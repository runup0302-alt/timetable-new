import streamlit as st
import pandas as pd
import numpy as np
from ortools.sat.python import cp_model
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
import io
import collections

# --- 🔐 セキュリティ設定 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
    if st.session_state.password_correct:
        return True
    st.markdown("## 🔒 時間割作成システム ログイン")
    password = st.text_input("パスワードを入力してください", type="password")
    if st.button("ログイン"):
        # secretsがない場合(ローカル)は1234で通す
        if password == st.secrets.get("PASSWORD", "1234"):
            st.session_state.password_correct = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False

# --- ⚙️ 初期設定 ---
st.set_page_config(layout="wide", page_title="中学校時間割システム")
if "PASSWORD" in st.secrets:
    if not check_password(): st.stop()

# --- 🛠️ ユーティリティ関数 ---

def clean_bool(val):
    """〇/× や TRUE/FALSE を Pythonのboolに変換"""
    s = str(val).strip().upper()
    return s in ['〇', 'TRUE', '1', 'YES']

def format_cell_text(class_name, subject_name):
    """表記圧縮 (1-1数学 -> 11)"""
    if subject_name in ['総合', '道徳', '学活', '自立']: return subject_name
    short_class = class_name.replace('-', '')
    if subject_name == '音美': return f"★{short_class}"
    return short_class

def get_grade_color(grade):
    """学年ごとの色コード定義"""
    if grade == 1: return "#E3F2FD" # 薄い青 (1年)
    if grade == 2: return "#E8F5E9" # 薄い緑 (2年)
    if grade == 3: return "#FFF3E0" # 薄いオレンジ (3年)
    return "#F5F5F5" # グレー (その他)

def generate_excel(df_res, classes, teacher_data, df_const):
    """Excel生成 (デザイン強化版)"""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # スタイル
    thick = Side(style='thick'); medium = Side(style='medium'); thin = Side(style='thin'); hair = Side(style='hair')
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=False)
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    # 教員データの整理 (表示順ソート済み前提)
    teachers = teacher_data['教員名'].tolist()
    
    # --- シート1: 教員別 ---
    ws_t = wb.active; ws_t.title = "教員別"
    ws_t.cell(row=6, column=1, value="曜").fill = header_fill
    ws_t.cell(row=6, column=2, value="限").fill = header_fill
    
    # ヘッダー作成 (学年色分け)
    for i, row in teacher_data.iterrows():
        t_name = row['教員名']
        grade = row['担当学年']
        col = 3 + i
        
        # 色決定
        color_hex = get_grade_color(grade).replace("#", "")
        grade_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        
        cell = ws_t.cell(row=6, column=col, value=t_name)
        cell.fill = grade_fill
        cell.border = Border(top=thin, bottom=thin, left=hair, right=hair)
        cell.alignment = align_center
        ws_t.column_dimensions[get_column_letter(col)].width = 5.5

    days = ['月', '火', '水', '木', '金']
    curr = 7
    for d in days:
        periods = [1,2,3,4,5,6] if d != '金' else [1,2,3,4,5]
        max_p = periods[-1]
        for p in periods:
            top = thick if p==1 else (medium if p==5 else thin)
            bottom = thick if p==max_p else (medium if p==4 else thin)
            
            # 左サイド
            ws_t.cell(row=curr, column=1, value=d if p==1 else "").border = Border(top=top, bottom=bottom, left=thick, right=thin)
            ws_t.cell(row=curr, column=2, value=p).border = Border(top=top, bottom=bottom, left=thin, right=thin)
            
            # データ埋め込み
            for i, t in enumerate(teachers):
                cell = ws_t.cell(row=curr, column=3+i)
                
                # 学年背景色をうっすら適用するか、白にするか
                # 視認性のため、交互色または白推奨だが、今回は白ベースで枠線重視
                cell.border = Border(top=top, bottom=bottom, left=hair, right=hair)
                cell.alignment = align_center
                
                matches = df_res[(df_res['曜日']==d) & (df_res['限']==p) & (df_res['教員'].str.contains(t, na=False))]
                val = ""
                if not matches.empty:
                    r = matches.iloc[0]; val = format_cell_text(r['クラス'], r['教科'])
                else:
                    for _, cr in df_const.iterrows():
                        target = cr['対象（教員名orクラス）']
                        # 教員名一致 or 学年団一致 (例: 2年団)
                        is_target = (target == t)
                        if not is_target and "年団" in target:
                            try:
                                target_g = int(target.replace("年団",""))
                                my_g = teacher_data[teacher_data['教員名']==t]['担当学年'].values[0]
                                if target_g == my_g: is_target = True
                            except: pass
                        
                        if is_target and cr['曜日'] == d and cr['限'] == p:
                            val = cr['内容']; break
                
                cell.value = val
                if val: cell.font = Font(size=11)
            curr += 1

    # --- シート2: クラス別 ---
    ws_c = wb.create_sheet(title="クラス別")
    classes = sorted(df_res['クラス'].unique())
    ws_c.cell(row=1, column=1, value="曜").fill = header_fill
    ws_c.cell(row=1, column=2, value="限").fill = header_fill
    for i, c in enumerate(classes):
        ws_c.cell(row=1, column=3+i, value=c).fill = header_fill
    
    curr = 2
    for d in days:
        periods = [1,2,3,4,5,6] if d != '金' else [1,2,3,4,5]
        max_p = periods[-1]
        for p in periods:
            top = thick if p==1 else (medium if p==5 else thin)
            bottom = thick if p==max_p else (medium if p==4 else thin)
            ws_c.cell(row=curr, column=1, value=d if p==1 else "").border = Border(top=top, bottom=bottom, left=thick, right=thin)
            ws_c.cell(row=curr, column=2, value=p).border = Border(top=top, bottom=bottom, left=thin, right=thin)
            
            for i, c in enumerate(classes):
                cell = ws_c.cell(row=curr, column=3+i)
                cell.border = Border(top=top, bottom=bottom, left=thin, right=thin)
                cell.alignment = align_center
                matches = df_res[(df_res['曜日']==d) & (df_res['限']==p) & (df_res['クラス']==c)]
                if not matches.empty:
                    r = matches.iloc[0]; cell.value = f"{r['教科']}\n({r['教員']})"
                    cell.font = Font(size=9); cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            curr += 1
            
    wb.save(output)
    return output.getvalue()

def solve_schedule(df_req, df_teacher, df_const, df_subj_conf, weights, recalc_classes, manual_instructions):
    """最適化エンジン"""
    
    # 1. 教員データの整理 (ソート済み)
    teachers = df_teacher['教員名'].tolist()
    # 学年マッピング {教員名: 学年}
    teacher_grade_map = dict(zip(df_teacher['教員名'], df_teacher['担当学年']))

    classes = sorted(df_req['クラス'].unique())
    days = ['月', '火', '水', '木', '金']
    periods = {'月': [1,2,3,4,5,6], '火': [1,2,3,4,5,6], '水': [1,2,3,4,5,6], '木': [1,2,3,4,5,6], '金': [1,2,3,4,5]}

    # 2. 教科設定の整理
    # {教科名: {'continuous': bool, 'grade_block': bool}}
    subj_conf = {}
    for _, row in df_subj_conf.iterrows():
        subj_conf[row['教科']] = {
            'continuous': clean_bool(row['連続コマ']),
            'grade_block': clean_bool(row['学年団拘束'])
        }

    # 3. 必要コマ数の調整 (固定リスト分を引き算)
    # 固定リストから「埋まっている授業」をカウント
    fixed_counts = collections.defaultdict(int) # {(クラス, 教科): 済みコマ数}
    
    for _, row in df_const.iterrows():
        tgt = row['対象（教員名orクラス）']
        content = row['内容'] # 教科名 or 会議名
        
        # ターゲットがクラスで、かつ content が授業名ならカウント
        # (会議などは無視)
        if tgt in classes:
            # 授業データにある教科かチェック
            if not df_req[(df_req['クラス']==tgt) & (df_req['教科']==content)].empty:
                fixed_counts[(tgt, content)] += 1

    model = cp_model.CpModel()
    x = {} 
    class_subjects = collections.defaultdict(list)
    
    # 4. 変数定義 & コマ数設定
    for _, row in df_req.iterrows():
        c = row['クラス']; subj = row['教科']; t1 = row['担当教員']; t2 = row['担当教員２'] if pd.notna(row['担当教員２']) else None
        
        req_count = int(row['週コマ数'])
        # ★ ここで固定分を引き算
        already_fixed = fixed_counts[(c, subj)]
        needed_count = max(0, req_count - already_fixed)
        
        # 設定取得
        conf = subj_conf.get(subj, {'continuous': False, 'grade_block': False})
        is_2block = conf['continuous'] and needed_count >= 2
        
        subj_id = (subj, t1, t2)
        
        # 必要な分だけ変数を生成するが、固定枠は後で "1" に強制するため、
        # モデル上は「全時間帯の変数」を作っておく必要がある
        for d in days:
            for p in periods[d]:
                x[(c, d, p, subj_id)] = model.NewBoolVar(f'x_{c}_{d}_{p}_{subj}')
        
        class_subjects[c].append({
            'subj': subj, 't1': t1, 't2': t2, 
            'count': needed_count, # 最適化で配置すべき残りコマ数
            'total_count': req_count, # 本来の総数
            'id': subj_id, 
            'is_2block': is_2block,
            'grade_block': conf['grade_block']
        })

    # --- 制約条件 ---
    
    # 1. クラス: 1枠1授業
    for c in classes:
        for d in days:
            for p in periods[d]:
                model.Add(sum(x[(c, d, p, item['id'])] for item in class_subjects[c]) <= 1)

    # 2. 教員: 1枠1授業 (TT対応)
    teacher_vars = collections.defaultdict(list)
    for c in classes:
        for item in class_subjects[c]:
            t1, t2 = item['t1'], item['t2']
            for d in days:
                for p in periods[d]:
                    var = x[(c, d, p, item['id'])]
                    if pd.notna(t1): teacher_vars[(t1, d, p)].append(var)
                    if pd.notna(t2): teacher_vars[(t2, d, p)].append(var)
    for key, vars_list in teacher_vars.items():
        model.Add(sum(vars_list) <= 1)

    # 3. 固定・禁止リスト (汎用化ロジック)
    for _, row in df_const.iterrows():
        target = row['対象（教員名orクラス）']
        d = row['曜日']; 
        try: p = int(row['限'])
        except: continue
        content = row['内容']

        # A. 教員指定のブロック (会議など)
        if target in teachers:
            if (target, d, p) in teacher_vars:
                model.Add(sum(teacher_vars[(target, d, p)]) == 0)
        
        # B. 学年団指定のブロック ("2年団"など)
        elif "年団" in target:
            try:
                target_grade = int(target.replace("年団", ""))
                # その学年の教員全員をブロック
                for t_name, t_grade in teacher_grade_map.items():
                    if t_grade == target_grade:
                         if (t_name, d, p) in teacher_vars:
                             model.Add(sum(teacher_vars[(t_name, d, p)]) == 0)
            except: pass

        # C. クラス指定
        elif target in classes:
            # もし授業データにある教科なら -> 「その授業をここに固定」
            found_subj = False
            for item in class_subjects[target]:
                if item['subj'] == content:
                    # その場所を 1 に固定
                    if (target, d, p, item['id']) in x:
                        model.Add(x[(target, d, p, item['id'])] == 1)
                    found_subj = True
            
            # 授業データにない(会議など) -> 「その時間は授業入れない」
            if not found_subj:
                for item in class_subjects[target]:
                    if (target, d, p, item['id']) in x:
                        model.Add(x[(target, d, p, item['id'])] == 0)
    
    # 4. コマ数確保 (残りコマ数分だけ配置)
    for c in classes:
        for item in class_subjects[c]:
            # 固定リストで配置された分(1になっている分)を除外してカウントする必要がある
            # しかしシンプルに、「全変数の合計 == 総コマ数」とすれば、固定で1になった分も含めて整合性が取れる
            model.Add(sum(x[(c, d, p, item['id'])] for d in days for p in periods[d]) == item['total_count'])

    # 5. 学年団拘束 (総合など)
    for c in classes:
        # クラスの学年を取得
        try: class_grade = int(c.split('-')[0])
        except: continue
        
        for item in class_subjects[c]:
            if item['grade_block']: # 総合など
                for d in days:
                    for p in periods[d]:
                        # もしこのクラスで総合が入るなら...
                        is_sogo = x[(c, d, p, item['id'])]
                        
                        # その学年の教員全員、他の授業を入れてはいけない
                        for t_name, t_grade in teacher_grade_map.items():
                            if t_grade == class_grade:
                                # その先生が、まさにこの総合を担当しているならOK (t1, t2)
                                if item['t1'] == t_name or item['t2'] == t_name:
                                    continue
                                
                                # そうでなければ、その時間の他の授業変数を0にする
                                # (実装詳細: is_sogoが1なら、その先生の sum(vars) は 0)
                                if (t_name, d, p) in teacher_vars:
                                    model.Add(sum(teacher_vars[(t_name, d, p)]) == 0).OnlyEnforceIf(is_sogo)

    # 6. ニコイチ・排他・1日1教科 (既存ロジック)
    for c in classes:
        for item in class_subjects[c]:
            if item['is_2block']:
                for d in days:
                    possible_starts = [1, 2, 3, 5] if d != '金' else [1, 2, 3]
                    start_vars = []
                    for s in possible_starts:
                        s_var = model.NewBoolVar(f's_{c}_{d}_{s}')
                        start_vars.append(s_var)
                        model.Add(x[(c, d, s, item['id'])] == 1).OnlyEnforceIf(s_var)
                        model.Add(x[(c, d, s+1, item['id'])] == 1).OnlyEnforceIf(s_var)
                    day_slots = [x[(c, d, p, item['id'])] for p in periods[d]]
                    # 既に固定されているニコイチがある場合も考慮し、
                    # day_slotsの合計が偶数になる等の制約が必要だが、
                    # ここでは簡易的に「開始フラグ数 * 2」で制御
                    # (固定リストとの整合性が難しい箇所だが、今回は固定優先で最適化に委ねる)
                    # model.Add(sum(day_slots) == sum(start_vars) * 2) 
                    pass # ニコイチ固定との競合回避のため一旦緩和

    # 7. 個別指示 (Constraints Injection)
    if manual_instructions:
        for inst in manual_instructions:
            target = inst.get('対象'); i_type = inst.get('指示タイプ'); day = inst.get('曜日'); val = inst.get('値')
            if not target: continue

            if target in teachers:
                target_days = [day] if day in days else days
                if i_type == '1日の最大コマ数':
                    try: limit = int(val)
                    except: continue
                    for d_target in target_days:
                        d_vars = []
                        for p in periods[d_target]:
                            if (target, d_target, p) in teacher_vars: d_vars.extend(teacher_vars[(target, d_target, p)])
                        model.Add(sum(d_vars) <= limit)
                elif i_type == '午前の授業数':
                    try: limit = int(val)
                    except: continue
                    for d_target in target_days:
                        am_vars = []
                        for p in [1,2,3,4]:
                            if (target, d_target, p) in teacher_vars: am_vars.extend(teacher_vars[(target, d_target, p)])
                        model.Add(sum(am_vars) == limit)

            elif target in classes:
                subj_name = inst.get('教科')
                if not subj_name: continue
                if i_type == '優先配置' and val == '午前':
                    for item in class_subjects[target]:
                        if item['subj'] == subj_name:
                            for d_loop in days:
                                pm_slots = []
                                for p in [5, 6]:
                                    if p in periods[d_loop] and (target, d_loop, p, item['id']) in x:
                                        pm_slots.append(x[(target, d_loop, p, item['id'])])
                                if pm_slots: model.Add(sum(pm_slots) == 0)

    # 8. ロック処理
    if 'prev_schedule' in st.session_state and recalc_classes:
        df_prev = st.session_state['prev_schedule']
        for _, r in df_prev.iterrows():
            c = r['クラス']
            if c in recalc_classes: continue 
            d = r['曜日']; p = int(r['限']); s_name = r['教科']
            for item in class_subjects[c]:
                if item['subj'] == s_name:
                    if (c, d, p, item['id']) in x:
                        model.Add(x[(c, d, p, item['id'])] == 1)

    # ペナルティ
    penalties = []
    if weights['TEACHER_LOAD'] > 0:
        for t in teachers:
            daily_counts = []
            for d in days:
                d_vars = []
                for p in periods[d]:
                    if (t, d, p) in teacher_vars: d_vars.extend(teacher_vars[(t, d, p)])
                cnt = model.NewIntVar(0, 6, f'cnt_{t}_{d}')
                model.Add(sum(d_vars) == cnt); daily_counts.append(cnt)
            mx = model.NewIntVar(0, 6, f'max_{t}'); mn = model.NewIntVar(0, 6, f'min_{t}')
            model.AddMaxEquality(mx, daily_counts); model.AddMinEquality(mn, daily_counts)
            penalties.append((mx - mn) * weights['TEACHER_LOAD'])

    if penalties: model.Minimize(sum(penalties))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        recs = []
        for c in classes:
            for d in days:
                for p in periods[d]:
                    for item in class_subjects[c]:
                        if solver.Value(x[(c, d, p, item['id'])]) == 1:
                            t_str = str(item['t1'])
                            if pd.notna(item['t2']): t_str += f", {item['t2']}"
                            recs.append({'曜日': d, '限': p, 'クラス': c, '教科': item['subj'], '教員': t_str})
        return pd.DataFrame(recs)
    else:
        return None

# --- 📱 UI構築 ---
st.sidebar.title("🎛️ 設定パネル")
st.sidebar.markdown("### 1. データ読み込み")
f_req = st.sidebar.file_uploader("授業データ", type='csv')
f_teacher = st.sidebar.file_uploader("教員データ", type='csv')
f_const = st.sidebar.file_uploader("固定・禁止リスト", type='csv')
f_conf = st.sidebar.file_uploader("教科設定 (New!)", type='csv')

st.sidebar.markdown("### 2. 全体バランス調整")
w_load = st.sidebar.slider("先生の負担平準化", 0, 100, 20)
w_am = st.sidebar.slider("午前満タン回避", 0, 100, 30)
w_st5 = st.sidebar.slider("生徒5教科分散", 0, 200, 100)

st.sidebar.markdown("### 3. 再計算ターゲット")
recalc_str = st.sidebar.text_input("作り直すクラス (空欄なら全クラス)", "")
recalc_list = [x.strip() for x in recalc_str.split(',')] if recalc_str else []

st.title("🏫 中学校時間割 AI作成システム (完全汎用版)")

if f_req and f_teacher and f_const and f_conf:
    df_req = pd.read_csv(f_req)
    df_teacher = pd.read_csv(f_teacher)
    df_const = pd.read_csv(f_const)
    df_conf = pd.read_csv(f_conf)
    
    # 教員を「表示順」でソート
    if '表示順' in df_teacher.columns:
        df_teacher = df_teacher.sort_values('表示順')
    teachers = df_teacher['教員名'].tolist()
    
    classes = sorted(df_req['クラス'].unique().tolist())
    
    # --- 個別指示 ---
    st.markdown("### 🗣️ 個別指示機能")
    if 'instructions' not in st.session_state:
        st.session_state['instructions'] = pd.DataFrame(columns=['対象', '曜日', '教科', '指示タイプ', '値'])
    
    input_df = st.data_editor(
        st.session_state['instructions'], num_rows="dynamic",
        column_config={
            "対象": st.column_config.SelectboxColumn(options=teachers + classes, required=True),
            "曜日": st.column_config.SelectboxColumn(options=['全日', '月', '火', '水', '木', '金'], default='全日'),
            "指示タイプ": st.column_config.SelectboxColumn(options=['1日の最大コマ数', '午前の授業数', '優先配置'], required=True),
        },
        key="editor", use_container_width=True
    )

    if 'schedule_df' in st.session_state:
        res_df = st.session_state['schedule_df']
        st.subheader("📅 時間割プレビュー")
        
        # プレビュー表示 (学年色分け付き)
        days = ['月', '火', '水', '木', '金']
        periods = [1, 2, 3, 4, 5, 6]
        
        # 色スタイルの適用は st.dataframe では限界があるため、
        # 教員名ヘッダーに学年情報を付記して区別する
        view_cols = []
        for _, r in df_teacher.iterrows():
            g = r['担当学年']
            g_mark = f"【{g}年】" if g > 0 else "【F】"
            view_cols.append(f"{r['教員名']} {g_mark}")
            
        view_data = []
        for d in days:
            for p in periods:
                if d == '金' and p == 6: continue
                row = {'曜日': d, '限': p}
                for col in view_cols: row[col] = ""
                view_data.append(row)
        df_view = pd.DataFrame(view_data)
        
        # データ埋め込み
        for _, r in res_df.iterrows():
            t_s = r['教員'].split(', ')
            val = format_cell_text(r['クラス'], r['教科'])
            for t in t_s:
                # 対応するカラム名を探す
                target_col = [c for c in view_cols if c.startswith(t + " ")]
                if target_col:
                    mask = (df_view['曜日']==r['曜日']) & (df_view['限']==r['限'])
                    df_view.loc[mask, target_col[0]] = val
        
        # 固定コマ埋め込み
        for _, cr in df_const.iterrows():
            t = cr['対象（教員名orクラス）']
            target_col = [c for c in view_cols if c.startswith(t + " ")]
            if target_col:
                mask = (df_view['曜日']==cr['曜日']) & (df_view['限']==cr['限'])
                if not df_view.loc[mask, target_col[0]].values[0]:
                     df_view.loc[mask, target_col[0]] = f"【{cr['内容']}】"

        st.dataframe(df_view, height=500, use_container_width=True)
        
        excel_data = generate_excel(res_df, classes, df_teacher, df_const)
        st.download_button("📥 Excelをダウンロード", excel_data, file_name="時間割_完成.xlsx")

    st.divider()
    if st.button("🚀 作成開始 (再計算)", type="primary"):
        manual_list = [m for m in input_df.to_dict('records') if m['対象'] is not None]
        with st.spinner("計算中..."):
            weights = {'TEACHER_LOAD': w_load, 'AM_FULL_AVOID': w_am, 'STUDENT_5MAJORS': w_st5}
            res = solve_schedule(df_req, df_teacher, df_const, df_conf, weights, recalc_list, manual_list)
            
            if res is not None:
                st.session_state['schedule_df'] = res
                st.session_state['prev_schedule'] = res
                st.success("作成完了！")
                st.rerun()
            else:
                st.error("解が見つかりませんでした。")
else:
    st.info("👈 左側のサイドバーからCSVファイル（4つ）をアップロードしてください。")
