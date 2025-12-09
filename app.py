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
        if password == st.secrets.get("PASSWORD", "1234"): # ローカルテスト用デフォルト
            st.session_state.password_correct = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False

# --- ⚙️ 定数・設定 ---
st.set_page_config(layout="wide", page_title="中学校時間割システム")
if "PASSWORD" in st.secrets:
    if not check_password(): st.stop()

MAJOR_SUBJECTS = ['国語', '社会', '数学', '理科', '英語']
SKILL_SUBJECTS = ['音楽', '美術', '体育', '技術', '家庭科', '技術家庭']
PRIORITIZE_AM_SUBJECTS = ['数学', '英語', '国語']
MAX_SKILL_SUBJECTS_PER_DAY = 2

# --- 🛠️ 関数群 ---

def format_cell_text(class_name, subject_name):
    if subject_name in ['総合', '道徳', '学活']: return subject_name
    short_class = class_name.replace('-', '')
    if subject_name == '音美': return f"★{short_class}"
    return short_class

def generate_excel(df_res, classes, teachers, df_const):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    thick = Side(style='thick'); medium = Side(style='medium'); thin = Side(style='thin'); hair = Side(style='hair')
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=False)
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    # 教員別
    ws_t = wb.active; ws_t.title = "教員別"
    ws_t.cell(row=6, column=1, value="曜").fill = header_fill
    ws_t.cell(row=6, column=2, value="限").fill = header_fill
    for i, t in enumerate(teachers):
        col = 3 + i
        ws_t.cell(row=6, column=col, value=t).fill = header_fill
        ws_t.column_dimensions[get_column_letter(col)].width = 5.5

    days = ['月', '火', '水', '木', '金']
    curr = 7
    for d in days:
        periods = [1,2,3,4,5,6] if d != '金' else [1,2,3,4,5]
        max_p = periods[-1]
        for p in periods:
            top = thick if p==1 else (medium if p==5 else thin)
            bottom = thick if p==max_p else (medium if p==4 else thin)
            ws_t.cell(row=curr, column=1, value=d if p==1 else "").border = Border(top=top, bottom=bottom, left=thick, right=thin)
            ws_t.cell(row=curr, column=2, value=p).border = Border(top=top, bottom=bottom, left=thin, right=thin)
            for i, t in enumerate(teachers):
                cell = ws_t.cell(row=curr, column=3+i)
                cell.border = Border(top=top, bottom=bottom, left=hair, right=hair)
                cell.alignment = align_center
                matches = df_res[(df_res['曜日']==d) & (df_res['限']==p) & (df_res['教員'].str.contains(t, na=False))]
                val = ""
                if not matches.empty:
                    r = matches.iloc[0]; val = format_cell_text(r['クラス'], r['教科'])
                else:
                    for _, cr in df_const.iterrows():
                        if cr['対象（教員名orクラス）'] == t and cr['曜日'] == d and cr['限'] == p:
                            val = cr['内容']; break
                cell.value = val
                if val: cell.font = Font(size=11)
            curr += 1
            
    # クラス別
    ws_c = wb.create_sheet(title="クラス別")
    ws_c.cell(row=1, column=1, value="曜").fill = header_fill
    ws_c.cell(row=1, column=2, value="限").fill = header_fill
    for i, c in enumerate(classes):
        col = 3 + i
        ws_c.cell(row=1, column=col, value=c).fill = header_fill
        ws_c.column_dimensions[get_column_letter(col)].width = 10
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
                cell.border = Border(top=top, bottom=bottom, left=thin, right=thin); cell.alignment = align_center
                matches = df_res[(df_res['曜日']==d) & (df_res['限']==p) & (df_res['クラス']==c)]
                if not matches.empty:
                    r = matches.iloc[0]; cell.value = f"{r['教科']}\n({r['教員']})"
                    cell.font = Font(size=9); cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            curr += 1

    wb.save(output)
    return output.getvalue()

def diagnose_schedule(df_schedule, teachers, classes):
    """現在のスケジュールを診断する"""
    warnings = []
    
    # 1. 教員の過密チェック
    for t in teachers:
        t_df = df_schedule[df_schedule['教員'].str.contains(t, na=False)]
        for d in ['月', '火', '水', '木', '金']:
            count = len(t_df[t_df['曜日'] == d])
            if count >= 5:
                warnings.append(f"⚠️ {t}: {d}曜に {count}コマ 入っています (過密)")
            
            # 午前満タンチェック
            am_count = len(t_df[(t_df['曜日'] == d) & (t_df['限'] <= 4)])
            if am_count >= 4:
                warnings.append(f"⚠️ {t}: {d}曜の午前が満タン(4コマ)です")

    # 2. クラスのバランスチェック
    for c in classes:
        c_df = df_schedule[df_schedule['クラス'] == c]
        for d in ['月', '火', '水', '木', '金']:
            day_df = c_df[c_df['曜日'] == d]
            subjects = day_df['教科'].tolist()
            majors = [s for s in subjects if s in MAJOR_SUBJECTS]
            if len(majors) >= 5:
                warnings.append(f"⚠️ {c}: {d}曜に主要5教科が全部入っています")
    
    return warnings

def solve_schedule(df_req, df_teacher, df_const, weights, recalc_classes, manual_instructions):
    """最適化計算"""
    
    # データ前処理
    for df in [df_req, df_teacher, df_const]:
        for col in df.columns:
            if df[col].dtype == object: df[col] = df[col].str.replace('ニシダ', 'ニシタ')

    classes = sorted(df_req['クラス'].unique())
    teachers = df_teacher['教員名'].unique().tolist()
    days = ['月', '火', '水', '木', '金']
    periods = {'月': [1,2,3,4,5,6], '火': [1,2,3,4,5,6], '水': [1,2,3,4,5,6], '木': [1,2,3,4,5,6], '金': [1,2,3,4,5]}

    model = cp_model.CpModel()
    x = {} 
    class_subjects = collections.defaultdict(list)
    
    for _, row in df_req.iterrows():
        c = row['クラス']; subj = row['教科']; t1 = row['担当教員']; t2 = row['担当教員２'] if pd.notna(row['担当教員２']) else None
        count = int(row['週コマ数'])
        if count == 0: continue
        is_2block = (subj in ['技術', '家庭科', '技術家庭'] and count >= 2)
        subj_id = (subj, t1, t2)
        for d in days:
            for p in periods[d]:
                x[(c, d, p, subj_id)] = model.NewBoolVar(f'x_{c}_{d}_{p}_{subj}')
        class_subjects[c].append({'subj': subj, 't1': t1, 't2': t2, 'count': count, 'id': subj_id, 'is_2block': is_2block})

    # 制約: 基本
    for c in classes:
        for d in days:
            for p in periods[d]: model.Add(sum(x[(c, d, p, item['id'])] for item in class_subjects[c]) <= 1)
    for c in classes:
        for item in class_subjects[c]: model.Add(sum(x[(c, d, p, item['id'])] for d in days for p in periods[d]) == item['count'])
    
    teacher_vars = collections.defaultdict(list)
    for c in classes:
        for item in class_subjects[c]:
            t1, t2 = item['t1'], item['t2']
            for d in days:
                for p in periods[d]:
                    var = x[(c, d, p, item['id'])]
                    if pd.notna(t1): teacher_vars[(t1, d, p)].append(var)
                    if pd.notna(t2): teacher_vars[(t2, d, p)].append(var)
    for key, vars_list in teacher_vars.items(): model.Add(sum(vars_list) <= 1)

    # 固定禁止
    for _, row in df_const.iterrows():
        target = row['対象（教員名orクラス）']; d = row['曜日']; content = row['内容']
        try: p = int(row['限'])
        except: continue
        if target in teachers:
            if (target, d, p) in teacher_vars: model.Add(sum(teacher_vars[(target, d, p)]) == 0)
        elif target in classes:
             for item in class_subjects[target]:
                if content in ['総合', '学活']:
                    if item['subj'] == content:
                         if (target, d, p, item['id']) in x: model.Add(x[(target, d, p, item['id'])] == 1)
                    else:
                         if (target, d, p, item['id']) in x: model.Add(x[(target, d, p, item['id'])] == 0)
        elif '全員' in target or '全学年' in target:
             target_grades = [1, 2, 3] 
             if '1年' in target: target_grades = [1]
             if '2,3年' in target: target_grades = [2, 3]
             for c in classes:
                 if int(c.split('-')[0]) in target_grades:
                     for item in class_subjects[c]:
                         if content in ['総合', '学活']:
                             if item['subj'] == content:
                                 if (c, d, p, item['id']) in x: model.Add(x[(c, d, p, item['id'])] == 1)
                             else:
                                 if (c, d, p, item['id']) in x: model.Add(x[(c, d, p, item['id'])] == 0)
    
    # 特殊授業
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
                    model.Add(sum(day_slots) == sum(start_vars) * 2)

    # 📌 【重要】個別指示の実装
    # manual_instructions = [{'target': '田中', 'type': '最大コマ数', 'day': '水', 'value': 4}, ...]
    if manual_instructions:
        for inst in manual_instructions:
            target = inst.get('対象')
            i_type = inst.get('指示タイプ')
            day = inst.get('曜日') # '月', '全日' etc
            val = inst.get('値')

            if not target: continue

            # 教員への指示
            if target in teachers:
                target_days = [day] if day in days else days
                
                # 1. 1日の最大コマ数制限 (例: 水曜は4コマまで)
                if i_type == '1日の最大コマ数':
                    try: limit = int(val)
                    except: continue
                    for d_target in target_days:
                        d_vars = []
                        for p in periods[d_target]:
                            if (target, d_target, p) in teacher_vars:
                                d_vars.extend(teacher_vars[(target, d_target, p)])
                        model.Add(sum(d_vars) <= limit)
                
                # 2. 午前/午後指定 (例: 午前を空ける -> 午前コマ数0)
                elif i_type == '午前の授業数':
                    try: limit = int(val)
                    except: continue
                    for d_target in target_days:
                        am_vars = []
                        for p in [1,2,3,4]:
                            if (target, d_target, p) in teacher_vars:
                                am_vars.extend(teacher_vars[(target, d_target, p)])
                        model.Add(sum(am_vars) == limit) # 厳密に指定

            # クラスへの指示 (例: 1-1 国語 午前)
            elif target in classes:
                # 教科指定がある場合を想定 (UI側で教科を入力させる必要あり)
                # 今回は簡易的に「教科」カラムがある前提
                subj_name = inst.get('教科')
                if not subj_name: continue
                
                if i_type == '優先配置':
                    if val == '午前':
                        for item in class_subjects[target]:
                            if item['subj'] == subj_name:
                                for d_loop in days:
                                    # 午後(5,6)を禁止にする
                                    pm_slots = []
                                    for p in [5, 6]:
                                        if p in periods[d_loop] and (target, d_loop, p, item['id']) in x:
                                            pm_slots.append(x[(target, d_loop, p, item['id'])])
                                    if pm_slots: model.Add(sum(pm_slots) == 0)

    # ロック処理 (再計算対象以外)
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

    # ペナルティ (スライダー)
    penalties = []
    
    # 先生負荷平準化
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

    # 午前満タン回避
    if weights['AM_FULL_AVOID'] > 0:
        for t in teachers:
            for d in days:
                am_vars = []
                for p in [1, 2, 3, 4]:
                    if (t, d, p) in teacher_vars: am_vars.extend(teacher_vars[(t, d, p)])
                # 固定部会も考慮
                mtg = sum(1 for _, r in df_const.iterrows() if r['対象（教員名orクラス）'] == t and r['曜日'] == d and r['限'] in [1,2,3,4])
                total = model.NewIntVar(0, 4, f'am_{t}_{d}')
                model.Add(total == sum(am_vars) + mtg)
                full = model.NewBoolVar(f'full_{t}_{d}')
                model.Add(total == 4).OnlyEnforceIf(full)
                model.Add(total < 4).OnlyEnforceIf(full.Not())
                penalties.append(full * weights['AM_FULL_AVOID'])

    # 生徒5教科分散
    if weights['STUDENT_5MAJORS'] > 0:
        for c in classes:
            for d in days:
                mj_vars = []
                for p in periods[d]:
                    for item in class_subjects[c]:
                        if item['subj'] in MAJOR_SUBJECTS:
                            if (c, d, p, item['id']) in x: mj_vars.append(x[(c, d, p, item['id'])])
                is_full = model.NewBoolVar(f'st5_{c}_{d}')
                model.Add(sum(mj_vars) >= 5).OnlyEnforceIf(is_full)
                model.Add(sum(mj_vars) < 5).OnlyEnforceIf(is_full.Not())
                penalties.append(is_full * weights['STUDENT_5MAJORS'])

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

st.sidebar.markdown("### 2. 全体バランス調整 (重み)")
w_load = st.sidebar.slider("先生の負担平準化", 0, 100, 20)
w_am = st.sidebar.slider("午前満タン回避", 0, 100, 30)
w_st5 = st.sidebar.slider("生徒5教科分散", 0, 200, 100)
w_skill = st.sidebar.slider("技能教科詰め込み回避", 0, 100, 50)
w_sandwich = st.sidebar.slider("サンドイッチ回避", 0, 100, 40)
w_am_place = st.sidebar.slider("主要科目(数英)の午前配置", 0, 100, 50)

st.sidebar.markdown("### 3. 再計算ターゲット")
recalc_str = st.sidebar.text_input("作り直すクラス (空欄なら全クラス)", "")
recalc_list = [x.strip() for x in recalc_str.split(',')] if recalc_str else []

# --- メインエリア ---
st.title("🏫 中学校時間割 AI作成システム")

if f_req and f_teacher and f_const:
    df_req = pd.read_csv(f_req)
    df_teacher = pd.read_csv(f_teacher)
    df_const = pd.read_csv(f_const)
    teachers = sorted(df_teacher['教員名'].unique().tolist())
    classes = sorted(df_req['クラス'].unique().tolist())
    
    # --- A. 個別指示エリア ---
    st.markdown("### 🗣️ 個別指示機能 (わがままリスト)")
    st.info("特定の先生やクラスに対して、個別のルールを追加できます。AIはこのルールを最優先で守ります。")
    
    # 個別指示の入力テーブル
    if 'instructions' not in st.session_state:
        st.session_state['instructions'] = pd.DataFrame(columns=['対象', '曜日', '教科', '指示タイプ', '値'])
    
    # 編集用データフレーム
    input_df = st.data_editor(
        st.session_state['instructions'],
        num_rows="dynamic",
        column_config={
            "対象": st.column_config.SelectboxColumn(options=teachers + classes, required=True),
            "曜日": st.column_config.SelectboxColumn(options=['全日', '月', '火', '水', '木', '金'], default='全日'),
            "教科": st.column_config.TextColumn(help="クラスへの指示の場合に入力"),
            "指示タイプ": st.column_config.SelectboxColumn(
                options=['1日の最大コマ数', '午前の授業数', '優先配置'], 
                required=True
            ),
            "値": st.column_config.TextColumn(help="数字 または '午前' など"),
        },
        key="editor",
        use_container_width=True
    )

    # --- B. 診断とプレビュー ---
    if 'schedule_df' in st.session_state:
        res_df = st.session_state['schedule_df']
        
        st.divider()
        st.subheader("🩺 AI診断レポート")
        warnings = diagnose_schedule(res_df, teachers, classes)
        if warnings:
            with st.expander(f"⚠️ {len(warnings)} 件の改善ポイントが見つかりました", expanded=True):
                for w in warnings:
                    st.write(f"- {w}")
        else:
            st.success("🎉 目立った問題点は見つかりませんでした！")

        st.subheader("📅 時間割プレビュー")
        
        # プレビュー表示
        days = ['月', '火', '水', '木', '金']
        periods = [1, 2, 3, 4, 5, 6]
        view_data = []
        for d in days:
            for p in periods:
                if d == '金' and p == 6: continue
                row = {'曜日': d, '限': p}
                for t in teachers: row[t] = ""
                view_data.append(row)
        df_view = pd.DataFrame(view_data)
        
        for _, r in res_df.iterrows():
            t_s = r['教員'].split(', ')
            val = format_cell_text(r['クラス'], r['教科'])
            for t in t_s:
                if t in df_view.columns:
                    mask = (df_view['曜日']==r['曜日']) & (df_view['限']==r['限'])
                    df_view.loc[mask, t] = val
        
        # 部会
        for _, cr in df_const.iterrows():
            t = cr['対象（教員名orクラス）']
            if t in teachers:
                mask = (df_view['曜日']==cr['曜日']) & (df_view['限']==cr['限'])
                current = df_view.loc[mask, t].values[0]
                if not current: df_view.loc[mask, t] = f"【{cr['内容']}】"

        st.dataframe(df_view, height=500, use_container_width=True)
        
        excel_data = generate_excel(res_df, classes, teachers, df_const)
        st.download_button("📥 Excelをダウンロード", excel_data, file_name="時間割_完成.xlsx")

    # --- 実行ボタン ---
    st.divider()
    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("🚀 作成開始 (再計算)", type="primary", use_container_width=True):
            # manual_instructions の作成
            manual_list = input_df.to_dict('records')
            # 空行削除
            manual_list = [m for m in manual_list if m['対象'] is not None]

            with st.spinner("AIがパズルを解いています... (約1分)"):
                weights = {
                    'TEACHER_LOAD': w_load, 'AM_FULL_AVOID': w_am,
                    'STUDENT_5MAJORS': w_st5, 'SKILL_OVERLOAD': w_skill,
                    'SANDWICH': w_sandwich, 'AM_PLACEMENT': w_am_place
                }
                
                res = solve_schedule(
                    df_req, df_teacher, df_const, 
                    weights, recalc_list, manual_list
                )
                
                if res is not None:
                    st.session_state['schedule_df'] = res
                    st.session_state['prev_schedule'] = res
                    st.success("作成完了！診断レポートを確認してください。")
                    st.rerun()
                else:
                    st.error("解が見つかりませんでした。個別指示が厳しすぎる可能性があります。")

else:
    st.info("👈 左側のサイドバーからCSVファイルを3つアップロードしてください。")
