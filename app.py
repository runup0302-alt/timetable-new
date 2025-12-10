import streamlit as st
import pandas as pd
import numpy as np
from ortools.sat.python import cp_model
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
import io
import collections
import re

# ==========================================
# ⚙️ 設定・定数
# ==========================================
st.set_page_config(layout="wide", page_title="中学校時間割作成システム")

# 🔐 パスワード設定 (secrets.toml または デフォルト)
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
    if st.session_state.password_correct:
        return True
    
    st.markdown("## 🔒 ログイン")
    password = st.text_input("パスワード", type="password")
    if st.button("ログイン"):
        # secretsがない場合のバックアップ
        correct = st.secrets["PASSWORD"] if "PASSWORD" in st.secrets else "1234"
        if password == correct:
            st.session_state.password_correct = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False

if not check_password():
    st.stop()

# 教科定義
MAJOR_SUBJECTS = ['国語', '社会', '数学', '理科', '英語']
SKILL_SUBJECTS = ['音楽', '美術', '体育', '技術', '家庭科', '技術家庭']
FORCE_FIX_SUBJECTS = ['総合', '学活', '道徳', 'ＬＨＲ', 'LHR'] # 固定リストで強制配置する教科

# 表記ゆれ辞書
NAME_CORRECTIONS = {
    "ニシダ": "ニシタ",
    "オオシマ": "オシマ",
}

# ==========================================
# 🛠️ ヘルパー関数
# ==========================================
def clean_name(name):
    """名前の空白除去と表記ゆれ補正"""
    if pd.isna(name) or name == "":
        return ""
    name = str(name).replace(" ", "").replace("　", "")
    return NAME_CORRECTIONS.get(name, name)

def find_col(df, keywords):
    """列名をあいまい検索"""
    for col in df.columns:
        for kw in keywords:
            if kw in col:
                return col
    return None

def format_cell_text(class_name, subject_name):
    """Excelセル内の表記短縮"""
    if subject_name in FORCE_FIX_SUBJECTS: return subject_name
    short_class = class_name.replace('-', '')
    if '音美' in subject_name: return f"★{short_class}"
    return short_class

def parse_manual_overrides(text):
    """手動ピン留めテキスト解析"""
    overrides = []
    if not text: return overrides
    for line in text.split('\n'):
        parts = [p.strip() for p in line.split(',')]
        if len(parts) >= 4:
            # 教員orクラス, 曜日, 限, 教科
            overrides.append({'target': parts[0], 'day': parts[1], 'period': int(parts[2]), 'subj': parts[3]})
    return overrides

def get_target_classes(target_str, all_classes):
    """固定リストの対象（'1年', '全学年', '2,3年'など）をクラスリストに変換"""
    target_str = str(target_str)
    targets = []
    
    if target_str in all_classes:
        return [target_str]
    
    # 学年指定の解析
    if '全' in target_str:
        return all_classes
    
    # "1年", "2,3年" などの解析
    target_grades = []
    if '1' in target_str: target_grades.append('1')
    if '2' in target_str: target_grades.append('2')
    if '3' in target_str: target_grades.append('3')
    
    for c in all_classes:
        g = c.split('-')[0]
        if g in target_grades:
            targets.append(c)
            
    return targets

# ==========================================
# 📊 Excel生成ロジック (マトリックス形式)
# ==========================================
def generate_excel(df_res, classes, teachers, df_const):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # スタイル定義
    thick = Side(style='thick')
    medium = Side(style='medium')
    thin = Side(style='thin')
    hair = Side(style='hair')
    
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    side_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    days = ['月', '火', '水', '木', '金']

    # ---------------------------------------------------------
    # シート1: 教員別 (縦:時間, 横:教員)
    # ---------------------------------------------------------
    ws_t = wb.active
    ws_t.title = "教員別"
    
    ws_t.cell(row=6, column=1, value="曜").fill = header_fill
    ws_t.cell(row=6, column=2, value="限").fill = header_fill
    
    for i, t in enumerate(teachers):
        col = 3 + i
        ws_t.cell(row=6, column=col, value=t).fill = header_fill
        ws_t.column_dimensions[get_column_letter(col)].width = 6

    curr = 7
    for d in days:
        periods = [1,2,3,4,5,6] if d != '金' else [1,2,3,4,5]
        max_p = periods[-1]
        for p in periods:
            top = thick if p==1 else (medium if p==5 else thin)
            bottom = thick if p==max_p else (medium if p==4 else thin)
            
            c_day = ws_t.cell(row=curr, column=1, value=d if p==1 else "")
            c_day.fill = side_fill
            c_day.border = Border(top=top, bottom=bottom, left=thick, right=thin)
            c_day.alignment = align_center
            
            c_p = ws_t.cell(row=curr, column=2, value=p)
            c_p.fill = side_fill
            c_p.border = Border(top=top, bottom=bottom, left=thin, right=thin)
            c_p.alignment = align_center
            
            for i, t in enumerate(teachers):
                cell = ws_t.cell(row=curr, column=3+i)
                cell.border = Border(top=top, bottom=bottom, left=hair, right=hair)
                cell.alignment = align_center
                
                # ★修正: 教員1または教員2として含まれている授業を探す
                # df_resの '教員' 列には "田中, 鈴木" のように入っている想定
                matches = df_res[(df_res['曜日']==d) & (df_res['限']==p) & (df_res['教員'].str.contains(t, na=False))]
                
                val = ""
                if not matches.empty:
                    # 授業がある場合
                    r = matches.iloc[0]
                    val = format_cell_text(r['クラス'], r['教科'])
                else:
                    # 授業がない場合、固定リスト（会議等）を確認
                    # ここでは簡易的に df_const（辞書リスト）を走査
                    for fix in df_const:
                        # 対象が教員名と一致する場合
                        if fix['target'] == t and fix['day'] == {'月':0,'火':1,'水':2,'木':3,'金':4}[d] and fix['period'] == p:
                            # 授業として割り当てられていない会議等を表示
                            val = f"【{fix['content']}】"
                            break
                
                cell.value = val
                if val: cell.font = Font(size=11)
            curr += 1

    # ---------------------------------------------------------
    # シート2: クラス別 (縦:時間, 横:クラス)
    # ---------------------------------------------------------
    ws_c = wb.create_sheet(title="クラス別")
    
    ws_c.cell(row=1, column=1, value="曜").fill = header_fill
    ws_c.cell(row=1, column=2, value="限").fill = header_fill
    
    for i, c in enumerate(classes):
        col = 3 + i
        cell = ws_c.cell(row=1, column=col, value=c)
        cell.fill = header_fill
        cell.alignment = align_center
        ws_c.column_dimensions[get_column_letter(col)].width = 12

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
                    r = matches.iloc[0]
                    # 教科名と教員名を表示
                    cell.value = f"{r['教科']}\n{r['教員']}"
                    cell.font = Font(size=9)
            curr += 1
            
    wb.save(output)
    return output.getvalue()


# ==========================================
# 🧩 最適化ロジック (全ルール適用版)
# ==========================================
def solve_schedule(teachers, req_list, fixed_list, weights, recalc_classes, manual_overrides, prev_df):
    model = cp_model.CpModel()
    DAYS = 5
    days_map = {0:'月', 1:'火', 2:'水', 3:'木', 4:'金'}
    
    classes = sorted(list(set(r['class'] for r in req_list)))
    
    # 変数 X[req_id, day, period]
    X = {}
    class_subjects = collections.defaultdict(list)

    # 1. 変数定義 & 基本制約
    for r in req_list:
        rid = r['id']
        class_subjects[r['class']].append(r)
        
        slots = []
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                X[(rid, d, p)] = model.NewBoolVar(f'r{rid}_d{d}_p{p}')
                slots.append(X[(rid, d, p)])
        
        # コマ数制約
        model.Add(sum(slots) == r['num'])
        
        # 連続制約 (ニコイチ)
        # 設定がTrue かつ 週2コマ以上の場合のみ
        if r['continuous'] and r['num'] >= 2:
            pair_vars = []
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                # 昼休み跨ぎ(4-5)禁止
                pairs = [(1,2), (2,3), (3,4)]
                if p_max >= 6: pairs.append((5,6))
                
                for (p1, p2) in pairs:
                    b_pair = model.NewBoolVar(f'pair_{rid}_{d}_{p1}')
                    model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] == 2).OnlyEnforceIf(b_pair)
                    model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] != 2).OnlyEnforceIf(b_pair.Not())
                    pair_vars.append(b_pair)
            
            # 少なくとも (コマ数 // 2) 組のペアを作る
            model.Add(sum(pair_vars) >= r['num'] // 2)

    # 2. クラス内 重複禁止
    for cls in classes:
        cls_reqs = class_subjects[cls]
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                vars_here = [X[(r['id'], d, p)] for r in cls_reqs if (r['id'], d, p) in X]
                model.Add(sum(vars_here) <= 1)

    # 3. ★修正: 教員重複禁止 (T1もT2も考慮)
    # 教員ごとの担当授業リストを作成
    teacher_assignments = {t: [] for t in teachers}
    for r in req_list:
        # T1として担当
        if r['t1'] in teachers:
            teacher_assignments[r['t1']].append(r)
        # T2として担当 (ここが重要！)
        if r['t2'] in teachers:
            teacher_assignments[r['t2']].append(r)
            
    for t in teachers:
        t_reqs = teacher_assignments[t]
        
        # 固定リストの処理 (禁止 or 強制)
        # ここでは「授業禁止（会議等）」の処理を行う
        # 「強制配置（総合等）」は後述
        for fix in fixed_list:
            if fix['target'] == t:
                # 内容が「授業名」でない場合 -> 会議とみなしてブロック
                # (簡易判定: FORCE_FIX_SUBJECTS に含まれないなら会議)
                if fix['content'] not in FORCE_FIX_SUBJECTS:
                    d, p = fix['day'], fix['period']
                    vars_here = [X[(r['id'], d, p)] for r in t_reqs if (r['id'], d, p) in X]
                    if vars_here:
                        model.Add(sum(vars_here) == 0)

        # 重複禁止 (T1, T2すべての担当授業について、同時刻は1つまで)
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                vars_here = [X[(r['id'], d, p)] for r in t_reqs if (r['id'], d, p) in X]
                if vars_here:
                    model.Add(sum(vars_here) <= 1)

    # 4. 学年排他 (体育・理科)
    grade_reqs = {}
    for r in req_list:
        g = r['class'].split('-')[0] # "1-1" -> "1"
        if g not in grade_reqs: grade_reqs[g] = []
        grade_reqs[g].append(r)
    
    excl_subjs = ["体育", "理科", "音楽", "美術"]
    for g, reqs in grade_reqs.items():
        for subj_name in excl_subjs:
            target_reqs = [r for r in reqs if subj_name in r['subject'] or "音美" in r['subject']]
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                for p in range(1, p_max + 1):
                    vars_here = [X[(r['id'], d, p)] for r in target_reqs if (r['id'], d, p) in X]
                    if vars_here:
                        model.Add(sum(vars_here) <= 1)

    # 5. 音美ルール (音美がある日は、単独の音楽/美術は禁止)
    for cls in classes:
        cls_reqs = class_subjects[cls]
        has_onbi = any("音美" in r['subject'] for r in cls_reqs)
        if has_onbi:
            reqs_onbi = [r for r in cls_reqs if "音美" in r['subject']]
            reqs_single = [r for r in cls_reqs if r['subject'] in ["音楽", "美術"]]
            
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                is_onbi_day = model.NewBoolVar(f'onbi_day_{cls}_{d}')
                onbi_vars = []
                for p in range(1, p_max + 1):
                    for r in reqs_onbi:
                        if (r['id'], d, p) in X: onbi_vars.append(X[(r['id'], d, p)])
                
                # 音美があれば is_onbi_day = 1
                model.Add(sum(onbi_vars) >= 1).OnlyEnforceIf(is_onbi_day)
                model.Add(sum(onbi_vars) == 0).OnlyEnforceIf(is_onbi_day.Not())
                
                # 音美の日は単独科目禁止
                for p in range(1, p_max + 1):
                    for r in reqs_single:
                        if (r['id'], d, p) in X:
                            model.Add(X[(r['id'], d, p)] == 0).OnlyEnforceIf(is_onbi_day)

    # 6. ★修正: 固定リストによる「強制配置」 (総合、学活など)
    for fix in fixed_list:
        if fix['content'] in FORCE_FIX_SUBJECTS:
            # 対象クラスを取得 (1年, 2,3年, 全学年対応)
            targets = get_target_classes(fix['target'], classes)
            d, p = fix['day'], fix['period']
            
            for cls in targets:
                # そのクラスの該当教科の授業IDを探す
                found = False
                for r in class_subjects[cls]:
                    if r['subject'] == fix['content']:
                        if (r['id'], d, p) in X:
                            model.Add(X[(r['id'], d, p)] == 1)
                            found = True
                            # 1コマ分埋めたらbreak (週1コマの場合などのため)
                            # 週2コマ以上ある場合は、他の曜日も指定されているはず
                            break 

    # 7. 再計算ロック
    if prev_df is not None:
        try:
            for index, row in prev_df.iterrows():
                d_str = row.get('曜', row.get('曜日'))
                p = int(row['限'])
                d_idx = {'月':0, '火':1, '水':2, '木':3, '金':4}.get(d_str, -1)
                
                if d_idx == -1: continue

                for col_cls in prev_df.columns:
                    if col_cls not in classes: continue
                    if col_cls in recalc_classes: continue # 再計算クラスは無視
                    
                    cell_val = str(row[col_cls])
                    if cell_val == 'nan' or cell_val == '':
                        # 空きコマ固定
                        for r in class_subjects[col_cls]:
                            if (r['id'], d_idx, p) in X: model.Add(X[(r['id'], d_idx, p)] == 0)
                    else:
                        # 授業固定 (教科名マッチング)
                        subj_name = cell_val.split('\n')[0].strip()
                        for r in class_subjects[col_cls]:
                            if r['subject'] == subj_name:
                                if (r['id'], d_idx, p) in X:
                                    model.Add(X[(r['id'], d_idx, p)] == 1)
                                    break
        except:
            pass

    # 8. 手動ピン留め
    for o in manual_overrides:
        tgt, d, p, s_name = o['target'], {'月':0, '火':1, '水':2, '木':3, '金':4}.get(o['day'], -1), o['period'], o['subj']
        if d == -1: continue
        
        # クラス指定
        if tgt in classes:
            for r in class_subjects[tgt]:
                if r['subject'] == s_name:
                    if (r['id'], d, p) in X: model.Add(X[(r['id'], d, p)] == 1)
        # 教員指定
        elif tgt in teachers:
            for r in teacher_assignments[tgt]:
                if r['subject'] == s_name:
                    if (r['id'], d, p) in X: model.Add(X[(r['id'], d, p)] == 1)

    # 目的関数
    obj_terms = []
    # 前詰め
    for (rid, d, p), var in X.items():
        obj_terms.append(var * p * weights['AM_PLACEMENT'])
    
    # 先生の負担分散
    if weights['TEACHER_LOAD'] > 0:
        for t in teachers:
            daily_counts = []
            for d in range(DAYS):
                d_vars = []
                p_max = 5 if d == 4 else 6
                for p in range(1, p_max+1):
                    # T1, T2 両方カウント
                    for r in teacher_assignments[t]:
                        if (r['id'], d, p) in X: d_vars.append(X[(r['id'], d, p)])
                cnt = model.NewIntVar(0, 6, f'tc_{t}_{d}')
                model.Add(sum(d_vars) == cnt)
                daily_counts.append(cnt)
            mx = model.NewIntVar(0, 6, f'max_{t}')
            mn = model.NewIntVar(0, 6, f'min_{t}')
            model.AddMaxEquality(mx, daily_counts)
            model.AddMinEquality(mn, daily_counts)
            obj_terms.append((mx - mn) * weights['TEACHER_LOAD'])

    model.Minimize(sum(obj_terms))

    # ソルバー実行
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        recs = []
        for r in req_list:
            rid = r['id']
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                for p in range(1, p_max + 1):
                    if (rid, d, p) in X and solver.Value(X[(rid, d, p)]) == 1:
                        # T1とT2を結合して表示
                        t_str = r['t1']
                        if r['t2']: t_str += f", {r['t2']}"
                        recs.append({
                            '曜日': days_map[d],
                            '限': p,
                            'クラス': r['class'],
                            '教科': r['subject'],
                            '教員': t_str
                        })
        return pd.DataFrame(recs)
    else:
        return None

# ==========================================
# 📱 メイン画面
# ==========================================
st.title("🏫 中学校 時間割作成システム (決定版)")
st.info("左側のサイドバーからCSVファイルをアップロードしてください。")

st.sidebar.header("1. データアップロード")
f_teacher = st.sidebar.file_uploader("教員データ", type='csv', key="up_t")
f_subject = st.sidebar.file_uploader("教科設定", type='csv', key="up_s")
f_req = st.sidebar.file_uploader("授業データ", type='csv', key="up_r")
f_fixed = st.sidebar.file_uploader("固定・禁止リスト", type='csv', key="up_f")

st.sidebar.markdown("---")
f_prev = st.sidebar.file_uploader("🔄 再計算用Excel (前回データ)", type='xlsx', key="up_prev")
recalc_str = st.sidebar.text_input("作り直すクラス (例: 1-1, 1-2)", "")

st.sidebar.header("2. 設定")
w_load = st.sidebar.slider("教員負担の平準化", 0, 100, 20)
manual_str = st.sidebar.text_area("手動ピン留め (例: 1-1,月,1,国語)", height=100)

if st.sidebar.button("🚀 作成開始"):
    if not all([f_teacher, f_subject, f_req]):
        st.error("⚠️ 必須ファイルが不足しています。")
    else:
        with st.spinner("計算中..."):
            try:
                # -----------------------
                # データ読み込み & 前処理
                # -----------------------
                # 教員
                df_teacher = pd.read_csv(f_teacher, encoding='utf-8-sig')
                c_name = find_col(df_teacher, ['教員名', '氏名', '名前'])
                df_teacher['教員名'] = df_teacher[c_name].apply(clean_name)
                teachers = df_teacher['教員名'].unique().tolist()
                
                # 教科
                df_subj = pd.read_csv(f_subject, encoding='utf-8-sig')
                c_sname = find_col(df_subj, ['教科名', '教科'])
                c_cont = find_col(df_subj, ['連続'])
                continuous_flags = {}
                for _, row in df_subj.iterrows():
                    s_name = str(row[c_sname]).strip()
                    is_cont = False
                    if c_cont:
                        val = str(row[c_cont])
                        if "〇" in val or "TRUE" in val.upper(): is_cont = True
                    continuous_flags[s_name] = is_cont
                
                # 授業
                df_req = pd.read_csv(f_req, encoding='utf-8-sig')
                c_cls = find_col(df_req, ['クラス'])
                c_sub = find_col(df_req, ['教科'])
                c_t1 = find_col(df_req, ['担当教員', '教員1'])
                c_num = find_col(df_req, ['週コマ', '数'])
                c_t2 = find_col(df_req, ['担当教員2', '教員2'])
                
                req_list = []
                req_id = 0
                for _, row in df_req.iterrows():
                    cls = str(row[c_cls]).strip()
                    subj = str(row[c_sub]).strip()
                    t1 = clean_name(row[c_t1])
                    t2 = clean_name(row[c_t2]) if c_t2 else ""
                    try: num = int(row[c_num])
                    except: continue
                    if num > 0:
                        is_cont = continuous_flags.get(subj, False)
                        if num < 2: is_cont = False
                        req_list.append({
                            'id': req_id, 'class': cls, 'subject': subj,
                            't1': t1, 't2': t2, 'num': num, 'continuous': is_cont
                        })
                        req_id += 1
                
                # 固定リスト
                fixed_list = []
                if f_fixed:
                    df_fix = pd.read_csv(f_fixed, encoding='utf-8-sig')
                    c_tar = find_col(df_fix, ['対象', '教員'])
                    c_day = find_col(df_fix, ['曜日'])
                    c_per = find_col(df_fix, ['限'])
                    c_con = find_col(df_fix, ['内容'])
                    if c_tar:
                        for _, row in df_fix.iterrows():
                            target = clean_name(row[c_tar])
                            day_str = row[c_day]
                            try: p = int(row[c_per])
                            except: p = 0
                            content = row[c_con] if c_con else "用務"
                            w_map = {'月':0, '火':1, '水':2, '木':3, '金':4}
                            if day_str in w_map and p > 0:
                                fixed_list.append({'target': target, 'day': w_map[day_str], 'period': p, 'content': content})

                # 再計算・手動
                recalc_classes = [x.strip() for x in recalc_str.split(',')] if recalc_str else []
                prev_df = pd.read_excel(f_prev, sheet_name='クラス別') if f_prev else None
                manual_overrides = parse_manual_overrides(manual_str)
                
                weights = {'TEACHER_LOAD': w_load, 'AM_PLACEMENT': 20} # AM配置は固定

                # -----------------------
                # 実行
                # -----------------------
                df_res = solve_schedule(teachers, req_list, fixed_list, weights, recalc_classes, manual_overrides, prev_df)
                
                if df_res is not None:
                    st.success("🎉 時間割が完成しました！")
                    excel_data = generate_excel(df_res, sorted(list(set(r['class'] for r in req_list))), teachers, fixed_list)
                    st.download_button("📥 Excelをダウンロード", excel_data, "時間割.xlsx")
                else:
                    st.error("❌ 解が見つかりませんでした。条件を緩和してください。")

            except Exception as e:
                st.error(f"エラー: {e}")
