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
    """パスワード認証機能"""
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    st.markdown("## 🔒 時間割作成システム ログイン")
    password = st.text_input("パスワードを入力してください", type="password")
    
    if st.button("ログイン"):
        # secretsが設定されていない場合（ローカル等）のための回避策
        correct_password = st.secrets["PASSWORD"] if "PASSWORD" in st.secrets else "1234"
        
        if password == correct_password:
            st.session_state.password_correct = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False

# --- ⚙️ 定数・設定 ---
st.set_page_config(layout="wide", page_title="中学校時間割システム")

# パスワードチェック
if not check_password():
    st.stop()

# --- 🛠️ ヘルパー関数群 ---
NAME_CORRECTIONS = {
    "ニシダ": "ニシタ",
    "オオシマ": "オシマ",
}

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
    if subject_name in ['総合', '道徳', '学活']: return subject_name
    short_class = class_name.replace('-', '')
    if '音美' in subject_name: return f"★{short_class}"
    return short_class

# --- 📊 Excel生成ロジック (ご希望のフォーマット) ---
def generate_excel(df_res, classes, teachers, df_const):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # スタイル定義
    thick = Side(style='thick'); medium = Side(style='medium'); thin = Side(style='thin'); hair = Side(style='hair')
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    # ---------------------------------------------------------
    # シート1: クラス別 (横にクラス、縦に時間)
    # ---------------------------------------------------------
    ws_c = wb.active
    ws_c.title = "クラス別"
    
    # ヘッダー作成
    ws_c.cell(row=1, column=1, value="曜").fill = header_fill
    ws_c.cell(row=1, column=2, value="限").fill = header_fill
    
    for i, c in enumerate(classes):
        col = 3 + i
        cell = ws_c.cell(row=1, column=col, value=c)
        cell.fill = header_fill
        cell.alignment = align_center
        ws_c.column_dimensions[get_column_letter(col)].width = 12

    days = ['月', '火', '水', '木', '金']
    curr = 2
    for d in days:
        periods = [1,2,3,4,5,6] if d != '金' else [1,2,3,4,5]
        max_p = periods[-1]
        for p in periods:
            # 罫線設定
            top = thick if p==1 else (medium if p==5 else thin)
            bottom = thick if p==max_p else (medium if p==4 else thin)
            
            # 曜日・限
            c_day = ws_c.cell(row=curr, column=1, value=d if p==1 else "")
            c_day.border = Border(top=top, bottom=bottom, left=thick, right=thin)
            c_day.alignment = align_center
            
            c_p = ws_c.cell(row=curr, column=2, value=p)
            c_p.border = Border(top=top, bottom=bottom, left=thin, right=thin)
            c_p.alignment = align_center
            
            # データ埋め込み
            for i, c in enumerate(classes):
                cell = ws_c.cell(row=curr, column=3+i)
                cell.border = Border(top=top, bottom=bottom, left=thin, right=thin)
                cell.alignment = align_center
                
                # 該当する授業を探す
                matches = df_res[(df_res['曜日']==d) & (df_res['限']==p) & (df_res['クラス']==c)]
                if not matches.empty:
                    r = matches.iloc[0]
                    # 表示形式: 教科(改行)教員名
                    txt = f"{r['教科']}\n{r['教員']}"
                    cell.value = txt
                    cell.font = Font(size=9)
            curr += 1

    # ---------------------------------------------------------
    # シート2: 教員別 (横に教員、縦に時間)
    # ---------------------------------------------------------
    ws_t = wb.create_sheet(title="教員別")
    
    ws_t.cell(row=1, column=1, value="曜").fill = header_fill
    ws_t.cell(row=1, column=2, value="限").fill = header_fill
    
    for i, t in enumerate(teachers):
        col = 3 + i
        cell = ws_t.cell(row=1, column=col, value=t)
        cell.fill = header_fill
        cell.alignment = align_center
        ws_t.column_dimensions[get_column_letter(col)].width = 6

    curr = 2
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
                
                # 授業検索
                matches = df_res[(df_res['曜日']==d) & (df_res['限']==p) & (df_res['教員'].str.contains(t, na=False))]
                val = ""
                if not matches.empty:
                    r = matches.iloc[0]
                    val = format_cell_text(r['クラス'], r['教科'])
                else:
                    # 固定リスト(部会など)検索
                    # df_constは標準化済みと仮定
                    for fix in df_const:
                        if fix['target'] == t and fix['day'] == {'月':0,'火':1,'水':2,'木':3,'金':4}[d] and fix['period'] == p:
                            val = f"【{fix['content']}】"
                            break
                            
                cell.value = val
                if val: cell.font = Font(size=11)
            curr += 1

    wb.save(output)
    return output.getvalue()


# --- 🧩 最適化ロジック (修正版) ---
def solve_schedule(teachers, req_list, fixed_list):
    model = cp_model.CpModel()
    DAYS = 5
    
    # 変数 X[req_id, day, period]
    X = {}
    
    # 1. 授業配置
    for r in req_list:
        rid = r['id']
        slots = []
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                X[(rid, d, p)] = model.NewBoolVar(f'r{rid}_d{d}_p{p}')
                slots.append(X[(rid, d, p)])
        model.Add(sum(slots) == r['num'])
        
        # 連続制約 (今日の修正点)
        if r['continuous'] and r['num'] == 2:
            pair_vars = []
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                pairs = [(1,2), (2,3), (3,4)]
                if p_max >= 6: pairs.append((5,6))
                for (p1, p2) in pairs:
                    b_pair = model.NewBoolVar(f'pair_{rid}_{d}_{p1}')
                    model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] == 2).OnlyEnforceIf(b_pair)
                    model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] != 2).OnlyEnforceIf(b_pair.Not())
                    pair_vars.append(b_pair)
            model.Add(sum(pair_vars) >= 1)

    # 2. クラス重複禁止
    classes = sorted(list(set(r['class'] for r in req_list)))
    for cls in classes:
        cls_reqs = [r for r in req_list if r['class'] == cls]
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                vars_here = [X[(r['id'], d, p)] for r in cls_reqs if (r['id'], d, p) in X]
                if vars_here:
                    model.Add(sum(vars_here) <= 1)

    # 3. 教員重複 & 固定リスト
    t_map = {t: [] for t in teachers}
    for r in req_list:
        if r['t1'] in teachers: t_map[r['t1']].append(r)
        if r['t2'] in teachers: t_map[r['t2']].append(r)
    
    for t in teachers:
        # 固定リスト適用
        for fix in fixed_list:
            if fix['target'] == t:
                d, p = fix['day'], fix['period']
                vars_here = [X[(r['id'], d, p)] for r in t_map[t] if (r['id'], d, p) in X]
                if vars_here:
                    model.Add(sum(vars_here) == 0)
        
        # 重複禁止
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                vars_here = [X[(r['id'], d, p)] for r in t_map[t] if (r['id'], d, p) in X]
                if vars_here:
                    model.Add(sum(vars_here) <= 1)

    # 4. 学年排他（体育など）
    grade_reqs = {}
    for r in req_list:
        g = r['class'].split('-')[0]
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

    # 目的関数 (授業をなるべく前に)
    obj_terms = []
    for (rid, d, p), var in X.items():
        obj_terms.append(var * p)
    model.Minimize(sum(obj_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        # 結果をDataFrame化して返す
        recs = []
        days_map = {0:'月', 1:'火', 2:'水', 3:'木', 4:'金'}
        for r in req_list:
            rid = r['id']
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                for p in range(1, p_max + 1):
                    if (rid, d, p) in X and solver.Value(X[(rid, d, p)]) == 1:
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

# --- 📱 メインアプリ画面 ---
st.title("🏫 中学校 時間割作成システム")

st.sidebar.header("1. データアップロード")
# keyを指定してリロード時の挙動を安定化
f_teacher = st.sidebar.file_uploader("教員データ", type='csv', key="up_t")
f_subject = st.sidebar.file_uploader("教科設定", type='csv', key="up_s")
f_req = st.sidebar.file_uploader("授業データ", type='csv', key="up_r")
f_fixed = st.sidebar.file_uploader("固定・禁止リスト", type='csv', key="up_f")

if st.sidebar.button("🚀 作成開始"):
    if not all([f_teacher, f_subject, f_req]):
        st.error("⚠️ 必須ファイル（教員、教科、授業）が不足しています。")
    else:
        with st.spinner("データの読み込みと診断中..."):
            try:
                # --- データ読み込み (Shift-JIS/UTF-8自動対応 & 列名検索) ---
                
                # 1. 教員
                df_teacher = pd.read_csv(f_teacher, encoding='utf-8-sig')
                c_name = find_col(df_teacher, ['教員名', '氏名', '名前'])
                if not c_name: raise ValueError("教員データに名前の列が見つかりません")
                df_teacher['教員名'] = df_teacher[c_name].apply(clean_name)
                teachers = df_teacher['教員名'].unique().tolist()
                
                # 2. 教科設定
                df_subj = pd.read_csv(f_subject, encoding='utf-8-sig')
                c_sname = find_col(df_subj, ['教科名', '教科'])
                c_cont = find_col(df_subj, ['連続'])
                continuous_flags = {}
                if c_sname:
                    for _, row in df_subj.iterrows():
                        s_name = str(row[c_sname]).strip()
                        is_cont = False
                        if c_cont:
                            val = str(row[c_cont])
                            if "〇" in val or "TRUE" in val.upper():
                                is_cont = True
                        continuous_flags[s_name] = is_cont
                
                # 3. 授業データ
                df_req = pd.read_csv(f_req, encoding='utf-8-sig')
                c_cls = find_col(df_req, ['クラス'])
                c_sub = find_col(df_req, ['教科'])
                c_t1 = find_col(df_req, ['担当教員', '教員1'])
                c_num = find_col(df_req, ['週コマ', '数'])
                c_t2 = find_col(df_req, ['担当教員2', '教員2', 'Ｔ２'])
                
                if not (c_cls and c_sub and c_t1 and c_num):
                    raise ValueError("授業データに必要な列（クラス、教科、担当教員、週コマ数）が不足しています")
                
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
                        # ★ 修正ポイント: 技術家庭科週1コマなら連続させない
                        is_cont = continuous_flags.get(subj, False)
                        if num < 2: is_cont = False
                        
                        req_list.append({
                            'id': req_id, 'class': cls, 'subject': subj,
                            't1': t1, 't2': t2, 'num': num, 'continuous': is_cont
                        })
                        req_id += 1

                # 4. 固定リスト
                fixed_list = []
                if f_fixed:
                    df_fix = pd.read_csv(f_fixed, encoding='utf-8-sig')
                    c_tar = find_col(df_fix, ['対象', '教員'])
                    c_day = find_col(df_fix, ['曜日'])
                    c_per = find_col(df_fix, ['限'])
                    c_con = find_col(df_fix, ['内容'])
                    
                    if c_tar and c_day and c_per:
                        for _, row in df_fix.iterrows():
                            target = clean_name(row[c_tar])
                            day_str = row[c_day]
                            try: p = int(row[c_per])
                            except: p = 0
                            content = row[c_con] if c_con else "用務"
                            
                            w_map = {'月':0, '火':1, '水':2, '木':3, '金':4}
                            if day_str in w_map and p > 0:
                                fixed_list.append({
                                    'target': target,
                                    'day': w_map[day_str],
                                    'period': p,
                                    'content': content
                                })

                # 計算実行
                st.info("計算を開始します...")
                df_result = solve_schedule(teachers, req_list, fixed_list)
                
                if df_result is not None:
                    st.success("🎉 時間割が完成しました！")
                    excel_data = generate_excel(df_result, sorted(list(set(r['class'] for r in req_list))), teachers, fixed_list)
                    
                    st.download_button(
                        label="📥 完成したExcelをダウンロード",
                        data=excel_data,
                        file_name="時間割_完成.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 解が見つかりませんでした。条件を緩和してください。")
                    
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
