import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# ==========================================
# 設定・定数
# ==========================================
st.set_page_config(page_title="時間割作成システム", layout="wide")

# 表記ゆれ吸収辞書
NAME_CORRECTIONS = {
    "ニシダ": "ニシタ",
    "オオシマ": "オシマ",
    # 必要に応じて追加
}

def clean_name(name):
    """名前の空白除去と表記ゆれ補正"""
    if pd.isna(name) or name == "":
        return ""
    # 全角・半角スペース除去
    name = str(name).replace(" ", "").replace("　", "")
    return NAME_CORRECTIONS.get(name, name)

# ==========================================
# メインアプリ処理
# ==========================================
def main():
    st.title("🏫 中学校 時間割作成システム (Streamlit版)")
    st.markdown("""
    **手順:**
    1. 左側のサイドバーから必要なCSVファイルを5つアップロードしてください。
    2. 「作成開始」ボタンを押してください。
    3. 完成したExcelファイルをダウンロードできます。
    """)

    # --- サイドバー：ファイルアップロード ---
    st.sidebar.header("📂 データアップロード")
    
    file_teacher = st.sidebar.file_uploader("教員データ", type=["csv"])
    file_subject = st.sidebar.file_uploader("教科設定 - 年間", type=["csv"])
    file_req = st.sidebar.file_uploader("授業データ (前期or後期)", type=["csv"])
    file_fixed = st.sidebar.file_uploader("固定・禁止リスト (前期or後期)", type=["csv"])
    
    # 実行ボタン
    if st.sidebar.button("🚀 作成開始"):
        if not all([file_teacher, file_subject, file_req]): 
            st.error("⚠️ 必須ファイル（教員、教科、授業）が不足しています。")
            return

        with st.spinner("⏳ 最適化計算を実行中..."):
            try:
                # --------------------------------------
                # 1. データ読み込み処理
                # --------------------------------------
                
                # (1) 教員データ
                df_teacher = pd.read_csv(file_teacher)
                df_teacher['教員名'] = df_teacher['教員名'].apply(clean_name)
                teachers = df_teacher['教員名'].unique().tolist()
                
                # (2) 教科設定（連続フラグの取得）
                df_subj = pd.read_csv(file_subject)
                # 列名の空白削除
                df_subj.columns = [c.strip() for c in df_subj.columns]

                continuous_flags = {}
                
                # 【修正】列名自動検出ロジック
                # 「教科」または「教科名」が含まれる列を探す
                col_subj_name = None
                col_cont = None
                
                for c in df_subj.columns:
                    if "教科" in c:  # "教科" or "教科名"
                        col_subj_name = c
                    if "連続" in c:
                        col_cont = c
                
                if not col_subj_name:
                    st.error("エラー: 教科設定ファイルに『教科』または『教科名』の列が見つかりません。")
                    return

                # 設定読み込み
                for _, row in df_subj.iterrows():
                    s_name = str(row[col_subj_name]).strip()
                    
                    # 連続列がある場合のみ判定、なければFalse
                    is_cont_flag = False
                    if col_cont:
                        val = str(row[col_cont])
                        if "〇" in val or "TRUE" in val.upper():
                            is_cont_flag = True
                    
                    continuous_flags[s_name] = is_cont_flag
                
                # (3) 授業データ
                df_req = pd.read_csv(file_req)
                df_req.columns = [c.strip() for c in df_req.columns]
                
                req_list = []
                req_id = 0
                for _, row in df_req.iterrows():
                    cls = str(row['クラス']).strip()
                    subj = str(row['教科']).strip()
                    t1 = clean_name(row['担当教員'])
                    t2 = clean_name(row.get('担当教員２', '')) 
                    num = int(row['週コマ数'])
                    
                    if num > 0:
                        # 連続判定: 設定ファイルでTrue かつ 週2コマ以上
                        # (技術家庭科の週1コマはここでFalseになる)
                        is_cont = continuous_flags.get(subj, False)
                        if num < 2:
                            is_cont = False 
                        
                        req_list.append({
                            'id': req_id,
                            'class': cls,
                            'subject': subj,
                            't1': t1,
                            't2': t2,
                            'num': num,
                            'continuous': is_cont
                        })
                        req_id += 1

                # (4) 固定・禁止リスト
                fixed_list = []
                if file_fixed:
                    df_fix = pd.read_csv(file_fixed)
                    # 列名クリーニング
                    df_fix.columns = [c.strip() for c in df_fix.columns]
                    
                    for _, row in df_fix.iterrows():
                        target = clean_name(row['対象'])
                        day_str = row['曜日']
                        period = int(row['限'])
                        content = row['内容']
                        
                        w_map = {'月':0, '火':1, '水':2, '木':3, '金':4}
                        if day_str in w_map:
                            fixed_list.append({
                                'target': target,
                                'day': w_map[day_str],
                                'period': period,
                                'content': content
                            })

                # --------------------------------------
                # 2. 最適化実行
                # --------------------------------------
                result_file = solve_schedule(teachers, req_list, fixed_list)
                
                if result_file:
                    st.success("🎉 時間割が完成しました！")
                    st.download_button(
                        label="📥 Excelをダウンロード",
                        data=result_file,
                        file_name="完成時間割.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 解が見つかりませんでした。条件を緩和してください。")

            except Exception as e:
                st.error(f"エラー詳細: {e}")

# ==========================================
# ソルバーロジック
# ==========================================
def solve_schedule(teachers, req_list, fixed_list):
    model = cp_model.CpModel()
    DAYS = 5
    X = {}
    
    # 1. 変数作成
    for r in req_list:
        rid = r['id']
        slots = []
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                X[(rid, d, p)] = model.NewBoolVar(f'r{rid}_d{d}_p{p}')
                slots.append(X[(rid, d, p)])
        model.Add(sum(slots) == r['num'])
        
        # 連続制約
        if r['continuous'] and r['num'] == 2:
            pair_vars = []
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                # 昼休み(4-5)跨ぎNG
                pairs = [(1,2), (2,3), (3,4)]
                if p_max >= 6: pairs.append((5,6))
                
                for (p1, p2) in pairs:
                    b_pair = model.NewBoolVar(f'pair_{rid}_{d}_{p1}')
                    model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] == 2).OnlyEnforceIf(b_pair)
                    model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] != 2).OnlyEnforceIf(b_pair.Not())
                    pair_vars.append(b_pair)
            model.Add(sum(pair_vars) >= 1)

    # 2. クラス重複
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
        # 固定リスト
        for fix in fixed_list:
            if fix['target'] == t:
                d, p = fix['day'], fix['period']
                vars_here = []
                for r in t_map[t]:
                    if (r['id'], d, p) in X:
                        vars_here.append(X[(r['id'], d, p)])
                if vars_here:
                    model.Add(sum(vars_here) == 0)
        # 重複
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                vars_here = []
                for r in t_map[t]:
                    if (r['id'], d, p) in X:
                        vars_here.append(X[(r['id'], d, p)])
                if vars_here:
                    model.Add(sum(vars_here) <= 1)

    # 4. 同学年排他
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

    # 目的関数
    obj_terms = []
    for (rid, d, p), var in X.items():
        obj_terms.append(var * p)
    model.Minimize(sum(obj_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        return generate_excel(solver, X, req_list, teachers, fixed_list)
    else:
        return None

def generate_excel(solver, X, req_list, teachers, fixed_list):
    output = io.BytesIO()
    days = ['月', '火', '水', '木', '金']
    
    data_cls = {}
    data_tch = {}
    
    for fix in fixed_list:
        t = fix['target']
        data_tch[(t, fix['day'], fix['period'])] = f"【{fix['content']}】"

    for r in req_list:
        rid = r['id']
        for d in range(5):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                if (rid, d, p) in X and solver.Value(X[(rid, d, p)]) == 1:
                    txt_c = f"{r['subject']}\n{r['t1']}"
                    if r['t2']: txt_c += f"/{r['t2']}"
                    data_cls[(r['class'], d, p)] = txt_c
                    
                    txt_t = f"{r['class']} {r['subject']}"
                    data_tch[(r['t1'], d, p)] = txt_t
                    if r['t2']: data_tch[(r['t2'], d, p)] = txt_t

    rows_c = []
    all_classes = sorted(list(set(r['class'] for r in req_list)))
    for c in all_classes:
        for p in range(1, 7):
            row = {'クラス': c, '限': p}
            for di, dw in enumerate(days):
                if di == 4 and p == 6: row[dw] = ""
                else: row[dw] = data_cls.get((c, di, p), "")
            rows_c.append(row)
    df_c = pd.DataFrame(rows_c)

    rows_t = []
    for t in teachers:
        for p in range(1, 7):
            row = {'教員名': t, '限': p}
            for di, dw in enumerate(days):
                if di == 4 and p == 6: row[dw] = ""
                else: row[dw] = data_tch.get((t, di, p), "")
            rows_t.append(row)
    df_t = pd.DataFrame(rows_t)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_c.to_excel(writer, sheet_name='クラス別', index=False)
        df_t.to_excel(writer, sheet_name='教員別', index=False)
        wb = writer.book
        fmt = wb.add_format({'text_wrap': True, 'valign': 'top'})
        for ws in writer.sheets.values():
            ws.set_column('A:G', 15, fmt)

    output.seek(0)
    return output

if __name__ == "__main__":
    main()
