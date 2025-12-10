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
    1. 左側のサイドバー（スマホなら上部）から必要なCSVファイルを5つアップロードしてください。
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
        if not all([file_teacher, file_subject, file_req]): # 固定リストは任意でも可とするが基本は必須
            st.error("⚠️ 必須ファイル（教員、教科、授業）が不足しています。")
            return

        with st.spinner("⏳ 最適化計算を実行中...（これには1〜2分かかる場合があります）"):
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
                continuous_flags = {}
                # 列名検索（"連続"が含まれる列を探す）
                col_cont = next((c for c in df_subj.columns if "連続" in c), None)
                
                if col_cont:
                    for _, row in df_subj.iterrows():
                        s_name = str(row['教科名']).strip()
                        val = str(row[col_cont])
                        # 〇, TRUE, True なら連続希望とみなす
                        if "〇" in val or "TRUE" in val.upper():
                            continuous_flags[s_name] = True
                        else:
                            continuous_flags[s_name] = False
                
                # (3) 授業データ
                df_req = pd.read_csv(file_req)
                # カラム名空白除去
                df_req.columns = [c.strip() for c in df_req.columns]
                
                req_list = []
                req_id = 0
                for _, row in df_req.iterrows():
                    cls = str(row['クラス']).strip()
                    subj = str(row['教科']).strip()
                    t1 = clean_name(row['担当教員'])
                    t2 = clean_name(row.get('担当教員２', '')) # 列がない場合に備える
                    num = int(row['週コマ数'])
                    
                    if num > 0:
                        # 【修正箇所】連続設定の判定ロジック
                        # 設定ファイルでTrue、かつ「週2コマ以上」の場合のみ連続とする
                        is_cont = continuous_flags.get(subj, False)
                        if num < 2:
                            is_cont = False # 1コマなら強制的に単発扱い
                        
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
                # 2. 最適化モデル構築 & 解決
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
                st.error(f"エラーが発生しました: {e}")
                st.write("詳細:", e)

# ==========================================
# ソルバーロジック
# ==========================================
def solve_schedule(teachers, req_list, fixed_list):
    model = cp_model.CpModel()
    DAYS = 5
    
    # 変数定義 X[req_id, day, period]
    X = {}
    
    # 1. 変数作成と基本制約（コマ数確保）
    for r in req_list:
        rid = r['id']
        slots = []
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                X[(rid, d, p)] = model.NewBoolVar(f'r{rid}_d{d}_p{p}')
                slots.append(X[(rid, d, p)])
        model.Add(sum(slots) == r['num'])
        
        # 連続制約（簡易版: 同じ日に2コマあるなら連続させる）
        if r['continuous'] and r['num'] == 2:
            # ペア変数の作成
            pair_vars = []
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                # 昼休み跨ぎ(4-5)を除く連続ペア
                pairs = [(1,2), (2,3), (3,4)]
                if p_max >= 6:
                    pairs.append((5,6))
                
                for (p1, p2) in pairs:
                    b_pair = model.NewBoolVar(f'pair_{rid}_{d}_{p1}')
                    # 両方1ならpairも1
                    model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] == 2).OnlyEnforceIf(b_pair)
                    model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] != 2).OnlyEnforceIf(b_pair.Not())
                    pair_vars.append(b_pair)
            
            # 少なくとも1つはペアであること
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

    # 3. 教員重複禁止 & 固定リスト
    # 担当授業のマッピング
    t_map = {t: [] for t in teachers}
    for r in req_list:
        if r['t1'] in teachers: t_map[r['t1']].append(r)
        if r['t2'] in teachers: t_map[r['t2']].append(r)
    
    for t in teachers:
        # 固定リスト（禁止時間）
        for fix in fixed_list:
            if fix['target'] == t:
                # その時間は授業入れない
                d, p = fix['day'], fix['period']
                vars_here = []
                for r in t_map[t]:
                    if (r['id'], d, p) in X:
                        vars_here.append(X[(r['id'], d, p)])
                if vars_here:
                    model.Add(sum(vars_here) == 0)
        
        # 重複禁止
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                vars_here = []
                for r in t_map[t]:
                    if (r['id'], d, p) in X:
                        vars_here.append(X[(r['id'], d, p)])
                if vars_here:
                    model.Add(sum(vars_here) <= 1)

    # 4. 同学年排他（体育・理科など）
    # 簡易的に学年抽出
    grade_reqs = {} # "1": [reqs], "2": [reqs]
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

    # 5. 目的関数（午前優先など）
    obj_terms = []
    for (rid, d, p), var in X.items():
        # pが大きいほどペナルティ（午後の授業を減らしたい＝午前優先）
        obj_terms.append(var * p)
    
    model.Minimize(sum(obj_terms))

    # ソルバー実行
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
    
    # データを整形
    # assigned[(cls, d, p)] = "国語\n田中"
    data_cls = {}
    data_tch = {}
    
    # 固定リスト（教員用）
    for fix in fixed_list:
        t = fix['target']
        key = (t, fix['day'], fix['period'])
        data_tch[key] = f"【{fix['content']}】"

    for r in req_list:
        rid = r['id']
        for d in range(5):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                if (rid, d, p) in X and solver.Value(X[(rid, d, p)]) == 1:
                    # クラス用
                    txt_c = f"{r['subject']}\n{r['t1']}"
                    if r['t2']: txt_c += f"/{r['t2']}"
                    data_cls[(r['class'], d, p)] = txt_c
                    
                    # 教員用
                    txt_t = f"{r['class']} {r['subject']}"
                    data_tch[(r['t1'], d, p)] = txt_t
                    if r['t2']: data_tch[(r['t2'], d, p)] = txt_t

    # DataFrame化
    # 1. クラス別
    rows_c = []
    all_classes = sorted(list(set(r['class'] for r in req_list)))
    for c in all_classes:
        for p in range(1, 7):
            row = {'クラス': c, '限': p}
            for di, dw in enumerate(days):
                if di == 4 and p == 6:
                     row[dw] = ""
                else:
                    row[dw] = data_cls.get((c, di, p), "")
            rows_c.append(row)
    df_c = pd.DataFrame(rows_c)

    # 2. 教員別
    rows_t = []
    for t in teachers:
        for p in range(1, 7):
            row = {'教員名': t, '限': p}
            for di, dw in enumerate(days):
                 if di == 4 and p == 6:
                     row[dw] = ""
                 else:
                    # 既に固定リストが入っているかも確認しつつ
                    val = data_tch.get((t, di, p), "")
                    row[dw] = val
            rows_t.append(row)
    df_t = pd.DataFrame(rows_t)

    # Excel書き出し
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_c.to_excel(writer, sheet_name='クラス別', index=False)
        df_t.to_excel(writer, sheet_name='教員別', index=False)
        
        # 見た目の調整（改行有効化など）
        workbook = writer.book
        fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        
        # 全セルに適用
        for worksheet in writer.sheets.values():
            worksheet.set_column('A:G', 15, fmt)

    output.seek(0)
    return output

if __name__ == "__main__":
    main()
