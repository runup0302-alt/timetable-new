# ==========================================
# 時間割作成システム (修正版)
# ==========================================
import pandas as pd
import numpy as np
from ortools.sat.python import cp_model
import openpyxl

# ------------------------------------------
# 1. 設定エリア (Config)
# ------------------------------------------
TERM = "後期"  # "前期" or "後期"
MAX_TIME_LIMIT = 120.0  # 計算時間の上限(秒)

# ファイル名設定
FILE_TEACHER = "教員データ - 教員.csv"
FILE_SUBJECT = "教科設定 - 年間.csv"
FILE_CLASS_REQ = f"授業データ - {TERM}.csv"
FILE_FIXED = f"固定・禁止リスト - {TERM}.csv"

# 重み付け（ペナルティの大きさ）
WEIGHTS = {
    'minimize_days': 20,      # 同じ教科をなるべく分散させる
    'morning_class': 10,      # 主要教科を午前に
    'teacher_dispersion': 50, # 教員の1日のコマ数を平準化
    'fill_morning': 5,       # 午前を埋める（空きコマ減）
    'pattern_balance': 10     # その他のバランス
}

# 表記ゆれ辞書 (CSVの入力ミスをここで吸収)
NAME_CORRECTIONS = {
    "ニシダ": "ニシタ",
    "オオシマ": "オシマ",
    # 必要に応じて追加してください
}

# ------------------------------------------
# 2. データ読み込み・前処理
# ------------------------------------------
def clean_name(name):
    """名前の空白除去と表記ゆれ補正"""
    if pd.isna(name) or name == "":
        return ""
    name = str(name).replace(" ", "").replace("　", "")
    return NAME_CORRECTIONS.get(name, name)

def load_data():
    print("📂 データを読み込んでいます...")
    
    # 1. 教員データ
    try:
        df_teacher = pd.read_csv(FILE_TEACHER)
        df_teacher['教員名'] = df_teacher['教員名'].apply(clean_name)
        teachers = df_teacher['教員名'].unique().tolist()
    except Exception as e:
        print(f"❌ 教員データの読み込みに失敗: {e}")
        return None, None, None, None

    # 2. 教科設定 (New列などは無視、連続設定を取得)
    try:
        # 必要な列だけ読むか、全部読んでから処理
        df_subj_settings = pd.read_csv(FILE_SUBJECT)
        # 連続設定の読み取り (〇/TRUEならTrue, それ以外False)
        continuous_flags = {}
        
        # 列名ゆれ対応
        col_cont = None
        for c in df_subj_settings.columns:
            if "連続" in c:
                col_cont = c
                break
        
        if col_cont:
            for _, row in df_subj_settings.iterrows():
                subj = str(row['教科名']).strip()
                val = str(row[col_cont])
                # 〇またはTRUEなら連続扱い
                if "〇" in val or "TRUE" in val.upper() or "True" in val:
                    continuous_flags[subj] = True
                else:
                    continuous_flags[subj] = False
        else:
            print("⚠️ 教科設定に「連続」列が見つかりません。デフォルト設定を使います。")
            continuous_flags = {} # 空なら適用しない

    except Exception as e:
        print(f"⚠️ 教科設定の読み込みエラー（標準設定で続行）: {e}")
        continuous_flags = {}

    # 3. 授業データ
    try:
        df_req = pd.read_csv(FILE_CLASS_REQ)
        # カラム名のクリーニング
        df_req.columns = [c.strip() for c in df_req.columns]
        
        req_list = []
        req_id = 0
        
        for _, row in df_req.iterrows():
            cls = str(row['クラス']).strip()
            subj = str(row['教科']).strip()
            t1 = clean_name(row['担当教員'])
            t2 = clean_name(row.get('担当教員２', '')) # 列がない場合に対応
            num = int(row['週コマ数'])
            
            if num > 0:
                # 連続設定の判定: 設定ファイルでTrue かつ コマ数が2以上
                # ★ここでお客様の要望通り「技術家庭」の×設定が効きます
                is_continuous = continuous_flags.get(subj, False)
                if num < 2:
                    is_continuous = False # 1コマなら物理的に連続不可
                
                req_list.append({
                    'id': req_id,
                    'class': cls,
                    'subject': subj,
                    't1': t1,
                    't2': t2,
                    'num': num,
                    'continuous': is_continuous
                })
                req_id += 1
                
    except Exception as e:
        print(f"❌ 授業データの読み込みに失敗: {e}")
        return None, None, None, None

    # 4. 固定・禁止リスト
    fixed_list = []
    try:
        df_fix = pd.read_csv(FILE_FIXED)
        for _, row in df_fix.iterrows():
            target = clean_name(row['対象'])
            day = row['曜日']
            period = row['限']
            content = row['内容']
            
            # 曜日変換 (月->0, 火->1...)
            w_map = {'月':0, '火':1, '水':2, '木':3, '金':4}
            d_idx = w_map.get(day, -1)
            
            if d_idx != -1:
                fixed_list.append({
                    'target': target,
                    'day': d_idx,
                    'period': int(period),
                    'content': content
                })
    except Exception as e:
        print(f"⚠️ 固定リストなし、または読み込みエラー: {e}")

    return teachers, req_list, fixed_list, continuous_flags

# ------------------------------------------
# 3. 最適化モデル構築
# ------------------------------------------
def solve_schedule(teachers, req_list, fixed_list):
    model = cp_model.CpModel()
    
    # 基本定数
    DAYS = 5 # 月〜金
    PERIODS = 6 # 最大6限
    
    # 変数作成: X[授業ID, 曜日, 限] = 1 (授業が入る)
    X = {}
    # 授業IDごとの配置情報を保存する辞書
    req_vars = {} 

    print("🧩 モデルを構築中...")

    for r in req_list:
        rid = r['id']
        req_vars[rid] = []
        
        # コマ数分配置
        # 全スロット分の変数を作る
        slots = []
        for d in range(DAYS):
            # 金曜は5限まで
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                X[(rid, d, p)] = model.NewBoolVar(f'req_{rid}_d{d}_p{p}')
                slots.append(X[(rid, d, p)])
        
        # 制約: 指定コマ数分配置する
        model.Add(sum(slots) == r['num'])
        req_vars[rid] = slots

        # --- 連続授業の制約 (ニコイチ) ---
        if r['continuous']:
            # 連続は「2コマ単位」で扱う簡易ロジック
            # 日ごとに、(p, p+1) のペアが少なくとも1つあることなどを強制するのではなく
            # "同じ日に2コマあるなら連続していなければならない" という制約を加える
            
            # 簡易実装: 週2コマなら「1セットの連続」がある
            if r['num'] == 2:
                # どこか1箇所で連続している
                # 連続可能な箇所: (d, 1-2), (d, 2-3), (d, 3-4), (d, 5-6) ※昼休み跨ぎNG
                possible_pairs = []
                for d in range(DAYS):
                    p_max = 5 if d == 4 else 6
                    # 昼休み(4-5)を除くペア
                    pairs = [(1,2), (2,3), (3,4)]
                    if p_max >= 6:
                        pairs.append((5,6))
                    
                    for (p1, p2) in pairs:
                        # 両方1ならOK
                        pair_bool = model.NewBoolVar(f'pair_{rid}_{d}_{p1}{p2}')
                        model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] == 2).OnlyEnforceIf(pair_bool)
                        model.Add(X[(rid, d, p1)] + X[(rid, d, p2)] != 2).OnlyEnforceIf(pair_bool.Not())
                        possible_pairs.append(pair_bool)
                
                # 少なくとも1組はペアである (sum >= 1)
                # コマ数が2なので、ペアが1つあればそれで全て
                model.Add(sum(possible_pairs) >= 1)

    # --- クラスごとの制約 ---
    # 1. 同じクラスは同時刻に1つだけ
    # クラスリスト抽出
    classes = sorted(list(set(r['class'] for r in req_list)))
    for cls in classes:
        cls_reqs = [r for r in req_list if r['class'] == cls]
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                # このクラスのこの時間の授業変数の合計 <= 1
                cls_vars = [X[(r['id'], d, p)] for r in cls_reqs]
                model.Add(sum(cls_vars) <= 1)
                
    # 2. 教員の重複禁止 & 固定禁止リスト
    # 教員ごとの担当授業を集める
    teacher_assignments = {t: [] for t in teachers}
    for r in req_list:
        if r['t1'] in teachers:
            teacher_assignments[r['t1']].append(r)
        if r['t2'] and r['t2'] in teachers:
            teacher_assignments[r['t2']].append(r)
            
    for t in teachers:
        t_reqs = teacher_assignments[t]
        
        # 固定禁止リストの適用
        # その教員に関連する禁止時間
        for fix in fixed_list:
            # 対象が「教員名」または「部会名(全員)」など
            # ここでは簡易的に教員名マッチ or 全教員対象の場合を考慮
            # ※本来は「部会」判定ロジックが必要だが、まずは名前一致で弾く
            if fix['target'] == t:
                d = fix['day']
                p = fix['period']
                # その時間は授業禁止 -> 変数の和を0にする
                # ただし、授業変数が存在する場合のみ
                vars_at_slot = []
                for r in t_reqs:
                    if (r['id'], d, p) in X:
                        vars_at_slot.append(X[(r['id'], d, p)])
                if vars_at_slot:
                    model.Add(sum(vars_at_slot) == 0)

        # 重複禁止
        for d in range(DAYS):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                # この時間にこの先生が入っている授業変数の和 <= 1
                vars_at_slot = []
                for r in t_reqs:
                    if (r['id'], d, p) in X:
                        vars_at_slot.append(X[(r['id'], d, p)])
                if vars_at_slot:
                    model.Add(sum(vars_at_slot) <= 1)

    # 3. 同学年排他（体育、理科など）
    # 学年ごとにクラスをグルーピング
    grade_map = {}
    for cls in classes:
        # "1-1" -> grade "1"
        g = cls.split('-')[0]
        if g not in grade_map: grade_map[g] = []
        grade_map[g].append(cls)
        
    exclusive_subjects = ["体育", "理科", "音楽", "美術"]
    for g, g_classes in grade_map.items():
        for subj in exclusive_subjects:
            # この学年のこの教科の授業IDリスト
            target_reqs = [r for r in req_list if r['class'] in g_classes and (subj in r['subject'] or "音美" in r['subject'])]
            
            for d in range(DAYS):
                p_max = 5 if d == 4 else 6
                for p in range(1, p_max + 1):
                    # 同時実施不可なら <= 1
                    # 施設数に応じて調整可能（体育館が2つあるなら <= 2）
                    # ここでは厳しく <= 1 とする
                    vars_at_slot = [X[(r['id'], d, p)] for r in target_reqs if (r['id'], d, p) in X]
                    if vars_at_slot:
                        model.Add(sum(vars_at_slot) <= 1)

    # --- 目的関数（ソフト制約） ---
    obj_terms = []
    
    # バランス: 同じクラス・教科はなるべく連日続けない、等
    # ここは簡易的に「教員の空きコマ分散」などをスコア化する例
    # 実際にはご要望の「1日4教科まで」などをここに追加します
    
    # とりあえず「解を見つけること」を最優先にするため、目的関数はシンプルに設定
    # 授業がなるべく前の方（1限〜）に入るように重みづけ
    for (rid, d, p), var in X.items():
        # pが大きいほどペナルティ（午前に詰めたい）
        obj_terms.append(var * (p * WEIGHTS['morning_class']))

    model.Minimize(sum(obj_terms))

    # ------------------------------------------
    # 4. ソルバー実行
    # ------------------------------------------
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = MAX_TIME_LIMIT
    print("⏳ 計算を開始しました...")
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        print(f"✅ 解が見つかりました！ ({solver.ObjectiveValue()})")
        return export_excel(solver, X, req_list, teachers, fixed_list)
    else:
        print("❌ 解が見つかりませんでした。制約が厳しすぎる可能性があります。")
        return None

# ------------------------------------------
# 5. Excel出力
# ------------------------------------------
def export_excel(solver, X, req_list, teachers, fixed_list):
    # データフレーム用配列
    # クラス別
    data_class = []
    # 教員別
    data_teacher = [] # これはあとでピボットする

    days = ['月', '火', '水', '木', '金']
    
    # 授業配置を取り出す
    assigned_map = {} # (class, day, period) -> info
    teacher_map = {}  # (teacher, day, period) -> info

    # 固定リストの内容をマッピング
    for fix in fixed_list:
        t = fix['target']
        d = fix['day']
        p = fix['period']
        c = fix['content']
        teacher_map[(t, d, p)] = f"【{c}】"

    for r in req_list:
        rid = r['id']
        for d in range(5):
            p_max = 5 if d == 4 else 6
            for p in range(1, p_max + 1):
                if (rid, d, p) in X and solver.Value(X[(rid, d, p)]) == 1:
                    # クラス向け文字列
                    info_c = f"{r['subject']}\n{r['t1']}"
                    if r['t2']: info_c += f"/{r['t2']}"
                    
                    assigned_map[(r['class'], d, p)] = info_c
                    
                    # 教員向け文字列
                    info_t = f"{r['class']} {r['subject']}"
                    
                    # T1用
                    teacher_map[(r['t1'], d, p)] = info_t
                    # T2用
                    if r['t2']:
                         teacher_map[(r['t2'], d, p)] = info_t

    # クラス別シート作成
    rows_c = []
    classes = sorted(list(set(r['class'] for r in req_list)))
    for cls in classes:
        for p in range(1, 7):
            row = {'クラス': cls, '限': p}
            for d_idx, day_name in enumerate(days):
                # 金曜6限は除外（表示したいなら空文字）
                if d_idx == 4 and p == 6:
                    row[day_name] = ""
                else:
                    row[day_name] = assigned_map.get((cls, d_idx, p), "")
            rows_c.append(row)
            
    df_out_class = pd.DataFrame(rows_c)
    
    # 教員別シート作成
    rows_t = []
    for t in teachers:
        for p in range(1, 7):
            row = {'教員名': t, '限': p}
            for d_idx, day_name in enumerate(days):
                if d_idx == 4 and p == 6:
                    row[day_name] = ""
                else:
                    row[day_name] = teacher_map.get((t, d_idx, p), "")
            rows_t.append(row)
    
    df_out_teacher = pd.DataFrame(rows_t)

    # 保存
    output_file = '完成時間割.xlsx'
    with pd.ExcelWriter(output_file) as writer:
        df_out_class.to_excel(writer, sheet_name='クラス別', index=False)
        df_out_teacher.to_excel(writer, sheet_name='教員別', index=False)
    
    print(f"🎉 '{output_file}' が作成されました！ダウンロードしてください。")
    return output_file

# ==========================================
# 実行部
# ==========================================
if __name__ == "__main__":
    t_data, r_data, f_data, c_flags = load_data()
    if t_data and r_data:
        solve_schedule(t_data, r_data, f_data)
