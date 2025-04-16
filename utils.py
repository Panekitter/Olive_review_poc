import time
from openai import OpenAI
from gspread_formatting import get_user_entered_format

def get_context(rows, index):
    """前後文脈の抽出（対象行の1つ上と1つ下のA列の値）"""
    prev_line = rows[index - 1][0] if index > 1 else ""
    target_line = rows[index][0]
    next_line = rows[index + 1][0] if index + 1 < len(rows) else ""
    return prev_line, target_line, next_line

def rgb_to_hex(color):
    def to_255(v): 
        return int(round(v * 255))
    r = to_255(color.red if color.red is not None else 1)
    g = to_255(color.green if color.green is not None else 1)
    b = to_255(color.blue if color.blue is not None else 1)
    return "#{:02X}{:02X}{:02X}".format(r, g, b)

def is_white_background(cell_format):
    color = cell_format.backgroundColor
    if not color:
        return False
    hex_color = rgb_to_hex(color)
    return hex_color == "#FFFFFF"

def process_review_file(spreadsheet, openai_key):
    """
    対象のファイル内の行データ（C列が空かつ背景色が白の行）を一括でGPTに渡してレビューし、
    バッチ更新で結果を書き込む。
    """
    client = OpenAI(api_key=openai_key)
    worksheet = spreadsheet.worksheet("Task")
    rows = worksheet.get_all_values()  # すべての行を取得

    # 対象行のインデックス（ヘッダーは除くので1以降）
    eligible_indices = []
    for i in range(1, len(rows)):
        # C列（インデックス2）が空白かどうかチェック
        if rows[i][2].strip():
            continue

        # C列の背景色を取得（各セルごとに呼ぶが、これをまとめて処理はできないので必要回数発生）
        try:
            cell_format = get_user_entered_format(worksheet, f"C{i+1}")
            # ※ API制限対策としては、ここで minimal sleep を入れる
            time.sleep(0.15)
            if not is_white_background(cell_format):
                continue
            eligible_indices.append(i)
        except Exception as e:
            print(f"Skipping row {i+1} due to error in background color check: {e}")
            continue

    if not eligible_indices:
        print("該当する対象行は見つかりませんでした。")
        return

    # --- まとめて1回のAPI呼び出し用プロンプト作成 ---
    # 行ごとに番号とともに、前文、本文、後文、初回訳（B列）の情報を記述
    prompt = "以下は翻訳レビュー対象データです。それぞれの行について、修正翻訳、エラー分類、エラー理由（エラー分類がotherの場合のみ）を、行番号ごとに以下のフォーマットで返してください。\n"
    prompt += "【出力フォーマット】\n"
    prompt += "行番号: 修正翻訳 | エラー分類 | エラー理由（必要な場合）\n"
    prompt += "-----------------------------\n\n"

    for i in eligible_indices:
        prev, target, next_ = get_context(rows, i)
        # i+1 は実際のシートの行番号
        prompt += f"行 {i+1}:\n"
        prompt += f"前文: {prev}\n"
        prompt += f"本文: {target}\n"
        prompt += f"後文: {next_}\n"
        prompt += f"初回日本語訳: {rows[i][1]}\n"
        prompt += "-----------------------------\n"

    # --- 1回のAPI呼び出し ---
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.5,
    )

    result_text = response.choices[0].message.content
    # 期待する出力形式の例（GPTに依頼するフォーマットに沿って）：
    # 行 2: 修正訳文 | 誤訳 | 理由（例: 原文の「...」は正しい訳ではない）
    # 行 5: 修正訳文 | 不自然 | 
    # …
    # 結果をパースする
    results = {}
    for line in result_text.splitlines():
        line = line.strip()
        if not line:
            continue
        # 例： "行 2: 修正訳文 | 誤訳 | 理由"
        if line.startswith("行"):
            try:
                # 行番号部分と本文を分割
                parts = line.split(":", 1)
                row_num_str = parts[0].replace("行", "").strip()
                row_num = int(row_num_str)
                details = parts[1].strip().split("|")
                revised = details[0].strip() if len(details) > 0 else ""
                category = details[1].strip() if len(details) > 1 else ""
                explanation = details[2].strip() if len(details) > 2 else ""
                results[row_num] = (revised, category, explanation)
            except Exception as e:
                print(f"エラー発生、行パース失敗: {line} : {e}")
                continue

    # --- バッチでシートを更新 ---
    cell_updates = []
    for i in eligible_indices:
        row_num = i + 1
        if row_num in results:
            revised, category, explanation = results[row_num]
            cell_updates.append({
                "range": f"C{row_num}",
                "values": [[revised]]
            })
            cell_updates.append({
                "range": f"D{row_num}",
                "values": [[category]]
            })
            if category.lower() == "other":
                cell_updates.append({
                    "range": f"E{row_num}",
                    "values": [[explanation]]
                })
    if cell_updates:
        worksheet.batch_update(cell_updates)
    else:
        print("更新するセルがありませんでした。")
