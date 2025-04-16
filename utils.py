import time
import re
import openai
from gspread_formatting import get_user_entered_format  # 個別取得は個々のセル用

DEBUG = True

def get_user_entered_formats(worksheet, range_str):
    """
    指定された範囲 (例: "C2:C{total_rows}") の各セル書式情報を取得して、
    セルアドレスをキーとする辞書を返す関数です。
    ※ 一括で取得する関数は gspread_formatting には用意されていないため、
      指定範囲のセルを個別に取得し、辞書にまとめる実装例です。
    """
    # 範囲文字列が "C<start>:C<end>" 形式であることを確認
    m = re.match(r"C(\d+):C(\d+)", range_str)
    if not m:
        raise ValueError(f"セル範囲の形式が正しくありません: {range_str}")
    start = int(m.group(1))
    end = int(m.group(2))
    formats = {}
    for row in range(start, end + 1):
        cell_address = f"C{row}"
        try:
            fmt = get_user_entered_format(worksheet, cell_address)
            formats[cell_address] = fmt
            time.sleep(0.1)  # API呼び出し間の待機でレート制限軽減
        except Exception as e:
            formats[cell_address] = None
            print(f"{cell_address} 書式取得エラー: {e}")
    return formats

def get_context(rows, index):
    """対象行の前後文脈（A列）の値を取得する"""
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
    """背景色が完全な白 (#FFFFFF) か判定する"""
    color = cell_format.backgroundColor
    if not color:
        return False
    hex_color = rgb_to_hex(color)
    return hex_color == "#FFFFFF"

def process_review_file(spreadsheet, openai_key):
    """
    対象のファイル内の全行（ヘッダー除く）のうち、C列の背景色が白 (#FFFFFF) 
    の行を対象とし、まとめて1回のGPT呼び出しで各行のレビュー結果を取得して、
    batch_update() によりシートに反映する。
    """
    openai.api_key = openai_key
    worksheet = spreadsheet.worksheet("Task")
    rows = worksheet.get_all_values()  # シート全行取得
    total_rows = len(rows)
    # バッチで C列 (範囲: C2 から C{total_rows}) の書式情報を取得
    range_str = f"C2:C{total_rows}"
    try:
        cell_formats = get_user_entered_formats(worksheet, range_str)
    except Exception as e:
        print(f"Error in batch retrieving formats for range {range_str}: {e}")
        return

    eligible_indices = []
    for i in range(1, total_rows):
        cell_address = f"C{i+1}"
        cell_format = cell_formats.get(cell_address)
        if not cell_format:
            if DEBUG:
                print(f"Row {i+1}: No format data for {cell_address}")
            continue

        try:
            hex_color = rgb_to_hex(cell_format.backgroundColor) if cell_format.backgroundColor else "None"
        except Exception as e:
            hex_color = f"取得エラー: {e}"
        if DEBUG:
            c_value = rows[i][2] if len(rows[i]) > 2 else ""
            print(f"Row {i+1}: Cセル値 = '{c_value}', 背景色 = {hex_color}")
        # 対象条件：背景色が #FFFFFF
        if not is_white_background(cell_format):
            continue
        eligible_indices.append(i)

    if not eligible_indices:
        print("該当する対象行は見つかりませんでした。")
        return

    # --- まとめて1回のGPT API呼び出し用プロンプト作成 ---
    prompt = ("以下は翻訳レビュー対象データです。それぞれの行について、"
              "修正翻訳、エラー分類、エラー理由（エラー分類がotherの場合のみ）を、"
              "行番号ごとに以下のフォーマットで返してください。\n")
    prompt += "【出力フォーマット】\n"
    prompt += "行番号: 修正翻訳 | エラー分類 | エラー理由（必要な場合）\n"
    prompt += "-----------------------------\n\n"

    for i in eligible_indices:
        prev, target, next_ = get_context(rows, i)
        prompt += f"行 {i+1}:\n"
        prompt += f"前文: {prev}\n"
        prompt += f"本文: {target}\n"
        prompt += f"後文: {next_}\n"
        prompt += f"初回日本語訳: {rows[i][1]}\n"
        prompt += "-----------------------------\n"

    # --- 1回のGPT API呼び出し ---
    response = openai.ChatCompletion.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.5,
    )

    result_text = response.choices[0].message.content
    results = {}
    for line in result_text.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith("行"):
            try:
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

    # --- バッチ更新でシートに書き込み ---
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
