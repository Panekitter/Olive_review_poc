import time
import re
import openai
from gspread_formatting import get_user_entered_format

DEBUG = True

def get_user_entered_formats(worksheet, range_str):
    """
    指定された範囲（例："C2:C{total_rows}"）の各セル書式情報を取得し、
    セルアドレス（例 "C2", "C3", …）をキーとする辞書を返す関数です。
    ※ gspread_formatting には一括取得関数がないため、各セル毎に個別取得してまとめます。
    """
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
            time.sleep(0.1)  # レート制限対策
        except Exception as e:
            formats[cell_address] = None
            print(f"{cell_address} 書式取得エラー: {e}")
    return formats

def get_context(data, j):
    """
    data: ヘッダーを除いたデータ行のリスト
    j: data における行番号 (0-indexed)
    前後文脈として、前のデータ行と次のデータ行の A列の値を返す。
    ヘッダーは除くので、利用できない場合は空文字を返します。
    """
    prev_line = data[j-1][0] if j > 0 else ""
    target_line = data[j][0]
    next_line = data[j+1][0] if j+1 < len(data) else ""
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
    対象のファイル内のデータ行（ヘッダー除く）のうち、C列の背景色が白 (#FFFFFF)
    のセルを対象として、前後文脈を利用した翻訳レビューのプロンプトを構築し、
    1回のGPT呼び出しでレビュー結果を取得し、バッチ更新でシートに反映する。
    
    ※ 書式情報が取得できなかった場合は、背景色は白 (#FFFFFF) と仮定します。
    """
    openai.api_key = openai_key
    worksheet = spreadsheet.worksheet("Task")
    rows = worksheet.get_all_values()  # 全行取得（ヘッダー含む）
    if len(rows) < 2:
        print("データ行がありません。")
        return

    # header を除くデータ行のみを抽出
    data = rows[1:]
    total_data_rows = len(data)
    # C列 (データ部分) の書式情報取得範囲は「C2:C{total_data_rows+1}」
    range_str = f"C2:C{total_data_rows+1}"
    try:
        cell_formats = get_user_entered_formats(worksheet, range_str)
    except Exception as e:
        print(f"Error in batch retrieving formats for range {range_str}: {e}")
        return

    eligible_indices = []
    for j in range(total_data_rows):
        cell_address = f"C{j+2}"  # ヘッダー行が1行目なので、最初のデータ行は行2
        cell_format = cell_formats.get(cell_address)
        # 書式情報が取得できない場合は、デフォルト背景＝白 (#FFFFFF) と仮定
        if cell_format is None:
            is_white = True
            hex_color = "#FFFFFF"
        else:
            try:
                hex_color = rgb_to_hex(cell_format.backgroundColor) if cell_format.backgroundColor else "None"
            except Exception as e:
                hex_color = f"取得エラー: {e}"
            is_white = is_white_background(cell_format)
        if DEBUG:
            c_value = data[j][2] if len(data[j]) > 2 else ""
            # 実際のシート行番号は j+2
            print(f"Row {j+2}: Cセル値 = '{c_value}', 背景色 = {hex_color}")
        if not is_white:
            continue
        eligible_indices.append(j)

    if not eligible_indices:
        print("該当する対象行は見つかりませんでした。")
        return

    # --- GPT API 用プロンプト作成 ---
    prompt = ("以下は翻訳レビュー対象データです。それぞれの行について、"
            "以下の【エラー分類の選択肢】の中から最も該当するものを1つ選び、"
            "修正翻訳、エラー分類、エラー理由（エラー分類が 'Other' の場合のみ）を返してください。\n\n")
    prompt += "【エラー分類の選択肢】\n"
    prompt += "Major Error - Difficult or impossible to understand\n"
    prompt += "Major Error - Meaning is not preserved\n"
    prompt += "Major Error - Translation is offensive (when the original was not)\n"
    prompt += "Major Error - Translation is offensive - misgendering (when the original was not)\n"
    prompt += "Major Error - Source text is incomprehensible\n"
    prompt += "Minor Error - Meaning preserved but awkward or slightly difficult to understand\n"
    prompt += "Minor Error - Source and translation text are misaligned\n"
    prompt += "Other\n"
    prompt += "※ 'misaligned' とは、該当のセルの翻訳が前後のセルとずれて入ってしまっているというエラーを意味します。\n\n"
    prompt += "【出力フォーマット】\n"
    prompt += "行番号: 修正翻訳 | エラー分類 | エラー理由（必要な場合）\n"
    prompt += "-----------------------------\n\n"


    for j in eligible_indices:
        prev, target, next_ = get_context(data, j)
        # シート上の行番号は j+2
        prompt += f"行 {j+2}:\n"
        prompt += f"前文: {prev}\n"
        prompt += f"本文: {target}\n"
        prompt += f"後文: {next_}\n"
        prompt += f"初回日本語訳: {data[j][1]}\n"
        prompt += "-----------------------------\n"

    # --- GPT API 呼び出し ---
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
    for j in eligible_indices:
        row_num = j+2
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
