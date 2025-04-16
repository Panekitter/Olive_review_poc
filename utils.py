import time
import re
import openai
import json
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

DEBUG = True

# ----- セルの背景色を dict から hex 表記に変換する関数 ----- 
def rgb_to_hex_obj(bg):
    # bg は辞書。たとえば { "red": 1, "green": 1, "blue": 1 } なら白
    r = bg.get('red', 1)
    g = bg.get('green', 1)
    b = bg.get('blue', 1)
    return "#{:02X}{:02X}{:02X}".format(int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))

# ----- 指定されたシートの範囲（例："Task!C2:C"）の書式情報を一括取得する -----
def get_c_column_formatting(spreadsheet_id, sheet_title, creds_dict):
    # 認証情報を作成
    scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    service = build('sheets', 'v4', credentials=creds)
    # 対象範囲はシート名と "C2:C" として、ヘッダー（1行目）は除く
    range_str = f"{sheet_title}!C2:C"
    # fields パラメータを設定して、userEnteredFormat 情報のみを取得する
    fields = "sheets(data(rowData(values(userEnteredFormat))))"
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[range_str],
        fields=fields
    ).execute()
    formats = {}
    # result には sheets のリストがあります。対象シートタイトルのものを探す。
    for sheet in result.get('sheets', []):
        props = sheet.get('properties', {})
        if props.get('title') == sheet_title:
            # data はリスト（通常 1つ）
            data = sheet.get('data', [])
            if not data:
                break
            rowData = data[0].get('rowData', [])
            for i, row in enumerate(rowData):
                # 実際のシート行番号は、ヘッダーが1行目なので i+2
                row_number = i + 2
                values = row.get('values', [])
                if len(values) > 2 and 'userEnteredFormat' in values[2]:
                    bg = values[2]['userEnteredFormat'].get('backgroundColor', None)
                    if bg is not None:
                        try:
                            hex_color = rgb_to_hex_obj(bg)
                        except Exception as e:
                            hex_color = f"Error: {e}"
                        formats[row_number] = hex_color
                    else:
                        formats[row_number] = None
                else:
                    formats[row_number] = None
            break
    return formats

def get_context(data, j):
    """data: ヘッダー除くデータ行リスト, j: 0-index（データ行の番号）"""
    prev_line = data[j-1][0] if j > 0 else ""
    target_line = data[j][0]
    next_line = data[j+1][0] if j+1 < len(data) else ""
    return prev_line, target_line, next_line

def process_review_file(spreadsheet, openai_key, creds_dict):
    """
    対象ファイル内のデータ行（ヘッダー除く）のうち、C列の背景色が白 (#FFFFFF) の行を対象として、
    前後文脈を含むプロンプトを作成し、GPTにレビュー依頼、得られた結果でシートを更新する。
    """
    openai.api_key = openai_key
    worksheet = spreadsheet.worksheet("Task")
    rows = worksheet.get_all_values()
    if len(rows) < 2:
        print("データ行がありません。")
        return
    # ヘッダー行を除いたデータ部分
    data = rows[1:]
    total_data_rows = len(data)
    spreadsheet_id = spreadsheet.id
    sheet_title = "Task"
    # 1回の API 呼び出しで列Cの書式情報をまとめて取得
    format_dict = get_c_column_formatting(spreadsheet_id, sheet_title, creds_dict)
    
    eligible_indices = []  # 0-indexed data 行番号
    for j in range(total_data_rows):
        # 実際のシート行番号 = j+2
        row_num = j + 2
        # 書式情報がない場合は、今回は対象外とする
        hex_color = format_dict.get(row_num)
        if DEBUG:
            c_value = data[j][2] if len(data[j]) > 2 else ""
            print(f"Row {row_num}: Cセル値 = '{c_value}', 背景色 = {hex_color}")
        if hex_color != "#FFFFFF":
            continue
        eligible_indices.append(j)
    
    if not eligible_indices:
        print("該当する対象行は見つかりませんでした。")
        return

    # --- GPT API 用プロンプト作成 ---
    prompt = ("以下は翻訳レビュー対象データです。それぞれの行について、"
            "以下の【エラー分類の選択肢】の中から最も該当するものを1つ選び、"
            "修正翻訳、エラー分類、エラー理由（エラー分類が 'Other' の場合のみ）を返してください。\n\n")
    prompt += "ただし、エラーの分類は翻訳を修正した場合のみ返してください。\n"
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
        row_num = j + 2
        prompt += f"行 {row_num}:\n"
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
        row_num = j + 2
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