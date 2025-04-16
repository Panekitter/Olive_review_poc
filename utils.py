import time
import re
import openai
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

DEBUG = True

def rgb_to_hex_obj(bg):
    """
    bg は辞書 { "red": float, "green": float, "blue": float } として渡されると仮定し、
    それを #RRGGBB 形式の文字列に変換します。
    """
    r = bg.get('red', 1)
    g = bg.get('green', 1)
    b = bg.get('blue', 1)
    return "#{:02X}{:02X}{:02X}".format(int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))

def get_c_column_formatting(spreadsheet_id, sheet_title, creds_dict):
    """
    指定されたシート（sheet_title）の、C列（ヘッダー除く "C2:C"）について、
    effectiveFormat の情報を 1 回の API 呼び出しで取得し、
    シート上の行番号（2行目～）をキーとする辞書を返します。
    """
    scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    service = build('sheets', 'v4', credentials=creds)
    # 範囲は、シート名と "C2:C" として指定（ヘッダーは1行目）
    range_str = f"{sheet_title}!C2:C"
    # effectiveFormat を取得（userEnteredFormat ではなく、計算済みの効果的な書式情報）
    fields = "sheets(data(rowData(values(effectiveFormat))))"
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[range_str],
        fields=fields
    ).execute()
    formats = {}
    # result の中から対象シートを探す
    for sheet in result.get('sheets', []):
        props = sheet.get('properties', {})
        if props.get('title') == sheet_title:
            data = sheet.get('data', [])
            if not data:
                break
            rowData = data[0].get('rowData', [])
            for i, row in enumerate(rowData):
                # 実際のシート行番号は、ヘッダーが1行目なので i + 2
                row_number = i + 2
                values = row.get('values', [])
                # 範囲 "C2:C" は1列だけなので、0 番目の値が該当するセル
                if len(values) > 0:
                    effective_fmt = values[0].get('effectiveFormat', {})
                    bg = effective_fmt.get('backgroundColor', None)
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
    """
    data: ヘッダー除くのデータ行リスト（各行はリスト）
    j: data 内の 0-indexed 行番号
    前後文脈として、前行と次行の A列の値を返します。
    """
    prev_line = data[j-1][0] if j > 0 else ""
    target_line = data[j][0]
    next_line = data[j+1][0] if j+1 < len(data) else ""
    return prev_line, target_line, next_line

def is_white_background(hex_color):
    """
    文字列としての背景色 (例: "#FFFFFF") を受け取り、白であるか判定します。
    """
    return hex_color == "#FFFFFF"

def process_review_file(spreadsheet, openai_key, creds_dict):
    """
    対象ファイル内のデータ行（ヘッダー除く）のうち、C列の効果的背景色が白 (#FFFFFF) である行を対象とし、
    前後文脈とともにプロンプトを作成、GPTに依頼し、結果でシートをバッチ更新します。
    """
    openai.api_key = openai_key
    worksheet = spreadsheet.worksheet("Task")
    rows = worksheet.get_all_values()  # 全行取得（ヘッダー含む）
    if len(rows) < 2:
        print("データ行がありません。")
        return
    # ヘッダー行を除いたデータリスト
    data = rows[1:]
    total_data_rows = len(data)
    spreadsheet_id = spreadsheet.id
    sheet_title = "Task"
    # C列の書式（効果的背景色）を一括取得（実際のシート行番号がキー）
    format_dict = get_c_column_formatting(spreadsheet_id, sheet_title, creds_dict)
    
    eligible_indices = []  # data 内のインデックス (0-indexed)
    for j in range(total_data_rows):
        row_num = j + 2  # シート上の行番号
        hex_color = format_dict.get(row_num)
        if DEBUG:
            c_value = data[j][2] if len(data[j]) > 2 else ""
            print(f"Row {row_num}: Cセル値 = '{c_value}', 背景色 = {hex_color}")
        # 対象条件：背景色が白 (#FFFFFF)
        if not hex_color or not is_white_background(hex_color):
            continue
        eligible_indices.append(j)
    
    if not eligible_indices:
        print("該当する対象行は見つかりませんでした。")
        return

    # --- GPT API 用プロンプト作成 ---
    prompt = ("以下は翻訳レビュー対象データです。それぞれの行について、"
              "修正翻訳、エラー分類、エラー理由（エラー分類がotherの場合のみ）を、"
              "シート上の実際の行番号ごとに以下のフォーマットで返してください。\n")
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
