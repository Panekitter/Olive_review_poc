import time
import re
import openai
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

DEBUG = True

def rgb_to_hex_obj(bg):
    """
    bg は辞書 { "red": float, "green": float, "blue": float } として渡されると仮定し、
    それを "#RRGGBB" 形式の文字列に変換します。
    """
    r = bg.get('red', 1)
    g = bg.get('green', 1)
    b = bg.get('blue', 1)
    return "#{:02X}{:02X}{:02X}".format(int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))

def get_c_column_formatting(spreadsheet_id, sheet_title, creds_dict):
    """
    指定されたシート（sheet_title）の、C列（ヘッダー除く "C2:C"）
    に対し、userEnteredFormat の情報を 1 回の API 呼び出しで取得し、
    シート上の行番号（2 行目～）をキーとする辞書を返します。
    
    ※ セルに対して明示的に書式が設定されていない場合は、書式情報は返りません。
    """
    scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    service = build('sheets', 'v4', credentials=creds)
    # 範囲は "Task!C2:C"（ヘッダーは1行目なので除く）
    range_str = f"{sheet_title}!C2:C"
    # fields には userEnteredFormat のみを指定
    fields = "sheets(data(rowData(values(userEnteredFormat))))"
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[range_str],
        fields=fields
    ).execute()
    formats = {}
    # 対象シートを探す
    for sheet in result.get('sheets', []):
        props = sheet.get('properties', {})
        if props.get('title') == sheet_title:
            data = sheet.get('data', [])
            if not data:
                break
            rowData = data[0].get('rowData', [])
            for i, row in enumerate(rowData):
                # 実際のシート行番号（ヘッダーが1行目の場合）＝ i + 2
                row_number = i + 2
                values = row.get('values', [])
                # 範囲 "C2:C" は1列のみなので、値は先頭の要素
                if len(values) > 0 and 'userEnteredFormat' in values[0]:
                    fmt = values[0]['userEnteredFormat']
                    bg = fmt.get('backgroundColor', None)
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
    data: ヘッダー除くデータ行のリスト（各行はリスト）
    j: data 内の 0-indexed 行番号
    前後文脈として、前の行および次の行の A列の値を返します。
    """
    prev_line = data[j-1][0] if j > 0 else ""
    target_line = data[j][0]
    next_line = data[j+1][0] if j+1 < len(data) else ""
    return prev_line, target_line, next_line

def is_white_background(hex_color):
    """
    文字列としての背景色 (例: "#FFFFFF") を受け取り、白であるか判定します。
    ※ この場合、明示的に "#FFFFFF" と設定されているセルのみが対象となります。
    """
    return hex_color == "#FFFFFF"

def process_review_file(spreadsheet, openai_key, creds_dict):
    """
    対象ファイル内のデータ行（ヘッダー除く）のうち、C列の userEnteredFormat で
    背景色が白 (#FFFFFF) と明示的に設定されている行のみを対象とし、
    前後文脈を含むプロンプトを作成して GPT にレビューを依頼し、
    返された結果でシートをバッチ更新します。
    """
    openai.api_key = openai_key
    worksheet = spreadsheet.worksheet("Task")
    rows = worksheet.get_all_values()  # 全行取得（ヘッダー含む）
    if len(rows) < 2:
        print("データ行がありません。")
        return
    # ヘッダー除くデータ行を抽出
    data = rows[1:]
    total_data_rows = len(data)
    spreadsheet_id = spreadsheet.id
    sheet_title = "Task"
    # 列C (ヘッダー除く) の書式情報（userEnteredFormat）を 1 回で取得
    format_dict = get_c_column_formatting(spreadsheet_id, sheet_title, creds_dict)
    
    eligible_indices = []  # data 内の 0-indexed 行番号
    for j in range(total_data_rows):
        # シート上の行番号 = j + 2
        row_num = j + 2
        hex_color = format_dict.get(row_num)  # 明示的に設定されていない場合は None
        if DEBUG:
            c_value = data[j][2] if len(data[j]) > 2 else ""
            print(f"Row {row_num}: Cセル値 = '{c_value}', 背景色 = {hex_color}")
        # 対象は、書式情報が存在し、かつ背景色が明示的に "#FFFFFF" である行のみ
        if hex_color != "#FFFFFF":
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
