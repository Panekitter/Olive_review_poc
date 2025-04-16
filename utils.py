from openai import OpenAI
from gspread_formatting import get_user_entered_format
from gspread_formatting.dataframe import Color

def is_white_background(cell_format):
    color = cell_format.backgroundColor
    if not color:
        return False
    r = color.get('red', 1)
    g = color.get('green', 1)
    b = color.get('blue', 1)
    return r == 1 and g == 1 and b == 1  # 完全に白（RGB 1,1,1）

def process_review_file(spreadsheet, openai_key):
    client = OpenAI(api_key=openai_key)
    worksheet = spreadsheet.worksheet("Task")
    rows = worksheet.get_all_values()

    for i in range(1, len(rows)):
        cell_format = get_user_entered_format(worksheet, f"C{i+1}")
        if not is_white_background(cell_format):
            continue

        prev, target, next_ = get_context(rows, i)
        prompt = f"""以下の英文は口語表現が含まれた文字起こしです。
中央の英文を翻訳レビューしてください。前後の文脈も参考にして翻訳精度を改善してください。

前文: {prev}
本文: {target}
後文: {next_}
日本語訳（初回）: {rows[i][1]}

出力:
1. 修正翻訳（1行）
2. エラー分類（誤訳、不自然、訳抜け、other）
3. otherの理由（あれば、英語で簡潔に）
"""

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5,
        )

        result = response.choices[0].message.content.splitlines()
        revised = result[0].strip()
        category = result[1].strip() if len(result) > 1 else ""
        explanation = result[2].strip() if len(result) > 2 else ""

        worksheet.update_cell(i+1, 3, revised)
        worksheet.update_cell(i+1, 4, category)
        if category.lower() == "other":
            worksheet.update_cell(i+1, 5, explanation)
