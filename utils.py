import openai

def get_context(rows, index):
    prev_line = rows[index - 1][0] if index > 1 else ""
    target_line = rows[index][0]
    next_line = rows[index + 1][0] if index + 1 < len(rows) else ""
    return prev_line, target_line, next_line

def process_review_file(spreadsheet, openai_key):
    openai.api_key = openai_key
    worksheet = spreadsheet.worksheet("Task")
    rows = worksheet.get_all_values()

    for i in range(1, len(rows)):
        ja_translation = rows[i][2]  # C列
        if ja_translation.strip():  # C列が空白でないならスキップ
            continue

        prev, target, next_ = get_context(rows, i)
        prompt = f"""
以下の英文は口語表現が含まれた文字起こしです。
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
        res = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5
        )
        result = res['choices'][0]['message']['content'].splitlines()
        revised = result[0].strip()
        category = result[1].strip() if len(result) > 1 else ""
        explanation = result[2].strip() if len(result) > 2 else ""

        worksheet.update_cell(i+1, 3, revised)
        worksheet.update_cell(i+1, 4, category)
        if category.lower() == "other":
            worksheet.update_cell(i+1, 5, explanation)
