# ask_gemini.py
import os
import sys
import google.generativeai as genai
import pandas as pd
import re

REPORT_PATH = "file_report.xlsx"

def load_api_key(filepath="Gemini_API_key.txt"):
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return f.read().strip()
    except FileNotFoundError:
        print(f"❌ 找不到 API 金鑰檔案：{filepath}")
        sys.exit(1)

api_key = load_api_key()
genai.configure(api_key=api_key)

PROMPT_TEMPLATE = """
请从以下文件中分析以下信息，并输出分析结果的纯文本，只需输出分析结果，且不需要格式化。

1. 需方：
2. 合同款项总额（不含税）、税额（税率13%）、合同款项总额（含税）：
3. 付款方式：
4. 需方账户：
5. 是否有预开发票风险：
6. 是开口还是闭口合同：

文件内容如下：
\"\"\"
{text}
\"\"\"

关于合同款项总额（含税或不含税）的判断：如果合同中仅写明金额的总和，则该总金额视为合同款项总额（含税）。
当税率不等于中国标准增值税（VAT）税率13%时，将提示警告，返回「未明确提及」。

付款方式：TT（电汇）、承兑汇票。例如，若合同写道：“甲方付款给乙方时，应以电汇支付”，则付款方式为电汇。
若合同未说明付款方式，则返回「未明确提及」。

关于是否存在预开发票风险的判断：预开发票是指开票时间早于发货时间。如果合同信息中包含开票时间早于发货时间的情形，则判断为「存在预开发票风险」。
相反，如果开票时间晚于发货时间，则判断为「不存在预开发票风险」。
例如：“合同中的产品全部发货后九十（90）天，甲方向乙方支付全部合同款项。乙方应在甲方付款的前一周，开出对应金额的发票。在乙方全额收到款项之前，产品的所有权归乙方所有。乙方收到全部款项后，产品的所有权归甲方所有。”此例中明确说明发货后才会付款及开票，因此判断为「不存在预开发票风险」。
如果无法判断开票与发货的时间先后，则返回「未明确提及」，并说明相关信息，例如产品所有权的转移时间。

关于是否为开口或闭口合同的判断：若合同中明确说明款项回收时间，则为闭口合同；否则为开口合同。
以下三种付款方式均属于闭口合同：
1. 乙方完成全部合同义务后的/_日内，甲方一次性支付合同总费用。
2. 乙方完成全部合同义务后的/_日内，甲方将合同费用总额的/_%，计/_元以/_方式支付乙方；剩余部分费用计/_元，甲方应于/_年/_月/_日前支付乙方。
3. 甲方收到合法有效的增值税专用发票并完成资金支付审批程序后，在45个工作日内通过银行转账方式向乙方支付业务外包费用。
例如：“合同中的产品全部发货后九十（90）天，甲方向乙方支付全部合同款项”，此例明确说明发货后90天内付款，因此具有明确的款项回收时间，为闭口合同。
"""

def ask_gemini(text: str):
    prompt = PROMPT_TEMPLATE.format(text=text)
    model = genai.GenerativeModel('gemini-1.5-flash')
    response = model.generate_content(prompt)
    return response.text

def is_suspicious(response_text: str) -> bool:
    lowered = response_text.lower()
    return (
        "预开发票风险" in response_text and "不存在预开发票风险" not in response_text
        or "未明确提及" in response_text
        or "未提及" in response_text 
        or "开口合同" in response_text
    )

def extract_suspicious_part(result_text: str) -> str:
    suspicious_fields = []

    # 依照換行分段
    for paragraph in result_text.splitlines():
        paragraph = paragraph.strip()

        if not paragraph:
            continue  # 跳過空行

        # 預開發票風險疑慮
        if "预开发票风险" in paragraph and "不存在预开发票风险" not in paragraph:
            suspicious_fields.append(paragraph)

        # 未提及類疑慮
        elif "未明确提及" in paragraph or "未提及" in paragraph:
            suspicious_fields.append(paragraph)

        # 開口合同疑慮：出現「開口合同」兩次以上才判定
        elif "开口合同" in paragraph:
            if paragraph.count("开口合同") >= 1:
                suspicious_fields.append(paragraph)

    if suspicious_fields:
        return "\n\n".join(suspicious_fields)
    return "未找到疑慮段落"

def log_suspicious(filename: str, analysis: str):
    summary = extract_suspicious_part(analysis)
    new_row = {
        "檔案名稱": filename,
        "疑慮摘要": summary
    }

    if os.path.exists(REPORT_PATH):
        df_existing = pd.read_excel(REPORT_PATH)
        df_updated = pd.concat([df_existing, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df_updated = pd.DataFrame([new_row])

    df_updated.to_excel(REPORT_PATH, index=False)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("用法：python ask_gemini.py <txt檔案路徑>")
        sys.exit(1)

    txt_path = sys.argv[1]
    with open(txt_path, 'r', encoding='utf-8') as f:
        content = f.read()

    result = ask_gemini(content)

    output_path = os.path.splitext(txt_path)[0] + "_result.txt"
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(result)

    # 判斷是否有疑慮，若有就寫入 Excel
    if is_suspicious(result):
        print(f"⚠️ 有疑慮，記錄到 {REPORT_PATH}")
        suspicious_summary = extract_suspicious_part(result)
        log_suspicious(os.path.basename(txt_path), suspicious_summary)


    print(result)
