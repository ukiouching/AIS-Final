# ask_gemini.py
import os
import sys
import google.generativeai as genai

def load_api_key(filepath="Gemini_API_key.txt"):
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return f.read().strip()
    except FileNotFoundError:
        print(f"❌ 找不到 API 金鑰檔案：{filepath}")
        sys.exit(1)

# 讀取 API 金鑰
api_key = load_api_key()
genai.configure(api_key=api_key)

# for m in genai.list_models():
#   if 'generateContent' in m.supported_generation_methods:
#     print(m.name)

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

关于合同款项总额（含不含税）的判断：如果合同中只写明金额的合计数，则视该合计数为合同款项总额（含税）。
关于是否有预开发票风险的判断：预开发票是指开票时间早于发货时间，若合同中写明开票时间早于发货时间，则判断为有预开发票风险；相反地，若开票时间晚于发货时间，则没有预开发票风险。如果合同中未明确提及两者的时间点，则返回“未明确提及”，并尽量说明相关信息，包括产品所有权转移时间。例如：“在乙方全额收到款项之前，产品的所有权归乙方所有。乙方收到全部款项后，产品的所有权归甲方所有。”该例中明确说明所有权转移时间在付款之后，因此没有预开发票风险。
关于是开口还是闭口合同的判断：若合同中明确说明款项回收的时间，则为开口合同。
"""

def ask_gemini(text: str):
    prompt = PROMPT_TEMPLATE.format(text=text)
    model = genai.GenerativeModel('gemini-1.5-pro-latest')
    response = model.generate_content(prompt)
    return response.text

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
    print(result)
