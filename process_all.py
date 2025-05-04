import os
import subprocess
import shutil
import time

PDF_DIR = "PDF"
TXT_DIR = "TXT"
RESULT_DIR = "RESULT"

os.makedirs(TXT_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

for filename in os.listdir(PDF_DIR):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(PDF_DIR, filename)
        base_name = os.path.splitext(filename)[0]
        txt_filename = base_name + ".txt"
        txt_path = os.path.join(TXT_DIR, txt_filename)
        result_path = os.path.join(RESULT_DIR, txt_filename)

        # Step 1: Extract text
        print(f"📄 Extracting: {filename}")
        try:
            result = subprocess.run(
                ["python", "read_doc.py", pdf_path],
                check=True,
                capture_output=True,
                text=True
            )
            output_text = result.stdout
            if "⚠️ PDF" in output_text:
                print(f"🧐 {filename} → 使用 OCR 擷取")
            else:
                print(f"✅ {filename} → 使用普通文字擷取")
        except subprocess.CalledProcessError as e:
            print(f"❌ 讀取 PDF 失敗：{filename}")
            continue

        raw_txt_path = os.path.splitext(pdf_path)[0] + ".txt"
        if os.path.exists(raw_txt_path):
            shutil.move(raw_txt_path, txt_path)
        else:
            print(f"⚠️ 找不到輸出的 TXT 檔案：{raw_txt_path}")
            continue

        # Step 2: Analyze with Gemini
        print(f"🤖 Analyzing: {txt_filename}")
        try:
            subprocess.run(["python", "ask_gemini.py", txt_path], check=True)
        except subprocess.CalledProcessError:
            print(f"❌ Gemini 分析失敗：{txt_filename}")
            continue

        # Step 3: Move result
        raw_result_path = os.path.splitext(txt_path)[0] + "_result.txt"
        if os.path.exists(raw_result_path):
            shutil.move(raw_result_path, result_path)
        else:
            print(f"⚠️ 找不到分析結果：{raw_result_path}")

        # Step 4: Sleep to respect quota limits
        print(f"⏳ 等待 40 秒避免配額限制...")
        time.sleep(40)

print("✅ 所有檔案處理完畢！")