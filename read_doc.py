import os
from PyPDF2 import PdfReader
from docx import Document
from pdf2image import convert_from_path
import pytesseract

def read_text_from_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.pdf':
        text = read_pdf(file_path)
        if not text.strip():
            print(f"⚠️ PDF {file_path} 沒有文字層，使用 OCR 處理...")
            text = ocr_pdf(file_path)
        return text
    elif ext == '.docx':
        return read_docx(file_path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

def read_pdf(file_path):
    reader = PdfReader(file_path)
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text
    return text

def ocr_pdf(file_path):
    images = convert_from_path(file_path)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang='chi_sim+eng')  # 中文+英文 OCR
    return text

def read_docx(file_path):
    doc = Document(file_path)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

if __name__ == '__main__':
    import sys
    if len(sys.argv) != 2:
        print("Usage: python read_doc.py <file_path>")
        sys.exit(1)

    file_path = sys.argv[1]
    try:
        content = read_text_from_file(file_path)
        output_path = os.path.splitext(file_path)[0] + '.txt'
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print("=== File Content ===")
        print(content)
    except Exception as e:
        print(f"Error: {e}")