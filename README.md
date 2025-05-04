## 建立虛擬環境

1. 開啟一個 virtual environment (虛擬環境)，用於管理安裝的套件

```
python3 -m venv .venv
```

2. 進入虛擬環境

```
source .venv/bin/activate
```

## 安裝套件

- 安裝 docx 與 PyPDF2 (用於讀取 pdf 與 word)

```
python3 -m pip install python-docx PyPDF2
```

- 安裝 Python SDK

```
python3 -m pip install google-generativeai
```

- 安裝 pytesseract

```
python3 -m pip install pytesseract pdf2image Pillow
```

- 安裝 Tesseract OCR

```
brew install tesseract
```

## 執行程式

- 掃描 pdf / doc

```
python3 read_doc.py <要掃描的檔案路徑>
```

- 交給 Gemini 分析

```
python3 ask_gemini.py <掃描好的txt檔案路徑>
```

- 一次處理 /PDF 中的所有 pdf 檔案

```
python3 process_all.py
```
