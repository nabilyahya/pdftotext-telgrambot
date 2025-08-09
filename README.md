# 📄 Telegram OCR & PDF Converter Bot

## 📌 Overview
This project is a **Telegram Bot** that processes PDF files (scanned or text-based) and converts them into different output formats with **OCR (Optical Character Recognition)** support for **Arabic, Turkish, and English**.  
It can generate:
- Searchable PDF
- Plain text (TXT) with correct **RTL** handling for Arabic
- DOCX (text only)
- DOCX (text + inline images)

The bot is optimized for high accuracy in Arabic text recognition and preserves layout as much as possible.

---

## ✨ Features
- **Multi-language OCR**: Arabic (`ara`), Turkish (`tur`), English (`eng`)
- **Image Preprocessing**:
  - Deskew (automatic rotation correction)
  - Noise reduction & adaptive thresholding
- **Output Formats**:
  - 📄 **Searchable PDF** (keeps images, adds selectable text layer)
  - 📜 **TXT** (UTF-8 with RTL support for Arabic)
  - 📝 **DOCX** (text only)
  - 📝 **DOCX with inline images**
- **Arabic Text Processing**:
  - Normalization (remove Tatweel, unify character forms)
  - Visual order correction for OCR outputs
  - RTL embedding in TXT for correct display
- **Performance**:
  - Multi-threaded OCR with `ThreadPoolExecutor`
  - DPI control for better OCR accuracy

---

## 🛠 Tech Stack
- **Python 3.10+**
- [python-telegram-bot](https://python-telegram-bot.org/) – Telegram Bot API
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) – Optical Character Recognition
- [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/) – PDF processing
- [OpenCV](https://opencv.org/) – Image preprocessing
- [Pillow (PIL)](https://python-pillow.org/) – Image handling
- [python-docx](https://python-docx.readthedocs.io/) – DOCX file generation
- [ocrmypdf](https://ocrmypdf.readthedocs.io/) – Searchable PDF generation (optional)

---

## 📂 Project Structure
