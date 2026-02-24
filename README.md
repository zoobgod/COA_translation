# COA Translation — Pharmaceutical Certificate of Analysis Translator

Streamlit application that translates pharmaceutical Certificate of Analysis (COA) documents from English to Russian using OpenAI, with a specialized pharmaceutical and medicinal glossary.

## Features

- **PDF Upload** — Upload COA documents in PDF format
- **Multi-method text extraction** — Uses pdfplumber (primary), PyMuPDF (fallback), and OCR via pytesseract (for scanned documents)
- **AI Translation with pharma glossary** — Translates via OpenAI models with a 200+ term pharmaceutical glossary enforcing standard Russian pharmaceutical terminology
- **Formatted Word output** — Generates a professionally formatted `.docx` file with the translated content
- **Download** — One-click download of the translated document

## Setup

### Prerequisites

- Python 3.10+
- OpenAI API key
- (Optional) Tesseract OCR installed on the system for scanned PDF support

### Installation

```bash
pip install -r requirements.txt
```

For OCR support (scanned PDFs), install Tesseract:

```bash
# Ubuntu/Debian
sudo apt-get install tesseract-ocr

# macOS
brew install tesseract
```

### Generate the Word template (optional)

A pre-prepared Word template can be generated for docxtpl-based rendering:

```bash
python -m modules.create_template
```

If the template is not present, the app falls back to generating documents from scratch using python-docx.

## Running the App

```bash
streamlit run app.py
```

Then open the URL shown in the terminal (typically `http://localhost:8501`).

## Usage

1. Enter your OpenAI API key in the sidebar
2. Select a translation model (gpt-4o recommended for best quality)
3. Upload a COA PDF file
4. Review the extracted text
5. Click **Translate to Russian**
6. Download the translated `.docx` file

## Project Structure

```
├── app.py                    # Main Streamlit application
├── requirements.txt          # Python dependencies
├── .streamlit/
│   └── config.toml           # Streamlit theme and server config
├── modules/
│   ├── __init__.py
│   ├── glossary.py           # Pharmaceutical EN→RU glossary (200+ terms)
│   ├── pdf_extractor.py      # PDF text extraction (pdfplumber, PyMuPDF, OCR)
│   ├── translator.py         # OpenAI translation with pharma context
│   ├── doc_generator.py      # Word document generation
│   └── create_template.py    # Script to generate the docxtpl template
└── templates/
    └── coa_template.docx     # Generated Word template (created by create_template.py)
```

## Glossary

The pharmaceutical glossary (`modules/glossary.py`) includes standard translations for:

- Document headers and metadata fields
- Physical and chemical test parameters
- Analytical methods (HPLC, GC, MS, etc.)
- Microbiological tests
- Dosage forms
- Units and measurements
- Regulatory references (USP, EP, BP, etc.)
- Common COA result phrases

The glossary terms are injected into the OpenAI system prompt to ensure consistent, accurate pharmaceutical translations.
