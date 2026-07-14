# AI PDF Scanner

PDF to Excel extractor with a Streamlit UI, optional OCR, OpenRouter-powered structured extraction, and built-in Excel export.

## What it does

- Upload one or more PDF files and extract structured data into Excel
- Upload CSV files and combine them into the same Excel output
- Use OCR when text extraction from a PDF is incomplete
- Use reference grids for supported documents such as Cala de Mar and Casa Colonial
- Expose a small Flask entrypoint for Vercel deployment

## Project files

- `pdf_to_excel.py` - main Streamlit app and extraction logic
- `app.py` - Flask entrypoint for Vercel and a simple health route
- `reference_grids/` - CSV reference data used for supported PDFs
- `requirements.txt` - Python dependencies

## Requirements

- Python 3.12 or newer
- An `OPENROUTER_API_KEY` if you want AI extraction for PDFs that are not covered by reference grids
- Tesseract and Poppler if you want OCR support to work locally

## Install

```bash
python -m pip install -r requirements.txt
```

## Environment variables

Create a `.env` file in the project root or set environment variables in your shell.

```env
OPENROUTER_API_KEY=your_key_here
OPENROUTER_MODEL=openai/gpt-4o-mini
OPENROUTER_SITE_URL=http://localhost:8501
OPENROUTER_APP_NAME=PDF to Excel Extractor
```

## Run locally

Start the Streamlit app:

```bash
python -m streamlit run pdf_to_excel.py
```

The app lets you:

1. Upload PDF and CSV files
2. Extract structured data
3. Download the generated Excel file

## Vercel entrypoint

`app.py` is a lightweight Flask app intended for deployment platforms such as Vercel. It serves:

- `/` - a basic HTML info page
- `/health` - a JSON health check

## Supported extraction flow

- If a PDF matches Cala de Mar or Casa Colonial reference content, the app loads the matching CSV grids
- Otherwise it sends extracted text to OpenRouter for structured JSON extraction
- It also tries to fill missing PDF text with OCR when the required packages are installed

## Notes

- The Streamlit app is the main local interface
- OCR may require extra system setup beyond Python packages
- Large or image-only PDFs may take longer to process
