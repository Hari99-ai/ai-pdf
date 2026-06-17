import io
import json
import os
from pathlib import Path

import requests
import streamlit as st

try:
    from pypdf import PdfReader
    PDF_IMPORT_ERROR = None
except ImportError as exc:
    PdfReader = None
    PDF_IMPORT_ERROR = exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_IMPORT_ERROR = None
except ImportError as exc:
    Workbook = None
    Alignment = Border = Font = PatternFill = Side = get_column_letter = None
    OPENPYXL_IMPORT_ERROR = exc


def load_env_value(name: str) -> str:
    value = os.getenv(name, "").strip()
    if value:
        return value

    env_path = Path(__file__).with_name(".env")
    if not env_path.exists():
        return ""

    for line in env_path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue

        key, raw_value = stripped.split("=", 1)
        if key.strip() == name:
            return raw_value.strip().strip('"').strip("'")

    return ""


OPENROUTER_API_KEY = load_env_value("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openai/gpt-4o-mini").strip() or "openai/gpt-4o-mini"
OPENROUTER_SITE_URL = os.getenv("OPENROUTER_SITE_URL", "http://localhost:8501").strip() or "http://localhost:8501"
OPENROUTER_APP_NAME = os.getenv("OPENROUTER_APP_NAME", "PDF to Excel Extractor").strip() or "PDF to Excel Extractor"


PROMPT = """Analyze this PDF and extract ALL data into a structured JSON object.

Rules:
1. Find every logical section: tables, fee lists, policies, contact info, dates, rates, etc.
2. Name each section with a short snake_case key, for example: room_rates, cancellation_policy
3. Tables -> list of dicts (one dict per row; keys = column headers in snake_case)
4. Bullet lists -> list of strings
5. Key-value pairs -> dict
6. Plain text -> string
7. Return ONLY valid JSON. No markdown fences, no explanation, no preamble.

Example structure:
{
  "agreement_info": {"hotel": "Cala de Mar", "period": "Jan 2026 - Mar 2027"},
  "room_rates": [
    {"dates": "Jan 03 - Mar 31 2026", "cliffside": "$320", "romance_deluxe": "$450"}
  ],
  "extra_charges": [
    {"item": "Extra Adult", "charge": "$100 USD"},
    {"item": "Pet Fee", "charge": "$85 USD/night"}
  ],
  "promotions": ["Early Booking: 5% off if booked 45 days ahead", "4th Night Free"],
  "cancellation_policy": {"deadline": "4 days before arrival", "penalty": "100% of total"},
  "contact_info": {"phone": "+52 755 555 1100", "email": "reservations@calademar.com"}
}"""


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed")

    reader = PdfReader(io.BytesIO(pdf_bytes))
    parts = []
    for page in reader.pages:
        text = page.extract_text() or ""
        text = text.strip()
        if text:
            parts.append(text)
    return "\n\n".join(parts)


def strip_code_fences(text: str) -> str:
    raw = text.strip()
    if raw.startswith("```"):
        raw = "\n".join(line for line in raw.splitlines() if not line.strip().startswith("```")).strip()
    return raw


def excel_safe_value(value):
    """Convert nested JSON values into something openpyxl can store."""
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, list):
        if not value:
            return ""
        if all(isinstance(item, (str, int, float, bool)) or item is None for item in value):
            return ", ".join("" if item is None else str(item) for item in value)
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def extract_data_with_openrouter(pdf_bytes: bytes) -> dict:
    pdf_text = extract_text_from_pdf(pdf_bytes)
    if not pdf_text.strip():
        raise ValueError("Could not extract text from the PDF. Try a text-based PDF or OCR first.")

    prompt = (
        f"{PROMPT}\n\n"
        "The PDF text is below. Preserve tables as structured data when possible.\n\n"
        f"PDF TEXT:\n{pdf_text}"
    )

    response = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "Content-Type": "application/json",
            "HTTP-Referer": OPENROUTER_SITE_URL,
            "X-OpenRouter-Title": OPENROUTER_APP_NAME,
        },
        json={
            "model": OPENROUTER_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1,
            "max_tokens": 8192,
        },
        timeout=180,
    )
    response.raise_for_status()

    payload = response.json()
    try:
        content = payload["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise ValueError(f"Unexpected OpenRouter response shape: {payload}") from exc

    return json.loads(strip_code_fences(content))


if OPENPYXL_IMPORT_ERROR is None:
    def make_fill(hex_color: str):
        return PatternFill("solid", start_color=hex_color)

    C_DARK_BLUE = "1F4E79"
    C_MID_BLUE = "2E75B6"
    C_LIGHT_BLUE = "D6E4F0"
    C_ACCENT = "E8F4FD"
    C_WHITE = "FFFFFF"
    C_TEXT_DARK = "1A1A2E"
    C_BORDER = "AEC6CF"

    def f_section(size=13):
        return Font(bold=True, color=C_WHITE, name="Calibri", size=size)

    def f_col_hdr():
        return Font(bold=True, color=C_WHITE, name="Calibri", size=10)

    def f_key():
        return Font(bold=True, color=C_TEXT_DARK, name="Calibri", size=10)

    def f_val():
        return Font(color=C_TEXT_DARK, name="Calibri", size=10)

    def f_bullet():
        return Font(color=C_TEXT_DARK, name="Calibri", size=10)

    _side = Side(style="thin", color=C_BORDER)
    BORDER = Border(left=_side, right=_side, top=_side, bottom=_side)
    AL_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    AL_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    AL_LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)
else:
    def make_fill(hex_color: str):
        return None

    def f_section(size=13):
        return None

    def f_col_hdr():
        return None

    def f_key():
        return None

    def f_val():
        return None

    def f_bullet():
        return None

    BORDER = AL_CENTER = AL_LEFT = AL_LEFT_TOP = None


def build_excel(extracted: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    row = 1
    for section_name, section_data in extracted.items():
        ws.cell(row=row, column=1, value=section_name.replace("_", " ").title())
        row += 1

        if isinstance(section_data, list) and section_data and isinstance(section_data[0], dict):
            headers = list(section_data[0].keys())
            for col_num, header in enumerate(headers, start=1):
                cell = ws.cell(row=row, column=col_num, value=header.replace("_", " ").title())
                cell.font = f_col_hdr()
                cell.fill = make_fill(C_MID_BLUE)
                cell.alignment = AL_CENTER
                cell.border = BORDER
            row += 1

            for i, record in enumerate(section_data):
                bg = make_fill(C_LIGHT_BLUE if i % 2 == 0 else C_WHITE)
                for col_num, header in enumerate(headers, start=1):
                    cell = ws.cell(row=row, column=col_num, value=excel_safe_value(record.get(header, "")))
                    cell.font = f_val()
                    cell.fill = bg
                    cell.alignment = AL_LEFT
                    cell.border = BORDER
                row += 1

        elif isinstance(section_data, dict):
            headers = list(section_data.keys())
            for col_num, header in enumerate(headers, start=1):
                cell = ws.cell(row=row, column=col_num, value=header.replace("_", " ").title())
                cell.font = f_col_hdr()
                cell.fill = make_fill(C_MID_BLUE)
                cell.alignment = AL_CENTER
                cell.border = BORDER
            row += 1

            for col_num, header in enumerate(headers, start=1):
                cell = ws.cell(row=row, column=col_num, value=excel_safe_value(section_data[header]))
                cell.font = f_val()
                cell.fill = make_fill(C_WHITE)
                cell.alignment = AL_LEFT_TOP
                cell.border = BORDER
            row += 1

        elif isinstance(section_data, list):
            for item in section_data:
                cell = ws.cell(row=row, column=1, value=excel_safe_value(item))
                cell.font = f_bullet()
                cell.fill = make_fill(C_WHITE)
                cell.alignment = AL_LEFT_TOP
                cell.border = BORDER
                row += 1

        else:
            cell = ws.cell(row=row, column=1, value=excel_safe_value(section_data))
            cell.font = f_val()
            cell.alignment = AL_LEFT_TOP
            row += 1

        row += 1

    for column in ws.columns:
        max_length = 0
        try:
            column_letter = column[0].column_letter
        except Exception:
            continue
        for cell in column:
            try:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[column_letter].width = min(max_length + 2, 50)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


st.set_page_config(page_title="PDF to Excel Extractor", page_icon="PDF", layout="centered")
st.title("PDF to Excel Extractor")
st.markdown("Upload a PDF and extract structured data into a clean Excel file.")

with st.sidebar:
    st.header("How it works")
    st.markdown("1. Upload a PDF\n2. Click Extract\n3. Download the Excel file")
    st.markdown("---")
    st.caption(f"Model: `{OPENROUTER_MODEL}`")

uploaded_file = st.file_uploader("Upload your PDF", type=["pdf"])

if not OPENROUTER_API_KEY:
    st.warning("Add OPENROUTER_API_KEY to your .env file before extracting a PDF.")

if PDF_IMPORT_ERROR:
    st.error("Missing dependency: pypdf. Install project dependencies with `python -m pip install -r requirements.txt`.")
    st.stop()

if OPENPYXL_IMPORT_ERROR:
    st.error("Missing dependency: openpyxl. Install project dependencies with `python -m pip install -r requirements.txt`.")
    st.stop()

if uploaded_file:
    st.info(f"{uploaded_file.name} - {uploaded_file.size / 1024:.1f} KB")

    if st.button("Extract and Generate Excel", type="primary", use_container_width=True):
        if not OPENROUTER_API_KEY:
            st.error("OPENROUTER_API_KEY is missing. Add it to your .env file, then reload the page.")
            st.stop()

        pdf_bytes = uploaded_file.read()

        with st.spinner("OpenRouter is reading your PDF..."):
            try:
                extracted = extract_data_with_openrouter(pdf_bytes)
            except json.JSONDecodeError as e:
                st.error(f"Could not parse AI response as JSON: {e}")
                st.stop()
            except requests.HTTPError as e:
                st.error(f"OpenRouter API error: {e}")
                st.stop()
            except Exception as e:
                st.error(f"OpenRouter error: {e}")
                st.stop()

        st.success(f"Extracted {len(extracted)} sections from the PDF.")

        with st.expander("Preview extracted data (JSON)", expanded=False):
            st.json(extracted)

        with st.spinner("Building Excel file..."):
            excel_bytes = build_excel(extracted)

        filename = uploaded_file.name.replace(".pdf", "_extracted.xlsx")
        st.download_button(
            label="Download Excel File",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
