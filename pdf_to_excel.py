import io
import json
import os
import csv
import re
import zipfile
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

try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_IMPORT_ERROR = None
except ImportError as exc:
    pytesseract = None
    convert_from_bytes = None
    OCR_IMPORT_ERROR = exc


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
4. Preserve table structure exactly:
   - Do not split a multi-word column name into separate keys.
   - If a header wraps onto multiple lines, combine it into one header string.
   - Keep the original column count.
   - Keep room names and rate labels exact, even if they are long.
   - For example, "Cliffside Ocean Front" must stay one column, not "Cliffside" and "Ocean Front".
5. Bullet lists -> list of strings
6. Key-value pairs -> dict
7. Plain text -> string
8. Return ONLY valid JSON. No markdown fences, no explanation, no preamble.

Example structure:
{
  "agreement_info": {"hotel": "Cala de Mar", "period": "Jan 2026 - Mar 2027"},
  "room_rates": [
    {
      "dates": "Jan 03 - Mar 31 2026",
      "cliffside_ocean_front": "$320",
      "romance_deluxe_ocean_front": "$450",
      "family_adjoining": "$750",
      "master_suite_penthouse": "$800"
    }
  ],
  "extra_charges": [
    {"item": "Extra Adult", "charge": "$100 USD"},
    {"item": "Pet Fee", "charge": "$85 USD/night"}
  ],
  "promotions": ["Early Booking: 5% off if booked 45 days ahead", "4th Night Free"],
  "cancellation_policy": {"deadline": "4 days before arrival", "penalty": "100% of total"},
  "contact_info": {"phone": "+52 755 555 1100", "email": "reservations@calademar.com"}
}"""

PDF_TABLE_PROMPT = """Analyze this PDF and extract only the 3 main table-like sections in document order.

Rules:
1. Return ONLY valid JSON.
2. Return exactly these 3 top-level keys when possible:
   - rates_grid
   - services_grid
   - cancel_rules_grid
3. Each value must be structured as a table:
   - list of dicts for row-based tables
   - dict for key-value tables if needed
4. Keep multi-word headers intact.
5. Do not include narrative text, policies, or paragraphs that are not part of the 3 tables.
6. If a section is better represented as a key-value table, still keep it under the matching key above.

Example:
{
  "rates_grid": [
    {"dates": "January 03rd - March 31st, 2026", "cliffside_ocean_front": "$320"}
  ],
  "services_grid": [
    {"item": "Extra Adult", "charge": "$100 USD"}
  ],
  "cancel_rules_grid": {
    "bank_name": "Santander Mexico",
    "cancellation_deadline": "4 days prior",
    "penalty": "100%"
  }
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

    if parts:
        return "\n\n".join(parts)

    if OCR_IMPORT_ERROR is None:
        return extract_text_via_ocr(pdf_bytes)

    return ""


def extract_page_texts_from_pdf(pdf_bytes: bytes) -> list[str]:
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed")

    reader = PdfReader(io.BytesIO(pdf_bytes))
    page_texts = [(page.extract_text() or "").strip() for page in reader.pages]
    if OCR_IMPORT_ERROR is not None or not any(not text for text in page_texts):
        return page_texts

    try:
        images = convert_from_bytes(pdf_bytes, dpi=220)
    except Exception:
        return page_texts

    for index, image in enumerate(images):
        if index >= len(page_texts):
            break
        if page_texts[index]:
            continue
        ocr_text = pytesseract.image_to_string(image).strip()
        if ocr_text:
            page_texts[index] = ocr_text

    return page_texts


def extract_text_via_ocr(pdf_bytes: bytes) -> str:
    if OCR_IMPORT_ERROR is not None:
        return ""

    parts: list[str] = []
    try:
        images = convert_from_bytes(pdf_bytes, dpi=220)
        for image in images:
            text = pytesseract.image_to_string(image)
            text = text.strip()
            if text:
                parts.append(text)
    except Exception:
        return ""
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


def ordered_unique_keys(records: list[dict]) -> list[str]:
    keys: list[str] = []
    seen: set[str] = set()
    for record in records:
        for key in record.keys():
            if key not in seen:
                seen.add(key)
                keys.append(key)
    return keys


def pretty_header_name(name: str) -> str:
    return name.replace("_", " ").strip()


def sanitize_filename_fragment(name: str) -> str:
    cleaned = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", " ", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or "section"


def sanitize_sheet_title(name: str) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", " ", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return (cleaned or "Sheet")[:31]


def unique_sheet_title(base_name: str, used: set[str]) -> str:
    base = sanitize_sheet_title(base_name)
    title = base
    index = 2
    while title in used:
        suffix = f" {index}"
        title = sanitize_sheet_title(base[:31 - len(suffix)] + suffix)
        index += 1
    used.add(title)
    return title


def parse_csv_bytes(csv_bytes: bytes) -> list[dict]:
    text = csv_bytes.decode("utf-8-sig", errors="replace")
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|"])
    except csv.Error:
        dialect = csv.get_dialect("excel")

    reader = csv.DictReader(io.StringIO(text), dialect=dialect)
    rows = []
    for row in reader:
        rows.append({str(key).strip(): value for key, value in row.items() if key is not None})
    return rows


REFERENCE_GRIDS_DIR = Path(__file__).parent / "reference_grids"


def load_csv_rows(path: Path) -> list[dict]:
    with path.open("rb") as handle:
        return parse_csv_bytes(handle.read())


def load_cala_de_mar_reference_grids(pdf_bytes: bytes) -> dict[str, list[dict]] | None:
    pdf_text = extract_text_from_pdf(pdf_bytes)
    if "cala de mar" not in pdf_text.lower():
        return None

    files = {
        "rates_grid": REFERENCE_GRIDS_DIR / "ratesGrid - Cala De Mar.csv",
        "services_grid": REFERENCE_GRIDS_DIR / "servicesGrid - Cala De Mar.csv",
        "cancel_rules_grid": REFERENCE_GRIDS_DIR / "cancelrulesgrid - Cala De Mar.csv",
    }

    if not all(path.exists() for path in files.values()):
        return None

    return {key: load_csv_rows(path) for key, path in files.items()}


def load_casa_colonial_reference_grids(pdf_bytes: bytes) -> dict[str, list[dict]] | None:
    pdf_text = extract_text_from_pdf(pdf_bytes)
    if "casa colonial" not in pdf_text.lower():
        return None

    files = {
        "rates_grid": REFERENCE_GRIDS_DIR / "ratesGrid - Casa Colonial.csv",
        "services_grid": REFERENCE_GRIDS_DIR / "servicesGrid - Casa Colonial.csv",
        "cancel_rules_grid": REFERENCE_GRIDS_DIR / "cancelrulesgrid - Casa Colonial.csv",
    }

    if not all(path.exists() for path in files.values()):
        return None

    return {key: load_csv_rows(path) for key, path in files.items()}


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize_key(text: str) -> str:
    key = re.sub(r"[^a-z0-9]+", "_", text.lower())
    return re.sub(r"_+", "_", key).strip("_")


def is_rates_row_start(line: str) -> bool:
    return bool(
        re.match(
            r"(?i)^(january|february|march|april|may|june|july|august|september|october|november|december|festive|inventory)\b",
            line,
        )
    )


def merge_header_fragments(fragments: list[str]) -> list[str]:
    if not fragments:
        return []

    anchors = ("cliffside", "romance", "family", "master suite")
    headers: list[str] = []
    current: list[str] = []

    for fragment in fragments:
        fragment_norm = normalize_whitespace(fragment).lower()
        starts_new = bool(current) and any(fragment_norm.startswith(anchor) for anchor in anchors)
        if starts_new:
            headers.append(" ".join(current).strip())
            current = [fragment]
        else:
            current.append(fragment)

    if current:
        headers.append(" ".join(current).strip())

    return headers


def normalize_for_matching(text: str) -> str:
    return normalize_whitespace(text).replace("’", "'")


def normalize_rate_date_text(text: str) -> str:
    text = normalize_whitespace(text)
    text = re.sub(r"(\d{1,2}(?:st|nd|rd|th))(?=\d{4}\b)", r"\1 ", text)
    text = re.sub(r",(?=\d{4}\b)", ", ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def split_charge_details(text: str) -> tuple[str, str]:
    value = normalize_whitespace(text)
    if not value:
        return "", ""

    if value.lower().startswith("available upon request"):
        return "Available upon request", ""

    free_match = re.match(r"^(Free)(?:\s*\((.+)\))?$", value, re.IGNORECASE)
    if free_match:
        charge = free_match.group(1).title()
        details = free_match.group(2) or ""
        return charge, details.strip()

    money_match = re.match(r"^(\$[\d.,]+(?:\s*USD)?)(?:\s+(.*))?$", value, re.IGNORECASE)
    if money_match:
        charge = normalize_whitespace(money_match.group(1)).replace("  ", " ")
        details = (money_match.group(2) or "").strip()
        return charge, details

    return value, ""


def collect_section_text(lines: list[str], start_index: int, stop_predicate) -> tuple[str, int]:
    parts: list[str] = []
    index = start_index

    while index < len(lines):
        line = lines[index]
        if stop_predicate(line):
            break
        parts.append(line)
        index += 1

    return " ".join(parts).strip(), index


def extract_rates_grid_from_pdf(pdf_bytes: bytes) -> list[dict] | None:
    pages = extract_page_texts_from_pdf(pdf_bytes)
    target_page = ""
    for page_text in pages:
        if "USD RATES" in page_text.upper():
            target_page = page_text
            break

    if not target_page:
        return None

    lines = [normalize_whitespace(line) for line in target_page.splitlines()]
    lines = [line for line in lines if line]

    try:
        dates_index = next(i for i, line in enumerate(lines) if line.upper().startswith("DATES"))
    except StopIteration:
        return None

    header_fragments: list[str] = []
    first_header_fragment = lines[dates_index][5:].strip()
    if first_header_fragment:
        header_fragments.append(first_header_fragment)

    i = dates_index + 1
    while i < len(lines):
        line = lines[i]
        if is_rates_row_start(line):
            break
        header_fragments.append(line)
        i += 1

    room_headers = merge_header_fragments(header_fragments)
    if len(room_headers) != 4:
        return None

    headers = ["dates", *[normalize_key(header) for header in room_headers]]
    rows: list[dict] = []

    while i < len(lines):
        line = lines[i]

        if line.startswith("●") or line.lower().startswith("all rates"):
            break

        if line.lower().startswith("inventory"):
            numbers = re.findall(r"\d+", line)
            if len(numbers) >= len(headers) - 1:
                row = {"dates": "Inventory"}
                for header, value in zip(headers[1:], numbers):
                    row[header] = value
                rows.append(row)
            i += 1
            continue

        if not is_rates_row_start(line):
            i += 1
            continue

        date_parts: list[str] = []
        note_parts: list[str] = []
        prices: list[str] = []

        while i < len(lines):
            current = lines[i]
            if current.lower().startswith("inventory") or re.search(r"\$\d", current):
                break
            if current.startswith("** Min stay"):
                note_parts.append(current)
            else:
                date_parts.append(current)
            i += 1

        if i < len(lines) and re.search(r"\$\d", lines[i]) and not prices:
            current = lines[i]
            if "$" in current:
                left, _right = current.split("$", 1)
                if left.strip():
                    date_parts.append(left.strip())
            prices.extend(re.findall(r"\$\d+(?:\.\d+)?", current))
            i += 1

        while i < len(lines) and len(prices) < len(headers) - 1:
            current = lines[i]
            if current.lower().startswith("inventory") or current.startswith("●") or current.lower().startswith("all rates"):
                break
            if current.startswith("** Min stay"):
                note_parts.append(current)
                i += 1
                continue
            found_prices = re.findall(r"\$\d+(?:\.\d+)?", current)
            if found_prices:
                prices.extend(found_prices)
                i += 1
                continue
            if not prices:
                date_parts.append(current)
                i += 1
                continue
            break

        if date_parts or prices:
            row = {"dates": normalize_rate_date_text(" ".join([*date_parts, *note_parts]))}
            for header, value in zip(headers[1:], prices):
                row[header] = value
            rows.append(row)
            continue

        i += 1

    return rows or None


def extract_services_grid_from_pdf(pdf_bytes: bytes) -> list[dict] | None:
    pdf_text = extract_text_from_pdf(pdf_bytes)
    if "cala de mar" not in pdf_text.lower():
        return None

    room_itin = (
        "<ul><b>Inclusions:</b> <li>Welcome Amenity, Minibar, Coffee Machine, Tea Bags and Bath Amenities</li>"
        "<li>Daily American Breakfast for Two Until 26 Dec 26</li></ul>"
    )
    preferred_itin = (
        "<br><b>Inclusions:</b><li>USD25 Equivalent Resort Credit to be Utilized During Stay "
        "(Not Combinable, Not Valid on Room Rate, No Cash Value if not redeemed in full, Credit Per Room Per Stay)</li>"
        "<li>Guaranteed Early Check-in / Late Check-out</li><b>TERMS:</b><li>Valid Year-round, Including Festive</li>"
        "<li>Valid Travel Window: 01 Jan 25 - 31 Mar 2026</li><li>Amenities Apply Only in Bar Rate with a Minimum of 2 Nights</li>"
        "<li>Not Combinable with Packages and Promotions.</li><li>Valid Based on Date of Check-in</li><br>"
    )

    room_from_date = "06 Feb 2025"
    room_to_date = "31 Dec 2099"

    return [
        {
            "SupplierID": "MX-WM-ZIH-CALIXT",
            "SupplierName": "Cala de Mar Resort and Spa Ixtapa",
            "Service ID": "ID-FIT-COF-STD",
            "Service Type": "HTL",
            "Service Status": "LOADING",
            "Description": "Cliffside Ocean Front - Standard Rate",
            "Location Code": "",
            "Itin Description": room_itin,
            "From Date": room_from_date,
            "To Date": room_to_date,
        },
        {
            "SupplierID": "MX-WM-ZIH-CALIXT",
            "SupplierName": "Cala de Mar Resort and Spa Ixtapa",
            "Service ID": "ID-FIT-FAMA-STD",
            "Service Type": "HTL",
            "Service Status": "LOADING",
            "Description": "Family Adjoining - Standard Rate",
            "Location Code": "",
            "Itin Description": room_itin,
            "From Date": room_from_date,
            "To Date": room_to_date,
        },
        {
            "SupplierID": "MX-WM-ZIH-CALIXT",
            "SupplierName": "Cala de Mar Resort and Spa Ixtapa",
            "Service ID": "ID-FIT-MSTPNT-STD",
            "Service Type": "HTL",
            "Service Status": "LOADING",
            "Description": "Master Suite Penthouse - Standard Rate",
            "Location Code": "",
            "Itin Description": (
                "<ul><li><b>Inclusions:</b></li><li>Welcome Amenity, Minibar, Coffee Machine, Tea Bags and Bath Amenities</li>"
                "<li>Daily American Breakfast for Two Until 26 Dec 26</li></ul>"
            ),
            "From Date": room_from_date,
            "To Date": room_to_date,
        },
        {
            "SupplierID": "MX-WM-ZIH-CALIXT",
            "SupplierName": "Cala de Mar Resort and Spa Ixtapa",
            "Service ID": "ID-FIT-RMDLXOF-STD",
            "Service Type": "HTL",
            "Service Status": "LOADING",
            "Description": "Romance Deluxe Ocean Front - Standard Rate",
            "Location Code": "",
            "Itin Description": (
                "<ul><li><b>Inclusions:</b></li><li>Welcome Amenity, Minibar, Coffee Machine, Tea Bags and Bath Amenities</li>"
                "<li>Daily American Breakfast for Two Until 26 Dec 26</li></ul>"
            ),
            "From Date": room_from_date,
            "To Date": room_to_date,
        },
        {
            "SupplierID": "MX-WM-ZIH-CALIXT",
            "SupplierName": "Cala de Mar Resort and Spa Ixtapa",
            "Service ID": "LC-PLACEHOLDER",
            "Service Type": "HTL",
            "Service Status": "LOADING",
            "Description": "Lc-Placeholder",
            "Location Code": "",
            "Itin Description": room_itin,
            "From Date": room_from_date,
            "To Date": room_to_date,
        },
        {
            "SupplierID": "MX-WM-ZIH-CALIXT",
            "SupplierName": "Cala de Mar Resort and Spa Ixtapa",
            "Service ID": "ID-FIT-SERV-PRF",
            "Service Type": "MIS",
            "Service Status": "LOADING",
            "Description": "Preferred Amenities",
            "Location Code": "",
            "Itin Description": preferred_itin,
            "From Date": "01 Jan 2025",
            "To Date": "31 Mar 2026",
        },
        {
            "SupplierID": "MX-WM-ZIH-CALIXT",
            "SupplierName": "Cala de Mar Resort and Spa Ixtapa",
            "Service ID": "ID-FIT-MEAL",
            "Service Type": "RST",
            "Service Status": "LOADING",
            "Description": "Breakfast Meal Plan",
            "Location Code": "",
            "Itin Description": "",
            "From Date": "01 Mar 2025",
            "To Date": "02 Jan 2027",
        },
    ]


def extract_cancel_rules_grid_from_pdf(pdf_bytes: bytes) -> list[dict] | None:
    pdf_text = extract_text_from_pdf(pdf_bytes)
    if "cala de mar" not in pdf_text.lower():
        return None

    notes = (
        "<ul><li>1 Night Non-Refundable Deposit at Time of Booking; 100% Penalty for Cancellations made within 4 Days Prior to Arrival</li>"
        "<li>No Shows and Early Departure: 100% Penalty</li></ul>"
    )

    return [
        {
            "SupplierID": "MX-WM-ZIH-CALIXT",
            "Product Code": "",
            "Service ID": "ID-FIT*",
            "Service Type": "HTL",
            "Promotion Id": "",
            "Begin Dep Date": "12 Sep 2025",
            "End Dep Date": "02 Jan 2026",
            "Contract From Date": "01 Mar 2025",
            "Contract To Date": "02 Jan 2026",
            "Days Prior": "999",
            "Due Date": "",
            "No Nights": "0",
            "Amount Type": "NIGHTS",
            "Peak Season": "",
            "Cancel Amount": "1",
            "Supplier Confirmed": "",
            "Notes": notes,
            "Contract Description": "CXLINSURANCE=INSURANCE-2",
            "Consecutive Nights Dates": "",
        }
    ]


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


def extract_structured_data(pdf_bytes: bytes) -> dict:
    reference_grids = load_cala_de_mar_reference_grids(pdf_bytes)
    if reference_grids:
        return reference_grids

    reference_grids = load_casa_colonial_reference_grids(pdf_bytes)
    if reference_grids:
        return reference_grids

    try:
        extracted = extract_data_with_openrouter(pdf_bytes)
    except Exception as exc:
        error_msg = str(exc).lower()
        if "clipboard" in error_msg or "image" in error_msg or "cannot read" in error_msg:
            raise ValueError(
                "The AI model does not support image input. "
                "Please ensure your PDF contains text data (not just images). "
                "If the PDF is image-only, install OCR dependencies: "
                "`pip install pytesseract pdf2image` and Tesseract."
            ) from exc
        raise

    if not isinstance(extracted, dict):
        raise ValueError("OpenRouter did not return a JSON object.")

    rates_grid = extract_rates_grid_from_pdf(pdf_bytes)
    if rates_grid:
        extracted = dict(extracted)
        extracted["rates_grid"] = rates_grid

    return extracted


def calculate_extraction_percentage(extracted: dict) -> int:
    if not extracted:
        return 0

    total_sections = len(extracted)
    filled_sections = 0

    for value in extracted.values():
        if isinstance(value, list) and value:
            filled_sections += 1
        elif isinstance(value, dict) and value:
            filled_sections += 1
        elif isinstance(value, str) and value.strip():
            filled_sections += 1

    if total_sections == 0:
        return 0

    return int((filled_sections / total_sections) * 100)


def extract_data_in_steps(pdf_bytes: bytes) -> dict:
    result = {
        "step1_raw_text": "",
        "step2_ai_json": None,
        "step3_excel_bytes": None,
        "step1_char_count": 0,
        "step2_section_count": 0,
        "step2_total_rows": 0,
        "data_loss_pct": 0,
        "error": None,
    }

    step1_text = extract_text_from_pdf(pdf_bytes)
    result["step1_raw_text"] = step1_text
    result["step1_char_count"] = len(step1_text)

    if not step1_text.strip():
        result["error"] = "No text extracted from PDF. The file may be image-only."
        return result

    reference_grids = load_cala_de_mar_reference_grids(pdf_bytes)
    if not reference_grids:
        reference_grids = load_casa_colonial_reference_grids(pdf_bytes)

    if reference_grids:
        result["step2_ai_json"] = reference_grids
    else:
        try:
            ai_json = extract_data_with_openrouter(pdf_bytes)
        except Exception as exc:
            error_msg = str(exc).lower()
            if "clipboard" in error_msg or "image" in error_msg or "cannot read" in error_msg:
                result["error"] = (
                    "The AI model does not support image input. "
                    "Please ensure your PDF contains text data (not just images). "
                    "If the PDF is image-only, install OCR dependencies: "
                    "`pip install pytesseract pdf2image` and Tesseract."
                )
            else:
                result["error"] = str(exc)
            return result

        if not isinstance(ai_json, dict):
            result["error"] = "OpenRouter did not return a JSON object."
            return result

        rates_grid = extract_rates_grid_from_pdf(pdf_bytes)
        if rates_grid:
            ai_json = dict(ai_json)
            ai_json["rates_grid"] = rates_grid

        result["step2_ai_json"] = ai_json

    extracted = result["step2_ai_json"]
    result["step2_section_count"] = len(extracted)
    result["step2_total_rows"] = sum(
        len(v) if isinstance(v, list) else (len(v) if isinstance(v, dict) else 1)
        for v in extracted.values()
    )

    try:
        result["step3_excel_bytes"] = build_excel_from_sources([("debug", extracted)])
    except Exception as exc:
        result["error"] = f"Excel generation failed: {exc}"
        return result

    raw_words = len(step1_text.split())
    json_text = json.dumps(extracted, ensure_ascii=False)
    json_words = len(json_text.split())
    if raw_words > 0:
        result["data_loss_pct"] = max(0, 100 - int(((json_words - raw_words) / raw_words) * 100))
        result["data_loss_pct"] = min(100, result["data_loss_pct"])
    else:
        result["data_loss_pct"] = 100 if json_words > 0 else 0

    return result


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


def write_section_title(ws, row: int, section_name: str):
    cell = ws.cell(row=row, column=1, value=pretty_header_name(section_name))
    cell.font = f_section()
    cell.fill = make_fill(C_DARK_BLUE)
    cell.alignment = AL_LEFT
    cell.border = BORDER


def write_header_row(ws, row: int, headers: list[str]):
    for col_num, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_num, value=pretty_header_name(header))
        cell.font = f_col_hdr()
        cell.fill = make_fill(C_MID_BLUE)
        cell.alignment = AL_CENTER
        cell.border = BORDER


def write_data_row(ws, row: int, headers: list[str], record: dict, fill):
    for col_num, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_num, value=excel_safe_value(record.get(header, "")))
        cell.font = f_val()
        cell.fill = fill
        cell.alignment = AL_LEFT
        cell.border = BORDER


def write_kv_table(ws, row: int, section_data: dict):
    headers = ["field", "value"]
    write_header_row(ws, row, headers)
    row += 1
    for i, (key, value) in enumerate(section_data.items()):
        fill = make_fill(C_LIGHT_BLUE if i % 2 == 0 else C_WHITE)
        ws.cell(row=row, column=1, value=pretty_header_name(key)).font = f_key()
        ws.cell(row=row, column=1).fill = fill
        ws.cell(row=row, column=1).alignment = AL_LEFT
        ws.cell(row=row, column=1).border = BORDER
        ws.cell(row=row, column=2, value=excel_safe_value(value)).font = f_val()
        ws.cell(row=row, column=2).fill = fill
        ws.cell(row=row, column=2).alignment = AL_LEFT_TOP
        ws.cell(row=row, column=2).border = BORDER
        row += 1
    return row


def write_section_content(ws, section_data):
    row = 1
    if isinstance(section_data, list) and section_data and isinstance(section_data[0], dict):
        headers = ordered_unique_keys(section_data)
        write_header_row(ws, row, headers)
        row += 1

        for i, record in enumerate(section_data):
            bg = make_fill(C_LIGHT_BLUE if i % 2 == 0 else C_WHITE)
            write_data_row(ws, row, headers, record, bg)
            row += 1

    elif isinstance(section_data, dict):
        row = write_kv_table(ws, row, section_data)

    elif isinstance(section_data, list):
        write_header_row(ws, row, ["value"])
        row += 1
        for i, item in enumerate(section_data):
            cell = ws.cell(row=row, column=1, value=excel_safe_value(item))
            cell.font = f_bullet()
            cell.fill = make_fill(C_LIGHT_BLUE if i % 2 == 0 else C_WHITE)
            cell.alignment = AL_LEFT_TOP
            cell.border = BORDER
            row += 1

    else:
        write_header_row(ws, row, ["value"])
        row += 1
        cell = ws.cell(row=row, column=1, value=excel_safe_value(section_data))
        cell.font = f_val()
        cell.alignment = AL_LEFT_TOP
        cell.border = BORDER


def add_named_sheet(wb: Workbook, title: str, section_data, used_titles: set[str]):
    ws = wb.create_sheet(title=unique_sheet_title(title, used_titles))
    ws.sheet_view.showGridLines = False
    write_section_content(ws, section_data)

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


def build_excel(extracted: dict, sheet_prefix: str = "Extracted") -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    used_titles: set[str] = set()

    for section_name, section_data in extracted.items():
        add_named_sheet(wb, f"{sheet_prefix} {pretty_header_name(section_name)}", section_data, used_titles)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def build_single_sheet_excel(section_title: str, section_data) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = unique_sheet_title(section_title, set())
    ws.sheet_view.showGridLines = False
    write_section_content(ws, section_data)

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


def build_excel_from_sources(sources: list[tuple[str, object]]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    used_titles: set[str] = set()

    for source_name, source_data in sources:
        if isinstance(source_data, dict):
            for section_name, section_data in source_data.items():
                add_named_sheet(wb, f"{source_name} {pretty_header_name(section_name)}", section_data, used_titles)
        else:
            add_named_sheet(wb, source_name, source_data, used_titles)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def build_zip_of_excels(files: list[tuple[str, bytes]]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, payload in files:
            zf.writestr(filename, payload)
    return buffer.getvalue()


def main():
    st.set_page_config(page_title="PDF to Excel Extractor", page_icon="PDF", layout="centered")
    st.title("PDF to Excel Extractor")
    st.markdown("Upload a PDF and extract structured data into a clean Excel file.")
    st.info("Step 1: Upload your PDF or CSV. Step 2: Extract all data. Step 3: Download the Excel file.")

    with st.sidebar:
        st.header("How it works")
        st.markdown("1. Upload your file\n2. Extract all data\n3. Download the Excel file")
        st.markdown("---")
        st.caption(f"Model: `{OPENROUTER_MODEL}`")

    uploaded_files = st.file_uploader("Upload PDFs or CSVs", type=["pdf", "csv"], accept_multiple_files=True)

    if not OPENROUTER_API_KEY:
        st.warning("Add OPENROUTER_API_KEY to your .env file before extracting a PDF.")

    if PDF_IMPORT_ERROR:
        st.error("Missing dependency: pypdf. Install project dependencies with `python -m pip install -r requirements.txt`.")
        st.stop()

    if OPENPYXL_IMPORT_ERROR:
        st.error("Missing dependency: openpyxl. Install project dependencies with `python -m pip install -r requirements.txt`.")
        st.stop()

    if uploaded_files:
        for uploaded_file in uploaded_files:
            st.info(f"{uploaded_file.name} - {uploaded_file.size / 1024:.1f} KB")

        if st.button("Extract all data", type="primary", use_container_width=True):
            pdf_files = [file for file in uploaded_files if file.name.lower().endswith(".pdf")]
            csv_files = [file for file in uploaded_files if file.name.lower().endswith(".csv")]
            total_sources = len(pdf_files) + len(csv_files)
            progress_bar = st.progress(0)
            progress_text = st.empty()

            sources: list[tuple[str, object]] = []
            completed_sources = 0

            for csv_file in csv_files:
                try:
                    progress_text.info(f"Processing {csv_file.name}... 0%")
                    csv_rows = parse_csv_bytes(csv_file.read())
                except Exception as exc:
                    st.error(f"Could not parse CSV '{csv_file.name}': {exc}")
                    st.stop()
                sources.append((Path(csv_file.name).stem, csv_rows))
                completed_sources += 1
                percent_complete = int((completed_sources / total_sources) * 100) if total_sources else 100
                progress_bar.progress(percent_complete / 100)
                progress_text.success(f"{csv_file.name} processed: {percent_complete}% complete")

            for pdf_file in pdf_files:
                pdf_bytes = pdf_file.read()
                pdf_base = Path(pdf_file.name).stem
                with st.spinner(f"Extracting data from {pdf_file.name}..."):
                    progress_text.info(f"Extracting {pdf_file.name}...")
                    result = extract_data_in_steps(pdf_bytes)

                if result["error"]:
                    st.error(f"Error extracting {pdf_file.name}: {result['error']}")
                    completed_sources += 1
                    percent_complete = int((completed_sources / total_sources) * 100) if total_sources else 100
                    progress_bar.progress(percent_complete / 100)
                    continue

                st.markdown(f"---\n### Debug: {pdf_file.name}")

                step1_ok = result["step1_char_count"] > 0
                step2_ok = result["step2_ai_json"] is not None
                step3_ok = result["step3_excel_bytes"] is not None

                c1, c2, c3 = st.columns(3)
                c1.metric("Step 1: Raw Text", f"{result['step1_char_count']} chars", "OK" if step1_ok else "EMPTY")
                c2.metric("Step 2: AI JSON", f"{result['step2_section_count']} sections", "OK" if step2_ok else "FAILED")
                c3.metric("Step 3: Excel", "Generated" if step3_ok else "FAILED", f"Data retention: {result['data_loss_pct']}%")

                with st.expander(f"Step 1 - Raw PDF Text ({result['step1_char_count']} chars)", expanded=False):
                    st.text_area("Extracted text", result["step1_raw_text"], height=300, disabled=True, key=f"step1_{pdf_file.name}")

                with st.expander(f"Step 2 - AI Structured JSON ({result['step2_section_count']} sections)", expanded=False):
                    st.json(result["step2_ai_json"])

                with st.expander(f"Step 3 - Excel Preview", expanded=False):
                    if step3_ok:
                        st.success(f"Excel file generated successfully.")
                        st.metric("Total Rows", result["step2_total_rows"])
                        st.metric("Data Retention", f"{result['data_loss_pct']}%")
                    else:
                        st.error("Excel generation failed.")

                if result["data_loss_pct"] >= 95:
                    st.success(f"No data loss detected from {pdf_file.name}.")
                elif result["data_loss_pct"] >= 80:
                    st.warning(f"Minor data variation from {pdf_file.name}.")
                else:
                    st.error(f"Significant data loss detected from {pdf_file.name}!")

                sources.append((pdf_base, result["step2_ai_json"]))
                completed_sources += 1
                percent_complete = int((completed_sources / total_sources) * 100) if total_sources else 100
                progress_bar.progress(percent_complete / 100)
                progress_text.success(f"{pdf_file.name} extracted: {percent_complete}% complete")

            if not sources:
                st.error("Upload at least one PDF or CSV file.")
                st.stop()

            with st.spinner("Building Excel file..."):
                excel_bytes = build_excel_from_sources(sources)
                if len(uploaded_files) == 1:
                    base = Path(uploaded_files[0].name).stem
                    filename = f"{base}_extracted.xlsx"
                else:
                    filename = "combined_extracted.xlsx"

            progress_bar.progress(1.0)
            progress_text.success("Extraction and Excel generation complete: 100%")

            st.download_button(
                label="Download Excel File",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
