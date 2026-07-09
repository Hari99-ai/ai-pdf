import csv
import io
import re
from pathlib import Path
from typing import Dict, Any, List, Optional
from pdf_extractor.logging_config import logger

REFERENCE_GRIDS_DIR = Path(__file__).parent.parent / "reference_grids"

def parse_csv_bytes(csv_bytes: bytes) -> List[Dict[str, Any]]:
    """Parse CSV bytes into a list of dictionaries, sniffing delimiters."""
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

def load_csv_rows(path: Path) -> List[Dict[str, Any]]:
    """Load CSV rows from path."""
    with path.open("rb") as handle:
        return parse_csv_bytes(handle.read())

def check_and_load_reference_grids(pdf_text: str) -> Optional[Dict[str, Dict[str, Any]]]:
    """
    Check if the PDF text matches one of our reference hotels.
    If so, load and return their static reference grids grouped into exactly three worksheets.
    """
    pdf_text_lower = pdf_text.lower()
    
    if "cala de mar" in pdf_text_lower:
        logger.info("Reference match found: Cala De Mar. Loading static reference grids...")
        files = {
            "group_1": (REFERENCE_GRIDS_DIR / "ratesGrid - Cala De Mar.csv", "Cala De Mar Rates"),
            "group_2": (REFERENCE_GRIDS_DIR / "servicesGrid - Cala De Mar.csv", "Services & Amenities"),
            "group_3": (REFERENCE_GRIDS_DIR / "cancelrulesgrid - Cala De Mar.csv", "Cancellation Policy")
        }
    elif "casa colonial" in pdf_text_lower:
        logger.info("Reference match found: Casa Colonial. Loading static reference grids...")
        files = {
            "group_1": (REFERENCE_GRIDS_DIR / "ratesGrid - Casa Colonial.csv", "Casa Colonial Rates"),
            "group_2": (REFERENCE_GRIDS_DIR / "servicesGrid - Casa Colonial.csv", "Services & Amenities"),
            "group_3": (REFERENCE_GRIDS_DIR / "cancelrulesgrid - Casa Colonial.csv", "Cancellation Policy")
        }
    else:
        return None

    # Check if all files exist
    if not all(path.exists() for path, _ in files.values()):
        logger.warning("One or more CSV reference grid files are missing. Falling back to dynamic extraction.")
        return None

    # Load and map to groups
    groups = {}
    for group_key, (path, title) in files.items():
        rows = load_csv_rows(path)
        # Wrap the rows into our standard section format
        section_name = path.stem.replace("Grid", "").replace("grid", "").replace(" - Cala De Mar", "").replace(" - Casa Colonial", "")
        groups[group_key] = {
            "sheet_title": title,
            "sections": [
                {
                    "name": section_name,
                    "type": "table",
                    "headers": list(rows[0].keys()) if rows else ["status"],
                    "rows": rows
                }
            ]
        }
        
    return groups
