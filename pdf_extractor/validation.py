import re
from typing import Dict, Any, List, Tuple
from pdf_extractor.logging_config import logger

def extract_numbers_and_currencies(text: str) -> set[str]:
    """Find all potential currency values and percentages in the raw text to audit preservation."""
    # Matches $320, 16%, $100.50, €50, etc.
    pattern = r"[\$€£]?\d+(?:\.\d+)?%?"
    matches = re.findall(pattern, text)
    # Filter out small numbers like page numbers (1-9) or short decimals
    return {m for m in matches if len(m) > 1 and not m.isdigit()}

def validate_extraction(
    groups: Dict[str, Dict[str, Any]],
    raw_pdf_text: str
) -> Tuple[bool, List[str]]:
    """
    Validate the extracted data before exporting it to Excel.
    Returns (is_valid, list_of_warnings_or_errors).
    """
    errors = []
    warnings = []
    
    # 1. Verify exactly 3 worksheets are present
    if len(groups) != 3:
        errors.append(f"Expected exactly 3 worksheets, but found {len(groups)}")

    total_rows = 0
    all_extracted_values = []
    
    for group_key, group_data in groups.items():
        title = group_data.get("sheet_title", "")
        sections = group_data.get("sections", [])
        
        if not title:
            errors.append(f"Worksheet key {group_key} has no sheet title.")
            
        for sec in sections:
            sec_name = sec.get("name", "")
            sec_type = sec.get("type", "table")
            rows = sec.get("rows", [])
            
            # Check headers
            if sec_type == "table":
                headers = sec.get("headers", [])
                if not headers:
                    errors.append(f"Table section '{sec_name}' has no headers defined.")
                elif any(not str(h).strip() for h in headers):
                    errors.append(f"Table section '{sec_name}' contains empty header names.")
            
            # Gather row count
            if isinstance(rows, list):
                total_rows += len(rows)
                for r in rows:
                    if isinstance(r, dict):
                        all_extracted_values.extend(str(v) for v in r.values())
                    else:
                        all_extracted_values.append(str(r))
            elif isinstance(rows, dict):
                total_rows += len(rows)
                all_extracted_values.extend(str(v) for v in rows.values())
                all_extracted_values.extend(str(k) for k in rows.keys())
                
    # 2. If PDF has text, we should have extracted at least some rows
    if len(raw_pdf_text.strip()) > 200 and total_rows == 0:
        errors.append("PDF contains substantial text, but zero rows of structured data were extracted.")

    # 3. Numeric accuracy: check if currency and percentage strings are preserved
    if raw_pdf_text:
        source_currencies = extract_numbers_and_currencies(raw_pdf_text)
        logger.info(f"Identified {len(source_currencies)} currency/percentage values in source PDF text.")
        
        extracted_text_block = " ".join(all_extracted_values)
        missing_values = []
        for val in source_currencies:
            if val not in extracted_text_block:
                missing_values.append(val)
                
        # If a significant number of currency/percentage symbols are missing, trigger a warning
        if missing_values and len(source_currencies) > 0:
            missing_ratio = len(missing_values) / len(source_currencies)
            if missing_ratio > 0.4:  # If more than 40% are missing
                warnings.append(
                    f"Warning: {len(missing_values)} currency/percentage values from PDF text were not found in Excel: {missing_values[:10]}"
                )

    # 4. Check for duplicate rows in sections
    for group_key, group_data in groups.items():
        sections = group_data.get("sections", [])
        for sec in sections:
            sec_name = sec.get("name", "")
            rows = sec.get("rows", [])
            if isinstance(rows, list) and len(rows) > 0:
                row_strs = [str(r) for r in rows]
                if len(row_strs) != len(set(row_strs)):
                    warnings.append(f"Section '{sec_name}' in sheet '{group_data.get('sheet_title')}' contains duplicate rows.")

    is_valid = len(errors) == 0
    all_messages = errors + warnings
    
    if errors:
        logger.error(f"Validation failed with errors: {errors}")
    if warnings:
        logger.warning(f"Validation warnings: {warnings}")
        
    return is_valid, all_messages
