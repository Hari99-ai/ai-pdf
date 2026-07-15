import json
import re
import requests
from typing import Dict, Any, List, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from pdf_extractor.logging_config import logger
from pdf_extractor.pdf_parser import PDFPage
from pdf_extractor.ocr_engine import encode_image_to_base64

DETECTION_PROMPT = """Analyze this document page and extract ALL structured data (tables, price lists, policies, terms, rates, fee schedules, promotions, amenities) into a single JSON object.

Output Schema:
{
  "sections": [
    {
      "name": "section_descriptive_name_in_snake_case",
      "type": "table",  // Must be one of: "table", "key_value", "list"
      "headers": ["column_1", "column_2"], // Required if type is "table"
      "rows": [
        {"column_1": "value_1", "column_2": "value_2"}
      ] // If type is "table", this is a list of row objects. If type is "key_value", this is a single object of key-value pairs. If type is "list", this is a list of strings.
    }
  ]
}

Rules for Extraction:
1. Preserve all data exactly. Do NOT round numbers, convert currencies, recalculate values, translate text, or change decimal places.
2. Ignore all document noise: headers, footers, page numbers, logos, signatures, watermarks, decorative lines.
3. For tables: keep multi-word column names combined into a single key. If a header wraps multiple lines, merge it into a single clean string. Maintain original columns and row ordering.
4. Return ONLY valid JSON matching the schema. No explanations, no markdown block fences (e.g. do not wrap in ```json), no preamble.
"""

def clean_json_text(text: str) -> str:
    """Clean the raw LLM response to get pure JSON."""
    raw = text.strip()
    # Strip markdown block fences if present
    if raw.startswith("```"):
        raw = "\n".join(line for line in raw.splitlines() if not line.strip().startswith("```")).strip()
    return raw

def call_openrouter_extraction(
    prompt_payload: Dict[str, Any],
    api_key: str,
    model: str
) -> Dict[str, Any]:
    """Call OpenRouter with the given payload and parse the JSON response."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    response = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers=headers,
        json={
            "model": model,
            "messages": [{"role": "user", "content": prompt_payload}],
            "temperature": 0.0,
            "max_tokens": 4090
        },
        timeout=180
    )
    response.raise_for_status()
    payload = response.json()
    content = payload["choices"][0]["message"]["content"]
    cleaned = clean_json_text(content)
    return json.loads(cleaned)

def extract_from_digital_page(
    page: PDFPage,
    api_key: str,
    model: str
) -> Dict[str, Any]:
    """Extract tables and data from a digital page using its text content."""
    prompt_content = [
        {"type": "text", "text": DETECTION_PROMPT},
        {"type": "text", "text": f"PAGE TEXT:\n{page.text}"}
    ]
    try:
        return call_openrouter_extraction(prompt_content, api_key, model)
    except Exception as e:
        logger.error(f"Failed to extract from digital page {page.page_num}: {e}")
        # Return empty sections to avoid crash
        return {"sections": []}

def extract_from_scanned_page(
    page: PDFPage,
    api_key: str,
    model: str
) -> Dict[str, Any]:
    """Extract tables and data from a scanned page using its image representation."""
    try:
        image_bytes = page.render_to_png(dpi=150)
        base64_image = encode_image_to_base64(image_bytes)
        
        prompt_content = [
            {"type": "text", "text": DETECTION_PROMPT},
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{base64_image}"
                }
            }
        ]
        return call_openrouter_extraction(prompt_content, api_key, model)
    except Exception as e:
        logger.error(f"Failed to extract from scanned page {page.page_num}: {e}")
        return {"sections": []}

def process_single_page(
    page: PDFPage,
    api_key: str,
    model: str
) -> Tuple[int, Dict[str, Any]]:
    """Process a single page (either digital or scanned) and return page index and sections."""
    logger.info(f"Starting page {page.page_num} processing (Scanned={page.is_scanned})...")
    if page.is_scanned:
        result = extract_from_scanned_page(page, api_key, model)
    else:
        result = extract_from_digital_page(page, api_key, model)
    logger.info(f"Finished page {page.page_num} processing. Found {len(result.get('sections', []))} sections.")
    return page.page_num, result

def process_all_pages_parallel(
    pages: List[PDFPage],
    api_key: str,
    model: str,
    progress_callback=None,
    max_workers: int = 5
) -> List[Dict[str, Any]]:
    """Extract data from all pages in parallel, invoking the progress callback as they complete."""
    results_dict = {}
    total_pages = len(pages)
    
    logger.info(f"Beginning parallel page extraction with {max_workers} workers...")
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(process_single_page, page, api_key, model): page
            for page in pages
        }
        
        completed_count = 0
        for future in as_completed(futures):
            page_num, result = future.result()
            results_dict[page_num] = result
            completed_count += 1
            if progress_callback:
                progress_callback(completed_count, total_pages, f"Extracted page {page_num} ({completed_count}/{total_pages})")
                
    # Sort results by page number to preserve document order
    sorted_results = [results_dict[i] for i in range(1, total_pages + 1)]
    return sorted_results

def clean_repeated_headers(rows: List[Dict[str, Any]], headers: List[str]) -> List[Dict[str, Any]]:
    """Detect and remove rows that contain duplicate headers (e.g. repeated table headers on page breaks)."""
    cleaned_rows = []
    header_lower_set = {h.lower().strip() for h in headers}
    
    for row in rows:
        # Check if the row matches the header names themselves
        row_values_lower = {str(val).lower().strip() for val in row.values()}

        # Only drop rows that are an exact header repeat.
        # A subset check is too aggressive and can remove legitimate data rows.
        if row_values_lower and row_values_lower == header_lower_set:
            logger.info(f"Removing repeated header row: {row}")
            continue
            
        cleaned_rows.append(row)
    return cleaned_rows

def clean_duplicate_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Preserve all rows to avoid silently dropping repeated values."""
    # Duplicate-looking rows can still be valid data in pricing/policy tables.
    # Keep them all and let downstream consumers decide how to present them.
    return rows

def merge_consecutive_tables(sections: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Merge consecutive table sections that have identical or highly similar schemas (headers)."""
    if not sections:
        return []
        
    merged_sections: List[Dict[str, Any]] = []
    
    for current in sections:
        if not merged_sections:
            merged_sections.append(current)
            continue
            
        previous = merged_sections[-1]
        
        # Check if they can be merged:
        # 1. Both must be tables
        # 2. They must have similar keys/names or column layouts
        if current.get("type") == "table" and previous.get("type") == "table":
            prev_headers = [h.lower().strip() for h in previous.get("headers", [])]
            curr_headers = [h.lower().strip() for h in current.get("headers", [])]
            
            # Simple set comparison to see if headers match
            headers_match = set(prev_headers) == set(curr_headers) and len(prev_headers) > 0
            
            # Or if names are extremely similar and column count is identical
            names_match = (
                re.sub(r'[^a-zA-Z]', '', previous.get("name", "")).lower() == 
                re.sub(r'[^a-zA-Z]', '', current.get("name", "")).lower()
            ) and len(prev_headers) == len(curr_headers)
            
            if headers_match or names_match:
                logger.info(f"Merging consecutive tables: '{previous.get('name')}' and '{current.get('name')}'")
                
                # Merge rows
                prev_rows = previous.get("rows", [])
                curr_rows = current.get("rows", [])
                
                # Standardize column keys by name, not by position.
                # This avoids corrupting values when header order changes between pages.
                standardized_curr_rows = []
                for row in curr_rows:
                    new_row = {}
                    prev_headers = previous.get("headers", [])
                    curr_headers = current.get("headers", [])
                    curr_lookup = {str(h).strip().lower(): h for h in curr_headers}
                    for prev_h in prev_headers:
                        curr_key = curr_lookup.get(str(prev_h).strip().lower())
                        new_row[prev_h] = row.get(curr_key, "") if curr_key is not None else ""
                    standardized_curr_rows.append(new_row)
                
                # Combine rows and clean up
                combined_rows = prev_rows + standardized_curr_rows
                combined_rows = clean_repeated_headers(combined_rows, previous.get("headers", []))
                combined_rows = clean_duplicate_rows(combined_rows)
                
                previous["rows"] = combined_rows
                continue
                
        merged_sections.append(current)
        
    return merged_sections

def compile_and_merge_sections(page_results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Compile all sections from all pages and merge consecutive matches."""
    all_sections = []
    for page_res in page_results:
        sections = page_res.get("sections", [])
        for sec in sections:
            # Simple schema validation
            if not sec.get("name") or not sec.get("type"):
                continue
            all_sections.append(sec)
            
    logger.info(f"Total sections extracted before merging: {len(all_sections)}")
    merged = merge_consecutive_tables(all_sections)
    logger.info(f"Total sections after merging: {len(merged)}")
    return merged
