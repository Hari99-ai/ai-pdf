import json
import re
import requests
from typing import Dict, Any, List, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from pdf_extractor.logging_config import logger
from pdf_extractor.pdf_parser import PDFPage
from pdf_extractor.ocr_engine import encode_image_to_base64

DETECTION_PROMPT = """You are a table extraction engine. Your ONLY job is to find every table on this page and extract it exactly as structured data.

CRITICAL: Detect every single table, grid, matrix, rate sheet, schedule, price list, fee schedule, comparison chart, and any data arranged in rows and columns — even if it has no visible borders. If data is aligned in columns, treat it as a table.

Output Schema:
{
  "sections": [
    {
      "name": "descriptive_table_name_in_snake_case",
      "type": "table",
      "headers": ["column_1", "column_2", "column_3", "context"],
      "rows": [
        {"column_1": "value_1", "column_2": "value_2", "column_3": "value_3", "context": "Brief explanation of what this row represents"}
      ]
    }
  ]
}

Rules:
1. EVERY table on the page must be a separate section with type "table".
2. Preserve ALL data exactly — do NOT round numbers, convert currencies, translate text, or change decimal places.
3. For tables: merge multi-word column names into a single key. Merge wrapped header lines into one string. Preserve original column order and row order.
4. Include ALL rows, even if they look like subtotals, totals, notes, or continuation rows.
5. If a table spans this page, include only what is visible on THIS page (merging across pages happens later).
6. If NO table exists on this page, return {"sections": []}.
7. Return ONLY valid JSON. No markdown fences, no explanations, no preamble.
8. IMPORTANT: Every row MUST include a "context" field with a brief, one-line explanation of what the row data means (e.g., "Rate per night for oceanfront room", "Seasonal pricing for holiday period", "Extra person charge"). This helps users understand the data at a glance.
"""

DETECTION_PROMPT_WITH_CONTEXT = """You are a table extraction engine processing a multi-page document sequentially. You are currently viewing PAGE {current_page} of {total_pages}.

You have already extracted data from previous pages. Here is a summary of what was found so far:
{previous_context}

CRITICAL: Detect every single table, grid, matrix, rate sheet, schedule, price list, fee schedule, comparison chart, and any data arranged in rows and columns — even if it has no visible borders. If data is aligned in columns, treat it as a table.

IMPORTANT: If this page continues a table from a previous page, use the same section name and headers to enable proper merging. For example, if "rates_grid" was started on page 1 and continues on page 2, use the same name "rates_grid" and matching headers.

Output Schema:
{
  "sections": [
    {
      "name": "descriptive_table_name_in_snake_case",
      "type": "table",
      "headers": ["column_1", "column_2", "column_3", "context"],
      "rows": [
        {"column_1": "value_1", "column_2": "value_2", "column_3": "value_3", "context": "Brief explanation of what this row represents"}
      ]
    }
  ]
}

Rules:
1. EVERY table on the page must be a separate section with type "table".
2. Preserve ALL data exactly — do NOT round numbers, convert currencies, translate text, or change decimal places.
3. For tables: merge multi-word column names into a single key. Merge wrapped header lines into one string. Preserve original column order and row order.
4. Include ALL rows, even if they look like subtotals, totals, notes, or continuation rows.
5. If a table spans this page, include only what is visible on THIS page (merging across pages happens later).
6. If NO table exists on this page, return {"sections": []}.
7. Return ONLY valid JSON. No markdown fences, no explanations, no preamble.
8. IMPORTANT: Every row MUST include a "context" field with a brief, one-line explanation of what the row data means (e.g., "Rate per night for oceanfront room", "Seasonal pricing for holiday period", "Extra person charge"). This helps users understand the data at a glance.
9. Use the previous context to maintain consistency in naming, headers, and data formatting across pages.
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
    model: str,
    previous_context: str = "",
    current_page: int = 1,
    total_pages: int = 1
) -> Dict[str, Any]:
    """Extract tables and data from a digital page using its text content."""
    if previous_context:
        prompt_text = DETECTION_PROMPT_WITH_CONTEXT.format(
            current_page=current_page,
            total_pages=total_pages,
            previous_context=previous_context
        )
    else:
        prompt_text = DETECTION_PROMPT

    prompt_content = [
        {"type": "text", "text": prompt_text},
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
    model: str,
    previous_context: str = "",
    current_page: int = 1,
    total_pages: int = 1
) -> Dict[str, Any]:
    """Extract tables and data from a scanned page.

    Strategy:
    1. Try sending the page image directly to the multimodal LLM.
    2. If the model rejects images, fall back to local OCR (pytesseract) to
       get text, then send that text to the text-based extraction path.
    """
    image_bytes = page.render_to_png(dpi=150)

    # --- attempt 1: image-based extraction ---
    try:
        base64_image = encode_image_to_base64(image_bytes)
        if previous_context:
            prompt_text = DETECTION_PROMPT_WITH_CONTEXT.format(
                current_page=current_page,
                total_pages=total_pages,
                previous_context=previous_context
            )
        else:
            prompt_text = DETECTION_PROMPT

        prompt_content = [
            {"type": "text", "text": prompt_text},
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{base64_image}"
                }
            }
        ]
        return call_openrouter_extraction(prompt_content, api_key, model)
    except Exception as e:
        logger.warning(f"Image extraction failed for page {page.page_num}: {e}")
        error_str = str(e).lower()
        image_not_supported = any(
            kw in error_str
            for kw in ("clipboard", "image", "cannot read", "not support")
        )

    # --- attempt 2: OCR → text-based extraction ---
    try:
        from pdf_extractor.ocr_engine import extract_text_with_local_tesseract
        ocr_text = extract_text_with_local_tesseract(image_bytes)
        if ocr_text:
            logger.info(f"OCR succeeded for page {page.page_num}, retrying as digital page.")
            return extract_from_digital_page(
                PDFPage(text=ocr_text, page_num=page.page_num, is_scanned=False),
                api_key,
                model,
                previous_context,
                current_page,
                total_pages
            )
    except Exception as ocr_err:
        logger.error(f"OCR fallback failed for page {page.page_num}: {ocr_err}")

    logger.error(f"All extraction methods failed for scanned page {page.page_num}")
    return {"sections": []}

def process_single_page(
    page: PDFPage,
    api_key: str,
    model: str,
    previous_context: str = "",
    current_page: int = 1,
    total_pages: int = 1
) -> Tuple[int, Dict[str, Any]]:
    """Process a single page (either digital or scanned) and return page index and sections."""
    logger.info(f"Starting page {page.page_num} processing (Scanned={page.is_scanned})...")
    if page.is_scanned:
        result = extract_from_scanned_page(page, api_key, model, previous_context, current_page, total_pages)
    else:
        result = extract_from_digital_page(page, api_key, model, previous_context, current_page, total_pages)
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

def build_context_summary(sections: List[Dict[str, Any]], page_num: int) -> str:
    """Build a concise summary of extracted sections from a page for context passing."""
    if not sections:
        return f"Page {page_num}: No tables found."

    summary_parts = []
    for sec in sections:
        name = sec.get("name", "unknown")
        sec_type = sec.get("type", "table")
        headers = sec.get("headers", [])
        rows = sec.get("rows", [])
        row_count = len(rows) if isinstance(rows, list) else 0

        header_str = ", ".join(headers[:5]) if headers else "no headers"
        if len(headers) > 5:
            header_str += f" (+{len(headers) - 5} more)"

        summary_parts.append(
            f"- Section '{name}' ({sec_type}): {row_count} rows, headers: [{header_str}]"
        )

    return f"Page {page_num}:\n" + "\n".join(summary_parts)

def process_all_pages_sequential(
    pages: List[PDFPage],
    api_key: str,
    model: str,
    progress_callback=None,
    previous_context: str = "",
) -> List[Dict[str, Any]]:
    """Extract data from all pages sequentially, passing context from each page to the next.

    This approach improves extraction quality for multi-page documents by providing
    the LLM with a summary of what was found on previous pages, enabling it to:
    - Maintain consistent naming for tables that span multiple pages
    - Use matching headers for proper merging
    - Understand the document structure and flow

    Args:
        previous_context: Optional context string from processing earlier pages
            (e.g. from a previous chunk). This is injected into the context
            history so the first page of this batch has awareness of prior extractions.
    """
    total_pages = len(pages)
    results = []
    context_history = []

    # Seed context history with any prior context from previous chunks
    if previous_context:
        context_history.append(previous_context)

    logger.info(f"Beginning sequential page extraction for {total_pages} pages...")
    for idx, page in enumerate(pages):
        current_page = idx + 1
        ctx = "\n".join(context_history) if context_history else ""

        logger.info(f"Processing page {current_page}/{total_pages} (Scanned={page.is_scanned})...")
        page_num, result = process_single_page(
            page, api_key, model, ctx, current_page, total_pages
        )
        results.append(result)

        # Build context summary for the next page
        sections = result.get("sections", [])
        summary = build_context_summary(sections, current_page)
        context_history.append(summary)

        # Keep context manageable - only keep last 3 pages of detailed context
        if len(context_history) > 3:
            # Summarize older pages into a compact form
            older = context_history[:-3]
            compact = f"Previously extracted: {len(older)} pages with tables from sections: "
            compact += ", ".join(
                sec.get("name", "?")
                for older_page_summary in older
                for line in older_page_summary.split("\n")
                if line.startswith("- Section ")
                for sec in [{"name": line.split("'")[1] if "'" in line else "unknown"}]
            )
            context_history = [compact] + context_history[-3:]

        if progress_callback:
            progress_callback(current_page, total_pages, f"Extracted page {current_page}/{total_pages}")

    logger.info(f"Sequential extraction complete. Processed {total_pages} pages.")
    return results

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
