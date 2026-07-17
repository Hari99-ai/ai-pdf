import json
import re
import requests
from typing import Dict, Any, List, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from pdf_extractor.logging_config import logger
from pdf_extractor.pdf_parser import PDFPage
from pdf_extractor.ocr_engine import encode_image_to_base64

# ---------------------------------------------------------------------------
# LLM Extraction Prompts
# ---------------------------------------------------------------------------

DETECTION_PROMPT = """You are an expert data analyst and extraction engine for hotel/travel contract PDFs. Your job is to extract all tables, grids, list items, key-value configurations, policies, and property metadata from this page as structured sections.

CRITICAL:
1. Detect every single table, grid, matrix, rate sheet, schedule, price list, fee schedule, comparison chart, and any data arranged in rows and columns.
2. If the page contains unstructured information (such as Property Name, Address, Contact details, Validity dates, Currency, Tax rates, Booking policies, Cancellation clauses, etc.), you MUST also extract them as structured "key_value" or "list" sections. Do NOT ignore them.

Output Schema:
{
  "sections": [
    {
      "name": "descriptive_section_name_in_snake_case",
      "type": "table", // Use "table", "key_value", or "list"
      "headers": ["column_1", "column_2", "context"], // Required ONLY if type is "table". Merge multi-word headers and preserve column order.
      "rows": [
        // For type "table", this is a list of dicts (one per row), e.g., [{"col_1": "val_1", "col_2": "val_2", "context": "..."}]
        // For type "key_value", this is a single dict of key-value pairs, e.g., {"property_name": "Ocean Breeze Resort", "currency": "INR"}
        // For type "list", this is a list of strings, e.g., ["Free cancellation up to 7 days before check-in.", "Breakfast included."]
      ]
    }
  ]
}

HIERARCHY RULES (critical for hotel contracts):
Hotel data is hierarchical: Season > Date Range > Room Type > Rate. You MUST preserve this hierarchy in every row using these fields when applicable:
- "season": The season/period name (e.g., "High Season", "Low Season", "Peak", "Festive")
- "date_from": The start date of the period (e.g., "Jan 01", "2025-01-01")
- "date_to": The end date of the period (e.g., "Mar 31", "2025-03-31")
- "meal_plan": The meal plan (e.g., "BB", "HB", "FB", "AI", "Room Only")
- "min_stay": Minimum stay requirement (e.g., "3 nights", "5 nights")

If the PDF shows a rate grid where seasons have sub-rows for different date ranges and room types, EACH data row MUST include the parent season name, date range, meal plan, and min stay as separate fields. Do NOT leave these blank — copy them down from the parent row.

ROW ALIGNMENT RULES (critical for "Other Charges" and fee tables):
When a table has multiple fee/charge items, each row MUST be self-contained with ALL columns filled. If a row's column is empty because the value was stated once in a merged cell above, REPEAT the value in every row. For example:
- If "Other Charges" has sub-items (Resort Fee, Parking, etc.) under a category "Mandatory", every sub-item row must have "Mandatory" in the category column.
- If a rate applies to multiple room types, create a separate row for EACH room type with the rate repeated.

General Rules:
1. Every table, key-value block, or list block must be a separate section with the correct type.
2. Preserve ALL data exactly — do NOT round numbers, convert currencies, translate text, or change decimal places. Keep dates, prices, phone numbers, and percentage values (like tax or commission rates) exactly as written.
3. Every row in type "table" must include a "context" field explaining the row data.
4. If no relevant information of any kind exists on this page, return {"sections": []}.
5. Return ONLY valid JSON. No markdown fences, no explanations, no preamble.
"""

DETECTION_PROMPT_WITH_CONTEXT = """You are an expert data analyst and extraction engine processing a multi-page hotel/travel contract document sequentially. You are currently viewing PAGE {current_page} of {total_pages}.

You have already extracted data from previous pages. Here is a summary of what was found so far:
{previous_context}

Your job is to extract all tables, grids, list items, key-value configurations, policies, and property metadata from this page as structured sections.

CRITICAL:
1. Detect every single table, grid, matrix, rate sheet, schedule, price list, fee schedule, comparison chart, and any data arranged in rows and columns.
2. If the page contains unstructured information (such as Property Name, Address, Contact details, Validity dates, Currency, Tax rates, Booking policies, cancellation clauses, etc.), you MUST also extract them as structured "key_value" or "list" sections. Do NOT ignore them.
3. If this page continues a table from a previous page, use the same section name and headers to enable proper merging. For example, if "rates_grid" was started on page 1 and continues on page 2, use the same name "rates_grid" and matching headers.

Output Schema:
{
  "sections": [
    {
      "name": "descriptive_section_name_in_snake_case",
      "type": "table", // Use "table", "key_value", or "list"
      "headers": ["column_1", "column_2", "context"], // Required ONLY if type is "table". Merge multi-word headers and preserve column order.
      "rows": [
        // For type "table", this is a list of dicts (one per row), e.g., [{"col_1": "val_1", "col_2": "val_2", "context": "..."}]
        // For type "key_value", this is a single dict of key-value pairs, e.g., {"property_name": "Ocean Breeze Resort", "currency": "INR"}
        // For type "list", this is a list of strings, e.g., ["Free cancellation up to 7 days before check-in.", "Breakfast included."]
      ]
    }
  ]
}

HIERARCHY RULES (critical for hotel contracts):
Hotel data is hierarchical: Season > Date Range > Room Type > Rate. You MUST preserve this hierarchy in every row using these fields when applicable:
- "season": The season/period name (e.g., "High Season", "Low Season", "Peak", "Festive")
- "date_from": The start date of the period (e.g., "Jan 01", "2025-01-01")
- "date_to": The end date of the period (e.g., "Mar 31", "2025-03-31")
- "meal_plan": The meal plan (e.g., "BB", "HB", "FB", "AI", "Room Only")
- "min_stay": Minimum stay requirement (e.g., "3 nights", "5 nights")

If the PDF shows a rate grid where seasons have sub-rows for different date ranges and room types, EACH data row MUST include the parent season name, date range, meal plan, and min stay as separate fields. Do NOT leave these blank — copy them down from the parent row.

ROW ALIGNMENT RULES (critical for "Other Charges" and fee tables):
When a table has multiple fee/charge items, each row MUST be self-contained with ALL columns filled. If a row's column is empty because the value was stated once in a merged cell above, REPEAT the value in every row.

General Rules:
1. Every table, key-value block, or list block must be a separate section with the correct type.
2. Preserve ALL data exactly — do NOT round numbers, convert currencies, translate text, or change decimal places. Keep dates, prices, phone numbers, and percentage values (like tax or commission rates) exactly as written.
3. Every row in type "table" must include a "context" field explaining the row data.
4. If no relevant information of any kind exists on this page, return {"sections": []}.
5. Return ONLY valid JSON. No markdown fences, no explanations, no preamble.
6. Use the previous context to maintain consistency in naming, headers, and data formatting across pages.
"""


# ---------------------------------------------------------------------------
# JSON helpers
# ---------------------------------------------------------------------------

def clean_json_text(text: str) -> str:
    """Clean the raw LLM response to get pure JSON."""
    raw = text.strip()
    # Strip markdown block fences if present
    if raw.startswith("```"):
        raw = "\n".join(line for line in raw.splitlines() if not line.strip().startswith("```")).strip()
    return raw


def _repair_truncated_json(text: str) -> str:
    """Attempt to repair JSON that was truncated by max_tokens limit."""
    if not text:
        return text

    try:
        json.loads(text)
        return text
    except json.JSONDecodeError:
        pass

    repaired = text.rstrip()

    brace_positions = [i for i, ch in enumerate(repaired) if ch == "}"]
    best_result = None

    for pos in reversed(brace_positions):
        candidate = repaired[:pos + 1].rstrip().rstrip(",")
        if not candidate:
            continue

        stack = []
        in_str = False
        esc = False
        for ch in candidate:
            if esc:
                esc = False
                continue
            if ch == "\\":
                esc = True
                continue
            if ch == '"':
                in_str = not in_str
                continue
            if in_str:
                continue
            if ch in ("{", "["):
                stack.append(ch)
            elif ch == "}" and stack and stack[-1] == "{":
                stack.pop()
            elif ch == "]" and stack and stack[-1] == "[":
                stack.pop()

        if in_str:
            continue

        closer = ""
        for opener in reversed(stack):
            closer += "]" if opener == "[" else "}"

        attempt = candidate + closer
        try:
            result = json.loads(attempt)
            if isinstance(result, dict) and result.get("sections"):
                logger.info(f"Repaired truncated JSON at pos {pos}, preserving data.")
                return attempt
            if best_result is None and isinstance(result, dict):
                best_result = attempt
        except json.JSONDecodeError:
            continue

    if best_result:
        logger.info("Using best-effort repair (empty sections).")
        return best_result

    logger.warning("All JSON repair attempts failed. Returning empty sections.")
    return '{"sections": []}'


# ---------------------------------------------------------------------------
# LLM API call
# ---------------------------------------------------------------------------

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
            "max_tokens": 8192
        },
        timeout=180
    )
    response.raise_for_status()
    payload = response.json()
    content = payload["choices"][0]["message"]["content"]
    cleaned = clean_json_text(content)

    # Try direct parse first
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        logger.warning(f"Initial JSON parse failed ({e}). Attempting repair...")
        repaired = _repair_truncated_json(cleaned)
        return json.loads(repaired)


# ---------------------------------------------------------------------------
# Page extraction
# ---------------------------------------------------------------------------

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
        return {"sections": []}


def extract_from_scanned_page(
    page: PDFPage,
    api_key: str,
    model: str,
    previous_context: str = "",
    current_page: int = 1,
    total_pages: int = 1
) -> Dict[str, Any]:
    """Extract tables and data from a scanned page."""
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

    # --- attempt 2: OCR -> text-based extraction ---
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
    """Process a single page and return page index and sections."""
    logger.info(f"Starting page {page.page_num} processing (Scanned={page.is_scanned})...")
    if page.is_scanned:
        result = extract_from_scanned_page(page, api_key, model, previous_context, current_page, total_pages)
    else:
        result = extract_from_digital_page(page, api_key, model, previous_context, current_page, total_pages)
    logger.info(f"Finished page {page.page_num} processing. Found {len(result.get('sections', []))} sections.")
    return page.page_num, result


# ---------------------------------------------------------------------------
# Parallel and sequential processing
# ---------------------------------------------------------------------------

def process_all_pages_parallel(
    pages: List[PDFPage],
    api_key: str,
    model: str,
    progress_callback=None,
    max_workers: int = 5
) -> List[Dict[str, Any]]:
    """Extract data from all pages in parallel."""
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

        header_str = ", ".join(headers[:6]) if headers else "no headers"
        if len(headers) > 6:
            header_str += f" (+{len(headers) - 6} more)"

        # Include hierarchy fields if present in rows
        hierarchy_info = ""
        if rows and isinstance(rows, list) and isinstance(rows[0], dict):
            sample = rows[0]
            hier_fields = ["season", "date_from", "date_to", "meal_plan", "min_stay"]
            found = [f"{f}={sample[f]}" for f in hier_fields if f in sample and sample[f]]
            if found:
                hierarchy_info = f", hierarchy: [{', '.join(found)}]"

        summary_parts.append(
            f"- Section '{name}' ({sec_type}): {row_count} rows, headers: [{header_str}]{hierarchy_info}"
        )

    return f"Page {page_num}:\n" + "\n".join(summary_parts)


def process_all_pages_sequential(
    pages: List[PDFPage],
    api_key: str,
    model: str,
    progress_callback=None,
    previous_context: str = "",
) -> List[Dict[str, Any]]:
    """Extract data from all pages sequentially, passing context from each page to the next."""
    total_pages = len(pages)
    results = []
    context_history = []

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

        sections = result.get("sections", [])
        summary = build_context_summary(sections, current_page)
        context_history.append(summary)

        # Keep last 5 pages of detailed context (increased from 3)
        if len(context_history) > 5:
            older = context_history[:-5]
            compact = f"Previously extracted: {len(older)} pages with tables from sections: "
            compact += ", ".join(
                sec.get("name", "?")
                for older_page_summary in older
                for line in older_page_summary.split("\n")
                if line.startswith("- Section ")
                for sec in [{"name": line.split("'")[1] if "'" in line else "unknown"}]
            )
            context_history = [compact] + context_history[-5:]

        if progress_callback:
            progress_callback(current_page, total_pages, f"Extracted page {current_page}/{total_pages}")

    logger.info(f"Sequential extraction complete. Processed {total_pages} pages.")
    return results


# ---------------------------------------------------------------------------
# Post-extraction cleanup & merge
# ---------------------------------------------------------------------------

def clean_repeated_headers(rows: List[Dict[str, Any]], headers: List[str]) -> List[Dict[str, Any]]:
    """Detect and remove rows that contain duplicate headers."""
    cleaned_rows = []
    header_lower_set = {h.lower().strip() for h in headers}

    for row in rows:
        row_values_lower = {str(val).lower().strip() for val in row.values()}
        if row_values_lower and row_values_lower == header_lower_set:
            logger.info(f"Removing repeated header row: {row}")
            continue
        cleaned_rows.append(row)
    return cleaned_rows


def clean_duplicate_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Preserve all rows to avoid silently dropping repeated values."""
    return rows


def _rows_are_hierarchy_compatible(prev_rows: List[Dict], curr_rows: List[Dict]) -> bool:
    """Check if rows from two sections can be merged without misalignment.

    Returns False if the sections have overlapping row keys but different
    hierarchy context (season/date_from/date_to), which would cause misalignment.
    """
    if not prev_rows or not curr_rows:
        return True

    hier_keys = {"season", "date_from", "date_to", "meal_plan", "min_stay"}

    prev_has_hier = any(k in row for row in prev_rows[:3] for k in hier_keys if row.get(k))
    curr_has_hier = any(k in row for row in curr_rows[:3] for k in hier_keys if row.get(k))

    # If both have hierarchy fields, check they're compatible (not conflicting)
    if prev_has_hier and curr_has_hier:
        # Check if the last row of previous and first row of current have
        # different season/date values — this means they're different tables
        prev_last = prev_rows[-1]
        curr_first = curr_rows[0]

        for key in ["season", "date_from", "date_to"]:
            pv = str(prev_last.get(key, "")).strip().lower()
            cv = str(curr_first.get(key, "")).strip().lower()
            if pv and cv and pv != cv:
                logger.info(f"Rejecting merge: hierarchy mismatch on '{key}' ('{pv}' vs '{cv}')")
                return False

    return True


def merge_consecutive_tables(sections: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Merge consecutive table sections that have identical schemas.

    Merge is only allowed when:
    1. Headers match as ordered sets (not just set equality)
    2. Section names are similar
    3. Row hierarchy context is compatible
    """
    if not sections:
        return []

    merged_sections: List[Dict[str, Any]] = []

    for current in sections:
        if not merged_sections:
            merged_sections.append(current)
            continue

        previous = merged_sections[-1]

        if current.get("type") == "table" and previous.get("type") == "table":
            prev_headers = [h.lower().strip() for h in previous.get("headers", [])]
            curr_headers = [h.lower().strip() for h in current.get("headers", [])]

            # Check 1: Exact header set match
            headers_match = set(prev_headers) == set(curr_headers) and len(prev_headers) > 0

            # Check 2: Ordered header prefix match (first N headers match in order)
            # This is more reliable for continuation tables
            min_len = min(len(prev_headers), len(curr_headers))
            prefix_match = (
                min_len >= 3
                and prev_headers[:min_len] == curr_headers[:min_len]
            )

            # Check 3: Name similarity + same column count
            names_match = (
                re.sub(r'[^a-zA-Z]', '', previous.get("name", "")).lower() ==
                re.sub(r'[^a-zA-Z]', '', current.get("name", "")).lower()
            ) and len(prev_headers) == len(curr_headers)

            # Check 4: Name is a strict prefix/suffix of the other
            prev_name = re.sub(r'[^a-zA-Z]', '', previous.get("name", "")).lower()
            curr_name = re.sub(r'[^a-zA-Z]', '', current.get("name", "")).lower()
            name_contains = (
                (prev_name and curr_name and (prev_name in curr_name or curr_name in prev_name))
                and len(prev_headers) == len(curr_headers)
            )

            can_merge = (headers_match or prefix_match or names_match or name_contains)

            # Additional safety: check hierarchy compatibility
            if can_merge:
                prev_rows = previous.get("rows", [])
                curr_rows = current.get("rows", [])
                if not _rows_are_hierarchy_compatible(prev_rows, curr_rows):
                    can_merge = False
                    logger.info(f"Rejecting merge of '{previous.get('name')}' and '{current.get('name')}' due to hierarchy mismatch")

            if can_merge:
                logger.info(f"Merging consecutive tables: '{previous.get('name')}' and '{current.get('name')}'")

                prev_rows = previous.get("rows", [])
                curr_rows = current.get("rows", [])

                # Standardize column keys by name, not by position
                standardized_curr_rows = []
                for row in curr_rows:
                    new_row = {}
                    prev_h = previous.get("headers", [])
                    curr_h = current.get("headers", [])
                    curr_lookup = {str(h).strip().lower(): h for h in curr_h}
                    for ph in prev_h:
                        curr_key = curr_lookup.get(str(ph).strip().lower())
                        new_row[ph] = row.get(curr_key, "") if curr_key is not None else ""
                    standardized_curr_rows.append(new_row)

                combined_rows = prev_rows + standardized_curr_rows
                combined_rows = clean_repeated_headers(combined_rows, previous.get("headers", []))
                combined_rows = clean_duplicate_rows(combined_rows)

                previous["rows"] = combined_rows
                # Update headers to use the (potentially larger) union
                all_headers = list(dict.fromkeys(previous.get("headers", []) + current.get("headers", [])))
                previous["headers"] = all_headers
                continue

        merged_sections.append(current)

    return merged_sections


def compile_and_merge_sections(page_results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Compile all sections from all pages and merge consecutive matches."""
    all_sections = []
    for page_res in page_results:
        sections = page_res.get("sections", [])
        for sec in sections:
            if not sec.get("name") or not sec.get("type"):
                continue
            all_sections.append(sec)

    logger.info(f"Total sections extracted before merging: {len(all_sections)}")
    merged = merge_consecutive_tables(all_sections)
    logger.info(f"Total sections after merging: {len(merged)}")
    return merged
