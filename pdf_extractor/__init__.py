import os
from typing import Dict, Any, List, Tuple, Callable, Optional
from pdf_extractor.logging_config import logger
from pdf_extractor.pdf_parser import parse_pdf, check_pdf_validity, PDFPage
from pdf_extractor.table_detector import process_all_pages_sequential, compile_and_merge_sections, build_context_summary
from pdf_extractor.classifier import classify_sections
from pdf_extractor.excel_generator import build_excel_workbook
from pdf_extractor.validation import validate_extraction
from pdf_extractor.reference_grids import check_and_load_reference_grids

DEFAULT_CHUNK_SIZE = 15


def _chunk_pages(pages: List[PDFPage], chunk_size: int) -> List[List[PDFPage]]:
    """Split a list of PDF pages into chunks of at most chunk_size pages each."""
    if chunk_size <= 0:
        chunk_size = DEFAULT_CHUNK_SIZE
    return [pages[i:i + chunk_size] for i in range(0, len(pages), chunk_size)]


def _build_chunk_context_summary(page_results: List[Dict[str, Any]], chunk_start: int) -> str:
    """Build a compact summary of a chunk's extracted sections for cross-chunk context."""
    all_section_names = []
    total_rows = 0
    for page_res in page_results:
        for sec in page_res.get("sections", []):
            name = sec.get("name", "unknown")
            all_section_names.append(name)
            rows = sec.get("rows", [])
            if isinstance(rows, list):
                total_rows += len(rows)

    if not all_section_names:
        return f"Pages {chunk_start}-{chunk_start + len(page_results) - 1}: No tables found."

    unique_names = list(dict.fromkeys(all_section_names))
    names_str = ", ".join(unique_names[:10])
    if len(unique_names) > 10:
        names_str += f" (+{len(unique_names) - 10} more)"

    return f"Pages {chunk_start}-{chunk_start + len(page_results) - 1}: {total_rows} total rows from sections: [{names_str}]"


def process_pdf(
    pdf_bytes: bytes,
    api_key: str,
    model: str = "openai/gpt-4o-mini",
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    max_workers: int = 5,
    chunk_size: int = DEFAULT_CHUNK_SIZE,
) -> Tuple[bytes, Dict[str, Any], List[str]]:
    """
    Orchestrate the entire PDF-to-Excel workflow with self-healing retries.

    For large PDFs, pages are split into chunks (default 15 pages each) and
    processed separately. Chunk results are then merged so that tables spanning
    chunk boundaries are properly combined.

    Args:
        chunk_size: Maximum number of pages per processing chunk.
                    Defaults to 15. Set to 0 to process all pages at once
                    (original behavior).

    Returns:
        - Excel workbook bytes
        - Classified groups dictionary (useful for preview)
        - List of validation warnings or messages
    """
    # Step 1: Validate PDF file structure and integrity
    is_valid_pdf, error_msg = check_pdf_validity(pdf_bytes)
    if not is_valid_pdf:
        raise ValueError(error_msg)

    # Step 2: Parse PDF pages (check digital text & prepare pages)
    pages = parse_pdf(pdf_bytes)
    raw_pdf_text = "\n\n".join(page.text for page in pages if page.text)
    total_pages = len(pages)

    logger.info(f"PDF has {total_pages} pages. Chunk size: {chunk_size}")

    # Step 2.5: Check for static reference grids first (backward compatibility)
    ref_groups = check_and_load_reference_grids(raw_pdf_text)
    if ref_groups:
        logger.info("Matched static reference grids. Skipping LLM extraction.")
        if progress_callback:
            progress_callback(95, 100, "Reference match found! Loading predefined sheets...")
        excel_bytes = build_excel_workbook(ref_groups)
        if progress_callback:
            progress_callback(100, 100, "Success!")
        return excel_bytes, ref_groups, []

    # Determine whether to use chunking
    use_chunking = chunk_size > 0 and total_pages > chunk_size
    if use_chunking:
        chunks = _chunk_pages(pages, chunk_size)
        logger.info(f"Split {total_pages} pages into {len(chunks)} chunks of up to {chunk_size} pages.")
    else:
        # Process all pages at once (original behavior)
        chunks = [pages]
        logger.info("Processing all pages in a single pass (no chunking).")

    max_retries = 3
    attempt = 1

    while attempt <= max_retries:
        logger.info(f"--- Extraction Attempt {attempt}/{max_retries} ---")

        all_page_results: List[Dict[str, Any]] = []
        chunk_context_history: List[str] = []

        for chunk_idx, chunk_pages in enumerate(chunks):
            chunk_start_page = chunk_pages[0].page_num
            chunk_end_page = chunk_pages[-1].page_num
            chunk_num = chunk_idx + 1
            total_chunks = len(chunks)

            logger.info(
                f"Processing chunk {chunk_num}/{total_chunks} "
                f"(pages {chunk_start_page}-{chunk_end_page}, {len(chunk_pages)} pages)..."
            )

            if progress_callback:
                page_progress = int((chunk_idx / total_chunks) * 80)
                progress_callback(
                    page_progress, 100,
                    f"Chunk {chunk_num}/{total_chunks}: Extracting pages {chunk_start_page}-{chunk_end_page}..."
                )

            # Build cross-chunk context from the previous chunk's summary
            chunk_previous_context = "\n".join(chunk_context_history) if chunk_context_history else ""

            # Process this chunk's pages sequentially (with intra-chunk context)
            chunk_results = process_all_pages_sequential(
                pages=chunk_pages,
                api_key=api_key,
                model=model,
                previous_context=chunk_previous_context,
                progress_callback=None,  # Suppress per-page callbacks; we report at chunk level
            )

            all_page_results.extend(chunk_results)

            # Build context summary from this chunk for the next chunk
            chunk_summary = _build_chunk_context_summary(chunk_results, chunk_start_page)
            chunk_context_history.append(chunk_summary)

            # Keep only the last chunk's context to avoid prompt bloat
            if len(chunk_context_history) > 1:
                older = chunk_context_history[:-1]
                compact = f"Previously processed {len(older)} chunk(s) covering sections: "
                compact += "; ".join(c for c in older)
                chunk_context_history = [compact] + chunk_context_history[-1:]

        if progress_callback:
            progress_callback(82, 100, "All chunks processed. Merging cross-chunk tables...")

        # Step 4: Compile and merge consecutive tables (including cross-chunk merges)
        merged_sections = compile_and_merge_sections(all_page_results)

        if progress_callback:
            progress_callback(88, 100, "Classifying data and generating worksheet names...")

        # Step 5: Classify sections into the exactly three target worksheets
        groups = classify_sections(merged_sections, api_key, model)

        # Step 6: Validate output structures and numbers
        if progress_callback:
            progress_callback(94, 100, "Performing pre-export validation audit...")
        is_valid, validation_msgs = validate_extraction(groups, raw_pdf_text)

        if is_valid:
            logger.info("Extraction validation passed successfully!")
            if progress_callback:
                progress_callback(98, 100, "Generating styled Excel workbook...")
            excel_bytes = build_excel_workbook(groups)
            if progress_callback:
                progress_callback(100, 100, "Success!")
            return excel_bytes, groups, validation_msgs
        else:
            logger.warning(f"Validation failed on attempt {attempt}: {validation_msgs}")
            attempt += 1
            if attempt > max_retries:
                logger.error("All extraction retries exhausted. Exporting best-effort data.")
                excel_bytes = build_excel_workbook(groups)
                return excel_bytes, groups, validation_msgs + [
                    "CRITICAL: Failed to validate data structure after 3 attempts. "
                    "Table schema may contain gaps."
                ]

            logger.info("Retrying extraction pipeline...")

    # Fallback return (should not be reached)
    raise RuntimeError("Extraction failed.")

