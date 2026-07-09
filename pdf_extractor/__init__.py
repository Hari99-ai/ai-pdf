import os
from typing import Dict, Any, List, Tuple, Callable, Optional
from pdf_extractor.logging_config import logger
from pdf_extractor.pdf_parser import parse_pdf, check_pdf_validity, PDFPage
from pdf_extractor.table_detector import process_all_pages_parallel, compile_and_merge_sections
from pdf_extractor.classifier import classify_sections
from pdf_extractor.excel_generator import build_excel_workbook
from pdf_extractor.validation import validate_extraction
from pdf_extractor.reference_grids import check_and_load_reference_grids

def process_pdf(
    pdf_bytes: bytes,
    api_key: str,
    model: str = "openai/gpt-4o-mini",
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    max_workers: int = 5
) -> Tuple[bytes, Dict[str, Any], List[str]]:
    """
    Orchestrate the entire PDF-to-Excel workflow with self-healing retries.
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
    
    max_retries = 3
    attempt = 1
    
    while attempt <= max_retries:
        logger.info(f"--- Extraction Attempt {attempt}/{max_retries} ---")
        
        if progress_callback:
            progress_callback(0, 100, f"Attempt {attempt}/{max_retries}: Launching parallel page analysis...")

        # Step 3: Run page-by-page structured extraction (handles OCR automatically for scanned pages)
        page_results = process_all_pages_parallel(
            pages=pages,
            api_key=api_key,
            model=model,
            progress_callback=progress_callback,
            max_workers=max_workers
        )
        
        # Step 4: Compile and merge consecutive tables (multi-page table merging)
        if progress_callback:
            progress_callback(85, 100, "Merging consecutive tables and removing repeated headers...")
        merged_sections = compile_and_merge_sections(page_results)
        
        # Step 5: Classify sections into the exactly three target worksheets
        if progress_callback:
            progress_callback(90, 100, "Classifying data and generating worksheet names...")
        groups = classify_sections(merged_sections, api_key, model)
        
        # Step 6: Validate output structures and numbers
        if progress_callback:
            progress_callback(95, 100, "Performing pre-export validation audit...")
        is_valid, validation_msgs = validate_extraction(groups, raw_pdf_text)
        
        if is_valid:
            logger.info("Extraction validation passed successfully!")
            # Step 7: Generate formatted Excel Workbook
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
                # If we've run out of retries, build the Excel workbook anyway but include the warnings
                logger.error("All extraction retries exhausted. Exporting best-effort data.")
                excel_bytes = build_excel_workbook(groups)
                return excel_bytes, groups, validation_msgs + ["CRITICAL: Failed to validate data structure after 3 attempts. Table schema may contain gaps."]
                
            logger.info("Retrying extraction pipeline...")
            
    # Fallback return (should not be reached)
    raise RuntimeError("Extraction failed.")

