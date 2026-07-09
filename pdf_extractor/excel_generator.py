import io
import re
from typing import Dict, Any, List, Union
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from pdf_extractor.logging_config import logger

def excel_safe_value(value: Any) -> Union[str, int, float, bool]:
    """Convert any nested JSON values into a clean string representation for Excel cells."""
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, dict):
        parts = []
        for k, v in value.items():
            val_str = str(excel_safe_value(v))
            if not val_str:
                continue
            parts.append(f"{k.replace('_', ' ').title()}: {val_str}")
        return "; ".join(parts)
    if isinstance(value, list):
        if not value:
            return ""
        # If list of dicts, format them nicely
        return ", ".join(str(excel_safe_value(item)) for item in value if item is not None)
    return str(value)

def pretty_header_name(name: str) -> str:
    """Format snake_case header name to Title Case with spaces."""
    # Special acronyms or cases
    special_cases = {
        "id": "ID",
        "usd": "USD",
        "ocr": "OCR",
        "pdf": "PDF",
        "fit": "FIT"
    }
    words = name.replace("_", " ").split()
    formatted_words = [special_cases.get(w.lower(), w.capitalize()) for w in words]
    return " ".join(formatted_words)

def union_align_sections(sections: List[Dict[str, Any]]) -> Tuple[List[str], List[Dict[str, Any]]]:
    """
    Union-align multiple sections into a single list of rows with a consistent schema.
    Converts list and key_value types to tables.
    """
    all_rows: List[Dict[str, Any]] = []
    seen_headers: List[str] = []
    
    for sec in sections:
        sec_type = sec.get("type", "table")
        sec_name = sec.get("name", "data")
        raw_rows = sec.get("rows", [])
        
        # 1. Convert key_value sections to tables (Attribute, Value)
        if sec_type == "key_value":
            if isinstance(raw_rows, dict):
                table_rows = []
                for k, v in raw_rows.items():
                    table_rows.append({
                        "attribute": pretty_header_name(k),
                        "value": excel_safe_value(v)
                    })
                raw_rows = table_rows
            elif isinstance(raw_rows, list):
                # If already list of key-values
                pass
                
        # 2. Convert list sections to tables (Description)
        elif sec_type == "list":
            if isinstance(raw_rows, list):
                raw_rows = [{"description": excel_safe_value(item)} for item in raw_rows]
                
        if not isinstance(raw_rows, list):
            logger.warning(f"Section {sec_name} rows is not a list. Skipping.")
            continue
            
        # Collect headers in order of appearance
        for row in raw_rows:
            if not isinstance(row, dict):
                continue
            for k in row.keys():
                if k not in seen_headers:
                    seen_headers.append(k)
            all_rows.append(row)
            
    # Sort headers so that common keywords appear first
    priority_keywords = ["date", "room", "season", "service", "description", "item", "charge", "price", "rate", "policy"]
    def header_sort_key(header: str) -> int:
        header_lower = header.lower()
        for idx, kw in enumerate(priority_keywords):
            if kw in header_lower:
                return idx
        return len(priority_keywords)
        
    seen_headers.sort(key=header_sort_key)
    
    return seen_headers, all_rows

def build_excel_workbook(groups: Dict[str, Dict[str, Any]]) -> bytes:
    """Build the single Excel workbook with exactly 3 worksheets and custom styles."""
    logger.info("Initializing openpyxl Workbook...")
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    # We must process exactly group_1, group_2, group_3
    keys = ["group_1", "group_2", "group_3"]
    
    for key in keys:
        group_data = groups.get(key, {})
        sheet_title = group_data.get("sheet_title", pretty_header_name(key))
        sections = group_data.get("sections", [])
        
        ws = wb.create_sheet(title=sheet_title)
        
        # Ensure grid lines are visible
        ws.views.sheetView[0].showGridLines = True
        
        # Align sections into a single relational table
        headers, rows = union_align_sections(sections)
        
        # If no data found, insert placeholder
        if not rows:
            headers = ["status"]
            rows = [{"status": "No data detected for this category."}]
            logger.info(f"No data for sheet '{sheet_title}', creating placeholder row.")
            
        # Write headers
        ws.append([pretty_header_name(h) for h in headers])
        
        # Write rows
        for row in rows:
            ws.append([excel_safe_value(row.get(h, "")) for h in headers])
            
        # Freeze top row
        ws.freeze_panes = "A2"
        
        # Style cells
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        border_thin = Side(border_style="thin", color="D3D3D3")
        cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
        
        # Center alignment for header
        align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
        # Left alignment with wrap text for data cells
        align_data_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        # Center alignment for specific columns (dates, rates, statuses, codes)
        align_data_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Apply header formatting
        ws.row_dimensions[1].height = 26
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_header
            cell.border = cell_border
            
        # Apply data formatting
        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 20
            # Zebra striping (alternating white and light blue)
            row_fill = PatternFill(
                start_color="F2F6FA" if row_idx % 2 == 0 else "FFFFFF",
                end_color="F2F6FA" if row_idx % 2 == 0 else "FFFFFF",
                fill_type="solid"
            )
            
            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(name="Segoe UI", size=10)
                cell.fill = row_fill
                cell.border = cell_border
                
                # Check header name to choose alignment
                h_lower = header.lower()
                if any(kw in h_lower for kw in ["date", "rate", "price", "status", "id", "percent"]):
                    cell.alignment = align_data_center
                else:
                    cell.alignment = align_data_left
                    
        # Apply Excel Table format
        try:
            # Table name must be alphanumeric and start with a letter, no spaces
            clean_name = re.sub(r"[^a-zA-Z0-9]", "", sheet_title)
            if not clean_name or not clean_name[0].isalpha():
                clean_name = f"Table{clean_name}"
            # Ensure name is unique inside workbook (append sheet index if needed)
            clean_name = clean_name[:250]
            
            table_ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
            excel_table = Table(displayName=clean_name, ref=table_ref)
            
            # Medium style with row stripes
            table_style = TableStyleInfo(
                name="TableStyleMedium9",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False
            )
            excel_table.tableStyleInfo = table_style
            ws.add_table(excel_table)
            logger.info(f"Added Excel Table '{clean_name}' for range {table_ref}.")
        except Exception as e:
            logger.warning(f"Could not add native Excel table styling: {e}. Falling back to cell styles.")
            
        # Enable Auto Filter
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
        
        # Auto column width
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                val_str = str(cell.value or "")
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)
            
    logger.info("Workbook generation finished.")
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
