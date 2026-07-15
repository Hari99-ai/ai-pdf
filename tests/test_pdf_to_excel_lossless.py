from io import BytesIO

from openpyxl import load_workbook

from pdf_to_excel import analyze_excel, build_excel_from_sources, split_text_for_excel, validate_extraction


def test_split_text_for_excel_keeps_all_non_empty_lines():
    text = "line 1\n\nline 2\n  \nline 3"

    assert split_text_for_excel(text) == ["line 1", "line 2", "line 3"]


def test_build_excel_from_sources_includes_raw_text_sheet():
    raw_text = "Rates 100 200\nPolicy 50%"
    workbook_bytes = build_excel_from_sources(
        [
            ("raw_pdf_text", split_text_for_excel(raw_text)),
            ("debug", {"section_a": {"field": "value"}}),
        ]
    )

    wb = load_workbook(BytesIO(workbook_bytes))

    assert "raw_pdf_text" in wb.sheetnames
    raw_sheet = wb["raw_pdf_text"]
    assert raw_sheet["A1"].value == "value"
    assert raw_sheet["A2"].value == "Rates 100 200"
    assert raw_sheet["A3"].value == "Policy 50%"


def test_validate_extraction_treats_raw_text_sheet_as_lossless():
    raw_text = "Rates 100 200\nPolicy 50%"
    workbook_bytes = build_excel_from_sources(
        [
            ("raw_pdf_text", split_text_for_excel(raw_text)),
            ("debug", {"section_a": {"field": "value"}}),
        ]
    )

    excel_info = analyze_excel(workbook_bytes)
    pdf_info = {
        "total_pages": 1,
        "total_chars": len(raw_text),
        "total_words": len(raw_text.split()),
        "sections_found": [],
        "currency_values": ["$100"],
        "dates": [],
        "room_types": [],
        "lines": raw_text.splitlines(),
        "percentages": [],
        "emails": [],
        "phones": [],
    }

    report = validate_extraction(pdf_info, excel_info, {"section_a": {"field": "value"}})

    assert report["data_loss_pct"] == 0.0
    assert report["accuracy_pct"] == 100.0
    assert report["status"] == "PASS"
