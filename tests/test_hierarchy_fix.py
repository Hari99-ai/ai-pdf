"""Quick smoke test for hierarchy fixes."""
import json
from pdf_extractor.table_detector import (
    _repair_truncated_json,
    merge_consecutive_tables,
    clean_repeated_headers,
    _rows_are_hierarchy_compatible,
)
from pdf_extractor.excel_generator import union_align_sections


def test_json_repair_truncated_string():
    broken = '{"sections": [{"name": "rates", "type": "table", "headers": ["room", "rate"], "rows": [{"room": "Suite", "rate": "200"}, {"room": "Villa", "rate":'
    repaired = _repair_truncated_json(broken)
    parsed = json.loads(repaired)
    assert len(parsed["sections"]) == 1
    assert len(parsed["sections"][0]["rows"]) >= 1
    print("PASS: test_json_repair_truncated_string")


def test_json_repair_trailing_comma():
    broken = '{"sections": [{"name": "x", "type": "table", "headers": ["a"], "rows": [{"a": "1"},]},'
    repaired = _repair_truncated_json(broken)
    parsed = json.loads(repaired)
    assert len(parsed["sections"]) == 1
    print("PASS: test_json_repair_trailing_comma")


def test_hierarchy_rejects_different_seasons():
    sec1 = {
        "name": "rates", "type": "table",
        "headers": ["season", "room", "rate"],
        "rows": [
            {"season": "High", "room": "Suite", "rate": "200"},
            {"season": "High", "room": "Villa", "rate": "350"},
        ]
    }
    sec2 = {
        "name": "rates", "type": "table",
        "headers": ["season", "room", "rate"],
        "rows": [
            {"season": "Low", "room": "Suite", "rate": "150"},
        ]
    }
    result = merge_consecutive_tables([sec1, sec2])
    assert len(result) == 2, f"Expected 2 sections, got {len(result)}"
    print("PASS: test_hierarchy_rejects_different_seasons")


def test_hierarchy_accepts_same_season():
    sec1 = {
        "name": "rates", "type": "table",
        "headers": ["season", "room", "rate"],
        "rows": [
            {"season": "High", "room": "Suite", "rate": "200"},
        ]
    }
    sec2 = {
        "name": "rates", "type": "table",
        "headers": ["season", "room", "rate"],
        "rows": [
            {"season": "High", "room": "Villa", "rate": "350"},
        ]
    }
    result = merge_consecutive_tables([sec1, sec2])
    assert len(result) == 1, f"Expected 1 section, got {len(result)}"
    print("PASS: test_hierarchy_accepts_same_season")


def test_hierarchy_compatible_check():
    prev_rows = [{"season": "High", "room": "Suite"}]
    curr_rows = [{"season": "Low", "room": "Villa"}]
    assert not _rows_are_hierarchy_compatible(prev_rows, curr_rows)
    curr_rows2 = [{"season": "High", "room": "Deluxe"}]
    assert _rows_are_hierarchy_compatible(prev_rows, curr_rows2)
    print("PASS: test_hierarchy_compatible_check")


def test_union_align_puts_hierarchy_first():
    sections = [
        {"name": "rates", "type": "table",
         "headers": ["room", "rate", "season", "meal_plan", "date_from"],
         "rows": [{"room": "Suite", "rate": "200", "season": "High", "meal_plan": "BB", "date_from": "Jan 01"}]},
    ]
    headers, _ = union_align_sections(sections)
    # Hierarchy fields should come first
    assert headers[0] == "season", f"Expected 'season' first, got '{headers[0]}'"
    assert "date_from" in headers[:5], f"Expected 'date_from' in first 5, got {headers[:5]}"
    assert "meal_plan" in headers[:5], f"Expected 'meal_plan' in first 5, got {headers[:5]}"
    print("PASS: test_union_align_puts_hierarchy_first")


if __name__ == "__main__":
    test_json_repair_truncated_string()
    test_json_repair_trailing_comma()
    test_hierarchy_rejects_different_seasons()
    test_hierarchy_accepts_same_season()
    test_hierarchy_compatible_check()
    test_union_align_puts_hierarchy_first()
    print("\nAll tests passed!")
