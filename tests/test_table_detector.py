from pdf_extractor.table_detector import clean_repeated_headers, merge_consecutive_tables


def test_clean_repeated_headers_only_removes_exact_header_rows():
    headers = ["dates", "cliffside ocean front", "romance deluxe ocean front"]

    exact_repeat = {
        "a": "DATES",
        "b": "Cliffside Ocean Front",
        "c": "Romance Deluxe Ocean Front",
    }
    subset_like_row = {
        "a": "DATES",
        "b": "Cliffside Ocean Front",
        "c": "320",
    }

    cleaned = clean_repeated_headers([exact_repeat, subset_like_row], headers)

    assert subset_like_row in cleaned
    assert exact_repeat not in cleaned


def test_merge_consecutive_tables_preserves_row_values_by_header_name():
    sections = [
        {
            "name": "rates_page_1",
            "type": "table",
            "headers": ["dates", "room_a", "room_b"],
            "rows": [
                {"dates": "Jan 1", "room_a": "$100", "room_b": "$200"},
                {"dates": "Jan 2", "room_a": "$110", "room_b": "$210"},
            ],
        },
        {
            "name": "rates_page_2",
            "type": "table",
            "headers": ["room_b", "dates", "room_a"],
            "rows": [
                {"room_b": "$220", "dates": "Jan 3", "room_a": "$120"},
            ],
        },
    ]

    merged = merge_consecutive_tables(sections)

    assert len(merged) == 1
    merged_rows = merged[0]["rows"]
    assert merged_rows[0]["room_a"] == "$100"
    assert merged_rows[1]["room_b"] == "$210"
    assert merged_rows[2]["dates"] == "Jan 3"
    assert merged_rows[2]["room_a"] == "$120"
    assert merged_rows[2]["room_b"] == "$220"


def test_merge_consecutive_tables_keeps_duplicate_rows():
    sections = [
        {
            "name": "policy_page_1",
            "type": "table",
            "headers": ["rule", "value"],
            "rows": [
                {"rule": "Deposit", "value": "100%"},
            ],
        },
        {
            "name": "policy_page_2",
            "type": "table",
            "headers": ["rule", "value"],
            "rows": [
                {"rule": "Deposit", "value": "100%"},
            ],
        },
    ]

    merged = merge_consecutive_tables(sections)

    assert len(merged) == 1
    assert len(merged[0]["rows"]) == 2
