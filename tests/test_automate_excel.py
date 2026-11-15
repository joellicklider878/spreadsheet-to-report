from pathlib import Path
import openpyxl
import pandas as pd
import pytest

from automate_excel import (
    MAX_ORDER_AMOUNT,
    REQUIRED_HEADERS,
    business_rules,
    create_sample_workbooks,
    generate_report,
    process_sales_data,
    validate_and_load_files,
)


def create_workbook(path: Path, sheet_name: str, rows: list[list]) -> None:
    """Helper to build test Excel workbooks."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in rows:
        ws.append(row)
    wb.save(path)


# ==============================================================================
# 1. SAMPLE WORKBOOK & SCHEMA VALIDATION TESTS
# ==============================================================================

def test_create_sample_workbooks(tmp_path: Path):
    create_sample_workbooks(tmp_path)

    expected_files = {"sales_january.xlsx", "sales_february.xlsx", "sales_march.xlsx"}
    created_files = {p.name for p in tmp_path.glob("sales_*.xlsx")}
    assert created_files == expected_files


def test_validate_and_load_valid_workbook(tmp_path: Path):
    file_path = tmp_path / "sales_january.xlsx"
    create_workbook(
        file_path,
        "Sales",
        [
            REQUIRED_HEADERS,
            ["ORD-001", "2026-01-10", "North", 2, 50.0],
            [None, None, None, None, None],  # Empty row should be skipped
            ["ORD-002", "2026-01-15", "South", 1, 25.0],
        ],
    )

    df, errors = validate_and_load_files(tmp_path)

    assert errors == []
    assert len(df) == 2
    assert list(df["order_id"]) == ["ORD-001", "ORD-002"]
    assert list(df["source_row"]) == [2, 4]
    assert all(df["source_file"] == "sales_january.xlsx")


def test_validate_missing_sales_worksheet(tmp_path: Path):
    file_path = tmp_path / "sales_wrong_sheet.xlsx"
    create_workbook(file_path, "Sheet1", [REQUIRED_HEADERS])

    df, errors = validate_and_load_files(tmp_path)

    assert df.empty
    assert len(errors) == 1
    assert errors[0]["source_file"] == "sales_wrong_sheet.xlsx"
    assert "Missing 'Sales' worksheet" in errors[0]["reason"]


@pytest.mark.parametrize(
    "bad_headers",
    [
        ["order_id", "order_date", "region", "quantity"],  # Missing header
        ["order_id", "region", "order_date", "quantity", "unit_price"],  # Wrong order
        ["order_id", "order_date", "region", "quantity", "unit_price", "extra"],  # Extra header
        ["order_id", "order_date", "region", "quantity", "quantity"],  # Duplicate header
    ],
)
def test_validate_invalid_headers(tmp_path: Path, bad_headers: list[str]):
    file_path = tmp_path / "sales_bad_headers.xlsx"
    create_workbook(file_path, "Sales", [bad_headers])

    df, errors = validate_and_load_files(tmp_path)

    assert df.empty
    assert len(errors) == 1
    assert errors[0]["source_row"] == 1
    assert "Invalid headers" in errors[0]["reason"]


# ==============================================================================
# 2. BUSINESS RULES TESTS
# ==============================================================================

def test_business_rules_valid_record():
    seen_ids: set[str] = set()
    violations = business_rules(
        order_id="ORD-100",
        region="North",
        amount=500.00,
        seen_order_ids=seen_ids,
    )

    assert violations == []
    assert "ORD-100" in seen_ids


def test_business_rules_disallowed_region():
    violations = business_rules(
        order_id="ORD-101",
        region="Central",
        amount=100.00,
        seen_order_ids=set(),
    )

    assert len(violations) == 1
    assert "Region 'Central' not in" in violations[0]


def test_business_rules_duplicate_order_id():
    seen_ids = {"ORD-102"}
    violations = business_rules(
        order_id="ORD-102",
        region="East",
        amount=100.00,
        seen_order_ids=seen_ids,
    )

    assert len(violations) == 1
    assert "Duplicate order_id 'ORD-102'" in violations[0]


def test_business_rules_exceeds_max_amount():
    over_limit = MAX_ORDER_AMOUNT + 0.01
    violations = business_rules(
        order_id="ORD-103",
        region="West",
        amount=over_limit,
        seen_order_ids=set(),
    )

    assert len(violations) == 1
    assert "exceeds" in violations[0]


# ==============================================================================
# 3. PROCESSING & REPORT GENERATION TESTS
# ==============================================================================

def test_process_sales_data_routing():
    raw_df = pd.DataFrame([
        {
            "order_id": "ORD-1",
            "order_date": "2026-01-10",
            "region": "North",
            "quantity": 4,
            "unit_price": 25.0,
            "source_file": "sales_january.xlsx",
            "source_row": 2,
        },
        {
            "order_id": "ORD-2",
            "order_date": "bad-date",
            "region": "Central",
            "quantity": -5,
            "unit_price": 10.0,
            "source_file": "sales_february.xlsx",
            "source_row": 3,
        },
    ])

    accepted_df, review_df, sales_value_df = process_sales_data(raw_df, file_errors=[])

    assert len(accepted_df) == 1
    assert accepted_df.iloc[0]["amount"] == 100.0

    assert len(review_df) == 1
    reason = review_df.iloc[0]["reason"]
    assert "Invalid order_date" in reason
    assert "Quantity must be > 0" in reason
    assert "Region 'Central' not in" in reason

    assert len(sales_value_df) == 2
    assert list(sales_value_df["status"]) == ["accepted", "rejected"]


def test_generate_report_sheets_and_totals(tmp_path: Path):
    accepted_df = pd.DataFrame([
        {
            "order_id": "ORD-1",
            "order_date": "2026-01-10",
            "region": "North",
            "quantity": 2,
            "unit_price": 50.0,
            "amount": 100.0,
            "source_file": "sales_january.xlsx",
            "source_row": 2,
        },
        {
            "order_id": "ORD-2",
            "order_date": "2026-02-05",
            "region": "South",
            "quantity": 3,
            "unit_price": 50.0,
            "amount": 150.0,
            "source_file": "sales_february.xlsx",
            "source_row": 2,
        },
    ])
    review_df = pd.DataFrame([
        {"source_file": "sales_march.xlsx", "source_row": 1, "reason": "Bad header"}
    ])
    sales_value_df = accepted_df.assign(status="accepted")
    output_file = tmp_path / "sales_report.xlsx"

    generate_report(accepted_df, review_df, sales_value_df, output_file)

    assert output_file.exists()
    wb = openpyxl.load_workbook(output_file)
    assert wb.sheetnames == ["summary", "accepted_data", "needs_review", "sales_value"]

    summary_ws = wb["summary"]
    rows = list(summary_ws.iter_rows(values_only=True))
    assert rows[0] == ("region", "record_count", "total_sales")
    assert rows[-1] == ("TOTAL", 2, 250.0)
