from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd

# 1. CONFIGURATION & CONSTANTS
REQUIRED_HEADERS = ["order_id", "order_date", "region", "quantity", "unit_price"]
ALLOWED_REGIONS = {"North", "South", "East", "West"}
MAX_ORDER_AMOUNT = 10000.00


# 2. SAMPLE INPUT GENERATION
def create_sample_workbooks(base_dir: Path) -> None:
    """Create sales_january.xlsx, sales_february.xlsx, and sales_march.xlsx if missing."""
    if list(base_dir.glob("sales_*.xlsx")):
        return

    samples = {
        "sales_january.xlsx": [
            REQUIRED_HEADERS,
            ["ORD-101", "2026-01-10", "North", 5, 20.00],
            ["ORD-102", "2026-01-15", "South", 10, 15.50],
            [None, None, None, None, None],  # Completely empty row (skipped)
        ],
        "sales_february.xlsx": [
            REQUIRED_HEADERS,
            ["ORD-201", "2026-02-05", "East", 8, 50.00],
            ["ORD-202", "2026-02-12", "West", -3, 25.00],       # Invalid quantity
            ["ORD-203", "2026-02-20", "Central", 4, 30.00],     # Disallowed region
        ],
        "sales_march.xlsx": [
            REQUIRED_HEADERS,
            ["ORD-301", "2026-03-03", "North", 12, 40.00],
            ["ORD-101", "2026-03-11", "South", 2, 20.00],       # Duplicate order_id
            ["ORD-302", "2026-03-25", "West", 500, 30.00],      # Exceeds max amount
        ],
    }

    for filename, rows in samples.items():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sales"
        for row in rows:
            ws.append(row)
        wb.save(base_dir / filename)


# 3. SCHEMA VALIDATION & DATA LOADING
def validate_and_load_files(base_dir: Path) -> tuple[pd.DataFrame, list[dict]]:
    """Inspect headers with openpyxl before loading non-empty rows into a DataFrame."""
    records = []
    file_errors = []

    for file_path in sorted(base_dir.glob("sales_*.xlsx")):
        wb = openpyxl.load_workbook(file_path, data_only=True)

        # Check for required 'Sales' worksheet
        if "Sales" not in wb.sheetnames:
            file_errors.append({
                "source_file": file_path.name,
                "source_row": "N/A",
                "reason": "Missing 'Sales' worksheet",
            })
            continue

        ws = wb["Sales"]

        # Read row 1 headers safely and trim trailing blank cells
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        raw_headers = list(first_row)
        while raw_headers and raw_headers[-1] is None:
            raw_headers.pop()

        # Enforce exact 5 headers, strict order, and no duplicates or extra columns
        if (
            len(raw_headers) != len(REQUIRED_HEADERS)
            or len(set(raw_headers)) != len(REQUIRED_HEADERS)
            or raw_headers != REQUIRED_HEADERS
        ):
            file_errors.append({
                "source_file": file_path.name,
                "source_row": 1,
                "reason": f"Invalid headers: found {raw_headers}, expected {REQUIRED_HEADERS}",
            })
            continue

        # Load data rows starting at Excel row 2, skipping completely blank rows
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            row_values = list(row[: len(REQUIRED_HEADERS)])
            if all(v is None or str(v).strip() == "" for v in row_values):
                continue

            record = dict(zip(REQUIRED_HEADERS, row_values))
            record["source_file"] = file_path.name
            record["source_row"] = row_idx
            records.append(record)

    return pd.DataFrame(records), file_errors


# 4. BUSINESS RULES & RECORD PROCESSING
def business_rules(
    order_id: str,
    region: str,
    amount: float | None,
    seen_order_ids: set[str],
) -> list[str]:
    """Enforce territory, uniqueness, and order threshold rules."""
    violations = []

    if region and region not in ALLOWED_REGIONS:
        violations.append(f"Region '{region}' not in {sorted(ALLOWED_REGIONS)}")

    if order_id:
        if order_id in seen_order_ids:
            violations.append(f"Duplicate order_id '{order_id}'")
        else:
            seen_order_ids.add(order_id)

    if amount is not None and amount > MAX_ORDER_AMOUNT:
        violations.append(f"Amount ${amount:,.2f} exceeds ${MAX_ORDER_AMOUNT:,.2f} limit")

    return violations


def process_sales_data(
    df: pd.DataFrame,
    file_errors: list[dict],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Validate field values, calculate order amounts, and route rows to output tables."""
    accepted_columns = [*REQUIRED_HEADERS, "amount", "source_file", "source_row"]
    review_columns = [*accepted_columns, "reason"]
    sales_value_columns = [*accepted_columns, "status"]

    if df.empty:
        return (
            pd.DataFrame(columns=accepted_columns),
            pd.DataFrame(file_errors, columns=review_columns),
            pd.DataFrame(columns=sales_value_columns),
        )

    accepted_rows = []
    review_rows = list(file_errors)
    sales_value_rows = []
    seen_order_ids: set[str] = set()

    for _, row in df.iterrows():
        reasons = []

        order_id = str(row["order_id"]).strip() if pd.notna(row["order_id"]) else ""
        region = str(row["region"]).strip() if pd.notna(row["region"]) else ""

        if not order_id:
            reasons.append("Missing order_id")
        if not region:
            reasons.append("Missing region")

        parsed_date = pd.to_datetime(row["order_date"], errors="coerce")
        if pd.isna(parsed_date):
            reasons.append("Invalid order_date")

        qty = pd.to_numeric(row["quantity"], errors="coerce")
        price = pd.to_numeric(row["unit_price"], errors="coerce")

        if pd.isna(qty) or qty <= 0:
            reasons.append("Quantity must be > 0")
        if pd.isna(price) or price < 0:
            reasons.append("Unit price must be >= 0")

        # Calculate order total
        amount = round(float(qty * price), 2) if (pd.notna(qty) and pd.notna(price)) else None

        reasons.extend(business_rules(order_id, region, amount, seen_order_ids))

        record = {
            "order_id": order_id or row["order_id"],
            "order_date": parsed_date.strftime("%Y-%m-%d") if pd.notna(parsed_date) else row["order_date"],
            "region": region or row["region"],
            "quantity": qty if pd.notna(qty) else row["quantity"],
            "unit_price": price if pd.notna(price) else row["unit_price"],
            "amount": amount,
            "source_file": row["source_file"],
            "source_row": row["source_row"],
        }

        # Record complete value of every incoming order
        sales_value_rows.append({
            **record,
            "status": "rejected" if reasons else "accepted",
        })

        if reasons:
            review_rows.append({**record, "reason": "; ".join(reasons)})
        else:
            accepted_rows.append(record)

    return (
        pd.DataFrame(accepted_rows, columns=accepted_columns),
        pd.DataFrame(review_rows, columns=review_columns),
        pd.DataFrame(sales_value_rows, columns=sales_value_columns),
    )


# 5. REPORT GENERATION
def generate_report(
    accepted_df: pd.DataFrame,
    review_df: pd.DataFrame,
    sales_value_df: pd.DataFrame,
    output_path: Path,
) -> None:
    """Export summary, accepted_data, needs_review, and sales_value sheets to Excel."""
    if not accepted_df.empty:
        summary_df = (
            accepted_df.groupby("region", as_index=False)
            .agg(record_count=("order_id", "count"), total_sales=("amount", "sum"))
            .sort_values("region")
        )
        grand_total = pd.DataFrame([{
            "region": "TOTAL",
            "record_count": int(summary_df["record_count"].sum()),
            "total_sales": round(float(summary_df["total_sales"].sum()), 2),
        }])
        summary_df = pd.concat([summary_df, grand_total], ignore_index=True)
    else:
        summary_df = pd.DataFrame(columns=["region", "record_count", "total_sales"])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="summary", index=False)
        accepted_df.to_excel(writer, sheet_name="accepted_data", index=False)
        review_df.to_excel(writer, sheet_name="needs_review", index=False)
        sales_value_df.to_excel(writer, sheet_name="sales_value", index=False)

        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        header_font = Font(bold=True)

        for ws in writer.book.worksheets:
            header_cells = next(ws.iter_rows(min_row=1, max_row=1), ())
            for cell in header_cells:
                cell.fill = header_fill
                cell.font = header_font

            for col_idx, cell in enumerate(header_cells, start=1):
                if cell.value in {"unit_price", "amount", "total_sales"}:
                    for row_idx in range(2, ws.max_row + 1):
                        ws.cell(row=row_idx, column=col_idx).number_format = "$#,##0.00"


# 6. ENTRY POINT
def main() -> None:
    base_dir = Path(".")
    create_sample_workbooks(base_dir)
    raw_df, file_errors = validate_and_load_files(base_dir)
    accepted_df, review_df, sales_value_df = process_sales_data(raw_df, file_errors)
    generate_report(accepted_df, review_df, sales_value_df, base_dir / "sales_report.xlsx")
    print("Report generated: sales_report.xlsx")


if __name__ == "__main__":
    main()
