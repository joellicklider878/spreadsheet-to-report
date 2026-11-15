# Spreadsheet to Report

A Python automation tool that validates incoming regional sales workbooks (`sales_*.xlsx`), enforces business rules, and compiles a formatted multi-sheet summary report (`sales_report.xlsx`).

## Features

- **Strict Schema Validation:** Inspects workbooks using `openpyxl` before loading data to verify the `Sales` worksheet exists and matches the exact required column order (`order_id`, `order_date`, `region`, `quantity`, `unit_price`).
- **Source Traceability:** Skips blank rows while tagging every record with its originating `source_file` and `source_row` number.
- **Business Rule Enforcement:**
  - Restricts `region` values to approved territories (`North`, `South`, `East`, `West`).
  - Detects duplicate `order_id` entries across all scanned workbooks.
  - Flags orders exceeding `$10,000.00` for manual review.
- **Formatted Excel Output:** Generates `sales_report.xlsx` with bolded headers, currency formatting (`$#,##0.00`), and four structured tabs:
  - `summary` — Accepted order counts and revenue totals grouped by region, plus a grand total row.
  - `accepted_data` — Validated orders with calculated `amount` values.
  - `needs_review` — File-level schema errors and rejected records with detailed failure reasons.
  - `sales_value` — Complete audit log of all incoming orders and their validation status.

## Project Structure

```text
spreadsheet-to-report/
├── pyproject.toml
├── README.md
├── src/
│   └── automate_excel.py
└── tests/
    └── test_automate_excel.py
```

## Prerequisites

- Python **3.11** or higher

## Installation

1. Clone the repository and navigate into the project directory:
   ```bash
   git clone <your-repo-url>
   cd spreadsheet-to-report
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. Install the package and development dependencies:
   ```bash
   pip install -e ".[dev]"
   ```

## Usage

Place any `sales_*.xlsx` workbooks in your working directory and run the CLI command:

```bash
spreadsheet-to-report
```

*(Alternatively, run the script directly with `python src/automate_excel.py`.)*

If no `sales_*.xlsx` files exist in the directory, the tool automatically generates sample workbooks (`sales_north.xlsx` and `sales_south.xlsx`) demonstrating valid records, schema checks, and business rule violations before outputting `sales_report.xlsx`.

## Running Tests & Linting

Run the `pytest` test suite:
```bash
pytest -v
```

Check code formatting and linting with `ruff`:
```bash
ruff check .
```

## License

MIT

