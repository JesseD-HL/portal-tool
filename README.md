# Portal Tool

A Flask backend (`server.py`) + single-page React frontend (`portal.html`) that validates apparel prebuys/buysheets against SAP HANA and generates buysheets, creation templates, and purchase-requisition CSVs.

This is a sanitized version for demonstration: all company-specific SAP schema, view, table, and transaction-code names have been replaced with generic placeholders (`YOUR_COMPANY_SCHEMA`, `YOUR_ARTICLE_MASTER_VIEW`, `YOUR_TXN_1`, etc.). No credentials are stored in the code — connection settings come from environment variables or a local `config.json` (gitignored).

## Setup

1. Copy `config.example.json` to `config.json` and fill in your own HANA host, port, user, and password — or set the equivalent environment variables (`HANA_HOST`, `HANA_PORT`, `HANA_USER`, `HANA_PASSWORD`, `TOOLS_DIR`, `STOCK_VIEW`).
2. Replace the placeholder SAP schema/view/table/transaction names in `server.py` with your own HANA calculation-view package and object names.
3. Install dependencies: `pip install flask hdbcli openpyxl` (plus whatever `server.py`'s other imports require).
4. Run: `python server.py`
5. Open `http://localhost:5050` — it serves `portal.html` directly.

## What it does

- Validates a prebuy/buysheet upload against live SAP HANA data (article existence, prepack composition, pricing, color codes).
- Generates the buysheet workbook, GENERIC/PPK creation templates, and a validation report.
- Builds a purchase-requisition CSV (including prepack PAC quantities, not just open stock).
- Pulls live on-hand stock from a configurable HANA view, or accepts a Cut-and-Sold export.
- Includes a read-only SAP explorer for tracing which tables back a given transaction.
