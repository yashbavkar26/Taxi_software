# Taxi Software - Modular Application & Excel-to-MySQL Pipeline

A modular, extensible software architecture for taxi dispatch, bookings management, duty slips, billing, and Excel-to-MySQL import pipelines.

---

## Directory & Package Architecture

The project is organized under the scalable `app/` layout:

```text
Taxi Software/
├── app/
│   ├── __init__.py
│   ├── core/                  # Core infrastructure
│   │   ├── config.py          # Environment settings loader (.env)
│   │   ├── database.py        # MySQL connection & atomic transaction manager
│   │   └── logging.py         # Centralized console/file logging
│   │
│   ├── common/                # Shared utilities across modules
│   │   ├── normalizer.py      # Type cleaners (dates, times, phones, integers, zeros)
│   │   └── schema_base.py     # Base TableSchema definitions
│   │
│   ├── modules/               # Domain feature modules
│   │   ├── drivers/           # Drivers & Contacts domain
│   │   │   ├── schema.py      # Drivers table schema & header aliases
│   │   │   └── repository.py  # Parameterized batch insert & duplicate handling
│   │   │
│   │   ├── bookings/          # Bookings domain
│   │   │   ├── schema.py      # Bookings table schema & header aliases
│   │   │   └── repository.py  # Parameterized batch insert into 'bookings'
│   │   │
│   │   ├── importers/         # Data ingestion pipelines
│   │   │   ├── excel_reader.py    # Low-memory streaming reader (openpyxl)
│   │   │   ├── excel_validator.py # Multi-sheet header & row validation
│   │   │   ├── models.py          # RowError, SheetSummary, ImportResult
│   │   │   └── service.py         # ExcelImportService orchestrator
│   │   │
│   │   ├── dutyslips/         # [Ready for Duty Slip Entry & KM Verification]
│   │   ├── billing/           # [Ready for Bill Book, Vouchers, GST Reports]
│   │   ├── messaging/         # [Ready for Automated Driver & Guest Texts]
│   │   └── vendors/           # Vendor Databook, Masterfile & Annexures
│   │       ├── schema.py      # Masterfile table schema & header aliases
│   │       ├── template.py    # Masterfile template generator (Excel)
│   │       └── vendor_generator.py # Live-syncing vendor workbook generator
│   │
│   └── cli/                   # Command-line entry points
│       ├── import_bookings.py       # CLI: python -m app.cli.import_bookings
│       └── generate_vendor_files.py # CLI: python -m app.cli.generate_vendor_files
│
├── tests/                     # Unit and integration test suite (84 tests)
├── requirements.txt
├── .env.example
├── taxi_schema.sql
├── bookings_template.xlsx     # Generated bookings workbook template
├── masterfile_template.xlsx   # Generated masterfile workbook template
├── vendor_template.xlsx       # Live-syncing vendor template with FILTER formula
├── Arvind.xlsx ... Yogesh.xlsx# 14 pre-generated vendor workbooks
└── README.md
```

---

## Key Excel Import Rules

1. **Multi-Sheet Mapping**:
   - **`Contacts` Sheet** (case-insensitive): Contains driver contacts (`Sr. no`, `Driver`, `Contact`) and is imported into the **`drivers`** table.
   - **All Other Sheets** (`Feb 01`, `Feb 02`, etc.): Contain date-wise taxi bookings and are imported into **ONE single MySQL table**: **`bookings`**.
   - Worksheet names do **not** determine destination table names.

2. **Automatic Driver Registration**:
   - Drivers from the `Contacts` worksheet are imported first.
   - When new drivers are encountered in `bookings`, the system checks if the driver already exists in the `drivers` table; if not, the driver is automatically inserted with their contact number. This can be toggled via `--no-auto-register-drivers` or `AUTO_REGISTER_DRIVERS=false`.

3. **Reversed Sheet Ingestion (Earliest Date First)**:
   - Workbooks arranged from end-of-month to start-of-month (e.g., Dec 31 to Dec 01) are ingested in reverse order so MySQL IDs reflect chronological progression (`REVERSE_SHEETS=true` by default). Can be toggled via `--no-reverse-sheets`.

4. **Company Support & Generated Stored Columns**:
   - MySQL column `company` defaults to `'GIM'` (customizable per booking row or Excel sheet).
   - MySQL columns `driver_message` and `guest_message` are `STORED GENERATED ALWAYS AS (...)`.
   - `guest_message` dynamically outputs `From [Company Name]\nDriver [Name] [Contact] will come to pick you tomorrow at [Time] at [Location]`.
   - The importer automatically detects and excludes trailing formula/message columns from the `INSERT` statement to avoid MySQL Error 3105.


5. **Streaming & Bounded Memory**:
   - Uses `openpyxl` with `read_only=True` and `data_only=True`.
   - Rows are streamed row-by-row with $O(1)$ memory usage.

6. **ACID Transaction Safety**:
   - Pre-validates schemas across all worksheets before starting any database operations.
   - If any sheet has missing columns, extra columns, or duplicate headers, the import aborts immediately.
   - Database writes run inside a single atomic transaction: commits on complete success; rolls back entirely on any error.

---

## Installation & Setup

### 1. Requirements

- Python 3.10+
- MySQL 8.0+

Install dependencies:

```bash
pip install -r requirements.txt
```

### 2. Configuration

Copy `.env.example` to `.env` and fill in your MySQL credentials:

```ini
# MySQL Database Configuration
DB_HOST=localhost
DB_PORT=3306
DB_USER=root
DB_PASSWORD=your_password
DB_NAME=taxi_database
DB_CHARSET=utf8mb4

# Importer Configuration
BATCH_SIZE=1000
DUPLICATE_MODE=skip          # 'skip', 'upsert', or 'fail'
AUTO_REGISTER_DRIVERS=true   # Auto-insert unknown driver names into drivers table
LOG_LEVEL=INFO
```

---

## Usage

### Command Line Interface (CLI)

#### Dry-Run Mode (Validate without database writes)
```bash
python -m app.cli.import_bookings "Requirements/Input/Bookings Sample.xlsx" --dry-run
```

#### Live Database Import
```bash
python -m app.cli.import_bookings "Requirements/Input/Bookings Sample.xlsx"
```

#### CLI Options
```bash
python -m app.cli.import_bookings <file.xlsx> [OPTIONS]

Options:
  --dry-run                    Validate workbook and rows without making database writes
  --batch-size INT             Number of rows per batch insert (default: 1000)
  --duplicate-mode {skip,upsert,fail}
                               Duplicate strategy for unique keys (default: skip)
  --no-auto-register-drivers   Do not auto-register unknown driver names
  --log-level {DEBUG,INFO,WARNING,ERROR}
                               Set console log verbosity
```

### Programmatic Python API

```python
from app import Config, ExcelImportService

# Initialize configuration
config = Config.from_env()

# Create importer service
service = ExcelImportService(config=config)

# Run dry-run or live import
result = service.import_workbook("Requirements/Input/Bookings Sample.xlsx", dry_run=False)

print(result.summary_text())
if result.status == "SUCCESS":
    print(f"Successfully imported {result.rows_inserted} rows!")
else:
    print(f"Import failed: {result.message}")
```

---

## Live Vendor Synchronization System (Excel Native)

Each vendor's trip records are dynamically synced from `masterfile_template.xlsx` using native Excel dynamic array formulas:

### 1. How It Works
- **Zero Scripts at Runtime**: In Microsoft Excel, opening `masterfile_template.xlsx` alongside any vendor workbook (e.g. `Arvind.xlsx`) updates trips and totals **in real time**.
- **Dynamic Array Formula (Cell A7)**:
  ```excel
  =IF(OR(ISBLANK($B$2),$B$2=""),"Please select a vendor in cell B2",FILTER('[masterfile_template.xlsx]Masterfile'!$A$2:$T$5000,'[masterfile_template.xlsx]Masterfile'!$B$2:$B$5000=$B$2,"No trips recorded for "&$B$2))
  ```
- **Live Summary KPI Cards (Rows 3–4)**:
  - **TOTAL TRIPS**: `=COUNTIF('[masterfile_template.xlsx]Masterfile'!$B$2:$B$5000, $B$2)`
  - **TOTAL DISTANCE**: `=SUMIFS('[masterfile_template.xlsx]Masterfile'!$H$2:$H$5000, '[masterfile_template.xlsx]Masterfile'!$B$2:$B$5000, $B$2)`
  - **TOTAL AMOUNT**: `=SUMIFS('[masterfile_template.xlsx]Masterfile'!$N$2:$N$5000, '[masterfile_template.xlsx]Masterfile'!$B$2:$B$5000, $B$2)`
  - **PAYMENT PENDING**: `=COUNTIFS('[masterfile_template.xlsx]Masterfile'!$B$2:$B$5000, $B$2, '[masterfile_template.xlsx]Masterfile'!$O$2:$O$5000, FALSE)`

### 2. Adding a New Vendor (Zero Code)
1. Copy `vendor_template.xlsx` in the root directory.
2. Rename the file to `<VendorName>.xlsx` (e.g. `Rahul.xlsx`).
3. Open the file and select or type the vendor name in cell **B2**. All records spill instantly.

### 3. Generating Vendor Files via CLI
Generate all 14 active vendor files and `vendor_template.xlsx`:
```bash
python -m app.cli.generate_vendor_files
```

Generate a single vendor file:
```bash
python -m app.cli.generate_vendor_files --vendor Arvind
```

---

## Running Automated Tests

Run the complete test suite:

```bash
pytest -v
```

All 84 test cases pass in ~3s.
