# ⚡ Blocked VILL Meter Classifier (`bb-automation`)

An enterprise-grade automation tool for matching and classifying blocked VILL (Village / Rural) electricity meters against WC, LTCT, and HTCT consumer installation master records.

Built with Python and Streamlit, it replaces manual, error-prone spreadsheet lookups with high-performance set-based matching, duplicate detection, and automated multi-sheet branded Excel reporting.

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Key Features](#-key-features)
- [Technology Stack](#-technology-stack)
- [Repository Structure](#-repository-structure)
- [Prerequisites](#-prerequisites)
- [Installation & Setup](#-installation--setup)
- [Configuration](#-configuration)
- [Execution](#-execution)
- [How the Engine Works](#-how-the-engine-works)
- [Excel Report Breakdown](#-excel-report-breakdown)
- [Testing & Validation](#-testing--validation)
- [Troubleshooting & FAQ](#-troubleshooting--faq)
- [Security & Data Privacy](#-security--data-privacy)

---

## 🔍 Overview

Utility billing automation often faces challenges when identifying blocked meters against multiple disparate consumer categories:
- **WC** (Water Works / Commercial / Wholesale Installation Records)
- **LTCT** (Low Tension Current Transformer Installations)
- **HTCT** (High Tension Current Transformer Installations)

Traditional Excel `VLOOKUP` processes across multi-hundred-thousand-row sheets suffer from application freezing, accidental truncation of leading zeroes, and scientific notation corruption.

**Blocked VILL Meter Classifier** automates this end-to-end:
1. Reads all files as raw text (preventing leading zero truncation).
2. Sanitises and normalises identifiers (stripping whitespace, collapsing internal spaces, uppercase conversion).
3. Executes exact-match set lookups in seconds.
4. Generates an executive summary and a 7-sheet styled Excel workbook ready for auditing.

---

## ✨ Key Features

- **Leading Zero & Format Preservation**: Reads all columns as string/text (`dtype=str`) across `.xlsx`, `.xls`, and `.csv` files so numeric IDs are never altered.
- **Robust Identifier Normalisation**: Strips whitespace, standardises letter casing, and removes null/NaN representations before comparison.
- **Fast In-Memory Set Lookups**: O(1) membership queries process hundreds of thousands of records in seconds.
- **Internal Duplicate Detection**: Detects and alerts on duplicated identifiers within individual installation files.
- **Cross-Category Conflict Detection**: Flags meters that match more than one installation category (e.g., both WC and LTCT).
- **Formatted Multi-Sheet Export**: Generates an Excel report with color-coded status badges, frozen headers, and auto-filters using `openpyxl`.
- **Modern Dark UI**: Streamlit web interface configured with glassmorphic cards, responsive controls, and progress indicators.

---

## 🛠 Technology Stack

- **Language**: Python 3.9+ (tested on Python 3.11)
- **Web Framework**: [Streamlit](https://streamlit.io/) (>=1.30.0)
- **Data Processing**: [Pandas](https://pandas.pydata.org/) (>=2.0.0)
- **Spreadsheet Engines**:
  - [OpenPyXL](https://openpyxl.readthedocs.io/) (>=3.1.0) for modern `.xlsx` generation and styling
  - [XLRD](https://xlrd.readthedocs.io/) (>=2.0.0) for legacy `.xls` format support

---

## 📂 Repository Structure

```text
bb-automation/
├── .streamlit/
│   └── config.toml         # Streamlit dark theme & upload size limits (3000MB)
├── .env.example            # Environment variable template
├── .gitignore              # Git ignore rules for Python, OS, & secrets
├── app.py                  # Streamlit web application & UI workflow
├── engine.py               # Core data processing & Excel export engine
├── README.md               # Project documentation
└── requirements.txt        # Production dependency manifest
```

---

## ⚙️ Prerequisites

- **Python**: Version 3.9 or higher (recommended: Python 3.10 or 3.11)
- **Package Manager**: `pip` (Python package installer)
- **Git**: Version 2.25+

---

## 🚀 Installation & Setup

### 1. Clone the Repository

```bash
git clone https://github.com/sarjeetshekhawat/bb-automation.git
cd bb-automation
```

### 2. Create a Virtual Environment (Recommended)

```bash
# Windows (PowerShell)
python -m venv .venv
.venv\Scripts\Activate.ps1

# Linux / macOS
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

---

## 🔧 Configuration

### Streamlit Configuration (`.streamlit/config.toml`)

The application includes pre-configured server and theme settings in `.streamlit/config.toml`:
- `server.maxUploadSize = 3000`: Allows uploading datasets up to 3 GB.
- `server.maxMessageSize = 3000`: Allows processing large payload transfers.
- `browser.gatherUsageStats = false`: Disables external telemetry.

### Optional Environment Variables (`.env`)

To customize port or server binding, copy `.env.example` to `.env`:

```bash
# Windows
copy .env.example .env

# Linux / macOS
cp .env.example .env
```

Available variables:
- `STREAMLIT_SERVER_PORT`: Custom listening port (default: `8501`).
- `STREAMLIT_SERVER_ADDRESS`: IP/host to bind to (default: `localhost`).
- `STREAMLIT_SERVER_HEADLESS`: Set to `true` when running on headless servers or containers.

---

## 🖥 Execution

Start the Streamlit application:

```bash
streamlit run app.py
```

Upon launch, Streamlit will open your default browser at:
```text
http://localhost:8501
```

### Step-by-Step Usage

1. **Upload Files**: In the upload card, supply the four required files:
   - **Blocked Bill File** (primary dataset)
   - **WC Consumer Installation File**
   - **LTCT Consumer Installation File**
   - **HTCT Consumer Installation File**
2. **Select Identifier Columns**: The application automatically inspects headers and suggests candidate identifier columns (e.g., `Meter Number`, `Consumer ID`, `Badge Number`). Confirm or adjust the column selection.
3. **Run Classification**: Click **Classify Meters**. The engine checks for duplicate keys in the installation files and categorises each blocked record.
4. **Inspect Dashboard**: Review the KPI metrics, category breakdowns, and tabbed data previews.
5. **Download Report**: Click **Download Classification Report** to receive `Blocked_VILL_Meter_Classification.xlsx`.

---

## 🧠 How the Engine Works

1. **Safe Loading (`read_excel_safe`)**:
   - Inspects file extensions (`.xlsx`, `.xls`, `.csv`).
   - Uses `openpyxl` with fallback to `xlrd` and `pandas.read_csv`.
   - Forces `dtype=str` across all columns.
2. **Identifier Normalisation (`normalise_key`)**:
   - Strips leading and trailing whitespace.
   - Converts to uppercase.
   - Cleanses placeholders (`NAN`, `NONE`, empty strings).
3. **Lookup Compilation (`build_lookup_set`)**:
   - Compiles sets for WC, LTCT, and HTCT identifiers for $O(1)$ membership checks.
4. **Classification (`classify_meters`)**:
   - Compares each blocked record key against the lookup sets.
   - Assigns category: `WC`, `LTCT`, `HTCT`, `NOT FOUND`, or `MULTIPLE MATCH`.
   - Populates match flags (`Matched_WC`, `Matched_LTCT`, `Matched_HTCT`, `Match_Status`).

---

## 📊 Excel Report Breakdown

The exported workbook (`Blocked_VILL_Meter_Classification.xlsx`) contains 7 worksheets:

| Sheet Name | Description |
|---|---|
| **Summary** | Executive overview with total counts and classification percentage |
| **Classified Data** | Complete blocked master dataset with classification columns appended |
| **WC Meters** | Records exclusively matching WC installation records (Green badge) |
| **LTCT Meters** | Records exclusively matching LTCT installation records (Blue badge) |
| **HTCT Meters** | Records exclusively matching HTCT installation records (Amber badge) |
| **Not Found** | Meters not located in any installation dataset (Red badge) |
| **Multiple Matches** | Discrepancy report for meters found in more than one installation file (Purple badge) |

---

## 🧪 Testing & Validation

### Syntax & Compilation Check

Verify all Python modules compile without syntax errors:

```bash
python -m py_compile app.py engine.py
```

### Engine Smoke Test

Run an in-memory classification verification test:

```bash
python -c "import pandas as pd; from engine import normalise_key, build_lookup_set, classify_meters, compute_summary; df_wc = pd.DataFrame({'meter': ['M1', 'M2']}); df_b = pd.DataFrame({'meter': ['m1', 'M3']}); s_wc = build_lookup_set(df_wc, 'meter'); res = classify_meters(df_b, 'meter', s_wc, set(), set()); sm = compute_summary(res); assert sm['WC Meters'] == 1 and sm['Not Found'] == 1; print('All smoke checks passed!')"
```

---

## ❓ Troubleshooting & FAQ

### 1. Large File Upload Errors
- If a file exceeds default Streamlit upload limits, check `.streamlit/config.toml`. It is configured to handle files up to 3000 MB (`maxUploadSize = 3000`).

### 2. "Permission Denied" When Writing Report
- Ensure you do not have a previously downloaded file with the same name open in Microsoft Excel while downloading a new copy.

### 3. File Encoding Issues in CSV
- CSV files should ideally be UTF-8 or standard ASCII encoded. If reading fails, save the file as `.xlsx` in Excel and upload.

---

## 🔒 Security & Data Privacy

- **100% Local Processing**: All file processing occurs entirely within local system memory.
- **Zero Remote Telemetry**: Telemetry gathering is disabled (`gatherUsageStats = false`).
- **No External API Dependencies**: No cloud APIs, credentials, or remote database connections are used.
