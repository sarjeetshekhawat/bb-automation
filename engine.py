"""
Core classification engine for Blocked VILL Meter Classification.

Handles all data processing: reading Excel files, performing exact-match
VLOOKUP-equivalent lookups, classifying meters, detecting duplicates,
and generating the final multi-sheet Excel report.
"""

import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ──────────────────────────────────────────────────────────────────────
# 1.  Reading & Normalising Excel Files
# ──────────────────────────────────────────────────────────────────────

def read_excel_safe(file, sheet_name=0) -> pd.DataFrame:
    """Read an Excel or CSV file into a DataFrame.

    Accepts either a file-like object (from Streamlit upload) or a
    string file path (for direct disk access on large files).

    All columns are read as **text** so leading zeros and long numbers
    are preserved exactly as they appear in the spreadsheet.
    Supports .xlsx (openpyxl), .xls (xlrd), and .csv files.
    """
    import os

    # Determine filename for extension detection
    if isinstance(file, str):
        filename = os.path.basename(file).lower()
    elif hasattr(file, "name"):
        filename = file.name.lower()
    else:
        filename = ""

    def _try_seek():
        """Reset seek on file-like objects; no-op for paths."""
        if not isinstance(file, str) and hasattr(file, "seek"):
            try:
                file.seek(0)
            except Exception:
                pass

    # CSV path — detected by extension
    if filename.endswith(".csv"):
        _try_seek()
        return pd.read_csv(file, dtype=str, keep_default_na=False)

    # Excel path — try openpyxl first, then xlrd, then CSV as last resort
    try:
        return pd.read_excel(file, sheet_name=sheet_name, dtype=str,
                             engine="openpyxl")
    except Exception:
        _try_seek()
        try:
            return pd.read_excel(file, sheet_name=sheet_name, dtype=str,
                                 engine="xlrd")
        except Exception:
            # Last resort: try reading as CSV
            _try_seek()
            return pd.read_csv(file, dtype=str, keep_default_na=False)


def get_column_names(file, sheet_name=0) -> list[str]:
    """Return column header names from an Excel or CSV file.

    For large CSV files passed as a path, only the header row is read
    to avoid loading gigabytes of data just for column names.
    """
    import os

    # Fast path for CSV files passed as a string path
    if isinstance(file, str):
        ext = os.path.splitext(file)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(file, dtype=str, nrows=0)
            return list(df.columns)
        # Excel path — read just the header via nrows trick isn't
        # supported the same way, so fall through to read_excel_safe
    elif hasattr(file, "name") and file.name.lower().endswith(".csv"):
        # File-like CSV — read only header
        try:
            file.seek(0)
        except Exception:
            pass
        df = pd.read_csv(file, dtype=str, nrows=0)
        return list(df.columns)

    df = read_excel_safe(file, sheet_name)
    return list(df.columns)


def normalise_key(series: pd.Series) -> pd.Series:
    """Strip whitespace, collapse internal spaces, upper-case, remove NaN."""
    return (
        series
        .astype(str)
        .str.strip()
        .str.upper()
        .replace({"NAN": "", "NONE": "", "": pd.NA})
    )


# ──────────────────────────────────────────────────────────────────────
# 2.  Duplicate Detection Inside a Single Installation File
# ──────────────────────────────────────────────────────────────────────

def find_internal_duplicates(df: pd.DataFrame, key_col: str) -> pd.DataFrame:
    """Return rows whose key value appears more than once in *df*."""
    keys = normalise_key(df[key_col])
    mask = keys.duplicated(keep=False) & keys.notna()
    return df[mask].copy()


# ──────────────────────────────────────────────────────────────────────
# 3.  Building Lookup Sets
# ──────────────────────────────────────────────────────────────────────

def build_lookup_set(df: pd.DataFrame, key_col: str) -> set:
    """Build a set of normalised, non-null identifiers from *df[key_col]*."""
    keys = normalise_key(df[key_col])
    return set(keys.dropna())


# ──────────────────────────────────────────────────────────────────────
# 4.  Classification
# ──────────────────────────────────────────────────────────────────────

def classify_meters(
    blocked_df: pd.DataFrame,
    blocked_key_col: str,
    wc_set: set,
    ltct_set: set,
    htct_set: set,
    progress_callback=None,
) -> pd.DataFrame:
    """Classify every row in *blocked_df* against the three lookup sets.

    Returns a copy of *blocked_df* with five new columns appended:
        Matched_WC, Matched_LTCT, Matched_HTCT,
        Final_Meter_Category, Match_Status
    """
    result = blocked_df.copy()
    keys = normalise_key(result[blocked_key_col])

    n = len(result)
    matched_wc = []
    matched_ltct = []
    matched_htct = []
    categories = []
    statuses = []

    for idx, key in enumerate(keys):
        if progress_callback and idx % max(1, n // 100) == 0:
            progress_callback(idx / n)

        if pd.isna(key) or key == "":
            matched_wc.append("No")
            matched_ltct.append("No")
            matched_htct.append("No")
            categories.append("NOT FOUND")
            statuses.append("Blank/Missing Identifier")
            continue

        in_wc = key in wc_set
        in_ltct = key in ltct_set
        in_htct = key in htct_set

        matched_wc.append("Yes" if in_wc else "No")
        matched_ltct.append("Yes" if in_ltct else "No")
        matched_htct.append("Yes" if in_htct else "No")

        match_count = sum([in_wc, in_ltct, in_htct])

        if match_count == 0:
            categories.append("NOT FOUND")
            statuses.append("Not Found")
        elif match_count == 1:
            if in_wc:
                categories.append("WC")
            elif in_ltct:
                categories.append("LTCT")
            else:
                categories.append("HTCT")
            statuses.append("Matched")
        else:
            # Matched in more than one file
            files = []
            if in_wc:
                files.append("WC")
            if in_ltct:
                files.append("LTCT")
            if in_htct:
                files.append("HTCT")
            categories.append("MULTIPLE MATCH")
            statuses.append(f"Multiple Match ({', '.join(files)})")

    result["Matched_WC"] = matched_wc
    result["Matched_LTCT"] = matched_ltct
    result["Matched_HTCT"] = matched_htct
    result["Final_Meter_Category"] = categories
    result["Match_Status"] = statuses

    if progress_callback:
        progress_callback(1.0)

    return result


# ──────────────────────────────────────────────────────────────────────
# 5.  Summary Statistics
# ──────────────────────────────────────────────────────────────────────

def compute_summary(classified_df: pd.DataFrame) -> dict:
    """Compute classification summary statistics."""
    total = len(classified_df)
    cats = classified_df["Final_Meter_Category"]

    wc = int((cats == "WC").sum())
    ltct = int((cats == "LTCT").sum())
    htct = int((cats == "HTCT").sum())
    not_found = int((cats == "NOT FOUND").sum())
    multiple = int((cats == "MULTIPLE MATCH").sum())
    classified = wc + ltct + htct
    pct = round(classified / total * 100, 2) if total else 0.0

    return {
        "Total Blocked Meters Processed": total,
        "WC Meters": wc,
        "LTCT Meters": ltct,
        "HTCT Meters": htct,
        "Not Found": not_found,
        "Multiple Matches": multiple,
        "Successfully Classified": classified,
        "Classification %": pct,
    }


# ──────────────────────────────────────────────────────────────────────
# 6.  Excel Report Generation (openpyxl)
# ──────────────────────────────────────────────────────────────────────

# ── Colour Palette ────────────────────────────────────────────────────
_HEADER_FILL    = PatternFill("solid", fgColor="4F46E5")   # Indigo
_HEADER_FONT    = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
_DATA_FONT      = Font(name="Calibri", size=11)
_THIN_BORDER    = Border(
    left=Side(style="thin", color="334155"),
    right=Side(style="thin", color="334155"),
    top=Side(style="thin", color="334155"),
    bottom=Side(style="thin", color="334155"),
)

_CATEGORY_FILLS = {
    "WC":             PatternFill("solid", fgColor="10B981"),  # Emerald
    "LTCT":           PatternFill("solid", fgColor="3B82F6"),  # Blue
    "HTCT":           PatternFill("solid", fgColor="F59E0B"),  # Amber
    "NOT FOUND":      PatternFill("solid", fgColor="EF4444"),  # Red
    "MULTIPLE MATCH": PatternFill("solid", fgColor="A855F7"),  # Purple
}
_CATEGORY_FONTS = {
    k: Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    for k in _CATEGORY_FILLS
}


def _write_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame):
    """Write a DataFrame to a named worksheet with formatting."""
    ws = wb.create_sheet(title=sheet_name)
    cols = list(df.columns)

    # ── Headers ───────────────────────────────────────────────────────
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = _THIN_BORDER

    # ── Data Rows ─────────────────────────────────────────────────────
    cat_col_idx = None
    if "Final_Meter_Category" in cols:
        cat_col_idx = cols.index("Final_Meter_Category") + 1  # 1-based

    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx,
                           value="" if pd.isna(value) else str(value))
            cell.font = _DATA_FONT
            cell.border = _THIN_BORDER
            cell.alignment = Alignment(horizontal="left")

        # Colour-code the category cell
        if cat_col_idx:
            cat_cell = ws.cell(row=r_idx, column=cat_col_idx)
            cat_val = cat_cell.value
            if cat_val in _CATEGORY_FILLS:
                cat_cell.fill = _CATEGORY_FILLS[cat_val]
                cat_cell.font = _CATEGORY_FONTS[cat_val]
                cat_cell.alignment = Alignment(horizontal="center")

    # ── Auto-width ────────────────────────────────────────────────────
    for c_idx, col_name in enumerate(cols, start=1):
        max_len = len(str(col_name))
        for r_idx in range(2, min(len(df) + 2, 102)):  # sample first 100 rows
            val = ws.cell(row=r_idx, column=c_idx).value or ""
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 4, 50)

    # ── Freeze header ─────────────────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Auto-filter ───────────────────────────────────────────────────
    if len(df) > 0:
        ws.auto_filter.ref = ws.dimensions


def _write_summary_sheet(wb: Workbook, summary: dict):
    """Add a formatted summary sheet at the beginning of the workbook."""
    ws = wb.create_sheet(title="Summary", index=0)

    title_font = Font(name="Calibri", bold=True, size=16, color="4F46E5")
    label_font = Font(name="Calibri", bold=True, size=12, color="E2E8F0")
    value_font = Font(name="Calibri", size=12, color="FFFFFF")
    label_fill = PatternFill("solid", fgColor="1E293B")
    value_fill = PatternFill("solid", fgColor="334155")
    pct_font   = Font(name="Calibri", bold=True, size=14, color="10B981")

    ws.merge_cells("A1:C1")
    title_cell = ws.cell(row=1, column=1,
                         value="Blocked VILL Meter Classification — Summary")
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    row = 3
    for label, value in summary.items():
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = label_font
        lc.fill = label_fill
        lc.border = _THIN_BORDER
        lc.alignment = Alignment(horizontal="left")

        vc = ws.cell(row=row, column=2,
                     value=f"{value}%" if label == "Classification %" else value)
        if label == "Classification %":
            vc.font = pct_font
        else:
            vc.font = value_font
        vc.fill = value_fill
        vc.border = _THIN_BORDER
        vc.alignment = Alignment(horizontal="center")
        row += 1

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 5


def generate_report(classified_df: pd.DataFrame, summary: dict) -> bytes:
    """Build the multi-sheet Excel report and return it as bytes."""
    wb = Workbook()
    # Remove the default sheet created by openpyxl
    wb.remove(wb.active)

    # Summary sheet (inserted at index 0 inside the helper)
    _write_summary_sheet(wb, summary)

    # Sheet 1 — Classified Data
    _write_sheet(wb, "Classified Data", classified_df)

    # Sheet 2–6 — category slices
    for cat, sheet_name in [
        ("WC", "WC Meters"),
        ("LTCT", "LTCT Meters"),
        ("HTCT", "HTCT Meters"),
        ("NOT FOUND", "Not Found"),
        ("MULTIPLE MATCH", "Multiple Matches"),
    ]:
        subset = classified_df[
            classified_df["Final_Meter_Category"] == cat
        ].reset_index(drop=True)
        _write_sheet(wb, sheet_name, subset)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
