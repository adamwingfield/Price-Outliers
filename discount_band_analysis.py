"""Generate discount band analysis charts from customer sales data,
analysed separately for each PRODUCT GROUP.

Product-group version: each analysed record is one customer/group in one
product group (e.g. VIAMED - KNEES, VIAMED - SURGICAL), and a separate trend
line, SD bands, size bins and percentiles are calculated for each product
group, because typical discount levels differ widely between product groups.
One chart is produced per product group; all results go into one Excel file
with a Product Group column, a Product Group Summary sheet and a Charts
sheet.  Use --combined to switch back to one record per customer across all
product groups.

Adapted from ``price_band_analysis.py``.  Instead of comparing average
procedure price against procedure quantity (log-log), this version compares
the total discount from list price (Discount %) against total list value
(customer size measured at list price, independent of the discount given).

The script reads a CSV or Excel file containing a customer (or customer
group) column, a total list-price value column, and a total net-sales
column.  It then:

  1. Standardizes the input columns (see ``COLUMN_ALIASES``).
  1b. Applies the optional include / exclude filters (customer type,
     customer name, product group -- see ``FILTER_DEFINITIONS``): include
     lists keep only matching rows, then exclude lists remove rows.
  2. Aggregates rows that share the same customer/group name into a single
     record: list value and net sales are summed and the discount is
     RECALCULATED from the totals, ``1 - sum(net) / sum(list)``, which makes
     it a list-value-weighted discount (configurable via
     ``CONFIG["aggregate_by_customer"]`` or ``--no-aggregate``).  Any
     discount % column in the source file is ignored because averaging row
     percentages would give an unweighted, misleading result.
  3. Applies the minimum-list-value and discount-range thresholds to the
     aggregated records (records with non-positive list or net value, e.g.
     credit / return-only customers, are excluded).
  4. Fits a trend line in SEMI-LOG space (Discount % on a linear scale vs
     log10 total list value).  List value is used for the x-axis instead
     of net sales because net sales already contains the discount
     (net = list x (1 - discount)), which would build part of the slope in
     by construction.  A log scale cannot be used for the discount axis
     because zero and negative discounts (net above list) are common.
     Each record is classified against +/-1, +/-2 and +/-3 standard-deviation
     bands (measured in discount percentage points), and within-bin and
     percentile-based outlier checks are run.
  5. Displays one scatter plot per product group (log x-axis, linear %
     y-axis) and exports the full analysis table, product-group summary,
     charts and filter audit to a single formatted Excel file.

Sign convention: Discount % is POSITIVE when net sales are below list
(e.g. 0.25 = 25 % off list) and NEGATIVE when net sales exceed list.  Note
that this is the opposite sign to the "Net Discount %" column in the
Pricemart extract, which reports (net - list) / list.

Usage:
    python discount_band_analysis.py
    python discount_band_analysis.py input_file.csv
    python discount_band_analysis.py input_file.xlsx --sheet "Data Table"
    python discount_band_analysis.py input_file.xlsx --no-aggregate
    python discount_band_analysis.py input_file.xlsx --combined
    python discount_band_analysis.py input_file.xlsx \
        --exclude-customer-type "Distributor" \
        --exclude-customer "*PHARMACY*" --exclude-product-group "CEMENT"
    python discount_band_analysis.py input_file.xlsx \
        --include-product-group "KNEES" --include-product-group "HIPS"
    python discount_band_analysis.py input_file.xlsx --filters-file filters.xlsx

Dependencies:
    pandas, numpy, matplotlib, openpyxl (for .xlsx input/output),
    Pillow (required by openpyxl to embed the chart images)
"""

from __future__ import annotations

import argparse
import fnmatch
import re
import shutil
import sys
import tempfile
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import PercentFormatter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError:
    tk = None
    filedialog = None
    messagebox = None


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Accepted aliases for each required input column, in match-priority order.
# Matching is case-insensitive and treats underscores as spaces (see
# ``normalize_column_name``).  "group or customer" is listed first so that,
# when a file carries both a "Customer" column and a "Group or Customer"
# column, the group-level column is used.
COLUMN_ALIASES: dict[str, list[str]] = {
    "customer_name": [
        "group or customer",
        "customer name",
        "customer group",
        "customer group l3",
        "business partner",
        "customer",
        "account",
        "account name",
        "name",
    ],
    "list_value": [
        "country list price",
        "list price",
        "list value",
        "total list price",
        "list sales",
        "gross sales",
        "extended list price",
    ],
    "net_sales": [
        "net price including payment terms cost",
        "net sales",
        "total net sales",
        "net price",
        "net revenue",
        "net value",
        "invoice price",
    ],
}

# Display names used after column standardization.
STANDARD_COLUMN_NAMES: dict[str, str] = {
    "customer_name": "Customer Name",
    "list_value": "List Value",
    "net_sales": "Net Sales",
}
CUSTOMER_COL = STANDARD_COLUMN_NAMES["customer_name"]
LIST_COL = STANDARD_COLUMN_NAMES["list_value"]
NET_COL = STANDARD_COLUMN_NAMES["net_sales"]

# Derived columns (calculated, never read from the input).
DISCOUNT_COL = "Discount %"              # 1 - Net Sales / List Value
DISCOUNT_AMOUNT_COL = "Discount Amount"  # List Value - Net Sales

# Column used for the x-axis, trend fit, size bins and size percentiles.
X_COL = LIST_COL

# Number of source rows combined into each analysed record (added only when
# aggregation is enabled).
RECORDS_AGGREGATED_COL = "Records Aggregated"

# Placeholder used for blank / missing customer names.  Rows carrying this
# placeholder are never aggregated together, because they are not known to
# belong to the same customer.
UNKNOWN_CUSTOMER = "Unknown"

# Severity labels in the order they should appear in the Excel export
# (most severe first).
SEVERITY_ORDER: list[str] = [
    "Extreme Concern Outlier",
    "High Concern Outlier",
    "Lower Concern Outlier",
    "Normal",
    "Not Assessed",
]

# Product-group column (standard name) and placeholder for blank values.
PRODUCT_GROUP_COL = "Product Group"
UNKNOWN_PRODUCT_GROUP = "Unknown"
# Label used for the single segment when product groups are not analysed
# separately (--combined, or no product-group column in the input).
ALL_PRODUCTS_LABEL = "All Product Groups"

# Chart size in inches, and the approximate plot-area dimensions (in
# typographic points) for that figure size.  The plot-area values are used
# solely to convert point offsets into proportional plot-space distances for
# label-collision scoring -- update them together if FIGSIZE changes.
FIGSIZE: tuple[float, float] = (12, 8)
_PLOT_WIDTH_PTS: float = 650.0
_PLOT_HEIGHT_PTS: float = 450.0

# ---------------------------------------------------------------------------
# Optional include / exclude filters
# ---------------------------------------------------------------------------
# Filters are applied to the raw input rows BEFORE aggregation, thresholds
# and trend fitting.  Matching is case-insensitive and ignores leading /
# trailing spaces; "*" and "?" wildcards are supported (e.g. "*PHARMACY*").
#
# For each filter type:
#   * include_values -- if not empty, ONLY rows matching at least one value
#                       are kept ("include only these").
#   * exclude_values -- rows matching any value are removed.
# Include lists are applied first, then exclude lists, so an exclusion
# always wins (e.g. include product group "KNEES*" but exclude "KNEES - REV").
# When several filter types have include lists, a row must satisfy all of
# them (e.g. product group in list AND customer type in list).
#
# Missing columns:
#   * An EXCLUDE filter whose column is not in the input file is skipped with
#     a warning (nothing to exclude).
#   * An INCLUDE filter whose column is not in the input file stops the run
#     with an error, because analysing every row would silently ignore the
#     "include only" request.
#
# "customer_name" values are checked against EVERY customer-name column
# present (e.g. "Customer", "Group or Customer", "IDN Child / Buying Grp 3").
# Including / excluding a group name therefore keeps / removes all of its
# member accounts; excluding a member account removes it from its group's
# totals.
#
# Values can also be supplied on the command line (--include-* / --exclude-*)
# or in a filters file (--filters-file); all sources are combined.
FILTER_DEFINITIONS: dict[str, dict] = {
    "customer_type": {
        "label": "Customer Type",
        # Underscore variants are not needed: aliases are matched via
        # normalize_column_name, which treats "_" as a space.
        "aliases": ["customer type", "account type"],
        "match_all_columns": False,
        "include_values": [],          # e.g. ["Hospital", "ASC"]
        # NOTE: active default -- distributor-type customers are excluded on
        # every run unless this list is emptied.
        "exclude_values": ["*DIST*"],  # e.g. ["Distributor", "Pharmacy"]
    },
    "customer_name": {
        "label": "Customer Name",
        "aliases": [
            "group or customer", "customer", "customer name",
            "idn child / buying grp 3", "customer group", "customer group l3",
            "business partner", "account", "account name",
        ],
        "match_all_columns": True,
        "include_values": [],          # e.g. ["0000124806 HELIOS*"]
        "exclude_values": [],          # e.g. ["*PHARMACY*"]
    },
    "product_group": {
        "label": "Product Group",
        "aliases": [
            "product group", "productgroup", "product type", "product line",
            "product category", "product family", "product hierarchy",
        ],
        "match_all_columns": False,
        "include_values": [],          # e.g. ["KNEES", "HIPS"]
        "exclude_values": [],          # e.g. ["CEMENT", "SURGICAL"]
    },
}

FILTER_ACTIONS = ("include", "exclude")

# Accepted values in the "Filter Type" column of a filters file, mapped to
# FILTER_DEFINITIONS keys.  Looked up via normalize_column_name, so
# underscore variants ("customer_type") are matched automatically.
FILTER_FILE_TYPE_ALIASES: dict[str, str] = {
    "customer type": "customer_type",
    "type": "customer_type",
    "customer name": "customer_name",
    "customer": "customer_name",
    "product group": "product_group",
    "product type": "product_group",
    "product": "product_group",
}

CONFIG: dict = {
    # Product-group analysis: when True (default for this version) each
    # record is customer/group x product group and every product group gets
    # its own trend line, bands, bins and percentiles.  Falls back to the
    # combined analysis (with a warning) if the input has no product-group
    # column.  Override with --combined.
    "analyse_by_product_group": True,
    # Product groups with fewer analysable records than this are listed in
    # the output as "Not Assessed" (no trend line is fitted for them).
    # Values below 3 are raised to 3 (the minimum for a residual SD).
    "min_records_per_product_group": 10,
    # Optional per-product-group minimum list value, overriding
    # min_list_value below for smaller-value product lines, e.g.
    # {"CEMENT": 2000, "SURGICAL": 2000}.  Keys are matched
    # case-insensitively.
    "min_list_value_by_product_group": {},
    # Aggregation: when True, rows sharing the same customer/group name are
    # combined (list and net summed, discount recalculated) BEFORE the
    # thresholds below are applied.
    "aggregate_by_customer": True,
    # Data-filtering thresholds (applied to aggregated records).
    # Records must have List Value >= min_list_value (the x-axis measure).
    "min_list_value": 10000,
    # Discount range kept for analysis, as fractions (0.25 = 25 % off list).
    # Negative values mean net sales above list.  Records outside this range
    # are treated as data issues (credits, rebills, mis-mapped list prices).
    "min_discount_pct": -0.50,
    "max_discount_pct": 0.95,
    # Bin configuration for within-bin outlier analysis
    "max_list_value_bins": 4,
    "min_records_per_bin": 5,
    # When True, Discount Position is left blank for points within 1 SD;
    # when False, those points are labelled "In Line".
    "band_blank_within_1sd_discount_position": True,
    # Points beyond this many SDs from the trend line receive scatter labels
    "scatter_label_sd_threshold": 1.5,
    # Within-bin outlier-detection parameters
    "robust_zscore_threshold": 3.5,
    "robust_iqr_multiplier": 1.5,
    "percentile_lower": 0.05,
    "percentile_upper": 0.95,
    # Excel output settings
    "default_output_filename": "discount_band_outliers_by_product_group.xlsx",
    "summary_sheet_name": "Product Group Summary",
    "charts_sheet_name": "Charts",
    "worksheet_name": "Discount Band Analysis",
    "filters_sheet_name": "Filters Applied",
    "removed_rows_sheet_name": "Removed Rows",
    "header_fill_color": "002060",
    "header_font_color": "FFFFFF",
    "severity_fill_colors": {
        "Normal": "DDEBF7",
        "Not Assessed": "EDEDED",
        "Lower Concern Outlier": "FCE4D6",
        "High Concern Outlier": "F8CBAD",
        "Extreme Concern Outlier": "FF8B8B",
    },
    # Percentile-based outlier flag thresholds (0-100 scale).
    # Records at or below the low threshold on list value AND at or above the
    # high threshold on discount are flagged "Outlier - High Discount, Low
    # List Value" (the commercially most concerning quadrant).  Records at or
    # above the high threshold on list value AND at or below the low
    # threshold on discount are flagged "Outlier - Low Discount, High List
    # Value".
    "outlier_low_percentile_threshold": 30,
    "outlier_high_percentile_threshold": 70,
}

# Excel number formats, keyed by column header.  One mapping serves every
# output sheet (analysis, summary, filter audit); headers not listed are left
# with the default format.
EXCEL_NUMBER_FORMATS: dict[str, str] = {
    **dict.fromkeys(
        [
            LIST_COL, NET_COL, DISCOUNT_AMOUNT_COL, "Net Sales Impact at Trend",
            "Net Sales Upside if Over-Discounted Moved to Trend",
            "List Value Removed", "Net Sales Removed",
        ],
        "$#,##0",
    ),
    **dict.fromkeys(
        [
            "Overall Discount %", "Median Record Discount %",
            "Trend Slope (pts per 10x List Value)", "Residual SD (pts)",
        ],
        "0.0%",
    ),
    **dict.fromkeys(
        [
            "Discount %", "Trend Discount %", "Discount Gap vs Trend",
            "Bin Median Discount %", "Bin IQR Lower Bound", "Bin IQR Upper Bound",
            "Bin Lower Percentile Threshold", "Bin Upper Percentile Threshold",
        ],
        "0.0%",
    ),
    **dict.fromkeys(["Bin Percentile Rank"], "0.0000"),
    **dict.fromkeys(["Bin Robust Z Score", "SD From Trend"], "0.000"),
    **dict.fromkeys(["List Value Percentile", "Discount Percentile"], "0.0"),
}


# ---------------------------------------------------------------------------
# Column-name helpers
# ---------------------------------------------------------------------------

def normalize_column_name(column_name: str) -> str:
    """Lowercase, strip, convert underscores to spaces, and collapse whitespace."""
    return " ".join(str(column_name).strip().lower().replace("_", " ").split())


def find_matching_column(columns: pd.Index, aliases: list[str]) -> str:
    """Return the first column whose normalized name matches an alias.

    Aliases are tried in list order, so earlier aliases take priority when a
    file contains several candidate columns.

    Raises ``KeyError`` if no alias matches any column in *columns*.
    """
    normalized_map = {normalize_column_name(col): col for col in columns}
    for alias in aliases:
        normalized_alias = normalize_column_name(alias)
        if normalized_alias in normalized_map:
            return normalized_map[normalized_alias]

    raise KeyError(f"Could not find a matching column for aliases: {aliases}")


def to_plot_space(list_value, discount) -> np.ndarray:
    """Convert (list value, discount) to chart coordinates (log10 x, linear y)."""
    return np.column_stack([
        np.log10(np.asarray(list_value, dtype=float)),
        np.asarray(discount, dtype=float),
    ])


# ---------------------------------------------------------------------------
# GUI message helpers
# ---------------------------------------------------------------------------

def _show_tk_message(title: str, message: str, *, level: str = "info") -> None:
    """Display a message via a tkinter dialog, falling back to the console
    when tkinter or a display is unavailable.

    Parameters
    ----------
    title : str
        Dialog / console prefix.
    message : str
        Body text.
    level : str
        ``"info"`` for informational messages (stdout) or ``"warning"`` for
        warnings (stderr).
    """
    if messagebox and tk:
        try:
            root = tk.Tk()
        except tk.TclError:
            # tkinter is installed but no display is available (e.g. a
            # headless server); fall through to console output.
            pass
        else:
            root.withdraw()
            try:
                dialog_fn = (
                    messagebox.showwarning if level == "warning"
                    else messagebox.showinfo
                )
                dialog_fn(title, message)
            finally:
                root.destroy()
            return

    stream = sys.stderr if level == "warning" else sys.stdout
    print(f"{title}: {message}", file=stream)


def show_info_message(title: str, message: str) -> None:
    """Display an informational message (GUI dialog or console)."""
    _show_tk_message(title, message, level="info")


def show_warning_message(title: str, message: str) -> None:
    """Display a warning message (GUI dialog or console)."""
    _show_tk_message(title, message, level="warning")


# ---------------------------------------------------------------------------
# File-dialog helpers
# ---------------------------------------------------------------------------

def _require_tk(purpose: str) -> None:
    """Raise a helpful error when tkinter dialogs are unavailable."""
    if not tk or not filedialog:
        raise RuntimeError(
            f"{purpose} dialog is not available in this Python environment. "
            "Please provide the file as a command-line argument."
        )


def choose_input_file() -> Path:
    """Open a file-selection dialog and return the chosen input path."""
    _require_tk("File")

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected_file = filedialog.askopenfilename(
            title="Select input file for discount band analysis",
            filetypes=[
                ("Data files", "*.csv *.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected_file:
        raise ValueError("No input file selected.")

    return Path(selected_file)


def choose_output_file(default_name: str) -> Path:
    """Open a save-as dialog and return the chosen Excel output path."""
    _require_tk("Save")

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected_file = filedialog.asksaveasfilename(
            title="Select location for Excel output",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected_file:
        raise ValueError("No output file selected.")

    return Path(selected_file)


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_excel_sheet(
    file_path: Path,
    sheet_name: str | None = None,
) -> pd.DataFrame:
    """Load a specific Excel sheet or auto-detect the most likely data sheet.

    When *sheet_name* is ``None`` the function checks each sheet for a column
    that matches one of the ``net_sales`` aliases (this skips metadata sheets
    such as the Pricemart "Settings" tab).  The first match wins; if none
    match the first sheet is returned.
    """
    if sheet_name:
        return pd.read_excel(file_path, sheet_name=sheet_name)

    net_aliases = {normalize_column_name(a) for a in COLUMN_ALIASES["net_sales"]}

    # Only the header row of each sheet is read during detection; the full
    # data is then read once, from the selected sheet only.
    with pd.ExcelFile(file_path) as workbook:
        selected = workbook.sheet_names[0]
        for candidate in workbook.sheet_names:
            header = workbook.parse(candidate, nrows=0)
            if net_aliases & {normalize_column_name(c) for c in header.columns}:
                selected = candidate
                break
        return workbook.parse(selected)


def load_input_file(
    file_path: Path,
    sheet_name: str | None = None,
) -> pd.DataFrame:
    """Load a CSV or Excel input file into a DataFrame."""
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(file_path)
    if suffix in {".xlsx", ".xls"}:
        return load_excel_sheet(file_path, sheet_name)

    raise ValueError("Unsupported file type.  Use CSV or Excel (.xlsx/.xls).")


# ---------------------------------------------------------------------------
# Include / exclude filters
# ---------------------------------------------------------------------------

REMOVED_BY_COL = "Removed By"

FILTER_LOG_COLUMNS = [
    "Action", "Filter", "Value", "Column(s) Checked", "Status",
    "Rows Matched", "Rows Removed", "List Value Removed", "Net Sales Removed",
]


def find_all_matching_columns(columns: pd.Index, aliases: list[str]) -> list[str]:
    """Return every column whose normalized name matches any alias."""
    wanted = {normalize_column_name(a) for a in aliases}
    return [c for c in columns if normalize_column_name(c) in wanted]


def _normalize_value(value) -> str:
    """Lowercase, strip and collapse whitespace (underscores are kept)."""
    return " ".join(str(value).strip().lower().split())


def _normalize_series(series: pd.Series) -> pd.Series:
    """Vectorised ``_normalize_value`` for a whole column (NaN -> "")."""
    return (
        series.fillna("").astype(str).str.lower()
        .str.split().str.join(" ")
    )


def _value_matches(normalized_values: pd.Series, pattern: str) -> pd.Series:
    """Case-insensitive exact or wildcard match of already-normalized values
    (see ``_normalize_series``) against *pattern*."""
    pat = _normalize_value(pattern)
    if any(ch in pat for ch in "*?["):
        regex = re.compile(fnmatch.translate(pat))
        return normalized_values.map(lambda v: regex.match(v) is not None)
    return normalized_values == pat


def _empty_filter_lists() -> dict[str, dict[str, list[str]]]:
    return {action: {key: [] for key in FILTER_DEFINITIONS} for action in FILTER_ACTIONS}


def load_filters_file(path: Path) -> dict[str, dict[str, list[str]]]:
    """Read a filters list (CSV or Excel; first sheet for Excel).

    Columns:
      * "Action"      -- Include or Exclude (optional; blank = Exclude, so
                         older two-column exclusion lists still work)
      * "Filter Type" -- Customer Type / Customer Name / Product Group
      * "Value"       -- value or wildcard pattern
    """
    if path.suffix.lower() == ".csv":
        table = pd.read_csv(path, dtype=str)
    else:
        table = pd.read_excel(path, dtype=str)
    type_col = find_matching_column(table.columns, ["filter type", "filter", "type"])
    # "exclude" / "exclusion value" are legacy headers from the earlier
    # exclusions-only file format and are kept for backward compatibility.
    value_col = find_matching_column(
        table.columns, ["value", "exclude", "exclusion value"]
    )
    try:
        action_col = find_matching_column(
            table.columns, ["action", "include / exclude", "include or exclude"]
        )
    except KeyError:
        action_col = None

    result = _empty_filter_lists()
    problems: set[str] = set()
    for row in table.to_dict("records"):
        ftype, value = row[type_col], row[value_col]
        if pd.isna(ftype) or pd.isna(value) or not str(value).strip():
            continue
        key = FILTER_FILE_TYPE_ALIASES.get(normalize_column_name(ftype))
        if key is None:
            problems.add(f"unrecognised filter type '{ftype}'")
            continue
        raw_action = row[action_col] if action_col is not None else None
        action = (
            _normalize_value(raw_action)
            if raw_action is not None and not pd.isna(raw_action) and str(raw_action).strip()
            else "exclude"
        )
        if action not in FILTER_ACTIONS:
            problems.add(f"unrecognised action '{raw_action}'")
            continue
        result[action][key].append(str(value).strip())
    if problems:
        show_warning_message(
            "Filters File",
            "Rows ignored: " + "; ".join(sorted(problems)),
        )
    return result


def build_filter_lists(
    *sources: dict[str, dict[str, list[str]]] | None,
) -> dict[str, dict[str, list[str]]]:
    """Combine CONFIG values with any number of {action: {key: [values]}}
    sources (command line, filters file), de-duplicated per list."""
    combined = _empty_filter_lists()
    for action in FILTER_ACTIONS:
        for key, spec in FILTER_DEFINITIONS.items():
            values = list(spec[f"{action}_values"])
            for src in sources:
                values += ((src or {}).get(action) or {}).get(key) or []
            # De-duplicate on the normalized value, keeping the first spelling.
            unique: dict[str, str] = {}
            for v in values:
                unique.setdefault(_normalize_value(v), v)
            combined[action][key] = list(unique.values())
    return combined


def apply_filters(
    df: pd.DataFrame,
    filters: dict[str, dict[str, list[str]]],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Apply include lists, then exclude lists, to the raw input rows.

    Returns
    -------
    kept : DataFrame
        Input rows that remain for analysis.
    removed : DataFrame
        Removed input rows (original columns) plus a "Removed By" column.
    log : DataFrame
        Audit table (see ``FILTER_LOG_COLUMNS``).  For include filters each
        value shows the rows it matched, and a summary line per filter shows
        the rows removed for not matching any listed value.
    """
    list_src = net_src = None
    try:
        list_src = find_matching_column(df.columns, COLUMN_ALIASES["list_value"])
        net_src = find_matching_column(df.columns, COLUMN_ALIASES["net_sales"])
    except KeyError:
        pass

    def _sum(mask: pd.Series, col: str | None) -> float:
        if col is None:
            return np.nan
        return float(pd.to_numeric(df.loc[mask, col], errors="coerce").sum())

    def _columns_for(spec: dict) -> list[str]:
        if spec["match_all_columns"]:
            return find_all_matching_columns(df.columns, spec["aliases"])
        try:
            return [find_matching_column(df.columns, spec["aliases"])]
        except KeyError:
            return []

    # Each filter column is normalized once and reused for every value.
    normalized_cache: dict[str, pd.Series] = {}

    def _matches(columns: list[str], value: str) -> pd.Series:
        hit = pd.Series(False, index=df.index)
        for col in columns:
            if col not in normalized_cache:
                normalized_cache[col] = _normalize_series(df[col])
            hit |= _value_matches(normalized_cache[col], value)
        return hit

    removed_mask = pd.Series(False, index=df.index)
    reasons = pd.Series("", index=df.index, dtype=object)
    log_rows: list[dict] = []

    def _record_removal(mask: pd.Series, label: str) -> pd.Series:
        """Mark *mask* rows as removed; return rows newly removed."""
        new = mask & ~removed_mask
        existing = reasons.loc[mask]
        reasons.loc[mask] = existing.where(existing == "", existing + "; ") + label
        removed_mask.loc[mask] = True
        return new

    # -- 1. Include lists -----------------------------------------------------
    missing_include: list[str] = []
    for key, values in filters["include"].items():
        if not values:
            continue
        spec = FILTER_DEFINITIONS[key]
        columns = _columns_for(spec)
        if not columns:
            missing_include.append(spec["label"])
            continue
        keep = pd.Series(False, index=df.index)
        for value in values:
            hit = _matches(columns, value)
            keep |= hit
            log_rows.append({
                "Action": "Include", "Filter": spec["label"], "Value": value,
                "Column(s) Checked": ", ".join(map(str, columns)),
                "Status": "Matched" if hit.any() else "No matching rows",
                "Rows Matched": int(hit.sum()),
                "Rows Removed": np.nan,
                "List Value Removed": np.nan, "Net Sales Removed": np.nan,
            })
        new = _record_removal(~keep, f"Not in {spec['label']} include list")
        log_rows.append({
            "Action": "Include", "Filter": spec["label"],
            "Value": "(rows not matching any include value)",
            "Column(s) Checked": ", ".join(map(str, columns)),
            "Status": "Applied",
            "Rows Matched": np.nan,
            "Rows Removed": int(new.sum()),
            "List Value Removed": _sum(new, list_src),
            "Net Sales Removed": _sum(new, net_src),
        })

    if missing_include:
        message = (
            "Include filters were requested but the input file has no matching "
            "column: " + ", ".join(missing_include) + ".  Remove these include "
            "values or use a file that contains the column."
        )
        show_warning_message("Include Filters", message)
        raise ValueError(message)

    # -- 2. Exclude lists -----------------------------------------------------
    missing_exclude: list[str] = []
    for key, values in filters["exclude"].items():
        if not values:
            continue
        spec = FILTER_DEFINITIONS[key]
        columns = _columns_for(spec)
        if not columns:
            missing_exclude.append(spec["label"])
            log_rows.extend(
                {
                    "Action": "Exclude", "Filter": spec["label"], "Value": value,
                    "Column(s) Checked": "",
                    "Status": "Column not in input file - not applied",
                    "Rows Matched": 0, "Rows Removed": 0,
                    "List Value Removed": 0.0, "Net Sales Removed": 0.0,
                }
                for value in values
            )
            continue
        for value in values:
            hit = _matches(columns, value)
            # Rows Removed counts only rows not already removed by an earlier
            # filter, so the Rows Removed column sums to the total removed.
            new = _record_removal(hit, f"Excluded {spec['label']}: {value}")
            log_rows.append({
                "Action": "Exclude", "Filter": spec["label"], "Value": value,
                "Column(s) Checked": ", ".join(map(str, columns)),
                "Status": "Applied" if hit.any() else "No matching rows",
                "Rows Matched": int(hit.sum()),
                "Rows Removed": int(new.sum()),
                "List Value Removed": _sum(new, list_src),
                "Net Sales Removed": _sum(new, net_src),
            })

    if missing_exclude:
        show_warning_message(
            "Exclude Filters",
            "These exclude filters were requested but the input file has no "
            "matching column, so they were NOT applied: "
            + ", ".join(sorted(set(missing_exclude))),
        )

    removed = df.loc[removed_mask].copy()
    removed.insert(0, REMOVED_BY_COL, reasons.loc[removed_mask])
    log = pd.DataFrame(log_rows, columns=FILTER_LOG_COLUMNS)

    kept = df.loc[~removed_mask].copy()
    if log_rows:
        print(f"Filters removed {int(removed_mask.sum())} of {len(df)} input rows.")
    if kept.empty:
        raise ValueError("No input rows remain after applying include / exclude filters.")
    return kept, removed, log


# ---------------------------------------------------------------------------
# Data cleaning, aggregation & standardization
# ---------------------------------------------------------------------------

def find_product_group_column(columns: pd.Index) -> str | None:
    """Return the input's product-group column, or None if there is none."""
    try:
        return find_matching_column(
            columns, FILTER_DEFINITIONS["product_group"]["aliases"]
        )
    except KeyError:
        return None


def standardize_columns(
    df: pd.DataFrame,
    *,
    product_group_column: str | None = None,
) -> pd.DataFrame:
    """Select the analysis columns and rename them to standard names.

    When *product_group_column* is given it is carried through as
    ``PRODUCT_GROUP_COL`` (blank values become ``UNKNOWN_PRODUCT_GROUP``).
    """
    source_columns = [
        find_matching_column(df.columns, COLUMN_ALIASES[key])
        for key in STANDARD_COLUMN_NAMES
    ]
    out = df[source_columns].copy()
    out.columns = list(STANDARD_COLUMN_NAMES.values())
    if product_group_column is not None:
        out[PRODUCT_GROUP_COL] = _clean_labels(
            df[product_group_column], UNKNOWN_PRODUCT_GROUP, blanks=("", "nan", "None")
        )
    return out


def _clean_labels(
    values: pd.Series,
    placeholder: str,
    blanks: tuple[str, ...] = ("", "nan", "None", "--"),
) -> pd.Series:
    """Strip whitespace and replace blank / missing labels with *placeholder*."""
    cleaned = values.fillna("").astype(str).str.strip()
    return cleaned.mask(cleaned.isin(blanks), placeholder)


def normalize_customer_names(names: pd.Series) -> pd.Series:
    """Strip whitespace and replace blank / missing names with a placeholder."""
    return _clean_labels(names, UNKNOWN_CUSTOMER)


def add_discount_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate Discount Amount and Discount % from list value and net sales.

    Discount % = 1 - Net / List.  It is NaN when List Value is not positive,
    so those records are removed by the threshold filter.
    """

    out = df.copy()
    out[DISCOUNT_AMOUNT_COL] = out[LIST_COL] - out[NET_COL]
    out[DISCOUNT_COL] = 1.0 - out[NET_COL] / out[LIST_COL].where(out[LIST_COL] > 0)
    return out


def aggregate_by_customer(df: pd.DataFrame) -> pd.DataFrame:
    """Collapse rows sharing the same customer/group name (and, when the
    Product Group column is present, the same product group) into one record.

    * List Value          -> sum
    * Net Sales           -> sum
    * Records Aggregated  -> number of source rows combined

    Discount % is recalculated afterwards from the summed totals (see
    ``add_discount_columns``), giving a list-value-weighted discount.

    Rows named ``UNKNOWN_CUSTOMER`` are passed through individually.
    """

    keys = [CUSTOMER_COL]
    if PRODUCT_GROUP_COL in df.columns:
        keys.append(PRODUCT_GROUP_COL)

    unknown_mask = df[CUSTOMER_COL] == UNKNOWN_CUSTOMER
    known = df.loc[~unknown_mask]

    grouped = (
        known.groupby(keys, sort=False, as_index=False)
        .agg(**{
            LIST_COL: (LIST_COL, "sum"),
            NET_COL: (NET_COL, "sum"),
            RECORDS_AGGREGATED_COL: (NET_COL, "size"),
        })
    )
    unknown = df.loc[unknown_mask].assign(**{RECORDS_AGGREGATED_COL: 1})

    columns = keys + [LIST_COL, NET_COL, RECORDS_AGGREGATED_COL]
    return pd.concat([grouped[columns], unknown[columns]], ignore_index=True)


def clean_input_data(
    df: pd.DataFrame,
    *,
    aggregate: bool = True,
    by_product_group: bool = False,
) -> pd.DataFrame:
    """Standardize, aggregate, validate, and filter the raw input for analysis.

    Steps:
      1. Map source columns to standard names and normalize customer names.
      2. Coerce list value and net sales to numeric and drop rows where either
         is non-numeric or missing.  (Negative rows such as credits are kept
         at this stage so they net off within their customer group.)
      3. When *aggregate* is True, combine rows sharing the same customer /
         group name -- and the same product group when *by_product_group* is
         True.  This happens before thresholds so that small member accounts
         count toward their group.
      4. Calculate Discount Amount and Discount % from the (aggregated) totals.
      5. Keep records with List Value >= ``min_list_value`` (and > 0) and
         ``min_discount_pct`` <= Discount % <= ``max_discount_pct``.
      6. Report what was dropped / combined.
    """

    pg_source = find_product_group_column(df.columns) if by_product_group else None
    if by_product_group and pg_source is None:
        show_warning_message(
            "Product Group",
            "No product-group column was found in the input file, so the "
            "analysis is run across all product groups combined.",
        )
    cleaned = standardize_columns(df, product_group_column=pg_source)
    input_rows = len(cleaned)
    cleaned[CUSTOMER_COL] = normalize_customer_names(cleaned[CUSTOMER_COL])

    # -- Remove non-numeric rows -------------------------------------------
    cleaned[LIST_COL] = pd.to_numeric(cleaned[LIST_COL], errors="coerce")
    cleaned[NET_COL] = pd.to_numeric(cleaned[NET_COL], errors="coerce")
    numeric_mask = cleaned[LIST_COL].notna() & cleaned[NET_COL].notna()
    non_numeric_count = int((~numeric_mask).sum())
    cleaned = cleaned.loc[numeric_mask]

    # -- Aggregate to one record per customer / group ----------------------
    if aggregate:
        cleaned = aggregate_by_customer(cleaned)
    cleaned = add_discount_columns(cleaned)
    record_count = len(cleaned)

    # -- Apply thresholds --------------------------------------------------
    # NaN discounts (List Value <= 0) fail these comparisons and are excluded.
    list_ok = cleaned[LIST_COL] > 0
    min_value = pd.Series(float(CONFIG["min_list_value"]), index=cleaned.index)
    overrides = {
        _normalize_value(k): float(v)
        for k, v in CONFIG["min_list_value_by_product_group"].items()
    }
    overrides_applied = bool(overrides) and PRODUCT_GROUP_COL in cleaned.columns
    if overrides_applied:
        mapped = cleaned[PRODUCT_GROUP_COL].map(_normalize_value).map(overrides)
        min_value = mapped.fillna(min_value)
    size_ok = cleaned[LIST_COL] >= min_value
    disc_ok = cleaned[DISCOUNT_COL].between(
        CONFIG["min_discount_pct"], CONFIG["max_discount_pct"]
    )
    valid_mask = list_ok & size_ok & disc_ok
    no_list_count = int((~list_ok).sum())
    low_value_count = int((list_ok & ~size_ok).sum())
    disc_range_count = int((list_ok & size_ok & ~disc_ok).sum())
    cleaned = cleaned.loc[valid_mask].reset_index(drop=True)

    # -- Reporting ---------------------------------------------------------
    summary_lines = [f"Input rows: {input_rows}"]
    if non_numeric_count:
        summary_lines.append(f"Rows dropped (non-numeric values): {non_numeric_count}")
    if aggregate:
        unit = (
            "Customer / group x product group records"
            if PRODUCT_GROUP_COL in cleaned.columns
            else "Customers / groups"
        )
        summary_lines.append(f"{unit} after aggregation: {record_count}")
    summary_lines += [
        f"Records dropped (List Value <= 0): {no_list_count}",
        f"Records dropped (List Value < {CONFIG['min_list_value']:,}"
        + (" or product-group minimum" if overrides_applied else "")
        + f"): {low_value_count}",
        "Records dropped (Discount outside "
        f"{CONFIG['min_discount_pct']:.0%} to {CONFIG['max_discount_pct']:.0%}): "
        f"{disc_range_count}",
        f"Records analysed: {len(cleaned)}",
    ]
    summary = "\n".join(summary_lines)

    if cleaned.empty:
        show_warning_message("No Valid Data", f"No valid data remains.\n\n{summary}")
        raise ValueError("No valid rows found after cleaning input data.")

    if non_numeric_count:
        show_info_message("Data Cleaning Complete", summary)
    else:
        print(summary)

    return cleaned


# ---------------------------------------------------------------------------
# Trend fitting
# ---------------------------------------------------------------------------

def fit_semilog_trend(df: pd.DataFrame) -> tuple[np.ndarray, float, float, float]:
    """Fit Discount % = slope * log10(List Value) + intercept.

    Returns
    -------
    trend_values : ndarray
        Predicted Discount % (fraction), one per row of *df*.
    residual_std : float
        Standard deviation of residuals in discount fraction units (ddof=1);
        0.0 when there are too few points to estimate it.
    slope, intercept : float
        Fitted coefficients (slope = change in discount per 10x list value).
    """
    if len(df) < 2:
        raise ValueError("At least two rows are required to fit a trend line.")

    log_x = np.log10(df[X_COL].to_numpy(dtype=float))
    y = df[DISCOUNT_COL].to_numpy(dtype=float)

    slope, intercept = np.polyfit(log_x, y, 1)
    trend = slope * log_x + intercept
    residual_std = float(np.std(y - trend, ddof=1)) if len(df) > 2 else 0.0

    return trend, residual_std, float(slope), float(intercept)


# ---------------------------------------------------------------------------
# List-value binning
# ---------------------------------------------------------------------------

def build_list_value_bins(
    value_series: pd.Series,
) -> tuple[pd.Series, pd.DataFrame]:
    """Assign each row to a quantile-based list-value bin.

    Returns
    -------
    bin_labels : Series
        Maps each row index to its bin name (``"Bin 1"``, ``"Bin 2"``, ...).
    bin_summary : DataFrame
        *List Value Bin* and *List Value Bin Range* columns.
    """
    bin_count = min(CONFIG["max_list_value_bins"], len(value_series))

    if bin_count <= 1:
        bin_labels = pd.Series("Bin 1", index=value_series.index)
    else:
        rank_values = value_series.rank(method="first")
        bin_numbers = pd.qcut(rank_values, q=bin_count, labels=False)
        bin_labels = bin_numbers.astype(int).add(1).map(lambda v: f"Bin {v}")

    bin_summary = (
        pd.DataFrame({"List Value Bin": bin_labels, "List Value": value_series})
        .groupby("List Value Bin", sort=True)["List Value"]
        .agg(["min", "max"])
        .reset_index()
    )
    bin_summary["List Value Bin Range"] = bin_summary.apply(
        lambda row: f"${row['min']:,.0f} to ${row['max']:,.0f}",
        axis=1,
    )

    return bin_labels, bin_summary[["List Value Bin", "List Value Bin Range"]]


# ---------------------------------------------------------------------------
# Within-bin outlier detection
# ---------------------------------------------------------------------------

def calculate_robust_z_scores(series: pd.Series) -> pd.Series:
    """Return Modified Z-scores based on the Median Absolute Deviation (MAD)."""
    median_val = series.median()
    mad = np.median(np.abs(series - median_val))

    if mad == 0 or np.isnan(mad):
        return pd.Series(0.0, index=series.index)

    return 0.6745 * (series - median_val) / mad


def append_bin_outlier_metrics(analysis_df: pd.DataFrame) -> pd.DataFrame:
    """Add list-value bin labels and per-bin Discount % outlier diagnostics.

    Three independent methods flag discount outliers within each bin:
    robust Z-score, robust IQR, and percentile extremes.  Bins with fewer
    than ``min_records_per_bin`` peers suppress all outlier flags.
    """
    enriched = analysis_df.copy()
    bin_labels, bin_summary = build_list_value_bins(enriched[X_COL])
    enriched["List Value Bin"] = bin_labels
    enriched = enriched.merge(bin_summary, on="List Value Bin", how="left")

    enriched["Bin Peer Count"] = 0
    enriched["Bin Median Discount %"] = np.nan
    enriched["Bin Robust Z Score"] = 0.0
    enriched["Bin Robust Z Score Outlier"] = False
    enriched["Bin IQR Lower Bound"] = np.nan
    enriched["Bin IQR Upper Bound"] = np.nan
    enriched["Bin Robust IQR Outlier"] = False
    enriched["Bin Percentile Rank"] = np.nan
    enriched["Bin Lower Percentile Threshold"] = np.nan
    enriched["Bin Upper Percentile Threshold"] = np.nan
    enriched["Bin Percentile Extreme"] = False
    enriched["Bin Outlier Methods Triggered"] = 0
    enriched["Bin Outlier Flag"] = False
    enriched["Bin Outlier Methods"] = ""

    for bin_idx in enriched.groupby("List Value Bin").groups.values():
        discounts = enriched.loc[bin_idx, DISCOUNT_COL]
        peer_count = len(discounts)

        q1, q3 = discounts.quantile(0.25), discounts.quantile(0.75)
        iqr = q3 - q1
        lower_bound = q1 - CONFIG["robust_iqr_multiplier"] * iqr
        upper_bound = q3 + CONFIG["robust_iqr_multiplier"] * iqr
        robust_z = calculate_robust_z_scores(discounts)
        pct_ranks = discounts.rank(method="average", pct=True)

        if peer_count < CONFIG["min_records_per_bin"]:
            z_outlier = iqr_outlier = pct_outlier = pd.Series(False, index=discounts.index)
        else:
            z_outlier = robust_z.abs() > CONFIG["robust_zscore_threshold"]
            iqr_outlier = (discounts < lower_bound) | (discounts > upper_bound)
            pct_outlier = (
                (pct_ranks <= CONFIG["percentile_lower"])
                | (pct_ranks >= CONFIG["percentile_upper"])
            )

        methods_triggered = (
            z_outlier.astype(int) + iqr_outlier.astype(int) + pct_outlier.astype(int)
        )
        method_labels = (
            z_outlier.map({True: "Robust Z score, ", False: ""})
            + iqr_outlier.map({True: "Robust IQR, ", False: ""})
            + pct_outlier.map({True: "Percentile extremes, ", False: ""})
        ).str.removesuffix(", ")

        enriched.loc[bin_idx, "Bin Peer Count"] = peer_count
        enriched.loc[bin_idx, "Bin Median Discount %"] = discounts.median()
        enriched.loc[bin_idx, "Bin Robust Z Score"] = robust_z
        enriched.loc[bin_idx, "Bin Robust Z Score Outlier"] = z_outlier
        enriched.loc[bin_idx, "Bin IQR Lower Bound"] = lower_bound
        enriched.loc[bin_idx, "Bin IQR Upper Bound"] = upper_bound
        enriched.loc[bin_idx, "Bin Robust IQR Outlier"] = iqr_outlier
        enriched.loc[bin_idx, "Bin Percentile Rank"] = pct_ranks
        enriched.loc[bin_idx, "Bin Lower Percentile Threshold"] = discounts.quantile(
            CONFIG["percentile_lower"]
        )
        enriched.loc[bin_idx, "Bin Upper Percentile Threshold"] = discounts.quantile(
            CONFIG["percentile_upper"]
        )
        enriched.loc[bin_idx, "Bin Percentile Extreme"] = pct_outlier
        enriched.loc[bin_idx, "Bin Outlier Methods Triggered"] = methods_triggered
        enriched.loc[bin_idx, "Bin Outlier Flag"] = methods_triggered > 0
        enriched.loc[bin_idx, "Bin Outlier Methods"] = method_labels

    return enriched


def append_percentile_outlier_flags(df: pd.DataFrame) -> pd.DataFrame:
    """Add list-value and discount percentile ranks with a quadrant flag.

    Percentile ranks use the PERCENTRANK.INC convention (0-100) and are
    calculated within the segment passed in (i.e. per product group).

    * ``"Outlier - High Discount, Low List Value"``: list-value percentile
      <= low threshold AND discount percentile >= high threshold.
    * ``"Outlier - Low Discount, High List Value"``: list-value percentile
      >= high threshold AND discount percentile <= low threshold.
    """
    enriched = df.copy()
    n = len(enriched)

    def percent_rank(series: pd.Series) -> pd.Series:
        if n <= 1:
            return pd.Series(0.0, index=series.index)
        return ((series.rank(method="min") - 1) / (n - 1) * 100).round(1)

    enriched["List Value Percentile"] = percent_rank(enriched[X_COL])
    enriched["Discount Percentile"] = percent_rank(enriched[DISCOUNT_COL])

    low_thresh = CONFIG["outlier_low_percentile_threshold"]
    high_thresh = CONFIG["outlier_high_percentile_threshold"]

    high_disc_low_value = (
        (enriched["List Value Percentile"] <= low_thresh)
        & (enriched["Discount Percentile"] >= high_thresh)
    )
    low_disc_high_value = (
        (enriched["List Value Percentile"] >= high_thresh)
        & (enriched["Discount Percentile"] <= low_thresh)
    )

    enriched["Outlier Flag"] = ""
    enriched.loc[high_disc_low_value, "Outlier Flag"] = "Outlier - High Discount, Low List Value"
    enriched.loc[low_disc_high_value, "Outlier Flag"] = "Outlier - Low Discount, High List Value"

    return enriched


# ---------------------------------------------------------------------------
# Outlier classification (trend-based)
# ---------------------------------------------------------------------------

def classify_outliers(
    df: pd.DataFrame,
    trend_values: np.ndarray,
    residual_std: float,
) -> pd.DataFrame:
    """Build the full analysis table from the cleaned data and trend fit.

    Trend-band classification (by absolute residual in discount points):
      * **Within 1 SD** -- Normal
      * **Between 1 and 2 SD** -- Lower Concern Outlier
      * **Between 2 and 3 SD** -- High Concern Outlier
      * **Outside 3 SD** -- Extreme Concern Outlier

    ``Discount Position`` is assigned only to points outside 1 SD:
    above trend = "Potentially Over-Discounted" (customer paying less than
    peers of similar size); below trend = "Potentially Under-Discounted".
    """
    out = df.copy()
    out["Trend Discount %"] = trend_values
    # Residual in discount fraction units (0.05 = 5 percentage points).
    out["Discount Gap vs Trend"] = out[DISCOUNT_COL] - out["Trend Discount %"]
    out["SD From Trend"] = (
        out["Discount Gap vs Trend"] / residual_std if residual_std != 0 else 0.0
    )
    # Approximate $ impact of moving the customer to the trend discount.
    out["Net Sales Impact at Trend"] = out["Discount Gap vs Trend"] * out[LIST_COL]
    out["Band"] = "Within 1 SD"
    out["Severity"] = "Normal"
    out["Discount Position"] = (
        "" if CONFIG["band_blank_within_1sd_discount_position"] else "In Line"
    )

    if residual_std != 0:
        abs_resid = out["Discount Gap vs Trend"].abs()
        outside_1sd = abs_resid > residual_std
        beyond_2sd = abs_resid > 2 * residual_std
        beyond_3sd = abs_resid > 3 * residual_std
        between_1_2 = outside_1sd & ~beyond_2sd
        between_2_3 = beyond_2sd & ~beyond_3sd

        out.loc[between_1_2, "Band"] = "Between 1 and 2 SD"
        out.loc[between_2_3, "Band"] = "Between 2 and 3 SD"
        out.loc[beyond_3sd, "Band"] = "Outside 3 SD"
        out.loc[between_1_2, "Severity"] = "Lower Concern Outlier"
        out.loc[between_2_3, "Severity"] = "High Concern Outlier"
        out.loc[beyond_3sd, "Severity"] = "Extreme Concern Outlier"

        out.loc[outside_1sd & (out["Discount Gap vs Trend"] > 0), "Discount Position"] = (
            "Potentially Over-Discounted"
        )
        out.loc[outside_1sd & (out["Discount Gap vs Trend"] < 0), "Discount Position"] = (
            "Potentially Under-Discounted"
        )

    return append_percentile_outlier_flags(append_bin_outlier_metrics(out))


# ---------------------------------------------------------------------------
# Plotting helpers -- band lines
# ---------------------------------------------------------------------------

def get_band_lines(
    trend_values: np.ndarray,
    residual_std: float,
) -> dict[str, np.ndarray]:
    """Return ADDITIVE +/-1 SD, +/-2 SD and +/-3 SD band lines (discount points).

    The y-axis is linear, so bands are trend +/- k*SD (additive) rather
    than multiplicative.  Empty dict when SD is zero.
    """
    if residual_std == 0:
        return {}

    return {
        "+1 SD": trend_values + residual_std,
        "-1 SD": trend_values - residual_std,
        "+2 SD": trend_values + 2 * residual_std,
        "-2 SD": trend_values - 2 * residual_std,
        "+3 SD": trend_values + 3 * residual_std,
        "-3 SD": trend_values - 3 * residual_std,
    }


# ---------------------------------------------------------------------------
# Plotting helpers -- scatter-label placement
# ---------------------------------------------------------------------------

def _generate_candidate_offsets(
    num_angles: int = 12,
    radii: tuple[float, ...] = (14.0, 24.0, 36.0),
) -> np.ndarray:
    """Generate candidate label offsets arranged in concentric rings.

    Returns an ``(num_angles * len(radii), 2)`` array of ``(dx, dy)`` pairs
    in display-point units (default: 36 candidates).  Successive rings are
    rotated by half an angular step so that inner and outer candidates do
    not share the same radial line, improving coverage in dense clusters.

    Parameters
    ----------
    num_angles : int
        Number of evenly-spaced directions per ring.
    radii : tuple[float, ...]
        Distance of each ring from the anchor, in display points.
    """
    offsets: list[tuple[float, float]] = []
    for ring_idx, radius in enumerate(radii):
        phase = (ring_idx * np.pi) / num_angles
        for i in range(num_angles):
            angle = phase + 2.0 * np.pi * i / num_angles
            offsets.append((radius * np.cos(angle), radius * np.sin(angle)))
    return np.asarray(offsets)


# Candidate label offsets are the same for every chart, so build them once.
_CANDIDATE_OFFSETS: np.ndarray = _generate_candidate_offsets()


def _offsets_to_plot_space(
    anchor_x: float,
    anchor_y: float,
    offsets: np.ndarray,
    x_span: float,
    y_span: float,
) -> tuple[np.ndarray, np.ndarray]:
    """Map point-unit ``(dx, dy)`` offsets to approximate plot-space positions
    (log10 list value, discount fraction) around an anchor point."""
    return (
        anchor_x + (offsets[:, 0] / _PLOT_WIDTH_PTS) * x_span,
        anchor_y + (offsets[:, 1] / _PLOT_HEIGHT_PTS) * y_span,
    )


def _select_best_offset(
    cx: np.ndarray,
    cy: np.ndarray,
    placed_labels: list[tuple[float, float]],
    data_points: np.ndarray,
    candidate_offsets: np.ndarray,
) -> int:
    """Return the index of the candidate offset that best avoids clutter.

    Scoring (lower is better) combines three terms:

    * **Label repulsion** -- sum of inverse distances to already-placed
      labels.  Heavily weighted to prevent label-on-label overlap.
    * **Point repulsion** -- sum of inverse distances to all scatter-plot
      data points.  Moderately weighted so labels avoid obscuring data.
    * **Anchor distance** -- slight preference for keeping the label close
      to its data point to maintain readability.

    All position comparisons happen in plot space (log10 list value on x,
    linear discount fraction on y), so scoring follows the chart layout.

    Parameters
    ----------
    cx, cy : ndarray, shape (C,)
        Plot-space positions of each candidate (see
        :func:`_offsets_to_plot_space`).
    placed_labels : list[tuple[float, float]]
        Plot-space positions of labels already committed to the chart.
    data_points : ndarray, shape (N, 2)
        Plot-space coordinates of every scatter-plot point.
    candidate_offsets : ndarray, shape (C, 2)
        Point-unit ``(dx, dy)`` candidates matching *cx*, *cy*.
    """
    n_candidates = len(candidate_offsets)

    # -- Label repulsion (sum of inverse distances to placed labels) --------
    label_penalty = np.zeros(n_candidates)
    for lx, ly in placed_labels:
        # Labels are much wider than tall, so horizontal separation counts
        # for less than vertical separation when scoring overlap.
        dists = np.hypot((cx - lx) / 4.0, cy - ly)
        label_penalty += 1.0 / np.maximum(dists, 1e-9)

    # -- Point repulsion (sum of inverse distances to data points) ----------
    point_penalty = np.zeros(n_candidates)
    if data_points.size > 0:
        # Broadcast: (C,1) - (1,N) -> (C,N) pairwise distance matrix.
        dx = cx[:, np.newaxis] - data_points[np.newaxis, :, 0]
        dy = cy[:, np.newaxis] - data_points[np.newaxis, :, 1]
        dists = np.hypot(dx, dy)
        point_penalty = np.sum(1.0 / np.maximum(dists, 1e-9), axis=1)

    # -- Anchor distance (prefer positions closer to the data point) --------
    anchor_dist = np.hypot(candidate_offsets[:, 0], candidate_offsets[:, 1])

    # Weighted composite score.
    scores = 8.0 * label_penalty + 1.5 * point_penalty + 0.01 * anchor_dist
    return int(np.argmin(scores))


def annotate_outliers(
    ax,
    plot_df: pd.DataFrame,
    labeled_points_df: pd.DataFrame,
    label_color: str = "#d62728",
) -> None:
    """Annotate outlier points with customer/group names, minimizing overlap.

    Positions are scored in plot space (log10 list value, linear discount),
    largest absolute discount gap first.
    """
    data_points = to_plot_space(plot_df[X_COL], plot_df[DISCOUNT_COL])
    x_span, y_span = np.maximum(np.ptp(data_points, axis=0), 1e-6)

    placed_labels: list[tuple[float, float]] = []
    x_lo, x_hi = np.log10(ax.get_xlim())
    y_lo, y_hi = ax.get_ylim()
    x_margin, y_margin = 0.05 * (x_hi - x_lo), 0.02 * (y_hi - y_lo)

    prioritized = labeled_points_df.sort_values(
        "Discount Gap vs Trend",
        ascending=False,
        key=abs,
    )

    for name, x_val, y_val in prioritized[[CUSTOMER_COL, X_COL, DISCOUNT_COL]].itertuples(
        index=False, name=None
    ):
        anchor_x, anchor_y = float(np.log10(x_val)), float(y_val)
        cand_x, cand_y = _offsets_to_plot_space(
            anchor_x, anchor_y, _CANDIDATE_OFFSETS, x_span, y_span
        )

        # Drop candidate positions that would push the label outside the
        # plot area (e.g. over the chart title); keep all if none remain.
        inside = (
            (cand_x > x_lo + x_margin) & (cand_x < x_hi - x_margin)
            & (cand_y > y_lo + y_margin) & (cand_y < y_hi - y_margin)
        )
        if not inside.any():
            inside[:] = True

        usable = _CANDIDATE_OFFSETS[inside]
        best = _select_best_offset(
            cand_x[inside], cand_y[inside], placed_labels, data_points, usable
        )
        placed_labels.append((float(cand_x[inside][best]), float(cand_y[inside][best])))
        offset = (float(usable[best, 0]), float(usable[best, 1]))

        ax.annotate(
            name,
            (x_val, y_val),
            xytext=offset,
            textcoords="offset points",
            fontsize=9,
            color=label_color,
            bbox={
                "boxstyle": "round,pad=0.2",
                "fc": "white",
                "ec": "none",
                "alpha": 0.75,
            },
            arrowprops={
                "arrowstyle": "-",
                "color": "#999999",
                "lw": 0.8,
                "alpha": 0.7,
            },
        )


# ---------------------------------------------------------------------------
# Main chart & analysis
# ---------------------------------------------------------------------------

def analyse_segment(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Fit the trend and classify one segment (one product group, or all).

    Returns the analysis table and a dict of fit statistics.
    """
    trend_values, residual_std, slope, intercept = fit_semilog_trend(df)
    analysis_df = classify_outliers(df, trend_values, residual_std)
    stats = {"slope": slope, "intercept": intercept, "residual_std": residual_std}
    return analysis_df, stats


def plot_segment_chart(
    analysis_df: pd.DataFrame,
    residual_std: float,
    chart_title: str,
    *,
    show: bool = True,
    chart_path: Path | None = None,
) -> None:
    """Draw the semi-log discount-band chart for one analysed segment."""
    plot_df = analysis_df.sort_values(X_COL).reset_index(drop=True)

    x_sorted = plot_df[X_COL].to_numpy(dtype=float)
    trend_sorted = plot_df["Trend Discount %"].to_numpy(dtype=float)

    if residual_std != 0:
        label_thresh = float(CONFIG["scatter_label_sd_threshold"])
        labeled_df = plot_df.loc[
            plot_df["Discount Gap vs Trend"].abs() > label_thresh * residual_std
        ]
    else:
        labeled_df = plot_df.iloc[0:0]

    fig, ax = plt.subplots(figsize=FIGSIZE)

    scatter_styles = {
        "Within 1 SD": {"alpha": 0.75, "s": 60, "color": "#1f77b4"},
        "Between 1 and 2 SD": {
            "alpha": 0.85, "s": 75, "color": "#ff7f0e",
            "edgecolors": "black", "linewidths": 0.5,
        },
        "Between 2 and 3 SD": {
            "alpha": 0.95, "s": 90, "color": "#d62728", "edgecolors": "black",
        },
        "Outside 3 SD": {
            "alpha": 1.0, "s": 110, "color": "#7f0000", "edgecolors": "black",
            "linewidths": 1.0,
        },
    }
    for band, style in scatter_styles.items():
        band_df = plot_df.loc[plot_df["Band"] == band]
        if not band_df.empty:
            ax.scatter(band_df[X_COL], band_df[DISCOUNT_COL], label=band, **style)

    ax.plot(
        x_sorted, trend_sorted,
        color="black", linewidth=2.2, label="Semi-log trend line",
    )

    bands = get_band_lines(trend_sorted, residual_std)
    if bands:
        ax.plot(x_sorted, bands["+1 SD"], ":", color="#2ca02c", lw=1.8, label="+/- 1 SD")
        ax.plot(x_sorted, bands["-1 SD"], ":", color="#2ca02c", lw=1.8)
        ax.plot(x_sorted, bands["+2 SD"], ":",  color="#ff7f0e", lw=1.8, label="+/- 2 SD")
        ax.plot(x_sorted, bands["-2 SD"], ":",  color="#ff7f0e", lw=1.8)
        ax.plot(x_sorted, bands["+3 SD"], ":", color="#d62728", lw=1.8, label="+/- 3 SD")
        ax.plot(x_sorted, bands["-3 SD"], ":", color="#d62728", lw=1.8)

    # List value on a log axis; discount on a LINEAR axis (zero / negative
    # discounts cannot be shown on a log scale).
    ax.set_xscale("log")
    ax.yaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=0))
    ax.axhline(0, color="grey", lw=0.8)
    # Keep the y-range to the data (bands can otherwise extend past 100 %).
    y_vals = plot_df[DISCOUNT_COL]
    y_pad = 0.10  # head-room for outlier labels
    ax.set_ylim(min(y_vals.min(), 0) - y_pad, min(max(y_vals.max(), 0) + y_pad, 1.05))

    if not labeled_df.empty:
        annotate_outliers(ax, plot_df, labeled_df)

    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel("Total List Value (log scale)", fontsize=15)
    ax.set_ylabel("Total Discount from List (%)", fontsize=15)
    ax.grid(True, which="both", linestyle=":", alpha=0.5)
    ax.legend()
    fig.tight_layout()
    if chart_path is not None:
        fig.savefig(chart_path, dpi=150)
    if show:
        plt.show()
    plt.close(fig)


def _safe_filename(text: str) -> str:
    keep = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in str(text))
    return keep.strip("_")[:60] or "segment"


def _chart_copy_path(
    save_chart: Path,
    product_group: str | None,
    used: set[Path],
) -> Path:
    """Path for the --save-chart copy of one chart.

    In product-group mode the product group is appended to the file name; a
    numeric suffix is added if two product groups sanitise to the same name
    (*used* holds the paths already written in this run).
    """
    suffix = save_chart.suffix or ".png"
    if product_group is None:
        target = save_chart.with_suffix(suffix)
    else:
        base = f"{save_chart.stem}_{_safe_filename(product_group)}"
        target = save_chart.with_name(f"{base}{suffix}")
        counter = 2
        while target in used:
            target = save_chart.with_name(f"{base}_{counter}{suffix}")
            counter += 1
    used.add(target)
    return target


def run_analysis(
    cleaned_df: pd.DataFrame,
    chart_title: str,
    *,
    show: bool = True,
    chart_dir: Path,
    save_chart: Path | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[tuple[str, Path]]]:
    """Analyse each product group separately (or all records as one segment
    when there is no Product Group column).

    Returns
    -------
    analysis_df : DataFrame
        All records; product groups with too few records are included with
        Severity "Not Assessed".
    summary_df : DataFrame
        One row per product group with totals and fit statistics.
    charts : list of (product group, PNG path)
    """
    df = cleaned_df.copy()
    segmented = PRODUCT_GROUP_COL in df.columns
    if not segmented:
        df[PRODUCT_GROUP_COL] = ALL_PRODUCTS_LABEL

    min_records = max(int(CONFIG["min_records_per_product_group"]), 3)

    # Largest product groups (by list value) first.
    order = (
        df.groupby(PRODUCT_GROUP_COL)[LIST_COL].sum()
        .sort_values(ascending=False).index.tolist()
    )

    results: list[pd.DataFrame] = []
    summary_rows: list[dict] = []
    charts: list[tuple[str, Path]] = []
    saved_charts: set[Path] = set()

    for pg in order:
        seg = df.loc[df[PRODUCT_GROUP_COL] == pg].reset_index(drop=True)
        total_list = float(seg[LIST_COL].sum())
        total_net = float(seg[NET_COL].sum())
        row = {
            PRODUCT_GROUP_COL: pg,
            "Status": "",
            "Records": len(seg),
            LIST_COL: total_list,
            NET_COL: total_net,
            "Overall Discount %": 1 - total_net / total_list if total_list else np.nan,
            "Median Record Discount %": float(seg[DISCOUNT_COL].median()),
            "Trend Slope (pts per 10x List Value)": np.nan,
            "Residual SD (pts)": np.nan,
            "Extreme Concern Outliers": 0,
            "High Concern Outliers": 0,
            "Lower Concern Outliers": 0,
            "Potentially Over-Discounted": 0,
            "Potentially Under-Discounted": 0,
            "Net Sales Upside if Over-Discounted Moved to Trend": 0.0,
        }

        if len(seg) < min_records:
            row["Status"] = f"Not assessed - fewer than {min_records} records"
            seg = seg.assign(Band="Not Assessed", Severity="Not Assessed")
            results.append(seg)
            summary_rows.append(row)
            continue

        analysis, stats = analyse_segment(seg)
        results.append(analysis)

        title = chart_title if not segmented else f"{chart_title} - {pg}"
        png = chart_dir / f"chart_{len(charts) + 1:02d}_{_safe_filename(pg)}.png"
        plot_segment_chart(analysis, stats["residual_std"], title, show=show, chart_path=png)
        charts.append((pg, png))
        if save_chart is not None:
            shutil.copyfile(
                png, _chart_copy_path(save_chart, pg if segmented else None, saved_charts)
            )

        severity_counts = analysis["Severity"].value_counts()
        position_counts = analysis["Discount Position"].value_counts()
        over = analysis["Discount Position"] == "Potentially Over-Discounted"
        row.update({
            "Status": "Analysed",
            "Trend Slope (pts per 10x List Value)": stats["slope"],
            "Residual SD (pts)": stats["residual_std"],
            "Extreme Concern Outliers": int(
                severity_counts.get("Extreme Concern Outlier", 0)
            ),
            "High Concern Outliers": int(severity_counts.get("High Concern Outlier", 0)),
            "Lower Concern Outliers": int(severity_counts.get("Lower Concern Outlier", 0)),
            "Potentially Over-Discounted": int(over.sum()),
            "Potentially Under-Discounted": int(
                position_counts.get("Potentially Under-Discounted", 0)
            ),
            "Net Sales Upside if Over-Discounted Moved to Trend": float(
                analysis.loc[over, "Net Sales Impact at Trend"].sum()
            ),
        })
        summary_rows.append(row)

    analysis_df = pd.concat(results, ignore_index=True)
    summary_df = pd.DataFrame(summary_rows)
    if not (summary_df["Status"] == "Analysed").any():
        raise ValueError(
            f"No product group has at least {min_records} records to analyse."
        )
    return analysis_df, summary_df, charts


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def _move_after(columns: list[str], to_move: list[str], anchor: str) -> list[str]:
    """Return *columns* with *to_move* (those present) placed directly after
    *anchor*, in the order given.  Unchanged if *anchor* is absent."""
    if anchor not in columns:
        return columns
    moving = [c for c in to_move if c in columns and c != anchor]
    remaining = [c for c in columns if c not in moving]
    pos = remaining.index(anchor) + 1
    return remaining[:pos] + moving + remaining[pos:]


def order_export_columns(columns: list[str]) -> list[str]:
    """Apply the Excel export column order.

    * Records Aggregated directly after Customer Name (only present when
      aggregation is enabled).
    * SD From Trend, Band, Severity and Discount Position directly after
      Discount %.
    * List Value Percentile, Discount Percentile and Outlier Flag directly
      after Discount Position (the "Potentially Over-/Under-Discounted"
      column).
    """
    columns = _move_after(
        columns,
        [RECORDS_AGGREGATED_COL],
        CUSTOMER_COL,
    )
    columns = _move_after(
        columns,
        ["SD From Trend", "Band", "Severity", "Discount Position"],
        DISCOUNT_COL,
    )
    columns = _move_after(
        columns,
        ["List Value Percentile", "Discount Percentile", "Outlier Flag"],
        "Discount Position",
    )
    # Product Group first, so Records Aggregated stays directly after
    # Customer Name.
    if PRODUCT_GROUP_COL in columns:
        columns = [PRODUCT_GROUP_COL] + [c for c in columns if c != PRODUCT_GROUP_COL]
    return columns


def _solid_fill(color: str) -> PatternFill:
    return PatternFill(fill_type="solid", start_color=color, end_color=color)


def _format_sheet(
    ws,
    *,
    max_width: int = 50,
    number_formats: dict[str, str] | None = None,
    severity_fills: dict[str, PatternFill] | None = None,
) -> None:
    """Apply the standard sheet formatting in a single pass over the cells.

    * Header row: fill, bold white font, centred; freeze panes; auto-filter.
    * Number formats by column header (*number_formats*).
    * Optional whole-row fill keyed by the row's "Severity" value.
    * Column widths fitted to the longest value, capped at *max_width*.
    """
    headers = [cell.value for cell in ws[1]]
    header_fill = _solid_fill(CONFIG["header_fill_color"])
    header_font = Font(name="Arial", bold=True, color=CONFIG["header_font_color"])
    header_alignment = Alignment(horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    column_formats = [(number_formats or {}).get(h) for h in headers]
    severity_pos = (
        headers.index("Severity")
        if severity_fills and "Severity" in headers
        else None
    )
    max_lens = [len(str(h)) if h is not None else 0 for h in headers]

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_fill = (
            severity_fills.get(row[severity_pos].value)
            if severity_pos is not None
            else None
        )
        for idx, cell in enumerate(row):
            if cell.value is not None:
                max_lens[idx] = max(max_lens[idx], len(str(cell.value)))
            if column_formats[idx]:
                cell.number_format = column_formats[idx]
            if row_fill is not None:
                cell.fill = row_fill

    for idx, length in enumerate(max_lens, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = min(length + 2, max_width)


def _group_columns(ws, headers_to_group: list[str], *, hidden: bool = False) -> None:
    """Apply a level-1 outline grouping to each listed column (by header).

    Headers not present on the sheet are skipped (e.g. Records Aggregated
    when --no-aggregate is used).  Set *hidden* to True to collapse the
    group when the workbook is opened.
    """
    headers = [cell.value for cell in ws[1]]
    for header in headers_to_group:
        if header not in headers:
            continue
        letter = get_column_letter(headers.index(header) + 1)
        dim = ws.column_dimensions[letter]
        dim.outline_level = 1
        dim.hidden = hidden
    # Put the group's expand / collapse button to the right of the group.
    ws.sheet_properties.outlinePr.summaryRight = True


def _format_summary_sheet(ws) -> None:
    """Product Group Summary: standard formatting plus wrapped headers and
    narrower numeric columns."""
    _format_sheet(ws, number_formats=EXCEL_NUMBER_FORMATS)
    for cell in ws[1]:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 45
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        ws.column_dimensions[letter].width = max(min(ws.column_dimensions[letter].width, 22), 12)
    # Product Group and Status columns (first two columns of the summary).
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 40


def _add_charts_sheet(ws, charts: list[tuple[str, Path]]) -> None:
    """Stack one titled chart image per product group down the sheet."""
    row = 1
    for pg, png in charts:
        ws.cell(row=row, column=1, value=str(pg)).font = Font(bold=True, size=14)
        img = XLImage(str(png))
        scale = 960 / img.width
        img.width, img.height = int(img.width * scale), int(img.height * scale)
        ws.add_image(img, f"A{row + 1}")
        # Default row height is 15 pt (~20 px).
        row += int(img.height / 20) + 4


def prepare_analysis_export(
    analysis_df: pd.DataFrame,
    product_group_order: list[str] | None = None,
) -> pd.DataFrame:
    """Sort and order the columns of the analysis table for export.

    Rows are sorted by product group (largest list value first), then
    most-severe first, then by bin, number of within-bin methods triggered,
    list value, and discount.
    """
    severity_rank = {sev: i for i, sev in enumerate(SEVERITY_ORDER)}
    pg_rank = {pg: i for i, pg in enumerate(product_group_order or [])}

    def _sort_key(col: pd.Series) -> pd.Series:
        if col.name == "Severity":
            return col.map(severity_rank)
        if col.name == PRODUCT_GROUP_COL:
            return col.map(pg_rank).fillna(len(pg_rank))
        if col.name == "List Value Bin":
            # Numeric sort so that "Bin 10" follows "Bin 9".
            return pd.to_numeric(col.str.removeprefix("Bin "), errors="coerce")
        return col

    sort_spec = [
        (PRODUCT_GROUP_COL, True), ("Severity", True), ("List Value Bin", True),
        ("Bin Outlier Methods Triggered", False), (X_COL, True), (DISCOUNT_COL, False),
    ]
    present = [(c, a) for c, a in sort_spec if c in analysis_df.columns]
    export_df = analysis_df.sort_values(
        [c for c, _ in present],
        ascending=[a for _, a in present],
        key=_sort_key,
        na_position="last",
    )
    return export_df[order_export_columns(list(export_df.columns))]


def write_output_workbook(
    output_path: Path,
    analysis_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    charts: list[tuple[str, Path]],
    filter_log: pd.DataFrame | None = None,
    removed_rows: pd.DataFrame | None = None,
) -> None:
    """Write every output sheet in one pass and save the workbook once.

    Sheet order: analysis table, Product Group Summary, Charts, and -- only
    when at least one include / exclude value was requested -- Filters
    Applied and Removed Rows.
    """
    export_df = prepare_analysis_export(
        analysis_df, summary_df[PRODUCT_GROUP_COL].tolist()
    )
    write_filters = filter_log is not None and not filter_log.empty

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name=CONFIG["worksheet_name"])
        summary_df.to_excel(writer, index=False, sheet_name=CONFIG["summary_sheet_name"])
        book = writer.book

        if charts:
            _add_charts_sheet(book.create_sheet(CONFIG["charts_sheet_name"]), charts)

        if write_filters:
            filter_log.to_excel(
                writer, index=False, sheet_name=CONFIG["filters_sheet_name"]
            )
            rows = removed_rows if removed_rows is not None else pd.DataFrame()
            if rows.empty:
                rows = pd.DataFrame({REMOVED_BY_COL: ["No input rows were removed"]})
            rows.to_excel(
                writer, index=False, sheet_name=CONFIG["removed_rows_sheet_name"]
            )
            _format_sheet(
                book[CONFIG["filters_sheet_name"]],
                number_formats=EXCEL_NUMBER_FORMATS,
            )
            # Removed Rows holds the original input columns, so no analysis
            # number formats are applied to it.
            _format_sheet(book[CONFIG["removed_rows_sheet_name"]])

        _format_sheet(
            book[CONFIG["worksheet_name"]],
            max_width=40,
            number_formats=EXCEL_NUMBER_FORMATS,
            severity_fills={
                sev: _solid_fill(color)
                for sev, color in CONFIG["severity_fill_colors"].items()
            },
        )
        # Outline-group the Records Aggregated column (after formatting so
        # the fitted column width is kept).
        _group_columns(book[CONFIG["worksheet_name"]], [RECORDS_AGGREGATED_COL])
        _format_summary_sheet(book[CONFIG["summary_sheet_name"]])
    # Workbook is saved when the ExcelWriter context closes.


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Generate a semi-log discount band analysis plot (Discount % vs "
            "total list value) with trend, standard-deviation bands, and "
            "configurable customer labels."
        ),
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        help="Input CSV or Excel file containing customer, list value, and net sales columns.",
    )
    parser.add_argument("--sheet", default=None, help="Excel sheet name for .xlsx or .xls inputs.")
    parser.add_argument(
        "--output-outliers",
        default=None,
        help="Output Excel filename for the analysis export.",
    )
    parser.add_argument(
        "--save-chart",
        default=None,
        help="Optional PNG filename to save the chart(s).  In product-group "
             "mode one file per product group is saved, with the product "
             "group appended to the name.",
    )
    parser.add_argument(
        "--no-show",
        dest="show",
        action="store_false",
        help="Do not display the chart on screen (useful for batch runs).",
    )
    parser.add_argument(
        "--title",
        default="Discount Band Analysis",
        help="Chart title.",
    )
    parser.add_argument(
        "--combined",
        dest="by_product_group",
        action="store_false",
        default=CONFIG["analyse_by_product_group"],
        help="Analyse each customer/group across all product groups combined "
             "(one trend line) instead of separately for each product group.",
    )
    parser.add_argument(
        "--no-aggregate",
        dest="aggregate",
        action="store_false",
        default=CONFIG["aggregate_by_customer"],
        help="Analyse each input row separately instead of aggregating rows "
             "that share the same customer/group name.",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Show the full error traceback instead of a one-line message.",
    )
    filter_group = parser.add_argument_group(
        "include / exclude filters",
        "Applied to input rows before analysis.  Each option can be repeated; "
        "matching is case-insensitive and supports * and ? wildcards.  "
        "Include lists keep ONLY matching rows; exclude lists are applied "
        "afterwards and always win.",
    )
    filter_group.add_argument(
        "--include-customer-type", action="append", default=[], metavar="VALUE",
        help="Keep only rows with this customer type.",
    )
    filter_group.add_argument(
        "--include-customer", action="append", default=[], metavar="VALUE",
        help="Keep only rows with this customer or group name (checked against "
             "every customer-name column).",
    )
    filter_group.add_argument(
        "--include-product-group", action="append", default=[], metavar="VALUE",
        help="Keep only rows with this product group / product type.",
    )
    filter_group.add_argument(
        "--exclude-customer-type", action="append", default=[], metavar="VALUE",
        help="Remove rows with this customer type.",
    )
    filter_group.add_argument(
        "--exclude-customer", action="append", default=[], metavar="VALUE",
        help="Remove rows with this customer or group name (checked against "
             "every customer-name column).",
    )
    filter_group.add_argument(
        "--exclude-product-group", action="append", default=[], metavar="VALUE",
        help="Remove rows with this product group / product type.",
    )
    filter_group.add_argument(
        "--filters-file", "--exclusions-file", dest="filters_file",
        default=None, metavar="FILE",
        help='CSV or Excel list with columns "Action" (Include / Exclude; '
             'blank = Exclude), "Filter Type" (Customer Type / Customer Name / '
             'Product Group) and "Value".',
    )
    return parser.parse_args()


def main() -> int:
    """Run the discount band analysis workflow."""
    args = parse_args()

    try:
        input_path = Path(args.input_file) if args.input_file else choose_input_file()

        raw_df = load_input_file(input_path, args.sheet)

        cli_filters = {
            action: {
                "customer_type": getattr(args, f"{action}_customer_type"),
                "customer_name": getattr(args, f"{action}_customer"),
                "product_group": getattr(args, f"{action}_product_group"),
            }
            for action in FILTER_ACTIONS
        }
        filters = build_filter_lists(
            cli_filters,
            load_filters_file(Path(args.filters_file)) if args.filters_file else None,
        )
        raw_df, removed_rows, filter_log = apply_filters(raw_df, filters)

        cleaned_df = clean_input_data(
            raw_df, aggregate=args.aggregate, by_product_group=args.by_product_group
        )

        # Resolve the output location BEFORE the (potentially long) analysis
        # and chart display, so a cancelled Save dialog does not discard work.
        output_path = (
            Path(args.output_outliers)
            if args.output_outliers
            else choose_output_file(CONFIG["default_output_filename"])
        )

        # Chart PNGs live in a temporary folder that is removed automatically,
        # including when an error occurs part-way through.
        with tempfile.TemporaryDirectory(prefix="discount_band_charts_") as tmp:
            analysis_df, summary_df, charts = run_analysis(
                cleaned_df,
                args.title,
                show=args.show,
                chart_dir=Path(tmp),
                save_chart=Path(args.save_chart) if args.save_chart else None,
            )
            write_output_workbook(
                output_path, analysis_df, summary_df, charts, filter_log, removed_rows
            )

        extreme_concern = int((analysis_df["Band"] == "Outside 3 SD").sum())
        high_concern = int((analysis_df["Band"] == "Between 2 and 3 SD").sum())
        # Not-assessed rows have no bin flag (NaN), so count explicit Trues.
        bin_outliers = int(analysis_df["Bin Outlier Flag"].eq(True).sum())
        print("Analysis complete.")
        print(f"Output file saved to: {output_path}")
        assessed = analysis_df["Severity"] != "Not Assessed"
        print(
            f"Records assessed: {int(assessed.sum())} "
            f"(not assessed - product group too small: {int((~assessed).sum())})"
        )
        for _, r in summary_df.iterrows():
            print(
                f"  {r[PRODUCT_GROUP_COL]}: {r['Records']} records, "
                f"{r['Extreme Concern Outliers']} extreme concern, "
                f"{r['High Concern Outliers']} high concern - {r['Status']}"
            )
        if not filter_log.empty:
            print(f"Input rows removed by include / exclude filters: {len(removed_rows)}")
        print(f"Extreme concern outliers (>3 SD from trend): {extreme_concern}")
        print(f"High concern outliers (2-3 SD from trend): {high_concern}")
        print(f"Records flagged by at least one within-bin method: {bin_outliers}")
        return 0
    except Exception as exc:
        if args.debug:
            raise
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
