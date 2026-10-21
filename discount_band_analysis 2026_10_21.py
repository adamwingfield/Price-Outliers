"""Generate a discount band analysis chart from customer sales data.

Adapted from ``price_band_analysis.py``.  Instead of comparing average
procedure price against procedure quantity (log-log), this version compares
the total discount from list price (Discount %) against total list value
(customer size measured at list price, independent of the discount given).

The script reads a CSV or Excel file containing a customer (or customer
group) column, a total list-price value column, and a total net-sales
column.  It then:

  1. Standardizes the input columns (see ``COLUMN_ALIASES``).
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
     Each record is classified against +/-1 and +/-2 standard-deviation
     bands (measured in discount percentage points), and within-bin and
     percentile-based outlier checks are run.
  5. Displays the scatter plot (log x-axis, linear % y-axis) and exports the
     full analysis table to a formatted Excel file.

Sign convention: Discount % is POSITIVE when net sales are below list
(e.g. 0.25 = 25 % off list) and NEGATIVE when net sales exceed list.  Note
that this is the opposite sign to the "Net Discount %" column in the
Pricemart extract, which reports (net - list) / list.

Usage:
    python discount_band_analysis.py
    python discount_band_analysis.py input_file.csv
    python discount_band_analysis.py input_file.xlsx --sheet "Data Table"
    python discount_band_analysis.py input_file.xlsx --no-aggregate

Dependencies:
    pandas, numpy, matplotlib, openpyxl (for .xlsx input/output)
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import PercentFormatter
from openpyxl import load_workbook
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

# Derived columns (calculated, never read from the input).
DISCOUNT_COL = "Discount %"            # 1 - Net Sales / List Value
DISCOUNT_AMOUNT_COL = "Discount Amount"  # List Value - Net Sales

# Column used for the x-axis, trend fit, size bins and size percentiles.
X_COL = STANDARD_COLUMN_NAMES["list_value"]

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
    "High Concern Outlier",
    "Lower Concern Outlier",
    "Normal",
]

# Approximate plot-area dimensions (in typographic points) for a (12, 8)-inch
# figure.  Used solely to convert pixel offsets into proportional plot-space
# distances for label-collision scoring.
_PLOT_WIDTH_PTS: float = 650.0
_PLOT_HEIGHT_PTS: float = 450.0

CONFIG: dict = {
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
    "scatter_label_sd_threshold": 2.0,
    # Within-bin outlier-detection parameters
    "robust_zscore_threshold": 3.5,
    "robust_iqr_multiplier": 1.5,
    "percentile_lower": 0.05,
    "percentile_upper": 0.95,
    # Excel output settings
    "default_output_filename": "discount_band_outliers.xlsx",
    "worksheet_name": "Discount Band Analysis",
    "header_fill_color": "002060",
    "header_font_color": "FFFFFF",
    "severity_fill_colors": {
        "Normal": "DDEBF7",
        "Lower Concern Outlier": "FCE4D6",
        "High Concern Outlier": "F8CBAD",
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

# Excel number formats applied to output columns, keyed by column header.
EXCEL_NUMBER_FORMATS: dict[str, str] = {
    **dict.fromkeys(
        ["List Value", "Net Sales", "Discount Amount"],
        "$#,##0",
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

    workbook = pd.read_excel(file_path, sheet_name=None)
    net_aliases = {normalize_column_name(a) for a in COLUMN_ALIASES["net_sales"]}

    for candidate_df in workbook.values():
        sheet_columns = {normalize_column_name(c) for c in candidate_df.columns}
        if net_aliases & sheet_columns:
            return candidate_df

    return next(iter(workbook.values()))


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
# Data cleaning, aggregation & standardization
# ---------------------------------------------------------------------------

def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Select the three analysis columns and rename them to standard names."""
    source_columns = [
        find_matching_column(df.columns, COLUMN_ALIASES[key])
        for key in STANDARD_COLUMN_NAMES
    ]
    out = df[source_columns].copy()
    out.columns = list(STANDARD_COLUMN_NAMES.values())
    return out


def normalize_customer_names(names: pd.Series) -> pd.Series:
    """Strip whitespace and replace blank / missing names with a placeholder."""
    cleaned = names.fillna("").astype(str).str.strip()
    return cleaned.mask(cleaned.isin(["", "nan", "None", "--"]), UNKNOWN_CUSTOMER)


def add_discount_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate Discount Amount and Discount % from list value and net sales.

    Discount % = 1 - Net / List.  It is NaN when List Value is not positive,
    so those records are removed by the threshold filter.
    """
    list_col = STANDARD_COLUMN_NAMES["list_value"]
    net_col = STANDARD_COLUMN_NAMES["net_sales"]

    out = df.copy()
    out[DISCOUNT_AMOUNT_COL] = out[list_col] - out[net_col]
    out[DISCOUNT_COL] = 1.0 - out[net_col] / out[list_col].where(out[list_col] > 0)
    return out


def aggregate_by_customer(df: pd.DataFrame) -> pd.DataFrame:
    """Collapse rows sharing the same customer/group name into one record.

    * List Value          -> sum
    * Net Sales           -> sum
    * Records Aggregated  -> number of source rows combined

    Discount % is recalculated afterwards from the summed totals (see
    ``add_discount_columns``), giving a list-value-weighted discount.

    Rows named ``UNKNOWN_CUSTOMER`` are passed through individually.
    """
    customer_col = STANDARD_COLUMN_NAMES["customer_name"]
    list_col = STANDARD_COLUMN_NAMES["list_value"]
    net_col = STANDARD_COLUMN_NAMES["net_sales"]

    unknown_mask = df[customer_col] == UNKNOWN_CUSTOMER
    known = df.loc[~unknown_mask]

    grouped = (
        known.groupby(customer_col, sort=False, as_index=False)
        .agg(**{
            list_col: (list_col, "sum"),
            net_col: (net_col, "sum"),
            RECORDS_AGGREGATED_COL: (net_col, "size"),
        })
    )
    unknown = df.loc[unknown_mask].assign(**{RECORDS_AGGREGATED_COL: 1})

    columns = [customer_col, list_col, net_col, RECORDS_AGGREGATED_COL]
    return pd.concat([grouped[columns], unknown[columns]], ignore_index=True)


def clean_input_data(df: pd.DataFrame, *, aggregate: bool = True) -> pd.DataFrame:
    """Standardize, aggregate, validate, and filter the raw input for analysis.

    Steps:
      1. Map source columns to standard names and normalize customer names.
      2. Coerce list value and net sales to numeric and drop rows where either
         is non-numeric or missing.  (Negative rows such as credits are kept
         at this stage so they net off within their customer group.)
      3. When *aggregate* is True, combine rows sharing the same customer /
         group name.  This happens before thresholds so that small member
         accounts count toward their group.
      4. Calculate Discount Amount and Discount % from the (aggregated) totals.
      5. Keep records with List Value >= ``min_list_value`` (and > 0) and
         ``min_discount_pct`` <= Discount % <= ``max_discount_pct``.
      6. Report what was dropped / combined.
    """
    customer_col = STANDARD_COLUMN_NAMES["customer_name"]
    list_col = STANDARD_COLUMN_NAMES["list_value"]
    net_col = STANDARD_COLUMN_NAMES["net_sales"]

    cleaned = standardize_columns(df)
    input_rows = len(cleaned)
    cleaned[customer_col] = normalize_customer_names(cleaned[customer_col])

    # -- Remove non-numeric rows -------------------------------------------
    cleaned[list_col] = pd.to_numeric(cleaned[list_col], errors="coerce")
    cleaned[net_col] = pd.to_numeric(cleaned[net_col], errors="coerce")
    numeric_mask = cleaned[list_col].notna() & cleaned[net_col].notna()
    non_numeric_count = int((~numeric_mask).sum())
    cleaned = cleaned.loc[numeric_mask]

    # -- Aggregate to one record per customer / group ----------------------
    if aggregate:
        cleaned = aggregate_by_customer(cleaned)
    cleaned = add_discount_columns(cleaned)
    record_count = len(cleaned)

    # -- Apply thresholds --------------------------------------------------
    # NaN discounts (List Value <= 0) fail these comparisons and are excluded.
    list_ok = cleaned[list_col] > 0
    size_ok = cleaned[list_col] >= CONFIG["min_list_value"]
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
        summary_lines.append(f"Customers / groups after aggregation: {record_count}")
    summary_lines += [
        f"Records dropped (List Value <= 0): {no_list_count}",
        f"Records dropped (List Value < {CONFIG['min_list_value']:,}): {low_value_count}",
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

def fit_semilog_trend(df: pd.DataFrame) -> tuple[np.ndarray, float]:
    """Fit Discount % = slope * log10(List Value) + intercept.

    Returns
    -------
    trend_values : ndarray
        Predicted Discount % (fraction), one per row of *df*.
    residual_std : float
        Standard deviation of residuals in discount fraction units (ddof=1);
        0.0 when there are too few points to estimate it.
    """
    if len(df) < 2:
        raise ValueError("At least two rows are required to fit a trend line.")

    log_x = np.log10(df[X_COL].to_numpy(dtype=float))
    y = df[DISCOUNT_COL].to_numpy(dtype=float)

    slope, intercept = np.polyfit(log_x, y, 1)
    trend = slope * log_x + intercept
    residual_std = float(np.std(y - trend, ddof=1)) if len(df) > 2 else 0.0

    return trend, residual_std


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
        method_flags = pd.DataFrame({
            "Robust Z score": z_outlier,
            "Robust IQR": iqr_outlier,
            "Percentile extremes": pct_outlier,
        })
        method_labels = method_flags.apply(
            lambda row: ", ".join(row.index[row.to_numpy(dtype=bool)]),
            axis=1,
        )

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
    """Add net-sales and discount percentile ranks with a quadrant flag.

    Percentile ranks use the PERCENTRANK.INC convention (0-100).

    * ``"Outlier - High Discount, Low List Value"``: net-sales percentile <= low
      threshold AND discount percentile >= high threshold.
    * ``"Outlier - Low Discount, High List Value"``: net-sales percentile >= high
      threshold AND discount percentile <= low threshold.
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

    high_disc_low_sales = (
        (enriched["List Value Percentile"] <= low_thresh)
        & (enriched["Discount Percentile"] >= high_thresh)
    )
    low_disc_high_sales = (
        (enriched["List Value Percentile"] >= high_thresh)
        & (enriched["Discount Percentile"] <= low_thresh)
    )

    enriched["Outlier Flag"] = ""
    enriched.loc[high_disc_low_sales, "Outlier Flag"] = "Outlier - High Discount, Low List Value"
    enriched.loc[low_disc_high_sales, "Outlier Flag"] = "Outlier - Low Discount, High List Value"

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
      * **Outside 2 SD** -- High Concern Outlier

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
    out["Net Sales Impact at Trend"] = (
        (out["Discount Gap vs Trend"] * out[STANDARD_COLUMN_NAMES["list_value"]])
    )
    out["Band"] = "Within 1 SD"
    out["Severity"] = "Normal"
    out["Discount Position"] = (
        "" if CONFIG["band_blank_within_1sd_discount_position"] else "In Line"
    )

    if residual_std != 0:
        abs_resid = out["Discount Gap vs Trend"].abs()
        outside_1sd = abs_resid > residual_std
        beyond_2sd = abs_resid > 2 * residual_std
        between_mask = outside_1sd & ~beyond_2sd

        out.loc[between_mask, "Band"] = "Between 1 and 2 SD"
        out.loc[beyond_2sd, "Band"] = "Outside 2 SD"
        out.loc[between_mask, "Severity"] = "Lower Concern Outlier"
        out.loc[beyond_2sd, "Severity"] = "High Concern Outlier"

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
    """Return ADDITIVE +/-1 SD and +/-2 SD band lines (discount points).

    The y-axis is linear, so bands are trend +/- k*SD rather than the
    multiplicative bands used for price.  Empty dict when SD is zero.
    """
    if residual_std == 0:
        return {}

    return {
        "+1 SD": trend_values + residual_std,
        "-1 SD": trend_values - residual_std,
        "+2 SD": trend_values + 2 * residual_std,
        "-2 SD": trend_values - 2 * residual_std,
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


def _select_best_offset(
    anchor_log_x: float,
    anchor_log_y: float,
    x_log_span: float,
    y_log_span: float,
    placed_labels: list[tuple[float, float]],
    data_points_log: np.ndarray,
    candidate_offsets: np.ndarray,
) -> tuple[float, float]:
    """Choose the candidate offset that best avoids visual clutter.

    Scoring (lower is better) combines three terms:

    * **Label repulsion** -- sum of inverse distances to already-placed
      labels.  Heavily weighted to prevent label-on-label overlap.
    * **Point repulsion** -- sum of inverse distances to all scatter-plot
      data points.  Moderately weighted so labels avoid obscuring data.
    * **Anchor distance** -- slight preference for keeping the label close
      to its data point to maintain readability.

    All position comparisons happen in an abstract space proportional to the
    log-scaled axes, so the scoring respects the visual layout of the chart.

    Parameters
    ----------
    anchor_log_x, anchor_log_y : float
        Plot-space coordinates (log10 list value, discount fraction) of the
        data point being labelled.
    x_log_span, y_log_span : float
        Full plot-space range of each axis (used for coordinate conversion).
    placed_labels : list[tuple[float, float]]
        Plot-space positions of labels already committed to the chart.
    data_points_log : ndarray, shape (N, 2)
        Plot-space coordinates of every scatter-plot point.
    candidate_offsets : ndarray, shape (C, 2)
        Pixel-space ``(dx, dy)`` candidates generated by
        :func:`_generate_candidate_offsets`.
    """
    n_candidates = len(candidate_offsets)

    # Map pixel offsets to approximate log-space positions.
    cx = anchor_log_x + (candidate_offsets[:, 0] / _PLOT_WIDTH_PTS) * x_log_span
    cy = anchor_log_y + (candidate_offsets[:, 1] / _PLOT_HEIGHT_PTS) * y_log_span

    # -- Label repulsion (sum of inverse distances to placed labels) --------
    label_penalty = np.zeros(n_candidates)
    for lx, ly in placed_labels:
        dists = np.hypot(cx - lx, cy - ly)
        label_penalty += 1.0 / np.maximum(dists, 1e-9)

    # -- Point repulsion (sum of inverse distances to data points) ----------
    point_penalty = np.zeros(n_candidates)
    if data_points_log.size > 0:
        # Broadcast: (C,1) - (1,N) -> (C,N) pairwise distance matrix.
        dx = cx[:, np.newaxis] - data_points_log[np.newaxis, :, 0]
        dy = cy[:, np.newaxis] - data_points_log[np.newaxis, :, 1]
        dists = np.hypot(dx, dy)
        point_penalty = np.sum(1.0 / np.maximum(dists, 1e-9), axis=1)

    # -- Anchor distance (prefer positions closer to the data point) --------
    anchor_dist = np.hypot(candidate_offsets[:, 0], candidate_offsets[:, 1])

    # Weighted composite score.
    scores = 8.0 * label_penalty + 1.5 * point_penalty + 0.01 * anchor_dist

    best_idx = int(np.argmin(scores))
    return (float(candidate_offsets[best_idx, 0]),
            float(candidate_offsets[best_idx, 1]))


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
    customer_col = STANDARD_COLUMN_NAMES["customer_name"]

    data_points = to_plot_space(plot_df[X_COL], plot_df[DISCOUNT_COL])
    x_span, y_span = np.maximum(np.ptp(data_points, axis=0), 1e-6)

    candidate_offsets = _generate_candidate_offsets()
    placed_labels: list[tuple[float, float]] = []

    prioritized = labeled_points_df.sort_values(
        "Discount Gap vs Trend",
        ascending=False,
        key=abs,
    )

    for _, row in prioritized.iterrows():
        x_val = float(row[X_COL])
        y_val = float(row[DISCOUNT_COL])
        anchor_px = np.log10(x_val)
        anchor_py = y_val

        offset = _select_best_offset(
            anchor_px,
            anchor_py,
            x_span,
            y_span,
            placed_labels,
            data_points,
            candidate_offsets,
        )

        placed_labels.append((
            anchor_px + (offset[0] / _PLOT_WIDTH_PTS) * x_span,
            anchor_py + (offset[1] / _PLOT_HEIGHT_PTS) * y_span,
        ))

        ax.annotate(
            row[customer_col],
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

def plot_analysis(
    df: pd.DataFrame,
    chart_title: str,
    *,
    show: bool = True,
    chart_path: Path | None = None,
) -> pd.DataFrame:
    """Build the semi-log discount-band chart and return the analysis table."""
    trend_values, residual_std = fit_semilog_trend(df)
    analysis_df = classify_outliers(df, trend_values, residual_std)
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

    fig, ax = plt.subplots(figsize=(12, 8))

    scatter_styles = {
        "Within 1 SD": {"alpha": 0.75, "s": 60, "color": "#1f77b4"},
        "Between 1 and 2 SD": {
            "alpha": 0.85, "s": 75, "color": "#ff7f0e",
            "edgecolors": "black", "linewidths": 0.5,
        },
        "Outside 2 SD": {
            "alpha": 0.95, "s": 90, "color": "#d62728", "edgecolors": "black",
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
        ax.plot(x_sorted, bands["+1 SD"], "--", color="#ff7f0e", lw=1.8, label="+/- 1 SD")
        ax.plot(x_sorted, bands["-1 SD"], "--", color="#ff7f0e", lw=1.8)
        ax.plot(x_sorted, bands["+2 SD"], ":",  color="#2ca02c", lw=2.0, label="+/- 2 SD")
        ax.plot(x_sorted, bands["-2 SD"], ":",  color="#2ca02c", lw=2.0)

    # Net sales on a log axis; discount on a LINEAR axis (zero / negative
    # discounts cannot be shown on a log scale).
    ax.set_xscale("log")
    ax.yaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=0))
    ax.axhline(0, color="grey", lw=0.8)
    # Keep the y-range to the data (bands can otherwise extend past 100 %).
    y_vals = plot_df[DISCOUNT_COL]
    y_pad = 0.05
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

    return analysis_df


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
    * List Value Percentile, Discount Percentile and Outlier Flag directly
      after Discount Position (the "Potentially Over-/Under-Discounted"
      column).
    """
    columns = _move_after(
        columns,
        [RECORDS_AGGREGATED_COL],
        STANDARD_COLUMN_NAMES["customer_name"],
    )
    columns = _move_after(
        columns,
        ["List Value Percentile", "Discount Percentile", "Outlier Flag"],
        "Discount Position",
    )
    return columns


def save_outliers(analysis_df: pd.DataFrame, output_path: Path) -> None:
    """Save the analysis table to a formatted Excel workbook.

    Rows are sorted most-severe first, then by bin, number of within-bin
    methods triggered, list value, and discount.
    """
    severity_rank = {sev: i for i, sev in enumerate(SEVERITY_ORDER)}
    export_df = analysis_df.sort_values(
        [
            "Severity",
            "List Value Bin",
            "Bin Outlier Methods Triggered",
            X_COL,
            DISCOUNT_COL,
        ],
        ascending=[True, True, False, True, False],
        key=lambda col: col.map(severity_rank) if col.name == "Severity" else col,
    )
    export_df = export_df[order_export_columns(list(export_df.columns))]
    export_df.to_excel(
        output_path,
        index=False,
        sheet_name=CONFIG["worksheet_name"],
    )

    workbook = load_workbook(output_path)
    ws = workbook[CONFIG["worksheet_name"]]

    header_fill = PatternFill(
        fill_type="solid",
        start_color=CONFIG["header_fill_color"],
        end_color=CONFIG["header_fill_color"],
    )
    header_font = Font(name="Arial", bold=True, color=CONFIG["header_font_color"])
    header_alignment = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    severity_fills = {
        sev: PatternFill(fill_type="solid", start_color=c, end_color=c)
        for sev, c in CONFIG["severity_fill_colors"].items()
    }
    headers = list(export_df.columns)
    severity_pos = headers.index("Severity")
    number_formats = {**EXCEL_NUMBER_FORMATS, "Net Sales Impact at Trend": "$#,##0"}
    column_formats = [number_formats.get(name) for name in headers]

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_fill = severity_fills.get(row[severity_pos].value)
        for cell, number_format in zip(row, column_formats):
            if number_format:
                cell.number_format = number_format
            if row_fill is not None:
                cell.fill = row_fill

    for col_idx, column_cells in enumerate(ws.iter_cols(), start=1):
        max_len = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in column_cells
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    workbook.save(output_path)


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
        help="Optional PNG filename to save the chart.",
    )
    parser.add_argument(
        "--no-show",
        dest="show",
        action="store_false",
        help="Do not display the chart on screen (useful for batch runs).",
    )
    parser.add_argument(
        "--title",
        default="Discount Band Analysis - Discount from List vs Total List Value",
        help="Chart title.",
    )
    parser.add_argument(
        "--no-aggregate",
        dest="aggregate",
        action="store_false",
        default=CONFIG["aggregate_by_customer"],
        help="Analyse each input row separately instead of aggregating rows "
             "that share the same customer/group name.",
    )
    return parser.parse_args()


def main() -> int:
    """Run the discount band analysis workflow."""
    args = parse_args()

    try:
        input_path = Path(args.input_file) if args.input_file else choose_input_file()

        raw_df = load_input_file(input_path, args.sheet)
        cleaned_df = clean_input_data(raw_df, aggregate=args.aggregate)
        analysis_df = plot_analysis(
            cleaned_df,
            args.title,
            show=args.show,
            chart_path=Path(args.save_chart) if args.save_chart else None,
        )

        output_path = (
            Path(args.output_outliers)
            if args.output_outliers
            else choose_output_file(CONFIG["default_output_filename"])
        )
        save_outliers(analysis_df, output_path)

        high_concern = int((analysis_df["Band"] == "Outside 2 SD").sum())
        bin_outliers = int(analysis_df["Bin Outlier Flag"].sum())
        print("Analysis complete.")
        print(f"Output file saved to: {output_path}")
        print(f"Records analysed: {len(analysis_df)}")
        print(f"High concern outliers (>2 SD from trend): {high_concern}")
        print(f"Records flagged by at least one within-bin method: {bin_outliers}")
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
