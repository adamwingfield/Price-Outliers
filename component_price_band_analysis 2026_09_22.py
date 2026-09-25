"""
Price Band / ASP Benchmarking Analysis  (Dual-Benchmark Version)
=================================================================
Identifies pricing inconsistencies and revenue-uplift opportunities
by comparing each transaction against two complementary benchmarks:

  Market benchmark  – ASP percentiles across ALL customers.
  Peer benchmark    – ASP percentiles among customers sharing one or
                      more of the peer dimensions configured in
                      PEER_DIMENSIONS (by default: Avalanche Segment,
                      GPO, Account Type, Revenue Tier).  A cascade
                      tries the most specific grouping first and falls
                      back to broader pools when a group has fewer than
                      MIN_PEER_GROUP customers.

PEER DIMENSIONS ARE FULLY DYNAMIC
---------------------------------
The single source of truth is the PEER_DIMENSIONS list in the
CONFIGURATION section.  The cascade (_PEER_LEVELS) is generated from it
automatically as every meaningful combination of the listed dimensions,
ordered most-specific first.  Add, remove, or reorder an entry in that
list and the band builder, cascade resolver, both workbooks, the
Information sheets, the Executive Summary and the console output all
follow — no other edits required.

  • Remove "Revenue Tier" from the list and no grouping will use it.
  • Add a new dimension (e.g. Market Segment) and it is immediately
    included in the cascade.
  • Set the list to [] to run a market-benchmark-only analysis.

Each dimension carries a significance flag; dimensions flagged as
high-significance (Avalanche Segment by default) are preferred, so
combinations containing them sit in the upper levels of the cascade.

Outputs two workbooks:
  <input>_PriceBandAnalysis.xlsx  – the main analysis
  <input>_PeerBands.xlsx          – full peer benchmarks per level

Usage:  python <script>.py [path_to_file.xlsx]
        If no path is given a file-picker dialog opens.
"""

import sys, os, subprocess, warnings, datetime
from pathlib import Path

# -- Dependency bootstrap ------------------------------------------------
_PKGS = {"pandas": "pandas", "openpyxl": "openpyxl", "xlsxwriter": "xlsxwriter"}
def _ensure():
    miss = []
    for i, p in _PKGS.items():
        try: __import__(i)
        except ImportError: miss.append(p)
    if not miss: return
    print(f"Installing {', '.join(miss)} ...")
    try: subprocess.check_call([sys.executable, "-m", "pip", "install", "-q"] + miss)
    except subprocess.CalledProcessError:
        sys.exit(f"pip install failed. Run:  pip install {' '.join(miss)}")
_ensure()

import pandas as pd
import numpy as np
from difflib import get_close_matches

# ═══════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════
NUM_TIERS       = 3       # Revenue tiers (Tier 1 = largest)
MIN_PEER_GROUP  = 10      # Min distinct customers in a peer group;
                          # groups below this cascade to the next level.
EXCLUDE_ZERO_QTY_AND_REVENUE = True   # Exclude lines where both quantity
                                      # and revenue total are zero.

# Source column names for enrichment dimensions
GPO_COL       = "GPO / Buying Grp 1"
IDN_COL       = "IDN Parent / Buying Grp 2"
ACCT_TYPE_COL = "Account Type"
AVAL_SEG_COL  = "Avalanche Segment"
TIER_COL      = "Revenue_Tier"   # derived in assign_revenue_tiers()

# ───────────────────────────────────────────────────────────────────
# PEER-GROUP DIMENSIONS  ← THE ONE LIST THAT CONTROLS PEER MATCHING
# ───────────────────────────────────────────────────────────────────
# Every peer grouping used anywhere in the analysis is generated from
# this list.  Remove an entry and the analysis stops using that
# dimension entirely; add an entry and it is picked up automatically.
#
# Keys per entry:
#   column   – column name in the data (after alias resolution).  May be
#              a derived column such as Revenue_Tier.
#   short    – abbreviation used when the dimension appears in a
#              multi-dimension level label (keep it brief: labels double
#              as Excel sheet names, limit 31 characters).
#   solo     – label used when the dimension forms a level on its own.
#              Optional; defaults to `short`.
#   high_significance – True marks the dimension as a preferred matching
#              criterion.  Combinations containing it are placed above
#              equally specific combinations that do not.
#   allow_solo – False prevents the dimension from forming a level by
#              itself (it is considered too broad to define a peer group
#              on its own).  Defaults to True.
#   definition – plain-English description written to the Information
#              sheets of both workbooks.
#
# Set PEER_DIMENSIONS = [] to run a market-benchmark-only analysis.
PEER_DIMENSIONS = [
    {
        "column": GPO_COL,
        "short": "GPO",
        "solo": "GPO",
        "high_significance": False,
        "allow_solo": True,
        "definition": "Group Purchasing Organisation affiliation.",
    },
    {
        "column": ACCT_TYPE_COL,
        "short": "Type",
        "solo": "Account Type",
        "high_significance": False,
        "allow_solo": True,
        "definition": "Facility classification (e.g., HOSPITAL, ASC, CLINIC).",
    },
    {
        "column": AVAL_SEG_COL,
        "short": "Aval",
        "solo": "Avalanche",
        "high_significance": True,
        "allow_solo": True,
        "definition": "Customer segmentation by volume and share: PLATINUM (high volume, "
                      "high share), GOLD (high volume, low share), SILVER (low volume, "
                      "high share), BRONZE (low volume, low share).",
    },
    {
        "column": TIER_COL,
        "short": "Tier",
        "solo": "Revenue Tier",
        "high_significance": False,
        "allow_solo": False,   # too broad to define a peer group alone
        "derived": True,       # calculated by the script, not a source column
        "definition": f"Customers divided into {NUM_TIERS} tiers by total revenue "
                      f"(Tier 1 = largest).",
    },
]

# Optional cap on the number of cascade levels generated.  The number of
# levels grows as 2^n with the number of dimensions (4 dimensions -> up
# to 15 levels), so a cap is useful when experimenting with 6+
# dimensions.  None = no cap.
MAX_CASCADE_LEVELS = None


# The cascade is GENERATED from PEER_DIMENSIONS — it is never hand-
# maintained.  Each generated entry is
# (display label, internal prefix, [grouping columns]).
#
# Generation rules:
#   1. Every combination of the configured dimensions is a candidate
#      level, except single-dimension levels for dimensions flagged
#      allow_solo = False.
#   2. Levels are ordered most specific first:
#        a. more high-significance dimensions first,
#        b. then more dimensions overall,
#        c. then the order the dimensions appear in PEER_DIMENSIONS.
#   3. Labels double as sheet names in the peer-bands workbook, so they
#      are trimmed to Excel's 31-character limit and de-duplicated.
from itertools import combinations


def _dim(d: dict) -> dict:
    """Normalise one PEER_DIMENSIONS entry, applying defaults."""
    if isinstance(d, str):                      # bare column name allowed
        d = {"column": d}
    col = d["column"]
    short = d.get("short") or col
    return {
        "column": col,
        "short": short,
        "solo": d.get("solo") or short,
        "high_significance": bool(d.get("high_significance", False)),
        "allow_solo": bool(d.get("allow_solo", True)),
        "derived": bool(d.get("derived", False)),
        "definition": d.get("definition", ""),
    }


def _validate_peer_dimensions(dims: list[dict]) -> list[dict]:
    """Reject duplicate columns and empty definitions early, with a clear
    message, rather than failing deep inside a groupby."""
    seen = set()
    clean = []
    for raw in dims:
        d = _dim(raw)
        if not d["column"]:
            sys.exit("ERROR – a PEER_DIMENSIONS entry has no 'column' value.")
        if d["column"] in seen:
            sys.exit(f"ERROR – PEER_DIMENSIONS lists '{d['column']}' more than once.")
        seen.add(d["column"])
        clean.append(d)
    return clean


_PEER_DIMS = _validate_peer_dimensions(PEER_DIMENSIONS)

# Columns the peer analysis needs present in the data
PEER_DIM_COLS = [d["column"] for d in _PEER_DIMS]


def _build_peer_levels(dims: list[dict]) -> list[tuple]:
    """Return the cascade [(label, prefix, [columns]), ...] derived from
    the configured dimensions, most specific first."""
    n = len(dims)
    if n == 0:
        return []

    candidates = []
    for size in range(n, 0, -1):
        for combo in combinations(range(n), size):
            if size == 1 and not dims[combo[0]]["allow_solo"]:
                continue
            candidates.append(combo)

    # Sort: high-significance dimensions first, then specificity, then
    # the declared order of the dimensions themselves.
    candidates.sort(key=lambda c: (
        -sum(dims[i]["high_significance"] for i in c),
        -len(c),
        c,
    ))

    if MAX_CASCADE_LEVELS:
        candidates = candidates[:MAX_CASCADE_LEVELS]

    levels, used_labels = [], set()
    for idx, combo in enumerate(candidates, start=1):
        prefix = f"L{idx}"
        if len(combo) == 1:
            label = dims[combo[0]]["solo"]
        else:
            label = " + ".join(dims[i]["short"] for i in combo)
        # Excel forbids  : \ / ? * [ ]  in sheet names, and caps them at
        # 31 characters.  Labels double as sheet names, so clean them.
        for bad in ":\\/?*[]":
            label = label.replace(bad, " ")
        label = " ".join(label.split())[:31].strip()
        if label in used_labels:                 # guarantee uniqueness
            label = f"{label[:31 - len(prefix) - 1]} {prefix}"
        used_labels.add(label)
        levels.append((label, prefix, [dims[i]["column"] for i in combo]))
    return levels


_PEER_LEVELS = _build_peer_levels(_PEER_DIMS)

# Display order for peer-level breakdowns (cascade order + Market)
_PEER_LEVEL_LABELS = [lbl for lbl, _, _ in _PEER_LEVELS] + ["Market"]

# Human-readable list of the active dimensions, used in on-sheet text
_DIM_NAMES = [d["column"].replace("_", " ") for d in _PEER_DIMS]
if not _DIM_NAMES:
    _DIM_TEXT = "no peer dimensions configured – market benchmark only"
elif len(_DIM_NAMES) == 1:
    _DIM_TEXT = _DIM_NAMES[0]
else:
    _DIM_TEXT = ", ".join(_DIM_NAMES[:-1]) + " and " + _DIM_NAMES[-1]

# Cascade endpoints for narrative text (safe when the cascade is empty)
_CASCADE_FIRST = _PEER_LEVELS[0][0] if _PEER_LEVELS else "n/a"
_CASCADE_LAST  = _PEER_LEVELS[-1][0] if _PEER_LEVELS else "n/a"


def check_peer_dim_columns(df: pd.DataFrame) -> None:
    """Confirm every configured peer dimension exists in the data."""
    missing = [c for c in PEER_DIM_COLS if c not in df.columns]
    if missing:
        sys.exit(
            f"ERROR – peer dimension column(s) not found in the data: {missing}\n"
            f"Either add the column to the input file (or to COLUMN_ALIASES) or "
            f"remove the dimension from PEER_DIMENSIONS."
        )



# ═══════════════════════════════════════════════════════════════════
# COLUMN NAME RESOLUTION – handles common variations
# ═══════════════════════════════════════════════════════════════════

# Canonical name -> list of known alternative spellings / abbreviations.
# All comparisons are done after normalisation (lowercase, whitespace-
# collapsed, underscores replaced with spaces), so only genuinely
# different *words* need to be listed here.
COLUMN_ALIASES: dict[str, list[str]] = {
    "Business Partner": [
            "business partner name",
            "bp", "customer", "customer name",
        ],
        "Account Type": [
            "acct type", "accttype", "account category", "customer type",
        ],
        "GPO / Buying Grp 1": [
            "gpo", "buying group", "buying grp",
        ],
        "IDN Parent / Buying Grp 2": [
            "idn parent", "buying group 2", "buying grp 2",
        ],
        "Product Group": [
            "prod group", "prod grp", "product grp",
        ],
        "Market Segment": [
            "mkt segment", "market seg",
        ],
        "Material Group": [
            "mat group", "matl group", "material grp", "mat grp",
        ],
        "Pricing Component Type": [
            "product type", "US Pricing Component Type", "component type", "comp type",
        ],
        "Pricing Component": [
            "L5 Product Brand", "US Pricing Component", "price component",
        ],
        "Avalanche Segment": [
            "avalanche seg", "aval segment", "avalanche",
        ],
        "Selling UOM Unit Qty": [
            "Selling UOM Invoice Qty", "selling uom unit quantity", "selling uom qty",
            "uom unit qty", "uom unit quantity", "uom qty",
            "unit qty", "unit quantity", "units", "quantity",
        ],
        "Discount %": [
            "discount pct", "discount percent", "disc %", "disc pct",
            "discount rate",
        ],
        "ASP": [
            "invoice asp", "inv asp", "asp", "avg selling price",
            "average selling price", "net asp",
        ],
        "Revenue Total": [
            "invoice total", "inv total", "invoice amount", "total invoice",
            "net total", "total revenue", "revenue",
        ],
        "Standard COGS": [
            "std cogs", "cogs", "cost of goods sold", "standard cost",
        ],
        "Std Margin": [
            "standard margin", "standard margin $", "margin", "net std margin",
        ],
        "Margin %": [
            "margin pct", "margin %", "margin percent", "margin rate",
        ],
}


def _normalise(name: str) -> str:
    """Lowercase, strip, replace underscores/hyphens with spaces, collapse
    consecutive spaces and strip trailing/leading whitespace."""
    s = str(name).strip().lower()
    s = s.replace("_", " ").replace("-", " ")
    while "  " in s:
        s = s.replace("  ", " ")
    return s


def _build_alias_lookup() -> dict[str, str]:
    """Return a dict mapping every normalised alias to its canonical name.
    The canonical name's own normalised form is also included so that
    simple case / whitespace differences are resolved automatically."""
    lookup: dict[str, str] = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        lookup[_normalise(canonical)] = canonical
        for alias in aliases:
            norm = _normalise(alias)
            if norm not in lookup:           # first-registered wins
                lookup[norm] = canonical
    return lookup

_ALIAS_LOOKUP = _build_alias_lookup()


def _resolve_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename DataFrame columns that are not exact canonical matches but
    can be identified via normalisation or known aliases.

    Returns the DataFrame with renamed columns and prints each mapping
    applied so the user can verify correctness.
    """
    canonical_set = set(COLUMN_ALIASES.keys())
    rename_map: dict[str, str] = {}
    already_taken: set[str] = set(df.columns) & canonical_set  # already correct

    for col in df.columns:
        if col in canonical_set:
            continue                         # exact match, no rename needed
        norm = _normalise(col)
        candidate = _ALIAS_LOOKUP.get(norm)
        if candidate and candidate not in already_taken:
            rename_map[col] = candidate
            already_taken.add(candidate)

    if rename_map:
        print("  Column-name variations resolved:")
        for old, new in rename_map.items():
            print(f"    '{old}'  →  '{new}'")
        df = df.rename(columns=rename_map)

    return df


def _suggest_matches(missing: list[str], available: list[str]) -> str:
    """For each missing column, suggest close matches from the available
    columns to help the user fix their input file."""
    lines: list[str] = []
    avail_lower = [c.lower() for c in available]
    for col in missing:
        close = get_close_matches(col.lower(), avail_lower, n=3, cutoff=0.5)
        if close:
            suggestions = [available[avail_lower.index(c)] for c in close]
            lines.append(f"  • '{col}'  – did you mean: {suggestions}?")
        else:
            lines.append(f"  • '{col}'  – no close match found")
    return "\n".join(lines)

# ═══════════════════════════════════════════════════════════════════
# EXCLUSION FILTERS – add values to exclude matching rows
# ═══════════════════════════════════════════════════════════════════
EXCLUDE_BUSINESS_PARTNER  = []
EXCLUDE_PRICING_COMPONENT = ["ROSA BRAIN - SERVICE"]
EXCLUDE_COMPONENT_TYPE    = []
EXCLUDE_MATERIAL_GROUP    = []
EXCLUDE_PRODUCT_GROUP     = []
EXCLUDE_GPO               = []
EXCLUDE_ACCOUNT_TYPE      = ["DIST-DISTRIBUTOR"]
EXCLUDE_AVALANCHE_SEG     = []

# ═══════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════
def _pct(s, q):
    """Percentile that handles empty / all-NaN series."""
    v = s.dropna()
    return np.percentile(v, q) if len(v) else np.nan

def _safe_div(num, denom):
    """Element-wise division replacing zero denominators with NaN."""
    return num / denom.replace(0, np.nan)


def _as_bool(s: pd.Series) -> pd.Series:
    """NaN-safe conversion of a flag column to a genuine boolean Series.

    A left merge on a boolean column produces dtype *object* holding
    True / False / NaN.  Calling `.fillna(False).astype(bool)` on that
    triggers the pandas FutureWarning about silent downcasting of object
    arrays (and the behaviour changes in a future pandas release).
    Comparing to True is equivalent — NaN compares False — and always
    returns bool dtype, whatever the incoming dtype.
    """
    return s.eq(True)


# Below-market flag sets used in multiple roll-ups
_BELOW_MED = {"Below Median", "Below P25", "Below P10", "Zero ASP"}
_BELOW_P25 = {"Below P25", "Below P10", "Zero ASP"}
_BELOW_PEER_MED = {"Below Peer Median", "Below Peer P25", "Below Peer P10", "Zero ASP"}
_BELOW_PEER_P25 = {"Below Peer P25", "Below Peer P10", "Zero ASP"}


def _pick_file() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        p = filedialog.askopenfilename(
            title="Select Customer Component Data File",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("All", "*.*")])
        root.destroy()
        if not p: print("No file selected."); sys.exit(0)
        return p
    except ImportError:
        sys.exit("tkinter unavailable – pass file path as argument.")


# ═══════════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════════
REQUIRED_COLS = [
    "Business Partner", ACCT_TYPE_COL, GPO_COL, IDN_COL,
    "Product Group", "Market Segment", "Material Group",
    "Pricing Component Type", "Pricing Component", AVAL_SEG_COL,
    "Selling UOM Unit Qty", "Discount %",
    "ASP", "Revenue Total",
    "Standard COGS", "Std Margin", "Margin %",
]

# Any peer dimension sourced from the input file (i.e. not derived by
# this script) must also be present.  Adding a new dimension to
# PEER_DIMENSIONS therefore extends the load-time validation
# automatically.
for _d in _PEER_DIMS:
    if not _d["derived"] and _d["column"] not in REQUIRED_COLS:
        REQUIRED_COLS.append(_d["column"])


# ───────────────────────────────────────────────────────────────────
# NUMERIC HYGIENE
# ───────────────────────────────────────────────────────────────────
# Excel formula errors ("Invalid types combination used in formula",
# "#VALUE!", "#DIV/0!", "#N/A", ...) arrive as text and turn an entire
# numeric column into an object column, which silently breaks means,
# medians, percentiles and comparisons.  Each column below is coerced to
# numeric; anything that will not convert is replaced with the fill
# value (None = leave as NaN so the row is ignored by the statistics).
NUMERIC_COLS = {
    "Discount %":           0.0,    # error text -> 0%
    "Margin %":             None,
    "ASP":                  None,
    "Revenue Total":        None,
    "Standard COGS":        None,
    "Std Margin":           None,
    "Selling UOM Unit Qty": None,
}

# Show at most this many distinct offending values per column in the log
_MAX_BAD_SAMPLES = 3


def _coerce_numeric_column(df: pd.DataFrame, col: str, fill) -> pd.DataFrame:
    """Force `col` to a numeric dtype.

    Values that cannot be converted (Excel error text, stray labels) are
    set to `fill`, or left as NaN when `fill` is None.  Thousands
    separators, currency symbols and a trailing '%' are tolerated.  Every
    substitution is reported so bad source data is never silently
    absorbed.
    """
    if col not in df.columns:
        return df

    original = df[col]
    if pd.api.types.is_numeric_dtype(original):
        return df                                   # already clean

    as_text = (original.astype(str)
                       .str.strip()
                       .str.replace(",", "", regex=False)
                       .str.replace("$", "", regex=False))
    # "25%" typed as text means 0.25 in a column that otherwise holds
    # decimal fractions, so rescale those entries rather than losing them.
    pct_text = as_text.str.endswith("%")
    numeric = pd.to_numeric(as_text.str.rstrip("%"), errors="coerce")
    numeric = numeric.where(~pct_text, numeric / 100.0)

    bad = numeric.isna() & original.notna() & (original.astype(str).str.strip() != "")
    n_bad = int(bad.sum())
    if n_bad:
        samples = (original[bad].astype(str).str.strip()
                   .value_counts().head(_MAX_BAD_SAMPLES))
        shown = "; ".join(f"'{v}' x{c:,}" for v, c in samples.items())
        more = "" if samples.sum() >= n_bad else ", ..."
        action = "set to 0" if fill == 0 else (
            f"set to {fill}" if fill is not None else "left blank (excluded from stats)")
        print(f"  WARNING – '{col}': {n_bad:,} non-numeric value(s) {action}  [{shown}{more}]")

    if fill is not None:
        numeric = numeric.fillna(fill)

    df = df.copy()
    df[col] = numeric
    return df


def sanitise_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Apply _coerce_numeric_column to every column in NUMERIC_COLS."""
    for col, fill in NUMERIC_COLS.items():
        df = _coerce_numeric_column(df, col, fill)
    return df


def load_data(path: str) -> pd.DataFrame:
    print(f"Loading  {path} ...")
    df = pd.read_excel(path, sheet_name="Data Table")

    # --- Strip stray whitespace from column headers -----------------
    df.columns = [str(c).strip() for c in df.columns]

    # --- Resolve common column-name variations ----------------------
    df = _resolve_columns(df)

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        hints = _suggest_matches(missing, list(df.columns))
        sys.exit(
            f"ERROR – {len(missing)} required column(s) not found after "
            f"alias resolution:\n{hints}\n\n"
            f"Columns present in file:\n  {list(df.columns)}\n\n"
            f"TIP: Add new aliases to the COLUMN_ALIASES dict at the top "
            f"of this script to handle your file's naming convention."
        )
    print(f"  {len(df):,} rows loaded")

    # --- Numeric hygiene: neutralise Excel error text ---------------
    # Must run before any filtering or aggregation, otherwise text
    # values in e.g. Discount % propagate into the statistics.
    df = sanitise_numeric_columns(df)

    filters = {
        "Business Partner":       EXCLUDE_BUSINESS_PARTNER,
        "Pricing Component":      EXCLUDE_PRICING_COMPONENT,
        "Pricing Component Type": EXCLUDE_COMPONENT_TYPE,
        "Material Group":         EXCLUDE_MATERIAL_GROUP,
        "Product Group":          EXCLUDE_PRODUCT_GROUP,
        GPO_COL:                  EXCLUDE_GPO,
        ACCT_TYPE_COL:            EXCLUDE_ACCOUNT_TYPE,
        AVAL_SEG_COL:             EXCLUDE_AVALANCHE_SEG,
    }
    for col, vals in filters.items():
        if vals and col in df.columns:
            pre = len(df)
            df = df[~df[col].isin(vals)]
            n = pre - len(df)
            if n: print(f"  Excluded {n:,} rows  ({col} in {vals})")

    # Remove non-transactional lines (zero quantity and zero revenue)
    if EXCLUDE_ZERO_QTY_AND_REVENUE:
        pre = len(df)
        df = df[~((df["Selling UOM Unit Qty"] == 0) & (df["Revenue Total"] == 0))]
        n = pre - len(df)
        if n: print(f"  Excluded {n:,} rows with zero quantity and zero revenue total")

    print(f"  {len(df):,} rows after filters  |  "
          f"{df['Business Partner'].nunique()} customers  |  "
          f"{df['Pricing Component'].nunique()} components")
    return df


# ═══════════════════════════════════════════════════════════════════
# REVENUE TIER ASSIGNMENT
# ═══════════════════════════════════════════════════════════════════
def assign_revenue_tiers(df: pd.DataFrame) -> pd.DataFrame:
    print(f"Assigning {NUM_TIERS} revenue tiers ...")
    cr = (df.groupby("Business Partner")["Revenue Total"]
          .sum().reset_index(name="_CustRev"))
    labels = [f"Tier {i}" for i in range(1, NUM_TIERS + 1)]
    cr["Revenue_Tier"] = pd.qcut(
        cr["_CustRev"].rank(method="first"), NUM_TIERS, labels=labels[::-1])
    for t in labels:
        m = cr["Revenue_Tier"] == t
        print(f"  {t}: {m.sum():,} customers, "
              f"${cr.loc[m, '_CustRev'].sum():,.0f}")
    return df.merge(cr[["Business Partner", "_CustRev", "Revenue_Tier"]],
                    on="Business Partner", how="left")


# ═══════════════════════════════════════════════════════════════════
# 1. GLOBAL (MARKET) COMPONENT BANDS
# ═══════════════════════════════════════════════════════════════════
def build_market_bands(df: pd.DataFrame) -> pd.DataFrame:
    print("Building market (global) price bands ...")
    g = df.groupby("Pricing Component")
    b = g.agg(
        Component_Type    = ("Pricing Component Type", "first"),
        Material_Group    = ("Material Group", "first"),
        Product_Group     = ("Product Group", "first"),
        Record_Lines      = ("ASP", "size"),
        Customers         = ("Business Partner", "nunique"),
        Total_Revenue     = ("Revenue Total", "sum"),
        Total_Units       = ("Selling UOM Unit Qty", "sum"),
        Mean_ASP          = ("ASP", "mean"),
        P10_ASP           = ("ASP", lambda s: _pct(s, 10)),
        P25_ASP           = ("ASP", lambda s: _pct(s, 25)),
        Median_ASP        = ("ASP", lambda s: _pct(s, 50)),
        P75_ASP           = ("ASP", lambda s: _pct(s, 75)),
        P90_ASP           = ("ASP", lambda s: _pct(s, 90)),
        Min_ASP           = ("ASP", "min"),
        Max_ASP           = ("ASP", "max"),
        Std_ASP           = ("ASP", "std"),
        Mean_Discount     = ("Discount %", "mean"),
        Median_Discount   = ("Discount %", "median"),
        Mean_Margin_Pct   = ("Margin %", "mean"),
        Median_Margin_Pct = ("Margin %", "median"),
    ).reset_index()
    b["CoV"]              = _safe_div(b["Std_ASP"], b["Mean_ASP"])
    b["ASP_Range"]        = b["Max_ASP"] - b["Min_ASP"]
    b["ASP_Range_vs_Med"] = _safe_div(b["ASP_Range"], b["Median_ASP"])
    return b.sort_values("Total_Revenue", ascending=False)


# ═══════════════════════════════════════════════════════════════════
# 2. PEER BANDS  (cascading)
# ═══════════════════════════════════════════════════════════════════
# The cascade levels (_PEER_LEVELS) are generated from PEER_DIMENSIONS
# in the CONFIGURATION section — nothing below needs editing when
# dimensions are added or removed.


def _band_at_level(df, group_cols, prefix):
    """Compute ASP statistics (min, P10, P25, median, mean, P75, P90, max)
    and customer count per peer group."""
    g = df.groupby(["Pricing Component"] + group_cols, observed=True)
    out = g.agg(
        _custs  = ("Business Partner", "nunique"),
        _min    = ("ASP", "min"),
        _p10    = ("ASP", lambda s: _pct(s, 10)),
        _p25    = ("ASP", lambda s: _pct(s, 25)),
        _median = ("ASP", lambda s: _pct(s, 50)),
        _mean   = ("ASP", "mean"),
        _p75    = ("ASP", lambda s: _pct(s, 75)),
        _p90    = ("ASP", lambda s: _pct(s, 90)),
        _max    = ("ASP", "max"),
    ).reset_index()
    out.rename(columns={
        "_custs":  f"{prefix}_Custs",
        "_min":    f"{prefix}_Min",
        "_p10":    f"{prefix}_P10",
        "_p25":    f"{prefix}_P25",
        "_median": f"{prefix}_Median",
        "_mean":   f"{prefix}_Mean",
        "_p75":    f"{prefix}_P75",
        "_p90":    f"{prefix}_P90",
        "_max":    f"{prefix}_Max",
    }, inplace=True)
    out[f"{prefix}_Valid"] = out[f"{prefix}_Custs"] >= MIN_PEER_GROUP
    return out


def build_peer_bands(df):
    """Build peer benchmarks for every cascade level defined in
    _PEER_LEVELS.  Returns a dict keyed by level prefix."""
    if not _PEER_LEVELS:
        print("No peer dimensions configured – peer benchmarks will fall back "
              "to the market benchmark for every line.")
        return {}
    print(f"Building peer bands ({len(_PEER_LEVELS)}-level cascade on "
          f"{len(_PEER_DIMS)} dimension(s): {', '.join(PEER_DIM_COLS)}) ...")
    frames = {}
    for label, prefix, grp_cols in _PEER_LEVELS:
        frames[prefix] = _band_at_level(df, grp_cols, prefix)
        v = frames[prefix][f"{prefix}_Valid"].sum()
        print(f"  {prefix:>3} {label:<26}: {len(frames[prefix]):>6,} groups, {v:>5,} valid")
    return frames


# ═══════════════════════════════════════════════════════════════════
# 3. FLAG RECORDS  (market + peer benchmarks)
# ═══════════════════════════════════════════════════════════════════
def flag_records(df, mkt_bands, peer_frames):
    print("Flagging below-benchmark records ...")

    # --- Market benchmarks ---
    mkt = mkt_bands.set_index("Pricing Component")[
        ["P10_ASP", "P25_ASP", "Median_ASP", "P75_ASP", "Mean_ASP",
         "Customers", "Record_Lines"]]
    out = df.merge(mkt, left_on="Pricing Component",
                   right_index=True, how="left", suffixes=("", "_mkt"))

    out["Price_Band"] = "At/Above Median"
    out.loc[out["ASP"] < out["Median_ASP"], "Price_Band"] = "Below Median"
    out.loc[out["ASP"] < out["P25_ASP"],    "Price_Band"] = "Below P25"
    out.loc[out["ASP"] < out["P10_ASP"],    "Price_Band"] = "Below P10"
    out.loc[out["ASP"] == 0,                "Price_Band"] = "Zero ASP"

    out["ASP_vs_Mkt_Med_Pct"] = _safe_div(out["ASP"] - out["Median_ASP"], out["Median_ASP"])
    out["ASP_vs_Mkt_P25_Pct"] = _safe_div(out["ASP"] - out["P25_ASP"],    out["P25_ASP"])

    out["Mkt_Uplift_Med"] = 0.0
    m = (out["ASP"] < out["Median_ASP"]) & (out["Median_ASP"] > 0)
    out.loc[m, "Mkt_Uplift_Med"] = (out.loc[m, "Median_ASP"] - out.loc[m, "ASP"]) * out.loc[m, "Selling UOM Unit Qty"]

    out["Mkt_Uplift_P25"] = 0.0
    m2 = (out["ASP"] < out["P25_ASP"]) & (out["P25_ASP"] > 0)
    out.loc[m2, "Mkt_Uplift_P25"] = (out.loc[m2, "P25_ASP"] - out.loc[m2, "ASP"]) * out.loc[m2, "Selling UOM Unit Qty"]

    # --- Peer benchmarks (cascade defined by _PEER_LEVELS) ---
    # Merge every level's benchmarks onto the transaction rows.
    for label, prefix, grp_cols in _PEER_LEVELS:
        out = out.merge(peer_frames[prefix],
                        on=["Pricing Component"] + grp_cols, how="left")

    # Resolve the winning level.  Applied in reverse cascade order
    # (least specific first) so that more specific levels overwrite.
    out["Peer_P10"]    = np.nan
    out["Peer_P25"]    = np.nan
    out["Peer_Median"] = np.nan
    out["Peer_Custs"]  = np.nan
    out["Peer_Level"]  = "Market"

    for label, prefix, grp_cols in reversed(_PEER_LEVELS):
        v = _as_bool(out[f"{prefix}_Valid"])
        out.loc[v, "Peer_P10"]    = out.loc[v, f"{prefix}_P10"]
        out.loc[v, "Peer_P25"]    = out.loc[v, f"{prefix}_P25"]
        out.loc[v, "Peer_Median"] = out.loc[v, f"{prefix}_Median"]
        out.loc[v, "Peer_Custs"]  = out.loc[v, f"{prefix}_Custs"]
        out.loc[v, "Peer_Level"]  = label

    # Market fallback for anything still NaN
    fb = out["Peer_Median"].isna()
    out.loc[fb, "Peer_Median"] = out.loc[fb, "Median_ASP"]
    out.loc[fb, "Peer_P25"]    = out.loc[fb, "P25_ASP"]
    out.loc[fb, "Peer_P10"]    = out.loc[fb, "P10_ASP"]

    # Peer price band (mirrors market band categories)
    out["Peer_Band"] = "At/Above Peer Median"
    out.loc[out["ASP"] < out["Peer_Median"], "Peer_Band"] = "Below Peer Median"
    out.loc[out["ASP"] < out["Peer_P25"],    "Peer_Band"] = "Below Peer P25"
    out.loc[out["ASP"] < out["Peer_P10"],    "Peer_Band"] = "Below Peer P10"
    out.loc[out["ASP"] == 0,                 "Peer_Band"] = "Zero ASP"

    out["ASP_vs_Peer_Med_Pct"] = _safe_div(out["ASP"] - out["Peer_Median"], out["Peer_Median"])
    out["ASP_vs_Peer_P25_Pct"] = _safe_div(out["ASP"] - out["Peer_P25"],    out["Peer_P25"])

    out["Peer_Uplift_Med"] = 0.0
    pm = (out["ASP"] < out["Peer_Median"]) & (out["Peer_Median"] > 0)
    out.loc[pm, "Peer_Uplift_Med"] = (out.loc[pm, "Peer_Median"] - out.loc[pm, "ASP"]) * out.loc[pm, "Selling UOM Unit Qty"]

    out["Peer_Uplift_P25"] = 0.0
    pp = (out["ASP"] < out["Peer_P25"]) & (out["Peer_P25"] > 0)
    out.loc[pp, "Peer_Uplift_P25"] = (out.loc[pp, "Peer_P25"] - out.loc[pp, "ASP"]) * out.loc[pp, "Selling UOM Unit Qty"]

    # Drop intermediate per-level columns
    _prefixes = tuple(f"{p}_" for _, p, _ in _PEER_LEVELS)
    drop = [c for c in out.columns if c.startswith(_prefixes)]
    out.drop(columns=drop, inplace=True)

    return out


# ═══════════════════════════════════════════════════════════════════
# 4-5. ROLL-UP FUNCTIONS
# ═══════════════════════════════════════════════════════════════════
def customer_benchmark(fl):
    print("Building customer benchmark ...")
    c = fl.groupby(["Business Partner", "Product Group", AVAL_SEG_COL]).agg(
        Account_Type          = (ACCT_TYPE_COL, "first"),
        GPO                   = (GPO_COL, "first"),
        IDN_Parent            = (IDN_COL, "first"),
        Revenue_Tier          = ("Revenue_Tier", "first"),
        Total_Lines           = ("ASP", "size"),
        Total_Revenue         = ("Revenue Total", "sum"),
        Total_Std_Margin      = ("Std Margin", "sum"),
        Avg_Discount          = ("Discount %", "mean"),
        Lines_Below_Med       = ("Price_Band",  lambda s: s.isin(_BELOW_MED).sum()),
        Lines_Below_P25       = ("Price_Band",  lambda s: s.isin(_BELOW_P25).sum()),
        Lines_Zero            = ("Price_Band",  lambda s: (s == "Zero ASP").sum()),
        Mkt_Uplift_Med        = ("Mkt_Uplift_Med", "sum"),
        Mkt_Uplift_P25        = ("Mkt_Uplift_P25", "sum"),
        Peer_Uplift_Med       = ("Peer_Uplift_Med", "sum"),
        Peer_Uplift_P25       = ("Peer_Uplift_P25", "sum"),
        Lines_Below_Peer      = ("Peer_Band", lambda s: s.isin(_BELOW_PEER_MED).sum()),
        Lines_Below_Peer_P25  = ("Peer_Band", lambda s: s.isin(_BELOW_PEER_P25).sum()),
    ).reset_index()
    c["Wtd_Margin_Pct"]       = _safe_div(c["Total_Std_Margin"], c["Total_Revenue"])
    c["Pct_Lines_Below_Peer_P25"] = c["Lines_Below_Peer_P25"] / c["Total_Lines"]
    c = c[[
        "Business Partner", "Account_Type", "GPO", "IDN_Parent", "Revenue_Tier",
        "Product Group", AVAL_SEG_COL, "Total_Revenue", "Total_Std_Margin",
        "Wtd_Margin_Pct", "Avg_Discount", "Total_Lines",
        "Lines_Below_Med", "Lines_Below_P25", "Lines_Zero",
        "Mkt_Uplift_Med", "Mkt_Uplift_P25",
        "Lines_Below_Peer", "Lines_Below_Peer_P25",
        "Peer_Uplift_Med", "Peer_Uplift_P25",
        "Pct_Lines_Below_Peer_P25",
    ]]
    return c.sort_values("Peer_Uplift_P25", ascending=False)


def comp_type_uplift(fl):
    print("Building component-type uplift ...")
    ct = fl.groupby("Pricing Component Type").agg(
        Total_Lines       = ("ASP", "size"),
        Total_Revenue     = ("Revenue Total", "sum"),
        Total_Std_Margin  = ("Std Margin", "sum"),
        Lines_Below_Med   = ("Price_Band", lambda s: s.isin(_BELOW_MED).sum()),
        Lines_Below_P25   = ("Price_Band", lambda s: s.isin(_BELOW_P25).sum()),
        Mkt_Uplift_Med    = ("Mkt_Uplift_Med", "sum"),
        Mkt_Uplift_P25    = ("Mkt_Uplift_P25", "sum"),
        Peer_Uplift_Med   = ("Peer_Uplift_Med", "sum"),
        Peer_Uplift_P25   = ("Peer_Uplift_P25", "sum"),
    ).reset_index()
    ct["Wtd_Margin_Pct"] = _safe_div(ct["Total_Std_Margin"], ct["Total_Revenue"])
    return ct.sort_values("Peer_Uplift_P25", ascending=False)


# ═══════════════════════════════════════════════════════════════════
# 6. EXECUTIVE SUMMARY KPIs
# ═══════════════════════════════════════════════════════════════════
def build_summary(df, bands, fl, cust, ct):
    n = len(fl)
    tr = df["Revenue Total"].sum()
    peer_levels = fl["Peer_Level"].value_counts()
    return dict(
        total_lines=n, total_rev=tr,
        n_cust=df["Business Partner"].nunique(),
        n_comp=df["Pricing Component"].nunique(),
        below_med=fl["Price_Band"].isin(_BELOW_MED).sum(),
        below_p25=fl["Price_Band"].isin(_BELOW_P25).sum(),
        zero=int((fl["Price_Band"] == "Zero ASP").sum()),
        mkt_up_med=fl["Mkt_Uplift_Med"].sum(),
        mkt_up_p25=fl["Mkt_Uplift_P25"].sum(),
        peer_up_med=fl["Peer_Uplift_Med"].sum(),
        peer_up_p25=fl["Peer_Uplift_P25"].sum(),
        peer_below_med=fl["Peer_Band"].isin(_BELOW_PEER_MED).sum(),
        peer_below_p25=fl["Peer_Band"].isin(_BELOW_PEER_P25).sum(),
        peer_levels=peer_levels,
        top5_cust=cust.groupby("Business Partner")["Peer_Uplift_P25"].sum()
            .nlargest(5).reset_index().values.tolist(),
        top5_ct=ct.head(5)[["Pricing Component Type","Peer_Uplift_P25"]].values.tolist(),
        top5_cov=bands[bands["Record_Lines"] >= 10].nlargest(5, "CoV")[
            ["Pricing Component","CoV","Record_Lines","Total_Revenue"]].values.tolist(),
    )


# ═══════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ═══════════════════════════════════════════════════════════════════
def write_output(path, bands, fl, cust, ct, S):
    print(f"Writing  {path} ...")
    wr = pd.ExcelWriter(path, engine="xlsxwriter")
    wb = wr.book

    # -- Formats --
    H  = wb.add_format({"bold":1,"bg_color":"#1F3864","font_color":"#FFFFFF","border":1,
                         "text_wrap":1,"valign":"vcenter","font_name":"Arial","font_size":10})
    # Market-comparison column headers
    HM = wb.add_format({"bold":1,"bg_color":"#00AEEF","font_color":"#FFFFFF","border":1,
                         "text_wrap":1,"valign":"vcenter","font_name":"Arial","font_size":10})
    # Peer-comparison column headers
    HP = wb.add_format({"bold":1,"bg_color":"#0081C9","font_color":"#FFFFFF","border":1,
                         "text_wrap":1,"valign":"vcenter","font_name":"Arial","font_size":10})
    C  = wb.add_format({"num_format":"$#,##0",   "font_name":"Arial","font_size":10})
    C2 = wb.add_format({"num_format":"$#,##0.00","font_name":"Arial","font_size":10})
    P  = wb.add_format({"num_format":"0.0%",     "font_name":"Arial","font_size":10})
    P2 = wb.add_format({"num_format":"0.00%",    "font_name":"Arial","font_size":10})
    N  = wb.add_format({"num_format":"#,##0",    "font_name":"Arial","font_size":10})
    D  = wb.add_format({"num_format":"0.00",     "font_name":"Arial","font_size":10})
    T  = wb.add_format({"font_name":"Arial","font_size":10})
    TI = wb.add_format({"bold":1,"font_name":"Arial","font_size":14,"font_color":"#1F3864"})
    SB = wb.add_format({"bold":1,"font_name":"Arial","font_size":11,"font_color":"#1F3864","bottom":1})
    LB = wb.add_format({"font_name":"Arial","font_size":10,"indent":1})
    KV = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"num_format":"$#,##0"})
    KN = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"num_format":"#,##0"})
    KP = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"num_format":"0.0%"})
    NT = wb.add_format({"font_name":"Arial","font_size":9,"italic":1,"text_wrap":1,"font_color":"#555"})
    IH = wb.add_format({"bold":1,"font_name":"Arial","font_size":10,"bg_color":"#D6E4F0",
                         "border":1,"text_wrap":1,"valign":"vcenter"})
    IC = wb.add_format({"font_name":"Arial","font_size":10,"text_wrap":1,"valign":"top","border":1})
    IT = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"valign":"top","border":1})

    def _ws(df, name, cols, tab="#4472C4", fz=1, mc=None, pc=None, groups=None):
        """Write a DataFrame sheet.

        mc     = market-comparison columns (cyan header)
        pc     = peer-comparison columns (blue header)
        groups = list of (first_col, last_col) ranges to place in a
                 collapsible outline group so they can be hidden.

        Columns in mc / pc / groups may be given as HEADER NAMES (strings)
        or as 0-based indices.  Names are strongly preferred: they are
        resolved against `cols` itself, so inserting, removing or
        reordering a column can never silently shift a colour or an
        outline group onto its neighbour.  An unknown name is a hard
        error rather than a silent mis-format.

        Returns the {header name -> index} map so callers can position
        conditional formats by name as well.
        """
        ix = {cn: i for i, (cn, _, _) in enumerate(cols)}

        def _res(col, where):
            """Resolve one column reference to an index."""
            if isinstance(col, str):
                if col not in ix:
                    sys.exit(f"ERROR – sheet '{name}': {where} refers to column "
                             f"'{col}', which is not one of its headers:\n"
                             f"  {list(ix)}")
                return ix[col]
            if not 0 <= col < len(cols):
                sys.exit(f"ERROR – sheet '{name}': {where} refers to column index "
                         f"{col}, outside the {len(cols)} columns defined.")
            return col

        df.to_excel(wr, sheet_name=name, startrow=1, index=False, header=False)
        ws = wr.sheets[name]; ws.set_tab_color(tab)
        ws.freeze_panes(1, fz); ws.autofilter(0, 0, len(df), len(df.columns)-1)
        mc = {_res(c, "mc") for c in (mc or ())}
        pc = {_res(c, "pc") for c in (pc or ())}
        overlap = mc & pc
        if overlap:
            names = [cols[i][0] for i in sorted(overlap)]
            sys.exit(f"ERROR – sheet '{name}': column(s) {names} are listed as both "
                     f"market and peer comparison columns.")
        grouped = set()
        for first, last in (groups or []):
            f_i, l_i = _res(first, "groups"), _res(last, "groups")
            if f_i > l_i:
                f_i, l_i = l_i, f_i
            grouped.update(range(f_i, l_i + 1))
        for i, (cn, w, f) in enumerate(cols):
            hdr = HM if i in mc else (HP if i in pc else H)
            ws.write(0, i, cn, hdr)
            if i in grouped:
                ws.set_column(i, i, w, f, {"level": 1})
            else:
                ws.set_column(i, i, w, f)
        return ix

    # ────────────── Information ──────────────────────────────────────
    wi = wb.add_worksheet("Information")
    wi.hide_gridlines(2); wi.set_tab_color("#305496")
    wi.set_column("A:A",4); wi.set_column("B:B",30); wi.set_column("C:C",95)
    r=1; wi.merge_range(r,1,r,2,"Price Band / ASP Benchmarking Analysis",TI); r+=1
    wi.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=2
    wi.merge_range(r,1,r,2,"Workbook Sheets",SB); r+=1
    wi.write(r,1,"Sheet",IH); wi.write(r,2,"Description",IH); r+=1
    for nm,ds in [
        ("Information","Overview of each tab and key term definitions."),
        ("Executive Summary","Headline KPIs under market and peer benchmarks."),
        ("Component Price Bands","Percentile ASP benchmarks per Pricing Component (market-wide)."),
        ("Customer Benchmark","One row per customer x product group x Avalanche Segment with "
         "revenue, margin, below-benchmark line counts, dual uplift, and % of lines below peer P25."),
        ("Flagged Records","Lines below the market P25 or the peer P25, with market and peer "
         "benchmarks side by side.  Includes the Peer Level column showing which cascade level "
         "supplied the benchmark for each line."),
        ("Uplift by Comp Type","Market and peer uplift aggregated by Pricing Component Type."),
        ("Column colours","Market-comparison columns use a cyan header; peer-comparison columns "
         "use a blue header.  Market columns are placed in a collapsible outline group — use the "
         "-/+ toggle above the column headers to hide them."),
    ]:
        wi.write(r,1,nm,IC); wi.write(r,2,ds,IC); wi.set_row(r,48); r+=1
    r+=1; wi.merge_range(r,1,r,2,"Key Terms",SB); r+=1
    wi.write(r,1,"Term",IH); wi.write(r,2,"Definition",IH); r+=1
    for tm,df_ in [
        ("Market Benchmark","ASP percentiles across ALL customers for a Pricing Component."),
        ("Peer Benchmark",f"ASP percentiles among customers sharing one or more of: {_DIM_TEXT}.  "
         "The most specific grouping with enough peers is used for each line."),
        ("Peer Level",f"Which of the {len(_PEER_LEVELS)} cascade levels supplied the benchmark, "
         "or 'Market' where no peer group met the minimum size.  The cascade runs from "
         f"'{_CASCADE_FIRST}' (most specific) down to "
         f"'{_CASCADE_LAST}'.  See the separate Peer Bands workbook for the full "
         "benchmark detail at every level."),
        # One row per configured peer dimension — generated from PEER_DIMENSIONS
        *[(d["column"].replace("_", " "),
           (d["definition"] + "  " if d["definition"] else "")
           + "Used as a peer-grouping dimension"
           + (" (high significance)." if d["high_significance"] else "."))
          for d in _PEER_DIMS],
        ("IDN Parent / Buying Grp 2","Integrated Delivery Network affiliation.  Shows the parent "
         "IDN that the customer belongs to, if any.  Informational only — not used in peer matching."),
        (f"Min Peer Group ({MIN_PEER_GROUP})",f"A peer group needs at least {MIN_PEER_GROUP} distinct "
         "customers to be used; smaller groups cascade to the next level."),
        ("ASP","Average Selling Price per unit (Revenue Total / Units)."),
        ("Market / Peer Price Band","Classification of the line's ASP against the relevant "
         "benchmark distribution: At/Above Median, Below Median, Below P25, Below P10, or Zero ASP."),
        ("P10–P90","Percentile benchmarks of ASP within the relevant population."),
        ("CoV","Coefficient of Variation (Std Dev / Mean ASP).  Higher = more dispersed pricing."),
        ("Uplift to P25","Incremental revenue if the ASP were raised to the P25 benchmark, holding volume constant."),
        ("Uplift to Median","Incremental revenue if the ASP were raised to the Median benchmark, holding volume constant."),
    ]:
        wi.write(r,1,tm,IT); wi.write(r,2,df_,IC); wi.set_row(r,38); r+=1

    # ────────────── Executive Summary ────────────────────────────────
    ws = wb.add_worksheet("Executive Summary")
    ws.hide_gridlines(2); ws.set_tab_color("#1F3864")
    ws.set_column("A:A",4); ws.set_column("B:B",50)
    ws.set_column("C:C",22); ws.set_column("D:D",22); ws.set_column("E:E",22)
    r=1; ws.merge_range(r,1,r,4,"Price Band / ASP Benchmarking Analysis",TI); r+=1
    ws.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=2

    ws.merge_range(r,1,r,4,"Dataset Overview",SB); r+=1
    for l,v,f in [("Total Records",S["total_lines"],KN),("Total Revenue",S["total_rev"],KV),
                   ("Unique Customers",S["n_cust"],KN),("Unique Pricing Components",S["n_comp"],KN),
                   ("Revenue Tiers",NUM_TIERS,KN)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Market Benchmark",SB); r+=1
    for l,v,f in [("Lines Below Market Median",S["below_med"],KN),
                   ("Lines Below Market P25",S["below_p25"],KN),
                   ("Lines at Zero ASP",S["zero"],KN),
                   ("Mkt Uplift to Median",S["mkt_up_med"],KV),
                   ("Mkt Uplift to P25",S["mkt_up_p25"],KV),
                   ("  -- as % of Revenue",S["mkt_up_p25"]/max(S["total_rev"],1),KP)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Peer Benchmark",SB); r+=1
    for l,v,f in [("Lines Below Peer Median",S["peer_below_med"],KN),
                   ("Lines Below Peer P25",S["peer_below_p25"],KN),
                   ("Lines at Zero ASP",S["zero"],KN),
                   ("Peer Uplift to Median",S["peer_up_med"],KV),
                   ("Peer Uplift to P25",S["peer_up_p25"],KV),
                   ("  -- as % of Revenue",S["peer_up_p25"]/max(S["total_rev"],1),KP)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1
    # Peer-level breakdown
    r+=1; ws.merge_range(r,1,r,4,"Peer Level Distribution",SB); r+=1
    ws.write(r,1,"Peer Level",H); ws.write(r,2,"Records",H); r+=1
    for lvl in _PEER_LEVEL_LABELS:
        cnt = S["peer_levels"].get(lvl, 0)
        ws.write(r,1,lvl,T); ws.write(r,2,cnt,N); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Top 5 Customers by Peer Uplift to P25",SB); r+=1
    ws.write(r,1,"Customer",H); ws.write(r,2,"Peer Uplift to P25",H); r+=1
    for nm,v in S["top5_cust"]:
        ws.write(r,1,nm,T); ws.write(r,2,v,C); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Top 5 Component Types by Peer Uplift to P25",SB); r+=1
    ws.write(r,1,"Component Type",H); ws.write(r,2,"Peer Uplift to P25",H); r+=1
    for nm,v in S["top5_ct"]:
        ws.write(r,1,nm,T); ws.write(r,2,v,C); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Top 5 Highest Price Variation (>=10 lines)",SB); r+=1
    ws.write(r,1,"Pricing Component",H); ws.write(r,2,"CoV",H)
    ws.write(r,3,"Lines",H); ws.write(r,4,"Revenue",H); r+=1
    for comp,cov,lines,rev in S["top5_cov"]:
        ws.write(r,1,comp,T); ws.write(r,2,cov,D)
        ws.write(r,3,lines,N); ws.write(r,4,rev,C); r+=1

    r+=2; ws.merge_range(r,1,r,4,"Methodology",SB); r+=1
    ws.merge_range(r,1,r+4,4,
        "Market benchmarks: ASP percentiles across ALL customers per Pricing Component.  "
        f"Peer benchmarks: ASP percentiles among customers sharing one or more of {_DIM_TEXT}, "
        "per Pricing Component.  A "
        f"{len(_PEER_LEVELS)}-level cascade applies the most specific grouping that has at "
        f"least {MIN_PEER_GROUP} distinct customers, running from "
        f"'{_CASCADE_FIRST}' down to '{_CASCADE_LAST}', "
        "falling back to the market benchmark where no peer group qualifies.  "
        f"Revenue tiers ({NUM_TIERS}): Tier 1 = largest customers by total revenue.  "
        "Uplift = incremental revenue if below-benchmark ASPs were raised to the benchmark, "
        "holding unit volume constant.  Full peer benchmark detail for every cascade level is "
        "provided in the accompanying Peer Bands workbook.", NT)

    # ────────────── Component Price Bands ────────────────────────────
    _bix = _ws(bands, "Component Price Bands", [
        ("Pricing Component",32,T),("Component Type",30,T),("Material Group",20,T),
        ("Product Group",18,T),("Lines",8,N),("Customers",10,N),
        ("Total Revenue",16,C),("Total Units",12,N),("Mean ASP",12,C2),
        ("P10 ASP",12,C2),("P25 ASP",12,C2),("Median ASP",12,C2),
        ("P75 ASP",12,C2),("P90 ASP",12,C2),("Min ASP",12,C2),
        ("Max ASP",12,C2),("Std Dev ASP",12,C2),("Avg Discount",12,P2),
        ("Median Discount",14,P2),("Avg Margin %",12,P),("Median Margin %",14,P),
        ("CoV",8,D),("ASP Range ($)",14,C2),("Range / Median",14,D),
    ])
    _cov = _bix["CoV"]
    wr.sheets["Component Price Bands"].conditional_format(1,_cov,len(bands),_cov,{
        "type":"3_color_scale","min_color":"#63BE7B","mid_color":"#FFEB84","max_color":"#F8696B"})

    # ────────────── Customer Benchmark ───────────────────────────────
    _cix = _ws(cust, "Customer Benchmark", [
        ("Business Partner",38,T),("Account Type",16,T),
        ("GPO",18,T),("IDN Parent",22,T),
        ("Revenue Tier",12,T),("Product Group",18,T),
        ("Avalanche Segment",18,T),
        ("Total Revenue",16,C),("Total Std Margin",16,C),
        ("Wtd Margin %",12,P),("Avg Discount",12,P2),("Total Lines",10,N),
        ("Lines Below Mkt Med",16,N),("Lines Below Mkt P25",16,N),
        ("Lines Zero ASP",12,N),
        ("Mkt Uplift to Median",18,C),("Mkt Uplift to P25",16,C),
        ("Lines Below Peer Med",16,N),("Lines Below Peer P25",16,N),
        ("Peer Uplift to Median",18,C),("Peer Uplift to P25",16,C),
        ("% Lines Below Peer P25",18,P),
    ], tab="#548235",
       # Market comparison columns (cyan header)
       mc=["Lines Below Mkt Med","Lines Below Mkt P25","Lines Zero ASP",
           "Mkt Uplift to Median","Mkt Uplift to P25"],
       # Peer comparison columns (blue header)
       pc=["Lines Below Peer Med","Lines Below Peer P25",
           "Peer Uplift to Median","Peer Uplift to P25",
           "% Lines Below Peer P25"],
       # Collapsible market block
       groups=[("Lines Below Mkt Med","Mkt Uplift to P25")])
    _cbar = _cix["Peer Uplift to P25"]
    wr.sheets["Customer Benchmark"].conditional_format(
        1,_cbar,len(cust),_cbar,{"type":"data_bar","bar_color":"#548235"})

    # ────────────── Flagged Records ──────────────────────────────────
    # Include records below EITHER the market P25 or the peer P25 so
    # that peer uplift totals reconcile to the Executive Summary.
    bel = fl[fl["Price_Band"].isin(_BELOW_P25)
             | fl["Peer_Band"].isin(_BELOW_PEER_P25)].copy()
    bel.sort_values("Peer_Uplift_P25", ascending=False, inplace=True)
    fo = bel[[
        "Business Partner",ACCT_TYPE_COL,GPO_COL,IDN_COL,"Revenue_Tier",
        AVAL_SEG_COL,"Pricing Component","Pricing Component Type",
        "Material Group","Product Group",
        "Selling UOM Unit Qty","Revenue Total","ASP",
        "Discount %","Margin %","Price_Band",
        "Median_ASP","ASP_vs_Mkt_Med_Pct","P25_ASP","ASP_vs_Mkt_P25_Pct",
        "Peer_Band",
        "Peer_Median","ASP_vs_Peer_Med_Pct","Peer_P25","ASP_vs_Peer_P25_Pct",
        "Peer_Custs","Peer_Level",
        "Mkt_Uplift_Med","Mkt_Uplift_P25","Peer_Uplift_Med","Peer_Uplift_P25",
    ]].copy()
    fo.columns = [
        "Business Partner","Account Type","GPO","IDN Parent","Revenue Tier",
        "Avalanche Segment","Pricing Component","Component Type",
        "Material Group","Product Group",
        "Units","Revenue Total","ASP",
        "Discount %","Margin %","Market Price Band",
        "Mkt Median ASP","ASP vs Mkt Med %","Mkt P25 ASP","ASP vs Mkt P25 %",
        "Peer Price Band",
        "Peer Median ASP","ASP vs Peer Med %","Peer P25 ASP","ASP vs Peer P25 %",
        "Peer Customers","Peer Level",
        "Mkt Uplift to Median","Mkt Uplift to P25","Peer Uplift to Median","Peer Uplift to P25",
    ]
    fcols = [
        ("Business Partner",38,T),("Account Type",16,T),
        ("GPO",18,T),("IDN Parent",22,T),("Revenue Tier",12,T),
        ("Avalanche Segment",18,T),
        ("Pricing Component",30,T),("Component Type",28,T),
        ("Material Group",20,T),("Product Group",18,T),
        ("Units",8,N),("Revenue Total",14,C),("ASP",12,C2),
        ("Discount %",12,P2),("Margin %",10,P),
        ("Market Price Band",16,T),
        ("Mkt Median ASP",14,C2),("ASP vs Mkt Med %",16,P),
        ("Mkt P25 ASP",12,C2),("ASP vs Mkt P25 %",14,P),
        ("Peer Price Band",16,T),
        ("Peer Median ASP",14,C2),("ASP vs Peer Med %",16,P),
        ("Peer P25 ASP",12,C2),("ASP vs Peer P25 %",14,P),
        ("Peer Customers",14,N),("Peer Level",18,T),
        ("Mkt Uplift to Median",18,C),("Mkt Uplift to P25",16,C),
        ("Peer Uplift to Median",18,C),("Peer Uplift to P25",16,C),
    ]
    _fix = _ws(fo,"Flagged Records",fcols,tab="#C00000",
        # Market comparison columns (cyan header): the market band, its
        # benchmark/gap pairs, and BOTH market uplift columns.
        mc=["Market Price Band",
            "Mkt Median ASP","ASP vs Mkt Med %",
            "Mkt P25 ASP","ASP vs Mkt P25 %",
            "Mkt Uplift to Median","Mkt Uplift to P25"],
        # Peer comparison columns (blue header): the peer band, its
        # benchmark/gap pairs, the peer group size, the cascade level
        # that supplied it, and BOTH peer uplift columns.
        pc=["Peer Price Band",
            "Peer Median ASP","ASP vs Peer Med %",
            "Peer P25 ASP","ASP vs Peer P25 %",
            "Peer Customers","Peer Level",
            "Peer Uplift to Median","Peer Uplift to P25"],
        # Collapsible market blocks — market columns only, so hiding the
        # market view never pulls a peer column with it.
        groups=[("Market Price Band","ASP vs Mkt P25 %"),
                ("Mkt Uplift to Median","Mkt Uplift to P25")])
    wf = wr.sheets["Flagged Records"]
    # Market Price Band conditional formatting
    _mb = _fix["Market Price Band"]
    for v,clr in [("Zero ASP","#C00000"),("Below P10","#ED7D31"),("Below P25","#FFC000")]:
        wf.conditional_format(1,_mb,len(fo),_mb,{
            "type":"text","criteria":"containing","value":v,
            "format":wb.add_format({"bg_color":clr,"font_color":"#FFFFFF",
                                     "font_name":"Arial","font_size":10})})
    # Peer Price Band conditional formatting
    _pb = _fix["Peer Price Band"]
    for v,clr in [("Zero ASP","#C00000"),("Below Peer P10","#ED7D31"),
                   ("Below Peer P25","#FFC000"),("Below Peer Median","#FFEDA0")]:
        wf.conditional_format(1,_pb,len(fo),_pb,{
            "type":"text","criteria":"containing","value":v,
            "format":wb.add_format({"bg_color":clr,"font_color":"#FFFFFF",
                                     "font_name":"Arial","font_size":10})})

    # ────────────── Uplift by Comp Type ──────────────────────────────
    _ws(ct, "Uplift by Comp Type", [
        ("Component Type",30,T),("Total Lines",10,N),("Total Revenue",16,C),
        ("Total Std Margin",16,C),
        ("Lines Below Mkt Med",16,N),("Lines Below Mkt P25",14,N),
        ("Mkt Uplift to Median",18,C),("Mkt Uplift to P25",16,C),
        ("Peer Uplift to Median",18,C),("Peer Uplift to P25",16,C),
        ("Wtd Margin %",12,P),
    ], tab="#7030A0",
       mc=["Lines Below Mkt Med","Lines Below Mkt P25",
           "Mkt Uplift to Median","Mkt Uplift to P25"],
       pc=["Peer Uplift to Median","Peer Uplift to P25"],
       groups=[("Lines Below Mkt Med","Mkt Uplift to P25")])

    wr.close()
    print(f"Done -- {path}")


# ═══════════════════════════════════════════════════════════════════
# PEER BANDS WORKBOOK
# ═══════════════════════════════════════════════════════════════════
def write_peer_bands(path, peer_frames):
    """Write a workbook with one sheet per cascade level showing the
    full peer-group benchmarks, plus an Information sheet.
    Sheets and their grouping columns are driven by _PEER_LEVELS."""
    print(f"Writing  {path} ...")
    wr = pd.ExcelWriter(path, engine="xlsxwriter")
    wb = wr.book

    H  = wb.add_format({"bold":1,"bg_color":"#1F3864","font_color":"#FFFFFF","border":1,
                         "text_wrap":1,"valign":"vcenter","font_name":"Arial","font_size":10})
    C2 = wb.add_format({"num_format":"$#,##0.00","font_name":"Arial","font_size":10})
    N  = wb.add_format({"num_format":"#,##0",    "font_name":"Arial","font_size":10})
    T  = wb.add_format({"font_name":"Arial","font_size":10})
    TI = wb.add_format({"bold":1,"font_name":"Arial","font_size":14,"font_color":"#1F3864"})
    SB = wb.add_format({"bold":1,"font_name":"Arial","font_size":11,"font_color":"#1F3864","bottom":1})
    NT = wb.add_format({"font_name":"Arial","font_size":9,"italic":1,"text_wrap":1,"font_color":"#555"})
    IH = wb.add_format({"bold":1,"font_name":"Arial","font_size":10,"bg_color":"#D6E4F0",
                         "border":1,"text_wrap":1,"valign":"vcenter"})
    IC = wb.add_format({"font_name":"Arial","font_size":10,"text_wrap":1,"valign":"top","border":1})
    IT = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"valign":"top","border":1})
    # Boolean format: show Yes/No
    BT = wb.add_format({"font_name":"Arial","font_size":10,"font_color":"#548235"})
    BF = wb.add_format({"font_name":"Arial","font_size":10,"font_color":"#C00000"})

    def _ws(df, name, cols, tab="#4472C4", fz=1):
        """Write a sheet and return the {header name -> index} map, so
        conditional formats can be positioned by name rather than by a
        hard-coded offset."""
        df.to_excel(wr, sheet_name=name, startrow=1, index=False, header=False)
        ws = wr.sheets[name]; ws.set_tab_color(tab)
        ws.freeze_panes(1, fz); ws.autofilter(0, 0, len(df), len(df.columns)-1)
        for i, (cn, w, f) in enumerate(cols):
            ws.write(0, i, cn, H); ws.set_column(i, i, w, f)
        return {cn: i for i, (cn, _, _) in enumerate(cols)}

    # ── Information sheet ──
    wi = wb.add_worksheet("Information")
    wi.hide_gridlines(2); wi.set_tab_color("#305496")
    wi.set_column("A:A",4); wi.set_column("B:B",28); wi.set_column("C:C",95)
    r=1; wi.merge_range(r,1,r,2,"Peer Group Price Benchmarks",TI); r+=1
    wi.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=2

    wi.merge_range(r,1,r,2,"Peer Dimensions In Use",SB); r+=1
    wi.write(r,1,"Dimension",IH); wi.write(r,2,"Role",IH); r+=1
    if _PEER_DIMS:
        for d in _PEER_DIMS:
            wi.write(r,1,d["column"].replace("_"," "),IT)
            wi.write(r,2,(d["definition"] + "  " if d["definition"] else "")
                        + ("High-significance dimension — combinations containing it are "
                           "tried first." if d["high_significance"]
                           else "Standard peer-matching dimension.")
                        + ("" if d["allow_solo"] else "  Not used as a peer group on its own."),IC)
            wi.set_row(r,38); r+=1
    else:
        wi.write(r,1,"None",IT)
        wi.write(r,2,"No peer dimensions are configured; every line uses the market "
                     "benchmark.",IC); r+=1
    r+=1

    wi.merge_range(r,1,r,2,"Workbook Contents",SB); r+=1
    wi.write(r,1,"Sheet",IH); wi.write(r,2,"Description",IH); r+=1
    for label, prefix, grp_cols in _PEER_LEVELS:
        dims = ", ".join(grp_cols)
        wi.write(r,1,label,IC)
        wi.write(r,2,f"Peer benchmarks grouped by [{dims}] per Pricing Component.  "
                      f"Each row is a unique combination of Pricing Component and the "
                      f"grouping dimensions.  'Valid' indicates whether the group has "
                      f"at least {MIN_PEER_GROUP} distinct customers.",IC)
        wi.set_row(r,48); r+=1

    r+=1; wi.merge_range(r,1,r,2,"Key Terms",SB); r+=1
    wi.write(r,1,"Term",IH); wi.write(r,2,"Definition",IH); r+=1
    for tm,df_ in [
        ("Peer Customers","Number of distinct customers in this peer group for this Pricing Component."),
        ("Peer Min ASP","Minimum ASP within the peer group."),
        ("Peer P10 ASP","10th percentile of ASP within the peer group."),
        ("Peer P25 ASP","25th percentile of ASP within the peer group."),
        ("Peer Median ASP","50th percentile (median) of ASP within the peer group."),
        ("Peer Mean ASP","Arithmetic mean of ASP within the peer group."),
        ("Peer P75 ASP","75th percentile of ASP within the peer group."),
        ("Peer P90 ASP","90th percentile of ASP within the peer group."),
        ("Peer Max ASP","Maximum ASP within the peer group."),
        ("Valid",f"Whether the peer group meets the minimum size threshold "
                 f"({MIN_PEER_GROUP} customers).  Groups below this threshold are "
                 f"not used for benchmarking; the cascade falls through to the next level."),
    ]:
        wi.write(r,1,tm,IT); wi.write(r,2,df_,IC); wi.set_row(r,38); r+=1

    # ── One sheet per cascade level ──
    for label, prefix, grp_cols in _PEER_LEVELS:
        df_out = peer_frames[prefix].copy()

        # Rename prefixed columns to friendly names
        df_out.rename(columns={
            f"{prefix}_Custs":  "Peer Customers",
            f"{prefix}_Min":    "Peer Min ASP",
            f"{prefix}_P10":    "Peer P10 ASP",
            f"{prefix}_P25":    "Peer P25 ASP",
            f"{prefix}_Median": "Peer Median ASP",
            f"{prefix}_Mean":   "Peer Mean ASP",
            f"{prefix}_P75":    "Peer P75 ASP",
            f"{prefix}_P90":    "Peer P90 ASP",
            f"{prefix}_Max":    "Peer Max ASP",
            f"{prefix}_Valid":  "Valid",
        }, inplace=True)

        # Convert boolean Valid to Yes/No string for clean Excel display
        df_out["Valid"] = df_out["Valid"].map({True: "Yes", False: "No"})

        # Build column definitions: grouping cols + metric cols
        cols = [("Pricing Component", 32, T)]
        for gc in grp_cols:
            w = 22 if len(gc) > 15 else 14
            cols.append((gc, w, T))
        cols += [
            ("Peer Customers",  14, N),
            ("Peer Min ASP",    12, C2),
            ("Peer P10 ASP",    12, C2),
            ("Peer P25 ASP",    12, C2),
            ("Peer Median ASP", 14, C2),
            ("Peer Mean ASP",   12, C2),
            ("Peer P75 ASP",    12, C2),
            ("Peer P90 ASP",    12, C2),
            ("Peer Max ASP",    12, C2),
            ("Valid",           8,  T),
        ]

        # Select columns in the right order
        col_order = ["Pricing Component"] + grp_cols + [
            "Peer Customers", "Peer Min ASP", "Peer P10 ASP", "Peer P25 ASP",
            "Peer Median ASP", "Peer Mean ASP", "Peer P75 ASP", "Peer P90 ASP",
            "Peer Max ASP", "Valid"]
        df_out = df_out[col_order]

        # Truncate sheet name to 31 chars (Excel limit)
        sheet_name = label[:31]
        _lix = _ws(df_out, sheet_name, cols, tab="#4472C4", fz=1)

        # Colour-code the Valid column
        ws = wr.sheets[sheet_name]
        valid_col = _lix["Valid"]
        ws.conditional_format(1, valid_col, len(df_out), valid_col, {
            "type": "text", "criteria": "containing", "value": "Yes",
            "format": BT})
        ws.conditional_format(1, valid_col, len(df_out), valid_col, {
            "type": "text", "criteria": "containing", "value": "No",
            "format": BF})

        print(f"  {label}: {len(df_out):,} rows")

    wr.close()
    print(f"  Peer bands workbook saved.")


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════
def main():
    warnings.filterwarnings("ignore", category=UserWarning)
    path = sys.argv[1] if len(sys.argv) > 1 else _pick_file()
    if not os.path.isfile(path): sys.exit(f"Not found: {path}")

    stem    = Path(path).stem
    out_dir = Path(path).parent
    out     = str(out_dir / f"{stem}_PriceBandAnalysis.xlsx")
    out_pb  = str(out_dir / f"{stem}_PeerBands.xlsx")

    df = load_data(path)
    df = assign_revenue_tiers(df)
    check_peer_dim_columns(df)          # configured dimensions must exist
    bands = build_market_bands(df)
    peer_frames = build_peer_bands(df)
    fl = flag_records(df, bands, peer_frames)
    cust = customer_benchmark(fl)
    ct = comp_type_uplift(fl)
    S = build_summary(df, bands, fl, cust, ct)

    write_output(out, bands, fl, cust, ct, S)
    write_peer_bands(out_pb, peer_frames)

    print(f"\n{'='*65}")
    print("  HEADLINE RESULTS")
    print(f"{'='*65}")
    print(f"  Total Revenue:              ${S['total_rev']:>14,.0f}")
    print(f"  Market Uplift to P25:       ${S['mkt_up_p25']:>14,.0f}  "
          f"({S['mkt_up_p25']/S['total_rev']:.1%})")
    print(f"  Peer Uplift to P25:         ${S['peer_up_p25']:>14,.0f}  "
          f"({S['peer_up_p25']/S['total_rev']:.1%})")
    print(f"  Lines Below Peer Median:    {S['peer_below_med']:>14,}  "
          f"({S['peer_below_med']/S['total_lines']:.1%})")
    print(f"\n  Peer Level Distribution:")
    for lvl in _PEER_LEVEL_LABELS:
        cnt = S["peer_levels"].get(lvl, 0)
        if cnt: print(f"    {lvl:<26s}  {cnt:>7,}  ({cnt/S['total_lines']:.1%})")
    print(f"{'='*65}")


if __name__ == "__main__":
    main()
