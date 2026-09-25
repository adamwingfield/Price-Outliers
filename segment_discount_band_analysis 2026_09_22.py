"""
Discount Band Benchmarking Analysis  (Dual-Benchmark Version)
=============================================================
Identifies pricing inconsistencies and revenue-uplift opportunities by
comparing each customer line against typical DISCOUNT OFF LIST PRICE
rather than against absolute price points.

BENCHMARK UNIT
--------------
Every benchmark is computed within a *segment*, defined as the
combination configured in BENCHMARK_KEY:

        Product Group  +  Market Segment  +  Material Group

Two complementary benchmarks are produced for each segment:

  Market benchmark  – discount percentiles across ALL customers trading
                      in that segment.
  Peer benchmark    – discount percentiles among customers in that
                      segment who share one or more of the peer
                      dimensions configured in PEER_DIMENSIONS (by
                      default: Avalanche Segment, GPO, Account Type,
                      Revenue Tier).  A cascade tries the most specific
                      grouping first and falls back to broader pools
                      when a group has fewer than MIN_PEER_GROUP
                      customers.

DIRECTION OF THE METRIC
-----------------------
Discount is expressed as a positive fraction off list:

        Discount % = 1 - (Invoice Total / List Price Total)

A HIGHER discount is a WORSE commercial outcome, so the bands run the
opposite way to an ASP analysis: lines are flagged when their discount
sits ABOVE the segment median / P75 / P90.  Uplift is the incremental
revenue obtained by pulling an over-discounted line back to the
benchmark discount, holding list price and volume constant:

        Uplift = List Price Total x (Actual Discount - Benchmark Discount)

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
  • Add a new dimension and it is immediately included in the cascade.
  • Set the list to [] to run a market-benchmark-only analysis.

BENCHMARK_KEY is dynamic in the same way: drop "Material Group" from it
and the whole analysis re-bases onto Product Group + Market Segment.

Outputs two workbooks:
  <input>_DiscountBandAnalysis.xlsx  – the main analysis
  <input>_PeerDiscountBands.xlsx     – full peer benchmarks per level

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
SHEET_NAME      = "Data Table"   # worksheet holding the transaction data
NUM_TIERS       = 3       # Revenue tiers (Tier 1 = largest)
MIN_PEER_GROUP  = 5      # Min distinct customers in a peer group;
                          # groups below this cascade to the next level.
EXCLUDE_ZERO_QTY_AND_REVENUE = True   # Exclude lines where both quantity
                                      # and invoice total are zero.

# ───────────────────────────────────────────────────────────────────
# BENCHMARK KEY  ← the level at which discounts are compared
# ───────────────────────────────────────────────────────────────────
# Every market and peer benchmark is calculated within this combination.
# Edit the list to re-base the whole analysis (e.g. drop "Material
# Group" to benchmark at Product Group + Market Segment only).
BENCHMARK_KEY = ["Product Group", "Market Segment", "Material Group"]

# Short label used in headings / console output
BENCHMARK_LABEL = " + ".join(BENCHMARK_KEY)

# Minimum number of distinct customers a SEGMENT needs before its market
# benchmark is considered reliable.  Segments below this are still
# reported but their lines are marked "Thin Segment" and contribute no
# uplift, preventing spurious opportunities from 2-customer segments.
MIN_SEGMENT_CUSTOMERS = 5

# ───────────────────────────────────────────────────────────────────
# DISCOUNT METRIC
# ───────────────────────────────────────────────────────────────────
# The analysed metric is discount off list, expressed as a POSITIVE
# fraction (0.40 = 40% off list).  It is recomputed from list price and
# invoice value wherever possible because source discount columns are
# frequently sign-inverted or corrupted on credit / zero-list lines.
RECOMPUTE_DISCOUNT = True        # True  -> 1 - Invoice Total / List Price Total
                                 # False -> use the source Discount % column
DISCOUNT_SIGN = "auto"           # Sign handling for the SOURCE column:
                                 #   "auto"     – flip if predominantly negative
                                 #   "positive" – source already positive
                                 #   "negative" – source is negative, always flip

# The source Discount % column sometimes carries Excel formula error
# text ("Invalid types combination used in formula", "#VALUE!", "#N/A",
# ...) instead of a number.  True treats those rows as a 0% discount;
# False leaves them blank, so the row falls back to the recomputed
# discount and is excluded from the benchmarks if that is unavailable.
# The substitution is applied AFTER the points/fraction rescale and the
# sign detection below, so injected zeros cannot distort either test.
DISCOUNT_TEXT_ERROR_TO_ZERO = True

# Rows whose discount falls outside this range are treated as data
# artefacts (credits, zero/negative list price, rebilled lines).  They
# are retained in the output but excluded from every percentile and
# from all uplift totals.
DISCOUNT_VALID_MIN = -0.50       # up to 50% ABOVE list (premium pricing)
DISCOUNT_VALID_MAX =  1.00       # 100% discount = zero invoice value

# Percentile pair used for banding and uplift.  Because a higher
# discount is worse, the CONSERVATIVE target is the upper percentile
# (pull the worst discounts back to P75) and the AGGRESSIVE target is
# the median.
UPPER_PCTL = 75                  # conservative uplift target
LOWER_PCTL = 25                  # reported for context only

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
        "definition": "Facility classification (e.g., HOSPITAL, ASC, CLINIC, DISTRIBUTOR).",
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


def _validate_peer_dimensions(dims: list) -> list:
    """Reject duplicate columns and columns that are already part of the
    benchmark key, with a clear message, rather than failing deep inside
    a groupby."""
    seen = set()
    clean = []
    for raw in dims:
        d = _dim(raw)
        if not d["column"]:
            sys.exit("ERROR – a PEER_DIMENSIONS entry has no 'column' value.")
        if d["column"] in seen:
            sys.exit(f"ERROR – PEER_DIMENSIONS lists '{d['column']}' more than once.")
        if d["column"] in BENCHMARK_KEY:
            sys.exit(f"ERROR – '{d['column']}' is part of BENCHMARK_KEY and cannot "
                     f"also be a peer dimension (it is already held constant).")
        seen.add(d["column"])
        clean.append(d)
    return clean


_PEER_DIMS = _validate_peer_dimensions(PEER_DIMENSIONS)

# Columns the peer analysis needs present in the data
PEER_DIM_COLS = [d["column"] for d in _PEER_DIMS]


def _build_peer_levels(dims: list) -> list:
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


def check_required_dim_columns(df: pd.DataFrame) -> None:
    """Confirm every configured peer dimension and benchmark-key column
    exists in the data."""
    missing = [c for c in BENCHMARK_KEY if c not in df.columns]
    if missing:
        sys.exit(
            f"ERROR – benchmark key column(s) not found in the data: {missing}\n"
            f"Either add the column to the input file (or to COLUMN_ALIASES) or "
            f"remove it from BENCHMARK_KEY."
        )
    # Derived dimensions (e.g. Revenue_Tier) are created later by the
    # script, so they are not expected in the source file.
    missing = [d["column"] for d in _PEER_DIMS
               if not d["derived"] and d["column"] not in df.columns]
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
COLUMN_ALIASES = {
    "Business Partner": [
        "business partner name",
        "bp", "customer", "customer name", "sold to", "sold to party",
    ],
    "Account Type": [
        "acct type", "accttype", "account category", "customer type",
    ],
    "GPO / Buying Grp 1": [
        "gpo", "buying group", "buying grp", "gpo buying grp 1",
    ],
    "IDN Parent / Buying Grp 2": [
        "idn parent", "buying group 2", "buying grp 2", "idn parent buying grp 2",
    ],
    "Product Group": [
        "prod group", "prod grp", "product grp", "franchise",
    ],
    "Market Segment": [
        "mkt segment", "market seg", "segment",
    ],
    "Material Group": [
        "mat group", "matl group", "material grp", "mat grp",
    ],
    "Avalanche Segment": [
        "avalanche seg", "aval segment", "avalanche",
    ],
    "Selling UOM Unit Qty": [
        "selling uom invoice qty", "selling uom unit quantity", "selling uom qty",
        "uom unit qty", "uom unit quantity", "uom qty",
        "unit qty", "unit quantity", "units", "quantity", "invoice qty",
    ],
    "List Price Total": [
        "list price", "list total", "total list price", "gross price",
        "gross sales", "list value", "list amount", "list price value",
        "Country List Price",
    ],
    "Discount %": [
        "discount pct", "discount percent", "disc %", "disc pct",
        "discount rate", "discount off list",
    ],
    "Invoice ASP": [
        "inv asp", "asp", "avg selling price", "average selling price", "net asp",
    ],
    "Invoice Total": [
        "inv total", "invoice amount", "total invoice", "net sales", "net revenue", "net total"
    ],
    "Standard COGS": [
        "std cogs", "cogs", "cost of goods sold", "standard cost",
    ],
    "Avg COGS": [
        "average cogs", "avg cost", "unit cogs",
    ],
    "Std Margin": [
        "standard margin", "standard margin $", "margin",
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


def _build_alias_lookup() -> dict:
    """Return a dict mapping every normalised alias to its canonical name.
    The canonical name's own normalised form is also included so that
    simple case / whitespace differences are resolved automatically."""
    lookup = {}
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
    rename_map = {}
    already_taken = set(df.columns) & canonical_set  # already correct

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
            print(f"    '{old}'  ->  '{new}'")
        df = df.rename(columns=rename_map)

    return df


def _suggest_matches(missing: list, available: list) -> str:
    """For each missing column, suggest close matches from the available
    columns to help the user fix their input file."""
    lines = []
    avail_lower = [c.lower() for c in available]
    for col in missing:
        close = get_close_matches(col.lower(), avail_lower, n=3, cutoff=0.5)
        if close:
            suggestions = [available[avail_lower.index(c)] for c in close]
            lines.append(f"  - '{col}'  – did you mean: {suggestions}?")
        else:
            lines.append(f"  - '{col}'  – no close match found")
    return "\n".join(lines)

# ═══════════════════════════════════════════════════════════════════
# EXCLUSION FILTERS – add values to exclude matching rows
# ═══════════════════════════════════════════════════════════════════
EXCLUDE_BUSINESS_PARTNER  = []
EXCLUDE_MATERIAL_GROUP    = ["SERVICE - SERVICE"]
EXCLUDE_PRODUCT_GROUP     = []
EXCLUDE_MARKET_SEGMENT    = []
EXCLUDE_GPO               = []
EXCLUDE_ACCOUNT_TYPE      = ["DIST-DISTRIBUTOR"]
EXCLUDE_AVALANCHE_SEG     = []

# ═══════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════
def _pct(s, q):
    """Percentile that handles empty / all-NaN series."""
    v = pd.to_numeric(s, errors="coerce").dropna()
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


# Band labels.  A HIGHER discount is worse, so the flagged bands sit
# ABOVE the benchmark percentiles.
BAND_OK          = "At/Below Median"
BAND_ABOVE_MED   = "Above Median"
BAND_ABOVE_UPPER = f"Above P{UPPER_PCTL}"
BAND_ABOVE_P90   = "Above P90"
BAND_ZERO_REV    = "Zero Revenue"
BAND_NOT_BM      = "Not Benchmarked"

PEER_BAND_OK          = "At/Below Peer Median"
PEER_BAND_ABOVE_MED   = "Above Peer Median"
PEER_BAND_ABOVE_UPPER = f"Above Peer P{UPPER_PCTL}"
PEER_BAND_ABOVE_P90   = "Above Peer P90"

# Flag sets used in multiple roll-ups
_ABOVE_MED   = {BAND_ABOVE_MED, BAND_ABOVE_UPPER, BAND_ABOVE_P90, BAND_ZERO_REV}
_ABOVE_UPPER = {BAND_ABOVE_UPPER, BAND_ABOVE_P90, BAND_ZERO_REV}
_ABOVE_PEER_MED   = {PEER_BAND_ABOVE_MED, PEER_BAND_ABOVE_UPPER,
                     PEER_BAND_ABOVE_P90, BAND_ZERO_REV}
_ABOVE_PEER_UPPER = {PEER_BAND_ABOVE_UPPER, PEER_BAND_ABOVE_P90, BAND_ZERO_REV}


def _pick_file() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        p = filedialog.askopenfilename(
            title="Select Customer Product Segment Data File",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("All", "*.*")])
        root.destroy()
        if not p: print("No file selected."); sys.exit(0)
        return p
    except ImportError:
        sys.exit("tkinter unavailable – pass file path as argument.")


# ═══════════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════════
# Columns that must be present for the analysis to run at all.
REQUIRED_COLS = [
    "Business Partner",
    "Selling UOM Unit Qty", "List Price Total", "Invoice Total",
]

# Columns that enrich the output but are not essential.  Missing ones
# are created empty (numeric -> NaN, text -> "--") so the rest of the
# analysis, and both workbooks, still build.
OPTIONAL_NUMERIC_COLS = [
    "Invoice ASP", "Discount %", "Standard COGS", "Avg COGS",
    "Std Margin", "Margin %",
]
OPTIONAL_TEXT_COLS = [IDN_COL]

# The benchmark key and any peer dimension sourced from the input file
# (i.e. not derived by this script) must also be present.  Adding a new
# dimension to PEER_DIMENSIONS or BENCHMARK_KEY therefore extends the
# load-time validation automatically.
for _c in BENCHMARK_KEY:
    if _c not in REQUIRED_COLS:
        REQUIRED_COLS.append(_c)
for _d in _PEER_DIMS:
    if not _d["derived"] and _d["column"] not in REQUIRED_COLS:
        REQUIRED_COLS.append(_d["column"])


# ───────────────────────────────────────────────────────────────────
# NUMERIC HYGIENE
# ───────────────────────────────────────────────────────────────────
# Excel formula errors ("Invalid types combination used in formula",
# "#VALUE!", "#DIV/0!", "#N/A", ...) arrive as text and turn an entire
# numeric column into an object column.  pd.to_numeric(errors="coerce")
# already neutralises them, but silently — so the columns below are
# routed through a helper that converts, tolerates thousands separators
# / currency symbols / a trailing '%', reports what it changed, and
# optionally substitutes a fill value.
#
# Fill value None = leave as NaN, so the row is simply excluded from the
# statistics rather than being distorted by a fabricated number.
# Discount % is deliberately left as NaN HERE; the 0% substitution for
# text errors is applied in prepare_discount(), after the rescale and
# sign tests, controlled by DISCOUNT_TEXT_ERROR_TO_ZERO.
NUMERIC_COLS = {
    "Selling UOM Unit Qty": None,
    "List Price Total":     None,
    "Invoice Total":        None,
    "Invoice ASP":          None,
    "Discount %":           None,
    "Standard COGS":        None,
    "Avg COGS":             None,
    "Std Margin":           None,
    "Margin %":             None,
}

# Helper column recording which rows held non-numeric text in Discount %
DISCOUNT_TEXT_FLAG = "_Discount_Text_Error"

# Show at most this many distinct offending values per column in the log
_MAX_BAD_SAMPLES = 3


def _coerce_numeric_column(df: pd.DataFrame, col: str, fill,
                           flag_col: str = None) -> pd.DataFrame:
    """Force `col` to a numeric dtype.

    Values that cannot be converted (Excel error text, stray labels) are
    set to `fill`, or left as NaN when `fill` is None.  Thousands
    separators, currency symbols and a trailing '%' are tolerated.  When
    `flag_col` is given, a boolean column of that name marks the rows
    that held unconvertible text, so a later step can decide what to do
    with them.  Every substitution is reported so bad source data is
    never silently absorbed.
    """
    if col not in df.columns:
        return df

    df = df.copy()
    original = df[col]

    if pd.api.types.is_numeric_dtype(original):
        if flag_col:
            df[flag_col] = False
        return df

    as_text = (original.astype(str)
                       .str.strip()
                       .str.replace(",", "", regex=False)
                       .str.replace("$", "", regex=False))
    # "25%" typed as text means 0.25 in a column that otherwise holds
    # decimal fractions, so rescale those entries rather than losing them.
    pct_text = as_text.str.endswith("%")
    numeric = pd.to_numeric(as_text.str.rstrip("%"), errors="coerce")
    numeric = numeric.where(~pct_text, numeric / 100.0)

    bad = numeric.isna() & original.notna() & (as_text != "")
    n_bad = int(bad.sum())
    if n_bad:
        samples = original[bad].astype(str).str.strip().value_counts().head(_MAX_BAD_SAMPLES)
        shown = "; ".join(f"'{v}' x{c:,}" for v, c in samples.items())
        more = "" if samples.sum() >= n_bad else ", ..."
        if flag_col:
            action = "flagged – see DISCOUNT_TEXT_ERROR_TO_ZERO"
        elif fill is not None:
            action = f"set to {fill}"
        else:
            action = "left blank (excluded from stats)"
        print(f"  WARNING – '{col}': {n_bad:,} non-numeric value(s) {action}  [{shown}{more}]")

    if fill is not None:
        numeric = numeric.fillna(fill)

    df[col] = numeric
    if flag_col:
        df[flag_col] = _as_bool(bad)
    return df


def sanitise_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Apply _coerce_numeric_column to every column in NUMERIC_COLS."""
    for col, fill in NUMERIC_COLS.items():
        flag = DISCOUNT_TEXT_FLAG if col == "Discount %" else None
        df = _coerce_numeric_column(df, col, fill, flag_col=flag)
    if DISCOUNT_TEXT_FLAG not in df.columns:
        df[DISCOUNT_TEXT_FLAG] = False
    return df


def _read_any(path: str) -> pd.DataFrame:
    """Read the transaction data from the configured worksheet, falling
    back to the first sheet, and supporting CSV input."""
    if str(path).lower().endswith((".csv", ".txt")):
        return pd.read_csv(path)
    xl = pd.ExcelFile(path)
    sheet = SHEET_NAME if SHEET_NAME in xl.sheet_names else xl.sheet_names[0]
    if sheet != SHEET_NAME:
        print(f"  Worksheet '{SHEET_NAME}' not found – using '{sheet}'.")
    return xl.parse(sheet)


def load_data(path: str) -> pd.DataFrame:
    print(f"Loading  {path} ...")
    df = _read_any(path)

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

    # --- Fill in optional columns so downstream code is unconditional
    for c in OPTIONAL_NUMERIC_COLS:
        if c not in df.columns:
            print(f"  Optional column '{c}' not present – filled with blanks.")
            df[c] = np.nan
    for c in OPTIONAL_TEXT_COLS:
        if c not in df.columns:
            print(f"  Optional column '{c}' not present – filled with '--'.")
            df[c] = "--"

    print(f"  {len(df):,} rows loaded")

    # --- Numeric coercion -------------------------------------------
    # Converts every numeric column, neutralises Excel error text and
    # flags the Discount % rows that held it (see NUMERIC_COLS).
    df = sanitise_numeric_columns(df)

    # --- Text tidy-up on grouping columns ---------------------------
    for c in set(BENCHMARK_KEY + PEER_DIM_COLS + ["Business Partner", IDN_COL]):
        if c in df.columns and c != TIER_COL:
            df[c] = df[c].astype("object").fillna("--").astype(str).str.strip()
            df.loc[df[c] == "", c] = "--"

    filters = {
        "Business Partner":       EXCLUDE_BUSINESS_PARTNER,
        "Material Group":         EXCLUDE_MATERIAL_GROUP,
        "Product Group":          EXCLUDE_PRODUCT_GROUP,
        "Market Segment":         EXCLUDE_MARKET_SEGMENT,
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
        df = df[~((df["Selling UOM Unit Qty"].fillna(0) == 0)
                  & (df["Invoice Total"].fillna(0) == 0))]
        n = pre - len(df)
        if n: print(f"  Excluded {n:,} rows with zero quantity and zero invoice total")

    df = df.reset_index(drop=True)
    n_seg = df.groupby(BENCHMARK_KEY, observed=True).ngroups
    print(f"  {len(df):,} rows after filters  |  "
          f"{df['Business Partner'].nunique()} customers  |  "
          f"{n_seg} segments ({BENCHMARK_LABEL})")
    return df


# ═══════════════════════════════════════════════════════════════════
# DISCOUNT NORMALISATION
# ═══════════════════════════════════════════════════════════════════
def prepare_discount(df: pd.DataFrame) -> pd.DataFrame:
    """Create the analysed metric `Discount_Pct` (positive fraction off
    list) plus the validity flag `Benchmarkable`.

    Precedence:
      1. Recompute from List Price Total and Invoice Total (preferred –
         immune to sign conventions and rounding in the source column).
      2. Fall back to the source Discount % column, sign-corrected, for
         rows where list price is missing or non-positive.
    """
    print("Preparing discount metric ...")

    src = df["Discount %"].copy()
    # Source columns are often expressed in points (e.g. -47.1) rather
    # than as a fraction; detect and rescale.
    finite = src.replace([np.inf, -np.inf], np.nan).dropna()
    if len(finite) and finite.abs().median() > 1.5:
        src = src / 100.0
        print("  Source Discount % appeared to be in percentage points – rescaled to a fraction.")

    sign = DISCOUNT_SIGN
    if sign == "auto":
        finite = src.replace([np.inf, -np.inf], np.nan).dropna()
        neg_share = (finite < 0).mean() if len(finite) else 0.0
        sign = "negative" if neg_share > 0.5 else "positive"
        print(f"  Source Discount % sign detected as '{sign}' "
              f"({neg_share:.0%} of values negative).")
    if sign == "negative":
        src = -src

    # Excel error text in the source column.  Substituted only now, so
    # the points/fraction rescale above and the sign detection are both
    # based on genuine data.
    text_err = df.get(DISCOUNT_TEXT_FLAG)
    if text_err is not None:
        text_err = _as_bool(text_err)
        n_err = int(text_err.sum())
        if n_err:
            if DISCOUNT_TEXT_ERROR_TO_ZERO:
                src = src.where(~text_err, 0.0)
                print(f"  {n_err:,} source Discount % value(s) held non-numeric "
                      f"text – treated as 0%.")
            else:
                print(f"  {n_err:,} source Discount % value(s) held non-numeric "
                      f"text – left blank (recomputed discount used where possible).")

    lp  = df["List Price Total"]
    inv = df["Invoice Total"]
    calc = pd.Series(np.nan, index=df.index, dtype=float)
    usable = lp > 0
    calc[usable] = 1.0 - (inv[usable] / lp[usable])

    if RECOMPUTE_DISCOUNT:
        disc = calc.where(calc.notna(), src)
        src_used = calc.isna() & src.notna()
        if src_used.any():
            print(f"  {src_used.sum():,} rows had no usable list price – "
                  f"source Discount % used instead.")
    else:
        disc = src.where(src.notna(), calc)

    disc = disc.replace([np.inf, -np.inf], np.nan)
    df["Discount_Pct"] = disc

    # The sanitised source column is what the detail sheets display, so
    # keep it consistent with what the analysis actually used, then drop
    # the internal flag.
    if text_err is not None:
        if DISCOUNT_TEXT_ERROR_TO_ZERO:
            df.loc[text_err, "Discount %"] = 0.0
        df.drop(columns=[DISCOUNT_TEXT_FLAG], inplace=True)

    # A row can only drive a benchmark when its discount is meaningful
    # AND it carries a positive list price (needed to value the uplift).
    in_range = disc.between(DISCOUNT_VALID_MIN, DISCOUNT_VALID_MAX)
    df["Benchmarkable"] = in_range & (lp > 0)

    bad = int((~df["Benchmarkable"]).sum())
    print(f"  {int(df['Benchmarkable'].sum()):,} benchmarkable rows  |  "
          f"{bad:,} excluded from benchmarks "
          f"(no positive list price, or discount outside "
          f"{DISCOUNT_VALID_MIN:.0%}..{DISCOUNT_VALID_MAX:.0%})")
    if df["Benchmarkable"].any():
        d = df.loc[df["Benchmarkable"], "Discount_Pct"]
        print(f"  Discount off list – median {d.median():.1%}, "
              f"P25 {np.percentile(d, 25):.1%}, P75 {np.percentile(d, 75):.1%}")
    return df


# ═══════════════════════════════════════════════════════════════════
# REVENUE TIER ASSIGNMENT
# ═══════════════════════════════════════════════════════════════════
def assign_revenue_tiers(df: pd.DataFrame) -> pd.DataFrame:
    print(f"Assigning {NUM_TIERS} revenue tiers ...")
    cr = (df.groupby("Business Partner")["Invoice Total"]
          .sum().reset_index(name="_CustRev"))
    labels = [f"Tier {i}" for i in range(1, NUM_TIERS + 1)]
    n_unique = cr["_CustRev"].rank(method="first").nunique()
    if n_unique < NUM_TIERS:
        cr["Revenue_Tier"] = labels[0]
    else:
        cr["Revenue_Tier"] = pd.qcut(
            cr["_CustRev"].rank(method="first"), NUM_TIERS, labels=labels[::-1])
    cr["Revenue_Tier"] = cr["Revenue_Tier"].astype(str)
    for t in labels:
        m = cr["Revenue_Tier"] == t
        print(f"  {t}: {m.sum():,} customers, "
              f"${cr.loc[m, '_CustRev'].sum():,.0f}")
    return df.merge(cr[["Business Partner", "_CustRev", "Revenue_Tier"]],
                    on="Business Partner", how="left")


# ═══════════════════════════════════════════════════════════════════
# 1. MARKET (SEGMENT-WIDE) DISCOUNT BANDS
# ═══════════════════════════════════════════════════════════════════
def build_market_bands(df: pd.DataFrame) -> pd.DataFrame:
    """Discount percentiles per segment, across ALL customers."""
    print(f"Building market discount bands by {BENCHMARK_LABEL} ...")
    bm = df[df["Benchmarkable"]]

    g = bm.groupby(BENCHMARK_KEY, observed=True)
    b = g.agg(
        Record_Lines      = ("Discount_Pct", "size"),
        Customers         = ("Business Partner", "nunique"),
        Total_Revenue     = ("Invoice Total", "sum"),
        Total_List        = ("List Price Total", "sum"),
        Total_Units       = ("Selling UOM Unit Qty", "sum"),
        Mean_Disc         = ("Discount_Pct", "mean"),
        P10_Disc          = ("Discount_Pct", lambda s: _pct(s, 10)),
        P25_Disc          = ("Discount_Pct", lambda s: _pct(s, LOWER_PCTL)),
        Median_Disc       = ("Discount_Pct", lambda s: _pct(s, 50)),
        P75_Disc          = ("Discount_Pct", lambda s: _pct(s, UPPER_PCTL)),
        P90_Disc          = ("Discount_Pct", lambda s: _pct(s, 90)),
        Min_Disc          = ("Discount_Pct", "min"),
        Max_Disc          = ("Discount_Pct", "max"),
        Std_Disc          = ("Discount_Pct", "std"),
        Mean_ASP          = ("Invoice ASP", "mean"),
        Median_ASP        = ("Invoice ASP", "median"),
        Mean_Margin_Pct   = ("Margin %", "mean"),
        Median_Margin_Pct = ("Margin %", "median"),
    ).reset_index()

    # Revenue-weighted discount: the discount the segment actually
    # realises in aggregate, as opposed to the average of its lines.
    b["Wtd_Disc"]        = 1.0 - _safe_div(b["Total_Revenue"], b["Total_List"])
    b["Disc_Spread"]     = b["P90_Disc"] - b["P10_Disc"]
    b["Disc_IQR"]        = b["P75_Disc"] - b["P25_Disc"]
    b["Disc_Range"]      = b["Max_Disc"] - b["Min_Disc"]
    b["Segment_Valid"]   = b["Customers"] >= MIN_SEGMENT_CUSTOMERS

    thin = int((~b["Segment_Valid"]).sum())
    print(f"  {len(b):,} segments, {len(b) - thin:,} with at least "
          f"{MIN_SEGMENT_CUSTOMERS} customers "
          f"({thin:,} thin segments carry no uplift)")
    return b.sort_values("Total_Revenue", ascending=False)


# ═══════════════════════════════════════════════════════════════════
# 2. PEER BANDS  (cascading, within each segment)
# ═══════════════════════════════════════════════════════════════════
# The cascade levels (_PEER_LEVELS) are generated from PEER_DIMENSIONS
# in the CONFIGURATION section — nothing below needs editing when
# dimensions are added or removed.


def _band_at_level(df, group_cols, prefix):
    """Compute discount statistics (min, P10, P25, median, mean, P75,
    P90, max) and customer count per peer group WITHIN each segment."""
    bm = df[df["Benchmarkable"]]
    g = bm.groupby(BENCHMARK_KEY + group_cols, observed=True)
    out = g.agg(
        _custs  = ("Business Partner", "nunique"),
        _min    = ("Discount_Pct", "min"),
        _p10    = ("Discount_Pct", lambda s: _pct(s, 10)),
        _p25    = ("Discount_Pct", lambda s: _pct(s, LOWER_PCTL)),
        _median = ("Discount_Pct", lambda s: _pct(s, 50)),
        _mean   = ("Discount_Pct", "mean"),
        _p75    = ("Discount_Pct", lambda s: _pct(s, UPPER_PCTL)),
        _p90    = ("Discount_Pct", lambda s: _pct(s, 90)),
        _max    = ("Discount_Pct", "max"),
        _rev    = ("Invoice Total", "sum"),
        _list   = ("List Price Total", "sum"),
    ).reset_index()
    out["_wtd"] = 1.0 - _safe_div(out["_rev"], out["_list"])
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
        "_wtd":    f"{prefix}_Wtd",
        "_rev":    f"{prefix}_Rev",
        "_list":   f"{prefix}_List",
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
    print(f"Building peer discount bands ({len(_PEER_LEVELS)}-level cascade on "
          f"{len(_PEER_DIMS)} dimension(s): {', '.join(PEER_DIM_COLS)}) "
          f"within {BENCHMARK_LABEL} ...")
    frames = {}
    for label, prefix, grp_cols in _PEER_LEVELS:
        frames[prefix] = _band_at_level(df, grp_cols, prefix)
        v = frames[prefix][f"{prefix}_Valid"].sum()
        print(f"  {prefix:>3} {label:<26}: {len(frames[prefix]):>6,} groups, {v:>5,} valid")
    return frames


# ═══════════════════════════════════════════════════════════════════
# 3. FLAG RECORDS  (market + peer benchmarks)
# ═══════════════════════════════════════════════════════════════════
def _uplift(list_total, actual_disc, bench_disc, active):
    """Revenue recovered by pulling `actual_disc` back to `bench_disc`,
    holding list price and volume constant."""
    up = pd.Series(0.0, index=list_total.index)
    m = active & (actual_disc > bench_disc) & bench_disc.notna() & (list_total > 0)
    up.loc[m] = list_total.loc[m] * (actual_disc.loc[m] - bench_disc.loc[m])
    return up


def flag_records(df, mkt_bands, peer_frames):
    print("Flagging above-benchmark discount records ...")

    # --- Market benchmarks (per segment) ---
    mkt = mkt_bands[BENCHMARK_KEY + [
        "P10_Disc", "P25_Disc", "Median_Disc", "P75_Disc", "P90_Disc",
        "Mean_Disc", "Wtd_Disc", "Customers", "Record_Lines", "Segment_Valid"]]
    out = df.merge(mkt, on=BENCHMARK_KEY, how="left")
    out["Segment_Valid"] = _as_bool(out["Segment_Valid"])

    # Only benchmarkable lines in sufficiently populated segments are
    # scored; everything else is reported but never flagged.
    active = out["Benchmarkable"] & out["Segment_Valid"]
    d = out["Discount_Pct"]

    out["Price_Band"] = BAND_OK
    out.loc[d > out["Median_Disc"], "Price_Band"] = BAND_ABOVE_MED
    out.loc[d > out["P75_Disc"],    "Price_Band"] = BAND_ABOVE_UPPER
    out.loc[d > out["P90_Disc"],    "Price_Band"] = BAND_ABOVE_P90
    out.loc[out["Invoice Total"] <= 0, "Price_Band"] = BAND_ZERO_REV
    out.loc[~active, "Price_Band"] = BAND_NOT_BM

    # Gap in percentage POINTS (e.g. +8.4 pts more discount than median)
    out["Disc_vs_Mkt_Med_Pts"] = np.where(active, d - out["Median_Disc"], np.nan)
    out["Disc_vs_Mkt_P75_Pts"] = np.where(active, d - out["P75_Disc"],    np.nan)

    out["Mkt_Uplift_Med"] = _uplift(out["List Price Total"], d, out["Median_Disc"], active)
    out["Mkt_Uplift_P75"] = _uplift(out["List Price Total"], d, out["P75_Disc"],    active)

    # --- Peer benchmarks (cascade defined by _PEER_LEVELS) ---
    for label, prefix, grp_cols in _PEER_LEVELS:
        cols = BENCHMARK_KEY + grp_cols + [
            f"{prefix}_Custs", f"{prefix}_P10", f"{prefix}_P25",
            f"{prefix}_Median", f"{prefix}_P75", f"{prefix}_P90",
            f"{prefix}_Mean", f"{prefix}_Wtd", f"{prefix}_Valid"]
        out = out.merge(peer_frames[prefix][cols],
                        on=BENCHMARK_KEY + grp_cols, how="left")

    # Resolve the winning level.  Applied in reverse cascade order
    # (least specific first) so that more specific levels overwrite.
    out["Peer_P25"]    = np.nan
    out["Peer_Median"] = np.nan
    out["Peer_P75"]    = np.nan
    out["Peer_P90"]    = np.nan
    out["Peer_Custs"]  = np.nan
    out["Peer_Level"]  = "Market"

    for label, prefix, grp_cols in reversed(_PEER_LEVELS):
        v = _as_bool(out[f"{prefix}_Valid"])
        out.loc[v, "Peer_P25"]    = out.loc[v, f"{prefix}_P25"]
        out.loc[v, "Peer_Median"] = out.loc[v, f"{prefix}_Median"]
        out.loc[v, "Peer_P75"]    = out.loc[v, f"{prefix}_P75"]
        out.loc[v, "Peer_P90"]    = out.loc[v, f"{prefix}_P90"]
        out.loc[v, "Peer_Custs"]  = out.loc[v, f"{prefix}_Custs"]
        out.loc[v, "Peer_Level"]  = label

    # Market fallback for anything still NaN
    fb = out["Peer_Median"].isna()
    out.loc[fb, "Peer_Median"] = out.loc[fb, "Median_Disc"]
    out.loc[fb, "Peer_P25"]    = out.loc[fb, "P25_Disc"]
    out.loc[fb, "Peer_P75"]    = out.loc[fb, "P75_Disc"]
    out.loc[fb, "Peer_P90"]    = out.loc[fb, "P90_Disc"]

    # Peer discount band (mirrors market band categories)
    out["Peer_Band"] = PEER_BAND_OK
    out.loc[d > out["Peer_Median"], "Peer_Band"] = PEER_BAND_ABOVE_MED
    out.loc[d > out["Peer_P75"],    "Peer_Band"] = PEER_BAND_ABOVE_UPPER
    out.loc[d > out["Peer_P90"],    "Peer_Band"] = PEER_BAND_ABOVE_P90
    out.loc[out["Invoice Total"] <= 0, "Peer_Band"] = BAND_ZERO_REV
    out.loc[~active, "Peer_Band"] = BAND_NOT_BM

    out["Disc_vs_Peer_Med_Pts"] = np.where(active, d - out["Peer_Median"], np.nan)
    out["Disc_vs_Peer_P75_Pts"] = np.where(active, d - out["Peer_P75"],    np.nan)

    out["Peer_Uplift_Med"] = _uplift(out["List Price Total"], d, out["Peer_Median"], active)
    out["Peer_Uplift_P75"] = _uplift(out["List Price Total"], d, out["Peer_P75"],    active)

    # Drop intermediate per-level columns
    _prefixes = tuple(f"{p}_" for _, p, _ in _PEER_LEVELS)
    drop = [c for c in out.columns if c.startswith(_prefixes)]
    out.drop(columns=drop, inplace=True)

    n_flag = int(out["Price_Band"].isin(_ABOVE_UPPER).sum())
    n_pflag = int(out["Peer_Band"].isin(_ABOVE_PEER_UPPER).sum())
    print(f"  {n_flag:,} lines above market P{UPPER_PCTL}  |  "
          f"{n_pflag:,} lines above peer P{UPPER_PCTL}")
    return out


# ═══════════════════════════════════════════════════════════════════
# 4-5. ROLL-UP FUNCTIONS
# ═══════════════════════════════════════════════════════════════════
def customer_benchmark(fl):
    print("Building customer benchmark ...")
    c = fl.groupby(["Business Partner", "Product Group", AVAL_SEG_COL],
                   observed=True).agg(
        Account_Type          = (ACCT_TYPE_COL, "first"),
        GPO                   = (GPO_COL, "first"),
        IDN_Parent            = (IDN_COL, "first"),
        Revenue_Tier          = ("Revenue_Tier", "first"),
        Total_Lines           = ("Discount_Pct", "size"),
        Total_Revenue         = ("Invoice Total", "sum"),
        Total_List            = ("List Price Total", "sum"),
        Total_Std_Margin      = ("Std Margin", "sum"),
        Avg_Discount          = ("Discount_Pct", "mean"),
        Lines_Above_Med       = ("Price_Band",  lambda s: s.isin(_ABOVE_MED).sum()),
        Lines_Above_P75       = ("Price_Band",  lambda s: s.isin(_ABOVE_UPPER).sum()),
        Lines_Zero_Rev        = ("Price_Band",  lambda s: (s == BAND_ZERO_REV).sum()),
        Mkt_Uplift_Med        = ("Mkt_Uplift_Med", "sum"),
        Mkt_Uplift_P75        = ("Mkt_Uplift_P75", "sum"),
        Peer_Uplift_Med       = ("Peer_Uplift_Med", "sum"),
        Peer_Uplift_P75       = ("Peer_Uplift_P75", "sum"),
        Lines_Above_Peer      = ("Peer_Band", lambda s: s.isin(_ABOVE_PEER_MED).sum()),
        Lines_Above_Peer_P75  = ("Peer_Band", lambda s: s.isin(_ABOVE_PEER_UPPER).sum()),
        Avg_Gap_vs_Peer_Med   = ("Disc_vs_Peer_Med_Pts", "mean"),
    ).reset_index()
    c["Wtd_Discount"]   = 1.0 - _safe_div(c["Total_Revenue"], c["Total_List"])
    c["Wtd_Margin_Pct"] = _safe_div(c["Total_Std_Margin"], c["Total_Revenue"])
    c["Pct_Lines_Above_Peer_P75"] = c["Lines_Above_Peer_P75"] / c["Total_Lines"]
    c = c[[
        "Business Partner", "Account_Type", "GPO", "IDN_Parent", "Revenue_Tier",
        "Product Group", AVAL_SEG_COL, "Total_Revenue", "Total_List",
        "Total_Std_Margin", "Wtd_Margin_Pct",
        "Wtd_Discount", "Avg_Discount", "Total_Lines",
        "Lines_Above_Med", "Lines_Above_P75", "Lines_Zero_Rev",
        "Mkt_Uplift_Med", "Mkt_Uplift_P75",
        "Lines_Above_Peer", "Lines_Above_Peer_P75",
        "Peer_Uplift_Med", "Peer_Uplift_P75",
        "Avg_Gap_vs_Peer_Med", "Pct_Lines_Above_Peer_P75",
    ]]
    return c.sort_values("Peer_Uplift_P75", ascending=False)


def segment_uplift(fl):
    """Uplift aggregated at the benchmark level itself — the primary
    'where is the money' view for this analysis."""
    print(f"Building uplift by {BENCHMARK_LABEL} ...")
    sg = fl.groupby(BENCHMARK_KEY, observed=True).agg(
        Customers         = ("Business Partner", "nunique"),
        Total_Lines       = ("Discount_Pct", "size"),
        Total_Revenue     = ("Invoice Total", "sum"),
        Total_List        = ("List Price Total", "sum"),
        Total_Std_Margin  = ("Std Margin", "sum"),
        Lines_Above_Med   = ("Price_Band", lambda s: s.isin(_ABOVE_MED).sum()),
        Lines_Above_P75   = ("Price_Band", lambda s: s.isin(_ABOVE_UPPER).sum()),
        Mkt_Uplift_Med    = ("Mkt_Uplift_Med", "sum"),
        Mkt_Uplift_P75    = ("Mkt_Uplift_P75", "sum"),
        Peer_Uplift_Med   = ("Peer_Uplift_Med", "sum"),
        Peer_Uplift_P75   = ("Peer_Uplift_P75", "sum"),
    ).reset_index()
    sg["Wtd_Discount"]   = 1.0 - _safe_div(sg["Total_Revenue"], sg["Total_List"])
    sg["Wtd_Margin_Pct"] = _safe_div(sg["Total_Std_Margin"], sg["Total_Revenue"])
    sg["Uplift_Pct_Rev"] = _safe_div(sg["Peer_Uplift_P75"], sg["Total_Revenue"])
    return sg.sort_values("Peer_Uplift_P75", ascending=False)


def product_group_uplift(fl):
    print("Building product-group uplift ...")
    ct = fl.groupby("Product Group", observed=True).agg(
        Total_Lines       = ("Discount_Pct", "size"),
        Total_Revenue     = ("Invoice Total", "sum"),
        Total_List        = ("List Price Total", "sum"),
        Total_Std_Margin  = ("Std Margin", "sum"),
        Lines_Above_Med   = ("Price_Band", lambda s: s.isin(_ABOVE_MED).sum()),
        Lines_Above_P75   = ("Price_Band", lambda s: s.isin(_ABOVE_UPPER).sum()),
        Mkt_Uplift_Med    = ("Mkt_Uplift_Med", "sum"),
        Mkt_Uplift_P75    = ("Mkt_Uplift_P75", "sum"),
        Peer_Uplift_Med   = ("Peer_Uplift_Med", "sum"),
        Peer_Uplift_P75   = ("Peer_Uplift_P75", "sum"),
    ).reset_index()
    ct["Wtd_Discount"]   = 1.0 - _safe_div(ct["Total_Revenue"], ct["Total_List"])
    ct["Wtd_Margin_Pct"] = _safe_div(ct["Total_Std_Margin"], ct["Total_Revenue"])
    return ct.sort_values("Peer_Uplift_P75", ascending=False)


# ═══════════════════════════════════════════════════════════════════
# 6. EXECUTIVE SUMMARY KPIs
# ═══════════════════════════════════════════════════════════════════
def build_summary(df, bands, fl, cust, sg, pg):
    n = len(fl)
    tr = df["Invoice Total"].sum()
    tl = df["List Price Total"].sum()
    peer_levels = fl["Peer_Level"].value_counts()
    bm = fl[fl["Benchmarkable"]]
    return dict(
        total_lines=n, total_rev=tr, total_list=tl,
        wtd_disc=(1.0 - tr / tl) if tl else np.nan,
        med_disc=bm["Discount_Pct"].median() if len(bm) else np.nan,
        n_cust=df["Business Partner"].nunique(),
        n_seg=len(bands),
        n_seg_valid=int(bands["Segment_Valid"].sum()),
        not_benchmarked=int((fl["Price_Band"] == BAND_NOT_BM).sum()),
        above_med=int(fl["Price_Band"].isin(_ABOVE_MED).sum()),
        above_p75=int(fl["Price_Band"].isin(_ABOVE_UPPER).sum()),
        zero=int((fl["Price_Band"] == BAND_ZERO_REV).sum()),
        mkt_up_med=fl["Mkt_Uplift_Med"].sum(),
        mkt_up_p75=fl["Mkt_Uplift_P75"].sum(),
        peer_up_med=fl["Peer_Uplift_Med"].sum(),
        peer_up_p75=fl["Peer_Uplift_P75"].sum(),
        peer_above_med=int(fl["Peer_Band"].isin(_ABOVE_PEER_MED).sum()),
        peer_above_p75=int(fl["Peer_Band"].isin(_ABOVE_PEER_UPPER).sum()),
        peer_levels=peer_levels,
        top5_cust=cust.groupby("Business Partner")["Peer_Uplift_P75"].sum()
            .nlargest(5).reset_index().values.tolist(),
        top5_seg=sg.head(5)[BENCHMARK_KEY + ["Peer_Uplift_P75"]]
            .assign(_lbl=lambda x: x[BENCHMARK_KEY].agg(" | ".join, axis=1))
            [["_lbl", "Peer_Uplift_P75"]].values.tolist(),
        top5_pg=pg.head(5)[["Product Group", "Peer_Uplift_P75"]].values.tolist(),
        top5_spread=bands[(bands["Record_Lines"] >= 10) & bands["Segment_Valid"]]
            .nlargest(5, "Disc_Spread")
            .assign(_lbl=lambda x: x[BENCHMARK_KEY].agg(" | ".join, axis=1))
            [["_lbl", "Disc_Spread", "Median_Disc", "Record_Lines", "Total_Revenue"]]
            .values.tolist(),
    )


# ═══════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ═══════════════════════════════════════════════════════════════════
def write_output(path, bands, fl, cust, sg, pg, S):
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
    PT = wb.add_format({"num_format":"+0.0\"pts\";-0.0\"pts\";0.0\"pts\"",
                         "font_name":"Arial","font_size":10})
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
        mc     = set of 0-indexed market-comparison columns (cyan header)
        pc     = set of 0-indexed peer-comparison columns (blue header)
        groups = list of (first_col, last_col) ranges to place in a
                 collapsible outline group so they can be hidden."""
        df.to_excel(wr, sheet_name=name, startrow=1, index=False, header=False)
        ws = wr.sheets[name]; ws.set_tab_color(tab)
        ws.freeze_panes(1, fz); ws.autofilter(0, 0, len(df), len(df.columns)-1)
        mc, pc = mc or set(), pc or set()
        grouped = set()
        for first, last in (groups or []):
            grouped.update(range(first, last + 1))
        for i, (cn, w, f) in enumerate(cols):
            hdr = HM if i in mc else (HP if i in pc else H)
            ws.write(0, i, cn, hdr)
            if i in grouped:
                ws.set_column(i, i, w, f, {"level": 1})
            else:
                ws.set_column(i, i, w, f)

    # ────────────── Information ──────────────────────────────────────
    wi = wb.add_worksheet("Information")
    wi.hide_gridlines(2); wi.set_tab_color("#305496")
    wi.set_column("A:A",4); wi.set_column("B:B",30); wi.set_column("C:C",95)
    r=1; wi.merge_range(r,1,r,2,"Discount Band Benchmarking Analysis",TI); r+=1
    wi.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=1
    wi.write(r,2,f"Benchmark level: {BENCHMARK_LABEL}",NT); r+=2
    wi.merge_range(r,1,r,2,"Workbook Sheets",SB); r+=1
    wi.write(r,1,"Sheet",IH); wi.write(r,2,"Description",IH); r+=1
    for nm,ds in [
        ("Information","Overview of each tab and key term definitions."),
        ("Executive Summary","Headline KPIs under market and peer discount benchmarks."),
        ("Segment Discount Bands",f"Percentile discount-off-list benchmarks for every "
         f"{BENCHMARK_LABEL} combination, across all customers."),
        ("Segment Uplift",f"Uplift and realised discount aggregated by {BENCHMARK_LABEL} — "
         "the primary view of where the discount leakage sits."),
        ("Customer Benchmark","One row per customer x product group x Avalanche Segment with "
         "revenue, realised discount, margin, above-benchmark line counts and dual uplift."),
        ("Flagged Records",f"Lines discounted above the market P{UPPER_PCTL} or the peer "
         f"P{UPPER_PCTL}, with market and peer benchmarks side by side.  Includes the Peer "
         "Level column showing which cascade level supplied the benchmark for each line."),
        ("Uplift by Product Group","Market and peer uplift aggregated by Product Group."),
        ("Column colours","Market-comparison columns use a cyan header; peer-comparison columns "
         "use a blue header.  Market columns are placed in a collapsible outline group — use the "
         "-/+ toggle above the column headers to hide them."),
    ]:
        wi.write(r,1,nm,IC); wi.write(r,2,ds,IC); wi.set_row(r,48); r+=1
    r+=1; wi.merge_range(r,1,r,2,"Key Terms",SB); r+=1
    wi.write(r,1,"Term",IH); wi.write(r,2,"Definition",IH); r+=1
    for tm,df_ in [
        ("Discount %","Discount off list price, expressed as a positive fraction: "
         "1 - (Invoice Total / List Price Total).  Recomputed from list and invoice values "
         "so that source sign conventions do not affect the result.  A HIGHER discount is a "
         "worse commercial outcome."),
        ("Segment",f"The benchmark unit: a unique combination of {BENCHMARK_LABEL}.  "
         "All comparisons are made within a segment, never across segments."),
        ("Market Benchmark","Discount percentiles across ALL customers within the segment."),
        ("Peer Benchmark",f"Discount percentiles within the segment among customers sharing "
         f"one or more of: {_DIM_TEXT}.  The most specific grouping with enough peers is used "
         "for each line."),
        ("Peer Level",f"Which of the {len(_PEER_LEVELS)} cascade levels supplied the benchmark, "
         "or 'Market' where no peer group met the minimum size.  The cascade runs from "
         f"'{_CASCADE_FIRST}' (most specific) down to "
         f"'{_CASCADE_LAST}'.  See the separate Peer Discount Bands workbook for the full "
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
        (f"Min Segment Customers ({MIN_SEGMENT_CUSTOMERS})",
         f"A segment needs at least {MIN_SEGMENT_CUSTOMERS} distinct customers before its "
         "benchmark is applied.  Lines in thinner segments are reported as 'Not Benchmarked' "
         "and generate no uplift."),
        ("Discount Band",f"Classification of the line's discount against the relevant benchmark "
         f"distribution: At/Below Median, Above Median, Above P{UPPER_PCTL}, Above P90, "
         "Zero Revenue, or Not Benchmarked."),
        ("P10–P90","Percentile benchmarks of discount off list within the relevant population.  "
         "P90 is the deepest-discount end of the distribution."),
        ("Weighted Discount","1 - (Total Invoice / Total List) for the group — the discount "
         "actually realised in aggregate, as opposed to the average of the individual lines."),
        ("Discount Spread","P90 minus P10, in percentage points.  Higher = more inconsistent "
         "discounting within the segment."),
        ("Gap vs Benchmark","Difference between the line's discount and the benchmark discount, "
         "in percentage POINTS.  Positive = discounted more deeply than the benchmark."),
        (f"Uplift to P{UPPER_PCTL}",f"Incremental revenue if lines discounted more deeply than "
         f"the P{UPPER_PCTL} benchmark were pulled back to it: List Price Total x (actual "
         f"discount - benchmark discount), holding list price and volume constant.  The "
         "conservative case."),
        ("Uplift to Median","Same calculation against the median benchmark — the more "
         "ambitious case, since it targets every line above the midpoint."),
        ("Not Benchmarked","Lines with no positive list price, a discount outside the plausible "
         f"range ({DISCOUNT_VALID_MIN:.0%} to {DISCOUNT_VALID_MAX:.0%}), or sitting in a thin "
         "segment.  Retained for completeness, excluded from all percentiles and uplift."),
    ]:
        wi.write(r,1,tm,IT); wi.write(r,2,df_,IC); wi.set_row(r,38); r+=1

    # ────────────── Executive Summary ────────────────────────────────
    ws = wb.add_worksheet("Executive Summary")
    ws.hide_gridlines(2); ws.set_tab_color("#1F3864")
    ws.set_column("A:A",4); ws.set_column("B:B",50)
    ws.set_column("C:C",22); ws.set_column("D:D",22); ws.set_column("E:E",22)
    r=1; ws.merge_range(r,1,r,4,"Discount Band Benchmarking Analysis",TI); r+=1
    ws.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=1
    ws.write(r,1,f"Benchmark level: {BENCHMARK_LABEL}",NT); r+=2

    ws.merge_range(r,1,r,4,"Dataset Overview",SB); r+=1
    for l,v,f in [("Total Records",S["total_lines"],KN),
                   ("Total Invoice Revenue",S["total_rev"],KV),
                   ("Total List Value",S["total_list"],KV),
                   ("Realised Discount off List",S["wtd_disc"],KP),
                   ("Median Line Discount",S["med_disc"],KP),
                   ("Unique Customers",S["n_cust"],KN),
                   ("Segments Benchmarked",S["n_seg_valid"],KN),
                   ("Segments Total",S["n_seg"],KN),
                   ("Lines Not Benchmarked",S["not_benchmarked"],KN),
                   ("Revenue Tiers",NUM_TIERS,KN)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Market Benchmark (segment-wide)",SB); r+=1
    for l,v,f in [("Lines Above Market Median Discount",S["above_med"],KN),
                   (f"Lines Above Market P{UPPER_PCTL} Discount",S["above_p75"],KN),
                   ("Lines at Zero / Negative Revenue",S["zero"],KN),
                   ("Mkt Uplift to Median",S["mkt_up_med"],KV),
                   (f"Mkt Uplift to P{UPPER_PCTL}",S["mkt_up_p75"],KV),
                   ("  -- as % of Revenue",S["mkt_up_p75"]/max(S["total_rev"],1),KP)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Peer Benchmark (within segment)",SB); r+=1
    for l,v,f in [("Lines Above Peer Median Discount",S["peer_above_med"],KN),
                   (f"Lines Above Peer P{UPPER_PCTL} Discount",S["peer_above_p75"],KN),
                   ("Lines at Zero / Negative Revenue",S["zero"],KN),
                   ("Peer Uplift to Median",S["peer_up_med"],KV),
                   (f"Peer Uplift to P{UPPER_PCTL}",S["peer_up_p75"],KV),
                   ("  -- as % of Revenue",S["peer_up_p75"]/max(S["total_rev"],1),KP)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1

    # Peer-level breakdown
    r+=1; ws.merge_range(r,1,r,4,"Peer Level Distribution",SB); r+=1
    ws.write(r,1,"Peer Level",H); ws.write(r,2,"Records",H); r+=1
    for lvl in _PEER_LEVEL_LABELS:
        cnt = int(S["peer_levels"].get(lvl, 0))
        ws.write(r,1,lvl,T); ws.write(r,2,cnt,N); r+=1

    r+=1; ws.merge_range(r,1,r,4,f"Top 5 Customers by Peer Uplift to P{UPPER_PCTL}",SB); r+=1
    ws.write(r,1,"Customer",H); ws.write(r,2,"Peer Uplift",H); r+=1
    for nm,v in S["top5_cust"]:
        ws.write(r,1,nm,T); ws.write(r,2,v,C); r+=1

    r+=1; ws.merge_range(r,1,r,4,f"Top 5 Segments by Peer Uplift to P{UPPER_PCTL}",SB); r+=1
    ws.write(r,1,BENCHMARK_LABEL,H); ws.write(r,2,"Peer Uplift",H); r+=1
    for nm,v in S["top5_seg"]:
        ws.write(r,1,nm,T); ws.write(r,2,v,C); r+=1

    r+=1; ws.merge_range(r,1,r,4,f"Top 5 Product Groups by Peer Uplift to P{UPPER_PCTL}",SB); r+=1
    ws.write(r,1,"Product Group",H); ws.write(r,2,"Peer Uplift",H); r+=1
    for nm,v in S["top5_pg"]:
        ws.write(r,1,nm,T); ws.write(r,2,v,C); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Top 5 Most Inconsistent Segments (>=10 lines)",SB); r+=1
    ws.write(r,1,"Segment",H); ws.write(r,2,"Spread P90-P10",H)
    ws.write(r,3,"Median Discount",H); ws.write(r,4,"Revenue",H); r+=1
    for seg,spread,med,lines,rev in S["top5_spread"]:
        ws.write(r,1,seg,T); ws.write(r,2,spread,P)
        ws.write(r,3,med,P); ws.write(r,4,rev,C); r+=1

    r+=2; ws.merge_range(r,1,r,4,"Methodology",SB); r+=1
    ws.merge_range(r,1,r+5,4,
        "The analysed metric is discount off list price, 1 - (Invoice Total / List Price "
        "Total), expressed as a positive fraction; a higher discount is a worse commercial "
        f"outcome.  Benchmarks are computed within each {BENCHMARK_LABEL} segment.  "
        "Market benchmarks: discount percentiles across ALL customers in the segment.  "
        f"Peer benchmarks: discount percentiles within the segment among customers sharing "
        f"one or more of {_DIM_TEXT}.  A "
        f"{len(_PEER_LEVELS)}-level cascade applies the most specific grouping that has at "
        f"least {MIN_PEER_GROUP} distinct customers, running from "
        f"'{_CASCADE_FIRST}' down to '{_CASCADE_LAST}', "
        "falling back to the segment-wide market benchmark where no peer group qualifies.  "
        f"Segments with fewer than {MIN_SEGMENT_CUSTOMERS} customers, lines without a "
        "positive list price, and lines whose discount falls outside "
        f"{DISCOUNT_VALID_MIN:.0%}..{DISCOUNT_VALID_MAX:.0%} are reported but excluded from "
        "percentiles and uplift.  "
        f"Revenue tiers ({NUM_TIERS}): Tier 1 = largest customers by total invoice revenue.  "
        "Uplift = List Price Total x (actual discount - benchmark discount) for lines "
        "discounted more deeply than the benchmark, holding list price and volume constant.  "
        "Full peer benchmark detail for every cascade level is provided in the accompanying "
        "Peer Discount Bands workbook.", NT)

    # ────────────── Segment Discount Bands ───────────────────────────
    bo = bands[BENCHMARK_KEY + [
        "Record_Lines","Customers","Total_Revenue","Total_List","Total_Units",
        "Wtd_Disc","Mean_Disc","Min_Disc","P10_Disc","P25_Disc","Median_Disc",
        "P75_Disc","P90_Disc","Max_Disc","Std_Disc","Disc_IQR","Disc_Spread",
        "Mean_Margin_Pct","Median_Margin_Pct","Segment_Valid"]].copy()
    bo["Segment_Valid"] = bo["Segment_Valid"].map({True:"Yes", False:"No"})
    bcols = [(k, 26, T) for k in BENCHMARK_KEY] + [
        ("Lines",8,N),("Customers",10,N),
        ("Total Revenue",16,C),("Total List",16,C),("Total Units",12,N),
        ("Weighted Discount",14,P2),("Mean Discount",12,P2),
        ("Min Discount",12,P2),("P10 Discount",12,P2),("P25 Discount",12,P2),
        ("Median Discount",14,P2),(f"P{UPPER_PCTL} Discount",12,P2),
        ("P90 Discount",12,P2),("Max Discount",12,P2),
        ("Std Dev",10,P2),("IQR (P75-P25)",13,P2),("Spread (P90-P10)",15,P2),
        ("Avg Margin %",12,P),("Median Margin %",14,P),("Benchmarked",12,T),
    ]
    _ws(bo, "Segment Discount Bands", bcols, tab="#2E75B6")
    _spread_col = len(bcols) - 4
    wr.sheets["Segment Discount Bands"].conditional_format(
        1,_spread_col,len(bo),_spread_col,{
        "type":"3_color_scale","min_color":"#63BE7B","mid_color":"#FFEB84","max_color":"#F8696B"})

    # ────────────── Segment Uplift ───────────────────────────────────
    so = sg[BENCHMARK_KEY + [
        "Customers","Total_Lines","Total_Revenue","Total_List","Wtd_Discount",
        "Wtd_Margin_Pct","Lines_Above_Med","Lines_Above_P75",
        "Mkt_Uplift_Med","Mkt_Uplift_P75","Peer_Uplift_Med","Peer_Uplift_P75",
        "Uplift_Pct_Rev"]].copy()
    nk = len(BENCHMARK_KEY)
    scols = [(k, 26, T) for k in BENCHMARK_KEY] + [
        ("Customers",10,N),("Lines",8,N),("Total Revenue",16,C),("Total List",16,C),
        ("Weighted Discount",14,P2),("Wtd Margin %",12,P),
        ("Lines Above Mkt Med",16,N),(f"Lines Above Mkt P{UPPER_PCTL}",16,N),
        ("Mkt Uplift to Median",18,C),(f"Mkt Uplift to P{UPPER_PCTL}",16,C),
        ("Peer Uplift to Median",18,C),(f"Peer Uplift to P{UPPER_PCTL}",16,C),
        ("Peer Uplift % of Rev",16,P),
    ]
    _six = {name: i for i, (name, _, _) in enumerate(scols)}
    _smkt = [n for n in ("Lines Above Mkt Med", f"Lines Above Mkt P{UPPER_PCTL}",
                         "Mkt Uplift to Median", f"Mkt Uplift to P{UPPER_PCTL}")
             if n in _six]
    _speer = [n for n in ("Peer Uplift to Median", f"Peer Uplift to P{UPPER_PCTL}",
                          "Peer Uplift % of Rev") if n in _six]
    _ws(so, "Segment Uplift", scols, tab="#BF8F00",
        mc={_six[n] for n in _smkt},
        pc={_six[n] for n in _speer},
        groups=[(_six[_smkt[0]], _six[_smkt[-1]])])
    _sbar = _six[f"Peer Uplift to P{UPPER_PCTL}"]
    wr.sheets["Segment Uplift"].conditional_format(
        1,_sbar,len(so),_sbar,{"type":"data_bar","bar_color":"#BF8F00"})

    # ────────────── Customer Benchmark ───────────────────────────────
    _ws(cust, "Customer Benchmark", [
        ("Business Partner",38,T),("Account Type",16,T),                 # 0-1
        ("GPO",18,T),("IDN Parent",22,T),                               # 2-3
        ("Revenue Tier",12,T),("Product Group",18,T),                   # 4-5
        ("Avalanche Segment",18,T),                                      # 6
        ("Total Revenue",16,C),("Total List",16,C),                     # 7-8
        ("Total Std Margin",16,C),("Wtd Margin %",12,P),                # 9-10
        ("Weighted Discount",14,P2),("Avg Discount",12,P2),             # 11-12
        ("Total Lines",10,N),                                            # 13
        ("Lines Above Mkt Med",16,N),(f"Lines Above Mkt P{UPPER_PCTL}",16,N),  # 14-15
        ("Lines Zero Revenue",14,N),                                     # 16
        ("Mkt Uplift to Median",18,C),(f"Mkt Uplift to P{UPPER_PCTL}",16,C),   # 17-18
        ("Lines Above Peer Med",16,N),(f"Lines Above Peer P{UPPER_PCTL}",16,N),# 19-20
        ("Peer Uplift to Median",18,C),(f"Peer Uplift to P{UPPER_PCTL}",16,C), # 21-22
        ("Avg Gap vs Peer Med (pts)",18,P2),                             # 23
        (f"% Lines Above Peer P{UPPER_PCTL}",18,P),                      # 24
    ], tab="#548235",
       mc={14,15,16,17,18},               # market comparison columns
       pc={19,20,21,22,23,24},            # peer comparison columns
       groups=[(14,18)])                  # collapsible market block
    wr.sheets["Customer Benchmark"].conditional_format(
        1,22,len(cust),22,{"type":"data_bar","bar_color":"#548235"})

    # ────────────── Flagged Records ──────────────────────────────────
    # Include records above EITHER the market upper percentile or the
    # peer upper percentile so peer uplift totals reconcile to the
    # Executive Summary.
    bel = fl[fl["Price_Band"].isin(_ABOVE_UPPER)
             | fl["Peer_Band"].isin(_ABOVE_PEER_UPPER)].copy()
    bel.sort_values("Peer_Uplift_P75", ascending=False, inplace=True)
    fo = bel[[
        "Business Partner",ACCT_TYPE_COL,GPO_COL,IDN_COL,"Revenue_Tier",
        AVAL_SEG_COL] + BENCHMARK_KEY + [
        "Selling UOM Unit Qty","List Price Total","Invoice Total","Invoice ASP",
        "Discount_Pct","Margin %","Price_Band",
        "Median_Disc","Disc_vs_Mkt_Med_Pts","P75_Disc","Disc_vs_Mkt_P75_Pts",
        "Peer_Band",
        "Peer_Median","Disc_vs_Peer_Med_Pts","Peer_P75","Disc_vs_Peer_P75_Pts",
        "Peer_Custs","Peer_Level",
        "Mkt_Uplift_Med","Mkt_Uplift_P75","Peer_Uplift_Med","Peer_Uplift_P75",
    ]].copy()
    fo.columns = [
        "Business Partner","Account Type","GPO","IDN Parent","Revenue Tier",
        "Avalanche Segment"] + BENCHMARK_KEY + [
        "Units","List Price Total","Invoice Total","Invoice ASP",
        "Discount %","Margin %","Market Discount Band",
        "Mkt Median Disc","Gap vs Mkt Med",f"Mkt P{UPPER_PCTL} Disc",
        f"Gap vs Mkt P{UPPER_PCTL}",
        "Peer Discount Band",
        "Peer Median Disc","Gap vs Peer Med",f"Peer P{UPPER_PCTL} Disc",
        f"Gap vs Peer P{UPPER_PCTL}",
        "Peer Customers","Peer Level",
        "Mkt Uplift to Median",f"Mkt Uplift to P{UPPER_PCTL}",
        "Peer Uplift to Median",f"Peer Uplift to P{UPPER_PCTL}",
    ]
    fcols = [
        ("Business Partner",38,T),("Account Type",16,T),                 # 0-1
        ("GPO",18,T),("IDN Parent",22,T),("Revenue Tier",12,T),         # 2-4
        ("Avalanche Segment",18,T),                                       # 5
    ] + [(k, 24, T) for k in BENCHMARK_KEY] + [                          # 6..6+nk-1
        ("Units",8,N),("List Price Total",16,C),("Invoice Total",14,C),
        ("Invoice ASP",12,C2),
        ("Discount %",12,P2),("Margin %",10,P),
        ("Market Discount Band",17,T),
        ("Mkt Median Disc",14,P2),("Gap vs Mkt Med",14,P2),
        (f"Mkt P{UPPER_PCTL} Disc",13,P2),(f"Gap vs Mkt P{UPPER_PCTL}",15,P2),
        ("Peer Discount Band",17,T),
        ("Peer Median Disc",14,P2),("Gap vs Peer Med",14,P2),
        (f"Peer P{UPPER_PCTL} Disc",13,P2),(f"Gap vs Peer P{UPPER_PCTL}",15,P2),
        ("Peer Customers",14,N),("Peer Level",18,T),
        ("Mkt Uplift to Median",18,C),(f"Mkt Uplift to P{UPPER_PCTL}",16,C),
        ("Peer Uplift to Median",18,C),(f"Peer Uplift to P{UPPER_PCTL}",16,C),
    ]
    # Column roles are resolved by NAME, not by arithmetic offsets: the
    # benchmark key is variable-length, so hard-coded indices drift as
    # soon as BENCHMARK_KEY changes and silently mis-colour the sheet.
    _ix = {name: i for i, (name, _, _) in enumerate(fcols)}

    def _cols(*names):
        return {_ix[n] for n in names if n in _ix}

    # Cyan (market) block: the market band, its two benchmark/gap pairs,
    # and BOTH market uplift columns.
    mkt_cols = _cols(
        "Market Discount Band",
        "Mkt Median Disc", "Gap vs Mkt Med",
        f"Mkt P{UPPER_PCTL} Disc", f"Gap vs Mkt P{UPPER_PCTL}",
        "Mkt Uplift to Median", f"Mkt Uplift to P{UPPER_PCTL}",
    )
    # Blue (peer) block: the peer band, its benchmark/gap pairs, the peer
    # group size, the cascade level that supplied it, and BOTH peer
    # uplift columns.
    peer_cols = _cols(
        "Peer Discount Band",
        "Peer Median Disc", "Gap vs Peer Med",
        f"Peer P{UPPER_PCTL} Disc", f"Gap vs Peer P{UPPER_PCTL}",
        "Peer Customers", "Peer Level",
        "Peer Uplift to Median", f"Peer Uplift to P{UPPER_PCTL}",
    )
    # Collapsible outline groups cover the market columns only, so the
    # market view can be hidden without touching the peer columns.
    mkt_band_block   = (_ix["Market Discount Band"], _ix[f"Gap vs Mkt P{UPPER_PCTL}"])
    mkt_uplift_block = (_ix["Mkt Uplift to Median"], _ix[f"Mkt Uplift to P{UPPER_PCTL}"])

    _ws(fo,"Flagged Records",fcols,tab="#C00000",
        mc=mkt_cols, pc=peer_cols,
        groups=[mkt_band_block, mkt_uplift_block])
    mb = _ix["Market Discount Band"]
    pb = _ix["Peer Discount Band"]
    wf = wr.sheets["Flagged Records"]
    # Market band conditional formatting
    for v,clr in [(BAND_ZERO_REV,"#C00000"),(BAND_ABOVE_P90,"#ED7D31"),
                   (BAND_ABOVE_UPPER,"#FFC000")]:
        wf.conditional_format(1,mb,len(fo),mb,{
            "type":"text","criteria":"containing","value":v,
            "format":wb.add_format({"bg_color":clr,"font_color":"#FFFFFF",
                                     "font_name":"Arial","font_size":10})})
    # Peer band conditional formatting
    for v,clr in [(BAND_ZERO_REV,"#C00000"),(PEER_BAND_ABOVE_P90,"#ED7D31"),
                   (PEER_BAND_ABOVE_UPPER,"#FFC000"),(PEER_BAND_ABOVE_MED,"#FFEDA0")]:
        wf.conditional_format(1,pb,len(fo),pb,{
            "type":"text","criteria":"containing","value":v,
            "format":wb.add_format({"bg_color":clr,"font_color":"#FFFFFF",
                                     "font_name":"Arial","font_size":10})})

    # ────────────── Uplift by Product Group ──────────────────────────
    po = pg[["Product Group","Total_Lines","Total_Revenue","Total_List",
             "Wtd_Discount","Total_Std_Margin","Lines_Above_Med","Lines_Above_P75",
             "Mkt_Uplift_Med","Mkt_Uplift_P75","Peer_Uplift_Med","Peer_Uplift_P75",
             "Wtd_Margin_Pct"]]
    _ws(po, "Uplift by Product Group", [
        ("Product Group",26,T),("Total Lines",10,N),("Total Revenue",16,C),
        ("Total List",16,C),("Weighted Discount",14,P2),("Total Std Margin",16,C),
        ("Lines Above Mkt Med",16,N),(f"Lines Above Mkt P{UPPER_PCTL}",16,N),
        ("Mkt Uplift to Median",18,C),(f"Mkt Uplift to P{UPPER_PCTL}",16,C),
        ("Peer Uplift to Median",18,C),(f"Peer Uplift to P{UPPER_PCTL}",16,C),
        ("Wtd Margin %",12,P),
    ], tab="#7030A0",
       mc={6,7,8,9}, pc={10,11}, groups=[(6,9)])

    wr.close()
    print(f"Done -- {path}")


# ═══════════════════════════════════════════════════════════════════
# PEER BANDS WORKBOOK
# ═══════════════════════════════════════════════════════════════════
def write_peer_bands(path, peer_frames):
    """Write a workbook with one sheet per cascade level showing the
    full peer-group discount benchmarks, plus an Information sheet.
    Sheets and their grouping columns are driven by _PEER_LEVELS."""
    print(f"Writing  {path} ...")
    wr = pd.ExcelWriter(path, engine="xlsxwriter")
    wb = wr.book

    H  = wb.add_format({"bold":1,"bg_color":"#1F3864","font_color":"#FFFFFF","border":1,
                         "text_wrap":1,"valign":"vcenter","font_name":"Arial","font_size":10})
    C  = wb.add_format({"num_format":"$#,##0",   "font_name":"Arial","font_size":10})
    P2 = wb.add_format({"num_format":"0.00%",    "font_name":"Arial","font_size":10})
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
        df.to_excel(wr, sheet_name=name, startrow=1, index=False, header=False)
        ws = wr.sheets[name]; ws.set_tab_color(tab)
        ws.freeze_panes(1, fz); ws.autofilter(0, 0, len(df), len(df.columns)-1)
        for i, (cn, w, f) in enumerate(cols):
            ws.write(0, i, cn, H); ws.set_column(i, i, w, f)

    # ── Information sheet ──
    wi = wb.add_worksheet("Information")
    wi.hide_gridlines(2); wi.set_tab_color("#305496")
    wi.set_column("A:A",4); wi.set_column("B:B",28); wi.set_column("C:C",95)
    r=1; wi.merge_range(r,1,r,2,"Peer Group Discount Benchmarks",TI); r+=1
    wi.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=1
    wi.write(r,2,f"All benchmarks are computed within {BENCHMARK_LABEL}.",NT); r+=2

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
        wi.write(r,2,f"Discount benchmarks grouped by [{dims}] within each "
                      f"{BENCHMARK_LABEL} segment.  Each row is a unique combination of "
                      f"the segment and the grouping dimensions.  'Valid' indicates "
                      f"whether the group has at least {MIN_PEER_GROUP} distinct customers.",IC)
        wi.set_row(r,48); r+=1

    r+=1; wi.merge_range(r,1,r,2,"Key Terms",SB); r+=1
    wi.write(r,1,"Term",IH); wi.write(r,2,"Definition",IH); r+=1
    for tm,df_ in [
        ("Discount","Discount off list price: 1 - (Invoice Total / List Price Total), "
                    "as a positive fraction.  Higher = deeper discount = worse outcome."),
        ("Peer Customers","Number of distinct customers in this peer group for this segment."),
        ("Peer Min Disc","Shallowest discount within the peer group."),
        ("Peer P10 Disc","10th percentile of discount within the peer group (shallow end)."),
        ("Peer P25 Disc","25th percentile of discount within the peer group."),
        ("Peer Median Disc","50th percentile (median) discount within the peer group."),
        ("Peer Mean Disc","Arithmetic mean of discount within the peer group."),
        (f"Peer P{UPPER_PCTL} Disc",f"{UPPER_PCTL}th percentile of discount within the peer "
                                    "group — the conservative uplift target."),
        ("Peer P90 Disc","90th percentile of discount within the peer group (deep end)."),
        ("Peer Max Disc","Deepest discount within the peer group."),
        ("Peer Wtd Disc","1 - (Total Invoice / Total List) for the peer group — the "
                         "discount actually realised in aggregate."),
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
            f"{prefix}_Min":    "Peer Min Disc",
            f"{prefix}_P10":    "Peer P10 Disc",
            f"{prefix}_P25":    "Peer P25 Disc",
            f"{prefix}_Median": "Peer Median Disc",
            f"{prefix}_Mean":   "Peer Mean Disc",
            f"{prefix}_P75":    f"Peer P{UPPER_PCTL} Disc",
            f"{prefix}_P90":    "Peer P90 Disc",
            f"{prefix}_Max":    "Peer Max Disc",
            f"{prefix}_Wtd":    "Peer Wtd Disc",
            f"{prefix}_Rev":    "Peer Revenue",
            f"{prefix}_List":   "Peer List Value",
            f"{prefix}_Valid":  "Valid",
        }, inplace=True)

        # Convert boolean Valid to Yes/No string for clean Excel display
        df_out["Valid"] = df_out["Valid"].map({True: "Yes", False: "No"})

        # Build column definitions: segment cols + grouping cols + metrics
        cols = [(k, 24, T) for k in BENCHMARK_KEY]
        for gc in grp_cols:
            w = 22 if len(gc) > 15 else 14
            cols.append((gc, w, T))
        cols += [
            ("Peer Customers",  14, N),
            ("Peer Min Disc",   12, P2),
            ("Peer P10 Disc",   12, P2),
            ("Peer P25 Disc",   12, P2),
            ("Peer Median Disc",15, P2),
            ("Peer Mean Disc",  13, P2),
            (f"Peer P{UPPER_PCTL} Disc", 13, P2),
            ("Peer P90 Disc",   12, P2),
            ("Peer Max Disc",   12, P2),
            ("Peer Wtd Disc",   13, P2),
            ("Peer Revenue",    16, C),
            ("Peer List Value", 16, C),
            ("Valid",           8,  T),
        ]

        # Select columns in the right order
        col_order = BENCHMARK_KEY + grp_cols + [
            "Peer Customers", "Peer Min Disc", "Peer P10 Disc", "Peer P25 Disc",
            "Peer Median Disc", "Peer Mean Disc", f"Peer P{UPPER_PCTL} Disc",
            "Peer P90 Disc", "Peer Max Disc", "Peer Wtd Disc",
            "Peer Revenue", "Peer List Value", "Valid"]
        df_out = df_out[col_order]

        # Truncate sheet name to 31 chars (Excel limit)
        sheet_name = label[:31]
        _ws(df_out, sheet_name, cols, tab="#4472C4", fz=1)

        # Colour-code the Valid column
        ws = wr.sheets[sheet_name]
        valid_col = len(cols) - 1
        ws.conditional_format(1, valid_col, len(df_out), valid_col, {
            "type": "text", "criteria": "containing", "value": "Yes",
            "format": BT})
        ws.conditional_format(1, valid_col, len(df_out), valid_col, {
            "type": "text", "criteria": "containing", "value": "No",
            "format": BF})

        print(f"  {label}: {len(df_out):,} rows")

    wr.close()
    print(f"  Peer discount bands workbook saved.")


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════
def main():
    warnings.filterwarnings("ignore", category=UserWarning)
    path = sys.argv[1] if len(sys.argv) > 1 else _pick_file()
    if not os.path.isfile(path): sys.exit(f"Not found: {path}")

    stem    = Path(path).stem
    out_dir = Path(path).parent
    out     = str(out_dir / f"{stem}_DiscountBandAnalysis.xlsx")
    out_pb  = str(out_dir / f"{stem}_PeerDiscountBands.xlsx")

    df = load_data(path)
    check_required_dim_columns(df)       # benchmark key + dimensions must exist
    df = prepare_discount(df)
    df = assign_revenue_tiers(df)
    bands = build_market_bands(df)
    peer_frames = build_peer_bands(df)
    fl = flag_records(df, bands, peer_frames)
    cust = customer_benchmark(fl)
    sg = segment_uplift(fl)
    pgr = product_group_uplift(fl)
    S = build_summary(df, bands, fl, cust, sg, pgr)

    write_output(out, bands, fl, cust, sg, pgr, S)
    write_peer_bands(out_pb, peer_frames)

    print(f"\n{'='*70}")
    print("  HEADLINE RESULTS")
    print(f"{'='*70}")
    print(f"  Benchmark level:            {BENCHMARK_LABEL}")
    print(f"  Total Revenue:              ${S['total_rev']:>14,.0f}")
    print(f"  Total List Value:           ${S['total_list']:>14,.0f}")
    print(f"  Realised Discount off List:  {S['wtd_disc']:>14.1%}")
    print(f"  Market Uplift to P{UPPER_PCTL}:        ${S['mkt_up_p75']:>14,.0f}  "
          f"({S['mkt_up_p75']/max(S['total_rev'],1):.1%})")
    print(f"  Peer Uplift to P{UPPER_PCTL}:          ${S['peer_up_p75']:>14,.0f}  "
          f"({S['peer_up_p75']/max(S['total_rev'],1):.1%})")
    print(f"  Peer Uplift to Median:      ${S['peer_up_med']:>14,.0f}  "
          f"({S['peer_up_med']/max(S['total_rev'],1):.1%})")
    print(f"  Lines Above Peer Median:    {S['peer_above_med']:>14,}  "
          f"({S['peer_above_med']/max(S['total_lines'],1):.1%})")
    print(f"  Lines Not Benchmarked:      {S['not_benchmarked']:>14,}")
    print(f"\n  Peer Level Distribution:")
    for lvl in _PEER_LEVEL_LABELS:
        cnt = int(S["peer_levels"].get(lvl, 0))
        if cnt: print(f"    {lvl:<26s}  {cnt:>7,}  ({cnt/max(S['total_lines'],1):.1%})")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()
