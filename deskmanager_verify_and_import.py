#!/usr/bin/env python3
"""
deskmanager_verify_and_import.py

Phase 1 — Data clean & verify
    Reads  : dm import master report.csv  (or DESKMANAGER_INVENTORY_CSV)
    Cross-refs: Database (shared).xlsx    (or DESKMANAGER_DATABASE_XLSX)
    Writes : dm_import_master_cleaned.csv
             dm_verification_issues.xlsx

Phase 2 — DeskManager sync
    Edits existing DM units with corrected data.
    Adds units that are not yet in DM.
    Writes : dm_sync_report.xlsx

Environment variables
    AUTOMANAGER_USERNAME / AUTOMANAGER_PASSWORD   required for Phase 2
    DESKMANAGER_INVENTORY_CSV     default: dm import master report.csv
    DESKMANAGER_CLEANED_CSV       default: ~/Desktop/dm_import_master_cleaned.csv (Phase 2-only input override)
    DESKMANAGER_DATABASE_XLSX     default: Database (shared).xlsx
    DESKMANAGER_OUTPUT_DIR        default: ~/Desktop
    DESKMANAGER_DRY_RUN=true      Phase 1 only (skip DM)
    DESKMANAGER_PHASE=1|2         run only one phase
    DESKMANAGER_START_FROM        resume from this stock/unit number
    PLAYWRIGHT_HEADLESS=true
"""

import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DESKTOP = Path.home() / "Desktop"

INVENTORY_CSV  = Path(os.getenv("DESKMANAGER_INVENTORY_CSV",  str(DESKTOP / "dm import master report.csv")))
CLEANED_CSV_OVERRIDE = os.getenv("DESKMANAGER_CLEANED_CSV", "").strip()
OUTPUT_DIR     = Path(os.getenv("DESKMANAGER_OUTPUT_DIR",     str(DESKTOP)))
START_FROM     = os.getenv("DESKMANAGER_START_FROM", "").strip()
DRY_RUN        = os.getenv("DESKMANAGER_DRY_RUN", "false").lower() in {"1", "true", "yes"}
PHASE_ONLY     = os.getenv("DESKMANAGER_PHASE", "").strip()   # "1", "2", or "" (both)
ALLOW_VIN_MATCH_EDIT = os.getenv("DESKMANAGER_ALLOW_VIN_MATCH_EDIT", "false").lower() in {"1", "true", "yes"}

# Database: try Desktop copy first, then OneDrive
_DB_CANDIDATES = [
    os.getenv("DESKMANAGER_DATABASE_XLSX", ""),
    str(DESKTOP / "Database (shared).xlsx"),
    str(Path.home() / "Library/CloudStorage/OneDrive-Personal/2026 Database_Share_2026-03-25.xlsx"),
    str(Path.home() / "Library/CloudStorage/OneDrive-Personal/Database/2026 Database_Share_2026-03-25.xlsx"),
]

def _find_database() -> Path:
    for p in _DB_CANDIDATES:
        if p and Path(p).exists():
            return Path(p)
    raise FileNotFoundError("Could not locate Database Excel. Set DESKMANAGER_DATABASE_XLSX.")

DATABASE_XLSX = _find_database()

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(line_buffering=True)
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(line_buffering=True)

# ---------------------------------------------------------------------------
# Lookup tables & normalization constants
# ---------------------------------------------------------------------------

MAKE_ALIASES: Dict[str, str] = {
    "HYUNDAI TRANSLEAD": "Hyundai",
    "HYUNDAI": "Hyundai",
    "UTIL": "Utility",
    "UTILITY TRAILER": "Utility",
    "WABASH NATIONAL": "Wabash",
    "WABASH": "Wabash",
    "GREAT DANE": "Great Dane",
    "TRAILMOBILE": "Trailmobile",
    "VANGUARD": "Vanguard",
    "KENTUCKY": "Kentucky",
    "FRUEHAUF": "Fruehauf",
    "STOUGHTON": "Stoughton",
    "FONTAINE": "Fontaine",
    "EAST MANUFACTURING": "East Manufacturing",
    "ALLOY": "Alloy",
}

# City abbreviation / shorthand fixes
CITY_FIXES: Dict[str, str] = {
    "CRST CA": "Jurupa Valley",
    "CRST-CA": "Jurupa Valley",
    "CRST": "Jurupa Valley",
    "TEC Equipment": "",      # not a real city – clear it
    "Elite": "",              # not a real city – clear it
    "CHINO": "Chino",
    "TOLLESON": "Tolleson",
    "ONTARIO": "Ontario",
}

# City → two-letter state (for State Registered inference)
CITY_STATE: Dict[str, str] = {
    "Nuevo": "CA", "Ontario": "CA", "Jurupa Valley": "CA", "Manteca": "CA",
    "Rialto": "CA", "Chino": "CA", "Madera": "CA", "Fresno": "CA",
    "Belmont": "CA", "Lathrop": "CA", "Orange": "CA", "Albany": "CA",
    "Fullerton": "CA", "Perris": "CA", "San Bernardino": "CA",
    "Las Vegas": "NV",
    "Tolleson": "AZ", "Phoenix": "AZ",
    "Hutchins": "TX", "Cedar Hill, TX": "TX", "Dallas": "TX",
    "Louisville": "KY",
    "East Peoria": "IL",
    "Cedar Rapids": "IA",
    "Birmingham": "AL",
    "Tupelo, MS": "MS",
    "Cincinnati": "OH",
}

# DB "Loc" → DM "Location"
DB_LOC_MAP: Dict[str, str] = {
    "Big Res": "Big Reservoir",
    "Big Res ": "Big Reservoir",
    "Small Res": "Small Reservoir",
    "small Res": "Small Reservoir",
    "North Lot": "North Lot",
    "Row 1": "Row",
    "Row 2": "Row 2",
    "Row 3": "Row 3",
    "ROW 3": "Row 3",
    "Row 4": "Row 4",
    "Row 1-4": "Row",
    "Horseshoe": "Big Reservoir",
    "HorseShoe": "Big Reservoir",
    "Wash Bay / Warehouse": "Wash",
    "Wash": "Wash",
    "sold Row": "Sold Row",
    "Sold Row": "Sold Row",
    "Elite": "",
    "Madera": "",
    "truck": "",
}

# DB "Title Status" → Title-In checkbox (Yes/No)
DB_TITLE_STATUS_TITLE_IN: Dict[str, str] = {
    "S": "Yes",    # Signed
    "T": "Yes",    # Title received
    "Signed": "Yes",
    "ok Title": "Yes",
    "ok title": "Yes",
    "ok Customer": "Yes",
    "ok customer": "Yes",
    "ok DMV": "Yes",
    "ok Reg": "Yes",
    "ok Consign": "Yes",
    "ok signed": "Yes",
    "ok Title VV": "Yes",
    "ok Wall": "Yes",
    "ok wall": "Yes",
    "ok title": "Yes",
    "Not signed": "No",
    "SPS": "No",   # Signed, Pending Something
}

VALID_BODY_STYLES = {"Dry Van", "Reefer", "Flatbed", "Chassis", "Container", "Dolly", "Truck", "Other", "Forklift"}
VALID_VEHICLE_TYPES = {"Trailer", "Truck", "Chassis", "Flatbed"}  # DM accepts these

# Purchase Method canonical value
PURCHASE_METHOD_CANONICAL = "Purchase from Seller"

BOS_DIR_PRIMARY = Path("/Users/seanclarke/Library/CloudStorage/OneDrive-Personal/sean's  docs/Bill of Sale")
BOS_DIR_SECONDARY = Path("/Users/seanclarke/Library/CloudStorage/OneDrive-Personal/Michelle/PC_Riverside Equipment Sales/Bill of Sales")
BOS_SKIP_PATTERN = re.compile(
    r"cancel|cancelled|canceled|credit|wrong\s*vin|wrong\s*title|unit\s*error|wire\s*instruct",
    re.IGNORECASE,
)

BOS_SUPPORTED_EXT = {
    ".pdf", ".tif", ".tiff", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".xlsx", ".xls", ".doc", ".docx"
}

SELLER_HINTS: List[Tuple[re.Pattern, str]] = [
    (re.compile(r"keystone", re.IGNORECASE), "Keystone Utility"),
    (re.compile(r"bob'?s\s*buys", re.IGNORECASE), "Bob's Buys"),
    (re.compile(r"wolf\s*logistics", re.IGNORECASE), "Wolf Logistics"),
    (re.compile(r"don\s*woods", re.IGNORECASE), "Don Woods Auctions"),
    (re.compile(r"knight", re.IGNORECASE), "Knight Trucking Trailer Sales"),
    (re.compile(r"crst", re.IGNORECASE), "CRST"),
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_DATE_FMTS = ["%m/%d/%y", "%m/%d/%Y", "%m-%d-%y", "%m-%d-%Y",
              "%Y-%m-%d", "%Y/%m/%d", "%m.%d.%y", "%m.%d.%Y"]


def normalize_date(val: str) -> str:
    """Parse a date string and return MM/DD/YYYY, or original if unparseable."""
    v = str(val).strip()
    if not v:
        return ""
    for fmt in _DATE_FMTS:
        try:
            dt = datetime.strptime(v, fmt)
            # Two-digit years: python maps 00-68 → 2000-2068, 69-99 → 1969-1999 — correct for our use
            return dt.strftime("%m/%d/%Y")
        except ValueError:
            pass
    return v  # return as-is if unparseable


def normalize_vin(val: str) -> str:
    """Strip spaces/dashes, uppercase, validate 17 chars."""
    v = re.sub(r"[^A-Z0-9]", "", str(val).strip().upper())
    return v


def is_valid_vin(v: str) -> bool:
    return bool(v) and len(v) == 17 and not re.search(r"[IOQ]", v)


def normalize_make(val: str) -> str:
    v = str(val).strip()
    return MAKE_ALIASES.get(v.upper(), v)


def infer_body_style_from_model(model: str) -> str:
    """Return Body Style inferred from model description."""
    m = model.lower()
    if "flatbed" in m or "flat bed" in m:
        return "Flatbed"
    if "reefer" in m or "refrigerated" in m or "refrig" in m:
        return "Reefer"
    if "dry van" in m or "dryvan" in m:
        return "Dry Van"
    if "chassis" in m:
        return "Chassis"
    if "container" in m:
        return "Container"
    if "dolly" in m:
        return "Dolly"
    if "forklift" in m:
        return "Forklift"
    return ""


def bos_date_to_key_number(bos_date: str) -> str:
    """Convert BOS date (MM/DD/YYYY) to key number format MM.DD.YY."""
    v = normalize_date(bos_date)
    if not v:
        return ""
    try:
        dt = datetime.strptime(v, "%m/%d/%Y")
        return dt.strftime("%m.%d.%y")
    except ValueError:
        return ""


def clean_currency(val: str) -> str:
    """Strip $ signs and commas from currency strings."""
    return re.sub(r"[$,]", "", str(val).strip())


def build_description(unit, year, make, model, vin, bos_date) -> str:
    """Rebuild description as two lines:
    1) Stock-Year-Make-Model
    2) Vin - (VIN) (BOS date)
    """
    u = str(unit).strip()
    y = str(year).strip()
    mk = str(make).strip()
    md = str(model).strip()
    v = str(vin).strip()
    bd = normalize_date(str(bos_date).strip())

    base = " ".join([p for p in [y, mk, md] if p]).strip()

    if u and base:
        line1 = f"{u} - {base}"
    elif u:
        line1 = u
    else:
        line1 = base

    paren_parts = []
    if v:
        paren_parts.append(f"({v})")
    if bd:
        paren_parts.append(f"({bd})")

    if paren_parts:
        line2 = f"Vin - {' '.join(paren_parts)}"
        return f"{line1}\n{line2}" if line1 else line2

    return line1


def _extract_invoice_token(*values: str) -> str:
    for val in values:
        txt = str(val).strip()
        if not txt:
            continue
        m = re.search(r"(?<!\d)(\d{5,6})(?!\d)", txt)
        if m:
            return m.group(1)
    return ""


def _bos_quality_score(path: Path, is_primary: bool) -> Tuple[int, float]:
    name = path.name.lower()
    # Lower is better for first item in tuple.
    quality = 3
    if "signed" in name:
        quality = 0
    elif "final" in name:
        quality = 1
    elif path.suffix.lower() == ".pdf":
        quality = 2
    if is_primary:
        quality -= 1
    return (quality, -path.stat().st_mtime)


def _infer_seller_from_path(path: Path) -> str:
    full = str(path).lower()
    for rx, seller in SELLER_HINTS:
        if rx.search(full):
            return seller
    return ""


def build_bos_invoice_map() -> Dict[str, Path]:
    """Map invoice number -> best BOS file from primary+secondary folders (recursive)."""
    invoice_map: Dict[str, Path] = {}
    score_map: Dict[str, Tuple[int, float]] = {}

    for root, is_primary in [(BOS_DIR_PRIMARY, True), (BOS_DIR_SECONDARY, False)]:
        if not root.exists():
            continue
        for f in root.rglob("*"):
            if not f.is_file():
                continue
            if f.suffix.lower() not in BOS_SUPPORTED_EXT:
                continue
            if BOS_SKIP_PATTERN.search(f.name):
                continue
            inv = _extract_invoice_token(f.name)
            if not inv:
                continue
            score = _bos_quality_score(f, is_primary)
            if inv not in invoice_map or score < score_map[inv]:
                invoice_map[inv] = f
                score_map[inv] = score

    return invoice_map


# ---------------------------------------------------------------------------
# Phase 1 – load, cross-ref, clean
# ---------------------------------------------------------------------------

def load_database(path: Path) -> pd.DataFrame:
    """Load the primary database sheet and return a clean DataFrame."""
    xl = pd.ExcelFile(path)
    # Prefer the most recent non-archived sheet
    preferred = [s for s in xl.sheet_names if "database" in s.lower() and "do not" not in s.lower()]
    sheet = preferred[0] if preferred else xl.sheet_names[0]
    print(f"  [DB] Using sheet '{sheet}' from {path.name}")
    df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def build_db_indexes(db: pd.DataFrame):
    """Build VIN and Unit lookup indexes from database."""
    # Normalize column names
    col_map = {}
    for c in db.columns:
        key = c.lower().replace("\n", " ").replace(" ", "").replace("_", "")
        col_map[key] = c

    vin_col   = col_map.get("vin", "VIN")
    unit_col  = col_map.get("unit", "Unit")
    year_col  = col_map.get("year", "Year")
    make_col  = col_map.get("make", "Make")
    model_col = col_map.get("model", "Model")
    city_col  = col_map.get("city", "City")
    loc_col   = col_map.get("loc", "Loc")
    title_col = col_map.get("titlestatus", col_map.get("title status", "Title Status"))
    inv_col   = col_map.get("billofsoleinvoice#", col_map.get("billofsaleinvoice#", "BillofSale\nInvoice #"))

    idx_vin:  Dict[str, dict] = {}
    idx_unit: Dict[str, dict] = {}

    for _, row in db.iterrows():
        rec = {
            "Year":         str(row.get(year_col,  "")).strip(),
            "Make":         str(row.get(make_col,  "")).strip(),
            "Model":        str(row.get(model_col, "")).strip(),
            "City":         str(row.get(city_col,  "")).strip(),
            "Loc":          str(row.get(loc_col,   "")).strip(),
            "TitleStatus":  str(row.get(title_col, "")).strip(),
            "InvoiceNo":    str(row.get(inv_col,   "")).strip(),
        }
        vin  = re.sub(r"[^A-Z0-9]", "", str(row.get(vin_col,  "")).strip().upper())
        unit = str(row.get(unit_col, "")).strip().upper()
        if vin:
            idx_vin[vin]  = rec
        if unit:
            idx_unit[unit] = rec

    return idx_vin, idx_unit


def _inventory_col_map(columns) -> Dict[str, str]:
    col_map: Dict[str, str] = {}
    for col in columns:
        key = str(col).lower().replace("\n", " ").replace(" ", "").replace("_", "")
        col_map[key] = str(col)
    return col_map


def _inventory_col(col_map: Dict[str, str], *keys: str) -> str:
    for key in keys:
        if key in col_map:
            return col_map[key]
    return ""


def _looks_like_match_export(df: pd.DataFrame) -> bool:
    col_map = _inventory_col_map(df.columns)
    return (
        "vehicletype" not in col_map
        and "eqtype" in col_map
        and "billofsaleinvoice#" in col_map
        and "unit" in col_map
        and "vin" in col_map
    )


def infer_length_from_model(model: str) -> str:
    text = str(model).strip().lower()
    if not text:
        return ""

    range_match = re.search(r"\b(\d{2})\s*[-/]\s*(\d{2})\b", text)
    if range_match:
        return f"{range_match.group(1)}/{range_match.group(2)}"

    length_match = re.search(r"\b(\d{2})\s*'?\b", text)
    if length_match:
        return f"{length_match.group(1)}'"

    return ""


def infer_axles_from_model(model: str) -> str:
    text = str(model).strip().lower()
    if not text:
        return ""
    if re.search(r"\bs/?a\b|single axle", text):
        return "Single"

    axle_match = re.search(r"\b([1-4])\s*ax(?:el|le)\b", text)
    if not axle_match:
        return ""

    axle_map = {
        "1": "Single",
        "2": "Tandem",
        "3": "Tridem",
        "4": "Quad",
    }
    return axle_map.get(axle_match.group(1), "")


def _normalize_match_export_inventory(df: pd.DataFrame) -> pd.DataFrame:
    col_map = _inventory_col_map(df.columns)

    unit_col = _inventory_col(col_map, "unit")
    alt_unit_col = _inventory_col(col_map, "altunit")
    vin_col = _inventory_col(col_map, "vin")
    year_col = _inventory_col(col_map, "year")
    eq_type_col = _inventory_col(col_map, "eqtype")
    make_col = _inventory_col(col_map, "make")
    model_col = _inventory_col(col_map, "model")
    invoice_col = _inventory_col(col_map, "billofsaleinvoice#")
    bos_date_col = _inventory_col(col_map, "billofsaleinvoicedtd")
    cost_col = _inventory_col(col_map, "cost")
    miles_col = _inventory_col(col_map, "currentmiles")
    city_col = _inventory_col(col_map, "city")
    sold_to_col = _inventory_col(col_map, "soldto")
    sold_date_col = _inventory_col(col_map, "solddt")
    sold_price_col = _inventory_col(col_map, "saleprice")
    loc_col = _inventory_col(col_map, "loc")
    title_notes_col = _inventory_col(col_map, "titlenotes")
    review_notes_col = _inventory_col(col_map, "reviewnotes")
    validated_data_col = _inventory_col(col_map, "validateddata")

    rows: List[Dict[str, str]] = []
    for _, row in df.iterrows():
        unit = str(row.get(unit_col, "")).strip()
        vin = str(row.get(vin_col, "")).strip()
        year = str(row.get(year_col, "")).strip()
        make = str(row.get(make_col, "")).strip()
        model = str(row.get(model_col, "")).strip()
        bos_date = str(row.get(bos_date_col, "")).strip()
        eq_type = str(row.get(eq_type_col, "")).strip()
        vehicle_type = eq_type if eq_type in VALID_VEHICLE_TYPES else ""
        body_style = infer_body_style_from_model(model)
        sticky_note = next(
            (
                str(row.get(col, "")).strip()
                for col in (title_notes_col, review_notes_col, validated_data_col)
                if col and str(row.get(col, "")).strip()
            ),
            "",
        )

        rows.append({
            "Unit": unit,
            "Alt Stock Number": str(row.get(alt_unit_col, "")).strip(),
            "VIN": vin,
            "Year": year,
            "Vehicle Type": vehicle_type,
            "Make": make,
            "Model": model,
            "Mileage (Current)": str(row.get(miles_col, "")).strip(),
            "Length": infer_length_from_model(model),
            "Body Style": body_style,
            "Axles": infer_axles_from_model(model),
            "Suspension": "",
            "New / Used": "Used",
            "Engine": "",
            "Series": "",
            "Weight": "",
            "Tire Size": "",
            "State Registered": "",
            "Key Number": "",
            "Title-In": "",
            "Bill Of Sale Date": bos_date,
            "Bill Of Sale Number": str(row.get(invoice_col, "")).strip(),
            "City": str(row.get(city_col, "")).strip(),
            "Sold To": str(row.get(sold_to_col, "")).strip(),
            "Sold Date": str(row.get(sold_date_col, "")).strip(),
            "Sold Price": str(row.get(sold_price_col, "")).strip(),
            "Description": build_description(unit, year, make, model, vin, bos_date),
            "Purchase Cost": str(row.get(cost_col, "")).strip(),
            "Purchase Date": bos_date,
            "Purchased From": "",
            "Purchase Method": PURCHASE_METHOD_CANONICAL,
            "Purchase Channel": "",
            "Reference Number": normalize_date(bos_date),
            "Invoice No.": str(row.get(invoice_col, "")).strip(),
            "Status": "In Inventory",
            "Sub Status": "",
            "Location": DB_LOC_MAP.get(str(row.get(loc_col, "")).strip(), ""),
            "Inventory Date": bos_date,
            "Sticky Note": sticky_note,
        })

    return pd.DataFrame(rows)


def phase1_clean(issues: List[dict]) -> pd.DataFrame:
    """Load inventory CSV, clean, cross-reference, return cleaned DataFrame."""
    print(f"\n[Phase 1] Loading inventory: {INVENTORY_CSV}")
    if INVENTORY_CSV.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(INVENTORY_CSV, dtype=str, keep_default_na=False)
    else:
        df = pd.read_csv(INVENTORY_CSV, dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]

    if _looks_like_match_export(df):
        print("  Detected DeskManager match-export schema; normalizing to canonical import columns")
        df = _normalize_match_export_inventory(df)

    print(f"  {len(df)} rows, {len(df.columns)} columns")
    print(f"\n[Phase 1] Loading database: {DATABASE_XLSX}")
    db = load_database(DATABASE_XLSX)
    idx_vin, idx_unit = build_db_indexes(db)
    print(f"  DB VIN index: {len(idx_vin)} entries  |  Unit index: {len(idx_unit)} entries")
    bos_invoice_map = build_bos_invoice_map()
    print(f"  BOS invoice index: {len(bos_invoice_map)} entries")

    def flag(unit, field, old, new, reason):
        issues.append({
            "Unit": unit, "Field": field,
            "Original": old, "Fixed To": new, "Reason": reason
        })

    for i, row in df.iterrows():
        unit = str(row.get("Unit", "")).strip()

        # ── DB cross-reference ──────────────────────────────────────────────
        vin_key  = re.sub(r"[^A-Z0-9]", "", str(row.get("VIN", "")).strip().upper())
        unit_key = unit.upper()
        db_rec   = idx_vin.get(vin_key) or idx_unit.get(unit_key) or {}

        # ── VIN normalization ───────────────────────────────────────────────
        if vin_key:
            clean_vin = vin_key
            if clean_vin != str(row.get("VIN", "")).strip():
                flag(unit, "VIN", row["VIN"], clean_vin, "Normalized (stripped non-alnum, uppercased)")
                df.at[i, "VIN"] = clean_vin
            if not is_valid_vin(clean_vin):
                flag(unit, "VIN", clean_vin, "", f"Invalid VIN (len={len(clean_vin)}, has I/O/Q: {bool(re.search(r'[IOQ]', clean_vin))})")
        else:
            flag(unit, "VIN", "", "", "Missing VIN — could not auto-fill")

        # ── Fill Year / Make / Model from DB where blank ────────────────────
        for field in ("Year", "Make", "Model"):
            db_field_map = {"Year": "Year", "Make": "Make", "Model": "Model"}
            current = str(row.get(field, "")).strip()
            db_val  = db_rec.get(db_field_map[field], "")
            if not current and db_val:
                df.at[i, field] = db_val
                flag(unit, field, "", db_val, "Filled from database")

        # ── Make normalization ──────────────────────────────────────────────
        make = str(df.at[i, "Make"]).strip()
        norm_make = normalize_make(make)
        if norm_make != make and make:
            df.at[i, "Make"] = norm_make
            flag(unit, "Make", make, norm_make, "Normalized make alias")

        # ── Vehicle Type / Body Style fix ───────────────────────────────────
        vtype  = str(df.at[i, "Vehicle Type"]).strip()
        bstyle = str(df.at[i, "Body Style"]).strip()
        model  = str(df.at[i, "Model"]).strip()

        # "Flatbed" in Vehicle Type → should be Trailer; set Body Style = Flatbed
        if vtype == "Flatbed":
            df.at[i, "Vehicle Type"] = "Trailer"
            flag(unit, "Vehicle Type", "Flatbed", "Trailer", "Vehicle Type 'Flatbed' corrected to 'Trailer'")
            if not bstyle or bstyle == "Dry Van":
                df.at[i, "Body Style"] = "Flatbed"
                flag(unit, "Body Style", bstyle, "Flatbed", "Body Style set to Flatbed (Vehicle Type was Flatbed)")
                bstyle = "Flatbed"

        # Body Style blank → infer from model
        if not bstyle and model:
            inferred = infer_body_style_from_model(model)
            if inferred:
                df.at[i, "Body Style"] = inferred
                flag(unit, "Body Style", "", inferred, f"Inferred from model: '{model}'")
                bstyle = inferred

        # Body Style "Dry Van" but model clearly says Flatbed
        if bstyle == "Dry Van" and "flatbed" in model.lower():
            df.at[i, "Body Style"] = "Flatbed"
            flag(unit, "Body Style", "Dry Van", "Flatbed", "Model indicates Flatbed but Body Style was Dry Van")

        # Vehicle Type blank → infer
        vtype = str(df.at[i, "Vehicle Type"]).strip()
        if not vtype:
            if str(df.at[i, "Body Style"]).strip() in {"Dry Van", "Reefer", "Flatbed", "Container"}:
                df.at[i, "Vehicle Type"] = "Trailer"
                flag(unit, "Vehicle Type", "", "Trailer", "Inferred from Body Style")
            elif "truck" in model.lower() or "forklift" in model.lower():
                df.at[i, "Vehicle Type"] = "Truck"
                flag(unit, "Vehicle Type", "", "Truck", "Inferred from Model")

        # ── Axles normalization ─────────────────────────────────────────────
        axles = str(df.at[i, "Axles"]).strip()
        axle_map = {
            "2": "Tandem", "tandem": "Tandem", "ta": "Tandem",
            "1": "Single",  "single": "Single", "s/a": "Single", "sa": "Single",
            "3": "Tridem",  "tridem": "Tridem",
            "4": "Quad",    "quad": "Quad",
        }
        norm_axles = axle_map.get(axles.lower(), axles)
        if norm_axles != axles and axles:
            df.at[i, "Axles"] = norm_axles
            flag(unit, "Axles", axles, norm_axles, "Normalized axles value")

        # ── Date normalization ──────────────────────────────────────────────
        for date_col in ("Bill Of Sale Date", "Purchase Date", "Sold Date", "Inventory Date"):
            raw = str(df.at[i, date_col]).strip()
            if raw:
                norm = normalize_date(raw)
                if norm != raw:
                    df.at[i, date_col] = norm
                    # Only flag if it looks like data changed meaningfully (not just format)
                    if norm == raw.replace("-", "/"):
                        pass  # minor separator change, no need to flag
                    else:
                        flag(unit, date_col, raw, norm, "Normalized date format to MM/DD/YYYY")

        # ── Key Number: derive from BOS Date if missing ─────────────────────
        bos_date = str(df.at[i, "Bill Of Sale Date"]).strip()
        key_num  = str(df.at[i, "Key Number"]).strip()
        if not key_num and bos_date:
            derived = bos_date_to_key_number(bos_date)
            if derived:
                df.at[i, "Key Number"] = derived
                flag(unit, "Key Number", "", derived, "Derived from Bill Of Sale Date")
        elif key_num and bos_date:
            expected = bos_date_to_key_number(bos_date)
            if expected and key_num != expected:
                df.at[i, "Key Number"] = expected
                flag(unit, "Key Number", key_num, expected, "Normalized from Bill Of Sale Date")

        # ── Currency fields ─────────────────────────────────────────────────
        for cur_col in ("Purchase Cost", "Sold Price"):
            raw = str(df.at[i, cur_col]).strip()
            if raw:
                cleaned = clean_currency(raw)
                if cleaned != raw:
                    df.at[i, cur_col] = cleaned
                    flag(unit, cur_col, raw, cleaned, "Stripped $ and commas")

        # ── Purchase Method normalize (always) ───────────────────────────────
        pm = str(df.at[i, "Purchase Method"]).strip()
        if pm != PURCHASE_METHOD_CANONICAL:
            df.at[i, "Purchase Method"] = PURCHASE_METHOD_CANONICAL
            flag(unit, "Purchase Method", pm, PURCHASE_METHOD_CANONICAL, "Normalized to canonical value")

        # ── Purchased From: infer from BOS file match (no blanket CRST) ─────
        pf = str(df.at[i, "Purchased From"]).strip() if "Purchased From" in df.columns else ""
        if not pf:
            invoice_token = _extract_invoice_token(row.get("Invoice No.", ""), row.get("Bill of sale Number", ""), row.get("Bill Of Sale Number", ""))
            if invoice_token and invoice_token in bos_invoice_map:
                seller = _infer_seller_from_path(bos_invoice_map[invoice_token])
                if seller:
                    df.at[i, "Purchased From"] = seller
                    flag(unit, "Purchased From", "", seller, f"Inferred from BOS file: {bos_invoice_map[invoice_token].name}")

        # ── City: fix known bad values ───────────────────────────────────────
        city = str(df.at[i, "City"]).strip()
        fixed_city = CITY_FIXES.get(city, city)
        if not fixed_city and city and city in CITY_FIXES:
            flag(unit, "City", city, "", f"Cleared invalid city value: '{city}'")
            df.at[i, "City"] = ""
        elif fixed_city != city and city:
            df.at[i, "City"] = fixed_city
            flag(unit, "City", city, fixed_city, "Fixed city abbreviation/shorthand")
        # Also fill from DB if blank
        if not str(df.at[i, "City"]).strip() and db_rec.get("City"):
            db_city = db_rec["City"].strip()
            db_city_fixed = CITY_FIXES.get(db_city, db_city)
            if db_city_fixed:
                df.at[i, "City"] = db_city_fixed
                flag(unit, "City", "", db_city_fixed, "Filled from database")

        # ── State Registered: do not infer (requires title verification) ────
        if "State Registered" not in df.columns:
            df["State Registered"] = ""

        # ── Title-In: fill from DB Title Status ──────────────────────────────
        if "Title-In" not in df.columns:
            df["Title-In"] = ""
        title_in = str(df.at[i, "Title-In"]).strip()
        if not title_in and db_rec.get("TitleStatus"):
            mapped = DB_TITLE_STATUS_TITLE_IN.get(db_rec["TitleStatus"].strip(), "")
            if mapped:
                df.at[i, "Title-In"] = mapped
                flag(unit, "Title-In", "", mapped, f"Mapped from DB Title Status: '{db_rec['TitleStatus']}'")

        # ── Location: fill from DB Loc if blank ──────────────────────────────
        loc = str(df.at[i, "Location"]).strip()
        if not loc and db_rec.get("Loc"):
            dm_loc = DB_LOC_MAP.get(db_rec["Loc"].strip(), "")
            if dm_loc:
                df.at[i, "Location"] = dm_loc
                flag(unit, "Location", "", dm_loc, f"Filled from DB Loc: '{db_rec['Loc']}'")

        # ── Reference Number: always mirror BOS date ────────────────────────
        ref = str(df.at[i, "Reference Number"]).strip()
        expected_ref = normalize_date(bos_date) if bos_date else ""
        if expected_ref and ref != expected_ref:
            df.at[i, "Reference Number"] = expected_ref
            flag(unit, "Reference Number", ref, expected_ref, "Set Reference Number from Bill Of Sale Date")

        # ── Status: fill blank ───────────────────────────────────────────────
        status = str(df.at[i, "Status"]).strip()
        if not status:
            df.at[i, "Status"] = "In Inventory"
            flag(unit, "Status", "", "In Inventory", "Defaulted blank Status to In Inventory")

        # ── New / Used: default ───────────────────────────────────────────────
        new_used = str(df.at[i, "New / Used"]).strip()
        if not new_used:
            df.at[i, "New / Used"] = "Used"
            flag(unit, "New / Used", "", "Used", "Defaulted blank New/Used to Used")

        # ── Description: rebuild for every row from row data ────────────────
        desc = str(df.at[i, "Description"]).strip()
        rebuilt = build_description(
            df.at[i, "Unit"], df.at[i, "Year"], df.at[i, "Make"],
            df.at[i, "Model"], df.at[i, "VIN"], df.at[i, "Bill Of Sale Date"]
        )
        if rebuilt != desc:
            df.at[i, "Description"] = rebuilt
            flag(unit, "Description", desc, rebuilt, "Rebuilt from row fields (Year/Make/Model/VIN/BOS date)")

        # ── Inventory Date: fill from BOS date if blank ───────────────────────
        inv_date = str(df.at[i, "Inventory Date"]).strip()
        if not inv_date and bos_date:
            df.at[i, "Inventory Date"] = normalize_date(bos_date)
            flag(unit, "Inventory Date", "", normalize_date(bos_date), "Filled from Bill Of Sale Date")

    # ── Rename columns to DM canonical names ────────────────────────────────
    df = df.rename(columns={
        "Bill of sale Number": "Bill Of Sale Number",
        "Reference Number":    "Reference No.",
        "Inventory Date":      "Date In",
    })

    return df


# ---------------------------------------------------------------------------
# Phase 2 – DeskManager sync
# ---------------------------------------------------------------------------

def build_dm_row(row: pd.Series) -> pd.Series:
    """
    Produce a Series ready for dvf.fill_vehicle_page / prepare_add_vehicle_page.
    Maps column names and sets derived fields.
    """
    dm = row.copy()

    # Stock Number = Unit
    dm["Stock Number"] = str(row.get("Unit", "")).strip()

    # Condition = New / Used if Condition blank
    if not str(dm.get("Condition", "")).strip():
        dm["Condition"] = str(row.get("New / Used", "")).strip()

    # Mileage column normalisation handled by dvf alias_map ("mileagecurrent")
    if "Mileage (Current)" in dm.index and "Mileage(Current)" not in dm.index:
        dm["Mileage(Current)"] = dm["Mileage (Current)"]

    # Status default
    if not str(dm.get("Status", "")).strip():
        dm["Status"] = "In Inventory"

    # Force Description into required two-line format for every Phase 2 update.
    dm["Description"] = build_description(
        row.get("Unit", ""),
        row.get("Year", ""),
        row.get("Make", ""),
        row.get("Model", ""),
        row.get("VIN", ""),
        row.get("Bill Of Sale Date", ""),
    )

    return dm


def build_dm_row_for_existing(dm_row: pd.Series) -> pd.Series:
    """For existing-unit edits, never change identity fields."""
    dm_edit = dm_row.copy()
    dm_edit["Stock Number"] = ""
    dm_edit["VIN"] = ""
    return dm_edit


def phase2_sync(cleaned_df: pd.DataFrame) -> pd.DataFrame:
    """Login and sync all rows to DeskManager."""
    import deskmanager_vehicle_fill as dvf
    from playwright.sync_api import sync_playwright

    dvf.STRICT_FILL = False   # don't abort on unfilled fields
    headless = os.getenv("PLAYWRIGHT_HEADLESS", "false").lower() in {"1", "true", "yes"}

    results: List[Dict] = []
    start_found = not bool(START_FROM)

    with sync_playwright() as p:
        print("[Phase 2] Authenticating to DeskManager...")
        browser, context, page = dvf.create_authenticated_session(p)
        print("[Phase 2] Authenticated. Starting unit sync loop...")
        session_errors = 0

        try:
            for idx, row in cleaned_df.iterrows():
                unit = str(row.get("Unit", "")).strip()

                if not start_found:
                    if unit == START_FROM:
                        start_found = True
                    else:
                        results.append({"Unit": unit, "Status": "Skipped", "Detail": "Before START_FROM"})
                        continue

                dm_row = build_dm_row(row)
                dm_row_existing = build_dm_row_for_existing(dm_row)
                stock = str(dm_row.get("Stock Number", unit)).strip()
                status_out = "Failed"
                detail = ""

                try:
                    # ── Try EDIT first ────────────────────────────────────────
                    try:
                        dvf.open_vehicle_form(page, stock)
                        dvf.fill_vehicle_page(page, dm_row_existing)
                        status_out = "Updated"
                        print(f"  ✓ Updated  {unit}")
                    except Exception as edit_err:
                        if dvf.is_recoverable_session_error(edit_err):
                            raise
                        err_str = str(edit_err)
                        if "Could not find exact vehicle row" in err_str or "not found" in err_str.lower():
                            # ── Before adding, check if this VIN already exists in DM ──
                            vin = str(dm_row.get("VIN", "")).strip().upper()
                            vin_match = dvf.find_inventory_match_by_vin(page, vin) if (vin and ALLOW_VIN_MATCH_EDIT) else None
                            if vin_match and ALLOW_VIN_MATCH_EDIT and dvf.inventory_match_has_stock(vin_match, stock):
                                # VIN already in DM under a different stock number — edit it
                                print(f"  ⚠ VIN {vin} already in DM (different stock#) — editing existing record")
                                try:
                                    vin_href = vin_match.get("href", "")
                                    if vin_href.startswith("/"):
                                        vin_href = f"https://dm.automanager.com{vin_href}"
                                    page.goto(vin_href, wait_until="domcontentloaded")
                                    dvf.wait_for_ready(page, timeout=15000)
                                    dvf.fill_vehicle_page(page, dm_row_existing)
                                    status_out = "Updated (vin match)"
                                    print(f"  ✓ Updated  {unit}  (vin match)")
                                except Exception as vin_err:
                                    if dvf.is_recoverable_session_error(vin_err):
                                        raise
                                    status_out = "Failed"
                                    detail = f"VIN-match edit failed: {vin_err}"
                                    print(f"  ✗ Failed   {unit}  {detail}")
                            elif vin_match and ALLOW_VIN_MATCH_EDIT:
                                status_out = "Failed"
                                detail = f"VIN {vin} exists under a different stock; expected exact stock {stock}"
                                print(f"  ✗ Failed   {unit}  {detail}")
                            else:
                                # ── ADD new vehicle ───────────────────────────────
                                try:
                                    dvf.open_new_vehicle_form(page, stock)
                                    dvf.prepare_add_vehicle_page(page, dm_row)
                                    dvf.fill_vehicle_page(page, dm_row)
                                    status_out = "Added"
                                    print(f"  + Added    {unit}")
                                except dvf.DuplicateVehicleBlocked:
                                    # DM detected a duplicate during add — reopen the existing record.
                                    # Prefer VIN here because the duplicate may already exist under a different stock.
                                    try:
                                        vin = str(dm_row.get("VIN", "")).strip().upper()
                                        vin_match = dvf.find_inventory_match_by_vin(page, vin) if vin else None
                                        if vin_match and dvf.inventory_match_has_stock(vin_match, stock):
                                            vin_href = vin_match.get("href", "")
                                            if vin_href.startswith("/"):
                                                vin_href = f"https://dm.automanager.com{vin_href}"
                                            page.goto(vin_href, wait_until="domcontentloaded")
                                            dvf.wait_for_ready(page, timeout=15000)
                                        elif vin_match:
                                            raise Exception(f"VIN {vin} exists under a different stock; expected exact stock {stock}")
                                        else:
                                            dvf.open_vehicle_form(page, stock)
                                        dvf.fill_vehicle_page(page, dm_row_existing)
                                        status_out = "Updated (dup→edit)"
                                        print(f"  ✓ Updated  {unit}  (dup→edit)")
                                    except Exception as dup_err:
                                        if dvf.is_recoverable_session_error(dup_err):
                                            raise
                                        status_out = "Failed"
                                        detail = f"Dup→edit failed: {dup_err}"
                                        print(f"  ✗ Failed   {unit}  {detail}")
                        else:
                            status_out = "Failed"
                            detail = err_str[:200]
                            print(f"  ✗ Failed   {unit}  {detail}")

                    session_errors = 0  # reset on success

                except Exception as outer_err:
                    if dvf.is_recoverable_session_error(outer_err) and session_errors < 2:
                        session_errors += 1
                        print(f"  ⚠ Session error for {unit}, restarting ({session_errors}/2)…")
                        try:
                            browser.close()
                        except Exception:
                            pass
                        browser, context, page = dvf.create_authenticated_session(p)
                        # Retry this row
                        try:
                            dvf.open_vehicle_form(page, stock)
                            dvf.fill_vehicle_page(page, dm_row)
                            status_out = "Updated (retry)"
                            print(f"  ✓ Updated  {unit}  (after session restart)")
                        except Exception as retry_err:
                            status_out = "Failed"
                            detail = f"Retry failed: {retry_err}"
                    else:
                        status_out = "Failed"
                        detail = str(outer_err)[:200]
                        print(f"  ✗ Failed   {unit}  {detail}")

                results.append({
                    "Unit":   unit,
                    "VIN":    str(row.get("VIN", "")),
                    "Year":   str(row.get("Year", "")),
                    "Make":   str(row.get("Make", "")),
                    "Model":  str(row.get("Model", "")),
                    "Status": status_out,
                    "Detail": detail,
                })

        finally:
            try:
                browser.close()
            except Exception:
                pass

    return pd.DataFrame(results)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    run_phase1 = PHASE_ONLY in ("", "1")
    run_phase2 = PHASE_ONLY in ("", "2")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    issues: List[dict] = []

    # ── Phase 1 ─────────────────────────────────────────────────────────────
    cleaned_df = None
    if run_phase1:
        cleaned_df = phase1_clean(issues)

        cleaned_csv = OUTPUT_DIR / "dm_import_master_cleaned.csv"
        cleaned_df.to_csv(cleaned_csv, index=False)
        print(f"\n[Phase 1] Cleaned CSV saved → {cleaned_csv}")
        print(f"  {len(cleaned_df)} rows")

        if issues:
            issues_df = pd.DataFrame(issues)
            issues_xlsx = OUTPUT_DIR / "dm_verification_issues.xlsx"
            with pd.ExcelWriter(issues_xlsx, engine="openpyxl") as writer:
                issues_df.to_excel(writer, sheet_name="Issues", index=False)

                # Summary pivot by field
                summary = issues_df.groupby("Field").size().reset_index(name="Count").sort_values("Count", ascending=False)
                summary.to_excel(writer, sheet_name="Summary by Field", index=False)

                # Unfilled critical fields
                critical = issues_df[issues_df["Fixed To"] == ""]
                if not critical.empty:
                    critical.to_excel(writer, sheet_name="Needs Manual Review", index=False)

            print(f"  {len(issues)} data changes/flags → {issues_xlsx}")

            # Console summary
            from collections import Counter
            field_counts = Counter(i["Field"] for i in issues)
            print("\n  Top issues by field:")
            for field, count in field_counts.most_common(15):
                print(f"    {field:35s} {count}")
        else:
            print("  No issues found.")
    else:
        # Phase 2 only — load the pre-cleaned CSV
        cleaned_csv = Path(CLEANED_CSV_OVERRIDE) if CLEANED_CSV_OVERRIDE else (OUTPUT_DIR / "dm_import_master_cleaned.csv")
        if not cleaned_csv.exists():
            print(f"ERROR: Cleaned CSV not found at {cleaned_csv}. Run Phase 1 first.")
            sys.exit(1)
        cleaned_df = pd.read_csv(cleaned_csv, dtype=str, keep_default_na=False)
        print(f"[Phase 2] Loaded cleaned CSV: {len(cleaned_df)} rows from {cleaned_csv}")

    # ── Phase 2 ─────────────────────────────────────────────────────────────
    if run_phase2 and (PHASE_ONLY == "2" or not DRY_RUN):
        if cleaned_df is None:
            print("ERROR: No cleaned data for Phase 2.")
            sys.exit(1)

        print(f"\n[Phase 2] Syncing {len(cleaned_df)} units to DeskManager…")
        if START_FROM:
            print(f"  Resuming from unit: {START_FROM}")

        report_df = phase2_sync(cleaned_df)

        report_xlsx = OUTPUT_DIR / "dm_sync_report.xlsx"
        with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
            report_df.to_excel(writer, sheet_name="Sync Report", index=False)

            summary = report_df.groupby("Status").size().reset_index(name="Count")
            summary.to_excel(writer, sheet_name="Summary", index=False)

        print(f"\n[Phase 2] Sync report → {report_xlsx}")
        counts = report_df["Status"].value_counts().to_dict()
        for k, v in sorted(counts.items()):
            print(f"  {k:30s} {v}")
    elif run_phase2 and DRY_RUN and PHASE_ONLY == "":
        print("\n[Phase 2] Skipped because DESKMANAGER_DRY_RUN=true (set DESKMANAGER_PHASE=2 to run a dry-run sync)")
    elif not run_phase2 and not DRY_RUN and PHASE_ONLY == "":
        print("\n[Phase 2] Skipped (set DESKMANAGER_DRY_RUN=false and ensure credentials are set to run Phase 2)")

    print("\nDone.")


if __name__ == "__main__":
    main()
