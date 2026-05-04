import os
import re
import sys
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE_URL = "https://dm.automanager.com/Account/Login?ReturnUrl=%2fDashboard"
INVENTORY_URL = "https://dm.automanager.com/Inventory?action=10"
DEFAULT_CSV_FILES = [
    "DM Vehicle master import.csv",
    "vehicle_updates.csv",
]

# Credentials must come from environment variables.
AUTOMANAGER_USERNAME = os.getenv("AUTOMANAGER_USERNAME", "").strip()
AUTOMANAGER_PASSWORD = os.getenv("AUTOMANAGER_PASSWORD", "").strip()
HEADLESS = os.getenv("PLAYWRIGHT_HEADLESS", "false").lower() in {"1", "true", "yes"}
SLOW_MO = int(os.getenv("PLAYWRIGHT_SLOW_MO", "250"))
VERBOSE_FORM_DUMP = os.getenv("DESKMANAGER_VERBOSE_FORM_DUMP", "false").lower() in {"1", "true", "yes"}
STRICT_FILL = os.getenv("DESKMANAGER_STRICT_FILL", "true").lower() in {"1", "true", "yes"}
DRY_RUN = os.getenv("DESKMANAGER_DRY_RUN", "false").lower() in {"1", "true", "yes"}
ONLY_STOCKS_RAW = os.getenv("DESKMANAGER_ONLY_STOCKS", "").strip()
FILL_REQUIRED_PLACEHOLDERS = os.getenv("DESKMANAGER_FILL_REQUIRED_PLACEHOLDERS", "false").lower() in {"1", "true", "yes"}
RUN_MODE = os.getenv("DESKMANAGER_RUN_MODE", "edit").strip().lower()
MANUAL_LOGIN = os.getenv("DESKMANAGER_MANUAL_LOGIN", "false").lower() in {"1", "true", "yes"}
MANUAL_LOGIN_TIMEOUT_MS = int(os.getenv("DESKMANAGER_MANUAL_LOGIN_TIMEOUT_MS", "600000"))
DEFAULT_EXCEL_UPDATE_FILE = "all_updated_units_wash_added.xlsx"
UPDATE_SOURCE_FILE = os.getenv("DESKMANAGER_UPDATE_SOURCE_FILE", DEFAULT_EXCEL_UPDATE_FILE).strip()
UPDATE_LOG_FILE = os.getenv("DESKMANAGER_UPDATE_LOG_FILE", "DeskManager_Update_Log.csv").strip()

DUPLICATE_CONFIRMATION_KEYWORDS = (
    "duplicate",
    "already exists",
    "already in inventory",
    "already on file",
    "already have",
)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(line_buffering=True)
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(line_buffering=True)


def parse_only_stocks(raw_value):
    if not raw_value:
        return []
    parts = [p.strip() for p in raw_value.replace(";", ",").split(",")]
    return [p for p in parts if p]


class DuplicateVehicleBlocked(Exception):
    pass


def duplicate_confirmation_text(text):
    normalized = " ".join(str(text or "").lower().split())
    if not normalized:
        return ""

    if any(token in normalized for token in DUPLICATE_CONFIRMATION_KEYWORDS):
        return normalized
    if "continue" in normalized and any(token in normalized for token in ["duplicate", "already", "exists"]):
        return normalized
    return ""


def clear_duplicate_prompt(page):
    setattr(page, "_deskmanager_duplicate_prompt", "")


def record_duplicate_prompt(page, message):
    normalized = duplicate_confirmation_text(message)
    if normalized:
        setattr(page, "_deskmanager_duplicate_prompt", normalized)
    return normalized


def get_duplicate_prompt(page):
    return getattr(page, "_deskmanager_duplicate_prompt", "")


def dump_form_elements(page):
    """Debug function to show all form elements on the page"""
    print("\n" + "="*60)
    print("DUMPING ALL FORM ELEMENTS ON PAGE")
    print("="*60)


def dump_note_controls(page):
    """Debug helper for note-related controls when add-note action is not found."""
    print("\n" + "="*60)
    print("DUMPING NOTE CONTROLS")
    print("="*60)

    noteish = page.locator(
        "a:visible, button:visible, [role='button']:visible, i:visible, span:visible"
    )
    shown = 0
    for i in range(min(noteish.count(), 300)):
        if shown >= 80:
            break
        item = noteish.nth(i)
        try:
            text = (item.inner_text() or "").strip()
        except Exception:
            text = ""
        try:
            title = item.get_attribute("title") or ""
            aria = item.get_attribute("aria-label") or ""
            cls = item.get_attribute("class") or ""
            onclick = item.get_attribute("onclick") or ""
            href = item.get_attribute("href") or ""
        except Exception:
            continue

        blob = " ".join([text, title, aria, cls, onclick, href]).lower()
        if not any(tok in blob for tok in ["note", "sticky", "task", "plus", "add", "new", "fa-plus", "icon-plus"]):
            continue

        shown += 1
        print(
            f"  note-control {shown}: text='{text[:80]}' title='{title[:60]}' aria='{aria[:60]}' class='{cls[:80]}' onclick='{onclick[:80]}' href='{href[:80]}'"
        )

    textareas = page.locator("textarea:visible")
    print(f"Visible textareas: {textareas.count()}")
    print("="*60)

    # Check for modals
    modals = page.locator("div.mantine-Modal-root, div[role='dialog'], div.modal")
    print(f"\nMODALS FOUND: {modals.count()}")
    for i in range(modals.count()):
        modal = modals.nth(i)
        visible = modal.is_visible()
        print(f"Modal {i}: visible={visible}")
        if visible:
            modal_inputs = modal.locator("input, select, textarea")
            modal_labels = modal.locator("label")
            print(f"  - Form elements in modal: {modal_inputs.count()}")
            print(f"  - Labels in modal: {modal_labels.count()}")

    # Check main page form elements
    print(f"\nMAIN PAGE FORM ELEMENTS:")
    all_inputs = page.locator("input:visible, select:visible, textarea:visible")
    all_labels = page.locator("label")

    print(f"Visible inputs/selects: {all_inputs.count()}")
    print(f"Labels: {all_labels.count()}")

    # Show all form elements
    for i in range(min(50, all_inputs.count())):
        inp = all_inputs.nth(i)
        try:
            tag = inp.evaluate("el => el.tagName.toLowerCase()")
            name = inp.get_attribute("name") or inp.get_attribute("id") or "?"
            placeholder = inp.get_attribute("placeholder") or ""
            value = inp.get_attribute("value") or ""
            print(f"  {i}: {tag} name='{name}' placeholder='{placeholder}' value='{value}'")
        except Exception as e:
            print(f"  {i}: ERROR - {e}")

    # Show labels
    for i in range(min(30, all_labels.count())):
        lbl = all_labels.nth(i)
        try:
            text = lbl.inner_text().strip()
            if text:
                print(f"  Label {i}: '{text}'")
        except Exception as e:
            print(f"  Label {i}: ERROR - {e}")

    # Check for any buttons
    buttons = page.locator("button")
    print(f"\nBUTTONS: {buttons.count()}")
    for i in range(min(15, buttons.count())):
        btn = buttons.nth(i)
        try:
            text = btn.inner_text().strip()
            visible = btn.is_visible()
            print(f"  Button {i}: '{text}' visible={visible}")
        except Exception as e:
            print(f"  Button {i}: ERROR - {e}")

    print("="*60)


def clean(v):
    if isinstance(v, pd.Series):
        # Duplicate CSV column names can produce a Series for a single key.
        for item in v.tolist():
            candidate = clean(item)
            if candidate:
                return candidate
        return ""
    if isinstance(v, (list, tuple, set)):
        for item in v:
            candidate = clean(item)
            if candidate:
                return candidate
        return ""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def normalize_column_key(value):
    return "".join(ch for ch in str(value).strip().lower() if ch.isalnum())


def normalize_csv_columns(columns):
    alias_map = {
        "unit": "Stock Number",
        "stocknumber": "Stock Number",
        "stockno": "Stock Number",
        "stock no": "Stock Number",
        "stock #": "Stock Number",
        "stock#": "Stock Number",
        "stock": "Stock Number",
        "location": "Location",
        "datein": "Date In",
        "substatus": "Sub Status",
        "vehicletype": "Vehicle Type",
        "newused": "New / Used",
        "bodystyle": "Body Style",
        "drivetrain": "Drivetrain",
        "colorexterior": "Exterior Color",
        "colorinterior": "Interior Color",
        "grossweight": "Gross Weight",
        "licenseplate": "License Plate",
        "stateregistered": "State Registered",
        "state": "Title DMV State",
        "keynumber": "Key Number",
        "billofsaledate": "Bill Of Sale Date",
        "billofsalenumber": "Bill Of Sale Number",
        "mileagecurrent": "Mileage(Current)",
        "titlein": "Title-In",
        "passtimeserialno": "Pass Time Serial No.",
        "passtimeserialnumber": "Pass Time Serial No.",
        "purchasecost": "Purchase Cost",
        "purchasemethod": "Purchase Method",
        "purchasemethod1": "Purchase Method",
        "purchasedate": "Purchase Date",
        "purchasechannel": "Purchase Channel",
        "purchasedfrom": "Purchased From",
        "purchasefrom": "Purchased From",
        "paymentmethod": "Payment Method",
        "referenceno": "Reference No.",
        "referencenumber": "Reference No.",
        "inventorydate": "Date In",
        "invoiceno": "Invoice No.",
        "mileagein": "Mileage In",
        "draftduedate": "Draft Due Date",
        "packamount": "Pack Amount",
        "odometerstatus": "Odometer Status",
        "draftpaiddate": "Draft Paid Date",
        "draftpaidto": "Draft Paid To",
        "purchasenotes": "Notes",
        "rostitlenumber": "ROS / Title Number",
        "titlestatus": "Title Status",
        "titledmvstate": "Title DMV State",
        "titleindate": "Title In",
        "titleoutdate": "Title Out",
        "totaltaxpaid": "Total Tax Paid",
        "previoustitleowner": "Previous Title Owner",
        "titledmvnote": "Note",
        "tempplatenumber1": "Temp Plate Number 1",
        "dateissued1": "Date Issued 1",
        "expirationdate1": "Expiration Date 1",
        "tempplatenumber2": "Temp Plate Number 2",
        "dateissued2": "Date Issued 2",
        "expirationdate2": "Expiration Date 2",
        "transfervin": "Transfer VIN",
        "transferplate": "Transfer Plate",
        "transferyear": "Transfer Year",
        "transfermake": "Transfer Make",
        "transferdate": "Transfer Date",
        "inspectiondate": "Inspection Date",
        "inspectorsname": "Inspector's Name",
        "inspectorsid": "Inspector's ID",
        "plateexpirationdate": "Plate Expiration Date",
    }

    normalized = []
    for col in columns:
        key = str(col).strip()
        normalized.append(alias_map.get(normalize_column_key(key), key))
    return normalized


def resolve_csv_file():
    env_path = os.getenv("DESKMANAGER_CSV_FILE", "").strip()
    candidates = [env_path] if env_path else DEFAULT_CSV_FILES

    for candidate in candidates:
        path = Path(candidate).expanduser()
        if not path.is_absolute():
            path = Path(__file__).with_name(candidate)
        if path.exists():
            return path

    if env_path:
        return Path(env_path).expanduser()
    return Path(__file__).with_name(DEFAULT_CSV_FILES[0])


def resolve_excel_update_file():
    candidate = UPDATE_SOURCE_FILE or DEFAULT_EXCEL_UPDATE_FILE
    path = Path(candidate).expanduser()
    if not path.is_absolute():
        path = Path(__file__).with_name(candidate)
    return path


def resolve_update_log_file():
    candidate = UPDATE_LOG_FILE or "DeskManager_Update_Log.csv"
    path = Path(candidate).expanduser()
    if not path.is_absolute():
        path = Path(__file__).with_name(candidate)
    return path


def load_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        first = f.readline()
        skiprows = 1 if first and "," not in first and "\t" not in first else 0

    df = pd.read_csv(path, skiprows=skiprows, dtype=str)
    df.columns = normalize_csv_columns(df.columns)
    return df

FIELD_LABEL_ALIASES = {
    "Stock Number": ["Stock Number", "Stock #", "Stock#", "Stock No", "Stock"],
    "Status": ["Status", "Vehicle Status", "Status Id", "Status"] ,
    "Sub Status": ["Sub Status", "Substatus", "Sub Status"],
    "Location": ["Location", "Select Lot"],
    "Date In": ["Date In", "Inventory Date"],
    "City": ["City", "Location City"],
    "Bill Of Sale Date": ["Bill Of Sale Date", "Bill of Sale", "Bill of Sale Dt", "Sale Date"],
    "Bill Of Sale Number": ["Bill Of Sale Number", "Bill of sale Number", "Bill Of Sale #"],
    "VIN": ["VIN", "Vehicle Identification Number"],
    "Vehicle Type": ["Vehicle Type", "Type"],
    "New / Used": ["New / Used", "New / Used", "Condition"],
    "Condition": ["Condition", "Vehicle Condition"],
    "Trailer Type": ["Trailer Type", "Trailer"],
    "Year": ["Year", "year", "Vehicle Year"],
    "Make": ["Make", "Manufacturer"],
    "Model": ["Model"],
    "Series": ["Series"],
    "Body Style": ["Body Style", "Bodystyle"],
    "Engine": ["Engine"],
    "Drivetrain": ["Drivetrain", "Drive Train"],
    "Overdrive": ["Overdrive"],
    "Length": ["Length"],
    "Exterior Color": ["Exterior Color", "Exterior Colour"],
    "Interior Color": ["Interior Color", "Interior Colour"],
    "Weight": ["Weight", "Gross Weight"],
    "Height": ["Height"],
    "Axels": ["Axels", "Axle", "Axles"],
    "Suspension": ["Suspension"],
    "Tire size": ["Tire size", "Tire Size", "Tire"],
    "Gross Weight": ["Gross Weight"],
    "Unladen Weight": ["Unladen Weight"],
    "License Plate": ["License Plate", "Plate"],
    "Last Registered": ["Last Registered", "Registration Date"],
    "State Registered": ["State Registered", "State"],
    "Bill of sale Date": ["Bill of sale Date", "Bill of sale", "Bill of Sale Date"],
    "Carrying Capacity": ["Carrying Capacity"],
    "Pass Time Serial No.": ["Pass Time Serial No.", "Pass Time Serial", "Pass Time"],
    "Key Number": ["Key Number"],
    "Title-In": ["Title-In", "Title In"],
    "Tagline": ["Tagline"],
    "Description": ["Description"],
    "Asking Price": ["Asking Price"],
    "Internet Price": ["Internet Price"],
    "MSRP": ["MSRP"],
    "Mileage(Current)": ["Mileage(Current)", "Mileage (Current)", "Mileage"],
    "Retail": ["Retail"],
    "Trade": ["Trade"],
    "Wholesale": ["Wholesale"],
    "Purchase Cost": ["Purchase Cost"],
    "Purchase Method": ["Purchase Method"],
    "Purchase Date": ["Purchase Date"],
    "Purchase Channel": ["Purchase Channel"],
    "Purchased From": ["Purchased From"],
    "Payment Method": ["Payment Method"],
    "Due Date": ["Due Date"],
    "Reference No.": ["Reference No.", "Reference No"],
    "Invoice No.": ["Invoice No.", "Invoice No"],
    "Buyer": ["Buyer"],
    "Mileage In": ["Mileage In"],
    "Draft Due Date": ["Draft Due Date"],
    "Pack Amount": ["Pack Amount"],
    "Odometer Status": ["Odometer Status"],
    "Draft Paid Date": ["Draft Paid Date"],
    "Draft Paid To": ["Draft Paid To"],
    "Notes": ["Notes", "Note"],
    "ROS / Title Number": ["ROS / Title Number", "ROS / Title", "Title Number"],
    "Title Status": ["Title Status"],
    "State": ["State"],
    "Title In": ["Title In"],
    "Title Out": ["Title Out"],
    "Total Tax Paid": ["Total Tax Paid"],
    "Previous Title Owner": ["Previous Title Owner"],
    "Note": ["Note", "Notes"],
    "Temp Plate Number 1": ["Temp Plate Number 1"],
    "Date Issued 1": ["Date Issued 1"],
    "Expiration Date 1": ["Expiration Date 1"],
    "Temp Plate Number 2": ["Temp Plate Number 2"],
    "Date Issued 2": ["Date Issued 2"],
    "Expiration Date 2": ["Expiration Date 2"],
    "Transfer Plate": ["Transfer Plate"],
    "Transfer Date": ["Transfer Date"],
    "Transfer VIN": ["VIN", "Transfer VIN"],
    "Transfer Year": ["Year", "Transfer Year"],
    "Transfer Make": ["Make", "Transfer Make"],
    "Inspection Date": ["Inspection Date"],
    "Inspector's Name": ["Inspector's Name"],
    "Inspector's ID": ["Inspector's ID"],
    "Plate Expiration Date": ["Plate Expiration Date"],
    "Sold To": ["Sold To"],
    "Sold Date": ["Sold Date"],
    "Sold Price": ["Sold Price"],
}

FIELD_SELECTORS = {
    "Stock Number": [
        "input[name='Vehicle.StockNumber']",
        "input[id='Vehicle_StockNumber']",
    ],
    "Status": [
        "select[name='Vehicle.StatusId']",
        "select[id='Vehicle_StatusId']",
        "select[name='StatusId']",
    ],
    "Sub Status": [
        "select[name='Vehicle.SubstatusId']",
        "select[id='Vehicle_SubstatusId']",
        "select[name='SubstatusId']",
    ],
    "Location": [
        "select[name='Vehicle.Location']",
        "select[name*='Location']",
        "select[id*='Location']",
        "select[name*='Lot']",
        "select[id*='Lot']",
        "input[name*='Lot']",
        "input[id*='Lot']",
    ],
    "Date In": [
        "input[name='Vehicle.InventoryDate']",
        "input[name*='InventoryDate']",
        "input[id*='InventoryDate']",
    ],
    "City": [
        "select[name*='City']",
        "select[id*='City']",
        "input[name*='City']",
        "input[id*='City']",
    ],
    "Title-In": [
        "input[type='checkbox'][name*='TitleIn']",
        "input[type='checkbox'][id*='TitleIn']",
    ],
    "VIN": [
        "input[name='Vehicle.Vin']",
        "input[id='Vehicle_Vin']",
        "input[name='Vin']",
    ],
    "Vehicle Type": [
        "select[name='Vehicle.CategoryId']",
        "select[id='Vehicle_CategoryId']",
        "select[name='CategoryId']",
    ],
    "New / Used": [
        "select[name='Vehicle.NewUsed']",
        "select[id='Vehicle_NewUsed']",
    ],
    "Condition": [
        "select[name='Vehicle.ConditionId']",
        "select[id='Vehicle_ConditionId']",
    ],
    "Trailer Type": [
        "select[name='VehicleTypeData.TrailerType']",
        "select[id*='TrailerType']",
    ],
    "Year": [
        "input[name='Vehicle.Year']",
        "input[id='Vehicle_Year']",
    ],
    "Make": [
        "input[name='Vehicle.Make']",
        "input[id='Vehicle_Make']",
    ],
    "Model": [
        "input[name='Vehicle.Model']",
        "input[id='Vehicle_Model']",
    ],
    "Series": [
        "input[name='Vehicle.Series']",
        "input[id='Vehicle_Series']",
    ],
    "Body Style": [
        "input[name='Vehicle.BodyStyle']",
        "input[id='Vehicle_BodyStyle']",
    ],
    "Engine": [
        "input[name='Vehicle.EngineSize']",
        "input[id='Vehicle_EngineSize']",
    ],
    "Drivetrain": [
        "input[name='Vehicle_DriveTrain-selectized']",
        "input[id*='DriveTrain'][id*='selectized']",
    ],
    "Overdrive": [
        "select[name='Vehicle.Overdrive']",
        "select[id='Vehicle_Overdrive']",
    ],
    "Length": [
        "input[name='Vehicle.Length']",
        "input[id='Vehicle_Length']",
    ],
    "Exterior Color": [
        "input[name='Vehicle_ExteriorColor-selectized']",
        "input[id*='ExteriorColor'][id*='selectized']",
    ],
    "Interior Color": [
        "input[name='Vehicle_InteriorColor-selectized']",
        "input[id*='InteriorColor'][id*='selectized']",
    ],
    "Weight": [
        "input[name='VehicleTypeData.Weight']",
        "input[id*='Weight']",
    ],
    "Height": [
        "input[name='VehicleTypeData.Height']",
        "input[id*='Height']",
    ],
    "Axels": [
        "select[name='VehicleTypeData.TruckAxles']",
        "select[id*='TruckAxles']",
    ],
    "Suspension": [
        "select[name='VehicleTypeData.TruckSuspension']",
        "select[id*='TruckSuspension']",
    ],
    "Tire size": [
        "select[name='VehicleTypeData.TruckTireSize']",
        "select[id*='TruckTireSize']",
    ],
    "Gross Weight": [
        "input[name='Vehicle.GrossWeight']",
        "input[id='Vehicle_GrossWeight']",
    ],
    "Unladen Weight": [
        "input[name='Vehicle.UnladenWeight']",
        "input[id='Vehicle_UnladenWeight']",
    ],
    "License Plate": [
        "input[name='Vehicle.LicensePlate']",
        "input[id='Vehicle_LicensePlate']",
    ],
    "Last Registered": [
        "input[name='Vehicle.LastRegistered']",
        "input[id='Vehicle_LastRegistered']",
    ],
    "State Registered": [
        "input[name='Vehicle.StateRegistered']",
        "input[id='Vehicle_StateRegistered']",
    ],
    "Key Number": [
        "input[name='Vehicle.KeyNumber']",
        "input[id='Vehicle_KeyNumber']",
    ],
    "Carrying Capacity": [
        "input[name='Vehicle.CarryCapacity']",
        "input[id*='CarryCapacity']",
    ],
    "Pass Time Serial No.": [
        "input[name='VehicleDataExtension.PassTimeSerialNumber']",
        "input[id*='PassTimeSerial']",
    ],
    "Purchase Cost": [
        "input[name='VehicleAcquisition.PurchasePrice']",
        "input[name*='PurchaseCost']",
        "input[id*='PurchaseCost']",
        "input[name*='Cost']",
        "input[id*='Cost']",
    ],
    "Purchase Date": [
        "input[name='VehicleAcquisition.PurchaseDate']",
        "input[name*='PurchaseDate']",
        "input[id*='PurchaseDate']",
    ],
    "Purchased From": [
        "input[name='VehicleAcquisition.PurchaseVendor']",
        "input[name*='PurchasedFrom']",
        "input[id*='PurchasedFrom']",
    ],
    "Purchase Method": [
        "select[name='VehicleAcquisition.PurchaseMethodId']",
        "select[name*='PurchaseMethod']",
        "select[id*='PurchaseMethod']",
    ],
    "Purchase Channel": [
        "select[name='VehicleAcquisition.PurchaseChannelId']",
        "select[name*='PurchaseChannel']",
        "select[id*='PurchaseChannel']",
        "input[name*='PurchaseChannel']",
        "input[id*='PurchaseChannel']",
    ],
    "Payment Method": [
        "select[name='VehicleAcquisition.PurchasePaymentMethodId']",
        "select[name*='PaymentMethod']",
        "select[id*='PaymentMethod']",
        "input[name*='PaymentMethod']",
        "input[id*='PaymentMethod']",
    ],
    "Due Date": [
        "input[name='VehicleAcquisition.PurchaseDueDate']",
        "input[name*='DueDate']",
        "input[id*='DueDate']",
    ],
    "Reference No.": [
        "input[name='VehicleAcquisition.PurchaseRefNumber']",
        "input[name*='Reference']",
        "input[id*='Reference']",
    ],
    "Invoice No.": [
        "input[name='VehicleAcquisition.PurchaseInvoiceNumber']",
        "input[name*='Invoice']",
        "input[id*='Invoice']",
    ],
    "Buyer": [
        "input[name='VehicleAcquisition.PurchasedBy']",
        "input[name*='Buyer']",
        "input[id*='Buyer']",
    ],
    "Mileage In": [
        "input[name='VehicleAcquisition.PurchaseMileage']",
        "input[name*='MileageIn']",
        "input[id*='MileageIn']",
    ],
    "Draft Due Date": [
        "input[name*='DraftDueDate']",
        "input[id*='DraftDueDate']",
    ],
    "Pack Amount": [
        "input[name*='PackAmount']",
        "input[id*='PackAmount']",
    ],
    "Odometer Status": [
        "select[name='VehicleAcquisition.PurchaseOdometerStatus']",
        "select[name*='OdometerStatus']",
        "select[id*='OdometerStatus']",
        "input[name*='OdometerStatus']",
        "input[id*='OdometerStatus']",
    ],
    "Draft Paid Date": [
        "input[name*='DraftPaidDate']",
        "input[id*='DraftPaidDate']",
    ],
    "Draft Paid To": [
        "input[name*='DraftPaidTo']",
        "input[id*='DraftPaidTo']",
    ],
    "Notes": [
        "textarea[name='VehicleAcquisition.Notes']",
        "textarea[name*='Notes']",
        "textarea[id*='Notes']",
        "input[name*='Notes']",
        "input[id*='Notes']",
    ],
    "Sold To": [
        "input[name*='SoldTo']",
        "input[id*='SoldTo']",
    ],
    "Sold Date": [
        "input[name*='SoldDate']",
        "input[id*='SoldDate']",
    ],
    "Sold Price": [
        "input[name*='SoldPrice']",
        "input[id*='SoldPrice']",
    ],
}

TAB_FIELD_CONFIG = [
    (
        "Details",
        [
            ("VIN", "VIN"),
            ("Vehicle Type", "Vehicle Type"),
            ("New / Used", "New / Used"),
            ("Condition", "Condition"),
            ("Trailer Type", "Trailer Type"),
            ("Year", "Year"),
            ("Make", "Make"),
            ("Model", "Model"),
            ("Series", "Series"),
            ("Body Style", "Body Style"),
            ("Engine", "Engine"),
            ("Drivetrain", "Drivetrain"),
            ("Overdrive", "Overdrive"),
            ("Length", "Length"),
            ("Exterior Color", "Exterior Color"),
            ("Interior Color", "Interior Color"),
            ("Weight", "Weight"),
            ("Height", "Height"),
            ("Axels", "Axles"),
            ("Suspension", "Suspension"),
            ("Tire size", "Tire Size"),
            ("Gross Weight", "Gross Weight"),
            ("Unladen Weight", "Unladen Weight"),
            ("License Plate", "License Plate"),
            ("Last Registered", "Last Registered"),
            ("State Registered", "State Registered"),
            ("Key Number", "Key Number"),
            ("Carrying Capacity", "Carrying Capacity"),
            ("Pass Time Serial No.", "Pass Time Serial No."),
        ],
    ),
    (
        "Values",
        [
            ("Asking Price", "Asking Price"),
            ("Internet Price", "Internet Price"),
            ("MSRP", "MSRP"),
            ("Mileage(Current)", "Mileage(Current)"),
            ("Retail", "Retail"),
            ("Trade", "Trade"),
            ("Wholesale", "Wholesale"),
        ],
    ),
    (
        "Other",
        [
            ("Title-In", "Title-In"),
            ("Bill Of Sale Date", "Bill Of Sale Date"),
            ("Bill Of Sale Number", "Bill Of Sale Number"),
            ("City", "City"),
            ("Sold To", "Sold To"),
            ("Sold Date", "Sold Date"),
            ("Sold Price", "Sold Price"),
        ],
    ),
    (
        "Description",
        [
            ("Tagline", "Tagline"),
            ("Description", "Description"),
        ],
    ),
    (
        "Title-DMV",
        [
            ("ROS / Title Number", "ROS / Title Number"),
            ("Title Status", "Title Status"),
            ("State", "Title DMV State"),
            ("Title In", "Title In"),
            ("Title Out", "Title Out"),
            ("Total Tax Paid", "Total Tax Paid"),
            ("Previous Title Owner", "Previous Title Owner"),
            ("Note", "Note"),
            ("Temp Plate Number 1", "Temp Plate Number 1"),
            ("Date Issued 1", "Date Issued 1"),
            ("Expiration Date 1", "Expiration Date 1"),
            ("Temp Plate Number 2", "Temp Plate Number 2"),
            ("Date Issued 2", "Date Issued 2"),
            ("Expiration Date 2", "Expiration Date 2"),
            ("Transfer VIN", "Transfer VIN"),
            ("Transfer Year", "Transfer Year"),
            ("Transfer Plate", "Transfer Plate"),
            ("Transfer Make", "Transfer Make"),
            ("Transfer Date", "Transfer Date"),
            ("Inspection Date", "Inspection Date"),
            ("Inspector's Name", "Inspector's Name"),
            ("Inspector's ID", "Inspector's ID"),
            ("Plate Expiration Date", "Plate Expiration Date"),
        ],
    ),
    (
        "Purchase Info",
        [
            ("Purchase Cost", "Purchase Cost"),
            ("Purchase Date", "Purchase Date"),
            ("Purchased From", "Purchased From"),
            ("Purchase Method", "Purchase Method"),
            ("Purchase Channel", "Purchase Channel"),
            ("Payment Method", "Payment Method"),
            ("Due Date", "Due Date"),
            ("Reference No.", "Reference No."),
            ("Invoice No.", "Invoice No."),
            ("Buyer", "Buyer"),
            ("Mileage In", "Mileage In"),
            ("Draft Due Date", "Draft Due Date"),
            ("Pack Amount", "Pack Amount"),
            ("Odometer Status", "Odometer Status"),
            ("Draft Paid Date", "Draft Paid Date"),
            ("Draft Paid To", "Draft Paid To"),
            ("Notes", "Notes"),
        ],
    ),
]

TOP_LEVEL_FIELD_CONFIG = [
    ("Stock Number", "Stock Number"),
    ("Status", "Status"),
    ("Sub Status", "Sub Status"),
    ("Location", "Location"),
    ("Date In", "Date In"),
]


def row_value(row, key):
    if key in row.index:
        return clean(row[key])
    return ""


def is_truthy(value):
    return str(value).strip().lower() in {"1", "true", "yes", "y", "checked", "on"}


NUMERIC_ONLY_FIELDS = {
    "Asking Price",
    "Internet Price",
    "MSRP",
    "Retail",
    "Trade",
    "Wholesale",
    "Purchase Cost",
    "Sold Price",
    "Pack Amount",
    "Total Tax Paid",
    "Mileage(Current)",
    "Mileage In",
}


LOCATION_VALUE_ALIASES = {
    "big resv": "Big Reservoir",
    "big resv.": "Big Reservoir",
    "big reservoir": "Big Reservoir",
}


PURCHASE_KEYWORDS = {
    "purchase",
    "payment",
    "due",
    "reference",
    "invoice",
    "buyer",
    "mileage",
    "draft",
    "pack",
    "odometer",
    "note",
    "cost",
    "sold",
}


IGNORED_ROW_COLUMNS = {
    "Stock Number",
    "Unit",
    "Photo 1",
    "Photo 2",
    "Photo 3",
    "Attachment 1",
    "Attachment 2",
}

# Optional placeholder strategy for required fields when CSV value is missing.
# Default is off because blank CSV fields should stay blank unless explicitly opted in.
PLACEHOLDER_TEXT_VALUE = "TBD"
PLACEHOLDER_DATE_VALUE = "1/1/2026"
NUMERIC_PLACEHOLDER_CANDIDATES = ["0.01", "0.1", "001", "1"]

# field label -> metadata about when/how to backfill placeholders
REQUIRED_FIELD_RULES = {
    "Purchase Cost": {"csv_key": "Purchase Cost", "kind": "numeric"},
    "Purchase Date": {"csv_key": "Purchase Date", "kind": "date"},
    "Purchased From": {"csv_key": "Purchased From", "kind": "text"},
}

def normalize_numeric_value(raw_value):
    text = str(raw_value).strip()
    if not text:
        return text

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    cleaned = text.replace("$", "").replace(",", "").replace(" ", "")
    filtered = "".join(ch for ch in cleaned if ch.isdigit() or ch in {".", "-"})
    if not filtered:
        return ""

    if negative and not filtered.startswith("-"):
        filtered = f"-{filtered}"
    return filtered


def normalize_field_value(label, value):
    value = clean(value)
    if not value:
        return ""
    if label == "Location":
        normalized_location = LOCATION_VALUE_ALIASES.get(value.strip().lower())
        if normalized_location and normalized_location != value:
            print(f"  DEBUG: Normalized location from '{value}' to '{normalized_location}'")
            return normalized_location
    if label in NUMERIC_ONLY_FIELDS:
        normalized = normalize_numeric_value(value)
        if normalized != value:
            print(f"  DEBUG: Normalized numeric field '{label}' from '{value}' to '{normalized}'")
        return normalized
    return value


def normalize_match_key(value):
    return "".join(ch for ch in str(value).lower() if ch.isalnum())


def is_purchase_related_column(column_name):
    key = normalize_column_key(column_name)
    return any(token in key for token in PURCHASE_KEYWORDS)


def tab_has_values(row, field_specs):
    return any(bool(row_value(row, key)) for _, key in field_specs)


def click_tab(page, tab_name):
    print(f"  Switching to tab '{tab_name}'...")
    accept_inventory_date_confirmation(page)
    close_active_modals(page)
    selectors = [
        f"a:has-text('{tab_name}')",
        f"button:has-text('{tab_name}')",
        f"li:has-text('{tab_name}')",
    ]
    if not first_visible_click(page, selectors, timeout=4000):
        try:
            page.get_by_text(tab_name, exact=True).click(timeout=4000)
        except Exception as exc:
            close_active_modals(page)
            accept_inventory_date_confirmation(page)
            raise Exception(f"Could not open tab '{tab_name}': {exc}")
    page.wait_for_timeout(800)


def fill_tab_fields(page, row, tab_name, field_specs):
    if not tab_has_values(row, field_specs):
        print(f"  Skipping tab '{tab_name}' (no CSV values for mapped fields)")
        return 0, 0

    click_tab(page, tab_name)

    attempted = 0
    filled = 0
    for label, key in field_specs:
        value = row_value(row, key)
        if not value:
            continue
        attempted += 1
        if fill_field(page, label, value):
            filled += 1

    print(f"  Tab '{tab_name}' summary: filled {filled}/{attempted} fields")
    return attempted, filled


def fill_purchase_tab_extras(page, row, already_mapped_keys):
    attempted = 0
    filled = 0
    for column_name in row.index:
        if column_name in already_mapped_keys:
            continue
        if column_name in IGNORED_ROW_COLUMNS:
            continue
        if not is_purchase_related_column(column_name):
            continue

        value = row_value(row, column_name)
        if not value:
            continue

        attempted += 1
        if fill_field(page, column_name, value):
            filled += 1

    if attempted:
        print(f"  Purchase tab extra pass: filled {filled}/{attempted} additional fields")
    return attempted, filled


def collect_unmapped_nonempty_values(row, mapped_keys):
    pending = {}
    for column_name in row.index:
        if column_name in mapped_keys:
            continue
        if column_name in IGNORED_ROW_COLUMNS:
            continue

        value = row_value(row, column_name)
        if not value:
            continue

        # Keep first non-empty instance for duplicated column names.
        if column_name not in pending:
            pending[column_name] = value
    return pending


def fill_remaining_columns_across_tabs(page, pending_values):
    if not pending_values:
        return 0, 0

    attempted = 0
    filled = 0
    print("  Running fallback pass for all remaining non-empty headers...")

    for tab_name, _ in TAB_FIELD_CONFIG:
        if not pending_values:
            break

        click_tab(page, tab_name)
        for column_name in list(pending_values.keys()):
            attempted += 1
            if fill_field(page, column_name, pending_values[column_name]):
                filled += 1
                del pending_values[column_name]

    if pending_values:
        preview = ", ".join(list(pending_values.keys())[:10])
        if len(pending_values) > 10:
            preview += f" ... (+{len(pending_values) - 10} more)"
        print(f"  WARNING: Could not map {len(pending_values)} non-empty header(s): {preview}")

    print(f"  Fallback header pass summary: filled {filled}/{attempted} attempts")
    return attempted, filled


def close_active_modals(page):
    try:
        dialogs = page.locator("div.bootstrap-dialog.in, div.modal.in, div[role='dialog'], div.ui-dialog, div.ui-dialog-content")
        for i in range(min(dialogs.count(), 5)):
            dialog = dialogs.nth(i)
            if not dialog.is_visible():
                continue

            try:
                dialog_text = dialog.inner_text()
            except Exception:
                dialog_text = ""

            if record_duplicate_prompt(page, dialog_text):
                cancel_btn = dialog.locator(
                    "button:has-text('No'), button:has-text('Cancel'), "
                    "button:has-text('Close'), button[data-bb-handler='cancel'], [data-dismiss='modal']"
                )
                if cancel_btn.count() > 0 and cancel_btn.first.is_visible() and cancel_btn.first.is_enabled():
                    cancel_btn.first.click(timeout=2000)
                    page.wait_for_timeout(300)
                    print("  DEBUG: Dismissed duplicate confirmation dialog")
                else:
                    try:
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(300)
                        print("  DEBUG: Closed duplicate confirmation dialog with Escape")
                    except Exception:
                        pass
                continue

            try:
                close_btn = dialog.locator(
                    "button.close, button:has-text('Close'), button:has-text('Cancel'), [data-dismiss='modal'], "
                    "button.ui-dialog-titlebar-close, .ui-dialog-titlebar-close, a.ui-dialog-titlebar-close"
                )
                if close_btn.count() > 0 and close_btn.first.is_visible() and close_btn.first.is_enabled():
                    close_btn.first.click(timeout=2000)
                    page.wait_for_timeout(200)
            except Exception:
                pass

        # If a jQuery UI overlay remains, Escape usually dismisses the top dialog.
        try:
            overlay = page.locator("div.ui-widget-overlay.ui-front:visible")
            for _ in range(min(overlay.count(), 3)):
                page.keyboard.press("Escape")
                page.wait_for_timeout(150)
        except Exception:
            pass
    except Exception:
        pass


def close_sticky_note_popup_if_present(page):
    """Close sticky-note popup shown on vehicle open so automation can proceed."""
    try:
        dialogs = page.locator("div[role='dialog'], div.modal, div.ui-dialog, div.ui-dialog-content")
        for i in range(min(dialogs.count(), 6)):
            dialog = dialogs.nth(i)
            if not dialog.is_visible():
                continue

            text = ""
            try:
                text = (dialog.inner_text() or "").lower()
            except Exception:
                pass

            has_note_field = False
            try:
                has_note_field = dialog.locator("textarea:visible, input[name*='Note']:visible, textarea[name*='Note']:visible").count() > 0
            except Exception:
                has_note_field = False

            if "sticky" not in text and "note" not in text and not has_note_field:
                continue

            closed = first_visible_click(
                dialog,
                [
                    "button.close",
                    "button:has-text('Close')",
                    "button:has-text('Cancel')",
                    "button:has-text('X')",
                    "[data-dismiss='modal']",
                    "button.ui-dialog-titlebar-close",
                    ".ui-dialog-titlebar-close",
                ],
                timeout=1200,
            )
            if not closed:
                try:
                    page.keyboard.press("Escape")
                    closed = True
                except Exception:
                    closed = False

            if closed:
                page.wait_for_timeout(300)
                print("  DEBUG: Closed existing sticky-note popup")
                return True
    except Exception:
        pass
    return False


def append_to_visible_sticky_popup_if_present(page, note_text):
    """If a sticky-note popup is already open, append note text and save it."""
    note_value = clean(note_text)
    if not note_value:
        return False, False, "No note text"

    try:
        dialogs = page.locator("div[role='dialog']:visible, div.modal:visible, div.ui-dialog:visible, div.ui-dialog-content:visible")
        for i in range(min(dialogs.count(), 6)):
            dialog = dialogs.nth(i)
            try:
                text = (dialog.inner_text() or "").lower()
            except Exception:
                text = ""

            note_fields = dialog.locator("textarea:visible")
            if note_fields.count() == 0:
                continue

            if "sticky" not in text and "note" not in text:
                # still allow if it clearly has note controls
                has_note_controls = dialog.locator("button:has-text('Save'), button:has-text('Save Note')").count() > 0
                if not has_note_controls:
                    continue

            note_input = note_fields.first
            existing = ""
            try:
                existing = (note_input.input_value() or "").strip()
            except Exception:
                existing = ""

            merged = note_value if not existing else (existing if note_value.lower() in existing.lower() else f"{existing}\n{note_value}")
            note_input.fill("")
            note_input.fill(merged)

            sticky_set = False
            sticky_checkbox_selectors = [
                "input[type='checkbox'][name*='Sticky']",
                "input[type='checkbox'][id*='Sticky']",
                "input[type='checkbox'][name*='Pin']",
                "input[type='checkbox'][id*='Pin']",
            ]
            for sel in sticky_checkbox_selectors:
                try:
                    candidate = dialog.locator(sel)
                    if candidate.count() > 0 and candidate.first.is_visible() and candidate.first.is_enabled():
                        if not candidate.first.is_checked():
                            candidate.first.check(timeout=1000)
                        sticky_set = True
                        break
                except Exception:
                    continue

            if not sticky_set:
                sticky_set = first_visible_click(
                    dialog,
                    ["label:has-text('Sticky')", "button:has-text('Sticky')", "label:has-text('Pin')"],
                    timeout=1000,
                )

            save_clicked = first_visible_click(
                dialog,
                [
                    "button:has-text('Save Note')",
                    "input[value='Save Note']",
                    "button:has-text('Save')",
                    "input[value='Save']",
                ],
                timeout=3000,
            )
            if not save_clicked:
                return True, sticky_set, "Sticky popup found but Save button not clickable"

            page.wait_for_timeout(1000)
            print("  DEBUG: Appended to existing sticky-note popup")
            return True, sticky_set, "Appended existing sticky note"
    except Exception as exc:
        return False, False, f"Sticky popup append failed: {exc}"

    return False, False, "No open sticky-note popup detected"


def accept_inventory_date_confirmation(page):
    try:
        dialogs = page.locator("div[role='dialog'], div.modal")
        if dialogs.count() == 0:
            return False

        for i in range(dialogs.count()):
            dialog = dialogs.nth(i)
            if not dialog.is_visible():
                continue

            dialog_text = dialog.inner_text().lower()
            if "inventory date" not in dialog_text and "default costs" not in dialog_text:
                continue

            ok_button = dialog.locator("button:has-text('OK'), button.btn-primary")
            if ok_button.count() == 0:
                continue

            ok_button.first.click(timeout=2000)
            try:
                dialog.wait_for(state="hidden", timeout=5000)
            except Exception:
                page.wait_for_timeout(1000)
            print("  DEBUG: Accepted inventory date confirmation dialog")
            return True

        return False
    except Exception:
        return False


def install_dialog_handler(page):
    def _handle_dialog(dialog):
        try:
            message = dialog.message or ""
            if message:
                print(f"  DEBUG: Browser dialog: '{message}'")

            if record_duplicate_prompt(page, message):
                try:
                    dialog.dismiss()
                    print("  DEBUG: Dismissed duplicate browser confirmation")
                    return
                except Exception as exc:
                    print(f"  DEBUG: Could not dismiss duplicate browser dialog: {exc}")

            dialog.accept()
        except Exception as exc:
            print(f"  DEBUG: Failed to handle browser dialog: {exc}")

    page.on("dialog", _handle_dialog)


def wait_for_ready(page, timeout=10000):
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        pass
    page.wait_for_timeout(500)


def first_visible_fill(page, selectors, value):
    for sel in selectors:
        try:
            loc = page.locator(sel)
            count = loc.count()
            for i in range(count):
                item = loc.nth(i)
                if item.is_visible() and item.is_enabled():
                    item.click(timeout=1000)
                    item.fill("")
                    item.fill(str(value))
                    return True
        except Exception:
            pass
    return False


def first_visible_click(page, selectors, timeout=3000):
    for sel in selectors:
        try:
            loc = page.locator(sel)
            count = loc.count()
            for i in range(count):
                item = loc.nth(i)
                if item.is_visible() and item.is_enabled():
                    item.click(timeout=timeout)
                    return True
        except Exception:
            pass
    return False


def open_new_vehicle_form(page, stock_number=""):
    print(f"Opening Add New Vehicle form for stock number: {stock_number}")
    page.goto(INVENTORY_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(3000)
    close_active_modals(page)

    selectors = [
        "a:has-text('Add New Vehicle')",
        "button:has-text('Add New Vehicle')",
        "a:has-text('Add Vehicle')",
        "button:has-text('Add Vehicle')",
        "a:has-text('New Vehicle')",
        "button:has-text('New Vehicle')",
        "a[href*='/Inventory/Edit']",
    ]

    for sel in selectors:
        try:
            loc = page.locator(sel)
            count = loc.count()
            for i in range(count):
                item = loc.nth(i)
                if not item.is_visible() or not item.is_enabled():
                    continue

                href = item.get_attribute("href") or ""
                text = ""
                try:
                    text = (item.inner_text() or "").strip()
                except Exception:
                    pass

                if text and not any(word in text.lower() for word in ["add", "new", "vehicle"]):
                    if href and "/Inventory/Edit/" in href:
                        continue

                print(f"Found add-vehicle control: selector='{sel}' text='{text}' href='{href}'")
                try:
                    item.click(timeout=5000)
                except Exception as exc:
                    if href:
                        if href.startswith("/"):
                            href = f"https://dm.automanager.com{href}"
                        print(f"Add control click failed: {exc}. Navigating directly to: {href}")
                        page.goto(href, wait_until="domcontentloaded")
                    else:
                        continue

                page.wait_for_timeout(3000)
                close_active_modals(page)
                accept_inventory_date_confirmation(page)
                ensure_on_vehicle_page(page)
                if is_vehicle_detail_page(page):
                    print("Add New Vehicle form is ready!")
                    if VERBOSE_FORM_DUMP:
                        dump_form_elements(page)
                    return
        except Exception:
            continue

    raise Exception("Could not open Add New Vehicle form.")


def open_vehicle_form(page, stock_number):
    search_inventory(page, stock_number)
    open_matching_unit(page, stock_number)


def is_add_vehicle_landing_page(page):
    url = page.url.lower()
    if "/inventory/new" not in url:
        return False
    try:
        continue_loc = page.locator("input[name='btnNewVehicleContinue'], input[value='Continue'], button:has-text('Continue')")
        return continue_loc.count() > 0
    except Exception:
        return False


def click_add_vehicle_continue(page):
    selectors = [
        "input[name='btnNewVehicleContinue']",
        "input[value='Continue']",
        "button:has-text('Continue')",
    ]
    if first_visible_click(page, selectors, timeout=5000):
        return True
    return False


def prepare_add_vehicle_page(page, row):
    print("  Filling add-vehicle first step...")
    initial_fields = [
        ("VIN", "VIN"),
        ("Mileage(Current)", "Mileage(Current)"),
        ("Vehicle Type", "Vehicle Type"),
        ("Location", "Location"),
        ("Date In", "Date In"),
        ("Status", "Status"),
        ("Sub Status", "Sub Status"),
    ]

    for label, key in initial_fields:
        value = row_value(row, key)
        if not value:
            continue
        fill_field(page, label, value)

    clear_duplicate_prompt(page)
    if not click_add_vehicle_continue(page):
        raise Exception("Could not click Continue on Add Vehicle page.")

    wait_for_ready(page, timeout=15000)
    close_active_modals(page)
    duplicate_prompt = get_duplicate_prompt(page)
    if duplicate_prompt:
        raise DuplicateVehicleBlocked(duplicate_prompt)

    ensure_on_vehicle_page(page)


def login(page):
    if MANUAL_LOGIN:
        page.goto(BASE_URL, wait_until="domcontentloaded")
        print("[login] Manual login mode is ON.")
        print("[login] Please sign in in the browser window. Script will continue automatically once logged in.")
        page.wait_for_timeout(1000)

        # Wait until user is truly authenticated on DeskManager (not just any other page).
        page.wait_for_function(
            """
            () => {
                const host = window.location.hostname.toLowerCase();
                const path = window.location.pathname.toLowerCase();
                const onDeskManager = host.includes('dm.automanager.com');
                const offLoginPage = !path.includes('/account/login');
                return onDeskManager && offLoginPage;
            }
            """,
            timeout=MANUAL_LOGIN_TIMEOUT_MS,
        )
        page.wait_for_load_state("networkidle", timeout=60000)

        url_lower = page.url.lower()
        if "dm.automanager.com" not in url_lower or "/account/login" in url_lower:
            raise Exception(
                "Manual login did not finish on DeskManager. "
                "Please complete login on https://dm.automanager.com and avoid unrelated tabs. "
                f"Current URL: {page.url}"
            )

        print(f"[login] Manual login complete. Current URL: {page.url}")
        return

    page.goto(BASE_URL, wait_until="networkidle")
    page.wait_for_selector('input[type="text"], input[type="email"]', timeout=10000)

    username_input = page.locator('input[type="text"], input[type="email"]').first
    password_input = page.locator('input[type="password"]').first

    username_input.fill(AUTOMANAGER_USERNAME)
    page.wait_for_timeout(400)
    password_input.fill(AUTOMANAGER_PASSWORD)
    page.wait_for_timeout(400)

    # Prefer DeskManager's actual submit control, then fall back to generic submit.
    clicked = False
    for selector in ["#btnLogin", 'input[type="submit"]', 'button[type="submit"]']:
        try:
            if page.locator(selector).count() > 0:
                page.locator(selector).first.click(timeout=5000)
                clicked = True
                break
        except Exception:
            pass

    if not clicked:
        try:
            page.evaluate("document.getElementById('btnLogin').click()")
            clicked = True
        except Exception:
            pass

    if not clicked:
        page.keyboard.press("Enter")

    page.wait_for_load_state("networkidle", timeout=60000)
    page.wait_for_timeout(2000)
    print(f"[login] Post-login URL: {page.url}")

    if "/account/login" in page.url.lower():
        msg_selectors = [
            ".validation-summary-errors",
            ".field-validation-error",
            ".text-danger",
            ".alert",
            ".error",
            ".message",
        ]
        login_errors = []
        for sel in msg_selectors:
            try:
                for raw in page.locator(sel).all_inner_texts():
                    text = (raw or "").strip()
                    if text and text not in login_errors:
                        login_errors.append(text)
            except Exception:
                pass

        detail = login_errors[0] if login_errors else "No server error text found"
        raise Exception(f"DeskManager login failed: {detail}")


def search_inventory(page, stock_number):
    # Navigate to inventory — if session redirects to login, wait and retry once
    for _attempt in range(2):
        page.goto(INVENTORY_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        if "/account/login" not in page.url.lower():
            break
        print(f"[search_inventory] Redirected to login, waiting 5s and retrying…")
        page.wait_for_timeout(5000)

    if "/account/login" in page.url.lower():
        raise Exception(f"Session expired: inventory navigation redirected to login for stock {stock_number}")

    # Use DeskManager's server-side inventory search form.
    try:
        page.select_option("#SearchTypeId", "20")  # Stock No.
        page.fill("#SearchText", str(stock_number).strip())
        page.evaluate("document.getElementById('SearchInventory').submit()")
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        page.wait_for_timeout(1500)
    except Exception as e:
        try:
            debug_path = Path("/tmp/dm_search_debug.html")
            debug_path.write_text(page.content(), encoding="utf-8")
            print(f"DEBUG: Wrote inventory page dump to {debug_path} (url={page.url})")
        except Exception as dump_err:
            print(f"DEBUG: Could not write search debug dump: {dump_err}")
        raise Exception(f"Could not run inventory stock search for {stock_number}: {e}")


def search_inventory_by_vin(page, vin):
    """Search DM inventory by VIN. Returns the detail href if exactly one match is found, else None."""
    match = find_inventory_match_by_vin(page, vin)
    if match:
        return match.get("href")
    return None


def find_inventory_match_by_vin(page, vin):
    """Search DM inventory by VIN. Returns the matched row dict if exactly one match is found, else None."""
    vin = str(vin).strip().upper()
    if not vin or len(vin) != 17:
        return None

    page.goto(INVENTORY_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)

    try:
        page.select_option("#SearchTypeId", "15")  # VIN
        page.fill("#SearchText", vin)
        page.evaluate("document.getElementById('SearchInventory').submit()")
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        page.wait_for_timeout(1500)
    except Exception:
        return None

    links = collect_inventory_match_links(page, vin)
    if len(links) == 1:
        return links[0]
    return None


def collect_inventory_match_links(page, search_token):
    token = str(search_token).strip().lower()
    if not token:
        return []

    stock_pattern = re.compile(rf"(^|[^a-z0-9]){re.escape(token)}([^a-z0-9]|$)")
    link_selector = "a[data-link-type='DetailPage'][href*='/Inventory/Edit/'], a[href*='/Inventory/Edit/']:not([href*='#Photos'])"
    matches = []
    seen_hrefs = set()

    try:
        rows = page.locator("tr")
        for i in range(rows.count()):
            row = rows.nth(i)
            try:
                if not row.is_visible():
                    continue
                row_text = " ".join((row.inner_text() or "").lower().split())
            except Exception:
                continue

            if not row_text or not stock_pattern.search(row_text):
                continue

            link = row.locator(link_selector)
            if link.count() == 0:
                continue

            href = link.first.get_attribute("href") or ""
            if not href:
                continue

            if href.startswith("/"):
                href = f"https://dm.automanager.com{href}"

            if href in seen_hrefs:
                continue

            seen_hrefs.add(href)
            matches.append({"href": href, "row_text": row_text})
    except Exception:
        return []

    return matches


def normalize_unit_text(raw_value):
    return clean(raw_value).upper()


def build_unit_search_candidates(unit_value):
    unit = clean(unit_value)
    if not unit:
        return []

    candidates = []
    seen = set()

    def add_candidate(value):
        candidate = clean(value)
        if not candidate:
            return
        key = candidate.upper()
        if key in seen:
            return
        seen.add(key)
        candidates.append(candidate)

    add_candidate(unit)
    add_candidate(unit.replace(" ", ""))

    alnum = "".join(ch for ch in unit if ch.isalnum())
    add_candidate(alnum)

    digits_only = "".join(ch for ch in alnum if ch.isdigit())
    if digits_only and digits_only != alnum:
        add_candidate(digits_only)

    return candidates


def inventory_match_has_stock(match, stock_number):
    row_text = " ".join(str(match.get("row_text", "")).lower().split())
    for candidate in build_unit_search_candidates(stock_number):
        token = str(candidate).strip().lower()
        if not token:
            continue
        stock_pattern = re.compile(rf"(^|[^a-z0-9]){re.escape(token)}([^a-z0-9]|$)")
        if stock_pattern.search(row_text):
            return True
    return False


def find_and_open_unit_for_excel_update(page, unit_value):
    original_unit = clean(unit_value)
    candidates = build_unit_search_candidates(original_unit)
    if not candidates:
        return {
            "status": "Unit not found",
            "matched": False,
            "message": "Unit value is blank",
            "search_query": "",
            "used_cleaned": False,
        }

    for idx, query in enumerate(candidates):
        search_inventory(page, query)
        matches = collect_inventory_match_links(page, query)

        if not matches:
            continue

        if len(matches) > 1:
            return {
                "status": "Multiple matches / needs review",
                "matched": False,
                "message": f"{len(matches)} matches found for search '{query}'",
                "search_query": query,
                "used_cleaned": query.upper() != original_unit.upper(),
            }

        detail_href = matches[0]["href"]
        page.goto(detail_href, wait_until="domcontentloaded")
        wait_for_ready(page, timeout=15000)
        ensure_on_vehicle_page(page)

        used_cleaned = query.upper() != original_unit.upper()
        match_note = "Matched original unit" if not used_cleaned else f"Matched cleaned unit '{query}'"
        return {
            "status": "Matched",
            "matched": True,
            "message": match_note,
            "search_query": query,
            "used_cleaned": used_cleaned,
        }

    return {
        "status": "Unit not found",
        "matched": False,
        "message": "No listing found using original/cleaned unit searches",
        "search_query": "",
        "used_cleaned": False,
    }


def add_sticky_note_to_listing(page, note_text):
    note_value = clean(note_text)
    if not note_value:
        return False, False, "No note text"

    vehicle_edit_url = page.url

    handled_popup, popup_sticky_set, popup_message = append_to_visible_sticky_popup_if_present(page, note_value)
    if handled_popup:
        return True, popup_sticky_set, popup_message

    new_note_selectors = [
        "a:has-text('New Note')",
        "button:has-text('New Note')",
        "span:has-text('New Note')",
        "a:has-text('Add Note')",
        "button:has-text('Add Note')",
        "a[title*='New Note' i]",
        "button[title*='New Note' i]",
        "a[aria-label*='New Note' i]",
        "button[aria-label*='New Note' i]",
        "a[title*='Add Note' i]",
        "button[title*='Add Note' i]",
        "a[onclick*='Note']",
        "button[onclick*='Note']",
        "a[href*='Note']",
        "button[class*='note' i]",
        "a[class*='note' i]",
    ]

    # In the recorded flow, New Note comes from the right "Vehicle Tasks & Notes" panel.
    add_clicked = first_visible_click(page, new_note_selectors, timeout=2500)
    if not add_clicked:
        # Try opening the right notes drawer/panel, then click New Note again.
        panel_openers = [
            "button[title*='note' i]",
            "button[aria-label*='note' i]",
            "a[title*='note' i]",
            "button:has-text('Notes')",
            "a:has-text('Notes')",
            "#intercom-launcher",
            ".intercom-launcher",
            ".chat-launcher",
            "button[class*='launcher']",
            "button[class*='chat']",
        ]
        first_visible_click(page, panel_openers, timeout=1500)
        page.wait_for_timeout(600)
        add_clicked = first_visible_click(page, new_note_selectors, timeout=3000)

    if not add_clicked:
        # Fall back to note controls inside the "Vehicle Tasks & Notes" panel.
        panel_scopes = page.locator("div:has-text('Vehicle Tasks & Notes'), section:has-text('Vehicle Tasks & Notes')")
        for i in range(min(panel_scopes.count(), 3)):
            scope = panel_scopes.nth(i)
            if not scope.is_visible():
                continue
            add_clicked = first_visible_click(
                scope,
                [
                    "a:has-text('New')",
                    "button:has-text('New')",
                    "a:has-text('Add')",
                    "button:has-text('Add')",
                    "a[title*='Note' i]",
                    "button[title*='Note' i]",
                    "a[onclick*='Note']",
                    "button[onclick*='Note']",
                    "a[href*='Note']",
                    "i.fa-plus",
                    ".fa-plus",
                    "i.icon-plus",
                    ".icon-plus",
                ],
                timeout=1500,
            )
            if add_clicked:
                break

    if not add_clicked:
        # Broader fallback: right-side containers with note/task words and plus/new/add icons.
        right_panel_scopes = page.locator(
            "aside:visible, div[class*='right' i]:visible, div[class*='sidebar' i]:visible, div[class*='panel' i]:visible"
        )
        for i in range(min(right_panel_scopes.count(), 8)):
            scope = right_panel_scopes.nth(i)
            try:
                panel_text = (scope.inner_text() or "").lower()
            except Exception:
                panel_text = ""
            if not any(tok in panel_text for tok in ["note", "task", "sticky"]):
                continue

            add_clicked = first_visible_click(
                scope,
                [
                    "a:has-text('New')",
                    "button:has-text('New')",
                    "a:has-text('Add')",
                    "button:has-text('Add')",
                    "a[title*='New' i]",
                    "button[title*='New' i]",
                    "a[title*='Add' i]",
                    "button[title*='Add' i]",
                    "a[aria-label*='New' i]",
                    "button[aria-label*='New' i]",
                    "a[aria-label*='Add' i]",
                    "button[aria-label*='Add' i]",
                    ".fa-plus",
                    ".icon-plus",
                    "[class*='plus' i]",
                    "[class*='add' i]",
                ],
                timeout=1500,
            )
            if add_clicked:
                break

    if not add_clicked:
        # Some unit pages expose note creation via a collapsed tasks menu and sticky-note icon link.
        first_visible_click(
            page,
            [
                "a[href='#menu-tasks']",
                "button[href='#menu-tasks']",
                "a:has(i.fa-tasks)",
                "button:has(i.fa-tasks)",
            ],
            timeout=1200,
        )
        page.wait_for_timeout(350)
        add_clicked = first_visible_click(
            page,
            [
                "a[href*='/Task?'][href*='ttid=20']",
                "a[href*='Task'][href*='ttid=20']",
                "a:has(i.fa-sticky-note)",
                "button:has(i.fa-sticky-note)",
                "i.fa-sticky-note",
                ".fa-sticky-note",
                "a[href*='ttid=20']",
                "a[href*='ttid='][href*='Task']",
            ],
            timeout=2500,
        )

    existing_editor_ready = False
    if not add_clicked:
        # Some listings load directly into an already-open note editor.
        try:
            existing_editor_ready = page.locator("textarea:visible").count() > 0
        except Exception:
            existing_editor_ready = False

    if not add_clicked and not existing_editor_ready:
        if VERBOSE_FORM_DUMP:
            dump_note_controls(page)
        return False, False, "Could not find New Note action in right panel"

    page.wait_for_timeout(500)

    scope = page
    try:
        visible_dialogs = page.locator("div[role='dialog']:visible, div.modal:visible, div.ui-dialog:visible, div.ui-dialog-content:visible")
        best_scope = None
        for i in range(min(visible_dialogs.count(), 8)):
            candidate = visible_dialogs.nth(i)
            try:
                has_note_editor = candidate.locator(
                    "textarea:visible, input[name*='Description' i]:visible, textarea[name*='Description' i]:visible, input[name*='Subject' i]:visible"
                ).count() > 0
            except Exception:
                has_note_editor = False

            if has_note_editor:
                best_scope = candidate
                break

        if best_scope is None and visible_dialogs.count() > 0:
            best_scope = visible_dialogs.first

        if best_scope is not None:
            scope = best_scope
    except Exception:
        pass

    # Subject in recorded UI is separate from note details body.
    subject_value = note_value.splitlines()[0].strip()[:80] if note_value.splitlines() else note_value[:80]
    if not subject_value:
        subject_value = "Note"

    subject_filled = False
    subject_selectors = [
        "input[name*='Subject']:visible",
        "input[id*='Subject']:visible",
        "input[placeholder*='Subject']:visible",
    ]
    for sel in subject_selectors:
        try:
            target = scope.locator(sel)
            if target.count() > 0 and target.first.is_visible() and target.first.is_enabled():
                target.first.fill("")
                target.first.fill(subject_value)
                subject_filled = True
                break
        except Exception:
            continue

    if not subject_filled:
        try:
            inputs = scope.locator("input:visible")
            for i in range(inputs.count()):
                item = inputs.nth(i)
                input_type = (item.get_attribute("type") or "").lower()
                if input_type in {"checkbox", "radio", "file", "hidden", "button", "submit"}:
                    continue
                item.fill("")
                item.fill(subject_value)
                subject_filled = True
                break
        except Exception:
            pass

    sticky_checkbox_selectors = [
        "input[type='checkbox'][name*='Sticky']",
        "input[type='checkbox'][id*='Sticky']",
        "input[type='checkbox'][name*='Pin']",
        "input[type='checkbox'][id*='Pin']",
    ]

    note_filled = False
    note_inputs = scope.locator("textarea:visible")
    for i in range(note_inputs.count()):
        item = note_inputs.nth(i)
        try:
            existing = ""
            try:
                existing = (item.input_value() or "").strip()
            except Exception:
                existing = ""

            merged = note_value
            if existing:
                if note_value.lower() in existing.lower():
                    merged = existing
                else:
                    merged = f"{existing}\n{note_value}"

            item.fill("")
            item.fill(merged)
            note_filled = True
            break
        except Exception:
            continue

    if not note_filled:
        # Some DeskManager note/task UIs use contenteditable divs.
        editable_blocks = scope.locator("[contenteditable='true']:visible")
        for i in range(editable_blocks.count()):
            item = editable_blocks.nth(i)
            try:
                item.click(timeout=1000)
                item.fill("")
                item.type(note_value)
                note_filled = True
                break
            except Exception:
                continue

    if not note_filled:
        # Fallback to note-specific text inputs used by some task forms.
        note_field_selectors = [
            "textarea[name*='Note' i]:visible",
            "textarea[id*='Note' i]:visible",
            "textarea[name*='Description' i]:visible",
            "textarea[id*='Description' i]:visible",
            "textarea[name*='Comment' i]:visible",
            "textarea[id*='Comment' i]:visible",
            "input[name*='Note' i]:visible",
            "input[id*='Note' i]:visible",
            "input[name*='Description' i]:visible",
            "input[id*='Description' i]:visible",
            "input[name*='Comment' i]:visible",
            "input[id*='Comment' i]:visible",
            "input[placeholder*='Note' i]:visible",
            "input[placeholder*='Comment' i]:visible",
            "textarea[placeholder*='Note' i]:visible",
            "textarea[placeholder*='Comment' i]:visible",
        ]
        for sel in note_field_selectors:
            try:
                target = scope.locator(sel)
                if target.count() == 0:
                    continue
                target.first.fill("")
                target.first.fill(note_value)
                note_filled = True
                break
            except Exception:
                continue

    if not note_filled:
        # Fallback: open the New Note task form directly when popup editor has no editable field.
        note_link_selectors = [
            "a:has-text('New Note')",
            "a[href*='/Task/New'][href*='cat=20']",
            "a[href*='Task/New'][href*='cat=20']",
            "a[href*='/Task/New'][href*='ttid=20']",
        ]
        note_href = ""
        for sel in note_link_selectors:
            try:
                loc = page.locator(sel)
                if loc.count() == 0:
                    continue
                candidate = loc.first.get_attribute("href") or ""
                if candidate:
                    note_href = candidate
                    break
            except Exception:
                continue

        if note_href:
            try:
                if note_href.startswith("/"):
                    note_href = f"https://dm.automanager.com{note_href}"

                page.goto(note_href, wait_until="domcontentloaded")
                wait_for_ready(page, timeout=15000)

                direct_subject_selectors = [
                    "input[name*='Subject' i]:visible",
                    "input[id*='Subject' i]:visible",
                    "input[name*='Title' i]:visible",
                    "input[id*='Title' i]:visible",
                    "input[name*='Name' i]:visible",
                    "input[id*='Name' i]:visible",
                ]
                for sel in direct_subject_selectors:
                    try:
                        loc = page.locator(sel)
                        if loc.count() > 0 and loc.first.is_visible() and loc.first.is_enabled():
                            loc.first.fill("")
                            loc.first.fill(subject_value)
                            break
                    except Exception:
                        continue

                direct_note_selectors = [
                    "textarea[name*='Description' i]:visible",
                    "textarea[id*='Description' i]:visible",
                    "textarea[name*='Note' i]:visible",
                    "textarea[id*='Note' i]:visible",
                    "textarea:visible",
                    "input[name*='Description' i]:visible",
                    "input[id*='Description' i]:visible",
                    "input[name*='Note' i]:visible",
                    "input[id*='Note' i]:visible",
                ]

                direct_note_filled = False
                for sel in direct_note_selectors:
                    try:
                        loc = page.locator(sel)
                        if loc.count() == 0:
                            continue
                        loc.first.fill("")
                        loc.first.fill(note_value)
                        direct_note_filled = True
                        break
                    except Exception:
                        continue

                if not direct_note_filled:
                    try:
                        editable = page.locator("[contenteditable='true']:visible")
                        if editable.count() > 0:
                            editable.first.click(timeout=1000)
                            editable.first.fill("")
                            editable.first.type(note_value)
                            direct_note_filled = True
                    except Exception:
                        pass

                direct_sticky_set = False
                for sel in sticky_checkbox_selectors:
                    try:
                        candidate = page.locator(sel)
                        if candidate.count() > 0 and candidate.first.is_visible() and candidate.first.is_enabled():
                            if not candidate.first.is_checked():
                                candidate.first.check(timeout=1000)
                            direct_sticky_set = True
                            break
                    except Exception:
                        continue

                direct_save_clicked = first_visible_click(
                    page,
                    [
                        "button:has-text('Save Note')",
                        "input[value='Save Note']",
                        "button:has-text('Save Task')",
                        "input[value='Save Task']",
                        "button:has-text('Save')",
                        "input[value='Save']",
                        "button:has-text('Create')",
                        "input[value='Create']",
                        "button:has-text('Add')",
                        "input[value='Add']",
                        "button:has-text('Submit')",
                        "input[value='Submit']",
                        "button[type='submit']",
                        "input[type='submit']",
                    ],
                    timeout=4000,
                )

                page.wait_for_timeout(1200)
                if vehicle_edit_url:
                    page.goto(vehicle_edit_url, wait_until="domcontentloaded")
                    wait_for_ready(page, timeout=15000)
                    ensure_on_vehicle_page(page)

                if direct_note_filled and direct_save_clicked:
                    return True, direct_sticky_set, "Note added via direct task form"
            except Exception:
                pass

        if VERBOSE_FORM_DUMP:
            print(f"  DEBUG: Current URL while adding note: {page.url}")
            dump_note_controls(page)
            dump_form_elements(page)
        return False, False, "Could not find editable note details field"

    sticky_set = False
    for sel in sticky_checkbox_selectors:
        try:
            candidate = scope.locator(sel)
            if candidate.count() > 0 and candidate.first.is_visible() and candidate.first.is_enabled():
                if not candidate.first.is_checked():
                    candidate.first.check(timeout=1000)
                sticky_set = True
                break
        except Exception:
            continue

    if not sticky_set:
        try:
            sticky_set = first_visible_click(
                scope,
                ["label:has-text('Sticky')", "button:has-text('Sticky')", "label:has-text('Pin')"],
                timeout=1000,
            )
        except Exception:
            sticky_set = False

    save_clicked = first_visible_click(
        scope,
        [
            "button:has-text('Save Note')",
            "input[value='Save Note']",
            "input[name='SaveNote']",
            "input[id*='SaveNote']",
            "a.saveTask",
            "button.saveTask",
            "input.saveTask",
            "button:has-text('Save Task')",
            "input[value='Save Task']",
            "button:has-text('Save')",
            "input[value='Save']",
            "button:has-text('Create')",
            "input[value='Create']",
            "button:has-text('Add')",
            "input[value='Add']",
            "button:has-text('Submit')",
            "input[value='Submit']",
            "button:has-text('Done')",
            "input[value='Done']",
        ],
        timeout=3000,
    )
    if not save_clicked:
        try:
            # Some note forms submit on Enter/Ctrl+Enter instead of a clickable button.
            page.keyboard.press("Control+Enter")
            page.wait_for_timeout(600)
            save_clicked = True
        except Exception:
            try:
                page.keyboard.press("Enter")
                page.wait_for_timeout(600)
                save_clicked = True
            except Exception:
                save_clicked = False
    if not save_clicked:
        return False, sticky_set, "Could not click Save for note"

    page.wait_for_timeout(1200)
    return True, sticky_set, "Note added"


def run_excel_location_note_updates():
    source_file = resolve_excel_update_file()
    if not source_file.exists():
        print(f"Missing Excel input file: {source_file}")
        return

    print(f"Using update input file: {source_file}")
    try:
        suffix = source_file.suffix.lower()
        if suffix == ".csv":
            df = pd.read_csv(source_file, dtype=str)
        else:
            df = pd.read_excel(source_file, dtype=str)
    except Exception as e:
        print(f"Could not read update file {source_file}: {e}")
        return

    required_columns = ["Unit", "Notes", "Last 6 VIN", "Location"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"Excel missing required columns: {missing_columns}")
        print(f"Found columns: {df.columns.tolist()}")
        return

    units_preview = [clean(v) for v in df["Unit"].fillna("").tolist() if clean(v)]
    preview_text = ", ".join(units_preview[:5])
    if len(units_preview) > 5:
        preview_text += f" ... (+{len(units_preview) - 5} more)"

    print(f"\nAbout to UPDATE {len(df)} unit(s) from Excel: {preview_text}")
    try:
        answer = input("Proceed? (y/N): ").strip().lower()
    except EOFError:
        answer = "y"
    if answer not in {"y", "yes"}:
        print("Aborted by user.")
        return

    log_rows = []

    with sync_playwright() as p:
        browser = None
        context = None
        page = None

        try:
            try:
                browser, context, page = create_authenticated_session(p)
                print("Login successful")
            except Exception as e:
                print(f"Could not start authenticated session: {e}")
                return

            total_rows = len(df)
            for idx, (_, row) in enumerate(df.iterrows(), 1):
                unit = clean(row.get("Unit", ""))
                location = clean(row.get("Location", ""))
                note_text = clean(row.get("Notes", ""))
                last6 = clean(row.get("Last 6 VIN", ""))

                entry = {
                    "Unit": unit,
                    "Last 6 VIN": last6,
                    "Location": location,
                    "Notes": note_text,
                    "Result": "",
                    "Updated": "No",
                    "Location changed": "No",
                    "Sticky note added": "No",
                    "Search Query": "",
                    "Message": "",
                }

                if not unit:
                    entry["Result"] = "Any error"
                    entry["Message"] = "Missing Unit value"
                    log_rows.append(entry)
                    print(f"[{idx}/{total_rows}] ERROR - missing Unit value")
                    continue

                print(f"\n[{idx}/{total_rows}] Processing Unit {unit}")

                try:
                    close_active_modals(page)
                    match = find_and_open_unit_for_excel_update(page, unit)
                    entry["Search Query"] = match.get("search_query", "")

                    if not match.get("matched"):
                        entry["Result"] = match.get("status", "Unit not found")
                        entry["Message"] = match.get("message", "")
                        log_rows.append(entry)
                        print(f"[{idx}/{total_rows}] {unit} - {entry['Result']}")
                        continue

                    sticky_added = False
                    if note_text:
                        note_added, sticky_set, note_message = add_sticky_note_to_listing(page, note_text)
                        if not note_added:
                            raise Exception(f"Failed to add note: {note_message}")
                        sticky_added = sticky_set
                        # Note/task modals can leave an overlay that blocks Location and Save clicks.
                        close_sticky_note_popup_if_present(page)
                        close_active_modals(page)

                    location_changed = False
                    if not note_text:
                        close_sticky_note_popup_if_present(page)
                    if location:
                        location_changed = fill_field(page, "Location", location)

                    save_vehicle(page, pd.Series({"Location": location}))

                    entry["Updated"] = "Yes"
                    entry["Location changed"] = "Yes" if location_changed else "No"
                    entry["Sticky note added"] = "Yes" if sticky_added else "No"
                    entry["Result"] = "Updated"
                    details = []
                    if location_changed:
                        details.append("Location changed")
                    if sticky_added:
                        details.append("Sticky note added")
                    if match.get("used_cleaned"):
                        details.append("Matched using cleaned unit")
                    entry["Message"] = "; ".join(details)
                    log_rows.append(entry)
                    print(f"[{idx}/{total_rows}] {unit} - COMPLETE")

                except Exception as e:
                    entry["Result"] = "Any error"
                    entry["Message"] = str(e)
                    log_rows.append(entry)
                    print(f"[{idx}/{total_rows}] ERROR processing {unit}: {e}")

        finally:
            log_file = resolve_update_log_file()
            try:
                pd.DataFrame(log_rows).to_csv(log_file, index=False)
                print(f"\nWrote update log: {log_file}")
            except Exception as e:
                print(f"\nCould not write update log: {e}")

            print("\nFinished processing. Closing browser...")
            try:
                if browser is not None:
                    browser.close()
            except Exception:
                pass


def open_matching_unit(page, stock_number):
    print(f"Searching for stock number: {stock_number}")
    page.wait_for_timeout(2000)

    stock_token = str(stock_number).strip().lower()
    stock_pattern = re.compile(rf"(^|[^a-z0-9]){re.escape(stock_token)}([^a-z0-9]|$)")
    link_selector = "a[data-link-type='DetailPage'][href*='/Inventory/Edit/'], a[href*='/Inventory/Edit/']:not([href*='#Photos'])"

    detail_href = ""
    max_pages = 1

    for page_idx in range(max_pages):
        page_number = page_idx + 1
        if page_idx > 0:
            page_url = f"{INVENTORY_URL}&page={page_number}"
            print(f"DEBUG: Loading inventory page {page_number} via URL")
            page.goto(page_url, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)
        else:
            print("DEBUG: Scanning inventory page 1")

        # Pass 1: check visible rows for exact stock token match.
        try:
            rows = page.locator("tr")
            for i in range(rows.count()):
                row = rows.nth(i)
                try:
                    if not row.is_visible():
                        continue
                    row_text = " ".join((row.inner_text() or "").lower().split())
                except Exception:
                    continue

                if not row_text or not stock_pattern.search(row_text):
                    continue

                link = row.locator(link_selector)
                if link.count() == 0:
                    continue

                href = link.first.get_attribute("href") or ""
                if href:
                    detail_href = href
                    print(f"Found detail link in matched row: {detail_href}")
                    break
        except Exception as e:
            print(f"DEBUG: Row scan error on page {page_number}: {e}")

        # Pass 2: fallback by validating each detail link's nearby text.
        if not detail_href:
            try:
                links = page.locator(link_selector)
                for i in range(links.count()):
                    link = links.nth(i)
                    if not link.is_visible():
                        continue

                    context_text = link.evaluate(
                        """
                        el => {
                            const row = el.closest('tr');
                            if (row && row.innerText) return row.innerText;
                            const card = el.closest('.card, .row, .item, .inventory-row');
                            if (card && card.innerText) return card.innerText;
                            return el.innerText || '';
                        }
                        """
                    )
                    normalized = " ".join(str(context_text or "").lower().split())
                    if not normalized or not stock_pattern.search(normalized):
                        continue

                    href = link.get_attribute("href") or ""
                    if href:
                        detail_href = href
                        print(f"Found detail link in matched context: {detail_href}")
                        break
            except Exception as e:
                print(f"DEBUG: Link scan error on page {page_number}: {e}")

        if detail_href:
            break

    if not detail_href:
        raise Exception(
            f"Could not find exact vehicle row/detail link for stock number {stock_number}. "
            "Skipping to avoid editing the wrong vehicle."
        )

    if detail_href.startswith('/'):
        detail_href = f"https://dm.automanager.com{detail_href}"

    print(f"Navigating directly to detail URL: {detail_href}")
    page.goto(detail_href, wait_until='domcontentloaded')
    wait_for_ready(page, timeout=15000)
    ensure_on_vehicle_page(page)
    print("Vehicle form is ready!")
    if VERBOSE_FORM_DUMP:
        dump_form_elements(page)

def is_logged_out(page):
    url = page.url.lower()
    if "loggedout" in url or "login" in url or "sign in" in page.content().lower():
        return True
    return False


def maybe_click_details_tab(page):
    try:
        page.get_by_text("Details", exact=True).click(timeout=3000)
        page.wait_for_timeout(1500)
    except Exception:
        pass


def get_visible_inputs(page):
    visible = []
    loc = page.locator("input")
    skip_types = {"hidden", "password", "checkbox", "radio", "submit", "reset", "file", "image", "button"}
    for i in range(loc.count()):
        item = loc.nth(i)
        try:
            input_type = (item.get_attribute("type") or "").lower()
            if input_type in skip_types:
                continue
            if item.is_visible() and item.is_enabled():
                visible.append(item)
        except Exception:
            pass
    return visible


def get_visible_selects(page):
    visible = []
    loc = page.locator("select")
    for i in range(loc.count()):
        item = loc.nth(i)
        try:
            name = (item.get_attribute("name") or "").lower()
            element_id = (item.get_attribute("id") or "").lower()
            if any(skip in name for skip in ["search", "filter", "sort", "page"]):
                continue
            if any(skip in element_id for skip in ["search", "filter", "sort", "page"]):
                continue
            if item.is_visible() and item.is_enabled():
                visible.append(item)
        except Exception:
            pass
    return visible


def is_vehicle_detail_page(page):
    url = page.url.lower()
    if "loggedout" in url or "login" in url:
        return False
    if any(token in url for token in ["detail", "edit", "/vehicle", "vehicle/"]):
        return True
    # Count visible inputs and selects
    visible_inputs = len(get_visible_inputs(page))
    visible_selects = len(get_visible_selects(page))
    print(f"  DEBUG: Found {visible_inputs} inputs, {visible_selects} selects")
    if visible_inputs + visible_selects >= 5:  # Lowered threshold from 10
        return True
    return False


def fill_field(page, label, value):
    if not value:
        return False
    value = normalize_field_value(label, value)
    if not value:
        return False

    print(f"  DEBUG: Looking for field '{label}' with value '{value}'")
    accept_inventory_date_confirmation(page)
    close_active_modals(page)

    def try_fill(loc):
        for i in range(loc.count()):
            item = loc.nth(i)
            try:
                if not item.is_visible() or not item.is_enabled():
                    continue
                tag = item.evaluate("el => el.tagName.toLowerCase()")
                print(f"    DEBUG: Found {tag} element, trying to fill...")
                if tag == "select":
                    accept_inventory_date_confirmation(page)
                    close_active_modals(page)
                    item.click(timeout=1000)
                    page.wait_for_timeout(200)

                    try:
                        page.wait_for_function(
                            "el => (el.options && el.options.length) ? true : false",
                            arg=item,
                            timeout=1500,
                        )
                    except Exception:
                        pass

                    options = item.evaluate(
                        """
                        el => Array.from(el.options || []).map(option => ({
                            label: (option.label || '').trim(),
                            value: option.value || ''
                        }))
                        """
                    )

                    try:
                        item.select_option(label=value, timeout=2000)
                        print(f"    DEBUG: Selected option '{value}' in select by label")
                        return True
                    except Exception:
                        pass

                    try:
                        item.select_option(value=value, timeout=2000)
                        print(f"    DEBUG: Selected option '{value}' in select by value")
                        return True
                    except Exception:
                        pass

                    target = value.lower()
                    target_norm = normalize_match_key(value)
                    target_tokens = [t for t in target_norm.split() if t]

                    # Fuzzy matching across normalized option label/value strings.
                    for option in options:
                        option_label = option["label"].lower()
                        option_value = str(option["value"]).lower()
                        option_label_norm = normalize_match_key(option_label)
                        option_value_norm = normalize_match_key(option_value)

                        exact_match = (
                            target == option_label
                            or target == option_value
                            or target_norm == option_label_norm
                            or target_norm == option_value_norm
                        )

                        partial_match = (
                            target in option_label
                            or (target_norm and target_norm in option_label_norm)
                            or (target_norm and target_norm in option_value_norm)
                        )

                        token_match = False
                        if target_tokens:
                            token_match = all(tok in option_label_norm for tok in target_tokens)

                        if exact_match or partial_match or token_match:
                            item.select_option(value=option["value"], timeout=2000)
                            print(f"    DEBUG: Selected option '{option['label']}' by fuzzy match")
                            return True

                    # Last attempt for selects: iterate options and pick first non-empty label containing target words.
                    if options:
                        for option in options:
                            label_norm = normalize_match_key(option["label"])
                            if target_norm and label_norm and (target_norm in label_norm or label_norm in target_norm):
                                item.select_option(value=option["value"], timeout=2000)
                                print(f"    DEBUG: Selected option '{option['label']}' by broad match")
                                return True

                    item.evaluate(
                        """
                        (el, val) => {
                            el.value = val;
                            el.dispatchEvent(new Event('input', { bubbles: true }));
                            el.dispatchEvent(new Event('change', { bubbles: true }));
                        }
                        """,
                        value,
                    )
                    print(f"    DEBUG: Set value '{value}' directly in select")
                    return True
                            
                if tag in {"input", "textarea"}:
                    accept_inventory_date_confirmation(page)
                    close_active_modals(page)
                    input_type = (item.get_attribute("type") or "").lower()
                    if input_type == "checkbox":
                        should_check = is_truthy(value)
                        if item.is_checked() != should_check:
                            if should_check:
                                item.check(timeout=1000)
                            else:
                                item.uncheck(timeout=1000)
                        print(f"    DEBUG: Set checkbox to '{should_check}'")
                        return True
                    try:
                        item.click(timeout=1000)
                        page.wait_for_timeout(100)
                    except Exception:
                        try:
                            item.focus()
                        except Exception:
                            pass
                    item.fill("")
                    page.wait_for_timeout(200)
                    item.fill(value)
                    if label == "Date In":
                        accept_inventory_date_confirmation(page)
                    print(f"    DEBUG: Filled '{value}' in {tag}")
                    return True
            except Exception as e:
                print(f"    DEBUG: Error filling element: {e}")
                continue
        return False

    variants = FIELD_LABEL_ALIASES.get(label, [label])

    # Primary strategy: exact selectors for known DeskManager fields.
    for sel in FIELD_SELECTORS.get(label, []):
        try:
            print(f"    DEBUG: Trying custom selector '{sel}'")
            if try_fill(page.locator(sel)):
                print(f"    DEBUG: Successfully filled by custom selector '{sel}'")
                return True
        except Exception as e:
            print(f"      DEBUG: Custom selector failed: {e}")

    # Secondary strategy: fill by associated label.
    for variant in variants:
        try:
            print(f"    DEBUG: Trying label variant '{variant}'")
            label_loc = page.get_by_label(variant, exact=False)
            if try_fill(label_loc):
                print(f"    DEBUG: Successfully filled by label '{variant}'")
                return True
        except Exception as e:
            print(f"    DEBUG: Label search failed: {e}")

        try:
            label_text = variant.replace("'", "\\'")
            nearby = page.locator(
                f"xpath=//label[contains(normalize-space(.), '{label_text}')]/following::*[self::input or self::select or self::textarea][1]"
            )
            if try_fill(nearby):
                print(f"    DEBUG: Successfully filled by nearby label '{variant}'")
                return True
        except Exception as e:
            print(f"    DEBUG: Nearby label search failed: {e}")

    # Fallback strategy: match by name or id patterns using alias variants.
    label_keys = ["".join(ch for ch in variant.lower() if ch.isalnum()) for variant in variants]
    seen = set()
    for key in label_keys:
        if key in seen:
            continue
        seen.add(key)
        print(f"    DEBUG: Trying name/id pattern '{key}'")
        selectors = [
            f"input[name*='{key}']",
            f"input[id*='{key}']",
            f"select[name*='{key}']",
            f"select[id*='{key}']",
            f"textarea[name*='{key}']",
            f"textarea[id*='{key}']",
        ]
        for sel in selectors:
            try:
                print(f"      DEBUG: Trying selector '{sel}'")
                if try_fill(page.locator(sel)):
                    print(f"    DEBUG: Successfully filled by selector '{sel}'")
                    return True
            except Exception as e:
                print(f"      DEBUG: Selector failed: {e}")

    # Last resort: try to find any input/select that might match
    print(f"    DEBUG: Trying generic input/select search...")
    try:
        # Look for all inputs and selects and try to match by placeholder or nearby text
        all_inputs = page.locator("input:visible, select:visible, textarea:visible")
        for i in range(min(20, all_inputs.count())):
            inp = all_inputs.nth(i)
            try:
                placeholder = (inp.get_attribute("placeholder") or "").lower()
                name = (inp.get_attribute("name") or "").lower()
                id_attr = (inp.get_attribute("id") or "").lower()

                # Check if any variant matches the placeholder, name, or id
                for variant in variants:
                    variant_lower = variant.lower()
                    if (variant_lower in placeholder or
                        variant_lower in name or
                        variant_lower in id_attr):
                        print(f"      DEBUG: Found matching element by attribute, trying to fill...")
                        if try_fill(inp):
                            print(f"    DEBUG: Successfully filled by attribute match")
                            return True
            except Exception:
                continue
    except Exception as e:
        print(f"    DEBUG: Generic search failed: {e}")

    print(f"Warning: Could not find field for '{label}'.")
    return False


def ensure_on_vehicle_page(page):
    # Wait for form fields to appear in modal or on page
    print(f"  Waiting for form fields to appear...")
    for attempt in range(10):
        page.wait_for_timeout(1000)

        # If the URL indicates a detail/edit page, we're ready
        if is_vehicle_detail_page(page):
            print(f"  Vehicle detail page detected (attempt {attempt+1})")
            return

        # Check if modal is visible with form fields
        try:
            modal = page.locator("div.mantine-Modal-root:visible, div[role='dialog']:visible")
            if modal.count() > 0:
                # Look for inputs/labels in modal
                modal_inputs = modal.first.locator("input:visible")
                modal_labels = modal.first.locator("label")
                if modal_inputs.count() > 3 or modal_labels.count() > 3:
                    print(f"  Modal with form detected (attempt {attempt+1})")
                    return
        except Exception:
            pass

        # Check if form is on main page (not just search fields)
        try:
            all_inputs = page.locator("input:visible")
            all_labels = page.locator("label")
            # If we have enough fields (more than just search), we're ready
            if all_inputs.count() > 5 and all_labels.count() > 10:
                print(f"  Form fields detected on page (attempt {attempt+1})")
                return
        except Exception:
            pass

    print(f"  Waiting for form... fields might load dynamically")


def click_save(page):
    try:
        page.get_by_text("Save", exact=True).click(timeout=5000)
        return True
    except Exception:
        return first_visible_click(page, ["button:has-text('Save')", "input[value='Save']"], timeout=5000)


def placeholder_for_kind(kind, numeric_index=0):
    if kind == "numeric":
        i = max(0, min(numeric_index, len(NUMERIC_PLACEHOLDER_CANDIDATES) - 1))
        return NUMERIC_PLACEHOLDER_CANDIDATES[i]
    if kind == "date":
        return PLACEHOLDER_DATE_VALUE
    return PLACEHOLDER_TEXT_VALUE


def fill_placeholder_for_required_field(page, field_label, placeholder_value):
    """Fill one required field placeholder and log it."""
    try:
        if fill_field(page, field_label, placeholder_value):
            print(f"  PLACEHOLDER FILLED: '{field_label}' = '{placeholder_value}' (user should update)")
            return True
    except Exception:
        pass
    return False


def fill_all_required_placeholders(page, row, numeric_index=0):
    """Fill placeholders only for required fields whose CSV value is missing."""
    if not FILL_REQUIRED_PLACEHOLDERS:
        return 0

    filled_count = 0
    for field_label, rule in REQUIRED_FIELD_RULES.items():
        csv_key = rule["csv_key"]
        if row_value(row, csv_key):
            continue
        placeholder_value = placeholder_for_kind(rule["kind"], numeric_index=numeric_index)
        if fill_placeholder_for_required_field(page, field_label, placeholder_value):
            filled_count += 1
    if filled_count > 0:
        print(f"  Filled {filled_count} required fields with placeholders")
    return filled_count


def collect_validation_errors(page):
    """Collect visible validation error messages after Save."""
    selectors = [
        "div.validation-summary-errors li",
        "span.field-validation-error",
        ".validation-summary-errors",
    ]
    messages = []
    seen = set()
    for sel in selectors:
        try:
            loc = page.locator(sel)
            count = loc.count()
            for i in range(count):
                item = loc.nth(i)
                if not item.is_visible():
                    continue
                text = (item.inner_text() or "").strip()
                if not text:
                    continue
                key = text.lower()
                if key in seen:
                    continue
                seen.add(key)
                messages.append(text)
        except Exception:
            continue
    return messages


def save_vehicle(page, row):
    max_attempts = len(NUMERIC_PLACEHOLDER_CANDIDATES) + 1

    for attempt in range(max_attempts):
        clear_duplicate_prompt(page)
        close_active_modals(page)
        if not click_save(page):
            raise Exception("Could not click Save button.")

        wait_for_ready(page, timeout=15000)
        close_active_modals(page)

        duplicate_prompt = get_duplicate_prompt(page)
        if duplicate_prompt:
            raise DuplicateVehicleBlocked(duplicate_prompt)

        if is_logged_out(page):
            raise Exception("Logged out after save attempt.")

        errors = collect_validation_errors(page)
        if not errors:
            # DeskManager may keep us on edit page even after successful save.
            if is_vehicle_detail_page(page):
                page.wait_for_timeout(1200)
                try:
                    page.goto(INVENTORY_URL, wait_until="domcontentloaded")
                    wait_for_ready(page)
                except Exception:
                    pass
            return

        if attempt >= len(NUMERIC_PLACEHOLDER_CANDIDATES):
            break

        if not FILL_REQUIRED_PLACEHOLDERS:
            break

        print("  WARNING: Save blocked by validation errors:")
        for msg in errors:
            print(f"    - {msg}")
        print("  Retrying with required placeholders...")
        fill_all_required_placeholders(page, row, numeric_index=attempt)
        page.wait_for_timeout(600)

    final_errors = collect_validation_errors(page)
    if final_errors:
        raise Exception(f"Save failed after retries. Validation errors: {final_errors}")
    raise Exception("Save failed after retries; validation is still blocking submission.")


def create_authenticated_session(playwright, max_attempts=3):
    last_error = None

    for attempt in range(1, max_attempts + 1):
        browser = None
        context = None
        page = None
        try:
            browser = playwright.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
            context = browser.new_context()
            page = context.new_page()
            install_dialog_handler(page)
            login(page)
            return browser, context, page
        except Exception as exc:
            last_error = exc
            print(f"Session startup attempt {attempt}/{max_attempts} failed: {exc}")
            try:
                if page is not None:
                    page.close()
            except Exception:
                pass
            try:
                if context is not None:
                    context.close()
            except Exception:
                pass
            try:
                if browser is not None:
                    browser.close()
            except Exception:
                pass

    raise Exception(f"Could not create authenticated session after {max_attempts} attempts: {last_error}")


def is_closed_page_error(exc):
    message = str(exc).lower()
    return (
        "target page, context or browser has been closed" in message
        or "page has been closed" in message
        or "browser has been closed" in message
        or "context has been closed" in message
    )


def is_recoverable_session_error(exc):
    message = str(exc).lower()
    return (
        is_closed_page_error(exc)
        or "waiting for locator(\"#searchtypeid\")" in message
        or "could not run inventory stock search" in message
        or "timeout" in message and "searchtypeid" in message
        or "execution context was destroyed" in message
        or "navigation" in message and "timeout" in message
    )


def fill_vehicle_page(page, row):
    print(f"  Starting to fill vehicle form...")
    page.wait_for_timeout(2000)
    maybe_click_details_tab(page)
    page.wait_for_timeout(1000)
    ensure_on_vehicle_page(page)
    page.wait_for_timeout(1000)
    
    # Debug: Show what we're trying to fill
    print(f"  DEBUG: CSV data for this vehicle:")
    for key, value in row.items():
        if pd.notna(value) and str(value).strip():
            print(f"    {key}: '{value}'")

    attempted_total = 0
    filled_total = 0

    mapped_keys = {key for _, key in TOP_LEVEL_FIELD_CONFIG}
    for _, fields in TAB_FIELD_CONFIG:
        for _, key in fields:
            mapped_keys.add(key)

    for label, key in TOP_LEVEL_FIELD_CONFIG:
        value = row_value(row, key)
        if not value:
            continue
        attempted_total += 1
        if fill_field(page, label, value):
            filled_total += 1

    # DeskManager data sometimes has Purchase Method and Payment Method values swapped in CSVs.
    purchase_method_val = row_value(row, "Purchase Method")
    payment_method_val = row_value(row, "Payment Method")
    payment_like_tokens = ["wire", "ach", "visa", "master", "discover", "amex", "cash", "check", "debit", "credit", "deposit", "online", "money order"]
    purchase_like_tokens = ["purchase from seller", "consignor"]
    purchase_looks_payment = any(tok in purchase_method_val.lower() for tok in payment_like_tokens)
    payment_looks_purchase = any(tok in payment_method_val.lower() for tok in purchase_like_tokens)
    if purchase_method_val and payment_method_val and purchase_looks_payment and payment_looks_purchase:
        print("  DEBUG: Detected swapped Purchase Method/Payment Method values in CSV; swapping for this unit")
        row = row.copy()
        row["Purchase Method"] = payment_method_val
        row["Payment Method"] = purchase_method_val

    for tab_name, field_specs in TAB_FIELD_CONFIG:
        attempted, filled = fill_tab_fields(page, row, tab_name, field_specs)
        attempted_total += attempted
        filled_total += filled
        if tab_name == "Purchase Info":
            extra_attempted, extra_filled = fill_purchase_tab_extras(page, row, mapped_keys)
            attempted_total += extra_attempted
            filled_total += extra_filled

    remaining_values = collect_unmapped_nonempty_values(row, mapped_keys)
    fallback_attempted, fallback_filled = fill_remaining_columns_across_tabs(page, remaining_values)
    attempted_total += fallback_attempted
    filled_total += fallback_filled

    # Re-assert inventory date right before save. DeskManager flows can override this
    # field during tab changes/modals and silently drift to today's date.
    date_in_value = row_value(row, "Date In")
    if date_in_value:
        if fill_field(page, "Date In", date_in_value):
            try:
                current_date_in = (page.locator("input[name='Vehicle.InventoryDate']").first.input_value() or "").strip()
                if current_date_in and current_date_in != date_in_value:
                    print(
                        f"  WARNING: Date In readback '{current_date_in}' differs from CSV '{date_in_value}'. "
                        "Reapplying once more."
                    )
                    fill_field(page, "Date In", date_in_value)
            except Exception:
                pass

    print(f"  Overall fill summary: {filled_total}/{attempted_total} fields filled")

    if attempted_total == 0:
        raise Exception("No mapped CSV fields contained values for this unit; refusing to save.")

    if STRICT_FILL and filled_total < attempted_total:
        raise Exception(
            f"Only filled {filled_total}/{attempted_total} fields; refusing to save (DESKMANAGER_STRICT_FILL=true)."
        )
    
    page.wait_for_timeout(1000)
    if DRY_RUN:
        print("  DRY RUN: skipping save (DESKMANAGER_DRY_RUN=true)")
        return
    save_vehicle(page, row)
    print("Vehicle updated successfully")


def main():
    if RUN_MODE in {"excel_update", "unit_location_note_update", "location_note_update"}:
        run_excel_location_note_updates()
        return

    csv_file = resolve_csv_file()
    if not csv_file.exists():
        print(f"Missing input file: {csv_file}")
        print(f"Looked for: {DEFAULT_CSV_FILES}")
        print("You can also set DESKMANAGER_CSV_FILE to a custom path.")
        return

    print(f"Using input file: {csv_file}")

    df = load_csv(csv_file)
    if "Stock Number" not in df.columns:
        print("CSV missing expected 'Stock Number' column.")
        print(f"Found columns: {df.columns.tolist()}")
        return

    only_stocks = parse_only_stocks(ONLY_STOCKS_RAW)
    if only_stocks:
        wanted = {s.strip().upper() for s in only_stocks}
        current_stocks = df["Stock Number"].fillna("").astype(str).str.strip()
        df = df[current_stocks.str.upper().isin(wanted)].copy()

        if df.empty:
            print(f"No rows matched DESKMANAGER_ONLY_STOCKS={only_stocks}")
            return

        found_stocks = df["Stock Number"].fillna("").astype(str).str.strip().tolist()
        missing = [s for s in only_stocks if s.upper() not in {f.upper() for f in found_stocks}]
        print(f"Filtered run: {len(df)} vehicle(s) selected from DESKMANAGER_ONLY_STOCKS")
        print(f"Selected: {found_stocks}")
        if missing:
            print(f"Not found in CSV: {missing}")

    print(f"Found {len(df)} vehicles to process")

    if DRY_RUN:
        print("*** DRY RUN MODE — fields will be filled but nothing will be saved ***")

    # ── Startup confirmation safeguard ──────────────────────────────────────
    stock_numbers = df["Stock Number"].dropna().astype(str).str.strip()
    stock_numbers = stock_numbers[stock_numbers != ""].tolist()
    preview = ", ".join(stock_numbers[:5])
    if len(stock_numbers) > 5:
        preview += f" … (+{len(stock_numbers) - 5} more)"
    action = "DRY-RUN edit" if DRY_RUN else "UPDATE"
    print(f"\nAbout to {action} {len(stock_numbers)} vehicle(s): {preview}")
    try:
        answer = input("Proceed? (y/N): ").strip().lower()
    except EOFError:
        answer = "y"  # non-interactive execution — proceed automatically
    if answer not in {"y", "yes"}:
        print("Aborted by user.")
        return
    # ────────────────────────────────────────────────────────────────────────

    with sync_playwright() as p:
        browser = None
        context = None
        page = None

        try:
            try:
                browser, context, page = create_authenticated_session(p)
                print("Login successful")
            except Exception as e:
                print(f"Could not start authenticated session: {e}")
                return

            for idx, (_, row) in enumerate(df.iterrows(), 1):
                stock_number = clean(row.get("Stock Number", ""))
                if not stock_number:
                    print(f"[{idx}/{len(df)}] Missing Stock Number, row skipped")
                    continue

                restart_attempts = 0
                max_restarts_for_stock = 2
                while True:
                    try:
                        close_active_modals(page)
                        print(f"\n[{idx}/{len(df)}] Processing {stock_number}")
                        open_vehicle_form(page, stock_number)
                        fill_vehicle_page(page, row)
                        print(f"[{idx}/{len(df)}] {stock_number} - COMPLETE")
                        break
                    except DuplicateVehicleBlocked as e:
                        print(f"[{idx}/{len(df)}] SKIPPED duplicate {stock_number}: {e}")
                        break
                    except Exception as e:
                        if is_recoverable_session_error(e) and restart_attempts < max_restarts_for_stock:
                            restart_attempts += 1
                            print(
                                f"[{idx}/{len(df)}] Recoverable session/search error; restarting session "
                                f"and retrying {stock_number} ({restart_attempts}/{max_restarts_for_stock})"
                            )
                            print(f"  DEBUG: Recovery reason: {e}")
                            try:
                                if browser is not None:
                                    browser.close()
                            except Exception:
                                pass
                            try:
                                browser, context, page = create_authenticated_session(p)
                                print("Login successful")
                                continue
                            except Exception as session_error:
                                print(
                                    f"[{idx}/{len(df)}] ERROR restarting session for {stock_number}: {session_error}"
                                )
                                break

                        print(f"[{idx}/{len(df)}] ERROR processing {stock_number}: {e}")
                        break

        finally:
            print("\nFinished processing. Closing browser...")
            try:
                if browser is not None:
                    browser.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()
