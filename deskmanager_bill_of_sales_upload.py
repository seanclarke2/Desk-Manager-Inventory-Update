import os
import re
import sys
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright

BASE_URL = "https://dm.automanager.com/Account/Login?ReturnUrl=%2fDashboard"
INVENTORY_URL = "https://dm.automanager.com/Inventory?action=10"

# Credentials must come from environment variables.
AUTOMANAGER_USERNAME = ""
AUTOMANAGER_PASSWORD = ""

# Bill of Sales folders — Sean's is primary (more accurate), Michelle's is fallback (has extras but also has
# cancelled/error files that must be skipped).
BOS_DIR_PRIMARY   = Path("/Users/seanclarke/Library/CloudStorage/OneDrive-Personal/sean's  docs/Bill of Sale")
BOS_DIR_SECONDARY = Path("/Users/seanclarke/Library/CloudStorage/OneDrive-Personal/Michelle/PC_Riverside Equipment Sales/Bill of Sales")
# Keep legacy name pointing at primary so any remaining references still work.
BILL_OF_SALES_DIR = BOS_DIR_PRIMARY

# Patterns in filenames that mark a file as invalid for use (cancelled, errors, wire instructions, etc.)
_BOS_SKIP_PATTERNS = re.compile(
    r'cancel|cancelled|canceled|credit|wrong\s*vin|wrong\s*title|unit\s*error|wire\s*instruct',
    re.IGNORECASE,
)

# Read credentials from environment variables.
AUTOMANAGER_USERNAME = os.getenv("AUTOMANAGER_USERNAME", AUTOMANAGER_USERNAME).strip()
AUTOMANAGER_PASSWORD = os.getenv("AUTOMANAGER_PASSWORD", AUTOMANAGER_PASSWORD).strip()
HEADLESS = os.getenv("PLAYWRIGHT_HEADLESS", "false").lower() in {"1", "true", "yes"}
SLOW_MO = int(os.getenv("PLAYWRIGHT_SLOW_MO", "250"))
DRY_RUN = os.getenv("DESKMANAGER_DRY_RUN", "false").lower() in {"1", "true", "yes"}
# Set this to a stock/unit number to skip all vehicles before it (leave blank to process all)
START_FROM_STOCK = os.getenv("DESKMANAGER_START_FROM", "53718").strip()

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(line_buffering=True)
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(line_buffering=True)


def extract_invoice_number_from_filename(filename):
    """
    Extract invoice number from bill of sale filename.
    Examples:
    - Bill of Sale_CRST_2024-04-16_42426 _Final.pdf -> 42426
    - Bill of Sale_CRST_2025-04-29_42526-Signed.pdf -> 42526
    """
    match = re.search(r'_(\d{5,6})(?:\s|_|-|\.)', filename)
    if match:
        return match.group(1)
    return None


def find_bill_of_sale(invoice_number, invoice_date=None):
    """
    Find the best bill of sale PDF file for an invoice.
    Searches Sean's folder first (primary/more accurate), then falls back to Michelle's
    folder for any files Sean is missing — skipping cancelled, error, and credit files.
    Prefers _Signed.pdf, then _Final.pdf, then others.

    Args:
        invoice_number: The invoice number to search for
        invoice_date: Optional invoice date for better matching

    Returns:
        Path to the best matching PDF file, or None if not found
    """

    def _is_valid(f: Path) -> bool:
        """Return True if this file should be considered (not a skip pattern, not a dir)."""
        return f.is_file() and f.suffix.lower() in {'.pdf', '.tif', '.tiff'} \
               and not _BOS_SKIP_PATTERNS.search(f.name)

    def _sort_key(f: Path):
        name = f.name.lower()
        if '_signed' in name:
            return (0, f.stat().st_mtime)
        elif '_final' in name:
            return (1, f.stat().st_mtime)
        else:
            return (2, f.stat().st_mtime)

    def _search_dir(folder: Path):
        if not folder.exists():
            return []
        return [
            f for f in folder.iterdir()
            if _is_valid(f) and invoice_number in f.name
        ]

    # Primary: Sean's folder
    candidates = _search_dir(BOS_DIR_PRIMARY)

    # Fallback: Michelle's folder for anything Sean doesn't have
    if not candidates:
        candidates = _search_dir(BOS_DIR_SECONDARY)

    if not candidates:
        return None

    candidates.sort(key=_sort_key)
    return candidates[0]


def login(page):
    """Login to Desk Manager"""
    print(f"Navigating to login page...")
    page.goto(BASE_URL, wait_until="networkidle")
    
    # Wait for login form to appear
    page.wait_for_selector('input[type="text"], input[type="email"]', timeout=10000)
    
    # Find and fill username
    username_input = page.locator('input[type="text"], input[type="email"]').first
    password_input = page.locator('input[type="password"]').first
    
    print(f"Filling login credentials...")
    username_input.fill(AUTOMANAGER_USERNAME)
    page.wait_for_timeout(500)
    password_input.fill(AUTOMANAGER_PASSWORD)
    page.wait_for_timeout(500)
    
    # Click login button
    login_button = page.locator('button:has-text("Log in"), button:has-text("LOGIN"), input[type="submit"]').first
    login_button.click()
    
    # Wait for dashboard to load
    page.wait_for_url("**/Dashboard**", timeout=30000)
    print("✓ Login successful")


def navigate_to_inventory(page):
    """Navigate to Inventory section"""
    print(f"Navigating to Inventory...")
    page.goto(INVENTORY_URL, wait_until="networkidle")
    page.wait_for_timeout(3000)
    print("✓ On Inventory page")


def get_vehicle_hrefs_from_inventory(page):
    """
    Get all vehicle detail page URLs from all inventory pages.
    Handles pagination automatically.
    """
    seen = set()
    vehicle_hrefs = []
    page_num = 1

    while True:
        page.wait_for_timeout(3000)
        all_links = page.locator('a[href]').all()

        before = len(seen)
        for link in all_links:
            try:
                href = (link.get_attribute("href") or "").strip()
                if not href:
                    continue
                href = href.split("#")[0]
                href_lower = href.lower()
                if not any(t in href_lower for t in ["/inventory/details/", "/inventory/edit/"]):
                    continue
                if href in seen:
                    continue
                seen.add(href)
                if not href.startswith("http"):
                    href = f"https://dm.automanager.com{href}"
                vehicle_hrefs.append(href)
            except Exception:
                pass

        added = len(seen) - before
        print(f"  Page {page_num}: +{added} vehicles (total so far: {len(vehicle_hrefs)})")

        # Try to click the Next page button (exact text match to avoid matching "Upcoming Payments - Next week")
        next_btn = None
        for next_sel in [
            'li:not([style*="display:none"]):not([style*="display: none"]) > a:text-is("Next")',
            'a:text-is("Next")',
            'li.active + li > a',
            'a[aria-label="Next"]',
        ]:
            try:
                el = page.locator(next_sel).last  # use last since there may be two (desktop/mobile)
                if el.count() > 0 and el.is_visible() and el.is_enabled():
                    next_btn = el
                    break
            except Exception:
                pass
        if next_btn:
            next_btn.click()
            page_num += 1
            page.wait_for_timeout(3000)
        else:
            break

    print(f"Found {len(vehicle_hrefs)} unique vehicle pages across {page_num} page(s)")
    return vehicle_hrefs


def read_label_field(page, label_text):
    """
    Find an input field by its nearby label text and return its value.
    """
    try:
        # Try finding label element, then get the associated input
        label = page.get_by_text(label_text, exact=True).first
        if not label.is_visible():
            return None
        # Try for= attribute on label
        for_id = label.evaluate("el => el.htmlFor || el.getAttribute('for') || ''")
        if for_id:
            inp = page.locator(f"#{for_id}")
            if inp.count() > 0:
                return inp.first.input_value().strip() or None
        # Try sibling/nearby input
        val = label.evaluate("""
            el => {
                let sib = el.nextElementSibling;
                while (sib) {
                    if (sib.tagName === 'INPUT' || sib.tagName === 'SELECT') return sib.value;
                    let inp = sib.querySelector('input, select');
                    if (inp) return inp.value;
                    sib = sib.nextElementSibling;
                }
                // Try parent row
                let row = el.closest('tr, div');
                if (row) {
                    let inp = row.querySelector('input, select');
                    if (inp) return inp.value;
                }
                return '';
            }
        """)
        return val.strip() if val and val.strip() else None
    except Exception:
        return None


def get_invoice_info_from_vehicle(page):
    """
    Extract Bill Of Sale Number and Date from the vehicle's "Other" tab.
    Falls back to Invoice No. on the "Purchase Info" tab.
    Returns: {'number': '42426', 'date': '2024-04-16'} or None
    """
    invoice_number = None
    invoice_date = None

    # --- Try "Other" tab first (has Bill Of Sale Number + Bill Of Sale Date) ---
    try:
        for sel in ["a:has-text('Other')", "button:has-text('Other')", "li:has-text('Other')"]:
            try:
                tab = page.locator(sel).first
                if tab.is_visible():
                    tab.click(timeout=3000)
                    page.wait_for_timeout(1200)
                    break
            except Exception:
                pass
        else:
            page.get_by_text("Other", exact=True).click(timeout=4000)
            page.wait_for_timeout(1200)

        invoice_number = read_label_field(page, "Bill Of Sale Number")
        invoice_date = read_label_field(page, "Bill Of Sale Date")
    except Exception:
        pass

    # --- Fall back to "Purchase Info" tab for Invoice No. ---
    if not invoice_number:
        try:
            for sel in ["a:has-text('Purchase Info')", "button:has-text('Purchase Info')", "li:has-text('Purchase Info')"]:
                try:
                    tab = page.locator(sel).first
                    if tab.is_visible():
                        tab.click(timeout=3000)
                        page.wait_for_timeout(1200)
                        break
                except Exception:
                    pass
            invoice_number = read_label_field(page, "Invoice No.")
        except Exception:
            pass

    if invoice_number:
        return {
            'number': str(invoice_number).strip(),
            'date': str(invoice_date).strip() if invoice_date else None
        }

    return None


def upload_bill_of_sale_to_attachments(page, bill_of_sale_path):
    """
    Upload a bill of sale file to the attachments section of current vehicle.
    """
    try:
        file_path = Path(bill_of_sale_path)
        print(f"  - Uploading: {file_path.name}")

        # Click Attachments tab - try multiple selectors
        clicked_tab = False
        for selector in [
            'a:has-text("Attachment")',
            'button:has-text("Attachment")',
            '[role="tab"]:has-text("Attachment")',
            'li:has-text("Attachment") > a',
            '.nav-link:has-text("Attachment")',
            'a[href*="Attachment" i]',
        ]:
            try:
                el = page.locator(selector).first
                if el.count() > 0 and el.is_visible():
                    el.click()
                    page.wait_for_timeout(2000)
                    clicked_tab = True
                    print(f"  - Clicked Attachments tab via: {selector}")
                    break
            except Exception:
                pass

        if not clicked_tab:
            # Print available tabs for debugging
            tabs_text = page.evaluate("""
                () => Array.from(document.querySelectorAll('a, [role="tab"], .nav-link'))
                    .filter(el => el.offsetParent !== null)
                    .map(el => el.textContent.trim())
                    .filter(t => t.length > 0 && t.length < 40)
            """)
            print(f"  ! Attachments tab not found. Visible tabs: {tabs_text[:10]}")

        # Strategy 1: Look for hidden file input directly and set files (bypasses button click)
        file_inputs = page.locator('input[type="file"]').all()
        if file_inputs:
            # Use the first file input, even if hidden
            fi = file_inputs[0]
            fi.set_input_files(str(bill_of_sale_path))
            page.wait_for_timeout(2000)
            # Look for a save/submit/upload button to confirm
            for btn_sel in [
                'button:has-text("Save")',
                'button:has-text("Upload")',
                'button:has-text("Submit")',
                'button:has-text("OK")',
                'button:has-text("Confirm")',
                'input[type="submit"]',
            ]:
                try:
                    btn = page.locator(btn_sel).last
                    if btn.count() > 0 and btn.is_visible():
                        btn.click()
                        page.wait_for_timeout(2000)
                        print(f"  - Confirmed with: {btn_sel}")
                        break
                except Exception:
                    pass
            print(f"  ✓ File uploaded: {file_path.name}")
            return True

        # Strategy 2: Click an upload button first, then look for file input
        upload_button = None
        for btn_sel in [
            'button:has-text("Upload")',
            'button:has-text("Add File")',
            'button:has-text("Add Attachment")',
            'label:has-text("Upload")',
            'a:has-text("Upload")',
            'a:has-text("Add File")',
        ]:
            try:
                el = page.locator(btn_sel).first
                if el.count() > 0 and el.is_visible():
                    upload_button = el
                    break
            except Exception:
                pass

        if upload_button is None:
            # Print visible buttons for debugging
            btns_text = page.evaluate("""
                () => Array.from(document.querySelectorAll('button, a, label'))
                    .filter(el => el.offsetParent !== null)
                    .map(el => el.textContent.trim())
                    .filter(t => t.length > 0 && t.length < 40)
            """)
            print(f"  ✗ Could not find upload button. Visible clickables: {btns_text[:15]}")
            return False

        # Use expect_file_chooser to handle the file picker dialog
        with page.expect_file_chooser() as fc_info:
            upload_button.click()
        file_chooser = fc_info.value
        file_chooser.set_files(str(bill_of_sale_path))
        page.wait_for_timeout(2000)

        # Confirm/save
        for btn_sel in [
            'button:has-text("Save")',
            'button:has-text("Upload")',
            'button:has-text("Submit")',
            'button:has-text("OK")',
        ]:
            try:
                btn = page.locator(btn_sel).last
                if btn.count() > 0 and btn.is_visible():
                    btn.click()
                    page.wait_for_timeout(2000)
                    break
            except Exception:
                pass

        print(f"  ✓ File uploaded: {file_path.name}")
        return True

    except Exception as e:
        print(f"  ✗ Error uploading file: {str(e)}")
        return False


def main():
    """Main function"""
    print("\n" + "="*60)
    print("DESK MANAGER - BILL OF SALES UPLOADER (FROM VEHICLES)")
    print("="*60)
    
    if DRY_RUN:
        print("[DRY RUN MODE] Script will run but not actually upload files.")
    
    # Start browser
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        
        try:
            # Login
            login(page)
            
            # Navigate to inventory
            navigate_to_inventory(page)
            
            # Get list of vehicles
            vehicle_hrefs = get_vehicle_hrefs_from_inventory(page)

            if len(vehicle_hrefs) == 0:
                print("No vehicles found on inventory page")
                return

            print(f"\nFound {len(vehicle_hrefs)} vehicles to process\n")

            uploaded_count = 0
            skipped_count = 0
            start_found = not bool(START_FROM_STOCK)  # If no start stock set, begin immediately

            # Process each vehicle
            for i, href in enumerate(vehicle_hrefs, 1):
                try:
                    # Navigate to vehicle page
                    if not href.startswith("http"):
                        href = f"https://dm.automanager.com{href}"
                    
                    page.goto(href, wait_until="domcontentloaded")
                    page.wait_for_timeout(2000)

                    # Check if we've reached the start stock number
                    if not start_found:
                        # Read the stock number from the page
                        try:
                            stock_input = page.locator('input[id*="stock" i], input[name*="stock" i]').first
                            stock_val = stock_input.input_value().strip() if stock_input.count() > 0 else ""
                            if not stock_val:
                                # Try reading from page title or header
                                stock_val = page.locator('h1, h2, .stock-number, [data-field*="stock"]').first.text_content().strip()
                        except Exception:
                            stock_val = ""
                        if START_FROM_STOCK in stock_val or stock_val in START_FROM_STOCK:
                            start_found = True
                            print(f"[{i}/{len(vehicle_hrefs)}] Starting here: {stock_val}")
                        else:
                            print(f"[{i}/{len(vehicle_hrefs)}] Skipping (before {START_FROM_STOCK}): {stock_val or href.split('/')[-1][:8]}")
                            continue
                    else:
                        print(f"[{i}/{len(vehicle_hrefs)}] {href.split('/')[-1][:8]}...")

                    # Get invoice info from "Other" tab
                    invoice_info = get_invoice_info_from_vehicle(page)

                    if not invoice_info or not invoice_info.get('number'):
                        print(f"  - No invoice number, skipping")
                        skipped_count += 1
                        continue

                    invoice_number = invoice_info['number']
                    invoice_date = invoice_info.get('date')
                    print(f"  - Invoice: {invoice_number} ({invoice_date or 'no date'})")

                    # Find matching bill of sale file
                    bill_of_sale = find_bill_of_sale(invoice_number, invoice_date)

                    if not bill_of_sale:
                        print(f"  - No matching bill of sale file, skipping")
                        skipped_count += 1
                        continue

                    print(f"  ✓ Found: {bill_of_sale.name}")

                    if not DRY_RUN:
                        success = upload_bill_of_sale_to_attachments(page, bill_of_sale)
                        if success:
                            uploaded_count += 1
                        else:
                            skipped_count += 1
                    else:
                        print(f"  [DRY RUN] Would upload: {bill_of_sale.name}")
                        uploaded_count += 1

                except Exception as e:
                    print(f"  - Skipping ({str(e)[:80]})")
                    skipped_count += 1
                    continue
            
            print(f"\n{'='*60}")
            print(f"COMPLETE: {uploaded_count} uploaded, {skipped_count} skipped")
            print("="*60)
            
        except Exception as e:
            print(f"\nERROR: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            browser.close()


if __name__ == "__main__":
    main()
