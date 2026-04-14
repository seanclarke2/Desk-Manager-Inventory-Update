import os
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE_URL = "https://dm.automanager.com/Account/Login?ReturnUrl=%2fDashboard"
INVENTORY_URL = "https://dm.automanager.com/Inventory?action=10"
CSV_FILE = str(Path(__file__).with_name("vehicle_updates.csv"))

# Put your real login here locally in VS Code if you want auto-login every time.
AUTOMANAGER_USERNAME = "deiseldeals2@gmail.com"
AUTOMANAGER_PASSWORD = "2$EanClarke2"

# Optional: environment variables override the hardcoded values above
AUTOMANAGER_USERNAME = os.getenv("AUTOMANAGER_USERNAME", AUTOMANAGER_USERNAME).strip()
AUTOMANAGER_PASSWORD = os.getenv("AUTOMANAGER_PASSWORD", AUTOMANAGER_PASSWORD).strip()


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


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


def login(page):
    page.goto(BASE_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2500)

    if not AUTOMANAGER_USERNAME or not AUTOMANAGER_PASSWORD:
        raise Exception("Missing username or password in script.")

    username_selectors = [
        "input[name='Email']",
        "input[name='UserName']",
        "input[name='Username']",
        "input[type='email']",
        "input[id*='email']",
        "input[id*='user']",
        "input[placeholder*='Email']",
        "input[placeholder*='User']",
        "input[type='text']",
    ]

    password_selectors = [
        "input[name='Password']",
        "input[type='password']",
        "input[id*='password']",
        "input[placeholder*='Password']",
    ]

    login_button_selectors = [
        "button[type='submit']",
        "input[type='submit']",
        "button:has-text('Log In')",
        "button:has-text('Login')",
        "input[value='Log In']",
        "input[value='Login']",
    ]

    username_ok = first_visible_fill(page, username_selectors, AUTOMANAGER_USERNAME)
    password_ok = first_visible_fill(page, password_selectors, AUTOMANAGER_PASSWORD)

    if not username_ok:
        raise Exception("Could not find username field.")
    if not password_ok:
        raise Exception("Could not find password field.")

    clicked = first_visible_click(page, login_button_selectors)
    if not clicked:
        page.keyboard.press("Enter")

    page.wait_for_timeout(5000)
    page.goto(INVENTORY_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)


def search_inventory(page, stock_number):
    page.goto(INVENTORY_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)

    search_selectors = [
        "input[placeholder*='Search']",
        "input[name*='search']",
        "input[id*='search']",
        "input[type='search']",
        "input.form-control",
    ]

    search_ok = first_visible_fill(page, search_selectors, stock_number)

    if not search_ok:
        loc = page.locator("input")
        for i in range(loc.count()):
            try:
                item = loc.nth(i)
                input_type = (item.get_attribute("type") or "").lower()
                if input_type in ["hidden", "password", "checkbox", "radio", "submit"]:
                    continue
                if item.is_visible() and item.is_enabled():
                    item.click(timeout=1000)
                    item.fill("")
                    item.fill(str(stock_number))
                    search_ok = True
                    break
            except Exception:
                pass

    if not search_ok:
        raise Exception("Could not find inventory search box.")

    page.keyboard.press("Enter")
    page.wait_for_timeout(3000)

    try:
        first_visible_click(
            page,
            [
                "button:has-text('Search')",
                "input[value='Search']",
                "button[title*='Search']",
            ],
            timeout=2000,
        )
        page.wait_for_timeout(2000)
    except Exception:
        pass


def open_matching_unit(page, stock_number):
    page.wait_for_timeout(2500)
    
    stock_texts = [
        f"Stock #: {stock_number}",
        f"Stock#:{stock_number}",
        f"Stock #:{stock_number}",
        str(stock_number),
    ]
    
    stock_locator = None
    for txt in stock_texts:
        try:
            loc = page.get_by_text(txt, exact=False)
            if loc.count() > 0 and loc.first.is_visible():
                stock_locator = loc.first
                break
        except Exception:
            pass
    
    if stock_locator is None:
        raise Exception(f"Could not find result row for stock number {stock_number}.")
    
    stock_locator.scroll_into_view_if_needed()
    page.wait_for_timeout(1000)
    
    containers = [
        "xpath=ancestor::tr[1]",
        "xpath=ancestor::div[contains(@class,'row')][1]",
        "xpath=ancestor::div[contains(@class,'card')][1]",
        "xpath=ancestor::div[contains(@class,'item')][1]",
        "xpath=ancestor::div[contains(@class,'result')][1]",
        "xpath=ancestor::li[1]",
        "xpath=ancestor::div[1]",
    ]
    
    for container_xpath in containers:
        try:
            container = stock_locator.locator(container_xpath)
            
            if container.count() == 0:
                continue
            
            links = container.locator("a")
            
            for i in range(links.count()):
                try:
                    link = links.nth(i)
                    
                    if not link.is_visible() or not link.is_enabled():
                        continue
                    
                    text = (link.inner_text() or "").strip()
                    href = (link.get_attribute("href") or "").strip().lower()
                    
                    if not text:
                        continue
                    



def maybe_click_details_tab(page):
    try:
        page.get_by_text("Details", exact=True).click(timeout=3000)
        page.wait_for_timeout(1000)
    except Exception:
        pass


def get_visible_inputs(page):
    visible = []
    loc = page.locator("input")
    for i in range(loc.count()):
        item = loc.nth(i)
        try:
            input_type = (item.get_attribute("type") or "").lower()
            if input_type == "hidden":
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
            if item.is_visible() and item.is_enabled():
                visible.append(item)
        except Exception:
            pass
    return visible


def fill_input_by_index(page, index, value):
    if not value:
        return
    inputs = get_visible_inputs(page)
    if index >= len(inputs):
        raise Exception(
            f"Visible input index {index} not found. Only found {len(inputs)} visible inputs."
        )
    loc = inputs[index]
    loc.click()
    loc.fill("")
    loc.fill(str(value))


def fill_select_by_index(page, index, value):
    if not value:
        return
    selects = get_visible_selects(page)
    if index >= len(selects):
        raise Exception(
            f"Visible select index {index} not found. Only found {len(selects)} visible selects."
        )
    loc = selects[index]
    loc.select_option(label=str(value))


def fill_vehicle_page(page, row):
    maybe_click_details_tab(page)

    fill_input_by_index(page, 0, clean(row.get("Stock Number", "")))
    fill_select_by_index(page, 0, clean(row.get("Status", "")))
    fill_select_by_index(page, 1, clean(row.get("Sub Status", "")))
    fill_select_by_index(page, 2, clean(row.get("City", "")))
    fill_input_by_index(page, 1, clean(row.get("Bill Of Sale Date", "")))

    fill_input_by_index(page, 2, clean(row.get("VIN", "")))
    fill_select_by_index(page, 3, clean(row.get("Vehicle Type", "")))
    fill_select_by_index(page, 4, clean(row.get("New / Used", "")))
    fill_select_by_index(page, 5, clean(row.get("Condition", "")))
    fill_select_by_index(page, 6, clean(row.get("Trailer Type", "")))

    fill_input_by_index(page, 3, clean(row.get("year", "")))
    fill_input_by_index(page, 4, clean(row.get("Make", "")))
    fill_input_by_index(page, 5, clean(row.get("Model", "")))
    fill_input_by_index(page, 6, clean(row.get("Series", "")))
    fill_input_by_index(page, 7, clean(row.get("Body Style", "")))
    fill_input_by_index(page, 8, clean(row.get("Engine", "")))
    fill_select_by_index(page, 7, clean(row.get("Drivetrain", "")))
    fill_select_by_index(page, 8, clean(row.get("Overdrive", "")))

    fill_input_by_index(page, 9, clean(row.get("Length", "")))
    fill_select_by_index(page, 9, clean(row.get("Exterior Color", "")))
    fill_select_by_index(page, 10, clean(row.get("Interior Color", "")))
    fill_input_by_index(page, 10, clean(row.get("Weight", "")))
    fill_input_by_index(page, 11, clean(row.get("Height", "")))
    fill_select_by_index(page, 11, clean(row.get("Axels", "")))
    fill_select_by_index(page, 12, clean(row.get("Suspension", "")))
    fill_select_by_index(page, 13, clean(row.get("Tire size", "")))

    fill_input_by_index(page, 12, clean(row.get("Gross Weight", "")))
    fill_input_by_index(page, 13, clean(row.get("Unladen Weight", "")))
    fill_input_by_index(page, 14, clean(row.get("License Plate", "")))
    fill_input_by_index(page, 15, clean(row.get("Last Registered", "")))
    fill_input_by_index(page, 16, clean(row.get("State Registered", "")))
    fill_input_by_index(page, 17, clean(row.get("Bill of sale Date", "")))
    fill_input_by_index(page, 18, clean(row.get("Carrying Capacity", "")))
    fill_input_by_index(page, 19, clean(row.get("Pass Time Serial No.", "")))

    try:
        page.get_by_text("Save", exact=True).click(timeout=3000)
    except Exception:
        first_visible_click(page, ["button:has-text('Save')", "input[value='Save']"])

    page.wait_for_timeout(2000)


def main():
    if not Path(CSV_FILE).exists():
        print(f"Missing {CSV_FILE}")
        return

    df = pd.read_csv(CSV_FILE)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=200)
        context = browser.new_context()
        page = context.new_page()

        login(page)

        for _, row in df.iterrows():
            stock_number = clean(row.get("Stock Number", ""))
            if not stock_number:
                print("Missing Stock Number, row skipped")
                continue

            print(f"Processing {stock_number}")
            search_inventory(page, stock_number)
            open_matching_unit(page, stock_number)
            fill_vehicle_page(page, row)

        input("\nPress Enter to close...")
        browser.close()


if __name__ == "__main__":
    main()
