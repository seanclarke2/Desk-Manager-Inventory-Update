#!/usr/bin/env python3
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from playwright.sync_api import sync_playwright

import deskmanager_vehicle_fill as dvf
from deskmanager_bill_of_sales_upload import upload_bill_of_sale_to_attachments

DESKTOP = Path.home() / "Desktop"
INPUT_CSV = Path(os.getenv("DESKMANAGER_INPUT_CSV", str(DESKTOP / "dm_import_master_cleaned.csv")))
OUTPUT_DIR = Path(os.getenv("DESKMANAGER_OUTPUT_DIR", str(DESKTOP)))
START_FROM = os.getenv("DESKMANAGER_START_FROM", "").strip()
DRY_RUN = os.getenv("DESKMANAGER_DRY_RUN", "false").lower() in {"1", "true", "yes"}

INVOICE_DIR_PRIMARY = Path(
    "/Users/seanclarke/Library/CloudStorage/OneDrive-Personal/Michelle/PC_Riverside Equipment Sales/Invoices"
)
INVOICE_DIR_SECONDARY = Path(
    "/Users/seanclarke/Library/CloudStorage/OneDrive-Personal/sean's  docs/Bill of Sale"
)
INVOICE_DIR_TERTIARY = Path(
    "/Users/seanclarke/Library/CloudStorage/OneDrive-Personal/Michelle/PC_Riverside Equipment Sales/Bill of Sales"
)

SUPPORTED_DOC_EXT = {
    ".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".webp", ".bmp", ".doc", ".docx", ".xlsx", ".xls"
}

SKIP_PATTERN = re.compile(
    r"cancel|cancelled|canceled|credit|wrong\s*vin|wrong\s*title|unit\s*error|wire\s*instruct",
    re.IGNORECASE,
)


def _extract_invoice_token(*values: object) -> str:
    for value in values:
        s = str(value or "").strip()
        if not s:
            continue
        m = re.search(r"\b(\d{4,7})\b", s)
        if m:
            return m.group(1)
    return ""


def _score_file(path: Path, root_rank: int) -> Tuple[int, int, float]:
    name = path.name.lower()
    quality = 5
    if "customer" in name:
        quality = 0
    elif "final" in name:
        quality = 1
    elif "signed" in name:
        quality = 2
    elif path.suffix.lower() == ".pdf":
        quality = 3
    else:
        quality = 4
    return (root_rank, quality, -path.stat().st_mtime)


def build_invoice_file_map() -> Dict[str, Path]:
    invoice_map: Dict[str, Path] = {}
    score_map: Dict[str, Tuple[int, int, float]] = {}

    roots = [
        (INVOICE_DIR_PRIMARY, 0),
        (INVOICE_DIR_SECONDARY, 1),
        (INVOICE_DIR_TERTIARY, 2),
    ]

    for root, root_rank in roots:
        if not root.exists():
            continue
        for f in root.rglob("*"):
            if not f.is_file():
                continue
            if f.suffix.lower() not in SUPPORTED_DOC_EXT:
                continue
            if SKIP_PATTERN.search(f.name):
                continue
            token = _extract_invoice_token(f.name)
            if not token:
                continue
            score = _score_file(f, root_rank)
            if token not in invoice_map or score < score_map[token]:
                invoice_map[token] = f
                score_map[token] = score

    return invoice_map


def _open_attachments_tab(page) -> None:
    selectors = [
        'a:has-text("Attachment")',
        'button:has-text("Attachment")',
        '[role="tab"]:has-text("Attachment")',
        'li:has-text("Attachment") > a',
        '.nav-link:has-text("Attachment")',
        'a[href*="Attachment" i]',
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible():
                loc.click(timeout=3000)
                page.wait_for_timeout(1200)
                return
        except Exception:
            pass


def _upload_via_browse_button(page, file_path: Path) -> bool:
    """Upload on Attachments via Browse click first, then file-input fallback."""
    try:
        dvf.close_sticky_note_popup_if_present(page)
        dvf.close_active_modals(page)
        _open_attachments_tab(page)
        dvf.close_sticky_note_popup_if_present(page)
        dvf.close_active_modals(page)

        # User workflow requires the blue Browse button at bottom-left.
        browse_candidates = page.locator(
            "button:has-text('Browse'), a:has-text('Browse'), label:has-text('Browse'), .btn:has-text('Browse'), [role='button']:has-text('Browse')"
        )

        chosen_idx = None
        chosen_score = None
        for i in range(browse_candidates.count()):
            try:
                candidate = browse_candidates.nth(i)
                if not candidate.is_visible() or not candidate.is_enabled():
                    continue
                box = candidate.bounding_box()
                if not box:
                    continue
                cls = (candidate.get_attribute("class") or "").lower()
                is_blue = any(tok in cls for tok in ["primary", "info", "blue", "btn-primary", "btn-info"])
                # Prefer blue, then lower on screen (larger y), then further left (smaller x).
                score = (1 if is_blue else 0, box["y"], -box["x"])
                if chosen_score is None or score > chosen_score:
                    chosen_score = score
                    chosen_idx = i
            except Exception:
                continue

        browse_clicked = False
        file_selected = False
        if chosen_idx is not None:
            browse_el = browse_candidates.nth(chosen_idx)
            try:
                with page.expect_file_chooser(timeout=5000) as fc_info:
                    browse_el.click(timeout=4000)
                chooser = fc_info.value
                chooser.set_files(str(file_path))
                page.wait_for_timeout(1200)
                print(f"  - Browse selected file: {file_path.name}")
                browse_clicked = True
                file_selected = True
            except Exception:
                # Some Browse controls are labels for hidden input and won't raise chooser event.
                try:
                    browse_el.click(timeout=3000)
                    page.wait_for_timeout(600)
                    browse_clicked = True
                except Exception:
                    browse_clicked = False

        # Fallback for drag/drop or hidden file inputs when no chooser opens.
        if not browse_clicked:
            return False

        if not file_selected:
            file_inputs = page.locator("input[type='file']")
            if file_inputs.count() > 0:
                try:
                    file_inputs.first.set_input_files(str(file_path))
                    page.wait_for_timeout(1200)
                    print(f"  - File input attached (drag/drop fallback): {file_path.name}")
                except Exception:
                    pass

        # Confirm upload/save if button appears.
        for sel in [
            "button:has-text('Save')",
            "button:has-text('Upload')",
            "button:has-text('Submit')",
            "button:has-text('OK')",
            "input[type='submit']",
        ]:
            try:
                btn = page.locator(sel).last
                if btn.count() > 0 and btn.is_visible() and btn.is_enabled():
                    btn.click(timeout=3000)
                    page.wait_for_timeout(1200)
                    print(f"  - Confirmed attachment with: {sel}")
                    break
            except Exception:
                pass

        return True
    except Exception:
        return False


def _upload_via_drag_drop_path(page, file_path: Path) -> bool:
    """Upload by targeting the attachment drop area/file input directly from file path."""
    try:
        dvf.close_sticky_note_popup_if_present(page)
        dvf.close_active_modals(page)
        _open_attachments_tab(page)
        dvf.close_sticky_note_popup_if_present(page)
        dvf.close_active_modals(page)

        # Prefer upload/dropzone file inputs near the bottom-left of the screen.
        inputs = page.locator("input[type='file']")
        if inputs.count() == 0:
            return False

        picked = None
        picked_score = None
        for i in range(inputs.count()):
            try:
                candidate = inputs.nth(i)
                box = candidate.bounding_box() or {"x": 99999, "y": -1}
                cls = (candidate.get_attribute("class") or "").lower()
                name = (candidate.get_attribute("name") or "").lower()
                input_id = (candidate.get_attribute("id") or "").lower()
                marker = f"{cls} {name} {input_id}"
                is_uploadish = any(tok in marker for tok in ["upload", "attach", "drop", "file", "qq", "fineuploader"])
                score = (1 if is_uploadish else 0, box["y"], -box["x"])
                if picked_score is None or score > picked_score:
                    picked_score = score
                    picked = candidate
            except Exception:
                continue

        if picked is None:
            return False

        picked.set_input_files(str(file_path))
        page.wait_for_timeout(1400)
        print(f"  - Drag/drop path attached file: {file_path.name}")

        for sel in [
            "button:has-text('Save')",
            "button:has-text('Upload')",
            "button:has-text('Submit')",
            "button:has-text('OK')",
            "input[type='submit']",
        ]:
            try:
                btn = page.locator(sel).last
                if btn.count() > 0 and btn.is_visible() and btn.is_enabled():
                    btn.click(timeout=3000)
                    page.wait_for_timeout(1200)
                    print(f"  - Confirmed drag/drop upload with: {sel}")
                    break
            except Exception:
                pass

        return True
    except Exception:
        return False


def _attachment_already_present(page, filename: str) -> bool:
    try:
        # Sticky-note popups can block tab clicks; dismiss before checking attachments.
        dvf.close_sticky_note_popup_if_present(page)
        dvf.close_active_modals(page)
        _open_attachments_tab(page)
        dvf.close_sticky_note_popup_if_present(page)
        dvf.close_active_modals(page)
        content = page.content().lower()
        return filename.lower() in content
    except Exception:
        return False


def _clean_text(value: object) -> str:
    return str(value or "").strip()


def _update_sold_information(page, row: pd.Series) -> Tuple[bool, str]:
    """Update sold fields (including Sub Status=Sold) and save the vehicle."""
    try:
        row_for_save = row.copy()
        row_for_save["Status"] = "Sold"
        row_for_save["Sub Status"] = "Sold"

        dvf.close_sticky_note_popup_if_present(page)
        dvf.close_active_modals(page)

        # Top-level status fields.
        dvf.fill_field(page, "Status", "Sold")
        dvf.fill_field(page, "Sub Status", "Sold")

        # Sold details are on Other tab in this workflow.
        dvf.click_tab(page, "Other")
        for label, key in [
            ("Sold Date", "Sold Date"),
            ("Sold To", "Sold To"),
            ("Sold Price", "Sold Price"),
            ("Bill Of Sale Date", "Bill Of Sale Date"),
            ("Bill Of Sale Number", "Bill Of Sale Number"),
        ]:
            value = _clean_text(row.get(key, ""))
            if value:
                dvf.fill_field(page, label, value)

        # Invoice number commonly sits on Purchase Info.
        dvf.click_tab(page, "Purchase Info")
        invoice_no = _clean_text(row.get("Invoice No.", ""))
        if invoice_no:
            dvf.fill_field(page, "Invoice No.", invoice_no)

        if DRY_RUN:
            print("  ~ sold fields updated (DRY RUN, no save)")
            return True, "DRY RUN sold update"

        dvf.save_vehicle(page, row_for_save)
        print("  + sold fields saved (Status/Sub Status/Other tab fields)")
        return True, "Sold fields saved"
    except Exception as e:
        return False, str(e)


def load_sold_rows(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip() for c in df.columns]
    if "Status" not in df.columns:
        raise Exception("Input CSV missing 'Status' column")

    sold = df[df["Status"].astype(str).str.strip().str.lower() == "sold"].copy()
    return sold


def main() -> None:
    if not INPUT_CSV.exists():
        raise Exception(f"Input CSV not found: {INPUT_CSV}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    report_path = OUTPUT_DIR / "dm_sold_invoice_upload_report.xlsx"

    sold_df = load_sold_rows(INPUT_CSV)
    invoice_map = build_invoice_file_map()

    print(f"[Sold Invoices] Sold rows in CSV: {len(sold_df)}")
    print(f"[Sold Invoices] Invoice index files: {len(invoice_map)}")
    if DRY_RUN:
        print("[Sold Invoices] DRY RUN mode ON (no uploads)")

    results: List[dict] = []
    start_found = not bool(START_FROM)

    with sync_playwright() as p:
        browser, context, page = dvf.create_authenticated_session(p)
        print("[Sold Invoices] Authenticated. Starting upload loop...")

        try:
            for _, row in sold_df.iterrows():
                stock = str(row.get("Unit", "")).strip()
                if not stock:
                    continue

                if not start_found:
                    if stock == START_FROM:
                        start_found = True
                    else:
                        results.append({
                            "Unit": stock,
                            "Invoice Token": "",
                            "File": "",
                            "Status": "Skipped",
                            "Detail": "Before START_FROM",
                        })
                        continue

                try:
                    dvf.close_sticky_note_popup_if_present(page)
                    dvf.close_active_modals(page)
                    dvf.open_vehicle_form(page, stock)
                    dvf.close_sticky_note_popup_if_present(page)
                    dvf.close_active_modals(page)
                except Exception as e:
                    # One quick recovery attempt after dismissing any blocking popup.
                    try:
                        dvf.close_sticky_note_popup_if_present(page)
                        dvf.close_active_modals(page)
                        dvf.open_vehicle_form(page, stock)
                        dvf.close_sticky_note_popup_if_present(page)
                        dvf.close_active_modals(page)
                    except Exception:
                        pass

                try:
                    dvf.close_sticky_note_popup_if_present(page)
                    dvf.close_active_modals(page)
                except Exception:
                    pass

                try:
                    # Verify we are on a vehicle page after potential retry.
                    dvf.wait_for_ready(page, timeout=12000)
                except Exception as e:
                    results.append({
                        "Unit": stock,
                        "Invoice Token": "",
                        "File": "",
                        "Status": "Failed",
                        "Detail": f"Could not open vehicle form: {e}",
                    })
                    print(f"  - {stock}: could not open vehicle page")
                    continue

                # Keep sold info synchronized before attachment work.
                sold_ok, sold_detail = _update_sold_information(page, row)
                if not sold_ok:
                    results.append({
                        "Unit": stock,
                        "Invoice Token": "",
                        "File": "",
                        "Status": "Failed",
                        "Detail": f"Sold update failed: {sold_detail}",
                    })
                    print(f"  - {stock}: sold update failed")
                    continue

                token = _extract_invoice_token(row.get("Invoice No.", ""), row.get("Bill Of Sale Number", ""))
                if not token:
                    results.append({
                        "Unit": stock,
                        "Invoice Token": "",
                        "File": "",
                        "Status": "Failed",
                        "Detail": "Sold updated; no Invoice No. / Bill Of Sale Number token",
                    })
                    print(f"  - {stock}: missing invoice token")
                    continue

                file_path = invoice_map.get(token)
                if not file_path:
                    results.append({
                        "Unit": stock,
                        "Invoice Token": token,
                        "File": "",
                        "Status": "Failed",
                        "Detail": "Sold updated; no matching file found in invoice folders",
                    })
                    print(f"  - {stock}: no file for token {token}")
                    continue

                # Save can return to inventory; reopen before attachment checks.
                if not DRY_RUN:
                    try:
                        dvf.open_vehicle_form(page, stock)
                        dvf.close_sticky_note_popup_if_present(page)
                        dvf.close_active_modals(page)
                        dvf.wait_for_ready(page, timeout=12000)
                    except Exception as e:
                        results.append({
                            "Unit": stock,
                            "Invoice Token": token,
                            "File": str(file_path),
                            "Status": "Failed",
                            "Detail": f"Sold updated but reopen for attachments failed: {e}",
                        })
                        print(f"  - {stock}: reopened failed after sold update")
                        continue

                if _attachment_already_present(page, file_path.name):
                    results.append({
                        "Unit": stock,
                        "Invoice Token": token,
                        "File": str(file_path),
                        "Status": "Skipped",
                        "Detail": "Sold updated; attachment already present",
                    })
                    print(f"  = {stock}: already has {file_path.name}")
                    continue

                if DRY_RUN:
                    results.append({
                        "Unit": stock,
                        "Invoice Token": token,
                        "File": str(file_path),
                        "Status": "Would Upload",
                        "Detail": "Sold updated; DRY RUN",
                    })
                    print(f"  ~ {stock}: would upload {file_path.name}")
                    continue

                # Preferred workflow: attach by drag/drop from file path on attachment screen.
                ok = _upload_via_drag_drop_path(page, file_path)
                if not ok:
                    # Fallback: click Browse and select file.
                    ok = _upload_via_browse_button(page, file_path)
                if not ok:
                    ok = upload_bill_of_sale_to_attachments(page, file_path)
                if ok:
                    results.append({
                        "Unit": stock,
                        "Invoice Token": token,
                        "File": str(file_path),
                        "Status": "Uploaded",
                        "Detail": "Sold updated",
                    })
                    print(f"  + {stock}: uploaded {file_path.name}")
                else:
                    results.append({
                        "Unit": stock,
                        "Invoice Token": token,
                        "File": str(file_path),
                        "Status": "Failed",
                        "Detail": "Upload function returned false",
                    })
                    print(f"  - {stock}: upload failed")

        finally:
            try:
                browser.close()
            except Exception:
                pass

    report_df = pd.DataFrame(results)
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        report_df.to_excel(writer, index=False, sheet_name="Results")
        if not report_df.empty:
            summary = report_df.groupby("Status").size().reset_index(name="Count").sort_values("Count", ascending=False)
        else:
            summary = pd.DataFrame([{"Status": "No Rows", "Count": 0}])
        summary.to_excel(writer, index=False, sheet_name="Summary")

    print(f"[Sold Invoices] Report: {report_path}")
    if report_df.empty:
        print("[Sold Invoices] No sold rows were processed.")
    else:
        print(report_df.groupby("Status").size().to_string())


if __name__ == "__main__":
    main()
