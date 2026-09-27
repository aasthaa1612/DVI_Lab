"""
eCourts High Court of Uttarakhand - Case Data Crawler
======================================================
Extracts:  Case Status (by Case Type) + Case Orders/Judgements (by Date)
Output:    ecourts_dataset.xlsx  (two sheets)

Requirements (install once):
    pip install selenium requests openpyxl pillow webdriver-manager

Usage:
    python ecourts_crawler.py
"""

import time
import re
import sys
import logging
from datetime import datetime, timedelta
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ── Selenium imports ──────────────────────────────────────────────────────────
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (
        TimeoutException, NoSuchElementException, StaleElementReferenceException
    )
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("[ERROR] Missing packages. Run:  pip install selenium openpyxl webdriver-manager")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
BASE_URL   = "https://hcservices.ecourts.gov.in/ecourtindiaHC"
STATE_CD   = "15"
DIST_CD    = "1"
COURT_CODE = "1"
STATE_NM   = "Uttarakhand"

OUTPUT_FILE = "ecourts_dataset.xlsx"

# How many pages to scrape per case type  (set None for unlimited)
MAX_PAGES_PER_TYPE = 5

# Delay between requests (seconds) — be respectful to the server
REQUEST_DELAY = 2

# Date range for Orders/Judgements  (YYYY-MM-DD)
ORDER_FROM_DATE = (datetime.today() - timedelta(days=30)).strftime("%d-%m-%Y")  # last 30 days
ORDER_TO_DATE   = datetime.today().strftime("%d-%m-%Y")

# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("ecourts")

# ─────────────────────────────────────────────────────────────────────────────
# HELPER – build common query string
# ─────────────────────────────────────────────────────────────────────────────
def _qs():
    return f"state_cd={STATE_CD}&dist_cd={DIST_CD}&court_code={COURT_CODE}&stateNm={STATE_NM}"


# ─────────────────────────────────────────────────────────────────────────────
# BROWSER SETUP
# ─────────────────────────────────────────────────────────────────────────────
def create_driver():
    log.info("Launching Chrome browser …")
    options = Options()
    # ── Uncomment the line below to run headless (no visible window) ──
    # options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=options)
    driver.set_window_size(1400, 900)
    log.info("Browser ready.")
    return driver


# ─────────────────────────────────────────────────────────────────────────────
# WAIT HELPER
# ─────────────────────────────────────────────────────────────────────────────
def wait_for(driver, by, locator, timeout=15):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, locator))
    )


def wait_click(driver, by, locator, timeout=15):
    el = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, locator))
    )
    el.click()
    return el


# ─────────────────────────────────────────────────────────────────────────────
# EXTRACT TABLE ROWS – generic
# ─────────────────────────────────────────────────────────────────────────────
def extract_table(driver, table_id=None):
    """Return list-of-dicts from the first result table found on page."""
    try:
        if table_id:
            table = wait_for(driver, By.ID, table_id, timeout=10)
        else:
            tables = driver.find_elements(By.TAG_NAME, "table")
            # Pick the biggest table (most rows → most data)
            table = max(tables, key=lambda t: len(t.find_elements(By.TAG_NAME, "tr")))

        rows = table.find_elements(By.TAG_NAME, "tr")
        if not rows:
            return []

        # Extract headers from first row
        headers = [th.text.strip() for th in rows[0].find_elements(By.TAG_NAME, "th")]
        if not headers:
            headers = [td.text.strip() for td in rows[0].find_elements(By.TAG_NAME, "td")]

        records = []
        for row in rows[1:]:
            cells = [td.text.strip() for td in row.find_elements(By.TAG_NAME, "td")]
            if cells and any(c for c in cells):          # skip empty rows
                record = dict(zip(headers, cells))
                # Pad missing keys
                for h in headers:
                    record.setdefault(h, "")
                records.append(record)
        return records

    except Exception as e:
        log.warning(f"extract_table failed: {e}")
        return []


# ─────────────────────────────────────────────────────────────────────────────
# CAPTCHA HANDLER  (manual solve)
# ─────────────────────────────────────────────────────────────────────────────
def handle_captcha(driver, captcha_input_id="captcha_text", captcha_img_id=None):
    """
    Waits for user to solve CAPTCHA manually.
    Returns after user presses Enter in the terminal.
    """
    log.warning("=" * 60)
    log.warning("CAPTCHA DETECTED — please solve it in the browser window,")
    log.warning("then press  Enter  here to continue …")
    log.warning("=" * 60)
    input()    # pause until user presses Enter


# ─────────────────────────────────────────────────────────────────────────────
# MODULE 1: CASE STATUS  by Case Type
# ─────────────────────────────────────────────────────────────────────────────
def scrape_case_status_by_type(driver):
    url = f"{BASE_URL}/cases/s_casetype.php?{_qs()}"
    log.info(f"Opening Case-Type search: {url}")
    driver.get(url)
    time.sleep(REQUEST_DELAY)

    all_cases = []

    try:
        # Get list of case types from dropdown
        select_el = wait_for(driver, By.NAME, "case_type_id", timeout=12)
        select    = Select(select_el)
        options   = [(o.get_attribute("value"), o.text.strip())
                     for o in select.options if o.get_attribute("value")]

        log.info(f"Found {len(options)} case types to iterate.")

        for val, label in options:
            log.info(f"  ▶ Case type: {label!r}")
            driver.get(url)   # fresh load
            time.sleep(REQUEST_DELAY)

            try:
                sel = Select(wait_for(driver, By.NAME, "case_type_id", timeout=10))
                sel.select_by_value(val)
                time.sleep(0.5)

                # Check for captcha
                captcha_fields = driver.find_elements(By.ID, "captcha_text")
                if captcha_fields:
                    handle_captcha(driver)

                # Submit
                try:
                    submit = driver.find_element(By.XPATH,
                        "//input[@type='submit'] | //button[@type='submit']")
                    submit.click()
                except NoSuchElementException:
                    driver.find_element(By.XPATH, "//input[@value='Go']").click()

                time.sleep(REQUEST_DELAY + 1)

                page_num = 1
                while True:
                    rows = extract_table(driver)
                    for r in rows:
                        r["Case Type"] = label
                    all_cases.extend(rows)
                    log.info(f"    Page {page_num}: {len(rows)} rows  (total={len(all_cases)})")

                    if MAX_PAGES_PER_TYPE and page_num >= MAX_PAGES_PER_TYPE:
                        break

                    # Try to go to next page
                    try:
                        next_btn = driver.find_element(By.LINK_TEXT, "Next")
                        next_btn.click()
                        time.sleep(REQUEST_DELAY)
                        page_num += 1
                    except NoSuchElementException:
                        break

            except Exception as e:
                log.warning(f"    Skipping {label!r}: {e}")
                continue

    except TimeoutException:
        log.error("Could not find case_type dropdown — page may require CAPTCHA or JS.")

    return all_cases


# ─────────────────────────────────────────────────────────────────────────────
# MODULE 2: CASE STATUS  by Party Name  (wildcard search)
# ─────────────────────────────────────────────────────────────────────────────
def scrape_case_status_by_party(driver):
    url = f"{BASE_URL}/cases/ki_petres.php?{_qs()}"
    log.info(f"Opening Party-Name search: {url}")
    driver.get(url)
    time.sleep(REQUEST_DELAY)

    all_cases = []

    # Search common single letters to get broad results
    search_terms = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")

    for term in search_terms:
        log.info(f"  ▶ Party name starts with: {term!r}")
        driver.get(url)
        time.sleep(REQUEST_DELAY)

        try:
            inp = wait_for(driver, By.NAME, "pet_name", timeout=10)
            inp.clear()
            inp.send_keys(term)
            time.sleep(0.3)

            captcha_fields = driver.find_elements(By.ID, "captcha_text")
            if captcha_fields:
                handle_captcha(driver)

            try:
                submit = driver.find_element(By.XPATH,
                    "//input[@type='submit'] | //button[@type='submit']")
                submit.click()
            except NoSuchElementException:
                pass

            time.sleep(REQUEST_DELAY + 1)

            page_num = 1
            while True:
                rows = extract_table(driver)
                for r in rows:
                    r["Search Term"] = term
                all_cases.extend(rows)
                log.info(f"    Page {page_num}: {len(rows)} rows  (total={len(all_cases)})")

                if MAX_PAGES_PER_TYPE and page_num >= MAX_PAGES_PER_TYPE:
                    break

                try:
                    next_btn = driver.find_element(By.LINK_TEXT, "Next")
                    next_btn.click()
                    time.sleep(REQUEST_DELAY)
                    page_num += 1
                except NoSuchElementException:
                    break

        except Exception as e:
            log.warning(f"    Skipping {term!r}: {e}")
            continue

    return all_cases


# ─────────────────────────────────────────────────────────────────────────────
# MODULE 3: ORDERS / JUDGEMENTS  by Date
# ─────────────────────────────────────────────────────────────────────────────
def scrape_orders_by_date(driver):
    url = f"{BASE_URL}/cases/s_orderdate.php?{_qs()}"
    log.info(f"Opening Orders-by-Date search: {url}")
    driver.get(url)
    time.sleep(REQUEST_DELAY)

    all_orders = []

    try:
        # Fill from-date
        from_field = wait_for(driver, By.NAME, "from_date", timeout=10)
        from_field.clear()
        from_field.send_keys(ORDER_FROM_DATE)

        to_field = driver.find_element(By.NAME, "to_date")
        to_field.clear()
        to_field.send_keys(ORDER_TO_DATE)

        captcha_fields = driver.find_elements(By.ID, "captcha_text")
        if captcha_fields:
            handle_captcha(driver)

        try:
            submit = driver.find_element(By.XPATH,
                "//input[@type='submit'] | //button[@type='submit']")
            submit.click()
        except NoSuchElementException:
            pass

        time.sleep(REQUEST_DELAY + 1)

        page_num = 1
        while True:
            rows = extract_table(driver)
            for r in rows:
                r["Order Date Range"] = f"{ORDER_FROM_DATE} to {ORDER_TO_DATE}"
            all_orders.extend(rows)
            log.info(f"  Page {page_num}: {len(rows)} rows  (total={len(all_orders)})")

            try:
                next_btn = driver.find_element(By.LINK_TEXT, "Next")
                next_btn.click()
                time.sleep(REQUEST_DELAY)
                page_num += 1
            except NoSuchElementException:
                break

    except TimeoutException:
        log.error("Could not find order date fields.")

    return all_orders


# ─────────────────────────────────────────────────────────────────────────────
# MODULE 4: ORDERS / JUDGEMENTS  by Judge
# ─────────────────────────────────────────────────────────────────────────────
def scrape_orders_by_judge(driver):
    url = f"{BASE_URL}/cases/s_order.php?{_qs()}"
    log.info(f"Opening Orders-by-Judge search: {url}")
    driver.get(url)
    time.sleep(REQUEST_DELAY)

    all_orders = []

    try:
        select_el = wait_for(driver, By.NAME, "judge_name", timeout=12)
        select    = Select(select_el)
        options   = [(o.get_attribute("value"), o.text.strip())
                     for o in select.options if o.get_attribute("value")]

        log.info(f"Found {len(options)} judges.")

        for val, label in options:
            log.info(f"  ▶ Judge: {label!r}")
            driver.get(url)
            time.sleep(REQUEST_DELAY)

            try:
                sel = Select(wait_for(driver, By.NAME, "judge_name", timeout=10))
                sel.select_by_value(val)
                time.sleep(0.3)

                captcha_fields = driver.find_elements(By.ID, "captcha_text")
                if captcha_fields:
                    handle_captcha(driver)

                try:
                    submit = driver.find_element(By.XPATH,
                        "//input[@type='submit'] | //button[@type='submit']")
                    submit.click()
                except NoSuchElementException:
                    pass

                time.sleep(REQUEST_DELAY + 1)

                page_num = 1
                while True:
                    rows = extract_table(driver)
                    for r in rows:
                        r["Judge"] = label
                    all_orders.extend(rows)
                    log.info(f"    Page {page_num}: {len(rows)} rows  (total={len(all_orders)})")

                    if MAX_PAGES_PER_TYPE and page_num >= MAX_PAGES_PER_TYPE:
                        break

                    try:
                        next_btn = driver.find_element(By.LINK_TEXT, "Next")
                        next_btn.click()
                        time.sleep(REQUEST_DELAY)
                        page_num += 1
                    except NoSuchElementException:
                        break

            except Exception as e:
                log.warning(f"    Skipping {label!r}: {e}")
                continue

    except TimeoutException:
        log.error("Could not find judge dropdown.")

    return all_orders


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL WRITER
# ─────────────────────────────────────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", fgColor="1F3864")   # dark navy
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=11)
ALT_FILL     = PatternFill("solid", fgColor="EBF1F8")   # light blue
NORMAL_FILL  = PatternFill("solid", fgColor="FFFFFF")

def write_sheet(ws, records, sheet_title):
    ws.title = sheet_title

    if not records:
        ws.append(["No data collected"])
        return

    # Collect all unique column names (preserve insertion order)
    cols = list(dict.fromkeys(k for rec in records for k in rec.keys()))

    # Header row
    ws.append(cols)
    for cell in ws[1]:
        cell.fill  = HEADER_FILL
        cell.font  = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Data rows
    for i, rec in enumerate(records, start=2):
        row = [rec.get(c, "") for c in cols]
        ws.append(row)
        fill = ALT_FILL if i % 2 == 0 else NORMAL_FILL
        for cell in ws[i]:
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=False)

    # Auto-width
    for col_cells in ws.columns:
        length = max(len(str(c.value or "")) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(length + 4, 50)

    log.info(f"Sheet '{sheet_title}': {len(records)} rows, {len(cols)} columns written.")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    log.info("=" * 60)
    log.info(" eCourts – High Court of Uttarakhand  Data Crawler")
    log.info("=" * 60)

    driver = create_driver()

    try:
        # ── 1. Case Status by Case Type ──────────────────────────────────────
        case_status = scrape_case_status_by_type(driver)

        # ── 2. Orders by Date ────────────────────────────────────────────────
        orders_date = scrape_orders_by_date(driver)

        # ── 3. Orders by Judge ───────────────────────────────────────────────
        orders_judge = scrape_orders_by_judge(driver)

        # Merge order data (deduplicate by case number if present)
        all_orders = orders_date + orders_judge

    finally:
        driver.quit()
        log.info("Browser closed.")

    # ── Write Excel ──────────────────────────────────────────────────────────
    log.info(f"Writing Excel file: {OUTPUT_FILE}")
    wb = Workbook()
    ws1 = wb.active
    write_sheet(ws1, case_status,  "Case Status")

    ws2 = wb.create_sheet()
    write_sheet(ws2, all_orders,   "Orders & Judgements")

    # Summary sheet
    ws3 = wb.create_sheet("Summary")
    ws3.append(["Dataset", "Records", "Extracted At"])
    ws3.append(["Case Status (by Case Type)", len(case_status),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws3.append(["Orders & Judgements",       len(all_orders),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws3.append(["Source URL",
                f"{BASE_URL}/index_highcourt.php?state_cd={STATE_CD}&dist_cd={DIST_CD}&stateNm={STATE_NM}",
                ""])
    for cell in ws3[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    wb.save(OUTPUT_FILE)
    log.info(f"✅  Done! File saved → {Path(OUTPUT_FILE).resolve()}")
    log.info(f"   Case Status rows  : {len(case_status)}")
    log.info(f"   Orders rows       : {len(all_orders)}")


if __name__ == "__main__":
    main()
