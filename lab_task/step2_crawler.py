"""
eCourts High Court Uttarakhand - Interactive Data Crawler v2
=============================================================
Strategy:
  - Opens Chrome visibly so YOU can solve CAPTCHAs
  - You fill the form and submit it in the browser
  - Script auto-harvests ALL result pages
  - Saves everything to Excel

Run:   python step2_crawler.py
"""

import sys, io
# Force UTF-8 output to avoid Windows cmd encoding crashes
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import time
from pathlib import Path
from datetime import datetime

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import NoSuchElementException, WebDriverException
    from webdriver_manager.chrome import ChromeDriverManager
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    print("Run:  pip install selenium webdriver-manager openpyxl")
    sys.exit(1)

# ===========================================================================
# CONFIGURATION
# ===========================================================================
OUTPUT_FILE   = "ecourts_dataset.xlsx"
REQUEST_DELAY = 2      # seconds between page loads
MAX_PAGES     = 9999   # max result pages to collect per search

BASE   = "https://hcservices.ecourts.gov.in/ecourtindiaHC"
PARAMS = "state_cd=15&dist_cd=1&court_code=1&stateNm=Uttarakhand"

SEARCH_PAGES = [
    ("Case Status - by Case Type",  f"{BASE}/cases/s_casetype.php?{PARAMS}"),
    ("Case Status - by Party Name", f"{BASE}/cases/ki_petres.php?{PARAMS}"),
    ("Case Status - by Advocate",   f"{BASE}/cases/qs_civil_advocate.php?{PARAMS}"),
    ("Orders - by Date",            f"{BASE}/cases/s_orderdate.php?{PARAMS}"),
    ("Orders - by Judge",           f"{BASE}/cases/s_order.php?{PARAMS}"),
    ("Orders - by Party Name",      f"{BASE}/cases/s_partyorder.php?{PARAMS}"),
]

# ===========================================================================
# BROWSER
# ===========================================================================
def create_driver():
    options = Options()
    # Browser stays VISIBLE so user can solve CAPTCHA
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
    return driver


# ===========================================================================
# TABLE EXTRACTOR
# ===========================================================================
def extract_all_tables(driver):
    """
    Grabs every HTML table on the current page.
    Returns a list of dicts (one dict per row).
    Smart enough to handle merged headers and skip nav tables.
    """
    time.sleep(2)   # let AJAX finish rendering
    all_rows = []

    tables = driver.find_elements(By.TAG_NAME, "table")
    if not tables:
        return []

    # Skip tiny tables (navigation bars etc.) - keep ones with >2 rows
    data_tables = [t for t in tables if len(t.find_elements(By.TAG_NAME, "tr")) > 2]
    if not data_tables:
        data_tables = tables

    for table_idx, table in enumerate(data_tables):
        try:
            rows = table.find_elements(By.TAG_NAME, "tr")
            if not rows:
                continue

            # Find header row (look for <th> tags in first 3 rows)
            headers = []
            data_start = 1
            for idx, row in enumerate(rows[:3]):
                ths = row.find_elements(By.TAG_NAME, "th")
                if ths:
                    headers = [h.text.strip().replace("\n", " ") for h in ths]
                    data_start = idx + 1
                    break

            # Fall back: use first row's <td> as headers
            if not headers:
                first_tds = rows[0].find_elements(By.TAG_NAME, "td")
                headers = [td.text.strip() for td in first_tds]
                if not any(headers):
                    continue

            # Clean up duplicate header names
            seen = {}
            clean_hdrs = []
            for h in headers:
                h = h or f"Col{len(clean_hdrs)+1}"
                count = seen.get(h, 0)
                seen[h] = count + 1
                clean_hdrs.append(h if count == 0 else f"{h}_{count}")

            # Extract data rows
            for row in rows[data_start:]:
                tds = row.find_elements(By.TAG_NAME, "td")
                if not tds:
                    continue
                cells = [td.text.strip() for td in tds]
                if not any(cells):
                    continue

                # Pad/trim to match header count
                while len(cells) < len(clean_hdrs):
                    cells.append("")
                cells = cells[:len(clean_hdrs)]

                record = dict(zip(clean_hdrs, cells))
                record["_table"] = table_idx
                all_rows.append(record)
        except Exception as e:
            print(f"    [WARNING] Error reading table {table_idx}: {e}")
            continue

    return all_rows


# ===========================================================================
# NEXT PAGE BUTTON FINDER
# ===========================================================================
def find_next_button(driver):
    """Try many selectors to find a Next page button/link."""
    candidates = [
        (By.LINK_TEXT,       "Next"),
        (By.LINK_TEXT,       "next"),
        (By.LINK_TEXT,       ">"),
        (By.LINK_TEXT,       ">>"),
        (By.PARTIAL_LINK_TEXT, "Next"),
        (By.XPATH, "//a[contains(text(),'Next')]"),
        (By.XPATH, "//input[@value='Next']"),
        (By.XPATH, "//button[contains(text(),'Next')]"),
        (By.XPATH, "//li[contains(@class,'next')]/a"),
        (By.XPATH, "//a[@aria-label='Next']"),
        (By.XPATH, "//a[contains(@class,'next')]"),
        (By.XPATH, "//td[contains(text(),'Next')]/parent::tr//a"),
        # eCourts specific pagination
        (By.XPATH, "//a[contains(@href,'page') and contains(text(),'>')]"),
        (By.XPATH, "//span[@id='next' or @class='next']/a"),
    ]
    for by, sel in candidates:
        try:
            el = driver.find_element(by, sel)
            if el.is_displayed() and el.is_enabled():
                return el
        except NoSuchElementException:
            pass
    return None


# ===========================================================================
# SECTION SCRAPER (interactive)
# ===========================================================================
def scrape_section(driver, section_name, url):
    """
    Opens a search page, waits for user to fill/submit form,
    then auto-collects all result pages.
    """
    print(f"\n{'='*65}")
    print(f"  SECTION: {section_name}")
    print(f"{'='*65}")

    driver.get(url)
    time.sleep(REQUEST_DELAY)

    print(f"\n  Page opened in Chrome browser.")
    print(f"\n  WHAT TO DO NOW:")
    print(f"  1. Switch to the Chrome window that opened")
    print(f"  2. Fill in the search form")
    print(f"     - Select case type OR enter dates OR enter name etc.")
    print(f"  3. If there is a CAPTCHA image - type the letters shown")
    print(f"  4. Click the SUBMIT / SEARCH / GO button")
    print(f"  5. Wait for the results table to appear on screen")
    print(f"  6. Come back here to this terminal window")
    print(f"\n  [Press ENTER after results are loaded in the browser]")
    input()

    all_data = []
    page_num  = 1

    while page_num <= MAX_PAGES:
        print(f"\n  Collecting page {page_num} ...")

        rows = extract_all_tables(driver)

        if rows:
            for r in rows:
                r["_section"] = section_name
                r["_page"]    = page_num
                r["_time"]    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            all_data.extend(rows)
            print(f"  Page {page_num}: {len(rows)} rows found. Total so far: {len(all_data)}")
        else:
            print(f"  Page {page_num}: No table data found on this page.")
            print(f"  If you can see data in the browser but got 0 rows here,")
            print(f"  type 'c' and press Enter to retry. Otherwise just press Enter to stop.")
            resp = input("  > ").strip().lower()
            if resp != "c":
                break

        # Check for next page
        next_btn = find_next_button(driver)
        if next_btn:
            print(f"  Found Next page button - clicking ...")
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", next_btn)
                time.sleep(REQUEST_DELAY)
                page_num += 1
            except Exception as e:
                print(f"  Could not click next: {e}")
                break
        else:
            print(f"  No more pages. Collected {page_num} page(s) for this search.")
            break

    print(f"\n  DONE: {len(all_data)} total rows collected for '{section_name}'")
    return all_data


# ===========================================================================
# EXCEL WRITER
# ===========================================================================
HEADER_FILL = PatternFill("solid", fgColor="1F3864")   # dark navy
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
ALT_FILL    = PatternFill("solid", fgColor="DCE6F1")   # light blue
ODD_FILL    = PatternFill("solid", fgColor="FFFFFF")


def write_excel(all_sections_data):
    wb = Workbook()
    wb.remove(wb.active)   # remove default sheet

    total = 0
    summary_rows = []

    for section_name, records in all_sections_data.items():
        # Sheet names max 31 chars, no special chars
        sname = section_name[:28].replace("/", "-").replace(":", "").replace("*", "").strip()
        ws = wb.create_sheet(title=sname)

        if not records:
            ws.append(["No data was collected for this section."])
            summary_rows.append((section_name, 0))
            continue

        # Build column list: user-visible cols first, then internal _cols
        all_keys = list(dict.fromkeys(k for r in records for k in r.keys()))
        visible_cols  = [k for k in all_keys if not k.startswith("_")]
        internal_cols = [k for k in all_keys if k.startswith("_")]
        cols = visible_cols + internal_cols

        # Header row
        ws.append(cols)
        for cell in ws[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28

        # Data rows
        for i, rec in enumerate(records, start=2):
            row = [rec.get(c, "") for c in cols]
            ws.append(row)
            fill = ALT_FILL if i % 2 == 0 else ODD_FILL
            for cell in ws[i]:
                cell.fill = fill
                cell.alignment = Alignment(vertical="center")

        # Auto column widths (capped at 55)
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 55)

        ws.freeze_panes = "A2"   # freeze header row

        total += len(records)
        summary_rows.append((section_name, len(records)))
        print(f"  Sheet '{sname}': {len(records)} rows written")

    # Summary sheet (inserted at position 0)
    ws_s = wb.create_sheet(title="Summary", index=0)
    ws_s.append(["Section", "Records", "Extracted At"])
    for cell in ws_s[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")

    for i, (sec, cnt) in enumerate(summary_rows, start=2):
        ws_s.append([sec, cnt, datetime.now().strftime("%Y-%m-%d %H:%M")])
        for cell in ws_s[i]:
            cell.fill = ALT_FILL if i % 2 == 0 else ODD_FILL

    ws_s.append([])
    ws_s.append(["TOTAL RECORDS", total, ""])
    ws_s.append(["Source URL", f"{BASE}/index_highcourt.php?{PARAMS}", ""])

    ws_s.column_dimensions["A"].width = 40
    ws_s.column_dimensions["B"].width = 18
    ws_s.column_dimensions["C"].width = 20

    wb.save(OUTPUT_FILE)
    return total


# ===========================================================================
# MAIN
# ===========================================================================
def main():
    print("=" * 65)
    print("  eCourts - High Court of Uttarakhand | Data Crawler v2")
    print("=" * 65)
    print("""
  HOW IT WORKS:
  -------------
  A Chrome window will open for each search page.
  YOU fill the search form and click Submit (this handles CAPTCHA!).
  The script then AUTOMATICALLY collects ALL result pages.
  At the end, everything is saved to:  ecourts_dataset.xlsx

  You can do MULTIPLE searches per section to collect different
  case types one by one.

  Press Enter to start ...""")
    input()

    driver = create_driver()
    all_data = {}

    try:
        for idx, (section_name, url) in enumerate(SEARCH_PAGES, start=1):
            print(f"\n\n{'#'*65}")
            print(f"  SECTION {idx}/{len(SEARCH_PAGES)}: {section_name}")
            print(f"{'#'*65}")

            print("\n  Press 's' + Enter to SKIP this section, or just Enter to proceed:")
            resp = input("  > ").strip().lower()
            if resp == "s":
                all_data[section_name] = []
                print("  Skipped.")
                continue

            section_rows = []

            # Allow user to run multiple searches within same section
            while True:
                new_rows = scrape_section(driver, section_name, url)
                section_rows.extend(new_rows)

                print(f"\n  Do you want to search AGAIN in '{section_name}'?")
                print(f"  (e.g., pick a different case type / date range)")
                print(f"  Type 'y' + Enter for yes, or just Enter to move on:")
                again = input("  > ").strip().lower()
                if again != "y":
                    break

            all_data[section_name] = section_rows
            print(f"\n  Section complete. Total rows: {len(section_rows)}")

    except KeyboardInterrupt:
        print("\n\n  Interrupted - saving collected data ...")
    finally:
        try:
            driver.quit()
        except:
            pass
        print("  Browser closed.")

    # Save to Excel
    print(f"\n\n{'='*65}")
    print("  SAVING DATA TO EXCEL ...")
    print(f"{'='*65}")

    total = write_excel(all_data)
    out   = Path(OUTPUT_FILE).resolve()

    print(f"\n  DONE!")
    print(f"  File saved: {out}")
    print(f"  Total rows collected: {total}")
    print(f"\n  Open ecourts_dataset.xlsx to see your data!")


if __name__ == "__main__":
    main()
