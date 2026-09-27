# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

"""
STEP 1 - eCourts Page Inspector
================================
Run this FIRST. It opens each search page, waits for JS to render,
saves the HTML source + a screenshot, then prints all form fields found.

Run:   python step1_inspect.py

After it finishes, open the 'page_dumps' folder and check the HTML files.
Then run step2_crawler.py
"""

import time, sys
from pathlib import Path

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("Run:  pip install selenium webdriver-manager")
    sys.exit(1)

DUMP_DIR = Path("page_dumps")
DUMP_DIR.mkdir(exist_ok=True)

PAGES = {
    "casetype":  "https://hcservices.ecourts.gov.in/ecourtindiaHC/cases/s_casetype.php?state_cd=15&dist_cd=1&court_code=1&stateNm=Uttarakhand",
    "partyname": "https://hcservices.ecourts.gov.in/ecourtindiaHC/cases/ki_petres.php?state_cd=15&dist_cd=1&court_code=1&stateNm=Uttarakhand",
    "orderdate": "https://hcservices.ecourts.gov.in/ecourtindiaHC/cases/s_orderdate.php?state_cd=15&dist_cd=1&court_code=1&stateNm=Uttarakhand",
    "judge":     "https://hcservices.ecourts.gov.in/ecourtindiaHC/cases/s_order.php?state_cd=15&dist_cd=1&court_code=1&stateNm=Uttarakhand",
    "casenumber":"https://hcservices.ecourts.gov.in/ecourtindiaHC/cases/case_no.php?state_cd=15&dist_cd=1&court_code=1&stateNm=Uttarakhand",
}

options = Options()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36")

service = Service(ChromeDriverManager().install())
driver  = webdriver.Chrome(service=service, options=options)
driver.set_window_size(1400, 900)

for name, url in PAGES.items():
    print(f"\n{'='*60}")
    print(f"  Inspecting: {name}")
    print(f"  URL: {url}")
    print(f"{'='*60}")

    driver.get(url)

    print("  Waiting 5s for JS to render …")
    time.sleep(5)

    # ── Check for CAPTCHA ────────────────────────────────────────────
    captcha = driver.find_elements(By.XPATH, "//*[contains(@id,'captcha') or contains(@class,'captcha')]")
    if captcha:
        print("\n  *** CAPTCHA DETECTED ***")
        print("  Please solve the CAPTCHA in the browser, then press Enter here …")
        input()
        time.sleep(2)

    # ── Save screenshot ──────────────────────────────────────────────
    shot_path = DUMP_DIR / f"{name}_screenshot.png"
    driver.save_screenshot(str(shot_path))
    print(f"  Screenshot saved -> {shot_path}")

    # ── Save full page HTML ──────────────────────────────────────────
    html_path = DUMP_DIR / f"{name}.html"
    html_path.write_text(driver.page_source, encoding="utf-8")
    print(f"  HTML saved      -> {html_path}")

    # ── Print all INPUT/SELECT elements found ────────────────────────
    print("\n  [FORM FIELDS FOUND]")
    fields = driver.find_elements(By.XPATH, "//input | //select | //textarea | //button")
    if not fields:
    if not fields:
        print("  (none found - page may still be loading or blocked by CAPTCHA)")
    for f in fields:
        tag   = f.tag_name
        fid   = f.get_attribute("id")   or ""
        fname = f.get_attribute("name") or ""
        ftype = f.get_attribute("type") or ""
        fval  = f.get_attribute("value") or ""
        fcls  = f.get_attribute("class") or ""
        print(f"    <{tag}> id={fid!r:25} name={fname!r:25} type={ftype!r:10} val={fval!r:20} class={fcls!r}")

    # ── Print all SELECT options ─────────────────────────────────────
    selects = driver.find_elements(By.TAG_NAME, "select")
    for sel in selects:
        sel_id   = sel.get_attribute("id") or sel.get_attribute("name") or "unknown"
        options_list = sel.find_elements(By.TAG_NAME, "option")
        print(f"\n  [SELECT] '{sel_id}' has {len(options_list)} options:")
        for opt in options_list[:10]:   # first 10
            print(f"    value={opt.get_attribute('value')!r:10}  text={opt.text.strip()!r}")
        if len(options_list) > 10:
            print(f"    … and {len(options_list)-10} more")

    # ── Print page title ─────────────────────────────────────────────
    print(f"\n  Page Title: {driver.title!r}")
    print(f"  Current URL: {driver.current_url}")

print("\n\nINSPECTION COMPLETE!")
print(f"   HTML files saved in: {DUMP_DIR.resolve()}")
print("   Open them to see the actual field names, then run:  python step2_crawler.py")

driver.quit()
