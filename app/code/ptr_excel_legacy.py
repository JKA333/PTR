#!/usr/bin/env python3
"""
PTR legacy Excel-only scraper (month-to-date → Excel)

Flow:
- eFD home → accept terms → Get Access/Continue
- Select Senator + PTR filters
- Search (1st of month → today)
- Collect up to 25 /search/view/ptr/<uuid>/ links
- For each link: open new tab → (optional) click Print Report → scrape table
- Write one combined sheet to Excel
"""

from __future__ import annotations

import argparse
import logging
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

BASE_URL = "https://efdsearch.senate.gov"
HOME_URL = f"{BASE_URL}/search/home/"
EXCEL_PATH = Path(__file__).resolve().parent.parent / "data" / "PTR_Legacy_Transactions.xlsx"


# ----------------------------
# selenium helpers
# ----------------------------

def create_driver(headless: bool) -> webdriver.Firefox:
    opts = FirefoxOptions()
    if headless:
        opts.add_argument("--headless")
    d = webdriver.Firefox(options=opts)
    try:
        d.maximize_window()
    except Exception:
        pass
    return d


def _safe_click(d: webdriver.Firefox, locator: tuple, timeout: int = 20) -> None:
    el = WebDriverWait(d, timeout).until(EC.element_to_be_clickable(locator))
    d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.click()
    except Exception:
        d.execute_script("arguments[0].click();", el)


def _safe_click_any(
    d: webdriver.Firefox,
    locators: List[tuple],
    timeout_each: int = 8,
    name: str = "button",
) -> None:
    last_exc: Optional[Exception] = None
    for loc in locators:
        try:
            _safe_click(d, loc, timeout=timeout_each)
            return
        except Exception as exc:
            last_exc = exc
    raise TimeoutException(f"Could not click {name}") from last_exc


def _safe_type(d: webdriver.Firefox, locator: tuple, text: str, timeout: int = 20) -> None:
    el = WebDriverWait(d, timeout).until(EC.presence_of_element_located(locator))
    d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    el.clear()
    el.send_keys(text)


def _ensure_checked(d: webdriver.Firefox, locator: tuple, timeout: int = 20) -> None:
    el = WebDriverWait(d, timeout).until(EC.presence_of_element_located(locator))
    d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        if not el.is_selected():
            try:
                el.click()
            except Exception:
                d.execute_script("arguments[0].click();", el)
    except Exception:
        try:
            el.click()
        except Exception:
            d.execute_script("arguments[0].click();", el)


def _month_to_date() -> tuple[str, str]:
    now = datetime.now()
    first = now.replace(day=1)
    return first.strftime("%m/%d/%Y"), now.strftime("%m/%d/%Y")


def _ptr_link_elements(d: webdriver.Firefox):
    return d.find_elements(By.XPATH, "//a[contains(@href, '/search/view/ptr/')]")


def _on_results_page(d: webdriver.Firefox) -> bool:
    try:
        if d.find_elements(By.CSS_SELECTOR, "nav ul.pagination"):
            return True
        if _ptr_link_elements(d):
            return True
    except Exception:
        return False
    return False


# ----------------------------
# page flow: terms + filters + search
# ----------------------------

def accept_terms_and_set_filters(d: webdriver.Firefox, from_date: str, to_date: str) -> None:
    d.get(HOME_URL)
    wait = WebDriverWait(d, 25)

    _ensure_checked(d, (By.ID, "agree_statement"), timeout=25)

    continue_locators = [
        (By.ID, "continueButton"),
        (By.CSS_SELECTOR, "button#continueButton"),
        (By.CSS_SELECTOR, "a#continueButton"),
        (By.CSS_SELECTOR, "input#continueButton"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Get Access')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Get Access')]"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Continue')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Continue')]"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "input[type='submit']"),
    ]
    _safe_click_any(d, continue_locators, timeout_each=10, name="Get Access/Continue")

    wait.until(EC.presence_of_element_located((By.ID, "fromDate")))

    _ensure_checked(
        d,
        (By.XPATH, "//label[@id='filerTypeLabelSenator']//input[@type='checkbox']"),
        timeout=15,
    )
    _ensure_checked(
        d,
        (By.XPATH, "//label[@id='reportTypeLabelPtr']//input[@type='checkbox']"),
        timeout=15,
    )

    _safe_type(d, (By.ID, "fromDate"), from_date, timeout=20)
    _safe_type(d, (By.ID, "toDate"), to_date, timeout=20)

    search_locators = [
        (By.ID, "searchReports"),
        (By.XPATH, "//form//button[contains(normalize-space(.), 'Search Reports')]"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Search Reports')]"),
        (By.XPATH, "//form//button[contains(normalize-space(.), 'Search')]"),
        (By.CSS_SELECTOR, "form button.btn-primary"),
        (By.CSS_SELECTOR, "button.btn.btn-primary"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "input[type='submit']"),
    ]
    _safe_click_any(d, search_locators, timeout_each=12, name="Search")

    time.sleep(1.5)
    if len(_ptr_link_elements(d)) == 0 and not _on_results_page(d):
        try:
            _safe_click_any(d, search_locators, timeout_each=6, name="Search (retry)")
        except Exception:
            pass

    WebDriverWait(d, 45).until(lambda x: _on_results_page(x) or len(_ptr_link_elements(x)) > 0)


# ----------------------------
# collect PTR links (cap 25)
# ----------------------------

def collect_ptr_links(d: webdriver.Firefox, max_reports: int = 25) -> List[str]:
    links: List[str] = []
    seen = set()

    for el in _ptr_link_elements(d):
        href = (el.get_attribute("href") or "").strip()
        if not href:
            continue
        if "/search/view/paper/" in href:
            continue
        if href in seen:
            continue
        seen.add(href)
        links.append(href)
        if len(links) >= max_reports:
            break

    return links


# ----------------------------
# scrape one PTR report table
# ----------------------------

def _switch_to_newest_tab(d: webdriver.Firefox) -> None:
    d.switch_to.window(d.window_handles[-1])


def _open_new_tab(d: webdriver.Firefox, url: str) -> None:
    d.execute_script("window.open(arguments[0], '_blank');", url)
    _switch_to_newest_tab(d)


def _ensure_printable_ptr(d: webdriver.Firefox) -> None:
    if "eFD: Print Periodic Transaction Report" in (d.title or ""):
        return

    print_locators = [
        (By.XPATH, "//a[contains(normalize-space(.), 'Print Report')]"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Print Report')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Print')]"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Print')]"),
    ]

    before = list(d.window_handles)
    try:
        _safe_click_any(d, print_locators, timeout_each=6, name="Print Report")
    except Exception:
        return

    try:
        WebDriverWait(d, 6).until(lambda x: len(x.window_handles) >= len(before))
    except Exception:
        pass

    if len(d.window_handles) > len(before):
        _switch_to_newest_tab(d)


def scrape_ptr_table(
    d: webdriver.Firefox,
    report_url: str,
    per_link_wait: float = 4.0,
) -> List[Dict[str, str]]:
    _open_new_tab(d, report_url)

    try:
        _ensure_printable_ptr(d)

        if per_link_wait and per_link_wait > 0:
            time.sleep(float(per_link_wait))

        wait = WebDriverWait(d, 25)

        senator_el = wait.until(
            EC.presence_of_element_located((By.XPATH, "//h2[contains(text(),'(')]"))
        )
        senator_full = (senator_el.text or "").strip()

        table = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "table")))
        wait.until(lambda _: len(table.find_elements(By.XPATH, ".//tr[td]")) > 0)

        headers = [h.text.strip() for h in table.find_elements(By.TAG_NAME, "th")]

        out: List[Dict[str, str]] = []
        for r in table.find_elements(By.TAG_NAME, "tr"):
            tds = r.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue
            vals = [td.text.strip() for td in tds]
            if len(vals) < len(headers):
                vals += [""] * (len(headers) - len(vals))
            row = dict(zip(headers, vals[: len(headers)]))
            row["Senator"] = senator_full
            row["Report URL"] = report_url
            out.append(row)

        return out

    finally:
        try:
            d.close()
        except Exception:
            pass
        try:
            d.switch_to.window(d.window_handles[0])
        except Exception:
            pass


# ----------------------------
# write combined excel
# ----------------------------

def write_excel(rows: List[Dict[str, str]], excel_path: Path) -> None:
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Master Transactions", index=False)


# ----------------------------
# main
# ----------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("--headless", action="store_true")
    p.add_argument("--max-reports", type=int, default=25)
    p.add_argument("--per-link-wait", type=float, default=4.0)
    return p.parse_args()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

    args = parse_args()
    from_date, to_date = _month_to_date()

    d = create_driver(headless=args.headless)
    try:
        accept_terms_and_set_filters(d, from_date, to_date)

        limit = min(int(args.max_reports), 25)
        links = collect_ptr_links(d, max_reports=limit)
        logging.info("Found %s PTR links (cap=%s).", len(links), limit)

        all_rows: List[Dict[str, str]] = []
        for i, url in enumerate(links, start=1):
            logging.info("Scraping %s/%s: %s", i, len(links), url)
            all_rows.extend(scrape_ptr_table(d, url, per_link_wait=float(args.per_link_wait)))

        write_excel(all_rows, EXCEL_PATH)
        logging.info("Wrote %s rows to %s", len(all_rows), EXCEL_PATH)

    finally:
        try:
            d.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
