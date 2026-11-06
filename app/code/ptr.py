#!/usr/bin/env python3
"""
PTR scraper (month-to-date → SQLite → Excel)

- Navigate eFD home → accept terms → Get Access
- Apply  Senator + PTR filters
- Search month-to-date (1st → today)
- Collect up to 25 /search/view/ptr/<uuid>/ links
- For each link: open new tab → ensure printable view → scrape table
- Store raw and normalized transactions in SQLite; export one Excel sheet
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import re
import sqlite3
import time
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


BASE_URL = "https://efdsearch.senate.gov"
HOME_URL = f"{BASE_URL}/search/home/"
SEARCH_URL = f"{BASE_URL}/search/"


APP_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = APP_DIR / "data"
EXCEL_PATH = DATA_DIR / "PTR_Transactions.xlsx"
DB_PATH = DATA_DIR / "PTR.sqlite3"
RAW_TABLE = "ptr_transactions"
TX_TABLE = "transactions"


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
    raise TimeoutException(f"Could not click {name}")


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
        # if is_selected() fails for any reason, just click once
        try:
            el.click()
        except Exception:
            d.execute_script("arguments[0].click();", el)


def _month_to_date() -> tuple[str, str]:
    now = datetime.now()
    first = now.replace(day=1)
    return first.strftime("%m/%d/%Y"), now.strftime("%m/%d/%Y")


# ----------------------------
# UID helpers
# ----------------------------

def _normalize_for_uid(row: Dict[str, str]) -> List[tuple[str, str]]:
    """Normalize a scraped row into a deterministic list of (key, value) pairs."""
    ignore = {"trade uid", "scraped at"}
    items: List[tuple[str, str]] = []
    for k, v in row.items():
        kk = (k or "").strip().lower()
        if not kk or kk in ignore:
            continue
        vv = (v or "").strip()
        items.append((kk, vv))
    items.sort(key=lambda x: x[0])
    return items


def make_trade_uid(row: Dict[str, str]) -> str:
    """Create a stable UID for a single transaction row (used for dedupe + tracking)."""
    senator = (row.get("Senator") or "").strip()
    report_url = (row.get("Report URL") or "").strip()
    parts = [f"senator={senator}", f"report_url={report_url}"]
    for k, v in _normalize_for_uid(row):
        parts.append(f"{k}={v}")
    raw = "|".join(parts)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


# ----------------------------
# 1) terms → get access → search page
# ----------------------------

def accept_terms_and_get_access(d: webdriver.Firefox) -> None:
    d.get(HOME_URL)

    # wait for the terms panel to render
    WebDriverWait(d, 25).until(EC.presence_of_element_located((By.ID, "agree_statement")))

    # checkbox (some browsers require clicking the label/parent)
    try:
        _ensure_checked(d, (By.ID, "agree_statement"), timeout=25)
    except Exception:
        pass

    try:
        label = d.find_element(By.CSS_SELECTOR, "label[for='agree_statement']")
        d.execute_script("arguments[0].scrollIntoView({block:'center'});", label)
        d.execute_script("arguments[0].click();", label)
    except Exception:
        pass

    # IMPORTANT: on eFD: Home, the button is usually "Get Access"; IDs vary.
    access_locators = [
        (By.ID, "continueButton"),
        (By.CSS_SELECTOR, "button#continueButton"),
        (By.CSS_SELECTOR, "a#continueButton"),
        (By.CSS_SELECTOR, "input#continueButton"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Get Access')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Get Access')]"),
        (By.XPATH, "//input[@type='submit' and contains(@value, 'Get Access')]"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Continue')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Continue')]"),
        (By.CSS_SELECTOR, "div.get-access button"),
        (By.CSS_SELECTOR, "button.btn.btn-primary"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "input[type='submit']"),
    ]

    before_handles = list(d.window_handles)
    _safe_click_any(d, access_locators, timeout_each=10, name="Get Access/Continue")

    # Sometimes a new tab opens; if so, switch.
    try:
        WebDriverWait(d, 6).until(lambda x: len(x.window_handles) >= len(before_handles))
    except Exception:
        pass
    if len(d.window_handles) > len(before_handles):
        d.switch_to.window(d.window_handles[-1])

    # Wait until we land on the search page (date inputs present)
    WebDriverWait(d, 30).until(
        lambda x: bool(x.find_elements(By.ID, "fromDate")) or "/search/" in (x.current_url or "")
    )
    WebDriverWait(d, 30).until(EC.presence_of_element_located((By.ID, "fromDate")))


# ----------------------------

def _ptr_link_elements(d: webdriver.Firefox):
    # Prefer href-based detection (text can wrap or change)
    return d.find_elements(By.XPATH, "//a[contains(@href, '/search/view/ptr/')]")


# ----------------------------
# 2) set filters + dates + search
# ----------------------------

def set_filters_and_search(d: webdriver.Firefox, from_date: str, to_date: str) -> None:
    WebDriverWait(d, 25).until(EC.presence_of_element_located((By.ID, "fromDate")))

    senator_locators = [
        (By.XPATH, "//label[@id='filerTypeLabelSenator']//input[@type='checkbox']"),
        (By.CSS_SELECTOR, "label#filerTypeLabelSenator input[type='checkbox']"),
        (By.XPATH, "//input[@type='checkbox' and (@value='1' or contains(@id,'Senator'))]")
    ]
    for loc in senator_locators:
        try:
            _ensure_checked(d, loc, timeout=12)
            break
        except Exception:
            continue

    ptr_locators = [
        (By.XPATH, "//label[@id='reportTypeLabelPtr']//input[@type='checkbox']"),
        (By.CSS_SELECTOR, "label#reportTypeLabelPtr input[type='checkbox']"),
        (By.XPATH, "//input[@type='checkbox' and (@value='11' or contains(@id,'Ptr'))]")
    ]
    for loc in ptr_locators:
        try:
            _ensure_checked(d, loc, timeout=12)
            break
        except Exception:
            continue

    _safe_type(d, (By.ID, "fromDate"), from_date, timeout=20)
    _safe_type(d, (By.ID, "toDate"), to_date, timeout=20)

    search_locators = [
        (By.ID, "searchReports"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Search Reports')]"),
        (By.XPATH, "//form//button[contains(normalize-space(.), 'Search')]"),
        (By.CSS_SELECTOR, "form button.btn-primary"),
        (By.CSS_SELECTOR, "button.btn.btn-primary"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "input[type='submit']"),
    ]

    _safe_click_any(d, search_locators, timeout_each=12, name="Search")

    # Headless sometimes ignores the first click; if results don't appear quickly, click once more.
    time.sleep(1.5)
    if len(_ptr_link_elements(d)) == 0:
        try:
            _safe_click_any(d, search_locators, timeout_each=6, name="Search (retry)")
        except Exception:
            pass

    WebDriverWait(d, 45).until(lambda x: len(_ptr_link_elements(x)) > 0)


# ----------------------------
# 3) collect PTR links (cap 25)
# ----------------------------

def collect_ptr_links(d: webdriver.Firefox, max_reports: int = 25) -> List[str]:
    links: List[str] = []
    seen = set()

    # only need first page if you cap 25 (most times page 1 already has <= 25),
    # but this still works if there are more on page 1.
    els = _ptr_link_elements(d)
    for el in els:
        href = (el.get_attribute("href") or "").strip()
        if not href:
            continue
        if "/search/view/ptr/" not in href:
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
# 4) scrape one PTR report table
# ----------------------------

def _switch_to_newest_tab(d: webdriver.Firefox) -> None:
    d.switch_to.window(d.window_handles[-1])


def _open_new_tab(d: webdriver.Firefox, url: str) -> None:
    d.execute_script("window.open(arguments[0], '_blank');", url)
    _switch_to_newest_tab(d)


def _ensure_printable_ptr(d: webdriver.Firefox) -> None:
    # If the page shows a "Print Report" control, click it (often opens printable table view).
    if "eFD: Print Periodic Transaction Report" in (d.title or ""):
        return

    print_locators = [
        (By.XPATH, "//a[contains(normalize-space(.), 'Print Report')]"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Print Report')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Print')]"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Print')]"),
    ]

    before_handles = list(d.window_handles)
    for loc in print_locators:
        try:
            _safe_click(d, loc, timeout=6)
            break
        except Exception:
            continue

    # if a new tab opened, switch
    try:
        WebDriverWait(d, 6).until(lambda x: len(x.window_handles) >= len(before_handles))
    except Exception:
        pass

    if len(d.window_handles) > len(before_handles):
        _switch_to_newest_tab(d)


def scrape_ptr_table(d: webdriver.Firefox, report_url: str, per_link_wait: float = 4.0) -> List[Dict[str, str]]:
    _open_new_tab(d, report_url)

    try:
        _ensure_printable_ptr(d)

        if per_link_wait and per_link_wait > 0:
            time.sleep(float(per_link_wait))

        wait = WebDriverWait(d, 25)

        # senator header (your known-working selector)
        senator_el = wait.until(
            EC.presence_of_element_located((By.XPATH, "//h2[contains(text(),'(')]"))
        )
        senator_full = (senator_el.text or "").strip()

        # transactions table (your known-working selector)
        table = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "table")))
        wait.until(lambda _: len(table.find_elements(By.XPATH, ".//tr[td]")) > 0)

        headers = [h.text.strip() for h in table.find_elements(By.TAG_NAME, "th")]
        rows = table.find_elements(By.TAG_NAME, "tr")

        out: List[Dict[str, str]] = []
        for r in rows:
            tds = r.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue
            vals = [td.text.strip() for td in tds]
            if len(vals) < len(headers):
                vals += [""] * (len(headers) - len(vals))
            row_dict = dict(zip(headers, vals[: len(headers)]))
            row_dict["Senator"] = senator_full
            row_dict["Report URL"] = report_url

            # Stable, deterministic identifier used for dedupe/tracking (must match SQLite UID logic).
            trade_uid = make_trade_uid(row_dict)
            row_dict["Trade UID"] = trade_uid
            # Back-compat for code that expects a generic UID column.
            row_dict["UID"] = trade_uid

            out.append(row_dict)

        return out

    finally:
        # close report tab and go back to results tab
        try:
            d.close()
        except Exception:
            pass
        try:
            d.switch_to.window(d.window_handles[0])
        except Exception:
            pass


# ----------------------------
# 5) write combined excel
# ----------------------------

def write_excel(rows: List[Dict[str, str]], excel_path: Path) -> None:
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)

    # Ensure both columns exist and are identical so downstream tools can rely on either name.
    if "Trade UID" in df.columns and "UID" not in df.columns:
        df["UID"] = df["Trade UID"]
    if "UID" in df.columns and "Trade UID" not in df.columns:
        df["Trade UID"] = df["UID"]

    # Keep identifiers first for readability.
    preferred = [c for c in ["Trade UID", "UID"] if c in df.columns]
    cols = preferred + [c for c in df.columns if c not in preferred]
    df = df[cols]

    # overwrite one sheet with combined output
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Master Transactions", index=False)


# ----------------------------
# SQLite (deduped store + Excel-from-SQLite)
# ----------------------------

def db_connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA temp_store=MEMORY;")
    return conn


def _drop_view_if_exists(conn: sqlite3.Connection, name: str) -> None:
    cur = conn.execute(
        "SELECT type FROM sqlite_master WHERE name = ? COLLATE NOCASE LIMIT 1",
        (name,),
    )
    row = cur.fetchone()
    if row and (row[0] or "").lower() == "view":
        logging.warning("Found VIEW '%s' in DB; dropping so we can create a real table.", name)
        conn.execute(f"DROP VIEW IF EXISTS {name};")


def db_init(conn: sqlite3.Connection) -> None:
    # If older experiments created views, drop them so indexing/inserts work.
    _drop_view_if_exists(conn, RAW_TABLE)
    _drop_view_if_exists(conn, TX_TABLE)

    # Raw JSON store (keeps full scraped row)
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {RAW_TABLE} (
            trade_uid    TEXT PRIMARY KEY,
            senator      TEXT,
            report_url   TEXT,
            scraped_at   TEXT NOT NULL,
            payload_json TEXT NOT NULL
        );
        """
    )
    conn.execute(f"CREATE INDEX IF NOT EXISTS idx_ptr_scraped_at ON {RAW_TABLE}(scraped_at);")
    conn.execute(f"CREATE INDEX IF NOT EXISTS idx_ptr_senator ON {RAW_TABLE}(senator);")

    # Normalized table (what bot.py expects)
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {TX_TABLE} (
            trade_uid         TEXT PRIMARY KEY,
            senator           TEXT,
            transaction_date  TEXT,
            owner             TEXT,
            ticker            TEXT,
            asset_name        TEXT,
            asset_type        TEXT,
            transaction_type  TEXT,
            amount_range      TEXT,
            comment           TEXT,
            report_url        TEXT,
            created_at        TEXT NOT NULL
        );
        """
    )
    conn.execute(f"CREATE INDEX IF NOT EXISTS idx_tx_created_at ON {TX_TABLE}(created_at);")
    conn.execute(f"CREATE INDEX IF NOT EXISTS idx_tx_senator ON {TX_TABLE}(senator);")
    conn.execute(f"CREATE INDEX IF NOT EXISTS idx_tx_ticker ON {TX_TABLE}(ticker);")

    conn.commit()


def _clean(v: Optional[str]) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    if not s or s == "--":
        return None
    return s


def db_insert_rows(conn: sqlite3.Connection, rows: List[Dict[str, str]]) -> int:
    """Insert scraped rows into SQLite (raw + normalized), deduped by trade_uid."""
    if not rows:
        return 0

    now_iso = datetime.now().isoformat(timespec="seconds")

    raw_payloads = []
    tx_payloads = []

    for r in rows:
        r = dict(r)
        r.setdefault("Scraped At", now_iso)

        uid = (r.get("Trade UID") or r.get("UID") or make_trade_uid(r)).strip()
        senator = _clean(r.get("Senator"))
        report_url = _clean(r.get("Report URL"))

        # Ensure Excel + DB always have these columns
        r["Trade UID"] = uid
        r["UID"] = uid

        raw_payloads.append(
            (
                uid,
                senator,
                report_url,
                now_iso,
                json.dumps(r, ensure_ascii=False),
            )
        )

        tx_payloads.append(
            (
                uid,
                senator,
                _clean(r.get("Transaction Date")),
                _clean(r.get("Owner")),
                _clean(r.get("Ticker")),
                _clean(r.get("Asset Name")),
                _clean(r.get("Asset Type")),
                _clean(r.get("Type")),
                _clean(r.get("Amount")),
                _clean(r.get("Comment")),
                report_url,
                now_iso,
            )
        )

    before = conn.total_changes

    conn.executemany(
        f"INSERT OR IGNORE INTO {RAW_TABLE} (trade_uid, senator, report_url, scraped_at, payload_json) VALUES (?, ?, ?, ?, ?)",
        raw_payloads,
    )
    conn.executemany(
        f"""
        INSERT OR IGNORE INTO {TX_TABLE}
        (trade_uid, senator, transaction_date, owner, ticker, asset_name, asset_type, transaction_type, amount_range, comment, report_url, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        tx_payloads,
    )

    conn.commit()
    return max(conn.total_changes - before, 0)


def write_excel_from_sqlite(conn: sqlite3.Connection, excel_path: Path) -> int:
    """Write the Master Transactions sheet from the normalized SQLite table (deduped)."""
    df = pd.read_sql_query(
        f"""
        SELECT
            trade_uid        AS "Trade UID",
            trade_uid        AS "UID",
            senator          AS "Senator",
            transaction_date AS "Transaction Date",
            owner            AS "Owner",
            COALESCE(ticker, '') AS "Ticker",
            asset_name       AS "Asset Name",
            asset_type       AS "Asset Type",
            transaction_type AS "Type",
            amount_range     AS "Amount",
            comment          AS "Comment",
            report_url       AS "Report URL",
            created_at       AS "Scraped At"
        FROM {TX_TABLE}
        ORDER BY created_at DESC
        """,
        conn,
    )

    excel_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Master Transactions", index=False)

    return len(df)


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

    conn = db_connect(DB_PATH)
    db_init(conn)

    d = create_driver(headless=args.headless)
    try:
        accept_terms_and_get_access(d)
        set_filters_and_search(d, from_date, to_date)

        links = collect_ptr_links(d, max_reports=min(args.max_reports, 25))
        logging.info("Found %s PTR links (cap=25).", len(links))

        total_scraped = 0
        total_inserted = 0
        for i, url in enumerate(links, start=1):
            logging.info("Scraping %s/%s: %s", i, len(links), url)
            rows = scrape_ptr_table(d, url, per_link_wait=args.per_link_wait)
            total_scraped += len(rows)
            total_inserted += db_insert_rows(conn, rows)

        excel_rows = write_excel_from_sqlite(conn, EXCEL_PATH)
        logging.info(
            "Scraped %s rows; inserted %s new; Excel now has %s deduped rows. DB=%s Excel=%s",
            total_scraped,
            total_inserted,
            excel_rows,
            DB_PATH,
            EXCEL_PATH,
        )

    finally:
        try:
            d.quit()
        except Exception:
            pass
        try:
            conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
