#!/usr/bin/env python3
"""
PTR extended Senate eFD scraper

- Selenium (Firefox) scraper for Senate eFD Periodic Transaction Reports (PTRs)
- Writes durable data to SQLite (dedupe by trade_uid)
- Also writes per-report sheets to Excel, then consolidates into:
  - Master Transactions
  - Legend

Notes:
- Correct domain is https://efdsearch.senate.gov (NO dash).
- Uses Homebrew geckodriver path by default on Apple Silicon:
  /opt/homebrew/bin/geckodriver

Operational principles (why this is built this way):
- Headless + WebDriverWait/EC (`wait.until(...)`) is critical for reliability (cloud/cron).
- Collect all report links first, then iterate: enables retries, logging, and dedupe.
- Electronically-filed PTRs are usually HTML (easy wins; no OCR required for tables).
- Daily cron runs fit the "forever database" model (SQLite as source of truth).
"""

from __future__ import annotations

import argparse
import hashlib
import logging
import os
import re
import sqlite3
import time
from io import StringIO
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Tuple, List, Dict, Set, Literal

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# =============================================================================
# PATHS / URLS
# =============================================================================

# Application root is the parent of code/.
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
LOGS_DIR = BASE_DIR / "logs"
IMAGES_DIR = DATA_DIR / "paper_images"
DEBUG_DIR = LOGS_DIR / "debug_screens"

DB_PATH = DATA_DIR / "PTR_extended.sqlite3"
EXCEL_PATH = DATA_DIR / "PTR_Extended_Transactions.xlsx"

# GeckoDriver path (Homebrew default on Apple Silicon)
GECKODRIVER_PATH = Path("/opt/homebrew/bin/geckodriver")

# Senate eFD URLs (NOTE: correct domain has NO dash)
BASE_URL = "https://efdsearch.senate.gov"
HOME_URL = f"{BASE_URL}/search/home/"
SEARCH_URL = f"{BASE_URL}/search/"


# =============================================================================
# LOGGING
# =============================================================================


def setup_logging() -> None:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(LOGS_DIR / "scraper.log"),
            logging.StreamHandler(),
        ],
    )


# =============================================================================
# SQLITE
# =============================================================================


def db_connect() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn


# --- DB schema migration helpers ---
def _table_columns(conn: sqlite3.Connection, table: str) -> Set[str]:
    try:
        rows = conn.execute(f"PRAGMA table_info({table});").fetchall()
        return {r[1] for r in rows}  # r[1] = column name
    except Exception:
        return set()


def db_migrate(conn: sqlite3.Connection) -> None:
    """Best-effort schema migration for existing DBs.

    `CREATE TABLE IF NOT EXISTS ...` won't modify existing tables.
    This function adds any missing columns we rely on so old DBs keep working.
    """

    expected: Dict[str, Dict[str, str]] = {
        "transactions": {
            "trade_uid": "TEXT",
            "report_url": "TEXT",
            "senator": "TEXT",
            "owner": "TEXT",
            "transaction_date": "TEXT",
            "ticker": "TEXT",
            "asset_name": "TEXT",
            "asset_type": "TEXT",
            "transaction_type": "TEXT",
            "amount_range": "TEXT",
            "source_page": "TEXT",
            "created_at": "TEXT",
        },
        "reports": {
            "report_url": "TEXT",
            "status": "TEXT",
            "last_error": "TEXT",
            "updated_at": "TEXT",
        },
        "paper_pages": {
            "report_url": "TEXT",
            "page_num": "INTEGER",
            "img_url": "TEXT",
            "sha256": "TEXT",
            "local_path": "TEXT",
        },
    }

    for table, cols in expected.items():
        existing = _table_columns(conn, table)
        if not existing:
            # Table might not exist yet; db_init() will create it.
            continue

        for col, col_type in cols.items():
            if col in existing:
                continue
            try:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type};")
                logging.info(
                    "DB migrate: added column %s.%s (%s)", table, col, col_type
                )
            except Exception as exc:
                logging.warning(
                    "DB migrate: failed adding column %s.%s: %s", table, col, exc
                )

    # Optional backfill if an older schema used senator_name
    try:
        tx_cols = _table_columns(conn, "transactions")
        if "senator" in tx_cols and "senator_name" in tx_cols:
            conn.execute(
                "UPDATE transactions SET senator = senator_name "
                "WHERE (senator IS NULL OR senator = '') AND senator_name IS NOT NULL;"
            )
    except Exception:
        pass

    # Backfill created_at if missing/null (keeps Excel export stable)
    try:
        tx_cols = _table_columns(conn, "transactions")
        if "created_at" in tx_cols:
            conn.execute(
                "UPDATE transactions SET created_at = COALESCE(created_at, '') "
                "WHERE created_at IS NULL;"
            )
    except Exception:
        pass

    conn.commit()


def db_init(conn: sqlite3.Connection) -> None:
    # Create tables if missing (does NOT modify existing schemas)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS transactions (
            trade_uid TEXT PRIMARY KEY,
            report_url TEXT NOT NULL,
            senator TEXT,
            owner TEXT,
            transaction_date TEXT,
            ticker TEXT,
            asset_name TEXT,
            asset_type TEXT,
            transaction_type TEXT,
            amount_range TEXT,
            source_page TEXT,
            created_at TEXT NOT NULL
        );
        """
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS reports (
            report_url TEXT PRIMARY KEY,
            status TEXT NOT NULL,
            last_error TEXT,
            updated_at TEXT NOT NULL
        );
        """
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS paper_pages (
            report_url TEXT NOT NULL,
            page_num INTEGER NOT NULL,
            img_url TEXT NOT NULL,
            sha256 TEXT,
            local_path TEXT,
            PRIMARY KEY (report_url, page_num)
        );
        """
    )

    conn.commit()

    # Migrate existing DBs (add any missing columns) before indexes.
    db_migrate(conn)

    # Indexes (best-effort; should succeed after migration)
    index_stmts = [
        "CREATE INDEX IF NOT EXISTS idx_transactions_ticker_date ON transactions(ticker, transaction_date);",
        "CREATE INDEX IF NOT EXISTS idx_transactions_senator_date ON transactions(senator, transaction_date);",
        "CREATE INDEX IF NOT EXISTS idx_transactions_uid ON transactions(trade_uid);",
        "CREATE INDEX IF NOT EXISTS idx_reports_status ON reports(status);",
        "CREATE INDEX IF NOT EXISTS idx_paper_pages_report ON paper_pages(report_url);",
    ]
    for stmt in index_stmts:
        try:
            conn.execute(stmt)
        except Exception as exc:
            logging.warning("DB init: index create failed: %s (%s)", stmt, exc)

    conn.commit()


# =============================================================================
# SELENIUM HELPERS
# =============================================================================

# Some older helper snippets (and copy/pasted code) may reference `driver`.
# Keep a module-level alias so those references never crash.
driver: Optional[webdriver.Firefox] = None


def create_driver(headless: bool = False) -> webdriver.Firefox:
    options = FirefoxOptions()
    if headless:
        options.add_argument("--headless")

    service = None
    if GECKODRIVER_PATH.exists():
        service = Service(str(GECKODRIVER_PATH))

    browser = (
        webdriver.Firefox(options=options, service=service)
        if service
        else webdriver.Firefox(options=options)
    )
    try:
        browser.maximize_window()
    except Exception:
        pass
    return browser


def _safe_click(browser: webdriver.Firefox, locator: tuple, timeout: int = 25) -> None:
    el = WebDriverWait(browser, timeout).until(EC.element_to_be_clickable(locator))
    browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
    try:
        el.click()
    except Exception:
        browser.execute_script("arguments[0].click();", el)


def _safe_type(
    browser: webdriver.Firefox, locator: tuple, text: str, timeout: int = 25
) -> None:
    el = WebDriverWait(browser, timeout).until(EC.presence_of_element_located(locator))
    browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
    el.clear()
    el.send_keys(text)


def _ensure_checked(
    browser: webdriver.Firefox, locator: tuple, timeout: int = 25
) -> None:
    el = WebDriverWait(browser, timeout).until(EC.presence_of_element_located(locator))
    browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
    try:
        selected = el.is_selected()
    except Exception:
        selected = False
    if not selected:
        try:
            el.click()
        except Exception:
            browser.execute_script("arguments[0].click();", el)


def _on_results_page(browser: webdriver.Firefox) -> bool:
    """Best-effort detection of the results listing page."""
    try:
        # Primary: your known results pagination nav
        if browser.find_elements(
            By.XPATH, "/html/body/div/main/div/div/div[2]/div/nav/ul"
        ):
            return True
        # Fallbacks: common results table + pagination
        if browser.find_elements(By.CSS_SELECTOR, "table") and browser.find_elements(
            By.CSS_SELECTOR, "nav ul.pagination"
        ):
            return True
        # Fallback: results table containing PTR links
        if browser.find_elements(
            By.XPATH, "//a[contains(normalize-space(.), 'Periodic Transaction Report')]"
        ):
            return True
    except Exception:
        return False
    return False


PageState = Literal[
    "FILTERS",
    "RESULTS",
    "REPORT_HTML",
    "REPORT_PAPER",
    "VIEW_REPORT",
    "UNKNOWN",
]


def detect_page_state(browser: webdriver.Firefox) -> PageState:
    """Classify the current page so downstream logic can react consistently."""
    try:
        title = (browser.title or "").strip()
        url = (browser.current_url or "").strip()

        # Terms/filters page
        if browser.find_elements(By.ID, "fromDate") and not _on_results_page(browser):
            return "FILTERS"

        # Results page
        if _on_results_page(browser):
            return "RESULTS"

        # Paper viewer
        if "/search/view/paper/" in url or browser.find_elements(
            By.CSS_SELECTOR, "img.filingImage"
        ):
            return "REPORT_PAPER"

        # Prefer printable PTR (this is the page we can reliably scrape)
        if "eFD: Print Periodic Transaction Report" in title:
            return "REPORT_HTML"

        # The alternate "view Report" page you want to skip
        if "eFD: view Report" in title:
            return "VIEW_REPORT"

        # HTML report view (table exists)
        if browser.find_elements(By.CSS_SELECTOR, "table"):
            return "REPORT_HTML"

    except Exception:
        return "UNKNOWN"

    return "UNKNOWN"


# =============================================================================
# HTML TABLE EXTRACTION (pandas-first)
# =============================================================================


# --- Helper: Find the transaction table on a PTR print page ---
def _find_transactions_table(browser: webdriver.Firefox):
    """Return the transaction table element on a PTR print page, or None."""
    try:
        tables = browser.find_elements(By.CSS_SELECTOR, "table")
    except Exception:
        return None

    for t in tables:
        try:
            # Prefer header-based detection (more stable than table class)
            headers = [
                h.text.strip() for h in t.find_elements(By.CSS_SELECTOR, "thead th")
            ]
            if not headers:
                headers = [
                    h.text.strip() for h in t.find_elements(By.CSS_SELECTOR, "tr th")
                ]
            header_blob = " | ".join([h for h in headers if h])
            if ("Transaction Date" in header_blob) and ("Owner" in header_blob):
                return t
        except Exception:
            continue

    return None


def _wait_for_transactions_table_loaded(
    browser: webdriver.Firefox, timeout: int = 25
) -> None:
    """Wait until the PTR transactions table exists and has at least one data row."""

    def _ready(d) -> bool:
        t = _find_transactions_table(d)
        if t is None:
            return False
        try:
            # Works whether tbody exists or not
            return len(t.find_elements(By.XPATH, ".//tr[td]")) > 0
        except Exception:
            return False

    WebDriverWait(browser, timeout).until(_ready)


def _clean_text(x: object) -> str:
    return re.sub(r"\s+", " ", str(x or "")).strip()


def _clean_ticker(x: object) -> Optional[str]:
    t = _clean_text(x)
    t = t.replace("—", "").replace("--", "").strip()
    return t or None


def _to_iso_date(maybe_date: str) -> str:
    s = _clean_text(maybe_date)
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if not m:
        return s
    mm, dd, yyyy = m.groups()
    return f"{int(yyyy):04d}-{int(mm):02d}-{int(dd):02d}"


def _make_trade_uid(parts: List[str]) -> str:
    raw = "|".join([_clean_text(p) for p in parts])
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _open_in_new_tab(browser: webdriver.Firefox, href: str) -> None:
    browser.execute_script("window.open(arguments[0], '_blank');", href)
    browser.switch_to.window(browser.window_handles[-1])


def _wait_for_report_content(browser: webdriver.Firefox, timeout: int = 25) -> None:
    def _ready(d):
        title = (d.title or "").strip()
        url = (d.current_url or "").strip()

        # Paper viewer
        if "/search/view/paper/" in url:
            return True
        if d.find_elements(By.CSS_SELECTOR, "img.filingImage"):
            return True

        # Printable PTR page (preferred)
        if "eFD: Print Periodic Transaction Report" in title:
            return True

        # Sometimes you land on /ptr/<uuid>/ first. If a Print control is present, we can proceed.
        if "/search/view/ptr/" in url:
            if d.find_elements(
                By.XPATH,
                "//a[contains(normalize-space(.), 'Print')] | //button[contains(normalize-space(.), 'Print')]",
            ):
                return True

        # HTML table present (any page that already contains the transactions)
        for t in d.find_elements(By.CSS_SELECTOR, "table"):
            txt = t.text or ""
            if "Transaction Date" in txt and "Owner" in txt:
                return True

        return False

    WebDriverWait(browser, timeout).until(_ready)


# Helper to click into the printable PTR view (and handle new tabs)
def _try_open_print_view(browser: webdriver.Firefox, timeout: int = 20) -> bool:
    """If we're on /view/ptr/<uuid>/, click into the printable view.

    Returns True if we believe we're on the printable page after this.
    """
    try:
        title = (browser.title or "").strip()
        if "eFD: Print Periodic Transaction Report" in title:
            return True

        url = (browser.current_url or "").strip()
        if "/search/view/ptr/" not in url:
            return False

        before_handles = list(browser.window_handles)
        before_url = url

        # Try several likely Print controls.
        print_locators = [
            (
                By.XPATH,
                "//a[contains(normalize-space(.), 'Print') and contains(normalize-space(.), 'Periodic Transaction')]",
            ),
            (
                By.XPATH,
                "//button[contains(normalize-space(.), 'Print') and contains(normalize-space(.), 'Periodic Transaction')]",
            ),
            (By.XPATH, "//a[contains(normalize-space(.), 'Print')]"),
            (By.XPATH, "//button[contains(normalize-space(.), 'Print')]"),
            (By.CSS_SELECTOR, "a[href*='print']"),
            (By.CSS_SELECTOR, "a.btn[href*='ptr']"),
        ]

        clicked = False
        for loc in print_locators:
            try:
                _safe_click(browser, loc, timeout=6)
                clicked = True
                break
            except Exception:
                continue

        if not clicked:
            return False

        # If a new tab opened, switch to it.
        try:
            WebDriverWait(browser, 6).until(
                lambda d: len(d.window_handles) >= len(before_handles)
            )
        except Exception:
            pass

        new_handles = list(browser.window_handles)
        if len(new_handles) > len(before_handles):
            browser.switch_to.window(new_handles[-1])

        # Wait for printable title or table presence.
        def _is_print_ready(d):
            t = (d.title or "").strip()
            if "eFD: Print Periodic Transaction Report" in t:
                return True
            if d.find_elements(By.CSS_SELECTOR, "table"):
                return True
            return False

        WebDriverWait(browser, timeout).until(_is_print_ready)

        # If the click didn't navigate at all, consider it a failure.
        if (
            browser.current_url or ""
        ).strip() == before_url and "eFD: Print Periodic Transaction Report" not in (
            browser.title or ""
        ):
            return False

        return "eFD: Print Periodic Transaction Report" in (
            browser.title or ""
        ) or bool(browser.find_elements(By.CSS_SELECTOR, "table"))

    except Exception:
        return False


# =============================================================================
# Debug artifact helper
# =============================================================================


def _save_debug_artifacts(browser: webdriver.Firefox, prefix: str) -> None:
    """Best-effort: save screenshot + HTML to logs/debug_screens for troubleshooting."""
    try:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        browser.save_screenshot(str(DEBUG_DIR / f"{prefix}-{ts}.png"))
        (DEBUG_DIR / f"{prefix}-{ts}.html").write_text(
            browser.page_source, encoding="utf-8"
        )
    except Exception:
        # Never allow debug capture to crash the run
        pass


def _get_ptr_link_elements(browser: webdriver.Firefox) -> List:
    els = browser.find_elements(
        By.XPATH,
        "//table//a[contains(normalize-space(.), 'Periodic Transaction Report')]",
    )
    if els:
        return els
    els = browser.find_elements(
        By.XPATH, "//a[contains(normalize-space(.), 'Periodic Transaction Report')]"
    )
    if els:
        return els
    return browser.find_elements(By.CSS_SELECTOR, "a[href*='/search/view/']")


def _click_next_results_page(browser: webdriver.Firefox) -> bool:
    candidates = browser.find_elements(
        By.XPATH, "//nav//ul//a[contains(normalize-space(.), 'Next')]"
    )
    if not candidates:
        candidates = browser.find_elements(
            By.XPATH,
            "//nav//ul//a[contains(@aria-label, 'Next') or contains(., '›') or contains(., '→')]",
        )
    if not candidates:
        return False

    before = browser.current_url
    _safe_click(
        browser,
        (
            By.XPATH,
            "(//nav//ul//a[contains(normalize-space(.), 'Next') or contains(@aria-label, 'Next') or contains(., '›') or contains(., '→')])[1]",
        ),
        timeout=15,
    )
    WebDriverWait(browser, 25).until(
        lambda d: d.current_url != before or len(_get_ptr_link_elements(d)) > 0
    )
    return True


# =============================================================================
# PAGE FLOW: TERMS + FILTERS
# =============================================================================


def accept_terms_and_set_filters(
    browser: webdriver.Firefox, from_date: str, to_date: str
) -> None:
    """Accept terms, set Senator + PTR filters, and run the search."""
    browser.get(HOME_URL)
    wait = WebDriverWait(browser, 25)

    # 1) Terms page
    _ensure_checked(browser, (By.ID, "agree_statement"), timeout=25)

    continue_locators = [
        (By.ID, "continueButton"),
        (By.CSS_SELECTOR, "button#continueButton"),
        (By.CSS_SELECTOR, "input#continueButton"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "input[type='submit']"),
        (By.XPATH, "//button[contains(normalize-space(.), 'Continue')]"),
        (
            By.XPATH,
            "//input[@type='submit' and (contains(@value,'Continue') or contains(@value,'Accept'))]",
        ),
        (By.XPATH, "//button[contains(normalize-space(.), 'Accept')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Continue')]"),
    ]

    clicked = False
    last_exc: Optional[Exception] = None
    for loc in continue_locators:
        try:
            _safe_click(browser, loc, timeout=8)
            clicked = True
            break
        except Exception as exc:
            last_exc = exc

    if not clicked:
        _save_debug_artifacts(browser, "continue-fail")
        raise TimeoutException(
            "Could not find/click Continue button on terms page"
        ) from last_exc

    # 2) Filters page
    wait.until(EC.presence_of_element_located((By.ID, "fromDate")))

    # Ensure checkboxes are selected (try multiple locators; first match wins)
    senator_locators = [
        (By.XPATH, "//label[@id='filerTypeLabelSenator']//input[@type='checkbox']"),
        (By.XPATH, "//label[@id='filerTypeLabelSenator']/input[@value='1']"),
        (By.CSS_SELECTOR, "label#filerTypeLabelSenator input[type='checkbox']"),
        (
            By.XPATH,
            "//input[@type='checkbox' and @value='1' and (contains(@name,'filer') or contains(@id,'filer'))]",
        ),
    ]
    for loc in senator_locators:
        try:
            _ensure_checked(browser, loc, timeout=10)
            break
        except Exception:
            continue

    ptr_locators = [
        (By.XPATH, "//label[@id='reportTypeLabelPtr']//input[@type='checkbox']"),
        (By.XPATH, "//label[@id='reportTypeLabelPtr']/input[@value='11']"),
        (By.CSS_SELECTOR, "label#reportTypeLabelPtr input[type='checkbox']"),
        (
            By.XPATH,
            "//input[@type='checkbox' and @value='11' and (contains(@name,'report') or contains(@id,'report'))]",
        ),
    ]
    for loc in ptr_locators:
        try:
            _ensure_checked(browser, loc, timeout=10)
            break
        except Exception:
            continue

    _safe_type(browser, (By.ID, "fromDate"), from_date, timeout=20)
    _safe_type(browser, (By.ID, "toDate"), to_date, timeout=20)

    # 3) Run search (primary + fallbacks)
    search_locators = [
        (By.XPATH, "//form//button[contains(normalize-space(.), 'Search Reports')]"),
        (By.XPATH, "//form//button[contains(normalize-space(.), 'Search')]"),
        (By.CSS_SELECTOR, "form button.btn-primary"),
        (By.CSS_SELECTOR, "button.btn.btn-primary"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "input[type='submit']"),
    ]

    clicked = False
    for loc in search_locators:
        try:
            _safe_click(browser, loc, timeout=10)
            clicked = True
            break
        except Exception:
            continue

    if not clicked:
        _save_debug_artifacts(browser, "search-click-fail")
        raise TimeoutException("Could not find/click Search button on filters page")

    # 4) Wait for results to actually load.
    # The site can ignore the first click in headless; if we're still on filters after a moment,
    # try clicking again once, then wait for results.
    time.sleep(1.5)

    def _still_on_filters(d) -> bool:
        try:
            return bool(d.find_elements(By.ID, "fromDate")) and not _on_results_page(d)
        except Exception:
            return False

    if _still_on_filters(browser):
        for loc in search_locators:
            try:
                _safe_click(browser, loc, timeout=6)
                break
            except Exception:
                continue

    try:
        WebDriverWait(browser, 30).until(
            lambda d: _on_results_page(d) or len(_get_ptr_link_elements(d)) > 0
        )
    except TimeoutException:
        _save_debug_artifacts(browser, "results-timeout")
        raise


# =============================================================================
# PAPER VIEWER
# =============================================================================


def is_paper_filing(browser: webdriver.Firefox) -> bool:
    try:
        if "/search/view/paper/" in browser.current_url:
            return True
        browser.find_element(By.CSS_SELECTOR, "img.filingImage")
        return True
    except NoSuchElementException:
        return False


def collect_paper_pages(
    browser: webdriver.Firefox,
    report_url: str,
    conn: sqlite3.Connection,
    download_images: bool = True,
) -> int:
    """Collect all image pages from the /paper/ viewer.

    Stores each page in SQLite `paper_pages` (dedupe by report_url + page_num).
    If `download_images` is True, saves the image bytes under data/paper_images/<uuid>/page_###.png.

    Safety:
    - hard cap on pages
    - stop if the image URL repeats
    - stop if clicking Next does not change the image src
    """

    report_dir: Optional[Path] = None
    if download_images:
        IMAGES_DIR.mkdir(parents=True, exist_ok=True)
        uuid_match = re.search(r"/paper/([a-f0-9-]+)/", report_url)
        report_uuid = uuid_match.group(1) if uuid_match else "unknown"
        report_dir = IMAGES_DIR / report_uuid
        report_dir.mkdir(parents=True, exist_ok=True)

    page_count = 0
    current_page = 1

    seen_img_urls: Set[str] = set()
    max_pages = 200

    while True:
        try:
            if current_page > max_pages:
                logging.warning(
                    "Hit max_pages=%s in paper viewer; stopping for %s",
                    max_pages,
                    report_url,
                )
                break

            img_element = WebDriverWait(browser, 12).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "img.filingImage"))
            )
            img_url = (img_element.get_attribute("src") or "").strip()
            if not img_url:
                break

            if img_url in seen_img_urls:
                logging.info(
                    "Paper viewer repeated the same image; stopping at page %s for %s",
                    current_page,
                    report_url,
                )
                break
            seen_img_urls.add(img_url)

            img_hash = hashlib.sha256(img_url.encode("utf-8")).hexdigest()

            local_path: Optional[str] = None
            if download_images and report_dir is not None:
                filename = f"page_{current_page:03d}.png"
                local_path = str(report_dir / filename)

                import requests

                s = requests.Session()
                # Copy Selenium cookies so the image request is authorized
                for c in browser.get_cookies():
                    try:
                        s.cookies.set(
                            c.get("name"), c.get("value"), domain=c.get("domain")
                        )
                    except Exception:
                        continue

                resp = s.get(img_url, timeout=30)
                resp.raise_for_status()
                Path(local_path).write_bytes(resp.content)

            cur = conn.execute(
                "SELECT 1 FROM paper_pages WHERE report_url = ? AND page_num = ?",
                (report_url, current_page),
            )
            if cur.fetchone() is None:
                conn.execute(
                    "INSERT INTO paper_pages (report_url, page_num, img_url, sha256, local_path) VALUES (?, ?, ?, ?, ?)",
                    (report_url, current_page, img_url, img_hash, local_path),
                )
                conn.commit()
                page_count += 1
                logging.info("Collected paper page %s for %s", current_page, report_url)

            # Find a Next control
            next_candidates = browser.find_elements(
                By.CSS_SELECTOR, "a[aria-label*='Next'], button[aria-label*='Next']"
            )
            if not next_candidates:
                next_candidates = browser.find_elements(
                    By.XPATH,
                    "//a[contains(@class,'next') or contains(., 'Next') or contains(., '›') or contains(., '→')]"
                    " | //button[contains(., 'Next') or contains(., '›') or contains(., '→')]",
                )
            if not next_candidates:
                break

            next_button = next_candidates[0]
            cls = (next_button.get_attribute("class") or "").lower()
            aria_disabled = (next_button.get_attribute("aria-disabled") or "").lower()
            if "disabled" in cls or aria_disabled == "true":
                break

            prev_src = img_url
            browser.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", next_button
            )
            try:
                next_button.click()
            except Exception:
                browser.execute_script("arguments[0].click();", next_button)

            # Wait for the src to change; if it doesn't, we are stuck/end.
            try:
                WebDriverWait(browser, 8).until(
                    lambda d: (
                        d.find_element(
                            By.CSS_SELECTOR, "img.filingImage"
                        ).get_attribute("src")
                        or ""
                    ).strip()
                    != prev_src
                )
            except Exception:
                logging.info(
                    "Paper viewer did not advance (src unchanged); stopping at page %s for %s",
                    current_page,
                    report_url,
                )
                break

            time.sleep(0.5)
            current_page += 1

        except TimeoutException:
            logging.warning(
                "Timeout waiting for paper page %s in %s", current_page, report_url
            )
            break
        except Exception as e:
            logging.error(
                "Error collecting paper page %s in %s: %s", current_page, report_url, e
            )
            break

    return page_count


# =============================================================================
# HTML TABLE EXTRACTION (pandas-first)
# =============================================================================


def save_to_excel(df: pd.DataFrame, file_path: Path, sheet_name: str) -> None:
    file_path = Path(file_path)
    file_path.parent.mkdir(parents=True, exist_ok=True)

    writer_args = {"engine": "openpyxl", "mode": "a", "if_sheet_exists": "overlay"}
    if not file_path.exists():
        writer_args.pop("mode")
        writer_args.pop("if_sheet_exists")

    with pd.ExcelWriter(file_path, **writer_args) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)


def extract_html_transactions(
    browser: webdriver.Firefox,
    report_url: str,
    conn: sqlite3.Connection,
    excel_sheet: Optional[str] = None,
) -> int:
    try:
        senator_full_name = ""
        for sel in ["h2", "h1"]:
            els = browser.find_elements(By.CSS_SELECTOR, sel)
            if els:
                txt = (els[0].text or "").strip()
                if txt:
                    senator_full_name = txt
                    break

        # Make sure the JS-rendered table has populated before we parse.
        _wait_for_transactions_table_loaded(browser, timeout=25)

        try:
            dfs = pd.read_html(StringIO(browser.page_source))
        except Exception as exc:
            # Common causes: missing `lxml`, or page not fully loaded.
            logging.warning(
                "pd.read_html failed for %s (%s). Will fall back to DOM extraction.",
                report_url,
                exc,
            )
            _save_debug_artifacts(browser, "read-html-failed")
            return 0
        chosen = None
        for df in dfs:
            cols = [str(c).strip() for c in df.columns]
            if any("Transaction Date" in c for c in cols) and any(
                "Owner" in c for c in cols
            ):
                chosen = df
                break

        if chosen is None:
            logging.warning("No PTR transaction table found in %s", report_url)
            return 0

        chosen.columns = [str(c).strip() for c in chosen.columns]
        chosen = chosen.rename(
            columns={c: re.sub(r"\s+", " ", str(c)).strip() for c in chosen.columns}
        )

        rows_inserted = 0
        out_rows: List[Dict[str, object]] = []

        conn.execute("BEGIN")
        try:
            for _, r in chosen.iterrows():
                transaction_date_raw = _clean_text(r.get("Transaction Date", ""))
                transaction_date = _to_iso_date(transaction_date_raw)

                owner = _clean_text(r.get("Owner", ""))
                ticker = _clean_ticker(r.get("Ticker", ""))
                asset_name = _clean_text(r.get("Asset Name", ""))
                asset_type = _clean_text(r.get("Asset Type", ""))
                tx_type = _clean_text(r.get("Type", ""))
                amount_range = _clean_text(r.get("Amount", ""))

                trade_uid = _make_trade_uid(
                    [
                        report_url,
                        transaction_date,
                        owner,
                        ticker or "",
                        asset_name,
                        asset_type,
                        tx_type,
                        amount_range,
                    ]
                )

                cur = conn.execute(
                    """
                    INSERT OR IGNORE INTO transactions
                    (trade_uid, report_url, senator, owner, transaction_date, ticker, asset_name, asset_type,
                     transaction_type, amount_range, source_page, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        trade_uid,
                        report_url,
                        senator_full_name,
                        owner,
                        transaction_date,
                        ticker,
                        asset_name,
                        asset_type,
                        tx_type,
                        amount_range,
                        browser.current_url,
                        datetime.now().isoformat(),
                    ),
                )

                if cur.rowcount == 1:
                    rows_inserted += 1

                out_rows.append(
                    {
                        "trade_uid": trade_uid,
                        "Senator": senator_full_name,
                        "Transaction Date": transaction_date,
                        "Owner": owner,
                        "Ticker": ticker or "",
                        "Asset Name": asset_name,
                        "Asset Type": asset_type,
                        "Type": tx_type,
                        "Amount": amount_range,
                        "Report URL": report_url,
                    }
                )

            conn.commit()
        except Exception:
            conn.rollback()
            raise

        if excel_sheet and out_rows:
            df_out = pd.DataFrame(out_rows)
            save_to_excel(df_out, EXCEL_PATH, excel_sheet)

        return rows_inserted

    except Exception as e:
        logging.error("Error extracting HTML transactions from %s: %s", report_url, e)
        return 0


# DOM-based fallback extractor (no lxml needed)
def extract_html_transactions_dom(
    browser: webdriver.Firefox,
    report_url: str,
    conn: sqlite3.Connection,
    excel_sheet: Optional[str] = None,
) -> int:
    """Fallback extractor that uses Selenium DOM parsing instead of pd.read_html.

    This helps when pd.read_html fails (missing lxml, weird markup, etc.).
    """
    try:
        senator_full_name = ""
        for sel in ["h2", "h1"]:
            els = browser.find_elements(By.CSS_SELECTOR, sel)
            if els:
                txt = (els[0].text or "").strip()
                if txt:
                    senator_full_name = txt
                    break

        _wait_for_transactions_table_loaded(browser, timeout=25)

        target = _find_transactions_table(browser)
        if target is None:
            logging.warning(
                "DOM fallback: no PTR transaction table found in %s", report_url
            )
            _save_debug_artifacts(browser, "dom-no-table")
            return 0

        headers = [
            h.text.strip() for h in target.find_elements(By.CSS_SELECTOR, "thead th")
        ]
        if not headers:
            headers = [
                h.text.strip() for h in target.find_elements(By.CSS_SELECTOR, "tr th")
            ]

        rows = []
        for tr in target.find_elements(By.XPATH, ".//tr[td]"):
            tds = tr.find_elements(By.CSS_SELECTOR, "td")
            if not tds:
                continue
            rows.append([td.text.strip() for td in tds])

        if not headers or not rows:
            logging.warning("DOM fallback: table had no headers/rows in %s", report_url)
            return 0

        # Align row lengths
        norm_rows = []
        for r in rows:
            if len(r) < len(headers):
                r = r + [""] * (len(headers) - len(r))
            norm_rows.append(r[: len(headers)])

        df = pd.DataFrame(norm_rows, columns=headers)
        df = df.rename(
            columns={c: re.sub(r"\s+", " ", str(c)).strip() for c in df.columns}
        )

        # Map columns to expected names (print view sometimes varies)
        col_map = {c: c for c in df.columns}
        # Ensure exact keys used below
        # (If the site changes slightly, these still work as long as words appear.)
        for c in df.columns:
            if "Transaction Date" in c:
                col_map[c] = "Transaction Date"
            elif c.strip() == "Owner" or "Owner" in c:
                col_map[c] = "Owner"
            elif c.strip() == "Ticker" or "Ticker" in c:
                col_map[c] = "Ticker"
            elif "Asset Name" in c:
                col_map[c] = "Asset Name"
            elif "Asset Type" in c:
                col_map[c] = "Asset Type"
            elif c.strip() == "Type" or "Type" in c:
                col_map[c] = "Type"
            elif c.strip() == "Amount" or "Amount" in c:
                col_map[c] = "Amount"

        df = df.rename(columns=col_map)

        rows_inserted = 0
        out_rows: List[Dict[str, object]] = []

        conn.execute("BEGIN")
        try:
            for _, r in df.iterrows():
                transaction_date_raw = _clean_text(r.get("Transaction Date", ""))
                transaction_date = _to_iso_date(transaction_date_raw)

                owner = _clean_text(r.get("Owner", ""))
                ticker = _clean_ticker(r.get("Ticker", ""))
                asset_name = _clean_text(r.get("Asset Name", ""))
                asset_type = _clean_text(r.get("Asset Type", ""))
                tx_type = _clean_text(r.get("Type", ""))
                amount_range = _clean_text(r.get("Amount", ""))

                trade_uid = _make_trade_uid(
                    [
                        report_url,
                        transaction_date,
                        owner,
                        ticker or "",
                        asset_name,
                        asset_type,
                        tx_type,
                        amount_range,
                    ]
                )

                cur = conn.execute(
                    """
                    INSERT OR IGNORE INTO transactions
                    (trade_uid, report_url, senator, owner, transaction_date, ticker, asset_name, asset_type,
                     transaction_type, amount_range, source_page, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        trade_uid,
                        report_url,
                        senator_full_name,
                        owner,
                        transaction_date,
                        ticker,
                        asset_name,
                        asset_type,
                        tx_type,
                        amount_range,
                        browser.current_url,
                        datetime.now().isoformat(),
                    ),
                )

                if cur.rowcount == 1:
                    rows_inserted += 1

                out_rows.append(
                    {
                        "trade_uid": trade_uid,
                        "Senator": senator_full_name,
                        "Transaction Date": transaction_date,
                        "Owner": owner,
                        "Ticker": ticker or "",
                        "Asset Name": asset_name,
                        "Asset Type": asset_type,
                        "Type": tx_type,
                        "Amount": amount_range,
                        "Report URL": report_url,
                    }
                )

            conn.commit()
        except Exception:
            conn.rollback()
            raise

        if excel_sheet and out_rows:
            df_out = pd.DataFrame(out_rows)
            save_to_excel(df_out, EXCEL_PATH, excel_sheet)

        return rows_inserted

    except Exception as e:
        logging.error(
            "DOM fallback: error extracting transactions from %s: %s", report_url, e
        )
        return 0


# Selenium-table extractor (senator01-style): waits for class="table" and scrapes via DOM
def extract_html_transactions_table_class(
    browser: webdriver.Firefox,
    report_url: str,
    conn: sqlite3.Connection,
    excel_sheet: Optional[str] = None,
    timeout: int = 25,
) -> int:
    """
    senator01-style extractor:
    - Wait for senator header (best-effort)
    - Wait for `table.table` to exist + have data rows
    - Scrape th/td via DOM
    - Insert into SQLite (dedupe by trade_uid)
    """
    try:
        wait = WebDriverWait(browser, timeout)

        # Senator header (best-effort)
        senator_full_name = ""
        try:
            senator_el = wait.until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(),'(')]"))
            )
            senator_full_name = (senator_el.text or "").strip()
        except Exception:
            for sel in ("h2", "h1"):
                els = browser.find_elements(By.CSS_SELECTOR, sel)
                if els:
                    senator_full_name = (els[0].text or "").strip()
                    if senator_full_name:
                        break

        # Wait for the main transactions table
        table = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "table")))
        wait.until(lambda d: len(table.find_elements(By.XPATH, ".//tr[td]")) > 0)

        headers = [h.text.strip() for h in table.find_elements(By.TAG_NAME, "th")]
        if not headers:
            first_row_th = table.find_elements(By.XPATH, ".//tr[th][1]//th")
            headers = [h.text.strip() for h in first_row_th]

        data: List[List[str]] = []
        for row in table.find_elements(By.TAG_NAME, "tr"):
            cols = row.find_elements(By.TAG_NAME, "td")
            if not cols:
                continue
            data.append([c.text.strip() for c in cols])

        if not headers or not data:
            logging.warning("table-class: headers/data missing for %s", report_url)
            _save_debug_artifacts(browser, "table-class-empty")
            return 0

        # Normalize widths
        norm_data: List[List[str]] = []
        for r in data:
            if len(r) < len(headers):
                r = r + [""] * (len(headers) - len(r))
            norm_data.append(r[: len(headers)])

        df = pd.DataFrame(
            norm_data,
            columns=[re.sub(r"\s+", " ", h).strip() for h in headers],
        )

        # Normalize column names (handles slight header variations)
        rename_map: Dict[str, str] = {}
        for c in df.columns:
            cc = str(c).strip()
            if "Transaction Date" in cc:
                rename_map[c] = "Transaction Date"
            elif "Owner" in cc:
                rename_map[c] = "Owner"
            elif "Ticker" in cc:
                rename_map[c] = "Ticker"
            elif "Asset Name" in cc:
                rename_map[c] = "Asset Name"
            elif "Asset Type" in cc:
                rename_map[c] = "Asset Type"
            elif cc == "Type" or "Type" in cc:
                rename_map[c] = "Type"
            elif "Amount" in cc:
                rename_map[c] = "Amount"
        df = df.rename(columns=rename_map)

        required = {"Transaction Date", "Owner", "Asset Name", "Type", "Amount"}
        if not required.issubset(set(df.columns)):
            logging.warning(
                "table-class: missing required cols for %s. Have=%s",
                report_url,
                list(df.columns),
            )
            _save_debug_artifacts(browser, "table-class-missing-cols")
            return 0

        rows_inserted = 0
        out_rows: List[Dict[str, object]] = []

        conn.execute("BEGIN")
        try:
            for _, r in df.iterrows():
                transaction_date_raw = _clean_text(r.get("Transaction Date", ""))
                transaction_date = _to_iso_date(transaction_date_raw)

                owner = _clean_text(r.get("Owner", ""))
                ticker = _clean_ticker(r.get("Ticker", ""))
                asset_name = _clean_text(r.get("Asset Name", ""))
                asset_type = _clean_text(r.get("Asset Type", ""))
                tx_type = _clean_text(r.get("Type", ""))
                amount_range = _clean_text(r.get("Amount", ""))

                trade_uid = _make_trade_uid(
                    [
                        report_url,
                        transaction_date,
                        owner,
                        ticker or "",
                        asset_name,
                        asset_type,
                        tx_type,
                        amount_range,
                    ]
                )

                cur = conn.execute(
                    """
                    INSERT OR IGNORE INTO transactions
                    (trade_uid, report_url, senator, owner, transaction_date, ticker, asset_name, asset_type,
                     transaction_type, amount_range, source_page, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        trade_uid,
                        report_url,
                        senator_full_name,
                        owner,
                        transaction_date,
                        ticker,
                        asset_name,
                        asset_type,
                        tx_type,
                        amount_range,
                        browser.current_url,
                        datetime.now().isoformat(),
                    ),
                )

                if cur.rowcount == 1:
                    rows_inserted += 1

                out_rows.append(
                    {
                        "trade_uid": trade_uid,
                        "Senator": senator_full_name,
                        "Transaction Date": transaction_date,
                        "Owner": owner,
                        "Ticker": ticker or "",
                        "Asset Name": asset_name,
                        "Asset Type": asset_type,
                        "Type": tx_type,
                        "Amount": amount_range,
                        "Report URL": report_url,
                    }
                )

            conn.commit()
        except Exception:
            conn.rollback()
            raise

        if excel_sheet and out_rows:
            save_to_excel(pd.DataFrame(out_rows), EXCEL_PATH, excel_sheet)

        return rows_inserted

    except Exception as exc:
        logging.error(
            "table-class: error extracting transactions from %s: %s", report_url, exc
        )
        _save_debug_artifacts(browser, "table-class-exception")
        return 0


# =============================================================================
# REPORT PROCESSING (senator01-style tab flow)
# =============================================================================


def process_report_link(
    browser: webdriver.Firefox,
    link: str,
    conn: sqlite3.Connection,
    attempts: int,
    download_images: bool,
    per_link_wait: float = 4.0,
    excel_sheet: Optional[str] = None,
) -> Tuple[bool, str]:
    last_err: Optional[str] = None

    for attempt in range(1, attempts + 1):
        # Guard: Firefox can sometimes lose its window in long runs.
        if not browser.window_handles:
            raise RuntimeError("Browser has no open windows (window_handles empty)")
        main_handle = browser.current_window_handle
        try:
            _open_in_new_tab(browser, link)

            # Ensure we left results page and content loaded
            WebDriverWait(browser, 25).until(lambda d: not _on_results_page(d))
            _wait_for_report_content(browser, timeout=25)

            # If we landed on /ptr/<uuid>/, click into the printable view first.
            # We prioritize the printable view because that's where the full transaction table lives.
            _try_open_print_view(browser)

            # Skip paper filings for now (OCR required)
            if "/search/view/paper/" in (browser.current_url or ""):
                conn.execute(
                    "INSERT OR REPLACE INTO reports (report_url, status, last_error, updated_at) VALUES (?, ?, ?, ?)",
                    (link, "paper_skipped", "OCR_required", datetime.now().isoformat()),
                )
                conn.commit()
                return False, "paper_skipped"

            state = detect_page_state(browser)

            if state == "VIEW_REPORT":
                # Skip this link type; return to results immediately.
                conn.execute(
                    "INSERT OR REPLACE INTO reports (report_url, status, last_error, updated_at) VALUES (?, ?, ?, ?)",
                    (link, "view_report_skipped", None, datetime.now().isoformat()),
                )
                conn.commit()
                return False, "view_report_skipped"

            if state == "REPORT_PAPER" or is_paper_filing(browser):
                conn.execute(
                    "INSERT OR REPLACE INTO reports (report_url, status, last_error, updated_at) VALUES (?, ?, ?, ?)",
                    (link, "paper_skipped", "OCR_required", datetime.now().isoformat()),
                )
                conn.commit()
                return False, "paper_skipped"

            # Give the page a moment to finish rendering dynamic bits (esp. headless)
            if per_link_wait and per_link_wait > 0:
                time.sleep(float(per_link_wait))
            # Give the page time to fully render (headless especially)
            time.sleep(4.0)

            # PRIORITY: senator01-style class="table" extraction
            rows_inserted = extract_html_transactions_table_class(
                browser, link, conn, excel_sheet=excel_sheet
            )

            # Fallback #1: DOM/header finder
            if rows_inserted == 0:
                rows_inserted = extract_html_transactions_dom(
                    browser, link, conn, excel_sheet=excel_sheet
                )

            # Fallback #2: pandas read_html (optional lxml)
            if rows_inserted == 0:
                rows_inserted = extract_html_transactions(
                    browser, link, conn, excel_sheet=excel_sheet
                )

            if rows_inserted > 0:
                conn.execute(
                    "INSERT OR REPLACE INTO reports (report_url, status, last_error, updated_at) VALUES (?, ?, ?, ?)",
                    (link, "html_processed", None, datetime.now().isoformat()),
                )
                conn.commit()
                return True, f"html_processed_{rows_inserted}_rows"

            last_err = "no_table"

        except Exception as exc:
            last_err = str(exc)
            logging.warning(
                "Attempt %s/%s failed for %s: %s", attempt, attempts, link, exc
            )
            time.sleep(2)

        finally:
            try:
                if len(browser.window_handles) > 1:
                    browser.close()
            except Exception:
                pass
            try:
                browser.switch_to.window(main_handle)
            except Exception:
                pass

    conn.execute(
        "INSERT OR REPLACE INTO reports (report_url, status, last_error, updated_at) VALUES (?, ?, ?, ?)",
        (link, "FAILED_NO_TABLE", last_err, datetime.now().isoformat()),
    )
    conn.commit()
    return False, "FAILED_NO_TABLE"


def collect_all_ptr_links(browser: webdriver.Firefox, max_pages: int) -> List[str]:
    """Collect PTR report hrefs across result pages by clicking Next."""
    seen: Set[str] = set()
    ordered: List[str] = []

    for _ in range(max(1, max_pages)):
        for el in _get_ptr_link_elements(browser):
            href = (el.get_attribute("href") or "").strip()
            if not href:
                continue
            # Focus only on electronically-filed PTR links
            if "/search/view/paper/" in href:
                continue
            if "/search/view/ptr/" not in href:
                continue
            if href not in seen:
                seen.add(href)
                ordered.append(href)

        if not _click_next_results_page(browser):
            break

        # Small pause to let the next page render
        time.sleep(1.0)

    return ordered


# =============================================================================
# EXCEL CONSOLIDATION (keep Master + Legend only)
# =============================================================================


def write_master_excel_from_sqlite(conn: sqlite3.Connection) -> None:
    """Write ONLY Master Transactions + Legend, generated from the durable SQLite store."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.read_sql_query(
        """
        SELECT
            trade_uid AS "Trade UID",
            senator AS "Senator",
            transaction_date AS "Transaction Date",
            owner AS "Owner",
            COALESCE(ticker, '') AS "Ticker",
            asset_name AS "Asset Name",
            asset_type AS "Asset Type",
            transaction_type AS "Type",
            amount_range AS "Amount",
            report_url AS "Report URL",
            created_at AS "Scraped At"
        FROM transactions
        ORDER BY transaction_date DESC, senator ASC
        """,
        conn,
    )

    if not df.empty:
        df.insert(0, "Unique ID", range(1, len(df) + 1))

    legend_df = pd.DataFrame(
        {
            "Field": [
                "Unique ID",
                "Trade UID",
                "Senator",
                "Transaction Date",
                "Owner",
                "Ticker",
                "Asset Name",
                "Asset Type",
                "Type",
                "Amount",
                "Report URL",
                "Scraped At",
            ],
            "Description": [
                "Sequential ID for Excel viewing",
                "Deterministic unique id (sha256) used to dedupe",
                "The filing senator name",
                "Date of transaction (ISO YYYY-MM-DD when available)",
                "Owner (self/spouse/dependent)",
                "Ticker symbol (blank if none)",
                "Asset name",
                "Asset type",
                "Purchase/Sale/etc.",
                "Reported amount bucket",
                "Source report link",
                "When the row was scraped",
            ],
        }
    )

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Master Transactions", index=False)
        legend_df.to_excel(writer, sheet_name="Legend", index=False)

    logging.info(
        "Wrote Excel: Master Transactions + Legend from SQLite (%s rows)", len(df)
    )


# =============================================================================
# CLI
# =============================================================================


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PTR extended Senate eFD scraper")
    parser.add_argument("--headless", action="store_true", help="Run in headless mode")
    parser.add_argument("--from", dest="from_date", help="Start date (MM/DD/YYYY)")
    parser.add_argument("--to", dest="to_date", help="End date (MM/DD/YYYY)")
    parser.add_argument("--month", help="Month to scrape (YYYY-MM)")
    parser.add_argument(
        "--max-pages", type=int, default=1, help="Maximum result pages to process"
    )
    parser.add_argument(
        "--max-reports",
        type=int,
        default=25,
        help="Maximum total reports to process (after collecting links). Hard-capped at 25.",
    )
    parser.add_argument(
        "--attempts", type=int, default=3, help="Retry attempts per report"
    )
    parser.add_argument(
        "--per-link-wait",
        type=float,
        default=4.0,
        help="Seconds to pause after opening each report tab before scraping (helps headless reliability).",
    )
    parser.add_argument(
        "--download-paper-images",
        action="store_true",
        default=True,
        help="Download paper filing images",
    )
    parser.add_argument(
        "--allow-custom-dates",
        action="store_true",
        help="Allow --from/--to/--month overrides. Default behavior is month-to-date (1st of month → today).",
    )
    return parser.parse_args()


def _resolve_dates(args: argparse.Namespace) -> Tuple[str, str]:
    """Resolve the date window for the Senate eFD search.

    Default (recommended for daily cron): month-to-date (1st of current month → today).
    If you *explicitly* pass `--allow-custom-dates`, then `--month` or `--from/--to` will override.
    """

    # Default: month-to-date (safe, repeatable daily runs)
    today = datetime.now()
    first_of_month = today.replace(day=1)
    default_from = first_of_month.strftime("%m/%d/%Y")
    default_to = today.strftime("%m/%d/%Y")

    if not getattr(args, "allow_custom_dates", False):
        return default_from, default_to

    # Optional overrides (only when explicitly enabled)
    if args.month:
        year, month = args.month.split("-")
        first_day = datetime(int(year), int(month), 1)
        if int(month) == 12:
            last_day = datetime(int(year) + 1, 1, 1) - timedelta(days=1)
        else:
            last_day = datetime(int(year), int(month) + 1, 1) - timedelta(days=1)
        return first_day.strftime("%m/%d/%Y"), last_day.strftime("%m/%d/%Y")

    if args.from_date and args.to_date:
        return args.from_date, args.to_date

    return default_from, default_to


def main() -> None:
    setup_logging()
    args = parse_args()
    from_date, to_date = _resolve_dates(args)

    conn = db_connect()
    db_init(conn)

    browser = create_driver(headless=args.headless)
    # Global alias for any legacy references that still use `driver`
    global driver
    driver = browser

    try:
        accept_terms_and_set_filters(browser, from_date, to_date)

        all_links = collect_all_ptr_links(browser, max_pages=args.max_pages)
        if not all_links:
            logging.warning(
                "No PTR report links found. The site layout may have changed."
            )

        total_stats = {
            "html_reports_scraped": 0,
            "html_rows_inserted": 0,
            "paper_skipped": 0,
            "view_report_skipped": 0,
            "reports_failed": 0,
        }

        # Hard cap to avoid runaway scraping during a single run
        limit = min(len(all_links), args.max_reports, 25)
        for i, link in enumerate(all_links[:limit], start=1):
            logging.info("Processing report %s/%s: %s", i, limit, link)

            success, status = process_report_link(
                browser,
                link,
                conn,
                attempts=args.attempts,
                download_images=args.download_paper_images,
                per_link_wait=args.per_link_wait,
                excel_sheet=None,
            )

            if status.startswith("html_processed"):
                total_stats["html_reports_scraped"] += 1
                rows_count = int(status.split("_")[2])
                total_stats["html_rows_inserted"] += rows_count
            elif status == "paper_skipped":
                total_stats["paper_skipped"] += 1
            elif status == "view_report_skipped":
                total_stats["view_report_skipped"] += 1
            else:
                total_stats["reports_failed"] += 1

        # Produce a clean workbook with ONLY Master Transactions + Legend
        write_master_excel_from_sqlite(conn)

        print("\n" + "=" * 60)
        print("SCRAPING SUMMARY")
        print("=" * 60)
        print(f"Date window: {from_date} → {to_date}")
        print(f"HTML Reports Scraped: {total_stats['html_reports_scraped']}")
        print(f"HTML Rows Inserted:  {total_stats['html_rows_inserted']}")
        print(f"Paper Reports Skipped (OCR required): {total_stats['paper_skipped']}")
        print(f"View Report Skipped:   {total_stats['view_report_skipped']}")
        print(f"Reports Failed:       {total_stats['reports_failed']}")
        print(f"SQLite DB:            {DB_PATH}")
        print(f"Excel:                {EXCEL_PATH}")
        print("=" * 60)

    finally:
        try:
            browser.quit()
        except Exception:
            pass
        try:
            conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
