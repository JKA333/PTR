#!/usr/bin/env python3
from __future__ import annotations

import argparse
import logging
import os
import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
from dotenv import load_dotenv

# Alpaca SDK (alpaca-py)
from alpaca.trading.client import TradingClient
from alpaca.trading.enums import OrderSide, TimeInForce
from alpaca.trading.requests import MarketOrderRequest


# -----------------------------
# Paths (match your project)
# -----------------------------
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
LOGS_DIR = BASE_DIR / "logs"

EXCEL_PATH = DATA_DIR / "PTR_Transactions.xlsx"
DB_PATH = DATA_DIR / "PTR.sqlite3"
TX_TABLE = "transactions"  # written by ptr.py
import hashlib


MASTER_SHEET = "Master Transactions"
TRADE_UID_COL = "Trade UID"

# Accept a few common UID column spellings from Excel
UID_CANDIDATES = [
    "Trade UID",
    "TRADE UID",
    "trade uid",
    "trade_uid",
    "TRADE_UID",
    "UID",
    "Uid",
    "uid",
]


# -----------------------------
# Logging
# -----------------------------
def setup_logging() -> None:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )


# -----------------------------
# Env / Alpaca
# -----------------------------
def load_env() -> None:
    """Load .env once and validate required Alpaca keys."""
    env_path = BASE_DIR / ".env"
    load_dotenv(dotenv_path=env_path, override=False)

    if not os.getenv("ALPACA_API_KEY") or not os.getenv("ALPACA_SECRET_KEY"):
        raise RuntimeError(
            "Missing Alpaca API credentials. Ensure app/.env has ALPACA_API_KEY and ALPACA_SECRET_KEY."
        )


def get_trading_client(paper: bool = True) -> TradingClient:
    """Return a TradingClient configured for paper or live trading.

    NOTE: `alpaca-py` already targets the Trading API v2 internally.
    If you set ALPACA_PAPER_URL to a URL that already ends with `/v2`,
    you can accidentally produce requests like `/v2/v2/account`.

    Safe values:
      - leave ALPACA_PAPER_URL unset (recommended)
      - or set it to: https://paper-api.alpaca.markets
    """

    def _sanitize_base_url(u: Optional[str]) -> Optional[str]:
        if not u:
            return None
        u = u.strip().rstrip("/")
        # Prevent double `/v2` (alpaca-py appends the v2 paths itself)
        if u.endswith("/v2"):
            u = u[: -len("/v2")]
            u = u.rstrip("/")
        return u or None

    key = os.getenv("ALPACA_API_KEY", "").strip()
    secret = os.getenv("ALPACA_SECRET_KEY", "").strip()

    base_url = _sanitize_base_url(os.getenv("ALPACA_PAPER_URL", "").strip() or None)

    kwargs = {"paper": paper}
    if base_url:
        # alpaca-py expects url_override (NOT base_url)
        kwargs["url_override"] = base_url

    return TradingClient(key, secret, **kwargs)


# -----------------------------
# SQLite state (so you don't double-trade)
# -----------------------------
def db_connect() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn


def db_init(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS executions (
            uid TEXT PRIMARY KEY,
            symbol TEXT NOT NULL,
            side TEXT NOT NULL,
            notional_usd REAL,
            qty REAL,
            alpaca_order_id TEXT,
            status TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        """
    )
    conn.commit()


def already_executed(conn: sqlite3.Connection, uid: str) -> bool:
    cur = conn.execute("SELECT 1 FROM executions WHERE uid = ? LIMIT 1;", (uid,))
    return cur.fetchone() is not None


def record_execution(
    conn: sqlite3.Connection,
    uid: str,
    symbol: str,
    side: str,
    status: str,
    alpaca_order_id: Optional[str] = None,
    notional_usd: Optional[float] = None,
    qty: Optional[float] = None,
) -> None:
    conn.execute(
        """
        INSERT OR REPLACE INTO executions
        (uid, symbol, side, notional_usd, qty, alpaca_order_id, status, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?);
        """,
        (
            uid,
            symbol,
            side,
            float(notional_usd) if notional_usd is not None else None,
            float(qty) if qty is not None else None,
            alpaca_order_id,
            status,
            datetime.now(timezone.utc).isoformat(),
        ),
    )
    conn.commit()


# -----------------------------
# Parsing helpers
# -----------------------------
_amount_re = re.compile(r"\$?([\d,]+)")


def parse_amount_range(amount_str: str) -> Tuple[Optional[float], Optional[float]]:
    """
    Examples:
      "$1,001 - $15,000"
      "$100,001 - $250,000"
      "Over $50,000,000"
    Returns (low, high). High may be None.
    """
    s = (amount_str or "").strip()
    if not s:
        return (None, None)

    if s.lower().startswith("over"):
        nums = _amount_re.findall(s)
        if nums:
            low = float(nums[0].replace(",", ""))
            return (low, None)
        return (None, None)

    nums = _amount_re.findall(s)
    if len(nums) >= 2:
        low = float(nums[0].replace(",", ""))
        high = float(nums[1].replace(",", ""))
        return (low, high)
    if len(nums) == 1:
        v = float(nums[0].replace(",", ""))
        return (v, v)

    return (None, None)


def conservative_notional(amount_str: str, cap_usd: float) -> Optional[float]:
    low, high = parse_amount_range(amount_str)
    if low is None:
        return None
    # conservative: use the LOW end, never exceed cap
    return float(min(low, cap_usd))


def normalize_symbol(sym: str) -> str:
    return (sym or "").strip().upper()


# -----------------------------
# Risk / Guards (minimal, safe defaults)
# -----------------------------
@dataclass
class RiskConfig:
    max_notional_per_trade: float = 50.0          # very conservative default
    max_trades_per_run: int = 3                   # throttle
    allow_sells: bool = False                     # start with buys-only
    require_ticker: bool = True                   # skip rows with blank tickers
    block_amount_over_usd: float = 5000.0         # if low-end is above this, skip
    dry_run: bool = True                          # default: no real orders


def should_trade_row(row: dict, risk: RiskConfig) -> Tuple[bool, str]:
    t = (row.get("Type") or "").lower()
    if "purchase" in t:
        side = "buy"
    elif "sale" in t:
        side = "sell"
    else:
        return (False, "type-not-buy-or-sell")

    if side == "sell" and not risk.allow_sells:
        return (False, "sells-disabled")

    sym = normalize_symbol(row.get("Ticker") or "")
    if risk.require_ticker and not sym:
        return (False, "missing-ticker")

    # Amount guard
    amt = row.get("Amount") or ""
    low, _ = parse_amount_range(amt)
    if low is None:
        return (False, "unknown-amount")
    if low > risk.block_amount_over_usd:
        return (False, f"amount-too-large(low={low})")

    return (True, "ok")


# -----------------------------
# Main execution
# -----------------------------

# -----------------------------
# SQLite Master Transaction Loader and Normalizer
# -----------------------------
def sqlite_table_exists(conn: sqlite3.Connection, name: str) -> bool:
    cur = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1;",
        (name,),
    )
    return cur.fetchone() is not None


def load_master_transactions_sqlite(conn: sqlite3.Connection) -> pd.DataFrame:
    """Load deduped PTR trades from SQLite (written by ptr.py) using Excel-compatible column names."""
    if not sqlite_table_exists(conn, TX_TABLE):
        raise RuntimeError(f"SQLite table not found: {TX_TABLE}")

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

    return df


def normalize_master_df(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure UID exists and key columns are consistently typed/cleaned."""
    if df is None or df.empty:
        return df

    # Prefer Trade UID (from ptr.py); otherwise fallback to any UID-like column.
    uid_col: Optional[str] = None
    for c in UID_CANDIDATES:
        if c in df.columns:
            uid_col = c
            break

    if uid_col:
        df["UID"] = df[uid_col]
    else:
        # Last-resort deterministic UID (keeps bot usable even if UID columns are missing)
        def _row_uid(r: pd.Series) -> str:
            parts = [
                str(r.get("Senator", "")),
                str(r.get("Transaction Date", "")),
                str(r.get("Owner", "")),
                str(r.get("Ticker", "")),
                str(r.get("Asset Name", "")),
                str(r.get("Asset Type", "")),
                str(r.get("Type", "")),
                str(r.get("Amount", "")),
                str(r.get("Report URL", "")),
            ]
            raw = "|".join(p.strip() for p in parts)
            return hashlib.sha256(raw.encode("utf-8")).hexdigest()

        df["UID"] = df.apply(_row_uid, axis=1)

    # Basic cleanup
    for col in ["UID", "Ticker", "Type", "Amount", "Senator", "Owner", "Transaction Date"]:
        if col in df.columns:
            df[col] = df[col].astype(str)

    df["UID"] = df["UID"].astype(str).str.strip()
    df = df[~df["UID"].isin({"", "nan", "None"})].copy()
    return df


def load_master_transactions() -> pd.DataFrame:
    """Load trades from Excel (fallback path)."""
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Excel not found: {EXCEL_PATH}")

    df = pd.read_excel(EXCEL_PATH, sheet_name=MASTER_SHEET, engine="openpyxl")
    return normalize_master_df(df)


def execute_orders(
    trading: TradingClient,
    conn: sqlite3.Connection,
    df: pd.DataFrame,
    risk: RiskConfig,
    paper: bool,
) -> None:
    traded = 0

    # quick connectivity check
    acct = trading.get_account()
    buying_power = float(acct.buying_power)
    logging.info("Alpaca connected. buying_power=%s paper=%s", buying_power, paper)

    # iterate newest first
    for _, r in df.iterrows():
        uid = str(r.get("UID") or "").strip()
        if not uid:
            continue

        if already_executed(conn, uid):
            continue

        row = r.to_dict()
        ok, reason = should_trade_row(row, risk)
        if not ok:
            record_execution(conn, uid, normalize_symbol(row.get("Ticker") or ""), "skip", reason)
            continue

        symbol = normalize_symbol(row.get("Ticker") or "")
        t = (row.get("Type") or "").lower()
        side = OrderSide.BUY if "purchase" in t else OrderSide.SELL

        # notional sizing (conservative)
        notional = conservative_notional(str(row.get("Amount") or ""), risk.max_notional_per_trade)
        if notional is None or notional <= 0:
            record_execution(conn, uid, symbol, side.value, "skip:notional")
            continue

        # affordability guard
        if notional > buying_power:
            record_execution(conn, uid, symbol, side.value, f"skip:insufficient-buying-power({buying_power})")
            continue

        if risk.dry_run:
            logging.info("[DRY RUN] %s %s notional=%s uid=%s", side.value.upper(), symbol, notional, uid)
            record_execution(conn, uid, symbol, side.value, "dry_run", notional_usd=notional)
        else:
            try:
                order = trading.submit_order(
                    MarketOrderRequest(
                        symbol=symbol,
                        side=side,
                        time_in_force=TimeInForce.DAY,
                        notional=notional,   # fractional where supported
                    )
                )
                logging.info("ORDER SUBMITTED: %s %s notional=%s id=%s", side.value.upper(), symbol, notional, order.id)
                record_execution(conn, uid, symbol, side.value, "submitted", alpaca_order_id=str(order.id), notional_usd=notional)
            except Exception as e:
                logging.error("Order failed for %s (%s): %s", symbol, uid, e)
                record_execution(conn, uid, symbol, side.value, f"error:{type(e).__name__}")

        traded += 1
        if traded >= risk.max_trades_per_run:
            logging.info("Reached max_trades_per_run=%s; stopping.", risk.max_trades_per_run)
            break


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Trade bot that mimics ptr.py Master Transactions (Alpaca paper).")
    ap.add_argument("--paper", action="store_true", help="Use Alpaca paper (recommended).")
    ap.add_argument("--execute", action="store_true", help="Actually place orders (default is dry-run).")
    ap.add_argument("--max-notional", type=float, default=50.0, help="Max USD notional per trade (default 50).")
    ap.add_argument("--max-trades", type=int, default=3, help="Max trades per run (default 3).")
    ap.add_argument("--allow-sells", action="store_true", help="Enable SELL orders (default off).")
    ap.add_argument("--block-amount-over", type=float, default=5000.0, help="Skip if low-end amount exceeds this.")
    return ap.parse_args()


def main() -> None:
    setup_logging()
    args = parse_args()

    load_env()

    risk = RiskConfig(
        max_notional_per_trade=args.max_notional,
        max_trades_per_run=args.max_trades,
        allow_sells=args.allow_sells,
        block_amount_over_usd=args.block_amount_over,
        dry_run=not args.execute,
    )

    conn = db_connect()
    db_init(conn)

    # Prefer SQLite (deduped source of truth written by ptr.py). Fallback to Excel.
    try:
        df = load_master_transactions_sqlite(conn)
        df = normalize_master_df(df)
        logging.info("Loaded %s rows from SQLite table '%s'.", len(df), TX_TABLE)
    except Exception as e:
        logging.warning("SQLite load failed (%s). Falling back to Excel.", e)
        df = load_master_transactions()
        logging.info("Loaded %s rows from Excel sheet '%s'.", len(df), MASTER_SHEET)

    if df.empty:
        logging.info("No rows available to trade.")
        return

    # Default is paper trading; pass --paper to be explicit.
    trading = get_trading_client(paper=True)
    execute_orders(trading, conn, df, risk, paper=True)

    logging.info("Done. Execution log stored in SQLite: %s", DB_PATH)


if __name__ == "__main__":
    main()
