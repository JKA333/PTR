import os
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv
from alpaca.trading.client import TradingClient

# Load environment
load_dotenv()

API_KEY = os.getenv("ALPACA_API_KEY")
API_SECRET = os.getenv("ALPACA_SECRET_KEY")

if not API_KEY or not API_SECRET:
    raise RuntimeError("Missing Alpaca API credentials")

client = TradingClient(
    api_key=API_KEY,
    secret_key=API_SECRET,
    paper=True
)

# Fetch all assets
assets = client.get_all_assets()

rows = []

for a in assets:
    if (
        a.tradable
        and a.status == "active"
        and a.asset_class == "us_equity"
    ):
        rows.append({
            "Symbol": a.symbol,
            "Name": a.name,
            "Exchange": a.exchange,
            "Tradable": a.tradable,
            "Shortable": a.shortable,
            "Fractionable": a.fractionable,
            "Easy to Borrow": a.easy_to_borrow,
        })

df = pd.DataFrame(rows).sort_values("Symbol")

print(f"Eligible symbols: {len(df)}")

# Write to Excel Legend sheet
EXCEL_PATH = Path(__file__).resolve().parent.parent / "data" / "PTR_Transactions.xlsx"

with pd.ExcelWriter(
    EXCEL_PATH,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace",
) as writer:
    df.to_excel(writer, sheet_name="Legend", index=False)

print("✅ Legend sheet updated with eligible Alpaca symbols")
