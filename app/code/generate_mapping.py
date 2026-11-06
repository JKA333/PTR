from pathlib import Path
import os
import csv

from dotenv import load_dotenv
from alpaca.trading.client import TradingClient

# ------------------------
# Load .env safely
# ------------------------
BASE_DIR = Path(__file__).resolve().parents[1]
ENV_PATH = BASE_DIR / ".env"

load_dotenv(dotenv_path=ENV_PATH)

API_KEY = os.getenv("ALPACA_API_KEY")
API_SECRET = os.getenv("ALPACA_SECRET_KEY")

if not API_KEY or not API_SECRET:
    raise RuntimeError("Missing Alpaca API credentials in .env")

# ------------------------
# Connect to Alpaca (PAPER)
# ------------------------
client = TradingClient(
    api_key=API_KEY,
    secret_key=API_SECRET,
    paper=True,
)

# ------------------------
# Fetch assets
# ------------------------
assets = client.get_all_assets()

tradable_assets = [
    a for a in assets
    if a.tradable and a.exchange in {"NYSE", "NASDAQ", "ARCA"}
]

print(f"Fetched {len(tradable_assets)} tradable assets")

# ------------------------
# Save mapping
# ------------------------
output_dir = BASE_DIR / "data"
output_dir.mkdir(exist_ok=True)

output_file = output_dir / "ticker_map.csv"

with open(output_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow([
        "symbol",
        "name",
        "exchange",
        "asset_class",
        "status",
        "shortable",
        "fractionable",
    ])

    for a in tradable_assets:
        writer.writerow([
            a.symbol,
            a.name,
            a.exchange,
            a.asset_class,
            a.status,
            a.shortable,
            a.fractionable,
        ])

print(f"Saved ticker map → {output_file}")