# update_prices.py
# Scarica i prezzi giornalieri dei 6 ETF da Yahoo Finance
# e aggiorna data/ETF_Prices.csv nel repository.
# Viene eseguito automaticamente da GitHub Actions ogni sera.

import yfinance as yf
import pandas as pd
import datetime
from pathlib import Path

# ── Configurazione ────────────────────────────────────────────────────────────
ISIN_TICKER_MAP = {
    "IE00B4L5Y983": "SWDA.MI",    # iShares Core MSCI World
    "LU3038520774": "DEFS.PA",    # Amundi Stoxx Europe Defense
    "IE00BMG6Z448": "EXCH.MI",    # iShares MSCI EM ex-China
    "IE00BZ0PKT83": "IFSW.MI",    # iShares STOXX World Multifactor
    "FR0013416716": "GOLD.AS",    # Amundi Physical Gold ETC
    "IE00BDBRDM35": "AGGH.MI",    # iShares Core Global Aggregate Bond
}

START_DATE = "2020-01-01"
OUTPUT_PATH = Path("data/ETF_Prices.csv")

# ── Load existing data ────────────────────────────────────────────────────────
OUTPUT_PATH.parent.mkdir(exist_ok=True)

if OUTPUT_PATH.exists():
    existing = pd.read_csv(OUTPUT_PATH, index_col=0, parse_dates=True)
    # Scarica solo dall'ultimo giorno disponibile
    start = existing.index.max().strftime("%Y-%m-%d")
    print(f"Existing data: {len(existing)} rows, last date: {start}")
else:
    existing = pd.DataFrame()
    start = START_DATE
    print(f"No existing data, downloading from {start}")

# ── Download new data ─────────────────────────────────────────────────────────
end = (datetime.date.today() + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
tickers = list(ISIN_TICKER_MAP.values())

print(f"Downloading {len(tickers)} ETFs from {start} to {end}...")
raw = yf.download(tickers, start=start, end=end, auto_adjust=True, progress=False)

if raw.empty:
    print("No new data from Yahoo Finance.")
else:
    new = raw["Close"].copy()
    # Rinomina da ticker a ISIN
    ticker_to_isin = {v: k for k, v in ISIN_TICKER_MAP.items()}
    new.rename(columns=ticker_to_isin, inplace=True)
    new.index.name = "Date"
    new.index = pd.to_datetime(new.index)

    # Unisci con storico esistente
    combined = pd.concat([existing, new])
    combined = combined[~combined.index.duplicated(keep="last")].sort_index()

    # Salva
    combined.to_csv(OUTPUT_PATH)
    print(f"Saved {len(combined)} rows to {OUTPUT_PATH}")
    print(f"New rows added: {len(combined) - len(existing)}")
    print(combined.tail(3).to_string())
