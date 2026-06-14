# update_excel.py
# Scarica i nuovi prezzi ETF da GitHub e aggiunge le nuove righe
# in fondo allo sheet "Raw Data" di Portafoglio Vittorio PRO.xlsx.
# I 3 fondi (LU1883307461, IT0001080446, LU1883328467) vengono
# mantenuti con forward fill — aggiornali manualmente quando hai i NAV.
#
# USO: python update_excel.py
# REQUISITI: pip install pandas openpyxl requests

import requests
import pandas as pd
from io import StringIO
from openpyxl import load_workbook

# ── Configurazione ────────────────────────────────────────────────────────────
GITHUB_USER   = "VitAluigi"
GITHUB_REPO   = "Download-Prices"
GITHUB_BRANCH = "main"
CSV_PATH      = "data/ETF_Prices.csv"
CSV_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{CSV_PATH}"

EXCEL_PATH = r"C:\Users\valuigi\OneDrive - KPMG\Desktop\Worth\Portafoglio Vittorio PRO.xlsx"
SHEET_NAME = "Raw Data"

ETF_COLS  = ["IE00B4L5Y983", "LU3038520774", "IE00BMG6Z448",
             "IE00BZ0PKT83", "FR0013416716", "IE00BDBRDM35"]
FUND_COLS = ["LU1883307461", "IT0001080446", "LU1883328467"]

# ── 1. Scarica prezzi da GitHub ───────────────────────────────────────────────
print("1. Downloading ETF prices from GitHub...")
r = requests.get(CSV_URL, timeout=30)
r.raise_for_status()
new_prices = pd.read_csv(StringIO(r.text), index_col=0, parse_dates=True)
new_prices.index.name = "Date"
print(f"   {len(new_prices)} rows, last date: {new_prices.index.max().date()}")

# ── 2. Apri Excel e leggi Raw Data ────────────────────────────────────────────
print(f"\n2. Opening Excel...")
wb = load_workbook(EXCEL_PATH)
ws = wb[SHEET_NAME]

# Header (riga 1) → mappa ISIN: colonna 1-based
headers = [cell.value for cell in ws[1]]
col_idx = {h: i+1 for i, h in enumerate(headers) if h}

# Ultima riga con data e ultimo NAV dei fondi (0-based row tuple)
last_date = None
last_fund_vals = {}
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is not None:
        last_date = pd.Timestamp(row[0])
    for fund in FUND_COLS:
        if fund in col_idx:
            val = row[col_idx[fund]-1]  # col_idx 1-based → tuple 0-based
            if val is not None:
                last_fund_vals[fund] = val

print(f"   Last date in Raw Data: {last_date.date() if last_date else 'N/A'}")
print(f"   Last fund NAVs: { {k: round(v,4) for k,v in last_fund_vals.items()} }")

# ── 3. Trova nuove date ───────────────────────────────────────────────────────
new_dates = new_prices[new_prices.index > last_date] if last_date else new_prices

if new_dates.empty:
    print("\nNessuna nuova riga. Raw Data già aggiornato.")
    wb.close()
    exit()

print(f"\n3. Nuove righe: {len(new_dates)} ({new_dates.index.min().date()} → {new_dates.index.max().date()})")

# ── 4. Scrivi nuove righe in fondo ───────────────────────────────────────────
write_row = ws.max_row + 1
print(f"   Writing from row {write_row}...")

for date, prices in new_dates.iterrows():
    ws.cell(row=write_row, column=col_idx["Date"], value=date.date())

    for isin in ETF_COLS:
        if isin in col_idx and isin in prices and pd.notna(prices[isin]):
            ws.cell(row=write_row, column=col_idx[isin], value=round(float(prices[isin]), 6))

    for fund in FUND_COLS:
        if fund in col_idx and fund in last_fund_vals:
            ws.cell(row=write_row, column=col_idx[fund], value=last_fund_vals[fund])

    write_row += 1

# ── 5. Salva ──────────────────────────────────────────────────────────────────
wb.save(EXCEL_PATH)
print(f"\nDone. Salvato: {EXCEL_PATH}")
print(f"Righe aggiunte: {len(new_dates)} | Ultima data: {new_dates.index.max().date()}")
print("\nRicorda di aggiornare manualmente i NAV dei 3 fondi quando disponibili.")
