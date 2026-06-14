# update_excel.py
# Scarica ETF_Prices.csv dal repository GitHub e aggiorna
# il file Excel Portafoglio_Vittorio_PRO.xlsx con i nuovi prezzi.
# I 3 fondi (LU1883307461, IT0001080446, LU1883328467) vengono mantenuti
# così come sono — aggiornali manualmente quando scarichi i NAV dai siti.
#
# USO:
#   python update_excel.py
#
# REQUISITI:
#   pip install pandas openpyxl requests

import pandas as pd
import requests
from pathlib import Path
from io import StringIO
import datetime

# ── Configurazione ────────────────────────────────────────────────────────────
# URL raw del CSV su GitHub (aggiorna con il tuo username/repo)
GITHUB_USER  = "VitAluigi"
GITHUB_REPO  = "ETF-Prices"          # nome della nuova repo
GITHUB_BRANCH = "main"
CSV_PATH_IN_REPO = "data/ETF_Prices.csv"

CSV_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{CSV_PATH_IN_REPO}"

# Percorso locale del tuo Excel
EXCEL_PATH = r"C:\Users\valuigi\OneDrive - KPMG\Desktop\Worth\Portafoglio_Vittorio_PRO.xlsx"
SHEET_NAME = "Raw Data"

# ISIN dei 3 fondi che aggiorni manualmente (non vengono toccati)
MANUAL_FUNDS = ["LU1883307461", "IT0001080446", "LU1883328467"]

# ── Download ETF prices from GitHub ──────────────────────────────────────────
print(f"Downloading ETF prices from GitHub...")
r = requests.get(CSV_URL, timeout=30)
r.raise_for_status()
new_prices = pd.read_csv(StringIO(r.text), index_col=0, parse_dates=True)
new_prices.index.name = "Date"
print(f"Downloaded: {len(new_prices)} rows, last date: {new_prices.index.max().date()}")

# ── Load existing Excel ───────────────────────────────────────────────────────
print(f"\nLoading Excel: {EXCEL_PATH}")
existing = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
existing["Date"] = pd.to_datetime(existing["Date"])
existing = existing.set_index("Date").sort_index()

last_excel_date = existing.index.max()
print(f"Excel last date: {last_excel_date.date()}")

# ── Find new rows to add ──────────────────────────────────────────────────────
etf_cols = [c for c in new_prices.columns if c not in MANUAL_FUNDS]
new_rows = new_prices[new_prices.index > last_excel_date][etf_cols]

if new_rows.empty:
    print("\nNessuna nuova riga da aggiungere. Excel già aggiornato.")
else:
    print(f"\nNuove righe da aggiungere: {len(new_rows)}")
    print(new_rows.to_string())

    # Per le nuove righe, i fondi manuali restano NaN (li aggiorni tu)
    for fund in MANUAL_FUNDS:
        new_rows[fund] = float("nan")

    # Calcola i pesi correnti dell'ultima riga valida (per scalare i prezzi in valore €)
    # NOTA: il tuo Excel contiene valori in € (non prezzi), quindi devi scalare
    # il prezzo per il numero di quote che possiedi.
    # Se hai le quote salvate da qualche parte aggiornale qui.
    # Per ora aggiungiamo i PREZZI grezzi — poi moltiplica per le quote manualmente.
    print("\nATTENZIONE: le nuove righe contengono prezzi ETF, non valori €.")
    print("Moltiplica per il numero di quote possedute per avere il valore in portafoglio.")

    # Aggiungi colonna Total (solo ETF, fondi NaN)
    new_rows["Total"] = float("nan")  # verrà calcolato nel tuo Excel

    # Unisci
    updated = pd.concat([existing, new_rows]).sort_index()
    updated = updated[~updated.index.duplicated(keep="last")]

    # ── Save back to Excel ────────────────────────────────────────────────────
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        updated.reset_index().to_excel(writer, sheet_name=SHEET_NAME, index=False)

    print(f"\nExcel aggiornato: {EXCEL_PATH}")
    print(f"Righe totali: {len(updated)}")
    print(f"Ultima data: {updated.index.max().date()}")

print("\nFatto.")
