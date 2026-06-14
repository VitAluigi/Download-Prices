# ETF Prices - Auto-updater

Aggiornamento automatico giornaliero dei prezzi ETF tramite GitHub Actions.

## Struttura

```
|-- .github/
│   |-- workflows/
│       |-- update_prices.yml   # GitHub Action (18:30 IT)
|-- data/
│   |-- ETF_Prices.csv
|-- update_prices.py
|-- update_excel.py
|__ README.md
```

## Come funziona

1. **GitHub Actions** esegue `update_prices.py` ogni giorno feriale alle 17:30 UTC
2. Il CSV `data/ETF_Prices.csv` viene aggiornato e committato automaticamente
3. Scarichi i nuovi prezzi eseguendo `update_excel.py` sul tuo PC

## Aggiornare l'Excel localmente

```bash
pip install pandas openpyxl requests
python update_excel.py
```
