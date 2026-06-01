# Loans & Interest

A web app that turns one or more QuickBooks **Transaction Reports** into a
single formula-driven workbook showing the **monthly interest on daily
balances** for each sub-account — with a shared **Loan Terms** tab and one
calculation tab per report.

The loan each report belongs to is **auto-detected from its contents**
(company name + account labels), so file names don't matter and nothing has
to be tagged by hand.

| Detected loan | Tab name | Default rate |
|---------------|----------|--------------|
| YS Affiliates notes receivable | `Affiliates` | 8.5% |
| Y&S Tickets → YS Affiliates | `YS Tix - YS Affiliates` | 12% |
| Mazel – Damona & Crew | `Mazel - Damona` | 13% |
| Mazel – Eitz Chaim / Ticket Vault | `Mazel Eitz Chaim, TV` | 13% |

Upload several reports together and they're combined into one workbook.

---

## How the interest is calculated

For every sub-account block in each report:

```
daily rate        = annual rate ÷ days in year
# days (per row)  = next transaction date − this date
                    (the last row of the month carries the balance to
                     month-end, so days always sum to the days in the month)
interest (row)    = daily rate × (sign × running balance) × # days
```

The `sign` is chosen automatically per tab so the **total interest is always
positive**, even when the running balances are negative (e.g. a payable to
Y&S Tickets). Everything in the output is a live Excel formula linked to the
`Loan Terms` tab.

### Loan Terms tab

Documents each rate block (annual rate, days in year, daily rate) and its
posting notes — for example the 12% block carries the Y&S Tickets / YS
Affiliates debit-credit instructions, and the 13% block the "Damona & crew /
Mazel Investing" note. All cells are plain black with no highlighting.

---

## Run locally

```bash
pip install -r requirements.txt
python app.py          # http://localhost:8000
```

Command line (single file, loan auto-detected):

```bash
python interest_calc.py input_report.xlsx output.xlsx
```

## Deploy on Railway

1. Push this folder to a GitHub repo.
2. Railway → New Project → Deploy from GitHub repo → pick it.
3. It installs `requirements.txt` and starts with the `Procfile` command
   (`gunicorn app:app`). Railway supplies `$PORT` automatically.

---

## How detection works

Each report's text (company in A1, account labels, names, memos) is scanned:

- "eitz chaim" / "ticket vault" / "subnotes" → **Mazel Eitz Chaim, TV**
- "damona" → **Mazel - Damona**
- "y&s tickets" / "tickets" / "tix" → **YS Tix - YS Affiliates**
- "affiliates" / "notes receivable" → **Affiliates**

In the UI each file shows an **Auto-detect** dropdown you can override if a
report is ever misclassified.

## Notes

- Rates are editable in the UI (Affiliates / YS Tix / Mazel) plus days-in-year
  (365; use 360 for a 30/360 convention).
- Tab names never contain a rate. Duplicate loan types get a numeric suffix.
- Max upload size is 50 MB; accepts `.xlsx` / `.xlsm`.

## Project layout

```
ys-interest-calculator/
├── app.py              # Flask routes: /, /process (multi-file), /download, /health
├── interest_calc.py    # Parser + detection + interest math + workbook generator
├── templates/index.html
├── sample_data/
├── requirements.txt
├── Procfile
├── railway.json
├── runtime.txt
└── .gitignore
```
