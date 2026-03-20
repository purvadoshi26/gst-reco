# GST Audit Reconciliation Tool

A free web app for GST audit reconciliation — no installation needed.

## Features

- **GSTR-2B vs Tally** — ITC Reconciliation (party-wise)
- **GSTR-1 vs Books** — Sales Reconciliation (invoice-wise)
- Auto-detects column names (Tally / SAP / any Excel format)
- Colour-coded Excel output download

## Deploy on Streamlit Cloud (Free)

1. Fork this repo on GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub account
4. Select this repo → `app.py` → Deploy

## File Format Requirements

### ITC Reco
- File 1: GSTR-2B Excel from portal (needs `B2B` sheet)
- File 2: Tally Purchase Register (Detailed export)

### Sales Reco
- File 1: E-Invoice Excel (`b2b, sez, de` sheet) or any GSTR-1 Excel
- File 2: Tally multi-sheet OR any flat Excel with `Invoice No` + `Taxable Value`
