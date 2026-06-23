# JELD-WEN BOM Converter

Convert costed Bill of Materials PDFs into editable Excel workbooks with live formulas. Includes:

- A **browser tool** for GitHub Pages (drag & drop, no server needed)
- A **Python CLI** for local batch conversion

## GitHub Pages

1. Push this repository to GitHub.
2. Open **Settings → Pages**.
3. Under **Build and deployment**, set **Source** to **Deploy from a branch**.
4. Choose the `main` branch and the `/ (root)` folder.
5. Save. After a minute or two your site will be live at:

   `https://<your-username>.github.io/<repo-name>/`

Anyone with the link can open the page, drop a BOM PDF, and download the Excel file.

## Local web preview

```bash
python3 -m http.server 8080
```

Then open [http://localhost:8080](http://localhost:8080).

## Python CLI

```bash
pip install -r requirements.txt
python3 bom_to_excel.py path/to/bom.pdf
```

## What the Excel output includes

- Horizontal header with item metadata
- Line items grouped by category (CHEM, HDWR, PKG, etc.)
- `Extended Cost = Qty × Unit Cost` formulas on each row
- Section totals and a cost summary panel on the right
- BOY columns removed
- Final **TOTAL COST** as a formula

## Privacy

The web tool runs entirely in the browser. PDFs are not sent to any server.
