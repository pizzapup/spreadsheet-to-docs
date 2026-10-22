# Spreadsheet to Documents

Turns each row of a spreadsheet into its own PDF or Word document and downloads them as a ZIP. It runs entirely in the browser as a static site.

## Features

- Open a CSV or `.xlsx` file and edit it in a table.
- Rename, reorder, hide and format columns. Exclude rows.
- Clean up values: blanks, find and replace, trimming, number/date/phone formats.
- Build file names from columns, e.g. `{Last name}-{First name}`.
- Output PDF, `.docx`, or both, in one ZIP. Other files can be added to the ZIP as-is.
- Export the edited data as CSV or XLSX.

## Data handling

Files are read and converted in the browser. Nothing is uploaded, and there is no backend or analytics. The site loads no third-party scripts.

## Formats and limits

- CSV: UTF-8, comma, tab or semicolon delimited.
- Excel: `.xlsx` only, first worksheet, first row as headers. Formulas are not recalculated. Legacy `.xls` is not supported.
- PDFs use a built-in font, so some non-Western characters may not display.
- Word output is plain paragraphs.
- Very large files are limited by browser memory.
- Not included: converting other documents, merging documents into one.

## Run locally

```sh
python3 -m http.server 8000
```

Then open http://localhost:8000. Tests need Node.js 20+:

```sh
npm test
```

## Deploy

`.github/workflows/pages.yml` runs the tests and publishes to GitHub Pages on every push to `main`. In the repo settings, set **Pages → Source** to **GitHub Actions**.

`data/data.csv` is made-up sample data.
