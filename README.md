# Spreadsheet to Documents

Turns each row of a spreadsheet into its own PDF or Word document and downloads them all as a ZIP.

**Use it:** https://pizzapup.github.io/spreadsheet-to-docs/

It runs entirely in your browser. Your files are not uploaded anywhere, and there is no backend, analytics or third-party script.

## How to use it

1. **Open a file.** Choose a CSV or `.xlsx` file. It appears as an editable table.
2. **Edit the data.** Fix values in the table, rename, reorder or hide columns, and exclude rows you don't want.
3. **Clean up values (optional).** Remove blanks, find and replace, trim spaces, and format numbers, dates and phone numbers.
4. **Set file names.** Build names from columns, e.g. `{Last name}-{First name}`.
5. **Choose output.** PDF, `.docx`, or both. You can add other files to the ZIP as-is.
6. **Download.** You get one ZIP with one document per row.

You can also export the edited data as CSV or XLSX.

## Supported files and limits

- CSV: UTF-8, comma, tab or semicolon delimited.
- Excel: `.xlsx` only. First worksheet, first row used as headers. Formulas are not recalculated. Legacy `.xls` is not supported.
- PDFs use a built-in font, so some non-Western characters may not display.
- Word output is plain paragraphs.
- Very large files are limited by browser memory.
- Not included: converting other document types, merging rows into one document.

## Run locally

```sh
python3 -m http.server 8000
```

Open http://localhost:8000. To run the tests (Node.js 20+):

```sh
npm test
```

## Deploy

`.github/workflows/pages.yml` runs the tests and publishes to GitHub Pages on every push to `main`. In the repo settings, set **Pages → Source** to **GitHub Actions**.

## Sample data

`data/data.csv` is made-up sample data you can open to try the app.
