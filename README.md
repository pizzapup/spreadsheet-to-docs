# Spreadsheet to Documents

Turns each row of a spreadsheet into its own PDF or Word document and downloads them all as a ZIP.

> [!NOTE]
> Your files stay on your computer. Everything runs in your browser, so nothing is uploaded, tracked or shared with third parties.

## Where to use it

| Option | How |
| --- | --- |
| Online | Open https://pizzapup.github.io/spreadsheet-to-docs/ |
| On your computer | Download the [latest release](https://github.com/pizzapup/spreadsheet-to-docs/releases/latest) or clone the repo, then see [Run locally](#run-locally). |

## How it works

| Step | What you do | Details |
| --- | --- | --- |
| 1. Open | Choose a CSV or `.xlsx` file | It appears as an editable table. |
| 2. Edit | Fix values, rename, reorder or hide columns, exclude rows | Optional cleanup: remove blanks, find and replace, trim spaces, format numbers, dates and phone numbers. |
| 3. Name | Set file names from columns | For example `{Last name}-{First name}`. |
| 4. Output | Pick PDF, `.docx`, or both | Other files can be added to the ZIP as-is. |
| 5. Download | Get one ZIP | One document per row. |

> [!TIP]
> `data/data.csv` is made-up sample data you can open to try the app. You can also export the edited data as CSV or XLSX.

## Supported files and limits

- CSV: UTF-8, comma, tab or semicolon delimited.
- Excel: `.xlsx` only. First worksheet, first row used as headers. Formulas are not recalculated. Legacy `.xls` is not supported.
- PDFs use a built-in font, so some non-Western characters may not display.
- Word output is plain paragraphs.
- Very large files are limited by browser memory.
- Not included: converting other document types, merging rows into one document.

## Data handling

The code makes no network calls with your data. The only `fetch` loads the bundled sample CSV from the same site. There's no cookie or browser storage use, no beacons, and no external URLs apart from XML namespaces. You can confirm this in `app.js` or in your browser's Network tab.

## Run locally

There is nothing to install. Serve the folder with any static file server, for example:

```sh
python3 -m http.server 8000
```

Open http://localhost:8000. To run the tests (Node.js 20+):

```sh
npm test
```

## Deploy your own copy

`.github/workflows/pages.yml` runs the tests and publishes to GitHub Pages on every push to `main`. In the repo settings, set Pages → Source to GitHub Actions.
