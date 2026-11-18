# Spreadsheet to Documents

Turns each row of a spreadsheet into its own PDF or Word document and downloads them all as a ZIP.

**Your files stay on your computer.** Everything runs in your browser, so nothing is uploaded, tracked or shared with third parties.

## Where to use it

| Option | How |
| --- | --- |
| **Online** | Open https://pizzapup.github.io/spreadsheet-to-docs/ |
| **On your computer** | Download the [latest release](https://github.com/pizzapup/spreadsheet-to-docs/releases/latest) or clone the repo, then follow [Run locally](#run-locally) below. |

## How to use it

1. **Open a file.** Choose a CSV or `.xlsx` file. It appears as an editable table.
2. **Edit the data.** Fix values in the table, rename, reorder or hide columns, and exclude rows you don't want.
3. **Clean up values (optional).** Remove blanks, find and replace, trim spaces, and format numbers, dates and phone numbers.
4. **Set file names.** Build names from columns, e.g. `{Last name}-{First name}`.
5. **Choose output.** PDF, `.docx`, or both. You can add other files to the ZIP as-is.
6. **Download.** You get one ZIP with one document per row.

You can also export the edited data as CSV or XLSX. `data/data.csv` is made-up sample data you can open to try the app.

## Data handling

The code makes no network calls with your data. The only `fetch` loads the bundled sample CSV from the same site. There's no cookie or browser storage use, no beacons, and no external URLs apart from XML namespaces. You can confirm this in `app.js` or in your browser's Network tab.

<details>
<summary><strong>Supported files and limits</strong></summary>

<br>

- CSV: UTF-8, comma, tab or semicolon delimited.
- Excel: `.xlsx` only. First worksheet, first row used as headers. Formulas are not recalculated. Legacy `.xls` is not supported.
- PDFs use a built-in font, so some non-Western characters may not display.
- Word output is plain paragraphs.
- Very large files are limited by browser memory.
- Not included: converting other document types, merging rows into one document.

</details>

<details id="run-locally">
<summary><strong>Run locally</strong></summary>

<br>

There is nothing to install. Serve the folder with any static file server, for example:

```sh
python3 -m http.server 8000
```

Open http://localhost:8000. To run the tests (Node.js 20+):

```sh
npm test
```

</details>

<details>
<summary><strong>Deploy your own copy</strong></summary>

<br>

`.github/workflows/pages.yml` runs the tests and publishes to GitHub Pages on every push to `main`. In the repo settings, set **Pages → Source** to **GitHub Actions**.

</details>
