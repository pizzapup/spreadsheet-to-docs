import assert from "node:assert/strict";
import test from "node:test";
import { classifyValue, columnInsights, conversionReport, detectColumnType, formatValue, parseDate, parseNumber, parsePhone, resolveSpec, similarValues, formatBool, parseBool } from "../format.js";
import { DEFAULT_NAME_OPTIONS, blankRule, createXlsx, replaceValue, toCsv, filenameTokens, resolveFilename, assignFilenames, createPdf, createDocx, createZip, inspectFilenameTemplate, isBlank, parseCsv, sanitizeFilename } from "../app.js";

test("parses CSV headers, quoted commas, escaped quotes, and line breaks", () => {
  const result = parseCsv('\uFEFFName,Description,Note\r\n"Lee, Sam","two\nlines","say ""hello"""\r\nMina,,done');
  assert.deepEqual(result.headers, ["Name", "Description", "Note"]);
  assert.deepEqual(result.rows, [
    { Name: "Lee, Sam", Description: "two\nlines", Note: 'say "hello"' },
    { Name: "Mina", Description: "", Note: "done" },
  ]);
});

test("detects semicolon-separated CSV and makes duplicate headings unique", () => {
  const result = parseCsv("Name;Name;Name (2);Amount\nA;B;C;1");
  assert.deepEqual(result.headers, ["Name", "Name (2)", "Name (2) (2)", "Amount"]);
  assert.equal(result.rows[0]["Name (2)"], "B");
  assert.equal(result.rows[0]["Name (2) (2)"], "C");
});

test("rejects malformed and header-only CSV files", () => {
  assert.throws(() => parseCsv('Name,Value\n"A,B'), /unclosed quoted field/);
  assert.throws(() => parseCsv("Name,Value"), /header row and at least one data row/);
});

test("builds a PDF with correct page objects and valid xref offsets", () => {
  const pdf = new TextDecoder().decode(createPdf(["Name: Lee (Sam)", "Town: Québec"]));
  assert.match(pdf, /^%PDF-1\.4/);
  assert.ok(pdf.includes("\\(Sam\\)"));
  assert.match(pdf, /Qu\\351bec/);
  assert.match(pdf, /\/Count 1/);
  const xrefAt = Number(pdf.match(/startxref\n(\d+)/)?.[1]);
  assert.equal(pdf.slice(xrefAt, xrefAt + 4), "xref");
  const entries = [...pdf.matchAll(/^(\d{10}) 00000 n /gm)].map((match) => Number(match[1]));
  assert.ok(entries.length >= 4);
  for (const offset of entries) assert.match(pdf.slice(offset), /^\d+ 0 obj/);
});

test("splits long PDF output across pages", () => {
  const pdf = new TextDecoder().decode(createPdf(Array.from({ length: 51 }, (_, index) => `Row ${index + 1}`)));
  assert.match(pdf, /\/Count 2/);
  assert.equal((pdf.match(/\/Type \/Page /g) ?? []).length, 2);
});

test("creates a readable ZIP with UTF-8 filenames and intact PDF data", () => {
  const entries = [
    { name: "A.pdf", data: createPdf(["First"]) },
    { name: "José.pdf", data: createPdf(["Second"]) },
  ];
  const archive = createZip(entries);
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
  let endOffset = -1;
  for (let offset = archive.length - 22; offset >= 0; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      endOffset = offset;
      break;
    }
  }
  assert.notEqual(endOffset, -1);
  assert.equal(view.getUint16(endOffset + 10, true), entries.length);
  let cursor = view.getUint32(endOffset + 16, true);
  const decoder = new TextDecoder();
  for (const expected of entries) {
    assert.equal(view.getUint32(cursor, true), 0x02014b50);
    const filenameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const expectedName = decoder.decode(archive.subarray(cursor + 46, cursor + 46 + filenameLength));
    assert.equal(expectedName, expected.name);
    assert.equal(view.getUint32(cursor + 16, true), crc32(expected.data));
    const localOffset = view.getUint32(cursor + 42, true);
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    const size = view.getUint32(cursor + 24, true);
    assert.deepEqual(archive.subarray(dataOffset, dataOffset + size), expected.data);
    cursor += 46 + filenameLength + extraLength + commentLength;
  }
});

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

test("sanitizes download filenames and uses a fallback for empty names", () => {
  assert.equal(sanitizeFilename("../Sales: 2026?.zip", "Documents"), ".._Sales_ 2026_.zip");
  assert.equal(sanitizeFilename("   ", "Documents"), "Documents");
});

test("checks filename values as a complete template rather than rejecting repeated columns", () => {
  const rows = [
    { First: "Alex", Last: "Lee", Code: "A/1" },
    { First: "Alex", Last: "Chen", Code: "B" },
    { First: "Taylor", Last: "Chen", Code: "" },
  ];
  const combined = inspectFilenameTemplate(["First", "Last", "Code"], rows, "{First}-{Last}");
  assert.equal(combined.collisions, 0);
  assert.equal(combined.columns.find((column) => column.name === "First").unique, 2);
  const repeated = inspectFilenameTemplate(["First", "Last", "Code"], rows, "{First}");
  assert.equal(repeated.collisions, 1);
  const report = inspectFilenameTemplate(["First", "Last", "Code"], rows, "{Code}-{Unknown}");
  assert.deepEqual(report.unknown, ["Unknown"]);
  assert.equal(report.columns.find((column) => column.name === "Code").special, 1);
  assert.equal(report.columns.find((column) => column.name === "Code").blanks, 1);
  const sanitizedCollision = inspectFilenameTemplate(
    ["Code"],
    [{ Code: "A/B" }, { Code: "A:B" }],
    "{Code}",
  );
  assert.equal(sanitizedCollision.collisions, 1);
  assert.equal(sanitizedCollision.special, 2);
  const longName = inspectFilenameTemplate(
    ["Code"],
    [{ Code: "A".repeat(260) }],
    "{Code}",
  );
  assert.equal(longName.long, 1);
});

test("assignFilenames numbers repeated names by mode", () => {
  const names = ["Smith", "Jones", "Smith", "Smith"];
  assert.deepEqual(assignFilenames(names), ["Smith", "Jones", "Smith-2", "Smith-3"]);
  assert.deepEqual(assignFilenames(names, "number-repeats"), ["Smith-1", "Jones", "Smith-2", "Smith-3"]);
  assert.deepEqual(assignFilenames(names, "all", "_"), ["Smith_1", "Jones_1", "Smith_2", "Smith_3"]);
  assert.deepEqual(assignFilenames(["A", "A-2", "A"]), ["A", "A-2", "A-3"]);
  assert.deepEqual(assignFilenames(names, "none"), names);
});

test("number formatting changes only numeric values", () => {
  const spec = { number: "$#,##0.00" };
  assert.equal(formatValue("64500", spec), "$64,500.00");
  assert.equal(formatValue("1,234.5", spec), "$1,234.50");
  assert.equal(formatValue("(12.5)", spec), "-$12.50");
  assert.equal(formatValue("0.125", { number: "0.0%" }), "12.5%");
  assert.equal(formatValue("12.5%", { number: "0.00" }), "0.13");
  assert.equal(formatValue("Sample row", spec), "Sample row");
  assert.equal(formatValue("02134", spec), "02134");
  assert.equal(formatValue("", spec), "");
  assert.equal(parseNumber("12abc"), null);
});

test("date formatting reads month/day/year and ISO dates", () => {
  assert.deepEqual(parseDate("3/2/2026"), { year: 2026, month: 3, day: 2 });
  assert.equal(parseDate("2/30/2026"), null);
  assert.equal(formatValue("2026-03-02", { date: "mm/dd/yyyy" }), "03/02/2026");
  assert.equal(formatValue("Mar 2, 2026", { date: "yyyy-mm-dd" }), "2026-03-02");
  assert.equal(formatValue("2026-03-02", { date: "mmmm d, yyyy" }), "March 2, 2026");
  assert.equal(formatValue("2026-03-02", { date: "dd-mmm-yy" }), "02-Mar-26");
  assert.equal(formatValue("2026", { date: "mm/dd/yyyy" }), "2026");
});

test("cell formats override column formats, which override the dataset", () => {
  const formats = {
    dataset: { number: "0" },
    columns: { Amount: { number: "0.00" } },
    cells: { "1|Amount": { number: "$#,##0" } },
  };
  assert.equal(resolveSpec(formats, "Other", 0).number, "0");
  assert.equal(resolveSpec(formats, "Amount", 0).number, "0.00");
  assert.equal(resolveSpec(formats, "Amount", 1).number, "$#,##0");
});

test("blank detection covers spaces and invisible characters", () => {
  assert.equal(isBlank(""), true);
  assert.equal(isBlank("  \u00a0 "), true);
  assert.equal(isBlank("\u200b"), true);
  assert.equal(isBlank(undefined), true);
  assert.equal(isBlank("0"), false);
});

test("filename column inspection lists the values behind each warning", () => {
  const rows = [{ Name: "Smith" }, { Name: "Smith" }, { Name: " " }, { Name: "A/B" }];
  const column = inspectFilenameTemplate(["Name"], rows, "{Name}").columns[0];
  assert.deepEqual(column.blankRows, [3]);
  assert.deepEqual(column.repeatedValues, [{ value: "Smith", count: 2 }]);
  assert.deepEqual(column.specialValues, ["A/B"]);
});

test("parseCsv reports load stats and drops empty rows and unnamed empty columns", () => {
  const { headers, rows, stats } = parseCsv("Name,,Note\nA,,x\n,,\nA,,x\nB,,");
  assert.deepEqual(headers, ["Name", "Note"]);
  assert.equal(rows.length, 3);
  assert.deepEqual(stats, { rows: 3, columns: 2, emptyRowsDropped: 1, emptyColumnsDropped: 1, duplicateRows: 1, blankCells: 1 });
});

test("parses {number}, {index} and limited column tokens", () => {
  const tokens = filenameTokens("{Name(4)}-{number}-{index}-{Other}", ["Name"]);
  assert.deepEqual(tokens.map((t) => t.kind), ["column", "number", "number", "unknown"]);
  assert.equal(tokens[0].limit, 4);
  assert.equal(tokens[0].column, "Name");
});

test("resolves filenames with limit, trim, spaces and number", () => {
  const options = { trim: true, spaces: "underscore", unsafe: "_", find: "", replaceWith: "", numbering: "repeats", separator: "-", numberStart: 5, numberPad: 3 };
  const name = resolveFilename("{Name(5)}-{number}", ["Name"], { Name: "  Mary Ann Lee " }, 0, options);
  assert.equal(name, "Mary_-005");
});

test("sanitizeFilename uses a custom replacement and numbering mode none keeps names", () => {
  assert.equal(sanitizeFilename("a/b", "x", "-"), "a-b");
  assert.deepEqual(assignFilenames(["A", "A"], "none"), ["A", "A"]);
});

test("replaceValue handles whole-cell, partial and case options", () => {
  assert.equal(replaceValue("T", "T", "True"), "True");
  assert.equal(replaceValue("Tx", "T", "True"), null);
  assert.equal(replaceValue("t", "T", "True", { matchCase: false }), "True");
  assert.equal(replaceValue("a-b-c", "-", "_", { whole: false }), "a_b_c");
  assert.equal(replaceValue("A.b", ".", "!", { whole: false, matchCase: false }), "A!b");
  assert.equal(replaceValue("x", "", "y"), null);
});

test("blankRule prefers a column override over the dataset default", () => {
  const meta = { blank: { dataset: { mode: "omit", value: "" }, columns: { A: { mode: "fill", value: "N/A" } } } };
  assert.equal(blankRule(meta, "A").mode, "fill");
  assert.equal(blankRule(meta, "B").mode, "omit");
  assert.equal(blankRule({}, "B").mode, "keep");
});

test("toCsv quotes fields and uses a BOM and CRLF", () => {
  assert.equal(toCsv([["a", "b,c"], ['say "hi"', "x\ny"]]), '\uFEFFa,"b,c"\r\n"say ""hi""","x\ny"\r\n');
});

test("createXlsx builds a zip with escaped cell text", () => {
  const bytes = createXlsx([["Name", "Note"], ["Lee <1>", "a & b"]]);
  assert.deepEqual([...bytes.slice(0, 2)], [0x50, 0x4b]);
  const text = new TextDecoder("latin1").decode(bytes);
  assert.ok(text.includes("xl/worksheets/sheet1.xml"));
  assert.ok(text.includes("Lee &lt;1&gt;"));
  assert.ok(text.includes("a &amp; b"));
});


test("parses long-form dates and timestamps as dates", () => {
  for (const value of ["2026-01-12 14:30:00", "2026-01-12T14:30:00Z", "1/12/2026 2:30 PM", "Monday, January 12, 2026", "2026/01/12"]) {
    assert.deepEqual(parseDate(value), { year: 2026, month: 1, day: 12 }, value);
  }
  assert.equal(formatValue("2026-01-12 14:30:00", { date: "mm/dd/yyyy" }), "01/12/2026");
});

test("formats phone numbers and leaves unrecognised values alone", () => {
  assert.equal(parsePhone("+1 (555) 010-2004"), "5550102004");
  assert.equal(formatValue("555.010.2004", { phone: "(###) ###-####" }), "(555) 010-2004");
  assert.equal(formatValue("call me", { phone: "###-###-####" }), "call me");
  assert.equal(parsePhone("12345"), null);
});

test("detects column types", () => {
  assert.equal(classifyValue("a@b.co"), "email");
  assert.equal(classifyValue("45%"), "percentage");
  assert.equal(classifyValue("$1,200.50"), "currency");
  assert.equal(detectColumnType(["Yes", "No", "yes"]).type, "bool");
  assert.equal(detectColumnType(["2026-01-01", "Jan 5, 2026", "1/2/2026 3:04 PM"]).type, "date");
  assert.equal(detectColumnType(["$10", "20", "30"]).type, "currency");
});

test("flags incomplete addresses", () => {
  const found = detectColumnType(["1 A St, Town, IL 62701", "2 B St, Town, IL 62702", "3 C St, Town, IL 62703", "62701"]);
  assert.equal(found.type, "address");
  assert.deepEqual(found.incompleteAddresses, [{ row: 4, value: "62701" }]);
});

test("column insights summarize values and numbers", () => {
  const info = columnInsights(["10", "20", "30", "", "10"]);
  assert.equal(info.empty, 1);
  assert.equal(info.numeric.min, 10);
  assert.equal(info.numeric.max, 30);
  assert.equal(info.top[0].value, "10");
  assert.equal(info.top[0].count, 2);
});

test("conversion report lists unconverted values", () => {
  const report = conversionReport(["2026-01-01", "soon", "", "soon"], "date");
  assert.equal(report.ok, 1);
  assert.equal(report.failedCount, 2);
  assert.deepEqual(report.failed, ["soon"]);
});

test("filename options support rules, capitalization and a default limit", () => {
  const options = { ...DEFAULT_NAME_OPTIONS, rules: [{ find: "-", to: "_" }], caps: "upper", limit: 5 };
  assert.equal(resolveFilename("{Name}", ["Name"], { Name: "ab-cd-ef" }, 0, options), "AB_CD");
  const title = { ...DEFAULT_NAME_OPTIONS, caps: "title" };
  assert.equal(resolveFilename("{Name}", ["Name"], { Name: "jANE doe" }, 0, title), "Jane Doe");
});

test("createDocx builds a ZIP containing the document text", () => {
  const bytes = createDocx(["Name: A & B"]);
  assert.equal(String.fromCharCode(bytes[0], bytes[1]), "PK");
  const text = new TextDecoder().decode(bytes);
  assert.ok(text.includes("word/document.xml"));
  assert.ok(text.includes("Name: A &amp; B"));
});

test("assignFilenames supports number position, start and digits", () => {
  assert.deepEqual(assignFilenames(["A", "A"], "all", "-", { position: "before", start: 5, pad: 2 }), ["05-A", "06-A"]);
  assert.deepEqual(assignFilenames(["A", "A"], "repeats", "_", { start: 1 }), ["A", "A_2"]);
});

test("trims edge spaces by default and lets a column opt out", () => {
  const formats = { dataset: {}, columns: { Keep: { trim: false } }, cells: {} };
  assert.equal(formatValue("  a  ", resolveSpec(formats, "Other", 0)), "a");
  assert.equal(formatValue("  a  ", resolveSpec(formats, "Keep", 0)), "  a  ");
});

test("suggests spelling fixes for rare values close to common ones", () => {
  const values = ["Birmingham", "Birmingham", "Birmingham", "Birmigham", "Boston", "Boston", "boston", "Austin"];
  const found = similarValues(values);
  assert.deepEqual(found.find((item) => item.from === "Birmigham"), { from: "Birmigham", to: "Birmingham", count: 1 });
  assert.equal(found.find((item) => item.from === "boston").to, "Boston");
  assert.equal(found.some((item) => item.from === "Austin"), false);
});

test("converts boolean-like values", () => {
  assert.equal(formatBool(parseBool("T"), "Yes/No"), "Yes");
  assert.equal(formatBool(parseBool("false"), "1/0"), "0");
  assert.equal(parseBool("maybe"), null);
  assert.equal(formatValue("t", { bool: "True/False" }), "True");
  assert.equal(formatValue("n", { bool: "True/False" }), "False");
});

test("date columns return a timeline histogram", () => {
  const info = columnInsights(["2026-01-01", "2026-01-02", "2026-03-01", "2026-03-01"], { name: "Joined" });
  assert.equal(info.dates.earliest, "2026-01-01");
  assert.equal(info.dates.latest, "2026-03-01");
  assert.equal(info.dates.bins.reduce((n, bin) => n + bin.count, 0), 4);
});
