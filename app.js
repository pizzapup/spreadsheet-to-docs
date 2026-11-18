import { PHONE_PRESETS, isPhoneName, parsePhoneParts, parsePhone, validPhonePattern, classifyValue, detectColumnType, columnInsights, TYPE_LABELS,
  DATE_PRESETS, NUMBER_PRESETS, BOOL_PRESETS, validBoolPattern, parseBool, similarValues, emptyFormats, formatValue, parseDate, parseNumber,
  resolveSpec, validDatePattern, validNumberPattern,
} from "./format.js";

const INVALID_FILENAME_CHARACTERS = /[<>:"/\\|?*\u0000-\u001f]/g;
let outputKind = "pdf";
let currentTable = null;
let originalTable = null;
let undoStates = [];
let redoStates = [];
let highlightCells = new Set();
let highlightHeaders = new Set();
let selectedColumns = new Set();
let selectionAnchor = null;
let highlightBlanks = new Set();
let previewPosition = 0;
let rowLimit = 500;
let selectedRows = new Set();
let rowAnchor = null;
let focusedRow = null;
let selectMode = false;
let loadStats = null;
const PENCIL = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zm17.71-10.21a1 1 0 0 0 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>';

export function parseCsv(input) {
  const text = input.replace(/^\uFEFF/, "");
  const firstLine = text.split(/\r?\n/, 1)[0] ?? "";
  const delimiter = detectDelimiter(firstLine);
  const records = [];
  let record = [];
  let value = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          value += '"';
          index += 1;
        } else {
          quoted = false;
        }
      } else {
        value += character;
      }
    } else if (character === '"' && value.length === 0) {
      quoted = true;
    } else if (character === delimiter) {
      record.push(value);
      value = "";
    } else if (character === "\n" || character === "\r") {
      record.push(value);
      records.push(record);
      record = [];
      value = "";
      if (character === "\r" && text[index + 1] === "\n") index += 1;
    } else {
      value += character;
    }
  }
  if (quoted) throw new Error("The CSV file has an unclosed quoted field.");
  if (record.length > 0 || value.length > 0) {
    record.push(value);
    records.push(record);
  }

  return tableFromRows(records);
}

function detectDelimiter(line) {
  let best = ",";
  let bestCount = -1;
  let quoted = false;
  for (const candidate of [",", "\t", ";"]) {
    let count = 0;
    quoted = false;
    for (let index = 0; index < line.length; index += 1) {
      if (line[index] === '"') {
        if (quoted && line[index + 1] === '"') index += 1;
        else quoted = !quoted;
      } else if (line[index] === candidate && !quoted) {
        count += 1;
      }
    }
    if (count > bestCount) {
      best = candidate;
      bestCount = count;
    }
  }
  return best;
}

function tableFromRows(rows) {
  const nonemptyRows = rows.filter((row) => row.some((cell) => cell.trim() !== ""));
  const emptyRowsDropped = rows.length - nonemptyRows.length;
  if (nonemptyRows.length < 2) {
    throw new Error("The spreadsheet needs a header row and at least one data row.");
  }
  const width = nonemptyRows.reduce((largest, row) => Math.max(largest, row.length), 0);
  const rawHeaders = Array.from({ length: width }, (_, index) =>
    (nonemptyRows[0][index] ?? "").trim() || `Column ${index + 1}`,
  );
  const usedHeaders = new Set();
  const headerCounts = new Map();
  const unnamed = new Set(
    rawHeaders.map((_, index) => index).filter((index) =>
      (nonemptyRows[0][index] ?? "").trim() === "" && nonemptyRows.slice(1).every((row) => (row[index] ?? "").trim() === ""),
    ),
  );
  const keptIndexes = rawHeaders.map((_, index) => index).filter((index) => !unnamed.has(index));
  const headers = keptIndexes.map((index) => rawHeaders[index]).map((header) => {
    let count = headerCounts.get(header) ?? 0;
    let candidate;
    do {
      count += 1;
      candidate = count === 1 ? header : `${header} (${count})`;
    } while (usedHeaders.has(candidate));
    headerCounts.set(header, count);
    usedHeaders.add(candidate);
    return candidate;
  });
  const data = nonemptyRows.slice(1).map((row) =>
    Object.fromEntries(headers.map((header, index) => [header, row[keptIndexes[index]] ?? ""])),
  );
  const seenRows = new Set();
  let duplicateRows = 0;
  for (const row of data) {
    const key = JSON.stringify(headers.map((header) => row[header]));
    if (seenRows.has(key)) duplicateRows += 1;
    seenRows.add(key);
  }
  const blankCells = data.reduce(
    (total, row) => total + headers.filter((header) => isBlank(row[header])).length,
    0,
  );
  return {
    headers,
    rows: data,
    stats: {
      rows: data.length,
      columns: headers.length,
      emptyRowsDropped,
      emptyColumnsDropped: unnamed.size,
      duplicateRows,
      blankCells,
    },
  };
}

export const DEFAULT_NAME_OPTIONS = {
  trim: true,
  spaces: "keep",
  unsafe: "_",
  find: "",
  replaceWith: "",
  numbering: "repeats",
  separator: "-",
  numberStart: 1,
  numberPad: 1,
  rules: [],
  caps: "keep",
  limit: 0,
};

export const MAX_BASE_LENGTH = 250;

export function parseFilenameToken(raw, headers) {
  if (headers.includes(raw)) return { kind: "column", column: raw, limit: 0 };
  if (raw === "number" || raw === "index") return { kind: "number" };
  const match = /^(.*)\((\d+)\)$/.exec(raw);
  if (match && headers.includes(match[1])) return { kind: "column", column: match[1], limit: Number(match[2]) };
  return { kind: "unknown" };
}

export function filenameTokens(template, headers) {
  return [...template.matchAll(/\{([^{}]+)\}/g)].map((match) => ({
    raw: match[1],
    start: match.index,
    end: match.index + match[0].length,
    ...parseFilenameToken(match[1], headers),
  }));
}

const EDGE_SPACE = /^[\s\u200b-\u200d\u2060\ufeff]+|[\s\u200b-\u200d\u2060\ufeff]+$/g;

function cleanValue(value, options, limit) {
  let text = String(value ?? "");
  let trimmed = false;
  if (options.trim) {
    const result = text.replace(EDGE_SPACE, "");
    trimmed = result !== text;
    text = result;
  }
  if (options.find) text = text.split(options.find).join(options.replaceWith);
  for (const rule of options.rules ?? []) {
    if (rule.find) text = text.split(rule.find).join(rule.to ?? "");
  }
  if (options.spaces !== "keep") {
    text = text.replace(/\s+/g, { hyphen: "-", underscore: "_", remove: "" }[options.spaces] ?? " ");
  }
  if (options.caps === "upper") text = text.toLocaleUpperCase();
  else if (options.caps === "lower") text = text.toLocaleLowerCase();
  else if (options.caps === "title") text = text.toLocaleLowerCase().replace(/(^|[\s_-])(\p{L})/gu, (_, gap, letter) => gap + letter.toLocaleUpperCase());
  const max = limit || options.limit;
  if (max) text = Array.from(text).slice(0, max).join("");
  return { text, trimmed };
}

function resolveFilenameDetailed(template, headers, row, sequence, options = DEFAULT_NAME_OPTIONS) {
  let trimmed = false;
  const text = template
    .replace(/\{([^{}]+)\}/g, (_, raw) => {
      const token = parseFilenameToken(raw, headers);
      if (token.kind === "number") {
        return String(Math.max(0, Number(options.numberStart) || 0) + sequence).padStart(Math.max(1, Number(options.numberPad) || 1), "0");
      }
      if (token.kind === "column") {
        const result = cleanValue(row[token.column], options, token.limit);
        trimmed ||= result.trimmed;
        return result.text;
      }
      return "";
    })
    .replace(/\.pdf$/i, "");
  return { text, trimmed };
}

export function resolveFilename(template, headers, row, sequence, options = DEFAULT_NAME_OPTIONS) {
  return resolveFilenameDetailed(template, headers, row, sequence, options).text;
}

export function inspectFilenameTemplate(headers, rows, template, options = DEFAULT_NAME_OPTIONS) {
  const tokens = filenameTokens(template, headers);
  const placeholders = tokens.map((token) => token.raw);
  const unknown = [...new Set(tokens.filter((token) => token.kind === "unknown").map((token) => token.raw))];
  const columns = [...new Set(tokens.filter((token) => token.kind === "column").map((token) => token.column))]
    .map((name) => inspectFilenameColumn(name, rows));
  const details = rows.map((row, index) => resolveFilenameDetailed(template, headers, row, index, options));
  const filenames = details.map((detail) => detail.text);
  const replacement = options.unsafe ?? "_";
  const safeNames = filenames.map((name, index) =>
    Array.from(sanitizeFilename(name, `Document-${index + 1}`, replacement)).slice(0, MAX_BASE_LENGTH).join(""),
  );
  const counts = new Map();
  for (const name of safeNames) {
    const normalized = name.toLocaleLowerCase();
    counts.set(normalized, (counts.get(normalized) ?? 0) + 1);
  }
  const collisions = [...counts.values()].reduce((total, count) => total + Math.max(0, count - 1), 0);
  return {
    placeholders,
    unknown,
    columns,
    collisions,
    trimmed: details.filter((detail) => detail.trimmed).length,
    special: filenames.filter((name) => UNSAFE_FILENAME.test(name)).length,
    long: filenames.filter((name) => Array.from(name).length > MAX_BASE_LENGTH).length,
  };
}

export function assignFilenames(bases, mode = "repeats", separator = "-", { position = "after", start = 1, pad = 1 } = {}) {
  if (mode === "none") return [...bases];
  const totals = new Map();
  for (const base of bases) {
    const key = base.toLocaleLowerCase();
    totals.set(key, (totals.get(key) ?? 0) + 1);
  }
  const counters = new Map();
  const used = new Set();
  return bases.map((base) => {
    const key = base.toLocaleLowerCase();
    const seen = (counters.get(key) ?? 0) + 1;
    counters.set(key, seen);
    let number = start + seen - 1;
    const join = (value) => {
      const digits = String(value).padStart(Math.max(1, pad), "0");
      return position === "before" ? `${digits}${separator}${base}` : `${base}${separator}${digits}`;
    };
    const numbered = mode === "all" || (mode === "number-repeats" && totals.get(key) > 1) ||
      (mode === "repeats" && seen > 1);
    let name = numbered ? join(number) : base;
    while (used.has(name.toLocaleLowerCase())) {
      number += 1;
      name = join(number);
    }
    used.add(name.toLocaleLowerCase());
    return name;
  });
}

export function isBlank(value) {
  return /^[\s\u200b-\u200d\u2060\ufeff]*$/.test(String(value ?? ""));
}

export const DEFAULT_BLANK = { mode: "keep", value: "N/A" };

export function blankRule(meta, header) {
  return meta?.blank?.columns?.[header] ?? meta?.blank?.dataset ?? DEFAULT_BLANK;
}

export function replaceValue(value, find, replacement, { whole = true, matchCase = true } = {}) {
  const text = String(value ?? "");
  if (!find) return null;
  let result;
  if (whole) {
    const same = matchCase ? text === find : text.toLocaleLowerCase() === find.toLocaleLowerCase();
    result = same ? replacement : text;
  } else if (matchCase) {
    result = text.split(find).join(replacement);
  } else {
    result = text.replace(new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"), () => replacement);
  }
  return result === text ? null : result;
}

function csvField(value) {
  const text = String(value ?? "");
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

export function toCsv(matrix) {
  return `\uFEFF${matrix.map((row) => row.map(csvField).join(",")).join("\r\n")}\r\n`;
}

function columnLetters(index) {
  let name = "";
  for (let number = index + 1; number > 0; number = Math.floor((number - 1) / 26)) {
    name = String.fromCharCode(65 + ((number - 1) % 26)) + name;
  }
  return name;
}

function xmlEscape(value) {
  return String(value).replace(/[&<>"\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, (character) => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[character] ?? ""
  ));
}

export function createXlsx(matrix) {
  const encoder = new TextEncoder();
  const sheetRows = matrix.map((row, rowIndex) => {
    const cells = row.map((value, columnIndex) =>
      `<c r="${columnLetters(columnIndex)}${rowIndex + 1}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value ?? "")}</t></is></c>`,
    ).join("");
    return `<row r="${rowIndex + 1}">${cells}</row>`;
  }).join("");
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const files = {
    "[Content_Types].xml": `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    "_rels/.rels": `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `${head}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": `${head}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetRows}</sheetData></worksheet>`,
  };
  return createZip(Object.entries(files).map(([name, text]) => ({ name, data: encoder.encode(text) })));
}

export function createDocx(lines) {
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const paragraphs = lines
    .map((line) => `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(line)}</w:t></w:r></w:p>`)
    .join("");
  const files = {
    "[Content_Types].xml": `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    "_rels/.rels": `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/document.xml": `${head}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  const encoder = new TextEncoder();
  return createZip(Object.entries(files).map(([name, text]) => ({ name, data: encoder.encode(text) })));
}

const UNSAFE_FILENAME = /[<>:"/\\|?*\u0000-\u001f]/;

function inspectFilenameColumn(name, rows) {
  const values = rows.map((row) => String(row[name] ?? ""));
  const blankRows = [];
  const counts = new Map();
  const specialValues = [];
  const longRows = [];
  values.forEach((value, index) => {
    if (isBlank(value)) {
      blankRows.push(index + 1);
      return;
    }
    counts.set(value, (counts.get(value) ?? 0) + 1);
    if (UNSAFE_FILENAME.test(value) && !specialValues.includes(value)) specialValues.push(value);
    if (value.length > MAX_BASE_LENGTH) longRows.push(index + 1);
  });
  const repeatedValues = [...counts].filter(([, count]) => count > 1).map(([value, count]) => ({ value, count }));
  return {
    name,
    blanks: blankRows.length,
    unique: counts.size,
    rows: values.length,
    special: values.filter((value) => UNSAFE_FILENAME.test(value)).length,
    long: values.filter((value) => value.length > MAX_BASE_LENGTH).length,
    shortest: values.length ? Math.min(...values.map((value) => value.length)) : 0,
    longest: values.length ? Math.max(...values.map((value) => value.length)) : 0,
    blankRows,
    repeatedValues,
    specialValues,
    longRows,
  };
}

function filenameColumnSummary(column) {
  const warnings = [];
  if (column.blanks) warnings.push(`${column.blanks} blank`);
  const repeated = column.rows - column.blanks - column.unique;
  if (repeated > 0) warnings.push(`${repeated} repeated`);
  if (column.special) warnings.push(`${column.special} unsafe`);
  if (column.long) warnings.push(`${column.long} long`);
  return warnings.length ? warnings.join(", ") : "no filename issues";
}

function filenameColumnDetails(column) {
  const quote = (value) => `"${value.length > 30 ? `${value.slice(0, 30)}…` : value}"`;
  const more = (total, shown) => (total > shown ? `, +${total - shown} more` : "");
  const details = [];
  if (column.blankRows.length) {
    details.push(`Blank rows: ${column.blankRows.slice(0, 10).join(", ")}${more(column.blankRows.length, 10)}`);
  }
  if (column.repeatedValues.length) {
    details.push("Repeated values:");
    column.repeatedValues.slice(0, 8).forEach(({ value, count }) => details.push(`  (×${count}) ${quote(value)}`));
    if (column.repeatedValues.length > 8) details.push(`  +${column.repeatedValues.length - 8} more`);
  }
  if (column.specialValues.length) {
    details.push(`Unsafe characters: ${column.specialValues.slice(0, 5).map(quote).join(", ")}${more(column.specialValues.length, 5)}`);
  }
  if (column.longRows.length) {
    details.push(`Over ${MAX_BASE_LENGTH} characters, rows: ${column.longRows.slice(0, 10).join(", ")}${more(column.longRows.length, 10)}`);
  }
  return details;
}

export async function parseXlsx(buffer) {
  const archive = await readZip(buffer);
  const workbookXml = parseXml(await readZipText(archive, "xl/workbook.xml"));
  const relationshipsXml = parseXml(await readZipText(archive, "xl/_rels/workbook.xml.rels"));
  const firstSheet = workbookXml.getElementsByTagNameNS("*", "sheet")[0];
  if (!firstSheet) throw new Error("The Excel workbook has no worksheets.");
  const relationId = firstSheet.getAttribute("r:id");
  const relationship = [...relationshipsXml.getElementsByTagNameNS("*", "Relationship")]
    .find((item) => item.getAttribute("Id") === relationId);
  if (!relationship) throw new Error("Could not find the first worksheet in this workbook.");

  const target = relationship.getAttribute("Target");
  const sheetPath = target.startsWith("/")
    ? target.slice(1)
    : normalizeArchivePath(`xl/${target}`);
  const sheetXml = parseXml(await readZipText(archive, sheetPath));
  const sharedStrings = archive.has("xl/sharedStrings.xml")
    ? readSharedStrings(parseXml(await readZipText(archive, "xl/sharedStrings.xml")))
    : [];
  const sheetData = sheetXml.getElementsByTagNameNS("*", "sheetData")[0];
  if (!sheetData) throw new Error("The first worksheet contains no tabular data.");

  const tableRows = [...sheetData.getElementsByTagNameNS("*", "row")].map((row) => {
    const values = [];
    for (const cell of row.getElementsByTagNameNS("*", "c")) {
      const column = columnNumber(cell.getAttribute("r"));
      if (column < 0) continue;
      const type = cell.getAttribute("t");
      const valueNode = cell.getElementsByTagNameNS("*", "v")[0];
      let value = valueNode?.textContent ?? "";
      if (type === "s") value = sharedStrings[Number(value)] ?? "";
      else if (type === "inlineStr") {
        value = [...cell.getElementsByTagNameNS("*", "t")]
          .map((node) => node.textContent ?? "").join("");
      } else if (type === "b") value = value === "1" ? "TRUE" : "FALSE";
      values[column] = value;
    }
    return Array.from({ length: values.length }, (_, index) => values[index] ?? "");
  });
  return tableFromRows(tableRows);
}

async function readZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const minimumOffset = Math.max(0, bytes.length - 22 - 0xffff);
  let endOffset = -1;
  for (let offset = bytes.length - 22; offset >= minimumOffset; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      endOffset = offset;
      break;
    }
  }
  if (endOffset < 0) throw new Error("This is not a valid XLSX workbook.");
  const entries = view.getUint16(endOffset + 10, true);
  let cursor = view.getUint32(endOffset + 16, true);
  const decoder = new TextDecoder();
  const files = new Map();

  for (let index = 0; index < entries; index += 1) {
    if (view.getUint32(cursor, true) !== 0x02014b50) {
      throw new Error("The XLSX archive has an invalid directory.");
    }
    const method = view.getUint16(cursor + 10, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const filenameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const localOffset = view.getUint32(cursor + 42, true);
    const name = decoder.decode(bytes.subarray(cursor + 46, cursor + 46 + filenameLength));
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.subarray(dataOffset, dataOffset + compressedSize);
    let contents;
    if (method === 0) contents = compressed.slice();
    else if (method === 8) {
      if (typeof DecompressionStream === "undefined") {
        throw new Error("This browser cannot decompress Excel files. Try a current version of Chrome, Edge, Firefox, or Safari.");
      }
      try {
        const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
        contents = new Uint8Array(await new Response(stream).arrayBuffer());
      } catch {
        throw new Error("Could not decompress this Excel file. Try a current browser or save the workbook again as .xlsx.");
      }
    } else {
      throw new Error("This Excel file uses an unsupported compression method.");
    }
    files.set(name, contents);
    cursor += 46 + filenameLength + extraLength + commentLength;
  }
  return files;
}

async function readZipText(archive, path) {
  const file = archive.get(path);
  if (!file) throw new Error(`The Excel workbook is missing ${path}.`);
  return new TextDecoder().decode(file);
}

function parseXml(source) {
  const xml = new DOMParser().parseFromString(source, "application/xml");
  if (xml.getElementsByTagName("parsererror").length > 0) {
    throw new Error("The Excel workbook contains invalid XML.");
  }
  return xml;
}

function readSharedStrings(xml) {
  return [...xml.getElementsByTagNameNS("*", "si")].map((item) =>
    [...item.getElementsByTagNameNS("*", "t")]
      .map((node) => node.textContent ?? "").join(""),
  );
}

function columnNumber(reference) {
  const letters = /^([A-Z]+)/i.exec(reference)?.[1]?.toUpperCase();
  if (!letters) return -1;
  return [...letters].reduce((column, letter) => column * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

function normalizeArchivePath(path) {
  const parts = [];
  for (const part of path.split("/")) {
    if (part === "..") parts.pop();
    else if (part && part !== ".") parts.push(part);
  }
  return parts.join("/");
}

export function sanitizeFilename(name, fallback, replacement = "_") {
  const safe = name.replace(INVALID_FILENAME_CHARACTERS, replacement).replace(/[. ]+$/g, "").trim();
  return safe || fallback;
}

export function createPdf(lines) {
  const outputLines = [];
  for (const line of lines) {
    const words = String(line).split(/\s+/);
    let current = "";
    for (const word of words) {
      if (word.length > 86) {
        if (current) outputLines.push(current);
        current = "";
        for (let offset = 0; offset < word.length; offset += 86) {
          const part = word.slice(offset, offset + 86);
          if (part.length === 86) outputLines.push(part);
          else current = part;
        }
      } else if (current && current.length + word.length + 1 > 86) {
        outputLines.push(current);
        current = word;
      } else {
        current = current ? `${current} ${word}` : word;
      }
    }
    outputLines.push(current);
  }

  const pageLines = [];
  for (let index = 0; index < outputLines.length; index += 50) {
    pageLines.push(outputLines.slice(index, index + 50));
  }
  if (pageLines.length === 0) pageLines.push([""]);
  const objects = [];
  const pageIds = pageLines.map((_, index) => 4 + index * 2);
  objects[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  objects[2] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;
  objects[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>";

  pageLines.forEach((linesForPage, index) => {
    const pageId = pageIds[index];
    const contentId = pageId + 1;
    const commands = ["BT", "/F1 11 Tf", "52 790 Td", "14 TL"];
    for (const [lineIndex, line] of linesForPage.entries()) {
      if (lineIndex > 0) commands.push("T*");
      commands.push(`(${pdfString(line)}) Tj`);
    }
    commands.push("ET");
    const content = commands.join("\n");
    objects[pageId] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 842] /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentId} 0 R >>`;
    objects[contentId] = `<< /Length ${content.length} >>\nstream\n${content}\nendstream`;
  });

  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (let id = 1; id < objects.length; id += 1) {
    offsets[id] = pdf.length;
    pdf += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let id = 1; id < objects.length; id += 1) {
    pdf += `${String(offsets[id]).padStart(10, "0")} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return new TextEncoder().encode(pdf);
}

function pdfString(value) {
  const cp1252 = new Map([
    [0x20ac, 0x80], [0x201a, 0x82], [0x0192, 0x83], [0x201e, 0x84], [0x2026, 0x85],
    [0x2020, 0x86], [0x2021, 0x87], [0x02c6, 0x88], [0x2030, 0x89], [0x0160, 0x8a],
    [0x2039, 0x8b], [0x0152, 0x8c], [0x017d, 0x8e], [0x2018, 0x91], [0x2019, 0x92],
    [0x201c, 0x93], [0x201d, 0x94], [0x2022, 0x95], [0x2013, 0x96], [0x2014, 0x97],
    [0x02dc, 0x98], [0x2122, 0x99], [0x0161, 0x9a], [0x203a, 0x9b], [0x0153, 0x9c],
    [0x017e, 0x9e], [0x0178, 0x9f],
  ]);
  let result = "";
  for (const character of value) {
    const code = character.codePointAt(0);
    const byte = cp1252.get(code) ?? (code >= 32 && code <= 255 && !(code >= 0x80 && code <= 0x9f) ? code : 63);
    if (byte === 0x28 || byte === 0x29 || byte === 0x5c) result += `\\${String.fromCharCode(byte)}`;
    else if (byte < 32 || byte > 126) result += `\\${byte.toString(8).padStart(3, "0")}`;
    else result += String.fromCharCode(byte);
  }
  return result;
}

export function createZip(entries) {
  if (entries.length > 0xffff) {
    throw new Error("A ZIP archive cannot contain more than 65,535 files.");
  }
  const localParts = [];
  const directoryParts = [];
  let offset = 0;
  const encoder = new TextEncoder();
  for (const { name, data } of entries) {
    const nameBytes = encoder.encode(name);
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    if (
      nameBytes.length > 0xffff
      || bytes.length > 0xffffffff
      || offset + 30 + nameBytes.length + bytes.length > 0xffffffff
    ) {
      throw new Error("The generated ZIP is too large for this browser app.");
    }
    const checksum = crc32(bytes);
    const local = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0x0800, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, bytes.length, true);
    localView.setUint32(22, bytes.length, true);
    localView.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    localParts.push(local, bytes);

    const directory = new Uint8Array(46 + nameBytes.length);
    const directoryView = new DataView(directory.buffer);
    directoryView.setUint32(0, 0x02014b50, true);
    directoryView.setUint16(4, 20, true);
    directoryView.setUint16(6, 20, true);
    directoryView.setUint16(8, 0x0800, true);
    directoryView.setUint32(16, checksum, true);
    directoryView.setUint32(20, bytes.length, true);
    directoryView.setUint32(24, bytes.length, true);
    directoryView.setUint16(28, nameBytes.length, true);
    directoryView.setUint32(42, offset, true);
    directory.set(nameBytes, 46);
    directoryParts.push(directory);
    offset += local.length + bytes.length;
  }

  const directoryLength = directoryParts.reduce((total, part) => total + part.length, 0);
  if (directoryLength > 0xffffffff) {
    throw new Error("The generated ZIP is too large for this browser app.");
  }
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, directoryLength, true);
  endView.setUint32(16, offset, true);
  return concatenate([...localParts, ...directoryParts, end]);
}

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

function concatenate(parts) {
  const result = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

if (typeof document !== "undefined") {
  initializeApp();
}

function initializeApp() {
  const nameConfig = document.querySelector("#name-config");
  let configToken = null;
  const fileInput = document.querySelector("#file-input");
  const workspace = document.querySelector("#workspace");
  const error = document.querySelector("#error");
  let editBefore = null;
  let lastCell = null;
  let limitTouched = false;
  let dataName = "data.csv";
  let extraFiles = [];
  const dataFileLabel = document.querySelector("#data-file");

  async function loadSample() {
    error.textContent = "";
    try {
      const response = await fetch(new URL("./data/data.csv", import.meta.url));
      if (!response.ok) throw new Error(`Could not load the sample data (${response.status}).`);
      showTable(parseCsv(await response.text()), "sample-data.csv");
    } catch (cause) {
      error.textContent = cause instanceof Error ? cause.message : "Could not load the sample data.";
    }
  }

  async function loadDataFile(file) {
    const extension = file.name.split(".").pop().toLowerCase();
    const data = extension === "csv" ? parseCsv(await file.text()) : await parseXlsx(await file.arrayBuffer());
    showTable(data, file.name);
  }

  async function handleFiles(list) {
    const files = [...list];
    if (!files.length) return;
    error.textContent = "";
    setStatus("");
    const isData = (file) => /\.(csv|xlsx)$/i.test(file.name);
    const legacy = files.filter((file) => /\.xls$/i.test(file.name));
    const dataFile = files.find(isData);
    const others = files.filter((file) => file !== dataFile && !legacy.includes(file));
    if (legacy.length) error.textContent = "Legacy .xls files are not supported. Save the file as .xlsx or .csv.";
    if (dataFile) {
      try {
        await loadDataFile(dataFile);
      } catch (cause) {
        error.textContent = cause instanceof Error ? cause.message : "Could not read this file.";
      }
    }
    others.forEach((file) => {
      extraFiles = extraFiles.filter((existing) => existing.name !== file.name);
      extraFiles.push(file);
    });
    renderExtraFiles();
    if (currentTable) updateFilenameGuidance();
  }

  function renderExtraFiles() {
    const list = document.querySelector("#extra-files");
    list.hidden = extraFiles.length === 0;
    list.replaceChildren();
    extraFiles.forEach((file, index) => {
      const item = document.createElement("li");
      const name = document.createElement("span");
      name.textContent = file.name;
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "icon-button";
      remove.dataset.removeFile = String(index);
      remove.textContent = "✕";
      remove.setAttribute("aria-label", `Remove ${file.name}`);
      item.append(name, remove);
      list.append(item);
    });
  }

  document.querySelector("#extra-files").addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-file]");
    if (!button) return;
    extraFiles.splice(Number(button.dataset.removeFile), 1);
    renderExtraFiles();
    if (currentTable) updateFilenameGuidance();
  });

  document.querySelector("#sample-button").addEventListener("click", loadSample);
  document.querySelector("#add-file-button").addEventListener("click", () => fileInput.click());
  document.querySelector("#replace-file").addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", async () => {
    await handleFiles(fileInput.files);
    fileInput.value = "";
  });
  const extraInput = document.querySelector("#extra-input");
  document.querySelector("#extra-add").addEventListener("click", () => extraInput.click());
  extraInput.addEventListener("change", () => {
    [...extraInput.files].forEach((file) => {
      extraFiles = extraFiles.filter((existing) => existing.name !== file.name);
      extraFiles.push(file);
    });
    extraInput.value = "";
    renderExtraFiles();
    if (currentTable) updateFilenameGuidance();
  });
  const zone = document.querySelector("#preview-panel");
  const hasFiles = (event) => [...(event.dataTransfer?.types ?? [])].includes("Files");
  window.addEventListener("dragover", (event) => {
    if (hasFiles(event)) event.preventDefault();
  });
  window.addEventListener("drop", (event) => {
    if (hasFiles(event)) event.preventDefault();
  });
  zone.addEventListener("dragover", (event) => {
    if (hasFiles(event)) zone.classList.add("is-dragging");
  });
  zone.addEventListener("dragleave", (event) => {
    if (!zone.contains(event.relatedTarget)) zone.classList.remove("is-dragging");
  });
  zone.addEventListener("drop", (event) => {
    zone.classList.remove("is-dragging");
    if (hasFiles(event)) handleFiles(event.dataTransfer.files);
  });

  function readNameOptions() {
    const value = (id) => document.querySelector(id).value;
    return {
      trim: document.querySelector("#nc-trim").checked,
      spaces: value("#nc-spaces"),
      unsafe: value("#nc-unsafe"),
      find: "",
      replaceWith: "",
      rules: [...document.querySelectorAll("#nc-rules li")].map((li) => ({
        find: li.querySelector(".rule-find").value,
        to: li.querySelector(".rule-with").value,
      })).filter((rule) => rule.find),
      caps: value("#nc-caps"),
      limit: document.querySelector("#nc-limit-all").checked ? Math.max(1, Math.floor(Number(value("#nc-limit-default"))) || 15) : 0,
      numbering: value("#name-numbering"),
      separator: sanitizeFilename(value("#name-separator"), "-", "_").slice(0, 5),
      numberPosition: value("#name-position"),
      repeatStart: Math.max(0, Math.floor(Number(value("#nc-repeat-start")))) || 0,
      numberStart: Number(value("#nc-start")) || 0,
      numberPad: Number(value("#nc-pad")) || 1,
    };
  }

  function planOutput() {
    const { labels, excluded, excludedRows } = currentTable.meta;
    const selected = currentTable.headers
      .filter((header) => !excluded.includes(header))
      .map((header) => ({ source: header, label: labels[header] || header, blank: blankRule(currentTable.meta, header) }));
    const template = document.querySelector("#filename-template").value.trim() || "Document-{number}";
    const options = readNameOptions();
    const skipped = new Set(excludedRows);
    const kept = [];
    currentTable.rows.forEach((row, index) => {
      if (skipped.has(index)) return;
      const rawFilename = resolveFilename(template, currentTable.headers, row, kept.length, options);
      const base = Array.from(sanitizeFilename(rawFilename, `Document-${index + 1}`, options.unsafe))
        .slice(0, MAX_BASE_LENGTH).join("");
      kept.push({ row, base, index });
    });
    const names = assignFilenames(kept.map(({ base }) => base), options.numbering, options.separator, { position: options.numberPosition, start: options.repeatStart ?? 1, pad: options.numberPad });
    const seen = new Set();
    let duplicates = 0;
    names.forEach((name) => {
      const key = name.toLocaleLowerCase();
      if (seen.has(key)) duplicates += 1;
      seen.add(key);
    });
    const files = kept.map((entry, position) => ({ ...entry, base: names[position], name: `${names[position]}.${outputKind === "docx" ? "docx" : "pdf"}` }));
    const lengths = files.map(({ name }) => name.length);
    return {
      duplicates,
      selected,
      options,
      template,
      excludedCount: skipped.size,
      shortest: lengths.length ? Math.min(...lengths) : 0,
      longest: lengths.length ? Math.max(...lengths) : 0,
      files,
    };
  }

  function saveBlob(bytes, type, filename) {
    const url = URL.createObjectURL(new Blob([bytes], { type }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  document.querySelector("#download-button").addEventListener("click", async () => {
    if (!currentTable) return;
    try {
      const { formats } = currentTable.meta;
      const { selected, duplicates, files: planned } = planOutput();
      if (!selected.length) throw new Error("Include at least one column.");
      if (duplicates) {
        throw new Error(`${duplicates} file name${duplicates === 1 ? " is" : "s are"} repeated. Turn numbering on in the file name options (⚙) or change the name template.`);
      }
      const zipName = sanitizeFilename(
        document.querySelector("#zip-name").value.trim().replace(/\.zip$/i, ""),
        "Documents",
      );
      const files = planned.map(({ row, index, base }) => {
        const lines = selected.flatMap(({ source, label, blank }) => {
          const value = String(row[source] ?? "");
          if (isBlank(value)) {
            if (blank.mode === "omit") return [];
            if (blank.mode === "fill") return [`${label}: ${blank.value}`];
          }
          return [`${label}: ${formatValue(value, resolveSpec(formats, source, index))}`];
        });
        return { base, lines };
      }).flatMap(({ base, lines }) => [
        ...(outputKind !== "docx" ? [{ name: `${base}.pdf`, data: createPdf(lines) }] : []),
        ...(outputKind !== "pdf" ? [{ name: `${base}.docx`, data: createDocx(lines) }] : []),
      ]);
      if (files.length === 0) throw new Error("No rows are included. Check at least one row.");
      const taken = new Set(files.map(({ name }) => name.toLocaleLowerCase()));
      for (const file of extraFiles) {
        const name = sanitizeFilename(file.name, "file");
        if (taken.has(name.toLocaleLowerCase())) throw new Error(`The added file ${name} has the same name as another file in the ZIP.`);
        taken.add(name.toLocaleLowerCase());
        files.push({ name, data: new Uint8Array(await file.arrayBuffer()) });
      }
      saveBlob(createZip(files), "application/zip", `${zipName}.zip`);
      setStatus(`Created ${files.length - extraFiles.length} document${files.length - extraFiles.length === 1 ? "" : "s"}${extraFiles.length ? ` and added ${extraFiles.length} other file${extraFiles.length === 1 ? "" : "s"}` : ""} in ${zipName}.zip.`);
    } catch (cause) {
      setStatus(cause instanceof Error ? cause.message : "Could not generate the documents.");
    }
  });

  function exportMatrix() {
    const { selected } = planOutput();
    const { formats, excludedRows } = currentTable.meta;
    const skipped = new Set(excludedRows);
    const matrix = [selected.map(({ label }) => label)];
    currentTable.rows.forEach((row, index) => {
      if (skipped.has(index)) return;
      matrix.push(selected.map(({ source, blank }) => {
        const value = String(row[source] ?? "");
        if (isBlank(value)) return blank.mode === "fill" ? blank.value : "";
        return formatValue(value, resolveSpec(formats, source, index));
      }));
    });
    return matrix;
  }

  function exportData(kind) {
    if (!currentTable) return;
    const base = dataName.replace(/\.[^.]+$/, "") || "data";
    const matrix = exportMatrix();
    if (matrix[0].length === 0) {
      setStatus("Include at least one column to export.");
      return;
    }
    if (kind === "csv") saveBlob(toCsv(matrix), "text/csv;charset=utf-8", `${base}-edited.csv`);
    else saveBlob(createXlsx(matrix), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", `${base}-edited.xlsx`);
    setStatus(`Exported ${matrix.length - 1} row${matrix.length === 2 ? "" : "s"} to ${base}-edited.${kind}.`);
  }
  document.querySelector("#export-csv").addEventListener("click", () => exportData("csv"));
  document.querySelector("#export-xlsx").addEventListener("click", () => exportData("xlsx"));

  document.querySelector("#row-limit").addEventListener("change", (event) => {
    rowLimit = event.target.value === "all" ? Infinity : Number(event.target.value);
    renderPreview();
  });
  document.querySelector("#preview-prev").addEventListener("click", () => stepPreview(-1));
  document.querySelector("#preview-next").addEventListener("click", () => stepPreview(1));
  function stepPreview(delta) {
    const count = planOutput().files.length;
    if (!count) return;
    previewPosition = (previewPosition + delta + count) % count;
    updateFilenameGuidance();
  }
  document.querySelectorAll("p.hint[data-as-info]").forEach((hint) => {
    const host = hint.closest("section[data-section]")?.querySelector("label, h4") ?? hint.closest("aside, dialog")?.querySelector("h3");
    if (host) {
      const icon = document.createElement("button");
      icon.type = "button";
      icon.className = "info";
      icon.dataset.info = hint.textContent;
      icon.textContent = "i";
      icon.setAttribute("aria-label", `Info: ${hint.textContent}`);
      host.append(" ", icon);
      hint.remove();
    }
  });
  const infoTip = document.createElement("div");
  infoTip.className = "info-tip";
  infoTip.hidden = true;
  document.body.append(infoTip);
  let infoPinned = null;
  const showInfo = (icon) => {
    infoTip.textContent = icon.dataset.info;
    infoTip.hidden = false;
    const box = icon.getBoundingClientRect();
    infoTip.style.left = `${Math.max(4, Math.min(box.left, window.innerWidth - 270))}px`;
    infoTip.style.top = `${box.bottom + 4}px`;
  };
  document.addEventListener("mouseover", (event) => {
    const icon = event.target.closest(".info");
    if (icon && !infoPinned) showInfo(icon);
  });
  document.addEventListener("mouseout", (event) => {
    if (event.target.closest(".info") && !infoPinned) infoTip.hidden = true;
  });
  document.addEventListener("click", (event) => {
    const icon = event.target.closest(".info");
    if (icon) {
      event.preventDefault();
      if (infoPinned === icon) { infoPinned = null; infoTip.hidden = true; } else { infoPinned = icon; showInfo(icon); }
    } else if (infoPinned) {
      infoPinned = null;
      infoTip.hidden = true;
    }
  });
  const shortNumber = (n) => (Number.isInteger(n) ? String(n) : n.toFixed(2));
  const cut = (value) => (value.length > 28 ? `${value.slice(0, 28)}…` : value);

  function dotPlot(values, min, max, mean, labels) {
    const span = max - min;
    const place = (n) => (span ? ((n - min) / span) * 100 : 50);
    const strip = document.createElement("div");
    strip.className = "strip";
    const step = Math.max(1, Math.ceil(values.length / 400));
    values.filter((_, index) => index % step === 0).forEach((n, index) => {
      const dot = document.createElement("span");
      dot.className = "strip-dot";
      dot.style.left = `${place(n)}%`;
      dot.style.top = `${15 + ((index * 37) % 71)}%`;
      strip.append(dot);
    });
    const avg = document.createElement("span");
    avg.className = "strip-mean";
    avg.style.left = `${place(mean)}%`;
    strip.append(avg);
    const axis = document.createElement("div");
    axis.className = "timeline-axis";
    labels.forEach((label) => {
      const span = document.createElement("span");
      span.textContent = label;
      axis.append(span);
    });
    return [strip, axis];
  }

  function renderStatsInto(box, header) {
    box.replaceChildren();
    const raw = currentTable.rows.map((row) => row[header]);
    const info = columnInsights(raw, { large: currentTable.rows.length > 1000, name: header });
    const distinct = new Set(raw.filter((value) => !isBlank(value)).map((value) => String(value).trim())).size;
    const filled = info.rows - info.empty;
    const title = document.createElement("div");
    title.className = "stats-title";
    title.textContent = currentTable.meta.labels[header] ?? header;
    box.append(title);
    const group = (heading, rows) => {
      if (heading) {
        const h = document.createElement("div");
        h.className = "stats-heading";
        h.textContent = heading;
        box.append(h);
      }
      const dl = document.createElement("dl");
      dl.className = "stats-grid";
      rows.forEach(([name, value]) => {
        const dt = document.createElement("dt");
        dt.textContent = name;
        const dd = document.createElement("dd");
        dd.textContent = String(value);
        dl.append(dt, dd);
      });
      box.append(dl);
    };
    const fillPct = info.rows ? Math.round((filled / info.rows) * 100) : 0;
    const fill = document.createElement("div");
    fill.className = "fill-chart";
    fill.innerHTML = `<div class="fill-bar" role="img" aria-label="${filled} of ${info.rows} cells filled"><span class="fill-on" style="width:${info.rows ? (filled / info.rows) * 100 : 0}%"></span></div><div class="fill-labels"><span>Filled ${fillPct}% (${filled})</span><span>${info.empty ? `Blank ${100 - fillPct}% (${info.empty})` : "No blanks"}</span></div>`;
    box.append(fill);
    group(null, [["Different values", distinct]]);
    if (info.types.length) group("Data types", info.types.map(({ type, count }) => [TYPE_LABELS[type], count]));
    if (info.numeric) {
      const { min, max, mean, stdDev } = info.numeric;
      const nums = raw.filter((value) => !isBlank(value)).map((value) => String(value)).filter((value) => !parseDate(value) && parseNumber(value) !== null).map((value) => parseNumber(value));
      const [strip, axis] = dotPlot(nums, min, max, mean, [`min ${shortNumber(min)}`, `avg ${shortNumber(mean)}`, `max ${shortNumber(max)}`]);
      const h = document.createElement("div");
      h.className = "stats-heading";
      h.textContent = "Numbers";
      box.append(h, strip, axis);
      group(null, [["Spread (std. dev.)", shortNumber(stdDev)]]);
    }
    if (info.dates) {
      const day = (iso) => String(iso).slice(0, 10);
      const h = document.createElement("div");
      h.className = "stats-heading";
      h.textContent = "Dates";
      const { times } = info.dates;
      const mean = times.reduce((sum, t) => sum + t, 0) / times.length;
      const [line, axis] = dotPlot(times, Math.min(...times), Math.max(...times), mean, [day(info.dates.earliest), `avg ${day(info.dates.average)}`, day(info.dates.latest)]);
      box.append(h, line, axis);
    }
    if (info.top.length) {
      const h = document.createElement("div");
      h.className = "stats-heading";
      h.textContent = "Most common values";
      const chart = document.createElement("div");
      chart.className = "bars";
      const addRow = (value, count, percent, muted) => {
        const row = document.createElement("div");
        row.className = `bar-row${muted ? " is-other" : ""}`;
        const name = document.createElement("span");
        name.className = "bar-name";
        name.textContent = cut(value);
        name.title = value;
        const track = document.createElement("span");
        track.className = "bar-track";
        const bar = document.createElement("span");
        bar.style.width = `${Math.max(2, percent)}%`;
        track.append(bar);
        const num = document.createElement("span");
        num.className = "bar-num";
        num.textContent = `${percent}% (${count})`;
        row.append(name, track, num);
        chart.append(row);
      };
      info.top.forEach(({ value, count, percent }) => addRow(value, count, percent));
      if (info.other) addRow("Other", info.other, Math.round((info.other / Math.max(1, filled)) * 100), true);
      box.append(h, chart);
    }
    if (info.incompleteAddresses?.length) {
      group("Incomplete addresses", info.incompleteAddresses.slice(0, 5).map((a) => [`Row ${a.row}`, cut(a.value)]));
    }
  }

  const hiddenPop = document.createElement("div");
  hiddenPop.id = "hidden-pop";
  hiddenPop.className = "stats-pop";
  hiddenPop.hidden = true;
  document.body.append(hiddenPop);
  document.addEventListener("pointerdown", (event) => {
    if (!hiddenPop.hidden && !event.target.closest("#hidden-pop, .hidden-marker")) hiddenPop.hidden = true;
  });
  function showHiddenColumns(names, anchor) {
    hiddenPop.replaceChildren();
    const title = document.createElement("div");
    title.className = "stats-title";
    title.textContent = names.length === 1 ? "Hidden column" : "Hidden columns";
    hiddenPop.append(title);
    const boxes = [];
    names.forEach((name) => {
      const label = document.createElement("label");
      label.className = "check hidden-item";
      const box = document.createElement("input");
      box.type = "checkbox";
      box.value = name;
      boxes.push(box);
      const link = document.createElement("button");
      link.type = "button";
      link.className = "link-button";
      link.textContent = currentTable.meta.labels[name] ?? name;
      link.addEventListener("click", (event) => { event.preventDefault(); hiddenPop.hidden = true; setExcluded([name], false); });
      label.append(box, link);
      hiddenPop.append(label);
    });
    if (names.length > 1) {
      const show = document.createElement("button");
      show.type = "button";
      show.textContent = "Show selected";
      show.addEventListener("click", () => {
        const chosen = boxes.filter((box) => box.checked).map((box) => box.value);
        hiddenPop.hidden = true;
        if (chosen.length) setExcluded(chosen, false);
      });
      const all = document.createElement("button");
      all.type = "button";
      all.textContent = "Show all";
      all.addEventListener("click", () => { hiddenPop.hidden = true; setExcluded(names, false); });
      const row = document.createElement("div");
      row.className = "panel-actions";
      row.append(show, all);
      hiddenPop.append(row);
    }
    const box = anchor.getBoundingClientRect();
    hiddenPop.hidden = false;
    hiddenPop.style.left = `${Math.max(4, Math.min(box.left, window.innerWidth - 240))}px`;
    hiddenPop.style.top = `${box.bottom + 2}px`;
  }
  document.querySelector("#show-changes").addEventListener("change", (event) => {
    document.querySelector("#table-wrap").classList.toggle("show-changes", event.target.checked);
  });
  const logDialog = document.querySelector("#filename-log");
  document.querySelectorAll(".menu-list button").forEach((item) =>
    item.addEventListener("click", () => item.closest("details").removeAttribute("open")));
  document.addEventListener("pointerdown", (event) => {
    const target = event.target;
    document.querySelectorAll("details.menu[open]").forEach((menu) => {
      if (!menu.contains(target)) menu.removeAttribute("open");
    });
    if (!target.closest("#name-config, .name-field, .filename-guidance, .chip-pop")) closeNameConfig();
    if (!target.closest("#filename-suggestions, #filename-template")) document.querySelector("#filename-suggestions").hidden = true;
    const editing = !panels.format.panel.hidden || !panels.stats.panel.hidden;
    if (selectedColumns.size && !selectMode && !(editing && target.closest("#preview-table td")) && !target.closest("th[data-header], .selection-bar, .side-panel, .menu, .table-head, dialog, .panel-rail")) {
      setSelection([]);
    }
  });
  document.querySelector("#filename-log-button").addEventListener("click", () => {
    renderFilenameLog();
    logDialog.showModal();
  });
  document.querySelector("#log-close").addEventListener("click", () => logDialog.close());
  function chooseLogRow(event) {
    const tr = event.target.closest("tr[data-position]");
    if (!tr) return;
    previewPosition = Number(tr.dataset.position);
    logDialog.querySelectorAll("tr[data-position]").forEach((row) => {
      const selected = row === tr;
      row.classList.toggle("is-selected", selected);
      row.setAttribute("aria-selected", String(selected));
    });
    updateFilenameGuidance();
  }
  document.querySelector("#log-table").addEventListener("click", chooseLogRow);
  document.querySelector("#log-table").addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      chooseLogRow(event);
    }
  });
  logDialog.addEventListener("click", (event) => {
    if (event.target === logDialog) logDialog.close();
  });

  document.querySelectorAll("#name-config input, #name-config select").forEach((control) => {
    if (control.id === "nc-limit-all" || control.id === "nc-limit-default") control.addEventListener("input", () => { limitTouched = true; });
    control.addEventListener("input", updateFilenameGuidance);
    control.addEventListener("change", updateFilenameGuidance);
  });
  document.querySelector("#ext-toggle").addEventListener("click", (event) => {
    outputKind = { pdf: "docx", docx: "both", both: "pdf" }[outputKind];
    const label = { pdf: ".pdf", docx: ".docx", both: ".pdf + .docx" }[outputKind];
    event.currentTarget.textContent = label;
    event.currentTarget.setAttribute("aria-label", `File type: ${label}. Click to change.`);
    updateFilenameGuidance();
  });
  const preview = document.querySelector("#preview-table");
  preview.addEventListener("click", (event) => {
    const check = event.target.closest("[data-check-column]");
    if (check) {
      selectColumn(check.dataset.checkColumn, { ctrlKey: !event.shiftKey, shiftKey: event.shiftKey });
      return;
    }
    const nameClick = event.target.closest(".col-name");
    if (nameClick && !event.target.closest("input")) {
      const header = nameClick.closest("th").dataset.header;
      if (!(selectedColumns.size > 1 && selectedColumns.has(header))) setSelection([header], header);
      openColumnPanel([...selectedColumns]);
      return;
    }
    const editButton = event.target.closest("[data-edit-column]");
    if (editButton) {
      const header = editButton.dataset.editColumn;
      if (!(selectedColumns.size > 1 && selectedColumns.has(header))) setSelection([header], header);
      openColumnPanel([...selectedColumns]);
      return;
    }
    const exclude = event.target.closest("[data-exclude-column]");
    if (exclude) {
      setExcluded([exclude.dataset.excludeColumn], true);
      return;
    }
    const marker = event.target.closest(".hidden-marker");
    if (marker) {
      showHiddenColumns(marker.dataset.hidden.split("\u0001"), marker);
      return;
    }
    if (event.target.closest("thead th.corner")) {
      setSelectMode(!selectModeOn());
      return;
    }
    if (event.target.closest(".add-column")) {
      togglePanel("columns");
      return;
    }
    const th = event.target.closest("th[data-header]");
    if (!th || event.target.closest("input")) return;
    selectColumn(th.dataset.header, event);
  });
  preview.addEventListener("dblclick", (event) => {
    const th = event.target.closest("th[data-header]");
    if (th && !event.target.closest("input, button")) startHeaderRename(th);
  });
  preview.addEventListener("keydown", (event) => {
    const th = event.target.closest("th[data-header]");
    if (!th || event.target !== th) return;
    const headers = currentTable.headers;
    const header = th.dataset.header;
    if (event.key === "Enter" || event.key === "F2") {
      event.preventDefault();
      startHeaderRename(th);
    } else if (event.key === " ") {
      event.preventDefault();
      selectColumn(header, { ctrlKey: true });
    } else if (event.key === "Escape") {
      setSelection([]);
    } else if (event.key === "ArrowLeft" || event.key === "ArrowRight") {
      event.preventDefault();
      const next = headers[headers.indexOf(header) + (event.key === "ArrowRight" ? 1 : -1)];
      if (!next) return;
      if (event.shiftKey) selectColumn(next, { shiftKey: true }, header);
      preview.querySelector(`th[data-header="${CSS.escape(next)}"]`)?.focus();
    }
  });

  function selectColumn(header, event, anchorOverride) {
    const headers = currentTable.headers;
    const anchor = anchorOverride ?? selectionAnchor;
    if (event.shiftKey && anchor && headers.includes(anchor)) {
      const from = headers.indexOf(anchor);
      const to = headers.indexOf(header);
      const range = headers.slice(Math.min(from, to), Math.max(from, to) + 1);
      setSelection(event.ctrlKey || event.metaKey ? [...selectedColumns, ...range] : range, anchor);
    } else if (event.ctrlKey || event.metaKey) {
      const next = new Set(selectedColumns);
      if (next.has(header)) next.delete(header);
      else next.add(header);
      setSelection([...next], header);
    } else if (selectedColumns.size === 1 && selectedColumns.has(header)) {
      setSelection([], header);
    } else {
      setSelection([header], header);
    }
  }

  function setSelectMode(on) {
    selectMode = on;
    if (!on) {
      selectedRows = new Set();
      syncRowSelection();
      setSelection([]);
    }
    syncSelectMode();
  }

  function setSelection(headers, anchor = selectionAnchor) {
    selectedColumns = new Set(headers);
    selectionAnchor = anchor;
    syncSelection();
    onSelectionChanged();
  }

  function startHeaderRename(th) {
    const header = th.dataset.header;
    const name = th.querySelector(".col-name");
    if (!name) return;
    const input = document.createElement("input");
    input.type = "text";
    input.className = "col-label";
    input.value = name.textContent;
    input.size = Math.max(8, Math.min(28, input.value.length + 2));
    input.setAttribute("aria-label", `Name of column ${header} in the output`);
    name.replaceWith(input);
    th.draggable = false;
    input.focus();
    input.select();
    let finished = false;
    const finish = (save) => {
      if (finished) return;
      finished = true;
      if (save) renameColumn(header, input.value);
      renderPreview();
      preview.querySelector(`th[data-header="${CSS.escape(header)}"]`)?.focus();
    };
    input.addEventListener("keydown", (event) => {
      event.stopPropagation();
      if (event.key === "Enter") {
        event.preventDefault();
        finish(true);
      } else if (event.key === "Escape") {
        event.preventDefault();
        finish(false);
      }
    });
    input.addEventListener("blur", () => finish(true));
  }

  preview.addEventListener("focusin", (event) => {
    const cell = event.target.closest("[data-row-index]");
    if (!cell) return;
    editBefore = cloneTable(currentTable);
    lastCell = { rowIndex: Number(cell.dataset.rowIndex), header: cell.dataset.header };
    focusedRow = lastCell.rowIndex;
    updateFooter();
    if (!panels.format.panel.hidden) renderCellException();
    const raw = currentTable.rows[lastCell.rowIndex][lastCell.header] ?? "";
    if (cell.textContent !== raw) cell.textContent = raw;
  });
  preview.addEventListener("focusout", (event) => {
    const cell = event.target.closest("[data-row-index]");
    if (!cell || !editBefore || !currentTable) return;
    const rowIndex = Number(cell.dataset.rowIndex);
    const header = cell.dataset.header;
    const value = cell.innerText.replace(/\u00a0/g, " ").replace(/^\n$/, "");
    if (currentTable.rows[rowIndex][header] !== value) {
      currentTable.rows[rowIndex][header] = value;
      commitHistory(editBefore, `Edit row ${rowIndex + 1}, ${header}`);
      updateFilenameGuidance();
    }
    editBefore = null;
    paintCell(cell);
  });
  let dragHeader = null;
  preview.addEventListener("dragstart", (event) => {
    const th = event.target.closest("th[data-header]");
    if (!th || event.target.closest("input")) return;
    dragHeader = th.dataset.header;
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", dragHeader);
  });
  preview.addEventListener("dragover", (event) => {
    const th = event.target.closest("th[data-header]");
    if (!dragHeader || !th) return;
    event.preventDefault();
    const after = event.clientX > th.getBoundingClientRect().left + th.offsetWidth / 2;
    preview.querySelectorAll(".drop-before, .drop-after").forEach((node) => node.classList.remove("drop-before", "drop-after"));
    th.classList.add(after ? "drop-after" : "drop-before");
  });
  preview.addEventListener("drop", (event) => {
    const th = event.target.closest("th[data-header]");
    if (event.target.closest("[data-row-index]")) event.preventDefault();
    if (!dragHeader || !th) return;
    event.preventDefault();
    const after = event.clientX > th.getBoundingClientRect().left + th.offsetWidth / 2;
    const moving = dragHeader;
    dragHeader = null;
    reorderColumn(moving, th.dataset.header, after);
  });
  preview.addEventListener("dragend", () => {
    dragHeader = null;
    preview.querySelectorAll(".drop-before, .drop-after").forEach((node) => node.classList.remove("drop-before", "drop-after"));
  });
  document.addEventListener("keydown", (event) => {
    const handle = event.target.closest?.(".drag-handle");
    if (!handle) return;
    const delta = { ArrowUp: -1, ArrowLeft: -1, ArrowDown: 1, ArrowRight: 1 }[event.key];
    if (!delta) return;
    event.preventDefault();
    const container = handle.closest("#columns-list") ? "#columns-list" : "#preview-table";
    if (moveColumn(handle.dataset.header, delta)) {
      document.querySelector(`${container} .drag-handle[data-header="${CSS.escape(handle.dataset.header)}"]`)?.focus();
    }
  });
  document.querySelector("#preview-table").addEventListener("paste", (event) => {
    const cell = event.target.closest("[data-row-index]");
    if (!cell) return;
    event.preventDefault();
    const text = event.clipboardData?.getData("text/plain") ?? "";
    const selection = window.getSelection();
    if (!selection?.rangeCount) {
      cell.append(document.createTextNode(text));
      return;
    }
    const range = selection.getRangeAt(0);
    range.deleteContents();
    const textNode = document.createTextNode(text);
    range.insertNode(textNode);
    range.setStartAfter(textNode);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  });
  document.querySelector("#undo-button").addEventListener("click", () => applyHistory(undoStates, redoStates));
  document.querySelector("#redo-button").addEventListener("click", () => applyHistory(redoStates, undoStates));
  document.querySelector("#revert-button").addEventListener("click", () => {
    if (!originalTable) return;
    currentTable = cloneTable(originalTable);
    undoStates = [];
    redoStates = [];
    closePanels();
    renderTable();
    syncHistoryButtons();
    updateFilenameGuidance();
    setStatus("Edits reverted to the original spreadsheet data.");
  });
  document.addEventListener("keydown", (event) => {
    if (!(event.ctrlKey || event.metaKey) || !currentTable || workspace.hidden) return;
    const active = document.activeElement;
    const editingCell = active.matches("[data-row-index]");
    if (!editingCell && active.matches("input, textarea, select, [contenteditable='true']")) return;
    if (event.key.toLowerCase() === "z" && !event.shiftKey) {
      event.preventDefault();
      if (editingCell) active.blur();
      applyHistory(undoStates, redoStates);
    } else if (event.key.toLowerCase() === "y" || (event.key.toLowerCase() === "z" && event.shiftKey)) {
      event.preventDefault();
      if (editingCell) active.blur();
      applyHistory(redoStates, undoStates);
    }
  });

  function renderDataFile(data, sourceName) {
    dataName = sourceName;
    loadStats = data.stats ?? null;
    dataFileLabel.textContent = sourceName;
  }

  function setStatus(message) {
    document.querySelector("#status").textContent = message;
  }

  function showTable(data, sourceName) {
    currentTable = cloneTable(data);
    originalTable = cloneTable(data);
    undoStates = [];
    redoStates = [];
    lastCell = null;
    previewPosition = 0;
    selectedColumns = new Set();
    selectionAnchor = null;
    selectedRows = new Set();
    rowAnchor = null;
    focusedRow = null;
    replaceLog = [];
    closePanels();
    document.querySelector("#preview-panel").classList.add("has-data");
    document.querySelector("#drop-empty").hidden = true;
    document.querySelector("#table-wrap").hidden = false;
    selectMode = false;
    renderTable();
    error.textContent = "";
    setStatus("");
    document.querySelector("#filename-template").value = defaultTemplate(data.headers);
    renderDataFile(data, sourceName);
    updateAssist();
    updateFilenameGuidance();
    syncHistoryButtons();
  }

  function refreshAfterColumnChange() {
    renderColumnsPanel();
    renderPreview();
    updateFilenameGuidance();
  }

  function moveColumn(header, delta) {
    const index = currentTable.headers.indexOf(header);
    const next = index + delta;
    if (index < 0 || next < 0 || next >= currentTable.headers.length) return false;
    const before = cloneTable(currentTable);
    currentTable.headers.splice(index, 1);
    currentTable.headers.splice(next, 0, header);
    commitHistory(before, `Move column ${header}`);
    refreshAfterColumnChange();
    return true;
  }

  function reorderColumn(moving, targetHeader, after) {
    if (!moving || moving === targetHeader) return;
    const before = cloneTable(currentTable);
    currentTable.headers.splice(currentTable.headers.indexOf(moving), 1);
    const target = currentTable.headers.indexOf(targetHeader);
    currentTable.headers.splice(target + (after ? 1 : 0), 0, moving);
    if (JSON.stringify(currentTable.headers) === JSON.stringify(before.headers)) return;
    commitHistory(before, `Move column ${moving}`);
    refreshAfterColumnChange();
  }

  function renameColumn(header, value, rebuildPanel = true) {
    const name = value.trim();
    const { labels } = currentTable.meta;
    if (name === (labels[header] ?? header)) return;
    const before = cloneTable(currentTable);
    if (!name || name === header) delete labels[header];
    else labels[header] = name;
    commitHistory(before, `Rename column ${header}`);
    if (rebuildPanel) renderColumnsPanel();
    renderPreview();
  }

  function moveColumnTo(header, position) {
    const target = Math.max(0, Math.min(currentTable.headers.length - 1, Math.round(position) - 1));
    const index = currentTable.headers.indexOf(header);
    if (!Number.isFinite(target) || target === index) return false;
    return moveColumn(header, target - index);
  }

  function setExcluded(headers, exclude) {
    const before = cloneTable(currentTable);
    const excluded = new Set(currentTable.meta.excluded);
    headers.forEach((header) => (exclude ? excluded.add(header) : excluded.delete(header)));
    currentTable.meta.excluded = currentTable.headers.filter((header) => excluded.has(header));
    if (JSON.stringify(currentTable.meta.excluded) === JSON.stringify(before.meta.excluded)) return;
    commitHistory(
      before,
      headers.length === 1
        ? `${exclude ? "Hide" : "Show"} column ${headers[0]}`
        : `${exclude ? "Hide" : "Show"} all columns`,
    );
    refreshAfterColumnChange();
  }

  function setRowsExcluded(indexes, exclude, label) {
    const before = cloneTable(currentTable);
    const out = new Set(currentTable.meta.excludedRows);
    indexes.forEach((index) => (exclude ? out.add(index) : out.delete(index)));
    currentTable.meta.excludedRows = [...out].sort((a, b) => a - b);
    if (JSON.stringify(currentTable.meta.excludedRows) === JSON.stringify(before.meta.excludedRows)) return;
    commitHistory(before, label);
    renderPreview();
    updateFilenameGuidance();
  }

  document.querySelector("#preview-table").addEventListener("click", (event) => {
    const th = event.target.closest('th[scope="row"]');
    if (!th) return;
    const index = Number(th.dataset.rowIndex);
    const x = event.target.closest(".row-x");
    if (x) {
      const out = currentTable.meta.excludedRows.includes(index);
      setRowsExcluded([index], !out, `${out ? "Include" : "Exclude"} row ${index + 1}`);
      return;
    }
    if (event.shiftKey && rowAnchor !== null) {
      const [from, to] = [rowAnchor, index].sort((a, b) => a - b);
      const range = Array.from({ length: to - from + 1 }, (_, i) => from + i);
      selectedRows = new Set([...selectedRows, ...range]);
    } else if (selectedRows.has(index)) {
      selectedRows.delete(index);
    } else {
      selectedRows.add(index);
    }
    rowAnchor = index;
    syncRowSelection();
  });
  document.querySelector("#row-selection-clear").addEventListener("click", () => {
    selectedRows = new Set();
    syncRowSelection();
    if (!selectedColumns.size) setSelectMode(false);
  });
  document.querySelector("#row-selection-exclude").addEventListener("click", () => {
    const rows = [...selectedRows];
    const allOut = rows.every((index) => currentTable.meta.excludedRows.includes(index));
    setRowsExcluded(rows, !allOut, `${allOut ? "Include" : "Exclude"} ${rows.length} rows`);
  });
  document.querySelector("#rows-all").addEventListener("click", () => {
    setRowsExcluded(currentTable.rows.map((_, index) => index), false, "Include all rows");
  });
  document.querySelector("#rows-blank").addEventListener("click", () => {
    const included = currentTable.headers.filter((header) => !currentTable.meta.excluded.includes(header));
    const blanks = currentTable.rows.map((row, index) => (included.some((header) => isBlank(row[header])) ? index : -1)).filter((index) => index >= 0);
    setRowsExcluded(blanks, true, "Exclude rows with blank cells");
  });

  const panels = {
    columns: { button: document.querySelector("#columns-button"), panel: document.querySelector("#columns-panel") },
    format: { button: document.querySelector("#selection-format"), panel: document.querySelector("#format-panel") },
    changes: { button: document.querySelector("#changes-button"), panel: document.querySelector("#changes-panel") },
    stats: { button: null, panel: document.querySelector("#stats-panel") },
  };

  function openPanel(name) {
    if (name === "columns") selectMode = true;
    else if (!panels.columns.panel.hidden) selectMode = false;
    document.querySelector("#panel-rail").hidden = true;
    if (name !== "format") document.querySelector("#replace-dialog")?.close();
    const hadHighlight = highlightCells.size > 0 || highlightHeaders.size > 0 || highlightBlanks.size > 0;
    document.querySelectorAll(".panel-tabs button").forEach((tab) => {
      tab.classList.toggle("is-active", tab.dataset.tab === name);
      tab.setAttribute("aria-selected", String(tab.dataset.tab === name));
    });
    Object.entries(panels).forEach(([key, { button, panel }]) => {
      const open = key === name;
      panel.hidden = !open;
      button?.setAttribute("aria-expanded", String(open));
      button?.classList.toggle("is-active", open);
    });
    if (name !== "format" && name !== "stats") {
      highlightCells = new Set();
      highlightHeaders = new Set();
    }
    if (name !== "format") highlightBlanks = new Set();
    if (hadHighlight || highlightBlanks.size) syncHighlights();
    const editing = name === "format" || name === "stats";
    const editButton = document.querySelector("#edit-all");
    editButton.classList.toggle("is-active", editing);
    editButton.setAttribute("aria-expanded", String(editing));
    syncSelectMode();
  }

  const rail = document.querySelector("#panel-rail");
  let collapsed = null;
  const RAIL_LABELS = { columns: "Columns", format: "Edit", stats: "Stats", changes: "History" };
  document.querySelector("#workspace").addEventListener("click", (event) => {
    if (!event.target.closest("[data-collapse-panel]")) return;
    collapsed = Object.keys(panels).find((key) => !panels[key].panel.hidden);
    if (!collapsed) return;
    panels[collapsed].panel.hidden = true;
    rail.textContent = `‹ ${RAIL_LABELS[collapsed]}`;
    rail.setAttribute("aria-label", `Show ${RAIL_LABELS[collapsed]} panel`);
    rail.hidden = false;
    rail.focus();
  });
  rail.addEventListener("click", () => {
    rail.hidden = true;
    if (collapsed) panels[collapsed].panel.hidden = false;
    collapsed = null;
  });

  function closePanels() {
    openPanel(null);
  }

  function togglePanel(name) {
    const { panel } = panels[name];
    if (!panel.hidden) {
      closePanels();
    } else {
      openPanel(name);
      if (name === "columns") {
        panel.querySelector("input[type=checkbox]")?.focus();
      }
    }
  }

  document.querySelector("#workspace").addEventListener("click", (event) => {
    if (event.target.closest("[data-close-panel]")) {
      const open = Object.values(panels).find(({ panel }) => !panel.hidden);
      closePanels();
      open?.button?.focus();
    }
  });

  Object.entries(panels).forEach(([name, { panel }]) => {
    panels[name].button?.addEventListener("click", () => togglePanel(name));
    panel.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        closePanels();
        panels[name].button?.focus();
      }
    });
  });
  const columnsPanel = document.querySelector("#columns-panel");
  columnsPanel.addEventListener("change", (event) => {
    const item = event.target.closest("li");
    if (!item) return;
    const header = item.dataset.source;
    if (event.target.matches(".col-check")) {
      setExcluded([header], !event.target.checked);
      document.querySelector(`#columns-list li[data-source="${CSS.escape(header)}"] .col-check`)?.focus();
    } else if (event.target.matches(".col-rename")) {
      renameColumn(header, event.target.value, false);
      paintColumnItem(item);
    }
  });
  columnsPanel.addEventListener("click", (event) => {
    const revert = event.target.closest(".rename-revert");
    if (!revert) return;
    const item = revert.closest("li");
    renameColumn(item.dataset.source, "", false);
    paintColumnItem(item);
  });
  columnsPanel.addEventListener("keydown", (event) => {
    if (event.target.matches(".col-rename") && event.key === "Enter") event.target.blur();
  });
  document.querySelector("#columns-all").addEventListener("click", () => setExcluded(currentTable.headers, false));
  document.querySelector("#columns-none").addEventListener("click", () => setExcluded(currentTable.headers, true));

  const columnList = document.querySelector("#columns-list");
  let dragListColumn = null;
  const clearListDrop = () => columnList.querySelectorAll(".drop-before, .drop-after").forEach((node) => node.classList.remove("drop-before", "drop-after"));
  columnList.addEventListener("dragstart", (event) => {
    const handle = event.target.closest(".drag-handle");
    if (!handle) return;
    dragListColumn = handle.dataset.header;
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", dragListColumn);
  });
  columnList.addEventListener("dragover", (event) => {
    const item = event.target.closest("li");
    if (!dragListColumn || !item) return;
    event.preventDefault();
    const box = item.getBoundingClientRect();
    clearListDrop();
    item.classList.add(event.clientY > box.top + box.height / 2 ? "drop-after" : "drop-before");
  });
  columnList.addEventListener("drop", (event) => {
    const item = event.target.closest("li");
    if (!dragListColumn || !item) return;
    event.preventDefault();
    const box = item.getBoundingClientRect();
    const moving = dragListColumn;
    dragListColumn = null;
    clearListDrop();
    reorderColumn(moving, item.dataset.source, event.clientY > box.top + box.height / 2);
  });
  columnList.addEventListener("dragend", () => {
    dragListColumn = null;
    clearListDrop();
  });

  document.querySelector("#changes-list").addEventListener("click", (event) => {
    const button = event.target.closest("[data-undo-count]");
    if (!button) return;
    const count = Number(button.dataset.undoCount);
    for (let step = 0; step < count; step += 1) applyHistory(undoStates, redoStates);
    closePanels();
  });

  const $ = (selector) => document.querySelector(selector);
  let editAll = false;

  function onSelectionChanged() {
    if (!currentTable) return;
    if (!panels.format.panel.hidden) {
      if (selectedColumns.size) openColumnPanel([...selectedColumns]);
      else if (!editAll) openColumnPanel(includedHeaders(), { all: true });
    }
    else if (!panels.stats.panel.hidden && selectedColumns.size) openStatsPanel([...selectedColumns]);
  }

  $("#selection-clear").addEventListener("click", () => {
    setSelection([]);
    if (!selectedRows.size) setSelectMode(false);
  });
  $("#selection-all").addEventListener("click", () => setSelection(currentTable.headers, currentTable.headers[0]));
  $("#selection-hide").addEventListener("click", () => {
    const headers = [...selectedColumns];
    if (!headers.length) return;
    const allHidden = headers.every((header) => currentTable.meta.excluded.includes(header));
    setExcluded(headers, !allHidden);
  });

  const KINDS = {
    number: { label: "Number", presets: NUMBER_PRESETS, valid: validNumberPattern, custom: "e.g. $#,##0.00", hint: "0 and # are digits: 0.0, $#,##0.00, 0%" },
    date: { label: "Date", presets: DATE_PRESETS, valid: validDatePattern, custom: "e.g. dd-mmm-yy", hint: "yyyy, yy, mmmm, mmm, mm, m, dd, d" },
    bool: { label: "Boolean", presets: BOOL_PRESETS, valid: validBoolPattern, noCustom: true },
    phone: { label: "Phone", presets: PHONE_PRESETS, valid: validPhonePattern, custom: "e.g. ###-###-####", hint: "# is a digit (10 digits). 7-digit numbers become ###-####; extensions are kept as x123" },
  };
  const kindTests = {
    date: (value) => Boolean(parseDate(value)),
    phone: (value) => !parseDate(value) && Boolean(parsePhoneParts(value)),
    bool: (value) => parseBool(value) !== null,
    number: (value, phoneColumn) => !parseDate(value) && !(phoneColumn && parsePhoneParts(value)) && parseNumber(value) !== null,
  };
  let formatTargets = [];
  let replaceLog = [];
  let draftSpec = {};
  let draftInvalid = false;
  let draftCaps = "keep";
  const replaceDialog = $("#replace-dialog");
  const targetLabel = () => (formatTargets.length === 1 ? `column ${formatTargets[0]}` : `${formatTargets.length} columns`);

  function groupCounts() {
    const counts = { total: 0, date: 0, number: 0, phone: 0, bool: 0, empty: 0, text: 0 };
    formatTargets.forEach((header) => {
      const phoneColumn = isPhoneName(header);
      currentTable.rows.forEach((row) => {
        const value = String(row[header] ?? "");
        if (isBlank(value)) { counts.empty += 1; return; }
        counts.total += 1;
        let matched = false;
        for (const kind of ["date", "number", "phone", "bool"]) {
          if (kind === "phone" && !phoneColumn) continue;
          if (kindTests[kind](value, phoneColumn)) { counts[kind] += 1; matched = true; }
        }
        if (!matched) counts.text += 1;
      });
    });
    return counts;
  }

  function renderEditTop() {
    const title = $("#format-title");
    const box = $("#ed-name");
    title.replaceChildren();
    box.replaceChildren();
    if (formatTargets.length === 1) {
      const header = formatTargets[0];
      const label = currentTable.meta.labels[header] ?? header;
      const edit = document.createElement("button");
      edit.type = "button";
      edit.className = "title-edit";
      const text = document.createElement("span");
      text.textContent = label;
      const pencil = document.createElement("span");
      pencil.className = "pencil";
      pencil.setAttribute("aria-hidden", "true");
      pencil.textContent = "✎";
      edit.append(text, pencil);
      edit.addEventListener("click", () => {
        const input = document.createElement("input");
        input.type = "text";
        input.className = "title-input";
        input.value = label;
        input.setAttribute("aria-label", "Column name");
        edit.replaceWith(input);
        input.focus();
        input.select();
        let done = false;
        const finish = (save) => {
          if (done) return;
          done = true;
          if (save) renameColumn(header, input.value);
          renderEditTop();
        };
        input.addEventListener("blur", () => finish(true));
        input.addEventListener("keydown", (event) => {
          if (event.key === "Enter") input.blur();
          else if (event.key === "Escape") { event.stopPropagation(); finish(false); }
        });
      });
      title.append(edit);
      if (label !== header) {
        const revert = document.createElement("button");
        revert.type = "button";
        revert.className = "icon-button revert";
        revert.textContent = "↺";
        revert.title = `Revert to “${header}”`;
        revert.setAttribute("aria-label", `Revert name to ${header}`);
        revert.addEventListener("click", () => { renameColumn(header, header); renderEditTop(); });
        title.append(revert);
      }
    } else {
      title.textContent = editAll ? "All columns" : `${formatTargets.length} columns`;
    }
    $("#rep-scope").textContent = editAll ? "In all columns." : `In ${formatTargets.length === 1 ? `column “${currentTable.meta.labels[formatTargets[0]] ?? formatTargets[0]}”` : `${formatTargets.length} selected columns`}.`;
    if (lastCell && formatTargets.length === 1 && formatTargets[0] === lastCell.header && currentTable.rows[lastCell.rowIndex]) {
      const { rowIndex, header } = lastCell;
      const wrap = document.createElement("div");
      wrap.className = "line";
      const label = document.createElement("label");
      label.textContent = `Row ${rowIndex + 1}`;
      const input = document.createElement("input");
      input.type = "text";
      input.id = "ed-cell";
      label.htmlFor = input.id;
      input.value = currentTable.rows[rowIndex][header] ?? "";
      input.addEventListener("change", () => {
        if ((currentTable.rows[rowIndex][header] ?? "") === input.value) return;
        const before = cloneTable(currentTable);
        currentTable.rows[rowIndex][header] = input.value;
        commitHistory(before, `Edit row ${rowIndex + 1}, ${header}`);
        renderPreview();
        updateFilenameGuidance();
        renderTypes();
        renderValues();
        refreshEdit();
      });
      input.addEventListener("keydown", (event) => { if (event.key === "Enter") input.blur(); });
      wrap.append(label, input);
      box.append(wrap);
    }
    const out = formatTargets.filter((header) => currentTable.meta.excluded.includes(header)).length;
    const include = $("#ed-include");
    include.checked = out === 0;
    include.indeterminate = out > 0 && out < formatTargets.length;
    $("#ed-body").classList.toggle("is-off", out === formatTargets.length);
  }

  $("#ed-include").addEventListener("change", (event) => {
    setExcluded(formatTargets, !event.target.checked);
    renderEditTop();
  });
  // Touching a control on a dimmed panel puts the column back in the output.
  const wake = (event) => {
    if (!$("#ed-body").classList.contains("is-off") || !event.target.closest("input, select, button")) return;
    setExcluded(formatTargets, false);
    renderEditTop();
  };
  $("#ed-body").addEventListener("pointerdown", wake, true);
  $("#ed-body").addEventListener("focusin", wake);

  function readSpec() {
    const spec = {};
    Object.keys(KINDS).forEach((kind) => { if (draftSpec[kind]) spec[kind] = draftSpec[kind]; });
    if (draftCaps !== "keep") spec.caps = draftCaps;
    spec.trim = $("#ed-trim").checked;
    return { spec, invalid: draftInvalid };
  }

  const CAPS = [["keep", "Use original data"], ["upper", "UPPERCASE"], ["lower", "lowercase"], ["title", "Title Case"]];

  function formatControl(kind) {
    const def = KINDS[kind];
    const cell = document.createElement("span");
    cell.className = "fmt-cell";
    const select = document.createElement("select");
    select.className = "fmt-select";
    select.setAttribute("aria-label", `${def.label} format`);
    select.append(new Option("none", ""));
    const groups = new Map();
    def.presets.forEach((preset) => {
      const option = new Option(`${preset.pattern}   (${preset.label})`, preset.pattern);
      if (!preset.group) { select.append(option); return; }
      if (!groups.has(preset.group)) {
        const group = document.createElement("optgroup");
        group.label = preset.group;
        groups.set(preset.group, group);
        select.append(group);
      }
      groups.get(preset.group).append(option);
    });
    if (!def.noCustom) select.append(new Option("Custom…", "__custom"));
    const pattern = draftSpec[kind] ?? "";
    const known = def.presets.some((preset) => preset.pattern === pattern);
    select.value = known ? pattern : pattern ? "__custom" : "";
    const custom = document.createElement("input");
    custom.type = "text";
    custom.className = "fmt-custom";
    custom.placeholder = def.custom ?? "";
    custom.title = def.hint ?? "";
    custom.setAttribute("aria-label", `Custom ${def.label.toLowerCase()} format. ${def.hint ?? ""}`);
    custom.hidden = select.value !== "__custom";
    if (select.value === "__custom") custom.value = pattern;
    const apply = () => {
      custom.hidden = select.value !== "__custom";
      const chosen = select.value === "__custom" ? custom.value.trim() : select.value;
      draftInvalid = Boolean(chosen) && !def.valid(chosen);
      custom.classList.toggle("is-invalid", draftInvalid);
      if (!chosen || draftInvalid) delete draftSpec[kind];
      else draftSpec[kind] = chosen;
      select.classList.toggle("is-set", Boolean(draftSpec[kind]));
      refreshEdit();
    };
    select.classList.toggle("is-set", Boolean(pattern));
    select.addEventListener("change", () => { apply(); if (!custom.hidden) custom.focus(); });
    custom.addEventListener("input", apply);
    cell.append(select, custom);
    return cell;
  }

  function capsControl() {
    const select = document.createElement("select");
    select.className = "fmt-select";
    select.setAttribute("aria-label", "Capitalization");
    CAPS.forEach(([value, label]) => select.append(new Option(value === "keep" ? "none" : label, value)));
    select.value = draftCaps;
    select.classList.toggle("is-set", draftCaps !== "keep");
    select.addEventListener("change", () => {
      draftCaps = select.value;
      select.classList.toggle("is-set", draftCaps !== "keep");
      refreshEdit();
    });
    const cell = document.createElement("span");
    cell.className = "fmt-cell";
    cell.append(select);
    return cell;
  }

  function renderTypes() {
    const box = $("#ed-types");
    box.replaceChildren();
    const counts = groupCounts();
    if (!counts.total) {
      const none = document.createElement("div");
      none.className = "muted small";
      none.textContent = "No values.";
      box.append(none);
    }
    const head = document.createElement("div");
    head.className = "type-row type-head";
    head.innerHTML = "<span>Type</span><span>Format</span><span>Values</span>";
    if (counts.total) box.append(head);
    [["text", "Text"], ["number", "Number"], ["date", "Date"], ["bool", "Boolean"], ["phone", "Phone"]].forEach(([kind, name]) => {
      if (!counts[kind]) return;
      const row = document.createElement("div");
      row.className = "type-row";
      const label = document.createElement("span");
      label.className = "type-name";
      label.textContent = name;
      const num = document.createElement("span");
      num.className = "stat-num";
      num.textContent = `${Math.round((counts[kind] / counts.total) * 100)}% (${counts[kind]}/${counts.total})`;
      row.append(label, kind === "text" ? capsControl() : formatControl(kind), num);
      box.append(row);
    });
    const extra = document.createElement("div");
    extra.id = "ed-extra";
    box.append(extra);
  }

  function renderValues() {
    const body = $("#ed-values-body");
    body.replaceChildren();
    const filled = [];
    formatTargets.forEach((header) => currentTable.rows.forEach((row) => {
      const value = String(row[header] ?? "").trim();
      if (value !== "") filled.push(value);
    }));
    const counts = new Map();
    filled.forEach((value) => counts.set(value, (counts.get(value) ?? 0) + 1));
    const ranked = [...counts].sort((x, y) => y[1] - x[1] || x[0].localeCompare(y[0]));
    const repeated = ranked.filter(([, n]) => n >= 2);
    const fixes = similarValues(filled);
    const showList = repeated.length >= 2;
    const pct = (n) => `${Math.round((n / filled.length) * 100)}%`;
    if (showList) {
      const limit = currentTable.rows.length > 1000 ? 10 : 8;
      const shown = repeated.slice(0, limit);
      const rest = ranked.filter(([value]) => !shown.some(([top]) => top === value));
      const selected = new Set();
      const boxes = new Map();
      const head = document.createElement("div");
      head.className = "stat-line value-head";
      head.innerHTML = "<span>Value</span><span>Count (share)</span>";
      const list = document.createElement("ul");
      list.className = "value-list";
      const bar = document.createElement("div");
      bar.className = "value-bar";
      bar.hidden = true;
      const barText = document.createElement("span");
      const barInput = document.createElement("input");
      barInput.type = "text";
      barInput.placeholder = "New value";
      barInput.setAttribute("aria-label", "New value for the selected values");
      barInput.autocomplete = "off";
      const barRun = document.createElement("button");
      barRun.type = "button";
      barRun.textContent = "Replace";
      barRun.disabled = true;
      const barClear = document.createElement("button");
      barClear.type = "button";
      barClear.className = "link-button";
      barClear.textContent = "Clear";
      bar.append(barText, barInput, barRun, barClear);
      const otherBox = document.createElement("input");
      otherBox.type = "checkbox";
      const syncOther = () => {
        const n = rest.filter(([value]) => selected.has(value)).length;
        otherBox.checked = rest.length > 0 && n === rest.length;
        otherBox.indeterminate = n > 0 && n < rest.length;
      };
      const updateBar = () => {
        bar.hidden = selected.size === 0;
        barText.textContent = `${selected.size} selected → replace with`;
        barRun.disabled = selected.size === 0 || barInput.value.trim() === "";
      };
      const item = (value, n, children) => {
        const li = document.createElement("li");
        const label = document.createElement("label");
        label.className = "value-item";
        const box = document.createElement("input");
        box.type = "checkbox";
        const text = document.createElement("span");
        text.className = "value-text";
        text.textContent = value;
        text.title = value;
        const num = document.createElement("span");
        num.className = "stat-num";
        num.textContent = `${n} (${pct(n)})`;
        label.append(box, text, num);
        li.append(label);
        if (children) li.append(children);
        return { li, box };
      };
      const add = (target, [value, n]) => {
        const { li, box } = item(value, n);
        boxes.set(value, box);
        box.addEventListener("change", () => {
          if (box.checked) selected.add(value); else selected.delete(value);
          syncOther();
          updateBar();
        });
        target.append(li);
      };
      shown.forEach((entry) => add(list, entry));
      if (rest.length) {
        const total = rest.reduce((sum, [, n]) => sum + n, 0);
        const nested = document.createElement("ul");
        nested.className = "value-list nested";
        nested.hidden = true;
        rest.slice(0, 50).forEach((entry) => add(nested, entry));
        const { li, box } = item(`${rest.length} other`, total);
        li.querySelector(".value-item").replaceChild(otherBox, box);
        const expand = document.createElement("button");
        expand.type = "button";
        expand.className = "link-button expand";
        expand.textContent = "Show";
        expand.setAttribute("aria-expanded", "false");
        expand.addEventListener("click", () => {
          nested.hidden = !nested.hidden;
          expand.textContent = nested.hidden ? "Show" : "Hide";
          expand.setAttribute("aria-expanded", String(!nested.hidden));
        });
        li.querySelector(".value-item").append(expand);
        li.append(nested);
        otherBox.addEventListener("change", () => {
          rest.forEach(([value]) => {
            if (otherBox.checked) selected.add(value); else selected.delete(value);
            const child = boxes.get(value);
            if (child) child.checked = otherBox.checked;
          });
          updateBar();
        });
        list.append(li);
      }
      const run = () => {
        const value = barInput.value;
        if (!selected.size || value.trim() === "") return;
        applyValueMap(new Map([...selected].map((from) => [from, value])));
      };
      barInput.addEventListener("input", updateBar);
      barInput.addEventListener("keydown", (event) => { if (event.key === "Enter") run(); });
      barRun.addEventListener("click", run);
      barClear.addEventListener("click", () => {
        selected.clear();
        boxes.forEach((box) => { box.checked = false; });
        syncOther();
        updateBar();
      });
      body.append(head, list, bar);
    }
    if (fixes.length) {
      const box = document.createElement("details");
      box.className = "spell-fixes";
      const title = document.createElement("summary");
      title.textContent = `Possible spelling fixes (${fixes.length})`;
      box.append(title);
      fixes.slice(0, 8).forEach(({ from, to, count }) => {
        const line = document.createElement("div");
        line.className = "stat-line";
        const text = document.createElement("span");
        text.className = "stat-name";
        text.textContent = `${from} → ${to}?`;
        const fix = document.createElement("button");
        fix.type = "button";
        fix.className = "link-button";
        fix.textContent = `Fix${count > 1 ? ` (${count})` : ""}`;
        fix.addEventListener("click", () => applyValueMap(new Map([[from, to]])));
        line.append(text, fix);
        box.append(line);
      });
      if (fixes.length > 1) {
        const all = document.createElement("button");
        all.type = "button";
        all.className = "link-button";
        all.textContent = "Fix all suggestions";
        all.addEventListener("click", () => applyValueMap(new Map(fixes.map(({ from, to }) => [from, to]))));
        box.append(all);
      }
      body.append(box);
    }
  }

  function applyValueMap(map) {
    const before = cloneTable(currentTable);
    const stats = new Map();
    currentTable.rows.forEach((row) => {
      formatTargets.forEach((header) => {
        const value = String(row[header] ?? "");
        const key = value.trim();
        if (!map.has(key) || map.get(key) === value) return;
        row[header] = map.get(key);
        const entry = stats.get(key) ?? { instances: 0, cells: 0 };
        entry.instances += 1;
        entry.cells += 1;
        stats.set(key, entry);
      });
    });
    if (!stats.size) return;
    commitHistory(before, `Replace ${stats.size} value${stats.size === 1 ? "" : "s"} in ${targetLabel()}`);
    [...stats].reverse().forEach(([from, { instances, cells }]) => {
      replaceLog.unshift({ find: from, replacement: map.get(from), instances, cells, where: targetLabel() });
    });
    setStatus(`Replaced ${stats.size} value${stats.size === 1 ? "" : "s"}.`);
    renderPreview();
    updateFilenameGuidance();
    renderTypes();
    renderValues();
    refreshEdit();
    renderReplaceLog();
  }

  function loadEditControls() {
    const first = formatTargets[0];
    const spec = currentTable.meta.formats.columns[first] ?? {};
    draftSpec = {};
    Object.keys(KINDS).forEach((kind) => { if (spec[kind]) draftSpec[kind] = spec[kind]; });
    draftInvalid = false;
    draftCaps = spec.caps ?? "keep";
    $("#ed-trim").checked = spec.trim !== false;
    const rule = blankRule(currentTable.meta, first);
    $(`input[name="blank-handling"][value="${rule.mode}"]`).checked = true;
    $("#blank-value").value = rule.value;
    $("#blank-edit").hidden = true;
    renderTypes();
    renderValues();
    refreshEdit();
  }

  function refreshEdit() {
    const { spec, invalid } = readSpec();
    const mode = $('input[name="blank-handling"]:checked').value;
    $("#blank-value").classList.toggle("is-inactive", mode !== "fill");
    let blankCells = 0;
    currentTable.rows.forEach((row) => {
      blankCells += formatTargets.filter((header) => isBlank(row[header])).length;
    });
    const mod = $("#blank-mod");
    $("#blank-line").classList.toggle("is-disabled", blankCells === 0);
    $("#blank-count").textContent = `(${blankCells})`;
    mod.hidden = blankCells === 0;
    mod.textContent = mode === "keep" ? "<keep>" : mode === "omit" ? "<omit>" : `<fill: ${$("#blank-value").value}>`;
    if (!blankCells) $("#blank-edit").hidden = true;
    const extra = $("#ed-extra");
    extra.replaceChildren();
    const active = ["date", "phone", "number", "bool"].filter((kind) => spec[kind]);
    if (active.length && !invalid) {
      const failed = [];
      let filled = 0;
      formatTargets.forEach((header) => {
        const phoneColumn = isPhoneName(header);
        currentTable.rows.forEach((row) => {
          const value = String(row[header] ?? "");
          if (isBlank(value)) return;
          filled += 1;
          if (!active.some((kind) => kindTests[kind](value, phoneColumn))) failed.push(value);
        });
      });
      const line = document.createElement("div");
      line.className = "small";
      const button = document.createElement("button");
      button.type = "button";
      button.className = "link-button";
      button.textContent = `${failed.length} of ${filled} values not converted`;
      button.disabled = failed.length === 0;
      const list = document.createElement("ul");
      list.className = "fmt-notconv";
      list.hidden = true;
      [...new Set(failed)].slice(0, 50).forEach((value) => {
        const item = document.createElement("li");
        item.textContent = value;
        list.append(item);
      });
      button.addEventListener("click", () => { list.hidden = !list.hidden; });
      const note = document.createElement("span");
      note.className = "muted";
      note.textContent = " Values that can't be converted stay as they are.";
      line.append(button, note);
      extra.append(line, list);
    }
    $("#fmt-error").textContent = invalid
      ? "A custom format is not valid. Number formats need 0 or #, date formats need yyyy, mm or dd, phone formats need ten #."
      : "";
    $("#fmt-apply").disabled = invalid;
    highlightHeaders = new Set(editAll ? [] : formatTargets);
    highlightBlanks = blankCells ? new Set(formatTargets) : new Set();
    syncHighlights();
    updateReplacePreview();
  }

  const includedHeaders = () => currentTable.headers.filter((header) => !currentTable.meta.excluded.includes(header));
  function openColumnPanel(headers, { all = false } = {}) {
    editAll = all;
    const chosen = headers?.length ? headers : [...selectedColumns];
    formatTargets = currentTable.headers.filter((header) => chosen.includes(header));
    if (!formatTargets.length) return;
    openPanel("format");
    renderEditTop();
    loadEditControls();
    renderReplaceLog();
  }

  function openStatsPanel(headers) {
    const chosen = headers?.length ? headers : [...selectedColumns];
    formatTargets = currentTable.headers.filter((header) => chosen.includes(header));
    if (!formatTargets.length) return;
    openPanel("stats");
    const body = $("#stats-body");
    body.replaceChildren();
    formatTargets.forEach((header) => {
      const section = document.createElement("section");
      section.className = "card stats-section";
      renderStatsInto(section, header);
      body.append(section);
    });
  }

  document.querySelectorAll("#format-panel, #stats-panel").forEach((panel) => {
    const nav = document.createElement("nav");
    nav.className = "panel-tabs";
    nav.setAttribute("role", "tablist");
    [["format", "Edit"], ["stats", "Stats"]].forEach(([key, text]) => {
      const tab = document.createElement("button");
      tab.type = "button";
      tab.dataset.tab = key;
      tab.setAttribute("role", "tab");
      tab.textContent = text;
      tab.addEventListener("click", () => showTab(key));
      nav.append(tab);
    });
    panel.querySelector(".side-head").after(nav);
  });

  function showTab(name) {
    const targets = selectedColumns.size ? [...selectedColumns] : formatTargets;
    if (name === "format") openColumnPanel(targets);
    else openStatsPanel(targets);
  }

  $("#ed-trim").addEventListener("change", refreshEdit);
  document.querySelectorAll('input[name="blank-handling"]').forEach((radio) => radio.addEventListener("change", refreshEdit));
  $("#blank-value").addEventListener("input", () => { $("#blank-fill").checked = true; refreshEdit(); });
  $("#blank-value").addEventListener("focus", () => { $("#blank-fill").checked = true; refreshEdit(); });
  $("#blank-mod").addEventListener("click", () => {
    const edit = $("#blank-edit");
    edit.hidden = !edit.hidden;
    $("#blank-mod").setAttribute("aria-expanded", String(!edit.hidden));
  });
  $("#open-replace").addEventListener("click", () => {
    if (replaceDialog.open) { replaceDialog.close(); return; }
    const panelBox = panels.format.panel.getBoundingClientRect();
    replaceDialog.style.left = `${Math.max(8, panelBox.left - 332)}px`;
    replaceDialog.style.top = `${Math.max(8, panelBox.top + 40)}px`;
    replaceDialog.style.right = "auto";
    replaceDialog.show();
    renderReplaceLog();
    updateReplacePreview();
    $("#rep-find").focus();
  });
  $("#replace-close").addEventListener("click", () => replaceDialog.close());
  replaceDialog.addEventListener("close", () => { highlightCells = new Set(); syncHighlights(); });
  replaceDialog.addEventListener("keydown", (event) => {
    if (event.key === "Escape") { event.stopPropagation(); replaceDialog.close(); }
  });
  document.addEventListener("pointerdown", (event) => {
    if (replaceDialog.open && !event.target.closest("#replace-dialog, #open-replace")) replaceDialog.close();
  });
  $("#edit-all").addEventListener("click", () => {
    if (!panels.format.panel.hidden || !panels.stats.panel.hidden) closePanels();
    else if (selectedColumns.size) openColumnPanel([...selectedColumns]);
    else openColumnPanel(includedHeaders(), { all: true });
  });
  $("#selection-format").addEventListener("click", () => {
    if (panels.format.panel.hidden) openColumnPanel([...selectedColumns]);
    else closePanels();
  });

  function replaceOptions() {
    return { find: $("#rep-find").value, replacement: $("#rep-with").value, whole: $("#rep-whole").checked, matchCase: $("#rep-case").checked };
  }

  function countInstances(text, { find, whole, matchCase }) {
    if (whole) return 1;
    const haystack = matchCase ? text : text.toLocaleLowerCase();
    return haystack.split(matchCase ? find : find.toLocaleLowerCase()).length - 1;
  }

  function replaceMatches() {
    const options = replaceOptions();
    const matches = [];
    if (!options.find) return matches;
    currentTable.rows.forEach((row, rowIndex) => {
      formatTargets.forEach((header) => {
        const next = replaceValue(row[header], options.find, options.replacement, options);
        if (next !== null) matches.push({ rowIndex, header, to: next, instances: countInstances(String(row[header] ?? ""), options) });
      });
    });
    return matches;
  }

  function updateReplacePreview() {
    if (!currentTable || panels.format.panel.hidden || !replaceDialog.open) return;
    const matches = replaceMatches();
    highlightCells = new Set(matches.map(({ rowIndex, header }) => `${rowIndex}|${header}`));
    const { find } = replaceOptions();
    $("#rep-preview").textContent = !find ? "" : matches.length ? `${matches.length} cell${matches.length === 1 ? "" : "s"} match` : "No matching cells";
    $("#rep-apply").disabled = matches.length === 0;
    syncHighlights();
  }

  function renderReplaceLog() {
    const list = $("#rep-log");
    list.replaceChildren();
    replaceLog.forEach(({ find, replacement, instances, cells, where }) => {
      const item = document.createElement("li");
      const change = document.createElement("span");
      change.textContent = `“${find}” → “${replacement}”`;
      const count = document.createElement("small");
      count.className = "muted";
      count.textContent = `${instances} instance${instances === 1 ? "" : "s"} in ${cells} cell${cells === 1 ? "" : "s"} · ${where}`;
      item.append(change, count);
      list.append(item);
    });
  }

  ["#rep-find", "#rep-with", "#rep-whole", "#rep-case"].forEach((selector) => {
    $(selector).addEventListener("input", updateReplacePreview);
    $(selector).addEventListener("change", updateReplacePreview);
  });
  $("#rep-apply").addEventListener("click", () => {
    const matches = replaceMatches();
    if (!matches.length) return;
    const { find, replacement } = replaceOptions();
    const instances = matches.reduce((total, match) => total + match.instances, 0);
    const before = cloneTable(currentTable);
    matches.forEach(({ rowIndex, header, to }) => { currentTable.rows[rowIndex][header] = to; });
    commitHistory(before, `Replace "${find}" with "${replacement}" (${matches.length} cell${matches.length === 1 ? "" : "s"})`);
    replaceLog.unshift({ find, replacement, instances, cells: matches.length, where: targetLabel() });
    setStatus(`Replaced ${instances} instance${instances === 1 ? "" : "s"} in ${matches.length} cell${matches.length === 1 ? "" : "s"}.`);
    $("#rep-find").value = "";
    $("#rep-with").value = "";
    renderPreview();
    updateFilenameGuidance();
    renderTypes();
    renderValues();
    refreshEdit();
    renderReplaceLog();
    $("#rep-find").focus();
  });

  function applyFormat(clear) {
    const before = cloneTable(currentTable);
    const { formats, blank } = currentTable.meta;
    const { spec } = readSpec();
    const mode = $('input[name="blank-handling"]:checked').value;
    formatTargets.forEach((header) => {
      if (clear) {
        delete formats.columns[header];
        delete blank.columns[header];
        return;
      }
      const next = { ...spec };
      if (next.trim !== false) delete next.trim;
      if (Object.keys(next).length) formats.columns[header] = next;
      else delete formats.columns[header];
      const rule = { mode, value: $("#blank-value").value };
      if (rule.mode === (blank.dataset.mode) && rule.value === blank.dataset.value) delete blank.columns[header];
      else blank.columns[header] = rule;
    });
    if (JSON.stringify(before.meta) !== JSON.stringify(currentTable.meta)) {
      commitHistory(before, `${clear ? "Clear changes for" : "Update"} ${targetLabel()}`);
      setStatus(`${clear ? "Cleared changes for" : "Updated"} ${targetLabel()}.`);
    }
    refreshAfterColumnChange();
    loadEditControls();
  }

  $("#fmt-apply").addEventListener("click", () => applyFormat(false));
  $("#fmt-reset").addEventListener("click", () => applyFormat(true));

  function renderFilenameLog() {
    const { files, excludedCount } = planOutput();
    const MAX_ROWS = 5000;
    document.querySelector("#log-summary").textContent =
      `${files.length} file${files.length === 1 ? "" : "s"} will be created${excludedCount ? `; ${excludedCount} row${excludedCount === 1 ? "" : "s"} excluded` : ""}.` +
      (files.length > MAX_ROWS ? ` Showing the first ${MAX_ROWS}.` : "");
    const table = document.querySelector("#log-table");
    const head = document.createElement("thead");
    const headRow = document.createElement("tr");
    ["Row", "File name"].forEach((text) => {
      const th = document.createElement("th");
      th.scope = "col";
      th.textContent = text;
      headRow.append(th);
    });
    head.append(headRow);
    const body = document.createElement("tbody");
    files.slice(0, MAX_ROWS).forEach(({ index, name }, position) => {
      const tr = document.createElement("tr");
      tr.tabIndex = 0;
      tr.dataset.position = String(position);
      tr.classList.toggle("is-selected", position === previewPosition);
      tr.setAttribute("aria-selected", String(position === previewPosition));
      const row = document.createElement("td");
      row.textContent = String(index + 1);
      const file = document.createElement("td");
      file.textContent = name;
      tr.append(row, file);
      body.append(tr);
    });
    table.replaceChildren(head, body);
  }

  const tplEl = document.querySelector("#filename-template");
  const suggestionsEl = document.querySelector("#filename-suggestions");
  const ZW = "\u200b";

  function readTemplate() {
    let out = "";
    tplEl.childNodes.forEach((node) => {
      if (node.nodeType === Node.TEXT_NODE) out += node.nodeValue.replace(/[\u200b\r\n]/g, "");
      else if (node.dataset?.raw !== undefined) out += `{${node.dataset.raw}}`;
      else if (node.tagName !== "BR") out += (node.textContent ?? "").replace(/[\u200b\r\n]/g, "");
    });
    return out;
  }

  function makeChip(token) {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.contentEditable = "false";
    chip.dataset.raw = token.raw;
    chip.dataset.kind = token.kind;
    chip.classList.toggle("is-generated", token.kind === "number");
    const label = document.createElement("span");
    label.className = "chip-label";
    label.textContent = token.raw;
    const help = document.createElement("button");
    help.type = "button";
    help.className = "chip-btn chip-help";
    help.tabIndex = -1;
    help.textContent = "?";
    help.setAttribute("aria-label", `Changes to ${token.raw}`);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "chip-btn chip-x";
    remove.tabIndex = -1;
    remove.textContent = "✕";
    remove.setAttribute("aria-label", `Remove ${token.raw}`);
    chip.append(label, help, remove);
    return chip;
  }

  function padChips() {
    tplEl.querySelectorAll(".chip").forEach((chip) => {
      if (!chip.nextSibling || chip.nextSibling.nodeType !== Node.TEXT_NODE) chip.after(document.createTextNode(ZW));
      if (!chip.previousSibling) chip.before(document.createTextNode(ZW));
    });
  }

  function renderTemplate(text) {
    tplEl.replaceChildren();
    const tokens = filenameTokens(text, currentTable?.headers ?? []).filter((token) => token.kind === "column" || token.kind === "number");
    let position = 0;
    for (const token of tokens) {
      if (token.start > position) tplEl.append(document.createTextNode(text.slice(position, token.start)));
      tplEl.append(makeChip(token));
      position = token.end;
    }
    if (position < text.length) tplEl.append(document.createTextNode(text.slice(position)));
    padChips();
  }

  Object.defineProperty(tplEl, "value", { get: readTemplate, set: renderTemplate, configurable: true });

  function normalizeTemplate() {
    let text = readTemplate();
    const names = new Map([["number", "number"], ...(currentTable?.headers ?? []).map((header) => [header.toLocaleLowerCase(), header])]);
    text = text.replace(/\{([^{}]+)(?=\{|$)/g, (whole, inner) => (names.has(inner.trim().toLocaleLowerCase()) ? `{${names.get(inner.trim().toLocaleLowerCase())}}` : whole));
    renderTemplate(text);
  }

  function placeCaret(node, offset) {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    const selection = getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function placeAfter(chip) {
    const next = chip.nextSibling;
    if (next?.nodeType === Node.TEXT_NODE) placeCaret(next, next.nodeValue.startsWith(ZW) ? 1 : 0);
  }

  function querySelection() {
    const selection = getSelection();
    if (!selection.rangeCount || !tplEl.contains(selection.anchorNode)) return null;
    const range = selection.getRangeAt(0);
    const node = range.startContainer;
    if (node.nodeType !== Node.TEXT_NODE) return null;
    const before = node.nodeValue.slice(0, range.startOffset);
    const open = before.lastIndexOf("{");
    if (open < 0 || before.slice(open).includes("}")) return null;
    const selecting = !range.collapsed && range.endContainer === node;
    return { node, open, end: selecting ? range.endOffset : range.startOffset, prefix: selecting ? "" : before.slice(open + 1), range };
  }

  function updateAssist() {
    const query = currentTable ? querySelection() : null;
    if (!query) {
      suggestionsEl.hidden = true;
      suggestionsEl.replaceChildren();
      return;
    }
    const needle = query.prefix.trim().toLocaleLowerCase();
    const names = ["number", ...currentTable.headers];
    const matches = [
      ...names.filter((name) => name.toLocaleLowerCase().startsWith(needle)),
      ...names.filter((name) => needle && !name.toLocaleLowerCase().startsWith(needle) && name.toLocaleLowerCase().includes(needle)),
    ];
    suggestionsEl.replaceChildren();
    matches.forEach((name) => {
      const item = document.createElement("button");
      item.type = "button";
      item.setAttribute("role", "option");
      item.dataset.value = name;
      const label = document.createElement("span");
      label.className = "suggestion-name";
      label.textContent = name;
      const help = document.createElement("span");
      help.className = "suggestion-help";
      if (name === "number" && !currentTable.headers.includes("number")) {
        item.classList.add("is-generated");
        label.textContent = "number";
        help.textContent = "counter: 1, 2, 3…";
      } else {
        const column = inspectFilenameColumn(name, currentTable.rows);
        const summary = filenameColumnSummary(column);
        const lines = filenameColumnDetails(column);
        if (lines.length) {
          item.classList.add("has-issues");
          help.textContent = summary;
          help.tabIndex = -1;
        }
      }
      const edit = document.createElement("span");
      edit.className = "suggestion-edit";
      edit.dataset.editName = name;
      edit.setAttribute("role", "button");
      edit.setAttribute("aria-label", `Clean-up options for ${name}`);
      edit.innerHTML = PENCIL;
      item.append(edit, label, help);
      suggestionsEl.append(item);
    });
    suggestionsEl.hidden = matches.length === 0;
  }

  function pickSuggestion(name) {
    const query = querySelection();
    const token = filenameTokens(`{${name}}`, currentTable.headers)[0] ?? { raw: name, kind: "column" };
    const chip = makeChip(token);
    if (query) {
      const text = query.node.nodeValue;
      const after = document.createTextNode(text.slice(query.end));
      query.node.nodeValue = text.slice(0, query.open);
      query.node.after(chip, after);
    } else {
      const selection = getSelection();
      if (selection.rangeCount && tplEl.contains(selection.anchorNode)) {
        const range = selection.getRangeAt(0);
        range.deleteContents();
        range.insertNode(chip);
      } else tplEl.append(chip);
    }
    padChips();
    placeAfter(chip);
    tplEl.focus();
    suggestionsEl.hidden = true;
    updateFilenameGuidance();
  }

  function unchip(chip, { dropLast = false } = {}) {
    const raw = dropLast ? chip.dataset.raw.slice(0, -1) : chip.dataset.raw;
    const node = document.createTextNode(`{${raw}`);
    chip.replaceWith(node);
    tplEl.focus();
    const range = document.createRange();
    if (dropLast) range.setStart(node, node.nodeValue.length);
    else range.setStart(node, 1);
    range.setEnd(node, node.nodeValue.length);
    if (dropLast) range.collapse(false);
    const selection = getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    updateAssist();
    updateFilenameGuidance();
  }

  function chipBeforeCaret() {
    const selection = getSelection();
    if (!selection.rangeCount || !selection.isCollapsed || !tplEl.contains(selection.anchorNode)) return null;
    let node = selection.anchorNode;
    let offset = selection.anchorOffset;
    let previous;
    if (node === tplEl) previous = tplEl.childNodes[offset - 1];
    else if (node.nodeType === Node.TEXT_NODE && node.nodeValue.slice(0, offset).replace(/\u200b/g, "") === "") previous = node.previousSibling;
    while (previous?.nodeType === Node.TEXT_NODE && previous.nodeValue.replace(/\u200b/g, "") === "") previous = previous.previousSibling;
    return previous?.classList?.contains("chip") ? previous : null;
  }

  tplEl.addEventListener("input", () => {
    padChips();
    updateAssist();
    updateFilenameGuidance();
  });
  tplEl.addEventListener("keyup", (event) => {
    if (!["Escape", "ArrowDown", "ArrowUp", "Enter"].includes(event.key)) updateAssist();
  });
  tplEl.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      const first = suggestionsEl.querySelector("[data-value]");
      if (!suggestionsEl.hidden && first) pickSuggestion(first.dataset.value);
    } else if (event.key === "Backspace") {
      const chip = chipBeforeCaret();
      if (chip) {
        event.preventDefault();
        unchip(chip, { dropLast: true });
      }
    } else if (event.key === "ArrowDown" && !suggestionsEl.hidden) {
      event.preventDefault();
      suggestionsEl.querySelector("[data-value]")?.focus();
    } else if (event.key === "Escape") {
      suggestionsEl.hidden = true;
    }
  });
  tplEl.addEventListener("paste", (event) => {
    event.preventDefault();
    const text = (event.clipboardData?.getData("text") ?? "").replace(/[\r\n]+/g, " ");
    document.execCommand("insertText", false, text);
  });
  tplEl.addEventListener("blur", () => {
    setTimeout(() => {
      if (document.activeElement === tplEl || suggestionsEl.contains(document.activeElement)) return;
      normalizeTemplate();
      updateFilenameGuidance();
    }, 150);
  });
  tplEl.addEventListener("mousedown", (event) => {
    if (event.target.closest(".chip-btn")) event.preventDefault();
  });
  tplEl.addEventListener("click", (event) => {
    const chip = event.target.closest(".chip");
    if (!chip) {
      updateAssist();
      return;
    }
    const token = { raw: chip.dataset.raw, kind: chip.dataset.kind, column: chip.dataset.raw.replace(/\(\d+\)$/, "") };
    if (event.target.closest(".chip-x")) {
      chip.remove();
      tplEl.focus();
      updateAssist();
      updateFilenameGuidance();
    } else if (event.target.closest(".chip-help")) {
      showChipNote(chip, token);
    } else {
      unchip(chip);
    }
  });
  const chipPop = document.createElement("div");
  chipPop.className = "chip-pop";
  chipPop.hidden = true;
  document.querySelector(".name-field").append(chipPop);
  function showChipNote(chip, token) {
    chipPop.replaceChildren();
    (chip.dataset.note ?? "").split("\n").filter(Boolean).forEach((line) => {
      const p = document.createElement("p");
      p.textContent = line;
      chipPop.append(p);
    });
    const change = document.createElement("button");
    change.type = "button";
    change.className = "link-button";
    change.textContent = "Change";
    change.addEventListener("click", () => {
      chipPop.hidden = true;
      openNameConfig("warnings", token);
    });
    chipPop.append(change);
    const field = chipPop.parentElement.getBoundingClientRect();
    const box = chip.getBoundingClientRect();
    chipPop.style.left = `${Math.max(0, box.left - field.left)}px`;
    chipPop.hidden = false;
  }
  document.addEventListener("pointerdown", (event) => {
    if (!chipPop.hidden && !event.target.closest(".chip-pop, .chip-help")) chipPop.hidden = true;
  });
  suggestionsEl.addEventListener("mousedown", (event) => event.preventDefault());
  suggestionsEl.addEventListener("click", (event) => {
    const edit = event.target.closest("[data-edit-name]");
    if (edit) {
      const name = edit.dataset.editName;
      suggestionsEl.hidden = true;
      if (name === "number" && !currentTable.headers.includes("number")) openNameConfig("format");
      else openNameConfig("warnings", { raw: name, kind: "column", column: name });
      return;
    }
    const option = event.target.closest("[data-value]");
    if (option) pickSuggestion(option.dataset.value);
  });
  suggestionsEl.addEventListener("keydown", (event) => {
    const options = [...suggestionsEl.querySelectorAll("[data-value]")];
    const index = options.indexOf(document.activeElement);
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      options[(index + (event.key === "ArrowDown" ? 1 : -1) + options.length) % options.length]?.focus();
    } else if (event.key === "Escape") {
      suggestionsEl.hidden = true;
      tplEl.focus();
    } else if (event.key === "Enter") {
      event.preventDefault();
      if (index >= 0) pickSuggestion(options[index].dataset.value);
    }
  });
  renderTemplate("Document-{number}");

  function showNameTab(name) {
    nameConfig.querySelectorAll("[data-section]").forEach((node) => {
      node.hidden = node.dataset.section !== name;
    });
    nameConfig.querySelectorAll("[data-tab]").forEach((tab) => tab.setAttribute("aria-selected", String(tab.dataset.tab === name)));
  }

  function openNameConfig(tab = "format", token = null) {
    document.querySelector("#filename-suggestions").hidden = true;
    configToken = token;
    const info = token?.kind === "column" ? inspectFilenameColumn(token.column, currentTable.rows) : null;
    document.querySelector("#name-config-title").textContent = token?.kind === "column" ? `Clean up: ${token.column}` : "Clean up file names";
    document.querySelector("#nc-column-info").textContent = info
      ? `Values in this column range from ${info.shortest} to ${info.longest} characters long.`
      : "";
    showNameTab(tab);
    nameConfig.hidden = false;
  }

  function closeNameConfig() {
    if (nameConfig.hidden) return;
    nameConfig.hidden = true;
    configToken = null;
  }

  nameConfig.querySelectorAll("[data-tab]").forEach((tab) => tab.addEventListener("click", () => showNameTab(tab.dataset.tab)));
  document.querySelector("#name-config-close").addEventListener("click", () => {
    closeNameConfig();
    document.querySelector("#filename-template").focus();
  });
  nameConfig.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeNameConfig();
      document.querySelector("#filename-template").focus();
    }
  });
  function addRuleRow(find = "", to = "") {
    const item = document.createElement("li");
    const label = document.createElement("span");
    label.textContent = "Replace";
    const findInput = document.createElement("input");
    findInput.type = "text";
    findInput.className = "rule-find";
    findInput.placeholder = "text";
    findInput.value = find;
    findInput.setAttribute("aria-label", "Text to find");
    const arrow = document.createElement("span");
    arrow.textContent = "with";
    const withInput = document.createElement("input");
    withInput.type = "text";
    withInput.className = "rule-with";
    withInput.setAttribute("aria-label", "Replacement text");
    withInput.value = to;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "icon-button";
    remove.textContent = "✕";
    remove.setAttribute("aria-label", "Remove this rule");
    remove.addEventListener("click", () => {
      item.remove();
      updateFilenameGuidance();
    });
    [findInput, withInput].forEach((input) => input.addEventListener("input", updateFilenameGuidance));
    item.append(label, findInput, arrow, withInput, remove);
    document.querySelector("#nc-rules").append(item);
    findInput.focus();
  }
  document.querySelector("#nc-add-rule").addEventListener("click", () => addRuleRow());
  function shortenValues() {
    limitTouched = true;
    document.querySelector("#nc-limit-all").checked = true;
    if (!Number(document.querySelector("#nc-limit-default").value)) document.querySelector("#nc-limit-default").value = "25";
    updateFilenameGuidance();
    openNameConfig("format");
  }

  function updateFilenameGuidance() {
    if (!currentTable) return;
    if (!limitTouched) {
      const used = filenameTokens(document.querySelector("#filename-template").value, currentTable.headers).filter((token) => token.kind === "column");
      if (used.some((token) => inspectFilenameColumn(token.column, currentTable.rows).longest > 25)) {
        limitTouched = true;
        document.querySelector("#nc-limit-all").checked = true;
        document.querySelector("#nc-limit-default").value = "25";
      }
    }
    const template = document.querySelector("#filename-template").value;
    const guidance = document.querySelector("#filename-guidance");
    const options = readNameOptions();
    const report = inspectFilenameTemplate(currentTable.headers, currentTable.rows, template, options);
    const plan = planOutput();
    if (previewPosition >= plan.files.length) previewPosition = 0;
    const example = plan.files[previewPosition];
    document.querySelector("#filename-preview").textContent = example ? example.name : "No rows to export";
    const rowLabel = document.querySelector("#preview-row-label");
    rowLabel.textContent = example ? `row ${example.index + 1}` : "no rows";
    document.querySelector("#preview-prev").disabled = document.querySelector("#preview-next").disabled = plan.files.length < 2;
    document.querySelector("#filename-stats").textContent = plan.files.length
      ? `${plan.shortest === plan.longest ? plan.longest : `${plan.shortest}–${plan.longest}`} characters`
      : "";
    const count = plan.files.length;
    document.querySelector("#download-count").textContent =
      `${count * (outputKind === "both" ? 2 : 1)} ${outputKind === "pdf" ? "PDF" : outputKind === "docx" ? "Word document" : "document"}${count * (outputKind === "both" ? 2 : 1) === 1 ? "" : "s"} will be created` +
      (extraFiles.length ? ` + ${extraFiles.length} other file${extraFiles.length === 1 ? "" : "s"}` : "") +
      (plan.excludedCount ? ` · ${plan.excludedCount} row${plan.excludedCount === 1 ? "" : "s"} excluded` : "");
    document.querySelector("#download-button").disabled = count === 0;

    const sep = options.separator || "-";
    const messages = [];
    const add = (text, kind, action) => messages.push({ text, kind, action });
    if (report.unknown.length) add(`Unknown fields: ${report.unknown.join(", ")}. Use a suggested field or check its spelling.`, "warning");
    const repeats = `${report.collisions} filename${report.collisions === 1 ? "" : "s"} repeat`;
    const changeRepeats = { label: "Change", run: () => openNameConfig("format") };
    if (report.collisions && options.numbering === "none") {
      add(`${repeats}. Numbering is off, so the ZIP cannot be created.`, "warning", changeRepeats);
    } else if (report.collisions) {
      const example = options.numbering === "repeats" ? `Name${sep}2` : `Name${sep}1, Name${sep}2`;
      add(`${repeats} and will be numbered (for example ${example}).`, "note", changeRepeats);
    } else if (options.numbering === "all") {
      add(`Every filename will end in a number (for example Name${sep}1).`, "note", changeRepeats);
    }
    if (report.trimmed && options.trim) {
      add(`Extra spaces were trimmed from ${report.trimmed} name${report.trimmed === 1 ? "" : "s"}.`, "note", {
        label: "Keep spaces",
        run: () => {
          document.querySelector("#nc-trim").checked = false;
          updateFilenameGuidance();
        },
      });
    }
    if (report.special) {
      add(`Invalid characters were replaced with "${options.unsafe}" in ${report.special} name${report.special === 1 ? "" : "s"}.`, "note", {
        label: "Change",
        run: () => openNameConfig("replace"),
      });
    }
    if (report.long) {
      const hasColumns = report.columns.length > 0;
      add(`${report.long} name${report.long === 1 ? " is" : "s are"} over ${MAX_BASE_LENGTH} characters and will be cut off at the end.`, "note",
        hasColumns ? { label: "Shorten values", run: shortenValues } : null);
    }
    const warningsBox = document.querySelector("#nc-warnings");
    warningsBox.replaceChildren();
    const templateTokens = filenameTokens(template, currentTable.headers);
    const columnNotes = new Map();
    report.columns.forEach((column) => {
      const block = document.createElement("div");
      block.className = "warn-block";
      const title = document.createElement("strong");
      title.textContent = column.name;
      block.append(title);
      const tokenLimit = templateTokens.find((token) => token.kind === "column" && token.column === column.name && token.limit)?.limit ?? 0;
      const caps = [options.limit, tokenLimit].filter(Boolean);
      const effective = caps.length ? Math.min(...caps) : Infinity;
      const issues = [];
      if (column.blanks) issues.push({ ok: false, text: `${column.blanks} blank value${column.blanks === 1 ? "" : "s"}` });
      if (column.repeatedValues.length) {
        issues.push(options.numbering === "none"
          ? { ok: false, text: `${column.repeatedValues.length} repeated value${column.repeatedValues.length === 1 ? "" : "s"}. Numbering is off.` }
          : { ok: true, text: `${column.repeatedValues.length} repeated value${column.repeatedValues.length === 1 ? "" : "s"}, resolved by numbering` });
      }
      if (column.special) issues.push({ ok: true, text: `${column.special} value${column.special === 1 ? "" : "s"} with invalid characters, ${options.unsafe ? `replaced with "${options.unsafe}"` : "removed"}` });
      if (column.longest > 25) {
        issues.push(column.longest <= effective
          ? { ok: false, text: `Longest value is ${column.longest} characters` }
          : { ok: true, text: `Longest value is ${column.longest} characters, shortened to ${effective}` });
      }
      if (issues.length) columnNotes.set(column.name, issues.map(({ text }) => `${text}.`).join("\n"));
      if (!issues.length) issues.push({ ok: true, text: "No issues" });
      issues.forEach(({ ok, text }) => {
        const row = document.createElement("div");
        row.className = `warn-item ${ok ? "is-resolved" : "is-open"}`;
        row.textContent = `${ok ? "✓" : "⚠"} ${text}`;
        block.append(row);
      });
      const details = filenameColumnDetails(column);
      if (details.length) {
        const more = document.createElement("details");
        const summary = document.createElement("summary");
        summary.textContent = "Values";
        more.append(summary);
        details.forEach((line) => {
          const item = document.createElement("div");
          item.className = "hint";
          item.textContent = line;
          more.append(item);
        });
        block.append(more);
      }
      warningsBox.append(block);
    });
    const repeatNote = document.querySelector("#nc-repeat-note");
    repeatNote.textContent = report.collisions ? `${report.collisions} filename${report.collisions === 1 ? "" : "s"} currently repeat.` : "No repeated names right now. These settings apply only if names repeat later.";
    const sampleOptions = { position: options.numberPosition, start: options.repeatStart ?? 1, pad: options.numberPad };
    document.querySelector("#nc-repeat-example").textContent = assignFilenames(["Smith", "Smith", "Smith"], options.numbering === "none" ? "repeats" : options.numbering, options.separator, sampleOptions).join(", ");
    if (!report.columns.length) warningsBox.textContent = "No columns are used in the name yet.";
    tplEl.querySelectorAll('.chip[data-kind="column"]').forEach((chip) => {
      const note = columnNotes.get(chip.dataset.raw.replace(/\(\d+\)$/, ""));
      chip.classList.toggle("is-changed", Boolean(note));
      if (note) chip.dataset.note = note;
      else delete chip.dataset.note;
    });
    guidance.replaceChildren();
    const notes = messages.filter(({ kind }) => kind !== "warning");
    let changes = null;
    if (notes.length) {
      changes = document.createElement("details");
      changes.className = "fn-changes";
      const summary = document.createElement("summary");
      summary.textContent = `${notes.length} change${notes.length === 1 ? "" : "s"} to names`;
      changes.append(summary);
    }
    messages.forEach(({ text, kind, action }) => {
      const item = document.createElement("p");
      item.className = kind === "warning" ? "guidance-warning" : kind === "ok" ? "guidance-ok" : "guidance-note";
      item.append(text);
      if (action) {
        item.append(" ");
        const button = document.createElement("button");
        button.type = "button";
        button.className = "link-button";
        button.textContent = action.label;
        button.addEventListener("click", action.run);
        item.append(button);
      }
      (kind === "warning" ? guidance : changes).append(item);
    });
    if (changes) guidance.append(changes);
  }
}

function defaultTemplate(headers) {
  return headers.includes("First name") && headers.includes("Last name")
    ? "{First name}-{Last name}"
    : "Document-{number}";
}

function renderTable() {
  renderColumnsPanel();
  renderPreview();
}

function renderColumnsPanel() {
  const list = document.querySelector("#columns-list");
  if (!list || !currentTable) return;
  list.replaceChildren();
  currentTable.headers.forEach((header, index) => {
    const item = document.createElement("li");
    item.dataset.source = header;
    const box = document.createElement("input");
    box.type = "checkbox";
    box.className = "col-check";
    box.checked = !currentTable.meta.excluded.includes(header);
    box.setAttribute("aria-label", `Include ${header} in the output`);
    const name = document.createElement("input");
    name.type = "text";
    name.className = "col-rename";
    name.setAttribute("aria-label", `Name of ${header} in the output`);
    name.autocomplete = "off";
    const revert = document.createElement("button");
    revert.type = "button";
    revert.className = "rename-revert";
    revert.textContent = "↶";
    const handle = document.createElement("button");
    handle.type = "button";
    handle.className = "drag-handle";
    handle.draggable = true;
    handle.dataset.header = header;
    handle.textContent = "⠿";
    handle.setAttribute("aria-label", `Reorder ${header}. Use the up and down arrow keys.`);
    item.append(box, name, revert, handle);
    list.append(item);
    paintColumnItem(item);
  });
  const included = currentTable.headers.length - currentTable.meta.excluded.length;
  document.querySelector("#columns-summary").textContent = `${included} of ${currentTable.headers.length} included`;
}

function paintColumnItem(item) {
  const header = item.dataset.source;
  const label = currentTable.meta.labels[header];
  const renamed = Boolean(label);
  item.classList.toggle("is-renamed", renamed);
  item.classList.toggle("is-hidden-col", currentTable.meta.excluded.includes(header));
  const name = item.querySelector(".col-rename");
  name.value = label ?? header;
  const revert = item.querySelector(".rename-revert");
  revert.hidden = !renamed;
  revert.setAttribute("aria-label", `Restore the original name "${header}"`);
}

function columnCounts(header) {
  let numbers = 0;
  let dates = 0;
  let phones = 0;
  const phoneColumn = isPhoneName(header);
  for (const row of currentTable.rows) {
    const raw = String(row[header] ?? "");
    if (parseDate(raw)) dates += 1;
    else if (phoneColumn && parsePhoneParts(raw)) phones += 1;
    else if (parseNumber(raw) !== null) numbers += 1;
  }
  return { numbers, dates, phones };
}

function describeCounts({ numbers, dates, phones = 0 }) {
  const parts = [];
  if (phones) parts.push(`${phones} phone number${phones === 1 ? "" : "s"}`);
  if (numbers) parts.push(`${numbers} number${numbers === 1 ? "" : "s"}`);
  if (dates) parts.push(`${dates} date${dates === 1 ? "" : "s"}`);
  return parts.length ? parts.join(", ") : "no numbers or dates";
}

function describeCountsText(counts) {
  return counts.numbers || counts.dates || counts.phones ? describeCounts(counts) : "numbers, dates and phone numbers";
}

function columnHasFormat(header) {
  const { formats } = currentTable.meta;
  const spec = formats.columns[header];
  if (spec?.number || spec?.date || spec?.phone) return true;
  return Object.keys(formats.cells).some((key) => key.slice(key.indexOf("|") + 1) === header);
}

function selectModeOn() {
  return selectMode || selectedColumns.size > 0 || selectedRows.size > 0;
}

function syncSelectMode() {
  const on = selectModeOn();
  document.querySelector("#table-wrap").classList.toggle("select-mode", on);
  document.querySelector("#preview-table th.corner")?.setAttribute("aria-pressed", String(on));
}

function syncSelection() {
  const table = currentTable;
  if (!table) return;
  selectedColumns = new Set([...selectedColumns].filter((header) => table.headers.includes(header)));
  document.querySelectorAll("#preview-table th[data-header]").forEach((th) => {
    const on = selectedColumns.has(th.dataset.header);
    th.classList.toggle("is-selected", on);
    th.setAttribute("aria-selected", String(on));
  });
  document.querySelectorAll("#preview-table td[data-header]").forEach((td) => {
    td.classList.toggle("col-selected", selectedColumns.has(td.dataset.header));
  });
  const bar = document.querySelector("#selection-bar");
  bar.hidden = selectedColumns.size === 0;
  const n = selectedColumns.size;
  document.querySelector("#selection-count").textContent = `${n} column${n === 1 ? "" : "s"} selected`;
  const allHidden = n > 0 && [...selectedColumns].every((header) => table.meta.excluded.includes(header));
  document.querySelector("#selection-hide").textContent = allHidden ? "Include in output" : "Exclude from output";
  syncSelectMode();
}

function syncRowSelection() {
  const table = currentTable;
  if (!table) return;
  selectedRows = new Set([...selectedRows].filter((index) => index < table.rows.length));
  document.querySelectorAll("#preview-table tbody tr[data-row-index]").forEach((tr) => {
    const on = selectedRows.has(Number(tr.dataset.rowIndex));
    tr.classList.toggle("row-selected", on);
    const box = tr.querySelector('th[scope="row"] input');
    if (box) box.checked = on;
  });
  const n = selectedRows.size;
  document.querySelector("#row-selection-bar").hidden = n === 0;
  document.querySelector("#row-selection-count").textContent = `${n} row${n === 1 ? "" : "s"} selected`;
  const allOut = n > 0 && [...selectedRows].every((index) => table.meta.excludedRows.includes(index));
  document.querySelector("#row-selection-exclude").textContent = allOut ? "Include in output" : "Exclude from output";
  syncSelectMode();
  updateFooter();
}

function updateFooter() {
  const table = currentTable;
  if (!table) return;
  const count = (n, word) => `${n.toLocaleString()} ${word}${n === 1 ? "" : "s"}`;
  const parts = [count(table.rows.length, "row")];
  if (selectedRows.size) parts.push(`${selectedRows.size.toLocaleString()} selected`);
  if (focusedRow !== null && focusedRow < table.rows.length) parts.push(`Row ${focusedRow + 1} focused`);
  if (table.meta.excludedRows.length) parts.push(`${table.meta.excludedRows.length.toLocaleString()} excluded`);
  if (table.meta.excluded.length) parts.push(`${count(table.meta.excluded.length, "column")} hidden`);
  if (loadStats?.emptyRowsDropped) parts.push(`${count(loadStats.emptyRowsDropped, "empty row")} skipped`);
  if (rowLimit < table.rows.length) parts.push(`showing first ${rowLimit.toLocaleString()}`);
  document.querySelector("#table-summary").textContent = parts.join(" · ");
}

function syncHighlights() {
  document.querySelectorAll("#preview-table th[data-header]").forEach((th) => {
    th.classList.toggle("fmt-target", highlightHeaders.has(th.dataset.header));
  });
  document.querySelectorAll("#preview-table td[data-row-index]").forEach((td) => {
    const { rowIndex, header } = td.dataset;
    td.classList.toggle("fmt-hit", highlightCells.has(`${rowIndex}|${header}`));
    td.classList.toggle("blank-hit", highlightBlanks.has(header) && isBlank(currentTable.rows[Number(rowIndex)]?.[header]));
  });
}

function paintCell(cell) {
  const rowIndex = Number(cell.dataset.rowIndex);
  const header = cell.dataset.header;
  const raw = String(currentTable.rows[rowIndex][header] ?? "");
  const original = String(originalTable?.rows[rowIndex]?.[header] ?? raw);
  const shown = formatValue(raw, resolveSpec(currentTable.meta.formats, header, rowIndex));
  const edited = shown !== original;
  const formatted = shown !== raw;
  cell.textContent = shown;
  cell.classList.toggle("is-edited", edited);
  cell.classList.toggle("is-formatted", formatted);
  const notes = [];
  if (edited) notes.push(`Original value: ${original === "" ? "(blank)" : original}`);
  if (formatted) notes.push(`Value in data: ${raw}`);
  if (notes.length) cell.title = notes.join("\n");
  else cell.removeAttribute("title");
}

function renderPreview() {
  const table = currentTable;
  if (!table) return;
  const { labels, excluded } = table.meta;
  const preview = document.querySelector("#preview-table");
  preview.replaceChildren();
  const head = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const rowNumberHeader = document.createElement("th");
  rowNumberHeader.scope = "col";
  rowNumberHeader.className = "corner";
  rowNumberHeader.tabIndex = 0;
  rowNumberHeader.setAttribute("aria-label", "Select rows and columns");
  rowNumberHeader.innerHTML = '<span class="corner-box" aria-hidden="true"></span>';
  headerRow.append(rowNumberHeader);
  const runs = [];
  table.headers.forEach((header, index) => {
    if (!excluded.includes(header)) return;
    const last = runs[runs.length - 1];
    if (last && last.end === index - 1) { last.end = index; last.names.push(header); }
    else runs.push({ start: index, end: index, names: [header] });
  });
  const markerAt = (index) => runs.find((run) => run.start === index);
  const makeMarker = (tag, run) => {
    const cell = document.createElement(tag);
    cell.className = "hidden-marker";
    cell.dataset.hidden = run.names.join("\u0001");
    cell.setAttribute("aria-label", `${run.names.length} hidden column${run.names.length === 1 ? "" : "s"}: ${run.names.join(", ")}`);
    cell.tabIndex = 0;
    cell.textContent = run.names.length > 1 ? String(run.names.length) : "";
    return cell;
  };
  table.headers.forEach((header, headerIndex) => {
    const run = markerAt(headerIndex);
    if (run) {
      headerRow.append(makeMarker("th", run));
    }
    if (excluded.includes(header)) return;
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.dataset.header = header;
    cell.classList.toggle("is-excluded", excluded.includes(header));
    cell.tabIndex = 0;
    cell.draggable = true;
    cell.classList.toggle("is-renamed", Boolean(labels[header]));
    const wrap = document.createElement("div");
    wrap.className = "col-head";
    const name = document.createElement("span");
    name.className = "col-name";
    name.textContent = labels[header] ?? header;
    const formatted = columnHasFormat(header);
    wrap.append(name);
    {
      const check = document.createElement("button");
      check.type = "button";
      check.className = "col-check";
      check.dataset.checkColumn = header;
      check.setAttribute("aria-label", `Select ${header}`);
      wrap.prepend(check);
      if (formatted) {
        const mark = document.createElement("span");
        mark.className = "col-mark";
        mark.textContent = "ƒ";
        mark.setAttribute("aria-label", "Formatted");
        wrap.append(mark);
      }
      const hide = document.createElement("button");
      hide.type = "button";
      hide.className = "col-exclude";
      hide.dataset.excludeColumn = header;
      hide.textContent = "✕";
      hide.setAttribute("aria-label", `Exclude ${header} from output`);
      const edit = document.createElement("button");
      edit.type = "button";
      edit.className = "col-edit";
      edit.dataset.editColumn = header;
      edit.setAttribute("aria-label", `Edit ${header}`);
      edit.innerHTML = PENCIL;
      wrap.append(edit, hide);
    }
    const rule = table.meta.blank.columns[header];
    if (rule && rule.mode !== "keep") {
      const mark = document.createElement("span");
      mark.className = "col-mark";
      mark.textContent = "∅";
      mark.setAttribute("aria-label", rule.mode === "omit" ? "Blank cells left out" : `Blank cells show ${rule.value}`);
      wrap.append(mark);
    }
    cell.append(wrap);
    headerRow.append(cell);
  });
  const addHead = document.createElement("th");
  addHead.className = "add-col";
  const add = document.createElement("button");
  add.type = "button";
  add.className = "add-column";
  add.textContent = "+";
  add.setAttribute("aria-label", "Open the Columns panel");
  addHead.append(add);
  headerRow.append(addHead);
  head.append(headerRow);
  const body = document.createElement("tbody");
  const rowsOut = new Set(table.meta.excludedRows);
  table.rows.slice(0, rowLimit).forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    const rowNumber = document.createElement("th");
    rowNumber.scope = "row";
    const isOut = rowsOut.has(rowIndex);
    rowNumber.dataset.rowIndex = String(rowIndex);
    const rowBox = document.createElement("input");
    rowBox.type = "checkbox";
    rowBox.tabIndex = -1;
    rowBox.checked = selectedRows.has(rowIndex);
    rowBox.setAttribute("aria-label", `Select row ${rowIndex + 1}`);
    const rowText = document.createElement("span");
    rowText.className = "row-num";
    rowText.textContent = String(rowIndex + 1);
    const rowX = document.createElement("button");
    rowX.type = "button";
    rowX.className = "row-x";
    rowX.textContent = isOut ? "↺" : "✕";
    rowX.setAttribute("aria-label", isOut ? "Include this row in the output" : "Exclude this row from the output");
    rowNumber.append(rowBox, rowText, rowX);
    tr.append(rowNumber);
    tr.dataset.rowIndex = String(rowIndex);
    tr.classList.toggle("is-excluded-row", rowsOut.has(rowIndex));
    table.headers.forEach((header, headerIndex) => {
      const run = markerAt(headerIndex);
      if (run) {
        const filler = makeMarker("td", run);
        filler.removeAttribute("title");
        filler.tabIndex = -1;
        tr.append(filler);
      }
      if (excluded.includes(header)) return;
      const cell = document.createElement("td");
      cell.contentEditable = "true";
      cell.spellcheck = false;
      cell.dataset.rowIndex = String(rowIndex);
      cell.dataset.header = header;
      cell.setAttribute("aria-label", `Row ${rowIndex + 1}, ${header}. Editable.`);
      cell.classList.toggle("is-excluded", excluded.includes(header));
      paintCell(cell);
      tr.append(cell);
    });
    body.append(tr);
  });
  preview.append(head, body);
  syncHighlights();
  document.querySelector("#row-limit-wrap").hidden = table.rows.length <= 250;
  document.querySelector("#columns-button").textContent = `Columns (${table.headers.length - excluded.length}/${table.headers.length})`;
  document.querySelector("#rows-button").firstChild.textContent = `Rows (${table.rows.length - rowsOut.size}/${table.rows.length})`;
  syncSelection();
  syncRowSelection();
}

function cloneTable(table) {
  const meta = JSON.parse(JSON.stringify(table.meta ?? { labels: {}, excluded: [], formats: emptyFormats() }));
  meta.excludedRows ??= [];
  meta.blank ??= { dataset: { ...DEFAULT_BLANK }, columns: {} };
  return {
    headers: [...table.headers],
    rows: table.rows.map((row) => ({ ...row })),
    meta,
  };
}

function commitHistory(before, label) {
  undoStates.push({ table: before, label });
  if (undoStates.length > 100) undoStates.shift();
  redoStates = [];
  syncHistoryButtons();
}

function applyHistory(from, to) {
  if (!from.length) return;
  const entry = from.pop();
  to.push({ table: cloneTable(currentTable), label: entry.label });
  currentTable = cloneTable(entry.table);
  renderColumnsPanel();
  renderPreview();
  syncHistoryButtons();
  updateFilenameGuidance();
}

function syncHistoryButtons() {
  const undo = document.querySelector("#undo-button");
  if (!undo) return;
  const redo = document.querySelector("#redo-button");
  undo.disabled = undoStates.length === 0;
  redo.disabled = redoStates.length === 0;
  const undoTitle = undoStates.length ? `Undo: ${undoStates.at(-1).label}` : "Nothing to undo";
  const redoTitle = redoStates.length ? `Redo: ${redoStates.at(-1).label}` : "Nothing to redo";
  undo.title = undoTitle;
  undo.setAttribute("aria-label", undoTitle);
  redo.title = redoTitle;
  redo.setAttribute("aria-label", redoTitle);
  document.querySelector("#revert-button").disabled =
    !originalTable || JSON.stringify(currentTable) === JSON.stringify(originalTable);
  const history = document.querySelector("#changes-button");
  history.setAttribute("aria-label", `History (${undoStates.length} change${undoStates.length === 1 ? "" : "s"})`);
  history.title = `History (${undoStates.length})`;
  const list = document.querySelector("#changes-list");
  list.replaceChildren();
  [...undoStates].reverse().forEach((entry, index) => {
    const item = document.createElement("li");
    const text = document.createElement("span");
    text.textContent = entry.label;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "icon-button";
    button.dataset.undoCount = String(index + 1);
    button.textContent = "↶";
    button.title = index === 0 ? "Undo this change" : "Undo this and all later changes";
    button.setAttribute("aria-label", `${button.title}: ${entry.label}`);
    item.append(text, button);
    list.append(item);
  });
  if (!undoStates.length) list.textContent = "No changes yet.";
}
