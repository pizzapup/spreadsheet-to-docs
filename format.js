export const NUMBER_PRESETS = [
  { pattern: "0", label: "1235" },
  { pattern: "0.00", label: "1234.57" },
  { pattern: "#,##0", label: "1,235" },
  { pattern: "#,##0.00", label: "1,234.57" },
  { pattern: "$#,##0.00", label: "$1,234.57", group: "Currency" },
  { pattern: "$#,##0", label: "$1,235", group: "Currency" },
  { pattern: "0%", label: "12%", group: "Percent" },
  { pattern: "0.0%", label: "12.3%", group: "Percent" },
];

export const DATE_PRESETS = [
  { pattern: "mm/dd/yyyy", label: "01/31/2026" },
  { pattern: "dd/mm/yyyy", label: "31/01/2026" },
  { pattern: "yyyy-mm-dd", label: "2026-01-31" },
  { pattern: "mmm d, yyyy", label: "Jan 31, 2026" },
  { pattern: "mmmm d, yyyy", label: "January 31, 2026" },
  { pattern: "d mmm yyyy", label: "31 Jan 2026" },
  { pattern: "yyyy", label: "2026", group: "Year or month only" },
  { pattern: "mmmm", label: "January", group: "Year or month only" },
  { pattern: "mmmm yyyy", label: "January 2026", group: "Year or month only" },
];

export const PHONE_PRESETS = [
  { pattern: "(###) ###-####", label: "(555) 123-4567" },
  { pattern: "###-###-####", label: "555-123-4567" },
  { pattern: "###.###.####", label: "555.123.4567" },
  { pattern: "+1 ### ### ####", label: "+1 555 123 4567" },
  { pattern: "##########", label: "5551234567" },
];

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

export const BOOL_PRESETS = [
  { pattern: "True/False", label: "True, False" },
  { pattern: "Yes/No", label: "Yes, No" },
  { pattern: "Y/N", label: "Y, N" },
  { pattern: "T/F", label: "T, F" },
  { pattern: "1/0", label: "1, 0" },
];

export function parseBool(text) {
  const value = String(text ?? "").trim().toLowerCase();
  if (["true", "yes", "y", "t", "1"].includes(value)) return true;
  if (["false", "no", "n", "f", "0"].includes(value)) return false;
  return null;
}

export function validBoolPattern(pattern) {
  return BOOL_PRESETS.some((preset) => preset.pattern === pattern);
}

export function formatBool(flag, pattern) {
  const [yes, no] = pattern.split("/");
  return flag ? yes : no;
}

function editDistance(a, b) {
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i += 1) {
    const row = [i];
    for (let j = 1; j <= b.length; j += 1) {
      row[j] = Math.min(prev[j] + 1, row[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = row;
  }
  return prev[b.length];
}

// Rare values that look like a typo of a more common value in the same data.
export function similarValues(values) {
  const counts = new Map();
  values.forEach((value) => {
    const text = String(value ?? "").trim();
    if (text) counts.set(text, (counts.get(text) ?? 0) + 1);
  });
  const found = [];
  for (const [from, count] of counts) {
    if (from.length < 4 || /^[\d\s.,:/-]+$/.test(from)) continue;
    let best = null;
    for (const [to, other] of counts) {
      if (to === from || other <= count) continue;
      const a = from.toLocaleLowerCase();
      const b = to.toLocaleLowerCase();
      const limit = Math.min(2, Math.max(1, Math.floor(Math.min(a.length, b.length) / 5)));
      if (Math.abs(a.length - b.length) > limit) continue;
      const distance = a === b ? 0 : editDistance(a, b);
      if (distance <= limit && (!best || distance < best.distance || (distance === best.distance && other > best.other))) best = { to, other, distance };
    }
    if (best) found.push({ from, to: best.to, count });
  }
  return found.sort((x, y) => y.count - x.count || x.from.localeCompare(y.from));
}

export function parseNumber(text) {
  let value = String(text ?? "").trim();
  if (!value || /^0\d/.test(value.replace(/^[-+(]?\$?/, ""))) return null;
  let negative = false;
  const parens = /^\((.*)\)$/.exec(value);
  if (parens) {
    negative = true;
    value = parens[1].trim();
  }
  const match = /^([-+]?)\s*\$?\s*([-+]?)(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?\s*(%?)$/.exec(value);
  if (!match) return null;
  const digits = match[3].replaceAll(",", "");
  if (digits.length > 15) return null;
  let number = Number(digits + (match[4] ?? ""));
  if (match[5]) number /= 100;
  if (match[1] === "-" || match[2] === "-") negative = !negative;
  return negative ? -number : number;
}

const TIME_SUFFIX = /(?:[T ]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?\s*(?:[AaPp][Mm])?\s*(?:Z|[+-]\d{2}:?\d{2}|[A-Z]{2,4})?)$/;

export function parseDate(text) {
  const value = String(text ?? "")
    .trim()
    .replace(/^(?:mon|tue|tues|wed|thu|thur|thurs|fri|sat|sun)[a-z]*\.?,?\s+/i, "")
    .replace(TIME_SUFFIX, "")
    .replace(/,\s*$/, "")
    .trim();
  let year;
  let month;
  let day;
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(value);
  if (match) {
    [year, month, day] = [match[1], match[2], match[3]].map(Number);
  } else if ((match = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/.exec(value))) {
    [year, month, day] = [match[1], match[2], match[3]].map(Number);
  } else if ((match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(value))) {
    [month, day, year] = [match[1], match[2], match[3]].map(Number);
  } else if ((match = /^([A-Za-z]{3,9})\.? (\d{1,2}),? (\d{4})$/.exec(value))) {
    month = monthIndex(match[1]);
    day = Number(match[2]);
    year = Number(match[3]);
  } else if ((match = /^(\d{1,2}) ([A-Za-z]{3,9})\.?,? (\d{4})$/.exec(value))) {
    day = Number(match[1]);
    month = monthIndex(match[2]);
    year = Number(match[3]);
  } else {
    return null;
  }
  if (!month) return null;
  const check = new Date(Date.UTC(year, month - 1, day));
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) return null;
  return { year, month, day };
}

function monthIndex(name) {
  const lower = name.toLocaleLowerCase();
  const index = MONTHS.findIndex((month) => month.toLocaleLowerCase() === lower || (lower.length === 3 && month.toLocaleLowerCase().startsWith(lower)));
  return index + 1;
}

export function isPhoneName(name) {
  return /phone|mobile|\bcell\b|cellular|\bhome\b|\btel\b|telephone|\bfax\b/i.test(String(name ?? ""));
}

// Accepts 7 and 10 digit numbers, a leading 1, and extensions ("x12", "ext. 12", or a trailing extra group).
export function parsePhoneParts(text) {
  const value = String(text ?? "").trim();
  if (!/^\+?[\d\s().-]+(?:\s*(?:x|ext\.?|extension)\s*\d+)?$/i.test(value)) return null;
  const marker = /\s*(?:x|ext\.?|extension)\s*(\d+)$/i.exec(value);
  const main = marker ? value.slice(0, marker.index) : value;
  let digits = main.replace(/\D/g, "");
  let ext = marker ? marker[1] : "";
  if (digits.length === 11 && digits.startsWith("1")) digits = digits.slice(1);
  else if (!ext && digits.length > 10 && digits.length <= 14 && /\D/.test(main)) {
    ext = digits.slice(10);
    digits = digits.slice(0, 10);
  }
  if (digits.length !== 10 && digits.length !== 7) return null;
  if (/^[01]/.test(digits)) return null;
  return { digits, ext };
}

export function parsePhone(text) {
  return parsePhoneParts(text)?.digits ?? null;
}

export function formatPhone(parts, pattern) {
  const { digits, ext } = typeof parts === "string" ? { digits: parts, ext: "" } : parts;
  let index = 0;
  const base = digits.length === 7
    ? `${digits.slice(0, 3)}-${digits.slice(3)}`
    : pattern.replace(/#/g, () => digits[index++] ?? "");
  return ext ? `${base} x${ext}` : base;
}

export function validPhonePattern(pattern) {
  return (pattern.match(/#/g) ?? []).length === 10;
}

export function validNumberPattern(pattern) {
  return /[0#]/.test(pattern);
}

export function validDatePattern(pattern) {
  return /yy|mm?|dd?/i.test(pattern);
}

export function formatNumber(number, pattern) {
  const core = /[#0,.]*[#0][#0,.]*/.exec(pattern);
  if (!core) return null;
  const prefix = pattern.slice(0, core.index).replaceAll('"', "");
  const suffix = pattern.slice(core.index + core[0].length).replaceAll('"', "");
  const [whole, fraction = ""] = core[0].split(".");
  const decimals = fraction.replaceAll(",", "").length;
  const minDecimals = (fraction.match(/0/g) ?? []).length;
  const scaled = suffix.includes("%") ? number * 100 : number;
  const body = Math.abs(scaled).toLocaleString("en-US", {
    useGrouping: whole.includes(","),
    minimumFractionDigits: minDecimals,
    maximumFractionDigits: decimals,
  });
  const sign = scaled < 0 && Number(body.replaceAll(",", "")) !== 0 ? "-" : "";
  return `${sign}${prefix}${body}${suffix}`;
}

export function formatDate({ year, month, day }, pattern) {
  return pattern.replace(/yyyy|yy|mmmm|mmm|mm|m|dd|d/gi, (token) => {
    switch (token.toLowerCase()) {
      case "yyyy": return String(year).padStart(4, "0");
      case "yy": return String(year % 100).padStart(2, "0");
      case "mmmm": return MONTHS[month - 1];
      case "mmm": return MONTHS[month - 1].slice(0, 3);
      case "mm": return String(month).padStart(2, "0");
      case "m": return String(month);
      case "dd": return String(day).padStart(2, "0");
      default: return String(day);
    }
  });
}

const EDGE = /^[\s\u200b-\u200d\u2060\ufeff]+|[\s\u200b-\u200d\u2060\ufeff]+$/g;

export function applyCaps(text, caps) {
  if (caps === "upper") return text.toLocaleUpperCase();
  if (caps === "lower") return text.toLocaleLowerCase();
  if (caps === "title") return text.toLocaleLowerCase().replace(/(^|[\s_-])(\p{L})/gu, (_, gap, letter) => gap + letter.toLocaleUpperCase());
  return text;
}

export function formatValue(raw, spec) {
  let text = String(raw ?? "");
  if (!spec) return text;
  if (spec.trim) text = text.replace(EDGE, "");
  const converted = convertValue(text, spec);
  return spec.caps && spec.caps !== "keep" && converted === text ? applyCaps(text, spec.caps) : converted;
}

function convertValue(text, spec) {
  if (spec.bool) {
    const flag = parseBool(text);
    if (flag !== null) return formatBool(flag, spec.bool);
  }
  if (spec.date) {
    const date = parseDate(text);
    if (date) return formatDate(date, spec.date);
  }
  if (spec.phone) {
    const parts = parsePhoneParts(text);
    if (parts) return formatPhone(parts, spec.phone);
  }
  if (spec.number) {
    const number = parseNumber(text);
    if (number !== null) return formatNumber(number, spec.number) ?? text;
  }
  return text;
}

export function resolveSpec(formats, header, rowIndex) {
  if (!formats) return null;
  const cell = formats.cells?.[`${rowIndex}|${header}`];
  // A cell can opt out of its column's formatting.
  if (cell?.skip) return null;
  const layers = [formats.dataset, formats.columns?.[header], cell];
  const spec = { trim: true };
  for (const layer of layers) {
    for (const key of ["number", "date", "phone", "bool"]) if (layer?.[key]) spec[key] = layer[key];
    if (layer && "trim" in layer) spec.trim = Boolean(layer.trim);
    if (layer?.caps && layer.caps !== "keep") spec.caps = layer.caps;
  }
  return spec;
}

export function emptyFormats() {
  return { dataset: {}, columns: {}, cells: {} };
}

const EMAIL = /^[^\s@,;]+@[^\s@,;]+\.[A-Za-z]{2,}$/;
const BOOLEANS = new Set(["true", "false", "yes", "no", "y", "n", "t", "f"]);
const US_ADDRESS = /^.+,\s*.+,\s*[A-Za-z]{2}\s+\d{5}(?:-\d{4})?$/;

// Phone numbers are only recognised in columns whose name suggests a phone number.
export function classifyValue(raw, { phoneColumn = false } = {}) {
  const text = String(raw ?? "").trim();
  if (!text) return "empty";
  if (EMAIL.test(text)) return "email";
  if (BOOLEANS.has(text.toLowerCase())) return "bool";
  if (US_ADDRESS.test(text)) return "address";
  if (/^\$|^\(?-?\$/.test(text) && parseNumber(text) !== null) return "currency";
  if (/%$/.test(text) && parseNumber(text) !== null) return "percentage";
  if (parseDate(text)) return "date";
  if (phoneColumn && parsePhoneParts(text)) return "phone";
  if (parseNumber(text) !== null) return "number";
  return "text";
}

export const TYPE_LABELS = {
  email: "Email", bool: "Yes/No", address: "Address", currency: "Currency", percentage: "Percentage",
  date: "Date", phone: "Phone", number: "Number", text: "Text",
};

// Mixed columns report every recognised type; the first is the dominant one.
export function detectColumnType(values, { name } = {}) {
  const phoneColumn = isPhoneName(name);
  const counts = {};
  let filled = 0;
  for (const value of values) {
    const kind = classifyValue(value, { phoneColumn });
    if (kind === "empty") continue;
    filled += 1;
    counts[kind] = (counts[kind] ?? 0) + 1;
  }
  // Plain numbers and currency are the same kind of value written differently.
  if (counts.currency && counts.number) {
    counts.currency += counts.number;
    delete counts.number;
  }
  const types = Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([type, count]) => ({ type, count }));
  const top = types[0];
  let type = top && top.count / filled >= 0.6 ? top.type : filled ? "mixed" : "empty";
  if (type === "mixed" || type === "text") {
    const addresses = values.filter((value) => String(value ?? "").trim() && /\d/.test(String(value)) && /,/.test(String(value))).length;
    if (filled && addresses / filled >= 0.6) type = "address";
  }
  const incompleteAddresses = [];
  if (type === "address") {
    values.forEach((value, index) => {
      const text = String(value ?? "").trim();
      if (text && !US_ADDRESS.test(text)) incompleteAddresses.push({ row: index + 1, value: text });
    });
  }
  return { type, filled, types, incompleteAddresses };
}

const BLANKISH = /[\s\u00a0\u200b-\u200d\ufeff]/g;

export function columnInsights(values, { large = false, name } = {}) {
  const list = values.map((value) => String(value ?? ""));
  const filledValues = list.filter((value) => value.replace(BLANKISH, ""));
  const empty = list.length - filledValues.length;
  const detected = detectColumnType(list, { name });
  const counts = new Map();
  filledValues.forEach((value) => counts.set(value.trim(), (counts.get(value.trim()) ?? 0) + 1));
  const sorted = [...counts].sort((a, b) => b[1] - a[1]);
  const dateColumn = detected.type === "date";
  const listAll = dateColumn && sorted.length <= 5;
  const repeating = sorted.length > 0 && (detected.type === "bool" || sorted.length <= Math.max(10, filledValues.length / 3)) && sorted[0][1] > 1;
  const limit = large ? 10 : 5;
  const show = listAll || (repeating && !dateColumn) ? sorted : [];
  const top = show.slice(0, limit).map(([value, count]) => ({ value, count, percent: Math.round((count / list.length) * 100) }));
  const other = show.slice(limit).reduce((total, [, count]) => total + count, 0);
  let numeric = null;
  if (["number", "currency", "percentage"].includes(detected.type)) {
    const numbers = filledValues.map(parseNumber).filter((number) => number !== null);
    if (numbers.length) {
      const mean = numbers.reduce((a, b) => a + b, 0) / numbers.length;
      const variance = numbers.reduce((a, b) => a + (b - mean) ** 2, 0) / numbers.length;
      numeric = { min: Math.min(...numbers), max: Math.max(...numbers), mean, stdDev: Math.sqrt(variance) };
    }
  }
  let dates = null;
  if (dateColumn) {
    const times = filledValues.map(parseDate).filter(Boolean).map((d) => Date.UTC(d.year, d.month - 1, d.day));
    if (times.length) {
      const iso = (ms) => new Date(ms).toISOString().slice(0, 10);
      const min = Math.min(...times);
      const max = Math.max(...times);
      const binCount = max === min ? 1 : Math.min(16, new Set(times).size);
      const span = max - min || 1;
      const bins = Array.from({ length: binCount }, (_, i) => ({ from: iso(min + (span * i) / binCount), to: iso(min + (span * (i + 1)) / binCount), count: 0 }));
      times.forEach((time) => { bins[Math.min(binCount - 1, Math.floor(((time - min) / span) * binCount))].count += 1; });
      dates = { earliest: iso(min), average: iso(times.reduce((a, b) => a + b, 0) / times.length), latest: iso(max), bins, times, count: times.length };
    }
  }
  return { ...detected, rows: list.length, empty, unique: counts.size, top, other, numeric, dates };
}

export function conversionReport(values, kind) {
  const parse = { date: parseDate, number: parseNumber, phone: parsePhone }[kind];
  let ok = 0;
  let failedCount = 0;
  const failed = [];
  for (const value of values) {
    const text = String(value ?? "");
    if (!text.trim()) continue;
    const result = parse(text);
    if (result !== null && result !== undefined) ok += 1;
    else {
      failedCount += 1;
      if (!failed.includes(text)) failed.push(text);
    }
  }
  return { ok, failedCount, failed };
}
