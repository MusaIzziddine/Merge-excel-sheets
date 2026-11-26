import * as XLSX from 'xlsx';

const AR_DIGITS = '٠١٢٣٤٥٦٧٨٩';
const toLatinDigits = (s) => s.replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)));

/** Normalizes a header so "اسم  القطعة" and "اسم القطعة" become the same column. */
export const normHeader = (h) =>
  toLatinDigits(String(h ?? '')).replace(/ـ/g, '').replace(/\s+/g, ' ').trim().toLowerCase();

/** Normalizes a part number so " p-1001" and "P-1001" match. */
export const normKey = (v) => toLatinDigits(String(v ?? '')).replace(/\s+/g, '').toUpperCase();

const STRONG_KEY = /رقم\s*القطع|part\s*(no|num|#)/i;
const WEAK_KEY = /رقم\s*الصنف|كود|رمز|^code$|^sku$|item\s*(no|code)|رقم/i;
export const KEY_ID = '__key';

let nextId = 0;

/** Wraps a parsed workbook into a "source file" the merger works with. */
export function makeSource(name, wb) {
  const src = { id: ++nextId, name, wb, sheet: wb.SheetNames[0], hdr: 1, hdrSet: false, keySet: false, headers: [], data: [], key: 0 };
  return loadSheet(src);
}

export function readFile(name, arrayBuffer) {
  return makeSource(name, XLSX.read(arrayBuffer, { type: 'array' }));
}

/** (Re)reads the selected sheet: finds the header row, the rows and the key column. */
export function loadSheet(src) {
  const s = { ...src };
  const aoa = XLSX.utils.sheet_to_json(s.wb.Sheets[s.sheet], { header: 1, defval: '', raw: true, blankrows: false });
  if (!s.hdrSet) {
    const i = aoa.findIndex((r) => r.filter((c) => String(c).trim() !== '').length >= 2);
    s.hdr = (i < 0 ? 0 : i) + 1;
  }
  const hi = Math.min(Math.max(s.hdr - 1, 0), Math.max(aoa.length - 1, 0));
  s.headers = (aoa[hi] || []).map((h, i) => String(h).trim() || `عمود ${i + 1}`);
  s.data = aoa.slice(hi + 1).filter((r) => r.some((c) => String(c).trim() !== ''));
  if (!s.keySet) {
    let k = s.headers.findIndex((h) => STRONG_KEY.test(h));
    if (k < 0) k = s.headers.findIndex((h) => WEAK_KEY.test(h));
    s.key = k < 0 ? 0 : k;
  }
  if (s.key >= s.headers.length) s.key = 0;
  s.filters = [];
  return s;
}

/** Text form of a cell used for filtering; empty cells become ''. */
export const cellText = (v) => String(v ?? '').trim();

/**
 * Filter shapes:
 *   { col, op: 'in',          values: ['a', 'b'] }  keep rows whose cell is one of the values
 *   { col, op: 'notIn',       values: [...] }       drop rows whose cell is one of the values
 *   { col, op: 'contains',    text: '...' }         keep rows whose cell contains the text
 *   { col, op: 'notContains', text: '...' }
 *   { col, op: 'gte' | 'lte', text: '10' }          numeric comparison
 * Filters on one file are combined with AND. An incomplete filter (no values / no text) is ignored.
 */
export function rowPasses(row, f) {
  const t = cellText(row[f.col]);
  switch (f.op) {
    case 'in':
      return !f.values?.length || f.values.includes(t);
    case 'notIn':
      return !f.values?.length || !f.values.includes(t);
    case 'contains':
    case 'notContains': {
      const q = cellText(f.text).toLowerCase();
      if (!q) return true;
      const hit = t.toLowerCase().includes(q);
      return f.op === 'contains' ? hit : !hit;
    }
    case 'gte':
    case 'lte': {
      const lim = parseFloat(f.text);
      if (!Number.isFinite(lim)) return true;
      const n = parseFloat(t);
      if (!Number.isFinite(n)) return false;
      return f.op === 'gte' ? n >= lim : n <= lim;
    }
    default:
      return true;
  }
}

export function filteredRows(src) {
  const fs = src.filters || [];
  return fs.length ? src.data.filter((r) => fs.every((f) => rowPasses(r, f))) : src.data;
}

/** Distinct values of a column with counts, most frequent first. */
export function distinctValues(src, col) {
  const counts = new Map();
  for (const r of src.data) {
    const t = cellText(r[col]);
    counts.set(t, (counts.get(t) || 0) + 1);
  }
  return [...counts.entries()]
    .map(([value, count]) => ({ value, count }))
    .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value, 'ar', { numeric: true }));
}

/** Output columns: key first, then every distinct header; keeps the user's order/labels/toggles. */
export function buildColumns(sources, prev = []) {
  const found = [];
  const seen = new Set([KEY_ID]);
  const first = sources[0];
  found.push({ id: KEY_ID, label: first ? first.headers[first.key] : 'رقم القطعة', on: true });
  for (const s of sources) {
    s.headers.forEach((h, i) => {
      if (i === s.key) return;
      const id = normHeader(h);
      if (!id || seen.has(id)) return;
      seen.add(id);
      found.push({ id, label: h, on: true });
    });
  }
  const prevById = new Map(prev.map((c, i) => [c.id, { ...c, i }]));
  const kept = found
    .filter((c) => prevById.has(c.id))
    .map((c) => {
      const p = prevById.get(c.id);
      return { ...c, label: p.edited ? p.label : c.label, edited: p.edited, on: p.on, i: p.i };
    })
    .sort((a, b) => a.i - b.i)
    .map(({ i, ...c }) => c);
  return [...kept, ...found.filter((c) => !prevById.has(c.id))];
}

/**
 * Joins all sources on the key column.
 * opts.mode: 'all' | 'base' (keys in first file) | 'both' (keys in every file)
 * opts.sumDuplicates: add numeric values when a key repeats inside the same file
 * opts.total: { qtyId, priceId } to append a quantity × price column
 */
export function mergeSources(sources, columns, opts = {}) {
  const { mode = 'all', sumDuplicates = false, total = null } = opts;
  const map = new Map();
  let duplicates = 0;

  sources.forEach((s, fi) => {
    const seen = new Set();
    for (const row of filteredRows(s)) {
      const raw = row[s.key];
      const k = normKey(raw);
      if (!k) continue;
      let rec = map.get(k);
      if (!rec) {
        rec = { key: typeof raw === 'string' ? raw.trim() : raw, values: {}, from: new Set() };
        map.set(k, rec);
      }
      const dup = seen.has(k);
      if (dup) duplicates++;
      seen.add(k);
      rec.from.add(fi);
      s.headers.forEach((h, i) => {
        if (i === s.key) return;
        const id = normHeader(h);
        const v = row[i];
        if (v === '' || v == null) return;
        if (dup) {
          if (sumDuplicates && typeof v === 'number' && typeof rec.values[id] === 'number') rec.values[id] += v;
          return;
        }
        if (rec.values[id] == null || rec.values[id] === '') rec.values[id] = v;
      });
    }
  });

  let records = [...map.values()];
  if (mode === 'base') records = records.filter((r) => r.from.has(0));
  if (mode === 'both') records = records.filter((r) => r.from.size === sources.length);

  const visible = columns.filter((c) => c.on);
  const headers = visible.map((c) => c.label);
  const withTotal = Boolean(total && total.qtyId && total.priceId);
  if (withTotal) headers.push('الإجمالي');

  const rows = records.map((r) => {
    const values = visible.map((c) => (c.id === KEY_ID ? r.key : r.values[c.id] ?? ''));
    if (withTotal) {
      const q = parseFloat(r.values[total.qtyId]);
      const p = parseFloat(r.values[total.priceId]);
      values.push(Number.isFinite(q) && Number.isFinite(p) ? q * p : '');
    }
    return { values, missing: values.some((v) => v === '') };
  });

  return {
    headers,
    rows,
    duplicates,
    total: withTotal ? { qtyIndex: visible.findIndex((c) => c.id === total.qtyId), priceIndex: visible.findIndex((c) => c.id === total.priceId) } : null,
  };
}

const colLetter = (n) => XLSX.utils.encode_col(n);

/** Builds the output .xlsx (right-to-left sheet, filters, widths, total as a formula). */
export function buildWorkbook(result) {
  const { headers, rows, total } = result;
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows.map((r) => r.values)]);
  if (total && total.qtyIndex >= 0 && total.priceIndex >= 0) {
    const t = headers.length - 1;
    rows.forEach((_, i) => {
      const r = i + 2;
      const cell = ws[`${colLetter(t)}${r}`];
      if (cell && cell.t === 'n') cell.f = `${colLetter(total.qtyIndex)}${r}*${colLetter(total.priceIndex)}${r}`;
    });
  }
  ws['!cols'] = headers.map((h, i) => ({
    wch: Math.min(45, Math.max(String(h).length, ...rows.slice(0, 500).map((r) => String(r.values[i] ?? '').length)) + 4),
  }));
  if (headers.length) ws['!autofilter'] = { ref: `A1:${colLetter(headers.length - 1)}${rows.length + 1}` };
  const wb = XLSX.utils.book_new();
  wb.Workbook = { Views: [{ RTL: true }] };
  XLSX.utils.book_append_sheet(wb, ws, 'الدمج');
  return wb;
}

export function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename);
}

/** Two sample files matching the example: prices + withdrawals. */
export function sampleSources() {
  const mk = (name, aoa) => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Sheet1');
    return makeSource(name, wb);
  };
  return [
    mk('الأسعار.xlsx', [
      ['رقم القطعة', 'اسم القطعة', 'سعر القطعة'],
      ['P-1001', 'فلتر زيت', 35],
      ['P-1002', 'فلتر هواء', 48],
      ['P-1003', 'بواجي', 22.5],
      ['P-1004', 'سير مكينة', 120],
      ['P-1005', 'فحمات فرامل', 165],
    ]),
    mk('السحوبات.xlsx', [
      ['اسم القطعة', 'رقم القطعة', 'الكمية المسحوبة', 'المستودع'],
      ['فلتر زيت', 'P-1001', 12, 'الرئيسي'],
      ['بواجي', 'P-1003', 40, 'الرئيسي'],
      ['سير مكينة', 'P-1004', 3, 'الفرعي'],
      ['فحمات فرامل', 'P-1005', 8, 'الرئيسي'],
      ['مساحات', 'P-1009', 6, 'الفرعي'],
    ]),
  ];
}
