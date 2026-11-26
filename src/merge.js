import * as XLSX from 'xlsx';

const AR_DIGITS = '٠١٢٣٤٥٦٧٨٩';
const toLatinDigits = (s) => s.replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)));

/** Normalizes a header so "اسم  القطعة" and "اسم القطعة" become the same column. */
export const normHeader = (h) =>
  toLatinDigits(String(h ?? '')).replace(/ـ/g, '').replace(/\s+/g, ' ').trim().toLowerCase();

export const DEFAULT_KEY_OPTIONS = { ignoreSymbols: true, suffixes: [] };

/**
 * Normalizes a part number for matching.
 * Always: Arabic digits → Latin, spaces removed, upper case.
 * ignoreSymbols: also drops - / . _ etc. ("42200-31A-H01" = "4220031AH01").
 * suffixes: endings to drop ("96536712GST" = "96536712" with suffixes ["GST"]).
 */
export function normKey(v, opts = DEFAULT_KEY_OPTIONS) {
  let k = toLatinDigits(String(v ?? '')).replace(/\s+/g, '').toUpperCase();
  if (opts.ignoreSymbols) k = k.replace(/[^0-9A-Z؀-ۿ]/g, '');
  for (const suf of opts.suffixes || []) {
    const s = normKey(suf, { ...opts, suffixes: [] });
    if (s && k.length > s.length && k.endsWith(s)) {
      k = k.slice(0, -s.length);
      break;
    }
  }
  return k;
}

const STRONG_KEY = /رقم\s*القطع|part\s*(no|num|#)/i;
const WEAK_KEY = /رقم\s*الصنف|كود|رمز|^code$|^sku$|item\s*(no|code)|oem|p\/n|رقم/i;
export const KEY_ID = '__key';

let nextId = 0;

/** Wraps a parsed workbook into a "source file" the merger works with. */
export function makeSource(name, wb) {
  const src = {
    id: ++nextId, name, wb, sheet: wb.SheetNames[0], hdr: 1, hdrSet: false, keySet: false,
    headers: [], data: [], key: 0, filters: [], sumDuplicates: false,
  };
  return loadSheet(src);
}

export function readFile(name, arrayBuffer) {
  return makeSource(name, XLSX.read(arrayBuffer, { type: 'array' }));
}

const blank = (c) => String(c ?? '').trim() === '';

/** (Re)reads the selected sheet: header row, rows, unique header names, guessed key column. */
export function loadSheet(src) {
  const s = { ...src };
  const aoa = XLSX.utils.sheet_to_json(s.wb.Sheets[s.sheet], { header: 1, defval: '', raw: true, blankrows: false });
  if (!s.hdrSet) {
    const i = aoa.findIndex((r) => r.filter((c) => !blank(c)).length >= 2);
    s.hdr = (i < 0 ? 0 : i) + 1;
  }
  const hi = Math.min(Math.max(s.hdr - 1, 0), Math.max(aoa.length - 1, 0));
  const headRow = aoa[hi] || [];
  const body = aoa.slice(hi + 1).filter((r) => r.some((c) => !blank(c)));
  const width = Math.max(headRow.length, ...body.map((r) => r.length), 0);

  // Drop columns that have neither a header nor any data (common trailing junk in Excel files).
  const keep = [];
  for (let i = 0; i < width; i++) if (!blank(headRow[i]) || body.some((r) => !blank(r[i]))) keep.push(i);

  // Make header names unique: "OEM NO", "OEM NO (2)".
  const seen = new Map();
  s.headers = keep.map((i) => {
    const base = blank(headRow[i]) ? `عمود ${i + 1}` : String(headRow[i]).trim();
    const n = (seen.get(normHeader(base)) || 0) + 1;
    seen.set(normHeader(base), n);
    return n === 1 ? base : `${base} (${n})`;
  });
  s.data = body.map((r) => keep.map((i) => (r[i] === undefined ? '' : r[i])));

  if (!s.keySet) {
    let k = s.headers.findIndex((h) => STRONG_KEY.test(h));
    if (k < 0) k = s.headers.findIndex((h) => WEAK_KEY.test(h));
    s.key = k < 0 ? 0 : k;
  }
  if (s.key >= s.headers.length) s.key = 0;
  s.filters = [];
  return s;
}

const keySet = (src, col, opts) => {
  const set = new Set();
  for (const r of src.data) {
    const k = normKey(r[col], opts);
    if (k) set.add(k);
  }
  return set;
};

/**
 * For every file after the first whose key the user has not chosen, picks the column whose
 * values match the first file's key column best (falls back to the header guess).
 */
export function autoKeys(sources, opts = DEFAULT_KEY_OPTIONS) {
  if (sources.length < 2) return sources;
  const base = keySet(sources[0], sources[0].key, opts);
  return sources.map((s, i) => {
    if (i === 0 || s.keySet) return s;
    let best = s.key;
    let bestHits = -1;
    s.headers.forEach((_, c) => {
      let hits = 0;
      for (const k of keySet(s, c, opts)) if (base.has(k)) hits++;
      if (hits > bestHits) [best, bestHits] = [c, hits];
    });
    return bestHits > 0 ? { ...s, key: best } : s;
  });
}

const matchCount = (a, b) => {
  let n = 0;
  for (const k of a) if (b.has(k)) n++;
  return n;
};

/**
 * Finds letter endings (like GST, GM) that some files add to part numbers and others don't.
 * A suffix is suggested only when dropping it makes more keys match between the files.
 */
export function detectSuffixes(sources, opts = DEFAULT_KEY_OPTIONS) {
  if (sources.length < 2) return [];
  const plain = { ...opts, suffixes: [] };
  const counts = new Map();
  for (const s of sources) {
    for (const k of keySet(s, s.key, plain)) {
      const m = k.match(/\d([A-Z]{2,4})$/);
      if (m) counts.set(m[1], (counts.get(m[1]) || 0) + 1);
    }
  }
  const total = (o) => {
    const sets = sources.map((s) => keySet(s, s.key, o));
    let n = 0;
    for (let i = 1; i < sets.length; i++) n += matchCount(sets[0], sets[i]);
    return n;
  };
  const chosen = [];
  let score = total(plain);
  const candidates = [...counts.entries()].filter(([, n]) => n >= 1).sort((a, b) => b[1] - a[1]).map(([s]) => s);
  for (const suf of candidates.slice(0, 8)) {
    const next = total({ ...plain, suffixes: [...chosen, suf] });
    if (next > score) {
      chosen.push(suf);
      score = next;
    }
  }
  return chosen;
}

/** Text form of a cell used for filtering; empty cells become ''. */
export const cellText = (v) => String(v ?? '').trim();

/**
 * Filter shapes:
 *   { col, op: 'in',          values: ['a', 'b'] }  keep rows whose cell is one of the values
 *   { col, op: 'notIn',       values: [...] }       drop rows whose cell is one of the values
 *   { col, op: 'notEmpty' } / { col, op: 'empty' }  cell has a value / is empty
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
    case 'notEmpty':
      return t !== '';
    case 'empty':
      return t === '';
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

/**
 * Output columns: key first, then every distinct header across files.
 * A column whose name appears in several files lists them in `files` (source ids);
 * `source` picks which file it is taken from ('' = first file that has a value).
 * Keeps the user's order, labels, toggles and sources from `prev`.
 */
export function buildColumns(sources, prev = []) {
  const found = [];
  const byId = new Map();
  const first = sources[0];
  const keyCol = { id: KEY_ID, label: first ? first.headers[first.key] : 'رقم القطعة', on: true, files: [], source: '' };
  found.push(keyCol);
  byId.set(KEY_ID, keyCol);
  for (const s of sources) {
    s.headers.forEach((h, i) => {
      if (i === s.key) return;
      const id = normHeader(h);
      if (!id) return;
      if (byId.has(id)) {
        byId.get(id).files.push(s.id);
        return;
      }
      const c = { id, label: h, on: true, files: [s.id], source: '' };
      byId.set(id, c);
      found.push(c);
    });
  }
  const prevById = new Map(prev.map((c, i) => [c.id, { ...c, i }]));
  const kept = found
    .filter((c) => prevById.has(c.id))
    .map((c) => {
      const p = prevById.get(c.id);
      return {
        ...c,
        label: p.edited ? p.label : c.label,
        edited: p.edited,
        on: p.on,
        source: c.files.includes(p.source) ? p.source : '',
        i: p.i,
      };
    })
    .sort((a, b) => a.i - b.i)
    .map(({ i, ...c }) => c);
  return [...kept, ...found.filter((c) => !prevById.has(c.id))];
}

/**
 * Joins all sources on their key column.
 * opts.mode: 'all' | 'base' (keys in first file) | 'both' (keys in every file)
 * opts.keyOptions: see normKey
 * opts.total: { qtyId, priceId } to append a quantity × price column
 * opts.rowsFrom: a source id → keep every row of that file (no collapsing of repeated part numbers);
 *   otherwise one row per part number.
 * Per source: src.filters, src.sumDuplicates (add numbers when a key repeats inside that file).
 */
export function mergeSources(sources, columns, opts = {}) {
  const { mode = 'all', total = null, keyOptions = DEFAULT_KEY_OPTIONS, rowsFrom = null } = opts;
  const map = new Map();
  const perFile = sources.map((s) => ({ id: s.id, name: s.name, rows: 0, keys: 0, duplicates: 0, conflicts: 0 }));

  sources.forEach((s, fi) => {
    const seen = new Set();
    const conflicted = new Set();
    const stat = perFile[fi];
    for (const row of filteredRows(s)) {
      const raw = row[s.key];
      const k = normKey(raw, keyOptions);
      if (!k) continue;
      stat.rows++;
      let rec = map.get(k);
      if (!rec) {
        rec = { key: typeof raw === 'string' ? raw.trim() : raw, values: {}, from: new Set() };
        map.set(k, rec);
      }
      const vals = (rec.values[s.id] ||= {});
      const dup = seen.has(k);
      if (dup) stat.duplicates++;
      else stat.keys++;
      seen.add(k);
      rec.from.add(fi);
      s.headers.forEach((h, i) => {
        if (i === s.key) return;
        const id = normHeader(h);
        const v = row[i];
        if (blank(v)) return;
        if (!dup || vals[id] == null) {
          vals[id] = v;
          return;
        }
        if (s.sumDuplicates && typeof v === 'number' && typeof vals[id] === 'number') vals[id] += v;
        else if (!s.sumDuplicates && cellText(v) !== cellText(vals[id]) && !conflicted.has(k)) {
          conflicted.add(k);
          stat.conflicts++;
        }
      });
    }
  });

  let records;
  let order = sources.map((s) => s.id);
  const baseIndex = rowsFrom == null ? -1 : sources.findIndex((s) => s.id === rowsFrom);
  if (baseIndex >= 0) {
    // One output row per (filtered) row of the base file, like VLOOKUP: its own values,
    // plus the other files' values for the same part number.
    const base = sources[baseIndex];
    order = [base.id, ...order.filter((id) => id !== base.id)];
    records = [];
    for (const row of filteredRows(base)) {
      const k = normKey(row[base.key], keyOptions);
      if (!k) continue;
      const rec = map.get(k);
      const own = {};
      base.headers.forEach((h, i) => {
        if (i !== base.key && !blank(row[i])) own[normHeader(h)] = row[i];
      });
      records.push({
        key: typeof row[base.key] === 'string' ? row[base.key].trim() : row[base.key],
        values: { ...rec.values, [base.id]: own },
        from: rec.from,
      });
    }
    if (mode === 'both') records = records.filter((r) => r.from.size === sources.length);
  } else {
    records = [...map.values()];
    if (mode === 'base') records = records.filter((r) => r.from.has(0));
    if (mode === 'both') records = records.filter((r) => r.from.size === sources.length);
  }

  const pick = (rec, c) => {
    if (c.id === KEY_ID) return rec.key;
    if (c.source) return rec.values[c.source]?.[c.id] ?? '';
    for (const id of order) {
      const v = rec.values[id]?.[c.id];
      if (v != null) return v;
    }
    return '';
  };

  const visible = columns.filter((c) => c.on);
  const headers = visible.map((c) => c.label);
  const byId = new Map(columns.map((c) => [c.id, c]));
  const withTotal = Boolean(total && byId.has(total.qtyId) && byId.has(total.priceId));
  if (withTotal) headers.push('الإجمالي');

  const rows = records.map((r) => {
    const values = visible.map((c) => pick(r, c));
    if (withTotal) {
      const q = parseFloat(pick(r, byId.get(total.qtyId)));
      const p = parseFloat(pick(r, byId.get(total.priceId)));
      values.push(Number.isFinite(q) && Number.isFinite(p) ? q * p : '');
    }
    return { values, missing: values.some((v) => v === '') };
  });

  for (const st of perFile) {
    st.matched = 0;
    for (const r of map.values()) if (r.from.size > 1 && r.from.has(perFile.indexOf(st))) st.matched++;
  }

  return {
    headers,
    rows,
    perFile,
    duplicates: perFile.reduce((n, f) => n + f.duplicates, 0),
    total: withTotal
      ? { qtyIndex: visible.findIndex((c) => c.id === total.qtyId), priceIndex: visible.findIndex((c) => c.id === total.priceId) }
      : null,
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
