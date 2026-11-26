import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import {
  KEY_ID, autoKeys, buildColumns, buildWorkbook, detectSuffixes, distinctValues, makeSource, mergeSources, normKey, sampleSources,
} from './merge.js';

const src = (name, aoa) => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'S');
  return makeSource(name, wb);
};

describe('mergeSources', () => {
  it('joins prices and withdrawals on part number, whatever the column order', () => {
    const sources = sampleSources();
    expect(sources[0].key).toBe(0);
    expect(sources[1].key).toBe(1);
    const order = [KEY_ID, 'اسم القطعة', 'الكمية المسحوبة', 'سعر القطعة'];
    const cols = buildColumns(sources)
      .filter((c) => order.includes(c.id))
      .sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
    const r = mergeSources(sources, cols);
    expect(r.headers).toEqual(['رقم القطعة', 'اسم القطعة', 'الكمية المسحوبة', 'سعر القطعة']);
    expect(r.rows.find((x) => x.values[0] === 'P-1001').values).toEqual(['P-1001', 'فلتر زيت', 12, 35]);
    expect(r.rows.find((x) => x.values[0] === 'P-1002').missing).toBe(true);
    expect(r.rows).toHaveLength(6);
  });

  it('filters by mode', () => {
    const s = sampleSources();
    const cols = buildColumns(s);
    expect(mergeSources(s, cols, { mode: 'base' }).rows).toHaveLength(5);
    expect(mergeSources(s, cols, { mode: 'both' }).rows).toHaveLength(4);
  });

  it('matches keys across number/text, spaces, case and Arabic digits', () => {
    const a = src('a', [['رقم القطعة', 'السعر'], [1001, 5], ['ab-1', 7]]);
    const b = src('b', [['رقم القطعة', 'الكمية'], ['١٠٠١', 2], [' AB-1 ', 3]]);
    const r = mergeSources([a, b], buildColumns([a, b]), { mode: 'both' });
    expect(r.rows.map((x) => x.values)).toEqual([[1001, 5, 2], ['ab-1', 7, 3]]);
    expect(r.perFile.map((f) => f.matched)).toEqual([2, 2]);
  });

  it('sums duplicate withdrawals only when asked, and adds a total', () => {
    const a = src('a', [['رقم القطعة', 'السعر'], ['X', 10]]);
    const b = src('b', [['رقم القطعة', 'الكمية'], ['X', 2], ['X', 3]]);
    const cols = buildColumns([a, b]);
    expect(mergeSources([a, b], cols).rows[0].values).toEqual(['X', 10, 2]);
    const r = mergeSources([a, { ...b, sumDuplicates: true }], cols, { total: { qtyId: 'الكمية', priceId: 'السعر' } });
    expect(r.duplicates).toBe(1);
    expect(r.rows[0].values).toEqual(['X', 10, 5, 50]);
    const ws = buildWorkbook(r).Sheets['الدمج'];
    expect(ws.D2.f).toBe('C2*B2');
  });

  it('skips title rows above the header', () => {
    const a = src('a', [['تقرير المخزن'], ['رقم القطعة', 'الاسم'], ['Z', 'برغي']]);
    expect(a.hdr).toBe(2);
    expect(a.data).toEqual([['Z', 'برغي']]);
  });
});

describe('filters', () => {
  const qty = () => src('q', [
    ['رقم القطعة', 'الكمية', 'المستودع', 'الحركة'],
    ['A', 5, 'الرئيسي', 'سحب'],
    ['A', 7, 'الفرعي', 'سحب'],
    ['B', 2, 'الرئيسي', 'إرجاع'],
    ['C', 9, 'الرئيسي', 'سحب مستعجل'],
  ]);
  const prices = () => src('p', [['رقم القطعة', 'السعر'], ['A', 1], ['B', 2], ['C', 3]]);
  const run = (filters, opts = {}) => {
    const p = prices();
    const q = { ...qty(), filters, sumDuplicates: Boolean(opts.sumDuplicates) };
    return mergeSources([p, q], buildColumns([p, q]), { mode: 'both', ...opts }).rows.map((r) => r.values);
  };

  it('keeps only chosen values, before matching and summing', () => {
    expect(run([{ col: 2, op: 'in', values: ['الفرعي'] }], { sumDuplicates: true })).toEqual([['A', 1, 7, 'الفرعي', 'سحب']]);
    expect(run([{ col: 2, op: 'in', values: ['الرئيسي'] }], { sumDuplicates: true }).map((r) => r[2])).toEqual([5, 2, 9]);
  });

  it('supports exclude, contains, numeric and combined (AND) filters', () => {
    expect(run([{ col: 3, op: 'notIn', values: ['إرجاع'] }]).map((r) => r[0])).toEqual(['A', 'C']);
    expect(run([{ col: 3, op: 'contains', text: 'سحب' }]).map((r) => r[0])).toEqual(['A', 'C']);
    expect(run([{ col: 3, op: 'notContains', text: 'سحب' }]).map((r) => r[0])).toEqual(['B']);
    expect(run([{ col: 1, op: 'gte', text: '6' }]).map((r) => r[0])).toEqual(['A', 'C']);
    expect(run([{ col: 2, op: 'notEmpty' }])).toHaveLength(3);
    expect(run([
      { col: 2, op: 'in', values: ['الرئيسي'] },
      { col: 3, op: 'contains', text: 'سحب' },
    ]).map((r) => r[2])).toEqual([5, 9]);
  });

  it('ignores an empty filter and lists distinct values', () => {
    expect(run([{ col: 2, op: 'in', values: [] }])).toHaveLength(3);
    expect(distinctValues(qty(), 2)).toEqual([{ value: 'الرئيسي', count: 3 }, { value: 'الفرعي', count: 1 }]);
  });
});

describe('real-world part numbers', () => {
  // Shapes taken from a prices file vs a delivery file: two "OEM NO" columns, GST suffix, dashes.
  const prices = () => src('prices', [
    ['OEM NO', 'OEM NO', 'ARABIC ', 'QTY', 'كلفة', '', ''],
    ['96536712GST', '96536712', 'خرطوم', 20, 5.5, '', ''],
    ['42200-31A-H01', '', 'هب عجل', 6, 31, '', '   '],
    ['42200-31A-H01', '', 'هب عجل', 6, 33, '', ''],
    ['94580175GST', '94580175', 'حساس', 5, 9, '', ''],
  ]);
  const qty = () => src('qty', [
    ['OEM NO', 'OEM NO', 'ARABIC ', 'QTY', 'القويسمه'],
    [96536712, '96536712', 'خرطوم', 20, 2],
    ['4220031AH01', '', 'هب عجل', 6, ''],
    ['94580175GM', '94580175', 'حساس', 4, 3],
  ]);

  it('drops empty columns and makes duplicate headers unique', () => {
    expect(prices().headers).toEqual(['OEM NO', 'OEM NO (2)', 'ARABIC', 'QTY', 'كلفة']);
  });

  it('normalizes symbols and suffixes', () => {
    const o = { ignoreSymbols: true, suffixes: ['GST', 'GM'] };
    expect(normKey('42200-31A-H01', o)).toBe(normKey('4220031AH01', o));
    expect(normKey('96536712GST', o)).toBe('96536712');
    expect(normKey('94580175GM', o)).toBe('94580175');
    expect(normKey('GST', o)).toBe('GST');
  });

  it('detects suffixes, merges, and takes a shared column from the chosen file', () => {
    let s = autoKeys([prices(), qty()]);
    const opts = { ignoreSymbols: true, suffixes: detectSuffixes(s) };
    expect(opts.suffixes).toEqual(['GST', 'GM']);
    s = autoKeys(s, opts);
    const cols = buildColumns(s).map((c) => (c.id === 'qty' ? { ...c, source: s[1].id } : c));
    const r = mergeSources(s, cols, { mode: 'both', keyOptions: opts });
    const q = r.headers.indexOf('QTY');
    expect(r.rows.map((x) => [x.values[0], x.values[q]])).toEqual([['96536712GST', 20], ['42200-31A-H01', 6], ['94580175GST', 4]]);
    expect(r.perFile[0].conflicts).toBe(1);
    expect(r.perFile[0].matched).toBe(3);
  });
});

