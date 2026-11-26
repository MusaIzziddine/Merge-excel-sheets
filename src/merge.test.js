import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { KEY_ID, buildColumns, buildWorkbook, makeSource, mergeSources, sampleSources } from './merge.js';

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
    const cols = buildColumns(sources).sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
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
  });

  it('sums duplicate withdrawals only when asked, and adds a total', () => {
    const a = src('a', [['رقم القطعة', 'السعر'], ['X', 10]]);
    const b = src('b', [['رقم القطعة', 'الكمية'], ['X', 2], ['X', 3]]);
    const cols = buildColumns([a, b]);
    expect(mergeSources([a, b], cols).rows[0].values).toEqual(['X', 10, 2]);
    const r = mergeSources([a, b], cols, { sumDuplicates: true, total: { qtyId: 'الكمية', priceId: 'السعر' } });
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
