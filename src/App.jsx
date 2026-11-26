import { useMemo, useRef, useState } from 'react';
import {
  KEY_ID, buildColumns, buildWorkbook, distinctValues, downloadWorkbook, filteredRows, loadSheet, mergeSources, readFile,
  sampleSources,
} from './merge.js';

const QTY_RX = /كمي|عدد|qty|quantity|مسحوب/i;
const PRICE_RX = /سعر|price|cost|تكلف/i;
const PREVIEW_LIMIT = 200;

function initialState() {
  const sources = sampleSources();
  const order = [KEY_ID, 'اسم القطعة', 'الكمية المسحوبة', 'سعر القطعة'];
  const rank = (id) => (order.includes(id) ? order.indexOf(id) : order.length);
  const columns = buildColumns(sources)
    .map((c) => ({ ...c, on: order.includes(c.id) }))
    .sort((a, b) => rank(a.id) - rank(b.id));
  return { sources, columns };
}

export default function App() {
  const [{ sources, columns }, setState] = useState(initialState);
  const [mode, setMode] = useState('all');
  const [sumDuplicates, setSumDuplicates] = useState(false);
  const [withTotal, setWithTotal] = useState(false);
  const [qtyId, setQtyId] = useState('');
  const [priceId, setPriceId] = useState('');
  const [fileName, setFileName] = useState('الملف-المدمج');
  const [errors, setErrors] = useState([]);

  const setSources = (next) =>
    setState((st) => {
      const list = typeof next === 'function' ? next(st.sources) : next;
      return { sources: list, columns: buildColumns(list, st.columns) };
    });
  const setColumns = (fn) => setState((st) => ({ ...st, columns: fn(st.columns) }));

  const valueCols = columns.filter((c) => c.id !== KEY_ID);
  const pick = (current, rx) =>
    valueCols.some((c) => c.id === current) ? current : valueCols.find((c) => rx.test(c.label))?.id ?? valueCols[0]?.id ?? '';
  const qty = pick(qtyId, QTY_RX);
  const price = pick(priceId, PRICE_RX);

  const result = useMemo(
    () =>
      sources.length
        ? mergeSources(sources, columns, { mode, sumDuplicates, total: withTotal ? { qtyId: qty, priceId: price } : null })
        : null,
    [sources, columns, mode, sumDuplicates, withTotal, qty, price],
  );

  async function addFiles(fileList) {
    const errs = [];
    const added = [];
    for (const f of fileList) {
      try {
        added.push(readFile(f.name, await f.arrayBuffer()));
      } catch {
        errs.push(`تعذّرت قراءة الملف: ${f.name}`);
      }
    }
    setErrors(errs);
    if (added.length) setSources((list) => [...list, ...added]);
  }

  function updateSource(id, patch) {
    setSources((list) =>
      list.map((s) => {
        if (s.id !== id) return s;
        const next = { ...s, ...patch };
        // Changing the sheet or header row re-reads the file (and clears its filters).
        return 'sheet' in patch || 'hdr' in patch ? loadSheet(next) : next;
      }),
    );
  }

  function download() {
    const safe = (fileName.trim() || 'الملف-المدمج').replace(/[\\/:*?"<>|]/g, '-');
    downloadWorkbook(buildWorkbook(result), `${safe}.xlsx`);
  }

  const missing = result ? result.rows.filter((r) => r.missing).length : 0;

  return (
    <>
      <header className="band">
        <h1>دمج ملفات الإكسل حسب رقم القطعة</h1>
        <p className="mut">
          ارفع ملفين أو أكثر، اختر عمود الربط في كل ملف، رتّب الأعمدة كما تريد، ثم نزّل ملفاً واحداً جاهزاً. كل المعالجة
          تتم داخل المتصفح، ولا تُرفع ملفاتك لأي مكان.
        </p>
      </header>

      <section className="band">
        <h2>الملفات</h2>
        <DropZone onFiles={addFiles} />
        <div className="row" style={{ marginTop: 10 }}>
          <button onClick={() => { setState(initialState()); setErrors([]); }}>جرّب بمثال جاهز</button>
          <button onClick={() => { setState({ sources: [], columns: [] }); setErrors([]); }}>مسح الكل</button>
        </div>
        <div className="files">
          {sources.map((s, i) => (
            <SourceCard
              key={s.id}
              src={s}
              isFirst={i === 0}
              onChange={(patch) => updateSource(s.id, patch)}
              onRemove={() => setSources((list) => list.filter((x) => x.id !== s.id))}
              onMakeFirst={() => setSources((list) => [s, ...list.filter((x) => x.id !== s.id)])}
            />
          ))}
        </div>
        {errors.map((e) => <div key={e} className="msg bad">{e}</div>)}
      </section>

      <section className="band">
        <div className="opts">
          <div>
            <h2>أعمدة الملف الناتج</h2>
            <p className="mut small">
              فعّل الأعمدة التي تريدها، غيّر اسمها إن أردت، ورتّبها بالأسهم. الأعمدة التي تحمل نفس الاسم في أكثر من ملف
              تُدمج في عمود واحد.
            </p>
            <ColumnList columns={columns} setColumns={setColumns} />
          </div>
          <div>
            <h2>خيارات الدمج</h2>
            <fieldset>
              <legend>أي القطع تظهر؟</legend>
              {[
                ['all', 'كل القطع الموجودة في أي ملف'],
                ['base', 'فقط القطع الموجودة في الملف الأول'],
                ['both', 'فقط القطع الموجودة في كل الملفات'],
              ].map(([v, label]) => (
                <label className="chk" key={v}>
                  <input type="radio" name="mode" checked={mode === v} onChange={() => setMode(v)} />
                  <span>{label}</span>
                </label>
              ))}
            </fieldset>
            <fieldset>
              <legend>القطع المكررة داخل نفس الملف</legend>
              <label className="chk">
                <input type="checkbox" checked={sumDuplicates} onChange={(e) => setSumDuplicates(e.target.checked)} />
                <span>
                  اجمع القيم الرقمية (مثلاً عدة سحوبات لنفس القطعة)
                  <br />
                  <span className="mut small">إذا لم يُفعّل، تؤخذ أول قيمة فقط.</span>
                </span>
              </label>
            </fieldset>
            <fieldset>
              <legend>عمود الإجمالي</legend>
              <label className="chk">
                <input type="checkbox" checked={withTotal} onChange={(e) => setWithTotal(e.target.checked)} />
                <span>أضف عمود «الإجمالي» = الكمية × السعر (كمعادلة في الإكسل)</span>
              </label>
              <label className="f">
                عمود الكمية
                <select value={qty} onChange={(e) => setQtyId(e.target.value)}>
                  {valueCols.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
                </select>
              </label>
              <label className="f">
                عمود السعر
                <select value={price} onChange={(e) => setPriceId(e.target.value)}>
                  {valueCols.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
                </select>
              </label>
            </fieldset>
          </div>
        </div>
      </section>

      <section className="band">
        <h2>النتيجة</h2>
        {result && (
          <div className="stats">
            {[
              [result.rows.length, 'قطعة في الملف الناتج'],
              [result.rows.length - missing, 'مكتملة البيانات'],
              [missing, 'ينقصها بيانات'],
              [result.duplicates, 'صف مكرر داخل الملفات'],
            ].map(([n, l]) => (
              <div className="stat" key={l}><b>{n.toLocaleString('ar')}</b><span>{l}</span></div>
            ))}
          </div>
        )}
        <div className="row" style={{ marginBottom: 12 }}>
          <label className="f" style={{ flex: '1 1 260px', margin: 0 }}>
            اسم الملف
            <input type="text" value={fileName} onChange={(e) => setFileName(e.target.value)} />
          </label>
          <button className="pri" disabled={!result?.rows.length || !result.headers.length} onClick={download}>
            تنزيل ملف الإكسل
          </button>
        </div>
        <Preview result={result} />
        <p className="mut small" style={{ marginTop: 8 }}>
          الصفوف المظللة ينقصها قيمة في أحد الأعمدة (غالباً القطعة غير موجودة في أحد الملفات).
        </p>
      </section>
    </>
  );
}

function DropZone({ onFiles }) {
  const input = useRef(null);
  const [over, setOver] = useState(false);
  return (
    <>
      <div
        className={`drop${over ? ' on' : ''}`}
        tabIndex={0}
        role="button"
        aria-label="اختر ملفات إكسل"
        onClick={() => input.current.click()}
        onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.current.click(); } }}
        onDragOver={(e) => { e.preventDefault(); setOver(true); }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => { e.preventDefault(); setOver(false); onFiles([...e.dataTransfer.files]); }}
      >
        <strong>اسحب ملفات الإكسل هنا أو اضغط للاختيار</strong>
        <span className="mut small">‎.xlsx ‎.xls ‎.csv — يمكنك اختيار أكثر من ملف</span>
      </div>
      <input
        ref={input}
        type="file"
        accept=".xlsx,.xls,.xlsm,.csv"
        multiple
        hidden
        onChange={(e) => { onFiles([...e.target.files]); e.target.value = ''; }}
      />
    </>
  );
}

function SourceCard({ src, isFirst, onChange, onRemove, onMakeFirst }) {
  const kept = src.filters.length ? filteredRows(src).length : null;
  return (
    <div className="card">
      <div className="nm">
        <span>
          <bdi>{src.name}</bdi>{' '}
          <small>
            {isFirst ? '(الملف الأول) · ' : ''}
            {kept === null ? `${src.data.length} صف` : <b className="acc">{kept} من {src.data.length} صف بعد الفلتر</b>}
          </small>
        </span>
        <button className="x" onClick={onRemove}>حذف</button>
      </div>
      {src.wb.SheetNames.length > 1 && (
        <label className="f">
          الورقة
          <select value={src.sheet} onChange={(e) => onChange({ sheet: e.target.value, hdrSet: false, keySet: false })}>
            {src.wb.SheetNames.map((n) => <option key={n}>{n}</option>)}
          </select>
        </label>
      )}
      <label className="f">
        صف العناوين
        <input
          type="number"
          min="1"
          value={src.hdr}
          onChange={(e) => onChange({ hdr: Math.max(1, +e.target.value || 1), hdrSet: true, keySet: false })}
        />
      </label>
      <label className="f">
        عمود الربط
        <select value={src.key} onChange={(e) => onChange({ key: +e.target.value, keySet: true })}>
          {src.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
        </select>
      </label>
      <Filters src={src} onChange={(filters) => onChange({ filters })} />
      {!isFirst && <button className="x" onClick={onMakeFirst}>اجعله الملف الأول</button>}
    </div>
  );
}

const OPS = [
  ['in', 'القيمة واحدة من'],
  ['notIn', 'استبعد القيم'],
  ['contains', 'يحتوي على'],
  ['notContains', 'لا يحتوي على'],
  ['gte', 'أكبر من أو يساوي'],
  ['lte', 'أصغر من أو يساوي'],
];
const LIST_OPS = new Set(['in', 'notIn']);

function Filters({ src, onChange }) {
  const set = (i, patch) => onChange(src.filters.map((f, j) => (j === i ? { ...f, ...patch } : f)));
  const add = () => {
    const col = src.headers.findIndex((_, i) => i !== src.key);
    onChange([...src.filters, { col: col < 0 ? 0 : col, op: 'in', values: [], text: '' }]);
  };
  return (
    <div className="filters">
      {src.filters.map((f, i) => (
        <div className="filter" key={i}>
          <div className="row between">
            <b className="small">فلتر {src.filters.length > 1 ? i + 1 : ''}</b>
            <button className="x" onClick={() => onChange(src.filters.filter((_, j) => j !== i))}>إزالة الفلتر</button>
          </div>
          <label className="f">
            العمود
            <select value={f.col} onChange={(e) => set(i, { col: +e.target.value, values: [] })}>
              {src.headers.map((h, j) => <option key={j} value={j}>{h}</option>)}
            </select>
          </label>
          <label className="f">
            الشرط
            <select value={f.op} onChange={(e) => set(i, { op: e.target.value })}>
              {OPS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
            </select>
          </label>
          {LIST_OPS.has(f.op) ? (
            <ValuePicker src={src} filter={f} onChange={(values) => set(i, { values })} />
          ) : (
            <label className="f">
              القيمة
              <input
                type={f.op === 'gte' || f.op === 'lte' ? 'number' : 'text'}
                value={f.text}
                onChange={(e) => set(i, { text: e.target.value })}
              />
            </label>
          )}
        </div>
      ))}
      <button className="x" onClick={add}>+ إضافة فلتر</button>
    </div>
  );
}

const PICKER_LIMIT = 300;

function ValuePicker({ src, filter, onChange }) {
  const [q, setQ] = useState('');
  const all = useMemo(() => distinctValues(src, filter.col), [src.data, filter.col]);
  const shown = all.filter((d) => !q || d.value.toLowerCase().includes(q.trim().toLowerCase()));
  const chosen = new Set(filter.values);
  const toggle = (v) => onChange(chosen.has(v) ? filter.values.filter((x) => x !== v) : [...filter.values, v]);
  return (
    <div className="picker">
      <div className="row">
        <input type="text" placeholder="ابحث عن قيمة…" value={q} onChange={(e) => setQ(e.target.value)} style={{ flex: '1 1 140px' }} />
        <button className="x" onClick={() => onChange([...new Set([...filter.values, ...shown.map((d) => d.value)])])}>تحديد الظاهر</button>
        <button className="x" onClick={() => onChange([])}>إلغاء التحديد</button>
      </div>
      <div className="vals">
        {shown.slice(0, PICKER_LIMIT).map((d) => (
          <label key={d.value} className="val">
            <input type="checkbox" checked={chosen.has(d.value)} onChange={() => toggle(d.value)} />
            <span>{d.value === '' ? <i className="mut">(فارغ)</i> : <bdi>{d.value}</bdi>}</span>
            <small className="mut">{d.count}</small>
          </label>
        ))}
        {shown.length > PICKER_LIMIT && <div className="mut small">يظهر أول {PICKER_LIMIT} قيمة، استخدم البحث للباقي.</div>}
        {!shown.length && <div className="mut small">لا توجد قيم مطابقة</div>}
      </div>
      <div className="mut small">
        {filter.values.length ? `محدد ${filter.values.length} من ${all.length} قيمة` : 'لم تحدد أي قيمة بعد، فالفلتر غير مفعّل'}
      </div>
    </div>
  );
}

function ColumnList({ columns, setColumns }) {
  if (!columns.length) return <ul className="cols"><li className="empty">ارفع الملفات لتظهر الأعمدة هنا</li></ul>;
  const move = (i, d) =>
    setColumns((cols) => {
      const next = [...cols];
      [next[i], next[i + d]] = [next[i + d], next[i]];
      return next;
    });
  const patch = (i, p) => setColumns((cols) => cols.map((c, j) => (j === i ? { ...c, ...p } : c)));
  return (
    <ul className="cols">
      {columns.map((c, i) => (
        <li key={c.id} className={c.on ? '' : 'off'}>
          <input type="checkbox" checked={c.on} aria-label="إظهار العمود" onChange={(e) => patch(i, { on: e.target.checked })} />
          <span>
            <input type="text" value={c.label} aria-label="اسم العمود" onChange={(e) => patch(i, { label: e.target.value, edited: true })} />
            {c.id === KEY_ID && <span className="tag">عمود الربط</span>}
          </span>
          <button className="mv" disabled={i === 0} aria-label="لأعلى" onClick={() => move(i, -1)}>↑</button>
          <button className="mv" disabled={i === columns.length - 1} aria-label="لأسفل" onClick={() => move(i, 1)}>↓</button>
        </li>
      ))}
    </ul>
  );
}

function Preview({ result }) {
  if (!result) return <div className="scroll"><div className="empty">لا توجد نتيجة بعد</div></div>;
  if (!result.rows.length) return <div className="scroll"><div className="empty">لا توجد قطع مطابقة لخيارات الدمج</div></div>;
  return (
    <div className="scroll">
      <table>
        <thead>
          <tr>{result.headers.map((h, i) => <th key={i}>{h}</th>)}</tr>
        </thead>
        <tbody>
          {result.rows.slice(0, PREVIEW_LIMIT).map((r, i) => (
            <tr key={i} className={r.missing ? 'miss' : ''}>
              {r.values.map((v, j) =>
                v === '' ? <td key={j} className="e" /> : (
                  <td key={j} className={typeof v === 'number' ? 'n' : ''}>{typeof v === 'number' ? +v.toFixed(4) : String(v)}</td>
                ),
              )}
            </tr>
          ))}
        </tbody>
      </table>
      {result.rows.length > PREVIEW_LIMIT && (
        <div className="empty">تُعرض أول {PREVIEW_LIMIT} صف من {result.rows.length}، والملف المنزّل يحتوي الكل.</div>
      )}
    </div>
  );
}
