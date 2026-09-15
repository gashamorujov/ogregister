import { useMemo, useRef, useState } from 'react';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import {
  CloseIcon, DocIcon, CheckIcon, ClipboardIcon, WarningIcon, ArrowLeftIcon,
} from './Icons';
import {
  TARGET_FIELDS, FIELD_LABELS, rowKey, detectColumnMapping, HEADER_MATCH_THRESHOLD,
  emptyFieldsOf, hasIdentitySignal,
} from '../lib/importMapping';

const COL_SCAN_LIMIT = 40;

const STEP_TITLES = {
  choose: 'Import',
  clipboard: 'Clipboard ilə idxal',
  file: 'Excel sənədi ilə idxal',
  preview: 'Önizləmə (Preview)',
};

const PREVIEW_COL_WIDTHS = {
  fullName: 210, serial: 110, idNumber: 120, birthDate: 100, phone: 140,
  email: 190, rank: 170, fullNameId: 200, rank2: 130, courseCode: 100,
  startDate: 110, finishDate: 110, note: 120, date: 100,
};

const ACCEPT_EXT = '.xlsx,.xls,.csv,.tsv,.txt';

function cleanVal(v) {
  return String(v === null || v === undefined ? '' : v).replace(/\[object Object\]/g, '').trim();
}

function toDateStr(d) {
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  return `${dd}.${mm}.${d.getUTCFullYear()}`;
}

// Reads one ExcelJS cell into plain display text (dates, cached formulas, rich text).
function parseCellText(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (v instanceof Date && !isNaN(v.getTime())) return toDateStr(v);
  if (v.richText && Array.isArray(v.richText)) return v.richText.map((rt) => rt.text || '').join('').trim();
  if (v.result !== undefined && v.result !== null) {
    if (v.result instanceof Date && !isNaN(v.result.getTime())) return toDateStr(v.result);
    return String(v.result).trim();
  }
  if (v.text) return String(v.text).trim();
  return cleanVal(v);
}

function findHeaderInSheet(sheet) {
  let best = null;
  for (let r = 1; r <= 5; r++) {
    const row = sheet.getRow(r);
    const texts = [];
    for (let c = 1; c <= COL_SCAN_LIMIT; c++) texts.push(parseCellText(row.getCell(c)));
    const { colIndexToField, matchedFields } = detectColumnMapping(texts);
    if (matchedFields.size >= HEADER_MATCH_THRESHOLD && (!best || matchedFields.size > best.matchedFields.size)) {
      best = { rowNumber: r, colIndexToField, matchedFields };
    }
  }
  return best;
}

// Quote-aware CSV line splitter (handles quoted fields with commas).
function splitLine(line, delimiter) {
  if (delimiter !== ',') return line.split(delimiter);
  const out = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQ) {
      if (ch === '"') {
        if (line[i + 1] === '"') { cur += '"'; i++; }
        else inQ = false;
      } else cur += ch;
    } else if (ch === '"') {
      inQ = true;
    } else if (ch === delimiter) {
      out.push(cur); cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

function detectDelimiter(text) {
  const line = (text.split(/\r?\n/).find(l => l.trim() !== '') || '');
  const counts = { '\t': (line.match(/\t/g) || []).length, ',': (line.match(/,/g) || []).length, ';': (line.match(/;/g) || []).length };
  if (counts['\t'] > 0) return '\t';
  if (counts[';'] > counts[',']) return ';';
  return ',';
}

// Converts raw text (clipboard or CSV/TSV file) into a 2D matrix of cells.
function textToMatrix(text, delimiter) {
  return text.split(/\r?\n/)
    .filter(l => l.trim() !== '')
    .map(l => splitLine(l, delimiter).map(x => cleanVal(x)));
}

// Header detection + row parsing shared by clipboard, CSV/TSV and .xls files.
function recordsFromMatrix(matrix) {
  let headerIdx = -1;
  let mapping = null;
  const scanLimit = Math.min(matrix.length, 5);
  for (let i = 0; i < scanLimit; i++) {
    const cells = matrix[i] || [];
    const { colIndexToField, matchedFields } = detectColumnMapping(cells);
    if (matchedFields.size >= HEADER_MATCH_THRESHOLD && (!mapping || matchedFields.size > mapping.matchedFields.size)) {
      headerIdx = i; mapping = { colIndexToField, matchedFields };
    }
  }

  let fellBack = false;
  if (!mapping) {
    fellBack = true;
    const colIndexToField = new Map();
    TARGET_FIELDS.forEach((f, i) => colIndexToField.set(i + 1, f));
    mapping = { colIndexToField, matchedFields: new Set(TARGET_FIELDS) };
    headerIdx = -1;
  }

  const records = [];
  for (let i = headerIdx + 1; i < matrix.length; i++) {
    const cells = matrix[i] || [];
    if (detectColumnMapping(cells).matchedFields.size >= HEADER_MATCH_THRESHOLD) continue; // repeated header guard
    const rec = {};
    mapping.colIndexToField.forEach((field, idx) => { rec[field] = cells[idx] || ''; });
    if (hasIdentitySignal(rec)) records.push(rec);
  }
  return { records, matchedFields: mapping.matchedFields, fellBack };
}

export default function ImportExcelModal({ existingRows = [], onConfirm, onCancel }) {
  const [step, setStep] = useState('choose');
  const [source, setSource] = useState('');
  const [fileName, setFileName] = useState('');
  const [clipboardText, setClipboardText] = useState('');
  const [previewRows, setPreviewRows] = useState(null);
  const [newCount, setNewCount] = useState(0);
  const [updateCount, setUpdateCount] = useState(0);
  const [skippedCount, setSkippedCount] = useState(0);
  const [emptyCount, setEmptyCount] = useState(0);
  const [missingCols, setMissingCols] = useState([]);
  const [usedFallback, setUsedFallback] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  const existingByKey = useMemo(() => {
    const m = new Map();
    (existingRows || []).forEach(r => { if (r && r._id) m.set(rowKey(r), r); });
    return m;
  }, [existingRows]);

  const processParsed = (records, matchedFields, fellBack) => {
    if (!records || records.length === 0) { setError('Heç bir məlumat tapılmadı.'); return; }
    const result = records.map((rec) => {
      const existing = existingByKey.get(rowKey(rec));
      if (!existing) return { ...rec, _status: 'new', EMPTY_FIELDS: emptyFieldsOf(rec), CHANGED_FIELDS: [] };
      const changes = [];
      TARGET_FIELDS.forEach(f => {
        const nv = String(rec[f] || '').trim();
        const ov = String(existing[f] || '').trim();
        if (nv && nv !== ov) changes.push(f);
      });
      if (changes.length === 0) {
        return { ...rec, _status: 'same', EMPTY_FIELDS: emptyFieldsOf(rec), CHANGED_FIELDS: [] };
      }
      return { _id: existing._id, ...existing, ...rec, _status: 'update', EMPTY_FIELDS: emptyFieldsOf(rec), CHANGED_FIELDS: changes };
    });
    setNewCount(result.filter(r => r._status === 'new').length);
    setUpdateCount(result.filter(r => r._status === 'update').length);
    setSkippedCount(result.filter(r => r._status === 'same').length);
    setEmptyCount(result.filter(r => (r.EMPTY_FIELDS || []).length > 0).length);
    setMissingCols(TARGET_FIELDS.filter(f => !matchedFields.has(f)).map(f => FIELD_LABELS[f]));
    setUsedFallback(!!fellBack);
    setPreviewRows(result);
    setStep('preview');
  };

  const resetOutcome = () => {
    setError(''); setPreviewRows(null); setNewCount(0); setUpdateCount(0);
    setSkippedCount(0); setEmptyCount(0); setMissingCols([]); setUsedFallback(false);
  };

  // ---------- Excel .xlsx path (REGİSTR-2026 sheet) ----------

  const parseXlsx = async (file) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const sheet = workbook.getWorksheet('REGİSTR-2026');
    if (!sheet) throw new Error('NO_SHEET');
    const header = findHeaderInSheet(sheet);
    if (!header) throw new Error('NO_HEADER');

    const parsed = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber <= header.rowNumber) return;
      const texts = [];
      for (let c = 1; c <= COL_SCAN_LIMIT; c++) texts.push(parseCellText(row.getCell(c)));
      if (detectColumnMapping(texts).matchedFields.size >= HEADER_MATCH_THRESHOLD) return;
      const rec = {};
      header.colIndexToField.forEach((field, idx) => { rec[field] = texts[idx] || ''; });
      if (hasIdentitySignal(rec)) parsed.push(rec);
    });
    processParsed(parsed, header.matchedFields, false);
  };

  // ---------- Legacy .xls path (SheetJS) ----------

  const parseXls = async (file) => {
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
    const sheetName = wb.SheetNames.find(n => /registr/i.test(n)) || wb.SheetNames[0];
    if (!sheetName) throw new Error('NO_SHEET');
    const ws = wb.Sheets[sheetName];
    const rows2d = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
    const matrix = rows2d.map(row => (row || []).map(cv => {
      if (cv instanceof Date && !isNaN(cv.getTime())) return toDateStr(cv);
      return cleanVal(cv);
    }));
    const { records, matchedFields, fellBack } = recordsFromMatrix(matrix);
    processParsed(records, matchedFields, fellBack);
  };

  // ---------- CSV / TSV text path ----------

  const parseText = async (file) => {
    const text = await file.text();
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    let delimiter;
    if (ext === 'tsv') delimiter = '\t';
    else if (ext === 'csv') delimiter = detectDelimiter(text) === '\t' ? ',' : detectDelimiter(text);
    else delimiter = detectDelimiter(text);
    if (textToMatrix(text, delimiter).length === 0) throw new Error('EMPTY');
    const matrix = textToMatrix(text, delimiter);
    const { records, matchedFields, fellBack } = recordsFromMatrix(matrix);
    processParsed(records, matchedFields, fellBack);
  };

  // ---------- File entry ----------

  const handleFile = async (file) => {
    resetOutcome();
    if (!file) return;
    setFileName(file.name);
    setSource('file');
    setLoading(true);
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    try {
      if (ext === 'xlsx') await parseXlsx(file);
      else if (ext === 'xls') await parseXls(file);
      else if (ext === 'csv' || ext === 'tsv' || ext === 'txt') await parseText(file);
      else throw new Error('FORMAT');
    } catch (err) {
      console.error('Fayl oxunma xətası:', err);
      if (err && err.message === 'NO_SHEET') setError('Faylda "REGİSTR-2026" səhifəsi tapılmadı.');
      else if (err && err.message === 'NO_HEADER') setError('Sütun başlıqları tanınmadı. Fayl strukturunu yoxlayın.');
      else if (err && err.message === 'EMPTY') setError('Faylda oxuna bilən məlumat yoxdur.');
      else if (err && err.message === 'FORMAT') setError('Dəstəklənən formatlar: XLSX, XLS, CSV, TSV.');
      else setError('Fayl oxunarkən xəta baş verdi. Düzgün Excel faylı seçdiyinizə əmin olun.');
    } finally {
      setLoading(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

  // ---------- Clipboard path ----------

  const handleClipboardParse = () => {
    resetOutcome();
    if (!clipboardText.trim()) { setError('Kopyaladığınız mətn boşdur.'); return; }
    setSource('clipboard');
    const matrix = textToMatrix(clipboardText, '\t');
    const { records, matchedFields, fellBack } = recordsFromMatrix(matrix);
    processParsed(records, matchedFields, fellBack);
  };

  // ---------- Shared ----------

  const stripMeta = (r) => {
    const copy = { ...r };
    delete copy._status; delete copy.EMPTY_FIELDS; delete copy.CHANGED_FIELDS;
    return copy;
  };

  const handleConfirm = () => {
    const added = (previewRows || []).filter(r => r._status === 'new').map(stripMeta);
    const updated = (previewRows || []).filter(r => r._status === 'update').map(stripMeta);
    onConfirm({ added, updated });
  };

  const goBack = () => {
    setError('');
    if (step === 'preview') {
      setPreviewRows(null);
      setStep(source === 'clipboard' ? 'clipboard' : 'file');
    } else {
      setStep('choose');
    }
  };

  const badgeFor = (status) => {
    if (status === 'new') return <span className="import-badge-new">Yeni</span>;
    if (status === 'update') return <span className="import-badge-new import-badge-update">Yenilənəcək</span>;
    return <span className="import-badge-new import-badge-existing">Mövcud</span>;
  };

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" style={{ maxWidth: step === 'preview' ? 1000 : 640 }} onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>{STEP_TITLES[step]}</h2>
          <button className="modal-close" onClick={onCancel} aria-label="Bağla"><CloseIcon /></button>
        </div>
        <div className="modal-body">

          {step === 'choose' && (
            <div className="import-choice-wrap">
              <p className="import-choice-lead">Məlumatları necə idxal etmək istəyirsiniz?</p>
              <div className="import-choice-grid">
                <button type="button" className="import-choice-card" onClick={() => { setSource('clipboard'); setStep('clipboard'); }}>
                  <span className="import-choice-icon"><ClipboardIcon /></span>
                  <span className="import-choice-title">Clipboard</span>
                  <span className="import-choice-desc">Excel-dən kopyaladığınız (Ctrl+A → Ctrl+C) məlumatları birbaşa yapışdırın</span>
                </button>
                <button type="button" className="import-choice-card" onClick={() => { setSource('file'); setStep('file'); }}>
                  <span className="import-choice-icon"><DocIcon /></span>
                  <span className="import-choice-title">Excel Document</span>
                  <span className="import-choice-desc">Fayl yükləyin — XLSX, XLS, CSV, TSV dəstəklənir</span>
                </button>
              </div>
            </div>
          )}

          {step === 'clipboard' && (
            <div className="import-step">
              <p className="import-clipboard-hint">
                Excel-də <b>REGİSTR-2026</b> səhifəsini açın, <b>Ctrl+A</b> ilə hamısını seçin, <b>Ctrl+C</b> edin və aşağıya yapışdırın.
              </p>
              <textarea
                className="import-clipboard-area"
                rows={9}
                autoFocus
                placeholder="Excel-dən kopyaladığınız mətni buraya yapışdırın (Ctrl+V)..."
                value={clipboardText}
                onChange={(e) => setClipboardText(e.target.value)}
              />
              <button className="btn-primary import-clipboard-btn" onClick={handleClipboardParse} disabled={!clipboardText.trim()}>
                <ClipboardIcon /> Məlumatları analiz et
              </button>
              {error && <div className="import-parse-error"><WarningIcon /> {error}</div>}
            </div>
          )}

          {step === 'file' && (
            <div className="import-step">
              <div
                className={`import-dropzone ${dragOver ? 'drag' : ''}`}
                onClick={() => !loading && inputRef.current?.click()}
                onDragOver={(e) => { e.preventDefault(); if (!loading) setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(e) => { if (loading) { e.preventDefault(); return; } handleDrop(e); }}
              >
                <div className="import-dropzone-icon"><DocIcon /></div>
                <div className="import-dropzone-text">{loading ? 'Fayl analiz edilir...' : 'Excel faylı seçin və ya buraya sürükləyin'}</div>
                <div className="import-dropzone-sub">XLSX · XLS · CSV · TSV — .xlsx üçün yalnız "REGİSTR-2026" səhifəsi oxunacaq</div>
                <input
                  ref={inputRef}
                  type="file"
                  accept={ACCEPT_EXT}
                  style={{ display: 'none' }}
                  onChange={(e) => handleFile(e.target.files?.[0])}
                />
              </div>
              {fileName && !loading && !error && <div className="import-file-name"><CheckIcon /> {fileName}</div>}
              {error && <div className="import-parse-error"><WarningIcon /> {error}</div>}
            </div>
          )}

          {step === 'preview' && previewRows && (
            <>
              <div className="import-note-new">
                <CheckIcon />
                <span>Preview-də mənbə ilə mövcud məlumatlar müqayisə olunur. Təsdiqdən sonra dəyişikliklər Firebase-ə yazılır və hər kəsə anında görünür.</span>
              </div>

              <div className="import-summary-stats">
                <div className="import-stat import-stat-new"><b>{newCount}</b><span>yeni əlavə olunacaq</span></div>
                {updateCount > 0 && (
                  <div className="import-stat import-stat-upd"><b>{updateCount}</b><span>dəyişəcək sətir</span></div>
                )}
                {skippedCount > 0 && (
                  <div className="import-stat import-stat-dup"><b>{skippedCount}</b><span>mövcuddur (dəyişməyib)</span></div>
                )}
                {emptyCount > 0 && (
                  <div className="import-stat import-stat-warn"><b>{emptyCount}</b><span>sətirdə boş xana var</span></div>
                )}
              </div>

              {missingCols.length > 0 && (
                <div className="import-note-warn">
                  <WarningIcon />
                  <span>Bu sütunlar mənbədə tapılmadı: {missingCols.join(', ')}</span>
                </div>
              )}
              {usedFallback && (
                <div className="import-note-warn">
                  <WarningIcon />
                  <span>Sütun başlıqları tanınmadı — məlumatlar defolt sütun ardıcıllığı ilə oxundu. Nəticəni diqqətlə yoxlayın.</span>
                </div>
              )}

              <div className="import-preview-table-wrap">
                <table className="import-preview-table import-preview-table-full">
                  <colgroup>
                    <col style={{ width: 110 }} />
                    {TARGET_FIELDS.map((f) => <col key={f} style={{ width: PREVIEW_COL_WIDTHS[f] || 140 }} />)}
                  </colgroup>
                  <thead>
                    <tr>
                      <th>Status</th>
                      {TARGET_FIELDS.map((f) => <th key={f}>{FIELD_LABELS[f]}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.length === 0 && (
                      <tr><td colSpan={TARGET_FIELDS.length + 1} className="import-empty">Heç bir məlumat tapılmadı.</td></tr>
                    )}
                    {previewRows.map((r, i) => (
                      <tr key={i}>
                        <td>{badgeFor(r._status)}</td>
                        {TARGET_FIELDS.map((f) => {
                          const val = r[f];
                          const isEmpty = !String(val || '').trim();
                          const changed = r._status === 'update' && (r.CHANGED_FIELDS || []).includes(f);
                          const cls = isEmpty ? 'import-cell-empty' : (changed ? 'import-cell-changed' : (r._status === 'new' ? 'import-colnew' : ''));
                          return (
                            <td key={f} className={cls}>
                              {isEmpty ? <span className="import-no-field">boş</span> : val}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}

        </div>
        <div className="modal-footer">
          {step === 'choose' ? (
            <button className="btn-secondary" onClick={onCancel}>Ləğv et</button>
          ) : (
            <button className="btn-secondary" onClick={goBack}><ArrowLeftIcon /> Geri</button>
          )}
          {step === 'preview' && (
            <button className="btn-primary" disabled={newCount + updateCount === 0} onClick={handleConfirm}>
              <CheckIcon /> Təsdiqlə ({newCount + updateCount})
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
