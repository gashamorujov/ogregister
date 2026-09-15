// ExcelGrid — Excel-like spreadsheet built on React + virtual scrolling.
// Supports: cell editing, multi-cell selection, copy/paste, drag fill,
// keyboard navigation, undo/redo, auto-save via onCellEdit, row context menu.
import {
  useState, useRef, useEffect, useCallback, useMemo,
  useImperativeHandle, forwardRef,
} from 'react';

export const ROW_HEIGHT = 44;
export const HEADER_HEIGHT = 54;
const BUFFER_ROWS = 15;
export const ROW_NUM_WIDTH = 56;
const MAX_HISTORY = 100;

export const GRID_COLUMNS = [
  { field: 'fullName', headerName: 'Soyad, Ad və Ata adı', width: 230, align: 'left' },
  { field: 'serial', headerName: 'Seriya nömrəsi', width: 130, align: 'center' },
  { field: 'idNumber', headerName: 'Fərdi ID nömrəsi', width: 140, align: 'center' },
  { field: 'birthDate', headerName: 'Doğum tarixi', width: 115, align: 'center' },
  { field: 'phone', headerName: 'Telefon', width: 185, align: 'left' },
  { field: 'email', headerName: 'Email', width: 235, align: 'left' },
  { field: 'rank', headerName: 'Rank (Working Diploma)', width: 190, align: 'left' },
  { field: 'fullNameId', headerName: 'Full Name (ID)', width: 220, align: 'left' },
  { field: 'rank2', headerName: 'Rank / Vəzifə', width: 155, align: 'left' },
  { field: 'courseCode', headerName: 'Course Code', width: 100, align: 'center' },
  { field: 'startDate', headerName: 'Başlama tarixi', width: 115, align: 'center' },
  { field: 'finishDate', headerName: 'Bitmə tarixi', width: 115, align: 'center' },
  { field: 'note', headerName: 'Qeyd', width: 160, align: 'left' },
  { field: 'date', headerName: 'Tarix', width: 100, align: 'center' },
];

function tryParseNumber(s) {
  if (s === '' || s == null) return null;
  const n = Number(String(s).trim());
  return Number.isFinite(n) ? n : null;
}

function parseDateStr(s) {
  if (!s) return null;
  const p = String(s).split('.');
  if (p.length !== 3) return null;
  const d = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
  return isNaN(d.getTime()) ? null : d;
}

function formatDate(d) {
  return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
}

// Excel drag-fill semantics: numbers increment by +dir, dates move day-by-day.
function advanceValue(val, dir) {
  const num = tryParseNumber(val);
  if (num !== null && String(val).trim() !== '') return String(num + dir);
  const date = parseDateStr(val);
  if (date) { date.setDate(date.getDate() + dir); return formatDate(date); }
  return val;
}

const ExcelGrid = forwardRef(function ExcelGrid(
  { rows, onCellEdit, onHeaderClick, onRowContextMenu },
  ref
) {
  const containerRef = useRef(null);   // scrolling body
  const headerRef = useRef(null);      // fixed header (horizontal sync)
  const inputRef = useRef(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [containerHeight, setContainerHeight] = useState(600);
  const [activeCell, setActiveCell] = useState(null);
  const [selectionRange, setSelectionRange] = useState(null);
  const [isSelecting, setIsSelecting] = useState(false);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');
  const [fillDragging, setFillDragging] = useState(false);
  const [fillEnd, setFillEnd] = useState(null);
  const undoStackRef = useRef([]);
  const redoStackRef = useRef([]);
  const rowsRef = useRef(rows);
  rowsRef.current = rows;

  const totalContentWidth = useMemo(() =>
    ROW_NUM_WIDTH + GRID_COLUMNS.reduce((s, c) => s + c.width, 0), []);
  const totalHeight = rows.length * ROW_HEIGHT;
  const startIdx = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - BUFFER_ROWS);
  const endIdx = Math.min(rows.length, Math.ceil((scrollTop + containerHeight) / ROW_HEIGHT) + BUFFER_ROWS);
  const visibleRows = useMemo(() => rows.slice(startIdx, endIdx), [rows, startIdx, endIdx]);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const obs = new ResizeObserver(entries => {
      for (const entry of entries) setContainerHeight(entry.contentRect.height);
    });
    obs.observe(el);
    setContainerHeight(el.clientHeight);
    return () => obs.disconnect();
  }, []);

  const handleScroll = useCallback(() => {
    const el = containerRef.current;
    if (!el) return;
    setScrollTop(el.scrollTop);
    if (headerRef.current) headerRef.current.scrollLeft = el.scrollLeft;
  }, []);

  const isCellInSelection = useCallback((row, col) => {
    if (!selectionRange) return false;
    const { startRow, startCol, endRow, endCol } = selectionRange;
    return row >= Math.min(startRow, endRow) && row <= Math.max(startRow, endRow) &&
           col >= Math.min(startCol, endCol) && col <= Math.max(startCol, endCol);
  }, [selectionRange]);

  const pushBatch = useCallback((batch) => {
    if (!batch || batch.length === 0) return;
    undoStackRef.current = [...undoStackRef.current.slice(-MAX_HISTORY + 1), batch];
    redoStackRef.current = [];
  }, []);

  const applyBatch = useCallback((batch) => {
    batch.forEach(ch => onCellEdit(ch.rowId, ch.field, ch.newValue));
  }, [onCellEdit]);

  // ---- Editing ----

  const commitEdit = useCallback(() => {
    if (!editingCell) return;
    const { row, col } = editingCell;
    const field = GRID_COLUMNS[col].field;
    const current = rowsRef.current;
    const oldValue = current[row]?.[field] == null ? '' : String(current[row][field]);
    const newValue = editValue;
    if (newValue !== oldValue) {
      const rowId = current[row]?._id;
      if (rowId) {
        const batch = [{ rowId, field, oldValue, newValue }];
        pushBatch(batch);
        applyBatch(batch);
      }
    }
    setEditingCell(null);
  }, [editingCell, editValue, pushBatch, applyBatch]);

  const cancelEdit = useCallback(() => setEditingCell(null), []);

  const startEditing = useCallback((row, col, initialValue) => {
    if (editingCell) commitEdit();
    const field = GRID_COLUMNS[col].field;
    const val = initialValue !== undefined
      ? initialValue
      : (rowsRef.current[row]?.[field] == null ? '' : String(rowsRef.current[row][field]));
    setEditingCell({ row, col });
    setEditValue(val);
  }, [editingCell, commitEdit]);

  const handleUndo = useCallback(() => {
    if (undoStackRef.current.length === 0 || editingCell) return;
    commitEdit();
    const batch = undoStackRef.current[undoStackRef.current.length - 1];
    undoStackRef.current = undoStackRef.current.slice(0, -1);
    redoStackRef.current = [...redoStackRef.current, batch];
    batch.forEach(ch => onCellEdit(ch.rowId, ch.field, ch.oldValue));
  }, [editingCell, commitEdit, onCellEdit]);

  const handleRedo = useCallback(() => {
    if (redoStackRef.current.length === 0 || editingCell) return;
    const batch = redoStackRef.current[redoStackRef.current.length - 1];
    redoStackRef.current = redoStackRef.current.slice(0, -1);
    undoStackRef.current = [...undoStackRef.current, batch];
    batch.forEach(ch => onCellEdit(ch.rowId, ch.field, ch.newValue));
  }, [editingCell, onCellEdit]);

  useImperativeHandle(ref, () => ({
    undo: handleUndo,
    redo: handleRedo,
    focusGrid: () => containerRef.current?.focus(),
  }));

  // ---- Selection / clipboard / fill ----

  const clearSelection = useCallback(() => {
    if (!selectionRange) return;
    const { startRow: sr, startCol: sc, endRow: er, endCol: ec } = selectionRange;
    const r1 = Math.min(sr, er), r2 = Math.max(sr, er);
    const c1 = Math.min(sc, ec), c2 = Math.max(sc, ec);
    const batch = [];
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const field = GRID_COLUMNS[c].field;
        const rowId = rowsRef.current[r]?._id;
        const oldValue = rowsRef.current[r]?.[field] == null ? '' : String(rowsRef.current[r][field]);
        if (rowId && oldValue !== '') batch.push({ rowId, field, oldValue, newValue: '' });
      }
    }
    if (batch.length) { pushBatch(batch); applyBatch(batch); }
  }, [selectionRange, pushBatch, applyBatch]);

  const copySelection = useCallback(async () => {
    if (!selectionRange) return;
    const { startRow: sr, startCol: sc, endRow: er, endCol: ec } = selectionRange;
    const r1 = Math.min(sr, er), r2 = Math.max(sr, er);
    const c1 = Math.min(sc, ec), c2 = Math.max(sc, ec);
    const lines = [];
    for (let r = r1; r <= r2; r++) {
      const cells = [];
      for (let c = c1; c <= c2; c++) {
        const v = rowsRef.current[r]?.[GRID_COLUMNS[c].field];
        cells.push(v == null ? '' : String(v));
      }
      lines.push(cells.join('\t'));
    }
    try { await navigator.clipboard.writeText(lines.join('\n')); } catch { /* clipboard unavailable */ }
  }, [selectionRange]);

  const pasteSelection = useCallback(async () => {
    if (!activeCell) return;
    try {
      const text = await navigator.clipboard.readText();
      if (!text) return;
      const lines = text.split(/\r?\n/).filter(l => l.length > 0);
      const { row: ar, col: ac } = activeCell;
      const batch = [];
      lines.forEach((line, li) => {
        line.split('\t').forEach((val, ci) => {
          const row = ar + li;
          const col = ac + ci;
          if (row < rowsRef.current.length && col < GRID_COLUMNS.length) {
            const field = GRID_COLUMNS[col].field;
            const rowId = rowsRef.current[row]?._id;
            const oldValue = rowsRef.current[row]?.[field] == null ? '' : String(rowsRef.current[row][field]);
            if (rowId && String(val) !== oldValue) batch.push({ rowId, field, oldValue, newValue: String(val) });
          }
        });
      });
      if (batch.length) { pushBatch(batch); applyBatch(batch); }
    } catch { /* clipboard unavailable */ }
  }, [activeCell, pushBatch, applyBatch]);

  const handleCellMouseDown = useCallback((row, col, e) => {
    if (e.button !== 0) return;
    if (editingCell) commitEdit();
    if (e.shiftKey && activeCell) {
      setSelectionRange(p => ({ ...p, endRow: row, endCol: col }));
    } else {
      setActiveCell({ row, col });
      setSelectionRange({ startRow: row, startCol: col, endRow: row, endCol: col });
      setIsSelecting(true);
    }
    containerRef.current?.focus();
  }, [activeCell, editingCell, commitEdit]);

  const handleCellMouseEnter = useCallback((row, col) => {
    if (isSelecting && activeCell) setSelectionRange(p => ({ ...p, endRow: row, endCol: col }));
    if (fillDragging) setFillEnd({ row, col });
  }, [isSelecting, activeCell, fillDragging]);

  const handleMouseUp = useCallback(() => {
    setIsSelecting(false);
  }, []);

  // Release-drag anywhere outside the grid still ends the selection drag.
  useEffect(() => {
    const up = () => setIsSelecting(false);
    window.addEventListener('mouseup', up);
    return () => window.removeEventListener('mouseup', up);
  }, []);

  // Fill: commit once the drag ends (works in all four directions).
  useEffect(() => {
    if (!fillDragging || !fillEnd || !selectionRange) return;
    const { startRow: sr, startCol: sc, endRow: er, endCol: ec } = selectionRange;
    const r1 = Math.min(sr, er), r2 = Math.max(sr, er);
    const c1 = Math.min(sc, ec), c2 = Math.max(sc, ec);
    const batch = [];
    if (fillEnd.row > r2) {
      for (let c = c1; c <= c2; c++) {
        const field = GRID_COLUMNS[c].field;
        const srcVal = rowsRef.current[r2]?.[field];
        for (let r = r2 + 1; r <= Math.min(fillEnd.row, rowsRef.current.length - 1); r++) {
          const rowId = rowsRef.current[r]?._id;
          if (!rowId) continue;
          const oldValue = rowsRef.current[r]?.[field] == null ? '' : String(rowsRef.current[r][field]);
          const newValue = advanceValue(srcVal, r - r2);
          if (oldValue !== newValue) batch.push({ rowId, field, oldValue, newValue });
        }
      }
    } else if (fillEnd.row < r1) {
      for (let c = c1; c <= c2; c++) {
        const field = GRID_COLUMNS[c].field;
        const srcVal = rowsRef.current[r1]?.[field];
        for (let r = Math.max(0, fillEnd.row); r < r1; r++) {
          const rowId = rowsRef.current[r]?._id;
          if (!rowId) continue;
          const oldValue = rowsRef.current[r]?.[field] == null ? '' : String(rowsRef.current[r][field]);
          const newValue = advanceValue(srcVal, r1 - r);
          if (oldValue !== newValue) batch.push({ rowId, field, oldValue, newValue });
        }
      }
    }
    if (fillEnd.col > c2) {
      for (let r = r1; r <= r2; r++) {
        const srcRow = rowsRef.current[r];
        const srcVal = srcRow?.[GRID_COLUMNS[c2].field];
        for (let c = c2 + 1; c <= Math.min(fillEnd.col, GRID_COLUMNS.length - 1); c++) {
          const rowId = srcRow?._id;
          if (!rowId) continue;
          const field = GRID_COLUMNS[c].field;
          const oldValue = srcRow?.[field] == null ? '' : String(srcRow[field]);
          const newValue = advanceValue(srcVal, c - c2);
          if (oldValue !== newValue) batch.push({ rowId, field, oldValue, newValue });
        }
      }
    } else if (fillEnd.col < c1) {
      for (let r = r1; r <= r2; r++) {
        const srcRow = rowsRef.current[r];
        const srcVal = srcRow?.[GRID_COLUMNS[c1].field];
        for (let c = Math.max(0, fillEnd.col); c < c1; c++) {
          const rowId = srcRow?._id;
          if (!rowId) continue;
          const field = GRID_COLUMNS[c].field;
          const oldValue = srcRow?.[field] == null ? '' : String(srcRow[field]);
          const newValue = advanceValue(srcVal, c1 - c);
          if (oldValue !== newValue) batch.push({ rowId, field, oldValue, newValue });
        }
      }
    }
    if (batch.length) { pushBatch(batch); applyBatch(batch); }
    setFillDragging(false);
    setFillEnd(null);
  }, [fillDragging, fillEnd, selectionRange, pushBatch, applyBatch]);

  // ---- Keyboard ----

  const scrollToRow = useCallback((idx) => {
    const el = containerRef.current;
    if (!el) return;
    const top = idx * ROW_HEIGHT, bottom = top + ROW_HEIGHT;
    if (top < el.scrollTop) el.scrollTop = top;
    else if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = bottom - el.clientHeight;
  }, []);

  useEffect(() => { if (activeCell) scrollToRow(activeCell.row); }, [activeCell, scrollToRow]);

  const handleContainerKeyDown = useCallback((e) => {
    if (e.ctrlKey && e.key === 'z') { e.preventDefault(); handleUndo(); return; }
    if (e.ctrlKey && e.key === 'y') { e.preventDefault(); handleRedo(); return; }
    if (e.ctrlKey && e.key === 'c') { e.preventDefault(); copySelection(); return; }
    if (e.ctrlKey && e.key === 'v') { e.preventDefault(); pasteSelection(); return; }
    if ((e.key === 'Delete' || (e.key === 'Backspace' && !e.ctrlKey))) {
      if (activeCell && editingCell == null) { e.preventDefault(); clearSelection(); return; }
    }
    if (e.key === 'F2' && activeCell && editingCell == null) {
      e.preventDefault();
      startEditing(activeCell.row, activeCell.col);
      return;
    }
    if (editingCell) return; // input handles its own keys
    if (!activeCell) return;
    let { row, col } = activeCell;
    const last = rowsRef.current.length - 1;
    switch (e.key) {
      case 'ArrowUp': e.preventDefault(); row = Math.max(0, row - 1); break;
      case 'ArrowDown': e.preventDefault(); row = Math.min(last, row + 1); break;
      case 'ArrowLeft': e.preventDefault(); col = Math.max(0, col - 1); break;
      case 'ArrowRight': e.preventDefault(); col = Math.min(GRID_COLUMNS.length - 1, col + 1); break;
      case 'Tab': e.preventDefault(); col = e.shiftKey ? Math.max(0, col - 1) : Math.min(GRID_COLUMNS.length - 1, col + 1); break;
      case 'Enter': e.preventDefault(); row = Math.min(last, row + 1); break;
      case 'PageDown': e.preventDefault(); row = Math.min(last, row + Math.max(1, Math.floor(containerHeight / ROW_HEIGHT))); break;
      case 'PageUp': e.preventDefault(); row = Math.max(0, row - Math.max(1, Math.floor(containerHeight / ROW_HEIGHT))); break;
      case 'Home':
        if (e.ctrlKey) { e.preventDefault(); row = 0; col = 0; } else return;
        break;
      case 'End':
        if (e.ctrlKey) { e.preventDefault(); row = last; col = GRID_COLUMNS.length - 1; } else return;
        break;
      default:
        if (e.key.length === 1 && !e.metaKey) {
          e.preventDefault();
          startEditing(row, col, e.key);
        }
        return;
    }
    setActiveCell({ row, col });
    setSelectionRange({ startRow: row, startCol: col, endRow: row, endCol: col });
  }, [
    activeCell, editingCell, containerHeight,
    handleUndo, handleRedo, copySelection, pasteSelection,
    clearSelection, startEditing,
  ]);

  const cellCls = useCallback((row, col) => {
    const cls = [];
    if (activeCell?.row === row && activeCell?.col === col) cls.push('eg-active');
    if (isCellInSelection(row, col)) cls.push('eg-selected');
    if (editingCell?.row === row && editingCell?.col === col) cls.push('eg-editing');
    return cls.join(' ');
  }, [activeCell, isCellInSelection, editingCell]);

  // Right-click -> row context menu (index is relative to `rows` prop).
  const handleContextMenu = useCallback((e) => {
    e.preventDefault();
    if (!onRowContextMenu) return;
    const target = e.target.closest('.eg-row');
    let rowIdx = 0;
    if (target) {
      const id = target.getAttribute('data-row-id');
      const vi = visibleRows.findIndex(r => r._id === id);
      if (vi >= 0) rowIdx = startIdx + vi;
    } else {
      const rect = containerRef.current?.getBoundingClientRect();
      if (rect) {
        const y = e.clientY - rect.top + containerRef.current.scrollTop;
        const idx = Math.floor(y / ROW_HEIGHT);
        rowIdx = Math.max(0, Math.min(rows.length - 1, idx));
      }
    }
    onRowContextMenu(rowIdx, Math.min(e.clientX, window.innerWidth - 240), Math.min(e.clientY, window.innerHeight - 180));
  }, [onRowContextMenu, visibleRows, startIdx, rows.length]);

  return (
    <div className="excel-grid-wrap">
      {/* Fixed header (horizontally synced with the body) */}
      <div className="eg-header" ref={headerRef}>
        <div className="eg-header-inner" style={{ width: totalContentWidth }}>
          <div className="eg-header-cell eg-row-num-header" style={{ width: ROW_NUM_WIDTH, minWidth: ROW_NUM_WIDTH, maxWidth: ROW_NUM_WIDTH }}>№</div>
          {GRID_COLUMNS.map((col) => (
            <div
              className="eg-header-cell"
              key={col.field}
              style={{ width: col.width, minWidth: col.width, maxWidth: col.width }}
              onClick={(e) => { e.stopPropagation(); if (onHeaderClick) onHeaderClick(col.field); }}
              title={col.headerName}
            >
              <span className="eg-header-text">{col.headerName}</span>
              <span className="eg-filter-indicator" aria-hidden="true">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="12" height="12">
                  <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
                </svg>
              </span>
            </div>
          ))}
        </div>
      </div>

      {/* Scrollable body with virtual rows */}
      <div
        ref={containerRef}
        className="eg-scroll"
        tabIndex={0}
        onScroll={handleScroll}
        onKeyDown={handleContainerKeyDown}
        onMouseUp={handleMouseUp}
        onContextMenu={handleContextMenu}
      >
        <div className="eg-body" style={{ width: totalContentWidth, height: totalHeight }}>
          {visibleRows.map((row, vi) => {
            const ri = startIdx + vi;
            return (
              <div
                className="eg-row"
                key={row._id || ri}
                data-row-id={row._id}
                style={{ top: ri * ROW_HEIGHT, height: ROW_HEIGHT }}
              >
                <div className="eg-cell eg-row-num" style={{ width: ROW_NUM_WIDTH, minWidth: ROW_NUM_WIDTH, maxWidth: ROW_NUM_WIDTH }}>{ri + 1}</div>
                {GRID_COLUMNS.map((col, ci) => {
                  const val = row[col.field] == null ? '' : String(row[col.field]);
                  const isEditing = editingCell?.row === ri && editingCell?.col === ci;
                  return (
                    <div
                      className={`eg-cell ${cellCls(ri, ci)} ${col.align === 'center' ? 'eg-center' : ''}`}
                      key={col.field}
                      style={{ width: col.width, minWidth: col.width, maxWidth: col.width }}
                      onMouseDown={(e) => handleCellMouseDown(ri, ci, e)}
                      onMouseEnter={() => handleCellMouseEnter(ri, ci)}
                      onTouchStart={() => {
                        if (editingCell && !(editingCell.row === ri && editingCell.col === ci)) commitEdit();
                        setActiveCell({ row: ri, col: ci });
                        setSelectionRange({ startRow: ri, startCol: ci, endRow: ri, endCol: ci });
                      }}
                      onDoubleClick={() => startEditing(ri, ci)}
                    >
                      {isEditing ? (
                        <input
                          ref={inputRef}
                          className="eg-edit-input"
                          value={editValue}
                          autoFocus
                          onChange={(e) => setEditValue(e.target.value)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.stopPropagation(); commitEdit();
                              setActiveCell(p => {
                                if (!p) return p;
                                const nr = Math.min(rowsRef.current.length - 1, p.row + 1);
                                setSelectionRange({ startRow: nr, startCol: p.col, endRow: nr, endCol: p.col });
                                return { ...p, row: nr };
                              });
                            } else if (e.key === 'Tab') {
                              e.stopPropagation(); commitEdit();
                              setActiveCell(p => {
                                if (!p) return p;
                                const nc = e.shiftKey ? Math.max(0, p.col - 1) : Math.min(GRID_COLUMNS.length - 1, p.col + 1);
                                setSelectionRange({ startRow: p.row, startCol: nc, endRow: p.row, endCol: nc });
                                return { ...p, col: nc };
                              });
                            } else if (e.key === 'Escape') { e.stopPropagation(); cancelEdit(); }
                          }}
                          onBlur={commitEdit}
                        />
                      ) : (
                        <span className="eg-cell-text">{val}</span>
                      )}
                      {activeCell?.row === ri && activeCell?.col === ci && !isEditing && (
                        <div
                          className="eg-fill-handle"
                          onMouseDown={(e) => {
                            e.stopPropagation();
                            setFillDragging(true);
                            setFillEnd({ row: ri, col: ci });
                          }}
                        />
                      )}
                    </div>
                  );
                })}
              </div>
            );
          })}
          {visibleRows.length === 0 && (
            <div className="eg-empty" style={{ width: totalContentWidth }}>
              Heç bir məlumat yoxdur
            </div>
          )}
        </div>
      </div>
    </div>
  );
});

export default ExcelGrid;
