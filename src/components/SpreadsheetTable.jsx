import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import ContextMenu from './ContextMenu';
import FilterPanel from './FilterPanel';
import TrainingPlanModal from './TrainingPlanModal';
import ImportExcelModal from './ImportExcelModal';
import ExcelGrid from './ExcelGrid';
import { getUniqueCourseGroups, generateTrainingPlan } from '../lib/excelGenerator';
import { rowKey, FIELD_LABELS } from '../lib/importMapping';
import useFirebaseData from '../lib/useFirebaseData';
import logoUrl from '../assets/ist-logo.png?url';
import {
  SearchIcon, CloseIcon, ResetFilterIcon, ImportIcon, WarningIcon,
} from './Icons';

export default function SpreadsheetTable() {
  const {
    rows, loading, connected, syncError,
    canUndo, canRedo,
    updateCell, addRow, deleteRow, importRows, undo, redo,
  } = useFirebaseData();

  const [searchText, setSearchText] = useState('');
  const [columnFilters, setColumnFilters] = useState({});
  const [menuState, setMenuState] = useState(null);
  const [activeFilterColumn, setActiveFilterColumn] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [modalGroups, setModalGroups] = useState([]);
  const [filteredForTemplate, setFilteredForTemplate] = useState([]);
  const [importOpen, setImportOpen] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const gridRef = useRef(null);

  // Filter + search (unchanged logic, now backed by real-time rows).
  const filteredData = useMemo(() => {
    let data = rows;
    const af = Object.entries(columnFilters);
    if (af.length > 0) {
      data = data.filter(row =>
        af.every(([field, vals]) => {
          if (!vals || vals.length === 0) return true;
          return vals.some(v => String(row[field] || '').toLowerCase().includes(v.toLowerCase()));
        })
      );
    }
    if (searchText.trim()) {
      const q = searchText.toLowerCase().trim();
      data = data.filter(row =>
        row && typeof row === 'object' && Object.values(row).some(v => String(v || '').toLowerCase().includes(q))
      );
    }
    return data;
  }, [rows, columnFilters, searchText]);

  const getSynchronizedValues = useCallback((field) => {
    const af = Object.entries(columnFilters).filter(([f]) => f !== field);
    let data = rows;
    if (af.length > 0) {
      data = data.filter(row =>
        af.every(([f, vals]) => {
          if (!vals || vals.length === 0) return true;
          return vals.some(v => String(row[f] || '').toLowerCase().includes(v.toLowerCase()));
        })
      );
    }
    const vals = new Set();
    data.forEach(r => { const v = r[field]; if (v) vals.add(String(v)); });
    return Array.from(vals).sort();
  }, [rows, columnFilters]);

  // Clicking a column header opens the filter directly (no extra button).
  const handleHeaderClick = useCallback((field) => {
    if (!field) return;
    setActiveFilterColumn(field);
  }, []);

  // Map a filtered-row index back to its real index in the full dataset.
  const realIndex = useCallback((filteredIdx) => {
    const row = filteredData[filteredIdx];
    if (!row) return -1;
    return rows.indexOf(row);
  }, [filteredData, rows]);

  // ---- Context menu + row operations ----

  const handleRowContextMenu = useCallback((rowIndex, x, y) => {
    setMenuState({ x, y, rowIndex });
  }, []);

  const insertRow = useCallback((position) => {
    if (!menuState) return;
    const idx = realIndex(menuState.rowIndex);
    if (idx >= 0) addRow(position, idx);
    setMenuState(null);
  }, [menuState, realIndex, addRow]);

  const requestDelete = useCallback((rowIndex) => {
    const row = filteredData[rowIndex];
    const actualIdx = row ? rows.indexOf(row) : rowIndex;
    setConfirmDelete({ rowIndex: actualIdx, name: row?.fullName || '' });
  }, [filteredData, rows]);

  const confirmDeleteRow = useCallback(() => {
    if (confirmDelete == null) return;
    deleteRow(confirmDelete.rowIndex);
    setConfirmDelete(null);
  }, [confirmDelete, deleteRow]);

  // ---- Training plan (unchanged) ----

  const handleTrainingPlan = useCallback(() => {
    if (menuState == null) return;
    const groups = getUniqueCourseGroups(filteredData);
    if (groups.length === 0) { setMenuState(null); return; }
    setFilteredForTemplate(filteredData);
    setModalGroups(groups);
    setModalOpen(true);
    setMenuState(null);
  }, [menuState, filteredData]);

  const handleConfirmTrainingPlan = useCallback(async (entries) => {
    try {
      const logoResp = await fetch(logoUrl);
      const logoBuffer = await logoResp.arrayBuffer();
      await generateTrainingPlan(filteredForTemplate, entries, logoBuffer);
    } catch (err) {
      console.error('Training plan error:', err);
    }
    setModalOpen(false);
  }, [filteredForTemplate]);

  // ---- Cell edit -> auto-save to Firebase ----
  const handleCellEdit = useCallback((rowId, field, value) => {
    updateCell(rowId, field, value);
  }, [updateCell]);

  // ---- Import confirm: write new + changed rows to Firebase ----
  const handleImportConfirm = useCallback(({ added, updated }) => {
    if ((!added || added.length === 0) && (!updated || updated.length === 0)) {
      setImportOpen(false);
      return;
    }
    const existing = new Set(rows.map(r => rowKey(r)));
    const freshAdded = (added || []).filter(r => !existing.has(rowKey(r)));
    if (freshAdded.length === 0 && (!updated || updated.length === 0)) { setImportOpen(false); return; }
    importRows(freshAdded, updated || []);
    setImportOpen(false);
  }, [rows, importRows]);

  // ---- Keyboard shortcuts: Undo / Redo (Ctrl+Z / Ctrl+Y) ----
  useEffect(() => {
    const handler = (e) => {
      const target = e.target;
      const inInput = target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA');
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !inInput) {
        e.preventDefault();
        if (e.shiftKey) redo(); else undo();
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y' && !inInput) {
        e.preventDefault();
        redo();
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [undo, redo]);

  if (loading) {
    return (
      <div className="spreadsheet-root">
        <div className="loading-screen">
          <div className="loading-spinner" />
          <div className="loading-text">Məlumatlar yüklənir...</div>
          <div className="loading-sub">Firebase bağlantısı qurulur</div>
        </div>
      </div>
    );
  }

  return (
    <div className="spreadsheet-root">
      <div className="toolbar">
        <div className="toolbar-left">
          <span className="row-count">{filteredData.length} / {rows.length} sətir</span>
          <button className={`btn-control ${canUndo ? '' : 'disabled'}`} onClick={undo} disabled={!canUndo} title="Geri al (Ctrl+Z)">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="16" height="16" aria-hidden="true">
              <polyline points="1 4 1 10 7 10" />
              <path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10" />
            </svg>
          </button>
          <button className={`btn-control ${canRedo ? '' : 'disabled'}`} onClick={redo} disabled={!canRedo} title="İrəli get (Ctrl+Y)">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="16" height="16" aria-hidden="true">
              <polyline points="23 4 23 10 17 10" />
              <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10" />
            </svg>
          </button>
          <span className={`sync-status ${connected ? 'online' : 'offline'}`} title={syncError || (connected ? 'Canlı' : 'Kəsildi')}>
            <span className="sync-dot" />{connected ? 'Canlı' : 'Kəsildi'}
          </span>
        </div>

        <div className="toolbar-center">
          <div className="search-box">
            <span className="search-icon"><SearchIcon /></span>
            <input type="text" className="search-input" placeholder="Axtar..." value={searchText} onChange={(e) => setSearchText(e.target.value)} />
            {searchText && <button className="search-clear" onClick={() => setSearchText('')} aria-label="Axtarışı təmizlə"><CloseIcon /></button>}
          </div>
        </div>

        <div className="toolbar-right">
          <div className="control-group">
            <button
              className={`btn-control reset ${Object.keys(columnFilters).length > 0 ? 'active' : ''}`}
              onClick={() => setColumnFilters({})}
              disabled={Object.keys(columnFilters).length === 0}
              title="Filtirləri sıfırla"
            >
              <ResetFilterIcon />
            </button>
            <button className="btn-control import" onClick={() => setImportOpen(true)} title="Excel-dən yeni məlumat idxal et">
              <ImportIcon /> Import
            </button>
          </div>
        </div>
      </div>

      {Object.keys(columnFilters).length > 0 && (
        <div className="active-filters-bar">
          {Object.entries(columnFilters).map(([field, values]) => (
            <span className="filter-chip" key={field}>
              {FIELD_LABELS[field] || field}: {values.length}
              <button onClick={() => setColumnFilters(prev => { const n = { ...prev }; delete n[field]; return n; })} aria-label={`${field} filtrini sil`}>
                <CloseIcon />
              </button>
            </span>
          ))}
        </div>
      )}

      <ExcelGrid
        ref={gridRef}
        rows={filteredData}
        onCellEdit={handleCellEdit}
        onHeaderClick={handleHeaderClick}
        onRowContextMenu={handleRowContextMenu}
      />

      {menuState && (
        <ContextMenu
          x={menuState.x} y={menuState.y}
          onClose={() => setMenuState(null)}
          onInsertAbove={() => insertRow('above')}
          onInsertBelow={() => insertRow('below')}
          onDelete={() => requestDelete(menuState.rowIndex)}
          onTrainingPlan={handleTrainingPlan}
        />
      )}

      {activeFilterColumn && (
        <FilterPanel
          field={activeFilterColumn}
          headerName={FIELD_LABELS[activeFilterColumn] || activeFilterColumn}
          values={getSynchronizedValues(activeFilterColumn)}
          selected={columnFilters[activeFilterColumn] || []}
          onApply={(field, values) => {
            setColumnFilters(prev => {
              const next = { ...prev };
              if (values && values.length > 0) next[field] = values;
              else delete next[field];
              return next;
            });
            setActiveFilterColumn(null);
          }}
          onClose={() => setActiveFilterColumn(null)}
        />
      )}

      {confirmDelete && (
        <div className="modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div className="modal confirm-modal" onClick={(e) => e.stopPropagation()}>
            <div className="confirm-body">
              <div className="confirm-icon"><WarningIcon /></div>
              <div className="confirm-title">Sətir silinsin?</div>
              <div className="confirm-message">
                {confirmDelete.name
                  ? `"${confirmDelete.name}" məlumatı silinəcək.`
                  : 'Bu sətir tamamilə silinəcək.'} Bu əməliyyat geri qaytarıla bilməz.
              </div>
            </div>
            <div className="confirm-actions">
              <button className="btn btn-secondary" onClick={() => setConfirmDelete(null)}>Ləğv et</button>
              <button className="btn btn-danger" onClick={confirmDeleteRow}>Sil</button>
            </div>
          </div>
        </div>
      )}

      {modalOpen && (
        <TrainingPlanModal
          groups={modalGroups}
          onConfirm={handleConfirmTrainingPlan}
          onCancel={() => setModalOpen(false)}
        />
      )}

      {importOpen && (
        <ImportExcelModal
          existingRows={rows}
          onConfirm={handleImportConfirm}
          onCancel={() => setImportOpen(false)}
        />
      )}
    </div>
  );
}
