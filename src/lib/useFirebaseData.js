// Real-time data hook: subscribes to the Firebase records node, keeps the UI
// optimistically in sync, auto-saves every edit and heals offline edits when
// the connection comes back.
import { useState, useEffect, useRef, useCallback } from 'react';
import {
  recordsRef,
  singleRecordRef,
  onValue,
  set as fbSet,
  update as fbUpdate,
  remove as fbRemove,
  db,
  ref,
} from './firebase';
import registeredData from '../data/registrData.json';

function generateId() {
  return `rec-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

const EMPTY_ROW = {
  fullName: '', serial: '', idNumber: '', birthDate: '',
  phone: '', email: '', rank: '', fullNameId: '', rank2: '',
  courseCode: '', startDate: '', finishDate: '', note: '', date: '',
};

const MAX_HISTORY = 100;

// Firebase can hand back objects (rich text, text wrappers…) — flatten to
// plain display strings so cells always show clean text.
function sanitize(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v.richText && Array.isArray(v.richText)) return v.richText.map(rt => rt.text || '').join('');
    if (v.text !== undefined) return String(v.text);
    if (v instanceof Date && !isNaN(v.getTime())) {
      const d = String(v.getUTCDate()).padStart(2, '0');
      const m = String(v.getUTCMonth() + 1).padStart(2, '0');
      return `${d}.${m}.${v.getUTCFullYear()}`;
    }
    return '';
  }
  const s = String(v);
  return s.includes('[object Object]') ? '' : s;
}

function sanitizeRecords(val) {
  return Object.entries(val || {}).map(([id, data]) => {
    const clean = {};
    for (const [k, v] of Object.entries(data || {})) clean[k] = sanitize(v);
    return { _id: id, ...clean };
  });
}

function serializeRecord(row) {
  const { _id, ...rest } = row || {};
  return rest;
}

export default function useFirebaseData() {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [connected, setConnected] = useState(true);
  const [syncError, setSyncError] = useState('');

  const rowsRef = useRef(rows);
  rowsRef.current = rows;
  const historyRef = useRef([]);
  const futureRef = useRef([]);
  const skipNextRef = useRef(0);
  const dirtyRef = useRef(false);
  const seededRef = useRef(false);
  const mountedRef = useRef(true);

  useEffect(() => {
    mountedRef.current = true;
    return () => { mountedRef.current = false; };
  }, []);

  // Re-sync everything to Firebase (used by undo/redo and offline healing).
  const flushToFirebase = useCallback((nextRows) => {
    skipNextRef.current += 1;
    const data = {};
    nextRows.forEach(row => { data[row._id || generateId()] = serializeRecord(row); });
    return fbSet(recordsRef(), data)
      .then(() => { if (mountedRef.current) { dirtyRef.current = false; setSyncError(''); } })
      .catch(err => {
        console.error('Firebase write error:', err);
        dirtyRef.current = true;
        if (mountedRef.current) setSyncError('Bağlantı kəsildi — dəyişikliklər yenidən bağlandıqda sinxronlaşacaq.');
      });
  }, []);

  // Seed the database once with the built-in registry when it is empty.
  const seedOnce = useCallback(() => {
    seededRef.current = true;
    const seed = [
      ...registeredData.map((r, i) => ({ _id: `reg-${i}`, ...r })),
      ...Array.from({ length: 50 }, () => ({ _id: generateId(), ...EMPTY_ROW })),
    ];
    setRows(seed);
    const data = {};
    seed.forEach(row => { data[row._id] = serializeRecord(row); });
    skipNextRef.current += 1;
    fbUpdate(recordsRef(), data)
      .then(() => { dirtyRef.current = false; })
      .catch(err => console.error('Seed error:', err));
  }, []);

  // Live subscription — the table updates automatically for every session.
  useEffect(() => {
    const unsub = onValue(recordsRef(), (snap) => {
      if (!mountedRef.current) return;
      if (skipNextRef.current > 0) { skipNextRef.current -= 1; return; }
      if (dirtyRef.current) return; // keep local unsynced edits; flushed on reconnect
      const val = snap.val();
      if (val && typeof val === 'object' && Object.keys(val).length > 0) {
        setRows(sanitizeRecords(val));
      } else if (!seededRef.current) {
        seedOnce();
      } else {
        setRows([]);
      }
      setLoading(false);
    }, (err) => {
      console.error('Firebase read error:', err);
      if (mountedRef.current) {
        setLoading(false);
        setSyncError('Firebase oxunarkən xəta baş verdi: ' + (err?.message || err));
      }
    });
    return () => unsub();
  }, [seedOnce]);

  // Connection state + offline healing.
  useEffect(() => {
    const connRef = ref(db, '.info/connected');
    const unsub = onValue(connRef, (snap) => {
      if (!mountedRef.current) return;
      const isOnline = snap.val() === true;
      setConnected(isOnline);
      if (isOnline) {
        if (dirtyRef.current) {
          flushToFirebase(rowsRef.current);
        } else {
          setSyncError('');
        }
      }
    }, () => {});
    return () => unsub();
  }, [flushToFirebase]);

  const pushHistory = useCallback((before) => {
    historyRef.current = [...historyRef.current.slice(-MAX_HISTORY + 1), before];
    futureRef.current = [];
  }, []);

  // ---- Public operations (all auto-saved to Firebase) ----

  const updateCell = useCallback((rowId, field, value) => {
    setRows(prev => {
      const next = prev.map(r => (r._id === rowId ? { ...r, [field]: value } : r));
      pushHistory(prev);
      const row = next.find(r => r._id === rowId);
      if (row && rowId) {
        fbUpdate(singleRecordRef(rowId), serializeRecord(row))
          .then(() => {
            dirtyRef.current = false;
            if (mountedRef.current) setSyncError('');
          })
          .catch(err => {
            console.error('Cell save failed:', err);
            dirtyRef.current = true;
            if (mountedRef.current) setSyncError('Bağlantı kəsildi — dəyişikliklər yenidən bağlandıqda sinxronlaşacaq.');
          });
      }
      return next;
    });
  }, [pushHistory]);

  const addRow = useCallback((position, afterIndex) => {
    setRows(prev => {
      const newRow = { _id: generateId(), ...EMPTY_ROW };
      const next = [...prev];
      const idx = afterIndex !== undefined && afterIndex >= 0 ? afterIndex : prev.length - 1;
      if (position === 'above') next.splice(idx, 0, newRow);
      else next.splice(idx + 1, 0, newRow);
      pushHistory(prev);
      fbSet(singleRecordRef(newRow._id), EMPTY_ROW)
        .then(() => { dirtyRef.current = false; })
        .catch(err => {
          console.error('Row add failed:', err);
          dirtyRef.current = true;
          if (mountedRef.current) setSyncError('Bağlantı kəsildi — dəyişikliklər yenidən bağlandıqda sinxronlaşacaq.');
        });
      return next;
    });
  }, [pushHistory]);

  const deleteRow = useCallback((index) => {
    setRows(prev => {
      if (index < 0 || index >= prev.length) return prev;
      const target = prev[index];
      const next = prev.filter((_, i) => i !== index);
      pushHistory(prev);
      if (target?._id) {
        fbRemove(singleRecordRef(target._id))
          .then(() => { dirtyRef.current = false; })
          .catch(err => {
            console.error('Row delete failed:', err);
            dirtyRef.current = true;
            if (mountedRef.current) setSyncError('Bağlantı kəsildi — silinmə yenidən bağlandıqda sinxronlaşacaq.');
          });
      }
      return next;
    });
  }, [pushHistory]);

  // Import: `added` are brand-new records, `updated` replace existing ones by _id.
  const importRows = useCallback((added = [], updated = []) => {
    setRows(prev => {
      const existingById = new Map(prev.map(r => [r._id, r]));
      const next = [...prev];
      added.forEach(r => {
        const row = { _id: generateId(), ...r };
        next.push(row);
        existingById.set(row._id, row);
      });
      updated.forEach(r => {
        const idx = next.findIndex(x => x._id === r._id);
        if (idx >= 0) next[idx] = { ...next[idx], ...r };
        else next.push({ _id: r._id, ...r });
      });
      pushHistory(prev);
      const data = {};
      added.forEach(r => { const row = { _id: r._id || generateId(), ...r }; data[row._id] = serializeRecord(row); });
      updated.forEach(r => { if (r._id) data[r._id] = serializeRecord(r); });
      fbUpdate(recordsRef(), data)
        .then(() => { dirtyRef.current = false; if (mountedRef.current) setSyncError(''); })
        .catch(err => {
          console.error('Import write failed:', err);
          dirtyRef.current = true;
          if (mountedRef.current) setSyncError('Bağlantı kəsildi — idxal yenidən bağlandıqda sinxronlaşacaq.');
        });
      return next;
    });
  }, [pushHistory]);

  const undo = useCallback(() => {
    if (historyRef.current.length === 0) return;
    const snapshot = historyRef.current[historyRef.current.length - 1];
    historyRef.current = historyRef.current.slice(0, -1);
    futureRef.current = [rowsRef.current, ...futureRef.current];
    setRows(snapshot);
    flushToFirebase(snapshot);
  }, [flushToFirebase]);

  const redo = useCallback(() => {
    if (futureRef.current.length === 0) return;
    const snapshot = futureRef.current[0];
    futureRef.current = futureRef.current.slice(1);
    historyRef.current = [...historyRef.current, rowsRef.current];
    setRows(snapshot);
    flushToFirebase(snapshot);
  }, [flushToFirebase]);

  return {
    rows,
    loading,
    connected,
    syncError,
    canUndo: historyRef.current.length > 0,
    canRedo: futureRef.current.length > 0,
    updateCell,
    addRow,
    deleteRow,
    importRows,
    undo,
    redo,
  };
}
