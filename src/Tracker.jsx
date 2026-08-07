import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { loadPending, enqueue, removeSynced, pendingCount } from "./store";
import Medidas, { parseMedidas, MED_SHEET, MED_NOTES_COL, MED_DATE_COL, niceDomain } from "./Medidas";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";

const C = { accent: "#C8F135", dark: "#0D0D0D", card: "#161616", muted: "#555", muted2: "#2A2A2A", text: "#E8E8E8", red: "#FF6B6B", orange: "#F59E0B", blue: "#60A5FA", green: "#4ADE80" };

const WEEK_OFFSET = 9;
const FIRST_WEEK_COL = 7;
const MAX_WEEKS = 40;
const OFF = { kg: 3, reps: 4, rir: 5, notes: 7 };
// Within week block: base+1=REPS_OBJ, base+2=RIR_OBJ (0-indexed from base)
const OFF_REPS_OBJ = 1;
const SERIE_RE = /^[1-5]ª SERIE$/i;
const SESSION_NAMES_MAP = {
  "empujes a": "Empujes A", "empuje a": "Empujes A",
  "tirón a": "Tirón A", "tiron a": "Tirón A",
  "pierna": "Pierna", "descanso activo": "Descanso Activo",
  "empujes b": "Empujes B", "empuje b": "Empujes B",
  "cadena posterior": "Cadena Posterior",
};

// ── REPS OBJ PARSER: recovers "6-8" from Excel date (format d-m) ─────────────
// Excel stores "6-8" typed with format d-m as a date: day=6, month=8
// We recover it as day + "-" + month using LOCAL date methods (not UTC)
// to avoid timezone off-by-one shifting the month.
function parseRepsObj(raw) {
  if (raw == null) return null;
  if (raw instanceof Date || (typeof raw === "object" && raw.getTime)) {
    const d = new Date(raw);
    return `${d.getDate()}-${d.getMonth() + 1}`;
  }
  if (typeof raw === "number" && raw > 1000) {
    // Excel serial → JS Date (local)
    const d = new Date(Math.round((raw - 25569) * 86400 * 1000));
    return `${d.getDate()}-${d.getMonth() + 1}`;
  }
  const s = String(raw).trim();
  return /^\d+-\d+$/.test(s) ? s : (s || null);
}

// ── WEEK DATA ─────────────────────────────────────────────────────────────────
// Reads real dates from row 2. For wi with text ("SEMANA I") extrapolates
// backwards from the first real date. Uses dates as-is (no Monday snapping)
// because some weeks start on Saturday in the Excel.
// todayWeekIdx: finds which wi's date range [date, nextDate) contains today.
function buildWeekData(grid, weekBaseCols) {
  const get = (r, c) => grid[r]?.[c] ?? null;

  // Step 1: read raw ms timestamps from row 2 for each wi
  const rawMs = weekBaseCols.map((base) => {
    const raw = get(1, base);
    if (raw == null) return null;
    let d = null;
    if (raw instanceof Date || (typeof raw === "object" && raw?.getTime)) d = new Date(raw);
    else if (typeof raw === "number" && raw > 40000) d = new Date(Math.round((raw - 25569) * 86400 * 1000));
    return (d && d.getFullYear() > 2000) ? d.getTime() : null;
  });

  // Step 2: find first real date as anchor for back-extrapolation
  let anchorMs = null, anchorWi = -1;
  for (let wi = 0; wi < rawMs.length; wi++) {
    if (rawMs[wi] != null) { anchorMs = rawMs[wi]; anchorWi = wi; break; }
  }

  // Step 3: build final ms array — real where available, extrapolated elsewhere
  const finalMs = weekBaseCols.map((_, wi) => {
    if (rawMs[wi] != null) return rawMs[wi];
    if (anchorMs == null) return null;
    return anchorMs + (wi - anchorWi) * 7 * 24 * 3600 * 1000;
  });

  // Step 4: build labels from ms (DD/MM format)
  const fmt = ms => {
    if (ms == null) return null;
    const d = new Date(ms);
    return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}`;
  };

  return weekBaseCols.map((_, wi) => ({
    label: fmt(finalMs[wi]) || `S${wi+1}`,
    ms: finalMs[wi],
  }));
}

// Find which wi contains today using range detection:
// wi contains today if finalMs[wi] <= today < finalMs[wi+1]
// This works correctly even when Excel dates are Saturdays or have gaps.
function mondayMs(ms) {
  const d = new Date(ms);
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0, 0);
  x.setDate(x.getDate() - ((x.getDay() + 6) % 7));   // Mon=0 ... Sun=6
  return x.getTime();
}

function findCurrentWeekIdx(weekData) {
  const todayMs = new Date().setHours(12, 0, 0, 0);
  const thisMonday = mondayMs(todayMs);

  const known = weekData
    .map(({ ms }, wi) => ({ wi, ms }))
    .filter(w => w.ms != null)
    .sort((a, b) => a.ms - b.ms);
  if (!known.length) return -1;

  // The active block is the one belonging to the CALENDAR week we're in.
  // Which weekday you train, or how many weeks are missing before, is
  // irrelevant — what matters is the Mon–Sun week containing today.
  const match = known.find(w => mondayMs(w.ms) === thisMonday);
  if (match) return match.wi;

  // No block for this calendar week: before the calendar → first;
  // otherwise the next upcoming one, never a past week.
  if (thisMonday < mondayMs(known[0].ms)) return known[0].wi;
  const upcoming = known.find(w => mondayMs(w.ms) > thisMonday);
  if (upcoming) return upcoming.wi;
  return known[known.length - 1].wi;
}

// ── PROGRESSION SUGGESTION ────────────────────────────────────────────────────
// If last week reps > hi of range → suggest +5% kg, rounded to gym increment
// Returns { kg, reps, reason } or null
// Find the most recent week at or before `from` that has real data for this set.
// Weeks can be empty because training wasn't recorded (holidays, gaps), so
// looking only at from-1 loses the reference entirely. Returns null if none.
function findLastRecorded(slots, from) {
  for (let wi = from; wi >= 0; wi--) {
    const s = slots[wi];
    if (s && (s.kg || s.reps)) return { slot: s, weekIdx: wi };
  }
  return null;
}

function suggestProgression(prevKg, prevReps, repsObjStr) {
  const kg = parseFloat(prevKg);
  const reps = parseFloat(prevReps);
  if (isNaN(kg) || isNaN(reps) || !repsObjStr) return null;
  const match = repsObjStr.match(/^(\d+)-(\d+)$/);
  if (!match) return null;
  const [, lo, hi] = match.map(Number);

  // Reps at or above range top → suggest weight increase (+5%)
  if (reps >= hi) {
    const raw = kg * 1.05;
    const inc = kg >= 100 ? 5 : kg >= 40 ? 2.5 : 1.25;
    const newKg = Math.round(raw / inc) * inc;
    return { type: "weight", kg: newKg, reps: lo, reason: `${prevReps} reps supera el rango (${repsObjStr}) → sube peso` };
  }

  // Reps within range but not at top → suggest more reps
  if (reps >= lo && reps < hi) {
    const span = hi - lo;
    const gap = hi - reps;
    const increment = gap >= 2 ? (span >= 4 ? 1 : 2) : 1;
    const newReps = Math.min(reps + increment, hi);
    return { type: "reps", kg, reps: newReps, reason: `${prevReps} reps en rango (${repsObjStr}) → intenta ${newReps} reps` };
  }

  return null;
}

// ── LOAD RECOMMENDATION based on previous week reps vs objective range ─────────
function getLoadRec(prevReps, repsObjStr) {
  if (!prevReps || !repsObjStr) return null;
  const done = parseFloat(prevReps);
  if (isNaN(done)) return null;
  const match = repsObjStr.match(/^(\d+)-(\d+)$/);
  if (!match) return null;
  const [, lo, hi] = match.map(Number);
  if (done > hi) return { type: "up",   msg: `Hiciste ${done} reps (objetivo ${repsObjStr}) → sube peso` };
  if (done < lo) return { type: "down", msg: `Hiciste ${done} reps (objetivo ${repsObjStr}) → baja peso` };
  return { type: "ok",   msg: `${done} reps en rango (${repsObjStr}) ✓` };
}

// ── PR: best KG ever lifted on serie 1 ────────────────────────────────────────
function getPersonalRecord(exercise, weekLabels) {
  const serie1 = exercise.sets[0]; if (!serie1) return null;
  let best = null;
  serie1.slots.forEach((slot, wi) => {
    const kg = parseFloat(slot.kg);
    const reps = parseFloat(slot.reps);
    if (isNaN(kg) || isNaN(reps)) return;
    const label = weekLabels ? weekLabels[wi] : `S${wi + 1}`;
    if (!best || kg > best.kg || (kg === best.kg && reps > best.reps)) {
      best = { kg, reps, weekLabel: label };
    }
  });
  return best;
}

// ── STAGNATION ─────────────────────────────────────────────────────────────────
function detectStagnation(exercise, n = 3) {
  const serie1 = exercise.sets[0]; if (!serie1) return false;
  const filled = serie1.slots.filter(s => s.reps !== "" && !isNaN(parseFloat(s.reps)));
  if (filled.length < n + 1) return false;
  const base = filled[filled.length - n - 1];
  return filled.slice(-n).every(s =>
    (parseFloat(s.kg) || 0) <= (parseFloat(base.kg) || 0) &&
    (parseFloat(s.reps) || 0) <= (parseFloat(base.reps) || 0)
  );
}

// ── PARSER ─────────────────────────────────────────────────────────────────────
function parseSheet(ws) {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const grid = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell ? cell.v : null);
    }
    grid.push(row);
  }
  const get = (r, c) => grid[r]?.[c] ?? null;
  const str = v => v != null ? String(v).trim() : "";
  const weekBaseCols = Array.from({ length: MAX_WEEKS }, (_, i) => FIRST_WEEK_COL + i * WEEK_OFFSET);

  // Build week data: { label, ms } for each week from anchor date in row 2
  const weekData = buildWeekData(grid, weekBaseCols);
  const weekLabels = weekData.map(w => w.label);

  // Find which week index = today (by calendar)
  const todayWeekIdx = findCurrentWeekIdx(weekData);

  const sessions = {};
  let curSession = null;

  for (let r = 0; r < grid.length; r++) {
    const colB = str(get(r, 1));
    const colH = str(get(r, 7));
    const sk = Object.keys(SESSION_NAMES_MAP).find(k => colB.toLowerCase().includes(k));
    if (sk && colH.includes("CALENTAMIENTO")) {
      curSession = SESSION_NAMES_MAP[sk];
      if (!sessions[curSession]) sessions[curSession] = { exercises: [], todayWeekIdx };
      continue;
    }
    if (!curSession) continue;
    if (colB && /^1ª SERIE$/i.test(colH)) {
      const exName = colB.replace(/\n/g, " ").trim();
      const sets = [];
      let ri = r;
      while (ri < grid.length) {
        const rb = str(get(ri, 1));
        const rh = str(get(ri, 7));
        if (ri > r && rb !== "") break;
        if (ri > r && /^1ª SERIE$/i.test(rh)) break;
        if (!SERIE_RE.test(rh)) { ri++; if (sets.length > 0) break; continue; }
        const slots = weekBaseCols.map((base, wi) => ({
          weekIdx: wi, weekLabel: weekLabels[wi],
          kg: str(get(ri, base + OFF.kg)), reps: str(get(ri, base + OFF.reps)),
          rir: str(get(ri, base + OFF.rir)), notes: str(get(ri, base + OFF.notes)),
          colKg: base + OFF.kg, colReps: base + OFF.reps,
          colRir: base + OFF.rir, colNotes: base + OFF.notes, rowIdx: ri,
        }));
        const repsObjRaw = get(ri, FIRST_WEEK_COL + OFF_REPS_OBJ);
        const repsObj = parseRepsObj(repsObjRaw);
        sets.push({ label: rh, slots, repsObj, repsObjColIdx: FIRST_WEEK_COL + OFF_REPS_OBJ, repsObjRowIdx: ri });
        ri++;
      }
      if (sets.length > 0) {
        let lastFilled = -1;
        for (let wi = weekBaseCols.length - 1; wi >= 0; wi--) {
          // Check kg OR reps — some sessions (Descanso Activo) only have kg
          if (sets.some(s => {
            const vReps = s.slots[wi].reps;
            const vKg = s.slots[wi].kg;
            return (vReps !== "" && !isNaN(parseFloat(vReps))) ||
                   (vKg !== "" && !isNaN(parseFloat(vKg)) && parseFloat(vKg) > 0);
          })) { lastFilled = wi; break; }
        }
        // SEMANA ACTIVA: siempre la semana del lunes actual (de lun a dom).
        // No se adelanta nunca a la siguiente semana aunque ya esté rellena.
        // Si todayWeekIdx no se detecta (sin ancla), fallback a lastFilled+1.
        const nextWeek = todayWeekIdx >= 0 ? todayWeekIdx : lastFilled + 1;
        const ex = { name: exName, sets, lastFilledWeek: lastFilled, nextWeek, weekLabels };
        ex.isStagnant = detectStagnation(ex);
        ex.pr = getPersonalRecord(ex, weekLabels);
        sessions[curSession].exercises.push(ex);
      }
    }
  }
  // Expose the week calendar so callers can detect/repair date drift
  Object.values(sessions).forEach(sd => { sd.weekData = weekData; });
  return sessions;
}

// ── WRITER ─────────────────────────────────────────────────────────────────────
function ensureWeekColumn(ws, wi) {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const base = FIRST_WEEK_COL + wi * WEEK_OFFSET;
  const ha = XLSX.utils.encode_cell({ r: 1, c: base });
  if (!ws[ha] || !String(ws[ha]?.v || "").includes("SEMANA")) ws[ha] = { t: "s", v: `SEMANA ${wi + 1}` };
  ["SERIES", "REPS OBJ.", "RIR OBJ.", "KG", "REPS REALIZ.", "RIR REALIZ.", "PROGRESO", "ANOTACIONES"].forEach((h, i) => {
    const a = XLSX.utils.encode_cell({ r: 4, c: base + i });
    if (!ws[a]) ws[a] = { t: "s", v: h };
  });
  if (base + 7 > range.e.c) { range.e.c = base + 7; ws["!ref"] = XLSX.utils.encode_range(range); }
}

function writeSession(ws, sessionData, formSets, nextWeek, substitutions, editedRepsObj) {
  ensureWeekColumn(ws, nextWeek);
  sessionData.exercises.forEach(ex => {
    const fEx = formSets[ex.name]; if (!fEx) return;
    ex.sets.forEach((set, si) => {
      const f = fEx[si]; if (!f) return;
      const slot = set.slots[nextWeek]; if (!slot) return;
      const write = (col, val, t) => {
        if (val === "" || val == null) return;
        ws[XLSX.utils.encode_cell({ r: slot.rowIdx, c: col })] = { t, v: t === "n" ? parseFloat(val) : String(val) };
      };
      if (substitutions[ex.name]) ws[XLSX.utils.encode_cell({ r: slot.rowIdx, c: 1 })] = { t: "s", v: substitutions[ex.name] };
      // Write edited reps obj if changed
      const editedRo = editedRepsObj[ex.name]?.[si];
      if (editedRo != null) {
        ws[XLSX.utils.encode_cell({ r: set.repsObjRowIdx, c: set.repsObjColIdx })] = { t: "s", v: editedRo };
      }
      write(slot.colKg, f.kg, "n"); write(slot.colReps, f.reps, "n");
      write(slot.colRir, f.rir, "s"); write(slot.colNotes, f.notes, "s");
    });
  });
}

function buildChartData(exercise) {
  const s1 = exercise.sets[0]; if (!s1) return [];
  return s1.slots.reduce((acc, slot, wi) => {
    const reps = parseFloat(slot.reps);
    if (!slot.reps || isNaN(reps)) return acc;
    const kg = parseFloat(slot.kg);
    return [...acc, { week: slot.weekLabel || `S${wi+1}`, kg: isNaN(kg) ? null : kg, reps, wi }];
  }, []);
}

// ── APP ────────────────────────────────────────────────────────────────────────
export default function Tracker({ xlsxBuffer, fileName, onSave, onSignOut }) {
  const [screen, setScreen] = useState("home");
  const [state, setState] = useState(null);
  const [selSession, setSelSession] = useState(null);
  const [form, setForm] = useState({});
  const [openEx, setOpenEx] = useState(null);
  const [substitutions, setSubstitutions] = useState({});
  const [editedRepsObj, setEditedRepsObj] = useState({});
  const [newWeekCreated, setNewWeekCreated] = useState(false);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState(null);
  const [summary, setSummary] = useState(null);

  // Parse Excel when buffer changes (loaded from Drive)
  // Only reset to home on first load — background reloads preserve current screen
  const isFirstLoad = useRef(true);
  useEffect(() => {
    if (!xlsxBuffer) return;
    try {
      const buf = new Uint8Array(xlsxBuffer);
      // Check magic bytes — xlsx files start with PK (0x50 0x4B)
      // If it starts with < it's an HTML error page from Google — ignore silently
      if (buf[0] === 0x3C) {
        // HTML response (token expired or rate limit) — skip background reload
        if (!isFirstLoad.current) return;
        throw new Error("Sesión expirada. Por favor reconecta con Google Drive.");
      }
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const wsName = wb.SheetNames.find(n => n.toUpperCase().includes("FPARK")) || wb.SheetNames.at(-1);
      const sessions = parseSheet(wb.Sheets[wsName]);
      const weekDates = sessions[Object.keys(sessions)[0]]?.weekData || null;
      // MEDIDAS lives in the same workbook — parse it too (null if absent)
      const medidas = parseMedidas(wb.Sheets[MED_SHEET]);
      // Overlay pending local values (not yet synced to Sheets) so the UI
      // always shows the latest data even if a previous sync didn't finish.
      const pendingItems = loadPending();
      if (pendingItems.length) {
        Object.values(sessions).forEach(sd => sd.exercises.forEach(ex => ex.sets.forEach(set => set.slots.forEach(slot => {
          pendingItems.forEach(p => {
            if (p.row !== slot.rowIdx) return;
            if (p.col === slot.colKg) slot.kg = String(p.value);
            else if (p.col === slot.colReps) slot.reps = String(p.value);
            else if (p.col === slot.colRir) slot.rir = String(p.value);
            else if (p.col === slot.colNotes) slot.notes = String(p.value);
          });
        }))));
      }
      if (pendingItems.length && medidas) {
        medidas.forEach(rec => {
          pendingItems.forEach(pi => {
            if (pi.sheetName !== MED_SHEET || pi.row !== rec.rowIdx) return;
            const f = rec.__fields || null;
            if (pi.col === MED_NOTES_COL) { rec.notes = String(pi.value); return; }
            const key = Object.keys(rec.values).find((k, i) => i + 2 === pi.col);
            if (key) { rec.values[key] = String(pi.value); rec.hasData = true; }
          });
        });
      }
      // Never overwrite state while mid-session (screen === "log") — would invalidate
      // all numPad.slot references and cause saves to go to wrong cells.
      // On first load screen is null so we always apply it.
      if (isFirstLoad.current) {
        setState(prev => ({ wb, wsName, sessions, medidas, weekDates }));
        setScreen("home");
        isFirstLoad.current = false;
      } else {
        // On subsequent loads (shouldn't happen now background reload is removed),
        // only update if not in an active session
        setState(prev => {
          if (prev.__screen === "log") return prev; // guard — never happens now
          return { wb, wsName, sessions, medidas, weekDates };
        });
      }
    } catch (err) {
      if (!isFirstLoad.current) return; // Ignore errors on background reloads
      alert("Error leyendo Excel: " + err.message);
    }
  }, [xlsxBuffer]);

  const openSession = (name) => {
    const sd = state.sessions[name];
    const nextWeek = sd.exercises[0]?.nextWeek ?? 0;
    const sets = {};
    // Pre-fill from local state (already updated by autoSaveCell without reload)
    sd.exercises.forEach(ex => {
      sets[ex.name] = ex.sets.map(set => {
        const slot = set.slots[nextWeek];
        return {
          kg: slot?.kg ?? "",
          reps: slot?.reps ?? "",
          rir: slot?.rir ?? "",
          notes: slot?.notes ?? "",
        };
      });
    });
    setSelSession(name); setForm(sets); setSubstitutions({}); setEditedRepsObj({});
    setOpenEx(sd.exercises[0]?.name || null);
    setScreen("log");
  };

  // Cells pending sync — key: "exName-setIdx-field" (derived from queue meta)
  const [failedCells, setFailedCells] = useState({});
  // Numpad for the medidas screen — Log has its own, and its state is not in scope here
  const [medPad, setMedPad] = useState(null);
  const [pendingN, setPendingN] = useState(pendingCount());
  const flushing = useRef(false);

  // Rebuild failedCells map from the persistent queue
  const refreshPendingUI = () => {
    const items = loadPending();
    setPendingN(items.length);
    const fc = {};
    items.forEach(it => { if (it.meta) fc[it.meta] = true; });
    setFailedCells(fc);
  };

  // Push all pending writes to Sheets in one batch per sheet.
  // Called after every save, every 30s, on reconnect, and on app focus.
  const flushQueue = async () => {
    if (flushing.current) return;
    const items = loadPending();
    if (!items.length) { refreshPendingUI(); return; }
    flushing.current = true;
    setSaving(true);
    try {
      const bySheet = {};
      items.forEach(it => { (bySheet[it.sheetName] = bySheet[it.sheetName] || []).push(it); });
      for (const sheet of Object.keys(bySheet)) {
        const its = bySheet[sheet];
        await onSave(sheet, its.map(it => ({ row: it.row, col: it.col, value: it.value })));
        removeSynced(its.map(it => it.id));
      }
      setSaveError(null);
    } catch (e) {
      // Data is safe in localStorage — will retry automatically
      setSaveError(null);
    } finally {
      flushing.current = false;
      setSaving(false);
      refreshPendingUI();
      // Items enqueued while this flush was running would otherwise wait
      // for the 30s timer — send them now.
      if (loadPending().length) setTimeout(() => flushRef.current(), 0);
    }
  };
  const flushRef = useRef(flushQueue);
  flushRef.current = flushQueue;

  // Auto-sync: every 30s + when connection returns + when app regains focus
  useEffect(() => {
    const t = setInterval(() => flushRef.current(), 30000);
    const onOnline = () => flushRef.current();
    window.addEventListener('online', onOnline);
    window.addEventListener('focus', onOnline);
    return () => { clearInterval(t); window.removeEventListener('online', onOnline); window.removeEventListener('focus', onOnline); };
  }, []);

  // Save one measurement cell. Same queue, same guarantees as training data.
  const saveMedida = (rec, field, value) => {
    const num = parseFloat(String(value).replace(",", "."));
    if (isNaN(num)) return;
    setState(prev => {
      if (!prev.medidas) return prev;
      const medidas = prev.medidas.map(r => r.rowIdx === rec.rowIdx
        ? { ...r, values: { ...r.values, [field.key]: String(value) }, hasData: true }
        : r);
      return { ...prev, medidas };
    });
    enqueue({
      id: `M${rec.rowIdx}-${field.col}`,
      sheetName: MED_SHEET,
      row: rec.rowIdx, col: field.col, value: num,
      meta: `medida-${rec.week}-${field.key}`,
      ts: Date.now(),
    });
    // First value of a week with no date yet → stamp it so the row is identifiable.
    // Must be enqueued BEFORE flushing, or it misses this batch.
    if (!rec.date) {
      const d = new Date();
      const monday = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));
      stampDate(rec, monday);
    }
    refreshPendingUI();
    flushQueue();
  };

  // Queue a date write without flushing (caller decides when to flush)
  const stampDate = (rec, date) => {
    if (!date) return;
    const txt = `${String(date.getDate()).padStart(2,"0")}/${String(date.getMonth()+1).padStart(2,"0")}/${date.getFullYear()}`;
    setState(prev => {
      if (!prev.medidas) return prev;
      const medidas = prev.medidas.map(r =>
        r.rowIdx === rec.rowIdx ? { ...r, date: new Date(date) } : r);
      return { ...prev, medidas };
    });
    enqueue({
      id: `M${rec.rowIdx}-${MED_DATE_COL}`,
      sheetName: MED_SHEET,
      row: rec.rowIdx, col: MED_DATE_COL, value: txt,
      meta: `medida-${rec.week}-fecha`,
      ts: Date.now(),
    });
  };

  // The workbook's future weeks can start on the wrong Monday (e.g. the season
  // was planned to resume later than it actually did). Rewrite the dates of
  // every week with no data so the first one is THIS calendar week, keeping
  // 7-day spacing. Weeks that already hold data are never touched.
  const calendarDrift = (() => {
    const wds = state?.weekDates;
    if (!wds || !wds.length) return null;
    const today = new Date().setHours(12, 0, 0, 0);
    const thisMonday = mondayMs(today);
    const hasThisWeek = wds.some(w => w.ms != null && mondayMs(w.ms) === thisMonday);
    if (hasThisWeek) return null;
    // Only offer this if the current active block is a FUTURE empty one
    const firstEmpty = wds.findIndex((w, wi) => w.ms != null && mondayMs(w.ms) > thisMonday);
    if (firstEmpty < 0) return null;
    return { firstEmpty, thisMonday, current: wds[firstEmpty].ms };
  })();

  const realignCalendar = () => {
    if (!calendarDrift) return;
    const { firstEmpty, thisMonday } = calendarDrift;
    const cells = [];
    for (let wi = firstEmpty; wi < state.weekDates.length; wi++) {
      if (state.weekDates[wi]?.ms == null) continue;
      const d = new Date(thisMonday);
      d.setDate(d.getDate() + (wi - firstEmpty) * 7);
      const txt = `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
      cells.push({ row: 1, col: FIRST_WEEK_COL + wi * WEEK_OFFSET, value: txt });
    }
    if (!cells.length) return;
    cells.forEach(c => enqueue({
      id: `W1-${c.col}`, sheetName: state.wsName, row: c.row, col: c.col,
      value: c.value, meta: `cal-${c.col}`, ts: Date.now(),
    }));
    refreshPendingUI();
    flushQueue();
    alert("Calendario ajustado. Recarga la app para ver la semana correcta.");
  };

  // Fill in the dates of every recorded week that has none, extrapolating
  // 7 days per week from the first week that does. These are ESTIMATES: they
  // assume the weeks ran consecutively with no skipped weeks.
  const backfillDates = () => {
    const rows = state.medidas;
    if (!rows) return;
    const anchorIdx = rows.findIndex(r => r.date);
    if (anchorIdx < 0) return;
    const anchor = rows[anchorIdx].date;
    let n = 0;
    rows.forEach((r, i) => {
      if (!r.hasData || r.date) return;
      const d = new Date(anchor);
      d.setDate(d.getDate() + (i - anchorIdx) * 7);
      stampDate(r, d);
      n++;
    });
    if (n) { refreshPendingUI(); flushQueue(); }
  };

  // Write the week's date into the FECHA column (same local-first queue)
  const saveMedidaDate = (rec, date) => {
    stampDate(rec, date);
    refreshPendingUI();
    flushQueue();
  };

  // LOCAL-FIRST SAVE: value goes to localStorage instantly (infallible),
  // then the queue syncs to Sheets in background. Data can never be lost.
  const autoSaveCell = (exName, setIdx, field, value, slot) => {
    const { wsName } = state;
    const fieldToCol = { kg: slot.colKg, reps: slot.colReps, rir: slot.colRir, notes: slot.colNotes };
    const col = fieldToCol[field];
    if (col == null) return;
    const numericFields = ['kg', 'reps'];
    const cellValue = numericFields.includes(field) ? parseFloat(value) : value;
    if (isNaN(cellValue) && numericFields.includes(field)) return;

    // 1. Update UI state immediately
    setState(prev => {
      const newState = { ...prev };
      const sd = newState.sessions[selSession];
      const ex = sd.exercises.find(e => e.name === exName);
      if (ex) {
        const nextWeek = sd.exercises[0]?.nextWeek ?? 0;
        ex.sets[setIdx].slots[nextWeek] = {
          ...ex.sets[setIdx].slots[nextWeek],
          [field]: value,
        };
      }
      return newState;
    });

    // 2. Persist locally (instant, works offline)
    enqueue({
      id: `${slot.rowIdx}-${col}`,
      sheetName: wsName,
      row: slot.rowIdx, col, value: cellValue,
      meta: `${exName}-${setIdx}-${field}`,
      ts: Date.now(),
    });
    refreshPendingUI();

    // 3. Sync to Sheets in background (fire and forget)
    flushQueue();
  };

  // Called when user taps "Terminar sesión".
  // Data is already saved cell by cell — just build summary, create next week headers, go to Done.
  const finish = async () => {
    // Final sync — make sure the Excel is complete before showing summary.
    // If it fails, data stays safe in localStorage and syncs on next open.
    await flushQueue();
    const { sessions, wsName } = state;
    const sd = sessions[selSession];
    const nextWeek = sd.exercises[0]?.nextWeek ?? 0;
    const prevWeek = nextWeek - 1;
    const summaryItems = sd.exercises.map(ex => {
      // Use first set that has data (not necessarily set[0])
      const filledCur = ex.sets.map(s => s.slots[nextWeek]).find(s => s?.kg || s?.reps) || ex.sets[0]?.slots[nextWeek];
      const filledPrev = prevWeek >= 0
        ? (ex.sets.map(s => findLastRecorded(s.slots, prevWeek)?.slot).find(s => s?.kg || s?.reps)
           || findLastRecorded(ex.sets[0]?.slots || [], prevWeek)?.slot || null)
        : null;
      const cur = filledCur;
      const prev = filledPrev;
      const curKg = parseFloat(cur?.kg);
      const prevKg = parseFloat(prev?.kg);
      const curReps = parseFloat(cur?.reps);
      const prevReps = parseFloat(prev?.reps);
      let trend = "neutral";
      if (!isNaN(curKg) && !isNaN(prevKg)) {
        if (curKg > prevKg) trend = "up";
        else if (curKg < prevKg) trend = "down";
        else if (!isNaN(curReps) && !isNaN(prevReps)) {
          if (curReps > prevReps) trend = "up";
          else if (curReps < prevReps) trend = "down";
        }
      }
      return { name: ex.name, curKg, curReps, prevKg, prevReps, trend };
    }).filter(s => !isNaN(s.curKg) || !isNaN(s.curReps));
    setSummary(summaryItems);

    // Create headers for nextWeek+1 if they don't exist yet.
    // We check by seeing if the slot for nextWeek+1 already has a colKg pointing
    // to a column that exists — simplest proxy: does any exercise have a non-empty
    // weekLabel for nextWeek+1? If weekData has no date there, the column is missing.
    const futureWi = nextWeek + 1;
    const futureBase = FIRST_WEEK_COL + futureWi * WEEK_OFFSET;
    // Check if futureWi week label exists (means column was already created)
    const futureLabel = sd.exercises[0]?.weekLabels?.[futureWi];
    const alreadyExists = futureLabel && !futureLabel.startsWith("S"); // real date label vs fallback "S16"

    if (!alreadyExists) {
      try {
        // Calculate the Monday date for futureWi based on the anchor week
        // Find a known week to extrapolate from
        const knownWi = sd.exercises[0]?.sets[0]?.slots?.findIndex?.((_, i) => {
          const lbl = sd.exercises[0]?.weekLabels?.[i];
          return lbl && !lbl.startsWith("S");
        }) ?? -1;

        // Build header cells for the new week column
        // Row 1 (idx 1): "SEMANA X" label
        // Row 4 (idx 4): column headers
        const HEADERS = ["SERIES", "REPS OBJ.", "RIR OBJ.", "KG", "REPS REALIZ.", "RIR REALIZ.", "PROGRESO", "ANOTACIONES"];
        const cells = [
          { row: 1, col: futureBase, value: `SEMANA ${futureWi + 1}` },
          ...HEADERS.map((h, i) => ({ row: 4, col: futureBase + i, value: h })),
        ];

        // Also write the Monday date for this week in row 1, col futureBase+1
        // Extrapolate: each week = 7 days apart from any known anchor
        const anchorWi = sd.exercises[0]?.sets[0]?.slots?.[nextWeek]?.weekIdx ?? nextWeek;
        // Use the weekLabel of nextWeek as anchor if it looks like a date (dd/mm)
        const anchorLabel = sd.exercises[0]?.weekLabels?.[nextWeek];
        if (anchorLabel && /\d{2}\/\d{2}/.test(anchorLabel)) {
          const [d, m] = anchorLabel.split("/").map(Number);
          const anchorDate = new Date(new Date().getFullYear(), m - 1, d);
          const futureDate = new Date(anchorDate.getTime() + 7 * 24 * 3600 * 1000);
          const futureStr = `${String(futureDate.getDate()).padStart(2,"0")}/${String(futureDate.getMonth()+1).padStart(2,"0")}`;
          cells.push({ row: 1, col: futureBase + 1, value: futureStr });
        }

        await onSave(wsName, cells);
        setNewWeekCreated(true);
      } catch (e) {
        // Non-critical — don't block navigation
        setNewWeekCreated(false);
      }
    } else {
      setNewWeekCreated(false);
    }

    setScreen("done");
  };

  if (!state) return null;

  if (screen === "home")      return <Home sessions={state.sessions} fileName={fileName} onSession={openSession} onProgress={() => setScreen("progress")} onMedidas={() => setScreen("medidas")} drift={calendarDrift} onRealign={realignCalendar} onSignOut={onSignOut} />;
  if (screen === "log")       return <Log session={selSession} sd={state.sessions[selSession]} form={form} openEx={openEx} setOpenEx={setOpenEx} substitutions={substitutions} setSubstitutions={setSubstitutions} editedRepsObj={editedRepsObj} setEditedRepsObj={setEditedRepsObj} onSet={(ex, si, f, v) => setForm(p => { const s = [...(p[ex] || [])]; s[si] = { ...s[si], [f]: v }; return { ...p, [ex]: s }; })} onAutoSave={autoSaveCell} onFinish={finish} saving={saving} saveError={saveError} failedCells={failedCells} pendingN={pendingN} onSync={flushQueue} onBack={() => setScreen("home")} />;
  if (screen === "medidas")   return (<>
    {medPad && (
      <NumPad value={medPad.value} label={medPad.label} hint={medPad.hint}
        onValue={v => { if (v !== "") saveMedida(medPad.rec, medPad.field, v); }}
        onClose={() => setMedPad(null)} />
    )}
    {state.medidas
    ? <Medidas rows={state.medidas} pendingN={pendingN} saving={saving} onSync={flushQueue} onSetDate={saveMedidaDate} onBackfillDates={backfillDates}
        onOpenPad={(rec, f) => setMedPad({ rec, field: f, value: rec.values[f.key] || "",
          label: `${f.label} · ${f.unit}`, hint: f.unit })}
        onBack={() => setScreen("home")} />
    : <div style={{ background: D.bg, minHeight: "100vh", color: D.text, fontFamily: D.font, padding: 40, textAlign: "center" }}>
        <div style={{ fontSize: 15, marginBottom: 10 }}>No se encontró la hoja MEDIDAS</div>
        <div style={{ fontSize: 12, color: D.muted, marginBottom: 24 }}>Este archivo no incluye seguimiento de medidas.</div>
        <button onClick={() => setScreen("home")} style={{ background: D.card2, border: `1px solid ${D.border}`, color: D.text, borderRadius: 12, padding: "12px 24px", fontSize: 14 }}>Volver</button>
      </div>}
  </>);
  if (screen === "progress")  return <Progress sessions={state.sessions} onBack={() => setScreen("home")} />;
  if (screen === "done")      return <Done fileName={fileName} newWeekCreated={newWeekCreated} summary={summary} onBack={() => setScreen("home")} />;
}

// ═══════════════════════════════════════════════════════════════════════════════
// DESIGN SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════
const D = {
  bg: "#080808", card: "#111", card2: "#181818", border: "#222",
  accent: "#C8F135", accentDim: "#1A2A05",
  text: "#F0F0F0", muted: "#666", muted2: "#333",
  red: "#FF5555", orange: "#F59E0B", blue: "#60A5FA", green: "#4ADE80",
  done: "#22C55E", doneDim: "#052010",
  font: "system-ui, -apple-system, sans-serif",
  mono: "'SF Mono', 'Fira Code', monospace",
}
const s = (base, over = {}) => ({ ...base, ...over })
const row = (over = {}) => s({ display: "flex", alignItems: "center" }, over)
const col = (over = {}) => s({ display: "flex", flexDirection: "column" }, over)
const card = (over = {}) => s({ background: D.card, borderRadius: 16, border: `1px solid ${D.border}` }, over)

// ── NUMERIC KEYPAD ─────────────────────────────────────────────────────────────
function NumPad({ value, onValue, onClose, label, hint }) {
  const [local, setLocal] = useState(value || "")
  const press = (k) => {
    if (k === "⌫") { setLocal(p => p.slice(0, -1)); return }
    if (k === "✓") { onValue(local); onClose(); return }
    if (k === "." && local.includes(".")) return
    setLocal(p => p + k)
  }
  const keys = ["7","8","9","4","5","6","1","2","3",".","0","⌫"]
  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 200, display: "flex", flexDirection: "column", justifyContent: "flex-end", background: "rgba(0,0,0,0.7)", backdropFilter: "blur(4px)" }}>
      <div style={{ background: "#0F0F0F", borderRadius: "24px 24px 0 0", padding: "0 0 env(safe-area-inset-bottom,16px)" }}>
        {/* Header */}
        <div style={{ padding: "16px 20px 8px", borderBottom: `1px solid ${D.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 11, color: D.muted, textTransform: "uppercase", letterSpacing: 2, fontFamily: D.mono }}>{label}</div>
            {hint && <div style={{ fontSize: 10, color: D.muted, marginTop: 2 }}>{hint}</div>}
          </div>
          <div style={{ fontSize: 40, fontWeight: 800, color: local ? D.accent : D.muted, fontFamily: D.mono, letterSpacing: -1, minWidth: 120, textAlign: "right" }}>
            {local || "—"}
          </div>
        </div>
        {/* Adjust buttons */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, padding: "12px 16px 0" }}>
          {[
            { label: "-5", delta: -5 }, { label: "-2.5", delta: -2.5 },
            { label: "+2.5", delta: 2.5 }, { label: "+5", delta: 5 },
          ].map(({ label: bl, delta }) => (
            <button key={bl} onClick={() => setLocal(p => {
              const v = parseFloat(p) || 0
              const n = Math.max(0, v + delta)
              return n % 1 === 0 ? String(n) : n.toFixed(1)
            })}
              style={{ background: D.card2, border: `1px solid ${D.border}`, borderRadius: 10, padding: "10px 0", fontSize: 13, fontWeight: 700, color: delta > 0 ? D.accent : D.muted, cursor: "pointer", fontFamily: D.mono }}>
              {bl}
            </button>
          ))}
        </div>
        {/* Keypad grid */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, padding: "8px 16px 8px" }}>
          {keys.map(k => (
            <button key={k} onClick={() => press(k)}
              style={{ background: k === "✓" ? D.accent : k === "⌫" ? "#1A1A1A" : D.card2, border: `1px solid ${k === "✓" ? D.accent : D.border}`, borderRadius: 12, padding: "18px 0", fontSize: k === "⌫" || k === "✓" ? 20 : 24, fontWeight: 700, color: k === "✓" ? D.bg : D.text, cursor: "pointer", fontFamily: D.mono }}>
              {k}
            </button>
          ))}
        </div>
        {/* Confirm big button */}
        <div style={{ padding: "0 16px 8px" }}>
          <button onClick={() => { onValue(local); onClose(); }}
            style={{ width: "100%", background: D.accent, border: "none", borderRadius: 14, padding: "18px 0", fontSize: 15, fontWeight: 800, color: D.bg, cursor: "pointer", letterSpacing: 1, fontFamily: D.font }}>
            CONFIRMAR
          </button>
        </div>
      </div>
    </div>
  )
}

// ── TIMER OVERLAY ──────────────────────────────────────────────────────────────
function TimerOverlay({ timer, onStop }) {
  if (!timer) return null
  const pct = timer.remaining / timer.total
  const r = 70, circ = 2 * Math.PI * r
  const done = timer.remaining === 0
  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 150, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", background: "rgba(0,0,0,0.85)", backdropFilter: "blur(8px)" }}>
      <div style={{ position: "relative", width: 180, height: 180, marginBottom: 32 }}>
        <svg width="180" height="180" style={{ transform: "rotate(-90deg)" }}>
          <circle cx="90" cy="90" r={r} fill="none" stroke={D.muted2} strokeWidth="8" />
          <circle cx="90" cy="90" r={r} fill="none" stroke={done ? D.accent : D.muted} strokeWidth="8"
            strokeDasharray={circ} strokeDashoffset={circ * (1 - pct)}
            strokeLinecap="round" style={{ transition: "stroke-dashoffset 1s linear" }} />
        </svg>
        <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
          <div style={{ fontSize: 48, fontWeight: 800, color: done ? D.accent : D.text, fontFamily: D.mono, letterSpacing: -2 }}>
            {done ? "✓" : `${Math.floor(timer.remaining / 60)}:${String(timer.remaining % 60).padStart(2,"0")}`}
          </div>
          <div style={{ fontSize: 11, color: D.muted, marginTop: 4 }}>{done ? "¡A por la siguiente!" : "descansando"}</div>
        </div>
      </div>
      <button onClick={onStop}
        style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 30, padding: "14px 40px", fontSize: 13, fontWeight: 700, color: D.muted, cursor: "pointer", fontFamily: D.font, letterSpacing: 1 }}>
        CERRAR
      </button>
    </div>
  )
}

// ── HOME ───────────────────────────────────────────────────────────────────────
function Home({ sessions, fileName, onSession, onProgress, onMedidas, drift, onRealign, onSignOut }) {
  const stagnantCount = Object.values(sessions).flatMap(s => s.exercises).filter(e => e.isStagnant).length
  const today = new Date()
  const dayName = today.toLocaleDateString("es-ES", { weekday: "long" })
  const dateStr = today.toLocaleDateString("es-ES", { day: "numeric", month: "long" })

  return (
    <div style={{ background: D.bg, minHeight: "100vh", color: D.text, fontFamily: D.font, paddingBottom: 40 }}>
      <div style={{ maxWidth: 480, margin: "0 auto", padding: "0 16px" }}>

        {/* Header */}
        <div style={{ padding: "env(safe-area-inset-top,24px) 0 24px" }}>
          <div style={row({ justifyContent: "space-between", alignItems: "flex-start" })}>
            <div>
              <div style={{ fontSize: 13, color: D.muted, marginBottom: 4, textTransform: "capitalize" }}>{dayName}, {dateStr}</div>
              <div style={{ fontSize: 32, fontWeight: 800, color: D.accent, letterSpacing: -1, lineHeight: 1 }}>ALEJANDRO</div>
              <div style={{ fontSize: 11, color: D.muted, marginTop: 6 }}>Mesociclo I · FPARK</div>
            </div>
            <button onClick={onSignOut}
              style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 20, padding: "8px 14px", fontSize: 12, color: D.muted, cursor: "pointer", fontFamily: D.font, marginTop: 4 }}>
              Salir
            </button>
          </div>
        </div>

        {/* Stagnation alert */}
        {stagnantCount > 0 && (
          <div style={{ background: "#180F00", border: `1px solid ${D.orange}30`, borderRadius: 14, padding: "14px 16px", marginBottom: 16, display: "flex", gap: 12, alignItems: "center" }}>
            <div style={{ fontSize: 24 }}>⚠️</div>
            <div>
              <div style={{ fontSize: 14, fontWeight: 700, color: D.orange }}>Estancamiento</div>
              <div style={{ fontSize: 12, color: D.muted, marginTop: 2 }}>{stagnantCount} ejercicio{stagnantCount > 1 ? "s" : ""} sin mejora en 3+ semanas</div>
            </div>
          </div>
        )}

        {/* Progress card */}
        <div onClick={onProgress}
          style={{ ...card({ padding: "18px 20px", marginBottom: 24, cursor: "pointer", background: D.accentDim, border: `1px solid ${D.accent}20` }), display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, color: D.accent }}>Ver progreso</div>
            <div style={{ fontSize: 12, color: D.muted, marginTop: 3 }}>Gráficas · PRs · evolución</div>
          </div>
          <div style={{ fontSize: 28, color: D.accent }}>→</div>
        </div>

        {/* Calendar drift warning */}
        {drift && (
          <div style={{ background: "#2a1f00", border: "1px solid #8a6d00", borderRadius: 14, padding: "14px 16px", marginBottom: 18 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#ffc933" }}>El calendario no coincide con esta semana</div>
            <div style={{ fontSize: 11, color: D.muted, marginTop: 5, lineHeight: 1.5 }}>
              El siguiente bloque del Excel empieza el {new Date(drift.current).toLocaleDateString("es-ES", {day:"2-digit",month:"2-digit"})},
              pero estamos en la semana del {new Date(drift.thisMonday).toLocaleDateString("es-ES", {day:"2-digit",month:"2-digit"})}.
              Si registras ahora, irá al bloque equivocado.
            </div>
            <button onClick={onRealign}
              style={{ marginTop: 10, background: "#ffc933", color: "#000", border: "none", borderRadius: 10, padding: "10px 16px", fontSize: 12, fontWeight: 800, cursor: "pointer" }}>
              Ajustar calendario a esta semana
            </button>
          </div>
        )}

        {/* Medidas card */}
        <div onClick={onMedidas}
          style={{ ...card({ padding: "18px 20px", marginBottom: 24, cursor: "pointer" }), display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 16, fontWeight: 700 }}>Medidas</div>
            <div style={{ fontSize: 12, color: D.muted, marginTop: 3 }}>Peso · perímetros · evolución</div>
          </div>
          <div style={{ fontSize: 28, color: D.muted }}>→</div>
        </div>

        {/* Sessions */}
        <div style={{ fontSize: 11, color: D.muted, letterSpacing: 2, textTransform: "uppercase", marginBottom: 12, fontFamily: D.mono }}>Esta semana</div>
        <div style={col({ gap: 10 })}>
          {Object.entries(sessions).map(([name, sd]) => {
            const nw = sd.exercises[0]?.nextWeek ?? 0
            const weekLabel = sd.exercises[0]?.weekLabels?.[nw] || `S${nw+1}`
            const stagnant = sd.exercises.filter(e => e.isStagnant).length
            const exWithData = sd.exercises.filter(ex => {
              const slot = ex.sets[0]?.slots[nw]
              return slot && (
                (slot.reps !== "" && !isNaN(parseFloat(slot.reps))) ||
                (slot.kg !== "" && !isNaN(parseFloat(slot.kg)) && parseFloat(slot.kg) > 0)
              )
            })
            const sessionDone = exWithData.length >= 2
            const sessionPartial = !sessionDone && exWithData.length >= 1
            const exCount = sd.exercises.length

            return (
              <div key={name} onClick={() => onSession(name)}
                style={{ ...card({ padding: "18px 20px", cursor: "pointer", border: `1px solid ${sessionDone ? D.done+"30" : D.border}`, background: sessionDone ? D.doneDim : D.card }), display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ flex: 1, marginRight: 12 }}>
                  <div style={row({ gap: 8, marginBottom: 6 })}>
                    <div style={{ fontSize: 18, fontWeight: 700, color: sessionDone ? D.done : D.text }}>{name}</div>
                    {sessionDone && <div style={{ background: D.done+"20", border: `1px solid ${D.done}40`, borderRadius: 20, padding: "2px 8px", fontSize: 10, color: D.done, fontWeight: 700 }}>✓ HECHO</div>}
                    {sessionPartial && <div style={{ background: D.orange+"15", border: `1px solid ${D.orange}30`, borderRadius: 20, padding: "2px 8px", fontSize: 10, color: D.orange }}>EN CURSO</div>}
                    {stagnant > 0 && <div style={{ background: D.orange+"15", borderRadius: 20, padding: "2px 8px", fontSize: 10, color: D.orange }}>⚠ {stagnant}</div>}
                  </div>
                  <div style={{ fontSize: 12, color: D.muted }}>
                    {exCount} ejercicios · semana del {weekLabel}
                  </div>
                </div>
                <div style={{ fontSize: 24, color: sessionDone ? D.done : D.accent }}>›</div>
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}

// ── LOG ────────────────────────────────────────────────────────────────────────
function Log({ session, sd, form, openEx, setOpenEx, substitutions, setSubstitutions, editedRepsObj, setEditedRepsObj, onSet, onAutoSave, onFinish, saving, saveError, failedCells, pendingN, onSync, onBack }) {
  const [subModal, setSubModal] = useState(null)
  const [numPad, setNumPad] = useState(null) // { exName, si, field, value, hint }
  const [timer, setTimer] = useState(null)
  const timerRef = useRef(null)

  const startTimer = (secs) => {
    if (timerRef.current) clearInterval(timerRef.current)
    setTimer({ total: secs, remaining: secs, running: true })
    timerRef.current = setInterval(() => {
      setTimer(t => {
        if (!t || t.remaining <= 1) {
          clearInterval(timerRef.current)
          if (navigator.vibrate) navigator.vibrate([300, 100, 300])
          return { ...t, remaining: 0, running: false }
        }
        return { ...t, remaining: t.remaining - 1 }
      })
    }, 1000)
  }

  const stopTimer = () => { if (timerRef.current) clearInterval(timerRef.current); setTimer(null) }

  const exercises = sd?.exercises || []
  const nextWeek = exercises[0]?.nextWeek ?? 0
  const prevWeek = nextWeek - 1

  const openSubModal = (ex) => setSubModal({
    exName: ex.name, newName: substitutions[ex.name] || "",
    seriesRepsObj: ex.sets.map((set, si) => editedRepsObj[ex.name]?.[si] ?? set.repsObj ?? ""),
  })

  const confirmSub = () => {
    if (subModal.newName.trim()) setSubstitutions(p => ({ ...p, [subModal.exName]: subModal.newName.trim() }))
    const newEdited = {}
    subModal.seriesRepsObj.forEach((val, si) => { newEdited[si] = val })
    setEditedRepsObj(p => ({ ...p, [subModal.exName]: newEdited }))
    setSubModal(null)
  }

  // Count filled exercises
  const filledCount = exercises.filter(ex => {
    const f = form[ex.name] || []
    return f.some(s => s.kg || s.reps)
  }).length

  return (
    <div style={{ background: D.bg, minHeight: "100vh", color: D.text, fontFamily: D.font }}>

      {/* Numpad overlay */}
      {numPad && (
        <NumPad
          value={numPad.value}
          label={numPad.label}
          hint={numPad.hint}
          onValue={v => {
            if (numPad.med) { if (v !== "") saveMedida(numPad.med.rec, numPad.med.field, v); return; }
            onSet(numPad.exName, numPad.si, numPad.field, v);
            if (v !== '' && numPad.slot) {
              onAutoSave(numPad.exName, numPad.si, numPad.field, v, numPad.slot);
            }
          }}
          onClose={() => setNumPad(null)}
        />
      )}

      {/* Timer overlay */}
      <TimerOverlay timer={timer} onStop={stopTimer} />

      {/* Sub modal */}
      {subModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.8)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
          <div style={{ ...card({ padding: 24, width: "100%", maxWidth: 380, maxHeight: "80vh", overflowY: "auto" }) }}>
            <div style={{ fontSize: 11, color: D.muted, letterSpacing: 2, marginBottom: 6, fontFamily: D.mono }}>EDITAR EJERCICIO</div>
            <div style={{ fontSize: 13, color: D.muted, marginBottom: 18, lineHeight: 1.4 }}>{subModal.exName}</div>
            <div style={{ fontSize: 11, color: D.muted, marginBottom: 8 }}>Nombre nuevo (opcional)</div>
            <input type="text" value={subModal.newName} onChange={e => setSubModal(p => ({ ...p, newName: e.target.value }))}
              placeholder="Dejar vacío para mantener" style={{ ...inp, marginBottom: 20 }} />
            <div style={{ fontSize: 11, color: D.muted, marginBottom: 12 }}>Reps objetivo por serie</div>
            {subModal.seriesRepsObj.map((val, si) => (
              <div key={si} style={row({ gap: 12, marginBottom: 10 })}>
                <div style={{ fontSize: 12, color: D.muted, width: 56 }}>Serie {si+1}</div>
                <input type="text" value={val} onChange={e => setSubModal(p => { const s = [...p.seriesRepsObj]; s[si] = e.target.value; return { ...p, seriesRepsObj: s } })}
                  placeholder="6-8" style={{ ...inp, width: 80, textAlign: "center", fontSize: 16, fontWeight: 700 }} />
                <div style={{ fontSize: 11, color: D.muted }}>reps</div>
              </div>
            ))}
            <div style={row({ gap: 10, marginTop: 20 })}>
              <button onClick={confirmSub} style={{ flex: 1, background: D.accent, color: D.bg, fontFamily: D.font, fontWeight: 800, fontSize: 13, padding: "14px 0", borderRadius: 12, border: "none", cursor: "pointer" }}>CONFIRMAR</button>
              <button onClick={() => setSubModal(null)} style={{ flex: 1, background: D.card2, border: `1px solid ${D.border}`, color: D.muted, fontFamily: D.font, fontSize: 13, padding: "14px 0", borderRadius: 12, cursor: "pointer" }}>CANCELAR</button>
            </div>
          </div>
        </div>
      )}

      <div style={{ maxWidth: 480, margin: "0 auto", padding: "0 16px 140px" }}>

        {/* Header */}
        <div style={{ padding: "env(safe-area-inset-top,20px) 0 20px", display: "flex", alignItems: "center", gap: 14 }}>
          <button onClick={onBack} style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 12, width: 40, height: 40, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, color: D.muted, cursor: "pointer" }}>‹</button>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, color: D.muted, textTransform: "uppercase", letterSpacing: 2, fontFamily: D.mono }}>Sesión de hoy</div>
            <div style={{ fontSize: 22, fontWeight: 800, color: D.accent, letterSpacing: -0.5 }}>{session}</div>
            <div style={{ fontSize: 11, color: D.muted, marginTop: 2 }}>
              Semana {exercises[0]?.weekLabels?.[nextWeek] || nextWeek+1}
              {prevWeek >= 0 ? ` · ref. ${exercises[0]?.weekLabels?.[prevWeek] || prevWeek+1}` : ""}
            </div>
          </div>
          {/* Sync status badge */}
          <div onClick={pendingN > 0 ? onSync : undefined}
            style={{ display: "inline-flex", alignItems: "center", gap: 6, marginTop: 8, padding: "5px 12px", borderRadius: 20,
              background: pendingN > 0 ? "#3a2a00" : "#0f2a12",
              border: `1px solid ${pendingN > 0 ? "#8a6d00" : "#1e5228"}`,
              cursor: pendingN > 0 ? "pointer" : "default", fontSize: 11, fontFamily: D.mono,
              color: pendingN > 0 ? "#ffc933" : "#5fd97a" }}>
            {saving ? "↻ sincronizando..." : pendingN > 0 ? `● ${pendingN} pendiente${pendingN > 1 ? "s" : ""} — tocar para sincronizar` : "✓ todo sincronizado"}
          </div>
          {filledCount > 0 && (
            <div style={{ background: D.accentDim, border: `1px solid ${D.accent}30`, borderRadius: 20, padding: "4px 12px", fontSize: 11, color: D.accent, fontWeight: 700 }}>
              {filledCount}/{exercises.length}
            </div>
          )}
        </div>

        {/* Exercises */}
        {exercises.map(ex => {
          const isOpen = openEx === ex.name
          const fEx = form[ex.name] || []
          const filled = fEx.filter(s => s.kg || s.reps).length
          const displayName = substitutions[ex.name] || ex.name
          const currentSlot = ex.sets[0]?.slots[nextWeek]
          // alreadyDone: visual indicator only — all sets have saved data AND form is untouched
          // Never blocks inputs — user can always fill in the form
          const allSetsHaveData = ex.sets.every(set => {
            const slot = set.slots[nextWeek]
            return slot && (
              (slot.reps !== "" && !isNaN(parseFloat(slot.reps))) ||
              (slot.kg !== "" && !isNaN(parseFloat(slot.kg)) && parseFloat(slot.kg) > 0)
            )
          })
          const formHasData = fEx.some(s => s.kg || s.reps)
          const alreadyDone = allSetsHaveData && !formHasData

          return (
            <div key={ex.name} style={{ ...card({ marginBottom: 10, overflow: "hidden", border: `1px solid ${alreadyDone ? D.done+"30" : filled > 0 ? D.accent+"25" : D.border}` }) }}>

              {/* Exercise header */}
              <div onClick={() => setOpenEx(isOpen ? null : ex.name)}
                style={{ padding: "16px 18px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                <div style={{ flex: 1, marginRight: 12 }}>
                  <div style={row({ gap: 8, flexWrap: "wrap", marginBottom: 4 })}>
                    <div style={{ fontSize: 15, fontWeight: 700, color: alreadyDone ? D.done : filled > 0 ? D.accent : D.text }}>
                      {displayName}
                    </div>
                    {alreadyDone && <div style={{ fontSize: 10, color: D.done, background: D.done+"15", borderRadius: 20, padding: "2px 8px" }}>✓ HECHO</div>}
                    {substitutions[ex.name] && <div style={{ fontSize: 10, color: D.orange, background: D.orange+"15", borderRadius: 20, padding: "2px 8px" }}>CAMBIADO</div>}
                    {ex.isStagnant && !alreadyDone && <div style={{ fontSize: 10, color: D.orange }}>⚠ estancado</div>}
                  </div>
                  <div style={{ fontSize: 12, color: D.muted }}>
                    {ex.sets.length} series
                    {filled > 0 && !alreadyDone ? ` · ${filled} registradas` : ""}
                    {ex.pr ? ` · PR ${ex.pr.kg}kg` : ""}
                  </div>
                </div>
                <div style={{ fontSize: 20, color: isOpen ? D.accent : D.muted, transform: isOpen ? "rotate(180deg)" : "none", transition: "transform .2s" }}>⌄</div>
              </div>

              {isOpen && (
                <div style={{ padding: "0 16px 16px", borderTop: `1px solid ${D.border}` }}>

                  {/* Already done */}
                  {alreadyDone && (
                    <div style={{ background: D.doneDim, border: `1px solid ${D.done}20`, borderRadius: 12, padding: "14px 16px", marginTop: 14, marginBottom: 4 }}>
                      <div style={{ fontSize: 14, fontWeight: 700, color: D.done, marginBottom: 8 }}>✓ Ya registrado esta semana</div>
                      <div style={row({ gap: 8, flexWrap: "wrap" })}>
                        {ex.sets.map((set, si) => {
                          const slot = set.slots[nextWeek]
                          if (!slot?.kg && !slot?.reps) return null
                          return (
                            <div key={si} style={{ background: "#0A1A0A", borderRadius: 10, padding: "10px 14px", textAlign: "center", minWidth: 70 }}>
                              <div style={{ fontSize: 10, color: D.muted, marginBottom: 4, fontFamily: D.mono }}>{set.label.replace(" SERIE","ª")}</div>
                              <div style={{ fontSize: 20, fontWeight: 800, color: D.done }}>{slot.kg || "—"}</div>
                              <div style={{ fontSize: 10, color: D.muted }}>kg</div>
                              {slot.reps && <div style={{ fontSize: 14, fontWeight: 700, color: D.text, marginTop: 2 }}>{slot.reps}<span style={{ fontSize: 10, color: D.muted }}>r</span></div>}
                            </div>
                          )
                        })}
                      </div>
                    </div>
                  )}

                  {/* Stagnation */}
                  {ex.isStagnant && !formHasData && (
                    <div style={{ background: "#180F00", border: `1px solid ${D.orange}25`, borderRadius: 10, padding: "10px 14px", marginTop: 14, marginBottom: 4, fontSize: 12, color: D.orange }}>
                      ⚠️ Sin mejora en 3+ semanas · considera ajustar carga
                    </div>
                  )}

                  {/* Input section — always shown */}
                  {ex.sets.map((set, si) => {
                    const f = fEx[si] || { kg: "", reps: "", rir: "", notes: "" }
                    const lastRec = prevWeek >= 0 ? findLastRecorded(set.slots, prevWeek) : null
                    const prev = lastRec?.slot || null
                    const refWeekIdx = lastRec?.weekIdx ?? -1
                    const weeksAgo = refWeekIdx >= 0 ? nextWeek - refWeekIdx : 0
                    const pKg = prev?.kg || null
                    const pReps = prev?.reps || null
                    const repsObj = editedRepsObj[ex.name]?.[si] ?? set.repsObj
                    const rec = getLoadRec(prev?.reps, repsObj)
                    const suggestion = suggestProgression(prev?.kg, prev?.reps, repsObj)

                    return (
                      <div key={si} style={{ marginTop: 16 }}>
                        {/* Serie label */}
                        <div style={row({ justifyContent: "space-between", marginBottom: 10 })}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: si === 0 ? D.accent : D.muted, fontFamily: D.mono }}>
                            {set.label}
                          </div>
                          {repsObj && <div style={{ fontSize: 11, color: D.muted }}>obj. {repsObj} reps</div>}
                        </div>

                        {/* Load rec */}
                        {rec && rec.type !== "ok" && (
                          <div style={{ background: rec.type === "up" ? "#081A08" : "#180808", border: `1px solid ${rec.type === "up" ? D.green : D.red}20`, borderRadius: 8, padding: "8px 12px", marginBottom: 10, fontSize: 12, color: rec.type === "up" ? D.green : D.red, fontWeight: 600 }}>
                            {rec.type === "up" ? "⬆" : "⬇"} {rec.msg}
                          </div>
                        )}

                        {/* Progression suggestion — visually dominant */}
                        {suggestion && (
                          <div style={{ background: suggestion.type === "weight" ? "#0D1A00" : "#00101A", border: `2px solid ${suggestion.type === "weight" ? D.accent : D.blue}40`, borderRadius: 14, padding: "16px", marginBottom: 14 }}>
                            <div style={{ fontSize: 11, fontWeight: 800, color: suggestion.type === "weight" ? D.accent : D.blue, letterSpacing: 2, marginBottom: 12, fontFamily: D.mono }}>
                              {suggestion.type === "weight" ? "⬆ SUBE PESO" : "⬆ SUBE REPS"}
                            </div>
                            <div style={row({ gap: 10, alignItems: "center", marginBottom: 10 })}>
                              <div style={{ flex: 1, background: "rgba(0,0,0,0.3)", borderRadius: 10, padding: "10px 14px", textAlign: "center" }}>
                                <div style={{ fontSize: 9, color: suggestion.type === "weight" ? D.accent : D.muted, letterSpacing: 2, marginBottom: 4, fontFamily: D.mono }}>KG</div>
                                <div style={{ fontSize: 32, fontWeight: 900, color: suggestion.type === "weight" ? D.accent : D.muted, fontFamily: D.mono, letterSpacing: -1 }}>{suggestion.kg}</div>
                              </div>
                              <div style={{ fontSize: 18, color: D.muted }}>×</div>
                              <div style={{ flex: 1, background: "rgba(0,0,0,0.3)", borderRadius: 10, padding: "10px 14px", textAlign: "center" }}>
                                <div style={{ fontSize: 9, color: suggestion.type === "reps" ? D.blue : D.muted, letterSpacing: 2, marginBottom: 4, fontFamily: D.mono }}>REPS</div>
                                <div style={{ fontSize: 32, fontWeight: 900, color: suggestion.type === "reps" ? D.blue : D.muted, fontFamily: D.mono, letterSpacing: -1 }}>{suggestion.reps}</div>
                              </div>
                              <button onClick={() => {
                                const kgStr = String(suggestion.kg);
                                const repsStr = String(suggestion.reps);
                                onSet(ex.name, si, "kg", kgStr);
                                onSet(ex.name, si, "reps", repsStr);
                                const slot = set.slots[nextWeek];
                                if (slot) {
                                  onAutoSave(ex.name, si, "kg", kgStr, slot);
                                  onAutoSave(ex.name, si, "reps", repsStr, slot);
                                }
                              }}
                                style={{ background: suggestion.type === "weight" ? D.accent : D.blue, color: D.bg, border: "none", borderRadius: 12, padding: "14px 18px", fontSize: 14, fontWeight: 900, cursor: "pointer", fontFamily: D.font, letterSpacing: 0.5 }}>
                                USAR
                              </button>
                            </div>
                            <div style={{ fontSize: 10, color: D.muted, fontFamily: D.mono }}>{suggestion.reason}</div>
                          </div>
                        )}

                        {/* Prev week reference — smaller, below suggestion */}
                        {refWeekIdx >= 0 && (pKg || pReps) && (
                          <div style={{ background: D.card2, borderRadius: 10, padding: "8px 14px", marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div style={{ fontSize: 10, color: D.muted, fontFamily: D.mono }}>
                              {weeksAgo <= 1 ? "SEMANA ANT." : `HACE ${weeksAgo} SEMANAS`} · {ex.weekLabels?.[refWeekIdx]}
                            </div>
                            <div style={row({ gap: 14 })}>
                              {pKg && <div><span style={{ fontSize: 16, fontWeight: 700, color: D.muted }}>{pKg}</span><span style={{ fontSize: 10, color: D.muted }}> kg</span></div>}
                              {pReps && <div><span style={{ fontSize: 16, fontWeight: 700, color: D.muted }}>{pReps}</span><span style={{ fontSize: 10, color: D.muted }}> r</span></div>}
                            </div>
                          </div>
                        )}

                        {/* BIG INPUT BUTTONS */}
                        <div style={row({ gap: 8, marginBottom: 8 })}>
                          {[
                            { field: "kg", label: "KG", unit: "kg" },
                            { field: "reps", label: "REPS", unit: "reps" },
                            { field: "rir", label: "RIR", unit: "rir" },
                          ].map(({ field, label, unit }) => (
                            <button key={field} onClick={() => setNumPad({
                              exName: ex.name, si, field,
                              value: f[field] || "",
                              label: `${ex.name.substring(0,20)} · ${set.label} · ${label}`,
                              hint: field === "kg" && pKg ? `Semana anterior: ${pKg}kg` : field === "reps" && pReps ? `Objetivo: ${repsObj || pReps} reps` : null,
                              slot: set.slots[nextWeek],
                            })}
                              style={{ flex: field === "rir" ? 0.7 : 1, background: f[field] ? (field === "kg" ? D.accentDim : D.card2) : D.card2, border: `1px solid ${failedCells[`${ex.name}-${si}-${field}`] ? D.red : f[field] ? (field === "kg" ? D.accent+"40" : D.border) : D.border}`, borderRadius: 12, padding: "16px 8px", textAlign: "center", cursor: "pointer" }}>
                              <div style={{ fontSize: 10, color: failedCells[`${ex.name}-${si}-${field}`] ? D.red : D.muted, marginBottom: 4, fontFamily: D.mono }}>{failedCells[`${ex.name}-${si}-${field}`] ? "⚠ "+label : label}</div>
                              <div style={{ fontSize: f[field] ? 28 : 20, fontWeight: 800, color: failedCells[`${ex.name}-${si}-${field}`] ? D.red : f[field] ? (field === "kg" ? D.accent : D.text) : D.muted, fontFamily: D.mono }}>
                                {f[field] || "—"}
                              </div>
                              {f[field] && <div style={{ fontSize: 10, color: failedCells[`${ex.name}-${si}-${field}`] ? D.red : D.muted, marginTop: 2 }}>{failedCells[`${ex.name}-${si}-${field}`] ? "no guardado" : unit}</div>}
                            </button>
                          ))}
                        </div>

                        {/* Notes */}
                        <input type="text" value={f.notes || ""} onChange={e => onSet(ex.name, si, "notes", e.target.value)}
                          onBlur={e => { const v = e.target.value; if (v) { const slot = set.slots[nextWeek]; if (slot) onAutoSave(ex.name, si, "notes", v, slot); } }}
                          placeholder="Nota..." style={{ ...inp, fontSize: 13, padding: "10px 14px" }} />

                        {/* Timer — shown after reps filled */}
                        {f.reps && (
                          <div style={row({ gap: 8, marginTop: 10 })}>
                            <div style={{ fontSize: 11, color: D.muted }}>Descanso:</div>
                            {[["1'30\"", 90], ["2'", 120], ["3'", 180]].map(([label, secs]) => (
                              <button key={secs} onClick={() => startTimer(secs)}
                                style={{ background: D.card2, border: `1px solid ${D.border}`, borderRadius: 8, padding: "6px 12px", fontSize: 12, color: D.muted, cursor: "pointer", fontFamily: D.mono }}>
                                {label}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    )
                  })}

                  {/* Edit button */}
                  <div style={{ marginTop: 16, paddingTop: 14, borderTop: `1px solid ${D.border}`, display: "flex", gap: 10 }}>
                    <button onClick={() => openSubModal(ex)}
                      style={{ background: D.card2, border: `1px solid ${D.border}`, borderRadius: 10, padding: "10px 16px", fontSize: 12, color: D.muted, cursor: "pointer", fontFamily: D.font }}>
                      ✏ Editar ejercicio
                    </button>
                    {(substitutions[ex.name] || editedRepsObj[ex.name]) && (
                      <button onClick={() => { setSubstitutions(p => { const n = { ...p }; delete n[ex.name]; return n }); setEditedRepsObj(p => { const n = { ...p }; delete n[ex.name]; return n }) }}
                        style={{ background: "none", border: "none", color: D.red, fontSize: 12, cursor: "pointer", fontFamily: D.font }}>
                        Resetear
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>
          )
        })}
      </div>

      {/* Save bar */}
      <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: `linear-gradient(transparent, ${D.bg} 40%)`, padding: "20px 16px env(safe-area-inset-bottom,24px)" }}>
        <div style={{ maxWidth: 480, margin: "0 auto" }}>
          {saveError && (
            <div style={{ background: "#1A0808", border: `1px solid ${D.red}30`, borderRadius: 10, padding: "10px 14px", marginBottom: 10, fontSize: 12, color: D.red }}>
              {saveError}
            </div>
          )}
          <div style={{ position: "relative" }}>
            {saving && (
              <div style={{ textAlign: "center", fontSize: 11, color: D.muted, marginBottom: 8 }}>
                ☁ Guardando...
              </div>
            )}
            <button onClick={onFinish}
              style={{ width: "100%", background: D.accent, color: D.bg, fontFamily: D.font, fontWeight: 800, fontSize: 16, padding: "18px 0", borderRadius: 16, border: "none", cursor: "pointer", letterSpacing: 0.5 }}>
              ✓ Terminar sesión
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

// ── PROGRESS ───────────────────────────────────────────────────────────────────
function Progress({ sessions, onBack }) {
  const [selSession, setSelSession] = useState(Object.keys(sessions)[0])
  const [selEx, setSelEx] = useState(null)
  const [metric, setMetric] = useState("kg")

  const sessionExercises = sessions[selSession]?.exercises.filter(ex => buildChartData(ex).length >= 2) || []
  const activeEx = selEx ? sessionExercises.find(e => e.name === selEx) || sessionExercises[0] : sessionExercises[0]
  const chartData = activeEx ? buildChartData(activeEx) : []
  const firstPt = chartData[0]; const lastPt = chartData.at(-1)
  const kgDiff = firstPt && lastPt ? (lastPt.kg - firstPt.kg).toFixed(1) : null
  const repsDiff = firstPt && lastPt ? (lastPt.reps - firstPt.reps).toFixed(0) : null

  return (
    <div style={{ background: D.bg, minHeight: "100vh", color: D.text, fontFamily: D.font, paddingBottom: 40 }}>
      <div style={{ maxWidth: 480, margin: "0 auto", padding: "0 16px" }}>

        {/* Header */}
        <div style={{ padding: "env(safe-area-inset-top,20px) 0 20px", display: "flex", alignItems: "center", gap: 14 }}>
          <button onClick={onBack} style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 12, width: 40, height: 40, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, color: D.muted, cursor: "pointer" }}>‹</button>
          <div>
            <div style={{ fontSize: 11, color: D.muted, textTransform: "uppercase", letterSpacing: 2, fontFamily: D.mono }}>Progreso</div>
            <div style={{ fontSize: 22, fontWeight: 800, color: D.accent }}>EVOLUCIÓN</div>
          </div>
        </div>

        {/* Session tabs */}
        <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
          {Object.keys(sessions).map(s => (
            <button key={s} onClick={() => { setSelSession(s); setSelEx(null) }}
              style={{ background: selSession === s ? D.accent : D.card, color: selSession === s ? D.bg : D.muted, border: `1px solid ${selSession === s ? D.accent : D.border}`, borderRadius: 20, fontSize: 12, fontWeight: selSession === s ? 700 : 400, padding: "7px 14px", cursor: "pointer", fontFamily: D.font }}>
              {s}
            </button>
          ))}
        </div>

        {sessionExercises.length === 0 ? (
          <div style={{ textAlign: "center", color: D.muted, fontSize: 14, padding: 60 }}>Sin datos suficientes (mínimo 2 semanas)</div>
        ) : (
          <>
            {/* Exercise chips */}
            <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
              {sessionExercises.map(ex => (
                <button key={ex.name} onClick={() => setSelEx(ex.name)}
                  style={{ background: activeEx?.name === ex.name ? D.accentDim : D.card, color: activeEx?.name === ex.name ? D.accent : ex.isStagnant ? D.orange : D.muted, border: `1px solid ${activeEx?.name === ex.name ? D.accent+"40" : ex.isStagnant ? D.orange+"30" : D.border}`, borderRadius: 20, fontSize: 11, padding: "6px 12px", cursor: "pointer", fontFamily: D.font, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {ex.isStagnant ? "⚠ " : ""}{ex.name}
                </button>
              ))}
            </div>

            {activeEx && chartData.length >= 2 && (
              <>
                {/* Stats */}
                <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
                  {[
                    { label: "KG actual", value: lastPt.kg, unit: "kg", delta: kgDiff, color: D.accent },
                    { label: "Reps actual", value: lastPt.reps, unit: "rep", delta: repsDiff, color: D.blue },
                    { label: "Semanas", value: chartData.length, unit: "", delta: null, color: D.text },
                  ].map(({ label, value, unit, delta, color }) => (
                    <div key={label} style={{ ...card({ flex: 1, padding: "14px 12px" }) }}>
                      <div style={{ fontSize: 10, color: D.muted, marginBottom: 6, fontFamily: D.mono }}>{label.toUpperCase()}</div>
                      <div style={{ fontSize: 24, fontWeight: 800, color }}>{value}<span style={{ fontSize: 11, color: D.muted, marginLeft: 2 }}>{unit}</span></div>
                      {delta != null && (
                        <div style={{ fontSize: 11, color: parseFloat(delta) >= 0 ? D.accent : D.red, marginTop: 4 }}>
                          {parseFloat(delta) >= 0 ? "▲" : "▼"} {Math.abs(delta)}{unit}
                        </div>
                      )}
                    </div>
                  ))}
                </div>

                {/* PR */}
                {activeEx.pr && (
                  <div style={{ ...card({ padding: "14px 18px", marginBottom: 16, background: "#060B1A", border: `1px solid ${D.blue}20` }), display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div>
                      <div style={{ fontSize: 10, color: D.blue, letterSpacing: 2, fontFamily: D.mono, marginBottom: 2 }}>★ MEJOR MARCA</div>
                      <div style={{ fontSize: 12, color: D.muted }}>{activeEx.name}</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontSize: 24, fontWeight: 800, color: D.text }}>{activeEx.pr.kg}<span style={{ fontSize: 12, color: D.muted }}> kg</span></div>
                      <div style={{ fontSize: 12, color: D.muted }}>× {activeEx.pr.reps} reps · {activeEx.pr.weekLabel}</div>
                    </div>
                  </div>
                )}

                {/* Stagnation */}
                {activeEx.isStagnant && (
                  <div style={{ background: "#180F00", border: `1px solid ${D.orange}25`, borderRadius: 12, padding: "12px 16px", marginBottom: 16, fontSize: 13, color: D.orange }}>
                    ⚠️ 3+ semanas sin mejora en la 1ª serie
                  </div>
                )}

                {/* Metric toggle */}
                <div style={row({ gap: 8, marginBottom: 14 })}>
                  {[["kg","KG"],["reps","Reps"],["both","Ambos"]].map(([v,l]) => (
                    <button key={v} onClick={() => setMetric(v)}
                      style={{ background: metric === v ? D.accent : D.card, color: metric === v ? D.bg : D.muted, border: `1px solid ${metric === v ? D.accent : D.border}`, borderRadius: 20, fontSize: 12, fontWeight: metric === v ? 700 : 400, padding: "6px 14px", cursor: "pointer", fontFamily: D.font }}>
                      {l}
                    </button>
                  ))}
                </div>

                {/* Chart */}
                <div style={{ ...card({ padding: "20px 8px 8px", marginBottom: 20 }) }}>
                  <div style={{ fontSize: 11, color: D.muted, paddingLeft: 12, marginBottom: 12, fontFamily: D.mono }}>{activeEx.name.toUpperCase()}</div>
                  <ResponsiveContainer width="100%" height={200}>
                    <LineChart data={chartData} margin={{ top: 4, right: 16, left: -16, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1A1A1A" />
                      <XAxis dataKey="week" tick={{ fill: D.muted, fontSize: 9 }} axisLine={false} tickLine={false} />
                      <YAxis domain={niceDomain(chartData.map(d => d.kg), { minPad: 1.25 })} allowDecimals tick={{ fill: D.muted, fontSize: 9 }} axisLine={false} tickLine={false} />
                      <Tooltip contentStyle={{ background: "#111", border: "1px solid #222", borderRadius: 10, fontSize: 12, fontFamily: D.mono }} labelStyle={{ color: D.accent }} />
                      {(metric === "kg" || metric === "both") && <Line type="monotone" dataKey="kg" name="KG" stroke={D.accent} strokeWidth={2.5} dot={{ r: 4, fill: D.accent }} activeDot={{ r: 6 }} />}
                      {(metric === "reps" || metric === "both") && <Line type="monotone" dataKey="reps" name="Reps" stroke={D.blue} strokeWidth={2.5} dot={{ r: 4, fill: D.blue }} activeDot={{ r: 6 }} />}
                    </LineChart>
                  </ResponsiveContainer>
                </div>

                {/* Last session */}
                <div style={{ fontSize: 11, color: D.muted, letterSpacing: 2, textTransform: "uppercase", marginBottom: 10, fontFamily: D.mono }}>Última sesión</div>
                {activeEx.sets.map((set, si) => {
                  const slot = activeEx.lastFilledWeek >= 0 ? set.slots[activeEx.lastFilledWeek] : null
                  if (!slot || (!slot.kg && !slot.reps)) return null
                  return (
                    <div key={si} style={{ ...card({ padding: "14px 18px", marginBottom: 8, border: `1px solid ${si === 0 ? D.accent+"25" : D.border}` }), display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 12, fontWeight: 700, color: si === 0 ? D.accent : D.muted }}>{set.label}{si === 0 ? " ★" : ""}</div>
                        {set.repsObj && <div style={{ fontSize: 10, color: D.muted, marginTop: 2 }}>obj. {set.repsObj} reps</div>}
                      </div>
                      <div style={row({ gap: 16 })}>
                        <div><span style={{ fontSize: 20, fontWeight: 800 }}>{slot.kg}</span><span style={{ fontSize: 11, color: D.muted }}> kg</span></div>
                        {slot.reps && <div><span style={{ fontSize: 20, fontWeight: 800 }}>{slot.reps}</span><span style={{ fontSize: 11, color: D.muted }}> rep</span></div>}
                        {slot.rir && <div style={{ fontSize: 11, color: D.muted }}>RIR {slot.rir}</div>}
                      </div>
                    </div>
                  )
                })}
              </>
            )}
          </>
        )}
      </div>
    </div>
  )
}

// ── DONE ───────────────────────────────────────────────────────────────────────
function Done({ fileName, newWeekCreated, summary, onBack }) {
  const ups = summary?.filter(s => s.trend === "up").length || 0
  const downs = summary?.filter(s => s.trend === "down").length || 0
  const neutrals = summary?.filter(s => s.trend === "neutral").length || 0

  return (
    <div style={{ background: D.bg, minHeight: "100vh", color: D.text, fontFamily: D.font, paddingBottom: 40 }}>
      <div style={{ maxWidth: 480, margin: "0 auto", padding: "0 20px" }}>

        <div style={{ paddingTop: 60, textAlign: "center", marginBottom: 32 }}>
          <div style={{ fontSize: 64, marginBottom: 16 }}>💪</div>
          <div style={{ fontSize: 28, fontWeight: 800, color: D.accent, marginBottom: 6 }}>¡Sesión guardada!</div>
          <div style={{ fontSize: 13, color: D.muted }}>Datos escritos en Google Sheets</div>
        </div>

        {summary && summary.length > 0 && (
          <>
            <div style={{ display: "flex", gap: 10, marginBottom: 24 }}>
              {[
                { label: "Subida", value: ups, color: D.accent, icon: "⬆" },
                { label: "Igual", value: neutrals, color: D.muted, icon: "=" },
                { label: "Bajada", value: downs, color: D.red, icon: "⬇" },
              ].map(({ label, value, color, icon }) => (
                <div key={label} style={{ ...card({ flex: 1, padding: "18px 10px" }), textAlign: "center" }}>
                  <div style={{ fontSize: 24, marginBottom: 6 }}>{icon}</div>
                  <div style={{ fontSize: 28, fontWeight: 800, color }}>{value}</div>
                  <div style={{ fontSize: 10, color: D.muted, marginTop: 4, letterSpacing: 1, fontFamily: D.mono }}>{label.toUpperCase()}</div>
                </div>
              ))}
            </div>

            <div style={{ fontSize: 11, color: D.muted, letterSpacing: 2, textTransform: "uppercase", marginBottom: 12, fontFamily: D.mono }}>Detalle</div>
            <div style={col({ gap: 8, marginBottom: 28 })}>
              {summary.map((s, i) => (
                <div key={i} style={{ ...card({ padding: "14px 18px", border: `1px solid ${s.trend === "up" ? D.accent+"25" : s.trend === "down" ? D.red+"25" : D.border}` }), display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ fontSize: 13, color: D.muted, flex: 1, marginRight: 12, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.name}</div>
                  <div style={row({ gap: 10, alignItems: "center" })}>
                    {!isNaN(s.prevKg) && <span style={{ fontSize: 12, color: D.muted }}>{s.prevKg}kg →</span>}
                    <span style={{ fontSize: 15, fontWeight: 800, color: s.trend === "up" ? D.accent : s.trend === "down" ? D.red : D.text }}>
                      {s.trend === "up" ? "↑" : s.trend === "down" ? "↓" : "="} {s.curKg}kg
                    </span>
                    {s.curReps && <span style={{ fontSize: 11, color: D.muted }}>× {s.curReps}</span>}
                  </div>
                </div>
              ))}
            </div>
          </>
        )}

        {newWeekCreated && (
          <div style={{ background: "#150F00", border: `1px solid ${D.orange}30`, borderRadius: 14, padding: "14px 18px", marginBottom: 24 }}>
            <div style={{ fontSize: 14, fontWeight: 700, color: D.orange, marginBottom: 4 }}>🆕 Nuevo bloque generado</div>
            <div style={{ fontSize: 12, color: D.muted, lineHeight: 1.7 }}>Se han creado las columnas del siguiente bloque automáticamente.</div>
          </div>
        )}

        <button onClick={onBack}
          style={{ width: "100%", background: D.card, border: `1px solid ${D.accent}30`, color: D.accent, fontFamily: D.font, fontWeight: 700, fontSize: 16, padding: "18px 0", borderRadius: 16, cursor: "pointer" }}>
          ← Volver al inicio
        </button>
      </div>
    </div>
  )
}

const inp = { background: D.card2, border: `1px solid ${D.border}`, borderRadius: 10, color: D.text, padding: "12px 14px", fontFamily: D.font, outline: "none", width: "100%", boxSizing: "border-box", fontSize: 14 }
const ghostBtn = { background: "none", border: `1px solid ${D.border}`, color: D.muted, fontSize: 12, padding: "8px 14px", borderRadius: 10, cursor: "pointer", fontFamily: D.font }
const hdr = { fontSize: 9, color: D.muted, textAlign: "center", letterSpacing: 2, fontFamily: D.mono }
