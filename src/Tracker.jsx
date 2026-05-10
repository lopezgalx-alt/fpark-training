import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
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
function findCurrentWeekIdx(weekData) {
  const todayMs = new Date().setHours(12, 0, 0, 0); // midday to avoid DST edge cases

  // Find all wi with known ms, sorted
  const known = weekData
    .map(({ ms }, wi) => ({ wi, ms }))
    .filter(w => w.ms != null)
    .sort((a, b) => a.ms - b.ms);

  for (let i = 0; i < known.length; i++) {
    const curr = known[i];
    const next = known[i + 1];
    const from = curr.ms;
    const to = next ? next.ms : curr.ms + 7 * 24 * 3600 * 1000;
    if (todayMs >= from && todayMs < to) return curr.wi;
  }
  return -1;
}

// ── PROGRESSION SUGGESTION ────────────────────────────────────────────────────
// If last week reps > hi of range → suggest +5% kg, rounded to gym increment
// Returns { kg, reps, reason } or null
function suggestProgression(prevKg, prevReps, repsObjStr) {
  const kg = parseFloat(prevKg);
  const reps = parseFloat(prevReps);
  if (isNaN(kg) || isNaN(reps) || !repsObjStr) return null;
  const match = repsObjStr.match(/^(\d+)-(\d+)$/);
  if (!match) return null;
  const [, lo, hi] = match.map(Number);
  if (reps <= hi) return null; // still in range or below, no weight increase needed

  // +5% rounded to nearest gym increment
  const raw = kg * 1.05;
  const inc = kg >= 100 ? 5 : kg >= 40 ? 2.5 : 1.25;
  const newKg = Math.round(raw / inc) * inc;
  return { kg: newKg, reps: lo, reason: `${prevReps} reps con ${kg}kg supera el rango (${repsObjStr}) → +5%` };
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
  useEffect(() => {
    if (!xlsxBuffer) return;
    try {
      const buf = new Uint8Array(xlsxBuffer);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const wsName = wb.SheetNames.find(n => n.toUpperCase().includes("FPARK")) || wb.SheetNames.at(-1);
      const sessions = parseSheet(wb.Sheets[wsName]);
      setState({ wb, wsName, sessions });
      setScreen("home");
    } catch (err) {
      alert("Error leyendo Excel: " + err.message);
    }
  }, [xlsxBuffer]);

  const openSession = (name) => {
    const sd = state.sessions[name];
    const sets = {};
    sd.exercises.forEach(ex => { sets[ex.name] = ex.sets.map(() => ({ kg: "", reps: "", rir: "", notes: "" })); });
    setSelSession(name); setForm(sets); setSubstitutions({}); setEditedRepsObj({});
    setOpenEx(sd.exercises[0]?.name || null);
    setScreen("log");
  };

  const save = async () => {
    setSaving(true); setSaveError(null);
    try {
      const { wb, wsName, sessions } = state;
      const sd = sessions[selSession];
      const nextWeek = sd.exercises[0]?.nextWeek ?? 0;
      const prevWeek = nextWeek - 1;

      // Build summary
      const summaryItems = sd.exercises.map(ex => {
        const cur = ex.sets[0]?.slots[nextWeek];
        const prev = prevWeek >= 0 ? ex.sets[0]?.slots[prevWeek] : null;
        const curKg = parseFloat(form[ex.name]?.[0]?.kg || cur?.kg);
        const prevKg = parseFloat(prev?.kg);
        const curReps = parseFloat(form[ex.name]?.[0]?.reps || cur?.reps);
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
      }).filter(s => !isNaN(s.curKg));

      setSummary(summaryItems);
      setNewWeekCreated(nextWeek >= 15);

      // Collect ONLY the cells that need writing — no full file rewrite
      const cellUpdates = [];
      sd.exercises.forEach(ex => {
        const fEx = form[ex.name]; if (!fEx) return;
        ex.sets.forEach((set, si) => {
          const f = fEx[si]; if (!f) return;
          const slot = set.slots[nextWeek]; if (!slot) return;
          if (f.kg !== '' && f.kg != null)
            cellUpdates.push({ row: slot.rowIdx, col: slot.colKg, value: parseFloat(f.kg) });
          if (f.reps !== '' && f.reps != null)
            cellUpdates.push({ row: slot.rowIdx, col: slot.colReps, value: parseFloat(f.reps) });
          if (f.rir !== '' && f.rir != null)
            cellUpdates.push({ row: slot.rowIdx, col: slot.colRir, value: f.rir });
          if (f.notes !== '' && f.notes != null)
            cellUpdates.push({ row: slot.rowIdx, col: slot.colNotes, value: f.notes });
          // Substituted exercise name
          if (substitutions[ex.name])
            cellUpdates.push({ row: slot.rowIdx, col: 1, value: substitutions[ex.name] });
          // Edited reps objective
          const editedRo = editedRepsObj[ex.name]?.[si];
          if (editedRo != null)
            cellUpdates.push({ row: set.repsObjRowIdx, col: set.repsObjColIdx, value: editedRo });
        });
      });

      // Write only those cells via Sheets API — format is preserved
      await onSave(wsName, cellUpdates);
      setScreen("done");
    } catch (err) {
      setSaveError(err.message);
    } finally {
      setSaving(false);
    }
  };

  if (!state) return null;

  if (screen === "home")      return <Home sessions={state.sessions} fileName={fileName} onSession={openSession} onProgress={() => setScreen("progress")} onSignOut={onSignOut} />;
  if (screen === "log")       return <Log session={selSession} sd={state.sessions[selSession]} form={form} openEx={openEx} setOpenEx={setOpenEx} substitutions={substitutions} setSubstitutions={setSubstitutions} editedRepsObj={editedRepsObj} setEditedRepsObj={setEditedRepsObj} onSet={(ex, si, f, v) => setForm(p => { const s = [...(p[ex] || [])]; s[si] = { ...s[si], [f]: v }; return { ...p, [ex]: s }; })} onSave={save} saving={saving} saveError={saveError} onBack={() => setScreen("home")} />;
  if (screen === "progress")  return <Progress sessions={state.sessions} onBack={() => setScreen("home")} />;
  if (screen === "done")      return <Done fileName={fileName} newWeekCreated={newWeekCreated} summary={summary} onBack={() => setScreen("home")} />;
}

// ── UPLOAD ─────────────────────────────────────────────────────────────────────
function Upload({ onFile, fileRef }) {
  const [drag, setDrag] = useState(false);
  return (
    <div style={{ background: C.dark, minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 24, fontFamily: "'DM Mono','Courier New',monospace", color: C.text }}>
      <div style={{ fontSize: 10, color: C.muted, letterSpacing: 4, textTransform: "uppercase", marginBottom: 6 }}>Tracker de entreno</div>
      <div style={{ fontSize: 26, fontWeight: 700, color: C.accent, letterSpacing: -1, marginBottom: 2 }}>ALEJANDRO</div>
      <div style={{ fontSize: 10, color: C.muted, marginBottom: 44 }}>Mesociclo I · FPARK</div>
      <div onClick={() => fileRef.current.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); onFile(e.dataTransfer.files[0]); }}
        style={{ border: `2px dashed ${drag ? C.accent : C.muted2}`, borderRadius: 16, padding: "44px 36px", textAlign: "center", cursor: "pointer", background: drag ? "#1A1F0A" : C.card, maxWidth: 340, width: "100%", transition: "all .2s" }}>
        <div style={{ fontSize: 40, marginBottom: 14 }}>📊</div>
        <div style={{ fontSize: 13, color: C.text, fontWeight: 600, marginBottom: 6 }}>Sube tu Excel</div>
        <div style={{ fontSize: 10, color: C.muted }}>Toca o arrastra aquí</div>
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xlsm" onChange={e => onFile(e.target.files[0])} style={{ display: "none" }} />
    </div>
  );
}

// ── HOME ───────────────────────────────────────────────────────────────────────
function Home({ sessions, fileName, onSession, onProgress, onSignOut }) {
  const stagnantCount = Object.values(sessions).flatMap(s => s.exercises).filter(e => e.isStagnant).length;
  return (
    <div style={{ background: C.dark, minHeight: "100vh", color: C.text, fontFamily: "'DM Mono','Courier New',monospace", paddingBottom: 40 }}>
      <div style={{ maxWidth: 480, margin: "0 auto", padding: "0 16px" }}>
        <div style={{ padding: "28px 0 20px", display: "flex", justifyContent: "space-between", alignItems: "flex-end" }}>
          <div>
            <div style={{ fontSize: 10, color: C.muted, letterSpacing: 3, textTransform: "uppercase", marginBottom: 4 }}>Mesociclo I · FPARK</div>
            <div style={{ fontSize: 22, fontWeight: 700, color: C.accent, letterSpacing: -1 }}>ALEJANDRO</div>
            <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>📁 {fileName}</div>
          </div>
          <button onClick={onSignOut} style={ghostBtn}>↩ SALIR</button>
        </div>
        {stagnantCount > 0 && (
          <div style={{ background: "#1A0F00", border: `1px solid ${C.orange}40`, borderRadius: 10, padding: "12px 16px", marginBottom: 16, display: "flex", gap: 10, alignItems: "center" }}>
            <span style={{ fontSize: 18 }}>⚠️</span>
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.orange }}>Estancamiento detectado</div>
              <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>{stagnantCount} ejercicio{stagnantCount > 1 ? "s" : ""} sin mejora en 3+ semanas</div>
            </div>
          </div>
        )}
        <div onClick={onProgress} style={{ background: "#1A1F0A", border: `1px solid ${C.accent}25`, borderRadius: 12, padding: "14px 18px", marginBottom: 20, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 700, color: C.accent }}>📈 Ver progreso</div>
            <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>Gráficas · PRs · evolución semanal</div>
          </div>
          <div style={{ color: C.accent, fontSize: 18 }}>→</div>
        </div>
        <div style={{ fontSize: 10, color: C.muted, letterSpacing: 3, textTransform: "uppercase", marginBottom: 10 }}>Registrar sesión</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {Object.entries(sessions).map(([name, sd]) => {
            const nw = sd.exercises[0]?.nextWeek ?? 0;
            const weekLabel = sd.exercises[0]?.weekLabels?.[nw] || `Semana ${nw + 1}`;
            const stagnant = sd.exercises.filter(e => e.isStagnant).length;
            // Session done: at least 2 exercises have kg or reps data this week
            const exWithData = sd.exercises.filter(ex => {
              const slot = ex.sets[0]?.slots[nw];
              return slot && (
                (slot.reps !== "" && !isNaN(parseFloat(slot.reps))) ||
                (slot.kg !== "" && !isNaN(parseFloat(slot.kg)) && parseFloat(slot.kg) > 0)
              );
            });
            const sessionDone = exWithData.length >= 2;
            const sessionPartial = !sessionDone && exWithData.length === 1;
            return (
              <div key={name} onClick={() => onSession(name)}
                style={{ background: C.card, border: `1px solid ${sessionDone ? "#2A4A2A" : "#222"}`, borderRadius: 12, padding: "16px 18px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: sessionDone ? "#5A9A5A" : C.text }}>
                    {name}
                    {sessionDone && <span style={{ fontSize: 10, color: "#5A9A5A", marginLeft: 8 }}>✓</span>}
                  </div>
                  <div style={{ fontSize: 10, color: C.muted, marginTop: 3, display: "flex", gap: 8 }}>
                    <span>Semana del {weekLabel}</span>
                    {sessionDone && <span style={{ color: "#5A9A5A" }}>completada</span>}
                    {sessionPartial && <span style={{ color: C.orange }}>en curso</span>}
                    {stagnant > 0 && <span style={{ color: C.orange }}>⚠ {stagnant}</span>}
                  </div>
                </div>
                <div style={{ color: sessionDone ? "#5A9A5A" : C.accent, fontSize: 20 }}>→</div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ── LOG ────────────────────────────────────────────────────────────────────────
function Log({ session, sd, form, openEx, setOpenEx, substitutions, setSubstitutions, editedRepsObj, setEditedRepsObj, onSet, onSave, saving, saveError, onBack }) {
  const [subModal, setSubModal] = useState(null);
  const [timer, setTimer] = useState(null); // { secs, running, elapsed }
  const timerRef = useRef(null);

  // Timer controls
  const startTimer = (secs) => {
    if (timerRef.current) clearInterval(timerRef.current);
    setTimer({ total: secs, remaining: secs, running: true });
    timerRef.current = setInterval(() => {
      setTimer(t => {
        if (!t || t.remaining <= 1) {
          clearInterval(timerRef.current);
          // Vibrate on finish if supported
          if (navigator.vibrate) navigator.vibrate([200, 100, 200]);
          return { ...t, remaining: 0, running: false };
        }
        return { ...t, remaining: t.remaining - 1 };
      });
    }, 1000);
  };

  const stopTimer = () => {
    if (timerRef.current) clearInterval(timerRef.current);
    setTimer(null);
  }; // { exName, newName, seriesRepsObj[] }
  const exercises = sd?.exercises || [];
  const nextWeek = exercises[0]?.nextWeek ?? 0;
  const prevWeek = nextWeek - 1;

  const openSubModal = (ex) => {
    setSubModal({
      exName: ex.name,
      newName: substitutions[ex.name] || "",
      seriesRepsObj: ex.sets.map((set, si) => editedRepsObj[ex.name]?.[si] ?? set.repsObj ?? ""),
    });
  };

  const confirmSub = () => {
    if (subModal.newName.trim()) setSubstitutions(p => ({ ...p, [subModal.exName]: subModal.newName.trim() }));
    // Save edited repsObj
    const newEdited = {};
    subModal.seriesRepsObj.forEach((val, si) => { newEdited[si] = val; });
    setEditedRepsObj(p => ({ ...p, [subModal.exName]: newEdited }));
    setSubModal(null);
  };

  return (
    <div style={{ background: C.dark, minHeight: "100vh", color: C.text, fontFamily: "'DM Mono','Courier New',monospace" }}>
      <div style={{ maxWidth: 480, margin: "0 auto", padding: "0 16px 130px" }}>
        <div style={{ padding: "20px 0 12px", display: "flex", alignItems: "center", gap: 12 }}>
          <button onClick={onBack} style={{ background: "none", border: "none", color: C.muted, fontSize: 22, cursor: "pointer", padding: 0 }}>←</button>
          <div>
            <div style={{ fontSize: 10, color: C.muted, letterSpacing: 2, textTransform: "uppercase" }}>Sesión de hoy</div>
            <div style={{ fontSize: 18, fontWeight: 700, color: C.accent }}>{session}</div>
            <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>
              {exercises[0]?.weekLabels?.[nextWeek] || `Semana ${nextWeek + 1}`}
              {prevWeek >= 0 ? ` · ref. ${exercises[0]?.weekLabels?.[prevWeek] || `S${prevWeek+1}`}` : " · primera sesión"}
            </div>
          </div>
        </div>

        {/* Substitution modal */}
        {subModal && (
          <div style={{ position: "fixed", inset: 0, background: "#000C", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
            <div style={{ background: "#1A1A1A", borderRadius: 16, padding: 22, width: "100%", maxWidth: 380, border: "1px solid #333", maxHeight: "80vh", overflowY: "auto" }}>
              <div style={{ fontSize: 9, color: C.muted, letterSpacing: 2, marginBottom: 6 }}>EDITAR EJERCICIO</div>
              <div style={{ fontSize: 11, color: "#888", marginBottom: 14, lineHeight: 1.4 }}>{subModal.exName}</div>

              {/* New exercise name */}
              <div style={{ fontSize: 9, color: C.muted, letterSpacing: 1, marginBottom: 6 }}>NOMBRE NUEVO (opcional)</div>
              <input type="text" value={subModal.newName}
                onChange={e => setSubModal(p => ({ ...p, newName: e.target.value }))}
                placeholder="Dejar vacío para mantener el actual"
                style={{ ...inp, marginBottom: 18 }} />

              {/* Series reps obj editor */}
              <div style={{ fontSize: 9, color: C.muted, letterSpacing: 1, marginBottom: 10 }}>REPS OBJETIVO POR SERIE</div>
              {subModal.seriesRepsObj.map((val, si) => (
                <div key={si} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                  <div style={{ fontSize: 10, color: C.muted, width: 60 }}>Serie {si + 1}</div>
                  <input type="text" value={val}
                    onChange={e => setSubModal(p => {
                      const s = [...p.seriesRepsObj]; s[si] = e.target.value; return { ...p, seriesRepsObj: s };
                    })}
                    placeholder="ej. 6-8"
                    style={{ ...inp, width: 80, textAlign: "center", fontSize: 13, fontWeight: 600 }} />
                  <div style={{ fontSize: 9, color: C.muted }}>reps</div>
                </div>
              ))}

              <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
                <button onClick={confirmSub} style={{ flex: 1, background: C.accent, color: C.dark, fontFamily: "inherit", fontWeight: 700, fontSize: 12, padding: "11px 0", borderRadius: 20, border: "none", cursor: "pointer" }}>CONFIRMAR</button>
                <button onClick={() => setSubModal(null)} style={{ flex: 1, background: "none", border: "1px solid #333", color: C.muted, fontFamily: "inherit", fontSize: 12, padding: "11px 0", borderRadius: 20, cursor: "pointer" }}>CANCELAR</button>
              </div>
            </div>
          </div>
        )}

        {exercises.map(ex => {
          const isOpen = openEx === ex.name;
          const fEx = form[ex.name] || [];
          const filled = fEx.filter(s => s.kg || s.reps).length;
          const displayName = substitutions[ex.name] || ex.name;

          // Already done: current week has kg or reps data for this exercise
          const currentSlot = ex.sets[0]?.slots[nextWeek];
          const alreadyDone = currentSlot && (
            (currentSlot.reps !== "" && !isNaN(parseFloat(currentSlot.reps))) ||
            (currentSlot.kg !== "" && !isNaN(parseFloat(currentSlot.kg)) && parseFloat(currentSlot.kg) > 0)
          );

          return (
            <div key={ex.name} style={{ background: C.card, borderRadius: 12, marginBottom: 8, overflow: "hidden", border: `1px solid ${alreadyDone ? "#1A3A1A" : ex.isStagnant ? C.orange + "50" : filled > 0 ? "#2A3A10" : "#1E1E1E"}` }}>
              <div onClick={() => setOpenEx(isOpen ? null : ex.name)} style={{ padding: "13px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                <div style={{ flex: 1, marginRight: 8 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: alreadyDone ? "#5A9A5A" : filled > 0 ? C.accent : C.text, lineHeight: 1.3 }}>
                    {displayName}
                    {substitutions[ex.name] && <span style={{ fontSize: 9, color: C.orange, marginLeft: 6 }}>SUSTITUIDO</span>}
                    {alreadyDone && <span style={{ fontSize: 9, color: "#5A9A5A", marginLeft: 6 }}>✓ YA REGISTRADO</span>}
                  </div>
                  <div style={{ fontSize: 9, color: C.muted, marginTop: 3, display: "flex", gap: 8 }}>
                    <span>{ex.sets.length} series</span>
                    {alreadyDone && <span style={{ color: "#5A9A5A" }}>{ex.sets.filter(s => s.slots[nextWeek]?.reps).length} series ya hechas esta semana</span>}
                    {!alreadyDone && ex.isStagnant && <span style={{ color: C.orange }}>⚠ estancado</span>}
                  </div>
                </div>
                <span style={{ color: isOpen ? C.accent : C.muted, fontSize: 14 }}>{isOpen ? "▲" : "▼"}</span>
              </div>

              {isOpen && (
                <div style={{ padding: "0 16px 16px" }}>

                  {/* Already done warning */}
                  {alreadyDone && (
                    <div style={{ background: "#0A1A0A", border: "1px solid #2A4A2A", borderRadius: 8, padding: "12px 14px", marginBottom: 12 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "#5A9A5A", marginBottom: 4 }}>✓ Este ejercicio ya está registrado esta semana</div>
                      <div style={{ fontSize: 10, color: C.muted, lineHeight: 1.7 }}>
                        Ya tienes datos para la semana del <strong style={{ color: C.text }}>{ex.weekLabels?.[nextWeek]}</strong>. Si necesitas corregir algo, edita directamente el Excel.
                      </div>
                      <div style={{ marginTop: 10, display: "flex", gap: 8 }}>
                        {ex.sets.map((set, si) => {
                          const slot = set.slots[nextWeek];
                          if (!slot?.kg && !slot?.reps) return null;
                          return (
                            <div key={si} style={{ background: "#111", borderRadius: 6, padding: "6px 10px", textAlign: "center" }}>
                              <div style={{ fontSize: 8, color: C.muted, marginBottom: 2 }}>{set.label.replace(" SERIE","")}</div>
                              <div style={{ fontSize: 13, fontWeight: 700, color: "#5A9A5A" }}>{slot.kg}<span style={{ fontSize: 8, color: C.muted }}>kg</span></div>
                              <div style={{ fontSize: 11, color: C.text }}>{slot.reps}<span style={{ fontSize: 8, color: C.muted }}>r</span></div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {!alreadyDone && ex.isStagnant && (
                    <div style={{ background: "#1A0F00", borderRadius: 8, padding: "8px 12px", marginBottom: 10, fontSize: 10, color: C.orange, lineHeight: 1.6 }}>
                      ⚠️ Sin mejora en 3+ semanas · considera subir carga o cambiar el ejercicio
                    </div>
                  )}

                  {/* Input section — hidden if already done this week */}
                  {!alreadyDone && (<>

                  {/* Column headers */}
                  <div style={{ display: "grid", gridTemplateColumns: "20px 80px 1fr 1fr 1fr", gap: 6, marginBottom: 8, paddingBottom: 6, borderBottom: "1px solid #1E1E1E" }}>
                    <div /><div style={hdr} />
                    {["KG", "REPS", "RIR"].map(h => <div key={h} style={hdr}>{h}</div>)}
                  </div>

                  {ex.sets.map((set, si) => {
                    const f = fEx[si] || { kg: "", reps: "", rir: "", notes: "" };
                    const prev = prevWeek >= 0 ? set.slots[prevWeek] : null;
                    const pKg = prev?.kg || "—"; const pReps = prev?.reps || "—"; const pRir = prev?.rir || "—";
                    const repsObj = editedRepsObj[ex.name]?.[si] ?? set.repsObj;
                    const rec = getLoadRec(prev?.reps, repsObj);
                    const suggestion = suggestProgression(prev?.kg, prev?.reps, repsObj);
                    // Pre-fill suggestion into form if field is still empty
                    const sugKg = suggestion ? String(suggestion.kg) : pKg;
                    const sugReps = suggestion ? String(suggestion.reps) : pReps;

                    return (
                      <div key={si} style={{ marginBottom: si < ex.sets.length - 1 ? 16 : 0 }}>

                        {/* Load alert */}
                        {rec && rec.type !== "ok" && (
                          <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6, background: rec.type === "up" ? "#0A1A0A" : "#1A0A0A", borderRadius: 6, padding: "6px 10px" }}>
                            <span style={{ fontSize: 13 }}>{rec.type === "up" ? "⬆️" : "⬇️"}</span>
                            <span style={{ fontSize: 10, color: rec.type === "up" ? C.green : C.red, fontWeight: 600 }}>{rec.msg}</span>
                          </div>
                        )}
                        {rec && rec.type === "ok" && (
                          <div style={{ fontSize: 9, color: "#3A5A3A", marginBottom: 4, paddingLeft: 4 }}>✓ {rec.msg}</div>
                        )}

                        {/* Progression suggestion */}
                        {suggestion && (
                          <div style={{ background: "#0D1A0D", border: `1px solid ${C.accent}30`, borderRadius: 8, padding: "8px 12px", marginBottom: 8 }}>
                            <div style={{ fontSize: 9, color: C.accent, letterSpacing: 1, marginBottom: 6 }}>💡 PROPUESTA PROGRESIÓN</div>
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 6, alignItems: "center" }}>
                              <div>
                                <div style={{ fontSize: 8, color: C.muted, marginBottom: 3, textAlign: "center" }}>KG</div>
                                <input type="number" defaultValue={suggestion.kg}
                                  onChange={e => onSet(ex.name, si, "kg", e.target.value)}
                                  style={{ ...inp, textAlign: "center", fontSize: 15, fontWeight: 700, color: C.accent, borderColor: C.accent + "50" }} />
                              </div>
                              <div>
                                <div style={{ fontSize: 8, color: C.muted, marginBottom: 3, textAlign: "center" }}>REPS MÍN.</div>
                                <input type="number" defaultValue={suggestion.reps}
                                  onChange={e => onSet(ex.name, si, "reps", e.target.value)}
                                  style={{ ...inp, textAlign: "center", fontSize: 15, fontWeight: 700, color: C.accent, borderColor: C.accent + "50" }} />
                              </div>
                              <button onClick={() => { onSet(ex.name, si, "kg", String(suggestion.kg)); onSet(ex.name, si, "reps", String(suggestion.reps)); }}
                                style={{ background: C.accent, color: C.dark, border: "none", borderRadius: 8, padding: "8px 10px", fontSize: 9, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", letterSpacing: 1, marginTop: 14 }}>
                                USAR
                              </button>
                            </div>
                            <div style={{ fontSize: 8, color: C.muted, marginTop: 6, lineHeight: 1.5 }}>{suggestion.reason}</div>
                          </div>
                        )}

                        {/* Reps objective */}
                        {repsObj && (
                          <div style={{ fontSize: 8, color: C.muted, marginBottom: 3, paddingLeft: 104, letterSpacing: 0.5 }}>
                            obj. {repsObj} reps
                          </div>
                        )}

                        {/* Previous week row */}
                        {prevWeek >= 0 && (
                          <div style={{ display: "grid", gridTemplateColumns: "20px 80px 1fr 1fr 1fr", gap: 6, marginBottom: 3, alignItems: "center" }}>
                            <div style={{ fontSize: 9, color: C.muted, textAlign: "center" }}>{si + 1}</div>
                            <div style={{ fontSize: 9, color: C.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{ex.weekLabels?.[prevWeek] || `S${prevWeek+1}`}</div>
                            {[pKg, pReps, pRir].map((v, i) => (
                              <div key={i} style={{ textAlign: "center", fontSize: 12, color: "#555", background: "#0A0A0A", border: "1px solid #1A1A1A", borderRadius: 5, padding: "6px 4px" }}>{v}</div>
                            ))}
                          </div>
                        )}

                        {/* New week row */}
                        <div style={{ display: "grid", gridTemplateColumns: "20px 80px 1fr 1fr 1fr", gap: 6, alignItems: "center" }}>
                          <div />
                          <div style={{ fontSize: 9, color: C.accent, fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{ex.weekLabels?.[nextWeek] || `S${nextWeek+1}`}</div>
                          {["kg", "reps", "rir"].map(field => (
                            <input key={field} type="number" value={f[field] || ""}
                              onChange={e => onSet(ex.name, si, field, e.target.value)}
                              placeholder={field === "kg" ? sugKg : field === "reps" ? sugReps : pRir}
                              style={{ ...inp, textAlign: "center", fontSize: 13, fontWeight: 600 }} />
                          ))}
                        </div>
                        <div style={{ marginTop: 4, paddingLeft: 104 }}>
                          <input type="text" value={f.notes || ""} onChange={e => onSet(ex.name, si, "notes", e.target.value)}
                            placeholder="nota…" style={{ ...inp, fontSize: 10, color: "#888", width: "100%" }} />
                        </div>
                        {/* Timer trigger — only show after filling reps */}
                        {f.reps && (
                          <div style={{ paddingLeft: 104, marginTop: 6, display: "flex", gap: 6 }}>
                            {[90, 120, 180].map(s => (
                              <button key={s} onClick={() => startTimer(s)}
                                style={{ background: timer?.running ? "#1A1A1A" : "#111", border: `1px solid ${timer?.running ? "#333" : C.muted2}`, color: C.muted, fontFamily: "inherit", fontSize: 9, padding: "4px 8px", borderRadius: 10, cursor: "pointer" }}>
                                {s / 60}'{s % 60 > 0 ? s % 60 + '"' : ""}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })}

                  </>)} {/* end !alreadyDone */}

                  {/* Edit / substitute button — always visible */}
                  <div style={{ marginTop: 14, borderTop: "1px solid #1E1E1E", paddingTop: 12, display: "flex", gap: 8 }}>
                    <button onClick={() => openSubModal(ex)}
                      style={{ background: "none", border: "1px solid #333", color: C.muted, fontFamily: "inherit", fontSize: 10, padding: "6px 14px", borderRadius: 16, cursor: "pointer", letterSpacing: 1 }}>
                      ✏ EDITAR EJERCICIO
                    </button>
                    {(substitutions[ex.name] || editedRepsObj[ex.name]) && (
                      <button onClick={() => { setSubstitutions(p => { const n = { ...p }; delete n[ex.name]; return n; }); setEditedRepsObj(p => { const n = { ...p }; delete n[ex.name]; return n; }); }}
                        style={{ background: "none", border: "none", color: C.red, fontFamily: "inherit", fontSize: 10, padding: "6px 4px", cursor: "pointer" }}>
                        × resetear
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: `linear-gradient(transparent, ${C.dark} 55%)`, padding: "24px 16px 28px" }}>
        <div style={{ maxWidth: 480, margin: "0 auto" }}>

          {/* Timer bar */}
          {timer && (
            <div style={{ background: timer.remaining === 0 ? "#0A1A0A" : "#0D0D0D", border: `1px solid ${timer.remaining === 0 ? C.accent : "#333"}`, borderRadius: 16, padding: "10px 16px", marginBottom: 10, display: "flex", alignItems: "center", gap: 12 }}>
              {/* Progress ring */}
              <div style={{ position: "relative", width: 44, height: 44, flexShrink: 0 }}>
                <svg width="44" height="44" style={{ transform: "rotate(-90deg)" }}>
                  <circle cx="22" cy="22" r="18" fill="none" stroke="#222" strokeWidth="3" />
                  <circle cx="22" cy="22" r="18" fill="none"
                    stroke={timer.remaining === 0 ? C.accent : C.muted}
                    strokeWidth="3"
                    strokeDasharray={`${2 * Math.PI * 18}`}
                    strokeDashoffset={`${2 * Math.PI * 18 * (1 - timer.remaining / timer.total)}`}
                    strokeLinecap="round"
                    style={{ transition: "stroke-dashoffset 1s linear, stroke 0.3s" }}
                  />
                </svg>
                <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 700, color: timer.remaining === 0 ? C.accent : C.text, fontFamily: "'DM Mono',monospace" }}>
                  {timer.remaining > 0 ? `${Math.floor(timer.remaining / 60)}:${String(timer.remaining % 60).padStart(2,"0")}` : "✓"}
                </div>
              </div>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: timer.remaining === 0 ? C.accent : C.text }}>
                  {timer.remaining === 0 ? "¡A por la siguiente serie!" : "Descansando..."}
                </div>
                <div style={{ fontSize: 9, color: C.muted, marginTop: 2 }}>
                  {timer.remaining > 0 ? `${timer.remaining}s restantes de ${timer.total}s` : "Descanso completado"}
                </div>
              </div>
              <button onClick={stopTimer}
                style={{ background: "none", border: "1px solid #333", color: C.muted, fontFamily: "inherit", fontSize: 10, padding: "5px 10px", borderRadius: 10, cursor: "pointer" }}>
                ✕
              </button>
            </div>
          )}

          {saveError && (
            <div style={{ background: "#1A0A0A", border: "1px solid #4A2A2A", borderRadius: 10, padding: "8px 14px", marginBottom: 8, fontSize: 10, color: "#FF6B6B" }}>
              Error guardando: {saveError}
            </div>
          )}
          <button onClick={onSave} disabled={saving}
            style={{ width: "100%", background: saving ? "#2A3A10" : C.accent, color: C.dark, fontFamily: "'DM Mono',monospace", fontWeight: 700, fontSize: 13, padding: "16px 0", borderRadius: 30, border: "none", cursor: saving ? "default" : "pointer", letterSpacing: 2, opacity: saving ? 0.8 : 1 }}>
            {saving ? "☁ GUARDANDO EN DRIVE..." : "☁ GUARDAR EN DRIVE"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── PROGRESS ───────────────────────────────────────────────────────────────────
function Progress({ sessions, onBack }) {
  const [selSession, setSelSession] = useState(Object.keys(sessions)[0]);
  const [selEx, setSelEx] = useState(null);
  const [metric, setMetric] = useState("kg");

  const sessionExercises = sessions[selSession]?.exercises.filter(ex => buildChartData(ex).length >= 2) || [];
  const activeEx = selEx ? sessionExercises.find(e => e.name === selEx) || sessionExercises[0] : sessionExercises[0];
  const chartData = activeEx ? buildChartData(activeEx) : [];
  const firstPt = chartData[0]; const lastPt = chartData.at(-1);
  const kgDiff = firstPt && lastPt ? (lastPt.kg - firstPt.kg).toFixed(1) : null;
  const repsDiff = firstPt && lastPt ? (lastPt.reps - firstPt.reps).toFixed(0) : null;

  return (
    <div style={{ background: C.dark, minHeight: "100vh", color: C.text, fontFamily: "'DM Mono','Courier New',monospace", paddingBottom: 40 }}>
      <div style={{ maxWidth: 480, margin: "0 auto", padding: "0 16px" }}>
        <div style={{ padding: "20px 0 16px", display: "flex", alignItems: "center", gap: 12 }}>
          <button onClick={onBack} style={{ background: "none", border: "none", color: C.muted, fontSize: 22, cursor: "pointer", padding: 0 }}>←</button>
          <div>
            <div style={{ fontSize: 10, color: C.muted, letterSpacing: 2, textTransform: "uppercase" }}>1ª serie · progreso</div>
            <div style={{ fontSize: 18, fontWeight: 700, color: C.accent }}>EVOLUCIÓN</div>
          </div>
        </div>

        {/* Session tabs */}
        <div style={{ display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" }}>
          {Object.keys(sessions).map(s => (
            <button key={s} onClick={() => { setSelSession(s); setSelEx(null); }}
              style={{ background: selSession === s ? C.accent : C.card, color: selSession === s ? C.dark : C.muted, border: `1px solid ${selSession === s ? C.accent : "#333"}`, borderRadius: 20, fontSize: 10, padding: "5px 12px", cursor: "pointer", fontFamily: "inherit", fontWeight: selSession === s ? 700 : 400 }}>
              {s}
            </button>
          ))}
        </div>

        {sessionExercises.length === 0 ? (
          <div style={{ textAlign: "center", color: C.muted, fontSize: 12, padding: 40 }}>Sin datos suficientes (mínimo 2 semanas).</div>
        ) : (
          <>
            {/* Exercise chips */}
            <div style={{ display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" }}>
              {sessionExercises.map(ex => (
                <button key={ex.name} onClick={() => setSelEx(ex.name)}
                  style={{ background: activeEx?.name === ex.name ? "#1A1F0A" : C.card, color: activeEx?.name === ex.name ? C.accent : ex.isStagnant ? C.orange : C.muted, border: `1px solid ${activeEx?.name === ex.name ? C.accent : ex.isStagnant ? C.orange + "40" : "#222"}`, borderRadius: 20, fontSize: 9, padding: "5px 10px", cursor: "pointer", fontFamily: "inherit", maxWidth: 170, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {ex.isStagnant ? "⚠ " : ""}{ex.name}
                </button>
              ))}
            </div>

            {activeEx && chartData.length >= 2 && (
              <>
                {/* Stats */}
                <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
                  {[
                    { label: "KG actual", value: `${lastPt.kg}`, unit: "kg", delta: kgDiff, color: C.accent },
                    { label: "Reps actual", value: `${lastPt.reps}`, unit: "reps", delta: repsDiff, color: C.blue },
                    { label: "Semanas", value: `${chartData.length}`, unit: "", delta: null, color: C.text },
                  ].map(({ label, value, unit, delta, color }) => (
                    <div key={label} style={{ flex: 1, background: C.card, borderRadius: 10, padding: "12px 10px", border: "1px solid #222" }}>
                      <div style={{ fontSize: 8, color: C.muted, letterSpacing: 1, marginBottom: 4 }}>{label.toUpperCase()}</div>
                      <div style={{ fontSize: 18, fontWeight: 700, color }}>{value}<span style={{ fontSize: 9, color: C.muted, marginLeft: 2 }}>{unit}</span></div>
                      {delta != null && (
                        <div style={{ fontSize: 9, color: parseFloat(delta) >= 0 ? C.accent : C.red, marginTop: 2 }}>
                          {parseFloat(delta) >= 0 ? "▲" : "▼"} {Math.abs(delta)}{unit} vs S1
                        </div>
                      )}
                    </div>
                  ))}
                </div>

                {/* PR — shown per selected exercise */}
                {activeEx.pr && (
                  <div style={{ background: "#0A0F18", border: `1px solid ${C.blue}30`, borderRadius: 10, padding: "12px 16px", marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div>
                      <div style={{ fontSize: 9, color: C.blue, letterSpacing: 2, textTransform: "uppercase", marginBottom: 3 }}>★ Mejor marca personal</div>
                      <div style={{ fontSize: 10, color: C.muted }}>{activeEx.name}</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>{activeEx.pr.kg}<span style={{ fontSize: 9, color: C.muted }}> kg</span> × {activeEx.pr.reps}<span style={{ fontSize: 9, color: C.muted }}> reps</span></div>
                      <div style={{ fontSize: 9, color: C.blue, marginTop: 2 }}>{activeEx.pr.weekLabel}</div>
                    </div>
                  </div>
                )}

                {activeEx.isStagnant && (
                  <div style={{ background: "#1A0F00", border: `1px solid ${C.orange}40`, borderRadius: 8, padding: "10px 14px", marginBottom: 12, fontSize: 10, color: C.orange }}>
                    ⚠️ 3+ semanas sin mejora en la 1ª serie · ajusta carga con tu entrenador
                  </div>
                )}

                {/* Metric toggle */}
                <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
                  {[["kg", "KG"], ["reps", "Reps"], ["both", "Ambos"]].map(([v, l]) => (
                    <button key={v} onClick={() => setMetric(v)}
                      style={{ background: metric === v ? C.accent : C.card, color: metric === v ? C.dark : C.muted, border: `1px solid ${metric === v ? C.accent : "#333"}`, borderRadius: 16, fontSize: 9, padding: "4px 10px", cursor: "pointer", fontFamily: "inherit", fontWeight: metric === v ? 700 : 400 }}>
                      {l}
                    </button>
                  ))}
                </div>

                {/* Chart */}
                <div style={{ background: C.card, borderRadius: 12, padding: "16px 8px 8px", border: "1px solid #1E1E1E", marginBottom: 16 }}>
                  <div style={{ fontSize: 10, color: C.muted, letterSpacing: 1, paddingLeft: 8, marginBottom: 8 }}>{activeEx.name.toUpperCase()}</div>
                  <ResponsiveContainer width="100%" height={200}>
                    <LineChart data={chartData} margin={{ top: 4, right: 16, left: -16, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1E1E1E" />
                      <XAxis dataKey="week" tick={{ fill: C.muted, fontSize: 9 }} axisLine={false} tickLine={false} />
                      <YAxis tick={{ fill: C.muted, fontSize: 9 }} axisLine={false} tickLine={false} />
                      <Tooltip contentStyle={{ background: "#111", border: "1px solid #333", borderRadius: 8, fontSize: 11, fontFamily: "'DM Mono',monospace" }} labelStyle={{ color: C.accent }} />
                      {(metric === "kg" || metric === "both") && <Line type="monotone" dataKey="kg" name="KG" stroke={C.accent} strokeWidth={2} dot={{ r: 3, fill: C.accent }} activeDot={{ r: 5 }} />}
                      {(metric === "reps" || metric === "both") && <Line type="monotone" dataKey="reps" name="Reps" stroke={C.blue} strokeWidth={2} dot={{ r: 3, fill: C.blue }} activeDot={{ r: 5 }} />}
                    </LineChart>
                  </ResponsiveContainer>
                </div>

                {/* Last session breakdown */}
                <div style={{ fontSize: 9, color: C.muted, letterSpacing: 2, textTransform: "uppercase", marginBottom: 8 }}>Última sesión · todas las series</div>
                {activeEx.sets.map((set, si) => {
                  const slot = activeEx.lastFilledWeek >= 0 ? set.slots[activeEx.lastFilledWeek] : null;
                  if (!slot || (!slot.kg && !slot.reps)) return null;
                  return (
                    <div key={si} style={{ background: C.card, borderRadius: 8, padding: "10px 14px", marginBottom: 6, display: "flex", justifyContent: "space-between", alignItems: "center", border: `1px solid ${si === 0 ? "#2A3A10" : "#1E1E1E"}` }}>
                      <div>
                        <div style={{ fontSize: 11, color: si === 0 ? C.accent : C.muted }}>{set.label}{si === 0 ? " ★" : ""}</div>
                        {set.repsObj && <div style={{ fontSize: 8, color: C.muted, marginTop: 2 }}>obj. {set.repsObj} reps</div>}
                      </div>
                      <div style={{ display: "flex", gap: 12 }}>
                        <span style={{ fontSize: 14, fontWeight: 700 }}>{slot.kg}<span style={{ fontSize: 9, color: C.muted }}> kg</span></span>
                        <span style={{ fontSize: 14, fontWeight: 700 }}>{slot.reps}<span style={{ fontSize: 9, color: C.muted }}> reps</span></span>
                        {slot.rir && <span style={{ fontSize: 10, color: C.muted }}>RIR {slot.rir}</span>}
                      </div>
                    </div>
                  );
                })}
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// ── DONE ───────────────────────────────────────────────────────────────────────
function Done({ fileName, newWeekCreated, summary, onBack }) {
  const ups = summary?.filter(s => s.trend === "up").length || 0;
  const downs = summary?.filter(s => s.trend === "down").length || 0;
  const neutrals = summary?.filter(s => s.trend === "neutral").length || 0;

  return (
    <div style={{ background: C.dark, minHeight: "100vh", color: C.text, fontFamily: "'DM Mono',monospace", paddingBottom: 40 }}>
      <div style={{ maxWidth: 480, margin: "0 auto", padding: "0 16px" }}>
        <div style={{ paddingTop: 40, textAlign: "center", marginBottom: 24 }}>
          <div style={{ fontSize: 48, marginBottom: 12 }}>💪</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: C.accent, marginBottom: 4 }}>¡Sesión completada!</div>
          <div style={{ fontSize: 10, color: C.muted }}>{fileName} descargado</div>
        </div>

        {/* Stats row */}
        {summary && summary.length > 0 && (
          <>
            <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
              {[
                { label: "Subida", value: ups, color: C.accent, icon: "⬆" },
                { label: "Igual", value: neutrals, color: C.muted, icon: "→" },
                { label: "Bajada", value: downs, color: C.red, icon: "⬇" },
              ].map(({ label, value, color, icon }) => (
                <div key={label} style={{ flex: 1, background: C.card, borderRadius: 10, padding: "14px 8px", textAlign: "center", border: "1px solid #222" }}>
                  <div style={{ fontSize: 20, marginBottom: 4 }}>{icon}</div>
                  <div style={{ fontSize: 22, fontWeight: 700, color }}>{value}</div>
                  <div style={{ fontSize: 9, color: C.muted, marginTop: 2, letterSpacing: 1 }}>{label.toUpperCase()}</div>
                </div>
              ))}
            </div>

            {/* Per exercise breakdown */}
            <div style={{ fontSize: 9, color: C.muted, letterSpacing: 2, textTransform: "uppercase", marginBottom: 8 }}>Detalle por ejercicio</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 24 }}>
              {summary.map((s, i) => (
                <div key={i} style={{ background: C.card, borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", border: `1px solid ${s.trend === "up" ? "#2A3A10" : s.trend === "down" ? "#3A1A1A" : "#1E1E1E"}` }}>
                  <div style={{ fontSize: 11, color: C.muted, flex: 1, marginRight: 8, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.name}</div>
                  <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
                    {!isNaN(s.prevKg) && (
                      <span style={{ fontSize: 10, color: "#444" }}>{s.prevKg}kg</span>
                    )}
                    <span style={{ fontSize: 11, fontWeight: 700, color: s.trend === "up" ? C.accent : s.trend === "down" ? C.red : C.text }}>
                      {s.trend === "up" ? "↑" : s.trend === "down" ? "↓" : "="} {s.curKg}kg
                    </span>
                    <span style={{ fontSize: 9, color: C.muted }}>×{s.curReps}</span>
                  </div>
                </div>
              ))}
            </div>
          </>
        )}

        {newWeekCreated && (
          <div style={{ background: "#1A1200", border: `1px solid ${C.orange}50`, borderRadius: 10, padding: "12px 16px", marginBottom: 16 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.orange, marginBottom: 4 }}>🆕 Nuevo bloque generado</div>
            <div style={{ fontSize: 10, color: C.muted, lineHeight: 1.7 }}>Se han creado las columnas del siguiente bloque automáticamente.</div>
          </div>
        )}

        <button onClick={onBack} style={{ width: "100%", background: "none", border: `1px solid ${C.accent}`, color: C.accent, fontFamily: "inherit", fontSize: 12, padding: "14px 0", borderRadius: 25, cursor: "pointer", letterSpacing: 2 }}>← VOLVER</button>
      </div>
    </div>
  );
}

const inp = { background: "#0F0F0F", border: "1px solid #2A2A2A", borderRadius: 6, color: C.text, padding: "8px 5px", fontFamily: "'DM Mono','Courier New',monospace", outline: "none", width: "100%", boxSizing: "border-box" };
const ghostBtn = { background: "none", border: "1px solid #333", color: C.muted, fontSize: 10, padding: "6px 12px", borderRadius: 8, cursor: "pointer", fontFamily: "'DM Mono',monospace", letterSpacing: 1 };
const hdr = { fontSize: 8, color: C.muted, textAlign: "center", letterSpacing: 2 };
