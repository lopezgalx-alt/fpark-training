// ═══════════════════════════════════════════════════════════════════════════════
// MEDIDAS — weekly body measurement tracking
// Reads/writes the MEDIDAS sheet of the same workbook, using the same
// local-first queue as training. Layout (1-indexed in Excel):
//   row 4  = headers, row 5..64 = weeks 1..60
//   col A=SEMANA B=FECHA C=PESO D=CINTURA E=PECHO F=BÍCEPS
//   col G=CUÁDRICEPS H=GEMELO I=GEMELO... J..M = calculated, N=NOTAS
// 0-indexed for the parser/writer: header row 3, first data row 4, col C = 2
// ═══════════════════════════════════════════════════════════════════════════════
import { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";

export const MED_SHEET = "MEDIDAS";
const HDR_ROW = 3;        // 0-indexed (Excel row 4)
const FIRST_ROW = 4;      // 0-indexed (Excel row 5 = week 1)
const TOTAL_WEEKS = 60;

// field key → 0-indexed column
export const MED_FIELDS = [
  { key: "peso",    col: 2, label: "PESO",       unit: "kg", dec: true,  color: "#C8F135" },
  { key: "cintura", col: 3, label: "CINTURA",    unit: "cm", dec: false, color: "#FF6B6B" },
  { key: "pecho",   col: 4, label: "PECHO",      unit: "cm", dec: false, color: "#60A5FA" },
  { key: "biceps",  col: 5, label: "BÍCEPS",     unit: "cm", dec: true,  color: "#A78BFA" },
  { key: "cuad",    col: 6, label: "CUÁDRICEPS", unit: "cm", dec: true,  color: "#F59E0B" },
  { key: "gemelo",  col: 7, label: "GEMELO",     unit: "cm", dec: true,  color: "#34D399" },
  { key: "gluteo",  col: 8, label: "GLÚTEO",     unit: "cm", dec: true,  color: "#22D3EE" },
];
const NOTES_COL = 13;
export const MED_DATE_COL = 1;   // column B — FECHA
const DATE_COL = MED_DATE_COL;

// Excel stores dates as days since 1899-12-30
function excelSerialToDate(n) {
  if (!n || n < 20000) return null;
  return new Date(Math.round((n - 25569) * 86400 * 1000));
}

// Monday of the week containing `d`
export function mondayOf(d = new Date()) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const shift = (x.getDay() + 6) % 7;   // Mon=0 ... Sun=6
  x.setDate(x.getDate() - shift);
  return x;
}

export const fmtDate = d => d
  ? `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}`
  : null;

export const fmtDateFull = d => d
  ? `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`
  : null;

// Infer the date of a week from any other row that has one (weeks are 7 days apart)
export function inferDate(rows, idx) {
  if (rows[idx]?.date) return rows[idx].date;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i]?.date) {
      const d = new Date(rows[i].date);
      d.setDate(d.getDate() + (idx - i) * 7);
      return d;
    }
  }
  return null;
}

// ── Parse the MEDIDAS sheet into an array of week records ────────────────────
export function parseMedidas(ws) {
  if (!ws) return null;
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const get = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    return cell ? cell.v : null;
  };
  const rows = [];
  for (let i = 0; i < TOTAL_WEEKS; i++) {
    const r = FIRST_ROW + i;
    if (r > range.e.r) break;
    const rec = { week: i + 1, rowIdx: r, values: {}, notes: "", date: null };
    const rawDate = get(r, DATE_COL);
    if (rawDate instanceof Date) rec.date = rawDate;
    else if (typeof rawDate === "number") rec.date = excelSerialToDate(rawDate);
    else if (typeof rawDate === "string" && rawDate.trim()) {
      const m = rawDate.match(/(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?/);
      if (m) {
        const yr = m[3] ? (m[3].length === 2 ? 2000 + +m[3] : +m[3]) : new Date().getFullYear();
        rec.date = new Date(yr, +m[2] - 1, +m[1]);
      }
    }
    MED_FIELDS.forEach(f => {
      const v = get(r, f.col);
      rec.values[f.key] = v == null || v === "" ? "" : String(v);
    });
    const n = get(r, NOTES_COL);
    rec.notes = n == null ? "" : String(n);
    rec.hasData = MED_FIELDS.some(f => rec.values[f.key] !== "");
    rows.push(rec);
  }
  return rows;
}

export const medColOf = key => MED_FIELDS.find(f => f.key === key)?.col ?? null;
export const MED_NOTES_COL = NOTES_COL;

// Last week index (0-based) that has any data; -1 if none
export function lastFilledMed(rows) {
  for (let i = rows.length - 1; i >= 0; i--) if (rows[i].hasData) return i;
  return -1;
}

// ── UI ────────────────────────────────────────────────────────────────────────
const D = {
  bg: "#080808", card: "#111", card2: "#181818", border: "#222",
  accent: "#C8F135", accentDim: "#1A2A05",
  text: "#F0F0F0", muted: "#666", red: "#FF5555", green: "#4ADE80",
  font: "system-ui, -apple-system, sans-serif",
  mono: "'SF Mono', 'Fira Code', monospace",
};


// Y-axis bounds that frame the actual data instead of anchoring near zero.
// A weight series of 76-82 kg plotted from 0 looks flat; padding proportional
// to the observed range makes real week-to-week movement visible.
// Padding = 18% of range, with a floor so a near-flat series still gets air.
export function niceDomain(values, opts = {}) {
  const vals = values.filter(v => typeof v === "number" && !isNaN(v));
  if (!vals.length) return ["auto", "auto"];
  const lo = Math.min(...vals), hi = Math.max(...vals);
  const range = hi - lo;
  const floor = opts.minPad ?? 0.5;
  const pad = Math.max(range * 0.18, floor);
  const step = range > 20 ? 5 : range > 8 ? 2 : range > 3 ? 1 : 0.5;
  const min = Math.floor((lo - pad) / step) * step;
  const max = Math.ceil((hi + pad) / step) * step;
  return [min, max];
}

function fmtDelta(v, dec) {
  if (v == null || isNaN(v)) return null;
  const s = dec ? Math.abs(v).toFixed(1) : String(Math.round(Math.abs(v)));
  if (Math.abs(v) < 0.05) return { txt: "=", color: D.muted };
  return { txt: (v > 0 ? "+" : "−") + s, color: v > 0 ? "#60A5FA" : "#F59E0B" };
}

export default function Medidas({ rows, onOpenPad, onSetDate, onSync, pendingN, saving, onBack }) {
  const [tab, setTab] = useState("entrada");
  const [selField, setSelField] = useState("peso");

  const lastIdx = lastFilledMed(rows);
  // Default to the first empty week after the last filled one
  const [weekIdx, setWeekIdx] = useState(Math.min(lastIdx + 1, rows.length - 1));
  const rec = rows[weekIdx];
  // Prefer the stored date; otherwise extrapolate from any week that has one
  const shownDate = rec.date || inferDate(rows, weekIdx);
  const prevRec = useMemo(() => {
    for (let i = weekIdx - 1; i >= 0; i--) if (rows[i].hasData) return rows[i];
    return null;
  }, [rows, weekIdx]);

  const chartData = useMemo(() => {
    const f = MED_FIELDS.find(x => x.key === selField);
    return rows.filter(r => r.values[f.key] !== "")
      .map(r => ({ semana: r.week,
                   etq: fmtDate(r.date || inferDate(rows, rows.indexOf(r))) || `S${r.week}`,
                   v: parseFloat(String(r.values[f.key]).replace(",", ".")) }))
      .filter(d => !isNaN(d.v));
  }, [rows, selField]);

  const field = MED_FIELDS.find(f => f.key === selField);
  const first = chartData[0]?.v, last = chartData[chartData.length - 1]?.v;
  const totalDelta = (first != null && last != null) ? last - first : null;

  return (
    <div style={{ background: D.bg, minHeight: "100vh", color: D.text, fontFamily: D.font }}>
      <div style={{ padding: "16px 18px 12px", position: "sticky", top: 0, background: D.bg, zIndex: 10 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <button onClick={onBack} style={{ background: "none", border: "none", color: D.muted, fontSize: 14, padding: "6px 0", cursor: "pointer" }}>‹ Inicio</button>
          <div onClick={pendingN > 0 ? onSync : undefined}
            style={{ fontSize: 10, fontFamily: D.mono, padding: "4px 10px", borderRadius: 20,
              cursor: pendingN > 0 ? "pointer" : "default",
              background: pendingN > 0 ? "#3a2a00" : "#0f2a12",
              border: `1px solid ${pendingN > 0 ? "#8a6d00" : "#1e5228"}`,
              color: pendingN > 0 ? "#ffc933" : "#5fd97a" }}>
            {saving ? "↻ sincronizando" : pendingN > 0 ? `● ${pendingN} pendiente${pendingN > 1 ? "s" : ""}` : "✓ sincronizado"}
          </div>
        </div>
        <div style={{ fontSize: 22, fontWeight: 800, letterSpacing: -0.3 }}>Medidas</div>

        <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
          {[["entrada", "Registrar"], ["graficas", "Evolución"]].map(([k, lbl]) => (
            <button key={k} onClick={() => setTab(k)}
              style={{ flex: 1, padding: "10px 0", borderRadius: 12, fontSize: 13, fontWeight: 700, cursor: "pointer",
                background: tab === k ? D.accent : D.card, color: tab === k ? "#000" : D.muted,
                border: `1px solid ${tab === k ? D.accent : D.border}` }}>
              {lbl}
            </button>
          ))}
        </div>
      </div>

      {tab === "entrada" && (
        <div style={{ padding: "4px 18px 40px" }}>
          {/* Week selector */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "10px 0 16px" }}>
            <button onClick={() => setWeekIdx(i => Math.max(0, i - 1))} disabled={weekIdx === 0}
              style={{ background: D.card2, border: `1px solid ${D.border}`, color: weekIdx === 0 ? D.muted : D.text,
                borderRadius: 10, padding: "10px 16px", fontSize: 16, cursor: "pointer" }}>‹</button>
            <div style={{ flex: 1, textAlign: "center" }}>
              <div style={{ fontSize: 19, fontWeight: 800 }}>
                {shownDate ? fmtDateFull(shownDate) : `Semana ${rec.week}`}
              </div>
              <div style={{ fontSize: 10, color: D.muted, fontFamily: D.mono, marginTop: 3 }}>
                Semana {rec.week} · {rec.hasData ? "registrada" : "sin registrar"}
                {!rec.date && shownDate ? " · fecha estimada" : ""}
              </div>
            </div>
            <button onClick={() => setWeekIdx(i => Math.min(rows.length - 1, i + 1))} disabled={weekIdx === rows.length - 1}
              style={{ background: D.card2, border: `1px solid ${D.border}`, color: D.text,
                borderRadius: 10, padding: "10px 16px", fontSize: 16, cursor: "pointer" }}>›</button>
          </div>

          {/* Fields */}
          {MED_FIELDS.map(f => {
            const val = rec.values[f.key];
            const pv = prevRec ? parseFloat(String(prevRec.values[f.key]).replace(",", ".")) : NaN;
            const cv = parseFloat(String(val).replace(",", "."));
            const delta = (!isNaN(pv) && !isNaN(cv)) ? fmtDelta(cv - pv, f.dec) : null;
            return (
              <div key={f.key} onClick={() => onOpenPad(rec, f)}
                style={{ background: val ? D.card : D.card2, border: `1px solid ${val ? f.color + "35" : D.border}`,
                  borderRadius: 14, padding: "14px 16px", marginBottom: 10,
                  display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                <div>
                  <div style={{ fontSize: 11, color: D.muted, fontFamily: D.mono, letterSpacing: 0.5 }}>{f.label}</div>
                  {prevRec && !isNaN(pv) && (
                    <div style={{ fontSize: 10, color: D.muted, marginTop: 3, fontFamily: D.mono }}>
                      {fmtDate(prevRec.date || inferDate(rows, rows.indexOf(prevRec))) || `sem. ${prevRec.week}`}: {prevRec.values[f.key]} {f.unit}
                    </div>
                  )}
                </div>
                <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
                  {delta && <span style={{ fontSize: 12, fontWeight: 700, color: delta.color, fontFamily: D.mono }}>{delta.txt}</span>}
                  <span style={{ fontSize: 26, fontWeight: 800, color: val ? f.color : D.muted, fontFamily: D.mono }}>
                    {val || "—"}
                  </span>
                  <span style={{ fontSize: 11, color: D.muted }}>{f.unit}</span>
                </div>
              </div>
            );
          })}

          {!rec.date && (
            <button onClick={() => onSetDate(rec, shownDate || mondayOf())}
              style={{ width: "100%", background: D.card2, border: `1px dashed ${D.border}`, color: D.muted,
                borderRadius: 12, padding: "12px 0", fontSize: 12, cursor: "pointer", marginTop: 4 }}>
              Fijar fecha {fmtDateFull(shownDate || mondayOf())} en la hoja
            </button>
          )}

          <div style={{ fontSize: 10, color: D.muted, textAlign: "center", marginTop: 16, lineHeight: 1.6 }}>
            Mide siempre en las mismas condiciones: mismo punto, misma hora, en ayunas si puedes.
          </div>
        </div>
      )}

      {tab === "graficas" && (
        <div style={{ padding: "4px 18px 40px" }}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16 }}>
            {MED_FIELDS.map(f => (
              <button key={f.key} onClick={() => setSelField(f.key)}
                style={{ padding: "7px 12px", borderRadius: 20, fontSize: 11, fontWeight: 700, cursor: "pointer",
                  background: selField === f.key ? f.color : D.card,
                  color: selField === f.key ? "#000" : D.muted,
                  border: `1px solid ${selField === f.key ? f.color : D.border}` }}>
                {f.label}
              </button>
            ))}
          </div>

          {chartData.length === 0 ? (
            <div style={{ textAlign: "center", color: D.muted, fontSize: 13, padding: "40px 0" }}>
              Aún no hay datos de {field.label.toLowerCase()}.
            </div>
          ) : (
            <>
              <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "16px 12px 8px", marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", padding: "0 6px 12px" }}>
                  <div>
                    <div style={{ fontSize: 10, color: D.muted, fontFamily: D.mono }}>ACTUAL</div>
                    <div style={{ fontSize: 24, fontWeight: 800, color: field.color, fontFamily: D.mono }}>
                      {last}<span style={{ fontSize: 12, color: D.muted }}> {field.unit}</span>
                    </div>
                  </div>
                  {totalDelta != null && (
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontSize: 10, color: D.muted, fontFamily: D.mono }}>DESDE {chartData[0].etq}</div>
                      <div style={{ fontSize: 24, fontWeight: 800, fontFamily: D.mono,
                        color: totalDelta > 0 ? "#60A5FA" : "#F59E0B" }}>
                        {totalDelta > 0 ? "+" : "−"}{Math.abs(totalDelta).toFixed(1)}
                        <span style={{ fontSize: 12, color: D.muted }}> {field.unit}</span>
                      </div>
                    </div>
                  )}
                </div>
                <ResponsiveContainer width="100%" height={200}>
                  <LineChart data={chartData} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
                    <CartesianGrid stroke="#1e1e1e" vertical={false} />
                    <XAxis dataKey="etq" minTickGap={28} tick={{ fill: "#666", fontSize: 10 }} axisLine={{ stroke: "#222" }} tickLine={false} />
                    <YAxis domain={niceDomain(chartData.map(d => d.v))} allowDecimals tick={{ fill: "#666", fontSize: 10 }} axisLine={false} tickLine={false} />
                    <Tooltip contentStyle={{ background: "#181818", border: "1px solid #333", borderRadius: 10, fontSize: 12 }}
                      labelStyle={{ color: "#999" }} labelFormatter={l => l}
                      formatter={v => [`${v} ${field.unit}`, field.label]} />
                    <Line type="monotone" dataKey="v" stroke={field.color} strokeWidth={2.5} dot={false} activeDot={{ r: 4 }} />
                  </LineChart>
                </ResponsiveContainer>
              </div>

              <div style={{ display: "flex", gap: 10 }}>
                {[["MÍNIMO", Math.min(...chartData.map(d => d.v))],
                  ["MÁXIMO", Math.max(...chartData.map(d => d.v))],
                  ["REGISTROS", chartData.length]].map(([lbl, v]) => (
                  <div key={lbl} style={{ flex: 1, background: D.card2, borderRadius: 12, padding: "12px 8px", textAlign: "center" }}>
                    <div style={{ fontSize: 9, color: D.muted, fontFamily: D.mono }}>{lbl}</div>
                    <div style={{ fontSize: 17, fontWeight: 800, marginTop: 3, fontFamily: D.mono }}>{v}</div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      )}
    </div>
  );
}
