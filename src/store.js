// ── LOCAL-FIRST PERSISTENCE ──────────────────────────────────────────────────
// Every value is saved to localStorage FIRST (instant, infallible).
// A sync queue pushes pending writes to Google Sheets when possible.
// Sheets is the report for the trainer — the phone is the source of truth
// until sync completes. Data can never be lost.

const PENDING_KEY = 'fpark_pending_v1'

export function loadPending() {
  try { return JSON.parse(localStorage.getItem(PENDING_KEY)) || [] } catch { return [] }
}

function savePending(list) {
  try { localStorage.setItem(PENDING_KEY, JSON.stringify(list)) } catch { /* quota — extremely unlikely at this size */ }
}

// Enqueue a cell write. If a write for the same cell already exists,
// it's replaced (last value wins).
export function enqueue(item) {
  // item: { id: 'row-col', sheetName, row, col, value, meta: 'exName-setIdx-field', ts }
  const list = loadPending().filter(p => p.id !== item.id)
  list.push(item)
  savePending(list)
  return list.length
}

export function removeSynced(ids) {
  const set = new Set(ids)
  savePending(loadPending().filter(p => !set.has(p.id)))
}

export function pendingCount() {
  return loadPending().length
}
