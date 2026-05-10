// ── GOOGLE DRIVE + SHEETS API ─────────────────────────────────────────────────
// Auth: Google Identity Services token flow (no backend needed)
// Write strategy: use Sheets API batchUpdate to write only the specific cells
// that changed — never reupload the full file, preserving all formatting.

const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.readonly',
].join(' ')

// The .xlsx file ID in Drive
const XLSX_FILE_ID = '1WEMvbLfERQl0E09p7cyjObsIP_uMegzL'

let accessToken = null

// ── AUTH ──────────────────────────────────────────────────────────────────────
function loadGIS() {
  return new Promise((resolve) => {
    if (window.google?.accounts) { resolve(); return }
    const script = document.createElement('script')
    script.src = 'https://accounts.google.com/gsi/client'
    script.onload = resolve
    document.head.appendChild(script)
  })
}

export async function signIn() {
  await loadGIS()
  return new Promise((resolve, reject) => {
    const saved = sessionStorage.getItem('gd_token')
    if (saved) { accessToken = saved; resolve(saved); return }
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: (resp) => {
        if (resp.error) { reject(new Error(resp.error)); return }
        accessToken = resp.access_token
        sessionStorage.setItem('gd_token', accessToken)
        resolve(accessToken)
      },
    })
    client.requestAccessToken({ prompt: 'consent' })
  })
}

export function isSignedIn() {
  const saved = sessionStorage.getItem('gd_token')
  if (saved) { accessToken = saved }
  return !!accessToken
}

export function signOut() {
  if (accessToken) window.google?.accounts?.oauth2?.revoke(accessToken)
  accessToken = null
  sessionStorage.removeItem('gd_token')
}

async function apiRequest(url, options = {}) {
  if (!accessToken) throw new Error('No autenticado')
  const resp = await fetch(url, {
    ...options,
    headers: { Authorization: `Bearer ${accessToken}`, ...(options.headers || {}) },
  })
  if (resp.status === 401) { signOut(); throw new Error('TOKEN_EXPIRED') }
  return resp
}

// ── READ: download .xlsx for parsing ─────────────────────────────────────────
export async function findFile() {
  // Just verify the file exists and return its metadata
  const resp = await apiRequest(
    `https://www.googleapis.com/drive/v3/files/${XLSX_FILE_ID}?fields=id,name,modifiedTime`
  )
  const data = await resp.json()
  return data.id ? { ...data, id: XLSX_FILE_ID } : null
}

export async function downloadFile(fileId) {
  const resp = await apiRequest(
    `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`
  )
  return resp.arrayBuffer()
}

// ── WRITE: use Sheets API to update only specific cells ───────────────────────
// cellUpdates: array of { row (0-indexed), col (0-indexed), value }
// sheetName: the sheet tab name e.g. "MESOCICLO I FPARK"
export async function writeCells(sheetName, cellUpdates) {
  if (!cellUpdates || cellUpdates.length === 0) return

  // Convert to A1 notation and build valueRanges
  const valueRanges = cellUpdates
    .filter(u => u.value !== '' && u.value != null)
    .map(u => {
      const col = colToLetter(u.col)
      const row = u.row + 1 // Sheets API is 1-indexed
      return {
        range: `'${sheetName}'!${col}${row}`,
        values: [[u.value]],
      }
    })

  if (valueRanges.length === 0) return

  const resp = await apiRequest(
    `https://sheets.googleapis.com/v4/spreadsheets/${XLSX_FILE_ID}/values:batchUpdate`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        valueInputOption: 'RAW',
        data: valueRanges,
      }),
    }
  )

  if (!resp.ok) {
    const err = await resp.json()
    throw new Error(err.error?.message || 'Error al guardar en Sheets')
  }

  return resp.json()
}

// Convert 0-indexed column number to letter(s): 0→A, 25→Z, 26→AA, etc.
function colToLetter(col) {
  let letter = ''
  col += 1 // 1-indexed
  while (col > 0) {
    const rem = (col - 1) % 26
    letter = String.fromCharCode(65 + rem) + letter
    col = Math.floor((col - 1) / 26)
  }
  return letter
}
