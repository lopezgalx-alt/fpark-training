const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.readonly'
const SHEET_ID = '1TYwQaXpA7W1mFmrYn-wZI5_TJ7J4v_BxzNiDVcb4wjg'
const SHEET_NAME = 'MESOCICLO I FPARK'

// ── AUTH ──────────────────────────────────────────────────────────────────────
// Google access tokens live 1 hour. The browser implicit flow has no refresh
// token, but GIS can mint a new one with NO user interface via
// requestAccessToken({ prompt: '' }) as long as the user still has a Google
// session and already granted consent. So: track expiry, refresh silently
// before it runs out, and only fall back to the consent screen if that fails.
// Critically, a 401 must NOT revoke the grant — revoking is what forced the
// full consent screen every hour.

const EXPIRY_BUFFER_MS = 5 * 60 * 1000   // refresh 5 min early

let accessToken = null
let tokenExpiry = 0
let tokenClient = null
let inFlight = null                      // dedupe concurrent refreshes

function loadGIS() {
  return new Promise((resolve, reject) => {
    if (window.google?.accounts) { resolve(); return }
    const existing = document.querySelector('script[src*="accounts.google.com/gsi"]')
    if (existing) { existing.addEventListener('load', () => resolve()); return }
    const script = document.createElement('script')
    script.src = 'https://accounts.google.com/gsi/client'
    script.onload = () => resolve()
    script.onerror = () => reject(new Error('No se pudo cargar Google Sign-In'))
    document.head.appendChild(script)
  })
}

async function ensureClient() {
  await loadGIS()
  if (!tokenClient) {
    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: () => {},        // replaced per request
    })
  }
  return tokenClient
}

function persist(resp) {
  accessToken = resp.access_token
  tokenExpiry = Date.now() + (Number(resp.expires_in) || 3600) * 1000
  try {
    localStorage.setItem('gd_token', accessToken)
    localStorage.setItem('gd_token_exp', String(tokenExpiry))
  } catch { /* private mode — token still held in memory */ }
}

function restore() {
  try {
    const t = localStorage.getItem('gd_token')
    const e = Number(localStorage.getItem('gd_token_exp') || 0)
    if (t) { accessToken = t; tokenExpiry = e }
  } catch { /* ignore */ }
}
restore()

const tokenValid = () => !!accessToken && Date.now() < tokenExpiry - EXPIRY_BUFFER_MS

// prompt '' = silent (no UI). prompt 'consent' = full consent screen.
function requestToken(prompt) {
  if (inFlight) return inFlight
  inFlight = (async () => {
    const client = await ensureClient()
    return new Promise((resolve, reject) => {
      let settled = false
      const done = fn => (...a) => { if (!settled) { settled = true; fn(...a) } }
      const ok = done(resolve), fail = done(reject)
      // Silent refresh can hang if Google never calls back
      const timer = setTimeout(() => fail(new Error('SILENT_TIMEOUT')), 12000)
      client.callback = resp => {
        clearTimeout(timer)
        if (resp.error) { fail(new Error(resp.error)); return }
        persist(resp)
        ok(accessToken)
      }
      client.error_callback = err => { clearTimeout(timer); fail(new Error(err?.type || 'auth_error')) }
      try { client.requestAccessToken({ prompt }) }
      catch (e) { clearTimeout(timer); fail(e) }
    })
  })().finally(() => { inFlight = null })
  return inFlight
}

// Get a usable token, refreshing silently when needed.
// interactive=true allows falling back to the consent screen.
export async function getToken({ interactive = false } = {}) {
  if (tokenValid()) return accessToken
  try {
    return await requestToken('')
  } catch (e) {
    if (!interactive) throw new Error('TOKEN_EXPIRED')
    return requestToken('consent')
  }
}

// Called on app focus/resume so the token is fresh before the user taps anything
export async function warmToken() {
  if (tokenValid()) return true
  try { await requestToken(''); return true } catch { return false }
}

export async function signIn() {
  return getToken({ interactive: true })
}

// True if we hold a token, even an expired one — it can be refreshed silently.
export function isSignedIn() {
  restore()
  return !!accessToken
}

export function signOut() {
  if (accessToken) window.google?.accounts?.oauth2?.revoke(accessToken)
  accessToken = null
  tokenExpiry = 0
  try {
    localStorage.removeItem('gd_token')
    localStorage.removeItem('gd_token_exp')
  } catch { /* ignore */ }
}

async function apiRequest(url, options = {}, allowRetry = true) {
  const token = await getToken()          // refreshes silently if needed
  const resp = await fetch(url, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
  })
  if (resp.status === 401 && allowRetry) {
    // Token rejected despite looking valid — force one silent refresh and retry.
    // Never revoke here.
    accessToken = null
    tokenExpiry = 0
    try { await requestToken('') } catch { throw new Error('TOKEN_EXPIRED') }
    return apiRequest(url, options, false)
  }
  if (resp.status === 401) throw new Error('TOKEN_EXPIRED')
  return resp
}

// ── DATA ──────────────────────────────────────────────────────────────────────
export async function findFile() {
  return { id: SHEET_ID, name: 'ENTRENAMIENTO ALEJANDRO TEMPORADA 2' }
}

// Export Google Sheet as .xlsx for parsing with SheetJS
export async function downloadFile() {
  const resp = await apiRequest(
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`
  )
  return resp.arrayBuffer()
}

// Write only specific cells via Sheets API — never touches formatting
export async function writeCells(sheetName, cellUpdates) {
  if (!cellUpdates || cellUpdates.length === 0) return

  const valueRanges = cellUpdates
    .filter(u => u.value !== '' && u.value != null)
    .map(u => ({
      range: `'${sheetName}'!${colToLetter(u.col)}${u.row + 1}`,
      values: [[u.value]],
    }))

  if (valueRanges.length === 0) return

  const resp = await apiRequest(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values:batchUpdate`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ valueInputOption: 'RAW', data: valueRanges }),
    }
  )

  if (!resp.ok) {
    const err = await resp.json()
    throw new Error(err.error?.message || 'Error al guardar')
  }
  return resp.json()
}

function colToLetter(col) {
  let letter = ''
  col += 1
  while (col > 0) {
    const rem = (col - 1) % 26
    letter = String.fromCharCode(65 + rem) + letter
    col = Math.floor((col - 1) / 26)
  }
  return letter
}
