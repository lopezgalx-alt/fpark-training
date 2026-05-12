const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.readonly'
const SHEET_ID = '1lb3OXjy5mpAd4dbcE7y93yz7I7RGpE4I7APi3aZZBcE'

let accessToken = null
let tokenClient = null

function loadGIS() {
  return new Promise((resolve) => {
    if (window.google?.accounts) { resolve(); return }
    const script = document.createElement('script')
    script.src = 'https://accounts.google.com/gsi/client'
    script.onload = resolve
    document.head.appendChild(script)
  })
}

function initTokenClient() {
  return new Promise((resolve) => {
    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: (resp) => resolve(resp),
    })
  })
}

function refreshToken(prompt = '') {
  return new Promise((resolve, reject) => {
    if (!tokenClient) { reject(new Error('Token client not initialized')); return }
    tokenClient.callback = (resp) => {
      if (resp.error) { reject(new Error(resp.error)); return }
      accessToken = resp.access_token
      localStorage.setItem('gd_token', accessToken)
      resolve(accessToken)
    }
    tokenClient.requestAccessToken({ prompt })
  })
}

export async function signIn() {
  await loadGIS()
  await initTokenClient()
  const saved = localStorage.getItem('gd_token')
  if (saved) { accessToken = saved; return saved }
  return refreshToken('consent')
}

export function isSignedIn() {
  const saved = localStorage.getItem('gd_token')
  if (saved) { accessToken = saved }
  return !!accessToken
}

export function signOut() {
  if (accessToken) window.google?.accounts?.oauth2?.revoke(accessToken)
  accessToken = null
  localStorage.removeItem('gd_token')
}

async function apiRequest(url, options = {}, retry = true) {
  if (!accessToken) throw new Error('No autenticado')
  const resp = await fetch(url, {
    ...options,
    headers: { Authorization: `Bearer ${accessToken}`, ...(options.headers || {}) },
  })
  if (resp.status === 401 && retry) {
    try {
      await loadGIS()
      if (!tokenClient) await initTokenClient()
      await refreshToken('')
      return apiRequest(url, options, false)
    } catch {
      signOut()
      throw new Error('TOKEN_EXPIRED')
    }
  }
  return resp
}

export async function findFile() {
  return { id: SHEET_ID, name: 'ENTRENAMIENTO ALEJANDRO LOPEZ' }
}

export async function downloadFile() {
  const resp = await apiRequest(
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`
  )
  if (!resp.ok) {
    const text = await resp.text()
    throw new Error(`Error ${resp.status}: ${text.slice(0, 100)}`)
  }
  return resp.arrayBuffer()
}

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
