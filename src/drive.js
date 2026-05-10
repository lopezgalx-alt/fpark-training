// ── GOOGLE DRIVE OAUTH ────────────────────────────────────────────────────────
// Uses Google Identity Services (GIS) for OAuth 2.0 PKCE flow.
// No backend needed — tokens stored in sessionStorage (cleared on tab close).

const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID
const SCOPES = 'https://www.googleapis.com/auth/drive.file'
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const FILE_NAME = 'ENTRENAMIENTO ALEJANDRO LOPEZ.xlsx'

let tokenClient = null
let accessToken = null

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
    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: (resp) => {
        if (resp.error) { reject(new Error(resp.error)); return }
        accessToken = resp.access_token
        sessionStorage.setItem('gd_token', accessToken)
        resolve(accessToken)
      },
    })
    // Check for saved token first
    const saved = sessionStorage.getItem('gd_token')
    if (saved) { accessToken = saved; resolve(saved); return }
    tokenClient.requestAccessToken({ prompt: 'consent' })
  })
}

export function isSignedIn() {
  const saved = sessionStorage.getItem('gd_token')
  if (saved) { accessToken = saved }
  return !!accessToken
}

export function signOut() {
  if (accessToken) {
    window.google?.accounts?.oauth2?.revoke(accessToken)
  }
  accessToken = null
  sessionStorage.removeItem('gd_token')
}

async function driveRequest(url, options = {}) {
  if (!accessToken) throw new Error('No autenticado')
  const resp = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.headers || {}),
    },
  })
  if (resp.status === 401) {
    // Token expired — clear and prompt re-auth
    signOut()
    throw new Error('TOKEN_EXPIRED')
  }
  return resp
}

// Find the training Excel file in Drive
export async function findFile() {
  const q = encodeURIComponent(`name='${FILE_NAME}' and mimeType='${XLSX_MIME}' and trashed=false`)
  const resp = await driveRequest(
    `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,modifiedTime)&orderBy=modifiedTime+desc`
  )
  const data = await resp.json()
  return data.files?.[0] || null
}

// Download file content as ArrayBuffer
export async function downloadFile(fileId) {
  const resp = await driveRequest(
    `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`
  )
  return resp.arrayBuffer()
}

// Upload (update) existing file with new content
export async function uploadFile(fileId, arrayBuffer, fileName) {
  const metadata = { name: fileName || FILE_NAME, mimeType: XLSX_MIME }

  const form = new FormData()
  form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }))
  form.append('file', new Blob([arrayBuffer], { type: XLSX_MIME }))

  const url = fileId
    ? `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=multipart`
    : `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart`

  const resp = await driveRequest(url, {
    method: fileId ? 'PATCH' : 'POST',
    body: form,
  })
  return resp.json()
}
