// ── GOOGLE DRIVE OAUTH ────────────────────────────────────────────────────────
const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID
const SCOPES = 'https://www.googleapis.com/auth/drive.file'
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const FILE_NAME = 'ENTRENAMIENTO ALEJANDRO LOPEZ.xlsx'
const DRIVE_FILE_ID = '1WEMvbLfERQl0E09p7cyjObsIP_uMegzL'

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
    const saved = sessionStorage.getItem('gd_token')
    if (saved) { accessToken = saved; resolve(saved); return }
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
    tokenClient.requestAccessToken({ prompt: 'consent' })
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

async function driveRequest(url, options = {}) {
  if (!accessToken) throw new Error('No autenticado')
  const resp = await fetch(url, {
    ...options,
    headers: { Authorization: `Bearer ${accessToken}`, ...(options.headers || {}) },
  })
  if (resp.status === 401) { signOut(); throw new Error('TOKEN_EXPIRED') }
  return resp
}

export async function findFile() {
  const resp = await driveRequest(
    `https://www.googleapis.com/drive/v3/files/${DRIVE_FILE_ID}?fields=id,name,modifiedTime`
  )
  const data = await resp.json()
  return data.id ? { ...data, id: DRIVE_FILE_ID } : null
}

export async function downloadFile(fileId) {
  const resp = await driveRequest(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`)
  return resp.arrayBuffer()
}

export async function uploadFile(fileId, arrayBuffer) {
  const metadata = { name: FILE_NAME, mimeType: XLSX_MIME }
  const form = new FormData()
  form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }))
  form.append('file', new Blob([arrayBuffer], { type: XLSX_MIME }))
  const resp = await driveRequest(
    `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=multipart`,
    { method: 'PATCH', body: form }
  )
  return resp.json()
}
