import { useState, useEffect } from 'react'
import { signIn, signOut, isSignedIn, findFile, downloadFile, writeCells } from './drive.js'
import Tracker from './Tracker.jsx'

const C = { accent: '#C8F135', dark: '#0D0D0D', card: '#161616', muted: '#555', text: '#E8E8E8', red: '#FF6B6B' }

function withTimeout(promise, ms = 15000) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout — la conexión tardó demasiado')), ms))
  ])
}

export default function App() {
  const [authState, setAuthState] = useState('checking')
  const [error, setError] = useState(null)
  const [driveFile, setDriveFile] = useState(null)
  const [xlsxBuffer, setXlsxBuffer] = useState(null)

  useEffect(() => {
    if (isSignedIn()) loadFromDrive()
    else setAuthState('signed-out')
  }, [])

  async function handleSignIn() {
    try {
      setAuthState('loading'); setError(null)
      await withTimeout(signIn(), 20000)
      await loadFromDrive()
    } catch (e) {
      setError(e.message); setAuthState('signed-out')
    }
  }

  async function loadFromDrive() {
    try {
      setAuthState('loading')
      const file = await withTimeout(findFile())
      if (!file) throw new Error('No se encontró el archivo en Drive.')
      const buffer = await withTimeout(downloadFile(file.id), 20000)
      setDriveFile(file)
      setXlsxBuffer(buffer)
      setAuthState('ready')
    } catch (e) {
      if (e.message === 'TOKEN_EXPIRED') setAuthState('signed-out')
      else { setError(e.message); setAuthState('error') }
    }
  }

  async function handleSave(sheetName, cellUpdates) {
    await withTimeout(writeCells(sheetName, cellUpdates))
    const buffer = await withTimeout(downloadFile(driveFile.id), 20000)
    setXlsxBuffer(buffer)
  }

  if (authState === 'checking' || authState === 'loading')
    return <LoadingScreen message={authState === 'checking' ? 'Iniciando...' : 'Conectando con Drive...'} onCancel={() => { signOut(); setAuthState('signed-out') }} />
  if (authState === 'signed-out')
    return <SignInScreen onSignIn={handleSignIn} error={error} />
  if (authState === 'error')
    return <ErrorScreen message={error} onRetry={loadFromDrive} onSignOut={() => { signOut(); setAuthState('signed-out') }} />
  if (authState === 'ready' && xlsxBuffer)
    return <Tracker xlsxBuffer={xlsxBuffer} fileName={driveFile.name} onSave={handleSave} onSignOut={() => { signOut(); setAuthState('signed-out') }} />
  return null
}

function LoadingScreen({ message, onCancel }) {
  const [seconds, setSeconds] = useState(0)
  useEffect(() => {
    const t = setInterval(() => setSeconds(s => s + 1), 1000)
    return () => clearInterval(t)
  }, [])
  return (
    <div style={{ background: C.dark, minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', fontFamily: "'DM Mono','Courier New',monospace", color: C.text }}>
      <div style={{ fontSize: 10, color: C.muted, letterSpacing: 4, textTransform: 'uppercase', marginBottom: 6 }}>Tracker de entreno</div>
      <div style={{ fontSize: 26, fontWeight: 700, color: C.accent, letterSpacing: -1, marginBottom: 32 }}>ALEJANDRO</div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
        {[0,1,2].map(i => <div key={i} style={{ width: 8, height: 8, borderRadius: '50%', background: C.accent, animation: `pulse 1.2s ${i*0.2}s infinite` }} />)}
      </div>
      <div style={{ fontSize: 11, color: C.muted, marginBottom: 8 }}>{message}</div>
      {seconds > 5 && (
        <div style={{ marginTop: 24, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
          <div style={{ fontSize: 11, color: C.muted }}>Tardando más de lo normal...</div>
          <button onClick={onCancel} style={{ background: 'none', border: '1px solid #333', color: C.muted, fontFamily: 'inherit', fontSize: 12, padding: '10px 20px', borderRadius: 20, cursor: 'pointer' }}>
            Cancelar y reconectar
          </button>
        </div>
      )}
      <style>{`@keyframes pulse{0%,100%{opacity:.2}50%{opacity:1}}`}</style>
    </div>
  )
}

function SignInScreen({ onSignIn, error }) {
  return (
    <div style={{ background: C.dark, minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24, fontFamily: "'DM Mono','Courier New',monospace", color: C.text }}>
      <div style={{ fontSize: 10, color: C.muted, letterSpacing: 4, textTransform: 'uppercase', marginBottom: 6 }}>Tracker de entreno</div>
      <div style={{ fontSize: 26, fontWeight: 700, color: C.accent, letterSpacing: -1, marginBottom: 4 }}>ALEJANDRO</div>
      <div style={{ fontSize: 10, color: C.muted, marginBottom: 52 }}>Mesociclo I · FPARK</div>
      <button onClick={onSignIn} style={{ background: C.accent, color: C.dark, fontFamily: 'inherit', fontWeight: 700, fontSize: 13, padding: '16px 32px', borderRadius: 30, border: 'none', cursor: 'pointer', letterSpacing: 2 }}>
        CONECTAR CON GOOGLE DRIVE
      </button>
      {error && <div style={{ marginTop: 20, background: '#1A0A0A', border: '1px solid #4A2A2A', borderRadius: 10, padding: '12px 16px', maxWidth: 320, textAlign: 'center' }}>
        <div style={{ fontSize: 11, color: C.red }}>{error}</div>
      </div>}
      <div style={{ marginTop: 36, fontSize: 10, color: '#333', textAlign: 'center', lineHeight: 1.9, maxWidth: 280 }}>
        Solo accede al archivo de tu entreno.<br />No lee ni modifica ningún otro archivo.
      </div>
    </div>
  )
}

function ErrorScreen({ message, onRetry, onSignOut }) {
  return (
    <div style={{ background: C.dark, minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', fontFamily: "'DM Mono',monospace", color: C.text, padding: 24, textAlign: 'center' }}>
      <div style={{ fontSize: 36, marginBottom: 16 }}>⚠️</div>
      <div style={{ fontSize: 16, fontWeight: 700, color: C.red, marginBottom: 8 }}>Error al conectar</div>
      <div style={{ fontSize: 11, color: C.muted, maxWidth: 300, lineHeight: 1.8, marginBottom: 28 }}>{message}</div>
      <div style={{ display: 'flex', gap: 10 }}>
        <button onClick={onRetry} style={{ background: C.accent, color: C.dark, fontFamily: 'inherit', fontWeight: 700, fontSize: 12, padding: '12px 20px', borderRadius: 25, border: 'none', cursor: 'pointer' }}>REINTENTAR</button>
        <button onClick={onSignOut} style={{ background: 'none', border: '1px solid #333', color: C.muted, fontFamily: 'inherit', fontSize: 12, padding: '12px 20px', borderRadius: 25, cursor: 'pointer' }}>CERRAR SESIÓN</button>
      </div>
    </div>
  )
}
