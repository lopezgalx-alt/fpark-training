import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { signIn, signOut, isSignedIn, findFile, downloadFile, uploadFile } from './drive.js'
import Tracker from './Tracker.jsx'

const C = { accent: '#C8F135', dark: '#0D0D0D', card: '#161616', muted: '#555', text: '#E8E8E8', red: '#FF6B6B' }

export default function App() {
  const [authState, setAuthState] = useState('checking') // checking | signed-out | loading | ready | error
  const [error, setError] = useState(null)
  const [driveFile, setDriveFile] = useState(null) // { id, name }
  const [xlsxBuffer, setXlsxBuffer] = useState(null)

  useEffect(() => {
    if (isSignedIn()) {
      loadFromDrive()
    } else {
      setAuthState('signed-out')
    }
  }, [])

  async function handleSignIn() {
    try {
      setAuthState('loading')
      setError(null)
      await signIn()
      await loadFromDrive()
    } catch (e) {
      setError(e.message)
      setAuthState('signed-out')
    }
  }

  async function loadFromDrive() {
    try {
      setAuthState('loading')
      const file = await findFile()
      if (!file) throw new Error(`No se encontró "${FILE_NAME}" en tu Drive. Asegúrate de que el archivo existe.`)
      const buffer = await downloadFile(file.id)
      setDriveFile(file)
      setXlsxBuffer(buffer)
      setAuthState('ready')
    } catch (e) {
      if (e.message === 'TOKEN_EXPIRED') {
        setAuthState('signed-out')
      } else {
        setError(e.message)
        setAuthState('error')
      }
    }
  }

  async function handleSave(updatedBuffer) {
    try {
      await uploadFile(driveFile.id, updatedBuffer, driveFile.name)
      // Reload to reflect saved data
      const buffer = await downloadFile(driveFile.id)
      setXlsxBuffer(buffer)
      return true
    } catch (e) {
      if (e.message === 'TOKEN_EXPIRED') {
        setAuthState('signed-out')
      }
      throw e
    }
  }

  if (authState === 'checking' || authState === 'loading') {
    return <LoadingScreen message={authState === 'checking' ? 'Iniciando...' : 'Conectando con Drive...'} />
  }

  if (authState === 'signed-out') {
    return <SignInScreen onSignIn={handleSignIn} error={error} />
  }

  if (authState === 'error') {
    return <ErrorScreen message={error} onRetry={loadFromDrive} onSignOut={() => { signOut(); setAuthState('signed-out') }} />
  }

  if (authState === 'ready' && xlsxBuffer) {
    return (
      <Tracker
        xlsxBuffer={xlsxBuffer}
        fileName={driveFile.name}
        onSave={handleSave}
        onSignOut={() => { signOut(); setAuthState('signed-out') }}
      />
    )
  }

  return null
}

function LoadingScreen({ message }) {
  return (
    <div style={{ background: C.dark, minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', fontFamily: "'DM Mono','Courier New',monospace", color: C.text }}>
      <div style={{ fontSize: 10, color: C.muted, letterSpacing: 4, textTransform: 'uppercase', marginBottom: 6 }}>Tracker de entreno</div>
      <div style={{ fontSize: 26, fontWeight: 700, color: C.accent, letterSpacing: -1, marginBottom: 32 }}>ALEJANDRO</div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
        {[0, 1, 2].map(i => (
          <div key={i} style={{ width: 8, height: 8, borderRadius: '50%', background: C.accent, animation: `pulse 1.2s ${i * 0.2}s infinite` }} />
        ))}
      </div>
      <div style={{ fontSize: 11, color: C.muted }}>{message}</div>
      <style>{`@keyframes pulse { 0%,100%{opacity:.2} 50%{opacity:1} }`}</style>
    </div>
  )
}

function SignInScreen({ onSignIn, error }) {
  return (
    <div style={{ background: C.dark, minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24, fontFamily: "'DM Mono','Courier New',monospace", color: C.text }}>
      <div style={{ fontSize: 10, color: C.muted, letterSpacing: 4, textTransform: 'uppercase', marginBottom: 6 }}>Tracker de entreno</div>
      <div style={{ fontSize: 26, fontWeight: 700, color: C.accent, letterSpacing: -1, marginBottom: 4 }}>ALEJANDRO</div>
      <div style={{ fontSize: 10, color: C.muted, marginBottom: 52 }}>Mesociclo I · FPARK</div>

      <button onClick={onSignIn}
        style={{ background: C.accent, color: C.dark, fontFamily: 'inherit', fontWeight: 700, fontSize: 13, padding: '16px 32px', borderRadius: 30, border: 'none', cursor: 'pointer', letterSpacing: 2, display: 'flex', alignItems: 'center', gap: 10 }}>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none">
          <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#4285F4"/>
          <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#34A853"/>
          <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l3.66-2.84z" fill="#FBBC05"/>
          <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" fill="#EA4335"/>
        </svg>
        CONECTAR CON GOOGLE DRIVE
      </button>

      {error && (
        <div style={{ marginTop: 20, background: '#1A0A0A', border: '1px solid #4A2A2A', borderRadius: 10, padding: '12px 16px', maxWidth: 320, textAlign: 'center' }}>
          <div style={{ fontSize: 11, color: C.red }}>{error}</div>
        </div>
      )}

      <div style={{ marginTop: 36, fontSize: 10, color: '#333', textAlign: 'center', lineHeight: 1.9, maxWidth: 280 }}>
        Solo accede al archivo de tu entreno.<br />
        No lee ni modifica ningún otro archivo de Drive.
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

const FILE_NAME = 'ENTRENAMIENTO ALEJANDRO LOPEZ.xlsx'
