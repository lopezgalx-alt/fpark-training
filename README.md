# FPARK Training App — Deploy en Vercel

## Pasos para el deploy

### 1. Google OAuth Client ID

Ve a https://console.cloud.google.com y:

1. Crea un proyecto nuevo (o usa uno existente)
2. Ve a **APIs & Services → OAuth consent screen**
   - User type: External
   - App name: FPARK Training
   - Support email: tu email
   - Guarda
3. Ve a **APIs & Services → Credentials → Create Credentials → OAuth client ID**
   - Application type: **Web application**
   - Name: FPARK Training
   - Authorized JavaScript origins:
     - `http://localhost:5173` (para desarrollo)
     - `https://tu-app.vercel.app` (añadir después del deploy)
   - Authorized redirect URIs: (dejar vacío — usamos token flow)
   - Crear → Copia el **Client ID**
4. Ve a **APIs & Services → Library** → busca **Google Drive API** → Enable

### 2. Subir a GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/fpark-training.git
git push -u origin main
```

### 3. Deploy en Vercel

1. Ve a https://vercel.com → New Project → Import desde GitHub
2. En **Environment Variables** añade:
   ```
   VITE_GOOGLE_CLIENT_ID = tu-client-id.apps.googleusercontent.com
   ```
3. Deploy

### 4. Añadir tu URL de Vercel al OAuth

Una vez desplegado (ej: `https://fpark-training.vercel.app`):
1. Vuelve a Google Cloud Console → Credentials → tu OAuth client
2. Añade la URL en **Authorized JavaScript origins**
3. Guarda

### 5. Instalar en iPhone como PWA

1. Abre la URL en Safari
2. Toca el botón de compartir (cuadrado con flecha)
3. "Añadir a pantalla de inicio"
4. Ya la tienes como app nativa

## Desarrollo local

```bash
cp .env.example .env.local
# Rellena VITE_GOOGLE_CLIENT_ID en .env.local
npm install
npm run dev
```

## Nombre del archivo en Drive

La app busca automáticamente: `ENTRENAMIENTO ALEJANDRO LOPEZ.xlsx`
Asegúrate de que ese archivo existe en tu Drive con ese nombre exacto.
