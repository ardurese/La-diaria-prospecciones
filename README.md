# Relevamiento — La Diaria Equipamientos

App de relevamiento comercial conectada a Google Sheets, con descarga de histórico en Excel.

## Cómo publicarla (gratis, sin saber programar)

### Opción más simple: Vercel + GitHub (desde el celu o compu)

1. Subí esta carpeta a un repositorio de **GitHub** (podés crear una cuenta gratis en github.com)
   - La forma más fácil: entrá a github.com/new, creá un repo, y arrastrá/subí todos estos archivos desde "Add file → Upload files"
2. Entrá a **vercel.com** y creá una cuenta (podés entrar con tu cuenta de GitHub directo)
3. Tocá **"Add New" → "Project"**
4. Elegí el repositorio que subiste
5. Dejá toda la configuración en automático (Vercel detecta que es un proyecto Vite) y tocá **"Deploy"**
6. En 1-2 minutos te da una URL pública (algo como `la-diaria.vercel.app`)

Esa URL ya queda funcionando para siempre, gratis, y se puede agregar a la pantalla de inicio del celu como una app.

### Actualizaciones futuras

Cada vez que quieras cambiar algo (colores, textos, campos), se edita el archivo `src/App.jsx` y se sube de nuevo a GitHub — Vercel actualiza la web automáticamente sola.

## Qué hace esta app

- Formulario de relevamiento con asesor, cliente, zona, interés, competencia y foto
- Al guardar, envía los datos a tu Google Sheet vía Apps Script (ya configurado)
- Dashboard con conteo total y por asesor
- Botón para descargar todo el histórico cargado en la sesión como archivo Excel (.xlsx)

## Configuración ya incluida

El script de Google Apps ya está conectado en `src/App.jsx` (constante `SCRIPT_URL`). Si en algún momento necesitás cambiar de planilla, hay que:
1. Crear un nuevo script en script.google.com apuntando a la nueva planilla
2. Reemplazar la URL en esa constante
