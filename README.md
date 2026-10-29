# Plataforma Web de Carpetas y Entregables — CiberPaz 2026 (DocsShares Unipamplona)

Plataforma interactiva tipo **Google Drive** generada a partir de `URL.xlsx` (**632 archivos** y su jerarquía completa de carpetas en `/ENTREGABLES SEGUNDO DESEMBOLSO`), conectada directamente a cada URL de `docsshares.unipamplona.edu.co`.

## Características incluidas

1. **Explorador Interactivo Tipo Google Drive**:
   - **Árbol de carpetas lateral** desplegable con contadores de archivos por cada entregable y subcarpeta.
   - **Migas de pan (Breadcrumbs)** interactivas y botón para subir de nivel.
   - **Vista en Cuadrícula (Grid)** y **Vista en Lista Detallada (List)**.
   - **Modo Claro y Modo Oscuro**.
2. **Buscador Global Instantáneo (`Ctrl+K`) y Filtros por Tipo**:
   - Búsqueda en tiempo real insensible a tildes y mayúsculas (busca por municipio, nombre de archivo, líder de equipo, mes o entregable).
   - Filtros rápidos por tipo de archivo: **Todos (632)**, **PDF (461)**, **Imágenes (133)**, **Office Word/Excel/PPT (28)**, **Videos MP4 (4)** y **Comprimidos ZIP (6)**.
3. **Conexión Directa a cada URL + Vista Previa**:
   - Botón **Abrir URL** que lleva directamente al enlace de `docsshares.unipamplona.edu.co`.
   - Botón **Descarga directa** (`.../download`) y **Copiar URL**.
   - **Modal de Vista Previa y Detalles** con previsualización de imágenes/videos y navegación rápida a la carpeta contenedora.
4. **Sincronización con Google Sheets o Excel (`.xlsx` / `.csv`)**:
   - Ya incluye los **632 registros de `URL.xlsx`** pre-cargados en `data.js` (sin necesidad de configurar servidores externos ni credenciales).
   - Desde el botón **"Excel / Sheets"** en la barra superior puedes pegar el link de tu **Google Sheets** (compartido como *"Cualquier persona con el enlace puede ver"*) o arrastrar un nuevo archivo `.xlsx` para actualizar los datos al instante.
   - Si deseas dejar fijo un Google Sheet para todos los usuarios en Vercel, pega el enlace en la constante `DEFAULT_GOOGLE_SHEET_URL` en la línea 7 de `app.js`.

---

## Cómo abrirlo localmente

Puedes abrir directamente el archivo `index.html` haciendo doble clic sobre él en tu computador, o ejecutar un servidor local desde esta carpeta:

```bash
npx serve .
```

---

## Cómo subirlo a Vercel

### Opción 1: Desde GitHub (Recomendado)
1. Sube esta carpeta a un repositorio en GitHub (`index.html`, `styles.css`, `app.js`, `data.js`, `data.json`, `package.json`, `vercel.json`).
2. Entra a [vercel.com](https://vercel.com) → **Add New Project** → Importa tu repositorio.
3. Haz clic en **Deploy** (no requiere configuración adicional).

### Opción 2: Con Vercel CLI
Ejecuta en la terminal dentro de esta carpeta (`d:\USUARIO\Desktop\URLS`):
```bash
npx vercel
```

