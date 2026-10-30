import os
import re
import json

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(WORKSPACE_DIR, "Enlaces_DocsShares.csv")
DATA_JSON_PATH = os.path.join(WORKSPACE_DIR, "data.json")
DATA_JS_PATH = os.path.join(WORKSPACE_DIR, "data.js")
STYLES_PATH = os.path.join(WORKSPACE_DIR, "styles.css")
APP_JS_PATH = os.path.join(WORKSPACE_DIR, "app.js")
INDEX_HTML_PATH = os.path.join(WORKSPACE_DIR, "index.html")


def clean_text(t):
    if not t:
        return ""
    t = str(t).strip().replace("\ufffd", "ó")
    t = re.sub(r"ciber\s*paz", "ConTIC seguro", t, flags=re.IGNORECASE)
    return t


def read_csv_records(csv_path):
    records = []
    seen_paths = set()

    with open(csv_path, "r", encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            m = re.search(r'^(.*?)[,;]"?(/ENTREGABLES.*?)"?[,;]"?(https?://[^",\s]+)"?$', line)
            if not m:
                continue
            name, ruta, url = m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
            name = clean_text(name.strip('"'))
            ruta = clean_text(ruta.strip('"'))
            url = url.strip('"').strip()

            if ruta in seen_paths:
                continue
            seen_paths.add(ruta)

            parts = [p.strip() for p in ruta.split("/") if p.strip()]
            if parts and parts[-1].lower() == name.lower():
                folder_parts = parts[:-1]
            elif len(parts) > 1:
                folder_parts = parts[:-1]
            else:
                folder_parts = parts

            ext = name.rsplit(".", 1)[-1].strip().lower() if "." in name else "otro"

            records.append({
                "id": len(records) + 1,
                "name": name,
                "path": ruta,
                "folderPath": "/" + "/".join(folder_parts),
                "folders": folder_parts,
                "url": url,
                "ext": ext
            })
    return records


def build_self_contained_html(records):
    with open(STYLES_PATH, "r", encoding="utf-8") as f:
        css_content = f.read()
    with open(APP_JS_PATH, "r", encoding="utf-8") as f:
        js_content = f.read()

    data_json_str = json.dumps(records, ensure_ascii=False)
    total = len(records)

    html = f"""<!DOCTYPE html>
<html lang="es" data-theme="light">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Explorador de Carpetas y Entregables | DocsShares ConTIC seguro 2026</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js" defer></script>
  <style>
{css_content}
  </style>
</head>
<body>
  <div class="app-layout">
    <!-- Top Navigation Bar -->
    <header class="topbar">
      <div class="topbar-left">
        <button id="sidebarToggleBtn" class="icon-btn mobile-only" title="Mostrar/Ocultar árbol de carpetas" aria-label="Menú">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
        </button>
        <div class="brand" id="brandHomeBtn" title="Ir a la carpeta raíz">
          <div class="brand-icon">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
              <line x1="12" y1="11" x2="12" y2="17"></line>
              <line x1="9" y1="14" x2="15" y2="14"></line>
            </svg>
          </div>
          <div class="brand-text">
            <h1>Repositorio ConTIC seguro 2026</h1>
            <span class="brand-subtitle" id="dataSourceBadge">DocsShares Unipamplona · {total} archivos</span>
          </div>
        </div>
      </div>

      <!-- Search Bar -->
      <div class="search-container">
        <div class="search-input-wrapper">
          <svg class="search-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          <input
            type="text"
            id="searchInput"
            placeholder="Buscar carpetas o archivos por nombre, municipio, equipo... (Ctrl+K)"
            autocomplete="off"
          />
          <button id="clearSearchBtn" class="clear-search-btn hidden" title="Limpiar búsqueda">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
          <kbd class="search-shortcut">Ctrl+K</kbd>
        </div>
        <div class="search-scope-toggle">
          <label class="scope-label" title="Buscar en todas las carpetas o solo dentro de la carpeta actual">
            <input type="checkbox" id="searchCurrentFolderOnly" />
            <span>Solo carpeta actual</span>
          </label>
        </div>
      </div>

      <!-- Actions -->
      <div class="topbar-actions">
        <button id="openDataModalBtn" class="btn btn-secondary" title="Conectar Google Sheets o cargar archivo Excel (.xlsx)">
          <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="8" y1="13" x2="16" y2="13"></line>
            <line x1="8" y1="17" x2="16" y2="17"></line>
          </svg>
          <span>Excel / Sheets</span>
        </button>

        <button id="themeToggleBtn" class="icon-btn" title="Cambiar tema claro/oscuro" aria-label="Cambiar tema">
          <svg id="themeIconSun" class="hidden" width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>
          <svg id="themeIconMoon" width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>
        </button>
      </div>
    </header>

    <div class="main-body">
      <!-- Sidebar Tree Navigation -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-header">
          <span class="sidebar-title">ESTRUCTURA DE CARPETAS</span>
          <div class="sidebar-header-actions">
            <button id="expandAllTreeBtn" class="mini-btn" title="Expandir carpetas principales">Expandir</button>
            <button id="collapseAllTreeBtn" class="mini-btn" title="Contraer árbol">Contraer</button>
          </div>
        </div>
        <div class="tree-container" id="folderTree"></div>
        <div class="sidebar-footer">
          <div class="storage-summary">
            <div class="storage-row">
              <span>Total indexado</span>
              <strong id="sidebarTotalFiles">{total} archivos</strong>
            </div>
            <div class="storage-row">
              <span>Carpetas totales</span>
              <strong id="sidebarTotalFolders">0 carpetas</strong>
            </div>
          </div>
        </div>
      </aside>

      <div class="sidebar-backdrop hidden" id="sidebarBackdrop"></div>

      <!-- Main Explorer Content -->
      <main class="explorer-content">
        <section class="filters-bar" id="typeFiltersBar"></section>

        <section class="toolbar">
          <div class="breadcrumbs-wrapper">
            <button id="upFolderBtn" class="icon-btn up-btn" title="Subir un nivel de carpeta" disabled>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"/></svg>
            </button>
            <nav class="breadcrumbs" id="breadcrumbs" aria-label="Ruta de carpeta"></nav>
          </div>

          <div class="toolbar-controls">
            <div class="sort-control">
              <label for="sortSelect" class="sr-only">Ordenar por</label>
              <select id="sortSelect" class="select-input" title="Ordenar elementos">
                <option value="name-asc">Nombre (A → Z)</option>
                <option value="name-desc">Nombre (Z → A)</option>
                <option value="count-desc">Más archivos primero</option>
                <option value="type-asc">Tipo de archivo</option>
              </select>
            </div>

            <div class="view-switcher" role="group" aria-label="Modo de vista">
              <button id="gridViewBtn" class="view-btn active" title="Vista en cuadrícula (Iconos grandes)">
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
              </button>
              <button id="listViewBtn" class="view-btn" title="Vista en lista detallada">
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
              </button>
            </div>
          </div>
        </section>

        <section class="folder-info-banner" id="folderInfoBanner">
          <div class="folder-info-main">
            <h2 id="currentFolderTitle">ENTREGABLES SEGUNDO DESEMBOLSO</h2>
            <p id="currentFolderMeta">Cargando estructura de carpetas y archivos...</p>
          </div>
          <div class="folder-info-actions">
            <button id="copyFolderLinkBtn" class="btn btn-ghost btn-sm" title="Copiar enlace directo a esta carpeta">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>
              <span>Compartir vista</span>
            </button>
          </div>
        </section>

        <div class="explorer-scroll" id="explorerScroll">
          <section id="foldersSection" class="explorer-section">
            <div class="section-header">
              <h3>
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
                <span>Carpetas</span>
              </h3>
              <span class="section-badge" id="foldersCountBadge">0</span>
            </div>
            <div id="foldersContainer" class="items-grid"></div>
          </section>

          <section id="filesSection" class="explorer-section">
            <div class="section-header">
              <h3>
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                <span id="filesSectionTitle">Archivos</span>
              </h3>
              <span class="section-badge" id="filesCountBadge">0</span>
            </div>
            <div id="filesContainer" class="items-grid"></div>
          </section>

          <section id="emptyState" class="empty-state hidden">
            <div class="empty-icon">
              <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            </div>
            <h3>No se encontraron resultados</h3>
            <p id="emptyStateText">Intenta ajustar tu búsqueda o cambiar el filtro de tipo de archivo.</p>
            <button id="resetFiltersBtn" class="btn btn-primary">Ver todos los archivos</button>
          </section>
        </div>
      </main>
    </div>
  </div>

  <!-- File Details & Preview Modal -->
  <div class="modal-overlay hidden" id="fileModalOverlay">
    <div class="modal-card modal-lg" role="dialog" aria-modal="true" aria-labelledby="modalFileName">
      <div class="modal-header">
        <div class="modal-title-group">
          <span class="file-ext-pill" id="modalFileExt">PDF</span>
          <h3 id="modalFileName">Nombre del archivo</h3>
        </div>
        <button class="icon-btn" id="closeFileModalBtn" aria-label="Cerrar">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      </div>
      <div class="modal-body">
        <div class="file-meta-box">
          <div class="meta-item">
            <span class="meta-label">Ubicación en carpetas:</span>
            <div class="meta-path-chips" id="modalFolderChips"></div>
          </div>
          <div class="meta-item">
            <span class="meta-label">Ruta completa:</span>
            <code class="meta-code" id="modalFullPath"></code>
          </div>
          <div class="meta-item">
            <span class="meta-label">URL en DocsShares:</span>
            <div class="url-copy-row">
              <input type="text" id="modalFileUrlInput" readonly class="url-readonly-input" />
              <button id="modalCopyUrlBtn" class="btn btn-secondary btn-sm">Copiar URL</button>
            </div>
          </div>
        </div>

        <div class="preview-box" id="modalPreviewContainer"></div>
      </div>
      <div class="modal-footer">
        <button id="modalGoToFolderBtn" class="btn btn-ghost">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
          <span>Ir a su carpeta</span>
        </button>
        <div class="modal-footer-right">
          <a id="modalDownloadBtn" href="#" target="_blank" rel="noopener noreferrer" class="btn btn-secondary">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            <span>Descarga directa</span>
          </a>
          <a id="modalOpenExternalBtn" href="#" target="_blank" rel="noopener noreferrer" class="btn btn-primary">
            <span>Abrir en DocsShares</span>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
          </a>
        </div>
      </div>
    </div>
  </div>

  <!-- Data Source Modal -->
  <div class="modal-overlay hidden" id="dataModalOverlay">
    <div class="modal-card" role="dialog" aria-modal="true" aria-labelledby="dataModalTitle">
      <div class="modal-header">
        <div class="modal-title-group">
          <h3 id="dataModalTitle">Sincronizar con Google Sheets o Excel</h3>
        </div>
        <button class="icon-btn" id="closeDataModalBtn" aria-label="Cerrar">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      </div>
      <div class="modal-body">
        <p class="modal-desc">
          La plataforma ya incluye los <strong>{total} archivos de Enlaces_DocsShares.csv</strong> integrados directamente en este HTML listo para Vercel.
        </p>

        <div class="sync-panel">
          <h4>1. Conectar enlace de Google Sheets en vivo</h4>
          <p class="sync-hint">
            En Google Sheets, haz clic en <em>Compartir → Cualquier persona con el enlace (Lector)</em> y pega el link aquí:
          </p>
          <div class="input-action-row">
            <input
              type="url"
              id="googleSheetUrlInput"
              class="text-input"
              placeholder="https://docs.google.com/spreadsheets/d/1ABC.../edit"
            />
            <button id="syncSheetBtn" class="btn btn-primary">Conectar Sheet</button>
          </div>
          <p id="sheetSyncStatus" class="sync-status hidden"></p>
        </div>

        <hr class="divider" />

        <div class="sync-panel">
          <h4>2. Cargar un archivo Excel (.xlsx) o CSV local</h4>
          <p class="sync-hint">
            Selecciona o arrastra un archivo con las columnas <code>Nombre del archivo</code>, <code>Ruta del Archivo</code> y <code>URL</code>:
          </p>
          <label class="file-dropzone" id="excelDropzone">
            <input type="file" id="excelFileInput" accept=".xlsx,.xls,.csv" hidden />
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            <span>Haz clic para elegir archivo <strong>.xlsx / .csv</strong> o arrástralo aquí</span>
          </label>
        </div>

        <hr class="divider" />

        <div class="sync-panel sync-actions-row">
          <div>
            <h4>3. Datos integrados y exportación</h4>
            <p class="sync-hint">Puedes descargar el archivo <code>data.js</code> actualizado o volver a los {total} registros originales.</p>
          </div>
          <div class="sync-buttons">
            <button id="exportDataJsBtn" class="btn btn-secondary btn-sm">Descargar data.js</button>
            <button id="resetDefaultDataBtn" class="btn btn-ghost btn-sm">Restaurar original ({total})</button>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div id="toastContainer" class="toast-container"></div>

  <script>
window.INITIAL_FILES_DATA = {data_json_str};
  </script>
  <script>
{js_content}
  </script>
</body>
</html>
"""
    with open(INDEX_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)


def main():
    if not os.path.exists(CSV_PATH):
        print(f"No se encontró {CSV_PATH}")
        return

    records = read_csv_records(CSV_PATH)

    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    with open(DATA_JS_PATH, "w", encoding="utf-8") as f:
        f.write("window.INITIAL_FILES_DATA = ")
        json.dump(records, f, ensure_ascii=False, indent=2)
        f.write(";\n")

    build_self_contained_html(records)

    print(f"¡Listo! Se actualizaron index.html (todo-en-uno), data.js y data.json con {len(records)} archivos.")


if __name__ == "__main__":
    main()
