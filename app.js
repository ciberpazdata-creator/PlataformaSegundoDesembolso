/**
 * Repositorio Interactivo de Carpetas y Archivos - ConTIC seguro 2026 (DocsShares Unipamplona)
 * Soporta datos pre-integrados (data.js), sincronización en vivo con Google Sheets y carga de Excel (.xlsx/.csv).
 */

// Si deseas dejar fijo un link de Google Sheets para todos los visitantes en Vercel, pégalo aquí:
const DEFAULT_GOOGLE_SHEET_URL = "";

// Estado global de la aplicación
const state = {
  rawRecords: [],
  rootNode: null,
  folderMap: new Map(), // path -> folderNode
  currentPath: "",
  expandedPaths: new Set(),
  searchQuery: "",
  searchCurrentOnly: false,
  activeTypeFilter: "all", // all | pdf | img | office | video | zip
  sortBy: "name-asc",      // name-asc | name-desc | count-desc | type-asc
  viewMode: localStorage.getItem("contic_seguro_view_mode") || "grid",
  dataSourceLabel: "URL.xlsx integrado",
  activeModalFile: null
};

// Categorías de tipos de archivo
const FILE_CATEGORIES = {
  all: { label: "Todos", exts: null },
  pdf: { label: "PDF", exts: ["pdf"] },
  img: { label: "Imágenes", exts: ["png", "jpg", "jpeg", "gif", "webp", "svg"] },
  office: { label: "Office (Word/Excel/PPT)", exts: ["docx", "doc", "xlsx", "xls", "pptx", "ppt"] },
  video: { label: "Videos", exts: ["mp4", "mov", "avi", "mkv", "webm"] },
  zip: { label: "Comprimidos (ZIP)", exts: ["zip", "rar", "7z"] }
};

// ============================================================================
// 1. LIMPIEZA DE TEXTO Y CODIFICACIÓN (TILDES / UTF-8)
// ============================================================================

const MOJIBAKE_MAP = [
  [/Ã¡/g, "á"], [/Ã©/g, "é"], [/Ã­/g, "í"], [/Ã³/g, "ó"], [/Ãº/g, "ú"],
  [/Ã\x81/g, "Á"], [/Ã\x89/g, "É"], [/Ã\x8d/g, "Í"], [/Ã\x93/g, "Ó"], [/Ã\x9a/g, "Ú"],
  [/Ã±/g, "ñ"], [/Ã\x91/g, "Ñ"], [/Ã¼/g, "ü"],
  [/ejecucin/gi, "ejecución"],
  [/implementacin/gi, "implementación"],
  [/orientacin/gi, "orientación"],
  [/prevencin/gi, "prevención"],
  [/articulacin/gi, "articulación"],
  [/pedaggic/gi, "pedagógic"],
  [/metodolgic/gi, "metodológic"],
  [/tcnic/gi, "técnic"],
  [/cpsulas/gi, "cápsulas"],
  [/GUAINA/g, "GUAINÍA"],
  [/INRIDA/g, "INÍRIDA"],
  [/INDGENA/g, "INDÍGENA"],
  [/ACCIN/g, "ACCIÓN"],
  [/ciber\s*paz/gi, "ConTIC seguro"]
];

function cleanSpanishText(str) {
  if (!str) return "";
  let s = String(str).trim();
  for (const [pattern, replacement] of MOJIBAKE_MAP) {
    s = s.replace(pattern, replacement);
  }
  if (s.includes("\ufffd")) {
    s = s
      .replace(/ci\ufffdn/g, "ción")
      .replace(/CI\ufffdN/g, "CIÓN")
      .replace(/GUAIN\ufffdA/g, "GUAINÍA")
      .replace(/IN\ufffdRIDA/g, "INÍRIDA")
      .replace(/IND\ufffdGENA/g, "INDÍGENA")
      .replace(/pedag\ufffdg/g, "pedagóg")
      .replace(/metodol\ufffdg/g, "metodológ")
      .replace(/t\ufffdcnic/g, "técnic")
      .replace(/c\ufffdpsula/g, "cápsula")
      .replace(/\ufffd/g, "ó");
  }
  return s;
}

// Normaliza texto para búsqueda insensible a tildes y mayúsculas
function normalizeForSearch(str) {
  return String(str || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function getFileExtension(filename) {
  const clean = String(filename || "").trim();
  const idx = clean.lastIndexOf(".");
  if (idx === -1 || idx === clean.length - 1) return "archivo";
  return clean.slice(idx + 1).trim().toLowerCase();
}

function getFileCategoryClass(ext) {
  const e = ext.toLowerCase();
  if (e === "pdf") return "ext-pdf";
  if (["png", "jpg", "jpeg", "gif", "webp", "svg"].includes(e)) return "ext-img";
  if (["doc", "docx"].includes(e)) return "ext-doc";
  if (["xls", "xlsx", "csv"].includes(e)) return "ext-xls";
  if (["mp4", "mov", "avi", "webm", "pptx", "ppt"].includes(e)) return "ext-vid";
  if (["zip", "rar", "7z"].includes(e)) return "ext-zip";
  return "ext-doc";
}

function matchesTypeFilter(ext, filterKey) {
  if (filterKey === "all") return true;
  const cat = FILE_CATEGORIES[filterKey];
  if (!cat || !cat.exts) return true;
  return cat.exts.includes(ext.toLowerCase());
}

// ============================================================================
// 2. CONSTRUCCIÓN DEL ÁRBOL DE CARPETAS Y NORMALIZACIÓN DE REGISTROS
// ============================================================================

function normalizeRecords(rawList) {
  const normalized = [];
  let idCounter = 1;

  for (const item of rawList) {
    let name = cleanSpanishText(item.name || item["Nombre del archivo"] || item.Column1 || "");
    let path = cleanSpanishText(item.path || item["Ruta del Archivo"] || item.Column2 || "");
    let url = String(item.url || item["URL"] || item.Column3 || "").trim();

    // Si viene en una sola columna separada por comas (como la hoja 2 del Excel)
    if ((!path || !url) && name.includes(",/") && name.includes("http")) {
      const m = name.match(/^(.*?),(\/.*?),(https?:\/\/\S+)$/);
      if (m) {
        name = cleanSpanishText(m[1]);
        path = cleanSpanishText(m[2]);
        url = m[3].trim();
      }
    }

    if (!name || !url || !url.startsWith("http")) continue;
    if (name.toLowerCase() === "nombre del archivo" || name.toLowerCase() === "column1") continue;

    const parts = path.split("/").map(p => cleanSpanishText(p)).filter(Boolean);
    let folderParts = parts;
    if (parts.length > 0 && parts[parts.length - 1].toLowerCase() === name.toLowerCase()) {
      folderParts = parts.slice(0, -1);
    } else if (parts.length > 1 && parts[parts.length - 1].includes(".")) {
      folderParts = parts.slice(0, -1);
    }

    const ext = getFileExtension(name);
    const folderPath = "/" + folderParts.join("/");

    normalized.push({
      id: idCounter++,
      name,
      path: path || `${folderPath}/${name}`,
      folderPath,
      folders: folderParts,
      url,
      ext,
      searchKey: normalizeForSearch(`${name} ${folderParts.join(" ")}`)
    });
  }

  return normalized;
}

function buildFolderTree(records) {
  const root = {
    name: "Raíz",
    path: "/",
    depth: 0,
    parentPath: null,
    childrenMap: new Map(),
    children: [],
    files: [],
    totalFiles: 0,
    totalSubfolders: 0
  };

  const folderMap = new Map();
  folderMap.set("/", root);

  for (const file of records) {
    let current = root;
    let accumPath = "";

    for (let i = 0; i < file.folders.length; i++) {
      const segment = file.folders[i];
      accumPath += "/" + segment;

      let child = current.childrenMap.get(segment);
      if (!child) {
        child = {
          name: segment,
          path: accumPath,
          depth: i + 1,
          parentPath: current.path,
          childrenMap: new Map(),
          children: [],
          files: [],
          totalFiles: 0,
          totalSubfolders: 0
        };
        current.childrenMap.set(segment, child);
        current.children.push(child);
        folderMap.set(accumPath, child);
      }
      current = child;
    }

    current.files.push(file);
  }

  // Calcular totales recursivos y ordenar carpetas naturalmente (1., 3., 6., 8., 10.)
  function computeStats(node) {
    node.children.sort((a, b) =>
      a.name.localeCompare(b.name, "es", { numeric: true, sensitivity: "base" })
    );
    node.files.sort((a, b) =>
      a.name.localeCompare(b.name, "es", { numeric: true, sensitivity: "base" })
    );

    let fileSum = node.files.length;
    let folderSum = node.children.length;

    for (const child of node.children) {
      computeStats(child);
      fileSum += child.totalFiles;
      folderSum += child.totalSubfolders;
    }

    node.totalFiles = fileSum;
    node.totalSubfolders = folderSum;
  }

  computeStats(root);
  return { root, folderMap };
}

// Obtiene todos los archivos recursivamente dentro de un nodo de carpeta
function collectRecursiveFiles(node, out = []) {
  for (const f of node.files) out.push(f);
  for (const c of node.children) collectRecursiveFiles(c, out);
  return out;
}

// Obtiene todas las subcarpetas recursivamente dentro de un nodo
function collectRecursiveFolders(node, out = []) {
  for (const c of node.children) {
    out.push(c);
    collectRecursiveFolders(c, out);
  }
  return out;
}

// ============================================================================
// 3. RENDERIZADO DE LA INTERFAZ
// ============================================================================

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Resalta coincidencias de búsqueda (incluyendo palabras con tildes)
function highlightText(text, query) {
  const safeText = String(text || "");
  if (!query || !query.trim()) return escapeHtml(safeText);

  const normQuery = normalizeForSearch(query.trim());
  if (!normQuery) return escapeHtml(safeText);

  const normText = normalizeForSearch(safeText);
  const idx = normText.indexOf(normQuery);
  if (idx === -1) return escapeHtml(safeText);

  const before = safeText.slice(0, idx);
  const match = safeText.slice(idx, idx + normQuery.length);
  const after = safeText.slice(idx + normQuery.length);
  return `${escapeHtml(before)}<mark class="hl">${escapeHtml(match)}</mark>${escapeHtml(after)}`;
}

function getEffectiveRootNode() {
  if (!state.rootNode) return null;
  // Si la raíz solo tiene una gran carpeta contenedora ("ENTREGABLES SEGUNDO DESEMBOLSO"), usarla como inicio
  if (state.rootNode.children.length === 1 && state.rootNode.files.length === 0) {
    return state.rootNode.children[0];
  }
  return state.rootNode;
}

function ensureExpandedAncestors(path) {
  if (!path || path === "/") return;
  const parts = path.split("/").filter(Boolean);
  let accum = "";
  for (const p of parts) {
    accum += "/" + p;
    state.expandedPaths.add(accum);
  }
}

function renderSidebarTree() {
  const container = document.getElementById("folderTree");
  if (!state.rootNode) return;

  const startNode = getEffectiveRootNode();
  const nodesToRender = startNode === state.rootNode ? state.rootNode.children : [startNode];

  function buildNodeHtml(node) {
    const isExpanded = state.expandedPaths.has(node.path);
    const isActive = state.currentPath === node.path;
    const hasChildren = node.children.length > 0;

    let html = `<div class="tree-node">`;
    html += `
      <div class="tree-row ${isActive ? "active" : ""}" data-path="${escapeHtml(node.path)}" title="${escapeHtml(node.name)} (${node.totalFiles} archivos)">
        <span class="tree-toggle ${hasChildren ? (isExpanded ? "expanded" : "") : "empty"}" data-toggle-path="${escapeHtml(node.path)}">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="9 18 15 12 9 6"/></svg>
        </span>
        <span class="tree-folder-icon">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="1.5"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
        </span>
        <span class="tree-label">${escapeHtml(node.name)}</span>
        <span class="tree-count">${node.totalFiles}</span>
      </div>
    `;

    if (hasChildren && isExpanded) {
      html += `<div class="tree-children">`;
      for (const child of node.children) {
        html += buildNodeHtml(child);
      }
      html += `</div>`;
    }

    html += `</div>`;
    return html;
  }

  container.innerHTML = nodesToRender.map(buildNodeHtml).join("");

  // Actualizar contadores globales del pie de barra lateral
  document.getElementById("sidebarTotalFiles").textContent = `${state.rawRecords.length} archivos`;
  document.getElementById("sidebarTotalFolders").textContent = `${state.folderMap.size - 1} carpetas`;
  document.getElementById("dataSourceBadge").textContent = `${state.dataSourceLabel} · ${state.rawRecords.length} archivos`;
}

function renderFilterChips(baseFilesForCounts) {
  const bar = document.getElementById("typeFiltersBar");
  const counts = {
    all: baseFilesForCounts.length,
    pdf: 0,
    img: 0,
    office: 0,
    video: 0,
    zip: 0
  };

  for (const f of baseFilesForCounts) {
    for (const key of ["pdf", "img", "office", "video", "zip"]) {
      if (matchesTypeFilter(f.ext, key)) counts[key]++;
    }
  }

  bar.innerHTML = Object.entries(FILE_CATEGORIES)
    .map(([key, cfg]) => {
      const activeClass = state.activeTypeFilter === key ? "active" : "";
      return `
        <button class="filter-chip ${activeClass}" data-filter="${key}">
          <span>${escapeHtml(cfg.label)}</span>
          <span class="filter-chip-count">${counts[key]}</span>
        </button>
      `;
    })
    .join("");
}

function renderBreadcrumbs(currentNode) {
  const nav = document.getElementById("breadcrumbs");
  const upBtn = document.getElementById("upFolderBtn");
  const effectiveRoot = getEffectiveRootNode();

  const parts = currentNode.path.split("/").filter(Boolean);
  let accum = "";
  const crumbs = [];

  for (let i = 0; i < parts.length; i++) {
    accum += "/" + parts[i];
    crumbs.push({
      name: parts[i],
      path: accum,
      isCurrent: i === parts.length - 1
    });
  }

  if (crumbs.length === 0) {
    crumbs.push({ name: "Inicio", path: "/", isCurrent: true });
  }

  nav.innerHTML = crumbs
    .map((c, idx) => {
      const sep = idx > 0 ? `<span class="crumb-sep">/</span>` : "";
      return `${sep}<button class="crumb-btn ${c.isCurrent ? "current" : ""}" data-path="${escapeHtml(c.path)}" title="${escapeHtml(c.name)}">${escapeHtml(c.name)}</button>`;
    })
    .join("");

  // Botón subir nivel
  const canGoUp = currentNode.path !== "/" && currentNode.path !== effectiveRoot.path;
  upBtn.disabled = !canGoUp;
}

function sortFolders(folders) {
  const list = [...folders];
  if (state.sortBy === "name-desc") {
    list.sort((a, b) => b.name.localeCompare(a.name, "es", { numeric: true, sensitivity: "base" }));
  } else if (state.sortBy === "count-desc") {
    list.sort((a, b) => b.totalFiles - a.totalFiles || a.name.localeCompare(b.name, "es", { numeric: true }));
  } else {
    list.sort((a, b) => a.name.localeCompare(b.name, "es", { numeric: true, sensitivity: "base" }));
  }
  return list;
}

function sortFiles(files) {
  const list = [...files];
  if (state.sortBy === "name-desc") {
    list.sort((a, b) => b.name.localeCompare(a.name, "es", { numeric: true, sensitivity: "base" }));
  } else if (state.sortBy === "type-asc") {
    list.sort((a, b) => a.ext.localeCompare(b.ext) || a.name.localeCompare(b.name, "es", { numeric: true }));
  } else {
    list.sort((a, b) => a.name.localeCompare(b.name, "es", { numeric: true, sensitivity: "base" }));
  }
  return list;
}

function renderExplorer() {
  const currentNode = state.folderMap.get(state.currentPath) || getEffectiveRootNode() || state.rootNode;
  if (!currentNode) return;

  renderBreadcrumbs(currentNode);

  const q = state.searchQuery.trim();
  const normQ = normalizeForSearch(q);
  const isSearching = normQ.length > 0;
  const isFilteringType = state.activeTypeFilter !== "all";

  // Determinar el conjunto base de archivos para los contadores de chips de filtro
  const scopeBaseFiles = isSearching
    ? (state.searchCurrentOnly ? collectRecursiveFiles(currentNode) : state.rawRecords).filter(f =>
        f.searchKey.includes(normQ)
      )
    : collectRecursiveFiles(currentNode);

  renderFilterChips(scopeBaseFiles);

  // Determinar carpetas y archivos a mostrar
  let visibleFolders = [];
  let visibleFiles = [];

  if (isSearching) {
    const folderPool = state.searchCurrentOnly
      ? collectRecursiveFolders(currentNode)
      : collectRecursiveFolders(state.rootNode);

    visibleFolders = folderPool.filter(folder =>
      normalizeForSearch(folder.name).includes(normQ)
    );

    const filePool = state.searchCurrentOnly
      ? collectRecursiveFiles(currentNode)
      : state.rawRecords;

    visibleFiles = filePool.filter(
      f => f.searchKey.includes(normQ) && matchesTypeFilter(f.ext, state.activeTypeFilter)
    );
  } else if (isFilteringType) {
    // Si hay un filtro de tipo activo (ej. Videos o Excel), mostrar subcarpetas que tengan archivos de ese tipo + todos los archivos de ese tipo en la rama actual
    visibleFolders = currentNode.children.filter(child =>
      collectRecursiveFiles(child).some(f => matchesTypeFilter(f.ext, state.activeTypeFilter))
    );
    visibleFiles = collectRecursiveFiles(currentNode).filter(f =>
      matchesTypeFilter(f.ext, state.activeTypeFilter)
    );
  } else {
    // Navegación normal por carpeta
    visibleFolders = currentNode.children;
    visibleFiles = currentNode.files;
  }

  visibleFolders = sortFolders(visibleFolders);
  visibleFiles = sortFiles(visibleFiles);

  // Actualizar banner superior de carpeta
  const bannerTitle = document.getElementById("currentFolderTitle");
  const bannerMeta = document.getElementById("currentFolderMeta");
  const filesSectionTitle = document.getElementById("filesSectionTitle");

  if (isSearching) {
    bannerTitle.textContent = `Resultados para "${q}"`;
    bannerMeta.textContent = `Se encontraron ${visibleFolders.length} carpetas y ${visibleFiles.length} archivos (${state.searchCurrentOnly ? "en esta carpeta" : "en todo el repositorio"})`;
    filesSectionTitle.textContent = "Archivos encontrados";
  } else if (isFilteringType) {
    const filterName = FILE_CATEGORIES[state.activeTypeFilter]?.label || state.activeTypeFilter;
    bannerTitle.textContent = `${currentNode.name} · Filtro: ${filterName}`;
    bannerMeta.textContent = `Mostrando ${visibleFiles.length} archivos de tipo ${filterName} dentro de esta ubicación y sus subcarpetas`;
    filesSectionTitle.textContent = `Archivos (${filterName})`;
  } else {
    bannerTitle.textContent = currentNode.name;
    bannerMeta.textContent = `${currentNode.children.length} subcarpetas directas · ${currentNode.files.length} archivos en esta carpeta · ${currentNode.totalFiles} archivos en total incluyendo subcarpetas`;
    filesSectionTitle.textContent = "Archivos en esta carpeta";
  }

  // Contenedores y modo de vista (grid vs list)
  const foldersSection = document.getElementById("foldersSection");
  const filesSection = document.getElementById("filesSection");
  const foldersContainer = document.getElementById("foldersContainer");
  const filesContainer = document.getElementById("filesContainer");
  const emptyState = document.getElementById("emptyState");

  const viewClass = state.viewMode === "list" ? "items-list" : "items-grid";
  foldersContainer.className = viewClass;
  filesContainer.className = viewClass;

  document.getElementById("foldersCountBadge").textContent = visibleFolders.length;
  document.getElementById("filesCountBadge").textContent = visibleFiles.length;

  // Renderizar tarjetas de carpetas
  if (visibleFolders.length > 0) {
    foldersSection.classList.remove("hidden");
    foldersContainer.innerHTML = visibleFolders
      .map(folder => {
        const subCountText =
          folder.children.length > 0 ? `${folder.children.length} subcarpetas · ` : "";
        const parentHint =
          isSearching && folder.parentPath && folder.parentPath !== "/"
            ? `<div class="file-card-subpath" title="${escapeHtml(folder.path)}">${escapeHtml(folder.parentPath.replace(/^\/ENTREGABLES SEGUNDO DESEMBOLSO\/?/, ""))}</div>`
            : "";

        return `
          <article class="folder-card" data-folder-path="${escapeHtml(folder.path)}" tabindex="0" role="button" aria-label="Abrir carpeta ${escapeHtml(folder.name)}">
            <div class="folder-icon-box">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="1.5"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
            </div>
            <div class="folder-card-body">
              <div class="folder-card-title" title="${escapeHtml(folder.name)}">${highlightText(folder.name, q)}</div>
              ${parentHint}
              <div class="folder-card-meta">
                <span>${subCountText}<strong>${folder.totalFiles}</strong> ${folder.totalFiles === 1 ? "archivo" : "archivos"}</span>
              </div>
            </div>
          </article>
        `;
      })
      .join("");
  } else {
    foldersSection.classList.add("hidden");
    foldersContainer.innerHTML = "";
  }

  // Renderizar tarjetas de archivos
  if (visibleFiles.length > 0) {
    filesSection.classList.remove("hidden");
    filesContainer.innerHTML = visibleFiles
      .map(file => {
        const extClass = getFileCategoryClass(file.ext);
        const relativeFolder =
          file.folders.slice(1).join(" / ") || file.folders[0] || "Raíz";
        const downloadUrl = file.url.endsWith("/download")
          ? file.url
          : `${file.url.replace(/\/$/, "")}/download`;

        return `
          <article class="file-card" data-file-id="${file.id}">
            <div class="file-card-top" data-open-modal="${file.id}" title="Ver detalles y vista previa de ${escapeHtml(file.name)}">
              <div class="file-icon-box ${extClass}">
                <span>${escapeHtml(file.ext.slice(0, 4))}</span>
              </div>
              <div class="file-card-info">
                <div class="file-card-name" title="${escapeHtml(file.name)}">${highlightText(file.name, q)}</div>
                <div class="file-card-subpath" title="${escapeHtml(relativeFolder)}">${highlightText(relativeFolder, q)}</div>
              </div>
            </div>
            <div class="file-card-actions">
              <a href="${escapeHtml(file.url)}" target="_blank" rel="noopener noreferrer" class="file-action-link" title="Abrir URL en DocsShares Unipamplona">
                <span>Abrir URL</span>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
              </a>
              <div class="file-action-icons">
                <button class="icon-btn" data-open-modal="${file.id}" title="Vista previa y detalles">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
                </button>
                <a href="${escapeHtml(downloadUrl)}" target="_blank" rel="noopener noreferrer" class="icon-btn" title="Descarga directa">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                </a>
                <button class="icon-btn" data-copy-url="${escapeHtml(file.url)}" title="Copiar URL del archivo">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                </button>
              </div>
            </div>
          </article>
        `;
      })
      .join("");
  } else {
    filesSection.classList.add("hidden");
    filesContainer.innerHTML = "";
  }

  // Estado vacío si no hay ni carpetas ni archivos
  if (visibleFolders.length === 0 && visibleFiles.length === 0) {
    emptyState.classList.remove("hidden");
  } else {
    emptyState.classList.add("hidden");
  }
}

// ============================================================================
// 4. NAVEGACIÓN Y MODAL DE VISTA PREVIA DE ARCHIVOS
// ============================================================================

function navigateToFolder(path, pushHash = true) {
  if (!state.folderMap.has(path)) {
    const eff = getEffectiveRootNode();
    path = eff ? eff.path : "/";
  }
  state.currentPath = path;
  ensureExpandedAncestors(path);

  if (pushHash) {
    history.replaceState(null, "", `#folder=${encodeURIComponent(path)}`);
  }

  renderSidebarTree();
  renderExplorer();
  document.getElementById("explorerScroll").scrollTop = 0;

  // Cerrar sidebar en móvil
  document.getElementById("sidebar").classList.remove("open");
  document.getElementById("sidebarBackdrop").classList.add("hidden");
}

function openFileModal(fileId) {
  const file = state.rawRecords.find(r => r.id === Number(fileId));
  if (!file) return;
  state.activeModalFile = file;

  document.getElementById("modalFileExt").textContent = file.ext.toUpperCase();
  document.getElementById("modalFileName").textContent = file.name;
  document.getElementById("modalFullPath").textContent = file.path;
  document.getElementById("modalFileUrlInput").value = file.url;

  const downloadUrl = file.url.endsWith("/download")
    ? file.url
    : `${file.url.replace(/\/$/, "")}/download`;
  document.getElementById("modalDownloadBtn").href = downloadUrl;
  document.getElementById("modalOpenExternalBtn").href = file.url;

  // Chips de ruta interactivos
  let accum = "";
  const chipsHtml = file.folders
    .map(segment => {
      accum += "/" + segment;
      return `<button class="path-chip" data-modal-folder="${escapeHtml(accum)}">📁 ${escapeHtml(segment)}</button>`;
    })
    .join("");
  document.getElementById("modalFolderChips").innerHTML = chipsHtml;

  // Vista previa inteligente según extensión
  const previewContainer = document.getElementById("modalPreviewContainer");
  const ext = file.ext.toLowerCase();

  if (["png", "jpg", "jpeg", "gif", "webp"].includes(ext)) {
    previewContainer.innerHTML = `
      <img src="${escapeHtml(downloadUrl)}" alt="${escapeHtml(file.name)}" onerror="this.outerHTML='<p style=\\'font-size:13px;color:var(--text-muted)\\'>Haz clic en <strong>Abrir en DocsShares</strong> o <strong>Descarga directa</strong> para visualizar esta imagen.</p>'" />
    `;
  } else if (ext === "mp4" || ext === "webm") {
    previewContainer.innerHTML = `
      <video controls style="max-width:100%;max-height:340px;border-radius:8px;">
        <source src="${escapeHtml(downloadUrl)}" type="video/${ext}">
      </video>
    `;
  } else {
    previewContainer.innerHTML = `
      <div style="display:flex;flex-direction:column;align-items:center;gap:10px;padding:12px;">
        <div class="file-icon-box ${getFileCategoryClass(ext)}" style="width:54px;height:54px;font-size:13px;">
          ${escapeHtml(ext)}
        </div>
        <p style="font-size:13.5px;font-weight:600;color:var(--text-primary);">${escapeHtml(file.name)}</p>
        <p style="font-size:12.5px;color:var(--text-muted);max-width:440px;">
          Este documento está alojado en el servidor seguro <strong>DocsShares Unipamplona</strong>. Puedes abrir su visor oficial o descargarlo directamente con los botones inferiores.
        </p>
      </div>
    `;
  }

  document.getElementById("fileModalOverlay").classList.remove("hidden");
}

function showToast(message) {
  const container = document.getElementById("toastContainer");
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => {
    toast.remove();
  }, 2600);
}

async function copyToClipboard(text, label = "URL copiada al portapapeles") {
  try {
    await navigator.clipboard.writeText(text);
    showToast(label);
  } catch {
    const input = document.createElement("input");
    input.value = text;
    document.body.appendChild(input);
    input.select();
    document.execCommand("copy");
    input.remove();
    showToast(label);
  }
}

// ============================================================================
// 5. SINCRONIZACIÓN CON GOOGLE SHEETS Y ARCHIVOS EXCEL (.XLSX / .CSV)
// ============================================================================

function loadDataset(rawArray, sourceLabel = "URL.xlsx integrado") {
  const cleaned = normalizeRecords(rawArray);
  if (cleaned.length === 0) {
    showToast("No se encontraron filas válidas con URLs en el archivo.");
    return false;
  }

  state.rawRecords = cleaned;
  state.dataSourceLabel = sourceLabel;
  const { root, folderMap } = buildFolderTree(cleaned);
  state.rootNode = root;
  state.folderMap = folderMap;

  const effectiveRoot = getEffectiveRootNode();
  state.expandedPaths.clear();
  if (effectiveRoot) {
    state.expandedPaths.add(effectiveRoot.path);
  }

  // Verificar si hay carpeta en el hash de la URL
  const hashMatch = location.hash.match(/#folder=([^&]+)/);
  const requestedPath = hashMatch ? decodeURIComponent(hashMatch[1]) : "";

  if (requestedPath && folderMap.has(requestedPath)) {
    navigateToFolder(requestedPath, false);
  } else {
    navigateToFolder(effectiveRoot ? effectiveRoot.path : "/", false);
  }

  return true;
}

function buildGoogleSheetCsvUrl(inputUrl) {
  const trimmed = String(inputUrl || "").trim();
  if (!trimmed) return null;
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) return null;
  const sheetId = match[1];
  const gidMatch = trimmed.match(/[#&?]gid=([0-9]+)/);
  const gidParam = gidMatch ? `&gid=${gidMatch[1]}` : "";
  return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv${gidParam}`;
}

async function syncFromGoogleSheets(sheetUrl, silent = false) {
  const statusEl = document.getElementById("sheetSyncStatus");
  const csvUrl = buildGoogleSheetCsvUrl(sheetUrl);

  if (!csvUrl) {
    if (!silent) {
      statusEl.classList.remove("hidden");
      statusEl.style.color = "var(--pdf-color)";
      statusEl.textContent = "URL no válida. Asegúrate de pegar un enlace de Google Sheets (docs.google.com/spreadsheets/d/...).";
    }
    return;
  }

  if (!silent) {
    statusEl.classList.remove("hidden");
    statusEl.style.color = "var(--text-secondary)";
    statusEl.textContent = "Conectando con Google Sheets y descargando registros...";
  }

  try {
    const response = await fetch(csvUrl);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const csvText = await response.text();

    if (!window.XLSX) throw new Error("Librería XLSX aún cargando.");
    const workbook = window.XLSX.read(csvText, { type: "string", raw: true });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = window.XLSX.utils.sheet_to_json(firstSheet, { header: ["name", "path", "url"], defval: "" });

    const ok = loadDataset(rows, "Google Sheets en vivo");
    if (ok) {
      localStorage.setItem("contic_seguro_google_sheet_url", sheetUrl);
      if (!silent) {
        statusEl.style.color = "var(--xls-color)";
        statusEl.textContent = `¡Sincronizado con éxito! (${state.rawRecords.length} archivos cargados desde Google Sheets).`;
        showToast(`Sincronizado: ${state.rawRecords.length} archivos desde Google Sheets`);
      }
    }
  } catch (err) {
    if (!silent) {
      statusEl.classList.remove("hidden");
      statusEl.style.color = "var(--pdf-color)";
      statusEl.textContent = "No se pudo leer la hoja. Verifica que en Google Sheets esté compartido como 'Cualquier persona con el enlace puede ver'.";
    }
  }
}

function handleExcelFileUpload(file) {
  if (!file || !window.XLSX) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = window.XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = window.XLSX.utils.sheet_to_json(sheet, { header: ["name", "path", "url"], defval: "" });

      const ok = loadDataset(rows, `Excel: ${file.name}`);
      if (ok) {
        document.getElementById("dataModalOverlay").classList.add("hidden");
        showToast(`Archivo cargado: ${state.rawRecords.length} enlaces procesados`);
      }
    } catch (err) {
      showToast("Error al procesar el archivo Excel.");
    }
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================================
// 6. EVENTOS E INICIALIZACIÓN
// ============================================================================

function initTheme() {
  const savedTheme = localStorage.getItem("contic_seguro_theme") || "light";
  document.documentElement.setAttribute("data-theme", savedTheme);
  updateThemeIcons(savedTheme);
}

function updateThemeIcons(theme) {
  const sun = document.getElementById("themeIconSun");
  const moon = document.getElementById("themeIconMoon");
  if (theme === "dark") {
    sun.classList.remove("hidden");
    moon.classList.add("hidden");
  } else {
    sun.classList.add("hidden");
    moon.classList.remove("hidden");
  }
}

function initEventListeners() {
  // Click en el árbol lateral
  document.getElementById("folderTree").addEventListener("click", e => {
    const toggleBtn = e.target.closest("[data-toggle-path]");
    if (toggleBtn) {
      e.stopPropagation();
      const path = toggleBtn.getAttribute("data-toggle-path");
      if (state.expandedPaths.has(path)) {
        state.expandedPaths.delete(path);
      } else {
        state.expandedPaths.add(path);
      }
      renderSidebarTree();
      return;
    }

    const row = e.target.closest("[data-path]");
    if (row) {
      const path = row.getAttribute("data-path");
      state.expandedPaths.add(path);
      navigateToFolder(path);
    }
  });

  // Expandir / Contraer árbol
  document.getElementById("expandAllTreeBtn").addEventListener("click", () => {
    for (const path of state.folderMap.keys()) {
      const node = state.folderMap.get(path);
      if (node && node.depth <= 2) state.expandedPaths.add(path);
    }
    renderSidebarTree();
  });

  document.getElementById("collapseAllTreeBtn").addEventListener("click", () => {
    state.expandedPaths.clear();
    const eff = getEffectiveRootNode();
    if (eff) state.expandedPaths.add(eff.path);
    renderSidebarTree();
  });

  // Breadcrumbs y subir carpeta
  document.getElementById("breadcrumbs").addEventListener("click", e => {
    const btn = e.target.closest("[data-path]");
    if (btn) navigateToFolder(btn.getAttribute("data-path"));
  });

  document.getElementById("upFolderBtn").addEventListener("click", () => {
    const current = state.folderMap.get(state.currentPath);
    if (current && current.parentPath) {
      navigateToFolder(current.parentPath);
    }
  });

  document.getElementById("brandHomeBtn").addEventListener("click", () => {
    state.searchQuery = "";
    state.activeTypeFilter = "all";
    document.getElementById("searchInput").value = "";
    document.getElementById("clearSearchBtn").classList.add("hidden");
    const eff = getEffectiveRootNode();
    navigateToFolder(eff ? eff.path : "/");
  });

  // Click en tarjetas de carpeta
  document.getElementById("foldersContainer").addEventListener("click", e => {
    const card = e.target.closest("[data-folder-path]");
    if (card) {
      const targetPath = card.getAttribute("data-folder-path");
      // Si estábamos buscando, limpiar búsqueda al entrar a la carpeta seleccionada
      if (state.searchQuery) {
        state.searchQuery = "";
        document.getElementById("searchInput").value = "";
        document.getElementById("clearSearchBtn").classList.add("hidden");
      }
      navigateToFolder(targetPath);
    }
  });

  // Click en tarjetas de archivos (abrir modal o copiar URL)
  document.getElementById("filesContainer").addEventListener("click", e => {
    const copyBtn = e.target.closest("[data-copy-url]");
    if (copyBtn) {
      e.stopPropagation();
      copyToClipboard(copyBtn.getAttribute("data-copy-url"));
      return;
    }

    const modalTrigger = e.target.closest("[data-open-modal]");
    if (modalTrigger) {
      openFileModal(modalTrigger.getAttribute("data-open-modal"));
    }
  });

  // Filtros de tipo de archivo
  document.getElementById("typeFiltersBar").addEventListener("click", e => {
    const chip = e.target.closest("[data-filter]");
    if (chip) {
      state.activeTypeFilter = chip.getAttribute("data-filter");
      renderExplorer();
    }
  });

  // Buscador instantáneo
  const searchInput = document.getElementById("searchInput");
  const clearSearchBtn = document.getElementById("clearSearchBtn");

  searchInput.addEventListener("input", () => {
    state.searchQuery = searchInput.value;
    clearSearchBtn.classList.toggle("hidden", !state.searchQuery);
    renderExplorer();
  });

  clearSearchBtn.addEventListener("click", () => {
    state.searchQuery = "";
    searchInput.value = "";
    clearSearchBtn.classList.add("hidden");
    renderExplorer();
    searchInput.focus();
  });

  document.getElementById("searchCurrentFolderOnly").addEventListener("change", e => {
    state.searchCurrentOnly = e.target.checked;
    if (state.searchQuery) renderExplorer();
  });

  // Atajo de teclado Ctrl+K o /
  window.addEventListener("keydown", e => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
      e.preventDefault();
      searchInput.focus();
      searchInput.select();
    } else if (e.key === "Escape") {
      document.getElementById("fileModalOverlay").classList.add("hidden");
      document.getElementById("dataModalOverlay").classList.add("hidden");
    }
  });

  // Ordenamiento y cambio de vista (Cuadrícula / Lista)
  document.getElementById("sortSelect").addEventListener("change", e => {
    state.sortBy = e.target.value;
    renderExplorer();
  });

  const gridBtn = document.getElementById("gridViewBtn");
  const listBtn = document.getElementById("listViewBtn");

  function setViewMode(mode) {
    state.viewMode = mode;
    localStorage.setItem("contic_seguro_view_mode", mode);
    gridBtn.classList.toggle("active", mode === "grid");
    listBtn.classList.toggle("active", mode === "list");
    renderExplorer();
  }

  gridBtn.classList.toggle("active", state.viewMode === "grid");
  listBtn.classList.toggle("active", state.viewMode === "list");
  gridBtn.addEventListener("click", () => setViewMode("grid"));
  listBtn.addEventListener("click", () => setViewMode("list"));

  // Botón restablecer filtros en estado vacío
  document.getElementById("resetFiltersBtn").addEventListener("click", () => {
    state.searchQuery = "";
    state.activeTypeFilter = "all";
    searchInput.value = "";
    clearSearchBtn.classList.add("hidden");
    renderExplorer();
  });

  // Compartir vista actual
  document.getElementById("copyFolderLinkBtn").addEventListener("click", () => {
    copyToClipboard(window.location.href, "Enlace a esta carpeta copiado al portapapeles");
  });

  // Tema claro / oscuro
  document.getElementById("themeToggleBtn").addEventListener("click", () => {
    const current = document.documentElement.getAttribute("data-theme") || "light";
    const next = current === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem("contic_seguro_theme", next);
    updateThemeIcons(next);
  });

  // Sidebar móvil
  document.getElementById("sidebarToggleBtn").addEventListener("click", () => {
    document.getElementById("sidebar").classList.toggle("open");
    document.getElementById("sidebarBackdrop").classList.toggle("hidden");
  });

  document.getElementById("sidebarBackdrop").addEventListener("click", () => {
    document.getElementById("sidebar").classList.remove("open");
    document.getElementById("sidebarBackdrop").classList.add("hidden");
  });

  // Modal de detalles de archivo
  document.getElementById("closeFileModalBtn").addEventListener("click", () => {
    document.getElementById("fileModalOverlay").classList.add("hidden");
  });

  document.getElementById("fileModalOverlay").addEventListener("click", e => {
    if (e.target.id === "fileModalOverlay") {
      document.getElementById("fileModalOverlay").classList.add("hidden");
    }
  });

  document.getElementById("modalCopyUrlBtn").addEventListener("click", () => {
    if (state.activeModalFile) {
      copyToClipboard(state.activeModalFile.url);
    }
  });

  document.getElementById("modalGoToFolderBtn").addEventListener("click", () => {
    if (state.activeModalFile) {
      document.getElementById("fileModalOverlay").classList.add("hidden");
      state.searchQuery = "";
      document.getElementById("searchInput").value = "";
      document.getElementById("clearSearchBtn").classList.add("hidden");
      navigateToFolder(state.activeModalFile.folderPath);
    }
  });

  document.getElementById("modalFolderChips").addEventListener("click", e => {
    const chip = e.target.closest("[data-modal-folder]");
    if (chip) {
      document.getElementById("fileModalOverlay").classList.add("hidden");
      state.searchQuery = "";
      document.getElementById("searchInput").value = "";
      document.getElementById("clearSearchBtn").classList.add("hidden");
      navigateToFolder(chip.getAttribute("data-modal-folder"));
    }
  });

  // Modal de Fuente de Datos (Google Sheets / Excel)
  document.getElementById("openDataModalBtn").addEventListener("click", () => {
    const savedSheet = localStorage.getItem("contic_seguro_google_sheet_url") || DEFAULT_GOOGLE_SHEET_URL;
    if (savedSheet) {
      document.getElementById("googleSheetUrlInput").value = savedSheet;
    }
    document.getElementById("dataModalOverlay").classList.remove("hidden");
  });

  document.getElementById("closeDataModalBtn").addEventListener("click", () => {
    document.getElementById("dataModalOverlay").classList.add("hidden");
  });

  document.getElementById("dataModalOverlay").addEventListener("click", e => {
    if (e.target.id === "dataModalOverlay") {
      document.getElementById("dataModalOverlay").classList.add("hidden");
    }
  });

  document.getElementById("syncSheetBtn").addEventListener("click", () => {
    const url = document.getElementById("googleSheetUrlInput").value;
    syncFromGoogleSheets(url, false);
  });

  document.getElementById("excelFileInput").addEventListener("change", e => {
    if (e.target.files && e.target.files[0]) {
      handleExcelFileUpload(e.target.files[0]);
    }
  });

  const dropzone = document.getElementById("excelDropzone");
  dropzone.addEventListener("dragover", e => {
    e.preventDefault();
    dropzone.style.borderColor = "var(--accent)";
  });
  dropzone.addEventListener("dragleave", () => {
    dropzone.style.borderColor = "";
  });
  dropzone.addEventListener("drop", e => {
    e.preventDefault();
    dropzone.style.borderColor = "";
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleExcelFileUpload(e.dataTransfer.files[0]);
    }
  });

  // Exportar data.js actualizado
  document.getElementById("exportDataJsBtn").addEventListener("click", () => {
    const content = `window.INITIAL_FILES_DATA = ${JSON.stringify(state.rawRecords, null, 2)};\n`;
    const blob = new Blob([content], { type: "application/javascript;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "data.js";
    a.click();
    URL.revokeObjectURL(url);
    showToast("Archivo data.js descargado");
  });

  // Restaurar datos originales
  document.getElementById("resetDefaultDataBtn").addEventListener("click", () => {
    localStorage.removeItem("contic_seguro_google_sheet_url");
    document.getElementById("googleSheetUrlInput").value = "";
    document.getElementById("sheetSyncStatus").classList.add("hidden");
    if (window.INITIAL_FILES_DATA) {
      loadDataset(window.INITIAL_FILES_DATA, "DocsShares Unipamplona");
      showToast("Restaurados los 632 archivos originales de URL.xlsx");
    }
  });
}

// Inicialización al cargar el DOM
document.addEventListener("DOMContentLoaded", () => {
  initTheme();
  initEventListeners();

  if (window.INITIAL_FILES_DATA && Array.isArray(window.INITIAL_FILES_DATA)) {
    loadDataset(window.INITIAL_FILES_DATA, "DocsShares Unipamplona");
  }

  // Si el usuario dejó conectado un Google Sheet previamente, sincronizar en segundo plano
  const savedSheetUrl = DEFAULT_GOOGLE_SHEET_URL || localStorage.getItem("contic_seguro_google_sheet_url");
  if (savedSheetUrl) {
    setTimeout(() => syncFromGoogleSheets(savedSheetUrl, true), 600);
  }
});

