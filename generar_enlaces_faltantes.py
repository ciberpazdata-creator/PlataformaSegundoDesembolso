import os
import re
import csv
import json
import getpass
import zipfile
import urllib.parse
import urllib.request
import base64
import xml.etree.ElementTree as ET

NEXTCLOUD_BASE_URL = "https://docsshares.unipamplona.edu.co"
NEXTCLOUD_LOCAL_ROOT = r"C:\Users\USER\Nextcloud"
TARGET_FOLDER = r"C:\Users\USER\Nextcloud\ENTREGABLES SEGUNDO DESEMBOLSO"

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(WORKSPACE_DIR, "URL.xlsx")
CSV_DESKTOP_PATH = r"D:\USUARIO\Desktop\Enlaces_DocsShares.csv"
DATA_JSON_PATH = os.path.join(WORKSPACE_DIR, "data.json")
DATA_JS_PATH = os.path.join(WORKSPACE_DIR, "data.js")


def clean_display_text(t):
    if not t:
        return ""
    t = str(t).strip().replace("\ufffd", "ó")
    t = re.sub(r"ciber\s*paz", "ConTIC seguro", t, flags=re.IGNORECASE)
    return t


def load_existing_links_from_excel(xlsx_path):
    """Carga el mapa {ruta_original_nextcloud: url} de los 632 archivos que ya están en URL.xlsx."""
    existing = {}
    if not os.path.exists(xlsx_path):
        return existing

    with zipfile.ZipFile(xlsx_path, "r") as z:
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                texts = [t.text for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t") if t.text]
                shared_strings.append("".join(texts))

        s1_root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        for row in s1_root.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
            r_num = int(row.attrib.get("r", 0))
            if r_num <= 2:
                continue
            cells = {}
            for c in row.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                c_ref = c.attrib.get("r", "")
                col = re.sub(r"\d+", "", c_ref)
                c_type = c.attrib.get("t")
                v = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                val = v.text if v is not None else ""
                if c_type == "s" and val != "":
                    val = shared_strings[int(val)]
                cells[col] = val.strip()
            ruta = cells.get("B", "")
            url = cells.get("C", "")
            if ruta and url.startswith("http"):
                existing[ruta] = url
    return existing


def scan_local_nextcloud_files():
    """Escanea todos los archivos dentro de C:\\Users\\USER\\Nextcloud\\ENTREGABLES SEGUNDO DESEMBOLSO."""
    all_files = []
    for root_dir, dirs, files in os.walk(TARGET_FOLDER):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        for file in files:
            if file.startswith(".") or file.lower() in ("desktop.ini", "thumbs.db"):
                continue
            full_path = os.path.join(root_dir, file)
            rel_to_nc = os.path.relpath(full_path, NEXTCLOUD_LOCAL_ROOT).replace("\\", "/")
            nc_path = "/" + rel_to_nc
            all_files.append((file, nc_path))
    return all_files


def make_ocs_request(method, endpoint, username, password, data=None):
    url = f"{NEXTCLOUD_BASE_URL}{endpoint}"
    credentials = f"{username}:{password}"
    b64_auth = base64.b64encode(credentials.encode("utf-8")).decode("ascii")
    headers = {
        "Authorization": f"Basic {b64_auth}",
        "OCS-APIRequest": "true",
        "Accept": "application/json",
    }
    body = None
    if data is not None:
        body = urllib.parse.urlencode(data).encode("utf-8")
        headers["Content-Type"] = "application/x-www-form-urlencoded"

    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def fetch_existing_server_shares(username, password):
    """Consulta en DocsShares los enlaces compartidos que ya existen en el servidor."""
    server_map = {}
    try:
        res = make_ocs_request(
            "GET",
            "/ocs/v2.php/apps/files_sharing/api/v1/shares?format=json&subfiles=true",
            username,
            password,
        )
        shares = res.get("ocs", {}).get("data", [])
        if isinstance(shares, list):
            for s in shares:
                p = s.get("path")
                u = s.get("url")
                if p and u:
                    server_map[p] = u
    except Exception:
        pass
    return server_map


def create_share_link(nc_path, username, password):
    """Crea un enlace público (shareType=3) en Nextcloud para la ruta indicada."""
    res = make_ocs_request(
        "POST",
        "/ocs/v2.php/apps/files_sharing/api/v1/shares?format=json",
        username,
        password,
        data={"path": nc_path, "shareType": 3, "permissions": 1},
    )
    data = res.get("ocs", {}).get("data", {})
    return data.get("url")


def build_web_records(path_to_url_map):
    records = []
    for idx, (raw_path, url) in enumerate(sorted(path_to_url_map.items()), 1):
        raw_name = raw_path.split("/")[-1]
        name = clean_display_text(raw_name)
        ruta = clean_display_text(raw_path)

        parts = [p.strip() for p in ruta.split("/") if p.strip()]
        if parts and parts[-1].lower() == name.lower():
            folder_parts = parts[:-1]
        elif len(parts) > 1:
            folder_parts = parts[:-1]
        else:
            folder_parts = parts

        ext = name.rsplit(".", 1)[-1].strip().lower() if "." in name else "otro"

        records.append({
            "id": idx,
            "name": name,
            "path": ruta,
            "folderPath": "/" + "/".join(folder_parts),
            "folders": folder_parts,
            "url": url.strip(),
            "ext": ext,
        })
    return records


def main():
    print("=" * 70)
    print(" ESCANEO Y ACTUALIZACIÓN DE ENLACES FALTANTES EN NEXTCLOUD")
    print("=" * 70)

    existing_map = load_existing_links_from_excel(EXCEL_PATH)
    local_files = scan_local_nextcloud_files()

    missing = [(name, path) for (name, path) in local_files if path not in existing_map]

    print(f"• Archivos totales en tu carpeta Nextcloud : {len(local_files)}")
    print(f"• Archivos que ya tienen URL en URL.xlsx   : {len(existing_map)}")
    print(f"• Archivos NUEVOS que faltan por enlace    : {len(missing)}")
    print("-" * 70)

    if not missing:
        print("¡Todos los archivos de la carpeta ya tienen enlace! Actualizando data.js...")
    else:
        print("\nPara generar los enlaces públicos de los 123 archivos nuevos en DocsShares,")
        print("ingresa tu usuario y contraseña de https://docsshares.unipamplona.edu.co:")
        username = input("Usuario de DocsShares (o correo): ").strip()
        password = getpass.getpass("Contraseña de DocsShares (no se verá al escribir): ").strip()

        print("\nConectando con DocsShares Unipamplona...")
        created_count = 0
        errors = []

        for idx, (fname, nc_path) in enumerate(missing, 1):
            try:
                url = create_share_link(nc_path, username, password)
                if url:
                    existing_map[nc_path] = url
                    created_count += 1
                    print(f"[{idx}/{len(missing)}] OK -> {fname}")
                else:
                    errors.append((nc_path, "Respuesta sin URL"))
                    print(f"[{idx}/{len(missing)}] SIN URL -> {fname}")
            except Exception as e:
                errors.append((nc_path, str(e)))
                print(f"[{idx}/{len(missing)}] ERROR -> {fname}: {e}")

        print("-" * 70)
        print(f"Enlaces nuevos generados: {created_count} de {len(missing)}")
        if errors:
            print(f"Archivos con error: {len(errors)}")

    # Guardar CSV actualizado en el Escritorio
    try:
        with open(CSV_DESKTOP_PATH, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Nombre del archivo", "Ruta del Archivo", "URL"])
            for raw_path, url in sorted(existing_map.items()):
                fname = raw_path.split("/")[-1]
                writer.writerow([fname, raw_path, url])
        print(f"• CSV actualizado en: {CSV_DESKTOP_PATH}")
    except Exception as e:
        print(f"• Aviso al guardar CSV: {e}")

    # Actualizar data.json y data.js de la plataforma web
    records = build_web_records(existing_map)

    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    with open(DATA_JS_PATH, "w", encoding="utf-8") as f:
        f.write("window.INITIAL_FILES_DATA = ")
        json.dump(records, f, ensure_ascii=False, indent=2)
        f.write(";\n")

    print(f"• Plataforma web actualizada (data.js y data.json) con {len(records)} archivos totales.")
    print("=" * 70)


if __name__ == "__main__":
    main()

