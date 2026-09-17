# -*- coding: utf-8 -*-
import time
import unicodedata
import re
from typing import Dict, Any, Tuple, List, Optional
from urllib.parse import quote
import requests


def ActualizarLicenciasExcelEnSharepoint(
    CONFIG: Dict[str, Any],
    licencias: List[Dict[str, Any]],
    lectura_config: Dict[str, Any],
) -> Tuple[bool, str]:
    """
    Actualiza (in-place) o crea (desde plantilla) un Excel en SharePoint y escribe los datos
    agrupados por 'HojaExcel' (sin incluir esa columna como encabezado).

    - Si el archivo YA existe y behavior_when_exists='update' (por defecto) -> abre sesión y actualiza.
    - Si NO existe -> copia desde 'template_relative_path' (de lectura_config) y luego escribe.
    - 'entorno' en lectura_config: 'desarrollo' -> escribe en *_DEV.xlsx (o file_relative_path_dev),
                                   'produccion' -> escribe en el mismo file_relative_path.

    CONFIG:
        { "TENANT_ID": "...", "CLIENT_ID": "...", "CLIENT_SECRET": "..." }

    lectura_config (DATOS_EXCEL_SHAREPOINT):
        {
            "site_hostname": "...",
            "site_path": "...",
            "entorno": "desarrollo" | "produccion",
            "file_relative_path": "Licencias y Garantías/InvLicPython3.xlsx",
            # "file_relative_path_dev": "Licencias y Garantías/InvLicPython3_DEV.xlsx"  (opcional)
            "template_relative_path": "Licencias y Garantías/PlantillaInventLicPythonNOBORRAR.xlsx",

            # lectura
            "hojas_permitidas": [...],
            "campos_salida_orden": [...],
            "campos_fecha_ddmmyyyy": [...],
            "aplanar": True,
            "nombre_campo_hoja": "HojaExcel",

            # escritura (opcionales; defaults razonables)
            "behavior_when_exists": "update",
            "batch_size": 1000,
            "remove_other_sheets": True,
            # "preserve_sheets": ["Resumen"]
        }

    Retorna: (success: bool, message: str)
    """

    # ------------------------------ Validaciones ------------------------------
    TENANT_ID = str(CONFIG.get("TENANT_ID", "")).strip()
    CLIENT_ID = str(CONFIG.get("CLIENT_ID", "")).strip()
    CLIENT_SECRET = str(CONFIG.get("CLIENT_SECRET", "")).strip()
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        return False, "CONFIG incompleto: faltan TENANT_ID, CLIENT_ID o CLIENT_SECRET."

    if not isinstance(lectura_config, dict):
        return False, "El tercer argumento debe ser DATOS_EXCEL_SHAREPOINT (dict)."

    site_hostname = lectura_config.get("site_hostname")
    site_path = lectura_config.get("site_path")
    entorno = str(lectura_config.get("entorno", "desarrollo")).strip().lower()

    file_rel_prod = lectura_config.get("file_relative_path")
    file_rel_dev = lectura_config.get("file_relative_path_dev")  # opcional
    template_relative_path = lectura_config.get("template_relative_path")

    behavior_when_exists = str(lectura_config.get("behavior_when_exists", "update")).strip().lower()
    batch_size = int(lectura_config.get("batch_size", 1000))
    remove_other_sheets = bool(lectura_config.get("remove_other_sheets", False))
    preserve_sheets: List[str] = lectura_config.get("preserve_sheets", []) or []

    if not site_hostname or not site_path or not file_rel_prod:
        return False, "Faltan site_hostname, site_path o file_relative_path en lectura_config."

    if behavior_when_exists not in ("update", "overwrite_from_template"):
        return False, "behavior_when_exists debe ser 'update' o 'overwrite_from_template'."

    if not isinstance(licencias, list) or len(licencias) == 0:
        return False, "La lista de licencias está vacía o no es válida."

    # ------------------------------ Resolver destino por entorno ------------------------------
    def _derive_dev_path_from_prod(path: str) -> str:
        # Inserta "_DEV" antes de la extensión
        if not isinstance(path, str) or not path.strip():
            return path
        p = path.strip()
        if "." in p.rsplit("/", 1)[-1]:
            base, ext = p.rsplit(".", 1)
            return f"{base}_DEV.{ext}"
        else:
            return f"{p}_DEV"

    if entorno == "produccion":
        dest_file_relative_path = file_rel_prod
    else:
        dest_file_relative_path = file_rel_dev or _derive_dev_path_from_prod(file_rel_prod)

    # ------------------------------ Constantes ------------------------------
    HTTP_TIMEOUT = 30
    MAX_RETRIES = 5

    # ------------------------------ Utils / Graph helpers ------------------------------
    def _graph_request(method: str, url: str, token: str, *,
                       params=None, json_body=None, extra_headers: Optional[Dict[str, str]] = None):
        headers = {"Authorization": f"Bearer {token}"}
        if json_body is not None:
            headers["Content-Type"] = "application/json"
        if extra_headers:
            headers.update(extra_headers)
        backoff = 1.7
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                r = requests.request(method.upper(), url, headers=headers,
                                     params=params, json=json_body, timeout=HTTP_TIMEOUT)
                if r.status_code in (429, 500, 502, 503, 504):
                    ra = r.headers.get("Retry-After")
                    wait = float(ra) if ra else min(32, backoff ** attempt)
                    time.sleep(wait)
                    continue
                try:
                    return r.status_code, r.json()
                except Exception:
                    return r.status_code, r.text
            except Exception:
                if attempt == MAX_RETRIES:
                    raise
                time.sleep(min(30, backoff ** attempt))
        return 599, "Error desconocido"

    def _get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {"grant_type": "client_credentials","client_id": client_id,
                "client_secret": client_secret,"scope": "https://graph.microsoft.com/.default"}
        r = requests.post(url, data=data, timeout=HTTP_TIMEOUT)
        if r.status_code != 200: raise RuntimeError(f"Error token ({r.status_code}): {r.text[:400]}")
        tok = r.json().get("access_token")
        if not tok: raise RuntimeError("No se recibió access_token.")
        return tok

    def _get_site_id(token: str) -> str:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
        status, data = _graph_request("GET", url, token)
        if status != 200: raise RuntimeError(f"Error obteniendo siteId ({status}): {str(data)[:400]}")
        site_id = data.get("id")
        if not site_id: raise RuntimeError("Respuesta sin 'id' para el sitio.")
        return site_id

    def _get_default_drive(token: str, site_id: str) -> str:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        status, data = _graph_request("GET", url, token)
        if status != 200: raise RuntimeError(f"Error obteniendo drive por defecto ({status}): {str(data)[:400]}")
        drive_id = data.get("id")
        if not drive_id: raise RuntimeError("Respuesta sin 'id' para el drive.")
        return drive_id

    def _get_item_by_path(token: str, site_id: str, relative_path: str) -> Dict[str, Any]:
        encoded = quote(str(relative_path).strip("/"), safe="/-_.()")
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded}"
        status, data = _graph_request("GET", url, token)
        if status != 200:
            raise RuntimeError(f"Error obteniendo item por path '{relative_path}' ({status}): {str(data)[:400]}")
        return data

    # Workbook session helpers
    def _create_session(token: str, drive_id: str, item_id: str) -> str:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/createSession"
        status, data = _graph_request("POST", url, token, json_body={"persistChanges": True})
        if status not in (200, 201):
            raise RuntimeError(f"Error creando sesión de workbook ({status}): {str(data)[:300]}")
        sid = data.get("id")
        if not sid:
            raise RuntimeError("No se recibió 'id' de sesión de workbook.")
        return sid

    def _list_worksheets(token: str, drive_id: str, item_id: str, sid: str) -> List[Dict[str, Any]]:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets"
        status, data = _graph_request("GET", url, token, extra_headers={"workbook-session-id": sid})
        if status != 200:
            return []
        return data.get("value", []) if isinstance(data, dict) else []

    def _ensure_worksheet(token: str, drive_id: str, item_id: str, sheet_name: str, sid: str) -> None:
        hojas = _list_worksheets(token, drive_id, item_id, sid)
        if any(h.get("name") == sheet_name for h in hojas):
            return
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets/add"
        status, data = _graph_request("POST", url, token, json_body={"name": sheet_name}, extra_headers={"workbook-session-id": sid})
        if status not in (200, 201):
            raise RuntimeError(f"No se pudo crear la hoja '{sheet_name}' ({status}): {str(data)[:300]}")

    def _delete_worksheet_by_name(token: str, drive_id: str, item_id: str, sheet_name: str, sid: str) -> None:
        hojas = _list_worksheets(token, drive_id, item_id, sid)
        if len(hojas) <= 1:
            return
        for h in hojas:
            if h.get("name") == sheet_name:
                sheet_id = h.get("id")
                if sheet_id:
                    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_id}"
                    _graph_request("DELETE", url, token, extra_headers={"workbook-session-id": sid})
                break

    def _clear_sheet(token: str, drive_id: str, item_id: str, sheet_name: str, sid: str) -> None:
        sheet_enc = quote(sheet_name)
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_enc}/usedRange/clear"
        _graph_request("POST", url, token, json_body={"applyTo": "all"}, extra_headers={"workbook-session-id": sid})

    # Rango dimensionado
    def _col_letters(n: int) -> str:
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    def _col_index_from_letters(s: str) -> int:
        s = s.upper()
        n = 0
        for ch in s:
            if 'A' <= ch <= 'Z':
                n = n * 26 + (ord(ch) - 64)
        return n

    def _parse_cell(cell: str) -> Tuple[int, int]:
        m = re.match(r"^([A-Z]+)(\d+)$", cell.strip().upper())
        if not m:
            raise ValueError(f"Celda inválida: {cell}")
        return _col_index_from_letters(m.group(1)), int(m.group(2))

    def _address_from_block(start_cell: str, rows: int, cols: int) -> str:
        sc, sr = _parse_cell(start_cell)
        ec, er = sc + max(cols, 1) - 1, sr + max(rows, 1) - 1
        sl, el = _col_letters(sc), _col_letters(ec)
        return f"{sl}{sr}" if (rows == 1 and cols == 1) else f"{sl}{sr}:{el}{er}"

    def _patch_range_values(token: str, drive_id: str, item_id: str,
                            sheet_name: str, start_cell: str, values: List[List[Any]],
                            sid: str) -> None:
        values = values or []
        rows = len(values)
        cols = max((len(r) for r in values), default=0)
        if rows == 0 or cols == 0:
            return
        rng = _address_from_block(start_cell, rows, cols)
        sheet_enc = quote(sheet_name)
        addr_enc = quote(rng, safe="'!:$ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789")
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_enc}/range(address='{addr_enc}')"
        status, data = _graph_request("PATCH", url, token, json_body={"values": values}, extra_headers={"workbook-session-id": sid})
        if status not in (200, 201):
            raise RuntimeError(f"Error escribiendo valores en '{sheet_name}'!{rng} ({status}): {str(data)[:300]}")

    # ------------------------------ Encabezados (derivados) y agrupación ------------------------------
    # Encabezados: usa 'campos_salida_orden' si viene, si no, usa keys del primer dict (sin HojaExcel)
    campos_salida_orden = lectura_config.get("campos_salida_orden") or []
    if campos_salida_orden:
        headers = [h for h in campos_salida_orden if h and h != "HojaExcel"]
    else:
        first = licencias[0]
        headers = [k for k in first.keys() if k != "HojaExcel"]

    # Incluir cualquier clave presente en datos que no esté (excepto HojaExcel)
    present_keys = set()
    for d in licencias:
        if isinstance(d, dict):
            for k in d.keys():
                if k != "HojaExcel":
                    present_keys.add(k)
    for k in sorted(present_keys):
        if k not in headers:
            headers.append(k)

    # Agrupar por HojaExcel y filtrar por hojas_permitidas si corresponde
    def _sanitize_name(name: str) -> str:
        invalid = set(':\\/?*[]')
        cleaned = "".join(ch if ch not in invalid else "_" for ch in (name or "Datos"))
        cleaned = cleaned.strip()
        if not cleaned:
            cleaned = "Datos"
        if len(cleaned) > 31:
            cleaned = cleaned[:31]
        return cleaned.strip("'")

    grupos: Dict[str, List[Dict[str, Any]]] = {}
    for d in licencias:
        hoja_raw = str(d.get("HojaExcel", "Datos")).strip() or "Datos"
        hoja = _sanitize_name(hoja_raw)
        grupos.setdefault(hoja, []).append({k: v for k, v in d.items() if k != "HojaExcel"})

    hojas_permitidas = lectura_config.get("hojas_permitidas")
    if hojas_permitidas:
        hp = set([str(h).strip() for h in hojas_permitidas if isinstance(h, str) and h.strip()])
        grupos = {h: filas for h, filas in grupos.items() if h in hp}

    # Evitar colisiones de nombres (saneado/longitud)
    final_grupos: Dict[str, List[Dict[str, Any]]] = {}
    seen = set()
    for hoja, filas in grupos.items():
        base = hoja
        suf = 1
        while hoja in seen:
            core = base[:28] if len(base) > 28 else base
            hoja = f"{core}_{suf}"
            suf += 1
        final_grupos[hoja] = filas
        seen.add(hoja)

    # ------------------------------ Flujo principal ------------------------------
    try:
        token = _get_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        site_id = _get_site_id(token)
        drive_id = _get_default_drive(token, site_id)

        # ¿Existe ya el archivo destino?
        dest_exists = False
        dest_id = None
        try:
            dest_item = _get_item_by_path(token, site_id, dest_file_relative_path)
            dest_id = dest_item.get("id")
            dest_exists = bool(dest_id)
        except Exception:
            dest_exists = False

        # Crear desde plantilla si NO existe; si existe, actualizar según behavior_when_exists
        if not dest_exists:
            if not template_relative_path:
                return False, (
                    f"El archivo destino '{dest_file_relative_path}' no existe y no se indicó plantilla "
                    f"('template_relative_path') en DATOS_EXCEL_SHAREPOINT. Defínela."
                )
            # Copiar plantilla -> destino
            dest_rel_path = str(dest_file_relative_path).strip("/")
            if "/" in dest_rel_path:
                parent_path = dest_rel_path.rsplit("/", 1)[0]
                dest_name = dest_rel_path.rsplit("/", 1)[1]
            else:
                parent_path = ""
                dest_name = dest_rel_path
            parent_ref = {"path": f"/drive/root:/{parent_path}" if parent_path else "/drive/root:"}
            body = {"parentReference": parent_ref, "name": dest_name}

            tpl_item = _get_item_by_path(token, site_id, template_relative_path)
            tpl_id = tpl_item.get("id")
            if not tpl_id:
                return False, "No se pudo resolver el itemId de la plantilla."

            url_copy = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{tpl_id}/copy"
            status, data = _graph_request("POST", url_copy, token, json_body=body)
            if status not in (202, 200):
                return False, f"Error copiando plantilla -> destino ({status}): {str(data)[:400]}"

            # Esperar disponibilidad y abrir sesión
            t0 = time.time()
            ready = False
            session_id = None
            while time.time() - t0 < 90:
                try:
                    dest_item = _get_item_by_path(token, site_id, dest_file_relative_path)
                    dest_id = dest_item.get("id")
                    if not dest_id:
                        time.sleep(2); continue
                    session_id = _create_session(token, drive_id, dest_id)
                    ready = True
                    break
                except Exception:
                    time.sleep(2)
            if not ready:
                return False, "El archivo creado no quedó listo para abrirse con Excel (sesión no disponible)."

        else:
            if behavior_when_exists == "overwrite_from_template":
                if not template_relative_path:
                    return False, "Para 'overwrite_from_template' debes indicar 'template_relative_path'."
                # Borrar y recrear
                del_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{dest_id}"
                _graph_request("DELETE", del_url, token)

                dest_rel_path = str(dest_file_relative_path).strip("/")
                if "/" in dest_rel_path:
                    parent_path = dest_rel_path.rsplit("/", 1)[0]
                    dest_name = dest_rel_path.rsplit("/", 1)[1]
                else:
                    parent_path = ""
                    dest_name = dest_rel_path
                parent_ref = {"path": f"/drive/root:/{parent_path}" if parent_path else "/drive/root:"}
                body = {"parentReference": parent_ref, "name": dest_name}

                tpl_item = _get_item_by_path(token, site_id, template_relative_path)
                tpl_id = tpl_item.get("id")
                if not tpl_id:
                    return False, "No se pudo resolver el itemId de la plantilla."
                url_copy = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{tpl_id}/copy"
                status, data = _graph_request("POST", url_copy, token, json_body=body)
                if status not in (202, 200):
                    return False, f"Error copiando plantilla -> destino ({status}): {str(data)[:400]}"

                # Esperar y abrir sesión
                t0 = time.time()
                ready = False
                session_id = None
                while time.time() - t0 < 90:
                    try:
                        dest_item = _get_item_by_path(token, site_id, dest_file_relative_path)
                        dest_id = dest_item.get("id")
                        if not dest_id:
                            time.sleep(2); continue
                        session_id = _create_session(token, drive_id, dest_id)
                        ready = True
                        break
                    except Exception:
                        time.sleep(2)
                if not ready:
                    return False, "El archivo destino no quedó listo para abrirse con Excel (sesión no disponible)."
            else:
                # update in-place
                session_id = _create_session(token, drive_id, dest_id)

        # 4) (Opcional) eliminar hojas no usadas (respetando preserve_sheets y dejando mínimo 1)
        if remove_other_sheets:
            hojas_existentes = _list_worksheets(token, drive_id, dest_id, session_id)
            existentes = [h.get("name") for h in hojas_existentes if isinstance(h.get("name"), str)]
            objetivo = list(final_grupos.keys())
            keep_extra = set([s for s in preserve_sheets if isinstance(s, str)])
            for h in existentes:
                if h in keep_extra:
                    continue
                if h not in objetivo and len(existentes) > 1:
                    _delete_worksheet_by_name(token, drive_id, dest_id, h, session_id)

        # 5) Escribir por hoja (encabezados + datos)
        if batch_size <= 0:
            batch_size = 1000

        def _fila_to_row(obj: Dict[str, Any]) -> List[Any]:
            return [obj.get(h, None) for h in headers]

        total_filas = 0
        hojas_escritas = 0
        for hoja, filas in final_grupos.items():
            _ensure_worksheet(token, drive_id, dest_id, hoja, session_id)
            _clear_sheet(token, drive_id, dest_id, hoja, session_id)

            # Encabezados -> A1:{colFinal}1
            _patch_range_values(token, drive_id, dest_id, hoja, "A1", [headers], session_id)

            # Datos -> A2:{colFinal}{fin}, por lotes
            start_row = 2
            for i in range(0, len(filas), batch_size):
                chunk = filas[i:i + batch_size]
                values = [_fila_to_row(o) for o in chunk]
                k = len(headers)
                values = [r + [None] * (k - len(r)) if len(r) < k else r[:k] for r in values]
                _patch_range_values(token, drive_id, dest_id, hoja, f"A{start_row}", values, session_id)
                start_row += len(values)
                total_filas += len(values)

            hojas_escritas += 1

        msg_ok = (
            f"{'Actualizado in-place' if dest_exists and behavior_when_exists=='update' else 'Creado/actualizado'} "
            f"'{dest_file_relative_path}' con {hojas_escritas} hoja(s) y {total_filas} fila(s)."
        )
        return True, msg_ok

    except Exception as ex:
        return False, f"Error creando/actualizando Excel: {ex}"
