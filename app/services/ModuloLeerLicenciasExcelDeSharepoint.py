# -*- coding: utf-8 -*-
import time
import unicodedata
import re
from datetime import datetime, date, timedelta, timezone
from typing import Dict, Any, Tuple, List, Optional
from urllib.parse import quote
import requests


def LeerLicenciasExcelDeSharepoint(
    CONFIG: Dict[str, Any],
    lectura_config: Dict[str, Any],
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """
    Lee un Excel (XLSX) en SharePoint mediante Microsoft Graph (app-only) y retorna registros **aplanados**.

    Parámetros:
        CONFIG: dict con credenciales app-only:
            {
                "TENANT_ID": "<tenant_id>",
                "CLIENT_ID": "<app_id>",
                "CLIENT_SECRET": "<secret>"
            }

        lectura_config: dict con parámetros de lectura (se extraen internamente):
            {
                "site_hostname": "interlan.sharepoint.com",
                "site_path": "/sites/csa",
                "file_relative_path": "Licencias y Garantías/InvLicPython3.xlsx",

                # Opcionales (si no vienen, se aplican defaults)
                "hojas_permitidas": [ ... ]                   # None o lista -> filtra hojas
                "campos_salida_orden": [ ... ]                # None -> devolver todas las columnas
                "campos_fecha_ddmmyyyy": [ "Fecha X", ... ]
                "aplanar": True,
                "nombre_campo_hoja": "HojaExcel",

                # Puede traer claves extra (p. ej. 'entorno'), se ignoran.
            }

    Retorna:
        (success: bool, message: str, data: list[dict])
    """

    # ------------------------- Validaciones básicas -------------------------
    TENANT_ID = str(CONFIG.get("TENANT_ID", "")).strip()
    CLIENT_ID = str(CONFIG.get("CLIENT_ID", "")).strip()
    CLIENT_SECRET = str(CONFIG.get("CLIENT_SECRET", "")).strip()
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        return False, "CONFIG incompleto: faltan TENANT_ID, CLIENT_ID o CLIENT_SECRET.", []

    if not isinstance(lectura_config, dict):
        return False, "lectura_config debe ser un diccionario.", []

    # Extraer parámetros desde lectura_config (claves extra se ignoran)
    site_hostname: str = lectura_config.get("site_hostname", "")
    site_path: str = lectura_config.get("site_path", "")
    file_relative_path: str = lectura_config.get("file_relative_path", "")
    hojas_permitidas: Optional[List[str]] = lectura_config.get("hojas_permitidas", None)
    campos_salida_orden: Optional[List[str]] = lectura_config.get("campos_salida_orden", None)
    campos_fecha_ddmmyyyy: Optional[List[str]] = lectura_config.get("campos_fecha_ddmmyyyy", None)
    aplanar: bool = bool(lectura_config.get("aplanar", True))
    nombre_campo_hoja: Optional[str] = lectura_config.get("nombre_campo_hoja", "HojaExcel")

    if not site_hostname or not site_path or not file_relative_path:
        return False, "Parámetros obligatorios faltantes: site_hostname, site_path o file_relative_path.", []

    # ------------------------- Constantes locales -------------------------
    HTTP_TIMEOUT = 30
    MAX_RETRIES = 5

    # ------------------------- Helpers internos -------------------------
    def _key_norm(s: Any) -> str:
        if s is None:
            return ""
        s = str(s).strip()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = s.lower()
        s = "".join(ch for ch in s if ch.isalnum())
        return s

    def _row_is_meaningful_generic(d: Dict[str, Any]) -> bool:
        if not d:
            return False
        for _, v in d.items():
            if v is not None and str(v).strip() != "":
                return True
        return False

    # ---- Fechas
    def _parse_odata_date(s: str) -> Tuple[Optional[date], Optional[str]]:
        try:
            m = re.match(r"^/Date\((-?\d+)\)/$", s)
            if not m:
                return None, "OData pattern no coincide"
            ms = int(m.group(1))
            dt = datetime.fromtimestamp(ms / 1000.0, tz=timezone.utc)
            return dt.date(), None
        except Exception as ex:
            return None, f"OData no parseable: {ex}"

    def _parse_iso_or_local(s: str) -> Tuple[Optional[date], Optional[str]]:
        try:
            s2 = s[:-1] + "+00:00" if s.endswith("Z") else s
            dt = datetime.fromisoformat(s2)
            return dt.date(), None
        except Exception:
            pass
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S",
                    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.date(), None
            except Exception:
                continue
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
                    "%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.date(), None
            except Exception:
                continue
        return None, f"Formato no reconocido: {s}"

    def _excel_serial_to_date(n: float) -> Optional[date]:
        try:
            n = float(n)
            if n <= 0:
                return None
            base = date(1899, 12, 30)
            d = base + timedelta(days=n)
            if d.year < 1900 or d.year > 2100:
                return None
            return d
        except Exception:
            return None

    def _coerce_fecha_ddmmyyyy(value: Any) -> str:
        d: Optional[date] = None
        if value is None:
            return ""
        if isinstance(value, datetime):
            d = value.date()
        elif isinstance(value, date):
            d = value
        elif isinstance(value, (int, float)):
            d = _excel_serial_to_date(value)
        elif isinstance(value, dict):
            for k in ("dateTime", "DateTime", "value"):
                if k in value and isinstance(value[k], str) and value[k].strip():
                    s = value[k].strip()
                    d1, _ = _parse_iso_or_local(s)
                    if d1:
                        d = d1
                        break
                    d2, _ = _parse_odata_date(s)
                    if d2:
                        d = d2
                        break
        elif isinstance(value, str):
            s = value.strip()
            if s:
                if s.startswith("/Date("):
                    d3, _ = _parse_odata_date(s)
                    d = d3
                else:
                    d4, _ = _parse_iso_or_local(s)
                    d = d4
        return d.strftime("%d/%m/%Y") if d else ""

    # ---- Graph
    def _get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        r = requests.post(url, data=data, timeout=HTTP_TIMEOUT)
        if r.status_code != 200:
            raise RuntimeError(f"Error token ({r.status_code}): {r.text[:500]}")
        tok = r.json().get("access_token")
        if not tok:
            raise RuntimeError("No se recibió access_token.")
        return tok

    def _graph_request(method: str, url: str, token: str, *, params=None, json_body=None):
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        backoff = 1.7
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                r = requests.request(
                    method=method.upper(),
                    url=url,
                    headers=headers,
                    params=params,
                    json=json_body,
                    timeout=HTTP_TIMEOUT,
                )
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

    def _get_site_id(token: str) -> str:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
        status, data = _graph_request("GET", url, token)
        if status != 200:
            raise RuntimeError(f"Error obteniendo siteId ({status}): {str(data)[:500]}")
        site_id = data.get("id")
        if not site_id:
            raise RuntimeError("Respuesta sin 'id' para el sitio.")
        return site_id

    def _get_default_drive(token: str, site_id: str) -> str:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        status, data = _graph_request("GET", url, token)
        if status != 200:
            raise RuntimeError(f"Error obteniendo drive por defecto ({status}): {str(data)[:500]}")
        drive_id = data.get("id")
        if not drive_id:
            raise RuntimeError("Respuesta sin 'id' para el drive.")
        return drive_id

    def _get_item_by_path(token: str, site_id: str, relative_path: str) -> Dict[str, Any]:
        encoded_path = quote(str(relative_path).strip("/"), safe="/-_.()")
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded_path}"
        status, data = _graph_request("GET", url, token)
        if status != 200:
            raise RuntimeError(f"Error obteniendo item por path '{relative_path}' ({status}): {str(data)[:500]}")
        return data

    # ---- Excel
    def _list_worksheets(token: str, drive_id: str, item_id: str) -> List[str]:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets"
        status, data = _graph_request("GET", url, token)
        if status != 200:
            raise RuntimeError(f"Error listando hojas ({status}): {str(data)[:500]}")
        value = data.get("value") if isinstance(data, dict) else None
        if not isinstance(value, list):
            return []
        hojas = []
        for it in value:
            nm = it.get("name")
            if isinstance(nm, str):
                hojas.append(nm)
        return hojas

    def _read_used_range(token: str, drive_id: str, item_id: str, sheet_name: str) -> List[List[Any]]:
        sheet_enc = quote(sheet_name)
        url = (
            f"https://graph.microsoft.com/v1.0/"
            f"drives/{drive_id}/items/{item_id}/workbook/"
            f"worksheets/{sheet_enc}/usedRange(valuesOnly=true)"
        )
        status, data = _graph_request("GET", url, token)
        if status != 200:
            return []
        return data.get("values", []) if isinstance(data, dict) else []

    def _headers_to_index_map(headers_row: List[Any]) -> Dict[str, int]:
        idx_map: Dict[str, int] = {}
        for i, h in enumerate(headers_row or []):
            tok = _key_norm(h)
            if tok and tok not in idx_map:
                idx_map[tok] = i
        return idx_map

    def _extract_rows_from_sheet(values: List[List[Any]]) -> List[Dict[str, Any]]:
        """
        Devuelve lista de dicts basada en:
          - campos_salida_orden (si viene) o todas las columnas si es None.
          - campos_fecha_ddmmyyyy -> 'DD/MM/YYYY'.
          - Filtra filas vacías (al menos un valor no vacío).
        """
        if not values or len(values) < 2:
            return []

        headers = values[0]
        rows = values[1:]
        idx_map = _headers_to_index_map(headers)

        # Sin campos_salida_orden: devuelve todas las columnas
        if not campos_salida_orden:
            headers_text = [str(h) if h is not None else "" for h in headers]
            salida: List[Dict[str, Any]] = []
            for raw in rows:
                fila: Dict[str, Any] = {}
                for i, htxt in enumerate(headers_text):
                    val = raw[i] if i < len(raw) else None
                    if campos_fecha_ddmmyyyy and htxt in (campos_fecha_ddmmyyyy or []):
                        fila[htxt] = _coerce_fecha_ddmmyyyy(val)
                    else:
                        fila[htxt] = val
                if _row_is_meaningful_generic(fila):
                    salida.append(fila)
            return salida

        # Con campos_salida_orden: respeta ese orden y normaliza fechas listadas
        fecha_set = set(campos_fecha_ddmmyyyy or [])

        def _idx(campo: str) -> int:
            return idx_map.get(_key_norm(campo), -1)

        idx_out = {campo: _idx(campo) for campo in campos_salida_orden}

        salida: List[Dict[str, Any]] = []
        for raw in rows:
            fila: Dict[str, Any] = {}
            for campo in campos_salida_orden:
                i = idx_out.get(campo, -1)
                val = raw[i] if (i >= 0 and i < len(raw)) else None
                fila[campo] = _coerce_fecha_ddmmyyyy(val) if campo in fecha_set else val

            # Cast rápido de 'Cantidad' si vino como str
            if "Cantidad" in fila and isinstance(fila["Cantidad"], str):
                try:
                    fila["Cantidad"] = int(fila["Cantidad"])
                except Exception:
                    pass

            if _row_is_meaningful_generic(fila):
                salida.append(fila)

        return salida

    # ------------------------- Flujo principal -------------------------
    try:
        # 1) Token & Site
        token = _get_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        site_id = _get_site_id(token)
        drive_id = _get_default_drive(token, site_id)

        # 2) Item por ruta relativa
        item = _get_item_by_path(token, site_id, file_relative_path)
        item_id = item.get("id")
        if not item_id:
            return False, "No se pudo resolver el itemId del archivo objetivo.", []

        # 3) Hojas presentes
        hojas_presentes = _list_worksheets(token, drive_id, item_id)
        if not hojas_presentes:
            return False, "El workbook no tiene hojas o no se pudieron listar.", []

        # 4) Filtro de hojas (exacto; si None -> todas)
        if hojas_permitidas is None:
            hojas_a_leer = hojas_presentes
        else:
            hp = [h for h in hojas_permitidas if isinstance(h, str) and h.strip() != ""]
            hojas_a_leer = [h for h in hojas_presentes if h in hp] if hp else hojas_presentes

        if not hojas_a_leer:
            return False, "No se encontraron hojas a leer según el filtro.", []

        # 5) Lectura y aplanado
        registros: List[Dict[str, Any]] = []
        hojas_con_datos: List[str] = []
        total = 0

        for hoja in hojas_a_leer:
            values = _read_used_range(token, drive_id, item_id, hoja)
            if not values:
                continue
            filas = _extract_rows_from_sheet(values)
            if filas:
                if nombre_campo_hoja:
                    for f in filas:
                        f[nombre_campo_hoja] = hoja
                registros.extend(filas)
                hojas_con_datos.append(hoja)
                total += len(filas)

        if not registros:
            return False, "No se extrajeron registros desde las hojas seleccionadas.", []

        msg_ok = (
            f"Se leyeron {total} registros de {len(hojas_con_datos)} hoja(s): "
            f"{', '.join(hojas_con_datos)}."
        )
        return True, msg_ok, registros

    except Exception as ex:
        return False, f"Error en la lectura del Excel: {ex}", []
        