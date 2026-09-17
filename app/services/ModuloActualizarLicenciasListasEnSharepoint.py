# -*- coding: utf-8 -*-
import time
from typing import Dict, Any, Tuple, List
from urllib.parse import quote
import requests


def ActualizarLicenciasListasEnSharepoint(
        CONFIG: Dict[str, Any],
        licencias: List[Dict[str, Any]],
        excel_config: Dict[str, Any]
    ) -> Tuple[bool, str]:
    """
    Actualiza un Excel en SharePoint (una sola hoja) con datos de licencias.
    Crea el archivo desde plantilla si no existe.

    Columnas fijas:
    id, Consecutivo, Fabricante, Producto, Cantidad, Vigencia,
    Fecha_x0020_Fin, Nombre_x0020_Cliente,
    Asesor_x0020_ComercialLookupId, Alerta, webUrl
    """

    # ---------------- Validaciones ----------------
    TENANT_ID = str(CONFIG.get("TENANT_ID", "")).strip()
    CLIENT_ID = str(CONFIG.get("CLIENT_ID", "")).strip()
    CLIENT_SECRET = str(CONFIG.get("CLIENT_SECRET", "")).strip()

    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        return False, "CONFIG incompleto: faltan TENANT_ID, CLIENT_ID o CLIENT_SECRET."

    if not isinstance(licencias, list):
        return False, "La lista de licencias no es válida."

    site_hostname = excel_config.get("site_hostname")
    site_path = excel_config.get("site_path")
    entorno = str(excel_config.get("entorno", "desarrollo")).lower()

    file_prod = excel_config.get("file_relative_path")
    file_dev = excel_config.get("file_relative_path_dev")
    template_path = excel_config.get("template_relative_path")
    sheet_name = excel_config.get("sheet_name", "Inventario")

    if not site_hostname or not site_path or not file_prod or not template_path:
        return False, "Configuración de SharePoint incompleta."

    dest_file = file_prod if entorno == "produccion" else file_dev

    # ---------------- Constantes ----------------
    HTTP_TIMEOUT = 30

    HEADERS = [
        "id",
        "Consecutivo",
        "Fabricante",
        "Producto",
        "Cantidad",
        "Vigencia",
        "Fecha_x0020_Fin",
        "Nombre_x0020_Cliente",
        "Asesor_x0020_ComercialLookupId",
        "Alerta",
        "webUrl",
    ]

    # ---------------- Graph Helpers ----------------
    def _graph(method, url, token, json_body=None, headers_extra=None):
        h = {"Authorization": f"Bearer {token}"}
        if json_body is not None:
            h["Content-Type"] = "application/json"
        if headers_extra:
            h.update(headers_extra)
        resp = requests.request(
            method, url, headers=h, json=json_body, timeout=HTTP_TIMEOUT
        )
        if resp.status_code in (200, 201, 202):
            return resp.status_code, (resp.json() if resp.text else {})
        raise RuntimeError(f"Graph error {resp.status_code}: {resp.text[:400]}")

    def _get_token():
        url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
        }
        r = requests.post(url, data=data, timeout=HTTP_TIMEOUT)
        if r.status_code != 200:
            raise RuntimeError(f"Token error: {r.text}")
        return r.json()["access_token"]

    # ---------------- Flujo principal ----------------
    try:
        token = _get_token()

        # Site
        site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
        _, site_data = _graph("GET", site_url, token)
        site_id = site_data["id"]

        # Drive
        _, drive_data = _graph("GET", f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", token)
        drive_id = drive_data["id"]

        # Resolver archivo
        def _get_item(path):
            enc = quote(path.strip("/"), safe="/-_.()")
            return _graph(
                "GET",
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{enc}",
                token,
            )

        try:
            _, item = _get_item(dest_file)
            item_id = item["id"]
        except Exception:
            # Copiar desde plantilla
            _, tpl = _get_item(template_path)
            parent = dest_file.rsplit("/", 1)[0]
            name = dest_file.rsplit("/", 1)[1]
            body = {
                "parentReference": {"path": f"/drive/root:/{parent}"},
                "name": name,
            }
            _graph(
                "POST",
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{tpl['id']}/copy",
                token,
                json_body=body,
            )
            time.sleep(5)
            _, item = _get_item(dest_file)
            item_id = item["id"]

        # Sesión Excel
        _, sess = _graph(
            "POST",
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/createSession",
            token,
            json_body={"persistChanges": True},
        )
        sid = sess["id"]
        headers_sess = {"workbook-session-id": sid}

        # Limpiar hoja
        sheet_enc = quote(sheet_name)
        _graph(
            "POST",
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_enc}/usedRange/clear",
            token,
            json_body={"applyTo": "all"},
            headers_extra=headers_sess,
        )

        # Encabezados
        _graph(
            "PATCH",
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_enc}/range(address='A1')",
            token,
            json_body={"values": [HEADERS]},
            headers_extra=headers_sess,
        )

        # Datos
        rows = [[lic.get(h) for h in HEADERS] for lic in licencias]

        if rows:
            _graph(
                "PATCH",
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_enc}/range(address='A2')",
                token,
                json_body={"values": rows},
                headers_extra=headers_sess,
            )

        return True, f"Excel actualizado correctamente ({len(rows)} registros)."

    except Exception as ex:
        return False, f"Error actualizando Excel en SharePoint: {ex}"
        