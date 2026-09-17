# -*- coding: utf-8 -*-
"""
Módulo para leer licencias desde SharePoint (Microsoft Graph) con paginación robusta.
- Usa @odata.nextLink para traer TODOS los ítems.
- Maneja errores 422 (List View Threshold) con reintentos/fallbacks:
  * Quita $orderby si causa el umbral.
  * Reduce $top (tamaño de página) si es necesario.
- Retorna items aplanados (campos en un solo nivel) listos para procesar.

Entradas requeridas en CONFIG:
  - TENANT_ID, CLIENT_ID, CLIENT_SECRET
  - SITE_ID (formato: host,siteCollectionId,siteId)
  - LIST_DISPLAY_NAME (nombre visible de la lista)

Salida de LeerLicenciasDeSharePoint(CONFIG):
  (True, None, lista_items)  o  (False, "mensaje de error", None)
"""

from __future__ import annotations

import json
import time
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple
import requests

# =========================
# PARÁMETROS DE RED / RETRIES
# =========================
MAX_RETRIES = 5
INITIAL_BACKOFF = 1.5  # segundos
DEFAULT_TIMEOUT = 30   # segundos

# Proxy (si aplica en tu red)
PROXIES: Optional[Dict[str, str]] = None
# Ejemplo:
# PROXIES = {
#     "http":  "http://usuario:password@proxy.miempresa.com:8080",
#     "https": "http://usuario:password@proxy.miempresa.com:8080",
# }

# =========================
# HELPERS GENERALES
# =========================

def odata_escape_literal(s: str) -> str:
    """Escapa comillas simples para literales OData (' -> '')."""
    return s.replace("'", "''") if isinstance(s, str) else s

def fecha_ddmmaaaa_a_iso(fecha_str: str) -> str:
    """Convierte 'DD/MM/AAAA' a 'YYYY-MM-DDT00:00:00Z'."""
    dt = datetime.strptime(fecha_str, "%d/%m/%Y")
    return dt.strftime("%Y-%m-%dT00:00:00Z")

def ensure_iso_date(value) -> str:
    """
    Normaliza fechas a ISO 8601 UTC.
    Acepta: 'DD/MM/AAAA', 'YYYY-MM-DD', datetime.date, datetime, o ISO ya válido.
    """
    if isinstance(value, str):
        if "/" in value:  # DD/MM/AAAA
            return fecha_ddmmaaaa_a_iso(value)
        # YYYY-MM-DD
        try:
            dt = datetime.strptime(value, "%Y-%m-%d")
            return dt.strftime("%Y-%m-%dT00:00:00Z")
        except ValueError:
            pass
        # Ya en ISO con Z
        if value.endswith("Z") and "T" in value:
            return value
        raise ValueError(f"Formato de fecha no reconocido: {value}")
    elif isinstance(value, datetime):
        return value.strftime("%Y-%m-%dT%H:%M:%SZ")
    elif isinstance(value, date):
        return value.strftime("%Y-%m-%dT00:00:00Z")
    else:
        raise ValueError(f"Tipo de fecha no soportado: {type(value)}")

def request_with_retries(method: str, url: str, headers: dict, **kwargs) -> requests.Response:
    """
    Solicitudes HTTP con reintentos ante 429/5xx (respeta Retry-After si viene).
    Para 4xx distintos a 429 no reintenta (el caller decide si aplicar fallback).
    """
    attempt = 0
    backoff = INITIAL_BACKOFF
    kwargs.setdefault("proxies", PROXIES)
    kwargs.setdefault("timeout", DEFAULT_TIMEOUT)

    while True:
        resp = requests.request(method, url, headers=headers, **kwargs)
        # Éxito
        if resp.status_code < 400:
            return resp

        # Too Many Requests o 5xx -> aplicar backoff exponencial
        if resp.status_code == 429 or 500 <= resp.status_code < 600:
            attempt += 1
            if attempt > MAX_RETRIES:
                return resp
            wait = backoff
            retry_after = resp.headers.get("Retry-After")
            if retry_after:
                try:
                    wait = int(retry_after)
                except ValueError:
                    pass
            time.sleep(wait)
            backoff = min(backoff * 2, 30)
            continue

        # Otros errores 4xx: no reintentar aquí (devolver para que el caller decida)
        return resp

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Obtiene token de Graph (application permissions)."""
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(token_url, data=data, proxies=PROXIES, timeout=DEFAULT_TIMEOUT)
    try:
        resp.raise_for_status()
    except requests.HTTPError:
        raise RuntimeError(f"Error obteniendo token: {resp.status_code} {resp.text}")
    return resp.json()["access_token"]

def get_headers(access_token: str) -> dict:
    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

# =========================
# LISTA: RESOLVER Y VERIFICAR
# =========================

def obtener_list_id_por_nombre(site_id: str, display_name: str, headers: dict) -> str:
    """
    Busca la lista por su displayName y devuelve el list_id.
    """
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    params = {"$filter": f"displayName eq '{odata_escape_literal(display_name)}'"}
    resp = request_with_retries("GET", url, headers=headers, params=params)
    try:
        resp.raise_for_status()
    except requests.HTTPError:
        raise RuntimeError(f"Error listando listas: {resp.status_code} {resp.text}")
    value = resp.json().get("value", [])
    if not value:
        raise ValueError(f"No se encontró una lista con nombre: {display_name}")
    return value[0]["id"]

# =========================
# LECTURA TOTAL (PAGINADA) CON FALLBACKS 422
# =========================

def leer_todos_los_elementos(
    site_id: str,
    list_id: str,
    headers: dict,
    top: int = 200,
    orderby: Optional[str] = "id asc",
    fields_select: Optional[List[str]] = None,
    expand_fields: bool = True,
) -> List[dict]:
    """
    Lee TODOS los items de la lista paginando con @odata.nextLink.

    Estrategia ante List View Threshold (HTTP 422 notSupported/tooManyResources):
      1) Intento inicial: $expand=fields, $orderby=id asc, $top=top
      2) Si 422: reintentar SIN $orderby (muchas vistas no indexadas fallan al ordenar)
      3) Si aún 422: reducir $top a 50
      4) Si persiste: levantar error (recomendar indexar campos o usar $filter)

    Parámetros:
      - top: tamaño de página solicitado (Graph puede limitarlo internamente).
      - orderby: orden recomendado "id asc" (columna interna ID suele estar indexada).
      - fields_select: lista de campos a seleccionar (además de id/webUrl). Normalmente no se requiere si expandimos fields.
      - expand_fields: si True, añade "$expand=fields" para traer los campos de la lista.
    """
    base = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"

    def _build_params(_top: int, _orderby: Optional[str], _expand_fields: bool) -> dict:
        p = {"$top": int(_top)}
        if _expand_fields:
            p["$expand"] = "fields"
        if _orderby:
            p["$orderby"] = _orderby
        if fields_select:
            # Nota: $select en /items aplica a propiedades del item; los fields vienen en "fields"
            # Si quieres limitar fields, usa $select sobre /items?expand=fields($select=Campo1,Campo2)
            # Por simplicidad, no filtramos fields aquí.
            p["$select"] = "id,webUrl"
        return p

    def _descargar_en_bucle(url_inicial: str, params_iniciales: Optional[dict]) -> List[dict]:
        items: List[dict] = []
        url = url_inicial
        params = params_iniciales
        while url:
            resp = request_with_retries("GET", url, headers=headers, params=params)
            # Manejo explícito de 422 aquí (para aplicar fallback desde el caller)
            if resp.status_code == 422:
                # devolvemos la respuesta para que el caller elija un fallback
                return ["__HTTP_422__", resp]
            resp.raise_for_status()

            data = resp.json()
            lote = data.get("value", [])
            items.extend(lote)

            next_link = data.get("@odata.nextLink")
            if not next_link:
                break
            url = next_link
            params = None  # el nextLink ya incluye los params
            time.sleep(0.05)  # cortesía
        return items

    # ---- Intento 1: con orderby (id asc por defecto), top como pide CONFIG ----
    params = _build_params(top, orderby, expand_fields)
    items = _descargar_en_bucle(base, params)
    if isinstance(items, list) and items and items[0] == "__HTTP_422__":
        # ---- Intento 2: SIN orderby ----
        _, resp422 = items
        # Log opcional:
        print(f"[AVISO] 422 con orderby='{orderby}' y top={top}. Mensaje: {resp422.text[:200]}...")
        params2 = _build_params(top, None, expand_fields)
        items = _descargar_en_bucle(base, params2)

        if isinstance(items, list) and items and items[0] == "__HTTP_422__":
            # ---- Intento 3: SIN orderby y top reducido a 50 ----
            _, resp422b = items
            print(f"[AVISO] 422 aún sin orderby. Reduciendo top a 50. Mensaje: {resp422b.text[:200]}...")
            params3 = _build_params(50, None, expand_fields)
            items = _descargar_en_bucle(base, params3)

            if isinstance(items, list) and items and items[0] == "__HTTP_422__":
                # Sigue fallando: escalar error para que el caller informe
                _, resp422c = items
                raise RuntimeError(
                    "HTTP 422 (List View Threshold) persistente incluso sin orderby y top=50. "
                    "Sugerencias: crear índices en la lista, usar filtros ($filter) por columnas indexadas, "
                    "o reducir la complejidad de la vista."
                    f" Detalle: {resp422c.text}"
                )

    # Aquí items es la lista final de ítems crudos.
    return items

# =========================
# UTILIDADES (APLANADO y DUPLICADOS)
# =========================

def aplanar_items(items: List[dict]) -> List[dict]:
    """
    Convierte cada item a un dict plano con:
      - id (string)
      - webUrl (si está presente)
      - todos los fields (clave/valor) fusionados al mismo nivel
    """
    planos = []
    for it in items:
        plano = {}
        if "id" in it:
            plano["id"] = it["id"]
        if "webUrl" in it:
            plano["webUrl"] = it["webUrl"]

        f = it.get("fields", {}) or {}
        if isinstance(f, dict):
            for k, v in f.items():
                if str(k).startswith("@odata."):
                    continue
                plano[k] = v
        planos.append(plano)
    return planos

def _norm(s):
    """Normaliza ID0 para comparar: quita espacios, minúsculas y colapsa espacios internos."""
    if s is None:
        return None
    if not isinstance(s, str):
        s = str(s)
    s = s.strip().lower()
    s = " ".join(s.split())
    return s

def detectar_duplicados_por_id0(items_aplanados: List[dict]):
    """
    items_aplanados: salida de aplanar_items (lista de dicts con 'ID0' y 'id')
    Retorna:
      - duplicados (dict)
      - sin_id0 (lista)
      - resumen (dict)
    """
    from collections import defaultdict

    grupos = defaultdict(list)
    sin_id0 = []
    for it in items_aplanados:
        id0 = it.get("ID0")
        if id0 in (None, "", "null", "nan", "NaN", "NULL"):
            sin_id0.append({
                "id": it.get("id"),
                "webUrl": it.get("webUrl"),
                "ID0": id0
            })
            continue
        clave = _norm(id0)
        grupos[clave].append({
            "id": it.get("id"),
            "ID0": id0,
            "webUrl": it.get("webUrl")
        })

    duplicados: Dict[str, Any] = {}
    total_grupos = 0
    total_items_en_grupos = 0
    for clave, lista in grupos.items():
        total_grupos += 1
        total_items_en_grupos += len(lista)
        if len(lista) > 1:
            try:
                ordenados = sorted(lista, key=lambda x: int(x["id"]))
            except Exception:
                ordenados = sorted(lista, key=lambda x: str(x["id"]))
            duplicados[clave] = {
                "id0_originales": sorted({x["ID0"] for x in lista}),
                "items": ordenados
            }

    resumen = {
        "total_items_aplanados": len(items_aplanados),
        "total_grupos_por_id0": total_grupos,
        "total_items_en_grupos": total_items_en_grupos,
        "total_items_sin_id0": len(sin_id0),
        "total_claves_con_duplicados": len(duplicados),
        "total_items_en_duplicados": sum(len(v["items"]) for v in duplicados.values()),
    }
    return duplicados, sin_id0, resumen

# =========================
# FUNCIÓN PÚBLICA (RECIBE CONFIG) → (bool, str, List[Dict]])
# =========================

def LeerUnalistaDeSharepoint(CONFIG: Dict[str, Any]) -> Tuple[bool, str | None, List[Dict[str, Any]] | None]:
    """
    Lee licencias desde SharePoint (Microsoft Graph) usando parámetros de CONFIG.

    Retorna:
      - OK:    (True, None, lista_items_planos)
      - ERROR: (False, "mensaje de error", None)
    """
    print('LeerLicenciasDeSharePoint...')
    try:
        # ===== 1) Validación de entradas =====
        tenant_id        = CONFIG.get("TENANT_ID")
        client_id        = CONFIG.get("CLIENT_ID")
        client_secret    = CONFIG.get("CLIENT_SECRET")
        site_id          = CONFIG.get("SITE_ID")
        list_display     = CONFIG.get("LIST_DISPLAY_NAME")

        top              = int(CONFIG.get("TOP", 50))
        orderby          = CONFIG.get("ORDERBY", "id asc")
        _usuario_forti   = CONFIG.get("USUARIO_PLAT_FORTI")  # no usado aquí
        _passw_forti     = CONFIG.get("PASSW_PLAT_FORTI")    # no usado aquí
        _imprimir_columnas = bool(CONFIG.get("IMPRIMIR_COLUMNAS", False))

        missing = [k for k, v in {
            "TENANT_ID": tenant_id,
            "CLIENT_ID": client_id,
            "CLIENT_SECRET": client_secret,
            "SITE_ID": site_id,
            "LIST_DISPLAY_NAME": list_display
        }.items() if not v]
        if missing:
            return False, f"Error: faltan claves requeridas en CONFIG: {', '.join(missing)}", None

        # ===== 2) Token y headers =====
        access_token = get_access_token(tenant_id, client_id, client_secret)
        headers = get_headers(access_token)

        # ===== 3) Resolver lista =====
        list_id = obtener_list_id_por_nombre(site_id, list_display, headers)
        print(f"[INFO] list_id resuelto: {list_id}")

        # ===== 4) Leer items paginados (con fallbacks 422) =====
        items = leer_todos_los_elementos(
            site_id=site_id,
            list_id=list_id,
            headers=headers,
            top=top,
            orderby=orderby,        # "id asc" por defecto
            fields_select=None,     # traer todos los fields (expand=fields)
            expand_fields=True
        )
        print(f"[INFO] Elementos leídos (crudo): {len(items)}")

        # ===== 5) Ordenar crudo por id =====
        try:
            items_ordenados = sorted(items, key=lambda x: int(x.get("id", 0)))
        except Exception:
            items_ordenados = sorted(items, key=lambda x: str(x.get("id", "")))

        # ===== 6) Aplanar & ordenar plano =====
        planos = aplanar_items(items_ordenados)
        try:
            planos_ordenados = sorted(planos, key=lambda x: int(x.get("id", 0)))
        except Exception:
            planos_ordenados = sorted(planos, key=lambda x: str(x.get("id", "")))

        # ===== 7) Retorno esperado =====
        total = len(planos_ordenados)
        mensaje_ok = f"Se leyeron {total} registros desde la lista '{list_display}'."
        return True, mensaje_ok, planos_ordenados

    except requests.HTTPError as http_err:
        return False, f"Error: HTTP {http_err.response.status_code} {http_err.response.text}", None
    except Exception as e:
        return False, f"Error: {str(e)}", None
