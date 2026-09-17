# -*- coding: utf-8 -*-
"""
Obtiene usuarios del tenant (Microsoft Graph /users) para envío de encuestas.
- Paginación: sigue @odata.nextLink hasta agotar resultados
- Filtrado: accountEnabled eq true (solo habilitados); por defecto excluye invitados (userType eq 'Member')
- Campos seleccionados (select): id, displayName, mail, userPrincipalName, givenName, surname, jobTitle,
  department, mobilePhone, businessPhones, accountEnabled

Ubicación sugerida: app/services/obtener_usuarios.py
Asegúrate de tener:
  app/__init__.py
  app/services/__init__.py
"""

from typing import Any, Dict, List, Tuple, Optional
import json
import requests
from urllib.parse import urlencode

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Campos que usaremos para listar usuarios
USERS_SELECT = [
    "id",
    "displayName",
    "mail",
    "userPrincipalName",
    "givenName",
    "surname",
    "jobTitle",
    "department",
    "mobilePhone",
    "businessPhones",
    "accountEnabled",
    "userType"
]


def _get_graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """
    Obtiene access_token (client credentials) para Microsoft Graph.
    """
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(token_url, data=data, timeout=60)
    r.raise_for_status()
    return r.json()["access_token"]


def _graph_get(url: str, token: str, params: Optional[Dict[str, Any]] = None, headers_extra: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
    """
    GET con bearer token. Devuelve JSON dict (o lanza excepción HTTPError).
    """
    headers = {"Authorization": f"Bearer {token}"}
    if headers_extra:
        headers.update(headers_extra)
    r = requests.get(url, headers=headers, params=params, timeout=60)
    r.raise_for_status()
    return r.json()


def _iter_users(
    token: str,
    select_fields: List[str],
    page_size: int = 999,
    filter_: Optional[str] = None,
    include_consistency_header: bool = False
):
    """
    Itera TODOS los usuarios del tenant usando paginación por @odata.nextLink.
    Endpoint: GET /users con $select y $top (hasta 999 por página).

    Notas:
    - /users devuelve 100 por defecto y soporta $top con máximo 999.
    - Para ciertas consultas avanzadas puede requerirse ConsistencyLevel=eventual,
      aquí lo dejamos opcional (por defecto False).
    """
    base = f"{GRAPH_BASE}/users"

    odata = {
        "$select": ",".join(select_fields),
        "$top": str(page_size),
    }
    if filter_:
        odata["$filter"] = filter_

    headers_extra = {"ConsistencyLevel": "eventual"} if include_consistency_header else None
    url = base + "?" + urlencode(odata)

    while url:
        data = _graph_get(url, token, headers_extra=headers_extra)
        for user in data.get("value", []):
            yield user
        url = data.get("@odata.nextLink")  # Si hay más páginas, viene el link completo


def LeerUsuariosDelTenant(
    CONFIG: Dict[str, Any],
    *,
    only_enabled: bool = True,
    include_guests: bool = False,
    require_email: bool = True
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """
    Lee usuarios del tenant y devuelve una lista optimizada para envío de encuestas.

    Parámetros
    ----------
    CONFIG : Dict[str, Any] con credenciales:
      {
        "TENANT_ID": "...",
        "CLIENT_ID": "...",
        "CLIENT_SECRET": "..."
      }
    only_enabled : bool
        Si True, retorna solo usuarios con accountEnabled == True.
    include_guests : bool
        Si True, incluye invitados (userType == 'Guest'); por defecto solo 'Member'.
    require_email : bool
        Si True, filtra en cliente los usuarios que no tengan email resoluble (mail o UPN).

    Returns
    -------
    (success: bool, message: str, usuarios: List[Dict[str, Any]]), donde cada usuario:
      {
        "id": "...",
        "displayName": "...",
        "userPrincipalName": "...",
        "mail": "...",
        "email": "<mail or upn>",
        "givenName": "...",
        "surname": "...",
        "jobTitle": "...",
        "department": "...",
        "mobilePhone": "...",
        "businessPhones": [...],
        "accountEnabled": true/false,
        "userType": "Member"|"Guest"
      }
    """
    try:
        # Validación de credenciales
        for k in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"):
            if not CONFIG.get(k):
                return (False, f"CONFIG incompleto. Falta '{k}'.", [])

        # 1) Token
        token = _get_graph_token(CONFIG["TENANT_ID"], CONFIG["CLIENT_ID"], CONFIG["CLIENT_SECRET"])

        # 2) Construir filtro OData (en servidor)
        filters = []
        if only_enabled:
            filters.append("accountEnabled eq true")
        if not include_guests:
            filters.append("userType eq 'Member'")

        server_filter = " and ".join(filters) if filters else None

        # 3) Iterar /users con $select y $top=999, siguiendo @odata.nextLink
        usuarios: List[Dict[str, Any]] = []
        for u in _iter_users(
            token=token,
            select_fields=USERS_SELECT,
            page_size=999,
            filter_=server_filter,
            include_consistency_header=False  # ponlo True si luego usas $count/$search/$filter avanzados
        ):
            # Normalizar "email" (preferimos mail; si no, UPN)
            mail = u.get("mail")
            upn = u.get("userPrincipalName")
            email_resuelto = mail or upn

            # Si se requiere email, valida
            if require_email and not email_resuelto:
                continue

            usuarios.append({
                "id": u.get("id"),
                "displayName": u.get("displayName"),
                "userPrincipalName": upn,
                "mail": mail,
                "email": email_resuelto,
                "givenName": u.get("givenName"),
                "surname": u.get("surname"),
                "jobTitle": u.get("jobTitle"),
                "department": u.get("department"),
                "mobilePhone": u.get("mobilePhone"),
                "businessPhones": u.get("businessPhones", []),
                "accountEnabled": u.get("accountEnabled"),
                "userType": u.get("userType"),
            })

        return (True, f"Usuarios recuperados: {len(usuarios)}", usuarios)

    except requests.HTTPError as e:
        msg = ""
        try:
            msg = json.dumps(e.response.json(), ensure_ascii=False)
        except Exception:
            msg = getattr(e.response, "text", str(e))
        return (False, f"HTTPError: {msg}", [])
    except Exception as ex:
        return (False, f"Error inesperado: {ex}", [])
