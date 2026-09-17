# -*- coding: utf-8 -*-
"""
ValidarEmailComercialesExcel
----------------------------

Clasifica licencias por validez de email en el tenant (Graph):

Salida:
    {
      "ListaNotificarLicenciasConEmail": [...],           # email válido en formato y activo en el tenant
      "ListaNotificarLincenciasComercialInactivos": [...],# formato inválido, no existe en tenant o deshabilitado
      "ListaNotificarLicenciasSinEmail": [...],           # sin email
      # "UsuariosGraphPreview": [...],                    # opcional, comentado
      "Totales": {
          "usuarios_graph": int,
          "con_email": int,
          "inactivos": int,
          "sin_email": int
      }
    }
"""

import re
import time
import json
import requests
from typing import Any, Dict, List, Tuple, Optional, Union

EMAIL_REGEX = re.compile(r"^[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

def _log(msg: str) -> None:
    try:
        print(msg)
    except Exception:
        pass

def _ensure_list_of_dicts(x: Any) -> List[Dict[str, Any]]:
    """
    Normaliza la entrada a una lista de dicts.
      - dict con claves 'vencen_1_mes'/'vencen_2_meses'/'vencen_3_meses' -> concatena sus listas.
      - list[dict] -> se deja
      - None/otros -> []
    """
    if x is None:
        return []
    if isinstance(x, dict):
        buckets = []
        for k in ("vencen_1_mes", "vencen_2_meses", "vencen_3_meses"):
            v = x.get(k)
            if isinstance(v, list):
                buckets.extend([it for it in v if isinstance(it, dict)])
        if buckets:
            return buckets
        # fallback: primer valor que sea list[dict]
        for v in x.values():
            if isinstance(v, list) and all(isinstance(it, dict) for it in v):
                return list(v)
        return []
    if isinstance(x, list):
        return [it for it in x if isinstance(it, dict)]
    return []

def _split_candidate_emails(raw: Any) -> List[str]:
    """
    Acepta:
      - string con uno o varios emails (separados por ; , o espacios)
      - lista/tupla de strings
    Devuelve emails en bruto (sin validar formato), limpios de espacios.
    """
    out: List[str] = []
    if raw is None:
        return out

    if isinstance(raw, (list, tuple)):
        for item in raw:
            if isinstance(item, str):
                out.extend(re.split(r"[;, \t\r\n]+", item.strip()))
        return [s for s in out if s]

    if isinstance(raw, str):
        parts = re.split(r"[;, \t\r\n]+", raw.strip())
        return [s for s in parts if s]

    # Otros tipos no soportados
    return out

def _extract_email_from_item(
    it: Dict[str, Any],
    campos_email: Union[str, List[str], Tuple[str, ...]]
) -> Tuple[Optional[str], bool]:
    """
    Busca un email en el item a partir de 1 o varias llaves posibles.
    - Devuelve el primer email con formato válido (EMAIL_REGEX) o None.
    - Segundo valor indica si el campo existía pero el formato fue inválido (para clasificar como inactivo).
    """
    keys: List[str]
    if isinstance(campos_email, str):
        keys = [campos_email]
    else:
        keys = list(campos_email or [])

    any_value_present = False
    for key in keys:
        raw_val = it.get(key)
        if raw_val is None:
            continue
        candidates = _split_candidate_emails(raw_val)
        if candidates:
            any_value_present = True
        for cand in candidates:
            em = cand.strip()
            if em and EMAIL_REGEX.match(em):
                return em, True  # email válido y había valor
    # Si llegamos aquí: no se encontró formato válido, pero puede que sí hubiera valor
    return None, any_value_present

def _get_graph_access_token(tenant_id: str, client_id: str, client_secret: str, timeout: int = 30, retries: int = 3) -> Tuple[bool, str]:
    """
    Client Credentials para Microsoft Graph (v2 endpoint) con reintentos básicos.
    """
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    backoff = 1.0
    for attempt in range(1, retries + 1):
        try:
            resp = requests.post(token_url, data=data, timeout=timeout)
        except Exception as ex:
            if attempt == retries:
                return False, f"Error HTTP al solicitar token: {ex}"
            time.sleep(backoff)
            backoff *= 2
            continue
        if resp.status_code == 200:
            try:
                tok = resp.json().get("access_token")
            except Exception as ex:
                return False, f"No se pudo parsear access_token: {ex}"
            if not tok:
                return False, "Respuesta de token sin 'access_token'."
            return True, tok
        # Reintento en 429/5xx si no es el último intento
        if resp.status_code in (429, 500, 502, 503, 504) and attempt < retries:
            retry_after = 1
            try:
                retry_after = int(resp.headers.get("Retry-After", "1"))
            except Exception:
                pass
            time.sleep(max(1, retry_after))
            continue
        # Error definitivo
        try:
            err = resp.json()
        except Exception:
            err = resp.text
        return False, f"Error token ({resp.status_code}): {err}"
    return False, "No fue posible obtener token (intentos agotados)."

def _graph_list_users(access_token: str, timeout: int = 30, max_pages: int = 50) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """
    Lista usuarios del tenant con paginación.
    Selecciona campos:
        id, displayName, mail, userPrincipalName, accountEnabled, proxyAddresses, userType
    Retorna (ok, msg, users)
    """
    base_url = "https://graph.microsoft.com/v1.0/users"
    params = {
        "$select": "id,displayName,mail,userPrincipalName,accountEnabled,proxyAddresses,userType",
        "$top": "999"
    }
    headers = {"Authorization": f"Bearer {access_token}"}
    users: List[Dict[str, Any]] = []

    url = base_url
    page = 0
    while url and page < max_pages:
        page += 1
        try:
            resp = requests.get(url, headers=headers, params=params if page == 1 else None, timeout=timeout)
        except Exception as ex:
            return False, f"Error HTTP al listar usuarios (p{page}): {ex}", users

        if resp.status_code != 200:
            # Reintento ligero en 429/5xx
            if resp.status_code in (429, 500, 502, 503, 504) and page <= max_pages:
                retry_after = 1
                try:
                    retry_after = int(resp.headers.get("Retry-After", "1"))
                except Exception:
                    pass
                time.sleep(max(1, retry_after))
                # reintenta misma URL en siguiente vuelta
                continue

            try:
                err = resp.json()
            except Exception:
                err = resp.text
            return False, f"Error Graph /users (p{page}) [{resp.status_code}]: {err}", users

        data = resp.json()
        chunk = data.get("value", [])
        if isinstance(chunk, list):
            users.extend(chunk)

        url = data.get("@odata.nextLink")  # siguiente página

    return True, f"Usuarios leídos: {len(users)}", users

def _emails_map_from_graph(users: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """
    Construye un mapa {email_lower: {accountEnabled, source, displayName, id, userType}}.
    Considera:
      - 'mail' (preferido)
      - 'userPrincipalName'
      - 'proxyAddresses' (SMTP: y smtp:)
    """
    out: Dict[str, Dict[str, Any]] = {}
    for u in users:
        enabled = bool(u.get("accountEnabled", False))
        name = u.get("displayName")
        uid = u.get("id")
        user_type = u.get("userType")  # 'Member' | 'Guest' | None

        candidates: List[Tuple[str, str]] = []

        mail = (u.get("mail") or "").strip()
        if mail:
            candidates.append(("mail", mail))

        upn = (u.get("userPrincipalName") or "").strip()
        if upn:
            candidates.append(("upn", upn))

        # proxyAddresses puede tener entradas como "SMTP:primary@contoso.com" y "smtp:alias@contoso.com"
        proxies = u.get("proxyAddresses") or []
        if isinstance(proxies, list):
            # Primero primarias (prefijo "SMTP:"), luego alias ("smtp:")
            primarias = [p for p in proxies if isinstance(p, str) and p.startswith("SMTP:")]
            alias = [p for p in proxies if isinstance(p, str) and p.startswith("smtp:")]
            for p in primarias + alias:
                addr = p.split(":", 1)[1].strip() if ":" in p else p
                if addr:
                    candidates.append(("proxy", addr))

        for src, addr in candidates:
            low = addr.lower()
            if not EMAIL_REGEX.match(low):
                continue
            # Prioridad: mail > proxy (primaria) > upn > proxy (alias)
            # Implementamos simple preferencia: si ya existe y la fuente anterior fue 'mail', no reemplazar.
            prev = out.get(low)
            if prev:
                if prev.get("source") == "mail":
                    continue
            out[low] = {
                "accountEnabled": enabled,
                "source": src,
                "displayName": name,
                "id": uid,
                "userType": user_type
            }
    return out

def _preview_users_for_output(users: List[Dict[str, Any]], limit: int = 200) -> List[Dict[str, Any]]:
    preview = []
    for u in users[:max(0, limit)]:
        preview.append({
            "displayName": u.get("displayName"),
            "mail": u.get("mail"),
            "userPrincipalName": u.get("userPrincipalName"),
            "accountEnabled": u.get("accountEnabled", None),
            "userType": u.get("userType", None),
        })
    return preview

def ValidarEmailComercialesExcel(
    CONFIG: Dict[str, Any],
    CamposEmail: Union[str, List[str], Tuple[str, ...]],
    LicenciasParaNotificar: Any
) -> Tuple[bool, str, Dict[str, Any]]:

    # 1) Validar CONFIG
    TENANT_ID = str(CONFIG.get("TENANT_ID", "")).strip()
    CLIENT_ID = str(CONFIG.get("CLIENT_ID", "")).strip()
    CLIENT_SECRET = str(CONFIG.get("CLIENT_SECRET", "")).strip()
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        return False, "CONFIG incompleto: faltan TENANT_ID, CLIENT_ID o CLIENT_SECRET.", {}

    # 2) Token Graph
    ok, token_or_msg = _get_graph_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    if not ok:
        return False, token_or_msg, {}
    access_token = token_or_msg

    # 3) Usuarios Graph
    ok, msg_users, users = _graph_list_users(access_token)
    if not ok:
        return False, f"No se pudieron leer usuarios del tenant: {msg_users}", {}

    # 4) Índice de emails válidos/activos del tenant
    email_map = _emails_map_from_graph(users)  # {email_lower: {...}}
    _log(f"[INFO] Usuarios Graph: {len(users)} | Emails indexados: {len(email_map)}")

    # 5) Normalizar licencias de entrada
    licencias = _ensure_list_of_dicts(LicenciasParaNotificar)

    # 6) Clasificación
    out_con_email: List[Dict[str, Any]] = []
    out_inactivos: List[Dict[str, Any]] = []
    out_sin_email: List[Dict[str, Any]] = []

    for it in licencias:
        em, had_value = _extract_email_from_item(it, CamposEmail)

        if not had_value:
            # No hay valor en el(los) campo(s) de email -> SinEmail
            out_sin_email.append(it)
            continue

        if em is None:
            # Había valor pero ningún email con formato válido -> Inactivos/Inválidos
            out_inactivos.append(it)
            continue

        # Email con formato válido -> verificar existencia + habilitación en tenant
        low = em.lower()
        info = email_map.get(low)

        # Regla base: debe existir en el tenant y estar accountEnabled=True
        if info and bool(info.get("accountEnabled", False)):
            # (Opcional) Si quieres excluir invitados ('Guest'), descomenta:
            # if info.get("userType") == "Guest":
            #     out_inactivos.append(it)
            #     continue
            out_con_email.append(it)
        else:
            out_inactivos.append(it)

    result = {
        "ListaNotificarLicenciasConEmail": out_con_email,
        "ListaNotificarLincenciasComercialInactivos": out_inactivos,
        "ListaNotificarLicenciasSinEmail": out_sin_email,
        # "UsuariosGraphPreview": _preview_users_for_output(users, limit=200),
        "Totales": {
            "usuarios_graph": len(users),
            "con_email": len(out_con_email),
            "inactivos": len(out_inactivos),
            "sin_email": len(out_sin_email),
        }
    }

    msg = (
        f"Usuarios Graph: {len(users)} | "
        f"ConEmail: {len(out_con_email)} | "
        f"Inactivos/Inválidos: {len(out_inactivos)} | "
        f"SinEmail: {len(out_sin_email)}"
    )
    return True, msg, result
    