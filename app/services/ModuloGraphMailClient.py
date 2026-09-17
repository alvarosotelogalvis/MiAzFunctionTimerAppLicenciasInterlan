# -*- coding: utf-8 -*-
"""
Cliente compartido para autenticar y enviar correos vía Microsoft Graph.

Extrae la lógica que antes estaba duplicada en cada módulo de notificación
(ModuloNotificacionVencimientosExcel.py / ModuloNotificacionVencimientos2.py /
ModuloEnviarCorreoSeguimientoLogsAzureFunction.py). Lo usan el módulo fusionador
(ModuloEnviarNotificacionesCombinadas.py) y el correo de alerta de fallo, para no
volver a duplicar el mismo código de autenticación/envío una tercera vez.
"""

import re
import time
import requests
from typing import Any, Dict, List, Tuple

EMAIL_REGEX = re.compile(r"^[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")


def ObtenerTokenGraph(CONFIG: Dict[str, Any]) -> Tuple[bool, str]:
    """
    Obtiene un access_token app-only de Microsoft Graph (client credentials).

    Returns:
        (success, access_token_o_mensaje_de_error)
    """
    TENANT_ID = str(CONFIG.get("TENANT_ID", "")).strip()
    CLIENT_ID = str(CONFIG.get("CLIENT_ID", "")).strip()
    CLIENT_SECRET = str(CONFIG.get("CLIENT_SECRET", "")).strip()

    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        return False, "CONFIG incompleto: faltan TENANT_ID, CLIENT_ID o CLIENT_SECRET."

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    try:
        token_resp = requests.post(token_url, data=token_data, timeout=30)
    except Exception as ex:
        return False, f"Error al solicitar token a Azure AD: {ex}"
    if token_resp.status_code != 200:
        try:
            err = token_resp.json()
        except Exception:
            err = token_resp.text
        return False, f"Error al obtener token ({token_resp.status_code}): {err}"
    access_token = token_resp.json().get("access_token")
    if not access_token:
        return False, "No se recibió access_token en la respuesta de Azure AD."
    return True, access_token


def DedupYEnvolver(addresses: List[str]) -> List[Dict[str, Dict[str, str]]]:
    """Deduplica y valida una lista de correos, y los envuelve en el formato de Graph."""
    env = []
    vistos = set()
    for a in addresses or []:
        if not isinstance(a, str):
            continue
        low = a.strip().lower()
        if not (low and EMAIL_REGEX.match(low)):
            continue
        if low in vistos:
            continue
        vistos.add(low)
        env.append({"emailAddress": {"address": low}})
    return env


def EnviarCorreoGraph(
    access_token: str,
    sender: str,
    subject: str,
    html: str,
    to_recipients: List[Dict[str, Dict[str, str]]],
    max_attempts: int = 3
) -> Tuple[bool, str]:
    """Envía un correo HTML vía Graph sendMail, con reintentos ante errores transitorios."""
    if not to_recipients:
        return False, "Sin destinatarios TO."
    email_url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": to_recipients
        },
        "saveToSentItems": True
    }
    for attempt in range(1, max_attempts + 1):
        try:
            resp = requests.post(email_url, headers=headers, json=payload, timeout=30)
            status = resp.status_code
            if status in (200, 202):
                return True, ""
            if status in (429, 500, 502, 503, 504):
                time.sleep(min(30, 2 ** attempt))
                continue
            try:
                detalle = resp.json()
            except Exception:
                detalle = resp.text
            return False, f"[{status}] {str(detalle)[:300]}"
        except Exception as ex:
            if attempt == max_attempts:
                return False, f"[exception] {ex}"
            time.sleep(min(30, 2 ** attempt))
    return False, "Error desconocido al enviar correo."
