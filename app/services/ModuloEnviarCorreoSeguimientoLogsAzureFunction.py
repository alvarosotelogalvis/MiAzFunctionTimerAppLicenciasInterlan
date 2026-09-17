from typing import Dict, Any, Tuple, List
from datetime import datetime, timezone
import json
import requests
import html


# ============================================================
# Helpers internos
# ============================================================

def _dedup_and_wrap(emails: List[str]) -> List[Dict[str, Any]]:
    """
    Elimina duplicados, valida correos básicos
    y los envuelve en el formato requerido por Microsoft Graph
    """
    seen = set()
    recipients = []

    for e in emails:
        if not isinstance(e, str):
            continue

        email = e.strip().lower()
        if not email or "@" not in email or email in seen:
            continue

        seen.add(email)
        recipients.append({
            "emailAddress": {
                "address": email
            }
        })

    return recipients


def _esc(value: Any) -> str:
    """Escapa texto para HTML"""
    return html.escape(str(value)) if value is not None else ""


def _send_graph_email(
    access_token: str,
    sender: str,
    subject: str,
    html_body: str,
    to_recipients: List[Dict[str, Any]]
) -> Tuple[bool, str]:
    """
    Envía correo usando Microsoft Graph
    """
    url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_body
            },
            "toRecipients": to_recipients
        },
        "saveToSentItems": True
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=30)
    except Exception as ex:
        return False, f"Error llamando Microsoft Graph: {ex}"

    if resp.status_code not in (200, 202):
        try:
            return False, json.dumps(resp.json(), ensure_ascii=False)
        except Exception:
            return False, resp.text

    return True, ""


# ============================================================
# Función principal
# ============================================================

def EnviarCorreoSeguimientoLogsAzureFunction(
    CONFIG: Dict[str, Any],
    MSGLOGS: Dict[str, Any],
    Destinatarios: Dict[str, Any]
) -> Tuple[bool, str]:
    """
    Envía un correo HTML con el seguimiento completo
    de la ejecución de la Azure Function.

    Destinatarios: dict con listas de emails, ej. {"Seguimiento": [...]}
    (viene de variables de entorno, ver MAIL_SEGUIMIENTO en function_app.py).

    Returns:
        (success: bool, message: str)
    """

    # ------------------ 1) Configuración ------------------
    TENANT_ID = str(CONFIG.get("TENANT_ID", "")).strip()
    CLIENT_ID = str(CONFIG.get("CLIENT_ID", "")).strip()
    CLIENT_SECRET = str(CONFIG.get("CLIENT_SECRET", "")).strip()
    SUBJECT_PREFIX = str(CONFIG.get("SUBJECT_PREFIX", "")).strip()
    SENDER = str(CONFIG.get("SENDER", "")).strip()

    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        return False, "CONFIG incompleto: faltan TENANT_ID, CLIENT_ID o CLIENT_SECRET."

    # ------------------ 2) Token Azure AD ------------------
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
        return False, f"Error solicitando token Azure AD: {ex}"

    if token_resp.status_code != 200:
        try:
            err = token_resp.json()
        except Exception:
            err = token_resp.text
        return False, f"Error token Azure AD ({token_resp.status_code}): {err}"

    access_token = token_resp.json().get("access_token")
    if not access_token:
        return False, "No se recibió access_token desde Azure AD."

    # ------------------ 3) Destinatarios ------------------
    list_seguimiento = [
        x for x in (Destinatarios.get("Seguimiento") or [])
        if isinstance(x, str)
    ]

    to_recipients = _dedup_and_wrap(list_seguimiento)
    if not to_recipients:
        return False, "No hay destinatarios válidos en 'Seguimiento'."

    # ------------------ 4) HTML del log ------------------
    inicio = MSGLOGS.get("Inicio", "")
    eventos = MSGLOGS.get("Eventos", [])

    filas_eventos = []
    for ev in eventos:
        filas_eventos.append(f"""
        <tr>
            <td>{_esc(ev.get("time", ""))}</td>
            <td>{_esc(ev.get("step", ""))}</td>
            <td>{_esc(ev.get("message", ""))}</td>
        </tr>
        """)

    tabla_eventos = "".join(filas_eventos) or """
        <tr>
            <td colspan="3"><em>No se registraron eventos.</em></td>
        </tr>
    """

    fecha_ejecucion = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    html_body = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Seguimiento ejecución Azure Function</title>
<style>
body {{
    font-family: Arial, Helvetica, sans-serif;
    color: #333;
}}
h2 {{
    color: #1a73e8;
}}
table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 12px;
}}
th {{
    background-color: #1a73e8;
    color: #fff;
    padding: 8px;
    text-align: left;
}}
td {{
    padding: 8px;
    border-bottom: 1px solid #ddd;
    font-size: 13px;
    vertical-align: top;
}}
.small {{
    color: #666;
    font-size: 12px;
}}
</style>
</head>
<body>

<h2>Seguimiento ejecución Azure Function</h2>

<p><strong>Inicio:</strong> {_esc(inicio)}</p>
<p><strong>Fecha envío:</strong> {_esc(fecha_ejecucion)}</p>

<table>
<thead>
<tr>
    <th>Hora</th>
    <th>Proceso</th>
    <th>Mensaje</th>
</tr>
</thead>
<tbody>
{tabla_eventos}
</tbody>
</table>

<p class="small">
Este correo corresponde al seguimiento automático de la ejecución de la Azure Function
<b>AppLicenciasInterlan</b>.
</p>

</body>
</html>
"""

    # ------------------ 5) Envío ------------------
    hoy_txt = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    subject_core = f"Seguimiento • Ejecución Azure Function • {hoy_txt}"
    subject = f"{SUBJECT_PREFIX} {subject_core}".strip() if SUBJECT_PREFIX else subject_core

    ok, err = _send_graph_email(
        access_token=access_token,
        sender=SENDER,
        subject=subject,
        html_body=html_body,
        to_recipients=to_recipients
    )

    if not ok:
        return False, f"Error enviando correo de seguimiento: {err}"

    return True, "Correo de seguimiento enviado correctamente."
    