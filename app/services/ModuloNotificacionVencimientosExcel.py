# -*- coding: utf-8 -*-
"""
NotificarVencimientosListas (envíos + sello en campo Notificaciones + retorno de 3 listas)
-----------------------------------------------------------------------------------------

- Envía UN solo correo a Seguimiento con el resumen completo (con email, sin destinatario y emails no válidos/inactivos).
- Envía UN solo correo (consolidado) a SinDestinatarios (incluyendo sin destinatario e no válidos/inactivos).
- Mantiene correos individuales a comerciales con email válido.
- Sella el campo "Notificaciones" con 'dd-mm-yyyy-N-OK' (N=3,2,1) según días a vencer:
    N=1:  0 <= días < 30
    N=2: 30 <= días < 60
    N=3: 60 <= días <= 90
- **Devuelve** las 3 listas (con los ítems mutados en memoria, ya sellados cuando corresponde).

CONFIG opcional:
  - SELLAR_CONSOLIDADO: bool = True
  - CONTAR_VENCIDOS_COMO_N1: bool = False
  - SUBJECT_PREFIX: str
  - SENDER: str  (usuario/casilla que envía por Graph)

Requiere:
  - TENANT_ID, CLIENT_ID, CLIENT_SECRET (App Graph con permisos .default)
"""

import requests
import time
from typing import Any, Dict, List, Tuple, Optional, Union
from datetime import datetime, timezone, date
import json
import re

# --------------------- Constantes de campos ---------------------
CAMPO_VENCIMIENTO = "FechaDeVencimientoDDMMAAAA"
CAMPO_NOTIFICACIONES = "Notificaciones"

# --------------------- Utilidades generales ---------------------

EMAIL_REGEX = re.compile(r"^[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

def _log(msg: str) -> None:
    try:
        print(msg)
    except Exception:
        try:
            safe = msg.encode("ascii", errors="replace").decode("ascii", errors="ignore")
            print(safe)
        except Exception:
            pass

def _esc(x: Any) -> str:
    s = "" if x is None else str(x)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _try_json_loads(s: str):
    try:
        return json.loads(s)
    except Exception:
        return None

def _ensure_list_of_dicts(x: Any, label: str = "") -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    if x is None:
        return out
    if isinstance(x, str):
        parsed = _try_json_loads(s=x)
        if isinstance(parsed, list):
            x = parsed
        else:
            return out
    if isinstance(x, list):
        for it in x:
            if isinstance(it, dict):
                out.append(it)
            elif isinstance(it, str):
                parsed_item = _try_json_loads(it)
                if isinstance(parsed_item, dict):
                    out.append(parsed_item)
        return out
    return out

def _parse_fecha(fecha_val: Any) -> Optional[date]:
    if fecha_val is None:
        return None
    if isinstance(fecha_val, date) and not isinstance(fecha_val, datetime):
        return fecha_val
    if isinstance(fecha_val, datetime):
        return fecha_val.date()
    s = str(fecha_val).strip()
    if not s:
        return None
    try:
        if "T" in s:
            s2 = s.split("T", 1)[0]
            return datetime.strptime(s2, "%Y-%m-%d").date()
    except Exception:
        pass
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        pass
    try:
        return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        pass
    try:
        return datetime.strptime(s, "%d-%m-%Y").date()
    except Exception:
        pass
    return None

def _fmt_fecha_corta(value: Any) -> str:
    d = _parse_fecha(value)
    if d is None:
        return str(value) if value is not None else ""
    return d.strftime("%Y-%m-%d")

def _days_to_due(fecha_venc: Any, hoy: date) -> Optional[int]:
    d = _parse_fecha(fecha_venc)
    if d is None:
        return None
    return (d - hoy).days

def _clasificar_por_ventana(items: List[Dict[str, Any]], hoy: date) -> Dict[str, List[Dict[str, Any]]]:
    out = {"v30": [], "v60": [], "v90": []}
    for it in items:
        if not isinstance(it, dict):
            continue
        dias = _days_to_due(it.get(CAMPO_VENCIMIENTO), hoy)
        if dias is None:
            continue
        if 0 <= dias <= 30:
            out["v30"].append(it)
        elif 31 <= dias <= 60:
            out["v60"].append(it)
        elif 61 <= dias <= 90:
            out["v90"].append(it)
    return out

def _fila(item: Dict[str, Any]) -> str:
    fecha_corta = _fmt_fecha_corta(item.get(CAMPO_VENCIMIENTO))
    return (
        f"<tr>"
        f"<td>{_esc(item.get('LicenciaID', ''))}</td>"
        f"<td>{_esc(item.get('Serial', ''))}</td>"
        f"<td>{_esc(item.get('Tipo', ''))}</td>"
        f"<td>{_esc(item.get('Version', ''))}</td>"
        f"<td>{_esc(item.get('Cantidad', ''))}</td>"
        f"<td>{_esc(item.get('Cliente', ''))}</td>"
        f"<td>{_esc(item.get('Fabricante', ''))}</td>"
        f"<td>{_esc(fecha_corta)}</td>"
        f"</tr>"
    )

def _tabla(items: List[Dict[str, Any]]) -> str:
    return "\n".join(_fila(it) for it in items)

def _section_html(titulo: str, badge_html: str, items: List[Dict[str, Any]]) -> str:
    if not items:
        return ""
    rows = _tabla(items)
    return f"""
<h3>{titulo} {badge_html}</h3>
<table>
  <thead>
    <tr>
      <th>LicenciaID</th>
      <th>Serial</th>
      <th>Tipo</th>
      <th>Version</th>
      <th>Cantidad</th>
      <th>Cliente</th>
      <th>Fabricante</th>
      <th>FechaVencimiento</th>
    </tr>
  </thead>
  <tbody>
    {rows}
  </tbody>
</table>
"""

def _section_html_simple(titulo: str, items: List[Dict[str, Any]]) -> str:
    if not items:
        return ""
    rows = _tabla(items)
    return f"""
<h3>{_esc(titulo)}</h3>
<table>
  <thead>
    <tr>
      <th>LicenciaID</th>
      <th>Serial</th>
      <th>Tipo</th>
      <th>Version</th>
      <th>Cantidad</th>
      <th>Cliente</th>
      <th>Fabricante</th>
      <th>FechaVencimiento</th>
    </tr>
  </thead>
  <tbody>
    {rows}
  </tbody>
</table>
"""

def _contenido_v3090(v30: List[Dict[str, Any]],
                     v60: List[Dict[str, Any]],
                     v90: List[Dict[str, Any]]) -> str:
    badge_30 = '<span style="display:inline-block;padding:2px 8px;border-radius:12px;font-size:12px;color:#ffffff;background-color:#d93025;vertical-align:middle;margin-left:6px;">Prioridad Alta</span>'
    badge_60 = '<span style="display:inline-block;padding:2px 8px;border-radius:12px;font-size:12px;color:#ffffff;background-color:#f29900;vertical-align:middle;margin-left:6px;">Prioridad Media</span>'
    badge_90 = '<span style="display:inline-block;padding:2px 8px;border-radius:12px;font-size:12px;color:#222222;background-color:#e5b700;vertical-align:middle;margin-left:6px;">Seguimiento</span>'
    seccion_30 = _section_html("Vencimiento en menos de 30 días", badge_30, v30)
    seccion_60 = _section_html("Vencimiento entre 31 y 60 días", badge_60, v60)
    seccion_90 = _section_html("Vencimiento entre 61 y 90 días", badge_90, v90)
    contenido = "".join([seccion_30, seccion_60, seccion_90]).strip()
    return contenido if contenido else '<p><em>No se encontraron elementos para mostrar.</em></p>'

def _html_body_multi(secciones: List[Dict[str, Any]],
                     saludo: str = "Equipo Comercial",
                     titulo_principal: str = "Notificación de Vencimiento de Licencias y/o Servicios",
                     intro_html: Optional[str] = None) -> str:
    bloques = []
    for sec in secciones:
        t = _esc(sec.get("titulo", ""))
        nota_html = sec.get("nota_html") or ""
        contenido = _contenido_v3090(sec.get("v30", []), sec.get("v60", []), sec.get("v90", []))
        nota_block = ""
        if nota_html:
            nota_block = f"""
<div style="background:#fff8e1;border:1px solid #f1c40f;color:#7d6608;border-radius:6px;padding:10px 12px;margin:12px 0;">
  <strong>Nota:</strong> {nota_html}
</div>
"""
        bloques.append(f"""
<h2>{t}</h2>
{nota_block}
{contenido}
""")

    intro_block = ""
    if intro_html:
        intro_block = f"""
<div style="background:#eef6ff;border:1px solid #b3d4ff;color:#1a4d8f;border-radius:6px;padding:10px 12px;margin:12px 0;">
  {intro_html}
</div>
"""

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>{_esc(titulo_principal)}</title>
<style>
  body {{
    font-family: Arial, Helvetica, sans-serif;
    color: #333333;
    line-height: 1.35;
  }}
  h1, h2 {{
    color: #1a73e8;
    font-weight: 600;
    margin: 18px 0 6px 0;
  }}
  h3 {{
    color: #444444;
    margin: 14px 0 4px 0;
    font-weight: 600;
  }}
  p {{ margin: 6px 0; }}
  table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 12px;
  }}
  th {{
    background-color: #1a73e8;
    color: #ffffff;
    padding: 8px;
    text-align: left;
    font-size: 13px;
    border: 1px solid #d0d7de;
  }}
  td {{
    padding: 8px;
    border-bottom: 1px solid #dddddd;
    border-left: 1px solid #f0f0f0;
    border-right: 1px solid #f0f0f0;
    font-size: 13px;
    vertical-align: top;
  }}
  .small {{ color: #666; font-size: 12px; }}
</style>
</head>
<body>

<p>Estimado(a) <strong>{_esc(saludo)}</strong>,</p>

{intro_block}

{''.join(bloques)}

<p class="small">
Este mensaje es informativo y forma parte del proceso de aseguramiento de continuidad operativa y seguridad.
</p>

<p>Atentamente,<br>
<strong>Equipo comercial</strong><br>
<strong>Interlan</strong><br>
</p>

</body>
</html>"""

def _html_body_info_simple(fecha_txt: str, titulo: str, mensaje: str) -> str:
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>{_esc(titulo)}</title>
<style>
  body {{
    font-family: Arial, Helvetica, sans-serif;
    color: #333333;
    line-height: 1.35;
  }}
  h2 {{
    color: #1a73e8;
    font-weight: 600;
  }}
  p {{ margin: 6px 0; }}
  .small {{ color: #666; font-size: 12px; }}
</style>
</head>
<body>
  <h2>{_esc(titulo)}</h2>
  <p>Fecha de ejecución: <strong>{_esc(fecha_txt)}</strong></p>
  <p>{_esc(mensaje)}</p>
  <p class="small">Interlan – Calidad e Innovación</p>
</body>
</html>"""

def _dedup_and_wrap(addresses: List[str]) -> List[Dict[str, Dict[str, str]]]:
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

def _extract_email_from_item(it: Dict[str, Any], campos_email: Union[str, List[str], Tuple[str, ...]]) -> Optional[str]:
    if isinstance(campos_email, str):
        candidates = [campos_email]
    else:
        candidates = list(campos_email or [])
    for c in candidates:
        val = it.get(c)
        if isinstance(val, str):
            em = val.strip()
            if em and EMAIL_REGEX.match(em):
                return em
    return None

def _send_graph_email(access_token: str, sender: str, subject: str, html: str,
                      to_recipients: List[Dict[str, Dict[str, str]]],
                      max_attempts: int = 3) -> Tuple[bool, str]:
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
                _log(f"[SEND OK] Subject='{subject}' To={[r['emailAddress']['address'] for r in to_recipients]}")
                return True, ""
            if status in (429, 500, 502, 503, 504):
                _log(f"[SEND RETRY {attempt}] status={status} body={resp.text[:160]}")
                time.sleep(min(30, 2 ** attempt))
                continue
            try:
                detalle = resp.json()
            except Exception:
                detalle = resp.text
            _log(f"[SEND ERR] status={status} body={str(detalle)[:500]}")
            return False, f"[{status}] {str(detalle)[:300]}"
        except Exception as ex:
            if attempt == max_attempts:
                _log(f"[SEND EXC] {ex}")
                return False, f"[exception] {ex}"
            _log(f"[SEND EXC RETRY {attempt}] {ex}")
            time.sleep(min(30, 2 ** attempt))
    return False, "Error desconocido al enviar correo."

# --------------------- Helpers para campo Notificaciones ---------------------

_RE_SEG_NOTIF = re.compile(r"(?<!\d)(\d{2})\d{2}\d{4}-(1|2|3)-OK(?!\d)", re.IGNORECASE)

def _calc_n_por_dias(diff: int) -> Optional[int]:
    if diff < 0:
        return None
    if 0 <= diff < 30:
        return 1
    if 30 <= diff < 60:
        return 2
    if 60 <= diff <= 90:
        return 3
    return None

def _parse_notif_field(raw: str) -> Dict[int, Tuple[date, str]]:
    out: Dict[int, Tuple[date, str]] = {}
    if not isinstance(raw, str) or not raw.strip():
        return out
    _, _, cola = raw.partition("Notif:")
    texto = cola if cola else raw
    for m in _RE_SEG_NOTIF.finditer(texto):
        dd, mm, yyyy, n_txt = m.groups()
        n = int(n_txt)
        try:
            d = date(int(yyyy), int(mm), int(dd))
        except Exception:
            continue
        seg = f"{dd}-{mm}-{yyyy}-{n}-OK"
        prev = out.get(n)
        if (prev is None) or (d > prev[0]):
            out[n] = (d, seg)
    return out

def _fmt_hoy_dd_mm_yyyy_utc() -> str:
    return datetime.now(timezone.utc).strftime("%d-%m-%Y")

def _dias_hasta_venc(it: dict, campo_venc: str, hoy: date) -> Optional[int]:
    d = _parse_fecha(it.get(campo_venc))
    return None if d is None else (d - hoy).days

def _sellar_notificacion_en_item(
    it: dict,
    hoy: date,
    *,
    campo_vencimiento: str = CAMPO_VENCIMIENTO,
    campo_notif: str = CAMPO_NOTIFICACIONES,
    contar_negativos_como_n1: bool = False
) -> bool:
    diff = _dias_hasta_venc(it, campo_vencimiento, hoy)
    if diff is None:
        return False

    if diff < 0:
        if not contar_negativos_como_n1:
            return False
        n = 1
    else:
        n = _calc_n_por_dias(diff)
        if n is None:
            return False

    existente = _parse_notif_field(it.get(campo_notif, ""))
    hoy_str = _fmt_hoy_dd_mm_yyyy_utc()
    existente[n] = (hoy, f"{hoy_str}-{n}-OK")

    nuevo_valor = "Notif: " + ", ".join(existente[k][1] for k in (3, 2, 1) if k in existente)
    it[campo_notif] = nuevo_valor
    return True

def _sellar_en_lote(
    items: List[dict],
    hoy: date,
    *,
    campo_vencimiento: str = CAMPO_VENCIMIENTO,
    campo_notif: str = CAMPO_NOTIFICACIONES,
    contar_negativos_como_n1: bool = False
) -> Tuple[int, int]:
    upd = 0
    omit = 0
    for it in items or []:
        if not isinstance(it, dict):
            omit += 1
            continue
        ok = _sellar_notificacion_en_item(
            it, hoy,
            campo_vencimiento=campo_vencimiento,
            campo_notif=campo_notif,
            contar_negativos_como_n1=contar_negativos_como_n1
        )
        upd += 1 if ok else 0
        omit += 0 if ok else 1
    return upd, omit

# --------------------- Función principal ---------------------



def NotificarVencimientosExcel(
    CONFIG: Dict[str, Any],
    CamposEmail: Union[str, List[str], Tuple[str, ...]],
    ListaNotificarLicenciasConEmail: Any,
    ListaNotificarLicenciasSinEmail: Any,
    ListaNotificarLincenciasComercialInactivos: Any,
    ListaLicenciasFechaNoValida: Any,
    Destinatarios: Dict[str, Any]
) -> Tuple[bool, str, Dict[str, Any]]:
    """
    Envía correos (por comercial y consolidados), sella el campo "Notificaciones" y
    retorna las 3 listas con los ítems (mutados) para su persistencia posterior.

    Returns:
      (success: bool, message: str, payload: dict)
      payload = {
        "ListaNotificarLicenciasConEmail": [...],
        "ListaNotificarLicenciasSinEmail": [...],
        "ListaNotificarLincenciasComercialInactivos": [...],
        "Totales": { "enviados_ok": int, "fallidos": int, "sellos_comercial": int, "sellos_consolidado": int }
      }
    """

    # 1) CONFIG + flags
    TENANT_ID = str(CONFIG.get("TENANT_ID", "")).strip()
    CLIENT_ID = str(CONFIG.get("CLIENT_ID", "")).strip()
    CLIENT_SECRET = str(CONFIG.get("CLIENT_SECRET", "")).strip()
    SUBJECT_PREFIX = str(CONFIG.get("SUBJECT_PREFIX", "")).strip()
    SENDER = str(CONFIG.get("SENDER", ""))

    SELLAR_CONSOLIDADO = bool(CONFIG.get("SELLAR_CONSOLIDADO", True))
    CONTAR_VENCIDOS_COMO_N1 = bool(CONFIG.get("CONTAR_VENCIDOS_COMO_N1", False))

    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        return False, "CONFIG incompleto: faltan TENANT_ID, CLIENT_ID o CLIENT_SECRET.", {}

    # 2) Token Graph
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
        return False, f"Error al solicitar token a Azure AD: {ex}", {}
    if token_resp.status_code != 200:
        try:
            err = token_resp.json()
        except Exception:
            err = token_resp.text
        return False, f"Error al obtener token ({token_resp.status_code}): {err}", {}
    access_token = token_resp.json().get("access_token")
    if not access_token:
        return False, "No se recibió access_token en la respuesta de Azure AD.", {}

    # 3) Normalizar entradas (listas mutables)
    ListaNotificarLicenciasConEmail = _ensure_list_of_dicts(ListaNotificarLicenciasConEmail, "ConEmail")
    ListaNotificarLicenciasSinEmail = _ensure_list_of_dicts(ListaNotificarLicenciasSinEmail, "SinEmail")
    ListaNotificarLincenciasComercialInactivos = _ensure_list_of_dicts(ListaNotificarLincenciasComercialInactivos, "Inactivos")
    ListaLicenciasFechaNoValida = _ensure_list_of_dicts(ListaLicenciasFechaNoValida, "FechaNoValida")

    '''html_fechas_invalidas = ""
    if ListaLicenciasFechaNoValida:
        html_fechas_invalidas = _section_html_simple(
            "Licencias con fecha de vencimiento ausente o no válida",
            ListaLicenciasFechaNoValida
        )'''

    _log(f"[INFO] ConEmail={len(ListaNotificarLicenciasConEmail)} | SinEmail={len(ListaNotificarLicenciasSinEmail)} | Inactivos={len(ListaNotificarLincenciasComercialInactivos)}")

    # 4) Destinatarios (vienen de variables de entorno, ver function_app.py)
    list_destinatarios = [x for x in (Destinatarios.get("Destinatarios") or []) if isinstance(x, str)]
    list_seguimiento = [x for x in (Destinatarios.get("Seguimiento") or []) if isinstance(x, str)]
    list_sin_destinatarios = [x for x in (Destinatarios.get("SinDestinatarios") or []) if isinstance(x, str)]

    to_seg = _dedup_and_wrap(list_seguimiento)
    if not to_seg:
        return False, "No hay destinatarios en 'Seguimiento' (o no válidos). No se puede monitorear el proceso.", {}

    hoy_txt = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    hoy_date = datetime.now(timezone.utc).date()

    fallidos: List[str] = []
    enviados_ok = 0
    total_sellos_comercial = 0
    total_sellos_consolidado = 0

    # --- caso sin datos ---
    if not (ListaNotificarLicenciasConEmail or ListaNotificarLicenciasSinEmail or ListaNotificarLincenciasComercialInactivos):
        html_info = _html_body_info_simple(
            hoy_txt,
            "Seguimiento • Sin licencias para notificar",
            "El proceso se ejecutó correctamente y no encontró licencias que notificar en esta ejecución."
        )
        ok, err = _send_graph_email(access_token, SENDER,
                                    f"Seguimiento • Sin licencias para notificar • {hoy_txt}",
                                    html_info, to_seg)
        if not ok:
            fallidos.append(f"Seguimiento(SinNada): {err}")
        else:
            enviados_ok += 1
        # Retornar las 3 listas (vacías) igualmente
        payload = {
            "ListaNotificarLicenciasConEmail": ListaNotificarLicenciasConEmail,
            "ListaNotificarLicenciasSinEmail": ListaNotificarLicenciasSinEmail,
            "ListaNotificarLincenciasComercialInactivos": ListaNotificarLincenciasComercialInactivos,
            "Totales": {
                "enviados_ok": enviados_ok,
                "fallidos": len(fallidos),
                "sellos_comercial": total_sellos_comercial,
                "sellos_consolidado": total_sellos_consolidado
            }
        }
        return (False, "; ".join(fallidos)) if fallidos else (True, "Sin licencias para notificar"), payload

    # --- clasificación auxiliar para consolidado ---
    clasif_sin = _clasificar_por_ventana(ListaNotificarLicenciasSinEmail, hoy_date) if ListaNotificarLicenciasSinEmail else {"v30": [], "v60": [], "v90": []}
    clasif_inv = _clasificar_por_ventana(ListaNotificarLincenciasComercialInactivos, hoy_date) if ListaNotificarLincenciasComercialInactivos else {"v30": [], "v60": [], "v90": []}

    # --- envíos por comercial y sellado ---
    acumulado_con_email: List[Dict[str, Any]] = []
    if ListaNotificarLicenciasConEmail:
        grupos: Dict[str, List[Dict[str, Any]]] = {}
        for it in ListaNotificarLicenciasConEmail:
            if not isinstance(it, dict):
                continue
            em = _extract_email_from_item(it, CamposEmail)
            if not em:
                continue
            grupos.setdefault(em.strip().lower(), []).append(it)

        _log(f"[INFO] Comerciales válidos detectados: {len(grupos)}")

        for correo_lower, items in grupos.items():
            clasif_g = _clasificar_por_ventana(items, hoy_date)
            if not (clasif_g["v30"] or clasif_g["v60"] or clasif_g["v90"]):
                continue

            html_g = _html_body_multi(
                secciones=[{
                    "titulo": "Licencias/Servicios próximos a vencer (asignadas)",
                    "nota_html": None,
                    "v30": clasif_g["v30"],
                    "v60": clasif_g["v60"],
                    "v90": clasif_g["v90"]
                }],
                saludo="Equipo Comercial",
                titulo_principal="Próximos Vencimientos (Licencias, productos y/o servicios)",
                intro_html="Se han identificado licencias/servicios próximos a vencer. Detalle por proximidad."
            )

            subject_core = f"Próximos Vencimientos (Licencias, productos y/o servicios) • {hoy_txt}"
            subject = f"{SUBJECT_PREFIX} {subject_core}".strip() if SUBJECT_PREFIX else subject_core

            to_list = _dedup_and_wrap([correo_lower] + list_destinatarios)
            if not to_list:
                fallidos.append(f"{correo_lower or '(sin correo válido)'} [sin destinatarios TO válidos]")
                continue

            ok, err = _send_graph_email(access_token, SENDER, subject, html_g, to_list)
            if not ok:
                fallidos.append(f"{correo_lower}: {err}")
            else:
                enviados_ok += 1
                acumulado_con_email.extend(items)
                upd, omit = _sellar_en_lote(
                    items, hoy_date,
                    campo_vencimiento=CAMPO_VENCIMIENTO,
                    campo_notif=CAMPO_NOTIFICACIONES,
                    contar_negativos_como_n1=CONTAR_VENCIDOS_COMO_N1
                )
                total_sellos_comercial += upd
                _log(f"[STAMP Comerciales] {correo_lower}: actualizados={upd}, omitidos={omit}")

    # --- consolidado SinDestinatarios + (opcional) sellado ---
    secciones_sin_inv = []
    if ListaNotificarLicenciasSinEmail:
        secciones_sin_inv.append({
            "titulo": "Vencimientos sin comercial asignado o no válido",
            "nota_html": "No tienen destinatario (email) o no hay comercial asignado.",
            "v30": clasif_sin["v30"],
            "v60": clasif_sin["v60"],
            "v90": clasif_sin["v90"]
        })
    if ListaNotificarLincenciasComercialInactivos:
        secciones_sin_inv.append({
            "titulo": "Vencimientos con EmailComercial no válido/inactivo",
            "nota_html": "Destinatarios detectados inactivos o no válidos.",
            "v30": clasif_inv["v30"],
            "v60": clasif_inv["v60"],
            "v90": clasif_inv["v90"]
        })
    if ListaLicenciasFechaNoValida:
        secciones_sin_inv.append({
            "titulo": "Licencias con fecha de vencimiento ausente o no válida",
            "nota_html": (
                "Estas licencias no tienen una fecha válida y requieren corrección "
                "antes de poder ser notificadas."
            ),
            "v30": ListaLicenciasFechaNoValida,
            "v60": [],
            "v90": []
        })


    if secciones_sin_inv:
        to_sin = _dedup_and_wrap(list_sin_destinatarios)
        if to_sin:
            html_sin_inv = _html_body_multi(
                secciones=secciones_sin_inv,
                saludo="Equipo Comercial (Asignación/Actualización requerida)",
                titulo_principal="Vencimientos sin destinatario o con email no válido/inactivo",
                intro_html="Se consolidan vencimientos que requieren asignación y/o actualización de contacto."
            )
            '''if html_fechas_invalidas:
                html_sin_inv = html_sin_inv.replace(
                    "</body>",
                    f"{html_fechas_invalidas}</body>"
                )'''
            subject_sin = f"Vencimientos sin destinatario o con email no válido • {hoy_txt}"
            subject_sin = f"{SUBJECT_PREFIX} {subject_sin}".strip() if SUBJECT_PREFIX else subject_sin

            ok, err = _send_graph_email(access_token, SENDER, subject_sin, html_sin_inv, to_sin)
            if not ok:
                fallidos.append(f"SinDestinatarios(Consolidado): {err}")
            else:
                enviados_ok += 1
                if SELLAR_CONSOLIDADO:
                    todos_consolidados: List[Dict[str, Any]] = []
                    for sec in secciones_sin_inv:
                        todos_consolidados.extend(sec["v30"])
                        todos_consolidados.extend(sec["v60"])
                        todos_consolidados.extend(sec["v90"])
                    upd, omit = _sellar_en_lote(
                        todos_consolidados, hoy_date,
                        campo_vencimiento=CAMPO_VENCIMIENTO,
                        campo_notif=CAMPO_NOTIFICACIONES,
                        contar_negativos_como_n1=CONTAR_VENCIDOS_COMO_N1
                    )
                    total_sellos_consolidado += upd
                    _log(f"[STAMP Consolidados] actualizados={upd}, omitidos={omit}")
        else:
            _log("[WARN] No hay correos válidos en 'SinDestinatarios' para enviar consolidado.")

    # --- Seguimiento consolidado ---
    secciones_seg = []
    if acumulado_con_email:
        clasif_all = _clasificar_por_ventana(acumulado_con_email, hoy_date)
        secciones_seg.append({
            "titulo": "Envíos a comerciales válidos",
            "nota_html": "Se enviaron notificaciones a los comerciales con email válido.",
            "v30": clasif_all["v30"],
            "v60": clasif_all["v60"],
            "v90": clasif_all["v90"]
        })
    if ListaNotificarLicenciasSinEmail:
        secciones_seg.append({
            "titulo": "Sin comercial asignado o no válido",
            "nota_html": "Pendiente asignación de comercial/destinatario.",
            "v30": clasif_sin["v30"],
            "v60": clasif_sin["v60"],
            "v90": clasif_sin["v90"]
        })
    if ListaNotificarLincenciasComercialInactivos:
        secciones_seg.append({
            "titulo": "Comerciales no válidos/inactivos",
            "nota_html": "Pendiente actualización de información de contacto.",
            "v30": clasif_inv["v30"],
            "v60": clasif_inv["v60"],
            "v90": clasif_inv["v90"]
        })
    if ListaLicenciasFechaNoValida:
        secciones_seg.append({
            "titulo": "Licencias con fecha de vencimiento ausente o no válida",
            "nota_html": (
                "Detectadas durante la validación. No fueron notificadas ni selladas."
            ),
            "v30": ListaLicenciasFechaNoValida,
            "v60": [],
            "v90": []
        })


    if secciones_seg:
        html_seg = _html_body_multi(
            secciones=secciones_seg,
            saludo="Seguimiento",
            titulo_principal="Seguimiento • Resumen consolidado de notificaciones",
            intro_html="Resumen consolidado: envíos a comerciales válidos, sin destinatario y no válidos/inactivos."
        )
        '''if html_fechas_invalidas:
            html_seg = html_seg.replace(
                "</body>",
                f"{html_fechas_invalidas}</body>"
            )'''
        subject_seg = f"Seguimiento • Resumen consolidado • {hoy_txt}"
        subject_seg = f"{SUBJECT_PREFIX} {subject_seg}".strip() if SUBJECT_PREFIX else subject_seg
        ok, err = _send_graph_email(access_token, SENDER, subject_seg, html_seg, to_seg)
        if not ok:
            fallidos.append(f"Seguimiento(Consolidado): {err}")
        else:
            enviados_ok += 1
    else:
        html_info = _html_body_info_simple(
            hoy_txt,
            "Seguimiento • Sin licencias para notificar",
            "El proceso se ejecutó correctamente y no encontró licencias que notificar en esta ejecución."
        )
        ok, err = _send_graph_email(access_token, SENDER,
                                    f"Seguimiento • Sin licencias para notificar • {hoy_txt}",
                                    html_info, to_seg)
        if not ok:
            fallidos.append(f"Seguimiento(SinNada2): {err}")
        else:
            enviados_ok += 1

    # --- Payload final: devolver 3 listas (mutadas/selladas) ---
    payload = {
        "ListaNotificarLicenciasConEmail": ListaNotificarLicenciasConEmail,
        "ListaNotificarLicenciasSinEmail": ListaNotificarLicenciasSinEmail,
        "ListaNotificarLincenciasComercialInactivos": ListaNotificarLincenciasComercialInactivos,
        "ListaLicenciasFechaNoValida": ListaLicenciasFechaNoValida,
        "Totales": {
            "enviados_ok": enviados_ok,
            "fallidos": len(fallidos),
            "sellos_comercial": total_sellos_comercial,
            "sellos_consolidado": total_sellos_consolidado
        }
    }

    if fallidos:
        return False, f"Correos OK: {enviados_ok}. Fallidos: {len(fallidos)}. Detalle: " + "; ".join(fallidos), payload
    return True, f"Correos enviados OK: {enviados_ok}", payload
    