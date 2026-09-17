# -*- coding: utf-8 -*-
"""
PrepararNotificacionesExcel (clasificación en ventanas + agrupación por comercial)
-----------------------------------------------------------------------------------------

Clasifica (sin enviar correo ni sellar "Notificaciones") las licencias del pipeline
de Excel en ventanas de vencimiento v30/v60/v90, agrupadas por comercial y por
categoría "sin destinatario". El envío real por Microsoft Graph y el sellado del
campo "Notificaciones" ahora los hace EnviarNotificacionesCombinadas
(ModuloEnviarNotificacionesCombinadas.py), una sola vez, después de que tanto esta
pipeline como la de Listas hayan terminado de clasificar — así cada destinatario
recibe un solo correo combinando ambas fuentes en vez de uno por pipeline.

Sella el campo "Notificaciones" con 'dd-mm-yyyy-N-OK' (N=3,2,1) según días a vencer:
    N=1:  0 <= días < 30
    N=2: 30 <= días < 60
    N=3: 60 <= días <= 90
(la función de sellado, SellarNotificacionesExcel, la invoca el módulo fusionador).
"""

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

def ConstruirFragmentoSeccionesExcel(secciones: List[Dict[str, Any]]) -> str:
    """
    Construye el fragmento HTML (sin el envoltorio <html><head>) de una lista de
    secciones {titulo, nota_html, v30, v60, v90}. Lo usa el módulo fusionador para
    insertar las tablas de licencias de Excel dentro de un correo combinado junto
    con las de la pipeline de Listas.
    """
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
    return "".join(bloques)

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

def SellarNotificacionesExcel(
    items: List[Dict[str, Any]],
    hoy: date,
    contar_negativos_como_n1: bool = False
) -> Tuple[int, int]:
    """Wrapper público sobre _sellar_en_lote, para que el módulo fusionador no
    tenga que tocar internals con guion bajo de este módulo."""
    return _sellar_en_lote(
        items, hoy,
        campo_vencimiento=CAMPO_VENCIMIENTO,
        campo_notif=CAMPO_NOTIFICACIONES,
        contar_negativos_como_n1=contar_negativos_como_n1
    )

# --------------------- Función principal (clasificación, sin enviar) ---------------------

def PrepararNotificacionesExcel(
    CamposEmail: Union[str, List[str], Tuple[str, ...]],
    ListaNotificarLicenciasConEmail: Any,
    ListaNotificarLicenciasSinEmail: Any,
    ListaNotificarLincenciasComercialInactivos: Any,
    ListaLicenciasFechaNoValida: Any,
) -> Tuple[bool, str, Dict[str, Any]]:
    """
    Clasifica las licencias del pipeline de Excel en ventanas de vencimiento
    (v30/v60/v90), agrupadas por comercial y por categoría "sin destinatario".
    No envía correo ni sella "Notificaciones" — eso lo hace, una sola vez para
    ambas pipelines, EnviarNotificacionesCombinadas.

    Returns:
      (success, message, plan)
      plan = {
        "por_comercial": {email_lower: {"items": [...], "clasif": {"v30":[],"v60":[],"v90":[]}}},
        "secciones_sin_destinatario": [ {titulo, nota_html, v30, v60, v90}, ... ],
        "secciones_seguimiento_base": [ {titulo, nota_html, v30, v60, v90}, ... ],
        "raw_lists": {
          "ConEmail": [...], "SinEmail": [...], "Inactivos": [...], "FechaNoValida": [...]
        }
      }
    """
    ListaNotificarLicenciasConEmail = _ensure_list_of_dicts(ListaNotificarLicenciasConEmail, "ConEmail")
    ListaNotificarLicenciasSinEmail = _ensure_list_of_dicts(ListaNotificarLicenciasSinEmail, "SinEmail")
    ListaNotificarLincenciasComercialInactivos = _ensure_list_of_dicts(ListaNotificarLincenciasComercialInactivos, "Inactivos")
    ListaLicenciasFechaNoValida = _ensure_list_of_dicts(ListaLicenciasFechaNoValida, "FechaNoValida")

    raw_lists = {
        "ConEmail": ListaNotificarLicenciasConEmail,
        "SinEmail": ListaNotificarLicenciasSinEmail,
        "Inactivos": ListaNotificarLincenciasComercialInactivos,
        "FechaNoValida": ListaLicenciasFechaNoValida,
    }

    _log(f"[INFO] ConEmail={len(ListaNotificarLicenciasConEmail)} | SinEmail={len(ListaNotificarLicenciasSinEmail)} | Inactivos={len(ListaNotificarLincenciasComercialInactivos)}")

    if not (ListaNotificarLicenciasConEmail or ListaNotificarLicenciasSinEmail
            or ListaNotificarLincenciasComercialInactivos or ListaLicenciasFechaNoValida):
        plan = {
            "por_comercial": {},
            "secciones_sin_destinatario": [],
            "secciones_seguimiento_base": [],
            "raw_lists": raw_lists,
        }
        return True, "Sin licencias para notificar", plan

    hoy_date = datetime.now(timezone.utc).date()

    clasif_sin = _clasificar_por_ventana(ListaNotificarLicenciasSinEmail, hoy_date) if ListaNotificarLicenciasSinEmail else {"v30": [], "v60": [], "v90": []}
    clasif_inv = _clasificar_por_ventana(ListaNotificarLincenciasComercialInactivos, hoy_date) if ListaNotificarLincenciasComercialInactivos else {"v30": [], "v60": [], "v90": []}
    clasif_con_email = _clasificar_por_ventana(ListaNotificarLicenciasConEmail, hoy_date) if ListaNotificarLicenciasConEmail else {"v30": [], "v60": [], "v90": []}

    # --- agrupar por comercial ---
    por_comercial: Dict[str, Dict[str, Any]] = {}
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
            por_comercial[correo_lower] = {"items": items, "clasif": clasif_g}

    # --- secciones para el correo consolidado "sin destinatario" ---
    secciones_sin_destinatario = []
    if ListaNotificarLicenciasSinEmail:
        secciones_sin_destinatario.append({
            "titulo": "Vencimientos sin comercial asignado o no válido",
            "nota_html": "No tienen destinatario (email) o no hay comercial asignado.",
            "v30": clasif_sin["v30"], "v60": clasif_sin["v60"], "v90": clasif_sin["v90"],
        })
    if ListaNotificarLincenciasComercialInactivos:
        secciones_sin_destinatario.append({
            "titulo": "Vencimientos con EmailComercial no válido/inactivo",
            "nota_html": "Destinatarios detectados inactivos o no válidos.",
            "v30": clasif_inv["v30"], "v60": clasif_inv["v60"], "v90": clasif_inv["v90"],
        })
    if ListaLicenciasFechaNoValida:
        secciones_sin_destinatario.append({
            "titulo": "Licencias con fecha de vencimiento ausente o no válida",
            "nota_html": (
                "Estas licencias no tienen una fecha válida y requieren corrección "
                "antes de poder ser notificadas."
            ),
            "v30": ListaLicenciasFechaNoValida, "v60": [], "v90": [],
        })

    # --- secciones para el correo "Seguimiento • Resumen consolidado" ---
    secciones_seguimiento_base = []
    if ListaNotificarLicenciasConEmail:
        secciones_seguimiento_base.append({
            "titulo": "Envíos a comerciales válidos",
            "nota_html": "Licencias asignadas a comerciales con email válido.",
            "v30": clasif_con_email["v30"], "v60": clasif_con_email["v60"], "v90": clasif_con_email["v90"],
        })
    if ListaNotificarLicenciasSinEmail:
        secciones_seguimiento_base.append({
            "titulo": "Sin comercial asignado o no válido",
            "nota_html": "Pendiente asignación de comercial/destinatario.",
            "v30": clasif_sin["v30"], "v60": clasif_sin["v60"], "v90": clasif_sin["v90"],
        })
    if ListaNotificarLincenciasComercialInactivos:
        secciones_seguimiento_base.append({
            "titulo": "Comerciales no válidos/inactivos",
            "nota_html": "Pendiente actualización de información de contacto.",
            "v30": clasif_inv["v30"], "v60": clasif_inv["v60"], "v90": clasif_inv["v90"],
        })
    if ListaLicenciasFechaNoValida:
        secciones_seguimiento_base.append({
            "titulo": "Licencias con fecha de vencimiento ausente o no válida",
            "nota_html": "Detectadas durante la validación. No fueron notificadas ni selladas.",
            "v30": ListaLicenciasFechaNoValida, "v60": [], "v90": [],
        })

    plan = {
        "por_comercial": por_comercial,
        "secciones_sin_destinatario": secciones_sin_destinatario,
        "secciones_seguimiento_base": secciones_seguimiento_base,
        "raw_lists": raw_lists,
    }
    return True, f"OK: {len(por_comercial)} comercial(es), {len(secciones_sin_destinatario)} sección(es) sin destinatario.", plan
