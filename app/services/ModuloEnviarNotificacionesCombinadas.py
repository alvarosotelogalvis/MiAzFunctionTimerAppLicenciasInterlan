# -*- coding: utf-8 -*-
"""
EnviarNotificacionesCombinadas
-------------------------------
Fusiona los "planes" de clasificación calculados por PrepararNotificacionesExcel
(ModuloNotificacionVencimientosExcel.py) y PrepararNotificacionesLista
(ModuloNotificacionVencimientos2.py), y envía UNA SOLA VEZ por destinatario los
3 correos que antes cada pipeline mandaba por separado (duplicados):

1. "Próximos Vencimientos" — uno por comercial, con secciones tituladas por
   origen (Excel / Lista de Gestión) cuando ese comercial tiene datos de ambas.
2. "Vencimientos sin destinatario o con email no válido" — un solo consolidado.
3. "Seguimiento • Resumen consolidado" — un solo consolidado.

El sellado del campo "Notificaciones" (SellarNotificacionesExcel /
SellarNotificacionesLista) solo se aplica sobre los ítems de la fuente que
efectivamente logró enviarse — igual que antes, "enviar y sellar" siguen siendo
atómicos, solo que ahora ocurren aquí en vez de dentro de cada pipeline.

CONFIG opcional:
  - SELLAR_CONSOLIDADO: bool = True
  - CONTAR_VENCIDOS_COMO_N1: bool = False
  - SUBJECT_PREFIX: str
  - SENDER: str  (usuario/casilla que envía por Graph)

Requiere en CONFIG: TENANT_ID, CLIENT_ID, CLIENT_SECRET (App Graph con permisos .default)
"""

from typing import Any, Dict, List, Optional, Tuple
from datetime import datetime, timezone, date

from app.services.ModuloGraphMailClient import ObtenerTokenGraph, DedupYEnvolver, EnviarCorreoGraph
from app.services.ModuloNotificacionVencimientosExcel import (
    ConstruirFragmentoSeccionesExcel,
    SellarNotificacionesExcel,
)
from app.services.ModuloNotificacionVencimientos2 import (
    ConstruirFragmentoSeccionesLista,
    SellarNotificacionesLista,
)

CAMPO_VENCIMIENTO_LISTA_DEFAULT = "Fecha_x0020_Fin"
CAMPO_NOTIFICACIONES_LISTA_DEFAULT = "Notificaciones"


def _log(msg: str) -> None:
    try:
        print(msg)
    except Exception:
        pass


def _envolver_html(fragmento_html: str, saludo: str, titulo_principal: str, intro_html: Optional[str]) -> str:
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
<title>{titulo_principal}</title>
<style>
  body {{ font-family: Arial, Helvetica, sans-serif; color: #333333; line-height: 1.35; }}
  h1, h2 {{ color: #1a73e8; font-weight: 600; margin: 18px 0 6px 0; }}
  h3 {{ color: #444444; margin: 14px 0 4px 0; font-weight: 600; }}
  p {{ margin: 6px 0; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 12px; }}
  th {{ background-color: #1a73e8; color: #ffffff; padding: 8px; text-align: left; font-size: 13px; border: 1px solid #d0d7de; }}
  td {{ padding: 8px; border-bottom: 1px solid #dddddd; border-left: 1px solid #f0f0f0; border-right: 1px solid #f0f0f0; font-size: 13px; vertical-align: top; }}
  .small {{ color: #666; font-size: 12px; }}
  .origen {{ margin-top: 22px; padding-top: 6px; border-top: 2px solid #1a73e8; }}
</style>
</head>
<body>

<p>Estimado(a) <strong>{saludo}</strong>,</p>

{intro_block}

{fragmento_html}

<p class="small">
Este mensaje es informativo y forma parte del proceso de aseguramiento de continuidad operativa y seguridad.
</p>

<p>Atentamente,<br>
<strong>Equipo comercial</strong><br>
<strong>Interlan</strong><br>
</p>

</body>
</html>"""


def _html_sin_novedades(fecha_txt: str, titulo: str, mensaje: str) -> str:
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>{titulo}</title>
<style>
  body {{ font-family: Arial, Helvetica, sans-serif; color: #333333; line-height: 1.35; }}
  h2 {{ color: #1a73e8; font-weight: 600; }}
  p {{ margin: 6px 0; }}
  .small {{ color: #666; font-size: 12px; }}
</style>
</head>
<body>
  <h2>{titulo}</h2>
  <p>Fecha de ejecución: <strong>{fecha_txt}</strong></p>
  <p>{mensaje}</p>
  <p class="small">Interlan – Calidad e Innovación</p>
</body>
</html>"""


def _combinar_fragmentos_por_origen(
    secciones_excel: List[Dict[str, Any]],
    secciones_lista: List[Dict[str, Any]],
    campo_venc_lista: str
) -> str:
    bloques = []
    if secciones_excel:
        bloques.append(f'<div class="origen"><h2>Origen: Excel</h2>{ConstruirFragmentoSeccionesExcel(secciones_excel)}</div>')
    if secciones_lista:
        bloques.append(f'<div class="origen"><h2>Origen: Lista de Gestión</h2>{ConstruirFragmentoSeccionesLista(secciones_lista, campo_venc_lista)}</div>')
    return "".join(bloques)


def _flatten(secciones: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for sec in secciones or []:
        out.extend(sec.get("v30", []))
        out.extend(sec.get("v60", []))
        out.extend(sec.get("v90", []))
    return out


def EnviarNotificacionesCombinadas(
    CONFIG: Dict[str, Any],
    Destinatarios: Dict[str, Any],
    PlanExcel: Optional[Dict[str, Any]],
    PlanLista: Optional[Dict[str, Any]],
) -> Tuple[bool, str, Dict[str, Any]]:
    """
    Returns:
      (success, message, resultado)
      resultado = {
        "enviados_ok": int,
        "fallidos": int,
        "detalle_fallidos": [str, ...],
        "sellos": {"excel_comercial": int, "lista_comercial": int,
                   "excel_consolidado": int, "lista_consolidado": int}
      }
    """
    SUBJECT_PREFIX = str(CONFIG.get("SUBJECT_PREFIX", "")).strip()
    SENDER = str(CONFIG.get("SENDER", ""))
    SELLAR_CONSOLIDADO = bool(CONFIG.get("SELLAR_CONSOLIDADO", True))
    CONTAR_VENCIDOS_COMO_N1 = bool(CONFIG.get("CONTAR_VENCIDOS_COMO_N1", False))

    ok_token, access_token_o_error = ObtenerTokenGraph(CONFIG)
    if not ok_token:
        return False, access_token_o_error, {}
    access_token = access_token_o_error

    plan_excel = PlanExcel or {}
    plan_lista = PlanLista or {}

    campo_venc_lista = plan_lista.get("campo_vencimiento", CAMPO_VENCIMIENTO_LISTA_DEFAULT)
    campo_notif_lista = plan_lista.get("campo_notificaciones", CAMPO_NOTIFICACIONES_LISTA_DEFAULT)

    list_destinatarios = [x for x in (Destinatarios.get("Destinatarios") or []) if isinstance(x, str)]
    list_seguimiento = [x for x in (Destinatarios.get("Seguimiento") or []) if isinstance(x, str)]
    list_sin_destinatarios = [x for x in (Destinatarios.get("SinDestinatarios") or []) if isinstance(x, str)]

    to_seg = DedupYEnvolver(list_seguimiento)
    if not to_seg:
        return False, "No hay destinatarios en 'Seguimiento' (o no válidos). No se puede monitorear el proceso.", {}

    hoy_txt = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    hoy_date = datetime.now(timezone.utc).date()

    fallidos: List[str] = []
    enviados_ok = 0
    sellos_excel_comercial = 0
    sellos_lista_comercial = 0
    sellos_excel_consolidado = 0
    sellos_lista_consolidado = 0

    por_comercial_excel: Dict[str, Any] = plan_excel.get("por_comercial", {}) or {}
    por_comercial_lista: Dict[str, Any] = plan_lista.get("por_comercial", {}) or {}
    secciones_sin_destinatario_excel = plan_excel.get("secciones_sin_destinatario", []) or []
    secciones_sin_destinatario_lista = plan_lista.get("secciones_sin_destinatario", []) or []
    secciones_seguimiento_excel = plan_excel.get("secciones_seguimiento_base", []) or []
    secciones_seguimiento_lista = plan_lista.get("secciones_seguimiento_base", []) or []

    hay_algo_que_reportar = bool(
        por_comercial_excel or por_comercial_lista
        or secciones_sin_destinatario_excel or secciones_sin_destinatario_lista
        or secciones_seguimiento_excel or secciones_seguimiento_lista
    )

    # --- caso sin datos en ninguna de las 2 pipelines ---
    if not hay_algo_que_reportar:
        html_info = _html_sin_novedades(
            hoy_txt,
            "Seguimiento • Sin licencias para notificar",
            "El proceso se ejecutó correctamente y no encontró licencias que notificar en esta ejecución."
        )
        ok, err = EnviarCorreoGraph(access_token, SENDER,
                                     f"Seguimiento • Sin licencias para notificar • {hoy_txt}",
                                     html_info, to_seg)
        if not ok:
            fallidos.append(f"Seguimiento(SinNada): {err}")
        else:
            enviados_ok += 1
        resultado = {
            "enviados_ok": enviados_ok,
            "fallidos": len(fallidos),
            "detalle_fallidos": fallidos,
            "sellos": {
                "excel_comercial": 0, "lista_comercial": 0,
                "excel_consolidado": 0, "lista_consolidado": 0,
            },
        }
        return (not fallidos), ("Sin licencias para notificar" if not fallidos else "; ".join(fallidos)), resultado

    # --- 1) Próximos Vencimientos: un correo por comercial, combinando ambas fuentes ---
    todos_los_correos = set(por_comercial_excel.keys()) | set(por_comercial_lista.keys())
    for correo_lower in todos_los_correos:
        entry_excel = por_comercial_excel.get(correo_lower)
        entry_lista = por_comercial_lista.get(correo_lower)

        secciones_excel = []
        if entry_excel:
            secciones_excel = [{
                "titulo": "Licencias/Servicios próximos a vencer (asignadas)",
                "nota_html": None,
                **entry_excel["clasif"],
            }]
        secciones_lista = []
        if entry_lista:
            secciones_lista = [{
                "titulo": "Licencias/Servicios próximos a vencer (asignadas)",
                "nota_html": None,
                **entry_lista["clasif"],
            }]

        fragmento = _combinar_fragmentos_por_origen(secciones_excel, secciones_lista, campo_venc_lista)
        html_g = _envolver_html(
            fragmento,
            saludo="Equipo Comercial",
            titulo_principal="Próximos Vencimientos (Licencias, productos y/o servicios)",
            intro_html="Se han identificado licencias/servicios próximos a vencer. Detalle por proximidad."
        )

        subject_core = f"Próximos Vencimientos (Licencias, productos y/o servicios) • {hoy_txt}"
        subject = f"{SUBJECT_PREFIX} {subject_core}".strip() if SUBJECT_PREFIX else subject_core

        to_list = DedupYEnvolver([correo_lower] + list_destinatarios)
        if not to_list:
            fallidos.append(f"{correo_lower or '(sin correo válido)'} [sin destinatarios TO válidos]")
            continue

        ok, err = EnviarCorreoGraph(access_token, SENDER, subject, html_g, to_list)
        if not ok:
            fallidos.append(f"{correo_lower}: {err}")
            continue

        enviados_ok += 1
        if entry_excel:
            upd, _omit = SellarNotificacionesExcel(entry_excel["items"], hoy_date, CONTAR_VENCIDOS_COMO_N1)
            sellos_excel_comercial += upd
        if entry_lista:
            upd, _omit = SellarNotificacionesLista(
                entry_lista["items"], hoy_date, campo_venc_lista, campo_notif_lista, CONTAR_VENCIDOS_COMO_N1
            )
            sellos_lista_comercial += upd

    # --- 2) Vencimientos sin destinatario o con email no válido: un solo consolidado ---
    if secciones_sin_destinatario_excel or secciones_sin_destinatario_lista:
        to_sin = DedupYEnvolver(list_sin_destinatarios)
        if to_sin:
            fragmento_sin = _combinar_fragmentos_por_origen(
                secciones_sin_destinatario_excel, secciones_sin_destinatario_lista, campo_venc_lista
            )
            html_sin = _envolver_html(
                fragmento_sin,
                saludo="Equipo Comercial (Asignación/Actualización requerida)",
                titulo_principal="Vencimientos sin destinatario o con email no válido/inactivo",
                intro_html="Se consolidan vencimientos que requieren asignación y/o actualización de contacto."
            )
            subject_sin = f"Vencimientos sin destinatario o con email no válido • {hoy_txt}"
            subject_sin = f"{SUBJECT_PREFIX} {subject_sin}".strip() if SUBJECT_PREFIX else subject_sin

            ok, err = EnviarCorreoGraph(access_token, SENDER, subject_sin, html_sin, to_sin)
            if not ok:
                fallidos.append(f"SinDestinatarios(Consolidado): {err}")
            else:
                enviados_ok += 1
                if SELLAR_CONSOLIDADO:
                    if secciones_sin_destinatario_excel:
                        upd, _omit = SellarNotificacionesExcel(
                            _flatten(secciones_sin_destinatario_excel), hoy_date, CONTAR_VENCIDOS_COMO_N1
                        )
                        sellos_excel_consolidado += upd
                    if secciones_sin_destinatario_lista:
                        upd, _omit = SellarNotificacionesLista(
                            _flatten(secciones_sin_destinatario_lista), hoy_date,
                            campo_venc_lista, campo_notif_lista, CONTAR_VENCIDOS_COMO_N1
                        )
                        sellos_lista_consolidado += upd
        else:
            _log("[WARN] No hay correos válidos en 'SinDestinatarios' para enviar consolidado.")

    # --- 3) Seguimiento • Resumen consolidado: un solo consolidado (no se sella) ---
    if secciones_seguimiento_excel or secciones_seguimiento_lista:
        fragmento_seg = _combinar_fragmentos_por_origen(
            secciones_seguimiento_excel, secciones_seguimiento_lista, campo_venc_lista
        )
        html_seg = _envolver_html(
            fragmento_seg,
            saludo="Seguimiento",
            titulo_principal="Seguimiento • Resumen consolidado de notificaciones",
            intro_html="Resumen consolidado: envíos a comerciales válidos, sin destinatario y no válidos/inactivos."
        )
        subject_seg = f"Seguimiento • Resumen consolidado • {hoy_txt}"
        subject_seg = f"{SUBJECT_PREFIX} {subject_seg}".strip() if SUBJECT_PREFIX else subject_seg
        ok, err = EnviarCorreoGraph(access_token, SENDER, subject_seg, html_seg, to_seg)
        if not ok:
            fallidos.append(f"Seguimiento(Consolidado): {err}")
        else:
            enviados_ok += 1

    resultado = {
        "enviados_ok": enviados_ok,
        "fallidos": len(fallidos),
        "detalle_fallidos": fallidos,
        "sellos": {
            "excel_comercial": sellos_excel_comercial,
            "lista_comercial": sellos_lista_comercial,
            "excel_consolidado": sellos_excel_consolidado,
            "lista_consolidado": sellos_lista_consolidado,
        },
    }

    if fallidos:
        return False, f"Correos OK: {enviados_ok}. Fallidos: {len(fallidos)}. Detalle: " + "; ".join(fallidos), resultado
    return True, f"Correos enviados OK: {enviados_ok}", resultado
