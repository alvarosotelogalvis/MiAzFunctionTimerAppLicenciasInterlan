# -*- coding: utf-8 -*-
"""
Correos de seguimiento/alerta de la ejecución de la Azure Function.

- EnviarCorreoSeguimientoLogsCombinado: UN solo correo de "Seguimiento • Ejecución
  Azure Function" con el log de ambas pipelines (Excel + Listas), en vez de uno
  por pipeline como antes.
- EnviarCorreoDeErrorPipeline: correo de alerta real cuando un paso falla (antes no
  existía ningún aviso funcional — ver el checkpoint en function_app.py).
"""

from typing import Dict, Any, Tuple, List
from datetime import datetime, timezone
import html

from app.services.ModuloGraphMailClient import ObtenerTokenGraph, DedupYEnvolver, EnviarCorreoGraph


def _esc(value: Any) -> str:
    """Escapa texto para HTML"""
    return html.escape(str(value)) if value is not None else ""


def _construir_tabla_eventos(msglogs: Dict[str, Any]) -> str:
    inicio = msglogs.get("Inicio", "")
    eventos = msglogs.get("Eventos", [])

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

    return f"""
<p><strong>Inicio:</strong> {_esc(inicio)}</p>
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
"""


_ESTILO_COMUN = """
body { font-family: Arial, Helvetica, sans-serif; color: #333; }
h2 { color: #1a73e8; }
h3 { color: #444444; margin: 18px 0 4px 0; }
table { width: 100%; border-collapse: collapse; margin-top: 12px; }
th { background-color: #1a73e8; color: #fff; padding: 8px; text-align: left; }
td { padding: 8px; border-bottom: 1px solid #ddd; font-size: 13px; vertical-align: top; }
.small { color: #666; font-size: 12px; }
.banner-error {
    background: #fdecea; border: 1px solid #d93025; color: #a50e0e;
    border-radius: 6px; padding: 12px 14px; margin-bottom: 16px;
}
.banner-warning {
    background: #fff8e1; border: 1px solid #f1c40f; color: #7d6608;
    border-radius: 6px; padding: 12px 14px; margin-bottom: 16px;
}
"""


def EnviarCorreoSeguimientoLogsCombinado(
    CONFIG: Dict[str, Any],
    MSGLOGS: Dict[str, Any],
    MSGLOGS2: Dict[str, Any],
    Destinatarios: Dict[str, Any]
) -> Tuple[bool, str]:
    """
    Envía UN solo correo HTML con el seguimiento completo de la ejecución de la
    Azure Function, combinando el log de la pipeline de Excel (MSGLOGS) y el de
    la pipeline de Listas (MSGLOGS2).

    Returns:
        (success: bool, message: str)
    """
    SUBJECT_PREFIX = str(CONFIG.get("SUBJECT_PREFIX", "")).strip()
    SENDER = str(CONFIG.get("SENDER", "")).strip()

    ok_token, access_token_o_error = ObtenerTokenGraph(CONFIG)
    if not ok_token:
        return False, access_token_o_error
    access_token = access_token_o_error

    to_recipients = DedupYEnvolver(Destinatarios.get("Seguimiento") or [])
    if not to_recipients:
        return False, "No hay destinatarios válidos en 'Seguimiento'."

    fecha_ejecucion = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    html_body = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Seguimiento ejecución Azure Function</title>
<style>{_ESTILO_COMUN}</style>
</head>
<body>

<h2>Seguimiento ejecución Azure Function</h2>
<p><strong>Fecha envío:</strong> {_esc(fecha_ejecucion)}</p>

<h3>Pipeline Excel</h3>
{_construir_tabla_eventos(MSGLOGS)}

<h3>Pipeline Listas</h3>
{_construir_tabla_eventos(MSGLOGS2)}

<p class="small">
Este correo corresponde al seguimiento automático de la ejecución de la Azure Function
<b>AppLicenciasInterlan</b>.
</p>

</body>
</html>
"""

    hoy_txt = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    subject_core = f"Seguimiento • Ejecución Azure Function • {hoy_txt}"
    subject = f"{SUBJECT_PREFIX} {subject_core}".strip() if SUBJECT_PREFIX else subject_core

    ok, err = EnviarCorreoGraph(access_token, SENDER, subject, html_body, to_recipients)
    if not ok:
        return False, f"Error enviando correo de seguimiento: {err}"
    return True, "Correo de seguimiento enviado correctamente."


def EnviarCorreoDeErrorPipeline(
    CONFIG: Dict[str, Any],
    MSGLOGS_list: List[Dict[str, Any]],
    Destinatarios: Dict[str, Any],
    PipelineLabel: str,
    Paso: str,
    MensajeError: str,
    Severidad: str = "Falla",
) -> Tuple[bool, str]:
    """
    Envía un correo de alerta cuando un paso de la Azure Function falla o se degrada,
    indicando qué paso fue, su mensaje, y todo lo que se alcanzó a loguear en esa(s)
    pipeline(s) hasta ese momento.

    MSGLOGS_list: lista de dicts MSGLOGS (uno para una sola pipeline, o varios cuando
    el evento ocurre en el tramo final compartido y se quiere mostrar el log de ambas).

    Severidad: "Falla" (rojo, el proceso se detuvo en este punto — comportamiento
    original) o "Advertencia" (naranja, algo no salió del todo bien pero el proceso
    continuó con datos parciales/anteriores, ej. Fortinet no disponible).

    Returns:
        (success: bool, message: str)
    """
    SUBJECT_PREFIX = str(CONFIG.get("SUBJECT_PREFIX", "")).strip()
    SENDER = str(CONFIG.get("SENDER", "")).strip()
    es_advertencia = Severidad.strip().lower() == "advertencia"

    ok_token, access_token_o_error = ObtenerTokenGraph(CONFIG)
    if not ok_token:
        return False, access_token_o_error
    access_token = access_token_o_error

    recipients_raw = Destinatarios.get("AlertasFallo") or Destinatarios.get("Seguimiento") or []
    to_recipients = DedupYEnvolver(recipients_raw)
    if not to_recipients:
        return False, "No hay destinatarios válidos para la alerta ('AlertasFallo' ni 'Seguimiento')."

    fecha_ejecucion = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    tablas_html = "".join(
        f"<h3>Log {('Pipeline ' + str(i + 1)) if len(MSGLOGS_list) > 1 else ''}</h3>{_construir_tabla_eventos(m)}"
        for i, m in enumerate(MSGLOGS_list or [])
    )

    emoji = "⚠"
    titulo = f"{emoji} Advertencia en la ejecución de la Azure Function" if es_advertencia else f"{emoji} Falla en la ejecución de la Azure Function"
    banner_clase = "banner-warning" if es_advertencia else "banner-error"
    etiqueta_paso = "Paso con advertencia:" if es_advertencia else "Paso que falló:"
    nota_final = (
        "El proceso continuó con los datos que ya había disponibles para este paso. "
        "Este correo se generó automáticamente."
        if es_advertencia else
        "El proceso se detuvo en este punto — los pasos posteriores de esta corrida no se "
        "ejecutaron. Este correo se generó automáticamente."
    )

    html_body = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>{titulo}</title>
<style>{_ESTILO_COMUN}</style>
</head>
<body>

<h2>{titulo}</h2>

<div class="{banner_clase}">
  <p><strong>Pipeline:</strong> {_esc(PipelineLabel)}</p>
  <p><strong>{etiqueta_paso}</strong> {_esc(Paso)}</p>
  <p><strong>Mensaje:</strong> {_esc(MensajeError)}</p>
  <p><strong>Fecha:</strong> {_esc(fecha_ejecucion)}</p>
</div>

{tablas_html}

<p class="small">
{nota_final}
</p>

</body>
</html>
"""

    hoy_txt = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    etiqueta_asunto = "Advertencia" if es_advertencia else "Falla"
    subject_core = f"{emoji} {etiqueta_asunto} • Azure Function ({PipelineLabel}) • {Paso} • {hoy_txt}"
    subject = f"{SUBJECT_PREFIX} {subject_core}".strip() if SUBJECT_PREFIX else subject_core

    ok, err = EnviarCorreoGraph(access_token, SENDER, subject, html_body, to_recipients)
    if not ok:
        return False, f"Error enviando correo de alerta: {err}"
    return True, "Correo de alerta enviado correctamente."
