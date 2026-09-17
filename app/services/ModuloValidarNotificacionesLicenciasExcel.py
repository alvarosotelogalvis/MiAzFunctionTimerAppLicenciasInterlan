# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Tuple, Optional
from datetime import datetime, date, timezone
import re

__all__ = ["ValidarNotificacionesLicenciasExcel"]

def ValidarNotificacionesLicenciasExcel(
    lista_aplanada: List[Dict[str, Any]],
    campo_alerta_o_notificacion: str,
    *,
    campo_vencimiento: Optional[str] = None,
    ventana_dias: int = 120
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:
    """
    Valida el ciclo de notificaciones almacenado en `campo_alerta_o_notificacion` contra
    la fecha de vencimiento (`campo_vencimiento`), limpiando entradas inválidas y
    clasificando las licencias en:
      - LicenciasNotificacionesValidasCompletas: tienen 3 notificaciones válidas (N=3,2,1)
      - LicenciasNotificacionesIncompletas: resto (se resetea a 'Notif:')

    Formato por entrada: 'dd-mm-yyyy-N-OK' (también se acepta dd/mm/yyyy), con N en {1,2,3}.
    Ventanas sin solapes (diff = (FechaVenc - FechaNotif).days):
      - N=1 válido si  0 <= diff < 30
      - N=2 válido si 30 <= diff < 60
      - N=3 válido si 60 <= diff <= max90, donde max90 = min(ventana_dias, 90)

    No modifica las filas originales; trabaja con copias defensivas.

    Retorna:
      (success: bool, message: str, payload: Dict[str, List[Dict[str, Any]]])
        payload = {
          "LicenciasNotificacionesValidasCompletas": [...],
          "LicenciasNotificacionesIncompletas": [...]
        }
    """

    # ---------- Helpers ----------
    def _parse_fecha_flexible(fecha: Any) -> Optional[datetime]:
        """
        Acepta:
          - datetime/date: normaliza a datetime UTC.
          - str: ISO 'YYYY-MM-DD', 'YYYY-MM-DDTHH:MM:SS[Z|+offset]', 'DD/MM/YYYY', 'DD-MM-YYYY'.
        Devuelve datetime en UTC o None.
        """
        if fecha is None:
            return None
        if isinstance(fecha, datetime):
            dt = fecha
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        if isinstance(fecha, date):
            return datetime(fecha.year, fecha.month, fecha.day, tzinfo=timezone.utc)

        if not isinstance(fecha, str):
            return None

        f = fecha.strip()
        if not f:
            return None

        # ISO (con Z/offset o solo fecha)
        try:
            if f.endswith("Z"):
                f_iso = f.replace("Z", "+00:00")
                dt = datetime.fromisoformat(f_iso)
            else:
                dt = datetime.fromisoformat(f)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except Exception:
            pass

        # YYYY-MM-DD
        try:
            return datetime.strptime(f, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        except Exception:
            pass

        # DD/MM/YYYY
        try:
            return datetime.strptime(f, "%d/%m/%Y").replace(tzinfo=timezone.utc)
        except Exception:
            pass

        # DD-MM-YYYY
        try:
            return datetime.strptime(f, "%d-%m-%Y").replace(tzinfo=timezone.utc)
        except Exception:
            pass

        return None

    def _resolver_campo_venc(r: Dict[str, Any]) -> Optional[str]:
        """Si no se especifica `campo_vencimiento`, intenta resolverlo por nombres comunes en la fila."""
        if campo_vencimiento:
            return campo_vencimiento
        for k in (
            "Fecha Vencimiento (DD/MM/AAAA)",
            "FechaVencimiento (DD/MM/AAAA)",
            "FechaVencimiento",
            "Fecha Fin",
            "Fecha_x0020_Fin",
            "FechaFin",
        ):
            if k in r:
                return k
        return None

    def _formatear_dd_mm_yyyy(d: datetime) -> str:
        return d.strftime("%d-%m-%Y")

    # Patrón 'dd[-|/]mm[-|/]yyyy-N-OK' (con espacios permitidos)
    _re_notif = re.compile(
        r"(?<!\d)\s*(\d{2})[/-](\d{2})[/-](\d{4})\s*-\s*([123])\s*-\s*OK\s*(?!\d)",
        flags=re.IGNORECASE
    )

    # Clasificación de ventanas (sin solapes)
    def _n_por_dias(diff: int, max90: int) -> Optional[int]:
        if diff < 0:
            return None
        if 0 <= diff < 30:
            return 1
        if 30 <= diff < 60:
            return 2
        if 60 <= diff <= max90:
            return 3
        return None

    # ---------- Validaciones base ----------
    if not isinstance(lista_aplanada, list):
        return False, "Error: 'lista_aplanada' no es una lista.", {
            "LicenciasNotificacionesValidasCompletas": [],
            "LicenciasNotificacionesIncompletas": []
        }
    if not isinstance(campo_alerta_o_notificacion, str) or not campo_alerta_o_notificacion.strip():
        return False, "Error: 'campo_alerta_o_notificacion' debe ser un string no vacío.", {
            "LicenciasNotificacionesValidasCompletas": [],
            "LicenciasNotificacionesIncompletas": []
        }

    # Tope superior: 90 días para N=3
    try:
        max90 = max(0, min(int(ventana_dias), 90))
    except Exception:
        max90 = 90

    lic_validas_completas: List[Dict[str, Any]] = []
    lic_incompletas: List[Dict[str, Any]] = []

    total = 0
    m_sin_venc = 0
    m_venc_invalida = 0
    m_completas = 0
    m_reseteadas = 0
    m_limpiezas_inval = 0
    m_sin_contenido = 0
    m_no_str = 0

    for row in lista_aplanada:
        if not isinstance(row, dict):
            continue
        total += 1
        out = dict(row)  # copia defensiva

        # Resolver fecha de vencimiento
        campo_venc_resuelto = _resolver_campo_venc(out)
        if not campo_venc_resuelto:
            out[campo_alerta_o_notificacion] = "Notif:"
            lic_incompletas.append(out)
            m_sin_venc += 1
            continue

        dt_venc = _parse_fecha_flexible(out.get(campo_venc_resuelto, ""))
        if dt_venc is None:
            out[campo_alerta_o_notificacion] = "Notif:"
            lic_incompletas.append(out)
            m_venc_invalida += 1
            continue

        d_venc = dt_venc.date()

        raw = out.get(campo_alerta_o_notificacion, None)
        if raw is None or not isinstance(raw, str):
            out[campo_alerta_o_notificacion] = "Notif:"
            lic_incompletas.append(out)
            m_no_str += 1
            continue

        # Obtener la parte después de 'Notif:' (si no aparece, se toma todo el string)
        partes = raw.split("Notif:", 1)
        cola = partes[1] if len(partes) > 1 else raw
        segs = [s.strip() for s in cola.split(",") if s.strip()]

        if not segs:
            out[campo_alerta_o_notificacion] = "Notif:"
            lic_incompletas.append(out)
            m_sin_contenido += 1
            continue

        # Validar y limpiar entradas: máximo 1 por N, conservar la más reciente
        grupo_por_n: Dict[int, Dict[str, Any]] = {}
        limpiezas_locales = 0

        for seg in segs:
            m = _re_notif.search(seg)
            if not m:
                limpiezas_locales += 1
                continue

            dd, mm, yyyy, n_txt = m.groups()
            try:
                dt_notif = datetime(int(yyyy), int(mm), int(dd), tzinfo=timezone.utc)
            except Exception:
                limpiezas_locales += 1
                continue

            n_val = int(n_txt)
            d_notif = dt_notif.date()
            diff = (d_venc - d_notif).days

            n_esperado = _n_por_dias(diff, max90)
            if n_esperado is None or n_esperado != n_val:
                limpiezas_locales += 1
                continue

            # Conservar la más reciente por N
            prev = grupo_por_n.get(n_val)
            if (prev is None) or (d_notif > prev["date"]):
                grupo_por_n[n_val] = {
                    "date": d_notif,
                    "diff": diff,
                    "str": f"{_formatear_dd_mm_yyyy(dt_notif)}-{n_val}-OK",
                }

        m_limpiezas_inval += limpiezas_locales

        # ¿Está completa? Debe tener N={1,2,3}
        if set(grupo_por_n.keys()) == {1, 2, 3}:
            nuevo_valor = "Notif: " + ", ".join(grupo_por_n[n]["str"] for n in (3, 2, 1))
            out[campo_alerta_o_notificacion] = nuevo_valor
            lic_validas_completas.append(out)
            m_completas += 1
        else:
            out[campo_alerta_o_notificacion] = "Notif:"
            lic_incompletas.append(out)
            m_reseteadas += 1

    msg = (
        f"Validación de notificaciones (N1=[0,30), N2=[30,60), N3=[60,{max90}]) completada. "
        f"Entradas: {total}. "
        f"Completas: {m_completas}. "
        f"Incompletas reseteadas: {m_reseteadas}. "
        f"Sin campo/fecha de vencimiento: {m_sin_venc}. "
        f"Vencimiento inválido: {m_venc_invalida}. "
        f"Sin contenido o no-string: {m_sin_contenido + m_no_str}. "
        f"Segmentos inválidos descartados: {m_limpiezas_inval}."
    )

    payload: Dict[str, List[Dict[str, Any]]] = {
        "LicenciasNotificacionesValidasCompletas": lic_validas_completas,
        "LicenciasNotificacionesIncompletas": lic_incompletas,
    }
    return True, msg, payload
