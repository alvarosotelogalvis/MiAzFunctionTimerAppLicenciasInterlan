# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Tuple, Optional
from datetime import datetime, date, timezone, timedelta
import re

__all__ = ["ValidarVencimientosLicenciasExcel"]

def ValidarVencimientosLicenciasExcel(
    lista_aplanada: List[Dict[str, Any]],
    *,
    campo_vencimiento: Optional[str] = None,   # Si None, se infiere por nombres comunes
    ventana_dias: int = 90                     # Próximos N días (incluye hoy)
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:
    """
    Clasifica licencias según su fecha de vencimiento.

    Reglas:
      - Si la fecha de vencimiento está entre hoy y hoy + ventana_dias (inclusive)
        -> "LicenciasParaNotificar"
      - En caso contrario (sin fecha, inválida, ya vencida o > ventana)
        -> "LicenciasNoNotificar"

    No modifica la lista original; retorna copias defensivas.

    Retorna:
      (success: bool, message: str, payload: Dict[str, List[Dict[str, Any]]])
        payload = {
          "LicenciasNoNotificar": [...],
          "LicenciasParaNotificar": [...]
        }
    """

    # -------- Helpers --------
    def _parse_fecha_flexible(fecha: Any) -> Optional[datetime]:
        """
        Acepta:
          - datetime/date: los normaliza a datetime UTC.
          - str: ISO 'YYYY-MM-DD', 'YYYY-MM-DDTHH:MM:SS[Z|+offset]', 'DD/MM/YYYY', 'DD-MM-YYYY'.
          - Fallback: busca dd[/-]mm[/-]yyyy dentro de un texto.
        Devuelve datetime en UTC o None.
        """
        if fecha is None:
            return None

        # Ya es datetime o date
        if isinstance(fecha, datetime):
            dt = fecha
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        if isinstance(fecha, date):
            return datetime(fecha.year, fecha.month, fecha.day, tzinfo=timezone.utc)

        # Cadenas
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
            dt = datetime.strptime(f, "%Y-%m-%d").replace(tzinfo=timezone.utc)
            return dt
        except Exception:
            pass

        # DD/MM/YYYY
        try:
            dt = datetime.strptime(f, "%d/%m/%Y").replace(tzinfo=timezone.utc)
            return dt
        except Exception:
            pass

        # DD-MM-YYYY
        try:
            dt = datetime.strptime(f, "%d-%m-%Y").replace(tzinfo=timezone.utc)
            return dt
        except Exception:
            pass

        # Fallback: buscar dd[/|-]mm[/|-]yyyy dentro de un texto más largo
        try:
            m = re.search(r"(?<!\d)(\d{2})[/-](\d{2})[/-](\d{4})(?!\d)", f)
            if m:
                dd, mm, yyyy = m.groups()
                dt = datetime(int(yyyy), int(mm), int(dd), tzinfo=timezone.utc)
                return dt
        except Exception:
            pass

        return None

    def _resolver_campo_venc(r: Dict[str, Any]) -> Optional[str]:
        """Si no se especifica campo_vencimiento, intenta resolverlo por nombres comunes en la fila."""
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

    # -------- Validaciones base --------
    if not isinstance(lista_aplanada, list):
        return False, "Error de sistema: 'lista_aplanada' no es una lista.", {
            "LicenciasNoNotificar": [],
            "LicenciasParaNotificar": []
        }

    # Ventana no negativa
    try:
        ventana = max(0, int(ventana_dias))
    except Exception:
        ventana = 0

    # -------- Proceso --------
    lic_no_notificar: List[Dict[str, Any]] = []
    lic_para_notificar: List[Dict[str, Any]] = []

    hoy = datetime.now(timezone.utc).date()
    fin = hoy + timedelta(days=ventana)

    total = 0
    m_sin_venc = 0
    m_invalidas = 0
    m_ya_vencidas = 0
    m_fuera_ventana = 0
    m_para_notificar = 0

    for row in lista_aplanada:
        if not isinstance(row, dict):
            continue
        total += 1
        out = dict(row)  # copia defensiva

        campo_venc_resuelto = _resolver_campo_venc(out)
        if not campo_venc_resuelto:
            lic_no_notificar.append(out)
            m_sin_venc += 1
            continue

        dt_venc = _parse_fecha_flexible(out.get(campo_venc_resuelto, ""))
        if dt_venc is None:
            lic_no_notificar.append(out)
            m_invalidas += 1
            continue

        d_venc = dt_venc.date()
        if d_venc < hoy:
            # Ya vencida -> fuera del ciclo actual
            lic_no_notificar.append(out)
            m_ya_vencidas += 1
            continue

        # ¿Está entre hoy y fin (inclusive)?
        if hoy <= d_venc <= fin:
            lic_para_notificar.append(out)
            m_para_notificar += 1
        else:
            lic_no_notificar.append(out)
            m_fuera_ventana += 1

    msg = (
        f"Validación de vencimientos (próximos {ventana} días) completada. "
        f"Entradas: {total}. "
        f"ParaNotificar: {m_para_notificar}. "
        f"Ya vencidas: {m_ya_vencidas}. "
        f"Fuera de ventana: {m_fuera_ventana}. "
        f"Sin campo/fecha de vencimiento: {m_sin_venc}. "
        f"Fechas inválidas: {m_invalidas}."
    )

    payload: Dict[str, List[Dict[str, Any]]] = {
        "LicenciasNoNotificar": lic_no_notificar,
        "LicenciasParaNotificar": lic_para_notificar,
    }
    return True, msg, payload
    