# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Tuple, Optional
from datetime import datetime, date, timezone, timedelta
import re

__all__ = ["ValidarVencimientosLicenciasExcel"]

def ValidarVencimientosLicenciasExcel(
    lista_aplanada: List[Dict[str, Any]],
    *,
    campo_vencimiento: Optional[str] = None,
    ventana_dias: int = 90
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:

    # ---------- Helpers ----------
    def _parse_fecha_flexible(fecha: Any) -> Optional[datetime]:
        if fecha is None:
            return None

        if isinstance(fecha, datetime):
            if fecha.tzinfo is None:
                fecha = fecha.replace(tzinfo=timezone.utc)
            return fecha.astimezone(timezone.utc)

        if isinstance(fecha, date):
            return datetime(fecha.year, fecha.month, fecha.day, tzinfo=timezone.utc)

        if not isinstance(fecha, str):
            return None

        f = fecha.strip()
        if not f:
            return None

        # ISO 8601
        try:
            if f.endswith("Z"):
                f = f.replace("Z", "+00:00")
            dt = datetime.fromisoformat(f)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except Exception:
            pass

        # Formatos comunes
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(f, fmt).replace(tzinfo=timezone.utc)
            except Exception:
                pass

        # Buscar fecha dentro de texto
        m = re.search(r"(?<!\d)(\d{2})[/-](\d{2})[/-](\d{4})(?!\d)", f)
        if m:
            dd, mm, yyyy = m.groups()
            return datetime(int(yyyy), int(mm), int(dd), tzinfo=timezone.utc)

        return None

    def _resolver_campo_venc(r: Dict[str, Any]) -> Optional[str]:
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

    # ---------- Validación base ----------
    if not isinstance(lista_aplanada, list):
        return False, "Error de sistema: 'lista_aplanada' no es una lista.", {
            "LicenciasParaNotificar": [],
            "LicenciasNoNotificar": [],
            "LicenciasSinFecha": [],
            "LicenciasFechaInvalida": [],
        }

    try:
        ventana = max(0, int(ventana_dias))
    except Exception:
        ventana = 0

    hoy = datetime.now(timezone.utc).date()
    fin = hoy + timedelta(days=ventana)

    # ---------- Acumuladores ----------
    lic_para_notificar = []
    lic_no_notificar = []
    lic_sin_fecha = []
    lic_fecha_invalida = []

    total = 0

    for row in lista_aplanada:
        if not isinstance(row, dict):
            continue

        total += 1
        out = dict(row)

        campo_venc = _resolver_campo_venc(out)
        if not campo_venc or not out.get(campo_venc):
            lic_sin_fecha.append(out)
            continue

        dt_venc = _parse_fecha_flexible(out.get(campo_venc))
        if dt_venc is None:
            lic_fecha_invalida.append(out)
            continue

        d_venc = dt_venc.date()

        if d_venc < hoy:
            lic_no_notificar.append(out)
        elif hoy <= d_venc <= fin:
            lic_para_notificar.append(out)
        else:
            lic_no_notificar.append(out)

    mensaje = (
        f"Validación completada. "
        f"Entradas: {total}. "
        f"ParaNotificar: {len(lic_para_notificar)}. "
        f"NoNotificar: {len(lic_no_notificar)}. "
        f"SinFecha: {len(lic_sin_fecha)}. "
        f"FechaInvalida: {len(lic_fecha_invalida)}."
    )

    payload = {
        "LicenciasParaNotificar": lic_para_notificar,
        "LicenciasNoNotificar": lic_no_notificar,
        "LicenciasSinFecha": lic_sin_fecha,
        "LicenciasFechaInvalida": lic_fecha_invalida,
    }

    return True, mensaje, payload
    