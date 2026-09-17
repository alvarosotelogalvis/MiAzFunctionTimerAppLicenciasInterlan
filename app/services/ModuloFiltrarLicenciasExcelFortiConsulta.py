# -*- coding: utf-8 -*-
from datetime import datetime, date, timedelta, timezone
from typing import Any, Dict, List, Tuple, Optional
import re

__all__ = ["FiltrarLicenciasExcelFortinetParaConsultar"]

def FiltrarLicenciasExcelFortinetParaConsultar(
    lista_aplanada_fortinet: Optional[List[Dict[str, Any]]]
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:
    """
    Filtra licencias Fortinet (lista aplanada) en dos conjuntos:
      - ListasLicenciasFortinetNoConsultar:   las que NO se deben consultar
      - ListasLicenciasFortinetParaConsultar: las que SÍ se consultarán

    Reglas:
      1) Si NotificarVencimiento == "no"            -> NoConsultar
      2) Si fecha no es parseable                   -> ParaConsultar
      3) Si fecha < (hoy - 6 meses aproximados)     -> NoConsultar
      4) Si fecha > hoy                             -> NoConsultar
      5) Si Fabricante != "Fortinet"                -> NoConsultar (seguridad)

    Entrada:
      lista_aplanada_fortinet: List[dict] con filas aplanadas (ya filtradas a Fortinet idealmente)

    Retorno:
      (success: bool, message: str, data: dict)
      data = {
        "ListasLicenciasFortinetNoConsultar":   [ {licencia}, ... ],
        "ListasLicenciasFortinetParaConsultar": [ {licencia}, ... ],
      }

    Notas:
      - No modifica la lista original (hace copias con dict()).
      - Las filas no-dict se ignoran silenciosamente.
    """

    # ---- Utilidades de fecha (idénticas/compatibles con tu flujo previo) ----
    def _parse_iso_or_local(s: str) -> Tuple[Optional[date], Optional[str]]:
        # ISO 8601 (manejo 'Z' -> '+00:00')
        try:
            s2 = s[:-1] + "+00:00" if s.endswith("Z") else s
            dt = datetime.fromisoformat(s2)
            return dt.date(), None
        except Exception:
            pass
        # ISO comunes
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S",
                    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.date(), None
            except Exception:
                continue
        # Locales d/m/a y variantes
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
                    "%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.date(), None
            except Exception:
                continue
        return None, f"Formato no reconocido: {s}"

    def _parse_odata_date(s: str) -> Tuple[Optional[date], Optional[str]]:
        try:
            m = re.match(r"^/Date\((\-?\d+)\)/$", s)
            if not m:
                return None, "OData pattern no coincide"
            ms = int(m.group(1))
            dt = datetime.fromtimestamp(ms / 1000.0, tz=timezone.utc)
            return dt.date(), None
        except Exception as ex:
            return None, f"OData no parseable: {ex}"

    def _parse_excel_serial(n: float) -> Tuple[Optional[date], Optional[str]]:
        try:
            if n <= 0:
                return None, "Serial Excel no positivo"
            # Base 1899-12-30 por compatibilidad con Excel
            base = date(1899, 12, 30)
            d = base + timedelta(days=n)
            if d.year < 1970 or d.year > 2100:
                return None, f"Serial Excel fuera de rango: {n} -> {d.isoformat()}"
            return d, None
        except Exception as ex:
            return None, f"Serial Excel no parseable: {ex}"

    def parse_fecha_sp(value, permitir_excel_serial: bool = True) -> Tuple[Optional[date], Optional[str]]:
        # 0) None
        if value is None:
            return None, "FechaVencimiento nula"
        # 1) date/datetime
        if isinstance(value, date) and not isinstance(value, datetime):
            return value, None
        if isinstance(value, datetime):
            return value.date(), None
        # 2) dict típico con 'dateTime'
        if isinstance(value, dict):
            for k in ("dateTime", "DateTime", "value"):
                if k in value and isinstance(value[k], str) and value[k].strip():
                    s = value[k].strip()
                    d, err = _parse_iso_or_local(s)
                    if d:
                        return d, None
                    d2, err2 = _parse_odata_date(s)
                    if d2:
                        return d2, None
                    return None, f"Dict(dateTime) no parseable: {s} ({err or err2})"
            return None, "Dict sin 'dateTime' o equivalente"
        # 3) OData antiguo
        if isinstance(value, str) and value.startswith("/Date("):
            d, err = _parse_odata_date(value)
            if d:
                return d, None
            return None, err
        # 4) str genérico
        if isinstance(value, str):
            s = value.strip()
            if not s:
                return None, "FechaVencimiento vacía"
            d, err = _parse_iso_or_local(s)
            if d:
                return d, None
            return None, err
        # 5) número (posible serial Excel)
        if isinstance(value, (int, float)):
            if permitir_excel_serial:
                d, err = _parse_excel_serial(float(value))
                if d:
                    return d, None
                return None, f"Formato numérico no soportado (Excel serial inválido): {value}"
            else:
                return None, f"Formato numérico no soportado: {value}"
        # 6) otros tipos
        return None, f"Tipo de fecha no soportado: {type(value).__name__}"

    # ---- Config ventana temporal ----
    hoy: date = date.today()
    seis_meses_atras: date = hoy - timedelta(days=183)  # ~6 meses

    # ---- Contenedores de salida ----
    out_no_consultar: List[Dict[str, Any]] = []
    out_para_consultar: List[Dict[str, Any]] = []

    try:
        # Entrada vacía o inválida → éxito con listas vacías
        if not lista_aplanada_fortinet or not isinstance(lista_aplanada_fortinet, list):
            msg = (
                f"Fortinet -> para consultar: 0, descartadas: 0; "
                f"ventana: [{seis_meses_atras.isoformat()} .. {hoy.isoformat()}] (entrada vacía)"
            )
            return True, msg, {
                "ListasLicenciasFortinetNoConsultar": out_no_consultar,
                "ListasLicenciasFortinetParaConsultar": out_para_consultar,
            }

        total_para = 0
        total_no = 0
        total_entrada = 0

        for row in lista_aplanada_fortinet:
            if not isinstance(row, dict):
                # Ignoramos silenciosamente elementos no-dict
                continue
            total_entrada += 1

            # 1) NotificarVencimiento
            noti_val = row.get('NotificarVencimiento', '')
            if isinstance(noti_val, str) and noti_val.strip().lower() == "no":
                out_no_consultar.append(dict(row))
                total_no += 1
                continue

            # 2) Fecha (acepta distintos nombres de campo)
            def _get_fecha_val(r: Dict[str, Any]):
                for k in (
                    "FechaDeVencimientoDDMMAAAA",
                ):
                    if k in r:
                        return r.get(k)
                return None

            fecha_val = _get_fecha_val(row)
            fecha_v, err = parse_fecha_sp(fecha_val)

            if err or not fecha_v:
                # Se consulta a pesar del error de fecha
                out_para_consultar.append(dict(row))
                total_para += 1
                continue

            # 3) Ventana temporal
            if fecha_v < seis_meses_atras:
                out_no_consultar.append(dict(row))
                total_no += 1
                continue

            if fecha_v > hoy:
                out_no_consultar.append(dict(row))
                total_no += 1
                continue

            # 4) Fabricante = Fortinet (por seguridad, aunque debería venir filtrado)
            fab = row.get('Fabricante', '')
            if not (isinstance(fab, str) and fab.strip().lower() == "fortinet"):
                out_no_consultar.append(dict(row))
                total_no += 1
                continue

            # 5) Pasa filtros -> ParaConsultar
            out_para_consultar.append(dict(row))
            total_para += 1

        respuesta = {
            "ListasLicenciasFortinetNoConsultar": out_no_consultar,
            "ListasLicenciasFortinetParaConsultar": out_para_consultar,
        }
        msg = (
            f"Fortinet -> para consultar: {total_para}, descartadas: {total_no}; "
            f"entradas: {total_entrada}; ventana: [{seis_meses_atras.isoformat()} .. {hoy.isoformat()}]"
        )
        return True, msg, respuesta

    except Exception as ex:
        msg = f"Error en FiltrarLicenciasExcelFortinetParaConsultar: {ex}"
        return False, msg, {
            "ListasLicenciasFortinetNoConsultar": [],
            "ListasLicenciasFortinetParaConsultar": [],
        }
        