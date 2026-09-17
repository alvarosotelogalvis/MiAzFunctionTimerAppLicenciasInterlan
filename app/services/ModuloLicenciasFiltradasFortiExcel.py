# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Tuple

__all__ = ["FiltrarLicenciasDeExcelFortinet"]

def FiltrarLicenciasDeExcelFortinet(
    lista_aplanada: List[Dict[str, Any]]
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:
    """
    Filtra una lista aplanada de licencias en dos grupos:
      - Si 'Fabricante' == 'Fortinet' (comparación insensible a mayúsculas/minúsculas y acentos)
      - No Fortinet (cualquier otro valor, faltante o no-string)

    Parámetros
    ----------
    lista_aplanada : List[Dict[str, Any]]
        Lista de diccionarios (cada dict = fila aplanada del Excel).

    Retorna
    -------
    (success, message, data)
        success : bool
            True si el proceso fue exitoso; False si ocurrió un error de sistema.
        message : str
            Métricas del proceso o descripción del error.
        data : Dict[str, List[Dict[str, Any]]]
            {
                "LicenciasExcelPorHojasSiFortinet": [ {fila}, {fila}, ... ],
                "LicenciasExcelPorHojasNoFortinet": [ {fila}, {fila}, ... ],
            }

    Notas
    -----
    - No modifica la lista de entrada (copia cada dict conservado con dict()).
    - Filas no-dict se ignoran silenciosamente.
    - La comparación de 'Fabricante' es robusta: recorta espacios, quita acentos y compara en minúsculas.
    """
    # Fallback del normalizador si no existe _key_norm global en tu proyecto
    try:
        _key_norm  # type: ignore
    except NameError:
        import unicodedata
        def _key_norm(s: Any) -> str:
            if s is None:
                return ""
            s = str(s).strip()
            s = unicodedata.normalize("NFKD", s)
            s = "".join(ch for ch in s if not unicodedata.combining(ch))
            s = s.lower()
            s = "".join(ch for ch in s if ch.isalnum())
            return s

    try:
        if not isinstance(lista_aplanada, list):
            return False, "Error de sistema: 'lista_aplanada' no es una lista.", {}

        si_fortinet: List[Dict[str, Any]] = []
        no_fortinet: List[Dict[str, Any]] = []

        total_si = 0
        total_no = 0
        total_entrada = len(lista_aplanada)

        # Normalizamos la palabra "Fortinet" para comparar
        objetivo_norm = _key_norm("Fortinet")

        for fila in lista_aplanada:
            if not isinstance(fila, dict):
                # Ignora silenciosamente elementos no-dict
                continue

            fabricante = fila.get("Fabricante", None)
            # Si fabricante es str, normalizamos; si no, cae a None
            fabricante_norm = _key_norm(fabricante) if isinstance(fabricante, str) else None

            es_fortinet = (fabricante_norm == objetivo_norm)

            if es_fortinet:
                si_fortinet.append(dict(fila))   # copia defensiva
                total_si += 1
            else:
                no_fortinet.append(dict(fila))   # copia defensiva
                total_no += 1

        message = (
            "Filtro por Fabricante == 'Fortinet' aplicado sobre lista aplanada. "
            f"Entradas: {total_entrada}. Fortinet: {total_si}. No Fortinet: {total_no}."
        )

        data: Dict[str, List[Dict[str, Any]]] = {
            "LicenciasExcelPorHojasSiFortinet": si_fortinet,
            "LicenciasExcelPorHojasNoFortinet": no_fortinet,
        }
        return True, message, data

    except Exception as ex:
        return False, f"Error al filtrar por fabricante Fortinet: {ex}", {}
        