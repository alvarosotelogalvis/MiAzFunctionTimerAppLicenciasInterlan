# -*- coding: utf-8 -*-
"""
Filtra (descarta) licencias de una lista aplanada según la intención de notificación/alerta.

Reglas por origen (token principal = texto antes de la primera coma, normalizado sin acentos y en minúsculas):
- Excel  -> descarta si token == 'no'
- Lista1 -> descarta si token == 'no'
- Lista2 -> descarta si token == 'nook' o 'no'   (ej.: "NoOK, Notif: ...")

Entrada:
- lista_aplanada (list[dict])
- mapping CAMPOSDEALERTAONOTIFICACION y parámetro 'origen' para saber en qué campo leer.

Salida:
- success (bool), message (str), data (dict con dos listas):
    {
        "LicenciasDescartadas": [...],
        "LicenciasNoDescartadas": [...]
    }
"""

from typing import Any, Dict, List, Tuple

__all__ = ["DescartarLicenciasDeExcelPorNotificacionNo"]

def DescartarLicenciasDeExcelPorNotificacionNo(
    lista_aplanada: List[Dict[str, Any]],
    campos_alerta_o_notificacion: Dict[str, str],
    origen: str = "Excel"
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:
    """
    Descarta elementos de una lista aplanada cuya intención de notificación sea negativa
    según reglas por origen, leyendo el campo indicado por `campos_alerta_o_notificacion[origen]`.

    Regla de descarte (por origen):
      - Se toma el valor del campo indicado (si existe y es str).
      - Se obtiene el token principal (antes de la primera coma), se normaliza (minúsculas, sin acentos).
      - Si el token pertenece al conjunto de 'negativos' del origen -> DESCARTAR.
        Negativos por defecto:
          Excel  : {'no'}
          Lista1 : {'no'}
          Lista2 : {'nook', 'no'}

    Retorna:
      (success, message, {
          "LicenciasDescartadas": [...],
          "LicenciasNoDescartadas": [...]
      })
    """
    # Fallback de normalizador si no existe _key_norm global
    try:
        _key_norm  # type: ignore  # noqa: F821
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

    def _normalize_text_basic(s: str) -> str:
        """Minúsculas + quita acentos + recorta espacios (conserva comas)."""
        import unicodedata
        s2 = s.strip()
        s2 = "".join(ch for ch in unicodedata.normalize("NFKD", s2) if not unicodedata.combining(ch))
        return s2.lower()

    # Conjuntos de tokens 'negativos' por origen (editables)
    NEGATIVOS_POR_ORIGEN: Dict[str, set] = {
        "Excel": {"no"},
        "Lista1": {"no"},
        # Para Lista2, si 'Alerta' viene como "NoOK, ...", se normaliza a "nook"
        "Lista2": {"nook", "no"},
    }

    try:
        if not isinstance(lista_aplanada, list):
            return False, "Error de sistema: 'lista_aplanada' no es una lista.", {
                "LicenciasDescartadas": [],
                "LicenciasNoDescartadas": []
            }
        if not isinstance(campos_alerta_o_notificacion, dict):
            return False, "Error de sistema: 'campos_alerta_o_notificacion' no es un dict.", {
                "LicenciasDescartadas": [],
                "LicenciasNoDescartadas": []
            }

        campo = campos_alerta_o_notificacion.get(origen)
        if not isinstance(campo, str) or not campo.strip():
            return False, f"Error de sistema: no se encontró un campo válido para el origen '{origen}'.", {
                "LicenciasDescartadas": [],
                "LicenciasNoDescartadas": []
            }

        negativos = NEGATIVOS_POR_ORIGEN.get(origen, {"no"})  # por defecto, 'no'

        total_entrada = len(lista_aplanada)
        total_conservadas = 0
        total_descartadas = 0

        lista_descartadas: List[Dict[str, Any]] = []
        lista_conservadas: List[Dict[str, Any]] = []

        for fila in lista_aplanada:
            if not isinstance(fila, dict):
                # Ignora filas no dict (no cuenta ni como descartada ni como conservada)
                continue  # 👈 importante: saltar la iteración

            valor = fila.get(campo, None)

            # Si no es str, se conserva
            if not isinstance(valor, str):
                lista_conservadas.append(dict(fila))
                total_conservadas += 1
                continue

            # Token principal: texto antes de la primera coma
            token_principal = valor.split(",", 1)[0]
            token_norm = _normalize_text_basic(token_principal)

            # 'sí' -> 'si' ya lo cubre _normalize_text_basic (quita acentos)
            # 'NoOK' -> 'nook' por lower+sin acentos
            if token_norm in negativos:
                lista_descartadas.append(dict(fila))
                total_descartadas += 1
                continue

            lista_conservadas.append(dict(fila))
            total_conservadas += 1

        mensaje = (
            f"Filtrado aplicado sobre '{campo}' (origen='{origen}'). "
            f"Entradas: {total_entrada}. Conservadas: {total_conservadas}. Descartadas: {total_descartadas}."
        )
        data = {
            "LicenciasDescartadas": lista_descartadas,
            "LicenciasNoDescartadas": lista_conservadas
        }
        return True, mensaje, data

    except Exception as ex:
        return False, f"Error al descartar por notificación: {ex}", {
            "LicenciasDescartadas": [],
            "LicenciasNoDescartadas": []
        }