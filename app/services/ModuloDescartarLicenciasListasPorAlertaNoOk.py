# -*- coding: utf-8 -*-
import re
from typing import Any, Dict, List, Tuple

def _normalize_token_before_comma(value: Any) -> str:
    """
    Convierte el valor a str, toma la parte ANTES de la primera coma,
    recorta espacios, lo pasa a minúsculas y elimina separadores (espacios, '-' y '_').
    Ejemplos:
      "NoOK, Notif"    -> "nook"
      "no ok , algo"   -> "nook"
      "no-ok"          -> "nook"
      "no_ok"          -> "nook"
      " NO  "          -> "no"
      "" / None        -> ""
    """
    if value is None:
        return ""
    s = str(value)
    head = s.split(",", 1)[0].strip().lower()
    if not head:
        return ""
    # Elimina separadores comunes para homogeneizar
    head_norm = re.sub(r"[\s\-_]+", "", head)
    return head_norm

def DescartarLicenciasListasPorAlertaNoOk(
    lista_items: List[Dict[str, Any]],
    campo_alerta: str
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """
    Filtra la lista removiendo elementos cuyo valor en 'campo_alerta', tomando el
    token ANTES de la primera coma (normalizado), equivalga a 'NoOk' o 'no'.
    - Sensible a variantes: 'NoOK', 'no ok', 'no-ok', 'no_ok' -> 'nook' (descarta)
    - 'no' a secas (en cualquier mayúsculas/minúsculas/espaciado) -> 'no' (descarta)
    - Si el campo no existe o está vacío: conserva el ítem.

    Returns:
      (True, mensaje_resumen, lista_filtrada)
    """
    if not isinstance(lista_items, list):
        # No romper el flujo
        return True, "Entrada no es una lista; no se aplicó filtro.", []

    total = len(lista_items)
    descartados = 0
    filtrada: List[Dict[str, Any]] = []

    # Tokens que deseamos descartar (ya normalizados por _normalize_token_before_comma)
    tokens_descartar = {
        "nook",  # 'NoOK', 'no ok', 'no-ok', 'no_ok', etc.
        "no"     # palabra 'no' sola (nuevo requisito)
    }

    for it in lista_items:
        if not isinstance(it, dict):
            # Conservar para no perder datos
            filtrada.append(it)
            continue

        raw = it.get(campo_alerta, None)
        token = _normalize_token_before_comma(raw)

        if token in tokens_descartar:
            descartados += 1
            continue  # Lo descartamos
        else:
            filtrada.append(it)

    kept = len(filtrada)
    msg = (f"Filtro por alerta aplicado. Total: {total}. "
           f"Descartados por '{campo_alerta}' en ('NoOk' o 'no'): {descartados}. "
           f"Conservados: {kept}.")

    return True, msg, filtrada