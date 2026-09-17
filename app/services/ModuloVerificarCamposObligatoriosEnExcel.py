# -*- coding: utf-8 -*-
"""
Normaliza y verifica campos obligatorios en una lista aplanada proveniente del Excel de SharePoint.
- No devuelve error por campos faltantes; los agrega con valor por defecto "".
- Reconoce alias y crea el campo canónico con el valor del alias (sin borrar el alias).
- Solo falla (success=False) ante errores de sistema (tipos no válidos, etc.).

Firma:
    success, message, lista_aplanada = VerificarCamposObligatoriosLista(lista_aplanada, campos_obligatorios)
"""

from typing import List, Dict, Any, Tuple

__all__ = ["VerificarCamposObligatoriosLista"]

def VerificarCamposObligatoriosEnExcelDeSharepoint(
    lista_aplanada: List[Dict[str, Any]],
    campos_obligatorios: List[str]
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """
    Verifica/normaliza que cada dict (fila) de `lista_aplanada` contenga los `campos_obligatorios`.
    - Si falta un campo, lo agrega con valor por defecto "" (no falla).
    - Acepta alias comunes y crea el campo canónico con su valor.
    - Devuelve success=True salvo errores de sistema.

    Parámetros:
        lista_aplanada: Lista de diccionarios (ya aplanada) con las filas del Excel.
        campos_obligatorios: Nombres canónicos que deben existir en cada fila.

    Retorna:
        (success: bool, message: str, lista_aplanada_normalizada: list[dict])
    """
    DEFAULT_VALUE = ""  # Cambia a None si prefieres

    # Normalizador: usa el global si existe; si no, define fallback local
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
        # Validaciones de sistema (únicos casos que devuelven error)
        if not isinstance(lista_aplanada, list):
            return False, "Error de sistema: 'lista_aplanada' no es una lista.", []
        if not isinstance(campos_obligatorios, list):
            return False, "Error de sistema: 'campos_obligatorios' no es una lista.", lista_aplanada

        # Canonicalización
        requeridos = [(campo, _key_norm(campo)) for campo in campos_obligatorios]

        # Aliases: normalizado(alias) -> normalizado(canónico)
        alias_norm_to_canon_norm = {
            _key_norm("Comercial"): _key_norm("EmailComercial"),
            _key_norm("Email"): _key_norm("EmailComercial"),
            _key_norm("NombreComercial"): _key_norm("EmailComercial"),

            # Versión
            _key_norm("Version"): _key_norm("Versión"),
            _key_norm("Versión"): _key_norm("Versión"),

            # Días de Vencimiento (sin/with acento)
            _key_norm("Dias de Vencimiento"): _key_norm("Días de Vencimiento"),
            _key_norm("Días de Vencimiento"): _key_norm("Días de Vencimiento"),

            # Fecha Vencimiento (tal cual en tu Excel)
            _key_norm("Fecha Vencimiento (DD/MM/AAAA)"): _key_norm("Fecha Vencimiento (DD/MM/AAAA)"),
        }

        total_filas = len(lista_aplanada)
        total_campos_agregados = 0
        total_can_por_alias = 0

        for fila in lista_aplanada:
            if not isinstance(fila, dict):
                return False, "Error de sistema: una fila no es dict.", lista_aplanada

            # Mapa (normalizado -> nombre real presente en la fila)
            norm_to_real = {}
            for k in list(fila.keys()):
                norm_to_real[_key_norm(k)] = k

            # 1) Garantiza existencia de los canónicos (usando alias si aplica)
            for canon_text, canon_norm in requeridos:
                if canon_norm in norm_to_real:
                    # Existe el campo (aunque con variación del nombre). Refuerza nombre canónico exacto.
                    real_key = norm_to_real[canon_norm]
                    if canon_text not in fila:
                        fila[canon_text] = fila.get(real_key, DEFAULT_VALUE)
                        total_campos_agregados += 1
                    continue

                # Busca alias que apunte a este canónico
                alias_encontrado = None
                for alias_norm, target_norm in alias_norm_to_canon_norm.items():
                    if target_norm == canon_norm and alias_norm in norm_to_real:
                        alias_encontrado = alias_norm
                        break

                if alias_encontrado is not None:
                    real_alias_key = norm_to_real[alias_encontrado]
                    if canon_text not in fila:
                        fila[canon_text] = fila.get(real_alias_key, DEFAULT_VALUE)
                        total_can_por_alias += 1
                else:
                    # Ni canónico ni alias: crea el campo con valor por defecto
                    if canon_text not in fila:
                        fila[canon_text] = DEFAULT_VALUE
                        total_campos_agregados += 1

            # 2) (Opcional) Normalizaciones extra (ej. convertir cantidad a int)
            # if isinstance(fila.get("Cantidad"), str):
            #     try:
            #         fila["Cantidad"] = int(fila["Cantidad"])
            #     except Exception:
            #         pass

        msg_ok = (
            f"OK: {total_filas} registro(s) verificados. "
            f"Campos agregados por faltantes: {total_campos_agregados}. "
            f"Campos canónicos creados desde alias: {total_can_por_alias}."
        )
        return True, msg_ok, lista_aplanada

    except Exception as ex:
        # Único caso de error: falla de sistema
        return False, f"Error de sistema en verificación de campos obligatorios: {ex}", lista_aplanada
