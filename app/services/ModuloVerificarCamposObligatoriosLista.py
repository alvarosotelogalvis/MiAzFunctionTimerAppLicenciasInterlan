from typing import List, Dict, Any, Tuple

def VerificarCamposObligatoriosLista(
    lista_aplanada: List[Dict[str, Any]],
    campos_obligatorios: List[str]
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """
    Verifica/normaliza que cada dict (fila) de `lista_aplanada` contenga los `campos_obligatorios`.
    - Si falta un campo, lo **agrega** con valor por defecto "" (no falla).
    - Acepta alias comunes (p. ej., Comercial -> EmailComercial, Versión <-> Version, Dias de Vencimiento -> Días de Vencimiento).
    - Si encuentra un alias, crea el **campo canónico** con el valor del alias (sin borrar el alias).
    - Solo devuelve error (success=False) en **errores de sistema** (p. ej., tipos no esperados).

    Retorna:
        (success: bool, message: str, lista_aplanada_normalizada: list[dict])
    """
    DEFAULT_VALUE = ""  # Cambia a None si lo prefieres

    # --- Normalizador: usa el que ya tienes; si no existe, define fallback ---
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
        # Validaciones "suaves": si algo viene vacío, seguimos y devolvemos success=True igualmente.
        if not isinstance(lista_aplanada, list):
            return False, "Error de sistema: 'lista_aplanada' no es una lista.", []
        if not isinstance(campos_obligatorios, list):
            return False, "Error de sistema: 'campos_obligatorios' no es una lista.", lista_aplanada

        # Canonicalización de nombres requeridos
        requeridos = [(campo, _key_norm(campo)) for campo in campos_obligatorios]
        canon_norm_to_canon_text = {norm_c: canon for (canon, norm_c) in requeridos}

        # Aliases: normalizado(alias) -> normalizado(canónico)
        alias_norm_to_canon_norm = {
            _key_norm("Comercial"): _key_norm("EmailComercial"),
            _key_norm("Email"): _key_norm("EmailComercial"),
            _key_norm("NombreComercial"): _key_norm("EmailComercial"),

            # Variaciones de "Versión"
            _key_norm("Version"): _key_norm("Versión"),
            _key_norm("Versión"): _key_norm("Versión"),

            # Variaciones de "Días de Vencimiento"
            _key_norm("Dias de Vencimiento"): _key_norm("Días de Vencimiento"),
            _key_norm("Días de Vencimiento"): _key_norm("Días de Vencimiento"),

            # Fecha Vencimiento
            _key_norm("Fecha Vencimiento (DD/MM/AAAA)"): _key_norm("Fecha Vencimiento (DD/MM/AAAA)"),
        }

        total_filas = len(lista_aplanada)
        total_campos_agregados = 0
        total_campos_normalizados_por_alias = 0

        for fila in lista_aplanada:
            if not isinstance(fila, dict):
                # Si una fila no es dict, no podemos normalizar: error de sistema.
                return False, "Error de sistema: una fila no es dict.", lista_aplanada

            # Mapa (normalizado -> nombre real tal como aparece en la fila)
            norm_to_real = {}
            for k in list(fila.keys()):
                norm_to_real[_key_norm(k)] = k

            # 1) Cubrir faltantes con alias si es posible; si no, rellenar por defecto.
            for canon_text, canon_norm in requeridos:
                if canon_norm in norm_to_real:
                    # Ya existe el canónico (tal vez con variación de mayúsculas/acentos)
                    # Aseguramos que la clave canónica exacta exista (ajuste cosmético)
                    real_key = norm_to_real[canon_norm]
                    if canon_text not in fila:
                        fila[canon_text] = fila.get(real_key, DEFAULT_VALUE)
                        total_campos_agregados += 1
                    continue

                # Buscar alias para este canónico
                alias_encontrado = None
                for alias_norm, target_norm in alias_norm_to_canon_norm.items():
                    if target_norm == canon_norm and alias_norm in norm_to_real:
                        alias_encontrado = alias_norm
                        break

                if alias_encontrado is not None:
                    # Creamos el canónico con el valor del alias
                    real_alias_key = norm_to_real[alias_encontrado]
                    if canon_text not in fila:
                        fila[canon_text] = fila.get(real_alias_key, DEFAULT_VALUE)
                        total_campos_normalizados_por_alias += 1
                    # No removemos el alias, solo garantizamos el canónico
                else:
                    # No existe ni canónico ni alias: crear campo con valor por defecto
                    if canon_text not in fila:
                        fila[canon_text] = DEFAULT_VALUE
                        total_campos_agregados += 1

            # 2) (Opcional) Asegurar consistencia de tipos/casos particulares.
            #    Aquí podrías, por ejemplo, asegurar que "Cantidad" sea int si es posible, etc.
            #    No tocamos más para no alterar tu pipeline sin necesidad.

        msg_ok = (
            f"OK: {total_filas} registro(s) verificados. "
            f"Campos agregados por faltantes: {total_campos_agregados}. "
            f"Campos canónicos creados desde alias: {total_campos_normalizados_por_alias}."
        )
        return True, msg_ok, lista_aplanada

    except Exception as ex:
        # Único caso en que devolvemos error: fallo del sistema (excepción)
        return False, f"Error de sistema en verificación de campos obligatorios: {ex}", lista_aplanada
