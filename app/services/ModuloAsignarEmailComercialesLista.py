from typing import Dict, Any, Tuple, List

def AsignarEmailComercialesLista(
    ComercialesEmails: Dict[str, Any],
    CamposEmail: tuple,
    Licencias: List[Dict[str, Any]]
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """
    Reemplaza el campo indicado en CamposEmail[2] (ej: 'Asesor_x0020_ComercialLookupId')
    por el email correspondiente, usando el mapeo de ComercialesEmails (viene de la
    variable de entorno COMERCIALES_EMAILS_JSON, ver function_app.py).

    Si no encuentra el id → deja el campo en blanco "".

    Retorna:
        (success, message, lista_licencias_modificada)
    """

    try:
        # =========================
        # 1) Validaciones básicas
        # =========================
        if not isinstance(CamposEmail, (list, tuple)) or len(CamposEmail) < 3:
            return False, "CamposEmail inválido. Se esperan al menos 3 elementos.", Licencias

        if not isinstance(Licencias, list):
            return False, "Licencias no es una lista.", Licencias

        campo_lookup = CamposEmail[2]  # "Asesor_x0020_ComercialLookupId"

        # =========================
        # 2) Mapa de usuarios
        # =========================
        usuarios = (ComercialesEmails or {}).get("usuarios", [])
        if not isinstance(usuarios, list):
            return False, "Formato inválido en ComercialesEmails: se esperaba una lista en 'usuarios'.", Licencias

        # Mapa: id (str) -> email
        mapa_id_email = {
            str(u.get("id")).strip(): (u.get("email") or "")
            for u in usuarios
            if u.get("id") is not None
        }

        # =========================
        # 3) Procesar licencias
        # =========================
        total_registros = 0
        total_con_email = 0
        total_sin_email = 0

        for item in Licencias:
            if not isinstance(item, dict):
                continue

            total_registros += 1

            lookup_id = str(item.get(campo_lookup, "")).strip()
            email = mapa_id_email.get(lookup_id, "")

            if email:
                total_con_email += 1
            else:
                total_sin_email += 1

            # Reemplazo directo en el MISMO campo
            item[campo_lookup] = email

        # =========================
        # 4) Mensaje OK
        # =========================
        msg_ok = (
            f"OK: {total_registros} licencia(s) procesadas. "
            f"Con email asignado: {total_con_email}. "
            f"Sin email: {total_sin_email}."
        )

        return True, msg_ok, Licencias

    except Exception as ex:
        return False, f"Error validando emails comerciales Lista2: {ex}", Licencias
