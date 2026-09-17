from typing import List, Dict, Any, Tuple


def ValidarEmailComercialesLista(
    CamposEmail: tuple,
    Licencias: List[Dict[str, Any]],
    UsuariosTenant: List[Dict[str, Any]]
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:
    """
    Valida que el email comercial asignado a cada licencia:
      - Exista en el tenant
      - Corresponda a un usuario activo

    Clasifica las licencias en:
      - ListaNotificarLicenciasConEmail
      - ListaNotificarLicenciasSinEmail
      - ListaNotificarLincenciasComercialInactivos
    """

    try:
        # =========================
        # 1) Validaciones básicas
        # =========================
        if not isinstance(CamposEmail, (list, tuple)) or len(CamposEmail) < 3:
            return False, "CamposEmail inválido.", {}

        if not isinstance(Licencias, list):
            return False, "Licencias no es una lista.", {}

        if not isinstance(UsuariosTenant, list):
            return False, "UsuariosTenant no es una lista.", {}

        campo_email = CamposEmail[2]  # Asesor_x0020_ComercialLookupId

        # =========================
        # 2) Indexar usuarios del tenant por email
        # =========================
        mapa_email_usuario = {}

        for u in UsuariosTenant:
            if not isinstance(u, dict):
                continue

            email = (
                u.get("mail")
                or u.get("email")
                or u.get("userPrincipalName")
            )

            if not email:
                continue

            mapa_email_usuario[email.strip().lower()] = {
                "accountEnabled": bool(u.get("accountEnabled", False)),
                "userType": u.get("userType", ""),
                "id": u.get("id"),
                "displayName": u.get("displayName")
            }

        # =========================
        # 3) Clasificación de licencias
        # =========================
        con_email = []
        sin_email = []
        comerciales_inactivos = []

        for lic in Licencias:
            if not isinstance(lic, dict):
                continue

            email = str(lic.get(campo_email, "")).strip().lower()

            # --- Sin email ---
            if not email:
                sin_email.append(lic)
                continue

            usuario = mapa_email_usuario.get(email)

            # --- Email no existe en tenant ---
            if not usuario:
                comerciales_inactivos.append(lic)
                continue

            # --- Usuario existe pero está inactivo ---
            if not usuario.get("accountEnabled", False):
                comerciales_inactivos.append(lic)
                continue

            # ✅ Email válido y usuario activo
            con_email.append(lic)

        # =========================
        # 4) Resultado final
        # =========================
        LicenciasFiltradas = {
            "ListaNotificarLicenciasConEmail": con_email,
            "ListaNotificarLicenciasSinEmail": sin_email,
            "ListaNotificarLincenciasComercialInactivos": comerciales_inactivos,
        }

        msg_ok = (
            f"OK: {len(Licencias)} licencia(s) evaluadas. "
            f"Con email válido: {len(con_email)}. "
            f"Sin email: {len(sin_email)}. "
            f"Comerciales inactivos o no encontrados: {len(comerciales_inactivos)}."
        )

        return True, msg_ok, LicenciasFiltradas

    except Exception as ex:
        return False, f"Error validando emails comerciales: {ex}", {}
        