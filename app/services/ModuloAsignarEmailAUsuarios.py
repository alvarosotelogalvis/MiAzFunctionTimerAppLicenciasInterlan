# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple


def _norm_str(s) -> str:
    """Normaliza a string minúsculo sin dobles espacios para comparaciones."""
    if s is None:
        return ""
    s = str(s).strip()
    s = " ".join(s.split())
    return s.lower()


def AsignarEmailDeUsuariosALicencias(
    usuarios: List[Dict[str, Any]],
    licencias: List[Dict[str, Any]],
    *,
    campo_email_destino: str = "AsesorEmail",
    # En usuarios (de /users) estos son los campos típicos para email/UPN:
    campos_usuario_email: Tuple[str, ...] = ("mail", "email", "userPrincipalName"),
    # En licencias: posibles campos que podrían traer un email ya guardado:
    campos_licencia_email: Tuple[str, ...] = ("AsesorEmail", "Asesor", "EmailAsesor"),
    # En licencias: posible campo de texto con nombre de la persona (no email):
    campo_persona_texto: str = "Asesor_x0020_Comercial",
    # Si True, intentará empatar por nombre (displayName) cuando no encuentre email:
    intentar_por_display_name: bool = True,
    # Mapa opcional para renombrar claves de usuarios si tu JSON usa otras
    # p. ej. {"displayName": "nombreCompleto"}:
    mapeo_campos_usuario: Optional[Dict[str, str]] = None,
) -> Tuple[bool, Optional[str], Optional[List[Dict[str, Any]]]]:
    """
    Enriquecer 'licencias' asignando el email del usuario correspondiente.

    Estrategia de matching (en orden):
      1) Si la licencia ya trae un email en alguno de 'campos_licencia_email',
         lo valida contra usuarios y lo fija en 'campo_email_destino'.
      2) Si no hay email, y 'intentar_por_display_name' es True:
         - busca el usuario cuyo displayName coincida con el texto en 'campo_persona_texto'.

    NOTA: Esta función NO resuelve LookupId (p. ej. 'Asesor_x0020_ComercialLookupId').
          Si necesitas cruzar por LookupId (71 → email), primero resuélvelo contra
          la User Information List y pasa el email resultante dentro de la licencia.
    """
    try:
        # --- 1) Prepara índices de usuarios por email/UPN y por displayName ---
        # Opcionalmente remapea nombres de campos de usuarios
        def _usr_get(u: Dict[str, Any], key: str) -> Any:
            if mapeo_campos_usuario and key in mapeo_campos_usuario:
                return u.get(mapeo_campos_usuario[key])
            return u.get(key)

        idx_by_email: Dict[str, Dict[str, Any]] = {}
        idx_by_dn: Dict[str, Dict[str, Any]] = {}

        for u in usuarios:
            # Índice por email/UPN
            for k in campos_usuario_email:
                v = _usr_get(u, k)
                if v:
                    idx_by_email[_norm_str(v)] = u

            # Índice por displayName
            dn = _usr_get(u, "displayName")
            if dn:
                idx_by_dn[_norm_str(dn)] = u

        # --- 2) Recorre licencias y asigna email ---
        total = len(licencias)
        asignados = 0
        saltados = 0

        for it in licencias:
            # 2.1) ¿La licencia ya trae algún email útil?
            email_en_licencia = ""
            for campo in campos_licencia_email:
                val = it.get(campo)
                if isinstance(val, str) and "@" in val:
                    email_en_licencia = _norm_str(val)
                    break

            usuario_match: Optional[Dict[str, Any]] = None

            # 2.2) Cruce por email directo si lo trae y existe en usuarios
            if email_en_licencia:
                usuario_match = idx_by_email.get(email_en_licencia)

            # 2.3) Si no hay email, intenta por displayName (si está activado)
            if usuario_match is None and intentar_por_display_name:
                persona_txt = _norm_str(it.get(campo_persona_texto))
                if persona_txt:
                    usuario_match = idx_by_dn.get(persona_txt)

            # 2.4) Si encontramos un usuario, asignamos su email (preferencia: mail -> email -> upn)
            if usuario_match:
                # Extrae el mejor email del usuario
                email_user = None
                for k in campos_usuario_email:
                    v = _usr_get(usuario_match, k)
                    if isinstance(v, str) and "@" in v:
                        email_user = v.strip()
                        break

                if email_user:
                    it[campo_email_destino] = email_user
                    asignados += 1
                else:
                    # Usuario sin email utilizable
                    saltados += 1
            else:
                saltados += 1

        resumen = f"Total licencias: {total} | Asignados: {asignados} | Sin match: {saltados}"
        return True, resumen, licencias

    except Exception as e:
        return False, f"Error: {str(e)}", None
