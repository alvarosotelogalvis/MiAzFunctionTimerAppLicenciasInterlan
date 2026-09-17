from typing import List, Dict, Any, Tuple
from datetime import datetime, timedelta, timezone
import re


def ValidarNotificacionesLicenciasLista(
    Licencias: List[Dict[str, Any]],
    campo_notificacion: str,
    campo_vencimiento: str,
    ventana_dias: int = 90
) -> Tuple[bool, str, List[Dict[str, Any]]]:
    """
    Valida y normaliza el campo de notificación de licencias según
    el ciclo de vencimiento actual.

    Si las notificaciones NO corresponden al ciclo vigente,
    el campo se resetea a: 'OK, Notif:'.

    Retorna:
        (success, message, Licencias)
    """

    try:
        if not isinstance(Licencias, list):
            return False, "Licencias no es una lista.", Licencias

        hoy = datetime.now(timezone.utc).replace(
            hour=0, minute=0, second=0, microsecond=0
        )

        patron_notif = re.compile(r"(\d{2}-\d{2}-\d{4})-(\d)-OK")

        total_evaluadas = 0
        total_reseteadas = 0
        total_conservadas = 0

        for lic in Licencias:
            if not isinstance(lic, dict):
                continue

            total_evaluadas += 1

            fecha_raw = lic.get(campo_vencimiento)
            texto_notif = str(lic.get(campo_notificacion, "")).strip()

            # --- Si no hay fecha, resetear ---
            if not fecha_raw:
                lic[campo_notificacion] = "OK, Notif:"
                total_reseteadas += 1
                continue

            try:
                fecha_venc = datetime.fromisoformat(
                    str(fecha_raw).replace("Z", "+00:00")
                ).astimezone(timezone.utc)
            except Exception:
                lic[campo_notificacion] = "OK, Notif:"
                total_reseteadas += 1
                continue

            fecha_venc = fecha_venc.replace(
                hour=0, minute=0, second=0, microsecond=0
            )

            dias_restantes = (fecha_venc - hoy).days

            # --- Fuera de ventana → resetear ---
            if dias_restantes < 0 or dias_restantes > ventana_dias:
                lic[campo_notificacion] = "OK, Notif:"
                total_reseteadas += 1
                continue

            # --- Determinar ciclo actual ---
            if dias_restantes > 60:
                ciclo_actual = "3"
            elif dias_restantes > 30:
                ciclo_actual = "2"
            else:
                ciclo_actual = "1"

            # --- Extraer notificaciones existentes ---
            matches = patron_notif.findall(texto_notif)

            if not matches:
                # No hay notificaciones válidas
                lic[campo_notificacion] = "OK, Notif:"
                total_reseteadas += 1
                continue

            notificaciones_validas = []

            for fecha_txt, ciclo in matches:
                if ciclo == ciclo_actual:
                    notificaciones_validas.append(
                        f"{fecha_txt}-{ciclo}-OK"
                    )

            # --- Reconstruir campo ---
            if not notificaciones_validas:
                lic[campo_notificacion] = "OK, Notif:"
                total_reseteadas += 1
            else:
                lic[campo_notificacion] = "OK, Notif: " + ", ".join(
                    notificaciones_validas
                )
                total_conservadas += 1

        msg_ok = (
            f"OK: {total_evaluadas} licencia(s) evaluadas. "
            f"Campos de notificación reseteados: {total_reseteadas}. "
            f"Campos conservados para ciclo vigente: {total_conservadas}."
        )

        return True, msg_ok, Licencias

    except Exception as ex:
        return (
            False,
            f"Error validando notificaciones de licencias: {ex}",
            Licencias,
        )
        