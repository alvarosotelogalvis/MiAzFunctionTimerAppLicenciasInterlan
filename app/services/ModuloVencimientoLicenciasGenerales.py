from typing import List, Dict, Any, Tuple
from datetime import datetime, timedelta, timezone


def ValidarVencimientosLicencias(
    lista_licencias: List[Dict[str, Any]],
    campo_fecha_vencimiento: str
) -> Tuple[bool, str, List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Clasifica licencias en:
      1) Licencias con vencimiento entre hoy y los próximos 90 días
      2) Licencias sin fecha válida (vacía, nula o no parseable)

    Descarta:
      - Licencias vencidas
      - Licencias con vencimiento mayor a 90 días

    Retorna:
      (success, message, licencias_para_notificar, licencias_sin_fecha_valida)
    """

    try:
        if not isinstance(lista_licencias, list):
            return False, "Error de sistema: la lista de licencias no es una lista.", [], []

        hoy = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
        limite = hoy + timedelta(days=90)

        licencias_para_notificar = []
        licencias_sin_fecha_valida = []

        descartadas_vencidas = 0
        descartadas_fuera_rango = 0

        for lic in lista_licencias:
            if not isinstance(lic, dict):
                continue

            fecha_raw = lic.get(campo_fecha_vencimiento)

            # --- Caso 1: fecha vacía o inexistente ---
            if not fecha_raw:
                licencias_sin_fecha_valida.append(lic)
                continue

            # --- Intentar parseo ---
            try:
                fecha_venc = datetime.fromisoformat(
                    str(fecha_raw).replace("Z", "+00:00")
                ).astimezone(timezone.utc)
            except Exception:
                licencias_sin_fecha_valida.append(lic)
                continue

            fecha_venc = fecha_venc.replace(hour=0, minute=0, second=0, microsecond=0)

            # --- Clasificación ---
            if fecha_venc < hoy:
                descartadas_vencidas += 1
                continue

            if fecha_venc > limite:
                descartadas_fuera_rango += 1
                continue

            # ✅ Dentro del rango válido
            licencias_para_notificar.append(lic)

        msg_ok = (
            f"OK: {len(licencias_para_notificar)} licencia(s) con vencimiento "
            f"entre hoy y los próximos 90 días. "
            f"Licencias sin fecha válida: {len(licencias_sin_fecha_valida)}. "
            f"Descartadas vencidas: {descartadas_vencidas}. "
            f"Descartadas fuera de rango: {descartadas_fuera_rango}."
        )

        return True, msg_ok, licencias_para_notificar, licencias_sin_fecha_valida

    except Exception as ex:
        return False, f"Error validando vencimientos de licencias: {ex}", [], []
        