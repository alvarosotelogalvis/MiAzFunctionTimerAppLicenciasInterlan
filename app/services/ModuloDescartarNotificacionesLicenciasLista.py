
from typing import Any, Dict, List, Tuple
from datetime import datetime, timezone, date

def _parse_fecha(fecha_val):
    if fecha_val is None:
        return None
    if isinstance(fecha_val, date) and not isinstance(fecha_val, datetime):
        return fecha_val
    if isinstance(fecha_val, datetime):
        return fecha_val.date()

    s = str(fecha_val).strip()
    if not s:
        return None

    try:
        if "T" in s:
            s2 = s.split("T", 1)[0]
            return datetime.strptime(s2, "%Y-%m-%d").date()
    except Exception:
        pass
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        pass
    try:
        return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        pass
    try:
        return datetime.strptime(s, "%d-%m-%Y").date()
    except Exception:
        pass
    return None

def DescartarNotificacionesLicenciasLista(
    campo_vencimiento: str,
    campo_alerta: str,
    licencias: List[Dict[str, Any]]
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:
    """
    Descarta licencias que ya tienen registrada la notificación correspondiente
    a la ventana actual (1, 2 o 3), para evitar re-notificar.

    Retorna:
    {
      "LicenciasDescartadas": [...],
      "LicenciasNoDescartadas": [...]
    }
    """

    hoy = datetime.now(timezone.utc).date()

    descartadas: List[Dict[str, Any]] = []
    no_descartadas: List[Dict[str, Any]] = []

    for it in licencias:
        if not isinstance(it, dict):
            continue

        fecha_venc = _parse_fecha(it.get(campo_vencimiento))
        if fecha_venc is None:
            # No se puede evaluar, no se descarta
            no_descartadas.append(it)
            continue

        diff = (fecha_venc - hoy).days

        # Determinar ventana actual
        if 0 <= diff < 30:
            n_actual = 1
        elif 30 <= diff < 60:
            n_actual = 2
        elif 60 <= diff <= 90:
            n_actual = 3
        else:
            # Fuera de ventanas de notificación
            no_descartadas.append(it)
            continue

        alerta_raw = it.get(campo_alerta, "") or ""

        # Buscar si ya existe notificación para esta ventana
        ya_notificado = False

        if "Notif:" in alerta_raw:
            _, _, cola = alerta_raw.partition("Notif:")
            segmentos = cola.split(",")

            for seg in segmentos:
                seg = seg.strip()
                partes = seg.split("-")
                # Esperamos formato: dd-mm-yyyy-N-OK
                if len(partes) >= 5:
                    try:
                        n_registrado = int(partes[3])
                        if n_registrado == n_actual:
                            ya_notificado = True
                            break
                    except Exception:
                        continue

        if ya_notificado:
            descartadas.append(it)
        else:
            no_descartadas.append(it)

    resumen = {
        "LicenciasDescartadas": descartadas,
        "LicenciasNoDescartadas": no_descartadas
    }

    mensaje = (
        f"Descartadas: {len(descartadas)} | "
        f"No descartadas: {len(no_descartadas)}"
    )

    return True, mensaje, resumen
