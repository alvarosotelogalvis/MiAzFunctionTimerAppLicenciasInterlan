# -*- coding: utf-8 -*-
import requests
import time
from bs4 import BeautifulSoup
from datetime import datetime
from typing import Any, Dict, List, Tuple, Optional

__all__ = ["ValidarVencimLicenciasExcelPlatafFortinet"]

def _extract_all_inputs(soup: BeautifulSoup) -> Dict[str, str]:
    return {i.get("name"): i.get("value", "") for i in soup.find_all("input") if i.get("name")}

def ValidarVencimLicenciasExcelPlatafFortinet(
    config: Dict[str, Any],
    lista_aplanada_fortinet: List[Dict[str, Any]],
    *,
    camposvencimiento: Dict[str, str],             # ← tu CAMPOSDEVENCIMIENTO
    escribir_iso_adicional: bool = False,           # deja también "<campo_excel>ISO"
    borrar_alias: bool = True                      # elimina alias (p. ej. FechaVencimiento) si difieren
) -> Tuple[bool, str, Dict[str, List[Dict[str, Any]]]]:
    """
    Valida fechas de vencimiento (Fortinet) y escribe en el campo Excel indicado por CAMPOSDEVENCIMIENTO["Excel"].
    - Formatea el valor al contrato DD/MM/YYYY para el Excel.
    - Opcional: agrega <campo_excel>ISO con YYYY-MM-DDT00:00:00Z.
    - Opcional: elimina alias de fecha distintos del campo Excel canónico.
    """
    if not isinstance(lista_aplanada_fortinet, list):
        return False, "Entrada no es lista.", {"ListasLicenciasFortinetConsultadas": [], "ListasLicenciasFortinetNoConsultadas": []}
    if not isinstance(config, dict) or not config.get("USUARIO_PLAT_FORTI") or not config.get("PASSW_PLAT_FORTI"):
        return False, "Faltan USUARIO_PLAT_FORTI / PASSW_PLAT_FORTI.", {"ListasLicenciasFortinetConsultadas": [], "ListasLicenciasFortinetNoConsultadas": []}
    if not isinstance(camposvencimiento, dict) or not camposvencimiento.get("Excel"):
        return False, "Falta CAMPOSDEVENCIMIENTO['Excel'].", {"ListasLicenciasFortinetConsultadas": [], "ListasLicenciasFortinetNoConsultadas": []}

    campo_excel = camposvencimiento["Excel"]
    alias_conocidos = set([v for v in camposvencimiento.values() if isinstance(v, str)])  # {'FechaVencimiento', 'Fecha_x0020_Fin', 'Fecha Vencimiento (DD/MM/AAAA)'}

    total_entrada = total_skip_noti = total_sin_serial = 0
    total_consultadas_ok = total_http_404 = total_sin_tabla = total_parse_error = total_otros_errores = 0
    consultadas: List[Dict[str, Any]] = []
    no_consultadas: List[Dict[str, Any]] = []

    if len(lista_aplanada_fortinet) == 0:
        return True, "Sin licencias a validar (entrada vacía).", {
            "ListasLicenciasFortinetConsultadas": consultadas,
            "ListasLicenciasFortinetNoConsultadas": no_consultadas,
        }

    base_url = "https://partnerportal.fortinet.com/English/"
    final_url = "https://support.fortinet.com/PartnerRenewalTools/PartnerPortal/SNContractQuery.aspx"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Referer": base_url,
        "Origin": "https://partnerportal.fortinet.com",
    }

    try:
        session = requests.Session()
        # 1) GET inicial
        resp0 = session.get(base_url, headers=headers, timeout=30)
        soup0 = BeautifulSoup(resp0.text, "html.parser"); time.sleep(0.5)
        # 2) Usuario
        payload_user = _extract_all_inputs(soup0)
        payload_user["__EVENTTARGET"] = "ctl00$ctl00$ctl00$GlobalBodyContent$ExternalBodyContent$BodyContent$LoginControl$LinkButton1"
        payload_user["ctl00$ctl00$ctl00$GlobalBodyContent$ExternalBodyContent$BodyContent$LoginControl$UserName"] = config["USUARIO_PLAT_FORTI"]
        resp_user = session.post(base_url, headers=headers, data=payload_user, timeout=30); time.sleep(0.5)
        # 3) Password
        soup1 = BeautifulSoup(resp_user.text, "html.parser")
        payload_pass = _extract_all_inputs(soup1)
        payload_pass["__EVENTTARGET"] = "ctl00$ctl00$ctl00$GlobalBodyContent$ExternalBodyContent$BodyContent$LoginControl$btnSubmit"
        payload_pass["ctl00$ctl00$ctl00$GlobalBodyContent$ExternalBodyContent$BodyContent$LoginControl$Password"] = config["PASSW_PLAT_FORTI"]
        resp_pass = session.post(base_url, headers=headers, data=payload_pass, timeout=30); time.sleep(0.5)
        # 4) Renewal Assets → SSO
        soup2 = BeautifulSoup(resp_pass.text, "html.parser")
        renewal_link = soup2.find("a", string=lambda text: text and "Renewal Assets" in text); time.sleep(0.5)
        if not (renewal_link and renewal_link.get("href")):
            return False, "No se encontró enlace 'Renewal Assets'.", {"ListasLicenciasFortinetConsultadas": [], "ListasLicenciasFortinetNoConsultadas": []}
        href = renewal_link["href"]
        full_url = requests.compat.urljoin(base_url, href)
        resp_renewal = session.get(full_url, headers=headers, timeout=30)
        soup_sso = BeautifulSoup(resp_renewal.text, "html.parser")
        form = soup_sso.find("form")
        if not form:
            return False, "No se encontró el formulario SSO hacia soporte.", {"ListasLicenciasFortinetConsultadas": [], "ListasLicenciasFortinetNoConsultadas": []}
        action = form.get("action")
        sso_url = requests.compat.urljoin(resp_renewal.url, action)
        sso_payload = _extract_all_inputs(soup_sso)
        _ = session.post(sso_url, headers=headers, data=sso_payload, allow_redirects=True, timeout=30)
        # 5) Página final
        resp_final = session.get(final_url, headers=headers, timeout=30)
        soup_final = BeautifulSoup(resp_final.text, "html.parser")
        if not soup_final.find("form", {"id": "aspnetForm"}):
            return False, "No se encontró el formulario final de soporte.", {"ListasLicenciasFortinetConsultadas": [], "ListasLicenciasFortinetNoConsultadas": []}
    except Exception as e:
        return False, f"Error durante login/SSO Fortinet: {e}", {"ListasLicenciasFortinetConsultadas": [], "ListasLicenciasFortinetNoConsultadas": []}

    # ---- Consulta por fila ----
    for row in lista_aplanada_fortinet:
        if not isinstance(row, dict):
            continue
        total_entrada += 1
        row_out = dict(row)

        noti = str(row.get("NotificarVencimiento", "")).strip().lower()
        if noti == "no":
            total_skip_noti += 1
            no_consultadas.append(row_out); continue

        serial = str(row.get("Serial", "")).strip()
        if not serial:
            total_sin_serial += 1
            no_consultadas.append(row_out); continue

        try:
            soup_form = BeautifulSoup(resp_final.text, "html.parser")
            form_data = _extract_all_inputs(soup_form)
            form_data["ctl00$RenewTrackingMainContent$TB_SerialNumber"] = serial
            form_data["ctl00$RenewTrackingMainContent$Btn_Submit"] = "Submit"

            resp_serial = session.post(final_url, headers=headers, data=form_data, timeout=30)
            if resp_serial.status_code == 404:
                total_http_404 += 1; no_consultadas.append(row_out); continue

            soup_res = BeautifulSoup(resp_serial.text, "html.parser")
            tabla = (
                soup_res.find("table", {"id": "ctl00_RenewTrackingMainContent_UC_SupportCoverage1_GV_SupportInfo"})
                or soup_res.find("table", {"id": "ctl00_RenewTrackingMainContent_UC_HWWarrantyInfo1_GV_WarrantyTermInfo"})
                or soup_res.find("table", {"id": "ctl00_RenewTrackingMainContent_UC_WarrantyInfo1_GV_WarrantyTermInfo"})
            )
            if not tabla:
                total_sin_tabla += 1; no_consultadas.append(row_out); continue

            filas = tabla.find_all("tr")
            encontrado = False
            for i, fila in enumerate(filas):
                ths = fila.find_all("th")
                if ths:
                    headers_txt = [th.get_text(strip=True) for th in ths]
                    if "Activation Date" in headers_txt and "Expiration Date" in headers_txt:
                        if i + 1 < len(filas):
                            fila_datos = filas[i + 1]
                            tds = fila_datos.find_all("td")
                            if len(tds) >= 4:
                                fecha_fin_txt = tds[3].get_text(strip=True)  # 'YYYY-MM-DD'
                                try:
                                    fecha_obj = datetime.strptime(fecha_fin_txt, "%Y-%m-%d")
                                    # Campo canónico de Excel (dinámico)
                                    row_out[campo_excel] = fecha_obj.strftime("%d/%m/%Y")
                                    if escribir_iso_adicional:
                                        row_out[f"{campo_excel}ISO"] = fecha_obj.strftime("%Y-%m-%dT00:00:00Z")
                                    # Limpia alias para evitar columnas duplicadas
                                    if borrar_alias:
                                        for alias in alias_conocidos:
                                            if alias != campo_excel and alias in row_out:
                                                del row_out[alias]
                                    consultadas.append(row_out)
                                    total_consultadas_ok += 1
                                    encontrado = True
                                except Exception:
                                    total_parse_error += 1
                        break

            if not encontrado:
                no_consultadas.append(row_out)

            time.sleep(0.6)

        except Exception:
            total_otros_errores += 1
            no_consultadas.append(row_out)

    msg = (
        "Fortinet OK. "
        f"Entradas: {total_entrada}. OK: {total_consultadas_ok}. Notificar=no: {total_skip_noti}. "
        f"Sin Serial: {total_sin_serial}. Sin tabla: {total_sin_tabla}. Parse fecha: {total_parse_error}. "
        f"HTTP 404: {total_http_404}. Otros: {total_otros_errores}."
    )
    return True, msg, {
        "ListasLicenciasFortinetConsultadas": consultadas,
        "ListasLicenciasFortinetNoConsultadas": no_consultadas,
    }