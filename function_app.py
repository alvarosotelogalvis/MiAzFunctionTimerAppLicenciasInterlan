import logging
import azure.functions as func
import json
import datetime
import os

#Licencias Del Excel
from app.services.ModuloLeerLicenciasExcelDeSharepoint import LeerLicenciasExcelDeSharepoint
from app.services.ModuloVerificarCamposObligatoriosEnExcel import VerificarCamposObligatoriosEnExcelDeSharepoint
from app.services.ModuloDescartarLicenciasDeExcel import DescartarLicenciasDeExcelPorNotificacionNo
from app.services.ModuloLicenciasFiltradasFortiExcel import FiltrarLicenciasDeExcelFortinet
from app.services.ModuloFiltrarLicenciasExcelFortiConsulta import FiltrarLicenciasExcelFortinetParaConsultar
from app.services.ModuloLicenciasExcelFortinet import ValidarVencimLicenciasExcelPlatafFortinet
from app.services.ModuloValidarVencimientosLicenciasExcel import ValidarVencimientosLicenciasExcel
from app.services.ModuloValidarEmailComercialesExcel import ValidarEmailComercialesExcel
from app.services.ModuloValidarNotificacionesLicenciasExcel import ValidarNotificacionesLicenciasExcel
from app.services.ModuloNotificacionVencimientosExcel import PrepararNotificacionesExcel
from app.services.ModuloActualizarLicenciasExcelEnSharepoint import ActualizarLicenciasExcelEnSharepoint

#Licencias Lista2: Gestion de licencias
from app.services.ModuloLicenciasDeSharePointLista import LeerUnalistaDeSharepoint
from app.services.ModuloVerificarCamposObligatoriosLista import VerificarCamposObligatoriosLista
from app.services.ModuloDescartarLicenciasListasPorAlertaNoOk import DescartarLicenciasListasPorAlertaNoOk
from app.services.ModuloLeerUsuariosTenant import LeerUsuariosDelTenant
from app.services.ModuloVencimientoLicenciasGenerales import ValidarVencimientosLicencias
from app.services.ModuloAsignarEmailComercialesLista import AsignarEmailComercialesLista
from app.services.ModuloValidarEmailComercialesLista import ValidarEmailComercialesLista
from app.services.ModuloValidarNotificacionesLicenciasLista import ValidarNotificacionesLicenciasLista
from app.services.ModuloDescartarNotificacionesLicenciasLista import DescartarNotificacionesLicenciasLista
from app.services.ModuloNotificacionVencimientos2 import PrepararNotificacionesLista

#Envío combinado (un solo correo por tipo, fusionando Excel + Listas) y alertas de fallo
from app.services.ModuloEnviarNotificacionesCombinadas import EnviarNotificacionesCombinadas
from app.services.ModuloEnviarCorreoSeguimientoLogsAzureFunction import (
    EnviarCorreoSeguimientoLogsCombinado,
    EnviarCorreoDeErrorPipeline,
)

app = func.FunctionApp()

#schedule="0 0 7 1,16 * *" los dias 1 y 16 de cada mes a las 7 am
#schedule="0 0 * * * *" todos los dias cada hora en punto
#schedule="0 0 7 * * *" todos los dias a las 7 am
@app.timer_trigger(schedule="0 0 7 * * *", arg_name="myTimer",
              run_on_startup=os.environ.get("RUN_ON_STARTUP", "false").lower() == "true",
              use_monitor=False)
def timerAppLicenciasInterlan(myTimer: func.TimerRequest) -> None:
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    MSGLOGS = {
        "Inicio": f"[{timestamp}] Inicio Azure Function AppLicenciasInterlan",
        "Eventos": []
    }
    MSGLOGS2 = {
        "Inicio": f"[{timestamp}] Inicio Azure Function AppLicenciasInterlan",
        "Eventos": []
    }
    CONFIG = {
        "ENTORNO": "desarrollo", # "desarrollo" o "produccion"
        "USUARIO_PLAT_FORTI": os.environ.get("FORTINET_USER", ""),
        "PASSW_PLAT_FORTI": os.environ.get("FORTINET_PASSWORD", ""),
        #ID_aplicación: app licencias
        "TENANT_ID": os.environ.get("APP1_TENANT_ID", ""),
        "CLIENT_ID": os.environ.get("APP1_CLIENT_ID", ""),
        "CLIENT_SECRET": os.environ.get("APP1_CLIENT_SECRET", ""),
        # ID del sitio (formato host,siteCollectionId,siteId)
        "SITE_ID": "interlan.sharepoint.com,b06178d9-4ff3-4b70-94c4-06cce7edac8d,550dd537-61fd-4ca0-bed6-2edb25541cbc",
        "LIST_DISPLAY_NAME": "Gestion de Licencias",
        # opcionales:
        "TOP": 50,
        "ORDERBY": "id asc",
        "IMPRIMIR_COLUMNAS": False,
        "SENDER": os.environ.get("MAIL_SENDER", "")
    }
    CONFIG2 = {
        #ID_aplicación: app licencias
        "TENANT_ID": os.environ.get("APP2_TENANT_ID", ""),
        "CLIENT_ID": os.environ.get("APP2_CLIENT_ID", ""),
        "CLIENT_SECRET": os.environ.get("APP2_CLIENT_SECRET", ""),
    }

    def _split_env_emails(var_name: str) -> list:
        raw = os.environ.get(var_name, "")
        return [e.strip() for e in raw.split(",") if e.strip()]

    DESTINATARIOS = {
        "Destinatarios": _split_env_emails("MAIL_DESTINATARIOS"),
        "Seguimiento": _split_env_emails("MAIL_SEGUIMIENTO"),
        "SinDestinatarios": _split_env_emails("MAIL_SIN_DESTINATARIOS"),
        "AlertasFallo": _split_env_emails("MAIL_ALERTAS_FALLO"),
    }

    # --------------------------------------------------------------------------------
    # Checkpoint de fallos: reemplaza al viejo ManejarErrorCritico/EsProduccion (que
    # nunca se invocaban y llamaban a una función inexistente). Cada punto de fallo de
    # las pipelines llama a uno de estos checkpoints en vez de solo loguear y retornar;
    # el checkpoint loguea Y dispara un correo de alerta real con el paso, el mensaje
    # de error y todo lo que se alcanzó a loguear de esa(s) pipeline(s) hasta ahora.
    # --------------------------------------------------------------------------------
    def _hacer_checkpoint(msglogs_list, pipeline_label):
        def checkpoint(success, message, step_label):
            if success:
                return True
            logging.error(f"[{step_label}] {message}")
            try:
                ok, err = EnviarCorreoDeErrorPipeline(
                    CONFIG, msglogs_list, DESTINATARIOS, pipeline_label, step_label, message
                )
                if not ok:
                    logging.error(f"[ReportarFallo] No se pudo enviar el correo de alerta: {err}")
            except Exception as ex:
                logging.error(f"[ReportarFallo] Excepción enviando correo de alerta: {ex}")
            return False
        return checkpoint

    checkpoint_excel = _hacer_checkpoint([MSGLOGS], "Excel")
    checkpoint_lista = _hacer_checkpoint([MSGLOGS2], "Lista")
    checkpoint_final = _hacer_checkpoint([MSGLOGS, MSGLOGS2], "Consolidado")

    def _advertir(msglogs_list, pipeline_label, step_label, message):
        """
        Para pasos que se degradan en vez de fallar (ej. Fortinet no disponible tras
        reintentos): no corta el proceso, pero sí avisa con un correo de advertencia
        (naranja) distinto del de fallo (rojo) que dispara _hacer_checkpoint.
        """
        logging.warning(f"[{step_label}] {message}")
        try:
            ok, err = EnviarCorreoDeErrorPipeline(
                CONFIG, msglogs_list, DESTINATARIOS, pipeline_label, step_label, message,
                Severidad="Advertencia"
            )
            if not ok:
                logging.error(f"[Advertencia] No se pudo enviar el correo de advertencia: {err}")
        except Exception as ex:
            logging.error(f"[Advertencia] Excepción enviando correo de advertencia: {ex}")

    try:
        COMERCIALES_EMAILS = json.loads(os.environ.get("COMERCIALES_EMAILS_JSON", "{}"))
    except json.JSONDecodeError:
        COMERCIALES_EMAILS = {}
    LISTAS = {
        "SITE_ID_1": "interlan.sharepoint.com,b06178d9-4ff3-4b70-94c4-06cce7edac8d,550dd537-61fd-4ca0-bed6-2edb25541cbc",
        "LIST_DISPLAY_NAME_1": "ListaAppLicenciasInterlan",
        "SITE_ID_2": "interlan.sharepoint.com,b06178d9-4ff3-4b70-94c4-06cce7edac8d,550dd537-61fd-4ca0-bed6-2edb25541cbc",
        "LIST_DISPLAY_NAME_2": "Gestion de Licencias",
    }
    CAMPOS_OBLIGATORIOS = {
        "Lista1": ["IDLicencia", "Dominio", "Cliente", "Fabricante", "TIPO", "Elemento", "Producto", "Version", "Cantidad", "Serial", "FechaVencimiento", "Estado", "DiasDeVencimiento", "NombreComercial", "EmailComercial", "NotificarVencimiento", "SeNotifico", "Observaciones"],
        "Lista2": ["id", "Consecutivo", "Fabricante", "Producto", "Cantidad", "Vigencia", "Fecha_x0020_Fin", "Nombre_x0020_Cliente", "Asesor_x0020_ComercialLookupId", "Alerta"],
        "Excel": ["SerieID", "LicenciaID", "Producto", "Serial", "Tipo", "Version", "Cantidad", "Cliente", "Comercial", "Dominio", "DiasDeVencimiento", "Elemento", "EmailComercial", "Estado", "Fabricante", "FechaDeVencimientoDDMMAAAA", "NotificarVencimiento", "Notificaciones", "Observaciones"]
    }
    CAMPOSDEVENCIMIENTO = {
        "Lista1": "FechaVencimiento",
        "Lista2": "Fecha_x0020_Fin",
        "Excel": "FechaDeVencimientoDDMMAAAA"
    }
    CAMPOSDEALERTAONOTIFICACION = {
        "Lista1": "SeNotifico",
        "Lista2": "Alerta",
        "Excel": "Notificaciones"
    }
    CamposEmail = ("Asesor Comercial", "EmailComercial", "Asesor_x0020_ComercialLookupId") #tupla
    DOMINIOS_CORREO_PERMITIDOS = ("@interlan.com.co")
    DATOS_EXCEL_SHAREPOINT = {
        "site_hostname": "interlan.sharepoint.com",
        "site_path": "/sites/csa",
        # 👇 cambia a "produccion" cuando muevas a PROD
        "entorno": "desarrollo",   #"desarrollo" o "produccion"
        # Archivo de PRODUCCIÓN (se usa también para lectura)
        #"file_relative_path": "Licencias y Garantías/InvLicPython3.xlsx",
        #"file_relative_path": "Licencias y Garantías/INV-LICENCIAS CLIENTES.xlsx",
        "file_relative_path": "Licencias y Garantías/ArchivosPython/PYTHON_INV-LICENCIAS CLIENTES_NORMALIZADO.xlsx",
        # (opcional) destino DEV; si lo omites, se deriva como *_DEV.xlsx
        "file_relative_path_dev": "Licencias y Garantías/ArchivosPython/PYTHONInvLicPython3_NOBORRAR_DEV.xlsx",
        # Plantilla válida para crear si el destino NO existe (DEV/PROD)
        "template_relative_path": "Licencias y Garantías/ArchivosPython/PYTHONPlantillaInventarioLicenciasNOBORRAR.xlsx",
        "hojas_permitidas": [
            "SEG_EN_LA_RED", "MONITOREO", "CONTINUIDAD", "SEGURIDAD INFORMATICA",
            "SEGURIDAD DE LA INFORMACIÓN", "INFRAESTRUCTURA", "MEDR",
            "UP SEGURIDAD DE LA INFORMACION", "UP SEGURIDAD", "UP NETWORKING",
            "UP INFRAESTRUCTURA", "UP BACKUP", "UP SEGURIDAD - AVAL"
        ],
        "campos_salida_orden": [
            "SerieID", "LicenciaID", "Producto", "Serial", "Tipo", "Version",
            "Cantidad", "Cliente", "Comercial", "Dominio", "DiasDeVencimiento",
            "Elemento", "EmailComercial", "Estado", "Fabricante", "FechaDeVencimientoDDMMAAAA",
            "NotificarVencimiento", "Notificaciones", "Observaciones"
        ],
        "campos_fecha_ddmmyyyy": ["FechaDeVencimientoDDMMAAAA"],
        "aplanar": True,
        "nombre_campo_hoja": "HojaExcel",
        # === Escritura (opcional; defaults mostrados) ===
        "behavior_when_exists": "update", #"update" o "overwrite_from_template"
        "batch_size": 1000,
        "remove_other_sheets": True,
        # "preserve_sheets": ["Resumen"]
    }
    #DATOS_EXCEL_SHAREPOINT_LISTAS = {
    '''DATOS_EXCEL_SHAREPOINT_2 = {
        "site_hostname": "interlan.sharepoint.com",
        "site_path": "/sites/csa",
        "entorno": "desarrollo",  # desarrollo o "produccion"
        "file_relative_path": "Licencias y Garantías/PYTHON_INV-LICENCIAS CLIENTES_NORMALIZADO.xlsx",
        "file_relative_path_dev": "Licencias y Garantías/PYTHONInvLicPython3_NOBORRAR_DEV.xlsx",
        "template_relative_path": "Licencias y Garantías/PYTHONPlantillaInventarioGestionDeLicenciasNOBORRAR.xlsx",
        "sheet_name": "InvGestionLicencias"
    }'''
    DATOS_EXCEL_SHAREPOINT_LISTAS = {
        "site_hostname": "interlan.sharepoint.com",
        "site_path": "/sites/csa",
        # 👇 cambia a "produccion" cuando muevas a PROD
        "entorno": "desarrollo",   #"desarrollo" o "produccion"
        # Archivo de PRODUCCIÓN (se usa también para lectura)
        #"file_relative_path": "Licencias y Garantías/InvLicPython3.xlsx",
        #"file_relative_path": "Licencias y Garantías/INV-LICENCIAS CLIENTES.xlsx",
        "file_relative_path": "Licencias y Garantías/PYTHON_INV-LICENCIAS CLIENTES_NORMALIZADO.xlsx",
        # (opcional) destino DEV; si lo omites, se deriva como *_DEV.xlsx
        "file_relative_path_dev": "Licencias y Garantías/PYTHONInvLicPython3_NOBORRAR_DEV.xlsx",
        # Plantilla válida para crear si el destino NO existe (DEV/PROD)
        "template_relative_path": "Licencias y Garantías/PYTHONPlantillaInventarioGestionDeLicenciasNOBORRAR.xlsx",
        "hojas_permitidas": [
            "SEG_EN_LA_RED", "MONITOREO", "CONTINUIDAD", "SEGURIDAD INFORMATICA",
            "SEGURIDAD DE LA INFORMACIÓN", "INFRAESTRUCTURA", "MEDR",
            "UP SEGURIDAD DE LA INFORMACION", "UP SEGURIDAD", "UP NETWORKING",
            "UP INFRAESTRUCTURA", "UP BACKUP", "UP SEGURIDAD - AVAL"
        ],
        "campos_salida_orden": [
            "SerieID", "LicenciaID", "Producto", "Serial", "Tipo", "Version",
            "Cantidad", "Cliente", "Comercial", "Dominio", "DiasDeVencimiento",
            "Elemento", "EmailComercial", "Estado", "Fabricante", "FechaDeVencimientoDDMMAAAA",
            "NotificarVencimiento", "Notificaciones", "Observaciones"
        ],
        "campos_fecha_ddmmyyyy": ["FechaDeVencimientoDDMMAAAA"],
        "aplanar": True,
        "nombre_campo_hoja": "HojaExcel",
        # === Escritura (opcional; defaults mostrados) ===
        "behavior_when_exists": "update", #"update" o "overwrite_from_template"
        "batch_size": 1000,
        "remove_other_sheets": True,
        # "preserve_sheets": ["Resumen"]
    }

    # ===========================================================================================
    # LISTA DE LICENCIAS EXCEL (ARCHIVO EXCEL LICPYTHON.XLSX) EN SHAREPOINT.

    # Retorna la lista de licencias leídas del Excel por hojas -> SharePoint
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, ListaDeLicenciasSharePoint = LeerLicenciasExcelDeSharepoint(
        CONFIG,
        DATOS_EXCEL_SHAREPOINT
    )
    if not checkpoint_excel(success, message, "LeerLicenciasExcelDeSharepoint"):
        return
    # Solo se llega aquí si TODO salió bien
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Leer Las Licencias Del Excel desde SharePoint",
        "message": message
    })

    # Verifica los campos obligatorios en el excel de sharepoint.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, ListaDeLicenciasSharePoint = VerificarCamposObligatoriosEnExcelDeSharepoint(
        ListaDeLicenciasSharePoint,
        CAMPOS_OBLIGATORIOS["Excel"]
    )
    if not checkpoint_excel(success, message, "VerificarCamposObligatoriosEnExcelDeSharepoint"):
        return
    #print('Verificar Campos Obligatorios En Excel De Sharepoint:')
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Verificar Campos Obligatorios En Licencias Excel De Sharepoint:",
        "message": message
    })

    # Descartar licencias de excel por "NotificarVencimiento": "no"
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = DescartarLicenciasDeExcelPorNotificacionNo(
        ListaDeLicenciasSharePoint,
        CAMPOSDEALERTAONOTIFICACION,
        origen="Excel"  # aquí también podrías pasar "Lista1" o "Lista2" en otros flujos
    )
    if not checkpoint_excel(success, message, "DescartarLicenciasDeExcelPorNotificacionNo"):
        return
    #print('DescartarLicenciasDeExcelPorNotificacionNo: OK')
    LicenciasDescartadas = LicenciasFiltradas.get("LicenciasDescartadas", []) or []
    ListaDeLicenciasSharePoint = LicenciasFiltradas.get("LicenciasNoDescartadas", []) or []
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Descartar Licencias De Excel Por Notificacion 'No':",
        "message": message
    })

    # filtrar las licencias solo fortinet.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = FiltrarLicenciasDeExcelFortinet(ListaDeLicenciasSharePoint)
    if not checkpoint_excel(success, message, "FiltrarLicenciasDeExcelFortinet"):
        return
    #print('FiltrarLicenciasDeExcelFortinet: OK')
    ListaDeLicenciasSharePoint = LicenciasFiltradas["LicenciasExcelPorHojasNoFortinet"]
    ListaDeLicenciasSharePointSiFortinet    = LicenciasFiltradas["LicenciasExcelPorHojasSiFortinet"]
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Filtrar Licencias De Excel Fabricante Fortinet:",
        "message": message
    })

    #Paso: Filtrar licencias Fortinet que serán consultadas y las que no
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = FiltrarLicenciasExcelFortinetParaConsultar(
        ListaDeLicenciasSharePointSiFortinet
    )
    if not checkpoint_excel(success, message, "FiltrarLicenciasExcelFortinetParaConsultar"):
        return
    #print('FiltrarLicenciasExcelFortinetParaConsultar: OK')
    ListaLicenciasExcelFortinetParaConsultar = LicenciasFiltradas["ListasLicenciasFortinetParaConsultar"]
    #ListasLicenciasFortinetNoConsultar  = LicenciasFiltradas["ListasLicenciasFortinetNoConsultar"]
    ListaDeLicenciasSharePoint.extend(LicenciasFiltradas["ListasLicenciasFortinetNoConsultar"])
    #print('FiltrarLicenciasExcelFortinetParaConsultar: OK')
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Filtrar Licencias Excel Fabricante Fortinet Para Consultar:",
        "message": message
    })

    #Toma la lista de Licencias Forti y valida sus vencimientos en la plataforma Fortinent.
    #Devuelve una lista con los registros de licencias para ser actualizados en el share point de nuevo
    # Si no hay licencias Fortinet, puedes ahorrar la llamada externa:
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = ValidarVencimLicenciasExcelPlatafFortinet(
        CONFIG,
        ListaLicenciasExcelFortinetParaConsultar,
        camposvencimiento=CAMPOSDEVENCIMIENTO,
    )
    if not checkpoint_excel(success, message, "ValidarVencimLicenciasExcelPlatafFortinet"):
        return
    if LicenciasFiltradas.get("LoginFortinetFallido"):
        # No se pudo conectar a Fortinet tras varios reintentos: se degrada en vez de
        # abortar. Las licencias afectadas ya vienen en "ListasLicenciasFortinetNoConsultadas"
        # (conservan la fecha que ya tenían), y se avisa con un correo de advertencia
        # (no de fallo total) para que quede constancia sin frenar el resto del proceso.
        _advertir([MSGLOGS], "Excel", "ValidarVencimLicenciasExcelPlatafFortinet", message)
    # Extender ListaDeLicenciasSharePoint consultadas y no consultadas
    ListaDeLicenciasSharePoint.extend(LicenciasFiltradas["ListasLicenciasFortinetConsultadas"])
    ListaDeLicenciasSharePoint.extend(LicenciasFiltradas["ListasLicenciasFortinetNoConsultadas"])
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Validar Vencimiento licencias Excel En la Plataforma Fortinet:",
        "message": message
    })

    # Se valida el vencimiento de las licencias y se devuelve las que se deben notificar po fecha de vencimiento.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = ValidarVencimientosLicenciasExcel(
        ListaDeLicenciasSharePoint,
        campo_vencimiento=CAMPOSDEVENCIMIENTO["Excel"],
        ventana_dias=90
    )
    if not checkpoint_excel(success, message, "ValidarVencimientosLicenciasExcel"):
        return
    ListaDeLicenciasSharePoint = LicenciasFiltradas["LicenciasNoNotificar"]
    ListaLicenciasParaNotificar = LicenciasFiltradas["LicenciasParaNotificar"]
    LicenciasSinFecha = LicenciasFiltradas["LicenciasSinFecha"]
    LicenciasSinFecha.extend(LicenciasFiltradas["LicenciasFechaInvalida"])
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Validar Vencimientos Licencias Excel:",
        "message": message
    })

    # Se valida el vencimiento de las licencias corresponda con la ventana de 3 meses anteriores, si es asi se deja
    # el registro, sino se cambia por defecto pero no se descarta ninguno
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = ValidarNotificacionesLicenciasExcel(
        ListaLicenciasParaNotificar,
        CAMPOSDEALERTAONOTIFICACION["Excel"],
        campo_vencimiento=CAMPOSDEVENCIMIENTO["Excel"],
        ventana_dias=90
    )
    if not checkpoint_excel(success, message, "ValidarNotificacionesLicenciasExcel"):
        return
    ListaLicenciasParaNotificar = LicenciasFiltradas["LicenciasNotificacionesValidasCompletas"]
    ListaLicenciasParaNotificar.extend(LicenciasFiltradas["LicenciasNotificacionesIncompletas"])
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Validar Notificaciones Licencias Excel:",
        "message": message
    })

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = ValidarEmailComercialesExcel(
        CONFIG2, CamposEmail, ListaLicenciasParaNotificar
    )
    if not checkpoint_excel(success, message, "ValidarEmailComercialesExcel"):
        return
    ListaNotificarLicenciasConEmail = LicenciasFiltradas.get("ListaNotificarLicenciasConEmail", [])
    ListaNotificarLicenciasSinEmail = LicenciasFiltradas.get("ListaNotificarLicenciasSinEmail", [])
    ListaNotificarLincenciasComercialInactivos = LicenciasFiltradas.get("ListaNotificarLincenciasComercialInactivos", [])
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Validar Email Comerciales Excel:",
        "message": message
    })

    # Clasifica (sin enviar ni sellar) las licencias en ventanas de vencimiento,
    # agrupadas por comercial. El envío real y el sellado ocurren más abajo, una
    # sola vez para ambas pipelines, en EnviarNotificacionesCombinadas.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, PlanExcel = PrepararNotificacionesExcel(
        CamposEmail,
        ListaNotificarLicenciasConEmail,
        ListaNotificarLicenciasSinEmail,
        ListaNotificarLincenciasComercialInactivos,
        LicenciasSinFecha
    )
    if not checkpoint_excel(success, message, "PrepararNotificacionesExcel"):
        return
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Preparar Notificaciones Excel:",
        "message": message
    })

    ListaDeLicenciasSharePoint.extend(PlanExcel["raw_lists"]["ConEmail"])
    ListaDeLicenciasSharePoint.extend(PlanExcel["raw_lists"]["SinEmail"])
    ListaDeLicenciasSharePoint.extend(PlanExcel["raw_lists"]["Inactivos"])
    ListaDeLicenciasSharePoint.extend(LicenciasDescartadas)

    # ===========================================================================================
    # LISTA DE LICENCIAS COMO LISTA (LISTA2: GESTION DE LICENCIAS) EN SHAREPOINT.

    #Retorna una lista de licencias leidas del sharepoint.
    CONFIG["SITE_ID"] = LISTAS["SITE_ID_2"]
    CONFIG["LIST_DISPLAY_NAME"] = LISTAS["LIST_DISPLAY_NAME_2"]
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, ListaDeLicenciasSharePointLista = LeerUnalistaDeSharepoint(CONFIG)
    if not checkpoint_lista(success, message2, "LeerUnalistaDeSharepoint"):
        return
    print('LeerUnalistaDeSharepoint Lista2: OK')
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Leer Una lista De Sharepoint (Lista2):",
        "message": message2
    })
    CONFIG["SITE_ID"] = LISTAS["SITE_ID_1"]
    CONFIG["LIST_DISPLAY_NAME"] = LISTAS["LIST_DISPLAY_NAME_1"]

    # Verifica los campos obligatorios en el excel de sharepoint.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, ListaDeLicenciasSharePointLista = VerificarCamposObligatoriosLista(
        ListaDeLicenciasSharePointLista, CAMPOS_OBLIGATORIOS["Lista2"]
    )
    if not checkpoint_lista(success, message2, "VerificarCamposObligatoriosLista"):
        return
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Verificar Campos Obligatorios Lista (Lista2):",
        "message": message2
    })

    #Aqui se descartan pero no se devuelven las descartadas por lo que se "Pierden"
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, LicenciasParaNotificar = DescartarLicenciasListasPorAlertaNoOk(
        ListaDeLicenciasSharePointLista, CAMPOSDEALERTAONOTIFICACION["Lista2"]
    )
    if not checkpoint_lista(success, message2, "DescartarLicenciasListasPorAlertaNoOk"):
        return
    print('ValidarVencimLicenciasGenerales: OK')
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Descartar Licencias Listas Por Alerta NoOk (Lista2):",
        "message": message2
    })

    #Descripcion funcion.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, LicenciasParaNotificar, LicenciasSinFechaValida = ValidarVencimientosLicencias(
        LicenciasParaNotificar, CAMPOSDEVENCIMIENTO["Lista2"]
    )
    if not checkpoint_lista(success, message2, "ValidarVencimientosLicencias"):
        return
    print('ValidarVencimLicenciasGenerales: OK')
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Validar Vencimientos Licencias (Lista2):",
        "message": message2
    })

    # Lee los usuarios del tenant.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, UsuariosTenant = LeerUsuariosDelTenant(CONFIG)
    if not checkpoint_lista(success, message2, "LeerUsuariosDelTenant"):
        return
    print('LeerUsuariosDelTenant Lista: OK')
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Leer Usuarios Del Tenant (Lista2):",
        "message": message2
    })

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, LicenciasParaNotificar = AsignarEmailComercialesLista(
        COMERCIALES_EMAILS, CamposEmail, LicenciasParaNotificar
    )
    if not checkpoint_lista(success, message2, "AsignarEmailComercialesLista"):
        return
    print('ValidarEmailComerciales: OK')
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Asignar Email Comerciales Lista (Lista2):",
        "message": message2
    })

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, LicenciasParaNotificar = ValidarNotificacionesLicenciasLista(
        LicenciasParaNotificar,
        CAMPOSDEALERTAONOTIFICACION["Lista2"],
        campo_vencimiento=CAMPOSDEVENCIMIENTO["Lista2"],
        ventana_dias=90
    )
    if not checkpoint_lista(success, message2, "ValidarNotificacionesLicenciasLista"):
        return
    print('ValidarNotificacionesLicenciasLista: OK')
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Validar Notificaciones Licencias Lista (Lista2):",
        "message": message2
    })

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, LicenciasFiltradas = DescartarNotificacionesLicenciasLista(
        CAMPOSDEVENCIMIENTO["Lista2"],
        CAMPOSDEALERTAONOTIFICACION["Lista2"],
        LicenciasParaNotificar
    )
    if not checkpoint_lista(success, message2, "DescartarNotificacionesLicenciasLista"):
        return
    print('DescartarNotificacionesLicenciasLista: OK')
    LicenciasParaNotificar = LicenciasFiltradas.get("LicenciasNoDescartadas", [])
    LicenciasDescartadasLista  = LicenciasFiltradas.get("LicenciasDescartadas", [])
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Descartar Notificaciones Licencias Lista (Lista2):",
        "message": message2
    })

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, LicenciasFiltradas = ValidarEmailComercialesLista(
        CamposEmail,
        LicenciasParaNotificar,
        UsuariosTenant
    )
    if not checkpoint_lista(success, message2, "ValidarEmailComercialesLista"):
        return
    print('ValidarEmailComerciales: OK')
    ListaNotificarLicenciasConEmail = LicenciasFiltradas["ListaNotificarLicenciasConEmail"]
    ListaNotificarLicenciasSinEmail = LicenciasFiltradas["ListaNotificarLicenciasSinEmail"]
    ListaNotificarLincenciasComercialInactivos = LicenciasFiltradas["ListaNotificarLincenciasComercialInactivos"]
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Validar Email Comerciales Lista (Lista2):",
        "message": message2
    })

    # Clasifica (sin enviar ni sellar) las licencias en ventanas de vencimiento,
    # agrupadas por comercial. El envío real y el sellado ocurren más abajo, una
    # sola vez para ambas pipelines, en EnviarNotificacionesCombinadas.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, PlanLista = PrepararNotificacionesLista(
        CamposEmail,
        CAMPOSDEVENCIMIENTO["Lista2"],
        CAMPOSDEALERTAONOTIFICACION["Lista2"],
        ListaNotificarLicenciasConEmail,
        ListaNotificarLicenciasSinEmail,
        ListaNotificarLincenciasComercialInactivos,
        LicenciasSinFechaValida
    )
    if not checkpoint_lista(success, message2, "PrepararNotificacionesLista"):
        return
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Preparar Notificaciones Lista:",
        "message": message2
    })

    # ===========================================================================================
    # ENVÍO COMBINADO: ambas pipelines ya clasificaron sus licencias (PlanExcel/PlanLista).
    # A partir de aquí se envía UN SOLO correo por tipo (Próximos Vencimientos, Sin
    # destinatario, Seguimiento consolidado, Seguimiento de ejecución), combinando el
    # contenido de las dos fuentes, en vez de uno por pipeline como antes.

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, ResultadoEnvio = EnviarNotificacionesCombinadas(
        CONFIG, DESTINATARIOS, PlanExcel, PlanLista
    )
    if not checkpoint_final(success, message, "EnviarNotificacionesCombinadas"):
        return
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Enviar Notificaciones Combinadas:",
        "message": message
    })

    # Ahora sí, con el sellado de "Notificaciones" ya aplicado por EnviarNotificacionesCombinadas,
    # se escribe de vuelta al Excel de SharePoint.
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message = ActualizarLicenciasExcelEnSharepoint(
        CONFIG,
        ListaDeLicenciasSharePoint,
        DATOS_EXCEL_SHAREPOINT
    )
    if not checkpoint_final(success, message, "ActualizarLicenciasExcelEnSharepoint"):
        return
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Actualizar Licencias Excel En Sharepoint:",
        "message": message
    })

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message = EnviarCorreoSeguimientoLogsCombinado(CONFIG, MSGLOGS, MSGLOGS2, DESTINATARIOS)
    if not checkpoint_final(success, message, "EnviarCorreoSeguimientoLogsCombinado"):
        return
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Enviar Correo Seguimiento Logs Combinado:",
        "message": message
    })

    ############################################################
    # Fin de la ejecución
    logging.info(json.dumps({
        "success": True,
        "messageLogs": {"Excel": MSGLOGS, "Lista": MSGLOGS2},
    }, ensure_ascii=False, indent=2))
    ############################################################
