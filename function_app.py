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
from app.services.ModuloNotificacionVencimientosExcel import NotificarVencimientosExcel
from app.services.ModuloActualizarLicenciasExcelEnSharepoint import ActualizarLicenciasExcelEnSharepoint
from app.services.ModuloEnviarCorreoSeguimientoLogsAzureFunction import EnviarCorreoSeguimientoLogsAzureFunction

#Licencias Lista2: Gestion de licencias
from app.services.ModuloLicenciasDeSharePointLista import LeerUnalistaDeSharepoint
from app.services.ModuloVerificarCamposObligatoriosLista import VerificarCamposObligatoriosLista
from app.services.ModuloDescartarLicenciasListasPorAlertaNoOk import DescartarLicenciasListasPorAlertaNoOk
from app.services.ModuloLeerUsuariosTenant import LeerUsuariosDelTenant
from app.services.ModuloAsignarEmailAUsuarios import AsignarEmailDeUsuariosALicencias
from app.services.ModuloVencimientoLicenciasGenerales import ValidarVencimientosLicencias
from app.services.ModuloAsignarEmailComercialesLista import AsignarEmailComercialesLista
from app.services.ModuloValidarEmailComercialesLista import ValidarEmailComercialesLista
from app.services.ModuloValidarNotificacionesLicenciasLista import ValidarNotificacionesLicenciasLista
from app.services.ModuloDescartarNotificacionesLicenciasLista import DescartarNotificacionesLicenciasLista
from app.services.ModuloNotificacionVencimientos2 import NotificarVencimientosListas
#from app.services.ModuloActualizarLicenciasListasEnSharepoint import ActualizarLicenciasListasEnSharepoint

app = func.FunctionApp()

def EsProduccion(config: dict) -> bool:
    return config.get("ENTORNO") == "produccion"

def ManejarErrorCritico(
        *,
        success: bool,
        message: str,
        funcion: str,
        config: dict,
        logs: dict
    ):
        """
        Maneja errores críticos de forma uniforme:
        - Producción (Timer): envía correo y aborta con Exception
        - Desarrollo (HTTP): retorna HttpResponse 500 con detalle
        """
        if success:
            return None  # ✅ no hay error, el flujo continúa
        # 🔴 Log siempre
        logging.error(f"[{funcion}] Error crítico: {message}")
        if EsProduccion(config):
            # 🔴 PRODUCCIÓN → proceso automático
            EnviarCorreoDeError(
                config=config,
                proceso=funcion,
                mensaje=logs
            )
            # Aborta correctamente el Timer
            raise Exception(message)
        else:
            # 🔵 DESARROLLO → HTTP Trigger
            print("⚠️ Error en entorno DESARROLLO")
            print(logs)
            return func.HttpResponse(
                json.dumps(
                    {
                        "success": False,
                        "Funcion": funcion,
                        "message": message,
                        "logs": logs
                    },
                    indent=2
                ),
                status_code=500,
                mimetype="application/json"
            )

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
    }
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
    if not success:
        logging.error(f"[LeerLicenciasExcelDeSharepoint] {message}")
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
    if not success:
        logging.error(f"[VerificarCamposObligatoriosEnExcelDeSharepoint] {message}")
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
    if not success:
        logging.error(f"[DescartarLicenciasDeExcelPorNotificacionNo] {message}")
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
    if not success:
        logging.error(f"[FiltrarLicenciasDeExcelFortinet] {message}")
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
    if not success:
        logging.error(f"[FiltrarLicenciasExcelFortinetParaConsultar] {message}")
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
    if not success:
        logging.error(f"[ValidarVencimLicenciasExcelPlatafFortinet] {message}")
        return
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
    if not success:
        logging.error(f"[ValidarVencimientosLicenciasExcel] {message}")
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
    if not success:
        logging.error(f"[ValidarNotificacionesLicenciasExcel] {message}")
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
    if not success:
        logging.error(f"[ValidarEmailComercialesExcel] {message}")
        return
    ListaNotificarLicenciasConEmail = LicenciasFiltradas.get("ListaNotificarLicenciasConEmail", [])
    ListaNotificarLicenciasSinEmail = LicenciasFiltradas.get("ListaNotificarLicenciasSinEmail", [])
    ListaNotificarLincenciasComercialInactivos = LicenciasFiltradas.get("ListaNotificarLincenciasComercialInactivos", [])
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Validar Email Comerciales Excel:",
        "message": message
    })

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = NotificarVencimientosExcel(
        CONFIG,
        CamposEmail,
        ListaNotificarLicenciasConEmail,
        ListaNotificarLicenciasSinEmail,
        ListaNotificarLincenciasComercialInactivos,
        LicenciasSinFecha,
        DESTINATARIOS
    )
    if not success:
        logging.error(f"[NotificarVencimientosExcel] {message}")
        return
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Notificar Vencimientos Excel:",
        "message": message
    })

    '''timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message, LicenciasFiltradas = NotificarVencimientosExcel(
        CONFIG,
        CamposEmail,
        ListaNotificarLicenciasConEmail,
        ListaNotificarLicenciasSinEmail,
        ListaNotificarLincenciasComercialInactivos
    )
    if not success:
        body = {"success": False, "Funcion": "NotificarVencimientosExcel", "message": message}
        return func.HttpResponse(
            json.dumps(body),
            status_code=500,
            mimetype="application/json"
        )
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Notificar Vencimientos Excel:",
        "message": message
    })'''

    ListaDeLicenciasSharePoint.extend(LicenciasFiltradas["ListaNotificarLicenciasConEmail"])
    ListaDeLicenciasSharePoint.extend(LicenciasFiltradas["ListaNotificarLicenciasSinEmail"])
    ListaDeLicenciasSharePoint.extend(LicenciasFiltradas["ListaNotificarLincenciasComercialInactivos"])
    ListaDeLicenciasSharePoint.extend(LicenciasDescartadas)

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message = ActualizarLicenciasExcelEnSharepoint(
        CONFIG,
        ListaDeLicenciasSharePoint,
        DATOS_EXCEL_SHAREPOINT
    )
    if not success:
        logging.error(f"[ActualizarLicenciasExcelEnSharepoint] {message}")
        return
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Actualizar Licencias Excel En Sharepoint:",
        "message": message
    })
    
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message = EnviarCorreoSeguimientoLogsAzureFunction(
        CONFIG,
        MSGLOGS,
        DESTINATARIOS
    )
    if not success:
        logging.error(f"[EnviarCorreoSeguimientoLogsAzureFunction] {message}")
        return
    MSGLOGS["Eventos"].append({
        "time": timestamp,
        "step": "Enviar Correo Seguimiento Logs Azure Function:",
        "message": message
    })

    ############################################################
    # Mostrar el contenido retornado por la función
    '''return func.HttpResponse(
        json.dumps({
            "success": True,
            "messageLogs": MSGLOGS,
            #"LicenciasFiltradas": LicenciasFiltradas,
            #"ListaLicenciasExcelFortinetParaConsultar": ListaLicenciasExcelFortinetParaConsultar
        }, ensure_ascii=False, indent=2),
        status_code=200,
        mimetype="application/json"
    )'''
    ############################################################

    # ===========================================================================================
    # LISTA DE LICENCIAS COMO LISTA (LISTA2: GESTION DE LICENCIAS) EN SHAREPOINT.
    
    #Retorna una lista de licencias leidas del sharepoint.
    CONFIG["SITE_ID"] = LISTAS["SITE_ID_2"]
    CONFIG["LIST_DISPLAY_NAME"] = LISTAS["LIST_DISPLAY_NAME_2"]
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, ListaDeLicenciasSharePoint = LeerUnalistaDeSharepoint(CONFIG)
    if not success:
        logging.error(f"[LeerUnalistaDeSharepoint] {message2}")
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
    success, message2, ListaDeLicenciasSharePoint = VerificarCamposObligatoriosLista(
        ListaDeLicenciasSharePoint, CAMPOS_OBLIGATORIOS["Lista2"]
    )
    if not success:
        logging.error(f"[VerificarCamposObligatoriosLista] {message2}")
        return
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Verificar Campos Obligatorios Lista (Lista2):",
        "message": message2
    })

    #Aqui se descartan pero no se devuelven las descartadas por lo que se "Pierden"
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, LicenciasParaNotificar = DescartarLicenciasListasPorAlertaNoOk(
        ListaDeLicenciasSharePoint, CAMPOSDEALERTAONOTIFICACION["Lista2"]
    )
    if not success:
        logging.error(f"[DescartarLicenciasListasPorAlertaNoOk] {message2}")
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
    if not success:
        logging.error(f"[ValidarVencimientosLicencias] {message2}")
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
    if not success:
        logging.error(f"[LeerUsuariosDelTenant] {message2}")
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
    if not success:
        logging.error(f"[AsignarEmailComercialesLista] {message2}")
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
    if not success:
        logging.error(f"[ValidarNotificacionesLicenciasLista] {message2}")
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
    if not success:
        logging.error(f"[DescartarNotificacionesLicenciasLista] {message2}")
        return
    print('DescartarNotificacionesLicenciasLista: OK')
    LicenciasParaNotificar = LicenciasFiltradas.get("LicenciasNoDescartadas", [])
    LicenciasDescartadas  = LicenciasFiltradas.get("LicenciasDescartadas", [])
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
    if not success:
        logging.error(f"[ValidarEmailComercialesLista] {message2}")
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
    
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2, LicenciasFiltradas = NotificarVencimientosListas(
        CONFIG,
        CamposEmail,
        CAMPOSDEVENCIMIENTO["Lista2"],
        CAMPOSDEALERTAONOTIFICACION["Lista2"],
        ListaNotificarLicenciasConEmail,
        ListaNotificarLicenciasSinEmail,
        ListaNotificarLincenciasComercialInactivos,
        LicenciasSinFechaValida,
        DESTINATARIOS
    )
    if not success:
        logging.error(f"[NotificarVencimientosListas] {message2}")
        return
    #print('NotificarVencimientosExcel: OK')
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Notificar Vencimientos Excel:",
        "message": message2
    })
    
    '''timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2 = ActualizarLicenciasListasEnSharepoint(
        CONFIG,
        LicenciasParaNotificar,
        DATOS_EXCEL_SHAREPOINT_LISTAS,
        #DATOS_EXCEL_SHAREPOINT,
    )
    if not success:
        #return func.HttpResponse(f"Proceso exitoso: {message2}, {CONFIG}", status_code=200, mimetype="text/plain")
        body = {"success": False, "Funcion": "ActualizarLicenciasListasEnSharepoint", "message": message2}
        return func.HttpResponse(
            json.dumps(body),
            status_code=500,
            mimetype="application/json"
        )
    #print('NotificarVencimientosExcel: OK')
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Actualizar Licencias Listas En Sharepoint:",
        "message": message2
    })'''

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success, message2 = EnviarCorreoSeguimientoLogsAzureFunction(
        CONFIG,
        MSGLOGS2,
        DESTINATARIOS
    )
    if not success:
        logging.error(f"[EnviarCorreoSeguimientoLogsAzureFunction] {message2}")
        return
    #print("EnviarCorreoSeguimientoLogsAzureFunction -> OK:")
    MSGLOGS2["Eventos"].append({
        "time": timestamp,
        "step": "Enviar Correo Seguimiento Logs Azure Function:",
        "message": message2
    })
    
    ############################################################
    # Fin de la ejecución
    logging.info(json.dumps({
        "success": True,
        "messageLogs": MSGLOGS2,
    }, ensure_ascii=False, indent=2))
    ############################################################

    # ===========================================================================================