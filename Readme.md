# Azure Function - Licencias Interlan

## 📌 Descripción
Esta Azure Function en **Python** (v2 programming model) automatiza el seguimiento de vencimiento
de licencias para Interlan. Un único timer trigger (`timerAppLicenciasInterlan`) corre diariamente
y ejecuta dos pipelines independientes y secuenciales:

1. **Pipeline Excel**: licencias almacenadas en un workbook de Excel en SharePoint, leído/escrito
   vía la API de Microsoft Graph (sin parsear el archivo localmente). Además cruza las fechas reales
   de vencimiento consultando el **Fortinet Partner Portal** (web scraping).
2. **Pipeline de Listas**: licencias almacenadas en dos listas de SharePoint (`ListaAppLicenciasInterlan`
   y `Gestion de Licencias`).

Ambas pipelines: leen la fuente → validan campos obligatorios → descartan filas marcadas para no
notificar → calculan qué licencias están dentro de la ventana de vencimiento → resuelven el email
del comercial responsable → envían correos de notificación vía Microsoft Graph `sendMail` → escriben
los resultados de vuelta en la fuente → envían un correo de resumen/log de la corrida.

El objetivo es automatizar la verificación de licencias y alertar oportunamente para evitar
interrupciones en los servicios.

---

## ⚙️ Arquitectura
- **Azure Functions**: un solo **Timer Trigger** (`timerAppLicenciasInterlan`), sin trigger HTTP.
- **SharePoint Online**: fuente de datos, tanto el workbook de Excel como las dos listas de licencias.
- **Fortinet Partner Portal**: sitio web (sin API pública) donde se valida el vencimiento real de
  las licencias Fortinet.
- **Microsoft Graph**: usado tanto para leer/escribir SharePoint como para enviar los correos de
  notificación (`sendMail`). Se usan dos App Registrations distintas (ver sección de configuración).

---

## 🔐 Configuración (variables de entorno)
Todas las credenciales y los remitentes/destinatarios de correo se leen desde variables de entorno,
nunca están escritas en el código:

- Local: `local.settings.json` (gitignored, no se sube ni a git ni al deploy de Azure). Usa
  `local.settings.json.example` como plantilla — cada variable tiene un comentario
  `_comment_<VARIABLE>` explicando para qué sirve.
- Azure: Application Settings del Function App (Portal → Configuration, o `az functionapp config
  appsettings set`). Deben crearse ahí manualmente con los mismos nombres antes del primer deploy.

Variables usadas: `FORTINET_USER`, `FORTINET_PASSWORD`, `APP1_TENANT_ID`, `APP1_CLIENT_ID`,
`APP1_CLIENT_SECRET`, `APP2_TENANT_ID`, `APP2_CLIENT_ID`, `APP2_CLIENT_SECRET`, `MAIL_SENDER`,
`MAIL_DESTINATARIOS`, `MAIL_SEGUIMIENTO`, `MAIL_SIN_DESTINATARIOS`, `COMERCIALES_EMAILS_JSON`.

---

## ▶️ Correr en local
```
pip install -r requirements.txt
func start
```
Con `RUN_ON_STARTUP=true` en `local.settings.json`, el timer se dispara automáticamente al arrancar
`func start` (no hace falta esperar el horario cron).

---

## ⏱️ Ejecución programada
El Timer Trigger corre con la expresión CRON `0 0 7 * * *` (todos los días a las 7:00 AM UTC).
