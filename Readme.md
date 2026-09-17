# Azure Function - Licencias Fortinet

## 📌 Descripción
Esta Azure Function en **Python** permite:
1. Leer una lista de **seriales** almacenados en una lista de **SharePoint Online**.
2. Consultar el estado de las licencias en la plataforma de **Fortinet** mediante **web scraping**.
3. Notificar por correo electrónico las licencias próximas a vencerse.

El objetivo es automatizar la verificación de licencias y alertar oportunamente para evitar interrupciones en los servicios.

---

## ⚙️ Arquitectura
- **Azure Functions**:
  - **HTTP Trigger**: para pruebas en tiempo real y depuración manual.
  - **Timer Trigger**: para ejecución automática y periódica sin intervención humana.
- **SharePoint Online**: fuente de datos con los seriales de licencias.
- **Fortinet Platform**: sitio web donde se valida el vencimiento de cada licencia.
- **Servicio de correo**: envío de notificaciones a los responsables.

---

## ⏱️ Ejecución programada
La función puede configurarse con un **Timer Trigger** usando expresiones CRON.  
Ejemplo: ejecutar cada día a las 8:00 AM:
```json
{
  "schedule": "0 0 7 * * *"
}
