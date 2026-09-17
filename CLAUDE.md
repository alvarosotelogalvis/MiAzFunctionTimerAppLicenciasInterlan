# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Python Azure Function (v2 programming model, single `function_app.py`) that automates license
expiration tracking for Interlan. It has one timer trigger, `timerAppLicenciasInterlan`, which runs
daily and drives two independent, sequential pipelines:

1. **Excel pipeline** — licenses stored in an Excel workbook (`PYTHON_INV-LICENCIAS CLIENTES_NORMALIZADO.xlsx`)
   in SharePoint, read/written via the Microsoft Graph workbook API (no local file parsing, no openpyxl).
2. **List pipeline** — licenses stored in two SharePoint lists (`ListaAppLicenciasInterlan` = "Lista1",
   `Gestion de Licencias` = "Lista2").

Both pipelines: read the source → validate required fields → discard rows flagged "don't notify" →
(Excel only) cross-check real expiration dates by scraping the Fortinet Partner Portal →
compute which licenses are within the expiration window → resolve the responsible commercial's email →
send expiration notification emails via Microsoft Graph `sendMail` → write results back to the source →
send a run/log summary email. Both pipelines send that summary email through the same function,
`EnviarCorreoSeguimientoLogsAzureFunction`.

## Commands

```
pip install -r requirements.txt
func start
```

There is a VS Code task (`func: host start`, F5-attachable via `.vscode/launch.json`) that runs the
same thing plus `pip install` first. There are no automated tests and no lint/build step configured.

## Architecture

### Pipeline-of-modules pattern

`function_app.py` is a long orchestrator: it imports ~25 single-purpose modules from
`app/services/Modulo*.py` and calls them in a fixed sequence, threading state between steps
via plain dicts/lists. **Every module follows the same contract**:

- Named `Modulo<VerboSustantivo>.py`, exporting one function with a matching PascalCase name
  (e.g. `ModuloDescartarLicenciasDeExcel.py` → `DescartarLicenciasDeExcelPorNotificacionNo`).
- Signature returns a tuple: `(success: bool, message: str, data)` (a couple of write-back steps
  return just `(success, message)`).
- `data` is usually a dict of named buckets the caller unpacks and re-merges before the next step,
  e.g. `{"LicenciasNoDescartadas": [...], "LicenciasDescartadas": [...]}` or
  `{"ListasLicenciasFortinetConsultadas": [...], "ListasLicenciasFortinetNoConsultadas": [...]}`.
- On `success is False`, `function_app.py` immediately returns a `func.HttpResponse` with status 500
  and the message — there is no shared error-handling helper actually wired in (`ManejarErrorCritico`
  exists at the top of the file but its call site is commented out).

When adding a step to either pipeline, match this shape rather than introducing a different return
convention, since every downstream unpack in `function_app.py` assumes it.

### Generic modules parametrized by source

Because the same kind of validation applies to Excel rows, Lista1 rows, and Lista2 rows, the field
names differ per source, so several modules take a `CAMPOS_OBLIGATORIOS` / `CAMPOSDEVENCIMIENTO` /
`CAMPOSDEALERTAONOTIFICACION` dict (keyed `"Excel"`, `"Lista1"`, `"Lista2"`) and look up the right
column name for the caller's source instead of hardcoding it. These maps are defined once, near the
top of `timerAppLicenciasInterlan`, and passed down through most of the pipeline — extend them there
rather than hardcoding a new field name inside a module.

### Auth and integrations

- **Microsoft Graph, app-only (client credentials)**: used for reading/writing the SharePoint Excel
  workbook, reading/writing the SharePoint lists, listing tenant users, and sending mail. Two separate
  app registrations are used — `CONFIG` (main Graph access) and `CONFIG2` (only for resolving
  commercial users' emails in `ModuloAsignarEmailComercialesLista.py` / `ModuloValidarEmailComercialesExcel.py`).
- **Fortinet Partner Portal**: no public API — `ModuloLicenciasExcelFortinet.py` scrapes it with
  `requests.Session` + `BeautifulSoup`, replaying a login → SSO redirect → renewal-tools form chain,
  then submits one serial number at a time against `SNContractQuery.aspx` to read back the real
  expiration date. This is the only step that hits an external, non-Microsoft site, and it's
  sequential per-serial (with `time.sleep` between requests), so it's the slowest and most fragile step.

### Where configuration lives

All runtime config — tenant/client credentials, site IDs, list names, required-field lists, the
Excel file paths/sheet allowlist, and the `ENTORNO` (desarrollo/producción) flag — is defined as
literal dicts inside `timerAppLicenciasInterlan` in `function_app.py`, not in `local.settings.json`
or environment variables. `local.settings.json` only holds the Functions runtime storage/worker
settings. When changing an endpoint, list name, file path, or field mapping, look in
`function_app.py` first, not in a settings/env file.

`app/services/DestinatariosConCopia.json` holds the recipient list for the run-summary email
(`EnviarCorreoSeguimientoLogsAzureFunction` reads it directly from disk next to the module).

### `EsProduccion` / dev vs prod branching

`CONFIG["ENTORNO"]` (`"desarrollo"` or `"produccion"`) is meant to control whether critical errors
abort the timer run and email an alert (prod) or return an inspectable HTTP 500 body (dev), via
`ManejarErrorCritico`. In the current code this helper is defined but not called anywhere in the
pipeline — every step instead does its own inline `if not success: return func.HttpResponse(...)`,
which only makes sense for the HTTP-trigger/dev testing path since a timer trigger has no caller to
receive that response. Keep this in mind if a run appears to "fail silently" in production.
