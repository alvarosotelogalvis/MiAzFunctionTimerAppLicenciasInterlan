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
  app registrations are used — `CONFIG` (main Graph access) and `CONFIG2` (Graph auth only for
  `ModuloValidarEmailComercialesExcel.py`, which looks up tenant users to validate a commercial's email
  in the Excel pipeline). `ModuloAsignarEmailComercialesLista.py` — the List-pipeline equivalent —
  doesn't call Graph at all; it looks up a static `{id: email}` table built from the
  `COMERCIALES_EMAILS_JSON` env var (see "Where configuration lives" below).
- **Fortinet Partner Portal**: no public API — `ModuloLicenciasExcelFortinet.py` scrapes it with
  `requests.Session` + `BeautifulSoup`, replaying a login → SSO redirect → renewal-tools form chain,
  then submits one serial number at a time against `SNContractQuery.aspx` to read back the real
  expiration date. This is the only step that hits an external, non-Microsoft site, and it's
  sequential per-serial (with `time.sleep` between requests), so it's the slowest and most fragile step.

### Where configuration lives

Secrets and per-environment values are read from environment variables via `os.environ.get(...)`
inside `timerAppLicenciasInterlan` — populated from `local.settings.json`'s `Values` block locally,
and from the Function App's Application Settings in Azure. `local.settings.json` is gitignored and
excluded from deployment via `.funcignore`, so it never leaves the machine; `local.settings.json.example`
(committed, placeholder values only) documents every key, each with a `_comment_<KEY>` sibling entry
explaining what it's for (JSON has no native comment syntax, so this is the convention used here).

Env vars consumed:
- `FORTINET_USER` / `FORTINET_PASSWORD` — Fortinet Partner Portal login.
- `APP1_TENANT_ID` / `APP1_CLIENT_ID` / `APP1_CLIENT_SECRET` — the "licencias" app registration (`CONFIG`).
- `APP2_TENANT_ID` / `APP2_CLIENT_ID` / `APP2_CLIENT_SECRET` — the "comerciales" app registration (`CONFIG2`).
- `MAIL_SENDER` — mailbox used for every Graph `sendMail` call.
- `MAIL_DESTINATARIOS` / `MAIL_SEGUIMIENTO` / `MAIL_SIN_DESTINATARIOS` — comma-separated recipient
  lists, parsed in `function_app.py` into a `DESTINATARIOS` dict and passed down to
  `NotificarVencimientosExcel`, `NotificarVencimientosListas`, and `EnviarCorreoSeguimientoLogsAzureFunction`.
- `COMERCIALES_EMAILS_JSON` — a JSON string `{"usuarios":[{"id":<id>,"email":"..."}]}`, parsed into
  `COMERCIALES_EMAILS` and passed to `AsignarEmailComercialesLista` to resolve each licence's commercial
  email in the List pipeline.

Everything else non-secret — site IDs, list names, required-field lists, the Excel file paths/sheet
allowlist, and the `ENTORNO` (desarrollo/producción) flag — is still a literal dict inside
`timerAppLicenciasInterlan` in `function_app.py`. When changing an endpoint, list name, file path, or
field mapping, look in `function_app.py` first; when changing a credential or a mail sender/recipient,
look in `local.settings.json` (local) or Application Settings (Azure) instead.

`app/services/DestinatariosConCopia.json` and `app/services/usuariosComerciales_emails*.json` used to
hold this recipient/commercial data on disk, with real emails committed to git — they were removed in
favor of the env vars above. Don't recreate them; extend `MAIL_*` / `COMERCIALES_EMAILS_JSON` instead.

### `EsProduccion` / dev vs prod branching

`CONFIG["ENTORNO"]` (`"desarrollo"` or `"produccion"`) is meant to control whether critical errors
abort the timer run and email an alert (prod) or return an inspectable HTTP 500 body (dev), via
`ManejarErrorCritico`. In the current code this helper is defined but not called anywhere in the
pipeline — every step instead does its own inline `if not success: return func.HttpResponse(...)`,
which only makes sense for the HTTP-trigger/dev testing path since a timer trigger has no caller to
receive that response. Keep this in mind if a run appears to "fail silently" in production.
