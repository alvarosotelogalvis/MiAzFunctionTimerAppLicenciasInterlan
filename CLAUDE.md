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
**classify** (not send) which licenses go to which recipient. Only after *both* pipelines finish
classifying does a single shared tail run: one combined "send everything" step
(`ModuloEnviarNotificacionesCombinadas.py`) merges both pipelines' classified data and sends each of
the 3 notification email types (per-commercial, "sin destinatario", "seguimiento resumen") **once**
per recipient per day — instead of once per pipeline, which is what happened before this was merged.
After that, the Excel write-back runs, then a single combined run/log summary email
(`EnviarCorreoSeguimientoLogsCombinado`) covering both pipelines' logs.

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
- On `success is False`, `function_app.py` calls a small closure (`checkpoint_excel` /
  `checkpoint_lista` / `checkpoint_final`, built by `_hacer_checkpoint` near the top of
  `timerAppLicenciasInterlan`) instead of just `logging.error(...)` — see "Failure alerts" below.

When adding a step to either pipeline, match this shape rather than introducing a different return
convention, since every downstream unpack in `function_app.py` assumes it.

The two notification modules (`ModuloNotificacionVencimientosExcel.py` /
`ModuloNotificacionVencimientos2.py`) are a partial exception to "one function per module": each
exports a `Preparar*` function (`PrepararNotificacionesExcel` / `PrepararNotificacionesLista`) that
only classifies license data into a "plan" dict — it does **not** send email or stamp
`Notificaciones` itself anymore. The actual Graph `sendMail` calls and stamping live in
`ModuloEnviarNotificacionesCombinadas.py::EnviarNotificacionesCombinadas`, which both pipelines feed
into once they've both finished classifying (see "What this is" above). Token acquisition and the
retrying `sendMail` call itself are shared via `ModuloGraphMailClient.py` (`ObtenerTokenGraph`,
`EnviarCorreoGraph`, `DedupYEnvolver`) rather than duplicated per module.

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
  `EnviarNotificacionesCombinadas` and `EnviarCorreoSeguimientoLogsCombinado`.
- `MAIL_ALERTAS_FALLO` — optional comma-separated recipient list for the failure-alert email (see
  "Failure alerts" below). Falls back to `MAIL_SEGUIMIENTO` when unset, so it's not required.
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

### Failure alerts

Every one of the ~23 failure points across both pipelines (`if not success: ...`) calls a checkpoint
closure instead of just logging: `checkpoint_excel`, `checkpoint_lista`, and `checkpoint_final` (for
the 3 shared steps after both pipelines have classified — `EnviarNotificacionesCombinadas`,
`ActualizarLicenciasExcelEnSharepoint`, `EnviarCorreoSeguimientoLogsCombinado`), each built once near
the top of `timerAppLicenciasInterlan` by `_hacer_checkpoint(msglogs_list, pipeline_label)`. On
failure, the checkpoint logs the error *and* sends a real alert email via
`EnviarCorreoDeErrorPipeline` (`ModuloEnviarCorreoSeguimientoLogsAzureFunction.py`) naming which step
failed, its message, and everything logged so far for that pipeline — then returns `False` so the
call site's `if not checkpoint_x(...): return` still short-circuits the run exactly like before.

This replaces an earlier `ManejarErrorCritico`/`EsProduccion` pair that looked like it did this but
didn't: it was never called anywhere, and it called an `EnviarCorreoDeError` function that didn't
exist in the codebase at all (would have raised `NameError` if it had ever been invoked). Both were
deleted along with `CONFIG["ENTORNO"]`'s only real consumer — that key still exists (hardcoded
`"desarrollo"`) and still needs to be flipped manually for a production deploy, but nothing branches
on it for error handling anymore; every failure now emails the same way regardless of `ENTORNO`.
