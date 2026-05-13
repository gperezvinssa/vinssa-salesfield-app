# CLAUDE.md — Vinssa Sales Field App + Sales Dashboard

This file is read by Claude Code at the start of every session. Keep it tight, specific, and current.

---

## What this repo is

A single GitHub Pages PWA at `https://gperezvinssa.github.io/vinssa-salesfield-app/` that contains **two related products** sharing one codebase:

1. **Field App** — mobile-first form used by asesores to register field activities (visits, demos, leads, opportunity updates). Captures GPS, requires check-in, writes to a SharePoint List.
2. **Sales Dashboard** — desktop-first reporting view used by Director, Gerentes, Líderes, and Asesores to see ventas, pipeline, oportunidades, presupuesto vs cumplimiento. Reads from xlsx files in SharePoint.

They share auth (MSAL/Azure), config catalogs (`js/config.js`), and the SharePoint backend. Treat changes that touch shared files as cross-cutting.

**Owner:** Gerardo Pérez, Director de Trazabilidad y Automatización, Vinssa (Versatilidad Industrial de Saltillo SA).

**End users:** Asesores comerciales, líderes de línea, gerentes de división, dirección. All Vinssa employees on tenant `versatilidadsaltillo.sharepoint.com`. UI language is Spanish.

---

## File map

```
vinssa-salesfield-app/
├── index.html          # Both apps share this entry; nav switches between screens
├── manifest.json       # PWA install metadata
├── icon.png
├── css/
│   └── styles.css      # All styles for both apps
└── js/
    ├── config.js       # Shared catalogs: divisiones, marcas, etapas, líderes, mapaAlias, mapaGrupos, asesores*
    ├── auth.js         # MSAL/Microsoft login (used by both apps)
    ├── geo.js          # Field App — 3-layer geolocation + check-in/checkout
    ├── sap.js          # Field App — writes to SharePoint List via Graph API
    ├── app.js          # Field App — form logic, validation, submit
    └── dashboard.js    # Sales Dashboard — DASHBOARD_CONFIG, role logic, render
```

When in doubt about what a function does, read the file. Don't infer from names — Vinssa-specific business logic is dense and hand-tuned.

---

## Stack

- **Frontend:** Vanilla JS, no framework. No build step. Edit files directly, commit, push.
- **Auth:** Microsoft MSAL (SPA). Tenant: `7a272c1a-ee40-4b22-a187-2656ea44b4c4`. Client ID: `45d6f369-f789-473b-8970-d9b25ff3225c`. Only Vinssa tenant users.
- **Backend (current):** SharePoint List "Visitas Field App" + xlsx files in SharePoint Documents root, accessed via Microsoft Graph API (`Sites.ReadWrite.All`).
- **Backend (planned):** SAP Business One Service Layer at `https://vinssa.vertrou.cloud:50000/b1s/v1`. Currently blocked by self-signed SSL cert at Vertrou. Direct browser → SAP doesn't work anyway (CORS) — Power Automate is the intended bridge.
- **Power Automate flow:** `Field App - Sincronizar registro con SAP` exists but paused on SSL cert.
- **Deploy:** GitHub Pages, branch `main`. Reflects in ~1 min. Force refresh with Ctrl+Shift+R.
- **Timezone:** America/Monterrey (Central Standard Time).

---

## Conventions — always do these

- **Comments and UI strings in Spanish.** Code identifiers in English are fine; user-facing text and code comments stay Spanish.
- **All money in MXP for the dashboard.** USD-origin lines convert via `TotalFrgn` with fallback to `ORTT.Rate`. See "Gotchas — SAP queries" below.
- **All dates as text** in xlsx files uploaded to SharePoint. Format `dd-MM-yyyy`. Format `dd/MM/yyyy HH:mm` for human-readable display (use `convertTimeZone(..., 'UTC', 'Central Standard Time', ...)`).
- **Asesor names match SAP exactly.** If a Presupuesto/dashboard name doesn't match SAP, add an entry to `DASHBOARD_CONFIG.mapaAlias` (Presupuesto → SAP) rather than renaming the SAP-side string.
- **Localstorage as fallback.** When writes to SharePoint can fail, write to localStorage first. See `sap.js` for the pattern.
- **Etapas reference `DASHBOARD_CONFIG.etapas`.** Six canonical stages with % and colors. Don't hardcode % or colors elsewhere.
- **Check-in is mandatory** in the Field App before any of the four registro buttons enable. Don't add a registro flow that bypasses this — it breaks the GPS-per-activity guarantee.
- **mapaGrupos and visible divisions must stay in sync.** If you add a SAP grupo to `mapaGrupos`, make sure the resulting división is one of: Trazabilidad, Visión, Robótica, Suministros, Servicios.

---

## Gotchas — never break these

### SharePoint List column "Competidor"
The display name is "Competidor" (Spanish) but the **internal name is "Competitor"** (English). `sap.js` writes to `Competitor:`. If you rename the column in SharePoint, update `sap.js`. If you "fix" the typo in `sap.js`, writes will fail silently.

### SAP queries — Ventas
- **Use `T2.TotalFrgn`, not `T2.LineTotal` directly.** MXP invoices have `DocRate=1`, which inflates USD totals if you divide by `DocRate`. The canonical pattern is `CASE WHEN T2.TotalFrgn <> 0 THEN T2.TotalFrgn WHEN T4b.Rate > 1 THEN T2.LineTotal / T4b.Rate ELSE T2.LineTotal / 17.5 END`.
- **UNION ALL with ORIN (notas de crédito) is required.** Without it, totals are inflated because credits aren't subtracted. ORIN rows use the same shape but with **negated** Total.
- **Canonical queries live in `vinssa-dashboard-status.md`** (the project doc). Treat that file as the source of truth for SAP query shape.

### Date columns in xlsx → SharePoint
If a date column is typed as Date in xlsx, Graph API serializes it to an Excel serial number (e.g. `45678`) and the dashboard parser breaks. Always format Fecha-style columns as **Text** before uploading. The parser expects `dd-MM-yyyy` strings.

### Graph API cache when refreshing data
If you upload a new xlsx and the dashboard shows stale data, the issue is the in-memory cache, not the data. Clear it from devtools console with:
```js
DASH_STATE.token = null;
DASH_STATE.driveId = null;
dashCargado = false;
navSwitch('screen-dashboard', null);
```
If you find yourself wanting to "fix" this with a cache-bust on every load, don't — the cache is intentional for cost/perf. Document the manual clear in any new operator-facing doc.

### Power Automate plan limits
Vinssa's plan is basic, no Premium. No HTTP trigger, no premium connectors. The SharePoint List → Power Automate → SAP pattern exists specifically because of this constraint. Don't propose architectures requiring Premium without flagging the cost.

### SAP Service Layer is not reachable from the browser
CORS blocks direct calls. The architecture uses Power Automate as the bridge. Don't propose `fetch('https://vinssa.vertrou.cloud:50000/...')` from `sap.js`.

### Etapa order is not enforced by SAP
SAP B1 accepts any stage transition (including backwards) and stores each as a separate line. The intended UX is: warn when new `% < current %`, but **do not block**. Don't add validation that prevents the save.

### Brand keywords (Field App lead detection — adjacent flow, but keep in mind)
Brand keyword detection elsewhere in Gerardo's automation suite excludes "telesis" as a keyword because it appears in every Telesis lead email signature and creates false positives. If you ever consolidate brand catalogs here, preserve that distinction.

---

## Domain knowledge that affects code decisions

### Roles and what they see (Dashboard)
- **Director, division "Todos"** — sees all asesores across all divisiones.
- **Gerente, division "Trazabilidad"** — sees Trazabilidad + Visión + Robótica (these three are sub-divisions of Trazabilidad in business terms, even though `mapaGrupos` distinguishes them).
- **Gerente, division "Suministros"** or **"Servicios"** — sees only that division.
- **Líder** — same scope as Gerente of the corresponding division.
- **Asesor** — only own numbers, regardless of division.

Roles live in `Lista Roles Dashboard.xlsx` in SharePoint Documents root, with columns `Email`, `Rol`, `Division`.

### Asesores activos en 2026
Three role-scoped lists in `DASHBOARD_CONFIG`:
- `asesoresSupply` — Suministros (~22 active)
- `asesoresServicio` — Servicios (~2 active)
- `asesoresAtencion` — Atención al Cliente interno (visible only to Dir·Todos)

When an asesor joins or leaves, update the appropriate list. If their name differs in Presupuesto vs SAP, also add to `mapaAlias`.

### Líderes de línea (Trazabilidad)
- Antonio Martinez — Marcaje, Visión
- Jose Juan Aguillon — Robótica
- Aldo Almaguer — Herramientas
- Jonathan Roche — Visión, Marcaje (note: name appears in Presupuesto as "JONATHAN ROCHE", in SAP as "JONATHAN ROCHE TOR" — already aliased)

### SAP grupo → División mapping (`mapaGrupos`)
- Identificacion, Herramienta de Ensamble → **Trazabilidad**
- Vision → **Visión**
- Robotica, Automatizacion → **Robótica**
- Quimicos, Abrasivos, Artículos, Otros Suministros, Soldadura y Corte, Marcadores, Seguridad → **Suministros**
- Servicio en Campo, Servicio Taller → **Servicios**
- Vending Machines → **ignored** (`null`)

### Etapas (canonical, with %)
Contacto Inicial 10%, Cotización 25%, Pruebas/Demostración 60%, Negociación 80%, Trámite con Compras 90%, Factura 95%. Colors defined once in `DASHBOARD_CONFIG.etapas` — reference them, don't redefine.

---

## Common tasks — where to edit

| Task | File | Section |
|------|------|---------|
| Add/remove an asesor (any division) | `js/config.js` and/or `js/dashboard.js` | `DASHBOARD_CONFIG.asesores*` |
| Asesor name mismatch Presupuesto ↔ SAP | `js/dashboard.js` | `DASHBOARD_CONFIG.mapaAlias` |
| New SAP grupo de artículos | `js/dashboard.js` | `mapaGrupos` |
| New competidor | `js/config.js` | `competidores` array |
| New marca for a línea | `js/config.js` | `divisiones[*].lineas[*].marcas` |
| New líder de línea | `js/config.js` | `lideres` |
| Change check-in proximity warning threshold | `js/geo.js` | search for 500 (meters) |
| Update SharePoint write logic | `js/sap.js` | — |
| Change role-based visibility | `js/dashboard.js` | `aplicarRol()` |

---

## Testing & validation

- **Local testing:** Just open `index.html` from `file://` for layout work, but MSAL auth requires `https://` so do auth flows directly on the deployed Pages URL.
- **Dashboard role testing:** From devtools console, call `dashTestMode('director', 'Todos')`, `dashTestMode('gerente', 'Trazabilidad')`, etc. This is what the visible TEST bar does (pending removal once all roles are validated with real users).
- **Field App test users:** `gperez@vinssa.com` (Gerardo) and `vinssabi@vinssa.com` (service account).
- **After any SAP query change:** Re-export to xlsx, format Fecha as text, upload to SharePoint, clear the Graph cache (snippet above), reload.
- **Spot-check numbers:** Validate against SAP directly before declaring a change correct. The current Dir·Todos numbers are still under validation.

---

## Pending work — context for decisions

Don't propose work already in motion or already decided against:

- Vertrou SSL cert install (`vinssa.vertrou.cloud:50000`) — **blocker for direct SAP**, requested, awaiting response. Don't propose workarounds that require Premium connectors.
- Power Automate flow to sync SharePoint List → SAP — **created, paused on SSL cert.** Don't recreate it.
- Eliminate TEST bar — pending all roles validated with real users.
- Add Presupuesto for Suministros and Servicios — currently shows $0 because data missing, not a code bug.
- Dir·Todos totals validation — under review.
- Clientes en riesgo currently 0 — pending threshold/range review.
- Drill-down asesor from Gerente view — placeholder `alert()` in place, needs detail view.
- Racha de actividad — hardcoded; will eventually pull from SharePoint List "Visitas Field App".
- Replicate Field App to asesores piloto Rafael Moyeda and Néstor Carranza — pending.

---

## Out of scope for this repo

These belong to other Vinssa automation projects, not here. Don't mix them in:

- PO detection/forwarding Power Automate flow (`PO - Detect and Forward`)
- Quote/lead detection Power Automate flow (`Quote - Detect and Route`)
- Tracker Leads xlsx
- App de inteligencia de precios (separate future project)

---

## Style for code changes

- Keep changes minimal and surgical. Vanilla JS, no framework, no build step — don't introduce one.
- No new top-level dependencies unless explicitly requested.
- Preserve existing function names and structure unless refactoring is the explicit task.
- Spanish for user-facing strings, comments helpful, console messages can be terse.
- Commit messages in Spanish or English, both fine, but descriptive.
