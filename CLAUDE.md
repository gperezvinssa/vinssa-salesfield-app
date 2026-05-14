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

## Data sources: present and future

La app actualmente lee datos de oportunidades desde `Oportunidades.xlsx` en SharePoint vía Microsoft Graph API. Este es un arreglo temporal mientras el certificado SSL de SAP B1 Service Layer (`vinssa.vertrou.cloud:50000`) está pendiente de instalación por Vertrou.

Cuando SAP esté disponible, la función `cargarOportunidadesAsesor()` en `sap.js` se reescribe internamente para llamar al endpoint `SalesOpportunities` de SAP Service Layer directamente. La firma externa de la función queda igual, así que el resto del frontend (`auth.js`, `app.js`) no requiere cambios. Es un cambio quirúrgico, no un rewrite.

Mismo principio aplica para otras lecturas futuras (clientes, productos, ventas, etc.): la app define funciones de lectura abstraídas; las implementaciones internas cambian de xlsx a SAP cuando llegue el momento.

El campo `OpportunidadID` guardado en la SharePoint List `Visitas Field App` es la pieza clave que permite a Power Automate (hoy) o a la app directamente (cuando SAP conecte) actualizar la oportunidad correcta en SAP B1 cuando se sincronice un cierre o avance de etapa.

---

## Conventions — always do these

- **Comments and UI strings in Spanish.** Code identifiers in English are fine; user-facing text and code comments stay Spanish.
- **All money in MXP for the dashboard.** USD-origin lines convert via `TotalFrgn` with fallback to `ORTT.Rate`. See "Gotchas — SAP queries" below.
- **All dates as text** in xlsx files uploaded to SharePoint. Format `dd-MM-yyyy`. Format `dd/MM/yyyy HH:mm` for human-readable display (use `convertTimeZone(..., 'UTC', 'Central Standard Time', ...)`).
- **Asesor names match SAP exactly.** If a Presupuesto/dashboard name doesn't match SAP, add an entry to `DASHBOARD_CONFIG.mapaAlias` (Presupuesto → SAP) rather than renaming the SAP-side string.
- **String matching against SAP-origin data uses exact normalized match.** Convención: `uppercase` + NFD strip de acentos + `trim`. Implementado como `dashNormNombre` (dashboard.js), `_normNombre` (sap.js), `normCliente` (app.js). Usar para matchear nombres de asesor y de cliente tipeados por el usuario contra valores que vienen de SAP/xlsx. No usar `includes` ni match tolerante salvo que se documente como concesión deliberada de UX.
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

### Etapas vs Status de oportunidad
En SAP B1, **etapa** (`StepLast` → tabla `OOST`) y **status de cierre** (`Status`: `O`=Abierta, otros valores=Cerrada Ganada/Perdida) son campos **separados**. El catálogo de etapas en `config.js` debe contener **solo etapas progresivas** (Contacto Inicial → Factura), **NO incluir 'Perdido' ni 'Ganado'** como etapa. Si el formulario necesita capturar cierre como ganada/perdida, debe ser un campo aparte del selector de etapa. Mantener esta separación previene corromper datos al sincronizar con SAP.

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

### Identidad del asesor (mapeo MSAL → SAP)
La app obtiene el email del usuario logueado vía MSAL (`account.username`, equivalente a `preferred_username` del idtoken). Ese email (lowercase) se mapea a un nombre de asesor SAP via la constante `EMAIL_A_ASESOR` en `config.js`. El nombre de asesor SAP debe coincidir EXACTAMENTE con la columna `Asesor` en `Oportunidades.xlsx` (que viene de `OSLP.SlpName` en SAP B1, típicamente en MAYÚSCULAS sin acentos). Cuando un asesor nuevo se sume al piloto, agregar entrada al mapeo. Cuando un asesor cambia su email (raro), actualizar la entrada.

Si el usuario logueado no está en el mapeo, `STATE.asesorSAP` queda `null` y los tres sub-flujos de Actualizar Oportunidad (Avanzó / Se ganó / Se perdió) muestran un mensaje de "función en piloto" dentro del form. Los flujos de Nueva Visita y Demo Realizada siguen funcionando normalmente para todos los usuarios.

### UX del combobox de cliente/oportunidad (Actualizar Oportunidad)

Decisiones que se tomaron durante el piloto, no obvias del código:

- **No hacer auto-focus al input de cliente** al entrar a Avanzó/Ganada/Perdida. En celular, auto-focus dispara el teclado virtual inmediatamente, tapando el form antes de que el asesor vea qué tiene que llenar. El asesor tapea cuando esté listo.
- **Auto-selección de la oportunidad cuando el cliente tiene UNA sola opp activa.** Reduce taps: el asesor selecciona cliente → si solo tiene una opp en SAP, esa queda pre-seleccionada con banner de contexto + monto pre-llenado (en Ganada). Implementado en `cbCommitCliente` (app.js): `if (ops.length === 1) STATE.oportunidadSeleccionada = ops[0]`.
- **Click fuera del combobox cierra la lista sin commit.** Listener global en `document.click` cierra cualquier combobox abierto cuyo click target no esté dentro de `.combobox`. Cerrar ≠ comittear: el commit (incluyendo fallback de cliente nuevo) ocurre vía `cbBlur` cuando el input pierde focus naturalmente, no por outside-click.

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
- **UI changes need mobile validation before commit.** Field App es mobile-first; los flujos críticos viven en celular. Cambios de UI (forms, combobox, layout) que se ven OK en desktop pueden fallar en celular por: zoom de iOS Safari (font-size < 16px en inputs), teclado virtual tapando viewport, scrollbars duplicados (choque de inline `display:block` vs CSS `display:flex`), o tap targets muy chicos. No commit hasta probar en celular real con `Ctrl+Shift+R` para invalidar service worker.

---

## Pending work — context for decisions

Don't propose work already in motion or already decided against:

- Vertrou SSL cert install (`vinssa.vertrou.cloud:50000`) — **blocker for direct SAP**, requested, awaiting response. Don't propose workarounds that require Premium connectors.
- Power Automate flow to sync SharePoint List → SAP — **created, paused on SSL cert.** Don't recreate it.
- TEST bar (Dashboard) — **stays visible** until every role (Director, Gerente per división, Líder, Asesor) has been validated end-to-end with real users on real data. Proactively removing or hiding it is **out of scope** until that validation is complete — don't propose it, don't gate it behind a flag, don't "clean it up" as part of an unrelated change. When the time comes, removal will be its own explicit task.
- Add Presupuesto for Suministros and Servicios — currently shows $0 because data missing, not a code bug.
- Dir·Todos totals validation — under review.
- Clientes en riesgo currently 0 — pending threshold/range review.
- Drill-down asesor from Gerente view — placeholder `alert()` in place, needs detail view.
- Racha de actividad — hardcoded; will eventually pull from SharePoint List "Visitas Field App".
- Replicate Field App to asesores piloto Rafael Moyeda and Néstor Carranza — pending.
- **OVERRIDE temporal de `gperez@vinssa.com → KIMBERLY PORTILLO`** en `EMAIL_A_ASESOR` (config.js). Eliminar esta entrada cuando se expanda el piloto a más asesores. La entrada existe solo para que el Director (Gerardo) pueda probar el flujo end-to-end durante la fase de pulido sin requerir login con cuenta de asesor.
- **Regenerar `Oportunidades.xlsx` con columna `DocCurrency`** cuando sea oportuno. Hoy la app asume MXP por defecto para pre-llenado de monto en cierres ganados. No es bloqueador para el piloto pero será importante cuando haya cierres en USD.
- **Match exacto normalizado de nombres de cliente** (uppercase + sin acentos + trim) es la convención actual para matchear cliente tipeado por asesor contra cliente en `Oportunidades.xlsx`. Si los asesores reportan fricción frecuente con esto, migrar a match tolerante (contains).
- **Búsqueda dentro del dropdown de oportunidades**: actualmente es dropdown plano. Cuando un asesor con muchas oportunidades reporte fricción al hacer scroll, agregar un campo de filtro por nombre/marca arriba del dropdown.
- **Flujo "actuar como asesor" para directores y gerentes**: hoy la app está diseñada para que cada usuario MSAL mapee 1:1 a un asesor SAP. Los directores y gerentes pueden necesitar cerrar oportunidades en nombre de un asesor (cuando él delega, está fuera, o renuncia). Diseño futuro: dropdown adicional "Actuando como: [asesor]" visible solo para roles Director/Gerente. El cierre queda auditado como "cerrado por X en nombre de Y".
- **Reasignación de oportunidades de asesores que salen**: cuando un asesor deja Vinssa, sus oportunidades activas en SAP quedan huérfanas si no se reasignan. Hoy la reasignación se hace manualmente en SAP. Futuro: vista para Gerentes que permita reasignar oportunidades de un asesor saliente a otro activo, con un solo flujo. Esto cierra el ciclo de gestión de carga por asesor.
- **Medir uso del fallback de texto libre**: cada vez que un asesor cierra/avanza una oportunidad usando el campo de texto libre (porque no había oportunidad en SAP), eso es señal de que la disciplina de captura de oportunidades en SAP necesita atención. Agregar reporte simple que cuente cuántos registros pasaron por fallback (`OpportunidadID` vacío y `OppNombre` con texto) vs cuántos por dropdown (`OpportunidadID` poblado), segmentado por asesor y por mes. Si el ratio de fallback es alto (>20%), abrir conversación con los gerentes sobre captura en SAP.

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
