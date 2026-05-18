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

**Lectura de Clientes Activos vía Graph API:** agregada función `cargarClientesActivos` en `sap.js`. Cuando SAP B1 Service Layer esté disponible, esta función se reescribe internamente para leer del endpoint `BusinessPartners` filtrado por `CardType='C'` `validFor='Y'`. La firma externa de la función queda igual.

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
- **CardCode es identidad canónica del cliente, Title es legibilidad humana.** En la SharePoint List `Visitas Field App`, el campo `CardCode` (columna agregada en mayo 2026) es la llave que permite cruzar registros contra SAP B1 (entidad `OCRD`). El campo `Title` guarda el nombre del cliente como string para legibilidad en vistas de SharePoint. **Cumplen funciones distintas — no eliminar `CardCode` pensando que `Title` es suficiente.** `CardCode` queda vacío (`null`) cuando el asesor cae al fallback de texto libre (prospecto/cliente sin entrada en `Clientes Activos.xlsx`); ese vacío es señal medible de captura faltante en SAP, no un bug.
- **Propiedades en `STATE.*` mantienen la capitalización original del xlsx/SAP.** No convertir a camelCase. Ejemplo: `STATE.oportunidades[i].NumOportunidad`, `STATE.clientesActivos[i].CardCode`, `.Cliente`, `.Asesor`. Esto preserva consistencia con código existente y simplifica migración futura a SAP B1 Service Layer (donde los campos vienen en PascalCase nativamente). El draft interno del check-in (`STATE.clienteCheckinDraft`) y `GEO.checkin` sí usan camelCase porque no son lectura directa del xlsx — son estado interno de la app.

---

## Gotchas — never break these

### SharePoint List column "Competidor"
The display name is "Competidor" (Spanish) but the **internal name is "Competitor"** (English). `sap.js` writes to `Competitor:`. If you rename the column in SharePoint, update `sap.js`. If you "fix" the typo in `sap.js`, writes will fail silently.

### SharePoint List column "Acompanante" (renamed from "Lideres")
La columna que guarda los acompañantes en visita/demo se renombró en SharePoint de display name `Lideres` → `Acompanante` (mayo 2026, cuando se agregaron ingenieros de aplicación y gerencia/dirección al modelo). **Internal name sigue siendo `Lideres`** porque SharePoint preserva el internal name al renombrar display. `sap.js` escribe a `Lideres:`. Si alguien "actualiza" el código a `Acompanante:` pensando que es el nombre real, los writes fallan silenciosamente — mismo patrón que `Competidor↔Competitor`.

### SAP queries — Ventas
- **Use `T2.TotalFrgn`, not `T2.LineTotal` directly.** MXP invoices have `DocRate=1`, which inflates USD totals if you divide by `DocRate`. The canonical pattern is `CASE WHEN T2.TotalFrgn <> 0 THEN T2.TotalFrgn WHEN T4b.Rate > 1 THEN T2.LineTotal / T4b.Rate ELSE T2.LineTotal / 17.5 END`.
- **UNION ALL with ORIN (notas de crédito) is required.** Without it, totals are inflated because credits aren't subtracted. ORIN rows use the same shape but with **negated** Total.
- **Canonical queries live in `vinssa-dashboard-status.md`** (the project doc). Treat that file as the source of truth for SAP query shape.

### SAP queries — Oportunidades
Las oportunidades en SAP de Vinssa usan **System Currency = USD**. Por eso `T0.MaxSumSys` y `T0.WtSumSys` en `OOPR` ya vienen en USD nativos. **NO se requiere** `CASE WHEN` ni JOIN a `ORTT` como en `OINV`/`ORDR`. Diferencia clave: facturas (`OINV`) y OVs (`ORDR`) tienen `DocCur` explícito por documento, oportunidades (`OOPR`) usan System Currency global del sistema. Verificado 2026-05-16 con `MaxSumLoc=2000 MXP` ↔ `MaxSumSys=108.73 USD` (ratio ≈ 18.4).

### SAP queries — OOPR campos de fecha
La tabla `OOPR` tiene dos campos de fecha que **NO son intercambiables**:
- `T0.OpenDate`: cuando se creó la oportunidad. Siempre poblado.
- `T0.CloseDate`: cuando se cerró efectivamente (Won/Lost). Vacío para `Status='O'`.
- `T0.PredDate`: fecha estimada de cierre, capturada al crear la oportunidad. Poblado para la mayoría de oportunidades abiertas.

Para reportes de pipeline futuro usar SIEMPRE `PredDate`. La query inicial usaba `CloseDate` por error, generando 868 oportunidades "sin fecha proyectada" en el dashboard cuando en SAP sí tenían fecha. Corregido el 2026-05-16.

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

### Acompañantes en Visita/Demo (líderes, ingenieros, gerencia/dirección)

El form de Visita y Demo permite marcar quién acompañó al asesor (opcional, multi-select). Los emails seleccionados se concatenan y se guardan en la SharePoint List `Visitas Field App`. El rol (líder/ingeniero/gerente/director) se **infiere** al hacer reporting cruzando email contra `CONFIG.acompanantes` en `config.js`.

**Líderes de línea (Trazabilidad y Automatización):**
- Antonio Martinez — Visión
- Jose Juan Aguillon — Trazabilidad (renderiza así en el form; en `divisiones[trazabilidad].marcas` la línea correspondiente se llama `Marcaje`)
- Aldo Almaguer — Robótica
- Jonathan Roche — Herramienta de Atornillado (en `divisiones[trazabilidad].marcas` la línea se llama `Herramientas`. Nota previa: name aparece en Presupuesto como "JONATHAN ROCHE", en SAP como "JONATHAN ROCHE TOR" — ya aliased en dashboard.js)

**Ingenieros de aplicación** (apoyan a un líder en demos/visitas):
- Luis Sanchez — Robótica (reporta a Aldo Almaguer)
- Jesús Ortíz — Visión (reporta a Antonio Martinez)

**Gerencia y Dirección:**
- Fernando Barajas — Gerente, Trazabilidad y Automatización
- Gerardo Pérez — Director Comercial

Cuando alguien entra/sale de cualquiera de estos roles, actualizar `CONFIG.acompanantes` en `config.js`. El historial guardado en SharePoint sigue válido (los emails se preservan; el rol se re-infiere contra la versión actual del catálogo).

**UI:** cada tarjeta muestra solo 2 líneas — `nombre` + `linea` o `division` (lo que aplique). `rol` y `reportaA` NO se renderizan en la tarjeta pero **sí permanecen en `CONFIG.acompanantes`** (datos disponibles para features futuras: notificaciones automáticas, organigrama, filtros). El buscador de la card filtra contra `data-search` (que incluye `nombre + rol + linea + division + reportaA + email`), así que "ingeniero", "Aldo", etc siguen matcheando aunque no aparezcan visualmente. Para Director Comercial u otros sin `linea`/`division`, la tarjeta muestra solo el nombre.

### SAP grupo → División mapping (`mapaGrupos`)
- Identificacion, Herramienta de Ensamble → **Trazabilidad**
- Vision → **Visión**
- Robotica, Automatizacion → **Robótica**
- Quimicos, Abrasivos, Artículos, Otros Suministros, Soldadura y Corte, Marcadores, Seguridad → **Suministros**
- Servicio en Campo, Servicio Taller → **Servicios**
- Vending Machines → **ignored** (`null`)

### Etapas (canonical, with %)
Contacto Inicial 10%, Cotización 25%, Pruebas/Demostración 60%, Negociación 80%, Trámite con Compras 90%, Factura 95%. Colors defined once in `DASHBOARD_CONFIG.etapas` — reference them, don't redefine.

### Pipeline por Mes (dashboard)

Vista por **fecha de cierre proyectada** de las oportunidades. Vive como **sección embebida al final del tab Pipeline** ("Cierre proyectado por mes") en las 3 vistas (Asesor, Líder, Gerente/Director) — comparte el helper `_pipelineMensualHtml(asesorNorm, divisionesVisibles)` que reutiliza exactamente el mismo patrón de filtrado por rol que `_pipelineHtml` (asesorNorm + `mapaLineas`). Orden visual del tab Pipeline: Oportunidades por etapa → OVs comprometidas → Cierre proyectado por mes.

**Buckets renderizados:**
- ⚠ Vencidas — `FechaCierre < hoy_00:00`. Se omite el bucket si no hay vencidas (no estorbar). Drill-down ordena por fecha ascendente (más antiguas primero).
- Próximos 6 meses — mes actual + 5 siguientes. Cada bucket se muestra siempre, aunque esté vacío (barra gris clara con `$0 · 0 opps`) para comunicar gaps explícitamente.

**Etapas incluidas** (segmentos de la barra):
- Cotización, Pruebas / Demostración, Negociación, Trámite con Compras.

**Etapas excluidas y razón:**
- Contacto Inicial — muy temprano en el ciclo, las fechas de cierre proyectadas no son confiables a esa altura.
- Factura — operativamente ya cerrada; no representa "pipeline futuro".
- Cualquier etapa fuera del catálogo canónico de 6 — se ignora silenciosamente.

**Escala visual:** las barras usan **escala global** — todas comparan contra el mes con más monto. Permite ver de un vistazo que un mes es 4× otro. La distribución por etapa dentro de cada mes se infiere por el ancho relativo de cada segmento.

**Drill-down:** inline expand (panel debajo de la fila al tap), paralelo al UX del tab Pipeline existente. Solo un panel abierto a la vez. Tap en un segmento → opps de esa etapa en ese mes ordenadas por monto desc. Tap en el área del label del mes (no en un segmento específico) → todas las opps del mes, todas las etapas, ordenadas por monto desc. Tap en la barra de Vencidas → todas las vencidas ordenadas por fecha asc.

**Filtrado por rol:** mismo patrón que el resto del dashboard.
- Asesor → solo sus opps (filtrado por `Asesor === asesorNorm`).
- Gerente/Líder → opps de su división (filtrado por `Linea` vía `DASHBOARD_CONFIG.mapaLineas`).
- Director → todas.

**Sin FechaCierre** → la opp se excluye totalmente del tab Por Mes (no aparece ni en vencidas ni en mes alguno). **Monto = 0/null** → no contribuye al ancho de la barra pero sí aparece en el drill-down con `$0` para que el asesor la vea.

**Montos:** este tab consume `MontoEstimado` tal cual viene del xlsx (sin conversión a USD pendiente — ver pending work). Los montos pueden estar en MXP raw mientras la query de Oportunidades no se actualice.

### Identidad del asesor (mapeo MSAL → SAP)
La app obtiene el email del usuario logueado vía MSAL (`account.username`, equivalente a `preferred_username` del idtoken). Ese email (lowercase) se mapea a un nombre de asesor SAP via la constante `EMAIL_A_ASESOR` en `config.js`. El nombre de asesor SAP debe coincidir EXACTAMENTE con la columna `Asesor` en `Oportunidades.xlsx` (que viene de `OSLP.SlpName` en SAP B1, típicamente en MAYÚSCULAS sin acentos). Cuando un asesor nuevo se sume al piloto, agregar entrada al mapeo. Cuando un asesor cambia su email (raro), actualizar la entrada.

Si el usuario logueado no está en el mapeo, `STATE.asesorSAP` queda `null` y los tres sub-flujos de Actualizar Oportunidad (Avanzó / Se ganó / Se perdió) muestran un mensaje de "función en piloto" dentro del form. Los flujos de Nueva Visita y Demo Realizada siguen funcionando normalmente para todos los usuarios.

### Cliente vs Lead/Prospecto

Vinssa distingue entre clientes con CardCode formal en SAP (entidad `OCRD`) y leads/prospectos aún no formalizados. En SAP, los leads se registran como entidades aparte hasta que generan primer pedido y se les crea CardCode. Hoy la app solo soporta selección de Clientes con CardCode existente (lectura de `Clientes Activos.xlsx`). El flujo de Lead (selección o creación) está pendiente hasta que SAP B1 Service Layer esté disponible vía Vertrou — entonces los leads se podrán crear directamente en SAP, evitando el paso manual de captura externa.

Mientras tanto, cuando un asesor visita un prospecto que no aparece en `Clientes Activos.xlsx`, el combobox del check-in cae a fallback de texto libre: el nombre se guarda en `Title` pero `CardCode` queda vacío. Esto es señal medible — ver pending work sobre medir uso del fallback.

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
- **OVERRIDE temporal de `gperez@vinssa.com → RAMON VILLEGAS`** en `EMAIL_A_ASESOR` (config.js). Eliminar esta entrada cuando se expanda el piloto a más asesores. La entrada existe solo para que el Director (Gerardo) pueda probar el flujo end-to-end durante la fase de pulido sin requerir login con cuenta de asesor. Mapeado a Ramón Villegas porque tiene cartera real significativa (~76 clientes formalmente asignados en SAP) lo cual permite probar el flujo con datos densos. Antes estuvo mapeado a Kimberly Portillo (1 cliente formalmente asignado), insuficiente para validación realista del autocomplete.
- **Regenerar `Oportunidades.xlsx` con columna `DocCurrency`** cuando sea oportuno. Hoy la app asume MXP por defecto para pre-llenado de monto en cierres ganados. No es bloqueador para el piloto pero será importante cuando haya cierres en USD.
- **Match exacto normalizado de nombres de cliente** (uppercase + sin acentos + trim) es la convención actual para matchear cliente tipeado por asesor contra cliente en `Oportunidades.xlsx`. Si los asesores reportan fricción frecuente con esto, migrar a match tolerante (contains).
- **Búsqueda dentro del dropdown de oportunidades**: actualmente es dropdown plano. Cuando un asesor con muchas oportunidades reporte fricción al hacer scroll, agregar un campo de filtro por nombre/marca arriba del dropdown.
- **Flujo "actuar como asesor" para directores y gerentes**: hoy la app está diseñada para que cada usuario MSAL mapee 1:1 a un asesor SAP. Los directores y gerentes pueden necesitar cerrar oportunidades en nombre de un asesor (cuando él delega, está fuera, o renuncia). Diseño futuro: dropdown adicional "Actuando como: [asesor]" visible solo para roles Director/Gerente. El cierre queda auditado como "cerrado por X en nombre de Y".
- **Reasignación de oportunidades de asesores que salen**: cuando un asesor deja Vinssa, sus oportunidades activas en SAP quedan huérfanas si no se reasignan. Hoy la reasignación se hace manualmente en SAP. Futuro: vista para Gerentes que permita reasignar oportunidades de un asesor saliente a otro activo, con un solo flujo. Esto cierra el ciclo de gestión de carga por asesor.
- **Medir uso del fallback de texto libre**: cada vez que un asesor cierra/avanza una oportunidad usando el campo de texto libre (porque no había oportunidad en SAP), eso es señal de que la disciplina de captura de oportunidades en SAP necesita atención. Agregar reporte simple que cuente cuántos registros pasaron por fallback (`OpportunidadID` vacío y `OppNombre` con texto) vs cuántos por dropdown (`OpportunidadID` poblado), segmentado por asesor y por mes. Si el ratio de fallback es alto (>20%), abrir conversación con los gerentes sobre captura en SAP. **Aplica el mismo principio al fallback de cliente en check-in**: registros con `CardCode` vacío indican que el asesor visitó un cliente/prospecto no presente en `Clientes Activos.xlsx`.
- **Flujo de Lead**: agregar capacidad de seleccionar/crear Leads en SAP cuando Vertrou resuelva el SSL del Service Layer. Hoy si un asesor visita un prospecto sin CardCode, cae a fallback de texto libre (CardCode = null en SharePoint). Cuando SAP conecte, este flujo se reemplaza por selección de Leads (entidad OCRD con `CardType='L'`) y opción de crear nuevo Lead directamente.
- **Re-export periódico de `Clientes Activos.xlsx`**: el archivo refleja un snapshot de SAP en el momento del export. Definir frecuencia de re-export (semanal, mensual) y/o automatizar con Power Automate cuando SAP esté disponible.
- **Priorización en autocomplete de Clientes Activos**: deferida hasta que los datos de asignación de asesor en SAP sean confiables o se diseñe esquema alternativo. Actualmente ~32% de clientes están bajo MOSTRADOR o sin asesor asignado, y muchas asignaciones humanas no reflejan operación real. Por eso priorización por asesor no aporta valor universal — el combobox de check-in usa orden alfabético puro. Cuando los datos de SAP mejoren o se diseñe esquema alternativo (geográfico, por planta, por frecuencia de visita real, etc.), regresar a este punto y revisar `_cbCliActivoOrdenar` en `app.js`.
- **Regenerar `Oportunidades.xlsx` con la query actualizada a `MaxSumSys`/`WtSumSys`** (USD nativos). La query en `docs/vinssa-dashboard-status.md` ya está actualizada (2026-05-16, ver gotcha "SAP queries — Oportunidades"). El código de lectura (`cargarOportunidadesAsesor` en `sap.js`) no requiere cambios: la columna del xlsx se sigue llamando `MontoEstimado`, solo cambian los valores subyacentes. Una vez subido el xlsx nuevo, los totales en el tab Pipeline y en "Cierre proyectado por mes" quedan en USD consistentes con Ventas/OVs (hoy están inflados ~17× para opps en MXP).
- **Evaluar si queries de OINV y ORDR pueden migrar a campos System** (`TotalSys` equivalente en `INV1`/`RDR1`) en lugar del `CASE WHEN T2.TotalFrgn ...` actual. Beneficio: simplificar lógica de las queries de Ventas y OVs alineándolas con el patrón ya usado en Oportunidades. Pendiente verificar disponibilidad de campos `Sys` en `INV1`/`RDR1`.
- **Cálculo de % de cierre real basado en histórico**: actualmente los % de probabilidad por etapa vienen de configuración (`DASHBOARD_CONFIG.etapas`: Cotización=25%, Negociación=80%, etc.) pero no reflejan tasa real de cierre histórica de Vinssa. Pendiente: agregar análisis que calcule histórico de oportunidades cerradas vs perdidas por etapa para sustituir o complementar los % configurados. Esto permitiría visualización de monto ponderado realista en el futuro (en el tab Pipeline y en Por Mes si se decide mostrar ponderado).
- **Pipeline por Mes — mejoras futuras**: filtros adicionales (por línea, por marca, por cliente específico), exportación a CSV, comparativa contra meses anteriores ("¿cómo lucía este mismo bucket hace 30 días?"), alertas automáticas para oportunidades vencidas (Power Automate notificando al asesor cuando una opp cruza vencida sin actualización).

### Hallazgos del 2026-05-15 — calidad de datos de Oportunidades

Descubiertos al probar la sección "Cierre proyectado por mes" con datos reales (override `gperez@vinssa.com → RAMON VILLEGAS`, 76 clientes formalmente asignados). La sección apareció vacía y la investigación reveló múltiples capas de problema:

- **Bug en lectura de `Oportunidades.xlsx`** ✓ **resuelto 2026-05-16 — Bug 1**: `cargarOportunidadesAsesor` en `sap.js` solo mapeaba 8 de las 13 columnas del xlsx. Faltaban: `Asesor`, `FechaApertura`, `MontoPonderado`, `EtapaCodigo`, `Probabilidad`. Verificado en consola: `STATE.oportunidades[0]` no tenía propiedad `Asesor`. Consecuencia inmediata: el filtro por rol no funcionaba en el tab Pipeline ni en la sección "Cierre proyectado por mes"; el dashboard renderizaba con datos truncados. Fix: mapper actualizado a 13 columnas preservando capitalización original (Mayúsculas, convención `STATE.*`).

- **Discrepancia SAP ↔ app en cantidad de oportunidades**: SAP reporta 868 oportunidades con `Status='O'`, pero `STATE.oportunidades.length = 37` para el usuario logueado. Investigar si el delta es filtro de asesor legítimo (solo carga las de Ramón Villegas por nombre) o si hay un filtro adicional silencioso. Cruzar con el bug de columnas faltantes anterior — si `Asesor` no se está leyendo, el filtro de asesor podría estar fallando y arrastrando otros efectos.

- **FechaCierre vacía sistemáticamente en oportunidades cargadas**: `STATE.oportunidades.filter(o => o.FechaCierre).length = 0`. Las 37 opps cargadas no tienen `FechaCierre` poblada. **NO es bug de código** — es problema de captura en SAP. Muchas oportunidades son antiguas (2018-2019) creadas y nunca actualizadas. Esta es la razón principal por la que la sección "Cierre proyectado por mes" no muestra datos útiles incluso después de arreglar el bug de carga.

- **Limpieza de datos en SAP — conversación pendiente con Fernando Barajas + gerentes**: las 868 oportunidades "abiertas" en SAP probablemente contienen mayoría de oportunidades fantasmas (creadas, nunca cerradas, sin FechaCierre, sin actualizaciones recientes). Temas a discutir: disciplina de captura (FechaCierre obligatoria, Asesor asignado, moneda explícita), cierre periódico de oportunidades obsoletas, criterio operativo para distinguir "abierta legítima" vs. "zombie". Sin esta limpieza, los reportes y vistas de pipeline tienen valor limitado por más que el código esté correcto.

- **Sección "Cierre proyectado por mes" bloqueada por calidad de datos**: la feature está implementada y validada operativamente (drill-down, filtros por rol, casos borde funcionan correctamente), pero no aporta valor visible al usuario hasta que (a) se arregle el bug de carga de columnas en `sap.js`, y (b) los datos en SAP se limpien (FechaCierre poblada en oportunidades vivas). Cuando ambos resuelvan, la feature debería revelar la distribución real del pipeline futuro.

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
