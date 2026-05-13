# Vinssa Field App — Documentación del Proyecto

## Estado actual
App construida y publicada el 6 de mayo 2026. Login con Microsoft funcionando (gperez@vinssa.com y vinssabi@vinssa.com). Formulario completo con campos dinámicos. Geolocalización 3 capas implementada. Check-in obligatorio antes de registrar. Registros guardándose correctamente en SharePoint List via Microsoft Graph API. Conexión SAP B1 pendiente de certificado SSL válido de Vertrou.

## URL de la app
```
https://gperezvinssa.github.io/vinssa-salesfield-app/
```

## Repositorio GitHub
```
https://github.com/gperezvinssa/vinssa-salesfield-app
```

## Estructura de archivos
```
vinssa-salesfield-app/
├── index.html          — Estructura HTML + pantalla de login
├── manifest.json       — Configuración PWA (instalar en celular)
├── icon.png            — Ícono de la app
├── css/
│   └── styles.css      — Estilos completos
└── js/
    ├── config.js       — Catálogos y configuración central
    ├── auth.js         — Login con Microsoft (MSAL)
    ├── geo.js          — Geolocalización 3 capas + check-in/checkout
    ├── sap.js          — Conexión SharePoint (Graph API) + SAP pendiente
    └── app.js          — Lógica completa del formulario
```

---

## Configuración Azure (Microsoft Login)

- **App registrada:** Vinssa Sales Field App
- **Application (client) ID:** 45d6f369-f789-473b-8970-d9b25ff3225c
- **Directory (tenant) ID:** 7a272c1a-ee40-4b22-a187-2656ea44b4c4
- **Redirect URI:** https://gperezvinssa.github.io/vinssa-salesfield-app/
- **Tipo:** Single-page application (SPA)
- **Acceso:** Solo inquilino Vinssa (VERSATILIDAD INDUSTRIAL DE SALTILLO SA)

### Permisos configurados en Azure (todos con admin consent ✅)
| API | Permiso | Tipo |
|---|---|---|
| Microsoft Graph | User.Read | Delegated |
| Microsoft Graph | Sites.ReadWrite.All | Delegated |
| SharePoint | AllSites.Write | Delegated |

---

## Configuración SAP B1 Service Layer

- **URL:** https://vinssa.vertrou.cloud:50000/b1s/v1
- **IP:** 148.244.98.206
- **Proveedor:** Vertrou Cloud
- **Estado:** Activo pero con certificado SSL auto-firmado ⚠️
- **Bloqueador:** Certificado SSL auto-firmado no aceptado por Azure/Power Automate
- **Solución pendiente:** Solicitar a Vertrou instalar certificado SSL válido (Let's Encrypt) en puerto 50000
- **Solicitud enviada a Vertrou:** Sí — pendiente de confirmación

---

## SharePoint — Lista de visitas ✅ FUNCIONANDO

- **Lista:** Visitas Field App
- **Sitio:** https://versatilidadsaltillo.sharepoint.com/sites/VINSSAAutomation
- **Estado:** Recibiendo registros correctamente desde la app ✅
- **Conexión:** Microsoft Graph API con scope Sites.ReadWrite.All

### Columnas de la lista
| Columna display | Nombre interno | Tipo |
|---|---|---|
| Title | Title | Single line (Cliente) |
| Asesor | Asesor | Single line |
| Tipo | Tipo | Single line |
| Marca | Marca | Single line |
| Producto | Producto | Single line |
| Etapa | Etapa | Single line |
| Monto | Monto | Number |
| Moneda | Moneda | Single line |
| Cierre | Cierre | Date |
| Competidor | Competitor | Single line ⚠️ nombre interno en inglés |
| Lideres | Lideres | Single line |
| Notas | Notas | Multiple lines |
| GPS_Lat | GPS_Lat | Number |
| GPS_Lng | GPS_Lng | Number |
| GPS_Precision | GPS_Precision | Number |
| Checkin_Hora | Checkin_Hora | Single line |
| Checkout_Hora | Checkout_Hora | Single line |
| Duracion_Min | Duracion_Min | Number |
| Contacto_Nuevo | Contacto_Nuevo | Single line |
| Contacto_Nombre | Contacto_Nombre | Single line |
| Contacto_Puesto | Contacto_Puesto | Single line |
| Contacto_Tel | Contacto_Tel | Single line |
| Contacto_Email | Contacto_Email | Single line |
| SAP_OppID | SAP_OppID | Single line |
| SAP_ActID | SAP_ActID | Single line |
| Estatus_SAP | Estatus_SAP | Single line |

**Nota:** La columna "Competidor" tiene nombre interno "Competitor" en inglés — manejado en sap.js con `Competitor:`.

---

## Geolocalización — 3 capas implementadas ✅

### Capa 1 — GPS al guardar registro
- Captura automática de coordenadas al momento de guardar
- Si el asesor está a más de 500m del cliente, muestra alerta suave (no bloquea)
- Coordenadas se guardan en SharePoint: GPS_Lat, GPS_Lng, GPS_Precision

### Capa 2 — Check-in / Check-out ✅
- **Check-in obligatorio** antes de poder registrar visita, demo, lead o actualizar oportunidad
- Los 4 botones de registro aparecen deshabilitados hasta que hay check-in activo
- Al hacer check-in: registra hora de llegada, nombre del cliente y coordenadas GPS
- Al hacer check-out: registra hora de salida y duración en minutos
- El botón cambia a "Terminar visita" con punto verde y header verde cuando hay visita activa
- Check-in persiste en localStorage si se cierra la app

### Capa 3 — Datos para Power BI
- Todos los registros GPS van a SharePoint automáticamente
- Conectar SharePoint List a Power BI para dashboard de zonas y rutas

---

## Power Automate — Flujo creado, pendiente de completar

**Nombre:** `Field App - Sincronizar registro con SAP`
**Estado:** Trigger configurado (SharePoint → Visitas Field App), conexión SAP OData creada pero bloqueada por certificado SSL

**Conexión SAP creada:** `SAP B1 Vinssa` (SAP OData connector)
- URL: https://vinssa.vertrou.cloud:50000/b1s/v1
- Auth: Basic

**Estructura planificada una vez resuelto el certificado:**
```
Trigger: When an item is created (SharePoint → Visitas Field App)
    ↓
Condition: Estatus_SAP = "Pendiente" AND Tipo ≠ "lead"
    ↓ TRUE
SAP OData: Create OData entity → SalesOpportunities
    ↓
SAP OData: Create OData entity → Activities
    ↓
Condition: Contacto_Nuevo = "Si"
  ↓ TRUE
SAP OData: Update OData entity → BusinessPartners (agregar contacto)
    ↓
SharePoint: Update item
  SAP_OppID = [ID devuelto]
  SAP_ActID = [ID devuelto]
  Estatus_SAP = "Sincronizado"
```

---

## config.js — Catálogos

### Divisiones y marcas
| División | Línea | Marcas |
|---|---|---|
| Trazabilidad y Automatización | Marcaje | Telesis |
| Trazabilidad y Automatización | Visión | Cognex |
| Trazabilidad y Automatización | Robótica | UR, MIR, Nabtesco, EasyRobotics |
| Trazabilidad y Automatización | Herramientas | Atlas Copco |
| Suministros Industriales | General | Por definir |
| Servicios Industriales | General | Por definir |

### Etapas (exactas de SAP B1)
| ID | Label | % Cierre |
|---|---|---|
| contacto_inicial | Contacto Inicial | 10% |
| cotizacion | Cotización | 25% |
| pruebas_demo | Pruebas / Demostración | 60% |
| negociacion | Negociación | 80% |
| tramite_compras | Trámite con Compras | 90% |
| factura | Factura | 95% |
| contacto_web | Contacto WEB/Teléfono | 10% |
| registro_lead | Registro de Lead Expos | 10% |
| perdido | Perdido | 0% |

### Líderes de línea — Trazabilidad
| Nombre | Email | Líneas |
|---|---|---|
| Antonio Martinez | amartinez@vinssa.com | Marcaje, Visión |
| Jose Juan Aguillon | jaguillon@vinssa.com | Robótica |
| Aldo Almaguer | aalmaguer@vinssa.com | Herramientas |
| Jonathan Roche | jroche@vinssa.com | Visión, Marcaje |

### Competidores
Keyence, SIC Marking, Ninguno, Otro

---

## Cómo actualizar catálogos (config.js)

1. Ir a GitHub → `js/config.js` → ícono de lápiz
2. Hacer el cambio en la sección correspondiente
3. Commit con mensaje descriptivo
4. Esperar ~1 minuto y recargar la app

---

## Roadmap de implementación

### Completado ✅
- [x] Estructura base PWA — publicada en GitHub Pages
- [x] Login con Microsoft (MSAL + Azure) — gperez y vinssabi funcionando
- [x] Formulario dinámico completo (4 tipos de registro)
- [x] Marcas dinámicas por división y línea
- [x] Etapas reales de SAP B1
- [x] Líderes de línea con búsqueda y múltiple selección
- [x] Selector MXP/USD
- [x] Contacto nuevo con toggle — campos nombre, puesto, tel, email
- [x] Geolocalización 3 capas (GPS al guardar, check-in/out, datos para Power BI)
- [x] Check-in obligatorio — 4 botones deshabilitados hasta check-in activo
- [x] SharePoint List creada con todas las columnas
- [x] Registros guardándose en SharePoint via Microsoft Graph API ✅
- [x] Conexión SAP OData creada en Power Automate

### Pendiente inmediato 🔄
- [ ] **Vertrou: instalar certificado SSL válido** en vinssa.vertrou.cloud:50000 — BLOQUEADOR PRINCIPAL
- [ ] Una vez resuelto: completar flujo Power Automate (SAP OData → crear oportunidad + actividad + contacto nuevo)
- [ ] Validación en app: si etapa nueva tiene % menor a la actual, mostrar advertencia antes de guardar

### Pendiente próximas sesiones 📋
- [ ] Actualizar oportunidad con datos pre-llenados desde SAP (requiere certificado SSL)
- [ ] Check-in con sugerencia de cliente por proximidad GPS (requiere conexión a lista de clientes SAP)
- [ ] Notificación a líder de línea vía Teams cuando se registra demo con su nombre
- [ ] Dashboard Power BI con mapa de zonas GPS y actividad por asesor
- [ ] Replicar a asesores de Trazabilidad (piloto) — Rafael Moyeda, Néstor Carranza
- [ ] Usuario técnico SAP dedicado para la integración (solicitar a Vertrou)
- [ ] Expandir a Suministros y Servicios

---

## Decisiones de diseño tomadas

### Check-in y KPIs
- Check-in es obligatorio antes de registrar — garantiza GPS ligado a cada actividad
- No hay tracking continuo — solo captura en check-in y al guardar
- No bloquea ni penaliza — muestra datos en dashboard para que el gerente vea patrones

### Etapas en SAP
- SAP B1 guarda cada stage como línea independiente — no valida orden progresivo
- Si asesor registra etapa menor a la actual, SAP lo guarda tal cual (dos líneas)
- Solución planeada: advertencia en app cuando % nueva < % actual, sin bloquear

### Geolocalización — comunicación al equipo
- Presentar como herramienta de optimización de rutas y planeación, no como vigilancia
- Los datos de GPS aparecen en Power BI para análisis de zonas — no para monitoreo individual

---

## Contexto de negocio

- **Empresa:** Vinssa (Versatilidad Industrial de Saltillo SA)
- **Piloto:** Equipo de Trazabilidad y Automatización
- **Asesores:** Rafael Moyeda, Néstor Carranza y otros
- **Líderes de línea:** Antonio Martinez, Jose Juan Aguillon, Aldo Almaguer, Jonathan Roche
- **Herramientas:** Microsoft 365, SAP Business One (Vertrou Cloud), GitHub Pages, SharePoint, Power Automate
- **Objetivo:** Registrar visitas y demos en campo desde celular → SharePoint → SAP automático → Power BI

## Notas técnicas importantes

- La app es una PWA — se instala en celular desde el navegador sin App Store
- En iPhone: Safari → Compartir → "Agregar a pantalla de inicio"
- En Android: Chrome muestra banner de instalación automático
- El login con Microsoft guarda la sesión — no pide contraseña cada vez
- GitHub Pages tarda ~1 minuto en reflejar cambios — forzar recarga: Ctrl+Shift+R
- Zona horaria: America/Monterrey (Central Standard Time)
- SAP Service Layer no acepta conexiones directas desde el navegador (CORS) — se usa Power Automate como puente
- SharePoint escribe via Microsoft Graph API con scope Sites.ReadWrite.All
- La columna "Competidor" tiene nombre interno "Competitor" en SharePoint — manejado en sap.js
- Los registros GPS se guardan en localStorage como respaldo si SharePoint falla
- Power Automate plan básico (sin Premium) — no tiene HTTP trigger — se usa SharePoint List como intermediario
- El flujo Power Automate "Field App - Sincronizar registro con SAP" está creado pero pausado hasta resolver certificado SSL
