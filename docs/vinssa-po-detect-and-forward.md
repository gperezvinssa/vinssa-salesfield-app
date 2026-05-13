# Vinssa — Flujo Power Automate: Detección y Reenvío de POs

## Estado actual
Flujo construido y probado exitosamente el 1 de mayo 2026. Escalación al gerente implementada el 1 de mayo 2026. Probado con PO real el 5 de mayo 2026. Filtro de Order Verifications de Telesis agregado el 11 de mayo 2026.

## Nombre del flujo en Power Automate
`PO - Detect and Forward`

## Estructura del flujo (versión actual funcional)

```
Trigger: When a new email arrives (V3) — gperez@vinssa.com
  → Solo con adjuntos (Only with Attachments: Yes)
  → Sin Subject Filter (vacío)
        ↓
Condition (principal)
  → Expression: if(or(contains(toLower(trigger()['outputs']['body']['subject']),'po'),
    contains(toLower(trigger()['outputs']['body']['subject']),'purchase order'),
    contains(toLower(trigger()['outputs']['body']['subject']),'orden de compra')),
    if(and(not(contains(toLower(trigger()['outputs']['body']['from']),'vinssa.com')),
    not(contains(toLower(trigger()['outputs']['body']['from']),'telesistech.com'))),
    true,false),false)
  → is equal to: true
  ← Filtra: emails internos @vinssa.com Y Order Verifications de @telesistech.com
        ↓ TRUE
For each 1 (itera sobre adjuntos del email)
  │
  ├─► Send email with options
  │     To: gperez@vinssa.com (CAMBIAR a email del asesor en producción)
  │     Subject: Confirmar reenvio - [Subject del trigger]
  │     Body: Posible PO recibida de: [From]
  │           Fecha: convertTimeZone(triggerOutputs()?['body/receivedDateTime'],'UTC','Central Standard Time','dd/MM/yyyy HH:mm')
  │           Asunto: [Subject]
  │           Mensaje: [Body Preview]
  │     User Options: Sí reenviar a SAC, No ignorar
  │     Attachments: name + contentBytes (del For each)
  │     Action Timeout: PT4H
  │
  ├─► Condition 2  (Run after: is successful)
  │     → SelectedOption is equal to: Sí reenviar a SAC
  │           ↓ TRUE
  │     Forward an email (V2) 1
  │       Message Id: [Message Id del trigger]
  │       To: jrodriguez@vinssa.com (email de SAC para USA — cambiar por territorio)
  │       Comment: PO recibida y confirmada por asesor. Favor de generar OV en SAP.
  │                Asesor: [From]
  │                Fecha de recepción: convertTimeZone(triggerOutputs()?['body/receivedDateTime'],'UTC','Central Standard Time','dd/MM/yyyy HH:mm')
  │           ↓
  │     Post message in a chat or channel 1
  │       Post as: Flow bot
  │       Post in: Chat with Flow bot
  │       Recipient: gperez@vinssa.com (CAMBIAR a email del asesor en producción)
  │       Message: ✅ PO enviada a Servicio al Cliente
  │                Cliente: [From]
  │                Fecha de envío: convertTimeZone(triggerOutputs()?['body/receivedDateTime'],'UTC','Central Standard Time','dd/MM/yyyy HH:mm')
  │                Servicio al Cliente fue notificado y generará la OV en SAP.
  │
  └─► Send an email (V2)  ← ESCALACIÓN AL GERENTE
        Run after: Send email with options → Has timed out
        To: gperez@vinssa.com ← temporalmente al mismo Gerardo durante pruebas
        Subject: ⚠️ PO sin confirmar — requiere atención: [Subject]
        Body: El asesor no confirmó la PO en el tiempo establecido.
              Cliente: [From]
              Asunto original: [Subject]
              Fecha de recepción: convertTimeZone(triggerOutputs()?['body/receivedDateTime'],'UTC','Central Standard Time','dd/MM/yyyy HH:mm')
              Por favor revisa el email adjunto y reenvía a Servicio al Cliente si aplica.
        Attachments: name + contentBytes (del For each)
        Importance: High
```

## Lo que funciona
- Detecta emails con PO en el asunto (po, purchase order, orden de compra) — case insensitive
- Filtra emails enviados desde @vinssa.com para evitar loops y falsos positivos ✅
- Filtra Order Verifications de @telesistech.com para evitar falsos positivos ✅
- Notificación al asesor con PDF adjunto y 2 botones (Sí reenviar a SAC / No ignorar)
- Notificación incluye: remitente, asunto, fecha legible en CST y body preview
- Asesor confirma con un toque
- Reenvío automático a Servicio al Cliente con contexto completo
- Fecha en formato legible dd/MM/yyyy HH:mm (convertida de UTC a Central Standard Time) ✅
- Confirmación al asesor por Teams (Workflows app)
- Escalación automática al gerente si el asesor no responde en 4 horas ✅
- Probado con PO real el 5 de mayo 2026 ✅

## Notas técnicas importantes
- La expression principal usa `trigger()['outputs']['body']` — NO usar `triggerOutputs()?['body/']` dentro de condiciones
- La expression principal usa `and(not(...), not(...))` para filtrar tanto @vinssa.com como @telesistech.com
- La Condition 2 debe estar DENTRO del For each 1, después del "Send email with options"
- El Flow bot de Teams requiere que cada usuario tenga la app "Workflows" instalada en Teams
- Sin Subject Filter en el trigger — el filtrado se hace con la expression en la Condition principal
- El Action Timeout usa formato ISO 8601: `PT4H` — escribir `P4H` sin la T no funciona
- El "Configure run after" del paso de escalación apunta a "Send email with options" con "Has timed out" marcado e "Is successful" desmarcado
- El paso de escalación usa "Send an email (V2)" — NO "Send email with options"
- Plataforma: Power Automate (make.powerautomate.com)
- Trigger: Office 365 Outlook — When a new email arrives (V3)

## Pendientes para implementación completa

### Prioridad alta — antes de activar con asesores
1. ~~**Escalación al gerente**~~ — ✅ Implementado el 1 de mayo 2026.
2. ~~**Fecha en formato legible**~~ — ✅ Implementado el 5 de mayo 2026.
3. ~~**Body y contexto del cliente en notificación**~~ — ✅ Implementado el 5 de mayo 2026.
4. ~~**Filtrar Order Verifications de Telesis**~~ — ✅ Implementado el 11 de mayo 2026.
5. **Cambiar email de escalación a Fernando Barajas** — actualmente va a gperez@vinssa.com durante pruebas. Cambiar a fbarajas@vinssa.com cuando se active en producción.
6. **Activar Flow bot de Teams** — requiere administrador de Microsoft 365 del tenant de Vinssa.
7. **Probar con PO real de cliente directo** — prueba del 5 de mayo fue con PO reenviada por Telesis.

### Prioridad media — replicación a asesores
8. **Replicar flujo a todos los asesores** — cada asesor crea el flujo en su cuenta. Email de SAC cambia según territorio: México vs USA. Estimado: 15 minutos por asesor con guía.
9. **Múltiples PDFs en un email** — actualmente el For each manda una notificación por adjunto. Solución: variable de array para consolidar adjuntos en un solo email.

### Prioridad baja — optimizaciones futuras
10. **Tracker de POs en SharePoint** — registrar cada PO en Excel con: asesor, cliente, fecha, estatus. Conectar a Power BI.
11. **Confirmación automática al cliente** — cuando SAC crea la OV en SAP, email automático al cliente.
12. **Extracción de datos con IA** — AI Builder para extraer del PDF: No. PO, cliente, producto, cantidad, precio. ROI claro con ~24 POs/día. Costo estimado <$50/mes.

## Equipo involucrado
- **Gerardo Pérez** (gperez@vinssa.com) — owner del flujo, Director de Trazabilidad y Automatización
- **Fernando Barajas** (fbarajas@vinssa.com) — Gerente división USA, receptor de escalaciones en producción
- **Servicio al Cliente USA** (jrodriguez@vinssa.com — confirmar email correcto) — destinatario de POs USA
- **Servicio al Cliente México** — email pendiente de confirmar

## Contexto de negocio
- Vinssa es distribuidor industrial mexicano con ~51 empleados
- Operaciones en Saltillo, Torreón, Durango, Monclova + expansión Texas
- Asesores reciben POs de clientes por email → deben reenviar a Servicio al Cliente → SAC genera OV en SAP B1
- Problema previo: asesores olvidaban reenviar o tardaban, causando retrasos en generación de OV
- Herramientas: Microsoft 365 (Outlook, Teams, Power Automate), SAP Business One, Zoho (SAC)
- Volumen: ~409 OVs en 17 días hábiles (abril-mayo 2026) = ~24 POs/día
