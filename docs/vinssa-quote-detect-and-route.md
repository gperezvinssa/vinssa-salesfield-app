# Vinssa — Flujo Power Automate: Detección y Enrutamiento de Cotizaciones USA

## Estado actual
Flujo construido el 4 de mayo 2026. Probado con leads reales de Telesis el 5-6 de mayo 2026. Tracker de leads en SharePoint implementado el 6 de mayo 2026.

## Nombre del flujo en Power Automate
`Quote - Detect and Route`

## Alcance actual
Flujo configurado para la cuenta de Gerardo Pérez (gperez@vinssa.com) — territorio USA (Texas, Louisiana, Oklahoma). Pendiente replicar a México una vez definido el proceso completo.

## Orígenes de leads detectados
| Origen | Remitente | Descripción |
|---|---|---|
| Web Vinssa | noreply@cliengo.com | Lead del chat de la página web de Vinssa |
| Telesis | inquiries@telesistech.com | Lead reenviado por Telesis Technologies |
| Cliente directo | Cualquier otro | No detectado por este flujo — llega directamente al asesor |

## Estructura del flujo

```
Trigger: When a new email arrives (V3) — gperez@vinssa.com
  → Only with Attachments: No
  → Sin Subject Filter
        ↓
Initialize variable
  Name: CustomerEmail
  Type: String
  Value: (vacío)
        ↓
Set variable
  Name: CustomerEmail
  Value: first(body('Extract_standard_entities')?['entities'])?['value']
  ← PENDIENTE: falla si AI Builder no tiene licencia o body es HTML
        ↓
Add a row into a table  ← TRACKER DE LEADS ✅
  Location: Group - VINSSA Automation
  Document Library: Documents
  File: /Tracker Leads.xlsx
  Table: Table1
  Fecha: formatDateTime(utcNow(), 'dd/MM/yyyy HH:mm')
  Origen: if(contains(toLower(from),'cliengo'),'Web Vinssa','Telesis')
  Cliente: [From]
  Email Cliente: [CustomerEmail]
  Asunto: [Subject]
  Territorio: if(contains(body,'texas'),'Texas',if(contains(body,'louisiana'),'Louisiana',if(contains(body,'oklahoma'),'Oklahoma','Sin detectar')))
  Estatus: Nuevo lead
  Notas: (vacío)
        ↓
Condition (principal) — ¿Es Cliengo o Telesis?
  toLower(from) contains 'noreply@cliengo.com'
  OR toLower(from) contains 'inquiries@telesistech.com'
        ↓ TRUE
  Condition 1 — ¿Es territorio válido? (TX, LA, OK)
  body contains 'texas' OR 'louisiana' OR 'oklahoma'
    │
    ├─► TRUE — Cliente en territorio
    │     Condition 3 — ¿Menciona producto específico?
    │     body contains: bonderite, nabtesco, cognex, easyrobotics, theo,
    │                    tmp, pinstamp, versa, triton
    │       │
    │       ├─► TRUE — Cliente sabe qué quiere
    │       │     Send email with options 1
    │       │       To: gperez@vinssa.com
    │       │       Subject: 🔵 Lead con producto específico — [Subject]
    │       │       Body: Nuevo lead recibido con producto identificado.
    │       │             Cliente: [From]
    │       │             Producto mencionado: [Body]
    │       │             Fecha: convertTimeZone(...)
    │       │             ¿Solicitar cotización a Jonas?
    │       │       Options: Sí solicitar a Jonas, No ignorar
    │       │           ↓
    │       │     Condition 4 — SelectedOption = Sí solicitar a Jonas
    │       │       ↓ TRUE
    │       │     Send an email (V2) 1
    │       │       To: jrodriguez@vinssa.com (Jonas — SAC)
    │       │       Subject: Cotización requerida — [From]
    │       │       Body: Hola Jonas, favor de cotizar:
    │       │             Cliente: [From] / Solicitud: [Body]
    │       │
    │       └─► FALSE — Cliente no especifica producto
    │             Send email with options 2
    │               To: gperez@vinssa.com
    │               Subject: 🟡 Lead sin producto específico — [Subject]
    │               Body: Nuevo lead sin producto identificado.
    │                     Cliente: [From] / Mensaje: [Body Preview]
    │                     Fecha: convertTimeZone(...)
    │               Options: Sí enviar template, No ignorar
    │                   ↓
    │             Condition 5 — SelectedOption = Sí enviar template
    │               ↓ TRUE
    │             Send an email (V2) 2
    │               To: CustomerEmail (variable) ← PENDIENTE resolver extracción
    │               Subject: Your Telesis Inquiry — A Few Quick Questions | VINSSA
    │               Body: Template de preguntas técnicas (ver sección de templates)
    │
    └─► FALSE — Cliente fuera de territorio
          Send email with options
            To: gperez@vinssa.com
            Subject: ⚠️ Lead fuera de territorio — [Subject]
            Body: Lead recibido fuera de territorio.
                  Cliente: [From] / Asunto: [Subject] / Mensaje: [Body Preview]
                  ¿Referir a Chuck Pemble (Telesis)?
            Options: Sí referir a Chuck, No ignorar
                ↓
          Condition 2 — SelectedOption = Sí referir a Chuck
            ↓ TRUE
          Send an email (V2)
            To: chuck.pemble@telesistech.com
            Subject: New Lead Referral | [From]
            Body: Hi Chuck, forwarding lead outside territory.
                  Customer: [From] / Inquiry: [Body]
                  Best, Gerardo
```

## Tracker de leads — SharePoint ✅
- **Archivo:** Tracker Leads.xlsx
- **Ubicación:** SharePoint → Group - VINSSA Automation → Documents
- **Tabla:** Table1
- **Columnas:** Fecha, Origen, Cliente, Email Cliente, Asunto, Territorio, Estatus, Notas
- **Se registra:** automáticamente con cada lead de Cliengo o Telesis, antes de cualquier decisión de enrutamiento

## Templates aprobados

### Template A — Preguntas técnicas (cliente sin producto específico)
```
Subject: Your Telesis Inquiry — A Few Quick Questions | VINSSA

Hi there,

Thank you for reaching out! I'm Gerardo Pérez from VINSSA, exclusive Telesis
distributor in Texas, Louisiana, and Oklahoma.

To recommend the right solution for your marking application, could you share
a few details?

1. What material are you marking?
2. Flat or cylindrical surface?
3. Required marking window size?
4. Marking depth needed?
5. Cycle time requirement?
6. Estimated production volume?
7. Budget and timeline?

A photo of the part would also be very helpful if you have one.

I'm also happy to jump on a quick call or schedule an on-site visit if that
works better for you, just let me know.

Looking forward to hearing from you!

Best regards,
Gerardo Pérez
VINSSA Industrial Solutions | Telesis Authorized Distributor, Texas, Louisiana & Oklahoma
gperez@vinssa.com | (857) 654 0925
```

### Template B — Referir a Chuck (cliente fuera de territorio)
```
Subject: New Lead Referral | [Customer Name]

Hi Chuck,

Forwarding a lead outside our territory.

Customer: [From]
Inquiry: [Body]

Best,
Gerardo
```

### Template C — Producto fuera de catálogo
```
Subject: Re: Your Inquiry | VINSSA Industrial Solutions

Hi [Name],

Thank you for reaching out! Unfortunately, [product] is not part of our
current portfolio so we are not able to assist with this request.

If you have any other industrial marking, vision, or automation needs in
the future, we would love to help.

Best regards,
Gerardo Pérez
gperez@vinssa.com | (857) 654 0925
```
*Nota: Template C está aprobado pero aún no implementado en el flujo.*

## Notas técnicas importantes
- El trigger NO filtra por adjuntos (los leads no traen adjuntos)
- La Condition principal usa OR para detectar ambos orígenes (Cliengo y Telesis)
- La detección de territorio busca "texas", "louisiana", "oklahoma" en el body — NO usar abreviaciones (tx, la, ok) — generan falsos positivos
- La detección de producto busca nombres de marca/modelo en el body — NO incluir "telesis" ya que aparece en todos los emails de Telesis (links, firma) y genera falsos positivos
- Keywords de producto actuales: bonderite, nabtesco, cognex, easyrobotics, theo, tmp, pinstamp, versa, triton
- Fecha en notificaciones: `convertTimeZone(triggerOutputs()?['body/receivedDateTime'],'UTC','Central Standard Time','dd/MM/yyyy HH:mm')`
- AI Builder "Extract standard entities" requiere licencia adicional (~$500/mes) — no justificado para volumen actual de leads
- Emails con body HTML complejo (ej. Caterpillar) muestran código en la notificación — pendiente resolver con Claude API cuando el volumen lo justifique

## Formatos de email por origen

### Cliengo (noreply@cliengo.com)
```
Nombre: Marcos Carrillo
Email: marcos.carrillo@prolec.energy
Empresa: GE Vernova
Ciudad: Shreveport, LA
```

### Telesis (inquiries@telesistech.com)
```
From: Carlos Rodriguez <crodriguez@powerseal.com>
Company: Powerseal
State: TX
Message: [descripción del requerimiento]
```

## Pendientes para implementación completa

### Prioridad alta
1. **Extracción del email del cliente** — AI Builder no disponible sin licencia. Alternativa futura: Claude API para extraer email y datos del cliente del body HTML. Pendiente hasta que el volumen justifique el costo.

2. **Template C — Producto fuera de catálogo** — aprobado pero no implementado. Agregar rama en Condition 3.

3. **Probar flujo completo con lead real de Cliengo** — pruebas han sido con Telesis.

### Prioridad media
4. **Actualizar Estatus en tracker** — actualmente solo registra "Nuevo lead". Agregar paso que actualice el estatus a "Template enviado", "Cotización solicitada" o "Referido a Chuck" según la decisión de Gerardo.

5. **Extracción del nombre del cliente** — para personalizar template ("Hi Marcos," en lugar de "Hi there,").

6. **Replicar a México** — una vez definido el proceso de cotizaciones México.

### Prioridad baja
7. **Dashboard Power BI** — conectar Tracker Leads.xlsx a Power BI para visibilidad gerencial de pipeline.

8. **Escalación por timeout** — si Gerardo no responde en X horas, escalar a Fernando Barajas.

9. **Confirmación automática al cliente** — cuando Gerardo aprueba enviar a Jonas, email al cliente confirmando que su solicitud está en proceso.

## Equipo involucrado
- **Gerardo Pérez** (gperez@vinssa.com) — owner del flujo, asesor USA
- **Jonas / SAC** (jrodriguez@vinssa.com — confirmar email) — genera cotizaciones
- **Chuck Pemble** (chuck.pemble@telesistech.com) — contacto Telesis para leads fuera de territorio
- **Fernando Barajas** (fbarajas@vinssa.com) — Gerente de la división USA

## Contexto de negocio
- Vinssa tiene exclusividad de Telesis en México y territorios USA: Texas, Louisiana, Oklahoma
- Leads llegan de dos fuentes automáticas: chat web Vinssa (Cliengo) y reenvíos de Telesis
- Líneas de producto USA: Telesis, Nabtesco, Cognex, Easy Robotics, Theo, Bonderite
- Proceso: Gerardo recibe lead → evalúa → pide cotización a Jonas o envía template → Jonas cotiza → Gerardo envía al cliente
- Herramientas: Microsoft 365 (Outlook, Teams, Power Automate), SAP Business One, SharePoint

## Próximos proyectos en roadmap
- **App de inteligencia de precios** (proyecto del hermano de Gerardo) — consultar historial de SAP B1 y recomendar precios por categoría de cliente A/B/C. Prerequisito: validar Service Layer de SAP B1 on-premise (URL: https://[servidor]:50000/b1s/v1). Stack propuesto: aplicación web + Claude API + SAP B1 Service Layer.
- **AI Builder / Claude API para extracción de datos** — cuando el volumen de leads y POs justifique el costo. ROI claro con ~24 POs/día.
- **Dashboard Power BI** — conectar Tracker Leads a Power BI para visibilidad del pipeline de ventas USA.
