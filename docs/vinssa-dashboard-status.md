# Vinssa Sales Dashboard PWA — Estado de Sesión
**Fecha:** 12 de mayo 2026  
**Archivo activo:** `js/dashboard.js` en GitHub → `gperezvinssa.github.io/vinssa-salesfield-app`

---

## Archivos en SharePoint (Documents root)

| Archivo | Contenido | Estado |
|---|---|---|
| `Lista Roles Dashboard.xlsx` | Email, Rol, Division por usuario | ✅ Activo |
| `Ventas Asesor v2.xlsx` | Facturas OINV + notas crédito ORIN | ✅ Activo (re-exportado hoy) |
| `OVs Asesor.xlsx` | Órdenes de venta abiertas ORDR | ✅ Activo |
| `Oportunidades.xlsx` | Oportunidades CRM OOPR | ✅ Activo |
| `Presupuesto Ventas.xlsx` | Meta por Asesor/Mes/División | ✅ Activo (solo Trazabilidad) |

---

## Query SAP — Ventas Asesor v2.xlsx (versión correcta actual)

Usa `TotalFrgn` con fallback via tabla `ORTT` para MXP sin tasa:

```sql
SELECT 
    T1.SlpName AS Asesor,
    T0.CardCode,
    T0.CardName AS Cliente,
    T0.DocNum AS NumFactura,
    CONVERT(varchar, T0.DocDate, 105) AS Fecha,
    T2.ItemCode,
    T2.Dscription AS Producto,
    T3.ItmsGrpNam AS GrupoArticulo,
    CASE 
        WHEN T2.TotalFrgn <> 0 THEN T2.TotalFrgn
        WHEN T4b.Rate > 1 THEN T2.LineTotal / T4b.Rate
        ELSE T2.LineTotal / 17.5
    END AS Total,
    T0.DocCur AS Moneda,
    T0.DocRate AS TipoCambio
FROM OINV T0
LEFT JOIN OSLP T1 ON T0.SlpCode = T1.SlpCode
LEFT JOIN INV1 T2 ON T0.DocEntry = T2.DocEntry
LEFT JOIN OITM T4 ON T2.ItemCode = T4.ItemCode
LEFT JOIN OITB T3 ON T4.ItmsGrpCod = T3.ItmsGrpCod
LEFT JOIN ORTT T4b ON T4b.Currency = 'USD' AND T4b.RateDate = T0.DocDate
WHERE T0.CANCELED = 'N' AND T0.DocDate >= '2024-01-01'

UNION ALL

SELECT 
    T1.SlpName AS Asesor,
    T0.CardCode,
    T0.CardName AS Cliente,
    T0.DocNum AS NumFactura,
    CONVERT(varchar, T0.DocDate, 105) AS Fecha,
    T2.ItemCode,
    T2.Dscription AS Producto,
    T3.ItmsGrpNam AS GrupoArticulo,
    CASE 
        WHEN T2.TotalFrgn <> 0 THEN -T2.TotalFrgn
        WHEN T4b.Rate > 1 THEN -T2.LineTotal / T4b.Rate
        ELSE -T2.LineTotal / 17.5
    END AS Total,
    T0.DocCur AS Moneda,
    T0.DocRate AS TipoCambio
FROM ORIN T0
LEFT JOIN OSLP T1 ON T0.SlpCode = T1.SlpCode
LEFT JOIN RIN1 T2 ON T0.DocEntry = T2.DocEntry
LEFT JOIN OITM T4 ON T2.ItemCode = T4.ItemCode
LEFT JOIN OITB T3 ON T4.ItmsGrpCod = T3.ItmsGrpCod
LEFT JOIN ORTT T4b ON T4b.Currency = 'USD' AND T4b.RateDate = T0.DocDate
WHERE T0.CANCELED = 'N' AND T0.DocDate >= '2024-01-01'
```

**Importante:** Al exportar, formatear columna Fecha como **Texto** antes de subir a SharePoint.

---

## Query SAP — OVs Asesor.xlsx

```sql
SELECT 
    T1.SlpName AS Asesor,
    T0.CardCode,
    T0.CardName AS Cliente,
    T0.DocNum AS NumOV,
    CONVERT(varchar, T0.DocDate, 105) AS Fecha,
    CONVERT(varchar, T0.DocDueDate, 105) AS FechaEntrega,
    T2.ItemCode,
    T2.Dscription AS Producto,
    T3.ItmsGrpNam AS GrupoArticulo,
    CASE 
        WHEN T2.TotalFrgn <> 0 THEN T2.TotalFrgn
        WHEN T4b.Rate > 1 THEN T2.LineTotal / T4b.Rate
        ELSE T2.LineTotal / 17.5
    END AS Total,
    T0.DocCur AS Moneda,
    T0.DocRate AS TipoCambio
FROM ORDR T0
LEFT JOIN OSLP T1 ON T0.SlpCode = T1.SlpCode
LEFT JOIN RDR1 T2 ON T0.DocEntry = T2.DocEntry
LEFT JOIN OITM T4 ON T2.ItemCode = T4.ItemCode
LEFT JOIN OITB T3 ON T4.ItmsGrpCod = T3.ItmsGrpCod
LEFT JOIN ORTT T4b ON T4b.Currency = 'USD' AND T4b.RateDate = T0.DocDate
WHERE T0.DocStatus = 'O' AND T0.CANCELED = 'N'
```

---

## Query SAP — Oportunidades.xlsx

```sql
SELECT 
    T1.SlpName AS Asesor,
    T0.CardName AS Cliente,
    T0.OpprId AS NumOportunidad,
    T0.Name AS Descripcion,
    CONVERT(varchar, T0.OpenDate, 105) AS FechaApertura,
    -- T0.PredDate = Predicted Closing Date (fecha proyectada). NO usar T0.CloseDate
    -- porque ese solo se popula cuando la oportunidad se cierra efectivamente (Status≠'O').
    CONVERT(varchar, T0.PredDate, 105) AS FechaCierre,
    T0.MaxSumSys AS MontoEstimado,
    T0.WtSumSys AS MontoPonderado,
    T0.StepLast AS EtapaCodigo,
    T2.Descript AS Etapa,
    T0.CloPrcnt AS Probabilidad,
    T0.U_Linea AS Linea,
    T0.U_Marca AS Marca
FROM OOPR T0
LEFT JOIN OSLP T1 ON T0.SlpCode = T1.SlpCode
LEFT JOIN OOST T2 ON T0.StepLast = T2.Num
WHERE T0.Status = 'O'
```

---

## Arquitectura del Dashboard

### Roles y divisiones (Lista Roles Dashboard.xlsx)

| Rol | División | Ve |
|---|---|---|
| Director | Todos | Todo — todos los asesores y divisiones |
| Gerente | Trazabilidad | Trazabilidad + Visión + Robótica |
| Gerente | Suministros | Solo Suministros |
| Gerente | Servicios | Solo Servicios |
| Lider | Trazabilidad | Igual que Gerente·Trazabilidad |
| Asesor | cualquiera | Solo sus propios números |

### Columnas requeridas en Lista Roles Dashboard.xlsx
`Email` | `Rol` | `Division`

---

## Listas de asesores en DASHBOARD_CONFIG (dashboard.js)

### asesoresSupply — Suministros (activos en 2026)
```
ALEJANDRO RODRIGUEZ, ARTURO CASTILLO, DANIEL ABASCAL, EDER MORENO,
EDUARDO ESPARZA, ENRIQUE CRUZ, FERNANDO MONTANA ALDAZ, FRANCISCO MARTINEZ,
HORACIO MORALES, JORGE OLAGUIVEL, LINO TOVAR, MARIA VILLALOBOS,
MAYELA HURTADO, MIGUEL PEÑALOZA, MOSTRADOR, MOSTRADOR TOR,
NEFTALI CRUZ, SALVADOR GARCIA, SERGIO PEREZ FLORES, WENDY SANCHEZ,
YATZMIN MONSIVAIS, YULIANA RODRIGUEZ
```

### asesoresServicio — Servicios (activos en 2026)
```
ALEJANDRO TRUJILLO, ADOLFO PACHECO
```

### asesoresAtencion — Atención al Cliente interno (solo Dir·Todos)
```
JONAS RODRIGUEZ
```

### Aliases en mapaAlias (Presupuesto → SAP)
```
JONATHAN ROCHE     → JONATHAN ROCHE TOR
JUAN DE DIOS LOPEZ → JUAN DE DIOS
KYMBERLY PORTILLO  → KIMBERLY PORTILLO
EDUARDO GONZALEZ   → EDUARDO CARRASCO
ADRIAN JIMENEZ     → JESUS ORTIZ
```

---

## Mapeo de grupos de artículos → División (mapaGrupos)

| GrupoArticulo SAP | División Dashboard |
|---|---|
| Identificacion | Trazabilidad |
| Herramienta de Ensamble | Trazabilidad |
| Vision | Visión |
| Robotica, Automatizacion | Robótica |
| Quimicos, Abrasivos, Artículos, Otros Suministros, Soldadura y Corte, Marcadores, Seguridad | Suministros |
| Servicio en Campo, Servicio Taller | Servicios |
| Vending Machines | ignorar (null) |

---

## Etapas de Oportunidades (DASHBOARD_CONFIG.etapas)

| Etapa | % | Color |
|---|---|---|
| Contacto Inicial | 10% | #B4B2A9 gris |
| Cotización | 25% | #BA7517 naranja |
| Pruebas / Demostración | 60% | #1D9E75 verde |
| Negociación | 80% | #185FA5 azul |
| Trámite con Compras | 90% | #5B2D8E morado |
| Factura | 95% | #639922 verde oscuro |

---

## Estado de validación de roles (barra de prueba TEST activa)

| Rol·División | Resumen | Pipeline | Asesores | En riesgo |
|---|---|---|---|---|
| Dir·Todos | ✅ ~$600K mayo, ~$8M Ene-May | ✅ todas divisiones | ⚠️ Pendiente confirmar todos | ✅ |
| Ger·Traz | ✅ números correctos | ✅ oportunidades y OVs solo Traz | ✅ solo Trazabilidad | ✅ |
| Ger·Sum | ⚠️ meta $0 (sin presupuesto) | ✅ | ⚠️ algunos en $0 (sin presupuesto) | ✅ |
| Líd·Traz | ✅ | ✅ | ✅ | ✅ |
| Ase·Traz | ✅ | ✅ | N/A | ✅ |
| Ase·Sum | ⚠️ meta $0 (sin presupuesto) | ✅ | N/A | ✅ |

---

## Pendientes para próxima sesión

### Alta prioridad
1. **Eliminar barra TEST** — una vez validados todos los roles con usuarios reales
2. **Presupuesto Suministros y Servicios** — agregar metas a `Presupuesto Ventas.xlsx` para que Ger·Sum y Ger·Servicios muestren % de cumplimiento
3. **Confirmar números Dir·Todos** — validar que el total empresa en Resumen cuadra con SAP después del último re-export
4. **Clientes en riesgo** — actualmente muestra 0; revisar umbral o ampliar rango de fechas en el archivo de ventas

### Media prioridad
5. **Acumulado anual vs 2025** — actualmente muestra -66%; revisar si el archivo de ventas tiene datos completos de 2025 o si se necesita re-exportar desde una fecha anterior
6. **Drill-down asesor desde Gerente** — actualmente muestra `alert()` placeholder; implementar vista detalle
7. **Racha de actividad** — hardcodeada; conectar con datos de visitas de Field App (SharePoint List "Visitas Field App")
8. **Meta anual $NaN** — ocurre cuando no hay presupuesto para todos los meses; verificar con datos completos de Suministros

### Baja prioridad
9. **Notas de crédito en OVs** — no aplica (las OVs son solo órdenes abiertas, no tienen notas)
10. **Conexión directa SAP** — cuando Vertrou resuelva el certificado SSL en `vinssa.vertrou.cloud:50000`, reemplazar archivos SharePoint con llamadas directas a Service Layer
11. **Replicar a asesores piloto** — Rafael Moyeda, Néstor Carranza primero; luego expandir

---

## Notas técnicas importantes

- **Formato de fechas:** El archivo de SharePoint debe tener la columna Fecha como **Texto** (`dd-MM-yyyy`) — si queda como tipo Date, Graph API la convierte a serial numérico y el parser falla
- **TotalFrgn para moneda:** Siempre usar `T2.TotalFrgn` en el query, nunca `LineTotal` directo — las facturas MXP con `DocRate=1` inflan artificialmente el total en USD
- **Notas de crédito:** El UNION ALL con ORIN y valores negativos es esencial — sin él los totales son inflados
- **Cache de Graph API:** Si los datos no se actualizan después de subir un archivo nuevo, usar en consola: `DASH_STATE.token = null; DASH_STATE.driveId = null; dashCargado = false; navSwitch('screen-dashboard', null)`
- **Barra de prueba:** Llamar `dashTestMode('gerente','Trazabilidad')` etc. desde consola para cambiar rol sin recargar
