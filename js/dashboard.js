// ── Vinssa Dashboard — js/dashboard.js ──────────────────────────────────────
// Lee roles de Lista Roles Dashboard.xlsx
// Lee ventas de Ventas Asesor.xlsx
// Lee presupuesto de Presupuesto Ventas.xlsx
// Todo vía Microsoft Graph API (mismo token que usa sap.js)

// ── Constantes SharePoint ────────────────────────────────────────────────────
const DASHBOARD_CONFIG = {
  siteId: 'versatilidadsaltillo.sharepoint.com,/sites/VINSSAAutomation',
  archivos: {
    roles:          'Lista Roles Dashboard.xlsx',
    ventas:         'Ventas Asesor v2.xlsx',
    ovs:            'OVs Asesor.xlsx',
    oportunidades:  'Oportunidades.xlsx',
    presupuesto:    'Presupuesto Ventas.xlsx'
  },
  // Mapeo GrupoArticulo SAP → División Dashboard
  mapaGrupos: {
    'Identificacion':          'Trazabilidad',
    'Herramienta de Ensamble': 'Trazabilidad',
    'Vision':                  'Visión',
    'Robotica':                'Robótica',
    'Automatizacion':          'Robótica',
    'Quimicos':                'Suministros',
    'Abrasivos':               'Suministros',
    'Artículos':               'Suministros',
    'Otros Suministros':       'Suministros',
    'Soldadura y Corte':       'Suministros',
    'Marcadores':              'Suministros',
    'Seguridad':               'Suministros',
    'Servicio en Campo':       'Servicios',
    'Servicio Taller':         'Servicios',
    'Vending Machines':        null  // ignorar
  },
  meses: ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'],
  // Etapas de oportunidades con orden, probabilidad y color
  etapas: {
    'Contacto Inicial':       { orden: 1, pct: 10, color: '#B4B2A9' },
    'Cotización':             { orden: 2, pct: 25, color: '#BA7517' },
    'Pruebas / Demostración': { orden: 3, pct: 60, color: '#1D9E75' },
    'Negociación':            { orden: 4, pct: 80, color: '#185FA5' },
    'Trámite con Compras':    { orden: 5, pct: 90, color: '#5B2D8E' },
    'Factura':                { orden: 6, pct: 95, color: '#639922' }
  },
  // Divisiones visibles por rol gerente (filtra pipeline)
  divisionesTrazabilidad: ['Trazabilidad', 'Visión', 'Robótica'],
  // Mapeo de valores del campo Linea en Oportunidades → División Dashboard
  // Cubre variantes de capitalización y nombres alternativos
  mapaLineas: {
    // Trazabilidad
    'identificacion': 'Trazabilidad', 'identificación': 'Trazabilidad',
    'telesis': 'Trazabilidad', 'marcaje': 'Trazabilidad', 'brady': 'Trazabilidad',
    'herramientas de ensamble': 'Trazabilidad', 'herramienta de ensamble': 'Trazabilidad',
    'atlas copco': 'Trazabilidad',
    // Visión
    'vision': 'Visión', 'visión': 'Visión', 'cognex': 'Visión',
    // Robótica
    'robotica': 'Robótica', 'robótica': 'Robótica',
    'robotica movil': 'Robótica', 'robótica móvil': 'Robótica',
    'automatizacion': 'Robótica', 'automatización': 'Robótica',
    'mir': 'Robótica', 'nabtesco': 'Robótica', 'ur': 'Robótica',
    // Suministros
    'quimicos': 'Suministros', 'químicos': 'Suministros',
    'abrasivos': 'Suministros', 'soldadura y corte': 'Suministros',
    'soldadura & corte': 'Suministros', 'soldadura': 'Suministros',
    'seguridad': 'Suministros', 'otros suministros': 'Suministros',
    'grasas': 'Suministros', 'equipos de dosificacion': 'Suministros',
    'epp': 'Suministros',
    // Servicios
    'servicio taller': 'Servicios', 'servicio en campo': 'Servicios',
    'servicios': 'Servicios'
  },
  // Asesores dedicados de Suministros — activos en 2026 (80%+ ventas en grupos Suministros)
  // Actualizar cuando haya cambios de personal
  asesoresSupply: new Set([
    'ALEJANDRO RODRIGUEZ','ARTURO CASTILLO','DANIEL ABASCAL','EDER MORENO',
    'EDUARDO ESPARZA','ENRIQUE CRUZ','FERNANDO MONTANA ALDAZ','FRANCISCO MARTINEZ',
    'HORACIO MORALES','JORGE OLAGUIVEL','LINO TOVAR','MARIA VILLALOBOS','MAYELA HURTADO','MOSTRADOR','MOSTRADOR TOR',
    'MIGUEL PEÑALOZA','NEFTALI CRUZ','SALVADOR GARCIA','SERGIO PEREZ FLORES','WENDY SANCHEZ',
    'YATZMIN MONSIVAIS','YULIANA RODRIGUEZ'
  ]),
  // Asesores dedicados de Servicios — activos en 2026
  asesoresServicio: new Set([
    'ALEJANDRO TRUJILLO','ADOLFO PACHECO'
  ]),
  // Asesores internos de atención al cliente — visibles solo en Dir·Todos
  asesoresAtencion: new Set([
    'JONAS RODRIGUEZ'
  ]),
  // Alias: nombre normalizado en Presupuesto → nombre normalizado en SAP
  mapaAlias: {
    'JONATHAN ROCHE':     'JONATHAN ROCHE TOR',
    'JUAN DE DIOS LOPEZ': 'JUAN DE DIOS',
    'KYMBERLY PORTILLO':  'KIMBERLY PORTILLO',
    'EDUARDO GONZALEZ':   'EDUARDO CARRASCO',
    'ADRIAN JIMENEZ':     'JESUS ORTIZ'
  }
};

// Estado global del dashboard
const DASH_STATE = {
  rol: null,          // 'asesor' | 'lider' | 'gerente'
  division: 'Todos',  // división del usuario — filtra pipeline y datos
  userEmail: null,
  ventas: [],         // facturas (OINV)
  ovs: [],            // órdenes de venta abiertas (ORDR)
  oportunidades: [],  // oportunidades CRM (OOPR)
  presupuesto: [],    // filas crudas del Excel de presupuesto
  mesActual: new Date().getMonth() + 1,
  anioActual: new Date().getFullYear(),
  mesSel: new Date().getMonth() + 1,
  anioSel: new Date().getFullYear(),
  token: null,
  driveId: null,
  asesoresSort: 'mas_ventas'  // 'mas_ventas' | 'menos_ventas' | 'mejor_pct' | 'peor_pct'
};

// ── Obtener token Graph ──────────────────────────────────────────────────────
async function dashGetToken() {
  if (DASH_STATE.token) return DASH_STATE.token;
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error('No hay sesión activa');
  const result = await acquireTokenSafe({
    scopes: ['Sites.ReadWrite.All', 'User.Read'],
    account: accounts[0]
  });
  DASH_STATE.token = result.accessToken;
  DASH_STATE.userEmail = accounts[0].username.toLowerCase();
  return DASH_STATE.token;
}

// ── Buscar archivo en SharePoint y obtener su driveItem ─────────────────────
async function dashGetFileId(token, nombre) {
  // Usar driveId cacheado si ya lo tenemos
  if (!DASH_STATE.driveId) {
    const url = 'https://graph.microsoft.com/v1.0/sites/versatilidadsaltillo.sharepoint.com:/sites/VINSSAAutomation:/drives';
    const drivesRes = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const drives = await drivesRes.json();
    if (!drives.value || !drives.value.length) throw new Error('No se encontró el drive de SharePoint');
    DASH_STATE.driveId = drives.value[0].id;
  }
  const driveId = DASH_STATE.driveId;
  const searchUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/root/children';
  const itemsRes = await fetch(searchUrl, { headers: { Authorization: 'Bearer ' + token } });
  const items = await itemsRes.json();
  const file = (items.value || []).find(f => f.name === nombre);
  if (!file) throw new Error('No se encontró el archivo: ' + nombre);
  return { driveId, fileId: file.id };
}

// ── Leer tabla de un Excel vía Graph ────────────────────────────────────────
async function dashLeerExcel(token, nombre) {
  try {
    const { driveId, fileId } = await dashGetFileId(token, nombre);
    // Leer la primera hoja como rango con valores
    const rangeUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/workbook/worksheets/Sheet1/usedRange`;
    const res = await fetch(rangeUrl, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      // Intentar con nombre de hoja en español
      const res2 = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/workbook/worksheets/Hoja1/usedRange`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!res2.ok) throw new Error(`Error leyendo ${nombre}`);
      const data2 = await res2.json();
      return excelValuesToObjects(data2.values);
    }
    const data = await res.json();
    return excelValuesToObjects(data.values);
  } catch(e) {
    console.error(`Error leyendo ${nombre}:`, e);
    return [];
  }
}

// Convierte array de arrays (primera fila = headers) a array de objetos
function excelValuesToObjects(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] ?? '');
      return obj;
    });
}

// ── Inicializar Dashboard ────────────────────────────────────────────────────
async function dashInit() {
  dashMostrarLoader(true);
  // Limpiar caché de driveId y token para forzar lectura fresca
  DASH_STATE.driveId = null;
  DASH_STATE.token = null;
  try {
    const token = await dashGetToken();

    // Leer archivos en paralelo — OVs y Oportunidades son opcionales
    const [roles, ventas, presupuesto] = await Promise.all([
      dashLeerExcel(token, DASHBOARD_CONFIG.archivos.roles),
      dashLeerExcel(token, DASHBOARD_CONFIG.archivos.ventas),
      dashLeerExcel(token, DASHBOARD_CONFIG.archivos.presupuesto)
    ]);
    let ovs = [], oportunidades = [];
    try {
      ovs = await dashLeerExcel(token, DASHBOARD_CONFIG.archivos.ovs);
    } catch(e) {
      console.warn('OVs Asesor.xlsx no encontrado:', e.message);
    }
    try {
      oportunidades = await dashLeerExcel(token, DASHBOARD_CONFIG.archivos.oportunidades);
    } catch(e) {
      console.warn('Oportunidades.xlsx no encontrado:', e.message);
    }

    // Determinar rol del usuario actual
    const emailActual = DASH_STATE.userEmail;
    const userRole = roles.find(r =>
      String(r.Email || '').toLowerCase().trim() === emailActual
    );
    DASH_STATE.rol = userRole ? String(userRole.Rol || '').toLowerCase().trim() : 'asesor';
    DASH_STATE.division = userRole ? String(userRole.Division || userRole.División || 'Todos').trim() : 'Todos';

    // Normalizar rol
    if (DASH_STATE.rol.includes('gerente')) DASH_STATE.rol = 'gerente';
    else if (DASH_STATE.rol.includes('lider') || DASH_STATE.rol.includes('líder')) DASH_STATE.rol = 'lider';
    else DASH_STATE.rol = 'asesor';

    DASH_STATE.ventas = ventas;
    DASH_STATE.ovs = ovs;
    DASH_STATE.oportunidades = oportunidades;
    DASH_STATE.presupuesto = presupuesto;

    dashRender();
  } catch(e) {
    console.error('Error inicializando dashboard:', e);
    dashMostrarError(e.message);
  } finally {
    dashMostrarLoader(false);
  }
}

// ── Obtener total en USD de una fila de ventas ───────────────────────────────
function dashGetTotal(v) {
  // Soporta columnas: Total, TotalUSD, TotalFrgn (incluye negativos para notas de crédito)
  const direct = parseFloat(v.Total !== undefined && v.Total !== '' ? v.Total : 
                            v.TotalUSD !== undefined && v.TotalUSD !== '' ? v.TotalUSD :
                            v.TotalFrgn !== undefined && v.TotalFrgn !== '' ? v.TotalFrgn : 0);
  if (direct !== 0) return direct;
  // Fallback: TotalMXP / TipoCambio
  const mxp = parseFloat(v.TotalMXP || 0);
  const tc  = parseFloat(v.TipoCambio || 1) || 1;
  return mxp !== 0 ? mxp / tc : 0;
}

// ── Normalizar nombre: quita acentos y pasa a mayúsculas ────────────────────
function dashNormNombre(str) {
  return String(str || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

// ── Normalizar nombre de presupuesto → SAP (aplica alias) ───────────────────
function dashNormPresup(str) {
  const norm = dashNormNombre(str);
  return DASHBOARD_CONFIG.mapaAlias[norm] || norm;
}

// ── Parser de fechas — soporta dd/MM/yyyy, yyyy-MM-dd y serial Excel ─────────
function dashParseFecha(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    // Serial date de Excel: días desde 1/1/1900
    // Usar UTC para evitar offset de zona horaria
    const ms = (val - 25569) * 86400 * 1000;
    const d = new Date(ms);
    // Retornar fecha local correcta usando componentes UTC
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  const s = String(val).trim();
  // dd-MM-yyyy (formato SAP con CONVERT 105)
  const m0 = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m0) return new Date(parseInt(m0[3]), parseInt(m0[2]) - 1, parseInt(m0[1]));
  // dd/MM/yyyy (formato mexicano con diagonal)
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) return new Date(parseInt(m1[3]), parseInt(m1[2]) - 1, parseInt(m1[1]));
  // yyyy-MM-dd
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return new Date(parseInt(m2[1]), parseInt(m2[2]) - 1, parseInt(m2[3]));
  // fallback
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

// ── Calcular métricas para un asesor/mes/año ─────────────────────────────────
function dashCalcMetricas(asesor, mes, anio, divisionesVisibles) {
  // asesor viene del presupuesto — aplicar alias para buscar en SAP
  const asesorNorm = dashNormPresup(asesor);
  const ventasFiltradas = DASH_STATE.ventas.filter(v => {
    if (!v.Fecha) return false;
    const fecha = dashParseFecha(v.Fecha);
    if (!fecha || isNaN(fecha)) return false;
    const vMes   = fecha.getMonth() + 1;
    const vAnio  = fecha.getFullYear();
    const vAsesor = dashNormNombre(v.Asesor);
    if (vMes !== mes || vAnio !== anio || vAsesor !== asesorNorm) return false;
    // Filtrar por división si aplica.
    // IMPORTANT: always include negative rows (credit notes) — they cancel sales
    // regardless of the GrupoArticulo they were categorized under in SAP.
    if (divisionesVisibles) {
      const total = dashGetTotal(v);
      if (total < 0) return true; // credit note — always include for correct netting
      const div = DASHBOARD_CONFIG.mapaGrupos[String(v.GrupoArticulo||'').trim()];
      return divisionesVisibles.includes(div);
    }
    return true;
  });

  // Agrupar por división usando el mapa de grupos
  const porDivision = {};
  ventasFiltradas.forEach(v => {
    const grupo = String(v.GrupoArticulo || '').trim();
    const division = DASHBOARD_CONFIG.mapaGrupos[grupo];
    if (!division) return; // ignorar (ej. Vending Machines)
    porDivision[division] = (porDivision[division] || 0) + dashGetTotal(v);
  });

  const totalVenta = ventasFiltradas.reduce((s, v) => s + dashGetTotal(v), 0);
  // totalVenta includes credit notes (negative) — this is correct net sales
  // Count unique positive OV/invoice numbers (exclude credit notes which are negative)
  const numOVs = new Set(ventasFiltradas.filter(v => dashGetTotal(v) > 0).map(v => v.NumFactura || v.NumOV)).size;

  // Presupuesto del asesor para ese mes
  // Buscar presupuesto: el nombre del asesor viene del presupuesto (puede tener alias)
  // El asesorNorm viene de SAP — necesitamos mapear presupuesto→SAP para comparar
  const presupFiltrado = DASH_STATE.presupuesto.filter(p => {
    const pAsesorSAP = dashNormPresup(p.Asesor); // convierte presupuesto→SAP norm
    const pMes = parseInt(p.Mes);
    // Año column may come with different encoding — try all variants
    const pAnioRaw = p['Año'] || p['Ano'] || p['A\u00f1o'] || p.Year || 0;
    const pAnio = parseInt(pAnioRaw);
    return pAsesorSAP === asesorNorm && pMes === mes && pAnio === anio;
  });

  // Meta total del asesor (suma de todas sus divisiones ese mes)
  const totalMeta = presupFiltrado.reduce((s, p) => {
    const m = parseFloat(p.Meta);
    return s + (isNaN(m) ? 0 : m);
  }, 0);

  // Meta por división
  const metaPorDivision = {};
  presupFiltrado.forEach(p => {
    const div = String(p.Division || p.División || '').trim();
    if (div) {
      const m = parseFloat(p.Meta);
      metaPorDivision[div] = (metaPorDivision[div] || 0) + (isNaN(m) ? 0 : m);
    }
  });

  const pct = totalMeta > 0 ? Math.round(totalVenta / totalMeta * 100) : 0;
  const falta = Math.max(0, totalMeta - totalVenta);
  const diasRestantes = mes === DASH_STATE.mesActual && anio === DASH_STATE.anioActual
    ? diasEnMes(mes, anio) - new Date().getDate()
    : 0;
  const ritmo = diasRestantes > 0 ? falta / diasRestantes : 0;

  return {
    totalVenta, totalMeta, pct, falta, diasRestantes,
    ritmo, numOVs, porDivision, metaPorDivision
  };
}

// ── Calcular histórico anual de un asesor ────────────────────────────────────
function dashHistoricoAnual(asesor, anio, divisionesVisibles) {
  const meses = [];
  for (let m = 1; m <= 12; m++) {
    const met = dashCalcMetricas(asesor, m, anio, divisionesVisibles);
    const metAnterior = dashCalcMetricas(asesor, m, anio - 1, divisionesVisibles);
    meses.push({
      mes: m,
      label: DASHBOARD_CONFIG.meses[m-1],
      venta: met.totalVenta,
      meta: met.totalMeta,
      ventaAnterior: metAnterior.totalVenta,
      pct: met.pct
    });
  }
  const acumulado = meses.reduce((s, m) => s + m.venta, 0);
  const acumuladoAnterior = meses.reduce((s, m) => s + m.ventaAnterior, 0);
  const mejorMes = meses.reduce((best, m) => m.venta > best.venta ? m : best, meses[0]);
  const metaAnual = meses.reduce((s, m) => s + m.meta, 0);
  // Proyección: promedio de meses con datos × 12
  const mesesConDatos = meses.filter(m => m.venta > 0);
  const proyeccion = mesesConDatos.length > 0
    ? (acumulado / mesesConDatos.length) * 12
    : 0;
  return { meses, acumulado, acumuladoAnterior, mejorMes, metaAnual, proyeccion };
}

// ── Obtener lista única de asesores con presupuesto (nombres SAP normalizados) ─
function dashGetAsesores() {
  // Returns the list of asesores to show in lider/gerente views.
  // Trazabilidad asesores come from the presupuesto file (authoritative).
  // Suministros asesores come from a hardcoded list derived from sales data.
  // Dir·Todos shows the union of both.
  const divs = dashGetDivisionesVisibles(); // null = Todos

  const seen = new Set();
  const result = [];

  const addFromPresupuesto = (filterDivs) => {
    DASH_STATE.presupuesto.forEach(p => {
      const nombre = String(p.Asesor || '').trim();
      if (!nombre || seen.has(nombre)) return;
      if (filterDivs) {
        const pDiv = String(p.Division || '').trim();
        if (!filterDivs.includes(pDiv)) return;
      }
      seen.add(nombre); result.push(nombre);
    });
  };

  const addFromSupplyList = () => {
    DASHBOARD_CONFIG.asesoresSupply.forEach(nombre => {
      if (!seen.has(nombre)) { seen.add(nombre); result.push(nombre); }
    });
  };

  const addFromServicioList = () => {
    DASHBOARD_CONFIG.asesoresServicio.forEach(nombre => {
      if (!seen.has(nombre)) { seen.add(nombre); result.push(nombre); }
    });
  };

  const addFromAtencionList = () => {
    DASHBOARD_CONFIG.asesoresAtencion.forEach(nombre => {
      if (!seen.has(nombre)) { seen.add(nombre); result.push(nombre); }
    });
  };

  if (!divs) {
    // Dir·Todos: all divisions + internal atencion al cliente
    addFromPresupuesto(null);   // Trazabilidad
    addFromSupplyList();         // Suministros
    addFromServicioList();       // Servicios
    addFromAtencionList();       // Atencion al cliente interno
  } else if (divs.includes('Suministros')) {
    addFromSupplyList();
  } else if (divs.includes('Servicios')) {
    addFromServicioList();
  } else {
    // Trazabilidad (or Visión/Robótica subset)
    addFromPresupuesto(divs);
  }

  return result;
}

// ── Calcular clientes en riesgo ──────────────────────────────────────────────
function dashClientesEnRiesgo(asesor, diasUmbral = 60) {
  const asesorNorm = asesor ? dashNormPresup(asesor) : null;
  const ultimaCompra = {};
  DASH_STATE.ventas
    .filter(v => !asesorNorm || dashNormNombre(v.Asesor) === asesorNorm)
    .forEach(v => {
      const key = v.CardCode;
      const fecha = dashParseFecha(v.Fecha);
      if (!fecha || isNaN(fecha)) return;
      if (!ultimaCompra[key] || fecha > ultimaCompra[key].fecha) {
        ultimaCompra[key] = {
          cardCode: v.CardCode,
          cliente: String(v.Cliente || '').trim(),
          asesor: String(v.Asesor || '').trim(),
          fecha,
          total: dashGetTotal(v),
          grupo: String(v.GrupoArticulo || '').trim()
        };
      }
    });

  const hoy = new Date();
  return Object.values(ultimaCompra)
    .map(c => ({
      ...c,
      dias: Math.floor((hoy - c.fecha) / (1000 * 60 * 60 * 24))
    }))
    .filter(c => c.dias >= diasUmbral)
    .sort((a, b) => b.dias - a.dias);
}

// ── Formatear moneda ─────────────────────────────────────────────────────────
function dashFmt(num) {
  const abs = Math.abs(num);
  const sign = num < 0 ? '-' : '';
  if (abs >= 1000000) return sign + '$' + (abs/1000000).toFixed(1) + 'M';
  if (abs >= 1000)    return sign + '$' + (abs/1000).toFixed(1) + 'K';
  return sign + '$' + Math.round(abs).toLocaleString('es-MX');
}

function diasEnMes(mes, anio) {
  return new Date(anio, mes, 0).getDate();
}

// ── Mensaje contextual según % de cumplimiento ───────────────────────────────
function dashMensajeCtx(pct, falta, diasRestantes, ritmo, isCurrent) {
  if (!isCurrent) {
    if (pct >= 100) return { tipo: 'ok', msg: `Cerraste en meta con <strong>${pct}%</strong> — excelente mes.` };
    if (pct >= 80)  return { tipo: 'info', msg: `Cerraste en <strong>${pct}%</strong> — cerca de la meta.` };
    return { tipo: 'warn', msg: `Cerraste en <strong>${pct}%</strong> de meta ese mes.` };
  }
  if (pct >= 100) return { tipo: 'ok', msg: `¡Ya alcanzaste la meta! Vas en <strong>${pct}%</strong>. Cada venta adicional suma al siguiente mes.` };
  if (pct >= 75)  return { tipo: 'info', msg: `Vas en <strong>${pct}%</strong> — buen ritmo. Necesitas <strong>${dashFmt(ritmo)}/día</strong> para cerrar en meta.` };
  if (pct >= 50)  return { tipo: 'warn', msg: `Vas en <strong>${pct}%</strong> a mitad de mes. Necesitas <strong>${dashFmt(ritmo)}/día</strong> — revisa tus oportunidades en negociación.` };
  return { tipo: 'warn', msg: `Vas en <strong>${pct}%</strong> — requiere atención. Necesitas <strong>${dashFmt(ritmo)}/día</strong> en los ${diasRestantes} días restantes.` };
}

// ── Colores semáforo ─────────────────────────────────────────────────────────
function dashColor(pct) {
  if (pct >= 100) return { bar: '#639922', pct: '#3B6D11', pill: 'pill-grn' };
  if (pct >= 65)  return { bar: '#185FA5', pct: '#185FA5', pill: 'pill-blu' };
  if (pct >= 40)  return { bar: '#EF9F27', pct: '#BA7517', pill: 'pill-amb' };
  return { bar: '#D85A30', pct: '#A32D2D', pill: 'pill-red' };
}

// ── Obtener divisiones visibles para el usuario actual ───────────────────────
function dashGetDivisionesVisibles() {
  const div = String(DASH_STATE.division || 'Todos').trim();
  if (div === 'Todos' || div === '' || div === 'todos') return null; // null = ver todo
  // Normalize accents
  const divNorm = div.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  if (divNorm === 'trazabilidad') return ['Trazabilidad', 'Visión', 'Robótica'];
  if (divNorm === 'suministros') return ['Suministros'];
  if (divNorm === 'servicios') return ['Servicios'];
  return null; // default: ver todo
}

// Verificar si un grupo de artículo es visible para el usuario
function dashGrupoVisible(grupoArticulo) {
  const divs = dashGetDivisionesVisibles();
  if (!divs) return true; // ver todo
  const div = DASHBOARD_CONFIG.mapaGrupos[String(grupoArticulo || '').trim()];
  return divs.includes(div);
}

// ── Render principal ─────────────────────────────────────────────────────────
// Entry point called after data loads. Routes to the correct view based on rol.
function dashRender() {
  const container = document.getElementById('dashboard-container');
  if (!container) return;

  const rol  = DASH_STATE.rol;
  const mes  = DASH_STATE.mesSel;
  const anio = DASH_STATE.anioSel;
  const isCurrent = mes === DASH_STATE.mesActual && anio === DASH_STATE.anioActual;
  const mesLabel  = DASHBOARD_CONFIG.meses[mes-1] + ' ' + anio;
  const divs      = dashGetDivisionesVisibles(); // null = Todos

  if (rol === 'asesor') {
    // Asesor: find their SAP name from ventas data
    const asesorNombre = _findAsesorNombre(DASH_STATE.userEmail);
    _renderAsesor(container, asesorNombre, mes, anio, isCurrent, mesLabel, divs);
  } else if (rol === 'lider') {
    _renderLider(container, mes, anio, isCurrent, mesLabel, divs);
  } else {
    // gerente or director
    _renderGerente(container, mes, anio, isCurrent, mesLabel, divs);
  }
}

// Find asesor name from roles file by email
function _findAsesorNombre(email) {
  // Look up their name in the presupuesto by matching email in roles
  // For now use the display name from MSAL
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length) return accounts[0].name || accounts[0].username;
  return email;
}

// ── SHARED: Period selector ──────────────────────────────────────────────────
function _periodSelector(mesLabel, isCurrent) {
  const meses     = DASHBOARD_CONFIG.meses;
  const mesActual = DASH_STATE.mesActual;
  return `
    <div class="dash-period-selector">
      <button class="dash-period-btn" onclick="dashTogglePeriodPanel()">
        <span>📅 ${mesLabel}</span><span style="font-size:10px">▾</span>
      </button>
      <div class="dash-period-badge" style="${isCurrent ? '' : 'background:#E6F1FB;color:#0C447C'}">
        ${isCurrent ? 'Mes actual' : 'Histórico'}
      </div>
    </div>
    <div class="dash-period-panel" id="dash-period-panel">
      <div class="dash-month-grid">
        ${meses.map((m, i) => {
          const num = i + 1;
          const futuro = num > mesActual && DASH_STATE.anioSel >= DASH_STATE.anioActual;
          const sel    = num === DASH_STATE.mesSel;
          return `<button class="dash-mg-btn${sel?' active':''}${futuro?' disabled':''}"
            ${futuro ? 'disabled' : `onclick="dashSelMes(${num},this)"`}
          >${m}${num===mesActual?' ●':''}</button>`;
        }).join('')}
      </div>
    </div>`;
}

// ── SHARED: Annual summary block ─────────────────────────────────────────────
function _anualHtml(hist, anio, mes) {
  const pctAnual = hist.metaAnual > 0 ? Math.round(hist.acumulado / hist.metaAnual * 100) : 0;
  const vsAnt    = hist.acumuladoAnterior > 0
    ? Math.round((hist.acumulado - hist.acumuladoAnterior) / hist.acumuladoAnterior * 100) : 0;
  const label    = `Ene–${DASHBOARD_CONFIG.meses[mes-1]}`;
  return `
    <div class="dash-annual">
      <div class="dash-annual-title">Resumen ${anio}</div>
      <div class="dash-annual-row">
        <div class="dash-ar-lbl">Acumulado ${label}</div>
        <div style="text-align:right">
          <div class="dash-ar-val">${dashFmt(hist.acumulado)}</div>
          <div class="dash-ar-sub ${vsAnt>=0?'dash-up':'dash-dn'}">${vsAnt>=0?'+':''}${vsAnt}% vs ${anio-1}</div>
        </div>
      </div>
      <div class="dash-annual-row">
        <div class="dash-ar-lbl">Meta anual</div>
        <div style="text-align:right">
          <div class="dash-ar-val">${dashFmt(hist.metaAnual)}</div>
          <div class="dash-ar-sub dash-nu">${pctAnual}% alcanzado</div>
        </div>
      </div>
      ${hist.mejorMes && hist.mejorMes.venta > 0 ? `
      <div class="dash-annual-row">
        <div class="dash-ar-lbl">Mejor mes · ${hist.mejorMes.label}</div>
        <div style="text-align:right">
          <div class="dash-ar-val">${dashFmt(hist.mejorMes.venta)}</div>
          <div class="dash-ar-sub" style="color:#BA7517">${hist.mejorMes.pct>=100?'🏆 ':''}${hist.mejorMes.pct}% de meta</div>
        </div>
      </div>` : ''}
      <div class="dash-annual-row">
        <div class="dash-ar-lbl">Proyección anual</div>
        <div style="text-align:right">
          <div class="dash-ar-val">${dashFmt(hist.proyeccion)}</div>
          <div class="dash-ar-sub ${hist.proyeccion>=hist.metaAnual?'dash-up':'dash-dn'}">
            ${hist.metaAnual>0 ? Math.round(hist.proyeccion/hist.metaAnual*100)+'% vs meta' : ''}
          </div>
        </div>
      </div>
      <canvas id="dash-chart-hist" style="margin-top:12px;width:100%;height:110px;max-height:110px"></canvas>
    </div>`;
}

// ── SHARED: Goal card ─────────────────────────────────────────────────────────
function _goalCard(venta, meta, pct, falta, diasRestantes, isCurrent, barId) {
  const col = dashColor(pct);
  return `
    <div class="dash-goal-card">
      <div class="dash-goal-top">
        <div class="dash-goal-amt">${dashFmt(venta)}</div>
        <div class="dash-goal-tgt">meta ${dashFmt(meta)}</div>
      </div>
      <div class="dash-pbar">
        <div class="dash-pfill" id="${barId}" style="width:0%;background:${col.bar}"></div>
      </div>
      <div class="dash-goal-row">
        <div class="dash-gs"><span class="dash-v" style="color:${col.pct}">${pct}%</span><span class="dash-l">alcanzado</span></div>
        <div class="dash-gs" style="text-align:center"><span class="dash-v">${dashFmt(Math.max(0,falta))}</span><span class="dash-l">${isCurrent?'te falta':'faltó'}</span></div>
        <div class="dash-gs" style="text-align:right"><span class="dash-v">${isCurrent&&diasRestantes>0?diasRestantes+' días':'Cerrado'}</span><span class="dash-l">${isCurrent&&diasRestantes>0?'restantes':''}</span></div>
      </div>
    </div>`;
}

// ── SHARED: Asesor card ───────────────────────────────────────────────────────
function _asesorCard(a, clickable) {
  const col = dashColor(a.pct);
  const slug = dashSlug(a.nombre);
  const ini  = a.nombre.split(' ').slice(0,2).map(n=>n[0]||'').join('').toUpperCase();
  const statusLabel = a.pct>=65?'En meta':a.pct>=40?'Seguimiento':'En riesgo';
  const statusClass = a.pct>=65?'pill-grn':a.pct>=40?'pill-amb':'pill-red';
  return `
    <div class="dash-ac ${a.pct<40?'dash-ac-risk':a.pct<65?'dash-ac-warn':''}"
      data-venta="${a.totalVenta}" data-pct="${a.pct}" data-haspresup="${a.totalMeta>0?1:0}"
      ${clickable?`onclick="dashDrillAsesor('${a.nombre}')"`:''} style="margin-bottom:8px">
      <div class="dash-ac-top">
        <div class="dash-acav" style="background:${col.bar}22;color:${col.pct}">${ini}</div>
        <div style="flex:1">
          <div class="dash-ac-name">${a.nombre}</div>
          <div class="dash-ac-sub">${dashFmt(a.totalVenta)} / ${dashFmt(a.totalMeta)}</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:15px;font-weight:500;color:${col.pct}">${a.pct}%</div>
          <span class="dash-pill ${statusClass}">${statusLabel}</span>
        </div>
      </div>
      <div class="dash-pbar" style="margin-bottom:8px">
        <div class="dash-pfill" id="dash-fill-${slug}" style="width:0%;background:${col.bar}"></div>
      </div>
      <div class="dash-ac-kpis">
        <div class="dash-ac-kpi"><span>${a.numOVs}</span>OVs</div>
        <div class="dash-ac-kpi"><span>${dashFmt(a.totalMeta)}</span>meta</div>
        <div class="dash-ac-kpi"><span>${dashFmt(Math.max(0,a.falta))}</span>falta</div>
        <div class="dash-ac-kpi"><span>${a.diasRestantes>0?a.diasRestantes+'d':'—'}</span>días</div>
      </div>
    </div>`;
}

// ── Ordenar equipo por venta o % de presupuesto ───────────────────────────────
// Para 'mejor_pct' y 'peor_pct': los asesores sin presupuesto (totalMeta=0) van
// al final — no deben "ganar" Peor % sólo porque su pct se calcula como 0.
function _ordenarEquipo(equipo, key) {
  const arr = equipo.slice();
  if (key === 'menos_ventas') {
    arr.sort((a, b) => a.totalVenta - b.totalVenta);
  } else if (key === 'mejor_pct' || key === 'peor_pct') {
    arr.sort((a, b) => {
      const aTiene = a.totalMeta > 0, bTiene = b.totalMeta > 0;
      if (aTiene !== bTiene) return aTiene ? -1 : 1;
      if (!aTiene) return 0;
      return key === 'mejor_pct' ? b.pct - a.pct : a.pct - b.pct;
    });
  } else {
    arr.sort((a, b) => b.totalVenta - a.totalVenta); // 'mas_ventas' default
  }
  return arr;
}

// ── SHARED: Asesores filter list ─────────────────────────────────────────────
function _asesoresListaHtml(equipo, clickable) {
  const sort = DASH_STATE.asesoresSort || 'mas_ventas';
  const ordenado = _ordenarEquipo(equipo, sort);
  return `
    <div class="dash-filter-row">
      <button class="dash-fchip active" onclick="dashFiltrarAsesores('todos',this)">Todos</button>
      <button class="dash-fchip" onclick="dashFiltrarAsesores('riesgo',this)">En riesgo</button>
      <button class="dash-fchip" onclick="dashFiltrarAsesores('meta',this)">En meta</button>
      <select class="dash-sort" aria-label="Ordenar por" onchange="dashOrdenarAsesores(this.value)">
        <option value="" disabled>Ordenar por</option>
        <option value="mas_ventas"   ${sort==='mas_ventas'  ?'selected':''}>Más ventas</option>
        <option value="menos_ventas" ${sort==='menos_ventas'?'selected':''}>Menos ventas</option>
        <option value="mejor_pct"    ${sort==='mejor_pct'   ?'selected':''}>Mejor % de presupuesto</option>
        <option value="peor_pct"     ${sort==='peor_pct'    ?'selected':''}>Peor % de presupuesto</option>
      </select>
    </div>
    <div id="dash-asesores-lista">
      ${ordenado.map(a => _asesorCard(a, clickable)).join('')}
    </div>
    <div style="height:16px"></div>`;
}

// ── SHARED: En riesgo ─────────────────────────────────────────────────────────
function _riesgoHtml(clientes, mostrarAsesor) {
  if (!clientes.length) return `
    <div style="padding:24px 16px;text-align:center;color:var(--color-text-secondary);font-size:13px">
      ✅ No hay clientes sin compra en más de 60 días.
    </div>`;
  return `
    <div class="dash-sec" style="padding-top:14px">
      <div class="dash-lbl">Clientes sin compra reciente</div>
      <div class="dash-ctx dash-ctx-warn" style="margin:0 0 4px">
        <div class="dash-cdot"></div>
        <p><strong>${clientes.length} clientes</strong> sin movimiento en +60 días.</p>
      </div>
    </div>
    <div class="dash-risk-list">
      ${clientes.slice(0,30).map(c=>`
        <div class="dash-rki">
          <div class="dash-rkdot" style="background:${c.dias>90?'#E24B4A':'#EF9F27'}"></div>
          <div class="dash-rk-info">
            <div class="dash-rk-name">${c.cliente}</div>
            <div class="dash-rk-sub">Último: ${c.fecha.toLocaleDateString('es-MX')} · ${dashFmt(c.total)}
              ${mostrarAsesor?`<br>Asesor: ${c.asesor}`:''}
            </div>
          </div>
          <div style="text-align:right;flex-shrink:0">
            <div style="font-size:13px;font-weight:500">${dashFmt(c.total)}</div>
            <div style="font-size:11px;color:${c.dias>90?'#A32D2D':'#BA7517'};margin-top:2px">${c.dias}d</div>
          </div>
        </div>`).join('')}
    </div>`;
}

// ── SHARED: Pipeline (oportunidades + OVs) ───────────────────────────────────
// divisionesVisibles: null = all, array = filter by those divisions
function _pipelineHtml(asesorNorm, divisionesVisibles) {
  const hoy = new Date();
  const etapasConfig = DASHBOARD_CONFIG.etapas;

  // ── Oportunidades ──────────────────────────────────────────────────────────
  const opps = DASH_STATE.oportunidades.filter(o => {
    if (asesorNorm && dashNormNombre(o.Asesor) !== asesorNorm) return false;
    if (divisionesVisibles) {
      const linea = String(o.Linea || '').trim().toLowerCase();
      const div = DASHBOARD_CONFIG.mapaLineas[linea];
      return div ? divisionesVisibles.includes(div) : false;
    }
    return true;
  });

  // Group by etapa, sort each group by monto desc
  const porEtapa = {};
  opps.forEach(o => {
    const etapa = String(o.Etapa || 'Sin etapa').trim();
    if (!porEtapa[etapa]) porEtapa[etapa] = [];
    const monto = parseFloat(o.MontoEstimado || 0);
    const pond  = parseFloat(o.MontoPonderado || 0);
    const fAp   = dashParseFecha(o.FechaApertura);
    const dias  = fAp ? Math.floor((hoy-fAp)/(1000*60*60*24)) : null;
    porEtapa[etapa].push({ ...o, monto, pond, diasAbierta: dias });
  });
  Object.values(porEtapa).forEach(arr => arr.sort((a,b) => b.monto - a.monto));

  const totalMonto = opps.reduce((s,o) => s + parseFloat(o.MontoEstimado||0), 0);
  const totalPond  = opps.reduce((s,o) => s + parseFloat(o.MontoPonderado||0), 0);

  // Render etapas in defined order, then any extra
  const etapasOrden = Object.keys(etapasConfig);
  const otrasEtapas = Object.keys(porEtapa).filter(e => !etapasOrden.includes(e));
  const todasEtapas = [...etapasOrden, ...otrasEtapas].filter(e => porEtapa[e]?.length > 0);

  const oppsHtml = todasEtapas.map(etapa => {
    const cfg  = etapasConfig[etapa] || { pct: 10, color: '#888' };
    const data = porEtapa[etapa];
    const total = data.reduce((s,o)=>s+o.monto,0);
    const pond  = data.reduce((s,o)=>s+o.pond,0);
    const pct   = totalMonto > 0 ? Math.round(total/totalMonto*100) : 0;
    const panelId = 'ep-' + etapa.replace(/[^a-zA-Z0-9]/g,'-');
    return `
      <div style="border-bottom:0.5px solid var(--color-border-tertiary)">
        <div style="display:flex;align-items:center;gap:10px;padding:10px 0;cursor:pointer"
          onclick="const p=document.getElementById('${panelId}');p.style.display=p.style.display==='none'?'block':'none';this.querySelector('.dchev').classList.toggle('rotated')">
          <div style="width:8px;height:8px;border-radius:50%;background:${cfg.color};flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-size:13px;font-weight:500">${etapa} · ${cfg.pct}%</div>
            <div style="font-size:11px;color:var(--color-text-secondary)">${data.length} oportunidades</div>
            <div style="height:3px;background:var(--color-border-tertiary);border-radius:2px;margin-top:4px;overflow:hidden">
              <div style="height:100%;width:${Math.min(pct,100)}%;background:${cfg.color};border-radius:2px"></div>
            </div>
          </div>
          <div style="text-align:right;flex-shrink:0">
            <div style="font-size:13px;font-weight:500">${dashFmt(total)}</div>
            <div style="font-size:11px;color:var(--color-text-secondary)">${dashFmt(pond)} pond.</div>
          </div>
          <div class="dchev" style="font-size:12px;color:var(--color-text-tertiary)">▾</div>
        </div>
        <div id="${panelId}" style="display:none;max-height:300px;overflow-y:auto;-webkit-overflow-scrolling:touch">
          ${data.map(o => `
            <div style="display:flex;gap:8px;padding:7px 0 7px 16px;border-top:0.5px solid var(--color-border-tertiary)">
              <div style="flex:1;min-width:0">
                <div style="font-size:12px;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${o.CardName||o.Cliente||''}</div>
                <div style="font-size:11px;color:var(--color-text-secondary);margin-top:2px">
                  ${o.Linea?o.Linea+' · ':''}${o.Marca?o.Marca+' · ':''}Cierre: ${o.FechaCierre||'—'}
                  ${!asesorNorm?' · '+o.Asesor:''}
                  ${o.diasAbierta!==null?` · <span style="color:${o.diasAbierta>90?'#A32D2D':o.diasAbierta>30?'#BA7517':'var(--color-text-tertiary)'}">${o.diasAbierta}d abierta</span>`:''}
                </div>
              </div>
              <div style="text-align:right;flex-shrink:0">
                <div style="font-size:12px;font-weight:500">${dashFmt(o.monto)}</div>
                <div style="font-size:10px;color:var(--color-text-secondary)">${o.Probabilidad||cfg.pct}%</div>
              </div>
            </div>`).join('')}
        </div>
      </div>`;
  }).join('');

  // ── OVs abiertas ─────────────────────────────────────────────────────────
  const ovsFiltradas = DASH_STATE.ovs.filter(v => {
    if (asesorNorm && dashNormNombre(v.Asesor) !== asesorNorm) return false;
    if (divisionesVisibles) {
      const div = DASHBOARD_CONFIG.mapaGrupos[String(v.GrupoArticulo||'').trim()];
      return divisionesVisibles.includes(div);
    }
    return true;
  });

  // Consolidate by OV number
  const ovsMap = {};
  ovsFiltradas.forEach(v => {
    const k = String(v.NumOV || v.NumFactura || '');
    if (!ovsMap[k]) {
      const fe = dashParseFecha(v.FechaEntrega);
      const dias = fe ? Math.floor((fe-hoy)/(1000*60*60*24)) : null;
      const sem  = dias===null?'ok':dias<0?'vencida':dias<=7?'urgente':'ok';
      ovsMap[k] = {
        numOV: k, cliente: String(v.Cliente||'').trim(),
        asesor: String(v.Asesor||'').trim(),
        fechaEntrega: v.FechaEntrega||'', dias, sem,
        division: DASHBOARD_CONFIG.mapaGrupos[String(v.GrupoArticulo||'').trim()]||'Otros',
        total: 0
      };
    }
    ovsMap[k].total += dashGetTotal(v);
  });

  const ovsArr = Object.values(ovsMap).sort((a,b) => {
    // Sort: vencidas first (most overdue), then urgentes, then ok by date
    const ord = {vencida:0, urgente:1, ok:2};
    if (ord[a.sem] !== ord[b.sem]) return ord[a.sem] - ord[b.sem];
    const da = a.dias ?? 9999, db = b.dias ?? 9999;
    return da - db;
  });

  const totalOVs   = ovsArr.reduce((s,v)=>s+v.total,0);
  const vencidas   = ovsArr.filter(o=>o.sem==='vencida').length;
  const urgentes   = ovsArr.filter(o=>o.sem==='urgente').length;

  const ovsHtml = !ovsArr.length
    ? `<div style="padding:16px;text-align:center;color:var(--color-text-secondary);font-size:13px">No hay OVs abiertas.</div>`
    : ovsArr.map(ov => {
        const sc = ov.sem==='vencida'?'#E24B4A':ov.sem==='urgente'?'#EF9F27':'#3B6D11';
        const sl = ov.sem==='vencida'?`⚠️ Vencida hace ${Math.abs(ov.dias)}d`
                  :ov.sem==='urgente'?`🔔 Entrega en ${ov.dias}d`
                  :ov.fechaEntrega?`Entrega: ${ov.fechaEntrega}`:'';
        return `
          <div style="display:flex;align-items:flex-start;gap:10px;padding:9px 0;border-bottom:0.5px solid var(--color-border-tertiary)">
            <div style="width:6px;height:6px;border-radius:50%;background:${sc};flex-shrink:0;margin-top:5px"></div>
            <div style="flex:1;min-width:0">
              <div style="font-size:13px;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${ov.cliente}</div>
              <div style="font-size:11px;color:var(--color-text-secondary);margin-top:1px">
                OV #${ov.numOV} · ${ov.division}${!asesorNorm?' · '+ov.asesor:''}
              </div>
              ${sl?`<div style="font-size:11px;color:${sc};font-weight:500;margin-top:1px">${sl}</div>`:''}
            </div>
            <div style="font-size:13px;font-weight:500;flex-shrink:0">${dashFmt(ov.total)}</div>
          </div>`;
      }).join('');

  return `
    <div style="padding:14px 0 80px">
      <!-- OPORTUNIDADES -->
      <div style="padding:0 16px 4px;display:flex;justify-content:space-between;align-items:baseline">
        <div>
          <div style="font-size:20px;font-weight:600">${dashFmt(totalMonto)}</div>
          <div style="font-size:12px;color:var(--color-text-secondary)">pipeline oportunidades · ${opps.length} abiertas</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:14px;font-weight:500">${dashFmt(totalPond)}</div>
          <div style="font-size:11px;color:var(--color-text-secondary)">ponderado</div>
        </div>
      </div>
      <div style="padding:0 16px">
        ${opps.length===0
          ? `<div style="padding:16px 0;text-align:center;color:var(--color-text-secondary);font-size:13px">No hay oportunidades abiertas.</div>`
          : oppsHtml}
      </div>

      <!-- POR MES (Cierre proyectado) — embebida entre Oportunidades y OVs -->
      ${_pipelineMensualHtml(asesorNorm, divisionesVisibles)}

      <!-- OVS -->
      <div style="padding:16px 16px 4px;display:flex;justify-content:space-between;align-items:baseline;margin-top:8px;border-top:2px solid var(--color-border-tertiary)">
        <div>
          <div style="font-size:20px;font-weight:600">${dashFmt(totalOVs)}</div>
          <div style="font-size:12px;color:var(--color-text-secondary)">OVs comprometidas · ${ovsArr.length} abiertas</div>
        </div>
        ${vencidas||urgentes?`
        <div style="text-align:right">
          ${vencidas?`<div style="font-size:11px;color:#A32D2D;font-weight:500">⚠️ ${vencidas} vencidas</div>`:''}
          ${urgentes?`<div style="font-size:11px;color:#BA7517;font-weight:500">🔔 ${urgentes} esta semana</div>`:''}
        </div>`:''}
      </div>
      <div style="max-height:400px;overflow-y:scroll;-webkit-overflow-scrolling:touch;padding:0 16px;border:0.5px solid var(--color-border-tertiary);border-radius:8px;margin:0 16px">
        ${ovsHtml}
      </div>
    </div>`;
}

// ── PIPELINE POR MES ─────────────────────────────────────────────────────────
// Vista por fecha de cierre proyectada. Bucket "Vencidas" (FechaCierre < hoy) +
// próximos 6 meses (mes actual + 5 siguientes). Barras horizontales apiladas con
// segmentos por etapa. Solo etapas progresivas relevantes para pipeline futuro.

// Etapas visibles en este tab. Las 6 etapas canónicas en orden progresivo (10% → 95%).
// Razón de incluir Contacto Inicial: ver opps en CI con fecha proyectada este mes es
// señal para que el gerente cuestione al asesor ("¿realmente cierra este mes?"). Razón
// de incluir Factura: confirma cierres en curso. Excluye etapas raras (Contacto WEB,
// Registro de Lead Expos, etc.) que añaden ruido visual sin volumen. El orden define
// el orden de los segmentos dentro de cada barra.
const _PM_ETAPAS_VISIBLES = ['Contacto Inicial', 'Cotización', 'Pruebas / Demostración', 'Negociación', 'Trámite con Compras', 'Factura'];
const _PM_COLOR_VENCIDA = '#E24B4A';
const _PM_COLOR_SIN_FECHA = '#9CA39E'; // gris medio — distinto del rojo de Vencidas y del gris claro de meses vacíos
const _PM_MESES_FUTURO = 5; // mes actual + 5 = 6 meses totales

// Agrupa STATE.oportunidades filtradas por rol en buckets {vencidas, [mes1..mesN]}.
// Cada bucket trae las opps separadas por etapa (solo etapas visibles).
function _pipelineMensualBuckets(asesorNorm, divisionesVisibles) {
  const hoy = new Date();
  const hoyInicio = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());

  // Filtro por rol — mismo patrón que _pipelineHtml.
  const oppsFiltradas = DASH_STATE.oportunidades.filter(o => {
    if (asesorNorm && dashNormNombre(o.Asesor) !== asesorNorm) return false;
    if (divisionesVisibles) {
      const linea = String(o.Linea || '').trim().toLowerCase();
      const div = DASHBOARD_CONFIG.mapaLineas[linea];
      return div ? divisionesVisibles.includes(div) : false;
    }
    return true;
  });

  // Buckets vacíos: sin fecha + vencidas + próximos N meses (mes actual + N futuros).
  // Orden visual: sin fecha (más arriba) → vencidas → meses futuros. Sin fecha y vencidas
  // se eliminan si quedan vacíos tras el bucketing — ver final de la función.
  const buckets = [];
  buckets.push({ id: 'sin-fecha', label: 'Sin fecha proyectada', esSinFecha: true, opps: [], porEtapa: {} });
  buckets.push({ id: 'vencidas', label: '⚠ Vencidas', esVencidas: true, opps: [], porEtapa: {} });
  for (let i = 0; i <= _PM_MESES_FUTURO; i++) {
    const d = new Date(hoy.getFullYear(), hoy.getMonth() + i, 1);
    const mes = d.getMonth();
    const anio = d.getFullYear();
    buckets.push({
      id: `m-${anio}-${mes+1}`,
      label: `${DASHBOARD_CONFIG.meses[mes]} ${anio}`,
      esVencidas: false,
      mes, anio,
      opps: [],
      porEtapa: {}
    });
  }

  const sinFechaBucket = buckets[0];
  const vencidasBucket = buckets[1];

  oppsFiltradas.forEach(o => {
    const etapa = String(o.Etapa || '').trim();
    const monto = parseFloat(o.MontoEstimado || 0);
    const f = dashParseFecha(o.FechaCierre);

    // Sin FechaCierre válida → bucket dedicado, sin filtrar por etapa visible
    // (queremos visibilidad total de capturas incompletas).
    if (!f || isNaN(f)) {
      const fechaApertura = dashParseFecha(o.FechaApertura);
      const oppEnriquecida = { ...o, _monto: monto, _fecha: null, _fechaApertura: fechaApertura };
      sinFechaBucket.opps.push(oppEnriquecida);
      if (!sinFechaBucket.porEtapa[etapa]) sinFechaBucket.porEtapa[etapa] = [];
      sinFechaBucket.porEtapa[etapa].push(oppEnriquecida);
      return;
    }

    const oppEnriquecida = { ...o, _monto: monto, _fecha: f };

    if (f < hoyInicio) {
      // Vencidas: incluye todas las etapas visibles. Etapas excluidas se ignoran.
      if (!_PM_ETAPAS_VISIBLES.includes(etapa)) return;
      vencidasBucket.opps.push(oppEnriquecida);
      if (!vencidasBucket.porEtapa[etapa]) vencidasBucket.porEtapa[etapa] = [];
      vencidasBucket.porEtapa[etapa].push(oppEnriquecida);
      return;
    }

    // Bucket por mes futuro: solo si cae dentro de los próximos 6 meses
    for (let i = 2; i < buckets.length; i++) {
      const b = buckets[i];
      if (f.getFullYear() === b.anio && f.getMonth() === b.mes) {
        if (!_PM_ETAPAS_VISIBLES.includes(etapa)) return;
        b.opps.push(oppEnriquecida);
        if (!b.porEtapa[etapa]) b.porEtapa[etapa] = [];
        b.porEtapa[etapa].push(oppEnriquecida);
        return;
      }
    }
    // FechaCierre > 6 meses al futuro → fuera del rango visible (descarta).
  });

  // Calcular totales por bucket y el monto pico (para escala global)
  buckets.forEach(b => {
    b.total = b.opps.reduce((s, o) => s + o._monto, 0);
    b.count = b.opps.length;
  });
  // Eliminar buckets opcionales vacíos (sin fecha y vencidas). Orden importante:
  // borrar vencidas primero (índice 1) para no shiftear el de sin fecha.
  if (vencidasBucket.opps.length === 0) buckets.splice(1, 1);
  if (sinFechaBucket.opps.length === 0) buckets.shift();

  const pico = Math.max(0, ...buckets.map(b => b.total));
  return { buckets, pico };
}

// Render de la sección "Cierre proyectado por mes" embebida al final del tab
// Pipeline. Toma asesorNorm + divisiones igual que _pipelineHtml.
function _pipelineMensualHtml(asesorNorm, divisionesVisibles) {
  const { buckets, pico } = _pipelineMensualBuckets(asesorNorm, divisionesVisibles);
  const etapasConfig = DASHBOARD_CONFIG.etapas;

  // Caso sin nada: ni vencidas ni opps en próximos 6m
  const totalVisible = buckets.reduce((s, b) => s + b.total, 0);
  const countTotal = buckets.reduce((s, b) => s + b.count, 0);
  if (countTotal === 0) {
    return `
      <div style="padding:14px 0 0;margin-top:12px;border-top:2px solid var(--color-border-tertiary)">
        <div style="padding:0 16px 4px">
          <div style="font-size:12px;text-transform:uppercase;letter-spacing:0.04em;color:var(--color-text-tertiary);margin-bottom:6px">
            Cierre proyectado por mes
          </div>
        </div>
        <div style="padding:24px 16px;text-align:center;color:var(--color-text-secondary);font-size:13px">
          Sin oportunidades en pipeline para los próximos 6 meses.
        </div>
      </div>`;
  }

  const legend = _PM_ETAPAS_VISIBLES.map(et => `
    <div class="dash-pm-leg-item">
      <span class="dash-pm-leg-dot" style="background:${etapasConfig[et]?.color || '#888'}"></span>
      <span>${et}</span>
    </div>`).join('');

  const filasHtml = buckets.map((b, idx) => {
    const drillId = `dash-pm-drill-${b.id}`;
    const widthPct = pico > 0 ? Math.max(2, (b.total / pico) * 100) : 0;

    // Segmentos: para vencidas, una sola barra roja sólida (no se desglosa por etapa visualmente
    // — el drill muestra la lista completa ordenada por fecha). Para meses, segmentos por etapa.
    let segmentos = '';
    if (b.total === 0 && !b.esSinFecha) {
      // Mes vacío — barra gris muy clarita para comunicar "sí, sin nada".
      // (Sin fecha proyectada cae al else de abajo: aunque _monto sea 0 queremos
      // mostrar la barra gris media y permitir drill-down con la lista.)
      segmentos = `<div class="dash-pm-seg dash-pm-seg-empty" style="width:100%"></div>`;
    } else if (b.esSinFecha) {
      segmentos = `<div class="dash-pm-seg" style="width:100%;background:${_PM_COLOR_SIN_FECHA};cursor:pointer"
        onclick="event.stopPropagation();dashPmDrill('${drillId}','${b.id}','sin-fecha')"
        title="Sin fecha proyectada · ${dashFmt(b.total)} · ${b.count} opps"></div>`;
    } else if (b.esVencidas) {
      segmentos = `<div class="dash-pm-seg" style="width:100%;background:${_PM_COLOR_VENCIDA};cursor:pointer"
        onclick="event.stopPropagation();dashPmDrill('${drillId}','${b.id}','vencidas')"
        title="Vencidas · ${dashFmt(b.total)} · ${b.count} opps"></div>`;
    } else {
      // Segmentos por etapa: cada uno con width proporcional al monto de la etapa
      // EN ESTE MES, dividido por el monto del mes con más pipeline (escala global).
      // Total width visible del row = (b.total / pico) * 100%.
      segmentos = _PM_ETAPAS_VISIBLES.map(et => {
        const ops = b.porEtapa[et] || [];
        const totalEt = ops.reduce((s, o) => s + o._monto, 0);
        if (totalEt <= 0) return '';
        const widthEt = pico > 0 ? (totalEt / pico) * 100 : 0;
        const c = etapasConfig[et]?.color || '#888';
        return `<div class="dash-pm-seg" style="width:${widthEt}%;background:${c};cursor:pointer"
          onclick="event.stopPropagation();dashPmDrill('${drillId}','${b.id}','${et.replace(/'/g,"\\'")}')"
          title="${et} · ${dashFmt(totalEt)} · ${ops.length} opps"></div>`;
      }).join('');
      // Espacio restante (gris claro) para completar la barra al 100% del width del contenedor
      const restante = 100 - widthPct;
      if (restante > 0) {
        segmentos += `<div class="dash-pm-seg dash-pm-seg-empty" style="width:${restante}%"></div>`;
      }
    }

    // Tap en row-head abre drill apropiado al tipo de bucket. Para Sin Fecha y
    // Vencidas el drill usa su ordenamiento dedicado (FechaApertura desc / FechaCierre asc);
    // para meses usa 'todas' (etapas mezcladas ordenadas por monto desc).
    const headEtapa = b.esSinFecha ? 'sin-fecha' : (b.esVencidas ? 'vencidas' : 'todas');
    const totalLabel = (b.esSinFecha && b.total === 0) ? '—' : dashFmt(b.total);

    return `
      <div class="dash-pm-row">
        <div class="dash-pm-row-head" onclick="dashPmDrill('${drillId}','${b.id}','${headEtapa}')" style="cursor:pointer">
          <div class="dash-pm-label">${b.label} · ${b.count} ${b.count === 1 ? 'opp' : 'opps'}</div>
          <div class="dash-pm-total">${totalLabel}</div>
        </div>
        <div class="dash-pm-bar">${segmentos}</div>
        <div id="${drillId}" class="dash-pm-drill" style="display:none"></div>
      </div>`;
  }).join('');

  // Outer wrapper: top border + margin como separador de sección (mismo patrón
  // que la sección OVs dentro de _pipelineHtml). Sin padding-bottom de 80px
  // porque ya no es el bottom de un tab page — vive embebida dentro del tab Pipeline.
  return `
    <div style="padding:14px 0 0;margin-top:12px;border-top:2px solid var(--color-border-tertiary)">
      <div style="padding:0 16px 4px">
        <div style="font-size:12px;text-transform:uppercase;letter-spacing:0.04em;color:var(--color-text-tertiary);margin-bottom:6px">
          Cierre proyectado por mes
        </div>
        <div style="font-size:20px;font-weight:600">${dashFmt(totalVisible)}</div>
        <div style="font-size:12px;color:var(--color-text-secondary)">pipeline visible · ${countTotal} oportunidades</div>
      </div>
      <div class="dash-pm-legend">${legend}</div>
      <div style="padding:0 16px">
        ${filasHtml}
      </div>
    </div>`;
}

// Drill-down inline: expande lista debajo de la barra. Solo un panel abierto a
// la vez — al abrir otro, cierra el anterior.
let _PM_DRILL_OPEN = null;
function dashPmDrill(drillId, bucketId, etapa) {
  const el = document.getElementById(drillId);
  if (!el) return;

  // Si está abierto y se hace tap al mismo segmento+mes → cerrar (toggle).
  const tag = `${drillId}::${etapa}`;
  if (_PM_DRILL_OPEN === tag) {
    el.style.display = 'none';
    _PM_DRILL_OPEN = null;
    return;
  }
  // Cerrar cualquier otro panel previamente abierto
  if (_PM_DRILL_OPEN) {
    const prev = _PM_DRILL_OPEN.split('::')[0];
    const prevEl = document.getElementById(prev);
    if (prevEl) prevEl.style.display = 'none';
  }

  // Reconstruir buckets para extraer los opps de este bucket+etapa
  // (No cachear: re-computar es barato y evita inconsistencia post-refresh)
  const asesorNorm = DASH_STATE.rol === 'asesor'
    ? dashNormPresup(_findAsesorNombre(DASH_STATE.userEmail))
    : null;
  const divs = dashGetDivisionesVisibles();
  const { buckets } = _pipelineMensualBuckets(asesorNorm, divs);
  const bucket = buckets.find(b => b.id === bucketId);
  if (!bucket) return;

  let opps = [];
  let titulo = '';
  const hoy = new Date();
  const hoyInicio = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());

  if (etapa === 'vencidas') {
    // Todas las vencidas, más antiguas primero (más vencidas arriba)
    opps = [...bucket.opps].sort((a, b) => a._fecha - b._fecha);
    titulo = `Vencidas · ${opps.length} opps`;
  } else if (etapa === 'sin-fecha') {
    // Sin fecha proyectada: ordenadas por FechaApertura desc (más recientes primero).
    // Si _fechaApertura es null cae al final.
    opps = [...bucket.opps].sort((a, b) => {
      const fa = a._fechaApertura ? a._fechaApertura.getTime() : -Infinity;
      const fb = b._fechaApertura ? b._fechaApertura.getTime() : -Infinity;
      return fb - fa;
    });
    titulo = `Sin fecha proyectada · ${opps.length} opps`;
  } else if (etapa === 'todas') {
    opps = [...bucket.opps].sort((a, b) => b._monto - a._monto);
    titulo = `${bucket.label.replace('⚠ ','')} · todas las etapas · ${opps.length} opps`;
  } else {
    opps = (bucket.porEtapa[etapa] || []).slice().sort((a, b) => b._monto - a._monto);
    titulo = `${bucket.label.replace('⚠ ','')} · ${etapa} · ${opps.length} opps`;
  }

  const itemsHtml = opps.map(o => {
    const dias = o._fecha ? Math.round((o._fecha - hoyInicio) / 86400000) : null;
    const diasLabel = dias === null ? ''
      : dias < 0 ? `<span style="color:${_PM_COLOR_VENCIDA};font-weight:500">vencida hace ${Math.abs(dias)}d</span>`
      : dias === 0 ? `<span style="color:#BA7517;font-weight:500">vence hoy</span>`
      : `faltan ${dias}d`;
    const etapaLabel = String(o.Etapa || '').trim();
    const cfg = DASHBOARD_CONFIG.etapas[etapaLabel];
    const dot = cfg ? `<span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:${cfg.color};margin-right:5px;vertical-align:middle"></span>` : '';
    return `
      <div class="dash-pm-drill-item">
        <div style="flex:1;min-width:0">
          <div class="dash-pm-drill-cliente">${o.Cliente || o.CardName || ''}</div>
          <div class="dash-pm-drill-sub">
            ${o.Descripcion ? o.Descripcion + ' · ' : ''}${dot}${etapaLabel || '—'}
          </div>
          <div class="dash-pm-drill-sub">
            Cierre ${o.FechaCierre || '—'}${diasLabel ? ' · ' + diasLabel : ''}
            ${!asesorNorm && o.Asesor ? ' · ' + o.Asesor : ''}
          </div>
        </div>
        <div class="dash-pm-drill-monto">${dashFmt(o._monto)}</div>
      </div>`;
  }).join('');

  el.innerHTML = `
    <div class="dash-pm-drill-head">${titulo}</div>
    ${opps.length === 0
      ? `<div style="padding:12px;text-align:center;color:var(--color-text-secondary);font-size:12px">Sin oportunidades en este corte.</div>`
      : itemsHtml}`;
  el.style.display = 'block';
  _PM_DRILL_OPEN = tag;
}

// ── ASESOR VIEW ──────────────────────────────────────────────────────────────
function _renderAsesor(container, asesor, mes, anio, isCurrent, mesLabel, divs) {
  const met    = dashCalcMetricas(asesor, mes, anio, divs);
  const hist   = dashHistoricoAnual(asesor, anio, divs);
  const riesgo = dashClientesEnRiesgo(asesor, 60);
  const ctx    = dashMensajeCtx(met.pct, met.falta, met.diasRestantes, met.ritmo, isCurrent);
  const col    = dashColor(met.pct);

  container.innerHTML = `
    <div class="dash-tabs">
      <button class="dash-tab active" onclick="dashTabSwitch('a-resumen',this)">Resumen</button>
      <button class="dash-tab" onclick="dashTabSwitch('a-pipeline',this)">Pipeline</button>
      <button class="dash-tab" onclick="dashTabSwitch('a-riesgo',this)">En riesgo</button>
    </div>

    <div id="dash-page-a-resumen" class="dash-page active" style="overflow-y:auto">
      ${_periodSelector(mesLabel, isCurrent)}
      <div class="dash-streak">
        <div style="font-size:26px">🔥</div>
        <div style="flex:1">
          <div class="dash-snum">Racha activa</div>
          <div class="dash-slbl">semanas con actividad registrada</div>
          <div class="dash-sdots">
            <div class="dash-sd done"></div><div class="dash-sd done"></div>
            <div class="dash-sd done"></div><div class="dash-sd cur"></div>
            <div class="dash-sd empty"></div><div class="dash-sd empty"></div>
          </div>
        </div>
      </div>
      <div class="dash-ctx dash-ctx-${ctx.tipo}"><div class="dash-cdot"></div><p>${ctx.msg}</p></div>
      <div class="dash-sec">
        <div class="dash-lbl">Progreso vs meta — ${mesLabel}</div>
        ${_goalCard(met.totalVenta, met.totalMeta, met.pct, met.falta, met.diasRestantes, isCurrent, 'pf-asesor')}
      </div>
      <div class="dash-kpis">
        <div class="dash-kpi"><div class="dash-kpi-v">${met.numOVs}</div><div class="dash-kpi-l">OVs generadas</div></div>
        <div class="dash-kpi"><div class="dash-kpi-v">${riesgo.length}</div><div class="dash-kpi-l">Clientes en riesgo</div><div class="dash-kpi-d" style="color:#A32D2D">+60 días sin compra</div></div>
        <div class="dash-kpi"><div class="dash-kpi-v">${Object.keys(met.porDivision).length}</div><div class="dash-kpi-l">Divisiones activas</div></div>
        <div class="dash-kpi"><div class="dash-kpi-v">${met.pct}%</div><div class="dash-kpi-l">Cumplimiento</div></div>
      </div>
      ${_divisionesHtml(met.porDivision, met.metaPorDivision)}
      ${_anualHtml(hist, anio, mes)}
      <div style="height:20px"></div>
    </div>

    <div id="dash-page-a-pipeline" class="dash-page" style="overflow-y:auto">
      ${_pipelineHtml(dashNormPresup(asesor), divs)}
    </div>

    <div id="dash-page-a-riesgo" class="dash-page" style="overflow-y:auto">
      ${_riesgoHtml(riesgo, false)}
    </div>`;

  setTimeout(() => {
    const b = document.getElementById('pf-asesor');
    if (b) b.style.width = Math.min(met.pct,100)+'%';
    _initChart('dash-chart-hist', hist);
  }, 100);
}

// ── LIDER VIEW ───────────────────────────────────────────────────────────────
function _renderLider(container, mes, anio, isCurrent, mesLabel, divs) {
  const asesores = dashGetAsesores();
  const equipo   = asesores.map(a => {
    const met = dashCalcMetricas(a, mes, anio, divs);
    const hasAnualVenta = DASH_STATE.ventas.some(v =>
      dashNormNombre(v.Asesor) === dashNormPresup(a) &&
      dashParseFecha(v.Fecha)?.getFullYear() === anio &&
      dashGetTotal(v) > 0
    );
    return { nombre: a, ...met, hasAnualVenta };
  }).filter(a => a.totalMeta > 0 || a.totalVenta > 0 || a.hasAnualVenta);
  const totalEq  = equipo.reduce((s,a)=>s+a.totalVenta,0);
  const metaEq   = equipo.reduce((s,a)=>s+a.totalMeta,0);
  const pctEq    = metaEq>0 ? Math.round(totalEq/metaEq*100) : 0;
  const col      = dashColor(pctEq);
  const hist     = dashHistoricoAnualEmpresa(asesores, anio, divs);
  const riesgo   = dashClientesEnRiesgo(null, 60);

  container.innerHTML = `
    <div class="dash-tabs">
      <button class="dash-tab active" onclick="dashTabSwitch('l-resumen',this)">Resumen</button>
      <button class="dash-tab" onclick="dashTabSwitch('l-pipeline',this)">Pipeline</button>
      <button class="dash-tab" onclick="dashTabSwitch('l-asesores',this)">Asesores</button>
      <button class="dash-tab" onclick="dashTabSwitch('l-riesgo',this)">En riesgo</button>
    </div>

    <div id="dash-page-l-resumen" class="dash-page active" style="overflow-y:auto">
      ${_periodSelector(mesLabel, isCurrent)}
      <div class="dash-ctx dash-ctx-${pctEq>=65?'info':'warn'}">
        <div class="dash-cdot"></div>
        <p>El equipo va en <strong>${pctEq}%</strong> de meta para ${mesLabel}. ${equipo.filter(a=>a.pct<50).length} asesores por debajo del 50%.</p>
      </div>
      <div class="dash-sec">
        <div class="dash-lbl">Equipo · ${mesLabel}</div>
        ${_goalCard(totalEq, metaEq, pctEq, Math.max(0,metaEq-totalEq), diasEnMes(mes,anio)-new Date().getDate(), isCurrent, 'pf-lider')}
      </div>
      ${_anualHtml(hist, anio, mes)}
      <div style="height:20px"></div>
    </div>

    <div id="dash-page-l-pipeline" class="dash-page" style="overflow-y:auto">
      ${_pipelineHtml(null, divs)}
    </div>

    <div id="dash-page-l-asesores" class="dash-page" style="overflow-y:auto">
      ${_asesoresListaHtml(equipo, false)}
    </div>

    <div id="dash-page-l-riesgo" class="dash-page" style="overflow-y:auto">
      ${_riesgoHtml(riesgo, true)}
    </div>`;

  setTimeout(() => {
    const b = document.getElementById('pf-lider');
    if (b) b.style.width = Math.min(pctEq,100)+'%';
    equipo.forEach(a => {
      const el = document.getElementById('dash-fill-'+dashSlug(a.nombre));
      if (el) setTimeout(()=>el.style.width=Math.min(a.pct,100)+'%',150);
    });
    _initChart('dash-chart-hist', hist);
  }, 100);
}

// ── GERENTE / DIRECTOR VIEW ──────────────────────────────────────────────────
function _renderGerente(container, mes, anio, isCurrent, mesLabel, divs) {
  const asesores = dashGetAsesores();
  const equipo   = asesores.map(a => {
    const met = dashCalcMetricas(a, mes, anio, divs);
    // Also check if asesor has any venta this year to avoid showing completely inactive asesores
    const hasAnualVenta = DASH_STATE.ventas.some(v =>
      dashNormNombre(v.Asesor) === dashNormPresup(a) &&
      dashParseFecha(v.Fecha)?.getFullYear() === anio &&
      dashGetTotal(v) > 0
    );
    return { nombre: a, ...met, hasAnualVenta };
  }).filter(a => a.totalMeta > 0 || a.totalVenta > 0 || a.hasAnualVenta);
  const totalEmp = equipo.reduce((s,a)=>s+a.totalVenta,0);
  const metaEmp  = equipo.reduce((s,a)=>s+a.totalMeta,0);
  const pctEmp   = metaEmp>0 ? Math.round(totalEmp/metaEmp*100) : 0;
  const col      = dashColor(pctEmp);
  const hist     = dashHistoricoAnualEmpresa(asesores, anio, divs);
  const riesgo   = dashClientesEnRiesgo(null, 60);

  container.innerHTML = `
    <div class="dash-tabs">
      <button class="dash-tab active" onclick="dashTabSwitch('g-resumen',this)">Resumen</button>
      <button class="dash-tab" onclick="dashTabSwitch('g-pipeline',this)">Pipeline</button>
      <button class="dash-tab" onclick="dashTabSwitch('g-asesores',this)">Asesores</button>
      <button class="dash-tab" onclick="dashTabSwitch('g-riesgo',this)">En riesgo</button>
    </div>

    <div id="dash-page-g-resumen" class="dash-page active" style="overflow-y:auto">
      ${_periodSelector(mesLabel, isCurrent)}
      <div class="dash-ctx dash-ctx-${pctEmp>=65?'info':'warn'}">
        <div class="dash-cdot"></div>
        <p>Empresa va en <strong>${pctEmp}%</strong> de meta para ${mesLabel}. ${equipo.filter(a=>a.pct<50).length} asesores requieren atención.</p>
      </div>
      <div class="dash-sec">
        <div class="dash-lbl">Vinssa · ${mesLabel}</div>
        ${_goalCard(totalEmp, metaEmp, pctEmp, Math.max(0,metaEmp-totalEmp), diasEnMes(mes,anio)-new Date().getDate(), isCurrent, 'pf-gerente')}
      </div>
      <div class="dash-grid2">
        <div class="dash-mc"><div class="dash-mc-v">${equipo.filter(a=>a.pct>=65).length}/${equipo.length}</div><div class="dash-mc-l">Asesores en meta</div></div>
        <div class="dash-mc"><div class="dash-mc-v">${riesgo.length}</div><div class="dash-mc-l">Clientes en riesgo</div><div class="dash-mc-d" style="color:#A32D2D">+60 días sin compra</div></div>
      </div>
      ${_anualHtml(hist, anio, mes)}
      <div style="height:20px"></div>
    </div>

    <div id="dash-page-g-pipeline" class="dash-page" style="overflow-y:auto">
      ${_pipelineHtml(null, divs)}
    </div>

    <div id="dash-page-g-asesores" class="dash-page" style="overflow-y:auto">
      ${_asesoresListaHtml(equipo, true)}
    </div>

    <div id="dash-page-g-riesgo" class="dash-page" style="overflow-y:auto">
      ${_riesgoHtml(riesgo, true)}
    </div>`;

  setTimeout(() => {
    const b = document.getElementById('pf-gerente');
    if (b) b.style.width = Math.min(pctEmp,100)+'%';
    equipo.forEach(a => {
      const el = document.getElementById('dash-fill-'+dashSlug(a.nombre));
      if (el) setTimeout(()=>el.style.width=Math.min(a.pct,100)+'%',150);
    });
    _initChart('dash-chart-hist', hist);
  }, 100);
}

// ── Divisiones bar chart ──────────────────────────────────────────────────────
function _divisionesHtml(porDivision, metaPorDivision) {
  const entries = Object.entries(porDivision).sort((a,b)=>b[1]-a[1]);
  if (!entries.length) return '';
  return `
    <div class="dash-sec" style="padding-top:12px">
      <div class="dash-lbl">Por división</div>
      ${entries.map(([div,venta])=>{
        const meta = metaPorDivision[div]||0;
        const pct  = meta>0?Math.round(venta/meta*100):0;
        const col  = dashColor(pct);
        return `<div style="margin-bottom:10px">
          <div style="display:flex;justify-content:space-between;margin-bottom:4px">
            <span style="font-size:12px;font-weight:500">${div}</span>
            <span style="font-size:12px;color:${col.pct};font-weight:500">${dashFmt(venta)}${meta>0?' / '+dashFmt(meta):''}</span>
          </div>
          <div class="dash-pbar"><div class="dash-pfill" style="width:${Math.min(pct,100)}%;background:${col.bar}"></div></div>
        </div>`;
      }).join('')}
    </div>`;
}

// ── Chart helper ──────────────────────────────────────────────────────────────
function _initChart(canvasId, hist) {
  const canvas = document.getElementById(canvasId);
  if (!canvas || !window.Chart) return;
  // Destroy previous chart if exists
  if (canvas._chartInstance) { canvas._chartInstance.destroy(); }
  const instance = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: hist.meses.map(m=>m.label),
      datasets: [
        { label:'Venta', data: hist.meses.map(m=>Math.round(m.venta)), backgroundColor:'#185FA5', borderRadius:3 },
        { label:'Meta',  data: hist.meses.map(m=>Math.round(m.meta)),  backgroundColor:'#D3D1C7', borderRadius:3 }
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false} },
      scales:{
        x:{grid:{display:false}, ticks:{font:{size:9},color:'#888780'}},
        y:{display:false}
      }
    }
  });
  canvas._chartInstance = instance;
}

// ── Histórico anual empresa ───────────────────────────────────────────────────
function dashHistoricoAnualEmpresa(asesores, anio, divisionesVisibles) {
  const meses = [];
  for (let m = 1; m <= 12; m++) {
    let venta=0, meta=0, ventaAnt=0;
    asesores.forEach(a => {
      const met = dashCalcMetricas(a, m, anio, divisionesVisibles);
      const ant = dashCalcMetricas(a, m, anio-1, divisionesVisibles);
      venta += met.totalVenta;
      meta  += met.totalMeta;
      ventaAnt += ant.totalVenta;
    });
    const pct = meta>0?Math.round(venta/meta*100):0;
    meses.push({mes:m, label:DASHBOARD_CONFIG.meses[m-1], venta, meta, ventaAnterior:ventaAnt, pct});
  }
  // Only count months up to current month for acumulado comparison
  const mesActual = new Date().getMonth()+1;
  const acumulado    = meses.slice(0,mesActual).reduce((s,m)=>s+m.venta,0);
  const acumAnt      = meses.slice(0,mesActual).reduce((s,m)=>s+m.ventaAnterior,0);
  const mejorMes     = meses.reduce((b,m)=>m.venta>b.venta?m:b, meses[0]);
  const metaAnual    = meses.reduce((s,m)=>s+m.meta,0);
  const mesesConDatos = meses.filter(m=>m.venta>0);
  const proyeccion   = mesesConDatos.length>0?(acumulado/mesesConDatos.length)*12:0;
  return { meses, acumulado, acumuladoAnterior:acumAnt, mejorMes, metaAnual, proyeccion };
}

// ── Navigation ────────────────────────────────────────────────────────────────
function dashTabSwitch(id, btn) {
  const container = document.getElementById('dashboard-container');
  container.querySelectorAll('.dash-tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  container.querySelectorAll('.dash-page').forEach(p=>p.classList.remove('active'));
  const page = document.getElementById('dash-page-'+id);
  if (page) page.classList.add('active');
}

function dashTogglePeriodPanel() {
  const p = document.getElementById('dash-period-panel');
  if (p) p.classList.toggle('open');
}

function dashSelMes(mes) {
  DASH_STATE.mesSel = mes;
  const p = document.getElementById('dash-period-panel');
  if (p) p.classList.remove('open');
  dashRender();
}

function dashFiltrarAsesores(filtro, btn) {
  document.querySelectorAll('.dash-fchip').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const lista = document.getElementById('dash-asesores-lista');
  if (!lista) return;
  lista.querySelectorAll('.dash-ac').forEach(card => {
    const pctEl = card.querySelector('[style*="font-size:15px"]');
    const pct   = parseInt(pctEl?.textContent || '100');
    if      (filtro==='todos')  card.style.display='';
    else if (filtro==='riesgo') card.style.display=pct<50?'':'none';
    else if (filtro==='meta')   card.style.display=pct>=65?'':'none';
  });
}

// Reordena los cards en el DOM SIN re-renderizar — preserva el filtro activo.
function dashOrdenarAsesores(key) {
  DASH_STATE.asesoresSort = key;
  const lista = document.getElementById('dash-asesores-lista');
  if (!lista) return;
  const cards = Array.from(lista.querySelectorAll('.dash-ac'));
  cards.sort((a, b) => {
    const av = parseFloat(a.dataset.venta) || 0;
    const bv = parseFloat(b.dataset.venta) || 0;
    const ap = parseFloat(a.dataset.pct)   || 0;
    const bp = parseFloat(b.dataset.pct)   || 0;
    const aH = a.dataset.haspresup === '1';
    const bH = b.dataset.haspresup === '1';
    if (key === 'menos_ventas') return av - bv;
    if (key === 'mejor_pct' || key === 'peor_pct') {
      if (aH !== bH) return aH ? -1 : 1;
      if (!aH) return 0;
      return key === 'mejor_pct' ? bp - ap : ap - bp;
    }
    return bv - av; // mas_ventas
  });
  cards.forEach(c => lista.appendChild(c));
}

function dashDrillAsesor(nombre) {
  alert(`Vista detalle de ${nombre} — próximamente`);
}

function dashMostrarLoader(show) {
  const l = document.getElementById('dash-loader');
  if (l) l.style.display=show?'flex':'none';
}

function dashMostrarError(msg) {
  const c = document.getElementById('dashboard-container');
  if (c) c.innerHTML=`
    <div style="padding:24px;text-align:center">
      <div style="font-size:32px;margin-bottom:12px">⚠️</div>
      <div style="font-size:14px;font-weight:500;margin-bottom:8px">Error cargando dashboard</div>
      <div style="font-size:12px;color:var(--color-text-secondary)">${msg}</div>
      <button onclick="dashInit()" style="margin-top:16px;padding:10px 20px;background:#185FA5;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px">Reintentar</button>
    </div>`;
}

function dashSlug(str) {
  return str.toLowerCase().replace(/[^a-z0-9]/g,'-');
}

// ── Inject CSS ────────────────────────────────────────────────────────────────
(function(){
  const s = document.createElement('style');
  s.textContent = `
    .dash-page { display:none; }
    .dash-page.active { display:block; }
    .dchev.rotated { transform:rotate(180deg); }
    .dash-page[style*="overflow-y:auto"] { -webkit-overflow-scrolling:touch; }
  `;
  document.head.appendChild(s);
})();

// ── Exponer globalmente ───────────────────────────────────────────────────────
window.dashTabSwitch         = dashTabSwitch;
window.dashTogglePeriodPanel = dashTogglePeriodPanel;
window.dashSelMes            = dashSelMes;
window.dashFiltrarAsesores   = dashFiltrarAsesores;
window.dashOrdenarAsesores   = dashOrdenarAsesores;
window.dashDrillAsesor       = dashDrillAsesor;
window.dashCambiarMes        = dashSelMes;
window.dashInit              = dashInit;
window.dashNormPresup        = dashNormPresup;
window.dashGetTotal          = dashGetTotal;
window.dashParseFecha        = dashParseFecha;
window.dashPmDrill           = dashPmDrill;

// ── TEST MODE BAR (eliminar antes de producción) ──────────────────────────────
window.dashTestMode = function(rol, division) {
  DASH_STATE.rol      = rol;
  DASH_STATE.division = division;
  console.log('Test mode:', rol, division);
  dashRender();
};
(function(){
  const observer = new MutationObserver(() => {
    if (document.getElementById('dash-test-bar')) return;
    const dash = document.getElementById('screen-dashboard');
    if (!dash) return;
    const bar = document.createElement('div');
    bar.id = 'dash-test-bar';
    bar.style.cssText = 'position:fixed;bottom:70px;left:0;right:0;background:#111827;color:white;font-size:10px;padding:5px 8px;z-index:9999;display:flex;gap:4px;flex-wrap:wrap;justify-content:center';
    bar.innerHTML = `
      <span style="opacity:0.5;align-self:center">TEST:</span>
      <button onclick="dashTestMode('gerente','Todos')" style="font-size:10px;padding:2px 6px;border:1px solid #444;background:#222;color:white;border-radius:4px;cursor:pointer">Dir·Todos</button>
      <button onclick="dashTestMode('gerente','Trazabilidad')" style="font-size:10px;padding:2px 6px;border:1px solid #444;background:#222;color:white;border-radius:4px;cursor:pointer">Ger·Traz</button>
      <button onclick="dashTestMode('gerente','Suministros')" style="font-size:10px;padding:2px 6px;border:1px solid #444;background:#222;color:white;border-radius:4px;cursor:pointer">Ger·Sum</button>
      <button onclick="dashTestMode('lider','Trazabilidad')" style="font-size:10px;padding:2px 6px;border:1px solid #444;background:#222;color:white;border-radius:4px;cursor:pointer">Líd·Traz</button>
      <button onclick="dashTestMode('asesor','Trazabilidad')" style="font-size:10px;padding:2px 6px;border:1px solid #444;background:#222;color:white;border-radius:4px;cursor:pointer">Ase·Traz</button>
      <button onclick="dashTestMode('asesor','Suministros')" style="font-size:10px;padding:2px 6px;border:1px solid #444;background:#222;color:white;border-radius:4px;cursor:pointer">Ase·Sum</button>`;
    dash.appendChild(bar);
  });
  observer.observe(document.body, { childList:true, subtree:true });
})();
