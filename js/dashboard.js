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
  driveId: null
};

// ── Obtener token Graph ──────────────────────────────────────────────────────
async function dashGetToken() {
  if (DASH_STATE.token) return DASH_STATE.token;
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error('No hay sesión activa');
  const result = await msalInstance.acquireTokenSilent({
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
  // Returns asesores to show in lider/gerente views.
  // Source of truth: presupuesto (has names + divisions).
  // For divisions not in presupuesto (e.g. Suministros), fall back to ventas.
  const divs = dashGetDivisionesVisibles(); // null = Todos

  const seen = new Set();
  const result = [];

  if (!divs) {
    // Dir·Todos: all asesores from presupuesto
    DASH_STATE.presupuesto.forEach(p => {
      const nombre = String(p.Asesor || '').trim();
      if (nombre && !seen.has(nombre)) { seen.add(nombre); result.push(nombre); }
    });
  } else {
    // Check if any presupuesto rows exist for these divisions
    const presupDivs = DASH_STATE.presupuesto.filter(p => {
      const pDiv = String(p.Division || '').trim();
      return divs.includes(pDiv);
    });

    if (presupDivs.length > 0) {
      // Use presupuesto — filter by division
      presupDivs.forEach(p => {
        const nombre = String(p.Asesor || '').trim();
        if (nombre && !seen.has(nombre)) { seen.add(nombre); result.push(nombre); }
      });
    } else {
      // No presupuesto for this division — derive asesores from ventas
      DASH_STATE.ventas.forEach(v => {
        const grupo = String(v.GrupoArticulo || '').trim();
        const div   = DASHBOARD_CONFIG.mapaGrupos[grupo];
        if (!divs.includes(div)) return;
        const nombre = String(v.Asesor || '').trim();
        if (nombre && !seen.has(nombre)) { seen.add(nombre); result.push(nombre); }
      });
    }
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
  if (num >= 1000000) return '$' + (num/1000000).toFixed(1) + 'M';
  if (num >= 1000)    return '$' + Math.round(num/1000) + 'K';
  return '$' + Math.round(num).toLocaleString('es-MX');
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

// ── SHARED: Asesores filter list ─────────────────────────────────────────────
function _asesoresListaHtml(equipo, clickable) {
  return `
    <div class="dash-filter-row">
      <button class="dash-fchip active" onclick="dashFiltrarAsesores('todos',this)">Todos</button>
      <button class="dash-fchip" onclick="dashFiltrarAsesores('riesgo',this)">En riesgo</button>
      <button class="dash-fchip" onclick="dashFiltrarAsesores('meta',this)">En meta</button>
    </div>
    <div id="dash-asesores-lista">
      ${equipo.map(a => _asesorCard(a, clickable)).join('')}
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
      const linea = String(o.Linea || '').trim();
      // Try direct division match first
      const divDirecta = DASHBOARD_CONFIG.mapaGrupos[linea];
      if (divDirecta) return divisionesVisibles.includes(divDirecta);
      // Try case-insensitive substring match against division names
      const lineaLow = linea.toLowerCase();
      const matched = divisionesVisibles.some(d => lineaLow.includes(d.toLowerCase()));
      if (matched) return true;
      // Also check marca field
      const marca = String(o.Marca || '').toLowerCase();
      const marcaKeywords = {
        'Trazabilidad': ['telesis','pinstamp','markem','imaje','datalogic'],
        'Visión': ['cognex','keyence','sick','banner'],
        'Robótica': ['nabtesco','ur','mir','easyrobotics','abb','fanuc'],
        'Suministros': ['loctite','bonderite','henkel','ansell','3m']
      };
      return divisionesVisibles.some(d => {
        const kws = marcaKeywords[d] || [];
        return kws.some(k => marca.includes(k));
      });
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
  const equipo   = asesores.map(a => ({ nombre: a, ...dashCalcMetricas(a, mes, anio, divs) }))
                           .filter(a => a.totalMeta > 0 || a.totalVenta > 0);
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
  const equipo   = asesores.map(a => ({ nombre: a, ...dashCalcMetricas(a, mes, anio, divs) }))
                           .filter(a => a.totalMeta > 0 || a.totalVenta > 0);
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
window.dashDrillAsesor       = dashDrillAsesor;
window.dashCambiarMes        = dashSelMes;
window.dashInit              = dashInit;
window.dashNormPresup        = dashNormPresup;
window.dashGetTotal          = dashGetTotal;
window.dashParseFecha        = dashParseFecha;

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
