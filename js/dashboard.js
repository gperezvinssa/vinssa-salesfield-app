// ── Vinssa Dashboard — js/dashboard.js ──────────────────────────────────────
// Lee roles de Lista Roles Dashboard.xlsx
// Lee ventas de Ventas Asesor.xlsx
// Lee presupuesto de Presupuesto Ventas.xlsx
// Todo vía Microsoft Graph API (mismo token que usa sap.js)

// ── Constantes SharePoint ────────────────────────────────────────────────────
const DASHBOARD_CONFIG = {
  siteId: 'versatilidadsaltillo.sharepoint.com,/sites/VINSSAAutomation',
  archivos: {
    roles:       'Lista Roles Dashboard.xlsx',
    ventas:      'Ventas Asesor.xlsx',
    ovs:         'OVs Asesor.xlsx',
    presupuesto: 'Presupuesto Ventas.xlsx'
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
  userEmail: null,
  ventas: [],         // facturas (OINV)
  ovs: [],            // órdenes de venta abiertas (ORDR)
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
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent('versatilidadsaltillo.sharepoint.com:/sites/VINSSAAutomation:')}/drives`;
  const drivesRes = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const drives = await drivesRes.json();
  const driveId = drives.value[0].id;
  DASH_STATE.driveId = driveId;

  const searchUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;
  const itemsRes = await fetch(searchUrl, { headers: { Authorization: `Bearer ${token}` } });
  const items = await itemsRes.json();
  const file = items.value.find(f => f.name === nombre);
  if (!file) throw new Error(`No se encontró el archivo: ${nombre}`);
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
  try {
    const token = await dashGetToken();

    // Leer los cuatro archivos en paralelo
    const [roles, ventas, ovs, presupuesto] = await Promise.all([
      dashLeerExcel(token, DASHBOARD_CONFIG.archivos.roles),
      dashLeerExcel(token, DASHBOARD_CONFIG.archivos.ventas),
      dashLeerExcel(token, DASHBOARD_CONFIG.archivos.ovs),
      dashLeerExcel(token, DASHBOARD_CONFIG.archivos.presupuesto)
    ]);

    // Determinar rol del usuario actual
    const emailActual = DASH_STATE.userEmail;
    const userRole = roles.find(r =>
      String(r.Email || '').toLowerCase().trim() === emailActual
    );
    DASH_STATE.rol = userRole ? String(userRole.Rol || '').toLowerCase().trim() : 'asesor';

    // Normalizar rol
    if (DASH_STATE.rol.includes('gerente')) DASH_STATE.rol = 'gerente';
    else if (DASH_STATE.rol.includes('lider') || DASH_STATE.rol.includes('líder')) DASH_STATE.rol = 'lider';
    else DASH_STATE.rol = 'asesor';

    DASH_STATE.ventas = ventas;
    DASH_STATE.ovs = ovs;
    DASH_STATE.presupuesto = presupuesto;

    dashRender();
  } catch(e) {
    console.error('Error inicializando dashboard:', e);
    dashMostrarError(e.message);
  } finally {
    dashMostrarLoader(false);
  }
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
    return new Date((val - 25569) * 86400 * 1000);
  }
  const s = String(val).trim();
  // dd/MM/yyyy
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
function dashCalcMetricas(asesor, mes, anio) {
  // asesor viene del presupuesto — aplicar alias para buscar en SAP
  const asesorNorm = dashNormPresup(asesor);
  const ventasFiltradas = DASH_STATE.ventas.filter(v => {
    if (!v.Fecha) return false;
    const fecha = dashParseFecha(v.Fecha);
    if (!fecha || isNaN(fecha)) return false;
    const vMes = fecha.getMonth() + 1;
    const vAnio = fecha.getFullYear();
    const vAsesor = dashNormNombre(v.Asesor);
    return vMes === mes && vAnio === anio && vAsesor === asesorNorm;
  });

  // Agrupar por división usando el mapa de grupos
  const porDivision = {};
  ventasFiltradas.forEach(v => {
    const grupo = String(v.GrupoArticulo || '').trim();
    const division = DASHBOARD_CONFIG.mapaGrupos[grupo];
    if (!division) return; // ignorar (ej. Vending Machines)
    porDivision[division] = (porDivision[division] || 0) + parseFloat(v.Total || 0);
  });

  const totalVenta = ventasFiltradas.reduce((s, v) => s + parseFloat(v.Total || 0), 0);
  const numOVs = new Set(ventasFiltradas.map(v => v.NumOV)).size;

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
    const div = String(p.Division || '').trim();
    // Si tiene columna División, sumar solo una vez por división
    // Si no tiene División, sumar la columna Meta directamente
    return s + parseFloat(p.Meta || 0);
  }, 0);

  // Meta por división
  const metaPorDivision = {};
  presupFiltrado.forEach(p => {
    const div = String(p.Division || p.División || '').trim();
    if (div) {
      metaPorDivision[div] = (metaPorDivision[div] || 0) + parseFloat(p.Meta || 0);
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
function dashHistoricoAnual(asesor, anio) {
  const meses = [];
  for (let m = 1; m <= 12; m++) {
    const met = dashCalcMetricas(asesor, m, anio);
    const metAnterior = dashCalcMetricas(asesor, m, anio - 1);
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
  // Retorna nombres como aparecen en el presupuesto (para mostrar en UI)
  // pero internamente se buscan en SAP con alias
  const seen = new Set();
  const result = [];
  DASH_STATE.presupuesto.forEach(p => {
    const nombre = String(p.Asesor || '').trim();
    if (nombre && !seen.has(nombre)) {
      seen.add(nombre);
      result.push(nombre);
    }
  });
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
          total: parseFloat(v.Total || 0),
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

// ── Render principal ─────────────────────────────────────────────────────────
function dashRender() {
  const container = document.getElementById('dashboard-container');
  if (!container) return;

  const rol = DASH_STATE.rol;
  const mes = DASH_STATE.mesSel;
  const anio = DASH_STATE.anioSel;
  const isCurrent = mes === DASH_STATE.mesActual && anio === DASH_STATE.anioActual;
  const mesLabel = DASHBOARD_CONFIG.meses[mes-1] + ' ' + anio;

  // Nombre del asesor actual (para vista asesor)
  const nombreAsesor = CONFIG.usuario.nombre;

  if (rol === 'asesor') {
    dashRenderAsesor(container, nombreAsesor, mes, anio, isCurrent, mesLabel);
  } else if (rol === 'lider') {
    dashRenderLider(container, mes, anio, isCurrent, mesLabel);
  } else {
    dashRenderGerente(container, mes, anio, isCurrent, mesLabel);
  }
}

// ── Render vista Asesor ──────────────────────────────────────────────────────
function dashRenderAsesor(container, asesor, mes, anio, isCurrent, mesLabel) {
  const met = dashCalcMetricas(asesor, mes, anio);
  const hist = dashHistoricoAnual(asesor, anio);
  const riesgo = dashClientesEnRiesgo(asesor, 60);
  const ctx = dashMensajeCtx(met.pct, met.falta, met.diasRestantes, met.ritmo, isCurrent);
  const col = dashColor(met.pct);

  container.innerHTML = `
    <div class="dash-tabs">
      <button class="dash-tab active" onclick="dashTabSwitch('resultados',this)">Resultados</button>
      <button class="dash-tab" onclick="dashTabSwitch('pipeline',this)">Pipeline</button>
      <button class="dash-tab" onclick="dashTabSwitch('riesgo',this)">En riesgo</button>
    </div>

    <div id="dash-page-resultados" class="dash-page active">
      ${dashPeriodSelector(mesLabel, isCurrent, 'dashCambiarMes')}
      ${dashStreakHtml()}
      <div class="dash-ctx dash-ctx-${ctx.tipo}"><div class="dash-cdot"></div><p>${ctx.msg}</p></div>
      <div class="dash-sec"><div class="dash-lbl">Progreso vs meta — ${mesLabel}</div>
        <div class="dash-goal-card">
          <div class="dash-goal-top"><div class="dash-goal-amt">${dashFmt(met.totalVenta)}</div><div class="dash-goal-tgt">meta ${dashFmt(met.totalMeta)}</div></div>
          <div class="dash-pbar"><div class="dash-pfill" id="dash-pf-main" style="width:0%;background:${col.bar}"></div></div>
          <div class="dash-goal-row">
            <div class="dash-gs"><span class="dash-v" style="color:${col.pct}">${met.pct}%</span><span class="dash-l">alcanzado</span></div>
            <div class="dash-gs" style="text-align:center"><span class="dash-v">${dashFmt(met.falta)}</span><span class="dash-l">${isCurrent ? 'te falta' : 'faltó'}</span></div>
            <div class="dash-gs" style="text-align:right"><span class="dash-v">${isCurrent ? met.diasRestantes + ' días' : 'Cerrado'}</span><span class="dash-l">${isCurrent ? 'restantes' : ''}</span></div>
          </div>
        </div>
      </div>
      <div class="dash-kpis">
        <div class="dash-kpi"><div class="dash-kpi-v">${met.numOVs}</div><div class="dash-kpi-l">OVs generadas</div></div>
        <div class="dash-kpi"><div class="dash-kpi-v">${riesgo.length}</div><div class="dash-kpi-l">Clientes en riesgo</div><div class="dash-kpi-d" style="color:#A32D2D">+60 días sin compra</div></div>
        <div class="dash-kpi"><div class="dash-kpi-v">${Object.keys(met.porDivision).length}</div><div class="dash-kpi-l">Divisiones activas</div></div>
        <div class="dash-kpi"><div class="dash-kpi-v">${dashFmt(met.totalMeta > 0 ? met.totalVenta/met.totalMeta*100 : 0).replace('$','').replace('K','') + '%'}</div><div class="dash-kpi-l">Cumplimiento</div></div>
      </div>
      ${dashVentasDivisionHtml(met.porDivision, met.metaPorDivision)}
      ${dashResumenAnualHtml(hist, anio)}
    </div>

    <div id="dash-page-pipeline" class="dash-page">
      ${dashPipelineHtml(asesor)}
    </div>

    <div id="dash-page-riesgo" class="dash-page">
      ${dashRiesgoHtml(riesgo)}
    </div>
  `;

  setTimeout(() => {
    document.getElementById('dash-pf-main').style.width = Math.min(met.pct, 100) + '%';
    dashInitChart(hist);
  }, 100);
}

// ── Render vista Líder ───────────────────────────────────────────────────────
function dashRenderLider(container, mes, anio, isCurrent, mesLabel) {
  const asesores = dashGetAsesores();
  // Calcular métricas de todos los asesores
  const equipo = asesores.map(a => {
    const met = dashCalcMetricas(a, mes, anio);
    return { nombre: a, ...met };
  }).filter(a => a.totalMeta > 0 || a.totalVenta > 0);

  const totalEquipo = equipo.reduce((s, a) => s + a.totalVenta, 0);
  const metaEquipo  = equipo.reduce((s, a) => s + a.totalMeta, 0);
  const pctEquipo   = metaEquipo > 0 ? Math.round(totalEquipo / metaEquipo * 100) : 0;
  const col = dashColor(pctEquipo);

  const enRiesgo = dashClientesEnRiesgo(null, 60);

  container.innerHTML = `
    <div class="dash-tabs">
      <button class="dash-tab active" onclick="dashTabSwitch('l-resultados',this)">Resultados</button>
      <button class="dash-tab" onclick="dashTabSwitch('l-riesgo',this)">En riesgo</button>
    </div>

    <div id="dash-page-l-resultados" class="dash-page active">
      ${dashPeriodSelector(mesLabel, isCurrent, 'dashCambiarMes')}
      <div class="dash-ctx dash-ctx-${pctEquipo >= 65 ? 'info' : 'warn'}">
        <div class="dash-cdot"></div>
        <p>El equipo va en <strong>${pctEquipo}%</strong> de meta para ${mesLabel}. ${equipo.filter(a=>a.pct<50).length} asesores por debajo del 50%.</p>
      </div>
      <div class="dash-sec"><div class="dash-lbl">Equipo · ${mesLabel}</div>
        <div class="dash-goal-card">
          <div class="dash-goal-top"><div class="dash-goal-amt">${dashFmt(totalEquipo)}</div><div class="dash-goal-tgt">meta ${dashFmt(metaEquipo)}</div></div>
          <div class="dash-pbar"><div class="dash-pfill" id="dash-pf-lider" style="width:0%;background:${col.bar}"></div></div>
          <div class="dash-goal-row">
            <div class="dash-gs"><span class="dash-v" style="color:${col.pct}">${pctEquipo}%</span><span class="dash-l">alcanzado</span></div>
            <div class="dash-gs" style="text-align:center"><span class="dash-v">${dashFmt(Math.max(0,metaEquipo-totalEquipo))}</span><span class="dash-l">falta al equipo</span></div>
            <div class="dash-gs" style="text-align:right"><span class="dash-v">${isCurrent ? diasEnMes(mes,anio)-new Date().getDate()+' días' : 'Cerrado'}</span><span class="dash-l">${isCurrent?'restantes':''}</span></div>
          </div>
        </div>
      </div>
      <div class="dash-lbl" style="padding:12px 16px 6px">Mi equipo · detalle</div>
      ${equipo.map(a => dashAsesorCardHtml(a, false)).join('')}
    </div>

    <div id="dash-page-l-riesgo" class="dash-page">
      ${dashRiesgoHtml(enRiesgo, true)}
    </div>
  `;

  setTimeout(() => {
    document.getElementById('dash-pf-lider').style.width = Math.min(pctEquipo, 100) + '%';
    equipo.forEach(a => {
      const el = document.getElementById('dash-fill-' + dashSlug(a.nombre));
      if (el) setTimeout(() => el.style.width = Math.min(a.pct, 100) + '%', 150);
    });
  }, 100);
}

// ── Render vista Gerente ─────────────────────────────────────────────────────
function dashRenderGerente(container, mes, anio, isCurrent, mesLabel) {
  const asesores = dashGetAsesores();
  const equipo = asesores.map(a => {
    const met = dashCalcMetricas(a, mes, anio);
    return { nombre: a, ...met };
  }).filter(a => a.totalMeta > 0 || a.totalVenta > 0);

  const totalEmpresa = equipo.reduce((s, a) => s + a.totalVenta, 0);
  const metaEmpresa  = equipo.reduce((s, a) => s + a.totalMeta, 0);
  const pctEmpresa   = metaEmpresa > 0 ? Math.round(totalEmpresa / metaEmpresa * 100) : 0;
  const col = dashColor(pctEmpresa);
  const enRiesgo = dashClientesEnRiesgo(null, 60);

  container.innerHTML = `
    <div class="dash-tabs">
      <button class="dash-tab active" onclick="dashTabSwitch('g-resumen',this)">Resumen</button>
      <button class="dash-tab" onclick="dashTabSwitch('g-asesores',this)">Asesores</button>
      <button class="dash-tab" onclick="dashTabSwitch('g-riesgo',this)">En riesgo</button>
    </div>

    <div id="dash-page-g-resumen" class="dash-page active">
      ${dashPeriodSelector(mesLabel, isCurrent, 'dashCambiarMes')}
      <div class="dash-ctx dash-ctx-${pctEmpresa >= 65 ? 'info' : 'warn'}">
        <div class="dash-cdot"></div>
        <p>Empresa va en <strong>${pctEmpresa}%</strong> de meta para ${mesLabel}. ${equipo.filter(a=>a.pct<50).length} asesores requieren atención.</p>
      </div>
      <div class="dash-sec"><div class="dash-lbl">Vinssa · ${mesLabel}</div>
        <div class="dash-goal-card">
          <div class="dash-goal-top"><div class="dash-goal-amt">${dashFmt(totalEmpresa)}</div><div class="dash-goal-tgt">meta ${dashFmt(metaEmpresa)}</div></div>
          <div class="dash-pbar"><div class="dash-pfill" id="dash-pf-gerente" style="width:0%;background:${col.bar}"></div></div>
          <div class="dash-goal-row">
            <div class="dash-gs"><span class="dash-v" style="color:${col.pct}">${pctEmpresa}%</span><span class="dash-l">alcanzado</span></div>
            <div class="dash-gs" style="text-align:center"><span class="dash-v">${dashFmt(Math.max(0,metaEmpresa-totalEmpresa))}</span><span class="dash-l">falta total</span></div>
            <div class="dash-gs" style="text-align:right"><span class="dash-v">${isCurrent ? diasEnMes(mes,anio)-new Date().getDate()+' días' : 'Cerrado'}</span><span class="dash-l">${isCurrent?'restantes':''}</span></div>
          </div>
        </div>
      </div>
      <div class="dash-grid2">
        <div class="dash-mc"><div class="dash-mc-v">${equipo.filter(a=>a.pct>=65).length}/${equipo.length}</div><div class="dash-mc-l">Asesores en meta</div></div>
        <div class="dash-mc"><div class="dash-mc-v">${enRiesgo.length}</div><div class="dash-mc-l">Clientes en riesgo</div><div class="dash-mc-d" style="color:#A32D2D">+60 días sin compra</div></div>
      </div>
    </div>

    <div id="dash-page-g-asesores" class="dash-page">
      <div class="dash-filter-row">
        <button class="dash-fchip active" onclick="dashFiltrarAsesores('todos',this,${JSON.stringify(equipo).replace(/'/g,"\\'")})">Todos</button>
        <button class="dash-fchip" onclick="dashFiltrarAsesores('riesgo',this)">En riesgo</button>
        <button class="dash-fchip" onclick="dashFiltrarAsesores('meta',this)">En meta</button>
      </div>
      <div id="dash-asesores-lista">
        ${equipo.map(a => dashAsesorCardHtml(a, true)).join('')}
      </div>
    </div>

    <div id="dash-page-g-riesgo" class="dash-page">
      ${dashRiesgoHtml(enRiesgo, true)}
    </div>
  `;

  setTimeout(() => {
    document.getElementById('dash-pf-gerente').style.width = Math.min(pctEmpresa, 100) + '%';
    equipo.forEach(a => {
      const el = document.getElementById('dash-fill-' + dashSlug(a.nombre));
      if (el) setTimeout(() => el.style.width = Math.min(a.pct, 100) + '%', 150);
    });
  }, 100);
}

// ── Componentes HTML ─────────────────────────────────────────────────────────

function dashPeriodSelector(mesLabel, isCurrent, fn) {
  const meses = DASHBOARD_CONFIG.meses;
  const mesActual = DASH_STATE.mesActual;
  return `
    <div class="dash-period-selector">
      <button class="dash-period-btn" onclick="dashTogglePeriodPanel()">
        <span>📅 ${mesLabel}</span>
        <span style="font-size:10px">▾</span>
      </button>
      ${isCurrent ? '<div class="dash-period-badge">Mes actual</div>' : '<div class="dash-period-badge" style="background:#E6F1FB;color:#0C447C">Histórico</div>'}
    </div>
    <div class="dash-period-panel" id="dash-period-panel">
      <div class="dash-month-grid">
        ${meses.map((m, i) => {
          const num = i + 1;
          const futuro = num > mesActual && DASH_STATE.anioSel >= DASH_STATE.anioActual;
          const seleccionado = num === DASH_STATE.mesSel;
          return `<button class="dash-mg-btn${seleccionado?' active':''}${futuro?' disabled':''}"
            ${futuro ? 'disabled' : `onclick="dashSelMes(${num},this)"`}
            >${m}${num === mesActual ? ' ●' : ''}</button>`;
        }).join('')}
      </div>
    </div>
  `;
}

function dashStreakHtml() {
  // Por ahora la racha la calculamos basándonos en semanas con OVs registradas
  // Cuando tengamos datos reales de visitas desde Field App se conecta
  return `
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
  `;
}

function dashVentasDivisionHtml(porDivision, metaPorDivision) {
  if (!Object.keys(porDivision).length) return '';
  const divs = Object.entries(porDivision).sort((a,b) => b[1]-a[1]);
  return `
    <div class="dash-sec" style="padding-top:12px">
      <div class="dash-lbl">Por división</div>
      ${divs.map(([div, venta]) => {
        const meta = metaPorDivision[div] || 0;
        const pct = meta > 0 ? Math.round(venta/meta*100) : 0;
        const col = dashColor(pct);
        return `
          <div style="margin-bottom:10px">
            <div style="display:flex;justify-content:space-between;margin-bottom:4px">
              <span style="font-size:12px;color:var(--color-text-primary);font-weight:500">${div}</span>
              <span style="font-size:12px;color:${col.pct};font-weight:500">${dashFmt(venta)}${meta>0?' / '+dashFmt(meta):''}</span>
            </div>
            <div class="dash-pbar"><div class="dash-pfill" style="width:${Math.min(pct,100)}%;background:${col.bar}"></div></div>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function dashResumenAnualHtml(hist, anio) {
  const pctAnual = hist.metaAnual > 0 ? Math.round(hist.acumulado / hist.metaAnual * 100) : 0;
  const vsAnterior = hist.acumuladoAnterior > 0
    ? Math.round((hist.acumulado - hist.acumuladoAnterior) / hist.acumuladoAnterior * 100)
    : 0;
  return `
    <div class="dash-annual">
      <div class="dash-annual-title">Resumen ${anio}</div>
      <div class="dash-annual-row">
        <div class="dash-ar-lbl">Acumulado</div>
        <div style="text-align:right">
          <div class="dash-ar-val">${dashFmt(hist.acumulado)}</div>
          <div class="dash-ar-sub ${vsAnterior>=0?'dash-up':'dash-dn'}">${vsAnterior>=0?'+':''}${vsAnterior}% vs ${anio-1}</div>
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
          <div class="dash-ar-sub" style="color:#BA7517">${hist.mejorMes.pct >= 100 ? '🏆 ' : ''}${hist.mejorMes.pct}% de meta</div>
        </div>
      </div>` : ''}
      <div class="dash-annual-row">
        <div class="dash-ar-lbl">Proyección anual</div>
        <div style="text-align:right">
          <div class="dash-ar-val">${dashFmt(hist.proyeccion)}</div>
          <div class="dash-ar-sub ${hist.proyeccion >= hist.metaAnual ? 'dash-up' : 'dash-dn'}">${hist.metaAnual > 0 ? Math.round(hist.proyeccion/hist.metaAnual*100)+'% vs meta' : ''}</div>
        </div>
      </div>
      <canvas id="dash-chart-hist" style="margin-top:12px;width:100%;height:120px;max-height:120px" role="img" aria-label="Histórico mensual de ventas"></canvas>
    </div>
  `;
}

function dashAsesorCardHtml(a, clickable) {
  const col = dashColor(a.pct);
  const slug = dashSlug(a.nombre);
  const initials = a.nombre.split(' ').slice(0,2).map(n=>n[0]||'').join('').toUpperCase();
  const statusLabel = a.pct >= 65 ? 'En meta' : a.pct >= 40 ? 'Seguimiento' : 'En riesgo';
  const statusClass = a.pct >= 65 ? 'pill-grn' : a.pct >= 40 ? 'pill-amb' : 'pill-red';
  return `
    <div class="dash-ac ${a.pct < 40 ? 'dash-ac-risk' : a.pct < 65 ? 'dash-ac-warn' : ''}"
      ${clickable ? `onclick="dashDrillAsesor('${a.nombre}')"` : ''} style="margin-bottom:8px">
      <div class="dash-ac-top">
        <div class="dash-acav" style="background:${col.bar}22;color:${col.pct}">${initials}</div>
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
        <div class="dash-ac-kpi"><span>${a.diasRestantes > 0 ? a.diasRestantes + 'd' : '—'}</span>días</div>
      </div>
    </div>
  `;
}

function dashPipelineHtml(asesor) {
  const asesorNorm = asesor ? dashNormPresup(asesor) : null;

  // Filtrar OVs por asesor si aplica
  const ovsFiltradas = DASH_STATE.ovs.filter(v => {
    if (!asesorNorm) return true;
    return dashNormNombre(v.Asesor) === asesorNorm;
  });

  if (!ovsFiltradas.length) {
    return `<div style="padding:24px 16px;text-align:center;color:var(--color-text-secondary);font-size:13px">
      No hay órdenes de venta abiertas.
    </div>`;
  }

  // Calcular total y agrupar por grupo de artículo / división
  const totalOVs = ovsFiltradas.reduce((s, v) => s + parseFloat(v.Total || 0), 0);
  const numOVs   = new Set(ovsFiltradas.map(v => v.NumOV)).size;

  // Agrupar por división
  const porDiv = {};
  ovsFiltradas.forEach(v => {
    const grupo = String(v.GrupoArticulo || '').trim();
    const div   = DASHBOARD_CONFIG.mapaGrupos[grupo] || 'Otros';
    if (!div || div === null) return;
    porDiv[div] = (porDiv[div] || 0) + parseFloat(v.Total || 0);
  });

  // Top 10 OVs por monto
  const topOVs = Object.values(
    ovsFiltradas.reduce((acc, v) => {
      const key = v.NumOV;
      if (!acc[key]) {
        acc[key] = {
          numOV: v.NumOV,
          cliente: String(v.Cliente || '').trim(),
          asesor: String(v.Asesor || '').trim(),
          fecha: v.Fecha,
          fechaEntrega: v.FechaEntrega || '',
          total: 0,
          division: DASHBOARD_CONFIG.mapaGrupos[String(v.GrupoArticulo||'').trim()] || 'Otros'
        };
      }
      acc[key].total += parseFloat(v.Total || 0);
      return acc;
    }, {})
  ).sort((a, b) => b.total - a.total).slice(0, 10);

  const divEntries = Object.entries(porDiv).sort((a, b) => b[1] - a[1]);
  const maxDiv = divEntries.length ? divEntries[0][1] : 1;

  return `
    <div class="pipe-wrap" style="padding-top:14px">
      <div class="pipe-hdr">
        <div>
          <div class="pipe-total">${dashFmt(totalOVs)}</div>
          <div class="pipe-sub">pipeline comprometido · ${numOVs} OVs abiertas</div>
        </div>
      </div>
      <div class="dash-ctx dash-ctx-info" style="margin:0 0 12px">
        <div class="dash-cdot"></div>
        <p>Órdenes de venta <strong>abiertas</strong> — comprometidas pero aún no facturadas.</p>
      </div>

      ${divEntries.map(([div, total]) => {
        const pct = Math.round(total / totalOVs * 100);
        const w   = Math.round(total / maxDiv * 100);
        return `<div style="margin-bottom:10px;padding:0 16px">
          <div style="display:flex;justify-content:space-between;margin-bottom:4px">
            <span style="font-size:12px;font-weight:500;color:var(--color-text-primary)">${div}</span>
            <span style="font-size:12px;color:var(--color-text-secondary)">${dashFmt(total)} · ${pct}%</span>
          </div>
          <div class="dash-pbar"><div class="dash-pfill" style="width:${w}%;background:#185FA5"></div></div>
        </div>`;
      }).join('')}

      <div style="padding:10px 16px 0">
        <div class="dash-lbl">Top OVs por monto</div>
        ${topOVs.map(ov => `
          <div style="display:flex;align-items:flex-start;gap:10px;padding:9px 0;border-bottom:0.5px solid var(--color-border-tertiary)">
            <div style="flex:1">
              <div style="font-size:13px;font-weight:500;color:var(--color-text-primary)">${ov.cliente}</div>
              <div style="font-size:11px;color:var(--color-text-secondary);margin-top:2px">
                OV #${ov.numOV} · ${ov.division}${ov.fechaEntrega ? ' · Entrega: '+ov.fechaEntrega : ''}
                ${!asesor ? ' · '+ov.asesor : ''}
              </div>
            </div>
            <div style="font-size:13px;font-weight:500;color:var(--color-text-primary);flex-shrink:0">${dashFmt(ov.total)}</div>
          </div>
        `).join('')}
      </div>
    </div>
  `;
}

function dashRiesgoHtml(clientes, mostrarAsesor = false) {
  if (!clientes.length) {
    return `<div style="padding:24px 16px;text-align:center;color:var(--color-text-secondary);font-size:13px">
      ✅ No hay clientes sin compra en más de 60 días.
    </div>`;
  }
  return `
    <div class="dash-sec" style="padding-top:14px">
      <div class="dash-lbl">Clientes sin compra reciente</div>
      <div class="dash-ctx dash-ctx-warn" style="margin:0 0 4px">
        <div class="dash-cdot"></div>
        <p><strong>${clientes.length} clientes</strong> sin movimiento en +60 días.</p>
      </div>
    </div>
    <div class="dash-risk-list">
      ${clientes.slice(0,20).map(c => `
        <div class="dash-rki">
          <div class="dash-rkdot" style="background:${c.dias > 90 ? '#E24B4A' : '#EF9F27'}"></div>
          <div class="dash-rk-info">
            <div class="dash-rk-name">${c.cliente}</div>
            <div class="dash-rk-sub">
              Último pedido: ${c.fecha.toLocaleDateString('es-MX')} · ${dashFmt(c.total)}
              ${mostrarAsesor ? `<br>Asesor: ${c.asesor}` : ''}
              ${c.grupo ? ` · ${DASHBOARD_CONFIG.mapaGrupos[c.grupo] || c.grupo}` : ''}
            </div>
          </div>
          <div style="text-align:right;flex-shrink:0">
            <div style="font-size:13px;font-weight:500;color:var(--color-text-primary)">${dashFmt(c.total)}</div>
            <div style="font-size:11px;color:${c.dias > 90 ? '#A32D2D' : '#BA7517'};margin-top:2px">${c.dias} días</div>
          </div>
        </div>
      `).join('')}
    </div>
  `;
}

// ── Inicializar gráfica histórica ────────────────────────────────────────────
function dashInitChart(hist) {
  const canvas = document.getElementById('dash-chart-hist');
  if (!canvas || !window.Chart) return;
  const labels = hist.meses.map(m => m.label);
  const ventas = hist.meses.map(m => Math.round(m.venta));
  const metas  = hist.meses.map(m => Math.round(m.meta));
  new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Venta', data: ventas, backgroundColor: '#185FA5', borderRadius: 3 },
        { label: 'Meta',  data: metas,  backgroundColor: '#D3D1C7', borderRadius: 3 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 9 }, color: '#888780' } },
        y: { display: false }
      }
    }
  });
}

// ── Navegación tabs ──────────────────────────────────────────────────────────
function dashTabSwitch(id, btn) {
  document.querySelectorAll('#dashboard-container .dash-tab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('#dashboard-container .dash-page').forEach(p => p.classList.remove('active'));
  const page = document.getElementById('dash-page-' + id);
  if (page) page.classList.add('active');
}

function dashTogglePeriodPanel() {
  const panel = document.getElementById('dash-period-panel');
  if (panel) panel.classList.toggle('open');
}

function dashSelMes(mes, btn) {
  DASH_STATE.mesSel = mes;
  const panel = document.getElementById('dash-period-panel');
  if (panel) panel.classList.remove('open');
  dashRender();
}

function dashFiltrarAsesores(filtro, btn) {
  document.querySelectorAll('.dash-fchip').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  const lista = document.getElementById('dash-asesores-lista');
  if (!lista) return;
  const cards = lista.querySelectorAll('.dash-ac');
  cards.forEach(card => {
    const pct = parseInt(card.querySelector('[style*="font-size:15px"]')?.textContent || '100');
    if (filtro === 'todos') card.style.display = '';
    else if (filtro === 'riesgo') card.style.display = pct < 50 ? '' : 'none';
    else if (filtro === 'meta') card.style.display = pct >= 65 ? '' : 'none';
  });
}

function dashDrillAsesor(nombre) {
  // Por ahora muestra un alert — en siguiente iteración abre vista drill-down
  alert(`Vista detalle de ${nombre} — próximamente`);
}

// ── Utilidades ────────────────────────────────────────────────────────────────
function dashSlug(str) {
  return str.toLowerCase().replace(/[^a-z0-9]/g, '-');
}

function dashMostrarLoader(show) {
  const loader = document.getElementById('dash-loader');
  if (loader) loader.style.display = show ? 'flex' : 'none';
}

function dashMostrarError(msg) {
  const container = document.getElementById('dashboard-container');
  if (container) container.innerHTML = `
    <div style="padding:24px;text-align:center">
      <div style="font-size:32px;margin-bottom:12px">⚠️</div>
      <div style="font-size:14px;color:var(--color-text-primary);font-weight:500;margin-bottom:8px">Error cargando dashboard</div>
      <div style="font-size:12px;color:var(--color-text-secondary)">${msg}</div>
      <button onclick="dashInit()" style="margin-top:16px;padding:10px 20px;background:#185FA5;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px">Reintentar</button>
    </div>
  `;
}

// Exponer globalmente
window.dashTabSwitch      = dashTabSwitch;
window.dashTogglePeriodPanel = dashTogglePeriodPanel;
window.dashSelMes         = dashSelMes;
window.dashFiltrarAsesores = dashFiltrarAsesores;
window.dashDrillAsesor    = dashDrillAsesor;
window.dashCambiarMes     = dashSelMes;
window.dashInit           = dashInit;
window.dashNormPresup     = dashNormPresup;
