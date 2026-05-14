// ── SAP B1 Service Layer + SharePoint ────────────────────────────────────────

const SAP = {
  baseUrl: CONFIG.sap.baseUrl,
  sessionId: null,
  sessionExpira: null
};

const SHAREPOINT = {
  site: 'https://versatilidadsaltillo.sharepoint.com/sites/VINSSAAutomation',
  tabla: 'TablaVisitas',
  archivo: 'Tracker Visitas Field App.xlsx'
};

// ── Autenticación SAP ────────────────────────────────────────────────────────

async function sapLogin(usuario, password) {
  try {
    const res = await fetch(`${SAP.baseUrl}/Login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        CompanyDB: CONFIG.sap.company,
        UserName: usuario,
        Password: password
      }),
      credentials: 'include'
    });

    if (!res.ok) throw new Error('Credenciales incorrectas');
    const data = await res.json();
    SAP.sessionId = data.SessionId;
    SAP.sessionExpira = Date.now() + (30 * 60 * 1000);
    localStorage.setItem('sap_session', JSON.stringify({
      id: SAP.sessionId,
      expira: SAP.sessionExpira,
      usuario: usuario
    }));
    return true;
  } catch(e) {
    console.error('SAP login error:', e);
    return false;
  }
}

async function sapEnsureSession() {
  const guardada = JSON.parse(localStorage.getItem('sap_session') || 'null');
  if (guardada && guardada.expira > Date.now()) {
    SAP.sessionId = guardada.id;
    return true;
  }
  return false;
}

function sapHeaders() {
  return {
    'Content-Type': 'application/json',
    'B1SESSION': SAP.sessionId,
    'Prefer': 'return=representation'
  };
}

// ── Crear oportunidad en SAP ─────────────────────────────────────────────────

async function sapCrearOportunidad(registro) {
  const etapaPct = CONFIG.etapas.find(e => e.id === registro.etapa)?.pct || 10;

  const body = {
    CardCode: registro.clienteCode || null,
    CardName: registro.cliente,
    ContactPersonCode: null,
    SalesPerson: CONFIG.usuario.sapCode || null,
    OpportunityName: registro.oppNombre || `${registro.marca}/${registro.producto}/${registro.cliente}`.slice(0, 50),
    Status: 'O',
    ClosingDate: registro.cierre || null,
    CloseDate: registro.cierre || null,
    MaxLocalTotal: parseFloat(registro.monto) || 0,
    WeightedAmountLocal: (parseFloat(registro.monto) || 0) * (etapaPct / 100),
    U_Division: registro.division,
    U_Marca: registro.marca,
    SalesOpportunityLines: [
      {
        SalesPerson: CONFIG.usuario.sapCode || null,
        StartDate: new Date().toISOString().split('T')[0],
        ClosingDate: registro.cierre || new Date().toISOString().split('T')[0],
        StageKey: etapaPct,
        PercentageRate: etapaPct
      }
    ],
    SalesOpportunityInterests: [],
    SalesOpportunityCompetitors: registro.competidores
      .filter(c => c !== 'Ninguno' && c !== 'Otro')
      .map(c => ({ CompetitorCode: null, CompetitorName: c }))
  };

  const res = await fetch(`${SAP.baseUrl}/SalesOpportunities`, {
    method: 'POST',
    headers: sapHeaders(),
    body: JSON.stringify(body),
    credentials: 'include'
  });

  if (!res.ok) {
    const err = await res.json();
    throw new Error(err?.error?.message?.value || 'Error al crear oportunidad en SAP');
  }
  return await res.json();
}

// ── Crear actividad en SAP ───────────────────────────────────────────────────

async function sapCrearActividad(registro, oppId) {
  const tipoActividad = registro.tipo === 'demo' ? 'D' : 'V';
  const now = new Date();

  const body = {
    CardCode: registro.clienteCode || null,
    CardName: registro.cliente,
    ActivityDate: now.toISOString().split('T')[0],
    ActivityTime: now.toTimeString().slice(0, 8),
    Details: registro.notas || `${registro.tipo.toUpperCase()} - ${registro.marca} ${registro.producto}`,
    Activity: 'C',
    ActivityType: tipoActividad === 'D' ? 'Demo' : 'Visita',
    SalesEmployee: CONFIG.usuario.sapCode || null,
    OpportunityId: oppId || null,
    U_GPS_Lat: registro.gps?.lat?.toString() || null,
    U_GPS_Lng: registro.gps?.lng?.toString() || null,
    Notes: [
      `Asesor: ${registro.asesor}`,
      `Marca: ${registro.marca} | Producto: ${registro.producto}`,
      `Etapa: ${registro.etapa} | Monto: ${registro.moneda} ${registro.monto}`,
      registro.competidores.length ? `Competidor: ${registro.competidores.join(', ')}` : '',
      registro.notas || ''
    ].filter(Boolean).join('\n')
  };

  const res = await fetch(`${SAP.baseUrl}/Activities`, {
    method: 'POST',
    headers: sapHeaders(),
    body: JSON.stringify(body),
    credentials: 'include'
  });

  if (!res.ok) {
    const err = await res.json();
    throw new Error(err?.error?.message?.value || 'Error al crear actividad en SAP');
  }
  return await res.json();
}

// ── Crear contacto nuevo en SAP ──────────────────────────────────────────────

async function sapCrearContacto(registro) {
  if (!registro.contactoNuevo || !registro.contactoNuevo.nombre) return null;
  if (!registro.clienteCode) return null;

  const c = registro.contactoNuevo;

  const res = await fetch(`${SAP.baseUrl}/BusinessPartners('${registro.clienteCode}')`, {
    method: 'PATCH',
    headers: sapHeaders(),
    body: JSON.stringify({
      ContactEmployees: [{
        Name: c.nombre,
        Position: c.puesto || '',
        Phone1: c.telefono || '',
        E_Mail: c.email || '',
        Active: 'tYES'
      }]
    }),
    credentials: 'include'
  });

  if (!res.ok) {
    console.warn('No se pudo agregar contacto a SAP:', await res.text());
    return null;
  }
  return true;
}

// ── Guardar en SharePoint (via Power Automate) ───────────────────────────────

async function guardarEnSharePoint(registro, sapOppId, sapActId) {
  const geoRegistros = JSON.parse(localStorage.getItem('vinssa_geo_registros') || '[]');
  const ultimoCheckin = geoRegistros.filter(r => r.tipo === 'checkin' && r.cliente === registro.cliente).pop();
  const ultimoCheckout = geoRegistros.filter(r => r.tipo === 'checkout' && r.cliente === registro.cliente).pop();
  const esGabinete = registro.modo === 'gabinete';

  const fields = {
    Title: registro.cliente || '',
    // Identidad canónica del cliente para cruce con SAP (entidad OCRD).
    // Vacío cuando es prospecto/cliente nuevo capturado por texto libre — esos
    // casos se cuentan aparte para medir disciplina de captura en SAP.
    CardCode: registro.cardCode || '',
    Asesor: registro.asesor || '',
    Tipo: registro.tipo || '',
    ModoRegistro: registro.modo || 'campo',
    ResultadoCierre: registro.resultadoCierre || '',
    RazonPerdida: registro.razonPerdida || '',
    OppNombre: registro.oppNombre || '',
    OpportunidadID: registro.opportunidadID || '',
    Marca: registro.marca || '',
    Producto: registro.producto || '',
    Etapa: registro.etapa || '',
    Monto: parseFloat(registro.monto) || 0,
    MontoFinal: registro.montoFinal ? parseFloat(registro.montoFinal) : null,
    Moneda: registro.moneda || 'MXP',
    Competitor: registro.competidores?.join(', ') || '',
    Lideres: registro.lideres?.join(', ') || '',
    Notas: registro.notas || '',
    GPS_Lat: esGabinete ? null : (registro.gps?.lat || 0),
    GPS_Lng: esGabinete ? null : (registro.gps?.lng || 0),
    GPS_Precision: esGabinete ? null : (registro.gps?.precision || 0),
    GPS_Disponible: registro.gpsDisponible ? 'Si' : 'No',
    Checkin_Hora: esGabinete ? '' : (ultimoCheckin?.hora || ''),
    Checkout_Hora: esGabinete ? '' : (ultimoCheckout?.horaSalida || ''),
    Duracion_Min: esGabinete ? null : (ultimoCheckout?.duracionMinutos || 0),
    Contacto_Nuevo: registro.contactoNuevo ? 'Si' : 'No',
    Contacto_Nombre: registro.contactoNuevo?.nombre || '',
    Contacto_Puesto: registro.contactoNuevo?.puesto || '',
    Contacto_Tel: registro.contactoNuevo?.telefono || '',
    Contacto_Email: registro.contactoNuevo?.email || '',
    SAP_OppID: sapOppId ? sapOppId.toString() : '',
    SAP_ActID: sapActId ? sapActId.toString() : '',
    Estatus_SAP: 'Pendiente'
  };

  // Agregar fecha de cierre solo si existe
  if (registro.cierre) {
    fields.Cierre = new Date(registro.cierre).toISOString();
  }
  if (registro.fechaCierreReal) {
    fields.FechaCierreReal = new Date(registro.fechaCierreReal).toISOString();
  }

  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ['Sites.ReadWrite.All'],
      account: msalInstance.getAllAccounts()[0]
    });

    // Obtener site ID via Graph
    const siteRes = await fetch(
      'https://graph.microsoft.com/v1.0/sites/versatilidadsaltillo.sharepoint.com:/sites/VINSSAAutomation',
      { headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` } }
    );
    const siteData = await siteRes.json();
    const siteId = siteData.id;

    // Obtener list ID
    const listRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$filter=displayName eq 'Visitas Field App'&$select=id`,
      { headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}` } }
    );
    const listData = await listRes.json();
    const listId = listData.value?.[0]?.id;

    if (!listId) throw new Error('Lista no encontrada');

    // Crear item en la lista
    const itemRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${tokenResponse.accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ fields })
      }
    );

    if (!itemRes.ok) {
      const err = await itemRes.json();
      throw new Error(err?.error?.message || 'Error al crear item');
    }

    const itemData = await itemRes.json();
    console.log('Guardado en SharePoint con Graph:', itemData.id);
    return itemData.id;

  } catch(e) {
    console.warn('SharePoint Graph error:', e.message);
    const respaldo = JSON.parse(localStorage.getItem('vinssa_registros_pendientes') || '[]');
    respaldo.push({ fields, timestamp: Date.now() });
    localStorage.setItem('vinssa_registros_pendientes', JSON.stringify(respaldo));
    return null;
  }
}
// ── Flujo principal de sincronización ───────────────────────────────────────

async function sincronizarConSAP(registro) {
  let sapOppId = null;
  let sapActId = null;

  // Guardar en SharePoint siempre — independiente de SAP
  await guardarEnSharePoint(registro, null, null);

  return { sapOppId, sapActId, errores: ['SAP: pendiente via Power Automate'] };
}

// ── Lectura de Oportunidades.xlsx vía Graph Workbook API ─────────────────────
// Cache propio — no compartir con DASH_STATE porque el dashboard lo limpia
// al entrar a la vista de resultados y eso invalidaría refrescos del field app.

const OPS_CACHE = { driveId: null };

// Normaliza nombres para match exacto con SAP: uppercase, sin acentos, trim.
// Misma convención que dashNormNombre en dashboard.js (copia local para no
// acoplar sap.js al orden de carga de scripts).
function _normNombre(str) {
  return String(str || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .trim();
}

// Devuelve array de oportunidades del asesor SAP que viene en STATE.asesorSAP.
// Si no hay asesor resuelto, o si Graph falla, devuelve []. No lanza errores
// hacia el caller: la app debe seguir usable sin oportunidades.
async function cargarOportunidadesAsesor() {
  if (typeof STATE === 'undefined' || !STATE.asesorSAP) return [];

  try {
    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) return [];
    const tokenRes = await msalInstance.acquireTokenSilent({
      scopes: ['Sites.ReadWrite.All'],
      account: accounts[0]
    });
    const token = tokenRes.accessToken;

    // Resolver driveId si no está cacheado
    if (!OPS_CACHE.driveId) {
      const drivesRes = await fetch(
        'https://graph.microsoft.com/v1.0/sites/versatilidadsaltillo.sharepoint.com:/sites/VINSSAAutomation:/drives',
        { headers: { Authorization: 'Bearer ' + token } }
      );
      const drives = await drivesRes.json();
      if (!drives.value || !drives.value.length) {
        console.warn('No se encontró drive de SharePoint');
        return [];
      }
      OPS_CACHE.driveId = drives.value[0].id;
    }
    const driveId = OPS_CACHE.driveId;

    // Buscar el archivo
    const itemsRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`,
      { headers: { Authorization: 'Bearer ' + token } }
    );
    const items = await itemsRes.json();
    const file = (items.value || []).find(f => f.name === 'Oportunidades.xlsx');
    if (!file) {
      console.warn('Oportunidades.xlsx no encontrado en SharePoint');
      return [];
    }

    // Leer la primera hoja — fallback a Hoja1 si Sheet1 no existe
    let res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${file.id}/workbook/worksheets/Sheet1/usedRange`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) {
      res = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${file.id}/workbook/worksheets/Hoja1/usedRange`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!res.ok) {
        console.warn('No se pudo leer Oportunidades.xlsx');
        return [];
      }
    }
    const data = await res.json();
    const values = data.values;
    if (!values || values.length < 2) return [];

    const headers = values[0].map(h => String(h).trim());
    const rows = values.slice(1)
      .filter(row => row.some(cell => cell !== '' && cell !== null))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i] ?? '');
        return obj;
      });

    // Filtrar por asesor del usuario logueado (match exacto normalizado)
    const asesorNorm = _normNombre(STATE.asesorSAP);
    const oportunidades = rows
      .filter(r => _normNombre(r.Asesor) === asesorNorm)
      .map(r => ({
        NumOportunidad: String(r.NumOportunidad || '').trim(),
        Cliente:        String(r.Cliente || '').trim(),
        Descripcion:    String(r.Descripcion || '').trim(),
        MontoEstimado:  parseFloat(r.MontoEstimado) || 0,
        Etapa:          String(r.Etapa || '').trim(),
        Marca:          String(r.Marca || '').trim(),
        Linea:          String(r.Linea || '').trim(),
        FechaCierre:    String(r.FechaCierre || '').trim()
      }));

    console.log(`Oportunidades cargadas para ${STATE.asesorSAP}: ${oportunidades.length}`);
    return oportunidades;
  } catch(e) {
    console.warn('Error cargando oportunidades:', e.message);
    return [];
  }
}

// ── Lectura de Clientes Activos.xlsx vía Graph Workbook API ──────────────────
// Devuelve array crudo de clientes; sap.js NO ordena por prioridad de asesor —
// eso es responsabilidad de app.js que conoce STATE.asesorSAP y el UX target.
// Reutiliza OPS_CACHE.driveId del cargador de oportunidades (mismo sitio).
//
// Cuando SAP B1 Service Layer esté disponible vía Vertrou, esta función se
// reescribe internamente para llamar a `BusinessPartners?$filter=CardType eq
// 'cCustomer' and validFor eq 'tYES'` y mapear los campos. La firma externa
// queda igual: array de objetos con las mismas keys.
async function cargarClientesActivos() {
  try {
    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) return [];
    const tokenRes = await msalInstance.acquireTokenSilent({
      scopes: ['Sites.ReadWrite.All'],
      account: accounts[0]
    });
    const token = tokenRes.accessToken;

    if (!OPS_CACHE.driveId) {
      const drivesRes = await fetch(
        'https://graph.microsoft.com/v1.0/sites/versatilidadsaltillo.sharepoint.com:/sites/VINSSAAutomation:/drives',
        { headers: { Authorization: 'Bearer ' + token } }
      );
      const drives = await drivesRes.json();
      if (!drives.value || !drives.value.length) {
        console.warn('No se encontró drive de SharePoint');
        return [];
      }
      OPS_CACHE.driveId = drives.value[0].id;
    }
    const driveId = OPS_CACHE.driveId;

    const itemsRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`,
      { headers: { Authorization: 'Bearer ' + token } }
    );
    const items = await itemsRes.json();
    const file = (items.value || []).find(f => f.name === 'Clientes Activos.xlsx');
    if (!file) {
      console.warn('Clientes Activos.xlsx no encontrado en SharePoint');
      return [];
    }

    let res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${file.id}/workbook/worksheets/Sheet1/usedRange`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) {
      res = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${file.id}/workbook/worksheets/Hoja1/usedRange`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!res.ok) {
        console.warn('No se pudo leer Clientes Activos.xlsx');
        return [];
      }
    }
    const data = await res.json();
    const values = data.values;
    if (!values || values.length < 2) return [];

    const headers = values[0].map(h => String(h).trim());
    const rows = values.slice(1)
      .filter(row => row.some(cell => cell !== '' && cell !== null))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i] ?? '');
        return obj;
      });

    // Mantener las llaves originales del xlsx (CardCode, Cliente, Asesor, Ciudad,
    // Estado, EstatusComercial) sin convertir a camelCase. Convención del proyecto:
    // STATE.* preserva la capitalización del xlsx/SAP — consistente con STATE.oportunidades
    // y simplifica migración futura a Service Layer (donde los campos son PascalCase).
    const clientes = rows
      .filter(r => r.CardCode && r.Cliente)
      .map(r => ({
        CardCode:        String(r.CardCode || '').trim(),
        Cliente:         String(r.Cliente || '').trim(),
        Asesor:          String(r.Asesor || '').trim(),
        Ciudad:          String(r.Ciudad || '').trim(),
        Estado:          String(r.Estado || '').trim(),
        EstatusComercial: String(r.EstatusComercial || '').trim()
      }));

    console.log(`Clientes activos cargados: ${clientes.length}`);
    return clientes;
  } catch(e) {
    console.warn('Error cargando clientes activos:', e.message);
    return [];
  }
}
