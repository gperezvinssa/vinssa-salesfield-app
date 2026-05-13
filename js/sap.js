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
    Asesor: registro.asesor || '',
    Tipo: registro.tipo || '',
    ModoRegistro: registro.modo || 'campo',
    Marca: registro.marca || '',
    Producto: registro.producto || '',
    Etapa: registro.etapa || '',
    Monto: parseFloat(registro.monto) || 0,
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
