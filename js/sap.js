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

  const item = {
    Title: registro.cliente,
    Asesor: registro.asesor,
    Tipo: registro.tipo,
    Marca: registro.marca,
    Producto: registro.producto,
    Etapa: registro.etapa || '',
    Monto: parseFloat(registro.monto) || 0,
    Moneda: registro.moneda || 'MXP',
    Cierre: registro.cierre || null,
    Competidor: registro.competidores.join(', '),
    Lideres: registro.lideres.join(', '),
    Notas: registro.notas || '',
    GPS_Lat: registro.gps?.lat || null,
    GPS_Lng: registro.gps?.lng || null,
    GPS_Precision: registro.gps?.precision || null,
    Checkin_Hora: ultimoCheckin?.hora || '',
    Checkout_Hora: ultimoCheckout?.horaSalida || '',
    Duracion_Min: ultimoCheckout?.duracionMinutos || null,
    Contacto_Nuevo: registro.contactoNuevo ? 'Sí' : 'No',
    Contacto_Nombre: registro.contactoNuevo?.nombre || '',
    Contacto_Puesto: registro.contactoNuevo?.puesto || '',
    Contacto_Tel: registro.contactoNuevo?.telefono || '',
    Contacto_Email: registro.contactoNuevo?.email || '',
    SAP_OppID: sapOppId?.toString() || '',
    SAP_ActID: sapActId?.toString() || '',
    Estatus_SAP: sapOppId ? 'Sincronizado' : 'Pendiente'
  };

  try {
    // Obtener token de SharePoint via Microsoft Graph
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ['https://versatilidadsaltillo.sharepoint.com/.default'],
      account: msalInstance.getAllAccounts()[0]
    });

    const res = await fetch(
      'https://versatilidadsaltillo.sharepoint.com/sites/VINSSAAutomation/_api/web/lists/getbytitle(\'Visitas Field App\')/items',
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${tokenResponse.accessToken}`,
          'Content-Type': 'application/json;odata=verbose',
          'Accept': 'application/json;odata=verbose',
          'X-RequestDigest': await getRequestDigest(tokenResponse.accessToken)
        },
        body: JSON.stringify({ __metadata: { type: 'SP.Data.Visitas_x0020_Field_x0020_AppListItem' }, ...item })
      }
    );

    if (!res.ok) throw new Error('Error al guardar en SharePoint');
    const data = await res.json();
    console.log('Guardado en SharePoint:', data.d?.Id);
    return data.d?.Id;
  } catch(e) {
    console.warn('SharePoint error:', e.message);
    // Guardar localmente como respaldo
    const respaldo = JSON.parse(localStorage.getItem('vinssa_registros_pendientes') || '[]');
    respaldo.push({ item, timestamp: Date.now() });
    localStorage.setItem('vinssa_registros_pendientes', JSON.stringify(respaldo));
    return null;
  }
}

async function getRequestDigest(token) {
  const res = await fetch(
    'https://versatilidadsaltillo.sharepoint.com/sites/VINSSAAutomation/_api/contextinfo',
    {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json;odata=verbose'
      }
    }
  );
  const data = await res.json();
  return data.d?.GetContextWebInformation?.FormDigestValue || '';
}

// ── Flujo principal de sincronización ───────────────────────────────────────

async function sincronizarConSAP(registro) {
  const sesionActiva = await sapEnsureSession();

  if (!sesionActiva) {
    const usuario = prompt('Usuario SAP (tu correo de SAP):');
    if (!usuario) return null;
    const password = prompt('Contraseña SAP:');
    if (!password) return null;

    const ok = await sapLogin(usuario, password);
    if (!ok) {
      alert('No se pudo conectar a SAP. Verifica tus credenciales.');
      return null;
    }
  }

  let sapOppId = null;
  let sapActId = null;
  const errores = [];

  try {
    if (registro.tipo !== 'lead') {
      const opp = await sapCrearOportunidad(registro);
      sapOppId = opp?.SequenceNo || opp?.OpportunityId || null;
    }
  } catch(e) {
    errores.push(`Oportunidad: ${e.message}`);
    console.error(e);
  }

  try {
    const act = await sapCrearActividad(registro, sapOppId);
    sapActId = act?.ActivityCode || null;
  } catch(e) {
    errores.push(`Actividad: ${e.message}`);
    console.error(e);
  }

  try {
    if (registro.contactoNuevo?.nombre) {
      await sapCrearContacto(registro);
    }
  } catch(e) {
    errores.push(`Contacto: ${e.message}`);
    console.error(e);
  }

  await guardarEnSharePoint(registro, sapOppId, sapActId);

  return { sapOppId, sapActId, errores };
}
