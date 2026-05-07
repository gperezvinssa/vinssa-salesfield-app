// ── Geolocalización — 3 capas ────────────────────────────────────────────────

const GEO = {
  checkin: null,
  watching: false
};

// ── Utilidades ───────────────────────────────────────────────────────────────

function obtenerPosicion() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error('GPS no disponible en este dispositivo'));
      return;
    }
    navigator.geolocation.getCurrentPosition(
      pos => resolve({
        lat: pos.coords.latitude,
        lng: pos.coords.longitude,
        precision: Math.round(pos.coords.accuracy),
        timestamp: new Date().toISOString()
      }),
      err => reject(err),
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
  });
}

function distanciaMetros(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) *
            Math.sin(dLng/2) * Math.sin(dLng/2);
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)));
}

function guardarRegistroGeo(registro) {
  const key = 'vinssa_geo_registros';
  const existentes = JSON.parse(localStorage.getItem(key) || '[]');
  existentes.push(registro);
  localStorage.setItem(key, JSON.stringify(existentes));
}

function obtenerRegistrosGeo() {
  return JSON.parse(localStorage.getItem('vinssa_geo_registros') || '[]');
}

// ── CAPA 1 — Captura GPS al guardar registro ─────────────────────────────────

async function capturarGPSAlGuardar(datosRegistro) {
  try {
    const pos = await obtenerPosicion();
    datosRegistro.gps = pos;

    if (datosRegistro.clienteCoords) {
      const distancia = distanciaMetros(
        pos.lat, pos.lng,
        datosRegistro.clienteCoords.lat,
        datosRegistro.clienteCoords.lng
      );
      datosRegistro.distanciaCliente = distancia;
      if (distancia > 500) {
        const continuar = confirm(
          `⚠️ Parece que estás a ${distancia}m del cliente registrado.\n\n¿Deseas guardar el registro de todas formas?`
        );
        if (!continuar) return null;
        datosRegistro.gpsAlerta = true;
      }
    }

    guardarRegistroGeo({
      tipo: 'registro',
      tipoRegistro: datosRegistro.tipo,
      cliente: datosRegistro.cliente,
      asesor: datosRegistro.asesor,
      gps: pos,
      fecha: new Date().toISOString()
    });

    return datosRegistro;
  } catch(e) {
    console.warn('GPS no disponible:', e.message);
    datosRegistro.gps = null;
    return datosRegistro;
  }
}

// ── CAPA 2 — Check-in / Check-out ───────────────────────────────────────────

async function checkin(clienteNombre) {
  try {
    const pos = await obtenerPosicion();
    GEO.checkin = {
      cliente: clienteNombre,
      gps: pos,
      horaEntrada: new Date().toISOString()
    };
    localStorage.setItem('vinssa_checkin_activo', JSON.stringify(GEO.checkin));

    guardarRegistroGeo({
      tipo: 'checkin',
      cliente: clienteNombre,
      asesor: CONFIG.usuario.email,
      gps: pos,
      hora: GEO.checkin.horaEntrada
    });

    actualizarBotonCheckin(true, clienteNombre);
    return GEO.checkin;
  } catch(e) {
    alert('No se pudo obtener tu ubicación. Verifica que el GPS esté activado.');
    return null;
  }
}

async function checkout() {
  const checkinActivo = GEO.checkin ||
    JSON.parse(localStorage.getItem('vinssa_checkin_activo') || 'null');

  if (!checkinActivo) {
    alert('No hay ninguna visita activa.');
    return null;
  }

  const entrada  = new Date(checkinActivo.horaEntrada);
  const salida   = new Date();
  const minutos  = Math.round((salida - entrada) / 60000);

  let gpsSalida = null;
  try {
    gpsSalida = await obtenerPosicion();
  } catch(e) {
    console.warn('GPS no disponible al hacer checkout:', e.message);
  }

  const registro = {
    tipo: 'checkout',
    cliente: checkinActivo.cliente,
    asesor: CONFIG.usuario.email,
    gpsEntrada: checkinActivo.gps,
    gpsSalida: gpsSalida,
    horaEntrada: checkinActivo.horaEntrada,
    horaSalida: salida.toISOString(),
    duracionMinutos: minutos
  };

  guardarRegistroGeo(registro);
  localStorage.removeItem('vinssa_checkin_activo');
  GEO.checkin = null;
  actualizarBotonCheckin(false, null);
  alert(`Visita terminada\n\nCliente: ${registro.cliente}\nDuración: ${minutos} minutos`);
  return registro;
}

function actualizarBotonCheckin(activo, cliente) {
  const btn  = document.getElementById('btn-checkin');
  const btns = document.querySelectorAll('.registro-btn');
  const topBar = document.querySelector('.top-bar');

  if (activo) {
    if (btn) {
      btn.className = 'action-card checkin-active';
      btn.onclick = () => checkout();
      btn.innerHTML = `
        <div class="action-icon icon-green" style="position:relative;background:#0F6E56;color:white">
          <span style="width:8px;height:8px;background:#22c55e;border-radius:50%;position:absolute;top:4px;right:4px;border:1.5px solid white"></span>
          C
        </div>
        <div>
          <div class="action-title">Terminar visita</div>
          <div class="action-sub">En: ${cliente}</div>
        </div>`;
    }
    btns.forEach(b => {
      b.disabled = false;
      const title = b.querySelector('.action-title')?.textContent;
      b.querySelector('.action-sub').textContent =
        title === 'Nueva visita'          ? 'Registrar cliente visitado' :
        title === 'Demo realizada'        ? 'Con o sin líder de línea'   :
        title === 'Lead / prospecto'      ? 'Evaluación inicial'         :
                                            'Cambiar etapa o datos';
    });
    if (topBar) topBar.style.background = '#0F6E56';
  } else {
    if (btn) {
      btn.className = 'action-card';
      btn.onclick = () => mostrarModalCheckin();
      btn.innerHTML = `
        <div class="action-icon icon-blue">C</div>
        <div>
          <div class="action-title">¿Estás con un cliente?</div>
          <div class="action-sub">Registra tu llegada primero</div>
        </div>`;
    }
    btns.forEach(b => {
      b.disabled = true;
      b.querySelector('.action-sub').textContent = 'Requiere check-in activo';
    });
    if (topBar) topBar.style.background = '#111827';
  }
}

// ── CAPA 3 — Exportar datos para Power BI ───────────────────────────────────

function exportarDatosGeo() {
  const registros = obtenerRegistrosGeo();
  if (registros.length === 0) {
    alert('No hay registros de GPS todavía.');
    return;
  }

  const csv = [
    'Tipo,Cliente,Asesor,Latitud,Longitud,Precision,Fecha,Hora,DuracionMin',
    ...registros.map(r => [
      r.tipo,
      `"${r.cliente || ''}"`,
      r.asesor || '',
      r.gps?.lat || r.gpsEntrada?.lat || '',
      r.gps?.lng || r.gpsEntrada?.lng || '',
      r.gps?.precision || '',
      r.fecha ? r.fecha.split('T')[0] : (r.hora ? r.hora.split('T')[0] : ''),
      r.fecha ? r.fecha.split('T')[1]?.slice(0,8) : (r.hora ? r.hora.split('T')[1]?.slice(0,8) : ''),
      r.duracionMinutos || ''
    ].join(','))
  ].join('\n');

  const blob = new Blob([csv], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `vinssa_geo_${new Date().toISOString().split('T')[0]}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── Init — restaurar checkin activo si existe ────────────────────────────────

window.addEventListener('DOMContentLoaded', () => {
  const checkinActivo = localStorage.getItem('vinssa_checkin_activo');
  if (checkinActivo) {
    const data = JSON.parse(checkinActivo);
    GEO.checkin = data;
    setTimeout(() => actualizarBotonCheckin(true, data.cliente), 200);
  }
});

// ── Exponer funciones globales ───────────────────────────────────────────────

window.checkin              = checkin;
window.checkout             = checkout;
window.exportarDatosGeo     = exportarDatosGeo;
window.capturarGPSAlGuardar = capturarGPSAlGuardar;
