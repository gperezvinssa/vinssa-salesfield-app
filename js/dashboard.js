// ── Normalizar nombre para comparación ───────────────────────────────────────
// Reemplaza la función existente en dashboard.js
// Convierte "Gerardo Pérez" → "GERARDO PEREZ" para comparar con SAP

function dashNormNombre(str) {
  return String(str || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // quita acentos
    .trim();
}

// ── Calcular métricas para un asesor/mes/año ──────────────────────────────────
// Reemplaza la función dashCalcMetricas existente en dashboard.js

function dashCalcMetricas(asesor, mes, anio) {
  const asesorNorm = dashNormNombre(asesor);

  const ventasFiltradas = DASH_STATE.ventas.filter(v => {
    if (!v.Fecha) return false;
    // Fecha puede venir como número serial de Excel o como string
    let fecha;
    if (typeof v.Fecha === 'number') {
      // Serial date de Excel: días desde 1/1/1900
      fecha = new Date((v.Fecha - 25569) * 86400 * 1000);
    } else {
      fecha = new Date(v.Fecha);
    }
    if (isNaN(fecha)) return false;
    const vMes  = fecha.getMonth() + 1;
    const vAnio = fecha.getFullYear();
    const vAsesor = dashNormNombre(v.Asesor);
    return vMes === mes && vAnio === anio && vAsesor === asesorNorm;
  });

  // Agrupar por división usando el mapa de grupos
  const porDivision = {};
  ventasFiltradas.forEach(v => {
    const grupo   = String(v.GrupoArticulo || '').trim();
    const division = DASHBOARD_CONFIG.mapaGrupos[grupo];
    if (!division) return;
    porDivision[division] = (porDivision[division] || 0) + parseFloat(v.Total || 0);
  });

  const totalVenta = ventasFiltradas.reduce((s, v) => s + parseFloat(v.Total || 0), 0);
  const numOVs    = new Set(ventasFiltradas.map(v => v.NumOV)).size;

  // Presupuesto del asesor para ese mes
  const presupFiltrado = DASH_STATE.presupuesto.filter(p => {
    const pAsesor = dashNormNombre(p.Asesor);
    const pMes   = parseInt(p.Mes);
    const pAnio  = parseInt(p['Año'] || p.Anio || p.Año || 0);
    return pAsesor === asesorNorm && pMes === mes && pAnio === anio;
  });

  const totalMeta = presupFiltrado.reduce((s, p) => s + parseFloat(p.Meta || 0), 0);

  const metaPorDivision = {};
  presupFiltrado.forEach(p => {
    const div = String(p.Division || p.División || '').trim();
    if (div) {
      metaPorDivision[div] = (metaPorDivision[div] || 0) + parseFloat(p.Meta || 0);
    }
  });

  const pct  = totalMeta > 0 ? Math.round(totalVenta / totalMeta * 100) : 0;
  const falta = Math.max(0, totalMeta - totalVenta);
  const diasRestantes = (mes === DASH_STATE.mesActual && anio === DASH_STATE.anioActual)
    ? diasEnMes(mes, anio) - new Date().getDate()
    : 0;
  const ritmo = diasRestantes > 0 ? falta / diasRestantes : 0;

  return { totalVenta, totalMeta, pct, falta, diasRestantes, ritmo, numOVs, porDivision, metaPorDivision };
}

// ── Calcular clientes en riesgo ───────────────────────────────────────────────
// Reemplaza la función dashClientesEnRiesgo existente en dashboard.js

function dashClientesEnRiesgo(asesor, diasUmbral = 60) {
  const asesorNorm = asesor ? dashNormNombre(asesor) : null;

  const ultimaCompra = {};
  DASH_STATE.ventas
    .filter(v => {
      if (!asesorNorm) return true;
      return dashNormNombre(v.Asesor) === asesorNorm;
    })
    .forEach(v => {
      const key = v.CardCode;
      let fecha;
      if (typeof v.Fecha === 'number') {
        fecha = new Date((v.Fecha - 25569) * 86400 * 1000);
      } else {
        fecha = new Date(v.Fecha);
      }
      if (isNaN(fecha)) return;
      if (!ultimaCompra[key] || fecha > ultimaCompra[key].fecha) {
        ultimaCompra[key] = {
          cardCode: v.CardCode,
          cliente:  String(v.Cliente || '').trim(),
          asesor:   String(v.Asesor  || '').trim(),
          fecha,
          total: parseFloat(v.Total || 0),
          grupo: String(v.GrupoArticulo || '').trim()
        };
      }
    });

  const hoy = new Date();
  return Object.values(ultimaCompra)
    .map(c => ({ ...c, dias: Math.floor((hoy - c.fecha) / (1000 * 60 * 60 * 24)) }))
    .filter(c => c.dias >= diasUmbral)
    .sort((a, b) => b.dias - a.dias);
}

// ── Obtener lista única de asesores con presupuesto ───────────────────────────
// Reemplaza la función dashGetAsesores existente en dashboard.js
// Solo devuelve asesores que tienen presupuesto definido — excluye mostrador, etc.

function dashGetAsesores() {
  const conPresupuesto = new Set(
    DASH_STATE.presupuesto
      .map(p => String(p.Asesor || '').trim())
      .filter(Boolean)
  );
  return [...conPresupuesto];
}

// ── Calcular histórico anual ──────────────────────────────────────────────────
// Reemplaza la función dashHistoricoAnual existente en dashboard.js

function dashHistoricoAnual(asesor, anio) {
  const meses = [];
  for (let m = 1; m <= 12; m++) {
    const met          = dashCalcMetricas(asesor, m, anio);
    const metAnterior  = dashCalcMetricas(asesor, m, anio - 1);
    meses.push({
      mes: m,
      label: DASHBOARD_CONFIG.meses[m - 1],
      venta:          met.totalVenta,
      meta:           met.totalMeta,
      ventaAnterior:  metAnterior.totalVenta,
      pct:            met.pct
    });
  }
  const acumulado         = meses.reduce((s, m) => s + m.venta, 0);
  const acumuladoAnterior = meses.reduce((s, m) => s + m.ventaAnterior, 0);
  const mejorMes          = meses.reduce((best, m) => m.venta > best.venta ? m : best, meses[0]);
  const metaAnual         = meses.reduce((s, m) => s + m.meta, 0);
  const mesesConDatos     = meses.filter(m => m.venta > 0);
  const proyeccion        = mesesConDatos.length > 0
    ? (acumulado / mesesConDatos.length) * 12
    : 0;
  return { meses, acumulado, acumuladoAnterior, mejorMes, metaAnual, proyeccion };
}
