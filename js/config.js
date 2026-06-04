const CONFIG = {

  usuario: {
    nombre: "Gerardo Pérez",
    iniciales: "GP",
    email: "gperez@vinssa.com",
    division_default: "trazabilidad"
  },

  divisiones: {
    trazabilidad: {
      label: "Trazabilidad y Automatización",
      marcas: {
        "Marcaje":      ["Telesis"],
        "Visión":       ["Cognex"],
        "Robótica":     ["UR", "MIR", "Nabtesco", "EasyRobotics"],
        "Herramientas": ["Atlas Copco"]
      }
    },
    suministros: {
      label: "Suministros Industriales",
      marcas: {
        "General": ["Por definir"]
      }
    },
    servicios: {
      label: "Servicios Industriales",
      marcas: {
        "General": ["Por definir"]
      }
    }
  },

  etapas: [
    { id: "contacto_inicial",    label: "Contacto Inicial",         pct: 10  },
    { id: "cotizacion",          label: "Cotización",               pct: 25  },
    { id: "pruebas_demo",        label: "Pruebas / Demostración",   pct: 60  },
    { id: "negociacion",         label: "Negociación",              pct: 80  },
    { id: "tramite_compras",     label: "Trámite con Compras",      pct: 90  },
    { id: "factura",             label: "Factura",                  pct: 95  }
  ],

  // Acompañantes que un asesor puede marcar en Visita/Demo: líderes de línea,
  // ingenieros de aplicación, y gerencia/dirección. Selección opcional, multi-pick.
  // El email es la identidad guardada en SharePoint (columna `Lideres` por internal
  // name aunque su display ahora es `Acompanante`). El rol se infiere al hacer
  // reporting cruzando el email contra este catálogo.
  acompanantes: [
    // Líderes de línea — uno por línea de Trazabilidad y Automatización.
    { nombre: "Antonio Martinez",   email: "amartinez@vinssa.com", rol: "Líder de línea",        linea: "Visión" },
    { nombre: "Jose Juan Aguillon", email: "jaguillon@vinssa.com", rol: "Líder de línea",        linea: "Trazabilidad" },
    { nombre: "Aldo Almaguer",      email: "aalmaguer@vinssa.com", rol: "Líder de línea",        linea: "Robótica" },
    { nombre: "Jonathan Roche",     email: "jroche@vinssa.com",    rol: "Líder de línea",        linea: "Herramienta de Atornillado" },
    // Ingenieros de aplicación — apoyan a un líder específico en demos/visitas.
    { nombre: "Luis Sanchez",       email: "lsanchez@vinssa.com",  rol: "Ingeniero de aplicación", linea: "Robótica", reportaA: "Aldo Almaguer" },
    { nombre: "Jesús Ortíz",        email: "jortiz@vinssa.com",    rol: "Ingeniero de aplicación", linea: "Visión",   reportaA: "Antonio Martinez" },
    // Gerencia y Dirección — scope por división o cross-división.
    { nombre: "Fernando Barajas",   email: "fbarajas@vinssa.com",  rol: "Gerente",          division: "Trazabilidad y Automatización" },
    { nombre: "Gerardo Pérez",      email: "gperez@vinssa.com",    rol: "Director Comercial" }
  ],

  competidores: [
    "Keyence", "SIC Marking", "Ninguno", "Otro"
  ],

  sap: {
    baseUrl: "https://vinssa.vertrou.cloud:50000/b1s/v1",
    company: "VINSSA"
  },

  timezone: "America/Monterrey"
};

// ── Mapeo MSAL email → nombre de asesor en SAP ──────────────────────────────
// El email del usuario logueado (account.username de MSAL, lowercase) se mapea
// al nombre de asesor SAP exacto (OSLP.SlpName, típicamente MAYÚSCULAS sin acentos).
// Este nombre se usa para filtrar Oportunidades.xlsx por la columna Asesor.
// Cuando un asesor nuevo se suma al piloto, agregar entrada aquí.
const EMAIL_A_ASESOR = {
  'kportillo@vinssa.com': 'KIMBERLY PORTILLO',
  'cangel@vinssa.com': 'CARLOS ANGEL',
  'rvillegas@vinssa.com': 'RAMON VILLEGAS',
  // OVERRIDE TEMPORAL: permite al Director (Gerardo) probar el flujo end-to-end
  // como un asesor sin colisionar con la cuenta real de Ramon (rvillegas), que ya
  // entró al piloto. Mapeado a RAFAEL MOYEDA — uno de los asesores piloto pendientes
  // (ver pending work en CLAUDE.md). Verificar que el SlpName en SAP sea exactamente
  // 'RAFAEL MOYEDA' (sin sufijo) la primera vez que se use; ajustar si SAP guarda
  // variante (ej. con apellido materno).
  // ELIMINAR esta entrada cuando Rafael Moyeda se sume con su propia cuenta MSAL.
  'gperez@vinssa.com': 'RAFAEL MOYEDA'
};

// ── Impersonación TEMPORAL (piloto) ──────────────────────────────────────────
// Permite a un usuario (típicamente gperez) ver el app como si fuera otro asesor.
// Solo afecta la lógica de rendering — el token MSAL y permisos Graph siguen
// siendo del usuario logueado. El email impersonado se usa para:
//   - resolver asesorSAP via EMAIL_A_ASESOR (auth.js)
//   - cargar rol via Lista Roles Dashboard.xlsx (cargarRolUsuario)
//   - resolver vista de dashboard (DASH_STATE.userEmail)
// Activar desde devtools: vinssaImpersonate('rvillegas@vinssa.com')
// Desactivar:             vinssaStopImpersonate()
// Eliminar este bloque cuando termine el piloto.
const IMPERSONATE_KEY = 'vinssa_impersonate_email';
function _getImpersonation() {
  try {
    const email = (localStorage.getItem(IMPERSONATE_KEY) || '').trim().toLowerCase();
    return { active: !!email, email };
  } catch(_) { return { active: false, email: '' }; }
}
function vinssaImpersonate(email) {
  if (!email || typeof email !== 'string') {
    console.warn('Uso: vinssaImpersonate("rvillegas@vinssa.com")');
    return;
  }
  localStorage.setItem(IMPERSONATE_KEY, email.toLowerCase().trim());
  console.log('Impersonando:', email, '— recargando...');
  location.reload();
}
function vinssaStopImpersonate() {
  localStorage.removeItem(IMPERSONATE_KEY);
  console.log('Impersonación desactivada — recargando...');
  location.reload();
}
window.vinssaImpersonate     = vinssaImpersonate;
window.vinssaStopImpersonate = vinssaStopImpersonate;

// Banner naranja fijo arriba cuando hay impersonación activa. Se inyecta en
// DOMContentLoaded y agrega padding-top al body para no tapar contenido.
function _renderImpersonationBanner() {
  const imp = _getImpersonation();
  if (!imp.active) return;
  if (document.getElementById('vinssa-imp-banner')) return;
  const banner = document.createElement('div');
  banner.id = 'vinssa-imp-banner';
  banner.style.cssText = 'position:fixed;top:0;left:0;right:0;background:#D85A30;color:white;font-size:12px;padding:8px 12px;z-index:99999;display:flex;justify-content:space-between;align-items:center;gap:8px;font-family:system-ui,sans-serif;box-shadow:0 1px 4px rgba(0,0,0,.25)';
  banner.innerHTML = '<span>\u{1F3AD} Impersonando: <strong>' + imp.email + '</strong></span>' +
    '<button onclick="vinssaStopImpersonate()" style="background:white;color:#D85A30;border:none;padding:4px 10px;border-radius:4px;font-size:11px;font-weight:600;cursor:pointer">Salir</button>';
  document.body.appendChild(banner);
  document.body.style.paddingTop = '36px';
}
window.addEventListener('DOMContentLoaded', _renderImpersonationBanner);
