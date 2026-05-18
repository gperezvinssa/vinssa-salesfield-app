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
  // OVERRIDE TEMPORAL: permite al Director (Gerardo) probar el flujo end-to-end
  // durante la fase de pulido del piloto sin requerir login con cuenta de asesor.
  // Mapeado a RAMON VILLEGAS porque tiene cartera real significativa (~76 clientes)
  // — permite probar el autocomplete y la priorización futura con datos densos.
  // ELIMINAR esta entrada antes de expandir el piloto a más asesores.
  'gperez@vinssa.com': 'RAMON VILLEGAS'
};
