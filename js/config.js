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

  lideres: [
    { nombre: "Antonio Martinez",  email: "amartinez@vinssa.com",  lineas: ["Marcaje", "Visión"] },
    { nombre: "Jose Juan Aguillon",email: "jaguillon@vinssa.com",  lineas: ["Robótica"] },
    { nombre: "Aldo Almaguer",     email: "aalmaguer@vinssa.com",  lineas: ["Herramientas"] },
    { nombre: "Jonathan Roche",    email: "jroche@vinssa.com",     lineas: ["Visión", "Marcaje"] }
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
  // OVERRIDE TEMPORAL: permite al Director (Gerardo) probar el flujo end-to-end
  // durante la fase de pulido del piloto sin requerir login con cuenta de asesor.
  // Mapeado a RAMON VILLEGAS porque tiene cartera real significativa (~76 clientes)
  // — permite probar el autocomplete y la priorización futura con datos densos.
  // ELIMINAR esta entrada antes de expandir el piloto a más asesores.
  'gperez@vinssa.com': 'RAMON VILLEGAS'
};
