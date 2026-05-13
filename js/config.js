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
