// ── Configuración MSAL ───────────────────────────────────────────────────────

const MSAL_CONFIG = {
  auth: {
    clientId: "45d6f369-f789-473b-8970-d9b25ff3225c",
    authority: "https://login.microsoftonline.com/7a272c1a-ee40-4b22-a187-2656ea44b4c4",
    redirectUri: "https://gperezvinssa.github.io/vinssa-salesfield-app/"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true  // ← true para compatibilidad Safari iOS
  }
};

const msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);

const LOGIN_SCOPES = {
  scopes: ["User.Read", "Sites.ReadWrite.All"]
};

// ── Iniciar app tras login exitoso ───────────────────────────────────────────

async function iniciarApp(account) {
  const nombre = account.name || account.username;
  const email  = account.username;

  CONFIG.usuario.nombre    = nombre;
  CONFIG.usuario.email     = email;
  CONFIG.usuario.iniciales = nombre.split(' ').slice(0,2).map(n=>n[0]).join('').toUpperCase();

  // Resolver asesor SAP del usuario logueado para piloto de Actualizar Oportunidad.
  // Si no está en el mapeo, queda null y los 3 sub-flujos muestran mensaje de piloto.
  if (typeof STATE !== 'undefined') {
    const emailKey = String(email || '').toLowerCase();
    STATE.asesorSAP = (typeof EMAIL_A_ASESOR !== 'undefined' && EMAIL_A_ASESOR[emailKey]) || null;
    STATE.oportunidades = [];
    // Carga async — no bloquea la entrada a Home. Cuando termina,
    // si el form Avanzó/Ganada/Perdida ya está abierto, app.js refresca el dropdown.
    if (STATE.asesorSAP) {
      STATE.opsLoading = true;
      cargarOportunidadesAsesor().then(ops => {
        STATE.oportunidades = ops;
        STATE.opsLoading = false;
        if (typeof onOportunidadesCargadas === 'function') onOportunidadesCargadas();
      });
    } else {
      // Usuario sin asesor → no hay nada que cargar, no mostrar "Cargando..." indefinido.
      STATE.opsLoading = false;
    }
  }

  document.getElementById('screen-login').style.display = 'none';
  document.getElementById('app').style.display = 'block';

  setTimeout(() => {
    const f = document.getElementById('fecha-hoy');
    if (f) f.textContent = new Date().toLocaleDateString('es-MX', {
      weekday: 'long', day: 'numeric', month: 'long', timeZone: 'America/Monterrey'
    });
    const u = document.getElementById('user-initials');
    if (u) u.textContent = CONFIG.usuario.iniciales;
    const n = document.getElementById('user-nombre');
    if (n) n.textContent = nombre.split(' ')[0];

    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    const home = document.getElementById('screen-home');
    if (home) home.classList.add('active');
    window.scrollTo(0, 0);
  }, 100);
}

// ── Login con redirect — funciona en Safari iOS, Chrome Android, Edge ────────

async function loginMicrosoft() {
  try {
    await msalInstance.loginRedirect(LOGIN_SCOPES);
    // La página se redirige a Microsoft y regresa — el resultado
    // se maneja en handleRedirectPromise dentro del DOMContentLoaded
  } catch(e) {
    console.error('Login error:', e);
    alert('No se pudo iniciar sesión. Intenta de nuevo.');
  }
}

async function cerrarSesion() {
  if (!confirm('¿Cerrar sesión?')) return;
  await msalInstance.logoutRedirect();
}

// ── Init — manejar redirect de regreso desde Microsoft ───────────────────────

window.addEventListener('DOMContentLoaded', async () => {
  await msalInstance.initialize();

  try {
    // Captura el resultado cuando Microsoft redirige de regreso a la app
    const result = await msalInstance.handleRedirectPromise();
    if (result) {
      await iniciarApp(result.account);
      return;
    }
  } catch(e) {
    console.error('Redirect error:', e);
  }

  // Si ya hay sesión guardada, entrar directo sin pedir login
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    await iniciarApp(accounts[0]);
  }
});

// ── Exponer funciones globales ───────────────────────────────────────────────

window.loginMicrosoft = loginMicrosoft;
window.cerrarSesion   = cerrarSesion;
