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

    // Clientes Activos: se carga para TODOS los usuarios (independiente de asesorSAP)
    // porque el autocomplete de check-in debe servir incluso a usuarios sin mapeo a
    // asesor SAP — caen al grupo 4 (alfabético) en la priorización.
    STATE.clientesActivos = [];
    STATE.clientesLoading = true;
    cargarClientesActivos().then(clientes => {
      STATE.clientesActivos = clientes;
      STATE.clientesLoading = false;
      if (typeof onClientesActivosCargados === 'function') onClientesActivosCargados();
    });

    // Rol del usuario: decide filtrado de clientes en check-in. Asesor ve solo
    // sus clientes asignados; gerente/líder/director ven todos (decisión 2026-05-20:
    // clientes son multi-división, ver CLAUDE.md). Default 'asesor' si email no
    // aparece en Lista Roles Dashboard.xlsx.
    STATE.rolUsuario = 'asesor';
    STATE.divisionUsuario = 'Todos';
    STATE.rolLoading = true;
    cargarRolUsuario(email).then(({ rol, division }) => {
      STATE.rolUsuario = rol;
      STATE.divisionUsuario = division;
      STATE.rolLoading = false;
      if (typeof onClientesActivosCargados === 'function') onClientesActivosCargados();
    });
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

    // Re-render home con CONFIG.usuario actualizado. app.js corre renderHome() en
    // DOMContentLoaded antes de que MSAL resuelva la sesión, así que el primer
    // render queda con el default "Gerardo Pérez" de config.js. Sin este re-render,
    // el saludo del home queda congelado con el default aunque el avatar (que sí
    // se actualiza arriba directo al DOM) ya muestre las iniciales correctas.
    if (typeof renderHome === 'function') renderHome();

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

// ── Token silencioso con fallback (popup desktop / redirect móvil) ───────────
// Wrapper de msalInstance.acquireTokenSilent que cae a interacción cuando MSAL no
// puede renovar el token en silencio. Casos cubiertos:
// - InteractionRequiredAuthError: token expirado, MFA requerida, consent_required.
// - BrowserAuthError con errorCode 'monitor_window_timeout': el iframe oculto que
//   MSAL usa para renovar excedió su timeout (típico tras horas de inactividad).
//
// Branch por dispositivo:
// - Desktop: acquireTokenPopup (UX 1-tap, mantiene estado de la app).
// - Móvil (viewport <768px o PWA standalone): acquireTokenRedirect. iOS Safari y
//   Chrome Android bloquean popups; aunque el usuario los apruebe, la PWA pierde
//   estado. Redirect navega a Microsoft y regresa; handleRedirectPromise refresca
//   la sesión. Tradeoff: el caller (ej. dashInit) NO se resume — el usuario debe
//   re-disparar la acción tras el regreso (2 taps en lugar de 1).
//
// Si la interacción falla, re-lanzamos el error ORIGINAL para que el caller maneje UX.
function _esMovil() {
  return window.matchMedia('(max-width: 768px)').matches
      || window.matchMedia('(display-mode: standalone)').matches;
}

async function acquireTokenSafe(request) {
  try {
    return await msalInstance.acquireTokenSilent(request);
  } catch(e) {
    const esInteraccionRequerida =
      (typeof msal !== 'undefined' && msal.InteractionRequiredAuthError && e instanceof msal.InteractionRequiredAuthError) ||
      (e && (e.errorCode === 'monitor_window_timeout'
          || e.errorCode === 'consent_required'
          || e.errorCode === 'login_required'
          || e.errorCode === 'interaction_required'));
    if (!esInteraccionRequerida) throw e;

    if (_esMovil()) {
      console.warn('acquireTokenSilent falló en móvil (' + (e.errorCode || e.message) + '), usando redirect');
      try {
        await msalInstance.acquireTokenRedirect(request);
      } catch(redirectErr) {
        console.error('acquireTokenRedirect también falló:', redirectErr);
        throw e; // re-lanza el original
      }
      // La página se redirige antes de que JS continúe; esta promise no resuelve
      // en el contexto actual. Devolvemos pending-forever para que el caller no
      // interprete falsamente "resolvió con undefined".
      return new Promise(() => {});
    }

    console.warn('acquireTokenSilent falló (' + (e.errorCode || e.message) + '), intentando popup');
    try {
      return await msalInstance.acquireTokenPopup(request);
    } catch(popupErr) {
      console.error('acquireTokenPopup también falló:', popupErr);
      throw e; // re-lanza el error original — el caller maneja la UX
    }
  }
}

// ── Exponer funciones globales ───────────────────────────────────────────────

window.loginMicrosoft   = loginMicrosoft;
window.cerrarSesion     = cerrarSesion;
window.acquireTokenSafe = acquireTokenSafe;
