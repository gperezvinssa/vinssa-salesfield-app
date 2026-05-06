const MSAL_CONFIG = {
  auth: {
    clientId: "45d6f369-f789-473b-8970-d9b25ff3225c",
    authority: "https://login.microsoftonline.com/7a272c1a-ee40-4b22-a187-2656ea44b4c4",
    redirectUri: "https://gperezvinssa.github.io/vinssa-salesfield-app/"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

const msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);

const LOGIN_SCOPES = { 
  scopes: [
    "User.Read",
    "https://versatilidadsaltillo.sharepoint.com/.default"
  ] 
};

async function iniciarApp(account) {
  const nombre = account.name || account.username;
  const email  = account.username;

  CONFIG.usuario.nombre    = nombre;
  CONFIG.usuario.email     = email;
  CONFIG.usuario.iniciales = nombre.split(' ').slice(0,2).map(n=>n[0]).join('').toUpperCase();

  document.getElementById('screen-login').style.display = 'none';
  document.getElementById('app').style.display = 'block';

  setTimeout(() => {
    const f = document.getElementById('fecha-hoy');
    if (f) f.textContent = new Date().toLocaleDateString('es-MX',{weekday:'long',day:'numeric',month:'long',timeZone:'America/Monterrey'});
    const u = document.getElementById('user-initials');
    if (u) u.textContent = CONFIG.usuario.iniciales;
    const n = document.getElementById('user-nombre');
    if (n) n.textContent = nombre.split(' ')[0];

    document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));
    const home = document.getElementById('screen-home');
    if (home) home.classList.add('active');
    window.scrollTo(0,0);
  }, 100);
}

async function loginMicrosoft() {
  try {
    const result = await msalInstance.loginPopup(LOGIN_SCOPES);
    await iniciarApp(result.account);
  } catch(e) {
    console.error('Login error:', e);
    alert('No se pudo iniciar sesión. Intenta de nuevo.');
  }
}

async function cerrarSesion() {
  if (!confirm('¿Cerrar sesión?')) return;
  await msalInstance.logoutPopup();
  location.reload();
}

window.addEventListener('DOMContentLoaded', async () => {
  await msalInstance.initialize();

  try {
    const result = await msalInstance.handleRedirectPromise();
    if (result) { await iniciarApp(result.account); return; }
  } catch(e) { console.error(e); }

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    await iniciarApp(accounts[0]);
  }
});
