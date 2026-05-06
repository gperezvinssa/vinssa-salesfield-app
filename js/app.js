// ── Utilidades ──────────────────────────────────────────────────────────────

function $(id) { return document.getElementById(id); }

function fechaHoy() {
  return new Date().toLocaleDateString('es-MX', {
    weekday: 'long', day: 'numeric', month: 'long',
    timeZone: CONFIG.timezone
  });
}

function iniciales(nombre) {
  return nombre.split(' ').slice(0,2).map(n => n[0]).join('').toUpperCase();
}

function mostrarScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  const s = $(id);
  if (s) { s.classList.add('active'); window.scrollTo(0,0); }
}

function irA(screenId, tipo) {
  STATE.tipo = tipo;
  STATE.division = CONFIG.usuario.division_default;
  STATE.marca = null;
  STATE.etapa = null;
  STATE.lideres = [];
  STATE.competidores = [];
  STATE.contactoNuevo = false;
  renderForm();
  mostrarScreen(screenId);
}

// ── Estado global ────────────────────────────────────────────────────────────

const STATE = {
  tipo: null,
  clienteExiste: true,
  division: 'trazabilidad',
  marca: null,
  etapa: null,
  lideres: [],
  competidores: [],
  contactoNuevo: false
};

// ── Colores por tipo ─────────────────────────────────────────────────────────

const TIPO_CONFIG = {
  visita:      { label: 'Nueva visita',           color: 'blue',   icon: 'V' },
  demo:        { label: 'Demo realizada',          color: 'green',  icon: 'D' },
  lead:        { label: 'Lead / prospecto',        color: 'amber',  icon: 'L' },
  oportunidad: { label: 'Actualizar oportunidad',  color: 'purple', icon: 'O' }
};

// ── Render formulario ────────────────────────────────────────────────────────

function renderForm() {
  const t = TIPO_CONFIG[STATE.tipo];
  const div = CONFIG.divisiones[STATE.division];

  // Marcas dinámicas por división
  let marcasHTML = '';
  Object.entries(div.marcas).forEach(([linea, marcas]) => {
    const pills = marcas.map(m => `
      <span class="pill ${STATE.marca === m ? 'sel-blue' : ''}"
            onclick="selMarca('${m}')">${m}</span>
    `).join('');
    marcasHTML += `
      <div class="field-label" style="margin-bottom:4px">${linea}</div>
      <div class="pill-row">${pills}</div>
    `;
  });

  // Etapas
  const etapasHTML = CONFIG.etapas.map(e => `
    <div class="stage-item ${STATE.etapa === e.id ? 'sel' : ''}"
         onclick="selEtapa('${e.id}')">${e.label}</div>
  `).join('');

  // Líderes
  const lideresHTML = CONFIG.lideres.map(l => {
    const sel = STATE.lideres.includes(l.email);
    return `
      <div class="lider-item ${sel ? 'sel' : ''}" onclick="toggleLider('${l.email}')">
        <div class="avatar" style="${sel ? 'background:#EAF3DE;color:#0F6E56' : ''}">${iniciales(l.nombre)}</div>
        <div>
          <div class="lider-name">${l.nombre}</div>
          <div class="lider-lineas">${l.lineas.join(' · ')}</div>
        </div>
        ${sel ? '<span style="margin-left:auto;font-size:12px;color:var(--green)">✓</span>' : ''}
      </div>
    `;
  }).join('');

  // Competidores
  const compHTML = CONFIG.competidores.map(c => `
    <span class="pill ${STATE.competidores.includes(c) ? 'sel-blue' : ''}"
          onclick="toggleComp('${c}')">${c}</span>
  `).join('');

  // Divisiones
  const divHTML = Object.entries(CONFIG.divisiones).map(([key, val]) => `
    <div class="type-btn ${STATE.division === key ? 'sel-blue' : ''}"
         onclick="selDivision('${key}')"
         style="font-size:11px">${val.label.split(' ')[0]}</div>
  `).join('');

  // Sección contacto nuevo (toggle)
  const contactoHTML = `
    <div class="card purple-card">
      <div class="card-title">¿Hay un contacto nuevo que agregar?</div>
      <div class="toggle-row">
        <div class="toggle ${STATE.contactoNuevo ? 'on' : ''}" onclick="toggleContacto()">
          <div class="toggle-dot"></div>
        </div>
        <span class="toggle-label">${STATE.contactoNuevo ? 'Sí, agregar contacto' : 'No por el momento'}</span>
      </div>
      ${STATE.contactoNuevo ? `
        <div class="field-label">Nombre completo <span class="req">*</span></div>
        <input type="text" id="c-nombre" placeholder="Nombre del contacto">
        <div class="field-row">
          <div>
            <div class="field-label">Puesto <span class="opt">opc</span></div>
            <input type="text" id="c-puesto" placeholder="Puesto">
          </div>
          <div>
            <div class="field-label">Teléfono <span class="opt">opc</span></div>
            <input type="tel" id="c-tel" placeholder="844 000 0000">
          </div>
        </div>
        <div class="field-label">Email <span class="opt">opcional</span></div>
        <input type="email" id="c-email" placeholder="correo@empresa.com">
      ` : ''}
    </div>
  `;

  // Sección líder (solo en demo)
  const liderSection = STATE.tipo === 'demo' ? `
    <div class="card">
      <div class="card-title">Líder de línea presente</div>
      <div class="search-box">
        <span class="search-icon">&#9906;</span>
        <input type="text" placeholder="Buscar líder..." oninput="filtrarLideres(this.value)" id="lider-search">
      </div>
      <div id="lideres-list">${lideresHTML}</div>
      ${STATE.lideres.length > 0 ? `
        <div class="nota-hint">Recibirán notificación para confirmar su asistencia</div>
      ` : ''}
    </div>
  ` : '';

  // Sección oportunidad (no en lead exploratorio)
  const opSection = STATE.tipo !== 'lead' ? `
    <div class="card">
      <div class="card-title">Etapa de la oportunidad <span class="req">*</span></div>
      <div class="stage-grid">${etapasHTML}</div>
      <div class="field-row">
        <div>
          <div class="field-label">Monto est. MXP <span class="req">*</span></div>
          <input type="number" id="monto" placeholder="0">
        </div>
        <div>
          <div class="field-label">Cierre estimado <span class="req">*</span></div>
          <input type="date" id="cierre">
        </div>
      </div>
      <div class="field-label">Presupuesto cliente <span class="opt">opcional</span></div>
      <input type="number" id="presupuesto" placeholder="¿El cliente mencionó presupuesto?">
      <div class="field-label">Nombre de la oportunidad <span class="req">*</span></div>
      <input type="text" id="opp-nombre" placeholder="Ej: PS/TMP1700/X1">
    </div>
  ` : '';

  $('screen-form').innerHTML = `
    <header class="top-bar ${t.color}">
      <button class="back-btn" onclick="mostrarScreen('screen-home')">← Inicio</button>
      <div style="color:white;font-size:14px;font-weight:500">${t.label}</div>
      <div style="width:50px"></div>
    </header>
    <div class="form-body">

      <div class="card">
        <div class="card-title">Cliente</div>
        <div style="display:flex;gap:6px;margin-bottom:8px">
          <div class="type-btn ${STATE.clienteExiste ? 'sel-blue' : ''}"
               style="flex:1" onclick="setClienteExiste(true)">Existente en SAP</div>
          <div class="type-btn ${!STATE.clienteExiste ? 'sel-amber' : ''}"
               style="flex:1" onclick="setClienteExiste(false)">Cliente nuevo</div>
        </div>
        <div class="field-label">Nombre del cliente <span class="req">*</span></div>
        <input type="text" id="cliente-nombre" placeholder="Buscar o escribir cliente...">
        <div class="field-label">Contacto principal <span class="opt">opcional</span></div>
        <input type="text" id="cliente-contacto" placeholder="Nombre del contacto">
      </div>

      ${contactoHTML}

      <div class="card">
        <div class="card-title">División <span class="req">*</span></div>
        <div class="type-grid" style="grid-template-columns:1fr 1fr 1fr;margin-bottom:10px">${divHTML}</div>
        <div class="card-title">Marca y producto <span class="req">*</span></div>
        ${marcasHTML}
        <div class="field-label">Producto / modelo <span class="req">*</span></div>
        <input type="text" id="producto" placeholder="Ej: LS-XPM 50W">
      </div>

      ${opSection}

      <div class="card">
        <div class="card-title">Competidor <span class="opt">opcional</span></div>
        <div class="pill-row">${compHTML}</div>
      </div>

      ${liderSection}

      <div class="card">
        <div class="card-title">Notas de la visita <span class="opt">opcional</span></div>
        <textarea id="notas" placeholder="Observaciones, siguientes pasos, acuerdos..."></textarea>
      </div>

      <button class="save-btn ${t.color}" onclick="guardar()">
        Guardar y sincronizar con SAP
      </button>

    </div>
  `;
}

// ── Selectores ───────────────────────────────────────────────────────────────

function selDivision(key) {
  STATE.division = key;
  STATE.marca = null;
  renderForm();
}

function selMarca(m) {
  STATE.marca = STATE.marca === m ? null : m;
  renderForm();
}

function selEtapa(id) {
  STATE.etapa = STATE.etapa === id ? null : id;
  renderForm();
}

function toggleLider(email) {
  const i = STATE.lideres.indexOf(email);
  if (i > -1) STATE.lideres.splice(i, 1);
  else STATE.lideres.push(email);
  renderForm();
}

function toggleComp(c) {
  const i = STATE.competidores.indexOf(c);
  if (i > -1) STATE.competidores.splice(i, 1);
  else STATE.competidores.push(c);
  renderForm();
}

function toggleContacto() {
  STATE.contactoNuevo = !STATE.contactoNuevo;
  renderForm();
}

function setClienteExiste(val) {
  STATE.clienteExiste = val;
  renderForm();
}

function filtrarLideres(q) {
  const lista = $('lideres-list');
  if (!lista) return;
  lista.querySelectorAll('.lider-item').forEach(item => {
    const nombre = item.querySelector('.lider-name').textContent.toLowerCase();
    item.style.display = nombre.includes(q.toLowerCase()) ? '' : 'none';
  });
}

// ── Guardar ──────────────────────────────────────────────────────────────────

function guardar() {
  const cliente = $('cliente-nombre')?.value?.trim();
  const producto = $('producto')?.value?.trim();

  if (!cliente) { alert('Falta el nombre del cliente'); return; }
  if (!STATE.marca) { alert('Selecciona una marca'); return; }
  if (!producto) { alert('Escribe el producto o modelo'); return; }
  if (STATE.tipo !== 'lead' && !STATE.etapa) { alert('Selecciona la etapa de la oportunidad'); return; }

  const registro = {
    tipo: STATE.tipo,
    fecha: new Date().toISOString(),
    cliente: cliente,
    clienteContacto: $('cliente-contacto')?.value?.trim(),
    clienteNuevo: !STATE.clienteExiste,
    division: STATE.division,
    marca: STATE.marca,
    producto: producto,
    etapa: STATE.etapa,
    monto: $('monto')?.value || null,
    cierre: $('cierre')?.value || null,
    presupuesto: $('presupuesto')?.value || null,
    oppNombre: $('opp-nombre')?.value?.trim() || null,
    competidores: STATE.competidores,
    lideres: STATE.lideres,
    notas: $('notas')?.value?.trim() || null,
    contactoNuevo: STATE.contactoNuevo ? {
      nombre:  $('c-nombre')?.value?.trim(),
      puesto:  $('c-puesto')?.value?.trim(),
      telefono:$('c-tel')?.value?.trim(),
      email:   $('c-email')?.value?.trim()
    } : null,
    asesor: CONFIG.usuario.email
  };

  console.log('Registro a enviar a SAP:', registro);
  alert(`✅ Registro guardado\n\n${registro.tipo.toUpperCase()} — ${registro.cliente}\n${registro.marca} · ${registro.producto}\n\nConexión SAP: próxima sesión`);
  mostrarScreen('screen-home');
}

// ── Init ─────────────────────────────────────────────────────────────────────

window.addEventListener('DOMContentLoaded', () => {
  const fechaEl = $('fecha-hoy');
  if (fechaEl) fechaEl.textContent = fechaHoy();
  const userEl = $('user-initials');
  if (userEl) userEl.textContent = iniciales(CONFIG.usuario.nombre);
  mostrarScreen('screen-home');
});
