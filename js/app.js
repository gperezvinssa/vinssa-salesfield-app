function $(id){return document.getElementById(id)}
function fechaHoy(){return new Date().toLocaleDateString('es-MX',{weekday:'long',day:'numeric',month:'long',timeZone:CONFIG.timezone})}
function iniciales(n){return n.split(' ').slice(0,2).map(x=>x[0]).join('').toUpperCase()}
function mostrarScreen(id){document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));const s=$(id);if(s){s.classList.add('active');window.scrollTo(0,0)}}

const STATE={
  modo:null,                           // null=selector | 'campo' | 'gabinete'
  tipo:null,division:'trazabilidad',marca:null,etapa:null,
  resultado:null,                      // null=sub-selector | 'avance' | 'ganada' | 'perdida'
  razonPerdida:null,                   // string elegido cuando resultado='perdida'
  lideres:[],competidores:[],contactoNuevo:false,moneda:'MXP',
  clienteExiste:true,
  campos:{}
};

function guardarCampos(){
  ['cliente-nombre','cliente-contacto','producto','monto','cierre','presupuesto','opp-nombre','notas','c-nombre','c-puesto','c-tel','c-email','monto-final','fecha-cierre-real','razon-perdida-detalle'].forEach(id=>{
    const el=$(id);if(el)STATE.campos[id]=el.value;
  });
}

function restaurarCampos(){
  Object.entries(STATE.campos).forEach(([id,val])=>{
    const el=$(id);if(el)el.value=val;
  });
}

function irA(screenId,tipo){
  STATE.tipo=tipo;STATE.division=CONFIG.usuario.division_default;
  STATE.marca=null;STATE.etapa=null;STATE.lideres=[];
  STATE.competidores=[];STATE.contactoNuevo=false;STATE.moneda='MXP';STATE.campos={};
  STATE.resultado=null;STATE.razonPerdida=null;
  renderForm();mostrarScreen(screenId);
}

// ── Render Home según STATE.modo ─────────────────────────────────────────────
function renderHome() {
  const body = $('home-body');
  if (!body) return;
  const nombre = CONFIG.usuario.nombre;

  // Reset top-bar a default; la rama campo lo pinta verde si hay visita activa
  const topBarReset = document.querySelector('.top-bar');
  if (topBarReset) topBarReset.style.background = '';

  if (STATE.modo === null) {
    body.innerHTML = `
      <div style="font-size:14px;color:var(--color-text-primary);margin-bottom:8px">Hola, <span style="font-weight:500">${nombre}</span></div>
      <div class="section-label">¿Cómo vas a registrar?</div>
      <button class="action-card mode-card" onclick="seleccionarModo('campo')">
        <div class="action-icon icon-blue">V</div>
        <div>
          <div class="action-title">Estoy con un cliente</div>
          <div class="action-sub">Visita o demo en persona · requiere check-in</div>
        </div>
      </button>
      <button class="action-card mode-card" onclick="seleccionarModo('gabinete')">
        <div class="action-icon icon-gray">O</div>
        <div>
          <div class="action-title">Trabajo de seguimiento</div>
          <div class="action-sub">Actualizar oportunidad desde oficina</div>
        </div>
      </button>`;
  } else if (STATE.modo === 'campo') {
    const enVisita = typeof GEO !== 'undefined' && GEO.checkin;
    if (!enVisita) {
      body.innerHTML = `
        <div style="font-size:14px;color:var(--color-text-primary);margin-bottom:16px">Hola, <span style="font-weight:500">${nombre}</span></div>
        <div class="info-banner">Iniciando visita...</div>`;
    } else {
      body.innerHTML = `
        <div style="font-size:14px;color:var(--color-text-primary);margin-bottom:16px">Hola, <span style="font-weight:500">${nombre}</span></div>
        <button class="action-card checkin-active" id="btn-checkin" onclick="checkout()">
          <div class="action-icon icon-green" style="position:relative;background:#0F6E56;color:white">
            <span style="width:8px;height:8px;background:#22c55e;border-radius:50%;position:absolute;top:4px;right:4px;border:1.5px solid white"></span>
            C
          </div>
          <div>
            <div class="action-title">Terminar visita</div>
            <div class="action-sub">En: ${GEO.checkin.cliente}</div>
          </div>
        </button>
        <div style="height:0.5px;background:var(--color-border-tertiary);margin:10px 0"></div>
        <div class="section-label">¿Qué vas a registrar?</div>
        <div id="registro-btns">
          <button class="action-card registro-btn" onclick="irA('screen-form','visita')">
            <div class="action-icon icon-blue">V</div>
            <div><div class="action-title">Nueva visita</div><div class="action-sub">Registrar cliente visitado</div></div>
          </button>
          <button class="action-card registro-btn" onclick="irA('screen-form','demo')">
            <div class="action-icon icon-green">D</div>
            <div><div class="action-title">Demo realizada</div><div class="action-sub">Con o sin líder de línea</div></div>
          </button>
          <button class="action-card registro-btn" onclick="irA('screen-form','oportunidad')">
            <div class="action-icon icon-gray">O</div>
            <div><div class="action-title">Actualizar oportunidad</div><div class="action-sub">Cambiar etapa o datos</div></div>
          </button>
        </div>`;
    }
    const topBar = document.querySelector('.top-bar');
    if (topBar) topBar.style.background = enVisita ? '#0F6E56' : '';
  } else if (STATE.modo === 'gabinete') {
    body.innerHTML = `
      <div style="font-size:14px;color:var(--color-text-primary);margin-bottom:12px">Hola, <span style="font-weight:500">${nombre}</span></div>
      <div class="info-banner">Modo seguimiento</div>
      <div class="section-label">¿Qué vas a registrar?</div>
      <button class="action-card" onclick="irA('screen-form','oportunidad')">
        <div class="action-icon icon-gray">O</div>
        <div><div class="action-title">Actualizar oportunidad</div><div class="action-sub">Desde oficina</div></div>
      </button>`;
  }

  const chip = $('btn-cambiar-modo');
  if (chip) chip.style.display = STATE.modo ? '' : 'none';
}

async function seleccionarModo(modo) {
  if (modo === 'gabinete') {
    STATE.modo = 'gabinete';
    renderHome();
    return;
  }
  if (modo === 'campo') {
    // Si ya hay check-in activo (reload con visita viva), entrar directo a campo
    if (typeof GEO !== 'undefined' && GEO.checkin) {
      STATE.modo = 'campo';
      renderHome();
      return;
    }
    // Sin check-in: lanzar prompt directamente, sin pantalla intermedia
    const nombre = prompt('¿Con qué cliente estás?');
    if (!nombre || !nombre.trim()) return; // cancelado: queda en selector
    STATE.modo = 'campo';
    renderHome(); // estado de carga mientras GPS captura
    const result = await checkin(nombre.trim());
    if (!result) {
      // GPS denegado o falló — regresar al selector
      STATE.modo = null;
      renderHome();
      return;
    }
    renderHome(); // re-render con GEO.checkin presente → estado activo
  }
}

async function cambiarModo() {
  if (STATE.modo === 'campo' && typeof GEO !== 'undefined' && GEO.checkin) {
    const ok = confirm(`Tienes una visita activa con ${GEO.checkin.cliente}. ¿Terminar visita y regresar?`);
    if (!ok) return;
    await checkout();
  }
  STATE.modo = null;
  renderHome();
}

const TIPO_CONFIG={
  visita:   {label:'Nueva visita',          color:'blue',  icon:'V'},
  demo:     {label:'Demo realizada',         color:'green', icon:'D'},
  oportunidad:{label:'Actualizar oportunidad',color:'purple',icon:'O'}
};

const RAZONES_PERDIDA = [
  'Precio',
  'Tiempo de entrega',
  'Especificaciones técnicas',
  'Sin presupuesto del cliente',
  'Decisión interna del cliente',
  'Competencia con relación previa',
  'Sin respuesta del cliente',
  'Otro'
];

function renderForm(){
  // CASO 1: Sub-selector de tipo de actualización (oportunidad sin resultado elegido)
  if (STATE.tipo === 'oportunidad' && STATE.resultado === null) {
    $('screen-form').innerHTML = `
      <header class="top-bar purple">
        <button class="back-btn" onclick="mostrarScreen('screen-home')">← Inicio</button>
        <div style="color:white;font-size:14px;font-weight:500">Actualizar oportunidad</div>
        <div style="width:50px"></div>
      </header>
      <div class="form-body">
        <div class="section-label" style="margin-top:4px;margin-bottom:12px">¿Qué pasó con la oportunidad?</div>
        <button class="action-card mode-card" onclick="selResultadoOpp('avance')">
          <div class="action-icon icon-blue">📈</div>
          <div>
            <div class="action-title">Avanzó de etapa</div>
            <div class="action-sub">Sigue activa, solo cambia el momento del pipeline.</div>
          </div>
        </button>
        <button class="action-card mode-card" onclick="selResultadoOpp('ganada')">
          <div class="action-icon icon-green">✅</div>
          <div>
            <div class="action-title">Se ganó</div>
            <div class="action-sub">Cliente confirmó la compra.</div>
          </div>
        </button>
        <button class="action-card mode-card" onclick="selResultadoOpp('perdida')">
          <div class="action-icon icon-red">❌</div>
          <div>
            <div class="action-title">Se perdió</div>
            <div class="action-sub">La oportunidad ya no procederá.</div>
          </div>
        </button>
      </div>`;
    return;
  }

  // CASO 2: Form Ganada
  if (STATE.tipo === 'oportunidad' && STATE.resultado === 'ganada') {
    const hoyISO = new Date().toISOString().split('T')[0];
    $('screen-form').innerHTML = `
      <header class="top-bar green">
        <button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>
        <div style="color:white;font-size:14px;font-weight:500">Oportunidad ganada</div>
        <div style="width:50px"></div>
      </header>
      <div class="form-body">
        <div class="card">
          <div class="card-title">Cliente</div>
          <div style="display:flex;gap:6px;margin-bottom:8px">
            <div class="type-btn ${STATE.clienteExiste?'sel-blue':''}" style="flex:1" onclick="setClienteExiste(true)">Existente en SAP</div>
            <div class="type-btn ${!STATE.clienteExiste?'sel-amber':''}" style="flex:1" onclick="setClienteExiste(false)">Cliente nuevo</div>
          </div>
          <div class="field-label">Nombre del cliente <span class="req">*</span></div>
          <input type="text" id="cliente-nombre" placeholder="Buscar o escribir cliente...">
          <div class="field-label">Nombre de la oportunidad <span class="opt">opc</span></div>
          <input type="text" id="opp-nombre" placeholder="Si lo recuerdas, ayuda a identificar la oportunidad">
        </div>
        <div class="card">
          <div class="card-title">Cierre ganado</div>
          <div class="field-row" style="align-items:end">
            <div>
              <div class="field-label">Monto final <span class="req">*</span></div>
              <div style="display:flex;gap:6px">
                <select id="moneda" onchange="selMoneda(this.value)" style="width:80px;margin-bottom:8px">
                  <option value="MXP" ${STATE.moneda==='MXP'?'selected':''}>MXP</option>
                  <option value="USD" ${STATE.moneda==='USD'?'selected':''}>USD</option>
                </select>
                <input type="number" id="monto-final" placeholder="0" style="flex:1">
              </div>
            </div>
            <div>
              <div class="field-label">Fecha de cierre real <span class="req">*</span></div>
              <input type="date" id="fecha-cierre-real" value="${STATE.campos['fecha-cierre-real'] || hoyISO}">
            </div>
          </div>
        </div>
        <div class="card">
          <div class="card-title">Notas de cierre <span class="opt">opc</span></div>
          <textarea id="notas" placeholder="Detalles del cierre, condiciones acordadas..."></textarea>
        </div>
        <button class="save-btn green" onclick="guardar()">Guardar cierre ganado</button>
      </div>`;
    restaurarCampos();
    return;
  }

  // CASO 3: Form Perdida
  if (STATE.tipo === 'oportunidad' && STATE.resultado === 'perdida') {
    const razonesHTML = RAZONES_PERDIDA.map(r =>
      `<div class="stage-item ${STATE.razonPerdida===r?'sel':''}" onclick="selRazonPerdida('${r.replace(/'/g,"\\'")}')">${r}</div>`
    ).join('');
    const detalleHTML = STATE.razonPerdida === 'Otro' ? `
      <div class="field-label" style="margin-top:10px">Detalle <span class="opt">opc</span></div>
      <textarea id="razon-perdida-detalle" placeholder="Especifica la razón (recomendado para análisis posterior)"></textarea>
    ` : '';
    $('screen-form').innerHTML = `
      <header class="top-bar red">
        <button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>
        <div style="color:white;font-size:14px;font-weight:500">Oportunidad perdida</div>
        <div style="width:50px"></div>
      </header>
      <div class="form-body">
        <div class="card">
          <div class="card-title">Cliente</div>
          <div style="display:flex;gap:6px;margin-bottom:8px">
            <div class="type-btn ${STATE.clienteExiste?'sel-blue':''}" style="flex:1" onclick="setClienteExiste(true)">Existente en SAP</div>
            <div class="type-btn ${!STATE.clienteExiste?'sel-amber':''}" style="flex:1" onclick="setClienteExiste(false)">Cliente nuevo</div>
          </div>
          <div class="field-label">Nombre del cliente <span class="req">*</span></div>
          <input type="text" id="cliente-nombre" placeholder="Buscar o escribir cliente...">
          <div class="field-label">Nombre de la oportunidad <span class="opt">opc</span></div>
          <input type="text" id="opp-nombre" placeholder="Si lo recuerdas, ayuda a identificar la oportunidad">
        </div>
        <div class="card">
          <div class="card-title">Razón de pérdida <span class="req">*</span></div>
          <div class="stage-grid">${razonesHTML}</div>
          ${detalleHTML}
        </div>
        <div class="card">
          <div class="card-title">Notas de cierre <span class="opt">opc</span></div>
          <textarea id="notas" placeholder="Contexto general (opcional)..."></textarea>
        </div>
        <button class="save-btn red" onclick="guardar()">Guardar cierre perdido</button>
      </div>`;
    restaurarCampos();
    return;
  }

  // CASO 4 (default): Avance, visita, demo. Para Avance, header back va al
  // sub-selector y opp-nombre sube al card Cliente. Para visita/demo, igual a hoy.
  const t=TIPO_CONFIG[STATE.tipo];
  const div=CONFIG.divisiones[STATE.division];
  const esAvance = STATE.tipo === 'oportunidad';

  let marcasHTML='';
  Object.entries(div.marcas).forEach(([linea,marcas])=>{
    const pills=marcas.map(m=>`<span class="pill ${STATE.marca===m?'sel-blue':''}" onclick="selMarca('${m}')">${m}</span>`).join('');
    marcasHTML+=`<div class="field-label" style="margin-bottom:4px">${linea}</div><div class="pill-row">${pills}</div>`;
  });

  const etapasHTML=CONFIG.etapas.map(e=>`<div class="stage-item ${STATE.etapa===e.id?'sel':''}" onclick="selEtapa('${e.id}')">${e.label}</div>`).join('');

  const lideresHTML=CONFIG.lideres.map(l=>{
    const sel=STATE.lideres.includes(l.email);
    return `<div class="lider-item" onclick="toggleLider('${l.email}')">
      <div class="avatar" style="${sel?'background:#EAF3DE;color:#0F6E56':''}">${iniciales(l.nombre)}</div>
      <div><div class="lider-name">${l.nombre}</div><div class="lider-lineas">${l.lineas.join(' · ')}</div></div>
      ${sel?'<span style="margin-left:auto;color:var(--green);font-size:13px">✓</span>':''}
    </div>`;
  }).join('');

  const compHTML=CONFIG.competidores.map(c=>`<span class="pill ${STATE.competidores.includes(c)?'sel-blue':''}" onclick="toggleComp('${c}')">${c}</span>`).join('');

  const divHTML=Object.entries(CONFIG.divisiones).map(([key,val])=>`<div class="type-btn ${STATE.division===key?'sel-blue':''}" onclick="selDivision('${key}')" style="font-size:11px">${val.label.split(' ')[0]}</div>`).join('');

  // opp-nombre: Avance lo renderiza en el card Cliente; visita/demo lo dejan en opSection
  const oppNombreEnCliente = esAvance ? `
        <div class="field-label">Nombre de la oportunidad <span class="req">*</span></div>
        <input type="text" id="opp-nombre" placeholder="Ej: PS/TMP1700/X1">` : '';

  const oppNombreEnOp = esAvance ? '' : `
    <div class="field-label">Nombre de la oportunidad <span class="req">*</span></div>
    <input type="text" id="opp-nombre" placeholder="Ej: PS/TMP1700/X1">`;

  const contactoHTML=`<div class="card purple-card">
    <div class="card-title">¿Hay un contacto nuevo que agregar?</div>
    <div class="toggle-row">
      <div class="toggle ${STATE.contactoNuevo?'on':''}" onclick="toggleContacto()"><div class="toggle-dot"></div></div>
      <span class="toggle-label">${STATE.contactoNuevo?'Sí, agregar contacto':'No por el momento'}</span>
    </div>
    ${STATE.contactoNuevo?`
      <div class="field-label">Nombre completo <span class="req">*</span></div>
      <input type="text" id="c-nombre" placeholder="Nombre del contacto">
      <div class="field-row">
        <div><div class="field-label">Puesto <span class="opt">opc</span></div><input type="text" id="c-puesto" placeholder="Puesto"></div>
        <div><div class="field-label">Teléfono <span class="opt">opc</span></div><input type="tel" id="c-tel" placeholder="844 000 0000"></div>
      </div>
      <div class="field-label">Email <span class="opt">opc</span></div>
      <input type="email" id="c-email" placeholder="correo@empresa.com">
    `:''}
  </div>`;

  const liderSection=STATE.tipo==='demo'?`<div class="card">
    <div class="card-title">Líder de línea presente</div>
    <div class="search-box">
      <span class="search-icon">&#9906;</span>
      <input type="text" placeholder="Buscar líder..." oninput="filtrarLideres(this.value)">
    </div>
    <div id="lideres-list">${lideresHTML}</div>
    ${STATE.lideres.length>0?'<div class="nota-hint">Recibirán notificación para confirmar su asistencia</div>':''}
  </div>`:'';

  const opSection=STATE.tipo!=='lead'?`<div class="card">
    <div class="card-title">Etapa de la oportunidad <span class="req">*</span></div>
    <div class="stage-grid">${etapasHTML}</div>
    <div class="field-row" style="align-items:end">
      <div>
        <div class="field-label">Monto estimado <span class="req">*</span></div>
        <div style="display:flex;gap:6px">
          <select id="moneda" onchange="selMoneda(this.value)" style="width:80px;margin-bottom:8px">
            <option value="MXP" ${STATE.moneda==='MXP'?'selected':''}>MXP</option>
            <option value="USD" ${STATE.moneda==='USD'?'selected':''}>USD</option>
          </select>
          <input type="number" id="monto" placeholder="0" style="flex:1">
        </div>
      </div>
      <div>
        <div class="field-label">Cierre estimado <span class="req">*</span></div>
        <input type="date" id="cierre">
      </div>
    </div>
    <div class="field-label">Presupuesto cliente <span class="opt">opc</span></div>
    <input type="number" id="presupuesto" placeholder="¿El cliente mencionó presupuesto?">${oppNombreEnOp}
  </div>`:'';

  // Header back button: Avance → sub-selector; visita/demo → Home
  const backBtn = esAvance
    ? `<button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>`
    : `<button class="back-btn" onclick="mostrarScreen('screen-home')">← Inicio</button>`;

  $('screen-form').innerHTML=`
    <header class="top-bar ${t.color}">
      ${backBtn}
      <div style="color:white;font-size:14px;font-weight:500">${t.label}</div>
      <div style="width:50px"></div>
    </header>
    <div class="form-body">
      <div class="card">
        <div class="card-title">Cliente</div>
        <div style="display:flex;gap:6px;margin-bottom:8px">
          <div class="type-btn ${STATE.clienteExiste?'sel-blue':''}" style="flex:1" onclick="setClienteExiste(true)">Existente en SAP</div>
          <div class="type-btn ${!STATE.clienteExiste?'sel-amber':''}" style="flex:1" onclick="setClienteExiste(false)">Cliente nuevo</div>
        </div>
        <div class="field-label">Nombre del cliente <span class="req">*</span></div>
        <input type="text" id="cliente-nombre" placeholder="Buscar o escribir cliente...">
        <div class="field-label">Contacto principal <span class="opt">opc</span></div>
        <input type="text" id="cliente-contacto" placeholder="Nombre del contacto">${oppNombreEnCliente}
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
        <div class="card-title">Competidor <span class="opt">opc</span></div>
        <div class="pill-row">${compHTML}</div>
      </div>
      ${liderSection}
      <div class="card">
        <div class="card-title">Notas de la visita <span class="opt">opc</span></div>
        <textarea id="notas" placeholder="Observaciones, siguientes pasos, acuerdos..."></textarea>
      </div>
      <button class="save-btn ${t.color}" onclick="guardar()">Guardar y sincronizar con SAP</button>
    </div>`;

  restaurarCampos();
}

function selDivision(key){guardarCampos();STATE.division=key;STATE.marca=null;renderForm()}
function selMarca(m){guardarCampos();STATE.marca=STATE.marca===m?null:m;renderForm()}
function selEtapa(id){guardarCampos();STATE.etapa=STATE.etapa===id?null:id;renderForm()}
function selMoneda(v){guardarCampos();STATE.moneda=v;renderForm()}
function toggleComp(c){guardarCampos();const i=STATE.competidores.indexOf(c);if(i>-1)STATE.competidores.splice(i,1);else STATE.competidores.push(c);renderForm()}
function toggleLider(email){guardarCampos();const i=STATE.lideres.indexOf(email);if(i>-1)STATE.lideres.splice(i,1);else STATE.lideres.push(email);renderForm()}
function toggleContacto(){guardarCampos();STATE.contactoNuevo=!STATE.contactoNuevo;renderForm()}
function setClienteExiste(val){guardarCampos();STATE.clienteExiste=val;renderForm()}

function selResultadoOpp(resultado){STATE.resultado=resultado;renderForm()}
function volverSelectorOpp(){guardarCampos();STATE.resultado=null;renderForm()}
function selRazonPerdida(razon){guardarCampos();STATE.razonPerdida=razon;renderForm()}

function filtrarLideres(q){
  const lista=$('lideres-list');if(!lista)return;
  lista.querySelectorAll('.lider-item').forEach(item=>{
    const n=item.querySelector('.lider-name').textContent.toLowerCase();
    item.style.display=n.includes(q.toLowerCase())?'':'none';
  });
}

async function guardar() {
  guardarCampos();
  const cliente = STATE.campos['cliente-nombre']?.trim();
  if (!cliente) { alert('Falta el nombre del cliente'); return; }

  const esGanada = STATE.tipo === 'oportunidad' && STATE.resultado === 'ganada';
  const esPerdida = STATE.tipo === 'oportunidad' && STATE.resultado === 'perdida';
  const esCierre = esGanada || esPerdida;

  // Validaciones por caso
  if (!esCierre) {
    // Avance, visita, demo
    const producto = STATE.campos['producto']?.trim();
    if (!STATE.marca) { alert('Selecciona una marca'); return; }
    if (!producto) { alert('Escribe el producto o modelo'); return; }
    if (STATE.tipo !== 'lead' && !STATE.etapa) { alert('Selecciona la etapa de la oportunidad'); return; }
  } else if (esGanada) {
    if (!STATE.campos['monto-final']?.trim()) { alert('Escribe el monto final del cierre'); return; }
    if (!STATE.campos['fecha-cierre-real']?.trim()) { alert('Indica la fecha de cierre real'); return; }
  } else if (esPerdida) {
    if (!STATE.razonPerdida) { alert('Selecciona la razón de pérdida'); return; }
  }

  // Construir RazonPerdida con sufijo si aplica
  let razonPerdidaFinal = '';
  if (esPerdida) {
    razonPerdidaFinal = STATE.razonPerdida;
    if (STATE.razonPerdida === 'Otro') {
      const detalle = STATE.campos['razon-perdida-detalle']?.trim();
      if (detalle) razonPerdidaFinal = `Otro: ${detalle}`;
    }
  }

  const registro = {
    tipo: STATE.tipo,
    modo: STATE.modo || 'campo',
    fecha: new Date().toISOString(),
    cliente: cliente,
    clienteContacto: STATE.campos['cliente-contacto'],
    clienteNuevo: !STATE.clienteExiste,
    division: esCierre ? '' : STATE.division,
    marca: esCierre ? '' : STATE.marca,
    producto: esCierre ? '' : STATE.campos['producto'],
    etapa: esCierre ? '' : STATE.etapa,
    moneda: esPerdida ? '' : STATE.moneda,
    monto: esCierre ? '' : STATE.campos['monto'],
    cierre: esCierre ? '' : STATE.campos['cierre'],
    presupuesto: esCierre ? '' : STATE.campos['presupuesto'],
    oppNombre: STATE.campos['opp-nombre'] || '',
    competidores: esCierre ? [] : STATE.competidores,
    lideres: esCierre ? [] : STATE.lideres,
    notas: STATE.campos['notas'],
    contactoNuevo: STATE.contactoNuevo ? {
      nombre: STATE.campos['c-nombre'], puesto: STATE.campos['c-puesto'],
      telefono: STATE.campos['c-tel'], email: STATE.campos['c-email']
    } : null,
    asesor: CONFIG.usuario.email,
    resultadoCierre: esGanada ? 'ganada' : (esPerdida ? 'perdida' : ''),
    razonPerdida: razonPerdidaFinal,
    montoFinal: esGanada ? STATE.campos['monto-final'] : '',
    fechaCierreReal: esGanada ? STATE.campos['fecha-cierre-real'] : ''
  };

  console.log('Registro a enviar a SAP:', registro);

  // Modo gabinete: sin GPS, sin alerta de proximidad. Modo campo: captura normal.
  let registroConGPS;
  if (STATE.modo === 'gabinete') {
    registroConGPS = registro;
  } else {
    registroConGPS = await capturarGPSAlGuardar(registro);
    if (!registroConGPS) return; // usuario canceló por alerta de proximidad
  }

  // Mostrar progreso
  const btnGuardar = document.querySelector('.save-btn');
  if (btnGuardar) { btnGuardar.textContent = 'Sincronizando con SAP...'; btnGuardar.disabled = true; }

  // Sincronizar con SAP y SharePoint
  const resultado = await sincronizarConSAP(registroConGPS);

  const gpsInfo = registroConGPS.gps
    ? `📍 ${registroConGPS.gps.lat.toFixed(4)}, ${registroConGPS.gps.lng.toFixed(4)}`
    : (STATE.modo === 'gabinete' ? '📝 Modo gabinete' : '📍 GPS no disponible');

  const titulo = esGanada ? 'OPORTUNIDAD GANADA'
               : esPerdida ? 'OPORTUNIDAD PERDIDA'
               : registro.tipo.toUpperCase();
  const subtitulo = esGanada ? `Monto final: ${STATE.moneda} ${STATE.campos['monto-final']}`
                  : esPerdida ? `Razón: ${razonPerdidaFinal}`
                  : `${registro.marca} · ${registro.producto}`;

  if (resultado) {
    const sapInfo = resultado.sapOppId
      ? `SAP Oportunidad #${resultado.sapOppId}`
      : resultado.errores.length ? `⚠️ ${resultado.errores[0]}` : 'SAP: registrado';
    alert(`✅ Registro guardado\n\n${titulo} — ${registro.cliente}\n${subtitulo}\n\n${gpsInfo}\n${sapInfo}`);
  } else {
    alert(`✅ Guardado localmente\n\n${titulo} — ${registro.cliente}\n${subtitulo}\n\n${gpsInfo}\nSAP: pendiente de sincronizar`);
  }

  mostrarScreen('screen-home');
}

// ── Modal Check-in ───────────────────────────────────────────────────────────

function mostrarModalCheckin() {
  const nombre = prompt('¿Con qué cliente estás?');
  if (nombre && nombre.trim()) {
    checkin(nombre.trim());
  }
}

// ── Exponer funciones globales ───────────────────────────────────────────────

window.irA                  = irA;
window.mostrarScreen        = mostrarScreen;
window.selDivision          = selDivision;
window.selMarca             = selMarca;
window.selEtapa             = selEtapa;
window.selMoneda            = selMoneda;
window.toggleComp           = toggleComp;
window.toggleLider          = toggleLider;
window.toggleContacto       = toggleContacto;
window.setClienteExiste     = setClienteExiste;
window.filtrarLideres       = filtrarLideres;
window.guardar              = guardar;
window.mostrarModalCheckin  = mostrarModalCheckin;
window.renderHome           = renderHome;
window.seleccionarModo      = seleccionarModo;
window.cambiarModo          = cambiarModo;
window.selResultadoOpp      = selResultadoOpp;
window.volverSelectorOpp    = volverSelectorOpp;
window.selRazonPerdida      = selRazonPerdida;

window.addEventListener('DOMContentLoaded',()=>{
  const f=$('fecha-hoy');if(f)f.textContent=fechaHoy();
  const u=$('user-initials');if(u)u.textContent=iniciales(CONFIG.usuario.nombre);
  // Si hay check-in activo persistido, forzar modo Campo automáticamente
  // (el chip "Cambiar modo" permite al asesor cerrar la visita y salir).
  STATE.modo = localStorage.getItem('vinssa_checkin_activo') ? 'campo' : null;
  renderHome();
  mostrarScreen('screen-home');
});
