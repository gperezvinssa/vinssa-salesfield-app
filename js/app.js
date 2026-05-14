function $(id){return document.getElementById(id)}
function fechaHoy(){return new Date().toLocaleDateString('es-MX',{weekday:'long',day:'numeric',month:'long',timeZone:CONFIG.timezone})}
function iniciales(n){return n.split(' ').slice(0,2).map(x=>x[0]).join('').toUpperCase()}
function mostrarScreen(id){
  // Gestiona inline style.display además de la clase .active para evitar choque
  // con navSwitch (bottom nav) que también escribe inline display. Sin esto,
  // navegar via bottom nav y luego tapear un sub-flujo deja el form oculto por
  // el inline display:none residual aunque tenga .active.
  document.querySelectorAll('.screen').forEach(s=>{
    s.classList.remove('active');
    s.style.display='none';
  });
  const s=$(id);
  if(s){
    s.classList.add('active');
    // 'flex' (no 'block') para que coincida con .screen.active { display: flex }
    // del CSS — necesario para que .form-body con flex:1 + overflow-y:auto
    // scrollee internamente en lugar de empujar el documento.
    s.style.display='flex';
    window.scrollTo(0,0);
  }
}

const STATE={
  modo:null,                           // null=selector | 'campo' | 'gabinete'
  tipo:null,division:'trazabilidad',marca:null,etapa:null,
  resultado:null,                      // null=sub-selector | 'avance' | 'ganada' | 'perdida'
  razonPerdida:null,                   // string elegido cuando resultado='perdida'
  lideres:[],competidores:[],contactoNuevo:false,moneda:'MXP',
  clienteExiste:true,
  campos:{},
  // Piloto de Actualizar Oportunidad: combobox de clientes y oportunidades.
  asesorSAP:null,                      // se resuelve en auth.js desde EMAIL_A_ASESOR
  oportunidades:[],                    // poblado async tras login
  opsLoading:true,                     // true hasta que termine la primera carga
  clienteCommit:'',                    // cliente "settled" (post-commit del combobox)
  oportunidadSeleccionada:null         // objeto opp completo cuando el asesor elige una
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
  STATE.oportunidadSeleccionada=null;
  STATE.clienteCommit='';
  renderForm();mostrarScreen(screenId);
}

// ── Helpers piloto de Actualizar Oportunidad ─────────────────────────────────
// Match exacto normalizado: uppercase + NFD strip acentos + trim.
function normCliente(str){
  return String(str||'').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,'').trim();
}

// Clientes únicos en STATE.oportunidades, ordenados alfabéticamente.
function clientesUnicos(){
  const set=new Set();
  STATE.oportunidades.forEach(o=>{ if(o.Cliente) set.add(o.Cliente); });
  return Array.from(set).sort();
}

// Oportunidades del cliente actual (match exacto normalizado).
function opsDelCliente(cliente){
  if(!cliente) return [];
  const norm=normCliente(cliente);
  return STATE.oportunidades.filter(o=>normCliente(o.Cliente)===norm);
}

// ── Combobox custom — vanilla autocomplete con lista flotante ───────────────
// Por qué no <select>: en celular abre selector nativo del OS, no es buscable,
// y los asesores no notan que es interactivo. Tampoco <input list=datalist>:
// el oninput re-renderiza el form lo cual destruye el input y el focus.
// Este módulo separa tipeo (filtrado DOM-local) de commit (re-render del form),
// preservando focus durante la búsqueda.

const CB_STATE = {};

function _cbInput(boxId){ return document.getElementById(boxId+'-input'); }
function _cbList(boxId){  return document.getElementById(boxId+'-list'); }
function _cbEsc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function _cbAttr(s){ return String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;'); }

// Inicializa estado del combobox. Llamada desde renderBloqueClienteYOpp.
function cbInit(boxId, items, kind){
  CB_STATE[boxId] = { items, filter:'', highlighted:-1, open:false, kind };
}

function cbOpen(boxId){
  const cb = CB_STATE[boxId]; if(!cb) return;
  cb.open = true; cb.filter = ''; cb.highlighted = -1;
  cbRenderList(boxId);
  setTimeout(()=>{
    const el = document.getElementById(boxId);
    // 'nearest' solo scrollea si el input NO está visible — evita forzar scroll
    // del documento al abrir la lista, que creaba un segundo scrollbar persistente.
    if(el) el.scrollIntoView({ block:'nearest', behavior:'smooth' });
  }, 50);
}

function cbClose(boxId){
  const cb = CB_STATE[boxId]; if(!cb) return;
  cb.open = false;
  const list = _cbList(boxId); if(list) list.hidden = true;
}

function cbFilter(boxId, value){
  const cb = CB_STATE[boxId]; if(!cb) return;
  cb.filter = value; cb.highlighted = -1; cb.open = true;
  cbRenderList(boxId);
}

function cbRenderList(boxId){
  const cb = CB_STATE[boxId]; const list = _cbList(boxId);
  if(!cb || !list) return;
  const f = normCliente(cb.filter);
  const matches = f
    ? cb.items.filter(i => normCliente((i.label||'')+' '+(i.subtext||'')).includes(f))
    : cb.items;
  cb._lastMatches = matches;
  if(matches.length === 0){
    const msg = cb.kind === 'cliente'
      ? 'Sin coincidencias en tus clientes con oportunidades activas. Sigue escribiendo — al confirmar se tratará como cliente sin opp en SAP.'
      : 'Sin coincidencias en este cliente.';
    list.innerHTML = `<div class="cb-empty">${msg}</div>`;
  } else {
    list.innerHTML = matches.map((m,i)=>{
      const hl = i === cb.highlighted ? ' cb-hl' : '';
      const sub = m.subtext ? `<div class="cb-option-sub">${_cbEsc(m.subtext)}</div>` : '';
      return `<div class="cb-option${hl}" data-v="${_cbAttr(m.value)}"
                   onmousedown="event.preventDefault(); cbSelect('${boxId}', this.dataset.v)">
        <div class="cb-option-label">${_cbEsc(m.label)}</div>${sub}
      </div>`;
    }).join('');
  }
  list.hidden = !cb.open;
  list.scrollTop = 0;
  // Scroll highlighted into view (keyboard nav)
  const hlEl = list.querySelector('.cb-hl');
  if(hlEl) hlEl.scrollIntoView({ block:'nearest' });
}

function cbSelect(boxId, value){
  const cb = CB_STATE[boxId]; if(!cb) return;
  const item = cb.items.find(i => String(i.value) === String(value));
  if(!item) return;
  const input = _cbInput(boxId);
  if(input) input.value = item.label;
  cbClose(boxId);
  if(cb.kind === 'cliente') cbCommitCliente(item.value);
  else cbCommitOpp(item.value);
}

function cbBlur(boxId){
  // Delay corto para tolerar variaciones de orden de evento entre browsers.
  // (mousedown.preventDefault() en options ya previene blur en el caso normal,
  // pero algunos browsers móviles disparan blur antes del mousedown.)
  setTimeout(()=>{
    const cb = CB_STATE[boxId]; if(!cb || !cb.open) return;
    cbClose(boxId);
    const input = _cbInput(boxId);
    const typed = input ? input.value.trim() : '';
    if(!typed) return;
    const exact = cb.items.find(i => normCliente(i.label) === normCliente(typed));
    if(exact){
      if(input) input.value = exact.label;
      if(cb.kind === 'cliente') cbCommitCliente(exact.value);
      else cbCommitOpp(exact.value);
    } else if(cb.kind === 'cliente'){
      // Sin match → commit como cliente sin opps en SAP (fallback de texto libre).
      cbCommitCliente(typed);
    } else {
      // Opp sin match → revertir al label de la opp previamente seleccionada (o vaciar).
      const sel = STATE.oportunidadSeleccionada;
      if(input) input.value = sel
        ? (sel.Descripcion + (sel.Marca ? ' · '+sel.Marca : '')).trim()
        : '';
    }
  }, 150);
}

function cbKeydown(event, boxId){
  const cb = CB_STATE[boxId]; if(!cb) return;
  if(event.key === 'ArrowDown'){
    event.preventDefault();
    if(!cb.open){ cbOpen(boxId); return; }
    const m = cb._lastMatches || cb.items;
    cb.highlighted = Math.min(cb.highlighted + 1, m.length - 1);
    cbRenderList(boxId);
  } else if(event.key === 'ArrowUp'){
    event.preventDefault();
    if(!cb.open) return;
    cb.highlighted = Math.max(cb.highlighted - 1, 0);
    cbRenderList(boxId);
  } else if(event.key === 'Enter'){
    if(!cb.open) return;
    event.preventDefault();
    const m = cb._lastMatches || cb.items;
    if(cb.highlighted >= 0 && m[cb.highlighted]) cbSelect(boxId, m[cb.highlighted].value);
    else if(m.length === 1) cbSelect(boxId, m[0].value);
    else { const inp = _cbInput(boxId); if(inp) inp.blur(); }
  } else if(event.key === 'Escape'){
    cbClose(boxId);
  }
}

// Commit del combobox de cliente. Idempotente: si el cliente no cambia,
// no resetea la opp seleccionada (evita perder selección al hacer blur sin cambio).
function cbCommitCliente(value){
  if(STATE.clienteCommit !== value){
    STATE.clienteCommit = value;
    STATE.oportunidadSeleccionada = null;
    const ops = opsDelCliente(value);
    if(ops.length === 1) STATE.oportunidadSeleccionada = ops[0];
    delete STATE.campos['monto-final'];
  }
  renderForm();
}

// Commit del combobox de oportunidad. Idempotente como cbCommitCliente.
function cbCommitOpp(numOpp){
  const sel = STATE.oportunidades.find(o => String(o.NumOportunidad) === String(numOpp)) || null;
  const cambia = !STATE.oportunidadSeleccionada
    || String(STATE.oportunidadSeleccionada.NumOportunidad) !== String(numOpp);
  if(cambia){
    STATE.oportunidadSeleccionada = sel;
    delete STATE.campos['monto-final'];
  }
  renderForm();
}

// Botón refresh inline en el form: recarga Oportunidades.xlsx.
async function refrescarOportunidades(){
  STATE.opsLoading = true;
  const btn = $('btn-refresh-opps');
  if(btn){ btn.disabled = true; btn.textContent = '...'; }
  STATE.oportunidades = await cargarOportunidadesAsesor();
  // Si la opp pre-seleccionada ya no existe en la nueva carga, limpiar.
  if(STATE.oportunidadSeleccionada){
    const sigue = STATE.oportunidades.find(o =>
      String(o.NumOportunidad) === String(STATE.oportunidadSeleccionada.NumOportunidad));
    STATE.oportunidadSeleccionada = sigue || null;
  }
  STATE.opsLoading = false;
  renderForm();
}

// Hook llamado por auth.js cuando termina la carga async inicial.
function onOportunidadesCargadas(){
  STATE.opsLoading = false;
  if(STATE.tipo === 'oportunidad' && STATE.resultado !== null){
    renderForm();
  }
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

// Mensaje cuando el usuario no es asesor en el piloto (asesorSAP null).
function renderMensajePiloto(){
  return `<div class="card msg-piloto">
    <div style="font-size:14px;font-weight:500;margin-bottom:6px">Función en piloto</div>
    <div style="font-size:13px;color:var(--color-text-secondary);line-height:1.5">
      Esta función está actualmente en piloto y disponible solo para asesores asignados.
      Si crees que esto es un error, contacta a Gerardo Pérez (gperez@vinssa.com).
    </div>
  </div>`;
}

// Bloque compartido: card Cliente (combobox custom autocomplete) + card Oportunidad
// (combobox cuando el cliente tiene opps, o fallback de texto libre).
// Devuelve { html, esFallback, opSel } para que cada caso del form decida qué hacer.
function renderBloqueClienteYOpp(){
  const clienteCommit = STATE.clienteCommit || '';
  const ops = opsDelCliente(clienteCommit);
  const opSel = STATE.oportunidadSeleccionada;
  const esFallback = clienteCommit.trim().length > 0 && ops.length === 0;
  const refreshBtn = `<button type="button" id="btn-refresh-opps" class="btn-refresh" onclick="refrescarOportunidades()" title="Recargar oportunidades de SAP">🔄</button>`;

  // Mientras se cargan oportunidades de SAP, solo el hint. Evita que el combobox
  // aparezca con 0 items y luego re-render destruya el input cuando termina la carga.
  if(STATE.opsLoading){
    return {
      html: `<div class="card">
        <div class="card-title-row">
          <div class="card-title">Cliente <span class="req">*</span></div>
          ${refreshBtn}
        </div>
        <div class="hint-piloto">Cargando tus oportunidades de SAP...</div>
      </div>`,
      esFallback: false,
      opSel: null
    };
  }

  // Combobox de cliente
  const clienteItems = clientesUnicos().map(c => ({ value: c, label: c }));
  cbInit('cb-cliente', clienteItems, 'cliente');
  const clienteListHTML = clienteItems.length > 0
    ? clienteItems.map(m => `<div class="cb-option" data-v="${_cbAttr(m.value)}"
        onmousedown="event.preventDefault(); cbSelect('cb-cliente', this.dataset.v)">
        <div class="cb-option-label">${_cbEsc(m.label)}</div>
      </div>`).join('')
    : `<div class="cb-empty">No tienes oportunidades activas en SAP. Escribe el cliente manualmente.</div>`;

  const clienteCard = `<div class="card">
    <div class="card-title-row">
      <div class="card-title">Cliente <span class="req">*</span></div>
      ${refreshBtn}
    </div>
    <div class="combobox" id="cb-cliente">
      <input type="text" id="cb-cliente-input"
             value="${_cbAttr(clienteCommit)}"
             placeholder="Buscar cliente..."
             oninput="cbFilter('cb-cliente', this.value)"
             onfocus="cbOpen('cb-cliente')"
             onblur="cbBlur('cb-cliente')"
             onkeydown="cbKeydown(event, 'cb-cliente')"
             autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
      <div class="combobox-list" id="cb-cliente-list" hidden
           onmousedown="event.preventDefault()">${clienteListHTML}</div>
    </div>
  </div>`;

  // Combobox de oportunidad o fallback
  let oppCard = '';
  if(clienteCommit && ops.length > 0){
    const oppItems = ops.map(o => ({
      value: String(o.NumOportunidad),
      label: (String(o.Descripcion||'(sin descripción)') + (o.Marca ? ' · '+o.Marca : '')).trim(),
      subtext: `$${(o.MontoEstimado||0).toLocaleString('es-MX')} MXP · etapa: ${o.Etapa||'—'}`
    }));
    cbInit('cb-opp', oppItems, 'opp');
    const oppInputValue = opSel
      ? (String(opSel.Descripcion||'(sin descripción)') + (opSel.Marca ? ' · '+opSel.Marca : '')).trim()
      : '';
    const oppListHTML = oppItems.map(m => `<div class="cb-option" data-v="${_cbAttr(m.value)}"
        onmousedown="event.preventDefault(); cbSelect('cb-opp', this.dataset.v)">
        <div class="cb-option-label">${_cbEsc(m.label)}</div>
        <div class="cb-option-sub">${_cbEsc(m.subtext)}</div>
      </div>`).join('');
    const banner = opSel ? `<div class="opp-banner">
      <strong>${_cbEsc(opSel.Descripcion||'(sin descripción)')}</strong>
      · ${_cbEsc(opSel.Marca||'(sin marca)')}
      · originalmente $${(opSel.MontoEstimado||0).toLocaleString('es-MX')} MXP
      · etapa actual: ${_cbEsc(opSel.Etapa||'(sin etapa)')}
    </div>` : '';
    oppCard = `<div class="card">
      <div class="card-title">Selecciona la oportunidad <span class="req">*</span></div>
      <div class="combobox" id="cb-opp">
        <input type="text" id="cb-opp-input"
               value="${_cbAttr(oppInputValue)}"
               placeholder="Buscar oportunidad..."
               oninput="cbFilter('cb-opp', this.value)"
               onfocus="cbOpen('cb-opp')"
               onblur="cbBlur('cb-opp')"
               onkeydown="cbKeydown(event, 'cb-opp')"
               autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
        <div class="combobox-list" id="cb-opp-list" hidden
             onmousedown="event.preventDefault()">${oppListHTML}</div>
      </div>
      ${banner}
    </div>`;
  } else if(clienteCommit){
    oppCard = `<div class="card">
      <div class="card-title">Oportunidad</div>
      <div class="hint-piloto">No hay oportunidades activas con este cliente en SAP. Captura el nombre manualmente.</div>
      <div class="field-label">Nombre de la oportunidad <span class="req">*</span></div>
      <input type="text" id="opp-nombre" placeholder="Describe la oportunidad...">
    </div>`;
  }

  return { html: clienteCard + oppCard, esFallback, opSel };
}

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
    // Bloqueo piloto: usuario logueado no está mapeado a ningún asesor SAP.
    if (!STATE.asesorSAP) {
      $('screen-form').innerHTML = `
        <header class="top-bar green">
          <button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>
          <div style="color:white;font-size:14px;font-weight:500">Oportunidad ganada</div>
          <div style="width:50px"></div>
        </header>
        <div class="form-body">${renderMensajePiloto()}</div>`;
      return;
    }
    const bloque = renderBloqueClienteYOpp();
    const opSel = bloque.opSel;
    // Auto-llenado: si hay opp seleccionada y el usuario aún no editó el monto,
    // pre-cargar MontoEstimado de la opp. STATE.campos['monto-final'] === undefined
    // = nunca tocado; '' = el usuario lo borró a propósito (respetar).
    const montoPre = STATE.campos['monto-final'] !== undefined
      ? STATE.campos['monto-final']
      : (opSel ? String(opSel.MontoEstimado || '') : '');
    // Solo mostrar la sección de cierre si ya hay opp seleccionada o si está en fallback con cliente escrito.
    const yaPuedeCerrar = !!opSel || bloque.esFallback;
    const cierreCard = yaPuedeCerrar ? `
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
                <input type="number" id="monto-final" placeholder="0" style="flex:1" value="${montoPre}">
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
        <button class="save-btn green" onclick="guardar()">Guardar cierre ganado</button>` : '';
    $('screen-form').innerHTML = `
      <header class="top-bar green">
        <button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>
        <div style="color:white;font-size:14px;font-weight:500">Oportunidad ganada</div>
        <div style="width:50px"></div>
      </header>
      <div class="form-body">
        ${bloque.html}
        ${cierreCard}
      </div>`;
    // Excluir monto-final de restaurarCampos en este render — ya lo pusimos
    // explícitamente vía value="" para que respete el auto-llenado.
    const guardado = STATE.campos['monto-final'];
    delete STATE.campos['monto-final'];
    restaurarCampos();
    if (guardado !== undefined) STATE.campos['monto-final'] = guardado;
    return;
  }

  // CASO 3: Form Perdida
  if (STATE.tipo === 'oportunidad' && STATE.resultado === 'perdida') {
    if (!STATE.asesorSAP) {
      $('screen-form').innerHTML = `
        <header class="top-bar red">
          <button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>
          <div style="color:white;font-size:14px;font-weight:500">Oportunidad perdida</div>
          <div style="width:50px"></div>
        </header>
        <div class="form-body">${renderMensajePiloto()}</div>`;
      return;
    }
    const bloque = renderBloqueClienteYOpp();
    const opSel = bloque.opSel;
    const yaPuedeCerrar = !!opSel || bloque.esFallback;
    const razonesHTML = RAZONES_PERDIDA.map(r =>
      `<div class="stage-item ${STATE.razonPerdida===r?'sel':''}" onclick="selRazonPerdida('${r.replace(/'/g,"\\'")}')">${r}</div>`
    ).join('');
    const detalleHTML = STATE.razonPerdida === 'Otro' ? `
      <div class="field-label" style="margin-top:10px">Detalle <span class="opt">opc</span></div>
      <textarea id="razon-perdida-detalle" placeholder="Especifica la razón (recomendado para análisis posterior)"></textarea>
    ` : '';
    const cierreCard = yaPuedeCerrar ? `
        <div class="card">
          <div class="card-title">Razón de pérdida <span class="req">*</span></div>
          <div class="stage-grid">${razonesHTML}</div>
          ${detalleHTML}
        </div>
        <div class="card">
          <div class="card-title">Notas de cierre <span class="opt">opc</span></div>
          <textarea id="notas" placeholder="Contexto general (opcional)..."></textarea>
        </div>
        <button class="save-btn red" onclick="guardar()">Guardar cierre perdido</button>` : '';
    $('screen-form').innerHTML = `
      <header class="top-bar red">
        <button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>
        <div style="color:white;font-size:14px;font-weight:500">Oportunidad perdida</div>
        <div style="width:50px"></div>
      </header>
      <div class="form-body">
        ${bloque.html}
        ${cierreCard}
      </div>`;
    restaurarCampos();
    return;
  }

  // CASO 4: Form Avanzó de etapa
  if (STATE.tipo === 'oportunidad' && STATE.resultado === 'avance') {
    if (!STATE.asesorSAP) {
      $('screen-form').innerHTML = `
        <header class="top-bar blue">
          <button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>
          <div style="color:white;font-size:14px;font-weight:500">Avanzó de etapa</div>
          <div style="width:50px"></div>
        </header>
        <div class="form-body">${renderMensajePiloto()}</div>`;
      return;
    }
    const bloque = renderBloqueClienteYOpp();
    const opSel = bloque.opSel;
    const yaPuedeAvanzar = !!opSel || bloque.esFallback;
    const etapasHTML = CONFIG.etapas.map(e =>
      `<div class="stage-item ${STATE.etapa===e.id?'sel':''}" onclick="selEtapa('${e.id}')">${e.label} <span style="opacity:.5;font-size:11px">${e.pct}%</span></div>`
    ).join('');
    const avanceCard = yaPuedeAvanzar ? `
        <div class="card">
          <div class="card-title">Nueva etapa <span class="req">*</span></div>
          <div class="stage-grid">${etapasHTML}</div>
        </div>
        <div class="card">
          <div class="card-title">Notas del avance <span class="opt">opc</span></div>
          <textarea id="notas" placeholder="Qué pasó en esta etapa, siguientes pasos..."></textarea>
        </div>
        <button class="save-btn blue" onclick="guardar()">Guardar avance de etapa</button>` : '';
    $('screen-form').innerHTML = `
      <header class="top-bar blue">
        <button class="back-btn" onclick="volverSelectorOpp()">← Atrás</button>
        <div style="color:white;font-size:14px;font-weight:500">Avanzó de etapa</div>
        <div style="width:50px"></div>
      </header>
      <div class="form-body">
        ${bloque.html}
        ${avanceCard}
      </div>`;
    restaurarCampos();
    return;
  }

  // CASO 5 (default): Visita y Demo. Forms originales sin cambios.
  const t=TIPO_CONFIG[STATE.tipo];
  const div=CONFIG.divisiones[STATE.division];
  const esAvance = false; // ya no entramos aquí desde avance; conservado por compatibilidad

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
function selMoneda(v){
  guardarCampos();
  STATE.moneda=v;
  // Reset explícito (string vacío, no delete) para que el asesor reescriba
  // el monto en la nueva moneda — evita conversiones erróneas auto-llenadas desde la opp.
  STATE.campos['monto-final']='';
  renderForm();
}
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

  const esGanada = STATE.tipo === 'oportunidad' && STATE.resultado === 'ganada';
  const esPerdida = STATE.tipo === 'oportunidad' && STATE.resultado === 'perdida';
  const esAvance = STATE.tipo === 'oportunidad' && STATE.resultado === 'avance';
  const esCierre = esGanada || esPerdida;
  const esSubflujoOpp = esCierre || esAvance;
  const opSel = STATE.oportunidadSeleccionada;

  // Cliente: sub-flujos leen del commit del combobox; visita/demo del input regular.
  const cliente = esSubflujoOpp
    ? (STATE.clienteCommit?.trim() || '')
    : (STATE.campos['cliente-nombre']?.trim() || '');
  if (!cliente) { alert('Falta el nombre del cliente'); return; }

  // Validaciones por caso
  if (esSubflujoOpp) {
    // Sin opp seleccionada y sin nombre libre → no se puede identificar la oportunidad
    if (!opSel && !STATE.campos['opp-nombre']?.trim()) {
      alert('Selecciona una oportunidad del listado o escribe su nombre');
      return;
    }
    if (esAvance && !STATE.etapa) { alert('Selecciona la nueva etapa'); return; }
    if (esGanada) {
      if (!STATE.campos['monto-final']?.trim()) { alert('Escribe el monto final del cierre'); return; }
      if (!STATE.campos['fecha-cierre-real']?.trim()) { alert('Indica la fecha de cierre real'); return; }
    }
    if (esPerdida && !STATE.razonPerdida) { alert('Selecciona la razón de pérdida'); return; }
  } else {
    // Visita / Demo
    const producto = STATE.campos['producto']?.trim();
    if (!STATE.marca) { alert('Selecciona una marca'); return; }
    if (!producto) { alert('Escribe el producto o modelo'); return; }
    if (STATE.tipo !== 'lead' && !STATE.etapa) { alert('Selecciona la etapa de la oportunidad'); return; }
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

  // Cliente canónico de SAP cuando hay opp seleccionada (evita inconsistencias por typos)
  const clienteFinal = opSel ? opSel.Cliente : cliente;
  // OpprId de SAP. Vacío en visita/demo y en fallback de texto libre.
  const opportunidadID = opSel ? String(opSel.NumOportunidad) : '';
  // OppNombre: con opp seleccionada queda vacío (el ID es suficiente);
  // en fallback o en visita/demo va el texto libre que tipeó el asesor.
  const oppNombreFinal = (esSubflujoOpp && opSel) ? '' : (STATE.campos['opp-nombre'] || '');

  const registro = {
    tipo: STATE.tipo,
    modo: STATE.modo || 'campo',
    fecha: new Date().toISOString(),
    cliente: clienteFinal,
    clienteContacto: esSubflujoOpp ? '' : STATE.campos['cliente-contacto'],
    clienteNuevo: esSubflujoOpp ? false : !STATE.clienteExiste,
    division: esSubflujoOpp ? '' : STATE.division,
    // Avanzó con opp seleccionada: heredar marca de la opp (informativo, SAP ya la tiene)
    marca: esCierre ? '' : (esAvance ? (opSel ? opSel.Marca : '') : STATE.marca),
    producto: esSubflujoOpp ? '' : STATE.campos['producto'],
    etapa: esCierre ? '' : STATE.etapa,
    moneda: esPerdida ? '' : STATE.moneda,
    monto: esSubflujoOpp ? '' : STATE.campos['monto'],
    cierre: esSubflujoOpp ? '' : STATE.campos['cierre'],
    presupuesto: esSubflujoOpp ? '' : STATE.campos['presupuesto'],
    oppNombre: oppNombreFinal,
    opportunidadID: opportunidadID,
    competidores: esSubflujoOpp ? [] : STATE.competidores,
    lideres: esSubflujoOpp ? [] : STATE.lideres,
    notas: STATE.campos['notas'],
    contactoNuevo: (esSubflujoOpp || !STATE.contactoNuevo) ? null : {
      nombre: STATE.campos['c-nombre'], puesto: STATE.campos['c-puesto'],
      telefono: STATE.campos['c-tel'], email: STATE.campos['c-email']
    },
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
               : esAvance ? 'AVANCE DE ETAPA'
               : registro.tipo.toUpperCase();
  const etapaLabel = CONFIG.etapas.find(e=>e.id===STATE.etapa)?.label || STATE.etapa || '';
  const subtitulo = esGanada ? `Monto final: ${STATE.moneda} ${STATE.campos['monto-final']}`
                  : esPerdida ? `Razón: ${razonPerdidaFinal}`
                  : esAvance ? `Nueva etapa: ${etapaLabel}`
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
window.cbOpen               = cbOpen;
window.cbClose              = cbClose;
window.cbFilter             = cbFilter;
window.cbSelect             = cbSelect;
window.cbBlur               = cbBlur;
window.cbKeydown            = cbKeydown;
window.refrescarOportunidades = refrescarOportunidades;
window.onOportunidadesCargadas = onOportunidadesCargadas;

window.addEventListener('DOMContentLoaded',()=>{
  const f=$('fecha-hoy');if(f)f.textContent=fechaHoy();
  const u=$('user-initials');if(u)u.textContent=iniciales(CONFIG.usuario.nombre);
  // Si hay check-in activo persistido, forzar modo Campo automáticamente
  // (el chip "Cambiar modo" permite al asesor cerrar la visita y salir).
  STATE.modo = localStorage.getItem('vinssa_checkin_activo') ? 'campo' : null;
  renderHome();
  // Solo navegar a screen-home si hay sesión MSAL activa. Sin sesión, dejar
  // screen-login visible para que el usuario pueda tapear el botón "Entrar".
  // iniciarApp en auth.js se encarga del switch tras login exitoso.
  let haySession = false;
  try { haySession = msalInstance.getAllAccounts().length > 0; } catch(_) {}
  if (haySession) mostrarScreen('screen-home');
});
