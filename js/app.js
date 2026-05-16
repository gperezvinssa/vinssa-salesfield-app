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
  acompanantes:[],competidores:[],contactoNuevo:false,moneda:'MXP',
  campos:{},
  // Piloto de Actualizar Oportunidad: combobox de clientes y oportunidades.
  asesorSAP:null,                      // se resuelve en auth.js desde EMAIL_A_ASESOR
  oportunidades:[],                    // poblado async tras login
  opsLoading:true,                     // true hasta que termine la primera carga
  clienteCommit:'',                    // cliente "settled" (post-commit del combobox)
  oportunidadSeleccionada:null,        // objeto opp completo cuando el asesor elige una
  // Check-in: autocomplete sobre Clientes Activos.xlsx (~2800). Carga async.
  clientesActivos:[],                  // poblado async tras login (auth.js)
  clientesLoading:true,                // true hasta que termine la primera carga
  // Draft del check-in: cliente seleccionado/tipeado ANTES de tap "Iniciar visita".
  // El draft NO toca GEO.checkin — eso solo lo hace checkin() tras capturar GPS.
  clienteCheckinDraft:null,            // { nombre, cardCode, ciudad, estatus, esNuevo }
  // True cuando el asesor entró a screen-checkin desde el botón [Cambiar] del form
  // (visita en curso) en vez de desde el selector inicial. Diferencia los dos flows:
  // false → capturar GPS; true → solo actualizar GEO.checkin sin re-GPS.
  cambiandoCliente:false,
  // True cuando screen-checkin se acaba de abrir (entry points). renderCheckin lo
  // consume una sola vez para hacer auto-focus + auto-open del dropdown, evitando
  // que re-renders disparados por cbCommitClienteCheckin reabran la lista (causaba
  // parpadeo en desktop y síntoma de "dropdown no cierra al seleccionar").
  checkinAcabaDeAbrirse:false
};

function guardarCampos(){
  // 'cliente-nombre' ya no existe en el form (heredado del check-in). Solo persisten
  // los inputs que el form sí renderiza, para sobrevivir re-renders por toggles.
  ['cliente-contacto','producto','monto','cierre','presupuesto','opp-nombre','notas','c-nombre','c-puesto','c-tel','c-email','monto-final','fecha-cierre-real','razon-perdida-detalle'].forEach(id=>{
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
  STATE.marca=null;STATE.etapa=null;STATE.acompanantes=[];
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

// Hook llamado por auth.js cuando termina la carga async de Clientes Activos.
// Si la screen-checkin ya está abierta esperando datos, re-renderiza para que
// la lista pre-poblada aparezca sin necesidad de reabrir.
function onClientesActivosCargados(){
  STATE.clientesLoading = false;
  if(document.getElementById('screen-checkin')?.classList.contains('active')){
    renderCheckin();
  }
}

// ── Screen de Check-in ──────────────────────────────────────────────────────
// Reemplaza el prompt() nativo de check-in. Captura cliente vía autocomplete
// sobre Clientes Activos.xlsx con priorización por asesor + estatus comercial.

function renderCheckin(){
  const draft = STATE.clienteCheckinDraft;
  const cambiando = STATE.cambiandoCliente;
  const titulo = cambiando ? 'Cambiar cliente' : 'Check-in';

  if(STATE.clientesLoading){
    document.getElementById('screen-checkin').innerHTML = `
      <header class="top-bar">
        <button class="back-btn" onclick="cancelarCheckin()">← Atrás</button>
        <div style="color:white;font-size:14px;font-weight:500">${titulo}</div>
        <div style="width:50px"></div>
      </header>
      <div class="form-body">
        <div class="hint-piloto" style="margin-top:24px">Cargando lista de clientes activos de SAP...</div>
      </div>`;
    return;
  }

  // Init del combobox: ordena por prioridad una vez. Si ya estaba init, sobrescribe
  // (los items son los mismos array reference, el sort es estable).
  cbCliActivoInit(STATE.clientesActivos);

  const inputValue = draft ? draft.nombre : '';
  const nuevoBadge = draft && draft.esNuevo
    ? `<div class="hint-piloto" style="margin-top:8px;color:#A85F0A">Cliente nuevo — se guardará sin CardCode.</div>`
    : '';
  const previaCard = draft && !draft.esNuevo
    ? `<div class="opp-banner" style="margin-top:8px">
        <strong>${_cbEsc(draft.nombre)}</strong>
        ${draft.ciudad ? '· '+_cbEsc(draft.ciudad) : ''}
        ${draft.estatus ? '· '+_cbEsc(draft.estatus) : ''}
      </div>`
    : '';

  const btnLabel = cambiando ? 'Confirmar cambio' : 'Iniciar visita';
  const btnDisabled = !draft ? 'disabled' : '';
  const btnHelp = !draft
    ? `<div class="hint-piloto" style="text-align:center;margin-top:8px">Selecciona un cliente o escribe el nombre para continuar.</div>`
    : '';

  document.getElementById('screen-checkin').innerHTML = `
    <header class="top-bar">
      <button class="back-btn" onclick="cancelarCheckin()">← Atrás</button>
      <div style="color:white;font-size:14px;font-weight:500">${titulo}</div>
      <div style="width:50px"></div>
    </header>
    <div class="form-body">
      <div class="card">
        <div class="card-title">¿Con qué cliente estás?</div>
        <div class="combobox" id="${CB_CLI_BOX}">
          <input type="text" id="${CB_CLI_BOX}-input"
                 value="${_cbAttr(inputValue)}"
                 placeholder="Buscar cliente..."
                 oninput="cbCliActivoFilter(this.value)"
                 onfocus="cbCliActivoOpen()"
                 onblur="cbCliActivoBlur()"
                 onkeydown="cbCliActivoKeydown(event)"
                 autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
          <div class="combobox-list" id="${CB_CLI_BOX}-list" hidden
               onmousedown="event.preventDefault()"></div>
        </div>
        ${previaCard}
        ${nuevoBadge}
      </div>
      <button class="save-btn" onclick="confirmarCheckin()" ${btnDisabled}>${btnLabel}</button>
      ${btnHelp}
    </div>`;

  // Auto-focus + lista pre-poblada SOLO en el mount inicial de screen-checkin.
  // Los re-renders disparados por cbCommitClienteCheckin (tras seleccionar opción)
  // NO deben reabrir la lista — eso causaba parpadeo en desktop y el síntoma de
  // "dropdown no cierra al seleccionar". El flag se setea en los 3 entry points
  // (seleccionarModo/mostrarModalCheckin/abrirCambiarCliente) y se consume aquí.
  if(STATE.checkinAcabaDeAbrirse){
    STATE.checkinAcabaDeAbrirse = false;
    setTimeout(() => {
      const input = _cbInput(CB_CLI_BOX);
      if(input){
        input.focus();
        cbCliActivoOpen();
      }
    }, 100);
  }
}

async function confirmarCheckin(){
  const draft = STATE.clienteCheckinDraft;
  if(!draft){ alert('Selecciona un cliente o escribe su nombre.'); return; }

  if(STATE.cambiandoCliente){
    // Flujo [Cambiar]: actualiza GEO.checkin sin re-capturar GPS.
    cambiarClienteCheckin(draft);
    STATE.cambiandoCliente = false;
    STATE.clienteCheckinDraft = null;
    mostrarScreen('screen-form');
    return;
  }

  // Flujo normal: switch a home ANTES de GPS para que el prompt de permisos del
  // browser se vea sobre home (no sobre screen-checkin), luego captura GPS.
  STATE.modo = 'campo';
  renderHome();
  mostrarScreen('screen-home');
  const result = await checkin(draft);
  if(!result){
    // GPS denegado: revertir al selector.
    STATE.modo = null;
    STATE.clienteCheckinDraft = null;
    renderHome();
    return;
  }
  STATE.clienteCheckinDraft = null;
  renderHome();
}

function cancelarCheckin(){
  if(STATE.cambiandoCliente){
    // [Cambiar] cancelado → vuelve al form con el cliente original intacto.
    STATE.cambiandoCliente = false;
    STATE.clienteCheckinDraft = null;
    mostrarScreen('screen-form');
    return;
  }
  // Check-in inicial cancelado → vuelve al selector de modo.
  STATE.modo = null;
  STATE.clienteCheckinDraft = null;
  renderHome();
  mostrarScreen('screen-home');
}

// Botón [Cambiar] desde el form: re-abre screen-checkin sin re-capturar GPS.
function abrirCambiarCliente(){
  guardarCampos();
  STATE.cambiandoCliente = true;
  // Pre-poblar el draft con el cliente actual del check-in para que el botón
  // "Confirmar cambio" funcione si el asesor no toca nada (idempotente).
  if(GEO.checkin){
    STATE.clienteCheckinDraft = {
      nombre: GEO.checkin.cliente,
      cardCode: GEO.checkin.cardCode,
      ciudad: GEO.checkin.ciudad || '',
      estatus: GEO.checkin.estatus || '',
      esNuevo: !!GEO.checkin.esNuevo
    };
  }
  STATE.checkinAcabaDeAbrirse = true;
  mostrarScreen('screen-checkin');
  renderCheckin();
}

// Click fuera de cualquier combobox abierto → cerrarlo. Necesario en mobile
// donde el blur nativo del input no siempre se dispara al tapear elementos
// no interactivos (cards, fondo, etc.) y la lista se queda abierta.
// Maneja ambos pools (cb-cliente/cb-opp del flujo Opp y cb-cli-activo del check-in).
document.addEventListener('click', (e) => {
  if(e.target.closest && e.target.closest('.combobox')) return;
  Object.keys(CB_STATE).forEach(boxId => {
    if(CB_STATE[boxId] && CB_STATE[boxId].open){
      if(CB_STATE[boxId].kind === 'clienteActivo') cbCliActivoClose(boxId);
      else cbClose(boxId);
    }
  });
});

// ── Combobox de Clientes Activos (check-in) ─────────────────────────────────
// Duplicación deliberada del combobox de Actualizar Oportunidad: comportamientos
// parecidos pero NO idénticos (priorización por asesor+estatus, dos líneas con
// subtext, fallback "cliente nuevo" sin CardCode, cap de 30 en lista inicial).
// Mantener funciones paralelas evita riesgo de regresión en el flujo Opp ya
// validado con Kimberly. Helpers genéricos (_cbEsc, _cbAttr, normCliente, CB_STATE)
// sí se reutilizan.

const CB_CLI_BOX = 'cb-cli-activo';
const CB_CLI_CAP = 30;

// Orden alfabético puro por Cliente. La priorización por asesor + estatus está
// deferida hasta que los datos de asignación en SAP sean confiables (~32% de
// clientes están bajo MOSTRADOR o sin asesor, muchas asignaciones no reflejan
// operación real). Ver pending work en CLAUDE.md.
function _cbCliActivoOrdenar(items){
  return [...items].sort((a, b) =>
    a.Cliente.localeCompare(b.Cliente, 'es', {sensitivity:'base'})
  );
}

// Subtext de cada opción: "Asesor · Ciudad" (la ciudad NO entra al filtro, es
// solo info para distinguir clientes con nombres similares).
function _cbCliActivoSubtext(c){
  const partes = [];
  if(c.Asesor) partes.push(c.Asesor);
  if(c.Ciudad) partes.push(c.Ciudad);
  return partes.join(' · ');
}

function cbCliActivoInit(items){
  const ordenados = _cbCliActivoOrdenar(items);
  CB_STATE[CB_CLI_BOX] = {
    items: ordenados,    // ya ordenados por prioridad — el orden se mantiene en filtros
    filter: '',
    highlighted: -1,
    open: false,
    kind: 'clienteActivo'
  };
}

function cbCliActivoOpen(){
  const cb = CB_STATE[CB_CLI_BOX]; if(!cb) return;
  cb.open = true;
  // NO reset filter aquí: el input puede traer texto residual (al regresar de
  // [Cambiar] o si el asesor reabre). cbCliActivoFilter es el path canónico.
  cb.highlighted = -1;
  cbCliActivoRenderList();
}

function cbCliActivoClose(){
  const cb = CB_STATE[CB_CLI_BOX]; if(!cb) return;
  cb.open = false;
  const list = _cbList(CB_CLI_BOX); if(list) list.hidden = true;
}

function cbCliActivoFilter(value){
  const cb = CB_STATE[CB_CLI_BOX]; if(!cb) return;
  cb.filter = value; cb.highlighted = -1; cb.open = true;
  cbCliActivoRenderList();
}

function cbCliActivoRenderList(){
  const cb = CB_STATE[CB_CLI_BOX]; const list = _cbList(CB_CLI_BOX);
  if(!cb || !list) return;
  const f = normCliente(cb.filter);
  // Filtrado SOLO contra Cliente (decisión explícita: la ciudad se ve en subtext
  // pero no es parámetro de búsqueda — evita ambigüedad y resultados anchos).
  let matches, hint = '';
  if(f){
    matches = cb.items.filter(c => normCliente(c.Cliente).includes(f));
  } else {
    // Sin filtro: top CB_CLI_CAP del orden priorizado. Si hay más, hint al fondo.
    matches = cb.items.slice(0, CB_CLI_CAP);
    if(cb.items.length > CB_CLI_CAP){
      hint = `Sigue escribiendo para buscar entre ${cb.items.length.toLocaleString('es-MX')} clientes`;
    }
  }
  cb._lastMatches = matches;

  if(matches.length === 0){
    list.innerHTML = `<div class="cb-empty">Sin coincidencias. Si es un prospecto nuevo, sigue escribiendo y al confirmar se guardará sin CardCode.</div>`;
  } else {
    const opts = matches.map((m, i) => {
      const hl = i === cb.highlighted ? ' cb-hl' : '';
      const sub = _cbCliActivoSubtext(m);
      const subHtml = sub ? `<div class="cb-option-sub">${_cbEsc(sub)}</div>` : '';
      return `<div class="cb-option${hl}" data-v="${_cbAttr(m.CardCode)}"
                   onmousedown="event.preventDefault(); cbCliActivoSelect(this.dataset.v)">
        <div class="cb-option-label">${_cbEsc(m.Cliente)}</div>${subHtml}
      </div>`;
    }).join('');
    const hintHtml = hint ? `<div class="cb-empty" style="font-size:11px;opacity:.7">${_cbEsc(hint)}</div>` : '';
    list.innerHTML = opts + hintHtml;
  }
  list.hidden = !cb.open;
  list.scrollTop = 0;
  const hlEl = list.querySelector('.cb-hl');
  if(hlEl) hlEl.scrollIntoView({ block:'nearest' });
}

function cbCliActivoSelect(cardCode){
  const cb = CB_STATE[CB_CLI_BOX]; if(!cb) return;
  const cli = cb.items.find(c => c.CardCode === cardCode);
  if(!cli) return;
  const input = _cbInput(CB_CLI_BOX);
  if(input) input.value = cli.Cliente;
  cbCliActivoClose();
  // El draft (clienteCheckinDraft) usa shape lowercase {nombre, cardCode, ...}
  // porque NO viene del xlsx — es estado interno consumido por geo.js (checkin/
  // cambiarClienteCheckin). Solo STATE.clientesActivos mantiene capitalización xlsx.
  cbCommitClienteCheckin({
    nombre: cli.Cliente,
    cardCode: cli.CardCode,
    ciudad: cli.Ciudad,
    estatus: cli.EstatusComercial,
    esNuevo: false
  });
}

function cbCliActivoBlur(){
  // Delay corto: tolera variaciones de orden de evento entre browsers móviles.
  setTimeout(() => {
    const cb = CB_STATE[CB_CLI_BOX]; if(!cb) return;
    if(cb.open) cbCliActivoClose();
    const input = _cbInput(CB_CLI_BOX);
    const typed = input ? input.value.trim() : '';
    if(!typed){
      // Vaciar draft si el asesor borró todo y se salió.
      STATE.clienteCheckinDraft = null;
      renderCheckin();
      return;
    }
    // Match exacto normalizado → tratar como selección formal del cliente real.
    const exact = cb.items.find(c => normCliente(c.Cliente) === normCliente(typed));
    if(exact){
      if(input) input.value = exact.Cliente;
      cbCommitClienteCheckin({
        nombre: exact.Cliente,
        cardCode: exact.CardCode,
        ciudad: exact.Ciudad,
        estatus: exact.EstatusComercial,
        esNuevo: false
      });
    } else {
      // Sin match → prospecto/cliente nuevo. CardCode null; el form mostrará el aviso.
      cbCommitClienteCheckin({
        nombre: typed,
        cardCode: null,
        ciudad: '',
        estatus: '',
        esNuevo: true
      });
    }
  }, 150);
}

function cbCliActivoKeydown(event){
  const cb = CB_STATE[CB_CLI_BOX]; if(!cb) return;
  if(event.key === 'ArrowDown'){
    event.preventDefault();
    if(!cb.open){ cbCliActivoOpen(); return; }
    const m = cb._lastMatches || cb.items;
    cb.highlighted = Math.min(cb.highlighted + 1, m.length - 1);
    cbCliActivoRenderList();
  } else if(event.key === 'ArrowUp'){
    event.preventDefault();
    if(!cb.open) return;
    cb.highlighted = Math.max(cb.highlighted - 1, 0);
    cbCliActivoRenderList();
  } else if(event.key === 'Enter'){
    if(!cb.open) return;
    event.preventDefault();
    const m = cb._lastMatches || cb.items;
    if(cb.highlighted >= 0 && m[cb.highlighted]) cbCliActivoSelect(m[cb.highlighted].CardCode);
    else if(m.length === 1) cbCliActivoSelect(m[0].CardCode);
    else { const inp = _cbInput(CB_CLI_BOX); if(inp) inp.blur(); }
  } else if(event.key === 'Escape'){
    cbCliActivoClose();
  }
}

// Commit del combobox: setea draft, re-renderiza screen-checkin (que actualiza
// el bloque de confirmación visible bajo el input). NO toca GEO.checkin todavía
// — eso espera al tap de "Iniciar visita" / "Confirmar cambio".
function cbCommitClienteCheckin(clienteInfo){
  STATE.clienteCheckinDraft = clienteInfo;
  renderCheckin();
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
    // Sin check-in: abrir screen-checkin con autocomplete sobre Clientes Activos.
    STATE.cambiandoCliente = false;
    STATE.clienteCheckinDraft = null;
    STATE.checkinAcabaDeAbrirse = true;
    mostrarScreen('screen-checkin');
    renderCheckin();
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

  // Acompañantes (líderes / ingenieros / gerencia-dirección). Lista plana con
  // badge de rol; reuso las clases .lider-item/.lider-name/.lider-lineas
  // (presentación genérica, no implica que todos sean líderes).
  const acompanantesHTML=CONFIG.acompanantes.map(a=>{
    const sel=STATE.acompanantes.includes(a.email);
    const scope = a.linea || a.division || '';
    const subtext = scope ? `${a.rol} · ${scope}` : a.rol;
    const reportaLine = a.reportaA
      ? `<div style="font-size:11px;color:var(--color-text-tertiary,#888);margin-top:1px">reporta a ${a.reportaA}</div>`
      : '';
    return `<div class="lider-item" onclick="toggleAcompanante('${a.email}')">
      <div class="avatar" style="${sel?'background:#EAF3DE;color:#0F6E56':''}">${iniciales(a.nombre)}</div>
      <div>
        <div class="lider-name">${a.nombre}</div>
        <div class="lider-lineas">${subtext}</div>
        ${reportaLine}
      </div>
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

  // Acompañantes: visible para Visita y Demo. Opcional en ambos casos.
  // Incluye líderes de línea, ingenieros de aplicación, gerencia y dirección.
  const liderSection=(STATE.tipo==='demo'||STATE.tipo==='visita')?`<div class="card">
    <div class="card-title">Acompañantes en la visita</div>
    <div class="search-box">
      <span class="search-icon">&#9906;</span>
      <input type="text" placeholder="Buscar por nombre, rol o línea..." oninput="filtrarAcompanantes(this.value)">
    </div>
    <div id="acompanantes-list">${acompanantesHTML}</div>
    ${STATE.acompanantes.length>0?'<div class="nota-hint">Recibirán notificación para confirmar su asistencia</div>':''}
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

  // Cliente heredado del check-in (GEO.checkin). El form ya no captura el nombre
  // del cliente — solo lo muestra como contexto read-only con botón [Cambiar].
  const checkinInfo = (typeof GEO !== 'undefined' && GEO.checkin) ? GEO.checkin : {};
  const visitandoSubInfo = checkinInfo.esNuevo
    ? 'Cliente nuevo · sin CardCode'
    : [checkinInfo.ciudad, checkinInfo.estatus].filter(Boolean).join(' · ');
  const visitandoCard = `<div class="card">
    <div class="card-title-row">
      <div>
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--color-text-tertiary);margin-bottom:4px">Visitando</div>
        <div style="font-size:15px;font-weight:500">${_cbEsc(checkinInfo.cliente || '(sin cliente)')}</div>
        ${visitandoSubInfo ? `<div style="font-size:12px;color:var(--color-text-secondary);margin-top:2px">${_cbEsc(visitandoSubInfo)}</div>` : ''}
      </div>
      <button type="button" class="btn-refresh" onclick="abrirCambiarCliente()" title="Cambiar cliente" style="font-size:12px;padding:6px 10px;width:auto">Cambiar</button>
    </div>
  </div>`;

  const contactoPrincipalCard = `<div class="card">
    <div class="card-title">Contacto principal <span class="opt">opc</span></div>
    <input type="text" id="cliente-contacto" placeholder="Nombre del contacto que atendió">${oppNombreEnCliente}
  </div>`;

  $('screen-form').innerHTML=`
    <header class="top-bar ${t.color}">
      ${backBtn}
      <div style="color:white;font-size:14px;font-weight:500">${t.label}</div>
      <div style="width:50px"></div>
    </header>
    <div class="form-body">
      ${visitandoCard}
      ${contactoPrincipalCard}
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
function toggleAcompanante(email){guardarCampos();const i=STATE.acompanantes.indexOf(email);if(i>-1)STATE.acompanantes.splice(i,1);else STATE.acompanantes.push(email);renderForm()}
function toggleContacto(){guardarCampos();STATE.contactoNuevo=!STATE.contactoNuevo;renderForm()}
function selResultadoOpp(resultado){STATE.resultado=resultado;renderForm()}
function volverSelectorOpp(){guardarCampos();STATE.resultado=null;renderForm()}
function selRazonPerdida(razon){guardarCampos();STATE.razonPerdida=razon;renderForm()}

function filtrarAcompanantes(q){
  const lista=$('acompanantes-list');if(!lista)return;
  // Filtra contra todo el textContent del item: nombre + rol + línea + "reporta a X".
  // Más útil que filtrar solo por nombre — permite "ingeniero", "visión", etc.
  const needle=q.toLowerCase();
  lista.querySelectorAll('.lider-item').forEach(item=>{
    const t=item.textContent.toLowerCase();
    item.style.display=t.includes(needle)?'':'none';
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

  // Cliente: sub-flujos leen del commit del combobox de Opp; visita/demo del check-in.
  const cliente = esSubflujoOpp
    ? (STATE.clienteCommit?.trim() || '')
    : (GEO.checkin?.cliente || '');
  if (!cliente) {
    alert(esSubflujoOpp ? 'Falta el nombre del cliente' : 'No hay un cliente activo. Haz check-in primero.');
    return;
  }

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

  // CardCode canónico: viene del check-in para visita/demo (cliente seleccionado
  // del autocomplete sobre Clientes Activos). null en fallback de texto libre y
  // en sub-flujos de oportunidad (estos identifican por OpportunidadID).
  const cardCodeFinal = esSubflujoOpp ? null : (GEO.checkin?.cardCode || null);

  const registro = {
    tipo: STATE.tipo,
    modo: STATE.modo || 'campo',
    fecha: new Date().toISOString(),
    cliente: clienteFinal,
    cardCode: cardCodeFinal,
    clienteContacto: esSubflujoOpp ? '' : STATE.campos['cliente-contacto'],
    clienteNuevo: esSubflujoOpp ? false : !!(GEO.checkin?.esNuevo),
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
    acompanantes: esSubflujoOpp ? [] : STATE.acompanantes,
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
// Wrapper para botones legacy (actualizarBotonCheckin lo llama cuando no hay
// check-in activo). Misma transición que seleccionarModo('campo').

function mostrarModalCheckin() {
  STATE.cambiandoCliente = false;
  STATE.clienteCheckinDraft = null;
  STATE.checkinAcabaDeAbrirse = true;
  mostrarScreen('screen-checkin');
  renderCheckin();
}

// ── Exponer funciones globales ───────────────────────────────────────────────

window.irA                  = irA;
window.mostrarScreen        = mostrarScreen;
window.selDivision          = selDivision;
window.selMarca             = selMarca;
window.selEtapa             = selEtapa;
window.selMoneda            = selMoneda;
window.toggleComp           = toggleComp;
window.toggleAcompanante    = toggleAcompanante;
window.toggleContacto       = toggleContacto;
window.filtrarAcompanantes  = filtrarAcompanantes;
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
window.cbCliActivoOpen      = cbCliActivoOpen;
window.cbCliActivoClose     = cbCliActivoClose;
window.cbCliActivoFilter    = cbCliActivoFilter;
window.cbCliActivoSelect    = cbCliActivoSelect;
window.cbCliActivoBlur      = cbCliActivoBlur;
window.cbCliActivoKeydown   = cbCliActivoKeydown;
window.confirmarCheckin     = confirmarCheckin;
window.cancelarCheckin      = cancelarCheckin;
window.abrirCambiarCliente  = abrirCambiarCliente;
window.onClientesActivosCargados = onClientesActivosCargados;

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
