/* ============================================================
   TOOLBOX CNA  —  app.js
   SPA completa, zero dipendenze esterne
============================================================ */

// ── Icone per regione ─────────────────────────────────────────
const ICONS = {
  Generali:        '🌐',
  Sicilia:         '🏝️',
  Lombardia:       '🏙️',
  'Emilia-Romagna':'🍝',
  Segreteria:      '📋',
  Amministrazione: '📊',
  _default:        '📍',
};
const ico = r => ICONS[r] || ICONS._default;

// ── Stato applicazione ────────────────────────────────────────
const ST = {
  tools:    [],   // lista completa dal server
  uid:      null, // tool selezionato
  tab:      'Tutti', // tab attiva in home
  q:        '',   // ricerca
  collapsed: {}, // regioni collassate in sidebar
};

// ── DOM shortcuts ─────────────────────────────────────────────
const g  = id  => document.getElementById(id);
const qs = sel => document.querySelector(sel);

// ── Escape HTML ───────────────────────────────────────────────
function h(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

// ── API ───────────────────────────────────────────────────────
async function GET(path) {
  const r = await fetch('/api' + path);
  if (!r.ok) { const t = await r.text(); throw new Error(t); }
  return r.json();
}
async function POST(path, body) {
  const r = await fetch('/api' + path, {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(body),
  });
  if (!r.ok) { const t = await r.text(); throw new Error(t); }
  return r.json();
}

// ── Avvio ─────────────────────────────────────────────────────
async function boot() {
  try {
    ST.tools = await GET('/tools');
  } catch(e) {
    toast('Errore caricamento tool: ' + e.message, 'err');
    ST.tools = [];
  }
  renderSidebar();
  showHome();

  // Carica config
  try {
    const cfg = await GET('/config');
    if (cfg.ai) {
      g('cfg-model').value = cfg.ai.model_id || '';
      g('cfg-url').value   = cfg.ai.base_url  || '';
      g('cfg-key').value   = cfg.ai.api_key   || '';
    }
  } catch(_) {}
}

// ═══════════════════════════════════════════════════════════════
//  SIDEBAR
// ═══════════════════════════════════════════════════════════════
function renderSidebar() {
  const q = ST.q.toLowerCase();

  const visible = q
    ? ST.tools.filter(t =>
        (t.name||'').toLowerCase().includes(q) ||
        (t.region||'').toLowerCase().includes(q))
    : ST.tools;

  const okCount = ST.tools.filter(t => !t.import_error).length;

  // Raggruppa per regione
  const byR = {};
  for (const t of visible) {
    const r = t.region || 'Generali';
    (byR[r] = byR[r] || []).push(t);
  }
  const regions = Object.keys(byR).sort((a,b) => {
    if (a==='Generali') return -1;
    if (b==='Generali') return 1;
    return a.localeCompare(b,'it');
  });

  let out = '';

  // Dashboard link
  out += `
    <div class="s-label">MAIN MENU</div>
    <button class="s-item${ST.uid === null ? ' on' : ''}" onclick="showHome()">
      <div class="s-item-ico">🏠</div>
      <span class="s-item-lbl">Dashboard</span>
      <span class="s-item-num">${okCount}</span>
    </button>`;

  // Tool per regione
  for (const reg of regions) {
    if (ST.collapsed[reg] === undefined) ST.collapsed[reg] = true;
    const hasActive = byR[reg].some(t => t.uid === ST.uid);
    const open = q ? true : (!ST.collapsed[reg] || hasActive);
    const regKey = reg.replace(/[^a-zA-Z0-9]/g, '_');
    const count  = byR[reg].length;

    out += `
      <div class="s-label s-label-toggle" onclick="toggleSbRegion('${h(reg)}','sr-${regKey}')">
        <span class="s-chev${open?' open':''}" id="sc-${regKey}"></span>
        <span class="s-label-toggle-txt">${h(reg.toUpperCase())}</span>
        <span class="s-region-count">${count}</span>
      </div>
      <div class="s-region-body${open?' open':''}" id="sr-${regKey}">
        <div class="s-region-inner">`;
    for (const t of byR[reg]) {
      const on = t.uid === ST.uid;
      out += `
          <button class="s-item${on?' on':''}${t.import_error?' opacity-60':''}"
            onclick="openTool('${h(t.uid)}')" title="${h(t.name)}">
            <div class="s-item-ico">${ico(t.region)}</div>
            <span class="s-item-lbl">${h(t.name)}</span>
          </button>`;
    }
    out += `</div></div>`;
  }

  if (regions.length === 0) {
    out += `<div style="padding:12px 8px;font-size:12.5px;color:var(--t-mute);">Nessun risultato.</div>`;
  }


  g('sb-nav').innerHTML = out;
}

// ═══════════════════════════════════════════════════════════════
//  TOPBAR
// ═══════════════════════════════════════════════════════════════
function setTopbar(title, sub) {
  g('tb-area').innerHTML = `<h2>${h(title)}</h2><p>${h(sub)}</p>`;
}

// ═══════════════════════════════════════════════════════════════
//  HOME
// ═══════════════════════════════════════════════════════════════
function showHome() {
  ST.uid = null;
  renderSidebar();
  g('rpanel').classList.remove('show');
  setTopbar('Benvenuto!', 'Seleziona un tool per iniziare.');

  const ok   = ST.tools.filter(t => !t.import_error);
  const bad  = ST.tools.filter(t =>  t.import_error);
  const regs = [...new Set(ST.tools.map(t => t.region || 'Generali'))];

  // Tab lista
  const allRegions = ['Tutti', ...regs.sort((a,b)=>{
    if (a==='Generali') return -1;
    if (b==='Generali') return 1;
    return a.localeCompare(b,'it');
  })];
  if (!allRegions.includes(ST.tab)) ST.tab = 'Tutti';

  // Filtra per tab
  const shown = ST.tab === 'Tutti'
    ? ok
    : ok.filter(t => (t.region||'Generali') === ST.tab);

  // Tabs HTML
  const tabsHtml = allRegions.map(r =>
    `<button class="tab${ST.tab===r?' on':''}" onclick="setTab('${h(r)}')">${h(r)}</button>`
  ).join('') +
  ``;

  // Righe tabella
  const rows = shown.length === 0
    ? `<tr><td colspan="5" style="text-align:center;padding:30px 16px;color:var(--t-mute);font-size:13px;">
         Nessun tool in questa categoria.
       </td></tr>`
    : shown.map(t => {
        const params = (t.params||[]).filter(p => !['info','warning','error','success','markdown','dynamic_info','file_path_info'].includes(p.type)).length;
        const files  = (t.inputs||[]).filter(i => !['info','warning','error','success','markdown'].includes(i.type)).length;
        return `
          <tr onclick="openTool('${h(t.uid)}')">
            <td class="name-cell">
              <strong>${ico(t.region)} ${h(t.name)}</strong>
              <span>${h(t.uid)}</span>
            </td>
            <td><span class="badge">${h(t.region||'Generali')}</span></td>
            <td><div class="status-cell"><span class="dot-ok">●</span> Attivo</div></td>
            <td style="font-size:12px;color:var(--t-mute);">
              ${files>0  ? `📁 ${files} file` : ''}
              ${params>0 ? `&thinsp;⚙️ ${params} param` : ''}
              ${files===0 && params===0 ? '—' : ''}
            </td>
            <td>
              <button class="btn-open pri"
                onclick="event.stopPropagation();openTool('${h(t.uid)}')">Apri →</button>
            </td>
          </tr>`;
      }).join('');

  // Righe errori
  const errRows = bad.map(t => `
    <tr style="opacity:.55;cursor:default;">
      <td class="name-cell">
        <strong>❌ ${h(t.name)}</strong>
        <span>${h(t.uid)}</span>
      </td>
      <td><span class="badge">${h(t.region||'')}</span></td>
      <td><div class="status-cell"><span class="dot-err">●</span> <span style="color:var(--red)">Errore</span></div></td>
      <td>—</td>
      <td><span style="font-size:12px;color:var(--t-mute);">Non disponibile</span></td>
    </tr>`).join('');

  g('content').innerHTML = `
    <div class="stats-row">
      <div class="stat-card">
        <div class="stat-ico">🧰</div>
        <div class="stat-num">${pad(ok.length)}</div>
        <div class="stat-lbl">Tool Attivi</div>
      </div>
      <div class="stat-card">
        <div class="stat-ico">📍</div>
        <div class="stat-num">${pad(regs.length)}</div>
        <div class="stat-lbl">Regioni</div>
      </div>
      <div class="stat-card">
        <div class="stat-ico">⚡</div>
        <div class="stat-num">${pad(ST.tools.length)}</div>
        <div class="stat-lbl">Tool Totali</div>
      </div>
    </div>

    <div class="sec-row">
      <div>
        <span class="sec-title">Lista Tool</span>
        <span class="sec-count">( ${shown.length} ${ST.tab==='Tutti'?'disponibili':'in '+ST.tab} )</span>
      </div>
    </div>

    <div class="tabs">${tabsHtml}</div>

    <div class="tbl-card">
      <table class="tbl">
        <thead>
          <tr>
            <th>Nome Tool</th>
            <th>Regione</th>
            <th>Stato</th>
            <th>Parametri</th>
            <th>Azione</th>
          </tr>
        </thead>
        <tbody>${rows}${errRows}</tbody>
      </table>
    </div>
  `;
}

function setTab(r) { ST.tab = r; showHome(); }
const pad = n => n < 10 ? '0' + n : String(n);

// ═══════════════════════════════════════════════════════════════
//  TOOL DETAIL
// ═══════════════════════════════════════════════════════════════
function openTool(uid) {
  ST.uid = uid;
  const t = ST.tools.find(x => x.uid === uid);
  if (!t) return;

  renderSidebar();
  setTopbar(`${ico(t.region)} ${t.name}`, `${t.region||'Generali'}  ·  ${(t.inputs||[]).length} input, ${(t.params||[]).length} parametri`);

  buildDetail(t);
  buildRightPanel(t);
}

function buildDetail(t) {
  const isOk = !t.import_error;

  g('content').innerHTML = `
    ${t.description ? `
    <div class="card card-collapsible" id="card-desc">
      <div class="card-head card-toggle" onclick="toggleCard('card-desc')">
        <span class="card-head-lbl">📋 Descrizione</span>
        <span class="card-chevron"></span>
      </div>
      <div class="card-body-collapsible">
        <div class="card-body-inner"><div class="desc-md">${md(t.description)}</div></div>
      </div>
    </div>` : ''}

    ${isOk ? `
    <div class="card">
      <div class="card-head"><span class="card-head-lbl">⚙️ Esegui</span></div>
      <div class="card-body">
        <form id="run-form" onsubmit="return false">
          ${buildInputFields(t)}
          ${buildParamFields(t)}
          <button type="button" id="run-btn" onclick="runTool('${h(uid)}')">
            <div class="run-spin"></div>
            <span class="run-label">▶  Esegui Tool</span>
          </button>
          <div id="result"></div>
        </form>
      </div>
    </div>
    ` : `<div class="alert alert-error">❌ Tool non disponibile — errore di import. Controlla il codice del tool.</div>`}
  `;

  hookDropzones(t);
  fetchDynamicInfos(t);
}

// shortcut per uid dal tool corrente nel click handler
const uid = t => t.uid;

function buildRightPanel(t) {
  const rp = g('rpanel');
  const inp = (t.inputs||[]).filter(i => !['info','warning','error','success','markdown'].includes(i.type));
  const par = (t.params||[]).filter(p => !['info','warning','error','success','markdown','dynamic_info'].includes(p.type));

  // Leggi config AI corrente
  const model = g('cfg-model') ? g('cfg-model').value.trim() : '';
  const url   = g('cfg-url')   ? g('cfg-url').value.trim()   : '';
  const aiOk  = !!(model && url);

  rp.innerHTML = `
    <div class="rp-tabs">
      <button class="rp-tab"    id="rpt-info" onclick="switchRpTab('info')">ℹ️ Dettagli</button>
      <button class="rp-tab on" id="rpt-ai"   onclick="switchRpTab('ai')">🤖 AI</button>
    </div>

    <!-- PANE: Dettagli Tool -->
    <div class="rp-pane" id="rpp-info">
      <div class="rp-sec">
        <div class="rp-sec-lbl">Identificativo</div>
        <div class="rp-row"><span class="rp-key">UID</span><span class="rp-val">${h(t.uid)}</span></div>
        <div class="rp-row"><span class="rp-key">ID</span><span class="rp-val">${h(t.id)}</span></div>
        <div class="rp-row"><span class="rp-key">Regione</span><span class="rp-val" style="font-family:var(--font)">${h(t.region||'Generali')}</span></div>
        <div class="rp-row">
          <span class="rp-key">Stato</span>
          <span style="font-size:12px;font-weight:700;color:${t.import_error?'var(--red)':'var(--green)'}">
            ${t.import_error ? '❌ Errore' : '✅ Attivo'}
          </span>
        </div>
      </div>

      <div class="rp-sec">
        <div class="rp-sec-lbl">File di input (${inp.length})</div>
        ${inp.length===0
          ? `<div style="font-size:12.5px;color:var(--t-mute)">Nessun file richiesto</div>`
          : inp.map(i=>`<div class="rp-row"><span class="rp-key">${h(i.label||i.key)}</span><span class="rp-val">${h(i.type)}</span></div>`).join('')}
      </div>

      <div class="rp-sec">
        <div class="rp-sec-lbl">Parametri (${par.length})</div>
        ${par.length===0
          ? `<div style="font-size:12.5px;color:var(--t-mute)">Nessun parametro</div>`
          : par.map(p=>`<div class="rp-row"><span class="rp-key">${h(p.label||p.key)}</span><span class="rp-val">${h(p.type)}</span></div>`).join('')}
      </div>

      <hr class="divider"/>
      <button class="rp-back" onclick="showHome()">← Torna alla lista</button>
    </div>

    <!-- PANE: AI Chat -->
    <div class="rp-pane on" id="rpp-ai">
      ${buildAIPaneHTML(aiOk)}
    </div>
  `;
  rp.classList.add('show');
  renderChatHistory();
}

function openAIPanel() {
  const rp    = g('rpanel');
  const model = g('cfg-model') ? g('cfg-model').value.trim() : '';
  const url   = g('cfg-url')   ? g('cfg-url').value.trim()   : '';
  const aiOk  = !!(model && url);

  rp.innerHTML = `
    <div class="rp-tabs">
      <button class="rp-tab on" id="rpt-ai">🤖 Chat AI</button>
    </div>
    <div class="rp-pane on" id="rpp-ai">
      ${buildAIPaneHTML(aiOk)}
    </div>
  `;
  rp.classList.add('show');
  renderChatHistory();
  const inp = g('chat-inp');
  if (inp) inp.focus();
}

function buildAIPaneHTML(aiOk) {
  if (!aiOk) return `
    <div class="ai-uncfg">
      <div class="ai-uncfg-ico">🤖</div>
      <div class="ai-uncfg-txt">AI non configurata</div>
      <div class="ai-uncfg-sub">Imposta Model ID, URL e API Key per abilitare la chat.</div>
      <button class="ai-edit-btn" onclick="openSettings()">⚙️ Configura AI</button>
    </div>`;
  return `
    <div class="chat-wrap">
      <div class="chat-msgs" id="chat-msgs"></div>
      <div class="chat-bar">
        <textarea class="chat-inp" id="chat-inp" rows="1"
          placeholder="Scrivi un messaggio…"
          onkeydown="chatKey(event)"></textarea>
        <button class="chat-send" id="chat-send" onclick="chatSend()">➤</button>
      </div>
    </div>`;
}

function renderChatHistory() {
  const box = g('chat-msgs');
  if (!box) return;
  box.innerHTML = '';
  for (const m of CHAT_HISTORY) {
    const div = document.createElement('div');
    div.className = `cmsg cmsg-${m.role}`;
    if (m.role === 'assistant') div.innerHTML = md(m.content);
    else div.textContent = m.content;
    box.appendChild(div);
  }
  box.scrollTop = box.scrollHeight;
}

function toggleSbRegion(reg, bodyId) {
  ST.collapsed[reg] = !ST.collapsed[reg];
  const open = !ST.collapsed[reg];
  const body = g(bodyId);
  const regKey = reg.replace(/[^a-zA-Z0-9]/g, '_');
  const chev  = g(`sc-${regKey}`);
  if (body) body.classList.toggle('open', open);
  if (chev) chev.classList.toggle('open', open);   // .s-chev
}

function toggleCard(id) {
  const card = g(id); if (!card) return;
  card.classList.toggle('open');
}

function switchRpTab(name) {
  ['info','ai'].forEach(n => {
    g(`rpt-${n}`).classList.toggle('on', n === name);
    g(`rpp-${n}`).classList.toggle('on', n === name);
  });
  if (name === 'ai') renderChatHistory();
}

// ═══════════════════════════════════════════════════════════════
//  FORM BUILDERS
// ═══════════════════════════════════════════════════════════════
function buildInputFields(t) {
  return (t.inputs||[]).map(inp => {
    const tp  = inp.type || 'file_single';
    const req = inp.required !== false;
    const lbl = inp.label || inp.key;

    // Alert e markdown: usa md() per renderizzare il testo
    if (['info','warning','error','success'].includes(tp))
      return `<div class="alert alert-${tp}">${md(lbl)}</div>`;
    if (tp === 'markdown')
      return `<div class="desc-md">${md(lbl)}</div>`;

    const multi  = ['txt_multi','file_multi'].includes(tp);
    const accept = acceptStr(tp);
    const hint   = hintStr(tp);

    return `
      <div class="fg" id="wrap-${h(inp.key)}">
        <label class="fl">${h(lbl)}${req?'<span class="req">*</span>':''}</label>
        <div class="dz" id="dz-${h(inp.key)}">
          <input type="file" id="inp-${h(inp.key)}" name="${h(inp.key)}"
            accept="${accept}" ${multi?'multiple':''} />
          <div class="dz-ico">📁</div>
          <div class="dz-title">${multi?'Trascina i file o clicca':'Clicca per scegliere'}</div>
          <div class="dz-hint">${hint}</div>
        </div>
        <div id="fp-${h(inp.key)}" class="fpills"></div>
        ${inp.description ? `<div class="fhint">${md(inp.description)}</div>` : ''}
      </div>`;
  }).join('');
}

function buildParamFields(t) {
  return (t.params||[]).map(p => {
    const tp  = p.type || 'text';
    const req = p.required !== false;
    const lbl = p.label || p.key;
    const pid = `p-${h(p.key)}`;

    // Alert e markdown: usa md()
    if (['info','warning','error','success'].includes(tp))
      return `<div class="alert alert-${tp}">${md(lbl)}</div>`;
    if (tp === 'markdown')
      return `<div class="desc-md">${md(lbl)}</div>`;

    // dynamic_info: sezione con loading spinner, contenuto caricato via API
    if (tp === 'dynamic_info') {
      const sectionLabel = p.section ? `<div class="dyninfo-section">${h(p.section)}</div>` : '';
      const headerLabel  = lbl && lbl !== p.key ? `<div class="dyninfo-label">${md(lbl)}</div>` : '';
      return `
        <div class="dyninfo-wrap" id="dynwrap-${h(p.key)}">
          ${sectionLabel}${headerLabel}
          <div class="dyninfo-body" id="dyn-${h(p.key)}">
            <div class="dyninfo-loading">
              <span class="dyninfo-spin"></span>
              <span>Caricamento info…</span>
            </div>
          </div>
        </div>`;
    }

    const labelEl = `<label class="fl" for="${pid}">${h(lbl)}${req?'<span class="req">*</span>':''}</label>`;
    const hint    = p.description ? `<div class="fhint">${md(p.description)}</div>` : '';

    if (tp === 'text')
      return `<div class="fg">${labelEl}
        <input class="fi" type="text" id="${pid}" name="${h(p.key)}"
          value="${h(p.default||'')}" placeholder="${h(p.placeholder||'')}"/>
        ${hint}</div>`;

    if (tp === 'textarea')
      return `<div class="fg">${labelEl}
        <textarea class="fi" id="${pid}" name="${h(p.key)}"
          placeholder="${h(p.placeholder||'')}">${h(p.default||'')}</textarea>
        ${hint}</div>`;

    if (tp === 'number') {
      const mn = p.min  != null ? `min="${p.min}"` : '';
      const mx = p.max  != null ? `max="${p.max}"` : '';
      const st = p.step != null ? `step="${p.step}"` : '';
      return `<div class="fg">${labelEl}
        <div class="num-row">
          <input type="number" id="${pid}" name="${h(p.key)}"
            value="${p.default != null ? p.default : ''}" ${mn} ${mx} ${st}/>
          <button type="button" class="num-btn" onclick="numSt('${pid}',-1,${p.step||1})">−</button>
          <button type="button" class="num-btn" onclick="numSt('${pid}', 1,${p.step||1})">+</button>
        </div>${hint}</div>`;
    }

    if (tp === 'checkbox') {
      const on = p.default === true;
      return `<div class="fg">
        <label class="chk-row" onclick="tglChk('chkb-${h(p.key)}','${pid}')">
          <div class="chk-box${on?' on':''}" id="chkb-${h(p.key)}">${on?'✓':''}</div>
          <input type="hidden" id="${pid}" name="${h(p.key)}" value="${on?'true':'false'}"/>
          <span style="font-size:13.5px;color:var(--t-body)">${h(lbl)}</span>
        </label>${hint}</div>`;
    }

    if (tp === 'radio') {
      const opts = p.options || [];
      const def  = p.default || opts[0] || '';
      return `<div class="fg">${labelEl}
        <div class="radio-g" id="rg-${h(p.key)}">
          ${opts.map(o=>`
            <div class="radio-p${o===def?' on':''}" data-v="${h(o)}"
              onclick="pickR('rg-${h(p.key)}','${pid}','${h(o)}')">${h(o)}</div>`).join('')}
        </div>
        <input type="hidden" id="${pid}" name="${h(p.key)}" value="${h(def)}"/>
        ${hint}</div>`;
    }

    if (tp === 'select') {
      const opts = p.options || [];
      return `<div class="fg">${labelEl}
        <select class="fi" id="${pid}" name="${h(p.key)}">
          ${opts.map(o=>`<option value="${h(o)}"${o===p.default?' selected':''}>${h(o)}</option>`).join('')}
        </select>${hint}</div>`;
    }

    if (tp === 'multiselect') {
      const opts = p.options || [];
      const defs = Array.isArray(p.default) ? p.default : [];
      return `<div class="fg">${labelEl}
        <div class="ms-g" id="ms-${h(p.key)}">
          ${opts.map(o=>`
            <div class="ms-c${defs.includes(o)?' on':''}" data-v="${h(o)}"
              onclick="tglMs('ms-${h(p.key)}','${h(o)}','mso-${h(p.key)}')">${h(o)}</div>`).join('')}
        </div>
        <input type="hidden" id="mso-${h(p.key)}" name="${h(p.key)}" value="${h(JSON.stringify(defs))}"/>
        ${hint}</div>`;
    }

    if (tp === 'folder')
      return `<div class="fg">${labelEl}
        <input class="fi" type="text" id="${pid}" name="${h(p.key)}"
          value="${h(p.default||'')}" placeholder="Percorso cartella…"/>
        ${hint}</div>`;

    if (tp === 'file_path_info') {
      const helpText = p.help ? `<div class="fhint">${md(p.help)}</div>` : hint;
      return `<div class="fg">${labelEl}
        <div class="filepath-wrap">
          <input class="fi filepath-inp" type="text" id="${pid}" name="${h(p.key)}"
            value="${h(p.default||'')}" placeholder="Percorso file…"/>
        </div>
        ${helpText}</div>`;
    }

    return '';
  }).join('');
}

// ═══════════════════════════════════════════════════════════════
//  DYNAMIC INFO — Fetch asincrono dal server
// ═══════════════════════════════════════════════════════════════
async function fetchDynamicInfos(t) {
  const dynParams = (t.params||[]).filter(p => p.type === 'dynamic_info');
  if (!dynParams.length) return;

  for (const p of dynParams) {
    const box = g(`dyn-${p.key}`);
    if (!box) continue;

    try {
      const url = `/api/tools/${encodeURIComponent(t.uid)}/dyninfo?key=${encodeURIComponent(p.key)}`;
      const resp = await fetch(url);
      if (!resp.ok) throw new Error(await resp.text());
      const data = await resp.json();

      const msgs = data.messages || [];

      if (msgs.length === 0) {
        // Nessun messaggio catturato: mostra placeholder discreto
        box.innerHTML = `<div class="dyninfo-empty">ℹ️ Nessuna informazione disponibile per questo parametro.</div>`;
      } else {
        box.innerHTML = msgs.map(m => {
          const tp = m.type === 'write' || m.type === 'markdown' ? 'info' : m.type;
          return `<div class="dyninfo-msg dyninfo-msg-${tp}">${md(m.text)}</div>`;
        }).join('');
      }
    } catch(e) {
      const box2 = g(`dyn-${p.key}`);
      if (box2) box2.innerHTML = `<div class="dyninfo-msg dyninfo-msg-error">⚠️ ${h(e.message)}</div>`;
    }
  }
}

// ═══════════════════════════════════════════════════════════════
//  FORM INTERACTIONS
// ═══════════════════════════════════════════════════════════════
function hookDropzones(t) {
  for (const inp of (t.inputs||[])) {
    const fi = g(`inp-${inp.key}`);
    const dz = g(`dz-${inp.key}`);
    if (!fi || !dz) continue;
    fi.addEventListener('change', () => updatePills(inp.key, fi));
    dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('drag'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
    dz.addEventListener('drop', e => {
      e.preventDefault(); dz.classList.remove('drag');
      fi.files = e.dataTransfer.files;
      updatePills(inp.key, fi);
    });
  }
}

function updatePills(key, fi) {
  const w = g(`fp-${key}`); if (!w) return;
  w.innerHTML = Array.from(fi.files).map((f,i) => `
    <div class="fpill">
      <span class="fpill-name">📄 ${h(f.name)}</span>
      <span class="fpill-size">${bytes(f.size)}</span>
      <button type="button" class="fpill-rm" onclick="rmFile('${key}',${i})">✕</button>
    </div>`).join('');
}

function rmFile(key, idx) {
  const fi = g(`inp-${key}`); if (!fi) return;
  const dt = new DataTransfer();
  Array.from(fi.files).forEach((f,i) => { if (i!==idx) dt.items.add(f); });
  fi.files = dt.files;
  updatePills(key, fi);
}

function pickR(gid, hid, val) {
  document.querySelectorAll(`#${gid} .radio-p`).forEach(el =>
    el.classList.toggle('on', el.dataset.v === val));
  const el = g(hid); if (el) el.value = val;
}

function tglChk(cbid, hid) {
  const cb = g(cbid); if (!cb) return;
  const on = !cb.classList.contains('on');
  cb.classList.toggle('on', on);
  cb.textContent = on ? '✓' : '';
  const el = g(hid); if (el) el.value = on ? 'true' : 'false';
}

function tglMs(msid, val, oid) {
  const chip = qs(`#${msid} [data-v="${val}"]`);
  if (!chip) return;
  chip.classList.toggle('on');
  const sel = [...document.querySelectorAll(`#${msid} .ms-c.on`)].map(c => c.dataset.v);
  const el = g(oid); if (el) el.value = JSON.stringify(sel);
}

function numSt(id, dir, step) {
  const el = g(id); if (!el) return;
  el.value = (parseFloat(el.value)||0) + dir * step;
}

// ═══════════════════════════════════════════════════════════════
//  ESECUZIONE
// ═══════════════════════════════════════════════════════════════
async function runTool(uid) {
  const btn = g('run-btn');
  const res = g('result');
  if (!btn || !res) return;

  btn.disabled = true;
  btn.classList.add('loading');
  btn.querySelector('.run-label').textContent = 'Elaborazione…';
  res.style.display = 'block';
  res.innerHTML = `
    <div style="padding:10px 0">
      <div class="prog-wrap"><div class="prog-fill"></div></div>
      <div style="font-size:12px;color:var(--t-mute);margin-top:6px;">Esecuzione in corso…</div>
    </div>`;

  try {
    const fd   = new FormData(g('run-form'));
    const tool = ST.tools.find(t => t.uid === uid);

    // Multiselect JSON → campi multipli
    if (tool) {
      for (const p of (tool.params||[])) {
        if (p.type === 'multiselect') {
          const v = fd.get(p.key);
          if (v) {
            fd.delete(p.key);
            try { JSON.parse(v).forEach(x => fd.append(p.key, x)); }
            catch(_) { fd.append(p.key, v); }
          }
        }
      }
    }

    const resp = await fetch(`/api/tools/${encodeURIComponent(uid)}/run`, {
      method: 'POST', body: fd,
    });
    if (!resp.ok) {
      let msg = resp.statusText;
      try { msg = (await resp.json()).detail || msg; } catch(_) {}
      throw new Error(msg);
    }

    const blob = await resp.blob();
    const url  = URL.createObjectURL(blob);
    const name = tool ? tool.id : 'output';

    res.innerHTML = `
      <div class="res-ok">
        <div class="res-title">✅ Completato!</div>
        <div class="res-msg">I file sono pronti per il download.</div>
        <a href="${url}" download="output_${h(name)}.zip" class="dl-btn">⬇ Scarica ZIP</a>
      </div>`;
    toast('Elaborazione completata!', 'ok');

  } catch(e) {
    res.innerHTML = `
      <div class="res-err">
        <div class="res-title">❌ Errore</div>
        <div class="res-msg">${h(e.message)}</div>
      </div>`;
    toast('Errore: ' + e.message, 'err');
  } finally {
    btn.disabled = false;
    btn.classList.remove('loading');
    btn.querySelector('.run-label').textContent = '▶  Esegui Tool';
  }
}

// ═══════════════════════════════════════════════════════════════
//  RELOAD
// ═══════════════════════════════════════════════════════════════
async function reloadTools() {
  try {
    await fetch('/api/tools/reload', { method: 'POST' });
    ST.tools = await GET('/tools');
    renderSidebar();
    showHome();
    toast('Tool ricaricati!', 'ok');
  } catch(e) { toast('Errore: ' + e.message, 'err'); }
}

// ═══════════════════════════════════════════════════════════════
//  RICERCA
// ═══════════════════════════════════════════════════════════════
function onSearch(val) {
  ST.q = val;
  renderSidebar();
  if (val) { ST.tab = 'Tutti'; showHome(); }
}

// ═══════════════════════════════════════════════════════════════
//  SETTINGS
// ═══════════════════════════════════════════════════════════════
function openSettings()  { g('cfg-modal').classList.add('open'); }
function closeSettings() { g('cfg-modal').classList.remove('open'); }

async function saveSettings() {
  const model = g('cfg-model').value.trim();
  const url   = g('cfg-url').value.trim();
  const key   = g('cfg-key').value.trim();
  try {
    await POST('/config', { ai: { model_id: model, base_url: url, api_key: key } });
    toast('Salvato!', 'ok');
    closeSettings();
  } catch(e) { toast('Errore: ' + e.message, 'err'); }
}

// ═══════════════════════════════════════════════════════════════
//  TOAST
// ═══════════════════════════════════════════════════════════════
function toast(msg, type) {
  const el = document.createElement('div');
  el.className = 'toast' + (type ? ' ' + type : '');
  el.textContent = msg;
  g('toasts').appendChild(el);
  requestAnimationFrame(() => el.classList.add('show'));
  setTimeout(() => {
    el.classList.remove('show');
    setTimeout(() => el.remove(), 250);
  }, 3000);
}

// ═══════════════════════════════════════════════════════════════
//  MARKDOWN minimale
// ═══════════════════════════════════════════════════════════════
function md(s) {
  if (!s) return '';
  let r = h(s);
  r = r.replace(/^#### (.+)$/gm, '<h4>$1</h4>');
  r = r.replace(/^### (.+)$/gm,  '<h3>$1</h3>');
  r = r.replace(/^## (.+)$/gm,   '<h2>$1</h2>');
  r = r.replace(/^# (.+)$/gm,    '<h1>$1</h1>');
  r = r.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  r = r.replace(/\*(.+?)\*/g,     '<em>$1</em>');
  r = r.replace(/`(.+?)`/g,       '<code>$1</code>');
  r = r.replace(/^[•\-\*] (.+)$/gm, '<li>$1</li>');
  r = r.replace(/(<li>.*?<\/li>\n?)+/gs, m => `<ul>${m}</ul>`);
  r = r.replace(/\n\n+/g, '</p><p>');
  r = r.replace(/\n/g, '<br>');
  return '<p>' + r + '</p>';
}

// ═══════════════════════════════════════════════════════════════
//  UTILITY
// ═══════════════════════════════════════════════════════════════
function bytes(n) {
  if (n < 1024) return n + ' B';
  if (n < 1048576) return (n/1024).toFixed(1) + ' KB';
  return (n/1048576).toFixed(1) + ' MB';
}
function acceptStr(tp) {
  if (tp === 'xlsx_single') return '.xlsx,.xls';
  if (tp === 'txt_single' || tp === 'txt_multi') return '.txt,.csv,.tsv';
  return '*';
}
function hintStr(tp) {
  if (tp === 'xlsx_single') return 'Formato: .xlsx / .xls';
  if (tp === 'txt_single' || tp === 'txt_multi') return 'Formato: .txt / .csv / .tsv';
  return 'Tutti i formati';
}

// ═══════════════════════════════════════════════════════════════
//  AI CHAT
// ═══════════════════════════════════════════════════════════════
const CHAT_HISTORY = [];  // { role, content }

function chatKey(e) {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); chatSend(); }
}

async function chatSend() {
  const inp  = g('chat-inp');
  const btn  = g('chat-send');
  if (!inp || !btn) return;
  const text = inp.value.trim();
  if (!text) return;

  inp.value = '';
  inp.style.height = '';
  chatAddMsg('user', text);
  CHAT_HISTORY.push({ role: 'user', content: text });

  btn.disabled = true;
  const typing = chatAddMsg('assistant', '…', true);

  try {
    const resp = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: CHAT_HISTORY }),
    });
    if (!resp.ok) {
      let err = resp.statusText;
      try { err = (await resp.json()).detail || err; } catch(_) {}
      throw new Error(err);
    }
    const data  = await resp.json();
    const reply = data.reply || '';
    typing.remove();
    chatAddMsg('assistant', reply);
    CHAT_HISTORY.push({ role: 'assistant', content: reply });
  } catch(e) {
    typing.remove();
    chatAddMsg('error', '❌ ' + e.message);
  } finally {
    btn.disabled = false;
    inp.focus();
  }
}

function chatAddMsg(role, text, isTyping = false) {
  const box = g('chat-msgs');
  if (!box) return null;
  const div = document.createElement('div');
  div.className = `cmsg cmsg-${role}`;
  if (isTyping) div.classList.add('typing');
  if (role === 'assistant' || role === 'error') {
    div.innerHTML = md(text);
  } else {
    div.textContent = text;
  }
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
  return div;
}

// ── Avvio ─────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', boot);
