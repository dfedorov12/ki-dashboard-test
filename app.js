'use strict';

// ═══════════════════════════════════════════════════════════════════
// CONFIG — bei Bedarf anpassen
// ═══════════════════════════════════════════════════════════════════
const CLIENT_ID    = '0eda579d-f557-430b-b69d-1afe3ee12fdd';
const TENANT_ID    = 'fdb70646-023a-403b-a4b9-1f474a935123';
const SP_HOST      = 'dihag.sharepoint.com';
const SP_SITE_PATH = '/sites/IT';
const LIST_ANTRAEGE  = 'KI_Antraege';
const LIST_LIZENZEN  = 'KI_Lizenzen';
const LIST_REGISTER  = 'KI_Register';

// KI-Register Spaltennamen
const COL_REG = {
  status:        'Status',
  risiko:        'Risikokategorie',
  verantw:       'VerantwortlicheStelle',
  hersteller:    'Hersteller',
  nutzungsart:   'InterneExterneNutzung',
  freigabeDatum: 'FreigabeDatum',
  anbieter:      'Anbieter',
  notizen:       'Notizen',
};

// SP-interne Spaltennamen (sofern abweichend von Anzeigenamme anpassen)
const COL = {
  status:           'Status',
  risiko:           'Risikokategorie',
  verantw:          'VerantwortlicheStelle',
  komponenten:      'KIKomponenten',
  hersteller:       'Hersteller',
  zweckHersteller:  'VerwendungszweckHersteller',
  zweckUnternehmen: 'AnwendungsbereichUnternehmen',
  nutzungsart:      'InterneExterneNutzung',
  projektplanung:   'Projektplanung',
  keyUser:          'KompetenzmassnahmeKeyUser',
  gremiumKommentar: 'GremiumKommentar',
  auflagen:         'Auflagen',
  freigabeDatum:    'FreigabeDatum',
  // Lizenzen
  kiSystem:         'KISystem',
  lizenztyp:        'Lizenztyp',
  anbieter:         'Anbieter',
  kosten:           'Kosten',
  rhythmus:         'Abrechnungsrhythmus',
  lizenzGesamt:     'LizenzenGesamt',
  lizenzBelegt:     'LizenzenBelegt',
  vertragsBeginn:   'VertragsBeginn',
  vertragsEnde:     'VertragsEnde',
  kuendigungsfrist: 'Kuendigungsfrist',
  autoRenewal:      'AutoRenewal',
  verantwIT:        'VerantwortlicherIT',
  notizen:          'Notizen',
};

const STATUS_OPTS = ['Eingereicht', 'In Prüfung', 'Genehmigt', 'Abgelehnt', 'Rückfrage'];

const ANTRAG_FIELDS = [
  { section: 'Grunddaten' },
  { key: 'Title',                        label: 'Bezeichnung des KI-Systems',        type: 'text',    req: true,
    hint: 'Interne Bezeichnung des KI-Use-Case oder der Software' },
  { key: COL.verantw,                    label: 'Verantwortliche Stelle',            type: 'text',    req: true,
    hint: 'Wer verantwortet die Lösung im Betrieb?' },
  { key: COL.hersteller,                 label: 'Hersteller / Entwickler',           type: 'text',    req: true,
    hint: 'Bezugsquelle, Dienstleister oder Lieferant' },
  { section: 'KI-Beschreibung' },
  { key: COL.komponenten,                label: 'KI-Komponente(n)',                  type: 'textarea', req: true,
    hint: 'Beschreibung der Funktionen, integrierten KI-Modelle und Verfahren' },
  { key: COL.zweckHersteller,            label: 'Verwendungszweck laut Hersteller',  type: 'textarea', req: true,
    hint: 'Wie definiert der Hersteller den Zweck? (Nutzungsbedingungen, Doku)' },
  { key: COL.zweckUnternehmen,           label: 'Anwendungsbereich im Unternehmen', type: 'textarea', req: true,
    hint: 'Zu welchem Zweck soll das System eingesetzt werden? Abweichung vom Herstellerzweck?' },
  { section: 'Klassifizierung' },
  { key: COL.risiko,                     label: 'Risikokategorie',                   type: 'choice',  req: true,
    choices: ['', 'Geringes Risiko', 'Normales Risiko', 'Hohes Risiko', 'Verboten'],
    hint: 'Eigene Einschätzung zu Sicherheit, Datenschutz und Grundrechten' },
  { key: COL.nutzungsart,                label: 'Nutzungsart',                       type: 'choice',  req: true,
    choices: ['', 'Intern', 'Extern', 'Intern & Extern'],
    hint: 'Nur intern oder auch als Angebot für Dritte/Vermarktung?' },
  { section: 'Umsetzung' },
  { key: COL.projektplanung,             label: 'Geplanter Einsatz ab',             type: 'date',    req: false },
  { key: COL.keyUser,                    label: 'Key User / Schulungsmaßnahme',     type: 'text',    req: false,
    hint: 'Wer ist Key User? Welche Schulungsmaßnahmen sind geplant?' },
];

const LIZENZ_FIELDS = [
  { key: 'Title',            label: 'Lizenzname',            type: 'text',     req: true },
  { key: COL.kiSystem,       label: 'KI-System',             type: 'text',     req: false },
  { key: COL.lizenztyp,      label: 'Lizenztyp',             type: 'choice',   req: true,
    choices: ['', 'Enterprise', 'Team', 'Pro', 'Free', 'API', 'Sonstiges'] },
  { key: COL.anbieter,       label: 'Anbieter',              type: 'text',     req: false },
  { key: COL.kosten,         label: 'Kosten (€/Periode)',    type: 'number',   req: false },
  { key: COL.rhythmus,       label: 'Abrechnungsrhythmus',   type: 'choice',   req: false,
    choices: ['', 'Monatlich', 'Jährlich', 'Einmalig'] },
  { key: COL.lizenzGesamt,   label: 'Lizenzen gesamt',       type: 'number',   req: false },
  { key: COL.lizenzBelegt,   label: 'Lizenzen belegt',       type: 'number',   req: false },
  { key: COL.vertragsBeginn, label: 'Vertragsbeginn',        type: 'date',     req: false },
  { key: COL.vertragsEnde,   label: 'Vertragsende',          type: 'date',     req: false },
  { key: COL.kuendigungsfrist,label:'Kündigungsfrist (Tage)',type: 'number',   req: false },
  { key: COL.autoRenewal,    label: 'Auto-Renewal',          type: 'choice',   req: false,
    choices: ['', 'Ja', 'Nein'] },
  { key: COL.verantwIT,      label: 'Verantwortlich IT',     type: 'text',     req: false },
  { key: COL.notizen,        label: 'Notizen',               type: 'textarea', req: false },
];

// ═══════════════════════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════════════════════
let msalInstance, account;
let siteId, listAntragId, listLizenzId, listRegisterId;
let isGremium = false;
let allAntraege = [], allLizenzen = [], allRegister = [];
let currentView = 'antrag';
let editLizenzId = null;

// ═══════════════════════════════════════════════════════════════════
// MSAL AUTH
// ═══════════════════════════════════════════════════════════════════
const SCOPES = ['https://graph.microsoft.com/Sites.ReadWrite.All', 'User.Read'];

async function initAuth() {
  const redirectUri = location.href.split('?')[0].split('#')[0];
  msalInstance = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}`, redirectUri },
    cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
  });
  await msalInstance.initialize();
  const r = await msalInstance.handleRedirectPromise();
  if (r) { account = r.account; msalInstance.setActiveAccount(r.account); return true; }
  const accs = msalInstance.getAllAccounts();
  if (accs.length) { account = accs[0]; msalInstance.setActiveAccount(accs[0]); return true; }
  return false;
}

async function doLogin() {
  $id('boot-btn').style.display = 'none';
  $id('boot-sub').textContent   = 'Weiterleitung zur Anmeldung…';
  $id('boot-spinner').style.display = 'block';
  try {
    await msalInstance.loginRedirect({ scopes: SCOPES });
  } catch(e) {
    $id('boot-err').textContent        = e.message;
    $id('boot-spinner').style.display  = 'none';
    $id('boot-btn').style.display      = 'block';
    $id('boot-btn').textContent        = 'Erneut versuchen';
  }
}

function logout() { msalInstance.logoutPopup({ postLogoutRedirectUri: location.href }); }

async function getToken() {
  try { return (await msalInstance.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
  catch(e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await msalInstance.acquireTokenRedirect({ scopes: SCOPES });
    }
    throw e;
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  $id('boot-spinner').style.display = 'block';
  try {
    const loggedIn = await initAuth();
    if (loggedIn) {
      await boot();
    } else {
      $id('boot-sub').textContent       = 'Bitte melden Sie sich an.';
      $id('boot-spinner').style.display = 'none';
      $id('boot-btn').style.display     = 'block';
    }
  } catch(e) {
    $id('boot-err').textContent        = e.message;
    $id('boot-spinner').style.display  = 'none';
    $id('boot-btn').style.display      = 'block';
  }
});

// ═══════════════════════════════════════════════════════════════════
// GRAPH API
// ═══════════════════════════════════════════════════════════════════
async function gFetch(url, opts = {}) {
  const token = await getToken();
  const full  = url.startsWith('http') ? url : `https://graph.microsoft.com/v1.0${url}`;
  const res   = await fetch(full, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', ...(opts.headers || {}) }
  });
  if (!res.ok) {
    const errBody = await res.json().catch(() => ({}));
    const detail  = errBody?.error?.message || errBody?.error?.code || res.statusText || res.status;
    console.error('Graph API error', res.status, JSON.stringify(errBody));
    throw Object.assign(new Error(`${res.status}: ${detail}`), { status: res.status, graphError: errBody });
  }
  return res.status === 204 ? null : res.json();
}
const gGet   = url        => gFetch(url);
const gPost  = (url, b)   => gFetch(url, { method: 'POST',   body: JSON.stringify(b) });
const gPatch = (url, b)   => gFetch(url, { method: 'PATCH',  body: JSON.stringify(b) });
const gDel   = url        => gFetch(url, { method: 'DELETE' });

// ═══════════════════════════════════════════════════════════════════
// STARTUP
// ═══════════════════════════════════════════════════════════════════
async function boot() {
  $id('boot-sub').textContent = 'Daten werden geladen…';
  try {
    const site = await gGet(`/sites/${SP_HOST}:${SP_SITE_PATH}`);
    siteId = site.id;
    console.log('Site ID:', siteId);

    // Listen direkt per internem URL-Namen abrufen (zuverlässiger als displayName-Filter)
    const [resA, resL, resR] = await Promise.allSettled([
      gGet(`/sites/${siteId}/lists/${LIST_ANTRAEGE}?$select=id,displayName,name`),
      gGet(`/sites/${siteId}/lists/${LIST_LIZENZEN}?$select=id,displayName,name`),
      gGet(`/sites/${siteId}/lists/${LIST_REGISTER}?$select=id,displayName,name`),
    ]);

    if (resA.status === 'fulfilled') {
      listAntragId = resA.value.id;
      console.log('Antraege list:', resA.value.displayName, listAntragId);
    } else {
      console.warn('Antraege list nicht gefunden:', resA.reason?.message);
    }

    if (resL.status === 'fulfilled') {
      listLizenzId = resL.value.id;
      console.log('Lizenzen list:', resL.value.displayName, listLizenzId);
      // Lesezugriff auf Lizenzen = Gremium-Mitglied
      try {
        await gGet(`/sites/${siteId}/lists/${listLizenzId}/items?$top=1`);
        isGremium = true;
      } catch(e) {
        if (e.status !== 403) console.warn('Lizenzliste Lesezugriff:', e.message);
      }
    } else {
      console.warn('Lizenzen list nicht gefunden:', resL.reason?.message);
    }

    if (resR.status === 'fulfilled') {
      listRegisterId = resR.value.id;
      console.log('Register list:', resR.value.displayName, listRegisterId);
    } else {
      console.warn('Register list nicht gefunden:', resR.reason?.message);
    }

    $id('boot').style.display = 'none';
    $id('app').style.display  = 'flex';

    const name = account?.name || account?.username || '';
    $id('user-name').textContent = name;

    if (isGremium) {
      $id('gremium-badge').classList.remove('hidden');
    } else {
      // Nicht-Gremium: Lizenzen ausblenden, Register bleibt sichtbar (Read-only)
      document.querySelector('[data-view="lizenzen"]').style.display = 'none';
    }

    // Register-Tab ausblenden wenn Liste nicht erreichbar
    if (!listRegisterId) {
      document.querySelector('[data-view="register"]').style.display = 'none';
    }

    renderAntragForm();
  } catch(e) {
    $id('boot-err').textContent       = 'Fehler: ' + e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-btn').style.display     = 'block';
    $id('boot-btn').textContent       = 'Erneut versuchen';
  }
}

// ═══════════════════════════════════════════════════════════════════
// VIEW SWITCHING
// ═══════════════════════════════════════════════════════════════════
async function switchView(view) {
  currentView = view;
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  $id(`view-${view}`)?.classList.add('active');
  document.querySelector(`[data-view="${view}"]`)?.classList.add('active');

  if (view === 'antraege' && !allAntraege.length)  await loadAntraege();
  if (view === 'lizenzen' && !allLizenzen.length)  await loadLizenzen();
  if (view === 'register' && !allRegister.length)  await loadRegister();
}

// ═══════════════════════════════════════════════════════════════════
// ANTRAG FORM
// ═══════════════════════════════════════════════════════════════════
function renderAntragForm() {
  let html = '';
  let inSection = false;
  let inRow = false;

  const closeRow = () => { if (inRow) { html += '</div>'; inRow = false; } };
  const closeSection = () => { closeRow(); if (inSection) { html += '</div>'; inSection = false; } };

  for (const f of ANTRAG_FIELDS) {
    if (f.section) {
      closeSection();
      html += `<div class="form-section"><div class="form-section-title">${esc(f.section)}</div><div class="form-row">`;
      inSection = true; inRow = true;
      continue;
    }
    const isWide = f.type === 'textarea';
    const cls = isWide ? 'form-group full' : 'form-group';
    html += `<div class="${cls}">
      <label class="form-label" for="f-${f.key}">${esc(f.label)}${f.req ? '<span class="req">*</span>' : ''}</label>`;

    if (f.type === 'textarea') {
      html += `<textarea id="f-${f.key}" name="${f.key}" class="form-control" rows="3"${f.req ? ' required' : ''}></textarea>`;
    } else if (f.type === 'choice') {
      html += `<select id="f-${f.key}" name="${f.key}" class="form-control"${f.req ? ' required' : ''}>`;
      for (const c of f.choices) html += `<option value="${esc(c)}">${esc(c) || '– bitte wählen –'}</option>`;
      html += '</select>';
    } else {
      html += `<input id="f-${f.key}" name="${f.key}" type="${f.type}" class="form-control"${f.req ? ' required' : ''}/>`;
    }

    if (f.hint) html += `<div class="form-hint">${esc(f.hint)}</div>`;
    html += '</div>';
  }

  closeSection();
  $id('form-antrag-fields').innerHTML = html;
}

async function submitAntrag(e) {
  e.preventDefault();
  const btn = $id('btn-submit');

  // Validation
  let valid = true;
  document.querySelectorAll('#form-antrag-fields [required]').forEach(el => {
    el.classList.remove('invalid');
    if (!el.value.trim()) { el.classList.add('invalid'); valid = false; }
  });
  if (!valid) return;

  btn.disabled = true; btn.textContent = 'Wird eingereicht …';
  $id('antrag-success').classList.add('hidden');

  const fields = { [COL.status]: 'Eingereicht' };
  for (const f of ANTRAG_FIELDS) {
    if (f.section) continue;
    const el = $id(`f-${f.key}`);
    if (!el) continue;
    const v = el.value.trim();
    if (v) fields[f.key] = spValue(f.type, v);
  }

  try {
    await gPost(`/sites/${siteId}/lists/${listAntragId}/items`, { fields });
    $id('form-antrag').reset();
    const s = $id('antrag-success');
    s.textContent = '✓ Ihr Antrag wurde eingereicht. Das KI-Koordinierungsgremium wird ihn prüfen und Sie per E-Mail informieren.';
    s.classList.remove('hidden');
    allAntraege = [];
    updateOpenBadge();
  } catch(err) {
    alert('Fehler beim Einreichen: ' + err.message);
  }

  btn.disabled = false; btn.textContent = 'Antrag einreichen';
}

// ═══════════════════════════════════════════════════════════════════
// ANTRÄGE LIST
// ═══════════════════════════════════════════════════════════════════
async function loadAntraege() {
  if (!listAntragId) {
    $id('antraege-loading').textContent = 'Liste "' + LIST_ANTRAEGE + '" nicht gefunden.';
    return;
  }
  $id('antraege-loading').classList.remove('hidden');
  $id('antraege-list').innerHTML = '';

  try {
    const data = await gGet(`/sites/${siteId}/lists/${listAntragId}/items?$expand=fields($select=*)&$top=999&$orderby=fields/Created desc`);
    allAntraege = data.value || [];
    renderAntraege();
    updateOpenBadge();
  } catch(e) {
    $id('antraege-loading').textContent = 'Fehler: ' + e.message;
  }
}

function filterAntraege() { renderAntraege(); }

function renderAntraege() {
  const statusF = $id('filter-status')?.value || '';
  const riskF   = $id('filter-risk')?.value   || '';

  let items = allAntraege.filter(i => {
    const f = i.fields;
    if (statusF && f[COL.status] !== statusF) return false;
    if (riskF   && f[COL.risiko] !== riskF)  return false;
    return true;
  });

  $id('antraege-loading').classList.add('hidden');
  $id('antraege-sub').textContent = isGremium
    ? `${items.length} Antrag/Anträge angezeigt`
    : 'Meine eingereichten Anträge';

  if (!items.length) {
    $id('antraege-list').innerHTML = '<div class="empty-state">Keine Anträge gefunden.</div>';
    return;
  }

  $id('antraege-list').innerHTML = items.map(i => {
    const f  = i.fields;
    const dt = fmtDate(f.Created || i.createdDateTime);
    const by = f.Author0LookupValue || f['Author/Title'] || '';
    return `<div class="item-card" onclick="openAntragPanel(${i.id})">
      <div class="card-top">
        <div class="card-title">${esc(f.Title || '–')}</div>
        ${statusBadge(f[COL.status])}
      </div>
      <div class="card-tags">
        ${riskBadge(f[COL.risiko])}
        ${f[COL.nutzungsart] ? `<span class="badge-type">${esc(f[COL.nutzungsart])}</span>` : ''}
      </div>
      <div class="card-meta">
        <span>${esc(f[COL.hersteller] || '')}</span>
        <span>📅 ${dt}</span>
        ${by ? `<span>👤 ${esc(by)}</span>` : ''}
      </div>
    </div>`;
  }).join('');
}

function updateOpenBadge() {
  const open = allAntraege.filter(i =>
    ['Eingereicht', 'In Prüfung', 'Rückfrage'].includes(i.fields?.[COL.status])
  ).length;
  const b = $id('badge-open');
  if (open) { b.textContent = open; b.classList.remove('hidden'); }
  else b.classList.add('hidden');
}

// ═══════════════════════════════════════════════════════════════════
// ANTRAG PANEL
// ═══════════════════════════════════════════════════════════════════
function openAntragPanel(itemId) {
  const item = allAntraege.find(i => i.id == itemId);
  if (!item) return;
  const f = item.fields;

  $id('panel-title').innerHTML = `${statusBadge(f[COL.status])} <span style="margin-left:8px">${esc(f.Title || '–')}</span>`;

  const row = (label, value, pre = false) =>
    `<div class="panel-field">
      <div class="panel-field-label">${esc(label)}</div>
      <div class="panel-field-value${pre ? ' pre' : ''}">${value || '<span style="color:#9ca3af">–</span>'}</div>
    </div>`;

  const rows1 = `
    <div class="panel-section">
      <div class="panel-section-title">Grunddaten</div>
      ${row('Bezeichnung',         esc(f.Title))}
      ${row('Verantwortl. Stelle', esc(f[COL.verantw]))}
      ${row('Hersteller',          esc(f[COL.hersteller]))}
      ${row('Nutzungsart',         esc(f[COL.nutzungsart]))}
      ${row('Risikokategorie',     riskBadge(f[COL.risiko]))}
      ${row('Geplanter Einsatz',   fmtDate(f[COL.projektplanung]))}
      ${row('Key User / Schulung', esc(f[COL.keyUser]))}
    </div>
    <div class="panel-section">
      <div class="panel-section-title">KI-Beschreibung</div>
      ${row('KI-Komponenten',             esc(f[COL.komponenten]), true)}
      ${row('Zweck laut Hersteller',      esc(f[COL.zweckHersteller]), true)}
      ${row('Anwendungsbereich intern',   esc(f[COL.zweckUnternehmen]), true)}
    </div>`;

  const gremiumSection = isGremium ? `
    <div class="panel-gremium">
      <div class="panel-gremium-title">⚖️ Gremium-Entscheidung</div>
      <div class="form-group">
        <label class="form-label">Status</label>
        <select id="pg-status" class="form-control">
          ${STATUS_OPTS.map(s => `<option value="${s}"${f[COL.status] === s ? ' selected' : ''}>${s}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">Kommentar / Begründung</label>
        <textarea id="pg-kommentar" class="form-control" rows="3">${esc(f[COL.gremiumKommentar] || '')}</textarea>
      </div>
      <div class="form-group">
        <label class="form-label">Auflagen / Bedingungen</label>
        <textarea id="pg-auflagen" class="form-control" rows="3">${esc(f[COL.auflagen] || '')}</textarea>
      </div>
      <div class="panel-actions">
        <button class="btn btn-success btn-sm" onclick="saveGremiumDecision(${item.id},'Genehmigt')">✓ Genehmigen</button>
        <button class="btn btn-danger btn-sm"  onclick="saveGremiumDecision(${item.id},'Abgelehnt')">✕ Ablehnen</button>
        <button class="btn btn-neutral btn-sm" onclick="saveGremiumDecision(${item.id},'Rückfrage')">? Rückfrage</button>
        <button class="btn btn-neutral btn-sm" onclick="saveGremiumDecision(${item.id})">💾 Speichern</button>
      </div>
    </div>` : (f[COL.gremiumKommentar] ? `
    <div class="panel-section">
      <div class="panel-section-title">Gremium-Rückmeldung</div>
      ${row('Kommentar', esc(f[COL.gremiumKommentar]), true)}
      ${f[COL.auflagen] ? row('Auflagen', esc(f[COL.auflagen]), true) : ''}
      ${f[COL.freigabeDatum] ? row('Freigabedatum', fmtDate(f[COL.freigabeDatum])) : ''}
    </div>` : '');

  $id('panel-body').innerHTML = rows1 + gremiumSection;
  openPanel();
}

async function saveGremiumDecision(itemId, forceStatus) {
  const status    = forceStatus || $id('pg-status')?.value;
  const kommentar = $id('pg-kommentar')?.value?.trim() || '';
  const auflagen  = $id('pg-auflagen')?.value?.trim()  || '';

  const fields = { [COL.status]: status };
  if (kommentar) fields[COL.gremiumKommentar] = kommentar;
  if (auflagen)  fields[COL.auflagen]         = auflagen;
  if (status === 'Genehmigt') fields[COL.freigabeDatum] = new Date().toISOString().slice(0, 10);

  try {
    await gPatch(`/sites/${siteId}/lists/${listAntragId}/items/${itemId}/fields`, fields);
    const idx = allAntraege.findIndex(i => i.id == itemId);
    if (idx >= 0) Object.assign(allAntraege[idx].fields, fields);
    closePanel();
    renderAntraege();
    updateOpenBadge();
  } catch(e) {
    alert('Fehler beim Speichern: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// LIZENZEN
// ═══════════════════════════════════════════════════════════════════
async function loadLizenzen() {
  if (!listLizenzId) {
    $id('lizenzen-loading').textContent = 'Liste "' + LIST_LIZENZEN + '" nicht gefunden oder kein Zugriff.';
    return;
  }
  $id('lizenzen-loading').classList.remove('hidden');
  $id('lizenzen-wrap').innerHTML = '';

  try {
    const data = await gGet(`/sites/${siteId}/lists/${listLizenzId}/items?$expand=fields($select=*)&$top=999`);
    allLizenzen = data.value || [];
    renderLizenzen();
  } catch(e) {
    $id('lizenzen-loading').textContent = 'Fehler: ' + e.message;
  }
}

function renderLizenzen() {
  $id('lizenzen-loading').classList.add('hidden');

  const today = new Date();
  const totalKosten = allLizenzen.reduce((s, i) => s + (parseFloat(i.fields?.[COL.kosten]) || 0), 0);
  const expireSoon  = allLizenzen.filter(i => {
    const d = i.fields?.[COL.vertragsEnde];
    if (!d) return false;
    const diff = (new Date(d) - today) / 86400000;
    return diff >= 0 && diff <= 60;
  }).length;

  const stats = `<div class="stats-row">
    <div class="stat-card accent"><div class="stat-value">${allLizenzen.length}</div><div class="stat-label">Lizenzen gesamt</div></div>
    <div class="stat-card green"><div class="stat-value">${fmtEuro(totalKosten)}</div><div class="stat-label">Kosten p.a. (geschätzt)</div></div>
    <div class="stat-card ${expireSoon ? 'red' : 'orange'}"><div class="stat-value">${expireSoon}</div><div class="stat-label">Ablauf in &lt; 60 Tagen</div></div>
  </div>`;

  if (!allLizenzen.length) {
    $id('lizenzen-wrap').innerHTML = stats + '<div class="empty-state">Noch keine Lizenzen erfasst.</div>';
    return;
  }

  const rows = allLizenzen.map(i => {
    const f      = i.fields;
    const ende   = f[COL.vertragsEnde];
    const diff   = ende ? (new Date(ende) - today) / 86400000 : null;
    const endeCls = diff === null ? 'expiry-ok' : diff < 0 ? 'expiry-alert' : diff < 30 ? 'expiry-alert' : diff < 60 ? 'expiry-warn' : 'expiry-ok';
    const endeLabel = ende ? fmtDate(ende) : '–';

    const gesamt = parseInt(f[COL.lizenzGesamt]) || 0;
    const belegt = parseInt(f[COL.lizenzBelegt]) || 0;
    const pct    = gesamt > 0 ? Math.min(100, Math.round(belegt / gesamt * 100)) : null;
    const barCls = pct === null ? '' : pct >= 90 ? 'util-full' : pct >= 70 ? 'util-warn' : 'util-ok';
    const util   = pct !== null ? `<div style="display:flex;align-items:center;gap:6px">
      <div class="util-bar-wrap"><div class="util-bar ${barCls}" style="width:${pct}%"></div></div>
      <span style="font-size:11px;color:#6b7280">${belegt}/${gesamt}</span>
    </div>` : '–';

    return `<tr onclick="openLizenzModal(${i.id})">
      <td><strong>${esc(f.Title || '–')}</strong></td>
      <td>${esc(f[COL.kiSystem] || '–')}</td>
      <td>${f[COL.lizenztyp] ? `<span class="badge-type">${esc(f[COL.lizenztyp])}</span>` : '–'}</td>
      <td>${esc(f[COL.anbieter] || '–')}</td>
      <td>${f[COL.kosten] ? fmtEuro(parseFloat(f[COL.kosten])) : '–'}</td>
      <td>${util}</td>
      <td class="${endeCls}">${endeLabel}${diff !== null && diff < 60 && diff >= 0 ? ` <small>(${Math.round(diff)}d)</small>` : ''}</td>
      <td><span style="font-size:11px">${f[COL.autoRenewal] || '–'}</span></td>
    </tr>`;
  }).join('');

  $id('lizenzen-wrap').innerHTML = stats + `<div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Lizenzname</th><th>KI-System</th><th>Typ</th><th>Anbieter</th>
          <th>Kosten</th><th>Auslastung</th><th>Vertragsende</th><th>Auto-Renewal</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

// ─── Lizenz Modal ────────────────────────────────────────────────
function openLizenzModal(itemId) {
  editLizenzId = itemId || null;
  const item = itemId ? allLizenzen.find(i => i.id == itemId) : null;
  const f = item?.fields || {};

  $id('modal-title').textContent = itemId ? 'Lizenz bearbeiten' : 'Neue Lizenz erfassen';

  let html = '<div class="form-row" style="grid-template-columns:1fr 1fr">';
  for (const field of LIZENZ_FIELDS) {
    const v   = f[field.key] ?? '';
    const cls = field.type === 'textarea' ? 'form-group full' : 'form-group';
    html += `<div class="${cls}">
      <label class="form-label" for="lf-${field.key}">${esc(field.label)}${field.req ? '<span class="req">*</span>' : ''}</label>`;

    if (field.type === 'textarea') {
      html += `<textarea id="lf-${field.key}" class="form-control" rows="2">${esc(v)}</textarea>`;
    } else if (field.type === 'choice') {
      html += `<select id="lf-${field.key}" class="form-control">`;
      for (const c of field.choices) html += `<option value="${esc(c)}"${v === c ? ' selected' : ''}>${esc(c) || '– wählen –'}</option>`;
      html += '</select>';
    } else {
      html += `<input id="lf-${field.key}" type="${field.type}" class="form-control" value="${esc(v)}"/>`;
    }
    html += '</div>';
  }
  html += '</div>';

  html += `<div class="modal-footer">
    ${itemId ? `<button class="btn btn-danger btn-sm" onclick="deleteLizenz(${itemId})">Löschen</button><span style="flex:1"></span>` : ''}
    <button class="btn btn-neutral btn-sm" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary btn-sm" onclick="saveLizenz()">Speichern</button>
  </div>`;

  $id('modal-body').innerHTML = html;
  $id('modal-overlay').classList.remove('hidden');
}

async function saveLizenz() {
  const fields = {};
  for (const f of LIZENZ_FIELDS) {
    const el = $id(`lf-${f.key}`);
    if (!el) continue;
    const v = el.value.trim();
    if (v !== '') fields[f.key] = spValue(f.type, v);
  }

  if (!fields.Title) { alert('Bitte Lizenzname eingeben.'); return; }

  try {
    if (editLizenzId) {
      await gPatch(`/sites/${siteId}/lists/${listLizenzId}/items/${editLizenzId}/fields`, fields);
    } else {
      await gPost(`/sites/${siteId}/lists/${listLizenzId}/items`, { fields });
    }
    closeModal();
    allLizenzen = [];
    await loadLizenzen();
  } catch(e) {
    alert('Fehler: ' + e.message);
  }
}

async function deleteLizenz(itemId) {
  const item = allLizenzen.find(i => i.id == itemId);
  if (!confirm(`Lizenz "${item?.fields?.Title}" wirklich löschen?`)) return;
  try {
    await gDel(`/sites/${siteId}/lists/${listLizenzId}/items/${itemId}`);
    closeModal();
    allLizenzen = [];
    await loadLizenzen();
  } catch(e) {
    alert('Fehler: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// KI-REGISTER
// ═══════════════════════════════════════════════════════════════════
async function loadRegister() {
  if (!listRegisterId) {
    $id('register-loading').textContent = 'Liste "' + LIST_REGISTER + '" nicht gefunden oder kein Zugriff.';
    return;
  }
  $id('register-loading').classList.remove('hidden');
  $id('register-wrap').innerHTML = '';

  try {
    const data = await gGet(`/sites/${siteId}/lists/${listRegisterId}/items?$expand=fields($select=*)&$top=999&$orderby=fields/Created desc`);
    allRegister = data.value || [];
    renderRegister();
  } catch(e) {
    $id('register-loading').textContent = 'Fehler: ' + e.message;
  }
}

function filterRegister() { renderRegister(); }

function renderRegister() {
  const statusF = $id('reg-filter-status')?.value || '';
  const riskF   = $id('reg-filter-risk')?.value   || '';

  let items = allRegister.filter(i => {
    const f = i.fields;
    if (statusF && f[COL_REG.status] !== statusF) return false;
    if (riskF   && f[COL_REG.risiko] !== riskF)  return false;
    return true;
  });

  $id('register-loading').classList.add('hidden');

  const aktiv = allRegister.filter(i => (i.fields?.[COL_REG.status] || '').toLowerCase() === 'aktiv').length;
  const stats = `<div class="stats-row">
    <div class="stat-card accent"><div class="stat-value">${allRegister.length}</div><div class="stat-label">Einträge gesamt</div></div>
    <div class="stat-card green"><div class="stat-value">${aktiv}</div><div class="stat-label">Aktive Systeme</div></div>
    <div class="stat-card orange"><div class="stat-value">${allRegister.length - aktiv}</div><div class="stat-label">Inaktiv / Archiviert</div></div>
  </div>`;

  if (!items.length) {
    $id('register-wrap').innerHTML = stats + '<div class="empty-state">Keine Einträge gefunden.</div>';
    return;
  }

  const rows = items.map(i => {
    const f = i.fields;
    return `<tr onclick="openRegisterPanel(${i.id})" style="cursor:pointer">
      <td><strong>${esc(f.Title || '–')}</strong></td>
      <td>${esc(f[COL_REG.verantw] || '–')}</td>
      <td>${esc(f[COL_REG.hersteller] || f[COL_REG.anbieter] || '–')}</td>
      <td>${riskBadge(f[COL_REG.risiko])}</td>
      <td>${f[COL_REG.nutzungsart] ? `<span class="badge-type">${esc(f[COL_REG.nutzungsart])}</span>` : '–'}</td>
      <td>${statusBadge(f[COL_REG.status])}</td>
      <td>${fmtDate(f[COL_REG.freigabeDatum])}</td>
    </tr>`;
  }).join('');

  $id('register-wrap').innerHTML = stats + `<div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>KI-System</th><th>Verantwortl. Stelle</th><th>Hersteller</th>
          <th>Risiko</th><th>Nutzungsart</th><th>Status</th><th>Freigabe</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

function openRegisterPanel(itemId) {
  const item = allRegister.find(i => i.id == itemId);
  if (!item) return;
  const f = item.fields;

  $id('panel-title').innerHTML = `${statusBadge(f[COL_REG.status])} <span style="margin-left:8px">${esc(f.Title || '–')}</span>`;

  const row = (label, value, pre = false) =>
    `<div class="panel-field">
      <div class="panel-field-label">${esc(label)}</div>
      <div class="panel-field-value${pre ? ' pre' : ''}">${value || '<span style="color:#9ca3af">–</span>'}</div>
    </div>`;

  $id('panel-body').innerHTML = `
    <div class="panel-section">
      <div class="panel-section-title">Stammdaten</div>
      ${row('Bezeichnung',           esc(f.Title))}
      ${row('Verantwortl. Stelle',   esc(f[COL_REG.verantw]))}
      ${row('Hersteller / Anbieter', esc(f[COL_REG.hersteller] || f[COL_REG.anbieter]))}
      ${row('Nutzungsart',           esc(f[COL_REG.nutzungsart]))}
      ${row('Risikokategorie',       riskBadge(f[COL_REG.risiko]))}
      ${row('Status',                statusBadge(f[COL_REG.status]))}
      ${row('Freigabedatum',         fmtDate(f[COL_REG.freigabeDatum]))}
    </div>
    ${f[COL_REG.notizen] ? `<div class="panel-section">
      <div class="panel-section-title">Notizen</div>
      ${row('', esc(f[COL_REG.notizen]), true)}
    </div>` : ''}`;

  openPanel();
}

// ═══════════════════════════════════════════════════════════════════
// PANEL / MODAL HELPERS
// ═══════════════════════════════════════════════════════════════════
function openPanel() {
  $id('panel-overlay').classList.remove('hidden');
  $id('side-panel').classList.remove('hidden');
}
function closePanel() {
  $id('panel-overlay').classList.add('hidden');
  $id('side-panel').classList.add('hidden');
}
function closeModal(e) {
  if (e && e.target !== $id('modal-overlay')) return;
  $id('modal-overlay').classList.add('hidden');
}

// ═══════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════
function $id(id) { return document.getElementById(id); }

// Konvertiert Formulareingabe in den von der Graph API erwarteten Typ
function spValue(type, v) {
  if (type === 'number') return parseFloat(v);
  if (type === 'date')   return v ? new Date(v).toISOString() : null;
  return v;
}

function esc(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function fmtDate(s) {
  if (!s) return '–';
  const d = new Date(s);
  if (isNaN(d)) return s;
  return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`;
}

function fmtEuro(n) {
  if (n === null || n === undefined || isNaN(n)) return '–';
  return new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(n);
}

function statusBadge(s) {
  const label = s || 'Eingereicht';
  const map = {
    'eingereicht':  ['s-eingereicht',  '📩'],
    'in prüfung':   ['s-in-pruefung',  '🔍'],
    'genehmigt':    ['s-genehmigt',    '✓'],
    'abgelehnt':    ['s-abgelehnt',    '✕'],
    'rückfrage':    ['s-rueckfrage',   '?'],
  };
  const [cls, icon] = map[label.toLowerCase().trim()] || ['s-eingereicht', '•'];
  return `<span class="badge-status ${cls}">${icon} ${esc(label)}</span>`;
}

function riskBadge(r) {
  if (!r) return '';
  const map = {
    'geringes risiko': 'r-gering',
    'normales risiko': 'r-normal',
    'hohes risiko':    'r-hoch',
    'verboten':        'r-verboten',
  };
  const cls = map[r.toLowerCase().trim()] || 'r-normal';
  return `<span class="badge-risk ${cls}">${esc(r)}</span>`;
}
