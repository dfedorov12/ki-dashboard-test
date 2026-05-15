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

// Bekannte GUIDs als Fallback (aus Browser-Netzwerklog ermittelt)
const KNOWN_SITE_ID        = 'dihag.sharepoint.com,1618712f-787b-4584-ad54-2bf68c110f15,b93e94cf-030f-4296-9756-15492a5409d9';
const KNOWN_ANTRAEGE_GUID  = '28d7d466-6239-4575-be27-4c3873634707';

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
  kiSystem:         'System',
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
  nutzer:           'KIUser',
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
  { key: COL.kiSystem,       label: 'KI-System',             type: 'text',     req: true },
  { key: COL.lizenztyp,      label: 'Lizenztyp',             type: 'combo',    req: false,
    choices: ['Enterprise', 'Team', 'Pro', 'Free', 'API', 'Sonstiges'] },
  { key: COL.anbieter,       label: 'Anbieter',              type: 'text',     req: false },
  { key: COL.kosten,         label: 'Kosten (€/Periode)',    type: 'number',   req: false },
  { key: COL.rhythmus,       label: 'Abrechnungsrhythmus',   type: 'choice',   req: false,
    choices: ['', 'Monatlich', 'Jährlich', 'Einmalig'] },
  { key: COL.lizenzGesamt,   label: 'Lizenzen gesamt',       type: 'number',   req: false },
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
let lizenzUsers = []; // current modal's user list
// Gültige schreibbare Spaltennamen der Listen (wird in boot() befüllt)
let antragCols = null;  // null = noch nicht geladen (= kein Filter)

// ═══════════════════════════════════════════════════════════════════
// MSAL AUTH
// ═══════════════════════════════════════════════════════════════════
const SCOPES = ['https://graph.microsoft.com/Sites.ReadWrite.All', 'User.Read', 'People.Read'];

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
  const token  = await getToken();
  const full   = url.startsWith('http') ? url : `https://graph.microsoft.com/v1.0${url}`;
  const method = (opts.method || 'GET').toUpperCase();
  // Prefer-Header nur bei GET-Abfragen (verhindert Probleme bei POST/PATCH)
  const preferHdr = method === 'GET'
    ? { 'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly' }
    : {};
  const res = await fetch(full, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...preferHdr,
      ...(opts.headers || {})
    }
  });
  if (!res.ok) {
    const errBody = await res.json().catch(() => ({}));
    const detail  = errBody?.error?.message || errBody?.error?.code || res.statusText || res.status;
    console.error('Graph API error', method, res.status, JSON.stringify(errBody));
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
    // Site-ID: zuerst dynamisch auflösen, sonst bekannte GUID als Fallback
    try {
      const site = await gGet(`/sites/${SP_HOST}:${SP_SITE_PATH}`);
      siteId = site.id;
    } catch(e) {
      console.warn('Site-Lookup fehlgeschlagen, nutze Fallback-ID:', e.message);
      siteId = KNOWN_SITE_ID;
    }

    // Alle Listen der Site abrufen; matcht intern (name) UND Anzeigenamen (displayName)
    let allLists = [];
    try {
      let nextUrl = `/sites/${siteId}/lists?$select=id,displayName,name&$top=200`;
      while (nextUrl) {
        const page = await gGet(nextUrl);
        allLists = allLists.concat(page.value || []);
        nextUrl = page['@odata.nextLink'] || null;
      }
    } catch(e) {
      console.warn('Listen-Übersicht fehlgeschlagen:', e.message);
    }

    const findList = key => allLists.find(
      l => l.name?.toLowerCase()        === key.toLowerCase() ||
           l.displayName?.toLowerCase() === key.toLowerCase()
    );

    const lA = findList(LIST_ANTRAEGE);
    const lL = findList(LIST_LIZENZEN);
    const lR = findList(LIST_REGISTER);

    // Antraege: dynamisch oder bekannte GUID als Fallback
    listAntragId = lA?.id || KNOWN_ANTRAEGE_GUID;
    if (lA) console.log('✓ KI_Antraege:', lA.displayName, listAntragId);
    else    console.log('⚠ KI_Antraege nicht via Listen-API gefunden → Fallback-GUID:', listAntragId);

    // Tatsächliche schreibbare Spaltennamen der Antraege-Liste abrufen
    // → verhindert 500-Fehler bei ungültigen Feldnamen
    try {
      const colData = await gGet(`/sites/${siteId}/lists/${listAntragId}/columns?$select=name,displayName,readOnly,hidden&$top=200`);
      antragCols = new Set(
        (colData.value || []).filter(c => !c.readOnly && !c.hidden).map(c => c.name)
      );
      console.log('✓ Antraege-Spalten:', [...antragCols].sort().join(', '));
    } catch(e) {
      console.warn('Spaltenabruf fehlgeschlagen (kein Filter aktiv):', e.message);
    }

    // Lizenzen: nur wenn gefunden
    if (lL) {
      listLizenzId = lL.id;
      console.log('✓ KI_Lizenzen:', lL.displayName, listLizenzId);
      try {
        await gGet(`/sites/${siteId}/lists/${listLizenzId}/items?$top=1`);
        isGremium = true;
        console.log('✓ Gremium-Zugriff bestätigt');

        // Echte interne Spaltennamen der Lizenzen-Liste ermitteln
        // (SharePoint-intern kann 'System' z.B. 'System0' heißen)
        try {
          const lizCols = await gGet(`/sites/${siteId}/lists/${listLizenzId}/columns?$select=name,displayName&$top=200`);
          for (const col of (lizCols.value || [])) {
            if (col.displayName === 'System') {
              const oldKey = COL.kiSystem;
              COL.kiSystem = col.name;
              // Auch LIZENZ_FIELDS-Key aktualisieren (wird zur Laufzeit genutzt)
              const lf = LIZENZ_FIELDS.find(f => f.key === oldKey);
              if (lf) lf.key = col.name;
              console.log('✓ KI-System Spalte:', oldKey, '→', col.name);
            }
            if (col.displayName === 'KI-User') {
              COL.nutzer = col.name;
              console.log('✓ KI-User Spalte:', col.name);
            }
          }
        } catch(eCols) {
          console.warn('Lizenzen Spalten-Lookup fehlgeschlagen:', eCols.message);
        }
      } catch(e) {
        if (e.status !== 403) console.warn('Lizenzen Lesezugriff:', e.message);
        else console.log('ℹ Kein Gremium-Zugriff auf Lizenzen (403)');
      }
    } else {
      console.warn('⚠ KI_Lizenzen nicht gefunden. Verfügbare Listen:',
        allLists.map(l => `${l.name}/${l.displayName}`).join(' | '));
    }

    // Register
    if (lR) {
      listRegisterId = lR.id;
      console.log('✓ KI_Register:', lR.displayName, listRegisterId);
    } else {
      console.warn('⚠ KI_Register nicht gefunden');
    }

    $id('boot').style.display = 'none';
    $id('app').style.display  = 'flex';

    const uName = account?.name || account?.username || '';
    $id('user-name').textContent = uName;

    // Tabs: Lizenzen nur für Gremium, Register nur wenn Liste bekannt
    if (isGremium) {
      $id('gremium-badge').classList.remove('hidden');
    } else {
      document.querySelector('[data-view="lizenzen"]').style.display = 'none';
    }
    if (!listRegisterId) {
      document.querySelector('[data-view="register"]').style.display = 'none';
    }

    renderAntragForm();
  } catch(e) {
    $id('boot-err').textContent       = 'Fehler beim Start: ' + e.message;
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
  // hidden-Klasse entfernen/setzen — .hidden hat display:none !important
  // und würde .view.active überschreiben wenn nicht explizit entfernt
  document.querySelectorAll('.view').forEach(v => {
    v.classList.remove('active');
    v.classList.add('hidden');
  });
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  const activeView = $id(`view-${view}`);
  if (activeView) { activeView.classList.remove('hidden'); activeView.classList.add('active'); }
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

  // Pflichtfelder prüfen
  let valid = true;
  document.querySelectorAll('#form-antrag-fields [required]').forEach(el => {
    el.classList.remove('invalid');
    if (!el.value.trim()) { el.classList.add('invalid'); valid = false; }
  });
  if (!valid) return;

  btn.disabled = true; btn.textContent = 'Wird eingereicht …';
  removeAntragError();

  // Spalten-Check-Funktion
  const colOk = name => !antragCols || antragCols.has(name);

  // Alle Formularwerte sammeln (ohne Status)
  const detailFields = {};
  for (const f of ANTRAG_FIELDS) {
    if (f.section) continue;
    const el = $id(`f-${f.key}`);
    if (!el || f.key === 'Title') continue;
    const v = el.value.trim();
    if (!v) continue;
    if (colOk(f.key)) detailFields[f.key] = spValue(f.type, v);
    else console.warn('Spalte nicht in SP-Liste gefunden, übersprungen:', f.key);
  }

  const titleEl = $id(`f-Title`);
  const titleVal = titleEl?.value.trim() || '–';

  // Debug-Ausgabe
  console.log('antragCols:', antragCols ? [...antragCols].sort().join(', ') : 'null (kein Filter)');
  console.log('STEP 1 – POST fields:', JSON.stringify({ Title: titleVal }));
  console.log('STEP 2 – PATCH fields:', JSON.stringify(detailFields));

  try {
    // ── Schritt 1: Item mit nur Title erstellen ──────────────────────
    // Minimal-POST vermeidet 500 durch ungültige Feldnamen
    const newItem = await gPost(`/sites/${siteId}/lists/${listAntragId}/items`,
      { fields: { Title: titleVal } });

    if (!newItem?.id) throw new Error('Kein Item-ID in der Antwort');

    // ── Schritt 2: Details + Status per PATCH setzen ─────────────────
    const patchPayload = { ...detailFields };
    if (colOk(COL.status)) patchPayload[COL.status] = 'Eingereicht';

    try {
      await gPatch(`/sites/${siteId}/lists/${listAntragId}/items/${newItem.id}/fields`,
        patchPayload);
    } catch(ePatch) {
      // PATCH fehlgeschlagen — Item existiert, aber Details fehlen
      console.warn('Detail-PATCH fehlgeschlagen:', ePatch.message,
        '\nPayload:', JSON.stringify(patchPayload));
      // Zumindest Status versuchen
      if (colOk(COL.status)) {
        try {
          await gPatch(`/sites/${siteId}/lists/${listAntragId}/items/${newItem.id}/fields`,
            { [COL.status]: 'Eingereicht' });
        } catch(e2) { console.warn('Status-only-PATCH auch fehlgeschlagen:', e2.message); }
      }
      showAntragError('Antrag erstellt, aber Details konnten nicht gespeichert werden: ' + ePatch.message);
    }

    $id('form-antrag').reset();
    const s = $id('antrag-success');
    s.textContent = '✓ Ihr Antrag wurde eingereicht. Das KI-Koordinierungsgremium wird ihn prüfen und Sie per E-Mail informieren.';
    s.classList.remove('hidden');
    allAntraege = [];
    updateOpenBadge();

  } catch(err) {
    console.error('Antrag-Submit fehlgeschlagen:', err.message);
    showAntragError('Fehler beim Erstellen: ' + err.message);
  }

  btn.disabled = false; btn.textContent = 'Antrag einreichen';
}

function showAntragError(msg) {
  removeAntragError();
  const el = document.createElement('div');
  el.id = 'antrag-err';
  el.style.cssText = 'color:#dc2626;background:#fef2f2;border:1px solid #fca5a5;padding:10px 14px;border-radius:6px;margin-top:12px;font-size:.85rem';
  el.textContent = '✕ ' + msg;
  $id('btn-submit').after(el);
}
function removeAntragError() {
  $id('antrag-err')?.remove();
  $id('antrag-success')?.classList.add('hidden');
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
    const data = await gGet(`/sites/${siteId}/lists/${listAntragId}/items?$expand=fields($select=*)&$top=999`);
    // Client-seitig sortieren — vermeidet 400 bei nicht-indizierten Feldern
    allAntraege = (data.value || []).sort((a, b) => {
      const da = new Date(a.fields?.Created || a.createdDateTime || 0);
      const db = new Date(b.fields?.Created || b.createdDateTime || 0);
      return db - da;
    });
    renderAntraege();
    updateOpenBadge();
  } catch(e) {
    $id('antraege-loading').textContent = 'Fehler beim Laden: ' + e.message;
    console.error('loadAntraege:', e);
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

    const gesamt  = parseInt(f[COL.lizenzGesamt]) || 0;
    const users   = parseLizenzUsers(f[COL.nutzer] || '');
    const belegt  = users.length || parseInt(f[COL.lizenzBelegt]) || 0;
    const pct     = gesamt > 0 ? Math.min(100, Math.round(belegt / gesamt * 100)) : null;
    const barCls  = pct === null ? '' : pct >= 90 ? 'util-full' : pct >= 70 ? 'util-warn' : 'util-ok';
    const util    = pct !== null ? `<div style="display:flex;align-items:center;gap:6px">
      <div class="util-bar-wrap"><div class="util-bar ${barCls}" style="width:${pct}%"></div></div>
      <span style="font-size:11px;color:#6b7280">${belegt}/${gesamt}</span>
    </div>` : (belegt ? `<span style="font-size:12px;color:#6b7280">${belegt} User</span>` : '–');

    return `<tr onclick="openLizenzModal(${i.id})">
      <td><strong>${esc(f[COL.kiSystem] || f.Title || '–')}</strong></td>
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
          <th>KI-System</th><th>Typ</th><th>Anbieter</th>
          <th>Kosten</th><th>Auslastung</th><th>Vertragsende</th><th>Auto-Renewal</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

// ─── Lizenz Modal ────────────────────────────────────────────────
function parseLizenzUsers(str) {
  if (!str) return [];
  return str.split(';').map(s => s.trim()).filter(Boolean);
}
function serializeLizenzUsers(arr) {
  return arr.filter(Boolean).join('; ');
}
function renderLizenzUserEditor() {
  const listEl  = $id('lz-user-list');
  const countEl = $id('lz-user-count');
  if (!listEl) return;
  const gesamt     = parseInt($id(`lf-${COL.lizenzGesamt}`)?.value) || 0;
  const verfuegbar = gesamt > 0 ? Math.max(0, gesamt - lizenzUsers.length) : null;
  listEl.innerHTML = lizenzUsers.length === 0
    ? '<div style="color:#9ca3af;font-size:13px;padding:6px 0">Noch keine User zugewiesen</div>'
    : lizenzUsers.map((u, i) => `
        <div style="display:flex;align-items:center;gap:8px;padding:5px 8px;background:#f9fafb;border-radius:6px;margin-bottom:4px">
          <span style="flex:1;font-size:13px">👤 ${esc(u)}</span>
          <button class="btn btn-ghost btn-sm" style="padding:2px 8px;color:#ef4444" onclick="removeLizenzUser(${i})">×</button>
        </div>`).join('');
  if (countEl) {
    countEl.textContent = lizenzUsers.length
      ? `${lizenzUsers.length} User zugewiesen${verfuegbar !== null ? ` · ${verfuegbar} von ${gesamt} verfügbar` : ''}`
      : '';
  }
}
function addLizenzUser() {
  const inp  = $id('lz-user-input');
  const name = inp?.value.trim();
  if (!name) return;
  if (!lizenzUsers.includes(name)) lizenzUsers.push(name);
  inp.value = '';
  renderLizenzUserEditor();
}
function removeLizenzUser(index) {
  lizenzUsers.splice(index, 1);
  renderLizenzUserEditor();
}

// ─── People-Autocomplete ─────────────────────────────────────────
let _peopleTimer = null;

function debounceUserSearch(q) {
  clearTimeout(_peopleTimer);
  if (!q || q.length < 2) { hidePeopleDrop(); return; }
  _peopleTimer = setTimeout(() => searchPeople(q), 300);
}

async function searchPeople(q) {
  try {
    const data = await gGet(`/me/people?$search=${encodeURIComponent(q)}&$top=8&$select=displayName,scoredEmailAddresses`);
    const people = (data.value || []).filter(p => p.displayName);
    showPeopleDrop(people);
  } catch(e) {
    console.warn('People-Suche fehlgeschlagen:', e.message);
    hidePeopleDrop();
  }
}

function showPeopleDrop(people) {
  const drop = $id('lz-people-drop');
  if (!drop) return;
  if (!people.length) { drop.style.display = 'none'; return; }
  drop.innerHTML = people.map(p => {
    const mail  = p.scoredEmailAddresses?.[0]?.address || '';
    const label = mail
      ? `${esc(p.displayName)} <span style="color:#9ca3af;font-size:11px">${esc(mail)}</span>`
      : esc(p.displayName);
    const val   = mail || p.displayName;
    // JSON.stringify liefert "...", dessen " im HTML-Attribut escaped werden müssen
    const safeVal = JSON.stringify(val).replace(/"/g, '&quot;');
    return `<div class="people-item"
      onmousedown="selectPerson(${safeVal})"
      onmouseover="this.classList.add('people-item-hover')"
      onmouseout="this.classList.remove('people-item-hover')"
    >👤 ${label}</div>`;
  }).join('');
  drop.style.display = 'block';
}

function hidePeopleDrop() {
  const drop = $id('lz-people-drop');
  if (drop) drop.style.display = 'none';
}

function selectPerson(val) {
  const inp = $id('lz-user-input');
  if (inp) inp.value = val;
  hidePeopleDrop();
  // User direkt hinzufügen
  addLizenzUser();
}

function openLizenzModal(itemId) {
  editLizenzId = itemId || null;
  const item = itemId ? allLizenzen.find(i => i.id == itemId) : null;
  const f = item?.fields || {};

  // Init user list from stored value
  lizenzUsers = parseLizenzUsers(f[COL.nutzer] || '');

  // Neue Lizenz: Verantwortlich IT = aktuell angemeldeter User
  if (!itemId && !f[COL.verantwIT]) {
    f[COL.verantwIT] = account?.name || account?.username || '';
  }

  $id('modal-title').textContent = itemId ? 'Lizenz bearbeiten' : 'Neue Lizenz erfassen';

  const USER_SECTION = `
    <div style="grid-column:1/-1;margin-top:4px;border-top:1px solid #e5e9ef;padding-top:14px">
      <div style="font-weight:600;font-size:.875rem;color:#374151;margin-bottom:10px">👥 KI-User (Lizenznehmer)</div>
      <div id="lz-user-list"></div>
      <div style="display:flex;gap:8px;margin-top:10px;align-items:flex-start">
        <div style="position:relative;flex:1">
          <input id="lz-user-input" type="text" class="form-control" autocomplete="off"
            placeholder="Name oder E-Mail eingeben…" style="width:100%"
            oninput="debounceUserSearch(this.value)"
            onblur="setTimeout(hidePeopleDrop,200)"
            onkeydown="if(event.key==='Enter'){event.preventDefault();addLizenzUser();}if(event.key==='Escape'){hidePeopleDrop();}">
          <div id="lz-people-drop" style="display:none;position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #d1d5db;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.12);z-index:9999;max-height:220px;overflow-y:auto;margin-top:3px"></div>
        </div>
        <button class="btn btn-primary btn-sm" style="white-space:nowrap" onclick="addLizenzUser()">+ Hinzufügen</button>
      </div>
      <div id="lz-user-count" style="font-size:12px;color:#6b7280;margin-top:6px"></div>
    </div>`;

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
    } else if (field.type === 'combo') {
      const dlId = `dl-lf-${field.key}`;
      html += `<input id="lf-${field.key}" type="text" list="${dlId}" class="form-control" value="${esc(v)}" placeholder="Auswählen oder eingeben…">
        <datalist id="${dlId}">${field.choices.map(c => `<option value="${esc(c)}">`).join('')}</datalist>`;
    } else {
      // KI-System beim Bearbeiten readonly (eindeutiger Schlüssel, darf nicht geändert werden)
      const isLocked = itemId && field.key === COL.kiSystem;
      html += `<input id="lf-${field.key}" type="${field.type}" class="form-control" value="${esc(v)}"
        ${field.key === COL.lizenzGesamt ? ' oninput="renderLizenzUserEditor()"' : ''}
        ${isLocked ? ' readonly style="background:#f3f4f6;cursor:not-allowed" title="KI-System kann nach dem Anlegen nicht mehr geändert werden"' : ''}/>`;
    }
    html += '</div>';

    // KI-User direkt nach "Lizenzen gesamt" einblenden
    if (field.key === COL.lizenzGesamt) html += USER_SECTION;
  }
  html += '</div>';

  html += `<div class="modal-footer">
    ${itemId ? `<button class="btn btn-danger btn-sm" onclick="deleteLizenz(${itemId})">Löschen</button><span style="flex:1"></span>` : ''}
    <button class="btn btn-neutral btn-sm" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary btn-sm" onclick="saveLizenz()">Speichern</button>
  </div>`;

  $id('modal-body').innerHTML = html;
  renderLizenzUserEditor();
  $id('modal-overlay').classList.remove('hidden');
}

function showLizenzError(msg) {
  let el = $id('lizenz-save-err');
  if (!el) {
    el = document.createElement('div');
    el.id = 'lizenz-save-err';
    el.style.cssText = 'color:#dc2626;background:#fef2f2;border:1px solid #fca5a5;padding:9px 13px;border-radius:6px;margin:0 0 10px;font-size:.83rem';
    $id('modal-body')?.querySelector('.modal-footer')?.before(el);
  }
  el.textContent = '✕ ' + msg;
  el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
function hideLizenzError() { $id('lizenz-save-err')?.remove(); }

async function saveLizenz() {
  hideLizenzError();
  const kiSysEl  = $id(`lf-${COL.kiSystem}`);
  const kiSysVal = kiSysEl?.value.trim();
  if (!kiSysVal) {
    showLizenzError('Bitte KI-System eingeben.');
    kiSysEl?.focus();
    return;
  }

  // Eindeutigkeit prüfen – darf nur einmal existieren
  if (!editLizenzId) {
    const dup = allLizenzen.find(i =>
      (i.fields?.[COL.kiSystem] || i.fields?.Title || '').toLowerCase() === kiSysVal.toLowerCase()
    );
    if (dup) {
      showLizenzError(`Ein KI-System mit dem Namen „${kiSysVal}" existiert bereits.`);
      kiSysEl?.focus();
      return;
    }
  }

  const fields = { Title: kiSysVal }; // Title = KI-System (SP requires Title)

  for (const f of LIZENZ_FIELDS) {
    const el = $id(`lf-${f.key}`);
    if (!el) continue;
    const v = el.value.trim();
    if (v !== '') fields[f.key] = spValue(f.type === 'combo' ? 'text' : f.type, v);
  }

  // Users → serialized string + auto-calculate belegt
  fields[COL.nutzer]      = serializeLizenzUsers(lizenzUsers);
  fields[COL.lizenzBelegt] = lizenzUsers.length;

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
    // Kein alert – Modal bleibt offen, Fehlermeldung inline, Formulardaten bleiben erhalten
    showLizenzError('Speichern fehlgeschlagen: ' + e.message);
    console.error('saveLizenz:', e);
  }
}

async function deleteLizenz(itemId) {
  const item = allLizenzen.find(i => i.id == itemId);
  const name = item?.fields?.[COL.kiSystem] || item?.fields?.Title || 'diese Lizenz';
  if (!confirm(`Lizenz "${name}" wirklich löschen?`)) return;
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
    const data = await gGet(`/sites/${siteId}/lists/${listRegisterId}/items?$expand=fields($select=*)&$top=999`);
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
