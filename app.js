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

// KI-Register Spaltennamen (kiSystem + nutzer werden in boot() dynamisch aufgelöst)
const COL_REG = {
  status:        'Status',
  risiko:        'Risikokategorie',
  verantw:       'VerantwortlicheStelle',
  hersteller:    'Hersteller',
  nutzungsart:   'InterneExterneNutzung',
  freigabeDatum: 'FreigabeDatum',
  anbieter:      'Anbieter',
  notizen:       'Notizen',
  kiSystem:      null,   // z.B. 'KI_x002d_System' – wird in boot() aufgelöst
  nutzer:        null,   // Person-Feld "User/Nutzer" (Lizenznehmer) – wird in boot() aufgelöst
  ansprechperson: null,  // Person-Feld "Ansprechperson" = Antragsteller – wird in boot() aufgelöst
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
  zugewieseneNutzer: null,
  genehmiger:       null,   // Person-Mehrfachauswahl "Genehmiger" in KI_Antraege – wird in boot() aufgelöst
};

const STATUS_OPTS = ['Eingereicht', 'In Prüfung', 'Genehmigt', 'Abgelehnt', 'Rückfrage'];

const ANLAGE_ROLLEN = ['Legal', 'Datenschutz', 'Compliance', 'IT', 'User', 'Sonstiges'];
const ROLLE_COLORS = {
  'Legal':      { bg: '#eff6ff', color: '#1d4ed8', border: '#bfdbfe' },
  'Datenschutz':{ bg: '#fdf4ff', color: '#7e22ce', border: '#e9d5ff' },
  'Compliance': { bg: '#fff7ed', color: '#c2410c', border: '#fed7aa' },
  'IT':         { bg: '#f0fdf4', color: '#15803d', border: '#bbf7d0' },
  'User':       { bg: '#f0f9ff', color: '#0369a1', border: '#bae6fd' },
  'Sonstiges':  { bg: '#f9fafb', color: '#374151', border: '#e5e7eb' },
};

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
    hint: 'Eigene Einschätzung gemäß EU AI Act. „Verboten" = Systeme nach Art. 5 EU AI Act (z.B. Social Scoring, manipulative KI, biometrische Massenüberwachung in der Öffentlichkeit).' },
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
  { key: COL.autoRenewal,    label: 'Auto-Renewal',          type: 'yesno',    req: false,
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
let isAdmin   = false;   // Nur administrator@dihag.com → Einstellungen-Tab
const ADMIN_UPN = 'administrator@dihag.com';
let allAntraege = [], allLizenzen = [], allRegister = [];
let currentView = 'antraege';
let editLizenzId = null;
// lizenzUsers: [{name: string, email: string, spId: number|null}]
let lizenzUsers = [];
let spUserMap    = {};   // email.toLowerCase()/name → SP-LookupId (Integer)
// Gültige schreibbare Spaltennamen der Listen (wird in boot() befüllt)
let antragCols   = null;
let registerCols = null;
let lizenzCols  = null;  // analog für KI_Lizenzen
let _cacheTs = { antraege: 0, lizenzen: 0, register: 0 };
const CACHE_TTL = 5 * 60 * 1000;  // 5 Minuten

// ═══════════════════════════════════════════════════════════════════
// MSAL AUTH
// ═══════════════════════════════════════════════════════════════════
const SCOPES = [
  'https://graph.microsoft.com/Sites.ReadWrite.All',
  'User.Read',
  'People.Read',
  'User.ReadBasic.All',  // für /users?$filter=startswith(...) Personensuche
  'Mail.Send',           // E-Mails direkt über Graph senden
];

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

// Optionaler SharePoint-REST-Token – versucht mehrere Scopes (wie im Ticketsystem)
// _spTokenAvailable: null=unbekannt, true=funktioniert, false=nicht verfügbar (kein Retry)
let _spTokenAvailable = null;
async function tryGetSpToken() {
  if (_spTokenAvailable === false) return null;   // bereits bekannt: SP-Scope nicht registriert
  for (const scope of [
    `https://${SP_HOST}/AllSites.FullControl`,
    `https://${SP_HOST}/Sites.ReadWrite.All`,
    `https://${SP_HOST}/AllSites.Write`,
  ]) {
    try {
      const tok = (await msalInstance.acquireTokenSilent({ scopes: [scope], account })).accessToken;
      _spTokenAvailable = true;
      return tok;
    } catch(e) { /* nächsten Scope versuchen */ }
  }
  // Alle Scopes fehlgeschlagen → SP-REST nicht verfügbar, kein weiterer Versuch nötig
  _spTokenAvailable = false;
  return null;
}

// SP-User via ensureUser (SharePoint REST) auflösen → LookupId
// Wichtig: logonName (Kleinbuchstabe n) + odata=verbose wie im SP-REST-Standard
async function ensureSpUserViaRest(email) {
  const token = await tryGetSpToken();
  if (!token) return null;
  try {
    const res = await fetch(
      `https://${SP_HOST}${SP_SITE_PATH}/_api/web/ensureUser`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': 'noreply',   // verhindert CSRF-Fehler in manchen SP-Konfigurationen
        },
        body: JSON.stringify({ logonName: `i:0#.f|membership|${email}` })
      }
    );
    if (!res.ok) {
      const errText = await res.text().catch(() => res.status);
      console.warn('ensureUser HTTP-Fehler:', res.status, errText);
      return null;
    }
    const d = await res.json();
    // odata=verbose: Antwort liegt in d.d
    const u = d.d ?? d;
    const id = u.Id ?? u.id ?? null;
    if (id) { seedSpUser(id, email, '', u.Title || ''); return id; }
  } catch(e) { console.warn('ensureUser fehlgeschlagen:', email, e.message); }
  return null;
}

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
  return (res.status === 204 || res.status === 202) ? null : res.json();
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

    const lA  = findList(LIST_ANTRAEGE);
    const lL  = findList(LIST_LIZENZEN);
    const lR  = findList(LIST_REGISTER);

    // Set IDs upfront before parallel tasks
    listAntragId  = lA?.id || KNOWN_ANTRAEGE_GUID;
    if (lA) console.log('✓ KI_Antraege:', lA.displayName, listAntragId);
    else    console.log('⚠ KI_Antraege nicht via Listen-API gefunden → Fallback-GUID:', listAntragId);
    if (lL) listLizenzId   = lL.id;
    if (lR) listRegisterId = lR.id;

    // Parallel: Spalten-Discovery für alle drei Listen
    await Promise.all([
      // ── Antraege-Spalten ──────────────────────────────────────────
      (async () => {
        try {
          const colData = await gGet(`/sites/${siteId}/lists/${listAntragId}/columns?$select=name,displayName,readOnly,hidden&$top=200`);
          const antragColArr = colData.value || [];
          antragCols = new Set(
            antragColArr.filter(c => !c.readOnly && !c.hidden).map(c => c.name)
          );
          for (const col of antragColArr) {
            const dn = (col.displayName || '').toLowerCase().trim();
            if (dn === 'genehmiger' || dn === 'approver' || dn === 'approvers') {
              COL.genehmiger = col.name;
              console.log('✓ Genehmiger-Spalte aufgelöst:', col.name);
            }
          }
          console.log('✓ Antraege-Spalten:', [...antragCols].sort().join(', '));
        } catch(e) {
          console.warn('Spaltenabruf fehlgeschlagen (kein Filter aktiv):', e.message);
        }
      })(),

      // ── Lizenzen: Zugriff prüfen + Spalten-Discovery ──────────────
      (async () => {
        if (!lL) {
          console.warn('⚠ KI_Lizenzen nicht gefunden. Verfügbare Listen:',
            allLists.map(l => `${l.name}/${l.displayName}`).join(' | '));
          return;
        }
        console.log('✓ KI_Lizenzen:', lL.displayName, listLizenzId);
        try {
          await gGet(`/sites/${siteId}/lists/${listLizenzId}/items?$top=1`);
          isGremium = true;
          console.log('✓ Gremium-Zugriff bestätigt');

          try {
            const lizColData = await gGet(`/sites/${siteId}/lists/${listLizenzId}/columns?$select=name,displayName,readOnly,hidden&$top=200`);
            lizenzCols = new Set(
              (lizColData.value || []).filter(c => !c.readOnly && !c.hidden).map(c => c.name)
            );
            console.log('✓ Lizenzen-Spalten:', [...lizenzCols].sort().join(', '));
            console.log('Lizenzen-Spalten (alle):', (lizColData.value || []).map(c => `${c.name}="${c.displayName}"`).join(' | '));

            for (const col of (lizColData.value || [])) {
              const dn = (col.displayName || '').toLowerCase().trim();
              if (dn === 'system' || dn === 'ki-system' || dn === 'kisystem') {
                const oldKey = COL.kiSystem;
                COL.kiSystem = col.name;
                const lf = LIZENZ_FIELDS.find(f => f.key === oldKey);
                if (lf) lf.key = col.name;
                console.log('✓ KI-System Spalte aufgelöst:', oldKey, '→', col.name);
              }
              if (dn === 'ki-user' || dn === 'ki user' || dn === 'kiuser' || dn === 'ki_user') {
                COL.nutzer = col.name;
                console.log('✓ KI-User Spalte aufgelöst:', col.name);
              }
              if (dn === 'notizen' || dn === 'notes' || dn === 'bemerkungen') {
                const oldNot = COL.notizen;
                COL.notizen = col.name;
                const lf = LIZENZ_FIELDS.find(f => f.key === oldNot);
                if (lf) lf.key = col.name;
                console.log('✓ Notizen Spalte aufgelöst:', oldNot, '→', col.name);
              }
              if (dn === 'zugewiesene nutzer' || dn === 'zugewiesene_nutzer' || dn === 'zugewiesenenutzer') {
                COL.zugewieseneNutzer = col.name;
                console.log('✓ ZugewieseneNutzer Spalte aufgelöst:', col.name);
              }
            }
          } catch(eCols) {
            console.warn('Lizenzen Spalten-Lookup fehlgeschlagen:', eCols.message);
          }
        } catch(e) {
          if (e.status !== 403) console.warn('Lizenzen Lesezugriff:', e.message);
          else console.log('ℹ Kein Gremium-Zugriff auf Lizenzen (403)');
        }
      })(),

      // ── Register-Spalten-Discovery ────────────────────────────────
      (async () => {
        if (!lR) {
          console.warn('⚠ KI_Register nicht gefunden');
          return;
        }
        console.log('✓ KI_Register:', lR.displayName, listRegisterId);
        try {
          const regColData = await gGet(`/sites/${siteId}/lists/${listRegisterId}/columns?$select=name,displayName,readOnly,hidden&$top=200`);
          registerCols = new Set((regColData.value || []).filter(c => !c.readOnly && !c.hidden).map(c => c.name));
          console.log('Register-Spalten (alle):', (regColData.value || []).map(c => `${c.name}="${c.displayName}"`).join(' | '));
          for (const col of (regColData.value || [])) {
            const dn = (col.displayName || '').toLowerCase().trim();
            if (dn === 'ki-system' || dn === 'ki system' || dn === 'kisystem' || dn === 'system') {
              COL_REG.kiSystem = col.name;
              console.log('✓ Register KI-System Spalte:', col.name);
            }
            if (dn === 'nutzer' || dn === 'benutzer' || dn === 'person' || dn === 'mitarbeiter' ||
                dn === 'ki-user' || dn === 'ki user' || dn === 'kiuser' || dn === 'user' || dn === 'users') {
              COL_REG.nutzer = col.name;
              console.log('✓ Register Nutzer-Spalte:', col.name);
            }
            if (dn === 'ansprechperson' || dn === 'ansprechpartner' || dn === 'kontakt' ||
                dn === 'contact' || dn === 'antragsteller') {
              COL_REG.ansprechperson = col.name;
              console.log('✓ Register Ansprechperson-Spalte:', col.name);
            }
          }
        } catch(e) { console.warn('Register-Spalten fehlgeschlagen:', e.message); }
      })(),
    ]);

    $id('boot').style.display = 'none';
    $id('app').style.display  = 'flex';

    const uName = account?.name || account?.username || '';
    $id('user-name').textContent = uName;

    // Admin-Flag: nur administrator@dihag.com darf Einstellungen sehen
    isAdmin = (account?.username || '').toLowerCase() === ADMIN_UPN.toLowerCase();

    // Tab-Sichtbarkeit: Gremium sieht alles, normale User nur Antrag + eigene Anträge
    if (isGremium) {
      $id('gremium-badge').classList.remove('hidden');
    } else {
      // Nicht-Gremium: Lizenzen, Register und Einstellungen ausblenden
      document.querySelector('[data-view="lizenzen"]').style.display = 'none';
      document.querySelector('[data-view="register"]').style.display = 'none';
      // Filter-Toolbar auf Anträge-View ausblenden (sehen nur eigene → kein Filter nötig)
      const toolbar = document.querySelector('#view-antraege .toolbar');
      if (toolbar) toolbar.style.display = 'none';
    }
    // Einstellungen nur für Admin
    if (isAdmin) {
      $id('tab-einstellungen').style.display = '';
    }
    // Register-Tab auch für Gremium ausblenden wenn Liste nicht gefunden
    if (!listRegisterId) {
      document.querySelector('[data-view="register"]').style.display = 'none';
    }

    renderAntragForm();

    // Standardansicht: Gremium → Anträge, normale User → Neuer Antrag
    const deepId = new URLSearchParams(location.search).get('antrag');
    if (deepId) {
      await switchView('antraege');   // Deep-Link immer auf Anträge
    } else if (isGremium) {
      await switchView('antraege');
    } else {
      await switchView('antrag');
    }

    // SP-User-Map: Stufe 2 – aus Author/Editor vorhandener Items befüllen
    // (funktioniert auch wenn UserInfo-Liste nicht erreichbar war)
    try {
      const seedData = await gGet(
        `/sites/${siteId}/lists/${listAntragId}/items` +
        `?$select=id&$expand=fields($select=Author0LookupId,Author0EMail,Author0LookupValue,` +
        `Editor0LookupId,Editor0EMail,Editor0LookupValue)&$top=100`
      );
      for (const item of (seedData.value || [])) {
        const f = item.fields || {};
        if (f.Author0LookupId) seedSpUser(f.Author0LookupId, f.Author0EMail || '', '', f.Author0LookupValue || '');
        if (f.Editor0LookupId) seedSpUser(f.Editor0LookupId, f.Editor0EMail || '', '', f.Editor0LookupValue || '');
      }
      console.log('✓ SP-User-Map nach Seeding:', Object.keys(spUserMap).length, 'Einträge');
    } catch(e) { console.warn('SP-User-Map Seeding fehlgeschlagen:', e.message); }

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
  // Zugriffsschutz: Nicht-Gremium darf nur antrag + antraege
  if (!isGremium && !['antrag', 'antraege'].includes(view)) return;
  // Einstellungen nur für Admin
  if (view === 'einstellungen' && !isAdmin) return;

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

  if (view === 'antraege' && (Date.now() - _cacheTs.antraege > CACHE_TTL)) await loadAntraege();
  if (view === 'lizenzen' && (Date.now() - _cacheTs.lizenzen > CACHE_TTL)) await loadLizenzen();
  if (view === 'register' && (Date.now() - _cacheTs.register > CACHE_TTL)) await loadRegister();
  if (view === 'einstellungen') renderEinstellungen();
}

// ═══════════════════════════════════════════════════════════════════
// ANTRAG FORM
// ═══════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════
// KI-RICHTLINIE MODAL (CO-10-01 – Entwurf Mai 2026)
// ═══════════════════════════════════════════════════════════════════
function openRichtlinieModal() {
  $id('modal-title').textContent = 'KI-Richtlinie CO-10-01 – Nutzung von KI-Anwendungen (Entwurf)';
  $id('modal-card').classList.add('modal-wide');
  $id('modal-body').innerHTML = `
<div class="rdoc">

  <div class="rdoc-meta-table">
    <table>
      <tr><td>Nummerierung</td><td><strong>CO-10-01</strong></td><td>Kurzbezeichnung</td><td><strong>KI-Richtlinie</strong></td></tr>
      <tr><td>Version</td><td>1.0 (Entwurf)</td><td>Verantwortlich</td><td>Chief Compliance Officer</td></tr>
      <tr><td>Genehmiger</td><td colspan="3">Dr. Alex Lissitsa · Viktor Babushchak · Dr. Hans-Jürgen Brenninger · Robert Schüller</td></tr>
      <tr><td>Geltungsbereich</td><td colspan="3">Geschäftsleitung, Führungskräfte, alle Mitarbeitenden der DIHAG-Gruppe</td></tr>
      <tr><td>Übergeordnet</td><td colspan="3">CO-01-01 Code of Conduct</td></tr>
    </table>
  </div>

  <h3>1 · Allgemeines, Geltungsbereich</h3>

  <h4>1.1 Begriffsdefinitionen</h4>
  <p>Die EU KI-Verordnung unterscheidet zwei wesentliche Begriffe:</p>
  <ul>
    <li><strong>KI-System</strong>: maschinengestütztes System, das autonom operiert, nach Inbetriebnahme anpassungsfähig sein kann und aus Eingaben Ausgaben (Vorhersagen, Inhalte, Empfehlungen, Entscheidungen) ableitet.</li>
    <li><strong>KI-Modell mit allgemeinem Verwendungszweck (GPAI)</strong>: erhebliche allgemeine Verwendbarkeit, kann ein breites Aufgabenspektrum erfüllen und in viele Systeme integriert werden – betrifft u. a. große Sprachmodelle (LLM) wie ChatGPT.</li>
  </ul>
  <p><strong>Anbieter</strong>: Person/Unternehmen, das ein KI-System entwickelt und unter eigenem Namen in Verkehr bringt. Wenn DIHAG-Gruppe KI-Anwendungen vermarktet, gelten die strengeren Anbieter-Anforderungen.</p>
  <p><strong>Betreiber</strong>: Person/Unternehmen, das ein KI-System zu geschäftlichen Zwecken einsetzt. Die DIHAG-Gruppe ist bei geschäftlichem Einsatz Betreiber und muss die Betreiber-Anforderungen erfüllen.</p>
  <p><strong>Personenbezogene Daten</strong>: alle Informationen, die sich auf eine identifizierte oder identifizierbare natürliche Person beziehen (Namen, Adressen, E-Mails, IP-Adressen, Benutzerkennungen …).</p>
  <p><strong>Geschäftsgeheimnisse</strong>: nicht offenkundige, einem begrenzten Personenkreis bekannte Tatsachen, an deren Geheimhaltung der Arbeitgeber ein berechtigtes wirtschaftliches Interesse hat.</p>

  <h4>1.2 Zweck</h4>
  <p>KI kann Arbeitsabläufe vereinfachen, birgt aber auch Risiken (Weiterverwendung von Eingaben, fehlerhafte Ausgaben). Diese Richtlinie regelt den Umgang mit KI-Systemen, gewährleistet die Einhaltung regulatorischer Anforderungen und minimiert Risiken unsachgemäßer Nutzung.</p>

  <h4>1.3 Geltungsbereich</h4>
  <p>DIHAG Holding GmbH und alle Tochtergesellschaften sowie Beteiligungsunternehmen, bei denen DIHAG &gt; 50 % der Anteile oder Stimmrechte hält bzw. beherrschenden Einfluss hat.</p>

  <h3>2 · Regulatorisches Umfeld</h3>

  <h4>2.1 EU KI-Verordnung (in Kraft seit 01.08.2024)</h4>
  <p>Risikobasierter Ansatz:</p>
  <ul>
    <li><span class="rdoc-badge rdoc-gering">Minimales Risiko</span> – z. B. Spamfilter, KI-Videospiele: keine besonderen Verpflichtungen.</li>
    <li><span class="rdoc-badge rdoc-normal">Transparenzpflichten</span> – Chatbots müssen als Maschine erkennbar sein; KI-erzeugte Inhalte sind zu kennzeichnen.</li>
    <li><span class="rdoc-badge rdoc-hoch">Hohes Risiko</span> – z. B. KI-Medizinsoftware, Personalrekrutierung: strenge Anforderungen (Risikominderung, Datensätze, menschliche Aufsicht).</li>
    <li><span class="rdoc-badge rdoc-verboten">Verboten</span> – Social Scoring, manipulative KI, biometrische Massenüberwachung in der Öffentlichkeit.</li>
  </ul>
  <p><strong>Wichtige Termine:</strong></p>
  <ul>
    <li>ab 02.02.2025: KI-Kompetenz im Unternehmen &amp; Verbot bestimmter KI-Anwendungen</li>
    <li>ab August 2025: Bestimmungen für Anbieter von GPAI-Modellen</li>
    <li>ab August 2026: Bestimmungen für Betreiber von GPAI-Modellen</li>
  </ul>

  <h4>2.2 Datenschutz</h4>
  <p>Bei Verarbeitung personenbezogener Daten in KI-Systemen sind DS-GVO und BDSG einzuhalten. Verantwortlich: betrieblicher Datenschutzbeauftragter.</p>

  <h4>2.3 Geschäftsgeheimnisse</h4>
  <p>Vertrauliche Informationen dürfen nur im zulässigen Rahmen in KI-Systemen verwendet werden. Verantwortlich: jeweilige Fachabteilung (ggf. Unterstützung durch CCO).</p>

  <h3>3 · Verwendung von KI-Anwendungen bei der DIHAG-Gruppe</h3>

  <h4>3.1 Grundsätze zum Einsatz von KI</h4>
  <p>Die DIHAG-Gruppe unterstützt grundsätzlich den Einsatz von KI-Systemen für geschäftliche Zwecke. Diese Richtlinie gibt einen Orientierungsrahmen für eine funktionale, ethische und rechtskonforme Nutzung. Die Richtlinie wird regelmäßig evaluiert und bei Bedarf aktualisiert.</p>

  <h4>3.2 KI-Koordinierungsgremium</h4>
  <p>Das <strong>KI-Koordinierungsgremium</strong> begleitet die Einführung und Verwendung von KI. Es besteht aus verantwortlichen Vertretern der Bereiche <em>Compliance, Datenschutz, IT und Legal</em> und wird durch die Leiterin Legal geleitet.</p>

  <h4>3.3 Antrag, Genehmigung, Zweck <span class="rdoc-highlight">← relevant für diesen Antrag</span></h4>
  <p><strong>KI-Systeme dürfen nur nach vorheriger Freigabe durch das KI-Koordinierungsgremium eingesetzt werden.</strong></p>
  <p>Geplante KI-Systeme werden von der Fachabteilung mit dem <em>Antrag zur Freigabe eines KI-Systems</em> (Anlage 1) beschrieben und dem Gremium frühzeitig vorgelegt. Das Gremium prüft die Risikokategorie, gibt die Anwendung frei und legt Rahmenbedingungen fest.</p>
  <p>Bereits freigegebene Systeme sind im Intranet zu finden. Bei der Nutzung sind deren Nutzungsbedingungen und Zweckbestimmungen zu beachten.</p>

  <h4>3.4 Verwendung von Trainingsdaten</h4>
  <p>Auch die Verwendung von Trainingsdaten bedarf der Genehmigung des KI-Koordinierungsgremiums. Personenbezogene Daten dürfen nur mit entsprechender Rechtsgrundlage verwendet werden.</p>

  <h4>3.5 Mitbestimmung</h4>
  <p>Der Konzernbetriebsrat ist frühzeitig einzubinden, sofern Mitbestimmungstatbestände nach BetrVG erfüllt sind.</p>

  <h4>3.6 Grundsätze für den Umgang mit KI-erzeugten Daten</h4>
  <ul>
    <li>Verantwortung für KI-Ausgaben trägt die Person, die die Daten eingibt.</li>
    <li>Ausgaben sind stets kritisch auf Korrektheit zu prüfen; im Zweifel 4-Augen-Prinzip.</li>
    <li>Keine automatisierten Entscheidungen über Personen ohne menschliche Prüfung.</li>
    <li>KI-erzeugte Texte, Dokumente oder Bilder sind als solche zu kennzeichnen.</li>
    <li>Keine rechtswidrige oder ethisch unzulässige Nutzung (Manipulation, Diskriminierung, Überwachung ohne Rechtsgrundlage).</li>
    <li>Auffällige Feststellungen (z. B. Halluzinationen) unverzüglich der zuständigen Stelle melden.</li>
  </ul>

  <h4>3.7 KI-Kompetenz &amp; Schulungen</h4>
  <p>Betreiber sind gesetzlich verpflichtet (Art. 4 EU KI-VO), sicherzustellen, dass ihr Personal über ausreichende KI-Kompetenz verfügt. Schulungsmaßnahmen werden vom KI-Koordinierungsgremium koordiniert und im Steckbrief festgehalten. DIHAG bildet einen <strong>Key User</strong> aus, der kompetenzvermittelnde Maßnahmen weitergibt.</p>

  <h4>3.8 Software mit KI-Bausteinen</h4>
  <p>Bei Beschaffung und Einsatz sonstiger Software, die KI-Bausteine enthält (z. B. CRM- oder ERP-Systeme), ist das KI-Koordinierungsgremium einzubinden.</p>

  <h4>3.9 Freie KI-Anwendungen / Open Source</h4>
  <p class="rdoc-warning">⚠ Die Nutzung freier KI-Anwendungen ist grundsätzlich <strong>verboten</strong>!</p>

  <h3>4 · Verstoß gegen die Richtlinie</h3>
  <p>Verstöße können arbeits-, zivil- und ggf. strafrechtliche Konsequenzen nach sich ziehen. Sanktionen richten sich nach Schwere, Häufigkeit und Vorsatz. Bei Zweifeln an der Zulässigkeit sind Mitarbeitende verpflichtet, vorab <em>Datenschutz, IT oder Compliance</em> zu konsultieren. Eine frühzeitige Meldung wirkt sanktionsmildernd.</p>

  <h3>Anlage 1 · Felder des Antragsformulars (Referenz)</h3>
  <table class="rdoc-anlage">
    <thead><tr><th>Feld</th><th>Erläuterung</th></tr></thead>
    <tbody>
      <tr><td>Bezeichnung des KI-Systems</td><td>Interne Bezeichnung des KI-Use-Case bzw. der Software</td></tr>
      <tr><td>Verantwortliche Stelle</td><td>Wer verantwortet die geplante Lösung im Betrieb?</td></tr>
      <tr><td>Hersteller / Entwickler</td><td>Bezugsquelle, Dienstleister oder Lieferant des KI-Systems</td></tr>
      <tr><td>KI-Komponente(n)</td><td>Beschreibung der Funktionen, integrierten KI-Modelle und Verfahren</td></tr>
      <tr><td>Verwendungszweck laut Hersteller</td><td>Wie definiert der Hersteller den Zweck? (Nutzungsbedingungen, Dokumentation)</td></tr>
      <tr><td>Anwendungsbereich im Unternehmen</td><td>Zu welchem Zweck intern einsetzen? Abweichung vom Herstellerzweck?</td></tr>
      <tr><td>Risikokategorie</td><td>Geringes · Normales · Hohes Risiko · Verboten – mit Begründung gemäß EU AI Act</td></tr>
      <tr><td>Nutzungsart</td><td>Nur intern oder auch externes Angebot / Vermarktung? (Abgrenzung Betreiber/Anbieter)</td></tr>
      <tr><td>Geplanter Einsatz ab</td><td>Zeitplan für den Einsatz des KI-Systems</td></tr>
      <tr><td>Key User / Schulungsmaßnahme</td><td>Wer ist Key User? Welche kompetenzvermittelnden Maßnahmen sind geplant?</td></tr>
    </tbody>
  </table>

  <p style="margin-top:20px;font-size:.78rem;color:#9ca3af;border-top:1px solid #e5e9ef;padding-top:12px">
    Entwurf Stand 05.05.2026 · Ansprechpartner: Karl Würz, ext. Chief Compliance Officer (wuerz@dihag.com)
  </p>
</div>`;
  $id('modal-overlay').classList.remove('hidden');
}

function closeModal(e) {
  if (e && e.target !== $id('modal-overlay')) return;
  $id('modal-overlay').classList.add('hidden');
  $id('modal-card').classList.remove('modal-wide');
}

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
  let firstInvalid = null;
  document.querySelectorAll('#form-antrag-fields [required]').forEach(el => {
    el.classList.remove('invalid');
    if (!el.value.trim()) {
      el.classList.add('invalid');
      if (!firstInvalid) firstInvalid = el;
      valid = false;
    }
  });
  if (!valid) {
    if (firstInvalid) {
      firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
      setTimeout(() => firstInvalid.focus(), 300);
    }
    return;
  }

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

    // Genehmiger-Feld aus Einstellungen befüllen (Person-Mehrfachauswahl)
    if (COL.genehmiger && colOk(COL.genehmiger)) {
      try {
        const _genList = loadSettings().genehmiger || [];
        const genIds = [];
        for (const g of _genList) {
          const id = await resolveSpUserId(g.email, g.name);
          if (id) genIds.push(id);
        }
        if (genIds.length) {
          patchPayload[COL.genehmiger + 'LookupId@odata.type'] = 'Collection(Edm.Int32)';
          patchPayload[COL.genehmiger + 'LookupId'] = genIds;
          console.log('✓ Genehmiger-LookupIds gesetzt:', genIds);
        }
      } catch(eGen) { console.warn('Genehmiger-LookupId fehlgeschlagen:', eGen.message); }
    }

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
    s.innerHTML = `✓ Ihr Antrag <strong>${esc(titleVal)}</strong> wurde eingereicht. Das KI-Koordinierungsgremium wird ihn prüfen.`;

    // Genehmiger automatisch per Graph-Mail benachrichtigen (wenn konfiguriert)
    // Antragsteller aus der Empfängerliste ausschließen – kein Self-Notify
    const _st = loadSettings();
    const _gen = _st.genehmiger || [];
    const _myEmail = (account?.username || '').toLowerCase();
    const _genToNotify = _gen.filter(g => g.email.toLowerCase() !== _myEmail);
    if (_st.benachrichtigung?.beiEinreichung !== false && _genToNotify.length) {
      const sender  = account?.name || account?.username || 'Antragsteller';
      const deepUrl = `${location.origin}${location.pathname}?antrag=${newItem.id}`;
      sendMail(
        _genToNotify.map(g => ({ address: g.email, name: g.name })),
        `[KI-Antrag] ${titleVal} – Prüfung erforderlich`,
        mailTemplate(
          'Neuer KI-Antrag zur Prüfung eingegangen',
          [
            ['Bezeichnung',    titleVal],
            ['Antragsteller',  sender],
            ['Eingereicht am', new Date().toLocaleDateString('de-DE')],
          ],
          '🔍 Antrag direkt öffnen',
          deepUrl
        )
      ).then(() => showToast('📧 Genehmiger wurden automatisch benachrichtigt.'))
       .catch(e => {
         console.warn('Mail an Genehmiger fehlgeschlagen:', e.message);
         showToast('Antrag eingereicht – E-Mail-Benachrichtigung fehlgeschlagen (' + e.message + ')', 'error', 7000);
       });
    }
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
    // Nicht-Gremium: nur eigene Items laden (server-seitig filtern → kein Fremddaten-Zugriff)
    let apiUrl = `/sites/${siteId}/lists/${listAntragId}/items?$expand=fields($select=*)&$top=999`;
    if (!isGremium) {
      const myEmailFilter = (account?.username || '').toLowerCase();
      if (myEmailFilter) {
        apiUrl += `&$filter=createdBy/user/email eq '${encodeURIComponent(myEmailFilter)}'`;
      }
    }
    let data;
    try {
      data = await gGet(apiUrl);
    } catch(filterErr) {
      // Fallback: Graph-Filter nicht unterstützt → alles laden, client-seitig filtern
      console.warn('Server-Filter nicht unterstützt, Fallback auf client-seitigen Filter:', filterErr.message);
      data = await gGet(`/sites/${siteId}/lists/${listAntragId}/items?$expand=fields($select=*)&$top=999`);
    }
    // Client-seitig sortieren — vermeidet 400 bei nicht-indizierten Feldern
    allAntraege = (data.value || []).sort((a, b) => {
      const da = new Date(a.fields?.Created || a.createdDateTime || 0);
      const db = new Date(b.fields?.Created || b.createdDateTime || 0);
      return db - da;
    });
    _cacheTs.antraege = Date.now();
    renderAntraege();
    updateOpenBadge();

    // Deep-Link: ?antrag=ID → Antrag-Panel direkt öffnen
    const deepId = new URLSearchParams(location.search).get('antrag');
    if (deepId) {
      const target = allAntraege.find(i => String(i.id) === String(deepId));
      if (target) {
        openAntragPanel(target.id);
        // URL sauber halten – ID aus der Adresszeile entfernen
        history.replaceState({}, '', location.pathname);
      }
    }
  } catch(e) {
    $id('antraege-loading').textContent = 'Fehler beim Laden: ' + e.message;
    console.error('loadAntraege:', e);
  }
}

function filterAntraege() { renderAntraege(); }

function renderAntraege() {
  const statusF = $id('filter-status')?.value || '';
  const riskF   = $id('filter-risk')?.value   || '';
  const searchQ = ($id('search-antraege')?.value || '').toLowerCase().trim();

  // Non-Gremium: nur eigene Anträge anzeigen
  const myEmail = (account?.username || account?.idTokenClaims?.preferred_username || '').toLowerCase();

  let items = allAntraege.filter(i => {
    const f = i.fields;
    if (statusF && f[COL.status] !== statusF) return false;
    if (riskF   && f[COL.risiko] !== riskF)  return false;
    if (!isGremium && myEmail && (f.Author0EMail || '').toLowerCase() !== myEmail) return false;
    if (searchQ) {
      const hay = [f.Title, f[COL.hersteller], f[COL.verantw], f[COL.komponenten]].filter(Boolean).join(' ').toLowerCase();
      if (!hay.includes(searchQ)) return false;
    }
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

  // Genehmiger-Namen aus Person-Feld auslesen
  const genehmigerNames = (() => {
    if (!COL.genehmiger || !f[COL.genehmiger]) return '';
    const pf = f[COL.genehmiger];
    if (Array.isArray(pf)) return pf.map(g => g?.LookupValue || String(g)).filter(Boolean).join(', ');
    if (typeof pf === 'object') return pf?.LookupValue || '';
    return String(pf);
  })();

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
      ${genehmigerNames ? row('Genehmiger', esc(genehmigerNames)) : ''}
    </div>
    <div class="panel-section">
      <div class="panel-section-title">KI-Beschreibung</div>
      ${row('KI-Komponenten',             esc(f[COL.komponenten]), true)}
      ${row('Zweck laut Hersteller',      esc(f[COL.zweckHersteller]), true)}
      ${row('Anwendungsbereich intern',   esc(f[COL.zweckUnternehmen]), true)}
    </div>`;

  // APPROVALS-Token aus Kommentar für die Anzeige entfernen
  const kommentarClean = (f[COL.gremiumKommentar] || '').replace(/\[APPROVALS:[^\]]*\]\n?/g, '').trim();

  // Für alle User: Status-Bereich (read-only für normale User)
  const statusSection = !isGremium ? `
    <div class="panel-section">
      <div class="panel-section-title">Status</div>
      ${row('Aktueller Status', statusBadge(f[COL.status]))}
      ${kommentarClean ? row('Gremium-Kommentar', esc(kommentarClean), true) : ''}
      ${f[COL.auflagen]         ? row('Auflagen',           esc(f[COL.auflagen]), true) : ''}
      ${f[COL.freigabeDatum]    ? row('Freigabedatum',      fmtDate(f[COL.freigabeDatum])) : ''}
    </div>` : '';

  // Rückfrage-Antwort-Sektion für nicht-Gremium-User
  const rueckfrageSection = (!isGremium && f[COL.status] === 'Rückfrage') ? `
    <div style="background:#faf5ff;border:1.5px solid #d8b4fe;border-radius:10px;padding:16px;margin-top:12px">
      <div style="font-size:.8rem;font-weight:700;color:#7e22ce;text-transform:uppercase;letter-spacing:.5px;margin-bottom:12px">💬 Rückfrage des Gremiums</div>
      ${f[COL.gremiumKommentar] ? `<div style="background:#fff;border:1px solid #e9d5ff;border-radius:7px;padding:10px 12px;font-size:.85rem;color:#1e2939;white-space:pre-wrap;margin-bottom:14px">${esc(f[COL.gremiumKommentar])}</div>` : ''}
      <div class="form-group" style="margin-bottom:10px">
        <label class="form-label">Ihre Antwort</label>
        <textarea id="rueck-antwort" class="form-control" rows="3" placeholder="Bitte beantworten Sie die Rückfrage hier…"></textarea>
      </div>
      <button class="btn btn-primary btn-sm" onclick="submitRueckfrageAntwort(${item.id})">Antwort senden</button>
    </div>` : '';

  const isDecided    = ['Genehmigt', 'Abgelehnt'].includes(f[COL.status]);

  // Einstimmig-Modus: Abstimmungsstand aus GremiumKommentar lesen
  const _stPanel       = loadSettings();
  const einstimmig     = (_stPanel.benachrichtigung?.genehmigungsmodus || 'einstimmig') === 'einstimmig';
  const _genPanel      = _stPanel.genehmiger || [];
  const panelApprovals = parseApprovals(f[COL.gremiumKommentar]);
  const myEmailPanel   = (account?.username || '').toLowerCase();

  // Self-Approval-Guard: Gremium-User ist zugleich der Antragsteller → keine Aktionsbuttons
  const antragAuthorEmail  = (f.Author0EMail || '').toLowerCase();
  const isOwnAntrag        = myEmailPanel && antragAuthorEmail && myEmailPanel === antragAuthorEmail;

  // Effektive Genehmiger = konfigurierte Liste ohne den Antragsteller (darf nicht selbst zustimmen)
  const effectiveGenehmiger  = _genPanel.filter(g => g.email.toLowerCase() !== antragAuthorEmail);
  const myApprovedAlready    = panelApprovals.includes(myEmailPanel);
  const showApprovalTracker  = !isDecided && !isOwnAntrag && einstimmig && effectiveGenehmiger.length >= 1;

  // Wiederverwendbarer Read-Only-Block (entschieden oder eigener Antrag)
  const decidedBlock = `
    <div style="margin-bottom:12px">
      ${statusBadge(f[COL.status])}
      ${f[COL.freigabeDatum] ? `<span style="font-size:.8rem;color:#6b7280;margin-left:8px">📅 ${fmtDate(f[COL.freigabeDatum])}</span>` : ''}
    </div>
    ${kommentarClean ? `
      <div style="margin-bottom:10px">
        <div class="panel-field-label">Begründung</div>
        <div class="panel-field-value pre" style="background:#f9fafb;padding:8px 10px;border-radius:7px;border:1px solid #e5e9ef">${esc(kommentarClean)}</div>
      </div>` : ''}
    ${f[COL.auflagen] ? `
      <div>
        <div class="panel-field-label">Auflagen / Bedingungen</div>
        <div class="panel-field-value pre" style="background:#fffbeb;padding:8px 10px;border-radius:7px;border:1px solid #fde68a">${esc(f[COL.auflagen])}</div>
      </div>` : ''}`;

  const gremiumSection = isGremium ? `
    <div class="panel-gremium">
      <div class="panel-gremium-title">⚖️ Gremium-Entscheidung</div>
      ${isOwnAntrag ? `
        <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:12px 14px;margin-bottom:12px;font-size:.82rem;color:#92400e">
          ℹ️ <strong>Eigener Antrag</strong> – du bist der Antragsteller und kannst diesen Antrag nicht selbst genehmigen.
        </div>
        ${decidedBlock}
      ` : isDecided ? `
        ${decidedBlock}
        <div style="margin-top:14px">
          <button class="btn btn-neutral btn-sm" onclick="saveGremiumDecision(${item.id},'Eingereicht')" title="Entscheidung zurücksetzen">↩ Zurücksetzen</button>
        </div>
      ` : `
        ${showApprovalTracker ? `
        <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:12px 14px;margin-bottom:14px;font-size:.82rem">
          <div style="font-weight:600;color:#15803d;margin-bottom:8px">⚖️ Einstimmig – Zustimmungsstand</div>
          ${effectiveGenehmiger.map(g => {
            const approved = panelApprovals.includes(g.email.toLowerCase());
            const rc2 = ROLLE_COLORS[g.rolle] || null;
            const rolleBadge2 = g.rolle && rc2
              ? `<span style="font-size:.65rem;font-weight:600;padding:1px 7px;border-radius:20px;background:${rc2.bg};color:${rc2.color};border:1px solid ${rc2.border}">${esc(g.rolle)}</span>`
              : '';
            return `<div style="display:flex;align-items:center;gap:6px;padding:2px 0">
              <span style="color:${approved ? '#15803d' : '#9ca3af'};font-size:1rem">${approved ? '✓' : '○'}</span>
              <span style="${approved ? 'color:#15803d;font-weight:500' : 'color:#6b7280'}">${esc(g.name || g.email)}</span>
              ${rolleBadge2}
            </div>`;
          }).join('')}
        </div>` : ''}
        <div class="form-group">
          <label class="form-label">Kommentar / Begründung <span style="color:#6b7280;font-weight:400">(optional)</span></label>
          <textarea id="pg-kommentar" class="form-control" rows="3" placeholder="Begründung der Entscheidung…">${esc(kommentarClean)}</textarea>
        </div>
        <div class="form-group">
          <label class="form-label">Auflagen / Bedingungen</label>
          <textarea id="pg-auflagen" class="form-control" rows="2" placeholder="Ggf. Auflagen oder Bedingungen…">${esc(f[COL.auflagen] || '')}</textarea>
        </div>
        <div class="panel-actions">
          <button class="btn btn-success btn-sm" onclick="saveGremiumDecision(${item.id},'Genehmigt')">${showApprovalTracker && !myApprovedAlready ? '✓ Zustimmen' : showApprovalTracker && myApprovedAlready ? '✓ Bereits zugestimmt' : '✓ Genehmigen'}</button>
          <button class="btn btn-danger btn-sm"  onclick="saveGremiumDecision(${item.id},'Abgelehnt')">✕ Ablehnen</button>
          <button class="btn btn-neutral btn-sm" onclick="saveGremiumDecision(${item.id},'Rückfrage')">? Rückfrage</button>
          <button class="btn btn-neutral btn-sm" onclick="saveGremiumDecision(${item.id},'${f[COL.status] || 'In Prüfung'}')">💾 Kommentar speichern</button>
        </div>
      `}
    </div>` : '';

  // Anhänge-Platzhalter (wird async befüllt)
  const attachSection = `
    <div class="panel-section" id="panel-attachments">
      <div class="panel-section-title">📎 Anhänge</div>
      <div id="att-list" class="att-list"><span style="color:#9ca3af;font-size:.8rem">Lade Anhänge…</span></div>
      ${isGremium ? `
        <div id="att-drop" class="att-drop" ondragover="attDragOver(event)" ondragleave="attDragLeave(event)" ondrop="attDrop(event,${item.id})">
          <span class="att-drop-icon">📁</span>
          <span>Datei hierher ziehen oder</span>
          <label class="att-drop-btn">
            Datei auswählen
            <input type="file" id="att-file-input" multiple style="display:none" onchange="attFileSelect(event,${item.id})">
          </label>
        </div>
      ` : ''}
    </div>`;

  $id('panel-body').innerHTML = rows1 + statusSection + rueckfrageSection + gremiumSection + attachSection;
  openPanel();

  // Anhänge asynchron nachladen (ohne await – kein Blockieren)
  renderAttachments(item.id);
}

async function saveGremiumDecision(itemId, forceStatus) {
  const status    = forceStatus || allAntraege.find(i => i.id == itemId)?.fields?.[COL.status] || 'In Prüfung';
  const kommentar = $id('pg-kommentar')?.value?.trim() || '';
  const auflagen  = $id('pg-auflagen')?.value?.trim()  || '';

  // Self-Approval-Guard: eigene Anträge können nicht selbst genehmigt/abgelehnt werden
  const prevItemCheck = allAntraege.find(i => i.id == itemId);
  const antragAuthorG = (prevItemCheck?.fields?.Author0EMail || '').toLowerCase();
  const myEmailG      = (account?.username || '').toLowerCase();
  if (myEmailG && antragAuthorG && myEmailG === antragAuthorG &&
      (status === 'Genehmigt' || status === 'Abgelehnt' || status === 'Rückfrage')) {
    showToast('Eigene Anträge können nicht selbst genehmigt werden.', 'error');
    return;
  }



  const prevItem = allAntraege.find(i => i.id == itemId);
  const prevKommentar = prevItem?.fields?.[COL.gremiumKommentar] || '';
  const now = new Date().toLocaleDateString('de-DE', {day:'2-digit',month:'2-digit',year:'numeric'}) + ' ' +
              new Date().toLocaleTimeString('de-DE', {hour:'2-digit', minute:'2-digit'});
  let newKommentar = prevKommentar;
  if (kommentar) {
    newKommentar = prevKommentar
      ? `[${now}] ${kommentar}\n──\n${prevKommentar}`
      : `[${now}] ${kommentar}`;
  }

  // ── Einstimmig-Modus: Teilzustimmung verfolgen ───────────────────
  if (status === 'Genehmigt') {
    const _stD = loadSettings();
    const modus = _stD.benachrichtigung?.genehmigungsmodus || 'einstimmig';
    if (modus === 'einstimmig') {
      // Effektive Genehmiger: Antragsteller aus der Pflicht-Zustimmungsliste ausschließen
      const _genDAll  = _stD.genehmiger || [];
      const _genD     = _genDAll.filter(g => g.email.toLowerCase() !== antragAuthorG);
      if (_genD.length >= 1) {
        const prevKomRaw = prevItem?.fields?.[COL.gremiumKommentar] || '';
        const approvals  = parseApprovals(prevKomRaw);
        const myEmailD   = (account?.username || '').toLowerCase();
        if (!approvals.includes(myEmailD)) approvals.push(myEmailD);
        const allApproved = _genD.every(g => approvals.includes(g.email.toLowerCase()));
        if (!allApproved) {
          // Noch nicht alle zugestimmt → Zwischenspeichern, Status auf 'In Prüfung' halten
          const appToken  = `[APPROVALS:${approvals.join('|')}]`;
          const cleanBase = prevKomRaw.replace(/\[APPROVALS:[^\]]*\]\n?/g, '').trim();
          let partialKom = cleanBase;
          if (kommentar) {
            partialKom = cleanBase
              ? `[${now}] ${kommentar}\n──\n${cleanBase}`
              : `[${now}] ${kommentar}`;
          }
          partialKom = partialKom ? `${appToken}\n${partialKom}` : appToken;
          try {
            await gPatch(`/sites/${siteId}/lists/${listAntragId}/items/${itemId}/fields`,
              { [COL.status]: 'In Prüfung', [COL.gremiumKommentar]: partialKom });
            const idx2 = allAntraege.findIndex(i => i.id == itemId);
            if (idx2 >= 0) Object.assign(allAntraege[idx2].fields, { [COL.status]: 'In Prüfung', [COL.gremiumKommentar]: partialKom });
          } catch(ePart) { console.warn('Einstimmig-PATCH fehlgeschlagen:', ePart.message); }

          const remaining = _genD.filter(g => !approvals.includes(g.email.toLowerCase()));
          const remainingNames = remaining.map(g => g.name || g.email);
          showToast(`✓ Deine Zustimmung gespeichert. Noch ausstehend: ${remainingNames.join(', ')}`);

          // Verbleibende Genehmiger per Mail benachrichtigen (außer dem Antragsteller)
          if (remaining.length && _stD.benachrichtigung?.beiEinreichung !== false) {
            const antragTitle = prevItem?.fields?.Title || '';
            const approverName = account?.name || account?.username || '';
            sendMail(
              remaining.map(g => ({ address: g.email, name: g.name })),
              `[KI-Antrag] ${antragTitle} – Zustimmung ausstehend`,
              mailTemplate(
                'Zustimmung zu einem KI-Antrag ausstehend',
                [
                  ['KI-System',    antragTitle],
                  ['Zugestimmt',   approverName],
                  ['Ausstehend',   remainingNames.join(', ')],
                  ['Datum',        new Date().toLocaleDateString('de-DE')],
                ],
                '✓ Jetzt zustimmen',
                `${location.origin}${location.pathname}?antrag=${itemId}`
              )
            ).catch(e => console.warn('Genehmiger-Reminder fehlgeschlagen:', e.message));
          }

          renderAntraege();
          updateOpenBadge();
          closePanel();
          return;
        }
        // Alle haben zugestimmt → APPROVALS-Token aus Kommentar entfernen
        newKommentar = newKommentar.replace(/\[APPROVALS:[^\]]*\]\n?/g, '').trim();
      }
    }
  }

  const fields = { [COL.status]: status };
  if (newKommentar) fields[COL.gremiumKommentar] = newKommentar;
  if (auflagen)  fields[COL.auflagen]         = auflagen;
  if (status === 'Genehmigt') fields[COL.freigabeDatum] = new Date().toISOString().slice(0, 10);

  try {
    await gPatch(`/sites/${siteId}/lists/${listAntragId}/items/${itemId}/fields`, fields);
    const idx = allAntraege.findIndex(i => i.id == itemId);
    if (idx >= 0) Object.assign(allAntraege[idx].fields, fields);
    const antragAfter = allAntraege.find(i => i.id == itemId);
    const savedName   = antragAfter?.fields?.Title || '';
    showToast(`✓ Entscheidung „${status}" gespeichert${savedName ? ' für ' + savedName : ''}.`);

    // Antragsteller automatisch per Graph-Mail benachrichtigen (wenn konfiguriert + Statusentscheidung)
    if ((status === 'Genehmigt' || status === 'Abgelehnt' || status === 'Rückfrage') && antragAfter) {
      const authorEmail = antragAfter.fields?.Author0EMail || '';
      const authorName  = antragAfter.fields?.Author0LookupValue || authorEmail;
      const _st2 = loadSettings();
      if (_st2.benachrichtigung?.beiEntscheidung !== false && authorEmail) {
        const statusEmoji = status === 'Genehmigt' ? '✅' : status === 'Abgelehnt' ? '❌' : '❓';
        const deepUrl = `${location.origin}${location.pathname}?antrag=${antragAfter.id}`;
        const infoRows = [
          ['KI-System',     savedName],
          ['Entscheidung',  `${statusEmoji} ${status}`],
          ['Entscheider',   account?.name || account?.username || ''],
          ['Datum',         new Date().toLocaleDateString('de-DE')],
        ];
        if (kommentar) infoRows.push(['Begründung', kommentar]);
        if (auflagen)  infoRows.push(['Auflagen / Bedingungen', auflagen]);

        sendMail(
          [{ address: authorEmail, name: authorName }],
          `[KI-Antrag] ${savedName} – ${status}`,
          mailTemplate(
            `Ihr KI-Antrag wurde ${status === 'Genehmigt' ? 'genehmigt' : status === 'Abgelehnt' ? 'abgelehnt' : 'mit einer Rückfrage versehen'}`,
            infoRows,
            status === 'Rückfrage' ? '💬 Rückfrage beantworten' : '📋 Antrag im Dashboard anzeigen',
            deepUrl
          )
        ).then(() => showToast(`📧 ${authorName || authorEmail} automatisch benachrichtigt.`))
         .catch(e => {
           console.warn('Mail an Antragsteller fehlgeschlagen:', e.message);
           showToast('Entscheidung gespeichert – E-Mail fehlgeschlagen: ' + e.message, 'error', 7000);
         });
      }
    }

    // Bei Genehmigung: automatisch Draft-Lizenz erstellen (falls noch keine existiert)
    if (status === 'Genehmigt' && listLizenzId) {
      const antrag = allAntraege.find(i => i.id == itemId);
      const systemName = antrag?.fields?.Title;
      if (systemName) {
        try {
          // Prüfen ob Lizenz schon vorhanden (lokal + ggf. remote)
          let lizenzen = allLizenzen;
          if (!lizenzen.length) {
            const d = await gGet(`/sites/${siteId}/lists/${listLizenzId}/items?$expand=fields($select=Title,${COL.kiSystem})&$top=999`);
            lizenzen = d.value || [];
          }
          const exists = lizenzen.some(l =>
            (l.fields?.[COL.kiSystem] || l.fields?.Title || '').toLowerCase() === systemName.toLowerCase()
          );
          if (!exists) {
            const newLiz = await gPost(`/sites/${siteId}/lists/${listLizenzId}/items`,
              { fields: { Title: systemName } });
            if (newLiz?.id) {
              const draftNote = '⚠ Automatisch erstellt – bitte Lizenzdetails ergänzen';
              const patchUrl  = `/sites/${siteId}/lists/${listLizenzId}/items/${newLiz.id}/fields`;

              // COL.kiSystem ist ein Lookup auf KI_Register → kein Text-Write möglich.
              // Der Lookup wird gesetzt, sobald der Register-Eintrag bei Lizenzierung erstellt wird.
              // Title reicht für die Anzeige in der Zwischenzeit.

              // Notizen setzen – mehrere Kandidaten durchprobieren
              let notesSaved = false;
              for (const nKey of [COL.notizen, 'Notizen', 'Notes', 'Bemerkungen']) {
                if (lizenzCols && !lizenzCols.has(nKey)) continue;
                try {
                  await gPatch(patchUrl, { [nKey]: draftNote });
                  COL.notizen = nKey; // korrekten Namen merken
                  notesSaved  = true;
                  break;
                } catch(e) { console.warn('Auto-Lizenz Notizen (', nKey, '):', e.message); }
              }
              if (!notesSaved) console.warn('Auto-Lizenz: Notizen-Spalte konnte nicht gesetzt werden');
            }
            allLizenzen = []; // Lizenz-Tab beim nächsten Öffnen neu laden
            console.log('✓ Draft-Lizenz automatisch erstellt:', systemName);
          }
        } catch(eLiz) {
          console.warn('Auto-Lizenz Erstellung fehlgeschlagen:', eLiz.message);
        }
      }
    }

    closePanel();
    renderAntraege();
    updateOpenBadge();
  } catch(e) {
    showToast('Fehler beim Speichern: ' + e.message, 'error');
  }
}

async function submitRueckfrageAntwort(itemId) {
  const text = $id('rueck-antwort')?.value?.trim() || '';
  if (!text) {
    showToast('Bitte eine Antwort eingeben.', 'error');
    return;
  }

  const item = allAntraege.find(i => i.id == itemId);
  const prevKommentar = item?.fields?.[COL.gremiumKommentar] || '';
  const now = new Date().toLocaleDateString('de-DE', {day:'2-digit',month:'2-digit',year:'numeric'}) + ' ' +
              new Date().toLocaleTimeString('de-DE', {hour:'2-digit', minute:'2-digit'});
  const newKommentar = prevKommentar
    ? `[${now}] Antwort: ${text}\n──\n${prevKommentar}`
    : `[${now}] Antwort: ${text}`;

  const fields = {
    [COL.status]: 'Eingereicht',
    [COL.gremiumKommentar]: newKommentar,
  };

  try {
    await gPatch(`/sites/${siteId}/lists/${listAntragId}/items/${itemId}/fields`, fields);
    const idx = allAntraege.findIndex(i => i.id == itemId);
    if (idx >= 0) Object.assign(allAntraege[idx].fields, fields);
    closePanel();
    renderAntraege();
    updateOpenBadge();
    showToast('Antwort eingereicht.');
  } catch(e) {
    showToast('Fehler beim Senden: ' + e.message, 'error');
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
    _cacheTs.lizenzen = Date.now();

    // spUserMap aus vorhandenen Lizenz-Einträgen befüllen (KI-User-Feld enthält LookupIds)
    for (const item of allLizenzen) {
      const f = item.fields || {};
      const names = Array.isArray(f[COL.nutzer]) ? f[COL.nutzer] : (f[COL.nutzer] ? [f[COL.nutzer]] : []);
      const ids   = Array.isArray(f[COL.nutzer+'LookupId']) ? f[COL.nutzer+'LookupId'] : (f[COL.nutzer+'LookupId'] ? [f[COL.nutzer+'LookupId']] : []);
      names.forEach((n, i) => {
        const name  = typeof n === 'string' ? n : (n?.LookupValue || '');
        const spId  = ids[i];
        if (name && spId) seedSpUser(spId, '', '', name);
      });
    }
    console.log('✓ SP-User-Map nach Lizenzen-Seeding:', Object.keys(spUserMap).length, 'Einträge');

    renderLizenzen();
  } catch(e) {
    $id('lizenzen-loading').textContent = 'Fehler: ' + e.message;
  }
}

function renderLizenzen() {
  $id('lizenzen-loading').classList.add('hidden');

  const today = new Date();
  // Stats always use all items
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

  // Build filtered items list
  const searchQ = ($id('search-lizenzen')?.value || '').toLowerCase().trim();
  const typF    = $id('filter-lizenztyp')?.value || '';
  const ablaufF = $id('filter-ablauf')?.value    || '';

  let items = allLizenzen.filter(i => {
    const f = i.fields;
    if (searchQ) {
      const hay = [f[COL.kiSystem], f.Title, f[COL.anbieter], f[COL.lizenztyp]].filter(Boolean).join(' ').toLowerCase();
      if (!hay.includes(searchQ)) return false;
    }
    if (typF && f[COL.lizenztyp] !== typF) return false;
    if (ablaufF) {
      const ende = f[COL.vertragsEnde];
      const diff = ende ? (new Date(ende) - today) / 86400000 : null;
      if (ablaufF === 'expired') {
        if (diff === null || diff >= 0) return false;
      } else {
        const days = parseInt(ablaufF);
        if (diff === null || diff < 0 || diff > days) return false;
      }
    }
    return true;
  });

  const rows = items.map(i => {
    const f      = i.fields;
    const ende   = f[COL.vertragsEnde];
    const diff   = ende ? (new Date(ende) - today) / 86400000 : null;
    const endeCls = diff === null ? 'expiry-ok' : diff < 0 ? 'expiry-alert' : diff < 30 ? 'expiry-alert' : diff < 60 ? 'expiry-warn' : 'expiry-ok';
    const endeLabel = ende ? fmtDate(ende) : '–';

    const gesamt  = parseInt(f[COL.lizenzGesamt]) || 0;
    const users   = parseLizenzUsersWithIds(f);
    const belegt  = users.length || parseInt(f[COL.lizenzBelegt]) || 0;
    const pct     = gesamt > 0 ? Math.min(100, Math.round(belegt / gesamt * 100)) : null;
    const barCls  = pct === null ? '' : pct >= 90 ? 'util-full' : pct >= 70 ? 'util-warn' : 'util-ok';
    const util    = pct !== null ? `<div style="display:flex;align-items:center;gap:6px">
      <div class="util-bar-wrap"><div class="util-bar ${barCls}" style="width:${pct}%"></div></div>
      <span style="font-size:11px;color:#6b7280">${belegt}/${gesamt}</span>
    </div>` : (belegt ? `<span style="font-size:12px;color:#6b7280">${belegt} User</span>` : '–');

    const isDraft = (f[COL.notizen] || '').startsWith('⚠ Automatisch erstellt');
    const rowStyle = isDraft ? ' style="opacity:.55"' : '';
    const draftBadge = isDraft ? ' <span style="font-size:10px;color:#9ca3af;font-weight:400">(Entwurf)</span>' : '';
    return `<tr onclick="openLizenzModal(${i.id})"${rowStyle}>
      <td><strong>${esc(f[COL.kiSystem] || f.Title || '–')}</strong>${draftBadge}</td>
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

// SP-Personenfeld (Graph) → [{name, email, spId}]
function parseLizenzUsers(val) {
  if (!val) return [];
  // Array: SP gibt Personenfelder als Array zurück
  if (Array.isArray(val)) {
    return val.map(u => {
      if (typeof u === 'object' && u !== null)
        return { name: u.LookupValue || u.Title || u.displayName || '', email: u.EMail || u.email || '', spId: u.LookupId || null };
      const s = String(u).trim();
      return { name: s, email: s.includes('@') ? s : '', spId: null };
    }).filter(u => u.name || u.email);
  }
  // Fallback: alter Semikolon-Text (Migrationspfad)
  if (typeof val === 'string') {
    return val.split(';').map(s => s.trim()).filter(Boolean).map(s => ({
      name: s, email: s.includes('@') ? s : '', spId: null
    }));
  }
  return [];
}

// LookupIds-Array-Variante aus fields (Graph gibt KIUserLookupId separat zurück)
function parseLizenzUsersWithIds(fields) {
  const nameVal = fields[COL.nutzer];
  const idVal   = fields[COL.nutzer + 'LookupId'];
  const names   = Array.isArray(nameVal) ? nameVal : (nameVal ? [nameVal] : []);
  const ids     = Array.isArray(idVal)   ? idVal   : (idVal   ? [idVal]   : []);
  if (!names.length && !ids.length) return [];
  return names.map((n, i) => ({
    name:  typeof n === 'string' ? n : (n?.LookupValue || String(n)),
    email: '',
    spId:  ids[i] ?? null
  }));
}

// Einen User in die Map eintragen (Email + claims-Login + Anzeigename)
function seedSpUser(id, email, loginName, displayName) {
  if (!id) return;
  const n = parseInt(id);
  if (email)       spUserMap[email.toLowerCase().trim()]          = n;
  if (loginName && loginName.includes('|'))
                   spUserMap[loginName.split('|').pop().toLowerCase().trim()] = n;
  if (displayName) spUserMap['__name__' + displayName.toLowerCase().trim()]   = n;
}

// E-Mail oder Anzeigename → SP-LookupId
// Stufe 1: spUserMap (befüllt aus Lizenzen-/Antrags-Items beim Laden)
// Stufe 2: SharePoint REST ensureUser (braucht SP-Scope in App-Registration)
async function resolveSpUserId(email, name) {
  if (email) {
    const byEmail = spUserMap[email.toLowerCase().trim()];
    if (byEmail) return byEmail;
  }
  if (name) {
    const byName = spUserMap['__name__' + name.toLowerCase().trim()];
    if (byName) return byName;
  }
  // Stufe 2: SharePoint REST ensureUser (funktioniert wenn SP-Scope in App-Registration)
  if (email) {
    const id = await ensureSpUserViaRest(email);
    if (id) return id;
  }
  console.warn('Kein SP-LookupId für:', name || email, '| Map-Größe:', Object.keys(spUserMap).length);
  return null;
}

// [{name,email,spId}] → [N, N2, …] – einfache Integer-Array für Graph Collection(Edm.Int32)
// (SP-REST erwartet {LookupId: N}, Graph erwartet plain integers + @odata.type-Annotation)
async function buildLookupIds(users) {
  const result = [];
  for (const u of users) {
    const id = u.spId || await resolveSpUserId(u.email, u.name);
    if (id) result.push(parseInt(id));
  }
  return result;
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
          <div style="flex:1;min-width:0">
            <div style="font-size:13px">👤 ${esc(u.name || u.email || '?')}</div>
            ${u.email && u.email !== u.name ? `<div style="font-size:11px;color:#9ca3af;margin-top:1px">${esc(u.email)}</div>` : ''}
          </div>
          <button class="btn btn-ghost btn-sm" style="padding:2px 8px;color:#ef4444;flex-shrink:0" onclick="removeLizenzUser(${i})">×</button>
        </div>`).join('');
  if (countEl) {
    countEl.textContent = lizenzUsers.length
      ? `${lizenzUsers.length} User zugewiesen${verfuegbar !== null ? ` · ${verfuegbar} von ${gesamt} verfügbar` : ''}`
      : '';
  }
}

function addLizenzUser() {
  const inp = $id('lz-user-input');
  const val = inp?.value.trim();
  if (!val) return;
  const isEmail = val.includes('@');
  const user = { name: val, email: isEmail ? val : '', spId: null };
  const dup  = lizenzUsers.some(u => (u.email && u.email === user.email) || u.name === user.name);
  if (!dup) lizenzUsers.push(user);
  inp.value = '';
  hidePeopleDrop();
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
  const enc = encodeURIComponent(q);
  try {
    // Graph /users: liefert displayName + mail → SP-freundlichste Option
    const data = await gGet(
      `/users?$filter=startswith(displayName,'${enc}') or startswith(mail,'${enc}')` +
      `&$select=id,displayName,mail,userPrincipalName&$top=8`
    );
    const people = (data.value || []).filter(p => p.displayName);
    if (people.length) { showPeopleDrop(people); return; }
  } catch(e) {
    console.warn('Benutzersuche (/users) fehlgeschlagen:', e.message);
  }
  // Fallback: /me/people (weniger Token-Anforderungen, aber keine Mail-Garantie)
  try {
    const data2 = await gGet(
      `/me/people?$search=${enc}&$top=8&$select=displayName,scoredEmailAddresses`
    );
    const people2 = (data2.value || [])
      .filter(p => p.displayName)
      .map(p => ({
        displayName:       p.displayName,
        mail:              p.scoredEmailAddresses?.[0]?.address || '',
        userPrincipalName: p.scoredEmailAddresses?.[0]?.address || '',
      }));
    showPeopleDrop(people2);
  } catch(e2) {
    console.warn('People-Suche fehlgeschlagen:', e2.message);
    hidePeopleDrop();
  }
}

function showPeopleDrop(people) {
  const drop = $id('lz-people-drop');
  if (!drop) return;
  if (!people.length) { drop.style.display = 'none'; return; }
  drop.innerHTML = people.map(p => {
    const mail  = p.mail || p.userPrincipalName || p.scoredEmailAddresses?.[0]?.address || '';
    const label = mail
      ? `${esc(p.displayName)} <span style="color:#9ca3af;font-size:11px">${esc(mail)}</span>`
      : esc(p.displayName);
    // data-Attribute statt Inline-String-Escaping – vermeidet SyntaxErrors
    return `<div class="people-item"
      data-name="${esc(p.displayName)}" data-mail="${esc(mail)}"
      onmousedown="selectPersonFromDrop(this)"
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

// Tickets-Muster: SP-LookupId sofort beim Auswählen aus dem Dropdown auflösen
async function selectPersonFromDrop(el) {
  const name  = el.dataset.name || '';
  const email = el.dataset.mail || '';
  hidePeopleDrop();
  const inp = $id('lz-user-input');
  if (inp) inp.value = '';

  const user = { name: name || email, email: email || '', spId: null };
  const dup  = lizenzUsers.some(u => (u.email && u.email === user.email) || u.name === user.name);
  if (dup) return;

  lizenzUsers.push(user);
  renderLizenzUserEditor();

  // SP-LookupId jetzt direkt auflösen (nicht erst beim Speichern)
  if (email) {
    try {
      const spId = await ensureSpUserViaRest(email);
      if (spId) {
        user.spId = spId;
        console.log('✓ SP-User-ID aufgelöst:', name, '→', spId);
      } else {
        console.warn('⚠ SP-User-ID nicht aufgelöst:', name, email, '– wird ggf. beim Speichern nochmals versucht');
      }
    } catch(e) {
      console.warn('ensureUser Fehler:', email, e.message);
    }
  }
}

// Fallback für addLizenzUser (manuelle Eingabe ohne Dropdown)
function selectPerson(name, email) {
  hidePeopleDrop();
  const user = { name: name || email, email: email || '', spId: null };
  const dup  = lizenzUsers.some(u => (u.email && u.email === user.email) || u.name === user.name);
  if (!dup) { lizenzUsers.push(user); renderLizenzUserEditor(); }
  const inp = $id('lz-user-input');
  if (inp) inp.value = '';
}

function openLizenzModal(itemId) {
  editLizenzId = itemId || null;
  const item = itemId ? allLizenzen.find(i => i.id == itemId) : null;
  const f = item?.fields || {};

  // Personenfeld lesen: Graph gibt Namen-Array + LookupId-Array separat zurück
  lizenzUsers = parseLizenzUsersWithIds(f);
  // Fallback: älteres Textformat (Semikolon-getrennt)
  if (!lizenzUsers.length && f[COL.nutzer]) lizenzUsers = parseLizenzUsers(f[COL.nutzer]);

  // Neue Lizenz: Verantwortlich IT = aktuell angemeldeter User
  if (!itemId && !f[COL.verantwIT]) {
    f[COL.verantwIT] = account?.name || account?.username || '';
  }

  $id('modal-title').textContent = itemId ? 'Lizenz bearbeiten' : 'Neue Lizenz erfassen';

  // War dieses Item ein Entwurf? → Notizen-Feld leer zeigen
  const isDraftItem = itemId && (f[COL.notizen] || '').startsWith('⚠ Automatisch erstellt');

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
    // Verantwortlich IT: wird beim Speichern automatisch gesetzt – nicht anzeigen
    if (field.key === COL.verantwIT) continue;

    // KI-System: Title als sicherer Fallback (wird immer korrekt gesetzt)
    // Entwurfs-Notiz: nicht vorausfüllen (leeres Feld zeigen)
    const rawVal = field.key === COL.kiSystem
      ? (f[COL.kiSystem] || f.Title || '')
      : (isDraftItem && field.key === COL.notizen)
        ? ''
        : (f[field.key] ?? '');
    // yesno-Felder: SP liefert Boolean → 'Ja'/'Nein'
    let v = spDisplayValue(field.type, rawVal);
    // date-Felder: SP liefert ISO-Datetime (z.B. "2025-06-01T00:00:00Z"),
    // <input type="date"> braucht "yyyy-MM-dd"
    if (field.type === 'date' && v) v = String(v).slice(0, 10);

    const cls = field.type === 'textarea' ? 'form-group full' : 'form-group';
    html += `<div class="${cls}">
      <label class="form-label" for="lf-${field.key}">${esc(field.label)}${field.req ? '<span class="req">*</span>' : ''}</label>`;

    if (field.type === 'textarea') {
      html += `<textarea id="lf-${field.key}" class="form-control" rows="2">${esc(v)}</textarea>`;
    } else if (field.type === 'choice' || field.type === 'yesno') {
      html += `<select id="lf-${field.key}" class="form-control">`;
      for (const c of field.choices) html += `<option value="${esc(c)}"${String(v) === String(c) ? ' selected' : ''}>${esc(c) || '– wählen –'}</option>`;
      html += '</select>';
    } else if (field.type === 'combo') {
      const dlId = `dl-lf-${field.key}`;
      html += `<input id="lf-${field.key}" type="text" list="${dlId}" class="form-control" value="${esc(v)}" placeholder="Auswählen oder eingeben…">
        <datalist id="${dlId}">${field.choices.map(c => `<option value="${esc(c)}">`).join('')}</datalist>`;
    } else {
      // KI-System beim Bearbeiten readonly (eindeutiger Schlüssel, darf nicht geändert werden)
      const isLocked = itemId && field.key === COL.kiSystem;
      html += `<input id="lf-${field.key}" type="${field.type === 'yesno' ? 'text' : field.type}" class="form-control" value="${esc(v)}"
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
    <button class="btn btn-primary btn-sm" id="lizenz-save-btn" onclick="saveLizenz()">Speichern</button>
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
  const saveBtn = $id('lizenz-save-btn');
  const setBusy = busy => {
    if (!saveBtn) return;
    saveBtn.disabled = busy;
    saveBtn.textContent = busy ? '⏳ Speichern…' : 'Speichern';
  };
  setBusy(true);

  const kiSysEl  = $id(`lf-${COL.kiSystem}`);
  const kiSysVal = kiSysEl?.value.trim();
  if (!kiSysVal) {
    showLizenzError('Bitte KI-System eingeben.');
    kiSysEl?.focus();
    setBusy(false);
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
      setBusy(false);
      return;
    }
  }

  // Spalten-Check: wenn lizenzCols geladen → nur bestätigte Spalten; sonst alles versuchen
  // (safePatch fängt 400 feldweise ab)
  const colOk = name => !lizenzCols || lizenzCols.has(name);

  // War das Item ein Entwurf (automatisch erstellt)?
  const oldLizenzItem = editLizenzId ? allLizenzen.find(i => i.id == editLizenzId) : null;
  const wasDraft = !!(oldLizenzItem?.fields?.[COL.notizen] || '').startsWith('⚠ Automatisch erstellt');

  // Detail-Felder sammeln (ohne Title und VerantwortlicherIT – werden separat gesetzt)
  const detailFields = {};
  for (const f of LIZENZ_FIELDS) {
    const el = $id(`lf-${f.key}`);
    if (!el || f.key === COL.kiSystem || f.key === COL.verantwIT) continue;
    const v = el.value.trim();
    if (v === '') continue;
    if (!colOk(f.key)) { console.warn('Lizenz-Spalte nicht in SP-Schreibliste, übersprungen:', f.key); continue; }
    const converted = spValue(f.type === 'combo' ? 'text' : f.type, v);
    // yesno: null bedeutet kein gültiger Wert → überspringen
    if (converted === null) continue;
    detailFields[f.key] = converted;
  }
  // COL.kiSystem ('System_x0028_LookupaufKI_Registe') ist ein Lookup auf KI_Register
  // → kann nicht als Text geschrieben werden; wird von createRegisterEntries via LookupId gesetzt
  // Zur Anzeige reicht Title (= Systemname), das wird separat gesetzt

  // Verantwortlich IT: wird immer auf den aktuell eingeloggten User gesetzt
  if (colOk(COL.verantwIT)) detailFields[COL.verantwIT] = account?.name || account?.username || '';

  // ZugewieseneNutzer: Textfeld mit den Namen aller zugewiesenen User befüllen
  if (COL.zugewieseneNutzer !== null && colOk(COL.zugewieseneNutzer) && lizenzUsers.length > 0) {
    detailFields[COL.zugewieseneNutzer] = lizenzUsers.map(u => u.name || u.email).filter(Boolean).join('; ');
  }

  // Entwurfs-Notiz entfernen: war das Item ein Entwurf, forcieren wir den Notizen-Wert
  // (auch wenn der User nichts eingetippt hat → leeres '' löscht die Marker-Notiz)
  if (wasDraft && colOk(COL.notizen)) {
    detailFields[COL.notizen] = $id(`lf-${COL.notizen}`)?.value.trim() || '';
  }

  // Personenfeld: LookupIds auflösen.
  // Graph nutzt 'LookupId'-Suffix für READ *und* WRITE (nicht 'Id' wie SP-REST!):
  //   "KI_x002d_UserLookupId@odata.type": "Collection(Edm.Int32)"
  //   "KI_x002d_UserLookupId": [42, 43]
  // Das ist exakt das gleiche Muster wie im Ticketsystem (stripReadOnly).
  const nutzerLookupKey  = COL.nutzer + 'LookupId';
  const nutzerOdataKey   = nutzerLookupKey + '@odata.type';
  if (lizenzUsers.length > 0) {
    const lookupIds = await buildLookupIds(lizenzUsers);
    if (lookupIds.length > 0) {
      detailFields[nutzerOdataKey]  = 'Collection(Edm.Int32)';
      detailFields[nutzerLookupKey] = lookupIds;
      console.log('SP-User LookupIds (Graph-Format):', JSON.stringify(lookupIds));
    } else {
      console.warn('Keine SP-LookupIds aufgelöst – User-Feld wird nicht gesetzt');
    }
  } else {
    // Leere Auswahl: Feld leeren
    detailFields[nutzerOdataKey]  = 'Collection(Edm.Int32)';
    detailFields[nutzerLookupKey] = [];
  }
  if (colOk(COL.lizenzBelegt)) detailFields[COL.lizenzBelegt] = lizenzUsers.length;

  console.log('saveLizenz – lizenzCols:', lizenzCols ? [...lizenzCols].sort().join(',') : 'null');
  console.log('saveLizenz – Title:', kiSysVal, '| detailFields:', JSON.stringify(detailFields));

  // Hilfsfunktion: PATCH-Versuch, bei Fehler feldweise retry
  // @odata.type-Annotationen werden immer zusammen mit dem Hauptfeld übertragen
  const safePatch = async (url, fields) => {
    if (!Object.keys(fields).length) return;
    try {
      await gPatch(url, fields);
    } catch(ePatch) {
      console.warn('Bulk-PATCH fehlgeschlagen, versuche feldweise:', ePatch.message);
      for (const [k, v] of Object.entries(fields)) {
        if (k.includes('@odata.type')) continue;  // kommt mit seinem Hauptfeld mit
        const odataKey   = k + '@odata.type';
        const fieldPatch = fields[odataKey] !== undefined
          ? { [odataKey]: fields[odataKey], [k]: v }
          : { [k]: v };
        try { await gPatch(url, fieldPatch); }
        catch(ef) { console.warn(`Feld "${k}" konnte nicht gespeichert werden:`, ef.message); }
      }
    }
  };

  try {
    if (editLizenzId) {
      // PATCH: Title immer + bestätigte Detail-Felder
      await safePatch(
        `/sites/${siteId}/lists/${listLizenzId}/items/${editLizenzId}/fields`,
        { Title: kiSysVal, ...detailFields }
      );
    } else {
      // Schritt 1: Item mit nur Title anlegen
      const newItem = await gPost(`/sites/${siteId}/lists/${listLizenzId}/items`,
        { fields: { Title: kiSysVal } });
      if (!newItem?.id) throw new Error('Kein Item-ID in der Antwort');
      // Schritt 2: Details per PATCH (mit feldweisem Fallback)
      await safePatch(
        `/sites/${siteId}/lists/${listLizenzId}/items/${newItem.id}/fields`,
        detailFields
      );
    }
    closeModal();
    allLizenzen = [];
    await loadLizenzen();

    // Wenn ein Entwurf vollständig ausgefüllt wurde → KI-Register-Eintrag anlegen + sofort laden
    if (wasDraft && listRegisterId) {
      const finalId = editLizenzId;  // editLizenzId noch gesetzt vor closeModal
      await createRegisterEntries(kiSysVal, lizenzUsers, finalId);
      allRegister = [];
      _cacheTs.register = 0;   // Cache erzwingen → nächstes Öffnen des Register-Tabs lädt neu
      if (currentView === 'register') {
        await loadRegister();   // Sofort aktualisieren wenn Register gerade geöffnet ist
        showToast('✓ KI-Register automatisch aktualisiert.');
      } else {
        showToast('✓ KI-Register wurde automatisch aktualisiert.');
      }
      console.log('✓ KI-Register nach Lizenzierung aktualisiert');
    }
  } catch(e) {
    showLizenzError('Speichern fehlgeschlagen: ' + e.message);
    console.error('saveLizenz:', e);
    setBusy(false);
  }
}

// Legt EINEN KI-Register-Eintrag pro System an (nicht pro Person).
// Die beteiligten Kollegen landen als Text in KeyUser.
// Nach der Erstellung wird das Lizenz-System-Lookup-Feld gesetzt.
async function createRegisterEntries(kiSystem, users, lizenzItemId) {
  if (!listRegisterId) return;
  const regColOk = k => k && (!registerCols || registerCols.has(k));

  // Duplikat-Prüfung: Register-Eintrag für dieses System schon vorhanden?
  const existing = allRegister.find(r =>
    (r.fields?.Title || '').toLowerCase() === kiSystem.toLowerCase()
  );
  if (existing) {
    console.log('Register-Eintrag bereits vorhanden:', kiSystem);
    // Trotzdem Lizenz-Lookup aktualisieren falls noch nicht gesetzt
    if (lizenzItemId) await updateLizenzSystemLookup(lizenzItemId, parseInt(existing.id));
    return;
  }

  // Verwandte Antrag-Daten (Hersteller, Beschreibung, Anwendungsbereiche, KeyUser, …)
  const antrag = allAntraege.find(i =>
    (i.fields?.Title || '').toLowerCase() === kiSystem.toLowerCase()
  );
  const af = antrag?.fields || {};

  try {
    // Schritt 1: Register-Item mit Title anlegen
    const regItem = await gPost(`/sites/${siteId}/lists/${listRegisterId}/items`,
      { fields: { Title: kiSystem } });
    if (!regItem?.id) return;

    // Schritt 2: Felder befüllen (nur bekannte, schreibbare Spalten)
    const pf = {};

    // Lookup auf den Antrag (AntragID-Feld)
    const antragLookupCol = 'AntragID_x0028_LookupaufKI_Antra';
    if (antrag?.id && regColOk(antragLookupCol)) {
      pf[antragLookupCol + 'LookupId'] = parseInt(antrag.id);
    }

    // Daten aus dem Antrag
    if (af[COL.hersteller]       && regColOk('Hersteller'))              pf['Hersteller']              = af[COL.hersteller];
    if (af[COL.komponenten]      && regColOk('Beschreibung'))             pf['Beschreibung']            = af[COL.komponenten];
    if (af[COL.zweckUnternehmen] && regColOk('Anwendungsbereiche'))       pf['Anwendungsbereiche']      = af[COL.zweckUnternehmen];
    if (af[COL.keyUser]          && regColOk('Schulungszielgruppe'))      pf['Schulungszielgruppe']     = af[COL.keyUser];
    if (af[COL.verantw]          && regColOk(COL_REG.verantw))           pf[COL_REG.verantw]           = af[COL.verantw];
    if (af[COL.risiko]           && regColOk(COL_REG.risiko))            pf[COL_REG.risiko]            = af[COL.risiko];
    if (af[COL.nutzungsart]      && regColOk(COL_REG.nutzungsart))       pf[COL_REG.nutzungsart]       = af[COL.nutzungsart];
    if (af[COL.freigabeDatum]    && regColOk(COL_REG.freigabeDatum))     pf[COL_REG.freigabeDatum]     = af[COL.freigabeDatum];
    // GueltigAb: Freigabedatum aus Antrag, sonst heute
    if (regColOk('GueltigAb'))   pf['GueltigAb'] = af[COL.freigabeDatum]
      ? String(af[COL.freigabeDatum]).slice(0, 10)
      : new Date().toISOString().slice(0, 10);
    if (regColOk('NaechstePruefung')) {
      const naechste = new Date();
      naechste.setFullYear(naechste.getFullYear() + 1);
      pf['NaechstePruefung'] = naechste.toISOString().slice(0, 10);
    }

    // Kollegen (KI-User) als Text in KeyUser (immer, als Fallback)
    if (users.length && regColOk('KeyUser')) {
      pf['KeyUser'] = users.map(u => u.name || u.email).filter(Boolean).join(', ');
    }

    // Nutzer als Person-Feld (Einzelauswahl) schreiben – erste Person aus der Lizenznehmerliste
    if (users.length && COL_REG.nutzer && regColOk(COL_REG.nutzer)) {
      try {
        const firstUser = users[0];
        const spId = firstUser.spId || await resolveSpUserId(firstUser.email, firstUser.name);
        if (spId) {
          pf[COL_REG.nutzer + 'LookupId'] = spId;   // Single-Value → plain integer (kein Collection)
          console.log('✓ Register Nutzer Person-LookupId:', spId);
        }
      } catch(eNU) { console.warn('Register Nutzer LookupId fehlgeschlagen:', eNU.message); }
    }

    // Antragsteller als Ansprechperson hinterlegen
    if (COL_REG.ansprechperson && regColOk(COL_REG.ansprechperson)) {
      try {
        const authorId = af.Author0LookupId
          ? parseInt(af.Author0LookupId)
          : await resolveSpUserId(af.Author0EMail || '', af.Author0LookupValue || '');
        if (authorId) {
          pf[COL_REG.ansprechperson + 'LookupId'] = authorId;
          console.log('✓ Register Ansprechperson (Antragsteller) LookupId:', authorId);
        }
      } catch(eAsp) { console.warn('Register Ansprechperson LookupId fehlgeschlagen:', eAsp.message); }
    }

    if (Object.keys(pf).length) {
      await gPatch(`/sites/${siteId}/lists/${listRegisterId}/items/${regItem.id}/fields`, pf)
        .catch(e => console.warn('Register-PATCH:', e.message));
    }

    // Schritt 3: Lizenz-System-Lookup auf den neuen Register-Eintrag setzen
    if (lizenzItemId) await updateLizenzSystemLookup(lizenzItemId, parseInt(regItem.id));

    console.log('✓ Register-Eintrag erstellt:', kiSystem, '(ID:', regItem.id, ')');
  } catch(e) {
    console.warn('Register-Eintrag fehlgeschlagen:', kiSystem, e.message);
  }
}

// Setzt das System-Lookup-Feld in der Lizenz auf einen KI_Register-Eintrag
async function updateLizenzSystemLookup(lizenzItemId, registerId) {
  if (!lizenzItemId || !registerId || !listLizenzId) return;
  // COL.kiSystem = 'System_x0028_LookupaufKI_Registe' (Lookup-Feld)
  // Schreiben: LookupId-Suffix + integer (kein Collection, da Single-Value-Lookup)
  const lookupKey = COL.kiSystem + 'LookupId';
  try {
    await gPatch(
      `/sites/${siteId}/lists/${listLizenzId}/items/${lizenzItemId}/fields`,
      { [lookupKey]: registerId }
    );
    console.log('✓ Lizenz System-Lookup gesetzt:', registerId);
  } catch(e) {
    console.warn('Lizenz System-Lookup PATCH fehlgeschlagen:', e.message);
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
    showToast('Fehler: ' + e.message, 'error');
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
    _cacheTs.register = Date.now();
    renderRegister();
  } catch(e) {
    $id('register-loading').textContent = 'Fehler: ' + e.message;
  }
}

function filterRegister() { renderRegister(); }

function renderRegister() {
  const riskF   = $id('reg-filter-risk')?.value    || '';
  const nutzerF = $id('reg-filter-nutzer')?.value  || '';
  const searchF = ($id('search-register')?.value   || '').toLowerCase().trim();

  // Hilfsfunktion: Person-Feld → lesbarer String (generisch für Nutzer und Ansprechperson)
  const getPersonText = (f, colKey, fallbackKey) => {
    if (colKey && f[colKey] != null) {
      const pf = f[colKey];
      if (Array.isArray(pf)) return pf.map(p => p?.LookupValue || String(p)).filter(Boolean).join(', ');
      if (typeof pf === 'object') return pf?.LookupValue || '';
      return String(pf);
    }
    return fallbackKey ? (f[fallbackKey] || '') : '';
  };
  const getNutzerText      = f => getPersonText(f, COL_REG.nutzer, 'KeyUser');
  const getAnsprechperson  = f => getPersonText(f, COL_REG.ansprechperson, null);

  // Nutzer-Dropdown dynamisch befüllen (Person-Feld hat Vorrang vor KeyUser-Text)
  const nutzerDrop = $id('reg-filter-nutzer');
  if (nutzerDrop) {
    const allNutzer = new Set();
    allRegister.forEach(i => {
      const f = i.fields;
      if (COL_REG.nutzer && f[COL_REG.nutzer] != null) {
        const pf = f[COL_REG.nutzer];
        const arr = Array.isArray(pf) ? pf : [pf];
        arr.forEach(p => {
          const n = p?.LookupValue || (typeof p === 'string' ? p : '');
          if (n) allNutzer.add(n);
        });
      } else {
        (f['KeyUser'] || '').split(/[,;]/).map(s => s.trim()).filter(Boolean).forEach(n => allNutzer.add(n));
      }
    });
    const currentVal = nutzerDrop.value;
    nutzerDrop.innerHTML = '<option value="">Alle Nutzer</option>' +
      [...allNutzer].sort().map(n => `<option value="${esc(n)}"${n === currentVal ? ' selected' : ''}>${esc(n)}</option>`).join('');
  }

  let items = allRegister.filter(i => {
    const f = i.fields;
    if (riskF && f[COL_REG.risiko] !== riskF) return false;
    if (nutzerF) {
      const nutzerText = getNutzerText(f);
      if (COL_REG.nutzer && f[COL_REG.nutzer] != null) {
        // Person-Feld: jeden Eintrag einzeln prüfen
        const pf = f[COL_REG.nutzer];
        const arr = Array.isArray(pf) ? pf : [pf];
        if (!arr.some(p => (p?.LookupValue || String(p)) === nutzerF)) return false;
      } else {
        const ku = nutzerText.split(/[,;]/).map(s => s.trim());
        if (!ku.includes(nutzerF)) return false;
      }
    }
    if (searchF) {
      const hay = [f.Title, f[COL_REG.verantw], f[COL_REG.hersteller], getNutzerText(f), getAnsprechperson(f)].join(' ').toLowerCase();
      if (!hay.includes(searchF)) return false;
    }
    return true;
  });

  $id('register-loading').classList.add('hidden');

  const mitFreigabe = allRegister.filter(i => !!i.fields?.GueltigAb).length;
  const stats = `<div class="stats-row">
    <div class="stat-card accent"><div class="stat-value">${allRegister.length}</div><div class="stat-label">Systeme gesamt</div></div>
    <div class="stat-card green"><div class="stat-value">${mitFreigabe}</div><div class="stat-label">Freigegeben</div></div>
    <div class="stat-card orange"><div class="stat-value">${allRegister.length - mitFreigabe}</div><div class="stat-label">In Vorbereitung</div></div>
  </div>`;

  if (!items.length) {
    $id('register-wrap').innerHTML = stats + '<div class="empty-state">Keine Einträge gefunden.</div>';
    return;
  }

  const rows = items.map(i => {
    const f = i.fields;
    const nutzerDisp     = getNutzerText(f);
    const ansprechDisp   = getAnsprechperson(f);
    return `<tr onclick="openRegisterPanel(${i.id})" style="cursor:pointer">
      <td>
        <strong>${esc(f.Title || '–')}</strong>
        ${nutzerDisp ? `<div style="font-size:11px;color:#6b7280;margin-top:2px">👤 ${esc(nutzerDisp)}</div>` : ''}
        ${ansprechDisp ? `<div style="font-size:11px;color:#9ca3af;margin-top:1px">✉ ${esc(ansprechDisp)}</div>` : ''}
      </td>
      <td>${esc(f[COL_REG.verantw] || '–')}</td>
      <td>${esc(f[COL_REG.hersteller] || f[COL_REG.anbieter] || '–')}</td>
      <td>${riskBadge(f[COL_REG.risiko])}</td>
      <td>${f[COL_REG.nutzungsart] ? `<span class="badge-type">${esc(f[COL_REG.nutzungsart])}</span>` : '–'}</td>
      <td>${fmtDate(f['GueltigAb'])}</td>
    </tr>`;
  }).join('');

  $id('register-wrap').innerHTML = stats + `<div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>KI-System / Nutzer</th><th>Verantwortl. Stelle</th><th>Hersteller</th>
          <th>Risiko</th><th>Nutzungsart</th><th>Geplant ab</th>
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

  $id('panel-title').innerHTML = `🤖 <span style="margin-left:4px">${esc(f.Title || '–')}</span>`;

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
      ${(() => {
        if (COL_REG.nutzer && f[COL_REG.nutzer] != null) {
          const pf = f[COL_REG.nutzer];
          const name = Array.isArray(pf)
            ? pf.map(p => p?.LookupValue || String(p)).filter(Boolean).join(', ')
            : (typeof pf === 'object' ? (pf?.LookupValue || '') : String(pf));
          return name ? row('Nutzer (Lizenznehmer)', esc(name)) : '';
        }
        return f['KeyUser'] ? row('Nutzer / Key User', esc(f['KeyUser'])) : '';
      })()}
      ${(() => {
        if (COL_REG.ansprechperson && f[COL_REG.ansprechperson] != null) {
          const pf = f[COL_REG.ansprechperson];
          const name = Array.isArray(pf)
            ? pf.map(p => p?.LookupValue || String(p)).filter(Boolean).join(', ')
            : (typeof pf === 'object' ? (pf?.LookupValue || '') : String(pf));
          return name ? row('Ansprechperson', esc(name)) : '';
        }
        return '';
      })()}
      ${f['GueltigAb']          ? row('Gültig ab',        fmtDate(f['GueltigAb'])) : ''}
      ${f['NaechstePruefung']   ? row('Nächste Prüfung',  fmtDate(f['NaechstePruefung'])) : ''}
      ${f['Beschreibung']       ? row('Beschreibung',     esc(f['Beschreibung']), true) : ''}
    </div>
    ${f['Schulungszielgruppe'] || f['Anwendungsbereiche'] ? `<div class="panel-section">
      <div class="panel-section-title">Einsatz & Schulung</div>
      ${f['Anwendungsbereiche']   ? row('Anwendungsbereiche', esc(f['Anwendungsbereiche']), true) : ''}
      ${f['Schulungszielgruppe']  ? row('Schulungszielgruppe', esc(f['Schulungszielgruppe']), true) : ''}
    </div>` : ''}
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
// closeModal is defined above (next to openRichtlinieModal) to support modal-wide cleanup

// ═══════════════════════════════════════════════════════════════════
// E-MAIL VIA GRAPH (Mail.Send)
// ═══════════════════════════════════════════════════════════════════
// toList: [{address, name}] oder ['email@...'] oder 'email@...'
async function sendMail(toList, subject, bodyHtml) {
  const toArr = (Array.isArray(toList) ? toList : [toList]).map(r =>
    typeof r === 'string'
      ? { emailAddress: { address: r } }
      : { emailAddress: { address: r.address || r.email || '', name: r.name || '' } }
  ).filter(r => r.emailAddress.address);

  if (!toArr.length) { console.warn('sendMail: keine Empfänger'); return; }

  await gPost('/me/sendMail', {
    message: {
      subject,
      body: { contentType: 'HTML', content: bodyHtml },
      toRecipients: toArr,
    },
    saveToSentItems: true,
  });
  console.log('✓ E-Mail gesendet:', subject, '→', toArr.map(r => r.emailAddress.address).join(', '));
}

// HTML-Template für KI-Benachrichtigungs-Mails
function mailTemplate(title, lines, ctaLabel, ctaUrl) {
  const href = ctaUrl || (location.origin + location.pathname);
  const cta = ctaLabel
    ? `<p style="margin:24px 0 0"><a href="${href}"
        style="background:#1a56db;color:#fff;padding:10px 22px;border-radius:7px;text-decoration:none;font-weight:600"
        >${ctaLabel}</a></p>`
    : '';
  const rows = lines.map(([k, v]) =>
    `<tr><td style="padding:5px 0;color:#6b7280;font-size:13px;width:160px">${k}</td>
         <td style="padding:5px 0;font-size:13px;font-weight:500">${v}</td></tr>`
  ).join('');
  return `<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;background:#f0f2f5;margin:0;padding:24px">
  <div style="max-width:560px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #e5e9ef">
    <div style="background:linear-gradient(135deg,#1648c5,#1a56db);padding:20px 28px">
      <div style="color:#fff;font-size:18px;font-weight:700">🤖 KI-Dashboard · DIHAG Gruppe</div>
    </div>
    <div style="padding:24px 28px">
      <h2 style="margin:0 0 18px;font-size:16px;color:#1e2939">${title}</h2>
      <table style="border-collapse:collapse;width:100%">${rows}</table>
      ${cta}
    </div>
    <div style="background:#f9fafb;padding:12px 28px;font-size:11px;color:#9ca3af;border-top:1px solid #e5e9ef">
      Diese Nachricht wurde automatisch vom KI-Dashboard generiert.
    </div>
  </div></body></html>`;
}

// ═══════════════════════════════════════════════════════════════════
// TOAST NOTIFICATIONS
// ═══════════════════════════════════════════════════════════════════
function showToast(msg, type = 'success', duration = 4000) {
  let container = $id('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = msg;
  container.appendChild(toast);
  requestAnimationFrame(() => { requestAnimationFrame(() => toast.classList.add('toast-show')); });
  setTimeout(() => {
    toast.classList.remove('toast-show');
    toast.addEventListener('transitionend', () => toast.remove(), { once: true });
  }, duration);
}

// ═══════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════
function $id(id) { return document.getElementById(id); }

// Konvertiert Formulareingabe in den von der Graph API erwarteten Typ
function spValue(type, v) {
  if (type === 'number') return parseFloat(v);
  if (type === 'date')   return v ? new Date(v).toISOString() : null;
  // yesno: SP-Boolean-Spalten erwarten true/false, wir zeigen 'Ja'/'Nein'
  if (type === 'yesno')  return v === 'Ja' ? true : v === 'Nein' ? false : null;
  return v;
}

// Umgekehrt: SP-Wert → Anzeige-String für yesno-Felder
function spDisplayValue(type, v) {
  if (type === 'yesno') {
    if (v === true  || v === 1) return 'Ja';
    if (v === false || v === 0) return 'Nein';
  }
  return v ?? '';
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

// ═══════════════════════════════════════════════════════════════════
// EINSTELLUNGEN (nur Gremium)
// ═══════════════════════════════════════════════════════════════════

// Parst die Zustimmungs-IDs aus dem [APPROVALS:email1|email2] Token im GremiumKommentar
function parseApprovals(kommentar) {
  const m = (kommentar || '').match(/\[APPROVALS:(.*?)\]/);
  return m ? m[1].split('|').filter(Boolean) : [];
}

function loadSettings() {
  try { return JSON.parse(localStorage.getItem('ki_settings') || '{}'); } catch { return {}; }
}
function saveSettingsData(data) {
  localStorage.setItem('ki_settings', JSON.stringify(data));
}

function renderEinstellungen() {
  const settings = loadSettings();
  const genehmiger = settings.genehmiger || [];
  const ben = settings.benachrichtigung || {};

  const genList = genehmiger.length
    ? genehmiger.map((g, i) => {
        const rc = ROLLE_COLORS[g.rolle] || ROLLE_COLORS['Sonstiges'];
        const rolleBadge = g.rolle
          ? `<span style="font-size:.7rem;font-weight:600;padding:2px 8px;border-radius:20px;background:${rc.bg};color:${rc.color};border:1px solid ${rc.border}">${esc(g.rolle)}</span>`
          : '';
        const rolleOpts = ANLAGE_ROLLEN.map(r =>
          `<option value="${r}" ${g.rolle === r ? 'selected' : ''}>${r}</option>`).join('');
        return `
        <div style="display:flex;align-items:center;gap:10px;padding:9px 12px;background:#f9fafb;border-radius:8px;margin-bottom:6px">
          <div style="flex:1;min-width:0">
            <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
              <span style="font-size:.875rem;font-weight:600">${esc(g.name || g.email)}</span>
              ${rolleBadge}
            </div>
            <div style="font-size:.75rem;color:#9ca3af"><span style="font-size:.68rem;font-weight:600;color:#d1d5db;text-transform:uppercase;letter-spacing:.3px">UPN </span>${esc(g.email)}</div>
          </div>
          <select class="filter-select" style="font-size:.75rem;padding:3px 6px;height:auto" onchange="updateGenehmigerRolle(${i}, this.value)">
            <option value="">Rolle…</option>
            ${rolleOpts}
          </select>
          <button class="btn btn-neutral btn-sm" onclick="removeGenehmiger(${i})">Entfernen</button>
        </div>`;
      }).join('')
    : `<div class="empty-state" style="padding:16px 10px;font-size:.82rem">Noch keine Genehmiger hinterlegt.</div>`;

  $id('einstellungen-body').innerHTML = `
    <div class="settings-grid">

      <div class="settings-card">
        <div class="settings-card-title">👤 Genehmiger verwalten</div>
        <p style="font-size:.82rem;color:#6b7280;margin-bottom:14px;line-height:1.5">
          Personen, die KI-Anträge prüfen und entscheiden. Bei neuen Anträgen kann ein vorausgefüllter E-Mail-Entwurf geöffnet werden.
        </p>
        <div id="gen-list">${genList}</div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:8px;margin-top:12px;align-items:end">
          <div>
            <label class="form-label">Anzeigename <span style="color:#9ca3af;font-weight:400">(optional)</span></label>
            <input id="new-gen-name" type="text" class="form-control" placeholder="Max Mustermann">
          </div>
          <div>
            <label class="form-label">UPN <span style="color:#6b7280;font-weight:400;font-size:.72rem">(Azure-Anmelde-E-Mail)</span></label>
            <input id="new-gen-email" type="email" class="form-control" placeholder="max@dihag.com">
          </div>
          <div>
            <label class="form-label">Rolle</label>
            <select id="new-gen-rolle" class="form-control">
              <option value="">– Keine –</option>
              ${ANLAGE_ROLLEN.map(r => `<option value="${r}">${r}</option>`).join('')}
            </select>
          </div>
          <button class="btn btn-primary btn-sm" onclick="addGenehmiger()" style="white-space:nowrap;margin-bottom:1px">+ Hinzufügen</button>
        </div>
        <div style="margin-top:8px;font-size:.75rem;color:#6b7280;line-height:1.4">
          ℹ️ Es werden ausschließlich UPN und Rolle gespeichert — kein E-Mail-Verlauf, keine Kommunikationsdaten.
        </div>
        <div style="margin-top:18px;padding-top:16px;border-top:1px solid #e5e9ef">
          <div style="font-size:.82rem;font-weight:600;color:#374151;margin-bottom:10px">⚖️ Genehmigungsmodus</div>
          <label class="settings-check">
            <input type="radio" name="genmodus" id="modus-einstimmig" value="einstimmig" ${(ben.genehmigungsmodus || 'einstimmig') === 'einstimmig' ? 'checked' : ''}>
            <span><strong>Einstimmig</strong> – alle Genehmiger müssen einzeln zustimmen</span>
          </label>
          <label class="settings-check" style="margin-top:6px">
            <input type="radio" name="genmodus" id="modus-einer" value="einer" ${ben.genehmigungsmodus === 'einer' ? 'checked' : ''}>
            <span><strong>Einer reicht</strong> – eine Zustimmung genügt zur Freigabe</span>
          </label>
        </div>
      </div>

      <div class="settings-card">
        <div class="settings-card-title">📧 E-Mail-Benachrichtigungen</div>
        <p style="font-size:.82rem;color:#6b7280;margin-bottom:14px;line-height:1.5">
          E-Mails werden <strong>vollautomatisch</strong> über Microsoft Graph (Ihr Konto) versendet –
          kein E-Mail-Programm nötig. Voraussetzung: App-Berechtigung <code style="background:#f3f4f6;padding:1px 5px;border-radius:4px">Mail.Send</code>
          in der Azure AD App-Registrierung.
        </p>
        <label class="settings-check">
          <input type="checkbox" id="notif-einreichung" ${ben.beiEinreichung !== false ? 'checked' : ''}>
          <span>Bei neuem Antrag → alle Genehmiger automatisch benachrichtigen</span>
        </label>
        <label class="settings-check">
          <input type="checkbox" id="notif-entscheidung" ${ben.beiEntscheidung !== false ? 'checked' : ''}>
          <span>Nach Gremium-Entscheidung → Antragsteller automatisch benachrichtigen</span>
        </label>
      </div>

    </div>
    <div style="margin-top:20px;display:flex;gap:10px;align-items:center">
      <button class="btn btn-primary" onclick="saveSettings()">💾 Einstellungen speichern</button>
      <span id="settings-saved" style="font-size:.82rem;color:#15803d;display:none">✓ Gespeichert</span>
    </div>`;
}

function saveSettings() {
  const prev  = loadSettings();
  const modus = document.querySelector('input[name="genmodus"]:checked')?.value || 'einstimmig';
  const data = {
    genehmiger: prev.genehmiger || [],
    benachrichtigung: {
      beiEinreichung:    $id('notif-einreichung')?.checked  ?? true,
      beiEntscheidung:   $id('notif-entscheidung')?.checked ?? true,
      genehmigungsmodus: modus,
    },
  };
  saveSettingsData(data);
  showToast('Einstellungen gespeichert.');
  const saved = $id('settings-saved');
  if (saved) { saved.style.display = ''; setTimeout(() => saved.style.display = 'none', 2500); }
}

function addGenehmiger() {
  const name  = $id('new-gen-name')?.value.trim();
  const email = $id('new-gen-email')?.value.trim();
  const rolle = $id('new-gen-rolle')?.value || '';
  if (!email) { showToast('Bitte E-Mail-Adresse eingeben.', 'error'); return; }
  const settings  = loadSettings();
  const list      = settings.genehmiger || [];
  if (list.some(g => g.email.toLowerCase() === email.toLowerCase())) {
    showToast('Diese E-Mail-Adresse ist bereits hinterlegt.', 'error'); return;
  }
  list.push({ name: name || email, email, rolle: rolle || '' });
  settings.genehmiger = list;
  saveSettingsData(settings);
  renderEinstellungen();
  showToast(`✓ ${name || email} als Genehmiger hinzugefügt.`);
}

function updateGenehmigerRolle(index, rolle) {
  const settings = loadSettings();
  const list = settings.genehmiger || [];
  if (list[index]) {
    list[index].rolle = rolle;
    settings.genehmiger = list;
    saveSettingsData(settings);
    renderEinstellungen();
    showToast(`✓ Rolle aktualisiert.`);
  }
}

function removeGenehmiger(index) {
  const settings = loadSettings();
  const list = settings.genehmiger || [];
  const removed = list.splice(index, 1)[0];
  settings.genehmiger = list;
  saveSettingsData(settings);
  renderEinstellungen();
  if (removed) showToast(`${removed.name || removed.email} entfernt.`, 'info');
}

// ═══════════════════════════════════════════════════════════════════
// ANHÄNGE — SharePoint REST (/_api/web/lists/.../items/{id}/AttachmentFiles)
// ═══════════════════════════════════════════════════════════════════

async function spAttachFetch(url, options = {}) {
  const token = await tryGetSpToken();
  if (!token) throw new Error('SP-Token nicht verfügbar');
  const headers = {
    'Accept': 'application/json;odata=verbose',
    'Authorization': `Bearer ${token}`,
    ...(options.headers || {}),
  };
  const res = await fetch(url, { ...options, headers });
  if (!res.ok) {
    const err = await res.text().catch(() => res.statusText);
    throw new Error(`SP ${res.status}: ${err.slice(0, 200)}`);
  }
  if (res.status === 200) {
    const ct = res.headers.get('content-type') || '';
    if (ct.includes('json')) return res.json();
    return res.arrayBuffer();
  }
  return null;
}

async function listAttachments(itemId) {
  const url = `https://${SP_HOST}${SP_SITE_PATH}/_api/web/lists/getbytitle('${LIST_ANTRAEGE}')/items(${itemId})/AttachmentFiles`;
  const data = await spAttachFetch(url);
  return (data?.d?.results || []);
}

async function uploadAttachment(itemId, file) {
  const safeName = file.name.replace(/[#%&{}\\<>*?/$!'":@+`|=]/g, '_');
  const url = `https://${SP_HOST}${SP_SITE_PATH}/_api/web/lists/getbytitle('${LIST_ANTRAEGE}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`;
  const buf = await file.arrayBuffer();
  await spAttachFetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/octet-stream' },
    body: buf,
  });
}

async function deleteAttachment(itemId, fileName) {
  const url = `https://${SP_HOST}${SP_SITE_PATH}/_api/web/lists/getbytitle('${LIST_ANTRAEGE}')/items(${itemId})/AttachmentFiles/getbyfilename('${encodeURIComponent(fileName)}')`;
  const token = await tryGetSpToken();
  if (!token) throw new Error('SP-Token nicht verfügbar');
  const res = await fetch(url, {
    method: 'DELETE',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=verbose',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'DELETE',
    },
  });
  if (!res.ok && res.status !== 404) throw new Error(`SP ${res.status}`);
}

async function renderAttachments(itemId) {
  const listEl = $id('att-list');
  if (!listEl) return;
  try {
    const files = await listAttachments(itemId);
    if (!files.length) {
      listEl.innerHTML = `<span style="color:#9ca3af;font-size:.8rem">Keine Anhänge vorhanden.</span>`;
      return;
    }
    listEl.innerHTML = files.map(att => {
      const fname = att.FileName || att.d?.FileName || 'Datei';
      const url = att.ServerRelativeUrl
        ? `https://${SP_HOST}${att.ServerRelativeUrl}`
        : (att.d?.ServerRelativeUrl ? `https://${SP_HOST}${att.d.ServerRelativeUrl}` : '#');
      const isGrem = isGremium;
      return `<div class="att-item">
        <a class="att-name" href="${esc(url)}" target="_blank" rel="noopener">📄 ${esc(fname)}</a>
        ${isGrem ? `<button class="att-del" onclick="attDelete(${itemId},'${esc(fname.replace(/'/g, "\\'"))}')">✕</button>` : ''}
      </div>`;
    }).join('');
  } catch(e) {
    listEl.innerHTML = `<span style="color:#ef4444;font-size:.8rem">Fehler beim Laden: ${esc(e.message)}</span>`;
    console.warn('Anhänge laden fehlgeschlagen:', e);
  }
}

async function attDelete(itemId, fileName) {
  if (!confirm(`Anhang „${fileName}" wirklich löschen?`)) return;
  try {
    await deleteAttachment(itemId, fileName);
    showToast(`✓ „${fileName}" gelöscht.`);
    await renderAttachments(itemId);
  } catch(e) {
    showToast(`Fehler beim Löschen: ${e.message}`, 'error');
  }
}

async function attUploadFiles(itemId, files) {
  if (!files || !files.length) return;
  const MAX_MB = 50;
  const dropEl = $id('att-drop');
  if (dropEl) dropEl.classList.add('att-uploading');
  let uploaded = 0, failed = 0;
  for (const file of files) {
    if (file.size > MAX_MB * 1024 * 1024) {
      showToast(`„${file.name}" ist zu groß (max. ${MAX_MB} MB).`, 'error'); failed++; continue;
    }
    try {
      await uploadAttachment(itemId, file);
      uploaded++;
    } catch(e) {
      showToast(`Fehler bei „${file.name}": ${e.message}`, 'error'); failed++;
    }
  }
  if (dropEl) dropEl.classList.remove('att-uploading');
  if (uploaded) showToast(`✓ ${uploaded} Datei${uploaded > 1 ? 'en' : ''} hochgeladen.`);
  await renderAttachments(itemId);
}

function attDragOver(e) {
  e.preventDefault();
  $id('att-drop')?.classList.add('drag-over');
}
function attDragLeave(e) {
  $id('att-drop')?.classList.remove('drag-over');
}
function attDrop(e, itemId) {
  e.preventDefault();
  $id('att-drop')?.classList.remove('drag-over');
  const files = [...(e.dataTransfer?.files || [])];
  if (files.length) attUploadFiles(itemId, files);
}
function attFileSelect(e, itemId) {
  const files = [...(e.target?.files || [])];
  if (files.length) attUploadFiles(itemId, files);
  if (e.target) e.target.value = '';
}
