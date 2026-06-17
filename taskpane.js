/**
 * taskpane.js — Thronsberg Email to Supabase Add-in
 * Eine einzige Datei — Supabase-Logik und UI-Logik sind direkt hier drin.
 */

// ── Supabase REST-Client (inline) ────────────────────────────────────────────

var _sbUrl    = '';
var _sbAnon   = '';
var _sbToken  = '';
var _sbRefresh = '';
var _sbUserId = '';

function sbInit(url, anon) {
  _sbUrl  = url.replace(/\/$/, '');
  _sbAnon = anon;
}

function sbIsConfigured() { return Boolean(_sbUrl && _sbAnon); }
function sbIsLoggedIn()    { return Boolean(_sbToken); }
function sbGetUserId()     { return _sbUserId; }

async function sbLogin(email, password) {
  var res = await fetch(_sbUrl + '/auth/v1/token?grant_type=password', {
    method: 'POST',
    headers: { 'apikey': _sbAnon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email: email, password: password }),
  });
  var data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.message || 'Auth-Fehler ' + res.status);
  _sbToken   = data.access_token;
  _sbRefresh = data.refresh_token;
  _sbUserId  = (data.user && data.user.id) || '';
  return data.user;
}

async function sbRefreshToken() {
  if (!_sbRefresh) throw new Error('Kein Refresh-Token — bitte neu einloggen.');
  var res = await fetch(_sbUrl + '/auth/v1/token?grant_type=refresh_token', {
    method: 'POST',
    headers: { 'apikey': _sbAnon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ refresh_token: _sbRefresh }),
  });
  var data = await res.json();
  if (!res.ok) { _sbToken = ''; _sbRefresh = ''; persistRefreshToken(); throw new Error('Session abgelaufen — bitte neu einloggen.'); }
  _sbToken   = data.access_token;
  _sbRefresh = data.refresh_token;
  _sbUserId  = (data.user && data.user.id) || _sbUserId;
  persistRefreshToken();
}

// Speichert das aktuelle Refresh-Token in den roamingSettings (überlebt App-Neustarts).
// Supabase rotiert das Token bei jeder Erneuerung — wir müssen das jeweils neue ablegen.
function persistRefreshToken() {
  var s = Office.context.roamingSettings;
  s.set('refresh_token', _sbRefresh || '');
  s.saveAsync(function() {});
}

function sbHeaders(extra) {
  var h = { 'apikey': _sbAnon, 'Authorization': 'Bearer ' + _sbToken, 'Content-Type': 'application/json' };
  if (extra) Object.assign(h, extra);
  return h;
}

async function sbSelect(path) {
  var res = await fetch(_sbUrl + '/rest/v1' + path, { headers: sbHeaders() });
  if (res.status === 401) { await sbRefreshToken(); return sbSelect(path); }
  if (!res.ok) { var e = await res.json().catch(function(){ return {}; }); throw new Error(e.message || 'DB-Fehler ' + res.status); }
  return res.json();
}

async function sbInsert(table, payload) {
  var res = await fetch(_sbUrl + '/rest/v1/' + table, {
    method: 'POST',
    headers: sbHeaders({ 'Prefer': 'return=representation' }),
    body: JSON.stringify(payload),
  });
  if (res.status === 401) { await sbRefreshToken(); return sbInsert(table, payload); }
  if (!res.ok) { var e = await res.json().catch(function(){ return {}; }); throw new Error(e.message || 'Insert-Fehler ' + res.status + ': ' + (e.details || e.hint || '')); }
  var rows = await res.json();
  return Array.isArray(rows) ? rows[0] : rows;
}

// ── DOM-Helfer ───────────────────────────────────────────────────────────────

function $(id) { return document.getElementById(id); }

function setStatus(type, msg, target) {
  var el = $(target || 'status');
  el.className = 'status ' + type + ' visible';
  el.textContent = msg;
}

function hideStatus(target) {
  var el = $(target || 'status');
  el.className = 'status';
  el.textContent = '';
}

// ── Globaler Zustand ─────────────────────────────────────────────────────────

var emailData      = null;
var matchedContact = null;
var isSaving       = false;

// ── Office.onReady ───────────────────────────────────────────────────────────

Office.onReady(function() {
  loadSettingsToForm();

  var cfg = getSavedConfig();
  if (!cfg.url || !cfg.anon) { show('notConfigured'); return; }

  sbInit(cfg.url, cfg.anon);

  if (cfg.refresh) {
    _sbRefresh = cfg.refresh;
    sbRefreshToken().then(function() {
      show('readyContent');
      setStatus('loading', '⏳ Email wird geladen…');
      $('saveBtn').disabled = true;
      Promise.all([loadEmailData(), loadProjects()]);
    }).catch(function() {
      show('notLoggedIn');
    });
  } else {
    show('notLoggedIn');
  }
});

// ── Settings ─────────────────────────────────────────────────────────────────

function getSavedConfig() {
  var s = Office.context.roamingSettings;
  return {
    url:     s.get('supabase_url')      || '',
    anon:    s.get('supabase_anon_key') || '',
    refresh: s.get('refresh_token')     || '',
  };
}

function loadSettingsToForm() {
  var cfg = getSavedConfig();
  $('settingsUrl').value  = cfg.url;
  $('settingsAnon').value = cfg.anon;
  // Email + Passwort werden nie gespeichert — Felder bleiben leer und dienen nur dem Login.
}

$('settingsToggle').addEventListener('click', function() {
  $('settingsPanel').classList.toggle('visible');
});

$('settingsSaveBtn').addEventListener('click', function() {
  var url      = $('settingsUrl').value.trim();
  var anon     = $('settingsAnon').value.trim();
  var email    = $('settingsEmail').value.trim();
  var password = $('settingsPassword').value;

  if (!url || !anon || !email || !password) {
    setStatus('error', '✗ Alle Felder ausfüllen.', 'settingsStatus');
    return;
  }

  $('settingsSaveBtn').disabled = true;
  setStatus('loading', '⏳ Verbindung wird getestet…', 'settingsStatus');

  sbInit(url, anon);
  sbLogin(email, password).then(function() {
    var s = Office.context.roamingSettings;
    s.set('supabase_url',      url);
    s.set('supabase_anon_key', anon);
    s.set('refresh_token',     _sbRefresh);
    // Passwort wird NICHT gespeichert. Alte Klartext-Felder aus früheren Versionen aktiv leeren.
    s.set('login_email',       '');
    s.set('login_password',    '');
    s.saveAsync(function(r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        $('settingsPassword').value = '';
        setStatus('success', '✓ Gespeichert und eingeloggt. Task Pane schließen und neu öffnen.', 'settingsStatus');
      } else {
        setStatus('error', '✗ Speichern fehlgeschlagen.', 'settingsStatus');
      }
      $('settingsSaveBtn').disabled = false;
    });
  }).catch(function(err) {
    setStatus('error', '✗ ' + err.message, 'settingsStatus');
    $('settingsSaveBtn').disabled = false;
  });
});

// ── Zustand wechseln ─────────────────────────────────────────────────────────

function show(sectionId) {
  ['notReady', 'notConfigured', 'notLoggedIn', 'readyContent'].forEach(function(id) {
    $(id).style.display = id === sectionId ? '' : 'none';
  });
}

// ── Email lesen ──────────────────────────────────────────────────────────────

function loadEmailData() {
  return new Promise(function(resolve) {
    try {
      var item      = Office.context.mailbox.item;
      var myAddress = Office.context.mailbox.userProfile.emailAddress.toLowerCase();

      var from = item.from
        ? { address: (item.from.emailAddress || ''), name: (item.from.displayName || '') }
        : { address: '', name: '' };

      var toList = (item.to || []).map(function(r) { return { address: r.emailAddress || '', name: r.displayName || '' }; });
      var ccList = (item.cc || []).map(function(r) { return { address: r.emailAddress || '', name: r.displayName || '' }; });

      var received = from.address.toLowerCase() !== myAddress;

      item.body.getAsync(Office.CoercionType.Text, function(result) {
        var body = '';
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          body = result.value || '';
        }

        var subject  = item.subject || '';
        var date     = item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString();
        var toStr    = toList.map(function(r) { return r.name ? r.name + ' <' + r.address + '>' : r.address; }).join(', ');
        var ccStr    = ccList.map(function(r) { return r.name ? r.name + ' <' + r.address + '>' : r.address; }).join(', ');
        var fromStr  = from.name ? from.name + ' <' + from.address + '>' : from.address;
        var header   = 'From: ' + fromStr + '\nTo: ' + toStr + (ccStr ? '\nCc: ' + ccStr : '') + '\nSubject: ' + subject + '\nDate: ' + date;

        emailData = { from: from, to: toList, cc: ccList, subject: subject, date: date, body: body, received: received, header: header };

        renderEmailCard();
        matchContact().then(function() {
          enableSaveBtn();
          resolve();
        });
      });
    } catch(err) {
      setStatus('error', '✗ Email konnte nicht geladen werden: ' + err.message);
      resolve();
    }
  });
}

function renderEmailCard() {
  if (!emailData) return;
  $('emailCard').classList.add('visible');
  $('emailFrom').textContent    = emailData.from.name || emailData.from.address;
  $('emailSubject').textContent = emailData.subject || '(kein Betreff)';

  var meta = $('emailMeta');
  meta.innerHTML = '';
  [emailData.received ? '📥 Eingehend' : '📤 Ausgehend',
   new Date(emailData.date).toLocaleDateString('de-DE')].forEach(function(t) {
    var span = document.createElement('span');
    span.className = 'meta-tag';
    span.textContent = t;
    meta.appendChild(span);
  });
}

// ── Kontakt-Matching ─────────────────────────────────────────────────────────

function matchContact() {
  if (!emailData) return Promise.resolve();

  var addr = emailData.received ? emailData.from.address : ((emailData.to[0] && emailData.to[0].address) || '');
  if (!addr) { showContactSearch(); return Promise.resolve(); }

  setStatus('loading', '⏳ Kontakt wird gesucht (' + addr + ')…');

  return sbSelect('/contacts?email=ilike.' + encodeURIComponent(addr) + '&select=contact_id,first_name,last_name,email&limit=5')
    .then(function(rows) {
      if (rows.length === 1) {
        matchedContact = rows[0];
        showMatchResult(rows[0]);
        hideStatus();
      } else if (rows.length > 1) {
        setStatus('warning', '⚠ ' + rows.length + ' Kontakte gefunden — bitte auswählen.');
        showContactSearch(rows);
      } else {
        setStatus('warning', '⚠ Kein Kontakt für „' + addr + '" gefunden.');
        showContactSearch();
      }
    })
    .catch(function(err) {
      setStatus('warning', '⚠ Kontakt-Suche fehlgeschlagen: ' + err.message);
      showContactSearch();
    });
}

function showMatchResult(contact) {
  matchedContact = contact;
  var fullName = [contact.first_name, contact.last_name].filter(Boolean).join(' ') || '(kein Name)';
  $('matchName').textContent  = fullName;
  $('matchEmail').textContent = contact.email || '';
  $('matchResult').classList.add('visible');
  $('contactSearch').style.display = 'none';
}

function showContactSearch(preload) {
  $('matchResult').classList.remove('visible');
  $('contactSearch').style.display = '';
  if (preload && preload.length) { renderSearchItems(preload); $('searchResults').classList.add('visible'); }
}

$('matchClearBtn').addEventListener('click', function() {
  matchedContact = null;
  $('matchResult').classList.remove('visible');
  $('contactSearch').style.display = '';
  $('contactSearchInput').focus();
});

$('skipContactBtn').addEventListener('click', function() {
  matchedContact = null;
  $('contactSearch').style.display = 'none';
  $('matchResult').classList.remove('visible');
  setStatus('warning', '⚠ Wird ohne Kontakt gespeichert.');
});

var searchTimer = null;
$('contactSearchInput').addEventListener('input', function() {
  clearTimeout(searchTimer);
  var q = $('contactSearchInput').value.trim();
  if (q.length < 2) { $('searchResults').classList.remove('visible'); $('searchResults').innerHTML = ''; return; }
  searchTimer = setTimeout(function() { runContactSearch(q); }, 280);
});

function runContactSearch(q) {
  var enc = encodeURIComponent('*' + q + '*');
  sbSelect('/contacts?or=(first_name.ilike.' + enc + ',last_name.ilike.' + enc + ',email.ilike.' + enc + ')&select=contact_id,first_name,last_name,email&limit=20')
    .then(renderSearchItems)
    .catch(function() {});
  $('searchResults').classList.add('visible');
}

function renderSearchItems(rows) {
  var container = $('searchResults');
  container.innerHTML = '';
  if (!rows.length) {
    var d = document.createElement('div');
    d.className = 'search-no-result';
    d.textContent = 'Keine Treffer gefunden.';
    container.appendChild(d);
    return;
  }
  rows.forEach(function(c) {
    var fullName = [c.first_name, c.last_name].filter(Boolean).join(' ') || '(kein Name)';

    var div = document.createElement('div');
    div.className = 'search-item';

    var nameDiv = document.createElement('div');
    nameDiv.className = 'search-item-name';
    nameDiv.textContent = fullName;

    var emailDiv = document.createElement('div');
    emailDiv.className = 'search-item-email';
    emailDiv.textContent = c.email || '–';

    div.appendChild(nameDiv);
    div.appendChild(emailDiv);
    div.addEventListener('click', function() { showMatchResult(c); hideStatus(); });
    container.appendChild(div);
  });
}

// ── Projekte laden ───────────────────────────────────────────────────────────

function loadProjects() {
  return sbSelect('/projects?select=project_id,job_name&order=job_name')
    .then(function(rows) {
      var sel = $('projectSelect');
      rows.forEach(function(p) {
        var opt = document.createElement('option');
        opt.value = p.project_id;
        opt.textContent = p.job_name || p.project_id.slice(0, 8);
        sel.appendChild(opt);
      });
    })
    .catch(function() {});
}

function enableSaveBtn() {
  $('saveBtn').disabled = false;
  $('saveBtnText').textContent = 'In Datenbank speichern';
}

// ── Speichern ────────────────────────────────────────────────────────────────

$('saveBtn').addEventListener('click', function() {
  if (isSaving) return;
  isSaving = true;
  $('saveBtn').disabled = true;
  $('saveBtnText').textContent = 'Wird gespeichert…';
  setStatus('loading', '⏳ Datensatz wird angelegt…');

  if (!emailData) {
    setStatus('error', '✗ Email-Daten fehlen — Task Pane schließen und neu öffnen.');
    isSaving = false;
    $('saveBtn').disabled = false;
    $('saveBtnText').textContent = 'Erneut versuchen';
    return;
  }

  var payload = {
    contact_id:   (matchedContact && matchedContact.contact_id) || null,
    project_id:   $('projectSelect').value || null,
    user_id:      sbGetUserId() || null,
    email_body:   emailData.body,
    received:     emailData.received,
    email_header: emailData.header,
  };

  sbInsert('email_activities', payload).then(function(result) {
    var contactLabel = matchedContact
      ? ([matchedContact.first_name, matchedContact.last_name].filter(Boolean).join(' ') || matchedContact.email)
      : 'kein Kontakt';
    var projectOpt = $('projectSelect').selectedOptions[0];
    var projectLabel = (projectOpt && projectOpt.text !== '— Kein Projekt —') ? projectOpt.text : 'kein Projekt';
    setStatus('success', '✓ Gespeichert\nKontakt: ' + contactLabel + '\nProjekt: ' + projectLabel + '\nID: ' + ((result && result.email_activity_id && result.email_activity_id.slice(0, 8)) || '…'));
    $('saveBtnText').textContent = 'Erneut speichern';
  }).catch(function(err) {
    setStatus('error', '✗ Fehler: ' + err.message);
    $('saveBtnText').textContent = 'Erneut versuchen';
  }).finally(function() {
    isSaving = false;
    $('saveBtn').disabled = false;
  });
});
