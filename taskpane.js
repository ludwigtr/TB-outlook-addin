/**
 * taskpane.js — Thronsberg Email to Supabase Add-in
 * Flow: Read email → editable preview → pick contact(s)/project/stage/status → Send to Database.
 * The (possibly edited) email is stored in `activities` (permanent log) — one row per contact.
 * If a project is selected and stage/status actually change, apply_status_change() also writes
 * a status_changes row linked to that email via activity_id (reason_header/footer stay empty;
 * the app timeline falls back to showing the linked email as the reason).
 *
 * Two modes, both sharing the same email/project/stage/status controls below:
 *  - Single candidate (default): auto-match by sender address, or search/pick one contact.
 *  - Bulk ("Apply to multiple candidates"): search and pick several contacts; the same email
 *    gets logged for each, and the same target stage/status is applied to whichever of them
 *    are already on the selected project (see saveBulk()).
 */

// ── Stage / Status options (mirror of src/lib/pipeline-options.ts in the frontend) ──────────

var STAGES = ["Search", "Contacted", "Shortlist", "CV Sent", "Interview", "Offer", "Placed"];

var STAGE_STATUSES = {
  "Search":    ["Search ID", "Direct Application", "Decline", "Info Call", "Info Mail", "Info F2F", "Info LinkedIn", "Absage"],
  "Contacted": ["Cand to Call", "Cand to Call (E)", "Cand to Call (L)", "Int Cons", "Absage"],
  "Shortlist": ["Cand Feedback", "CV 2B Sent", "CV SL Select", "CV T-Style", "CV Ready", "Absage"],
  "CV Sent":   ["CV Sent", "Invited", "Client: On Hold", "Absage"],
  "Interview": ["Interview", "Absage"],
  "Offer":     ["References", "Offer", "Absage"],
  "Placed":    ["Offer Accepted"],
};

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
  if (!res.ok) throw new Error(data.error_description || data.message || 'Auth error ' + res.status);
  _sbToken   = data.access_token;
  _sbRefresh = data.refresh_token;
  _sbUserId  = (data.user && data.user.id) || '';
  return data.user;
}

async function sbRefreshToken() {
  if (!_sbRefresh) throw new Error('No refresh token — please log in again.');
  var res = await fetch(_sbUrl + '/auth/v1/token?grant_type=refresh_token', {
    method: 'POST',
    headers: { 'apikey': _sbAnon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ refresh_token: _sbRefresh }),
  });
  var data = await res.json();
  if (!res.ok) { _sbToken = ''; _sbRefresh = ''; persistRefreshToken(); throw new Error('Session expired — please log in again.'); }
  _sbToken   = data.access_token;
  _sbRefresh = data.refresh_token;
  _sbUserId  = (data.user && data.user.id) || _sbUserId;
  persistRefreshToken();
}

// Persists the current refresh token in roamingSettings (survives app restarts).
// Supabase rotates the token on every refresh — we must store the new one each time.
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
  if (!res.ok) { var e = await res.json().catch(function(){ return {}; }); throw new Error(e.message || 'DB error ' + res.status); }
  return res.json();
}

async function sbInsert(table, payload) {
  var res = await fetch(_sbUrl + '/rest/v1/' + table, {
    method: 'POST',
    headers: sbHeaders({ 'Prefer': 'return=representation' }),
    body: JSON.stringify(payload),
  });
  if (res.status === 401) { await sbRefreshToken(); return sbInsert(table, payload); }
  if (!res.ok) { var e = await res.json().catch(function(){ return {}; }); throw new Error(e.message || 'Insert error ' + res.status + ': ' + (e.details || e.hint || '')); }
  var rows = await res.json();
  return Array.isArray(rows) ? rows[0] : rows;
}

// Calls a Postgres function (RPC). Returns parsed JSON when present, else null.
async function sbRpc(fn, args) {
  var res = await fetch(_sbUrl + '/rest/v1/rpc/' + fn, {
    method: 'POST',
    headers: sbHeaders(),
    body: JSON.stringify(args || {}),
  });
  if (res.status === 401) { await sbRefreshToken(); return sbRpc(fn, args); }
  if (!res.ok) { var e = await res.json().catch(function(){ return {}; }); throw new Error(e.message || 'RPC error ' + res.status + ': ' + (e.details || e.hint || '')); }
  var text = await res.text();
  return text ? JSON.parse(text) : null;
}

// ── DOM helpers ────────────────────────────────────────────────────────────

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

// ── Global state ─────────────────────────────────────────────────────────────

var emailData      = null;   // { from, to, cc, subject, date, body, received }
var matchedContact = null;
var isSaving       = false;
// Current stage/status of the matched contact on the selected project (null if not linked).
var currentLink    = null;   // { stage, status } or null when no contacts_projects row exists

// Bulk mode: apply the same email + target stage/status to several candidates at once
// (e.g. "client wants to meet these 5" — pick them, set the stage/status, done).
var bulkMode       = false;
var bulkContacts   = [];     // [{ contact_id, first_name, last_name, email }]

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
      loadProjects();
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
  // Email + password are never stored — fields stay empty and are only used to log in.
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
    setStatus('error', '✗ Please fill in all fields.', 'settingsStatus');
    return;
  }

  $('settingsSaveBtn').disabled = true;
  setStatus('loading', '⏳ Testing connection…', 'settingsStatus');

  sbInit(url, anon);
  sbLogin(email, password).then(function() {
    var s = Office.context.roamingSettings;
    s.set('supabase_url',      url);
    s.set('supabase_anon_key', anon);
    s.set('refresh_token',     _sbRefresh);
    s.set('login_email',       '');
    s.set('login_password',    '');
    s.saveAsync(function(r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        $('settingsPassword').value = '';
        setStatus('success', '✓ Saved and logged in. Close the task pane and reopen it.', 'settingsStatus');
      } else {
        setStatus('error', '✗ Save failed.', 'settingsStatus');
      }
      $('settingsSaveBtn').disabled = false;
    });
  }).catch(function(err) {
    setStatus('error', '✗ ' + err.message, 'settingsStatus');
    $('settingsSaveBtn').disabled = false;
  });
});

// ── Switch state ──────────────────────────────────────────────────────────────

function show(sectionId) {
  ['notReady', 'notConfigured', 'notLoggedIn', 'readyContent'].forEach(function(id) {
    $(id).style.display = id === sectionId ? '' : 'none';
  });
}

// ── Step 1: Read email ─────────────────────────────────────────────────────────

$('readBtn').addEventListener('click', function() {
  $('readBtn').disabled = true;
  $('readBtnText').textContent = 'Reading…';
  setStatus('loading', '⏳ Reading email…');
  loadEmailData().then(function() {
    $('readBtnText').textContent = 'Re-read email';
    $('readBtn').disabled = false;
  });
});

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

        var subject = item.subject || '';
        var date    = item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString();

        emailData = { from: from, to: toList, cc: ccList, subject: subject, date: date, body: body, received: received };

        $('previewArea').style.display = 'flex';
        renderMeta();
        $('emailSubjectInput').value = subject;
        $('emailBodyInput').value    = body;
        hideStatus();

        matchContact().then(function() {
          resolve();
        });
      });
    } catch(err) {
      setStatus('error', '✗ Could not read email: ' + err.message);
      $('readBtn').disabled = false;
      $('readBtnText').textContent = 'Read email';
      resolve();
    }
  });
}

function renderMeta() {
  if (!emailData) return;
  $('emailFrom').textContent = emailData.from.name || emailData.from.address || '(unknown sender)';
  var meta = $('emailMeta');
  meta.innerHTML = '';
  [emailData.received ? '📥 Inbound' : '📤 Outbound',
   new Date(emailData.date).toLocaleDateString('en-GB')].forEach(function(t) {
    var span = document.createElement('span');
    span.className = 'meta-tag';
    span.textContent = t;
    meta.appendChild(span);
  });
}

// Builds the composed header block for `activities.email_header`, using the (edited) subject.
function buildHeader() {
  var fromStr = emailData.from.name ? emailData.from.name + ' <' + emailData.from.address + '>' : emailData.from.address;
  var toStr   = emailData.to.map(function(r) { return r.name ? r.name + ' <' + r.address + '>' : r.address; }).join(', ');
  var ccStr   = emailData.cc.map(function(r) { return r.name ? r.name + ' <' + r.address + '>' : r.address; }).join(', ');
  var subject = $('emailSubjectInput').value;
  return 'From: ' + fromStr + '\nTo: ' + toStr + (ccStr ? '\nCc: ' + ccStr : '') + '\nSubject: ' + subject + '\nDate: ' + emailData.date;
}

// ── Contact matching ────────────────────────────────────────────────────────

function matchContact() {
  if (!emailData) return Promise.resolve();

  var addr = emailData.received ? emailData.from.address : ((emailData.to[0] && emailData.to[0].address) || '');
  if (!addr) { showContactSearch(); enableSaveBtn(); return Promise.resolve(); }

  setStatus('loading', '⏳ Looking up contact (' + addr + ')…');

  return sbSelect('/contacts?email=ilike.' + encodeURIComponent(addr) + '&select=contact_id,first_name,last_name,email&limit=5')
    .then(function(rows) {
      if (rows.length === 1) {
        showMatchResult(rows[0]);
        hideStatus();
      } else if (rows.length > 1) {
        setStatus('warning', '⚠ ' + rows.length + ' contacts found — please choose.');
        showContactSearch(rows);
      } else {
        setStatus('warning', '⚠ No contact found for “' + addr + '”.');
        showContactSearch();
      }
      enableSaveBtn();
    })
    .catch(function(err) {
      setStatus('warning', '⚠ Contact lookup failed: ' + err.message);
      showContactSearch();
      enableSaveBtn();
    });
}

function showMatchResult(contact) {
  matchedContact = contact;
  var fullName = [contact.first_name, contact.last_name].filter(Boolean).join(' ') || '(no name)';
  $('matchName').textContent  = fullName;
  $('matchEmail').textContent = contact.email || '';
  $('matchResult').classList.add('visible');
  $('contactSearch').style.display = 'none';
  refreshStageStatus();
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
  refreshStageStatus();
});

$('skipContactBtn').addEventListener('click', function() {
  matchedContact = null;
  $('contactSearch').style.display = 'none';
  $('matchResult').classList.remove('visible');
  setStatus('warning', '⚠ Will be saved without a contact.');
  refreshStageStatus();
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
    d.textContent = 'No matches found.';
    container.appendChild(d);
    return;
  }
  rows.forEach(function(c) {
    var fullName = [c.first_name, c.last_name].filter(Boolean).join(' ') || '(no name)';

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

// ── Bulk mode (multiple candidates, one email, one target stage/status) ────────

$('bulkModeToggle').addEventListener('click', function() {
  bulkMode = !bulkMode;
  $('singleContactMode').style.display = bulkMode ? 'none' : '';
  $('bulkContactMode').style.display   = bulkMode ? '' : 'none';
  $('bulkModeToggle').textContent      = bulkMode ? '← Back to single candidate' : 'Apply to multiple candidates';
  $('contactSectionLabel').textContent = bulkMode ? 'Candidates (multiple)' : 'Contact';

  // Switching modes always starts clean — no mixing a single match with a bulk list.
  matchedContact = null;
  $('matchResult').classList.remove('visible');
  $('contactSearch').style.display = '';
  bulkContacts = [];
  renderBulkSelected();
  $('bulkSearchInput').value = '';
  $('bulkSearchResults').classList.remove('visible');
  $('bulkSearchResults').innerHTML = '';

  refreshStageStatus();
});

var bulkSearchTimer = null;
$('bulkSearchInput').addEventListener('input', function() {
  clearTimeout(bulkSearchTimer);
  var q = $('bulkSearchInput').value.trim();
  if (q.length < 2) { $('bulkSearchResults').classList.remove('visible'); $('bulkSearchResults').innerHTML = ''; return; }
  bulkSearchTimer = setTimeout(function() { runBulkContactSearch(q); }, 280);
});

function runBulkContactSearch(q) {
  var enc = encodeURIComponent('*' + q + '*');
  sbSelect('/contacts?or=(first_name.ilike.' + enc + ',last_name.ilike.' + enc + ',email.ilike.' + enc + ')&select=contact_id,first_name,last_name,email&limit=20')
    .then(renderBulkSearchItems)
    .catch(function() {});
  $('bulkSearchResults').classList.add('visible');
}

function renderBulkSearchItems(rows) {
  var container = $('bulkSearchResults');
  container.innerHTML = '';

  var already = {};
  bulkContacts.forEach(function(c) { already[c.contact_id] = true; });
  var filtered = rows.filter(function(c) { return !already[c.contact_id]; });

  if (!filtered.length) {
    var d = document.createElement('div');
    d.className = 'search-no-result';
    d.textContent = rows.length ? 'Already selected.' : 'No matches found.';
    container.appendChild(d);
    return;
  }

  filtered.forEach(function(c) {
    var fullName = [c.first_name, c.last_name].filter(Boolean).join(' ') || '(no name)';

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
    div.addEventListener('click', function() {
      bulkContacts.push(c);
      renderBulkSelected();
      $('bulkSearchInput').value = '';
      $('bulkSearchResults').classList.remove('visible');
      $('bulkSearchResults').innerHTML = '';
      refreshStageStatus();
    });
    container.appendChild(div);
  });
}

function renderBulkSelected() {
  var container = $('bulkSelectedList');
  container.innerHTML = '';
  bulkContacts.forEach(function(c, idx) {
    var fullName = [c.first_name, c.last_name].filter(Boolean).join(' ') || c.email || '(no name)';

    var chip = document.createElement('span');
    chip.className = 'chip';
    chip.textContent = fullName;

    var rm = document.createElement('button');
    rm.className = 'chip-remove';
    rm.type = 'button';
    rm.textContent = '✕';
    rm.addEventListener('click', function() {
      bulkContacts.splice(idx, 1);
      renderBulkSelected();
      refreshStageStatus();
    });

    chip.appendChild(rm);
    container.appendChild(chip);
  });
  $('bulkSelectedHint').style.display = bulkContacts.length ? 'none' : '';
}

// ── Projects ─────────────────────────────────────────────────────────────────

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

$('projectSelect').addEventListener('change', refreshStageStatus);

// ── Stage / Status ─────────────────────────────────────────────────────────────

// Fills the stage dropdown once.
(function initStageOptions() {
  var sel = $('stageSelect');
  STAGES.forEach(function(s) {
    var opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    sel.appendChild(opt);
  });
})();

// Repopulates the status dropdown for a given stage, optionally preselecting a value.
function fillStatusOptions(stage, selected) {
  var sel = $('statusSelect');
  sel.innerHTML = '<option value="">— Status —</option>';
  var list = (stage && STAGE_STATUSES[stage]) ? STAGE_STATUSES[stage] : [];
  list.forEach(function(s) {
    var opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    if (s === selected) opt.selected = true;
    sel.appendChild(opt);
  });
  sel.disabled = !stage;
}

// When the user changes the stage manually, reset the status (mirrors the app).
$('stageSelect').addEventListener('change', function() {
  fillStatusOptions($('stageSelect').value, null);
});

// Loads the current stage/status of the matched contact on the selected project and prefills.
// In bulk mode there is no single "current" stage/status to prefill (each candidate may sit
// somewhere different) — the user just picks the target stage/status to apply to everyone.
function refreshStageStatus() {
  var projectId = $('projectSelect').value;
  var area = $('stageStatusArea');
  var hint = $('stageStatusHint');
  currentLink = null;

  var hasTarget = bulkMode ? bulkContacts.length > 0 : Boolean(matchedContact);

  if (!projectId || !hasTarget) {
    area.style.display = 'none';
    return;
  }

  area.style.display = '';
  $('stageSelect').disabled = false;
  $('stageSelect').value = '';
  fillStatusOptions('', null);

  if (bulkMode) {
    hint.textContent = 'Applies to every selected candidate who is already on this project. Candidates not yet on the project only get the email logged, stage/status is skipped for them.';
    hint.style.display = '';
    return;
  }

  hint.style.display = 'none';

  sbSelect('/contacts_projects?contact_id=eq.' + encodeURIComponent(matchedContact.contact_id) +
           '&project_id=eq.' + encodeURIComponent(projectId) +
           '&select=stage,status&limit=1')
    .then(function(rows) {
      if (rows && rows.length) {
        currentLink = { stage: rows[0].stage || '', status: rows[0].status || '' };
        $('stageSelect').value = currentLink.stage;
        fillStatusOptions(currentLink.stage, currentLink.status);
      } else {
        // Candidate is not on this project — stage/status can't be changed from here.
        currentLink = null;
        $('stageSelect').value = '';
        $('stageSelect').disabled = true;
        $('statusSelect').disabled = true;
        hint.textContent = 'This candidate is not on the selected project — the email will be saved, but stage/status can only be set once they are added to the project.';
        hint.style.display = '';
      }
    })
    .catch(function() {
      area.style.display = 'none';
    });
}

function enableSaveBtn() {
  $('saveBtn').disabled = false;
  $('saveBtnText').textContent = 'Send to Database';
}

// ── Step 3: Send ────────────────────────────────────────────────────────────

$('saveBtn').addEventListener('click', function() {
  if (isSaving) return;
  if (!emailData) {
    setStatus('error', '✗ No email loaded — click “Read email” first.');
    return;
  }

  isSaving = true;
  $('saveBtn').disabled = true;
  $('saveBtnText').textContent = 'Sending…';
  setStatus('loading', '⏳ Saving…');

  var projectId = $('projectSelect').value || null;
  var body      = $('emailBodyInput').value;   // subject is read inside buildHeader()

  if (bulkMode) {
    saveBulk(projectId, body);
    return;
  }

  var payload = {
    contact_id:   (matchedContact && matchedContact.contact_id) || null,
    project_id:   projectId,
    user_id:      sbGetUserId() || null,
    email_body:   body,
    received:     emailData.received,
    email_header: buildHeader(),
  };

  // Decide whether a status/stage change should be logged.
  var newStage  = $('stageSelect').value || null;
  var newStatus = $('statusSelect').value || null;
  var doStatusChange =
    projectId &&
    matchedContact &&
    currentLink &&                                   // candidate is linked to the project
    ((newStage || '') !== (currentLink.stage || '') ||
     (newStatus || '') !== (currentLink.status || ''));

  sbInsert('activities', payload)
    .then(function(result) {
      if (!doStatusChange) return { changed: false };
      // Link the status change to the saved email (activity_id). We do NOT copy subject/body
      // into reason_header/footer — the timeline falls back to the linked email's content,
      // and those fields stay empty unless someone edits the reason manually in the app.
      var activityId = (result && result.email_activity_id) || null;
      // Send both stage and status (current-or-edited) so the trigger logs only real diffs.
      return sbRpc('apply_status_change', {
        p_contact_id: matchedContact.contact_id,
        p_project_id: projectId,
        p_stage:      newStage,
        p_status:     newStatus,
        p_activity_id: activityId,
      }).then(function() { return { changed: true }; });
    })
    .then(function(outcome) {
      var contactLabel = matchedContact
        ? ([matchedContact.first_name, matchedContact.last_name].filter(Boolean).join(' ') || matchedContact.email)
        : 'no contact';
      var projectOpt = $('projectSelect').selectedOptions[0];
      var projectLabel = (projectId && projectOpt) ? projectOpt.text : 'no project';
      var changeLine = outcome.changed
        ? ('Stage/Status: ' + ($('stageSelect').value || '—') + ' · ' + ($('statusSelect').value || '—'))
        : 'No stage/status change';
      setStatus('success', '✓ Saved\nContact: ' + contactLabel + '\nProject: ' + projectLabel + '\n' + changeLine);
      $('saveBtnText').textContent = 'Send again';
      // Refresh the cached current link so a second send doesn't re-log the same change.
      if (outcome.changed) currentLink = { stage: $('stageSelect').value || '', status: $('statusSelect').value || '' };
    })
    .catch(function(err) {
      setStatus('error', '✗ Error: ' + err.message);
      $('saveBtnText').textContent = 'Try again';
    })
    .finally(function() {
      isSaving = false;
      $('saveBtn').disabled = false;
    });
});

// ── Step 3b: Send (bulk mode) ───────────────────────────────────────────────
// For every selected candidate: log the (possibly edited) email as its own `activities` row,
// then — only if a target stage AND status were picked, and that candidate is already linked
// to the project — call apply_status_change() with that row's activity_id. Same "reference,
// not copy" pattern as the single-candidate flow: reason_header/footer stay empty, the app
// timeline shows the linked email as the reason unless someone overrides it manually.
function saveBulk(projectId, body) {
  if (!bulkContacts.length) {
    setStatus('error', '✗ Select at least one candidate first.');
    isSaving = false;
    $('saveBtn').disabled = false;
    $('saveBtnText').textContent = 'Send to Database';
    return;
  }

  var header       = buildHeader();
  var newStage     = $('stageSelect').value || null;
  var newStatus    = $('statusSelect').value || null;
  var wantsChange  = Boolean(projectId && newStage && newStatus);

  var loggedCount  = 0;
  var changedCount = 0;
  var skipped      = []; // names not updated (not on project / error)

  var chain = Promise.resolve();
  bulkContacts.forEach(function(c) {
    var label = [c.first_name, c.last_name].filter(Boolean).join(' ') || c.email || c.contact_id.slice(0, 8);

    chain = chain.then(function() {
      var payload = {
        contact_id:   c.contact_id,
        project_id:   projectId,
        user_id:      sbGetUserId() || null,
        email_body:   body,
        received:     emailData.received,
        email_header: header,
      };

      return sbInsert('activities', payload).then(function(result) {
        loggedCount++;
        if (!wantsChange) return;

        return sbSelect('/contacts_projects?contact_id=eq.' + encodeURIComponent(c.contact_id) +
                 '&project_id=eq.' + encodeURIComponent(projectId) +
                 '&select=stage,status&limit=1')
          .then(function(rows) {
            if (!rows || !rows.length) {
              skipped.push(label + ' (not on project)');
              return;
            }
            var activityId = (result && result.email_activity_id) || null;
            return sbRpc('apply_status_change', {
              p_contact_id:  c.contact_id,
              p_project_id:  projectId,
              p_stage:       newStage,
              p_status:      newStatus,
              p_activity_id: activityId,
            }).then(function() { changedCount++; });
          });
      }).catch(function(err) {
        skipped.push(label + ' (error: ' + err.message + ')');
      });
    });
  });

  chain.then(function() {
    var msg = loggedCount
      ? ('✓ Email logged for ' + loggedCount + ' of ' + bulkContacts.length + ' candidate(s)')
      : '✗ Nothing was saved.';
    if (wantsChange) msg += '\nStage/Status applied: ' + changedCount + ' of ' + bulkContacts.length;
    if (skipped.length) msg += '\nSkipped: ' + skipped.join(', ');
    setStatus(loggedCount ? 'success' : 'error', msg);
    $('saveBtnText').textContent = 'Send again';
  }).catch(function(err) {
    setStatus('error', '✗ Error: ' + err.message);
    $('saveBtnText').textContent = 'Try again';
  }).finally(function() {
    isSaving = false;
    $('saveBtn').disabled = false;
  });
}
