/**
 * taskpane.js — Hauptlogik des Outlook Add-ins
 *
 * Ablauf:
 *  1. Office.onReady → Settings aus roamingSettings laden → Auto-Login
 *  2. Email-Daten lesen (from, to, subject, date, body)
 *  3. Kontakt-Matching per Absender-Email
 *  4. Projekte laden → Dropdown befüllen
 *  5. Speichern → INSERT in email_activities
 */

import * as db from './supabase.js';

// ── DOM-Helfer ──────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

function setStatus(type, msg, target = 'status') {
  const el = $(target);
  el.className = `status ${type} visible`;
  el.textContent = msg;
}

function hideStatus(target = 'status') {
  const el = $(target);
  el.className = 'status';
  el.textContent = '';
}

// ── Globaler Zustand ────────────────────────────────────────────────────────
let emailData     = null;   // { from, to, cc, subject, date, body, received }
let matchedContact = null;  // { contact_id, first_name, last_name, email } | null
let isSaving      = false;

// ── Office.onReady ──────────────────────────────────────────────────────────
Office.onReady(async () => {
  // Settings laden und in Felder schreiben
  loadSettingsToForm();

  const cfg = getSavedConfig();
  if (!cfg.url || !cfg.anon) {
    show('notConfigured');
    return;
  }

  db.init(cfg.url, cfg.anon);

  // Auto-Login wenn Credentials gespeichert
  if (cfg.email && cfg.password) {
    try {
      await db.login(cfg.email, cfg.password);
    } catch (_) {
      show('notLoggedIn');
      return;
    }
  } else {
    show('notLoggedIn');
    return;
  }

  // Eingeloggt: UI aufbauen
  show('readyContent');
  setStatus('loading', '⏳ Email wird geladen…');
  $('saveBtn').disabled = true;

  await Promise.all([loadEmailData(), loadProjects()]);
});

// ── Settings ─────────────────────────────────────────────────────────────────

function getSavedConfig() {
  const s = Office.context.roamingSettings;
  return {
    url:      s.get('supabase_url')      || '',
    anon:     s.get('supabase_anon_key') || '',
    email:    s.get('login_email')       || '',
    password: s.get('login_password')    || '',
  };
}

function loadSettingsToForm() {
  const cfg = getSavedConfig();
  $('settingsUrl').value      = cfg.url;
  $('settingsAnon').value     = cfg.anon;
  $('settingsEmail').value    = cfg.email;
  $('settingsPassword').value = cfg.password;
}

$('settingsToggle').addEventListener('click', () => {
  $('settingsPanel').classList.toggle('visible');
});

$('settingsSaveBtn').addEventListener('click', async () => {
  const url      = $('settingsUrl').value.trim();
  const anon     = $('settingsAnon').value.trim();
  const email    = $('settingsEmail').value.trim();
  const password = $('settingsPassword').value;

  if (!url || !anon || !email || !password) {
    setStatus('error', '✗ Alle Felder ausfüllen.', 'settingsStatus');
    return;
  }

  $('settingsSaveBtn').disabled = true;
  setStatus('loading', '⏳ Verbindung wird getestet…', 'settingsStatus');

  try {
    db.init(url, anon);
    await db.login(email, password);

    const s = Office.context.roamingSettings;
    s.set('supabase_url',      url);
    s.set('supabase_anon_key', anon);
    s.set('login_email',       email);
    s.set('login_password',    password);
    await new Promise((res, rej) => s.saveAsync(r =>
      r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(new Error(r.error?.message))));

    setStatus('success', '✓ Gespeichert und eingeloggt. Seite neu laden.', 'settingsStatus');
  } catch (err) {
    setStatus('error', `✗ ${err.message}`, 'settingsStatus');
  } finally {
    $('settingsSaveBtn').disabled = false;
  }
});

// ── Hilfsfunktion: sichtbaren Zustand wechseln ──────────────────────────────
function show(sectionId) {
  ['notReady', 'notConfigured', 'notLoggedIn', 'readyContent'].forEach(id => {
    $(id).style.display = id === sectionId ? '' : 'none';
  });
}

// ── Email-Daten lesen ────────────────────────────────────────────────────────

async function loadEmailData() {
  try {
    const item      = Office.context.mailbox.item;
    const myAddress = Office.context.mailbox.userProfile.emailAddress.toLowerCase();

    const from = item.from
      ? { address: item.from.emailAddress || '', name: item.from.displayName || '' }
      : { address: '', name: '' };

    const toList = (item.to || []).map(r => ({ address: r.emailAddress || '', name: r.displayName || '' }));
    const ccList = (item.cc || []).map(r => ({ address: r.emailAddress || '', name: r.displayName || '' }));

    // Richtung bestimmen: von uns gesendet oder eingegangen?
    const received = from.address.toLowerCase() !== myAddress;

    // Body als Plain Text (async)
    const body = await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || '');
        } else {
          reject(new Error(result.error?.message || 'Body konnte nicht geladen werden.'));
        }
      });
    });

    const subject = item.subject || '';
    const date    = item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString();

    // email_header als lesbaren Block
    const toStr = toList.map(r => r.name ? `${r.name} <${r.address}>` : r.address).join(', ');
    const ccStr = ccList.map(r => r.name ? `${r.name} <${r.address}>` : r.address).join(', ');
    const fromStr = from.name ? `${from.name} <${from.address}>` : from.address;

    let header = `From: ${fromStr}\nTo: ${toStr}`;
    if (ccStr) header += `\nCc: ${ccStr}`;
    header += `\nSubject: ${subject}\nDate: ${date}`;

    emailData = { from, to: toList, cc: ccList, subject, date, body, received, header };

    // UI aktualisieren
    renderEmailCard();
    await matchContact();
    enableSaveBtn();

  } catch (err) {
    setStatus('error', `✗ Email konnte nicht geladen werden: ${err.message}`);
  }
}

function renderEmailCard() {
  if (!emailData) return;
  $('emailCard').classList.add('visible');
  $('emailFrom').textContent    = emailData.from.name || emailData.from.address;
  $('emailSubject').textContent = emailData.subject || '(kein Betreff)';

  const meta = $('emailMeta');
  meta.innerHTML = '';
  const tags = [];
  tags.push(emailData.received ? '📥 Eingehend' : '📤 Ausgehend');
  tags.push(new Date(emailData.date).toLocaleDateString('de-DE'));
  tags.forEach(t => {
    const span = document.createElement('span');
    span.className = 'meta-tag';
    span.textContent = t;
    meta.appendChild(span);
  });
}

// ── Kontakt-Matching ─────────────────────────────────────────────────────────

async function matchContact() {
  if (!emailData) return;

  // Gegenpartei-Adresse: bei eingehend = Absender, bei ausgehend = erster Empfänger
  const counterpartAddress = emailData.received
    ? emailData.from.address
    : (emailData.to[0]?.address || '');

  if (!counterpartAddress) {
    showContactSearch();
    return;
  }

  setStatus('loading', `⏳ Kontakt wird gesucht (${counterpartAddress})…`);

  try {
    const encoded = encodeURIComponent(counterpartAddress);
    const rows = await db.select(
      `/contacts?email=ilike.${encoded}&select=contact_id,first_name,last_name,email&limit=5`
    );

    if (rows.length === 1) {
      // Eindeutiger Treffer
      matchedContact = rows[0];
      showMatchResult(rows[0]);
      hideStatus();
    } else if (rows.length > 1) {
      // Mehrere Treffer → manuelle Auswahl
      setStatus('warning', `⚠ ${rows.length} Kontakte mit dieser Adresse gefunden — bitte auswählen.`);
      showContactSearch(rows);
    } else {
      // Kein Treffer
      setStatus('warning', `⚠ Kein Kontakt für „${counterpartAddress}" gefunden.`);
      showContactSearch();
    }
  } catch (err) {
    setStatus('warning', `⚠ Kontakt-Suche fehlgeschlagen: ${err.message}`);
    showContactSearch();
  }
}

function showMatchResult(contact) {
  matchedContact = contact;
  const fullName = [contact.first_name, contact.last_name].filter(Boolean).join(' ') || '(kein Name)';
  $('matchName').textContent  = fullName;
  $('matchEmail').textContent = contact.email || '';
  $('matchResult').classList.add('visible');
  $('contactSearch').style.display = 'none';
}

function showContactSearch(preload = null) {
  $('matchResult').classList.remove('visible');
  $('contactSearch').style.display = '';

  if (preload?.length) {
    renderSearchItems(preload);
    $('searchResults').classList.add('visible');
  }
}

$('matchClearBtn').addEventListener('click', () => {
  matchedContact = null;
  $('matchResult').classList.remove('visible');
  $('contactSearch').style.display = '';
  $('contactSearchInput').focus();
});

$('skipContactBtn').addEventListener('click', () => {
  matchedContact = null;
  $('contactSearch').style.display = 'none';
  $('matchResult').classList.remove('visible');
  setStatus('warning', '⚠ Wird ohne Kontakt gespeichert.');
});

// Suche mit Debounce
let searchTimer = null;
$('contactSearchInput').addEventListener('input', () => {
  clearTimeout(searchTimer);
  const q = $('contactSearchInput').value.trim();
  if (q.length < 2) {
    $('searchResults').classList.remove('visible');
    $('searchResults').innerHTML = '';
    return;
  }
  searchTimer = setTimeout(() => runContactSearch(q), 280);
});

async function runContactSearch(q) {
  const enc = encodeURIComponent(`*${q}*`);
  try {
    const rows = await db.select(
      `/contacts?or=(first_name.ilike.${enc},last_name.ilike.${enc},email.ilike.${enc})&select=contact_id,first_name,last_name,email&limit=20`
    );
    renderSearchItems(rows);
    $('searchResults').classList.add('visible');
  } catch (_) {
    // Suche still scheitern lassen
  }
}

function renderSearchItems(rows) {
  const container = $('searchResults');
  container.innerHTML = '';
  if (!rows.length) {
    const div = document.createElement('div');
    div.className = 'search-no-result';
    div.textContent = 'Keine Treffer gefunden.';
    container.appendChild(div);
    return;
  }
  rows.forEach(c => {
    const fullName = [c.first_name, c.last_name].filter(Boolean).join(' ') || '(kein Name)';
    const div = document.createElement('div');
    div.className = 'search-item';
    div.innerHTML = `<div class="search-item-name">${fullName}</div><div class="search-item-email">${c.email || '–'}</div>`;
    div.addEventListener('click', () => {
      showMatchResult(c);
      hideStatus();
    });
    container.appendChild(div);
  });
}

// ── Projekte laden ───────────────────────────────────────────────────────────

async function loadProjects() {
  try {
    const rows = await db.select('/projects?select=project_id,job_name&order=job_name');
    const sel = $('projectSelect');
    rows.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.project_id;
      opt.textContent = p.job_name || p.project_id.slice(0, 8);
      sel.appendChild(opt);
    });
  } catch (_) {
    // Projekt-Dropdown optional — Fehler still ignorieren
  }
}

// ── Speichern-Button aktivieren ──────────────────────────────────────────────

function enableSaveBtn() {
  $('saveBtn').disabled = false;
  $('saveBtnText').textContent = 'In Datenbank speichern';
}

// ── INSERT email_activities ──────────────────────────────────────────────────

$('saveBtn').addEventListener('click', async () => {
  if (isSaving) return;
  isSaving = true;
  $('saveBtn').disabled = true;
  $('saveBtnText').textContent = 'Wird gespeichert…';
  setStatus('loading', '⏳ Datensatz wird angelegt…');

  try {
    if (!emailData) throw new Error('Email-Daten fehlen — bitte Task Pane schließen und neu öffnen.');

    // Feld-Mapping — zentrale Stelle für spätere Erweiterungen
    const payload = {
      contact_id:   matchedContact?.contact_id ?? null,
      project_id:   $('projectSelect').value || null,
      user_id:      db.getCurrentUserId() || null,
      email_body:   emailData.body,
      received:     emailData.received,
      email_header: emailData.header,
    };

    const result = await db.insert('email_activities', payload);

    const contactLabel = matchedContact
      ? `${[matchedContact.first_name, matchedContact.last_name].filter(Boolean).join(' ') || matchedContact.email}`
      : 'kein Kontakt';
    const projectLabel = $('projectSelect').selectedOptions[0]?.text !== '— Kein Projekt —'
      ? $('projectSelect').selectedOptions[0]?.text
      : 'kein Projekt';

    setStatus('success',
      `✓ Gespeichert\nKontakt: ${contactLabel}\nProjekt: ${projectLabel}\nID: ${result?.email_activity_id?.slice(0, 8) || '…'}`);

    $('saveBtnText').textContent = 'Erneut speichern';
  } catch (err) {
    setStatus('error', `✗ Fehler: ${err.message}`);
    $('saveBtnText').textContent = 'Erneut versuchen';
  } finally {
    isSaving = false;
    $('saveBtn').disabled = false;
  }
});
