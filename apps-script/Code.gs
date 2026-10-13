/**
 * MIYAN — backend de réservation.
 * Déployé comme Web App (exécution "Moi", accès "Tout le monde").
 * Base de données = ce classeur Google Sheet, feuilles "Reservations" et "Config".
 *
 * Contrat API consommé par reservation.html :
 *   GET  ?                                -> { ok, config }
 *   GET  ?action=slots&date=YYYY-MM-DD&party=N -> { ok, slots:[{time,service,available}], config }
 *   GET  ?action=cancel&id=..&token=..    -> page HTML de confirmation d'annulation
 *   GET  ?action=list&key=..&date=..      -> tableau de bord admin (admin.html)
 *   POST { date, time, party, name, phone, email, note, lang, website } -> { ok, id }
 */

const SHEET_RES = 'Reservations';
const SHEET_CONFIG = 'Config';
const TZ = 'Europe/Paris';
const RES_HEADERS = ['id', 'createdAt', 'date', 'time', 'service', 'party', 'name', 'phone', 'email', 'note', 'lang', 'status', 'cancelToken'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'slots') return json(getSlots(p.date, Number(p.party || 1)));
    if (p.action === 'cancel') return cancelPage(p.id, p.token);
    if (p.action === 'list') return json(adminList(p));
    return json({ ok: true, config: publicConfig() });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    return json(createReservation(data));
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────── Config ───────────────────────── */

function rawConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  const rows = sh.getDataRange().getValues().slice(1); // skip header row
  const cfg = {};
  rows.forEach(r => { if (r[0]) cfg[String(r[0]).trim()] = r[1]; });
  return cfg;
}

function getConfig() {
  const c = rawConfig_();
  return {
    maxParty: Number(c.maxParty || 8),
    maxDays: Number(c.maxDays || 30),
    phone: String(c.phone || ''),
    email: String(c.email || ''),
    timezone: String(c.timezone || TZ),
    slotMinutes: Number(c.slotMinutes || 15),
    capacityPerSlot: Number(c.capacityPerSlot || 20),
    services: JSON.parse(c.servicesJson || '[{"id":"midi","label":"Midi","start":"12:00","end":"14:00"},{"id":"soir","label":"Soir","start":"19:00","end":"22:00"}]'),
    closedWeekdays: JSON.parse(c.closedWeekdays || '[1]'),
    closedDates: String(c.closedDates || '').split(',').map(s => s.trim()).filter(Boolean),
    adminKey: String(c.adminKey || ''),
    notifyEmails: String(c.notifyEmails || c.email || '').split(',').map(s => s.trim()).filter(Boolean)
  };
}

// Version envoyée au client (reservation.html) : jamais la clé admin.
function publicConfig() {
  const c = getConfig();
  return { maxParty: c.maxParty, maxDays: c.maxDays, phone: c.phone, email: c.email };
}

/* ───────────────────────── Créneaux ───────────────────────── */

function dateInRange_(dateStr, cfg) {
  const d = new Date(dateStr + 'T00:00:00');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const max = new Date(today); max.setDate(max.getDate() + cfg.maxDays);
  return d >= today && d <= max;
}

// Date et heure actuelles dans le fuseau du restaurant (Config.timezone), en chaînes
// comparables directement à "date"/"time" : sert à exclure les créneaux déjà passés aujourd'hui.
function nowInTz_(cfg) {
  const now = new Date();
  return {
    date: Utilities.formatDate(now, cfg.timezone, 'yyyy-MM-dd'),
    time: Utilities.formatDate(now, cfg.timezone, 'HH:mm')
  };
}

function isClosed_(dateStr, cfg) {
  const d = new Date(dateStr + 'T00:00:00');
  if (cfg.closedDates.indexOf(dateStr) !== -1) return true;
  return cfg.closedWeekdays.indexOf(d.getDay()) !== -1;
}

function timesForService_(service, slotMinutes) {
  const [sh, sm] = service.start.split(':').map(Number);
  const [eh, em] = service.end.split(':').map(Number);
  const start = sh * 60 + sm, end = eh * 60 + em;
  const out = [];
  for (let m = start; m < end; m += slotMinutes) {
    out.push(`${pad_(Math.floor(m / 60))}:${pad_(m % 60)}`);
  }
  return out;
}

function pad_(n) { return String(n).padStart(2, '0'); }

// Jauge applicable à un service : sa propre "capacity" si elle est définie dans
// servicesJson, sinon la jauge globale capacityPerSlot.
function capacityFor_(service, cfg) {
  return service.capacity != null ? Number(service.capacity) : cfg.capacityPerSlot;
}

// Total des couverts déjà confirmés pour une date+heure+service donnés.
function bookedPartyBySlot_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
  const rows = sh.getDataRange().getValues().slice(1);
  const totals = {}; // key `${time}|${service}` -> total party
  rows.forEach(r => {
    const [, , date, time, service, party, , , , , , status] = r;
    if (date === dateStr && status === 'confirmed') {
      const key = time + '|' + service;
      totals[key] = (totals[key] || 0) + Number(party);
    }
  });
  return totals;
}

function getSlots(dateStr, party) {
  const cfg = getConfig();
  if (!dateStr || !dateInRange_(dateStr, cfg)) return { ok: true, slots: [], config: publicConfig() };
  if (isClosed_(dateStr, cfg)) return { ok: true, slots: [], config: publicConfig() };

  const booked = bookedPartyBySlot_(dateStr);
  const now = nowInTz_(cfg);
  const isToday = dateStr === now.date;
  const slots = [];
  cfg.services.forEach(svc => {
    timesForService_(svc, cfg.slotMinutes).forEach(time => {
      const used = booked[time + '|' + svc.id] || 0;
      slots.push({
        time,
        service: svc.id,
        available: !(isToday && time <= now.time) && party <= cfg.maxParty && (used + party) <= capacityFor_(svc, cfg)
      });
    });
  });
  return { ok: true, slots, config: publicConfig() };
}

/* ───────────────────────── Réservation ───────────────────────── */

function createReservation(data) {
  // Honeypot anti-spam : on simule un succès sans rien enregistrer.
  if (data.website) return { ok: true, id: 'OK' };

  const cfg = getConfig();
  const date = String(data.date || ''), time = String(data.time || '');
  const party = Number(data.party || 0);
  const name = String(data.name || '').trim();
  const phone = String(data.phone || '').trim();
  const email = String(data.email || '').trim();

  if (!dateInRange_(date, cfg) || isClosed_(date, cfg)) return { ok: false, error: 'closed' };
  if (!time || party < 1 || party > cfg.maxParty) return { ok: false, error: 'invalid' };
  const now = nowInTz_(cfg);
  if (date === now.date && time <= now.time) return { ok: false, error: 'full' };
  if (!name) return { ok: false, error: 'invalid' };
  if (phone.replace(/\D/g, '').length < 8) return { ok: false, error: 'invalid' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'invalid' };

  const service = cfg.services.find(s => timesForService_(s, cfg.slotMinutes).indexOf(time) !== -1);
  if (!service) return { ok: false, error: 'invalid' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const booked = bookedPartyBySlot_(date);
    const used = booked[time + '|' + service.id] || 0;
    if (used + party > capacityFor_(service, cfg)) return { ok: false, error: 'full' };

    const id = 'MY-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const cancelToken = Utilities.getUuid();
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
    sh.appendRow([
      id, new Date(), date, time, service.id, party, name, phone, email,
      String(data.note || '').trim(), String(data.lang || 'fr'), 'confirmed', cancelToken
    ]);

    const r = { id, date, time, service, party, name, phone, email, note: data.note, lang: data.lang, cancelToken };
    sendConfirmationEmail_(cfg, r);
    sendStaffNotification_(cfg, r);
    return { ok: true, id };
  } finally {
    lock.releaseLock();
  }
}

function sendConfirmationEmail_(cfg, r) {
  const cancelUrl = ScriptApp.getService().getUrl() + '?action=cancel&id=' + encodeURIComponent(r.id) + '&token=' + encodeURIComponent(r.cancelToken);
  const fr = r.lang !== 'en';
  const subject = fr ? 'Votre réservation chez MIYAN — ' + r.date : 'Your MIYAN booking — ' + r.date;
  const body = fr
    ? `Bonjour ${r.name},\n\nVotre table est réservée :\n${r.date} à ${r.time} — ${r.party} personne(s)\nRéférence : ${r.id}\n\nPour annuler : ${cancelUrl}\n\nÀ bientôt,\nMIYAN — 62 rue du Faubourg Montmartre, 75009 Paris`
    : `Hello ${r.name},\n\nYour table is booked:\n${r.date} at ${r.time} — ${r.party} guest(s)\nReference: ${r.id}\n\nTo cancel: ${cancelUrl}\n\nSee you soon,\nMIYAN — 62 rue du Faubourg Montmartre, 75009 Paris`;
  MailApp.sendEmail({ to: r.email, replyTo: cfg.email || undefined, subject, body });
}

// Email à l'équipe (Config.notifyEmails, ou Config.email si vide) à chaque nouvelle réservation,
// pour confirmer côté restaurant sans avoir besoin d'admin.html ouvert.
function sendStaffNotification_(cfg, r) {
  if (!cfg.notifyEmails.length) return;
  const serviceLabel = (r.service && r.service.label) || (r.service && r.service.id) || '';
  const subject = `Nouvelle réservation — ${r.date} ${r.time} (${r.party} pers.)`;
  const body = `${r.name} — ${r.phone} — ${r.email}\n${r.date} à ${r.time} (${serviceLabel}) — ${r.party} personne(s)\nRéférence : ${r.id}` +
    (r.note ? `\nNote : ${r.note}` : '') +
    `\n\nVoir le tableau de bord : ${ScriptApp.getService().getUrl()}`;
  MailApp.sendEmail({ to: cfg.notifyEmails.join(','), subject, body });
}

/* ───────────────────────── Annulation ───────────────────────── */

function cancelPage(id, token) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
  const range = sh.getDataRange();
  const rows = range.getValues();
  let message = 'Lien invalide ou déjà utilisé.';
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id && rows[i][12] === token) {
      if (rows[i][11] === 'confirmed') {
        sh.getRange(i + 1, 12).setValue('cancelled'); // colonne "status"
        message = 'Votre réservation a bien été annulée. Merci de nous avoir prévenus !';
      } else {
        message = 'Cette réservation était déjà annulée.';
      }
      break;
    }
  }
  return HtmlService.createHtmlOutput(
    `<meta charset="utf-8"><div style="font-family:system-ui;max-width:420px;margin:60px auto;text-align:center;padding:0 20px;">` +
    `<h1 style="font-size:20px;">MIYAN</h1><p>${message}</p>` +
    `<p><a href="https://miyan.fr/">Retour au site</a></p></div>`
  );
}

/* ───────────────────────── Admin (tableau de bord) ───────────────────────── */

function adminList(p) {
  const cfg = getConfig();
  if (!cfg.adminKey || p.key !== cfg.adminKey) return { ok: false, error: 'unauthorized' };

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
  const rows = sh.getDataRange().getValues().slice(1);
  const fromDate = p.date || isoToday_();

  const list = rows
    .filter(r => r[2] >= fromDate)
    .map(r => ({
      id: r[0], createdAt: r[1], date: r[2], time: r[3], service: r[4],
      party: r[5], name: r[6], phone: r[7], email: r[8], note: r[9],
      lang: r[10], status: r[11]
    }))
    .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));

  return { ok: true, reservations: list, config: { capacityPerSlot: cfg.capacityPerSlot, services: cfg.services } };
}

function isoToday_() {
  const d = new Date();
  return `${d.getFullYear()}-${pad_(d.getMonth() + 1)}-${pad_(d.getDate())}`;
}

/* ───────────────────────── Installation (à lancer une fois) ───────────────────────── */

// Dans l'éditeur Apps Script : sélectionner "setupSpreadsheet" puis Exécuter.
// Crée les feuilles "Reservations" et "Config" si elles n'existent pas, avec les en-têtes
// et des valeurs par défaut. Change ensuite "adminKey" et "email"/"phone" directement
// dans la feuille Config.
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActive();

  let res = ss.getSheetByName(SHEET_RES);
  if (!res) res = ss.insertSheet(SHEET_RES);
  if (res.getLastRow() === 0) res.appendRow(RES_HEADERS);
  // Colonnes date (C) et heure (D) en texte brut : sans ça, Sheets convertit
  // "2026-10-20" / "19:30" en vraies dates/heures et casse les comparaisons de chaînes.
  res.getRange('C:D').setNumberFormat('@');

  let cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) cfg = ss.insertSheet(SHEET_CONFIG);
  // Colonne B en texte brut : sinon Sheets essaie d'interpréter "+33 1 00 00 00 00"
  // (phone) comme une formule et affiche #ERROR!.
  cfg.getRange('B:B').setNumberFormat('@');
  if (cfg.getLastRow() === 0) {
    cfg.appendRow(['key', 'value']);
    [
      ['maxParty', 8],
      ['maxDays', 30],
      ['phone', '+33 1 00 00 00 00'],
      ['email', 'contact@miyan.fr'],
      ['notifyEmails', ''],
      ['timezone', 'Europe/Paris'],
      ['slotMinutes', 15],
      ['capacityPerSlot', 20],
      ['servicesJson', '[{"id":"midi","label":"Midi","start":"12:00","end":"14:00"},{"id":"soir","label":"Soir","start":"19:00","end":"22:00"}]'],
      ['closedWeekdays', '[1]'],
      ['closedDates', ''],
      ['adminKey', Utilities.getUuid()]
    ].forEach(row => cfg.appendRow(row));
  }

  try {
    SpreadsheetApp.getUi().alert('Feuilles "Reservations" et "Config" prêtes. Ouvrez l\'onglet Config pour ajuster adminKey, email, phone, capacité, etc.');
  } catch (e) {
    // getUi() échoue si le classeur n'est pas ouvert dans un onglet — sans conséquence, le setup est déjà fait.
  }
}
