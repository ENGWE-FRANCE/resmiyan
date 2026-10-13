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
 *   GET  ?action=config&key=..            -> config complète (admin.html, écran Paramètres)
 *   POST { date, time, party, firstName, lastName, phone, email, note, lang, website } -> { ok, id }
 *   POST { action:'updateConfig', key, config:{...} } -> { ok } (admin.html, écran Paramètres)
 *   POST { action:'setOverride', key, date, time, service, capacity } -> { ok } (forcer/lever un créneau)
 *   POST { action:'adminCancel', key, id, reason } -> { ok } (annulation côté restaurant, email au client)
 */

const SHEET_RES = 'Reservations';
const SHEET_CONFIG = 'Config';
const SHEET_OVERRIDES = 'Overrides';
const TZ = 'Europe/Paris';
const RES_HEADERS = ['id', 'createdAt', 'date', 'time', 'service', 'party', 'firstName', 'lastName', 'phone', 'email', 'note', 'lang', 'status', 'cancelToken'];

// Lit la ligne d'en-têtes de "Reservations" et renvoie {nomDeColonne: index}, pour ne jamais
// dépendre de positions figées dans le tableau (fragile si le schéma évolue).
function resCols_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  return idx;
}

// Une cellule date/heure peut être un texte ("2026-10-13") ou, si Sheets l'a réinterprétée
// malgré le format texte brut, un vrai objet Date : on normalise toujours en chaîne comparable.
function asDateStr_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '');
}
function asTimeStr_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v || '');
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'slots') return json(getSlots(p.date, Number(p.party || 1)));
    if (p.action === 'cancel') return cancelPage(p.id, p.token);
    if (p.action === 'list') return json(adminList(p));
    if (p.action === 'config') return json(adminConfig(p));
    return json({ ok: true, config: publicConfig() });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'updateConfig') return json(updateConfig_(data));
    if (data.action === 'setOverride') return json(setOverride_(data));
    if (data.action === 'adminCancel') return json(adminCancelReservation_(data));
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
    services: JSON.parse(c.servicesJson || '[{"id":"midi","label":"Midi","start":"11:30","end":"15:00"},{"id":"soir","label":"Soir","start":"18:00","end":"23:00"}]'),
    closedWeekdays: JSON.parse(c.closedWeekdays || '[0]'),
    closedDates: String(c.closedDates || '').split(',').map(s => s.trim()).filter(Boolean),
    // Horaires différents de "services" pour certains jours (ex. mardi ferme plus tôt le midi) :
    // { "<0-6>": { "<idService>": {start,end} } }. Un jour/service absent garde l'horaire par défaut.
    hoursOverrides: JSON.parse(c.hoursOverrides || '{}'),
    adminKey: String(c.adminKey || ''),
    notifyEmails: String(c.notifyEmails || c.email || '').split(',').map(s => s.trim()).filter(Boolean),
    // Durée pendant laquelle une table reste occupée après l'heure de réservation : une jauge
    // "par créneau" n'a de sens que si on tient compte des réservations des créneaux voisins
    // qui occupent encore une table (sinon un même covert pourrait être compté comme libre
    // à chaque créneau de 15 min, alors qu'un repas dure bien plus longtemps).
    turnMinutes: Number(c.turnMinutes || 90)
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

// Services applicables à une date donnée : les services de base (cfg.services), avec les
// horaires de ce jour de la semaine remplacés si un hoursOverrides existe pour ce jour+service.
function servicesForDate_(dateStr, cfg) {
  const weekday = new Date(dateStr + 'T00:00:00').getDay();
  const overrides = cfg.hoursOverrides[String(weekday)] || {};
  return cfg.services.map(s => {
    const o = overrides[s.id];
    return (o && (o.start || o.end)) ? Object.assign({}, s, { start: o.start || s.start, end: o.end || s.end }) : s;
  });
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

function toMinutes_(hhmm) {
  const [h, m] = String(hhmm).split(':').map(Number);
  return h * 60 + m;
}

// Réservations confirmées d'une date, prêtes pour le calcul de chevauchement (une seule
// lecture de la feuille, réutilisée pour tous les créneaux d'un même appel).
function confirmedReservations_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
  const col = resCols_(sh);
  return sh.getDataRange().getValues().slice(1)
    .filter(r => asDateStr_(r[col.date]) === dateStr && r[col.status] === 'confirmed')
    .map(r => ({ time: asTimeStr_(r[col.time]), service: r[col.service], party: Number(r[col.party]) }));
}

// Couverts qui occupent encore une table au moment où une réservation à "time" commencerait et
// durerait "turnMinutes" : toute réservation existante dont la période d'occupation chevauche
// celle-ci compte, pas seulement celles à l'heure exacte "time".
function occupiedAt_(reservations, time, serviceId, turnMinutes) {
  const reqStart = toMinutes_(time), reqEnd = reqStart + turnMinutes;
  let total = 0;
  reservations.forEach(r => {
    if (r.service !== serviceId) return;
    const bStart = toMinutes_(r.time), bEnd = bStart + turnMinutes;
    if (reqStart < bEnd && bStart < reqEnd) total += r.party;
  });
  return total;
}

// Forçages manuels (admin) pour une date : clé "time|service" -> jauge imposée pour ce
// créneau précis (0 = marqué complet, une valeur plus haute = places libérées manuellement).
function overridesForDate_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_OVERRIDES);
  if (!sh) return {};
  const rows = sh.getDataRange().getValues().slice(1);
  const map = {};
  rows.forEach(r => {
    if (asDateStr_(r[0]) === dateStr) map[asTimeStr_(r[1]) + '|' + r[2]] = Number(r[3]);
  });
  return map;
}

// Jauge effective d'un créneau : le forçage manuel s'il existe, sinon la jauge normale du service.
function effectiveCapacity_(time, service, overrides, cfg) {
  const key = time + '|' + service.id;
  return overrides[key] != null ? overrides[key] : capacityFor_(service, cfg);
}

// Pose ou retire un forçage manuel sur un créneau précis (admin.html, grille d'occupation).
// capacity vide/null -> retire le forçage (retour à la jauge automatique).
function setOverride_(data) {
  const cfg = getConfig();
  if (!cfg.adminKey || data.key !== cfg.adminKey) return { ok: false, error: 'unauthorized' };

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_OVERRIDES);
  const rows = sh.getDataRange().getValues();
  const date = String(data.date || ''), time = String(data.time || ''), service = String(data.service || '');
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (asDateStr_(rows[i][0]) === date && asTimeStr_(rows[i][1]) === time && rows[i][2] === service) { rowIndex = i; break; }
  }

  if (data.capacity === '' || data.capacity == null) {
    if (rowIndex !== -1) sh.deleteRow(rowIndex + 1);
    return { ok: true };
  }

  if (rowIndex !== -1) {
    sh.getRange(rowIndex + 1, 4).setValue(Number(data.capacity));
  } else {
    sh.appendRow(["'" + date, "'" + time, service, Number(data.capacity), new Date()]);
  }
  return { ok: true };
}

function getSlots(dateStr, party) {
  const cfg = getConfig();
  if (!dateStr || !dateInRange_(dateStr, cfg)) return { ok: true, slots: [], config: publicConfig() };
  if (isClosed_(dateStr, cfg)) return { ok: true, slots: [], config: publicConfig() };

  const reservations = confirmedReservations_(dateStr);
  const overrides = overridesForDate_(dateStr);
  const now = nowInTz_(cfg);
  const isToday = dateStr === now.date;
  const slots = [];
  servicesForDate_(dateStr, cfg).forEach(svc => {
    timesForService_(svc, cfg.slotMinutes).forEach(time => {
      const occupied = occupiedAt_(reservations, time, svc.id, cfg.turnMinutes);
      const cap = effectiveCapacity_(time, svc, overrides, cfg);
      slots.push({
        time,
        service: svc.id,
        available: !(isToday && time <= now.time) && party <= cfg.maxParty && (occupied + party) <= cap
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
  const firstName = String(data.firstName || '').trim();
  const lastName = String(data.lastName || '').trim();
  const phone = String(data.phone || '').trim();
  const email = String(data.email || '').trim();

  if (!dateInRange_(date, cfg) || isClosed_(date, cfg)) return { ok: false, error: 'closed' };
  if (!time || party < 1 || party > cfg.maxParty) return { ok: false, error: 'invalid' };
  const now = nowInTz_(cfg);
  if (date === now.date && time <= now.time) return { ok: false, error: 'full' };
  if (!firstName || !lastName) return { ok: false, error: 'invalid' };
  if (phone.replace(/\D/g, '').length < 8) return { ok: false, error: 'invalid' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'invalid' };

  const service = servicesForDate_(date, cfg).find(s => timesForService_(s, cfg.slotMinutes).indexOf(time) !== -1);
  if (!service) return { ok: false, error: 'invalid' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const occupied = occupiedAt_(confirmedReservations_(date), time, service.id, cfg.turnMinutes);
    const cap = effectiveCapacity_(time, service, overridesForDate_(date), cfg);
    if (occupied + party > cap) return { ok: false, error: 'full' };

    const id = 'MY-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const cancelToken = Utilities.getUuid();
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
    // L'apostrophe force Sheets à garder date/heure en texte brut, même si le format de
    // colonne venait à être perdu : sans ça "2026-10-13" peut être réinterprété en vraie date.
    sh.appendRow([
      id, new Date(), "'" + date, "'" + time, service.id, party, firstName, lastName, phone, email,
      String(data.note || '').trim(), String(data.lang || 'fr'), 'confirmed', cancelToken
    ]);

    const r = { id, date, time, service, party, firstName, lastName, phone, email, note: data.note, lang: data.lang, cancelToken };
    sendConfirmationEmail_(cfg, r);
    sendStaffNotification_(cfg, r);
    return { ok: true, id };
  } finally {
    lock.releaseLock();
  }
}

// Couleurs reprises de reservation.html / admin.html, pour que l'email ressemble au site.
const EMAIL_COLORS_ = { bg: '#2a0708', card: '#F4E5C4', paper: '#FBF3DF', ink: '#520D10', gold: '#C49A3A' };

function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Date longue dans la langue de la réservation, ex. "mardi 13 octobre" / "Tuesday 13 October".
function longDate_(dateStr, fr) {
  const d = new Date(dateStr + 'T00:00:00');
  return new Intl.DateTimeFormat(fr ? 'fr-FR' : 'en-GB', { weekday: 'long', day: 'numeric', month: 'long' }).format(d);
}

function emailShell_(bodyHtml) {
  const c = EMAIL_COLORS_;
  return `<div style="background:${c.bg};padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">` +
    `<div style="max-width:480px;margin:0 auto;background:${c.card};border-radius:16px;padding:28px 24px;border:1px solid ${c.gold};">` +
    `<div style="color:${c.ink};font-weight:900;letter-spacing:.08em;font-size:18px;margin:0 0 18px;">米安 MIYAN</div>` +
    bodyHtml +
    `</div></div>`;
}

function sendConfirmationEmail_(cfg, r) {
  const c = EMAIL_COLORS_;
  const fullName = `${r.firstName} ${r.lastName}`.trim();
  const cancelUrl = ScriptApp.getService().getUrl() + '?action=cancel&id=' + encodeURIComponent(r.id) + '&token=' + encodeURIComponent(r.cancelToken);
  const fr = r.lang !== 'en';
  const date = longDate_(r.date, fr);
  const guests = fr ? (r.party > 1 ? `${r.party} personnes` : '1 personne') : (r.party > 1 ? `${r.party} guests` : '1 guest');
  const subject = (fr ? 'Votre réservation chez MIYAN - ' : 'Your MIYAN booking - ') + r.date;

  const body = fr
    ? `Bonjour ${fullName},\n\nVotre table est réservée :\n${date} à ${r.time} - ${guests}\nRéférence : ${r.id}\n\nPour annuler : ${cancelUrl}\n\nÀ bientôt,\nMIYAN - 62 rue du Faubourg Montmartre, 75009 Paris`
    : `Hello ${fullName},\n\nYour table is booked:\n${date} at ${r.time} - ${guests}\nReference: ${r.id}\n\nTo cancel: ${cancelUrl}\n\nSee you soon,\nMIYAN - 62 rue du Faubourg Montmartre, 75009 Paris`;

  const htmlBody = emailShell_(
    `<h1 style="color:${c.ink};font-size:22px;margin:0 0 14px;text-transform:uppercase;">${fr ? "C'est réservé" : "You're booked"}</h1>` +
    `<p style="color:${c.ink};margin:0 0 16px;">${fr ? 'Bonjour' : 'Hello'} ${escapeHtml_(fullName)},</p>` +
    `<div style="background:${c.paper};border-left:4px solid ${c.gold};border-radius:8px;padding:12px 14px;color:${c.ink};font-weight:600;margin:0 0 18px;">` +
    `${date} ${fr ? 'à' : 'at'} ${r.time} - ${guests}<br>${fr ? 'Référence' : 'Reference'} : ${r.id}</div>` +
    `<p style="margin:0 0 22px;"><a href="${cancelUrl}" style="color:${c.ink};">${fr ? 'Annuler la réservation' : 'Cancel reservation'}</a></p>` +
    `<p style="color:${c.ink};margin:0;font-size:13px;">${fr ? 'À bientôt,' : 'See you soon,'}<br>MIYAN - 62 rue du Faubourg Montmartre, 75009 Paris</p>`
  );

  MailApp.sendEmail({ to: r.email, replyTo: cfg.email || undefined, subject, body, htmlBody });
}

// Email à l'équipe (Config.notifyEmails, ou Config.email si vide) à chaque nouvelle réservation,
// pour confirmer côté restaurant sans avoir besoin d'admin.html ouvert. Toujours en français.
function sendStaffNotification_(cfg, r) {
  if (!cfg.notifyEmails.length) return;
  const c = EMAIL_COLORS_;
  const fullName = `${r.firstName} ${r.lastName}`.trim();
  const serviceLabel = (r.service && r.service.label) || (r.service && r.service.id) || '';
  const date = longDate_(r.date, true);
  const subject = `Nouvelle réservation - ${r.date} ${r.time} (${r.party} pers.)`;

  const body = `${fullName} - ${r.phone} - ${r.email}\n${date} à ${r.time} (${serviceLabel}) - ${r.party} personne(s)\nRéférence : ${r.id}` +
    (r.note ? `\nNote : ${r.note}` : '') +
    `\n\nVoir le tableau de bord : ${ScriptApp.getService().getUrl()}`;

  const htmlBody = emailShell_(
    `<h1 style="color:${c.ink};font-size:20px;margin:0 0 14px;text-transform:uppercase;">Nouvelle réservation</h1>` +
    `<div style="background:${c.paper};border-left:4px solid ${c.gold};border-radius:8px;padding:12px 14px;color:${c.ink};font-weight:600;margin:0 0 16px;">` +
    `${date} à ${r.time} (${escapeHtml_(serviceLabel)}) - ${r.party} personne(s)<br>Référence : ${r.id}</div>` +
    `<p style="color:${c.ink};margin:0 0 6px;">${escapeHtml_(fullName)} - ${escapeHtml_(r.phone)} - ${escapeHtml_(r.email)}</p>` +
    (r.note ? `<p style="color:${c.ink};margin:0 0 16px;font-style:italic;">${escapeHtml_(r.note)}</p>` : '') +
    `<p style="margin:16px 0 0;"><a href="${ScriptApp.getService().getUrl()}" style="color:${c.ink};">Voir le tableau de bord</a></p>`
  );

  MailApp.sendEmail({ to: cfg.notifyEmails.join(','), subject, body, htmlBody });
}

/* ───────────────────────── Annulation ───────────────────────── */

// Annulation côté restaurant (admin.html) : contrairement au lien client, celle-ci envoie un
// email au client avec un motif éventuel.
function adminCancelReservation_(data) {
  const cfg = getConfig();
  if (!cfg.adminKey || data.key !== cfg.adminKey) return { ok: false, error: 'unauthorized' };

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
  const col = resCols_(sh);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][col.id] !== data.id) continue;
    if (rows[i][col.status] !== 'confirmed') return { ok: false, error: 'already_cancelled' };

    sh.getRange(i + 1, col.status + 1).setValue('cancelled');
    const r = {
      id: rows[i][col.id], date: asDateStr_(rows[i][col.date]), time: asTimeStr_(rows[i][col.time]),
      party: rows[i][col.party], firstName: rows[i][col.firstName], lastName: rows[i][col.lastName],
      email: rows[i][col.email], lang: rows[i][col.lang]
    };
    sendAdminCancellationEmail_(cfg, r, String(data.reason || '').trim());
    return { ok: true };
  }
  return { ok: false, error: 'not_found' };
}

function sendAdminCancellationEmail_(cfg, r, reason) {
  const c = EMAIL_COLORS_;
  const fullName = `${r.firstName} ${r.lastName}`.trim();
  const fr = r.lang !== 'en';
  const date = longDate_(r.date, fr);
  const guests = fr ? (r.party > 1 ? `${r.party} personnes` : '1 personne') : (r.party > 1 ? `${r.party} guests` : '1 guest');
  const subject = (fr ? 'Votre réservation chez MIYAN a été annulée - ' : 'Your MIYAN booking was cancelled - ') + r.date;

  const reasonLine = reason ? (fr ? `\nMotif : ${reason}` : `\nReason: ${reason}`) : '';
  const body = fr
    ? `Bonjour ${fullName},\n\nNous sommes désolés, votre réservation du ${date} à ${r.time} (${guests}) a dû être annulée par le restaurant.${reasonLine}\n\nPour toute question : ${cfg.email}\n\nMIYAN - 62 rue du Faubourg Montmartre, 75009 Paris`
    : `Hello ${fullName},\n\nWe're sorry, your booking on ${date} at ${r.time} (${guests}) had to be cancelled by the restaurant.${reasonLine}\n\nFor any question: ${cfg.email}\n\nMIYAN - 62 rue du Faubourg Montmartre, 75009 Paris`;

  const htmlBody = emailShell_(
    `<h1 style="color:${c.ink};font-size:20px;margin:0 0 14px;text-transform:uppercase;">${fr ? 'Réservation annulée' : 'Booking cancelled'}</h1>` +
    `<p style="color:${c.ink};margin:0 0 16px;">${fr ? 'Bonjour' : 'Hello'} ${escapeHtml_(fullName)},</p>` +
    `<p style="color:${c.ink};margin:0 0 16px;">${fr ? 'Nous sommes désolés, la réservation suivante a dû être annulée par le restaurant' : 'We are sorry, the following booking had to be cancelled by the restaurant'} :</p>` +
    `<div style="background:${c.paper};border-left:4px solid ${c.gold};border-radius:8px;padding:12px 14px;color:${c.ink};font-weight:600;margin:0 0 18px;">` +
    `${date} ${fr ? 'à' : 'at'} ${r.time} - ${guests}` +
    (reason ? `<br><span style="font-weight:400;font-style:italic;">${fr ? 'Motif' : 'Reason'} : ${escapeHtml_(reason)}</span>` : '') +
    `</div>` +
    `<p style="color:${c.ink};margin:0;font-size:13px;">${fr ? 'Pour toute question' : 'For any question'} : <a href="mailto:${cfg.email}" style="color:${c.ink};">${cfg.email}</a></p>`
  );

  MailApp.sendEmail({ to: r.email, replyTo: cfg.email || undefined, subject, body, htmlBody });
}

function cancelPage(id, token) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
  const col = resCols_(sh);
  const rows = sh.getDataRange().getValues();
  let message = 'Lien invalide ou déjà utilisé.';
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][col.id] === id && rows[i][col.cancelToken] === token) {
      if (rows[i][col.status] === 'confirmed') {
        sh.getRange(i + 1, col.status + 1).setValue('cancelled');
        message = 'Votre réservation a bien été annulée. Merci de nous avoir prévenus !';
      } else {
        message = 'Cette réservation était déjà annulée.';
      }
      break;
    }
  }
  const c = EMAIL_COLORS_;
  return HtmlService.createHtmlOutput(
    `<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">` +
    `<div style="background:${c.bg};min-height:100vh;margin:0;padding:48px 16px;font-family:Arial,Helvetica,sans-serif;box-sizing:border-box;">` +
    `<div style="max-width:420px;margin:0 auto;background:${c.card};border-radius:16px;padding:28px 24px;border:1px solid ${c.gold};text-align:center;">` +
    `<div style="color:${c.ink};font-weight:900;letter-spacing:.08em;font-size:18px;margin:0 0 18px;">米安 MIYAN</div>` +
    `<p style="color:${c.ink};font-size:16px;margin:0 0 20px;">${message}</p>` +
    `<a href="https://miyan.fr/" style="color:${c.ink};">Retour au site</a>` +
    `</div></div>`
  );
}

/* ───────────────────────── Admin (tableau de bord) ───────────────────────── */

function adminList(p) {
  const cfg = getConfig();
  if (!cfg.adminKey || p.key !== cfg.adminKey) return { ok: false, error: 'unauthorized' };

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RES);
  const col = resCols_(sh);
  const rows = sh.getDataRange().getValues().slice(1);
  const fromDate = p.date || isoToday_();

  const list = rows
    .map(r => ({
      id: r[col.id], createdAt: r[col.createdAt], date: asDateStr_(r[col.date]), time: asTimeStr_(r[col.time]),
      service: r[col.service], party: r[col.party], firstName: r[col.firstName], lastName: r[col.lastName],
      phone: r[col.phone], email: r[col.email], note: r[col.note], lang: r[col.lang], status: r[col.status]
    }))
    .filter(r => r.date === fromDate) // la grille d'occupation et les forçages sont par date précise
    .sort((a, b) => a.time.localeCompare(b.time));

  return {
    ok: true,
    reservations: list,
    config: { capacityPerSlot: cfg.capacityPerSlot, services: servicesForDate_(fromDate, cfg), slotMinutes: cfg.slotMinutes, turnMinutes: cfg.turnMinutes },
    overrides: overridesForDate_(fromDate)
  };
}

// Config complète éditable (sans adminKey) pour l'écran Paramètres de admin.html.
function adminConfig(p) {
  const cfg = getConfig();
  if (!cfg.adminKey || p.key !== cfg.adminKey) return { ok: false, error: 'unauthorized' };
  return {
    ok: true,
    config: {
      maxParty: cfg.maxParty, maxDays: cfg.maxDays, phone: cfg.phone, email: cfg.email,
      notifyEmails: cfg.notifyEmails.join(', '), slotMinutes: cfg.slotMinutes, turnMinutes: cfg.turnMinutes,
      capacityPerSlot: cfg.capacityPerSlot, services: cfg.services,
      closedWeekdays: cfg.closedWeekdays, closedDates: cfg.closedDates.join(', '),
      hoursOverrides: cfg.hoursOverrides
    }
  };
}

// Écrit les champs fournis dans l'onglet Config (clé/valeur), sans toucher aux clés absentes.
function updateConfig_(data) {
  const cfg = getConfig();
  if (!cfg.adminKey || data.key !== cfg.adminKey) return { ok: false, error: 'unauthorized' };

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  const rows = sh.getDataRange().getValues();
  const setKV = (key, value) => {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
    }
    sh.appendRow([key, value]);
  };

  const c = data.config || {};
  if (c.maxParty != null) setKV('maxParty', Number(c.maxParty));
  if (c.maxDays != null) setKV('maxDays', Number(c.maxDays));
  if (c.phone != null) setKV('phone', "'" + String(c.phone));
  if (c.email != null) setKV('email', String(c.email));
  if (c.notifyEmails != null) setKV('notifyEmails', String(c.notifyEmails));
  if (c.slotMinutes != null) setKV('slotMinutes', Number(c.slotMinutes));
  if (c.capacityPerSlot != null) setKV('capacityPerSlot', Number(c.capacityPerSlot));
  if (c.turnMinutes != null) setKV('turnMinutes', Number(c.turnMinutes));
  if (c.services != null) setKV('servicesJson', JSON.stringify(c.services));
  if (c.closedWeekdays != null) setKV('closedWeekdays', JSON.stringify(c.closedWeekdays));
  if (c.closedDates != null) setKV('closedDates', String(c.closedDates));
  if (c.hoursOverrides != null) setKV('hoursOverrides', JSON.stringify(c.hoursOverrides));

  // La feuille Config utilise aussi la colonne B en texte brut (cf. setupSpreadsheet) ;
  // on la réapplique ici au cas où une ligne ait été ajoutée après coup par appendRow.
  sh.getRange('B:B').setNumberFormat('@');
  return { ok: true };
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
      ['phone', "'07 70 93 79 36"],
      ['email', 'contact@miyan.fr'],
      ['notifyEmails', ''],
      ['timezone', 'Europe/Paris'],
      ['slotMinutes', 15],
      ['capacityPerSlot', 20],
      ['turnMinutes', 90],
      ['servicesJson', '[{"id":"midi","label":"Midi","start":"11:30","end":"15:00"},{"id":"soir","label":"Soir","start":"18:00","end":"23:00"}]'],
      ['closedWeekdays', '[0]'],
      ['closedDates', ''],
      ['hoursOverrides', '{"2":{"midi":{"start":"11:30","end":"14:30"},"soir":{"start":"18:30","end":"23:00"}},"5":{"midi":{"start":"11:30","end":"14:30"}}}'],
      ['adminKey', Utilities.getUuid()]
    ].forEach(row => cfg.appendRow(row));
  }

  let ov = ss.getSheetByName(SHEET_OVERRIDES);
  if (!ov) ov = ss.insertSheet(SHEET_OVERRIDES);
  if (ov.getLastRow() === 0) ov.appendRow(['date', 'time', 'service', 'capacity', 'createdAt']);
  ov.getRange('A:B').setNumberFormat('@');

  try {
    SpreadsheetApp.getUi().alert('Feuilles "Reservations", "Config" et "Overrides" prêtes. Ouvrez l\'onglet Config pour ajuster adminKey, email, phone, capacité, etc.');
  } catch (e) {
    // getUi() échoue si le classeur n'est pas ouvert dans un onglet — sans conséquence, le setup est déjà fait.
  }
}
