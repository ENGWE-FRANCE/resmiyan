/**
 * Réservations Miyan — backend de réservation du restaurant MIYAN.
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

// Logo Réservations Miyan (dérivé de Logo_Miyan.jpeg à la racine du dépôt), encodé en base64 pour
// ne dépendre d'aucune URL externe : utilisé comme favicon de la page d'annulation et comme image
// intégrée (cid) dans les emails, deux contextes où Apps Script n'a pas accès aux fichiers du repo.
const FAVICON_PNG_BASE64_ = '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';
const LOGO_PNG_BASE64_ = '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';

function logoBlob_() {
  return Utilities.newBlob(Utilities.base64Decode(LOGO_PNG_BASE64_), 'image/png', 'logo-miyan.png');
}
function faviconDataUri_() {
  return 'data:image/png;base64,' + FAVICON_PNG_BASE64_;
}
function logoDataUri_() {
  return 'data:image/png;base64,' + LOGO_PNG_BASE64_;
}

function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Date longue dans la langue de la réservation, ex. "mardi 13 octobre" / "Tuesday 13 October".
function longDate_(dateStr, fr) {
  const d = new Date(dateStr + 'T00:00:00');
  return new Intl.DateTimeFormat(fr ? 'fr-FR' : 'en-GB', { weekday: 'long', day: 'numeric', month: 'long' }).format(d);
}

// Image embarquée en pièce jointe "cid" (voir logoBlob_ / inlineImages sur MailApp.sendEmail) :
// fonctionne dans Gmail/Outlook/etc. contrairement à une data URI, souvent bloquée par les clients mail.
function emailShell_(bodyHtml) {
  const c = EMAIL_COLORS_;
  return `<div style="background:${c.bg};padding:32px 16px;font-family:Arial,Helvetica,sans-serif;">` +
    `<div style="max-width:480px;margin:0 auto;background:${c.card};border-radius:16px;padding:28px 24px;border:1px solid ${c.gold};">` +
    `<div style="text-align:center;margin:0 0 18px;"><img src="cid:logo" width="44" height="44" alt="Réservations Miyan" style="display:inline-block;border-radius:10px;vertical-align:middle;"></div>` +
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

  MailApp.sendEmail({ to: r.email, replyTo: cfg.email || undefined, subject, body, htmlBody, inlineImages: { logo: logoBlob_() } });
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

  MailApp.sendEmail({ to: cfg.notifyEmails.join(','), subject, body, htmlBody, inlineImages: { logo: logoBlob_() } });
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

  MailApp.sendEmail({ to: r.email, replyTo: cfg.email || undefined, subject, body, htmlBody, inlineImages: { logo: logoBlob_() } });
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
    `<title>Réservations Miyan</title>` +
    `<link rel="icon" type="image/png" href="${faviconDataUri_()}">` +
    `<div style="background:${c.bg};min-height:100vh;margin:0;padding:48px 16px;box-sizing:border-box;display:flex;align-items:center;justify-content:center;font-family:Arial,Helvetica,sans-serif;">` +
    `<div style="width:100%;max-width:420px;margin:0 auto;background:${c.card};border-radius:16px;padding:28px 24px;border:1px solid ${c.gold};text-align:center;">` +
    `<img src="${logoDataUri_()}" width="48" height="48" alt="Réservations Miyan" style="display:inline-block;border-radius:10px;margin:0 0 14px;">` +
    `<div style="color:${c.ink};font-weight:900;letter-spacing:.08em;font-size:18px;margin:0 0 18px;">米安 MIYAN</div>` +
    `<p style="color:${c.ink};font-size:16px;margin:0 0 20px;">${message}</p>` +
    `<a href="https://miyan.fr/" style="color:${c.ink};">Retour au site</a>` +
    `</div></div>`
  ).setTitle('Réservations Miyan');
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
