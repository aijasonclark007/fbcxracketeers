/**
 * Registration + proof-of-payment handler (Google Sheet web app).
 * Columns: A Timestamp | B Full Name | C Email | D Phone | E Skill Level
 *          F Status | G Emailed | H Token | I Payment Proof | J Paid At
 *
 * Waitlist: when a level already has its slots filled with approved players, a new registration is saved
 * with Status "Waitlist" and the player gets a waitlist email. Move them to "Approved" when a slot opens
 * and the normal approval email goes out.
 *
 * Jersey orders go to their own tab ("TournamentJersey", made automatically on the first order):
 *          A Timestamp | B Full Name | C Back Print Name | D Size | E Email
 *          F Status | G Emailed | H Token | I Payment Proof | J Paid At
 */
const SHEET_NAME = 'Sheet1';
const HEADERS = ['Timestamp', 'Full Name', 'Email', 'Phone Number', 'Skill Level'];
const TOKEN_COL = 8; // H

// Tournament Jersey orders
const JERSEY_SHEET = 'TournamentJersey';
const JERSEY_HEADERS = ['Timestamp', 'Full Name', 'Back Print Name', 'Size', 'Email', 'Status', 'Emailed', 'Token', 'Payment Proof', 'Paid At'];
const JERSEY_STATUSES = ['Pending', 'Send payment', 'Remind', 'Paid', 'Cancelled'];
const JERSEY_SIZES = ['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL'];
const JERSEY_PRICE = 'PHP 400';
const JERSEY_CLOSES = new Date('2026-11-02T00:00:00+08:00');   // orders close at the end of November 1, 2026 (Philippine time)

// Slots per level. The registration page reads these numbers from this script, so this is the only place to change them.
const SLOT_CAP = { High: 16, Mid: 32, Low: 24 };
const LEVELS = ['High', 'Mid', 'Low'];

// Who gets a short email every time someone registers. Kept out of the code on purpose (this file is public):
// in Apps Script open Project Settings > Script properties > Add script property:  NOTIFY_EMAIL = your@email.com
// If it is not set, the email goes to the account that owns this script.
function notifyEmail_() {
  return PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') || Session.getEffectiveUser().getEmail();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // players wait in line so two submissions never write to the same row
    const data = parseBody_(e);
    if (data.action === 'payment') return json_(handlePayment_(data));
    if (data.action === 'jersey') return json_(handleJersey_(data));

    const sheet = getSheet_();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    // older versions of the page marked waitlist sign-ups by adding [WAITLIST] to the name: clean it up
    data.name = String(data.name || '').replace(/\s*\[waitlist\]\s*$/i, '').trim();
    const level = LEVELS.filter(l => l.toLowerCase() === String(data.skill || '').trim().toLowerCase())[0];
    const waitlisted = !!level && countApproved_()[level] >= SLOT_CAP[level];   // checked here, so it cannot be skipped or go stale

    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 2, 1, 4).setNumberFormat('@'); // plain text, keeps leading 0 / +
    sheet.getRange(row, 1, 1, 5).setValues([[new Date(), data.name, data.email || '', data.phone || '', data.skill || '']]);
    if (waitlisted) {
      sheet.getRange(row, STATUS_COL).setValue('Waitlist');
      try {
        sendWaitlistEmail_(data.email, String(data.name).split(' ')[0], data.skill);
        sheet.getRange(row, EMAILED_COL).setValue('Waitlisted ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
      } catch (err) { sheet.getRange(row, EMAILED_COL).setValue('Failed: ' + err.message); }
    }
    notifyNewRegistration_(data, waitlisted);
    return json_({ result: 'success', waitlisted: waitlisted });
  } catch (err) {
    return json_({ result: 'error', message: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// Emails the organizer. A failure here never blocks the player's registration.
function notifyNewRegistration_(data, waitlisted) {
  try {
    MailApp.sendEmail({
      to: notifyEmail_(),
      name: 'FBC x Racketeers',
      subject: (waitlisted ? 'WAITLIST: ' : 'New registration: ') + (data.name || '') + ' (' + (data.skill || '') + ')',
      body:
        (waitlisted ? 'A player joined the WAITLIST (' + (data.skill || '') + ' is full).\n\n' : 'A new player just registered.\n\n') +
        'Name: ' + (data.name || '') + '\n' +
        'Email: ' + (data.email || '') + '\n' +
        'Phone: ' + (data.phone || '') + '\n' +
        'Skill level: ' + (data.skill || '') + '\n\n' +
        'Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });
  } catch (err) { /* ignore */ }
}

// Public slot counter for the registration page: how many players are approved in each level.
// Returns numbers only (no names or contact details). Players with status "Remind" are approved too.
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cache = CacheService.getScriptCache();
  let body = cache.get('slots2');                // cached for a minute so busy pages stay fast
  if (!body) {
    const out = countApproved_();
    out.cap = SLOT_CAP;                          // the page shows "x/16 slots" and opens the waitlist from these
    body = JSON.stringify(out);
    cache.put('slots2', body, 60);
  }
  const cb = String(p.callback || '').replace(/[^\w.$]/g, '');   // JSONP, so any website can read it
  return cb
    ? ContentService.createTextOutput(cb + '(' + body + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

// Approved players per level (Status "Approved" or "Remind"; waitlisted, pending and declined do not count)
function countApproved_() {
  const out = { High: 0, Mid: 0, Low: 0 };
  try {
    const rows = getSheet_().getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const level = LEVELS.filter(l => l.toLowerCase() === String(rows[i][4]).trim().toLowerCase())[0];
      const status = String(rows[i][5]).trim().toLowerCase();
      if (level && (status === 'approved' || status === 'remind')) out[level]++;
    }
  } catch (err) { /* leave the counts at 0 */ }
  return out;
}

// ----- Tournament Jersey orders -----
function jerseySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(JERSEY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(JERSEY_SHEET);
    sheet.getRange(1, 1, 1, JERSEY_HEADERS.length).setValues([JERSEY_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, STATUS_COL, 999, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(JERSEY_STATUSES, true).setAllowInvalid(false).build());
  }
  return sheet;
}

// Run once from the Apps Script editor to make the TournamentJersey tab right away (otherwise it appears with the first order)
function setupJerseySheet() {
  jerseySheet_();
}

function handleJersey_(data) {
  if (new Date() >= JERSEY_CLOSES) throw new Error('Jersey orders are closed.');
  const name = String(data.name || '').replace(/\s+/g, ' ').trim();
  const email = String(data.email || '').trim();
  const size = JERSEY_SIZES.filter(s => s === String(data.size || '').trim().toUpperCase())[0];
  const back = String(data.back || '').toUpperCase().replace(/[^A-Z\u00D1 .'\-]/g, '').replace(/\s+/g, ' ').trim().slice(0, 12);   // always capitals
  if (name.length < 2) throw new Error('Enter your full name.');
  if (back.length < 2) throw new Error('Enter the name to print on the back.');
  if (!size) throw new Error('Choose your size.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email.');

  const sheet = jerseySheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {          // one order per person: the same name with the same email cannot order twice
    if (String(rows[i][4]).trim().toLowerCase() === email.toLowerCase() && String(rows[i][1]).trim().toLowerCase() === name.toLowerCase() &&
        String(rows[i][5]).trim().toLowerCase() !== 'cancelled') throw new Error('This name already has a jersey order for that email.');
  }
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 2, 1, 4).setNumberFormat('@');
  sheet.getRange(row, 1, 1, 5).setValues([[new Date(), name, back, size, email]]);
  sheet.getRange(row, STATUS_COL).setValue('Pending');
  try {
    MailApp.sendEmail({
      to: notifyEmail_(),
      name: 'FBC x Racketeers',
      subject: 'New jersey order: ' + name + ' (' + size + ')',
      body: 'A new jersey order came in.\n\nName: ' + name + '\nBack print: ' + back + '\nSize: ' + size + '\nEmail: ' + email +
        '\n\nSet Status to "Send payment" in the TournamentJersey tab to email the payment link.\n\nSheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });
  } catch (err) { /* never blocks the order */ }
  return { result: 'success' };
}

// Saves the uploaded screenshot to Drive and links it on the player's row
function handlePayment_(data) {
  const sheet = data.kind === 'jersey' ? jerseySheet_() : getSheet_();      // jersey payment links carry kind=jersey
  const token = String(data.token || '');
  if (!token) throw new Error('Missing payment code');
  const rows = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][TOKEN_COL - 1]) === token) { row = i + 1; break; }
  }
  if (row < 0) throw new Error('Invalid payment code');

  const bytes = Utilities.base64Decode(data.data || '');
  if (!bytes.length || bytes.length > 8 * 1024 * 1024) throw new Error('Image is empty or too large');

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
  const blob = Utilities.newBlob(bytes, 'image/jpeg', rows[row - 1][1] + ' - payment - ' + stamp + '.jpg');
  const file = getProofFolder_().createFile(blob);

  sheet.getRange(row, 9).setValue(file.getUrl());
  sheet.getRange(row, 10).setValue(new Date());
  return { result: 'success' };
}

function getProofFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PROOF_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const folder = DriveApp.createFolder('Badminton Payment Proofs');
  props.setProperty('PROOF_FOLDER_ID', folder.getId());
  return folder;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function parseBody_(e) {
  if (e.parameter && Object.keys(e.parameter).length) return e.parameter;
  if (e.postData && e.postData.contents) return JSON.parse(e.postData.contents);
  return {};
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run once after pasting, to grant Drive + Mail permissions
function authorize() {
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
}