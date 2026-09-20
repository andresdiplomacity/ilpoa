/**
 * ILPOA member access backend.
 * Bound to the "ILPOA Members" Google Sheet (Extensions > Apps Script).
 * Deployed as a Web App (Execute as: Me, Who has access: Anyone) and
 * called from the website via fetch(). See ../backend/SETUP.md.
 */

const SHEET_ID = '18qg93cGRIY2YKEs2G3YNo-OHZargCL_gjNesFLNUHi0';
const REQUESTS_SHEET = 'Requests';
const MEMBERS_SHEET = 'Members';
const HASH_ROUNDS = 1000; // simple stretching - Apps Script has no native slow-hash function
const RESET_TTL_MINUTES = 30;
const CONTACT_EMAIL = 'contact@islandlakeassociation.ca';

function getAdminEmail() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || Session.getEffectiveUser().getEmail();
}

function getSiteUrl() {
  return PropertiesService.getScriptProperties().getProperty('SITE_URL') || 'https://andresdiplomacity.github.io/ilpoa/';
}

function getSheet(name) {
  // getActiveSpreadsheet() only resolves inside an editor/UI context - a Web
  // App request has none, so it returns null there even for a bound script.
  // openById works the same in every context (editor, trigger, or Web App).
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

/** Run this once manually from the Apps Script editor (and again after schema changes) to (re)write headers and trigger the auth prompt. Safe to re-run - it never touches existing data rows. */
function ensureHeaders() {
  const req = getSheet(REQUESTS_SHEET);
  req.getRange(1, 1, 1, 10).setValues([['Timestamp', 'First Name', 'Last Name', 'Email', 'Phone', 'Address', 'Status', 'Processed At', 'Password Hash', 'Password Salt']]);
  req.setFrozenRows(1);

  const mem = getSheet(MEMBERS_SHEET);
  mem.getRange(1, 1, 1, 10).setValues([['First Name', 'Last Name', 'Email', 'Phone', 'Address', 'Approved At', 'Password Hash', 'Password Salt', 'Reset Token', 'Reset Token Expiry']]);
  mem.setFrozenRows(1);
}

// ===== Password hashing =====
// Salted, stretched SHA-256. Not as strong as bcrypt/argon2 (Apps Script has
// no native slow-hash function), but proportionate here - plaintext is never
// stored, and this is a low-stakes community site, not a bank.

function makeSalt() {
  return Utilities.getUuid();
}

function hashPassword(password, salt) {
  let bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password);
  for (let i = 0; i < HASH_ROUNDS; i++) {
    const asString = bytes.map((b) => String.fromCharCode((b + 256) % 256)).join('');
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + asString);
  }
  return bytes.map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

// ===== Web app entry point =====

function doPost(e) {
  ensureHeaders();
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'bad_request' });
  }
  try {
    switch (body.action) {
      case 'request_access': return handleRequestAccess(body);
      case 'login': return handleLogin(body);
      case 'request_password_reset': return handleRequestPasswordReset(body);
      case 'reset_password': return handleResetPassword(body);
      case 'contact': return handleContact(body);
      default: return jsonOut({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function findMemberRow(email) {
  const mem = getSheet(MEMBERS_SHEET);
  const data = mem.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === email) return i + 1; // 1-based sheet row
  }
  return -1;
}

// ===== Actions =====

function handleRequestAccess(body) {
  const firstName = (body.firstName || '').trim();
  const lastName = (body.lastName || '').trim();
  const email = (body.email || '').trim().toLowerCase();
  const phone = (body.phone || '').trim();
  const address = (body.address || '').trim();
  const password = body.password || '';

  if (!firstName || !lastName || !email || !phone || !address || !password) {
    return jsonOut({ ok: false, error: 'missing_fields' });
  }
  if (password.length < 6) {
    return jsonOut({ ok: false, error: 'weak_password' });
  }
  if (findMemberRow(email) > 0) {
    return jsonOut({ ok: false, error: 'already_member' });
  }

  const req = getSheet(REQUESTS_SHEET);
  const data = req.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() === email && data[i][6] === 'Pending') {
      return jsonOut({ ok: false, error: 'already_pending' });
    }
  }

  const salt = makeSalt();
  const hash = hashPassword(password, salt);
  req.appendRow([new Date(), firstName, lastName, email, phone, address, 'Pending', '', hash, salt]);

  MailApp.sendEmail({
    to: getAdminEmail(),
    subject: 'ILPOA — New member access request',
    body: `${firstName} ${lastName} (${email}, ${phone}) requested access.\nAddress: ${address}\n\nApprove or decline by editing the Status column (set it to "Approved" or "Declined") in the Requests sheet.`
  });
  MailApp.sendEmail({
    to: email,
    subject: 'Island Lake Association — Request received',
    body: `Hi ${firstName},\n\nWe received your request for member access to the Island Lake Association website. The site administrator will review it and email you once it's approved. This usually takes a few days.\n\n— Island Lake Association`
  });

  return jsonOut({ ok: true });
}

function handleLogin(body) {
  const email = (body.email || '').trim().toLowerCase();
  const password = body.password || '';
  const genericError = jsonOut({ ok: false, error: 'invalid_credentials' });
  if (!email || !password) return genericError;

  const rowIndex = findMemberRow(email);
  if (rowIndex < 0) return genericError;

  const mem = getSheet(MEMBERS_SHEET);
  const storedHash = mem.getRange(rowIndex, 7).getValue();
  const storedSalt = mem.getRange(rowIndex, 8).getValue();
  if (!storedHash || !storedSalt) return genericError;

  if (hashPassword(password, storedSalt) !== storedHash) return genericError;

  const firstName = mem.getRange(rowIndex, 1).getValue();
  const lastName = mem.getRange(rowIndex, 2).getValue();
  return jsonOut({ ok: true, member: { firstName, lastName, email } });
}

function handleRequestPasswordReset(body) {
  const email = (body.email || '').trim().toLowerCase();
  if (!email) return jsonOut({ ok: true }); // never reveal whether an email is registered

  const rowIndex = findMemberRow(email);
  if (rowIndex > 0) {
    const mem = getSheet(MEMBERS_SHEET);
    const token = Utilities.getUuid();
    const expiry = new Date(Date.now() + RESET_TTL_MINUTES * 60 * 1000);
    mem.getRange(rowIndex, 9).setValue(token);
    mem.getRange(rowIndex, 10).setValue(expiry);

    const firstName = mem.getRange(rowIndex, 1).getValue();
    const base = getSiteUrl();
    const resetUrl = base + (base.indexOf('?') >= 0 ? '&' : '?') + 'reset_token=' + encodeURIComponent(token);
    MailApp.sendEmail({
      to: email,
      subject: 'Reset your Island Lake Association password',
      body: `Hi ${firstName},\n\nClick this link to set a new password (valid for ${RESET_TTL_MINUTES} minutes):\n${resetUrl}\n\nIf you didn't request this, you can ignore this email - your password will stay the same.\n\n— Island Lake Association`
    });
  }
  return jsonOut({ ok: true });
}

function handleResetPassword(body) {
  const token = (body.token || '').trim();
  const newPassword = body.newPassword || '';
  if (!token || !newPassword) return jsonOut({ ok: false, error: 'missing_fields' });
  if (newPassword.length < 6) return jsonOut({ ok: false, error: 'weak_password' });

  const mem = getSheet(MEMBERS_SHEET);
  const data = mem.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === token) {
      const expiry = data[i][9];
      if (!expiry || new Date(expiry) <= now) return jsonOut({ ok: false, error: 'expired' });

      const salt = makeSalt();
      const hash = hashPassword(newPassword, salt);
      const row = i + 1;
      mem.getRange(row, 7).setValue(hash);
      mem.getRange(row, 8).setValue(salt);
      mem.getRange(row, 9).setValue('');
      mem.getRange(row, 10).setValue('');
      return jsonOut({ ok: true });
    }
  }
  return jsonOut({ ok: false, error: 'invalid' });
}

function handleContact(body) {
  const name = (body.name || '').trim();
  const email = (body.email || '').trim();
  const phone = (body.phone || '').trim();
  const message = (body.message || '').trim();

  if (!name || !email || !message) {
    return jsonOut({ ok: false, error: 'missing_fields' });
  }

  MailApp.sendEmail({
    to: CONTACT_EMAIL,
    replyTo: email,
    subject: `ILPOA website — message from ${name}`,
    body: `${name} (${email}${phone ? ', ' + phone : ''}) sent a message through the website contact form:\n\n${message}`
  });

  return jsonOut({ ok: true });
}

// ===== Installable trigger =====
// Set up manually: Apps Script editor > Triggers (clock icon) > Add trigger
//   Function: onStatusEdit | Event source: From spreadsheet | Event type: On edit
function onStatusEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== REQUESTS_SHEET) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (col !== 7 || row === 1) return; // Status column only, skip header row

  const newStatus = String(e.value || '').trim();
  if (newStatus !== 'Approved' && newStatus !== 'Declined') return;

  const processedCell = sheet.getRange(row, 8);
  if (processedCell.getValue()) return; // already processed - avoid double-sending

  const rowData = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const firstName = rowData[1], lastName = rowData[2], email = rowData[3], phone = rowData[4], address = rowData[5];
  const passwordHash = rowData[8], passwordSalt = rowData[9];

  if (newStatus === 'Approved') {
    if (findMemberRow(email) < 0) {
      getSheet(MEMBERS_SHEET).appendRow([firstName, lastName, email, phone, address, new Date(), passwordHash, passwordSalt]);
    }
    MailApp.sendEmail({
      to: email,
      subject: 'Your Island Lake Association account is approved',
      body: `Hi ${firstName},\n\nYour member access request has been approved. Visit the site and sign in with your email (${email}) and the password you chose when you requested access.\n\n${getSiteUrl()}\n\n— Island Lake Association`
    });
  } else {
    MailApp.sendEmail({
      to: email,
      subject: 'Island Lake Association — Access request update',
      body: `Hi ${firstName},\n\nWe're sorry, but your access request could not be approved at this time. If you believe this is a mistake, please reply to this email.\n\n— Island Lake Association`
    });
  }

  processedCell.setValue(new Date());
}
