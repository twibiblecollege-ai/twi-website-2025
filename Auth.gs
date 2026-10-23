// ============================================================
// TWI BIBLE COLLEGE — AUTHENTICATION & SESSION GUARD
// Auth.gs
//
// WHY THIS FILE EXISTS
// The web app is deployed with access = ANYONE_ANONYMOUS. Every global
// function in this project is therefore callable by any visitor via
// google.script.run from the browser console of any served page.
// checkLogin() and checkAdminLogin() are login *screens* — they return a
// result to the browser but nothing on the server ever verified that a
// caller had logged in before running the next function.
//
// This file provides:
//   1. Salted, peppered SHA-256 password hashing (_hashPassword_)
//   2. Backward-compatible verification that transparently upgrades old
//      plaintext passwords on first successful login (_verifyPassword_)
//   3. Server-side sessions (_issueSession_ / _requireStudent_ / etc.)
//
// Guards FAIL CLOSED: they throw, which reaches the client's
// withFailureHandler. A caller with no token gets nothing.
//
// Requires: Code.gs (_ss, _norm)
// ============================================================

// CacheService caps expiry at 21600s (6 hours). Do not raise past this.
const BUILD_AUTH = '2026-10-23f';   // version stamp — see getServerBuild() in Code.gs

const SESSION_TTL_SECONDS = 21600;
const SESSION_PREFIX      = 'twisess_';
const PEPPER_PROP_KEY     = 'TWI_PW_PEPPER';

// Column 29 (index 28) of Master is 'Password'. Kept here so the migration
// helpers below have one place to change if the sheet layout ever moves.
const MASTER_PW_COL = 29;

// ============================================================
// PASSWORD HASHING
// ============================================================

// The pepper is a per-deployment secret held in Script Properties, NOT in
// this file. Stealing the spreadsheet alone is not enough to crack hashes.
function _pepper_() {
  const props = PropertiesService.getScriptProperties();
  let p = props.getProperty(PEPPER_PROP_KEY);
  if (!p) {
    p = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty(PEPPER_PROP_KEY, p);
  }
  return p;
}

function _toHex_(bytes) {
  return bytes.map(function (b) { return ((b & 0xFF) + 0x100).toString(16).slice(1); }).join('');
}

// Stored format: sha256$<32-char salt>$<64-char hex digest>
function _hashPassword_(plain, saltArg) {
  const salt = saltArg || Utilities.getUuid().replace(/-/g, '');
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + String(plain) + _pepper_(),
    Utilities.Charset.UTF_8
  );
  return 'sha256$' + salt + '$' + _toHex_(digest);
}

function _isHashed_(stored) {
  return /^sha256\$[0-9a-f]{32}\$[0-9a-f]{64}$/.test(String(stored || '').trim());
}

// Constant-time-ish comparison. JS strings can't be compared in true constant
// time, but this avoids the early-exit of === on long shared prefixes.
function _safeEquals_(a, b) {
  const x = String(a), y = String(b);
  if (x.length !== y.length) return false;
  let diff = 0;
  for (let i = 0; i < x.length; i++) diff |= (x.charCodeAt(i) ^ y.charCodeAt(i));
  return diff === 0;
}

// Returns { ok, needsUpgrade }.
// needsUpgrade === true means the sheet still held a plaintext password and
// the caller should rewrite that cell with _hashPassword_(plain). This is what
// lets every existing student keep logging in with their current password
// while the plaintext drains out of the sheet by itself.
function _verifyPassword_(plain, stored) {
  const s = String(stored || '').trim();
  const p = String(plain || '').trim();
  if (!s) return { ok: false, needsUpgrade: false };

  if (!_isHashed_(s)) {
    return { ok: _safeEquals_(p, s), needsUpgrade: true };
  }
  const parts = s.split('$');
  return { ok: _safeEquals_(_hashPassword_(p, parts[1]), s), needsUpgrade: false };
}

// Rewrites one cell with the hash of the password that just succeeded.
// Called on the success path of a login that hit a plaintext row.
function _upgradeStoredPassword_(sheet, rowNum, colNum, plain) {
  try {
    sheet.getRange(rowNum, colNum).setValue(_hashPassword_(plain));
  } catch (e) {
    // Never let a migration failure block a valid login.
    console.error('Password upgrade failed on row ' + rowNum + ': ' + e.message);
  }
}

// ============================================================
// SESSIONS
// ============================================================

// type: 'student' | 'instructor' | 'admin'
function _issueSession_(type, subjectId, role) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const payload = JSON.stringify({
    t: String(type),
    id: _norm(subjectId),
    role: _norm(role || ''),
    iat: Date.now()
  });
  CacheService.getScriptCache().put(SESSION_PREFIX + token, payload, SESSION_TTL_SECONDS);
  return token;
}

function _readSession_(token) {
  const t = String(token || '').trim();
  if (!t) return null;
  const raw = CacheService.getScriptCache().get(SESSION_PREFIX + t);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function _revokeSession_(token) {
  const t = String(token || '').trim();
  if (t) CacheService.getScriptCache().remove(SESSION_PREFIX + t);
}

// Extends a live session so an active user is not logged out mid-task.
function _touchSession_(token) {
  const s = _readSession_(token);
  if (!s) return null;
  CacheService.getScriptCache().put(SESSION_PREFIX + String(token), JSON.stringify(s), SESSION_TTL_SECONDS);
  return s;
}

// ============================================================
// GUARDS — every one of these THROWS on failure. Fail closed.
// ============================================================

// True only when a human is running the function from the Apps Script editor.
// In a web app deployed as USER_DEPLOYING with anonymous access,
// getActiveUser() is empty for visitors, so this cannot be reached from the
// browser — it exists so maintenance functions (adminSetup,
// migratePlaintextPasswords) remain runnable by hand.
function _isEditorContext_() {
  try {
    const active = Session.getActiveUser().getEmail();
    return !!active && active === Session.getEffectiveUser().getEmail();
  } catch (e) {
    return false;
  }
}

function _requireSession_(token) {
  const s = _touchSession_(token);
  if (!s) throw new Error('Your session has expired. Please log in again.');
  return s;
}

// A student may only ever act on their own record. Passing someone else's
// studentID with a valid session is rejected.
function _requireStudent_(token, studentID) {
  const s = _requireSession_(token);
  if (s.t === 'admin') return s;                       // admins may act on any student
  if (s.t !== 'student') throw new Error('Not authorised.');
  if (_norm(s.id).toUpperCase() !== _norm(studentID).toUpperCase()) {
    throw new Error('Not authorised for this student record.');
  }
  return s;
}

function _requireInstructor_(token, instructorUsername) {
  const s = _requireSession_(token);
  if (s.t === 'admin') return s;
  if (s.t !== 'instructor') throw new Error('Not authorised.');
  if (instructorUsername !== undefined && instructorUsername !== null &&
      _norm(s.id).toLowerCase() !== _norm(instructorUsername).toLowerCase()) {
    throw new Error('Not authorised for this instructor account.');
  }
  return s;
}

function _requireAdmin_(token) {
  if (_isEditorContext_()) return { t: 'admin', id: 'EDITOR', role: 'SUPERADMIN' };
  const s = _requireSession_(token);
  if (s.t !== 'admin') throw new Error('Administrator access required.');
  return s;
}

// Issued by checkAdminLogin in Admin.gs after a verified password.
function _issueAdminSession_(username, role) {
  return _issueSession_('admin', username, role);
}

// ============================================================
// ONE-TIME MIGRATION — run manually from the Apps Script editor.
// Safe to re-run; already-hashed rows are skipped.
// ============================================================

// Hashes every remaining plaintext password in Master. Run this only AFTER
// the reset broadcast has gone out and the grace period has passed — it
// permanently destroys the plaintext, so anyone who has not logged in by
// then will need an admin reset.
function migratePlaintextPasswords() {
  const sh = _ss().getSheetByName('Master');
  const last = sh.getLastRow();
  if (last < 2) return { success: true, message: 'No rows to migrate.' };

  const rng  = sh.getRange(2, MASTER_PW_COL, last - 1, 1);
  const vals = rng.getValues();
  let hashed = 0, already = 0, blank = 0;

  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || '').trim();
    if (!v)             { blank++;   continue; }
    if (_isHashed_(v))  { already++; continue; }
    vals[i][0] = _hashPassword_(v);
    hashed++;
  }
  rng.setValues(vals);

  return {
    success: true,
    message: 'Hashed ' + hashed + ' password(s). ' +
             already + ' already hashed, ' + blank + ' blank.'
  };
}

// Reports how far the transparent upgrade has got, without changing anything.
// Use this to decide when it is safe to run migratePlaintextPasswords().
function auditPasswordStorage() {
  const sh = _ss().getSheetByName('Master');
  const last = sh.getLastRow();
  if (last < 2) return { success: true, total: 0, hashed: 0, plaintext: 0, blank: 0 };

  const vals = sh.getRange(2, MASTER_PW_COL, last - 1, 1).getValues();
  let hashed = 0, plaintext = 0, blank = 0;
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || '').trim();
    if (!v)            blank++;
    else if (_isHashed_(v)) hashed++;
    else               plaintext++;
  }
  return {
    success: true,
    total: vals.length,
    hashed: hashed,
    plaintext: plaintext,
    blank: blank,
    message: hashed + ' hashed / ' + plaintext + ' still plaintext / ' + blank + ' blank.'
  };
}

// ============================================================
// PASSWORD RESET BROADCAST
// Run from the Apps Script editor. Emails every student asking them to log
// in and set a new password.
//
// WHY IT BATCHES: MailApp on a consumer Gmail account allows 100 recipients
// per day (Workspace: 1500). A single pass over a large Master sheet would
// stop silently partway through. This sends up to a daily cap, records who
// has been reached in Script Properties, and picks up where it stopped the
// next time it runs. Safe to re-run — nobody is emailed twice.
// ============================================================

const BROADCAST_SENT_KEY = 'TWI_RESET_BROADCAST_SENT';

// Leaves headroom under the quota for the college's ordinary mail
// (welcome emails, COR, invoices) on the same day.
const BROADCAST_DAILY_CAP = 70;

function _broadcastSentSet_() {
  const raw = PropertiesService.getScriptProperties().getProperty(BROADCAST_SENT_KEY);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function _broadcastSentSave_(obj) {
  PropertiesService.getScriptProperties()
    .setProperty(BROADCAST_SENT_KEY, JSON.stringify(obj));
}

// How many are left to email, without sending anything.
function previewResetBroadcast() {
  const sh = _ss().getSheetByName('Master');
  const data = sh.getDataRange().getValues();
  const sent = _broadcastSentSet_();
  let pending = 0, already = 0, noEmail = 0;

  for (let i = 1; i < data.length; i++) {
    const sid = _norm(data[i][1]);
    if (!sid) continue;
    const email = String(data[i][2] || '').trim().toLowerCase();
    if (!email)        { noEmail++; continue; }
    if (sent[sid])     { already++; continue; }
    pending++;
  }
  return {
    success: true, pending: pending, alreadySent: already, missingEmail: noEmail,
    dailyCap: BROADCAST_DAILY_CAP,
    remainingQuotaToday: MailApp.getRemainingDailyQuota(),
    message: pending + ' student(s) still to email. ' + already + ' already done, ' +
             noEmail + ' have no email on file. Gmail quota left today: ' +
             MailApp.getRemainingDailyQuota() + '.'
  };
}

// Sends the next batch. Run once a day until previewResetBroadcast() reports
// pending: 0.
function sendResetBroadcast() {
  const sh = _ss().getSheetByName('Master');
  const data = sh.getDataRange().getValues();
  const sent = _broadcastSentSet_();

  const quota = MailApp.getRemainingDailyQuota();
  const budget = Math.min(BROADCAST_DAILY_CAP, quota);
  if (budget <= 0) {
    return { success: false, sentNow: 0,
             message: 'No Gmail quota left today. Try again tomorrow.' };
  }

  let sentNow = 0, skippedNoEmail = 0, failed = 0;

  for (let i = 1; i < data.length && sentNow < budget; i++) {
    const sid = _norm(data[i][1]);
    if (!sid || sent[sid]) continue;

    const email = String(data[i][2] || '').trim().toLowerCase();
    if (!email) { skippedNoEmail++; continue; }
    const firstName = String(data[i][4] || '').trim() || 'Student';

    try {
      MailApp.sendEmail(email,
        'Action needed: set a new Student Portal password',
        'Dear ' + firstName + ',\n\n' +
        'We have upgraded how Thy Word Intl Bible College Bataan stores Student ' +
        'Portal passwords. Passwords are now encrypted, which means no one — ' +
        'including college staff — can read yours.\n\n' +
        'Please log in to the Student Portal and set a new password:\n\n' +
        '  Student ID (Username): ' + sid + '\n\n' +
        'Your current password still works for now. Once you are logged in, open ' +
        'the Profile tab and use "Change Password" to set a new one.\n\n' +
        'If you have forgotten your password, use "Forgot Password" on the portal ' +
        'and we will email you a temporary one.\n\n' +
        'Please do this at your earliest convenience.\n\n' +
        'God bless you!\n\n' +
        'Thy Word Intl Bible College Bataan Administration');

      sent[sid] = new Date().toISOString();
      sentNow++;
    } catch (e) {
      failed++;
      console.error('Broadcast to ' + sid + ' failed: ' + e.message);
    }
  }

  _broadcastSentSave_(sent);

  const after = previewResetBroadcast();
  return {
    success: true, sentNow: sentNow, failed: failed, skippedNoEmail: skippedNoEmail,
    stillPending: after.pending,
    message: 'Emailed ' + sentNow + ' student(s). ' + after.pending +
             ' still pending' + (failed ? ', ' + failed + ' failed' : '') +
             '. Run again tomorrow if pending is above zero.'
  };
}

// Clears the "already emailed" log so the broadcast can be run again from
// scratch. Only needed if you deliberately want to re-notify everyone.
function resetBroadcastLog() {
  PropertiesService.getScriptProperties().deleteProperty(BROADCAST_SENT_KEY);
  return { success: true, message: 'Broadcast log cleared. Everyone will be emailed again.' };
}
