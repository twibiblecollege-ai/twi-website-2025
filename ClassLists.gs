// ============================================================
// TWI BIBLE COLLEGE — CLASS LISTS TO INSTRUCTORS
// ClassLists.gs  (new file — the other script files are unchanged)
//
// On the enrollment closing date at 12:00 PM, every instructor receives one
// email with a CSV class list for each subject they handle (Instructors
// sheet, Subject Handle columns E onward). The CSV has the same columns as
// the Admin Dashboard's Enrolled tab.
//
//   Settings sheet, column C = closing date
//     row "Current Semester"     -> TWI-Philippines term
//     row "Current Semester INT" -> International term (Qatar / Canada / Europe)
//
// Every student sent is written to the "Class List Log" sheet, so nobody is
// ever sent twice. Students who enroll after the closing time appear in the
// Admin Dashboard (Enrolled tab) as LATE ENROLLEES and are sent from there.
//
// ONE-TIME SETUP (Apps Script editor, signed in as twibiblecollege@gmail.com):
//   run setupClassListTimer  ->  click Allow.
// Emails are sent FROM the account that ran the setup.
// ============================================================

const BUILD_CLASSLISTS   = '2026-11-03a';
const CL_ADMIN_EMAIL     = 'twibiblecollege@gmail.com';
const CL_LOG_SHEET       = 'Class List Log';
const CL_LOG_HEADERS     = ['Timestamp','Semester','Term','Instructor','Instructor Email','Subject',
                            'Student ID','Student Name','Batch','Status'];
const CL_SEND_HOUR       = 12;      // 12:00 PM on the closing date
const CL_TIMER_MINUTES   = 5;       // how often the timer checks
const CL_AUTO_WINDOW_DAYS = 3;      // the timer only auto-sends within 3 days after closing
const CL_TERMS = [
  { term: 'PH',  key: 'current semester',     label: 'TWI-Philippines', campus: 'TWI-PHILIPPINES' },
  { term: 'INT', key: 'current semester int', label: 'International',   campus: 'TWI-QATAR' }
];

// ── Settings: semester + closing date per term ──────────────────
function _clParseClosing_(v) {
  if (v === '' || v === null || v === undefined) return null;
  let ymd = '';
  if (v instanceof Date) {
    ymd = Utilities.formatDate(v, _ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  } else {
    const t = _norm(v);
    let m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);                     // 11/9/2026 (month/day/year)
    if (m) ymd = m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);                           // 2026-11-09
    if (m) ymd = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  }
  if (!ymd) return null;
  const hh = ('0' + CL_SEND_HOUR).slice(-2);
  return Utilities.parseDate(ymd + ' ' + hh + ':00', Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function _clTerms_() {
  const sh = _ss().getSheetByName('Settings');
  const rows = sh ? sh.getDataRange().getValues() : [];
  return CL_TERMS.map(t => {
    let closingRaw = '';
    for (let i = 1; i < rows.length; i++) {
      if (_norm(rows[i][0]).toLowerCase() === t.key) { closingRaw = rows[i].length > 2 ? rows[i][2] : ''; break; }
    }
    const closeAt = _clParseClosing_(closingRaw);
    return Object.assign({}, t, {
      semester: _getCurrentSemesterValue_(t.campus) || '',
      closeAt: closeAt,
      closeText: closeAt ? Utilities.formatDate(closeAt, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a') : '',
      badClosing: !!_norm(closingRaw) && !closeAt
    });
  });
}

function _clSentKey_(term, semester) {
  return 'CL_MAIN|' + term + '|' + String(semester).toLowerCase();
}

// ── Instructors and their class lists ───────────────────────────
function _clInstructors_() {
  const d = _ss().getSheetByName('Instructors').getDataRange().getValues();
  const hdr = d[0].map(h => _norm(h).toLowerCase());
  let emailCol = hdr.indexOf('email');
  if (emailCol < 0) emailCol = 3;
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const username = _norm(d[i][0]);
    if (!username) continue;
    const handles = [];
    for (let c = 4; c <= 25; c++) {
      if (c === emailCol) continue;
      const s = _norm(d[i][c]);
      if (s && handles.indexOf(s) < 0) handles.push(s);
    }
    out.push({ username: username, fullName: _norm(d[i][2]), email: _norm(d[i][emailCol]).toLowerCase(),
               handles: handles, handlesLower: new Set(handles.map(h => h.toLowerCase())) });
  }
  return out;
}

// For one term: instructor username -> [{ subject, students: [...] }].
// Same rule as the Instructor Portal: the enrollment's instructor is this
// instructor, or it is TBA/blank and the subject is assigned to them.
// A subject is listed when it has students or is offered on that term's sheet.
function _clRosters_(t) {
  const ss = _ss();
  const sem = String(t.semester).toLowerCase();
  const isIntTerm = t.term === 'INT';
  const info = {};
  ss.getSheetByName('Master').getDataRange().getValues().slice(1).forEach(r => {
    const sid = _norm(r[1]);
    if (!sid) return;
    info[sid.toUpperCase()] = { studentID: sid, email: _norm(r[2]), surname: _norm(r[3]), name: _norm(r[4]),
      middleName: _norm(r[5]), contact: _norm(r[7]), facebook: _norm(r[15]), campus: _norm(r[27]).toUpperCase(),
      fullName: _buildStudentName_(r) };
  });

  const offered = new Set();
  const offSheet = ss.getSheetByName(isIntTerm ? 'INT Subjects' : 'Subjects');
  if (offSheet) {
    const cols = _resolveSubjectColumns_(offSheet);
    offSheet.getDataRange().getValues().slice(1).forEach(r => { const s = _norm(r[cols.subject]); if (s) offered.add(s.toLowerCase()); });
  }

  const instructors = _clInstructors_();
  const lists = {};
  instructors.forEach(ins => { lists[ins.username] = {}; });
  ss.getSheetByName('Enrollments').getDataRange().getValues().slice(1).forEach(r => {
    if (_norm(r[1]).toLowerCase() !== sem) return;
    if (_norm(r[7]).toUpperCase() !== 'ENROLLED') return;
    const st = info[_norm(r[2]).toUpperCase()];
    if (!st || _isIntCampus_(st.campus) !== isIntTerm) return;
    const subject = _norm(r[4]), subLower = subject.toLowerCase();
    if (!subject) return;
    instructors.forEach(ins => {
      if (!ins.handlesLower.has(subLower)) return;
      if (!_rowIsMine_(r[5], subLower, ins.fullName, ins.handlesLower)) return;
      const bucket = lists[ins.username][subLower] = lists[ins.username][subLower] || { subject: subject, students: [] };
      if (!bucket.students.some(x => x.studentID === st.studentID)) {
        bucket.students.push(Object.assign({ enrolledAt: r[0] instanceof Date ? r[0] : null }, st));
      }
    });
  });

  return instructors.map(ins => {
    const items = [];
    ins.handles.forEach(h => {
      const b = lists[ins.username][h.toLowerCase()];
      if (b || offered.has(h.toLowerCase())) {
        const students = b ? b.students : [];
        students.sort((a, c) => (a.surname + a.name).localeCompare(c.surname + c.name));
        items.push({ subject: b ? b.subject : h, students: students });
      }
    });
    return { instructor: ins, items: items };
  }).filter(x => x.items.length);
}

// Enrolled students that NO instructor's list covers — usually because the
// Instructors sheet (Subject Handles) does not match who teaches the subject
// in the Subjects / INT Subjects sheet. Shown as a warning so nobody is missed.
// -> [{ subject, teacher (from the enrollment), students }]
function _clUncovered_(t, rosters) {
  const ss = _ss(), sem = String(t.semester).toLowerCase(), isIntTerm = t.term === 'INT';
  const covered = new Set();
  rosters.forEach(x => x.items.forEach(it => it.students.forEach(s => covered.add(it.subject.toLowerCase() + '|' + s.studentID.toUpperCase()))));
  const campus = {};
  ss.getSheetByName('Master').getDataRange().getValues().slice(1).forEach(r => { if (_norm(r[1])) campus[_norm(r[1]).toUpperCase()] = _norm(r[27]).toUpperCase(); });
  const groups = {};
  ss.getSheetByName('Enrollments').getDataRange().getValues().slice(1).forEach(r => {
    if (_norm(r[1]).toLowerCase() !== sem || _norm(r[7]).toUpperCase() !== 'ENROLLED') return;
    const sid = _norm(r[2]).toUpperCase(), subject = _norm(r[4]);
    if (!subject || !(sid in campus) || _isIntCampus_(campus[sid]) !== isIntTerm) return;
    if (covered.has(subject.toLowerCase() + '|' + sid)) return;
    const k = subject.toLowerCase() + '|' + _norm(r[5]).toLowerCase();
    groups[k] = groups[k] || { subject: subject, teacher: _norm(r[5]) || '(blank)', ids: new Set() };
    groups[k].ids.add(sid);
  });
  return Object.keys(groups).map(k => ({ subject: groups[k].subject, teacher: groups[k].teacher, students: groups[k].ids.size }))
    .sort((a, b) => a.subject.localeCompare(b.subject));
}

// Already-sent students for a semester: "instructor|subject|studentID" -> {at, batch}
function _clSentMap_(semester) {
  const sh = _ss().getSheetByName(CL_LOG_SHEET);
  const map = {};
  if (!sh) return map;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (_norm(r[1]).toLowerCase() !== String(semester).toLowerCase()) return;
    if (!_norm(r[6]) || _norm(r[9]).toUpperCase().indexOf('SENT') !== 0) return;
    map[[_norm(r[3]).toLowerCase(), _norm(r[5]).toLowerCase(), _norm(r[6]).toUpperCase()].join('|')] =
      { at: r[0] instanceof Date ? r[0] : null, batch: _norm(r[8]) };
  });
  return map;
}

function _clKey_(ins, subject, sid) {
  return [ins.fullName.toLowerCase(), String(subject).toLowerCase(), String(sid).toUpperCase()].join('|');
}

// Same 9 columns as the Admin Dashboard's Enrolled tab (ENROLLED_HEADERS).
function _clCsvBlob_(subject, semester, students, late) {
  const q = v => { const s = String(v == null ? '' : v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const lines = [ENROLLED_HEADERS.join(',')].concat(students.map(s =>
    [s.studentID, subject, s.email, s.surname, s.name, s.middleName, s.contact, s.facebook, s.campus].map(q).join(',')));
  const name = (late ? 'LATE ENROLLEES - ' : '') + subject + ' - ' + semester + '.csv';
  return Utilities.newBlob('﻿' + lines.join('\r\n'), 'text/csv', name.replace(/[\\/:*?"<>|]/g, '-'));
}

// Sends one email to one instructor. items: [{subject, students}].
// testTo: send to this address instead and log nothing.
function _clSendOne_(t, ins, items, batch, testTo) {
  const late = batch === 'LATE';
  const to = testTo || ins.email;
  if (!to) return { ok: false, status: 'NO EMAIL' };
  const attachments = items.filter(it => it.students.length).map(it => _clCsvBlob_(it.subject, t.semester, it.students, late));
  const lines = items.map(it => '  • ' + it.subject + ' — ' + it.students.length + ' student(s)' +
    (it.students.length ? '' : ' (no enrollees)')).join('\n');
  const subject = (testTo ? '[TEST — would go to ' + (ins.email || 'NO EMAIL ON FILE') + '] ' : '') +
    (late ? 'LATE ENROLLEES — ' + items.map(it => it.subject).join(', ') + ' — ' + t.semester
          : 'Class List — ' + t.semester + ' (' + t.label + ')');
  const body = 'Dear ' + ins.fullName + ',\n\n' +
    (late
      ? 'The following student(s) enrolled after the enrollment closing (' + t.closeText + ') in the subject(s) you handle for ' +
        t.semester + ' (' + t.label + '). They are NOT on the class list you received earlier:\n\n'
      : 'Enrollment for ' + t.semester + ' (' + t.label + ') closed on ' + t.closeText + '.\n' +
        'Attached is the class list (CSV) for each subject you handle:\n\n') +
    lines + '\n\n' +
    (late ? 'Please add them to your class list.\n\n'
          : 'Students who enroll late will be sent to you separately, marked "LATE ENROLLEES".\n' +
            'You can also see your students anytime in the Instructor Portal.\n\n') +
    'God bless you!\n\nThy Word Intl Bible College Bataan\nOffice of the Registrar';
  MailApp.sendEmail({ to: to, subject: subject, body: body, attachments: attachments });
  if (!testTo) {
    const log = _getOrCreateSheet(CL_LOG_SHEET, CL_LOG_HEADERS);
    items.forEach(it => {
      if (!it.students.length) {
        log.appendRow([new Date(), t.semester, t.term, ins.fullName, to, it.subject, '', '', batch, 'SENT (no students)']);
      }
      it.students.forEach(s => log.appendRow([new Date(), t.semester, t.term, ins.fullName, to, it.subject,
                                              s.studentID, s.fullName, batch, 'SENT']));
    });
  }
  return { ok: true, status: 'SENT', count: items.reduce((a, it) => a + it.students.length, 0) };
}

// Sends the main class lists for one term (once per semester).
function _clSendMain_(t) {
  const sentMap = _clSentMap_(t.semester);
  const report = [];
  const rosters = _clRosters_(t);
  const uncovered = _clUncovered_(t, rosters);
  rosters.forEach(x => {
    // Students already logged for this instructor/subject are never sent again.
    const items = x.items.map(it => ({ subject: it.subject,
      students: it.students.filter(s => !sentMap[_clKey_(x.instructor, it.subject, s.studentID)]) }));
    let r;
    try { r = _clSendOne_(t, x.instructor, items, 'MAIN', ''); }
    catch (e) { r = { ok: false, status: 'ERROR: ' + e.message }; }
    if (!r.ok) {
      _getOrCreateSheet(CL_LOG_SHEET, CL_LOG_HEADERS).appendRow([new Date(), t.semester, t.term, x.instructor.fullName,
        x.instructor.email, items.map(it => it.subject).join(', '), '', '', 'MAIN', r.status]);
    }
    report.push(x.instructor.fullName + ' <' + (x.instructor.email || 'no email') + '>: ' +
      (r.ok ? items.map(it => it.subject + ' (' + it.students.length + ')').join(', ') : r.status));
  });
  PropertiesService.getScriptProperties().setProperty(_clSentKey_(t.term, t.semester), new Date().toISOString());
  try {
    MailApp.sendEmail(CL_ADMIN_EMAIL, 'Class lists sent — ' + t.semester + ' (' + t.label + ')',
      'Class lists for ' + t.semester + ' (' + t.label + ') were sent after enrollment closed (' + t.closeText + ').\n\n' +
      report.join('\n') + '\n\n' +
      (uncovered.length
        ? '⚠ NOT SENT TO ANY INSTRUCTOR — no instructor has these subjects in the Instructors sheet (Subject Handles):\n' +
          uncovered.map(u => '  • ' + u.subject + ' — ' + u.students + ' student(s), teacher in the Subjects sheet: ' + u.teacher).join('\n') +
          '\nAdd the subject to the right instructor in the Instructors sheet, then send it from the Admin Dashboard as late enrollees.\n\n'
        : '') +
      'Late enrollees can be sent from the Admin Dashboard → Enrolled tab → Class Lists to Instructors.\n' +
      'Every student sent is recorded in the "' + CL_LOG_SHEET + '" sheet.');
  } catch (e) {}
  return report;
}

// ── Timer (time-driven trigger) ──────────────────────────────────
function classListTick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const now = new Date();
    _clTerms_().forEach(t => {
      if (!t.semester || !t.closeAt) return;
      if (props.getProperty(_clSentKey_(t.term, t.semester))) return;           // already sent this semester
      if (now < t.closeAt) return;                                              // not closed yet
      if (now - t.closeAt > CL_AUTO_WINDOW_DAYS * 86400000) return;             // too old: send from the dashboard
      _clSendMain_(t);
    });
  } finally { lock.releaseLock(); }
}

// Run ONCE from the Apps Script editor (signed in as twibiblecollege@gmail.com).
function setupClassListTimer() {
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'classListTick') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('classListTick').timeBased().everyMinutes(CL_TIMER_MINUTES).create();
  const msg = 'Class list timer is ON (checks every ' + CL_TIMER_MINUTES + ' minutes). Emails will be sent from ' +
    Session.getEffectiveUser().getEmail() + '.\n' +
    _clTerms_().map(t => t.label + ': ' + (t.semester || '(no semester)') + ' — ' +
      (t.closeAt ? 'sends ' + t.closeText : t.badClosing ? 'closing date in column C is not a date' : 'no closing date in column C')).join('\n');
  Logger.log(msg);
  return msg;
}

function removeClassListTimer() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'classListTick') { ScriptApp.deleteTrigger(tr); n++; }
  });
  Logger.log('Removed ' + n + ' class list timer(s).');
}

// ── Admin Dashboard (Enrolled tab) ───────────────────────────────
function adminClassListStatus(token) {
  _requireAdmin_(token);
  const props = PropertiesService.getScriptProperties();
  let timerOn = null;
  try { timerOn = ScriptApp.getProjectTriggers().some(tr => tr.getHandlerFunction() === 'classListTick'); } catch (e) {}
  const terms = _clTerms_().map(t => {
    const sentAt = t.semester ? props.getProperty(_clSentKey_(t.term, t.semester)) : '';
    const sentMap = t.semester ? _clSentMap_(t.semester) : {};
    const rosters = t.semester ? _clRosters_(t) : [];
    const uncovered = t.semester ? _clUncovered_(t, rosters) : [];
    const rows = rosters.reduce((acc, x) => acc.concat(x.items.map(it => {
      const sent = it.students.filter(s => sentMap[_clKey_(x.instructor, it.subject, s.studentID)]).length;
      return { username: x.instructor.username, instructor: x.instructor.fullName, email: x.instructor.email,
               subject: it.subject, enrolled: it.students.length, sent: sent,
               late: sentAt ? it.students.length - sent : 0 };
    })), []);
    return { term: t.term, label: t.label, semester: t.semester, closeText: t.closeText, badClosing: t.badClosing,
             closed: !!(t.closeAt && new Date() >= t.closeAt),
             sentAt: sentAt ? Utilities.formatDate(new Date(sentAt), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a') : '',
             rows: rows, uncovered: uncovered };
  });
  return { success: true, build: BUILD_CLASSLISTS, timerOn: timerOn, adminEmail: CL_ADMIN_EMAIL, terms: terms };
}

function _clFind_(term, username, subject) {
  const t = _clTerms_().filter(x => x.term === term)[0];
  if (!t || !t.semester) return { error: 'Unknown term.' };
  const x = _clRosters_(t).filter(r => r.instructor.username === username)[0];
  const it = x && x.items.filter(i => i.subject.toLowerCase() === String(subject).toLowerCase())[0];
  if (!it) return { error: 'That instructor/subject was not found for ' + t.semester + '.' };
  return { t: t, ins: x.instructor, item: it };
}

// Who was already sent, and who is late, for one instructor + subject.
function adminClassListDetail(term, username, subject, token) {
  _requireAdmin_(token);
  const f = _clFind_(term, username, subject);
  if (f.error) return { success: false, message: f.error };
  const sentMap = _clSentMap_(f.t.semester);
  const fmt = d => d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a') : '';
  const sent = [], late = [];
  f.item.students.forEach(s => {
    const m = sentMap[_clKey_(f.ins, f.item.subject, s.studentID)];
    const row = { studentID: s.studentID, name: s.fullName, campus: s.campus, enrolledAt: fmt(s.enrolledAt) };
    if (m) sent.push(Object.assign(row, { sentAt: fmt(m.at), batch: m.batch })); else late.push(row);
  });
  const mainSent = !!PropertiesService.getScriptProperties().getProperty(_clSentKey_(f.t.term, f.t.semester));
  return { success: true, instructor: f.ins.fullName, email: f.ins.email, subject: f.item.subject,
           semester: f.t.semester, label: f.t.label, mainSent: mainSent, sent: sent, late: mainSent ? late : [],
           notYetSent: mainSent ? [] : late };
}

// Sends ONLY the students not yet sent, marked as LATE ENROLLEES.
function adminSendLateList(term, username, subject, token) {
  _requireAdmin_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Another class-list send is running. Try again in a moment.' };
  try {
    const f = _clFind_(term, username, subject);
    if (f.error) return { success: false, message: f.error };
    if (!PropertiesService.getScriptProperties().getProperty(_clSentKey_(f.t.term, f.t.semester))) {
      return { success: false, message: 'The class lists for ' + f.t.semester + ' have not been sent yet.' };
    }
    if (!f.ins.email) return { success: false, message: f.ins.fullName + ' has no email in the Instructors sheet.' };
    const sentMap = _clSentMap_(f.t.semester);
    const late = f.item.students.filter(s => !sentMap[_clKey_(f.ins, f.item.subject, s.studentID)]);
    if (!late.length) return { success: true, sent: 0, message: 'No late enrollees to send — everyone was already sent.' };
    _clSendOne_(f.t, f.ins, [{ subject: f.item.subject, students: late }], 'LATE', '');
    return { success: true, sent: late.length,
             message: 'Sent ' + late.length + ' late enrollee(s) in ' + f.item.subject + ' to ' + f.ins.fullName + ' (' + f.ins.email + ').' };
  } finally { lock.releaseLock(); }
}

// Sends the main lists now (e.g. if the timer was not set up in time).
function adminSendClassListsNow(term, token) {
  _requireAdmin_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Another class-list send is running. Try again in a moment.' };
  try {
    const t = _clTerms_().filter(x => x.term === term)[0];
    if (!t || !t.semester) return { success: false, message: 'Unknown term.' };
    if (PropertiesService.getScriptProperties().getProperty(_clSentKey_(t.term, t.semester))) {
      return { success: false, message: 'Already sent for ' + t.semester + '. Use "Send late enrollees" for new students.' };
    }
    if (!t.closeText) t.closeText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
    const report = _clSendMain_(t);
    return { success: true, message: 'Class lists sent to ' + report.length + ' instructor(s) for ' + t.semester + '. A summary was emailed to ' + CL_ADMIN_EMAIL + '.' };
  } finally { lock.releaseLock(); }
}

// Preview: sends what ONE instructor would receive to the admin address only.
// Nothing is logged, so it does not count as sent.
function adminSendClassListTest(term, username, token) {
  _requireAdmin_(token);
  const t = _clTerms_().filter(x => x.term === term)[0];
  if (!t || !t.semester) return { success: false, message: 'Unknown term.' };
  const all = _clRosters_(t);
  const x = all.filter(r => r.instructor.username === username)[0] ||
            all.slice().sort((a, b) => b.items.reduce((s, i) => s + i.students.length, 0) - a.items.reduce((s, i) => s + i.students.length, 0))[0];
  if (!x) return { success: false, message: 'No instructor has subjects for ' + t.semester + '.' };
  if (!t.closeText) t.closeText = '(closing date not set)';
  _clSendOne_(t, x.instructor, x.items, 'MAIN', CL_ADMIN_EMAIL);
  return { success: true, message: 'Test email for ' + x.instructor.fullName + ' sent to ' + CL_ADMIN_EMAIL + '. Nothing was recorded as sent.' };
}
