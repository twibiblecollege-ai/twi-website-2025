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
// Admin Dashboard (Enrolled tab) as LATE ENROLLEES and are sent from there
// ("Send" per subject, or "Send all unsent lists").
//
// ONE-TIME SETUP (Apps Script editor, signed in as twibiblecollege@gmail.com):
//   run setupClassListTimer  ->  click Allow.
// Emails are sent FROM the account that ran the setup.
// ============================================================

const BUILD_CLASSLISTS   = '2026-11-09a';
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
  // Why each group is not covered, and what to change in the Instructors sheet.
  const instructors = _clInstructors_();
  return Object.keys(groups).map(k => {
    const g = groups[k], subLower = g.subject.toLowerCase(), tLower = g.teacher.toLowerCase();
    const byName = instructors.filter(i => i.fullName.toLowerCase() === tLower);
    const byHandle = instructors.filter(i => i.handlesLower.has(subLower));
    let fix;
    if (byName.length) {
      fix = 'Add "' + g.subject + '" to the Subject Handles of ' + byName[0].fullName + ' (' + byName[0].username + ').';
    } else {
      fix = '"' + g.teacher + '" (the teacher on these enrollments) is not in the Instructors sheet. Add a row with Full Name exactly "' +
            g.teacher + '", an email, and "' + g.subject + '" in Subject Handles.' +
            (byHandle.length ? ' (' + byHandle.map(i => i.fullName).join(', ') + ' also has "' + g.subject +
              '" — if that is the same person, make the Full Name match instead.)' : '');
    }
    return { subject: g.subject, teacher: g.teacher, students: g.ids.size, fix: fix };
  }).sort((a, b) => a.subject.localeCompare(b.subject));
}

// A student is a LATE ENROLLEE only if they enrolled after the closing time.
function _clIsLate_(t, s) {
  return !!(t.closeAt && s.enrolledAt && s.enrolledAt.getTime() > t.closeAt.getTime());
}

// Appends many rows to the Class List Log in one write.
function _clLog_(rows) {
  if (!rows.length) return;
  const sh = _getOrCreateSheet(CL_LOG_SHEET, CL_LOG_HEADERS);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, CL_LOG_HEADERS.length).setValues(rows);
}

// Sends one email to one instructor after the class lists went out.
// items: [{ subject, students, first }]
//   first = true  -> the instructor never received this subject: a complete class list.
//   first = false -> students not on the list sent earlier (late enrollees, or
//                    students missed because the Instructors sheet was out of date).
// Logs each student: 'MAIN' (first list), 'LATE' (enrolled after closing) or 'ADDED'.
function _clSendUnsent_(t, ins, items) {
  const firsts = items.filter(it => it.first), adds = items.filter(it => !it.first);
  const lateCount = it => it.students.filter(s => _clIsLate_(t, s)).length;
  const allLate = it => lateCount(it) === it.students.length;
  const attachments = items.map(it => _clCsvBlob_(it.subject, t.semester, it.students,
    it.first ? '' : allLate(it) ? 'LATE ENROLLEES' : 'ADDITIONAL STUDENTS'));
  const subject = !adds.length ? 'Class List — ' + t.semester + ' (' + t.label + ')'
    : !firsts.length && adds.every(allLate) ? 'LATE ENROLLEES — ' + adds.map(it => it.subject).join(', ') + ' — ' + t.semester
    : 'Class List update — ' + t.semester + ' (' + t.label + ')';
  const body = 'Dear ' + ins.fullName + ',\n\n' +
    (firsts.length
      ? 'Attached is the class list (CSV) for ' + t.semester + ' (' + t.label + ') for each subject below:\n\n' +
        firsts.map(it => '  • ' + it.subject + ' — ' + it.students.length + ' student(s)').join('\n') + '\n\n'
      : '') +
    (adds.length
      ? 'The following student(s) are NOT on the class list you received earlier for ' + t.semester + ' (' + t.label + '). ' +
        'Please add them to your class list:\n\n' +
        adds.map(it => { const n = lateCount(it);
          return '  • ' + it.subject + ' — ' + it.students.length + ' student(s)' +
                 (n === it.students.length ? ' (late enrollees)' : n ? ' (' + n + ' late enrollee(s))' : ''); }).join('\n') + '\n\n'
      : '') +
    (t.closeText ? 'Enrollment closed on ' + t.closeText + '.\n' : '') +
    'You can also see your students anytime in the Instructor Portal.\n\n' +
    'God bless you!\n\nThy Word Intl Bible College Bataan\nOffice of the Registrar';
  MailApp.sendEmail({ to: ins.email, subject: subject, body: body, attachments: attachments });
  const now = new Date(), rows = [];
  items.forEach(it => it.students.forEach(s => rows.push([now, t.semester, t.term, ins.fullName, ins.email, it.subject,
    s.studentID, s.fullName, it.first ? 'MAIN' : _clIsLate_(t, s) ? 'LATE' : 'ADDED', 'SENT'])));
  _clLog_(rows);
}

// For one roster entry: the subjects with students not sent yet -> [{ subject, students, first }]
function _clUnsentItems_(x, sentMap) {
  return x.items.map(it => {
    const isSent = s => !!sentMap[_clKey_(x.instructor, it.subject, s.studentID)];
    return { subject: it.subject, students: it.students.filter(s => !isSent(s)), first: !it.students.some(isSent) };
  }).filter(it => it.students.length);
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
// late: true = "LATE ENROLLEES - " file name prefix; a string = that prefix.
function _clCsvBlob_(subject, semester, students, late) {
  const q = v => { const s = String(v == null ? '' : v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const lines = [ENROLLED_HEADERS.join(',')].concat(students.map(s =>
    [s.studentID, subject, s.email, s.surname, s.name, s.middleName, s.contact, s.facebook, s.campus].map(q).join(',')));
  const name = (late === true ? 'LATE ENROLLEES - ' : late ? late + ' - ' : '') + subject + ' - ' + semester + '.csv';
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
        ? '⚠ NOT SENT TO ANY INSTRUCTOR — these students match no instructor in the Instructors sheet:\n' +
          uncovered.map(u => '  • ' + u.subject + ' — ' + u.students + ' student(s). ' + u.fix).join('\n') +
          '\nFix the Instructors sheet, then send them from the Admin Dashboard ("Send all unsent lists").\n\n'
        : '') +
      'Students not sent yet (late enrollees, or students missed because the Instructors sheet changed) can be sent from the Admin Dashboard → Enrolled tab → Class Lists to Instructors.\n' +
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
      const unsent = it.students.filter(s => !sentMap[_clKey_(x.instructor, it.subject, s.studentID)]);
      const late = unsent.filter(s => _clIsLate_(t, s)).length;
      // unsent = not sent yet; late = of those, enrolled after closing; notSent = the rest
      return { username: x.instructor.username, instructor: x.instructor.fullName, email: x.instructor.email,
               subject: it.subject, enrolled: it.students.length, sent: it.students.length - unsent.length,
               unsent: sentAt ? unsent.length : 0, late: sentAt ? late : 0, notSent: sentAt ? unsent.length - late : 0 };
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
    const row = { studentID: s.studentID, name: s.fullName, campus: s.campus, enrolledAt: fmt(s.enrolledAt), isLate: _clIsLate_(f.t, s) };
    if (m) sent.push(Object.assign(row, { sentAt: fmt(m.at), batch: m.batch })); else late.push(row);
  });
  const mainSent = !!PropertiesService.getScriptProperties().getProperty(_clSentKey_(f.t.term, f.t.semester));
  const lr = _clLastResend_(f.t.semester, f.ins, f.item.subject);
  return { success: true, instructor: f.ins.fullName, email: f.ins.email, subject: f.item.subject,
           semester: f.t.semester, label: f.t.label, mainSent: mainSent, sent: sent, late: mainSent ? late : [],
           notYetSent: mainSent ? [] : late, lastResend: lr ? fmt(lr.at) + ' to ' + lr.to : '' };
}

// SEND (one subject): emails the students of this subject not sent yet.
// Never received this subject -> a complete class list; otherwise only the
// new students, marked late enrollees when they enrolled after closing.
function adminSendSubjectList(term, username, subject, token) {
  _requireAdmin_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Another class-list send is running. Try again in a moment.' };
  try {
    const f = _clFind_(term, username, subject);
    if (f.error) return { success: false, message: f.error };
    if (!PropertiesService.getScriptProperties().getProperty(_clSentKey_(f.t.term, f.t.semester))) {
      return { success: false, message: 'The class lists for ' + f.t.semester + ' have not been sent yet. Use "Send class lists now".' };
    }
    if (!f.ins.email) return { success: false, message: f.ins.fullName + ' has no email in the Instructors sheet.' };
    const items = _clUnsentItems_({ instructor: f.ins, items: [f.item] }, _clSentMap_(f.t.semester));
    if (!items.length) return { success: true, sent: 0, message: 'Nothing to send — everyone in ' + f.item.subject + ' was already sent.' };
    _clSendUnsent_(f.t, f.ins, items);
    const it = items[0], late = it.students.filter(s => _clIsLate_(f.t, s)).length;
    return { success: true, sent: it.students.length,
             message: (it.first ? 'Sent the class list of ' + it.subject + ' (' + it.students.length + ' student(s))'
                                : 'Sent ' + it.students.length + ' new student(s) in ' + it.subject + (late ? ' (' + late + ' late enrollee(s))' : '')) +
                      ' to ' + f.ins.fullName + ' (' + f.ins.email + ').' };
  } finally { lock.releaseLock(); }
}
function adminSendLateList(term, username, subject, token) { return adminSendSubjectList(term, username, subject, token); }

// SEND ALL: every instructor gets ONE email with every subject that has
// students not sent yet (same rules as adminSendSubjectList).
function adminSendAllUnsent(term, token) {
  _requireAdmin_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Another class-list send is running. Try again in a moment.' };
  try {
    const t = _clTerms_().filter(x => x.term === term)[0];
    if (!t || !t.semester) return { success: false, message: 'Unknown term.' };
    if (!PropertiesService.getScriptProperties().getProperty(_clSentKey_(t.term, t.semester))) {
      return { success: false, message: 'The class lists for ' + t.semester + ' have not been sent yet. Use "Send class lists now".' };
    }
    const sentMap = _clSentMap_(t.semester);
    const done = [], skipped = [], failed = [];
    let students = 0;
    _clRosters_(t).forEach(x => {
      const items = _clUnsentItems_(x, sentMap);
      if (!items.length) return;
      const line = x.instructor.fullName + ': ' + items.map(it => it.subject + ' (' + it.students.length + ')').join(', ');
      if (!x.instructor.email) { skipped.push(line); return; }
      try { _clSendUnsent_(t, x.instructor, items); done.push(line); students += items.reduce((n, it) => n + it.students.length, 0); }
      catch (e) { failed.push(line + ' — ' + e.message); }
    });
    if (!done.length && !skipped.length && !failed.length) return { success: true, message: 'Nothing to send — every student was already sent.' };
    try {
      MailApp.sendEmail(CL_ADMIN_EMAIL, 'Class lists sent (students not sent before) — ' + t.semester + ' (' + t.label + ')',
        'Sent from the Admin Dashboard ("Send all unsent lists"):\n\n' + (done.join('\n') || '(none)') +
        (skipped.length ? '\n\nNOT SENT — no email in the Instructors sheet:\n' + skipped.join('\n') : '') +
        (failed.length ? '\n\nFAILED:\n' + failed.join('\n') : '') +
        '\n\nEvery student sent is recorded in the "' + CL_LOG_SHEET + '" sheet.');
    } catch (e) {}
    return { success: !failed.length,
             message: 'Sent ' + students + ' student(s) to ' + done.length + ' instructor(s).' +
                      (skipped.length ? ' NOT sent (no email in the Instructors sheet): ' + skipped.join('; ') + '.' : '') +
                      (failed.length ? ' FAILED: ' + failed.join('; ') + '.' : '') +
                      ' A summary was emailed to ' + CL_ADMIN_EMAIL + '.' };
  } finally { lock.releaseLock(); }
}

// Last time the complete list for one instructor + subject was resent -> {at, to} or null.
function _clLastResend_(semester, ins, subject) {
  const sh = _ss().getSheetByName(CL_LOG_SHEET);
  if (!sh) return null;
  let last = null;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (_norm(r[8]).toUpperCase() !== 'RESEND') return;
    if (_norm(r[1]).toLowerCase() !== String(semester).toLowerCase()) return;
    if (_norm(r[3]).toLowerCase() !== ins.fullName.toLowerCase()) return;
    if (_norm(r[5]).toLowerCase() !== String(subject).toLowerCase()) return;
    if (r[0] instanceof Date && (!last || r[0] >= last.at)) last = { at: r[0], to: _norm(r[4]) };
  });
  return last;
}

// RESEND: emails the COMPLETE current list for one instructor + subject —
// everyone enrolled now (the class list + all late enrollees) in one CSV.
// For when an instructor changed email or lost the earlier email.
// toEmail: the address to send to (blank = the email in the Instructors sheet).
// Log: students never sent before are logged 'SENT' (so they stop counting as
// late); students already sent are logged 'RESENT' (history only).
function adminResendClassList(term, username, subject, toEmail, token) {
  _requireAdmin_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Another class-list send is running. Try again in a moment.' };
  try {
    const f = _clFind_(term, username, subject);
    if (f.error) return { success: false, message: f.error };
    const t = f.t, ins = f.ins, it = f.item;
    if (!PropertiesService.getScriptProperties().getProperty(_clSentKey_(t.term, t.semester))) {
      return { success: false, message: 'The class lists for ' + t.semester + ' have not been sent yet, so there is nothing to resend.' };
    }
    const to = _norm(toEmail || ins.email).toLowerCase();
    if (!to) return { success: false, message: ins.fullName + ' has no email in the Instructors sheet. Type the email address to send to.' };
    if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(to)) return { success: false, message: '"' + to + '" is not a valid email address.' };
    if (!it.students.length) return { success: false, message: 'No students are enrolled in ' + it.subject + '.' };

    const sentMap = _clSentMap_(t.semester);
    const isNew = s => !sentMap[_clKey_(ins, it.subject, s.studentID)];
    const newCount = it.students.filter(isNew).length;
    const now = new Date();
    const asOf = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
    const body = 'Dear ' + ins.fullName + ',\n\n' +
      'Here is the complete, up-to-date class list for ' + it.subject + ', ' + t.semester + ' (' + t.label + ').\n' +
      'It has all ' + it.students.length + ' student(s) enrolled as of ' + asOf + ', including late enrollees, ' +
      'and replaces the list(s) you received earlier.\n\n' +
      'You can also see your students anytime in the Instructor Portal.\n\n' +
      'God bless you!\n\nThy Word Intl Bible College Bataan\nOffice of the Registrar';
    MailApp.sendEmail({ to: to, subject: 'Class List (complete, resent) — ' + it.subject + ' — ' + t.semester,
                        body: body, attachments: [_clCsvBlob_(it.subject, t.semester, it.students, false)] });

    const log = _getOrCreateSheet(CL_LOG_SHEET, CL_LOG_HEADERS);
    it.students.forEach(s => log.appendRow([now, t.semester, t.term, ins.fullName, to, it.subject,
                                            s.studentID, s.fullName, 'RESEND', isNew(s) ? 'SENT' : 'RESENT']));
    const other = ins.email && to !== ins.email;
    return { success: true, sent: it.students.length,
             message: 'Resent the complete list of ' + it.subject + ' (' + it.students.length + ' student(s)' +
                      (newCount ? ', including ' + newCount + ' late enrollee(s) not sent before' : '') + ') to ' + ins.fullName + ' (' + to + ').' +
                      (other ? ' Note: the Instructors sheet still has ' + ins.email + ' — update it there so future lists go to the new email.' : '') };
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
      return { success: false, message: 'Already sent for ' + t.semester + '. Use "Send all unsent lists" or "Send" per subject for students not sent yet.' };
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
