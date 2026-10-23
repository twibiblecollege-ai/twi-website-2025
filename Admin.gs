// ========================================
// TWI BIBLE COLLEGE — ADMIN DASHBOARD BACKEND
// Admin.gs
// Requires: Code.gs (_ss, _norm, _getOrCreateSheet, _getCurrentSemesterValue_,
//           _getCourseCatalogIndex_, _resolveSubjectColumns_, _computeFeesByUnits_,
//           _buildStudentName_, _peso_, _safeFile_, formatTimeAsText, MASTER_HEADERS)
//           Academic.gs (getAcademicRecord, _rebuildStudentSubjects_, PH_ALLOWED_DEFAULT)
// ========================================

const BUILD_ADMIN = '2026-10-23f';   // version stamp — see getServerBuild() in Code.gs

const ADMIN_CAMPUSES = ['TWI-PHILIPPINES','TWI-QATAR','TWI-CANADA','TWI-EUROPE'];

const INVOICE_HEADERS = [
  'Invoice No','Timestamp','Student ID','Student Name','Email','Campus','Semester',
  'Subjects','Total Units','Registration Fee','Miscellaneous Fee','Tuition Fee',
  'Total Assessment','Amount Paid','Balance','Status','Proof URL','Invoice PDF URL','Remarks'
];
const ADMIN_HEADERS = ['Username','Password','Full Name','Role'];
const ENROLLED_HEADERS = [
  'Student ID','SUBJECTS ENROLLED','EMAIL','SURNAME','NAME','MIDDLE NAME',
  'CONTACT NUMBER','FACEBOOK NAME','CAMPUS'
];

const CLASS_SHEET_MAP = {
  '1st Year CCM':'1st Year CCM',
  '2nd Year CCM':'2nd Year CCM',
  'CCM Evening Class':'CCM Evening Class',
  'BCM':'BCM'
};

// ============================================================
// ONE-TIME REPAIR — run manually from the editor. NEVER on login.
// ============================================================
function adminSetup(token) {
  _requireAdmin_(token);
  const ss = _ss();

  // 1. Master headers
  const master = ss.getSheetByName('Master');
  if (_norm(master.getRange(1,26).getValue()) === '') master.getRange(1,26).setValue('Classification');
  if (_norm(master.getRange(1,31).getValue()) === '') master.getRange(1,31).setValue('Allowed Programs');
  if (_norm(master.getRange(1,32).getValue()) === '') master.getRange(1,32).setValue('Completed Subjects');

  // 2. Core sheets
  setupInstructorModule();
  ADMIN_CAMPUSES.forEach(c => _getOrCreateSheet(c, MASTER_HEADERS));
  ADMIN_CAMPUSES.forEach(c => {
    _ensureHeader_(ss.getSheetByName(c), 'Allowed Programs');
    _ensureHeader_(ss.getSheetByName(c), 'Completed Subjects');
  });
  ['1st Year CCM','2nd Year CCM','CCM Evening Class','Evening Class','BCM'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh) { _ensureHeader_(sh,'Allowed Programs'); _ensureHeader_(sh,'Completed Subjects'); }
  });

  // 3. INT Subjects is missing its Program column -> program index was hitting "Click to Enroll"
  const intSh = ss.getSheetByName('INT Subjects');
  if (intSh) {
    const hdrs = intSh.getRange(1,1,1,Math.max(intSh.getLastColumn(),1)).getValues()[0]
                   .map(h => _norm(h).toLowerCase());
    if (!hdrs.some(h => h === 'program')) {
      const col = intSh.getLastColumn() + 1;
      intSh.getRange(1, col).setValue('Program');
      const n = intSh.getLastRow() - 1;
      if (n > 0) intSh.getRange(2, col, n, 1).setValues(Array(n).fill(['INT']));
    }
  }

  // 4. Invoices + Admin + clean ENROLLED STUDENTS
  _getOrCreateSheet('Invoices', INVOICE_HEADERS);
  const adm = _getOrCreateSheet('Admin', ADMIN_HEADERS);
  let seededAdminPw = '';
  if (adm.getLastRow() < 2) {
    // No shipped default. Generate a random one-time password, store only
    // its hash, and hand the plaintext back to whoever ran adminSetup.
    seededAdminPw = _generateTempPassword_();
    adm.appendRow(['admin', _hashPassword_(seededAdminPw), 'Administrator', 'SUPERADMIN']);
  }
  _getOrCreateSheet('ENROLLED STUDENTS', ENROLLED_HEADERS);

  // 5. Enrollments semester -> match Settings exactly
  adminNormalizeSemesters(token);

  // 6. Rebuild AA/AF + backfill AE, then sync everything out
  migrateAcademicColumns();
  adminSyncAll(token);

  return {
    success: true,
    message: 'Admin setup complete.' + (seededAdminPw
      ? ' An "admin" account was created with the one-time password: ' + seededAdminPw +
        ' — log in and change it immediately. This is the only time it is shown.'
      : '')
  };
}

function _ensureHeader_(sh, name) {
  if (!sh) return;
  const last = Math.max(sh.getLastColumn(), 1);
  const hdrs = sh.getRange(1, 1, 1, last).getValues()[0].map(h => _norm(h).toLowerCase());
  if (hdrs.some(h => h === name.toLowerCase())) return;
  for (let i = hdrs.length - 1; i >= 0; i--) {
    if (hdrs[i] === '') { sh.getRange(1, i + 1).setValue(name); return; }
  }
  sh.getRange(1, last + 1).setValue(name);
}

// ============================================================
// HEADER-MAPPED SYNC — fixes the column-drift bug.
// Master has 31-32 cols; campus sheets had 32 with a DIFFERENT col 31,
// so positional copying wrote Allowed Programs into Pastor Recommendation Letter.
// ============================================================
function _headerIndex_(sh) {
  const last = Math.max(sh.getLastColumn(), 1);
  const hdrs = sh.getRange(1, 1, 1, last).getValues()[0];
  const map = {};
  hdrs.forEach((h, i) => { const k = _norm(h); if (k) map[k.toLowerCase()] = i; });
  return { map: map, width: last };
}

function _rowToObj_(headersArr, row) {
  const o = {};
  headersArr.forEach((h, i) => { const k = _norm(h); if (k) o[k] = (row[i] !== undefined ? row[i] : ''); });
  return o;
}

function _upsertByHeaders_(targetSh, masterObj, studentID) {
  if (!targetSh) return;
  const hi = _headerIndex_(targetSh);
  const out = new Array(hi.width).fill('');
  Object.keys(masterObj).forEach(k => {
    const i = hi.map[k.toLowerCase()];
    if (i !== undefined) out[i] = masterObj[k];
  });
  const sidCol = hi.map['student id'];
  if (sidCol === undefined) return;
  const data = targetSh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (_norm(data[r][sidCol]) === _norm(studentID)) {
      targetSh.getRange(r + 1, 1, 1, hi.width).setValues([out]);
      return;
    }
  }
  targetSh.appendRow(out);
}

function _removeFromSheet_(sh, studentID) {
  if (!sh) return;
  const hi = _headerIndex_(sh);
  const sidCol = hi.map['student id'];
  if (sidCol === undefined) return;
  const d = sh.getDataRange().getValues();
  for (let r = d.length - 1; r >= 1; r--) {
    if (_norm(d[r][sidCol]) === _norm(studentID)) sh.deleteRow(r + 1);
  }
}

function _propagateStudent_(studentID) {
  const ss = _ss();
  const master = ss.getSheetByName('Master');
  const data = master.getDataRange().getValues();
  const hdrs = data[0];
  let obj = null;
  for (let r = 1; r < data.length; r++) {
    if (_norm(data[r][1]) === _norm(studentID)) { obj = _rowToObj_(hdrs, data[r]); break; }
  }
  if (!obj) return;

  const campus = _norm(obj['Campus']).toUpperCase();          // Master col AB
  const classification = _norm(obj['Classification']);

  ADMIN_CAMPUSES.forEach(c => {
    const sh = ss.getSheetByName(c);
    if (c === campus) _upsertByHeaders_(sh, obj, studentID);
    else _removeFromSheet_(sh, studentID);
  });

  const target = CLASS_SHEET_MAP[classification];
  Object.keys(CLASS_SHEET_MAP).forEach(k => {
    const n = CLASS_SHEET_MAP[k];
    const sh = ss.getSheetByName(n);
    if (n === target) _upsertByHeaders_(sh, obj, studentID);
    else _removeFromSheet_(sh, studentID);
  });
}

function adminSyncAll(token) {
  _requireAdmin_(token);
  const ss = _ss();
  const data = ss.getSheetByName('Master').getDataRange().getValues();
  let n = 0;
  for (let r = 1; r < data.length; r++) {
    const sid = _norm(data[r][1]);
    if (sid) { _propagateStudent_(sid); n++; }
  }
  return { success:true, message:'Synced ' + n + ' student(s) to campus + class sheets.' };
}

// Fixes spelling variants of the CURRENT semester only ("1st sem  2025-2026 "
// -> "1st Sem 2025-2026"). Rows from other semesters are history and are
// never touched.
function adminNormalizeSemesters(token) {
  _requireAdmin_(token);
  const ss = _ss();
  const targets = [_getCurrentSemesterValue_(), _getCurrentSemesterValue_('TWI-QATAR')].filter(Boolean);
  if (!targets.length) return { success:false, message:'Current Semester is not set in Settings.' };
  const sh = ss.getSheetByName('Enrollments');
  const last = sh.getLastRow();
  if (last < 2) return { success:true, message:'No enrollment rows.' };
  const rng = sh.getRange(2, 2, last - 1, 1);
  const vals = rng.getValues();
  const key = s => String(s).toLowerCase().replace(/\s+/g, ' ').trim();
  let n = 0;
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0]);
    const t = targets.find(x => v !== x && key(v) === key(x));
    if (t) { vals[i][0] = t; n++; }
  }
  rng.setValues(vals);
  return { success:true, message:'Normalized ' + n + ' enrollment row(s) to "' + targets.join('" / "') + '".' };
}

// ============================================================
// AUTH — deliberately lightweight. Does NOT call adminSetup().
// ============================================================
function checkAdminLogin(username, password) {
  try {
    const ss = _ss();
    const sh = ss.getSheetByName('Admin');
    // Never auto-create an admin here. A login screen that provisions a
    // known-password superadmin when a sheet is missing is a back door.
    // Run adminSetup() from the editor to create the account.
    if (!sh) return { success:false, message:'Admin account is not set up. Run adminSetup() from the Apps Script editor.' };

    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (_norm(d[i][0]) !== _norm(username)) continue;

      const check = _verifyPassword_(password, d[i][1]);
      if (!check.ok) break;
      if (check.needsUpgrade) _upgradeStoredPassword_(sh, i + 1, 2, password);

      const uname = _norm(d[i][0]);
      const role  = _norm(d[i][3]);
      return { success:true, username:uname,
               fullName:_norm(d[i][2]) || uname, role:role,
               token: _issueAdminSession_(uname, role) };
    }
    return { success:false, message:'Invalid admin username or password.' };
  } catch (e) {
    return { success:false, message:'Login error: ' + e.message };
  }
}

function adminChangePassword(username, oldPw, newPw, token) {
  const sess = _requireAdmin_(token);
  // An admin may only change their own password, never another admin's.
  if (sess.id !== 'EDITOR' && _norm(sess.id).toLowerCase() !== _norm(username).toLowerCase()) {
    return { success:false, message:'You may only change your own password.' };
  }
  const sh = _ss().getSheetByName('Admin');
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (_norm(d[i][0]) === _norm(username)) {
      if (!_verifyPassword_(oldPw, d[i][1]).ok)
        return { success:false, message:'Current password is incorrect.' };
      if (String(newPw || '').length < 8)
        return { success:false, message:'New password must be at least 8 characters.' };
      sh.getRange(i + 1, 2).setValue(_hashPassword_(newPw));
      return { success:true, message:'Admin password updated.' };
    }
  }
  return { success:false, message:'Admin not found.' };
}

// ============================================================
// OVERVIEW
// ============================================================
// ============================================================
// SEMESTER PICKER
// ============================================================
// Blank / missing semester argument means "the current semester" — and the
// Philippines and international campuses each have their own current term.
function _semOrCurrent_(semester, campus) {
  return _norm(semester) || _getCurrentSemesterValue_(campus);
}

// Returns match(label, campus): does an Enrollments/Grades/Invoices row with
// this semester label belong to the semester being viewed? With a blank
// semester ("Current terms"), each row is compared with its own campus's term.
function _semMatcher_(semester) {
  const explicit = _norm(semester).toLowerCase();
  const ph  = String(_getCurrentSemesterValue_() || '').toLowerCase();
  const int = String(_getCurrentSemesterValue_('TWI-QATAR') || '').toLowerCase();
  return (label, campus) => {
    const l = _norm(label).toLowerCase();
    if (explicit) return l === explicit;
    return l === (_isIntCampus_(campus) ? int : ph);
  };
}

// Human label for the view, e.g. "S.Y 2026-2027 / S.Y 2026 Period (current)".
function _semViewLabel_(semester) {
  if (_norm(semester)) return _norm(semester);
  const ph = _getCurrentSemesterValue_(), int = _getCurrentSemesterValue_('TWI-QATAR');
  return (ph === int || !int ? ph : 'PH ' + ph + ' / INT ' + int) + ' (current)';
}

// Student ID -> campus, from Master.
function _campusBySid_(masterValues) {
  const map = {};
  for (let r = 1; r < masterValues.length; r++) {
    const sid = _norm(masterValues[r][1]);
    if (sid) map[sid] = _norm(masterValues[r][27]).toUpperCase();
  }
  return map;
}

// Every semester that has records, newest first, current semester on top.
// "Newest" = appears furthest down the Enrollments / Grades sheets.
function adminListSemesters(token) {
  _requireAdmin_(token);
  const ss = _ss();
  const current = _getCurrentSemesterValue_();
  const order = new Map();   // lowercased -> { label, rank }
  let rank = 0;
  const add = v => {
    const s = _norm(v); if (!s) return;
    const k = s.toLowerCase();
    rank++;
    order.set(k, { label: (order.get(k) || {}).label || s, rank: rank });
  };
  [['Enrollments', 1], ['Grades', 1]].forEach(([name, col]) => {
    const sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, col + 1, sh.getLastRow() - 1, 1).getValues().forEach(r => add(r[0]));
  });
  ['Subjects', 'INT Subjects'].forEach(name => {
    const sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 2) return;
    const cols = _resolveSubjectColumns_(sh);
    if (cols.semester == null || cols.semester < 0) return;
    sh.getDataRange().getValues().slice(1).forEach(r => add(r[cols.semester]));
  });

  // Sort by school year then term read from the label ("2nd Sem 2024-2025",
  // "Summer 2025"); labels without a year fall back to sheet order.
  const semKey = s => {
    const y = String(s).match(/(19|20)\d{2}/);
    const t = /summer|mid-?year|3rd|third/i.test(s) ? 3 : /2nd|second/i.test(s) ? 2 : 1;
    return y ? Number(y[0]) * 10 + t : -1;
  };
  const currentInt = _getCurrentSemesterValue_('TWI-QATAR');
  [current, currentInt].forEach(add);   // always listed, even with no rows yet
  const list = Array.from(order.values())
    .sort((a, b) => (semKey(b.label) - semKey(a.label)) || (b.rank - a.rank))
    .map(o => o.label);
  return { success:true, current: current, currentInt: currentInt, semesters: list };
}

function adminGetOverview(semester, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const m = ss.getSheetByName('Master').getDataRange().getValues();
  const match = _semMatcher_(semester);
  const campusOf = _campusBySid_(m);

  const counts = {};
  ADMIN_CAMPUSES.forEach(c => counts[c] = 0);
  for (let r = 1; r < m.length; r++) {
    const c = _norm(m[r][27]).toUpperCase();   // AB
    if (counts[c] !== undefined) counts[c]++;
  }

  const e = ss.getSheetByName('Enrollments').getDataRange().getValues();
  const enrolledStudents = new Set();
  let enrollRows = 0;
  for (let r = 1; r < e.length; r++) {
    if (!match(e[r][1], campusOf[_norm(e[r][2])])) continue;
    if (_norm(e[r][7]).toUpperCase() !== 'ENROLLED') continue;
    enrolledStudents.add(_norm(e[r][2]));
    enrollRows++;
  }

  let unpaid = 0, collected = 0;
  const invSh = ss.getSheetByName('Invoices');
  if (invSh) {
    const inv = invSh.getDataRange().getValues();
    for (let r = 1; r < inv.length; r++) {
      if (!_norm(inv[r][0])) continue;
      if (!match(inv[r][6], inv[r][5])) continue;   // G = Semester, F = Campus
      if (_norm(inv[r][15]).toUpperCase() !== 'PAID') unpaid++;
      collected += Number(inv[r][13]) || 0;
    }
  }

  return {
    success:true, semester:_semViewLabel_(semester),
    totalStudents: Math.max(m.length - 1, 0),
    campusCounts: counts,
    enrolledStudents: enrolledStudents.size,
    enrollRows: enrollRows,
    unpaidInvoices: unpaid,
    collected: collected
  };
}

// ============================================================
// STUDENTS
// ============================================================
function adminSearchStudents(query, campusFilter, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const d = ss.getSheetByName('Master').getDataRange().getValues();
  const q = _norm(query).toLowerCase();
  const cf = _norm(campusFilter).toUpperCase();
  const out = [];
  for (let r = 1; r < d.length; r++) {
    const sid = _norm(d[r][1]);
    if (!sid) continue;
    const campus = _norm(d[r][27]).toUpperCase();
    if (cf && cf !== 'ALL' && campus !== cf) continue;
    const blob = [sid, _norm(d[r][2]), _norm(d[r][3]), _norm(d[r][4]), _norm(d[r][5])].join(' ').toLowerCase();
    if (q && blob.indexOf(q) === -1) continue;
    out.push({
      studentID: sid, email: _norm(d[r][2]), surname: _norm(d[r][3]),
      firstName: _norm(d[r][4]), middleName: _norm(d[r][5]), mobile: _norm(d[r][7]),
      facebook: _norm(d[r][15]), classification: _norm(d[r][25]),
      subjects: _norm(d[r][26]), campus: campus,
      allowed: _norm(d[r][30]), completed: _norm(d[r][31])
    });
    if (out.length >= 300) break;
  }
  return { success:true, students: out, count: out.length };
}

function adminGetStudent(studentID, token) {
  _requireAdmin_(token);
  const sh = _ss().getSheetByName('Master');
  const d = sh.getDataRange().getValues();
  const hdrs = d[0];
  for (let r = 1; r < d.length; r++) {
    if (_norm(d[r][1]) === _norm(studentID)) {
      const o = _rowToObj_(hdrs, d[r]);
      Object.keys(o).forEach(k => {
        if (o[k] instanceof Date) {
          o[k] = Utilities.formatDate(o[k], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      });
      o['_row'] = r + 1;
      return { success:true, student:o, headers: hdrs.map(h => _norm(h)).filter(Boolean) };
    }
  }
  return { success:false, message:'Student not found.' };
}

// updates = { 'Header Name': value, ... }
function adminUpdateStudent(studentID, updates, token) {
  _requireAdmin_(token);
  try {
    const sh = _ss().getSheetByName('Master');
    const hi = _headerIndex_(sh);
    const d = sh.getDataRange().getValues();
    for (let r = 1; r < d.length; r++) {
      if (_norm(d[r][1]) !== _norm(studentID)) continue;
      const row = d[r].slice();
      while (row.length < hi.width) row.push('');
      Object.keys(updates || {}).forEach(k => {
        const i = hi.map[String(k).toLowerCase()];
        if (i !== undefined && i !== 1) row[i] = updates[k];   // never overwrite Student ID
      });
      sh.getRange(r + 1, 1, 1, hi.width).setValues([row]);
      _propagateStudent_(studentID);
      return { success:true, message:'Student updated and synced.' };
    }
    return { success:false, message:'Student not found.' };
  } catch (e) {
    return { success:false, message:'Update error: ' + e.message };
  }
}

function adminDeleteStudent(studentID, token) {
  _requireAdmin_(token);
  try {
    const ss = _ss();
    // Master (Student ID = col idx 1), Enrollments (col idx 2), ENROLLED STUDENTS (col idx 0)
    [['Master',1], ['Enrollments',2], ['ENROLLED STUDENTS',0], ['Grades',3]].forEach(function(pair) {
      const sh = ss.getSheetByName(pair[0]);
      if (!sh) return;
      const col = pair[1];
      const d = sh.getDataRange().getValues();
      for (let r = d.length - 1; r >= 1; r--) {
        if (_norm(d[r][col]) === _norm(studentID)) sh.deleteRow(r + 1);
      }
    });
    ADMIN_CAMPUSES.forEach(c => _removeFromSheet_(ss.getSheetByName(c), studentID));
    Object.keys(CLASS_SHEET_MAP).forEach(k => _removeFromSheet_(ss.getSheetByName(CLASS_SHEET_MAP[k]), studentID));
    return { success:true, message:'Student ' + studentID + ' deleted from all sheets.' };
  } catch (e) {
    return { success:false, message:'Delete error: ' + e.message };
  }
}

// The admin does NOT choose the password. A random single-use one is
// generated here, so staff never know a student's working password and
// cannot set the same weak one across many accounts.
function adminResetStudentPassword(studentID, unusedPw, token) {
  _requireAdmin_(token);

  const s = adminGetStudent(studentID, token);
  if (!s.success) return s;
  const email = _norm(s.student && s.student['Email']).toLowerCase();
  if (!email) return { success:false, message:'That student has no email address on file, so a reset cannot be delivered.' };

  const tempPassword = _generateTempPassword_();
  const r = adminUpdateStudent(studentID, { 'Password': _hashPassword_(tempPassword) }, token);
  if (!r.success) return r;

  MailApp.sendEmail(email,
    'Your Student Portal password has been reset',
    'Dear ' + _norm(s.student['First Name']) + ',\n\n' +
    'A college administrator has reset your Student Portal password.\n\n' +
    'Student ID: ' + studentID + '\n' +
    'Temporary Password: ' + tempPassword + '\n\n' +
    'Please log in with this temporary password and change it immediately from ' +
    'the Profile tab. Your previous password no longer works.\n\n' +
    'If you did not expect this reset, contact the college office at once.\n\n' +
    'God bless you!\n\nThy Word Intl Bible College Bataan');

  // The plaintext is deliberately not returned to the dashboard — it goes to
  // the student's own inbox and nowhere else.
  return { success:true, message:'A temporary password has been emailed to ' + email + '.' };
}

// ============================================================
// CAMPUS ROSTERS
// ============================================================
function adminGetCampusRoster(campus, semester, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const c = _norm(campus).toUpperCase();
  const sem = _semOrCurrent_(semester, c);
  const d = ss.getSheetByName('Master').getDataRange().getValues();

  const e = ss.getSheetByName('Enrollments').getDataRange().getValues();
  const enrolled = {};
  for (let r = 1; r < e.length; r++) {
    if (String(_norm(e[r][1])).toLowerCase() !== String(sem).toLowerCase()) continue;
    if (_norm(e[r][7]).toUpperCase() !== 'ENROLLED') continue;
    const sid = _norm(e[r][2]);
    if (!enrolled[sid]) enrolled[sid] = [];
    enrolled[sid].push(_norm(e[r][4]));
  }

  const out = [];
  for (let r = 1; r < d.length; r++) {
    if (_norm(d[r][27]).toUpperCase() !== c) continue;
    const sid = _norm(d[r][1]);
    out.push({
      studentID: sid, surname: _norm(d[r][3]), firstName: _norm(d[r][4]),
      middleName: _norm(d[r][5]), email: _norm(d[r][2]), mobile: _norm(d[r][7]),
      facebook: _norm(d[r][15]), classification: _norm(d[r][25]),
      currentSubjects: (enrolled[sid] || []).join(', '),
      completed: _norm(d[r][31]),
      isEnrolled: !!enrolled[sid]
    });
  }
  return { success:true, campus:c, semester:sem, students:out, count:out.length };
}

// ============================================================
// COURSE CATALOG  (single source of truth for subject titles)
// ============================================================
function adminGetCatalog(token) {
  _requireAdmin_(token);
  const sh = _ss().getSheetByName('COURSE_CATALOG');
  if (!sh) return { success:false, message:'COURSE_CATALOG sheet not found.' };
  const d = sh.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < d.length; r++) {
    const title = _norm(d[r][2]);
    if (!title) continue;
    rows.push({
      row: r + 1, program: _norm(d[r][0]), code: _norm(d[r][1]), title: title,
      units: Number(d[r][3]) || 3, yearLevel: _norm(d[r][4]),
      semester: _norm(d[r][5]), sort: Number(d[r][6]) || 0
    });
  }
  rows.sort((a,b) => a.program.localeCompare(b.program) || a.sort - b.sort);
  return { success:true, catalog: rows };
}

function adminSaveCatalogRow(item, token) {
  _requireAdmin_(token);
  const sh = _ss().getSheetByName('COURSE_CATALOG');
  const vals = [
    _norm(item.program).toUpperCase(), _norm(item.code), _norm(item.title),
    Number(item.units) || 3, _norm(item.yearLevel), _norm(item.semester), Number(item.sort) || 0
  ];
  if (!vals[2]) return { success:false, message:'Subject title is required.' };

  const d = sh.getDataRange().getValues();
  for (let r = 1; r < d.length; r++) {
    if (_norm(d[r][2]).toLowerCase() === vals[2].toLowerCase() && (r + 1) !== Number(item.row)) {
      return { success:false, message:'A catalog subject with that exact title already exists.' };
    }
  }
  if (item.row && Number(item.row) > 1) sh.getRange(Number(item.row), 1, 1, 7).setValues([vals]);
  else sh.appendRow(vals);
  return { success:true, message:'Catalog saved.' };
}

function adminDeleteCatalogRow(rowNum, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const sh = ss.getSheetByName('COURSE_CATALOG');
  const title = _norm(sh.getRange(Number(rowNum), 3).getValue());

  // Block deletion while the subject is still offered anywhere.
  let used = false;
  ['Subjects','INT Subjects'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (!s || used) return;
    const cols = _resolveSubjectColumns_(s);
    const d = s.getDataRange().getValues();
    for (let r = 1; r < d.length; r++) {
      if (_norm(d[r][cols.subject]).toLowerCase() === title.toLowerCase()) { used = true; break; }
    }
  });
  if (used) return { success:false, message:'Cannot delete: "' + title + '" is still offered in Subjects / INT Subjects.' };

  sh.deleteRow(Number(rowNum));
  return { success:true, message:'Catalog entry deleted.' };
}

// ============================================================
// SUBJECT OFFERINGS  (sheetName = 'Subjects' | 'INT Subjects')
// ============================================================
function adminGetOfferings(sheetName, token) {
  _requireAdmin_(token);
  const sh = _ss().getSheetByName(sheetName);
  if (!sh) return { success:false, message: sheetName + ' sheet not found.' };
  const cols = _resolveSubjectColumns_(sh);
  const d = sh.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < d.length; r++) {
    const s = _norm(d[r][cols.subject]);
    if (!s) continue;
    rows.push({
      row: r + 1, subject: s, day: _norm(d[r][cols.day]),
      from: formatTimeAsText(d[r][cols.from]), to: formatTimeAsText(d[r][cols.to]),
      instructor: _norm(d[r][cols.instructor]), semester: _norm(d[r][cols.semester]),
      yearLevel: _norm(d[r][cols.yearLevel]), campus: _norm(d[r][cols.campus]),
      program: _norm(d[r][cols.program])
    });
  }
  return { success:true, offerings: rows, sheet: sheetName };
}

function adminSaveOffering(sheetName, item, token) {
  _requireAdmin_(token);
  const sh = _ss().getSheetByName(sheetName);
  if (!sh) return { success:false, message: sheetName + ' sheet not found.' };

  const title = _norm(item.subject);
  if (!title) return { success:false, message:'Subject is required.' };

  // HARD GUARD: the title must exist verbatim in COURSE_CATALOG.
  const cat = _getCourseCatalogIndex_();
  if (!cat.map.has(title.toLowerCase())) {
    return { success:false,
      message:'"' + title + '" is not in COURSE_CATALOG. Add it to the catalog first — this is what prevents the subject-name errors.' };
  }

  const cols = _resolveSubjectColumns_(sh);
  let maxIdx = 0;
  Object.keys(cols).forEach(k => { if (cols[k] > maxIdx) maxIdx = cols[k]; });
  const width = Math.max(sh.getLastColumn(), maxIdx + 1);

  let out;
  if (item.row && Number(item.row) > 1) {
    out = sh.getRange(Number(item.row), 1, 1, width).getValues()[0];
  } else {
    out = new Array(width).fill('');
  }
  while (out.length < width) out.push('');

  out[cols.subject]    = title;
  out[cols.day]        = _norm(item.day);
  out[cols.from]       = _norm(item.from);
  out[cols.to]         = _norm(item.to);
  out[cols.instructor] = _norm(item.instructor);
  out[cols.semester]   = _norm(item.semester) ||
    _getCurrentSemesterValue_(sheetName === 'INT Subjects' ? 'TWI-QATAR' : '');
  out[cols.yearLevel]  = _norm(item.yearLevel);
  out[cols.campus]     = _norm(item.campus).toUpperCase();
  out[cols.program]    = _norm(item.program) || (sheetName === 'INT Subjects' ? 'INT' : '');

  if (item.row && Number(item.row) > 1) sh.getRange(Number(item.row), 1, 1, width).setValues([out]);
  else sh.appendRow(out);

  return { success:true, message:'Offering saved.' };
}

function adminDeleteOffering(sheetName, rowNum, token) {
  _requireAdmin_(token);
  const sh = _ss().getSheetByName(sheetName);
  if (!sh) return { success:false, message: sheetName + ' not found.' };
  sh.deleteRow(Number(rowNum));
  return { success:true, message:'Offering deleted.' };
}

function adminGetInstructorNames(token) {
  _requireAdmin_(token);
  const sh = _ss().getSheetByName('Instructors');
  if (!sh) return { success:true, instructors: [] };
  const d = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const n = _norm(d[i][2]);
    if (n) out.push(n);
  }
  return { success:true, instructors: out };
}

// ============================================================
// ENROLLED THIS SEMESTER — instructor handoff
// ============================================================
function adminGetEnrolledThisSemester(campusFilter, semester, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const match = _semMatcher_(semester);
  const cf = _norm(campusFilter).toUpperCase();
  const m = ss.getSheetByName('Master').getDataRange().getValues();
  const campusOf = _campusBySid_(m);

  const e = ss.getSheetByName('Enrollments').getDataRange().getValues();
  const bySid = {};
  for (let r = 1; r < e.length; r++) {
    if (!match(e[r][1], campusOf[_norm(e[r][2])])) continue;
    if (_norm(e[r][7]).toUpperCase() !== 'ENROLLED') continue;
    const sid = _norm(e[r][2]);
    const sub = _norm(e[r][4]);
    if (!sid || !sub) continue;
    if (!bySid[sid]) bySid[sid] = [];
    if (bySid[sid].indexOf(sub) === -1) bySid[sid].push(sub);
  }

  const rows = [];
  for (let r = 1; r < m.length; r++) {
    const sid = _norm(m[r][1]);
    if (!bySid[sid]) continue;
    const campus = _norm(m[r][27]).toUpperCase();
    if (cf && cf !== 'ALL' && campus !== cf) continue;
    rows.push({
      studentID: sid,
      subjects: bySid[sid].join(', '),
      email: _norm(m[r][2]),
      surname: _norm(m[r][3]),
      name: _norm(m[r][4]),
      middleName: _norm(m[r][5]),
      contact: _norm(m[r][7]),
      facebook: _norm(m[r][15]),
      campus: campus
    });
  }
  rows.sort((a,b) => a.campus.localeCompare(b.campus) || a.surname.localeCompare(b.surname));
  return { success:true, semester:_semViewLabel_(semester), rows:rows, count:rows.length };
}

// Writes the exact 9 columns into ENROLLED STUDENTS.
// subjectFilter is optional: pass a subject title to export only that class.
function adminExportEnrolled(campusFilter, subjectFilter, semester, token) {
  _requireAdmin_(token);
  const res = _norm(subjectFilter)
    ? adminGetSubjectRoster(subjectFilter, campusFilter, semester, token)
    : adminGetEnrolledThisSemester(campusFilter, semester, token);

  const sh = _getOrCreateSheet('ENROLLED STUDENTS', ENROLLED_HEADERS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 9).clearContent();

  const rows = res.rows || [];
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 9).setValues(rows.map(r => [
      r.studentID, r.subjects, r.email, r.surname, r.name,
      r.middleName, r.contact, r.facebook, r.campus
    ]));
  }
  return { success:true, count: rows.length,
    message:'Exported ' + rows.length + ' student(s) to the ENROLLED STUDENTS sheet' +
            (_norm(subjectFilter) ? ' for "' + _norm(subjectFilter) + '"' : '') +
            ' (' + res.semester + ').' };
}

// ============================================================
// INVOICES
// ============================================================
function _nextInvoiceNo_() {
  const sh = _ss().getSheetByName('Invoices');
  const last = sh.getLastRow();
  const yr = new Date().getFullYear();
  if (last < 2) return 'INV-' + yr + '-0001';
  const prev = _norm(sh.getRange(last, 1).getValue());
  const n = parseInt(String(prev).split('-')[2], 10);
  return 'INV-' + yr + '-' + String((isNaN(n) ? 0 : n) + 1).padStart(4, '0');
}

// semester (optional, 3rd argument so older callers still work): blank =
// the student's current semester. The invoice lists ONLY the subjects the
// student is enrolled in for that one semester, so two semesters never mix.
function adminBuildInvoiceDraft(studentID, token, semester) {
  _requireAdmin_(token);
  const ss = _ss();
  const m = ss.getSheetByName('Master').getDataRange().getValues();
  let row = null;
  for (let r = 1; r < m.length; r++) {
    if (_norm(m[r][1]).toUpperCase() === _norm(studentID).toUpperCase()) { row = m[r]; break; }
  }
  if (!row) return { success:false, message:'Student not found.' };
  const sid = _norm(row[1]);

  const campus = _norm(row[27]).toUpperCase();
  const isInt = (campus !== 'TWI-PHILIPPINES');
  const currentSem = _getCurrentSemesterValue_(campus);
  const sem = _norm(semester) || currentSem;

  // Every semester this student has (non-dropped) enrollments in.
  const { map: catalog } = _getCourseCatalogIndex_();
  const semSeen = new Map(), subjects = [], seenSub = new Set();
  const e = ss.getSheetByName('Enrollments').getDataRange().getValues();
  for (let r = 1; r < e.length; r++) {
    if (_norm(e[r][2]).toUpperCase() !== sid.toUpperCase()) continue;
    if (_norm(e[r][7]).toUpperCase() !== 'ENROLLED') continue;
    const rs = _norm(e[r][1]);
    if (rs && !semSeen.has(rs.toLowerCase())) semSeen.set(rs.toLowerCase(), rs);
    if (rs.toLowerCase() !== sem.toLowerCase()) continue;
    const title = _norm(e[r][4]);
    if (!title || seenSub.has(title.toLowerCase())) continue;
    seenSub.add(title.toLowerCase());
    const info = catalog.get(title.toLowerCase()) || {};
    subjects.push({ code: info.code || '', title: title, units: Number(info.units) || 3, instructor: _norm(e[r][5]) });
  }
  if (currentSem && !semSeen.has(currentSem.toLowerCase())) semSeen.set(currentSem.toLowerCase(), currentSem);
  const semesters = Array.from(semSeen.values()).sort((a, b) =>
    (b.toLowerCase() === currentSem.toLowerCase()) - (a.toLowerCase() === currentSem.toLowerCase()) ||
    _semesterSortKey_(b) - _semesterSortKey_(a));

  // Invoices already issued to this student for this semester.
  const existing = [];
  const inv = ss.getSheetByName('Invoices');
  if (inv) inv.getDataRange().getValues().slice(1).forEach(r => {
    if (_norm(r[2]).toUpperCase() === sid.toUpperCase() && _norm(r[6]).toLowerCase() === sem.toLowerCase()) {
      existing.push({ invoiceNo: _norm(r[0]), status: _norm(r[15]), total: Number(r[12]) || 0 });
    }
  });

  if (!subjects.length) {
    return { success:false, message:'No enrolled subjects for ' + sem + '.', semesters: semesters,
             semester: sem, currentSemester: currentSem };
  }
  const units = subjects.reduce((a, s) => a + s.units, 0);
  const fees = _computeFeesByUnits_(units, isInt);

  return { success:true, semesters: semesters, currentSemester: currentSem, existingInvoices: existing, draft: {
    studentID: sid,
    studentName: _buildStudentName_(row),
    email: _norm(row[2]).toLowerCase(),
    campus: campus, semester: sem,
    classification: _norm(row[25]),
    subjects: subjects, totalUnits: units, fees: fees
  }};
}

function adminCreateInvoice(studentID, opts, token) {
  _requireAdmin_(token);
  try {
    opts = opts || {};
    const dr = adminBuildInvoiceDraft(studentID, token, opts.semester);
    if (!dr.success) return dr;
    const d = dr.draft;

    const num = (v, fb) => (v !== undefined && v !== '' && v !== null) ? Number(v) : fb;
    const fees = {
      registrationFee:  num(opts.registrationFee,  d.fees.registrationFee),
      miscellaneousFee: num(opts.miscellaneousFee, d.fees.miscellaneousFee),
      tuitionFee:       num(opts.tuitionFee,       d.fees.tuitionFee)
    };
    fees.totalAssessment = num(opts.totalAssessment,
      fees.registrationFee + fees.miscellaneousFee + fees.tuitionFee);

    // Payment status chosen in the form (optional). It must agree with the
    // amount paid; PAID with no amount means the full total was paid.
    const want = _norm(opts.status).toUpperCase();
    let paid = Number(opts.amountPaid) || 0;
    if (want === 'PAID' && paid <= 0) paid = fees.totalAssessment;
    if (want === 'UNPAID' && paid > 0) return { success:false, message:'Status is Unpaid but an amount paid was entered. Set the amount to 0 or choose Partial/Paid.' };
    if (want === 'PARTIAL' && (paid <= 0 || paid >= fees.totalAssessment)) {
      return { success:false, message:'Partial needs an amount paid greater than 0 and less than the total (' + _peso_(fees.totalAssessment) + ').' };
    }
    if (want === 'PAID' && paid < fees.totalAssessment) return { success:false, message:'Paid needs the full total (' + _peso_(fees.totalAssessment) + '). Use Partial for less.' };
    if (opts.proofData && !/^data:(image\/[a-z+.-]+|application\/pdf);base64,/.test(String(opts.proofData))) {
      return { success:false, message:'The proof of payment must be an image or a PDF.' };
    }
    const balance = fees.totalAssessment - paid;
    const status = paid <= 0 ? 'UNPAID' : (balance <= 0 ? 'PAID' : 'PARTIAL');
    const invNo = _nextInvoiceNo_();

    const ctx = {
      invoiceNo: invNo, studentID: d.studentID, studentName: d.studentName,
      email: d.email, campus: d.campus, semester: d.semester,
      classification: d.classification, subjects: d.subjects, totalUnits: d.totalUnits,
      fees: fees, amountPaid: paid, balance: balance, status: status,
      remarks: _norm(opts.remarks)
    };

    const pdf = HtmlService.createHtmlOutput(_buildInvoiceHtml_(ctx))
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .getBlob().setName(invNo + '_' + d.studentID + '.pdf').getAs(MimeType.PDF);

    const folder = _getOrCreateInvoiceFolder_(d.semester);
    const file = folder.createFile(pdf);
    // Private. The invoice is delivered to the student as an email
    // attachment; a public link would expose their name, program and fees.
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

    _ss().getSheetByName('Invoices').appendRow([
      invNo, new Date(), d.studentID, d.studentName, d.email, d.campus, d.semester,
      d.subjects.map(s => s.title).join(', '), d.totalUnits,
      fees.registrationFee, fees.miscellaneousFee, fees.tuitionFee, fees.totalAssessment,
      paid, balance, status, '', file.getUrl(), _norm(opts.remarks)
    ]);

    if (opts.sendEmail !== false && d.email) {
      MailApp.sendEmail({
        to: d.email,
        subject: 'Statement of Account / Invoice ' + invNo + ' - ' + d.semester,
        body: 'Dear ' + d.studentName + ',\n\n' +
              'Attached is your invoice (' + invNo + ') for ' + d.semester + '.\n\n' +
              'Total Assessment: ' + _peso_(fees.totalAssessment) + '\n' +
              'Amount Paid: ' + _peso_(paid) + '\n' +
              'Balance: ' + _peso_(balance) + '\n' +
              'Status: ' + status + '\n\n' +
              'Online copy: ' + file.getUrl() + '\n\n' +
              'Send your payment via GCash (Joana A. Taguiam - 09088982181) or BDO Acct No: 013340053140.\n' +
              'After paying, reply with your proof of payment screenshot.\n\n' +
              'God bless you!\n\nThy Word Intl Bible College Bataan',
        attachments: [pdf]
      });
    }

    // Proof of payment uploaded together with the invoice (Partial / Paid).
    let proofNote = '';
    if (opts.proofData) {
      const pr = adminAttachProof(invNo, opts.proofData, paid, _norm(opts.remarks), token);
      proofNote = pr.success ? ' Proof of payment attached.' : ' (Proof NOT attached: ' + pr.message + ' — use the Proof button.)';
    }

    return { success:true, invoiceNo: invNo, fileUrl: file.getUrl(), status: status,
      message: 'Invoice ' + invNo + ' created (' + status + ')' + (opts.sendEmail !== false ? ' and emailed.' : '.') + proofNote };
  } catch (err) {
    return { success:false, message:'Invoice error: ' + err.message };
  }
}

function adminListInvoices(query, statusFilter, token) {
  _requireAdmin_(token);
  const sh = _ss().getSheetByName('Invoices');
  if (!sh) return { success:true, invoices: [], count: 0 };
  const d = sh.getDataRange().getValues();
  const q = _norm(query).toLowerCase();
  const sf = _norm(statusFilter).toUpperCase();
  const rows = [];
  for (let r = 1; r < d.length; r++) {
    const no = _norm(d[r][0]);
    if (!no) continue;
    const st = _norm(d[r][15]).toUpperCase();
    if (sf && sf !== 'ALL' && st !== sf) continue;
    const blob = [no, _norm(d[r][2]), _norm(d[r][3])].join(' ').toLowerCase();
    if (q && blob.indexOf(q) === -1) continue;
    rows.push({
      row: r + 1, invoiceNo: no,
      timestamp: (d[r][1] instanceof Date)
        ? Utilities.formatDate(d[r][1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : _norm(d[r][1]),
      studentID: _norm(d[r][2]), studentName: _norm(d[r][3]), email: _norm(d[r][4]),
      campus: _norm(d[r][5]), semester: _norm(d[r][6]), subjects: _norm(d[r][7]),
      units: Number(d[r][8]) || 0, total: Number(d[r][12]) || 0,
      paid: Number(d[r][13]) || 0, balance: Number(d[r][14]) || 0,
      status: st, proofUrl: _norm(d[r][16]), pdfUrl: _norm(d[r][17]), remarks: _norm(d[r][18])
    });
  }
  rows.reverse();
  return { success:true, invoices: rows, count: rows.length };
}

// ── PAYMENT MONITOR ──
// Every student enrolled in the semester being viewed, with the payment
// status of their invoice for that semester. A student with no invoice yet
// counts as UNPAID. Read-only: nothing is written, no email is sent.
function adminPaymentMonitor(semester, campusFilter, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const match = _semMatcher_(semester);
  const cf = _norm(campusFilter).toUpperCase();
  const m = ss.getSheetByName('Master').getDataRange().getValues();
  const info = {};
  for (let r = 1; r < m.length; r++) {
    const sid = _norm(m[r][1]); if (!sid) continue;
    info[sid.toUpperCase()] = { sid: sid, name: _buildStudentName_(m[r]), campus: _norm(m[r][27]).toUpperCase(),
                                classification: _norm(m[r][25]) };
  }
  const enrolled = new Map();     // sid -> { sem, subjects[] }
  const e = ss.getSheetByName('Enrollments').getDataRange().getValues();
  for (let r = 1; r < e.length; r++) {
    const st = info[_norm(e[r][2]).toUpperCase()]; if (!st) continue;
    if (_norm(e[r][7]).toUpperCase() !== 'ENROLLED') continue;
    if (!match(e[r][1], st.campus)) continue;
    if (cf && cf !== 'ALL' && st.campus !== cf) continue;
    const k = st.sid.toUpperCase();
    if (!enrolled.has(k)) enrolled.set(k, { sem: _norm(e[r][1]), subjects: [] });
    const sub = _norm(e[r][4]);
    if (sub && enrolled.get(k).subjects.indexOf(sub) < 0) enrolled.get(k).subjects.push(sub);
  }
  const invByKey = new Map();     // sid|semester -> latest invoice
  const inv = ss.getSheetByName('Invoices');
  if (inv) inv.getDataRange().getValues().slice(1).forEach(r => {
    if (!_norm(r[0])) return;
    invByKey.set(_norm(r[2]).toUpperCase() + '|' + _norm(r[6]).toLowerCase(), {
      invoiceNo: _norm(r[0]), studentName: _norm(r[3]), total: Number(r[12]) || 0, paid: Number(r[13]) || 0,
      balance: Number(r[14]) || 0, status: _norm(r[15]).toUpperCase() || 'UNPAID', proofUrl: _norm(r[16]),
      pdfUrl: _norm(r[17]), remarks: _norm(r[18]) });
  });
  const rows = [];
  const counts = { PAID: 0, PARTIAL: 0, UNPAID: 0, noInvoice: 0 };
  enrolled.forEach((en, k) => {
    const st = info[k], iv = invByKey.get(k + '|' + en.sem.toLowerCase()) || null;
    const status = iv ? (['PAID','PARTIAL'].indexOf(iv.status) >= 0 ? iv.status : 'UNPAID') : 'UNPAID';
    counts[status]++; if (!iv) counts.noInvoice++;
    rows.push({ studentID: st.sid, name: st.name, campus: st.campus, classification: st.classification,
                semester: en.sem, subjects: en.subjects, status: status, invoice: iv });
  });
  const order = { UNPAID: 0, PARTIAL: 1, PAID: 2 };
  rows.sort((a, b) => order[a.status] - order[b.status] || a.name.localeCompare(b.name));
  return { success: true, semester: _semViewLabel_(semester), rows: rows, counts: counts, total: rows.length };
}

// Attach the student's proof-of-payment screenshot. imageData = data URL.
function adminAttachProof(invoiceNo, imageData, amountPaid, remarks, token) {
  _requireAdmin_(token);
  try {
    const sh = _ss().getSheetByName('Invoices');
    const d = sh.getDataRange().getValues();
    let rowNum = -1, rec = null;
    for (let r = 1; r < d.length; r++) {
      if (_norm(d[r][0]) === _norm(invoiceNo)) { rowNum = r + 1; rec = d[r]; break; }
    }
    if (rowNum < 0) return { success:false, message:'Invoice not found.' };

    let proofUrl = _norm(rec[16]);
    if (imageData) {
      const mt = String(imageData).match(/^data:(.+);base64,(.+)$/);
      if (!mt) return { success:false, message:'Invalid image data.' };
      const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
      const blob = Utilities.newBlob(Utilities.base64Decode(mt[2]), mt[1],
        invoiceNo + '_proof_' + stamp);
      const folder = _getOrCreateProofFolder_(_norm(rec[6]));
      const f = folder.createFile(blob);
      // Private, and never anything else. These are GCash/bank screenshots
      // containing account numbers and reference numbers. Staff open them
      // from Drive while signed in to the college account.
      f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      proofUrl = f.getUrl();
    }

    const total = Number(rec[12]) || 0;
    const paid = (amountPaid !== undefined && amountPaid !== '' && amountPaid !== null)
      ? Number(amountPaid) : (Number(rec[13]) || 0);
    const balance = total - paid;
    const status = paid <= 0 ? 'UNPAID' : (balance <= 0 ? 'PAID' : 'PARTIAL');

    sh.getRange(rowNum, 14).setValue(paid);
    sh.getRange(rowNum, 15).setValue(balance);
    sh.getRange(rowNum, 16).setValue(status);
    sh.getRange(rowNum, 17).setValue(proofUrl);
    if (remarks !== undefined && _norm(remarks)) sh.getRange(rowNum, 19).setValue(_norm(remarks));

    return { success:true, proofUrl: proofUrl, status: status, balance: balance,
             message:'Proof attached. Status: ' + status + '.' };
  } catch (e) {
    return { success:false, message:'Attach error: ' + e.message };
  }
}

function adminEmailReceipt(invoiceNo, token) {
  _requireAdmin_(token);
  const d = _ss().getSheetByName('Invoices').getDataRange().getValues();
  for (let r = 1; r < d.length; r++) {
    if (_norm(d[r][0]) !== _norm(invoiceNo)) continue;
    const email = _norm(d[r][4]).toLowerCase();
    if (!email) return { success:false, message:'No email on file for this invoice.' };
    MailApp.sendEmail(email,
      'Payment Received - Invoice ' + invoiceNo,
      'Dear ' + _norm(d[r][3]) + ',\n\nWe have received your payment.\n\n' +
      'Invoice: ' + invoiceNo + '\n' +
      'Semester: ' + _norm(d[r][6]) + '\n' +
      'Total Assessment: ' + _peso_(Number(d[r][12]) || 0) + '\n' +
      'Amount Paid: ' + _peso_(Number(d[r][13]) || 0) + '\n' +
      'Balance: ' + _peso_(Number(d[r][14]) || 0) + '\n' +
      'Status: ' + _norm(d[r][15]) + '\n\n' +
      'Thank you and God bless!\n\nThy Word Intl Bible College Bataan');
    return { success:true, message:'Receipt emailed to ' + email };
  }
  return { success:false, message:'Invoice not found.' };
}

function _getOrCreateInvoiceFolder_(sem) {
  const rootIt = DriveApp.getFoldersByName('TWI_COR');
  const root = rootIt.hasNext() ? rootIt.next() : DriveApp.createFolder('TWI_COR');
  const it = root.getFoldersByName('INVOICES');
  const inv = it.hasNext() ? it.next() : root.createFolder('INVOICES');
  const s = _safeFile_(sem || 'UNKNOWN');
  const sIt = inv.getFoldersByName(s);
  return sIt.hasNext() ? sIt.next() : inv.createFolder(s);
}

function _getOrCreateProofFolder_(sem) {
  const f = _getOrCreateInvoiceFolder_(sem);
  const it = f.getFoldersByName('PROOF_OF_PAYMENT');
  return it.hasNext() ? it.next() : f.createFolder('PROOF_OF_PAYMENT');
}

function _buildInvoiceHtml_(c) {
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');
  const rows = c.subjects.map(s =>
    '<tr><td class="td">' + (s.code || '') + '</td><td class="td">' + s.title +
    '</td><td class="td ctr">' + s.units + '</td><td class="td">' + (s.instructor || '') + '</td></tr>'
  ).join('');
  const badge = c.status === 'PAID' ? '#1a7f4b' : (c.status === 'PARTIAL' ? '#b8860b' : '#b03030');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
  '@page{size:A4;margin:16mm}body{font-family:Arial,sans-serif;color:#111;font-size:12px}' +
  '.hd{text-align:center}.sch{font-size:16px;font-weight:800}.sub{font-size:11px}' +
  '.ttl{margin:12px 0;text-align:center;font-size:15px;font-weight:800;letter-spacing:1px}' +
  '.box{border:1px solid #333;padding:10px;margin-bottom:10px}' +
  '.row{display:flex;gap:12px}.col{flex:1}' +
  '.lb{font-size:10px;color:#444}.vl{font-size:12px;font-weight:700}' +
  'table{width:100%;border-collapse:collapse}' +
  '.th{background:#eaeaea;border:1px solid #333;padding:6px;font-size:10px;text-align:left}' +
  '.td{border:1px solid #333;padding:6px;font-size:11px}.ctr{text-align:center}.rt{text-align:right}' +
  '.fees{width:58%;margin-top:10px;margin-left:auto}' +
  '.badge{display:inline-block;padding:4px 14px;border-radius:12px;color:#fff;font-weight:800;font-size:11px;background:' + badge + '}' +
  '.ft{margin-top:14px;font-size:10px;color:#444;line-height:1.6}' +
  '</style></head><body>' +
  '<div class="hd"><div class="sch">THY WORD INTL BIBLE COLLEGE BATAAN</div>' +
  '<div class="sub">205 Dahlia Street, DFS, Balanga City, Bataan | twibiblecollege@gmail.com</div></div>' +
  '<div class="ttl">STATEMENT OF ACCOUNT / INVOICE</div>' +
  '<div class="box"><div class="row">' +
  '<div class="col"><div class="lb">Invoice No.</div><div class="vl">' + c.invoiceNo + '</div></div>' +
  '<div class="col"><div class="lb">Student No.</div><div class="vl">' + c.studentID + '</div></div>' +
  '<div class="col"><div class="lb">Semester</div><div class="vl">' + c.semester + '</div></div>' +
  '<div class="col"><div class="lb">Status</div><div class="vl"><span class="badge">' + c.status + '</span></div></div>' +
  '</div><div class="row" style="margin-top:8px">' +
  '<div class="col"><div class="lb">Student Name</div><div class="vl">' + c.studentName + '</div></div>' +
  '<div class="col"><div class="lb">Campus</div><div class="vl">' + c.campus + '</div></div>' +
  '<div class="col"><div class="lb">Program</div><div class="vl">' + (c.classification || '') + '</div></div>' +
  '</div></div>' +
  '<div class="box"><table><thead><tr>' +
  '<th class="th" style="width:14%">CODE</th><th class="th">SUBJECT</th>' +
  '<th class="th ctr" style="width:10%">UNITS</th><th class="th" style="width:28%">FACULTY</th>' +
  '</tr></thead><tbody>' + rows +
  '<tr><td class="td rt" colspan="2"><b>TOTAL UNITS</b></td><td class="td ctr"><b>' + c.totalUnits +
  '</b></td><td class="td"></td></tr></tbody></table>' +
  '<table class="fees">' +
  '<tr><td class="td">Registration Fee</td><td class="td rt">' + _peso_(c.fees.registrationFee) + '</td></tr>' +
  '<tr><td class="td">Miscellaneous Fee</td><td class="td rt">' + _peso_(c.fees.miscellaneousFee) + '</td></tr>' +
  '<tr><td class="td">Tuition Fee</td><td class="td rt">' + _peso_(c.fees.tuitionFee) + '</td></tr>' +
  '<tr><td class="td"><b>TOTAL ASSESSMENT</b></td><td class="td rt"><b>' + _peso_(c.fees.totalAssessment) + '</b></td></tr>' +
  '<tr><td class="td">Amount Paid</td><td class="td rt">' + _peso_(c.amountPaid) + '</td></tr>' +
  '<tr><td class="td"><b>BALANCE DUE</b></td><td class="td rt"><b>' + _peso_(c.balance) + '</b></td></tr>' +
  '</table>' +
  '<div class="ft">Date Issued: ' + dt + (c.remarks ? '<br/>Remarks: ' + c.remarks : '') + '<br/><br/>' +
  '<b>Payment Options</b><br/>GCash - Joana A. Taguiam | 09088982181<br/>' +
  'BDO - Joana A. Taguiam | Acct No: 013340053140<br/>' +
  'After paying, send your proof of payment to twibiblecollege@gmail.com to confirm your enrollment.</div>' +
  '</div></body></html>';
}

// ============================================================
// OLD GRADES — encode grades from before the system existed
// ------------------------------------------------------------
// Old grades go into the Grades sheet (the same place instructor grades go),
// so they appear in the student's Milestone and, through GR_LOOKUP, on the
// TOR. They never touch Enrollments, so current enrollment, load limits,
// instructor class lists and invoices are unaffected.
// Each row is marked in Grades column I ("Source") so admin-encoded grades
// can be told apart from portal grades, and only they can be edited here.
// ============================================================
const OLD_GRADE_SOURCE     = 'Old record (admin)';
const OLD_GRADE_INSTRUCTOR = 'Registrar (old record)';
const OLD_GRADES_MAX_ROWS  = 300;      // per bulk paste (Apps Script time limit)
const OLD_STUDENT_CLASSES  = ['Alumni', '1st Year CCM', '2nd Year CCM', 'BCM', 'CCM Evening Class', 'INT'];
const GRADE_HISTORY_HEADERS = ['Timestamp','Semester','Subject','Student ID','Student Name','Instructor',
                               'Old Grade','New Grade','Old Remarks','New Remarks'];

function _gradesSheetWithSource_() {
  const sh = _ss().getSheetByName('Grades');
  if (_norm(sh.getRange(1, 9).getValue()) === '') sh.getRange(1, 9).setValue('Source');
  return sh;
}

function _logGradeHistory_(row) {
  _getOrCreateSheet('Grade History', GRADE_HISTORY_HEADERS).appendRow(row);
}

function _levenshtein_(a, b) {
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[n];
}

// Finds a COURSE_CATALOG subject by exact title (ignoring case and extra
// spaces) or by course code. Never guesses: a near miss only produces a
// "did you mean" suggestion and the row is rejected.
function _catalogMatcher_() {
  const { map } = _getCourseCatalogIndex_();
  const squash = s => _norm(s).toLowerCase().replace(/\s+/g, ' ');
  const byTitle = new Map(), byCode = new Map();
  map.forEach(v => { byTitle.set(squash(v.title), v); if (v.code) byCode.set(squash(v.code), v); });
  const titles = Array.from(map.values()).map(v => v.title);
  return {
    find: s => byTitle.get(squash(s)) || byCode.get(squash(s)) || null,
    suggest: s => {
      const q = squash(s);
      let best = '', bestD = 1e9;
      titles.forEach(t => { const d = _levenshtein_(q, t.toLowerCase()); if (d < bestD) { bestD = d; best = t; } });
      return bestD <= Math.max(3, Math.round(q.length * 0.3)) ? best : '';
    }
  };
}

function _checkGradeValue_(g) {
  const v = _norm(g);
  if (!v) return { ok: false, msg: 'Grade is empty' };
  const n = Number(v);
  if (isFinite(n)) {
    // Point scale 1.00–5.00, or a percentage 50–100. Anything in between
    // (e.g. 7.5) is almost certainly a typo, so it is rejected.
    if ((n >= 1 && n <= 5) || (n >= 50 && n <= 100)) return { ok: true, value: v };
    return { ok: false, msg: 'Grade "' + v + '" looks wrong — use 1.00–5.00 or a percentage from 50 to 100' };
  }
  const u = v.toUpperCase();
  if (['P','PASS','PASSED','S','CR','F','FAIL','FAILED','INC','DRP','D'].indexOf(u) >= 0) return { ok: true, value: u };
  return { ok: false, msg: 'Unrecognised grade "' + v + '"' };
}

// Pasted from Excel / Google Sheets (tab-separated) or CSV.
// Columns: Student ID | Subject | Grade | Semester | Remarks (optional) | Instructor (optional)
function _parseOldGradesText_(text) {
  const rows = [];
  String(text || '').split(/\r?\n/).forEach((line, i) => {
    if (!line.trim()) return;
    const cells = (line.indexOf('\t') >= 0 ? line.split('\t') : line.split(','))
      .map(c => c.trim().replace(/^"(.*)"$/, '$1').trim());
    if (rows.length === 0 && /student/i.test(cells[0] || '')) return;   // header row
    rows.push({ line: i + 1, studentID: cells[0] || '', subject: cells[1] || '', grade: cells[2] || '',
                semester: cells[3] || '', remarks: cells[4] || '', instructor: cells[5] || '' });
  });
  return rows;
}

// Validates rows WITHOUT writing anything. Used by the single form, the bulk
// "Check" step, and again by every save (the browser is never trusted).
function _validateOldGrades_(rows) {
  const ss = _ss();
  const master = ss.getSheetByName('Master').getDataRange().getValues();
  const students = {};
  for (let r = 1; r < master.length; r++) {
    const id = _norm(master[r][1]);
    if (!id) continue;
    students[id.toUpperCase()] = { id: id, name: _buildStudentName_(master[r]), campus: _norm(master[r][27]).toUpperCase() };
  }
  const cat = _catalogMatcher_();

  const existing = new Map();     // sid|subject -> [{semester, grade}]
  const g = ss.getSheetByName('Grades').getDataRange().getValues();
  for (let r = 1; r < g.length; r++) {
    const sid = _norm(g[r][3]).toUpperCase(), sub = _norm(g[r][2]).toLowerCase();
    if (!sid || !sub || !_norm(g[r][6])) continue;
    const k = sid + '|' + sub;
    if (!existing.has(k)) existing.set(k, []);
    existing.get(k).push({ semester: _norm(g[r][1]), grade: _norm(g[r][6]) });
  }

  const semPH = String(_getCurrentSemesterValue_() || '').toLowerCase();
  const semINT = String(_getCurrentSemesterValue_('TWI-QATAR') || '').toLowerCase();
  const enrolledNow = new Set();  // sid|subject currently enrolled
  const e = ss.getSheetByName('Enrollments').getDataRange().getValues();
  for (let r = 1; r < e.length; r++) {
    const s = students[_norm(e[r][2]).toUpperCase()];
    if (!s) continue;
    const cur = _isIntCampus_(s.campus) ? semINT : semPH;
    if (_norm(e[r][1]).toLowerCase() === cur && _norm(e[r][7]).toUpperCase() === 'ENROLLED') {
      enrolledNow.add(s.id.toUpperCase() + '|' + _norm(e[r][4]).toLowerCase());
    }
  }

  const seen = new Map();
  return rows.map(r => {
    const out = { line: r.line, studentID: _norm(r.studentID), subject: _norm(r.subject), grade: _norm(r.grade),
                  semester: _norm(r.semester), remarks: _norm(r.remarks), instructor: _norm(r.instructor),
                  studentName: '', code: '', errors: [], warnings: [] };
    const st = students[out.studentID.toUpperCase()];
    if (!out.studentID) out.errors.push('Student ID is empty');
    else if (!st) out.errors.push('Student ID "' + out.studentID + '" is not in the system — add their profile first');
    else { out.studentID = st.id; out.studentName = st.name; }

    const c = out.subject ? cat.find(out.subject) : null;
    if (!out.subject) out.errors.push('Subject is empty');
    else if (!c) {
      const sug = cat.suggest(out.subject);
      out.errors.push('Subject "' + out.subject + '" is not in COURSE_CATALOG' + (sug ? ' — did you mean "' + sug + '"?' : ''));
    } else { out.subject = c.title; out.code = c.code; }

    const gv = _checkGradeValue_(out.grade);
    if (!gv.ok) out.errors.push(gv.msg); else out.grade = gv.value;

    if (!out.semester) out.errors.push('Semester is empty');
    else if (!/(19|20)\d{2}/.test(out.semester)) out.warnings.push('Semester has no year in it');

    if (!out.errors.length) {
      const k = out.studentID.toUpperCase() + '|' + out.subject.toLowerCase();
      const same = (existing.get(k) || []).filter(x => x.semester.toLowerCase() === out.semester.toLowerCase());
      const other = (existing.get(k) || []).filter(x => x.semester.toLowerCase() !== out.semester.toLowerCase());
      const dupKey = k + '|' + out.semester.toLowerCase();
      if (same.length) out.errors.push('Already has a grade (' + same[0].grade + ') for this subject in ' + out.semester);
      else if (seen.has(dupKey)) out.errors.push('Duplicate of line ' + seen.get(dupKey));
      else seen.set(dupKey, out.line);
      if (other.length) out.warnings.push('Also has ' + other.map(x => x.grade + ' (' + x.semester + ')').join(', ') + ' — retake?');
      if (enrolledNow.has(k)) out.warnings.push('Currently enrolled in this subject this semester');
    }
    out.ok = out.errors.length === 0;
    return out;
  });
}

// Appends validated rows one at a time (appendRow is safe even while
// instructors are saving grades at the same moment).
function _writeOldGrades_(validRows) {
  const sh = _gradesSheetWithSource_();
  validRows.forEach(v => {
    sh.appendRow([new Date(), v.semester, v.subject, v.studentID, v.studentName,
                  v.instructor || OLD_GRADE_INSTRUCTOR, v.grade, v.remarks, OLD_GRADE_SOURCE]);
  });
}

// Refreshes each student's Subjects Enrolled / Completed Subjects columns,
// stopping before the Apps Script time limit. Returns the ones not reached.
function _refreshSummaries_(sids, budgetMs) {
  const start = Date.now(), pending = [];
  sids.forEach(sid => {
    if (Date.now() - start > budgetMs) { pending.push(sid); return; }
    try { _rebuildStudentSubjects_(sid); } catch (e) { pending.push(sid); }
  });
  return pending;
}

function _ensureGrLookupRows_() {
  const ss = _ss(), lk = ss.getSheetByName('GR_LOOKUP');
  if (!lk) return;
  const need = ss.getSheetByName('Grades').getLastRow() + 200;
  if (lk.getMaxRows() < need) lk.insertRowsAfter(lk.getMaxRows(), need - lk.getMaxRows());
}

// ── Student lookup for the Old Grades tab ──
function adminOldGradesStudent(studentID, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const q = _norm(studentID).toUpperCase();
  if (!q) return { success: false, message: 'Enter a Student ID.' };
  const m = ss.getSheetByName('Master').getDataRange().getValues();
  let row = null;
  for (let r = 1; r < m.length; r++) if (_norm(m[r][1]).toUpperCase() === q) { row = m[r]; break; }
  if (!row) return { success: true, found: false, studentID: _norm(studentID) };
  const sid = _norm(row[1]);
  const { map: catalog } = _getCourseCatalogIndex_();
  const grades = [];
  const g = ss.getSheetByName('Grades').getDataRange().getValues();
  for (let r = 1; r < g.length; r++) {
    if (_norm(g[r][3]).toUpperCase() !== q || !_norm(g[r][2])) continue;
    const info = catalog.get(_norm(g[r][2]).toLowerCase()) || {};
    const source = _norm(g[r][8]);
    grades.push({ rowNum: r + 1, semester: _norm(g[r][1]), subject: _norm(g[r][2]), code: info.code || '',
                  program: info.program || '', grade: _norm(g[r][6]), remarks: _norm(g[r][7]),
                  instructor: _norm(g[r][5]), status: _gradeStatus_(g[r][6]),
                  source: source === OLD_GRADE_SOURCE ? 'Old record' : 'Instructor portal',
                  editable: source === OLD_GRADE_SOURCE });
  }
  grades.sort((a, b) => (_semesterSortKey_(a.semester) - _semesterSortKey_(b.semester)) || a.subject.localeCompare(b.subject));
  const rec = _getAcademicRecord_(sid);
  return { success: true, found: true, student: {
    studentID: sid, name: _buildStudentName_(row), campus: _norm(row[27]), classification: _norm(row[25]),
    current: rec.current.map(s => s.title), unitsEarned: rec.unitsEarned
  }, grades: grades };
}

// ── One grade from the form ──
function adminSaveOldGrade(item, token) {
  _requireAdmin_(token);
  const v = _validateOldGrades_([Object.assign({ line: 1 }, item || {})])[0];
  if (!v.ok) return { success: false, message: v.errors.join('; '), row: v };
  _writeOldGrades_([v]);
  _ensureGrLookupRows_();
  _refreshSummaries_([v.studentID], 60000);
  return { success: true, message: 'Saved ' + v.subject + ' — ' + v.grade + ' (' + v.semester + ') for ' + v.studentName + '.' +
           (v.warnings.length ? ' Note: ' + v.warnings.join('; ') : ''), row: v };
}

// ── Bulk paste: step 1, check only ──
function adminCheckOldGrades(text, token) {
  _requireAdmin_(token);
  const rows = _parseOldGradesText_(text);
  if (!rows.length) return { success: false, message: 'Nothing to check. Paste rows first.' };
  if (rows.length > OLD_GRADES_MAX_ROWS) {
    return { success: false, message: 'Please paste at most ' + OLD_GRADES_MAX_ROWS + ' rows at a time (you pasted ' + rows.length + ').' };
  }
  const res = _validateOldGrades_(rows);
  return { success: true, rows: res, ok: res.filter(r => r.ok).length, errors: res.filter(r => !r.ok).length,
           warnings: res.filter(r => r.ok && r.warnings.length).length };
}

// ── Bulk paste: step 2, save the rows that pass (re-checked here) ──
function adminImportOldGrades(text, token) {
  _requireAdmin_(token);
  const rows = _parseOldGradesText_(text);
  if (!rows.length) return { success: false, message: 'Nothing to save.' };
  if (rows.length > OLD_GRADES_MAX_ROWS) return { success: false, message: 'At most ' + OLD_GRADES_MAX_ROWS + ' rows at a time.' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Another save is in progress. Try again in a moment.' };
  try {
    const res = _validateOldGrades_(rows);
    const good = res.filter(r => r.ok);
    _writeOldGrades_(good);
    _ensureGrLookupRows_();
    const sids = Array.from(new Set(good.map(r => r.studentID)));
    const pending = _refreshSummaries_(sids, 200000);
    return { success: true, saved: good.length, skipped: res.length - good.length, rows: res, pendingSummaries: pending,
             message: 'Saved ' + good.length + ' old grade(s).' +
                      (res.length - good.length ? ' Not saved: ' + (res.length - good.length) + ' row(s) with errors.' : '') +
                      (pending.length ? ' ' + pending.length + ' student summary column(s) still need "Refresh summaries".' : '') };
  } finally { lock.releaseLock(); }
}

function adminRefreshStudentSummaries(sids, token) {
  _requireAdmin_(token);
  const pending = _refreshSummaries_((sids || []).map(_norm).filter(Boolean), 240000);
  return { success: true, pending: pending,
           message: pending.length ? pending.length + ' still pending — click again.' : 'All student summaries refreshed.' };
}

// Finds an admin-encoded row and checks it is still the one the admin saw.
function _findOldGradeRow_(sh, rowNum, studentID, subject) {
  const n = Number(rowNum);
  if (!(n >= 2) || n > sh.getLastRow()) return null;
  const v = sh.getRange(n, 1, 1, 9).getValues()[0];
  if (_norm(v[3]).toUpperCase() !== _norm(studentID).toUpperCase()) return null;
  if (_norm(v[2]).toLowerCase() !== _norm(subject).toLowerCase()) return null;
  if (_norm(v[8]) !== OLD_GRADE_SOURCE) return null;     // portal grades are never touched here
  return v;
}

function adminUpdateOldGrade(rowNum, studentID, subject, item, token) {
  _requireAdmin_(token);
  item = item || {};
  const sh = _gradesSheetWithSource_();
  const v = _findOldGradeRow_(sh, rowNum, studentID, subject);
  if (!v) return { success: false, message: 'That grade changed or is not an old record. Reload the student and try again.' };
  const gv = _checkGradeValue_(item.grade);
  if (!gv.ok) return { success: false, message: gv.msg };
  const sem = _norm(item.semester);
  if (!sem) return { success: false, message: 'Semester is empty.' };
  if (sem.toLowerCase() !== _norm(v[1]).toLowerCase()) {
    const g = sh.getDataRange().getValues();
    for (let r = 1; r < g.length; r++) {
      if (r + 1 === Number(rowNum)) continue;
      if (_norm(g[r][3]).toUpperCase() === _norm(studentID).toUpperCase() &&
          _norm(g[r][2]).toLowerCase() === _norm(subject).toLowerCase() &&
          _norm(g[r][1]).toLowerCase() === sem.toLowerCase() && _norm(g[r][6])) {
        return { success: false, message: 'This student already has a grade for this subject in ' + sem + '.' };
      }
    }
  }
  const remarks = _norm(item.remarks), instructor = _norm(item.instructor) || OLD_GRADE_INSTRUCTOR;
  sh.getRange(Number(rowNum), 1, 1, 9).setValues([[new Date(), sem, v[2], v[3], v[4], instructor, gv.value, remarks, OLD_GRADE_SOURCE]]);
  _logGradeHistory_([new Date(), sem, v[2], v[3], v[4], 'Admin (old record)', _norm(v[6]), gv.value, _norm(v[7]), remarks]);
  _refreshSummaries_([_norm(v[3])], 60000);
  return { success: true, message: 'Updated ' + v[2] + ' for ' + v[4] + '.' };
}

// Clears the row instead of deleting it, so no other row moves (GR_LOOKUP
// and anything else that points at Grades rows stays lined up).
function adminDeleteOldGrade(rowNum, studentID, subject, token) {
  _requireAdmin_(token);
  const sh = _gradesSheetWithSource_();
  const v = _findOldGradeRow_(sh, rowNum, studentID, subject);
  if (!v) return { success: false, message: 'That grade changed or is not an old record. Reload the student and try again.' };
  sh.getRange(Number(rowNum), 1, 1, 9).clearContent();
  _logGradeHistory_([new Date(), v[1], v[2], v[3], v[4], 'Admin (old record)', _norm(v[6]), '(deleted)', _norm(v[7]), '']);
  _refreshSummaries_([_norm(v[3])], 60000);
  return { success: true, message: 'Deleted ' + v[2] + ' (' + _norm(v[6]) + ') for ' + v[4] + '.' };
}

// ── Old students who are not in the system yet ──
function adminSuggestStudentID(token) {
  _requireAdmin_(token);
  return { success: true, id: generateStudentID() };
}

function adminCreateOldStudent(p, token) {
  _requireAdmin_(token);
  p = p || {};
  const id = _norm(p.studentID).toUpperCase();
  if (!id) return { success: false, message: 'Student ID is required (use their old ID, or click "Suggest ID").' };
  if (!/^[A-Z0-9][A-Z0-9-]{2,30}$/.test(id)) return { success: false, message: 'Student ID may only contain letters, numbers and dashes.' };
  if (!_norm(p.surname) || !_norm(p.firstName)) return { success: false, message: 'Surname and First Name are required.' };
  const campus = _norm(p.campus).toUpperCase();
  if (ADMIN_CAMPUSES.indexOf(campus) < 0) return { success: false, message: 'Choose a campus.' };
  const cls = OLD_STUDENT_CLASSES.indexOf(_norm(p.classification)) >= 0 ? _norm(p.classification) : 'Alumni';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Busy — try again in a moment.' };
  try {
    const sh = _ss().getSheetByName('Master');
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues() : [];
    if (ids.some(r => _norm(r[0]).toUpperCase() === id)) return { success: false, message: 'Student ID ' + id + ' already exists.' };
    const obj = {
      'Timestamp': new Date(), 'Student ID': id, 'Email': _norm(p.email).toLowerCase(),
      'Surname': _norm(p.surname).toUpperCase(), 'First Name': _norm(p.firstName).toUpperCase(),
      'Middle Name': _norm(p.middleName).toUpperCase(), 'Address': _norm(p.address), 'Mobile': _norm(p.mobile),
      'Date of Birth': _norm(p.dob), 'Sex': _norm(p.sex), 'Civil Status': _norm(p.civilStatus),
      'Campus': campus, 'Classification': cls, 'New Student': 'No',
      'Password': ''      // blank = cannot log in until an admin resets the password
    };
    const hi = _headerIndex_(sh);
    const out = new Array(hi.width).fill('');
    Object.keys(obj).forEach(k => { const i = hi.map[k.toLowerCase()]; if (i !== undefined) out[i] = obj[k]; });
    sh.appendRow(out);
  } finally { lock.releaseLock(); }
  _propagateStudent_(id);
  return { success: true, studentID: id,
           message: 'Added ' + _norm(p.surname).toUpperCase() + ', ' + _norm(p.firstName).toUpperCase() + ' (' + id + ').' };
}

// ── TOR lookup (GR_LOOKUP) ──
// One formula that (1) covers every Grades row, however many, and (2) lists
// passing grades first, so the TOR's VLOOKUP shows the passing grade when a
// subject was retaken. Same 10 columns, in the same order, as before.
const GR_LOOKUP_FORMULA =
  '=IFERROR(ARRAYFORMULA(LET(g,FILTER(Grades!A2:H,Grades!D2:D<>"",Grades!G2:G<>""),' +
  'sub,INDEX(g,,3),gr,INDEX(g,,7),n,IFERROR(VALUE(gr),-1),' +
  'pass,IF(n>5,n>=75,IF(n>=1,n<=3,REGEXMATCH(UPPER(TO_TEXT(gr)),"^(P|PASS|PASSED|S|CR)$"))),' +
  'cat,COURSE_CATALOG!A2:D,' +
  'code,IFERROR(VLOOKUP(sub,{INDEX(cat,,3),INDEX(cat,,2)},2,FALSE),""),' +
  'units,IFERROR(VLOOKUP(sub,{INDEX(cat,,3),INDEX(cat,,4)},2,FALSE),""),' +
  'prog,IFERROR(VLOOKUP(sub,{INDEX(cat,,3),INDEX(cat,,1)},2,FALSE),""),' +
  'ARRAY_CONSTRAIN(SORT({INDEX(g,,4)&"|"&sub,gr,INDEX(g,,8),INDEX(g,,2),INDEX(g,,1),INDEX(g,,6),INDEX(g,,5),code,units,prog,pass},11,FALSE,5,FALSE),ROWS(g),10))),"")';

function _gradedRowKeys_() {
  const g = _ss().getSheetByName('Grades').getDataRange().getValues();
  const keys = [];
  for (let r = 1; r < g.length; r++) {
    if (_norm(g[r][3]) && _norm(g[r][6])) keys.push((_norm(g[r][3]) + '|' + _norm(g[r][2])).toLowerCase());
  }
  return keys;
}

function adminTorLookupStatus(token) {
  _requireAdmin_(token);
  const lk = _ss().getSheetByName('GR_LOOKUP');
  if (!lk) return { success: false, message: 'GR_LOOKUP sheet not found.' };
  const upgraded = lk.getRange('A2').getFormula() === GR_LOOKUP_FORMULA;
  const keys = new Set(_gradedRowKeys_());
  const shown = new Set(lk.getRange(2, 1, Math.max(lk.getLastRow() - 1, 1), 1).getValues()
    .map(r => _norm(r[0]).toLowerCase()).filter(Boolean));
  let missing = 0;
  keys.forEach(k => { if (!shown.has(k)) missing++; });
  return { success: true, upgraded: upgraded, total: keys.size, missing: missing };
}

// Installs GR_LOOKUP_FORMULA, checks the result, and puts the old formulas
// back automatically if anything is wrong.
function adminUpgradeTorLookup(token) {
  _requireAdmin_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Busy — try again in a moment.' };
  try {
    const ss = _ss(), lk = ss.getSheetByName('GR_LOOKUP');
    if (!lk) return { success: false, message: 'GR_LOOKUP sheet not found.' };
    const rng = lk.getRange(2, 1, lk.getMaxRows() - 1, 10);
    const oldF = rng.getFormulas(), oldV = rng.getValues();
    const restore = () => {
      rng.clearContent();
      rng.setFormulas(oldF);
      oldF.forEach((row, i) => row.forEach((f, j) => {
        if (!f && oldV[i][j] !== '') lk.getRange(i + 2, j + 1).setValue(oldV[i][j]);
      }));
      SpreadsheetApp.flush();
    };
    rng.clearContent();
    _ensureGrLookupRows_();
    lk.getRange('A2').setFormula(GR_LOOKUP_FORMULA);
    SpreadsheetApp.flush();

    const keys = _gradedRowKeys_();
    const out = lk.getRange(2, 1, Math.max(lk.getLastRow() - 1, 1), 2).getValues();
    const shown = new Set(out.map(r => _norm(r[0]).toLowerCase()).filter(Boolean));
    const bad = out.some(r => /^#/.test(_norm(r[0])) || /^#/.test(_norm(r[1])));
    const missing = keys.filter(k => !shown.has(k)).length;
    if (bad || missing || (keys.length && !shown.size)) {
      restore();
      return { success: false, message: 'The new lookup did not check out (' + (bad ? 'formula error' : missing + ' grade(s) missing') +
                                         '). Nothing was changed — the old formulas were put back.' };
    }
    return { success: true, message: 'TOR lookup upgraded. It now covers all ' + keys.length +
                                      ' graded rows and shows the passing grade for retaken subjects.' };
  } finally { lock.releaseLock(); }
}
