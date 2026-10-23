// ========================================
// TWI BIBLE COLLEGE — ADMIN DASHBOARD BACKEND
// Admin.gs
// Requires: Code.gs (_ss, _norm, _getOrCreateSheet, _getCurrentSemesterValue_,
//           _getCourseCatalogIndex_, _resolveSubjectColumns_, _computeFeesByUnits_,
//           _buildStudentName_, _peso_, _safeFile_, formatTimeAsText, MASTER_HEADERS)
//           Academic.gs (getAcademicRecord, _rebuildStudentSubjects_, PH_ALLOWED_DEFAULT)
// ========================================

const BUILD_ADMIN = '2026-10-23c';   // version stamp — see getServerBuild() in Code.gs

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

function adminBuildInvoiceDraft(studentID, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const sem = _getCurrentSemesterValue_(_studentCampus_(studentID));
  const m = ss.getSheetByName('Master').getDataRange().getValues();
  let row = null;
  for (let r = 1; r < m.length; r++) {
    if (_norm(m[r][1]) === _norm(studentID)) { row = m[r]; break; }
  }
  if (!row) return { success:false, message:'Student not found.' };

  const campus = _norm(row[27]).toUpperCase();
  const isInt = (campus !== 'TWI-PHILIPPINES');

  // Current load from the academic record (never the AA string).
  const rec = _getAcademicRecord_(studentID);
  if (!rec.current.length) return { success:false, message:'No enrolled subjects for ' + sem + '.' };

  const subjects = rec.current.map(s => ({
    code: s.code || '', title: s.title, units: s.units, instructor: s.instructor
  }));
  const units = rec.unitsCurrent;
  const fees = _computeFeesByUnits_(units, isInt);

  return { success:true, draft: {
    studentID: _norm(studentID),
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
    const dr = adminBuildInvoiceDraft(studentID, token);
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

    const paid = Number(opts.amountPaid) || 0;
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

    return { success:true, invoiceNo: invNo, fileUrl: file.getUrl(),
      message: 'Invoice ' + invNo + ' created' + (opts.sendEmail !== false ? ' and emailed.' : '.') };
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
