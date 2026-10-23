// ============================================================
// TWI BIBLE COLLEGE — ACADEMIC RECORD ENGINE
// New script file: Academic.gs
// Requires: Code.gs (for _ss, _norm, _getCurrentSemesterValue_,
//           _getCourseCatalogIndex_, _buildStudentName_)
//           Admin.gs (for _headerIndex_, _propagateStudent_, ADMIN_CAMPUSES)
// ============================================================

const BUILD_ACADEMIC = '2026-10-23c';   // version stamp — see getServerBuild() in Code.gs

// Master column numbers (1-based)
const COL_SUBJECTS_ENROLLED = 27;  // AA  -> CURRENT semester only
const COL_CAMPUS            = 28;  // AB
const COL_ALLOWED_PROGRAMS  = 31;  // AE
const COL_COMPLETED         = 32;  // AF  -> NEW: passed subjects w/ grades

const PH_ALLOWED_DEFAULT = '1st Year CCM, 2nd Year CCM, BCM, CCM Evening Class';

// ── GRADE INTERPRETATION ─────────────────────────────────────
// Two scales are in use in the Grades sheet:
//   Point scale:      1.00 (highest) .. 3.00 (lowest pass). 4.00 = conditional, 5.00 = fail.
//   Percentage scale: 75 and above passes (e.g. 89, 91); below 75 fails.
// A number above 5 is read as a percentage.
const PASSING_PERCENT = 75;
function _gradeStatus_(gradeRaw) {
  const g = _norm(gradeRaw);
  if (!g) return 'PENDING';
  const n = Number(g);
  if (!isFinite(n)) {
    // tolerate letter marks if you ever use them
    const u = g.toUpperCase();
    if (['P','PASS','PASSED','S','CR'].indexOf(u) >= 0) return 'PASSED';
    if (['F','FAIL','FAILED','INC','DRP','D'].indexOf(u) >= 0) return 'FAILED';
    return 'PENDING';
  }
  if (n > 5) return n >= PASSING_PERCENT ? 'PASSED' : 'FAILED';   // percentage scale
  if (n >= 1 && n <= 3) return 'PASSED';
  return 'FAILED';                    // 4.00, 5.00, 0, anything else numeric
}

function _isPassing_(gradeRaw) { return _gradeStatus_(gradeRaw) === 'PASSED'; }

// ── CORE: full academic record for one student ───────────────
// Returns { current:[], passed:[], failed:[], dropped:[], all:{} }
// Internal. No session check — callers that expose it to the browser must
// guard first. The public entry point is getAcademicRecord() below.
function _getAcademicRecord_(studentID) {
  const ss  = _ss();
  const sid = _norm(studentID);
  // "Current" depends on the student's campus: PH and INT run separate terms.
  const sem = _getCurrentSemesterValue_(_studentCampus_(sid));
  const { map: catalog } = _getCourseCatalogIndex_();

  // 1. Every grade ever recorded for this student (latest row per subject wins,
  //    so a retake overwrites the earlier failing mark).
  const gradeBySubject = new Map();   // lowercased title -> {grade, sem, status, instructor, remarks}
  const gsh = ss.getSheetByName('Grades');
  if (gsh) {
    const g = gsh.getDataRange().getValues();
    for (let r = 1; r < g.length; r++) {
      if (_norm(g[r][3]) !== sid) continue;
      const title = _norm(g[r][2]); if (!title) continue;
      const grade = _norm(g[r][6]);
      const st    = _gradeStatus_(grade);
      if (st === 'PENDING') continue;          // blank grade row = not yet graded
      const prev = gradeBySubject.get(title.toLowerCase());
      // A later PASSED always supersedes an earlier FAILED (retake).
      if (prev && prev.status === 'PASSED' && st !== 'PASSED') continue;
      gradeBySubject.set(title.toLowerCase(), {
        title: title, grade: grade, status: st,
        semester: _norm(g[r][1]), instructor: _norm(g[r][5]), remarks: _norm(g[r][7])
      });
    }
  }

  // 2. Every enrollment row for this student.
  const enrollments = [];
  const esh = ss.getSheetByName('Enrollments');
  if (esh) {
    const e = esh.getDataRange().getValues();
    for (let r = 1; r < e.length; r++) {
      if (_norm(e[r][2]) !== sid) continue;
      const title = _norm(e[r][4]); if (!title) continue;
      enrollments.push({
        row: r + 1, semester: _norm(e[r][1]), title: title,
        instructor: _norm(e[r][5]), yearLevel: _norm(e[r][6]),
        status: _norm(e[r][7]).toUpperCase()
      });
    }
  }

  const current = [], passed = [], failed = [], dropped = [];
  const seenCurrent = new Set();

  // CURRENT = enrolled this semester, not dropped, no grade yet.
  enrollments.forEach(en => {
    if (String(en.semester).toLowerCase() !== String(sem).toLowerCase()) return;
    if (en.status === 'DROPPED') { dropped.push(en); return; }
    if (en.status !== 'ENROLLED') return;
    const key = en.title.toLowerCase();
    if (gradeBySubject.has(key)) return;       // already graded -> not "current"
    if (seenCurrent.has(key)) return;
    seenCurrent.add(key);
    const info = catalog.get(key) || {};
    current.push({
      title: en.title, code: info.code || '', units: Number(info.units) || 3,
      instructor: en.instructor, semester: en.semester, yearLevel: en.yearLevel
    });
  });

  // PASSED / FAILED from the grade map.
  gradeBySubject.forEach((v, key) => {
    const info = catalog.get(key) || {};
    const rec = {
      title: v.title, code: info.code || '', units: Number(info.units) || 3,
      grade: v.grade, semester: v.semester, instructor: v.instructor, remarks: v.remarks
    };
    if (v.status === 'PASSED') passed.push(rec); else failed.push(rec);
  });

  const unitsEarned = passed.reduce((a, s) => a + (Number(s.units) || 0), 0);
  const unitsCurrent = current.reduce((a, s) => a + (Number(s.units) || 0), 0);

  return {
    success: true, studentID: sid, semester: sem,
    current: current, passed: passed, failed: failed, dropped: dropped,
    unitsEarned: unitsEarned, unitsCurrent: unitsCurrent,
    passedSet: passed.map(s => s.title.toLowerCase()),
    failedSet: failed.map(s => s.title.toLowerCase())
  };
}

// ── DERIVED COLUMNS: rebuild AA + AF from Enrollments/Grades ──
// This is the whole point: AA and AF are never authored by hand.
function _rebuildStudentSubjects_(studentID) {
  const ss = _ss();
  const sh = ss.getSheetByName('Master');
  const d  = sh.getDataRange().getValues();
  const rec = _getAcademicRecord_(studentID);

  const currentStr   = rec.current.map(s => s.title).join(', ');
  const completedStr = rec.passed
    .sort((a,b) => a.title.localeCompare(b.title))
    .map(s => s.title + ' (' + s.grade + ')').join(', ');

  for (let r = 1; r < d.length; r++) {
    if (_norm(d[r][1]) !== _norm(studentID)) continue;
    sh.getRange(r + 1, COL_SUBJECTS_ENROLLED).setValue(currentStr);
    sh.getRange(r + 1, COL_COMPLETED).setValue(completedStr);
    break;
  }
  return { current: currentStr, completed: completedStr };
}

function rebuildStudentSubjects(studentID, token) {
  _requireStudent_(token, studentID);
  const out = _rebuildStudentSubjects_(studentID);
  _propagateStudent_(studentID);        // push to campus + class sheets
  return { success: true, message: 'Rebuilt.', current: out.current, completed: out.completed };
}

// ── ONE-TIME MIGRATION ───────────────────────────────────────
// Adds the AF header, backfills AE, and rebuilds AA/AF for every student
// from the Enrollments + Grades sheets (the real source of truth).
function migrateAcademicColumns() {
  const ss = _ss();
  const sh = ss.getSheetByName('Master');

  // Header for the new AF column
  if (_norm(sh.getRange(1, COL_COMPLETED).getValue()) === '') {
    sh.getRange(1, COL_COMPLETED).setValue('Completed Subjects');
  }
  // Make sure every campus/class sheet has both headers too
  const targets = ADMIN_CAMPUSES.concat(['1st Year CCM','2nd Year CCM','CCM Evening Class','Evening Class','BCM']);
  targets.forEach(n => {
    const s = ss.getSheetByName(n);
    if (s) { _ensureHeader_(s, 'Allowed Programs'); _ensureHeader_(s, 'Completed Subjects'); }
  });

  const last = sh.getLastRow();
  if (last < 2) return { success: true, message: 'No students.' };

  // Backfill AE (Allowed Programs) by campus
  const campusVals  = sh.getRange(2, COL_CAMPUS, last - 1, 1).getValues();
  const allowedVals = sh.getRange(2, COL_ALLOWED_PROGRAMS, last - 1, 1).getValues();
  for (let i = 0; i < campusVals.length; i++) {
    const c = _norm(campusVals[i][0]).toUpperCase();
    if (!c) continue;
    allowedVals[i][0] = (c === 'TWI-PHILIPPINES') ? PH_ALLOWED_DEFAULT : 'INT';
  }
  sh.getRange(2, COL_ALLOWED_PROGRAMS, last - 1, 1).setValues(allowedVals);

  // Rebuild AA + AF in ONE batched write (not row-by-row -- avoids timeouts)
  const ids = sh.getRange(2, 2, last - 1, 1).getValues();
  const cur = [], comp = [];
  for (let i = 0; i < ids.length; i++) {
    const sid = _norm(ids[i][0]);
    if (!sid) { cur.push(['']); comp.push(['']); continue; }
    const rec = _getAcademicRecord_(sid);
    cur.push([ rec.current.map(s => s.title).join(', ') ]);
    comp.push([ rec.passed.sort((a,b)=>a.title.localeCompare(b.title))
                  .map(s => s.title + ' (' + s.grade + ')').join(', ') ]);
  }
  sh.getRange(2, COL_SUBJECTS_ENROLLED, last - 1, 1).setValues(cur);
  sh.getRange(2, COL_COMPLETED,         last - 1, 1).setValues(comp);

  return { success: true, message: 'Migrated ' + ids.length + ' students. AA = current semester, AF = completed.' };
}

// ── RE-ENROLLMENT GUARD ──────────────────────────────────────
// Returns null if OK, or a rejection message.
function _blockIfAlreadyTaken_(studentID, subjectName) {
  const rec = _getAcademicRecord_(studentID);
  const key = _norm(subjectName).toLowerCase();
  if (rec.passedSet.indexOf(key) >= 0) {
    const p = rec.passed.filter(s => s.title.toLowerCase() === key)[0];
    return 'You already passed "' + subjectName + '" with a grade of ' + p.grade +
           ' (' + p.semester + '). You cannot enroll in it again.';
  }
  if (rec.current.some(s => s.title.toLowerCase() === key)) {
    return 'You are already enrolled in "' + subjectName + '" this semester.';
  }
  return null;   // failed subjects fall through -> retake allowed
}

// ── SUBJECT ROSTER: all students in one subject, this semester ──
function adminGetSubjectRoster(subjectName, campusFilter, semester, token) {
  _requireAdmin_(token);
  const ss = _ss();
  const match = _semMatcher_(semester);
  const sub = _norm(subjectName).toLowerCase();
  const cf  = _norm(campusFilter).toUpperCase();

  // Master lookup
  const m = ss.getSheetByName('Master').getDataRange().getValues();
  const info = {};
  for (let r = 1; r < m.length; r++) {
    const sid = _norm(m[r][1]); if (!sid) continue;
    info[sid] = {
      email: _norm(m[r][2]), surname: _norm(m[r][3]), name: _norm(m[r][4]),
      middleName: _norm(m[r][5]), contact: _norm(m[r][7]), facebook: _norm(m[r][15]),
      classification: _norm(m[r][25]), campus: _norm(m[r][27]).toUpperCase()
    };
  }

  // Already-graded students shouldn't appear as active in the subject
  const graded = new Set();
  const gsh = ss.getSheetByName('Grades');
  if (gsh) {
    const g = gsh.getDataRange().getValues();
    for (let r = 1; r < g.length; r++) {
      if (_norm(g[r][2]).toLowerCase() !== sub) continue;
      if (_norm(g[r][6])) graded.add(_norm(g[r][3]));
    }
  }

  const rows = [];
  const e = ss.getSheetByName('Enrollments').getDataRange().getValues();
  for (let r = 1; r < e.length; r++) {
    if (_norm(e[r][4]).toLowerCase() !== sub) continue;
    if (_norm(e[r][7]).toUpperCase() !== 'ENROLLED') continue;
    const sid = _norm(e[r][2]);
    const s = info[sid]; if (!s) continue;
    if (!match(e[r][1], s.campus)) continue;
    if (cf && cf !== 'ALL' && s.campus !== cf) continue;
    if (rows.some(x => x.studentID === sid)) continue;
    rows.push({
      studentID: sid, subjects: _norm(e[r][4]), email: s.email, surname: s.surname,
      name: s.name, middleName: s.middleName, contact: s.contact, facebook: s.facebook,
      campus: s.campus, classification: s.classification,
      instructor: _norm(e[r][5]), graded: graded.has(sid)
    });
  }
  rows.sort((a, b) => a.surname.localeCompare(b.surname));
  return { success: true, subject: _norm(subjectName), semester: _semViewLabel_(semester), rows: rows, count: rows.length };
}

// Subjects for the chosen semester (default: current), from the offerings
// sheets plus anything with enrollments that semester, with a headcount.
// Feeds the Enrolled-tab filter.
function adminGetActiveSubjects(semester, token) {
  _requireAdmin_(token);
  const ss  = _ss();
  const match = _semMatcher_(semester);
  const campusOf = _campusBySid_(ss.getSheetByName('Master').getDataRange().getValues());

  // headcount per subject from Enrollments
  const counts = {};
  const e = ss.getSheetByName('Enrollments').getDataRange().getValues();
  for (let r = 1; r < e.length; r++) {
    if (!match(e[r][1], campusOf[_norm(e[r][2])])) continue;
    if (_norm(e[r][7]).toUpperCase() !== 'ENROLLED') continue;
    const t = _norm(e[r][4]); if (!t) continue;
    const k = t.toLowerCase();
    if (!counts[k]) counts[k] = { title: t, n: new Set() };
    counts[k].n.add(_norm(e[r][2]));
  }

  const out = [];
  const seen = new Set();
  ['Subjects', 'INT Subjects'].forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName); if (!sh) return;
    const cols = _resolveSubjectColumns_(sh);
    const d = sh.getDataRange().getValues();
    for (let r = 1; r < d.length; r++) {
      const title = _norm(d[r][cols.subject]); if (!title) continue;
      const rowSem = _norm(d[r][cols.semester]);
      // offering counts as active if its semester matches, or is blank
      if (rowSem && !match(rowSem, sheetName === 'INT Subjects' ? 'TWI-QATAR' : 'TWI-PHILIPPINES')) continue;
      const key = title.toLowerCase() + '|' + _norm(d[r][cols.campus]).toUpperCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({
        subject: title,
        instructor: _norm(d[r][cols.instructor]),
        campus: _norm(d[r][cols.campus]).toUpperCase(),
        program: _norm(d[r][cols.program]),
        day: _norm(d[r][cols.day]),
        source: sheetName,
        enrolled: counts[title.toLowerCase()] ? counts[title.toLowerCase()].n.size : 0
      });
    }
  });
  // Past semesters: the offerings sheets have usually moved on, so list every
  // subject that actually had enrollments too.
  const listed = new Set(out.map(o => o.subject.toLowerCase()));
  Object.keys(counts).forEach(k => {
    if (listed.has(k)) return;
    out.push({ subject: counts[k].title, instructor: '', campus: '', program: '', day: '',
               source: 'Enrollments', enrolled: counts[k].n.size });
  });

  out.sort((a, b) => b.enrolled - a.enrolled || a.subject.localeCompare(b.subject));
  return { success: true, semester: _semViewLabel_(semester), subjects: out, count: out.length };
}

// ── ADMIN: academic record view for the dashboard ─────────────
function adminGetAcademicRecord(studentID, token) {
  _requireAdmin_(token);
  const rec = _getAcademicRecord_(studentID);
  const ss  = _ss();
  const m   = ss.getSheetByName('Master').getDataRange().getValues();
  let who = null;
  for (let r = 1; r < m.length; r++) {
    if (_norm(m[r][1]) === _norm(studentID)) {
      who = {
        studentID: _norm(m[r][1]), name: _buildStudentName_(m[r]),
        campus: _norm(m[r][27]).toUpperCase(), classification: _norm(m[r][25])
      };
      break;
    }
  }
  if (!who) return { success: false, message: 'Student not found.' };
  rec.student = who;
  return rec;
}

// Admin can drop a student from a subject (marks Enrollments row DROPPED, rebuilds AA)
function adminDropSubject(studentID, subjectName, token) {
  _requireAdmin_(token);
  try {
    const ss  = _ss();
    const sem = _getCurrentSemesterValue_(_studentCampus_(studentID));
    const sh  = ss.getSheetByName('Enrollments');
    const d   = sh.getDataRange().getValues();
    let n = 0;
    for (let r = 1; r < d.length; r++) {
      if (_norm(d[r][2]) !== _norm(studentID)) continue;
      if (String(_norm(d[r][1])).toLowerCase() !== String(sem).toLowerCase()) continue;
      if (_norm(d[r][4]).toLowerCase() !== _norm(subjectName).toLowerCase()) continue;
      sh.getRange(r + 1, 8).setValue('DROPPED');
      n++;
    }
    if (!n) return { success: false, message: 'No active enrollment found for that subject.' };
    _rebuildStudentSubjects_(studentID);
    _propagateStudent_(studentID);
    return { success: true, message: 'Dropped "' + subjectName + '". Current subjects updated.' };
  } catch (e) { return { success: false, message: 'Drop error: ' + e.message }; }
}
// ============================================================
// PUBLIC ENTRY POINT — a student may read only their own record.
// ============================================================
function getAcademicRecord(studentID, token) {
  _requireStudent_(token, studentID);
  return _getAcademicRecord_(studentID);
}
