// ========================================
// THY WORD INTL BIBLE COLLEGE BATAAN
// Registration + Student Portal + Instructor Portal + Admin
// Google Apps Script Backend — Code.gs
// ========================================

const SPREADSHEET_ID = '1ISbv7Hso14xeupMog3OdwQS5oMDbLRPE7IKBcp2J_nI';

// Version stamp. Each .gs file has one (BUILD_CODE / BUILD_AUTH / BUILD_ADMIN /
// BUILD_ACADEMIC). The Admin Dashboard login screen compares them with what it
// expects and names any file that was not pasted into Apps Script.
const BUILD_CODE = '2026-10-23a';

// Public on purpose: returns only version stamps, no data.
function getServerBuild() {
  // typeof is safe on a name that does not exist (an old file without a stamp).
  return {
    code:     BUILD_CODE,
    auth:     typeof BUILD_AUTH     !== 'undefined' ? BUILD_AUTH     : '',
    admin:    typeof BUILD_ADMIN    !== 'undefined' ? BUILD_ADMIN    : '',
    academic: typeof BUILD_ACADEMIC !== 'undefined' ? BUILD_ACADEMIC : ''
  };
}

// =============================
// ROUTING (ONLY ONE doGet)
// =============================
function doGet(e) {
  e = e || {};
  const page = (e.parameter && e.parameter.page) || 'registration';

  if (page === 'registration') {
    return HtmlService.createHtmlOutputFromFile('Registration')
      .setTitle('Thy Word Intl Bible College Registration')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'portal') {
    return HtmlService.createHtmlOutputFromFile('StudentPortal')
      .setTitle('Student Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'instructor') {
    return HtmlService.createHtmlOutputFromFile('Instructor')
      .setTitle('Instructor Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('AdminDashboard')
      .setTitle('Admin Dashboard - TWI Bible College')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutput("Page not found");
}

// =============================
// SETUP / SHEET HELPERS
// =============================
function _ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function _norm(v) {
  return String(v || '').trim();
}

function _getOrCreateSheet(name, headers) {
  const ss = _ss();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && headers.length) {
    const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const empty = firstRow.every(v => String(v || '').trim() === '');
    if (empty) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sh;
}

function setupInstructorModule() {
  _getOrCreateSheet('Instructors', ['Username', 'Password', 'Full Name', 'Email', 'Subject Handle 1']);
  _getOrCreateSheet('Settings', ['Key', 'Value']);
  _getOrCreateSheet('Enrollments', ['Timestamp','Semester','Student ID','Student Name','Subject','Instructor','Year Level','Status']);
  _getOrCreateSheet('Grades', ['Timestamp','Semester','Subject','Student ID','Student Name','Instructor','Grade','Remarks']);
  const settings = _ss().getSheetByName('Settings');
  const data = settings.getDataRange().getValues();
  const has = data.some((r, i) => i > 0 && String(r[0]).trim() === 'Current Semester');
  if (!has) settings.appendRow(['Current Semester', '']);
}

// =============================
// PROGRAM CONSTANTS / HELPERS
// =============================
const PH_PROGRAMS = ['1st Year CCM', '2nd Year CCM', 'BCM', 'CCM Evening Class'];

function _getProgramCaps_(program) {
  const p = _norm(program).toLowerCase();
  if (p === 'ccm evening class') return { maxSubjects: 2, maxUnits: 6 };
  return { maxSubjects: 4, maxUnits: 12 };
}

// International campuses (Qatar / Canada / Europe): one load cap for all programs.
const INT_CAPS = { maxSubjects: 2, maxUnits: 6 };

function _programFamily_(program) {
  const p = _norm(program).toLowerCase();
  if (p === 'ccm evening class') return 'CCM Evening Class';
  if (p === 'bcm') return 'BCM';
  if (p === '1st year ccm' || p === '2nd year ccm' || p === 'ccm') return 'CCM';
  return _norm(program);
}

function _parseAllowedPrograms_(allowedCell, fallbackClassification) {
  const raw = _norm(allowedCell);
  let list = [];
  if (raw) {
    list = raw.split(',').map(s => _norm(s)).filter(Boolean);
  }
  if (list.length === 0 && _norm(fallbackClassification)) {
    list = [_norm(fallbackClassification)];
  }
  const seen = new Set();
  const out = [];
  list.forEach(p => {
    const k = p.toLowerCase();
    if (!seen.has(k)) { seen.add(k); out.push(p); }
  });
  return out;
}

// =============================
// MASTER SHEET HEADERS
// Col 27 (AA) = Subjects Enrolled  -> CURRENT semester only (derived)
// Col 28 (AB) = Campus
// Col 31 (AE) = Allowed Programs
// Col 32 (AF) = Completed Subjects -> passed subjects w/ grades (derived)
// =============================
const MASTER_HEADERS = [
  'Timestamp','Student ID','Email','Surname','First Name','Middle Name',
  'Address','Mobile','Tel','Date of Birth','Sex','Civil Status','Spouse',
  'Emergency Contact Person','Emergency Contact Number','Facebook','Are You From AG',
  'Church Name','Church Address','Pastor Name','Ministry in Church',
  'Religious Affiliation','Recommendation','School Last Attended',
  'New Student','Classification','Subjects Enrolled','Campus','Password','Profile Picture URL',
  'Allowed Programs','Completed Subjects'
];

// =============================
// CAMPUS SHEET AUTO-SYNC (legacy — superseded by _propagateStudent_ in Admin.gs)
// =============================
function _ensureCampusSheets_() {
  const campuses = ['TWI-PHILIPPINES', 'TWI-QATAR', 'TWI-CANADA', 'TWI-EUROPE'];
  campuses.forEach(c => _getOrCreateSheet(c, MASTER_HEADERS));
}

function syncCampusSheets(token) {
  // Delegates to the header-mapped sync so values can never land in the wrong column.
  // adminSyncAll enforces _requireAdmin_, so the token must be passed through —
  // otherwise this becomes an unauthenticated back door to the same work.
  return adminSyncAll(token);
}

// =============================
// CURRENT SEMESTER
// =============================
function _getSemesterColumnIndex_(subjectsSheet) {
  const lastCol = subjectsSheet.getLastColumn();
  if (lastCol < 1) return 6;
  const headers = subjectsSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => _norm(h).toLowerCase());
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].includes('semester')) return i + 1;
  }
  return 6;
}

const INT_CAMPUS_LIST = ['TWI-QATAR', 'TWI-CANADA', 'TWI-EUROPE'];
function _isIntCampus_(campus) {
  return INT_CAMPUS_LIST.indexOf(_norm(campus).toUpperCase()) >= 0;
}

// Campus (Master column AB) for one student, upper-cased; '' if not found.
function _studentCampus_(studentID) {
  const sid = _norm(studentID);
  const m = _ss().getSheetByName('Master').getDataRange().getValues();
  for (let r = 1; r < m.length; r++) {
    if (_norm(m[r][1]) === sid) return _norm(m[r][27]).toUpperCase();
  }
  return '';
}

// Current semester. The Philippines and the international campuses run
// separate terms:
//   Settings "Current Semester"      -> TWI-PHILIPPINES (and the default)
//   Settings "Current Semester INT"  -> TWI-QATAR / CANADA / EUROPE
// With no campus argument, the Philippines semester is returned.
// If a Settings row is blank, it falls back to the most common value in the
// Semester column of 'Subjects' (PH) or 'INT Subjects' (INT).
function _getCurrentSemesterValue_(campus) {
  const isInt = _isIntCampus_(campus);
  const ss = _ss();
  const settings = ss.getSheetByName('Settings');
  if (settings) {
    const wanted = isInt ? 'current semester int' : 'current semester';
    const setData = settings.getDataRange().getValues();
    for (let i = 1; i < setData.length; i++) {
      if (_norm(setData[i][0]).toLowerCase() === wanted) {
        const overrideVal = _norm(setData[i][1]);
        if (overrideVal) return overrideVal;
      }
    }
  }
  const subj = ss.getSheetByName(isInt ? 'INT Subjects' : 'Subjects');
  if (!subj) return '';
  const lastRow = subj.getLastRow();
  if (lastRow < 2) return '';
  const semCol = _getSemesterColumnIndex_(subj);
  const semValues = subj.getRange(2, semCol, lastRow - 1, 1).getValues().flat().map(_norm).filter(Boolean);
  if (semValues.length === 0) return '';
  const uniq = Array.from(new Set(semValues));
  if (uniq.length === 1) return uniq[0];
  const freq = new Map();
  semValues.forEach(s => freq.set(s, (freq.get(s) || 0) + 1));
  let best = ''; let bestCount = -1;
  uniq.forEach(s => {
    const c = freq.get(s) || 0;
    if (c > bestCount) { best = s; bestCount = c; }
    else if (c === bestCount && s.localeCompare(best) > 0) best = s;
  });
  return best;
}

function getCurrentSemester() {
  return { currentSemester: _getCurrentSemesterValue_(), currentSemesterInt: _getCurrentSemesterValue_('TWI-QATAR') };
}

// =============================
// GENERATE STUDENT ID
// =============================
function generateStudentID() {
  const ss = _ss();
  const masterSheet = ss.getSheetByName('Master');
  const lastRow = masterSheet.getLastRow();
  let newID;
  if (lastRow <= 1) {
    newID = 'TWIBC-2026-0001';
  } else {
    const lastID = masterSheet.getRange(lastRow, 2).getValue();
    const lastNumber = parseInt(String(lastID).split('-')[2], 10);
    const newNumber = String(lastNumber + 1).padStart(4, '0');
    newID = 'TWIBC-2026-' + newNumber;
  }
  return newID;
}

// =============================
// SUBMIT REGISTRATION
// =============================
function submitRegistration(formData) {
  try {
    _ensureCampusSheets_();
    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const studentID = generateStudentID();
    const timestamp = new Date();

    // AE (Allowed Programs) auto-fills by campus.
    const campusU = _norm(formData.campus).toUpperCase();
    const allowedPrograms = _norm(formData.allowedPrograms) ||
      (campusU === 'TWI-PHILIPPINES' ? PH_ALLOWED_DEFAULT : 'INT');

    const rowData = [
      timestamp, studentID, formData.email,
      formData.surname, formData.firstName, formData.middleName,
      formData.address, formData.mobile, formData.tel,
      formData.dateOfBirth, formData.sex, formData.civilStatus,
      formData.spouse, formData.emergencyContactPerson, formData.emergencyContactNumber,
      formData.facebook, formData.areYouFromAG, formData.churchName,
      formData.churchAddress, formData.pastorName, formData.ministryInChurch,
      formData.religiousAffiliation, formData.recommendation, formData.schoolLastAttended,
      formData.newStudent, formData.classification, '',   // AA starts empty (derived)
      formData.campus, _hashPassword_(formData.password), '',
      allowedPrograms, ''                                  // AE, AF
    ];

    masterSheet.appendRow(rowData);

    // Header-mapped propagation to campus + class sheets
    _propagateStudent_(studentID);

    // Evidence of consent. RA 10173 requires consent to be "evidenced by
    // written, electronic or recorded means" — a tick box that leaves no
    // record proves nothing later.
    _recordConsent_(studentID, formData);

    sendWelcomeEmail(String(formData.email || '').toLowerCase().trim(), studentID, formData.firstName);
    return { success: true, studentID, message: 'Registration successful! Check your email for login credentials.' };
  } catch (error) {
    return { success: false, message: 'Registration failed: ' + error.message };
  }
}

// The password is deliberately NOT included. The student chose it during
// registration and we no longer hold a readable copy of it.
function sendWelcomeEmail(email, studentID, firstName) {
  const subject = 'Welcome to Thy Word Intl Bible College Bataan';
  const body =
    'Dear ' + firstName + ',\n\n' +
    'Welcome to Thy Word Intl Bible College Bataan!\n\n' +
    'Your registration has been successfully completed. Your Student Portal username is:\n\n' +
    'Student ID (Username): ' + studentID + '\n\n' +
    'Log in using the password you chose on the registration form. If you forget it, ' +
    'use "Forgot Password" on the portal and we will email you a temporary one.\n\n' +
    'Please keep your login credentials secure.\n\n' +
    'Kindly send your Pastor Recommendation Letter to the following email address: twibiblecollege@gmail.com\n\n' +
    'God bless you!\n\n' +
    'Thy Word Intl Bible College Bataan Administration';
  MailApp.sendEmail(email, subject, body);
}

// =============================
// LOGIN
// =============================
function checkLogin(studentID, password) {
  try {
    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const data = masterSheet.getDataRange().getValues();
    const trimmedStudentID = String(studentID).trim();
    for (let i = 1; i < data.length; i++) {
      const sheetStudentID = String(data[i][1]).trim();
      if (sheetStudentID !== trimmedStudentID) continue;

      const check = _verifyPassword_(password, data[i][28]);
      if (!check.ok) break;

      // First successful login against a legacy plaintext row rewrites it as
      // a hash. The student notices nothing and is never locked out.
      if (check.needsUpgrade) {
        _upgradeStoredPassword_(masterSheet, i + 1, MASTER_PW_COL, password);
      }
      return {
        success: true,
        studentID: trimmedStudentID,
        token: _issueSession_('student', trimmedStudentID, 'STUDENT')
      };
    }
    return { success: false, message: 'Invalid Student ID or Password' };
  } catch (error) {
    return { success: false, message: 'Login error: ' + error.message };
  }
}

function sendForgotPasswordEmail(studentID, email) {
  try {
    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const data = masterSheet.getDataRange().getValues();
    const trimmedStudentID = String(studentID).trim();
    const trimmedEmail = String(email).trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const sheetStudentID = String(data[i][1]).trim();
      const sheetEmail = String(data[i][2]).trim().toLowerCase();
      const firstName = data[i][4] || '';
      if (sheetStudentID === trimmedStudentID && sheetEmail === trimmedEmail) {
        // Passwords are hashed, so the old one cannot be recovered or emailed.
        // Issue a single-use temporary password instead and store its hash.
        const tempPassword = _generateTempPassword_();
        masterSheet.getRange(i + 1, MASTER_PW_COL).setValue(_hashPassword_(tempPassword));
        _propagateStudent_(sheetStudentID);

        const subject = 'Password Reset - Thy Word Intl Bible College Bataan';
        const body =
          'Dear ' + firstName + ',\n\n' +
          'A password reset was requested for your Student Portal account.\n\n' +
          'Student ID (Username): ' + sheetStudentID + '\n' +
          'Temporary Password: ' + tempPassword + '\n\n' +
          'Please log in with this temporary password and change it immediately ' +
          'from the Profile tab of your portal.\n\n' +
          'If you did not request this reset, contact the college office at once — ' +
          'your previous password no longer works.\n\n' +
          'God bless you!\n\n' +
          'Thy Word Intl Bible College Bataan Administration';
        MailApp.sendEmail(sheetEmail, subject, body);
        return { success: true, message: 'A temporary password has been sent to your email address.' };
      }
    }
    return { success: false, message: 'Student ID and Email do not match our records.' };
  } catch (error) {
    return { success: false, message: 'Error sending password: ' + error.message };
  }
}

// Readable but unguessable: 12 chars from an alphabet with no 0/O/1/l/I.
function _generateTempPassword_() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let out = '';
  for (let i = 0; i < 12; i++) {
    out += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  }
  return out;
}

// =============================
// GET STUDENT DATA
// =============================
function getStudentData(studentId, token) {
  _requireStudent_(token, studentId);
  const ss = _ss();
  const sheet = ss.getSheetByName('Master');
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange('B2:AF' + lastRow).getValues();   // extended to AF
  const headers = [
    'Student ID','Email','Surname','First Name','Middle Name',
    'Address','Mobile','Tel','Date of Birth','Sex','Civil Status','Spouse',
    'Emergency Contact Person','Emergency Contact Number','Facebook','Are You From AG',
    'Church Name','Church Address','Pastor Name','Ministry in Church',
    'Religious Affiliation','Recommendation','School Last Attended',
    'New Student','Classification','Subjects Enrolled','Campus',
    'Password','Profile Picture URL','Allowed Programs','Completed Subjects'
  ];
  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    if (id === String(studentId).trim()) {
      let studentData = {};
      headers.forEach((h, j) => studentData[h] = (data[i][j] !== undefined ? data[i][j] : ''));
      studentData['Allowed Programs List'] =
        _parseAllowedPrograms_(studentData['Allowed Programs'], studentData['Classification']);
      return JSON.stringify(studentData);
    }
  }
  return null;
}

// =============================
// UPDATE STUDENT PROFILE (student-facing)
// =============================
function updateStudentProfile(studentID, updates, token) {
  try {
    _requireStudent_(token, studentID);
    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const data = masterSheet.getDataRange().getValues();

    if (updates.mobile) {
      const mob = String(updates.mobile).replace(/\D/g, '');
      if (mob.length !== 11 || !mob.startsWith('09')) {
        return { success: false, message: 'Mobile number must start with 09 and be exactly 11 digits.' };
      }
    }

    const classRank = { '1st Year CCM': 1, 'CCM Evening Class': 1, '2nd Year CCM': 2, 'BCM': 3 };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() !== String(studentID).trim()) continue;

      const currentClass = _norm(data[i][25]);
      const newClass = _norm(updates.classification);

      if (newClass && classRank[newClass] && classRank[currentClass]) {
        if (classRank[newClass] < classRank[currentClass]) {
          return { success: false, message: 'You cannot downgrade your classification.' };
        }
      }

      if (updates.address)              masterSheet.getRange(i + 1, 7).setValue(updates.address);
      if (updates.mobile)               masterSheet.getRange(i + 1, 8).setValue(updates.mobile);
      if (updates.civilStatus)          masterSheet.getRange(i + 1, 12).setValue(updates.civilStatus);
      if (updates.spouse !== undefined) masterSheet.getRange(i + 1, 13).setValue(updates.spouse);
      if (newClass && newClass !== currentClass) {
        masterSheet.getRange(i + 1, 26).setValue(newClass);
      }

      _propagateStudent_(studentID);
      return { success: true, message: 'Profile updated successfully.' };
    }
    return { success: false, message: 'Student not found.' };
  } catch (error) {
    return { success: false, message: 'Update error: ' + error.message };
  }
}

// =============================
// CHANGE PASSWORD
// =============================
function changePassword(studentID, oldPassword, newPassword, token) {
  try {
    _requireStudent_(token, studentID);
    if (String(newPassword || '').length < 8) {
      return { success: false, message: 'New password must be at least 8 characters.' };
    }
    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const data = masterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(studentID).trim()) {
        const check = _verifyPassword_(oldPassword, data[i][28]);
        if (check.ok) {
          masterSheet.getRange(i + 1, MASTER_PW_COL).setValue(_hashPassword_(newPassword));
          _propagateStudent_(studentID);
          return { success: true, message: 'Password changed successfully' };
        } else {
          return { success: false, message: 'Current password is incorrect' };
        }
      }
    }
    return { success: false, message: 'Student not found' };
  } catch (error) {
    return { success: false, message: 'Error changing password: ' + error.message };
  }
}

// =============================
// UPLOAD PROFILE PICTURE
// =============================
function uploadProfilePicture(studentID, imageData, token) {
  try {
    _requireStudent_(token, studentID);
    const matches = imageData.match(/^data:(.+);base64,(.+)$/);
    if (!matches) throw new Error("Invalid image data format");
    const contentType = matches[1];
    const base64Data = matches[2];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, studentID + "_profile.jpg");
    const folderName = "StudentProfilePictures";
    const folders = DriveApp.getFoldersByName(folderName);
    const driveFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const file = driveFolder.createFile(blob);
    // Private. Students fetch their own photo through getProfilePicture(),
    // which checks their session first. A public link would let anyone who
    // guessed or was forwarded the URL view a student's photo.
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    const fileId = file.getId();
    const driveUrl = file.getUrl();
    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const data = masterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(studentID).trim()) {
        const header = masterSheet.getRange(1, 30).getValue();
        if (!header) masterSheet.getRange(1, 30).setValue('Profile Picture URL');
        masterSheet.getRange(i + 1, 30).setValue(driveUrl);
        // Echo the image straight back as a data URL. The file itself is
        // private now, so the old drive.google.com/thumbnail link would 404
        // for the student's browser.
        return {
          success: true,
          message: 'Profile picture uploaded successfully',
          imageUrl: 'data:' + contentType + ';base64,' + base64Data
        };
      }
    }
    return { success: false, message: 'Student not found' };
  } catch (error) {
    return { success: false, message: 'Error uploading picture: ' + error.message };
  }
}

// =============================
// GET PROFILE PICTURE (session-guarded)
// Returns the student's own photo as a data URL. Replaces the public
// drive.google.com/thumbnail link the portal used to embed.
// =============================
function getProfilePicture(studentID, token) {
  try {
    _requireStudent_(token, studentID);
    const sh = _ss().getSheetByName('Master');
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() !== String(studentID).trim()) continue;

      const url = String(data[i][29] || '').trim();
      if (!url) return { success: false, message: 'No profile picture on file.' };

      const m = url.match(/[-\w]{25,}/);
      if (!m) return { success: false, message: 'Stored picture link is unreadable.' };

      const blob = DriveApp.getFileById(m[0]).getBlob();
      return {
        success: true,
        imageUrl: 'data:' + blob.getContentType() + ';base64,' +
                  Utilities.base64Encode(blob.getBytes())
      };
    }
    return { success: false, message: 'Student not found' };
  } catch (error) {
    return { success: false, message: 'Error loading picture: ' + error.message };
  }
}

// =============================
// FORMAT TIME
// =============================
function formatTimeAsText(timeValue) {
  if (!timeValue) return '';
  if (typeof timeValue === 'string') return timeValue;
  if (timeValue instanceof Date) {
    let hours = timeValue.getHours();
    let minutes = timeValue.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    minutes = minutes < 10 ? '0' + minutes : minutes;
    return hours + ':' + minutes + ' ' + ampm;
  }
  return String(timeValue);
}

// =============================
// SUBJECTS SHEET COLUMN RESOLUTION
// =============================
function _resolveSubjectColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => _norm(h).toLowerCase())
    : [];
  const find = (kw, fallback) => {
    for (let i = 0; i < headers.length; i++) {
      if (kw.some(k => headers[i].includes(k))) return i;
    }
    return fallback;
  };
  // Tries each keyword in priority order across all headers, skipping columns
  // already taken, so "Year Level" never lands on a "Semester / S.Y." column.
  const findFirst = (kw, fallback, taken) => {
    for (const k of kw) {
      for (let i = 0; i < headers.length; i++) {
        if (i !== taken && headers[i].includes(k)) return i;
      }
    }
    return fallback !== taken ? fallback : -1;
  };
  const semester  = findFirst(['semester', 'sem'], 5, -1);
  const yearLevel = findFirst(['year level', 'yr level', 'level', 'year', 'period'], 6, semester);
  return {
    subject:    find(['subject', 'course', 'title'], 0),
    day:        find(['day'], 1),
    from:       find(['from', 'start'], 2),
    to:         find(['to', 'end'], 3),
    instructor: find(['instructor', 'faculty', 'teacher'], 4),
    semester:   semester,
    yearLevel:  yearLevel,
    campus:     find(['campus'], 7),
    program:    find(['program'], 8)
  };
}

function _getSubjectsSheetForCampus_(campusUpper) {
  const ss = _ss();
  const isInt = (campusUpper === 'TWI-QATAR' || campusUpper === 'TWI-CANADA' || campusUpper === 'TWI-EUROPE');
  if (isInt) {
    return { sheet: ss.getSheetByName('INT Subjects'), isInt: true };
  }
  return { sheet: ss.getSheetByName('Subjects'), isInt: false };
}

// =============================
// GET AVAILABLE SUBJECTS
// Hides subjects the student already PASSED. Failed ones stay (retake).
// =============================
function getAvailableSubjects(studentCampus, studentClassification, allowedProgramsArg, studentID, token) {
  try {
    _requireStudent_(token, studentID);
    const ss = _ss();
    const campus = _norm(studentCampus).toUpperCase();
    const { sheet: subjectsSheet, isInt } = _getSubjectsSheetForCampus_(campus);

    if (!subjectsSheet) {
      return { success: false, message: (isInt ? "'INT Subjects'" : "'Subjects'") + ' sheet not found' };
    }

    const data = subjectsSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, subjects: [] };

    const cols = _resolveSubjectColumns_(subjectsSheet);

    // Academic record -> which subjects are already passed / currently taken
    let passedSet = [], currentSet = [];
    if (_norm(studentID)) {
      const rec = _getAcademicRecord_(studentID);
      passedSet  = rec.passedSet || [];
      currentSet = (rec.current || []).map(s => s.title.toLowerCase());
    }

    let allowedList = Array.isArray(allowedProgramsArg)
      ? allowedProgramsArg.map(_norm).filter(Boolean)
      : _parseAllowedPrograms_(allowedProgramsArg, studentClassification);
    const allowedFamilies = new Set(allowedList.map(p => _programFamily_(p)));

    const subjects = [];
    for (let i = 1; i < data.length; i++) {
      const subjName = _norm(data[i][cols.subject]);
      if (!subjName) continue;

      const key = subjName.toLowerCase();
      if (passedSet.indexOf(key) >= 0) continue;   // already passed -> hide entirely

      const subjCampus  = _norm(data[i][cols.campus]).toUpperCase();
      const subjProgram = _norm(data[i][cols.program]);
      const subjYear    = _norm(data[i][cols.yearLevel]);

      const campusMatch = isInt
        ? (subjCampus === campus || subjCampus === 'ALL' || subjCampus === '')
        : (subjCampus === 'TWI-PHILIPPINES' || subjCampus === 'ALL' || subjCampus === '');
      if (!campusMatch) continue;

      let allowed = true;
      if (!isInt) {
        const subjFamily = _programFamily_(subjProgram);
        allowed = allowedFamilies.size === 0 || allowedFamilies.has(subjFamily);
      }

      subjects.push({
        subject:    subjName,
        day:        data[i][cols.day] ? String(data[i][cols.day]) : '',
        from:       data[i][cols.from] ? formatTimeAsText(data[i][cols.from]) : '',
        to:         data[i][cols.to]   ? formatTimeAsText(data[i][cols.to])   : '',
        instructor: data[i][cols.instructor] ? String(data[i][cols.instructor]) : '',
        semester:   data[i][cols.semester]   ? String(data[i][cols.semester])   : '',
        period:     subjYear,
        program:    subjProgram,
        campus:     subjCampus,
        allowed:    allowed,
        isEnrolled: currentSet.indexOf(key) >= 0
      });
    }
    return { success: true, subjects };
  } catch (error) {
    return { success: false, message: 'Error loading subjects: ' + error.message };
  }
}

// =============================
// COURSE CATALOG
// =============================
function _getCourseCatalogIndex_() {
  const ss = _ss();
  const sh = ss.getSheetByName('COURSE_CATALOG');
  if (!sh) return { map: new Map(), hasSheet: false };
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { map: new Map(), hasSheet: true };
  const headers = values[0].map(h => _norm(h).toLowerCase());
  const findCol = (keywords, fallbackIdx) => {
    for (let i = 0; i < headers.length; i++) {
      if (keywords.some(k => headers[i].includes(k))) return i;
    }
    return fallbackIdx;
  };
  const colProgram = findCol(['program'], 0);
  const colCode    = findCol(['course code', 'subject code', 'code'], 1);
  const colTitle   = findCol(['subject', 'canonical', 'title', 'course name', 'description'], 2);
  const colUnits   = findCol(['units', 'unit'], 3);
  const map = new Map();
  for (let r = 1; r < values.length; r++) {
    const title = _norm(values[r][colTitle]);
    if (!title) continue;
    const program = _norm(values[r][colProgram]);
    const code = _norm(values[r][colCode]);
    let units = Number(values[r][colUnits]);
    if (!isFinite(units) || units <= 0) units = 3;
    map.set(title.toLowerCase(), { program, code, title, units });
  }
  return { map, hasSheet: true };
}

// =============================
// GET MILESTONE DATA
// Now driven by the academic record (passed/current), not the AA string.
// =============================
function getMilestoneData(studentID, token) {
  try {
    _requireStudent_(token, studentID);
    const ss = _ss();

    const master = ss.getSheetByName('Master');
    const mData = master.getDataRange().getValues();
    let studentRow = null;
    for (let i = 1; i < mData.length; i++) {
      if (_norm(mData[i][1]) === _norm(studentID)) { studentRow = mData[i]; break; }
    }
    if (!studentRow) return { success: false, message: 'Student not found.' };

    const campus = _norm(studentRow[27]).toUpperCase();
    const isInternational = (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE');

    const catalogSheet = ss.getSheetByName('COURSE_CATALOG');
    if (!catalogSheet) return { success: false, message: 'COURSE_CATALOG sheet not found.' };
    const catData = catalogSheet.getDataRange().getValues();
    const headers = catData[0].map(h => _norm(h).toLowerCase());
    const findCol = (kw, fb) => { for (let i=0;i<headers.length;i++) if(kw.some(k=>headers[i].includes(k))) return i; return fb; };
    const colProgram = findCol(['program'],0);
    const colCode    = findCol(['course code','subject code','code'],1);
    const colTitle   = findCol(['subject','canonical','title','course name','description'],2);
    const colUnits   = findCol(['units','unit'],3);

    const ccmSubjects = [], bcmSubjects = [], intSubjects = [];
    for (let r = 1; r < catData.length; r++) {
      const prog  = _norm(catData[r][colProgram]).toUpperCase();
      const code  = _norm(catData[r][colCode]);
      const title = _norm(catData[r][colTitle]);
      const units = Number(catData[r][colUnits]) || 3;
      if (!title) continue;
      const entry = { code, title, units, program: prog };
      if (prog === 'CCM') ccmSubjects.push(entry);
      else if (prog === 'BCM') bcmSubjects.push(entry);
      else if (prog === 'INT') intSubjects.push(entry);
    }

    const { sheet: subjSheet } = _getSubjectsSheetForCampus_(campus);
    const availableNow = new Set();
    if (subjSheet) {
      const sCols = _resolveSubjectColumns_(subjSheet);
      const sData = subjSheet.getDataRange().getValues();
      for (let r = 1; r < sData.length; r++) {
        const name = _norm(sData[r][sCols.subject]);
        if (name) availableNow.add(name.toLowerCase());
      }
    }

    // Truth comes from the academic record
    const rec = _getAcademicRecord_(studentID);
    const completedSubjects = new Set(rec.passedSet);
    const failedSubjects    = new Set(rec.failedSet);
    const enrolledSubjects  = new Set((rec.current || []).map(s => s.title.toLowerCase()));
    const gradeBySubject    = new Map();
    rec.passed.concat(rec.failed).forEach(g => gradeBySubject.set(g.title.toLowerCase(), g));

    function buildTrack(subjects) {
      return subjects.map(s => {
        const tl = s.title.toLowerCase();
        let status = 'grey';
        if (completedSubjects.has(tl)) status = 'red';           // passed
        else if (enrolledSubjects.has(tl)) status = 'orange';    // in progress
        else if (availableNow.has(tl)) status = 'green';         // can enroll
        const out = Object.assign({}, s, { status: status });
        if (failedSubjects.has(tl)) out.failed = true;           // retake needed
        const g = gradeBySubject.get(tl);
        if (g) { out.grade = g.grade; out.gradeSemester = g.semester; }
        return out;
      });
    }

    const intBuilt = buildTrack(intSubjects);
    const intByYear = [];
    for (let y = 0; y < 3; y++) intByYear.push(intBuilt.slice(y * 8, y * 8 + 8));

    function countProgress(track) {
      const total = track.length;
      const done  = track.filter(s => s.status === 'red').length;
      const units = track.filter(s => s.status === 'red').reduce((a,s) => a + s.units, 0);
      const totalUnits = track.reduce((a,s) => a + s.units, 0);
      return { total, done, units, totalUnits };
    }

    const ccmTrack = buildTrack(ccmSubjects);
    const bcmTrack = buildTrack(bcmSubjects);

    return {
      success: true, isInternational, campus,
      ccm: { subjects: ccmTrack, progress: countProgress(ccmTrack) },
      bcm: { subjects: bcmTrack, progress: countProgress(bcmTrack) },
      int: { byYear: intByYear, progress: countProgress(intBuilt) }
    };
  } catch (err) {
    return { success: false, message: 'Milestone error: ' + err.message };
  }
}

// =============================
// ENROLLMENT SUMMARY (per-program caps)
// =============================
function _getCurrentSemEnrollmentSummary_(studentID) {
  const ss = _ss();
  const { map: catalog } = _getCourseCatalogIndex_();

  const master = ss.getSheetByName('Master');
  const mData  = master.getDataRange().getValues();
  let classification = '', campus = '';
  for (let i = 1; i < mData.length; i++) {
    if (_norm(mData[i][1]) === _norm(studentID)) {
      classification = _norm(mData[i][25]);
      campus         = _norm(mData[i][27]).toUpperCase();
      break;
    }
  }
  const sem = _getCurrentSemesterValue_(campus) || '';
  const isInt = (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE');

  // Current load comes from the academic record, not the AA string.
  const rec = _getAcademicRecord_(studentID);
  const subjects = rec.current.map(s => s.title);
  const totalUnits = rec.unitsCurrent;
  const subjectCount = subjects.length;

  // Map each current subject to its program family
  const subjProgramMap = new Map();
  const { sheet: subjSheet } = _getSubjectsSheetForCampus_(campus);
  if (subjSheet) {
    const sCols = _resolveSubjectColumns_(subjSheet);
    const sData = subjSheet.getDataRange().getValues();
    for (let r = 1; r < sData.length; r++) {
      const name = _norm(sData[r][sCols.subject]);
      if (!name) continue;
      subjProgramMap.set(name.toLowerCase(), _programFamily_(_norm(sData[r][sCols.program])));
    }
  }

  const perProgram = {};
  rec.current.forEach(s => {
    const fam = subjProgramMap.get(s.title.toLowerCase()) || _programFamily_(classification);
    if (!perProgram[fam]) perProgram[fam] = { subjectCount: 0, units: 0 };
    perProgram[fam].subjectCount += 1;
    perProgram[fam].units += (Number(s.units) || 3);
  });

  if (isInt) {
    const reached3 = totalUnits === 3;
    const reached6 = totalUnits === 6;
    return {
      isInt: true, semester: sem, subjectCount, totalUnits, subjects, perProgram,
      maxReached: (subjectCount >= INT_CAPS.maxSubjects) || (totalUnits >= INT_CAPS.maxUnits),
      classification,
      reached3, reached6, reached12: false, showPopup: (reached3 || reached6),
      maxSubjects: INT_CAPS.maxSubjects, maxUnits: INT_CAPS.maxUnits
    };
  }

  const programStatus = {};
  let anyCapReached = false;
  Object.keys(perProgram).forEach(fam => {
    const caps = _getProgramCaps_(fam === 'CCM Evening Class' ? 'CCM Evening Class' : fam);
    const pc = perProgram[fam];
    const reached = (pc.subjectCount >= caps.maxSubjects) || (pc.units >= caps.maxUnits);
    programStatus[fam] = {
      subjectCount: pc.subjectCount, units: pc.units,
      maxSubjects: caps.maxSubjects, maxUnits: caps.maxUnits, maxReached: reached
    };
    if (reached) anyCapReached = true;
  });

  const activeFamilies = Object.keys(perProgram);
  const isEveningOnly = (activeFamilies.length === 1 && activeFamilies[0] === 'CCM Evening Class');
  // The load limit is per student (all programs combined), set by classification.
  const popupCaps = _getProgramCaps_(classification);
  if (subjectCount >= popupCaps.maxSubjects || totalUnits >= popupCaps.maxUnits) anyCapReached = true;
  const reached3  = totalUnits === 3;
  const reached6  = totalUnits === 6;
  const reached12 = !isEveningOnly && totalUnits === 12;

  return {
    isInt: false, semester: sem, subjectCount, totalUnits, subjects,
    perProgram: programStatus,
    reached3, reached6, reached12,
    showPopup: (reached3 || reached6 || reached12),
    maxReached: anyCapReached,
    maxSubjects: popupCaps.maxSubjects, maxUnits: popupCaps.maxUnits,
    classification
  };
}

function getCurrentSemEnrollmentSummary(studentID, token) {
  _requireStudent_(token, studentID);
  return _getCurrentSemEnrollmentSummary_(studentID);
}

// =============================
// ENROLL IN SUBJECT
// Enrollments is the source of truth. AA/AF are rebuilt, never appended.
// =============================
function enrollInSubject(studentID, subjectName, token) {
  try {
    _requireStudent_(token, studentID);
    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const data = masterSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() !== String(studentID).trim()) continue;

      const studentClass = _norm(data[i][25]);
      const campus       = _norm(data[i][27]).toUpperCase();
      const allowedCell  = _norm(data[i][30]);
      const isInt = (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE');

      const subjectInfo = _lookupSubjectInfo_(subjectName, campus);
      const subjFamily  = _programFamily_(subjectInfo.program);

      // ----- Already PASSED, or already enrolled this sem? -----
      const block = _blockIfAlreadyTaken_(studentID, subjectName);
      if (block) return { success: false, message: block };

      // ----- Eligibility (PH only) -----
      if (!isInt) {
        const allowedList = _parseAllowedPrograms_(allowedCell, studentClass);
        const allowedFamilies = new Set(allowedList.map(p => _programFamily_(p)));
        if (allowedFamilies.size > 0 && subjFamily && !allowedFamilies.has(subjFamily)) {
          return { success: false, message: 'You are not allowed to enroll in this subject. It belongs to a program you are not enrolled in.' };
        }
      }

      // ----- Load cap check -----
      const summaryBefore = _getCurrentSemEnrollmentSummary_(studentID);
      if (isInt) {
        const subjUnits = (subjectInfo.units && Number(subjectInfo.units)) || _catalogUnits_(subjectName) || 3;
        if ((summaryBefore.subjectCount || 0) + 1 > INT_CAPS.maxSubjects ||
            (summaryBefore.totalUnits || 0) + subjUnits > INT_CAPS.maxUnits) {
          return {
            success: false,
            message: 'International students are limited to ' + INT_CAPS.maxSubjects + ' subjects / ' +
                     INT_CAPS.maxUnits + ' units per semester. Maximum reached.',
            summary: summaryBefore
          };
        }
      } else {
        const subjUnits = (subjectInfo.units && Number(subjectInfo.units)) || _catalogUnits_(subjectName) || 3;

        // Total load across ALL programs: 4 subjects / 12 units per semester
        // (CCM Evening Class students: 2 subjects / 6 units). Matches the fee table.
        const totalCaps = _getProgramCaps_(studentClass);
        if ((summaryBefore.subjectCount || 0) + 1 > totalCaps.maxSubjects ||
            (summaryBefore.totalUnits || 0) + subjUnits > totalCaps.maxUnits) {
          return {
            success: false,
            message: 'You can enroll in up to ' + totalCaps.maxSubjects + ' subjects (' + totalCaps.maxUnits +
                     ' units) per semester. Maximum load reached.',
            summary: summaryBefore
          };
        }

        const caps = _getProgramCaps_(subjFamily === 'CCM Evening Class' ? 'CCM Evening Class' : subjFamily);
        const ps = (summaryBefore.perProgram && summaryBefore.perProgram[subjFamily]) || { subjectCount: 0, units: 0 };
        const wouldBeSubjects = (ps.subjectCount || 0) + 1;
        const wouldBeUnits    = (ps.units || 0) + subjUnits;

        if ((ps.subjectCount || 0) >= caps.maxSubjects || (ps.units || 0) >= caps.maxUnits) {
          const isEvening = (subjFamily === 'CCM Evening Class');
          return {
            success: false,
            message: isEvening
              ? 'CCM Evening Class is limited to 2 subjects / 6 units per semester. Maximum reached for this program.'
              : subjFamily + ' is limited to 4 subjects / 12 units per semester. Maximum reached for this program.',
            summary: summaryBefore
          };
        }
        if (wouldBeSubjects > caps.maxSubjects || wouldBeUnits > caps.maxUnits) {
          return {
            success: false,
            message: 'Enrolling in this subject would exceed the ' + subjFamily +
                     ' limit (' + caps.maxSubjects + ' subjects / ' + caps.maxUnits + ' units).',
            summary: summaryBefore
          };
        }
      }

      // ----- Write to Enrollments (the source of truth) -----
      _upsertEnrollmentFromStudentRow_(data[i], subjectName, campus);

      // ----- AA + AF are DERIVED -----
      const rebuilt = _rebuildStudentSubjects_(studentID);
      _propagateStudent_(studentID);

      const summaryAfter = _getCurrentSemEnrollmentSummary_(studentID);
      return {
        success: true,
        message: 'Successfully enrolled in ' + subjectName,
        enrolledSubjects: rebuilt.current,
        completedSubjects: rebuilt.completed,
        summary: summaryAfter
      };
    }
    return { success: false, message: 'Student not found' };
  } catch (error) {
    return { success: false, message: 'Enrollment error: ' + error.message };
  }
}

function _catalogUnits_(subjectName) {
  const { map } = _getCourseCatalogIndex_();
  const info = map.get(String(subjectName).trim().toLowerCase());
  return info ? (Number(info.units) || 0) : 0;
}

function _buildStudentName_(masterRow) {
  const surname = String(masterRow[3] || '').trim();
  const first   = String(masterRow[4] || '').trim();
  const middle  = String(masterRow[5] || '').trim();
  const middlePart = middle ? (' ' + middle) : '';
  return (surname + ', ' + first + middlePart).trim();
}

function _lookupSubjectInfo_(subjectName, campusUpper) {
  const { sheet } = _getSubjectsSheetForCampus_(_norm(campusUpper).toUpperCase());
  if (!sheet) return { instructor:'', semester:'', yearLevel:'', program:'', units:0 };
  const cols = _resolveSubjectColumns_(sheet);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const sub = _norm(data[i][cols.subject]);
    if (sub && sub.toLowerCase() === String(subjectName).trim().toLowerCase()) {
      return {
        instructor: _norm(data[i][cols.instructor]),
        semester:   _norm(data[i][cols.semester]),
        yearLevel:  _norm(data[i][cols.yearLevel]),
        program:    _norm(data[i][cols.program]),
        units:      _catalogUnits_(sub)
      };
    }
  }
  return { instructor:'', semester:'', yearLevel:'', program:'', units:0 };
}

function _upsertEnrollmentFromStudentRow_(masterRow, subjectName, campusUpper) {
  const ss = _ss();
  const enroll = ss.getSheetByName('Enrollments');
  const studentID   = String(masterRow[1] || '').trim();
  const studentName = _buildStudentName_(masterRow);
  const info        = _lookupSubjectInfo_(subjectName, campusUpper);
  // Always stamp with the CURRENT semester from Settings (single source of truth).
  const semester    = _getCurrentSemesterValue_(campusUpper) || info.semester || '';
  const instructor  = info.instructor || '';
  const yearLevel   = info.yearLevel || '';
  const data = enroll.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === semester &&
        String(data[i][2]).trim() === studentID &&
        String(data[i][4]).trim().toLowerCase() === String(subjectName).trim().toLowerCase()) {
      enroll.getRange(i + 1, 6).setValue(instructor);
      enroll.getRange(i + 1, 7).setValue(yearLevel);
      enroll.getRange(i + 1, 8).setValue('ENROLLED');
      return;
    }
  }
  enroll.appendRow([new Date(), semester, studentID, studentName, subjectName, instructor, yearLevel, 'ENROLLED']);
}

// =============================
// INSTRUCTOR PORTAL BACKEND
// =============================
function checkInstructorLogin(username, password) {
  const ss = _ss();
  const sh = ss.getSheetByName('Instructors');
  if (!sh) return { success: false, message: 'Instructors sheet not found.' };
  const data = sh.getDataRange().getValues();
  const u = String(username || '').trim();
  const p = String(password || '').trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== u) continue;

    const check = _verifyPassword_(p, data[i][1]);
    if (!check.ok) break;

    // Same transparent upgrade as students: a legacy plaintext instructor
    // password is rehashed on first successful login.
    if (check.needsUpgrade) _upgradeStoredPassword_(sh, i + 1, 2, p);

    return {
      success: true,
      username: u,
      fullName: String(data[i][2] || u).trim(),
      token: _issueSession_('instructor', u, 'INSTRUCTOR')
    };
  }
  return { success: false, message: 'Invalid instructor username or password.' };
}

function getInstructorSubjects(instructorUsername, token) {
  _requireInstructor_(token, instructorUsername);
  const ss = _ss();
  const instSheet = ss.getSheetByName('Instructors');
  const instData  = instSheet.getDataRange().getValues();
  let fullName = '';
  let subjectHandles = [];
  for (let i = 1; i < instData.length; i++) {
    if (_norm(instData[i][0]) === _norm(instructorUsername)) {
      fullName = _norm(instData[i][2]);
      for (let c = 4; c <= 25; c++) {
        const s = _norm(instData[i][c]);
        if (s) subjectHandles.push(s);
      }
      break;
    }
  }
  if (!fullName) return { success: false, message: 'Instructor not found in Instructors sheet.' };
  const semPH  = _getCurrentSemesterValue_();
  const semINT = _getCurrentSemesterValue_('TWI-QATAR');

  // Offering rows per subject, split by side: 'Subjects' = PH, 'INT Subjects' = INT.
  const me = fullName.toLowerCase();
  const offerings = { PH: new Map(), INT: new Map() };
  [['Subjects', 'PH'], ['INT Subjects', 'INT']].forEach(([sheetName, side]) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const cols = _resolveSubjectColumns_(sh);
    const sData = sh.getDataRange().getValues();
    for (let r = 1; r < sData.length; r++) {
      const subName = _norm(sData[r][cols.subject]);
      if (!subName) continue;
      const key = subName.toLowerCase();
      if (!offerings[side].has(key)) offerings[side].set(key, []);
      offerings[side].get(key).push({
        row: sData[r], cols,
        mine: _norm(sData[r][cols.instructor]).toLowerCase() === me
      });
    }
  });

  const card = (subName, side, rows) => {
    const r0 = rows[0];
    const out = { subject: subName, instructor: fullName, side: side,
                  semester: (side === 'INT' ? semINT : semPH) || '',
                  day: '', from: '', to: '', yearLevel: '', campuses: side === 'INT' ? 'International' : 'TWI-PHILIPPINES' };
    if (r0) {
      const row = r0.row, cols = r0.cols;
      out.day  = _norm(row[cols.day]);
      out.from = row[cols.from] ? formatTimeAsText(row[cols.from]) : '';
      out.to   = row[cols.to]   ? formatTimeAsText(row[cols.to])   : '';
      out.yearLevel = _norm(row[cols.yearLevel]);
      if (side === 'INT') {
        const list = rows.map(x => _norm(x.row[x.cols.campus]).toUpperCase()).filter(Boolean);
        if (list.length) out.campuses = Array.from(new Set(list)).join(', ');
        // Same subject at several campuses with different times: list each one.
        if (rows.length > 1) {
          out.schedule = rows.map(x => {
            const c = x.cols, rw = x.row;
            const f = rw[c.from] ? formatTimeAsText(rw[c.from]) : '', t = rw[c.to] ? formatTimeAsText(rw[c.to]) : '';
            return _norm(rw[c.day]) + ' ' + f + '–' + t + ' (' + _norm(rw[c.campus]).toUpperCase().replace('TWI-', '') + ')';
          }).join('  ·  ');
        }
      }
    }
    return out;
  };

  // Every subject in columns E–Z is shown — never hidden by how the Semester
  // column is spelled. A subject taught on both sides gets one card per side,
  // because PH and INT run separate terms (separate class lists and grades).
  const subjects = [];
  const seen = new Set();
  subjectHandles.forEach(subName => {
    const key = subName.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    const ph  = offerings.PH.get(key)  || [];
    const int = offerings.INT.get(key) || [];
    const phMine = ph.filter(x => x.mine), intMine = int.filter(x => x.mine);
    if (phMine.length)  subjects.push(card(subName, 'PH', phMine));
    if (intMine.length) subjects.push(card(subName, 'INT', intMine));
    if (!phMine.length && !intMine.length) {
      // Assigned here but the offering says TBA / someone else: show it anyway.
      if (ph.length || !int.length) subjects.push(card(subName, 'PH', ph));
      else subjects.push(card(subName, 'INT', int));
    }
  });
  return { success: true, subjects, currentSemester: semPH, currentSemesterInt: semINT, fullName };
}

function getEnrolledStudentsForSubject(instructorUsername, subjectName, semester, token) {
  _requireInstructor_(token, instructorUsername);
  const ss = _ss();
  const instSheet = ss.getSheetByName('Instructors');
  const instData  = instSheet.getDataRange().getValues();
  let fullName = '';
  const handles = new Set();
  for (let i = 1; i < instData.length; i++) {
    if (_norm(instData[i][0]) === _norm(instructorUsername)) {
      fullName = _norm(instData[i][2]);
      for (let c = 4; c <= 25; c++) { const h = _norm(instData[i][c]).toLowerCase(); if (h) handles.add(h); }
      break;
    }
  }
  if (!fullName) return { success: false, message: 'Instructor not found.' };
  const assignedToMe = handles.has(_norm(subjectName).toLowerCase());

  const enroll = ss.getSheetByName('Enrollments');
  const data   = enroll.getDataRange().getValues();
  const sem    = _norm(semester);
  const sub    = _norm(subjectName).toLowerCase();

  const gradesSheet = ss.getSheetByName('Grades');
  const grades      = gradesSheet.getDataRange().getValues();
  const gradeMap    = new Map();
  for (let g = 1; g < grades.length; g++) {
    const gSem = _norm(grades[g][1]);
    const gSub = _norm(grades[g][2]).toLowerCase();
    const gSid = _norm(grades[g][3]);
    if (!gSem || !gSub || !gSid) continue;
    const instructor = _norm(grades[g][5]);
    const locked     = (instructor === fullName);
    gradeMap.set([gSem,gSub,gSid].join('|'), { grade: grades[g][6] || '', remarks: grades[g][7] || '', locked });
  }

  const students = [];
  for (let i = 1; i < data.length; i++) {
    const rowSem  = _norm(data[i][1]);
    const rowSid  = _norm(data[i][2]);
    const rowName = _norm(data[i][3]);
    const rowSub  = _norm(data[i][4]).toLowerCase();
    const rowInst = _norm(data[i][5]);
    const status  = _norm(data[i][7]).toUpperCase();
    if (rowSem.toLowerCase() !== sem.toLowerCase()) continue;
    if (rowSub !== sub) continue;
    // Enrollments made while the offering said "TBA" (or had no instructor)
    // belong to whoever now has the subject in the Instructors sheet.
    const unassigned = !rowInst || /\btba\b|to be announced/i.test(rowInst);
    if (!(unassigned && assignedToMe) && rowInst.toLowerCase() !== fullName.toLowerCase()) continue;
    if (status === 'DROPPED') continue;
    const key      = [sem, sub, rowSid].join('|');
    const existing = gradeMap.get(key) || { grade:'', remarks:'', locked:false };
    students.push({ studentId: rowSid, studentName: rowName, grade: existing.grade, remarks: existing.remarks, locked: !!existing.locked });
  }
  return { success: true, students, instructor: fullName };
}

function saveGrades(instructorUsername, semester, subjectName, gradesArray, token) {
  _requireInstructor_(token, instructorUsername);
  const ss = _ss();
  const instSheet = ss.getSheetByName('Instructors');
  const instData  = instSheet.getDataRange().getValues();
  let fullName = '';
  for (let i = 1; i < instData.length; i++) {
    if (_norm(instData[i][0]) === _norm(instructorUsername)) { fullName = _norm(instData[i][2]); break; }
  }
  if (!fullName) return { success: false, message: 'Instructor not found.' };

  const gradesSheet = ss.getSheetByName('Grades');
  const data        = gradesSheet.getDataRange().getValues();
  const sem = _norm(semester);
  const sub = _norm(subjectName);
  const idx = new Map();
  for (let i = 1; i < data.length; i++) {
    const key = [_norm(data[i][1]), _norm(data[i][2]).toLowerCase(), _norm(data[i][3])].join('|');
    idx.set(key, { rowNum: i + 1, instructor: _norm(data[i][5]), grade: _norm(data[i][6]) });
  }

  let saved = 0, lockedCount = 0, skipped = 0;
  const results = [];
  gradesArray = gradesArray || [];
  gradesArray.forEach(g => {
    const sid     = _norm(g.studentId);
    const sname   = _norm(g.studentName);
    const grade   = _norm(g.grade);
    const remarks = _norm(g.remarks);
    if (!sid) { results.push({ studentId: '', status: 'skipped', reason: 'Missing Student ID' }); return; }
    // An empty box is "not graded yet" — never write it, or it locks the student.
    if (!grade) { skipped++; results.push({ studentId: sid, status: 'skipped', reason: 'No grade entered' }); return; }
    const key   = [sem, sub.toLowerCase(), sid].join('|');
    const found = idx.get(key);
    // A blank row left by the old Save All bug is filled in, not treated as locked.
    if (found && !found.grade) {
      gradesSheet.getRange(found.rowNum, 1, 1, 8).setValues([[new Date(), sem, sub, sid, sname, fullName, grade, remarks]]);
      saved++;
      results.push({ studentId: sid, status: 'saved' });
      return;
    }
    if (found && found.instructor === fullName) { lockedCount++; results.push({ studentId: sid, status: 'locked' }); return; }
    if (found && found.instructor && found.instructor !== fullName) { lockedCount++; results.push({ studentId: sid, status: 'locked', reason: 'Already graded by another instructor' }); return; }
    gradesSheet.appendRow([new Date(), sem, sub, sid, sname, fullName, grade, remarks]);
    saved++;
    results.push({ studentId: sid, status: 'saved' });
  });

  // A grade moves a subject CURRENT -> PASSED/FAILED. Rebuild AA + AF.
  const touched = new Set();
  gradesArray.forEach(g => { const s = _norm(g.studentId); if (s) touched.add(s); });
  touched.forEach(sid => {
    try { _rebuildStudentSubjects_(sid); _propagateStudent_(sid); } catch (e) {}
  });

  let msg = 'Saved grades for ' + saved + ' student(s).';
  if (lockedCount) msg += ' Locked: ' + lockedCount + ' (already saved before).';
  if (skipped) msg += ' Skipped: ' + skipped + ' (no grade entered).';
  return { success: true, message: msg, results };
}

// =============================
// CERTIFICATE OF REGISTRATION
// =============================
function createAndSendCertificateOfRegistration(studentID, token) {
  try {
    _requireStudent_(token, studentID);
    const ss = _ss();
    const master = ss.getSheetByName('Master');
    if (!master) throw new Error('Master sheet not found.');
    const mData = master.getDataRange().getValues();
    let row = null;
    for (let i = 1; i < mData.length; i++) {
      if (_norm(mData[i][1]) === _norm(studentID)) { row = mData[i]; break; }
    }
    if (!row) throw new Error('Student not found in Master sheet.');

    const sem = _getCurrentSemesterValue_(_norm(row[27])) || '';
    if (!sem) throw new Error('Current Semester is not set. Please set it in the Settings sheet.');

    const email           = _norm(row[2]).toLowerCase();
    const first           = _norm(row[4]);
    const address         = _norm(row[6]);
    const sex             = _norm(row[10]);
    const classification  = _norm(row[25]);
    const campus          = _norm(row[27]).toUpperCase();
    const isInternational = (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE');
    const studentName     = _buildStudentName_(row);
    const programType     = (classification.toUpperCase().includes('BCM')) ? 'BCM' : 'CCM';

    // COR uses the academic record's CURRENT subjects.
    const rec = _getAcademicRecord_(studentID);
    if (!rec.current.length) throw new Error('No enrolled subjects found for the current semester.');

    const subjects = rec.current.map(s => ({
      code: s.code || '', title: s.title, units: s.units, instructor: s.instructor, yearLevel: s.yearLevel
    }));
    const totalUnits = rec.unitsCurrent;
    const fees = _computeFeesByUnits_(totalUnits, isInternational);

    const html = _buildCORHtml_({
      semester: sem, programType, classification, campus,
      studentID: _norm(studentID), studentName, address, sex, email, subjects,
      totals: { totalUnits, subjectCount: subjects.length }, fees
    });

    const blob = HtmlService.createHtmlOutput(html).setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .getBlob().setName('COR_' + _norm(studentID) + '_' + _safeFile_(sem) + '.pdf');
    const pdf  = blob.getAs(MimeType.PDF);
    const folder = _getOrCreateCorFolder_(sem, programType);
    const file   = folder.createFile(pdf);
    // Private. The COR carries the student's name, address, sex, email and
    // fees — it goes out as an email attachment to the student, not as a
    // public link anyone could open.
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

    const subject = 'Certificate of Registration - ' + sem;
    const body =
      'Dear ' + (first || 'Student') + ',\n\n' +
      'Attached is your Certificate of Registration (COR) for ' + sem + '.\n\n' +
      'Student ID: ' + studentID + '\nName: ' + studentName + '\nCampus: ' + campus + '\nProgram: ' + classification + '\n\n' +
      'Your COR is attached to this email. Please keep it for your records.\n\nGod bless you!\n\nThy Word Intl Bible College Bataan';
    if (!email) throw new Error('Student email is empty in Master sheet.');
    MailApp.sendEmail({ to: email, subject, body, attachments: [pdf] });

    return { success: true, message: 'COR generated and sent to student email successfully.',
             fileUrl: file.getUrl(), fileId: file.getId(), semester: sem, programType,
             totalUnits, subjectCount: subjects.length };
  } catch (err) {
    return { success: false, message: 'COR error: ' + err.message };
  }
}

// =============================
// COMPUTE FEES BY UNITS
// =============================
// Total per semester, keyed by units (3 units = 1 subject).
const FEES_PH_TOTAL  = { 3: 2000, 6: 4000, 9: 5000, 12: 6000 };
const FEES_INT_TOTAL = { 3: 3000, 6: 6000 };  // international max is 6 units (INT_CAPS)

function _computeFeesByUnits_(totalUnits, isInternational) {
  // Registration + miscellaneous are part of the total; tuition is the rest.
  const reg  = isInternational ? 1000 : 400;
  const misc = isInternational ? 1000 : 550;
  let total;
  if (isInternational) {
    total = FEES_INT_TOTAL[totalUnits] || (totalUnits / 3) * 3000;
  } else {
    total = FEES_PH_TOTAL[totalUnits] || (totalUnits * 350 + reg + misc);
  }
  return { registrationFee: reg, miscellaneousFee: misc, tuitionFee: total - reg - misc, totalAssessment: total };
}

function checkDuplicateRegistration(email, surname, firstName, middleName) {
  try {
    const ss = _ss();
    const master = ss.getSheetByName('Master');
    const data = master.getDataRange().getValues();
    const e  = String(email      || '').trim().toLowerCase();
    const sn = String(surname    || '').trim().toUpperCase();
    const fn = String(firstName  || '').trim().toUpperCase();
    const mn = String(middleName || '').trim().toUpperCase();
    for (let i = 1; i < data.length; i++) {
      const rowEmail   = String(data[i][2] || '').trim().toLowerCase();
      const rowSurname = String(data[i][3] || '').trim().toUpperCase();
      const rowFirst   = String(data[i][4] || '').trim().toUpperCase();
      const rowMiddle  = String(data[i][5] || '').trim().toUpperCase();
      if (rowEmail === e && rowSurname === sn && rowFirst === fn && rowMiddle === mn) {
        // Deliberately does NOT return the Student ID. This function is
        // callable anonymously, so echoing the ID back would turn it into a
        // lookup oracle for anyone who knows a name and email. The student
        // is told to use Forgot Password, which mails the ID to the address
        // already on file.
        return { exists: true };
      }
    }
    return { exists: false };
  } catch (err) {
    return { exists: false };
  }
}

function _getOrCreateCorFolder_(semester, programType) {
  const rootIt = DriveApp.getFoldersByName('TWI_COR');
  const root   = rootIt.hasNext() ? rootIt.next() : DriveApp.createFolder('TWI_COR');
  const corIt  = root.getFoldersByName('COR');
  const cor    = corIt.hasNext() ? corIt.next() : root.createFolder('COR');
  const semName = _safeFile_(semester || 'UNKNOWN_SEM');
  const semIt   = cor.getFoldersByName(semName);
  const semFolder = semIt.hasNext() ? semIt.next() : cor.createFolder(semName);
  const prog    = (String(programType || 'CCM').toUpperCase().includes('BCM')) ? 'BCM' : 'CCM';
  const progIt  = semFolder.getFoldersByName(prog);
  return progIt.hasNext() ? progIt.next() : semFolder.createFolder(prog);
}

function _safeFile_(name) {
  return String(name || '').replace(/[\\\/:*?"<>|]+/g, ' ').trim();
}

function _peso_(n) {
  const num = Number(n) || 0;
  return '\u20B1' + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function _buildCORHtml_(ctx) {
  const datePrinted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');
  const subjectsRows = ctx.subjects.map(s => `
    <tr>
      <td class="td">${s.code||''}</td>
      <td class="td">${s.title||''}</td>
      <td class="td center">${s.units||''}</td>
      <td class="td">${s.instructor||''}</td>
    </tr>`).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Certificate of Registration</title>
  <style>
    @page{size:A4;margin:18mm}body{font-family:Arial,sans-serif;color:#111}
    .header{text-align:center}.school{font-size:16px;font-weight:700}.sub{font-size:12px;margin-top:2px}
    .title{margin:14px 0 10px;text-align:center;font-size:14px;font-weight:800;letter-spacing:1px}
    .box{border:1px solid #333;padding:10px;margin-bottom:10px}.row{display:flex;gap:12px}.col{flex:1}
    .label{font-size:11px;color:#333}.value{font-size:12px;font-weight:700;margin-top:2px}
    table{width:100%;border-collapse:collapse}
    .th{background:#e9ecef;border:1px solid #333;padding:6px;font-size:11px;text-align:left}
    .td{border:1px solid #333;padding:6px;font-size:11px}.center{text-align:center}.right{text-align:right}
    .fees{width:55%;margin-top:10px}.foot{margin-top:18px;font-size:10px;color:#444}
    .signRow{display:flex;gap:20px;margin-top:20px}.sign{flex:1;text-align:center}
    .line{border-top:1px solid #111;margin-top:30px}
  </style></head><body>
  <div class="header">
    <div class="school">Republic of the Philippines</div>
    <div class="school">THY WORD INTL BIBLE COLLEGE BATAAN</div>
    <div class="sub">City of Balanga, Bataan</div>
  </div>
  <div class="title">CERTIFICATE OF REGISTRATION</div>
  <div class="box">
    <div class="row">
      <div class="col"><div class="label">Student No.</div><div class="value">${ctx.studentID}</div></div>
      <div class="col"><div class="label">Student Name</div><div class="value">${ctx.studentName}</div></div>
      <div class="col"><div class="label">Semester</div><div class="value">${ctx.semester}</div></div>
    </div>
    <div class="row" style="margin-top:10px;">
      <div class="col"><div class="label">Campus</div><div class="value">${ctx.campus}</div></div>
      <div class="col"><div class="label">Type</div><div class="value">${ctx.programType}</div></div>
      <div class="col"><div class="label">Sex</div><div class="value">${ctx.sex}</div></div>
    </div>
    <div style="margin-top:10px;"><div class="label">Address</div><div class="value" style="font-weight:600;">${ctx.address||''}</div></div>
  </div>
  <div class="box">
    <div class="label" style="font-weight:700;margin-bottom:6px;">SUBJECTS ENROLLED</div>
    <table><thead><tr>
      <th class="th" style="width:15%;">CODE</th><th class="th">SUBJECT TITLE</th>
      <th class="th center" style="width:10%;">UNITS</th><th class="th" style="width:25%;">FACULTY</th>
    </tr></thead><tbody>
      ${subjectsRows}
      <tr><td class="td right" colspan="2"><b>TOTAL UNITS</b></td><td class="td center"><b>${ctx.totals.totalUnits}</b></td><td class="td"></td></tr>
    </tbody></table>
    <table class="fees">
      <tr><td class="td">Registration Fee</td><td class="td right">${_peso_(ctx.fees.registrationFee)}</td></tr>
      <tr><td class="td">Miscellaneous Fee</td><td class="td right">${_peso_(ctx.fees.miscellaneousFee)}</td></tr>
      <tr><td class="td">Tuition Fee</td><td class="td right">${_peso_(ctx.fees.tuitionFee)}</td></tr>
      <tr><td class="td"><b>Total Assessment</b></td><td class="td right"><b>${_peso_(ctx.fees.totalAssessment)}</b></td></tr>
    </table>
    <div class="signRow">
      <div class="sign"><div class="line"></div><div class="label">Student's Signature</div></div>
      <div class="sign"><div class="line"></div><div class="label">Sis. Judilyn C. Acda</div><div class="label">Registrar</div></div>
    </div>
    <div class="foot">
      Date Printed: ${datePrinted}<br/>
      Keep this certificate. You will be required to present this in all your dealings with the College.
    </div>
    <br/>Send your payment via any of the options below.<br/>
    After paying, send a message with your proof of payment to our email: twibiblecollege@gmail.com to confirm your enrollment.<br/>
    Gcash<br/>Joana A. Taguiam<br/>09088982181<br/>
    or Joana A. Taguiam<br/>BDO Acct No: 013340053140
  </div>
</body></html>`;
}

// =============================
// CONSENT LOG
// Append-only record of what each student agreed to, and when.
// =============================
const CONSENT_HEADERS = [
  'Timestamp','Student ID','Email','Full Name',
  'Processing (required)','Sensitive Personal Info (required)',
  'Photo/Media (optional)','Emergency Contact Notified (required)',
  'Terms Accepted (required)','Policy Version','Browser Timezone'
];

// Bump this whenever the wording of the consent block or the policies
// materially changes, so old records stay attributable to the text that was
// actually shown at the time.
const CONSENT_POLICY_VERSION = '2026-09-10';

function _recordConsent_(studentID, formData) {
  try {
    const sh = _getOrCreateSheet('Consent Log', CONSENT_HEADERS);
    const name = [formData.surname, formData.firstName, formData.middleName]
                   .map(function (v) { return _norm(v); })
                   .filter(Boolean).join(', ');
    sh.appendRow([
      new Date(),
      _norm(studentID),
      String(formData.email || '').toLowerCase().trim(),
      name,
      formData.consentProcessing ? 'YES' : 'NO',
      formData.consentSensitive  ? 'YES' : 'NO',
      formData.consentPhoto      ? 'YES' : 'NO',
      formData.consentEmergency  ? 'YES' : 'NO',
      formData.consentTerms      ? 'YES' : 'NO',
      CONSENT_POLICY_VERSION,
      _norm(formData.consentTimezone)
    ]);
  } catch (e) {
    // Never block a registration because the log failed, but make it loud.
    console.error('Consent log failed for ' + studentID + ': ' + e.message);
  }
}

// Lets a student see, and the College evidence, what was agreed.
function getConsentRecord(studentID, token) {
  _requireStudent_(token, studentID);
  const sh = _ss().getSheetByName('Consent Log');
  if (!sh) return { success: false, message: 'No consent log exists yet.' };
  const d = sh.getDataRange().getValues();
  for (let r = d.length - 1; r >= 1; r--) {
    if (_norm(d[r][1]) === _norm(studentID)) {
      return { success: true, record: {
        timestamp: d[r][0], processing: d[r][4], sensitive: d[r][5],
        photo: d[r][6], emergency: d[r][7], terms: d[r][8],
        policyVersion: d[r][9]
      }};
    }
  }
  return { success: false, message: 'No consent record found for that Student ID.' };
}
