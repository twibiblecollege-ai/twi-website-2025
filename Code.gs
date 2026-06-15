// ========================================
// THY WORD INTL BIBLE COLLEGE BATAAN
// Registration + Student Portal + Instructor Portal
// Google Apps Script Backend Code
// ========================================

const SPREADSHEET_ID = '1ISbv7Hso14xeupMog3OdwQS5oMDbLRPE7IKBcp2J_nI';

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
// PROGRAM CONSTANTS / HELPERS  (NEW)
// =============================
// Canonical PH programs the student can be admin-allowed into.
const PH_PROGRAMS = ['1st Year CCM', '2nd Year CCM', 'BCM', 'CCM Evening Class'];

// Per-program load caps (PH only). Evening = 2 subjects / 6 units, others = 4 / 12.
function _getProgramCaps_(program) {
  const p = _norm(program).toLowerCase();
  if (p === 'ccm evening class') return { maxSubjects: 2, maxUnits: 6 };
  return { maxSubjects: 4, maxUnits: 12 };
}

// Which "program family" a PH classification belongs to (for grouping subjects/caps).
// 1st & 2nd Year CCM both belong to the "CCM" family; Evening is its own; BCM its own.
function _programFamily_(program) {
  const p = _norm(program).toLowerCase();
  if (p === 'ccm evening class') return 'CCM Evening Class';
  if (p === 'bcm') return 'BCM';
  if (p === '1st year ccm' || p === '2nd year ccm' || p === 'ccm') return 'CCM';
  return _norm(program);
}

// Parse the comma-separated "Allowed Programs" cell. Falls back to single classification.
function _parseAllowedPrograms_(allowedCell, fallbackClassification) {
  const raw = _norm(allowedCell);
  let list = [];
  if (raw) {
    list = raw.split(',').map(s => _norm(s)).filter(Boolean);
  }
  if (list.length === 0 && _norm(fallbackClassification)) {
    list = [_norm(fallbackClassification)];
  }
  // de-dupe (case-insensitive) but preserve display form
  const seen = new Set();
  const out = [];
  list.forEach(p => {
    const k = p.toLowerCase();
    if (!seen.has(k)) { seen.add(k); out.push(p); }
  });
  return out;
}

// =============================
// CAMPUS SHEET HEADERS
// =============================
const MASTER_HEADERS = [
  'Timestamp','Student ID','Email','Surname','First Name','Middle Name',
  'Address','Mobile','Tel','Date of Birth','Sex','Civil Status','Spouse',
  'Emergency Contact Person','Emergency Contact Number','Facebook','Are You From AG',
  'Church Name','Church Address','Pastor Name','Ministry in Church',
  'Religious Affiliation','Recommendation','School Last Attended',
  'New Student','Classification','Subjects Enrolled','Campus','Password','Profile Picture URL',
  'Allowed Programs'  // NEW column 31
];

// =============================
// CAMPUS SHEET AUTO-SYNC
// =============================
function _ensureCampusSheets_() {
  const campuses = ['TWI-QATAR', 'TWI-CANADA', 'TWI-EUROPE'];
  campuses.forEach(c => _getOrCreateSheet(c, MASTER_HEADERS));
}

function syncCampusSheets() {
  try {
    _ensureCampusSheets_();
    const ss = _ss();
    const master = ss.getSheetByName('Master');
    const mData = master.getDataRange().getValues();

    const campusMap = {
      'TWI-QATAR':  ss.getSheetByName('TWI-QATAR'),
      'TWI-CANADA': ss.getSheetByName('TWI-CANADA'),
      'TWI-EUROPE': ss.getSheetByName('TWI-EUROPE')
    };

    const existingSets = {};
    Object.keys(campusMap).forEach(key => {
      const sh = campusMap[key];
      const d = sh.getDataRange().getValues();
      existingSets[key] = new Set();
      for (let i = 1; i < d.length; i++) {
        const sid = _norm(d[i][1]);
        if (sid) existingSets[key].add(sid);
      }
    });

    for (let i = 1; i < mData.length; i++) {
      const campus = _norm(mData[i][27]).toUpperCase();
      if (!campusMap[campus]) continue;
      const sid = _norm(mData[i][1]);
      if (!sid) continue;
      if (!existingSets[campus].has(sid)) {
        campusMap[campus].appendRow(mData[i]);
        existingSets[campus].add(sid);
      } else {
        const sh = campusMap[campus];
        const shData = sh.getDataRange().getValues();
        for (let j = 1; j < shData.length; j++) {
          if (_norm(shData[j][1]) === sid) {
            sh.getRange(j + 1, 1, 1, mData[i].length).setValues([mData[i]]);
            break;
          }
        }
      }
    }
    return { success: true, message: 'Campus sheets synced successfully.' };
  } catch (err) {
    return { success: false, message: 'Sync error: ' + err.message };
  }
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

function _getCurrentSemesterValue_() {
  setupInstructorModule();
  const ss = _ss();
  const settings = ss.getSheetByName('Settings');
  if (settings) {
    const setData = settings.getDataRange().getValues();
    for (let i = 1; i < setData.length; i++) {
      if (_norm(setData[i][0]) === 'Current Semester') {
        const overrideVal = _norm(setData[i][1]);
        if (overrideVal) return overrideVal;
      }
    }
  }
  const subj = ss.getSheetByName('Subjects');
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
  return { currentSemester: _getCurrentSemesterValue_() };
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

    // Default Allowed Programs = the chosen classification (admin can widen later).
    const allowedPrograms = _norm(formData.allowedPrograms) || _norm(formData.classification);

    const rowData = [
      timestamp, studentID, formData.email,
      formData.surname, formData.firstName, formData.middleName,
      formData.address, formData.mobile, formData.tel,
      formData.dateOfBirth, formData.sex, formData.civilStatus,
      formData.spouse, formData.emergencyContactPerson, formData.emergencyContactNumber,
      formData.facebook, formData.areYouFromAG, formData.churchName,
      formData.churchAddress, formData.pastorName, formData.ministryInChurch,
      formData.religiousAffiliation, formData.recommendation, formData.schoolLastAttended,
      formData.newStudent, formData.classification, formData.subjectsEnrolled,
      formData.campus, formData.password, '',
      allowedPrograms  // NEW column 31
    ];

    masterSheet.appendRow(rowData);

    let classSheet = null;
    if (formData.classification === '1st Year CCM') classSheet = ss.getSheetByName('1st Year CCM');
    else if (formData.classification === '2nd Year CCM') classSheet = ss.getSheetByName('2nd Year CCM');
    else if (formData.classification === 'CCM Evening Class') classSheet = ss.getSheetByName('Evening Class');
    else if (formData.classification === 'BCM') classSheet = ss.getSheetByName('BCM');
    if (classSheet) classSheet.appendRow(rowData);

    const campus = _norm(formData.campus).toUpperCase();
    if (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE') {
      const campusSheet = ss.getSheetByName(campus);
      if (campusSheet) campusSheet.appendRow(rowData);
    }

    sendWelcomeEmail(String(formData.email||'').toLowerCase().trim(), studentID, formData.password, formData.firstName);
    return { success: true, studentID, message: 'Registration successful! Check your email for login credentials.' };
  } catch (error) {
    return { success: false, message: 'Registration failed: ' + error.message };
  }
}

function sendWelcomeEmail(email, studentID, password, firstName) {
  const subject = 'Welcome to Thy Word Intl Bible College Bataan';
  const body =
    'Dear ' + firstName + ',\n\n' +
    'Welcome to Thy Word Intl Bible College Bataan!\n\n' +
    'Your registration has been successfully completed. Here are your login credentials for the Student Portal:\n\n' +
    'Student ID (Username): ' + studentID + '\n' +
    'Password: ' + password + '\n\n' +
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
    const trimmedPassword = String(password).trim();
    for (let i = 1; i < data.length; i++) {
      const sheetStudentID = String(data[i][1]).trim();
      const sheetPassword = String(data[i][28]).trim();
      if (sheetStudentID === trimmedStudentID && sheetPassword === trimmedPassword) {
        return { success: true, studentID: trimmedStudentID };
      }
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
      const password = String(data[i][28]).trim();
      const firstName = data[i][4] || '';
      if (sheetStudentID === trimmedStudentID && sheetEmail === trimmedEmail) {
        const subject = 'Your Student Portal Password - Thy Word Intl Bible College Bataan';
        const body =
          'Dear ' + firstName + ',\n\n' +
          'As requested, here are your login credentials for the Student Portal:\n\n' +
          'Student ID (Username): ' + sheetStudentID + '\n' +
          'Password: ' + password + '\n\n' +
          'God bless you!\n\n' +
          'Thy Word Intl Bible College Bataan Administration';
        MailApp.sendEmail(sheetEmail, subject, body);
        return { success: true, message: 'Your password has been sent to your email address.' };
      }
    }
    return { success: false, message: 'Student ID and Email do not match our records.' };
  } catch (error) {
    return { success: false, message: 'Error sending password: ' + error.message };
  }
}

// =============================
// GET STUDENT DATA
// =============================
function getStudentData(studentId) {
  const ss = _ss();
  const sheet = ss.getSheetByName('Master');
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange('B2:AE' + lastRow).getValues(); // extended to AE for Allowed Programs
  const headers = [
    'Student ID','Email','Surname','First Name','Middle Name',
    'Address','Mobile','Tel','Date of Birth','Sex','Civil Status','Spouse',
    'Emergency Contact Person','Emergency Contact Number','Facebook','Are You From AG',
    'Church Name','Church Address','Pastor Name','Ministry in Church',
    'Religious Affiliation','Recommendation','School Last Attended',
    'New Student','Classification','Subjects Enrolled','Campus',
    'Password','Profile Picture URL','Allowed Programs'
  ];
  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    if (id === String(studentId).trim()) {
      let studentData = {};
      headers.forEach((h, j) => studentData[h] = (data[i][j] !== undefined ? data[i][j] : ''));
      // Provide a resolved allowed-programs list for the front end.
      studentData['Allowed Programs List'] =
        _parseAllowedPrograms_(studentData['Allowed Programs'], studentData['Classification']);
      return JSON.stringify(studentData);
    }
  }
  return null;
}

// =============================
// UPDATE STUDENT PROFILE
// =============================
function updateStudentProfile(studentID, updates) {
  try {
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

      const oldClass = currentClass;

      if (updates.address)              masterSheet.getRange(i + 1, 7).setValue(updates.address);
      if (updates.mobile)               masterSheet.getRange(i + 1, 8).setValue(updates.mobile);
      if (updates.civilStatus)          masterSheet.getRange(i + 1, 12).setValue(updates.civilStatus);
      if (updates.spouse !== undefined) masterSheet.getRange(i + 1, 13).setValue(updates.spouse);
      if (newClass && newClass !== oldClass) {
        masterSheet.getRange(i + 1, 26).setValue(newClass);
      }

      const refreshed = masterSheet.getRange(i + 1, 1, 1, MASTER_HEADERS.length).getValues()[0];

      if (newClass && newClass !== oldClass) {
        const classSheetMap = {
          '1st Year CCM': '1st Year CCM',
          '2nd Year CCM': '2nd Year CCM',
          'CCM Evening Class': 'Evening Class',
          'BCM': 'BCM'
        };
        const oldSheetName = classSheetMap[oldClass];
        if (oldSheetName) {
          const oldSh = ss.getSheetByName(oldSheetName);
          if (oldSh) {
            const od = oldSh.getDataRange().getValues();
            for (let j = 1; j < od.length; j++) {
              if (_norm(od[j][1]) === studentID) { oldSh.deleteRow(j + 1); break; }
            }
          }
        }
        const newSheetName = classSheetMap[newClass];
        if (newSheetName) {
          const newSh = ss.getSheetByName(newSheetName);
          if (newSh) newSh.appendRow(refreshed);
        }
      } else {
        const classSheetMap = {
          '1st Year CCM': '1st Year CCM',
          '2nd Year CCM': '2nd Year CCM',
          'CCM Evening Class': 'Evening Class',
          'BCM': 'BCM'
        };
        const shName = classSheetMap[currentClass];
        if (shName) {
          const sh = ss.getSheetByName(shName);
          if (sh) {
            const sd = sh.getDataRange().getValues();
            for (let j = 1; j < sd.length; j++) {
              if (_norm(sd[j][1]) === studentID) {
                sh.getRange(j + 1, 1, 1, refreshed.length).setValues([refreshed]); break;
              }
            }
          }
        }
      }

      const campus = _norm(refreshed[27]).toUpperCase();
      if (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE') {
        _ensureCampusSheets_();
        const campusSh = ss.getSheetByName(campus);
        if (campusSh) {
          const cd = campusSh.getDataRange().getValues();
          let found = false;
          for (let j = 1; j < cd.length; j++) {
            if (_norm(cd[j][1]) === studentID) {
              campusSh.getRange(j + 1, 1, 1, refreshed.length).setValues([refreshed]);
              found = true; break;
            }
          }
          if (!found) campusSh.appendRow(refreshed);
        }
      }

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
function changePassword(studentID, oldPassword, newPassword) {
  try {
    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const data = masterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(studentID).trim()) {
        if (String(data[i][28]).trim() === String(oldPassword).trim()) {
          masterSheet.getRange(i + 1, 29).setValue(newPassword);
          const classification = data[i][25];
          let classSheet = null;
          if (classification === '1st Year CCM') classSheet = ss.getSheetByName('1st Year CCM');
          else if (classification === '2nd Year CCM') classSheet = ss.getSheetByName('2nd Year CCM');
          else if (classification === 'CCM Evening Class') classSheet = ss.getSheetByName('Evening Class');
          else if (classification === 'BCM') classSheet = ss.getSheetByName('BCM');
          if (classSheet) {
            const classData = classSheet.getDataRange().getValues();
            for (let j = 1; j < classData.length; j++) {
              if (String(classData[j][1]).trim() === String(studentID).trim()) {
                classSheet.getRange(j + 1, 29).setValue(newPassword); break;
              }
            }
          }
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
function uploadProfilePicture(studentID, imageData) {
  try {
    const matches = imageData.match(/^data:(.+);base64,(.+)$/);
    if (!matches) throw new Error("Invalid image data format");
    const contentType = matches[1];
    const base64Data = matches[2];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, studentID + "_profile.jpg");
    const folderName = "StudentProfilePictures";
    const folders = DriveApp.getFoldersByName(folderName);
    const driveFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const file = driveFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
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
        const directImageUrl = 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w400';
        return { success: true, message: 'Profile picture uploaded successfully', imageUrl: directImageUrl };
      }
    }
    return { success: false, message: 'Student not found' };
  } catch (error) {
    return { success: false, message: 'Error uploading picture: ' + error.message };
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
// SUBJECTS SHEET COLUMN RESOLUTION  (NEW)
// Layout: Subject, Day, From, To, Instructor, Semester, Year Level, Campus, Program
// We auto-detect the Program column by header, falling back to index 8.
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
  return {
    subject:    find(['subject', 'course', 'title'], 0),
    day:        find(['day'], 1),
    from:       find(['from', 'start'], 2),
    to:         find(['to', 'end'], 3),
    instructor: find(['instructor', 'faculty', 'teacher'], 4),
    semester:   find(['semester', 'sem'], 5),
    yearLevel:  find(['year level', 'year', 'level', 'period'], 6),
    campus:     find(['campus'], 7),
    program:    find(['program'], 8)
  };
}

// Pick the correct subjects sheet for a campus.
function _getSubjectsSheetForCampus_(campusUpper) {
  const ss = _ss();
  const isInt = (campusUpper === 'TWI-QATAR' || campusUpper === 'TWI-CANADA' || campusUpper === 'TWI-EUROPE');
  if (isInt) {
    return { sheet: ss.getSheetByName('INT Subjects'), isInt: true };
  }
  return { sheet: ss.getSheetByName('Subjects'), isInt: false };
}

// =============================
// GET AVAILABLE SUBJECTS  (REWRITTEN for multi-program / INT sheet)
// =============================
function getAvailableSubjects(studentCampus, studentClassification, allowedProgramsArg) {
  try {
    const ss = _ss();
    const campus = _norm(studentCampus).toUpperCase();
    const { sheet: subjectsSheet, isInt } = _getSubjectsSheetForCampus_(campus);

    if (!subjectsSheet) {
      return { success: false, message: (isInt ? "'INT Subjects'" : "'Subjects'") + ' sheet not found' };
    }

    const data = subjectsSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, subjects: [] };

    const cols = _resolveSubjectColumns_(subjectsSheet);

    // Resolve which programs this student may enroll in (PH only).
    let allowedList = Array.isArray(allowedProgramsArg)
      ? allowedProgramsArg.map(_norm).filter(Boolean)
      : _parseAllowedPrograms_(allowedProgramsArg, studentClassification);
    // Normalize to families so "1st Year CCM" / "2nd Year CCM" both map to CCM, etc.
    const allowedFamilies = new Set(allowedList.map(p => _programFamily_(p)));

    const subjects = [];
    for (let i = 1; i < data.length; i++) {
      const subjName = _norm(data[i][cols.subject]);
      if (!subjName) continue;

      const subjCampus  = _norm(data[i][cols.campus]).toUpperCase();
      const subjProgram = _norm(data[i][cols.program]);
      const subjYear    = _norm(data[i][cols.yearLevel]);

      if (isInt) {
        // INT students: subject must be for their campus (or ALL/blank). No program/cap filtering.
        const campusMatch = (subjCampus === campus || subjCampus === 'ALL' || subjCampus === '');
        if (!campusMatch) continue;
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
          allowed:    true
        });
      } else {
        // PH students: subject must belong to TWI-PHILIPPINES (or ALL/blank).
        const campusMatch = (subjCampus === 'TWI-PHILIPPINES' || subjCampus === 'ALL' || subjCampus === '');
        if (!campusMatch) continue;

        // A subject is "allowed" if its program family is in the student's allowed families.
        // If the student has no allowed programs resolved, fall back to allowing all PH subjects.
        const subjFamily = _programFamily_(subjProgram);
        const allowed = allowedFamilies.size === 0 || allowedFamilies.has(subjFamily);

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
          allowed:    allowed
        });
      }
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
// =============================
function getMilestoneData(studentID) {
  try {
    setupInstructorModule();
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

    const ccmSubjects = [];
    const bcmSubjects = [];
    const intSubjects = [];

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

    // Available-now set comes from the correct subjects sheet per campus.
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

    const gradesSheet = ss.getSheetByName('Grades');
    const completedSubjects = new Set();
    if (gradesSheet) {
      const gData = gradesSheet.getDataRange().getValues();
      for (let g = 1; g < gData.length; g++) {
        const sid   = _norm(gData[g][3]);
        const title = _norm(gData[g][2]);
        const grade = _norm(gData[g][6]);
        if (sid === _norm(studentID) && grade) completedSubjects.add(title.toLowerCase());
      }
    }

    const enrollSheet = ss.getSheetByName('Enrollments');
    const enrolledSubjects = new Set();
    if (enrollSheet) {
      const eData = enrollSheet.getDataRange().getValues();
      for (let e = 1; e < eData.length; e++) {
        const sid    = _norm(eData[e][2]);
        const title  = _norm(eData[e][4]);
        const status = _norm(eData[e][7]).toUpperCase();
        if (sid === _norm(studentID) && status === 'ENROLLED') {
          if (!completedSubjects.has(title.toLowerCase())) {
            enrolledSubjects.add(title.toLowerCase());
          }
        }
      }
    }

    function buildTrack(subjects) {
      return subjects.map(s => {
        const tl = s.title.toLowerCase();
        let status = 'grey';
        if (completedSubjects.has(tl)) status = 'red';
        else if (enrolledSubjects.has(tl)) status = 'orange';
        else if (availableNow.has(tl)) status = 'green';
        return { ...s, status };
      });
    }

    const intBuilt = buildTrack(intSubjects);
    const intByYear = [];
    for (let y = 0; y < 3; y++) {
      intByYear.push(intBuilt.slice(y * 8, y * 8 + 8));
    }

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
      success: true,
      isInternational,
      campus,
      ccm: { subjects: ccmTrack, progress: countProgress(ccmTrack) },
      bcm: { subjects: bcmTrack, progress: countProgress(bcmTrack) },
      int: { byYear: intByYear, progress: countProgress(intBuilt) }
    };
  } catch (err) {
    return { success: false, message: 'Milestone error: ' + err.message };
  }
}

// =============================
// ENROLLMENT SUMMARY  (REWRITTEN: per-program caps)
// =============================
function _getCurrentSemEnrollmentSummary_(studentID) {
  setupInstructorModule();
  const ss = _ss();
  const sem = _getCurrentSemesterValue_() || '';
  const enroll = ss.getSheetByName('Enrollments');
  const data = enroll.getDataRange().getValues();
  const { map: catalog } = _getCourseCatalogIndex_();

  const master = ss.getSheetByName('Master');
  const mData  = master.getDataRange().getValues();
  let classification = '';
  let allowedCell = '';
  let campus = '';
  for (let i = 1; i < mData.length; i++) {
    if (_norm(mData[i][1]) === _norm(studentID)) {
      classification = _norm(mData[i][25]);
      campus         = _norm(mData[i][27]).toUpperCase();
      allowedCell    = _norm(mData[i][30]); // Allowed Programs (col 31)
      break;
    }
  }

  const isInt = (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE');

  // Build a lookup: subject name -> program family, using the right subjects sheet.
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

  const sid = _norm(studentID);
  const subjects = [];
  let totalUnits = 0;

  // Per-program tallies (PH only).
  const perProgram = {}; // family -> { subjectCount, units }

  for (let i = 1; i < data.length; i++) {
    const rowSem = _norm(data[i][1]);
    const rowSid = _norm(data[i][2]);
    const rowSub = _norm(data[i][4]);
    const status = _norm(data[i][7]);
    if (rowSid !== sid) continue;
    if (sem && rowSem !== sem) continue;
    if (status && status.toUpperCase() !== 'ENROLLED') continue;
    if (!rowSub) continue;
    subjects.push(rowSub);
    const info = catalog.get(rowSub.toLowerCase());
    const units = info ? (Number(info.units) || 0) : 3;
    totalUnits += units;

    const fam = subjProgramMap.get(rowSub.toLowerCase()) || _programFamily_(classification);
    if (!perProgram[fam]) perProgram[fam] = { subjectCount: 0, units: 0 };
    perProgram[fam].subjectCount += 1;
    perProgram[fam].units += units;
  }

  const subjectCount = subjects.length;

  // INT students: no caps.
  if (isInt) {
    return {
      isInt: true,
      semester: sem,
      subjectCount,
      totalUnits,
      subjects,
      perProgram,
      maxReached: false,
      classification,
      // legacy fields kept for popup compatibility
      reached3: false, reached6: false, reached12: false, showPopup: false,
      maxSubjects: 9999, maxUnits: 9999
    };
  }

  // PH students: evaluate caps per program family.
  const programStatus = {};
  let anyCapReached = false;
  Object.keys(perProgram).forEach(fam => {
    const caps = _getProgramCaps_(fam === 'CCM Evening Class' ? 'CCM Evening Class' : fam);
    const pc = perProgram[fam];
    const reached = (pc.subjectCount >= caps.maxSubjects) || (pc.units >= caps.maxUnits);
    programStatus[fam] = {
      subjectCount: pc.subjectCount,
      units: pc.units,
      maxSubjects: caps.maxSubjects,
      maxUnits: caps.maxUnits,
      maxReached: reached
    };
    if (reached) anyCapReached = true;
  });

  // For the popup we use the dominant single-program situation when only one program is active.
  const activeFamilies = Object.keys(perProgram);
  const isEveningOnly = (activeFamilies.length === 1 && activeFamilies[0] === 'CCM Evening Class');
  const popupCaps = _getProgramCaps_(isEveningOnly ? 'CCM Evening Class' : classification);
  const reached3  = totalUnits === 3;
  const reached6  = totalUnits === 6;
  const reached12 = !isEveningOnly && totalUnits === 12;
  const showPopup = reached3 || reached6 || reached12;

  return {
    isInt: false,
    semester: sem,
    subjectCount,
    totalUnits,
    subjects,
    perProgram: programStatus,
    reached3,
    reached6,
    reached12,
    showPopup,
    maxReached: anyCapReached, // any program full → overall locked notice (per-program enforced on enroll)
    maxSubjects: popupCaps.maxSubjects,
    maxUnits: popupCaps.maxUnits,
    classification
  };
}

function getCurrentSemEnrollmentSummary(studentID) {
  return _getCurrentSemEnrollmentSummary_(studentID);
}

// =============================
// ENROLL IN SUBJECT  (REWRITTEN: multi-program + per-program caps + INT sheet)
// =============================
function enrollInSubject(studentID, subjectName) {
  try {
    setupInstructorModule();

    const ss = _ss();
    const masterSheet = ss.getSheetByName('Master');
    const data = masterSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() !== String(studentID).trim()) continue;

      const studentClass = _norm(data[i][25]);
      const campus       = _norm(data[i][27]).toUpperCase();
      const allowedCell  = _norm(data[i][30]);
      const isInt = (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE');

      // Look up the subject's program family from the appropriate subjects sheet.
      const subjectInfo  = _lookupSubjectInfo_(subjectName, campus);
      const subjFamily   = _programFamily_(subjectInfo.program);

      // ----- Eligibility check (PH only) -----
      if (!isInt) {
        const allowedList = _parseAllowedPrograms_(allowedCell, studentClass);
        const allowedFamilies = new Set(allowedList.map(p => _programFamily_(p)));
        if (allowedFamilies.size > 0 && subjFamily && !allowedFamilies.has(subjFamily)) {
          return {
            success: false,
            message: 'You are not allowed to enroll in this subject. It belongs to a program you are not enrolled in.'
          };
        }
      }

      // ----- Per-program cap check (PH only) -----
      const summaryBefore = _getCurrentSemEnrollmentSummary_(studentID);
      if (!isInt) {
        const caps = _getProgramCaps_(subjFamily === 'CCM Evening Class' ? 'CCM Evening Class' : subjFamily);
        const ps = (summaryBefore.perProgram && summaryBefore.perProgram[subjFamily]) || { subjectCount: 0, units: 0 };
        const wouldBeSubjects = (ps.subjectCount || 0) + 1;
        const subjUnits = (subjectInfo.units && Number(subjectInfo.units)) || _catalogUnits_(subjectName) || 3;
        const wouldBeUnits = (ps.units || 0) + subjUnits;

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

      // ----- Already enrolled? -----
      let currentSubjects = data[i][26] || '';
      if (String(currentSubjects).split(',').map(s => s.trim().toLowerCase()).includes(String(subjectName).trim().toLowerCase())) {
        return { success: false, message: 'You are already enrolled in this subject' };
      }

      // ----- Write enrollment -----
      const updatedSubjects = currentSubjects ? (currentSubjects + ', ' + subjectName) : subjectName;
      masterSheet.getRange(i + 1, 27).setValue(updatedSubjects);

      const classification = data[i][25];
      let classSheet = null;
      if (classification === '1st Year CCM') classSheet = ss.getSheetByName('1st Year CCM');
      else if (classification === '2nd Year CCM') classSheet = ss.getSheetByName('2nd Year CCM');
      else if (classification === 'CCM Evening Class') classSheet = ss.getSheetByName('Evening Class');
      else if (classification === 'BCM') classSheet = ss.getSheetByName('BCM');
      if (classSheet) {
        const classData = classSheet.getDataRange().getValues();
        for (let j = 1; j < classData.length; j++) {
          if (String(classData[j][1]).trim() === String(studentID).trim()) {
            classSheet.getRange(j + 1, 27).setValue(updatedSubjects); break;
          }
        }
      }

      _upsertEnrollmentFromStudentRow_(data[i], subjectName, campus);
      const summaryAfter = _getCurrentSemEnrollmentSummary_(studentID);
      return {
        success: true,
        message: 'Successfully enrolled in ' + subjectName,
        enrolledSubjects: updatedSubjects,
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

// Look up subject info from the campus-appropriate subjects sheet (Subjects or INT Subjects).
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
  const semester    = info.semester || _getCurrentSemesterValue_() || '';
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
  setupInstructorModule();
  const ss = _ss();
  const sh = ss.getSheetByName('Instructors');
  const data = sh.getDataRange().getValues();
  const u = String(username || '').trim();
  const p = String(password || '').trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === u && String(data[i][1]).trim() === p) {
      return { success: true, username: u, fullName: String(data[i][2] || u).trim() };
    }
  }
  return { success: false, message: 'Invalid instructor username or password.' };
}

function getInstructorSubjects(instructorUsername) {
  setupInstructorModule();
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
  const currentSem = _getCurrentSemesterValue_();

  // Build a combined subject map from BOTH the PH Subjects sheet and INT Subjects sheet,
  // so instructors who teach in either can see their handles.
  const subjMap = new Map();
  ['Subjects', 'INT Subjects'].forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const cols = _resolveSubjectColumns_(sh);
    const sData = sh.getDataRange().getValues();
    for (let r = 1; r < sData.length; r++) {
      const subName = _norm(sData[r][cols.subject]);
      if (!subName) continue;
      subjMap.set(subName.toLowerCase(), { row: sData[r], cols });
    }
  });

  const subjects = [];
  subjectHandles.forEach(subName => {
    const entry = subjMap.get(subName.toLowerCase());
    let day = '', from = '', to = '', semester = currentSem, yearLevel = '';
    if (entry) {
      const row = entry.row, cols = entry.cols;
      day = _norm(row[cols.day]);
      from = row[cols.from] ? formatTimeAsText(row[cols.from]) : '';
      to   = row[cols.to]   ? formatTimeAsText(row[cols.to])   : '';
      yearLevel = _norm(row[cols.yearLevel]);
      const semFromSheet = _norm(row[cols.semester]);
      semester = semFromSheet || currentSem;
    }
    if (currentSem && semester && semester !== currentSem) return;
    subjects.push({ subject: subName, day, from, to, instructor: fullName, semester: semester || currentSem || '', yearLevel });
  });
  return { success: true, subjects, currentSemester: currentSem, fullName };
}

function getEnrolledStudentsForSubject(instructorUsername, subjectName, semester) {
  setupInstructorModule();
  const ss = _ss();
  const instSheet = ss.getSheetByName('Instructors');
  const instData  = instSheet.getDataRange().getValues();
  let fullName = '';
  for (let i = 1; i < instData.length; i++) {
    if (_norm(instData[i][0]) === _norm(instructorUsername)) { fullName = _norm(instData[i][2]); break; }
  }
  if (!fullName) return { success: false, message: 'Instructor not found.' };
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
    if (rowSem !== sem) continue;
    if (rowSub !== sub) continue;
    if (rowInst !== fullName) continue;
    const key      = [sem, sub, rowSid].join('|');
    const existing = gradeMap.get(key) || { grade:'', remarks:'', locked:false };
    students.push({ studentId: rowSid, studentName: rowName, grade: existing.grade, remarks: existing.remarks, locked: !!existing.locked });
  }
  return { success: true, students, instructor: fullName };
}

function saveGrades(instructorUsername, semester, subjectName, gradesArray) {
  setupInstructorModule();
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
    idx.set(key, { rowNum: i + 1, instructor: _norm(data[i][5]) });
  }
  let saved = 0, lockedCount = 0;
  const results = [];
  gradesArray = gradesArray || [];
  gradesArray.forEach(g => {
    const sid     = _norm(g.studentId);
    const sname   = _norm(g.studentName);
    const grade   = _norm(g.grade);
    const remarks = _norm(g.remarks);
    if (!sid) { results.push({ studentId: '', status: 'skipped', reason: 'Missing Student ID' }); return; }
    const key   = [sem, sub.toLowerCase(), sid].join('|');
    const found = idx.get(key);
    if (found && found.instructor === fullName) { lockedCount++; results.push({ studentId: sid, status: 'locked' }); return; }
    if (found && found.instructor && found.instructor !== fullName) { lockedCount++; results.push({ studentId: sid, status: 'locked', reason: 'Already graded by another instructor' }); return; }
    gradesSheet.appendRow([new Date(), sem, sub, sid, sname, fullName, grade, remarks]);
    saved++;
    results.push({ studentId: sid, status: 'saved' });
  });
  let msg = 'Saved grades for ' + saved + ' student(s).';
  if (lockedCount) msg += ' Locked: ' + lockedCount + ' (already saved before).';
  return { success: true, message: msg, results };
}

// =============================
// CERTIFICATE OF REGISTRATION
// =============================
function createAndSendCertificateOfRegistration(studentID) {
  try {
    setupInstructorModule();
    const ss = _ss();
    const master = ss.getSheetByName('Master');
    if (!master) throw new Error('Master sheet not found.');
    const sem = _getCurrentSemesterValue_() || '';
    if (!sem) throw new Error('Current Semester is not set. Please set it in Settings sheet (Current Semester).');
    const mData = master.getDataRange().getValues();
    let row = null;
    for (let i = 1; i < mData.length; i++) {
      if (_norm(mData[i][1]) === _norm(studentID)) { row = mData[i]; break; }
    }
    if (!row) throw new Error('Student not found in Master sheet.');

    const email           = _norm(row[2]).toLowerCase();
    const first           = _norm(row[4]);
    const address         = _norm(row[6]);
    const sex             = _norm(row[10]);
    const classification  = _norm(row[25]);
    const campus          = _norm(row[27]).toUpperCase();
    const isInternational = (campus === 'TWI-QATAR' || campus === 'TWI-CANADA' || campus === 'TWI-EUROPE');
    const studentName     = _buildStudentName_(row);
    const programType     = (classification.toUpperCase().includes('BCM')) ? 'BCM' : 'CCM';

    const enroll  = ss.getSheetByName('Enrollments');
    const eData   = enroll.getDataRange().getValues();
    const { map: catalog } = _getCourseCatalogIndex_();
    const subjects = [];
    let totalUnits = 0;
    // COR counts ALL enrolled subjects for the semester, across every program.
    for (let i = 1; i < eData.length; i++) {
      const rowSem    = _norm(eData[i][1]);
      const rowSid    = _norm(eData[i][2]);
      const subj      = _norm(eData[i][4]);
      const instr     = _norm(eData[i][5]);
      const yearLevel = _norm(eData[i][6]);
      const status    = _norm(eData[i][7]);
      if (rowSid !== _norm(studentID)) continue;
      if (rowSem !== sem) continue;
      if (status && status.toUpperCase() !== 'ENROLLED') continue;
      if (!subj) continue;
      const info  = catalog.get(subj.toLowerCase()) || { code:'', units:3, title:subj };
      const units = Number(info.units) || 3;
      totalUnits += units;
      subjects.push({ code: info.code || '', title: subj, units, instructor: instr, yearLevel });
    }
    if (subjects.length === 0) throw new Error('No enrolled subjects found for the current semester.');

    const fees = _computeFeesByUnits_(totalUnits, isInternational);

    const html = _buildCORHtml_({
      semester:       sem,
      programType,
      classification,
      campus,
      studentID:      _norm(studentID),
      studentName,
      address,
      sex,
      email,
      subjects,
      totals: { totalUnits, subjectCount: subjects.length },
      fees
    });

    const blob = HtmlService.createHtmlOutput(html).setSandboxMode(HtmlService.SandboxMode.IFRAME).getBlob().setName(`COR_${_norm(studentID)}_${_safeFile_(sem)}.pdf`);
    const pdf  = blob.getAs(MimeType.PDF);
    const folder = _getOrCreateCorFolder_(sem, programType);
    const file   = folder.createFile(pdf);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const subject = `Certificate of Registration - ${sem}`;
    const body =
      `Dear ${first || 'Student'},\n\n` +
      `Attached is your Certificate of Registration (COR) for ${sem}.\n\n` +
      `Student ID: ${studentID}\nName: ${studentName}\nCampus: ${campus}\nProgram: ${classification}\n\n` +
      `You may also access your COR here:\n${file.getUrl()}\n\nGod bless you!\n\nThy Word Intl Bible College Bataan`;
    if (email) {
      MailApp.sendEmail({ to: email, subject, body, attachments: [pdf] });
    } else {
      throw new Error('Student email is empty in Master sheet.');
    }
    return { success: true, message: 'COR generated and sent to student email successfully.', fileUrl: file.getUrl(), fileId: file.getId(), semester: sem, programType, totalUnits, subjectCount: subjects.length };
  } catch (err) {
    return { success: false, message: 'COR error: ' + err.message };
  }
}

// =============================
// COMPUTE FEES BY UNITS
// =============================
function _computeFeesByUnits_(totalUnits, isInternational) {
  if (isInternational) {
    const intReg     = 1000;
    const intMisc    = 1000;
    const intTuition = (totalUnits / 3) * 2000;
    const intTotal   = intReg + intMisc + intTuition;
    return { registrationFee: intReg, miscellaneousFee: intMisc, tuitionFee: intTuition, totalAssessment: intTotal };
  }

  const reg = 400, misc = 550;
  let tuition = 0, total = 0;
  if (totalUnits === 3)       { tuition = 1050; total = 2000; }
  else if (totalUnits === 6)  { tuition = 3050; total = 4000; }
  else if (totalUnits === 12) { tuition = 5050; total = 6000; }
  else { tuition = totalUnits * 350; total = tuition + reg + misc; }
  return { registrationFee: reg, miscellaneousFee: misc, tuitionFee: tuition, totalAssessment: total };
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
      const rowEmail  = String(data[i][2] || '').trim().toLowerCase();
      const rowSurname= String(data[i][3] || '').trim().toUpperCase();
      const rowFirst  = String(data[i][4] || '').trim().toUpperCase();
      const rowMiddle = String(data[i][5] || '').trim().toUpperCase();
      if (rowEmail === e && rowSurname === sn && rowFirst === fn && rowMiddle === mn) {
        return { exists: true, studentID: String(data[i][1]).trim() };
      }
    }
    return { exists: false };
  } catch (err) {
    return { exists: false };
  }
}

function _getOrCreateCorFolder_(semester, programType) {
  const ssFolderName = 'TWI_COR';
  const rootIt = DriveApp.getFoldersByName(ssFolderName);
  const root   = rootIt.hasNext() ? rootIt.next() : DriveApp.createFolder(ssFolderName);
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
  return '₱' + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
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