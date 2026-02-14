/**
 * Clinic SaaS Queue System - Google Apps Script Backend
 * Beginner-friendly, modular implementation.
 */

const SHEETS = {
  QUEUE: 'Queue',
  PATIENTS: 'Patients',
  APPOINTMENTS: 'Appointments',
  USERS: 'Users',
  ACTIVITY: 'ActivityLogs',
  SETTINGS: 'Settings',
  REMINDERS: 'Reminders'
};

const STATUS = {
  WAITING: 'WAITING',
  SERVING: 'SERVING',
  COMPLETED: 'COMPLETED',
  SKIPPED: 'SKIPPED'
};

const TIMEZONE = Session.getScriptTimeZone() || 'Asia/Kolkata';

function doGet(e) {
  initializeSystem();
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  const allowedPages = ['index', 'login', 'dashboard', 'appointments', 'queue', 'display', 'patients', 'reminders', 'reports', 'admin'];
  const file = allowedPages.indexOf(page) > -1 ? page : 'index';
  return HtmlService.createTemplateFromFile(file)
    .evaluate()
    .setTitle(getSettingsObject().clinicName || 'Clinic Queue System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetIfMissing_(ss, SHEETS.QUEUE, [
    'Date', 'Token No', 'Patient Name', 'Phone', 'Appointment Time', 'Doctor',
    'Status', 'Check-in Time', 'Serve Time', 'Complete Time', 'Notes'
  ]);

  createSheetIfMissing_(ss, SHEETS.PATIENTS, [
    'Patient ID', 'Name', 'Phone', 'Visit Count', 'Last Visit Date', 'Created At', 'Updated At'
  ]);

  createSheetIfMissing_(ss, SHEETS.APPOINTMENTS, [
    'Appointment ID', 'Date', 'Token No', 'Patient Name', 'Phone', 'Appointment Time',
    'Doctor', 'Status', 'Created At', 'Updated At'
  ]);

  createSheetIfMissing_(ss, SHEETS.USERS, [
    'Email', 'Name', 'Role', 'Password', 'Is Active', 'Created At'
  ]);

  createSheetIfMissing_(ss, SHEETS.ACTIVITY, [
    'Timestamp', 'Actor', 'Action', 'Entity', 'Entity ID', 'Details'
  ]);

  createSheetIfMissing_(ss, SHEETS.SETTINGS, [
    'Key', 'Value'
  ]);

  createSheetIfMissing_(ss, SHEETS.REMINDERS, [
    'Timestamp', 'Patient Name', 'Phone', 'Appointment ID', 'Message', 'Status', 'Triggered By'
  ]);

  seedDefaults_();
}

function createSheetIfMissing_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  }
}

function seedDefaults_() {
  const settings = getSheet_(SHEETS.SETTINGS);
  if (settings.getLastRow() < 2) {
    const defaults = [
      ['clinicName', 'Sunrise Clinic'],
      ['clinicLogo', 'https://via.placeholder.com/120x120.png?text=Clinic'],
      ['welcomeMessage', 'Welcome to our digital appointment and token queue system.'],
      ['primaryColor', '#0F766E'],
      ['supportPhone', '+91-99999-99999']
    ];
    settings.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }

  const users = getSheet_(SHEETS.USERS);
  if (users.getLastRow() < 2) {
    const now = now_();
    users.getRange(2, 1, 3, 6).setValues([
      ['admin@clinic.com', 'Clinic Admin', 'Admin', 'admin123', 'YES', now],
      ['reception@clinic.com', 'Reception Desk', 'Reception', 'recep123', 'YES', now],
      ['doctor@clinic.com', 'Duty Doctor', 'Doctor', 'doc123', 'YES', now]
    ]);
  }
}

function getSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Missing required sheet: ' + name);
  return sh;
}

function now_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function today_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
}

function normalizePhone_(phone) {
  return String(phone || '').replace(/\D/g, '');
}

function getSettingsObject() {
  initializeSystem();
  const sh = getSheet_(SHEETS.SETTINGS);
  const data = sh.getDataRange().getValues();
  const obj = {};
  for (let i = 1; i < data.length; i++) {
    obj[data[i][0]] = data[i][1];
  }
  return obj;
}

function updateSettings(entries, actor) {
  initializeSystem();
  const sh = getSheet_(SHEETS.SETTINGS);
  const values = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) map[values[i][0]] = i + 1;

  Object.keys(entries || {}).forEach(function(key) {
    if (map[key]) {
      sh.getRange(map[key], 2).setValue(entries[key]);
    } else {
      sh.appendRow([key, entries[key]]);
    }
  });

  logActivity_(actor || 'SYSTEM', 'UPDATE_SETTINGS', 'Settings', '-', JSON.stringify(entries || {}));
  return { success: true, settings: getSettingsObject() };
}

function loginUser(email, password) {
  initializeSystem();
  const sh = getSheet_(SHEETS.USERS);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[0]).toLowerCase() === String(email).toLowerCase() && row[3] === password && String(row[4]).toUpperCase() === 'YES') {
      const user = {
        email: row[0],
        name: row[1],
        role: row[2]
      };
      logActivity_(user.email, 'LOGIN', 'Users', user.email, 'Role: ' + user.role);
      return { success: true, user: user };
    }
  }

  return { success: false, message: 'Invalid credentials or inactive user.' };
}

function getUsers() {
  initializeSystem();
  const rows = getSheet_(SHEETS.USERS).getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return {
      email: r[0],
      name: r[1],
      role: r[2],
      password: r[3],
      isActive: r[4],
      createdAt: r[5]
    };
  });
}

function addOrUpdateUser(user, actor) {
  initializeSystem();
  if (!user || !user.email) throw new Error('User email is required');
  const sh = getSheet_(SHEETS.USERS);
  const data = sh.getDataRange().getValues();
  const email = String(user.email).toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      sh.getRange(i + 1, 2, 1, 5).setValues([[
        user.name || data[i][1],
        user.role || data[i][2],
        user.password || data[i][3],
        user.isActive || data[i][4],
        data[i][5] || now_()
      ]]);
      logActivity_(actor || 'SYSTEM', 'UPDATE_USER', 'Users', email, JSON.stringify(user));
      return { success: true, mode: 'updated' };
    }
  }

  sh.appendRow([
    email,
    user.name || '',
    user.role || 'Reception',
    user.password || 'changeme123',
    user.isActive || 'YES',
    now_()
  ]);
  logActivity_(actor || 'SYSTEM', 'ADD_USER', 'Users', email, JSON.stringify(user));
  return { success: true, mode: 'added' };
}

function addAppointment(payload) {
  initializeSystem();
  validateAppointmentPayload_(payload);

  const date = payload.date || today_();
  const phone = normalizePhone_(payload.phone);
  const patientName = String(payload.patientName || '').trim();
  const doctor = String(payload.doctor || '').trim();
  const appointmentTime = String(payload.appointmentTime || '').trim();
  const actor = payload.actor || 'SYSTEM';

  if (isDuplicateAppointment_(date, phone, doctor, appointmentTime)) {
    return { success: false, message: 'Duplicate appointment exists for same date, phone, doctor and time.' };
  }

  const token = generateToken(date);
  const appointmentId = 'APT-' + date.replace(/-/g, '') + '-' + token;

  const appointmentSheet = getSheet_(SHEETS.APPOINTMENTS);
  appointmentSheet.appendRow([
    appointmentId,
    date,
    token,
    patientName,
    phone,
    appointmentTime,
    doctor,
    STATUS.WAITING,
    now_(),
    now_()
  ]);

  const queueSheet = getSheet_(SHEETS.QUEUE);
  queueSheet.appendRow([
    date,
    token,
    patientName,
    phone,
    appointmentTime,
    doctor,
    STATUS.WAITING,
    now_(),
    '',
    '',
    ''
  ]);

  upsertPatient_(patientName, phone, date);
  logActivity_(actor, 'ADD_APPOINTMENT', 'Appointments', appointmentId, 'Token ' + token);

  return {
    success: true,
    appointmentId: appointmentId,
    tokenNo: token,
    status: STATUS.WAITING,
    message: 'Appointment created successfully.'
  };
}

function validateAppointmentPayload_(payload) {
  if (!payload) throw new Error('Payload required.');
  if (!payload.patientName) throw new Error('Patient name required.');
  if (!payload.phone) throw new Error('Phone required.');
  if (!payload.appointmentTime) throw new Error('Appointment time required.');
  if (!payload.doctor) throw new Error('Doctor required.');
}

function isDuplicateAppointment_(date, phone, doctor, time) {
  const sh = getSheet_(SHEETS.APPOINTMENTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === date && normalizePhone_(rows[i][4]) === phone && rows[i][6] === doctor && rows[i][5] === time && rows[i][7] !== STATUS.COMPLETED) {
      return true;
    }
  }
  return false;
}

function generateToken(date) {
  initializeSystem();
  const d = date || today_();
  const sh = getSheet_(SHEETS.QUEUE);
  const rows = sh.getDataRange().getValues();
  let maxToken = 0;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d) {
      maxToken = Math.max(maxToken, Number(rows[i][1]) || 0);
    }
  }
  return maxToken + 1;
}

function callNextPatient(actor) {
  initializeSystem();
  const sh = getSheet_(SHEETS.QUEUE);
  const rows = sh.getDataRange().getValues();
  const date = today_();

  // enforce single SERVING
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === date && rows[i][6] === STATUS.SERVING) {
      return { success: false, message: 'A patient is already being served.', tokenNo: rows[i][1] };
    }
  }

  let candidate = null;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === date && rows[i][6] === STATUS.WAITING) {
      if (!candidate || Number(rows[i][1]) < Number(candidate.tokenNo)) {
        candidate = { rowIndex: i + 1, tokenNo: rows[i][1], patientName: rows[i][2] };
      }
    }
  }

  if (!candidate) {
    return { success: false, message: 'No WAITING patients in queue.' };
  }

  sh.getRange(candidate.rowIndex, 7).setValue(STATUS.SERVING);
  sh.getRange(candidate.rowIndex, 9).setValue(now_());
  syncAppointmentStatus_(date, candidate.tokenNo, STATUS.SERVING);

  logActivity_(actor || 'SYSTEM', 'CALL_NEXT', 'Queue', String(candidate.tokenNo), candidate.patientName);

  return {
    success: true,
    message: 'Patient moved to SERVING.',
    tokenNo: candidate.tokenNo,
    patientName: candidate.patientName
  };
}

function completePatient(tokenNo, actor, notes) {
  return updateQueueStatus_(tokenNo, STATUS.SERVING, STATUS.COMPLETED, actor, notes || 'Completed consultation');
}

function skipPatient(tokenNo, actor, notes) {
  return updateQueueStatus_(tokenNo, STATUS.SERVING, STATUS.SKIPPED, actor, notes || 'Skipped by doctor/reception');
}

function updateQueueStatus_(tokenNo, fromStatus, toStatus, actor, notes) {
  initializeSystem();
  const sh = getSheet_(SHEETS.QUEUE);
  const rows = sh.getDataRange().getValues();
  const date = today_();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === date && Number(rows[i][1]) === Number(tokenNo) && rows[i][6] === fromStatus) {
      sh.getRange(i + 1, 7).setValue(toStatus);
      if (toStatus === STATUS.COMPLETED || toStatus === STATUS.SKIPPED) {
        sh.getRange(i + 1, 10).setValue(now_());
      }
      sh.getRange(i + 1, 11).setValue(notes || '');
      syncAppointmentStatus_(date, tokenNo, toStatus);
      logActivity_(actor || 'SYSTEM', toStatus, 'Queue', String(tokenNo), notes || '');
      return { success: true, tokenNo: tokenNo, newStatus: toStatus };
    }
  }
  return { success: false, message: 'Token not found in ' + fromStatus + ' state.' };
}

function syncAppointmentStatus_(date, tokenNo, status) {
  const sh = getSheet_(SHEETS.APPOINTMENTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === date && Number(rows[i][2]) === Number(tokenNo)) {
      sh.getRange(i + 1, 8).setValue(status);
      sh.getRange(i + 1, 10).setValue(now_());
      return;
    }
  }
}

function upsertPatient_(name, phone, visitDate) {
  const sh = getSheet_(SHEETS.PATIENTS);
  const rows = sh.getDataRange().getValues();
  const normalized = normalizePhone_(phone);

  for (let i = 1; i < rows.length; i++) {
    if (normalizePhone_(rows[i][2]) === normalized) {
      const count = Number(rows[i][3]) || 0;
      sh.getRange(i + 1, 1, 1, 7).setValues([[
        rows[i][0],
        name || rows[i][1],
        normalized,
        count + 1,
        visitDate,
        rows[i][5] || now_(),
        now_()
      ]]);
      return;
    }
  }

  const patientId = 'PAT-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  sh.appendRow([patientId, name, normalized, 1, visitDate, now_(), now_()]);
}

function getQueueDisplayData() {
  initializeSystem();
  const date = today_();
  const queue = getTodayQueue_();
  const serving = queue.filter(function(q) { return q.status === STATUS.SERVING; })[0] || null;
  const waiting = queue.filter(function(q) { return q.status === STATUS.WAITING; }).slice(0, 10);

  return {
    currentServing: serving,
    nextWaiting: waiting,
    lastUpdated: now_(),
    isEmpty: queue.length === 0
  };
}

function getTodayQueue_() {
  const date = today_();
  const rows = getSheet_(SHEETS.QUEUE).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === date) {
      out.push({
        rowNumber: i + 1,
        date: rows[i][0],
        tokenNo: rows[i][1],
        patientName: rows[i][2],
        phone: rows[i][3],
        appointmentTime: rows[i][4],
        doctor: rows[i][5],
        status: rows[i][6],
        checkInTime: rows[i][7],
        serveTime: rows[i][8],
        completeTime: rows[i][9],
        notes: rows[i][10]
      });
    }
  }
  out.sort(function(a, b) { return Number(a.tokenNo) - Number(b.tokenNo); });
  return out;
}

function getQueueData() {
  return getTodayQueue_();
}

function getAppointmentsByDate(date) {
  initializeSystem();
  const selected = date || today_();
  const rows = getSheet_(SHEETS.APPOINTMENTS).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === selected) {
      out.push({
        appointmentId: rows[i][0],
        date: rows[i][1],
        tokenNo: rows[i][2],
        patientName: rows[i][3],
        phone: rows[i][4],
        appointmentTime: rows[i][5],
        doctor: rows[i][6],
        status: rows[i][7],
        createdAt: rows[i][8],
        updatedAt: rows[i][9]
      });
    }
  }
  out.sort(function(a, b) { return Number(a.tokenNo) - Number(b.tokenNo); });
  return out;
}

function getPatients(searchText) {
  initializeSystem();
  const rows = getSheet_(SHEETS.PATIENTS).getDataRange().getValues();
  const q = String(searchText || '').toLowerCase().trim();
  const out = rows.slice(1).map(function(r) {
    return {
      patientId: r[0],
      name: r[1],
      phone: r[2],
      visitCount: r[3],
      lastVisitDate: r[4],
      createdAt: r[5],
      updatedAt: r[6]
    };
  });

  if (!q) return out;
  return out.filter(function(p) {
    return String(p.name).toLowerCase().indexOf(q) > -1 || String(p.phone).toLowerCase().indexOf(q) > -1;
  });
}

function getPatientVisitHistory(phone) {
  initializeSystem();
  const normalized = normalizePhone_(phone);
  const rows = getSheet_(SHEETS.APPOINTMENTS).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (normalizePhone_(rows[i][4]) === normalized) {
      out.push({
        appointmentId: rows[i][0],
        date: rows[i][1],
        tokenNo: rows[i][2],
        doctor: rows[i][6],
        status: rows[i][7]
      });
    }
  }
  return out;
}

function triggerReminder(payload) {
  initializeSystem();
  if (!payload || !payload.phone || !payload.patientName) {
    throw new Error('patientName and phone are required for reminder.');
  }

  const message = payload.message || ('Dear ' + payload.patientName + ', this is a reminder for your clinic appointment at ' + (payload.appointmentTime || 'scheduled time') + '.');
  const random = Math.random();
  let status = 'Sent';
  if (random < 0.1) status = 'Failed';
  else if (random < 0.25) status = 'Scheduled';

  getSheet_(SHEETS.REMINDERS).appendRow([
    now_(),
    payload.patientName,
    normalizePhone_(payload.phone),
    payload.appointmentId || '',
    message,
    status,
    payload.actor || 'SYSTEM'
  ]);

  logActivity_(payload.actor || 'SYSTEM', 'WHATSAPP_REMINDER', 'Reminders', payload.appointmentId || '-', status);

  return {
    success: true,
    status: status,
    provider: 'MockWhatsAppGateway',
    message: message
  };
}

function getReminders(limit) {
  initializeSystem();
  const rows = getSheet_(SHEETS.REMINDERS).getDataRange().getValues();
  const out = rows.slice(1).map(function(r) {
    return {
      timestamp: r[0],
      patientName: r[1],
      phone: r[2],
      appointmentId: r[3],
      message: r[4],
      status: r[5],
      triggeredBy: r[6]
    };
  });
  out.reverse();
  return out.slice(0, Number(limit) || 100);
}

function getDashboardStats() {
  initializeSystem();
  const queue = getTodayQueue_();
  const counts = {
    todaysAppointments: queue.length,
    activeQueue: queue.filter(function(q) { return q.status === STATUS.WAITING || q.status === STATUS.SERVING; }).length,
    completed: queue.filter(function(q) { return q.status === STATUS.COMPLETED; }).length,
    waiting: queue.filter(function(q) { return q.status === STATUS.WAITING; }).length,
    serving: queue.filter(function(q) { return q.status === STATUS.SERVING; }).length,
    skipped: queue.filter(function(q) { return q.status === STATUS.SKIPPED; }).length,
    liveQueueSummary: queue.slice(0, 5)
  };
  return counts;
}

function getReports(date) {
  initializeSystem();
  const selected = date || today_();
  const rows = getSheet_(SHEETS.QUEUE).getDataRange().getValues();
  const filtered = rows.slice(1).filter(function(r) { return r[0] === selected; });

  const summary = {
    date: selected,
    total: filtered.length,
    waiting: filtered.filter(function(r) { return r[6] === STATUS.WAITING; }).length,
    serving: filtered.filter(function(r) { return r[6] === STATUS.SERVING; }).length,
    completed: filtered.filter(function(r) { return r[6] === STATUS.COMPLETED; }).length,
    skipped: filtered.filter(function(r) { return r[6] === STATUS.SKIPPED; }).length
  };

  return {
    summary: summary,
    queue: filtered.map(function(r) {
      return {
        tokenNo: r[1],
        patientName: r[2],
        phone: r[3],
        time: r[4],
        doctor: r[5],
        status: r[6],
        checkInTime: r[7],
        serveTime: r[8],
        completeTime: r[9],
        notes: r[10]
      };
    })
  };
}

function getActivityLogs(limit) {
  initializeSystem();
  const rows = getSheet_(SHEETS.ACTIVITY).getDataRange().getValues();
  const out = rows.slice(1).map(function(r) {
    return {
      timestamp: r[0],
      actor: r[1],
      action: r[2],
      entity: r[3],
      entityId: r[4],
      details: r[5]
    };
  });
  out.reverse();
  return out.slice(0, Number(limit) || 200);
}

function logActivity_(actor, action, entity, entityId, details) {
  getSheet_(SHEETS.ACTIVITY).appendRow([
    now_(),
    actor || 'SYSTEM',
    action || '',
    entity || '',
    entityId || '',
    details || ''
  ]);
}
