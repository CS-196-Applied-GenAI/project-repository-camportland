/**
 * Camp Kesem Event Revenue Tracking System — Production
 *
 * Includes:
 * - Admin/Viewer permissions (Option A)
 * - Duplicate event prevention + locking
 * - Processed Attendance column on Events
 * - Camp Kesem bold theme formatting (no competitive conditional highlights)
 * - Prettier Admin-only PDF via a dedicated print sheet (LOGO ON RIGHT)
 *
 * IMPORTANT LIMITATION:
 * Apps Script cannot prevent a user from viewing sheet data if they already have access to the spreadsheet.
 * To truly restrict what Viewers can see, do not share the master spreadsheet with them.
 */

/** -------------------------
 *  Constants / Configuration
 *  ------------------------- */

const SHEETS = {
  MEMBERS: 'Members',
  EVENTS: 'Events',
  PROCESSED: 'ProcessedData',
  MEMBER_CHART: 'MemberChart',
  MEMBER_CHART_PRINT: 'MemberChart_Print',
  MY_CHART: 'MyChart',
  ROLES: 'Roles'
};

const HEADERS = {
  MEMBERS: ['First Name', 'Last Name', 'Kesem Name', 'Email', 'Member ID'],
  EVENTS: ['Event Name', 'Date', 'Total Revenue', 'Raw Attendance List', 'Processed Attendance', 'Revenue per Shift'],
  PROCESSED: ['Event', 'Date', 'Member', 'Shifts', 'Revenue'],
  MEMBER_CHART: ['Kesem Name', 'First Name', 'Last Name', 'Events', 'Total Revenue'],
  MY_CHART: ['Kesem Name', 'First Name', 'Last Name', 'Events', 'Total Revenue'],
  ROLES: ['Email', 'Role']
};

const AUTOMATION = {
  minIntervalMs: 1500,
  lockWaitMs: 15000
};

// PDF logo settings
const PDF_LOGO = {
  driveFileId: '1SJ91797QUSOthzN4Drpw568CSwF20sKw',
  anchorCellA1: 'E1', // top-right
  widthPx: 170,
  heightPx: 56
};

/** -------------------------
 *  Fast-path caches (per execution)
 *  ------------------------- */

let _SHEETS_INITIALIZED_ONCE_ = false;
let _ROLE_CACHE_ = {}; // key: email -> role
let _CURRENT_USER_EMAIL_CACHE_ = null;

/** -------------------------
 *  Menu / Entry Points
 *  ------------------------- */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Kesem Revenue System');

  menu.addItem('View My Results', 'UiService_viewMyResults');

  const email = RoleService.getCurrentUserEmail();
  const role = RoleService.getRoleForEmail(email);

  // Admin + Super Admin can recompute
  if (role === RoleService.ROLE_SUPER_ADMIN || role === RoleService.ROLE_ADMIN) {
    menu.addSeparator();
    menu.addItem('Recompute / Refresh Charts', 'ProcessingService_processAllEvents');
  }

  // Super Admin: can do everything
  if (role === RoleService.ROLE_SUPER_ADMIN) {
    menu.addSeparator();
    menu.addItem('Add Member', 'UiService_showAddMemberSidebar');
    menu.addItem('Add Event', 'UiService_showAddEventSidebar');
    menu.addSeparator();
    menu.addItem('Generate PDF', 'PdfService_generateMemberChartPDF');
    menu.addSeparator();
    menu.addItem('Apply Camp Kesem Theme', 'FormattingService_applyTheme');
    menu.addSeparator();
    menu.addItem('Initialize Sheets', 'initializeSheets');
    menu.addItem('Validate System Health', 'validateSystemHealth');
  }

  menu.addToUi();
}

function UiService_viewMyResults() {
  return UiService.viewMyResults();
}

function UiService_showAddMemberSidebar() {
  SecurityService.requireSuperAdmin_('Open Add Member sidebar');
  UiService.showAddMemberSidebar();
}

function UiService_showAddEventSidebar() {
  SecurityService.requireSuperAdmin_('Open Add Event sidebar');
  UiService.showAddEventSidebar();
}

/** Menu wrapper: must be top-level. */
function PdfService_generateMemberChartPDF() {
  SecurityService.requireAdmin_('Generate PDF');

  const blob = PdfService.generateMemberChartPDF();
  const fileName = `CampKesem_MemberChart_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm')}.pdf`;
  const file = DriveApp.createFile(blob).setName(fileName);

  SpreadsheetApp.getUi().alert(
    'PDF Generated',
    `Created PDF in Drive:\n${file.getUrl()}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** -------------------------
 *  System Health Check
 *  ------------------------- */

function validateSystemHealth() {
  SecurityService.requireAdmin_('Validate System Health');

  const report = {
    ok: true,
    timestamp: new Date().toISOString(),
    checks: []
  };

  function addCheck(name, ok, detail) {
    report.checks.push({ name, ok, detail: detail || '' });
    if (!ok) report.ok = false;
  }

  try {
    initializeSheets();
    addCheck('Sheets initialized', true, 'All required sheets exist with expected headers.');
  } catch (e) {
    addCheck('Sheets initialized', false, String(e && e.message ? e.message : e));
  }

  try {
    const email = RoleService.getCurrentUserEmail();
    addCheck('Current user email detected', Boolean(email), email || '(blank)');
  } catch (e) {
    addCheck('Current user email detected', false, String(e && e.message ? e.message : e));
  }

  try {
    const email = RoleService.getCurrentUserEmail();
    const role = RoleService.getRoleForEmail(email);
    const okRole = ['Super Admin', 'Admin', 'Viewer'].includes(role);
    addCheck('Role lookup', okRole, `role="${role}"`);
  } catch (e) {
    addCheck('Role lookup', false, String(e && e.message ? e.message : e));
  }

  try {
    const blob = PdfService.generateMemberChartPDF();
    const okBlob = blob && typeof blob.getBytes === 'function' && blob.getContentType() === 'application/pdf';
    addCheck('PDF export (blob)', Boolean(okBlob), okBlob ? `size=${blob.getBytes().length} bytes` : 'Blob invalid');
  } catch (e) {
    addCheck('PDF export (blob)', false, String(e && e.message ? e.message : e));
  }

  const lines = report.checks.map(c => `${c.ok ? 'OK' : 'FAIL'} — ${c.name}: ${c.detail}`);
  SpreadsheetApp.getUi().alert(
    `System Health: ${report.ok ? 'OK' : 'ISSUES FOUND'}`,
    lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log(`[validateSystemHealth] ${JSON.stringify(report)}`);
  return report;
}

/** -------------------------
 *  Sheet Initialization Utils
 *  ------------------------- */

function initializeSheets() {
  // In a single Apps Script execution, the sheet layout won't change.
  // This prevents repeated expensive header/sheet checks.
  if (_SHEETS_INITIALIZED_ONCE_) return { ok: true, skipped: true };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const members = _ensureSheetWithHeader_(ss, SHEETS.MEMBERS, HEADERS.MEMBERS);
  const events = _ensureSheetWithHeader_(ss, SHEETS.EVENTS, HEADERS.EVENTS);
  const processed = _ensureSheetWithHeader_(ss, SHEETS.PROCESSED, HEADERS.PROCESSED);
  const memberChart = _ensureSheetWithHeader_(ss, SHEETS.MEMBER_CHART, HEADERS.MEMBER_CHART);
  const myChart = _ensureSheetWithHeader_(ss, SHEETS.MY_CHART, HEADERS.MY_CHART);
  const roles = _ensureSheetWithHeader_(ss, SHEETS.ROLES, HEADERS.ROLES);

  _ensureSheetExists_(ss, SHEETS.MEMBER_CHART_PRINT);

  // Theme only newly created sheets (preserves "one-time apply" feel)
  if (typeof FormattingService !== 'undefined' && FormattingService && typeof FormattingService._applySheet_ === 'function') {
    if (members.created) FormattingService._applySheet_(members.sheet, 'MEMBERS');
    if (events.created) FormattingService._applySheet_(events.sheet, 'EVENTS');
    if (processed.created) FormattingService._applySheet_(processed.sheet, 'PROCESSED');
    if (memberChart.created) FormattingService._applySheet_(memberChart.sheet, 'MEMBER_CHART');
    if (myChart.created) FormattingService._applySheet_(myChart.sheet, 'MY_CHART');
    if (roles.created) FormattingService._applySheet_(roles.sheet, 'ROLES');
  }

  _SHEETS_INITIALIZED_ONCE_ = true;

  return {
    ok: true,
    created: {
      MEMBERS: members.created,
      EVENTS: events.created,
      PROCESSED: processed.created,
      MEMBER_CHART: memberChart.created,
      MY_CHART: myChart.created,
      ROLES: roles.created
    }
  };
}

function _ensureSheetExists_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function _ensureSheetWithHeader_(ss, sheetName, headerValues) {
  let sheet = ss.getSheetByName(sheetName);
  const created = !sheet;
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const headerRange = sheet.getRange(1, 1, 1, headerValues.length);
  const existing = headerRange.getValues()[0];

  if (!_arraysEqual_(existing, headerValues)) headerRange.setValues([headerValues]);
  sheet.setFrozenRows(1);

  return { sheet, created };
}

function _arraysEqual_(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b)) return false;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (String(a[i]) !== String(b[i])) return false;
  }
  return true;
}

/** -------------------------
 *  UI Service / Sidebars
 *  ------------------------- */

const UiService = {
  showAddMemberSidebar: function () {
    const html = HtmlService.createTemplateFromFile('AddMemberSidebar')
      .evaluate()
      .setTitle('Add Member');
    SpreadsheetApp.getUi().showSidebar(html);
  },

  showAddEventSidebar: function () {
    const html = HtmlService.createTemplateFromFile('AddEventSidebar')
      .evaluate()
      .setTitle('Add Event');
    SpreadsheetApp.getUi().showSidebar(html);
  },

  viewMyResults: function () {
    initializeSheets();

    const email = RoleService.getCurrentUserEmail();
    if (!email) {
      SpreadsheetApp.getUi().alert('Unable to detect your email. Ask an Admin for help.');
      return { ok: false, message: 'No email detected.' };
    }

    const member = MemberService.findMemberByEmail(email);
    if (!member) {
      SpreadsheetApp.getUi().alert(
        'No Member record found for your email.',
        `Email: ${email}\nAsk an Admin to add your email to the Members sheet.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return { ok: false, message: 'Member not found for email.' };
    }

    const result = MyChartService.buildForMember(member);
    SpreadsheetApp.getUi().alert('My Results Updated', 'Your MyChart sheet has been updated.', SpreadsheetApp.getUi().ButtonSet.OK);
    return { ok: true, ...result };
  }
};

function handleAddMemberStub(formData) {
  try {
    SecurityService.requireSuperAdmin_('Add Member');
    return MemberService.addMember(formData);
  } catch (err) {
    return { ok: false, message: `Error adding member: ${err && err.message ? err.message : err}` };
  }
}

function handleAddEventStub(formData) {
  try {
    SecurityService.requireSuperAdmin_('Add Event');
    return EventService.addEvent(formData);
  } catch (err) {
    return { ok: false, message: `Error adding event: ${err && err.message ? err.message : err}` };
  }
}

/** -------------------------
 *  SecurityService
 *  ------------------------- */

const SecurityService = {
  requireAdmin_: function (actionName) {
    const email = RoleService.getCurrentUserEmail();
    const role = RoleService.getRoleForEmail(email);
    const ok = role === RoleService.ROLE_ADMIN || role === RoleService.ROLE_SUPER_ADMIN;
    if (!ok) throw new Error(`Permission denied for "${actionName}". Your role is "${role}".`);
  },

  requireSuperAdmin_: function (actionName) {
    const email = RoleService.getCurrentUserEmail();
    const role = RoleService.getRoleForEmail(email);
    const ok = role === RoleService.ROLE_SUPER_ADMIN;
    if (!ok) throw new Error(`Permission denied for "${actionName}". Your role is "${role}".`);
  }
};

/** -------------------------
 *  Automation Guard
 *  ------------------------- */

function _withProcessLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(AUTOMATION.lockWaitMs)) {
    Logger.log('[Automation] Could not acquire lock; skipping processing run.');
    return { ok: false, skipped: true, reason: 'lock_not_acquired' };
  }

  try {
    const props = PropertiesService.getDocumentProperties();
    const lastRun = Number(props.getProperty('PROCESSING_LAST_RUN_MS') || 0);
    const now = Date.now();

    if (now - lastRun < AUTOMATION.minIntervalMs) {
      Logger.log('[Automation] Debounced: processing ran too recently; skipping.');
      return { ok: false, skipped: true, reason: 'debounced' };
    }

    props.setProperty('PROCESSING_LAST_RUN_MS', String(now));
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** -------------------------
 *  RoleService
 *  ------------------------- */

const RoleService = {
  ROLE_SUPER_ADMIN: 'Super Admin',
  ROLE_ADMIN: 'Admin',
  ROLE_VIEWER: 'Viewer',

  getCurrentUserEmail: function () {
    if (_CURRENT_USER_EMAIL_CACHE_ != null) return _CURRENT_USER_EMAIL_CACHE_;

    let email = '';
    try {
      email = Session.getEffectiveUser().getEmail() || '';
    } catch (err) {
      email = '';
    }

    _CURRENT_USER_EMAIL_CACHE_ = String(email || '').trim().toLowerCase();
    return _CURRENT_USER_EMAIL_CACHE_;
  },

  getRoleForEmail: function (email) {
    initializeSheets();

    const e = String(email || '').trim().toLowerCase();
    if (!e) return RoleService.ROLE_VIEWER;

    if (_ROLE_CACHE_[e]) return _ROLE_CACHE_[e];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ROLES);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      _ROLE_CACHE_[e] = RoleService.ROLE_VIEWER;
      return _ROLE_CACHE_[e];
    }

    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowEmail = String(values[i][0] || '').trim().toLowerCase();
      if (rowEmail === e) {
        _ROLE_CACHE_[e] = RoleService._normalizeRole_(values[i][1]);
        return _ROLE_CACHE_[e];
      }
    }

    _ROLE_CACHE_[e] = RoleService.ROLE_VIEWER;
    return _ROLE_CACHE_[e];
  },

  _normalizeRole_: function (role) {
    const r = String(role || '').trim().toLowerCase();
    if (r === 'super admin' || r === 'superadmin' || r === 'super_admin') return RoleService.ROLE_SUPER_ADMIN;
    if (r === 'admin') return RoleService.ROLE_ADMIN;
    return RoleService.ROLE_VIEWER;
  }
};

/** -------------------------
 *  MemberService
 *  ------------------------- */

const MemberService = {
  addMember: function (data) {
    const normalized = MemberService._normalizeMemberData_(data);
    MemberService._validateMemberData_(normalized);

    initializeSheets();

    const existing = MemberService.findMemberByEmail(normalized.email);
    if (existing) throw new Error(`A member with email "${normalized.email}" already exists.`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MEMBERS);

    const existingIds = MemberService._getExistingMemberIds_(sheet);
    const memberId = MemberService._generateUniqueMemberId_(existingIds);

    sheet.appendRow([normalized.firstName, normalized.lastName, normalized.kesemName, normalized.email, memberId]);

    // Prevent "white text" or other inherited formatting on newly appended row
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, HEADERS.MEMBERS.length)
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontColor(FormattingService && FormattingService.COLORS ? FormattingService.COLORS.text : '#111827');

    _withProcessLock_(function () {
      ProcessingService.processAllEvents();
      return { ok: true };
    });

    return { ok: true, message: `Member added successfully: ${normalized.firstName} ${normalized.lastName}`, memberId };
  },

  getAllMembers: function () {
    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MEMBERS);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.MEMBERS.length).getValues();
    return values
      .filter(r => r.some(cell => String(cell).trim() !== ''))
      .map(r => ({
        firstName: String(r[0] || '').trim(),
        lastName: String(r[1] || '').trim(),
        kesemName: String(r[2] || '').trim(),
        email: String(r[3] || '').trim().toLowerCase(),
        memberId: String(r[4] || '').trim()
      }));
  },

  findMemberByEmail: function (email) {
    const e = String(email || '').trim().toLowerCase();
    if (!e) return null;

    const members = MemberService.getAllMembers();
    for (let i = 0; i < members.length; i++) {
      if (String(members[i].email || '').trim().toLowerCase() === e) return members[i];
    }
    return null;
  },

  _normalizeMemberData_: function (data) {
    const safe = data || {};
    return {
      firstName: String(safe.firstName || '').trim(),
      lastName: String(safe.lastName || '').trim(),
      kesemName: String(safe.kesemName || '').trim(),
      email: String(safe.email || '').trim().toLowerCase()
    };
  },

  _validateMemberData_: function (data) {
    const missing = [];
    if (!data.firstName) missing.push('First Name');
    if (!data.lastName) missing.push('Last Name');
    if (!data.kesemName) missing.push('Kesem Name');
    if (!data.email) missing.push('Email');
    if (missing.length) throw new Error(`Missing required field(s): ${missing.join(', ')}`);
    if (!/^.+@.+\..+$/.test(data.email)) throw new Error('Email must be a valid email address.');
  },

  _getExistingMemberIds_: function (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return new Set();

    const idValues = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    return new Set(idValues.map(r => String(r[0] || '').trim()).filter(v => v !== ''));
  },

  _generateUniqueMemberId_: function (existingIdsSet) {
    for (let attempts = 0; attempts < 10; attempts++) {
      const id = Utilities.getUuid();
      if (!existingIdsSet.has(id)) return id;
    }
    throw new Error('Unable to generate unique Member ID.');
  }
};

/** -------------------------
 *  EventService
 *  ------------------------- */

const EventService = {
  addEvent: function (data) {
    const normalized = EventService._normalizeEventData_(data);
    EventService._validateEventData_(normalized);

    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.EVENTS);

    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(15000)) return { ok: false, message: 'System busy—please try again.' };

    try {
      const dup = EventService._isDuplicateEvent_(sheet, normalized);
      if (dup) return { ok: false, duplicate: true, message: `Duplicate event detected (same Event Name + Date as row ${dup.row}).` };

      sheet.appendRow([normalized.eventName, normalized.eventDate, normalized.totalRevenue, normalized.rawAttendance, '', '']);

      // Prevent "white text" or other inherited formatting on newly appended row
      const newRow = sheet.getLastRow();
      sheet.getRange(newRow, 1, 1, HEADERS.EVENTS.length)
        .setFontFamily('Arial')
        .setFontSize(10)
        .setFontColor(FormattingService && FormattingService.COLORS ? FormattingService.COLORS.text : '#111827');

      sheet.getRange(newRow, 3).setNumberFormat('$#,##0.00');
      sheet.getRange(newRow, 2).setNumberFormat('MM/dd/yyyy');
    } finally {
      lock.releaseLock();
    }

    _withProcessLock_(function () {
      ProcessingService.processAllEvents();
      return { ok: true };
    });

    return { ok: true, message: `Event added successfully: ${normalized.eventName}` };
  },

  _normalizeEventData_: function (data) {
    const safe = data || {};
    const eventName = String(safe.eventName || '').trim();
    const rawDate = String(safe.eventDate || '').trim();
    const revenueStr = String(safe.totalRevenue || '').trim();
    const rawAttendance = String(safe.rawAttendance || '');

    // Parse date safely in *local timezone* to avoid off-by-one from UTC parsing.
    let eventDate = null;
    if (rawDate) {
      const m = rawDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m) {
        const y = Number(m[1]);
        const mo = Number(m[2]) - 1;
        const d = Number(m[3]);
        // Noon local time avoids DST/UTC edge cases
        eventDate = new Date(y, mo, d, 12, 0, 0);
      } else {
        const parsed = new Date(rawDate);
        eventDate = parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : null;
      }
    }

    return {
      eventName,
      eventDate,
      totalRevenue: revenueStr === '' ? NaN : Number(revenueStr),
      rawAttendance
    };
  },

  _validateEventData_: function (data) {
    const missing = [];
    if (!data.eventName) missing.push('Event Name');
    const dateIsValid = data.eventDate instanceof Date && !isNaN(data.eventDate.getTime());
    if (!dateIsValid) missing.push('Date');
    const revenueIsValid = typeof data.totalRevenue === 'number' && isFinite(data.totalRevenue) && data.totalRevenue > 0;
    if (!revenueIsValid) missing.push('Total Revenue');
    if (missing.length) throw new Error(`Missing/invalid field(s): ${missing.join(', ')}`);
  },

  _normalizeEventNameKey_: function (s) {
    return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  },

  // IMPORTANT: kept as MM/dd to preserve your current duplicate-detection behavior exactly.
  // NOTE: This means events on same month/day in different years are considered duplicates.
  _dateKey_: function (d) {
    if (d instanceof Date && !isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd');
    }

    const s = String(d || '').trim();
    if (!s) return '';

    const parsed = new Date(s);
    if (parsed instanceof Date && !isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'MM/dd');
    }

    return s;
  },

  _isDuplicateEvent_: function (sheet, normalized) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const targetName = EventService._normalizeEventNameKey_(normalized.eventName);
    const targetDate = EventService._dateKey_(normalized.eventDate);

    for (let i = 0; i < values.length; i++) {
      const name = EventService._normalizeEventNameKey_(values[i][0]);
      const date = EventService._dateKey_(values[i][1]);
      if (name === targetName && date === targetDate) return { row: i + 2 };
    }
    return null;
  }
};

/** -------------------------
 *  ProcessingService (Admin-only)
 *  ------------------------- */

const ProcessingService = {
  parseAttendance: function (rawString) {
    const input = rawString == null ? '' : String(rawString);
    const parts = input.split(/,|\r?\n/);

    const counts = {};
    for (let i = 0; i < parts.length; i++) {
      const raw = String(parts[i] || '').trim();
      if (!raw) continue;
      const key = _normalizeNameKey_(raw);
      counts[key] = (counts[key] || 0) + 1;
    }
    return counts;
  },

  matchAttendanceToMembers: function (attendanceMap) {
    const map = attendanceMap || {};
    const attendanceNames = Object.keys(map);

    initializeSheets();

    const members = MemberService.getAllMembers();
    const lookup = {};

    for (let i = 0; i < members.length; i++) {
      const m = members[i];
      const kesem = String(m.kesemName || '').trim();
      const full = `${String(m.firstName || '').trim()} ${String(m.lastName || '').trim()}`.trim();
      const displayName = kesem || full || 'Unknown Member';
      const memberId = String(m.memberId || '').trim();
      if (!memberId) continue;

      if (kesem) lookup[_normalizeNameKey_(kesem)] = { memberId, name: displayName };
      if (full) lookup[_normalizeNameKey_(full)] = { memberId, name: displayName };
    }

    const matched = [];
    const unmatched = [];

    for (let i = 0; i < attendanceNames.length; i++) {
      const rawName = attendanceNames[i];
      const shifts = Number(map[rawName]) || 0;
      const hit = lookup[_normalizeNameKey_(rawName)];
      if (hit && shifts > 0) matched.push({ memberId: hit.memberId, name: hit.name, shifts });
      else unmatched.push(rawName);
    }

    if (unmatched.length) Logger.log(`[ProcessingService.matchAttendanceToMembers] Unmatched: ${JSON.stringify(unmatched)}`);
    return { matched, unmatched };
  },

  calculateRevenueDistribution: function (totalRevenue, matchedList) {
    const revenueNum = Number(totalRevenue);
    if (!isFinite(revenueNum) || revenueNum <= 0) throw new Error('totalRevenue must be a number > 0');

    const list = Array.isArray(matchedList) ? matchedList : [];
    const totalShifts = list.reduce((sum, m) => sum + (Number(m.shifts) || 0), 0);

    if (!totalShifts || totalShifts <= 0) {
      return list.map(m => ({
        memberId: m.memberId,
        name: m.name,
        shifts: Number(m.shifts) || 0,
        revenuePerPerson: 0
      }));
    }

    const revenuePerShift = revenueNum / totalShifts;

    return list.map(m => {
      const shifts = Number(m.shifts) || 0;
      const raw = shifts * revenuePerShift;
      return { memberId: m.memberId, name: m.name, shifts, revenuePerPerson: Math.round(raw * 100) / 100 };
    });
  },

  processAllEvents: function () {
    SecurityService.requireAdmin_('Process all events');

    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName(SHEETS.EVENTS);
    const processedSheet = ss.getSheetByName(SHEETS.PROCESSED);

    ProcessingService._clearProcessedData_(processedSheet);

    const lastRow = eventsSheet.getLastRow();
    const outputRows = [];

    if (lastRow >= 2) {
      const eventValues = eventsSheet.getRange(2, 1, lastRow - 1, HEADERS.EVENTS.length).getValues();

      // Batch-write Processed Attendance (col 5) and Revenue per Shift (col 6)
      const processedAttendanceCol = [];
      const revenuePerShiftCol = [];

      for (let i = 0; i < eventValues.length; i++) {
        const row = eventValues[i];

        const eventName = String(row[0] || '').trim();
        const eventDate = row[1];
        const totalRevenue = row[2];
        const rawAttendance = row[3];

        const isBlank = !eventName && !eventDate && !totalRevenue && !rawAttendance;
        if (isBlank) {
          processedAttendanceCol.push(['']);
          revenuePerShiftCol.push(['']);
          continue;
        }

        try {
          const attendanceMap = ProcessingService.parseAttendance(rawAttendance);

          const processedAttendanceText = Object.keys(attendanceMap)
            .sort()
            .map(k => `${k} x${attendanceMap[k]}`)
            .join('\n');

          processedAttendanceCol.push([processedAttendanceText]);

          const totalShifts = Object.keys(attendanceMap).reduce((sum, k) => sum + (Number(attendanceMap[k]) || 0), 0);

          const revenueNum = Number(totalRevenue);
          const revenuePerShift =
            isFinite(revenueNum) && revenueNum > 0 && totalShifts > 0
              ? Math.round((revenueNum / totalShifts) * 100) / 100
              : 0;

          revenuePerShiftCol.push([revenuePerShift]);

          const matchResult = ProcessingService.matchAttendanceToMembers(attendanceMap);
          const enriched = ProcessingService.calculateRevenueDistribution(totalRevenue, matchResult.matched);

          for (let j = 0; j < enriched.length; j++) {
            const m = enriched[j];
            outputRows.push([eventName, eventDate, m.name, m.shifts, m.revenuePerPerson]);
          }
        } catch (err) {
          Logger.log(`[ProcessingService.processAllEvents] ERROR event "${eventName}": ${err && err.message ? err.message : err}`);
          // Keep alignment of batch arrays even on error
          processedAttendanceCol.push(['']);
          revenuePerShiftCol.push(['']);
        }
      }

      const writeRows = eventValues.length;
      if (writeRows > 0) {
        eventsSheet.getRange(2, 5, writeRows, 1).setValues(processedAttendanceCol).setWrap(true).setVerticalAlignment('top');
        eventsSheet.getRange(2, 6, writeRows, 1).setValues(revenuePerShiftCol).setNumberFormat('$#,##0.00');
      }
    }

    if (outputRows.length) {
      processedSheet.getRange(2, 1, outputRows.length, HEADERS.PROCESSED.length).setValues(outputRows);
      processedSheet.getRange(2, 5, outputRows.length, 1).setNumberFormat('$#,##0.00');
    }

    ChartService.buildMemberChart();
  },

  _clearProcessedData_: function (processedSheet) {
    processedSheet.getRange(1, 1, 1, HEADERS.PROCESSED.length).setValues([HEADERS.PROCESSED]);
    processedSheet.setFrozenRows(1);

    const lastRow = processedSheet.getLastRow();
    if (lastRow >= 2) processedSheet.getRange(2, 1, lastRow - 1, HEADERS.PROCESSED.length).clearContent();
  }
};

/** -------------------------
 *  ChartService
 *  ------------------------- */

const ChartService = {
  aggregateMemberTotals: function () {
    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.PROCESSED);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    const values = sheet
      .getRange(2, 1, lastRow - 1, HEADERS.PROCESSED.length)
      .getValues();

    const acc = {};

    for (let i = 0; i < values.length; i++) {
      const r = values[i];

      const eventName = String(r[0] || '').trim();
      const eventDateRaw = r[1];
      const memberName = String(r[2] || '').trim();
      const shifts = Number(r[3]) || 0;
      const revenue = Number(r[4]) || 0;

      if (!memberName) continue;

      const eventDate =
        eventDateRaw instanceof Date && !isNaN(eventDateRaw.getTime())
          ? eventDateRaw
          : null;

      if (!acc[memberName]) {
        acc[memberName] = { memberName, totalRevenue: 0, _eventsMeta: [] };
      }

      acc[memberName].totalRevenue += revenue;

      const mmdd = eventDate
        ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'MM/dd')
        : '??/??';

      const safeEventName = eventName || '(Unnamed Event)';

      acc[memberName]._eventsMeta.push({
        date: eventDate || new Date(0),
        text: `${safeEventName} x${shifts} (${mmdd})`
      });
    }

    const result = {};
    const memberNames = Object.keys(acc);

    for (let i = 0; i < memberNames.length; i++) {
      const name = memberNames[i];
      const entry = acc[name];

      entry._eventsMeta.sort((a, b) => a.date.getTime() - b.date.getTime());

      result[name] = {
        memberName: entry.memberName,
        totalRevenue: Math.round(entry.totalRevenue * 100) / 100,
        events: entry._eventsMeta.map(e => e.text)
      };
    }

    return result;
  },

  buildMemberChart: function () {
    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MEMBER_CHART);

    sheet.clearContents();

    sheet.getRange(1, 1, 1, HEADERS.MEMBER_CHART.length).setValues([HEADERS.MEMBER_CHART]);
    sheet.setFrozenRows(1);

    const aggregated = ChartService.aggregateMemberTotals();
    const memberKeys = Object.keys(aggregated);

    if (!memberKeys.length) {
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setHorizontalAlignment('center');
      return { membersWritten: 0 };
    }

    const members = MemberService.getAllMembers();
    const lookup = {};
    for (let i = 0; i < members.length; i++) {
      const m = members[i];
      const kesem = String(m.kesemName || '').trim();
      const first = String(m.firstName || '').trim();
      const last = String(m.lastName || '').trim();
      const full = `${first} ${last}`.trim();

      if (kesem) lookup[_normalizeNameKey_(kesem)] = m;
      if (full) lookup[_normalizeNameKey_(full)] = m;
    }

    memberKeys.sort((a, b) => {
      const ra = Number(aggregated[a].totalRevenue) || 0;
      const rb = Number(aggregated[b].totalRevenue) || 0;
      if (rb !== ra) return rb - ra;
      return String(a).localeCompare(String(b));
    });

    const rows = memberKeys.map(keyName => {
      const entry = aggregated[keyName];
      const hit = lookup[_normalizeNameKey_(keyName)];

      const kesemName = hit ? String(hit.kesemName || '').trim() : String(entry.memberName || '').trim();
      const firstName = hit ? String(hit.firstName || '').trim() : '';
      const lastName = hit ? String(hit.lastName || '').trim() : '';
      const events = (entry.events || []).join('\n');
      const totalRevenue = Number(entry.totalRevenue) || 0;

      return [kesemName, firstName, lastName, events, totalRevenue];
    });

    sheet.getRange(2, 1, rows.length, HEADERS.MEMBER_CHART.length).setValues(rows);

    const startRow = 2;
    const n = rows.length;

    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(startRow, 1, n, 3).setHorizontalAlignment('left').setWrap(false);
    sheet.getRange(startRow, 4, n, 1).setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
    sheet.getRange(startRow, 5, n, 1).setNumberFormat('$#,##0.00').setHorizontalAlignment('right');

    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 320);
    sheet.setColumnWidth(5, 140);

    return { membersWritten: n };
  }
};

/** -------------------------
 *  MyChartService
 *  ------------------------- */

const MyChartService = {
  buildForMember: function (member) {
    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MY_CHART);

    // Clear values only (keep theme)
    sheet.clearContents();

    // Header
    sheet.getRange(1, 1, 1, HEADERS.MY_CHART.length).setValues([HEADERS.MY_CHART]);
    sheet.setFrozenRows(1);

    const m = member || {};
    const kesemName = String(m.kesemName || '').trim();
    const firstName = String(m.firstName || '').trim();
    const lastName = String(m.lastName || '').trim();

    const aggregated = ChartService.aggregateMemberTotals();

    const full = `${firstName} ${lastName}`.trim();
    const keys = Object.keys(aggregated);

    let hitKey = '';
    const kesemKey = _normalizeNameKey_(kesemName);
    const fullKey = _normalizeNameKey_(full);

    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const nk = _normalizeNameKey_(k);
      if (nk === kesemKey || nk === fullKey) {
        hitKey = k;
        break;
      }
    }

    const entry = hitKey ? aggregated[hitKey] : null;
    const events = entry && entry.events ? entry.events.join('\n') : '';
    const totalRevenue = entry ? Number(entry.totalRevenue) || 0 : 0;

    // Write the single output row
    sheet.getRange(2, 1, 1, HEADERS.MY_CHART.length).setValues([
      [kesemName || (entry ? entry.memberName : ''), firstName, lastName, events, totalRevenue]
    ]);

    // Force body row font color (fixes "white text" in MyChart)
    sheet.getRange(2, 1, 1, HEADERS.MY_CHART.length)
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontColor(FormattingService && FormattingService.COLORS ? FormattingService.COLORS.text : '#111827');

    // Minimal formatting (similar to MemberChart)
    sheet.getRange(2, 4).setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
    sheet.getRange(2, 5).setNumberFormat('$#,##0.00').setHorizontalAlignment('right');
    sheet.getRange(2, 1, 1, 3).setHorizontalAlignment('left').setWrap(false);

    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 320);
    sheet.setColumnWidth(5, 140);

    return { ok: true };
  }
};

/** -------------------------
 *  PdfLayoutService (Print Sheet Builder)
 *  ------------------------- */

const PdfLayoutService = {
  buildMemberChartPrintSheet_: function () {
    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const source = ss.getSheetByName(SHEETS.MEMBER_CHART);
    const printSheet = ss.getSheetByName(SHEETS.MEMBER_CHART_PRINT) || ss.insertSheet(SHEETS.MEMBER_CHART_PRINT);

    // Clean slate (formatting + content)
    printSheet.clear({ contentsOnly: false });
    printSheet.setHiddenGridlines(true);
    printSheet.setFrozenRows(0);

    // Ensure no old merged ranges remain in the header area (prevents merge errors)
    printSheet.getRange('A1:E3').breakApart();

    // Column widths (give logo room on the right)
    printSheet.setColumnWidth(1, 170);
    printSheet.setColumnWidth(2, 140);
    printSheet.setColumnWidth(3, 140);
    printSheet.setColumnWidth(4, 340);
    printSheet.setColumnWidth(5, 160); // E wider for logo

    // Header area (give room for logo)
    printSheet.setRowHeight(1, 54);
    printSheet.setRowHeight(2, 24);
    printSheet.setRowHeight(3, 14);

    const title = 'Camp Kesem — Member Revenue Summary';
    const subtitle = `Generated ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy h:mm a')}`;

    // Title in A1:D1 (merged) — logo sits in E1
    const titleRange = printSheet.getRange(1, 1, 1, 4); // A1:D1
    titleRange.merge();
    titleRange
      .setValue(title)
      .setFontFamily('Arial')
      .setFontSize(18)
      .setFontWeight('bold')
      .setFontColor('#0B4F8A')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left');

    // Subtitle in A2:D2 (merged)
    const subtitleRange = printSheet.getRange(2, 1, 1, 4); // A2:D2
    subtitleRange.merge();
    subtitleRange
      .setValue(subtitle)
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontColor('#374151')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left');

    // Place logo at E1 (top-right)
    _insertOrReplaceLogoOnSheet_(printSheet);

    // Divider line (row 3)
    printSheet.getRange(3, 1, 1, 5).setBackground('#E5E7EB');

    // Copy MemberChart (values only) starting row 4
    const sourceLastRow = source.getLastRow();
    const sourceLastCol = source.getLastColumn();

    if (sourceLastRow >= 1 && sourceLastCol >= 1) {
      const tableValues = source.getRange(1, 1, sourceLastRow, sourceLastCol).getValues();
      printSheet.getRange(4, 1, tableValues.length, tableValues[0].length).setValues(tableValues);
    }

    // Table header row (row 4)
    const tableHeader = printSheet.getRange(4, 1, 1, 5);
    tableHeader
      .setBackground('#0B4F8A')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    printSheet.setRowHeight(4, 28);

    // Body formatting
    const bodyStart = 5;
    const bodyRows = Math.max(0, sourceLastRow - 1);
    if (bodyRows > 0) {
      const bandings = printSheet.getBandings();
      for (let i = 0; i < bandings.length; i++) bandings[i].remove();

      const tableRange = printSheet.getRange(4, 1, bodyRows + 1, 5);
      const banding = tableRange.applyRowBanding();
      banding.setHeaderRowColor('#0B4F8A');
      banding.setFirstRowColor('#F3F8FF');
      banding.setSecondRowColor('#FFFFFF');

      tableRange.setBorder(true, true, true, true, true, true, '#D0D7DE', SpreadsheetApp.BorderStyle.SOLID);

      printSheet.getRange(bodyStart, 4, bodyRows, 1).setWrap(true).setVerticalAlignment('top');
      printSheet.getRange(bodyStart, 5, bodyRows, 1).setNumberFormat('$#,##0.00').setHorizontalAlignment('right');

      printSheet.setRowHeights(bodyStart, bodyRows, 24);
    }

    return { ok: true, rows: sourceLastRow };
  }
};

/** -------------------------
 *  PdfService (Admin only, prettier layout)
 *  ------------------------- */

const PdfService = {
  generateMemberChartPDF: function () {
    SecurityService.requireAdmin_('Generate MemberChart PDF');

    PdfLayoutService.buildMemberChartPrintSheet_();

    // Helps ensure images and formatting render before export
    SpreadsheetApp.flush();
    Utilities.sleep(400);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const printSheet = ss.getSheetByName(SHEETS.MEMBER_CHART_PRINT);

    const ssId = ss.getId();
    const gid = printSheet.getSheetId();

    const params = {
      format: 'pdf',
      size: 'letter',
      portrait: 'false',
      fitw: 'true',
      gridlines: 'false',
      printtitle: 'false',
      pagenumbers: 'false',
      sheetnames: 'false',
      attachment: 'false',
      margin_top: '0.50',
      margin_bottom: '0.50',
      margin_left: '0.50',
      margin_right: '0.50'
    };

    const query = Object.keys(params).map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`).join('&');
    const url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(ssId)}/export?gid=${encodeURIComponent(gid)}&${query}`;

    const resp = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` } });
    return resp.getBlob().setName('CampKesem_MemberChart.pdf');
  }
};

/** -------------------------
 *  FormattingService (Presentation Only; no competitive highlighting)
 *  ------------------------- */

function FormattingService_applyTheme() {
  SecurityService.requireAdmin_('Apply Camp Kesem Theme');
  return FormattingService.applyTheme();
}

const FormattingService = {
  COLORS: {
    headerBlue: '#0B4F8A',
    accentGreen: '#1E8E3E',
    bandOdd: '#F3F8FF',
    bandEven: '#F4FBF6',
    grid: '#D0D7DE',
    text: '#111827',
    mutedText: '#6B7280'
  },

  applyTheme: function () {
    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    FormattingService._applySheet_(ss.getSheetByName(SHEETS.MEMBERS), 'MEMBERS');
    FormattingService._applySheet_(ss.getSheetByName(SHEETS.EVENTS), 'EVENTS');
    FormattingService._applySheet_(ss.getSheetByName(SHEETS.PROCESSED), 'PROCESSED');
    FormattingService._applySheet_(ss.getSheetByName(SHEETS.MEMBER_CHART), 'MEMBER_CHART');
    FormattingService._applySheet_(ss.getSheetByName(SHEETS.MY_CHART), 'MY_CHART');
    FormattingService._applySheet_(ss.getSheetByName(SHEETS.ROLES), 'ROLES');

    SpreadsheetApp.getUi().alert('Theme Applied', 'Camp Kesem formatting has been applied (presentation only).', SpreadsheetApp.getUi().ButtonSet.OK);
    return { ok: true };
  },

  _applySheet_: function (sheet, kind) {
    if (!sheet) return;

    sheet.setHiddenGridlines(true);
    sheet.getDataRange().setFontFamily('Arial').setFontSize(10).setFontColor(FormattingService.COLORS.text);

    const lastCol = sheet.getLastColumn() || 1;
    const header = sheet.getRange(1, 1, 1, lastCol);
    header
      .setBackground(FormattingService.COLORS.headerBlue)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    sheet.setRowHeight(1, 34);

    FormattingService._applyBanding_(sheet);
    FormattingService._applyBorders_(sheet);
    FormattingService._ensureFilter_(sheet);

    if (kind === 'MEMBERS') FormattingService._formatMembers_(sheet);
    if (kind === 'EVENTS') FormattingService._formatEvents_(sheet);
    if (kind === 'PROCESSED') FormattingService._formatProcessed_(sheet);
    if (kind === 'MEMBER_CHART') FormattingService._formatMemberChartLike_(sheet);
    if (kind === 'MY_CHART') FormattingService._formatMemberChartLike_(sheet);
    if (kind === 'ROLES') FormattingService._formatRoles_(sheet);

    const lr = sheet.getLastRow();
    if (lr >= 2) sheet.setRowHeights(2, lr - 1, 24);
  },

  _applyBanding_: function (sheet) {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = Math.max(sheet.getLastColumn(), 1);

    const bandings = sheet.getBandings();
    for (let i = 0; i < bandings.length; i++) bandings[i].remove();

    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const banding = range.applyRowBanding();

    banding.setHeaderRowColor(FormattingService.COLORS.headerBlue);
    banding.setFirstRowColor(FormattingService.COLORS.bandOdd);
    banding.setSecondRowColor(FormattingService.COLORS.bandEven);
    banding.setFooterRowColor(null);
  },

  _applyBorders_: function (sheet) {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    sheet.getRange(1, 1, lastRow, lastCol)
      .setBorder(true, true, true, true, true, true, FormattingService.COLORS.grid, SpreadsheetApp.BorderStyle.SOLID);
  },

  _ensureFilter_: function (sheet) {
    if (sheet.getFilter()) return;
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  },

  _formatMembers_: function (sheet) {
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(4, 220);
    sheet.setColumnWidth(5, 220);

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('@STRING@');
      sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('@STRING@').setFontColor(FormattingService.COLORS.mutedText);
    }
  },

  _formatEvents_: function (sheet) {
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 110);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 260);
    sheet.setColumnWidth(5, 260);
    sheet.setColumnWidth(6, 140);

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('MM/dd/yyyy');
      sheet.getRange(2, 3, lastRow - 1, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('$#,##0.00');

      sheet.getRange(2, 4, lastRow - 1, 2).setWrap(true).setVerticalAlignment('top');
      sheet.getRange(2, 5, lastRow - 1, 1).setBackground('#EAF6EE');
    }
  },

  _formatProcessed_: function (sheet) {
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 110);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 120);

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('MM/dd/yyyy');
      sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('0');
      sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('$#,##0.00');
    }
  },

  _formatMemberChartLike_: function (sheet) {
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 320);
    sheet.setColumnWidth(5, 140);

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(2, 4, lastRow - 1, 1).setWrap(true).setVerticalAlignment('top');
      sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('$#,##0.00').setHorizontalAlignment('right');
      sheet.getRange(2, 1, lastRow - 1, 3).setHorizontalAlignment('left');

      sheet.getRange(1, 5).setBackground(FormattingService.COLORS.accentGreen).setFontColor('#FFFFFF');
      sheet.setConditionalFormatRules([]); // ensure no old rules remain
    }
  },

  _formatRoles_: function (sheet) {
    sheet.setColumnWidth(1, 240);
    sheet.setColumnWidth(2, 140);

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('@STRING@');
  }
};

/** -------------------------
 *  Helpers
 *  ------------------------- */

function _insertOrReplaceLogoOnSheet_(sheet) {
  if (!sheet) return { ok: false, reason: 'no_sheet' };
  if (!PDF_LOGO || !PDF_LOGO.driveFileId) return { ok: false, reason: 'no_logo_file_id' };

  // Remove existing images to avoid stacking logos each run
  const imgs = sheet.getImages();
  for (let i = 0; i < imgs.length; i++) imgs[i].remove();

  // Clear cells behind the logo so nothing "duplicates" visually under transparent pixels
  sheet.getRange('E1:E2').clearContent().setBackground('#FFFFFF');

  const file = DriveApp.getFileById(PDF_LOGO.driveFileId);
  const blob = file.getBlob();

  const anchor = sheet.getRange(PDF_LOGO.anchorCellA1); // E1
  const img = sheet.insertImage(blob, anchor.getColumn(), anchor.getRow());

  img.setWidth(PDF_LOGO.widthPx);
  img.setHeight(PDF_LOGO.heightPx);

  // Nudge toward the top-right corner
  if (typeof img.setOffsetX === 'function') img.setOffsetX(10);
  if (typeof img.setOffsetY === 'function') img.setOffsetY(2);

  return { ok: true, fileName: file.getName() };
}

function _normalizeNameKey_(name) {
  return String(name || '').trim().toLowerCase();
}

// Admin + Super Admin menu wrapper for recompute
function ProcessingService_processAllEvents() {
  SecurityService.requireAdmin_('Process all events');
  return _withProcessLock_(function () {
    ProcessingService.processAllEvents();
    SpreadsheetApp.getUi().alert(
      'Recompute Complete',
      'ProcessedData + charts have been refreshed.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return { ok: true };
  });
}
