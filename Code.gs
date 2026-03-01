/**
 * Camp Kesem Event Revenue Tracking System — Production
 *
 * Includes:
 * - Admin/Viewer permissions (Option A)
 * - Duplicate event prevention + locking
 * - Processed Attendance column on Events
 * - Camp Kesem bold theme formatting (no competitive conditional highlights)
 * - Prettier Admin-only PDF via a dedicated print sheet (NO LOGO)
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

/** -------------------------
 *  Menu / Entry Points
 *  ------------------------- */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Kesem Revenue System');

  menu.addItem('View My Results', 'UiService_viewMyResults');

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

  menu.addToUi();
}

function UiService_viewMyResults() {
  return UiService.viewMyResults();
}

function UiService_showAddMemberSidebar() {
  SecurityService.requireAdmin_('Open Add Member sidebar');
  UiService.showAddMemberSidebar();
}

function UiService_showAddEventSidebar() {
  SecurityService.requireAdmin_('Open Add Event sidebar');
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSheetWithHeader_(ss, SHEETS.MEMBERS, HEADERS.MEMBERS);
  _ensureSheetWithHeader_(ss, SHEETS.EVENTS, HEADERS.EVENTS);
  _ensureSheetWithHeader_(ss, SHEETS.PROCESSED, HEADERS.PROCESSED);
  _ensureSheetWithHeader_(ss, SHEETS.MEMBER_CHART, HEADERS.MEMBER_CHART);
  _ensureSheetWithHeader_(ss, SHEETS.MY_CHART, HEADERS.MY_CHART);
  _ensureSheetWithHeader_(ss, SHEETS.ROLES, HEADERS.ROLES);

  _ensureSheetExists_(ss, SHEETS.MEMBER_CHART_PRINT);
}

function _ensureSheetExists_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function _ensureSheetWithHeader_(ss, sheetName, headerValues) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const headerRange = sheet.getRange(1, 1, 1, headerValues.length);
  const existing = headerRange.getValues()[0];

  if (!_arraysEqual_(existing, headerValues)) headerRange.setValues([headerValues]);
  sheet.setFrozenRows(1);
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
    SecurityService.requireAdmin_('Add Member');
    return MemberService.addMember(formData);
  } catch (err) {
    return { ok: false, message: `Error adding member: ${err && err.message ? err.message : err}` };
  }
}

function handleAddEventStub(formData) {
  try {
    SecurityService.requireAdmin_('Add Event');
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
    let email = '';
    try {
      email = Session.getEffectiveUser().getEmail() || '';
    } catch (err) {
      email = '';
    }
    return String(email || '').trim().toLowerCase();
  },

  getRoleForEmail: function (email) {
    initializeSheets();

    const e = String(email || '').trim().toLowerCase();
    if (!e) return RoleService.ROLE_VIEWER;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ROLES);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return RoleService.ROLE_VIEWER;

    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowEmail = String(values[i][0] || '').trim().toLowerCase();
      if (rowEmail === e) return RoleService._normalizeRole_(values[i][1]);
    }

    return RoleService.ROLE_VIEWER;
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
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 3).setNumberFormat('$#,##0.00');
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

    return {
      eventName,
      eventDate: rawDate ? new Date(rawDate) : null,
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

  _dateKey_: function (d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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

      for (let i = 0; i < eventValues.length; i++) {
        const row = eventValues[i];

        const eventName = String(row[0] || '').trim();
        const eventDate = row[1];
        const totalRevenue = row[2];
        const rawAttendance = row[3];

        const isBlank = !eventName && !eventDate && !totalRevenue && !rawAttendance;
        if (isBlank) continue;

        try {
          const attendanceMap = ProcessingService.parseAttendance(rawAttendance);

          const processedAttendanceText = Object.keys(attendanceMap)
            .sort()
            .map(k => `${k} x${attendanceMap[k]}`)
            .join('\n');
          eventsSheet.getRange(i + 2, 5).setValue(processedAttendanceText).setWrap(true);

          const totalShifts = Object.keys(attendanceMap).reduce((sum, k) => sum + (Number(attendanceMap[k]) || 0), 0);

          const revenueNum = Number(totalRevenue);
          const revenuePerShift =
            isFinite(revenueNum) && revenueNum > 0 && totalShifts > 0
              ? Math.round((revenueNum / totalShifts) * 100) / 100
              : 0;

          eventsSheet.getRange(i + 2, 6).setValue(revenuePerShift).setNumberFormat('$#,##0.00');

          const matchResult = ProcessingService.matchAttendanceToMembers(attendanceMap);
          const enriched = ProcessingService.calculateRevenueDistribution(totalRevenue, matchResult.matched);

          for (let j = 0; j < enriched.length; j++) {
            const m = enriched[j];
            outputRows.push([eventName, eventDate, m.name, m.shifts, m.revenuePerPerson]);
          }
        } catch (err) {
          Logger.log(`[ProcessingService.processAllEvents] ERROR event "${eventName}": ${err && err.message ? err.message : err}`);
        }
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
 *  PdfLayoutService (Print Sheet Builder) — NO LOGO
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

    // Column widths
    printSheet.setColumnWidth(1, 170);
    printSheet.setColumnWidth(2, 140);
    printSheet.setColumnWidth(3, 140);
    printSheet.setColumnWidth(4, 340);
    printSheet.setColumnWidth(5, 140);

    // Header area
    printSheet.setRowHeight(1, 48);
    printSheet.setRowHeight(2, 24);
    printSheet.setRowHeight(3, 14);

    const title = 'Camp Kesem — Member Revenue Summary';
    const subtitle = `Generated ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy h:mm a')}`;

    // Title in A1:E1 (merged)
    const titleRange = printSheet.getRange(1, 1, 1, 5);
    titleRange.merge();
    titleRange
      .setValue(title)
      .setFontFamily('Arial')
      .setFontSize(18)
      .setFontWeight('bold')
      .setFontColor('#0B4F8A')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left');

    // Subtitle in A2:E2 (merged)
    const subtitleRange = printSheet.getRange(2, 1, 1, 5);
    subtitleRange.merge();
    subtitleRange
      .setValue(subtitle)
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontColor('#374151')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left');

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

function _normalizeNameKey_(name) {
  return String(name || '').trim().toLowerCase();
}
