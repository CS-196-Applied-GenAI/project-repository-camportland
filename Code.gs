/**
 * Camp Kesem Event Revenue Tracking System — Production
 * Prompts 1–18 (Final Wiring & Cleanup)
 *
 * Key behavior:
 * - Members + Events managed via sidebars (AddMemberSidebar / AddEventSidebar)
 * - Adding a member OR adding an event automatically triggers full recalculation:
 *     Events -> ProcessedData -> MemberChart (formatted)
 * - Roles stored in Roles sheet (detection only; not enforced yet)
 * - PDF export of MemberChart via menu item "Generate PDF" (creates Drive file + shows link)
 * - Health check: validateSystemHealth()
 *
 * Notes:
 * - This file intentionally contains no orphaned stubs or unused test helpers.
 * - UI HTML files must exist in the project: AddMemberSidebar, AddEventSidebar
 */

/** -------------------------
 *  Constants / Configuration
 *  ------------------------- */

const SHEETS = {
  MEMBERS: 'Members',
  EVENTS: 'Events',
  PROCESSED: 'ProcessedData',
  MEMBER_CHART: 'MemberChart',
  ROLES: 'Roles'
};

const HEADERS = {
  MEMBERS: ['First Name', 'Last Name', 'Kesem Name', 'Member ID'],
  EVENTS: ['Event Name', 'Date', 'Total Revenue', 'Raw Attendance List'],
  PROCESSED: ['Event', 'Date', 'Member', 'Shifts', 'Revenue'],
  MEMBER_CHART: ['Member Name', 'Events', 'Total Revenue'],
  ROLES: ['Email', 'Role']
};

/**
 * Automation settings (debounce prevents accidental back-to-back recalcs, e.g. double-submit UI).
 * Adjust minIntervalMs if you find it too aggressive.
 */
const AUTOMATION = {
  minIntervalMs: 1500,
  lockWaitMs: 15000
};

/** -------------------------
 *  Menu / Entry Points
 *  ------------------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Kesem Revenue System')
    .addItem('Add Member', 'UiService_showAddMemberSidebar')
    .addItem('Add Event', 'UiService_showAddEventSidebar')
    .addSeparator()
    .addItem('Generate PDF', 'PdfService_generateMemberChartPDF')
    .addSeparator()
    .addItem('Initialize Sheets', 'initializeSheets')
    .addItem('Validate System Health', 'validateSystemHealth')
    .addToUi();
}

/** Menu wrapper: must be top-level. */
function PdfService_generateMemberChartPDF() {
  const blob = PdfService.generateMemberChartPDF();

  const fileName = `MemberChart_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm')}.pdf`;
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

/**
 * Prompt 18: validateSystemHealth()
 *
 * Performs a non-destructive health check:
 * - Ensures sheets + headers exist
 * - Verifies current user's email can be detected (Workspace)
 * - Verifies role lookup returns a valid role string
 * - Verifies MemberChart PDF export can produce a blob
 *
 * Returns a report object and also shows a UI alert summary.
 */
function validateSystemHealth() {
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
  _ensureSheetWithHeader_(ss, SHEETS.ROLES, HEADERS.ROLES);
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

function UiService_showAddMemberSidebar() {
  UiService.showAddMemberSidebar();
}

function UiService_showAddEventSidebar() {
  UiService.showAddEventSidebar();
}

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
  }
};

/**
 * Called by AddMemberSidebar.html via google.script.run
 * (kept as stable endpoint for UI; not a stub)
 */
function handleAddMemberStub(formData) {
  try {
    return MemberService.addMember(formData);
  } catch (err) {
    return { ok: false, message: `Error adding member: ${err && err.message ? err.message : err}` };
  }
}

/**
 * Called by AddEventSidebar.html via google.script.run
 * (kept as stable endpoint for UI; not a stub)
 */
function handleAddEventStub(formData) {
  try {
    return EventService.addEvent(formData);
  } catch (err) {
    return { ok: false, message: `Error adding event: ${err && err.message ? err.message : err}` };
  }
}

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
 *  RoleService (Detection Only)
 *  ------------------------- */

const RoleService = {
  ROLE_SUPER_ADMIN: 'Super Admin',
  ROLE_ADMIN: 'Admin',
  ROLE_VIEWER: 'Viewer',

  getCurrentUserEmail: function () {
    // Workspace-friendly identity
    let email = '';
    try {
      email = Session.getEffectiveUser().getEmail() || '';
    } catch (err) {
      email = '';
    }

    email = String(email || '').trim().toLowerCase();
    Logger.log(`[RoleService.getCurrentUserEmail] email="${email || '(blank)'}"`);
    return email;
  },

  getRoleForEmail: function (email) {
    initializeSheets();

    const e = String(email || '').trim().toLowerCase();
    if (!e) {
      Logger.log('[RoleService.getRoleForEmail] blank email -> Viewer');
      return RoleService.ROLE_VIEWER;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ROLES);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log(`[RoleService.getRoleForEmail] Roles empty -> Viewer for "${e}"`);
      return RoleService.ROLE_VIEWER;
    }

    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowEmail = String(values[i][0] || '').trim().toLowerCase();
      if (rowEmail === e) {
        const role = RoleService._normalizeRole_(values[i][1]);
        Logger.log(`[RoleService.getRoleForEmail] email="${e}" role="${role}"`);
        return role;
      }
    }

    Logger.log(`[RoleService.getRoleForEmail] email="${e}" role="Viewer" (default)`);
    return RoleService.ROLE_VIEWER;
  },

  /**
   * Roles sheet management (still not enforced; included because you requested adding admins earlier).
   */
  setRoleForEmail: function (email, role) {
    initializeSheets();

    const e = String(email || '').trim().toLowerCase();
    if (!e) throw new Error('Email is required.');

    const normalizedRole = RoleService._normalizeRole_(role);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ROLES);

    const lastRow = sheet.getLastRow();
    const values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];

    let foundRow = -1;
    for (let i = 0; i < values.length; i++) {
      const rowEmail = String(values[i][0] || '').trim().toLowerCase();
      if (rowEmail === e) {
        foundRow = i + 2; // sheet row
        break;
      }
    }

    if (foundRow > 0) sheet.getRange(foundRow, 1, 1, 2).setValues([[e, normalizedRole]]);
    else sheet.appendRow([e, normalizedRole]);

    Logger.log(`[RoleService.setRoleForEmail] email="${e}" role="${normalizedRole}"`);
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

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MEMBERS);

    const existingIds = MemberService._getExistingMemberIds_(sheet);
    const memberId = MemberService._generateUniqueMemberId_(existingIds);

    sheet.appendRow([normalized.firstName, normalized.lastName, normalized.kesemName, memberId]);

    // Prompt 13: auto-recalc
    _withProcessLock_(function () {
      Logger.log('[MemberService.addMember] Auto-processing after member add.');
      ProcessingService.processAllEvents();
      return { ok: true };
    });

    return {
      ok: true,
      message: `Member added successfully: ${normalized.firstName} ${normalized.lastName}`,
      memberId
    };
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
        memberId: String(r[3] || '').trim()
      }));
  },

  _normalizeMemberData_: function (data) {
    const safe = data || {};
    return {
      firstName: String(safe.firstName || '').trim(),
      lastName: String(safe.lastName || '').trim(),
      kesemName: String(safe.kesemName || '').trim()
    };
  },

  _validateMemberData_: function (data) {
    const missing = [];
    if (!data.firstName) missing.push('First Name');
    if (!data.lastName) missing.push('Last Name');
    if (!data.kesemName) missing.push('Kesem Name');
    if (missing.length) throw new Error(`Missing required field(s): ${missing.join(', ')}`);
  },

  _getExistingMemberIds_: function (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return new Set();

    const idValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    return new Set(idValues.map(r => String(r[0] || '').trim()).filter(v => v !== ''));
  },

  _generateUniqueMemberId_: function (existingIdsSet) {
    for (let attempts = 0; attempts < 10; attempts++) {
      const id = Utilities.getUuid();
      if (!existingIdsSet.has(id)) return id;
    }
    throw new Error('Unable to generate unique Member ID (UUID collision loop).');
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

    sheet.appendRow([
      normalized.eventName,
      normalized.eventDate,
      normalized.totalRevenue,
      normalized.rawAttendance
    ]);

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3).setNumberFormat('$#,##0.00');

    // Prompt 13: auto-recalc
    _withProcessLock_(function () {
      Logger.log('[EventService.addEvent] Auto-processing after event add.');
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

    const revenueIsValid =
      typeof data.totalRevenue === 'number' &&
      isFinite(data.totalRevenue) &&
      data.totalRevenue > 0;
    if (!revenueIsValid) missing.push('Total Revenue');

    if (missing.length) throw new Error(`Missing/invalid field(s): ${missing.join(', ')}`);
  }
};

/** -------------------------
 *  ProcessingService
 *  ------------------------- */

const ProcessingService = {
  parseAttendance: function (rawString) {
    const input = rawString == null ? '' : String(rawString);
    const parts = input.split(/,|\r?\n/);

    const counts = {};
    for (let i = 0; i < parts.length; i++) {
      const name = String(parts[i] || '').trim();
      if (!name) continue;
      counts[name] = (counts[name] || 0) + 1;
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

    if (unmatched.length) {
      Logger.log(`[ProcessingService.matchAttendanceToMembers] Unmatched: ${JSON.stringify(unmatched)}`);
    }

    return { matched, unmatched };
  },

  calculateRevenueDistribution: function (totalRevenue, matchedList) {
    const revenueNum = Number(totalRevenue);
    if (!isFinite(revenueNum) || revenueNum <= 0) {
      throw new Error('totalRevenue must be a number > 0');
    }

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
      return {
        memberId: m.memberId,
        name: m.name,
        shifts,
        revenuePerPerson: Math.round(raw * 100) / 100
      };
    });
  },

  processAllEvents: function () {
    Logger.log('--- [ProcessingService.processAllEvents] START ---');

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
          const matchResult = ProcessingService.matchAttendanceToMembers(attendanceMap);
          const enriched = ProcessingService.calculateRevenueDistribution(totalRevenue, matchResult.matched);

          for (let j = 0; j < enriched.length; j++) {
            const m = enriched[j];
            outputRows.push([eventName, eventDate, m.name, m.shifts, m.revenuePerPerson]);
          }
        } catch (err) {
          Logger.log(
            `[ProcessingService.processAllEvents] ERROR event "${eventName}": ${err && err.message ? err.message : err}`
          );
        }
      }
    }

    if (outputRows.length) {
      processedSheet.getRange(2, 1, outputRows.length, HEADERS.PROCESSED.length).setValues(outputRows);
      processedSheet.getRange(2, 5, outputRows.length, 1).setNumberFormat('$#,##0.00');
    }

    ChartService.buildMemberChart();

    Logger.log(`--- [ProcessingService.processAllEvents] END — processedRows=${outputRows.length} ---`);
  },

  _clearProcessedData_: function (processedSheet) {
    processedSheet.getRange(1, 1, 1, HEADERS.PROCESSED.length).setValues([HEADERS.PROCESSED]);
    processedSheet.setFrozenRows(1);

    const lastRow = processedSheet.getLastRow();
    if (lastRow >= 2) {
      processedSheet.getRange(2, 1, lastRow - 1, HEADERS.PROCESSED.length).clearContent();
    }
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

    const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.PROCESSED.length).getValues();
    const acc = {};

    for (let i = 0; i < values.length; i++) {
      const r = values[i];

      const eventName = String(r[0] || '').trim();
      const eventDateRaw = r[1];
      const memberName = String(r[2] || '').trim();
      const shifts = Number(r[3]) || 0;
      const revenue = Number(r[4]) || 0;

      if (!eventName && !memberName) continue;

      const eventDate =
        eventDateRaw instanceof Date && !isNaN(eventDateRaw.getTime()) ? eventDateRaw : null;

      if (!acc[memberName]) {
        acc[memberName] = { memberName, totalRevenue: 0, _eventsMeta: [] };
      }

      acc[memberName].totalRevenue += revenue;

      const mmdd = eventDate
        ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'MM/dd')
        : '??/??';

      acc[memberName]._eventsMeta.push({
        date: eventDate || new Date(0),
        text: `${eventName} x${shifts} (${mmdd})`
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

    // Clear MemberChart content, then rebuild
    sheet.clearContents();

    // Header
    sheet.getRange(1, 1, 1, HEADERS.MEMBER_CHART.length).setValues([HEADERS.MEMBER_CHART]);
    sheet.setFrozenRows(1);

    const aggregated = ChartService.aggregateMemberTotals();
    const memberNames = Object.keys(aggregated);

    if (!memberNames.length) {
      ChartService.formatMemberChart(1);
      return { membersWritten: 0 };
    }

    // Sorted results (by revenue desc, name asc)
    memberNames.sort((a, b) => {
      const ra = Number(aggregated[a].totalRevenue) || 0;
      const rb = Number(aggregated[b].totalRevenue) || 0;
      if (rb !== ra) return rb - ra;
      return String(a).localeCompare(String(b));
    });

    const rows = memberNames.map(name => {
      const entry = aggregated[name];
      return [entry.memberName, (entry.events || []).join('\n'), entry.totalRevenue];
    });

    sheet.getRange(2, 1, rows.length, HEADERS.MEMBER_CHART.length).setValues(rows);
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('$#,##0.00');
    sheet.getRange(2, 2, rows.length, 1).setWrap(true);

    ChartService.formatMemberChart(rows.length + 1);

    return { membersWritten: rows.length };
  },

  formatMemberChart: function (lastRowToFormat) {
    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MEMBER_CHART);

    const cols = HEADERS.MEMBER_CHART.length; // A:C
    const lastRow = Math.max(1, Number(lastRowToFormat) || sheet.getLastRow() || 1);

    // Bold header
    sheet.getRange(1, 1, 1, cols).setFontWeight('bold');

    // Alternating row colors (apply banding only to A:C area)
    const bandingRange = sheet.getRange(1, 1, lastRow, cols);
    const bandings = sheet.getBandings();
    for (let i = 0; i < bandings.length; i++) bandings[i].remove();
    bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

    // Auto-resize A:C
    sheet.autoResizeColumns(1, cols);
  }
};

/** -------------------------
 *  PdfService (Prompt 17)
 *  ------------------------- */

const PdfService = {
  /**
   * Export MemberChart sheet to PDF and return a Blob.
   * - Landscape
   * - Fit to width
   * - Reasonable margins
   * - Exports only MemberChart (by gid)
   * - Uses current sheet state (including filter state); does not modify filters.
   */
  generateMemberChartPDF: function () {
    initializeSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MEMBER_CHART);

    const ssId = ss.getId();
    const gid = sheet.getSheetId();

    // Google Sheets export parameters
    const params = {
      format: 'pdf',
      size: 'letter',
      portrait: 'false', // landscape
      fitw: 'true',      // fit to width
      gridlines: 'false',
      printtitle: 'false',
      pagenumbers: 'false',
      sheetnames: 'false',
      attachment: 'false',

      // Margins in inches
      margin_top: '0.50',
      margin_bottom: '0.50',
      margin_left: '0.50',
      margin_right: '0.50'
    };

    const query = Object.keys(params)
      .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
      .join('&');

    const url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(ssId)}/export?gid=${encodeURIComponent(gid)}&${query}`;

    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
    });

    return resp.getBlob().setName('MemberChart.pdf');
  }
};

/** -------------------------
 *  Helpers
 *  ------------------------- */

function _normalizeNameKey_(name) {
  return String(name || '').trim().toLowerCase();
}
