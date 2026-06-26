// ============================================================
// TIME TRACKING DASHBOARD — Google Apps Script Backend
// Code.gs
// ============================================================

// ============================================================
// SHEET NAMES
// ============================================================
var SHEET_TIME_ENTRIES = "time_entries";
var SHEET_USERS        = "users";           // single sheet: email | role | display_name
var SHEET_TEAM_ASSIGN  = "team_assignments";
var SHEET_CASES        = "Cases";           // fraud case solves from the Chrome extension

// Hidden roster names — analysts no longer on the team. Historical data
// stays in the sheets but is filtered out of the dashboard everywhere
// (Team Directory, exports, analyst dropdown, case views, etc.).
var HIDDEN_ROSTER_NAMES = [
  "cassandra buss",
  // Mexico analysts (removed from team)
  "alex mendez",
  "daniel jullian",
  "daniela lópez",
  "daniela lopez",
  "daniela suarez",
  "eddie escamilla",
  "enrique becerril",
  "pamela minero",
  "ricardo rico",
  "tori segura",
  "mauricio gavito paredes",
  "michael bracamontes",
  "diego garcía",
  "diego garcia"
];

function isHiddenRosterName(name) {
  return HIDDEN_ROSTER_NAMES.indexOf(String(name || "").trim().toLowerCase()) !== -1;
}

// ============================================================
// SPREADSHEET ID
// ============================================================
var SPREADSHEET_ID = "1Kl57TacbVJmTAJTLqJ_vVIFqTkQMxY0vC1ejBULya5M";

// ============================================================
// SHEET SETUP — run setupSheets() once after deployment
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var te = getOrCreateSheet(ss, SHEET_TIME_ENTRIES);
  if (te.getLastRow() === 0) {
    te.appendRow(["id","analyst","project","task","duration","date","category",
                  "start_time","end_time","edit_reason","edited_at","original_duration","sheet_row_id"]);
    te.getRange(1,1,1,13).setFontWeight("bold");
  }

  var us = getOrCreateSheet(ss, SHEET_USERS);
  if (us.getLastRow() === 0) {
    us.appendRow(["email","role","display_name"]);
    us.getRange(1,1,1,3).setFontWeight("bold");
    us.appendRow(["admin@example.com","admin","Admin User"]);
  }

  var ta = getOrCreateSheet(ss, SHEET_TEAM_ASSIGN);
  if (ta.getLastRow() === 0) {
    ta.appendRow(["analyst_name","supervisor_email","team_name","role"]);
    ta.getRange(1,1,1,4).setFontWeight("bold");
  }

  var cs = getOrCreateSheet(ss, SHEET_CASES);
  if (cs.getLastRow() === 0) {
    cs.appendRow(["Date","Analyst","Platform","Case ID","Source",
                  "Handle (sec)","Handle (min)","Solved At","dedupe_key"]);
    cs.getRange(1,1,1,9).setFontWeight("bold");
    cs.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert("Sheets created/verified successfully!");
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok:       true,
        version:  '3.8-server',
        serverTs: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Time Tracking Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ============================================================
// ROLE HELPERS
// ============================================================
function isPrivileged(role) {
  var r = (role || "").toLowerCase().trim();
  return r === "admin" || r === "supervisor";
}

// ============================================================
// USER LOOKUP — reads from single "users" sheet
// ============================================================
function getUserRecord(email) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) return null;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        return {
          email:        data[i][0].toString(),
          role:         data[i][1].toString(),
          display_name: data[i][2].toString() || null
        };
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

// ============================================================
// AUTH — two paths:
//   getCurrentUser()         legacy — used internally (debugAuth etc.)
//   getUserByEmail(email)    NEW — called from the frontend after Google Sign-In
//                            verifies the user's identity via JWT. The frontend
//                            decodes the JWT client-side (signature already
//                            verified by Google's library) and passes the email
//                            to this function to look up role + display_name.
// ============================================================
function getCurrentUser() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return { authenticated: false };

    var record      = getUserRecord(email) || {};
    var role        = record.role        || "analyst";
    var displayName = record.display_name || email.split("@")[0];

    return {
      authenticated: true,
      email:         email,
      role:          role,
      isSupervisor:  isPrivileged(role),
      displayName:   displayName
    };
  } catch(e) {
    return { authenticated: false, error: e.toString() };
  }
}

// NEW: called by the frontend after Google Sign-In succeeds.
// email is taken from the verified Google JWT — we trust it because
// Google's GSI library has already validated the signature before
// the frontend calls this function.
function getUserByEmail(email) {
  try {
    if (!email) return { authenticated: false, error: "No email provided" };

    // Only allow fetchrewards.com accounts
    var domain = email.split("@")[1] || "";
    if (domain.toLowerCase() !== "fetchrewards.com") {
      return { authenticated: false, error: "Please sign in with your Fetch Rewards account" };
    }

    var record      = getUserRecord(email) || {};
    var role        = record.role        || "analyst";
    var displayName = record.display_name || email.split("@")[0];

    return {
      authenticated: true,
      email:         email,
      role:          role,
      isSupervisor:  isPrivileged(role),
      displayName:   displayName
    };
  } catch(e) {
    return { authenticated: false, error: e.toString() };
  }
}

// ============================================================
// TIME ENTRIES — fetch
// ============================================================
function getTimeEntries(startDate, endDate) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    var headers = ["id","analyst","project","task","duration","date","category",
                   "start_time","end_time","edit_reason","edited_at","original_duration","sheet_row_id"];

    var start = new Date(startDate + "T00:00:00");
    var end   = new Date(endDate   + "T23:59:59");

    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;

      var dateVal   = row[5];
      var entryDate = (dateVal instanceof Date)
        ? dateVal
        : new Date(dateVal + "T00:00:00");

      if (entryDate >= start && entryDate <= end) {
        var entry = {};
        for (var j = 0; j < headers.length; j++) {
          var val = row[j];
          entry[headers[j]] = (val instanceof Date)
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd")
            : (val === "" ? null : val);
        }
        if (typeof entry.date === "object") {
          entry.date = Utilities.formatDate(new Date(entry.date), Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        results.push(entry);
      }
    }
    return results;
  } catch(e) {
    console.error("getTimeEntries error:", e);
    return { error: e.toString() };
  }
}

function getAvailableMonths() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data   = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues();
    var months = {};

    for (var i = 0; i < data.length; i++) {
      var val = data[i][0];
      if (!val) continue;
      var dateStr = (val instanceof Date)
        ? Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : val.toString();
      if (dateStr.length >= 7) months[dateStr.substring(0, 7)] = true;
    }

    return Object.keys(months).sort(function(a,b){ return b.localeCompare(a); });
  } catch(e) {
    console.error("getAvailableMonths error:", e);
    return [];
  }
}

// ============================================================
// TIME ENTRIES — update
// ============================================================
function updateTimeEntry(id, updates) {
  try {
    var email       = Session.getActiveUser().getEmail();
    var record      = getUserRecord(email) || {};
    var role        = record.role || "analyst";
    var displayName = record.display_name || "";

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
    if (!sheet || sheet.getLastRow() < 2) return { error: "Sheet not found" };

    var data     = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    var rowIndex = -1, entryAnalyst = "";

    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        rowIndex     = i + 2;
        entryAnalyst = data[i][1].toString();
        break;
      }
    }

    if (rowIndex === -1) return { error: "Entry not found" };
    if (!isPrivileged(role) && entryAnalyst !== displayName) {
      return { error: "Permission denied" };
    }

    var existingOriginal = data[rowIndex - 2][11];
    var originalDuration = data[rowIndex - 2][4];

    sheet.getRange(rowIndex, 4).setValue(updates.task     || data[rowIndex - 2][3]);
    sheet.getRange(rowIndex, 5).setValue(Number(updates.duration) || data[rowIndex - 2][4]);
    sheet.getRange(rowIndex, 8).setValue(updates.start_time || "");
    sheet.getRange(rowIndex, 9).setValue(updates.end_time   || "");
    sheet.getRange(rowIndex, 10).setValue(updates.edit_reason || "");
    sheet.getRange(rowIndex, 11).setValue(new Date().toISOString());
    if (!existingOriginal) sheet.getRange(rowIndex, 12).setValue(originalDuration);

    return { success: true };
  } catch(e) {
    console.error("updateTimeEntry error:", e);
    return { error: e.toString() };
  }
}

// ============================================================
// TEAM DIRECTORY
// ============================================================
function getTeamData(startDate, endDate) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var taSheet = ss.getSheetByName(SHEET_TEAM_ASSIGN);

    if (!taSheet || taSheet.getLastRow() < 2) return { assignments: [], entries: [], cases: [] };

    var assignData  = taSheet.getRange(2, 1, taSheet.getLastRow() - 1, 4).getValues();
    var assignments = assignData
      .filter(function(r){ return r[0] && !isHiddenRosterName(r[0]); })
      .map(function(r){
        return {
          analyst_name:     r[0].toString(),
          supervisor_email: r[1].toString(),
          team_name:        r[2].toString(),
          role:             r[3].toString()
        };
      });

    var visibleAnalysts = {};
    assignments.forEach(function(a){ visibleAnalysts[a.analyst_name] = true; });

    var entries = getTimeEntries(startDate, endDate);
    if (entries.error) return { error: entries.error };

    entries = entries.filter(function(e){
      return !isHiddenRosterName(e.analyst) && !!visibleAnalysts[e.analyst];
    });

    var cases = _readAllCases(startDate, endDate).filter(function(c){
      return !isHiddenRosterName(c.analyst) && !!visibleAnalysts[c.analyst];
    });

    return { assignments: assignments, entries: entries, cases: cases };
  } catch(e) {
    console.error("getTeamData error:", e);
    return { error: e.toString() };
  }
}

// ============================================================
// FRAUD CASES
// ============================================================
function _readAllCases(startDate, endDate) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_CASES);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    var tz   = Session.getScriptTimeZone();

    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      var dateStr;
      if (row[0] instanceof Date) {
        dateStr = Utilities.formatDate(row[0], tz, "yyyy-MM-dd");
      } else {
        dateStr = String(row[0] || "").substring(0, 10);
      }
      if (!dateStr) continue;
      if (dateStr < startDate || dateStr > endDate) continue;

      var handleSec = (typeof row[5] === "number" && row[5] > 0) ? row[5] : null;

      results.push({
        date:           dateStr,
        analyst:        String(row[1] || ""),
        platform:       String(row[2] || ""),
        case_id:        String(row[3] || ""),
        source:         String(row[4] || ""),
        handle_seconds: handleSec,
        solved_at:      String(row[7] || "")
      });
    }
    return results;
  } catch(e) {
    console.error("_readAllCases error:", e);
    return [];
  }
}

function getCaseEntries(startDate, endDate) {
  try {
    return _readAllCases(startDate, endDate).filter(function(c){
      return !isHiddenRosterName(c.analyst);
    });
  } catch(e) {
    console.error("getCaseEntries error:", e);
    return { error: e.toString() };
  }
}

// ============================================================
// CSV IMPORT
// ============================================================
function importCSV(csvText) {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!isPrivileged((getUserRecord(email) || {}).role)) return { error: "Admin/Supervisor only" };

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);

    var lines       = csvText.split("\n");
    var inserted    = 0, skipped = 0;
    var existingIds = {};

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 13, sheet.getLastRow() - 1, 1).getValues()
        .forEach(function(r){ if (r[0]) existingIds[r[0].toString()] = true; });
    }

    var maxId = 0;
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
        .forEach(function(r){ var n = parseInt(r[0]); if (!isNaN(n) && n > maxId) maxId = n; });
    }

    var newRows = [];
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      var cols = parseCSVLine(line);
      if (cols.length < 10) continue;

      var dateStr    = cols[1] || "";
      var project    = cols[3] || "Unknown";
      var analyst    = cols[4] || "";
      var startedStr = cols[7] || "";
      var stoppedStr = cols[8] || "";
      var durStr     = cols[9] || "0:00:00";
      if (!dateStr || !analyst) continue;

      var startTime  = extractTimeFromISO(startedStr);
      var endTime    = extractTimeFromISO(stoppedStr);
      var duration   = parseDuration(durStr);
      var task       = normalizeTaskName(project);
      var sheetRowId = generateRowId(analyst, dateStr, startTime, task);

      if (existingIds[sheetRowId]) { skipped++; continue; }

      maxId++;
      newRows.push([maxId, analyst, "Fetch Rewards", task, duration, dateStr, "Work",
                    startTime, endTime, "", "", "", sheetRowId]);
      existingIds[sheetRowId] = true;
      inserted++;
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 13).setValues(newRows);
    }

    return { success: true, inserted: inserted, skipped: skipped };
  } catch(e) {
    console.error("importCSV error:", e);
    return { error: e.toString() };
  }
}

// ============================================================
// USER MANAGEMENT
// ============================================================
function getAllUsers() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!isPrivileged((getUserRecord(email) || {}).role)) return { error: "Admin/Supervisor only" };

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) return [];

    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
      .filter(function(r){ return r[0] && !isHiddenRosterName(r[2]); })
      .map(function(r){
        return {
          email:        r[0].toString(),
          role:         r[1].toString(),
          display_name: r[2].toString()
        };
      });
  } catch(e) {
    return { error: e.toString() };
  }
}

function upsertUser(email, role, displayName) {
  try {
    var callerEmail = Session.getActiveUser().getEmail();
    if (!isPrivileged((getUserRecord(callerEmail) || {}).role)) return { error: "Admin/Supervisor only" };

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_USERS);
    upsertSheetRow(sheet, email, [email, role, displayName], 1);

    return { success: true };
  } catch(e) {
    return { error: e.toString() };
  }
}

function upsertSheetRow(sheet, keyValue, rowData, keyCol) {
  var found = false;
  if (sheet.getLastRow() > 1) {
    var vals = sheet.getRange(2, keyCol, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0].toString().toLowerCase() === keyValue.toLowerCase()) {
        sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
        found = true;
        break;
      }
    }
  }
  if (!found) sheet.appendRow(rowData);
}

// ============================================================
// HELPER UTILITIES
// ============================================================
function parseCSVLine(line) {
  var result = [], current = "", inQuotes = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (ch === '"') { inQuotes = !inQuotes; }
    else if (ch === ',' && !inQuotes) { result.push(current.trim()); current = ""; }
    else { current += ch; }
  }
  result.push(current.trim());
  return result;
}

function parseDuration(s) {
  var parts = s.split(":");
  if (parts.length === 3) return parseInt(parts[0])*3600 + parseInt(parts[1])*60 + parseInt(parts[2]);
  if (parts.length === 2) return parseInt(parts[0])*3600 + parseInt(parts[1])*60;
  return 0;
}

function extractTimeFromISO(s) {
  if (!s) return null;
  var m = s.match(/T(\d{2}:\d{2}:\d{2})/);
  return m ? m[1] : null;
}

function normalizeTaskName(task) {
  return task.trim().replace(/\s*-?\s*Sweeping/gi, "🧹Sweeping");
}

function generateRowId(analyst, date, startTime, task) {
  return ("csv_" + analyst + "_" + date + "_" + (startTime || "notime") + "_" + task)
    .replace(/[^a-zA-Z0-9_\-]/g, "_")
    .substring(0, 200);
}

// ============================================================
// DEBUG HELPER
// ============================================================
function debugAuth() {
  var email     = Session.getActiveUser().getEmail();
  var effEmail  = Session.getEffectiveUser().getEmail();
  var record    = getUserRecord(email);
  Logger.log("Active user email: "   + email);
  Logger.log("Effective user email: " + effEmail);
  Logger.log("User record: "         + JSON.stringify(record));
}

// ============================================================
// CHROME EXTENSION ENDPOINT
// ============================================================
var EXTENSION_SECRET = "fetch-fraud-squad";

function doPost(e) {
  var corsHeaders = {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type":                 "application/json"
  };

  try {
    var payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch(parseErr) {
      return buildResponse({ error: "Invalid JSON" }, 400, corsHeaders);
    }

    if (payload.secret !== EXTENSION_SECRET) {
      return buildResponse({ error: "Unauthorized" }, 401, corsHeaders);
    }

    if (payload.kind === 'case') {
      return handleCaseRow(payload, corsHeaders);
    }

    var required = ["analyst", "task", "date", "duration"];
    for (var i = 0; i < required.length; i++) {
      if (!payload[required[i]]) {
        return buildResponse({ error: "Missing field: " + required[i] }, 400, corsHeaders);
      }
    }

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);

    var maxId = 0;
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
        .forEach(function(r){ var n = parseInt(r[0]); if (!isNaN(n) && n > maxId) maxId = n; });
    }
    maxId++;

    var sheetRowId = generateRowId(
      payload.analyst,
      payload.date,
      payload.start_time || null,
      payload.task
    );

    if (sheet.getLastRow() > 1) {
      var existingIds = sheet.getRange(2, 13, sheet.getLastRow() - 1, 1).getValues();
      for (var j = 0; j < existingIds.length; j++) {
        if (existingIds[j][0].toString() === sheetRowId) {
          return buildResponse({ success: true, skipped: true, reason: "Duplicate entry" }, 200, corsHeaders);
        }
      }
    }

    sheet.appendRow([
      maxId,
      payload.analyst,
      payload.project    || "Fetch Rewards",
      payload.task,
      Number(payload.duration),
      payload.date,
      payload.category   || "Work",
      payload.start_time || null,
      payload.end_time   || null,
      null, null, null,
      sheetRowId
    ]);

    return buildResponse({ success: true, id: maxId }, 200, corsHeaders);

  } catch(err) {
    console.error("doPost error:", err);
    return buildResponse({ error: err.toString() }, 500, corsHeaders);
  }
}

function buildResponse(data, code, headers) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleCaseRow(payload, corsHeaders) {
  try {
    var required = ["analyst", "platform", "case_id", "date"];
    for (var i = 0; i < required.length; i++) {
      if (!payload[required[i]]) {
        return buildResponse({ error: "Missing field: " + required[i] }, 400, corsHeaders);
      }
    }

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(ss, SHEET_CASES);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Date","Analyst","Platform","Case ID","Source",
                       "Handle (sec)","Handle (min)","Solved At","dedupe_key"]);
      sheet.getRange(1,1,1,9).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var dedupeKey = [payload.date, payload.platform, payload.case_id].join('|');

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var keys = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
      for (var k = 0; k < keys.length; k++) {
        if (keys[k][0] === dedupeKey) {
          return buildResponse({ success: true, skipped: true, reason: "Duplicate case" }, 200, corsHeaders);
        }
      }
    }

    var handleSeconds = (payload.handle_seconds === null || payload.handle_seconds === undefined)
      ? '' : payload.handle_seconds;
    var handleMinutes = (handleSeconds === '') ? '' : Math.round((handleSeconds / 60) * 100) / 100;

    sheet.appendRow([
      payload.date      || '',
      payload.analyst   || '',
      payload.platform  || '',
      payload.case_id   || '',
      payload.source    || '',
      handleSeconds,
      handleMinutes,
      payload.solved_at || '',
      dedupeKey
    ]);

    return buildResponse({ success: true }, 200, corsHeaders);

  } catch(err) {
    console.error("handleCaseRow error:", err);
    return buildResponse({ error: err.toString() }, 500, corsHeaders);
  }
}
