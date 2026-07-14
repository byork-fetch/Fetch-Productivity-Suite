// ============================================================
// PI PRODUCTIVITY SUITE — Google Apps Script Backend
// Code.gs — v3.9 (GitHub Pages / fetch() compatible)
// ============================================================
// All data reads now go through doGet with ?action=... params.
// The dashboard is hosted on GitHub Pages and calls this
// endpoint via fetch() instead of google.script.run.
// doPost is unchanged — extension sync still works as before.
// ============================================================

var SHEET_TIME_ENTRIES = "time_entries";
var SHEET_USERS        = "users";
var SHEET_TEAM_ASSIGN  = "team_assignments";
var SHEET_CASES        = "Cases";
var SPREADSHEET_ID     = "1Kl57TacbVJmTAJTLqJ_vVIFqTkQMxY0vC1ejBULya5M";
var EXTENSION_SECRET   = "fetch-fraud-squad";
// Dashboard secret — used by the GitHub Pages frontend to authenticate
// data fetch requests. Keep this in sync with DASHBOARD_SECRET in index.html.
var DASHBOARD_SECRET   = "fps-dashboard-2024";

var HIDDEN_ROSTER_NAMES = [
  "cassandra buss",
  "alex mendez", "daniel jullian", "daniela lópez", "daniela lopez",
  "daniela suarez", "eddie escamilla", "enrique becerril", "pamela minero",
  "ricardo rico", "tori segura", "mauricio gavito paredes",
  "michael bracamontes", "diego garcía", "diego garcia"
];

function isHiddenRosterName(name) {
  return HIDDEN_ROSTER_NAMES.indexOf(String(name || "").trim().toLowerCase()) !== -1;
}

function isPrivileged(role) {
  var r = (role || "").toLowerCase().trim();
  return r === "admin" || r === "supervisor";
}

// ============================================================
// CORS HEADERS — required for fetch() from GitHub Pages
// ============================================================
function corsHeaders() {
  return {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type":                 "application/json"
  };
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doGet — handles both dashboard data requests and ping
// ============================================================
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  // Ping check (used by extension)
  if (params.ping) {
    return jsonResponse({ ok: true, version: "3.9-server", serverTs: new Date().toISOString() });
  }

  // Data requests from the GitHub Pages dashboard
  var action = params.action || "";
  if (action) {
    // Verify dashboard secret on all data requests
    if (params.secret !== DASHBOARD_SECRET) {
      return jsonResponse({ error: "Unauthorized" });
    }

    try {
      if (action === "getAvailableMonths") {
        return jsonResponse(getAvailableMonths());
      }
      if (action === "getTimeEntries") {
        return jsonResponse(getTimeEntries(params.start, params.end));
      }
      if (action === "getCaseEntries") {
        return jsonResponse(getCaseEntries(params.start, params.end));
      }
      if (action === "getTeamData") {
        return jsonResponse(getTeamData(params.start, params.end));
      }
      if (action === "getAllUsers") {
        return jsonResponse(getAllUsers(params.email));
      }
      if (action === "getUserByEmail") {
        return jsonResponse(getUserByEmail(params.email));
      }
      return jsonResponse({ error: "Unknown action: " + action });
    } catch(err) {
      return jsonResponse({ error: err.toString() });
    }
  }

  // No action — serve the dashboard HTML (fallback, not used when hosted on GitHub Pages)
  return HtmlService.createHtmlOutput("<p>PI Productivity Suite API</p>")
    .setTitle("PI Productivity Suite API");
}

// ============================================================
// USER LOOKUP
// ============================================================
function getUserRecord(email) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        return { email: data[i][0].toString(), role: data[i][1].toString(), display_name: data[i][2].toString() || null };
      }
    }
    return null;
  } catch(e) { return null; }
}

function getUserByEmail(email) {
  try {
    if (!email) return { authenticated: false, error: "No email provided" };
    var domain = (email.split("@")[1] || "").toLowerCase();
    if (domain !== "fetchrewards.com") {
      return { authenticated: false, error: "Please sign in with your Fetch Rewards account" };
    }
    var record      = getUserRecord(email) || {};
    var role        = record.role        || "analyst";
    var displayName = record.display_name || email.split("@")[0];
    return { authenticated: true, email: email, role: role, isSupervisor: isPrivileged(role), displayName: displayName };
  } catch(e) { return { authenticated: false, error: e.toString() }; }
}

// ============================================================
// TIME ENTRIES
// ============================================================
function getTimeEntries(startDate, endDate) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    var headers = ["id","analyst","project","task","duration","date","category","start_time","end_time","edit_reason","edited_at","original_duration","sheet_row_id"];
    var start = new Date(startDate + "T00:00:00"), end = new Date(endDate + "T23:59:59");
    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      var dateVal   = row[5];
      var entryDate = (dateVal instanceof Date) ? dateVal : new Date(dateVal + "T00:00:00");
      if (entryDate >= start && entryDate <= end) {
        var entry = {};
        for (var j = 0; j < headers.length; j++) {
          var val = row[j];
          entry[headers[j]] = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd") : (val === "" ? null : val);
        }
        results.push(entry);
      }
    }
    return results;
  } catch(e) { return { error: e.toString() }; }
}

function getAvailableMonths() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues();
    var months = {};
    for (var i = 0; i < data.length; i++) {
      var val = data[i][0];
      if (!val) continue;
      var dateStr = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd") : val.toString();
      if (dateStr.length >= 7) months[dateStr.substring(0, 7)] = true;
    }
    return Object.keys(months).sort(function(a,b){ return b.localeCompare(a); });
  } catch(e) { return []; }
}

// ============================================================
// TIME ENTRIES — update (called via POST from dashboard)
// ============================================================
function updateTimeEntryData(id, updates, callerEmail) {
  try {
    var record      = getUserRecord(callerEmail) || {};
    var role        = record.role || "analyst";
    var displayName = record.display_name || "";
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
    if (!sheet || sheet.getLastRow() < 2) return { error: "Sheet not found" };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    var rowIndex = -1, entryAnalyst = "";
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) { rowIndex = i + 2; entryAnalyst = data[i][1].toString(); break; }
    }
    if (rowIndex === -1) return { error: "Entry not found" };
    if (!isPrivileged(role) && entryAnalyst !== displayName) return { error: "Permission denied" };
    var existingOriginal = data[rowIndex - 2][11];
    var originalDuration = data[rowIndex - 2][4];
    sheet.getRange(rowIndex, 4).setValue(updates.task || data[rowIndex - 2][3]);
    sheet.getRange(rowIndex, 5).setValue(Number(updates.duration) || data[rowIndex - 2][4]);
    sheet.getRange(rowIndex, 8).setValue(updates.start_time || "");
    sheet.getRange(rowIndex, 9).setValue(updates.end_time   || "");
    sheet.getRange(rowIndex, 10).setValue(updates.edit_reason || "");
    sheet.getRange(rowIndex, 11).setValue(new Date().toISOString());
    if (!existingOriginal) sheet.getRange(rowIndex, 12).setValue(originalDuration);
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
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
      .map(function(r){ return { analyst_name: r[0].toString(), supervisor_email: r[1].toString(), team_name: r[2].toString(), role: r[3].toString() }; });
    var visibleAnalysts = {};
    assignments.forEach(function(a){ visibleAnalysts[a.analyst_name] = true; });
    var entries = getTimeEntries(startDate, endDate);
    if (entries.error) return { error: entries.error };
    entries = entries.filter(function(e){ return !isHiddenRosterName(e.analyst) && !!visibleAnalysts[e.analyst]; });
    var cases = _readAllCases(startDate, endDate).filter(function(c){ return !isHiddenRosterName(c.analyst) && !!visibleAnalysts[c.analyst]; });
    return { assignments: assignments, entries: entries, cases: cases };
  } catch(e) { return { error: e.toString() }; }
}

// ============================================================
// PI CASES
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
      var dateStr = (row[0] instanceof Date) ? Utilities.formatDate(row[0], tz, "yyyy-MM-dd") : String(row[0] || "").substring(0, 10);
      if (!dateStr || dateStr < startDate || dateStr > endDate) continue;
      results.push({ date: dateStr, analyst: String(row[1] || ""), platform: String(row[2] || ""), case_id: String(row[3] || ""), source: String(row[4] || ""), handle_seconds: (typeof row[5] === "number" && row[5] > 0) ? row[5] : null, solved_at: String(row[7] || "") });
    }
    return results;
  } catch(e) { return []; }
}

function getCaseEntries(startDate, endDate) {
  try {
    return _readAllCases(startDate, endDate).filter(function(c){ return !isHiddenRosterName(c.analyst); });
  } catch(e) { return { error: e.toString() }; }
}

// ============================================================
// USER MANAGEMENT
// ============================================================
function getAllUsers(callerEmail) {
  try {
    if (!isPrivileged((getUserRecord(callerEmail) || {}).role)) return { error: "Admin/Supervisor only" };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
      .filter(function(r){ return r[0] && !isHiddenRosterName(r[2]); })
      .map(function(r){ return { email: r[0].toString(), role: r[1].toString(), display_name: r[2].toString() }; });
  } catch(e) { return { error: e.toString() }; }
}

function upsertUserData(email, role, displayName, callerEmail) {
  try {
    if (!isPrivileged((getUserRecord(callerEmail) || {}).role)) return { error: "Admin/Supervisor only" };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_USERS);
    upsertSheetRow(sheet, email, [email, role, displayName], 1);
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

function upsertSheetRow(sheet, keyValue, rowData, keyCol) {
  var found = false;
  if (sheet.getLastRow() > 1) {
    var vals = sheet.getRange(2, keyCol, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0].toString().toLowerCase() === keyValue.toLowerCase()) {
        sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
        found = true; break;
      }
    }
  }
  if (!found) sheet.appendRow(rowData);
}

// ============================================================
// CSV IMPORT
// ============================================================
function importCSVData(csvText, callerEmail) {
  try {
    if (!isPrivileged((getUserRecord(callerEmail) || {}).role)) return { error: "Admin/Supervisor only" };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
    var lines = csvText.split("\n");
    var inserted = 0, skipped = 0, existingIds = {}, maxId = 0;
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 13, sheet.getLastRow() - 1, 1).getValues().forEach(function(r){ if (r[0]) existingIds[r[0].toString()] = true; });
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(r){ var n = parseInt(r[0]); if (!isNaN(n) && n > maxId) maxId = n; });
    }
    var newRows = [];
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      var cols = parseCSVLine(line);
      if (cols.length < 10) continue;
      var dateStr = cols[1]||"", project = cols[3]||"Unknown", analyst = cols[4]||"";
      if (!dateStr || !analyst) continue;
      var startTime = extractTimeFromISO(cols[7]||""), endTime = extractTimeFromISO(cols[8]||"");
      var duration = parseDuration(cols[9]||"0:00:00"), task = normalizeTaskName(project);
      var sheetRowId = generateRowId(analyst, dateStr, startTime, task);
      if (existingIds[sheetRowId]) { skipped++; continue; }
      maxId++;
      newRows.push([maxId, analyst, "Fetch Rewards", task, duration, dateStr, "Work", startTime, endTime, "", "", "", sheetRowId]);
      existingIds[sheetRowId] = true; inserted++;
    }
    if (newRows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 13).setValues(newRows);
    return { success: true, inserted: inserted, skipped: skipped };
  } catch(e) { return { error: e.toString() }; }
}

// ============================================================
// HELPERS
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
function parseDuration(s) { var p=s.split(":"); if(p.length===3)return parseInt(p[0])*3600+parseInt(p[1])*60+parseInt(p[2]); if(p.length===2)return parseInt(p[0])*3600+parseInt(p[1])*60; return 0; }
function extractTimeFromISO(s) { if(!s)return null; var m=s.match(/T(\d{2}:\d{2}:\d{2})/); return m?m[1]:null; }
function normalizeTaskName(task) { return task.trim().replace(/\s*-?\s*Sweeping/gi,"🧹Sweeping"); }
function generateRowId(analyst,date,startTime,task) { return ("csv_"+analyst+"_"+date+"_"+(startTime||"notime")+"_"+task).replace(/[^a-zA-Z0-9_\-]/g,"_").substring(0,200); }

function getOrCreateSheet(ss, name) { var s=ss.getSheetByName(name); if(!s)s=ss.insertSheet(name); return s; }

function debugAuth() {
  Logger.log("Active user: " + Session.getActiveUser().getEmail());
  Logger.log("Effective user: " + Session.getEffectiveUser().getEmail());
}

// ============================================================
// doPost — extension sync + dashboard mutations
// ============================================================
function doPost(e) {
  try {
    var payload;
    try { payload = JSON.parse(e.postData.contents); }
    catch(err) { return jsonResponse({ error: "Invalid JSON" }); }

    // Dashboard mutation actions (updateTimeEntry, upsertUser, importCSV)
    if (payload.dashSecret === DASHBOARD_SECRET) {
      if (payload.action === "updateTimeEntry") {
        return jsonResponse(updateTimeEntryData(payload.id, payload.updates, payload.callerEmail));
      }
      if (payload.action === "upsertUser") {
        return jsonResponse(upsertUserData(payload.email, payload.role, payload.displayName, payload.callerEmail));
      }
      if (payload.action === "importCSV") {
        return jsonResponse(importCSVData(payload.csvText, payload.callerEmail));
      }
      return jsonResponse({ error: "Unknown action" });
    }

    // Extension sync (existing flow — unchanged)
    if (payload.secret !== EXTENSION_SECRET) {
      return jsonResponse({ error: "Unauthorized" });
    }
    if (payload.kind === "case") return handleCaseRow(payload);
    var required = ["analyst","task","date","duration"];
    for (var i = 0; i < required.length; i++) {
      if (!payload[required[i]]) return jsonResponse({ error: "Missing field: " + required[i] });
    }
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
    var maxId = 0;
    if (sheet.getLastRow() > 1) sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().forEach(function(r){var n=parseInt(r[0]);if(!isNaN(n)&&n>maxId)maxId=n;});
    maxId++;
    var sheetRowId = generateRowId(payload.analyst, payload.date, payload.start_time||null, payload.task);
    if (sheet.getLastRow() > 1) {
      var existingIds = sheet.getRange(2,13,sheet.getLastRow()-1,1).getValues();
      for (var j = 0; j < existingIds.length; j++) {
        if (existingIds[j][0].toString() === sheetRowId) return jsonResponse({ success:true, skipped:true, reason:"Duplicate entry" });
      }
    }
    sheet.appendRow([maxId, payload.analyst, payload.project||"Fetch Rewards", payload.task, Number(payload.duration), payload.date, payload.category||"Work", payload.start_time||null, payload.end_time||null, null, null, null, sheetRowId]);
    return jsonResponse({ success: true, id: maxId });
  } catch(err) { return jsonResponse({ error: err.toString() }); }
}

function handleCaseRow(payload) {
  try {
    var required = ["analyst","platform","case_id","date"];
    for (var i = 0; i < required.length; i++) {
      if (!payload[required[i]]) return jsonResponse({ error: "Missing field: " + required[i] });
    }
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(ss, SHEET_CASES);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Date","Analyst","Platform","Case ID","Source","Handle (sec)","Handle (min)","Solved At","dedupe_key"]);
      sheet.getRange(1,1,1,9).setFontWeight("bold"); sheet.setFrozenRows(1);
    }
    var dedupeKey = [payload.date, payload.platform, payload.case_id].join("|");
    if (sheet.getLastRow() > 1) {
      var keys = sheet.getRange(2,9,sheet.getLastRow()-1,1).getValues();
      for (var k = 0; k < keys.length; k++) {
        if (keys[k][0] === dedupeKey) return jsonResponse({ success:true, skipped:true, reason:"Duplicate case" });
      }
    }
    var hs = (payload.handle_seconds===null||payload.handle_seconds===undefined)?'':payload.handle_seconds;
    var hm = hs===''?'':Math.round((hs/60)*100)/100;
    sheet.appendRow([payload.date||'', payload.analyst||'', payload.platform||'', payload.case_id||'', payload.source||'', hs, hm, payload.solved_at||'', dedupeKey]);
    return jsonResponse({ success: true });
  } catch(err) { return jsonResponse({ error: err.toString() }); }
}
