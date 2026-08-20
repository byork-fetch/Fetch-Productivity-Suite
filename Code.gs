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
var SHEET_ATTENTION_REVIEWS = "attention_reviews";
var SHEET_DIGEST_RECIPIENTS = "digest_recipients";
var SPREADSHEET_ID     = "1Kl57TacbVJmTAJTLqJ_vVIFqTkQMxY0vC1ejBULya5M";
var EXTENSION_SECRET   = "fetch-fraud-squad";
// Dashboard secret — used by the GitHub Pages frontend to authenticate
// data fetch requests. Keep this in sync with DASHBOARD_SECRET in index.html.
var DASHBOARD_SECRET   = "fps-dashboard-2024";
// GitHub Pages URL the digest email links out to. "#team" tells index.html
// (see the DOMContentLoaded hash check there) to open straight to Team
// Directory instead of the default Dashboard view.
var DASHBOARD_URL      = "https://byork-fetch.github.io/Fetch-Productivity-Suite/#team";

function isPrivileged(role) {
  var r = (role || "").toLowerCase().trim();
  return r === "admin" || r === "supervisor";
}

function isAdminRole(role) {
  return (role || "").toLowerCase().trim() === "admin";
}

// ============================================================
// CACHING — Sheets reads are the slowest part of every request.
// We cache computed results (not raw sheet data) in Script Cache,
// keyed with a version number so any write anywhere instantly
// invalidates every cached read, rather than us trying to guess
// which specific cache keys a given write affects.
// ============================================================
var CACHE_DEFAULT_TTL = 120; // seconds

function getCacheVersion() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty("cacheVersion") || "0";
}

function bumpCacheVersion() {
  var props = PropertiesService.getScriptProperties();
  var v = parseInt(props.getProperty("cacheVersion") || "0", 10) + 1;
  props.setProperty("cacheVersion", String(v));
}

// Wraps an expensive computeFn() in a cached read. Falls back to
// computing fresh if the cache is unavailable, empty, or the result
// is too large to cache (Script Cache has a 100KB per-key limit) —
// caching is a speed optimization here, never a hard dependency.
function cachedCall(key, ttlSeconds, computeFn) {
  var cache = CacheService.getScriptCache();
  var fullKey = key + ":v" + getCacheVersion();
  try {
    var hit = cache.get(fullKey);
    if (hit) return JSON.parse(hit);
  } catch(e) { /* fall through and compute fresh */ }

  var result = computeFn();
  try {
    var json = JSON.stringify(result);
    if (json.length < 100000) cache.put(fullKey, json, ttlSeconds || CACHE_DEFAULT_TTL);
  } catch(e) { /* not cacheable — fine, just skip caching this result */ }
  return result;
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
      if (action === "getAssignments") {
        return jsonResponse(getAssignmentsList());
      }
      if (action === "getAllUsers") {
        return jsonResponse(getAllUsers(params.email));
      }
      if (action === "getUserByEmail") {
        return jsonResponse(getUserByEmail(params.email));
      }
      if (action === "getReviewedFlags") {
        return jsonResponse(getReviewedFlagKeys());
      }
      if (action === "getReviewedFlagDetails") {
        return jsonResponse(getReviewedFlagDetails());
      }
      if (action === "getDigestRecipients") {
        return jsonResponse(getDigestRecipients(params.email));
      }
      if (action === "getCaseTrends") {
        return jsonResponse(getCaseTrends(params.weeksBack, params.analyst));
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
    return cachedCall("entries:" + startDate + ":" + endDate, 120, function() {
      var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEET_TIME_ENTRIES);
      if (!sheet || sheet.getLastRow() < 2) return [];
      var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
      var headers = ["id","analyst","project","task","duration","date","category","start_time","end_time","edit_reason","edited_at","original_duration","sheet_row_id"];
      var start = new Date(startDate + "T00:00:00"), end = new Date(endDate + "T23:59:59");
      var results = [];
      var skipped = 0;
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        if (!row[0] && !row[1]) continue;
        var dateVal   = row[5];
        var entryDate = (dateVal instanceof Date) ? dateVal : new Date(dateVal + "T00:00:00");
        if (isNaN(entryDate.getTime())) { skipped++; continue; } // malformed date — skip, but track it
        if (entryDate >= start && entryDate <= end) {
          var entry = {};
          for (var j = 0; j < headers.length; j++) {
            var val = row[j];
            entry[headers[j]] = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd") : (val === "" ? null : val);
          }
          results.push(entry);
        }
      }
      if (skipped > 0) Logger.log("getTimeEntries: skipped " + skipped + " row(s) with unparseable dates for range " + startDate + " to " + endDate);
      return results;
    });
  } catch(e) { return { error: e.toString() }; }
}

function getAvailableMonths() {
  try {
    return cachedCall("months", 300, function() {
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
    });
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
    if (!isPrivileged(role) && entryAnalyst !== displayName) {
      // Log this rather than failing silently — a mismatch here is usually
      // "entryAnalyst" (as logged by the extension) not matching
      // "users.display_name" exactly, which is otherwise invisible from the
      // dashboard's point of view (it just shows "Permission denied").
      Logger.log("updateTimeEntryData permission denied — entry analyst='" + entryAnalyst + "', caller displayName='" + displayName + "', caller email=" + callerEmail);
      return { error: "Permission denied" };
    }
    var existingOriginal = data[rowIndex - 2][11];
    var originalDuration = data[rowIndex - 2][4];

    // IMPORTANT: use explicit "was this actually provided" checks instead of
    // `updates.x || fallback`. `||` treats a legitimate value of 0 (e.g. an
    // edit that intentionally zeroes out a duration) or an emptied-out task
    // name as "not provided," and silently falls back to the OLD value while
    // still returning { success: true } — the dashboard shows "Entry
    // updated!" even though nothing changed. That's the bug that was hiding
    // behind reports of edits "not saving."
    var newTask = (updates.task !== undefined && updates.task !== null && updates.task !== "")
      ? updates.task : data[rowIndex - 2][3];
    var newDuration = (updates.duration !== undefined && updates.duration !== null && !isNaN(Number(updates.duration)))
      ? Number(updates.duration) : data[rowIndex - 2][4];

    sheet.getRange(rowIndex, 4).setValue(newTask);
    sheet.getRange(rowIndex, 5).setValue(newDuration);
    sheet.getRange(rowIndex, 8).setValue(updates.start_time || "");
    sheet.getRange(rowIndex, 9).setValue(updates.end_time   || "");
    sheet.getRange(rowIndex, 10).setValue(updates.edit_reason || "");
    sheet.getRange(rowIndex, 11).setValue(new Date().toISOString());

    // Only stamp original_duration the first time this row is edited. Guard
    // against the same falsy-0 trap: if the row's true original duration was
    // ever legitimately 0, "existingOriginal" reads back as 0 (falsy) forever,
    // and every future edit would re-stamp original_duration with whatever
    // the CURRENT (already-edited) duration is — permanently losing the real
    // original and corrupting the "% Changed" column in Recently Edited.
    if (existingOriginal === "" || existingOriginal === null || existingOriginal === undefined) {
      sheet.getRange(rowIndex, 12).setValue(originalDuration);
    }
    bumpCacheVersion();
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

// ============================================================
// TEAM DIRECTORY
// ============================================================
function getTeamData(startDate, endDate) {
  try {
    return cachedCall("team:" + startDate + ":" + endDate, 120, function() {
      var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
      var taSheet = ss.getSheetByName(SHEET_TEAM_ASSIGN);
      if (!taSheet || taSheet.getLastRow() < 2) return { assignments: [], entries: [], cases: [] };
      var assignData  = taSheet.getRange(2, 1, taSheet.getLastRow() - 1, 4).getValues();
      var assignments = assignData
        .filter(function(r){ return r[0]; })
        .map(function(r){ return { analyst_name: r[0].toString(), supervisor_email: r[1].toString(), team_name: r[2].toString(), role: r[3].toString() }; });
      var visibleAnalysts = {};
      assignments.forEach(function(a){ visibleAnalysts[a.analyst_name] = true; });
      var entries = getTimeEntries(startDate, endDate);
      if (entries.error) return { error: entries.error };
      entries = entries.filter(function(e){ return !!visibleAnalysts[e.analyst]; });
      var cases = _readAllCases(startDate, endDate).filter(function(c){ return !!visibleAnalysts[c.analyst]; });
      return { assignments: assignments, entries: entries, cases: cases };
    });
  } catch(e) { return { error: e.toString() }; }
}

// Lightweight, date-range-independent version of the assignments half of
// getTeamData() — just team_assignments rows, no entries/cases join. Used
// by the Dashboard (not just Team Directory) to populate its Team/Role
// filter dropdowns without waiting on a full getTeamData() call tied to
// whatever date range happens to be selected. Same shape as getTeamData()'s
// assignments array so both callers can share client-side logic.
function getAssignmentsList() {
  try {
    return cachedCall("assignments_list", 300, function() {
      var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
      var taSheet = ss.getSheetByName(SHEET_TEAM_ASSIGN);
      if (!taSheet || taSheet.getLastRow() < 2) return [];
      var assignData = taSheet.getRange(2, 1, taSheet.getLastRow() - 1, 4).getValues();
      return assignData
        .filter(function(r){ return r[0]; })
        .map(function(r){ return { analyst_name: r[0].toString(), supervisor_email: r[1].toString(), team_name: r[2].toString(), role: r[3].toString() }; });
    });
  } catch(e) { return []; }
}

// ============================================================
// PI CASES
// ============================================================
function _readAllCases(startDate, endDate) {
  try {
    return cachedCall("cases_raw:" + startDate + ":" + endDate, 120, function() {
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
    });
  } catch(e) { return []; }
}

function getCaseEntries(startDate, endDate) {
  try {
    return _readAllCases(startDate, endDate);
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
      .filter(function(r){ return r[0]; })
      .map(function(r){ return { email: r[0].toString(), role: r[1].toString(), display_name: r[2].toString() }; });
  } catch(e) { return { error: e.toString() }; }
}

function upsertUserData(email, role, displayName, callerEmail) {
  try {
    if (!isPrivileged((getUserRecord(callerEmail) || {}).role)) return { error: "Admin/Supervisor only" };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_USERS);
    upsertSheetRow(sheet, email, [email, role, displayName], 1);
    bumpCacheVersion();
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
    if (inserted > 0) bumpCacheVersion();
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

// ============================================================
// NEEDS ATTENTION — REVIEWED STATUS
// Needs Attention flags are computed fresh client-side on every load (not
// stored records), so "reviewed" status is tracked separately here, keyed
// by a stable flag_key the dashboard builds (analyst+date for the
// zero-cases flag, analyst+platform for the AHT-above-average flag — see
// index.html's computeAttentionFlags). Shared across everyone who opens
// the dashboard, not per-browser.
// ============================================================
function getReviewedFlagKeys() {
  try {
    return cachedCall("reviewedFlagKeys", 60, function() {
      var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEET_ATTENTION_REVIEWS);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
        .map(function(r){ return String(r[0] || ""); })
        .filter(function(k){ return k; });
    });
  } catch(e) { return []; }
}

// Full rows behind those bare keys — analyst, message, who reviewed it, and
// when — for the dashboard's admin-only "Reviewed / Archived" section.
// Deliberately a separate function/cache key from getReviewedFlagKeys()
// rather than changing that function's return shape, since existing callers
// (the Needs Attention exclusion filter) expect a plain array of key strings.
function getReviewedFlagDetails() {
  try {
    return cachedCall("reviewedFlagDetails", 60, function() {
      var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEET_ATTENTION_REVIEWS);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
        .filter(function(r){ return r[0]; })
        .map(function(r){
          return {
            flag_key:    String(r[0] || ""),
            analyst:     String(r[1] || ""),
            message:     String(r[2] || ""),
            reviewed_by: String(r[3] || ""),
            reviewed_at: String(r[4] || "")
          };
        });
    });
  } catch(e) { return []; }
}

function markFlagReviewed(flagKey, analyst, message, callerEmail) {
  try {
    if (!flagKey) return { error: "Missing flagKey" };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(ss, SHEET_ATTENTION_REVIEWS);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["flag_key","analyst","message","reviewed_by","reviewed_at"]);
      sheet.getRange(1,1,1,5).setFontWeight("bold"); sheet.setFrozenRows(1);
    }
    // Dedupe — if this exact flag was already marked reviewed, don't add a
    // second row for it.
    if (sheet.getLastRow() > 1) {
      var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (existing[i][0] === flagKey) return { success: true, alreadyReviewed: true };
      }
    }
    sheet.appendRow([flagKey, analyst || "", message || "", callerEmail || "", new Date().toISOString()]);
    bumpCacheVersion();
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

// Restores an archived flag back into Needs Attention by deleting its row
// in attention_reviews. Flags aren't stored entities in their own right —
// they're recomputed fresh from entries/cases every load — so "restoring"
// one is just removing the "reviewed" marker; computeAttentionFlags() on
// the dashboard will surface it again automatically as long as the
// underlying data (the zero-cases day, the elevated AHT, etc.) still holds.
// Admin-only, matching the dashboard's admin-only visibility of the
// Needs Attention / Archived sections.
function unmarkFlagReviewed(flagKey, callerEmail) {
  try {
    if (!flagKey) return { error: "Missing flagKey" };
    var record = getUserRecord(callerEmail) || {};
    if (!isAdminRole(record.role)) return { error: "Admin only" };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_ATTENTION_REVIEWS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, notFound: true };
    var keys = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (keys[i][0] === flagKey) {
        sheet.deleteRow(i + 2);
        bumpCacheVersion();
        return { success: true };
      }
    }
    return { success: true, notFound: true };
  } catch(e) { return { error: e.toString() }; }
}

function debugAuth() {
  Logger.log("Active user: " + Session.getActiveUser().getEmail());
  Logger.log("Effective user: " + Session.getEffectiveUser().getEmail());
}

// ============================================================
// AUTOMATED WEEKLY EMAIL DIGEST — admin-only opt-in list of recipients,
// stored in the digest_recipients sheet. A time-driven trigger (created
// automatically the first time someone is added, so no manual Apps Script
// setup is needed) fires sendWeeklyDigest() every Saturday in the 10 PM
// hour, which reads the recipient list fresh and mails a summary of the
// current Sun–Sat week to everyone on it. Empty recipient list = trigger
// still fires but sends nothing, which is harmless and self-correcting
// once someone's added.
// ============================================================
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email||"").trim());
}

function getDigestRecipients(callerEmail) {
  try {
    if (!isAdminRole((getUserRecord(callerEmail) || {}).role)) return { error: "Admin only" };
    return cachedCall("digestRecipients", 60, function() {
      var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEET_DIGEST_RECIPIENTS);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
        .filter(function(r){ return r[0]; })
        .map(function(r){ return { email:String(r[0]||""), added_by:String(r[1]||""), added_at:String(r[2]||"") }; });
    });
  } catch(e) { return { error: e.toString() }; }
}

function addDigestRecipient(email, callerEmail) {
  try {
    var record = getUserRecord(callerEmail) || {};
    if (!isAdminRole(record.role)) return { error: "Admin only" };
    email = String(email||"").trim().toLowerCase();
    if (!isValidEmail(email)) return { error: "Please enter a valid email address" };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(ss, SHEET_DIGEST_RECIPIENTS);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["email","added_by","added_at"]);
      sheet.getRange(1,1,1,3).setFontWeight("bold"); sheet.setFrozenRows(1);
    }
    if (sheet.getLastRow() > 1) {
      var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]||"").toLowerCase() === email) return { success: true, alreadyAdded: true };
      }
    }
    sheet.appendRow([email, callerEmail || "", new Date().toISOString()]);
    ensureDigestTriggerExists();
    bumpCacheVersion();
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

function removeDigestRecipient(email, callerEmail) {
  try {
    if (!isAdminRole((getUserRecord(callerEmail) || {}).role)) return { error: "Admin only" };
    email = String(email||"").trim().toLowerCase();
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_DIGEST_RECIPIENTS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, notFound: true };
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]||"").toLowerCase() === email) {
        sheet.deleteRow(i + 2);
        bumpCacheVersion();
        return { success: true };
      }
    }
    return { success: true, notFound: true };
  } catch(e) { return { error: e.toString() }; }
}

// Creates the Saturday 10 PM(ish) trigger exactly once — checked by handler
// function name so re-adding recipients later never creates duplicates.
// atHour(22) fires somewhere in the 22:00–23:00 window (Apps Script's
// time-based triggers are approximate, not to-the-minute), in the script's
// own timezone (Session.getScriptTimeZone()), which satisfies "Saturday
// after 10 PM" without needing an exact-minute guarantee.
function ensureDigestTriggerExists() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendWeeklyDigest") return; // already set up
  }
  ScriptApp.newTrigger("sendWeeklyDigest")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(22)
    .create();
}

function _digestGetWeekStart(d) {
  var s = new Date(d); s.setDate(d.getDate() - d.getDay()); s.setHours(0,0,0,0); return s;
}
function _digestFormatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// Mirrors PLATFORM_TASK_MAP in index.html — kept as a separate server-side
// copy rather than shared, since the dashboard's version lives in client
// JS and isn't reachable from Apps Script.
var DIGEST_PLATFORM_TASK_MAP = { Kount: "Rules Review", Zendesk: "Dispute Review", RADAR: "RADAR" };
var DIGEST_PLATFORM_COLORS = { Kount: "#2f6fed", Zendesk: "#159a6c", RADAR: "#d99e2b" };

// Server-side port of index.html's computeAttentionFlags(), scoped to a
// rolling 42-day window (same span Team Directory defaults to) and
// excluding flags already marked reviewed — so this is the same "open flag"
// count an admin would see if they opened the dashboard right now. Returns
// a count only; the digest doesn't need per-flag detail, just "how many
// need a look."
function _digestComputeOpenAttentionFlagCount() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var taSheet = ss.getSheetByName(SHEET_TEAM_ASSIGN);
    if (!taSheet || taSheet.getLastRow() < 2) return 0;
    var privileged = _digestGetPrivilegedNames();
    var validNames = {};
    taSheet.getRange(2, 1, taSheet.getLastRow() - 1, 1).getValues().forEach(function(r){
      if (r[0] && !privileged[_digestNormName(r[0])]) validNames[r[0].toString()] = true;
    });

    var end = new Date(), start = new Date(end.getTime() - 42*86400000);
    var startStr = _digestFormatDate(start), endStr = _digestFormatDate(end);

    var entriesRaw = getTimeEntries(startStr, endStr);
    var entries = (Array.isArray(entriesRaw) ? entriesRaw : []).filter(function(e){ return validNames[e.analyst]; });
    var cases = _readAllCases(startStr, endStr).filter(function(c){ return validNames[c.analyst]; });

    var caseTaskNames = Object.keys(DIGEST_PLATFORM_TASK_MAP).map(function(k){ return DIGEST_PLATFORM_TASK_MAP[k]; });
    var casesByAnalystDate = {};
    cases.forEach(function(c){ var k=c.analyst+"|"+c.date; casesByAnalystDate[k]=(casesByAnalystDate[k]||0)+1; });
    var loggedCaseDays = {};
    entries.forEach(function(e){
      if (caseTaskNames.indexOf(e.task)===-1) return;
      if (!(e.duration>0)) return;
      loggedCaseDays[e.analyst+"|"+e.date] = true;
    });
    var flagKeys = [];
    Object.keys(loggedCaseDays).forEach(function(key){
      if (!casesByAnalystDate[key]) {
        var parts = key.split("|");
        flagKeys.push("zerocases:"+parts[0]+":"+parts[1]);
      }
    });

    ["Kount","Zendesk","RADAR"].forEach(function(platform){
      var teamTimed = cases.filter(function(c){ return c.platform===platform && typeof c.handle_seconds==="number" && c.handle_seconds>0; });
      if (teamTimed.length < 5) return;
      var teamAvg = teamTimed.reduce(function(a,c){return a+c.handle_seconds;},0)/teamTimed.length;
      var byAnalyst = {};
      teamTimed.forEach(function(c){ (byAnalyst[c.analyst]=byAnalyst[c.analyst]||[]).push(c.handle_seconds); });
      Object.keys(byAnalyst).forEach(function(name){
        var arr = byAnalyst[name];
        if (arr.length < 3) return;
        var avg = arr.reduce(function(a,b){return a+b;},0)/arr.length;
        if (avg > teamAvg*1.5) flagKeys.push("aht:"+name+":"+platform);
      });
    });

    var reviewed = {};
    getReviewedFlagKeys().forEach(function(k){ reviewed[k]=true; });
    return flagKeys.filter(function(k){ return !reviewed[k]; }).length;
  } catch(e) { return 0; }
}

// ============================================================
// CASE TRENDS (dashboard) — powers the "Case Trends (Last 12 Weeks)" chart.
// Aggregates weekly case volume + avg AHT server-side and returns only the
// small per-week summary rows, instead of the old approach of shipping
// every individual case for the whole window to the browser and bucketing
// there. At real-world case volumes that raw payload easily runs into the
// megabytes, which is both slow over the wire and too large for
// CacheService to cache at all (100KB per-key limit) — so every chart open
// was redoing a full uncached sheet scan. The aggregated result here is a
// few dozen small objects regardless of case volume, so it comfortably
// fits in cache and stays fast even as the Cases sheet keeps growing.
// analystFilter mirrors the dashboard's own case-scoping rule: a specific
// analyst name narrows to just them; "all" (or omitted) excludes admin/
// supervisor accounts from the aggregate, same as the "All Analysts" view
// does elsewhere via privilegedNames.
function getCaseTrends(weeksBack, analystFilter) {
  try {
    weeksBack = parseInt(weeksBack, 10) || 12;
    analystFilter = analystFilter || "all";
    return cachedCall("caseTrends:" + weeksBack + ":" + analystFilter, 300, function() {
      var now = new Date();
      var thisWeekStart = _digestGetWeekStart(now);
      var weeks = [];
      for (var i = weeksBack - 1; i >= 0; i--) {
        var ws = new Date(thisWeekStart.getTime() - i*7*86400000);
        var we = new Date(ws.getTime() + 6*86400000);
        if (we > now) we = now; // cap the current week's end at "now" rather than querying into the future
        weeks.push({ start: _digestFormatDate(ws), end: _digestFormatDate(we) });
      }
      var overallStart = weeks[0].start, overallEnd = weeks[weeks.length-1].end;
      var cases = _readAllCases(overallStart, overallEnd);

      if (analystFilter !== "all") {
        cases = cases.filter(function(c){ return c.analyst === analystFilter; });
      } else {
        var privileged = {};
        var usersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
        if (usersSheet && usersSheet.getLastRow() > 1) {
          usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 3).getValues().forEach(function(r){
            if (isPrivileged(r[1])) privileged[String(r[2]||"")] = true;
          });
        }
        cases = cases.filter(function(c){ return !privileged[c.analyst]; });
      }

      return weeks.map(function(w){
        var inWeek = cases.filter(function(c){ return c.date >= w.start && c.date <= w.end; });
        var ahtSecs = inWeek.filter(function(c){ return typeof c.handle_seconds==="number" && c.handle_seconds>0; }).map(function(c){ return c.handle_seconds; });
        return {
          start: w.start, end: w.end, count: inWeek.length,
          avgAhtMin: ahtSecs.length ? Math.round(ahtSecs.reduce(function(a,b){return a+b;},0)/ahtSecs.length/60*10)/10 : null
        };
      });
    });
  } catch(e) { return { error: e.toString() }; }
}

// Case volume + avg AHT for each of the last `weeksBack` Sunday-start weeks,
// oldest first, ending with the current (possibly partial) week. Used for
// the digest's week-over-week trend table.
function _digestGetWeeklyStats(weeksBack) {
  var now = new Date();
  var thisWeekStart = _digestGetWeekStart(now);
  var weeks = [];
  for (var i = weeksBack - 1; i >= 0; i--) {
    var ws = new Date(thisWeekStart.getTime() - i*7*86400000);
    var we = new Date(ws.getTime() + 6*86400000);
    if (we > now) we = now; // cap the current week's end at "now" rather than querying into the future
    var wsStr = _digestFormatDate(ws), weStr = _digestFormatDate(we);
    var cases = _readAllCases(wsStr, weStr);
    var ahtSecs = cases.filter(function(c){ return typeof c.handle_seconds==="number" && c.handle_seconds>0; }).map(function(c){ return c.handle_seconds; });
    var avgAhtMin = ahtSecs.length ? Math.round(ahtSecs.reduce(function(a,b){return a+b;},0)/ahtSecs.length/60*10)/10 : null;
    weeks.push({ start: wsStr, end: weStr, count: cases.length, avgAhtMin: avgAhtMin });
  }
  return weeks;
}

// Normalizes a name for matching, same convention used throughout for
// as normName() in index.html — trims + lowercases so a stray space or a
// capitalization mismatch between users.display_name and Cases.analyst
// doesn't silently break the exclusion.
function _digestNormName(s) { return String(s||"").trim().toLowerCase(); }

// Mirrors the dashboard's computePrivilegedNames()/privilegedNames exclusion:
// admins and supervisors (Bronte included) are always fully visible if
// looked up by name directly, but never counted in team-wide aggregates —
// case totals, AHT, top analysts, idle time. The digest previously only
// excluded hidden-roster rows, so privileged accounts were silently still
// being counted here even though every other aggregate view excludes them.
function _digestGetPrivilegedNames() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) return {};
    var set = {};
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function(r){
      if (isPrivileged(r[1])) set[_digestNormName(r[2])] = true;
    });
    return set;
  } catch(e) { return {}; }
}

// Server-side port of index.html's parseEntryTimeMs/computeIdleGaps — those
// only existed client-side before, so the digest previously had no way to
// surface idle time at all. Same 10-minute threshold and same UTC-time
// parsing convention as the dashboard (start_time/end_time are stored as
// bare "HH:MM:SS" UTC strings unless they already contain a full ISO "T").
var DIGEST_IDLE_GAP_THRESHOLD_SEC = 600;

function _digestParseEntryTimeMs(date, utcTime) {
  if (!utcTime) return null;
  try {
    var d = String(utcTime).indexOf("T") !== -1 ? new Date(utcTime) : new Date(date + "T" + utcTime + "Z");
    var t = d.getTime();
    return isNaN(t) ? null : t;
  } catch(e) { return null; }
}

function _digestComputeIdleGaps(entries) {
  var byKey = {};
  entries.forEach(function(e){
    var startMs = _digestParseEntryTimeMs(e.date, e.start_time), endMs = _digestParseEntryTimeMs(e.date, e.end_time);
    if (startMs == null || endMs == null) return;
    var key = e.analyst + "|" + e.date;
    (byKey[key] = byKey[key] || []).push({ startMs: startMs, endMs: endMs, analyst: e.analyst, date: e.date });
  });
  var gaps = [];
  Object.keys(byKey).forEach(function(key){
    var list = byKey[key].sort(function(a,b){ return a.startMs - b.startMs; });
    for (var i = 0; i < list.length - 1; i++) {
      var gapSec = Math.round((list[i+1].startMs - list[i].endMs) / 1000);
      if (gapSec > DIGEST_IDLE_GAP_THRESHOLD_SEC) {
        gaps.push({ analyst: list[i].analyst, date: list[i].date, gapSec: gapSec });
      }
    }
  });
  return gaps;
}

// Total idle minutes + top 3 analysts by idle time this week, for the
// digest's Idle Time section. entries should already be scoped to the
// digest week and hidden-roster-filtered (sendWeeklyDigest does both before
// calling this).
function _digestSummarizeIdleGaps(entries) {
  var gaps = _digestComputeIdleGaps(entries);
  var totalSec = gaps.reduce(function(a,g){ return a + g.gapSec; }, 0);
  var byAnalyst = {};
  gaps.forEach(function(g){ byAnalyst[g.analyst] = (byAnalyst[g.analyst] || 0) + g.gapSec; });
  var topOffenders = Object.keys(byAnalyst)
    .map(function(name){ return { name: name, sec: byAnalyst[name] }; })
    .sort(function(a,b){ return b.sec - a.sec; })
    .slice(0, 3);
  return { gapCount: gaps.length, totalSec: totalSec, topOffenders: topOffenders };
}

// This-week vs prior-full-week avg AHT per platform, for the "vs last week"
// arrows on the By Platform table. Prior week is the last complete Sun–Sat
// week before the digest's (possibly partial) current week, so a Tuesday
// send is compared against a fair full week rather than another partial one.
function _digestComputePlatformTrend(weekStart, privileged) {
  var priorEnd = new Date(weekStart.getTime() - 86400000);
  var priorStart = new Date(priorEnd.getTime() - 6*86400000);
  var priorCases = _readAllCases(_digestFormatDate(priorStart), _digestFormatDate(priorEnd))
    .filter(function(c){ return !privileged[_digestNormName(c.analyst)]; });
  var byPlatform = { Kount: [], Zendesk: [], RADAR: [] };
  priorCases.forEach(function(c){
    if (byPlatform[c.platform] && typeof c.handle_seconds === "number" && c.handle_seconds > 0) byPlatform[c.platform].push(c.handle_seconds);
  });
  var result = {};
  ["Kount","Zendesk","RADAR"].forEach(function(p){
    var arr = byPlatform[p];
    result[p] = arr.length ? Math.round(arr.reduce(function(a,b){return a+b;},0)/arr.length) : null;
  });
  return result;
}

// The function the Saturday trigger actually calls. No callerEmail here —
// it runs unattended, so recipient-list access bypasses the admin-only
// gate getDigestRecipients() normally enforces (there's no "caller" to
// check) and reads the sheet directly instead.
function sendWeeklyDigest() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var recipSheet = ss.getSheetByName(SHEET_DIGEST_RECIPIENTS);
    if (!recipSheet || recipSheet.getLastRow() < 2) return; // no one signed up yet — nothing to do
    var recipients = recipSheet.getRange(2, 1, recipSheet.getLastRow() - 1, 1).getValues()
      .map(function(r){ return String(r[0]||"").trim(); })
      .filter(function(e){ return isValidEmail(e); });
    if (!recipients.length) return;

    var now = new Date();
    var weekStart = _digestGetWeekStart(now);
    var weekStartStr = _digestFormatDate(weekStart), todayStr = _digestFormatDate(now);

    var privileged = _digestGetPrivilegedNames();
    var cases = _readAllCases(weekStartStr, todayStr).filter(function(c){ return !privileged[_digestNormName(c.analyst)]; });
    var byPlatform = { Kount:{count:0,ahtSecs:[]}, Zendesk:{count:0,ahtSecs:[]}, RADAR:{count:0,ahtSecs:[]} };
    var byAnalyst = {};
    cases.forEach(function(c){
      if (byPlatform[c.platform]) {
        byPlatform[c.platform].count++;
        if (typeof c.handle_seconds==="number" && c.handle_seconds>0) byPlatform[c.platform].ahtSecs.push(c.handle_seconds);
      }
      byAnalyst[c.analyst] = (byAnalyst[c.analyst]||0) + 1;
    });
    var topAnalysts = Object.keys(byAnalyst).map(function(name){ return {name:name, count:byAnalyst[name]}; })
      .sort(function(a,b){ return b.count-a.count; }).slice(0,3);

    function avgMin(secs) { if (!secs.length) return null; return Math.round(secs.reduce(function(a,b){return a+b;},0)/secs.length/60*10)/10; }
    function esc(s) { return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
    function fmtMinShort(sec) {
      if (!sec) return "0m";
      var h = Math.floor(sec/3600), m = Math.round((sec%3600)/60);
      return h > 0 ? (h+"h "+m+"m") : (m+"m");
    }

    // Prior-week AHT per platform, for the ▲/▼ context next to this week's number.
    var priorAht = _digestComputePlatformTrend(weekStart, privileged);

    var platformRows = ["Kount","Zendesk","RADAR"].map(function(p){
      var d = byPlatform[p];
      var curAvg = avgMin(d.ahtSecs);
      var curSec = curAvg != null ? curAvg*60 : null;
      var priorSec = priorAht[p];
      var deltaHtml = "";
      if (curSec != null && priorSec != null && priorSec > 0) {
        var diff = curSec - priorSec;
        if (Math.abs(diff) >= 5) { // ignore sub-5-second noise
          var pct = Math.round(Math.abs(diff)/priorSec*100);
          var faster = diff < 0; // lower AHT is better
          deltaHtml = ' <span style="font-size:11px;font-weight:600;color:'+(faster?'#1e8a5a':'#c0392b')+'">'+(faster?'▼':'▲')+' '+pct+'%</span>';
        }
      }
      var dotColor = DIGEST_PLATFORM_COLORS[p] || "#888";
      return '<tr>'
        + '<td style="padding:10px 12px;border-bottom:1px solid #eee"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:'+dotColor+';margin-right:8px"></span>'+p+'</td>'
        + '<td style="padding:10px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:600">'+d.count+'</td>'
        + '<td style="padding:10px 12px;border-bottom:1px solid #eee;text-align:right">'+(curAvg!=null?curAvg+' min':'—')+deltaHtml+'</td>'
        + '</tr>';
    }).join("");

    var topRows = topAnalysts.length ? topAnalysts.map(function(a,i){
      var medal = i===0 ? "🥇" : i===1 ? "🥈" : "🥉";
      return '<tr><td style="padding:8px 12px;border-bottom:1px solid #eee">'+medal+' '+esc(a.name)+'</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:600">'+a.count+' cases</td></tr>';
    }).join("") : '<tr><td style="padding:8px 12px;color:#888" colspan="2">No cases logged this week</td></tr>';

    var attentionCount = _digestComputeOpenAttentionFlagCount();
    var attentionBanner = attentionCount > 0
      ? '<div style="background:#fff8e6;border:1px solid #f5deb3;color:#7a5b00;padding:12px 16px;border-radius:10px;font-weight:600;font-size:13px;margin:0 0 12px">⚠️ '+attentionCount+' Needs Attention flag'+(attentionCount===1?'':'s')+' awaiting review on the dashboard</div>'
      : '<div style="background:#eaf7ee;border:1px solid #bfe6c9;color:#1e6b34;padding:12px 16px;border-radius:10px;font-weight:600;font-size:13px;margin:0 0 12px">✅ No open Needs Attention flags</div>';
    var reviewButton = '<div style="text-align:center;margin:0 0 20px">'
      + '<a href="'+DASHBOARD_URL+'" style="display:inline-block;background:linear-gradient(90deg,#e35c3c 0%,#c23d6e 100%);color:#fff;text-decoration:none;font-weight:700;font-size:13px;padding:10px 24px;border-radius:8px">Click here to review →</a>'
      + '</div>';

    var weeklyStats = _digestGetWeeklyStats(4);
    var trendRows = weeklyStats.map(function(w, idx){
      var deltaStr = "";
      var isPartialCurrent = (idx === weeklyStats.length - 1) && (todayStr !== _digestFormatDate(new Date(weekStart.getTime()+6*86400000)));
      if (idx > 0 && !isPartialCurrent) {
        var prev = weeklyStats[idx-1];
        var diff = w.count - prev.count;
        if (diff !== 0) {
          var arrow = diff > 0 ? "▲" : "▼";
          var pct = prev.count ? Math.round(Math.abs(diff)/prev.count*100) : null;
          deltaStr = ' <span style="color:'+(diff>0?'#1e8a3a':'#c0392b')+';font-size:11px;font-weight:600">'+arrow+(pct!=null?' '+pct+'%':'')+'</span>';
        }
      }
      var label = w.start+' – '+w.end + (isPartialCurrent ? ' <span style="color:#999;font-weight:400;font-size:11px">(partial)</span>' : '');
      return '<tr><td style="padding:8px 12px;border-bottom:1px solid #eee">'+label+'</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right">'+w.count+deltaStr+'</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right">'+(w.avgAhtMin!=null?w.avgAhtMin+' min':'—')+'</td></tr>';
    }).join("");

    // Idle Time — ported from the dashboard's client-side idle-gap logic so
    // the digest can flag it without anyone opening the dashboard first.
    var weekEntriesRaw = getTimeEntries(weekStartStr, todayStr);
    var weekEntries = (Array.isArray(weekEntriesRaw) ? weekEntriesRaw : []).filter(function(e){ return !privileged[_digestNormName(e.analyst)]; });
    var idleSummary = _digestSummarizeIdleGaps(weekEntries);
    var idleSection = "";
    if (idleSummary.gapCount > 0) {
      var idleRows = idleSummary.topOffenders.map(function(o){
        return '<tr><td style="padding:8px 12px;border-bottom:1px solid #eee">'+esc(o.name)+'</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:600">'+fmtMinShort(o.sec)+'</td></tr>';
      }).join("");
      idleSection = '<h3 style="margin:24px 0 4px;font-size:14px;color:#333">⏱️ Idle Time Gaps</h3>'
        + '<p style="color:#888;font-size:12px;margin:0 0 8px">'+idleSummary.gapCount+' unaccounted gap'+(idleSummary.gapCount===1?'':'s')+' of 10+ min this week, totaling '+fmtMinShort(idleSummary.totalSec)+'</p>'
        + '<table style="border-collapse:collapse;width:100%;font-size:13px"><tbody>'+idleRows+'</tbody></table>';
    } else {
      idleSection = '<h3 style="margin:24px 0 4px;font-size:14px;color:#333">⏱️ Idle Time Gaps</h3>'
        + '<p style="color:#888;font-size:12px;margin:0">No unaccounted 10+ min gaps this week. 👍</p>';
    }

    var totalCases = cases.length;
    var gradientBg = "background:#d85a30;background:linear-gradient(135deg,#e35c3c 0%,#c23d6e 100%)";
    var html = '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;color:#222;max-width:560px;margin:0 auto">'
      + '<div style="'+gradientBg+';border-radius:14px 14px 0 0;padding:24px 24px 20px">'
      +   '<h1 style="color:#fff;font-size:19px;font-weight:700;margin:0 0 4px">PI Productivity Suite, Weekly Digest</h1>'
      +   '<p style="color:rgba(255,255,255,.85);font-size:13px;margin:0">'+weekStartStr+' – '+todayStr+' · '+totalCases+' total cases</p>'
      + '</div>'
      + '<div style="border:1px solid #eee;border-top:none;border-radius:0 0 14px 14px;padding:20px 24px 24px">'
      + attentionBanner
      + reviewButton
      + '<h3 style="margin:0 0 6px;font-size:14px;color:#333">By Platform (This Week)</h3>'
      + '<table style="border-collapse:collapse;width:100%;font-size:14px"><thead><tr><th style="text-align:left;padding:6px 12px;border-bottom:2px solid #333;font-size:12px;color:#666">Platform</th><th style="text-align:right;padding:6px 12px;border-bottom:2px solid #333;font-size:12px;color:#666">Cases</th><th style="text-align:right;padding:6px 12px;border-bottom:2px solid #333;font-size:12px;color:#666">Avg AHT vs last wk</th></tr></thead><tbody>'
      + platformRows + '</tbody></table>'
      + '<h3 style="margin:22px 0 6px;font-size:14px;color:#333">4-Week Trend</h3>'
      + '<table style="border-collapse:collapse;width:100%;font-size:13px"><thead><tr><th style="text-align:left;padding:6px 12px;border-bottom:2px solid #333;font-size:12px;color:#666">Week</th><th style="text-align:right;padding:6px 12px;border-bottom:2px solid #333;font-size:12px;color:#666">Cases</th><th style="text-align:right;padding:6px 12px;border-bottom:2px solid #333;font-size:12px;color:#666">Avg AHT</th></tr></thead><tbody>'
      + trendRows + '</tbody></table>'
      + idleSection
      + '<h3 style="margin:22px 0 6px;font-size:14px;color:#333">Top Analysts This Week</h3>'
      + '<table style="border-collapse:collapse;width:100%;font-size:13px"><tbody>' + topRows + '</tbody></table>'
      + '<p style="color:#999;font-size:11px;margin-top:22px;padding-top:14px;border-top:1px solid #eee">Sent automatically every Saturday.</p>'
      + '</div>'
      + '</div>';

    MailApp.sendEmail({
      to: recipients.join(","),
      subject: "PI Productivity Suite, Weekly Digest (" + weekStartStr + " – " + todayStr + ")",
      htmlBody: html
    });
  } catch(e) {
    Logger.log("sendWeeklyDigest failed: " + e.toString());
  }
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
      if (payload.action === "markFlagReviewed") {
        return jsonResponse(markFlagReviewed(payload.flagKey, payload.analyst, payload.message, payload.callerEmail));
      }
      if (payload.action === "unmarkFlagReviewed") {
        return jsonResponse(unmarkFlagReviewed(payload.flagKey, payload.callerEmail));
      }
      if (payload.action === "upsertUser") {
        return jsonResponse(upsertUserData(payload.email, payload.role, payload.displayName, payload.callerEmail));
      }
      if (payload.action === "importCSV") {
        return jsonResponse(importCSVData(payload.csvText, payload.callerEmail));
      }
      if (payload.action === "addDigestRecipient") {
        return jsonResponse(addDigestRecipient(payload.email, payload.callerEmail));
      }
      if (payload.action === "removeDigestRecipient") {
        return jsonResponse(removeDigestRecipient(payload.email, payload.callerEmail));
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
    bumpCacheVersion();
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
    bumpCacheVersion();
    return jsonResponse({ success: true });
  } catch(err) { return jsonResponse({ error: err.toString() }); }
}
