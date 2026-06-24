// ─────────────────────────────────────────────
// Drone Fleet Manager — Google Apps Script v2
// Supports: Operator App + Technician App
// ─────────────────────────────────────────────

const SHEET_NAME = "Drones";

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  // CORS headers for PWA access
  const params = e.parameter;
  const action = params.action;
  try {
    let result;
    if      (action === "getDrones")       result = getDrones();
    else if (action === "updateDrone")     result = updateDrone(params.droneId, params.status, params.reason || "", params.notes || "");
    else if (action === "getFailsByDate")  result = getFailsByDate();
    else if (action === "batchUpdate")     result = batchUpdate(params.droneIds, params.status);
    else if (action === "autoTransition")  result = autoTransition();
    else result = { error: "Unknown action: " + action };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── SHARED: find column index ──
function findCol(headers, candidates) {
  for (const c of candidates) {
    const idx = headers.findIndex(h => h === c || h.includes(c));
    if (idx >= 0) return idx;
  }
  return -1;
}

// ── SHARED: get sheet + headers + col map ──
function getSheetData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found`);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col = {
    id:       findCol(headers, ["droneid", "drone id", "id"]),
    name:     findCol(headers, ["drone name", "dronename", "name"]),
    status:   findCol(headers, ["status"]),
    reason:   findCol(headers, ["reason", "fail reason"]),
    notes:    findCol(headers, ["notes", "note", "remarks"]),
    updated:  findCol(headers, ["updated", "last updated", "timestamp"]),
    fixed_at: findCol(headers, ["fixed_at", "fixed at", "fixedat"]),
  };
  return { sheet, data, headers, col };
}

// ── OPERATOR: get all drones ──
function getDrones() {
  const { data, col } = getSheetData();
  if (data.length < 2) return { drones: [] };
  const drones = data.slice(1)
    .filter(row => row[col.name >= 0 ? col.name : 0])
    .map((row, i) => ({
      rowIndex: i + 2,
      id:       col.id       >= 0 ? String(row[col.id]).trim()     : `D${String(i+1).padStart(2,"0")}`,
      name:     col.name     >= 0 ? String(row[col.name]).trim()   : String(row[0]),
      status:   col.status   >= 0 ? String(row[col.status]).trim() : "Unknown",
      reason:   col.reason   >= 0 ? String(row[col.reason]).trim() : "",
      notes:    col.notes    >= 0 ? String(row[col.notes]).trim()  : "",
      updated:  col.updated  >= 0 ? formatDate(row[col.updated])   : "",
      fixed_at: col.fixed_at >= 0 ? formatDate(row[col.fixed_at])  : "",
    }));
  return { drones };
}

// ── OPERATOR: update single drone ──
function updateDrone(droneId, status, reason, notes) {
  const { sheet, data, col } = getSheetData();
  const rowIndex = findDroneRow(data, col, droneId);
  if (rowIndex === -1) return { error: `Drone "${droneId}" not found` };

  const now = new Date();
  if (col.status   >= 0) sheet.getRange(rowIndex, col.status   + 1).setValue(status);
  if (col.reason   >= 0) sheet.getRange(rowIndex, col.reason   + 1).setValue(reason);
  if (col.notes    >= 0) sheet.getRange(rowIndex, col.notes    + 1).setValue(notes);
  if (col.updated  >= 0) sheet.getRange(rowIndex, col.updated  + 1).setValue(now);
  // clear fixed_at if moving back to Fail/WIP; set it if Good
  if (col.fixed_at >= 0) {
    sheet.getRange(rowIndex, col.fixed_at + 1).setValue(status === "Good" ? now : "");
  }
  return { success: true };
}

// ── TECHNICIAN: get fails grouped by date ──
function getFailsByDate() {
  const { data, col } = getSheetData();
  if (data.length < 2) return { dates: [] };

  // Group rows by date (from "updated" column)
  const dateMap = {}; // { "2026-06-24": { fail:0, wip:0, good:0, drones:[] } }

  data.slice(1).forEach((row, i) => {
    const status = col.status >= 0 ? String(row[col.status]).trim() : "";
    const updatedRaw = col.updated >= 0 ? row[col.updated] : null;
    if (!updatedRaw || !status) return;

    // Only include Fail and WIP drones
    if (status !== "Fail" && status !== "Work In Progress") return;

    const dateKey = toDateKey(updatedRaw); // "2026-06-24"
    if (!dateKey) return;

    if (!dateMap[dateKey]) dateMap[dateKey] = { fail: 0, wip: 0, drones: [] };
    if (status === "Fail")             dateMap[dateKey].fail++;
    if (status === "Work In Progress") dateMap[dateKey].wip++;

    dateMap[dateKey].drones.push({
      rowIndex: i + 2,
      id:       col.id      >= 0 ? String(row[col.id]).trim()     : `D${String(i+1).padStart(2,"0")}`,
      name:     col.name    >= 0 ? String(row[col.name]).trim()   : String(row[0]),
      status,
      reason:   col.reason  >= 0 ? String(row[col.reason]).trim() : "",
      updated:  formatDate(updatedRaw),
      fixed_at: col.fixed_at >= 0 ? formatDate(row[col.fixed_at]) : "",
    });
  });

  // Sort dates newest first
  const dates = Object.entries(dateMap)
    .sort((a, b) => b[0].localeCompare(a[0]))
    .map(([date, val]) => ({
      date,
      displayDate: formatDisplayDate(date),
      fail:  val.fail,
      wip:   val.wip,
      drones: val.drones.sort((a, b) => a.id.localeCompare(b.id)),
    }));

  // Total counts across all dates
  const totalFail = dates.reduce((s, d) => s + d.fail, 0);
  const totalWip  = dates.reduce((s, d) => s + d.wip,  0);

  return { dates, totalFail, totalWip };
}

// ── TECHNICIAN: batch update multiple drones ──
function batchUpdate(droneIdsJson, status) {
  const droneIds = JSON.parse(droneIdsJson || "[]");
  if (!droneIds.length) return { error: "No drone IDs provided" };

  const { sheet, data, col } = getSheetData();
  const now = new Date();
  let updated = 0;
  const errors = [];

  droneIds.forEach(droneId => {
    const rowIndex = findDroneRow(data, col, droneId);
    if (rowIndex === -1) { errors.push(`Not found: ${droneId}`); return; }
    if (col.status   >= 0) sheet.getRange(rowIndex, col.status   + 1).setValue(status);
    if (col.updated  >= 0) sheet.getRange(rowIndex, col.updated  + 1).setValue(now);
    if (col.fixed_at >= 0) sheet.getRange(rowIndex, col.fixed_at + 1).setValue(status === "Good" ? now : "");
    updated++;
  });

  return { success: true, updated, errors };
}

// ── AUTO-TRANSITION: Fail → WIP at 11:59 PM ──
// Set this as a daily time-based trigger in Apps Script
function autoTransition() {
  const { sheet, data, col } = getSheetData();
  const now     = new Date();
  const today   = toDateKey(now);
  const wip     = "Work In Progress";
  let   count   = 0;

  data.slice(1).forEach((row, i) => {
    const status     = col.status  >= 0 ? String(row[col.status]).trim()  : "";
    const updatedRaw = col.updated >= 0 ? row[col.updated]                : null;
    if (status !== "Fail" || !updatedRaw) return;

    const rowDate = toDateKey(updatedRaw);
    // Flip to WIP if drone was marked Fail today and is still Fail
    if (rowDate === today) {
      const rowIndex = i + 2;
      if (col.status  >= 0) sheet.getRange(rowIndex, col.status  + 1).setValue(wip);
      if (col.updated >= 0) sheet.getRange(rowIndex, col.updated + 1).setValue(now);
      count++;
    }
  });

  return { success: true, transitioned: count };
}

// ── HELPERS ──
function findDroneRow(data, col, droneId) {
  for (let i = 1; i < data.length; i++) {
    const rowId   = col.id   >= 0 ? String(data[i][col.id]).trim()   : "";
    const rowName = col.name >= 0 ? String(data[i][col.name]).trim() : String(data[i][0]).trim();
    if (rowId === droneId || rowName === droneId) return i + 1;
  }
  return -1;
}

function toDateKey(val) {
  if (!val) return null;
  try {
    const d = val instanceof Date ? val : new Date(val);
    if (isNaN(d)) return null;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  } catch(e) { return null; }
}

function formatDate(val) {
  if (!val) return "";
  try {
    const d = val instanceof Date ? val : new Date(val);
    if (isNaN(d)) return String(val);
    return d.toLocaleString("en-IN", { day:"2-digit", month:"short", year:"numeric", hour:"2-digit", minute:"2-digit" });
  } catch(e) { return String(val); }
}

function formatDisplayDate(dateKey) {
  // "2026-06-24" → "24 Jun 2026"
  try {
    const [y, m, d] = dateKey.split("-");
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    return `${d} ${months[parseInt(m)-1]} ${y}`;
  } catch(e) { return dateKey; }
}

// ── HOW TO SET UP AUTO-TRANSITION TRIGGER ──
// In Apps Script editor:
// 1. Click "Triggers" (clock icon on left)
// 2. Add trigger → function: autoTransition
// 3. Time-based → Day timer → 11 PM to midnight
// 4. Save
