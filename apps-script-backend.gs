/**
 * Shared-state backend for the Livewire dashboard.
 *
 * Stores one JSON blob (commitments table, each rep's Top Target / Biggest
 * Win, and the Josh/Ben notes) in cell A1 of a dedicated "DashboardState"
 * tab in this spreadsheet, so every dashboard viewer reads/writes the same
 * data. GET returns the current blob; POST replaces it.
 *
 * --- Deploy steps (one time) ---
 * 1. In this Google Sheet: Extensions > Apps Script.
 * 2. Delete any placeholder code in Code.gs and paste this whole file in.
 * 3. Save (the disk icon or Cmd/Ctrl+S).
 * 4. Deploy > New deployment.
 * 5. Click the gear next to "Select type" > Web app.
 * 6. Execute as: Me. Who has access: Anyone.
 * 7. Click Deploy, then authorize it (it's your own script, on your own
 *    sheet — the scary-looking Google warning is normal for any Apps
 *    Script you deploy yourself).
 * 8. Copy the "Web app URL" it gives you (ends in /exec) and send it back
 *    so it can be pasted into SHARED_STATE_URL in index.html.
 *
 * To update the script later (e.g. if you ever change this file), use
 * Deploy > Manage deployments > (pencil icon) > New version > Deploy —
 * a plain re-save of Code.gs does NOT update the live Web App URL.
 */

const STATE_SHEET_NAME = 'DashboardState';

function getStateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STATE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STATE_SHEET_NAME);
    sheet.getRange('A1').setValue('{}');
    sheet.hideSheet();
  }
  return sheet;
}

function doGet(e) {
  const sheet = getStateSheet_();
  const json = sheet.getRange('A1').getValue() || '{}';
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// Merges an incoming save into whatever is already stored, instead of blindly overwriting
// it. Each browser tab only knows about its own edits — if it just replaced the whole blob,
// two people editing different reps (or different fields on the same rep) around the same
// time would stomp each other's changes. Commitments are merged per rep (by name) and per
// field within that rep's row; notes are merged per key.
function mergeState_(existing, incoming) {
  const merged = {};
  for (const key in existing) merged[key] = existing[key];

  if (incoming.commitments) {
    const byName = {};
    (existing.commitments || []).forEach(function (row) { byName[row.name] = row; });
    incoming.commitments.forEach(function (row) {
      const prev = byName[row.name] || { name: row.name };
      const next = {};
      for (const k in prev) next[k] = prev[k];
      for (const k in row) next[k] = row[k];
      byName[row.name] = next;
    });
    merged.commitments = Object.keys(byName).map(function (name) { return byName[name]; });
  }

  if (incoming.notes) {
    const notes = {};
    const prevNotes = existing.notes || {};
    for (const k in prevNotes) notes[k] = prevNotes[k];
    for (const k in incoming.notes) notes[k] = incoming.notes[k];
    merged.notes = notes;
  }

  return merged;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getStateSheet_();
    const body = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const incoming = JSON.parse(body); // throws (and aborts) if the body isn't valid JSON
    const existingRaw = sheet.getRange('A1').getValue() || '{}';
    let existing;
    try { existing = JSON.parse(existingRaw); } catch (err) { existing = {}; }
    const merged = mergeState_(existing, incoming);
    sheet.getRange('A1').setValue(JSON.stringify(merged));
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
