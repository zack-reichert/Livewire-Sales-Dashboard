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

function doPost(e) {
  const sheet = getStateSheet_();
  const body = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
  // Validate it's actually JSON before writing, so a bad request can't corrupt the store.
  JSON.parse(body);
  sheet.getRange('A1').setValue(body);
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
