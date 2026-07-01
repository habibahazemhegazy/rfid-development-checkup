// ── Google Apps Script: shared API for BOTH forms (RFID checkup + Manager feedback) ──
// This file owns the ONLY doGet/doPost. manager-script.gs provides manager helpers.
// Deploy: Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy.

// RFID column order. KEYS drives read AND write so they can't drift apart.
const KEYS = [
  "submittedAt","name","month",
  "g1name","g1status","g2name","g2status","g3name","g3status",
  "g4name","g4status","g5name","g5status","s1comment",
  "blocker","blockerDesc",
  "c1name","c1status","c1score","c2name","c2status","c2score",
  "c3name","c3status","c3score","c4name","c4status","c4score",
  "c5name","c5status","c5score","s4a","s4b","action"
];
const HEADERS = ["Timestamp","Name","Month",
  "Goal 1","G1 Status","Goal 2","G2 Status","Goal 3","G3 Status",
  "Goal 4","G4 Status","Goal 5","G5 Status","Section 1 Comment",
  "Blocker","Blocker Description",
  "Course 1","C1 Status","C1 Score","Course 2","C2 Status","C2 Score",
  "Course 3","C3 Status","C3 Score","Course 4","C4 Status","C4 Score",
  "Course 5","C5 Status","C5 Score",
  "Key Development Need","Micro-Win","Action Required"];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// RFID data lives on its own named tab (kept separate from the manager tab).
function rfidSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName("Responses");
  if (!s) {
    s = ss.insertSheet("Responses");
    s.appendRow(HEADERS);
    s.getRange(1, 1, 1, HEADERS.length)
     .setFontWeight("bold").setBackground("#1A3A5C").setFontColor("#FFFFFF");
  }
  return s;
}

// ── RFID: READ ──
function listResponses_() {
  const rows = rfidSheet_().getDataRange().getValues().slice(1);
  return rows.map((r, i) => {
    const o = { id: i };                 // id = 0-based data row → sheet row is id + 2
    KEYS.forEach((k, j) => o[k] = r[j]);
    return o;
  });
}

// ── RFID: CREATE ──
function createResponse_(d) {
  const sheet = rfidSheet_();
  sheet.appendRow(KEYS.map(k =>
    k === "submittedAt" ? new Date().toLocaleString("en-GB") : (d[k] != null ? d[k] : "")
  ));
  const cell = sheet.getRange(sheet.getLastRow(), KEYS.length); // last col = "action"
  const c = { "No Meeting": ["#C6EFCE","#276221"],
              "15-Min Call": ["#FFEB9C","#9C6500"],
              "Full Meeting": ["#FFC7CE","#9C0006"] }[d.action];
  if (c) cell.setBackground(c[0]).setFontColor(c[1]);
  return { status: "ok" };
}

// ── RFID: DELETE ──
function deleteResponse_(d) {
  rfidSheet_().deleteRow(d.id + 2);       // +1 header, +1 for 1-based rows
  return { status: "ok" };
}

// ── RFID: DELETE ALL (one request) ──
function clearResponses_() {
  const sheet = rfidSheet_();
  const n = sheet.getLastRow() - 1;       // data rows (below the header)
  if (n > 0) sheet.deleteRows(2, n);
  return { status: "ok" };
}

// ── ROUTERS (only ones in the project) ──
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : "";
  if (action === "getResponses") return json_({ responses: listResponses_() }); // RFID
  if (action === "getAll")       return json_(managerList_());                   // Manager
  return ContentService.createTextOutput("Pulse Check API running.");
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === "delete")        return json_(deleteResponse_(d)); // RFID
    if (d.action === "clearAll")      return json_(clearResponses_());  // RFID
    if (d.action === "managerSubmit") return json_(managerCreate_(d));  // Manager
    if (d.action === "managerDelete") return json_(managerDelete_(d));  // Manager
    if (d.action === "clearManager")  return json_(clearManager_());    // Manager
    return json_(createResponse_(d));                                   // RFID create
  } catch (err) {
    return json_({ status: "error", message: err.toString() });
  }
}

// One runnable check: run once from the editor, look at the log.
function test_() {
  if (KEYS.length !== HEADERS.length)
    throw new Error("KEYS/HEADERS length mismatch: " + KEYS.length + " vs " + HEADERS.length);
  Logger.log("ok: %s RFID columns; manager tab: %s", KEYS.length, MSHEET);
}
