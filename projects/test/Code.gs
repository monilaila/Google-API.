// ===== Reviews test endpoint (standalone Apps Script project) =====
// Verifies a Google Sign-In ID token, then appends a row to the "Reviews" sheet.

const CLIENT_ID = "PASTE_YOUR_CLIENT_ID.apps.googleusercontent.com";
const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID";
const SHEET_NAME = "Reviews";
const MAX_MESSAGE_LENGTH = 1000;

function doGet() {
  return json_({ ok: true, info: "Reviews test endpoint is running" });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const message = String(body.message || "").trim();

    if (!body.token) return json_({ ok: false, error: "Missing sign-in token" });
    if (!message) return json_({ ok: false, error: "Message is empty" });
    if (message.length > MAX_MESSAGE_LENGTH) {
      return json_({ ok: false, error: "Message is too long (max " + MAX_MESSAGE_LENGTH + ")" });
    }

    // The important part: Google confirms who this really is.
    const user = verifyToken_(body.token);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    let row;
    try {
      const sheet = getSheet_();
      sheet.appendRow([new Date(), user.name || "", user.email, safeCell_(message)]);
      row = sheet.getLastRow();
    } finally {
      lock.releaseLock();
    }

    return json_({ ok: true, email: user.email, row: row, sheet: SHEET_NAME });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function verifyToken_(token) {
  const res = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(token),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error("Google rejected the token (it may have expired - sign in again)");
  }
  const info = JSON.parse(res.getContentText());

  if (info.aud !== CLIENT_ID) throw new Error("Token was issued for a different app");
  if (info.iss !== "accounts.google.com" && info.iss !== "https://accounts.google.com") {
    throw new Error("Unexpected token issuer");
  }
  if (String(info.email_verified) !== "true") throw new Error("Email is not verified by Google");
  if (Number(info.exp) * 1000 < Date.now()) throw new Error("Token expired - sign in again");
  if (!info.email) throw new Error("No email in token");

  return info;
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Date", "Name", "Google Email", "Message"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Stops a message starting with = + - @ from being treated as a spreadsheet formula.
function safeCell_(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
