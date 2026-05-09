// ============================================================
//  Jackson & Crystal RSVP — Google Apps Script
//  Paste this entire file into your Apps Script editor,
//  then deploy as a Web App (see instructions below).
// ============================================================

// ── CONFIG ───────────────────────────────────────────────────
const NOTIFICATION_EMAIL = "lyonshuynh@gmail.com"; // where RSVP emails go
const SHEET_NAME         = "RSVPs";                // tab name in your spreadsheet
// ─────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    // Parse the incoming JSON body
    const data = JSON.parse(e.postData.contents);

    const sheet = getOrCreateSheet();
    const now   = new Date();

    // Append one row per submission
    sheet.appendRow([
      now,                               // Timestamp
      data.name        || "",            // Full name (submitter)
      data.email       || "",            // Email
      data.attendance  || "",            // hike / dinner / both / cannot-attend
      data.guests      || "",            // Number of guests
      data.guestNames  || "",            // All guest names (pipe-separated)
      data.meals       || "",            // Per-guest meal choices (pipe-separated)
      data.notes       || "",            // Dietary needs / notes
    ]);

    // Send notification email
    sendNotification(data, now);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── HELPERS ──────────────────────────────────────────────────

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Add header row on first run
    sheet.appendRow([
      "Timestamp",
      "Name",
      "Email",
      "Attending",
      "Guests",
      "Guest Names",
      "Meal Selections",
      "Notes",
    ]);
    // Bold + freeze the header
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function attendanceLabel(value) {
  const map = {
    "hike":          "The Summit — Mount Elbert (Sep 4)",
    "dinner":        "The Celebration — The Wright Room (Sep 6)",
    "both":          "Both — The Full Adventure (Sep 4 + 6)",
    "cannot-attend": "Cannot Attend",
  };
  return map[value] || value;
}

function mealLabel(value) {
  const map = {
    "chicken":   "Pancetta Chicken",
    "prime-rib": "Herb Rubbed King Cut Prime Rib of Beef",
    "sole":      "Citrus Basil Crab Stuffed Sole",
  };
  // meals arrive pipe-separated, e.g. "chicken | prime-rib"
  if (!value) return "—";
  return value.split(" | ").map((v, i) => `Guest ${i + 1}: ${map[v.trim()] || v.trim()}`).join("\n           ");
}

function sendNotification(data, timestamp) {
  const label   = attendanceLabel(data.attendance);
  const subject = `New RSVP from ${data.name || "a guest"}`;

  const body = `
You have a new RSVP for Jackson & Crystal's wedding!

──────────────────────────
Name:      ${data.name       || "—"}
Email:     ${data.email      || "—"}
Attending: ${label}
Guests:    ${data.guests     || "—"}
Names:     ${data.guestNames || "—"}
Meals:     ${mealLabel(data.meals)}
Notes:     ${data.notes      || "—"}
──────────────────────────
Received:  ${timestamp.toLocaleString()}

View all responses:
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
  `.trim();

  MailApp.sendEmail({
    to:      NOTIFICATION_EMAIL,
    subject: subject,
    body:    body,
  });
}

// ── OPTIONAL: run this once manually to test the sheet setup ──
function testSetup() {
  const sheet = getOrCreateSheet();
  Logger.log("Sheet ready: " + sheet.getName());
}
