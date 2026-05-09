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
      data.dietary     || "",            // Per-guest dietary restrictions (pipe-separated)
      data.notes       || "",            // Additional notes
    ]);

    // Send notification email to organizers
    sendNotification(data, now);
    // Send confirmation email to guest
    sendConfirmation(data);

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
      "Dietary Restrictions",
      "Notes",
    ]);
    // Bold + freeze the header
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold");
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
Dietary:   ${data.dietary    || "—"}
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

function sendConfirmation(data) {
  if (!data.email) return;

  const MEAL_LABELS = {
    "chicken":   "Pancetta Chicken",
    "prime-rib": "Herb Rubbed King Cut Prime Rib of Beef",
    "sole":      "Citrus Basil Crab Stuffed Sole",
  };
  const ATT_LABELS = {
    "hike":          "The Summit — Mount Elbert (Sep 4)",
    "dinner":        "The Celebration — The Wright Room (Sep 6)",
    "both":          "Both — The Full Adventure (Sep 4 + 6)",
    "cannot-attend": "Cannot Attend",
  };

  const guests = data.guestData ? JSON.parse(data.guestData) : [{ name: data.name, attendance: data.attendance, meal: data.meals, dietary: data.dietary }];

  const anyHike   = guests.some(function(g) { return g.attendance === "hike"   || g.attendance === "both"; });
  const anyDinner = guests.some(function(g) { return g.attendance === "dinner" || g.attendance === "both"; });
  const allOut    = guests.every(function(g) { return g.attendance === "cannot-attend"; });

  const guestLines = guests.map(function(g) {
    var line = (g.name || "Guest") + "\n  Joining for: " + (ATT_LABELS[g.attendance] || g.attendance);
    if (g.meal)    line += "\n  Entrée:      " + (MEAL_LABELS[g.meal] || g.meal);
    if (g.dietary) line += "\n  Dietary:     " + g.dietary;
    return line;
  }).join("\n\n");

  var eventSection = "";
  if (anyHike) {
    eventSection +=
      "\n──────────────────────────\n" +
      "PART I — THE SUMMIT\n" +
      "Mount Elbert · 14,440 ft · Lake County, CO\n" +
      "Friday, September 4, 2026\n" +
      "Depart trailhead at 4:00 AM · Summit by lunchtime\n" +
      "Rain check date: Saturday, September 5\n";
  }
  if (anyDinner) {
    eventSection +=
      "\n──────────────────────────\n" +
      "PART II — THE CELEBRATION\n" +
      "The Wright Room · Denver, Colorado\n" +
      "Sunday, September 6, 2026 · Evening dinner\n";
  }

  const opening = allOut
    ? "We're sorry you can't make it, but we're so grateful you took the time to let us know. You'll be in our hearts on the day."
    : "We're so excited to celebrate with you. Here's a summary of your RSVP — save this for your records.";

  const body = [
    "Jackson & Crystal — September 2026",
    "",
    opening,
    "",
    "YOUR RSVP SUMMARY",
    "──────────────────────────",
    guestLines,
    data.notes ? "\nNotes: " + data.notes : "",
    eventSection,
    "──────────────────────────",
    "",
    "If you have any questions, just reply to this email.",
    "",
    "With love,",
    "Jackson & Crystal",
  ].join("\n").trim();

  MailApp.sendEmail({
    to:      data.email,
    subject: "Your RSVP — Jackson & Crystal · September 2026",
    body:    body,
  });
}

// ── OPTIONAL: run this once manually to test the sheet setup ──
function testSetup() {
  const sheet = getOrCreateSheet();
  Logger.log("Sheet ready: " + sheet.getName());
}
