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
    const data   = JSON.parse(e.postData.contents);
    const sheet  = getOrCreateSheet();
    const now    = new Date();
    const guests = data.guestData
      ? JSON.parse(data.guestData)
      : [{ name: data.name, attendance: data.attendance, meal: data.meals, dietary: data.dietary }];

    // One row per guest
    guests.forEach(function(g) {
      sheet.appendRow([
        now,
        data.email             || "",
        data.phone             || "",
        data.address           || "",
        g.name                 || "",
        attendanceLabel(g.attendance),
        mealLabel(g.meal),
        g.dietary              || "",
        data.notes             || "",
      ]);
    });

    // One notification email per submission (not per guest)
    sendNotification(data, guests, now);
    // One confirmation email to guest per submission
    sendConfirmation(data, guests);

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
    sheet.appendRow([
      "Timestamp",
      "Email",
      "Phone",
      "Address",
      "Guest Name",
      "Joining For",
      "Meal",
      "Dietary Restrictions",
      "Notes",
    ]);
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
  return map[value] || value || "";
}

function mealLabel(value) {
  const map = {
    "chicken":   "Pancetta Chicken",
    "bass":      "Colorado Striped Bass",
    "prime-rib": "Herb Rubbed King Cut Prime Rib of Beef",
  };
  return map[value] || value || "";
}

function sendNotification(data, guests, timestamp) {
  const guestLines = guests.map(function(g, i) {
    var line = "Guest " + (i + 1) + ": " + (g.name || "—");
    line += "\n  Joining for: " + attendanceLabel(g.attendance);
    if (g.meal)    line += "\n  Entrée:      " + mealLabel(g.meal);
    if (g.dietary) line += "\n  Dietary:     " + g.dietary;
    return line;
  }).join("\n\n");

  const body = [
    "You have a new RSVP for Jackson & Crystal's wedding!",
    "",
    "──────────────────────────",
    "Submitted by: " + (data.name || "—") + " (" + (data.email || "—") + (data.phone ? " · " + data.phone : "") + ")",
    data.address ? "Address:      " + data.address : "",
    "Total guests: " + (data.guests || guests.length),
    "──────────────────────────",
    "",
    guestLines,
    data.notes ? "\nNotes: " + data.notes : "",
    "",
    "──────────────────────────",
    "Received:  " + timestamp.toLocaleString(),
    "",
    "View all responses:",
    SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  ].join("\n").trim();

  MailApp.sendEmail({
    to:      NOTIFICATION_EMAIL,
    subject: "New RSVP from " + (data.name || "a guest"),
    body:    body,
  });
}

function sendConfirmation(data, guests) {
  if (!data.email) return;

  const ATT_LABELS = {
    "hike":          "The Summit — Mount Elbert (Sep 4)",
    "dinner":        "The Celebration — The Wright Room (Sep 6)",
    "both":          "Both — The Full Adventure (Sep 4 + 6)",
    "cannot-attend": "Cannot Attend",
  };

  const anyHike   = guests.some(function(g) { return g.attendance === "hike"   || g.attendance === "both"; });
  const anyDinner = guests.some(function(g) { return g.attendance === "dinner" || g.attendance === "both"; });
  const allOut    = guests.every(function(g) { return g.attendance === "cannot-attend"; });

  const guestLines = guests.map(function(g) {
    var line = (g.name || "Guest") + "\n  Joining for: " + (ATT_LABELS[g.attendance] || g.attendance);
    if (g.meal)    line += "\n  Entrée:      " + mealLabel(g.meal);
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
      "Depart trailhead at 4:00 AM · Summit by sunrise\n" +
      "Rain check date: Saturday, September 5\n\n" +
      "Trailhead: North Mt. Elbert Trailhead\n" +
      "Via County Road 11 (Halfmoon Creek Road) near Leadville, CO\n" +
      "GPS: Forest Service Road 110, Leadville, CO 80461\n" +
      "5 miles up a dirt road west of CO 300, near Elbert Creek Campground\n";
  }
  if (anyDinner) {
    eventSection +=
      "\n──────────────────────────\n" +
      "PART II — THE CELEBRATION\n" +
      "The Wright Room\n" +
      "535 16th St Mall, Suite 240, Denver, CO 80202\n" +
      "Inside the historic Denver Masonic Building\n" +
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
