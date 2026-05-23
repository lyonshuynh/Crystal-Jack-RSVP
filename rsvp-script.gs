// ============================================================
//  Jackson & Crystal RSVP — Google Apps Script
//  Paste this entire file into your Apps Script editor,
//  then deploy as a Web App (see instructions below).
// ============================================================

// ── CONFIG ───────────────────────────────────────────────────
const NOTIFICATION_EMAIL = "lyonshuynh@gmail.com"; // where RSVP emails go
const SHEET_NAME         = "RSVPs";                // tab name in your spreadsheet
// ─────────────────────────────────────────────────────────────

// Health check — open the /exec URL in a browser to confirm the deploy is live.
function doGet() {
  return jsonResp({ result: "ok", message: "RSVP endpoint is live." });
}

function doPost(e) {
  // Serialize submissions so concurrent writes can't interleave the
  // delete-old + append-new pair below.
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (_) {
    return jsonResp({ result: "error", message: "Server busy, please try again." });
  }

  try {
    const data = JSON.parse(e.postData.contents);

    // Spam honeypot: legitimate guests never populate this field.
    if (data.company) return jsonResp({ result: "ok" });

    const emailRaw = String(data.email || "").trim();
    const emailKey = emailRaw.toLowerCase();

    // Rate limit: 5-second floor between submissions from the same email.
    if (emailKey) {
      const props = PropertiesService.getScriptProperties();
      const pkey  = "ts_" + emailKey;
      const last  = Number(props.getProperty(pkey) || 0);
      const now   = Date.now();
      if (now - last < 5000) {
        return jsonResp({ result: "error", message: "We just received your RSVP — please wait a few seconds before resubmitting." });
      }
      props.setProperty(pkey, String(now));
    }

    const sheet  = getOrCreateSheet();
    const stamp  = new Date();
    const guests = data.guestData
      ? JSON.parse(data.guestData)
      : [{ name: data.name, attendance: data.attendance, meal: data.meals, dietary: data.dietary }];

    // Latest-wins idempotency: drop any prior rows for this email before writing
    // the new set. Without this, a guest who edits their RSVP creates duplicates.
    if (emailKey) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // column B = Email
        const toDelete = [];
        for (let i = 0; i < values.length; i++) {
          if (String(values[i][0] || "").trim().toLowerCase() === emailKey) toDelete.push(i + 2);
        }
        for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
      }
    }

    // One row per guest.
    guests.forEach(function(g) {
      sheet.appendRow([
        stamp,
        data.email   || "",
        data.phone   || "",
        data.address || "",
        g.name       || "",
        attendanceLabel(g.attendance),
        mealLabel(g.meal),
        g.dietary    || "",
        data.notes   || "",
      ]);
    });

    // Mail sends are independent — a quota or transient failure must not lose
    // a successful sheet write or mislead the guest about their RSVP status.
    safeSend(function() { sendNotification(data, guests, stamp); }, "notify");
    safeSend(function() { sendConfirmation(data, guests); },        "confirm");

    return jsonResp({ result: "ok" });

  } catch (err) {
    return jsonResp({ result: "error", message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeSend(fn, label) {
  try { fn(); }
  catch (err) {
    try {
      Logger.log("[" + label + "] mail failed: " + err);
      SpreadsheetApp.getActiveSpreadsheet().toast(label + " email failed: " + err, "RSVP", 8);
    } catch (_) {}
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
    "chicken": "Chicken Piccata",
    "bass":    "Colorado Striped Bass",
    "strip":   "Green Chili Rubbed NY Strip",
    "special": "Special Request (see dietary notes)",
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

  // GmailApp uses the user's ~1500/day quota rather than MailApp's ~100/day.
  GmailApp.sendEmail(
    NOTIFICATION_EMAIL,
    "New RSVP from " + (data.name || "a guest"),
    body
  );
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
      "Rain check date: Saturday, September 5 — dinner on September 6 is unaffected\n\n" +
      "The Weekend\n" +
      "Wed Sep 2 — Arrive at the lodge\n" +
      "Thu Sep 3 — Acclimate: explore Leadville, board games, take it easy\n" +
      "Fri Sep 4 — Hike day: depart trailhead at 4:00 AM\n" +
      "Sat Sep 5 — Recovery: make your way back to Denver (2.5 hrs)\n" +
      "Sun Sep 6 — The Celebration: The Wright Room, Denver\n" +
      "Rain check plan: Friday becomes a bonus lodge day — Saturday we hike, then drive to Denver.\n\n" +
      "Trailhead: North Mt. Elbert Trailhead\n" +
      "Via County Road 11 (Halfmoon Creek Road) near Leadville, CO\n" +
      "GPS: Forest Service Road 110, Leadville, CO 80461\n" +
      "5 miles up a dirt road west of CO 300, near Elbert Creek Campground\n\n" +
      "Lodging\n" +
      "The Lodges — We've booked two lodges near the trailhead for Wed Sep 2 through Sat Sep 5, with a mix of king, queen, and single rooms. Rooms are first-come, first-served. $150 per person for all three nights. Email LyonsHuynh@gmail.com to claim a spot.\n" +
      "On Your Own — Leadville has a range of hotels, vacation rentals, and B&Bs nearby.\n" +
      "Camping — For the bold: dispersed camping near the trailhead is always an option.\n\n" +
      "What to Wear\n" +
      "Dress for the hike — comfortable layers, broken-in boots, whatever gets you to 14,440 ft. No need to pack dress clothes. The whole idea is to hike, exchange rings somewhere breathtaking, and head back down. The dinner on Sunday is where we'll take the nice photos.\n";
  }
  if (anyDinner) {
    eventSection +=
      "\n──────────────────────────\n" +
      "PART II — THE CELEBRATION\n" +
      "The Wright Room\n" +
      "535 16th St Mall, Suite 240, Denver, CO 80202\n" +
      "Inside the historic Denver Masonic Building\n" +
      "Sunday, September 6, 2026\n\n" +
      "Schedule\n" +
      "5:00 PM — Cocktail hour\n" +
      "6:00 PM — Dinner begins\n" +
      "9:00 PM — We say goodnight\n\n" +
      "Parking: Denver Pavilion next door — covered by us. Rideshare and public transit also convenient.\n\n" +
      "Attire: Semi-formal\n\n" +
      "Lodging: Aloft by Marriott Denver Downtown\n" +
      "800 15th St, Denver, CO 80202 · 3 blocks from The Wright Room\n" +
      "Book the group rate: https://app.marriott.com/reslink?id=1779477866112&key=GRP&app=resvlink\n" +
      "The group rate is available 2 days before and 2 days after the contracted dates — choose your preferred arrival and departure when checking availability.\n" +
      "Reservation questions? Glenn Wentzel at Aloft: 720-240-5101\n";
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

  GmailApp.sendEmail(
    data.email,
    "Your RSVP — Jackson & Crystal · September 2026",
    body
  );
}

// ── OPTIONAL: run this once manually to test the sheet setup ──
function testSetup() {
  const sheet = getOrCreateSheet();
  Logger.log("Sheet ready: " + sheet.getName());
}
