// ============================================================
//  Jackson & Crystal RSVP — Google Apps Script
//  Paste this entire file into your Apps Script editor,
//  then deploy as a Web App (see instructions below).
// ============================================================

// ── CONFIG ───────────────────────────────────────────────────
const NOTIFICATION_EMAIL = "lyonshuynh@gmail.com"; // where RSVP emails go
const SHEET_NAME         = "RSVPs";                // tab name in your spreadsheet
// ─────────────────────────────────────────────────────────────

// ── CONFIG (lookup) ──────────────────────────────────────────
const SITE_URL          = "https://lyonshuynh.github.io/Crystal-Jack-RSVP/";
const TOKEN_TTL_MS      = 24 * 60 * 60 * 1000;   // magic-link valid for 24h, reusable until expiry
const LOOKUP_MAX_PER_HR = 5;                     // lookup-link requests per email per hour
// ─────────────────────────────────────────────────────────────

// Health check (GET /exec) — or, with ?action=fetch&token=..., return RSVP data.
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action === "fetch" && params.token) return fetchByToken(params.token);
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

    // Lookup flow: a guest is asking us to email them a magic link to edit
    // their RSVP. ALWAYS return ok — we don't leak which addresses are on
    // the guest list, regardless of whether a match exists.
    if (data.action === "lookup") {
      const lkEmail = String(data.email || "").trim().toLowerCase();
      sweepExpiredTokens();
      if (lkEmail && checkLookupRate(lkEmail)) startLookup(lkEmail);
      return jsonResp({ result: "ok" });
    }

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
    // the new set. If the guest is editing via magic link and changed the
    // email field, `prevEmail` carries the original key so we still find them.
    const dedupeKey = String(data.prevEmail || "").trim().toLowerCase() || emailKey;
    if (dedupeKey) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // column B = Email
        const toDelete = [];
        for (let i = 0; i < values.length; i++) {
          if (String(values[i][0] || "").trim().toLowerCase() === dedupeKey) toDelete.push(i + 2);
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

// ── LOOKUP / MAGIC-LINK ──────────────────────────────────────

// Build a guestList payload for the form to pre-fill, by scanning the sheet
// for rows matching this email. Returns null if no match.
function lookupRsvp(emailKey) {
  const sheet   = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const matched = rows.filter(function(r) {
    return String(r[1] || "").trim().toLowerCase() === emailKey;
  });
  if (matched.length === 0) return null;
  // After latest-wins idempotency all rows for an email share the same
  // submission timestamp, phone, address, and notes — use the first.
  const head = matched[0];
  return {
    email:   head[1] || "",
    phone:   head[2] || "",
    address: head[3] || "",
    notes:   head[8] || "",
    guests:  matched.map(function(r) {
      return {
        name:       r[4] || "",
        attendance: attendanceValueFromLabel(r[5]),
        meal:       mealValueFromLabel(r[6]),
        mealName:   r[6] || "",
        dietary:    r[7] || "",
      };
    }),
  };
}

// Inverse of attendanceLabel() — the sheet stores labels, but the form pre-fill
// needs the value codes ("hike", "dinner", "both", "cannot-attend").
function attendanceValueFromLabel(label) {
  if (!label) return "";
  const s = String(label);
  if (s.indexOf("Summit")        !== -1) return "hike";
  if (s.indexOf("Celebration")   !== -1) return "dinner";
  if (s.indexOf("Both")          !== -1) return "both";
  if (s.indexOf("Cannot Attend") !== -1) return "cannot-attend";
  return "";
}

function mealValueFromLabel(label) {
  if (!label) return "";
  const s = String(label);
  if (s.indexOf("Piccata")  !== -1) return "chicken";
  if (s.indexOf("Bass")     !== -1) return "bass";
  if (s.indexOf("NY Strip") !== -1) return "strip";
  if (s.indexOf("Special")  !== -1) return "special";
  return "";
}

// Always returns 200 with the same shape — the *only* signal of a hit is whether
// the guest receives an email at the address they entered.
function startLookup(emailKey) {
  const sheet   = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Find canonical (original-case) email + check existence in one pass.
  const range = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let canonical = "";
  for (let i = 0; i < range.length; i++) {
    const v = String(range[i][0] || "").trim();
    if (v.toLowerCase() === emailKey) { canonical = v; break; }
  }
  if (!canonical) return; // silent — no email sent, response to client unchanged

  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  const exp   = Date.now() + TOKEN_TTL_MS;
  PropertiesService.getScriptProperties()
    .setProperty("tok_" + token, JSON.stringify({ email: canonical, exp: exp }));

  const link = SITE_URL + "?token=" + token;
  const body = [
    "Hi,",
    "",
    "You asked to edit your RSVP for Jackson & Crystal's wedding.",
    "Open this link to view and update your response:",
    "",
    link,
    "",
    "The link is good for 24 hours. If you didn't request it, just ignore this email — your RSVP hasn't changed.",
    "",
    "With love,",
    "Jackson & Crystal",
  ].join("\n");

  safeSend(function() {
    GmailApp.sendEmail(canonical, "Edit your RSVP — Jackson & Crystal", body);
  }, "lookup-link");
}

function fetchByToken(token) {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty("tok_" + token);
  if (!raw) return jsonResp({ result: "error", message: "expired" });
  let info;
  try { info = JSON.parse(raw); }
  catch (_) { props.deleteProperty("tok_" + token); return jsonResp({ result: "error", message: "expired" }); }
  if (!info.exp || info.exp < Date.now()) {
    props.deleteProperty("tok_" + token);
    return jsonResp({ result: "error", message: "expired" });
  }
  const data = lookupRsvp(String(info.email || "").toLowerCase());
  if (!data) return jsonResp({ result: "error", message: "expired" });
  // Token is REUSABLE until expiry — same email channel authorizes either way,
  // and reusable tokens survive page refreshes / accidental tab closes.
  return jsonResp({ result: "ok", data: data });
}

// Allow N lookup-link requests per email per hour, then quietly stop sending
// (the response shape doesn't change — still {result:"ok"}).
function checkLookupRate(emailKey) {
  const props = PropertiesService.getScriptProperties();
  const key   = "lk_" + emailKey;
  const raw   = props.getProperty(key);
  const now   = Date.now();
  let times = [];
  try { times = raw ? JSON.parse(raw) : []; } catch (_) {}
  times = times.filter(function(t) { return now - t < 60 * 60 * 1000; });
  if (times.length >= LOOKUP_MAX_PER_HR) {
    props.setProperty(key, JSON.stringify(times));
    return false;
  }
  times.push(now);
  props.setProperty(key, JSON.stringify(times));
  return true;
}

// Lazy sweep — expired tokens are dropped opportunistically when a new lookup
// arrives. PropertiesService is fine for a few hundred entries.
function sweepExpiredTokens() {
  const props = PropertiesService.getScriptProperties();
  const all   = props.getProperties();
  const now   = Date.now();
  Object.keys(all).forEach(function(k) {
    if (k.indexOf("tok_") !== 0) return;
    try {
      const info = JSON.parse(all[k]);
      if (!info.exp || info.exp < now) props.deleteProperty(k);
    } catch (_) {
      props.deleteProperty(k);
    }
  });
}

// ── OPTIONAL: run this once manually to test the sheet setup ──
function testSetup() {
  const sheet = getOrCreateSheet();
  Logger.log("Sheet ready: " + sheet.getName());
}
