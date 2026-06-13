const BOOKINGS_SHEET_NAME = "Bookings";
const GROUP_STUDY_LIMIT = 5;
const TUTORING_LIMIT = 1;
const NOTIFICATION_EMAILS = [
  "deyaanyemul@gmail.com",
  "shayodhanm2@gmail.com",
  "rishabhisnow@gmail.com"
];
const BOOKING_HEADERS = [
  "bookingId",
  "requestId",
  "createdISO",
  "date",
  "time",
  "section",
  "name",
  "email",
  "topic",
  "sessionType"
];

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const action = String(p.action || "").toLowerCase();
  const callback = String(p.callback || "").trim();

  try {
    if (action === "availability") {
      return respond({ ok: true, availability: getAvailability(cleanDate(p.date)) }, callback);
    }

    if (action === "book") {
      return respond(createBooking(p), callback);
    }

    return respond({ ok: false, error: "Invalid action" }, callback);
  } catch (err) {
    return respond({ ok: false, error: String(err.message || err) }, callback);
  }
}

function doPost(e) {
  const p = e && e.parameter ? e.parameter : {};
  const action = String(p.action || "").toLowerCase();
  const requestId = cleanText(p.requestId, 100);
  let result;

  try {
    result = action === "book"
      ? createBooking(p)
      : { ok: false, error: "Invalid action" };
  } catch (err) {
    result = { ok: false, error: String(err.message || err) };
  }

  const message = JSON.stringify({
    type: "history-help-booking-result",
    requestId: requestId,
    result: result
  }).replace(/</g, "\\u003c");

  return HtmlService.createHtmlOutput(
    "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>" +
    "<script>(function(){var message=" + message + ";" +
    "window.parent.postMessage(message, '*');" +
    "if(window.top !== window.parent) window.top.postMessage(message, '*');" +
    "}());<\/script>" +
    "</body></html>"
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createBooking(p) {
  const date = cleanDate(p.date);
  const time = cleanTime(p.time);
  const section = sectionForTime(time);
  const name = cleanText(p.name, 120);
  const email = cleanText(p.email, 200);
  const topic = cleanText(p.topic, 300);
  const sessionType = cleanText(p.sessionType, 50);
  const requestId = cleanText(p.requestId, 100);

  if (!date || !time || !section || !name || !email || !topic || !sessionType || !requestId) {
    throw new Error("Missing required booking information");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getBookingsSheet();
    const existing = findRequest(sheet, requestId);
    if (existing) return { ok: true, bookingId: existing, duplicate: true };

    const availability = getAvailabilityForSheet(sheet, date);
    const slot = availability[time];
    if (!slot || slot.full) throw new Error("That time is no longer available");

    const bookingId = Utilities.getUuid();
    sheet.appendRow([
      bookingId,
      requestId,
      new Date().toISOString(),
      date,
      time,
      section,
      name,
      email,
      topic,
      sessionType
    ]);
    sendBookingNotification({
      bookingId: bookingId,
      date: date,
      time: time,
      section: section,
      name: name,
      email: email,
      topic: topic,
      sessionType: sessionType
    });
    return { ok: true, bookingId: bookingId };
  } finally {
    lock.releaseLock();
  }
}

function sendBookingNotification(booking) {
  const sectionName = booking.section === "group-study" ? "Group Study" : "Tutoring";
  const sessionTypeName = booking.sessionType === "by-peers" ? "By peers" : "By tutor";
  const subject = "New History Help reservation: " + booking.date + " at " + booking.time;
  const body = [
    "A new reservation has been made.",
    "",
    "Date: " + booking.date,
    "Time: " + booking.time,
    "Section: " + sectionName,
    "Type of session: " + sessionTypeName,
    "Name: " + booking.name,
    "Student email: " + booking.email,
    "Topic to study: " + booking.topic,
    "Booking ID: " + booking.bookingId
  ].join("\n");

  MailApp.sendEmail({
    to: NOTIFICATION_EMAILS.join(","),
    subject: subject,
    body: body,
    replyTo: booking.email,
    name: "History Help Reservations"
  });
}

function getAvailability(date) {
  if (!date) throw new Error("Missing or invalid date");
  return getAvailabilityForSheet(getBookingsSheet(), date);
}

function getAvailabilityForSheet(sheet, date) {
  const counts = {};
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, BOOKING_HEADERS.length).getValues();
    rows.forEach(function(row) {
      if (String(row[3] || "") !== date) return;
      const time = String(row[4] || "");
      counts[time] = (counts[time] || 0) + 1;
    });
  }

  const result = {};
  allSlots().forEach(function(time) {
    const section = sectionForTime(time);
    const limit = section === "group-study" ? GROUP_STUDY_LIMIT : TUTORING_LIMIT;
    const booked = counts[time] || 0;
    result[time] = { booked: booked, limit: limit, full: booked >= limit };
  });
  return result;
}

function getBookingsSheet() {
  const book = SpreadsheetApp.getActive();
  let sheet = book.getSheetByName(BOOKINGS_SHEET_NAME);
  if (!sheet) sheet = book.insertSheet(BOOKINGS_SHEET_NAME);
  BOOKING_HEADERS.forEach(function(header, index) {
    if (!sheet.getRange(1, index + 1).getValue()) sheet.getRange(1, index + 1).setValue(header);
  });
  return sheet;
}

function findRequest(sheet, requestId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "";
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][1] || "") === requestId) return String(rows[i][0] || "");
  }
  return "";
}

function allSlots() {
  return [
    "9:00 AM", "10:00 AM", "11:00 AM", "12:00 PM", "1:00 PM",
    "2:00 PM", "3:00 PM", "4:00 PM", "5:00 PM", "6:00 PM"
  ];
}

function sectionForTime(time) {
  if (["9:00 AM", "10:00 AM", "11:00 AM", "12:00 PM", "1:00 PM"].indexOf(time) !== -1) return "group-study";
  if (["2:00 PM", "3:00 PM", "4:00 PM", "5:00 PM", "6:00 PM"].indexOf(time) !== -1) return "tutoring";
  return "";
}

function cleanDate(value) {
  const date = String(value || "").trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : "";
}

function cleanTime(value) {
  const time = cleanText(value, 20);
  return allSlots().indexOf(time) !== -1 ? time : "";
}

function cleanText(value, maxLength) {
  return String(value || "").trim().slice(0, maxLength);
}

function respond(obj, callbackName) {
  const json = JSON.stringify(obj);
  if (callbackName) {
    const safeCallback = callbackName.replace(/[^a-zA-Z0-9_$\.]/g, "");
    return ContentService.createTextOutput(safeCallback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
