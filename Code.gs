/**
 * M-600 Desk Reservation — Google Apps Script Backend
 *
 * Sheet columns (Row 1 headers):
 *   A: desk | B: name | C: startTime | D: endTime | E: days | F: id
 *
 * After updating, Deploy > Manage deployments > Edit > New version > Deploy
 */

/**
 * Normalize any time value (Date object, date string, or "HH:MM") to "HH:MM" string.
 * This ensures the frontend always receives a consistent, parseable format.
 */
function normalizeTime(val) {
  if (!val) return '';
  // Google Sheets may return a native Date object for time columns
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  var s = String(val).trim();
  if (!s) return '';
  // If it's already "HH:MM" or "HH:MM:SS", extract just HH:MM
  var match = s.match(/^(\d{1,2}):(\d{2})/);
  if (match) {
    return String(parseInt(match[1])).padStart(2, '0') + ':' + match[2];
  }
  // If it's a full date string, parse it
  try {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      var h = d.getHours();
      var m = d.getMinutes();
      return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
    }
  } catch(e) {}
  return s;
}

function timeToMin(t) {
  if (!t) return null;
  var s = String(t).trim();
  var h, m;
  // Handle Google Sheets Date objects and date strings
  if (s.indexOf('GMT') !== -1 || s.indexOf('T') !== -1 || s.indexOf('Z') !== -1 || s.match(/\d{4}/)) {
    try {
      var d = new Date(s);
      if (!isNaN(d.getTime())) {
        h = d.getHours();
        m = d.getMinutes();
      }
    } catch(e) {}
  }
  // Handle "HH:MM" or "HH:MM:SS"
  if (h === undefined || h === null) {
    var parts = s.split(':');
    h = parseInt(parts[0]);
    m = parseInt(parts[1] || 0);
  }
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function safeMin(val, fallback) {
  var m = timeToMin(val);
  return (m !== null && !isNaN(m)) ? m : fallback;
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var reservations = {};

  // Skip header row, build arrays per desk
  for (var i = 1; i < data.length; i++) {
    var desk = String(data[i][0]).trim();
    var name = String(data[i][1]).trim();
    if (!desk || !name) continue;

    var startTime = normalizeTime(data[i][2]);
    var endTime   = normalizeTime(data[i][3]);
    var daysRaw   = data[i][4] ? String(data[i][4]).trim() : '';
    var id        = data[i][5] ? String(data[i][5]).trim() : String(i);

    var entry = {
      id: id,
      name: name,
      startTime: startTime,
      endTime: endTime,
      days: daysRaw ? daysRaw.split(',').map(function(d) { return d.trim(); }).filter(function(d) { return d; }) : []
    };

    if (!reservations[desk]) {
      reservations[desk] = [];
    }
    reservations[desk].push(entry);
  }

  return ContentService
    .createTextOutput(JSON.stringify(reservations))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var params = JSON.parse(e.postData.contents);
  var action = params.action;

  if (action === 'reserve') {
    var deskId    = params.desk;
    var name      = params.name;
    var startTime = params.startTime || '';
    var endTime   = params.endTime || '';
    var days      = Array.isArray(params.days) ? params.days.join(', ') : '';
    var id        = params.id || String(new Date().getTime());

    // Server-side conflict check before saving
    var newDays = days ? days.split(',').map(function(d) { return d.trim(); }).filter(function(d) { return d; }) : ['Mon','Tue','Wed','Thu','Fri'];
    var newStart = safeMin(startTime, 0);
    var newEnd = safeMin(endTime, 1439);

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== deskId) continue;

      var eDaysRaw = data[i][4] ? String(data[i][4]).trim() : '';
      var eDays = eDaysRaw ? eDaysRaw.split(',').map(function(d) { return d.trim(); }).filter(function(d) { return d; }) : ['Mon','Tue','Wed','Thu','Fri'];
      var eStart = safeMin(normalizeTime(data[i][2]), 0);
      var eEnd = safeMin(normalizeTime(data[i][3]), 1439);

      // Check shared days
      var shared = false;
      for (var j = 0; j < newDays.length; j++) {
        if (eDays.indexOf(newDays[j]) !== -1) { shared = true; break; }
      }
      if (!shared) continue;

      // Check time overlap
      if (newStart < eEnd && eStart < newEnd) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'conflict', conflictWith: String(data[i][1]).trim() }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    sheet.appendRow([deskId, name, startTime, endTime, days, id]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'release') {
    var releaseId = params.id;

    if (releaseId) {
      // Delete the row with matching id
      var data = sheet.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][5]).trim() === String(releaseId)) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}
