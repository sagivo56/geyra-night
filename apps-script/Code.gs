/**
 * Google Apps Script — קליטת תוצאות "מבחן לילה" וכתיבתן לגיליון Google Sheets.
 *
 * הגיליון ייכתב עם העמודות:
 *   תאריך ושעה | שם תלמיד | ציון | עבר/נכשל
 *
 * הוראות פריסה מלאות נמצאות בקובץ README.md (סעיף "פריסת ה-Web App").
 */

// שם הלשונית (Sheet) שאליה נכתבות התוצאות. תיווצר אוטומטית אם אינה קיימת.
var SHEET_NAME = 'תוצאות';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var sheet = getOrCreateSheet_();

    // תאריך ושעה מקומיים (אזור זמן של הגיליון), נגזר מה-timestamp שנשלח.
    var ts = data.timestamp ? new Date(data.timestamp) : new Date();
    var tz = Session.getScriptTimeZone();
    var when = Utilities.formatDate(ts, tz, 'dd/MM/yyyy HH:mm:ss');

    var scoreText = (data.score != null && data.totalQuestions != null)
      ? (data.score + '/' + data.totalQuestions)
      : String(data.score);

    var passText = data.passed ? 'עבר' : 'נכשל';

    sheet.appendRow([
      when,
      data.studentName || '',
      scoreText,
      passText
    ]);

    return jsonOut_({ status: 'ok' });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

// מאפשר בדיקה מהירה שהפריסה חיה (פתיחת ה-URL בדפדפן).
function doGet() {
  return jsonOut_({ status: 'alive', service: 'night-quiz-results' });
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // כותרות — נוספות רק אם הגיליון ריק.
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['תאריך ושעה', 'שם תלמיד', 'ציון', 'עבר/נכשל']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
