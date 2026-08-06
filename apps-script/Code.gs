/**
 * Google Apps Script — קליטת תוצאות "מבחן לילה" וכתיבתן לגיליון Google Sheets.
 *
 * הגיליון ייכתב עם העמודות:
 *   תאריך ושעה | שם תלמיד | ציון | עבר/נכשל | שאלות שגויות ותשובה נכונה
 *
 * העמודה האחרונה מפרטת, לכל שאלה שנטעתה, את נוסח השאלה + התשובה הנכונה +
 * מה שהתלמיד סימן. זה למעקב המורה בלבד (התלמיד עצמו לא רואה את התשובה הנכונה).
 *
 * הוראות פריסה מלאות נמצאות בקובץ README.md (סעיף "פריסת ה-Web App").
 */

// שם הלשונית (Sheet) שאליה נכתבות התוצאות. תיווצר אוטומטית אם אינה קיימת.
var SHEET_NAME = 'תוצאות';

var HEADERS = ['תאריך ושעה', 'שם תלמיד', 'ציון', 'עבר/נכשל', 'שאלות שגויות ותשובה נכונה'];

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

    var wrongText = buildWrongText_(data.wrongDetails);

    sheet.appendRow([
      when,
      data.studentName || '',
      scoreText,
      passText,
      wrongText
    ]);

    return jsonOut_({ status: 'ok' });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

// בונה טקסט קריא לתא: כל שאלה שגויה בשורה נפרדת, עם התשובה הנכונה ומה שסומן.
function buildWrongText_(wrongDetails) {
  if (!wrongDetails || !wrongDetails.length) return '';
  return wrongDetails.map(function (d) {
    var line = d.number + '. ' + d.question +
               '\n    ✔ תשובה נכונה: ' + d.correct;
    if (d.chosen) {
      line += '\n    ✗ התלמיד סימן: ' + d.chosen;
    }
    return line;
  }).join('\n\n');
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
  // מוודא שקיימת שורת כותרות מלאה (כולל העמודה החדשה) — גם בגיליון קיים.
  var firstCell = sheet.getRange(1, 1).getValue();
  if (!firstCell) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (!sheet.getRange(1, HEADERS.length).getValue()) {
    // כותרות קיימות אך חסרה העמודה החדשה — מוסיפים רק אותה.
    sheet.getRange(1, HEADERS.length).setValue(HEADERS[HEADERS.length - 1]).setFontWeight('bold');
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
