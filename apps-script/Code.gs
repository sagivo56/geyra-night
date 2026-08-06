/**
 * Google Apps Script — קליטת נתונים מהאפליקציה וכתיבתם לגיליון Google Sheets.
 *
 * שתי לשוניות (נוצרות אוטומטית אם אינן קיימות):
 *   1) "תוצאות"      — תוצאות מבחן:  תאריך ושעה | שם תלמיד | ציון | עבר/נכשל | שאלות שגויות ותשובה נכונה
 *   2) "צ'ק ליסט"    — סיום צ'ק ליסט: תאריך ושעה | שם תלמיד
 *
 * סוג הנתון נקבע לפי השדה kind בבקשה ("quiz" או "checklist").
 * הוראות פריסה מלאות נמצאות בקובץ README.md (סעיף "פריסת ה-Web App").
 */

var RESULTS_SHEET   = 'תוצאות';
var CHECKLIST_SHEET = "צ'ק ליסט";

var RESULTS_HEADERS   = ['תאריך ושעה', 'שם תלמיד', 'ציון', 'עבר/נכשל', 'שאלות שגויות ותשובה נכונה'];
var CHECKLIST_HEADERS = ['תאריך ושעה', 'שם תלמיד'];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var when = formatWhen_(data.timestamp);

    if (data.kind === 'checklist') {
      // ----- סיום צ'ק ליסט -----
      var clSheet = getOrCreateSheet_(CHECKLIST_SHEET, CHECKLIST_HEADERS);
      clSheet.appendRow([when, data.studentName || '']);
      return jsonOut_({ status: 'ok', kind: 'checklist' });
    }

    // ----- תוצאת מבחן (ברירת מחדל) -----
    var sheet = getOrCreateSheet_(RESULTS_SHEET, RESULTS_HEADERS);

    var scoreText = (data.score != null && data.totalQuestions != null)
      ? (data.score + '/' + data.totalQuestions)
      : String(data.score);

    var passText = data.passed ? 'עבר' : 'נכשל';
    var wrongText = buildWrongText_(data.wrongDetails);

    sheet.appendRow([when, data.studentName || '', scoreText, passText, wrongText]);
    return jsonOut_({ status: 'ok', kind: 'quiz' });

  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

// תאריך ושעה מקומיים (אזור זמן של הגיליון), נגזר מה-timestamp שנשלח.
function formatWhen_(timestamp) {
  var ts = timestamp ? new Date(timestamp) : new Date();
  return Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

// בונה טקסט קריא לתא: כל שאלה שגויה בשורה נפרדת, עם התשובה הנכונה ומה שסומן.
function buildWrongText_(wrongDetails) {
  if (!wrongDetails || !wrongDetails.length) return '';
  return wrongDetails.map(function (d) {
    var line = d.number + '. ' + d.question + '\n    ✔ תשובה נכונה: ' + d.correct;
    if (d.chosen) line += '\n    ✗ התלמיד סימן: ' + d.chosen;
    return line;
  }).join('\n\n');
}

// מאפשר בדיקה מהירה שהפריסה חיה (פתיחת ה-URL בדפדפן).
function doGet() {
  return jsonOut_({ status: 'alive', service: 'night-quiz-results' });
}

// מחזיר לשונית קיימת (או יוצר חדשה) ומוודא שקיימת שורת כותרות מלאה.
function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  var firstCell = sheet.getRange(1, 1).getValue();
  if (!firstCell) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (!sheet.getRange(1, headers.length).getValue()) {
    // כותרות קיימות אך חסרה העמודה האחרונה — מוסיפים רק אותה.
    sheet.getRange(1, headers.length).setValue(headers[headers.length - 1]).setFontWeight('bold');
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
