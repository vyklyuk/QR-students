// Доступ до вкладок Google Таблиці: пошук студента, робота з "Відвідуваність" і "Лог".

// Складає рядок сесії "yyyy-MM-dd п.N" за поточною датою сервера (Script Timezone) —
// клієнту достатньо передати лише номер пари, дату вгадувати/пересилати не треба.
function buildTodaySession(pair) {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return today + ' п.' + pair;
}

// Перший рядок вкладки "Налаштування" містить перелік шифрів груп, кожен —
// у своїй колонці, починаючи з колонки B (колонка A — службова, напр. підпис
// "Групи"); кожен шифр відповідає назві вкладки зі списком студентів.
function getConfiguredGroups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastColumn() < 2) {
    return [];
  }
  var row = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1).getValues()[0];
  return row
    .map(function (value) { return String(value).trim(); })
    .filter(function (value) { return value.length > 0; });
}

// Шукає студента за шифром групи (назва вкладки) і номером (номер рядка - 1).
function findStudent(group, number) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(group);
  if (!sheet) {
    return null;
  }
  var rowIndex = parseInt(number, 10) + 1; // рядок 1 — заголовки
  if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
    return null;
  }
  var name = sheet.getRange(rowIndex, 2).getValue(); // колонка B — Прізвище, ім'я
  if (!name) {
    return null;
  }
  return { name: String(name) };
}

// Повертає номер колонки заняття у вкладці "Відвідуваність", створюючи її за потреби.
function getOrCreateSessionColumn(sheet, session) {
  var lastColumn = Math.max(sheet.getLastColumn(), 2);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (var i = 2; i < headers.length; i++) { // колонки з C — заняття
    if (headers[i] === session) {
      return i + 1;
    }
  }
  var newColumn = lastColumn + 1;
  sheet.getRange(1, newColumn).setValue(session);
  return newColumn;
}

// Шукає рядок студента у "Відвідуваність" за групою і прізвищем; додає рядок, якщо його ще немає.
function findAttendanceRow(sheet, group, name) {
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === group && data[i][1] === name) {
        return i + 2;
      }
    }
  }
  var newRow = lastRow + 1;
  sheet.getRange(newRow, 1, 1, 2).setValues([[group, name]]);
  return newRow;
}

function countPresent(sheet, group, column) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }
  var groups = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var marks = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < groups.length; i++) {
    if (groups[i][0] === group && (marks[i][0] === '+' || marks[i][0] === '~')) {
      count++;
    }
  }
  return count;
}

// Ставить відмітку ('+' або '~') у клітинку заняття. Якщо там уже є відмітка — не перезаписує.
function markAttendance(group, name, session, mark) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  var column = getOrCreateSessionColumn(sheet, session);
  var row = findAttendanceRow(sheet, group, name);
  var cell = sheet.getRange(row, column);
  var existing = cell.getValue();

  if (existing === '+' || existing === '~') {
    return { duplicate: true, presentCount: countPresent(sheet, group, column) };
  }

  cell.setValue(mark);
  return { duplicate: false, presentCount: countPresent(sheet, group, column) };
}

// Повертає прізвища студентів групи, вже відмічених на цьому занятті.
function getCheckedNames(group, session) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2 || lastColumn < 3) {
    return [];
  }

  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var column = -1;
  for (var i = 2; i < headers.length; i++) {
    if (headers[i] === session) {
      column = i + 1;
      break;
    }
  }
  if (column === -1) {
    return [];
  }

  var groups = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var marks = sheet.getRange(2, column, lastRow - 1, 1).getValues();

  var checked = [];
  for (var j = 0; j < groups.length; j++) {
    if (groups[j][0] === group && (marks[j][0] === '+' || marks[j][0] === '~')) {
      checked.push(names[j][0]);
    }
  }
  return checked;
}

// Додає рядок у "Лог" і повертає затримку в секундах від попереднього сканування (будь-якого).
function writeLog(group, number, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_LOG);
  var now = new Date();
  var lastRow = sheet.getLastRow();
  var delaySeconds = '';

  if (lastRow >= 2) {
    var previousTime = sheet.getRange(lastRow, 1).getValue();
    if (previousTime instanceof Date) {
      delaySeconds = Math.round((now.getTime() - previousTime.getTime()) / 1000);
    }
  }

  sheet.appendRow([now, group, number, status, delaySeconds]);
  return delaySeconds;
}
