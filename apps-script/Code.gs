// Точка входу веб-застосунку. Диспетчер за полем "action".
//
// ВАЖЛИВО: Apps Script не віддає CORS-заголовки на preflight-запит.
// Клієнт має слати Content-Type: text/plain;charset=utf-8, тіло — JSON-рядком,
// тоді браузер вважає запит "простим" і preflight не виникає.

function doPost(e) {
  var request;
  try {
    request = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Некоректний JSON у тілі запиту' });
  }

  try {
    switch (request.action) {
      case 'checkin':
        return jsonResponse(handleCheckin(request));
      case 'roster':
        return jsonResponse(handleRoster(request));
      case 'manual':
        return jsonResponse(handleManual(request));
      case 'groups':
        return jsonResponse(handleGroups(request));
      case 'cards':
        return jsonResponse(handleCards(request));
      default:
        return jsonResponse({ status: 'error', message: 'Невідома дія: ' + request.action });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

// Проста перевірка працездатності — відкриття URL веб-застосунку в браузері (SETUP.md, Крок 8).
function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'QR-відмітка: бекенд працює (' + QR_VERSION + ')' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// action: "checkin" — вхід {payload, session}. Група студента визначається з
// самого payload (вона там уже підписана), тому клієнту вказувати її не треба —
// один сканер обслуговує одразу всі групи, перелічені в "Налаштуваннях".
function handleCheckin(request) {
  var secret = getHmacSecret();
  var parsed = parsePayload(request.payload);

  if (!parsed) {
    writeLog('?', '?', 'forged');
    return { status: 'forged', group: '' };
  }

  if (!verifySignature(parsed, secret)) {
    writeLog(parsed.group, parsed.number, 'forged');
    return { status: 'forged', group: parsed.group };
  }

  var student = findStudent(parsed.group, parsed.number);
  if (!student) {
    writeLog(parsed.group, parsed.number, 'unknown');
    return { status: 'unknown', group: parsed.group };
  }

  var result = markAttendance(parsed.group, student.name, request.session, '+');
  var status = result.duplicate ? 'duplicate' : 'ok';
  writeLog(parsed.group, parsed.number, status);

  return {
    status: status,
    name: student.name,
    group: parsed.group,
    presentCount: result.presentCount
  };
}

// action: "roster" — вхід {session}. Список студентів усіх груп, перелічених
// у першому рядку "Налаштувань", з позначками, хто вже відмічений на цьому занятті.
function handleRoster(request) {
  var groups = getConfiguredGroups();
  if (groups.length === 0) {
    return { status: 'error', message: 'У вкладці "Налаштування" не вказано жодної групи в першому рядку.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = [];

  groups.forEach(function (group) {
    var groupSheet = ss.getSheetByName(group);
    if (!groupSheet) {
      return; // вкладку групи ще не створено — пропускаємо
    }

    var lastRow = groupSheet.getLastRow();
    if (lastRow < 2) {
      return;
    }

    var data = groupSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var checkedNames = getCheckedNames(group, request.session);

    data.forEach(function (row) {
      var number = row[0];
      var name = row[1];
      if (!name) {
        return;
      }
      roster.push({
        group: group,
        number: number,
        name: name,
        checked: checkedNames.indexOf(name) !== -1
      });
    });
  });

  return { status: 'ok', groups: groups, roster: roster };
}

// action: "manual" — вхід {group, number, session}. Ручна відмітка, пише "~" замість "+".
function handleManual(request) {
  var group = request.group;
  var number = request.number;

  var student = findStudent(group, number);
  if (!student) {
    writeLog(group, number, 'unknown');
    return { status: 'unknown', group: group };
  }

  var result = markAttendance(group, student.name, request.session, '~');
  var status = result.duplicate ? 'duplicate' : 'ok';
  writeLog(group, number, status);

  return {
    status: status,
    name: student.name,
    group: group,
    presentCount: result.presentCount
  };
}

// action: "groups" — перелік груп для випадаючого списку на сторінці генератора карток.
function handleGroups(request) {
  return { status: 'ok', groups: getConfiguredGroups() };
}

// action: "cards" — вхід {group}. Підписаний payload для кожного студента групи —
// секрет лишається на бекенді, у браузер іде вже готовий рядок для QR-коду.
function handleCards(request) {
  var group = request.group;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(group);
  if (!sheet) {
    return { status: 'error', message: 'Групу не знайдено: ' + group };
  }

  var secret = getHmacSecret();
  var lastRow = sheet.getLastRow();
  var students = [];

  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(function (row) {
      var number = row[0];
      var name = row[1];
      if (!name) {
        return;
      }
      students.push({
        number: number,
        name: name,
        payload: buildSignedPayload(group, String(number), secret)
      });
    });
  }

  return { status: 'ok', group: group, students: students };
}
