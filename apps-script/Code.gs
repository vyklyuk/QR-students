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

// action: "checkin" — вхід {payload, group, session}.
function handleCheckin(request) {
  var secret = getHmacSecret();
  var parsed = parsePayload(request.payload);

  if (!parsed) {
    writeLog(request.group || '?', '?', 'forged');
    return { status: 'forged', group: request.group || '' };
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

// action: "roster" — вхід {group, session}. Список групи з позначками, хто вже відмічений.
function handleRoster(request) {
  var group = request.group;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var groupSheet = ss.getSheetByName(group);
  if (!groupSheet) {
    return { status: 'error', message: 'Групу не знайдено: ' + group };
  }

  var lastRow = groupSheet.getLastRow();
  var students = [];
  if (lastRow >= 2) {
    var data = groupSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][1]) {
        students.push({ number: data[i][0], name: data[i][1] });
      }
    }
  }

  var checkedNames = getCheckedNames(group, request.session);
  var roster = students.map(function (student) {
    return {
      number: student.number,
      name: student.name,
      checked: checkedNames.indexOf(student.name) !== -1
    };
  });

  return { status: 'ok', group: group, roster: roster };
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
