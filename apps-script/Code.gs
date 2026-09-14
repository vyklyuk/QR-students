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
      case 'batchCheckin':
        return jsonResponse(handleBatchCheckin(request));
      case 'roster':
        return jsonResponse(handleRoster(request));
      case 'manual':
        return jsonResponse(handleManual(request));
      case 'groups':
        return jsonResponse(handleGroups(request));
      case 'cards':
        return jsonResponse(handleCards(request));
      case 'sendCard':
        return jsonResponse(handleSendCard(request));
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

// action: "batchCheckin" — вхід {payloads: [рядок, ...], pair} (або {..., session},
// якщо треба задати конкретну дату вручну — напр. для тестів). Той самий
// алгоритм, що й у checkin, для кожного коду по черзі, але за ОДНЕ звернення
// до Apps Script (одне відкриття таблиці) замість окремого виклику на кожного
// студента. Призначено для сценарію "спершу сканувати офлайн (напр. Команди
// на iPhone накопичують коди без мережі), потім відправити все одним пакетом".
//
// Дату сесії бере сам сервер (за своїм годинником) — клієнту (телефону)
// вистачає передати лише номер пари, без ризику розбіжності часових поясів
// телефон/сервер чи забутої дати.
// "Команди" на iPhone не завжди вдається налаштувати поле payloads як
// справжній JSON-масив (тип поля "Масив" у деяких версіях складно знайти чи
// підв'язати) — тому приймаємо або реальний масив, або звичайний рядок,
// де кожен код на своєму рядку (чи через кому): так його передає звичайне
// текстове поле зі списком "Скани".
function normalizePayloads(raw) {
  if (!raw) {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw;
  }
  return String(raw)
    .split(/\r?\n|,/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length > 0; });
}

function handleBatchCheckin(request) {
  var secret = getHmacSecret();
  var session = request.session || buildTodaySession(request.pair);
  var payloads = normalizePayloads(request.payloads);
  var results = [];

  payloads.forEach(function (payload) {
    var parsed = parsePayload(payload);

    if (!parsed) {
      writeLog('?', '?', 'forged');
      results.push({ payload: payload, status: 'forged', group: '', name: '' });
      return;
    }

    if (!verifySignature(parsed, secret)) {
      writeLog(parsed.group, parsed.number, 'forged');
      results.push({ payload: payload, status: 'forged', group: parsed.group, name: '' });
      return;
    }

    var student = findStudent(parsed.group, parsed.number);
    if (!student) {
      writeLog(parsed.group, parsed.number, 'unknown');
      results.push({ payload: payload, status: 'unknown', group: parsed.group, name: '' });
      return;
    }

    var result = markAttendance(parsed.group, student.name, session, '+');
    var status = result.duplicate ? 'duplicate' : 'ok';
    writeLog(parsed.group, parsed.number, status);
    results.push({
      payload: payload,
      status: status,
      group: parsed.group,
      name: student.name,
      presentCount: result.presentCount
    });
  });

  var summary = { ok: 0, duplicate: 0, forged: 0, unknown: 0 };
  results.forEach(function (r) {
    if (summary[r.status] !== undefined) {
      summary[r.status]++;
    }
  });

  // Готовий рядок із прізвищами (через кому) — щоб "Команди" на телефоні
  // могли одразу показати, кого зарахували, без розбору масиву results.
  var presentNames = results
    .filter(function (r) { return r.status === 'ok' || r.status === 'duplicate'; })
    .map(function (r) { return r.name; })
    .join(', ');

  return { status: 'ok', results: results, summary: summary, presentNames: presentNames };
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
    // Колонки: A № | B Прізвище, ім'я | C Email | D Картку надіслано
    var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    data.forEach(function (row) {
      var number = row[0];
      var name = row[1];
      var email = row[2];
      var cardSent = row[3];
      if (!name) {
        return;
      }
      students.push({
        number: number,
        name: name,
        email: email || '',
        cardSent: cardSent ? true : false,
        payload: buildSignedPayload(group, String(number), secret)
      });
    });
  }

  return { status: 'ok', group: group, students: students };
}

// action: "sendCard" — вхід {group, number, pngBase64}. Надсилає лист із QR-карткою
// (PNG у base64, без префіксу data:image/png;base64,) на пошту студента з колонки C,
// позначає сьогоднішню дату в колонці D ("Картку надіслано"). Лист іде від імені
// того акаунта, під яким розгорнутий сам Apps Script (SETUP.md, Крок 0).
function handleSendCard(request) {
  var group = request.group;
  var number = request.number;
  var pngBase64 = request.pngBase64;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(group);
  if (!sheet) {
    return { status: 'error', message: 'Групу не знайдено: ' + group };
  }

  var rowIndex = parseInt(number, 10) + 1; // рядок 1 — заголовки
  if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
    return { status: 'error', message: 'Студента не знайдено: №' + number };
  }

  var row = sheet.getRange(rowIndex, 1, 1, 3).getValues()[0];
  var name = row[1];
  var email = row[2];

  if (!name) {
    return { status: 'error', message: 'Студента не знайдено: №' + number };
  }
  if (!email) {
    return { status: 'error', message: name + ': не вказано email' };
  }
  if (!pngBase64) {
    return { status: 'error', message: name + ': не передано зображення картки' };
  }
  if (MailApp.getRemainingDailyQuota() <= 0) {
    return { status: 'error', message: 'Вичерпано денну квоту листів Gmail на сьогодні' };
  }

  var imageBlob = Utilities.newBlob(Utilities.base64Decode(pngBase64), 'image/png', 'qr-картка.png');
  var subject = 'QR-картка для відмітки присутності — ' + group;
  var body =
    'Вітаємо!\n\n' +
    'У вкладенні — твоя персональна QR-картка для відмітки присутності на заняттях.\n\n' +
    'Що з нею робити:\n' +
    '- Збережи зображення з вкладення в Фото на телефоні.\n' +
    '- Перед початком (або після) заняття покажи фото з екрана телефона викладачу для сканування.\n\n' +
    'Важливо:\n' +
    '- Картка персональна — не передавай її іншим студентам, вона прив\'язана саме до твого номера в групі.\n' +
    '- Якщо картку загубив, пошкодив чи не отримав — напиши викладачу, картку перевидадуть.\n\n' +
    'Гарного навчання!';

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    attachments: [imageBlob]
  });

  sheet.getRange(rowIndex, 4).setValue(new Date());

  return { status: 'ok', name: name, email: email };
}
