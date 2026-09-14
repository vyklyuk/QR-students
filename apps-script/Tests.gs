// Тестова функція для Етапу 1. Запускати вручну з редактора Apps Script (вибрати
// testCheckinScenarios у випадаючому списку функцій і натиснути "Виконати").
//
// Створює тимчасову вкладку групи "ТЕСТ-01" з вигаданими студентами, проганяє
// п'ять сценаріїв із ROADMAP.md і перевіряє статуси. Записи в "ТЕСТ-01",
// "Відвідуваність" і "Лог" лишаються в таблиці — так замовник бачить результат.

function testCheckinScenarios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var secret = getHmacSecret();
  var testGroup = 'ТЕСТ-01';
  var testSession = 'тест ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  setUpTestGroup(ss, testGroup);

  var results = [];

  // 1. Коректний код → зараховано.
  results.push(runScenario('Коректний код', {
    action: 'checkin',
    payload: buildSignedPayload(testGroup, '1', secret),
    group: testGroup,
    session: testSession
  }, 'ok'));

  // 2. Підроблений підпис → forged.
  var forgedPayload = buildSignedPayload(testGroup, '1', secret).slice(0, -1) + 'X';
  results.push(runScenario('Підроблений підпис', {
    action: 'checkin',
    payload: forgedPayload,
    group: testGroup,
    session: testSession
  }, 'forged'));

  // 3. Неіснуючий номер → unknown.
  results.push(runScenario('Неіснуючий номер', {
    action: 'checkin',
    payload: buildSignedPayload(testGroup, '99', secret),
    group: testGroup,
    session: testSession
  }, 'unknown'));

  // 4. Повторне сканування того самого студента на тому самому занятті → duplicate.
  results.push(runScenario('Повторне сканування', {
    action: 'checkin',
    payload: buildSignedPayload(testGroup, '1', secret),
    group: testGroup,
    session: testSession
  }, 'duplicate'));

  // 5. Ручна відмітка іншого студента → ok, у клітинці "~".
  results.push(runScenario('Ручна відмітка', {
    action: 'manual',
    group: testGroup,
    number: '2',
    session: testSession
  }, 'ok'));

  Logger.log(results.join('\n'));

  var failed = results.filter(function (line) {
    return line.indexOf('ПОМИЛКА') === 0;
  });

  if (failed.length > 0) {
    throw new Error('Тест провалено:\n' + failed.join('\n'));
  }

  Logger.log(
    'Усі 5 сценаріїв пройшли успішно. Перевір вкладки "' + testGroup + '", "' +
    SHEET_ATTENDANCE + '" і "' + SHEET_LOG + '".'
  );
}

// Тест для batchCheckin — один пакетний запит на кілька кодів одразу
// (той сценарій, що й "спершу сканувати офлайн, потім відправити все разом").
function testBatchCheckin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var secret = getHmacSecret();
  var testGroup = 'ТЕСТ-01';
  var testSession = 'тест-пакет ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  setUpTestGroup(ss, testGroup);

  var validPayload = buildSignedPayload(testGroup, '1', secret);
  var forgedPayload = buildSignedPayload(testGroup, '2', secret).slice(0, -1) + 'X';
  var unknownPayload = buildSignedPayload(testGroup, '99', secret);

  var request = {
    action: 'batchCheckin',
    session: testSession,
    payloads: [validPayload, forgedPayload, unknownPayload, validPayload] // останній — дублікат
  };
  var fakeEvent = { postData: { contents: JSON.stringify(request) } };
  var response = JSON.parse(doPost(fakeEvent).getContent());

  var expected = { ok: 1, forged: 1, unknown: 1, duplicate: 1 };
  var mismatches = [];
  ['ok', 'forged', 'unknown', 'duplicate'].forEach(function (key) {
    if (response.summary[key] !== expected[key]) {
      mismatches.push(key + ': очікувалось ' + expected[key] + ', отримано ' + response.summary[key]);
    }
  });

  if (mismatches.length > 0) {
    throw new Error('testBatchCheckin провалено:\n' + mismatches.join('\n'));
  }

  Logger.log('testBatchCheckin: усі 4 записи в пакеті отримали правильний статус. Підсумок: ' + JSON.stringify(response.summary));
  Logger.log('presentNames: ' + response.presentNames);
}

// Той самий сценарій, але payloads переданий одним рядком через переноси
// рядків — так виходить, коли в "Командах" на iPhone поле не вдалось
// налаштувати як справжній JSON-масив (normalizePayloads() у Code.gs).
function testBatchCheckinTextPayloads() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var secret = getHmacSecret();
  var testGroup = 'ТЕСТ-01';
  var testSession = 'тест-пакет-текст ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  setUpTestGroup(ss, testGroup);

  var validPayload = buildSignedPayload(testGroup, '1', secret);
  var unknownPayload = buildSignedPayload(testGroup, '99', secret);

  var request = {
    action: 'batchCheckin',
    session: testSession,
    payloads: validPayload + '\n' + unknownPayload // рядок, не масив
  };
  var fakeEvent = { postData: { contents: JSON.stringify(request) } };
  var response = JSON.parse(doPost(fakeEvent).getContent());

  if (response.summary.ok !== 1 || response.summary.unknown !== 1) {
    throw new Error('testBatchCheckinTextPayloads провалено. Підсумок: ' + JSON.stringify(response.summary));
  }

  Logger.log('testBatchCheckinTextPayloads: рядок коректно розібрано на 2 коди. Підсумок: ' + JSON.stringify(response.summary));
}

function runScenario(title, request, expectedStatus) {
  var fakeEvent = { postData: { contents: JSON.stringify(request) } };
  var response = JSON.parse(doPost(fakeEvent).getContent());
  var ok = response.status === expectedStatus;
  var prefix = ok ? 'OK' : 'ПОМИЛКА';
  return prefix + ' — ' + title + ': очікувалось "' + expectedStatus + '", отримано "' + response.status + '"';
}

// Вигадана тестова група — жодних реальних прізвищ чи адрес.
function setUpTestGroup(ss, testGroup) {
  var existing = ss.getSheetByName(testGroup);
  if (existing) {
    ss.deleteSheet(existing);
  }
  var sheet = ss.insertSheet(testGroup);
  sheet.getRange(1, 1, 1, 4).setValues([['№', "Прізвище, ім'я", 'Email', 'Картку надіслано']]);
  sheet.getRange(2, 1, 2, 2).setValues([
    [1, 'Тестовий Іван Іванович'],
    [2, 'Тестова Марія Петрівна']
  ]);
}

// Допоміжна функція для налагодження сторінки сканера (Етап 2), поки немає
// друкованих карток чи генератора (Етап 4). Запускати вручну з редактора.
//
// Бере вкладку `group` (за замовчуванням "ТЕСТ" — заведи її сам, за зразком
// setUpTestGroup, із 3-5 вигаданими прізвищами), формує підписаний payload
// для кожного студента і виводить у Журнал виконання. Кожен рядок можна
// перетворити на QR-код будь-яким онлайн-генератором і показати сторінці
// сканера з екрана другого пристрою. Додатково виводить один навмисно
// зіпсований код і один із неіснуючим номером — щоб перевірити статуси
// "forged" і "unknown".
function logTestCodes(group) {
  group = group || 'ТЕСТ';
  var secret = getHmacSecret();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(group);
  if (!sheet) {
    Logger.log('Вкладку "' + group + '" не знайдено. Створи її з кількома вигаданими студентами.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('У вкладці "' + group + '" немає студентів.');
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var lines = [];

  data.forEach(function (row) {
    var number = row[0];
    var name = row[1];
    if (!name) return;
    var payload = buildSignedPayload(group, String(number), secret);
    lines.push(name + ' (№' + number + '): ' + payload);
  });

  if (lines.length === 0) {
    Logger.log('У вкладці "' + group + '" немає заповнених рядків.');
    return;
  }

  var forgedPayload = buildSignedPayload(group, String(data[0][0]), secret).slice(0, -1) + 'X';
  lines.push('Навмисно зіпсований підпис (очікуємо "forged"): ' + forgedPayload);

  var unknownNumber = Number(data[data.length - 1][0]) + 1000;
  var unknownPayload = buildSignedPayload(group, String(unknownNumber), secret);
  lines.push('Неіснуючий номер (очікуємо "unknown"): ' + unknownPayload);

  Logger.log(lines.join('\n'));
}

// Те саме, що logTestCodes(), але бере студентів не з вигаданої тестової
// вкладки, а випадково з реальних груп, перелічених у першому рядку
// "Налаштувань" (getConfiguredGroups() з SheetHelpers.gs). Зручно, коли
// хочеш прогнати сканер на живих номерах/групах, а генератора карток
// (Етап 4) ще немає. Виводить прізвища лише в приватний Журнал виконання
// Apps Script — ці дані не потрапляють у репозиторій.
function logRandomStudentCodes(count) {
  count = count || 5;
  var secret = getHmacSecret();
  var groups = getConfiguredGroups();
  Logger.log('Групи з першого рядка "Налаштувань": [' + groups.join(', ') + ']');
  if (groups.length === 0) {
    Logger.log('У вкладці "Налаштування" не вказано жодної групи в першому рядку.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allStudents = [];
  var diagnostics = [];

  groups.forEach(function (group) {
    var sheet = ss.getSheetByName(group);
    if (!sheet) {
      diagnostics.push('"' + group + '" — вкладку з такою назвою не знайдено (перевір, чи збігається символ у символ, без зайвих пробілів)');
      return;
    }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      diagnostics.push('"' + group + '" — вкладка є, але рядків зі студентами немає (lastRow=' + lastRow + ')');
      return;
    }
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var countInGroup = 0;
    data.forEach(function (row) {
      var number = row[0];
      var name = row[1];
      if (name) {
        allStudents.push({ group: group, number: number, name: name });
        countInGroup++;
      }
    });
    diagnostics.push('"' + group + '" — знайдено ' + countInGroup + ' студент(ів) із заповненим прізвищем (колонка B)');
  });

  Logger.log(diagnostics.join('\n'));

  if (allStudents.length === 0) {
    Logger.log('У жодній із перелічених груп немає студентів.');
    return;
  }

  var picked = pickRandom(allStudents, Math.min(count, allStudents.length));
  var lines = picked.map(function (student) {
    var payload = buildSignedPayload(student.group, String(student.number), secret);
    return student.name + ' (' + student.group + ', №' + student.number + '): ' + payload;
  });

  Logger.log(lines.join('\n'));
}

// Перетасовка Фішера-Єйтса — повертає n випадкових елементів масиву без повторів.
function pickRandom(array, n) {
  var copy = array.slice();
  for (var i = copy.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = copy[i];
    copy[i] = copy[j];
    copy[j] = tmp;
  }
  return copy.slice(0, n);
}
