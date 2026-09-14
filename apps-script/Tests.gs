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
