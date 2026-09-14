// Розбір payload з QR-коду та перевірка HMAC-підпису.
// Секрет живе лише у Script Properties, у код не потрапляє.

function getHmacSecret() {
  var secret = PropertiesService.getScriptProperties().getProperty('HMAC_SECRET');
  if (!secret) {
    throw new Error('HMAC_SECRET не налаштовано у Script Properties');
  }
  return secret;
}

// "v1|КН-31|07|a7Kd92Lm" → {version, group, number, signature} або null, якщо формат невірний.
function parsePayload(payload) {
  var parts = String(payload || '').split('|');
  if (parts.length !== 4) {
    return null;
  }
  var version = parts[0];
  var group = parts[1];
  var number = parts[2];
  var signature = parts[3];
  if (version !== QR_VERSION || !group || !number || !signature) {
    return null;
  }
  return { version: version, group: group, number: number, signature: signature };
}

function computeSignature(message, secret) {
  var rawBytes = Utilities.computeHmacSha256Signature(message, secret);
  var base64url = Utilities.base64EncodeWebSafe(rawBytes);
  return base64url.substring(0, SIGNATURE_LENGTH);
}

function verifySignature(parsed, secret) {
  var message = parsed.version + '|' + parsed.group + '|' + parsed.number;
  var expected = computeSignature(message, secret);
  return expected === parsed.signature;
}

// Складає підписаний payload — використовується генератором карток (Етап 4) і тестами.
function buildSignedPayload(group, number, secret) {
  var message = QR_VERSION + '|' + group + '|' + number;
  return message + '|' + computeSignature(message, secret);
}
