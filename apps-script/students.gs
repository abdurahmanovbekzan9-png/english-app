/**
 * Вход в Mini App только для студентов: проверка номера телефона по листу «Студенты».
 *
 * КАК ПОДКЛЮЧИТЬ (один раз):
 * 1. В таблице создайте лист «Студенты». В первой строке заголовки, например:
 *      Имя | Телефон
 *    Дальше по одной строке на студента. Телефон можно писать как угодно:
 *    +998 90 123 45 67, 998901234567, 90 123-45-67. Несколько номеров через запятую тоже можно.
 *    Колонки «Telegram ID» и «Вход» скрипт добавит сам.
 * 2. В проекте Apps Script: «+» → «Скрипт» → назовите students → вставьте этот код.
 * 3. В doPost(e), рядом со строкой для заявок, добавьте:
 *      if (body.action == "verify") return handleVerify(body);
 * 4. Нужен BOT_TOKEN в «Свойствах скрипта» (тот же, что для заявок и бота).
 * 5. Новое развёртывание.
 *
 * Чтобы закрыть доступ, просто удалите строку студента из листа.
 */

var STUDENTS_SHEET = "Студенты";

function handleVerify(body) {
  var out;
  try { out = verifyStudent_(body); } catch (err) { console.error("verify: " + err); out = { ok: null, reason: "error" }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function verifyStudent_(body) {
  var token = PropertiesService.getScriptProperties().getProperty("BOT_TOKEN");
  if (!token) return { ok: null, reason: "no_token" };

  var sh = studentsSheet_();
  if (!sh) return { ok: null, reason: "no_sheet" };
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function (h) { return String(h).toLowerCase(); });
  var colPhone = findCol_(head, ["телефон", "phone", "номер"]);
  var colName  = findCol_(head, ["имя", "фио", "name", "студент"]);
  if (colPhone < 0) return { ok: null, reason: "no_phone_column" };
  var colTg = ensureCol_(sh, data, "Telegram ID");
  var colAt = ensureCol_(sh, data, "Вход");

  // 1) Вход по номеру: подписанный Telegram ответ на «Поделиться номером»
  if (body.contact) {
    var c = checkTelegramSignature_(String(body.contact), token);
    if (!c || !c.contact) return { ok: null, reason: "bad_signature" };
    var contact = JSON.parse(c.contact);
    var phone = normPhone_(contact.phone_number);
    for (var r = 1; r < data.length; r++) {
      var phones = String(data[r][colPhone]).split(/[,;\/]/).map(normPhone_);
      if (phone && phones.indexOf(phone) >= 0) {
        sh.getRange(r + 1, colTg + 1).setValue(String(contact.user_id));
        sh.getRange(r + 1, colAt + 1).setValue(new Date());
        return { ok: true, name: colName >= 0 ? String(data[r][colName]) : "", phone: "+" + phone };
      }
    }
    return { ok: false, reason: "not_found", phone: "+" + phone };
  }

  // 2) Повторная проверка при запуске: по Telegram ID из подписанных данных Mini App
  var init = checkTelegramSignature_(String(body.init || ""), token);
  if (!init || !init.user) return { ok: null, reason: "bad_signature" };
  var uid = String(JSON.parse(init.user).id);
  for (var k = 1; k < data.length; k++) {
    if (String(data[k][colTg]) == uid) return { ok: true, name: colName >= 0 ? String(data[k][colName]) : "" };
  }
  return { ok: false, reason: "not_found" };
}

/** Проверка подписи Telegram (как для initData). Возвращает поля или null. */
function checkTelegramSignature_(raw, token) {
  if (!raw) return null;
  var fields = {}, hash = "";
  raw.split("&").forEach(function (pair) {
    var i = pair.indexOf("="); if (i < 0) return;
    var k = decodeURIComponent(pair.slice(0, i)), v = decodeURIComponent(pair.slice(i + 1).replace(/\+/g, "%20"));
    if (k == "hash") hash = v; else fields[k] = v;
  });
  if (!hash) return null;
  var dcs = Object.keys(fields).filter(function (k) { return k != "signature"; }).sort()
    .map(function (k) { return k + "=" + fields[k]; }).join("\n");
  var secret = Utilities.computeHmacSha256Signature(Utilities.newBlob(token).getBytes(), Utilities.newBlob("WebAppData").getBytes());
  var sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(dcs).getBytes(), secret);
  var hex = sig.map(function (b) { return ("0" + (b & 0xff).toString(16)).slice(-2); }).join("");
  if (hex != hash) return null;
  if (fields.auth_date && Date.now() / 1000 - Number(fields.auth_date) > 86400 * 30) return null;
  return fields;
}

/** Номер → только цифры с кодом Узбекистана: 998XXXXXXXXX */
function normPhone_(p) {
  var d = String(p || "").replace(/\D/g, "");
  if (d.length == 9) d = "998" + d;
  if (d.length == 10 && d.charAt(0) == "8") d = "998" + d.slice(1);
  return d;
}

function studentsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById("1o_NBtHhbU1r1NmKpcDRkcYRNEfMRQmQyqHGPIWt-GSU");
  return ss.getSheetByName(STUDENTS_SHEET);
}
function findCol_(head, names) {
  for (var i = 0; i < head.length; i++) for (var j = 0; j < names.length; j++) if (head[i].indexOf(names[j]) >= 0) return i;
  return -1;
}
function ensureCol_(sh, data, title) {
  var i = data[0].map(String).indexOf(title);
  if (i >= 0) return i;
  i = data[0].length;
  sh.getRange(1, i + 1).setValue(title);
  data[0].push(title);
  for (var r = 1; r < data.length; r++) data[r].push("");
  return i;
}
