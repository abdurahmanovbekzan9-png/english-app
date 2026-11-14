/**
 * Быстрые заявки из Mini App (кнопка 📞).
 *
 * КАК ПОДКЛЮЧИТЬ (один раз):
 * 1. Откройте ваш проект Google Apps Script (тот, чья ссылка стоит в API_URL в index.html).
 * 2. Нажмите «+» рядом с «Файлы» → «Скрипт», назовите leads и вставьте весь этот код.
 * 3. В вашей функции doPost(e), сразу после строки, где разбирается JSON (что-то вроде
 *      var body = JSON.parse(e.postData.contents);
 *    добавьте одну строку:
 *      if (body.action == "lead") return handleLead(body);
 * 4. «Настройки проекта» (шестерёнка) → «Свойства скрипта» → добавьте:
 *      BOT_TOKEN     — токен вашего бота от @BotFather
 *      LEADS_CHAT_ID — куда присылать заявки: ID менеджера или группы
 *                      (у группы ID начинается с минуса, например -1001234567890;
 *                       бот должен быть добавлен в эту группу)
 * 5. «Начать развёртывание» → «Управление развёртываниями» → карандаш → Версия: «Новая версия» → «Развернуть».
 *    Ссылка API_URL при этом не меняется.
 */
function handleLead(body) {
  function f(k, n) { return String(body[k] || "").slice(0, n || 100); }
  var name = f("name"), phone = f("phone", 30), fmt = f("fmt", 20), time = f("time", 20),
      age = f("age", 20), city = f("city", 50), msg = f("msg", 500), lvl = f("lvl", 20),
      user = f("username", 60), uid = f("uid", 30);

  // 1) Сохраняем в лист «Заявки» той же таблицы
  var ss = SpreadsheetApp.getActiveSpreadsheet() ||
           SpreadsheetApp.openById("1o_NBtHhbU1r1NmKpcDRkcYRNEfMRQmQyqHGPIWt-GSU");
  var sh = ss.getSheetByName("Заявки") || ss.insertSheet("Заявки");
  if (sh.getLastRow() == 0) sh.appendRow(["Дата", "Имя", "Телефон", "Формат", "Когда удобно", "Возраст", "Город", "Комментарий", "Уровень в приложении", "Telegram", "ID"]);
  sh.appendRow([new Date(), name, phone, fmt, time, age, city, msg, lvl, user ? "@" + user : "", uid]);

  // 2) Присылаем менеджеру в Telegram
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("BOT_TOKEN");
  var chat  = props.getProperty("LEADS_CHAT_ID");
  if (token && chat) {
    var text = "🔔 Заявка на диагностику (Mini App)\n\n" +
      "👤 " + name + "\n" +
      "📞 " + phone + "\n" +
      "🏫 Формат: " + fmt + "\n" +
      "🕐 Когда: " + time + "\n" +
      "🎂 Возраст: " + age + "\n" +
      "📍 Город: " + city + "\n" +
      (msg ? "💬 " + msg + "\n" : "") +
      (lvl ? "🎯 Уровень в приложении: " + lvl + "\n" : "") +
      (user ? "✈️ @" + user + "\n" : "");
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chat, text: text }),
      muteHttpExceptions: true
    });
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
