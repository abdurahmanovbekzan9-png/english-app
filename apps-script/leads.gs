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
 *    Для amoCRM (необязательно) добавьте ещё:
 *      AMO_DOMAIN     — адрес вашей amoCRM без https://, например orientenglish.amocrm.ru
 *      AMO_TOKEN      — долгосрочный токен (amoCRM → amoМаркет → ⋮ → «Создать интеграцию»
 *                       → «Внешняя интеграция» → после сохранения вкладка «Ключи и доступы»
 *                       → «Долгосрочный токен» → выберите срок → скопируйте)
 *      AMO_PIPELINE_ID — (необязательно) ID воронки, если сделки нужны не в основной воронке
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
  // 3) Создаём сделку + контакт в amoCRM
  try {
    sendToAmo({ name: name, phone: phone, fmt: fmt, time: time, age: age, city: city, msg: msg, lvl: lvl, user: user });
  } catch (err) {
    console.error("amoCRM: " + err);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendToAmo(d) {
  var props  = PropertiesService.getScriptProperties();
  var domain = String(props.getProperty("AMO_DOMAIN") || "").replace(/^https?:\/\//, "").replace(/\/.*$/, "");
  var token  = props.getProperty("AMO_TOKEN");
  if (!domain || !token) return;
  var base = "https://" + domain + "/api/v4";
  var headers = { Authorization: "Bearer " + token };

  var lead = {
    name: "Диагностика (Mini App): " + d.name,
    _embedded: {
      tags: [{ name: "Mini App" }, { name: "Диагностика" }],
      contacts: [{
        first_name: d.name,
        custom_fields_values: [{ field_code: "PHONE", values: [{ value: d.phone, enum_code: "MOB" }] }]
      }]
    }
  };
  var pipeline = Number(props.getProperty("AMO_PIPELINE_ID"));
  if (pipeline) lead.pipeline_id = pipeline;

  var res = UrlFetchApp.fetch(base + "/leads/complex", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify([lead]), muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error(res.getResponseCode() + " " + res.getContentText().slice(0, 300));
  var created = JSON.parse(res.getContentText());
  var leadId = created && created[0] && created[0].id;
  if (!leadId) return;

  // Все ответы из анкеты — примечанием в сделке
  var note = "Заявка из Telegram Mini App\n" +
    "Формат: " + d.fmt + "\nКогда удобно: " + d.time + "\nВозраст: " + d.age + "\nГород: " + d.city +
    (d.msg ? "\nКомментарий: " + d.msg : "") + (d.lvl ? "\nУровень в приложении: " + d.lvl : "") +
    (d.user ? "\nTelegram: @" + d.user : "");
  UrlFetchApp.fetch(base + "/leads/" + leadId + "/notes", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify([{ note_type: "common", params: { text: note } }]), muteHttpExceptions: true
  });
}

/** Для проверки: выберите эту функцию вверху редактора и нажмите «Выполнить». В amoCRM появится тестовая сделка. */
function testAmo() {
  sendToAmo({ name: "Тест Mini App", phone: "+998900000000", fmt: "Онлайн", time: "Вечером", age: "25–34", city: "Ташкент", msg: "Тестовая заявка, можно удалить", lvl: "", user: "" });
}
