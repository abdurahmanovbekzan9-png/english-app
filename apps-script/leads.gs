/**
 * Заявки из Mini App. Два вида:
 *   • «Быстрая заявка» (стартовый экран): имя + телефон → задача колл-центру «перезвонить сегодня».
 *   • «Запись на диагностику» (кнопка в боте): ученик сам выбирает день и время →
 *     задача колл-центру «подтвердить запись и передать админу». Занятое время другим уже не показывается.
 *
 * КАК ПОДКЛЮЧИТЬ (один раз):
 * 1. Откройте ваш проект Google Apps Script (тот, чья ссылка стоит в API_URL в index.html).
 * 2. Файл leads: замените весь код на этот (или «+» → «Скрипт» → leads, если файла ещё нет).
 * 3. В вашей функции doPost(e), сразу после строки, где разбирается JSON (что-то вроде
 *      var body = JSON.parse(e.postData.contents);
 *    должна быть строка (если уже есть — ничего не делайте):
 *      if (body.action == "lead") return handleLead(body);
 * 4. «Настройки проекта» (шестерёнка) → «Свойства скрипта»:
 *      BOT_TOKEN     — токен вашего бота от @BotFather
 *      LEADS_CHAT_ID — куда присылать заявки: ID менеджера или группы
 *                      (у группы ID начинается с минуса, например -1001234567890;
 *                       бот должен быть добавлен в эту группу)
 *    Для amoCRM:
 *      AMO_DOMAIN     — адрес вашей amoCRM без https://, например orientenglish.amocrm.ru
 *      AMO_TOKEN      — долгосрочный токен (amoCRM → amoМаркет → ⋮ → «Создать интеграцию»
 *                       → «Внешняя интеграция» → после сохранения вкладка «Ключи и доступы»
 *                       → «Долгосрочный токен» → выберите срок → скопируйте)
 *      AMO_PIPELINE_ID — (необязательно) ID воронки, если сделки нужны не в основной воронке
 *      AMO_STATUS_ID  — (необязательно) ID этапа воронки для новых заявок. Создать этап «ТЕЛЕГРАМ БОТ»
 *                       и прописать его сюда автоматически: выберите функцию amoCreateStage → «Выполнить».
 *      AMO_CALLCENTER_USER_ID — (необязательно) ID сотрудника колл-центра, кому ставить задачи.
 *                       Узнать ID: выберите вверху функцию amoUsers → «Выполнить» → смотрите «Журнал».
 *                       Если не указать — задача достанется ответственному по сделке
 *                       (как настроено распределение в вашей amoCRM).
 * 5. «Начать развёртывание» → «Управление развёртываниями» → карандаш → Версия: «Новая версия» → «Развернуть».
 *    Ссылка API_URL при этом не меняется.
 *
 * ЗАПИСИ НА ДИАГНОСТИКУ хранятся на листе «Диагностика» (создастся сам).
 * Чтобы освободить время (ученик отменил) — напишите в колонке «Статус» слово «Отмена»
 * или удалите строку. Это время снова появится в списке.
 */

// ===== Расписание диагностики (меняйте здесь) =====
var DIAG = {
  days: [1, 2, 3, 4, 5, 6],   // дни недели: 1 = пн … 6 = сб, 0 = вс
  from: "06:00",              // первая запись
  to: "22:00",                // к этому времени диагностика должна закончиться
  step: 30,                   // минут на одну диагностику
  perSlot: 1,                 // сколько учеников на одно время
  aheadDays: 7,               // на сколько рабочих дней вперёд показывать время
  noticeMin: 60               // нельзя записаться раньше, чем через столько минут
};
var DIAG_SHEET = "Диагностика";
var TZ_OFFSET_H = 5;          // Ташкент, UTC+5
var WD = ["вс", "пн", "вт", "ср", "чт", "пт", "сб"];

function handleLead(body) {
  var out;
  try {
    if (body.kind == "slots") out = { ok: true, days: freeSlots_() };
    else out = saveLead_(body);
  } catch (err) {
    console.error("lead: " + err);
    out = { ok: false, reason: "error" };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function saveLead_(body) {
  function f(k, n) { return String(body[k] || "").slice(0, n || 100); }
  var kind = body.kind == "book" ? "book" : "quick";
  var name = f("name"), phone = f("phone", 30), fmt = f("fmt", 20), time = f("time", 20),
      age = f("age", 20), city = f("city", 50), msg = f("msg", 500), lvl = f("lvl", 20),
      user = f("username", 60), uid = f("uid", 30), slot = f("slot", 20), lang = f("lang", 20);
  var ss = book_();
  var slotDate = null, slotText = "";

  // 0) Запись на конкретное время: проверяем, что время ещё свободно, и сразу занимаем его
  if (kind == "book") {
    slotDate = parseSlot_(slot);
    if (!slotDate) return { ok: false, reason: "bad_slot" };
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var free = freeSlots_().some(function (d) { return d.date == slot.slice(0, 10) && d.slots.indexOf(slot.slice(11)) >= 0; });
      if (!free) return { ok: false, reason: "taken" };
      var dsh = diagSheet_(ss);
      dsh.appendRow([slotDate, name, phone, fmt, age, city, msg, user ? "@" + user : "", uid, "Новая", new Date(), lang]);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    slotText = slotLabel_(slotDate);
  }

  // 1) Сохраняем в лист «Заявки» той же таблицы
  var sh = ss.getSheetByName("Заявки") || ss.insertSheet("Заявки");
  if (sh.getLastRow() == 0) sh.appendRow(["Дата", "Имя", "Телефон", "Формат", "Когда удобно", "Возраст", "Город", "Комментарий", "Уровень в приложении", "Telegram", "ID", "Тип"]);
  if (!sh.getRange(1, 12).getValue()) sh.getRange(1, 12).setValue("Тип");
  if (!sh.getRange(1, 13).getValue()) sh.getRange(1, 13).setValue("Язык");
  sh.appendRow([new Date(), name, phone, fmt, kind == "book" ? slotText : time, age, city, msg, lvl, user ? "@" + user : "", uid,
                kind == "book" ? "Запись на диагностику" : "Быстрая заявка", lang]);

  // 2) Присылаем менеджеру в Telegram
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("BOT_TOKEN");
  var chat  = props.getProperty("LEADS_CHAT_ID");
  if (token && chat) {
    var text = (kind == "book" ? "🗓 Запись на диагностику: " + slotText + "\n\n" : "🔔 Быстрая заявка — перезвонить\n\n") +
      "👤 " + name + "\n" +
      "📞 " + phone + "\n" +
      (lang ? "🗣 Язык диагностики: " + lang + "\n" : "") +
      (fmt ? "🏫 Формат: " + fmt + "\n" : "") +
      (kind != "book" && time ? "🕐 Когда: " + time + "\n" : "") +
      (age ? "🎂 Возраст: " + age + "\n" : "") +
      (city ? "📍 Город: " + city + "\n" : "") +
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

  // 3) Сделка + контакт + задача колл-центру в amoCRM
  try {
    sendToAmo({ kind: kind, slot: slotDate, lang: lang, name: name, phone: phone, fmt: fmt, time: time, age: age, city: city, msg: msg, lvl: lvl, user: user });
  } catch (err) {
    console.error("amoCRM: " + err);
  }

  return { ok: true, slot: slotText };
}

function sendToAmo(d) {
  var props  = PropertiesService.getScriptProperties();
  var domain = String(props.getProperty("AMO_DOMAIN") || "").replace(/^https?:\/\//, "").replace(/\/.*$/, "");
  var token  = props.getProperty("AMO_TOKEN");
  if (!domain || !token) return;
  var base = "https://" + domain + "/api/v4";
  var headers = { Authorization: "Bearer " + token };
  var book = d.kind == "book" && d.slot;
  var when = book ? slotLabel_(d.slot) : "";
  var cc = Number(props.getProperty("AMO_CALLCENTER_USER_ID"));
  var langTxt = d.lang ? " (язык: " + d.lang + ")" : "";

  var lead = {
    name: (book ? "Диагностика " + when + ": " : "Быстрая заявка (Mini App): ") + d.name,
    _embedded: {
      tags: (book ? [{ name: "Mini App" }, { name: "Диагностика" }, { name: "Запись через бота" }]
                  : [{ name: "Mini App" }, { name: "Быстрая заявка" }]).concat(d.lang ? [{ name: d.lang }] : []),
      contacts: [{
        first_name: d.name,
        custom_fields_values: [{ field_code: "PHONE", values: [{ value: d.phone, enum_code: "MOB" }] }]
      }]
    }
  };
  var pipeline = Number(props.getProperty("AMO_PIPELINE_ID"));
  if (pipeline) lead.pipeline_id = pipeline;
  var status = Number(props.getProperty("AMO_STATUS_ID"));   // этап воронки, например «ТЕЛЕГРАМ БОТ»
  if (status) lead.status_id = status;
  if (cc) lead.responsible_user_id = cc;

  var res = UrlFetchApp.fetch(base + "/leads/complex", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify([lead]), muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error(res.getResponseCode() + " " + res.getContentText().slice(0, 300));
  var created = JSON.parse(res.getContentText());
  var leadId = created && created[0] && created[0].id;
  if (!leadId) return;

  // Все ответы из анкеты — примечанием в сделке
  var note = (book ? "Запись на диагностику через Telegram-бота\nВремя: " + when : "Быстрая заявка из Telegram Mini App") +
    (d.lang ? "\nЯзык диагностики: " + d.lang : "") +
    (d.fmt ? "\nФормат: " + d.fmt : "") + (!book && d.time ? "\nКогда удобно: " + d.time : "") +
    (d.age ? "\nВозраст: " + d.age : "") + (d.city ? "\nГород: " + d.city : "") +
    (d.msg ? "\nКомментарий: " + d.msg : "") + (d.lvl ? "\nУровень в приложении: " + d.lvl : "") +
    (d.user ? "\nTelegram: @" + d.user : "");
  UrlFetchApp.fetch(base + "/leads/" + leadId + "/notes", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify([{ note_type: "common", params: { text: note } }]), muteHttpExceptions: true
  });

  // Задача колл-центру
  var task = {
    entity_id: leadId,
    entity_type: "leads",
    task_type_id: 1, // «Связаться»
    complete_till: Math.floor(taskDeadline_(d.kind, d.slot, new Date()).getTime() / 1000),
    text: book
      ? "Ученик записался на диагностику: " + when + (d.fmt ? " (" + d.fmt + ")" : "") + langTxt + ". Позвонить, подтвердить запись и поставить задачу админу."
      : "Быстрая заявка из Mini App" + langTxt + ". Перезвонить: " + d.name + ", " + d.phone + "."
  };
  if (cc) task.responsible_user_id = cc;
  var tr = UrlFetchApp.fetch(base + "/tasks", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify([task]), muteHttpExceptions: true
  });
  if (tr.getResponseCode() >= 300) throw new Error("task " + tr.getResponseCode() + " " + tr.getContentText().slice(0, 300));
}

// ===== Время и слоты (всё считается по Ташкенту, UTC+5) =====

// Дата-время по Ташкенту → настоящий момент времени
function tashDate_(y, m, d, hh, mm) { return new Date(Date.UTC(y, m, d, hh || 0, mm || 0) - TZ_OFFSET_H * 3600000); }
// Момент времени → части даты по Ташкенту
function tashParts_(dt) {
  var t = new Date(dt.getTime() + TZ_OFFSET_H * 3600000);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth(), d: t.getUTCDate(), wd: t.getUTCDay(), hh: t.getUTCHours(), mm: t.getUTCMinutes() };
}
function pad_(n) { return (n < 10 ? "0" : "") + n; }
function hm_(s) { var p = String(s).split(":"); return Number(p[0]) * 60 + Number(p[1] || 0); }
function slotKey_(dt) { var p = tashParts_(dt); return p.y + "-" + pad_(p.m + 1) + "-" + pad_(p.d) + " " + pad_(p.hh) + ":" + pad_(p.mm); }
function slotLabel_(dt) { var p = tashParts_(dt); return WD[p.wd] + " " + pad_(p.d) + "." + pad_(p.m + 1) + " в " + pad_(p.hh) + ":" + pad_(p.mm); }
function parseSlot_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/.exec(String(s));
  if (!m) return null;
  return tashDate_(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
}

// Все возможные слоты на ближайшие рабочие дни, без прошедших и занятых
function allSlots_(now, taken) {
  var days = [], p = tashParts_(now), earliest = now.getTime() + DIAG.noticeMin * 60000;
  var from = hm_(DIAG.from), to = hm_(DIAG.to);
  for (var i = 0; days.length < DIAG.aheadDays && i < 31; i++) {
    var day = tashDate_(p.y, p.m, p.d + i, 0, 0), dp = tashParts_(day);
    if (DIAG.days.indexOf(dp.wd) < 0) continue;
    var slots = [];
    for (var t = from; t + DIAG.step <= to; t += DIAG.step) {
      var s = tashDate_(dp.y, dp.m, dp.d, Math.floor(t / 60), t % 60);
      if (s.getTime() < earliest) continue;
      if ((taken[slotKey_(s)] || 0) >= DIAG.perSlot) continue;
      slots.push(pad_(Math.floor(t / 60)) + ":" + pad_(t % 60));
    }
    if (slots.length) days.push({ date: dp.y + "-" + pad_(dp.m + 1) + "-" + pad_(dp.d), label: WD[dp.wd] + ", " + dp.d + "." + pad_(dp.m + 1), today: i == 0, slots: slots });
  }
  return days;
}

function freeSlots_() {
  var taken = {}, sh = book_().getSheetByName(DIAG_SHEET);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues().forEach(function (r) {
      if (!(r[0] instanceof Date)) return;
      if (/отмен|cancel/i.test(String(r[9]))) return;
      var k = slotKey_(r[0]);
      taken[k] = (taken[k] || 0) + 1;
    });
  }
  return allSlots_(new Date(), taken);
}

// Срок задачи колл-центру.
// Быстрая заявка — сегодня до конца рабочего дня; если заявка пришла вечером/в выходной — следующий рабочий день до 12:00.
// Запись — до конца рабочего дня, но не позже чем за час до диагностики (и не раньше чем через 15 минут).
function taskDeadline_(kind, slot, now) {
  var p = tashParts_(now), to = hm_(DIAG.to);
  var endToday = tashDate_(p.y, p.m, p.d, Math.floor(to / 60), to % 60);
  var deadline;
  if (DIAG.days.indexOf(p.wd) >= 0 && now.getTime() < endToday.getTime() - 60 * 60000) {
    deadline = endToday;
  } else {
    for (var i = 1; i < 8; i++) {
      var d = tashDate_(p.y, p.m, p.d + i, 12, 0);
      if (DIAG.days.indexOf(tashParts_(d).wd) >= 0) { deadline = d; break; }
    }
  }
  if (kind == "book" && slot) {
    var before = new Date(slot.getTime() - 60 * 60000);
    if (before < deadline) deadline = before;
  }
  var min = new Date(now.getTime() + 15 * 60000);
  return deadline < min ? min : deadline;
}

function book_() {
  return SpreadsheetApp.getActiveSpreadsheet() ||
         SpreadsheetApp.openById("1o_NBtHhbU1r1NmKpcDRkcYRNEfMRQmQyqHGPIWt-GSU");
}
function diagSheet_(ss) {
  var sh = ss.getSheetByName(DIAG_SHEET) || ss.insertSheet(DIAG_SHEET);
  if (sh.getLastRow() == 0) {
    sh.appendRow(["Время диагностики", "Имя", "Телефон", "Формат", "Возраст", "Город", "Комментарий", "Telegram", "ID", "Статус", "Создано", "Язык"]);
    sh.getRange("A:A").setNumberFormat("dd.MM.yyyy HH:mm");
    sh.setFrozenRows(1);
  }
  return sh;
}

// ===== Для проверки: выберите функцию вверху редактора и нажмите «Выполнить» =====

/** Создаёт тестовую быструю заявку в amoCRM (сделка + задача колл-центру). */
function testAmo() {
  sendToAmo({ kind: "quick", lang: "Русский", name: "Тест Mini App", phone: "+998900000000", fmt: "", time: "", age: "", city: "", msg: "Тестовая заявка, можно удалить", lvl: "", user: "" });
}
/** Создаёт тестовую запись на диагностику в amoCRM на первое свободное время (в таблицу не пишет). */
function testAmoBooking() {
  var d = freeSlots_().filter(function (x) { return x.slots.length; })[0];
  sendToAmo({ kind: "book", lang: "Узбекский", slot: parseSlot_(d.date + " " + d.slots[0]), name: "Тест запись", phone: "+998900000000", fmt: "Офлайн", time: "", age: "25–34", city: "Ташкент", msg: "Тест, можно удалить", lvl: "", user: "" });
}
/** Показывает в «Журнале» свободное время на ближайшие дни. */
function testSlots() {
  freeSlots_().forEach(function (d) { console.log(d.label + ": " + d.slots.join(" ")); });
}
/** Показывает в «Журнале» сотрудников amoCRM и их ID (для AMO_CALLCENTER_USER_ID). */
function amoUsers() {
  var props  = PropertiesService.getScriptProperties();
  var domain = String(props.getProperty("AMO_DOMAIN") || "").replace(/^https?:\/\//, "").replace(/\/.*$/, "");
  var res = UrlFetchApp.fetch("https://" + domain + "/api/v4/users?limit=250", {
    headers: { Authorization: "Bearer " + props.getProperty("AMO_TOKEN") }, muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) { console.log("Ошибка " + res.getResponseCode() + ": " + res.getContentText().slice(0, 300)); return; }
  (JSON.parse(res.getContentText())._embedded.users || []).forEach(function (u) { console.log(u.id + " — " + u.name + " (" + u.email + ")"); });
}

/** Создаёт в воронке этап «ТЕЛЕГРАМ БОТ» (в самом начале, перед остальными этапами)
 *  и сохраняет его ID в свойство AMO_STATUS_ID. Если этап уже есть — просто сохраняет его ID. */
function amoCreateStage() {
  var NAME = "ТЕЛЕГРАМ БОТ";
  var props  = PropertiesService.getScriptProperties();
  var domain = String(props.getProperty("AMO_DOMAIN") || "").replace(/^https?:\/\//, "").replace(/\/.*$/, "");
  var headers = { Authorization: "Bearer " + props.getProperty("AMO_TOKEN") };
  var base = "https://" + domain + "/api/v4/leads/pipelines";
  var pipes = JSON.parse(UrlFetchApp.fetch(base, { headers: headers }).getContentText())._embedded.pipelines;
  var pid = Number(props.getProperty("AMO_PIPELINE_ID"));
  var pipe = pipes.filter(function (p) { return pid ? p.id == pid : p.is_main; })[0];
  var sts = pipe._embedded.statuses;
  var ex = sts.filter(function (s) { return String(s.name).toUpperCase() == NAME; })[0];
  var id = ex && ex.id;
  if (!id) {
    var work = sts.filter(function (s) { return s.type == 0 && s.id != 142 && s.id != 143; });
    var minSort = Math.min.apply(null, work.map(function (s) { return s.sort; }));
    var res = UrlFetchApp.fetch(base + "/" + pipe.id + "/statuses", {
      method: "post", contentType: "application/json", headers: headers, muteHttpExceptions: true,
      payload: JSON.stringify([{ name: NAME, sort: minSort - 1, color: "#98cbff" }])
    });
    if (res.getResponseCode() >= 300) { console.log("Ошибка " + res.getResponseCode() + ": " + res.getContentText().slice(0, 400)); return; }
    id = JSON.parse(res.getContentText())._embedded.statuses[0].id;
  }
  props.setProperty("AMO_STATUS_ID", String(id));
  console.log("Этап «" + NAME + "» в воронке «" + pipe.name + "»: ID " + id + " — сохранён в AMO_STATUS_ID");
}
