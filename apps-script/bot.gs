/**
 * Красивое приветствие бота Orient English Express.
 * Работает в том же проекте Google Apps Script, что и leads.gs.
 *
 * КАК ПОДКЛЮЧИТЬ (один раз):
 * 1. В проекте Apps Script: «+» рядом с «Файлы» → «Скрипт» → назовите bot → вставьте весь этот код.
 * 2. В вашей функции doPost(e), рядом со строкой для заявок, добавьте ещё одну:
 *      if (body.update_id) return handleTelegramUpdate(body);
 * 3. «Свойства скрипта»: BOT_TOKEN — токен НОВОГО бота (уже нужен для заявок).
 *    Необязательно: APP_URL — адрес Mini App, если он не такой, как ниже в DEFAULT_APP_URL.
 * 4. Сделайте новое развёртывание (Управление развёртываниями → карандаш → Новая версия).
 * 5. Вверху редактора выберите функцию setupBot и нажмите «Выполнить» (один раз).
 *    Она сама подключит бота к скрипту, поставит описание, меню команд и кнопку «Открыть».
 */

var DEFAULT_APP_URL = "https://abdurahmanovbekzan9-png.github.io/english-app/";
var WELCOME_IMG = "https://raw.githubusercontent.com/abdurahmanovbekzan9-png/english-app/main/bot/welcome.png";
var SITE_URL = "https://orientenglish.com/";
var PHONE = "+998 78 555 08 07";

function appUrl_() {
  return PropertiesService.getScriptProperties().getProperty("APP_URL") || DEFAULT_APP_URL;
}

function leadUrl_() {
  var u = appUrl_();
  return u + (u.indexOf("?") < 0 ? "?" : "&") + "lead=1";
}

function tg_(method, payload) {
  var token = PropertiesService.getScriptProperties().getProperty("BOT_TOKEN");
  var res = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/" + method, {
    method: "post", contentType: "application/json",
    payload: JSON.stringify(payload || {}), muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText());
}

function mainKeyboard_() {
  var url = appUrl_();
  return { inline_keyboard: [
    [{ text: "📚 Открыть приложение", web_app: { url: url } }],
    [{ text: "📞 Бесплатная диагностика", web_app: { url: leadUrl_() } }],
    [{ text: "💬 Связаться с нами", callback_data: "contact" }, { text: "🌐 Сайт", url: SITE_URL }]
  ]};
}

function sendWelcome_(chatId, firstName) {
  var caption =
    "<b>Привет" + (firstName ? ", " + escapeHtml_(firstName) : "") + "! 👋</b>\n" +
    "Это <b>Orient English Express</b>: английский за 30 дней.\n\n" +
    "📚 <b>Учитесь у нас?</b> Откройте приложение: книги, слова и упражнения.\n" +
    "📞 <b>Новенький?</b> Запишитесь на бесплатную диагностику с завучем.";
  var r = tg_("sendPhoto", { chat_id: chatId, photo: WELCOME_IMG, caption: caption, parse_mode: "HTML", reply_markup: mainKeyboard_() });
  if (!r.ok) tg_("sendMessage", { chat_id: chatId, text: caption, parse_mode: "HTML", reply_markup: mainKeyboard_() });
}

function sendContact_(chatId) {
  tg_("sendMessage", {
    chat_id: chatId, parse_mode: "HTML",
    text: "<b>Мы на связи</b> 💬\n\n📞 " + PHONE + "\n🌐 " + SITE_URL +
          "\n\nИли оставьте заявку, и завуч сам вам позвонит:",
    reply_markup: { inline_keyboard: [[{ text: "📞 Оставить заявку", web_app: { url: leadUrl_() } }]] }
  });
}

function handleTelegramUpdate(update) {
  try {
    var m = update.message;
    if (m && m.chat && m.chat.type == "private") {
      var text = String(m.text || "").split("@")[0].trim();
      if (text.indexOf("/start") == 0 || text == "/menu") sendWelcome_(m.chat.id, m.from && m.from.first_name);
      else if (text == "/app") tg_("sendMessage", { chat_id: m.chat.id, text: "📚 Ваше приложение:", reply_markup: { inline_keyboard: [[{ text: "Открыть приложение", web_app: { url: appUrl_() } }]] } });
      else if (text == "/diagnostika") tg_("sendMessage", { chat_id: m.chat.id, text: "📞 Бесплатная диагностика с завучем: 2 минуты, чтобы записаться.", reply_markup: { inline_keyboard: [[{ text: "Записаться", web_app: { url: leadUrl_() } }]] } });
      else if (text == "/contact") sendContact_(m.chat.id);
      else sendWelcome_(m.chat.id, m.from && m.from.first_name);
    }
    var cq = update.callback_query;
    if (cq) {
      tg_("answerCallbackQuery", { callback_query_id: cq.id });
      if (cq.data == "contact") sendContact_(cq.message.chat.id);
    }
  } catch (err) {
    console.error("bot: " + err);
  }
  return ContentService.createTextOutput("ok");
}

/** Запустите один раз вручную после развёртывания. */
function setupBot() {
  var hook = ScriptApp.getService().getUrl();
  var log = [];
  log.push(["setWebhook", tg_("setWebhook", { url: hook, allowed_updates: ["message", "callback_query"], drop_pending_updates: true })]);
  log.push(["setMyCommands", tg_("setMyCommands", { commands: [
    { command: "start", description: "Главное меню" },
    { command: "app", description: "Открыть приложение" },
    { command: "diagnostika", description: "Бесплатная диагностика" },
    { command: "contact", description: "Связаться с нами" }
  ]})]);
  log.push(["setChatMenuButton", tg_("setChatMenuButton", { menu_button: { type: "web_app", text: "Открыть", web_app: { url: appUrl_() } } })]);
  log.push(["setMyShortDescription", tg_("setMyShortDescription", { short_description: "Английский за 30 дней. Приложение для учеников и бесплатная диагностика с завучем." })]);
  log.push(["setMyDescription", tg_("setMyDescription", { description:
    "Orient English Express: английский за 30 дней 🇬🇧\n\n" +
    "📚 Для учеников: книги A1–A2, слова и упражнения прямо в Telegram.\n" +
    "📞 Для новых: бесплатная диагностика с завучем. Разберём цель, уровень и график.\n\n" +
    "Нажмите «Старт» 👇" })]);
  log.forEach(function (x) { console.log(x[0] + ": " + (x[1].ok ? "OK" : JSON.stringify(x[1]))); });
}

function escapeHtml_(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
