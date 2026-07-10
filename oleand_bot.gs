// ============================================
//  OLEAND Помощник — Telegram-бот, этап 1
//  Учёт белья: сдать в прачку, принять, остатки
// ============================================

// ⬇️⬇️⬇️  ШАГ 1: ВСТАВЬТЕ ВАШ ТОКЕН ОТ BOTFATHER МЕЖДУ КАВЫЧКАМИ  ⬇️⬇️⬇️
var TOKEN = 'ВСТАВЬТЕ_ТОКЕН_СЮДА';
// ⬆️⬆️⬆️  БОЛЬШЕ НИЧЕГО В КОДЕ МЕНЯТЬ НЕ НУЖНО  ⬆️⬆️⬆️

var OBJECTS = [
  'Петроградка б+м',
  'Саперный б+м',
  'Царская столица',
  'BLUE Железнодор.',
  'GREEN Тосина',
  'NUDE Манчестерская',
  'TIFFANY Ветеранов',
  'GRAY Лиговский проспект'
];

var ITEMS = [
  'Пододеяльник 2,0',
  'Простыня 2,0',
  'Наволочка 2,0',
  'Пододеяльник 1,5',
  'Простыня 1,5',
  'Наволочка 1,5',
  'Полотенце большое',
  'Полотенце среднее',
  'Полотенце малютка',
  'Полотенце для ног',
  'Салфетка кухонная',
  'Салфетка (клетка)'
];

// ---------- Telegram ----------

function tg(method, payload) {
  return UrlFetchApp.fetch('https://api.telegram.org/bot' + TOKEN + '/' + method, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function send(chatId, text, keyboard) {
  var p = { chat_id: chatId, text: text, parse_mode: 'HTML' };
  if (keyboard) p.reply_markup = { inline_keyboard: keyboard };
  tg('sendMessage', p);
}

// ---------- Хранилище ----------

function getSheetFile() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('OLEAND бот — данные');
    props.setProperty('SHEET_ID', ss.getId());
  }
  if (!ss.getSheetByName('Журнал')) {
    var j = ss.insertSheet('Журнал');
    j.appendRow(['Дата', 'Операция', 'Объект', 'Позиция', 'Кол-во', 'Кто']);
  }
  if (!ss.getSheetByName('Всего')) {
    var t = ss.insertSheet('Всего');
    t.appendRow(['Объект', 'Позиция', 'Всего на объекте']);
    var rows = [];
    for (var i = 0; i < OBJECTS.length; i++)
      for (var k = 0; k < ITEMS.length; k++)
        rows.push([OBJECTS[i], ITEMS[k], 0]);
    t.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  var s1 = ss.getSheetByName('Лист1');
  if (s1 && ss.getSheets().length > 2) ss.deleteSheet(s1);
  return ss;
}

// ---------- Состояние диалога ----------

function getState(chatId) {
  var raw = CacheService.getScriptCache().get('st_' + chatId);
  return raw ? JSON.parse(raw) : {};
}

function setState(chatId, st) {
  CacheService.getScriptCache().put('st_' + chatId, JSON.stringify(st), 21600);
}

// ---------- Клавиатуры ----------

function kbMain() {
  return [
    [{ text: '🧺 Сдать в прачку', callback_data: 'op|send' }],
    [{ text: '📦 Приняли из прачки', callback_data: 'op|ret' }],
    [{ text: '📊 Остатки', callback_data: 'op|bal' }]
  ];
}

function kbObjects() {
  var kb = [];
  for (var i = 0; i < OBJECTS.length; i += 2) {
    var row = [{ text: OBJECTS[i], callback_data: 'ob|' + i }];
    if (OBJECTS[i + 1]) row.push({ text: OBJECTS[i + 1], callback_data: 'ob|' + (i + 1) });
    kb.push(row);
  }
  kb.push([{ text: '✖️ Отмена', callback_data: 'cancel' }]);
  return kb;
}

function kbItems(cart) {
  var kb = [];
  for (var i = 0; i < ITEMS.length; i += 2) {
    var row = [];
    for (var d = 0; d < 2; d++) {
      var j = i + d;
      if (j < ITEMS.length) {
        var n = cart[j] ? ' ✅' + cart[j] : '';
        row.push({ text: ITEMS[j] + n, callback_data: 'it|' + j });
      }
    }
    kb.push(row);
  }
  kb.push([{ text: '✅ Готово', callback_data: 'done' }, { text: '✖️ Отмена', callback_data: 'cancel' }]);
  return kb;
}

function kbQty() {
  return [
    [1, 2, 3, 4, 5].map(function (n) { return { text: String(n), callback_data: 'q|' + n }; }),
    [6, 7, 8, 10, 12].map(function (n) { return { text: String(n), callback_data: 'q|' + n }; }),
    [{ text: '← Назад к списку', callback_data: 'back' }]
  ];
}

// ---------- Остатки ----------

function balanceText(obIdx) {
  var ss = getSheetFile();
  var obj = OBJECTS[obIdx];
  var data = ss.getSheetByName('Журнал').getDataRange().getValues();
  var laundry = {};
  for (var r = 1; r < data.length; r++) {
    if (data[r][2] !== obj) continue;
    var item = data[r][3], qty = Number(data[r][4]) || 0;
    if (laundry[item] === undefined) laundry[item] = 0;
    if (data[r][1] === 'Сдано в прачку') laundry[item] += qty;
    if (data[r][1] === 'Принято из прачки') laundry[item] -= qty;
  }
  var totals = {};
  var tdata = ss.getSheetByName('Всего').getDataRange().getValues();
  for (var r2 = 1; r2 < tdata.length; r2++)
    if (tdata[r2][0] === obj) totals[tdata[r2][1]] = Number(tdata[r2][2]) || 0;

  var out = '📊 <b>' + obj + '</b>\n\n';
  var anyLaundry = false;
  for (var k = 0; k < ITEMS.length; k++) {
    var it = ITEMS[k];
    var inL = laundry[it] || 0;
    var tot = totals[it] || 0;
    if (inL === 0 && tot === 0) continue;
    var clean = tot > 0 ? Math.max(tot - inL, 0) : null;
    out += '• ' + it + ': в прачке <b>' + inL + '</b>';
    if (clean !== null) out += ', чистое <b>' + clean + '</b>';
    out += '\n';
    if (inL > 0) anyLaundry = true;
  }
  if (out.indexOf('•') < 0) out += 'Пока нет данных — сдайте первую накладную 🙂';
  else if (!anyLaundry) out += '\nВсё бельё на объекте, в прачке ничего нет ✅';
  return out;
}

// ---------- Главный обработчик ----------

function doPost(e) {
  try {
    var u = JSON.parse(e.postData.contents);
    if (u.message) handleMessage(u.message);
    if (u.callback_query) handleCallback(u.callback_query);
  } catch (err) {
    // молчим, чтобы Telegram не повторял запрос бесконечно
  }
  return HtmlService.createHtmlOutput('ok');
}

function handleMessage(msg) {
  var chatId = msg.chat.id;
  send(chatId, '👋 Привет! Я помощник OLEAND.\nЧто делаем?', kbMain());
}

function handleCallback(cb) {
  var chatId = cb.message.chat.id;
  var who = cb.from.first_name || '';
  var d = cb.data.split('|');
  var st = getState(chatId);
  tg('answerCallbackQuery', { callback_query_id: cb.id });

  if (d[0] === 'cancel') {
    setState(chatId, {});
    send(chatId, 'Отменено. Что делаем?', kbMain());

  } else if (d[0] === 'op') {
    st = { op: d[1], cart: {} };
    setState(chatId, st);
    var title = d[1] === 'send' ? '🧺 Сдать в прачку' : d[1] === 'ret' ? '📦 Приняли из прачки' : '📊 Остатки';
    send(chatId, title + '\nВыберите объект:', kbObjects());

  } else if (d[0] === 'ob') {
    st.ob = Number(d[1]);
    setState(chatId, st);
    if (st.op === 'bal') {
      send(chatId, balanceText(st.ob), kbMain());
      setState(chatId, {});
    } else {
      send(chatId, '🏠 ' + OBJECTS[st.ob] + '\nВыберите позицию:', kbItems(st.cart));
    }

  } else if (d[0] === 'it') {
    st.item = Number(d[1]);
    setState(chatId, st);
    send(chatId, ITEMS[st.item] + ' — сколько штук?', kbQty());

  } else if (d[0] === 'q') {
    var q = Number(d[1]);
    st.cart[st.item] = (st.cart[st.item] || 0) + q;
    setState(chatId, st);
    send(chatId, '✅ ' + ITEMS[st.item] + ' — ' + st.cart[st.item] + ' шт.\nЕщё позиции?', kbItems(st.cart));

  } else if (d[0] === 'back') {
    send(chatId, 'Выберите позицию:', kbItems(st.cart));

  } else if (d[0] === 'done') {
    var keys = Object.keys(st.cart || {});
    if (!keys.length || st.ob === undefined) {
      send(chatId, 'Ничего не выбрано. Что делаем?', kbMain());
      setState(chatId, {});
      return;
    }
    var ss = getSheetFile();
    var j = ss.getSheetByName('Журнал');
    var opName = st.op === 'send' ? 'Сдано в прачку' : 'Принято из прачки';
    var now = new Date();
    var total = 0, lines = '';
    keys.forEach(function (k) {
      var qty = st.cart[k];
      total += qty;
      lines += '• ' + ITEMS[k] + ' — ' + qty + ' шт\n';
      j.appendRow([now, opName, OBJECTS[st.ob], ITEMS[k], qty, who]);
    });
    var head = st.op === 'send' ? '🧺 Сдано в прачку' : '📦 Принято из прачки';
    send(chatId,
      head + '\n🏠 ' + OBJECTS[st.ob] + '\n\n' + lines + '\n📦 Итого: ' + total + ' шт\n👤 ' + who,
      kbMain());
    setState(chatId, {});
  }
}

// ---------- Запустить ОДИН раз после развёртывания ----------

function setup() {
  getSheetFile();
  var url = ScriptApp.getService().getUrl();
  var resp = tg('setWebhook', { url: url });
  Logger.log(resp.getContentText());
}

// Проверить состояние webhook (показывает очередь и ошибки Telegram)
function checkWebhook() { Logger.log(tg('getWebhookInfo', {}).getContentText()); }

// Привязать бота к рабочему адресу /exec (запускать после каждого нового развёртывания)
function fixWebhook() { var u = 'https://script.google.com/macros/s/AKfycbw5yFaycoIs3tgKE7U2uElBz8UuxdUcOY7RxLABN_fenDEut2E2lqBWltO-1lo-Dz_CNA/exec'; Logger.log(tg('setWebhook', { url: u, drop_pending_updates: true }).getContentText()); Logger.log(tg('getWebhookInfo', {}).getContentText()); }

// Проверить, что ссылка календаря отвечает (смотрим длину и начало файла)
function testCal() { var t = UrlFetchApp.fetch('https://realtycalendar.ru/apartments/export.ics?q=MTQxMTQx').getContentText(); Logger.log('LEN=' + t.length); Logger.log(t.slice(0, 3000)); }

// ========== ЭТАП 2: КАЛЕНДАРЬ ЗАСЕЛЕНИЙ (RealtyCalendar) ==========
// Новые kbMain и doPost ниже заменяют старые версии выше — так задумано

var CALENDARS = [{ obj: 'Царская столица', url: 'https://realtycalendar.ru/apartments/export.ics?q=MTQxMTQx' }];

function parseIcs(t) { var ev = []; var b = t.split('BEGIN:VEVENT'); for (var i = 1; i < b.length; i++) { var s = b[i].match(/DTSTART[^:]*:(\d{8})/); var e2 = b[i].match(/DTEND[^:]*:(\d{8})/); if (s && e2) ev.push({ s: s[1], e: e2[1] }); } return ev; }

function planText() { var d = new Date(); var today = Utilities.formatDate(d, 'Europe/Moscow', 'yyyyMMdd'); var out = '📅 <b>План на ' + Utilities.formatDate(d, 'Europe/Moscow', 'dd.MM.yyyy') + '</b>\n'; var any = false; for (var c = 0; c < CALENDARS.length; c++) { var txt = ''; try { txt = UrlFetchApp.fetch(CALENDARS[c].url).getContentText(); } catch (er) { out += '\n⚠ ' + CALENDARS[c].obj + ': календарь недоступен\n'; continue; } var ev = parseIcs(txt); var vin = 0; var vout = 0; for (var k = 0; k < ev.length; k++) { if (ev[k].s === today) vin++; if (ev[k].e === today) vout++; } if (vin || vout) { any = true; out += '\n🏠 <b>' + CALENDARS[c].obj + '</b>\n'; if (vout) out += '🔴 Выезд сегодня: ' + vout + ' — уборка и смена белья\n'; if (vin) out += '🟢 Заезд сегодня: ' + vin + '\n'; } } if (!any) out += '\nСегодня заездов и выездов нет ✅'; return out; }

function addPlanSub(chatId) { var p = PropertiesService.getScriptProperties(); var s = p.getProperty('PLAN_SUBS') || ''; if (s.indexOf('[' + chatId + ']') < 0) p.setProperty('PLAN_SUBS', s + '[' + chatId + ']'); }

function sendMorningPlan() { var s = PropertiesService.getScriptProperties().getProperty('PLAN_SUBS') || ''; var ids = s.match(/-?\d+/g) || []; var t = planText(); for (var i = 0; i < ids.length; i++) send(ids[i], t, kbMain()); }

// Запустить ОДИН раз: создаёт ежедневную рассылку плана в 9 утра
function setupMorningTrigger() { var tr = ScriptApp.getProjectTriggers(); for (var i = 0; i < tr.length; i++) if (tr[i].getHandlerFunction() === 'sendMorningPlan') ScriptApp.deleteTrigger(tr[i]); ScriptApp.newTrigger('sendMorningPlan').timeBased().atHour(9).everyDays(1).create(); Logger.log('OK: триггер на 9 утра создан'); }

// Посмотреть план в журнале, не отправляя в Telegram
function testPlan() { Logger.log(planText()); }

function kbMain() { return [[{ text: '📅 План на сегодня', callback_data: 'pl|0' }], [{ text: '🧺 Сдать в прачку', callback_data: 'op|send' }], [{ text: '📦 Приняли из прачки', callback_data: 'op|ret' }], [{ text: '📊 Остатки', callback_data: 'op|bal' }]]; }

function doPost(e) { try { var u = JSON.parse(e.postData.contents); if (u.callback_query && u.callback_query.data === 'pl|0') { tg('answerCallbackQuery', { callback_query_id: u.callback_query.id }); addPlanSub(u.callback_query.message.chat.id); send(u.callback_query.message.chat.id, planText(), kbMain()); return HtmlService.createHtmlOutput('ok'); } if (u.message) { addPlanSub(u.message.chat.id); handleMessage(u.message); } if (u.callback_query) handleCallback(u.callback_query); } catch (err) {} return HtmlService.createHtmlOutput('ok'); }

// ========== ЭТАП 3: ФИРМЕННАЯ ФОРМА НАКЛАДНОЙ (Mini App) ==========
// Новые kbMain и doPost ниже снова заменяют версии выше — так задумано

var FORM_URL = 'https://qnesterowa-hash.github.io/oleand/nak%D0%BB%D0%B0%D0%B4%D0%BD%D0%B0%D1%8F_v2.html';

function kbMain() { return [[{ text: '📅 План на сегодня', callback_data: 'pl|0' }], [{ text: '🧺 Сдать в прачку', web_app: { url: FORM_URL } }], [{ text: '📦 Приняли из прачки', callback_data: 'op|ret' }], [{ text: '📊 Остатки', callback_data: 'op|bal' }]]; }

function doPost(e) { try { var u = JSON.parse(e.postData.contents); if (u.oleand === 'nakl' && u.chat && u.text) { tg('sendMessage', { chat_id: u.chat, text: u.text }); return HtmlService.createHtmlOutput('ok'); } if (u.callback_query && u.callback_query.data === 'pl|0') { tg('answerCallbackQuery', { callback_query_id: u.callback_query.id }); addPlanSub(u.callback_query.message.chat.id); send(u.callback_query.message.chat.id, planText(), kbMain()); return HtmlService.createHtmlOutput('ok'); } if (u.message) { addPlanSub(u.message.chat.id); handleMessage(u.message); } if (u.callback_query) handleCallback(u.callback_query); } catch (err) {} return HtmlService.createHtmlOutput('ok'); }

// ========== ЭТАП 4: ПРИЁМ С ГАЛОЧКАМИ И ОСТАТКИ (Mini App) ==========
var PRIEM_URL = 'https://qnesterowa-hash.github.io/oleand/priem.html';
var OSTATKI_URL = 'https://qnesterowa-hash.github.io/oleand/ostatki.html';
function getLaundrySheet() { var ss = getSheetFile(); var sh = ss.getSheetByName('Прачка'); if (!sh) { sh = ss.insertSheet('Прачка'); sh.appendRow(['ID', 'Дата сдачи', 'Объект', 'Позиция', 'Кол-во', 'Статус', 'Дата приёма']); } return sh; }
function saveNakl(obj, itemsArr, who) { var sh = getLaundrySheet(); var j = getSheetFile().getSheetByName('Журнал'); var d = Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm'); for (var i = 0; i < itemsArr.length; i++) { var id = Date.now() + '_' + i; sh.appendRow([id, d, obj, itemsArr[i].name, itemsArr[i].qty, 'в прачке', '']); j.appendRow([d, 'сдано в прачку', obj, itemsArr[i].name, itemsArr[i].qty, who || '']); } }
function pendingJson() { var v = getLaundrySheet().getDataRange().getValues(); var out = []; for (var i = 1; i < v.length; i++) if (v[i][5] === 'в прачке') out.push({ id: String(v[i][0]), date: String(v[i][1]), obj: v[i][2], name: v[i][3], qty: v[i][4] }); return out; }
function receiveIds(ids, chat) { var sh = getLaundrySheet(); var v = sh.getDataRange().getValues(); var j = getSheetFile().getSheetByName('Журнал'); var d = Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm'); var got = {}; for (var i = 1; i < v.length; i++) { if (ids.indexOf(String(v[i][0])) >= 0 && v[i][5] === 'в прачке') { sh.getRange(i + 1, 6).setValue('принято'); sh.getRange(i + 1, 7).setValue(d); j.appendRow([d, 'принято из прачки', v[i][2], v[i][3], v[i][4], '']); if (!got[v[i][2]]) got[v[i][2]] = []; got[v[i][2]].push(' • ' + v[i][3] + ' — ' + v[i][4] + ' шт'); } } var t = '📦 ПРИНЯТО ИЗ ПРАЧКИ\n' + Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy') + '\n'; for (var o in got) t += '\n🏠 ' + o + '\n' + got[o].join('\n') + '\n'; send(chat, t, kbMain()); }
function balanceJson() { var tv = getSheetFile().getSheetByName('Всего').getDataRange().getValues(); var pend = pendingJson(); var pmap = {}; for (var i = 0; i < pend.length; i++) { var k = pend[i].obj + '|' + pend[i].name; pmap[k] = (pmap[k] || 0) + Number(pend[i].qty); } var res = {}; for (var r = 1; r < tv.length; r++) { var ob = tv[r][0]; if (!res[ob]) res[ob] = []; res[ob].push({ name: tv[r][1], total: tv[r][2], laundry: pmap[ob + '|' + tv[r][1]] || 0 }); } var arr = []; for (var o2 in res) arr.push({ obj: o2, items: res[o2] }); return arr; }
function jsonOut(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function fmtCell(x) { return (x && x.getTime) ? Utilities.formatDate(x, 'Europe/Moscow', 'dd.MM.yyyy HH:mm') : String(x || ''); }
function doGet(e) {
  var api = e && e.parameter ? e.parameter.api : '';
  if (api === 'pending') return jsonOut({ ok: true, items: pendingJson() });
  if (api === 'balance') return jsonOut({ ok: true, objects: balanceJson() });
  if (api === 'nakl') { // фирменная накладная-черновик по ID (ЭТАП 9)
    var id = e.parameter.id || '';
    var v = getLaundrySheet().getDataRange().getValues();
    var items = [], obj = '', date = '', status = '';
    for (var i = 1; i < v.length; i++)
      if (String(v[i][0]).indexOf(id + '_') === 0) {
        obj = v[i][2]; date = fmtCell(v[i][1]); status = String(v[i][5]);
        if (Number(v[i][4]) > 0) items.push({ name: v[i][3], qty: Number(v[i][4]) });
      }
    return jsonOut({ ok: true, type: 'draft', id: id, obj: obj, date: date, status: status, items: items });
  }
  if (api === 'ship') { // фирменная накладная-факт отгрузки по ID (ЭТАП 9)
    var id2 = e.parameter.id || '';
    var s = getShipSheet().getDataRange().getValues();
    var items2 = [], date2 = '';
    for (var r = 1; r < s.length; r++)
      if (String(s[r][5]) === id2) { date2 = fmtCell(s[r][0]); items2.push({ obj: s[r][2], name: s[r][3], qty: Number(s[r][4]) }); }
    return jsonOut({ ok: true, type: 'ship', id: id2, date: date2, items: items2 });
  }
  return ContentService.createTextOutput('OLEAND bot ok');
}
function kbMain() { return [[{ text: '📅 План на сегодня', callback_data: 'pl|0' }], [{ text: '🧺 Сдать в прачку', web_app: { url: FORM_URL } }], [{ text: '📦 Приняли из прачки', web_app: { url: PRIEM_URL } }], [{ text: '📊 Остатки', web_app: { url: OSTATKI_URL } }]]; }
function doPost(e) { try { var u = JSON.parse(e.postData.contents); if (u.oleand === 'nakl' && u.chat && u.text) { if (u.obj && u.items) saveNakl(u.obj, u.items, u.who); tg('sendMessage', { chat_id: u.chat, text: u.text }); return HtmlService.createHtmlOutput('ok'); } if (u.oleand === 'recv' && u.chat && u.ids) { receiveIds(u.ids, u.chat); return HtmlService.createHtmlOutput('ok'); } if (u.callback_query && u.callback_query.data === 'pl|0') { tg('answerCallbackQuery', { callback_query_id: u.callback_query.id }); addPlanSub(u.callback_query.message.chat.id); send(u.callback_query.message.chat.id, planText(), kbMain()); return HtmlService.createHtmlOutput('ok'); } if (u.message) { addPlanSub(u.message.chat.id); handleMessage(u.message); } if (u.callback_query) handleCallback(u.callback_query); } catch (err) {} return HtmlService.createHtmlOutput('ok'); }

// ЭТАП 5: все календари RealtyCalendar (Петроградка и Саперный — по 2 квартиры в плане, бельё общее)
var CALENDARS = [{ obj: 'Царская столица', url: 'https://realtycalendar.ru/apartments/export.ics?q=MTQxMTQx' }, { obj: 'Петроградка Малая Посадская 1', url: 'https://realtycalendar.ru/apartments/export.ics?q=MjMwOTUy' }, { obj: 'Петроградка Малая Посадская 2', url: 'https://realtycalendar.ru/apartments/export.ics?q=MjMwOTc3' }, { obj: 'Саперный 1 маленькая', url: 'https://realtycalendar.ru/apartments/export.ics?q=MTQxNDU0' }, { obj: 'Саперный 2 большая', url: 'https://realtycalendar.ru/apartments/export.ics?q=MTQxNDUy' }, { obj: 'BLUE Железнодорожный', url: '' }, { obj: 'GREEN Тосина', url: 'https://realtycalendar.ru/apartments/export.ics?q=MzEzOTM5' }, { obj: 'NUDE Манчестерская', url: 'https://realtycalendar.ru/apartments/export.ics?q=Mjg0NjA2' }, { obj: 'TIFFANY Ветеранов', url: 'https://realtycalendar.ru/apartments/export.ics?q=MzY0ODM3' }, { obj: 'GRAY Лиговский проспект', url: 'https://realtycalendar.ru/apartments/export.ics?q=MzY1Mjk5' }];

// ЭТАП 5.1: объекты без ссылки пропускаем (BLUE Железнодорожный пока в стопе)
function planText() { var d = new Date(); var today = Utilities.formatDate(d, 'Europe/Moscow', 'yyyyMMdd'); var out = '📅 <b>План на ' + Utilities.formatDate(d, 'Europe/Moscow', 'dd.MM.yyyy') + '</b>\n'; var any = false; for (var c = 0; c < CALENDARS.length; c++) { if (!CALENDARS[c].url) continue; var txt = ''; try { txt = UrlFetchApp.fetch(CALENDARS[c].url).getContentText(); } catch (er) { out += '\n⚠ ' + CALENDARS[c].obj + ': календарь недоступен\n'; continue; } var ev = parseIcs(txt); var vin = 0; var vout = 0; for (var k = 0; k < ev.length; k++) { if (ev[k].s === today) vin++; if (ev[k].e === today) vout++; } if (vin || vout) { any = true; out += '\n🏠 <b>' + CALENDARS[c].obj + '</b>\n'; if (vout) out += '🔴 Выезд сегодня: ' + vout + ' — уборка и смена белья\n'; if (vin) out += '🟢 Заезд сегодня: ' + vin + '\n'; } } if (!any) out += '\nСегодня заездов и выездов нет ✅'; return out; }

// ========== ЭТАП 6: РОЛИ — ХОЗЯЙКА / ГОРНИЧНЫЕ / ПРАЧКА ==========
function roleOf(id) { var p = PropertiesService.getScriptProperties(); id = String(id); if (p.getProperty('ADMIN_CHAT') === id) return 'admin'; if (p.getProperty('LAUNDRY_CHAT') === id) return 'laundry'; if ((p.getProperty('MAID_CHATS') || '').indexOf('[' + id + ']') > -1) return 'maid'; return 'guest'; }
function kbFor(id) { var r = roleOf(id); if (r === 'laundry' || r === 'guest') return null; var rows = []; if (r === 'admin') rows.push([{ text: '📅 План на сегодня', callback_data: 'pl|0' }]); rows.push([{ text: '🧺 Сдать в прачку', web_app: { url: FORM_URL } }]); rows.push([{ text: '📦 Приняли из прачки', web_app: { url: PRIEM_URL } }]); rows.push([{ text: '📊 Остатки', web_app: { url: OSTATKI_URL } }]); return rows; }
function maidName(o) { var M = { 'Саперный 1 маленькая': 'Саперный — уборка маленькая квартира', 'Саперный 2 большая': 'Саперный — уборка большая квартира', 'Петроградка Малая Посадская 1': 'Петроградка — уборка маленькая квартира', 'Петроградка Малая Посадская 2': 'Петроградка — уборка большая квартира' }; return M[o] || (o + ' — уборка'); }
function maidText() { var d = new Date(); var today = Utilities.formatDate(d, 'Europe/Moscow', 'yyyyMMdd'); var out = '🧹 <b>Уборки на ' + Utilities.formatDate(d, 'Europe/Moscow', 'dd.MM.yyyy') + '</b>\n'; var any = false; for (var c = 0; c < CALENDARS.length; c++) { if (!CALENDARS[c].url) continue; var txt = ''; try { txt = UrlFetchApp.fetch(CALENDARS[c].url).getContentText(); } catch (er) { continue; } var ev = parseIcs(txt); var vout = 0; var vin = 0; for (var k = 0; k < ev.length; k++) { if (ev[k].e === today) vout++; if (ev[k].s === today) vin++; } if (vout) { any = true; out += '\n• ' + maidName(CALENDARS[c].obj) + (vin ? ' — сегодня заезд, убрать к заселению!' : '') + '\n'; } } if (!any) out += '\nСегодня уборок нет ✅'; return out; }
function sendMorningPlan() { var p = PropertiesService.getScriptProperties(); var a = p.getProperty('ADMIN_CHAT'); if (a) { send(a, planText(), kbFor(a)); } else { var s = (p.getProperty('PLAN_SUBS') || '').match(/-?\d+/g) || []; if (s.length) { var pt = planText(); for (var i = 0; i < s.length; i++) send(s[i], pt); } } var m = (p.getProperty('MAID_CHATS') || '').match(/-?\d+/g) || []; if (m.length) { var t = maidText(); for (var j = 0; j < m.length; j++) send(m[j], t, kbFor(m[j])); } }
function doPost(e) { try { var u = JSON.parse(e.postData.contents); if (u.oleand === 'nakl' && u.chat && u.text) { if (u.obj && u.items) saveNakl(u.obj, u.items, u.who); tg('sendMessage', { chat_id: u.chat, text: u.text }); var L = PropertiesService.getScriptProperties().getProperty('LAUNDRY_CHAT'); if (L && String(L) !== String(u.chat)) tg('sendMessage', { chat_id: L, text: u.text }); return HtmlService.createHtmlOutput('ok'); } if (u.oleand === 'recv' && u.chat && u.ids) { receiveIds(u.ids, u.chat); return HtmlService.createHtmlOutput('ok'); } if (u.callback_query && u.callback_query.data === 'pl|0') { tg('answerCallbackQuery', { callback_query_id: u.callback_query.id }); var cid = u.callback_query.message.chat.id; if (roleOf(cid) === 'admin') { addPlanSub(cid); send(cid, planText(), kbFor(cid)); } else { send(cid, 'План на день доступен только хозяйке 🙂', kbFor(cid)); } return HtmlService.createHtmlOutput('ok'); } if (u.message) { var id = u.message.chat.id; var t = String(u.message.text || '').toLowerCase().trim(); var p = PropertiesService.getScriptProperties(); if (t === 'я хозяйка') { p.setProperty('ADMIN_CHAT', String(id)); send(id, 'Готово! Вы — хозяйка 👑 Вам доступно всё, включая план на день.', kbFor(id)); return HtmlService.createHtmlOutput('ok'); } if (t === 'я горничная') { var m = p.getProperty('MAID_CHATS') || ''; if (m.indexOf('[' + id + ']') < 0) p.setProperty('MAID_CHATS', m + '[' + id + ']'); send(id, 'Готово! Вы — горничная 🧹 Каждое утро сюда будут приходить уборки на день.', kbFor(id)); return HtmlService.createHtmlOutput('ok'); } if (t === 'я прачка') { p.setProperty('LAUNDRY_CHAT', String(id)); send(id, 'Готово! Вы — прачка 🧺 Накладные будут приходить в этот чат автоматически.'); return HtmlService.createHtmlOutput('ok'); } var r = roleOf(id); if (r === 'laundry') { send(id, 'Накладные будут приходить сюда автоматически 🧺'); } else if (r === 'guest') { send(id, 'Привет! Это бот OLEAND. Напишите кодовое слово, которое вам дали.'); } else { send(id, 'Меню 👇', kbFor(id)); } return HtmlService.createHtmlOutput('ok'); } if (u.callback_query) handleCallback(u.callback_query); } catch (err) {} return HtmlService.createHtmlOutput('ok'); }
// Посмотреть текст уборок для горничных в журнале, не отправляя
function testMaid() { Logger.log(maidText()); }

// ========== ЭТАП 6.1: НЕСКОЛЬКО ХОЗЯЕК / МЕНЕДЖЕРОВ ==========
function addAdmin(id) { var p = PropertiesService.getScriptProperties(); var a = p.getProperty('ADMIN_CHATS') || ''; if (a.indexOf('[' + id + ']') < 0) p.setProperty('ADMIN_CHATS', a + '[' + id + ']'); }
function roleOf(id) { var p = PropertiesService.getScriptProperties(); id = String(id); if (p.getProperty('ADMIN_CHAT') === id) return 'admin'; if ((p.getProperty('ADMIN_CHATS') || '').indexOf('[' + id + ']') > -1) return 'admin'; if (p.getProperty('LAUNDRY_CHAT') === id) return 'laundry'; if ((p.getProperty('MAID_CHATS') || '').indexOf('[' + id + ']') > -1) return 'maid'; return 'guest'; }
function adminIds() { var p = PropertiesService.getScriptProperties(); var a = (p.getProperty('ADMIN_CHATS') || '').match(/-?\d+/g) || []; var one = p.getProperty('ADMIN_CHAT'); if (one && a.indexOf(one) < 0) a.push(one); return a; }
function sendMorningPlan() { var a = adminIds(); if (a.length) { var pt = planText(); for (var i = 0; i < a.length; i++) send(a[i], pt, kbFor(a[i])); } var p = PropertiesService.getScriptProperties(); var m = (p.getProperty('MAID_CHATS') || '').match(/-?\d+/g) || []; if (m.length) { var t = maidText(); for (var j = 0; j < m.length; j++) send(m[j], t, kbFor(m[j])); } }
function doPost(e) { try { var u = JSON.parse(e.postData.contents); if (u.oleand === 'nakl' && u.chat && u.text) { if (u.obj && u.items) saveNakl(u.obj, u.items, u.who); tg('sendMessage', { chat_id: u.chat, text: u.text }); var L = PropertiesService.getScriptProperties().getProperty('LAUNDRY_CHAT'); if (L && String(L) !== String(u.chat)) tg('sendMessage', { chat_id: L, text: u.text }); return HtmlService.createHtmlOutput('ok'); } if (u.oleand === 'recv' && u.chat && u.ids) { receiveIds(u.ids, u.chat); return HtmlService.createHtmlOutput('ok'); } if (u.callback_query && u.callback_query.data === 'pl|0') { tg('answerCallbackQuery', { callback_query_id: u.callback_query.id }); var cid = u.callback_query.message.chat.id; if (roleOf(cid) === 'admin') { addPlanSub(cid); send(cid, planText(), kbFor(cid)); } else { send(cid, 'План на день доступен только хозяйке и менеджеру 🙂', kbFor(cid)); } return HtmlService.createHtmlOutput('ok'); } if (u.message) { var id = u.message.chat.id; var t = String(u.message.text || '').toLowerCase().trim(); var p = PropertiesService.getScriptProperties(); if (t === 'я хозяйка' || t === 'я менеджер') { addAdmin(String(id)); send(id, 'Готово! 👑 Вам доступно всё, включая план на день.', kbFor(id)); return HtmlService.createHtmlOutput('ok'); } if (t === 'я горничная') { var m = p.getProperty('MAID_CHATS') || ''; if (m.indexOf('[' + id + ']') < 0) p.setProperty('MAID_CHATS', m + '[' + id + ']'); send(id, 'Готово! Вы — горничная 🧹 Каждое утро сюда будут приходить уборки на день.', kbFor(id)); return HtmlService.createHtmlOutput('ok'); } if (t === 'я прачка') { p.setProperty('LAUNDRY_CHAT', String(id)); send(id, 'Готово! Вы — прачка 🧺 Накладные будут приходить в этот чат автоматически.'); return HtmlService.createHtmlOutput('ok'); } var r = roleOf(id); if (r === 'laundry') { send(id, 'Накладные будут приходить сюда автоматически 🧺'); } else if (r === 'guest') { send(id, 'Привет! Это бот OLEAND. Напишите кодовое слово, которое вам дали.'); } else { send(id, 'Меню 👇', kbFor(id)); } return HtmlService.createHtmlOutput('ok'); } if (u.callback_query) handleCallback(u.callback_query); } catch (err) {} return HtmlService.createHtmlOutput('ok'); }

// ========== ЭТАП 7: НАПОМИНАНИЕ О ЗАВТРАШНИХ ВЫЕЗДАХ (13:00) ==========
var SMS_GUEST = 'Здравствуйте! Надеемся, вам у нас понравилось 🤍 Завтра — день вашего выезда. Если хочется остаться подольше — просто напишите, и мы с радостью продлим бронь. Если планируете выезжать, подскажите, пожалуйста, во сколько. Ждём новых встреч!\n\nС любовью и заботой,\nOLEAND Apartments';
function checkoutText() { var d = new Date(Date.now() + 86400000); var tom = Utilities.formatDate(d, 'Europe/Moscow', 'yyyyMMdd'); var dd = Utilities.formatDate(d, 'Europe/Moscow', 'dd.MM.yyyy'); var list = []; for (var c = 0; c < CALENDARS.length; c++) { if (!CALENDARS[c].url) continue; var txt = ''; try { txt = UrlFetchApp.fetch(CALENDARS[c].url).getContentText(); } catch (er) { continue; } var ev = parseIcs(txt); var n = 0; for (var k = 0; k < ev.length; k++) { if (ev[k].e === tom) n++; } if (n) list.push('🏠 ' + CALENDARS[c].obj + (n > 1 ? ' — выездов: ' + n : '')); } if (!list.length) return ''; return '🔔 <b>Завтра (' + dd + ') выезжают гости:</b>\n\n' + list.join('\n') + '\n\n✍️ Отправьте гостям сообщение — предложите продлить или уточните время выезда. Готовый текст следующим сообщением 👇'; }
function sendCheckoutReminder() { var t = checkoutText(); if (!t) return; var a = adminIds(); for (var i = 0; i < a.length; i++) { send(a[i], t); send(a[i], SMS_GUEST); } }
function setupCheckoutTrigger() { var tr = ScriptApp.getProjectTriggers(); for (var i = 0; i < tr.length; i++) { if (tr[i].getHandlerFunction() === 'sendCheckoutReminder') ScriptApp.deleteTrigger(tr[i]); } ScriptApp.newTrigger('sendCheckoutReminder').timeBased().everyDays(1).atHour(13).create(); Logger.log('Триггер на 13:00 создан'); }
function testCheckout() { Logger.log(checkoutText() || 'Завтра выездов нет'); }

// ========== ЭТАП 8: АВТО-НАКЛАДНЫЕ ПО КАЛЕНДАРЮ ==========
// Полный цикл: выезд по RealtyCalendar → черновик накладной по норме объекта →
// горничная подтверждает или убавляет → прачка подтверждает забор → накладная-факт →
// запись в лист «Отгрузки» → 1-го числа месяца сверка приходит в бот.
// После вставки: Развернуть новую ВЕРСИЮ того же развёртывания (адрес /exec не менять!)
// и запустить ОДИН раз setupStage8().

// Позиции «евро» для кровати 260 в Сапёрном 2 — добавлены в конец списка
var ITEMS = [
  'Пододеяльник 2,0', 'Простыня 2,0', 'Наволочка 2,0',
  'Пододеяльник 1,5', 'Простыня 1,5', 'Наволочка 1,5',
  'Полотенце большое', 'Полотенце среднее', 'Полотенце малютка',
  'Полотенце для ног', 'Салфетка кухонная', 'Салфетка (клетка)',
  'Пододеяльник евро', 'Простыня евро'
];

// Нормы = МАКСИМАЛЬНЫЙ комплект на объект. Горничная убавляет, если гостей было меньше.
// Названия объектов = названия в CALENDARS, названия позиций = списку ITEMS (менять нельзя!)
var NORMS = {
  'GRAY Лиговский проспект': { 'Пододеяльник 2,0': 3, 'Простыня 2,0': 3, 'Наволочка 2,0': 6, 'Полотенце большое': 6, 'Полотенце среднее': 6, 'Полотенце для ног': 2, 'Салфетка кухонная': 1 },
  'Петроградка Малая Посадская 1': { 'Пододеяльник 2,0': 1, 'Пододеяльник 1,5': 1, 'Простыня 2,0': 1, 'Простыня 1,5': 1, 'Наволочка 2,0': 3, 'Полотенце большое': 3, 'Полотенце среднее': 3, 'Полотенце для ног': 1, 'Салфетка кухонная': 1 },
  'Петроградка Малая Посадская 2': { 'Пододеяльник 2,0': 1, 'Пододеяльник 1,5': 1, 'Простыня 2,0': 1, 'Простыня 1,5': 1, 'Наволочка 2,0': 3, 'Полотенце большое': 3, 'Полотенце среднее': 3, 'Полотенце для ног': 1, 'Салфетка кухонная': 1 },
  'Саперный 1 маленькая': { 'Пододеяльник 2,0': 1, 'Простыня 2,0': 1, 'Наволочка 2,0': 2, 'Полотенце большое': 2, 'Полотенце среднее': 2, 'Полотенце для ног': 1, 'Салфетка кухонная': 1 },
  'Саперный 2 большая': { 'Пододеяльник евро': 1, 'Простыня евро': 1, 'Пододеяльник 2,0': 1, 'Простыня 2,0': 1, 'Наволочка 2,0': 4, 'Полотенце большое': 4, 'Полотенце среднее': 4, 'Полотенце для ног': 2, 'Салфетка кухонная': 1 },
  'TIFFANY Ветеранов': { 'Пододеяльник 2,0': 1, 'Пододеяльник 1,5': 1, 'Простыня 2,0': 1, 'Простыня 1,5': 1, 'Наволочка 2,0': 3, 'Полотенце большое': 3, 'Полотенце среднее': 3, 'Полотенце для ног': 1, 'Салфетка кухонная': 1 },
  'Царская столица': { 'Пододеяльник 2,0': 1, 'Пододеяльник 1,5': 3, 'Простыня 2,0': 1, 'Простыня 1,5': 3, 'Наволочка 2,0': 6, 'Полотенце большое': 6, 'Полотенце среднее': 6, 'Полотенце для ног': 2, 'Салфетка кухонная': 1 },
  'GREEN Тосина': { 'Пододеяльник 2,0': 1, 'Пододеяльник 1,5': 1, 'Простыня 2,0': 1, 'Простыня 1,5': 1, 'Наволочка 2,0': 3, 'Полотенце большое': 3, 'Полотенце среднее': 3, 'Полотенце для ног': 1, 'Салфетка кухонная': 1 },
  'BLUE Железнодорожный': { 'Пододеяльник 2,0': 2, 'Простыня 2,0': 2, 'Наволочка 2,0': 4, 'Полотенце большое': 4, 'Полотенце среднее': 4, 'Полотенце для ног': 2, 'Салфетка кухонная': 1 },
  'NUDE Манчестерская': { 'Пододеяльник 2,0': 3, 'Простыня 2,0': 3, 'Наволочка 2,0': 6, 'Полотенце большое': 6, 'Полотенце среднее': 6, 'Полотенце для ног': 2, 'Салфетка кухонная': 1 }
};

// Фирменная страница накладной (цвета бренда) на GitHub Pages
var VIEW_URL = 'https://qnesterowa-hash.github.io/oleand/nakladnaya.html';
function kbView(id) { return [[{ text: '🧾 Открыть фирменную накладную', web_app: { url: VIEW_URL + '?id=' + id } }]]; }

// ---------- Лист «Отгрузки» (факты отгрузки белья в прачку) ----------
function getShipSheet() {
  var ss = getSheetFile();
  var sh = ss.getSheetByName('Отгрузки');
  if (!sh) { sh = ss.insertSheet('Отгрузки'); sh.appendRow(['Дата', 'Месяц', 'Объект', 'Позиция', 'Кол-во', 'ID']); }
  return sh;
}

// ---------- Черновики накладных (хранятся в листе «Прачка», статус «черновик») ----------
function createDraft(obj) {
  var norm = NORMS[obj];
  if (!norm) return null;
  var sh = getLaundrySheet();
  var id = 'D' + Date.now().toString(36);
  var d = Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm');
  var i = 0;
  for (var name in norm) { sh.appendRow([id + '_' + i, d, obj, name, norm[name], 'черновик', '']); i++; }
  return id;
}

function draftRows(id) {
  var v = getLaundrySheet().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++)
    if (String(v[i][0]).indexOf(id + '_') === 0)
      out.push({ row: i + 1, id: String(v[i][0]), obj: v[i][2], name: v[i][3], qty: Number(v[i][4]), status: v[i][5] });
  return out;
}

function draftText(id, title) {
  var rows = draftRows(id);
  if (!rows.length) return null;
  var t = title + '\n🏠 <b>' + rows[0].obj + '</b>\n\n';
  var total = 0;
  for (var i = 0; i < rows.length; i++) { if (rows[i].qty > 0) { t += '• ' + rows[i].name + ' — ' + rows[i].qty + ' шт\n'; total += rows[i].qty; } }
  t += '\n📦 Итого: ' + total + ' шт';
  return t;
}

function kbDraft(id) {
  return [
    [{ text: '🧾 Открыть фирменную накладную', web_app: { url: VIEW_URL + '?id=' + id } }],
    [{ text: '✅ Всё верно — подготовлено', callback_data: 'dok|' + id }],
    [{ text: '✏️ Изменить количество', callback_data: 'ded|' + id }],
    [{ text: '❌ Уборки не было — отменить', callback_data: 'dno|' + id }]
  ];
}

function kbDraftEdit(id) {
  var rows = draftRows(id);
  var kb = [];
  for (var i = 0; i < rows.length; i++) {
    var suffix = rows[i].id.split('_')[1];
    kb.push([{ text: '➖ ' + rows[i].name + ' (' + rows[i].qty + ')', callback_data: 'dmn|' + id + '|' + suffix }]);
  }
  kb.push([{ text: '⬅️ Готово, вернуться', callback_data: 'dbk|' + id }]);
  return kb;
}

function editMsg(chat, mid, text, kb) {
  var p = { chat_id: chat, message_id: mid, text: text, parse_mode: 'HTML' };
  if (kb) p.reply_markup = { inline_keyboard: kb };
  tg('editMessageText', p);
}

// ---------- Ежедневный черновик в день выезда (триггер 10:00) ----------
function sendDraftNakl() {
  var today = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyyMMdd');
  var p = PropertiesService.getScriptProperties();
  var maids = (p.getProperty('MAID_CHATS') || '').match(/-?\d+/g) || [];
  var admins = adminIds();
  var targets = maids.concat(admins.filter(function (a) { return maids.indexOf(a) < 0; }));
  for (var c = 0; c < CALENDARS.length; c++) {
    if (!CALENDARS[c].url || !NORMS[CALENDARS[c].obj]) continue;
    var txt = '';
    try { txt = UrlFetchApp.fetch(CALENDARS[c].url).getContentText(); } catch (er) { continue; }
    var ev = parseIcs(txt);
    var vout = 0;
    for (var k = 0; k < ev.length; k++) if (ev[k].e === today) vout++;
    if (!vout) continue;
    var id = createDraft(CALENDARS[c].obj);
    if (!id) continue;
    var t = draftText(id, '📋 <b>Накладная — черновик (выезд сегодня)</b>') +
      '\n\nЭто максимальный комплект по норме. Если гостей было меньше — нажмите «Изменить» и уберите лишнее 👇';
    for (var m = 0; m < targets.length; m++) send(targets[m], t, kbDraft(id));
  }
}

// ---------- Обработка кнопок черновика и забора ----------
function handleDraftCb(cb) {
  var chat = cb.message.chat.id;
  var mid = cb.message.message_id;
  var d = cb.data.split('|');
  tg('answerCallbackQuery', { callback_query_id: cb.id });

  if (d[0] === 'dok') {
    var rows = draftRows(d[1]);
    var fresh = rows.filter(function (r) { return r.status === 'черновик'; });
    if (!fresh.length) { editMsg(chat, mid, (draftText(d[1], '📋 <b>Накладная</b>') || '') + '\n\n☑️ Уже обработана ранее.'); return; }
    var sh = getLaundrySheet();
    for (var i = 0; i < fresh.length; i++) sh.getRange(fresh[i].row, 6).setValue('подготовлено');
    editMsg(chat, mid, draftText(d[1], '✅ <b>Подготовлено к сдаче в прачку</b>') + '\n\nБельё копится до забора прачкой.', kbView(d[1]));
    notifyPrepared();

  } else if (d[0] === 'ded') {
    editMsg(chat, mid, draftText(d[1], '✏️ <b>Уберите лишнее</b> (каждое нажатие −1 шт)'), kbDraftEdit(d[1]));

  } else if (d[0] === 'dmn') {
    var rows2 = draftRows(d[1]);
    for (var j = 0; j < rows2.length; j++) {
      if (rows2[j].id === d[1] + '_' + d[2] && rows2[j].status === 'черновик' && rows2[j].qty > 0) {
        getLaundrySheet().getRange(rows2[j].row, 5).setValue(rows2[j].qty - 1);
        break;
      }
    }
    editMsg(chat, mid, draftText(d[1], '✏️ <b>Уберите лишнее</b> (каждое нажатие −1 шт)'), kbDraftEdit(d[1]));

  } else if (d[0] === 'dbk') {
    editMsg(chat, mid, draftText(d[1], '📋 <b>Накладная — черновик</b>') + '\n\nПроверьте и подтвердите 👇', kbDraft(d[1]));

  } else if (d[0] === 'dno') {
    var rows3 = draftRows(d[1]);
    var sh3 = getLaundrySheet();
    for (var n = 0; n < rows3.length; n++) if (rows3[n].status === 'черновик') sh3.getRange(rows3[n].row, 6).setValue('отменено');
    editMsg(chat, mid, '❌ Накладная отменена (' + (rows3.length ? rows3[0].obj : '') + ').');

  } else if (d[0] === 'pk') {
    doPickup(chat, mid);
  }
}

// Сколько подготовлено — сообщение прачке и хозяйке с кнопкой забора
function preparedRows() {
  var v = getLaundrySheet().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++)
    if (v[i][5] === 'подготовлено')
      out.push({ row: i + 1, obj: v[i][2], name: v[i][3], qty: Number(v[i][4]) });
  return out;
}

function preparedText() {
  var rows = preparedRows();
  if (!rows.length) return null;
  var byObj = {}, total = 0;
  for (var i = 0; i < rows.length; i++) {
    if (!byObj[rows[i].obj]) byObj[rows[i].obj] = [];
    byObj[rows[i].obj].push('• ' + rows[i].name + ' — ' + rows[i].qty + ' шт');
    total += rows[i].qty;
  }
  var t = '🧺 <b>Подготовлено к сдаче в прачку:</b>\n';
  for (var o in byObj) t += '\n🏠 <b>' + o + '</b>\n' + byObj[o].join('\n') + '\n';
  t += '\n📦 Всего: ' + total + ' шт';
  return t;
}

function notifyPrepared() {
  var t = preparedText();
  if (!t) return;
  var kb = [[{ text: '🚚 Подтвердить забор', callback_data: 'pk|0' }]];
  var L = PropertiesService.getScriptProperties().getProperty('LAUNDRY_CHAT');
  if (L) send(L, t + '\n\nКогда заберёте бельё — нажмите кнопку 👇', kb);
  var a = adminIds();
  for (var i = 0; i < a.length; i++) send(a[i], t, kb);
}

// Забор прачкой: подготовлено → в прачке + накладная-факт + запись в «Отгрузки»
function doPickup(chat, mid) {
  var rows = preparedRows();
  if (!rows.length) { editMsg(chat, mid, 'Сейчас нет подготовленного белья — забирать нечего 🙂'); return; }
  var sh = getLaundrySheet();
  var ship = getShipSheet();
  var j = getSheetFile().getSheetByName('Журнал');
  var now = new Date();
  var d = Utilities.formatDate(now, 'Europe/Moscow', 'dd.MM.yyyy');
  var mon = Utilities.formatDate(now, 'Europe/Moscow', 'yyyy-MM');
  var shipId = 'S' + Date.now().toString(36);
  var byObj = {}, total = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].qty > 0) {
      sh.getRange(rows[i].row, 6).setValue('в прачке');
      ship.appendRow([d, mon, rows[i].obj, rows[i].name, rows[i].qty, shipId]);
      j.appendRow([d, 'сдано в прачку', rows[i].obj, rows[i].name, rows[i].qty, 'забор прачкой']);
      if (!byObj[rows[i].obj]) byObj[rows[i].obj] = [];
      byObj[rows[i].obj].push('• ' + rows[i].name + ' — ' + rows[i].qty + ' шт');
      total += rows[i].qty;
    } else {
      sh.getRange(rows[i].row, 6).setValue('отменено');
    }
  }
  var t = '🚚 <b>НАКЛАДНАЯ — ФАКТ ОТГРУЗКИ</b>\n📅 ' + d + '\n';
  for (var o in byObj) t += '\n🏠 <b>' + o + '</b>\n' + byObj[o].join('\n') + '\n';
  t += '\n📦 Итого отгружено: ' + total + ' шт';
  editMsg(chat, mid, t, kbView(shipId));
  var L = PropertiesService.getScriptProperties().getProperty('LAUNDRY_CHAT');
  if (L && String(L) !== String(chat)) send(L, t, kbView(shipId));
  var a = adminIds();
  for (var k = 0; k < a.length; k++) if (String(a[k]) !== String(chat)) send(a[k], t, kbView(shipId));
}

// ---------- Месячная сверка (триггер: 1-е число, 10:00) ----------
var MONTH_RU = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];

function monthlyText() {
  var d = new Date();
  d.setDate(0); // последний день прошлого месяца
  var mon = Utilities.formatDate(d, 'Europe/Moscow', 'yyyy-MM');
  var name = MONTH_RU[d.getMonth()] + ' ' + d.getFullYear();
  var v = getShipSheet().getDataRange().getValues();
  var byObj = {}, grand = 0, days = {};
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][1]) !== mon) continue;
    var o = v[i][2], nm = v[i][3], q = Number(v[i][4]);
    if (!byObj[o]) byObj[o] = {};
    byObj[o][nm] = (byObj[o][nm] || 0) + q;
    grand += q;
    days[String(v[i][0])] = true;
  }
  var nShip = Object.keys(days).length;
  if (!grand) return '📊 <b>Сверка за ' + name + '</b>\n\nОтгрузок в прачку не было.';
  var t = '📊 <b>СВЕРКА С ПРАЧКОЙ — ' + name + '</b>\n🚚 Дней с отгрузками: ' + nShip + '\n';
  for (var ob in byObj) {
    var sub = 0, lines = '';
    for (var it in byObj[ob]) { lines += '• ' + it + ' — ' + byObj[ob][it] + ' шт\n'; sub += byObj[ob][it]; }
    t += '\n🏠 <b>' + ob + '</b>\n' + lines + 'Итого: ' + sub + ' шт\n';
  }
  t += '\n📦 <b>ВСЕГО ЗА МЕСЯЦ: ' + grand + ' шт</b>';
  t += '\n\n📄 Подробности по датам — лист «Отгрузки»:\n' + getSheetFile().getUrl();
  return t;
}

function sendMonthlyReport() {
  var t = monthlyText();
  var a = adminIds();
  for (var i = 0; i < a.length; i++) send(a[i], t);
}

// ---------- Меню с учётом ролей (прачка теперь тоже с кнопкой) ----------
function kbFor(id) {
  var r = roleOf(id);
  if (r === 'guest') return null;
  if (r === 'laundry') return [[{ text: '🚚 Подтвердить забор', callback_data: 'pk|0' }]];
  var rows = [];
  if (r === 'admin') rows.push([{ text: '📅 План на сегодня', callback_data: 'pl|0' }]);
  rows.push([{ text: '🧺 Сдать в прачку', web_app: { url: FORM_URL } }]);
  rows.push([{ text: '📦 Приняли из прачки', web_app: { url: PRIEM_URL } }]);
  rows.push([{ text: '📊 Остатки', web_app: { url: OSTATKI_URL } }]);
  if (r === 'admin') rows.push([{ text: '🚚 Прачка забрала бельё', callback_data: 'pk|0' }, { text: '📊 Сверка за месяц', callback_data: 'rep|0' }]);
  return rows;
}

// ---------- Итоговый doPost (заменяет все версии выше — так задумано) ----------
function doPost(e) {
  try {
    var u = JSON.parse(e.postData.contents);

    // Mini App: накладная из формы
    if (u.oleand === 'nakl' && u.chat && u.text) {
      if (u.obj && u.items) saveNakl(u.obj, u.items, u.who);
      tg('sendMessage', { chat_id: u.chat, text: u.text });
      var L = PropertiesService.getScriptProperties().getProperty('LAUNDRY_CHAT');
      if (L && String(L) !== String(u.chat)) tg('sendMessage', { chat_id: L, text: u.text });
      return HtmlService.createHtmlOutput('ok');
    }
    // Mini App: приём из прачки
    if (u.oleand === 'recv' && u.chat && u.ids) { receiveIds(u.ids, u.chat); return HtmlService.createHtmlOutput('ok'); }

    if (u.callback_query) {
      var cd = String(u.callback_query.data || '');
      // Кнопки черновиков и забора (ЭТАП 8)
      if (/^(dok|ded|dmn|dpl|dzz|dbk|dno|pk)\|/.test(cd)) { handleDraftCb(u.callback_query); return HtmlService.createHtmlOutput('ok'); }
      // Сверка по запросу
      if (cd === 'rep|0') {
        tg('answerCallbackQuery', { callback_query_id: u.callback_query.id });
        var rc = u.callback_query.message.chat.id;
        if (roleOf(rc) === 'admin') send(rc, monthlyText(), kbFor(rc));
        return HtmlService.createHtmlOutput('ok');
      }
      // План на сегодня
      if (cd === 'pl|0') {
        tg('answerCallbackQuery', { callback_query_id: u.callback_query.id });
        var cid = u.callback_query.message.chat.id;
        if (roleOf(cid) === 'admin') { addPlanSub(cid); send(cid, planText(), kbFor(cid)); }
        else { send(cid, 'План на день доступен только хозяйке и менеджеру 🙂', kbFor(cid)); }
        return HtmlService.createHtmlOutput('ok');
      }
    }

    if (u.message) {
      var id = u.message.chat.id;
      var t = String(u.message.text || '').toLowerCase().trim();
      var p = PropertiesService.getScriptProperties();
      if (t === 'я хозяйка' || t === 'я менеджер') { addAdmin(String(id)); send(id, 'Готово! 👑 Вам доступно всё, включая план на день.', kbFor(id)); return HtmlService.createHtmlOutput('ok'); }
      if (t === 'я горничная') { var m = p.getProperty('MAID_CHATS') || ''; if (m.indexOf('[' + id + ']') < 0) p.setProperty('MAID_CHATS', m + '[' + id + ']'); send(id, 'Готово! Вы — горничная 🧹 В день выезда гостей сюда будет приходить накладная — проверьте и подтвердите её.', kbFor(id)); return HtmlService.createHtmlOutput('ok'); }
      if (t === 'я прачка') { p.setProperty('LAUNDRY_CHAT', String(id)); send(id, 'Готово! Вы — прачка 🧺 Когда бельё подготовлено, сюда придёт список с кнопкой «Подтвердить забор».', kbFor(id)); return HtmlService.createHtmlOutput('ok'); }
      var r = roleOf(id);
      if (r === 'laundry') { var pt2 = preparedText(); send(id, pt2 ? pt2 + '\n\nКогда заберёте бельё — нажмите кнопку 👇' : 'Пока ничего не подготовлено. Списки будут приходить сюда автоматически 🧺', kbFor(id)); }
      else if (r === 'guest') { send(id, 'Привет! Это бот OLEAND. Напишите кодовое слово, которое вам дали.'); }
      else { send(id, 'Меню 👇', kbFor(id)); }
      return HtmlService.createHtmlOutput('ok');
    }

    if (u.callback_query) handleCallback(u.callback_query);
  } catch (err) {}
  return HtmlService.createHtmlOutput('ok');
}

// ---------- Запустить ОДИН раз после развёртывания новой версии ----------
function setupStage8() {
  getShipSheet();
  var tr = ScriptApp.getProjectTriggers();
  for (var i = 0; i < tr.length; i++) {
    var f = tr[i].getHandlerFunction();
    if (f === 'sendDraftNakl' || f === 'sendMonthlyReport') ScriptApp.deleteTrigger(tr[i]);
  }
  ScriptApp.newTrigger('sendDraftNakl').timeBased().atHour(10).everyDays(1).create();
  ScriptApp.newTrigger('sendMonthlyReport').timeBased().onMonthDay(1).atHour(10).create();
  Logger.log('OK: черновики в 10:00 ежедневно, сверка 1-го числа в 10:00');
}

// Проверки без отправки в Telegram
function testDraftToday() { sendDraftNakl(); Logger.log('Черновики разосланы (если сегодня есть выезды)'); }
function testMonthly() { Logger.log(monthlyText()); }

// ========== ЭТАП 9.1: УДОБНОЕ ИЗМЕНЕНИЕ КОЛИЧЕСТВА (➖ и ➕) ==========
function kbDraftEdit(id) {
  var rows = draftRows(id);
  var kb = [];
  for (var i = 0; i < rows.length; i++) {
    var sfx = rows[i].id.split('_')[1];
    kb.push([
      { text: '➖', callback_data: 'dmn|' + id + '|' + sfx },
      { text: rows[i].name + ': ' + rows[i].qty, callback_data: 'dzz|0' },
      { text: '➕', callback_data: 'dpl|' + id + '|' + sfx }
    ]);
  }
  kb.push([{ text: '✅ Готово', callback_data: 'dbk|' + id }]);
  return kb;
}

function editText(id) {
  var rows = draftRows(id);
  var t = '✏️ <b>Изменить количество</b>\nНажимайте ➖ или ➕ рядом с позицией. Когда всё верно — «✅ Готово».';
  if (!rows.length) return t;
  var total = 0;
  for (var i = 0; i < rows.length; i++) total += rows[i].qty;
  return t + '\n\n🏠 <b>' + rows[0].obj + '</b>\n📦 Итого: ' + total + ' шт';
}

function handleDraftCb(cb) {
  var chat = cb.message.chat.id;
  var mid = cb.message.message_id;
  var d = cb.data.split('|');
  tg('answerCallbackQuery', { callback_query_id: cb.id });

  if (d[0] === 'dok') {
    var rows = draftRows(d[1]);
    var fresh = rows.filter(function (r) { return r.status === 'черновик'; });
    if (!fresh.length) { editMsg(chat, mid, (draftText(d[1], '📋 <b>Накладная</b>') || '') + '\n\n☑️ Уже обработана ранее.'); return; }
    var sh = getLaundrySheet();
    for (var i = 0; i < fresh.length; i++) sh.getRange(fresh[i].row, 6).setValue('подготовлено');
    editMsg(chat, mid, draftText(d[1], '✅ <b>Подготовлено к сдаче в прачку</b>') + '\n\nБельё копится до забора прачкой.', kbView(d[1]));
    notifyPrepared();

  } else if (d[0] === 'ded') {
    editMsg(chat, mid, editText(d[1]), kbDraftEdit(d[1]));

  } else if (d[0] === 'dmn' || d[0] === 'dpl') {
    var rows2 = draftRows(d[1]);
    var changed = false;
    for (var j = 0; j < rows2.length; j++) {
      if (rows2[j].id === d[1] + '_' + d[2] && rows2[j].status === 'черновик') {
        var q = rows2[j].qty + (d[0] === 'dpl' ? 1 : -1);
        if (q < 0) q = 0;
        if (q > 30) q = 30;
        if (q !== rows2[j].qty) { getLaundrySheet().getRange(rows2[j].row, 5).setValue(q); changed = true; }
        break;
      }
    }
    if (changed) editMsg(chat, mid, editText(d[1]), kbDraftEdit(d[1]));

  } else if (d[0] === 'dbk') {
    editMsg(chat, mid, draftText(d[1], '📋 <b>Накладная — черновик</b>') + '\n\nПроверьте и подтвердите 👇', kbDraft(d[1]));

  } else if (d[0] === 'dno') {
    var rows3 = draftRows(d[1]);
    var sh3 = getLaundrySheet();
    for (var n = 0; n < rows3.length; n++) if (rows3[n].status === 'черновик') sh3.getRange(rows3[n].row, 6).setValue('отменено');
    editMsg(chat, mid, '❌ Накладная отменена (' + (rows3.length ? rows3[0].obj : '') + ').');

  } else if (d[0] === 'pk') {
    doPickup(chat, mid);
  }
}
