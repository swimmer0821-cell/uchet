/**
 * ============================================================
 *  Учёт доходов и расходов SUB PROFI — серверная часть (версия 5)
 *  (Google Apps Script над Google Таблицей)
 * ============================================================
 *
 *  Что делает этот код:
 *   1. Функция setup() один раз создаёт в таблице листы
 *      «Операции», «Справочники», «Отчёты» и заполняет справочник статей.
 *   2. Функция doPost() принимает запросы от приложения (index.html):
 *      записать, изменить, отменить операцию, показать историю, посчитать отчёт.
 *
 *  Все расчёты отчётов делаются здесь, в приложение приходит готовый результат.
 */

// ===================== НАСТРОЙКИ =====================
const SHEET_OPS = 'Операции';
const SHEET_REF = 'Справочники';
const SHEET_REP = 'Отчёты';

const CANCEL_MINUTES = 10;          // сколько минут сотрудник может отменить свою запись (изменять — в любое время)
const HISTORY_LIMIT = 50;           // сколько записей показывать в истории
const MAX_AMOUNT = 10000000000;     // 10 млрд сум: больше не принимаем (защита от лишних нулей)
const TYPICAL_DAYS = 120;           // за сколько дней смотреть «обычную сумму» статьи
const TIMEZONE = 'Asia/Tashkent';

const TYPES = ['Расход', 'Доход', 'Вложение', 'Перевод'];
const STATUS_CANCELLED = 'Отменена';
const ROLE_BOSS = 'Руководитель';

// Колонки листа «Операции»
const OPS_HEADERS = ['ID', 'Дата', 'Тип', 'Статья', 'Подстатья', 'Сумма', 'Способ оплаты',
  'Контрагент', 'Комментарий', 'Кто внёс', 'Время записи', 'Статус', 'Изменено',
  'Наименование', 'Количество', 'Ед.'];

// Единицы измерения для количества
const UNITS = [['кг', 'kg'], ['г', 'g'], ['л', 'l'], ['шт', 'dona'], ['пачка', 'pachka'], ['рулон', 'rulon'], ['м', 'm']];

// Колонки листа «Справочники»: статьи A:H, способы оплаты J:K, контрагенты M, пользователи O:Q
const REF_HEADERS = ['Тип', 'Статья', 'Статья (узб)', 'Подстатья', 'Подстатья (узб)',
  'Быстрая плитка №', 'Надпись плитки', 'Надпись плитки (узб)'];

// ===================== СПРАВОЧНИК ПО УМОЛЧАНИЮ =====================
// [Тип, Статья, Статья узб, Подстатья, Подстатья узб, № быстрой плитки, Надпись, Надпись узб]
// Быстрые плитки — частые траты, которые сотрудник видит сразу на первом экране.
const DEFAULT_ARTICLES = [
  ['Расход', 'Сырьё', 'Xom ashyo', 'Фрукты и ягоды', 'Mevalar va rezavorlar', 8, 'Фрукты и ягоды', 'Mevalar'],
  ['Расход', 'Сырьё', 'Xom ashyo', 'Молочное', 'Sut mahsulotlari', '', '', ''],
  ['Расход', 'Сырьё', 'Xom ashyo', 'Шоколад, орехи, добавки', 'Shokolad, yongʻoq, qoʻshimchalar', '', '', ''],
  ['Расход', 'Сырьё', 'Xom ashyo', 'Курут', 'Qurut', '', '', ''],
  ['Расход', 'Сырьё', 'Xom ashyo', 'Сулугуни', 'Suluguni', '', '', ''],
  ['Расход', 'Сырьё', 'Xom ashyo', 'Услуга сублимации', 'Sublimatsiya xizmati', '', '', ''],
  ['Расход', 'Сырьё', 'Xom ashyo', 'Образцы', 'Namunalar', 7, 'Образцы', 'Namunalar'],

  ['Расход', 'Упаковка', 'Qadoqlash', 'Банки и контейнеры', 'Bankalar va idishlar', '', '', ''],
  ['Расход', 'Упаковка', 'Qadoqlash', 'Этикетки и стикеры', 'Etiketka va stikerlar', '', '', ''],
  ['Расход', 'Упаковка', 'Qadoqlash', 'Пакеты и плёнка', 'Paketlar va plyonka', '', '', ''],
  ['Расход', 'Упаковка', 'Qadoqlash', 'Коробки', 'Qutilar', '', '', ''],
  ['Расход', 'Упаковка', 'Qadoqlash', 'Ложки и мелочи', 'Qoshiqlar va mayda narsalar', '', '', ''],

  ['Расход', 'Зарплата', 'Ish haqi', 'Аванс на карты', 'Kartaga avans', '', '', ''],
  ['Расход', 'Зарплата', 'Ish haqi', 'Наличная часть', 'Naqd qism', '', '', ''],
  ['Расход', 'Зарплата', 'Ish haqi', 'Еженедельные выплаты', 'Haftalik toʻlovlar', '', '', ''],
  ['Расход', 'Зарплата', 'Ish haqi', 'Наёмные рабочие', 'Yollanma ishchilar', 9, 'ЗП наёмным', 'Yollanma ishchilarga'],
  ['Расход', 'Зарплата', 'Ish haqi', 'Премии и переработки', 'Mukofot va qoʻshimcha ish', '', '', ''],

  ['Расход', 'Аренда', 'Ijara', 'Помещение и офис', 'Bino va ofis', '', '', ''],
  ['Расход', 'Аренда', 'Ijara', 'Прочая аренда', 'Boshqa ijara', '', '', ''],

  ['Расход', 'Коммунальные', 'Kommunal', 'Электричество', 'Elektr energiya', '', '', ''],
  ['Расход', 'Коммунальные', 'Kommunal', 'Газ', 'Gaz', '', '', ''],
  ['Расход', 'Коммунальные', 'Kommunal', 'Вода', 'Suv', '', '', ''],
  ['Расход', 'Коммунальные', 'Kommunal', 'Вывоз мусора', 'Chiqindi olib ketish', 6, 'Вывоз мусора', 'Chiqindi'],
  ['Расход', 'Коммунальные', 'Kommunal', 'Связь и интернет', 'Aloqa va internet', '', '', ''],

  ['Расход', 'Транспорт и логистика', 'Transport va logistika', 'Доставка', 'Yetkazib berish', 2, 'Доставка', 'Yetkazib berish'],
  ['Расход', 'Транспорт и логистика', 'Transport va logistika', 'Такси', 'Taksi', 3, 'Такси', 'Taksi'],
  ['Расход', 'Транспорт и логистика', 'Transport va logistika', 'Бензин', 'Benzin', 5, 'Бензин', 'Benzin'],
  ['Расход', 'Транспорт и логистика', 'Transport va logistika', 'Стоянка', 'Avtoturargoh', 4, 'Стоянка', 'Turargoh'],
  ['Расход', 'Транспорт и логистика', 'Transport va logistika', 'Сухой лёд и термотара', 'Quruq muz va termoidish', '', '', ''],
  ['Расход', 'Транспорт и логистика', 'Transport va logistika', 'Машины: кредит, страховка, ремонт', 'Mashinalar: kredit, sugʻurta, taʼmir', '', '', ''],

  ['Расход', 'Оборудование и ремонт', 'Uskuna va taʼmir', 'Покупка оборудования', 'Uskuna sotib olish', '', '', ''],
  ['Расход', 'Оборудование и ремонт', 'Uskuna va taʼmir', 'Морозильная камера', 'Muzlatish kamerasi', '', '', ''],
  ['Расход', 'Оборудование и ремонт', 'Uskuna va taʼmir', 'Ремонт и запчасти', 'Taʼmir va ehtiyot qismlar', '', '', ''],
  ['Расход', 'Оборудование и ремонт', 'Uskuna va taʼmir', 'Расходники производства', 'Ishlab chiqarish sarf materiallari', '', '', ''],

  ['Расход', 'Налоги и сборы', 'Soliq va yigʻimlar', 'Налоги с зарплаты', 'Ish haqi soliqlari', '', '', ''],
  ['Расход', 'Налоги и сборы', 'Soliq va yigʻimlar', 'Налог на аренду', 'Ijara soligʻi', '', '', ''],
  ['Расход', 'Налоги и сборы', 'Soliq va yigʻimlar', 'Прочие налоги', 'Boshqa soliqlar', '', '', ''],
  ['Расход', 'Налоги и сборы', 'Soliq va yigʻimlar', 'Госуслуги и проверки', 'Davlat xizmatlari va tekshiruvlar', '', '', ''],
  ['Расход', 'Налоги и сборы', 'Soliq va yigʻimlar', 'Комиссии банка', 'Bank komissiyasi', '', '', ''],

  ['Расход', 'Реклама и продвижение', 'Reklama', 'СММ и таргет', 'SMM va target', '', '', ''],
  ['Расход', 'Реклама и продвижение', 'Reklama', 'Блогеры', 'Blogerlar', '', '', ''],
  ['Расход', 'Реклама и продвижение', 'Reklama', 'Фото, видео, дизайн', 'Foto, video, dizayn', '', '', ''],
  ['Расход', 'Реклама и продвижение', 'Reklama', 'Печать', 'Bosma mahsulot', '', '', ''],
  ['Расход', 'Реклама и продвижение', 'Reklama', 'Выставки и мерч', 'Koʻrgazma va merch', '', '', ''],

  ['Расход', 'Прочее', 'Boshqa', 'Питание персонала', 'Xodimlar ovqati', 1, 'Хлеб и питание', 'Non va ovqat'],
  ['Расход', 'Прочее', 'Boshqa', 'Хозтовары и уборка', 'Xoʻjalik mollari', 10, 'Хозтовары', 'Xoʻjalik mollari'],
  ['Расход', 'Прочее', 'Boshqa', 'Программы и сервисы', 'Dasturlar va xizmatlar', '', '', ''],
  ['Расход', 'Прочее', 'Boshqa', 'Бухгалтерия', 'Buxgalteriya', '', '', ''],
  ['Расход', 'Прочее', 'Boshqa', 'Поиск персонала', 'Xodim qidirish', '', '', ''],
  ['Расход', 'Прочее', 'Boshqa', 'Другое', 'Boshqa', '', '', ''],

  ['Доход', 'Продажи оптом', 'Ulgurji savdo', 'Дистрибьютор', 'Distribyutor', '', '', ''],
  ['Доход', 'Продажи оптом', 'Ulgurji savdo', 'Другие оптовые покупатели', 'Boshqa ulgurji xaridorlar', '', '', ''],
  ['Доход', 'Прочие поступления', 'Boshqa tushumlar', 'Продажа имущества', 'Mol-mulk sotish', '', '', ''],
  ['Доход', 'Прочие поступления', 'Boshqa tushumlar', 'Возмещения', 'Qoplamalar', '', '', ''],
  ['Доход', 'Прочие поступления', 'Boshqa tushumlar', 'Другое', 'Boshqa', '', '', ''],

  ['Вложение', 'Вложения и займы', 'Investitsiya va qarzlar', 'Деньги от АЛ', 'AL dan pul', '', '', ''],
  ['Вложение', 'Вложения и займы', 'Investitsiya va qarzlar', 'Займ у Бест биотик', 'Best biotikdan qarz', '', '', ''],
  ['Вложение', 'Вложения и займы', 'Investitsiya va qarzlar', 'Другие займы', 'Boshqa qarzlar', '', '', ''],

  ['Перевод', 'Перевод между своими', 'Oʻz hisoblar oʻrtasida', 'Инкассация (сдали в банк)', 'Inkassatsiya (bankka topshirildi)', '', '', ''],
  ['Перевод', 'Перевод между своими', 'Oʻz hisoblar oʻrtasida', 'Сняли с карты или счёта', 'Karta yoki hisobdan yechildi', '', '', ''],
  ['Перевод', 'Перевод между своими', 'Oʻz hisoblar oʻrtasida', 'Пополнили счёт', 'Hisob toʻldirildi', '', '', ''],
  ['Перевод', 'Перевод между своими', 'Oʻz hisoblar oʻrtasida', 'Уставной фонд', 'Ustav fondi', '', '', '']
];

const DEFAULT_PAYS = [['Наличные', 'Naqd'], ['Карта', 'Karta'], ['Перечисление', 'Bank orqali']];

const DEFAULT_PARTIES = ['Дистрибьютор (Баходыр)', 'АЛ', 'Бест биотик', 'Бест пласт', 'Мирворис',
  'Хикмат', 'Секта', 'Илья (СММ)', 'Чагай плюс', 'Фуд пак'];

// [Имя, Роль, PIN]. PIN можно оставить пустым — тогда вход по нажатию на имя.
const DEFAULT_USERS = [['Руководитель', ROLE_BOSS, ''], ['Шахноза', 'Сотрудник', ''], ['Елена', 'Сотрудник', '']];


// ===================== МЕНЮ В ТАБЛИЦЕ =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Учёт')
    .addItem('Настроить таблицу', 'setup')
    .addToUi();
}


// ===================== НАСТРОЙКА ТАБЛИЦЫ (запустить один раз) =====================
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TIMEZONE);
  ss.setSpreadsheetLocale('ru_RU');

  // --- Лист «Операции» ---
  const ops = ss.getSheetByName(SHEET_OPS) || ss.insertSheet(SHEET_OPS);
  if (ops.getLastRow() === 0) {
    ops.getRange(1, 1, 1, OPS_HEADERS.length).setValues([OPS_HEADERS]).setFontWeight('bold').setBackground('#e8eef3');
    ops.setFrozenRows(1);
  }
  ops.getRange('B2:B').setNumberFormat('dd.mm.yyyy');
  ops.getRange('F2:F').setNumberFormat('#,##0');
  ops.getRange('K2:K').setNumberFormat('dd.mm.yyyy hh:mm');
  ops.setColumnWidth(1, 90);
  [4, 5, 9].forEach(c => ops.setColumnWidth(c, 200));
  ensureCancelFormat_(ops);
  ops.getRange('C2:C').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).setAllowInvalid(false).build());

  // --- Лист «Справочники» (заполняется, только если пустой) ---
  const ref = ss.getSheetByName(SHEET_REF) || ss.insertSheet(SHEET_REF);
  if (ref.getLastRow() === 0) {
    ref.getRange(1, 1, 1, REF_HEADERS.length).setValues([REF_HEADERS]);
    ref.getRange(2, 1, DEFAULT_ARTICLES.length, REF_HEADERS.length).setValues(DEFAULT_ARTICLES);
    ref.getRange(1, 10, 1, 2).setValues([['Способ оплаты', 'Способ оплаты (узб)']]);
    ref.getRange(2, 10, DEFAULT_PAYS.length, 2).setValues(DEFAULT_PAYS);
    ref.getRange(1, 13).setValue('Контрагенты');
    ref.getRange(2, 13, DEFAULT_PARTIES.length, 1).setValues(DEFAULT_PARTIES.map(p => [p]));
    ref.getRange(1, 15, 1, 3).setValues([['Пользователь', 'Роль', 'PIN (необязательно)']]);
    ref.getRange(2, 15, DEFAULT_USERS.length, 3).setValues(DEFAULT_USERS);
    ref.getRange('Q2:Q').setNumberFormat('@'); // PIN хранится как текст, чтобы не терялись нули
    ref.getRange('1:1').setFontWeight('bold').setBackground('#e8eef3');
    ref.setFrozenRows(1);
    [2, 4, 5].forEach(c => ref.setColumnWidth(c, 210));
    ref.getRange('P2:P').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList([ROLE_BOSS, 'Сотрудник'], true).build());
  }
  ops.getRange('G2:G').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(ref.getRange('J2:J'), true).setAllowInvalid(false).build());

  // --- Лист «Отчёты» (перезаписывается целиком, только формулы) ---
  const rep = ss.getSheetByName(SHEET_REP) || ss.insertSheet(SHEET_REP);
  rep.clear();
  buildReportSheet_(rep);

  // Убираем пустой стартовый лист, если он есть
  ['Лист1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  ss.toast('Листы «Операции», «Справочники», «Отчёты» готовы.', 'Готово', 10);
  Logger.log('Готово. Таблица настроена.');
}

// Формулы листа «Отчёты». Менять руками можно только даты в C3 и C4.
// Записывает формулу. В таблицах с русскими настройками разделитель — «;», а не «,»:
// если формула с запятыми дала ошибку, пишем её же с точками с запятой.
function setF_(sh, a1, formula) {
  const r = sh.getRange(a1);
  r.setFormula(formula);
  SpreadsheetApp.flush();
  if (/^#(ERROR|ОШИБКА)/i.test(String(r.getDisplayValue()))) r.setFormula(formula.replace(/,/g, ';'));
  return r;
}

function buildReportSheet_(rep) {
  const base = (type) =>
    `SUMIFS(${SHEET_OPS}!$F:$F,${SHEET_OPS}!$C:$C,"${type}",${SHEET_OPS}!$B:$B,">="&$C$3,${SHEET_OPS}!$B:$B,"<="&$C$4,${SHEET_OPS}!$L:$L,"<>${STATUS_CANCELLED}")`;

  rep.getRange('A1').setValue('Отчёт считается сам. Руками можно менять только даты в C3 и C4.').setFontWeight('bold');
  rep.getRange('B3:C4').setValues([['С', ''], ['По', '']]);
  setF_(rep, 'C3', '=DATE(YEAR(TODAY()),MONTH(TODAY()),1)');
  setF_(rep, 'C4', '=TODAY()');
  rep.getRange('C3:C4').setNumberFormat('dd.mm.yyyy').setBackground('#fff8d6');

  rep.getRange('B6:B10').setValues([['Доходы'], ['Расходы'], ['Разница'], ['Вложения и займы (не доход)'], ['Переводы между своими (не расход)']]);
  setF_(rep, 'C6', '=' + base('Доход'));
  setF_(rep, 'C7', '=' + base('Расход'));
  setF_(rep, 'C8', '=C6-C7');
  setF_(rep, 'C9', '=' + base('Вложение'));
  setF_(rep, 'C10', '=' + base('Перевод'));
  rep.getRange('C6:C10').setNumberFormat('#,##0').setFontWeight('bold');

  const byCat = (type, catCol, totalCell) =>
    `=ARRAYFORMULA(IF(${catCol}14:${catCol}60="","",SUMIFS(${SHEET_OPS}!$F:$F,${SHEET_OPS}!$C:$C,"${type}",${SHEET_OPS}!$D:$D,${catCol}14:${catCol}60,${SHEET_OPS}!$B:$B,">="&$C$3,${SHEET_OPS}!$B:$B,"<="&$C$4,${SHEET_OPS}!$L:$L,"<>${STATUS_CANCELLED}")))`;

  rep.getRange('A12').setValue('Расходы по статьям').setFontWeight('bold');
  rep.getRange('A13:C13').setValues([['Статья', 'Сумма', 'Доля']]).setFontWeight('bold');
  setF_(rep, 'A14', `=UNIQUE(FILTER(${SHEET_REF}!$B$2:$B,${SHEET_REF}!$A$2:$A="Расход"))`);
  setF_(rep, 'B14', byCat('Расход', 'A'));
  setF_(rep, 'C14', '=ARRAYFORMULA(IF(A14:A60="","",IFERROR(B14:B60/$C$7,0)))');

  rep.getRange('E12').setValue('Доходы по статьям').setFontWeight('bold');
  rep.getRange('E13:G13').setValues([['Статья', 'Сумма', 'Доля']]).setFontWeight('bold');
  setF_(rep, 'E14', `=UNIQUE(FILTER(${SHEET_REF}!$B$2:$B,${SHEET_REF}!$A$2:$A="Доход"))`);
  setF_(rep, 'F14', byCat('Доход', 'E'));
  setF_(rep, 'G14', '=ARRAYFORMULA(IF(E14:E60="","",IFERROR(F14:F60/$C$6,0)))');

  rep.getRange('I12').setValue('По способу оплаты').setFontWeight('bold');
  rep.getRange('I13:K13').setValues([['Способ', 'Расход', 'Доход']]).setFontWeight('bold');
  setF_(rep, 'I14', `=FILTER(${SHEET_REF}!$J$2:$J,${SHEET_REF}!$J$2:$J<>"")`);
  const byPay = (type) =>
    `=ARRAYFORMULA(IF(I14:I30="","",SUMIFS(${SHEET_OPS}!$F:$F,${SHEET_OPS}!$C:$C,"${type}",${SHEET_OPS}!$G:$G,I14:I30,${SHEET_OPS}!$B:$B,">="&$C$3,${SHEET_OPS}!$B:$B,"<="&$C$4,${SHEET_OPS}!$L:$L,"<>${STATUS_CANCELLED}")))`;
  setF_(rep, 'J14', byPay('Расход'));
  setF_(rep, 'K14', byPay('Доход'));

  rep.getRange('B14:B60').setNumberFormat('#,##0');
  rep.getRange('F14:F60').setNumberFormat('#,##0');
  rep.getRange('J14:K30').setNumberFormat('#,##0');
  rep.getRange('C14:C60').setNumberFormat('0%');
  rep.getRange('G14:G60').setNumberFormat('0%');
  rep.setColumnWidth(1, 220); rep.setColumnWidth(2, 260); rep.setColumnWidth(5, 220); rep.setColumnWidth(9, 150);
  rep.protect().setWarningOnly(true).setDescription('Отчёт считается формулами');
}


// ===================== ПРИЁМ ЗАПРОСОВ ОТ ПРИЛОЖЕНИЯ =====================
function doGet() {
  return out_({ ok: true, message: 'Учёт SUB PROFI работает' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ ok: false, error: 'ERR_DATA' }); }

  try {
    // Список пользователей нужен до входа, поэтому без проверки
    if (req.action === 'users') return out_({ ok: true, data: listUsers_() });
    // Первый вход: человек без PIN сам придумывает PIN
    if (req.action === 'setPin') return out_({ ok: true, data: setPin_(req.user, req.newPin) });

    const user = auth_(req.user, req.pin);
    let data;
    switch (req.action) {
      case 'init':    data = init_(user); break;
      case 'add':     data = addOp_(user, req.op); break;
      case 'history': data = history_(user, req.filter || {}); break;
      case 'cancel':  data = cancel_(user, req.id); break;
      case 'edit':    data = editOp_(user, req.id, req.op); break;
      case 'report':  data = report_(user, req.from, req.to); break;
      default: throw new Error('ERR_DATA');
    }
    return out_({ ok: true, data: data });
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


// ===================== ПОЛЬЗОВАТЕЛИ =====================
function listUsers_() {
  return readRef_().users.map(u => ({ name: u.name, role: u.role, hasPin: !!u.pin }));
}

// Проверяем, что пользователь есть в справочнике и PIN верный.
// Без PIN войти нельзя: сначала его нужно придумать (setPin_).
function auth_(name, pin) {
  const u = readRef_().users.find(x => x.name === String(name || '').trim());
  if (!u) throw new Error('ERR_USER');
  if (!u.pin) throw new Error('ERR_NO_PIN');
  if (String(pin || '') !== u.pin) throw new Error('ERR_PIN');
  return u;
}

// Сохранить новый PIN. Можно только если PIN ещё не задан
// (или руководитель очистил ячейку PIN в листе «Справочники»).
function setPin_(name, newPin) {
  const pin = String(newPin || '');
  if (!/^\d{4}$/.test(pin)) throw new Error('ERR_PIN_FORMAT');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SHEET_REF);
    const v = sh.getDataRange().getValues();
    for (let r = 1; r < v.length; r++) {
      if (String(v[r][14] || '').trim() === String(name || '').trim() && String(name || '').trim()) {
        if (String(v[r][16] || '').trim()) throw new Error('ERR_PIN_EXISTS');
        sh.getRange(r + 1, 17).setNumberFormat('@').setValue(pin);
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
    throw new Error('ERR_USER');
  } finally {
    lock.releaseLock();
  }
}


// ===================== ЧТЕНИЕ СПРАВОЧНИКОВ =====================
function readRef_() {
  const v = sheet_(SHEET_REF).getDataRange().getValues();
  const s = (x) => String(x === null || x === undefined ? '' : x).trim();
  const articles = [], pays = [], parties = [], users = [];
  for (let r = 1; r < v.length; r++) {
    const row = v[r];
    if (s(row[0]) && s(row[1]) && s(row[3])) {
      articles.push({
        type: s(row[0]), cat: s(row[1]), catUz: s(row[2]) || s(row[1]),
        sub: s(row[3]), subUz: s(row[4]) || s(row[3]),
        quick: Number(row[5]) || 0, tile: s(row[6]) || s(row[3]), tileUz: s(row[7]) || s(row[4]) || s(row[6]) || s(row[3])
      });
    }
    if (s(row[9])) pays.push({ ru: s(row[9]), uz: s(row[10]) || s(row[9]) });
    if (s(row[12])) parties.push(s(row[12]));
    if (s(row[14])) users.push({ name: s(row[14]), role: s(row[15]) === ROLE_BOSS ? 'boss' : 'staff', pin: s(row[16]) });
  }
  return { articles, pays, parties, users };
}


// ===================== ДАННЫЕ ДЛЯ ЗАПУСКА ПРИЛОЖЕНИЯ =====================
function init_(user) {
  const ref = readRef_();
  const allowed = user.role === 'boss' ? TYPES : ['Расход'];
  return {
    user: { name: user.name, role: user.role },
    articles: ref.articles.filter(a => allowed.indexOf(a.type) >= 0),
    pays: ref.pays,
    parties: ref.parties,
    typical: typical_(),
    names: recentNames_(),
    units: UNITS.map(u => ({ ru: u[0], uz: u[1] })),
    today: today_(),
    cancelMinutes: CANCEL_MINUTES
  };
}

// «Обычная сумма» по каждой подстатье — медиана за последние TYPICAL_DAYS дней.
// Нужна, чтобы предупредить, если сумма в 5 раз больше обычной.
function typical_() {
  const since = addDays_(today_(), -TYPICAL_DAYS);
  const groups = {};
  readOps_(3000).forEach(o => {
    if (o.status === STATUS_CANCELLED || o.date < since) return;
    const k = o.type + '|' + o.cat + '|' + o.sub;
    (groups[k] = groups[k] || []).push(o.sum);
  });
  const res = {};
  Object.keys(groups).forEach(k => {
    const a = groups[k].sort((x, y) => x - y);
    if (a.length >= 3) res[k] = a[Math.floor(a.length / 2)];
  });
  return res;
}


// Последние названия по каждой подстатье (до 6 штук) — для подсказок кнопками
function recentNames_() {
  const res = {};
  readOps_(3000).reverse().forEach(o => {
    if (!o.name || o.status === STATUS_CANCELLED) return;
    const k = o.type + '|' + o.cat + '|' + o.sub;
    const list = res[k] = res[k] || [];
    if (list.length < 6 && !list.some(x => x.name.toLowerCase() === o.name.toLowerCase())) list.push({ name: o.name, unit: o.unit });
  });
  return res;
}

// Отменённые записи подсвечиваются красным и зачёркиваются (условное форматирование)
function ensureCancelFormat_(sh) {
  const formula = '=$L2="' + STATUS_CANCELLED + '"';
  const rules = sh.getConditionalFormatRules();
  const has = rules.some(r => {
    const c = r.getBooleanCondition();
    return c && c.getCriteriaValues && String(c.getCriteriaValues()[0]) === formula;
  });
  if (has) return;
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground('#f4cccc').setFontColor('#990000').setStrikethrough(true)
    .setRanges([sh.getRange('A2:P')]).build();
  sh.setConditionalFormatRules(rules.concat([rule]));
}

// Заголовки новых колонок, если таблица создана старой версией
function ensureHeaders_(sh) {
  if (sh.getRange(1, OPS_HEADERS.length).getValue() !== OPS_HEADERS[OPS_HEADERS.length - 1]) {
    sh.getRange(1, 1, 1, OPS_HEADERS.length).setValues([OPS_HEADERS]).setFontWeight('bold').setBackground('#e8eef3');
  }
}


// ===================== ЗАПИСЬ ОПЕРАЦИИ =====================
function addOp_(user, op) {
  if (!op) throw new Error('ERR_DATA');

  // ID создаёт приложение. Повторная отправка с тем же ID не создаст дубль.
  const id = String(op.id || '').trim();
  if (!/^[A-Za-z0-9-]{8,64}$/.test(id)) throw new Error('ERR_DATA');

  const v = validateOp_(user, op);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (readOps_(1000).some(o => o.id === id)) return { id: id, duplicate: true };
    const sh = sheet_(SHEET_OPS);
    ensureHeaders_(sh);
    sh.appendRow([
      id, dateSerial_(v.date), v.type, v.cat, v.sub, v.amount, v.pay,
      v.party, v.comment, user.name, new Date(), '', '', v.name, v.qty, v.unit
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { id: id, duplicate: false };
}


// Проверка полей операции — общая для записи и изменения
function validateOp_(user, op) {
  if (!op) throw new Error('ERR_DATA');
  const type = String(op.type || '');
  if (TYPES.indexOf(type) < 0) throw new Error('ERR_DATA');
  if (user.role !== 'boss' && type !== 'Расход') throw new Error('ERR_ROLE');

  const amount = Math.round(Number(op.amount));
  if (!isFinite(amount) || amount <= 0) throw new Error('ERR_AMOUNT');
  if (amount > MAX_AMOUNT) throw new Error('ERR_AMOUNT_BIG');

  const ref = readRef_();
  if (!ref.articles.some(a => a.type === type && a.cat === op.cat && a.sub === op.sub)) throw new Error('ERR_ARTICLE');
  if (!ref.pays.some(p => p.ru === op.pay)) throw new Error('ERR_PAY');

  const date = String(op.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('ERR_DATE');
  if (date > today_()) throw new Error('ERR_DATE_FUTURE');

  // Наименование обязательно для расхода. Количество — по желанию, но если есть, то больше нуля и с единицей.
  const name = safeText_(op.name, 100);
  if (type === 'Расход' && !name) throw new Error('ERR_NAME');
  let qty = '', unit = '';
  if (op.qty !== undefined && op.qty !== null && String(op.qty).trim() !== '') {
    qty = Number(String(op.qty).replace(',', '.'));
    if (!isFinite(qty) || qty <= 0 || qty > 1000000000) throw new Error('ERR_QTY');
    unit = String(op.unit || '');
    if (!UNITS.some(u => u[0] === unit)) throw new Error('ERR_QTY');
  }

  return { type: type, cat: op.cat, sub: op.sub, amount: amount, pay: op.pay, date: date,
    party: safeText_(op.party, 100), comment: safeText_(op.comment, 500), name: name, qty: qty, unit: unit };
}

// Изменить: руководитель — любую запись, сотрудник — свою. В любое время.
function canEdit_(user, o) {
  if (o.status === STATUS_CANCELLED) return false;
  return user.role === 'boss' || o.who === user.name;
}

// Отменить: руководитель — любую в любое время,
// сотрудник — только свою и только CANCEL_MINUTES минут после записи.
function canCancel_(user, o, now) {
  if (!canEdit_(user, o)) return false;
  return user.role === 'boss' || now - o.created <= CANCEL_MINUTES * 60000;
}


// ===================== ИЗМЕНЕНИЕ ЗАПИСИ =====================
// Меняет дату, тип, статью, сумму, оплату, контрагента и комментарий.
// В колонке «Изменено» остаётся отметка: кто и когда исправил.
function editOp_(user, id, op) {
  const v = validateOp_(user, op);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const o = readOps_(5000).find(x => x.id === String(id));
    if (!o) throw new Error('ERR_NOT_FOUND');
    if (o.status === STATUS_CANCELLED) throw new Error('ERR_CANCELLED');
    if (!canEdit_(user, o)) throw new Error('ERR_ROLE');
    const sh = sheet_(SHEET_OPS);
    ensureHeaders_(sh);
    sh.getRange(o.row, 2).setNumberFormat('dd.mm.yyyy');
    sh.getRange(o.row, 2, 1, 8).setValues([[dateSerial_(v.date), v.type, v.cat, v.sub, v.amount, v.pay, v.party, v.comment]]);
    sh.getRange(o.row, 13, 1, 4).setValues([[user.name + ', ' + Utilities.formatDate(new Date(), TIMEZONE, 'dd.MM.yyyy HH:mm'), v.name, v.qty, v.unit]]);
    SpreadsheetApp.flush();
    return { id: o.id };
  } finally {
    lock.releaseLock();
  }
}


// ===================== ИСТОРИЯ =====================
// Руководитель видит все записи, сотрудник — только свои.
function history_(user, filter) {
  const now = Date.now();
  const res = [];
  const ops = readOps_(5000).reverse();
  for (let i = 0; i < ops.length && res.length < HISTORY_LIMIT; i++) {
    const o = ops[i];
    if (user.role !== 'boss' && o.who !== user.name) continue;
    if (filter.type && o.type !== filter.type) continue;
    if (filter.cat && o.cat !== filter.cat) continue;
    res.push({
      id: o.id, date: o.date, type: o.type, cat: o.cat, sub: o.sub, sum: o.sum, pay: o.pay,
      party: o.party, comment: o.comment, who: o.who, created: o.created,
      cancelled: o.status === STATUS_CANCELLED, edited: o.edited, name: o.name, qty: o.qty, unit: o.unit,
      canEdit: canEdit_(user, o),
      canCancel: canCancel_(user, o, now)
    });
  }
  return { items: res, now: now };
}


// ===================== ОТМЕНА ЗАПИСИ =====================
// Строка не удаляется, а получает статус «Отменена» и не попадает в отчёты.
function cancel_(user, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const o = readOps_(5000).find(x => x.id === String(id));
    if (!o) throw new Error('ERR_NOT_FOUND');
    if (o.status === STATUS_CANCELLED) return { id: o.id };
    if (!canEdit_(user, o)) throw new Error('ERR_ROLE');
    if (!canCancel_(user, o, Date.now())) throw new Error('ERR_CANCEL_TIME');
    const sh = sheet_(SHEET_OPS);
    sh.getRange(o.row, OPS_HEADERS.indexOf('Статус') + 1).setValue(STATUS_CANCELLED);
    ensureCancelFormat_(sh);
    return { id: o.id };
  } finally {
    lock.releaseLock();
  }
}


// ===================== ОТЧЁТ =====================
// Считает выбранный период и такой же по длине период перед ним (для «Динамики»).
function report_(user, from, to) {
  if (user.role !== 'boss') throw new Error('ERR_ROLE');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) throw new Error('ERR_DATE');

  const days = Math.round((Date.UTC(+to.slice(0, 4), +to.slice(5, 7) - 1, +to.slice(8)) -
    Date.UTC(+from.slice(0, 4), +from.slice(5, 7) - 1, +from.slice(8))) / 86400000) + 1;
  // Если период начинается с 1-го числа и не длиннее месяца — сравниваем с теми же днями прошлого месяца.
  // Иначе — с таким же по длине периодом прямо перед выбранным.
  let prevFrom, prevTo;
  if (from.slice(8) === '01' && days <= 31) {
    prevFrom = shiftMonth_(from);
    prevTo = shiftMonth_(to);
  } else {
    prevTo = addDays_(from, -1);
    prevFrom = addDays_(prevTo, -(days - 1));
  }

  const ops = readOps_(0).filter(o => o.status !== STATUS_CANCELLED);
  const cur = aggregate_(ops, from, to);
  const prev = aggregate_(ops, prevFrom, prevTo);

  const toList = (obj, total) => Object.keys(obj)
    .map(k => ({ name: k, sum: obj[k], share: total ? obj[k] / total : 0 }))
    .sort((a, b) => b.sum - a.sum);

  const cats = {};
  Object.keys(cur.expByCat).concat(Object.keys(prev.expByCat)).forEach(k => { cats[k] = true; });
  const dynamics = Object.keys(cats).map(k => {
    const now = cur.expByCat[k] || 0, before = prev.expByCat[k] || 0;
    return { name: k, now: now, before: before, diff: now - before, pct: before ? (now - before) / before : null };
  }).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));

  const pays = {};
  Object.keys(cur.payExp).concat(Object.keys(cur.payInc)).forEach(k => { pays[k] = true; });

  return {
    from: from, to: to, prevFrom: prevFrom, prevTo: prevTo, count: cur.count,
    totals: { income: cur.income, expense: cur.expense, diff: cur.income - cur.expense, invest: cur.invest, transfer: cur.transfer },
    prevTotals: { income: prev.income, expense: prev.expense, diff: prev.income - prev.expense },
    expByCat: toList(cur.expByCat, cur.expense),
    incByCat: toList(cur.incByCat, cur.income),
    dynamics: dynamics,
    byPay: Object.keys(pays).map(k => ({ name: k, expense: cur.payExp[k] || 0, income: cur.payInc[k] || 0 }))
  };
}

function aggregate_(ops, from, to) {
  const a = { income: 0, expense: 0, invest: 0, transfer: 0, count: 0, expByCat: {}, incByCat: {}, payExp: {}, payInc: {} };
  ops.forEach(o => {
    if (o.date < from || o.date > to) return;
    a.count++;
    if (o.type === 'Расход') {
      a.expense += o.sum;
      a.expByCat[o.cat] = (a.expByCat[o.cat] || 0) + o.sum;
      a.payExp[o.pay] = (a.payExp[o.pay] || 0) + o.sum;
    } else if (o.type === 'Доход') {
      a.income += o.sum;
      a.incByCat[o.cat] = (a.incByCat[o.cat] || 0) + o.sum;
      a.payInc[o.pay] = (a.payInc[o.pay] || 0) + o.sum;
    } else if (o.type === 'Вложение') {
      a.invest += o.sum;
    } else if (o.type === 'Перевод') {
      a.transfer += o.sum;
    }
  });
  return a;
}


// ===================== ВСПОМОГАТЕЛЬНОЕ =====================
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('ERR_SETUP');
  return sh;
}

// Читает последние `limit` строк листа «Операции» (0 = все строки)
function readOps_(limit) {
  const sh = sheet_(SHEET_OPS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = limit ? Math.max(2, last - limit + 1) : 2;
  const vals = sh.getRange(start, 1, last - start + 1, OPS_HEADERS.length).getValues();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const s = (x) => String(x === null || x === undefined ? '' : x).trim();
  const out = [];
  vals.forEach((r, i) => {
    if (!s(r[0])) return;
    out.push({
      row: start + i, id: s(r[0]),
      date: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd')
        : typeof r[1] === 'number' ? serialToDate_(r[1]) : s(r[1]),
      type: s(r[2]), cat: s(r[3]), sub: s(r[4]), sum: Number(r[5]) || 0, pay: s(r[6]),
      party: s(r[7]), comment: s(r[8]), who: s(r[9]),
      created: r[10] instanceof Date ? r[10].getTime() : 0, status: s(r[11]), edited: s(r[12]),
      name: s(r[13]), qty: r[14] === '' || r[14] === null ? '' : Number(r[14]), unit: s(r[15])
    });
  });
  return out;
}

function today_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
}

// Прибавить n дней к дате 'yyyy-MM-dd'
function addDays_(s, n) {
  const d = new Date(Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10) + n));
  return d.toISOString().slice(0, 10);
}

// Та же дата месяцем раньше (31 марта → 28/29 февраля)
function shiftMonth_(s) {
  const y = +s.slice(0, 4), m = +s.slice(5, 7) - 1, d = +s.slice(8, 10);
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate(); // дней в прошлом месяце
  return new Date(Date.UTC(y, m - 1, Math.min(d, last))).toISOString().slice(0, 10);
}

// Дата как число дней (так её хранит Google Таблица) — не зависит от часового пояса
function dateSerial_(s) {
  return (Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000;
}

// Обратно: число дней → 'yyyy-MM-dd' (если в ячейке дата сохранилась как число)
function serialToDate_(n) {
  return new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000).toISOString().slice(0, 10);
}

// Обрезает текст и не даёт ему стать формулой в таблице
function safeText_(x, max) {
  let t = String(x === null || x === undefined ? '' : x).trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}
