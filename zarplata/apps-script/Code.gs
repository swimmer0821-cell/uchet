/**
 * =============================================================
 *  Зарплата SUB PROFI — серверная часть (Google Apps Script)
 * =============================================================
 *  Листы таблицы:
 *   «Сотрудники» — кто работает, как считается зарплата, сколько обычно на карту
 *   «Ведомость»  — по строке на сотрудника в каждом месяце
 *   «Выплаты»    — каждая выдача денег (карта или наличные), с датой
 *   «Пользователи» — кто входит в приложение и его PIN
 *
 *  Как считается (премии в расчёт не входят):
 *   Полная ЗП      = оклад за месяц, или ставка × смены, или ставка × дни
 *   На карту       = официальная часть «на руки» (уже после налога)
 *   Наличными      = Полная ЗП − На карту
 *   Официально начислено = На карту ÷ 0,88;  налог 12% = начислено − на карту
 *   Соцналог 12% от начисления платит компания сверху
 *   Затраты на зарплату (для себестоимости) = Полная ЗП + налог 12% + соцналог 12%
 *   Осталось выдать = (к выдаче + долг прошлого месяца) − выплаты
 *  Отменённые выплаты не считаются и не удаляются.
 * =============================================================
 */

const SH_EMP = 'Сотрудники';
const SH_ROWS = 'Ведомость';
const SH_PAY = 'Выплаты';
const SH_USERS = 'Пользователи';
const TIMEZONE = 'Asia/Tashkent';
const TAX = 0.12;          // подоходный налог + ИНПС
const INPS = 0.001;        // ИНПС внутри 12%
const SOC = 0.12;          // социальный налог, платит компания сверх начисления
const CANCELLED = 'Отменена';
const TYPES = ['Месяц', 'Смена', 'День'];
const PAY_KINDS = ['Карта', 'Наличные'];

// Связь с приложением «Деньги»: выдача НАЛИЧНЫМИ сразу записывается там расходом.
const MONEY_SHEET_ID = '18mgpXfwxlSQjfXNWjlOHpLGu_0Or5V9Z1ynmxXtFwVM';   // таблица «Учёт денег SUB PROFI»
const MONEY_LINK_KINDS = ['Наличные'];                                   // карту пока не связываем
const MONEY_ART = { type: 'Расход', cat: 'Зарплата', sub: 'Наличная часть', pay: 'Наличные' };

const EMP_H = ['ID', 'Имя', 'ФИО (официально)', 'Должность / заметка', 'Способ', 'Ставка', 'На карту (обычно)', 'Работает', 'Комментарий'];
const ROW_H = ['Месяц', 'ID сотрудника', 'Имя', 'Способ', 'Ставка', 'Кол-во смен/дней', 'Полная ЗП', 'На карту', 'Комментарий', 'Изменил', 'Когда'];
const PAY_H = ['ID', 'Месяц', 'ID сотрудника', 'Дата', 'Вид', 'Сумма', 'Комментарий', 'Кто', 'Когда', 'Статус'];
const USER_H = ['Пользователь', 'PIN'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Зарплата').addItem('Настроить таблицу', 'setup').addToUi();
}

// Запустить один раз: создаёт недостающие листы и заголовки. Данные не трогает.
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TIMEZONE);
  const mk = (name, head) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8eef3');
    sh.setFrozenRows(1);
    return sh;
  };
  const emp = mk(SH_EMP, EMP_H), rows = mk(SH_ROWS, ROW_H), pay = mk(SH_PAY, PAY_H), users = mk(SH_USERS, USER_H);
  rows.getRange('A2:A').setNumberFormat('@');
  pay.getRange('B2:B').setNumberFormat('@');
  pay.getRange('D2:D').setNumberFormat('dd.mm.yyyy');
  users.getRange('B2:B').setNumberFormat('@');
  ['F2:G', 'F2:F'].forEach(a => emp.getRange(a).setNumberFormat('#,##0'));
  rows.getRange('E2:E').setNumberFormat('#,##0'); rows.getRange('G2:H').setNumberFormat('#,##0');
  pay.getRange('F2:F').setNumberFormat('#,##0');
  emp.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).build());
  emp.getRange('H2:H').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Да', 'Нет'], true).build());
  pay.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PAY_KINDS, true).build());
  // отменённые выплаты — красным и зачёркнуто
  const f = '=$J2="' + CANCELLED + '"';
  const rules = pay.getConditionalFormatRules().filter(r => { const c = r.getBooleanCondition(); return !(c && String(c.getCriteriaValues()[0]) === f); });
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setBackground('#f4cccc').setFontColor('#990000')
    .setStrikethrough(true).setRanges([pay.getRange('A2:J')]).build());
  pay.setConditionalFormatRules(rules);
  ['Лист1', 'Sheet1'].forEach(n => { const s = ss.getSheetByName(n); if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
  Logger.log('Готово. Таблица настроена.');
}

// ===================== ПРИЁМ ЗАПРОСОВ =====================
function doGet() { return out_({ ok: true, message: 'Зарплата SUB PROFI работает' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'ERR_DATA' }); }
  try {
    if (req.action === 'users') return out_({ ok: true, data: readUsers_().map(u => ({ name: u.name, hasPin: !!u.pin })) });
    if (req.action === 'setPin') return out_({ ok: true, data: setPin_(req.user, req.newPin) });
    const user = auth_(req.user, req.pin);
    let data;
    switch (req.action) {
      case 'init':      data = init_(user); break;
      case 'month':     data = month_(String(req.month || '')); break;
      case 'saveRow':   data = withLock_(() => saveRow_(user, req.row)); break;
      case 'addPay':    data = withLock_(() => addPay_(user, req.pay)); break;
      case 'payAll':    data = withLock_(() => payAll_(user, String(req.month || ''), req.kind, req.empId)); break;
      case 'cancelPay': data = withLock_(() => cancelPay_(user, req.id)); break;
      case 'newMonth':  data = withLock_(() => newMonth_(user, String(req.month || ''))); break;
      case 'saveEmp':   data = withLock_(() => saveEmp_(user, req.emp)); break;
      default: throw new Error('ERR_DATA');
    }
    return out_({ ok: true, data: data });
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

// ===================== ПОЛЬЗОВАТЕЛИ =====================
function readUsers_() {
  return values_(SH_USERS).filter(r => s_(r[0])).map(r => ({ name: s_(r[0]), pin: s_(r[1]) }));
}
function auth_(name, pin) {
  const u = readUsers_().find(x => x.name === s_(name));
  if (!u) throw new Error('ERR_USER');
  if (!u.pin) throw new Error('ERR_NO_PIN');
  if (String(pin || '') !== u.pin) throw new Error('ERR_PIN');
  return u;
}
// Первый вход: человек сам придумывает PIN. Сброс — очистить ячейку PIN.
function setPin_(name, newPin) {
  const pin = String(newPin || '');
  if (!/^\d{4}$/.test(pin)) throw new Error('ERR_PIN_FORMAT');
  return withLock_(() => {
    const sh = sheet_(SH_USERS), v = sh.getDataRange().getValues();
    for (let r = 1; r < v.length; r++) {
      if (s_(v[r][0]) && s_(v[r][0]) === s_(name)) {
        if (s_(v[r][1])) throw new Error('ERR_PIN_EXISTS');
        sh.getRange(r + 1, 2).setNumberFormat('@').setValue(pin);
        return { ok: true };
      }
    }
    throw new Error('ERR_USER');
  });
}

// ===================== ЧТЕНИЕ =====================
function readEmps_() {
  return values_(SH_EMP).map((r, i) => ({
    row: i + 2, id: s_(r[0]), name: s_(r[1]), fio: s_(r[2]), role: s_(r[3]),
    type: TYPES.indexOf(s_(r[4])) >= 0 ? s_(r[4]) : 'Месяц', rate: n_(r[5]), card: n_(r[6]),
    active: s_(r[7]) !== 'Нет', comment: s_(r[8])
  })).filter(e => e.id && e.name);
}
function readRows_() {
  return values_(SH_ROWS).map((r, i) => ({
    row: i + 2, month: monthKey_(r[0]), empId: s_(r[1]), name: s_(r[2]), type: s_(r[3]) || 'Месяц', rate: n_(r[4]),
    qty: r[5] === '' || r[5] === null ? '' : n_(r[5]), full: n_(r[6]), card: n_(r[7]), comment: s_(r[8]),
    by: s_(r[9]), at: r[10] instanceof Date ? fmtDT_(r[10]) : s_(r[10])
  })).filter(x => x.month && x.empId);
}
function readPays_() {
  return values_(SH_PAY).map((r, i) => ({
    row: i + 2, id: s_(r[0]), month: monthKey_(r[1]), empId: s_(r[2]),
    date: r[3] instanceof Date ? Utilities.formatDate(r[3], TIMEZONE, 'yyyy-MM-dd') : s_(r[3]),
    kind: s_(r[4]), sum: n_(r[5]), comment: s_(r[6]), by: s_(r[7]), at: r[8] instanceof Date ? fmtDT_(r[8]) : s_(r[8]),
    cancelled: s_(r[9]) === CANCELLED
  })).filter(p => p.id && p.month);
}

function init_(user) {
  const months = Array.from(new Set(readRows_().map(r => r.month))).sort();
  return { user: user.name, months: months, today: today_(), emps: readEmps_() };
}

// Месяц целиком: строки, выплаты, долг с прошлого месяца и итоги
function month_(m) {
  if (!/^\d{4}-\d{2}$/.test(m)) throw new Error('ERR_DATA');
  const rows = readRows_(), pays = readPays_().filter(p => !p.cancelled);
  const months = Array.from(new Set(rows.map(r => r.month))).sort().filter(x => x <= m);
  // долг переносится из месяца в месяц по каждому сотруднику
  let carry = {};
  let result = null;
  months.forEach(mm => {
    const list = rows.filter(r => r.month === mm).map(r => {
      const c = carry[r.empId] || { card: 0, cash: 0 };
      const cash = Math.max(0, r.full - r.card);
      const paidCard = sum_(pays.filter(p => p.month === mm && p.empId === r.empId && p.kind === 'Карта'));
      const paidCash = sum_(pays.filter(p => p.month === mm && p.empId === r.empId && p.kind === 'Наличные'));
      const gross = Math.round(r.card / (1 - TAX));
      return Object.assign({}, r, {
        cash: cash, prevCard: c.card, prevCash: c.cash, paidCard: paidCard, paidCash: paidCash,
        leftCard: r.card + c.card - paidCard, leftCash: cash + c.cash - paidCash,
        gross: gross, tax: gross - r.card, inps: Math.round(gross * INPS), soc: Math.round(gross * SOC),
        cost: r.full + (gross - r.card) + Math.round(gross * SOC)
      });
    });
    const next = {};
    list.forEach(x => { next[x.empId] = { card: x.leftCard, cash: x.leftCash }; });
    // если сотрудника нет в этом месяце, его долг переносится дальше
    Object.keys(carry).forEach(k => { if (!next[k]) next[k] = carry[k]; });
    carry = next;
    if (mm === m) result = list;
  });
  if (!result) result = [];
  const t = (k) => result.reduce((a, x) => a + (Number(x[k]) || 0), 0);
  return {
    month: m, rows: result,
    pays: readPays_().filter(p => p.month === m),
    total: { full: t('full'), card: t('card'), cash: t('cash'), gross: t('gross'), tax: t('tax'), soc: t('soc'), cost: t('cost'),
      paidCard: t('paidCard'), paidCash: t('paidCash'), leftCard: t('leftCard'), leftCash: t('leftCash'), people: result.length }
  };
}

// ===================== ИЗМЕНЕНИЯ =====================
function saveRow_(user, x) {
  if (!x) throw new Error('ERR_DATA');
  const m = s_(x.month), id = s_(x.empId);
  const r = readRows_().find(z => z.month === m && z.empId === id);
  if (!r) throw new Error('ERR_NOT_FOUND');
  const type = TYPES.indexOf(s_(x.type)) >= 0 ? s_(x.type) : r.type;
  const rate = money_(x.rate), full = money_(x.full), card = money_(x.card);
  const qty = x.qty === '' || x.qty === null || x.qty === undefined ? '' : Number(String(x.qty).replace(',', '.'));
  if (qty !== '' && (!isFinite(qty) || qty < 0 || qty > 31)) throw new Error('ERR_QTY');
  sheet_(SH_ROWS).getRange(r.row, 4, 1, 8).setValues([[type, rate, qty, full, card, safe_(x.comment, 300), user.name, new Date()]]);
  return { ok: true };
}

function addPay_(user, p) {
  if (!p) throw new Error('ERR_DATA');
  const m = s_(p.month), id = s_(p.empId);
  if (!readRows_().some(z => z.month === m && z.empId === id)) throw new Error('ERR_NOT_FOUND');
  if (PAY_KINDS.indexOf(s_(p.kind)) < 0) throw new Error('ERR_DATA');
  const sum = money_(p.sum);
  if (sum <= 0) throw new Error('ERR_AMOUNT');
  const date = s_(p.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date > today_()) throw new Error('ERR_DATE');
  const pid = s_(p.id) || Utilities.getUuid();
  if (readPays_().some(z => z.id === pid)) return { id: pid, duplicate: true };
  const row = [pid, "'" + m, id, dateObj_(date), s_(p.kind), sum, safe_(p.comment, 200), user.name, new Date(), ''];
  moneyAdd_(user, [row]);                      // сначала «Деньги»: если не получилось — выплату не записываем
  sheet_(SH_PAY).appendRow(row);
  return { id: pid };
}

// «Выдано всё»: записывает выплаты на весь остаток — одному сотруднику или всем в месяце
function payAll_(user, m, kind, empId) {
  if (PAY_KINDS.indexOf(s_(kind)) < 0) throw new Error('ERR_DATA');
  const d = month_(m), now = new Date(), date = today_();
  const out = [];
  d.rows.filter(r => !empId || r.empId === s_(empId)).forEach(r => {
    const left = kind === 'Карта' ? r.leftCard : r.leftCash;
    if (left > 0) out.push([Utilities.getUuid(), "'" + m, r.empId, dateObj_(date), kind, left, 'выдано всё', user.name, now, '']);
  });
  if (out.length) { moneyAdd_(user, out); const sh = sheet_(SH_PAY); sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out); }
  return { count: out.length };
}

function cancelPay_(user, id) {
  const p = readPays_().find(z => z.id === s_(id));
  if (!p) throw new Error('ERR_NOT_FOUND');
  if (!p.cancelled) {
    if (MONEY_LINK_KINDS.indexOf(p.kind) >= 0) moneyCancel_(p.id);
    sheet_(SH_PAY).getRange(p.row, 10).setValue(CANCELLED);
  }
  return { ok: true };
}

// Новый месяц: строки для всех, кто работает, по данным карточки сотрудника
function newMonth_(user, m) {
  if (!/^\d{4}-\d{2}$/.test(m)) throw new Error('ERR_DATA');
  const rows = readRows_();
  if (rows.some(r => r.month === m)) return { exists: true };
  const now = new Date();
  const out = readEmps_().filter(e => e.active).map(e => {
    const qty = e.type === 'Месяц' ? '' : 0;
    const full = e.type === 'Месяц' ? e.rate : 0;
    return ["'" + m, e.id, e.name, e.type, e.rate, qty, full, e.card, '', user.name, now];
  });
  if (!out.length) throw new Error('ERR_NO_EMPS');
  const sh = sheet_(SH_ROWS);
  sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  return { count: out.length };
}

// Добавить или изменить сотрудника. Новый сотрудник сразу попадает в последний открытый месяц.
function saveEmp_(user, e) {
  if (!e || !s_(e.name)) throw new Error('ERR_NAME');
  const type = TYPES.indexOf(s_(e.type)) >= 0 ? s_(e.type) : 'Месяц';
  const vals = [s_(e.name).slice(0, 60), safe_(e.fio, 80), safe_(e.role, 80), type, money_(e.rate), money_(e.card), e.active === false ? 'Нет' : 'Да', safe_(e.comment, 300)];
  const sh = sheet_(SH_EMP), emps = readEmps_();
  if (s_(e.id)) {
    const old = emps.find(z => z.id === s_(e.id));
    if (!old) throw new Error('ERR_NOT_FOUND');
    sh.getRange(old.row, 2, 1, vals.length).setValues([vals]);
    return { id: old.id };
  }
  const id = 'E' + String(emps.reduce((a, z) => Math.max(a, Number(z.id.replace(/\D/g, '')) || 0), 0) + 1).padStart(3, '0');
  sh.appendRow([id].concat(vals));
  const months = Array.from(new Set(readRows_().map(r => r.month))).sort();
  const last = months[months.length - 1];
  if (last && e.addToMonth !== false) {
    sheet_(SH_ROWS).appendRow(["'" + last, id, vals[0], type, vals[4], type === 'Месяц' ? '' : 0, type === 'Месяц' ? vals[4] : 0, vals[5], '', user.name, new Date()]);
  }
  return { id: id };
}

// ===================== СВЯЗЬ С «ДЕНЬГАМИ» =====================
// Строка листа «Операции» в «Учёт денег»: ID, Дата, Тип, Статья, Подстатья, Сумма, Способ оплаты,
// Контрагент, Комментарий, Кто внёс, Время записи, Статус, Изменено, Наименование, Количество, Ед., Куда
function moneySheet_() {
  try { return SpreadsheetApp.openById(MONEY_SHEET_ID).getSheetByName('Операции'); }
  catch (e) { throw new Error('ERR_MONEY'); }
}
function moneyAdd_(user, payRows) {
  const rows = payRows.filter(r => MONEY_LINK_KINDS.indexOf(r[4]) >= 0);
  if (!rows.length) return;
  const sh = moneySheet_();
  if (!sh) throw new Error('ERR_MONEY');
  const emps = readEmps_();
  const out = rows.map(r => {
    const m = String(r[1]).replace(/^'/, ''), e = emps.find(x => x.id === r[2]) || { name: r[2] };
    const note = 'Зарплата за ' + monthName_(m) + (r[6] ? ' · ' + r[6] : '') + ' · из приложения «Зарплата»';
    return ['zp-' + r[0], r[3], MONEY_ART.type, MONEY_ART.cat, MONEY_ART.sub, r[5], MONEY_ART.pay,
      '', note, user.name, new Date(), '', '', e.name, '', '', ''];
  });
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, out.length, out[0].length).setValues(out);
  sh.getRange(start, 2, out.length, 1).setNumberFormat('dd.mm.yyyy');
}
function moneyCancel_(payId) {
  const sh = moneySheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === 'zp-' + payId) { sh.getRange(i + 2, 12).setValue('Отменена'); return; }
  }
}
function monthName_(m) {
  const n = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь'];
  return n[Number(m.slice(5)) - 1] + ' ' + m.slice(0, 4);
}
// Запустить один раз из редактора: Google спросит разрешение на доступ к таблице «Учёт денег».
function checkMoneyLink() {
  Logger.log('Связь с «Деньгами» работает: ' + SpreadsheetApp.openById(MONEY_SHEET_ID).getName());
}

// ===================== МЕЛОЧИ =====================
function sheet_(n) { const sh = SpreadsheetApp.getActive().getSheetByName(n); if (!sh) throw new Error('ERR_SETUP'); return sh; }
function values_(n) { const sh = sheet_(n), last = sh.getLastRow(); if (last < 2) return []; return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues(); }
function s_(x) { return String(x === null || x === undefined ? '' : x).trim(); }
function n_(x) { const v = Number(String(x === null || x === undefined ? '' : x).replace(/\s/g, '').replace(',', '.')); return isFinite(v) ? v : 0; }
function money_(x) { const v = Math.round(n_(x)); if (v < 0 || v > 1000000000) throw new Error('ERR_AMOUNT'); return v; }
function sum_(list) { return list.reduce((a, p) => a + p.sum, 0); }
function safe_(x, max) { let t = s_(x).slice(0, max); if (/^[=+\-@]/.test(t)) t = "'" + t; return t; }
function monthKey_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, TIMEZONE, 'yyyy-MM');
  const t = s_(x).replace(/^'/, '');
  return /^\d{4}-\d{2}$/.test(t) ? t : '';
}
function today_() { return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'); }
function dateObj_(s) { return new Date(Number(s.slice(0, 4)), Number(s.slice(5, 7)) - 1, Number(s.slice(8, 10)), 12); }
function fmtDT_(d) { return Utilities.formatDate(d, TIMEZONE, 'dd.MM.yyyy HH:mm'); }
