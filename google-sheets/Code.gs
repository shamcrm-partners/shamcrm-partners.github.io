/**
 * Приём анкет и результатов экзамена со страницы партнёров shamCRM
 * (https://shamcrm-partners.github.io/) в эту Google Таблицу.
 *
 * Анкета    → новая строка.
 * Экзамен   → дописывается в строку того же партнёра.
 *
 * Как установить: Расширения → Apps Script → вставить этот код целиком →
 * Сохранить → Начать развёртывание → Новое развёртывание → Веб-приложение →
 * «Запуск от имени»: Я, «У кого есть доступ»: Все → Развернуть → Разрешить доступ.
 * Полученный адрес (…/exec) вписывается в config.js сайта в поле sheetUrl.
 */

var COLUMNS = [
  '№', 'Статус', 'Дата анкеты', 'Компания', 'Юр. наименование', 'ИНН', 'Сайт',
  'Руководитель', 'Контактное лицо', 'Должность', 'Телефон', 'Почта',
  'Потенц. клиентов', 'Сотрудников на shamCRM', 'Откуда узнали', 'Анкета полностью',
  'Экзамен: дата', 'Экзамен: ФИО', 'Экзамен: формат', 'Экзамен: результат',
  'Экзамен: %', 'Экзамен: сдан', 'Экзамен: по модулям',
  'Комментарий менеджера', 'ID устройства', 'Страница'
];
var STATUSES = ['Новая анкета', 'Связались', 'Договор отправлен', 'Договор подписан',
  'Обучение', 'Экзамен сдан', 'Работает', 'Отказ'];

// Номер колонки по названию (1, 2, 3…)
function col(name) {
  var i = COLUMNS.indexOf(name);
  if (i < 0) throw new Error('Нет колонки ' + name);
  return i + 1;
}

function doGet() {
  return ContentService.createTextOutput('Приём анкет shamCRM работает');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = {};
    try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return out({ ok: false, error: 'bad_json' }); }
    if (d.type !== 'anketa' && d.type !== 'exam') return out({ ok: false, error: 'bad_type' });

    var sh = getSheet();
    var row;

    if (d.type === 'anketa') {
      if (!d.phone && !d.email && !d.contact_name) return out({ ok: false, error: 'no_contact' });
      // Повторная отправка (исправили ответы) — та же строка, если совпадает устройство
      row = rowByLeadId(sh, d);
      var isNew = !row;
      if (isNew) row = Math.max(sh.getLastRow(), 1) + 1;
      setCells(sh, row, {
        '№': row - 1,
        'Дата анкеты': toDate(d.submitted_at),
        'Компания': d.company,
        'Юр. наименование': d.legal_name,
        'ИНН': d.inn,
        'Сайт': d.site,
        'Руководитель': d.ceo,
        'Контактное лицо': d.contact_name,
        'Должность': d.position,
        'Телефон': d.phone,
        'Почта': d.email,
        'Потенц. клиентов': d.potential_clients,
        'Сотрудников на shamCRM': d.team_size,
        'Откуда узнали': d.source,
        'Анкета полностью': d.text,
        'ID устройства': d.client_id,
        'Страница': d.page_url
      });
      if (isNew) setStatus(sh, row, 'Новая анкета');
      return out({ ok: true, lead_id: String(row) });
    }

    // Экзамен: ищем строку партнёра — по номеру, по устройству, по телефону/почте
    row = rowByLeadId(sh, d) || lastRowWhere(sh, 'ID устройства', d.client_id) ||
          lastRowWhere(sh, 'Телефон', d.phone) || lastRowWhere(sh, 'Почта', d.email);
    var newRow = !row;
    if (newRow) {
      row = Math.max(sh.getLastRow(), 1) + 1;
      setCells(sh, row, {
        '№': row - 1,
        'Компания': d.company,
        'Контактное лицо': d.name,
        'Телефон': d.phone,
        'Почта': d.email,
        'Анкета полностью': 'Экзамен без анкеты',
        'ID устройства': d.client_id,
        'Страница': d.page_url
      });
    }
    var modules = (d.modules || []).map(function (m) { return m.correct + '/' + m.total; }).join(' · ');
    setCells(sh, row, {
      'Экзамен: дата': toDate(d.finished_at),
      'Экзамен: ФИО': d.name,
      'Экзамен: формат': d.mode === 'full_240' ? 'Полный тест (240)' : 'Экзамен партнёра (40)',
      'Экзамен: результат': num(d.correct) + ' / ' + num(d.total),
      'Экзамен: %': num(d.percent),
      'Экзамен: сдан': d.passed === true ? 'Да' : 'Нет',
      'Экзамен: по модулям': modules
    });
    var status = sh.getRange(row, col('Статус')).getValue();
    if (d.passed === true && status !== 'Отказ' && status !== 'Работает') {
      setStatus(sh, row, 'Экзамен сдан');
    } else if (newRow) {
      setStatus(sh, row, 'Новая анкета');
    }
    return out({ ok: true, lead_id: String(row) });
  } finally {
    lock.releaseLock();
  }
}

// Первый лист таблицы; при первом запуске пишет заголовки и оформление
function getSheet() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getRange(1, 1).getValue() !== COLUMNS[0]) {
    sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  }
  if (sh.getFrozenRows() < 1) {
    sh.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Статус с выпадающим списком в ячейке
function setStatus(sh, row, value) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(true).build();
  sh.getRange(row, col('Статус')).setDataValidation(rule).setValue(value);
}

// Строка по номеру из ответа сайту — только если совпадает устройство партнёра
function rowByLeadId(sh, d) {
  var r = parseInt(d.lead_id, 10);
  if (!r || r < 2 || r > sh.getLastRow() || !d.client_id) return 0;
  return sh.getRange(r, col('ID устройства')).getValue() === d.client_id ? r : 0;
}

// Последняя строка, где в колонке name стоит value
function lastRowWhere(sh, name, value) {
  if (!value) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, col(name), last - 1, 1).getValues();
  var want = String(value).trim().replace(/^'/, '');
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim().replace(/^'/, '') === want) return i + 2;
  }
  return 0;
}

function setCells(sh, row, map) {
  Object.keys(map).forEach(function (name) {
    var v = map[name];
    if (v === undefined || v === null || v === '') return;
    sh.getRange(row, col(name)).setValue(v instanceof Date || typeof v === 'number' ? v : safe(v));
  });
}

// Текст без формул (= + - @ в начале) и не длиннее 20 000 символов
function safe(v) {
  var s = String(v).slice(0, 20000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function toDate(iso) {
  var dt = new Date(iso);
  return isNaN(dt.getTime()) ? '' : dt;
}

function num(v) {
  var n = Number(v);
  return isFinite(n) ? n : '';
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
