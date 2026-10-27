/** Administration: every web endpoint checks a short-lived server-side session. */
var OprosAdmin = (function () {
  'use strict';
  function hash(value) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value)).map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
  }
  function requireSession(token) {
    if (typeof token !== 'string' || !/^[a-f0-9-]{72}$/.test(token)) throw new Error('Войдите как администратор.');
    var credential = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD_HASH');
    if (!credential || CacheService.getScriptCache().get('admin:session:' + hash(token)) !== credential) throw new Error('Сессия истекла. Войдите повторно.');
  }
  function snapshot(book) {
    var config = {};
    ['Settings', 'Sections', 'Questions', 'Lists'].forEach(function (name) { config[name.toLowerCase()] = OprosStore.readRows(book, name); });
    config.revision = hash(JSON.stringify(config));
    return config;
  }
  function withLock(fn) {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) throw new Error('Сервис занят. Повторите через несколько секунд.');
    try { return fn(); } finally { lock.releaseLock(); }
  }
  function login(password) {
    return withLock(function () {
      var expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD_HASH');
      if (!expected) throw new Error('Владелец должен один раз запустить setupAdmin в редакторе Apps Script.');
      var cache = CacheService.getScriptCache(), key = 'admin:attempts:' + Math.floor(Date.now() / 300000);
      var count = Number(cache.get(key) || 0);
      if (count >= 10) throw new Error('Слишком много попыток входа. Подождите 5 минут.');
      cache.put(key, String(count + 1), 300);
      if (typeof password !== 'string' || password.length > 200 || hash(password) !== expected) throw new Error('Неверный пароль администратора.');
      var token = Utilities.getUuid() + Utilities.getUuid();
      cache.put('admin:session:' + hash(token), expected, 1800);
      return { token: token, expiresIn: 1800 };
    });
  }
  function logout(token) {
    if (typeof token === 'string' && /^[a-f0-9-]{72}$/.test(token)) CacheService.getScriptCache().remove('admin:session:' + hash(token));
    return { ok: true };
  }
  function load(token) { requireSession(token); return snapshot(OprosStore.workbook()); }
  var fields = {
    settings: ['key', 'value', 'description'],
    sections: ['id', 'title', 'description', 'position', 'enabled'],
    questions: ['id', 'section_id', 'title', 'description', 'type', 'required', 'options_source', 'options_json', 'visible_if_json', 'validation_json', 'position', 'enabled']
  };
  function prepare(book, key, rows) {
    if (!Array.isArray(rows) || !rows.length || rows.length > (key === 'questions' ? 100 : 200)) throw new Error('Недопустимое число строк: ' + key);
    var name = key[0].toUpperCase() + key.slice(1), sheet = book.getSheetByName(name);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.length !== fields[key].length || headers.some(function (h) { return fields[key].indexOf(h) < 0; })) throw new Error('Редактор поддерживает стандартные столбцы листа ' + name + '.');
    var clean = rows.map(function (row) {
      if (!row || typeof row !== 'object' || Array.isArray(row)) throw new Error('Некорректная строка ' + name);
      var result = {};
      fields[key].forEach(function (field) {
        var value = row[field] == null ? '' : row[field];
        if (['string', 'number', 'boolean'].indexOf(typeof value) < 0 || String(value).length > 20000) throw new Error('Некорректное поле ' + field);
        result[field] = value;
      });
      return result;
    });
    var rowsCount = Math.max(sheet.getLastRow(), clean.length + 1);
    var data = [headers].concat(clean.map(function (row) { return headers.map(function (h) { return OprosStore.safeCell(row[h]); }); }));
    while (data.length < rowsCount) data.push(headers.map(function () { return ''; }));
    var old = sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), headers.length).getValues();
    while (old.length < rowsCount) old.push(headers.map(function () { return ''; }));
    return { sheet: sheet, clean: clean, data: data, old: old };
  }
  function save(token, payload) {
    requireSession(token);
    if (!payload || typeof payload !== 'object' || JSON.stringify(payload).length > 500000) throw new Error('Настройки слишком большие или повреждены.');
    return withLock(function () {
      requireSession(token);
      var book = OprosStore.workbook(), current = snapshot(book);
      if (current.revision !== payload.revision) throw new Error('Настройки уже изменены в таблице или другом окне. Обновите редактор перед сохранением.');
      var jobs = Object.keys(fields).map(function (key) { return prepare(book, key, payload[key]); });
      // Validate all edits, including hidden/disabled rows becoming enabled, before writing any sheet.
      OprosStore.schemaFromRows(book, jobs[0].clean, jobs[1].clean, jobs[2].clean, current.lists);
      var written = [];
      try {
        jobs.forEach(function (job) {
          if (job.data.length > job.sheet.getMaxRows()) job.sheet.insertRowsAfter(job.sheet.getMaxRows(), job.data.length - job.sheet.getMaxRows());
          written.push(job);
          job.sheet.getRange(1, 1, job.data.length, job.data[0].length).setValues(job.data);
        });
        SpreadsheetApp.flush();
      } catch (error) {
        var rollbackFailed = false;
        written.forEach(function (job) { try { job.sheet.getRange(1, 1, job.old.length, job.old[0].length).setValues(job.old); } catch (ignored) { rollbackFailed = true; } });
        if (rollbackFailed) throw new Error('Сохранение прервано. Проверьте листы Settings, Sections и Questions: восстановление выполнено не полностью.');
        throw new Error('Не удалось сохранить. Прежние настройки восстановлены.');
      }
      return snapshot(book);
    });
  }
  return { hash: hash, login: login, logout: logout, load: load, save: save };
})();

function adminLogin(password) { return OprosAdmin.login(password); }
function adminLogout(token) { return OprosAdmin.logout(token); }
function adminGetConfig(token) { return OprosAdmin.load(token); }
function adminSaveConfig(token, payload) { return OprosAdmin.save(token, payload); }

/** Run manually in the bound spreadsheet editor; web-app calls have no spreadsheet UI. */
function setupAdmin() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Запустите setupAdmin в редакторе Apps Script связанной Google Таблицы.');
  var ui = SpreadsheetApp.getUi();
  if (book.getId() !== OprosStore.workbook().getId()) throw new Error('Откройте редактор нужной таблицы.');
  var answer = ui.alert('Пароль администратора', 'Создать новый пароль? Действующие сессии администратора будут завершены.', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  // Generated high-entropy password: only its SHA-256 digest is persisted.
  var password = 'opros-' + Utilities.getUuid() + Utilities.getUuid();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Сервис занят. Повторите позже.');
  try { PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD_HASH', OprosAdmin.hash(password)); }
  finally { lock.releaseLock(); }
  ui.alert('Сохраните пароль', password + '\n\nСкопируйте в менеджер паролей. Он понадобится для кнопки «Вход администратора» на форме. В таблице и исходниках пароль не сохраняется.', ui.ButtonSet.OK);
}
