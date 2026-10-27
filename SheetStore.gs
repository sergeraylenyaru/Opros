/** Spreadsheet persistence. Administration is deliberately not exposed to HTML clients. */
var OprosStore = (function () {
  'use strict';
  var HEADERS = {
    Settings: ['key', 'value', 'description'],
    Sections: ['id', 'title', 'description', 'position', 'enabled'],
    Questions: ['id', 'section_id', 'title', 'description', 'type', 'required', 'options_source', 'options_json', 'visible_if_json', 'validation_json', 'position', 'enabled'],
    People: ['ФИО'],
    Lists: ['list_id', 'value', 'label', 'position', 'enabled'],
    Responses: ['submission_id', 'submitted_at', 'respondent_email']
  };
  function workbook() {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('Владелец должен запустить setupForm из редактора Apps Script.');
    return SpreadsheetApp.openById(id);
  }
  function readRows(book, name) {
    var sheet = book.getSheetByName(name);
    if (!sheet || !sheet.getLastRow()) throw new Error('Не настроен лист ' + name + '. Запустите инициализацию.');
    var values = sheet.getDataRange().getValues();
    var headers = values.shift().map(function (x) { return String(x).trim(); });
    HEADERS[name].forEach(function (key) { if (headers.indexOf(key) < 0) throw new Error('На листе ' + name + ' нет столбца ' + key); });
    if (new Set(headers.filter(Boolean)).size !== headers.filter(Boolean).length) throw new Error('Повторные заголовки на листе ' + name);
    return values.filter(function (row) { return row.some(function (x) { return x !== ''; }); }).map(function (row) {
      var item = Object.create(null);
      headers.forEach(function (key, i) { if (key) item[key] = row[i]; });
      return item;
    });
  }
  function schema(book) {
    return schemaFromRows(book, readRows(book, 'Settings'), readRows(book, 'Sections'), readRows(book, 'Questions'), readRows(book, 'Lists'));
  }
  function schemaFromRows(book, settings, sections, questions, lists) {
    var resolved = questions.map(function(row){var copy=Object.assign({},row);if(copy.enabled!==false&&!/^(false|0|no)$/i.test(String(copy.enabled))&&String(copy.options_source||'').indexOf('sheet:')===0){copy.options_json=JSON.stringify(externalOptions(JSON.parse(copy.options_source.slice(6))));copy.options_source='';}return copy;});
    var result = OprosDomain.buildSchema(settings, sections, resolved, lists);
    if (result.sections.some(function (s) { return s.questions.some(function (q) { return q.id === 'respondent_fio'; }); })) throw new Error('ID respondent_fio зарезервирован для обязательного ФИО.');
    var range = settings.filter(function (r) { return r.key === 'fio_range'; })[0];
    result.sections[0].questions.unshift({ id: 'respondent_fio', title: 'ФИО', description: 'Выберите себя в списке. Для поиска введите часть ФИО.', type: 'person', required: true, options: people(book, range ? range.value : 'People!A2:A'), visibleIf: null, validation: {} });
    return result;
  }
  function people(book, reference) {
    var match = /^('?)(.+?)\1!([A-Za-z]+)([1-9][0-9]*):([A-Za-z]+)([1-9][0-9]*)?$/.exec(String(reference).trim());
    if (!match || match[3].toUpperCase() !== match[5].toUpperCase()) throw new Error('fio_range: укажите один столбец, например People!A2:A или Сотрудники!B2:B500.');
    var sheet = book.getSheetByName(match[2].replace(/''/g, "'"));
    if (!sheet) {
      if (String(reference) === 'People!A2:A') return [];
      throw new Error('Не найден лист ФИО: ' + match[2]);
    }
    var first = Number(match[4]), last = Math.min(match[6] ? Number(match[6]) : sheet.getLastRow(), sheet.getLastRow());
    if (last < first) return [];
    if (last - first + 1 > 10000) throw new Error('Диапазон ФИО не должен превышать 10000 строк.');
    var column = match[3].toUpperCase().split('').reduce(function (n, c) { return n * 26 + c.charCodeAt(0) - 64; }, 0);
    var seen = Object.create(null);
    return sheet.getRange(first, column, last - first + 1, 1).getDisplayValues().map(function (row) { return String(row[0]).trim(); }).filter(function (name) {
      if (!name || seen[name]) return false;
      if (name.length > 500) throw new Error('ФИО длиннее 500 символов.');
      seen[name] = true; return true;
    }).map(function (name) { return { value: name, label: name }; });
  }
  function externalOptions(source) {
    if (!source || typeof source!=='object') throw new Error('Укажите файл, лист и столбец справочника.');
    var file=String(source.fileId||'').trim(),url=/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/.exec(file);if(url)file=url[1];
    if(!/^[A-Za-z0-9_-]{10,200}$/.test(file)||!String(source.sheet||'').trim()||!/^[A-Za-z]{1,3}$/.test(String(source.column||'')))throw new Error('Неверный файл, лист или столбец справочника.');
    var start=Number(source.startRow||2);if(!Number.isInteger(start)||start<1)throw new Error('Начальная строка должна быть целой от 1.');
    var book;try{book=SpreadsheetApp.openById(file);}catch(error){throw new Error('Нет доступа к файлу справочника. Предоставьте доступ владельцу приложения.');}
    var sheet=book.getSheetByName(String(source.sheet).trim());if(!sheet)throw new Error('Не найден лист справочника: '+source.sheet);
    var last=sheet.getLastRow();if(last<start)return [];if(last-start+1>10000)throw new Error('Справочник должен содержать не более 10000 строк.');
    var column=String(source.column).toUpperCase().split('').reduce(function(n,c){return n*26+c.charCodeAt(0)-64;},0),seen=Object.create(null),result=[];
    sheet.getRange(start,column,last-start+1,1).getDisplayValues().forEach(function(row){var value=String(row[0]).trim();if(!value||seen[value])return;if(value.length>500)throw new Error('Слишком длинное значение справочника.');seen[value]=true;result.push({value:value,label:value});});
    if(result.length>1000)throw new Error('Не более 1000 уникальных вариантов на вопрос.');return result;
  }
  function safeCell(value) {
    // Force user-controlled formula-like strings to remain text in Google Sheets.
    return typeof value === 'string' && /^[\s\u0000-\u001f]*[=+\-@]/.test(value) ? "'" + value : value;
  }
  function writeRows(sheet, headers, rows) {
    if (!rows.length) return;
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows.map(function (r) { return headers.map(function (h) { return safeCell(r[h] == null ? '' : r[h]); }); }));
  }
  function initialize(book) {
    var seeds = {
      Settings: [
        { key: 'form_title', value: 'Давайте познакомимся', description: 'Название формы' },
        { key: 'form_description', value: 'Короткая анкета с вопросами, которые зависят от ваших ответов.', description: 'Описание' },
        { key: 'confirmation_message', value: 'Спасибо! Ваш ответ сохранён.', description: 'Текст после отправки' },
        { key: 'accept_responses', value: true, description: 'TRUE — приём открыт, FALSE — закрыт' },
        { key: 'collect_email', value: true, description: 'Показывать поле электронной почты' },
        { key: 'email_required', value: false, description: 'Требовать заполнение почты; адрес указывает сам участник' },
        { key: 'submit_label', value: 'Отправить ответ', description: 'Текст кнопки' },
        { key: 'closed_message', value: 'Приём ответов завершён. Спасибо за интерес!', description: 'Сообщение закрытой формы' },
        { key: 'fio_range', value: 'People!A2:A', description: 'Диапазон ФИО на отдельном листе (один столбец)' },
        { key: 'max_responses', value: 0, description: 'Максимум ответов, 0 — без ограничения' }
      ],
      Sections: [
        { id: 'about', title: 'О вас', description: 'Начнём с самого простого', position: 1, enabled: true },
        { id: 'details', title: 'Ваши предпочтения', description: 'Помогите нам подготовить подходящее предложение', position: 2, enabled: true }
      ],
      Questions: [
        { id: 'city', section_id: 'about', title: 'Ваш город', type: 'select', required: true, options_source: 'list:cities', position: 2, enabled: true },
        { id: 'role', section_id: 'about', title: 'Как вы планируете использовать инструмент?', type: 'radio', required: true, options_json: '[{"value":"personal","label":"Для себя"},{"value":"team","label":"Для команды"}]', position: 3, enabled: true },
        { id: 'team_size', section_id: 'about', title: 'Сколько человек в команде?', type: 'number', required: true, visible_if_json: '{"question":"role","operator":"equals","value":"team"}', validation_json: '{"min":2,"max":10000}', position: 4, enabled: true },
        { id: 'interests', section_id: 'details', title: 'Какие возможности вам интересны?', description: 'Можно выбрать до трёх вариантов', type: 'multiselect', required: true, options_source: 'list:features', validation_json: '{"min":1,"max":3}', position: 5, enabled: true },
        { id: 'rating', section_id: 'details', title: 'Насколько полезен такой инструмент?', description: '1 — пока не уверен, 5 — очень полезен', type: 'scale', required: true, validation_json: '{"min":1,"max":5}', position: 6, enabled: true },
        { id: 'comment', section_id: 'details', title: 'Что ещё стоит учесть?', type: 'paragraph', required: false, validation_json: '{"maxLength":2000}', position: 7, enabled: true },
        { id: 'consent', section_id: 'details', title: 'Согласен передать указанные сведения владельцу анкеты', description: 'Перед использованием замените этот пример своим текстом и укажите условия обработки данных.', type: 'checkbox', required: true, position: 8, enabled: true }
      ],
      Lists: [
        { list_id: 'cities', value: 'moscow', label: 'Москва', position: 1, enabled: true },
        { list_id: 'cities', value: 'spb', label: 'Санкт-Петербург', position: 2, enabled: true },
        { list_id: 'cities', value: 'other', label: 'Другой город', position: 3, enabled: true },
        { list_id: 'features', value: 'conditions', label: 'Условные вопросы', position: 1, enabled: true },
        { list_id: 'features', value: 'lists', label: 'Списки из таблицы', position: 2, enabled: true },
        { list_id: 'features', value: 'validation', label: 'Проверка ответов', position: 3, enabled: true },
        { list_id: 'features', value: 'sections', label: 'Разделы и шаги', position: 4, enabled: true }
      ],
      People: [],
      Responses: []
    };
    Object.keys(HEADERS).forEach(function (name) {
      var sheet = book.getSheetByName(name) || book.insertSheet(name);
      if (!sheet.getLastRow()) {
        var headers = HEADERS[name];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8eefb');
        writeRows(sheet, headers, seeds[name]);
        sheet.autoResizeColumns(1, headers.length);
      }
    });
    // Add the new setting without changing existing settings or responses.
    if (!readRows(book, 'Settings').some(function (r) { return r.key === 'fio_range'; })) {
      var settingsSheet = book.getSheetByName('Settings');
      var settingHeaders = settingsSheet.getRange(1, 1, 1, settingsSheet.getLastColumn()).getValues()[0];
      var settingRow = { key: 'fio_range', value: 'People!A2:A', description: 'Диапазон ФИО на отдельном листе' };
      ensureCapacity(settingsSheet, settingsSheet.getLastRow() + 1, settingHeaders.length);
      settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, 1, settingHeaders.length).setValues([settingHeaders.map(function (key) { return settingRow[key] || ''; })]);
    }
    // Existing question rows and responses are never reseeded or overwritten.
    schema(book);
    return book;
  }
  function responses(book) {
    var sheet = book.getSheetByName('Responses');
    if (!sheet || sheet.getLastRow() < 1) throw new Error('Лист Responses не настроен. Запустите инициализацию.');
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    HEADERS.Responses.forEach(function (key, i) { if (headers[i] !== key) throw new Error('Не меняйте первые три столбца листа Responses.'); });
    if (new Set(headers).size !== headers.length) throw new Error('Повторные заголовки Responses.');
    return sheet;
  }
  function findSubmission(sheet, id) {
    if (sheet.getLastRow() < 2) return false;
    return !!sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).matchCase(true).useRegularExpression(false).findNext();
  }
  function ensureCapacity(sheet, rows, columns) {
    if (columns > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), columns - sheet.getMaxColumns());
    if (rows > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  }
  function append(sheet, schemaValue, submission) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columns = headers.slice();
    schemaValue.sections.forEach(function (s) { s.questions.forEach(function (q) { if (columns.indexOf('q:' + q.id) < 0) columns.push('q:' + q.id); }); });
    ensureCapacity(sheet, sheet.getLastRow() + 1, columns.length);
    if (columns.length !== headers.length) sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
    var row = columns.map(function (key) {
      if (key === 'submission_id') return submission.submissionId;
      if (key === 'submitted_at') return new Date();
      if (key === 'respondent_email') return safeCell(submission.respondentEmail);
      var id = key.indexOf('q:') === 0 ? key.slice(2) : '';
      var value = Object.prototype.hasOwnProperty.call(submission.answers, id) ? submission.answers[id] : '';
      return safeCell(value !== null && typeof value === 'object' ? JSON.stringify(value) : value);
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
  }
  return { externalOptions: externalOptions, readRows: readRows, schemaFromRows: schemaFromRows, workbook: workbook, schema: schema, initialize: initialize, responses: responses, findSubmission: findSubmission, append: append, safeCell: safeCell };
})();
