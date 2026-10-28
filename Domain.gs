/** Pure schema and submission rules, also exercised locally with Node. */
var OprosDomain = (function () {
  'use strict';
  var TYPES = ['text', 'paragraph', 'email', 'number', 'date', 'select', 'multiselect', 'radio', 'checkbox', 'scale', 'person', 'map'];
  var CHOICES = ['select', 'multiselect', 'radio', 'person'];
  var OPS = ['equals', 'not_equals', 'includes', 'empty', 'not_empty'];
  function fail(message) { throw new Error(message); }
  function string(value) { return value == null ? '' : String(value); }
  function bool(value, fallback) {
    if (value === '' || value == null) return fallback;
    if (value === true || value === 1 || /^(true|yes|1)$/i.test(string(value).trim())) return true;
    if (value === false || value === 0 || /^(false|no|0)$/i.test(string(value).trim())) return false;
    return fail('Ожидается TRUE или FALSE: ' + string(value));
  }
  function json(value, fallback, context) {
    if (value === '' || value == null) return fallback;
    try { return JSON.parse(string(value)); } catch (e) { return fail('Некорректный JSON: ' + context); }
  }
  function object(value) { return value !== null && typeof value === 'object' && !Array.isArray(value); }
  function own(value, key) { return Object.prototype.hasOwnProperty.call(value, key); }
  function identifier(value, label) {
    var id = string(value).trim();
    if (!/^[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(id)) fail('Некорректный ID (' + label + '): ' + id);
    if (['constructor', 'prototype', '__proto__'].indexOf(id) >= 0) fail('Зарезервированный ID: ' + id);
    return id;
  }
  function ordered(rows) {
    return rows.filter(function (row) { return bool(row.enabled, true); }).slice().sort(function (a, b) {
      return (Number(a.position) || 0) - (Number(b.position) || 0);
    });
  }
  function normalizeOptions(raw, context) {
    if (!Array.isArray(raw) || raw.length > 1000) fail('Ожидается список до 1000 вариантов: ' + context);
    var seen = Object.create(null);
    return raw.map(function (item) {
      var value = object(item) ? string(item.value) : string(item);
      var label = object(item) ? string(item.label || item.value) : value;
      if (!value || value.length > 500 || seen[value]) fail('Пустой, слишком длинный или повторный вариант: ' + context);
      seen[value] = true;
      return { value: value, label: label };
    });
  }
  function buildSchema(settingsRows, sectionRows, questionRows, listRows) {
    var settings = Object.create(null);
    settingsRows.forEach(function (r) { if (string(r.key).trim()) settings[string(r.key).trim()] = r.value; });
    var schema = {
      title: string(settings.form_title || 'Расширенная анкета'),
      description: string(settings.form_description),
      confirmationMessage: string(settings.confirmation_message || 'Спасибо! Ваш ответ сохранён.'),
      acceptResponses: bool(settings.accept_responses, true),
      closedMessage: string(settings.closed_message || 'Приём ответов завершён.'),
      collectEmail: bool(settings.collect_email, false),
      emailRequired: bool(settings.email_required, false),
      submitLabel: string(settings.submit_label || 'Отправить ответ'),
      maxResponses: Number(settings.max_responses || 0),
      mapProvider: settings.map_provider === 'yandex' ? 'yandex' : 'osm',
      yandexMapsKey: settings.map_provider === 'yandex' ? string(settings.yandex_maps_key).trim() : '',
      sections: []
    };
    if (!Number.isInteger(schema.maxResponses) || schema.maxResponses < 0) fail('max_responses должен быть целым числом от 0.');
    if (schema.emailRequired && !schema.collectEmail) fail('Для email_required включите collect_email.');
    var sections = Object.create(null);
    ordered(sectionRows).forEach(function (r) {
      var id = identifier(r.id, 'раздел');
      if (sections[id]) fail('Повторный ID раздела: ' + id);
      var section = { id: id, title: string(r.title || id), description: string(r.description), questions: [] };
      sections[id] = section;
      schema.sections.push(section);
    });
    if (!schema.sections.length) fail('Нужен хотя бы один включённый раздел.');
    var seen = Object.create(null);
    ordered(questionRows).forEach(function (r) {
      var sectionId = string(r.section_id).trim();
      // A disabled section deliberately hides all of its questions.
      if (!sections[sectionId]) {
        if (sectionRows.some(function (s) { return string(s.id).trim() === sectionId && !bool(s.enabled, true); })) return;
        fail('Не найден раздел вопроса: ' + sectionId);
      }
      var id = identifier(r.id, 'вопрос');
      if (seen[id]) fail('Повторный ID вопроса: ' + id);
      seen[id] = true;
      var type = string(r.type).trim() || 'text';
      if (TYPES.indexOf(type) < 0) fail('Неизвестный тип вопроса: ' + type);
      var validation = json(r.validation_json, {}, id);
      if (!object(validation)) fail('validation_json должен быть объектом: ' + id);
      ['min', 'max', 'minLength', 'maxLength'].forEach(function (k) {
        if (own(validation, k) && (typeof validation[k] !== 'number' || !isFinite(validation[k]))) fail('Некорректное ограничение ' + k + ': ' + id);
      });
      ['minLength', 'maxLength'].forEach(function (k) {
        if (own(validation, k) && (!Number.isInteger(validation[k]) || validation[k] < 0 || validation[k] > 10000)) fail('Некорректная длина: ' + id);
      });
      if (validation.min != null && validation.max != null && validation.min > validation.max) fail('min больше max: ' + id);
      if (validation.minLength != null && validation.maxLength != null && validation.minLength > validation.maxLength) fail('minLength больше maxLength: ' + id);
      if (validation.pattern != null) {
        if (typeof validation.pattern !== 'string' || validation.pattern.length > 300) fail('Некорректный pattern: ' + id);
        try { new RegExp(validation.pattern); } catch (e) { fail('Некорректное регулярное выражение: ' + id); }
      }
      if (type === 'scale') {
        if (validation.min == null) validation.min = 1;
        if (validation.max == null) validation.max = 5;
        if (!Number.isInteger(validation.min) || !Number.isInteger(validation.max) || validation.min > validation.max || validation.max - validation.min > 20) fail('Шкала: нужны целые границы и до 21 значения: ' + id);
      }
      var options = [];
      if (CHOICES.indexOf(type) >= 0) {
        var source = string(r.options_source).trim();
        if (source) {
          if (source.indexOf('list:') !== 0 || !source.slice(5)) fail('options_source: ожидается list:ID: ' + id);
          options = normalizeOptions(ordered(listRows.filter(function (l) { return string(l.list_id).trim() === source.slice(5); })).map(function (l) { return { value: l.value, label: l.label }; }), id);
        } else options = normalizeOptions(json(r.options_json, [], id), id);
        if (!options.length) fail('Не найдены варианты ответа: ' + id);
      }
      var rule = json(r.visible_if_json, null, id);
      if (rule !== null && (!object(rule) || OPS.indexOf(rule.operator) < 0 || typeof rule.question !== 'string')) fail('Некорректное условие показа: ' + id);
      sections[sectionId].questions.push({
        id: id, title: string(r.title || id), description: string(r.description), type: type,
        required: bool(r.required, false), options: options, visibleIf: rule, validation: validation
      });
    });
    var earlier = Object.create(null), count = 0;
    schema.sections.forEach(function (s) { s.questions.forEach(function (q) {
      if (q.visibleIf && !earlier[q.visibleIf.question]) fail('Условие должно ссылаться на предыдущий вопрос: ' + q.id);
      earlier[q.id] = true; count++;
    }); });
    if (!count || count > 100) fail('Форма должна содержать от 1 до 100 вопросов.');
    return schema;
  }
  function empty(value) { return value == null || value === '' || value === false || (Array.isArray(value) && !value.length); }
  function visible(rule, answers) {
    if (!rule) return true;
    var value = own(answers, rule.question) ? answers[rule.question] : undefined;
    // An unanswered or hidden parent cannot accidentally enable not_equals.
    if (rule.operator === 'empty') return empty(value);
    if (rule.operator === 'not_empty') return !empty(value);
    if (empty(value)) return false;
    if (rule.operator === 'includes') return Array.isArray(value) ? value.indexOf(rule.value) >= 0 : String(value).indexOf(String(rule.value)) >= 0;
    if (rule.operator === 'equals') return value === rule.value;
    if (rule.operator === 'not_equals') return value !== rule.value;
    return false;
  }
  function activeQuestions(schema, answers) {
    var effective = Object.create(null), result = [];
    schema.sections.forEach(function (s) { s.questions.forEach(function (q) {
      if (visible(q.visibleIf, effective)) {
        result.push(q);
        if (own(answers, q.id)) effective[q.id] = answers[q.id];
      }
    }); });
    return result;
  }
  function validEmail(value) { return value.length <= 254 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value); }
  function validateSubmission(schema, payload) {
    if (!object(payload) || JSON.stringify(payload).length > 100000) fail('Некорректный или слишком большой ответ.');
    if (payload.honeypot) fail('Не удалось отправить ответ.');
    if (!/^[A-Za-z0-9_-]{16,80}$/.test(string(payload.submissionId))) fail('Некорректный идентификатор отправки. Обновите страницу.');
    if (!object(payload.answers)) fail('Ответы должны быть объектом.');
    var email = schema.collectEmail ? string(payload.respondentEmail).trim() : '';
    if (schema.collectEmail && ((schema.emailRequired && !email) || (email && !validEmail(email)))) fail('Укажите корректный адрес электронной почты.');
    var answers = Object.create(null);
    schema.sections.reduce(function (all, section) { return all.concat(section.questions); }, []).forEach(function (q) {
      if (!visible(q.visibleIf, answers)) return;
      var value = own(payload.answers, q.id) ? payload.answers[q.id] : undefined;
      if (value == null || value === '' || (Array.isArray(value) && !value.length) || (q.type === 'checkbox' && value === false)) {
        if (q.required) fail('Заполните обязательное поле: ' + q.title);
        return;
      }
      var v = q.validation || {};
      if (q.type === 'checkbox') {
        if (typeof value !== 'boolean') fail('Ожидается отметка: ' + q.title);
      } else if (q.type === 'map') {
        if (!object(value) || typeof value.lat !== 'number' || typeof value.lng !== 'number' || !isFinite(value.lat) || !isFinite(value.lng) || Math.abs(value.lat) > 90 || Math.abs(value.lng) > 180) fail('Выберите корректную точку на карте: ' + q.title);
        var address = value.address;
        if (address != null && (typeof address !== 'string' || address.length > 1000)) fail('Некорректный адрес точки: ' + q.title);
        value = { lat: value.lat, lng: value.lng };
        if (typeof address === 'string' && address.trim()) value.address = address.trim();
      } else if (q.type === 'number' || q.type === 'scale') {
        if ((typeof value !== 'number' && typeof value !== 'string') || (typeof value === 'string' && !value.trim())) fail('Ожидается число: ' + q.title);
        value = Number(value);
        if (!isFinite(value) || (q.type === 'scale' && !Number.isInteger(value))) fail('Некорректное число: ' + q.title);
        if ((v.min != null && value < v.min) || (v.max != null && value > v.max)) fail('Число вне допустимого диапазона: ' + q.title);
      } else if (q.type === 'multiselect') {
        if (!Array.isArray(value) || value.some(function (a) { return typeof a !== 'string' || !q.options.some(function (o) { return o.value === a; }); })) fail('Выберите варианты из списка: ' + q.title);
        value = value.filter(function (item, index, all) { return all.indexOf(item) === index; });
        if ((v.min != null && value.length < v.min) || (v.max != null && value.length > v.max)) fail('Недопустимое число выбранных вариантов: ' + q.title);
      } else {
        if (typeof value !== 'string' || value.length > 10000) fail('Ожидается текст до 10000 символов: ' + q.title);
        value = value.trim();
        if (!value) { if (q.required) fail('Заполните обязательное поле: ' + q.title); return; }
        if (q.type === 'email' && !validEmail(value)) fail('Некорректная почта: ' + q.title);
        if (q.type === 'date') {
          if (!/^\d{4}-\d{2}-\d{2}$/.test(value) || !isFinite(Date.parse(value)) || new Date(value).toISOString().slice(0, 10) !== value) fail('Некорректная дата: ' + q.title);
        }
        if (q.type === 'select' || q.type === 'radio' || q.type === 'person') {
          if (!q.options.some(function (o) { return o.value === value; })) fail('Выберите вариант из списка: ' + q.title);
        }
        if ((v.minLength != null && value.length < v.minLength) || (v.maxLength != null && value.length > v.maxLength)) fail('Недопустимая длина ответа: ' + q.title);
        if (v.pattern && !new RegExp(v.pattern).test(value)) fail(string(v.patternMessage || ('Ответ не соответствует формату: ' + q.title)));
      }
      answers[q.id] = value;
    });
    return { submissionId: string(payload.submissionId), respondentEmail: email, answers: answers };
  }
  return { buildSchema: buildSchema, validateSubmission: validateSubmission, activeQuestions: activeQuestions, visible: visible };
})();
