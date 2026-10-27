'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const context = vm.createContext({ console });
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Domain.gs'), 'utf8'), context, {
  filename: 'Domain.gs'
});
const domain = context.OprosDomain;
const normalize = value => JSON.parse(JSON.stringify(value));

function question(id, type = 'text', extra = {}) {
  return { id, section_id: 'main', title: id, type, required: false,
    position: 1, enabled: true, ...extra };
}

function schema(questions, extra = {}) {
  return domain.buildSchema(
    extra.settings || [{ key: 'form_title', value: 'Тестовая анкета' }],
    extra.sections || [{ id: 'main', title: 'Основное', position: 1, enabled: true }],
    questions,
    extra.lists || []
  );
}

function payload(answers = {}, extra = {}) {
  return { submissionId: 'test_submission_0001', answers, respondentEmail: '',
    honeypot: '', ...extra };
}

function condition(questionId, operator, value) {
  return JSON.stringify({ question: questionId, operator, value });
}

test('builds a schema from sheet rows and parses TRUE/FALSE rather than JS truthiness', () => {
  const form = schema([
    question('optional', 'text', { required: 'FALSE', enabled: 'TRUE' }),
    question('required', 'text', { required: 'TRUE', position: 2 }),
    question('disabled', 'text', { enabled: 'FALSE', position: 3 })
  ], { settings: [
    { key: 'form_title', value: 'Анкета из таблицы' },
    { key: 'accept_responses', value: 'FALSE' },
    { key: 'collect_email', value: 'TRUE' },
    { key: 'email_required', value: 'FALSE' }
  ] });
  assert.equal(form.title, 'Анкета из таблицы');
  assert.equal(form.acceptResponses, false);
  assert.equal(form.collectEmail, true);
  assert.equal(form.emailRequired, false);
  assert.deepEqual(normalize(form.sections[0].questions.map(q => [q.id, q.required])), [
    ['optional', false], ['required', true]
  ]);
});

test('reads enabled list choices in configured order and preserves values separately from labels', () => {
  const form = schema([question('city', 'select', { options_source: 'list:cities' })], {
    lists: [
      { list_id: 'cities', value: 'spb', label: 'Санкт-Петербург', position: 2, enabled: true },
      { list_id: 'cities', value: 'old', label: 'Неактивный', position: 0, enabled: false },
      { list_id: 'other', value: 'wrong', label: 'Другой список', position: 0, enabled: true },
      { list_id: 'cities', value: 'moscow', label: 'Москва', position: 1, enabled: true }
    ]
  });
  assert.deepEqual(normalize(form.sections[0].questions[0].options), [
    { value: 'moscow', label: 'Москва' },
    { value: 'spb', label: 'Санкт-Петербург' }
  ]);
  assert.equal(domain.validateSubmission(form, payload({ city: 'moscow' })).answers.city, 'moscow');
  assert.throws(() => domain.validateSubmission(form, payload({ city: 'Москва' })));
  assert.throws(() => domain.validateSubmission(form, payload({ city: 'old' })));
});

test('supports inline choices and rejects choices that were not configured', () => {
  const form = schema([
    question('color', 'radio', { options_json: '["Красный","Синий"]' }),
    question('tools', 'multiselect', { position: 2, options_json:
      '[{"value":"js","label":"JavaScript"},{"value":"gs","label":"Apps Script"}]' })
  ]);
  const result = domain.validateSubmission(form, payload({ color: 'Красный', tools: ['gs', 'js'] }));
  assert.equal(result.answers.color, 'Красный');
  assert.deepEqual(normalize(result.answers.tools), ['gs', 'js']);
  assert.throws(() => domain.validateSubmission(form, payload({ tools: ['unlisted'] })));
  assert.throws(() => domain.validateSubmission(form, payload({ tools: 'gs' })));
});

test('rejects duplicate question IDs, malformed JSON, unknown types and invalid patterns at configuration time', () => {
  assert.throws(() => schema([question('same'), question('same', 'text', { position: 2 })]));
  assert.throws(() => schema([question('bad', 'radio', { options_json: '{broken' })]));
  assert.throws(() => schema([question('bad', 'unsupported')]));
  assert.throws(() => schema([question('bad', 'text', { validation_json: '{"pattern":"["}' })]));
});

test('rejects missing and forward visibility references instead of allowing cycles', () => {
  assert.throws(() => schema([question('one', 'text', {
    visible_if_json: condition('missing', 'equals', 'yes')
  })]));
  assert.throws(() => schema([
    question('one', 'text', { visible_if_json: condition('two', 'equals', 'yes') }),
    question('two', 'text', { position: 2 })
  ]));
});

test('hidden required questions do not block submission and injected hidden answers cannot activate a descendant', () => {
  const form = schema([
    question('gate', 'radio', { required: true, options_json: '["yes","no"]' }),
    question('detail', 'text', { required: true, position: 2,
      visible_if_json: condition('gate', 'equals', 'yes') }),
    question('followup', 'text', { required: true, position: 3,
      visible_if_json: condition('detail', 'equals', 'trigger') })
  ]);
  const answers = { gate: 'no', detail: 'trigger', followup: 'injected' };
  assert.deepEqual(normalize(domain.activeQuestions(form, answers).map(q => q.id)), ['gate']);
  assert.deepEqual(normalize(domain.validateSubmission(form, payload(answers)).answers), { gate: 'no' });
  assert.throws(() => domain.validateSubmission(form, payload({ gate: 'yes' })));
  assert.deepEqual(normalize(domain.validateSubmission(form, payload({ gate: 'yes', detail: 'ordinary' })).answers), {
    gate: 'yes', detail: 'ordinary'
  });
});

test('visibility supports includes, not_equals, empty and not_empty', () => {
  const form = schema([
    question('choice', 'multiselect', { options_json: '["a","b"]' }),
    question('text', 'text', { position: 2 }),
    question('included', 'text', { position: 3, visible_if_json: condition('choice', 'includes', 'a') }),
    question('different', 'text', { position: 4, visible_if_json: condition('text', 'not_equals', 'stop') }),
    question('empty', 'text', { position: 5, visible_if_json: condition('text', 'empty') }),
    question('notempty', 'text', { position: 6, visible_if_json: condition('text', 'not_empty') })
  ]);
  assert.deepEqual(normalize(domain.activeQuestions(form, { choice: ['a'], text: '' }).map(q => q.id)),
    ['choice', 'text', 'included', 'empty']);
  assert.deepEqual(normalize(domain.activeQuestions(form, { choice: ['b'], text: 'stop' }).map(q => q.id)),
    ['choice', 'text', 'notempty']);
});

test('validates text length and configured regex on the server', () => {
  const form = schema([question('code', 'text', { required: true,
    validation_json: '{"minLength":3,"maxLength":5,"pattern":"^[A-Z]+$","patternMessage":"Только заглавные буквы"}' })]);
  assert.equal(domain.validateSubmission(form, payload({ code: 'ABC' })).answers.code, 'ABC');
  for (const code of ['', 'AB', 'ABCDEF', 'abc']) {
    assert.throws(() => domain.validateSubmission(form, payload({ code })));
  }
});

test('rejects missing required values but does not treat zero as empty', () => {
  const form = schema([question('count', 'number', { required: true })]);
  assert.equal(domain.validateSubmission(form, payload({ count: 0 })).answers.count, 0);
  assert.throws(() => domain.validateSubmission(form, payload({})));
  assert.throws(() => domain.validateSubmission(form, payload({ count: '' })));
});

test('enforces finite numeric bounds and integer scale boundaries', () => {
  const form = schema([
    question('age', 'number', { validation_json: '{"min":18,"max":100}' }),
    question('rating', 'scale', { position: 2 })
  ]);
  const result = domain.validateSubmission(form, payload({ age: 18, rating: 5 }));
  assert.equal(result.answers.age, 18);
  assert.equal(result.answers.rating, 5);
  for (const age of [17, 101, Infinity, 'NaN']) {
    assert.throws(() => domain.validateSubmission(form, payload({ age })));
  }
  for (const rating of [0, 6, 2.5]) {
    assert.throws(() => domain.validateSubmission(form, payload({ rating })));
  }
});

test('accepts a real leap day and rejects impossible calendar dates', () => {
  const form = schema([question('date', 'date', { required: true })]);
  assert.equal(domain.validateSubmission(form, payload({ date: '2024-02-29' })).answers.date, '2024-02-29');
  for (const date of ['2026-02-29', '2026-02-30', '2026-13-01', '27.10.2026']) {
    assert.throws(() => domain.validateSubmission(form, payload({ date })));
  }
});

test('required consent must be boolean true, not a truthy string', () => {
  const form = schema([question('consent', 'checkbox', { required: true })]);
  assert.equal(domain.validateSubmission(form, payload({ consent: true })).answers.consent, true);
  for (const consent of [false, 'true', 'yes', 1]) {
    assert.throws(() => domain.validateSubmission(form, payload({ consent })));
  }
});

test('requires and validates respondent email only when collection is enabled', () => {
  const form = schema([question('name')], { settings: [
    { key: 'collect_email', value: true },
    { key: 'email_required', value: true }
  ] });
  assert.equal(domain.validateSubmission(form, payload({}, { respondentEmail: 'user@example.org' })).respondentEmail,
    'user@example.org');
  assert.throws(() => domain.validateSubmission(form, payload({})));
  assert.throws(() => domain.validateSubmission(form, payload({}, { respondentEmail: 'not-email' })));
  const privateForm = schema([question('name')], { settings: [{ key: 'collect_email', value: false }] });
  assert.equal(domain.validateSubmission(privateForm, payload({}, { respondentEmail: 'not-email' })).respondentEmail, '');
});

test('email question validation works independently of respondent email collection', () => {
  const form = schema([question('contact', 'email', { required: true })]);
  assert.equal(domain.validateSubmission(form, payload({ contact: 'name@example.com' })).answers.contact,
    'name@example.com');
  assert.throws(() => domain.validateSubmission(form, payload({ contact: 'name at example dot com' })));
});

test('rejects honeypot, malformed submission IDs, oversized payloads and non-object answers', () => {
  const form = schema([question('name')]);
  assert.throws(() => domain.validateSubmission(form, payload({}, { honeypot: 'bot' })));
  for (const submissionId of ['', 'short', 'contains spaces and symbols!', 'a'.repeat(81)]) {
    assert.throws(() => domain.validateSubmission(form, payload({}, { submissionId })));
  }
  assert.throws(() => domain.validateSubmission(form, payload({ name: 'x'.repeat(100001) })));
  assert.throws(() => domain.validateSubmission(form, payload([])));
  assert.throws(() => domain.validateSubmission(form, payload(null)));
});

test('visibility of required follow-up uses canonical numeric answers, preventing string-type bypass', () => {
 const form=schema([question('score','number'),question('followup','text',{position:2,required:true,visible_if_json:condition('score','equals',5)})]);
 assert.throws(()=>domain.validateSubmission(form,payload({score:'5'})));
 assert.equal(domain.validateSubmission(form,payload({score:'5',followup:'details'})).answers.followup,'details');
});

test('map coordinates must be finite numbers within geographic bounds',()=>{
 const form=schema([question('location','map',{required:true})]);
 const point=domain.validateSubmission(form,payload({location:{lat:55.75,lng:37.61,unexpected:'ignored'}})).answers.location;
 assert.deepEqual(normalize(point),{lat:55.75,lng:37.61});
 for(const value of ['',{},'55,37',{lat:'55',lng:37},{lat:NaN,lng:0},{lat:0,lng:Infinity},{lat:-91,lng:0},{lat:0,lng:181}])assert.throws(()=>domain.validateSubmission(form,payload({location:value})));
 assert.equal(domain.validateSubmission(form,payload({location:{lat:0,lng:0}})).answers.location.lat,0);
});
