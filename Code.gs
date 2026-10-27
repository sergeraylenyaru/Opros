/** Public web endpoints. Admin handlers end in '_' and cannot be called by HTML clients. */
function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.initialFormId = OprosForms.id(arguments[0] && arguments[0].parameter && arguments[0].parameter.form);
  template.appUrl = ScriptApp.getService().getUrl() || '';
  return template.evaluate()
    .setTitle('Анкета · Opros')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getFormSchema(formId) {
  var book = OprosStore.workbook();
  var schema = OprosForms.schema(book, formId);
  var sheet = OprosStore.responses(OprosForms.responseBook(book, formId));
  if (schema.maxResponses && sheet.getLastRow() - 1 >= schema.maxResponses) schema.acceptResponses = false;
  // A stable non-reversible namespace keeps drafts from different tables separate.
  schema.formId = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, book.getId() + ":" + OprosForms.id(formId))
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('').slice(0, 24);
  delete schema.maxResponses;
  return schema;
}

function submitResponse(payload) {
  var formId = OprosForms.id(payload && payload.formId);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Сервис занят. Повторите отправку через несколько секунд.');
  try {
    var book = OprosStore.workbook();
    var schema = OprosForms.schema(book, formId);
    var result = OprosDomain.validateSubmission(schema, payload);
    var sheet = OprosStore.responses(OprosForms.responseBook(book, formId));
    // A network retry must succeed even if the first write filled the response limit.
    if (OprosStore.findSubmission(sheet, result.submissionId)) {
      return { ok: true, submissionId: result.submissionId, message: schema.confirmationMessage };
    }
    if (!schema.acceptResponses || (schema.maxResponses && sheet.getLastRow() - 1 >= schema.maxResponses)) throw new Error(schema.closedMessage);
    OprosStore.append(sheet, schema, result);
    return { ok: true, submissionId: result.submissionId, message: schema.confirmationMessage };
  } finally { lock.releaseLock(); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Расширенная форма')
    .addItem('Первый запуск / подготовить структуру', 'startApp')
    .addItem('Проверить настройки', 'validateConfiguration_')
    .addItem('Создать / сменить пароль администратора', 'setupAdmin')
    .addSeparator().addItem('Как опубликовать', 'showPublishingGuide_').addToUi();
}

function initializeSpreadsheet_() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Откройте Apps Script через меню «Расширения» нужной Google Таблицы.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Настройка уже выполняется. Повторите позже.');
  try {
    OprosStore.initialize(book);
    OprosForms.registry(book);
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', book.getId());
  } finally { lock.releaseLock(); }
  setupNotice_('Структура готова, привязка сохранена. Запустите startApp для настройки администратора.');
}

function validateConfiguration_() {
  var schema = OprosStore.schema(OprosStore.workbook());
  OprosStore.responses(OprosStore.workbook());
  var count = schema.sections.reduce(function (n, s) { return n + s.questions.length; }, 0);
  SpreadsheetApp.getUi().alert('Настройки корректны: разделов — ' + schema.sections.length + ', вопросов — ' + count + '.');
}

function showPublishingGuide_() {
  SpreadsheetApp.getUi().alert('Публикация', 'Откройте Расширения → Apps Script → Развернуть → Новое развёртывание → Веб-приложение. Выполнять от вашего имени; доступ выберите под свою аудиторию. Скопируйте адрес, заканчивающийся на /exec. После изменения кода обновите версию развёртывания. Данные таблицы обновляются без новой версии.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Editor entry point; initialization checks for a spreadsheet UI context. */
function setupForm() {
  initializeSpreadsheet_();
}

/** Single first-run entry point, visible in Code.gs in the Apps Script editor. */
function startApp() {
  initializeSpreadsheet_();
  var properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('ADMIN_PASSWORD_HASH') || properties.getProperty('ADMIN_SETUP_KEY')) setupAdmin();
  else setupNotice_('Настройка завершена. Пароль администратора сохранён без изменений.');
}

/** Kept in Code.gs so both setup actions are visible in the same editor file. */
function setupAdmin() {
  configureAdmin_();
}

function setupNotice_(message) {
  if (typeof Logger !== 'undefined') Logger.log(message);
}
