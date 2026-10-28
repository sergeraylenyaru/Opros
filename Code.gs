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

/** Keyless geocoding. Cache + global throttle respect the public Nominatim service. */
function lookupMapAddress(formId, query) {
  if (!query || typeof query !== 'object') throw new Error('Некорректный запрос адреса.');
  var config = OprosForms.load(OprosStore.workbook(), OprosForms.id(formId));
  if (!config.questions.some(function(q){return q.type==='map' && q.enabled!==false && !/^(false|0|no)$/i.test(String(q.enabled));})) throw new Error('В этой форме нет вопроса с картой.');
  var parameters, mode;
  if (query.text != null) {
    var text=String(query.text).trim();if(text.length<3||text.length>300)throw new Error('Введите адрес от 3 до 300 символов.');
    parameters='search?format=jsonv2&limit=5&q='+encodeURIComponent(text);mode='search';
  } else {
    if(typeof query.lat!=='number'||typeof query.lng!=='number'||!isFinite(query.lat)||!isFinite(query.lng)||Math.abs(query.lat)>90||Math.abs(query.lng)>180)throw new Error('Некорректные координаты.');
    parameters='reverse?format=jsonv2&lat='+query.lat.toFixed(6)+'&lon='+query.lng.toFixed(6);mode='reverse';
  }
  var cache=CacheService.getScriptCache(),key='geocode:'+OprosAdmin.hash(parameters),cached=cache.get(key);
  if(cached)return JSON.parse(cached);
  var lock=LockService.getScriptLock(),geoLocked=false;if(!lock.tryLock(1000))throw new Error('Поиск занят. Повторите через секунду.');geoLocked=true;
  try {
    cached=cache.get(key);if(cached)return JSON.parse(cached);
    var geoProperties=PropertiesService.getScriptProperties();
    if(Date.now()-Number(geoProperties.getProperty('GEOCODER_LAST_REQUEST_AT')||0)<1100)throw new Error('Повторите поиск через секунду.');
    geoProperties.setProperty('GEOCODER_LAST_REQUEST_AT',String(Date.now()));
    lock.releaseLock();geoLocked=false;
    var response=UrlFetchApp.fetch('https://nominatim.openstreetmap.org/'+parameters+'&accept-language=ru',{headers:{'User-Agent':'Opros-AppsScript/1.0 (+https://github.com/sergeraylenyaru/Opros)','Accept':'application/json'},muteHttpExceptions:true});
    if(response.getResponseCode()!==200)throw new Error('Адресный сервис временно недоступен. Можно сохранить координаты без адреса.');
    var raw;try{raw=JSON.parse(response.getContentText());}catch(error){throw new Error('Адресный сервис вернул непонятный ответ. Повторите позже.');}
    var rows=mode==='search'?raw:(raw.error?[]:[raw]);
    if(!Array.isArray(rows))throw new Error('Не удалось прочитать результат поиска.');
    var results=rows.slice(0,5).filter(function(row){return row&&isFinite(Number(row.lat))&&isFinite(Number(row.lon))&&Math.abs(Number(row.lat))<=90&&Math.abs(Number(row.lon))<=180;}).map(function(row){return {lat:Number(row.lat),lng:Number(row.lon),address:String(row.display_name||'').slice(0,1000)};});
    cache.put(key,JSON.stringify(results),3600);return results;
  } finally {if(geoLocked)lock.releaseLock();}
}
