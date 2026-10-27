/** Each new form owns exactly two sheets; the shared registry contains no responses. */
var OprosForms = (function () {
  var REGISTRY=['id','title','questions_sheet','responses_sheet'];
  var COLUMNS=['kind','id','section_id','title','description','type','required','options_source','options_json','visible_if_json','validation_json','position','enabled','key','value'];
  function id(value){value=value||'default';if(!/^(default|f_[a-f0-9]{16})$/.test(value))throw new Error('Некорректная ссылка на форму.');return value;}
  function registry(book){var sheet=book.getSheetByName('Forms');if(!sheet){sheet=book.insertSheet('Forms');sheet.getRange(1,1,2,4).setValues([REGISTRY,['default','Основная форма','Questions','Responses']]);sheet.setFrozenRows(1);}return sheet;}
  function rows(book){var sheet=book.getSheetByName('Forms');if(!sheet)return [{id:'default',title:'Основная форма',questions_sheet:'Questions',responses_sheet:'Responses'}];return sheet.getDataRange().getValues().slice(1).filter(function(r){return r[0];}).map(function(r){var result={};REGISTRY.forEach(function(k,i){result[k]=r[i];});return result;});}
  function find(book,value){value=id(value);var entry=rows(book).filter(function(r){return r.id===value;})[0];if(!entry)throw new Error('Форма не найдена.');return entry;}
  function load(book,value){
    var entry=find(book,value),config={formId:entry.id};
    if(entry.id==='default') ['Settings','Sections','Questions','Lists'].forEach(function(name){config[name.toLowerCase()]=OprosStore.readRows(book,name);});
    else{
      config.settings=[];config.sections=[];config.questions=[];config.lists=OprosStore.readRows(book,'Lists');
      var values=book.getSheetByName(entry.questions_sheet).getDataRange().getValues();var headers=values.shift();
      if(headers.join('|')!==COLUMNS.join('|'))throw new Error('Не меняйте заголовки листа '+entry.questions_sheet);
      values.filter(function(row){return row[0];}).forEach(function(row){var item={};headers.forEach(function(k,i){if(k!=='kind')item[k]=row[i];});var kind=row[0];if(['settings','sections','questions'].indexOf(kind)<0)throw new Error('Некорректный тип строки формы.');config[kind].push(item);});
    }
    config.revision=OprosAdmin.hash(JSON.stringify(config));return config;
  }
  function data(config){var result=[COLUMNS];['settings','sections','questions'].forEach(function(kind){config[kind].forEach(function(row){result.push(COLUMNS.map(function(key){return key==='kind'?kind:OprosStore.safeCell(row[key]==null?'':row[key]);}));});});return result;}
  function create(book,title){
    title=String(title||'Новая форма').trim();if(!title||title.length>150)throw new Error('Укажите название до 150 символов.');
    var formId='f_'+Utilities.getUuid().replace(/-/g,'').slice(0,16),entry={id:formId,title:title,questions_sheet:'Вопросы_'+formId.slice(2),responses_sheet:'Ответы_'+formId.slice(2)};
    var config={formId:formId,settings:[{key:'form_title',value:title},{key:'accept_responses',value:false},{key:'confirmation_message',value:'Спасибо! Ответ сохранён.'},{key:'fio_range',value:'People!A2:A'}],sections:[{id:'main',title:'Вопросы',description:'',position:1,enabled:true}],questions:[{id:'question_1',section_id:'main',title:'Ваш комментарий',type:'paragraph',position:1,enabled:true,required:false}],lists:[]};
    var created=[],registrySheet=null,registryRow=0;
    try{var qs=book.insertSheet(entry.questions_sheet);created.push(qs);var matrix=data(config);qs.getRange(1,1,matrix.length,COLUMNS.length).setValues(matrix);qs.setFrozenRows(1);var rs=book.insertSheet(entry.responses_sheet);created.push(rs);rs.getRange(1,1,1,3).setValues([['submission_id','submitted_at','respondent_email']]);rs.setFrozenRows(1);var reg=registry(book);registrySheet=reg;registryRow=reg.getLastRow()+1;reg.getRange(registryRow,1,1,4).setValues([REGISTRY.map(function(k){return OprosStore.safeCell(entry[k]);})]);SpreadsheetApp.flush();return load(book,formId);}catch(error){if(registrySheet&&registryRow)registrySheet.getRange(registryRow,1,1,4).setValues([['','','','']]);created.forEach(function(sheet){book.deleteSheet(sheet);});throw error;}
  }
  function save(book,payload){
    var current=load(book,payload.formId);if(current.revision!==payload.revision)throw new Error('Настройки уже изменены. Обновите редактор.');
    ['settings','sections','questions'].forEach(function(key){if(!Array.isArray(payload[key])||!payload[key].length||payload[key].length>200)throw new Error('Некорректные строки '+key);payload[key].forEach(function(row){if(!row||typeof row!=='object'||Array.isArray(row))throw new Error('Некорректная строка');Object.keys(row).forEach(function(k){if(COLUMNS.indexOf(k)<0)delete row[k];else if(row[k]!=null&&['string','number','boolean'].indexOf(typeof row[k])<0)throw new Error('Некорректное значение');});});});
    OprosStore.schemaFromRows(book,payload.settings,payload.sections,payload.questions,current.lists);
    var entry=find(book,payload.formId),sheet=book.getSheetByName(entry.questions_sheet),matrix=data(payload),old=sheet.getDataRange().getValues();while(matrix.length<old.length)matrix.push(COLUMNS.map(function(){return '';}));
    if(matrix.length>sheet.getMaxRows())sheet.insertRowsAfter(sheet.getMaxRows(),matrix.length-sheet.getMaxRows());
    try{sheet.getRange(1,1,matrix.length,COLUMNS.length).setValues(matrix);SpreadsheetApp.flush();}catch(error){while(old.length<matrix.length)old.push(COLUMNS.map(function(){return '';}));sheet.getRange(1,1,old.length,COLUMNS.length).setValues(old);throw new Error('Сохранение не удалось; настройки восстановлены.');}
    return load(book,payload.formId);
  }
  function schema(book,value){if(id(value)==='default')return OprosStore.schema(book);var config=load(book,value);return OprosStore.schemaFromRows(book,config.settings,config.sections,config.questions,config.lists);}
  function responseBook(book,value){var entry=find(book,value);return {getSheetByName:function(name){return book.getSheetByName(name==='Responses'?entry.responses_sheet:name);}};}
  return {id:id,registry:registry,rows:rows,load:load,create:create,save:save,schema:schema,responseBook:responseBook};
})();
function adminListForms(token){OprosAdmin.requireSession(token);var book=OprosStore.workbook();return OprosForms.rows(book).map(function(entry){var config=OprosForms.load(book,entry.id);var title=config.settings.filter(function(r){return r.key==='form_title';})[0];var active=config.settings.filter(function(r){return r.key==='accept_responses';})[0];return {id:entry.id,title:title?String(title.value):entry.title,acceptResponses:!active||/^(true|1)$/i.test(String(active.value)),questions:config.questions.length,responses:Math.max(0,book.getSheetByName(entry.responses_sheet).getLastRow()-1)};});}
function adminCreateForm(token,title){OprosAdmin.requireSession(token);var lock=LockService.getScriptLock();if(!lock.tryLock(20000))throw new Error('Сервис занят');try{OprosAdmin.requireSession(token);return OprosForms.create(OprosStore.workbook(),title);}finally{lock.releaseLock();}}
function adminPreviewSource(token,source){OprosAdmin.requireSession(token);return OprosStore.externalOptions(source);}
