const {test}=require('node:test');
const assert=require('node:assert/strict');
const {runtime}=require('./fakes');
function payload(id='submission_1234567890'){return {submissionId:id,respondentEmail:'test@example.org',answers:{respondent_fio:'Иванов Иван Иванович',name:'Иван',city:'moscow',role:'personal',interests:['lists'],rating:5,consent:true,comment:'=IMPORTXML("evil")'},honeypot:''};}
function setting(r,key,value){const row=r.sheets.Settings.data.find(row=>row[0]===key);row[1]=value;}
test('initialize, load, submit, deduplicate, and preserve configured data on reinitialization',()=>{
 const r=runtime(), c=r.context;c.initializeSpreadsheet_();
 assert.equal(Object.keys(r.sheets).length,6);
 const schema=c.getFormSchema();assert.equal(schema.sections.length,2);assert.equal(schema.formId.length,24);assert.ok(!JSON.stringify(schema).includes(r.book.getId()));
 assert.equal(c.submitResponse(payload()).ok,true);assert.equal(r.sheets.Responses.getLastRow(),2);
 const headers=r.sheets.Responses.data[0], row=r.sheets.Responses.data[1];
 assert.equal(row[headers.indexOf('q:comment')],"'=IMPORTXML(\"evil\")");
 assert.equal(row[headers.indexOf('q:team_size')],'');assert.equal(row[headers.indexOf('q:interests')],'["lists"]');
 assert.equal(c.submitResponse(payload()).ok,true);assert.equal(r.sheets.Responses.getLastRow(),2);
 setting(r,'form_title','Свой опрос');c.initializeSpreadsheet_();assert.equal(c.getFormSchema().title,'Свой опрос');assert.equal(r.sheets.Responses.getLastRow(),2);assert.equal(r.locked,false);
});
test('limits and closure enforced inside lock while a saved retry remains idempotent',()=>{
 const r=runtime(),c=r.context;c.initializeSpreadsheet_();setting(r,'max_responses',1);c.submitResponse(payload());
 assert.equal(c.getFormSchema().acceptResponses,false);assert.throws(()=>c.submitResponse(payload('submission_2222222222')));assert.equal(c.submitResponse(payload()).ok,true);assert.equal(r.locked,false);
 setting(r,'max_responses',0);setting(r,'accept_responses',false);assert.throws(()=>c.submitResponse(payload('submission_3333333333')));assert.equal(r.sheets.Responses.getLastRow(),2);
});
test('server rejects invalid visible required fields, retains lock health, and accepts corrected retry',()=>{
 const r=runtime(),c=r.context;c.initializeSpreadsheet_();const p=payload();p.answers.role='team';assert.throws(()=>c.submitResponse(p));assert.equal(r.locked,false);assert.equal(r.sheets.Responses.getLastRow(),1);
 p.answers.team_size=3;assert.equal(c.submitResponse(p).ok,true);
});
test('response columns expand without moving historic answer cells',()=>{
 const r=runtime(),c=r.context;c.initializeSpreadsheet_();c.submitResponse(payload());const prior=r.sheets.Responses.data[1].slice();
 const h=r.sheets.Questions.data[0], data={id:'extra',section_id:'details',title:'Дополнение',type:'text',enabled:true,position:99};r.sheets.Questions.data.push(h.map(k=>data[k]??''));
 const p=payload('submission_4444444444');p.answers.extra='Новое';c.submitResponse(p);assert.deepEqual(r.sheets.Responses.data[1],prior);assert.equal(r.sheets.Responses.data[2].at(-1),'Новое');
});
test('formula-looking text is escaped; actual numeric values are retained',()=>{const {context:c}=runtime();for(const input of ['=1+1','+cmd','-cmd','@SUM(1)',' \t=1'])assert.equal(c.OprosStore.safeCell(input),"'"+input);assert.equal(c.OprosStore.safeCell(-2),-2);});
test('response grid expands beyond Sheets default column count',()=>{
 const r=runtime(),c=r.context;c.initializeSpreadsheet_();const headers=r.sheets.Questions.data[0];
 for(let n=0;n<30;n++){const data={id:'extra_'+n,section_id:'details',title:'Extra '+n,type:'text',position:100+n,enabled:true};r.sheets.Questions.data.push(headers.map(k=>data[k]??''));}
 c.submitResponse(payload());assert.ok(r.sheets.Responses.getMaxColumns()>26);assert.equal(r.sheets.Responses.data[0].length,41);
});
test('FIO is always required and first, source range follows settings, invalid names rejected',()=>{
 const r=runtime(),c=r.context;c.setupForm();
 const first=c.getFormSchema().sections[0].questions[0];assert.equal(first.id,'respondent_fio');assert.equal(first.required,true);assert.equal(first.type,'person');assert.equal(first.options.length,2);
 const p=payload();delete p.answers.respondent_fio;assert.throws(()=>c.submitResponse(p));p.answers.respondent_fio='Посторонний';assert.throws(()=>c.submitResponse(p));
 const custom=r.book.insertSheet('Сотрудники');custom.data=[['ID','ФИО'],['1','Сидоров Сидор'],['2',''],['3','Сидоров Сидор'],['4','Другой сотрудник']];setting(r,'fio_range',"'Сотрудники'!B2:B4");
 const options=c.getFormSchema().sections[0].questions[0].options;assert.equal(options.length,1);assert.equal(options[0].value,'Сидоров Сидор');p.answers.respondent_fio='Сидоров Сидор';c.submitResponse(p);
 const h=r.sheets.Responses.data[0];assert.equal(r.sheets.Responses.data[1][h.indexOf('q:respondent_fio')],'Сидоров Сидор');
});
test('initialization creates empty People without invented names and preserves existing answers',()=>{
 const r=runtime();delete r.sheets.People;r.context.setupForm();assert.equal(r.sheets.People.data.length,1);assert.equal(r.context.getFormSchema().sections[0].questions[0].options.length,0);
});
test('map point persisted as JSON and rejected before writing when malformed',()=>{
 const r=runtime(),c=r.context;c.setupForm();const headers=r.sheets.Questions.data[0];const q={id:'place',section_id:'details',title:'Место',type:'map',required:true,enabled:true,position:99};r.sheets.Questions.data.push(headers.map(key=>q[key]??''));
 const p=payload();p.answers.place={lat:55,lng:37};c.submitResponse(p);const h=r.sheets.Responses.data[0];assert.deepEqual(JSON.parse(r.sheets.Responses.data[1][h.indexOf('q:place')]),{lat:55,lng:37});
 p.submissionId='submission_invalid_map';p.answers.place={lat:91,lng:37};assert.throws(()=>c.submitResponse(p));assert.equal(r.sheets.Responses.getLastRow(),2);
});
test('upgrade adds fio_range to existing settings without overwriting customized rows',()=>{
 const r=runtime(),c=r.context;c.setupForm();const data=r.sheets.Settings.data;data.splice(data.findIndex(row=>row[0]==='fio_range'),1);setting(r,'form_title','Мой опрос');c.setupForm();assert.equal(c.getFormSchema().title,'Мой опрос');assert.equal(data.filter(row=>row[0]==='fio_range').length,1);
});
