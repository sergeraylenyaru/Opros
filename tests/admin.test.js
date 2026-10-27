const {test}=require('node:test');
const assert=require('node:assert/strict');
const {runtime}=require('./fakes');
function fixture(){const r=runtime();r.context.setupForm();r.context.setupAdmin();const password=r.alerts.at(-1)[1].split('\n')[0];return {...r,password};}
function login(r){return r.context.adminLogin(r.password).token;}
function serialized(value){return JSON.stringify(value);}

test('generated admin password stored only as digest, absent from public schema',()=>{
 const r=fixture();assert.match(r.password,/^opros-[a-f0-9-]{72}$/);assert.notEqual(r.properties.ADMIN_PASSWORD_HASH,r.password);assert.equal(serialized(r.context.getFormSchema()).includes('ADMIN_PASSWORD'),false);assert.equal(serialized(r.context.getFormSchema()).includes(r.password),false);
});
test('all admin reads and saves require a server session; malformed or guessed tokens fail',()=>{
 const r=fixture();for(const token of [undefined,'',{},'a'.repeat(72),r.password]){assert.throws(()=>r.context.adminGetConfig(token));assert.throws(()=>r.context.adminSaveConfig(token,{}));}
 const token=login(r);assert.equal(r.context.adminGetConfig(token).questions.length,7);
});
test('incorrect passwords rejected; login rate limited after ten attempts per time window',()=>{
 const r=fixture();for(let n=0;n<10;n++)assert.throws(()=>r.context.adminLogin('wrong'));
 assert.throws(()=>r.context.adminLogin(r.password),/попыток/);
});
test('logout, expired cache entry, and password rotation invalidate sessions',()=>{
 const r=fixture();let token=login(r);r.context.adminLogout(token);assert.throws(()=>r.context.adminGetConfig(token));
 token=login(r);const key=[...r.cache.keys()].find(key=>key.startsWith('admin:session:'));r.cache.get(key).expiresAt=0;assert.throws(()=>r.context.adminGetConfig(token));
 token=login(r);r.context.setupAdmin();assert.throws(()=>r.context.adminGetConfig(token));assert.throws(()=>r.context.adminLogin(r.password));
});
test('web context cannot initialize or rotate admin password',()=>{
 const r=fixture(),before=r.properties.ADMIN_PASSWORD_HASH;r.context.SpreadsheetApp.getActiveSpreadsheet=()=>null;assert.throws(()=>r.context.setupAdmin());assert.equal(r.properties.ADMIN_PASSWORD_HASH,before);
});
test('admin saves question edits, new choices and settings; preserves FIO and answers',()=>{
 const r=fixture(),token=login(r),c=r.context,config=c.adminGetConfig(token);const before=serialized(r.sheets.Responses.data),people=serialized(r.sheets.People.data);
 config.settings.find(row=>row.key==='form_title').value='Новая анкета';config.questions[0].title='Выберите город проживания';config.questions.push({id:'new_choice',section_id:'details',title:'Новый выбор',type:'radio',required:false,enabled:true,position:100,options_json:'["Первый","Второй"]'});
 const result=c.adminSaveConfig(token,config);assert.notEqual(result.revision,config.revision);const schema=c.getFormSchema();assert.equal(schema.title,'Новая анкета');assert.equal(schema.sections[0].questions[0].id,'respondent_fio');assert.equal(schema.sections[1].questions.at(-1).options.length,2);assert.equal(serialized(r.sheets.Responses.data),before);assert.equal(serialized(r.sheets.People.data),people);
});
test('invalid configuration and stale revision rejected without any write',()=>{
 const r=fixture(),c=r.context,token=login(r),before=serialized(r.sheets.Questions.data);let config=c.adminGetConfig(token);config.questions[0].type='unknown';assert.throws(()=>c.adminSaveConfig(token,config));assert.equal(serialized(r.sheets.Questions.data),before);
 config=c.adminGetConfig(token);r.sheets.Settings.data[1][1]='Внешнее изменение';assert.throws(()=>c.adminSaveConfig(token,config),/уже изменены/);assert.equal(serialized(r.sheets.Questions.data),before);
});
test('reorder or deletion leaving invalid conditions rejected, existing answers untouched',()=>{
 const r=fixture(),c=r.context,token=login(r),config=c.adminGetConfig(token);config.questions=config.questions.filter(q=>q.id!=='role');assert.throws(()=>c.adminSaveConfig(token,config),/предыдущий вопрос/);assert.equal(r.sheets.Responses.getLastRow(),1);
});
test('write failure rolls prior sheets back instead of leaving partial configuration',()=>{
 const r=fixture(),c=r.context,token=login(r),config=c.adminGetConfig(token),before=c.adminGetConfig(token).revision;
 config.settings[0].value='Changed';const getRange=r.sheets.Questions.getRange.bind(r.sheets.Questions);let fail=true;
 r.sheets.Questions.getRange=(...args)=>{const range=getRange(...args),write=range.setValues.bind(range);range.setValues=values=>{if(fail){fail=false;throw Error('Simulated quota failure');}return write(values);};return range;};
 assert.throws(()=>c.adminSaveConfig(token,config),/восстановлены/);assert.equal(c.adminGetConfig(token).revision,before);
});
test('headless bootstrap and admin setup never require a spreadsheet dialog or leak the setup key',()=>{
 const r=runtime(),c=r.context;c.SpreadsheetApp.getUi=()=>{throw Error('Cannot call SpreadsheetApp.getUi() from this context.');};c.startApp();assert.equal(r.properties.SPREADSHEET_ID,r.book.getId());assert.equal(r.properties.ADMIN_PASSWORD_HASH,undefined);
 const key=require('node:crypto').randomBytes(32).toString('hex');r.properties.ADMIN_SETUP_KEY=key;c.startApp();assert.equal(r.properties.ADMIN_SETUP_KEY,undefined);assert.ok(c.adminLogin(key).token);const hash=r.properties.ADMIN_PASSWORD_HASH;c.startApp();assert.equal(r.properties.ADMIN_PASSWORD_HASH,hash);
});
test('invalid pending setup key is rejected without replacing an existing password',()=>{const r=fixture(),hash=r.properties.ADMIN_PASSWORD_HASH;r.properties.ADMIN_SETUP_KEY='weak';assert.throws(()=>r.context.setupAdmin(),/64/);assert.equal(r.properties.ADMIN_PASSWORD_HASH,hash);});
