/* Real Chromium + local Apps Script service doubles. No npm dependencies. */
const http=require('node:http');
const fs=require('node:fs');
const path=require('node:path');
const os=require('node:os');
const {spawn}=require('node:child_process');
const {runtime}=require('../tests/fakes');
const r=runtime();r.context.initializeSpreadsheet_();
const questionHeaders=r.sheets.Questions.data[0];
r.sheets.Questions.data.push(questionHeaders.map(key=>({id:'location',section_id:'details',title:'Где вы находитесь?',type:'map',required:true,position:99,enabled:true}[key]??'')));
r.context.setupAdmin();const testPassword=r.alerts.at(-1)[1].split('\n')[0];
let submissions=0;
const bridge=`<script>window.testAdminPassword=${JSON.stringify(testPassword)};window.google={script:{run:{withSuccessHandler(fn){this.success=fn;return this;},withFailureHandler(fn){this.failure=fn;return this;},adminLogin(p){this.call('adminLogin',p);},adminGetConfig(t){this.call('adminGetConfig',t);},adminSaveConfig(t,p){this.call('adminSaveConfig',{token:t,config:p});},adminLogout(t){this.call('adminLogout',t);},getFormSchema(){this.call('schema');},submitResponse(p){this.call('submit',p);},call(method,payload){const good=this.success,bad=this.failure;fetch('/api/'+method,{method:'POST',body:JSON.stringify(payload||{})}).then(r=>r.json()).then(r=>r.error?bad({message:r.error}):good(r));}}}};</script>`;
const automation=`<script>
(async()=>{const pause=()=>new Promise(r=>setTimeout(r,25));const assert=(v,m)=>{if(!v)throw Error(m);};const by=id=>document.getElementById(id);async function wait(fn){for(let i=0;i<180;i++){if(fn())return;await pause();}throw Error('Timeout '+fn);}
try {await wait(()=>!by('form-area').hidden);assert(!by('save-draft').checked,'Draft must be opt-in');by('next').click();assert(!by('announcement').hidden,'Required validation');
function set(id,value,event='input'){const el=by(id);assert(el,'Missing '+id);el.value=value;el.dispatchEvent(new Event(event,{bubbles:true}));}
assert(by('questions').firstElementChild.id==='question-respondent_fio','FIO first');by('field-respondent_fio').click();set('person-search','пётр');assert(by('person-results').children.length===1,'Search FIO');by('person-results').firstElementChild.click();set('field-city','moscow','change');by('field-role-1').click();assert(by('field-team_size'),'Conditional team question');set('field-team_size','4');by('field-role').click();assert(!by('field-team_size'),'Hidden team question');by('field-role-1').click();
set('respondent-email','test@example.org');by('save-draft').click();assert(localStorage.length>0,'Draft stored');by('next').click();assert(by('section-title').textContent==='Ваши предпочтения','Navigation');by('field-interests').click();by('field-rating-4').click();by('field-consent').click();by('field-location').click();const view=by('map-view');view.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));assert(by('map-lat').value!=='','Map selection');set('map-lat','-91');by('map-confirm').click();assert(document.querySelector('dialog'),'Invalid coordinate blocked');set('map-lat','55.75');set('map-lng','37.61');by('map-confirm').click();assert(!document.querySelector('dialog[open]'),'Map dialog closed');by('submit').click();await wait(()=>!by('announcement').hidden&&by('announcement').textContent.includes('Тестовый сбой'));by('submit').click();await wait(()=>!by('result').hidden);assert(by('result-title').textContent==='Спасибо за ответы!','Success');assert(localStorage.length===0,'Draft cleared after success');
by('admin-login').click();set('admin-password','wrong');document.querySelector('dialog[open] form').requestSubmit();await wait(()=>document.querySelector('dialog[open] .error').textContent.includes('Неверный пароль'));set('admin-password',window.testAdminPassword);document.querySelector('dialog[open] form').requestSubmit();await wait(()=>!by('admin-panel').hidden);
const city=document.querySelector('[data-question-id="city"]');city.open=true;const title=city.querySelector('input');title.value='Город проживания';title.dispatchEvent(new Event('input',{bubbles:true}));
Array.from(by('admin-fields').querySelectorAll('button')).find(button=>button.textContent==='+ Добавить вопрос').click();let added=Array.from(document.querySelectorAll('[data-question-id]')).find(card=>card.querySelector('input').value==='Новый вопрос');const newId=added.dataset.questionId;let newTitle=added.querySelector('input');newTitle.value='Новый вопрос с галочками';newTitle.dispatchEvent(new Event('input',{bubbles:true}));let type=added.querySelector('select');type.value='multiselect';type.dispatchEvent(new Event('change',{bubbles:true}));added=document.querySelector('[data-question-id="'+newId+'"]');let choices=added.querySelectorAll('textarea')[1];choices.value='yes | Да' + String.fromCharCode(10) + 'no | Нет';choices.dispatchEvent(new Event('input',{bubbles:true}));
by('admin-save').click();await wait(()=>by('admin-message').textContent.includes('Настройки сохранены'));assert(by('admin-state').textContent==='Все изменения сохранены','Admin save');by('admin-exit').click();await wait(()=>!by('form-area').hidden);assert(by('question-city').textContent.includes('Город проживания'),'Updated question visible');assert(by('question-'+newId).querySelectorAll('input[type=checkbox]').length===2,'New multiselect question visible');assert(by('questions').firstElementChild.id==='question-respondent_fio','Mandatory FIO preserved');
document.body.dataset.smoke='passed';}
catch(e){document.body.dataset.smoke='failed';document.body.appendChild(Object.assign(document.createElement('pre'),{id:'smoke-error',textContent:e.stack}));}
})();</script>`;
const server=http.createServer(async(req,res)=>{try{
 if(req.url.startsWith('/api/')){let raw='';for await(const chunk of req)raw+=chunk;let result;if(req.url==='/api/schema')result=r.context.getFormSchema();else if(req.url==='/api/adminLogin')result=r.context.adminLogin(JSON.parse(raw));else if(req.url==='/api/adminGetConfig')result=r.context.adminGetConfig(JSON.parse(raw));else if(req.url==='/api/adminLogout')result=r.context.adminLogout(JSON.parse(raw));else if(req.url==='/api/adminSaveConfig'){const data=JSON.parse(raw);result=r.context.adminSaveConfig(data.token,data.config);}else{result=r.context.submitResponse(JSON.parse(raw));submissions++;if(submissions===1)result={error:'Тестовый сбой связи после записи. Повторите отправку.'};}res.setHeader('Content-Type','application/json');res.end(JSON.stringify(result));return;}
 const html=fs.readFileSync(path.join(__dirname,'../Index.html'),'utf8');res.setHeader('Content-Security-Policy',"img-src 'self' data:");res.setHeader('Content-Type','text/html; charset=utf-8');res.end(html.replace('<script>',bridge+'<script>').replace('</body>',automation+'</body>'));
}catch(e){res.end(JSON.stringify({error:e.message}));}});
(async()=>{
 const profile=fs.mkdtempSync(path.join(os.tmpdir(),'opros-chromium-'));
 await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
 const proc=spawn(process.env.CHROMIUM||'chromium',['--headless','--no-sandbox','--disable-dev-shm-usage','--disable-gpu','--no-first-run','--disable-background-networking','--no-proxy-server','--user-data-dir='+profile,'--remote-debugging-port=0','about:blank'],{detached:true});
 let socket; const timeout=setTimeout(()=>{console.error('Browser test timed out');process.kill(-proc.pid,'SIGKILL');},25000);
 try {
  const address=await new Promise((resolve,reject)=>{let output='';proc.stderr.on('data',chunk=>{output+=chunk;const match=output.match(/DevTools listening on (ws:\/\/[^\s]+)/);if(match)resolve(match[1]);});proc.on('error',reject);proc.on('exit',code=>reject(Error('Chromium exited '+code)));});
  socket=new WebSocket(address);await new Promise((resolve,reject)=>{socket.onopen=resolve;socket.onerror=reject;});
  let id=0;const pending=new Map();socket.onmessage=event=>{const msg=JSON.parse(event.data);if(msg.id&&pending.has(msg.id)){const {resolve,reject}=pending.get(msg.id);pending.delete(msg.id);msg.error?reject(Error(msg.error.message)):resolve(msg.result);}};
  const call=(method,params={},sessionId)=>new Promise((resolve,reject)=>{pending.set(++id,{resolve,reject});socket.send(JSON.stringify({id,method,params,sessionId}));});
  const {targetId}=await call('Target.createTarget',{url:'http://127.0.0.1:'+server.address().port});
  const {sessionId}=await call('Target.attachToTarget',{targetId,flatten:true});
  await call('Runtime.enable',{},sessionId);
  let result;for(let i=0;i<150;i++){result=await call('Runtime.evaluate',{expression:'JSON.stringify({status:document.body?.dataset.smoke,error:document.getElementById("smoke-error")?.textContent})',returnByValue:true},sessionId);result=JSON.parse(result.result.value||'{}');if(result.status)break;await new Promise(resolve=>setTimeout(resolve,50));}
  if(result.status!=='passed'||r.sheets.Responses.getLastRow()!==2||submissions!==2)throw Error(result.error||'Browser flow incomplete: '+JSON.stringify(result));
  console.log('PASS Chromium: conditional fields, validation, navigation, opt-in draft, server write, lost-response retry without duplicate, draft cleanup, admin login/save/logout.');
 } finally {clearTimeout(timeout);if(socket)socket.close();try{process.kill(-proc.pid,'SIGKILL');}catch{}server.closeAllConnections();server.close();await new Promise(resolve=>proc.once('close',resolve));fs.rmSync(profile,{recursive:true,force:true});}
})().catch(e=>{server.closeAllConnections();server.close();console.error(e);process.exitCode=1;});
