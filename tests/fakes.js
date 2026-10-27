const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
class Range {
  constructor(sheet, row, col, rows, cols) {Object.assign(this,{sheet,row,col,rows,cols});}
  getValues() {return Array.from({length:this.rows},(_,r)=>Array.from({length:this.cols},(_,c)=>this.sheet.data[this.row+r-1]?.[this.col+c-1]??''));}
  setValues(values) {values.forEach((row,r)=>row.forEach((v,c)=>{(this.sheet.data[this.row+r-1]??=[])[this.col+c-1]=v;}));return this;}
  getDisplayValues(){return this.getValues().map(row=>row.map(value=>String(value)));}
  setFontWeight(){return this;} setBackground(){return this;}
  createTextFinder(id) {const found=this.getValues().some(row=>row[0]===id);const finder={matchEntireCell(){return this;},matchCase(){return this;},useRegularExpression(){return this;},findNext(){return found?{}:null;}};return finder;}
}
class Sheet {
  constructor(){this.data=[];this.maxRows=1000;this.maxColumns=26;}
  getMaxRows(){return this.maxRows;} getMaxColumns(){return this.maxColumns;}
  insertColumnsAfter(after,count){this.maxColumns+=count;} insertRowsAfter(after,count){this.maxRows+=count;}
  getLastRow(){return this.data.length;}
  getLastColumn(){return Math.max(0,...this.data.map(r=>r.length));}
  getRange(...args){return new Range(this,...args);}
  getDataRange(){return this.getRange(1,1,this.getLastRow(),this.getLastColumn());}
  setFrozenRows(){} autoResizeColumns(){}
}
function runtime() {
  const sheets={}, properties={}, cache=new Map(), alerts=[];let locked=false;
  const book={getId:()=> 'local-fixture-sheet',getSheetByName:name=>sheets[name]||null,insertSheet(name){return sheets[name]=new Sheet();},deleteSheet(sheet){for(const name of Object.keys(sheets))if(sheets[name]===sheet)delete sheets[name];}};
  const ui={ButtonSet:{OK:'OK',YES_NO:'YES_NO'},Button:{YES:'YES',NO:'NO'},alert(...args){alerts.push(args);return 'YES';}};
  const context=vm.createContext({console,CacheService:{getScriptCache:()=>({get:key=>{const entry=cache.get(key);return entry&&entry.expiresAt>Date.now()?entry.value:null;},put:(key,value,seconds)=>cache.set(key,{value,expiresAt:Date.now()+seconds*1000}),remove:key=>cache.delete(key)})},SpreadsheetApp:{openById:id=>{if(id!==book.getId())throw Error('wrong ID');return book;},getActiveSpreadsheet:()=>book,getUi:()=>ui,flush(){}},PropertiesService:{getScriptProperties:()=>({getProperty:k=>properties[k],setProperty:(k,v)=>{properties[k]=v;},deleteProperty:k=>{delete properties[k];}})},LockService:{getScriptLock:()=>({tryLock(){if(locked)return false;locked=true;return true;},releaseLock(){locked=false;}})},Utilities:{getUuid:()=>crypto.randomUUID(),DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(algorithm,value)=>Array.from(crypto.createHash(algorithm).update(value).digest())}});
  for(const file of ['Domain.gs','SheetStore.gs','Code.gs','Admin.gs','Forms.gs'])vm.runInContext(fs.readFileSync(path.join(__dirname,'..',file),'utf8'),context,{filename:file});
  book.insertSheet('People').data=[['ФИО'],['Иванов Иван Иванович'],['Петров Пётр Петрович']];
  return {context,book,sheets,properties,cache,alerts,get locked(){return locked;}};
}
module.exports={runtime};
