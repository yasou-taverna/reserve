const SHEET_ID = '1Y6744iP3IjAQQtPrGrs6B-1whqefeb0h-gQG0ZwUtFg';
const SHEET_NAME = 'Reservations';

function doGet(e){return handle(e)}
function doPost(e){return handle(e)}

function handle(e){
  const a=e.parameter.action||''; let o={ok:true};
  try{
    const s=SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if(a==='getReservations')o.reservations=getAll_(s);
    if(a==='addReservation'){const d=JSON.parse(e.postData?.contents||'{}');upsert_(s,d);o.saved=true}
    if(a==='cancelReservation'){const d=JSON.parse(e.postData?.contents||'{}');cancel_(s,d.id)}
  }catch(err){o={ok:false,error:String(err)}}
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function getAll_(s){const v=s.getDataRange().getValues();const h=v.shift();return v.map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);o.guests=Number(o.guests)||0;return o})}
function upsert_(s,r){const d=s.getDataRange().getValues();const h=d[0];const id=h.indexOf('id');const row=h.map(k=>r[k]??'');if(!r.createdAt)row[h.indexOf('createdAt')]=new Date().toISOString();let idx=-1;for(let i=1;i<d.length;i++)if(String(d[i][id])===String(r.id))idx=i+1;idx>0?s.getRange(idx,1,1,row.length).setValues([row]):s.appendRow(row)}
function cancel_(s,id){const d=s.getDataRange().getValues();const h=d[0];const i=h.indexOf('id');const st=h.indexOf('status');for(let r=1;r<d.length;r++)if(String(d[r][i])===String(id))s.getRange(r+1,st+1).setValue('cancelled')}
