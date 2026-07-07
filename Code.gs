const SHEET_ID='1Y6744iP3IjAQQtPrGrs6B-1whqefeb0h-gQG0ZwUtFg';
const SHEET_NAME='Reservations';
function doGet(e){return handle(e)}
function doPost(e){return handle(e)}
function handle(e){
  let o={ok:true};
  try{
    const s=SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const p=e.parameter.action||JSON.parse(e.postData?.contents||'{}').action;
    if(p=='getReservations')o.reservations=getAll(s);
    if(p=='addReservation')upsert(s,JSON.parse(e.postData.contents));
    if(p=='updateStatus'){const b=JSON.parse(e.postData.contents);updateStatus(s,b.id,b.status)}
  }catch(err){o={ok:false,error:String(err)}}
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)
}
function getAll(s){const v=s.getDataRange().getValues();const h=v.shift();return v.map(r=>{let o={};h.forEach((k,i)=>o[k]=r[i]);return o})}
function upsert(s,r){const d=s.getDataRange().getValues();const h=d[0]||['id','customerName','phone','date','time','guests','zone','tableId','status','notes','createdAt'];if(d.length==0)s.appendRow(h);const row=h.map(k=>r[k]||'');let idx=-1;for(let i=1;i<d.length;i++)if(d[i][0]==r.id)idx=i+1;idx>0?s.getRange(idx,1,1,row.length).setValues([row]):s.appendRow(row)}
function updateStatus(s,id,status){const d=s.getDataRange().getValues();for(let i=1;i<d.length;i++)if(d[i][0]==id){s.getRange(i+1,9).setValue(status);break}}
