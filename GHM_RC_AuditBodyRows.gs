function GHM_RC_AuditBodyRows(){
  const ss=SpreadsheetApp.getActive();
  const out=[["Sheet","Body rows (2+)","Non-empty rows (any cell)"]];
  ss.getSheets().forEach(sh=>{
    const lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
    if (lastCol<1) return;
    const body=Math.max(0,lastRow-1);
    if (!body) return;
    const data=sh.getRange(2,1,body,lastCol).getValues();
    const nonEmpty=data.filter(row=>row.some(v=>v!=="" && v!=null)).length;
    out.push([sh.getName(), body, nonEmpty]);
  });
  const rep=ss.getSheetByName("_GHM_BODY_AUDIT_")||ss.insertSheet("_GHM_BODY_AUDIT_");
  rep.clear(); rep.getRange(1,1,out.length,out[0].length).setValues(out);
}

