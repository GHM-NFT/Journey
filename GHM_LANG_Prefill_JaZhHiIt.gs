function GHM_LANG_Prefill_JaZhHiIt(){
  const ss = SpreadsheetApp.getActive();
  let plan = ss.getSheetByName("Control_Languages");
  if (!plan){
    // create minimal plan sheet
    plan = ss.insertSheet("Control_Languages");
    plan.getRange(1,1,1,3).setValues([["tab_name","langs","notes"]]);
  }
  // index current rows
  const last = plan.getLastRow();
  const map = new Map(); // tab_name -> row
  if (last > 1){
    const vals = plan.getRange(2,1,last-1,1).getValues();
    vals.forEach((v,i)=>{ const t=String(v[0]||"").trim(); if(t) map.set(t, i+2); });
  }

  const rowsToUpsert = [];
  ss.getSheets().forEach(sh=>{
    const name = sh.getName();
    // skip admin/system sheets
    const skip = new Set([
      "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
      "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
      "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
      "Compact_View","Metadata_QC","Marketplace_Preview","ERC1155 - Editions",
      "_GHM_BODY_AUDIT_","_GHM_COMPACT_AUDIT_","_GHM_CANON_REPORT_","_GHM_SHOWONLY_REPORT_","Control_Languages"
    ]);
    if (skip.has(name) || sh.getLastColumn()<1) return;

    const nLower = name.toLowerCase();
    let langs = "";
    if (nLower.includes("japanese")) langs = "ja";
    else if (nLower.includes("chinese")) langs = "zh-Hans"; // add ",zh-Hant" if you want both
    else if (nLower.includes("roman")) langs = "it";
    else if (nLower.includes("hindu")) langs = "hi";

    if (langs){
      const row = map.get(name);
      if (row){ plan.getRange(row,2).setValue(langs); }
      else { rowsToUpsert.push([name, langs, "prefilled by script"]); }
    }
  });

  if (rowsToUpsert.length){
    const start = plan.getLastRow()+1;
    plan.getRange(start,1,rowsToUpsert.length,3).setValues(rowsToUpsert);
  }
  plan.setFrozenRows(1);
  SpreadsheetApp.getActive().toast("Control_Languages prefilled (ja / zh-Hans / hi / it). Review & adjust as needed.", "GHM Languages", 6);
}
