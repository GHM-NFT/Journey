function GHM_CompactFromActive_FORCE(){
  const WANT = ["Series","Character","Character_Variant","token_id","standard","contract_factory",
                "name_final","slug","operator_filter","category_id","subcategory_id","taxonomy_tags","alt_text_en"];
  const norm = s => (s||"").toString().trim().toLowerCase().replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");
  const ss = SpreadsheetApp.getActive(), src = ss.getActiveSheet();
  if (!src){ SpreadsheetApp.getUi().alert("Open a collection tab first."); return; }
  const SKIP = new Set(["All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
                        "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
                        "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
                        "Compact_View","Metadata_QC","Marketplace_Preview","ERC1155 - Editions"]);
  if (SKIP.has(src.getName())){ SpreadsheetApp.getUi().alert("Open a collection tab (not a system tab)."); return; }

  const rows = Math.max(0, src.getLastRow()-1);
  const headVals = src.getRange(1,1,1,Math.max(1,src.getLastColumn())).getValues()[0];
  const map = {}; headVals.forEach((h,i)=>{ const k=norm(h); if(k && !(k in map)) map[k]=i+1; });
  const headers = WANT.filter(h => map[norm(h)]);   // only those that exist

  // nuke old compact variants
  ["Compact_View","Compact_view","Compact view","CompactView"].forEach(n=>{
    const old = ss.getSheetByName(n); if (old) ss.deleteSheet(old);
  });

  const view = ss.insertSheet("Compact_View");
  if (!headers.length){
    view.getRange(1,1,1,1).setValues([["(no matching Compact headers on source tab)"]]);
  } else {
    view.getRange(1,1,1,headers.length).setValues([headers]);
    if (rows > 0){
      const out=[];
      for (let r=2; r<=rows+1; r++){
        out.push(headers.map(h => src.getRange(r, map[norm(h)]).getValue()));
      }
      view.getRange(2,1,out.length,headers.length).setValues(out);
    }
    // trim columns to exactly the headers we wrote
    if (view.getMaxColumns() > Math.max(1, headers.length))
      view.deleteColumns(Math.max(1, headers.length)+1, view.getMaxColumns()-Math.max(1, headers.length));
  }

  // tidy rows
  const needRows = Math.max(2, (rows>0 ? rows+1 : 1));
  if (view.getMaxRows() > needRows) view.deleteRows(needRows+1, view.getMaxRows()-needRows);

  // style
  view.setFrozenRows(1);
  view.getRange(1,1,1,Math.max(1,headers.length)).setBackground("#1f2937").setFontColor("#ffffff")
      .setFontWeight("bold").setFontFamily("Roboto Condensed").setFontSize(10);
  SpreadsheetApp.getActive().toast(`Compact_View built from: ${src.getName()} (rows: ${rows})`, "GHM", 6);
}
