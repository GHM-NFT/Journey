/***** GHM — RESCUE COMPACT (standalone, no dependencies) *****/

// Desired columns (only those that exist on a tab will be used)
const RC_WANT = [
  "Series","Character","Character_Variant",
  "token_id","standard","contract_factory",
  "name_final","slug","operator_filter",
  "category_id","subcategory_id","taxonomy_tags",
  "alt_text_en"
];

// Skip “system” tabs
const RC_SKIP = new Set([
  "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
  "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
  "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
  "Compact_View","Compact_view","Metadata_QC","Marketplace_Preview","ERC1155 - Editions"
]);

const rcNorm = s => (s||"").toString().trim().toLowerCase().replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");

// === 0) Verify (why compact might not build) ===
function GHM_RC_VerifyActive(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh) { SpreadsheetApp.getUi().alert("Open a collection tab first."); return; }
  const name = sh.getName();
  if (RC_SKIP.has(name)) { SpreadsheetApp.getUi().alert(`'${name}' is a system tab. Open a collection tab.`); return; }

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  const body = Math.max(0, lastRow - 1);

  let protections = 0;
  try { protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length; } catch(e){}

  let filters = !!sh.getFilter();
  const head = lastCol>=1 ? sh.getRange(1,1,1,lastCol).getValues()[0] : [];
  const map = {};
  head.forEach((h,i)=>{ const k=rcNorm(h); if(k && !(k in map)) map[k]=i+1; });

  const found = RC_WANT.filter(h => map[rcNorm(h)]);
  const missing = RC_WANT.filter(h => !map[rcNorm(h)]);

  const msg = [
    `Tab: ${name}`,
    `Headers: ${lastCol || 0}`,
    `Body rows: ${body}`,
    `Protections: ${protections}`,
    `Basic filter present: ${filters ? "yes" : "no"}`,
    `Found columns (${found.length}): ${found.join(", ") || "-"}`,
    `Missing columns (${missing.length}): ${missing.join(", ") || "-"}`
  ].join("\n");

  SpreadsheetApp.getUi().alert(msg);
}

// === 1) Build compact for ACTIVE tab → CV_<TabName> (no collisions) ===
function GHM_RC_CompactFromActive(){
  const ss = SpreadsheetApp.getActive();
  const src = ss.getActiveSheet();
  if (!src) { SpreadsheetApp.getUi().alert("Open a collection tab first."); return; }
  const name = src.getName();
  if (RC_SKIP.has(name)) { SpreadsheetApp.getUi().alert(`'${name}' is a system tab. Open a collection tab.`); return; }

  const lastCol = src.getLastColumn();
  if (lastCol < 1){ SpreadsheetApp.getUi().alert("No headers on this tab."); return; }
  const lastRow = src.getLastRow();
  const body = Math.max(0, lastRow - 1);

  const head = src.getRange(1,1,1,lastCol).getValues()[0];
  const map = {};
  head.forEach((h,i)=>{ const k=rcNorm(h); if(k && !(k in map)) map[k]=i+1; });
  const headers = RC_WANT.filter(h => map[rcNorm(h)]);
  if (!headers.length){ SpreadsheetApp.getUi().alert("None of the Compact columns exist on this tab."); return; }

  const viewName = `CV_${name}`;
  let view = ss.getSheetByName(viewName);
  if (view) ss.deleteSheet(view);
  view = ss.insertSheet(viewName);

  // write header
  view.getRange(1,1,1,headers.length).setValues([headers]);

  // write body (if any)
  if (body > 0){
    const out = [];
    const data = src.getRange(2,1,body,lastCol).getValues(); // body block once
    const headNorm = head.map(rcNorm);
    const idxs = headers.map(h => headNorm.indexOf(rcNorm(h)));
    for (let r=0; r<body; r++){
      const row = new Array(headers.length);
      for (let i=0;i<idxs.length;i++){ const j = idxs[i]; row[i] = (j>=0 ? data[r][j] : ""); }
      out.push(row);
    }
    if (out.length) view.getRange(2,1,out.length,headers.length).setValues(out);
  }

  // trim + style
  if (view.getMaxColumns() > headers.length) view.deleteColumns(headers.length+1, view.getMaxColumns()-headers.length);
  const needRows = Math.max(2, (body>0 ? body+1 : 2));
  if (view.getMaxRows() > needRows) view.deleteRows(needRows+1, view.getMaxRows()-needRows);

  view.setFrozenRows(1);
  view.getRange(1,1,1,headers.length)
      .setBackground("#1f2937").setFontColor("#ffffff").setFontWeight("bold")
      .setFontFamily("Roboto Condensed").setFontSize(10);

  SpreadsheetApp.getActive().toast(`Built ${viewName} (rows: ${body})`, "GHM Rescue", 6);
}

// === 2) Build a single combined Compact_View from ALL collections (fast) ===
function GHM_RC_CompactAll(){
  const ss = SpreadsheetApp.getActive();
  const tabs = ss.getSheets().filter(sh => !RC_SKIP.has(sh.getName()) && sh.getLastColumn()>=1);

  // union headers present anywhere (keep RC_WANT order)
  const present = new Set(); const headers = [];
  tabs.forEach(sh=>{
    const c = sh.getLastColumn(); if (c<1) return;
    const raw = sh.getRange(1,1,1,c).getValues()[0];
    const map = {}; raw.forEach((h,i)=>{ const k=rcNorm(h); if(k && !(k in map)) map[k]=i; });
    RC_WANT.forEach(h=>{ const k=rcNorm(h); if(map[k]!=null && !present.has(h)){ present.add(h); headers.push(h); }});
  });
  if (!headers.length){ SpreadsheetApp.getUi().alert("No Compact columns found on any collection tab."); return; }

  // clean any old variants to avoid confusion
  ["Compact_View","Compact_view","Compact view","CompactView"].forEach(n=>{
    const z = ss.getSheetByName(n); if (z) ss.deleteSheet(z);
  });

  const view = ss.insertSheet("Compact_View");
  view.getRange(1,1,1,headers.length).setValues([headers]);

  const allRows = [];
  tabs.forEach(sh=>{
    const rows = Math.max(0, sh.getLastRow()-1); if (!rows) return;
    const c = sh.getLastColumn();
    const block = sh.getRange(1,1,rows+1,c).getValues(); // header + body
    const head = block[0].map(rcNorm);
    const idxs = headers.map(h => head.indexOf(rcNorm(h)));
    for (let r=1; r<block.length; r++){
      const src = block[r];
      if (src.every(v=>v==="")) continue;
      const row = new Array(headers.length);
      for (let i=0;i<idxs.length;i++){ const j = idxs[i]; row[i] = (j>=0 ? src[j] : ""); }
      allRows.push(row);
    }
  });

  if (allRows.length){
    view.getRange(2,1,allRows.length,headers.length).setValues(allRows);
  }

  // trim + style
  if (view.getMaxColumns() > headers.length) view.deleteColumns(headers.length+1, view.getMaxColumns()-headers.length);
  const needRows = Math.max(2, allRows.length+1);
  if (view.getMaxRows() > needRows) view.deleteRows(needRows+1, view.getMaxRows()-needRows);
  view.setFrozenRows(1);
  view.getRange(1,1,1,headers.length)
      .setBackground("#1f2937").setFontColor("#ffffff").setFontWeight("bold")
      .setFontFamily("Roboto Condensed").setFontSize(10);

  SpreadsheetApp.getActive().toast(`Compact_View built (rows: ${allRows.length})`, "GHM Rescue", 6);
}
function GHM_RC_CompactFromActive_KEYS(){
  const WANT=["Series","Character","Character_Variant","token_id","standard","contract_factory",
              "name_final","slug","operator_filter","category_id","subcategory_id","taxonomy_tags","alt_text_en"];
  const KEYS=new Set(["token_id","name_final","slug"]);
  const norm=s=>(s||"").toString().trim().toLowerCase().replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");
  const ss=SpreadsheetApp.getActive(), src=ss.getActiveSheet();
  const SKIP=new Set(["All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
                      "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
                      "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
                      "Compact_View","Compact_view","Metadata_QC","Marketplace_Preview","ERC1155 - Editions"]);
  if (!src || SKIP.has(src.getName())){ SpreadsheetApp.getUi().alert("Open a collection tab."); return; }
  const rows=Math.max(0,src.getLastRow()-1), cols=src.getLastColumn(); if (cols<1){ SpreadsheetApp.getUi().alert("No headers."); return; }
  const head=src.getRange(1,1,1,cols).getValues()[0], hNorm=head.map(norm);
  const map={}; hNorm.forEach((h,i)=>{ if(h && map[h]==null) map[h]=i; });
  const headers=WANT.filter(h=>map[norm(h)]!=null);

  const name=`CV_${src.getName()}`; const old=ss.getSheetByName(name); if (old) ss.deleteSheet(old);
  const view=ss.insertSheet(name);

  if (!headers.length){ view.getRange(1,1,1,1).setValue("(no matching headers)"); return; }
  view.getRange(1,1,1,headers.length).setValues([headers]);

  if (rows>0){
    const data=src.getRange(2,1,rows,cols).getValues();
    const idxs=headers.map(h=>map[norm(h)]);
    const keyIdxs=Array.from(KEYS).map(k=>map[norm(k)]).filter(i=>i!=null);
    const out=[];
    for (let r=0;r<rows;r++){
      const row=data[r];
      const keep=keyIdxs.some(i=>{ const v=row[i]; return v!=="" && v!=null; });
      if (!keep) continue;
      out.push(idxs.map(i=>row[i]));
    }
    if (out.length) view.getRange(2,1,out.length,headers.length).setValues(out);
  }

  if (view.getMaxColumns()>headers.length) view.deleteColumns(headers.length+1, view.getMaxColumns()-headers.length);
  const needRows=Math.max(2, view.getLastRow());
  if (view.getMaxRows()>needRows) view.deleteRows(needRows+1, view.getMaxRows()-needRows);

  view.setFrozenRows(1);
  view.getRange(1,1,1,headers.length).setBackground("#1f2937").setFontColor("#fff")
      .setFontWeight("bold").setFontFamily("Roboto Condensed").setFontSize(10);
  SpreadsheetApp.getActive().toast(`Built ${name}`, "GHM", 5);
}
