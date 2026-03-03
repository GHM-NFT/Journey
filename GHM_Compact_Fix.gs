/***** GHM — Compact Fix (active tab, hard reset + audit) *****/

// Columns we want in Compact (will include only those present on the tab)
const WANT = ["Series","Character","Character_Variant","token_id","standard","contract_factory",
              "name_final","slug","operator_filter","category_id","subcategory_id","taxonomy_tags","alt_text_en"];

// Normalize header text
const norm = s => (s||"").toString().trim().toLowerCase().replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");

// Remove protections & filter views on a sheet if any (best-effort)
function _unprotectAndClearFilters_(sh){
  try {
    const p = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    p.forEach(pt=>{ try{ pt.remove(); }catch(e){} });
  } catch(e){}
  try {
    // Clear basic filter (not filter views)
    const f = sh.getFilter();
    if (f) f.remove();
  } catch(e){}
}

// Delete ‘Compact_View’ and common aliases
function _deleteCompactAliases_(ss){
  ["Compact_View","Compact_view","Compact view","CompactView"].forEach(n=>{
    const sh = ss.getSheetByName(n);
    if (sh){ _unprotectAndClearFilters_(sh); ss.deleteSheet(sh); }
  });
}

// AUDIT: show which compact columns are present/missing on the ACTIVE tab
function GHM_Compact_Audit_Active(){
  const ss = SpreadsheetApp.getActive();
  const src = ss.getActiveSheet();
  if (!src){ SpreadsheetApp.getUi().alert("Open your collection tab first."); return; }
  const cols = src.getLastColumn();
  if (cols < 1){ SpreadsheetApp.getUi().alert("This tab has no headers."); return; }
  const head = src.getRange(1,1,1,cols).getValues()[0];
  const map = {}; head.forEach((h,i)=>{ const k=norm(h); if(k && !(k in map)) map[k]=i+1; });

  const found = [], missing = [];
  WANT.forEach(h => (map[norm(h)] ? found : missing).push(h));

  const rep = ss.getSheetByName("_GHM_COMPACT_AUDIT_") || ss.insertSheet("_GHM_COMPACT_AUDIT_");
  rep.clear();
  rep.getRange(1,1,1,2).setValues([["Found_on_active_tab","Missing_on_active_tab"]]);
  const maxLen = Math.max(found.length, missing.length, 1);
  const out = [];
  for (let i=0;i<maxLen;i++){ out.push([found[i]||"", missing[i]||""]); }
  rep.getRange(2,1,out.length,2).setValues(out);
  SpreadsheetApp.getActive().toast(`Audit written → _GHM_COMPACT_AUDIT_ (Found: ${found.length}, Missing: ${missing.length})`, "GHM", 6);
}

// BUILD: compact from ACTIVE tab only (hard reset, ignores system tabs)
function GHM_CompactFromActive_NOW(){
  const ss = SpreadsheetApp.getActive();
  const src = ss.getActiveSheet();
  if (!src){ SpreadsheetApp.getUi().alert("Open your collection tab first."); return; }

  // Quick guard: skip known system tabs
  const SKIP = new Set(["All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
                        "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
                        "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
                        "Compact_View","Metadata_QC","Marketplace_Preview","ERC1155 - Editions"]);
  if (SKIP.has(src.getName())){ SpreadsheetApp.getUi().alert("Open a collection tab (not a system tab)."); return; }

  const rows = Math.max(0, src.getLastRow()-1);
  if (!rows){
    GHM_Compact_Audit_Active();
    SpreadsheetApp.getUi().alert("No data rows found under the headers on this tab.");
    return;
  }

  // Map headers
  const cols = src.getLastColumn();
  const head = src.getRange(1,1,1,cols).getValues()[0];
  const map = {}; head.forEach((h,i)=>{ const k=norm(h); if(k && !(k in map)) map[k]=i+1; });

  // Determine which WANT headers exist on this tab
  const headers = WANT.filter(h => map[norm(h)]);
  if (!headers.length){
    GHM_Compact_Audit_Active();
    SpreadsheetApp.getUi().alert("None of the Compact headers were found on this tab (see _GHM_COMPACT_AUDIT_).");
    return;
  }

  // Remove old compact tabs and protections
  _deleteCompactAliases_(ss);

  // Create fresh Compact_View
  const view = ss.insertSheet("Compact_View");
  _unprotectAndClearFilters_(view);

  // Write headers
  view.getRange(1,1,1,headers.length).setValues([headers]);

  // Read body in one go, write in one go
  const out = [];
  for (let r=2; r<=rows+1; r++){
    out.push(headers.map(h => src.getRange(r, map[norm(h)]).getValue()));
  }
  if (out.length){
    view.getRange(2,1,out.length,headers.length).setValues(out);
  }

  // Trim & style
  if (view.getMaxColumns() > headers.length)
    view.deleteColumns(headers.length+1, view.getMaxColumns()-headers.length);
  const needRows = Math.max(2, out.length + 1);
  if (view.getMaxRows() > needRows)
    view.deleteRows(needRows+1, view.getMaxRows() - needRows);

  view.setFrozenRows(1);
  view.getRange(1,1,1,headers.length)
      .setBackground("#1f2937").setFontColor("#ffffff")
      .setFontWeight("bold").setFontFamily("Roboto Condensed").setFontSize(10);
  view.getRange(2,1,Math.max(1,out.length),headers.length)
      .setFontFamily("Roboto Condensed").setFontSize(10);

  // Drop a quick audit so you can see exactly what was included/missing
  GHM_Compact_Audit_Active();

  SpreadsheetApp.getActive().toast(`Compact_View rebuilt from: ${src.getName()}`, "GHM", 6);
}
