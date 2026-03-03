/***** GHM — Show Only Desired Columns (hide the rest) *****/

/* 1) Your desired columns (exact list you provided) */
const GHM_WANTED_HEADERS = [
  "token_id","Title/Name","Pantheon","Frame","Pallette","Format/Medium","Stylisation",
  "License_URL","image_filename","animation_filename","background_color",
  "title_en","description_en","title_zh-Hans","description_zh-Hans",
  "edition_size","token_range","Series","Character","Character_Variant","Frame_Style",
  "Colorway","Edition_Type","Medium","name_final","slug","alt_text_en","alt_text_zh-Hans",
  "attributes_json","image_mime","image_bytes","animation_mime","animation_bytes",
  "collection_path","deity_or_collection","collection_item","Meaning/Story","Tier","token_name",
  "Masters","Previews","Contract","price_native","currency","chain","license","external_url","description",
  "standard","contract_factory","unlockable_zip_filename","unlockable_zip_url","unlockable_zip_bytes",
  "unlockable_zip_sha256","unlockable_notes","operator_filter","operator_policy_note",
  "category_id","subcategory_id","taxonomy_tags","schema_version","Add","FFFFFF"
];

/* 2) System / index tabs we should NOT touch */
const GHM_EXCLUDE_TABS = new Set([
  "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
  "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
  "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
  "Compact_View","Compact_view","Metadata_QC","Marketplace_Preview","ERC1155 - Editions",
  "_GHM_BODY_AUDIT_","_GHM_COMPACT_AUDIT_","_GHM_CANON_REPORT_"
]);

/* 3) Normalizer + alias helper so header variations still match */
const _norm = s => (s||"").toString().trim().toLowerCase()
  .replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");
function _aliases(h){
  const n = _norm(h);
  // add lightweight aliases here for common variants you use
  if (n==="title/name" || n==="title / name") return "title/name";
  if (n==="format/medium" || n==="format medium") return "format/medium";
  if (n==="frame style" || n==="frame_style") return "frame_style";
  if (n==="alt_text_zh-hans" || n==="alt text zh-hans") return "alt_text_zh-hans";
  if (n==="title zh-hans" || n==="title_zh-hans") return "title_zh-hans";
  if (n==="description zh-hans" || n==="description_zh-hans") return "description_zh-hans";
  if (n==="animation filename" || n==="animations_filename") return "animation_filename";
  if (n==="license url" || n==="license_url") return "license_url";
  if (n==="contract") return "contract"; // keep as-is; you also have 'contract_factory'
  return n;
}

/* Precompute wanted set (normalized+aliased) */
const _WANTED_SET = new Set(GHM_WANTED_HEADERS.map(h => _aliases(h)));

/* --- Core: hide all non-wanted columns on a given sheet --- */
function GHM_ShowOnlyColumns_OnSheet_(sh){
  if (!sh || GHM_EXCLUDE_TABS.has(sh.getName())) return;
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;

  // Unhide everything first so we start clean
  sh.showColumns(1, lastCol);

  // Build keep mask by header row
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const keep = headers.map(h => _WANTED_SET.has(_aliases(h)));

  // Find consecutive runs of columns to HIDE (keep==false) and hide in batches
  let start = -1;
  for (let c = 0; c < lastCol; c++){
    const isKeep = keep[c] === true;
    if (!isKeep && start === -1){ start = c+1; }          // 1-based
    if ((isKeep || c === lastCol - 1) && start !== -1){
      const end = isKeep ? c : c+1;                       // end 1-based inclusive
      const len = end - start + 1;
      try { sh.hideColumns(start, len); } catch(e) {}
      start = -1;
    }
  }
}

/* 4-A) Apply to the ACTIVE tab */
function GHM_ShowOnlyColumns_OnActiveTab(){
  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh){ SpreadsheetApp.getUi().alert("Open a collection tab."); return; }
  if (GHM_EXCLUDE_TABS.has(sh.getName())){ SpreadsheetApp.getUi().alert("You're on a system tab. Open a collection tab."); return; }
  GHM_ShowOnlyColumns_OnSheet_(sh);
  SpreadsheetApp.getActive().toast(`Shown only desired columns on: ${sh.getName()}`, "GHM", 5);
}

/* 4-B) Apply to ALL collection tabs */
function GHM_ShowOnlyColumns_OnAllCollections(){
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh=>{
    if (GHM_EXCLUDE_TABS.has(sh.getName())) return;
    GHM_ShowOnlyColumns_OnSheet_(sh);
  });
  SpreadsheetApp.getActive().toast("Shown only desired columns on all collection tabs.", "GHM", 6);
}

/* 5) Unhide everything on ACTIVE tab (quick toggle) */
function GHM_Unhide_All_OnActive(){
  const sh = SpreadsheetApp.getActiveSheet(); if (!sh) return;
  const lastCol = sh.getLastColumn(); if (lastCol<1) return;
  sh.showColumns(1, lastCol);
  SpreadsheetApp.getActive().toast(`All columns visible on: ${sh.getName()}`, "GHM", 4);
}

/* 6) Unhide everything on ALL collection tabs */
function GHM_Unhide_All_OnAllCollections(){
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh=>{
    if (sh.getLastColumn()<1) return;
    const lastCol = sh.getLastColumn();
    try { sh.showColumns(1, lastCol); } catch(e){}
  });
  SpreadsheetApp.getActive().toast("All columns visible on all tabs.", "GHM", 4);
}

/* 7) Report: which desired columns are missing per tab (helps you fill gaps) */
function GHM_Report_MissingDesiredColumns(){
  const ss = SpreadsheetApp.getActive();
  const out = [["Sheet","Missing_Count","First_20_Missing"]];
  ss.getSheets().forEach(sh=>{
    if (GHM_EXCLUDE_TABS.has(sh.getName()) || sh.getLastColumn()<1) return;
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>_aliases(h));
    const have = new Set(headers);
    const missing = [];
    _WANTED_SET.forEach(w => { if (!have.has(w)) missing.push(w); });
    out.push([sh.getName(), missing.length, missing.slice(0,20).join(", ")]);
  });
  const rep = ss.getSheetByName("_GHM_SHOWONLY_REPORT_") || ss.insertSheet("_GHM_SHOWONLY_REPORT_");
  rep.clear();
  rep.getRange(1,1,out.length,out[0].length).setValues(out);
  SpreadsheetApp.getActive().toast("Wrote _GHM_SHOWONLY_REPORT_ (missing desired headers by tab).", "GHM", 6);
}
