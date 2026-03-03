/***** GHM — Languages L3 Fix Kit (namespaced; safe, anchored, dynamic) *****/

const L3_SKIP = new Set([
  "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
  "TOC","Control","Control_Languages","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
  "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
  "Compact_View","Compact_view","Metadata_QC","Marketplace_Preview","ERC1155 - Editions",
  "_GHM_BODY_AUDIT_","_GHM_COMPACT_AUDIT_","_GHM_CANON_REPORT_","_GHM_SHOWONLY_REPORT_"
]);

const L3_N = s => (s||"").toString().trim().toLowerCase().replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");

function L3_hdrMap(sh){
  const lc = sh.getLastColumn(); if (lc<1) return {raw:[], map:{}};
  const raw = sh.getRange(1,1,1,lc).getValues()[0];
  const map = {}; raw.forEach((h,i)=>{ const k=L3_N(h); if(k && map[k]==null) map[k]=i+1; });
  return {raw, map};
}

// insert a brand-new column at 1-based index 'idx', shifting existing to the right
function L3_insertAt(sh, idx, headerText){
  sh.insertColumns(idx, 1);
  sh.getRange(1, idx).setValue(headerText);
  return idx;
}

// Ensure a language column exists immediately AFTER the English anchor;
// if the lang column exists elsewhere, move it next to the anchor.
function L3_ensureAnchored(sh, anchorExact, langHeaderExact){
  const hm = L3_hdrMap(sh);
  const anchorCol = hm.map[L3_N(anchorExact)];
  if (!anchorCol) return 0; // no anchor on this tab

  const existing = hm.map[L3_N(langHeaderExact)];
  if (existing === anchorCol+1) return existing; // already anchored

  // Case 1: lang col exists but not in the right spot -> move it
  if (existing){
    // copy header + body to temp, clear old, insert new at anchor+1, paste back
    const rows = Math.max(1, sh.getLastRow()); // include header
    const vals = sh.getRange(1, existing, rows, 1).getValues();
    sh.deleteColumn(existing);
    const newCol = L3_insertAt(sh, anchorCol+1, langHeaderExact);
    // write body back if any
    if (rows>1) sh.getRange(2, newCol, rows-1, 1).setValues(vals.slice(1));
    return newCol;
  }

  // Case 2: doesn’t exist -> create at anchor+1
  const newCol = L3_insertAt(sh, anchorCol+1, langHeaderExact);
  return newCol;
}

// Fill blanks in dstCol with GOOGLETRANSLATE from srcCol
function L3_fillTranslateIfBlank(sh, srcCol, dstCol, code){
  const rows = Math.max(0, sh.getLastRow()-1); if (!rows || !srcCol || !dstCol) return;
  const rng = sh.getRange(2, dstCol, rows, 1);
  const cur = rng.getValues();
  for (let r=0;r<rows;r++){
    if (!cur[r][0]) {
      const a1 = sh.getRange(r+2, srcCol).getA1Notation();
      cur[r][0] = `=IFERROR(GOOGLETRANSLATE(${a1},"en","${code}"),"")`;
    }
  }
  rng.setValues(cur);
}

// Restore a terminal "FFFFFF" column if it was overwritten
function L3_restoreFFFFFF(sh){
  const lc = sh.getLastColumn(); if (lc<1) return;
  const lastHeader = String(sh.getRange(1, lc).getValue()||"").trim();
  if (lastHeader !== "FFFFFF"){
    // add one final col named FFFFFF (do not delete anything)
    sh.insertColumnAfter(lc);
    sh.getRange(1, lc+1).setValue("FFFFFF");
  }
}

// --- PUBLIC 1: Fix active tab for one code (e.g., "ja","hi","it","zh-Hans") ---
function GHM_L3_Fix_OnActive_ForCode(code){
  if (!code){ SpreadsheetApp.getUi().alert("Provide a language code, e.g. ja / hi / it / zh-Hans"); return; }
  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh || L3_SKIP.has(sh.getName())){ SpreadsheetApp.getUi().alert("Open a collection tab."); return; }
  if (sh.getLastColumn()<1){ SpreadsheetApp.getUi().alert("No headers on this tab."); return; }

  const hm = L3_hdrMap(sh).map;
  const tEn = hm[L3_N("title_en")], dEn = hm[L3_N("description_en")], aEn = hm[L3_N("alt_text_en")];

  const tLoc = L3_ensureAnchored(sh, "title_en",       `title_${code}`);
  const dLoc = L3_ensureAnchored(sh, "description_en", `description_${code}`);
  const aLoc = L3_ensureAnchored(sh, "alt_text_en",    `alt_text_${code}`);

  if (tEn && tLoc) L3_fillTranslateIfBlank(sh, tEn, tLoc, code);
  if (dEn && dLoc) L3_fillTranslateIfBlank(sh, dEn, dLoc, code);
  if (aEn && aLoc) L3_fillTranslateIfBlank(sh, aEn, aLoc, code);

  // Repair any alt_test_* typo on this tab
  L3_renameAltTest_OnSheet(sh);

  // Ensure we still have a terminal FFFFFF column
  L3_restoreFFFFFF(sh);

  SpreadsheetApp.getActive().toast(`Anchored language trio for ${code} on '${sh.getName()}'`, "GHM L3", 6);
}

// --- PUBLIC 2: Apply plan (Control_Languages) with safe anchored placement ---
function GHM_L3_ApplyPlan_SafeAnchored(){
  const ss = SpreadsheetApp.getActive();
  const plan = ss.getSheetByName("Control_Languages");
  if (!plan || plan.getLastRow()<2){ SpreadsheetApp.getUi().alert("No Control_Languages plan found."); return; }

  // build map: tab -> [codes]
  const rows = plan.getRange(2,1,plan.getLastRow()-1,2).getValues();
  const perTab = new Map();
  rows.forEach(([tab,langs])=>{
    if (!tab) return;
    const codes = String(langs||"").split(",").map(s=>s.trim()).filter(Boolean);
    perTab.set(String(tab).trim(), codes);
  });

  let touched = 0;
  ss.getSheets().forEach(sh=>{
    if (L3_SKIP.has(sh.getName()) || sh.getLastColumn()<1) return;
    const codes = perTab.get(sh.getName());
    if (!codes || !codes.length) return;

    const hm = L3_hdrMap(sh).map;
    const tEn = hm[L3_N("title_en")], dEn = hm[L3_N("description_en")], aEn = hm[L3_N("alt_text_en")];

    codes.forEach(code=>{
      const tLoc = L3_ensureAnchored(sh, "title_en",       `title_${code}`);
      const dLoc = L3_ensureAnchored(sh, "description_en", `description_${code}`);
      const aLoc = L3_ensureAnchored(sh, "alt_text_en",    `alt_text_${code}`);
      if (tEn && tLoc) L3_fillTranslateIfBlank(sh, tEn, tLoc, code);
      if (dEn && dLoc) L3_fillTranslateIfBlank(sh, dEn, dLoc, code);
      if (aEn && aLoc) L3_fillTranslateIfBlank(sh, aEn, aLoc, code);
    });

    L3_renameAltTest_OnSheet(sh);
    L3_restoreFFFFFF(sh);
    touched++;
  });

  SpreadsheetApp.getActive().toast(`Applied anchored language plan to ${touched} tab(s).`, "GHM L3", 6);
}

// --- PUBLIC 3: Dynamic Show Only (keeps language cols per plan on each tab) ---
function GHM_ShowOnlyColumns_OnAllCollections_DYNAMIC_LANG(){
  const ss = SpreadsheetApp.getActive();
  const plan = ss.getSheetByName("Control_Languages");
  const planMap = new Map();
  if (plan && plan.getLastRow()>1){
    const r = plan.getRange(2,1,plan.getLastRow()-1,2).getValues();
    r.forEach(([tab,langs])=>{
      if (!tab) return; planMap.set(String(tab).trim(),
        String(langs||"").split(",").map(s=>s.trim()).filter(Boolean));
    });
  }

  ss.getSheets().forEach(sh=>{
    const name = sh.getName();
    if (L3_SKIP.has(name) || sh.getLastColumn()<1) return;

    const langs = planMap.get(name) || []; // keep only planned languages here
    const keepBase = new Set([
      "token_id","title/name","pantheon","frame","pallette","format/medium","stylisation",
      "license_url","image_filename","animation_filename","background_color",
      "title_en","description_en","edition_size","token_range","series","character","character_variant",
      "frame_style","colorway","edition_type","medium","name_final","slug","alt_text_en",
      "attributes_json","image_mime","image_bytes","animation_mime","animation_bytes",
      "collection_path","deity_or_collection","collection_item","meaning/story","tier","token_name",
      "masters","previews","contract","contract address","price_native","currency","chain","license","external_url","description",
      "standard","contract_factory","unlockable_zip_filename","unlockable_zip_url","unlockable_zip_bytes",
      "unlockable_zip_sha256","unlockable_notes","operator_filter","operator_policy_note",
      "category_id","subcategory_id","taxonomy_tags","schema_version","add","ffffff"
    ]);

    // add planned language columns for this tab
    langs.forEach(code=>{
      keepBase.add(`title_${code}`);
      keepBase.add(`description_${code}`);
      keepBase.add(`alt_text_${code}`);
    });

    // now hide all columns not in keepBase
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h=>L3_N(h));
    sh.showColumns(1, lastCol);
    let start=-1;
    for (let c=1;c<=lastCol;c++){
      const h = headers[c-1];
      const keep = !!keepBase.has(h);
      if (!keep && start===-1) start=c;
      if ((keep || c===lastCol) && start!==-1){
        const end = keep ? c-1 : c;
        const len = end - start + 1;
        try{ sh.hideColumns(start, len); }catch(e){}
        start=-1;
      }
    }
  });

  SpreadsheetApp.getActive().toast("Dynamic Show Only applied (per-tab languages from Control_Languages).","GHM View",6);
}

// --- PUBLIC 4: Rename any alt_test_* → alt_text_* across all tabs ---
function GHM_L3_RenameAltTest_All(){
  const ss=SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh=>{
    if (sh.getLastColumn()<1) return;
    L3_renameAltTest_OnSheet(sh);
  });
  SpreadsheetApp.getActive().toast("Renamed any alt_test_* to alt_text_* where found.","GHM L3",5);
}

function L3_renameAltTest_OnSheet(sh){
  const lc = sh.getLastColumn(); if (lc<1) return;
  const hdrs = sh.getRange(1,1,1,lc).getValues()[0];
  const re = /^alt[_ ]?test[_-](.+)$/i;
  let changed=0;
  for (let i=0;i<hdrs.length;i++){
    const h = String(hdrs[i]||"");
    const m = h.match(re);
    if (m){
      const code = m[1];
      sh.getRange(1, i+1).setValue(`alt_text_${code}`);
      changed++;
    }
  }
  return changed;
}
