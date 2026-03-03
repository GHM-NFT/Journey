/***** GHM — Languages SAFE v2b (namespaced; no global collisions) *****/

const L2_LANG_SKIP = new Set([
  "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
  "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
  "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
  "Compact_View","Compact_view","Metadata_QC","Marketplace_Preview","ERC1155 - Editions",
  "_GHM_BODY_AUDIT_","_GHM_COMPACT_AUDIT_","_GHM_CANON_REPORT_","_GHM_SHOWONLY_REPORT_","Control_Languages"
]);

const L2_NORM = s => (s||"").toString().trim().toLowerCase()
  .replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");

function L2_headerIdxMap(sh){
  const lc = sh.getLastColumn(); if (lc<1) return {raw:[], map:{}};
  const raw = sh.getRange(1,1,1,lc).getValues()[0];
  const map = {};
  raw.forEach((h,i)=>{ const k=L2_NORM(h); if(k && map[k]==null) map[k]=i+1; });
  return {raw, map};
}

// Insert newHeader immediately AFTER the anchor column if it exists; else append.
// Returns 1-based column index of the new/existing header.
function L2_ensureAfter(sh, anchorNorm, newHeaderText){
  const hm = L2_headerIdxMap(sh);
  const existing = hm.map[L2_NORM(newHeaderText)];
  if (existing) return existing;

  const anchorCol = hm.map[anchorNorm];
  if (anchorCol){
    sh.insertColumnAfter(anchorCol);
    const newCol = anchorCol + 1;
    sh.getRange(1,newCol).setValue(newHeaderText);
    return newCol;
  } else {
    sh.insertColumnAfter(sh.getLastColumn());
    const newCol = sh.getLastColumn();
    sh.getRange(1,newCol).setValue(newHeaderText);
    return newCol;
  }
}

function L2_fillTranslateIfBlank(sh, srcCol, dstCol, langCode){
  const rows = Math.max(0, sh.getLastRow()-1); if (!rows || !srcCol || !dstCol) return;
  const rng = sh.getRange(2,dstCol,rows,1);
  const cur = rng.getValues();
  for (let r=0;r<rows;r++){
    if (!cur[r][0]){
      const a1 = sh.getRange(r+2, srcCol).getA1Notation();
      cur[r][0] = `=IFERROR(GOOGLETRANSLATE(${a1},"en","${langCode}"),"")`;
    }
  }
  rng.setValues(cur);
}

// Remove accidental empty trailing columns (blank header + no data)
function L2_cleanupTrailingBlankColumns(sh){
  let trimmed = 0;
  while (sh.getLastColumn()>0){
    const lastCol = sh.getLastColumn();
    const head = String(sh.getRange(1,lastCol).getValue()||"").trim();
    const rows = Math.max(0, sh.getLastRow()-1);
    const isBlankHeader = head==="";
    let isBlankBody = true;
    if (rows>0){
      const vals = sh.getRange(2,lastCol,rows,1).getValues();
      isBlankBody = vals.every(r=>r[0]==="" || r[0]===null);
    }
    if (isBlankHeader && isBlankBody){
      sh.deleteColumn(lastCol);
      trimmed++;
      continue;
    }
    break;
  }
  if (trimmed>0) SpreadsheetApp.getActive().toast(`Trimmed ${trimmed} trailing empty column(s).`,"GHM",3);
}

/* ===== PUBLIC: Apply plan from Control_Languages across collections ===== */
function GHM_L2_Apply_From_Plan_SAFE(){
  const ss = SpreadsheetApp.getActive();
  const plan = ss.getSheetByName("Control_Languages");
  const control = ss.getSheetByName("Control");

  // Optional fallback from Control.secondary_languages
  let fallback = "";
  if (control){
    const vals = control.getDataRange().getValues();
    for (let i=1;i<vals.length;i++){
      if (String(vals[i][0]).trim()==="secondary_languages"){
        fallback = String(vals[i][1]||"").trim(); break;
      }
    }
  }

  const perTab = new Map();
  if (plan && plan.getLastRow()>1){
    const rows = plan.getRange(2,1,plan.getLastRow()-1,2).getValues();
    rows.forEach(([tab, langs])=>{
      if (!tab) return;
      perTab.set(String(tab).trim(), String(langs||"").trim());
    });
  }

  let touched=0;
  ss.getSheets().forEach(sh=>{
    if (L2_LANG_SKIP.has(sh.getName()) || sh.getLastColumn()<1) return;

    const raw = perTab.has(sh.getName()) ? perTab.get(sh.getName()) : fallback;
    const codes = (raw||"").split(",").map(s=>s.trim()).filter(Boolean);
    if (!codes.length) return;

    const idx = L2_headerIdxMap(sh).map;
    const titleEn = idx["title_en"], descEn = idx["description_en"], altEn = idx["alt_text_en"];

    codes.forEach(code=>{
      const tLoc = L2_ensureAfter(sh, "title_en",        `title_${code}`);
      const dLoc = L2_ensureAfter(sh, "description_en",  `description_${code}`);
      const aLoc = L2_ensureAfter(sh, "alt_text_en",     `alt_text_${code}`); // correct: alt_text_*

      if (titleEn) L2_fillTranslateIfBlank(sh, titleEn, tLoc, code);
      if (descEn)  L2_fillTranslateIfBlank(sh, descEn,  dLoc, code);
      if (altEn)   L2_fillTranslateIfBlank(sh, altEn,   aLoc, code);
    });

    L2_cleanupTrailingBlankColumns(sh);
    touched++;
  });

  SpreadsheetApp.getActive().toast(`Language plan applied safely to ${touched} tab(s).`, "GHM Languages v2b", 6);
}

/* ===== PUBLIC: Add one language to the ACTIVE tab ===== */
function GHM_L2_Add_To_Active_SAFE(langCode){
  if (!langCode){ SpreadsheetApp.getUi().alert("Provide a language code, e.g. ja / hi / it / zh-Hans"); return; }
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh || L2_LANG_SKIP.has(sh.getName())){ SpreadsheetApp.getUi().alert("Open a collection tab."); return; }
  if (sh.getLastColumn()<1){ SpreadsheetApp.getUi().alert("No headers on this tab."); return; }

  const idx = L2_headerIdxMap(sh).map;
  const titleEn = idx["title_en"], descEn = idx["description_en"], altEn = idx["alt_text_en"];

  const tLoc = L2_ensureAfter(sh, "title_en",        `title_${langCode}`);
  const dLoc = L2_ensureAfter(sh, "description_en",  `description_${langCode}`);
  const aLoc = L2_ensureAfter(sh, "alt_text_en",     `alt_text_${langCode}`);

  if (titleEn) L2_fillTranslateIfBlank(sh, titleEn, tLoc, langCode);
  if (descEn)  L2_fillTranslateIfBlank(sh, descEn,  dLoc, langCode);
  if (altEn)   L2_fillTranslateIfBlank(sh, altEn,   aLoc, langCode);

  L2_cleanupTrailingBlankColumns(sh);
  SpreadsheetApp.getActive().toast(`Added language columns for ${langCode} on '${sh.getName()}'`, "GHM Languages v2b", 6);
}
Through