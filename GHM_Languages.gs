/***** GHM — Languages (Plan-driven per collection) *****/

const GHM_LANG_SKIP = new Set([
  "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
  "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
  "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
  "Compact_View","Compact_view","Metadata_QC","Marketplace_Preview","ERC1155 - Editions",
  "_GHM_BODY_AUDIT_","_GHM_COMPACT_AUDIT_","_GHM_CANON_REPORT_","_GHM_SHOWONLY_REPORT_","Control_Languages"
]);
const _gl_norm = s => (s||"").toString().trim().toLowerCase();

/* ---------- MENU ---------- */
function onOpen(){ GHM_LANG_BuildMenu(); }
function GHM_LANG_BuildMenu(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("GHM Languages")
    .addItem("Initialize Language Plan", "GHM_LANG_Init_Plan")
    .addItem("Open Language Plan", "GHM_LANG_Open_Plan")
    .addSeparator()
    .addItem("Add Language(s) to Active Tab…", "GHM_LANG_Add_To_Active_Prompt")
    .addItem("Apply Plan to Collections", "GHM_LANG_Apply_From_Plan")
    .addToUi();
}

/* ---------- 1) Create/refresh Control_Languages tab ---------- */
function GHM_LANG_Init_Plan(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Control_Languages");
  if (!sh) sh = ss.insertSheet("Control_Languages");

  const headers = ["tab_name","langs","notes"];
  sh.clear();
  sh.getRange(1,1,1,headers.length).setValues([headers]);

  // Gather collection tabs
  const rows = [];
  ss.getSheets().forEach(s=>{
    const n = s.getName();
    if (GHM_LANG_SKIP.has(n)) return;
    if (s.getLastColumn() < 1) return;
    rows.push([n, "", "comma-separated BCP-47 codes (e.g. ja,el,zh-Hant). Leave blank for no translation."]);
  });

  if (rows.length){
    sh.getRange(2,1,rows.length,3).setValues(rows);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,3).setBackground("#1f2937").setFontColor("#ffffff").setFontWeight("bold");
  SpreadsheetApp.getActive().toast(`Control_Languages initialized (${rows.length} collections)`, "GHM Languages", 6);
}

/* ---------- 2) Open Control_Languages quickly ---------- */
function GHM_LANG_Open_Plan(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Control_Languages");
  if (!sh){ GHM_LANG_Init_Plan(); sh = ss.getSheetByName("Control_Languages"); }
  ss.setActiveSheet(sh);
}

/* ---------- 3) Prompt to add langs for the active collection ---------- */
function GHM_LANG_Add_To_Active_Prompt(){
  const ss = SpreadsheetApp.getActive();
  const active = ss.getActiveSheet();
  if (!active || GHM_LANG_SKIP.has(active.getName())){
    SpreadsheetApp.getUi().alert("Open a collection tab (not a system/index sheet).");
    return;
  }
  const tab = active.getName();
  let plan = ss.getSheetByName("Control_Languages"); if (!plan) { GHM_LANG_Init_Plan(); plan = ss.getSheetByName("Control_Languages"); }
  const last = plan.getLastRow();
  // find or add the row for this tab
  let row = -1;
  if (last >= 2){
    const vals = plan.getRange(2,1,last-1,1).getValues();
    for (let i=0;i<vals.length;i++){ if (String(vals[i][0]).trim() === tab){ row = i+2; break; } }
  }
  if (row === -1){
    row = last + 1;
    plan.getRange(row,1,1,3).setValues([[tab,"",""]]);
  }
  // prompt
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Add language code(s) for this collection", `Tab: ${tab}\nEnter comma-separated codes (e.g. ja,el,zh-Hant). Leave blank for none.`, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const langs = String(resp.getResponseText()||"").trim();
  plan.getRange(row,2).setValue(langs);
  SpreadsheetApp.getActive().toast(`Saved languages for '${tab}': ${langs||"(none)"}`, "GHM Languages", 5);
}

/* ---------- 4) Apply plan: add/fill language columns only where requested ---------- */
function GHM_LANG_Apply_From_Plan(){
  const ss = SpreadsheetApp.getActive();
  const plan = ss.getSheetByName("Control_Languages");
  const control = ss.getSheetByName("Control");

  // Optional global fallback from Control.secondary_languages (used only if tab not listed)
  let fallback = "";
  if (control){
    const vals = control.getDataRange().getValues();
    for (let i=1;i<vals.length;i++){
      if (String(vals[i][0]).trim()==="secondary_languages"){
        fallback = String(vals[i][1]||"").trim();
        break;
      }
    }
  }

  // Build per-tab plan map
  const map = new Map();
  if (plan && plan.getLastRow() > 1){
    const rows = plan.getRange(2,1,plan.getLastRow()-1,2).getValues();
    rows.forEach(([tab,langs])=>{
      if (!tab) return;
      map.set(String(tab).trim(), String(langs||"").trim());
    });
  }

  let touched = 0;
  ss.getSheets().forEach(sh=>{
    const name = sh.getName();
    if (GHM_LANG_SKIP.has(name) || sh.getLastColumn()<1) return;

    // Decide which codes to apply
    const raw = map.has(name) ? map.get(name) : fallback;
    const codes = (raw||"").split(",").map(s=>s.trim()).filter(Boolean);
    if (!codes.length) return;

    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1,1,1,lastCol).getValues()[0];
    const idxBy = Object.fromEntries(head.map((h,i)=>[_gl_norm(h), i+1]));

    const ensureCol = (label) => {
      const k = _gl_norm(label);
      if (idxBy[k]) return idxBy[k];
      sh.insertColumnAfter(sh.getLastColumn());
      const col = sh.getLastColumn();
      sh.getRange(1,col).setValue(label);
      idxBy[k]=col; return col;
    };

    const rows = Math.max(0, sh.getLastRow()-1); if (!rows) return;

    const colOf = (h)=>idxBy[_gl_norm(h)]||0;
    const tEn=colOf("title_en"), dEn=colOf("description_en"), aEn=colOf("alt_text_en");

    codes.forEach(code=>{
      const tLoc = ensureCol(`title_${code}`);
      const dLoc = ensureCol(`description_${code}`);
      const aLoc = ensureCol(`alt_text_${code}`);

      // Only fill blanks with GOOGLETRANSLATE drafts; keep any existing human edits
      if (tEn){
        const rng = sh.getRange(2,tLoc,rows,1), vals=rng.getValues();
        for (let r=0;r<rows;r++){
          if (!vals[r][0]) vals[r][0] = `=IFERROR(GOOGLETRANSLATE(${sh.getRange(r+2,tEn).getA1Notation()},"en","${code}"),"")`;
        }
        rng.setValues(vals);
      }
      if (dEn){
        const rng = sh.getRange(2,dLoc,rows,1), vals=rng.getValues();
        for (let r=0;r<rows;r++){
          if (!vals[r][0]) vals[r][0] = `=IFERROR(GOOGLETRANSLATE(${sh.getRange(r+2,dEn).getA1Notation()},"en","${code}"),"")`;
        }
        rng.setValues(vals);
      }
      if (aEn){
        const rng = sh.getRange(2,aLoc,rows,1), vals=rng.getValues();
        for (let r=0;r<rows;r++){
          if (!vals[r][0]) vals[r][0] = `=IFERROR(GOOGLETRANSLATE(${sh.getRange(r+2,aEn).getA1Notation()},"en","${code}"),"")`;
        }
        rng.setValues(vals);
      }
    });

    touched++;
  });

  SpreadsheetApp.getActive().toast(`Applied language plan to ${touched} collection tab(s).`, "GHM Languages", 6);
}
