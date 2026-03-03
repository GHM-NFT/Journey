/***** GHM — Canonicalize Headers (standalone, no dependencies) *****/
/* Fixes: different header order/missing columns per tab.
   Workflow:
     1) Open your best tab (e.g., “Greek - The Olympians”)
     2) Run GHM_CANON_Capture_FromActive()
     3) Run GHM_CANON_Apply_To_AllTabs()
*/

const GHM_CANON_PROP_KEY = "GHM_CANON_HEADERS_V1";

// System tabs to skip (edit if you’ve renamed anything)
const GHM_CANON_SKIP = new Set([
  "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
  "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
  "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
  "Compact_View","Metadata_QC","Marketplace_Preview","ERC1155 - Editions"
]);

const ghm_norm = s => (s||"").toString().trim().toLowerCase()
  .replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");

/* ——— Capture canonical header row from ACTIVE sheet ——— */
function GHM_CANON_Capture_FromActive(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh){ SpreadsheetApp.getUi().alert("Open a collection tab first."); return; }
  if (GHM_CANON_SKIP.has(sh.getName())){ SpreadsheetApp.getUi().alert("Open a collection tab (not a system tab)."); return; }
  const c = sh.getLastColumn(); if (c < 1){ SpreadsheetApp.getUi().alert("No headers on this tab."); return; }
  const headers = sh.getRange(1,1,1,c).getValues()[0].map(h => (h==null?"":String(h)));
  PropertiesService.getScriptProperties().setProperty(GHM_CANON_PROP_KEY, JSON.stringify(headers));
  SpreadsheetApp.getActive().toast(`Captured ${headers.length} headers from '${sh.getName()}'`, "GHM Canon", 6);
}

/* ——— Apply canonical headers to ALL collection tabs ——— */
function GHM_CANON_Apply_To_AllTabs(){
  const ss = SpreadsheetApp.getActive();
  const prop = PropertiesService.getScriptProperties().getProperty(GHM_CANON_PROP_KEY);
  if (!prop){ SpreadsheetApp.getUi().alert("Run GHM_CANON_Capture_FromActive() first."); return; }
  const CANON = JSON.parse(prop) || [];
  if (!CANON.length){ SpreadsheetApp.getUi().alert("Canonical header list is empty."); return; }

  const wantNorm = CANON.map(ghm_norm);

  let fixed = 0, skipped = 0;

  ss.getSheets().forEach(src=>{
    const name = src.getName();
    if (GHM_CANON_SKIP.has(name) || src.getLastColumn()<1){ skipped++; return; }

    const rows = Math.max(0, src.getLastRow()-1);
    const c = src.getLastColumn();
    const block = src.getRange(1,1,Math.max(1,rows)+1,c).getValues(); // header + data
    const headNorm = block[0].map(ghm_norm);
    const srcIndexByNorm = {}; headNorm.forEach((h,i)=>{ if(h && srcIndexByNorm[h]==null) srcIndexByNorm[h]=i; });

    // Create a fresh temp sheet to avoid “destination index within span” errors
    const tmp = ss.insertSheet(`_GHM_TMP_${Date.now()}`);
    // Ensure capacity
    if (tmp.getMaxColumns() < CANON.length) tmp.insertColumnsAfter(tmp.getMaxColumns(), CANON.length - tmp.getMaxColumns());
    if (tmp.getMaxRows() < Math.max(2, rows+1)) tmp.insertRowsAfter(tmp.getMaxRows(), Math.max(2, rows+1) - tmp.getMaxRows());

    // Write canonical headers exactly as captured
    tmp.getRange(1,1,1,CANON.length).setValues([CANON]);

    // Map & write body in one pass
    if (rows > 0){
      const out = new Array(rows);
      for (let r=0; r<rows; r++){
        const srcRow = block[r+1];
        const newRow = new Array(CANON.length).fill("");
        for (let i=0;i<CANON.length;i++){
          const idx = srcIndexByNorm[ wantNorm[i] ];
          if (idx != null) newRow[i] = srcRow[idx];
        }
        out[r] = newRow;
      }
      tmp.getRange(2,1,rows,CANON.length).setValues(out);
    }

    // Trim temp to exact size
    if (tmp.getMaxColumns() > CANON.length) tmp.deleteColumns(CANON.length+1, tmp.getMaxColumns()-CANON.length);
    const needRows = Math.max(2, rows+1);
    if (tmp.getMaxRows() > needRows) tmp.deleteRows(needRows+1, tmp.getMaxRows()-needRows);

    // Replace original sheet at the same index
    const idx = src.getIndex();
    ss.deleteSheet(src);
    tmp.setName(name);
    ss.setActiveSheet(tmp);
    ss.moveActiveSheet(idx);

    fixed++;
  });

  SpreadsheetApp.getActive().toast(`Canonicalized: ${fixed} tab(s). Skipped: ${skipped}`, "GHM Canon", 6);

  // Optional: re-apply your validations/colours if those functions exist
  try{ Control_Apply_Safe && Control_Apply_Safe(); }catch(e){}
  try{ Color_Repaint_ThisTab && Color_Repaint_ThisTab(); }catch(e){}
}

/* ——— Report tabs that were missing canonical headers (before/after) ——— */
function GHM_CANON_Report_Missing(){
  const ss = SpreadsheetApp.getActive();
  const prop = PropertiesService.getScriptProperties().getProperty(GHM_CANON_PROP_KEY);
  if (!prop){ SpreadsheetApp.getUi().alert("Run GHM_CANON_Capture_FromActive() first."); return; }
  const CANON = JSON.parse(prop) || [];
  const want = new Set(CANON.map(ghm_norm));
  const out = [["Sheet","Missing_Count","First_10_Missing"]];

  ss.getSheets().forEach(sh=>{
    if (GHM_CANON_SKIP.has(sh.getName()) || sh.getLastColumn()<1) return;
    const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(ghm_norm);
    const have = new Set(head);
    const miss = [];
    want.forEach(h=>{ if (!have.has(h)) miss.push(h); });
    out.push([sh.getName(), miss.length, miss.slice(0,10).join(", ")]);
  });

  const rep = ss.getSheetByName("_GHM_CANON_REPORT_") || ss.insertSheet("_GHM_CANON_REPORT_");
  rep.clear();
  rep.getRange(1,1,out.length,out[0].length).setValues(out);
  SpreadsheetApp.getActive().toast("Wrote _GHM_CANON_REPORT_", "GHM Canon", 5);
}
