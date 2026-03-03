/***** GHM — FAST COMPACT (standalone, no conflicts) *****/

// Desired columns (order kept if present)
const GHM_FC_WANT = [
  "Series","Character","Character_Variant",
  "token_id","standard","contract_factory",
  "name_final","slug","operator_filter",
  "category_id","subcategory_id","taxonomy_tags",
  "alt_text_en"
];

// System tabs to skip
const GHM_FC_SKIP = new Set([
  "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
  "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
  "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
  "Compact_View","Metadata_QC","Marketplace_Preview","ERC1155 - Editions"
]);

const ghm_fc_norm = s => (s||"").toString().trim().toLowerCase()
  .replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");

/* Build Compact_View from ALL collection tabs in one go (fast, batched reads/writes) */
function GHM_FastCompact_All(){
  const ss = SpreadsheetApp.getActive();
  const tabs = ss.getSheets().filter(sh => !GHM_FC_SKIP.has(sh.getName()) && sh.getLastColumn() >= 1);

  // Build union headers in desired order
  const present = new Set();
  const headers = [];
  tabs.forEach(sh => {
    const c = sh.getLastColumn(); if (c < 1) return;
    const raw = sh.getRange(1,1,1,c).getValues()[0];
    const map = {};
    raw.forEach((h,i) => { const k = ghm_fc_norm(h); if (k && !(k in map)) map[k] = i; });
    GHM_FC_WANT.forEach(h => { const k = ghm_fc_norm(h);
      if (map[k] != null && !present.has(h)) { present.add(h); headers.push(h); }
    });
  });
  if (!headers.length){ SpreadsheetApp.getUi().alert("No Compact columns found on any tab."); return; }

  // Fresh Compact_View
  const old = ss.getSheetByName("Compact_View"); if (old) ss.deleteSheet(old);
  const view = ss.insertSheet("Compact_View");

  // Header row
  view.getRange(1,1,1,headers.length).setValues([headers]);

  // Collect data efficiently
  const allRows = [];
  tabs.forEach(sh => {
    const rows = Math.max(0, sh.getLastRow() - 1); if (!rows) return;
    const c = sh.getLastColumn();
    const block = sh.getRange(1,1,rows+1,c).getValues(); // header + rows
    const head = block[0].map(ghm_fc_norm);
    const idxs = headers.map(h => head.indexOf(ghm_fc_norm(h)));
    for (let r = 1; r < block.length; r++){
      const src = block[r];
      if (src.every(v => v === "")) continue;
      const out = new Array(headers.length);
      for (let i = 0; i < idxs.length; i++){ const j = idxs[i]; out[i] = (j >= 0 ? src[j] : ""); }
      allRows.push(out);
    }
  });

  if (allRows.length){
    view.getRange(2,1,allRows.length,headers.length).setValues(allRows);
  }

  // Trim + style
  if (view.getMaxColumns() > headers.length)
    view.deleteColumns(headers.length + 1, view.getMaxColumns() - headers.length);
  const needRows = Math.max(2, allRows.length + 1);
  if (view.getMaxRows() > needRows)
    view.deleteRows(needRows + 1, view.getMaxRows() - needRows);

  view.setFrozenRows(1);
  view.getRange(1,1,1,headers.length)
      .setBackground("#1f2937").setFontColor("#ffffff")
      .setFontWeight("bold").setFontFamily("Roboto Condensed").setFontSize(10);
  view.getRange(2,1,Math.max(1,allRows.length),headers.length)
      .setFontFamily("Roboto Condensed").setFontSize(10);

  SpreadsheetApp.getActive().toast(`Compact_View built (rows: ${allRows.length})`, "GHM", 6);
}

/* ---- Batched version to avoid timeouts ---- */
const GHM_FC_BATCH_SIZE = 3;

function GHM_FastCompact_Batch_Reset(){
  PropertiesService.getScriptProperties().deleteProperty("GHM_FC_IDX");
  PropertiesService.getScriptProperties().deleteProperty("GHM_FC_HEADERS");
  const ss = SpreadsheetApp.getActive();
  const old = ss.getSheetByName("Compact_View"); if (old) ss.deleteSheet(old);
  SpreadsheetApp.getActive().toast("Compact batch reset", "GHM", 4);
}

function GHM_FastCompact_Batch_Next(){
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();

  const tabs = ss.getSheets().filter(sh => !GHM_FC_SKIP.has(sh.getName()) && sh.getLastColumn() >= 1);
  let idx = parseInt(props.getProperty("GHM_FC_IDX") || "0", 10);

  // Initialize headers + sheet on first run
  let headers = props.getProperty("GHM_FC_HEADERS");
  let view = ss.getSheetByName("Compact_View");
  if (!headers){
    const present = new Set(); const h = [];
    tabs.forEach(sh => {
      const c = sh.getLastColumn(); if (c < 1) return;
      const raw = sh.getRange(1,1,1,c).getValues()[0];
      const map = {};
      raw.forEach((x,i) => { const k = ghm_fc_norm(x); if (k && !(k in map)) map[k] = i; });
      GHM_FC_WANT.forEach(w => { const k = ghm_fc_norm(w);
        if (map[k] != null && !present.has(w)) { present.add(w); h.push(w); }
      });
    });
    if (!h.length){ SpreadsheetApp.getUi().alert("No Compact columns found on any tab."); return; }
    headers = JSON.stringify(h);
    props.setProperty("GHM_FC_HEADERS", headers);
    if (view) ss.deleteSheet(view);
    view = ss.insertSheet("Compact_View");
    view.getRange(1,1,1,h.length).setValues([h]);
    view.setFrozenRows(1);
    view.getRange(1,1,1,h.length).setBackground("#1f2937").setFontColor("#ffffff")
        .setFontWeight("bold").setFontFamily("Roboto Condensed").setFontSize(10);
  }
  const HEADERS = JSON.parse(headers);

  // Process next chunk of tabs
  const end = Math.min(idx + GHM_FC_BATCH_SIZE, tabs.length);
  for (let t = idx; t < end; t++){
    const sh = tabs[t];
    const rows = Math.max(0, sh.getLastRow() - 1); if (!rows) continue;
    const c = sh.getLastColumn();
    const block = sh.getRange(1,1,rows+1,c).getValues();
    const head = block[0].map(ghm_fc_norm);
    const idxs = HEADERS.map(h => head.indexOf(ghm_fc_norm(h)));

    const out = [];
    for (let r = 1; r < block.length; r++){
      const src = block[r];
      if (src.every(v => v === "")) continue;
      const row = new Array(HEADERS.length);
      for (let i = 0; i < idxs.length; i++){ const j = idxs[i]; row[i] = (j >= 0 ? src[j] : ""); }
      out.push(row);
    }
    if (out.length){
      const start = ss.getSheetByName("Compact_View").getLastRow() + 1;
      ss.getSheetByName("Compact_View").getRange(start,1,out.length,HEADERS.length).setValues(out);
    }
  }

  props.setProperty("GHM_FC_IDX", String(end));
  SpreadsheetApp.getActive().toast(`Compact batch: processed ${end - idx} tab(s). Remaining: ${tabs.length - end}`, "GHM", 6);
}
