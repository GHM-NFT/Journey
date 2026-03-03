/***** GHM — FINAL AUDIT (read-only; namespaced as AUD4_) *****/

const AUD4_SKIP = new Set([
  "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
  "TOC","Control","Control_Languages","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
  "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
  "Compact_View","Compact_view","Metadata_QC","Marketplace_Preview","ERC1155 - Editions",
  "_GHM_BODY_AUDIT_","_GHM_COMPACT_AUDIT_","_GHM_CANON_REPORT_","_GHM_SHOWONLY_REPORT_","_GHM_FINAL_AUDIT_"
]);

const AUD4_N = s => (s||"").toString().trim().toLowerCase().replace(/\s*\/\s*/g,"/").replace(/\s+/g," ");

function AUD4_hdrMap(sh){
  const lc = sh.getLastColumn(); if (lc<1) return {raw:[], map:{}, lastNonEmptyIdx:0};
  const raw = sh.getRange(1,1,1,lc).getValues()[0];
  const map = {};
  raw.forEach((h,i)=>{ const k=AUD4_N(h); if(k && map[k]==null) map[k]=i+1; });
  // find last non-empty header cell left-to-right
  let lastNonEmptyIdx = 0;
  for (let i=raw.length-1;i>=0;i--){
    if (String(raw[i]||"").trim()!==""){ lastNonEmptyIdx = i+1; break; }
  }
  return {raw, map, lastNonEmptyIdx};
}

function AUD4_readPlan(){
  const ss = SpreadsheetApp.getActive();
  const plan = ss.getSheetByName("Control_Languages");
  const m = new Map();
  if (plan && plan.getLastRow()>1){
    const rows = plan.getRange(2,1,plan.getLastRow()-1,2).getValues();
    rows.forEach(([tab,langs])=>{
      if (!tab) return;
      const codes = String(langs||"").split(",").map(s=>s.trim()).filter(Boolean);
      m.set(String(tab).trim(), codes);
    });
  }
  return m;
}

function AUD4_pct(valid, total){
  if (!total) return "";
  return Math.round((valid/total)*100) + "%";
}

function GHM_AUD4_Run(){
  const ss = SpreadsheetApp.getActive();
  const plan = AUD4_readPlan();

  const out = [[
    "Sheet",
    "BodyRows",
    "HasFFFFFF",
    "FFFFFF_is_last",
    "Title_en_col",
    "Desc_en_col",
    "Alt_en_col",
    "Planned_Langs",
    "Missing_Lang_Codes",      // any code where one or more of the trio is missing
    "Misplaced_Lang_Codes",    // any code where the trio is not anchored next to EN
    "Extraneous_Lang_Codes",   // present but not planned for this tab
    "Alt_Test_Typos",
    "Duplicate_Headers",
    "Has_Contract_or_Address",
    "standard_721_1155_ok",
    "operator_filter_on_off_ok"
  ]];

  ss.getSheets().forEach(sh=>{
    const name = sh.getName();
    if (AUD4_SKIP.has(name)) return;
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastCol<1) return;

    const body = Math.max(0, lastRow-1);
    const {raw, map, lastNonEmptyIdx} = AUD4_hdrMap(sh);

    // FFFFFF checks
    const hasFFFF = !!map[AUD4_N("FFFFFF")];
    const ffffCol = map[AUD4_N("FFFFFF")] || 0;
    const ffffIsLast = !!(hasFFFF && (ffffCol === lastNonEmptyIdx));

    // English anchors
    const tEn = map[AUD4_N("title_en")] || 0;
    const dEn = map[AUD4_N("description_en")] || 0;
    const aEn = map[AUD4_N("alt_text_en")] || 0;

    // planned languages for this tab
    const planned = plan.get(name) || [];

    // collect found language codes from headers
    const foundLangs = new Set();
    const reLang = /^(title|description|alt_text)_(.+)$/i;
    raw.forEach(h=>{
      const m = String(h||"").match(reLang);
      if (m && m[2]) foundLangs.add(m[2]);
    });

    // check missing/misplaced for planned codes
    const missing = [];
    const misplaced = [];
    planned.forEach(code=>{
      const tLoc = map[AUD4_N(`title_${code}`)] || 0;
      const dLoc = map[AUD4_N(`description_${code}`)] || 0;
      const aLoc = map[AUD4_N(`alt_text_${code}`)] || 0;
      // missing if any of trio absent
      if (!tLoc || !dLoc || !aLoc) missing.push(code);
      // misplaced if present but not anchored right after EN
      const badAnchor = (tLoc && tEn && tLoc !== tEn+1) ||
                        (dLoc && dEn && dLoc !== dEn+1) ||
                        (aLoc && aEn && aLoc !== aEn+1);
      if (badAnchor) misplaced.push(code);
    });

    // extraneous languages = found - planned - en
    const extra = [];
    foundLangs.forEach(code=>{
      if (code.toLowerCase()==="en") return;
      if (!planned.includes(code)) extra.push(code);
    });

    // alt_test_* typos
    const typos = raw.filter(h => /^alt[_ ]?test[_-]/i.test(String(h||""))).length;

    // duplicate headers (by normalized name)
    const counts = {};
    raw.forEach(h=>{
      const k = AUD4_N(h);
      if (!k) return;
      counts[k] = (counts[k]||0)+1;
    });
    const dupCount = Object.values(counts).filter(c=>c>1).length;

    // Contract / Contract address present?
    const hasContract = !!(map[AUD4_N("contract")] || map[AUD4_N("contract address")] || map[AUD4_N("collection address")]);

    // Value quality checks (read-only)
    let stdOK = "", opOK = "";
    if (body>0){
      // standard: 721 / 1155
      const cStd = map[AUD4_N("standard")] || 0;
      if (cStd){
        const vals = sh.getRange(2, cStd, body, 1).getValues().map(r=>String(r[0]||"").trim());
        const ok = vals.filter(v=>v==="721"||v==="1155").length;
        stdOK = AUD4_pct(ok, body);
      }
      // operator_filter: on / off
      const cOp = map[AUD4_N("operator_filter")] || 0;
      if (cOp){
        const vals = sh.getRange(2, cOp, body, 1).getValues().map(r=>String(r[0]||"").trim().toLowerCase());
        const ok = vals.filter(v=>v==="on"||v==="off").length;
        opOK = AUD4_pct(ok, body);
      }
    }

    out.push([
      name,
      body,
      hasFFFF ? "YES" : "NO",
      ffffIsLast ? "YES" : "NO",
      tEn||"",
      dEn||"",
      aEn||"",
      planned.join(","),
      missing.join(","),
      Array.from(new Set(misplaced)).join(","), // uniq
      extra.join(","),
      typos,
      dupCount,
      hasContract ? "YES" : "NO",
      stdOK,
      opOK
    ]);
  });

  const rep = ss.getSheetByName("_GHM_FINAL_AUDIT_") || ss.insertSheet("_GHM_FINAL_AUDIT_");
  rep.clear();
  rep.getRange(1,1,out.length,out[0].length).setValues(out);
  rep.setFrozenRows(1);
  rep.getRange(1,1,1,out[0].length).setBackground("#1f2937").setFontColor("#ffffff").setFontWeight("bold");
  SpreadsheetApp.getActive().toast("Final audit written to _GHM_FINAL_AUDIT_.", "GHM Audit", 6);
}
