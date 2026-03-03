Through// QA the ACTIVE sheet (read-only). Writes an "Active_QA_Report" tab.
function RUN_QA_active(){
  var sh = SpreadsheetApp.getActiveSheet();
  var ss = SpreadsheetApp.getActive();
  var HEADER_ROW = 1;
  var REPORT_SHEET = 'Active_QA_Report';
  var MAX_SCAN = 300; // limit for speed

  // helpers
  function headers(){
    var lastCol = sh.getLastColumn(); if (!lastCol) return [];
    var vals = sh.getRange(HEADER_ROW,1,1,lastCol).getValues()[0];
    var out=[]; for (var i=0;i<vals.length;i++) out.push(String(vals[i]||'').trim()); return out;
  }
  function idx(H,name){ for (var i=0;i<H.length;i++) if (H[i]===name) return i; return -1; }
  function pick(H,nameAltArr){ for (var i=0;i<nameAltArr.length;i++){ var j=idx(H,nameAltArr[i]); if (j!==-1) return j; } return -1; }
  function sampleCol(colIdx1){
    var lastRow = sh.getLastRow(); var n = Math.max(0,lastRow-HEADER_ROW); if (!n) return [];
    var rows = Math.min(n, MAX_SCAN);
    return sh.getRange(HEADER_ROW+1, colIdx1, rows, 1).getValues().map(function(r){ return String(r[0]||'').trim(); });
  }
  function mkReport(){
    var rep = ss.getSheetByName(REPORT_SHEET) || ss.insertSheet(REPORT_SHEET);
    rep.clear();
    rep.appendRow(['sheet','checked_at','check','status','issues','notes / first few rows']);
    return rep;
  }
  function row(rep, check, ok, issues, notes){
    rep.appendRow([sh.getName(), new Date(), check, ok?'OK':'ISSUE', issues, notes||'']);
  }

  var H = headers();
  var rep = mkReport();

  // 1) token_range should be like "1-10" or "5" (NOT JSON traits)
  var iTR = idx(H,'token_range');
  if (iTR === -1){
    row(rep,'token_range column present', false, 1, 'Missing column');
  } else {
    var vals = sampleCol(iTR+1), bad=[], jsony=[];
    for (var r=0;r<vals.length;r++){
      var v = vals[r];
      if (!v) continue;
      if (/^\d+(\s*-\s*\d+)?$/.test(v)) continue;
      if (v.indexOf('"trait_type"')!==-1 || v.charAt(0)=='[' || v.charAt(0)=='{') jsony.push((r+2)+': '+v.slice(0,60));
      else bad.push((r+2)+': '+v);
    }
    var issues = bad.length + jsony.length;
    row(rep,'token_range format', issues===0, issues, jsony.length?('Looks like traits JSON in token_range → e.g. '+jsony.slice(0,3).join(' | ')):(bad.slice(0,3).join(' | ')));
  }

  // 2) attributes_json should contain `[{"trait_type":...}]`
  var iAJ = idx(H,'attributes_json');
  if (iAJ === -1){
    row(rep,'attributes_json column present', false, 1, 'Missing column');
  } else {
    var vals2 = sampleCol(iAJ+1), wrong=[];
    for (var r2=0;r2<vals2.length;r2++){
      var v2 = vals2[r2]; if (!v2) continue;
      if (!(v2.charAt(0)=='[' && v2.indexOf('"trait_type"')!==-1)) wrong.push((r2+2)+': '+v2.slice(0,60));
    }
    row(rep,'attributes_json looks like traits', wrong.length===0, wrong.length, wrong.slice(0,3).join(' | '));
  }

  // 3) Tier in allowed set
  var iTier = idx(H,'Tier');
  if (iTier !== -1){
    var valsT = sampleCol(iTier+1), badT=[];
    var allowed = {'Mythic Icon':1,'Signature Edition':1,'Companion Piece':1,'Limited Edition Print':1,'Relic':1};
    for (var t=0;t<valsT.length;t++){ var v=valsT[t]; if (!v) continue; if (!allowed[v]) badT.push((t+2)+': '+v); }
    row(rep,'Tier allowed values', badT.length===0, badT.length, badT.slice(0,3).join(' | '));
  }

  // 4) background_hex is valid hex
  var iBh = idx(H,'background_hex');
  if (iBh !== -1){
    var valsH = sampleCol(iBh+1), badH=[];
    for (var h=0;h<valsH.length;h++){
      var v=valsH[h]; if (!v) continue;
      var norm = v.charAt(0)=='#'? v.slice(1): v;
      if (!/^[0-9a-fA-F]{6}$/.test(norm)) badH.push((h+2)+': '+v);
    }
    row(rep,'background_hex valid', badH.length===0, badH.length, badH.slice(0,3).join(' | '));
  }

  // 5) edition_size positive integer
  var iEd = idx(H,'edition_size');
  if (iEd !== -1){
    var valsE = sampleCol(iEd+1), badE=[];
    for (var e=0;e<valsE.length;e++){ var v=valsE[e]; if (!v) continue; if (!/^\d+$/.test(v) || parseInt(v,10)<=0) badE.push((e+2)+': '+v); }
    row(rep,'edition_size integer > 0', badE.length===0, badE.length, badE.slice(0,3).join(' | '));
  }

  // 6) Style should not be a version string like "v1.1"
  var iSt = idx(H,'Style');
  if (iSt !== -1){
    var valsS = sampleCol(iSt+1), ver=[];
    for (var s=0;s<valsS.length;s++){ var v=valsS[s]; if (/^v\d+(\.\d+)*$/i.test(v)) ver.push((s+2)+': '+v); }
    row(rep,'Style is descriptive (not version)', ver.length===0, ver.length, ver.slice(0,3).join(' | '));
  }

  // 7) Series shouldn’t equal Pantheon
  var iSe = idx(H,'Series'), iPa = (idx(H,'Pantheon')!==-1? idx(H,'Pantheon'): idx(H,'Culture'));
  if (iSe!==-1 && iPa!==-1){
    var valsSe = sampleCol(iSe+1), valsPa = sampleCol(iPa+1), eq=[];
    for (var q=0;q<Math.min(valsSe.length, valsPa.length); q++){
      var a=valsSe[q].toLowerCase(), b=valsPa[q].toLowerCase();
      if (a && b && a===b) eq.push((q+2)+': '+valsSe[q]);
    }
    row(rep,'Series ≠ Pantheon', eq.length===0, eq.length, eq.slice(0,3).join(' | '));
  }

  // 8) CIDs (image_cid) look like ipfs CIDs
  var iCID = pick(H,['image_cid','cid_Poster_Image']);
  if (iCID !== -1){
    var valsC = sampleCol(iCID+1), badC=[];
    for (var c=0;c<valsC.length;c++){
      var v=valsC[c]; if (!v) continue;
      if (!( /^Qm[1-9A-HJ-NP-Za-km-z]{44}$/.test(v) || /^bafy[0-9A-Za-z]+$/.test(v) )) badC.push((c+2)+': '+v);
    }
    row(rep,'image CID looks valid', badC.length===0, badC.length, badC.slice(0,3).join(' | '));
  }

  // 9) Filenames: json_filename ends with .json
  var iJF = idx(H,'json_filename');
  if (iJF !== -1){
    var valsJ = sampleCol(iJF+1), badJ=[];
    for (var j=0;j<valsJ.length;j++){ var v=valsJ[j]; if (!v) continue; if (!/\.json$/i.test(v)) badJ.push((j+2)+': '+v); }
    row(rep,'json_filename ends .json', badJ.length===0, badJ.length, badJ.slice(0,3).join(' | '));
  }

  // 10) Slug: lowercase letters, numbers, hyphen
  var iSl = idx(H,'slug');
  if (iSl !== -1){
    var valsL = sampleCol(iSl+1), badL=[];
    for (var l=0;l<valsL.length;l++){ var v=valsL[l]; if (!v) continue; if (!/^[a-z0-9-]+$/.test(v)) badL.push((l+2)+': '+v); }
    row(rep,'slug format (a-z0-9-)', badL.length===0, badL.length, badL.slice(0,3).join(' | '));
  }
}
// 1) Set Series from the tab name (part after " - "), but only when Series is blank or equals Pantheon/Culture.
function RUN_setSeriesFromSheetName_active(){
  var sh = SpreadsheetApp.getActiveSheet(), HR=1, name=sh.getName();
  var sep=name.indexOf(' - '), series = sep>-1 ? name.substring(sep+3).trim() : '';
  if (!series) return;
  var lastCol=sh.getLastColumn(); if (!lastCol) return;
  var H=sh.getRange(HR,1,1,lastCol).getValues()[0].map(function(v){return String(v).trim();});
  function idx(n){for (var i=0;i<H.length;i++) if (H[i]===n) return i; return -1;}
  var iSe=idx('Series'), iPa=idx('Pantheon'); if (iPa===-1) iPa=idx('Culture');
  if (iSe===-1) return;
  var LR=sh.getLastRow(); if (LR<=HR) return; var n=LR-HR;
  var rngS=sh.getRange(HR+1,iSe+1,n,1), S=rngS.getValues();
  var P=iPa!==-1 ? sh.getRange(HR+1,iPa+1,n,1).getValues() : null, changed=false;
  for (var r=0;r<n;r++){
    var cur=(S[r][0]||'').toString().trim();
    var pan=P? (P[r][0]||'').toString().trim() : '';
    if (!cur || (pan && cur.toLowerCase()===pan.toLowerCase())){ S[r][0]=series; changed=true; }
  }
  if (changed) rngS.setValues(S);
}

// 2) Normalize Tier values to the 5 allowed labels.
function RUN_fixTierCasing_active(){
  var sh=SpreadsheetApp.getActiveSheet(), HR=1, LC=sh.getLastColumn(); if(!LC) return;
  var H=sh.getRange(HR,1,1,LC).getValues()[0].map(function(v){return String(v).trim();});
  function idx(n){for(var i=0;i<H.length;i++) if(H[i]===n) return i; return -1;}
  var iT=idx('Tier'); if(iT===-1) return;
  var LR=sh.getLastRow(); if(LR<=HR) return; var n=LR-HR;
  var rng=sh.getRange(HR+1,iT+1,n,1), V=rng.getValues(), changed=false;
  function norm(s){return s.toLowerCase().replace(/_/g,' ').replace(/\s+/g,' ').trim();}
  function canon(s){
    var t=norm(s);
    if(t==='mythic icon') return 'Mythic Icon';
    if(t==='signature edition') return 'Signature Edition';
    if(t==='companion piece') return 'Companion Piece';
    if(t==='limited edition print') return 'Limited Edition Print';
    if(t==='relic') return 'Relic';
    return null;
  }
  for(var r=0;r<n;r++){
    var v=(V[r][0]||'').toString().trim(); if(!v) continue;
    var c=canon(v); if(c && c!==v){ V[r][0]=c; changed=true; }
  }
  if(changed) rng.setValues(V);
}

// 3) Clean token_range: if it looks like traits JSON and attributes_json is empty → move it; else just clear token_range.
function RUN_cleanTokenRange_active(){
  var sh=SpreadsheetApp.getActiveSheet(), HR=1, LC=sh.getLastColumn(); if(!LC) return;
  var H=sh.getRange(HR,1,1,LC).getValues()[0].map(function(v){return String(v).trim();});
  function idx(n){for(var i=0;i<H.length;i++) if(H[i]===n) return i; return -1;}
  var iTR=idx('token_range'), iAJ=idx('attributes_json');
  if(iTR===-1){ SpreadsheetApp.getUi().alert('No token_range column'); return; }
  if(iAJ===-1){ SpreadsheetApp.getUi().alert('No attributes_json column'); return; }
  var LR=sh.getLastRow(); if(LR<=HR) return; var n=LR-HR;
  var Rtr=sh.getRange(HR+1,iTR+1,n,1).getValues();
  var Raj=sh.getRange(HR+1,iAJ+1,n,1).getValues();
  var moved=0, cleared=0;
  for(var r=0;r<n;r++){
    var tr=(Rtr[r][0]||'').toString().trim();
    var aj=(Raj[r][0]||'').toString().trim();
    var looksJSON = tr && (tr.charAt(0)=='[' || tr.charAt(0)=='{') && tr.indexOf('"trait_type"')!==-1;
    if(looksJSON){
      if(!aj){ Raj[r][0]=tr; moved++; }
      Rtr[r][0]=''; cleared++;
    }
  }
  if(moved) sh.getRange(HR+1,iAJ+1,n,1).setValues(Raj);
  if(cleared) sh.getRange(HR+1,iTR+1,n,1).setValues(Rtr);
  SpreadsheetApp.getUi().alert('token_range: moved '+moved+' → attributes_json, cleared '+cleared+'.');
}
// Run the three QA fixes across ALL collection tabs (skips ops/utility tabs).
function RUN_applyQAfixes_all(){
  var ss = SpreadsheetApp.getActive();
  // Skip ops/utility tabs so we don't touch audits, control, etc.
  var EXCLUDE = new RegExp('^(Control.*|_?GHM.*|.*_AUDIT_.*|.*_REPORT$|Marketplace_Preview|Metadata_QC|Taxonomy_.*|Traits_Dictionary|All|ERC1155 - Editions|Collections_Index|Characters_Index|Series_Character_Matrix|Control_Languages)$','i');

  var sheets = ss.getSheets(), touched = 0;
  for (var i=0;i<sheets.length;i++){
    var sh = sheets[i];
    if (EXCLUDE.test(sh.getName())) continue;

    ss.setActiveSheet(sh);  // reuse the single-tab helpers safely
    try {
      // 1) Series from sheet name (only if blank or equals Pantheon)
      if (typeof RUN_setSeriesFromSheetName_active === 'function') RUN_setSeriesFromSheetName_active();

      // 2) Tier casing → 5 allowed labels
      if (typeof RUN_fixTierCasing_active === 'function') RUN_fixTierCasing_active();

      // 3) token_range: move traits JSON → attributes_json (if empty), then clear
      if (typeof RUN_cleanTokenRange_active === 'function') RUN_cleanTokenRange_active();

      touched++;
    } catch(e){
      // Keep going; you can check the Execution Log for any sheet that threw.
    }
  }
  SpreadsheetApp.getUi().alert('Applied QA fixes on '+touched+' tabs (collections only).');
}
