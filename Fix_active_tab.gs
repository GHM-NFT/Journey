// One-tab fixer: adds schema_version, moves "v1.x" out of Style, merges S -> Series.
// ES5-safe, no external helpers needed.
function RUN_buildSchemaAndFixActive(){
  var sh = SpreadsheetApp.getActiveSheet();
  var HEADER_ROW = 1;

  function readHeaders(){
    var lastCol = sh.getLastColumn();
    if (!lastCol) return [];
    var vals = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    var out = [];
    for (var i=0;i<vals.length;i++){ out.push(String(vals[i]||'').trim()); }
    return out;
  }
  function findIdx(name, H){
    for (var i=0;i<H.length;i++){ if (H[i]===name) return i; }
    return -1;
  }
  function refresh(){ return readHeaders(); }

  var H = readHeaders();
  var lastRow = sh.getLastRow();

  // 1) Ensure schema_version exists (place near admin if possible, else append)
  var iSchema = findIdx('schema_version', H);
  if (iSchema === -1){
    var insertAfter = -1;
    var prefs = ['last_modified','ready_for_export','qc_notes','qc_status'];
    for (var p=0;p<prefs.length;p++){
      var j = findIdx(prefs[p], H);
      if (j !== -1){ insertAfter = j; break; }
    }
    if (insertAfter === -1) insertAfter = H.length - 1; // append at end
    sh.insertColumnAfter(insertAfter+1);
    sh.getRange(HEADER_ROW, insertAfter+2).setValue('schema_version');
    H = refresh();
    iSchema = findIdx('schema_version', H);
  }

  // 2) Move version-like values from Style -> schema_version (clear Style)
  var iStyle = findIdx('Style', H);
  if (iStyle !== -1 && lastRow > 1){
    var styleR = sh.getRange(2, iStyle+1, lastRow-1, 1);
    var schemR = sh.getRange(2, iSchema+1, lastRow-1, 1);
    var styleV = styleR.getValues();
    var schemV = schemR.getValues();
    var changedS = false, changedT = false;

    for (var r=0;r<styleV.length;r++){
      var s = String(styleV[r][0]||'').trim();
      var v = String(schemV[r][0]||'').trim();
      // looks like "v1", "v1.1", "V2.0.3" etc.
      if (/^v\d+(\.\d+)*$/i.test(s) && !v){
        schemV[r][0] = s;  changedT = true;
        styleV[r][0] = ''; changedS = true;
      }
    }
    if (changedT) schemR.setValues(schemV);
    if (changedS) styleR.setValues(styleV);
  }

  // 3) Legacy "S" column -> Series (merge then drop)
  H = refresh();
  var iS = findIdx('S', H);
  var iSeries = findIdx('Series', H);
  if (iS !== -1){
    if (iSeries === -1){
      // rename header S -> Series
      sh.getRange(HEADER_ROW, iS+1).setValue('Series');
      H = refresh();
      iSeries = findIdx('Series', H);
    } else if (lastRow > 1){
      var sVals  = sh.getRange(2, iS+1, lastRow-1, 1).getValues();
      var serOut = sh.getRange(2, iSeries+1, lastRow-1, 1).getValues();
      var changed = false;
      for (var k=0;k<sVals.length;k++){
        var s2 = String(sVals[k][0]||'').trim();
        var v2 = String(serOut[k][0]||'').trim();
        if (!v2 && s2){ serOut[k][0] = s2; changed = true; }
      }
      if (changed) sh.getRange(2, iSeries+1, serOut.length, 1).setValues(serOut);
      // delete S column (use fresh index because positions shift)
    }
    // Find S again then delete
    H = refresh();
    iS = findIdx('S', H);
    if (iS !== -1) sh.deleteColumn(iS+1);
  }
}
