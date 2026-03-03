function RUN_fix_S_and_Style_active(){
  var sh = SpreadsheetApp.getActiveSheet();
  var H  = getHeaders_(sh,1);
  var last = sh.getLastRow();

  // --- 1) S -> Series (legacy deity_or_collection)
  var iS = indexOf_(H,'S');
  var iSeries = indexOf_(H,'Series');

  if (iS !== -1) {
    if (iSeries === -1) {
      // If there's no Series yet, just rename S to Series
      renameHeaderIfPresent_(sh, 1, 'S', 'Series');
      H = getHeaders_(sh,1);
      iSeries = indexOf_(H,'Series');
    } else {
      // Copy S values into empty Series cells
      if (last > 1){
        var sVals   = sh.getRange(2, iS+1, last-1, 1).getValues();
        var serVals = sh.getRange(2, iSeries+1, last-1, 1).getValues();
        var changed = false;
        for (var r=0;r<sVals.length;r++){
          var s = (sVals[r][0]||'').toString().trim();
          var v = (serVals[r][0]||'').toString().trim();
          if (!v && s){ serVals[r][0] = s; changed = true; }
        }
        if (changed) sh.getRange(2, iSeries+1, serVals.length, 1).setValues(serVals);
      }
      // Remove the stray S column
      sh.deleteColumn(iS+1);
      H = getHeaders_(sh,1);
    }
  }

  // --- 2) Fix Style that is all/mostly "v1.1"
  var iStyle = indexOf_(H,'Style');
  if (iStyle !== -1 && last > 1){
    var styleVals = sh.getRange(2, iStyle+1, last-1, 1).getValues();
    var nonblank = 0, v11 = 0;
    for (var i=0;i<styleVals.length;i++){
      var val = (styleVals[i][0]||'').toString().trim().toLowerCase();
      if (val){ nonblank++; if (val === 'v1.1') v11++; }
    }
    if (nonblank && v11/nonblank >= 0.8){
      // Try rebuild from Format/Medium + Stylisation
      H = getHeaders_(sh,1);
      var iFmt = indexOf_(H,'Format/Medium');
      var iSty = indexOf_(H,'Stylisation');
      if (iFmt !== -1 || iSty !== -1){
        var len = last-1;
        var F = (iFmt !== -1) ? sh.getRange(2, iFmt+1, len, 1).getValues() : null;
        var S = (iSty !== -1) ? sh.getRange(2, iSty+1, len, 1).getValues() : null;
        var out = [];
        for (var j=0;j<len;j++){
          var a = F ? (F[j][0]||'').toString().trim() : '';
          var b = S ? (S[j][0]||'').toString().trim() : '';
          out.push([ (a && b) ? (a+' '+b) : (a || b) ]);
        }
        sh.getRange(2, iStyle+1, len, 1).setValues(out);
      } else {
        // No sources to rebuild; just clear the "v1.1" entries
        for (var k=0;k<styleVals.length;k++){
          var vv = (styleVals[k][0]||'').toString().trim();
          if (vv.toLowerCase() === 'v1.1') styleVals[k][0] = '';
        }
        sh.getRange(2, iStyle+1, styleVals.length, 1).setValues(styleVals);
      }
    }
  }
}
