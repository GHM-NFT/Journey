// Fix obvious legacy headers + ensure core fields on the ACTIVE sheet only
function RUN_fixLegacyHeadersActive(){
  var sh = SpreadsheetApp.getActiveSheet();
  var H = getHeaders_(sh,1); // uses helpers already in your project

  // Common renames to canonical
  renameHeaderIfPresent_(sh, 1, 'Title/Name', 'name_final');
  renameHeaderIfPresent_(sh, 1, 'License_URL', 'license_url');
  renameHeaderIfPresent_(sh, 1, 'Contract address', 'contract_address');

  // Poster Image -> cid_Poster_Image (migrate values if target empty) then drop old
  H = getHeaders_(sh,1);
  if (indexOf_(H,'Poster Image') !== -1){
    ensureColumnExists_(sh, 1, 'cid_Poster_Image');
    migrateIfEmpty_(sh, 1, 'Poster Image', 'cid_Poster_Image');
    dropHeaderIfExists_(sh, 1, 'Poster Image');
  }

  // Ensure these exist (we’ll place/order them later)
  var must = ['Style','Tier','myth_scene','meaning_line','caption_300','symbols'];
  for (var i=0;i<must.length;i++){
    H = getHeaders_(sh,1);
    if (indexOf_(H, must[i]) === -1){
      ensureColumnExists_(sh, 1, must[i]);
    }
  }
}
