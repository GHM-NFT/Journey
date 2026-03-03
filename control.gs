function GHM_Control_Setup(){
  const ss = SpreadsheetApp.getActive();
  const name = "Control";

  const defaults = [
    ["Setting",                 "Value",                         "Notes"],
    ["project_name",            "GHM – Mythic Icons",            "Display only"],
    ["default_locale",          "en",                            "On-chain JSON language (recommended: en)"],
    ["secondary_languages",     "zh-Hans",                       "Comma-separated: e.g. zh-Hans,el,ja"],
    ["onchain_lang",            "en",                            "Keep on-chain in English; localize the web view"],
    ["external_url_base",       "https://ghm.art/nft",           "Base for external_url; slug will append"],
    ["operator_filter_default", "on",                            "on/off (listings compatibility stance)"],
    ["operator_policy_note",    "Respect creator royalties; allow major marketplaces.", "Short policy note"],
    ["freeze_after_days",       "7",                             "QA window (days) before freezing metadata"],
    ["standard_default",        "721",                           "721 for 1/1; 1155 for editions"],
    ["edition_size_default",    "",                              "Optional default for editions"],
    ["contract_factory",        "OpenSea",                       "Manifold, Zora, OpenSea, Custom"],
    ["collection_address",      "",                              "Canonical collection address (per series if needed)"],
    ["collection_slug",         "",                              "Platform slug (for URLs)"],
    ["license_url_default",     "",                              "Default license link (applied if blank)"],
    ["background_color_default","",                              "HEX (no #). e.g. FFFFFF"]
  ];

  // Create or get the Control sheet
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    // Fully clear content, formatting, and ANY old validations
    sh.clear(); // contents + formats
    const maxR = Math.max(sh.getMaxRows(), 1000);
    const maxC = Math.max(sh.getMaxColumns(), 26);
    sh.getRange(1,1, maxR, maxC).clearDataValidations();
  }

  // Write fresh table
  sh.getRange(1,1,defaults.length, defaults[0].length).setValues(defaults);
  sh.setFrozenRows(1);

  // Style header/body
  const HEADER_BG = "#1f2937", HEADER_FG = "#ffffff", CONTROL_BG = "#e5eaff";
  const FONT = "Roboto Condensed";
  const cols = defaults[0].length;

  sh.getRange(1,1,1,cols)
    .setBackground(HEADER_BG).setFontColor(HEADER_FG)
    .setFontWeight("bold").setFontFamily(FONT).setFontSize(10);

  if (defaults.length>1){
    sh.getRange(2,1,defaults.length-1,cols)
      .setBackground(CONTROL_BG).setFontFamily(FONT).setFontSize(10);
  }

  // Recreate named lists cleanly (clear existing named ranges with same names)
  const removeNamed = (n)=>{
    const nr = ss.getNamedRanges().find(x=>x.getName()===n);
    if (nr) nr.remove();
  };
  removeNamed("GHM_Standard_List");
  removeNamed("GHM_Factory_List");

  // Hidden list sheet + named ranges
  let listSh = ss.getSheetByName("_GHM_LISTS_");
  if (!listSh) { listSh = ss.insertSheet("_GHM_LISTS_"); listSh.hideSheet(); }
  // Clear old lists in case they collide
  listSh.clear();

  // Write lists in two columns
  const std = ["721","1155"].map(v=>[v]);
  const fac = ["Manifold","Zora","OpenSea","Custom"].map(v=>[v]);
  listSh.getRange(1,1,std.length,1).setValues(std);
  listSh.getRange(1,2,fac.length,1).setValues(fac);

  ss.setNamedRange("GHM_Standard_List", listSh.getRange(1,1,std.length,1));
  ss.setNamedRange("GHM_Factory_List",  listSh.getRange(1,2,fac.length,1));

  SpreadsheetApp.getUi().alert("Control sheet reset ✅  — Fill values, then run ‘GHM Control → Apply Control → Tabs’. ");
}
