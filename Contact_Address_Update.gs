function GHM_RenameOrAdd_ContractAddress_All(){
  const ss=SpreadsheetApp.getActive();
  const SKIP=new Set([
    "All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary",
    "TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping",
    "_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT",
    "Compact_View","Compact_view","Metadata_QC","Marketplace_Preview","ERC1155 - Editions",
    "_GHM_BODY_AUDIT_","_GHM_COMPACT_AUDIT_","_GHM_CANON_REPORT_","_GHM_SHOWONLY_REPORT_"
  ]);
  const norm=s=>(s||"").toString().trim().toLowerCase();
  ss.getSheets().forEach(sh=>{
    if (SKIP.has(sh.getName()) || sh.getLastColumn()<1) return;
    const lastCol=sh.getLastColumn();
    const headers=sh.getRange(1,1,1,lastCol).getValues()[0];
    const hNorm=headers.map(norm);

    // If "Contract" exists, rename it to "Contract address"
    let i=hNorm.indexOf("contract");
    if (i>=0){ sh.getRange(1,i+1).setValue("Contract address"); return; }

    // Else if a variant exists, rename it
    const variants=["contract address","contract_address","collection address","collection_address"];
    for (const v of variants){
      const j=hNorm.indexOf(v);
      if (j>=0){ sh.getRange(1,j+1).setValue("Contract address"); return; }
    }

    // Else add a new "Contract address" column after contract_factory (or append)
    const k=hNorm.indexOf("contract_factory");
    if (k>=0){ sh.insertColumnAfter(k+1); sh.getRange(1,k+2).setValue("Contract address"); }
    else { sh.insertColumnAfter(lastCol); sh.getRange(1,lastCol+1).setValue("Contract address"); }
  });
}
