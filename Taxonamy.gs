function GHM_Fill_Taxonomy_From_Mapping(){
  const ss=SpreadsheetApp.getActive();
  const mapSh=ss.getSheetByName("Taxonomy_Mapping");
  if(!mapSh){ SpreadsheetApp.getUi().alert("No 'Taxonomy_Mapping' sheet found."); return; }

  // Read mapping table
  const lastRow=mapSh.getLastRow(), lastCol=mapSh.getLastColumn();
  if(lastRow<2 || lastCol<1){ SpreadsheetApp.getUi().alert("Taxonomy_Mapping is empty."); return; }
  const hdr=mapSh.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||"").trim().toLowerCase());
  const colIdx = name => hdr.indexOf(name);
  const iSeries = colIdx("series"), iChar = colIdx("character"), iVar = colIdx("character_variant");
  const iCat = colIdx("category_id"), iSub = colIdx("subcategory_id"), iTags = colIdx("taxonomy_tags");
  if(iSeries<0||iChar<0||iVar<0||iCat<0){ SpreadsheetApp.getUi().alert("Taxonomy_Mapping needs headers: Series, Character, Character_Variant, category_id (and optionally subcategory_id, taxonomy_tags)."); return; }
  const mapRows=mapSh.getRange(2,1,lastRow-1,lastCol).getValues();
  const norm=s=>(s||"").toString().trim().toLowerCase();
  const key=(S,C,V)=>`${norm(S)}|${norm(C)}|${norm(V)}`;
  const dict=new Map();
  mapRows.forEach(r=>{
    const k=key(r[iSeries], r[iChar], r[iVar]);
    dict.set(k, {
      category: r[iCat]||"",
      subcat: (iSub>=0? r[iSub] : "")||"",
      tags:   (iTags>=0? r[iTags]: "")||""
    });
  });

  // Fill on every collection tab
  const SKIP=new Set(["All","Collections_Index","Characters_Index","Series_Character_Matrix","Traits_Dictionary","TOC","Control","Globals","_GHM_LISTS_","Taxonomy_Categories","Taxonomy_Mapping","_GHM_DEBUG_","_GHM_BU_AUDIT_","GHM_CONTROL_APPLY_REPORT","Compact_View","Metadata_QC","Marketplace_Preview","ERC1155 - Editions"]);
  ss.getSheets().forEach(sh=>{
    if (SKIP.has(sh.getName()) || sh.getLastColumn()<1) return;
    const head=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||"").trim().toLowerCase());
    const col = name => head.indexOf(name);
    const cSeries = col("series"), cChar = col("character"), cVar = col("character_variant");
    const cCat = col("category_id"), cSub = col("subcategory_id"), cTags = col("taxonomy_tags");
    if (cSeries<0 || cChar<0 || cVar<0) return; // skip tabs without the key columns

    // Ensure target columns exist
    const ensure = (label)=>{ if (col(label)>=0) return col(label); const nc=sh.getLastColumn()+1; sh.getRange(1,nc).setValue(label); head[nc-1]=label; return nc-1; };
    const Cc = (cCat>=0? cCat : ensure("category_id"));
    const Cs = (cSub>=0? cSub : ensure("subcategory_id"));
    const Ct = (cTags>=0? cTags : ensure("taxonomy_tags"));

    const rows=Math.max(0, sh.getLastRow()-1); if (!rows) return;
    const data=sh.getRange(2,1,rows,sh.getLastColumn()).getValues();
    const catCol=[], subCol=[], tagCol=[];
    for (let r=0;r<rows;r++){
      const S=data[r][cSeries], C=data[r][cChar], V=data[r][cVar];
      const d=dict.get(key(S,C,V));
      const curCat = data[r][Cc], curSub = data[r][Cs], curTags = data[r][Ct];
      catCol.push([curCat? curCat : (d? d.category : "")]);
      subCol.push([curSub? curSub : (d? d.subcat  : "")]);
      tagCol.push([curTags? curTags : (d? d.tags   : "")]);
    }
    sh.getRange(2,Cc+1,rows,1).setValues(catCol);
    sh.getRange(2,Cs+1,rows,1).setValues(subCol);
    sh.getRange(2,Ct+1,rows,1).setValues(tagCol);
  });

  SpreadsheetApp.getActive().toast("Filled taxonomy columns from Taxonomy_Mapping (blanks only).", "GHM", 6);
}
