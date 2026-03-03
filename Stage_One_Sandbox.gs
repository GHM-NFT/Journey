/***************************************
 * GHM — Column Autos (No-Preview build)
 * - No preview_thumb / thumb_url logic
 * - Your manual "preview" column is ignored
 ***************************************/

/*** GLOBAL CONFIG ***/
var CONFIG = {
  BASE_SITE: 'https://godsheroesmyths.com/',
  IPFS_GATEWAY: 'https://cloudflare-ipfs.com/ipfs/',
  SOT_TAB_NAME: 'GHM_SoT'
};
/*** POLICY QA / AUTOFILL CONFIG (keep exactly one copy) ***/
if (typeof POLICY_FIELDS === 'undefined') {
  // Guarded declaration to avoid "already declared" errors across files
  var POLICY_FIELDS = [
    'chain','operator_filter','royalty_bps',
    'reveal_mode','reveal_date','freeze_policy','freeze_date',
    'primary_mint_platform','est_gas_mint','est_gas_batch',
    'royalty_recipient','license_type','license_url','terms_url','physical_terms_url',
    'canonical_domain_or_ENS','marketplace_targets','target_fiat_price'
  ];
}


const CONTROL = {
  TAB: 'Control_GHM',
  ENFORCE_WRITE: true,    // fill only when Manifest cell is empty
  WARN_AUTOFILL: true,    // note 'autofilled: <field>' in warnings
  AUTOFILL: {
    chain:true, operator_filter:true, royalty_bps:true,
    reveal_mode:true, reveal_date:true, freeze_policy:true, freeze_date:true,
    primary_mint_platform:true, est_gas_mint:true, est_gas_batch:true,
    royalty_recipient:false, license_type:false, license_url:false, terms_url:false, physical_terms_url:false,
    canonical_domain_or_ENS:false, marketplace_targets:false, target_fiat_price:false
  }
};

/*** UTILITIES ***/
function getHeaderRowValues_(sh){
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(v){return String(v).trim();});
}
function getHeaderMap_(sh){
  var row = getHeaderRowValues_(sh), map = {};
  row.forEach(function(h,i){ if(h) map[h]=i+1; });
  return map;
}
function headerMap_(sh){ return getHeaderMap_(sh); } // legacy alias
function getOrCreateCol_(sh, header){
  var H = getHeaderMap_(sh);
  if (H[header]) return H[header];
  var c = sh.getLastColumn()+1;
  sh.getRange(1,c).setValue(header);
  return c;
}
function A1_(col){ var s='',n=col; while(n>0){var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=(n-m-1)/26;} return s; }
function driveUrl_(fileId, mode){ return fileId ? ('https://drive.google.com/uc?export='+(mode||'view')+'&id='+fileId) : ''; }

/*** GUARDS ***/
function normalizeTier_(sh, H){
  if (!H['Tier']) return;
  var rows = Math.max(0, sh.getLastRow()-1); if (!rows) return;
  var rng = sh.getRange(2,H['Tier'],rows,1), vals=rng.getValues();
  var CANON=['Mythic Icons','Signature Editions','Companion Pieces','Limited Edition Prints','Relics'];
  var map = new Map([
    ['mythic','Mythic Icons'],['mythic icons','Mythic Icons'],
    ['signature','Signature Editions'],['signatures','Signature Editions'],['signature editions','Signature Editions'],
    ['companion','Companion Pieces'],['companion pieces','Companion Pieces'],
    ['limited prints','Limited Edition Prints'],['limited edition prints','Limited Edition Prints'],['print','Limited Edition Prints'],
    ['relic','Relics'],['relics','Relics']
  ]);
  var canonSet = new Set(CANON.map(function(s){return s.toLowerCase();}));
  for (var i=0;i<rows;i++){
    var v=String(vals[i][0]||'').trim(); if(!v) continue;
    var k=v.toLowerCase(); if (canonSet.has(k)) continue;
    if (map.has(k)) vals[i][0]=map.get(k);
  }
  rng.setValues(vals);
}
function clearLangErrors_(sh, H, prefixes){
  var rows = Math.max(0, sh.getLastRow()-1); if(!rows) return;
  var headers = getHeaderRowValues_(sh);
  var ERR=/^#(N\/A|REF|VALUE|ERROR|DIV\/0|NAME|NULL)/i;
  function isNonEN(h){ var l=h.toLowerCase(); var starts=prefixes.some(function(p){return l.startsWith(p);}); var isEn=/(_en$|_en\b)/i.test(l); return starts && !isEn; }
  headers.forEach(function(h,idx){
    if(!isNonEN(h)) return;
    var rng=sh.getRange(2,idx+1,rows,1), vals=rng.getValues(), dirty=false;
    for (var i=0;i<rows;i++){ var v=vals[i][0]; if (typeof v==='string' && ERR.test(v)){ vals[i][0]=''; dirty=true; } }
    if (dirty) rng.setValues(vals);
  });
}
function enforceHex_(sh,H,name){
  var c=H[name]; if(!c) return; var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var rng=sh.getRange(2,c,rows,1), vals=rng.getValues(), HEX=/^#[0-9A-Fa-f]{6}$/; var dirty=false;
  for (var i=0;i<rows;i++){ var v=String(vals[i][0]||'').trim(); if(!v) continue; if(!HEX.test(v)){ vals[i][0]=''; dirty=true; } }
  if (dirty) rng.setValues(vals);
}
function enforceTokenRange_(sh,H,name){
  var c=H[name]; if(!c) return; var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var rng=sh.getRange(2,c,rows,1), vals=rng.getValues(), RANGE=/^\d+(-\d+)?$/; var dirty=false;
  for (var i=0;i<rows;i++){ var v=String(vals[i][0]||'').trim(); if(!v) continue; if(!RANGE.test(v)){ vals[i][0]=''; dirty=true; } }
  if (dirty) rng.setValues(vals);
}

/*** DERIVATIONS (no preview) ***/
// Media URLs (IPFS preferred; Drive fallback)
function fillMediaUrls_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var colImgUrl=H['image_url']||getOrCreateCol_(sh,'image_url');
  var colAnimUrl=H['animation_url']||getOrCreateCol_(sh,'animation_url');
  var imgCid=H['image_cid'], animCid=H['animation_cid'], imgDrv=H['drive_image_id'], animDrv=H['drive_animation_id'];
  var gw=CONFIG.IPFS_GATEWAY, outImg=[], outAni=[];
  for (var r=0;r<rows;r++){
    var ic=imgCid?String(sh.getRange(2+r,imgCid).getValue()||'').trim():'';
    var ac=animCid?String(sh.getRange(2+r,animCid).getValue()||'').trim():'';
    var idI=imgDrv?String(sh.getRange(2+r,imgDrv).getValue()||'').trim():'';
    var idA=animDrv?String(sh.getRange(2+r,animDrv).getValue()||'').trim():'';
    outImg.push([ ic?gw+ic : (idI?driveUrl_(idI,'view'):'') ]);
    outAni.push([ ac?gw+ac : (idA?driveUrl_(idA,'view'):'') ]);
  }
  sh.getRange(2,colImgUrl,rows,1).setValues(outImg);
  sh.getRange(2,colAnimUrl,rows,1).setValues(outAni);
}
// External URL (built)
function fillExternalUrlBuilt_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows||!H['slug']) return;
  var outCol=H['external_url_built']||getOrCreateCol_(sh,'external_url_built');
  var slugs=sh.getRange(2,H['slug'],rows,1).getValues();
  var out=slugs.map(function(r){var s=String(r[0]||'').trim(); return [s?CONFIG.BASE_SITE+s:''];});
  sh.getRange(2,outCol,rows,1).setValues(out);
}
// Filenames (json / edition / media)
function fillFilenames_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows||!H['slug']) return;
  var colJson=H['json_filename']||getOrCreateCol_(sh,'json_filename');
  var colEd=H['edition_string']||getOrCreateCol_(sh,'edition_string');
  var colMedia=H['media_filename']||getOrCreateCol_(sh,'media_filename');
  var slugs=sh.getRange(2,H['slug'],rows,1).getValues();
  var edCol=H['edition_size']?sh.getRange(2,H['edition_size'],rows,1).getValues():[];
  var mimeC=H['image_mime']?sh.getRange(2,H['image_mime'],rows,1).getValues():[];
  function extFromMime(m){m=(m||'').toLowerCase();
    if(m.indexOf('png')>-1)return'.png';
    if(m.indexOf('jpeg')>-1||m.indexOf('jpg')>-1)return'.jpg';
    if(m.indexOf('webp')>-1)return'.webp';
    if(m.indexOf('gif')>-1)return'.gif';
    if(m.indexOf('mp4')>-1)return'.mp4';
    if(m.indexOf('quicktime')>-1||m.indexOf('mov')>-1)return'.mov';
    return'';}
  var outJ=[],outE=[],outM=[];
  for (var i=0;i<rows;i++){
    var slug=String(slugs[i][0]||'').trim();
    var ed=H['edition_size']?Number(edCol[i][0]||''):'';
    var ext=H['image_mime']?extFromMime(mimeC[i][0]):'';
    outJ.push([slug?slug+'.json':'']);
    outE.push([ed?('Edition of '+ed):'']);
    outM.push([slug?slug+ext:'']);
  }
  sh.getRange(2,colJson,rows,1).setValues(outJ);
  sh.getRange(2,colEd,rows,1).setValues(outE);
  sh.getRange(2,colMedia,rows,1).setValues(outM);
}
// Alt-text EN
function fillAltTextEn_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var outCol=H['alt_text_en_auto']||getOrCreateCol_(sh,'alt_text_en_auto');
  function pick(h){return H[h]?sh.getRange(2,H[h],rows,1).getValues():[];}
  var Ch=pick('Character'),V=pick('Character_Variant'),St=pick('Style'),Fr=pick('Frame'),Cw=pick('Colorway'),Ml=pick('meaning_line'),T=pick('Tier');
  var out=[];
  for (var i=0;i<rows;i++){
    var c=Ch[i]&&Ch[i][0]||'', v=V[i]&&V[i][0]?(' ('+V[i][0]+')'):'', st=St[i]&&St[i][0]||'',
        fr=Fr[i]&&Fr[i][0]?(' with '+Fr[i][0]):'', cw=Cw[i]&&Cw[i][0]?(' '+Cw[i][0]+' palette'):'',
        ml=Ml[i]&&Ml[i][0]?('; '+Ml[i][0]):'', t=T[i]&&T[i][0]?('. '+T[i][0]+' edition.'):'.';
    out.push([ c?('Round canvas portrait of '+c+v+', '+st+fr+','+cw+ml+t):'' ]);
  }
  sh.getRange(2,outCol,rows,1).setValues(out);
}
// Alt-text zh-Hans (only if column exists)
function fillAltTextZH_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var col=H['alt_text_zh-Hans_auto']; if(!col) return;
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  function pick(h){return H[h]?sh.getRange(2,H[h],rows,1).getValues():[];}
  var Ch=pick('Character'),V=pick('Character_Variant'),St=pick('Style'),Fr=pick('Frame'),Cw=pick('Colorway'),Ml=pick('meaning_line'),T=pick('Tier');
  var out=[];
  for (var i=0;i<rows;i++){
    var c=Ch[i]&&Ch[i][0]||'', v=V[i]&&V[i][0]?('（'+V[i][0]+'）'):'', st=St[i]&&St[i][0]||'',
        fr=Fr[i]&&Fr[i][0]?('，配'+Fr[i][0]):'', cw=Cw[i]&&Cw[i][0]?('，'+Cw[i][0]+'配色'):'',
        ml=Ml[i]&&Ml[i][0]?('；'+Ml[i][0]):'', t=T[i]&&T[i][0]?('。'+T[i][0]+'版本。'):'。';
    out.push([ c?('圆形画布肖像：'+c+v+'，'+st+fr+cw+ml+t):'' ]);
  }
  sh.getRange(2,col,rows,1).setValues(out);
}
// Slug (unique; append -token_id if dup)
function fillSlug_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var outCol=H['slug']||getOrCreateCol_(sh,'slug');
  function pick(h){return H[h]?sh.getRange(2,H[h],rows,1).getValues():null;}
  var Series=pick('Series'),Char=pick('Character'),Var=pick('Character_Variant'),Scene=pick('myth_scene'),Style=pick('Style'),Tok=pick('token_id');
  function slugify(s){return s.toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/-+/g,'-').replace(/^-|-$/g,'');}
  var seen=new Map(), out=[];
  for (var i=0;i<rows;i++){
    var parts=[Series&&Series[i]&&Series[i][0],Char&&Char[i]&&Char[i][0],Var&&Var[i]&&Var[i][0],Scene&&Scene[i]&&Scene[i][0],Style&&Style[i]&&Style[i][0]]
      .map(function(s){return String(s||'').trim();}).filter(Boolean);
    var base=slugify(parts.join('-')); if(!base){out.push(['']); continue;}
    var final=base;
    if (seen.has(base)){ var tok=Tok&&Tok[i]&&Tok[i][0]; final=tok?(base+'-'+tok):(base+'-dup'+(seen.get(base)+1)); }
    seen.set(base,(seen.get(base)||0)+1);
    out.push([final]);
  }
  sh.getRange(2,outCol,rows,1).setValues(out);
}
// Hook line
function fillHookLine_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var outCol=H['hook_line']||getOrCreateCol_(sh,'hook_line');
  function pick(h){return H[h]?sh.getRange(2,H[h],rows,1).getValues():[];}
  var n=pick('name_final'), c=pick('Character'), v=pick('Character_Variant'), st=pick('Style'), t=pick('Tier'), ms=pick('myth_scene');
  var out=[];
  for (var i=0;i<rows;i++){
    var s1=String(n[i]&&n[i][0]||'').trim();
    var s2=[c[i]&&c[i][0], v[i]&&v[i][0], st[i]&&st[i][0], t[i]&&t[i][0]].filter(Boolean).join(' · ');
    var s3=String(ms[i]&&ms[i][0]||'').trim();
    out.push([ [s1,s2,s3].filter(Boolean).join(' — ').substring(0,140) ]);
  }
  sh.getRange(2,outCol,rows,1).setValues(out);
}
// Attributes JSON
function fillTraitJson_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var outCol=H['attributes_json']||getOrCreateCol_(sh,'attributes_json');
  var fields=['Pantheon','Series','Character','Character_Variant','myth_scene','Style','Tier','edition_size','Colorway','Frame'];
  var picks={}; fields.forEach(function(f){picks[f]=H[f]?sh.getRange(2,H[f],rows,1).getValues():[];});
  // prefer leftmost 'symbols'
  var symCol=(function(){ var hdrs=getHeaderRowValues_(sh); for (var i=0;i<hdrs.length;i++){ if (hdrs[i]==='symbols') return i+1; } return null; })();
  var symVals= symCol? sh.getRange(2,symCol,rows,1).getValues() : [];
  var out=[];
  for (var i=0;i<rows;i++){
    var attrs=[];
    function add(trait,val){ var v=String(val||'').trim(); if(!v) return; attrs.push({trait_type:trait, value:v}); }
    add('Pantheon',picks['Pantheon'][i]&&picks['Pantheon'][i][0]);
    add('Series',picks['Series'][i]&&picks['Series'][i][0]);
    add('Character',picks['Character'][i]&&picks['Character'][i][0]);
    add('Variant/Location',picks['Character_Variant'][i]&&picks['Character_Variant'][i][0]);
    add('Myth Scene',picks['myth_scene'][i]&&picks['myth_scene'][i][0]);
    add('Style/Format',picks['Style'][i]&&picks['Style'][i][0]);
    add('Tier',picks['Tier'][i]&&picks['Tier'][i][0]);
    add('Edition Size',picks['edition_size'][i]&&picks['edition_size'][i][0]);
    String(symVals[i]&&symVals[i][0]||'').split(/[,;|]/).map(function(s){return s.trim();}).filter(Boolean)
      .forEach(function(s){ attrs.push({trait_type:'Symbols', value:s}); });
    add('Colorway',picks['Colorway'][i]&&picks['Colorway'][i][0]);
    add('Frame',picks['Frame'][i]&&picks['Frame'][i][0]);
    out.push([JSON.stringify(attrs)]);
  }
  sh.getRange(2,outCol,rows,1).setValues(out);
}
// Lengths & status lights
function fillLengthsAndStatus_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var SPEC={
    name_final:[60,80], slug:[100,120], description_en:[300,450], caption_300:[300,450],
    Series:[24,30], Character:[24,30], Character_Variant:[24,30], myth_scene:[24,30], Style:[24,30], Tier:[24,30], edition_size:[24,30]
  };
  Object.keys(SPEC).forEach(function(h){
    if(!H[h]) return; var g=SPEC[h][0], a=SPEC[h][1];
    var lenCol=getOrCreateCol_(sh,'len_'+h), stCol=getOrCreateCol_(sh,'status_'+h);
    var vals=sh.getRange(2,H[h],rows,1).getValues().map(function(r){return String(r[0]||'');});
    var lenOut=[], stOut=[];
    for (var i=0;i<rows;i++){ var L=vals[i].length; lenOut.push([L]); stOut.push([ L<=g?'GREEN':(L<=a?'AMBER':'RED') ]); }
    sh.getRange(2,lenCol,rows,1).setValues(lenOut);
    sh.getRange(2,stCol,rows,1).setValues(stOut);
  });
}
// STRICT: RED lengths block JSON emission
function fillJsonGateAndWarnings_(_sh){
  var sh = _sh || SpreadsheetApp.getActiveSheet();
  var H  = getHeaderMap_(sh);
  var rows = Math.max(0, sh.getLastRow()-1); if (!rows) return;

  var willCol = H['json_will_emit'] || getOrCreateCol_(sh,'json_will_emit');
  var missCol = H['missing_fields'] || getOrCreateCol_(sh,'missing_fields');
  var warnCol = H['warnings']       || getOrCreateCol_(sh,'warnings');

  // helpers
  function need(h){ return H[h] ? sh.getRange(2,H[h],rows,1).getValues() : []; }
  var imgCid = need('image_cid');
  var licUrl = need('license_url');
  var bgHex  = need('background_hex');

  // pull all status_* columns we care about (if present)
  var STATUS_HEADERS = [
    'status_name_final','status_slug','status_description_en','status_caption_300',
    'status_Series','status_Character','status_Character_Variant','status_myth_scene',
    'status_Style','status_Tier','status_edition_size'
  ];
  var S = {}; // header -> values[]
  STATUS_HEADERS.forEach(function(h){
    if (H[h]) S[h] = sh.getRange(2,H[h],rows,1).getValues().map(function(r){ return String(r[0]||''); });
  });

  // build outputs
  var outWill=[], outMiss=[], outWarn=[];
  for (var i=0;i<rows;i++){
    var missing = [];
    if (!String(imgCid[i] && imgCid[i][0] || '').trim()) missing.push('image_cid');
    if (!String(licUrl[i] && licUrl[i][0] || '').trim()) missing.push('license_url');
    var hex = String(bgHex[i] && bgHex[i][0] || '').trim();
    if (hex && !/^#[0-9A-Fa-f]{6}$/.test(hex)) missing.push('background_hex');

    // strict length check: any RED?
    var redReasons = [];
    Object.keys(S).forEach(function(h){
      if (S[h][i].toUpperCase() === 'RED') redReasons.push(h.replace(/^status_/, ''));
    });
    var hasRed = redReasons.length > 0;

    // warnings (non-blocking notes)
    var warns = [];
    if (hasRed) warns.push('lengths in RED: ' + redReasons.join(', '));
    if (hex && !/^#[0-9A-Fa-f]{6}$/.test(hex)) warns.push('invalid background_hex');

    // final gate
    var ok = (missing.length === 0) && !hasRed;
    outWill.push([ok]);
    outMiss.push([missing.join(', ')]);
    outWarn.push([warns.join('; ')]);
  }

  sh.getRange(2,willCol,rows,1).setValues(outWill);
  sh.getRange(2,missCol,rows,1).setValues(outMiss);
  sh.getRange(2,warnCol,rows,1).setValues(outWarn);
}

// SoT status & pulls
function fillSoTStatus_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), ss=SpreadsheetApp.getActive();
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var outCol=getOrCreateCol_(sh,'sot_status');
  var sot=ss.getSheetByName(CONFIG.SOT_TAB_NAME);
  if(!sot){ sh.getRange(2,outCol,rows,1).setValues(Array(rows).fill([false])); return; }
  var H=getHeaderMap_(sh), Hs=getHeaderMap_(sot); if(!H['slug']||!Hs['slug']) return;
  var manifest=sh.getRange(2,H['slug'],rows,1).getValues().map(function(r){return String(r[0]||'').trim();});
  var sotRows=Math.max(0,sot.getLastRow()-1);
  var sotSlugs=sot.getRange(2,Hs['slug'],sotRows,1).getValues().map(function(r){return String(r[0]||'').trim();});
  var set=new Set(sotSlugs);
  sh.getRange(2,outCol,rows,1).setValues(manifest.map(function(s){return [set.has(s)];}));
}
function pullSoT_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), ss=SpreadsheetApp.getActive();
  var sot=ss.getSheetByName(CONFIG.SOT_TAB_NAME); if(!sot) return;
  var H=getHeaderMap_(sh), Hs=getHeaderMap_(sot); if(!Hs['slug']||!H['slug']) return;
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var manifest=sh.getRange(2,H['slug'],rows,1).getValues().map(function(r){return String(r[0]||'').trim();});
  var sotRows=Math.max(0,sot.getLastRow()-1);
  var data=sot.getRange(2,1,sotRows,sot.getLastColumn()).getValues();
  var hdrs=getHeaderRowValues_(sot), idx={}; hdrs.forEach(function(h,i){idx[h]=i;});
  var m=new Map(), si=idx['slug']; data.forEach(function(r){ var s=String(r[si]||'').trim(); if(s) m.set(s,r); });
  function setColIf(header,getter){
    var col=H[header]||getOrCreateCol_(sh,header), out=[];
    for (var i=0;i<rows;i++){ var row=m.get(manifest[i]); out.push([ row ? getter(row) : '' ]); }
    sh.getRange(2,col,rows,1).setValues(out);
  }
  if (idx['caption_long_en']!==undefined) setColIf('caption_long_en',function(r){return r[idx['caption_long_en']];});
  if (idx['research_notes']!==undefined) setColIf('research_notes',function(r){return r[idx['research_notes']];});
  if (idx['sources_bibliography']!==undefined) setColIf('sources_bibliography',function(r){return r[idx['sources_bibliography']];});
  if (idx['sot_doc_url']!==undefined){
    var col=H['sot_link']||getOrCreateCol_(sh,'sot_link'), out=[];
    for (var i=0;i<rows;i++){ var row=m.get(manifest[i]); var url=row?String(row[idx['sot_doc_url']]||''):''; out.push([ url ? '=HYPERLINK("'+url+'","SoT Doc")' : '' ]); }
    sh.getRange(2,col,rows,1).setValues(out);
  }
  fillSoTStatus_(sh);
}
// Checksum OK
function fillChecksumOk_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var outCol=H['checksum_ok']||getOrCreateCol_(sh,'checksum_ok');
  function pick(h){return H[h]?sh.getRange(2,H[h],rows,1).getValues():[];}
  var bytes=pick('image_bytes'), sha=pick('image_sha256');
  var out=[]; for (var i=0;i<rows;i++){ var ok=(Number(bytes[i]&&bytes[i][0]||0)>0)&&(!sha.length||String(sha[i]&&sha[i][0]||'').length>=40); out.push([ok]); }
  sh.getRange(2,outCol,rows,1).setValues(out);
}
// Taxonomy IDs
function fillTaxonomyIds_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), ss=SpreadsheetApp.getActive();
  var mapSh=ss.getSheetByName('Taxonomy_Map')||ss.getSheetByName('Taxonomy_mapping'); if(!mapSh) return;
  var mapVals=mapSh.getDataRange().getValues(), headers=mapVals.shift().map(function(h){return String(h).trim().toLowerCase();});
  var si=headers.indexOf('series_key'), ci=headers.indexOf('category_id'), sbi=headers.indexOf('subcategory_id'); if(si<0||ci<0||sbi<0) return;
  var lut=new Map(); mapVals.forEach(function(r){ var k=String(r[si]||'').toLowerCase(); if(k) lut.set(k,{cat:r[ci],sub:r[sbi]}); });
  var H=getHeaderMap_(sh), rows=Math.max(0,sh.getLastRow()-1); if(!rows||!H['Series']) return;
  var seriesVals=sh.getRange(2,H['Series'],rows,1).getValues(), catCol=H['category_id']||getOrCreateCol_(sh,'category_id'), subCol=H['subcategory_id']||getOrCreateCol_(sh,'subcategory_id');
  var outCat=[], outSub=[];
  for (var i=0;i<rows;i++){ var key=String(seriesVals[i][0]||'').toLowerCase(); var hit=lut.get(key); outCat.push([hit?hit.cat:'']); outSub.push([hit?hit.sub:'']); }
  sh.getRange(2,catCol,rows,1).setValues(outCat);
  sh.getRange(2,subCol,rows,1).setValues(outSub);
}
// Print-eligible (suggestion)
function fillPrintEligibleInferred_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows) return;
  var outCol=H['print_eligible_inferred']||getOrCreateCol_(sh,'print_eligible_inferred');
  function pick(h){return H[h]?sh.getRange(2,H[h],rows,1).getValues():[];}
  var T=pick('Tier'), W=pick('image_width'), B=pick('image_bytes'); var okTiers=new Set(['Mythic Icons','Signature Editions','Limited Edition Prints']);
  var out=[]; for (var i=0;i<rows;i++){ var yes=okTiers.has(String(T[i]&&T[i][0]||''))&&(Number(W[i]&&W[i][0]||0)>=4000||Number(B[i]&&B[i][0]||0)>=8000000); out.push([ yes?'YES':'NO' ]); }
  sh.getRange(2,outCol,rows,1).setValues(out);
}
// Provenance SHA
function fillProvenanceSha_(_sh){
  var sh=_sh||SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  var rows=Math.max(0,sh.getLastRow()-1); if(!rows||!H['slug']||!H['image_cid']) return;
  var colSha=H['provenance_block_sha']||getOrCreateCol_(sh,'provenance_block_sha');
  var slugs=sh.getRange(2,H['slug'],rows,1).getValues().map(function(r){return String(r[0]||'');});
  var cids= sh.getRange(2,H['image_cid'],rows,1).getValues().map(function(r){return String(r[0]||'');});
  var order= H['display_order']? sh.getRange(2,H['display_order'],rows,1).getValues().map(function(r,i){return {i:i,o:Number(r[0]||0)};}): slugs.map(function(_,i){return {i:i,o:i};});
  order.sort(function(a,b){return (a.o-b.o)||slugs[a.i].localeCompare(slugs[b.i]);});
  var joined=''; order.forEach(function(o){ var i=o.i; if(slugs[i]&&cids[i]) joined+=slugs[i]+'|'+cids[i]+'\n'; });
  var bytes=Utilities.newBlob(joined).getBytes(), digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,bytes);
  var hex=digest.map(function(b){return ('0'+(b&0xFF).toString(16)).slice(-2)}).join('');
  sh.getRange(2,colSha,rows,1).setValues(Array(rows).fill([hex]));
}

/*** MENU & ORCHESTRATORS ***/
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('GHM Autos')
    .addItem('1) Media URLs','fillMediaUrls_')
    .addItem('2) Preview URLs','fillPreviewUrls_')
    .addItem('3) Inject preview thumbs','fillPreviewThumb_')
    .addItem('4) External URL (built)','fillExternalUrlBuilt_')
    .addItem('5) Filenames','fillFilenames_')
    .addItem('6) Alt-text EN','fillAltTextEn_')
    .addItem('7) Alt-text zh-Hans (opt)','fillAltTextZH_')
    .addItem('8) Taxonomy IDs','fillTaxonomyIds_')
    .addItem('9) Lengths & Status','fillLengthsAndStatus_')
    .addItem('10) JSON Gate + Control','jsonGatePlusControl__menu')
    .addItem('11) Print-eligible inferred','fillPrintEligibleInferred_')
    .addItem('12) Provenance SHA','fillProvenanceSha_')
    .addSeparator()
    .addItem('SEO autos (meta title/desc/og)', 'fillSeoAutos_')  // ← the one you want
    .addSeparator()
    .addItem('Run All (Sandbox)','runAllSandbox_')
    .addToUi();
}

function runAllSandbox_(){
  const sh = SpreadsheetApp.getActiveSheet();
  const H  = getHeaderMap_(sh);

  // 1) Normalize & guards
  try{ normalizeTier_(sh, H); }catch(e){}
  try{ clearLangErrors_(sh, H, ['title_','description_','alt_text_']); }catch(e){}
  try{ enforceHex_(sh, H, 'background_hex'); }catch(e){}
  try{ enforceTokenRange_(sh, H, 'token_range'); }catch(e){}

  // 2) Derivations
  fillSlug_(sh, H);
  fillHookLine_(sh, H);
  fillAltTextEn_(sh);                 // gets H internally
  fillAltTextZH_(sh);                 // gets H internally (safe if column exists)
  fillTraitJson_(sh, H);
  fillTaxonomyIds_(sh);               // looks up Taxonomy_Map internally
  fillMediaUrls_(sh);                 // builds image_url / animation_url
  fillFilenames_(sh);                 // media_filename / json_filename / edition_string
  fillExternalUrlBuilt_(sh);          // per-token page
  fillLengthsAndStatus_(sh, H);
  fillJsonGateAndWarnings_(sh);       // STRICT gate (RED blocks)
   const ctl = getControl_();
  try { enforceControlPolicies_(sh, H, ctl); } catch(e){}
  try { fillSeoAutos_(sh); } catch(e){}
  // 3) Control enforcement (append warnings / optionally fill empties)
  try { enforceControlPolicies_(sh, H, ctl); } catch(e){}

  // 4) SoT & other greens
  fillSoTStatus_(sh, H);
  fillChecksumOk_(sh, H);
  fillPrintEligibleInferred_(sh);     // gets H internally
  fillProvenanceSha_(sh);             // gets H internally
  pullSoT_(sh, H);                    // caption_long_en / research_notes / sources_bibliography / sot_link
}

function RUN_updateAll_now(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const H  = getHeaderMap_(sh);

  // 1) Normalize & guards
  try{ normalizeTier_(sh, H); }catch(e){}
  try{ clearLangErrors_(sh, H, ['title_','description_','alt_text_']); }catch(e){}
  try{ enforceHex_(sh, H, 'background_hex'); }catch(e){}
  try{ enforceTokenRange_(sh, H, 'token_range'); }catch(e){}

  // 2) Derivations (same order as sandbox)
  fillSlug_(sh, H);
  fillHookLine_(sh, H);
  try{ fillAltTextEn_(sh); }catch(e){}
  try{ fillAltTextZH_(sh); }catch(e){}           // safe if zh column not present
  fillTraitJson_(sh, H);
  fillTaxonomyIds_(sh);                           // reads Taxonomy_Map / Taxonomy_mapping
  fillMediaUrls_(sh);                             // builds image_url / animation_url
  fillFilenames_(sh);                             // media_filename / json_filename / edition_string
  fillExternalUrlBuilt_(sh);                      // per-token page
  fillLengthsAndStatus_(sh, H);
  fillJsonGateAndWarnings_(sh);                   // STRICT: RED length = block
// AFTER fillJsonGateAndWarnings_(sh);
const ctl = getControl_();
try { enforceControlPolicies_(sh, H, ctl); } catch(e){}
try { fillSeoAutos_(sh); } catch(e){}

  // 3) Control enforcement (append warnings, and auto-fill empties if enabled)
  try { enforceControlPolicies_(sh, H, ctl); } catch(e){}

  // 4) SoT & other greens
  fillSoTStatus_(sh, H);
  fillChecksumOk_(sh, H);
  try{ fillPrintEligibleInferred_(sh); }catch(e){}
  try{ fillProvenanceSha_(sh); }catch(e){}
  pullSoT_(sh, H);                                // caption_long_en / research_notes / sources_bibliography / sot_link
}


/*** OPTIONAL ONE-OFF CLEANUP: remove legacy preview columns ***/
function GHM_RemoveLegacyPreviewColsOnce(){
  var sh=SpreadsheetApp.getActiveSheet(), H=getHeaderMap_(sh);
  ['preview_thumb','thumb_url','thumb_full_url'].forEach(function(h){
    var col=H[h]; if(col){ sh.deleteColumn(col); }
  });
}
function GHM_AuditSheets(){
  const ss = SpreadsheetApp.getActive();
  const outName = 'GHM_Audit';
  const existing = ss.getSheetByName(outName);
  if (existing) ss.deleteSheet(existing);
  const out = ss.insertSheet(outName);

  const header = [
    'sheet_name','rows','cols',
    'has_slug','has_attributes_json','has_json_will_emit',
    'has_image_cid','has_image_url',
    'has_category_id','has_subcategory_id',
    'has_preview','has_preview_thumb_cols',
    'has_GHM_SoT','has_Taxonomy_Map',
    'manifest_candidate_score',
    'first_40_headers'
  ];
  out.getRange(1,1,1,header.length).setValues([header]);

  const hasSheet = (name)=> !!(ss.getSheetByName(name));
  const sheets = ss.getSheets();
  let r = 2;

  sheets.forEach(sh=>{
    const rows = Math.max(0, sh.getLastRow()-1);
    const cols = sh.getLastColumn();
    if (cols === 0) return;

    const headers = sh.getRange(1,1,1,cols).getValues()[0].map(h=>String(h).trim());
    const H = Object.fromEntries(headers.map((h,i)=>[h,i+1]));
    const has = (h)=> !!H[h];

    // quick score to identify a manifest-like sheet
    let score = 0;
    if (has('slug')) score+=2;
    if (has('attributes_json')) score+=2;
    if (has('json_will_emit')) score+=1;
    if (has('category_id') && has('subcategory_id')) score+=1;
    if (has('image_cid') || has('image_url')) score+=1;

    const row = [
      sh.getName(), rows, cols,
      has('slug'), has('attributes_json'), has('json_will_emit'),
      has('image_cid'), has('image_url'),
      has('category_id'), has('subcategory_id'),
      has('preview'),
      (has('preview_thumb')||has('thumb_url')||has('thumb_full_url')),
      hasSheet('GHM_SoT') ? 'YES':'NO',
      (hasSheet('Taxonomy_Map')||hasSheet('Taxonomy_mapping')) ? 'YES':'NO',
      score,
      headers.slice(0,40).join(' | ')
    ];
    out.getRange(r,1,1,row.length).setValues([row]);
    r++;
  });

  out.autoResizeColumns(1, header.length);
  SpreadsheetApp.getUi().alert('GHM_Audit created. Open the tab to review.');
}
// 1) Create (or fix) the SoT tab with exact headers
function GHM_SoT_CreateOrFixHeaders(){
  const ss = SpreadsheetApp.getActive();
  const name = 'GHM_SoT';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  // wipe header row and set exact headers
  sh.getRange(1,1,1,sh.getLastColumn()||1).clearContent();
  const hdrs = ['slug','caption_long_en','research_notes','sources_bibliography','sot_doc_url'];
  sh.getRange(1,1,1,hdrs.length).setValues([hdrs]);
  sh.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('GHM_SoT headers are set.');
}

// 2) Seed N slugs from your Manifest into GHM_SoT so lookups have something to match
function GHM_SoT_Seed(n){
  const ss = SpreadsheetApp.getActive();
  const man = ss.getActiveSheet();                         // run this *from* your Manifest sheet
  const Hm = man.getRange(1,1,1,man.getLastColumn()).getValues()[0]
              .reduce((m,h,i)=>(m[String(h).trim()]=i+1,m),{});
  if (!Hm['slug']) { SpreadsheetApp.getUi().alert('No "slug" column on this sheet.'); return; }

  const sot = ss.getSheetByName('GHM_SoT') || ss.insertSheet('GHM_SoT');
  const Hs = sot.getRange(1,1,1,Math.max(5,sot.getLastColumn())).getValues()[0]
              .reduce((m,h,i)=>(m[String(h).trim()]=i+1,m),{});
  // ensure headers exist
  const needed = ['slug','caption_long_en','research_notes','sources_bibliography','sot_doc_url'];
  needed.forEach((h,i)=>{ if(!Hs[h]) { sot.getRange(1,i+1).setValue(h); Hs[h]=i+1; } });

  const rows = Math.max(0, man.getLastRow()-1);
  const take = Math.min(n||10, rows);
  const slugs = man.getRange(2,Hm['slug'],rows,1).getValues().map(r=>String(r[0]||'').trim());
  const uniq = Array.from(new Set(slugs.slice(0,take).filter(Boolean)));

  const existing = new Set(
    (sot.getLastRow()>1 ? sot.getRange(2,Hs['slug'],sot.getLastRow()-1,1).getValues() : [])
    .map(r=>String(r[0]||'').trim())
  );

  // append new rows for any seed slugs not already in SoT
  const app = uniq.filter(s=>!existing.has(s)).map(s=>[s,'','','','']);
  if (app.length){
    sot.getRange(sot.getLastRow()+1,1,app.length,5).setValues(app);
    SpreadsheetApp.getUi().alert('Seeded '+app.length+' slug(s) into GHM_SoT.');
  }else{
    SpreadsheetApp.getUi().alert('No new slugs to seed (SoT already has these).');
  }
}

// (Optional) 3) Debug report of matches vs. misses
function GHM_SoT_Debug(){
  const ss = SpreadsheetApp.getActive();
  const man = ss.getActiveSheet();
  const sot = ss.getSheetByName('GHM_SoT');
  if (!sot){ SpreadsheetApp.getUi().alert('GHM_SoT not found. Run GHM_SoT_CreateOrFixHeaders first.'); return; }

  const Hm = man.getRange(1,1,1,man.getLastColumn()).getValues()[0]
              .reduce((m,h,i)=>(m[String(h).trim()]=i+1,m),{});
  const Hs = sot.getRange(1,1,1,sot.getLastColumn()).getValues()[0]
              .reduce((m,h,i)=>(m[String(h).trim()]=i+1,m),{});
  if(!Hm['slug']||!Hs['slug']){ SpreadsheetApp.getUi().alert('Missing "slug" on one of the sheets.'); return; }

  const rows = Math.max(0, man.getLastRow()-1);
  const slugs = man.getRange(2,Hm['slug'],rows,1).getValues().map(r=>String(r[0]||'').trim());
  const sotRows = Math.max(0, sot.getLastRow()-1);
  const sotSlugs = sot.getRange(2,Hs['slug'],sotRows,1).getValues().map(r=>String(r[0]||'').trim());
  const set = new Set(sotSlugs);

  const outName='GHM_SoT_Debug';
  const old = ss.getSheetByName(outName); if(old) ss.deleteSheet(old);
  const dbg = ss.insertSheet(outName);
  dbg.getRange(1,1,1,2).setValues([['slug','found_in_GHM_SoT']]);
  const out = slugs.filter(Boolean).map(s=>[s, set.has(s)?'YES':'NO']);
  if (out.length) dbg.getRange(2,1,out.length,2).setValues(out);
  dbg.autoResizeColumns(1,2);
  SpreadsheetApp.getUi().alert('Debug sheet created.');
}
// ---- Control integration (read Control_GHM row 2) ----

function getControl_(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONTROL.TAB);
  if (!sh || sh.getLastRow() < 2) return {};
  const hdrs = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const vals = sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0];
  const out = {}; hdrs.forEach((h,i)=> out[h] = vals[i]);
  return out;
}

function enforceControlPolicies_(sh, H, ctl){
  const rows = Math.max(0, sh.getLastRow()-1); if (!rows) return;
  const warnCol = H['warnings'] || getOrCreateCol_(sh,'warnings');
  const warnVals = sh.getRange(2,warnCol,rows,1).getValues();

  const pick = (h)=> H[h] ? sh.getRange(2,H[h],rows,1).getValues().map(r=>r[0]) : Array(rows).fill('');

  // current values
  const vals = Object.fromEntries(POLICY_FIELDS.map(f=>[f, pick(f)]));

  let dirtyWarn=false;
  const dirty = Object.fromEntries(POLICY_FIELDS.map(f=>[f,false]));

  for (let i=0;i<rows;i++){
    const adds=[];
    POLICY_FIELDS.forEach(f=>{
      const cur = vals[f][i];
      const want = ctl ? ctl[f] : undefined;
      const allowAuto = CONTROL.ENFORCE_WRITE && CONTROL.AUTOFILL && CONTROL.AUTOFILL[f];

      if ((cur==='' || cur==null) && allowAuto && want!==undefined && want!==''){
        vals[f][i] = want; dirty[f]=true;
        if (CONTROL.WARN_AUTOFILL) adds.push('autofilled: '+f);
      } else if (cur!=='' && want!==undefined && want!=='' &&
                 String(cur).toLowerCase() !== String(want).toLowerCase()){
        adds.push(f+' ≠ Control_GHM');
      }
    });

    // sanity checks for your new BLUEs
    const rm = String(vals['reveal_mode']?.[i]||'').toLowerCase();
    const rd = String(vals['reveal_date']?.[i]||'').trim();
    if (rd && rm && rm!=='delayed') adds.push('reveal_date set but reveal_mode ≠ Delayed');

    const fp = String(vals['freeze_policy']?.[i]||'').toLowerCase();
    const fd = String(vals['freeze_date']?.[i]||'').trim();
    if (fd && !fp) adds.push('freeze_date set but freeze_policy empty');

    ['est_gas_mint','est_gas_batch','target_fiat_price'].forEach(n=>{
      const v = vals[n]?.[i];
      if (v!=='' && isNaN(Number(v))) adds.push(n+' not numeric');
    });

    if (adds.length){
      const prev = String(warnVals[i][0]||'').trim();
      warnVals[i][0] = prev ? prev+'; '+adds.join('; ') : adds.join('; ');
      dirtyWarn = true;
    }
  }

  if (dirtyWarn) sh.getRange(2,warnCol,rows,1).setValues(warnVals);
  if (CONTROL.ENFORCE_WRITE){
    POLICY_FIELDS.forEach(f=>{
      if (!dirty[f]) return;
      const col = H[f]; if (!col) return;
      const out = vals[f].map(v=>[v]);
      sh.getRange(2,col,rows,1).setValues(out);
    });
  }
}

function jsonGatePlusControl__menu(){
  const sh = SpreadsheetApp.getActiveSheet();
  const H  = getHeaderMap_(sh);
  fillJsonGateAndWarnings_(sh);             // JSON gate (strict)
  const ctl = getControl_();                // read Control_GHM
  try { enforceControlPolicies_(sh, H, ctl); } catch(e){}
}

function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GHM Autos')
    .addItem('1) Media URLs','fillMediaUrls_')
    .addItem('2) External URL (built)','fillExternalUrlBuilt_')
    .addItem('3) Filenames','fillFilenames_')
    .addItem('4) Alt-text EN','fillAltTextEn_')
    .addItem('5) Alt-text zh-Hans (opt)','fillAltTextZH_')
    .addItem('6) Taxonomy IDs','fillTaxonomyIds_')
    .addItem('7) Lengths & Status','fillLengthsAndStatus_')
    .addItem('8) JSON Gate + Control','jsonGatePlusControl__menu')
    .addItem('9) Print-eligible inferred','fillPrintEligibleInferred_')
    .addItem('10) Provenance SHA','fillProvenanceSha_')
    .addItem('11) SoT pulls (opt)','pullSoT_')
    .addSeparator()
    .addItem('SEO autos (meta title/desc/og)', 'fillSeoAutos_')   // ← add this line
    .addSeparator()
    .addItem('Run All (Sandbox)','runAllSandbox_')
    .addToUi();
}

function fillSeoAutos_(sh){
  sh = sh || SpreadsheetApp.getActiveSheet();
  const H = getHeaderMap_(sh);
  const rows = Math.max(0, sh.getLastRow()-1); if (!rows) return;

  const colTitleA = getOrCreateCol_(sh,'meta_title_auto');
  const colDescA  = getOrCreateCol_(sh,'meta_description_auto');
  const colOgA    = getOrCreateCol_(sh,'og_image_auto');

  const colTitleM = H['meta_title'] || null;
  const colDescM  = H['meta_description'] || null;

  const pick = (h)=> H[h] ? sh.getRange(2,H[h],rows,1).getValues().map(r=>String(r[0]||'')) : Array(rows).fill('');
  const name   = pick('name_final');
  const series = pick('Series');
  const cap300 = pick('caption_300');
  const desc   = pick('description_en');
  const imgurl = pick('image_url');

  const ctl = getControl_ ? getControl_() : {};
  const DOMAIN = String(ctl.canonical_domain_or_ENS || 'Gods • Heroes • Myths');

  function trunc(s, n){ s = String(s||''); return s.length<=n ? s : s.slice(0,n-1)+'…'; }
  function pickDesc(i){
    const a = String(cap300[i]||'').trim();
    const b = String(desc[i]||'').trim();
    return a || b || '';
  }

  const outT=[], outD=[], outO=[], mirrorT=[], mirrorD=[];
  for (let i=0;i<rows;i++){
    const t = trunc([name[i], series[i] ? `— ${series[i]}` : '', `| ${DOMAIN}`].filter(Boolean).join(' '), 60);
    const d = trunc(pickDesc(i).replace(/\s+/g,' '), 155);
    const o = imgurl[i] || String(ctl.og_image || '');

    outT.push([t]); outD.push([d]); outO.push([o]);

    // mirror into manual fields if blank (optional, safe)
    if (colTitleM) mirrorT.push([t]);
    if (colDescM)  mirrorD.push([d]);
  }
  sh.getRange(2,colTitleA,rows,1).setValues(outT);
  sh.getRange(2,colDescA,rows,1).setValues(outD);
  sh.getRange(2,colOgA,rows,1).setValues(outO);

  // Only write manual fields if blank
  if (colTitleM){
    const cur = sh.getRange(2,colTitleM,rows,1).getValues();
    for (let i=0;i<rows;i++) if (!String(cur[i][0]||'').trim()) cur[i][0]=mirrorT[i][0];
    sh.getRange(2,colTitleM,rows,1).setValues(cur);
  }
  if (colDescM){
    const cur = sh.getRange(2,colDescM,rows,1).getValues();
    for (let i=0;i<rows;i++) if (!String(cur[i][0]||'').trim()) cur[i][0]=mirrorD[i][0];
    sh.getRange(2,colDescM,rows,1).setValues(cur);
  }
}
function setSchemaFreezeMetadata() {
  const ss = SpreadsheetApp.getActive();
  const CONTROL = 'Control_GHM'; // change if different
  const sheet = ss.getSheetByName(CONTROL);
  if (!sheet) throw new Error(CONTROL + ' not found');

  // Row/column where you store single-value control items (adjust if different)
  // This script writes keys into column A and values into column B; it will upsert.
  const map = {
    'schema_version': 'v2025-10-17.freeze1',
    'schema_status': 'FROZEN_PENDING_QA',
    'schema_frozen_by': Session.getActiveUser().getEmail() || '<owner_email_here>',
    'schema_frozen_at': new Date().toISOString()
  };

  const data = sheet.getRange(1,1,sheet.getLastRow(),2).getValues();
  const out = {};
  data.forEach(r => { if (r[0]) out[r[0]] = r[1]; });

  // Upsert keys
  Object.keys(map).forEach(function(key){
    let rowIdx = data.findIndex(r => r[0] === key);
    if (rowIdx === -1) {
      // append
      sheet.appendRow([key, map[key]]);
    } else {
      sheet.getRange(rowIdx+1, 2).setValue(map[key]);
    }
  });

  SpreadsheetApp.getUi().alert('Schema freeze metadata set in ' + CONTROL);
}
function protectHeadersAndCriticalCols() {
  const ss = SpreadsheetApp.getActive();
  const MAIN = 'Manifest_GHM_Olympians'; // change to your manifest sheet name
  const sheet = ss.getSheetByName(MAIN);
  if (!sheet) throw new Error(MAIN + ' not found');

  const allowedEditors = ['mark@moda.digital']; // replace with your email(s)
  const CRITICAL_HEADERS = ['Series','Character','slug','token_id','contract','alt_text_en','image_cid','json_path'];

  // Protect header row (row 1)
  const headerRange = sheet.getRange(1,1,1,sheet.getLastColumn());
  const headerProtection = headerRange.protect().setDescription('Header protection: schema frozen');
  headerProtection.removeEditors(headerProtection.getEditors());
  headerProtection.addEditors(allowedEditors);

  // Protect each critical column (protect full column but leave first row editable maybe)
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  CRITICAL_HEADERS.forEach(function(hdr){
    const col = headers.indexOf(hdr);
    if (col >= 0) {
      const range = sheet.getRange(2, col+1, sheet.getMaxRows()-1); // protect below header
      const p = range.protect().setDescription('Protected: ' + hdr);
      p.removeEditors(p.getEditors());
      p.addEditors(allowedEditors);
    }
  });

  SpreadsheetApp.getUi().alert('Header + critical columns protected on ' + MAIN);
}
function addHiddenOpsFlagAndHideCols() {
  const ss = SpreadsheetApp.getActive();
  const MAIN = 'Manifest_GHM_Olympians';
  const sheet = ss.getSheetByName(MAIN);
  if (!sheet) throw new Error(MAIN + ' not found');

  const HIDDEN_COLUMNS = [
    'Renders','Thumbnails','cid_Previews','cid_Renders','cid_Thumbnails',
    'image_path','json_path','json_cid','unlockable_path','cid_master','cid_video','cid_vr','cid_coa','cid_print_token'
  ];

  // Ensure HIDDEN_OPS flag column exists (append to right if missing)
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let flagColIndex = headers.indexOf('HIDDEN_OPS');
  if (flagColIndex === -1) {
    flagColIndex = headers.length;
    sheet.getRange(1, flagColIndex+1).setValue('HIDDEN_OPS');
  }

  // Hide the listed columns (if found)
  headers.forEach(function(name, i){
    if (HIDDEN_COLUMNS.indexOf(name) !== -1) {
      sheet.hideColumns(i+1);
      // mark hidden in HIDDEN_OPS on the header row
      sheet.getRange(1, flagColIndex+1).setValue('HIDDEN_OPS - columns hidden');
    }
  });

  SpreadsheetApp.getUi().alert('Hidden ops columns processed.');
}
function applyDataValidationRules() {
  const ss = SpreadsheetApp.getActive();
  const MAIN = 'Manifest_GHM_Olympians';
  const sheet = ss.getSheetByName(MAIN);
  if (!sheet) throw new Error(MAIN + ' not found');

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  const tierCol = headers.indexOf('Tier'); // change if different
  const langCol = headers.indexOf('Planned_Langs'); // example
  const hexCol = headers.indexOf('background_hex');

  // Tier options
  const tiers = ['Mythic Icon','Signature Edition','Companion Piece','Limited Edition Print','Relic'];
  if (tierCol >= 0) {
    const range = sheet.getRange(2, tierCol+1, sheet.getMaxRows()-1);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(tiers, true).setAllowInvalid(false).setHelpText('Pick a Tier').build();
    range.setDataValidation(rule);
  }

  // Language codes dropdown
  const langs = ['en','ja','hi','it','zh-Hans'];
  if (langCol >= 0) {
    const range = sheet.getRange(2, langCol+1, sheet.getMaxRows()-1);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(langs, true).setAllowInvalid(true).setHelpText('Use planned lang codes').build();
    range.setDataValidation(rule);
  }

  // background_hex regex custom formula (applies to whole column)
  if (hexCol >= 0) {
    const colLetter = String.fromCharCode('A'.charCodeAt(0) + hexCol); // simple A..Z only; if >Z adjust (user can adapt)
    const formula = `=OR(ISBLANK(${colLetter}2),REGEXMATCH(${colLetter}2,\"^#([A-Fa-f0-9]{6})$\"))`;
    const range = sheet.getRange(2, hexCol+1, sheet.getMaxRows()-1);
    const rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(formula).setAllowInvalid(false).setHelpText('Must be #RRGGBB or blank').build();
    range.setDataValidation(rule);
  }

  SpreadsheetApp.getUi().alert('Data validation rules applied (adjust header names if needed).');
}
function runSmokeTest() {
  const ss = SpreadsheetApp.getActive();
  const MAIN = 'Manifest_GHM_Olympians';
  const REPORT = 'QA_Report';
  const N = 10;

  const sheet = ss.getSheetByName(MAIN);
  if (!sheet) throw new Error(MAIN + ' not found');

  // Fields to validate (ensure these headers exist)
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const H = {
    name_final: headers.indexOf('name_final'),
    slug: headers.indexOf('slug'),
    title_en: headers.indexOf('title_en'),
    alt_text_en: headers.indexOf('alt_text_en'),
    license_url: headers.indexOf('license_url'),
    background_hex: headers.indexOf('background_hex'),
    image_cid: headers.indexOf('image_cid'),
    image_url: headers.indexOf('image_url')
  };

  const rows = sheet.getRange(2,1,Math.min(N, sheet.getLastRow()-1), sheet.getLastColumn()).getValues();

  // Ensure report sheet exists
  let rep = ss.getSheetByName(REPORT);
  if (!rep) rep = ss.insertSheet(REPORT);
  rep.clear();
  rep.appendRow(['row','issue','field','value','notes']);

  rows.forEach(function(r, idx){
    const rowNum = idx + 2;
    // 1) name length <= 60
    const name = r[H.name_final] || '';
    if (name.length > 60) rep.appendRow([rowNum,'length','name_final',name,'>60 chars']);

    // 2) slug length <= 120
    const slug = r[H.slug] || '';
    if (slug.length > 120) rep.appendRow([rowNum,'length','slug',slug,'>120 chars']);

    // 3) title_en not empty
    const title = r[H.title_en] || '';
    if (!title) rep.appendRow([rowNum,'missing','title_en',title,'required']);

    // 4) alt_text_en not empty and <=125
    const alt = r[H.alt_text_en] || '';
    if (!alt) rep.appendRow([rowNum,'missing','alt_text_en',alt,'recommended']);
    if (alt.length > 125) rep.appendRow([rowNum,'length','alt_text_en',alt,'>125 chars']);

    // 5) json gate: license_url + background_hex valid
    const lic = r[H.license_url] || '';
    const hex = r[H.background_hex] || '';
    if (!lic) rep.appendRow([rowNum,'missing','license_url',lic,'JSON gate fails']);
    if (hex && !/^#([A-Fa-f0-9]{6})$/.test(hex)) rep.appendRow([rowNum,'format','background_hex',hex,'invalid pattern']);

    // 6) image_cid or image_url present
    const cid = r[H.image_cid] || '';
    const url = r[H.image_url] || '';
    if (!cid && !url) rep.appendRow([rowNum,'missing','image_cid/image_url','', 'no media reference']);

    // (Optional) If you have a JSON emitter function, call it here (replace 'emitJsonForRow' with your func)
    // try { var json = emitJsonForRow(rowNum); } catch(e) { rep.appendRow([rowNum,'error','json_emit','', e.toString()]); }
  });

  SpreadsheetApp.getUi().alert('Smoke test complete — open sheet: ' + REPORT);
}
// fillTitlesAndAlts.gs
function fillTitlesAndAlts() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = ''; // <-- Put exact tab name here or leave '' to use active sheet
  const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found or no active sheet.');

  // 1) Backup the sheet
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
  const backupName = (sheet.getName() + '_Backup_' + now).substring(0,99);
  sheet.copyTo(ss).setName(backupName);

  // 2) Locate headers
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idx = (name) => headers.indexOf(name);
  const H = {
    name_final: idx('name_final'),
    title_en: idx('title_en'),
    title: idx('title'),
    slug: idx('slug'),
    myth_scene: idx('myth_scene'),
    meaning_line: idx('meaning_line'),
    colorway: idx('colorway'),
    Style: idx('Style'),
    Character: idx('Character'),
    alt_text_en: idx('alt_text_en')
  };

  if (H.title_en < 0) throw new Error('Header "title_en" not found.');
  // Prepare data range
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data rows found.'); return; }
  const data = sheet.getRange(2,1,lastRow-1,headers.length).getValues();

  const MAX_TITLE = 60;
  const MAX_ALT = 125;
  const updates = []; // rows changed

  for (let i=0;i<data.length;i++){
    const rIdx = i + 2;
    const row = data[i];

    // BUILD title_en if empty
    let titleVal = row[H.title_en] ? String(row[H.title_en]).trim() : '';
    if (!titleVal) {
      // source precedence: name_final -> title -> slug-cleaned
      let source = '';
      if (H.name_final >=0 && row[H.name_final]) source = String(row[H.name_final]).trim();
      else if (H.title >=0 && row[H.title]) source = String(row[H.title]).trim();
      else if (H.slug >=0 && row[H.slug]) {
        source = String(row[H.slug]).replace(/[-_]+/g,' ').replace(/\s+\d+$/,'').trim();
      }

      if (source) {
        // optionally append a small slug suffix if slug exists and source length allows
        if (H.slug>=0 && row[H.slug]) {
          const rawSlug = String(row[H.slug]).replace(/[-_]+/g,' ').trim();
          // take up to 3 words of slug as suffix
          const slugWords = rawSlug.split(/\s+/).slice(0,3).join(' ');
          if (slugWords && source.length + slugWords.length + 3 < MAX_TITLE) {
            source = `${source} (${slugWords})`;
          }
        }
        // enforce MAX_TITLE
        if (source.length > MAX_TITLE) source = source.substring(0, MAX_TITLE).trim();
        titleVal = source;
        sheet.getRange(rIdx, H.title_en+1).setValue(titleVal);
        updates.push([rIdx,'title_en','filled',titleVal]);
      } else {
        updates.push([rIdx,'title_en','left_blank','no source']);
      }
    }

    // SEED alt_text_en from title + descriptors (avoid shape/format)
    let altVal = row[H.alt_text_en] ? String(row[H.alt_text_en]).trim() : '';
    if (!altVal && titleVal) {
      // collect descriptor fields (priority) - take up to 3 short descriptors
      const descCandidates = [];
      if (H.myth_scene>=0 && row[H.myth_scene]) descCandidates.push(String(row[H.myth_scene]).trim());
      if (H.meaning_line>=0 && row[H.meaning_line]) descCandidates.push(String(row[H.meaning_line]).trim());
      if (H.colorway>=0 && row[H.colorway]) descCandidates.push(String(row[H.colorway]).trim());
      if (H.Style>=0 && row[H.Style]) descCandidates.push(String(row[H.Style]).trim());
      if (H.Character>=0 && row[H.Character]) descCandidates.push(String(row[H.Character]).trim());
      // take first 3 and sanitize to short phrases
      const descriptors = descCandidates.slice(0,3).map(s => s.replace(/\s+/g,' ').trim());
      let descPhrase = descriptors.join(', ');
      if (!descPhrase && H.Character>=0 && row[H.Character]) descPhrase = 'Artwork depicting ' + String(row[H.Character]).trim();
      // Build alt_text: "<title_en>. <descPhrase>."
      let candidateAlt = titleVal;
      if (descPhrase) candidateAlt = candidateAlt + '. ' + descPhrase + '.';
      // enforce MAX_ALT
      if (candidateAlt.length > MAX_ALT) candidateAlt = candidateAlt.substring(0, MAX_ALT).trim();
      altVal = candidateAlt;
      sheet.getRange(rIdx, H.alt_text_en+1).setValue(altVal);
      updates.push([rIdx,'alt_text_en','filled',altVal]);
    } else if (!altVal && !titleVal) {
      updates.push([rIdx,'alt_text_en','left_blank','no title source']);
    }
  }

  // Write QA report sheet
  const REP = 'QA_Report_AltUpdates';
  let rep = ss.getSheetByName(REP);
  if (!rep) rep = ss.insertSheet(REP);
  rep.clear();
  rep.appendRow(['row','field','action','value']);
  if (updates.length) rep.getRange(2,1,updates.length,updates[0].length).setValues(updates);

  SpreadsheetApp.getUi().alert('fillTitlesAndAlts complete. Backup: ' + backupName + '. See ' + REP);
}
/**
 * populateTitleDisplayAndShort.gs
 * - Builds title_display using:
 *   Series + ' | ' + Character + ' ' + Character_Variant + ' ' + myth_scene + ' | ' + Style + ' | ' + Colourway + ' | ' + Frame
 * - Builds title_en using:
 *   Character + ' — ' + Character_Variant + (short myth_scene fragment when space allows) <= 60 chars
 *
 * Config:
 *  - SHEET_NAME: exact tab name or '' to use active sheet
 *  - OVERWRITE_TITLE_EN: true = only fill empty title_en; true = overwrite all
 */

function populateTitleDisplayAndShort() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = 'Manifest_GHM_Olympians';              // <-- put exact sheet name or leave '' to use active sheet
  const OVERWRITE_TITLE_EN = true;   // set true to force overwrite of title_en
  const MAX_TITLE_EN = 60;

  const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found. Set SHEET_NAME or open the correct tab.');

  // Backup sheet
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
  const backupName = (sheet.getName() + '_Backup_' + now).substring(0, 99);
  sheet.copyTo(ss).setName(backupName);

  // Read headers and map indexes
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxOf = name => headers.indexOf(name);

  const H = {
    Series: idxOf('Series'),
    Character: idxOf('Character'),
    Character_Variant: idxOf('Character_Variant'),
    myth_scene: idxOf('myth_scene'),
    Style: idxOf('Style'),
    Colourway: idxOf('Colourway'),
    Frame: idxOf('Frame'),
    title_display: idxOf('title_display'),
    title_en: idxOf('title_en'),
    name_final: idxOf('name_final'),
    slug: idxOf('slug')
  };

  // Ensure title_display and title_en columns exist (append to right if missing)
  let colCount = headers.length;
  if (H.title_display === -1) {
    colCount++;
    sheet.getRange(1, colCount).setValue('title_display');
    H.title_display = colCount - 1;
    headers.push('title_display');
  }
  if (H.title_en === -1) {
    colCount++;
    sheet.getRange(1, colCount).setValue('title_en');
    H.title_en = colCount - 1;
    headers.push('title_en');
  }

  // Fetch data rows
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data rows found.'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const changes = []; // [row, field, old, new]

  // Helper: clean & collapse spaces
  function cleanText(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/\s+/g, ' ').trim();
  }

  // Helper: take myth_scene fragment that fits within remaining characters
  function makeShortMythFragment(baseLen, mythScene) {
    if (!mythScene) return '';
    mythScene = cleanText(mythScene);
    if (!mythScene) return '';
    const maxRem = MAX_TITLE_EN - baseLen - 3; // room for ' ()'
    if (maxRem <= 0) return '';
    // try by words
    const words = mythScene.split(/\s+/);
    let frag = '';
    for (let i = 0; i < words.length; i++) {
      const candidate = (frag ? frag + ' ' : '') + words[i];
      if (candidate.length > maxRem) break;
      frag = candidate;
    }
    // if nothing fits, fallback to truncated chars
    if (!frag) frag = mythScene.substring(0, Math.max(0, maxRem)).trim();
    return frag;
  }

  // Build and write
  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const row = data[i];

    const Series = cleanText(H.Series >= 0 ? row[H.Series] : '');
    const Character = cleanText(H.Character >= 0 ? row[H.Character] : '');
    const Variant = cleanText(H.Character_Variant >= 0 ? row[H.Character_Variant] : '');
    const mythScene = cleanText(H.myth_scene >= 0 ? row[H.myth_scene] : '');
    const Style = cleanText(H.Style >= 0 ? row[H.Style] : '');
    const Colourway = cleanText(H.Colourway >= 0 ? row[H.Colourway] : '');
    const Frame = cleanText(H.Frame >= 0 ? row[H.Frame] : '');

    // title_display: Series | Character Variant myth_scene | Style | Colourway | Frame
    // Build character block
    const charParts = [Character, Variant].filter(Boolean).join(' ');
    const charAndMyth = [charParts, mythScene].filter(Boolean).join(' ').trim();

    // assemble display pieces, skip empty pieces, use ' | ' between groups
    const displaySegments = [];
    if (Series) displaySegments.push(Series);
    const middle = charAndMyth || null;
    if (middle) displaySegments.push(middle);
    if (Style) displaySegments.push(Style);
    if (Colourway) displaySegments.push(Colourway);
    if (Frame) displaySegments.push(Frame);
    const titleDisplay = displaySegments.join(' | ').replace(/\s+\|\s+$/,'').trim();

    // write title_display (always overwrite to keep canonical)
    const oldDisplay = row[H.title_display] || '';
    if (String(oldDisplay).trim() !== titleDisplay) {
      sheet.getRange(rowNum, H.title_display + 1).setValue(titleDisplay);
      changes.push([rowNum, 'title_display', oldDisplay, titleDisplay]);
    }

    // title_en: Character — Variant (myth fragment) with MAX_TITLE_EN
    // Build base
    let base = Character || '';
    if (Variant) {
      base = base ? (base + ' — ' + Variant) : Variant;
    }
    let newTitleEn = base;
    if (mythScene && base) {
      const frag = makeShortMythFragment(base.length, mythScene);
      if (frag) {
        newTitleEn = base + ' (' + frag + ')';
      }
    } else if (!base) {
      // fallback: use name_final or slug
      const nameFinal = cleanText(H.name_final >= 0 ? row[H.name_final] : '');
      if (nameFinal) newTitleEn = nameFinal;
      else {
        const slug = cleanText(H.slug >= 0 ? row[H.slug] : '');
        if (slug) newTitleEn = slug.replace(/[-_]+/g, ' ').trim();
      }
    }

    // enforce MAX_TITLE_EN
    if (newTitleEn.length > MAX_TITLE_EN) {
      newTitleEn = newTitleEn.substring(0, MAX_TITLE_EN).trim();
    }

    const oldTitleEn = row[H.title_en] || '';
    const shouldWriteTitleEn = OVERWRITE_TITLE_EN || !oldTitleEn || String(oldTitleEn).trim() === '';
    if (shouldWriteTitleEn && String(oldTitleEn).trim() !== newTitleEn) {
      sheet.getRange(rowNum, H.title_en + 1).setValue(newTitleEn);
      changes.push([rowNum, 'title_en', oldTitleEn, newTitleEn]);
    }
  }

  // Write QA report
  const REP = 'QA_Report_TitleUpdates';
  let rep = ss.getSheetByName(REP);
  if (!rep) rep = ss.insertSheet(REP);
  rep.clear();
  if (changes.length) {
    rep.appendRow(['row','field','old','new']);
    rep.getRange(2,1,changes.length,changes[0].length).setValues(changes);
  } else {
    rep.appendRow(['no_changes']);
  }

  SpreadsheetApp.getUi().alert('populateTitleDisplayAndShort complete. Backup: ' + backupName + '. See ' + REP);
}
function headerInspector() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = ''; // set sheet name or leave '' to use active sheet
  const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  if (!sh) throw new Error('Sheet not found');
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  Logger.log('HEADERS:\n' + headers.join(' | '));
  // show first data row
  if (sh.getLastRow() >= 2) {
    const firstRow = sh.getRange(2,1,1,headers.length).getValues()[0];
    Logger.log('FIRST ROW SAMPLE:\n' + firstRow.join(' | '));
  } else {
    Logger.log('No data rows found.');
  }
  SpreadsheetApp.getUi().alert('Headers & sample logged. Open View → Logs.');
}
function dryRunTitleBuild() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = ''; // set or leave blank for active
  const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found');

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idxOf = name => headers.indexOf(name);

  const H = {
    Series: idxOf('Series'),
    Character: idxOf('Character'),
    Character_Variant: idxOf('Character_Variant'),
    myth_scene: idxOf('myth_scene'),
    Style: idxOf('Style'),
    Colourway: idxOf('Colourway'),
    Frame: idxOf('Frame'),
    title_display: idxOf('title_display'),
    title_en: idxOf('title_en'),
    name_final: idxOf('name_final'),
    slug: idxOf('slug')
  };

  // gather rows
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data rows.'); return; }
  const data = sheet.getRange(2,1,lastRow-1,headers.length).getValues();

  const MAX_TITLE_EN = 60;
  function clean(s){ return (s||'').toString().replace(/\s+/g,' ').trim(); }
  function makeShortMythFragment(baseLen, mythScene) {
    if (!mythScene) return '';
    mythScene = clean(mythScene);
    const maxRem = MAX_TITLE_EN - baseLen - 3;
    if (maxRem <= 0) return '';
    const words = mythScene.split(/\s+/);
    let frag='';
    for (let i=0;i<words.length;i++){
      const candidate = (frag?frag+' ':'') + words[i];
      if (candidate.length > maxRem) break;
      frag = candidate;
    }
    if (!frag) frag = mythScene.substring(0, Math.max(0,maxRem)).trim();
    return frag;
  }

  const report = [];
  for (let i=0;i<data.length;i++){
    const r = data[i];
    const rowNum = i+2;
    const Series = clean(H.Series>=0? r[H.Series] : '');
    const Character = clean(H.Character>=0? r[H.Character] : '');
    const Variant = clean(H.Character_Variant>=0? r[H.Character_Variant] : '');
    const mythScene = clean(H.myth_scene>=0? r[H.myth_scene] : '');
    const Style = clean(H.Style>=0? r[H.Style] : '');
    const Colourway = clean(H.Colourway>=0? r[H.Colourway] : '');
    const Frame = clean(H.Frame>=0? r[H.Frame] : '');
    const charParts = [Character, Variant].filter(Boolean).join(' ');
    const middle = [charParts, mythScene].filter(Boolean).join(' ').trim();
    const displaySegments = [];
    if (Series) displaySegments.push(Series);
    if (middle) displaySegments.push(middle);
    if (Style) displaySegments.push(Style);
    if (Colourway) displaySegments.push(Colourway);
    if (Frame) displaySegments.push(Frame);
    const titleDisplay = displaySegments.join(' | ').trim();

    let base = Character || '';
    if (Variant) base = base ? (base + ' — ' + Variant) : Variant;
    let newTitleEn = base;
    if (mythScene && base) {
      const frag = makeShortMythFragment(base.length, mythScene);
      if (frag) newTitleEn = base + ' (' + frag + ')';
    } else if (!base) {
      const nameFinal = clean(H.name_final>=0? r[H.name_final] : '');
      const slug = clean(H.slug>=0? r[H.slug] : '');
      newTitleEn = nameFinal || slug;
    }
    if (newTitleEn && newTitleEn.length > MAX_TITLE_EN) newTitleEn = newTitleEn.substring(0,MAX_TITLE_EN).trim();

    // existing values (if present)
    const existingDisplay = H.title_display>=0 ? clean(r[H.title_display]) : '';
    const existingShort = H.title_en>=0 ? clean(r[H.title_en]) : '';

    report.push([rowNum, existingDisplay, titleDisplay, existingShort, newTitleEn]);
  }

  // write report
  const RN = 'QA_Report_TitleDryRun';
  let rep = ss.getSheetByName(RN);
  if (!rep) rep = ss.insertSheet(RN);
  rep.clear();
  rep.appendRow(['row','existing_title_display','computed_title_display','existing_title_en','computed_title_en']);
  rep.getRange(2,1,report.length,report[0].length).setValues(report);
  SpreadsheetApp.getUi().alert('Dry-run complete — open sheet: ' + RN);
}
function normalizeAltTextEn() {
  const ss = SpreadsheetApp.getActive();
  const SHEET = ''; // leave '' to use the active sheet or put the exact tab name
  const sheet = SHEET ? ss.getSheetByName(SHEET) : ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found.');

  // backup
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
  sheet.copyTo(ss).setName(sheet.getName() + '_Backup_ALT_' + now);

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const iAlt = headers.indexOf('alt_text_en');
  const iTitle = headers.indexOf('title_en');
  const iMyth = headers.indexOf('myth_scene');
  const iColor = headers.indexOf('Colorway') >= 0 ? headers.indexOf('Colorway') : headers.indexOf('colourway');
  const iMeaning = headers.indexOf('meaning_line');
  const iCharacter = headers.indexOf('Character');

  if (iAlt < 0) throw new Error('alt_text_en header not found');

  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues();
  const report = [];
  const shapeStarts = ['round canvas','round canvas portrait','square canvas','oval canvas','portrait','landscape','round portrait'];

  for (let r=0; r<rows.length; r++){
    const rowNum = r+2;
    let alt = rows[r][iAlt] ? String(rows[r][iAlt]).trim() : '';
    const title = (iTitle>=0 && rows[r][iTitle]) ? String(rows[r][iTitle]).trim() : '';
    // If alt exists but starts with shape words, strip them
    if (alt) {
      let lower = alt.toLowerCase();
      // remove leading shape phrases
      shapeStarts.forEach(function(p){
        if (lower.startsWith(p)) {
          alt = alt.replace(new RegExp('^'+p,'i'),'').replace(/^[:,\-\s]+/,'').trim();
        }
      });
      // convert "palette: Gold" or "palette:Gold" to "Gold palette" (move palette word to end)
      alt = alt.replace(/palette\s*:\s*/i, '').replace(/\s*palette\s*$/i,'').trim();
      // ensure alt starts with title if title exists
      if (title && !alt.toLowerCase().startsWith(title.toLowerCase())) {
        alt = title + '. ' + alt;
      }
      // normalize extra punctuation and enforce length <=125
      if (alt.length > 125) alt = alt.substring(0,125).trim();
      sheet.getRange(rowNum, iAlt+1).setValue(alt);
      report.push([rowNum,'alt_text_en','cleaned_existing',alt]);
      continue;
    }

    // If alt is empty, build it from title + descriptors
    if (!alt && title) {
      const descriptors = [];
      if (iMyth>=0 && rows[r][iMyth]) descriptors.push(String(rows[r][iMyth]).trim());
      if (iMeaning>=0 && rows[r][iMeaning]) descriptors.push(String(rows[r][iMeaning]).trim());
      if (iColor>=0 && rows[r][iColor]) descriptors.push(String(rows[r][iColor]).trim());
      if (iCharacter>=0 && rows[r][iCharacter]) descriptors.push(String(rows[r][iCharacter]).trim());
      const desc = descriptors.slice(0,3).join(', ');
      let built = title + (desc ? ('. ' + desc + '.') : '.');
      if (built.length > 125) built = built.substring(0,125).trim();
      sheet.getRange(rowNum, iAlt+1).setValue(built);
      report.push([rowNum,'alt_text_en','generated',built]);
    }
  }

  // write report
  const RN = 'QA_Report_AltClean';
  let rep = ss.getSheetByName(RN);
  if (!rep) rep = ss.insertSheet(RN);
  rep.clear();
  if (report.length) {
    rep.appendRow(['row','field','action','value']);
    rep.getRange(2,1,report.length,report[0].length).setValues(report);
  } else rep.appendRow(['no_changes']);

  SpreadsheetApp.getUi().alert('normalizeAltTextEn done. See ' + RN);
}
function fixAltTextEn_Strict() {
  const ss = SpreadsheetApp.getActive();
  const SHEET = ''; // leave '' to use active sheet or put exact tab name
  const sheet = SHEET ? ss.getSheetByName(SHEET) : ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found.');

  // backup
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
  sheet.copyTo(ss).setName(sheet.getName() + '_Backup_ALTfix_' + now);

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const ix = name => headers.indexOf(name);
  const H = {
    alt: ix('alt_text_en'),
    title: ix('title_en'),
    name_final: ix('name_final'),
    Character: ix('Character'),
    myth: ix('myth_scene'),
    meaning: ix('meaning_line'),
    Colorway: ix('Colorway'),
    Style: ix('Style')
  };
  if (H.alt < 0) throw new Error('alt_text_en header not found');

  // common corrections map (add more as needed)
  const corrections = {
    'Posiedon': 'Poseidon',
    'posiedon': 'Poseidon',
    'HADES': 'Hades',
    'hades': 'Hades'
  };

  function applyCorrections(s) {
    if (!s) return s;
    Object.keys(corrections).forEach(k => {
      const re = new RegExp('\\b' + k + '\\b','g');
      s = s.replace(re, corrections[k]);
    });
    return s;
  }

  function cleanParenthetical(s) {
    if (!s) return s;
    return s.replace(/\s*\([^)]*\)/g,'').replace(/\s+/g,' ').trim();
  }

  function stripLeadingShapeOrPortrait(s) {
    if (!s) return s;
    // strip common leading shape phrases
    s = s.replace(/^\s*(round|square|oval|rectangular|circular)\s+canvas(\s+portrait)?[^\.\-\:]*[:\-\s]*/i,'');
    // strip leading 'portrait of' or 'portrait - of' etc
    s = s.replace(/^\s*portrait(?:\s*of|\s*-\s*of)?\s+/i,'');
    s = s.replace(/^\s*portrait[,]?\s+/i,'');
    return s.trim();
  }

  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues();
  const report = [];

  for (let r=0; r<rows.length; r++){
    const rowNum = r+2;
    let alt = rows[r][H.alt] ? String(rows[r][H.alt]).trim() : '';
    // build base title (prefer title_en, fallback to name_final, then Character)
    let baseTitle = (H.title>=0 && rows[r][H.title]) ? String(rows[r][H.title]).trim() : '';
    if (!baseTitle && H.name_final>=0 && rows[r][H.name_final]) baseTitle = String(rows[r][H.name_final]).trim();
    if (!baseTitle && H.Character>=0 && rows[r][H.Character]) baseTitle = String(rows[r][H.Character]).trim();
    baseTitle = applyCorrections(baseTitle);
    baseTitle = cleanParenthetical(baseTitle);

    // Clean existing alt if present
    if (alt) {
      // apply corrections first
      alt = applyCorrections(alt);
      // strip shape/portrait starts
      alt = stripLeadingShapeOrPortrait(alt);
      // remove any parenthetical repeats inside alt
      alt = cleanParenthetical(alt);
      // remove dangling tokens like 'palette:' -> move the word after
      alt = alt.replace(/palette\s*:\s*/i,'').trim();
      // ensure starts with baseTitle
      if (baseTitle && !alt.toLowerCase().startsWith(baseTitle.toLowerCase())) {
        alt = baseTitle + '. ' + alt;
      }
      // normalize punctuation
      alt = alt.replace(/\s+\./g,'.').replace(/\s{2,}/g,' ').trim();
      // enforce length
      if (alt.length > 125) alt = alt.substring(0,125).trim();
      sheet.getRange(rowNum, H.alt+1).setValue(alt);
      report.push([rowNum,'alt_text_en','cleaned_existing',alt]);
      continue;
    }

    // If alt empty -> build from baseTitle + descriptors
    const descriptors = [];
    if (H.myth>=0 && rows[r][H.myth]) descriptors.push(String(rows[r][H.myth]).trim());
    if (H.meaning>=0 && rows[r][H.meaning]) descriptors.push(String(rows[r][H.meaning]).trim());
    if (H.Colorway>=0 && rows[r][H.Colorway]) descriptors.push(String(rows[r][H.Colorway]).trim());
    if (H.Style>=0 && rows[r][H.Style]) descriptors.push(String(rows[r][H.Style]).trim());
    // pick up to first 3 short descriptors, clean them (no parenthetical)
    const desc = descriptors.map(d => applyCorrections(cleanParenthetical(d))).filter(Boolean).slice(0,3).join(', ');
    let built = baseTitle || '';
    if (desc) built = (built ? built + '. ' : '') + desc + '.';
    // fallback if still empty
    if (!built) built = 'Artwork image.';
    // enforce length
    if (built.length > 125) built = built.substring(0,125).trim();
    sheet.getRange(rowNum, H.alt+1).setValue(built);
    report.push([rowNum,'alt_text_en','generated',built]);
  }

  // write report
  const RN = 'QA_Report_AltFix';
  let rep = ss.getSheetByName(RN);
  if (!rep) rep = ss.insertSheet(RN);
  rep.clear();
  if (report.length) {
    rep.appendRow(['row','field','action','value']);
    rep.getRange(2,1,report.length,report[0].length).setValues(report);
  } else rep.appendRow(['no_changes']);

  SpreadsheetApp.getUi().alert('fixAltTextEn_Strict done. See ' + RN);
}
function forceRewriteAltTextEn() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = ''; // put tab name or leave '' to use active sheet
  const OVERWRITE_ALL = true; // set false to only fill empty alts
  const MAX_ALT = 125;

  const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  if (!sh) throw new Error('Sheet not found.');

  // backup
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
  sh.copyTo(ss).setName(sh.getName() + '_Backup_ALTforce_' + now);

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const ix = h => Math.max(0, headers.indexOf(h));
  const H = {
    alt: headers.indexOf('alt_text_en'),
    title: headers.indexOf('title_en'),
    myth: headers.indexOf('myth_scene'),
    meaning: headers.indexOf('meaning_line'),
    color: headers.indexOf('Colorway') >= 0 ? headers.indexOf('Colorway') : headers.indexOf('Colourway'),
    character: headers.indexOf('Character'),
    name_final: headers.indexOf('name_final')
  };
  if (H.alt < 0) throw new Error('alt_text_en header not found');

  function cleanParenthetical(s){ return s ? String(s).replace(/\s*\([^)]*\)/g,'').replace(/\s+/g,' ').trim() : ''; }
  function stripPortraitPaletteNoise(s){
    if(!s) return s;
    s = String(s);
    s = s.replace(/portrait\s*(?:of)?\s*/i,'');               // remove 'portrait' or 'portrait of'
    s = s.replace(/palette\s*[:\-]\s*/i,'');                  // remove 'palette:' label
    s = s.replace(/\bportrait\b/i,'');                        // any stray portrait
    s = s.replace(/\s{2,}/g,' ').trim();
    return s;
  }

  const rows = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const report = [];

  for (let r=0;r<rows.length;r++){
    const rowNum = r+2;
    const row = rows[r];

    // base title (prefer title_en, fallback to name_final or Character)
    let base = H.title>=0 && row[H.title] ? String(row[H.title]).trim() : '';
    if (!base && H.name_final>=0 && row[H.name_final]) base = String(row[H.name_final]).trim();
    if (!base && H.character>=0 && row[H.character]) base = String(row[H.character]).trim();
    base = cleanParenthetical(base);

    // gather descriptors (first non-empty of myth/meaning/color)
    const descPieces = [];
    if (H.myth >= 0 && row[H.myth]) descPieces.push(String(row[H.myth]).trim());
    if (H.meaning >= 0 && row[H.meaning]) descPieces.push(String(row[H.meaning]).trim());
    if (H.color >= 0 && row[H.color]) descPieces.push(String(row[H.color]).trim());
    // sanitize descriptors
    const cleanDescs = descPieces.map(d => stripPortraitPaletteNoise(cleanParenthetical(d))).filter(Boolean).slice(0,3);
    let desc = cleanDescs.join(', ');

    // build new alt: "<base>. <desc>."
    let newAlt = base || '';
    if (desc) newAlt = (newAlt ? newAlt + '. ' : '') + desc + '.';
    if (!newAlt) newAlt = 'Artwork image.';

    // enforce length
    if (newAlt.length > MAX_ALT) newAlt = newAlt.substring(0, MAX_ALT).trim();

    const existingAlt = row[H.alt] ? String(row[H.alt]).trim() : '';

    if (OVERWRITE_ALL || !existingAlt) {
      sh.getRange(rowNum, H.alt+1).setValue(newAlt);
      report.push([rowNum, existingAlt, newAlt]);
    }
  }

  // write report
  const RN = 'QA_Report_AltForce';
  let rep = ss.getSheetByName(RN);
  if (!rep) rep = ss.insertSheet(RN);
  rep.clear();
  if (report.length) {
    rep.appendRow(['row','old_alt','new_alt']);
    rep.getRange(2,1,report.length,report[0].length).setValues(report);
  } else {
    rep.appendRow(['no_changes']);
  }

  SpreadsheetApp.getUi().alert('forceRewriteAltTextEn complete — see sheet: ' + RN);
}
function fixCommonNames() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = 'Manifest_GHM_Olympians';
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + SHEET_NAME);

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idxChar = headers.indexOf('Character');
  const idxNameFinal = headers.indexOf('name_final');
  if (idxChar < 0 && idxNameFinal < 0) throw new Error('Character/name_final headers not found');

  const map = {'Posiedon':'Poseidon','posiedon':'Poseidon','HADES':'Hades','hades':'Hades'};
  const rows = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const updates = [];
  for (let r=0;r<rows.length;r++){
    const rowNum = r+2;
    if (idxChar >= 0 && rows[r][idxChar]) {
      const cur = String(rows[r][idxChar]);
      if (map[cur]) {
        sh.getRange(rowNum, idxChar+1).setValue(map[cur]);
        updates.push([rowNum,'Character',cur,map[cur]]);
      }
    }
    if (idxNameFinal >= 0 && rows[r][idxNameFinal]) {
      const cur2 = String(rows[r][idxNameFinal]);
      // replace only exact token occurrences
      let replaced = cur2;
      Object.keys(map).forEach(k => { replaced = replaced.replace(new RegExp('\\b'+k+'\\b','g'), map[k]); });
      if (replaced !== cur2) {
        sh.getRange(rowNum, idxNameFinal+1).setValue(replaced);
        updates.push([rowNum,'name_final',cur2,replaced]);
      }
    }
  }
  const repName = 'QA_Report_NameMapping';
  let rep = ss.getSheetByName(repName) || ss.insertSheet(repName);
  rep.clear(); rep.appendRow(['row','field','old','new']);
  if (updates.length) rep.getRange(2,1,updates.length,updates[0].length).setValues(updates);
  SpreadsheetApp.getUi().alert('Name mapping run complete. See ' + repName);
}
function forceRewriteAltTextEn_Targeted() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = 'Manifest_GHM_Olympians'; // <- exact tab name
  const OVERWRITE_ALL = true;
  const MAX_ALT = 125;

  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + SHEET_NAME);

  // backup
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
  sh.copyTo(ss).setName(sh.getName() + '_Backup_ALTforce_' + now);

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const H = {
    alt: headers.indexOf('alt_text_en'),
    title: headers.indexOf('title_en'),
    myth: headers.indexOf('myth_scene'),
    meaning: headers.indexOf('meaning_line'),
    color: headers.indexOf('Colorway') >= 0 ? headers.indexOf('Colorway') : headers.indexOf('Colourway'),
    character: headers.indexOf('Character'),
    name_final: headers.indexOf('name_final')
  };
  if (H.alt < 0) throw new Error('alt_text_en header not found on ' + SHEET_NAME);

  const rows = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const report = [];
  for (let r=0;r<rows.length;r++){
    const rowNum = r+2;
    const row = rows[r];

    let base = H.title>=0 && row[H.title] ? String(row[H.title]).trim() : '';
    if (!base && H.name_final>=0 && row[H.name_final]) base = String(row[H.name_final]).trim();
    if (!base && H.character>=0 && row[H.character]) base = String(row[H.character]).trim();
    base = base.replace(/\s*\([^)]*\)/g,'').trim();

    const descPieces = [];
    if (H.myth >= 0 && row[H.myth]) descPieces.push(String(row[H.myth]).trim());
    if (H.meaning >= 0 && row[H.meaning]) descPieces.push(String(row[H.meaning]).trim());
    if (H.color >= 0 && row[H.color]) descPieces.push(String(row[H.color]).trim());
    const cleanDescs = descPieces.map(d => String(d).replace(/\s*\([^)]*\)/g,'').replace(/palette\s*[:\-]\s*/i,'').trim()).filter(Boolean).slice(0,3);
    let desc = cleanDescs.join(', ');
    let newAlt = base || '';
    if (desc) newAlt = (newAlt ? newAlt + '. ' : '') + desc + '.';
    if (!newAlt) newAlt = 'Artwork image.';
    if (newAlt.length > MAX_ALT) newAlt = newAlt.substring(0, MAX_ALT).trim();

    const existingAlt = row[H.alt] ? String(row[H.alt]).trim() : '';
    if (OVERWRITE_ALL || !existingAlt) {
      try {
        sh.getRange(rowNum, H.alt+1).setValue(newAlt);
        report.push([rowNum, existingAlt, newAlt, 'ok']);
      } catch(e) {
        report.push([rowNum, existingAlt, newAlt, 'ERROR: ' + e.message]);
      }
    }
  }

  const RN = 'QA_Report_AltForce';
  let rep = ss.getSheetByName(RN) || ss.insertSheet(RN);
  rep.clear(); rep.appendRow(['row','old_alt','new_alt','status']);
  if (report.length) rep.getRange(2,1,report.length,report[0].length).setValues(report);
  SpreadsheetApp.getUi().alert('Alt force run complete. See ' + RN);
}
function listProtections() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = 'Manifest_GHM_Olympians';
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + SHEET_NAME);

  const protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(sh.getProtections(SpreadsheetApp.ProtectionType.SHEET));
  const out = protections.map(p => {
    return [p.getDescription() || '(no desc)', p.getEditors().map(e=>e.getEmail()).join(','), p.canDomainEdit()];
  });
  const RN = 'QA_Report_Protections';
  let rep = ss.getSheetByName(RN) || ss.insertSheet(RN);
  rep.clear(); rep.appendRow(['desc','editors','canDomainEdit']);
  if (out.length) rep.getRange(2,1,out.length,out[0].length).setValues(out);
  SpreadsheetApp.getUi().alert('Protection list written to ' + RN + '.');
}
function recordFullBackupMetadata(name) {
  const ss = SpreadsheetApp.getActive();
  const CONTROL = 'Control_GHM';
  const sheet = ss.getSheetByName(CONTROL);
  if (!sheet) throw new Error(CONTROL + ' not found');

  const now = new Date().toISOString();
  const user = Session.getActiveUser().getEmail() || '<owner_email_here>';
  const map = {
    'last_full_backup': name || ('GHM_Master_Manifest_BACKUP_' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm')),
    'last_backup_at': now,
    'last_backup_by': user
  };

  // read existing key/value rows (A:B)
  const data = sheet.getRange(1,1,sheet.getLastRow(),2).getValues();
  const out = {};
  data.forEach(r => { if (r[0]) out[r[0]] = r[1]; });

  Object.keys(map).forEach(function(key){
    let rowIdx = data.findIndex(r => r[0] === key);
    if (rowIdx === -1) {
      sheet.appendRow([key, map[key]]);
    } else {
      sheet.getRange(rowIdx+1, 2).setValue(map[key]);
    }
  });

  SpreadsheetApp.getUi().alert('Backup metadata recorded in ' + CONTROL);
}
