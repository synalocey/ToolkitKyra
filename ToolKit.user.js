// ==UserScript==
// @name         Toolkit Kyra
// @version      26.11.06
// @namespace    https://github.com/synalocey/ToolkitKyra
// @description  Toolkit Kyra
// @author       Syna
// @license      GPL-3.0
// @icon64       https://a.favicon.im/lamresearch.com
// @downloadURL  https://raw.githubusercontent.com/synalocey/ToolkitKyra/main/ToolKit.user.js
// @updateURL    https://raw.githubusercontent.com/synalocey/ToolkitKyra/main/ToolKit.user.js
// @require      https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js
// @require      https://cdn.jsdelivr.net/gh/JJJ/chosen@4.0.0/dist/js/chosen.jquery.min.js
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @require      https://cdn.jsdelivr.net/npm/@zip.js/zip.js@2.24.0/dist/zip.min.js
// @resource     excelJs https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @resource     tabulatorJs https://cdn.jsdelivr.net/npm/tabulator-tables@6.5.0/dist/js/tabulator.min.js
// @resource     tabulatorCss https://cdn.jsdelivr.net/npm/tabulator-tables@6.5.0/dist/css/tabulator.min.css
// @resource     chosenCss https://cdn.jsdelivr.net/gh/JJJ/chosen@4.0.0/dist/css/chosen.min.css
// @resource     xlsxJs https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @match        *://*.lamresearch.com/*
// @match        *://*.mylam.com/*
// @match        https://kmmatrix.fremont.lamrc.net/*
// @match        http://127.0.0.1/toolbox/part-changes
// @match        *://*.github.io/*
// @run-at       document-start
// @noframes
// @grant        GM_registerMenuCommand
// @grant        GM_setClipboard
// @grant        GM_xmlhttpRequest
// @grant        GM_getResourceText
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM.setValue
// @grant        GM.setValues
// @grant        GM_deleteValue
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// @grant        GM_openInTab
// @grant        GM_getTab
// @grant        GM_saveTab
// @grant        unsafeWindow
// @connect      127.0.0.1
// @connect      pcazbwpap04.fremont.lamrc.net
// @connect      synalocey.github.io
// @connect      raw.githubusercontent.com
// @connect      github.com
// ==/UserScript==
/* globals jQuery, XLSX, zip */
const chosenBaseCss=GM_getResourceText('chosenCss');
let excelJs;
function loadExcelJs(){
  if(excelJs)return excelJs;
  const lib={exports:{}};
  new Function('module','exports',GM_getResourceText('excelJs'))(lib,lib.exports);
  if(typeof lib.exports.Workbook!=='function')throw new Error('Excel workbook library is unavailable. Refresh the page and try again.');
  excelJs=lib.exports;return excelJs;
}

// BOM parsing shared by MyLam and SAP Toolbox.
const BOM_COMPARE_LABEL='BOM Compare';
const PART_CHANGES_MYLAM_URL='https://techinfo.mylam.com/portal/group/mylam/contenttyperesults?ctype=Tech%20Articles';
const PART_CHANGES_MYLAM_TAB='kyra-part-changes';
const PART_CHANGES_LAUNCH_PATH='/toolbox/part-changes';
const PART_CHANGES_ACTION_LABEL='🔩 Find Part Changes';
// MyLam reports "Comparing 45% · …" as the task phase while it builds the comparison files.
const PART_CHANGES_COMPARING='Comparing ';
const ZBOM_COMPARE_DEFAULT_TEMPLATE='{Customer}_{FID}{FIDver}_{System Description}',ZBOM_COMPARE_DEFAULT_SEPARATOR='vs',ZBOM_COMPARE_STANDALONE='__STANDALONE__';
const PART_CHANGES_DEFAULT_TEMPLATE='[BOM] '+ZBOM_COMPARE_DEFAULT_TEMPLATE;

function partChangesHandoff(hash){
  const match=/^#kyra-part-changes=([0-9a-f]{32})&port=(\d{1,5})(?:&systems=([^&]+))?$/i.exec(hash);
  if(!match||Number(match[2])<1||Number(match[2])>65535)return null;
  const request={id:match[1],port:Number(match[2])};
  if(match[3])try{
    const systems=JSON.parse(decodeURIComponent(match[3]));
    if(!systems||typeof systems!=='object'||Array.isArray(systems)||Object.entries(systems).some(([group,selection])=>{
      const legacy=Array.isArray(selection),labels=legacy?selection:selection?.rows;
      return !/^G\d{2,}$/.test(group)||!Array.isArray(labels)||!labels.length||
        labels.some(label=>typeof label!=='string'||(legacy||label!=='')&&!label.trim());
    }))return null;
    request.systems=systems;
  }catch(_){return null}
  return request;
}
const PART_CHANGES_MYLAM_SYSTEM_KEY='toolkit.partChanges.systems.v1';
// Tech Articles System labels. The MyLam helper refreshes the saved list from the live filter.
const PART_CHANGES_MYLAM_SYSTEMS=[
  "384", "384T", "2300 Ascent Cu CMP", "2300 Microwave Stripper", "2300 Transport Module 2LPM V2",
  "2300 Transport Module 3LPM V2", "2300 Transport Module V2 Dual", "2300 Transport Module V3",
  "2300 Transport Module V3 Dual", "2300 Transport Module V4", "2300e4 Transport Module", "2300e5 Transport Module",
  "2300e6 Transport Module", "4520XLE", "4520XLE Alliance", "4520XLEi", "4520XLEi Alliance", "4520XLi Alliance",
  "Aether GD", "Aether GDX", "Aether GP", "Aether GPX", "Aether GT", "Aether GTX", "Aether HT", "Akara AL", "Akara AP",
  "Akara AX", "Akara BL", "Akara BT", "Akara BV", "Akara BX", "Akara CX", "Alliance Transport Module", "ALTUS 450",
  "ALTUS Degas", "ALTUS DirectFill (WN)", "ALTUS DirectFill Max (WN Max)", "ALTUS Halo GW", "ALTUS Halo GX",
  "ALTUS Halo HX", "ALTUS Halo HXD", "ALTUS Halo Inflection", "ALTUS Halo JX", "ALTUS Inflection G",
  "ALTUS Inflection GS", "ALTUS LFW EL", "ALTUS LFW GL", "ALTUS LFWn", "ALTUS Max", "ALTUS PreClean (PC)",
  "ALTUS PreClean Max (PC Max)", "Argos", "Argos FXT", "Argos GT", "Argos HP", "ASIQ", "Aura", "Auriga", "AutoEtch 490",
  "AutoEtch 490/590", "AutoEtch 490B", "AutoEtch 490B/590B", "AutoEtch 590", "AutoEtch 590B", "AutoEtch 690",
  "AutoEtch 690B", "AutoEtch 790", "AutoEtch 790B", "Avanti", "C2 ALTUS", "C2 SEQUEL", "C2 SEQUEL Express", "C2 SPEED",
  "C2 SPEED NExT", "C3 ALTUS", "C3 ALTUSi", "C3 ALTUSi ICEFill", "C3 SEQUEL", "C3 SPEED", "C3 SPEED Max",
  "C3 SPEED NExT", "C3 SPEEDi", "C360", "Casper AL", "Casper AX", "CMS", "Concept 3", "Concept One", "Concept Two",
  "Continuum Process Module", "Continuum Transport Module", "Coronus", "Coronus D", "Coronus DX", "Coronus HP",
  "Coronus HPT", "Da Vinci (DV-24)", "Da Vinci (DV-28)", "Da Vinci (DV-34)", "Da Vinci (DV-38)", "Dextro Flex GB",
  "Dextro Flex HL", "Dextro Flex HX", "Dextro Flex HXE", "Dextro Kiyo FX", "Dextro Kiyo GX", "Dextro Sabre 3D",
  "Dextro Vantex B", "Dextro Vantex CX+", "DLCM", "DLCM-S", "DLCMi", "DSiE F Alliance", "DSiE G", "DSM 9800",
  "DSM 9900 Process Module", "DSM 9900 Transport Module", "DSQ Stripper", "DSQ Stripper Alliance (A4)",
  "DSQ Stripper Alliance (A6)", "DSS-150", "DSS-200 Series 0", "DSS-200 Series I", "DSS-200 Series II",
  "DSS-200 Series II CE", "DSS-200 Series IIi", "DV-Prime", "DV-Prime Plus (200MM)", "DV-Prime Plus (300MM)", "EDGE",
  "EFx", "EFx 450", "EI-AA", "EI-ARC", "EI-AutoTRD", "EI-CV", "EI-DA 2.x", "EI-DAe", "EI-DH", "EI-DPG-ABPC",
  "EI-DPG-AD", "EI-DPG-OPT", "EI-DPG-RTM", "EI-DS", "EI-DS2.0", "EI-FM", "EI-Math", "EI-PHA", "EI-PPD", "EI-PTA",
  "EI-Sandbox", "EI-SRC", "EI-SSHC", "EI-Workbench", "EIM", "ELD", "Envion", "Envoy", "EOS", "EOS DS L", "EOS DS L PTK",
  "EOS DS M", "EOS DS MT", "EOS ES H", "EOS ES L", "EOS ES M", "EOS ES N", "EOS GS L", "EPIC", "Esanti (ES-28)",
  "Esanti (ES-34)", "Esanti (ES-38)", "Exelan", "Exelan Alliance", "Exelan Flex", "Exelan Flex3x", "Exelan Flex45",
  "Exelan FlexCore", "Exelan HP Alliance", "Exelan HPT Alliance", "Extrima 300NX", "Extrima 3100i", "Extrima 3100P",
  "Extrima 6000", "FFW", "Flex DL", "Flex DS", "Flex DX", "Flex EL", "Flex ES", "Flex EX", "Flex EX+", "Flex FL",
  "Flex FS", "Flex FSE", "Flex FX", "Flex FXP", "Flex GB", "Flex GB+", "Flex GL", "Flex GS", "Flex GX", "Flex GXE",
  "Flex HL", "Flex HX", "Flex HX+", "Flex HXE", "Flex HXE Cryo", "Flex HXP", "Flex JL", "Flex JXE", "Flex JXP",
  "Flex45 DL", "Flex45 DS", "Flex45 DSP", "Flex45 DX", "Flex45 EL", "Flex45 ES", "Flex45 EX", "G3D", "G400",
  "Galaxy CX", "GAMMA 2100", "GAMMA 2130", "GAMMA 2130i", "GAMMA Express", "GAMMA xPR", "GAMMA xPRi", "GxT",
  "High Flow Metal Alliance", "ICE - RP", "ICEFill", "ICEFill Max G", "ICEFill Max H", "ICS", "ICS-HP", "ICS-S",
  "Ikano", "iLPC", "Inline Series", "INOVA", "INOVA 3D", "INOVA NExT", "INOVA xT", "Integrity", "Ion Beam Etch (IBE)",
  "Kallisto", "Kiyo CL", "Kiyo CM", "Kiyo CS", "Kiyo CX", "Kiyo ES", "Kiyo EX", "Kiyo EX with Hydra", "Kiyo EXP",
  "Kiyo FP", "Kiyo FT", "Kiyo FT+", "Kiyo FX", "Kiyo FXE", "Kiyo FXE+", "Kiyo FXM", "Kiyo FXT", "Kiyo GP", "Kiyo GX",
  "Kiyo GX DWG", "Kiyo GXE", "Kiyo GXT", "Kiyo MCX", "Kyber", "L-VHP", "L2200", "L3200", "L3300", "L3500", "LSrFM",
  "M Series", "MACH 1+", "MACH I", "MACH Ii", "MACH IQ", "MACH IT", "MACH IV", "MACH1+XR", "Megastrip", "Mentor DF3",
  "Mentor OC23", "Mentor OC23E", "Mentor SF3", "Metior HX", "Metior SRM", "Metryx 2W EFEM", "Metryx 3W CEFEM",
  "Microwave Stripper Alliance", "Microwave Stripper Alliance (A4)", "MInT", "Momentum", "Motif", "Oxide 9500",
  "Oxide 9500 Alliance", "PEP Iridia", "PES-2300", "Phoenix", "Prestis", "Prestis Degas", "Prestis PreClean",
  "Prevos FX", "Prevos FXT", "Prism", "QUAD", "Rainbow 4400", "Rainbow 4420", "Rainbow 4420 Alliance", "Rainbow 4420XL",
  "Rainbow 4420XL Alliance", "Rainbow 4500", "Rainbow 4500i", "Rainbow 4500is", "Rainbow 4500s", "Rainbow 4520",
  "Rainbow 4520 Alliance", "Rainbow 4520i", "Rainbow 4520is", "Rainbow 4520s", "Rainbow 4520XL",
  "Rainbow 4520XL Alliance", "Rainbow 4520XLi", "Rainbow 4600", "Rainbow 4620", "Rainbow 4700", "Rainbow 4700 Alliance",
  "Rainbow 4720", "Rainbow 4720 Alliance", "Rainbow 4720A Alliance", "Rainbow 4720XL", "Rainbow 4720XL Alliance",
  "SABRE 3D", "SABRE 3D EX", "SABRE 3D FP", "SABRE 3D Scimitar", "SABRE 3DxT", "SABRE 450", "SABRE Anneal Module",
  "SABRE Classic", "SABRE ELD", "SABRE EX", "SABRE Excel", "SABRE Excel SBOM12+", "SABRE Excel SBOM12+i",
  "SABRE Extreme", "SABRE FC", "SABRE FX", "SABRE FXG", "SABRE FXT", "SABRE HX", "SABRE Max", "SABRE Max SBOM12+",
  "SABRE NExT", "SABRE SSM", "SABRE xT", "SABRE xTi", "Selis FS", "Selis FX", "Selis FXT", "Sense.i",
  "Serene Module and FDS (V3 Version)", "Serene Module and Vacuum System  (V2 Version)",
  "Serene Process Module (BRS Version)", "Sierra", "SIG 450", "Silylation Alliance", "Single Suzi", "SOLA",
  "SOLA Excel", "SOLA Excel Clear 2", "SOLA Excel PC", "SOLA xT", "SOLA xTi", "Solo", "SP100", "SP101", "SP102",
  "SP201", "SP201J", "SP203", "SP203L", "SP223", "SP303", "SP304", "SP305", "SP323", "SP1200", "SP1300", "SP4200",
  "SP4300", "SP8200", "SRMi", "Strata GX", "Strata GXE", "Strata GXE PTK", "Strata GXE PTK2", "Strata GXE PTK3",
  "Strata HX", "Strata HX PTK", "Strata JX", "Striker Carbide CK", "Striker Carbide CKS", "Striker Carbide DA",
  "Striker Carbide DK", "Striker Carbide DKG", "Striker Ember MDM", "Striker Halo FXM", "Striker Halo FXML",
  "Striker Nitride", "Striker Nitride FN", "Striker Nitride GN", "Striker Nitride Versa-A", "Striker Oxide D",
  "Striker Oxide D on Excel Platform", "Striker Oxide EP", "Striker Oxide EPX", "Striker Oxide FE",
  "Striker Oxide FE PTK1", "Striker Oxide FH", "Striker Oxide FP", "Striker Oxide FS", "Striker Oxide GE",
  "Striker SPARC GL", "Striker SPARC GN PTK 1", "Striker SPARC GNE", "Strip45", "Syndion", "Syndion C", "Syndion FS",
  "Syndion GP", "Syndion GS", "Syndion HD", "Syndion HX", "Synergy", "Synergy Integra", "Synergy Performa",
  "TCP 9100 Alliance", "TCP 9100PTX Alliance", "TCP 9400", "TCP 9400 Alliance", "TCP 9400DFM", "TCP 9400DFM Alliance",
  "TCP 9400DFM-P Alliance", "TCP 9400DSiE Alliance", "TCP 9400DSiE II Alliance", "TCP 9400DSiE III Alliance",
  "TCP 9400PTX Alliance", "TCP 9400SE", "TCP 9400SE Alliance", "TCP 9400SE II", "TCP 9400SE II Alliance", "TCP 9600",
  "TCP 9600 Alliance", "TCP 9600DFM Alliance", "TCP 9600DFM-P Alliance", "TCP 9600NV Alliance C",
  "TCP 9600PTX Alliance", "TCP 9600SE", "TCP 9600SE Alliance", "TCP 9600SE II", "TCP 9600SE II Alliance", "TELOS",
  "Teres 200mm CMP", "Teres 300mm CMP", "Torus 200", "Torus 300", "Torus 300 B", "Torus 300 LP 2PM", "Torus 300 LP 3PM",
  "Torus 300 S", "TPD", "Treos Duo", "Treos Echo", "Triton", "UST", "Vantex B", "Vantex BX", "Vantex BX+", "Vantex C",
  "Vantex CX", "Vantex CX+", "Vantex CXE", "Vantex DS", "Vantex DX", "VECTOR", "VECTOR 450", "VECTOR AHM HCE",
  "VECTOR AHM HX", "VECTOR AHM HXE", "VECTOR AHM HXi", "VECTOR Core 3D", "VECTOR DT", "VECTOR DT E", "VECTOR DT EX",
  "VECTOR Excel ALD Nitride (Alpha)", "VECTOR Excel CLEAR", "VECTOR Excel DB", "VECTOR Excel DBi", "VECTOR Excel EBR",
  "VECTOR Excel LUMIER", "VECTOR Excel MD", "VECTOR Excel Versa", "VECTOR Excel Versa-D", "VECTOR Excel Versa-G",
  "VECTOR Excel Versa-M", "VECTOR Excel Versa-P", "VECTOR Excel Versa-QE", "VECTOR Excel Versa-T", "VECTOR Express",
  "VECTOR Express AHM", "VECTOR Express ALD", "VECTOR Extreme AHM GX", "VECTOR Extreme AHM xT", "VECTOR Extreme Core",
  "VECTOR Extreme Core Nano", "VECTOR Extreme Corei", "VECTOR Extreme EBR", "VECTOR Extreme ILDS",
  "VECTOR Extreme Strata", "VECTOR Extreme TEOS xT", "VECTOR F47", "VECTOR HPO", "VECTOR HPO PTK", "VECTOR HT",
  "VECTOR HTE", "VECTOR MD Duo", "VECTOR MD G", "VECTOR MD GL", "VECTOR MD H", "VECTOR Q Strata", "VECTOR TEOS 3D",
  "VECTORi", "Versys Bias Strip", "Versys Kiyo", "Versys Kiyo SNC", "Versys Kiyo3x", "Versys Kiyo45", "Versys Metal",
  "Versys Metal FX", "Versys Metal HP", "Versys Metal L", "Versys Metal M", "Versys Metal MX", "Versys Metal N",
  "Versys Metal NXP", "Versys Metal45", "Versys Silicon", "Versys Star", "Versys TCP Strip Module", "Versys W", "VPM",
  "VTMi", "WCO", "WTS", "WTS-HV", "WTS-Max", "WTSi", "Xceda", "XM Series",
];
// A row keeps its FID Systems as one text value: labels on separate lines, and an empty line for Auto.
function partChangesSystemList(value){return [...new Set(String(value??'').split('\n').map(label=>label.trim()))]}
function partChangesSystemValue(labels){
  const unique=[...new Set(labels)];
  return unique.every(label=>!label)?'':unique.join('\n');
}

// System Description joins SAP short codes with "/", for example E6/KiyoGX ARIA STANZAx3.
// Each code names a platform or a module. These rules spell a code the way MyLam names the System.
const PART_CHANGES_SYSTEM_CODES=[
  [/^E([456])$/,'2300e$1 Transport Module'],[/^A$/,'Alliance Transport Module'],[/^MWST$/i,'2300 Microwave Stripper'],[/^ST45$/i,'Strip45'],
  [/^PC[_ ]MAX$/i,'ALTUS PreClean Max (PC Max)'],[/^PC$/,'ALTUS PreClean (PC)'],
  [/^DF[_ ]WN[_ ]MAX/i,'ALTUS DirectFill Max (WN Max) '],[/^DF[_ ]WN/i,'ALTUS DirectFill (WN) '],
  [/^DV-?(\d\d)(?!\d)/i,'Da Vinci (DV-$1) '],[/^DV-?P(?![a-z])/,'DV-Prime '],[/^EOS\d/i,'EOS'],[/^WTSC/i,'WTS'],
  [/^EXHPT/i,'Exelan HPT Alliance '],[/^EXHP/i,'Exelan HP Alliance '],[/^94(?:00)?(?!\d)/,'TCP 9400'],[/^96(?:00)?(?!\d)/,'TCP 9600'],
  [/^23F45/i,'Exelan Flex45 '],[/^23F3X/i,'Exelan Flex3x '],[/^23FL(?:EX)?/i,'Exelan Flex '],
  [/^MET(?!AL|IOR|RYX)/i,'Versys Metal '],[/^Synd(?!ion)/i,'Syndion '],[/^STR[_ ]/i,'Strata '],[/^ALT[_ ]/i,'ALTUS '],
  [/^SAB(?!RE)[_ ]?/i,'SABRE '],[/^(SABRE )EXC(?![A-Z])/,'$1Excel '],[/^(SABRE )EXT(?![A-Z])/,'$1Extreme '],[/^(SABRE )ANNEAL/i,'$1Anneal Module '],
  [/^EX[TC][_ ]STR(?:IKER|K)[_ ]/i,'Striker '],[/^(Striker )OX[_ ]/,'$1Oxide '],[/^(Striker )C[_ ]/,'$1Carbide '],[/^(Striker )NI[_ ]/,'$1Nitride '],
  [/^EXC[_ ]SOLA/i,'SOLA Excel '],[/^EXC[_ ]VER-/i,'VECTOR Excel Versa-'],
  // A VECTOR module code may leave out the platform word that MyLam also leaves out.
  [/^EXT[_ ]/i,['VECTOR Extreme ','VECTOR ']],[/^EXC[_ ]/i,['VECTOR Excel ','VECTOR ']],
  [/[_ ]PLUS\b/i,'+'],[/[_ ]SPD\b/i,' SPEED'],[/[_ ]SQL\b/i,' SEQUEL'],
];

function createPartChangesSystemGuesser(labels){
  const key=value=>String(value||'').normalize('NFKC').toUpperCase().replace(/[^A-Z0-9+]/g,'');
  const tokens=value=>String(value||'').normalize('NFKC').match(/[A-Z]+(?![a-z])\+?|[A-Z]?[a-z]+\+?|\d+\+?/g)||[];
  const byKey=new Map(),endings=new Map(),guesses=new Map();
  for(const label of labels){
    const name=key(label);
    if(!name||byKey.has(name))continue;
    byKey.set(name,label);
    // "MetalNXP" and "Kiyo45" leave out the family word of "Versys Metal NXP" and "Versys Kiyo45".
    const words=tokens(label).map(key);
    for(let start=1;start<words.length;start++){
      const ending=words.slice(start).join('');
      if(ending.length>3)endings.set(ending,endings.has(ending)?'':label);
    }
  }
  const names=[...byKey.keys()];
  function spellings(code){
    let list=[code];
    for(const [pattern,replacement] of PART_CHANGES_SYSTEM_CODES)
      list=list.flatMap(text=>pattern.test(text)?[].concat(replacement).map(value=>text.replace(pattern,value)):[text]);
    return list;
  }
  function match(code){
    let best=[],score=0;
    const offer=(found,value)=>{if(value>score){best=found;score=value}};
    for(const spelling of spellings(code)){
      const words=tokens(spelling).map(key);
      for(let start=0;start<words.length;start++){
        const rest=words.slice(start).join('');
        // The whole remaining code names one System.
        if(byKey.has(rest)){offer([byKey.get(rest)],3000+rest.length);continue}
        if(!start&&endings.get(rest))offer([endings.get(rest)],2500+rest.length);
        // A family code such as ALT_LFW or ICEFILL MAX covers its few variants.
        const family=rest.length>(start?6:4)?names.filter(name=>name.startsWith(rest)):[];
        if(family.length&&family.length<4)offer(family.map(name=>byKey.get(name)),2000+rest.length);
        // Otherwise the System name starts the code, followed by a kit, chamber or variant suffix.
        let boundary='';
        const boundaries=new Set(words.slice(start).map(word=>boundary+=word));
        for(const name of names)
          if(rest.startsWith(name)&&name.length>(boundaries.has(name)?2:5))offer([byKey.get(name)],1000+name.length);
      }
    }
    return best;
  }
  function guessText(text){
    const codes=String(text||'').split(/[\/,;|\n]+/).map(code=>code.trim().replace(/x\d+$/,'').trim()).filter(Boolean);
    const alliance=codes[0]==='A',found=new Set();
    for(const code of codes){
      if(!guesses.has(code))guesses.set(code,match(code));
      for(const label of guesses.get(code)){
        // An Alliance platform uses the Alliance edition of its modules when MyLam lists one.
        const edition=alliance&&byKey.get(key(label+' Alliance'));
        found.add(edition||label);
      }
    }
    return [...found];
  }
  return (description,product)=>{
    const found=guessText(description);
    return found.length?found:guessText(product);
  };
}

function normalizedBomPartNumber(value){return String(value??'').normalize('NFKC').replace(/[\u2010-\u2015\u2212]/g,'-').replace(/\u00ad/g,'').replace(/\s+/g,'').toUpperCase();}

function bomPartDifferences(boms){
  const parts=new Map();
  boms.forEach((bom,index)=>{
    for(const key of bom.materials.keys()){
      if(!parts.has(key))parts.set(key,[]);
      parts.get(key).push(index);
    }
  });
  if(boms.length>1)for(const [key,indexes] of parts)if(indexes.length===boms.length)parts.delete(key);
  return parts;
}

function recoverMalformedBomRows(buffer){
  // Some BOM exports use numeric column positions in cell references (for example " 42" instead of D2).
  const workbook=XLSX.read(buffer,{type:'array',bookFiles:true}),content=workbook.files?.['xl/worksheets/sheet1.xml']?.content;
  if(!content)return null;
  const xml=new TextDecoder().decode(content),firstCell=xml.match(/<c\b[^>]*\br="([^"]*)"/);
  if(!firstCell||/^[A-Z]+\d+$/i.test(firstCell[1]))return null;
  const rows=[];
  for(const match of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)){
    const rowNumber=Number(match[1].match(/\br="(\d+)"/)?.[1]);if(!rowNumber)continue;
    const values=[];let nextColumn=0;
    for(const cell of match[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
      const reference=cell[1].match(/\br="([^"]*)"/)?.[1]||'',prefix=reference.endsWith(String(rowNumber))?reference.slice(0,-String(rowNumber).length).trim():'',oneBased=Number(prefix); const column=/^\d+$/.test(prefix)&&oneBased>0?oneBased-1:nextColumn;nextColumn=column+1;
      const raw=cell[2]?.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/)?.[1];if(raw===undefined||raw==='')continue;
      const type=cell[1].match(/\bt="([^"]+)"/)?.[1]; values[column]=type==='s'?workbook.Strings?.[Number(raw)]?.t??'':type==='n'||!type?(/^\d{16,}$/.test(raw)?raw:Number(raw)):type==='b'?raw==='1':raw;
    }
    rows[rowNumber-1]=values;
  }
  return rows.length?rows:null;
}

async function readBomFile(file,{requireDescription=false,materialsOnly=false}={}){
  if(typeof XLSX==='undefined'||!XLSX.read)throw new Error('Excel reader is unavailable. Refresh the page and try again.');
  let workbook,buffer;
  try{buffer=await file.arrayBuffer();workbook=XLSX.read(buffer,{type:'array',cellDates:true,cellStyles:!materialsOnly})} catch(error){throw new Error(`${file.name}: could not read the workbook (${error.message}).`)}
  const sourceSheetName=workbook.SheetNames[0],source=workbook.Sheets[sourceSheetName];
  if(!source)throw new Error(`${file.name}: the workbook has no worksheet.`);
  const decoded=XLSX.utils.decode_range(source['!ref']||'A1'),range={s:{r:0,c:0},e:decoded.e};
  let rows,display=(row,column)=>{const cell=source[XLSX.utils.encode_cell({r:row,c:column})]; return !cell?'':typeof cell.v==='number'&&Number.isInteger(cell.v)&&Math.abs(cell.v)>=
    1e15?String(cell.v):XLSX.utils.format_cell(cell)};
  let lastRow=decoded.e.r,lastColumn=decoded.e.c,headerRow=-1,materialColumn=-1;
  function findHeader(){findMaterial:for(let row=0;row<=Math.min(lastRow,199);row++)for(let column=0;column<=lastColumn;column++){if(String(display(row,column)).replace(/\s+/g,' ').trim().toUpperCase()==='MATERIAL'){headerRow=row;materialColumn=column;break findMaterial}}}
  findHeader();
  if(headerRow<0){const recovered=recoverMalformedBomRows(buffer); if(recovered){rows=recovered;lastRow=rows.length-1;lastColumn=Math.max(0,...rows.slice(0,200).map(row=>row?.length||0))-1;display=(row,column)=>rows[row]?.[column]??'';findHeader()}}
  if(headerRow<0)throw new Error(`${file.name}: no MATERIAL column was found in the first worksheet.`);
  const materials=new Map(),descriptions=new Map(),partKeys=[];
  let descriptionColumn=-1;
  if(!materialsOnly){
    for(let column=0;column<=lastColumn;column++)if(/^DESCRIPTION$/i.test(String(display(headerRow,column)).trim())){descriptionColumn=column;break}
  }
  if(requireDescription&&descriptionColumn<0)throw new Error(file.name+': no DESCRIPTION column was found.');
  for(let row=headerRow+1;row<=lastRow;row++){
    const shown=String(display(row,materialColumn)).replace(/[\u2010-\u2015\u2212]/g,'-').replace(/\u00ad/g,'').trim(),key=normalizedBomPartNumber(shown);
    if(!key)continue;
    if(!materials.has(key))materials.set(key,shown);
    if(!materialsOnly)partKeys[row]=key;
    if(descriptionColumn>=0&&!descriptions.has(key)){
      const description=String(display(row,descriptionColumn)).trim();
      if(description)descriptions.set(key,description);
    }
  }
  if(requireDescription&&!materials.size)throw new Error(file.name+': the MATERIAL column is empty.');
  if(materialsOnly)return {materials};
  rows??=XLSX.utils.sheet_to_json(source,{header:1,range,raw:true,defval:'',blankrows:true});
  return {file,fileName:file.name,sourceSheetName,rows,headerRow,materialColumn,materials,descriptions,partKeys, merges:(source['!merges']||[]).map(rangeItem=>XLSX.utils.encode_range(rangeItem)), columnWidths:(source['!cols']||[]).map(column=>Number(column.wch)||(Number(column.wpx)?Number(column.wpx)/
    7:0)), rowHeights:(source['!rows']||[]).map(row=>Number(row.hpt)||(Number(row.hpx)?Number(row.hpx)*0.75:0))};
}

(function(){
const WORK_NOTES_KEY='toolkitKyra.workNotes';
const FID_NOTES_KEY='toolkitKyra.fid.notes.v1',FID_ID_PATTERN=/^\d{5,10}$/;

function normalizeFidNotes(notes){
  return Object.fromEntries(Object.entries(notes||{})
    .filter(([fid,note])=>FID_ID_PATTERN.test(fid)&&typeof note==='string'&&note.trim())
    .map(([fid,note])=>[fid,note.replace(/\r\n?/g,'\n')]));
}

function readFidNotes(){
  const saved=GM_getValue(FID_NOTES_KEY,null);
  return normalizeFidNotes(typeof saved==='string'?JSON.parse(saved):saved);
}

function mountWorkNotes(onFidNotesChanged){
  const host=document.createElement('div');
  host.id='work-notes-host';
  const root=host.attachShadow({mode:'open'});
  root.innerHTML=`
    <style>
      :host{position:fixed;right:20px;bottom:20px;z-index:2147483647;font:13px/1.5 system-ui,sans-serif;color:#303030}
      :host([hidden]),.panel[hidden]{display:none}
      *,*::before,*::after{box-sizing:border-box}
      button,select,textarea{font:inherit;color:inherit}
      button{cursor:pointer}
      button:disabled{cursor:default;opacity:.5}
      .bubble{display:grid;place-items:center;width:49px;height:49px;padding:0;border:0;border-radius:16px;background:#ffda64;color:#795800;box-shadow:0 4px 15px #0002}
      .bubble:hover{background:#ffd24b}
      .bubble svg{width:27px;height:27px}
      .panel{display:flex;flex-direction:column;position:absolute;right:0;bottom:61px;width:min(325px,calc(100vw - 40px));height:min(365px,calc(100dvh - 101px));overflow:hidden;background:#fffefa;border:1px solid #e6d69b;border-radius:13px;box-shadow:0 9px 35px #0003}
      .header,.toolbar,.buttons{display:flex;align-items:center;gap:7px}
      .header{justify-content:space-between;padding:10px 12px;background:#ffedaa;font-weight:650}
      .close{border:0;padding:2px 6px;background:transparent;font-size:19px;line-height:1}
      .toolbar{padding:9px;border-bottom:1px solid #eee8d4}
      select{flex:1;min-width:0;height:32px;padding:4px 6px;border:1px solid #ddd2a6;border-radius:6px;background:white}
      .action{border:1px solid #e2d7b8;background:white;border-radius:7px;padding:5px 10px}
      .action:hover:not(:disabled){background:#fff6d5}
      .icon{display:grid;place-items:center;width:32px;height:32px;padding:0;flex:none}
      .icon svg{width:17px;height:17px}
      textarea{flex:1;min-height:0;width:100%;padding:12px;border:0;resize:none;background:transparent;line-height:1.65}
      textarea:focus-visible{outline:2px solid #bf8b12;outline-offset:-2px}
      .footer{padding:7px 9px 9px;border-top:1px solid #eee8d4}
      .status{display:block;margin-bottom:5px;color:#72674d;font-size:11px;overflow-wrap:anywhere}
      .buttons{justify-content:flex-end;flex-wrap:wrap}
    </style>
    <section id="panel" class="panel" aria-label="Work Notes" hidden>
      <div class="header"><span>📝 Notes</span><button class="close" type="button" id="close" aria-label="Close Notes" title="Close">×</button></div>
      <div class="toolbar">
        <select id="list" aria-label="Choose a note"></select>
        <button class="action icon" type="button" id="add" aria-label="New note" title="New note">＋</button>
        <button class="action icon" type="button" id="remove" aria-label="Delete note" title="Delete note">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 7h16M10 4h4M6.5 7l.7 13h9.6l.7-13M10 11v6M14 11v6"/></svg>
        </button>
      </div>
      <textarea id="editor" aria-label="Note text" placeholder="Write a note…"></textarea>
      <div class="footer">
        <span id="status" class="status" role="status">Saved locally</span>
        <div class="buttons"><button class="action" type="button" id="copy">Copy</button><button class="action" type="button" id="move">Move to FID Note</button></div>
      </div>
    </section>
    <button class="bubble" type="button" id="toggle" title="Notes" aria-label="Notes" aria-controls="panel" aria-expanded="false">
      <svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M9 3.8h13a5 5 0 0 1 5 5V18l-9 9H9a5 5 0 0 1-5-5V8.8a5 5 0 0 1 5-5Z"/>
        <path d="M18 27v-5a4 4 0 0 1 4-4h5"/>
      </svg>
    </button>`;
  const id=name=>root.getElementById(name);
  const toggle=id('toggle'),panel=id('panel'),list=id('list'),editor=id('editor'),status=id('status');
  const add=id('add'),remove=id('remove'),copy=id('copy'),move=id('move');
  let moving=false;

  function readDrafts(){
    const saved=GM_getValue(WORK_NOTES_KEY,[]);
    return Array.isArray(saved)?saved.filter(note=>typeof note?.id==='string'&&note.id&&typeof note.text==='string'):[];
  }

  function titleOf(note,index){
    const firstLine=note.text.split(/\r?\n/).find(line=>line.trim());
    return firstLine?firstLine.trim().slice(0,36):`Untitled ${index+1}`;
  }

  function updateControls(){
    add.disabled=moving;
    list.disabled=editor.disabled=remove.disabled=copy.disabled=moving||!list.value;
    move.disabled=editor.disabled||!editor.value.trim();
  }

  function showSelectedDraft(drafts){
    const text=drafts.find(note=>note.id===list.value)?.text??'';
    if(editor.value!==text)editor.value=text;
    updateControls();
  }

  function refreshList(drafts,selectedId=list.value){
    list.replaceChildren(...drafts.map((note,index)=>new Option(titleOf(note,index),note.id)));
    if(drafts.some(note=>note.id===selectedId))list.value=selectedId;
    showSelectedDraft(drafts);
  }

  function saveDrafts(drafts){
    try{
      GM_setValue(WORK_NOTES_KEY,drafts);
      status.textContent='Saved locally';
      return true;
    }catch(error){
      status.textContent='Could not save notes. '+error.message;
      return false;
    }
  }

  function addNote(drafts){
    const note={id:`${Date.now()}-${Math.random().toString(36).slice(2,7)}`,text:''};
    const next=[note,...drafts];
    if(saveDrafts(next)){
      refreshList(next,note.id);
      editor.focus({preventScroll:true});
    }
  }

  function closePanel(){
    panel.hidden=true;
    toggle.setAttribute('aria-expanded','false');
    toggle.focus({preventScroll:true});
  }

  toggle.addEventListener('click',()=>{
    if(!panel.hidden){closePanel();return}
    panel.hidden=false;
    toggle.setAttribute('aria-expanded','true');
    const drafts=readDrafts();
    refreshList(drafts);
    if(!drafts.length&&!moving)addNote(drafts);
    else editor.focus({preventScroll:true});
  });
  toggle.addEventListener('contextmenu',event=>{
    event.preventDefault();
    if(confirm('Hide Notes until this page is reloaded?')){
      closePanel();
      host.hidden=true;
    }
  });
  id('close').addEventListener('click',closePanel);
  add.addEventListener('click',()=>addNote(readDrafts()));
  list.addEventListener('change',()=>{
    showSelectedDraft(readDrafts());
    status.textContent='Saved locally';
    editor.focus({preventScroll:true});
  });
  editor.addEventListener('input',()=>{
    const drafts=readDrafts(),index=drafts.findIndex(note=>note.id===list.value);
    if(index===-1){refreshList(drafts);return}
    const note={...drafts[index],text:editor.value};
    drafts[index]=note;
    if(saveDrafts(drafts))list.selectedOptions[0].textContent=titleOf(note,index);
    updateControls();
  });
  copy.addEventListener('click',()=>GM_setClipboard(editor.value));
  remove.addEventListener('click',()=>{
    const selectedId=list.value;
    if(editor.value&&!confirm('Delete this note?'))return;
    const next=readDrafts().filter(note=>note.id!==selectedId);
    if(saveDrafts(next))refreshList(next);
  });
  move.addEventListener('click',async()=>{
    if(typeof GM==='undefined'||typeof GM.setValues!=='function'){
      status.textContent='Update Tampermonkey to move notes.';
      return;
    }
    const selectedId=list.value,text=editor.value;
    const input=prompt('Enter the target FID (5–10 digits).');
    if(input===null)return;
    const fid=input.trim();
    if(!FID_ID_PATTERN.test(fid)){status.textContent='Enter a FID with 5–10 digits.';return}

    moving=true;
    updateControls();
    try{
      const drafts=readDrafts(),draft=drafts.find(note=>note.id===selectedId);
      if(draft?.text.replace(/\r\n?/g,'\n')!==text)throw new Error('The draft changed. Review it before moving it.');
      const notes=readFidNotes(),previous=notes[fid];
      notes[fid]=previous?previous+'\n\n'+text:text;
      const next=drafts.filter(note=>note.id!==selectedId);
      // Store the destination and remove the draft in the same storage update.
      await GM.setValues({[FID_NOTES_KEY]:JSON.stringify(notes),[WORK_NOTES_KEY]:next});
      refreshList(readDrafts());
    }catch(error){
      status.textContent='Could not move note. '+error.message;
      return;
    }finally{
      moving=false;
      updateControls();
    }
    status.textContent='Moved to FID '+fid+'.';
    onFidNotesChanged?.();
    if(!panel.hidden)(editor.disabled?add:editor).focus({preventScroll:true});
  });
  root.addEventListener('keydown',event=>{
    if(panel.hidden)return;
    event.stopPropagation();
    if(event.key==='Escape'){
      event.preventDefault();
      closePanel();
    }
  });
  if(typeof GM_addValueChangeListener==='function')GM_addValueChangeListener(WORK_NOTES_KEY,(_key,_old,_value,remote)=>{
    if(remote&&!panel.hidden&&!moving)refreshList(readDrafts());
  });
  document.body.append(host);
  return host;
}
'use strict';
const [versionYear,versionMonth,versionDay]=GM_info.script.version.split('.').map(Number);
const date_v = new Date(2000+versionYear,versionMonth-1,versionDay); const APP_PATH='/k1.png'; const BRIDGE_PATH='/k2.png'; const FCID_PATH='/staceykyra.png'; const NSR_HEADER_BUTTON_ID='toolkitKyraControlTowerButton';
const GUIDE_URL='https://synalocey.github.io/ToolkitKyra/guide';
const FID_DASHBOARD_TITLE='🏭 FID & NSR Tracking Dashboard';
const DASHBOARD_BASE_CSS=chosenBaseCss+`
:root{--bg:#f3f6fc;--panel:#fff;--ink:#24354e;--muted:#53657d;--line:#d8e1ee;--blue:#2563eb;--teal:#0f766e;--green:#16a34a;--amber:#d97706;--red:#dc2626;--radius:12px;--font-sans:"Segoe UI Variable Text","Segoe UI","Microsoft YaHei UI",Arial,sans-serif;--shadow:0 3px 16px #253c6410}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:13px/1.45 var(--font-sans)}h1,h2,h3,b,strong{font-weight:600}button,input,select,textarea{font:inherit}button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible,summary:focus-visible{outline:3px solid #8ab4f2;outline-offset:2px}.card,.side{background:#fff;border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow)}
.chosen-container{font-family:var(--font-sans);width:100%!important}.chosen-container-multi .chosen-choices{background-image:none;min-height:34px}.chosen-container-multi .chosen-choices li.search-field input{height:26px;line-height:26px;font-family:inherit}.chosen-container-multi .chosen-choices li.search-choice .search-choice-close{top:50%;right:5px;width:14px;height:14px;margin-top:-7px;background:none;font-size:0}.chosen-container-multi .chosen-choices li.search-choice .search-choice-close::after{content:"×";display:block;color:#365c95;font-size:15px;line-height:13px;text-align:center}.chosen-container .chosen-drop{border-color:#b8c8d8;border-radius:0 0 9px 9px;box-shadow:0 9px 24px #263b6425}.chosen-container .chosen-results{max-height:230px}
.section-title{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:12px}.section-title h2{margin:0}.context-badge{display:inline-flex;align-items:center;background:#e7f6ed;color:#22604b;border:1px solid #b8dcc7;border-radius:18px;padding:4px 9px;font-size:12px;font-weight:600;white-space:nowrap}.toolbar-right{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:9px;background:#fff;scrollbar-width:thin}
.status-pill{display:inline-flex;border-radius:12px;padding:3px 7px;font-size:12px;line-height:1.3;background:#dbeafe;color:#1d4ed8}.status-pill.terminal{background:#dcfce7;color:#166534}.status-pill.stop{background:#fee2e2;color:#991b1b}.status-pill.unknown{background:#fef3c7;color:#92400e}
.cycle-shell{display:grid;gap:14px}.cycle-label{display:block;color:#425a7b;font-size:12px;font-weight:600;margin-bottom:6px}.cycle-input{display:block;width:100%;resize:vertical;min-height:110px;border:1px solid var(--line);border-radius:9px;background:#fff;color:var(--ink);padding:10px;font:13px/1.5 Consolas,"Segoe UI",sans-serif}.cycle-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:9px}.cycle-progress{display:block;width:100%;height:6px;margin-top:12px;accent-color:var(--teal);appearance:none;border:0;border-radius:9px;overflow:hidden;background:#e9eff4}.cycle-progress::-webkit-progress-bar{background:#e9eff4}.cycle-progress::-webkit-progress-value,.cycle-progress::-moz-progress-bar{background:var(--teal)}.cycle-status{margin-top:7px;color:var(--muted);font-size:12px}.cycle-status:empty{display:none}.ct-link{border:0;background:none;color:#0369a1;padding:0;font:inherit;text-decoration:underline;text-underline-offset:2px;cursor:pointer}.duration-over,.duration-over .ct-link{color:var(--red)!important}
.wf-modal:not([open]){display:none}.wf-modal{position:fixed;inset:0;z-index:1000;display:grid;place-items:center;width:100%;height:100%;max-width:none;max-height:none;margin:0;padding:18px;border:0;background:transparent;overflow:hidden}.wf-modal::backdrop{background:#152e505c;backdrop-filter:blur(2px)}.wf-bg{position:absolute;inset:0;width:100%;height:100%;border:0;background:transparent}.wf-box.fid-dialog{--table-heading-bg:initial;position:relative;width:min(1200px,calc(100vw - 36px));min-width:0;height:auto;min-height:0;max-height:90vh}.wf-head p{margin:4px 0 0;font-size:13px}#wfTable .tabulator-col-title{white-space:normal;line-height:17px}#wfTable .tabulator-row .tabulator-cell{height:auto;white-space:pre-wrap}#wfTable .cell-content{white-space:pre-wrap;max-height:none;overflow-wrap:anywhere}body.wf-open{overflow:hidden}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation:none!important;transition:none!important}}
`;
const NSR_CYCLE_HTML=`<section id="cycleTab" class="tab-page" role="tabpanel" aria-label="NSR Cycle Time">
  <div class="cycle-shell">
    <section class="card">
      <div class="section-title">
        <div><h2>NSR Cycle Time</h2></div>
        <span id="cycleBadge" class="context-badge">Ready</span>
      </div>
      <label class="cycle-label" for="cycleInput">NSR# list</label>
      <textarea id="cycleInput" class="cycle-input" rows="6" placeholder="Enter one NSR# per line, for example:&#10;N123456&#10;N123457"></textarea>
      <div class="cycle-actions">
        <button id="runCycleBtn" class="btn blue" type="button">Query</button>
        <button id="cancelCycleBtn" class="btn light" type="button" disabled>Cancel</button>
      </div>
      <progress id="cycleProgress" class="cycle-progress" max="1" value="0"></progress>
      <div id="cycleStatus" class="cycle-status"></div>
    </section>

    <section class="card">
      <div class="section-title detail-head">
        <h2>Cycle Time Detail</h2>
      </div>
      <div id="cycleTable" aria-label="NSR cycle time detail"></div>
    </section>
  </div>
</section>
<dialog id="wfModal" class="wf-modal" aria-labelledby="wfTitle">
  <button class="wf-bg" type="button" data-wf-close aria-label="Close workflow detail" tabindex="-1"></button>
  <section class="wf-box fid-dialog">
    <div class="section-head wf-head">
      <div><h1><span id="wfTitle">Workflow Detail</span><a id="wfNsrLink" class="fid-action" target="_blank" rel="noopener noreferrer" hidden>NSR ↗</a></h1><p id="wfMeta" class="muted"></p></div>
      <button id="wfClose" class="fid-close" type="button" data-wf-close aria-label="Close workflow detail">×</button>
    </div>
    <p id="wfStatus" class="muted" role="status" hidden></p>
    <div id="wfSummary" class="fid-dialog-grid"></div>
    <div id="wfAdditional"></div>
    <div id="wfTable" aria-label="Full NSR workflow detail"></div>
  </section>
</dialog>
`;
const WORKSPACE_CSS=`
.btn{border:1px solid transparent;border-radius:8px;background:var(--teal);color:#fff;font-weight:600;line-height:1.35;padding:8px 12px;cursor:pointer}.btn:disabled{cursor:not-allowed;opacity:.48}.btn:hover:not(:disabled){filter:brightness(.96)}
:is(.fcid-app,.toolbox-app) .btn{color:#fff;font-size:12px;padding:7px 10px;font-weight:600;white-space:nowrap}:is(.fcid-app,.toolbox-app) .btn.small{padding:4px 7px}:is(.fcid-app,.toolbox-app) .btn.light{background:#edf1f6;border:1px solid #ccd6e3;color:#354d6d}:is(.fcid-app,.toolbox-app) .btn.blue{background:#2863d6;color:#fff}
.fc-hero{padding:15px 18px;background:linear-gradient(110deg,#e8f3ff,#f0ebff 57%,#fff0e4);border-bottom:1px solid #ccd9ed}.fc-hero-inner{max-width:1900px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:14px}.fc-hero .fc-hero-inner h1{margin:3px 0 0;font-size:27px;letter-spacing:-.5px}
.fc-brand{font-size:12px;letter-spacing:1.5px;color:#4d5b89}.fc-heading{display:flex;align-items:center;gap:13px}.btnrow,.fc-hero-actions{display:flex;gap:7px;align-items:center;flex-wrap:wrap}
.layout{max-width:1948px;margin:auto;display:grid;grid-template-columns:var(--side) minmax(0,1fr);gap:14px;padding:14px;align-items:start}.layout>.main{min-width:0;display:grid;grid-template-columns:minmax(0,1fr);gap:12px}.layout>aside.side{position:sticky;top:12px;max-height:calc(100vh - 24px);overflow:auto;scrollbar-width:thin;padding:13px;min-width:0}
.section-head{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:12px}
.fc-source{position:relative;border:1px solid #dce7ed;background:#fbfdfe;border-radius:9px;padding:9px 8px;margin-top:8px}.fc-source-head{display:flex;align-items:center;justify-content:space-between;gap:5px;margin-bottom:7px}.fc-source .fc-source-head h3{display:flex;align-items:center;gap:5px;min-width:0;font-size:12px;color:#24364b}.fc-source-name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.fc-source-number{display:inline-flex;align-items:center;justify-content:center;flex:none;width:22px;height:22px;background:#376bd2;color:#fff;border-radius:6px;font-size:12px;font-weight:600}.fc-source[data-source=ctl] .fc-source-number{background:#8152bf}.fc-source[data-source=ctlNsr] .fc-source-number{background:#a85c07}.fc-source[data-source=nsr] .fc-source-number{background:#bf3d74}.fc-source[data-source=nce] .fc-source-number{background:#087e70}.fc-source-state{display:inline-flex;align-items:center;gap:5px;flex:none;max-width:105px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#53657d;font-size:12px;line-height:1.5;font-variant-numeric:tabular-nums}.fc-source-state.has-range{max-width:none}.fc-source-state::before{content:"";display:inline-block;flex:none;width:5px;height:5px;border-radius:50%;background:#94a3b8}.fc-source-state.ok{color:#166534}.fc-source-state.ok::before{background:#158653}.fc-source-state.stale{color:#d97706}.fc-source-state.stale::before{background:#d97706}.fc-source-state.loading{color:#92400e}.fc-source-state.loading::before{background:#d97706}.fc-source-state.error{cursor:help}.fc-source-toolbar{display:flex;align-items:center;gap:6px;min-width:0}.fc-source-toolbar>.btn{flex:none}.fc-source-file{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#53657d;font-size:12px;line-height:1.5}.fc-source-help{position:relative}.fc-source-help summary{list-style:none}.fc-source-help-bubble{position:absolute;right:0;top:30px;width:275px;z-index:30;padding:10px;background:#fff;border:1px solid #c4d0e4;border-radius:8px;box-shadow:0 8px 26px #23395630;font-size:12px}
.workspace-nav{display:flex;gap:12px;flex-wrap:wrap;padding:6px;background:#fff;border:1px solid var(--line);border-radius:10px}.workspace-nav button{border:0;background:transparent;color:#485d7b;font-size:13px;font-weight:600;padding:8px 10px;border-radius:7px;cursor:pointer}.workspace-nav button:is(.active,[aria-selected=true]){background:#e4edff;color:#1e52ad}.workspace-nav button:hover{background:#eef3fd}.workspace-nav button[data-view=rules]{margin-left:auto}
@media(max-width:1250px){.fc-hero .fc-hero-inner h1{font-size:25px}}
@media(max-width:760px){.layout{grid-template-columns:minmax(0,1fr);padding:10px;gap:10px}.layout>aside.side{position:static;max-height:none}.fc-hero{padding:12px}.fc-hero .fc-hero-inner h1{font-size:21px}.fc-hero-inner{align-items:flex-start;flex-wrap:wrap}.fc-hero-actions{margin-left:auto}.workspace-nav button{padding:7px}.section-head{flex-wrap:wrap}}`;
function workspaceHeader(title,refreshId,refreshAttributes,sideToggle=''){
  return `<header class="fc-hero"><div class="fc-hero-inner"><div class="fc-heading">${sideToggle}<div><div class="fc-brand">TOOLKIT</div><h1>${title}</h1></div></div>
    <div class="fc-hero-actions"><button id="${refreshId}" class="btn light" type="button" ${refreshAttributes}>Refresh</button><a class="btn light" href="${GUIDE_URL}" target="_blank" rel="noopener noreferrer">Help</a></div></div></header>`;
}

function sourceMenu(id,label,content,hidden=false){
  return `<details class="fc-source-help" id="${id}"${hidden?' hidden':''}><summary class="btn light small" aria-label="${label}">...</summary><div class="fc-source-help-bubble">${content}</div></details>`;
}

function closeSourceMenus(target){
  document.querySelectorAll('.fc-source-help[open]').forEach(menu=>{if(!target||!menu.contains(target))menu.open=false});
}

const FEP_ORIGIN='https://fep.lamresearch.com',CACHE_PREFIX='fcid-governance:',BW_CACHE_PREFIX=CACHE_PREFIX+'bobj:';
const BW_SOURCES={fcid:{label:'FID',query:'PHCPSYS10R_QR001_REG',file:'BOBJ_FID.xlsm',ids:['FID']},ctl:{label:'CTL',query:'PHMACTR1_ACTIVITY_CTL_REP',file:'BOBJ_CTL.xlsm',ids:['Opportunity ID','CTL ID']}};
const SOURCE_DATES={
  fcid:{fields:['Slot Actual'],label:'Slot Actual'}, ctl:{fields:['Creation Date','Create Date'],label:'Creation Date'}, ctlNsr:{fields:['Create Date','Creation Date'],label:'Create Date'},
  nsr:{fields:['Submit Date','Submitted Date','Submission Date'],label:'Submit Date'}, nce:{fields:['Open Date'],label:'Open Date'}
};
GM_registerMenuCommand(FID_DASHBOARD_TITLE,()=>unsafeWindow.open(FEP_ORIGIN+FCID_PATH+'?nsrct=1','_blank','noopener'));
GM_registerMenuCommand('🧰 SAP Toolbox',()=>unsafeWindow.open(FEP_ORIGIN+BRIDGE_PATH+'?nsrbridge=1','_blank','noopener')); GM_registerMenuCommand('📖 Help',()=>unsafeWindow.open(GUIDE_URL,'_blank','noopener'));

function toolkitExpired(){
  const expired=Number.isFinite(date_v.getTime())&&(Date.now()-date_v.getTime())/86400000>270;
  if(expired)document.documentElement.innerHTML='<head><meta charset="utf-8"><title>Toolkit update required</title></head><body><p>Something wrong with toolkit, please update to the latest version.</p><a href="'+GUIDE_URL+'">'+GUIDE_URL+'</a></body>';
  return expired;
}

if(location.origin==='https://synalocey.github.io'&&location.pathname.startsWith('/ToolkitKyra/')){
  const autoUpdate=typeof GM_info.scriptWillUpdate==='boolean'?GM_info.scriptWillUpdate:GM_info.script.options?.check_for_updates;
  const announce=()=>unsafeWindow.postMessage({type:'toolkit:ready',version:GM_info.script.version,autoUpdate:typeof autoUpdate==='boolean'?autoUpdate:null},location.origin);
  unsafeWindow.addEventListener('message',event=>{if(event.source===unsafeWindow&&event.origin===location.origin&&event.data?.type==='toolkit:check')announce();}); announce(); return;
}

const currentPath=unsafeWindow.location.pathname.toLowerCase();
const partChangesLaunch=location.hostname==='127.0.0.1'&&currentPath===PART_CHANGES_LAUNCH_PATH;
if(partChangesLaunch||currentPath===BRIDGE_PATH){
  const hash=location.hash,handoff=partChangesHandoff(hash);
  if(handoff&&(!partChangesLaunch||handoff.port===Number(location.port))){
    // Save the task in this tab before MyLam or SSO can drop the fragment.
    GM_getTab(tab=>GM_saveTab({...tab,[PART_CHANGES_MYLAM_TAB]:handoff},()=>unsafeWindow.location.replace(PART_CHANGES_MYLAM_URL+hash)));
    return;
  }
}
if(currentPath===APP_PATH){if(!toolkitExpired()){const target=new URL(FCID_PATH,unsafeWindow.location.href);target.search=unsafeWindow.location.search;target.hash='nsr';unsafeWindow.location.replace(target.href)}}else if(currentPath===FCID_PATH){whenBody(runFcidGovernance);}
  else if(currentPath!==BRIDGE_PATH&&unsafeWindow.location.hostname.toLowerCase()==='fep.lamresearch.com')addNsrHeaderButton();

function whenBody(callback){if(document.body)callback();else document.addEventListener('DOMContentLoaded',callback,{once:true});}

async function addNsrHeaderButton(){
  while(true){
    const target=document.getElementById('shellAppTitle-button'),title=target&&target.querySelector('.sapUshellAppTitleText');
    if(target&&String(title&&title.textContent||'').trim().toUpperCase()==='NSR'&&!document.getElementById(NSR_HEADER_BUTTON_ID)){
      const button=document.createElement('button');button.id=NSR_HEADER_BUTTON_ID;button.type='button';button.title='Open '+FID_DASHBOARD_TITLE;button.textContent=FID_DASHBOARD_TITLE; button.style.cssText='margin-left:8px;padding:4px 8px;cursor:pointer;white-space:nowrap';
      button.addEventListener('click',event=>{event.preventDefault();event.stopPropagation();unsafeWindow.open(FCID_PATH+'?nsrct=1','_blank','noopener')}); target.insertAdjacentElement('afterend',button);
      if(button.getBoundingClientRect().right>innerWidth)button.textContent=FID_DASHBOARD_TITLE.replace(' Tracking','');
    }
    await new Promise(resolve=>setTimeout(resolve,2000));
  }
}

function persistentStore(){try{return unsafeWindow.localStorage}catch(_){try{return localStorage}catch(__){return null}}}
function emojiFavicon(emoji){
  const icon=document.createElement('link');icon.rel='icon';icon.type='image/svg+xml';icon.sizes='any';
  icon.href='data:image/svg+xml,'+encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><text x="32" y="52" text-anchor="middle" font-size="54" font-family="Segoe UI Emoji, Apple Color Emoji, Noto Color Emoji, sans-serif">${emoji}</text></svg>`); return icon;
}

function setWorkspaceHead(label,emoji,css){
  const charset=document.createElement('meta');charset.charset='utf-8';
  const viewport=document.createElement('meta');viewport.name='viewport';viewport.content='width=device-width,initial-scale=1';
  const title=document.createElement('title');title.textContent=label;
  const style=document.createElement('style');style.textContent=css;
  document.head.replaceChildren(charset,viewport,title,emojiFavicon(emoji),style);
}


// Workers have separate globals; bootstrap them with the cached XLSX resource.
function createWorkbookWorker(code){
  const url=URL.createObjectURL(new Blob([GM_getResourceText('xlsxJs'),'\n;\n',code],{type:'text/javascript'})),worker=new Worker(url);
  URL.revokeObjectURL(url);return worker;
}
function bridgeTransportError(){const error=new Error('Cannot reach SAPGUIcontrol.exe');error.transport=true;return error;}
function bridgeHttpRequest(method,url,payload,timeout){
  return new Promise((resolve,reject)=>{
    if(typeof GM_xmlhttpRequest!=='function'){reject(new Error('Tampermonkey permission for SAPGUIcontrol is unavailable. Reinstall or update this userscript.'));return}
    const headers={Accept:'application/json'};
    if(payload!==undefined)headers['Content-Type']='application/json';
    const options={method,url,headers,timeout:timeout??20000,nocache:true,fetch:false,
      onload:response=>{
        if(!response||!Number(response.status)){reject(bridgeTransportError());return}
        let data={};
        try{data=response.responseText?JSON.parse(response.responseText):{}}catch(_){data={message:response.responseText||'Invalid SAPGUIcontrol response.'}}
        if(response.status>=200&&response.status<300){resolve(data);return}
        const error=new Error(data.message||('SAPGUIcontrol returned HTTP '+response.status+'.'));error.status=response.status;error.code=data.code;error.data=data;reject(error);
      },
      onerror:()=>reject(bridgeTransportError()),ontimeout:()=>reject(bridgeTransportError()),onabort:()=>reject(bridgeTransportError())
    };
    if(options.timeout===0)delete options.timeout;
    if(payload!==undefined)options.data=JSON.stringify(payload);
    try{GM_xmlhttpRequest(options)}catch(error){options.onerror(error?.message)}
  });
}

function createSapGuiControlClient(urls){
  let origin='',discovery=null;
  async function request(method,path,payload,timeout){
    if(typeof GM_xmlhttpRequest!=='function')throw new Error('Tampermonkey permission for SAPGUIcontrol is unavailable. Reinstall or update this userscript.');
    if(/^https?:\/\//i.test(path))return bridgeHttpRequest(method,path,payload,timeout);
    // Recheck identity before operations, including after an EXE restart. Concurrent callers share probes.
    if(!discovery)discovery=(async()=>{
      for(const url of origin?[origin,...urls.filter(url=>url!==origin)]:urls){
        try{const service=await bridgeHttpRequest('GET',url+'/',undefined,1500);if(service?.ok===true&&service.service==='SAPGUIcontrol'){origin=url;return url}}catch(_){/* try the next candidate */}
      }
      origin='';throw bridgeTransportError();
    })().finally(()=>{discovery=null});
    const url=await discovery;
    // Do not replay a failed operation: it may already have started in the EXE.
    return bridgeHttpRequest(method,new URL(path,url).href,payload,timeout);
  }
  function files(operation,options,timeout){return request('POST','/files',{operation,...options},timeout||600000)}
  return {request,files,get origin(){return origin}};
}
// Only SAP Toolbox initializes the local EXE connection, UI and task state.
if(currentPath===BRIDGE_PATH){
const BRIDGE_URLS=[25972,25973,25974,8765].map(port=>'http://127.0.0.1:'+port);
const bridgeClient=createSapGuiControlClient(BRIDGE_URLS),{request:bridgeRequest,files:bridgeFileRequest}=bridgeClient;
const SAPGUICONTROL_DOWNLOAD_URL='https://lamresearch-my.sharepoint.com/:u:/p/kyra_long/IQAkrJwix3nKTKijNazpuktkAUXL46RCibysVC7XkXJmYcQ?e=UUDyWl'; const SAPGUICONTROL_VERSION_URL='https://synalocey.github.io/ToolkitKyra/SAPGUIcontrol.version.json';
const bobjCards={fcid:{id:'bobj',label:'FID Database',endpoint:'/excel'},ctl:{id:'ctlBobj',label:'CTL Report',endpoint:'/excel/ctl'}};
const BRIDGE_CSS=chosenBaseCss+(`
:root{--bg:#f6f8fe;--panel:#fff;--ink:#2e486c;--muted:#526b8d;--line:#d6e0ef;--blue:#215fbb;--teal:#0f766e;--green:#16843f;--amber:#b77912;--red:#b91c1c;--shadow:0 2px 12px #253c6410;
--font-sans:"Segoe UI Variable Text","Segoe UI","Microsoft YaHei UI","Microsoft YaHei",Arial,sans-serif}*{box-sizing:border-box}[hidden]{display:none!important}body{margin:0;background:var(--bg);color
:var(--ink);font-family:var(--font-sans);font-weight:400;line-height:1.4}b,h1,h2,strong{font-weight:600}button,input,select,textarea{font:inherit}button:focus-visible,input:focus-visible,
select:focus-visible,textarea:focus-visible,summary:focus-visible{outline:3px solid #8ab4f2;outline-offset:2px}.toolbox-app{--side:clamp(330px,27vw,400px);font-size:13px}
.toolbox-app .layout>.side{overflow:visible;max-height:none;border-radius:11px;z-index:2}.toolbox-app .main{container-type:inline-size}.toolbox-app .section-head{margin-bottom:8px}.toolbox-app .section-head h2{font-size:16px}
.toolbox-app :is([data-tool-tab=zbom],[data-tool=zbom]){--tool-ink:#2458a6;--tool-tint:#eaf1ff;--tool-line:#bed0ef}
.toolbox-app :is([data-tool-tab=zbomCompare],[data-tool=zbomCompare]){--tool-ink:#7950a1;--tool-tint:#f3edfa;--tool-line:#d2bee8}
.toolbox-app :is([data-tool-tab=partChanges],[data-tool=partChanges]){--tool-ink:#475569;--tool-tint:#edf1f7;--tool-line:#cbd5e1}
.toolbox-app :is([data-tool-tab=createQuote],[data-tool=createQuote]){--tool-ink:#176973;--tool-tint:#e8f6f7;--tool-line:#b0d7db}
.toolbox-app .workspace-nav [data-tool-tab]{background:var(--tool-tint);color:var(--tool-ink);box-shadow:inset 0 0 0 1px var(--tool-line)}
.toolbox-app .workspace-nav [data-tool-tab]:hover{box-shadow:inset 0 0 0 1px var(--tool-ink)}
.toolbox-app .workspace-nav [data-tool-tab][aria-selected=true]{box-shadow:inset 0 0 0 2px var(--tool-ink)}
.toolbox-sidebar-section{border-top:1px solid var(--line);margin-top:12px;padding-top:12px}
.toolbox-task-list,.toolbox-task-details{display:grid;grid-template-columns:minmax(0,1fr)}
.toolbox-task-list{gap:9px}.toolbox-task-row{display:grid;grid-template-columns:minmax(0,1fr) auto;align-items:start;gap:7px 10px;border:1px solid var(--line);border-radius:9px;padding:11px;background:#fff}
#toolboxTaskMore{display:block;margin:12px auto 0}
.toolbox-task-open{border:0;background:none;padding:0;text-align:left;cursor:pointer;min-width:0}
.toolbox-task-open strong,#toolboxTaskTitle{color:var(--tool-ink,var(--ink))}.toolbox-task-open strong{display:block;font-size:14px;line-height:1.35}.toolbox-task-open:hover strong{text-decoration:underline}
.toolbox-task-date{display:block;color:var(--muted);font-size:12px;font-variant-numeric:tabular-nums}.toolbox-task-open .toolbox-task-date{margin-top:2px}
.toolbox-task-summary{display:grid;justify-items:end;gap:4px;font-size:12px;text-align:right}.toolbox-task-count{margin-inline-start:3px;color:var(--muted);font-variant-numeric:tabular-nums;white-space:nowrap}
.toolbox-task-summary .bridge-btn.small{padding:3px 7px;white-space:nowrap}
.toolbox-task-summary:has(.toolbox-task-delete){grid-template-columns:auto auto;align-items:center}
.toolbox-task-summary:has(.toolbox-task-delete) .bridge-btn{grid-column:1/-1}
.toolbox-task-delete{grid-area:1/2;width:28px;height:28px;padding:0;border:0;border-radius:6px;background:transparent;font-size:14px;cursor:pointer}
.toolbox-task-delete:hover:not(:disabled){background:#fee2e2}.toolbox-task-delete:disabled{opacity:.5;cursor:wait}
.toolbox-task-state::before{content:"";display:inline-block;width:7px;height:7px;border-radius:50%;background:#94a3b8;margin-right:5px}
.toolbox-task-row:is([data-state=Running],[data-state=Comparing]) .toolbox-task-state::before{background:var(--blue);animation:task-pulse 1.4s ease-in-out infinite}.toolbox-task-row[data-state=Completed] .toolbox-task-state::before{background:var(--green)}
.toolbox-task-row[data-state=Failed] .toolbox-task-state::before,.toolbox-task-row[data-state=Interrupted] .toolbox-task-state::before{background:var(--red)}.toolbox-task-row:is([data-state="Needs attention"],[data-state="Partial complete"]) .toolbox-task-state::before{background:var(--amber)}
.toolbox-task-progress{grid-column:1/-1;min-width:0;display:grid;gap:4px}.toolbox-task-progress .progress-track{height:5px}
.toolbox-task-phase{grid-column:1/-1;min-width:0;display:flex;gap:8px;justify-content:space-between;color:var(--muted);font-size:11px}.toolbox-task-phase span:first-child{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.toolbox-task-phase span+span{flex:none;font-variant-numeric:tabular-nums}
.toolbox-task-details{grid-auto-rows:max-content;gap:14px;padding-top:16px;overflow:hidden auto;overflow-wrap:anywhere}.toolbox-task-details .bridge-actions{margin:0}
#toolboxTaskResults{margin:0;overflow-x:hidden}#toolboxTaskResults th{white-space:normal}
#toolboxTaskDialog[data-tool=partChanges] #toolboxTaskResults{overflow-x:auto;max-height:360px}
.part-changes-system-cell{width:200px;min-width:190px;--chosen-choice-gap:3px;--chosen-choice-padding-block:2px;--chosen-border-color:#cbd5e1;--chosen-border-radius:7px}
.part-changes-system-cell .chosen-container-multi .chosen-choices{padding:1px 3px;color:var(--ink);font:11px/16px var(--font-sans)}
.part-changes-system-cell .chosen-container-multi .chosen-choices li.search-choice{margin:2px 3px 2px 0}
.part-changes-system-cell .chosen-container-multi .chosen-choices li.search-field input.chosen-search-input{padding:3px 0;font-size:11px}
.part-changes-system-cell .chosen-container-active .chosen-choices{border-color:#8ab4f2}.part-changes-system-cell .chosen-results{max-height:200px}
#partChangesPreviewBody .compare-file-name{min-width:380px}
#partChangesPreviewBody .chosen-with-drop,#toolboxTaskResults .chosen-with-drop{z-index:2}
#toolboxTaskStatus{display:grid;grid-template-columns:1fr auto;gap:8px 16px;font-size:13px}.toolbox-task-current{grid-column:1/-1;grid-row:2;color:var(--ink);font-weight:600}.toolbox-task-timing{color:var(--muted);font-variant-numeric:tabular-nums}
#toolboxTaskStatus .toolbox-task-state{color:var(--muted)}#toolboxTaskStatus:is([data-state=Running],[data-state=Comparing]) .toolbox-task-state::before{background:var(--blue)}#toolboxTaskStatus[data-state=Completed] .toolbox-task-state::before{background:var(--green)}
#toolboxTaskStatus[data-state=Failed] .toolbox-task-state::before,#toolboxTaskStatus[data-state=Interrupted] .toolbox-task-state::before{background:var(--red)}#toolboxTaskStatus:is([data-state="Needs attention"],[data-state="Partial complete"]) .toolbox-task-state::before{background:var(--amber)}
#toolboxTaskPrompt{font-size:13px;color:var(--red);padding:10px 12px;border:1px solid #e9c8c8;border-radius:7px;background:#fff7f7;white-space:pre-wrap}#toolboxTaskPrompt:empty{display:none}
@keyframes task-pulse{50%{opacity:.3}}@media(prefers-reduced-motion:reduce){.toolbox-task-state::before{animation:none!important}}
.toolbox-app .side .bridge-actions{margin-top:6px}
.zbom-grid,.create-quote-grid{display:grid;gap:16px}.zbom-grid{grid-template-columns:minmax(0,1.3fr) minmax(300px,.9fr)}` +
  `.create-quote-grid{grid-template-columns:minmax(0,1fr) minmax(0,1.3fr) minmax(0,1.2fr)}.bridge-card{min-width:0;
background:var(--panel);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow);padding:17px}.bridge-card.wide{grid-column:1/-1}.bridge-card h2{margin:0;font-size:17px;line-height:1.3
}.bridge-card p{margin:4px 0 0;color:var(--muted);font-size:11px;line-height:1.5}.bridge-card-head{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;margin-bottom:14px}
.bridge-card-tools{display:flex;align-items:center;gap:7px}.bridge-btn{border:1px solid transparent;border-radius:7px;background:var(--teal);color:#fff;font-size:13px;font-weight:600;padding:9px
13px;cursor:pointer;transition:background .12s ease}.bridge-btn:hover:not(:disabled){filter:brightness(1.05)}.bridge-btn.blue{background:var(--blue)}
.bridge-btn.light{border-color:#c5d2e5;background:#fff;color:var(--blue)}.bridge-btn.light:hover:not(:disabled){background:#e7edf8}.bridge-btn:disabled{cursor:not-allowed;opacity:.5}.bridge-btn.small{padding:6px
9px;font-size:12px}.bridge-badge{display:inline-flex;align-items:center;gap:7px;border-radius:999px;background:#f1f5f9;color:#475569;padding:6px 10px;font-size:11px;font-weight:600;white-space:nowrap}
.bridge-badge::before{content:"";width:8px;height:8px;border-radius:50%;background:#94a3b8}.bridge-badge.ok{background:#dcfce7;color:#166534}.bridge-badge.ok::before{background:var(--green)}
.bridge-badge.warn{background:#fef3c7;color:#92400e}.bridge-badge.warn::before{background:var(--amber)}.bridge-badge.error{background:#fee2e2;color:#991b1b}.bridge-badge.error::before{background:
var(--red)}.bridge-download{text-decoration:none}.toolbox-app .fc-source h3{margin:0}.toolbox-app .fc-source-file{font-variant-numeric:tabular-nums}
.bobj-refresh-state{display:block;margin-top:4px}.bobj-refresh-state.running{color:#92400e}.bobj-refresh-state.error{color:#b91c1c}
.bridge-field{display:grid;gap:6px;margin-top:11px}.bridge-field label{color:#475569;font-size:12px;` +
  `font-weight:600;
}.bridge-field input,.bridge-field select,.bridge-field textarea{display:block;width:100%;border:1px solid #c2d0e4;border-radius:7px;background:#fff;color
:var(--ink);padding:10px 11px;font-size:13px}.bridge-field textarea{min-height:88px;resize:vertical;font:12px/1.55 Consolas,"Courier New",monospace;tab-size:4}.bridge-field textarea.quotation-input{
white-space:pre;overflow:auto;overflow-wrap:normal;word-break:normal}.bridge-field input:disabled,.bridge-field select:disabled,.bridge-field textarea:disabled{background:#f1f5f9;color:#94a3b8}
.bridge-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:12px}.bridge-note{color:var(--muted);font-size:12px;line-height:1.5}.token-list{display:flex;flex-wrap:wrap;gap:6px;margin-top:
8px}.token-chip{border:1px solid #bfdbfe;border-radius:999px;background:#eff6ff;color:#1d4ed8;padding:4px 8px;font-size:10px;cursor:pointer}.folder-line{display:flex;align-items:center;gap:9px;
flex-wrap:wrap;margin-top:12px}.folder-name{min-width:0;color:#475569;font-size:11px;overflow-wrap:anywhere}.toolbox-modal-card{display:grid;grid-template-rows:auto minmax(0,1fr);width:min(1100px,96vw);max-width:96vw;max-height:92vh;overflow:hidden;border:1px solid #b7c7e0;border-radius:12px;background:#f6f8fe;color:var(--ink);
box-shadow:0 16px 70px #1e355550;padding:20px}.toolbox-modal-card:not([open]){display:none}.toolbox-modal-card::backdrop{background:#152e505c}.toolbox-modal-head{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;padding-bottom:14px;border-bottom:1px solid var(--line)}
.toolbox-modal-close{flex:none;margin:-8px -8px 0 0;width:36px;height:36px;padding:0;border:0;border-radius:7px;background:transparent;color:var(--muted);font-size:28px;line-height:1;cursor:pointer}.toolbox-modal-close:hover{background:#e7edf8}
.toolbox-modal-title{display:grid;gap:3px}.toolbox-modal-title h2{margin:0;font-size:17px}.toolbox-app:has(> .toolbox-modal-card[open]){overflow:hidden}
.toolbox-modal-card.toolbox-task-history{width:min(560px,96vw)}.toolbox-task-history .toolbox-task-list{align-content:start;overflow:hidden auto;padding:12px 4px 4px}
.preview-toolbar{display:flex;align-items:center;justify-content:
space-between;gap:10px;flex-wrap:wrap}.preview-count{color:#475569;font-size:11px}.zbom-table-wrap{margin-top:12px;border:1px solid #dbe5ee;border-radius:12px}
.zbom-table{width:100%;border-collapse:separate;border-spacing:0;font-size:12px;background:#fff;border-radius:11px}.zbom-table th{position:sticky;top:0;z-index:1;background:#edf2fb;color:#425c7f;text-align:left;font-size:12px;font-weight:600;
white-space:nowrap}.zbom-table td,.zbom-table th{padding:8px 9px;border-right:1px solid #e5edf4;border-bottom:1px solid #e5edf4;vertical-align:top}.zbom-table td:last-child,.zbom-table th:last-child{
border-right:0}.zbom-table tr:last-child td{border-bottom:0}.zbom-table .missing{color:#94a3b8}.row-ok{color:#166534}.row-warning{color:#92400e}.row-error{color:#b91c1c}.planned-name{min-width:270px;max-width:440px;
overflow-wrap:anywhere}.compare-template-control{display:grid;grid-template-columns:minmax(0,1fr) 64px;gap:8px;align-items:center}.compare-template-separator{text-align:center;font-weight:700}.compare-table td,.compare-table ` +
  `th{padding:4px 7px;vertical-align:middle}
.zbom-file-edit,.zbom-file-save,.zbom-file-reset{border:0;background:transparent;padding:0 4px;cursor:pointer;font:inherit;white-space:nowrap}.zbom-file-edit::before{content:"✏️"}.zbom-file-save::before{content:"✓"}.zbom-file-save{color:#15803d;font-size:15px;font-weight:700}.zbom-file-reset::before{content:"↩️"}.zbom-file-edit:disabled,.zbom-file-save:disabled,.zbom-file-reset:disabled{cursor:not-allowed;opacity:.45}
.zbom-file-editor{display:flex;align-items:center;gap:5px}.zbom-file-editor .zbom-file-input{flex:1;min-width:0}
.zbom-file-input{width:100%;min-width:220px;border:1px solid #94a3b8;border-radius:6px;padding:4px 6px;font:inherit}.zbom-file-input:invalid{border-color:var(--red);background:#fff1f2}
.zbom-quote-select{width:100%;min-width:104px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:var(--ink);padding:3px 6px;font-size:11px;font-variant-numeric:tabular-nums}
.zbom-quote-fallback{position:relative;display:block}.zbom-quote-fallback select{color:transparent}.zbom-quote-fallback select option{color:var(--ink)}
.zbom-quote-fallback>span{position:absolute;inset-inline:7px 20px;top:50%;transform:translateY(-50%);overflow:hidden;white-space:nowrap;pointer-events:none}
@supports(appearance:base-select){
  :is(.zbom-quote-select,.compare-group-select),:is(.zbom-quote-select,.compare-group-select)::picker(select){appearance:base-select}
  select.compare-group-select{display:flex}
  :is(.zbom-quote-select,.compare-group-select)::picker(select){min-width:236px;max-height:320px;overflow:auto;border:1px solid #cbd5e1;border-radius:8px;background:#fff;color:var(--ink);box-shadow:0 6px 18px #23395626}
  :is(.zbom-quote-select,.compare-group-select) option{display:grid;grid-template-columns:12px 1fr;gap:1px 6px;padding:6px 8px}
  :is(.zbom-quote-select,.compare-group-select) option::checkmark{grid-column:1;grid-row:1}
  .zbom-quote-select option>strong,.zbom-quote-select option>small{grid-column:2}
  .zbom-quote-select option>small{font-size:11px;line-height:1.5;color:#475569;white-space:pre-line}
}
.compare-fid{min-width:78px}.compare-customer{width:110px;min-width:100px;max-width:145px;overflow-wrap:anywhere}.compare-system-description{width:155px;min-width:140px;max-width:230px;overflow-wrap:anywhere}
.compare-group-cell{width:92px;min-width:92px}.compare-group-select{display:block;width:100%;min-width:0;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:var(--ink);padding:3px 6px;font-size:11px}` +
  `.compare-file-name{min-width:460px;max-width:900px;
overflow-wrap:anywhere}.compare-file-link{border:0;background:0 0;color:#0369a1;padding:0;text-align:left;font:inherit;text-decoration:underline;text-underline-offset:2px;cursor:pointer;overflow-wrap:anywhere}` +
  `.compare-file-link:hover{color:#0f766e}.compare-file-link:disabled{color:#94a3b8;cursor:wait}.compare-output-control{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;align-items:center}` +
  `
.compare-keep-temp{display:inline-flex;align-items:center;gap:6px;min-height:32px;padding:0;border:0;background:none;color:#475569;font-size:12px;font-weight:600;white-space:nowrap;cursor:pointer}
.compare-keep-temp:disabled{opacity:.5;cursor:not-allowed}
.compare-keep-temp-track{display:flex;align-items:center;width:32px;height:18px;padding:2px;border:1px solid currentColor;border-radius:999px;color:#64748b;background:currentColor}
.compare-keep-temp-track::before{content:"";width:12px;height:12px;border-radius:50%;background:#fff}
.compare-keep-temp[aria-checked=true] .compare-keep-temp-track{justify-content:flex-end;color:var(--teal)}
@media(forced-colors:active){
  .compare-keep-temp-track{background:ButtonFace}
  .compare-keep-temp-track::before{background:ButtonText}
}
.create-quote-reference-title{white-space:nowrap}.create-quote-number{max-width:112px;` +
  `font-variant-numeric:tabular-nums}.create-quote-setting-row{grid-template-columns:max-content minmax(0,1fr);align-items:center;column-gap:10px}.create-quote-days-control{display:grid;grid-template-columns:auto 62px auto;` +
  `align-items:center;justify-content:start;gap:7px}.create-quote-days-control input{text-align:center;font-variant-numeric:tabular-nums}.create-quote-unit{color:#475569;font-size:11px;font-weight:600}` +
  `.create-quote-suffix{max-width:76px}.create-quote-crd-input{display:block;width:150px;max-width:100%;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:var(--ink);padding:5px 7px;font-size:11px}` +
  `.create-quote-crd-input.invalid{border-color:var(--red);background:#fff1f2}.create-quote-version-note{font-weight:400;white-space:nowrap}.create-quote-table{table-layout:fixed}` +
  `.create-quote-table th:first-child,.create-quote-table td:first-child{width:58px}.create-quote-table th:not(:first-child),.create-quote-table td:not(:first-child){width:calc((100% - 58px)/3)}.compare-status{width:74px;` +
  `min-width:74px;white-space:nowrap}.progress-track{height:8px;border-radius:999px;background:#e2e8f0;overflow:hidden}
.progress-track::after{content:"";display:block;width:var(--task-progress,0%);height:100%;border-radius:inherit;background:linear-gradient(90deg,var(--blue),var(--teal))}
.progress-track[data-state=running]:not([aria-valuenow])::after{width:32%;animation:sap-progress 1.5s ease-in-out infinite}.progress-track[data-state=ok]::after{background:var(--green)}
.progress-track[data-state=warning]{background:#fef3c7}.progress-track[data-state=warning]::after{background:var(--amber)}
.progress-track[data-state=error]{background:#fee2e2}.progress-track[data-state=error]::after{background:var(--red)}
@keyframes sap-progress{from{transform:translateX(-100%)}to{transform:translateX(313%)}}@media(prefers-reduced-motion:reduce){.progress-track::after{animation:none!important}}
@container(max-width:820px){.create-quote-grid{grid-template-columns:minmax(0,1fr)}}
@container(max-width:650px){.zbom-grid{grid-template-columns:minmax(0,1fr)}}@media(max-width:760px){.planned-name{min-width:220px}.toolbox-modal-card{padding:16px;max-height:94vh}#toolboxTaskStatus{grid-template-columns:1fr}.toolbox-task-current{grid-row:auto}}`);
function toolboxTokenButtons(attribute,fields,action=''){
  return fields.map(field=>`<button class="token-chip" type="button" ${attribute}="{${field}}"${action}>{${field}}</button>`).join('\n          ');
}
function toolboxStartActions(tool,action=''){
  return `<div class="bridge-actions">
          <button id="${tool}Start" class="bridge-btn blue" type="button"${action} disabled>Start</button>
          <span id="${tool}ProgressText" class="bridge-note" role="status" hidden></span>
        </div>`;
}
function toolboxResetButton(id,action=''){
  return `<button id="${id}" class="bridge-btn light small" type="button"${action}>Reset</button>`;
}
const BRIDGE_HTML=(`
${workspaceHeader('<span aria-hidden="true">🧰</span> SAP Toolbox','bridgeRefresh','data-bridge-action')}
<div class="layout">
<aside class="side bridge-card" aria-label="SAP Toolbox sidebar">
  <section>
    <div class="section-head"><h2>SAPGUIcontrol</h2><span id="bridgeBadge" class="bridge-badge" role="status">Checking</span></div>
    <div class="bridge-actions"><strong id="bridgeService" hidden></strong>
      <a id="bridgeDownload" class="btn light small bridge-download" href="${SAPGUICONTROL_DOWNLOAD_URL}" download="SAPGUIcontrol.zip" target="_blank" rel="noopener noreferrer" hidden>Download</a>
    </div>
    <div id="bridgeStatus" class="bridge-note" role="status"></div>
  </section>
  <section class="toolbox-sidebar-section"><div class="section-head"><h2>Data Sources</h2></div>
    ${Object.entries(bobjCards).map(([key,card],index)=>bobjSourceCard(key,card,index+1)).join('')}
  </section>
  <section class="toolbox-sidebar-section"><div class="section-head"><h2>Tasks</h2></div>
    <div id="toolboxTaskList" class="toolbox-task-list" aria-live="polite"><span class="bridge-note">Connecting to SAPGUIcontrol…</span></div>
    <button id="toolboxTaskMore" class="bridge-btn light small" type="button" aria-haspopup="dialog" aria-controls="toolboxTaskHistory" hidden>View all tasks</button>
  </section>
</aside>
<main class="main">
  <nav class="workspace-nav" role="tablist" aria-label="SAP Toolbox functions">
    <button id="tabZbom" type="button" role="tab" aria-selected="true" aria-controls="panelZbom" data-tool-tab="zbom">ZBOM PDF</button>
    <button id="tabZbomCompare" type="button" role="tab" aria-selected="false" aria-controls="panelZbomCompare" data-tool-tab="zbomCompare">ZBOM Compare</button>
    <button id="tabPartChanges" type="button" role="tab" aria-selected="false" aria-controls="panelPartChanges" data-tool-tab="partChanges">${BOM_COMPARE_LABEL}</button>
    <button id="tabCreateQuote" type="button" role="tab" aria-selected="false" aria-controls="panelCreateQuote" data-tool-tab="createQuote">Quote Create</button>
  </nav>

  <section id="panelZbom" class="tool-panel" role="tabpanel" aria-labelledby="tabZbom" data-tool-panel="zbom">
    <div class="zbom-grid">
      <section class="bridge-card">
        <div class="bridge-card-head">
          <div><h2>1. Quotation Data</h2></div>
          <div class="bridge-card-tools">
            ${toolboxResetButton('zbomResetInput')}
          </div>
        </div>
        <div class="bridge-field"><textarea id="zbomInput" class="quotation-input" rows="4" wrap="off" spellcheck="false" aria-label="Quotation data" placeholder="FID    System Description    Quote&#10;225822-4    ` +
  `WTSC_MAX/HALT_MAXx2    20736497"></textarea></div>
        <div class="bridge-note">One entry per line. Mix 6-digit FIDs, 10-digit CTL IDs and 8-digit Quotes. Excel columns supported.</div>
      </section>

      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2>2. Filename</h2></div>${toolboxResetButton('zbomResetTemplate')}</div>
        <div class="bridge-field"><input id="zbomTemplate" type="text" maxlength="240" value="{FID}{FIDver}_{System Description}_{Quote}_{Today}" autocomplete="off"></div>
        <div class="token-list" aria-label="Available file-name tokens">
          ${toolboxTokenButtons('data-zbom-token',['FID','FIDver','System Description','Quote','Today'])}
        </div>
        <div class="folder-line"><span class="folder-name">Default output folder: C:\\PDFFILES</span></div>
      </section>

      <section class="bridge-card wide">
        <div class="preview-toolbar">
          <div><h2>3. Preview</h2></div>
          <div class="bridge-card-tools"><span id="zbomCount" class="preview-count">0 rows</span><button id="zbomSapLookup" class="btn light small" type="button" data-sap-lookup="zbom" disabled>Lookup in SAP</button></div>
        </div>
        <div class="zbom-table-wrap">
          <table class="zbom-table" aria-label="Detected ZBOM PDF input">
            <thead></thead>
            <tbody id="zbomPreviewBody"></tbody>
          </table>
        </div>
      </section>

      <section class="bridge-card wide">
        <div class="bridge-card-head"><div><h2>4. Export PDF Files</h2></div></div>
        ${toolboxStartActions('zbom')}
      </section>
    </div>
  </section>

  <section id="panelZbomCompare" class="tool-panel" role="tabpanel" aria-labelledby="tabZbomCompare" data-tool-panel="zbomCompare" hidden>
    <div class="zbom-grid">
      <section class="bridge-card">
        <div class="bridge-card-head">
          <div><h2>1. Quotation Data</h2></div>
          <div class="bridge-card-tools">
            ${toolboxResetButton('zbomCompareResetInput')}
          </div>
        </div>
        <div class="bridge-field"><textarea id="zbomCompareInput" class="quotation-input" rows="6" wrap="off" spellcheck="false" aria-label="ZBOM Compare quotation data" placeholder="Customer    FID    System Description    ` +
  `Quotation    Group&#10;Customer A    225822    WTSC_MAX/HALT_MAXx2    21314937    G1"></textarea></div>
        <div class="bridge-note">One entry per line. Mix 6-digit FIDs, 10-digit CTL IDs and 8-digit Quotes. Excel columns supported.</div>
      </section>

      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2>2. Filename</h2></div>${toolboxResetButton('zbomCompareResetTemplate')}</div>
        <div class="bridge-field">
          <div class="compare-template-control">
            <input id="zbomCompareTemplate" type="text" maxlength="240" value="{Customer}_{FID}{FIDver}_{System Description}" autocomplete="off" aria-label="ZBOM Compare filename template">
            <input id="zbomCompareSeparator" class="compare-template-separator" type="text" maxlength="24" value="vs" autocomplete="off" spellcheck="false" aria-label="Text between compared file names" title="Text between ` +
  `compared file names">
          </div>
        </div>
        <div class="token-list" aria-label="Available ZBOM Compare file-name tokens">
          ${toolboxTokenButtons('data-zbom-compare-token',['Customer','FID','FIDver','System Description','Quote'])}
        </div>
        <div class="bridge-note">The short field joins the file names of compared quotations.</div>
        <div class="bridge-field">
          <label for="zbomCompareOutputFolder">Output Folder</label>
          <div class="compare-output-control">
            <input id="zbomCompareOutputFolder" type="text" maxlength="1000" value="C:\\PDFFILES\\Compare" autocomplete="off" spellcheck="false" placeholder="C:\\PDFFILES\\Compare">
            <button id="zbomCompareKeepTemp" class="compare-keep-temp" type="button" role="switch" aria-checked="false" title="Keep the temp folder after Compare succeeds">
              <span class="compare-keep-temp-track" aria-hidden="true"></span>Keep temp
            </button>
          </div>
        </div>
      </section>

      <section class="bridge-card wide">
        <div class="preview-toolbar">
          <div><h2>3. Preview</h2></div>
          <div class="bridge-card-tools"><span id="zbomCompareCount" class="preview-count">0 quotations</span><button id="zbomCompareSapLookup" class="btn light small" type="button" data-sap-lookup="zbomCompare" disabled>Lookup in SAP</button></div>
        </div>
        <div class="zbom-table-wrap">
          <table class="zbom-table compare-table" aria-label="Detected ZBOM Compare input">
            <thead></thead>
            <tbody id="zbomComparePreviewBody"></tbody>
          </table>
        </div>
      </section>

      <section class="bridge-card wide">
        <div class="bridge-card-head"><div><h2>4. Compare</h2></div></div>
        ${toolboxStartActions('zbomCompare',' data-bridge-action')}
      </section>
    </div>
  </section>

  <section id="panelPartChanges" class="tool-panel" role="tabpanel" aria-labelledby="tabPartChanges" data-tool-panel="partChanges" hidden>
    <div class="zbom-grid">
      <section class="bridge-card">
        <div class="bridge-card-head"><h2>1. Quotation Data</h2>${toolboxResetButton('partChangesReset',' data-bridge-action')}</div>
        <div class="bridge-field"><textarea id="partChangesInput" class="quotation-input" rows="6" wrap="off" spellcheck="false" aria-label="${BOM_COMPARE_LABEL} quotation data" placeholder="Identifier    Group&#10;266845    G01&#10;21291911    G01&#10;P4000    G01"></textarea></div>
        <div class="bridge-note">One entry per line. Mix 6-digit FIDs, 8-digit Quotes and 5-character SO Numbers (e.g. P1234). Excel columns supported.</div>
      </section>
      <section class="bridge-card">
        <div class="bridge-card-head"><h2>2. Filename</h2>${toolboxResetButton('partChangesResetTemplate',' data-bridge-action')}</div>
        <div class="bridge-field">
          <div class="compare-template-control">
            <input id="partChangesTemplate" type="text" maxlength="240" autocomplete="off" aria-label="${BOM_COMPARE_LABEL} filename template">
            <input id="partChangesSeparator" class="compare-template-separator" type="text" maxlength="24" autocomplete="off" spellcheck="false" aria-label="${BOM_COMPARE_LABEL} filename separator">
          </div>
        </div>
        <div class="token-list" aria-label="Available ${BOM_COMPARE_LABEL} filename tokens">
          ${toolboxTokenButtons('data-part-changes-token',['Customer','FID','FIDver','System Description','Quote','SO'],' data-bridge-action')}
        </div>
        <div class="bridge-field">
          <label for="partChangesOutputFolder">Comparison Output Folder</label>
          <input id="partChangesOutputFolder" type="text" maxlength="1000" aria-label="${BOM_COMPARE_LABEL} output folder" value="C:\\PDFFILES\\${BOM_COMPARE_LABEL}" autocomplete="off" spellcheck="false">
        </div>
      </section>
      <section class="bridge-card wide">
        <div class="preview-toolbar"><h2>3. Preview</h2><div class="bridge-card-tools"><span id="partChangesCount" class="preview-count"></span>
          <button id="partChangesLookup" class="btn light small" type="button" disabled>Lookup in SAP</button></div></div>
        <div class="zbom-table-wrap"><table class="zbom-table" aria-label="${BOM_COMPARE_LABEL} preview">
          <thead><tr><th>FID</th><th>Sales Order / Quotation</th><th>FID System</th><th>Compare Group</th><th>File Name</th><th>Status</th></tr></thead><tbody id="partChangesPreviewBody"></tbody>
        </table></div>
      </section>
      <section class="bridge-card wide">
        <div class="bridge-card-head"><h2>4. Compare</h2></div>
        ${toolboxStartActions('partChanges',' data-bridge-action')}
        <div id="partChangesMyLamNote" class="bridge-note">After SAP exports the BOMs, SAPGUIcontrol opens MyLam in your default browser. You can close this tab after submission.</div>
      </section>
    </div>
  </section>

  <section id="panelCreateQuote" class="tool-panel" role="tabpanel" aria-labelledby="tabCreateQuote" data-tool-panel="createQuote" hidden>
    <div class="create-quote-grid">
      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2 class="create-quote-reference-title">1. Reference Quotation</h2></div></div>
        <div class="bridge-field"><input id="createQuoteQuotation" class="create-quote-number" type="text" inputmode="numeric" pattern="[0-9]*" maxlength="10" autocomplete="off" spellcheck="false" aria-label="Create Quote ` +
  `quotation number" placeholder="Quotation No."></div>
      </section>

      <section class="bridge-card">
        <div class="bridge-card-head">
          <div><h2>2. FID</h2></div>
          <div class="bridge-card-tools">
            ${toolboxResetButton('createQuoteResetInput',' data-bridge-action')}
          </div>
        </div>
        <div class="bridge-field"><textarea id="createQuoteInput" class="quotation-input" rows="4" wrap="off" spellcheck="false" aria-label="Create Quote FID and date data" placeholder="FID    CRD    ValidTo&#10;264072    ` +
  `08/01/2026    09/15/2026"></textarea></div>
      </section>

      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2>3. Settings</h2></div>${toolboxResetButton('createQuoteResetSettings',' data-bridge-action')}</div>
        <div class="bridge-field create-quote-setting-row">
          <label for="createQuoteDefaultValidTo">Default ValidTo</label>
          <div class="create-quote-days-control">
            <span class="create-quote-unit">+</span>
            <input id="createQuoteDefaultValidTo" type="text" inputmode="numeric" pattern="[0-9]*" maxlength="4" value="30" autocomplete="off" spellcheck="false" aria-label="Default ValidTo days">
            <span class="create-quote-unit">Days</span>
          </div>
        </div>
        <div class="bridge-field create-quote-setting-row">
          <label for="createQuoteFidSuffix">New FID suffix</label>
          <input id="createQuoteFidSuffix" class="create-quote-suffix" type="text" maxlength="6" value="-01" autocomplete="off" spellcheck="false">
        </div>
      </section>

      <section class="bridge-card wide">
        <div class="preview-toolbar">
          <div><h2>4. Preview</h2></div>
          <span id="createQuoteCount" class="preview-count">0 FIDs</span>
        </div>
        <div class="zbom-table-wrap">
          <table class="zbom-table create-quote-table" aria-label="Detected Create Quote input">
            <thead><tr><th>Item</th><th>CRD</th><th>Version <span class="bridge-note create-quote-version-note">(Incremented if same as Ref. Quotation)</span></th><th>Valid to</th></tr></thead>
            <tbody id="createQuotePreviewBody"><tr><td colspan="4" class="missing">No preview yet.</td></tr></tbody>
          </table>
        </div>
      </section>

      <section class="bridge-card wide">
        <div class="bridge-card-head"><div><h2>5. Create Quote</h2></div></div>
        ${toolboxStartActions('createQuote')}
      </section>
    </div>
  </section>

</main></div>
<dialog id="toolboxTaskHistory" class="toolbox-modal-card toolbox-task-history" aria-labelledby="toolboxTaskHistoryTitle">
  <div class="toolbox-modal-head"><div class="toolbox-modal-title"><h2 id="toolboxTaskHistoryTitle">All tasks</h2></div><button id="toolboxTaskHistoryClose" class="toolbox-modal-close" type="button" aria-label="Close all tasks" autofocus>×</button></div>
  <div id="toolboxTaskHistoryList" class="toolbox-task-list"></div>
</dialog>
<dialog id="toolboxTaskDialog" class="toolbox-modal-card" aria-labelledby="toolboxTaskTitle" aria-describedby="toolboxTaskCreated">
    <div class="toolbox-modal-head"><div class="toolbox-modal-title"><h2 id="toolboxTaskTitle">Task</h2><time id="toolboxTaskCreated" class="toolbox-task-date" hidden></time></div><button id="toolboxTaskClose" class="toolbox-modal-close" type="button" aria-label="Close task details">×</button></div>
    <div class="toolbox-task-details"><div id="toolboxTaskStatus" role="status"></div><div id="toolboxTaskAnimation" class="progress-track" role="progressbar" hidden></div>
      <div id="toolboxTaskActions" class="bridge-actions"></div><div id="toolboxTaskPrompt" class="bridge-note" role="alert"></div>
      <div id="toolboxTaskResults" class="zbom-table-wrap"></div>
    </div>
</dialog>`);
const ZBOM_PDF_FOLDER='C:\\PDFFILES'; const ZBOM_DEFAULT_TEMPLATE='{FID}{FIDver}_{System Description}_{Quote}_{Today}'; const ZBOM_STORAGE_KEYS={input:'nsrSapToolbox.zbomInput.v1',template:'nsrSapToolbox.zbomTemplate.v1',fileNames:'nsrSapToolbox.zbomFileNames.v1'};
const ZBOM_COMPARE_DEFAULT_GROUP='G01'; const ZBOM_COMPARE_DEFAULT_FOLDER='C:\\PDFFILES\\Compare';
const ZBOM_COMPARE_STORAGE_KEYS={input:'nsrSapToolbox.zbomCompareInput.v1',template:'nsrSapToolbox.zbomCompareTemplate.v1',separator:'nsrSapToolbox.zbomCompareSeparator.v1',folder:'nsrSapToolbox.zbomCompareFolder.v1',groups:'nsrSapToolbox.zbomCompareGroups.v1',fileNames:'nsrSapToolbox.zbomCompareFileNames.v1'};
const ZBOM_LATEST_LABEL='Latest in SAP';
const ZBOM_SAP_LOOKUP_TOOL='zbomLookup';
const PART_CHANGES_LOOKUP_TOOL='partChangesLookup',TOOLBOX_LOOKUP_TOOLS=new Set([ZBOM_SAP_LOOKUP_TOOL,PART_CHANGES_LOOKUP_TOOL]);
const PART_CHANGES_DEFAULT_FOLDER='C:\\PDFFILES\\'+BOM_COMPARE_LABEL;
const PART_CHANGES_TASK_SYSTEMS_KEY='toolkit.partChanges.taskSystems.v1.';
const PART_CHANGES_LATEST_SO='latest-so',PART_CHANGES_LATEST_QUOTE='latest-quote';
const PART_CHANGES_STORAGE={partChangesInput:'nsrSapToolbox.partChangesInput.v1',partChangesOutputFolder:'nsrSapToolbox.partChangesFolder.v1',
  partChangesTemplate:'nsrSapToolbox.partChangesTemplate.v1',partChangesSeparator:'nsrSapToolbox.partChangesSeparator.v1'};
const partChangesLookups=new Map(),partChangesSelections=new Map(),partChangesCompareTasks=new Set();
const PART_CHANGES_COMPARING_PHASE=new RegExp('^'+PART_CHANGES_COMPARING+'(\\d+)% · ');
let partChangesRows=[];
const CREATE_QUOTE_DEFAULT_VALID_TO_DAYS='30';
const CREATE_QUOTE_DEFAULT_FID_SUFFIX='-01';
const CREATE_QUOTE_STORAGE_KEYS={quotation:'nsrSapToolbox.createQuoteQuotation.v1',input:'nsrSapToolbox.createQuoteInput.v1',validToDays:'nsrSapToolbox.createQuoteValidToDays.v1',fidSuffix:'nsrSapToolbox.createQuoteFidSuffix.v1',crdOverrides:'nsrSapToolbox.createQuoteCrdOverrides.v1'};
const SAP_TOOL_LABELS={zbom:'ZBOM PDF',zbomCompare:'ZBOM Compare',partChanges:BOM_COMPARE_LABEL,createQuote:'Quote Create'}; const sapGuiControlUpdate={online:null,current:'',release:null,workflowReady:false,sapBatch:false,compareCalculation:false,rowBatches:false,partChangesMyLam:false,watchFileCount:false,mock:false};
const TOOLBOX_TASK_PREVIEW_LIMIT=5;
const TOOLBOX_TASK_DAY=new Intl.DateTimeFormat('en-GB',{day:'2-digit',month:'short'});
const sapTools=Object.fromEntries(Object.keys(SAP_TOOL_LABELS).map(name=>[name,{starting:false}]));
let bridgeTaskOnline=false,bridgeTaskPollTimer=0,bridgeTaskPolling=false;
let bridgeBusy=false,bridgeRefreshPending=false,bridgeReconnectTimer=0,zbomRows=[];
let zbomCompareRows=[],zbomCompareStats={ignored:0},zbomCompareSavedGroups=new Map();
let createQuoteRows=[],createQuoteStats={duplicates:0,ignored:0},createQuoteCrdOverrides=new Map(),createQuoteManualCrd=new Set();
let bridgeSupportsToolboxTasks=false,toolboxTasks=[],activeToolboxTask='',toolboxTaskData=null,toolboxTaskReturnFocus=null;
const pendingToolboxSubmissions=new Map();
const zbomQuoteSelections={zbom:new Map(),zbomCompare:new Map()};
const zbomFileNames={zbom:new Map(),zbomCompare:new Map()};
const zbomSapQuotes={zbom:new Map(),zbomCompare:new Map()};
const zbomSapVersions=new Map();
const bwData=createBwDataService(),bwCache=createNsrDataService(),toolboxSourceReader=createFidSourceReaders(fidSourceDefinitions());
let toolboxCacheReady,toolboxRefreshPending=false,bwFidRows=new Map(),bwFidHeaders=new Map();
const BOBJ_LINK_COLUMNS=['Opportunity ID','Quote ID','CTL ID'],bobjLinks={quotes:new Map(),fids:new Map(),ctls:new Map()};
const bobjActiveQuotes=new Map();
let bobjDataRows=new Map(),bobjDataHeaderIndexes=new Map(),bobjDataMissing=new Set(),bobjFileSource=null,bobjIndexedFile=null,bobjDataPromise=null;
let bobjWorker=null,bobjWorkerRequest=0,bobjWorkerPending=new Map();

whenBody(mountBridgeApp);

function readPersistent(key,fallback){const storage=persistentStore();if(!storage)return fallback; try{const value=storage.getItem(key);return value===null?fallback:value}catch(_){return fallback}}
function writePersistent(key,value){const storage=persistentStore();if(!storage)return; try{storage.setItem(key,String(value))}catch(_){/* ignore unavailable or full storage */}}

function loadZbomPreferences(){
  const input=bridgeId('zbomInput'),template=bridgeId('zbomTemplate'),savedInput=readPersistent(ZBOM_STORAGE_KEYS.input,null),savedTemplate=readPersistent(ZBOM_STORAGE_KEYS.template,null);
  if(savedInput!==null)input.value=savedInput;
  if(savedTemplate!==null){
    template.value=savedTemplate===ZBOM_DEFAULT_TEMPLATE.replace('{FIDver}','')?ZBOM_DEFAULT_TEMPLATE:savedTemplate;
    if(template.value!==savedTemplate)saveZbomTemplate();
  }
  loadZbomFileNames('zbom');
}

function saveZbomInput(){const input=bridgeId('zbomInput');if(input)writePersistent(ZBOM_STORAGE_KEYS.input,input.value);}
function saveZbomTemplate(){const input=bridgeId('zbomTemplate');if(input)writePersistent(ZBOM_STORAGE_KEYS.template,input.value);}

function zbomFileNameKey(tool,row){return tool==='zbom'||row.group===ZBOM_COMPARE_STANDALONE?'quote:'+row.quote:'group:'+row.group;}

function loadZbomFileNames(tool,values){
  const keys=tool==='zbom'?ZBOM_STORAGE_KEYS:ZBOM_COMPARE_STORAGE_KEYS;
  try{
    const names=values||JSON.parse(readPersistent(keys.fileNames,'{}'));
    zbomFileNames[tool]=new Map(Object.entries(names).filter(([key,name])=>
      (/^(?:quote:\d{8}|group:.+)$/.test(key)||key.startsWith('quote:')&&toolboxLatestQuoteFid(key.slice(6)))&&
      typeof name==='string'&&!zbomFileNameError(name,tool)));
  }catch(_){zbomFileNames[tool]=new Map()}
}

function saveZbomFileNames(tool){const keys=tool==='zbom'?ZBOM_STORAGE_KEYS:ZBOM_COMPARE_STORAGE_KEYS;writePersistent(keys.fileNames,JSON.stringify(Object.fromEntries(zbomFileNames[tool])));}

function toolboxOutputRoot(folder){
  // Older retries could leave generated subfolders in the saved output path.
  const root=folder.replace(/(?:[\\/]+Retry-[0-9a-f]{12})+[\\/]*$/i,'');
  return root!==folder&&/^[A-Za-z]:$/.test(root)?root+'\\':root;
}

function loadZbomComparePreferences(){
  const input=bridgeId('zbomCompareInput'),template=bridgeId('zbomCompareTemplate'),separator=bridgeId('zbomCompareSeparator'),folder=bridgeId('zbomCompareOutputFolder'),savedInput=readPersistent(ZBOM_COMPARE_STORAGE_KEYS.input,null),
    savedTemplate=readPersistent(ZBOM_COMPARE_STORAGE_KEYS.template,null),savedSeparator=readPersistent(ZBOM_COMPARE_STORAGE_KEYS.separator,null),savedFolder=readPersistent(ZBOM_COMPARE_STORAGE_KEYS.folder,null);
  if(savedInput!==null)input.value=savedInput;
  if(savedTemplate!==null){
    const oldDefault=/^\{fid\}$/i.test(savedTemplate.trim())||savedTemplate===ZBOM_COMPARE_DEFAULT_TEMPLATE.replace('{FIDver}','');
    template.value=oldDefault?ZBOM_COMPARE_DEFAULT_TEMPLATE:savedTemplate;
    if(template.value!==savedTemplate)saveZbomCompareTemplate();
  }
  if(savedSeparator!==null)separator.value=savedSeparator;
  if(savedFolder?.trim()){
    const root=toolboxOutputRoot(savedFolder);
    folder.value=root;
    if(root!==savedFolder)saveZbomCompareFolder();
  }else folder.value=ZBOM_COMPARE_DEFAULT_FOLDER;
  zbomCompareSavedGroups=new Map();
  try{
    const parsed=JSON.parse(readPersistent(ZBOM_COMPARE_STORAGE_KEYS.groups,'{}'));
    if(parsed&&typeof parsed==='object'&&!Array.isArray(parsed))Object.entries(parsed).forEach(([quote,value])=>{if(!(/^2\d{7}$/.test(quote)||toolboxLatestQuoteFid(quote))||typeof value!=='string')return; const group=value===ZBOM_COMPARE_STANDALONE||value===
      '__DO_NOT_COMPARE__'?ZBOM_COMPARE_STANDALONE:normalizeZbomCompareGroup(value,true); if(group)zbomCompareSavedGroups.set(quote,group);});
  }catch(_){zbomCompareSavedGroups=new Map()}
  loadZbomFileNames('zbomCompare');
}

function saveZbomCompareInput(){const input=bridgeId('zbomCompareInput');if(input)writePersistent(ZBOM_COMPARE_STORAGE_KEYS.input,input.value);}
function saveZbomCompareTemplate(){const input=bridgeId('zbomCompareTemplate');if(input)writePersistent(ZBOM_COMPARE_STORAGE_KEYS.template,input.value);}
function saveZbomCompareSeparator(){const input=bridgeId('zbomCompareSeparator');if(input)writePersistent(ZBOM_COMPARE_STORAGE_KEYS.separator,input.value);}
function saveZbomCompareFolder(){const input=bridgeId('zbomCompareOutputFolder');if(input)writePersistent(ZBOM_COMPARE_STORAGE_KEYS.folder,input.value);}
function sanitizeCreateQuoteQuotation(){const input=bridgeId('createQuoteQuotation');if(!input)return ''; const value=String(input.value||'').replace(/\D/g,'').slice(0,10);if(input.value!==value)input.value=value;return value;}
function normalizeCreateQuoteDays(value){return String(value||'').replace(/\D/g,'').slice(0,4);}
function sanitizeCreateQuoteDays(){const input=bridgeId('createQuoteDefaultValidTo');if(!input)return ''; const value=normalizeCreateQuoteDays(input.value);if(input.value!==value)input.value=value;return value;}

function loadCreateQuotePreferences(){
  const quotation=bridgeId('createQuoteQuotation'),input=bridgeId('createQuoteInput'),validToDays=bridgeId('createQuoteDefaultValidTo'),fidSuffix=bridgeId('createQuoteFidSuffix');
  const savedQuotation=readPersistent(CREATE_QUOTE_STORAGE_KEYS.quotation,null),savedInput=readPersistent(CREATE_QUOTE_STORAGE_KEYS.input,null),savedValidToDays=readPersistent(CREATE_QUOTE_STORAGE_KEYS.validToDays,null),
    savedFidSuffix=readPersistent(CREATE_QUOTE_STORAGE_KEYS.fidSuffix,null);
  if(savedQuotation!==null)quotation.value=savedQuotation;
  if(savedInput!==null)input.value=savedInput;
  validToDays.value=savedValidToDays===null?CREATE_QUOTE_DEFAULT_VALID_TO_DAYS:savedValidToDays; fidSuffix.value=savedFidSuffix===null?CREATE_QUOTE_DEFAULT_FID_SUFFIX:savedFidSuffix; const normalized=sanitizeCreateQuoteQuotation();
  if(savedQuotation!==null&&normalized!==savedQuotation)writePersistent(CREATE_QUOTE_STORAGE_KEYS.quotation,normalized);
  const normalizedDays=sanitizeCreateQuoteDays();
  if(savedValidToDays!==null&&normalizedDays!==savedValidToDays)writePersistent(CREATE_QUOTE_STORAGE_KEYS.validToDays,normalizedDays);
  createQuoteCrdOverrides=new Map();createQuoteManualCrd=new Set();let removedBlankOverride=false;
  try{
    const parsed=JSON.parse(readPersistent(CREATE_QUOTE_STORAGE_KEYS.crdOverrides,'{}'));
    if(parsed&&typeof parsed==='object'&&!Array.isArray(parsed))Object.entries(parsed).forEach(([fid,value])=>{if(!/^\d{6}$/.test(fid)||typeof value!=='string')return; if(value.trim())createQuoteCrdOverrides.set(fid,value);else removedBlankOverride=true;});
  }catch(_){createQuoteCrdOverrides=new Map()}
  if(removedBlankOverride)saveCreateQuoteCrdOverrides();
}

function saveCreateQuoteQuotation(){const input=bridgeId('createQuoteQuotation');if(input)writePersistent(CREATE_QUOTE_STORAGE_KEYS.quotation,input.value);}
function saveCreateQuoteInput(){const input=bridgeId('createQuoteInput');if(input)writePersistent(CREATE_QUOTE_STORAGE_KEYS.input,input.value);}
function saveCreateQuoteValidToDays(){const input=bridgeId('createQuoteDefaultValidTo');if(input)writePersistent(CREATE_QUOTE_STORAGE_KEYS.validToDays,input.value);}
function saveCreateQuoteFidSuffix(){const input=bridgeId('createQuoteFidSuffix');if(input)writePersistent(CREATE_QUOTE_STORAGE_KEYS.fidSuffix,input.value);}
function saveCreateQuoteCrdOverrides(){const values={};createQuoteCrdOverrides.forEach((value,fid)=>{values[fid]=value});writePersistent(CREATE_QUOTE_STORAGE_KEYS.crdOverrides,JSON.stringify(values));}
function resetCreateQuoteInput(submittedRows){
  bridgeId('createQuoteInput').value='';
  if(submittedRows){
    for(const {Fid} of submittedRows){createQuoteCrdOverrides.delete(Fid);createQuoteManualCrd.delete(Fid)}
  }else{createQuoteCrdOverrides.clear();createQuoteManualCrd.clear()}
  saveCreateQuoteInput();saveCreateQuoteCrdOverrides();previewCreateQuoteInput();
}
function keepZbomCompareTemp(){return bridgeId('zbomCompareKeepTemp').getAttribute('aria-checked')==='true'}
function saveZbomCompareGroups(){const values={};zbomCompareSavedGroups=new Map(); zbomCompareRows.forEach(row=>{if(row.quote&&row.group){values[row.quote]=row.group;zbomCompareSavedGroups.set(row.quote,row.group)}}); writePersistent(ZBOM_COMPARE_STORAGE_KEYS.groups,
  JSON.stringify(values));}



function toolboxTaskFinished(task){return ['Completed','Failed','Cancelled','Interrupted'].includes(task.state)}
// The MyLam comparison reports its own percentage after the SAP exports have completed.
function partChangesComparePercent(task){
  const match=task.tool==='partChanges'&&task.state==='Completed'&&PART_CHANGES_COMPARING_PHASE.exec(task.phase);
  return match?Number(match[1]):null;
}
// A task list entry has no rows, so remember which tasks end with a comparison once that is known.
function partChangesCompares(task){
  if(task.tool==='partChanges'&&(task.comparisonComplete||partChangesComparePercent(task)!==null||task.spec&&partChangesGroups(task.spec.Rows).size))
    partChangesCompareTasks.add(task.id);
  return partChangesCompareTasks.has(task.id);
}
function toolboxTaskDisplayState(task){
  if(task.state==='Failed'&&task.completed>0&&task.completed<task.count)return 'Partial complete';
  if(partChangesComparePercent(task)!==null)return 'Comparing';
  return task.tool==='partChanges'&&task.state==='Completed'&&!task.comparisonComplete?'Exported':task.state;
}
function toolboxTaskTime(task){const from=Date.parse(task.startedAt||task.createdAt),to=Date.parse(task.finishedAt||'')||Date.now(),seconds=Math.max(0,Math.floor((to-from)/1000));return Math.floor(seconds/60)+'m '+seconds%60+'s'}
function toolboxTaskPhase(task){
  // Older SAPGUIcontrol releases require one connection for automatic tasks.
  if(task.state==='Waiting for SAP window'&&task.phase==='Choose a connection and resubmit.')return 'Leave only the intended SAP connection open.';
  // The state and the progress text already show that a comparison is running and how far it is.
  return partChangesComparePercent(task)===null?task.phase:task.phase.replace(PART_CHANGES_COMPARING_PHASE,'');
}
function paintToolboxTaskDate(element,value,today,yesterday){
  const date=new Date(Date.parse(value));
  element.hidden=Number.isNaN(date.getTime());
  if(element.hidden){element.textContent='';element.removeAttribute('datetime');element.removeAttribute('title');return}
  const dateKey=bwCache.localDate(date),clock=date.toTimeString().slice(0,8),full='Created '+dateKey+' '+clock;
  element.dateTime=date.toISOString();element.title=full;
  if(!today){element.textContent=full;return}
  const year=dateKey.slice(0,4)===today.slice(0,4)?'':' '+date.getFullYear();
  const day=dateKey===today?'Today':dateKey===yesterday?'Yesterday':TOOLBOX_TASK_DAY.format(date)+year;
  element.textContent=day+' · '+clock;
}

function paintToolboxTaskStatus(task){
  const status=bridgeId('toolboxTaskStatus'),state=document.createElement('span'),phase=document.createElement('span'),detail=document.createElement('span');
  state.className='toolbox-task-state';state.textContent=toolboxTaskDisplayState(task);status.dataset.state=state.textContent;
  phase.className='toolbox-task-current';phase.textContent=toolboxTaskPhase(task);phase.hidden=task.phase===(task.prompt||task.error)||(task.state==='Completed'&&task.phase===task.state);
  detail.className='toolbox-task-timing';detail.textContent=toolboxTaskProgress(task).text+' — '+toolboxTaskTime(task);
  status.replaceChildren(state,phase,detail);
}

function toolboxTaskProgress(task){
  const count=Math.max(0,Number(task.count)||0),completed=Math.min(count,Math.max(0,Number(task.completed)||0));
  if(task.tool==='partChanges'){
    const files=task.state==='Running'?Math.max(0,Number(task.observedFiles)||0):0;
    // The report does not supply a total file count. Only published exports get full credit.
    const estimate=files?Math.min(0.9,files/(files+10*count)):0;
    // A task with Compare Groups completes its bar when the comparison files are ready, not when SAP finishes.
    const exportShare=partChangesCompares(task)?0.7:1,comparing=partChangesComparePercent(task);
    const compared=comparing!==null?comparing/100:task.comparisonComplete?1:0;
    const text=completed+' / '+count+' exported'+(files?' · '+files+' BOM '+(files===1?'file':'files')+' detected':'')+
      (comparing!==null?' · comparing '+comparing+'%':'');
    return {count,value:(completed+(count-completed)*estimate)*exportShare+count*(1-exportShare)*compared,text};
  }
  const observed=task.tool==='zbom'&&task.state==='Running'?Math.min(count,Number(task.observedFiles)||0):0;
  const value=Math.max(completed,observed),text=observed>completed
    ?observed+' / '+count+' PDFs detected · '+completed+' completed'
    :completed+' / '+count+' completed';
  return {count,value,text};
}

function paintToolboxTaskProgress(track,task){
  const {count,value,text}=toolboxTaskProgress(task),state=toolboxTaskDisplayState(task);
  const percent=count?Math.round(value/count*1e4)/100:0,indeterminate=task.state==='Running'&&!value;
  track.hidden=!count;
  track.dataset.state=task.state==='Running'||state==='Comparing'?'running':task.state==='Completed'&&percent>=100?'ok':
    ['Failed','Interrupted'].includes(state)?'error':['Partial complete','Needs attention'].includes(state)?'warning':'idle';
  track.style.setProperty('--task-progress',percent+'%');
  track.setAttribute('role','progressbar');track.setAttribute('aria-label',SAP_TOOL_LABELS[task.tool]+' progress');
  track.setAttribute('aria-valuemin','0');track.setAttribute('aria-valuemax','100');
  if(indeterminate)track.removeAttribute('aria-valuenow');else track.setAttribute('aria-valuenow',String(percent));
  track.setAttribute('aria-valuetext',text+' — '+state+' — '+toolboxTaskPhase(task));
  return text;
}

function renderToolboxTasks(){
  const count=toolboxTasks.length,title=bridgeId('toolboxTaskHistoryTitle'),label='All tasks ('+count+')';
  if(title.textContent!==label)title.textContent=label;
  bridgeId('toolboxTaskMore').hidden=count<=TOOLBOX_TASK_PREVIEW_LIMIT&&!toolboxTasks.some(task=>task.canDelete===true);
  renderToolboxTaskList(bridgeId('toolboxTaskList'),toolboxTasks.slice(0,TOOLBOX_TASK_PREVIEW_LIMIT));
  if(bridgeId('toolboxTaskHistory').open)renderToolboxTaskList(bridgeId('toolboxTaskHistoryList'),toolboxTasks);
}

function renderToolboxTaskList(list,tasks){
  const now=new Date(),today=bwCache.localDate(now),signature=today+JSON.stringify(tasks.length?tasks:bridgeSupportsToolboxTasks)+tasks.map(toolboxTaskPhase).join('\n');
  if(list.dataset.tasks===signature)return;
  const scrollTop=list.scrollTop,focused=document.activeElement;
  const focusedRow=list.contains(focused)?focused.closest('.toolbox-task-row'):null;
  let nextFocus;
  list.dataset.tasks=signature;list.replaceChildren();
  now.setDate(now.getDate()-1);const yesterday=bwCache.localDate(now);
  if(!tasks.length){const note=document.createElement('span');note.className='bridge-note';note.textContent=bridgeSupportsToolboxTasks?'No tasks yet.':'Update and restart SAPGUIcontrol to use Tasks.';list.appendChild(note)}
  tasks.forEach(task=>{
    const row=document.createElement('div'),open=document.createElement('button'),title=document.createElement('strong'),date=document.createElement('time');
    const summary=document.createElement('div'),state=document.createElement('span'),displayState=toolboxTaskDisplayState(task);
    // A running comparison keeps the progress bar of its exported task.
    const exported=task.state==='Completed',isComplete=exported&&displayState!=='Comparing';
    row.className='toolbox-task-row';row.dataset.state=displayState;row.dataset.taskId=task.id;row.dataset.tool=task.tool;
    open.type='button';open.className='toolbox-task-open';open.title='View task details';title.textContent=SAP_TOOL_LABELS[task.tool]||'Lookup in SAP';
    date.className='toolbox-task-date';paintToolboxTaskDate(date,task.createdAt,today,yesterday);
    open.append(title,date);open.addEventListener('click',()=>openToolboxTask(task.id,list));
    const phaseText=isComplete&&task.error?task.error:toolboxTaskPhase(task);
    summary.className='toolbox-task-summary';state.className='toolbox-task-state';state.textContent=displayState;state.title=phaseText;
    summary.appendChild(state);row.append(open,summary);
    if(isComplete){
      const count=document.createElement('span');count.className='toolbox-task-count';count.textContent=task.completed+' / '+task.count;state.append(' ',count);
    }
    if((exported||displayState==='Partial complete')&&['zbom','zbomCompare','partChanges'].includes(task.tool)){
      const folder=document.createElement('button');folder.type='button';folder.className='bridge-btn light small';folder.textContent='Open folder';
      folder.addEventListener('click',async()=>{
        folder.disabled=true;
        try{const {task:detail}=await bridgeRequest('GET','/toolbox/tasks/'+task.id,undefined,10000);await openBridgeFolder(toolboxTaskOutputFolder(detail))}
        catch(error){alert('Could not open the folder.\n\n'+bridgeError(error))}finally{folder.disabled=false}
      });summary.appendChild(folder);
    }
    if(list.id==='toolboxTaskHistoryList'&&task.canDelete===true){
      const remove=document.createElement('button');remove.type='button';remove.className='toolbox-task-delete';
      remove.textContent='🗑️';remove.title='Delete task';remove.setAttribute('aria-label',remove.title);
      remove.addEventListener('click',async()=>{
        if(!confirm('Delete this task and its local history files?'))return;
        remove.disabled=true;
        try{
          const result=await bridgeRequest('POST','/toolbox/tasks/'+task.id+'/delete',{},30000);
          if(task.tool==='partChanges')GM_deleteValue(PART_CHANGES_TASK_SYSTEMS_KEY+task.id);
          toolboxTasks=toolboxTasks.filter(entry=>entry.id!==task.id);
          if(bridgeId('toolboxTaskHistory').open&&(document.activeElement===remove||document.activeElement===document.body))
            bridgeId('toolboxTaskHistoryClose').focus();
          renderToolboxTasks();
          if(result.cleanupPending)alert('Task deleted. SAPGUIcontrol will retry cleaning up its local files on restart.');
        }catch(error){alert('Could not delete the task.\n\n'+bridgeError(error))}finally{remove.disabled=false}
      });summary.appendChild(remove);
    }
    if(!isComplete||(phaseText&&phaseText!==task.state)){
      const detail=document.createElement('div'),phase=document.createElement('span');
      detail.className='toolbox-task-phase';phase.textContent=phaseText;phase.title=task.error||phaseText;detail.appendChild(phase);
      if(isComplete)row.appendChild(detail);
      else{
        const progress=document.createElement('div'),track=document.createElement('div'),count=document.createElement('span');
        progress.className='toolbox-task-progress';track.className='progress-track';
        count.textContent=paintToolboxTaskProgress(track,task);detail.appendChild(count);progress.append(track,detail);row.appendChild(progress);
      }
    }
    list.appendChild(row);
    if(focusedRow?.dataset.taskId===task.id)nextFocus=focused.matches('.toolbox-task-open')?open:
      summary.querySelector(focused.matches('.toolbox-task-delete')?'.toolbox-task-delete':'.bridge-btn')||open;
  });
  if(scrollTop)list.scrollTop=scrollTop;
  nextFocus?.focus({preventScroll:true});
}

async function pollToolboxTasks(){
  if(!bridgeSupportsToolboxTasks)return;
  const data=await bridgeRequest('GET','/toolbox/tasks',undefined,10000);
  toolboxTasks=(data.tasks||[]).filter(task=>!TOOLBOX_LOOKUP_TOOLS.has(task.tool)||toolboxTaskFinished(task)&&task.window);renderToolboxTasks();
  const summary=toolboxTasks.find(task=>task.id===activeToolboxTask);
  if(summary&&summary.revision!==toolboxTaskData?.revision){const {task}=await bridgeRequest('GET','/toolbox/tasks/'+summary.id,undefined,10000);if(activeToolboxTask===summary.id)renderToolboxTask(task)}
  else if(summary&&toolboxTaskData&&summary.observedFiles!==toolboxTaskData.observedFiles){
    toolboxTaskData.observedFiles=summary.observedFiles;
    paintToolboxTaskProgress(bridgeId('toolboxTaskAnimation'),toolboxTaskData);paintToolboxTaskStatus(toolboxTaskData);
  }
}

function closeToolboxTask(){
  if(!activeToolboxTask)return;
  destroyPartChangesSystemPickers(bridgeId('toolboxTaskResults'));
  bridgeId('toolboxTaskDialog').close();activeToolboxTask='';toolboxTaskData=null;
  const {id,list}=toolboxTaskReturnFocus;
  const trigger=list.querySelector(`[data-task-id="${CSS.escape(id)}"] .toolbox-task-open`);
  const fallback=bridgeId(list.id==='toolboxTaskHistoryList'?'toolboxTaskHistoryClose':'toolboxTaskMore');
  (trigger||fallback).focus({preventScroll:true});toolboxTaskReturnFocus=null;
}

async function openToolboxTask(id,list){
  toolboxTaskReturnFocus={id,list};activeToolboxTask=id;toolboxTaskData=null;
  const dialog=bridgeId('toolboxTaskDialog');delete dialog.dataset.tool;bridgeId('toolboxTaskTitle').textContent='Loading task…';
  paintToolboxTaskDate(bridgeId('toolboxTaskCreated'));
  bridgeId('toolboxTaskAnimation').hidden=true;
  destroyPartChangesSystemPickers(bridgeId('toolboxTaskResults'));
  for(const name of ['Status','Actions','Prompt','Results']){const section=bridgeId('toolboxTask'+name);section.replaceChildren();delete section.dataset.task}
  dialog.showModal();
  bridgeId('toolboxTaskClose').focus();
  try{const data=await bridgeRequest('GET','/toolbox/tasks/'+id,undefined,10000);if(activeToolboxTask===id)renderToolboxTask(data.task)}catch(error){if(activeToolboxTask===id)bridgeId('toolboxTaskStatus').textContent=bridgeError(error)}
}

function toolboxTaskWindows(task){return task.windows||(task.window?[task.window]:[])}

function toolboxTaskOutputFolder(task){
  if(task.spec.OutputFolder||task.tool!=='partChanges')return task.spec.OutputFolder;
  const path=task.results.find(row=>row.Status==='Completed'&&row.Files?.length)?.Files[0];
  return path?.replace(/[^\\/]+$/,'')||'';
}

function paintToolboxComparisonResults(results,files){
  for(const cell of results.querySelectorAll('[data-export-group]'))cell.textContent=files?.[cell.dataset.exportGroup]?'Completed':'Exported';
  for(const link of results.querySelectorAll('[data-comparison-group]')){
    link.title=files?.[link.dataset.comparisonGroup]||'';
    link.parentElement.hidden=!link.title;
  }
}

function renderToolboxTask(task){
  toolboxTaskData=task;bridgeId('toolboxTaskDialog').dataset.tool=task.tool;bridgeId('toolboxTaskTitle').textContent=SAP_TOOL_LABELS[task.tool]||'Lookup in SAP';
  paintToolboxTaskDate(bridgeId('toolboxTaskCreated'),task.createdAt);
  paintToolboxTaskProgress(bridgeId('toolboxTaskAnimation'),task);paintToolboxTaskStatus(task);
  bridgeId('toolboxTaskPrompt').textContent=task.prompt||task.error;
  const folder=toolboxTaskOutputFolder(task),windows=toolboxTaskWindows(task);
  const partChanges=task.tool==='partChanges',finished=toolboxTaskFinished(task);
  const groups=partChanges&&finished&&task.finishedAt&&task.completed?partChangesGroups(task.spec.Rows):null;
  if(groups?.size){
    const exported=new Set(task.spec.Rows.filter((row,index)=>task.results[index]?.Status==='Completed'&&task.results[index].Files?.length===1));
    for(const [group,rows] of groups)if(rows.length<2||rows.some(row=>!exported.has(row)))groups.delete(group);
  }
  const results=bridgeId('toolboxTaskResults');
  const actions=bridgeId('toolboxTaskActions'),signature=JSON.stringify([task.state,task.finishedAt,windows,task.prompt,folder,Boolean(groups?.size)]);
  if(actions.dataset.task!==task.id+signature){
    actions.dataset.task=task.id+signature;actions.replaceChildren();
    const add=(label,run,style='light')=>{const button=document.createElement('button');button.type='button';button.className='bridge-btn '+style;button.textContent=label;button.addEventListener('click',async()=>{button.disabled=true;try{await run()}catch(error){bridgeId('toolboxTaskPrompt').textContent=bridgeError(error)}finally{button.disabled=false}});actions.appendChild(button)};
    if(folder)add('Open folder',()=>openBridgeFolder(folder),'blue');
    if(groups?.size){
      add(PART_CHANGES_ACTION_LABEL,async()=>{
        const port=new URL(bridgeClient.origin).port;
        const target=sapGuiControlUpdate.partChangesMyLam?bridgeClient.origin+PART_CHANGES_LAUNCH_PATH:location.origin+BRIDGE_PATH;
        const key=PART_CHANGES_TASK_SYSTEMS_KEY+task.id,saved=GM_getValue(key,{});
        const choices=[...results.querySelectorAll('select[data-group]')].map(select=>({Group:select.dataset.group,select}));
        const systems=Object.fromEntries([...partChangesGroups(choices)].flatMap(([group,rows])=>{
          const pickers=rows.map(row=>row.select);
          // Preserve an older group selection until a row in that group is edited.
          if(Array.isArray(saved[group])&&pickers.every(select=>!select.dataset.edited))return [[group,saved[group]]];
          return pickers.some(select=>select.dataset.systems!==select.dataset.original)?[[group,{rows:pickers.map(select=>select.dataset.systems)}]]:[];
        }));
        const custom=Object.keys(systems).length>0;
        if(custom)await GM.setValue(key,systems);
        else if(GM_getValue(key))GM_deleteValue(key);
        const scope=custom?'&systems='+encodeURIComponent(JSON.stringify(systems)):'';
        GM_openInTab(target+'#'+PART_CHANGES_MYLAM_TAB+'='+task.id+'&port='+port+scope,{active:true});
      });
    }
    if(!TOOLBOX_LOOKUP_TOOLS.has(task.tool))add('Edit as new task',async()=>{
      await restoreToolboxTask(task.id);
      if(activeToolboxTask!==task.id)return;
      closeToolboxTask();bridgeId('toolboxTaskHistory').close();
    });
    for(const [index,window] of windows.entries())if(window.windowHandle)add(windows.length>1?'Show SAP '+(index+1):'Show SAP',async()=>{
      const health=await bridgeRequest('GET','/health',undefined,10000),session=health.sessions.find(item=>item.sessionKey===window.sessionKey);
      if(!session)throw new Error('The task’s SAP window is closed or unavailable.');
      await bridgeRequest('POST','/windows/activate',{handle:session.windowHandle},20000);
    });
    if(task.prompt){add('Continue',()=>actToolboxTask(task.id,'continue'));add('Stop',()=>actToolboxTask(task.id,'stop'))}
    else if(!finished)add('Stop',()=>actToolboxTask(task.id,'cancel'));
  }
  const unfinishedReason=finished&&task.results?.some(row=>row.Status==='Pending'||row.Status==='Formatted')?
    task.error||toolboxTaskPhase(task)||'The task ended before this row finished.':'';
  const resultSignature=JSON.stringify([task.results,unfinishedReason,Boolean(groups?.size)]);
  if(results.dataset.task===task.id+resultSignature){
    if(partChanges)paintToolboxComparisonResults(results,task.comparisonFiles);
    return;
  }
  destroyPartChangesSystemPickers(results);
  results.dataset.task=task.id+resultSignature;results.replaceChildren();
  const systemChoices=new Map(),suggestSystems=groups?.size?createPartChangesSystemGuesser(partChangesSystemNames()):null;
  if(groups?.size){
    const saved=GM_getValue(PART_CHANGES_TASK_SYSTEMS_KEY+task.id,{});
    for(const [group,rows] of groups){
      const scope=saved[group],legacy=Array.isArray(scope)?scope:[];
      rows.forEach((row,index)=>{
        const original=row.Fields?.myLamSystem||'';
        const value=legacy.length?legacy[index%legacy.length]:scope?.rows?.[index]??original;
        systemChoices.set(row,{value,original,extra:legacy.length>rows.length?legacy:[]});
      });
    }
  }
  const table=document.createElement('table');table.className='zbom-table';const head=document.createElement('tr');
  for(const name of partChanges?['Sales Order / Quotation · FID','FID System','Group','Status','Result']:['Quotation / FID','Status','Result']){const th=document.createElement('th');th.textContent=name;head.appendChild(th)}table.appendChild(head);
  const addFileLink=(container,label,path='')=>{
    const link=document.createElement('button');link.type='button';link.className='compare-file-link';link.textContent=label;link.title=path;
    link.addEventListener('click',async()=>{
      link.disabled=true;
      try{await bridgeFileRequest('openFile',{path:link.title},30000)}
      catch(error){bridgeId('toolboxTaskPrompt').textContent=bridgeError(error)}finally{link.disabled=false}
    });
    container.appendChild(link);
    return link;
  };
  for(const [index,row] of (task.results||[]).entries()){
    let quote=toolboxLatestQuoteFid(row.Quote)?/^SAP quote (\d{8})\b/.exec(row.Message)?.[1]||ZBOM_LATEST_LABEL:row.Quote,message=row.Message;
    const input=partChanges?task.spec.Rows[index]:null;
    if(partChanges){
      const {document,message:detail}=partChangesResultDetails(input,row);
      if(document){quote=document.number;message=[document.version,'Created '+document.created,detail].join(' · ')}
      else if(!quote)quote=input.Fields.lookupPolicy==='latest-quote'?'Latest quotation':'Latest SO';
    }
    const tr=document.createElement('tr');
    tr.insertCell().textContent=[quote,row.Fid].filter(Boolean).join(' / ');
    if(partChanges){
      const system=tr.insertCell(),choice=systemChoices.get(input);
      if(choice){
        const subject=[quote,row.Fid?'FID '+row.Fid:'',input.Group].filter(Boolean).join(' · ');
        const select=partChangesSystemSelect(system,subject,choice.value,false,picker=>{picker.dataset.edited='true';system.removeAttribute('title')},
          suggestSystems(input.Fields.system||bobjValue(row.Fid||input.Fid,'System Description'),input.Fields.product));
        select.dataset.group=input.Group;select.dataset.original=choice.original;
        if(choice.extra.length)system.title='Saved group Systems · '+choice.extra.join(', ');
      }else system.textContent='—';
      tr.insertCell().textContent=input.Group===ZBOM_COMPARE_STANDALONE?'Standalone':input.Group;
    }
    const status=tr.insertCell(),result=tr.insertCell();
    const incomplete=finished&&['Pending','Formatted'].includes(row.Status);
    status.textContent=incomplete?'Not completed':row.Status;
    result.textContent=message||(incomplete?unfinishedReason:'');
    if(partChanges&&row.Status==='Completed')status.dataset.exportGroup=input.Group;
    for(const path of row.Files||[]){result.append(' ');addFileLink(result,path.split(/[\\/]/).pop(),path)}
    if(groups?.get(input.Group)?.[0]===input){
      const report=document.createElement('div');report.hidden=true;result.appendChild(report);
      addFileLink(report,'Comparison Excel').dataset.comparisonGroup=input.Group;
    }
    table.appendChild(tr);
  }
  results.appendChild(table);
  if(partChanges)paintToolboxComparisonResults(results,task.comparisonFiles);
}

async function actToolboxTask(id,action){const data=await bridgeRequest('POST','/toolbox/tasks/'+id+'/'+action,{},10000);if(activeToolboxTask===id)renderToolboxTask(data.task);await pollToolboxTasks()}

async function restoreToolboxTask(id){
  const {task}=await bridgeRequest('GET','/toolbox/tasks/'+id,undefined,10000),form=task.spec.Form||{};if(sapToolRunning(task.tool))return;
  if(task.tool==='partChanges'){
    restorePartChangesInputs(task);
    setToolTab(task.tool);setSapProgress(task.tool,'Inputs copied. Edit and press Start when ready.','idle');
    return;
  }
  for(const [name,value] of Object.entries(form)){
    if(name==='zbomCompareOutputFolder')continue;
    const field=bridgeId(name);if(field&&field.matches('input,textarea'))field.value=String(value??'');
  }
  if(zbomQuoteSelections[task.tool]){
    zbomQuoteSelections[task.tool]=new Map(Object.entries(form.quoteSelections||{}).filter(([,choice])=>
      /^\d{6}$/.test(choice?.fid)&&(/^\d{8}$/.test(choice.quote)||toolboxLatestQuoteFid(choice.quote)===choice.fid)));
    for(const {Quote,Version} of task.spec.Rows)if(/^\d{8}$/.test(Quote)&&isZbomFid(Version)&&!zbomSapVersions.has(Quote))zbomSapVersions.set(Quote,Version);
    loadZbomFileNames(task.tool,form.fileNames||{});saveZbomFileNames(task.tool);
  }
  if(form.groups)zbomCompareSavedGroups=new Map(Object.entries(form.groups));
  if(form.crdOverrides)createQuoteCrdOverrides=new Map(Object.entries(form.crdOverrides));if(form.manualCrd)createQuoteManualCrd=new Set(form.manualCrd);
  if(form.keepTemp!==undefined)bridgeId('zbomCompareKeepTemp').setAttribute('aria-checked',String(Boolean(form.keepTemp)));
  if(task.tool==='zbom'){saveZbomInput();saveZbomTemplate();previewZbomInput()}
  else if(task.tool==='zbomCompare'){saveZbomCompareInput();saveZbomCompareTemplate();saveZbomCompareSeparator();previewZbomCompareInput()}
  else{saveCreateQuoteQuotation();saveCreateQuoteInput();saveCreateQuoteValidToDays();saveCreateQuoteFidSuffix();saveCreateQuoteCrdOverrides();previewCreateQuoteInput()}
  setToolTab(task.tool);setSapProgress(task.tool,'Inputs copied. Edit and press Start when ready.','idle');
}

async function submitToolboxTask(tool){
  const editor=bridgeId(tool+'PreviewBody')?.querySelector('.zbom-file-input');if(editor){editor.focus();return}
  if(sapToolRunning(tool))return;sapTools[tool].starting=true;setBridgeBusy(bridgeBusy);setSapProgress(tool,'Preparing task…');
  try{
    let spec=pendingToolboxSubmissions.get(tool);
    if(!spec){
      if(!bridgeSupportsToolboxTasks)throw new Error('Update and restart SAPGUIcontrol to use Tasks.');
      if(!toolboxWorkflowReady(tool))throw new Error('Update and restart SAPGUIcontrol to run Tasks.');
      await ensureBobjDataSource();
      if(toolboxSourcesLoading())throw new Error('Wait for Data Sources to finish loading, then press Start.');
      if(tool==='partChanges'){
        spec=preparePartChangesTask();
        if(spec.Rows.some(row=>!row.Fid&&row.Group!==ZBOM_COMPARE_STANDALONE)){
          spec.Form.exportFidSource=await snapshotExportFidSource(spec.Template,['primary product','fid description']);
        }
      }
      else{
      let rows,names,template='',outputFolder='',form={};
      const fields={zbom:['zbomInput','zbomTemplate'],zbomCompare:['zbomCompareInput','zbomCompareTemplate','zbomCompareSeparator','zbomCompareOutputFolder'],createQuote:['createQuoteInput','createQuoteQuotation','createQuoteDefaultValidTo','createQuoteFidSuffix']}[tool];
      if(tool==='zbom'){
        previewZbomInput();rows=zbomRows;
        const info=zbomTemplateInfo();if(info.error)throw new Error(info.error);
        template=info.template;outputFolder=ZBOM_PDF_FOLDER;names=plannedZbomNames(rows,template);
        const count=rows.filter(row=>!row._error&&!row._duplicate).length;
        if(count>500)throw new Error('A maximum of 500 unique quotations can be exported at once.');
        if(count>100&&!unsafeWindow.confirm('You are about to print '+count+' quotations through SAP. Continue?')){setSapProgress(tool,'Ready.','idle');return}
      }else if(tool==='zbomCompare'){
        previewZbomCompareInput();rows=zbomCompareRows;
        template=bridgeId('zbomCompareTemplate').value.trim()||ZBOM_COMPARE_DEFAULT_TEMPLATE;outputFolder=normalizedZbomCompareFolder();
        names=plannedZbomCompareFileNames();form.groups=Object.fromEntries(rows.filter(row=>!row._error).map(row=>[row.quote,row.group]));form.keepTemp=keepZbomCompareTemp();
      }else{
        previewCreateQuoteInput();if(!/^\d{8,10}$/.test(sanitizeCreateQuoteQuotation())||!createQuoteRows.length)throw new Error('Enter a reference quotation and at least one FID.');
        if(!validateCreateQuotePreview())return;rows=createQuoteRows;form.crdOverrides=Object.fromEntries(createQuoteCrdOverrides);form.manualCrd=Array.from(createQuoteManualCrd);
      }
      if(!rows.length)throw new Error('Enter at least one quotation, FID or CTL ID.');
      fields.forEach(name=>{form[name]=bridgeId(name).value});
      if(zbomQuoteSelections[tool]){
        const choices=new Map(zbomQuoteSelections[tool]);
        for(const row of rows)if(row.quotePicker){
          const {key,fid}=row.quotePicker;
          choices.set(key,{...choices.get(key),fid,quote:row.quote});
        }
        if(choices.size)form.quoteSelections=Object.fromEntries(choices);
        const templateFields=bobjTemplateFields(template),lookups=rows.filter(row=>!row._duplicate&&!toolboxLatestQuoteFid(row.quote)&&
          needsZbomSapLookup(row,templateFields,zbomFileNames[tool].has(zbomFileNameKey(tool,row))));
        if(lookups.length)form.sapLookupFids=Object.fromEntries(lookups.map(row=>[row.quote,row.inputFid||(row.ctl?bobjFidKey(row.fid):'')]));
        form.fileNames=Object.fromEntries(zbomFileNames[tool]);
      }
      spec={Id:crypto.randomUUID().replace(/-/g,''),Tool:tool,Workflow:createToolboxWorkflow(tool,outputFolder),OutputFolder:outputFolder,Template:template,KeepTemp:Boolean(form.keepTemp),Form:form,Rows:rows.filter(row=>!row._duplicate).map(row=>{
        if(tool!=='createQuote'&&row._error){
          const [line,values]=JSON.parse(row.key);
          return {Quote:row.quote||values.filter(Boolean).join(' ').slice(0,120),Fid:row.fid||'',Fields:{inputError:'Row '+line+': '+row._error}};
        }
        return {
          Quote:row.quote||'',Fid:row.fid||'',Group:row.group||'',Crd:row.crd||'',ValidTo:row.validTo||'',Version:tool==='createQuote'?row.version||createQuotePreviewFid(row.fid):row.sapVersion||'',FileName:names?.get(row.quote)||'',FileNameEdited:Boolean(zbomFileNames[tool]?.has(zbomFileNameKey(tool,row))),FromBobj:hasBobjFid(bobjFidKey(row.fid)),
          Fields:Object.fromEntries(Array.from(template.matchAll(/\{([^{}]+)\}/g),match=>[normalizeBobjHeader(match[1]),bobjTemplateValue(row,match[1])]))
        };
      })};
      if(tool!=='createQuote'&&(form.sapLookupFids||spec.Rows.some(row=>!row.Fields.inputError&&(!bobjFidKey(row.Fid)||!row.Version&&!row.FileNameEdited&&row.Fields['system description']==='')))){
        setSapProgress(tool,'Preparing filename fields for export…');
        form.exportFidSource=await snapshotExportFidSource(template);
      }
      }
      pendingToolboxSubmissions.set(tool,spec);
    }
    setSapProgress(tool,'Submitting task…');const data=await bridgeRequest('POST','/toolbox/tasks',spec,30000);
    if(data.task?.id!==spec.Id||data.task?.tool!==tool)throw new Error('SAPGUIcontrol did not confirm the submitted task.');
    pendingToolboxSubmissions.delete(tool);
    const cleared=clearSubmittedToolboxInput(spec);
    if(tool==='partChanges')partChangesCompares(data.task);
    const {spec:_,results,log,...summary}=data.task;toolboxTasks=[summary,...toolboxTasks.filter(task=>task.id!==summary.id)];renderToolboxTasks();
    if(activeToolboxTask===summary.id)renderToolboxTask(data.task);
    setSapProgress(tool,cleared?'Submitted. Follow progress in Tasks.':'Submitted. Current input kept. Follow progress in Tasks.','ok');
  }catch(error){if(error.status)pendingToolboxSubmissions.delete(tool);const message=bridgeError(error)+(pendingToolboxSubmissions.has(tool)?' Press Start to check the same submission.':'');setSapProgress(tool,message,'error')}
  finally{sapTools[tool].starting=false;setBridgeBusy(bridgeBusy)}
}

function clearSubmittedToolboxInput({Tool:tool,Form:form,Rows:submittedRows}){
  const name=tool+'Input',input=bridgeId(name);
  // An idempotent retry can confirm an older submission after the user edits the next input.
  if(input.value!==String(form[name]??''))return false;
  if(tool==='createQuote'){
    const quotation=bridgeId('createQuoteQuotation');
    if(quotation.value!==form.createQuoteQuotation||JSON.stringify(Object.fromEntries(createQuoteCrdOverrides))!==JSON.stringify(form.crdOverrides)||
        JSON.stringify(Array.from(createQuoteManualCrd))!==JSON.stringify(form.manualCrd))return false;
    quotation.value='';saveCreateQuoteQuotation();resetCreateQuoteInput(submittedRows);return true;
  }
  if(tool==='partChanges'){
    if(partChangesRows.length!==submittedRows.length||partChangesRows.some((row,index)=>{
      const saved=submittedRows[index];
      return row.Group!==saved.Group||row.selection!==saved.Fields.selection||row.myLamSystem!==saved.Fields.myLamSystem;
    }))return false;
    resetPartChangesInput();return true;
  }
  if(JSON.stringify(Object.fromEntries(zbomFileNames[tool]))!==JSON.stringify(form.fileNames||{}))return false;
  const rows=tool==='zbom'?zbomRows:zbomCompareRows;
  if(rows.some(row=>row.quotePicker&&form.quoteSelections?.[row.quotePicker.key]?.quote!==row.quote))return false;
  if(tool==='zbomCompare'&&rows.some(row=>!row._error&&form.groups?.[row.quote]!==row.group))return false;
  input.value='';
  zbomFileNames[tool].clear();saveZbomFileNames(tool);
  if(tool==='zbom'){saveZbomInput();previewZbomInput()}
  else{saveZbomCompareInput();previewZbomCompareInput()}
  return true;
}

function mountBridgeApp(){
  if(toolkitExpired())return;
  if(document.documentElement.dataset.nsrSapBridgeMounted)return;
  document.documentElement.dataset.nsrSapBridgeMounted='1';document.documentElement.lang='en';
  setWorkspaceHead('SAP Toolbox','🧰',WORKSPACE_CSS+BRIDGE_CSS);
  document.body.className='toolbox-app';document.body.innerHTML=BRIDGE_HTML;
  mountWorkNotes();bindBridgeEvents();refreshToolbox();
}

function bridgeId(id){return document.getElementById(id)}
function setSapProgress(tool,phase,state='running',detail=''){
  sapTools[tool].phase=phase;const text=bridgeId(tool+'ProgressText');if(!text)return;
  text.textContent=detail?phase+' — '+detail:phase;text.hidden=state==='idle';text.dataset.state=state;
  text.className='bridge-note'+(state==='error'?' row-error':state==='warning'?' row-warning':state==='ok'?' row-ok':'');
}

async function openBridgeFolder(path){
  const result=await bridgeFileRequest('openDirectory',{path:path,activate:true},30000);
  if(result.foreground!==true)alert('The folder has been opened in File Explorer.\nPlease switch to it manually.\n\n'+path);
  return result;
}

function setBridgeBadge(text,kind,details=''){
  const badge=bridgeId('bridgeBadge');badge.textContent=text;badge.className='bridge-badge'+(kind?' '+kind:'');
  bridgeId('bridgeStatus').textContent=details;
}
function bridgeError(error){return error&&error.message?error.message:String(error||'Unknown SAPGUIcontrol error.');}
function bobjFidKey(value){const match=String(value||'').trim().match(/(?:^|\D)(\d{6})(?!\d)/);return match?match[1]:'';}
function bobjQuoteKey(value){return /^(?:00)?(\d{8})(?:\.0+)?$/.exec(String(value??'').trim())?.[1]||'';}
function bobjField(row,index){const value=String(row?.[index]??'').trim();return value==='#'?'':value;}
function bobjValue(fid,header){
  const key=bobjFidKey(fid),field=normalizeBobjHeader(header);
  return bobjField(bwFidRows.get(key),bwFidHeaders.get(field))||bobjField(bobjDataRows.get(key),bobjDataHeaderIndexes.get(field));
}

// Keep conflicting keys in the indexes so they can be distinguished from misses.
function indexUniqueBobjValue(map,key,value){map.set(key,map.has(key)&&map.get(key)!==value?'':value)}

function indexBobjLink(map,key,link){
  if(!map.has(key)){map.set(key,link);return}
  const previous=map.get(key);
  if(!previous)return;
  if(previous.fid!==link.fid||previous.quote&&link.quote&&previous.quote!==link.quote){map.set(key,null);return}
  if(!previous.quote&&link.quote||!previous.customer&&link.customer||!previous.systemDescription&&link.systemDescription){
    map.set(key,{fid:link.fid,quote:previous.quote||link.quote,customer:previous.customer||link.customer,systemDescription:previous.systemDescription||link.systemDescription});
  }
}

function indexBobjQuotes(rows,columns){
  const [fidIndex,quoteIndex,ctlIndex]=BOBJ_LINK_COLUMNS.map(field=>columns.indexOf(field));
  const customerIndex=columns.indexOf('Customer group 2'),systemIndex=columns.indexOf('System Description');
  const dateIndexes=['Create Date','Creation Date','Created Date'].map(column=>columns.indexOf(column)).filter(index=>index>=0);
  const changeIndex=columns.findIndex(column=>['Change Cat','Change Type','Change Category','CTL Change Type'].includes(column)),stageIndex=columns.findIndex(column=>['Stage','CTL Stage'].includes(column));
  for(const row of rows){
    const fid=bobjField(row,fidIndex),quote=bobjQuoteKey(row[quoteIndex]),ctl=bobjField(row,ctlIndex);
    if(!/^\d{6}$/.test(fid))continue;
    const link={fid,quote,customer:bobjField(row,customerIndex),systemDescription:bobjField(row,systemIndex)};
    const hasCtl=/^\d{10}$/.test(ctl);
    if(quote){
      indexBobjLink(bobjLinks.quotes,quote,link);
      if(!bobjLinks.fids.has(fid))bobjLinks.fids.set(fid,new Map());
      const quotes=bobjLinks.fids.get(fid);
      if(!quotes.has(quote))quotes.set(quote,new Map());
      if(hasCtl){
        const ctls=quotes.get(quote),previous=ctls.get(ctl);
        let created='';
        for(const index of dateIndexes){created=sourceDate(row[index],bwCache);if(created)break}
        ctls.set(ctl,{
          created:!previous?.created||created&&created<previous.created?created:previous.created,
          change:bobjField(row,changeIndex)||previous?.change||'',stage:bobjField(row,stageIndex)||previous?.stage||''
        });
      }
    }
    if(hasCtl)indexBobjLink(bobjLinks.ctls,ctl,link);
  }
}

function zbomQuoteFid(quote){
  return bobjFidKey(zbomSapVersions.get(quote))||(bobjLinks.quotes.has(quote)?bobjLinks.quotes.get(quote)?.fid:bobjActiveQuotes.get(quote));
}

function matchBobjInput(row,ctl,item,tool){
  const inputFid=row.inputFid=bobjFidKey(row.fid),choice=zbomQuoteSelections[tool].get(item.key);
  const lookupFid=inputFid||choice?.fid||zbomQuoteFid(row.quote)||bobjLinks.ctls.get(ctl)?.fid;
  const sapQuotes=zbomSapQuotes[tool].get(lookupFid),pickerFid=inputFid&&row.quote?'':inputFid||((sapQuotes||choice)&&lookupFid);
  if(pickerFid){
    const active=bobjQuoteKey(bobjValue(pickerFid,'Active Quote Number'));
    const choices=new Map(bobjLinks.fids.get(pickerFid));
    for(const quote of sapQuotes||[])if(!choices.has(quote))choices.set(quote,null);
    for(const quote of [active,row.quote])if(quote&&!choices.has(quote))choices.set(quote,null);
    const latest='sap-latest:'+pickerFid;
    row.quotePicker={key:item.key,fid:pickerFid,choices,latest:sapQuotes?'':latest};
    if(choice?.fid===pickerFid){
      row.quote=choice.quote;
      if(choice.resolvedLatest&&bobjLinks.ctls.get(ctl)?.quote!==row.quote)ctl=row.ctl='';
    }
    else if(!ctl)row.quote||=latest;
    if(row.quote===latest){
      ctl=row.ctl='';
      if(!sapQuotes)return row;
      selectZbomQuote(tool,row,sapQuotes[0],true);row.quote=sapQuotes[0];
    }
    if(row.quote&&!choices.has(row.quote)){
      choices.set(row.quote,null);
      if(!zbomSapVersions.has(row.quote))row._lookup='Selected quotation is unavailable in the data sources.';
    }
  }
  const ctlLink=bobjLinks.ctls.get(ctl);
  const matchedCtl=ctlLink&&(!inputFid||inputFid===ctlLink.fid)&&(!row.quote||!ctlLink.quote||row.quote===ctlLink.quote)?ctlLink:null;
  const fid=inputFid||matchedCtl?.fid||'';
  if(!row.quote){
    const choices=row.quotePicker?.choices||bobjLinks.fids.get(fid);
    const sole=choices?.size===1?choices.keys().next().value:'';
    const quote=matchedCtl?.quote||(inputFid?sole:bobjQuoteKey(bobjValue(fid,'Active Quote Number'))||sole);
    if(quote&&(inputFid||matchedCtl?.quote===quote||!bobjLinks.quotes.has(quote)||bobjLinks.quotes.get(quote)?.fid===fid))row.quote=quote;
  }
  row.sapVersion=zbomSapVersions.get(row.quote);
  const sapFid=bobjFidKey(row.sapVersion);
  const linked=zbomQuoteFid(row.quote);
  const sourceFid=matchedCtl?.fid||linked;
  if(!inputFid&&sourceFid)row.fid=sourceFid;
  if(matchedCtl)row.ctl=ctl;
  if(!row.quote&&(fid||ctl)){
    if(bobjLinks.ctls.has(ctl)&&!ctlLink)row._error='CTL ID matches multiple records. Enter the quotation.';
    else if(row.quotePicker)row._error='Choose a quotation or remove the CTL ID.';
    else row._error=bobjLinks.fids.has(fid)?'Enter the quotation for this FID.':'No quotation found for this '+(ctl?'CTL ID.':'FID.');
  }else if(ctlLink&&!matchedCtl)row._error='CTL ID does not match the entered FID or quotation.';
  else if(row.quote&&!bobjFidKey(row.fid)){
    row._afterExport=row.sapVersion!==undefined;
    row._lookup=row._afterExport?'FID will be matched after export.':'FID will be checked in SAP.';
  }
  else if(linked&&linked!==bobjFidKey(row.fid))row._error=(sapFid?'SAP links':'The data sources link')+' this quotation to FID '+linked+'.';
  return row;
}

function hasBobjFid(fid){return bwFidRows.has(fid)||bobjDataRows.has(fid);}

function needsZbomSapLookup(row,fields,edited){
  return !row._error&&row.quote&&(toolboxLatestQuoteFid(row.quote)||row.sapVersion===undefined&&
    (!row.fid||row._lookup||!edited&&(!bobjTemplateValue(row,'System Description')||bobjPreviewIssue(row,fields))));
}

async function fetchToolboxLookup(tool,lookupTool,rows){
  const health=await bridgeRequest('GET','/health',undefined,8000);
  if(health.mode!=='mock'&&!health.sessions?.some(session=>session.user))throw new Error('Sign in to SAP before looking up documents.');
  const id=crypto.randomUUID().replace(/-/g,''),route='/toolbox/tasks/'+id;
  const spec={Id:id,Tool:lookupTool,Workflow:createToolboxWorkflow(lookupTool),Rows:rows};
  let task;
  try{({task}=await bridgeRequest('POST','/toolbox/tasks',spec,30000))}
  catch(error){
    if(!error.transport)throw error;
    // The bridge may have accepted the lookup before the response was lost.
    ({task}=await bridgeRequest('GET',route,undefined,10000));
  }
  if(task?.id!==id||task.tool!==lookupTool)throw new Error('SAPGUIcontrol did not confirm the lookup.');
  try{
    const queueDeadline=Date.now()+60000;
    while(!toolboxTaskFinished(task)||!task.finishedAt||task.canDelete===false){
      if(['Queued','Waiting for SAP window'].includes(task.state)&&Date.now()>=queueDeadline){
        const message=task.state==='Queued'?'SAP did not make a window available for lookup.':toolboxTaskPhase(task);
        ({task}=await bridgeRequest('POST',route+'/cancel',{},10000));
        throw new Error(message);
      }
      setSapProgress(tool,task.phase==='Queued'?'Waiting for a SAP lookup window…':toolboxTaskPhase(task));
      await waitMs(750);
      ({task}=await bridgeRequest('GET',route,undefined,10000));
    }
    return task;
  }finally{
    if(task?.canDelete&&!task.window)try{await bridgeRequest('POST',route+'/delete',{},10000)}catch(_){/* The bridge can expire the finished lookup. */}
  }
}

function zbomLookupPending(tool,row){
  return row.quote&&!row._error&&!zbomSapQuotes[tool].has(bobjFidKey(row.fid));
}

async function lookupZbomSap(tool){
  const pdf=tool==='zbom',preview=pdf?previewZbomInput:previewZbomCompareInput,targets=new Map();
  for(const row of pdf?zbomRows:zbomCompareRows)if(zbomLookupPending(tool,row)){
    const fid=row.inputFid||bobjFidKey(row.sapVersion),quote=fid?'sap-latest:'+fid:row.quote;
    targets.set(quote,{Quote:quote,Fid:fid||row.fid});
  }
  if(!targets.size)return;
  const task=await fetchToolboxLookup(tool,ZBOM_SAP_LOOKUP_TOOL,[...targets.values()]),misses=[];
  for(const {Quote,Fid,Status,Message} of task.results||[]){
    if(Status!=='Completed')continue;
    const fid=bobjFidKey(Fid),expected=toolboxLatestQuoteFid(Quote);
    const found=JSON.parse(Message),quotes=[];
    if(fid&&(!expected||fid===expected)){
      if(!expected)zbomSapVersions.set(Quote,Fid);
      for(const [quote,version] of found)if(/^\d{8}$/.test(quote)&&isZbomFid(version)&&bobjFidKey(version)===fid){
        zbomSapVersions.set(quote,version);quotes.push(quote);
      }
    }
    if(quotes.length)zbomSapQuotes[tool].set(fid,quotes.sort((a,b)=>b.localeCompare(a)));
    else misses.push(fid||Quote);
  }
  preview();if(await ensureBobjDataSource())preview();
  if(task.state!=='Completed')throw new Error(task.error||toolboxTaskPhase(task)||'SAP lookup did not complete.');
  if(misses.length)throw new Error('No matching SAP quotations for '+misses.join(', ')+'.');
  setSapProgress(tool,'SAP lookup finished.','ok');
}

function bobjTemplateValue(row,header,omit){
  const key=normalizeBobjHeader(header),options=omit||{};
  if(key==='quote')return toolboxLatestQuoteFid(row?.quote)?'{Latest}':String(row?.quote||'');
  if(key==='fid')return bobjFidKey(row?.fid);
  if(key==='fidver')return row?.sapVersion?row.sapVersion.slice(6):'{FIDver}';
  if(key==='today')return todayZbomStamp();
  if(options.customer&&key==='customer')return '';
  if(options.systemDescription&&key==='system description')return '';
  const source=bobjValue(row&&row.fid,header);if(source)return source;
  const field=key==='customer'?'customer':key==='system description'?'systemDescription':'';
  if(!field)return '';
  const fid=bobjFidKey(row.fid),ctlLink=bobjLinks.ctls.get(row.ctl),quoteLink=bobjLinks.quotes.get(row.quote);
  return row[field]||(ctlLink?.fid===fid&&(!ctlLink.quote||ctlLink.quote===row.quote)?ctlLink[field]:'')||(quoteLink?.fid===fid?quoteLink[field]:'')||'';
}

function renderBobjTemplate(template,row,omit,pendingFid){
  return String(template||'').replace(/\{([^{}]+)\}/g,(token,header)=>{
    const value=bobjTemplateValue(row,header,omit);
    if(value)return value;
    if(pendingFid)return pendingFidFilenameToken(token,header,omit);
    return normalizeBobjHeader(header)==='system description'&&!omit?.systemDescription?'{System Description}':'';
  });
}

function bobjTemplateFields(template){
  return Array.from(String(template).matchAll(/\{([^{}]+)\}/g),match=>normalizeBobjHeader(match[1])).filter(field=>!['fid','fidver','quote','today'].includes(field));
}

function bobjPreviewIssue(row,fields){
  if(row._error||row._lookup)return row._lookup||'';
  const missing=fields.filter(field=>!bobjTemplateValue(row,field));
  return missing.length?'Filename fields unavailable: '+missing.join(', ')+'.':'';
}

async function waitForBobjRefresh(endpoint,status,update,onProgress,isCurrent=()=>true){
  let pendingStart=update&&!status.running,blocked=false;
  while(pendingStart||status.running){
    if(!isCurrent())return null;
    if(pendingStart&&status.running)pendingStart=false;
    if(pendingStart){
      if(blocked){
        const other=await bridgeRequest('GET',endpoint==='/excel'?'/excel/ctl/status':'/excel/status',undefined,10000);
        if(!isCurrent())return null;blocked=Boolean(other.running);
      }
      if(!blocked){
        try{status=await bridgeRequest('POST',endpoint+'/refresh',{},20000);pendingStart=false;continue}
        catch(error){if(error.status!==409)throw error;blocked=true}
      }
    }
    onProgress(status,pendingStart||status.stage==='queued');await waitMs(1000);if(!isCurrent())return null;
    status=await bridgeRequest('GET',endpoint+'/status',undefined,10000);
  }
  return isCurrent()?status:null;
}

function bobjWorkerMain(){
  'use strict';
  let rows=new Map(); const normalizeHeader=value=>String(value||'').trim().replace(/\s+/g,' ').toLocaleLowerCase(); const sourceFid=value=>{const text=String(value||'').trim();return /^\d{6}$/.test(text)?text:''};
  const base64Bytes=value=>{const binary=atob(String(value||'')),bytes=new Uint8Array(binary.length);for(let index=0;index<binary.length;index++)bytes[index]=binary.charCodeAt(index);return bytes};
  const findTable=(workbook,firstMatch)=>{
    let best=null;
    (workbook&&workbook.SheetNames||[]).forEach((sheetName,sheetIndex)=>{
      const sheet=workbook.Sheets[sheetName],reference=sheet&&sheet['!ref'];if(!reference)return;
      const range=XLSX.utils.decode_range(reference),options={header:1,defval:'',raw:false,dateNF:'mm/dd/yyyy',blankrows:true},scanRange={s:{r:range.s.r,c:range.s.c},e:{r:Math.min(range.e.r,range.s.r+49),c:range.e.c}},preview=XLSX.utils.sheet_to_json(sheet,Object.assign({},options,
        {range:XLSX.utils.encode_range(scanRange)})),headerOffset=preview.findIndex(row=>normalizeHeader(row&&row[0])==='fid');if(headerOffset<0)return;
      const dataRange={s:{r:range.s.r+headerOffset,c:range.s.c},e:{r:range.e.r,c:range.e.c}},matrix=XLSX.utils.sheet_to_json(sheet,Object.assign({},options,{range:XLSX.utils.encode_range(dataRange)}));
      if(firstMatch){best={matrix};return}
      const fidCount=matrix.reduce((count,row,index)=>count+(index&&sourceFid(row?.[0])?1:0),0);if(!best||fidCount>best.fidCount)best={matrix,fidCount};
    });
    return best;
  };
  self.onmessage=event=>{
    const message=event.data||{},reply=data=>self.postMessage(Object.assign({id:message.id,ok:true},data));
    try{
      if(message.type==='export'){
        const workbook=XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([message.columns,...message.rows],{dense:true}),message.sheet);
        reply({contentBase64:XLSX.write(workbook,{bookType:'xlsx',type:'base64',compression:true})});return;
      }
      if(message.type==='load'){
        const options={type:'array',cellDates:true,dense:true,sheets:['BOBJ'],cellFormula:false,cellHTML:false,cellStyles:false},bytes=base64Bytes(message.contentBase64);
        let workbook=XLSX.read(bytes,options),table=findTable(workbook,true);if(!table){delete options.sheets;workbook=XLSX.read(bytes,options);table=findTable(workbook,false)}
        if(!table)throw new Error('No worksheet with FID in its first column was found.');
        const headers=(table.matrix[0]||[]).map(value=>String(value||'').trim()),headerIndexes=new Map();headers.forEach((header,index)=>{const key=normalizeHeader(header);if(key&&!headerIndexes.has(key))headerIndexes.set(key,index)});
        const nextRows=new Map();
        for(let index=1;index<table.matrix.length;index++){const row=table.matrix[index],fid=sourceFid(row?.[0]);if(fid&&!nextRows.has(fid))nextRows.set(fid,row)}
        if(!nextRows.size)throw new Error('The FID BOBJ worksheet contains no readable FID rows.');
        rows=nextRows;reply({headerIndexes:Array.from(headerIndexes)});return;
      }
      if(message.type==='lookup'){const found=[];(message.fids||[]).forEach(fid=>{const row=rows.get(fid);if(row)found.push([fid,row])});reply({rows:found});return;}
      if(message.type==='project'){
        reply({rows:Array.from(rows,([fid,row])=>[fid,message.indexes.map(index=>row[index]??'')])});return;
      }
      throw new Error('Unknown FID BOBJ worker request.');
    }catch(error){self.postMessage({id:message.id,ok:false,error:error&&error.message?error.message:String(error)})}
  };
}

function createBobjWorker(){
  if(bobjWorker)return bobjWorker;
  if(typeof Worker==='undefined')throw new Error('This browser does not support background workbook reading.');
  const worker=createWorkbookWorker('('+bobjWorkerMain.toString()+')()');
  worker.onmessage=event=>{const data=event.data||{},pending=bobjWorkerPending.get(data.id);if(!pending)return;bobjWorkerPending.delete(data.id);if(data.ok)pending.resolve(data);else pending.reject(new Error(data.error||'BOBJ background worker failed.'))};
  worker.onerror=event=>{const error=new Error(event&&event.message||'BOBJ background worker failed.');bobjWorkerPending.forEach(pending=>pending.reject(error));bobjWorkerPending.clear();worker.terminate();if(bobjWorker===worker){bobjWorker=null;bobjIndexedFile=null}}; bobjWorker=worker;return worker;
}

function callBobjWorker(type,payload){return new Promise((resolve,reject)=>{const id=++bobjWorkerRequest;bobjWorkerPending.set(id,{resolve:resolve,reject:reject});try{createBobjWorker().postMessage(Object.assign({id:id,type:type},payload||{}))}
  catch(error){bobjWorkerPending.delete(id);reject(error)}});}

function refreshBobjConsumers(){
  if(!sapToolRunning('zbom'))previewZbomInput();
  if(!sapToolRunning('zbomCompare'))previewZbomCompareInput();
  if(!sapToolRunning('createQuote'))previewCreateQuoteInput();
  if(!sapToolRunning('partChanges'))renderPartChangesPreview();
}

async function loadBobjFileIndex(source){
  const data=await bridgeFileRequest('read',{path:source.path});
  if(!data?.contentBase64)throw new Error('SAPGUIcontrol returned no FID BOBJ workbook content.');
  const loading=callBobjWorker('load',{contentBase64:data.contentBase64});data.contentBase64='';
  const loaded=await loading;bobjIndexedFile=source;bobjDataRows.clear();bobjDataMissing.clear();bobjDataHeaderIndexes=new Map(loaded.headerIndexes);
}

async function ensureBobjFileIndex(source){
  if(bobjDataPromise)await bobjDataPromise;
  if(source!==bobjFileSource)return false;
  if(source!==bobjIndexedFile){
    bobjDataPromise=loadBobjFileIndex(source).finally(()=>{bobjDataPromise=null});
    await bobjDataPromise;
  }
  return source===bobjFileSource&&source===bobjIndexedFile;
}

async function snapshotExportFidSource(template,extraFields=[]){
  // Store only the attributes the task can display or use in its filenames.
  await toolboxCacheReady;
  const fields=[...new Set(['customer','system description',...bobjTemplateFields(template),...extraFields])];
  const source=bobjFileSource;
  let fileRows=new Map();
  if(source)try{
    if(await ensureBobjFileIndex(source)){
      const indexes=fields.map(field=>bobjDataHeaderIndexes.get(field)??-1);
      if(indexes.some(index=>index>=0)){
        const {rows}=await callBobjWorker('project',{indexes});
        if(source===bobjFileSource&&source===bobjIndexedFile)fileRows=new Map(rows);
      }
    }
  }catch(error){setBobjFileStatus('fcid','BOBJ file read failed: '+bridgeError(error),'error')}
  const rows=Object.create(null);
  for(const fid of new Set([...bwFidRows.keys(),...fileRows.keys()])){
    const values=fields.map((field,index)=>bobjField(bwFidRows.get(fid),bwFidHeaders.get(field))||bobjField(fileRows.get(fid),index));
    if(values.some(Boolean))rows[fid]=values;
  }
  return Object.keys(rows).length?JSON.stringify([fields,rows]):'';
}

async function loadBobjRows(fids,fields){
  const source=bobjFileSource;
  if(!source)return false;
  const needed=[...new Set(Array.from(fids,bobjFidKey))].filter(fid=>fid&&(!fields||fields.some(field=>!bobjField(bwFidRows.get(fid),bwFidHeaders.get(field)))));
  if(!needed.length)return false;
  try{
    const reindex=bobjIndexedFile!==source;
    if(reindex&&!(await ensureBobjFileIndex(source)))return false;
    const missing=needed.filter(fid=>!bobjDataRows.has(fid)&&!bobjDataMissing.has(fid));if(!missing.length)return reindex;
    const {rows}=await callBobjWorker('lookup',{fids:missing});
    if(source!==bobjFileSource||source!==bobjIndexedFile)return false;
    for(const [fid,row] of rows)bobjDataRows.set(fid,row);
    for(const fid of missing)if(!bobjDataRows.has(fid))bobjDataMissing.add(fid);
    return reindex||rows.length>0;
  }catch(error){setBobjFileStatus('fcid','BOBJ file read failed: '+bridgeError(error),'error');return false}
}

async function ensureBobjDataSource(){
  await toolboxCacheReady;
  if(!bobjFileSource)return;
  const needed=new Set(),find=(rows,fields)=>{
    for(const row of rows){
      const fid=bobjFidKey(row.fid||row.Fid),cached=bwFidRows.get(fid);
      if(fid&&fields.some(field=>!bobjField(cached,bwFidHeaders.get(field))))needed.add(fid);
    }
  };
  find(zbomRows,['system description',...bobjTemplateFields(bridgeId('zbomTemplate').value||ZBOM_DEFAULT_TEMPLATE)]);
  find(zbomCompareRows,['customer','system description',...bobjTemplateFields(bridgeId('zbomCompareTemplate').value||ZBOM_COMPARE_DEFAULT_TEMPLATE)]);
  find(createQuoteRows,['date: customer request']);
  find(partChangesRows.filter(row=>!row.error),['system description',...bobjTemplateFields(bridgeId('partChangesTemplate').value||PART_CHANGES_DEFAULT_TEMPLATE).filter(field=>field!=='so')]);
  const loaded=await loadBobjRows(needed);if(loaded)refreshBobjConsumers();return loaded;
}

function bobjSourceCard(key,{id,label},number){
  const actions=`<div class="fc-source-toolbar"><button id="${id}FileUpdate" class="btn light small" type="button" disabled>Update BOBJ Excel</button>
    <button id="${id}Open" class="btn light small" type="button" disabled>Open Excel</button></div>`;
  return `<section class="fc-source" data-source="${key}" aria-label="${label}">
    <div class="fc-source-head"><h3><span class="fc-source-number">${number}</span><span class="fc-source-name">${label}</span></h3><span id="${id}RefreshState" class="fc-source-state" role="status">Checking…</span></div>
    <div class="fc-source-toolbar"><span id="${id}DataSource" class="fc-source-file"></span>
      <button id="${id}Update" class="btn small light" type="button" disabled>Load</button>
      <button id="${id}Import" class="btn small light" type="button" disabled>Import</button>
      ${sourceMenu(id+'Menu',label+' file actions',actions)}
    </div>
    <input id="${id}ImportFile" type="file" aria-label="Import ${label}" accept=".xlsx,.xls,.xlsm,.csv" hidden>
    <span id="${id}FileState" class="bridge-note bobj-refresh-state" role="status" hidden></span>
  </section>`;
}

function setToolboxSourceStatus(key,message,kind='ok'){
  const {id,meta}=bobjCards[key];paintSourceStatus(bridgeId(id+'RefreshState'),meta,bwCache,message,kind);
}

function setBobjFileStatus(key,message,kind='',details=message){
  const note=bridgeId(bobjCards[key].id+'FileState');note.textContent=message;note.title=details;note.hidden=!message;note.className='bridge-note bobj-refresh-state '+kind;
}

function applyToolboxSource(key,cache,cached){
  const {meta,rows,columns}=cache,card=bobjCards[key];
  if(!cached||!meta.dateField)Object.assign(meta,sourceDateRange(rows,key,bwCache,columns));
  card.meta={...meta,cached};paintSourceFile(bridgeId(card.id+'DataSource'),meta,false);setToolboxSourceStatus(key,'Loaded');
  if(key==='fcid'){
    const headers=new Map(),indexed=new Map();
    columns.forEach((column,index)=>{const field=normalizeBobjHeader(column);if(!headers.has(field))headers.set(field,index)});
    const fidIndex=columns.indexOf('FID'),quoteIndex=headers.get('active quote number');
    bobjActiveQuotes.clear();
    for(const row of rows){
      const fid=bobjField(row,fidIndex);if(!/^\d{6}$/.test(fid))continue;
      if(!indexed.has(fid))indexed.set(fid,row);
      if(quoteIndex!==undefined){
        const quote=bobjQuoteKey(row[quoteIndex]);if(quote)indexUniqueBobjValue(bobjActiveQuotes,quote,fid);
      }
    }
    bwFidHeaders=headers;bwFidRows=indexed;
  }else{Object.values(bobjLinks).forEach(map=>map.clear());indexBobjQuotes(rows,columns)}
  refreshBobjConsumers();ensureBobjDataSource();
}

async function loadToolboxSource(key,file,range){
  const card=bobjCards[key];
  if(card.dataBusy)return;
  card.dataBusy=true;setBridgeBusy(bridgeBusy);
  setToolboxSourceStatus(key,file?'Reading '+file.name+'…':'Loading BW InA / 0%','loading');
  try{
    let data;
    if(file){
      const {rows,...recognition}=await toolboxSourceReader.readFile(file,key);
      if(!rows.length)throw new Error('The file contains no data rows.');
      data={rows,meta:{...recognition,file:file.name,kind:'file',savedAt:Date.now()}};
    }else{
      data=await readBwSnapshot(bwData,key,undefined,message=>setToolboxSourceStatus(key,message,'loading'),range||(card.meta?.customRange?card.meta:undefined));
    }
    const {rows,meta:{fields,...meta},warning,cleanupError}=data,notice=[warning,cleanupError].filter(Boolean).join(' ');
    const cache=packSourceCache(BW_CACHE_PREFIX+key,rows,meta,fields);
    applyToolboxSource(key,cache,false);
    try{await bwCache.writeDataset(cache)}catch(error){setToolboxSourceStatus(key,'Cache not saved: '+error.message,'error');return}
    if(notice)setToolboxSourceStatus(key,notice,'error');
  }catch(error){
    const message=(file?'Import failed: ':'BW load failed: ')+(error.userMessage||error.message);
    if(file)setToolboxSourceStatus(key,message+(card.meta?' Previous data retained.':''),'error');
    else{
      if(card.meta)Object.assign(card.meta,{cached:true,cacheReason:message});
      setToolboxSourceStatus(key,message,card.meta?'ok':'error');
    }
  }finally{card.dataBusy=false;setBridgeBusy(bridgeBusy)}
}

async function restoreToolboxSources(){
  const keys=Object.keys(bobjCards).filter(key=>!bobjCards[key].dataBusy),savedDay=key=>bwCache.localDate(new Date(bobjCards[key].meta.savedAt)),pending=[];
  for(const key of keys)bobjCards[key].dataBusy=true;
  updateBobjFileControls();
  try{
    toolboxCacheReady=Promise.all(keys.map(async key=>{
      const cache=await bwCache.readDataset(BW_CACHE_PREFIX+key).catch(()=>null);
      if(validSourceCache(cache,BW_SOURCES[key].ids,true))applyToolboxSource(key,cache,true);
    }));
    await toolboxCacheReady;
    const today=bwCache.localDate(new Date()),expired=keys.filter(key=>bobjCards[key].meta&&savedDay(key)!==today);
    pending.push(...keys.filter(key=>!bobjCards[key].meta));
    if(expired.length){
      await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
      if(confirmBwRefresh(expired,savedDay))pending.push(...expired);
    }
  }finally{
    for(const key of keys)bobjCards[key].dataBusy=false;
    updateBobjFileControls();
  }
  await Promise.all(pending.map(key=>loadToolboxSource(key)));
}

async function refreshToolbox(){
  if(toolboxRefreshPending)return;
  toolboxRefreshPending=true;setBridgeBusy(bridgeBusy);
  try{await Promise.all([restoreToolboxSources(),refreshBridge()])}
  finally{toolboxRefreshPending=false;setBridgeBusy(bridgeBusy)}
}

function updateBobjFileControls(){
  for(const {id,meta,dataBusy,fileStatus,fileBusy} of Object.values(bobjCards)){
    const update=bridgeId(id+'FileUpdate'),open=bridgeId(id+'Open'),busy=bridgeBusy||fileBusy||sapGuiControlUpdate.online!==true;
    for(const suffix of ['Update','Import','ImportFile'])bridgeId(id+suffix).disabled=Boolean(dataBusy);
    update.disabled=busy||fileStatus?.rfcReady!==true;open.disabled=busy||dataBusy||!meta||!fileStatus?.outputPath;
    update.title=fileStatus?.rfcReady?'Refresh this workbook through RFC/BICS.':fileStatus?.componentMessage||'SAPGUIcontrol with RFC/BICS support is required.';
    open.title=meta?'Save cached data as Excel and open it using SAPGUIcontrol.':'No cached data available.';
  }
}

async function readBobjFileSource(status,force=false){
  const path=status.hasWorkbook&&status.outputPath;
  if(!path){bobjFileSource=null;return}
  const data=await bridgeFileRequest('stat',{path},10000),entry=data.entries?.[0];
  if(!entry){bobjFileSource=null;return}
  const modified=Number(entry.lastModified)||0;
  if(force||path!==bobjFileSource?.path||modified!==bobjFileSource?.modified){bobjFileSource={path,modified};ensureBobjDataSource()}
}

async function refreshBobjFile(key,update=false){
  const card=bobjCards[key],{id,endpoint}=card;
  if(card.fileBusy)return;
  card.fileBusy=true;updateBobjFileControls();
  if(update){bridgeId(id+'Menu').open=false;setBobjFileStatus(key,'Starting BOBJ file update…','running')}
  try{
    let status=card.fileStatus=await bridgeRequest('GET',endpoint+'/status',undefined,10000);
    const running=status.running;
    if(key==='fcid'&&running&&!bobjFileSource)await readBobjFileSource(status);
    if(update&&!running&&status.rfcReady!==true)throw new Error(status.componentMessage||'SAPGUIcontrol with RFC/BICS support is required.');
    status=await waitForBobjRefresh(endpoint,status,update,(state,queued)=>{
      setBobjFileStatus(key,queued?'Waiting for another BOBJ file update…':'Updating BOBJ file / '+(state.stage||'running'),'running',state.message);
    });
    card.fileStatus=status;
    if(key==='fcid')await readBobjFileSource(status,(update||running)&&status.stage==='completed');
    if(status.stage==='failed')throw new Error(status.message||'RFC/BICS update failed.');
    if(update||running)setBobjFileStatus(key,status.hasWorkbook?'BOBJ file updated.':status.message||'No BOBJ file is available.');
  }catch(error){setBobjFileStatus(key,'BOBJ file: '+bridgeError(error),'error')}
  finally{card.fileBusy=false;updateBobjFileControls()}
}

async function openBobjData(key){
  const card=bobjCards[key],{id,fileStatus}=card;
  if(bridgeId(id+'Open').disabled)return;
  card.fileBusy=true;bridgeId(id+'Menu').open=false;updateBobjFileControls();
  setBobjFileStatus(key,'Preparing Excel…','running');
  try{
    const cache=await bwCache.readDataset(BW_CACHE_PREFIX+key);
    if(!validSourceCache(cache,BW_SOURCES[key].ids,true))throw new Error('No cached data available. Use Load first.');
    const {contentBase64}=await callBobjWorker('export',{columns:cache.columns,rows:cache.rows,sheet:key==='fcid'?'BOBJ':'CTL BOBJ'});
    const path=fileStatus.outputPath.replace(/\.xlsm$/i,'.xlsx');
    await bridgeFileRequest('write',{path,contentBase64,overwrite:true,createParents:true});
    await bridgeFileRequest('openFile',{path},30000);setBobjFileStatus(key,'Cache opened in Excel.');
  }
  catch(error){setBobjFileStatus(key,'Open failed: '+bridgeError(error),'error')}
  finally{card.fileBusy=false;updateBobjFileControls()}
}
function sapToolRunning(tool){return Boolean(sapTools[tool].starting)}
function toolboxSourcesLoading(){return toolboxRefreshPending||Object.values(bobjCards).some(card=>card.dataBusy)}
function toolboxWorkflowReady(tool){
  return sapGuiControlUpdate.workflowReady&&(tool!=='zbom'&&tool!=='zbomCompare'||sapGuiControlUpdate.sapBatch)&&
    (tool!=='zbomCompare'||sapGuiControlUpdate.compareCalculation)&&(tool!=='partChanges'||sapGuiControlUpdate.watchFileCount);
}
function canStartSapTool(tool,workflow=tool){return bridgeSupportsToolboxTasks&&toolboxWorkflowReady(workflow)&&bridgeTaskOnline&&!bridgeBusy&&!sapToolRunning(tool)&&!toolboxSourcesLoading()}

function setBridgeBusy(value){
  bridgeBusy=Boolean(value);
  document.querySelectorAll('[data-bridge-action]').forEach(button=>{const panel=button.closest('[data-tool-panel]'),tool=panel&&panel.dataset.toolPanel; button.disabled=bridgeBusy||Boolean(tool&&sapToolRunning(tool));});
  document.querySelectorAll('[data-tool-panel] input,[data-tool-panel] textarea,[data-tool-panel] select').forEach(input=>{
    const tool=input.closest('[data-tool-panel]').dataset.toolPanel,disabled=sapToolRunning(tool)||input.hasAttribute('data-preview-disabled');
    if(input.disabled===disabled)return;
    input.disabled=disabled;
    if(input.matches('.part-changes-system-select')&&typeof jQuery==='function')jQuery(input).trigger('chosen:updated');
  });
  document.querySelectorAll('[data-zbom-token],[data-zbom-compare-token],.zbom-file-edit,.zbom-file-save,.zbom-file-reset,#zbomResetInput,#zbomResetTemplate,#zbomCompareResetInput,#zbomCompareResetTemplate,#zbomCompareKeepTemp').forEach(button=>{button.disabled=
    sapToolRunning(button.closest('[data-tool-panel]').dataset.toolPanel)||button.hasAttribute('data-preview-disabled');});
  bridgeId('bridgeRefresh').disabled=bridgeBusy||toolboxRefreshPending;
  updateZbomStart();updateZbomCompareStart();updatePartChangesControls();updateCreateQuoteStart();updateBobjFileControls();
}

async function pollSapTasks(){
  if(bridgeTaskPolling||!bridgeSupportsToolboxTasks)return;
  bridgeTaskPolling=true;clearTimeout(bridgeTaskPollTimer);
  try{await pollToolboxTasks();bridgeTaskOnline=true} catch(_){bridgeTaskOnline=false} finally{bridgeTaskPolling=false;setBridgeBusy(bridgeBusy);bridgeTaskPollTimer=setTimeout(pollSapTasks,3000)}
}

async function runSapTool(tool,run){
  if(sapToolRunning(tool))return;
  sapTools[tool].starting=true;setBridgeBusy(bridgeBusy);
  setSapProgress(tool,'Preparing input…');
  try{await run()}
  catch(error){setSapProgress(tool,bridgeError(error),'error')}
  finally{
    if(sapTools[tool].phase==='Preparing input…')setSapProgress(tool,'Ready.','idle');
    sapTools[tool].starting=false;setBridgeBusy(bridgeBusy);
  }
}

function isSapGuiControlUpdateAvailable(current,latest){const local=current.replace(/^v/i,'').split('.'),remote=latest.replace(/^v/i,'').split('.'); for(let index=0;index<Math.max(local.length,remote.length);index++){const difference=Number(remote[index]||0)-Number(local[index]||0);
  if(difference)return difference>0;} return false;}

function renderSapGuiControlUpdate(){
  const state=sapGuiControlUpdate,release=state.release,download=bridgeId('bridgeDownload'); const newer=Boolean(release&&state.current&&isSapGuiControlUpdateAvailable(state.current,release.version))||!toolboxWorkflowReady('zbomCompare')||!toolboxWorkflowReady('partChanges')||!state.rowBatches||!state.partChangesMyLam; download.href=release&&release.downloadUrl||SAPGUICONTROL_DOWNLOAD_URL;
  download.hidden=state.online===null||Boolean(state.online&&!newer);download.textContent=state.online?'Update':'Download';
  download.title=(release?'Download SAPGUIcontrol '+release.version+' (ZIP).':'Download SAPGUIcontrol.zip.')+' Extract it and run SAPGUIcontrol.exe.';
  const service=bridgeId('bridgeService');
  service.textContent=state.online===null?'':state.online?'Version '+state.current:release?'Latest '+release.version:'';
  service.hidden=!service.textContent;
}

async function checkSapGuiControlUpdate(){
  renderSapGuiControlUpdate();
  try{const release=await bridgeRequest('GET',SAPGUICONTROL_VERSION_URL+'?t='+Date.now(),undefined,8000); sapGuiControlUpdate.release={version:release.version.trim(),downloadUrl:release.downloadUrl};}catch(_){sapGuiControlUpdate.release=null}
  finally{renderSapGuiControlUpdate()}
}

async function refreshBridge(background=false){
  if(bridgeRefreshPending)return;
  bridgeRefreshPending=true;clearTimeout(bridgeReconnectTimer);bridgeReconnectTimer=0;
  if(background!==true){setBridgeBusy(true);setBridgeBadge('Checking','');sapGuiControlUpdate.online=null;checkSapGuiControlUpdate()}
  try{
    const health=await bridgeRequest('GET','/health',undefined,background===true?3000:8000),mock=health.mode==='mock';
    if(background===true)setBridgeBusy(true);
    bridgeSupportsToolboxTasks=health.toolboxTasks===true;
    const workflow=health.toolboxWorkflow;
    sapGuiControlUpdate.workflowReady=bridgeSupportsToolboxTasks&&workflow?.api===1&&workflow.afterSuccess===true&&
      Array.isArray(workflow.operations)&&toolboxWorkflowOperations().every(operation=>workflow.operations.includes(operation));
    sapGuiControlUpdate.sapBatch=sapGuiControlUpdate.workflowReady&&workflow.operations.includes('sapBatch');
    sapGuiControlUpdate.compareCalculation=sapGuiControlUpdate.workflowReady&&['alignRows','scriptStage'].every(operation=>workflow.operations.includes(operation));
    sapGuiControlUpdate.rowBatches=workflow?.rowBatches===true;
    sapGuiControlUpdate.partChangesMyLam=workflow?.partChangesMyLam===true;
    sapGuiControlUpdate.watchFileCount=sapGuiControlUpdate.workflowReady&&workflow.operations.includes('watchFileCount');
    sapGuiControlUpdate.mock=mock;
    await pollSapTasks();sapGuiControlUpdate.online=true;sapGuiControlUpdate.current=health.version||'';
    renderSapGuiControlUpdate();renderToolboxTasks();
    if(!sapGuiControlUpdate.workflowReady)setBridgeBadge('Update needed','warn','Update SAPGUIcontrol for Tasks.');
    else if(!sapGuiControlUpdate.sapBatch)setBridgeBadge('Update needed','warn','Update SAPGUIcontrol for parallel ZBOM tasks.');
    else if(!sapGuiControlUpdate.compareCalculation)setBridgeBadge('Update needed','warn','Update SAPGUIcontrol for ZBOM Compare.');
    else if(!sapGuiControlUpdate.watchFileCount)setBridgeBadge('Update needed','warn','Update SAPGUIcontrol for BOM Compare.');
    else if(!mock&&health.sapAvailable===false)setBridgeBadge('SAP unavailable','warn');
    else setBridgeBadge(mock?'Mock ready':'Ready',mock?'warn':'ok');
    for(const key of Object.keys(bobjCards))refreshBobjFile(key);
  }catch(error){
    if(background===true)return;
    bridgeSupportsToolboxTasks=false;sapGuiControlUpdate.workflowReady=false;sapGuiControlUpdate.sapBatch=false;sapGuiControlUpdate.compareCalculation=false;sapGuiControlUpdate.rowBatches=false;sapGuiControlUpdate.partChangesMyLam=false;sapGuiControlUpdate.watchFileCount=false;bridgeTaskOnline=false;
    setBridgeBadge(error.status?'Error':'Offline','error',bridgeError(error));sapGuiControlUpdate.online=false;sapGuiControlUpdate.current='';renderSapGuiControlUpdate();
  }finally{bridgeRefreshPending=false;setBridgeBusy(false);if(sapGuiControlUpdate.online===false)bridgeReconnectTimer=setTimeout(()=>refreshBridge(true),10000);}
}

function setToolTab(name){
  document.querySelectorAll('[data-tool-tab]').forEach(button=>{const active=button.dataset.toolTab===name;button.setAttribute('aria-selected',active?'true':'false');});
  document.querySelectorAll('[data-tool-panel]').forEach(panel=>{panel.hidden=panel.dataset.toolPanel!==name});setBridgeBusy(bridgeBusy);
}

function splitZbomLine(line){
  const separator=['\t',',',';','，','；','、'].find(separator=>line.includes(separator));
  if(separator)return splitQuotedLine(line,separator).map(value=>value.trim());
  if(/\s{2,}/.test(line.trim()))return line.trim().split(/\s{2,}/).map(value=>value.trim());
  return [line.trim()];
}

function normalizedZbomHeader(value){return String(value||'').normalize('NFKC').trim().toLowerCase().replace(/[\s_.\-/#:]+/g,'');}

function zbomHeaderRole(value){
  const key=normalizedZbomHeader(value);
  if(['fid','fidno','fidnumber','facilityid','opportunityid'].includes(key))return 'fid';
  if(['ctl','ctlid','ctlno','ctlnumber'].includes(key))return 'ctl';
  if(['systemdescription','systemdesc','sysdescription','sysdesc','description'].includes(key))return 'systemDescription';
  if(['quote','quotation','quoteid','quoteno','quotationno','quotenumber','quotationnumber','vbeln','报价','报价号','报价单','报价单号'].includes(key))return 'quote';
  if(['id','ids','identifier','number','fidctlquote','fidctlquoteid','fidctlquotation'].includes(key))return 'identifier';
  return '';
}

function createQuoteHeaderRole(value){
  const key=normalizedZbomHeader(value);
  if(['fid','facilityid'].includes(key))return 'fid';
  if(['crd','customerrequestdate','customerrequesteddate','datecustomerrequest'].includes(key))return 'crd';
  if(['validto','validateto','validityto','validitydate'].includes(key))return 'validTo';
  return '';
}

function createQuoteUtcDate(year,month,day){const value=Date.UTC(Number(year),Number(month)-1,Number(day)),date=new Date(value); return date.getUTCFullYear()===Number(year)&&date.getUTCMonth()===Number(month)-1&&date.getUTCDate()===Number(day)?value:null;}

function createQuoteDateValue(value){
  const text=String(value||'').trim();if(!text||/^\d{6}$/.test(text))return null;
  let match=text.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if(match)return createQuoteUtcDate(match[1],match[2],match[3]);
  match=text.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if(match)return createQuoteUtcDate(match[1],match[2],match[3]);
  match=text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if(match)return createQuoteUtcDate(match[1],match[2],match[3]);
  match=text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if(match){const first=Number(match[1]),second=Number(match[2]),month=first>12?second:first,day=first>12?first:second; return createQuoteUtcDate(match[3],month,day);}
  if(/^\d{4,5}(?:\.\d+)?$/.test(text)){const serial=Number(text);if(serial>=20000&&serial<=80000)return Date.UTC(1899,11,30)+Math.round(serial*86400000);}
  // The browser parser also accepts plain numbers such as "1"; use it only for English month-name dates.
  if(!/\d{4}/.test(text)||!/[A-Za-z]{3}/.test(text))return null;
  const parsed=Date.parse(text);return Number.isFinite(parsed)?parsed:null;
}

function parseCreateQuoteInput(text){
  const source=String(text||'').replace(/\r/g,'').split('\n').map((line,index)=>{const values=splitZbomLine(line);return {line:index+1,values:values.length===1&&/\s/.test(line.trim())?line.trim().split(/\s+/).map(value=>value.trim()):values};}).filter(item=>item.values.some(Boolean));
  if(!source.length)return {rows:[],duplicates:0,ignored:0};
  const header=source[0].values.map(createQuoteHeaderRole),hasHeader=header.some(Boolean),data=hasHeader?source.slice(1):source,rows=[],seen=new Map();let duplicates=0,ignored=0;
  data.forEach(item=>{
    const cells=item.values.map((value,index)=>({value:String(value||'').trim(),index:index})).filter(cell=>cell.value);
    // Header columns apply to rows with the same layout; other rows find the FID and dates in order.
    const mapped=hasHeader&&item.values.length===header.length,column=role=>cells.find(cell=>cell.index===header.indexOf(role));
    const fidCell=mapped&&header.includes('fid')?column('fid'):cells.find(cell=>/^\d{6}$/.test(cell.value));
    if(!fidCell||!/^\d{6}$/.test(fidCell.value)){ignored++;return}
    const dates=mapped&&(header.includes('crd')||header.includes('validTo'))?[column('crd'),column('validTo')]:cells.filter(cell=>cell.index!==fidCell.index&&createQuoteDateValue(cell.value)!==null);
    const row={fid:fidCell.value,crd:dates[0]?dates[0].value:'',validTo:dates[1]?dates[1].value:''},existing=seen.get(fidCell.value);
    if(existing){duplicates++;if(!existing.crd&&row.crd)existing.crd=row.crd;if(!existing.validTo&&row.validTo)existing.validTo=row.validTo;return}
    seen.set(fidCell.value,row);rows.push(row);
  });
  return {rows:rows,duplicates:duplicates,ignored:ignored};
}

function formatCreateQuoteDate(date){return String(date.getMonth()+1).padStart(2,'0')+'/'+String(date.getDate()).padStart(2,'0')+'/'+date.getFullYear();}
function formatCreateQuoteDateParts(year,month,day){return createQuoteUtcDate(year,month,day)===null?'':String(Number(month)).padStart(2,'0')+'/'+String(Number(day)).padStart(2,'0')+'/'+String(Number(year)).padStart(4,'0');}

function normalizeCreateQuoteDate(value){
  const text=String(value||'').trim();if(!text||/^\d{6}$/.test(text))return '';
  let match=text.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if(match)return formatCreateQuoteDateParts(match[1],match[2],match[3]);
  match=text.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if(match)return formatCreateQuoteDateParts(match[1],match[2],match[3]);
  match=text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if(match)return formatCreateQuoteDateParts(match[1],match[2],match[3]);
  match=text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if(match){const first=Number(match[1]),second=Number(match[2]),month=first>12?second:first,day=first>12?first:second; return formatCreateQuoteDateParts(match[3],month,day);}
  if(/^\d{4,5}(?:\.\d+)?$/.test(text)){const serial=Number(text); if(serial>=20000&&serial<=80000){const date=new Date(Date.UTC(1899,11,30)+Math.round(serial*86400000));return formatCreateQuoteDateParts(date.getUTCFullYear(),date.getUTCMonth()+1,date.getUTCDate())}}
  if(!/\d{4}/.test(text)||!/[A-Za-z]{3}/.test(text))return '';
  const parsed=Date.parse(text);if(!Number.isFinite(parsed))return '';
  const formatted=formatCreateQuoteDate(new Date(parsed));return /^\d{2}\/\d{2}\/\d{4}$/.test(formatted)?formatted:'';
}

function defaultCreateQuoteValidTo(){const input=bridgeId('createQuoteDefaultValidTo'),raw=normalizeCreateQuoteDays(input&&input.value),days=raw===''?Number(CREATE_QUOTE_DEFAULT_VALID_TO_DAYS):Number(raw),date=new Date(); date.setHours(12,0,0,0);date.setDate(date.getDate()+days);
  return formatCreateQuoteDate(date);}
function createQuotePreviewFid(fid){const suffix=bridgeId('createQuoteFidSuffix');return String(fid||'')+String(suffix&&suffix.value||'');}

function renderCreateQuotePreview(){
  const body=bridgeId('createQuotePreviewBody');body.replaceChildren();
  if(!createQuoteRows.length){const tr=document.createElement('tr'),td=document.createElement('td');td.colSpan=4;td.className='missing';td.textContent='No preview yet.';tr.appendChild(td);body.appendChild(tr)}
  else createQuoteRows.forEach((row,index)=>{
    const tr=document.createElement('tr'),item=document.createElement('td'),crd=document.createElement('td'),crdInput=document.createElement('input'),version=document.createElement('td'),validTo=document.createElement('td'); item.textContent=String(index+1);version.textContent=
      row.version||createQuotePreviewFid(row.fid);crdInput.type='text';crdInput.disabled=sapToolRunning('createQuote');crdInput.className='create-quote-crd-input';crdInput.value=row.crd||'';crdInput.placeholder='CRD';
      crdInput.setAttribute('aria-label','CRD for FID '+row.fid); crdInput.addEventListener('input',()=>{crdInput.classList.remove('invalid');row.crd=crdInput.value;createQuoteManualCrd.add(row.fid);createQuoteCrdOverrides.set(row.fid,row.crd);saveCreateQuoteCrdOverrides()});
    crd.appendChild(crdInput);validTo.textContent=row.validTo||defaultCreateQuoteValidTo();tr.append(item,crd,version,validTo);body.appendChild(tr);
  });
  const count=createQuoteRows.length,parts=[count+' FID'+(count===1?'':'s')];
  if(createQuoteStats.duplicates)parts.push(createQuoteStats.duplicates+' duplicate'+(createQuoteStats.duplicates===1?'':'s')+' removed');
  if(createQuoteStats.ignored)parts.push(createQuoteStats.ignored+' row'+(createQuoteStats.ignored===1?'':'s')+' ignored');
  bridgeId('createQuoteCount').textContent=parts.join(' / '); updateCreateQuoteStart();
}

function previewCreateQuoteInput(){
  const parsed=parseCreateQuoteInput(bridgeId('createQuoteInput').value);createQuoteRows=parsed.rows;createQuoteRows.forEach(row=>{
    const override=createQuoteCrdOverrides.get(row.fid);
    if(createQuoteManualCrd.has(row.fid))row.crd=override||''; else if(String(override||'').trim())row.crd=override; else if(!row.crd){const source=bobjValue(row.fid,'Date: Customer Request');if(source&&source!=='#')row.crd=normalizeCreateQuoteDate(source)||source}
  });createQuoteStats={duplicates:parsed.duplicates,ignored:parsed.ignored};renderCreateQuotePreview();
}

function updateCreateQuoteStart(){
  const button=bridgeId('createQuoteStart');if(!button)return;
  button.disabled=!canStartSapTool('createQuote')||(!pendingToolboxSubmissions.has('createQuote')&&(!sanitizeCreateQuoteQuotation()||!createQuoteRows.length));
}

function validateCreateQuotePreview(){
  const inputs=Array.from(document.querySelectorAll('#createQuotePreviewBody .create-quote-crd-input')),normalized=[];let firstInvalid=null,error='';
  createQuoteRows.forEach((row,index)=>{
    const input=inputs[index],crdText=String(input?input.value:row.crd).trim(),crd=crdText?normalizeCreateQuoteDate(crdText):'',validTo=normalizeCreateQuoteDate(row.validTo||defaultCreateQuoteValidTo());
    if(input)input.classList.remove('invalid');
    if(crdText&&!crd){if(input){input.classList.add('invalid');if(!firstInvalid)firstInvalid=input} if(!error)error='Item '+(index+1)+': CRD is not a recognized date.'; return;}
    if(!validTo){if(!error)error='Item '+(index+1)+': Valid to is not a recognized date.';return}
    normalized[index]={crd:crd,validTo:validTo};
  });
  if(error){setSapProgress('createQuote',error,'error');if(firstInvalid)firstInvalid.focus();return false}
  normalized.forEach((value,index)=>{createQuoteRows[index].crd=value.crd;createQuoteRows[index].validTo=value.validTo;createQuoteCrdOverrides.set(createQuoteRows[index].fid,value.crd)}); saveCreateQuoteCrdOverrides();renderCreateQuotePreview();return true;
}


function zbomCompareHeaderRole(value){
  const role=zbomHeaderRole(value);if(role)return role;
  const key=normalizedZbomHeader(value);
  if(['group','groupid','comparegroup','comparisongroup','zbomgroup'].includes(key))return 'group';
  if(['customer','customername','account','accountname'].includes(key))return 'customer';
  return '';
}

function normalizeZbomCompareGroup(value,allowCustom){const text=String(value||'').trim(); if(!text)return ''; if(/^standalone$/i.test(text)||/^do[\s_-]*not[\s_-]*compare$/i.test(text))return ZBOM_COMPARE_STANDALONE; if(/^g\d+$/i.test(text))return 'G'+text.slice(1);
  return allowCustom&&/^[A-Za-z0-9]+$/.test(text)?text:'';}

function zbomCompareGroupLabel(number){return 'G'+String(number).padStart(2,'0');}

function compareGroupSelect(label,value,groups,disabled){
  const select=document.createElement('select');select.className='compare-group-select';select.setAttribute('aria-label',label);
  for(const group of groups)select.add(new Option(group===ZBOM_COMPARE_STANDALONE?'Standalone':group,group));
  select.value=value;select.disabled=disabled;
  return select;
}

function zbomCompareGroupOptions(){
  const selected=new Set(zbomCompareRows.map(row=>row.group)),groups=[zbomCompareGroupLabel(1)],custom=[],customKeys=new Set();let number=1;
  while(number<500&&selected.has(zbomCompareGroupLabel(number))){number++;groups.push(zbomCompareGroupLabel(number))}
  zbomCompareRows.forEach(row=>{if(row.group&&row.group!==ZBOM_COMPARE_STANDALONE&&!groups.includes(row.group)){const key=row.group;if(!customKeys.has(key)){customKeys.add(key);custom.push(row.group)}}}); return groups.concat(custom,ZBOM_COMPARE_STANDALONE);
}

function plannedZbomCompareFileNames(showPendingFid=false){
  const template=bridgeId('zbomCompareTemplate').value.trim()||ZBOM_COMPARE_DEFAULT_TEMPLATE;
  const separator=String(bridgeId('zbomCompareSeparator').value||'').trim()||ZBOM_COMPARE_DEFAULT_SEPARATOR;
  const pendingFid=row=>showPendingFid&&!bobjFidKey(row.fid);
  const render=(row,omit)=>compactZbomCompareName(renderBobjTemplate(template,row,omit,pendingFid(row)))||row.quote;
  const names=new Map(),groups=new Map();
  for(const row of zbomCompareRows){
    if(row._error||row._duplicate)continue;
    if(row.group===ZBOM_COMPARE_STANDALONE){
      names.set(row.quote,zbomFileNames.zbomCompare.get(zbomFileNameKey('zbomCompare',row))||finishZbomCompareFileName(render(row),undefined,pendingFid(row)));
      continue;
    }
    if(!groups.has(row.group))groups.set(row.group,[]);
    groups.get(row.group).push(row);
  }
  for(const rows of groups.values()){
    const name=zbomFileNames.zbomCompare.get(zbomFileNameKey('zbomCompare',rows[0]))||
      finishZbomCompareFileName(joinCompareFileNames(rows,separator,bobjTemplateValue,render),undefined,rows.some(pendingFid));
    for(const row of rows)names.set(row.quote,name);
  }
  return names;
}

function selectZbomQuote(tool,row,quote,resolvedLatest=false){
  const {key,fid,latest}=row.quotePicker;
  zbomQuoteSelections[tool].set(key,{fid,quote,resolvedLatest:resolvedLatest||!latest});
  if(tool==='zbomCompare'&&!zbomCompareSavedGroups.has(quote)){
    const group=zbomCompareSavedGroups.get(row.quote)||row.group;
    if(group)zbomCompareSavedGroups.set(quote,group);
  }
  if(resolvedLatest&&(tool==='zbom'||zbomCompareSavedGroups.get(quote)===ZBOM_COMPARE_STANDALONE)){
    const names=zbomFileNames[tool],oldName='quote:'+row.quote,newName='quote:'+quote;
    if(names.has(oldName)&&!names.has(newName)){
      names.set(newName,names.get(oldName));names.delete(oldName);saveZbomFileNames(tool);
    }
  }
}

function renderZbomQuoteCell(cell,row,tool){
  if(!row.quotePicker)return;
  const {key,fid,choices,latest}=row.quotePicker;
  const select=document.createElement('select'),rich=CSS.supports('appearance:base-select');
  select.className='zbom-quote-select';select.dataset.quoteKey=key;select.setAttribute('aria-label','Quote for FID '+row.fid);
  select.disabled=sapToolRunning(tool);
  const label=latest&&row.quote===latest?ZBOM_LATEST_LABEL:row.quote||'—';
  if(rich){
    const button=document.createElement('button');button.type='button';button.textContent=label;select.append(button);
    cell.replaceChildren(select);
  }else{
    const wrapper=document.createElement('span'),value=document.createElement('span');
    wrapper.className='zbom-quote-fallback';value.textContent=label;value.setAttribute('aria-hidden','true');
    wrapper.append(select,value);cell.replaceChildren(wrapper);
  }
  const sortedChoices=Array.from(choices,([quote,ctls])=>({
    quote,links:[...(ctls||[])].sort(([a,first],[b,second])=>second.created.localeCompare(first.created)||b.localeCompare(a))
  })).sort((a,b)=>b.quote.localeCompare(a.quote));
  if(latest)sortedChoices.unshift({quote:latest,links:[]});
  for(const {quote,links} of sortedChoices){
    const option=document.createElement('option');option.value=quote;
    const titleText=quote===latest?ZBOM_LATEST_LABEL:quote;
    const details=quote===latest?['Find by SAP creation time before export, including drafts.']:
      links.length?links.map(([ctl,detail])=>[ctl,detail.change,detail.stage,'Created '+(detail.created||'—')].filter(Boolean).join(' / ')):['Created —'];
    if(rich){
      const title=document.createElement('strong'),meta=document.createElement('small');
      title.textContent=titleText;meta.textContent=details.join('\n');option.append(title,meta);
    }else option.textContent=titleText+' / '+details.join(' | ');
    select.append(option);
  }
  if(row.quote!==latest)select.value=row.quote;
  select.addEventListener('change',()=>{
    const quote=select.value,body=cell.closest('tbody');
    if(!body||!quote||sapToolRunning(tool))return;
    selectZbomQuote(tool,row,quote);
    if(tool==='zbom')previewZbomInput();
    else previewZbomCompareInput();
    const picks=Array.from(body.querySelectorAll('.zbom-quote-select'));
    (picks.find(pick=>pick.dataset.quoteKey===key)||picks.find(pick=>pick.value===quote))?.focus({preventScroll:true});
  });
}

function renderZbomPreviewTable(tool,rows,template,fileNames,templateError=''){
  const compare=tool==='zbomCompare',body=bridgeId(tool+'PreviewBody'),head=body.closest('table').tHead,fields=bobjTemplateFields(template);
  const columns=[{label:'Quotation',field:'Quote'},{label:'FID',field:'FID',className:'compare-fid'},{label:'System Description',field:'System Description',className:'compare-system-description'}];
  const seen=new Set(columns.map(column=>normalizeBobjHeader(column.field)));
  seen.add('fidver');
  for(const [,token] of String(template).matchAll(/\{([^{}]+)\}/g)){
    const field=normalizeBobjHeader(token);if(seen.has(field))continue;seen.add(field);
    columns.push({label:token.trim(),field,className:field==='customer'?'compare-customer':''});
  }
  if(compare)columns.push({label:'Compare Group',kind:'group',className:'compare-group-cell'});
  columns.push({label:'File Name',kind:'file',className:compare?'compare-file-name':'planned-name'},{label:'Status',kind:'status',className:compare?'compare-status':''});
  const heading=document.createElement('tr');
  for(const column of columns){const cell=document.createElement('th');cell.textContent=column.label;heading.appendChild(cell)}
  head.replaceChildren(heading);body.replaceChildren();
  if(!rows.length){
    const tr=document.createElement('tr'),cell=document.createElement('td');cell.colSpan=columns.length;cell.className='missing';cell.textContent='No preview yet.';tr.appendChild(cell);body.appendChild(tr);return;
  }
  const groups=compare?zbomCompareGroupOptions():[];
  for(const row of rows){
    const latest=toolboxLatestQuoteFid(row.quote),quoteLabel=latest?'latest SAP quote for FID '+latest:'quotation '+row.quote;
    const key=zbomFileNameKey(tool,row),edited=zbomFileNames[tool].has(key);
    const tr=document.createElement('tr'),issue=bobjPreviewIssue(row,edited?[]:fields),lookup=needsZbomSapLookup(row,fields,edited);
    for(const column of columns){
      const cell=document.createElement('td');cell.className=column.className||'';
      if(column.kind==='group'){
        const select=compareGroupSelect('Compare group for '+quoteLabel,row.group,groups,sapToolRunning(tool)||Boolean(row._error));
        select.dataset.quote=row.quote;select.dataset.rowKey=row.key;select.toggleAttribute('data-preview-disabled',Boolean(row._error));
        select.addEventListener('change',()=>{
          for(const item of zbomCompareRows)if(item.quote===row.quote)item.group=select.value;
          saveZbomCompareGroups();renderZbomComparePreview();
          Array.from(body.querySelectorAll('.compare-group-select')).find(item=>item.dataset.rowKey===row.key)?.focus({preventScroll:true});
        });
        cell.appendChild(select);
      }else if(column.kind==='file'){
        if(row._duplicate&&!compare){tr.appendChild(cell);continue}
        const name=fileNames.get(row.quote)||'',value=document.createElement('span'),edit=document.createElement('button');
        const render=()=>{if(compare)renderZbomComparePreview();else renderZbomPreview();Array.from(body.querySelectorAll('.zbom-file-edit')).find(item=>item.dataset.rowKey===row.key)?.focus({preventScroll:true})};
        value.className='zbom-file-value';value.textContent=name;edit.type='button';edit.className='zbom-file-edit';
        if(!edited){
          if(latest)value.title='SAP resolves the quotation before export.';
          else if(!bobjFidKey(row.fid)&&name.includes('{'))value.title='Filename fields will be matched using the FID from SAP.';
          else if(name.includes('{System Description}'))value.title='System Description will be matched using the FID from SAP.';
          else if(name.includes('{FIDver}'))value.title='SAP fills {FIDver} after export when a suffix exists.';
        }
        edit.dataset.rowKey=row.key;
        edit.setAttribute('aria-label','Edit file name for '+quoteLabel);edit.title='Edit file name';edit.disabled=sapToolRunning(tool)||!name;
        edit.toggleAttribute('data-preview-disabled',!name);cell.append(value,edit);
        if(edited){
          const reset=document.createElement('button');reset.type='button';reset.className='zbom-file-reset';reset.title='Reset to template';reset.setAttribute('aria-label','Reset file name for '+quoteLabel);reset.disabled=sapToolRunning(tool);
          reset.addEventListener('click',()=>{zbomFileNames[tool].delete(key);saveZbomFileNames(tool);render()});cell.appendChild(reset);
        }
        edit.addEventListener('click',()=>{
          const input=document.createElement('input'),confirm=document.createElement('button'),editor=document.createElement('span'),extension=compare?'.xlsx':'.PDF';
          const planned=edited?name:(compare?plannedZbomCompareFileNames():plannedZbomNames(rows,template)).get(row.quote);
          const suggested=edited?planned:planned.replace(/\{(?:FIDver|System Description)\}/g,'');
          input.type='text';input.className='zbom-file-input';input.value=suggested;input.maxLength=255;input.setAttribute('aria-label','File name for '+quoteLabel);input.title='Click ✓ to save / Esc to cancel';
          confirm.type='button';confirm.className='zbom-file-save';confirm.title='Save file name for '+quoteLabel;editor.className='zbom-file-editor';editor.append(input,confirm);
          const changedName=()=>{const text=input.value.trim();return text&&!text.toLowerCase().endsWith(extension.toLowerCase())?text+extension:text};
          confirm.addEventListener('click',()=>{
            const next=changedName(),error=zbomFileNameError(next,tool);input.setCustomValidity(error);
            if(error){input.reportValidity();return}
            if(next!==suggested||edited)zbomFileNames[tool].set(key,next);
            saveZbomFileNames(tool);render();
          });
          input.addEventListener('input',()=>input.setCustomValidity(zbomFileNameError(changedName(),tool)));
          input.addEventListener('keydown',event=>{
            if(event.key==='Enter')event.preventDefault();
            else if(event.key==='Escape'){event.preventDefault();render()}
          });
          cell.replaceChildren(editor);if(compare)updateZbomCompareStart();else updateZbomStart();input.focus();input.setSelectionRange(0,Math.max(0,suggested.length-extension.length));
        });
      }
      else if(column.kind==='status'){
        const error=row._error||templateError;
        cell.textContent=error||(row._duplicate?'Duplicated':lookup?'Look up in SAP at start':row._afterExport?'Look up after export':row._lookup?'Check FID':issue?'Missing fields':'Ready');cell.title=issue;
        cell.classList.add(error?'row-error':row._duplicate||issue||lookup?'row-warning':'row-ok');
      }else{
        const value=bobjTemplateValue(row,column.field);cell.textContent=value||(row._afterExport&&column.field==='FID'?'After export':'—');if(!value)cell.classList.add('missing');
        if(column.field==='Quote')renderZbomQuoteCell(cell,row,tool);
      }
      tr.appendChild(cell);
    }
    body.appendChild(tr);
  }
}

function renderZbomComparePreview(){
  const template=bridgeId('zbomCompareTemplate').value.trim()||ZBOM_COMPARE_DEFAULT_TEMPLATE;
  renderZbomPreviewTable('zbomCompare',zbomCompareRows,template,plannedZbomCompareFileNames(true));
  const count=zbomCompareRows.filter(row=>!row._error).length,parts=[count+' quotation'+(count===1?'':'s')];
  if(zbomCompareRows.length!==count)parts.push((zbomCompareRows.length-count)+' need a quotation');
  if(zbomCompareStats.ignored)parts.push(zbomCompareStats.ignored+' line'+(zbomCompareStats.ignored===1?'':'s')+' ignored');
  bridgeId('zbomCompareCount').textContent=parts.join(' / ');updateZbomCompareStart();
}

function previewZbomCompareInput(keepGroups=true){
  const parsed=parseZbomInput(bridgeId('zbomCompareInput').value,'zbomCompare');zbomCompareRows=parsed.rows;zbomCompareStats={ignored:parsed.ignored};
  if(keepGroups)zbomCompareRows.forEach(row=>{if(zbomCompareSavedGroups.has(row.quote))row.group=zbomCompareSavedGroups.get(row.quote)});
  if(!sapToolRunning('zbomCompare')&&!pendingToolboxSubmissions.has('zbomCompare'))setSapProgress('zbomCompare','Ready.','idle');
  saveZbomCompareGroups();renderZbomComparePreview();
}

function updateZbomCompareStart(){
  const button=bridgeId('zbomCompareStart');if(!button)return;
  const path=bridgeId('zbomCompareOutputFolder').value.trim(),validPath=/^[A-Za-z]:[\\/]/.test(path);
  const ready=updateZbomLookupButton('zbomCompare',zbomCompareRows);
  button.disabled=!ready||(!pendingToolboxSubmissions.has('zbomCompare')&&(!zbomCompareRows.length||!validPath));
}
function normalizedZbomCompareFolder(){const input=bridgeId('zbomCompareOutputFolder'),raw=String(input&&input.value||'').trim().replace(/\//g,'\\'); if(!raw||!/^[A-Za-z]:\\/.test(raw))throw new Error('Use an absolute local output folder such as C:\\PDFFILES\\Compare.');
  return raw.length>3?raw.replace(/\\+$/,''):raw;}

function isZbomFid(value){return /^\d{6}(?:-[A-Za-z0-9-]+)?$/.test(String(value||'').trim());}
function pickZbomValue(values,index){return index>=0&&index<values.length?String(values[index]||'').trim():'';}
function zbomInputKey(line,values){return JSON.stringify([line,values]);}

function zbomInputLines(text){
  const unmapped={fid:-1,ctl:-1,quote:-1,customer:-1,systemDescription:-1,group:-1,identifier:-1};
  let mapping=unmapped,columns=0;
  const source=String(text||'').replace(/\r\n?/g,'\n').split('\n').map((line,index)=>{
    const values=splitZbomLine(line);return {line:index+1,values,key:zbomInputKey(index+1,values)};
  }).filter(item=>item.values.some(Boolean));
  for(const item of source){
    const roles=item.values.map(zbomCompareHeaderRole);
    item.header=roles.some(role=>['fid','ctl','quote','identifier'].includes(role))&&!item.values.some(value=>/\d{6}/.test(value.normalize('NFKC')));
    if(item.header){
      mapping=Object.fromEntries(Object.keys(mapping).map(role=>[role,-1]));
      roles.forEach((role,index)=>{if(role&&mapping[role]<0)mapping[role]=index});
      columns=item.values.length;
    }
    item.mapping=item.values.length===columns?mapping:unmapped;
  }
  return source;
}

function classifyZbomRows(item,tool){
  const {values,mapping}=item,roles=Object.keys(mapping),identifiers={fid:new Set(),ctl:new Set(),quote:new Set()};
  const labels={fid:'FID',ctl:'CTL ID',quote:'Quote'},row={key:item.key,fid:'',ctl:'',quote:'',customer:'',systemDescription:'',group:''},metadata=[];
  const pattern=/(?<![A-Za-z0-9_-])(?:(FID|Facility\s*ID|Opportunity\s*ID|CTL|Quote|Quotation|报价(?:单)?(?:号)?)(?:\s*(?:ID|No\.?|Number|#))?\s*[:#=]?\s*)?(\d{10}|\d{8}|\d{6}(?:-[A-Za-z0-9]+)*)(?:\.0+)?(?![A-Za-z0-9_-]|\.\d)/giu;
  let nsr=false;
  values.forEach((raw,index)=>{
    if(!raw)return;
    const hint=roles.find(role=>mapping[role]===index);
    if(!hint&&roles.some(role=>mapping[role]>=0))return;
    if(['customer','systemDescription','group'].includes(hint)){row[hint]=hint==='group'?normalizeZbomCompareGroup(raw,true):raw;return}
    const value=raw.normalize('NFKC');
    nsr||=/(?:^|[^\p{L}\p{N}_])(?:NSR\s*[:#=]?\s*)?[NDL]\d{5,8}(?!\d)/iu.test(value);
    let matched=false;
    const rest=value.replace(pattern,(match,label,number,offset)=>{
      if(!label&&/(?:\b(?:date|today|created|updated|so|sales\s*order)|日期|时间|订单号)\s*[:=]?\s*$/i.test(value.slice(0,offset)))return match;
      matched=true;
      // Length identifies the number even when neighbouring rows use a different column layout.
      let role=zbomHeaderRole(label);
      if(!role)role=/^\d{10}$/.test(number)?'ctl':/^\d{8}$/.test(number)?'quote':'fid';
      if(!label&&hint==='quote'&&/^00\d{8}$/.test(number))role='quote';
      let id=number;
      if(role==='quote')id=bobjQuoteKey(number);
      else if(role==='fid'&&!isZbomFid(number)||role==='ctl'&&!/^\d{10}$/.test(number))id='';
      if(id)identifiers[role].add(id);else row._error=labels[role]+' has an invalid number of digits.';
      return ' ';
    }).trim();
    if(identifiers[hint]&&!matched)row._error=labels[hint]+' was not detected in its column.';
    if(!matched&&/^['"]?\d+$/.test(value.trim())&&(hint||values.length===1))row._error='Use a 6-digit FID, 10-digit CTL ID or 8-digit Quote.';
    const remaining=matched?rest:raw;
    if(remaining&&!zbomCompareHeaderRole(remaining)&&/\p{L}/u.test(remaining)){
      const group=normalizeZbomCompareGroup(remaining,false);
      if(!row.group&&group)row.group=group;else metadata.push(remaining);
    }
  });
  for(const [role,ids] of Object.entries(identifiers)){
    row[role]=ids.values().next().value||'';
    if(ids.size>1&&role!=='quote')row._error='Multiple '+labels[role]+' numbers were entered in one row.';
  }
  if(nsr&&!Object.values(identifiers).some(ids=>ids.size))row._error='NSR numbers are not supported. Enter a FID, CTL ID or Quote.';
  if(!row.systemDescription){
    const marked=metadata.findIndex(value=>/[\/_]/.test(value));
    if(marked>=0)row.systemDescription=metadata.splice(marked,1)[0];
  }
  if(!row.customer&&!row.systemDescription){
    if(metadata.length===1)row.systemDescription=metadata[0];
    else if(metadata.length>1){row.customer=metadata[0];row.systemDescription=metadata[1]}
  }else if(!row.customer&&metadata.length)row.customer=metadata[0];else if(!row.systemDescription&&metadata.length)row.systemDescription=metadata[0];
  row._groupExplicit=Boolean(row.group);
  return Array.from(identifiers.quote.size?identifiers.quote:[''],quote=>{
    const input=identifiers.quote.size>1?{...item,key:JSON.stringify([item.line,values,quote])}:item;
    const entry={...row,key:input.key,quote};return entry._error?entry:matchBobjInput(entry,entry.ctl,input,tool);
  });
}

function parseZbomInput(text,tool='zbom'){
  const rows=[],seen=new Map(),selections=zbomQuoteSelections[tool];let ignored=0;
  for(const item of zbomInputLines(text).filter(item=>!item.header))for(const row of classifyZbomRows(item,tool)){
    if(tool==='zbomCompare'&&!row.quote&&!row._error){ignored++;continue}
    if(tool==='zbomCompare'&&!row.group)row.group=zbomCompareSavedGroups.get(row.quote)||ZBOM_COMPARE_DEFAULT_GROUP;
    if(!row.quote)row._error||='Quote was not detected.';else if(!/^\d{8}$/.test(row.quote)&&!toolboxLatestQuoteFid(row.quote))row._error='Quote must contain exactly 8 digits.';
    if(row._error){rows.push(row);continue}
    const previous=seen.get(row.quote);row._duplicate=Boolean(previous);
    if(previous){
      for(const field of ['fid','customer','systemDescription'])if(!previous[field]&&row[field])previous[field]=row[field];
      if(!previous._groupExplicit&&row._groupExplicit){previous.group=row.group;previous._groupExplicit=true}
    }
    else seen.set(row.quote,row);
    rows.push(row);
  }
  if(tool==='zbomCompare')for(const row of rows)if(row._duplicate)row.group=seen.get(row.quote).group;
  const keys=new Set(rows.map(row=>row.key));for(const key of selections.keys())if(!keys.has(key))selections.delete(key);
  return {rows,ignored};
}

function todayZbomStamp(){const date=new Date(),part=value=>String(value).padStart(2,'0'); return String(date.getFullYear())+part(date.getMonth()+1)+part(date.getDate());}

function zbomTemplateInfo(){
  const input=bridgeId('zbomTemplate'),template=String(input.value||'').trim()||ZBOM_DEFAULT_TEMPLATE;
  if(template.length>240)return {template:template,error:'The file-name template is longer than 240 characters.'};
  const matches=template.match(/\{[^{}]+\}/g)||[],empty=matches.find(token=>!token.slice(1,-1).trim());
  if(empty)return {template:template,error:'The file-name template contains an empty token.'};
  const remainder=template.replace(/\{[^{}]+\}/g,'');
  if(remainder.indexOf('{')>=0||remainder.indexOf('}')>=0)return {template:template,error:'The file-name template contains an incomplete token.'};
  return {template:template,error:''};
}

function renderZbomFileName(row,template,pendingFid){
  const name=renderBobjTemplate(template||ZBOM_DEFAULT_TEMPLATE,row,undefined,pendingFid);
  let stem=sanitizeToolboxFileName(name,'-',pendingFid).replace(/_{2,}/g,'_').replace(/-{2,}/g,'-').replace(/\s{2,}/g,' ').replace(/^[\s._-]+|[\s._-]+$/g,'');
  if(!stem)stem=row.quote||'ZBOM_PDF';
  if(!pendingFid&&stem.length>190)stem=stem.slice(0,190).replace(/[\s._-]+$/g,'');
  return stem+'.PDF';
}

function appendQuoteToName(name,quote){const dot=name.toLowerCase().lastIndexOf('.pdf'),stem=dot>=0?name.slice(0,dot):name; return stem+'_'+quote+'.PDF';}

function plannedZbomNames(rows,template,showPendingFid=false){
  const claims=new Map(),names=new Map();
  for(const row of rows){
    if(row._error||row._duplicate)continue;
    const name=zbomFileNames.zbom.get(zbomFileNameKey('zbom',row))||renderZbomFileName(row,template,showPendingFid&&!bobjFidKey(row.fid));
    names.set(row.quote,claimZbomOutputName(name,row.quote,claims));
  }
  return names;
}

function renderZbomPreview(){
  const {template,error}=zbomTemplateInfo();
  renderZbomPreviewTable('zbom',zbomRows,template,plannedZbomNames(zbomRows,template,true),error);
  const valid=zbomRows.filter(row=>!row._error).length;bridgeId('zbomCount').textContent=valid+' valid row'+(valid===1?'':'s')+(zbomRows.length!==valid?' / '+(zbomRows.length-valid)+' invalid':'');
  updateZbomStart();
}

function previewZbomInput(){
  zbomRows=parseZbomInput(bridgeId('zbomInput').value).rows;
  renderZbomPreview();
}

function insertZbomToken(tool,token,update){
  const input=bridgeId(tool+'Template'),start=input.selectionStart??input.value.length,end=input.selectionEnd??start;
  input.value=input.value.slice(0,start)+token+input.value.slice(end);input.focus();
  input.setSelectionRange(start+token.length,start+token.length);update();
}

function updateZbomLookupButton(tool,rows){
  const ready=canStartSapTool(tool)&&!bridgeId(tool+'PreviewBody').querySelector('.zbom-file-input');
  const lookup=bridgeId(tool+'SapLookup'),pending=rows.some(row=>zbomLookupPending(tool,row));
  lookup.hidden=!pending&&rows.some(row=>row.quote&&!row._error);
  lookup.disabled=!ready||!pending;
  return ready;
}

function updateZbomStart(){
  const button=bridgeId('zbomStart');if(!button)return;
  button.disabled=!updateZbomLookupButton('zbom',zbomRows)||(!pendingToolboxSubmissions.has('zbom')&&!zbomRows.length);
}

function parsePartChangesInput(value){
  let columns=[];
  const rows=[];
  const quotationPair=(first,second='')=>isZbomFid(first)&&bobjQuoteKey(second)?[first,second]:
    isZbomFid(second)&&bobjQuoteKey(first)?[second,first]:null;
  for(const [index,line] of value.replace(/\r\n?/g,'\n').split('\n').entries()){
    const fields=splitZbomLine(line);
    if(!fields.some(Boolean))continue;
    const entry=fields[0].normalize('NFKC'),words=entry.split(/\s*\+\s*|\s+/);
    if(quotationPair(words[0],words[1])||fields.length===1&&/^(?:\S+\s+(?:G\d+|Standalone|__STANDALONE__)|FID\s+Group)$/i.test(entry))
      fields.splice(0,1,...words);
    const roles=fields.map(field=>/^(?:so|sonumber|salesorder|salesordernumber|salesorderquotation|document|documentnumber)$/.test(normalizedZbomHeader(field))
      ?'document':zbomCompareHeaderRole(field));
    if(roles.some(role=>['fid','identifier','document','quote'].includes(role))&&fields.every((field,i)=>!field||roles[i])){
      columns=roles;continue;
    }
    const mapped=columns.length===fields.length,read=role=>mapped?pickZbomValue(fields,columns.indexOf(role)):'';
    const [first,second='']=fields.map(field=>field.normalize('NFKC'));
    const pair=!mapped&&quotationPair(first,second);
    const fid=read('fid')||pair?.[0],document=read('document')||read('quote')||read('identifier')||pair?.[1];
    const raw=(document||(mapped?fid:first)||'').normalize('NFKC').trim().toUpperCase();
    const Input=bobjQuoteKey(raw)||raw,type=partChangesDocumentType(Input);
    const group=read('group')||(!mapped?fields[pair?2:1]:'')||ZBOM_COMPARE_DEFAULT_GROUP;
    const groupNumber=/^G(0*[1-9]\d*)$/i.exec(group)?.[1]?.replace(/^0+/,'');
    const standalone=group===ZBOM_COMPARE_STANDALONE||/^standalone$/i.test(group);
    const expectedFid=fid&&document&&type?partChangesFidOf(fid):'';
    const error=!type&&!/^\d{6}$/.test(Input)?'Enter a six-digit FID, an eight-digit quotation, or a letter-prefixed SO.':
      fid&&document&&type&&!expectedFid?'Enter a valid FID in the FID column.':
      (!mapped&&fields.length>(pair?3:2))||!groupNumber&&!standalone?'Use an identifier and optional Group, such as G01 or Standalone.':'';
    rows.push({key:zbomInputKey(index+1,fields),Input,type,expectedFid,Group:standalone?ZBOM_COMPARE_STANDALONE:groupNumber?'G'+groupNumber.padStart(2,'0'):ZBOM_COMPARE_DEFAULT_GROUP,
      customer:read('customer'),systemDescription:read('systemDescription'),error});
  }
  return rows;
}

function previewPartChanges(){
  partChangesRows=parsePartChangesInput(bridgeId('partChangesInput').value).map(row=>{
    const choice=partChangesSelections.get(row.key),lookup=partChangesLookups.get(row.Input),options=lookup?.documents;
    const selection=choice?.document||(!row.type?PART_CHANGES_LATEST_SO:'');
    const policy=!row.type&&[PART_CHANGES_LATEST_SO,PART_CHANGES_LATEST_QUOTE].includes(selection)?selection:'';
    const soIssue=lookup?.issues.find(issue=>issue.startsWith('SO: '));
    const order=policy===PART_CHANGES_LATEST_QUOTE?options?.find(item=>item.type==='quotation'):
      policy===PART_CHANGES_LATEST_SO?options?.find(item=>item.type==='so')||(!soIssue?options?.find(item=>item.type==='quotation'):undefined):
      selection?options?.find(item=>partChangesDocumentKey(item.type,item.number)===selection):options?.[0];
    const Fid=order?partChangesFidOf(order.version):row.expectedFid||(!row.type?row.Input:'');
    const error=row.error||(order&&row.expectedFid&&Fid!==row.expectedFid?'SAP did not confirm FID '+row.expectedFid+' for '+row.Input+'.':'');
    const Quote=order?.number||(row.type?row.Input:''),type=order?.type||row.type;
    return {...row,Fid,Quote,order,options,selection,policy,myLamSystem:choice?.myLamSystem||'',documentKey:Quote?partChangesDocumentKey(type,Quote):policy?row.Input+':'+policy:'',
      Group:choice?.Group||row.Group,lookupIssue:lookup?.issues.join(' · ')||'',error};
  });
  renderPartChangesPreview();
}

function resetPartChangesInput(){
  bridgeId('partChangesInput').value='';partChangesSelections.clear();
  writePersistent(PART_CHANGES_STORAGE.partChangesInput,'');previewPartChanges();
}

function partChangesPreviewValue(row,header,omit){
  return partChangesTemplateValue(row,header,omit,(row,header,omit)=>bobjTemplateValue({
    fid:row.Fid,quote:(row.order?.type||row.type)==='quotation'?row.Quote:'',customer:row.customer,systemDescription:row.systemDescription
  },header,omit));
}

function previewPartChangesFileNames(groups,showPendingFid=false){
  return plannedPartChangesFileNames(groups,bridgeId('partChangesTemplate').value.trim()||PART_CHANGES_DEFAULT_TEMPLATE,
    bridgeId('partChangesSeparator').value.trim()||ZBOM_COMPARE_DEFAULT_SEPARATOR,undefined,partChangesPreviewValue,showPendingFid);
}

function destroyPartChangesSystemPickers(container){
  if(typeof jQuery!=='function'||!jQuery.fn?.chosen)return;
  const pickers=jQuery(container).find('.part-changes-system-select');
  pickers.chosen('destroy');pickers.off('.partChangesSystems');
}

function partChangesSystemNames(){return GM_getValue(PART_CHANGES_MYLAM_SYSTEM_KEY,PART_CHANGES_MYLAM_SYSTEMS)}
function partChangesProduct(fid){return bobjValue(fid,'Primary Product')||bobjValue(fid,'FID Description')}

// Auto stays selected beside the Systems it suggests for the row. select.dataset.systems holds the saved value.
function partChangesSystemSelect(cell,subject,value,disabled,onChange,suggested=[]){
  const select=document.createElement('select');select.className='zbom-quote-select part-changes-system-select';
  select.multiple=true;select.disabled=disabled;select.dataset.systems=value;
  select.setAttribute('aria-label','FID System for '+subject);
  select.add(new Option('Auto',''));
  cell.classList.add('part-changes-system-cell');cell.append(select);
  const picker=typeof jQuery==='function'&&jQuery.fn?.chosen?jQuery(select):null;
  let shown=[],automatic=new Set();
  const paint=()=>{
    const saved=partChangesSystemList(select.dataset.systems),auto=saved.includes('');
    automatic=new Set(auto?suggested.filter(label=>!saved.includes(label)):[]);
    shown=auto?['',...automatic,...saved.filter(Boolean)]:saved;
    const listed=new Set(Array.from(select.options,option=>option.value));
    for(const label of shown)if(!listed.has(label))select.add(new Option(label,label));
    for(const option of select.options)option.selected=shown.includes(option.value);
    picker?.trigger('chosen:updated');
  };
  const loadOptions=()=>{
    const listed=new Set(Array.from(select.options,option=>option.value));
    for(const name of partChangesSystemNames())if(!listed.has(name))select.add(new Option(name,name));
    picker?.trigger('chosen:updated');
  };
  const change=()=>{
    const chosen=Array.from(select.selectedOptions,option=>option.value);
    // Removing a suggestion turns the remaining Systems into a manual selection.
    const auto=chosen.includes('')&&!shown.some(label=>automatic.has(label)&&!chosen.includes(label));
    select.dataset.systems=partChangesSystemValue(auto?['',...chosen.filter(label=>!automatic.has(label))]:chosen.filter(Boolean));
    paint();onChange(select);
  };
  paint();
  if(picker){
    picker.one('chosen:showing_dropdown.partChangesSystems',loadOptions).on('change.partChangesSystems',change);
    picker.chosen({width:'100%',search_contains:true,dropdown_position:'fixed',dropdown_width:'250px'});
  }else{
    select.addEventListener('focus',loadOptions,{once:true});
    select.addEventListener('change',change);
  }
  return select;
}

function renderPartChangesPreview(){
  const groups=partChangesGroups(partChangesRows),options=[...groups.keys()];let next=1;
  while(groups.has(zbomCompareGroupLabel(next)))next++;
  options.push(zbomCompareGroupLabel(next));options.sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  options.push(ZBOM_COMPARE_STANDALONE);
  const names=previewPartChangesFileNames(groups),previewNames=previewPartChangesFileNames(groups,true);
  const suggestSystems=createPartChangesSystemGuesser(partChangesSystemNames());
  const body=bridgeId('partChangesPreviewBody');destroyPartChangesSystemPickers(body);body.replaceChildren();
  for(const row of partChangesRows){
    const tr=document.createElement('tr');
    tr.insertCell().textContent=row.Fid||'—';
    const sales=tr.insertCell(),system=tr.insertCell(),group=tr.insertCell(),file=tr.insertCell(),status=tr.insertCell(),peers=groups.get(row.Group),standalone=row.Group===ZBOM_COMPARE_STANDALONE;
    const name=names.get(row.Group);
    if(!row.type&&!row.error){
      const select=document.createElement('select');select.className='zbom-quote-select';
      select.setAttribute('aria-label','Sales Order / Quotation for FID '+row.Fid);
      select.add(new Option('Latest SO (fallback latest quote)',PART_CHANGES_LATEST_SO));
      select.add(new Option('Latest quote in SAP',PART_CHANGES_LATEST_QUOTE));
      (row.options||[]).forEach(item=>select.add(new Option(partChangesDocumentLabel(item.type)+' '+item.number+' · '+item.version+' · '+item.created+
        (item.latest?' · Latest '+(item.type==='quotation'?'Quote':'SO')+' in SAP':''),partChangesDocumentKey(item.type,item.number))));
      select.value=row.selection;select.disabled=sapToolRunning('partChanges');
      select.title=row.order?partChangesDocumentLabel(row.order.type)+' '+row.order.number+' · '+row.order.version+' · '+row.order.created:'';
      select.addEventListener('change',()=>{partChangesSelections.set(row.key,{...partChangesSelections.get(row.key),document:select.value});previewPartChanges()});
      sales.append(select);
    }else sales.textContent=row.type?partChangesDocumentLabel(row.type)+' '+row.Input:'—';
    if(!row.error&&!standalone){
      partChangesSystemSelect(system,row.Fid?'FID '+row.Fid:row.Input,row.myLamSystem,sapToolRunning('partChanges'),select=>{
        row.myLamSystem=select.dataset.systems;
        partChangesSelections.set(row.key,{...partChangesSelections.get(row.key),myLamSystem:row.myLamSystem});
      },suggestSystems(partChangesPreviewValue(row,'System Description'),partChangesProduct(row.Fid)));
    }else system.textContent='—';
    if(!row.error){
      const select=compareGroupSelect('Group for FID '+row.Fid,row.Group,options,sapToolRunning('partChanges'));
      group.className='compare-group-cell';
      select.addEventListener('change',()=>{partChangesSelections.set(row.key,{...partChangesSelections.get(row.key),Group:select.value});previewPartChanges()});
      group.append(select);
    }
    file.className='compare-file-name';file.textContent=standalone?'SAP default filename':previewNames.get(row.Group)||'—';
    if(standalone)file.title='SAP saves this workbook in its default export folder.';
    const issue=row.error||(!standalone&&peers.length<2?'Add another document to '+row.Group+'.':
      !standalone&&row.documentKey&&peers.some(other=>other!==row&&other.documentKey===row.documentKey)?'This document is already in '+row.Group+'.':
      !row.options?'Lookup in SAP runs in Tasks.':!row.order?(row.lookupIssue||'No matching SAP document found.'):
      standalone?'':zbomFileNameError(name||'.xlsx','partChanges'));
    status.title=row.lookupIssue;
    status.textContent=issue||'Ready';status.className=issue?'row-warning':'row-ok';body.append(tr);
  }
  const standalone=partChangesRows.filter(row=>row.Group===ZBOM_COMPARE_STANDALONE).length;
  bridgeId('partChangesCount').textContent=partChangesRows.length+' rows · '+groups.size+' groups'+(standalone?' · '+standalone+' standalone':'');
  bridgeId('partChangesMyLamNote').hidden=Boolean(standalone&&!groups.size);
  updatePartChangesControls(names);
}

function updatePartChangesControls(names){
  const start=bridgeId('partChangesStart');if(!start)return;
  const ready=canStartSapTool('partChanges'),valid=partChangesRows.length>0&&partChangesRows.every(row=>!row.error);
  bridgeId('partChangesLookup').disabled=!canStartSapTool('partChanges',PART_CHANGES_LOOKUP_TOOL)||!valid;
  const groups=partChangesGroups(partChangesRows);
  const grouped=[...groups.values()].every(rows=>rows.length>=2&&new Set(rows.map(row=>row.documentKey)).size===rows.length);
  const folder=bridgeId('partChangesOutputFolder').value.trim();
  names||=previewPartChangesFileNames(groups);
  bridgeId('partChangesMyLamNote').textContent=!sapGuiControlUpdate.partChangesMyLam?'Update and restart SAPGUIcontrol to open MyLam automatically after export.':
    sapGuiControlUpdate.mock?'Mock mode records the MyLam launch. Use '+PART_CHANGES_ACTION_LABEL+' in Tasks to run the comparison.':
      'After SAP exports the BOMs, SAPGUIcontrol opens MyLam in your default browser. You can close this tab after submission.';
  start.disabled=!ready||(!pendingToolboxSubmissions.has('partChanges')&&
    (!valid||!grouped||!partChangesRows.every(row=>row.order||row.policy||row.type&&!row.options)||
     [...names.values()].some(name=>zbomFileNameError(name,'partChanges'))||groups.size&&(!sapGuiControlUpdate.partChangesMyLam||!/^[A-Za-z]:[\\/]/.test(folder))));
}

async function lookupPartChanges(){
  const inputs=[...new Set(partChangesRows.filter(row=>!row.error).map(row=>row.Input))];
  if(!inputs.length)return;
  const task=await fetchToolboxLookup('partChanges',PART_CHANGES_LOOKUP_TOOL,inputs.map(input=>partChangesDocumentType(input)?{Quote:input}:{Fid:input}));
  for(const row of task.results)if(row.Status==='Completed')partChangesLookups.set(row.Quote||row.Fid,JSON.parse(row.Message));
  previewPartChanges();
  if(task.state!=='Completed')throw new Error(task.error||'SAP could not finish the document lookup.');
  if(partChangesRows.some(row=>!row.order&&row.options))setSapProgress('partChanges','Review the SAP lookup results in Preview.','warning');
  else setSapProgress('partChanges','Select the documents and groups, then press Start.','ok');
}

function preparePartChangesTask(){
  const output=bridgeId('partChangesOutputFolder').value.trim().replace(/\//g,'\\');
  const Template=bridgeId('partChangesTemplate').value.trim()||PART_CHANGES_DEFAULT_TEMPLATE;
  if(!partChangesRows.length)throw new Error('Enter at least one document.');
  const invalid=partChangesRows.find(row=>row.error||!row.order&&!row.policy&&(!row.type||row.options));
  if(invalid)throw new Error(invalid.error||invalid.lookupIssue||'No matching SAP document found for '+invalid.Input+'.');
  const groups=partChangesGroups(partChangesRows),names=previewPartChangesFileNames(groups);
  if(groups.size&&!sapGuiControlUpdate.partChangesMyLam)throw new Error('Update and restart SAPGUIcontrol to open MyLam automatically.');
  const filenameError=[...names.values()].map(name=>zbomFileNameError(name,'partChanges')).find(Boolean);
  if(filenameError)throw new Error(filenameError);
  const deferredGroups=new Set([...groups].filter(([,members])=>members.some(row=>!row.order||row.policy)).map(([group])=>group));
  const filenameFields=[...new Set(['customer','system description',...bobjTemplateFields(Template).filter(field=>field!=='so')])];
  const rows=partChangesRows.map(row=>{
    const {Fid,Group,policy}=row,order=policy?null:row.order,deferred=deferredGroups.has(Group);
    const filenameValues=deferred?JSON.stringify(Object.fromEntries([...filenameFields,'today'].map(field=>[field,partChangesPreviewValue(row,field)]))):'';
    return {Fid,Quote:policy?'':row.Quote,Group,Version:order?.version||'',FileName:deferred?'':names.get(Group)||'',
      Fields:{input:row.Input,expectedFid:row.expectedFid,selection:row.selection,lookupPolicy:policy,documentType:order?.type||row.type,created:order?.created||'',
        customer:partChangesPreviewValue(row,'Customer'),system:partChangesPreviewValue(row,'System Description'),
        product:partChangesProduct(Fid),myLamSystem:row.myLamSystem,filenameValues}};
  });
  const spec={Id:crypto.randomUUID().replace(/-/g,''),Tool:'partChanges',OutputFolder:groups.size?output:'',Template,KeepTemp:groups.size>0,Rows:rows,
    Form:{partChangesInput:bridgeId('partChangesInput').value,partChangesThrough:bwCache.localDate(new Date()),
      partChangesSeparator:bridgeId('partChangesSeparator').value.trim()||ZBOM_COMPARE_DEFAULT_SEPARATOR},
    Workflow:createPartChangesWorkflow('partChanges',groups.size>0)};
  const batches=sapGuiControlUpdate.rowBatches&&partChangesRowBatches(rows);
  if(batches)spec.Form.rowBatches=batches;
  validatePartChangesTask(spec);
  return spec;
}

function partChangesTaskInput(rows){
  const cell=value=>'"'+String(value||'').replace(/\r\n?|\n/g,' ').replace(/"/g,'""')+'"';
  return ['Identifier\tFID\tGroup\tCustomer\tSystem Description',...rows.map(row=>
    [row.Fields.input||row.Fid||row.Quote,row.Fields.expectedFid,row.Group===ZBOM_COMPARE_STANDALONE?'Standalone':row.Group,row.Fields.customer,row.Fields.system].map(cell).join('\t'))].join('\n');
}

function restorePartChangesInputs(task){
  const input=bridgeId('partChangesInput');
  input.value=task.spec.Form?.partChangesInput||partChangesTaskInput(task.spec.Rows);
  partChangesSelections.clear();
  for(const [index,row] of parsePartChangesInput(input.value).entries()){
    const saved=task.spec.Rows[index],{document:resolved}=partChangesResultDetails(saved,task.results[index]);
    const document=resolved||(saved.Fields.created?{number:saved.Quote,version:saved.Version,created:saved.Fields.created,type:partChangesRowType(saved)}:null);
    const key=document?partChangesDocumentKey(document.type,document.number):'';
    // A saved latest result is only one candidate. A manual lookup still loads the full list.
    if(document){
      const lookup=partChangesLookups.get(row.Input)||{documents:[],issues:[]};
      if(!lookup.documents.some(item=>partChangesDocumentKey(item.type,item.number)===key))lookup.documents.push(document);
      partChangesLookups.set(row.Input,lookup);
    }
    partChangesSelections.set(row.key,{Group:saved.Group,document:saved.Fields.selection||key,myLamSystem:saved.Fields.myLamSystem});
  }
  const template=task.spec.Template||PART_CHANGES_DEFAULT_TEMPLATE;
  bridgeId('partChangesTemplate').value=template==='[BOM]'+ZBOM_COMPARE_DEFAULT_TEMPLATE?PART_CHANGES_DEFAULT_TEMPLATE:template;
  bridgeId('partChangesSeparator').value=task.spec.Form?.partChangesSeparator||ZBOM_COMPARE_DEFAULT_SEPARATOR;
  for(const id of ['partChangesInput','partChangesTemplate','partChangesSeparator'])writePersistent(PART_CHANGES_STORAGE[id],bridgeId(id).value);
  previewPartChanges();
}

function bindPartChangesEvents(){
  const defaults={partChangesInput:'',partChangesOutputFolder:PART_CHANGES_DEFAULT_FOLDER,partChangesTemplate:PART_CHANGES_DEFAULT_TEMPLATE,partChangesSeparator:ZBOM_COMPARE_DEFAULT_SEPARATOR};
  const update=id=>{
    writePersistent(PART_CHANGES_STORAGE[id],bridgeId(id).value);
    if(id==='partChangesInput')previewPartChanges();
    else renderPartChangesPreview();
    if(id==='partChangesInput'||id==='partChangesTemplate')void ensureBobjDataSource();
  };
  for(const [id,key] of Object.entries(PART_CHANGES_STORAGE)){
    const field=bridgeId(id);field.value=readPersistent(key,defaults[id]);
    if(id==='partChangesTemplate'&&field.value==='[BOM]'+ZBOM_COMPARE_DEFAULT_TEMPLATE){field.value=defaults[id];writePersistent(key,field.value)}
    field.addEventListener('input',()=>update(id));
  }
  bridgeId('partChangesReset').addEventListener('click',()=>{
    resetPartChangesInput();void ensureBobjDataSource();bridgeId('partChangesInput').focus();
  });
  bridgeId('partChangesResetTemplate').addEventListener('click',()=>{
    for(const id of ['partChangesTemplate','partChangesSeparator','partChangesOutputFolder']){bridgeId(id).value=defaults[id];update(id)}
    bridgeId('partChangesTemplate').focus();
  });
  document.querySelectorAll('[data-part-changes-token]').forEach(button=>button.addEventListener('click',()=>{
    const input=bridgeId('partChangesTemplate');input.setRangeText(button.dataset.partChangesToken,input.selectionStart,input.selectionEnd,'end');
    update(input.id);input.focus();
  }));
  bridgeId('partChangesLookup').addEventListener('click',()=>runSapTool('partChanges',lookupPartChanges));
  bridgeId('partChangesStart').addEventListener('click',()=>submitToolboxTask('partChanges'));
  previewPartChanges();
}




function waitMs(milliseconds){return new Promise(resolve=>setTimeout(resolve,milliseconds))}


function claimZbomOutputName(baseName,quote,claims=new Map()){
  let name=baseName,key=name.toLowerCase(),owner=claims.get(key);
  if(owner&&owner!==quote){name=appendQuoteToName(name,quote);key=name.toLowerCase();let suffix=2;while(claims.has(key)&&claims.get(key)!==quote){name=appendQuoteToName(baseName,quote+'_'+suffix++);key=name.toLowerCase()}}
  claims.set(key,quote);return name;
}








function bindBridgeEvents(){
  const taskHistory=bridgeId('toolboxTaskHistory');
  bridgeId('toolboxTaskMore').addEventListener('click',()=>{
    const list=bridgeId('toolboxTaskHistoryList');
    renderToolboxTaskList(list,toolboxTasks);taskHistory.showModal();list.scrollTop=0;
  });
  for(const [dialog,button,close] of [
    [bridgeId('toolboxTaskDialog'),bridgeId('toolboxTaskClose'),closeToolboxTask],
    [taskHistory,bridgeId('toolboxTaskHistoryClose'),()=>taskHistory.close()]
  ]){
    button.addEventListener('click',close);
    dialog.addEventListener('cancel',event=>{event.preventDefault();close()});
    dialog.addEventListener('click',event=>{
      if(event.target!==dialog)return;
      const {left,right,top,bottom}=dialog.getBoundingClientRect();
      if(event.clientX<left||event.clientX>right||event.clientY<top||event.clientY>bottom)close();
    });
  }
  setInterval(()=>{
    if(toolboxTaskData&&!toolboxTaskFinished(toolboxTaskData))paintToolboxTaskStatus(toolboxTaskData);
  },1000);
  bridgeId('bridgeRefresh').addEventListener('click',refreshToolbox);
  for(const [key,{id}] of Object.entries(bobjCards)){
    const picker=bridgeId(id+'ImportFile');
    bridgeId(id+'Update').addEventListener('click',()=>openBwLoadDialog(bwData,key,bobjCards[key].meta,range=>loadToolboxSource(key,undefined,range)));
    bridgeId(id+'Import').addEventListener('click',()=>picker.click());
    picker.addEventListener('change',()=>{const file=picker.files?.[0];picker.value='';if(file)loadToolboxSource(key,file)});
    bridgeId(id+'FileUpdate').addEventListener('click',()=>refreshBobjFile(key,true));
    bridgeId(id+'Open').addEventListener('click',()=>openBobjData(key));
  }
  document.addEventListener('click',event=>closeSourceMenus(event.target));
  document.querySelectorAll('[data-sap-lookup]').forEach(button=>button.addEventListener('click',()=>runSapTool(button.dataset.sapLookup,()=>lookupZbomSap(button.dataset.sapLookup))));
  window.addEventListener('beforeunload',event=>{if(Object.keys(sapTools).some(sapToolRunning)){event.preventDefault();event.returnValue=''}});
  document.querySelectorAll('[data-tool-tab]').forEach(button=>button.addEventListener('click',()=>{setToolTab(button.dataset.toolTab);ensureBobjDataSource()}));
  bridgeId('zbomInput').addEventListener('input',()=>{saveZbomInput();previewZbomInput();ensureBobjDataSource()});
  bridgeId('zbomTemplate').addEventListener('input',()=>{saveZbomTemplate();renderZbomPreview();ensureBobjDataSource()});
  document.querySelectorAll('[data-zbom-token]').forEach(button=>button.addEventListener('click',()=>insertZbomToken('zbom',button.dataset.zbomToken,()=>{saveZbomTemplate();renderZbomPreview()})));
  bridgeId('createQuoteQuotation').addEventListener('input',()=>{sanitizeCreateQuoteQuotation();saveCreateQuoteQuotation();updateCreateQuoteStart()});
  bridgeId('createQuoteInput').addEventListener('input',()=>{saveCreateQuoteInput();previewCreateQuoteInput();ensureBobjDataSource()});
  bridgeId('createQuoteDefaultValidTo').addEventListener('input',()=>{sanitizeCreateQuoteDays();saveCreateQuoteValidToDays();renderCreateQuotePreview()}); bridgeId('createQuoteFidSuffix').addEventListener('input',()=>{saveCreateQuoteFidSuffix();renderCreateQuotePreview()});
  bridgeId('zbomCompareInput').addEventListener('input',()=>{saveZbomCompareInput();previewZbomCompareInput(false);ensureBobjDataSource()});
  bridgeId('zbomCompareTemplate').addEventListener('input',()=>{saveZbomCompareTemplate();renderZbomComparePreview();ensureBobjDataSource()});
  bridgeId('zbomCompareSeparator').addEventListener('input',()=>{saveZbomCompareSeparator();renderZbomComparePreview()});
  bridgeId('zbomCompareOutputFolder').addEventListener('input',()=>{saveZbomCompareFolder();renderZbomComparePreview();updateZbomCompareStart()});
  bridgeId('zbomCompareKeepTemp').addEventListener('click',event=>{event.currentTarget.setAttribute('aria-checked',String(!keepZbomCompareTemp()))});
  document.querySelectorAll('[data-zbom-compare-token]').forEach(button=>button.addEventListener('click',()=>insertZbomToken('zbomCompare',button.dataset.zbomCompareToken,()=>{saveZbomCompareTemplate();renderZbomComparePreview()})));
  document.addEventListener('keydown',event=>{if(event.key==='Escape')closeSourceMenus()});
  bridgeId('zbomResetInput').addEventListener('click',()=>{const input=bridgeId('zbomInput');input.value='';zbomFileNames.zbom.clear();saveZbomFileNames('zbom');saveZbomInput();previewZbomInput();input.focus();});
  bridgeId('zbomResetTemplate').addEventListener('click',()=>{const input=bridgeId('zbomTemplate');input.value=ZBOM_DEFAULT_TEMPLATE;zbomFileNames.zbom.clear();saveZbomFileNames('zbom');saveZbomTemplate();renderZbomPreview();input.focus();});
  bridgeId('zbomCompareResetInput').addEventListener('click',()=>{const input=bridgeId('zbomCompareInput');input.value='';zbomFileNames.zbomCompare.clear();saveZbomFileNames('zbomCompare');saveZbomCompareInput();previewZbomCompareInput();input.focus();});
  bridgeId('zbomCompareResetTemplate').addEventListener('click',()=>{
    const input=bridgeId('zbomCompareTemplate'),separator=bridgeId('zbomCompareSeparator'),folder=bridgeId('zbomCompareOutputFolder');input.value=ZBOM_COMPARE_DEFAULT_TEMPLATE;separator.value=ZBOM_COMPARE_DEFAULT_SEPARATOR;folder.value=ZBOM_COMPARE_DEFAULT_FOLDER;
      zbomFileNames.zbomCompare.clear();saveZbomFileNames('zbomCompare');
      saveZbomCompareTemplate();saveZbomCompareSeparator();saveZbomCompareFolder();renderZbomComparePreview();updateZbomCompareStart();input.focus();
  });
  bridgeId('createQuoteResetInput').addEventListener('click',()=>{
    resetCreateQuoteInput();if(!pendingToolboxSubmissions.has('createQuote'))setSapProgress('createQuote','Ready.','idle');bridgeId('createQuoteInput').focus();
  });
  bridgeId('createQuoteResetSettings').addEventListener('click',()=>{
    const validToDays=bridgeId('createQuoteDefaultValidTo'),fidSuffix=bridgeId('createQuoteFidSuffix');validToDays.value=CREATE_QUOTE_DEFAULT_VALID_TO_DAYS;fidSuffix.value=CREATE_QUOTE_DEFAULT_FID_SUFFIX;saveCreateQuoteValidToDays();saveCreateQuoteFidSuffix();
      renderCreateQuotePreview();validToDays.focus();
  });
  bridgeId('zbomStart').addEventListener('click',()=>submitToolboxTask('zbom')); bridgeId('zbomCompareStart').addEventListener('click',()=>submitToolboxTask('zbomCompare'));
  bindPartChangesEvents();
  bridgeId('createQuoteStart').addEventListener('click',()=>submitToolboxTask('createQuote')); loadZbomPreferences();previewZbomInput();
    loadZbomComparePreferences();previewZbomCompareInput();loadCreateQuotePreferences();previewCreateQuoteInput();setToolTab('zbom');
}

} // SAP Toolbox

function splitQuotedLine(line,delimiter){
  const values=[];let value='',quoted=false;
  for(let index=0;index<line.length;index++){const character=line[index];if(character==='"'){if(quoted&&line[index+1]==='"'){value+='"';index++}else quoted=!quoted;}else if(character===delimiter&&!quoted){values.push(value);value=''}else value+=character;}
  values.push(value);return values;
}

// The task defines matching and layout. The bridge computes alignment and writes the workbook plan.
function createToolboxCompareTools(splitLine=splitQuotedLine){
  const ZBOM_COMPARE_FILE_COLORS=['FFFFC7CE','FFC6EFCE','FFFFEB9C','FFBDD7EE','FFE4DFEC','FFFCE4D6','FFDDEBF7','FFE2F0D9','FFF4B183','FFD9E1F2'];

  function parseZbomCompareTsvRows(text){
    return String(text||'').replace(/^\uFEFF/,'').split(/\r\n?|\n/).map(line=>{
      const values=line.includes('"')?splitLine(line,'\t'):line.split('\t');
      return values.map(value=>value.replace(/\u00a0/g,' ').trim());
    });
  }
  function extractZbomCompareFid(rows){const pattern=/Quotation\s*No\.?\s*[:#]?\s*(\d{6}(?:[-_][A-Za-z0-9]+)*)/i; for(const row of rows){const match=row.join(' ').match(pattern); if(match)return match[1];} return '';}

  function parseZbomCompareOptions(rows,gasStart){
    const options=[];let optionTypeColumn=-1,selectionColumn=-1,effectiveType=''; const end=gasStart>=0?gasStart:rows.length;
    for(let rowIndex=0;rowIndex<end;rowIndex++){
      const row=rows[rowIndex],typeHeader=row.findIndex(value=>normalizeZbomGasHeader(value)==='optiontype'),selectionHeader=row.findIndex(value=>normalizeZbomGasHeader(value)==='optionselection');
      if(typeHeader>=0&&selectionHeader>=0){optionTypeColumn=typeHeader+1;selectionColumn=selectionHeader+1;continue}
      if(optionTypeColumn<0||selectionColumn<0)continue;
      const optionType=String(row[optionTypeColumn]||'').trim(),selection=String(row[selectionColumn]||'').trim();
      if(!optionType&&!selection)continue;
      if(/^(?:Ref\s*#|Page\s+\d+|Quotation\s*No\.?)/i.test(optionType+' '+selection))continue;
      if(optionType)effectiveType=optionType;
      options.push({optionType:optionType, effectiveType:effectiveType||optionType, selection:selection, comment:'', key:(effectiveType||optionType)+'\u0000'+selection});
    }
    return options;
  }

  function normalizeZbomGasHeader(value){return String(value||'').trim().toLowerCase().replace(/[^a-z0-9]+/g,'');}

  function canonicalZbomGasHeader(value){
    const key=normalizeZbomGasHeader(value);
    if(key==='module'||key==='processmodule'||key==='pm')return {compareKey:'module',label:'Module'};
    if(key==='line'||key==='linenumber'||key==='lineno')return {compareKey:'line',label:'Line'};
    if(key==='mfctype'||key==='mfc')return {compareKey:'mfctype',label:'MFC Type'};
    if(key==='gastype'||key==='gas')return {compareKey:'gastype',label:'Gas Type'};
    if(key==='gasflow'||key==='flow')return {compareKey:'gasflow',label:'Gas Flow'};
    return {compareKey:key,label:String(value||'').trim()};
  }

  function firstZbomGasValue(row,start,end){const last=Math.min(row.length-1,Math.max(start,end)); for(let column=start;column<=last;column++){const value=String(row[column]||'').trim();if(value)return value;} return '';}

  function findZbomCompareGasHeaderRow(rows,gasStart){
    if(gasStart<0)return -1;
    for(let rowIndex=gasStart+1;rowIndex<rows.length;rowIndex++){const headers=rows[rowIndex].map(normalizeZbomGasHeader); if(headers.includes('line')&&headers.includes('mfctype')&&headers.includes('gastype')&&headers.includes('gasflow'))return rowIndex;}
    return -1;
  }

  function parseZbomCompareGas(rows,gasStart){
    if(gasStart<0)return null;
    const headerRow=findZbomCompareGasHeaderRow(rows,gasStart);
    if(headerRow<0)return null;
    const sourceHeaders=rows[headerRow].map((value,column)=>({value:String(value||'').trim(),column:column})).filter(item=>item.value); const seenKeys=new Map(),headers=[];
    sourceHeaders.forEach((item,index)=>{
      const canonical=canonicalZbomGasHeader(item.value),count=(seenKeys.get(canonical.compareKey)||0)+1,key=canonical.compareKey+(count>1?'#'+count:''); seenKeys.set(canonical.compareKey,count);
      headers.push({key:key,compareKey:canonical.compareKey,label:canonical.label,start:item.column,end:index+1<sourceHeaders.length?sourceHeaders[index+1].column-1:Math.max(item.column,...rows.slice(headerRow+1).map(row=>row.length-1))});
    });
    if(!headers.some(header=>header.compareKey==='module'))headers.unshift({key:'module',compareKey:'module',label:'Module',start:-1,end:-1,synthetic:true});
    const lineHeader=headers.find(header=>header.compareKey==='line');
    if(!lineHeader)throw new Error('Gas section was found, but its Line column could not be identified.');
    const records=[];let currentModule='';
    for(let rowIndex=headerRow+1;rowIndex<rows.length;rowIndex++){
      const row=rows[rowIndex],joined=row.join(' ').trim();
      if(!joined)continue;
      if(/(?:Ref\s*#|Page\s+\d+|Quotation\s*No\.?)/i.test(joined))continue;
      const line=firstZbomGasValue(row,lineHeader.start,lineHeader.end);
      if(!line){const moduleCandidate=row.map(value=>String(value||'').trim()).find(value=>/^(?:PM|MODULE)\s*[-_A-Z0-9]+$/i.test(value)); if(moduleCandidate)currentModule=moduleCandidate; continue;}
      const values={}; headers.forEach(header=>{values[header.key]=header.synthetic?currentModule:firstZbomGasValue(row,header.start,header.end);}); const moduleHeader=headers.find(header=>header.compareKey==='module'),moduleValue=String(values[moduleHeader.key]||currentModule).trim();
      if(moduleValue)currentModule=moduleValue;
      values[moduleHeader.key]=currentModule; records.push({values:values,key:currentModule+'\u0000'+line,module:currentModule,line:line});
    }
    return {headers:headers,records:records};
  }

  function parseZbomCompareTsv(text,quote){
    const rows=parseZbomCompareTsvRows(text),fid=extractZbomCompareFid(rows),gasMarker=rows.findIndex(row=>row.some(value=>normalizeZbomGasHeader(value).includes('gasboxconfiguration'))),gasStart=findZbomCompareGasHeaderRow(rows,gasMarker)>=0?gasMarker:-1;
    const options=parseZbomCompareOptions(rows,gasStart),gas=parseZbomCompareGas(rows,gasStart);
    if(!options.length)throw Object.assign(new Error('No option rows were found in quotation '+quote+'.'),{rowError:true});
    return {fid,options,gas,warning:''};
  }

  // Keep two score rows and one byte per traceback step; both alignment paths use the same tie rules.
  function alignZbomRows(left,right,gap,mismatch){
    const rowCount=left.length,columnCount=right.length,directions=Array.from({length:rowCount+1},()=>new Uint8Array(columnCount+1));
    let previous=new Float64Array(columnCount+1),current=new Float64Array(columnCount+1);
    for(let column=1;column<=columnCount;column++)previous[column]=column*gap;
    for(let row=1;row<=rowCount;row++){
      current[0]=row*gap;
      for(let column=1;column<=columnCount;column++){
        let score=mismatch;
        for(const key of left[row-1])score=Math.max(score,right[column-1][key]??mismatch);
        const diagonal=previous[column-1]+score,above=previous[column]+gap,across=current[column-1]+gap;
        current[column]=Math.max(diagonal,above,across);
        directions[row][column]=diagonal===current[column]&&(score>0||(diagonal>above&&diagonal>across))?1:above===current[column]?2:3;
      }
      [previous,current]=[current,previous];
    }
    const actions=[];let row=rowCount,column=columnCount;
    while(row||column){
      const direction=row&&column?directions[row][column]:row?2:3;
      if(direction===1){actions.push(['match',row-1,column-1]);row--;column--}
      else if(direction===2){actions.push(['missing',row-1,-1]);row--}
      else{actions.push(['insert',-1,column-1]);column--}
    }
    return actions.reverse();
  }
  function alignZbomCompareRecordSets(recordSets,keyOf,similarKeys,align){
    if(!recordSets.length)return [];
    const keys=record=>['exact:'+keyOf(record),...similarKeys(record)];
    let slots=(recordSets[0]||[]).map(record=>({records:[record],keys:new Set(keys(record))}));
    for(let fileIndex=1;fileIndex<recordSets.length;fileIndex++){
      const incoming=recordSets[fileIndex]||[];
      if(!incoming.length)continue;
      const incomingKeys=incoming.map(keys),left=slots.map(slot=>[...slot.keys]);
      const right=incomingKeys.map(values=>Object.fromEntries(values.map((key,index)=>[key,index?4:12])));
      slots=align(left,right,-2,-5).map(([action,row,column])=>{
        if(action==='insert')return {records:Array(fileIndex).fill(null).concat(incoming[column]),keys:new Set(incomingKeys[column])};
        const slot=slots[row];
        while(slot.records.length<fileIndex)slot.records.push(null);
        slot.records.push(action==='match'?incoming[column]:null);
        if(action==='match')for(const key of incomingKeys[column])slot.keys.add(key);
        return slot;
      });
    }
    slots.forEach(slot=>{while(slot.records.length<recordSets.length)slot.records.push(null)});return slots;
  }

  function zbomCompareFileColor(index){
    if(index<ZBOM_COMPARE_FILE_COLORS.length)return ZBOM_COMPARE_FILE_COLORS[index];
    const hue=(index*137.508)%360,saturation=.48,lightness=.84,component=value=>{const amount=saturation*Math.min(lightness,1-lightness),part=(value+hue/30)%12,color=lightness-amount*Math.max(-1,Math.min(part-3,9-part,1)); return Math.round(255*color).toString(16).padStart(2,'0');};
    return 'FF'+component(0)+component(8)+component(4);
  }

  function normalizedZbomCompareCell(value){return String(value===null||value===undefined?'':value).trim()}
  function zbomCompareOptionKey(option){return [option.optionType,option.selection,option.comment].map(normalizedZbomCompareCell).join('\u0000');}

  function zbomCompareOptionSimilarity(option){
    const keys=[],selection=normalizedZbomCompareCell(option.selection),type=normalizedZbomCompareCell(option.effectiveType);
    if(selection)keys.push('selection:'+selection);
    if(type)keys.push('type:'+type);
    return keys;
  }

  function isZbomCompareSectionOption(option){const title=normalizedZbomCompareCell(option&&option.optionType),selection=normalizedZbomCompareCell(option&&option.selection); return Boolean(title&&!selection&&/[A-Z]/.test(title)&&title===title.toUpperCase());}

  function splitZbomOptionSections(options){
    const sections=[];let section=null;
    (options||[]).forEach(option=>{
      if(isZbomCompareSectionOption(option)){const title=normalizedZbomCompareCell(option.optionType).replace(/\s+/g,' ').toUpperCase(); section={key:'SECTION:'+title,records:[]};sections.push(section);}else if(!section){section={key:'PREAMBLE',records:[],preamble:true};
        sections.push(section)}
      section.records.push(option);
    });
    return sections;
  }

  function scoreZbomSectionContent(left,right){
    const records=section=>section.records.slice(section.preamble?0:1).map(zbomCompareOptionKey),leftKeys=records(left),rightKeys=records(right),counts=new Map(); leftKeys.forEach(key=>counts.set(key,(counts.get(key)||0)+1));let matches=0;
    rightKeys.forEach(key=>{const count=counts.get(key)||0;if(count){matches++;counts.set(key,count-1)}}); return leftKeys.length+rightKeys.length?Math.round(200*matches/(leftKeys.length+rightKeys.length)):100;
  }

  function scoreZbomCompareSectionMatch(slot,section){const existing=slot.sections.filter(Boolean).filter(item=>item.key===section.key); if(!existing.length)return -1; return 1000+Math.max(...existing.map(item=>scoreZbomSectionContent(item,section)));}

  function alignZbomOptionSections(optionSets,align){
    if(!optionSets.length)return [];
    let slots=splitZbomOptionSections(optionSets[0]).map(section=>({sections:[section]}));
    for(let fileIndex=1;fileIndex<optionSets.length;fileIndex++){
      const incoming=splitZbomOptionSections(optionSets[fileIndex]),assignments=new Map(),usedSlots=new Set(),slotsByTitle=new Map();
      slots.forEach((slot,index)=>{
        while(slot.sections.length<=fileIndex)slot.sections.push(null);
        const title=slot.sections.find(Boolean).key;
        if(!slotsByTitle.has(title))slotsByTitle.set(title,[]);
        slotsByTitle.get(title).push(index);
      });
      const scores=incoming.map(section=>Object.fromEntries((slotsByTitle.get(section.key)||[]).map(index=>[index,scoreZbomCompareSectionMatch(slots[index],section)])));
      // Preserve the original Section order wherever an ordered match is possible.
      for(const [action,slotIndex,incomingIndex] of align(slots.map((_,index)=>[String(index)]),scores,0,-1)){
        if(action!=='match')continue;
        slots[slotIndex].sections[fileIndex]=incoming[incomingIndex];assignments.set(incomingIndex,slots[slotIndex]);usedSlots.add(slotIndex);
      }

      // Only unmatched Sections with the same title may move to a different position.
      const moved=[];
      scores.forEach((candidates,incomingIndex)=>{
        if(assignments.has(incomingIndex)||incoming[incomingIndex].preamble)return;
        for(const [index,score] of Object.entries(candidates)){
          const slotIndex=Number(index);
          if(!usedSlots.has(slotIndex))moved.push({slot:slots[slotIndex],slotIndex,incomingIndex,score,distance:Math.abs(slotIndex-incomingIndex)});
        }
      });
      moved.sort((left,right)=>right.score-left.score||left.distance-right.distance||left.slotIndex-right.slotIndex||left.incomingIndex-right.incomingIndex);
      moved.forEach(match=>{if(usedSlots.has(match.slotIndex)||assignments.has(match.incomingIndex))return; match.slot.sections[fileIndex]=incoming[match.incomingIndex];assignments.set(match.incomingIndex,match.slot);usedSlots.add(match.slotIndex);});

      // Keep genuinely new Sections whole and place them beside their nearest ordered match.
      incoming.forEach((section,incomingIndex)=>{
        if(assignments.has(incomingIndex))return;
        const newSlot={sections:Array(fileIndex).fill(null).concat(section)};let anchor=null;
        for(let previous=incomingIndex-1;previous>=0&&!anchor;previous--)anchor=assignments.get(previous)||null;
        if(anchor)slots.splice(slots.indexOf(anchor)+1,0,newSlot); else{for(let next=incomingIndex+1;next<incoming.length&&!anchor;next++)anchor=assignments.get(next)||null; slots.splice(anchor?slots.indexOf(anchor):slots.length,0,newSlot);}
        assignments.set(incomingIndex,newSlot);
      });
    }
    return slots.flatMap(slot=>alignZbomCompareRecordSets(slot.sections.map(section=>section?section.records:[]),zbomCompareOptionKey,zbomCompareOptionSimilarity,align));
  }

  function plan(models,grouped,align=alignZbomRows){
    const sheets=[],headerFill='FF0070C0';
    const colors=models.map((_,index)=>zbomCompareFileColor(index));
    function add(value){
      const base=String(value||'Sheet').replace(/[\[\]:*?/\\]/g,',').replace(/^'+|'+$/g,'').trim()||'Sheet';
      let name=base.slice(0,31),number=2;
      while(sheets.some(sheet=>sheet.Name.toLowerCase()===name.toLowerCase())){
        const suffix=' ('+number+++')';name=base.slice(0,31-suffix.length)+suffix;
      }
      const sheet={Name:name,Cells:new Map(),Blocks:[],Differences:[],Widths:[],Merges:[],FilterRange:'',FrozenRows:2,ShowGridLines:false};
      sheets.push(sheet);return sheet;
    }
    function cell(sheet,Row,Column,text='',Header=false,Bold=Header){
      const value={Row,Column,Text:String(text??''),Header,Bold,Fill:Header?headerFill:''};
      sheet.Cells.set(Row+':'+Column,value);return value;
    }
    function optionBlock(sheet,model,start,rows){
      const labels=['Option Type','Option Selection','Comment'],fields=['optionType','selection','comment'],widths=[30,38,22];
      labels.forEach((label,index)=>{
        cell(sheet,1,start+index,index===0?'FID No.':index===1?model.fid:'',true);
        cell(sheet,2,start+index,label,true);
        // Cells stay on one line, so a column grows with its longest text up to 60.
        const longest=rows.reduce((length,record)=>Math.max(length,String(record?.[fields[index]]??'').length),0);
        sheet.Widths[start+index-1]=Math.max(widths[index],Math.min(60,longest+2));
      });
      sheet.Widths[start+2]=3;
      sheet.Blocks.push({Row:3,Column:start,Values:rows.map(record=>[record?.optionType,record?.selection,record?.comment]),
        BoldRows:rows.map(isZbomCompareSectionOption)});
    }
    function gasBlock(sheet,model,start,rows){
      model.gas.headers.forEach((header,index)=>{
        cell(sheet,1,start+index,index===0?'FID No.':index===1?model.fid:'',true);
        cell(sheet,2,start+index,header.label,true);
        sheet.Widths[start+index-1]=Math.max(11,Math.min(28,header.label.length+4));
      });
      sheet.Blocks.push({Row:3,Column:start,Values:rows.map(record=>model.gas.headers.map(header=>record?.values[header.key]))});
    }
    function compareGas(){
      const available=models.map((model,file)=>({model,file})).filter(item=>item.model.gas);
      if(!available.length)return;
      const sheet=add('Compare_Gas'),starts=[];
      const slots=alignZbomCompareRecordSets(available.map(item=>item.model.gas.records),row=>row.key,row=>row.module?['module:'+row.module]:[],align);
      let start=1;
      models.forEach((model,file)=>{
        starts[file]=start;
        const width=model.gas?Math.max(2,model.gas.headers.length):3;
        for(let column=0;column<width;column++)cell(sheet,1,start+column,column===0?'FID No.':column===1?model.fid:'',true);
        if(model.gas){
          const index=available.findIndex(item=>item.file===file);
          gasBlock(sheet,model,start,slots.map(slot=>slot.records[index]));
        }else{
          for(let column=0;column<width;column++)cell(sheet,2,start+column,column===0?'No Gas':'',true).Fill=colors[file];
          sheet.Merges.push([2,start,2,start+width-1]);
          [11,14,5].forEach((value,index)=>{sheet.Widths[start+index-1]=value});
        }
        sheet.Widths[start+width-1]=3;start+=width+1;
      });
      const common=[...new Set(available[0].model.gas.headers.map(header=>header.compareKey))]
        .filter(key=>available.every(item=>item.model.gas.headers.some(header=>header.compareKey===key)));
      for(const name of common)sheet.Differences.push({Row:3,Count:slots.length,Colors:available.map(item=>colors[item.file]),
        Columns:available.map(item=>starts[item.file]+item.model.gas.headers.findIndex(header=>header.compareKey===name))});
    }
    if(grouped&&models.length>1){
      const sheet=add('Compare'),slots=alignZbomOptionSections(models.map(model=>model.options),align);
      models.forEach((model,file)=>optionBlock(sheet,model,file*4+1,slots.map(slot=>slot.records[file])));
      for(let column=1;column<=3;column++)sheet.Differences.push({Row:3,Count:slots.length,Columns:models.map((_,file)=>file*4+column),Colors:colors});
      compareGas();
    }
    for(const model of models){
      const options=add(model.fid);optionBlock(options,model,1,model.options);options.FilterRange='A2:C2';
      if(model.gas){
        const sheet=add(model.fid+'_Gas');gasBlock(sheet,model,1,model.gas.records);
        model.gas.headers.forEach((header,index)=>{
          const longest=model.gas.records.reduce((width,record)=>Math.max(width,record.values[header.key].length),header.label.length);
          sheet.Widths[index]=Math.max(11,Math.min(35,longest+2));
        });
      }
    }
    return {BorderColor:'FFD6E0EA',HeaderFontColor:'FFFFFFFF',Sheets:sheets.map(sheet=>({...sheet,Cells:[...sheet.Cells.values()]}))};
  }

  return {parse:parseZbomCompareTsv,plan,align:alignZbomRows};
}
function zbomFileNameError(name,tool){
  if(!name||/^[. ]+$/.test(name))return 'Enter a file name.';
  if(/[<>:"/\\|?*\u0000-\u001F]/.test(name)||/[. ]$/.test(name))return 'Use a file name without folder paths or these characters: < > : " / \\ | ? *';
  if(/^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)/i.test(name))return 'This name is reserved by Windows. Choose another name.';
  const extension=tool==='zbom'?'.pdf':'.xlsx';
  return name.length>255?'Keep the file name within 255 characters.':!name.toLowerCase().endsWith(extension)?'Use the '+extension+' file extension.':'';
}

function compactZbomCompareName(value){return String(value||'').trim().replace(/_{2,}/g,'_').replace(/^[\s_.-]+|[\s_.-]+$/g,'');}

function sanitizeToolboxFileName(value,replacement,pendingFid){
  const invalid=/[<>:"/\\|?*\u0000-\u001F]/g;
  if(!pendingFid)return value.replace(invalid,replacement);
  return value.split(/(\{[^{}]+\})/).map((part,index)=>index%2?part:part.replace(invalid,replacement)).join('');
}

function finishZbomCompareFileName(value,fallback='ZBOM Compare',pendingFid){
  const safe=sanitizeToolboxFileName(compactZbomCompareName(value),',',pendingFid).replace(/\.xlsx$/i,'');
  return (safe||fallback)+'.xlsx';
}
function zbomCompareDuplicateKey(value){return String(value||'').trim().toLocaleLowerCase()}

function joinCompareFileNames(rows,separator,field,render){
  const seenCustomers=new Set(),seenDescriptions=new Set(),parts=[];
  rows.forEach(row=>{
    const customerKey=zbomCompareDuplicateKey(field(row,'Customer')),descriptionKey=zbomCompareDuplicateKey(field(row,'System Description'));
    parts.push(render(row,{customer:Boolean(customerKey&&seenCustomers.has(customerKey)),systemDescription:Boolean(descriptionKey&&seenDescriptions.has(descriptionKey))}));
    if(customerKey)seenCustomers.add(customerKey);
    if(descriptionKey)seenDescriptions.add(descriptionKey);
  });
  return parts.join(' '+separator+' ');
}

// Task-local JavaScript: serialized at submission and executed without the browser.
function toolboxWorkflowOperations(){
  return ['sap','sapScript','progress','result','ask','path','mkdir','files','waitFiles','watchFiles','inspectFile','readFile','publish','mockFiles','writeWorkbook'];
}

function toolboxLatestQuoteFid(quote){
  return /^sap-latest:(\d{6})$/.exec(quote)?.[1]||'';
}

function createToolboxWorkflow(tool,outputFolder){
  if(tool==='partChanges'||tool==='partChangesLookup')return createPartChangesWorkflow(tool);
  const lookupOnly=tool==='zbomLookup';
  const run={zbom:toolboxPdfWorkflow,zbomCompare:toolboxCompareWorkflow,createQuote:toolboxCreateQuoteWorkflow,zbomLookup:resolveToolboxQuotes}[tool];
  if(!run)throw new Error('Unknown task workflow: '+tool);
  const dates='const workflowDate=('+toolboxWorkflowDate.toString()+');\n';
  const selection=tool==='zbom'?'const toolboxZbomSelectionSource='+JSON.stringify(toolboxSapScriptSource(toolboxZbomSelection))+';\n':'';
  const compare=tool==='zbomCompare'?'const toolboxCompareStageSource='+JSON.stringify('"use strict";\n'+
    'const compare=('+createToolboxCompareTools.toString()+')('+splitQuotedLine.toString()+');\n'+
    '('+toolboxCompareStage.toString()+')(JSON.parse(inputJson),bridge,compare);')+';\n'+
    'const toolboxCompareExportSource='+JSON.stringify(toolboxSapScriptSource(toolboxCompareExport))+';\n':'';
  const latestKey=toolboxLatestQuoteFid.toString()+'\n';
  const lookup=tool==='createQuote'?'':latestKey+(lookupOnly?'':resolveToolboxQuotes.toString()+'\n')+
    'const toolboxSapQuoteSource='+JSON.stringify(toolboxSapScriptSource(toolboxLookupSapQuote,toolboxSearchSapVersions))+';\n';
  return {Api:1,KeepWindow:tool==='createQuote',CompletionMessage:tool==='createQuote'?'Items entered. SAP window kept open for review.':'Completed',
    // SAP's PDF download needs a short path. The bridge allocates a separate directory per task.
    OutputWorkFolder:tool==='zbom'?'.kyra':'',
    AfterSuccess:lookupOnly||tool==='createQuote'?[]:[{Operation:'openDirectory',Path:outputFolder,Activate:true}],
    Operations:lookupOnly?['sap','sapScript','sapBatch','progress','result']:
      tool==='createQuote'?toolboxWorkflowOperations():[...toolboxWorkflowOperations(),'sapBatch',...(tool==='zbomCompare'?['alignRows','scriptStage']:[])],
    Validation:'"use strict";\n'+dates+latestKey+'('+validateToolboxWorkflowInput.toString()+')(JSON.parse(inputJson));',
    Source:'"use strict";\n'+dates+selection+compare+lookup+'('+toolboxWorkflowRuntime.toString()+')(JSON.parse(contextJson),bridge,('+run.toString()+'));'};
}

function toolboxSapScriptSource(run,...dependencies){
  return dependencies.map(helper=>helper.toString()+'\n').join('')+'('+toolboxSapScope.toString()+')(inputJson,sap,('+run.toString()+'));';
}

function toolboxWorkflowDate(value){
  const text=String(value).trim();let parts=text.match(/^(\d{4})-(\d{2})-(\d{2})$/)||text.match(/^(\d{4})(\d{2})(\d{2})$/),year,month,day;
  if(parts){[,year,month,day]=parts}else if((parts=text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))){[,month,day,year]=parts}
  else if((parts=text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/))){[,day,month,year]=parts}else throw new Error('Use a valid date with a four-digit year.');
  const date=new Date(Date.UTC(Number(year),Number(month)-1,Number(day)));
  if(date.getUTCFullYear()!==Number(year)||date.getUTCMonth()+1!==Number(month)||date.getUTCDate()!==Number(day))throw new Error('Use a valid date with a four-digit year.');
  return date.toISOString().slice(0,10);
}

function validateToolboxWorkflowInput(spec){
  const create=spec.Tool==='createQuote',lookup=spec.Tool==='zbomLookup',seen=new Set();
  for(const row of spec.Rows){
    if(!create&&!lookup&&row.Fields?.inputError)continue;
    const key=create?row.Fid:row.Quote;if(seen.has(key))throw new Error('Remove duplicate task rows.');seen.add(key);
    if(create){
      if(!/^\d{6}$/.test(row.Fid))throw new Error('A FID must have six digits.');
      if(row.Crd)workflowDate(row.Crd);workflowDate(row.ValidTo);
    }else{
      const latestFid=toolboxLatestQuoteFid(row.Quote);
      if(latestFid){
        if(!/^\d{6}(?:-[A-Za-z0-9-]+)?$/.test(row.Fid)||!row.Fid.startsWith(latestFid))throw new Error('Latest SAP quote requires a matching FID.');
      }else if(!/^\d{8}$/.test(row.Quote))throw new Error('A quotation must have eight digits.');
      if(lookup)continue;
      const extension=spec.Tool==='zbom'?'.pdf':'.xlsx';
      if(!row.FileName.toLowerCase().endsWith(extension)||/[<>:"/\\|?*\u0000-\u001F]/.test(row.FileName))throw new Error('Invalid planned filename.');
    }
  }
  if(create){if(!/^\d{8,10}$/.test(spec.Form.createQuoteQuotation))throw new Error('Enter a valid reference quotation.')}
  else if(!lookup&&!/^[A-Za-z]:[\\/]/.test(spec.OutputFolder))throw new Error('Use an absolute local output folder.');
}

function toolboxWorkflowRuntime(context,bridge,run){
  function invoke(op,args={}){
    const reply=JSON.parse(bridge(JSON.stringify({op,...args})));
    if(!reply.ok){const error=new Error(reply.error.message);error.code=reply.error.code;throw error}
    return reply.value;
  }
  const fidParts=/^(\d{6})((?:[-_][A-Za-z0-9]+)*)$/;
  const deferredDescription=row=>!row.FileNameEdited&&row.Fields?.['system description']==='';
  let fidSource;
  function exportedFidFields(fid){
    const key=/^[0-9]{6}(?![0-9])/.exec(fid)?.[0],snapshot=context.spec.Form?.exportFidSource;
    if(!key||!snapshot)return {};
    const [fields,rows]=fidSource||(fidSource=JSON.parse(snapshot)),values=rows[key];
    return values?Object.fromEntries(fields.map((field,index)=>[field,values[index]]).filter(([,value])=>value)):{};
  }
  const api={
    invoke,
    sap:(operation,control='',value='',options={})=>invoke('sap',{operation,control,value:String(value??''),dateFormats:[],...options}),
    sapScript:(source,input)=>invoke('sapScript',{source,input}),
    sapBatch:(message,transaction,source,jobs)=>jobs.length?invoke('sapBatch',{message,transaction,source,jobs}):{results:[]},
    lookupQuote:input=>{
      api.sap('transaction','','VA23');
      return api.sapScript(toolboxSapQuoteSource,{...input,reportRowError:true});
    },
    progress:message=>invoke('progress',{message}),
    result:(index,status,message='',file,fid)=>invoke('result',{index,status,message,...file!==undefined?{file}:{},...fid!==undefined?{fid}:{}}),
    fail:(index,row,message)=>{
      api.result(index,'Failed',[row.SapLookup,message].filter(Boolean).join(' / '));
      row.Error=message;return row;
    },
    ask:message=>invoke('ask',{message}),
    path:(...parts)=>invoke('path',{parts}),
    mkdir:path=>invoke('mkdir',{path}),
    files:(directory,pattern='*',recursive=false)=>invoke('files',{directory,pattern,recursive}),
    waitFiles:(directory,pattern,recursive,required=[])=>invoke('waitFiles',{directory,pattern,recursive,required}),
    inspect:path=>invoke('inspectFile',{path}),
    publish:(source,destination)=>invoke('publish',{source,destination}),
    mockFiles:files=>invoke('mockFiles',{files}),
    mockFid:row=>{
      const fid=row.Version||row.Fid||'300000';
      return /^\d{6}$/.test(fid)?fid+'-01':fid;
    },
    focus:control=>api.sap('sapCall',control,'',{steps:[{Action:'invoke',Member:'SetFocus',Arguments:[]}]}),
    read:(control,optional=false)=>String(api.sap('readText',control,'',{optional}).value??''),
    usesFidVersion:/\{\s*fidver\s*\}/i.test(context.spec.Template),
    deferredDescription,
    exportedFid:(row,observed,lookupVersion,lookupError)=>{
      const sapFid=fidParts.test(row.Version)?row.Version:'';
      if(row.SapLookup&&fidParts.exec(sapFid)?.[2])return sapFid;
      if(fidParts.exec(observed)?.[2])return observed;
      if(sapFid&&(!observed||sapFid.slice(0,6)===observed.slice(0,6)))return sapFid;
      if(lookupVersion&&!sapFid&&!row.SapLookup){
        if(lookupError)throw new Error(lookupError);
        api.progress('Reading the FID version for quotation '+row.Quote);
        const version=api.lookupQuote({quoteForVersion:row.Quote});
        if(fidParts.test(version))return version;
      }
      return observed;
    },
    renderName:(row,fid,customers,descriptions)=>{
      const [,base='',fidver='']=fidParts.exec(fid)||[];
      const fields=Object.fromEntries(Object.entries(row.Fields||{}).map(([field,value])=>[field.toLowerCase(),value]));
      for(const [field,value] of Object.entries(exportedFidFields(fid)))if(!fields[field]||base!==row.Fid.slice(0,6))fields[field]=value;
      Object.assign(fields,{fid:base,fidver,quote:row.Quote});
      for(const [field,seen] of [['customer',customers],['system description',descriptions]])if(seen&&fields[field]){
        const key=fields[field].toLowerCase();if(seen.has(key))fields[field]='';else seen.add(key);
      }
      return context.spec.Template.replace(/\{([^{}]+)\}/g,(_,field)=>fields[field.trim().replace(/\s+/g,' ').toLowerCase()]||'');
    },
    exportedFidMessage:(row,fid)=>{
      const fields=exportedFidFields(fid),missingDescription=deferredDescription(row)&&!fields['system description']?'System Description unavailable.':'';
      if(row.Fid)return missingDescription;
      if(!fid)return 'The SAP export did not provide a FID.';
      const details=Object.entries(fields).map(([field,value])=>field[0].toUpperCase()+field.slice(1)+' '+value);
      const note=details.length?'Matched after export · '+details.join(' · '):'FID read from SAP. Details are unavailable in Data Sources.';
      return [note,missingDescription].filter(Boolean).join(' ');
    },
    date:workflowDate
  };
  run(context,api);
}

function toolboxSapScope(inputJson,sap,run){
  function request(command){
    const reply=JSON.parse(sap(JSON.stringify(command)));
    if(!reply.ok){const error=new Error(reply.error.message);error.code=reply.error.code;throw error}
    return reply.value;
  }
  const argument=value=>({Type:typeof value==='number'?'int':'string',Value:String(value)});
  const call=(control,step,via=[])=>request({Operation:'sapCall',ControlId:control,Steps:[
    {Action:'raw'},...via.map(Member=>({Action:'follow',Member})),...[].concat(step)
  ]});
  const api={
    call,
    state:()=>request({Operation:'state'}),
    properties:(control,names,via)=>call(control,{Action:'properties',Arguments:names.map(argument)},via),
    children:(control,names)=>call(control,{Action:'children',Arguments:names.map(argument)}),
    set:(control,Member,value,via)=>call(control,{Action:'set',Member,Value:argument(value)},via),
    press:control=>call(control,{Action:'invoke',Member:'Press'}),
    failRow(message,cancelDialog=false){
      // A recoverable document error must leave the guarded SAP command in a known state.
      const {transaction,window}=api.state();
      if(cancelDialog&&window==='wnd[1]')call(window,{Action:'invoke',Member:'SendVKey',Arguments:[argument(12)]});
      const ready=api.state();
      if(ready.window!=='wnd[0]'||ready.transaction!==transaction)throw new Error(message);
      call('',{Action:'invoke',Member:'StartTransaction',Arguments:[argument(transaction)]});
      const reset=api.state();
      if(reset.window!=='wnd[0]'||reset.transaction!==transaction||!['','S','I'].includes(reset.statusType))
        throw new Error(message+' SAP did not return to its initial screen.');
      throw Object.assign(new Error(message),{rowError:true});
    }
  };
  try{return JSON.stringify(run(JSON.parse(inputJson),api)??null)}
  catch(error){if(!error.rowError)throw error;return JSON.stringify({rowError:error.message})}
}

function resolveToolboxQuotes(ctx,api){
  const lookupOnly=ctx.spec.Tool==='zbomLookup';
  const lookups=ctx.spec.Form?.sapLookupFids,versionFid=/^(\d{6})(?:-[A-Za-z0-9-]+)?$/;
  const jobs=[],jobByRow=new Map();
  for(const [index,row] of ctx.spec.Rows.entries()){
    if(row.Fields?.inputError)continue;
    const latest=toolboxLatestQuoteFid(row.Quote);
    if(lookupOnly?latest:!latest&&lookups?.[row.Quote]===undefined)continue;
    const input=latest?{fid:latest}:{quoteForVersion:row.Quote};
    input.reportRowError=true;
    if(ctx.mock)input.mockResult=latest?{quote:'29'+latest,version:latest+'-01'}:api.mockFid(row);
    jobByRow.set(index,jobs.length);
    jobs.push({input,label:'Looking up '+(latest?'FID '+latest:'quotation '+row.Quote)+' in SAP'});
  }
  const batch=api.sapBatch('Looking up quotations in SAP','VA23',toolboxSapQuoteSource,jobs);
  const rows=ctx.spec.Rows.map((row,index)=>{
    if(row.Fields?.inputError)return api.fail(index,row,row.Fields.inputError);
    const latest=toolboxLatestQuoteFid(row.Quote),requiredFid=lookups?.[row.Quote];
    if(lookupOnly&&latest)return {...row,Version:latest};
    if(!jobByRow.has(index))return row;
    const reply=batch.results[jobByRow.get(index)];
    if(!reply)return api.fail(index,row,'Not looked up: the task stopped before this row.');
    if(reply.error||reply.value?.rowError)return api.fail(index,row,reply.error||reply.value.rowError);
    if(!latest){
      const version=reply.value,fid=versionFid.exec(version)?.[1]||'';
      if(lookupOnly)return {...row,Version:version||''};
      if(requiredFid&&fid&&fid!==requiredFid)return api.fail(index,row,'SAP links quotation '+row.Quote+' to FID '+fid+' (expected '+requiredFid+').');
      return {...row,Version:fid?version:'',SapLookup:fid?'FID '+version+' read from SAP.':'No FID version on SAP quotation.'};
    }
    const found=reply.value;
    if(!found)return api.fail(index,row,'No SAP quotation found for FID '+latest+'.');
    const lookup=['SAP quote '+found.quote,found.version,found.created?'Created '+found.created:''].filter(Boolean).join(' / ');
    api.progress('FID '+latest+' / '+lookup);
    api.result(index,'Pending',lookup);
    return {...row,Quote:found.quote,Version:found.version,SapLookup:lookup};
  });
  // Failed lookups are already reported on their rows; the other rows continue.
  if(lookupOnly){
    const fids=rows.map(row=>row.Error?'':versionFid.exec(row.Version)?.[1]||'');
    const unique=[...new Set(fids.filter(Boolean))];
    const lists=api.sapBatch('Reading SAP quotation lists','VA23',toolboxSapQuoteSource,unique.map(fid=>({
      label:'Reading quotation list for FID '+fid,
      input:{fid,list:true,mockResult:ctx.mock?[['28'+fid,fid+'-02'],['29'+fid,fid+'-01']]:undefined}
    })));
    const byFid=new Map(unique.map((fid,index)=>[fid,lists.results[index]]));
    rows.forEach((row,index)=>{
      if(row.Error)return;
      const reply=byFid.get(fids[index]);
      if(reply?.error)api.fail(index,row,reply.error);
      else if(reply||!fids[index])api.result(index,'Completed',JSON.stringify(reply?.value||[]),undefined,row.Version);
    });
    if(lists.error)throw new Error(lists.error);
  }
  return rows;
}

function toolboxSearchSapVersions(sap,{fid,group,label,requireWindow,documentOf}){
  const text=control=>String(sap.properties(control,['Text']).Text??'').trim();
  const invoke=(control,Member,Arguments=[])=>sap.call(control,{Action:'invoke',Member,Arguments});
  const sendKey=(window,key)=>invoke(window,'SendVKey',[{Type:'int',Value:String(key)}]);
  function setText(control,value){
    sap.set(control,'Text',value);
    if(text(control)!==value)throw new Error('SAP did not retain the search value for FID '+fid+'.');
  }
  const numberField='wnd[0]/usr/ctxtVBAK-VBELN';
  requireWindow('wnd[0]');
  invoke(numberField,'SetFocus');
  sendKey('wnd[0]',4);
  requireWindow('wnd[1]');
  const tab='wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001';
  if(!/^A:/.test(text(tab)))throw new Error('The SAP Version search tab is unavailable for FID '+fid+'.');
  invoke(tab,'Select');
  const preselection=tab+'/ssubSUBSCR_PRESEL:SAPLSDH4:0220',fields=preselection+'/sub:SAPLSDH4:0220';
  if(text(fields+'/txtG_SELFLD_TAB-KEYWORD[2,0]')!=='Version'||text(fields+'/ctxtG_SELFLD_TAB-LOW[12,24]')!==group)
    throw new Error('SAP did not show the expected '+label+' Version search for FID '+fid+'.');
  setText(fields+'/txtG_SELFLD_TAB-LOW[2,24]',fid+'*');
  const limit=500;
  setText(preselection+'/txtDDSHF4CTRL-MAXRECORDS',String(limit));
  sendKey('wnd[1]',0);
  const search=requireWindow('wnd[1]');
  const cancel=()=>{sap.press('wnd[1]/tbar[0]/btn[12]');requireWindow('wnd[0]')};
  if(/no values for this selection/i.test(search.statusText)){cancel();return []}
  const count=Number(/^Sales Document\s+(\d+)\s+Entr(?:y|ies)$/i.exec(text('wnd[1]'))?.[1]);
  if(!Number.isInteger(count)||count<1||count>=limit)throw new Error('SAP did not show a complete '+label+' list for FID '+fid+'.');
  const resultArea='wnd[1]/usr',records=new Map();
  let versionCol,documentCol,itemCol;
  while(records.size<count){
    const lines=new Map();
    for(const {Id,Text} of sap.children(resultArea,['Id','Text'])){
      const cell=/\/lbl\[(\d+),(\d+)\]$/.exec(Id);
      if(!cell)continue;
      const [,column,line]=cell;
      if(!lines.has(line))lines.set(line,new Map());
      lines.get(line).set(column,String(Text??'').trim());
    }
    if(versionCol===undefined){
      for(const line of lines.values()){
        const headers=new Map([...line].map(([column,label])=>[label,column]));
        if(headers.has('Version')&&headers.has('Document')&&headers.has('Item')){
          versionCol=headers.get('Version');documentCol=headers.get('Document');itemCol=headers.get('Item');
          break;
        }
      }
      if(versionCol===undefined)throw new Error('SAP '+label+' results are missing the Version or Document column.');
    }
    let pageRows=0;
    for(const line of lines.values()){
      const version=line.get(versionCol),document=documentOf(line.get(documentCol)),item=line.get(itemCol);
      if(!version||!document||!/^\d{6}$/.test(item))continue;
      records.set([document,version,item].join(':'),{document,version});pageRows++;
    }
    if(records.size>=count)break;
    const via=['VerticalScrollbar'],scroll=()=>sap.properties(resultArea,['Position','Maximum','PageSize'],via);
    const {Position,Maximum,PageSize}=scroll();
    const next=Math.min(Maximum,Position+Math.max(1,Math.min(PageSize,pageRows)));
    if(!Number.isInteger(next)||next<=Position)throw new Error('SAP did not expose every '+label+' result for FID '+fid+'.');
    sap.set(resultArea,'Position',next,via);
    if(scroll().Position!==next)throw new Error('SAP could not scroll the '+label+' results for FID '+fid+'.');
  }
  if(records.size!==count)throw new Error('SAP '+label+' result count changed for FID '+fid+'.');
  cancel();
  return [...records.values()];
}

function toolboxLookupSapQuote({fid,quoteForVersion,list,details,byNumber,reportRowError,mockResult},sap){
  if(mockResult!==undefined)return mockResult;
  const target=fid?'FID '+fid:'quotation '+quoteForVersion;
  function requireWindow(expected){
    const state=sap.state();
    if(reportRowError&&state.transaction==='VA23'&&state.window==='wnd[0]'&&state.statusType==='E')
      sap.failRow('SAP could not look up '+target+': '+state.statusText);
    if(state.transaction!=='VA23'||state.window!==expected)throw new Error('SAP changed screens while looking up '+target+'.');
    if(!['','S','I'].includes(state.statusType))throw new Error('SAP could not look up '+target+': '+state.statusText);
    return state;
  }
  const text=control=>String(sap.properties(control,['Text']).Text??'').trim();
  const invoke=(control,Member,Arguments=[])=>sap.call(control,{Action:'invoke',Member,Arguments});
  const sendKey=(window,key)=>invoke(window,'SendVKey',[{Type:'int',Value:String(key)}]);
  function setText(control,value){
    sap.set(control,'Text',value);
    if(text(control)!==value)throw new Error('SAP did not retain the search value for '+target+'.');
  }
  const numberField='wnd[0]/usr/ctxtVBAK-VBELN';
  const header='wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301';
  function openQuote(quote){
    const overviewNumber='wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN';
    const salesMenu='wnd[0]/mbar/menu[2]/menu[1]/menu[0]';
    requireWindow('wnd[0]');
    setText(numberField,quote);
    sendKey('wnd[0]',0);
    requireWindow('wnd[0]');
    if(text(overviewNumber).replace(/^00(?=\d{8}$)/,'')!==quote||text(salesMenu)!=='Sales')
      throw new Error('SAP did not open the expected quotation '+quote+'.');
    invoke(salesMenu,'Select');
    requireWindow('wnd[0]');
    return text(header+'/txtVBAK-VSNMR_V');
  }
  function creationStamp(date,time){
    let year,month,day,parts;
    if((parts=/^(\d{4})[./-]?(\d{2})[./-]?(\d{2})$/.exec(date)))[,year,month,day]=parts;
    else if((parts=/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/.exec(date)))[,month,day,year]=parts;
    else if((parts=/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(date)))[,day,month,year]=parts;
    const calendar=new Date(Date.UTC(Number(year),Number(month)-1,Number(day)));
    if(calendar.getUTCFullYear()!==Number(year)||calendar.getUTCMonth()+1!==Number(month)||calendar.getUTCDate()!==Number(day)||
      !/^(?:[01]\d|2[0-3]):[0-5]\d:[0-5]\d$/.test(time))throw new Error('SAP returned an unreadable quotation creation date or time for '+target+'.');
    return calendar.toISOString().slice(0,10)+' '+time;
  }
  const createdAt=()=>creationStamp(text(header+'/ctxtVBAK-ERDAT'),text(header+'/ctxtVBAK-ERZET'));
  if(quoteForVersion){
    const version=openQuote(quoteForVersion);
    return details?{quote:quoteForVersion,version,created:createdAt()}:version;
  }

  const records=toolboxSearchSapVersions(sap,{fid,group:'2',label:'quotation',requireWindow,
    documentOf:value=>/^(?:00)?(\d{8})$/.exec(value)?.[1]});
  const candidates=new Map(),versionPattern=new RegExp('^'+fid+'(?:-[A-Za-z0-9-]+)?$','i');
  for(const {document:quote,version} of records){
    if(!byNumber&&!versionPattern.test(version))continue;
    if(candidates.has(quote)&&candidates.get(quote)!==version)throw new Error('SAP returned conflicting versions for quotation '+quote+'.');
    candidates.set(quote,version);
  }
  if(list&&!details)return [...candidates];
  let latest=null;const documents=[],ordered=[...candidates];
  if(byNumber)ordered.sort((a,b)=>b[0].localeCompare(a[0]));
  for(const [quote,version] of byNumber&&!list?ordered.slice(0,1):ordered){
    if(latest)invoke('','StartTransaction',[{Type:'string',Value:'VA23'}]);
    if(openQuote(quote)!==version)throw new Error('SAP quotation '+quote+' no longer matches Version '+version+'.');
    const created=createdAt(),document={quote,version,created};
    if(list)documents.push(document);
    if(!latest||created>latest.created||created===latest.created&&quote>latest.quote)latest=document;
  }
  return list?(byNumber?documents:documents.sort((a,b)=>b.created.localeCompare(a.created)||b.quote.localeCompare(a.quote))):latest;
}

function toolboxZbomSelection({quotations,folder,version},sap){
  const window='wnd[1]',strip=window+'/usr/tabsTAB_STRIP',valueName='RSCSEL_255-SLOW_I',rangeName='RSCSEL_255-SHIGH_I';
  const outputType='wnd[0]/usr/ctxtRG_KSCHL-LOW',outputPath='wnd[0]/usr/ctxtP_FILE';
  const relative=id=>id.replace(/^.*?(?=wnd\[)/,'');
  function requireWindow(expected,printing=false){
    const state=sap.state();
    if(printing&&state.transaction==='ZMASSPRINT'&&state.window==='wnd[0]'&&state.statusType==='E')
      sap.failRow('SAP refused ZBOM printing: '+state.statusText);
    if(state.window!==expected||state.transaction!=='ZMASSPRINT')throw new Error('SAP changed screens while selecting the quotations.');
    if(folder&&!['','S','I'].includes(state.statusType))throw new Error('SAP '+state.statusType+': '+state.statusText);
  }
  const fieldText=control=>String(sap.properties(control,['Text']).Text??'').trim();
  function setField(control,value){
    sap.set(control,'Text',value);
    requireWindow('wnd[0]');
    if(fieldText(control)!==value.trim())throw new Error('SAP did not retain the print field '+control+'.');
  }
  function checkTabs(expected){
    const selected=sap.properties(strip,['Name'],['SelectedTab']),tabs=sap.children(strip,['Name','Text']);
    if(selected.Name!=='SIVA'||tabs.length!==4)throw new Error('SAP is not on the single-value quotation tab.');
    for(const tab of tabs){
      const count=Number(/[(（]([0-9]+)[)）]$/.exec(String(tab.Text??'').trim())?.[1]||0);
      if(count!==(tab.Name==='SIVA'?expected:0))throw new Error(expected===0
        ?'The SAP quotation selection already contains entries. Retry in a new SAP window.'
        :'SAP selection counts differ from the requested quotation list. Check single values, ranges and exclusions.');
    }
  }
  function findTable(parent,tables){
    for(const child of sap.children(parent,['Id','Type','ContainerType'])){
      const id=relative(child.Id);
      if(child.Type==='GuiTableControl'){
        if(sap.children(id,['Name']).some(cell=>cell.Name===valueName))tables.push(id);
      }else if(child.ContainerType)findTable(id,tables);
    }
  }
  function rowIndex(id){
    const match=/,([0-9]+)\]$/.exec(id);
    if(!match)throw new Error('Cannot read SAP quotation row positions.');
    return Number(match[1]);
  }
  if(!quotations.length||!quotations.every(quote=>/^[0-9]{8}$/.test(quote)))throw new Error('Expected eight-digit quotations.');
  if(folder){
    requireWindow('wnd[0]');
    setField(outputType,'ZBOM');setField(outputPath,folder);
  }
  sap.press('wnd[0]/usr/btn%_RG_VBELN_%_APP_%-VALU_PUSH');
  requireWindow(window);
  checkTabs(0);
  const tables=[];
  findTable(window,tables);
  if(tables.length!==1)throw new Error('Cannot identify one readable SAP quotation selection table.');
  const table=tables[0],scrollbar=['VerticalScrollbar'];
  const scroll=()=>sap.properties(table,['Position','Maximum'],scrollbar);
  const move=position=>sap.set(table,'Position',position,scrollbar);
  const cells=()=>sap.children(table,['Id','Name','Text']);
  if(scroll().Position!==0)move(0);
  let written=0;
  while(written<quotations.length){
    const position=scroll().Position,visible=sap.properties(table,['VisibleRowCount']).VisibleRowCount;
    const end=Math.min(position+visible,quotations.length);
    if(position<0||position>written||written>=end)throw new Error('SAP did not expose the next quotation row.');
    const values=new Map(cells().filter(cell=>cell.Name===valueName).map(cell=>[rowIndex(cell.Id),cell]));
    for(;written<end;written++){
      const cell=values.get(written-position);
      if(!cell)throw new Error('Cannot find a writable SAP quotation row.');
      if(String(cell.Text??'').trim())throw new Error('A target SAP quotation cell already contains a value. Retry in a new SAP window.');
      sap.set(relative(cell.Id),'Text',quotations[written]);
    }
    sap.press(window+'/tbar[0]/btn[0]');
    requireWindow(window);
    checkTabs(written);
    if(written<quotations.length)move(Math.min(written,scroll().Maximum));
  }

  // SAP can clamp scrolling to the last populated row even though RowCount includes blank placeholders.
  const rows=new Map(),original=scroll().Position;
  let previous=-1,requested=0;
  try{
    if(original!==0)move(0);
    for(;;){
      const {Position:position,Maximum:maximum}=scroll();
      const {RowCount:count,VisibleRowCount:visible}=sap.properties(table,['RowCount','VisibleRowCount']);
      if(position<=previous||position>requested||visible<=0)throw new Error('SAP quotation table did not scroll as expected.');
      previous=position;
      for(const cell of cells()){
        const value=String(cell.Text??'').trim();
        if(cell.Name===rangeName&&value)throw new Error('SAP contains an unexpected quotation range.');
        if(cell.Name!==valueName)continue;
        const row=position+rowIndex(cell.Id);
        if(value)rows.set(row,value);
        else rows.delete(row);
      }
      if(rows.size>=quotations.length||position+visible>=count)break;
      requested=Math.min(position+visible,maximum);
      if(requested<=position)throw new Error('Cannot verify the remaining SAP quotation rows.');
      move(requested);
    }
  }finally{if(scroll().Position!==original)move(original)}
  checkTabs(quotations.length);
  const actual=[...rows].sort((left,right)=>left[0]-right[0]).map(([,value])=>value);
  if(actual.length!==quotations.length||actual.some((value,index)=>value!==quotations[index]))
    throw new Error('SAP quotation list differs from the requested list (expected '+quotations.length+', read '+actual.length+').');
  sap.press(window+'/tbar[0]/btn[8]');
  requireWindow('wnd[0]');
  if(folder){
    const radio='wnd[0]/usr/radP_RAD2',process='wnd[0]/tbar[1]/btn[14]';
    sap.call(radio,{Action:'invoke',Member:'Select'});requireWindow('wnd[0]');
    setField('wnd[0]/usr/ctxtPM_VERMO',String(version));
    if(fieldText(outputType)!=='ZBOM'||fieldText(outputPath)!==folder.trim()||
      String(sap.properties(radio,['Selected']).Selected).toLowerCase()!=='true')throw new Error('SAP changed the verified ZBOM print settings.');
    sap.press('wnd[0]/tbar[1]/btn[8]');requireWindow('wnd[0]',true);
    try{sap.properties(process,['Text'])}
    catch(error){if(error.code==='SAP_CONTROL_NOT_FOUND')return actual.length;throw error}
    sap.press('wnd[0]/tbar[1]/btn[5]');requireWindow('wnd[0]',true);
    sap.press(process);requireWindow('wnd[0]',true);
  }
  return actual.length;
}

function toolboxPdfWorkflow(ctx,api){
  const rows=resolveToolboxQuotes(ctx,api),completed=new Set(),claims=new Map(),byQuote=new Map(),failures=new Map();
  rows.forEach((row,index)=>{
    if(row.Error)return;
    if(!byQuote.has(row.Quote))byQuote.set(row.Quote,[]);
    byQuote.get(row.Quote).push(index);
  });
  if(!byQuote.size)return;
  const readPdfs=()=>api.files(ctx.work,'*.pdf',true).sort((a,b)=>a.directory<b.directory?1:a.directory>b.directory?-1:0);
  const quoteOf=file=>/^00([0-9]{8})_/.exec(file.name)?.[1]||'';
  const pdfPattern=quote=>'^00'+quote+'_';
  const remaining=()=>[...byQuote.keys()].filter(quote=>!completed.has(quote));
  function collect(files){
    for(const file of files){
      const quote=quoteOf(file),indexes=byQuote.get(quote);
      if(!indexes)throw new Error('An exported PDF could not be matched to a quotation: '+file.name);
      if(completed.has(quote))continue;
      const inspection=api.inspect(file.path);
      if(inspection.length<12||inspection.head!=='%PDF-'||!inspection.tail.includes('%%EOF')){
        failures.set(quote,'SAP produced an incomplete or invalid PDF: '+file.name);continue;
      }
      const row=rows[indexes.find(index=>rows[index].FileNameEdited)??indexes[0]];
      const detected=/_([0-9]{6}(?:[-_][A-Za-z0-9]+)*)\.pdf$/i.exec(file.name)?.[1]||'';
      const sapRow=indexes.map(index=>rows[index]).find(item=>item.SapLookup)||row;
      const fid=api.exportedFid(sapRow,detected,api.usesFidVersion&&!row.FileNameEdited);
      const newFid=fid&&fid.slice(0,6)!==row.Fid.slice(0,6);
      let name=row.FileName;
      if(!row.FileNameEdited&&(api.usesFidVersion||row.SapLookup||newFid||api.deferredDescription(row))){
        const stem=api.renderName(row,fid||row.Fid.slice(0,6)).replace(/[<>:"/\\|?*\u0000-\u001F]/g,'-').replace(/_{2,}/g,'_').replace(/-{2,}/g,'-').replace(/\s{2,}/g,' ').replace(/^[ ._-]+|[ ._-]+$/g,'')||row.Quote;
        name=stem.slice(0,190).replace(/[ ._-]+$/g,'')+'.PDF';
        if(newFid)api.progress(row.Quote+' / FID read from SAP: '+fid);
      }
      const stem=name.replace(/\.pdf$/i,'');let suffix=0;
      while(claims.has(name.toLowerCase())&&claims.get(name.toLowerCase())!==row.Quote)name=stem+'_'+row.Quote+(++suffix>1?'_'+suffix:'')+'.PDF';
      claims.set(name.toLowerCase(),row.Quote);
      const destination=api.path(ctx.spec.OutputFolder,name);api.publish(file.path,destination);
      for(const index of indexes){
        const item=rows[index],message=[item.SapLookup,api.exportedFidMessage(item,fid)].filter(Boolean).join(' / ');
        api.result(index,'Completed',message,destination,fid||item.Fid);
      }
      completed.add(quote);
    }
  }
  function print(quotations,version){
    quotations=quotations.filter(quote=>!completed.has(quote));
    if(!quotations.length)return;
    const folder=api.path(ctx.work,'v'+version);api.mkdir(folder);
    api.progress('Version '+version+' / Preparing '+quotations.length+' quotations');
    api.sap('transaction','','ZMASSPRINT');
    api.progress('Version '+version+' / Waiting for SAP to print');
    const error=api.sapScript(toolboxZbomSelectionSource,{quotations,folder:ctx.mock?'':folder,version})?.rowError;
    if(ctx.mock)api.mockFiles(quotations.map(quote=>({path:api.path(folder,'00'+quote+'_ZBOM_MOCK_'+api.mockFid(rows[byQuote.get(quote)[0]])+'.pdf'),text:'%PDF-1.4\n% Mock SAP output, for testing only\n%%EOF\n'})));
    else if(!error){api.progress('Version '+version+' / Waiting for PDF files');api.waitFiles(ctx.work,'*.pdf',true,quotations.map(pdfPattern))}
    collect(readPdfs());
    if(!error)return;
    const pending=quotations.filter(quote=>!completed.has(quote));
    if(pending.length===1){failures.set(pending[0],error);return}
    // Isolate a document error only after a verified SAP refusal of the batch.
    if(pending.length>1){
      const middle=Math.ceil(pending.length/2);
      print(pending.slice(0,middle),version);print(pending.slice(middle),version);
    }
  }
  api.invoke('watchFiles',{directory:ctx.work,pattern:'*.pdf',recursive:true,required:remaining().map(pdfPattern)});
  for(const version of [2,1]){
    let pending=remaining();if(!pending.length)break;
    if(version===1&&readPdfs().some(file=>pending.includes(quoteOf(file)))){
      api.progress('Checking late PDFs before Version 1');
      api.waitFiles(ctx.work,'*.pdf',true,pending.map(pdfPattern));collect(readPdfs());pending=remaining();
    }
    if(pending.length)print(pending,version);
  }
  for(const quote of remaining())for(const index of byQuote.get(quote))
    api.fail(index,rows[index],failures.get(quote)||'SAP produced no PDF after Versions 2 and 1.');
}

function toolboxCompareExport({quote,directory,path},sap){
  function requireWindow(expected){
    const state=sap.state();
    const type=String(state.statusType??'').trim().toUpperCase();
    if(state.transaction==='VA23'&&state.window==='wnd[0]'&&type==='E')sap.failRow('SAP could not export quotation '+quote+': '+state.statusText);
    if(String(state.transaction).toUpperCase()!=='VA23'||state.window!==expected)throw new Error('SAP changed screens while exporting the quotation. Check SAP before retrying.');
    if(['E','A','X','W'].includes(type))throw new Error('SAP '+type+': '+state.statusText+' Check SAP before continuing.');
    if(state.statusText==null||state.statusType==null||!['','S','I'].includes(type))throw new Error('SAP status could not be verified. Check SAP before continuing.');
  }
  function optionalText(control){
    try{return String(sap.properties(control,['Text']).Text??'')}
    catch(error){if(error.code==='SAP_CONTROL_NOT_FOUND')return null;throw error}
  }
  function setText(control,value){
    sap.set(control,'Text',value);
    if(String(sap.properties(control,['Text']).Text??'').trim()!==value.trim())throw new Error('SAP did not retain the expected field value: '+control+'.');
  }
  const invoke=(control,Member)=>sap.call(control,{Action:'invoke',Member});
  requireWindow('wnd[0]');
  setText('wnd[0]/usr/ctxtVBAK-VBELN',quote);
  requireWindow('wnd[0]');
  invoke('wnd[0]/mbar/menu[0]/menu[5]','Select');
  requireWindow('wnd[1]');
  const popup=optionalText('wnd[1]/usr/txtMESSTXT1');
  if(popup)sap.failRow('SAP: '+popup,true);
  const table='wnd[1]/usr/tblSAPLVMSGTABCONTROL';
  const outputRow=[4,...Array.from({length:31},(_,row)=>row).filter(row=>row!==4)]
    .find(row=>optionalText(table+'/txtNAST-KSCHL[0,'+row+']')?.trim()==='ZBOM');
  if(outputRow===undefined)sap.failRow('ZBOM output type was not found.',true);
  sap.call(table,[{Action:'invoke',Member:'GetAbsoluteRow',Arguments:[{Type:'int',Value:String(outputRow)}]},
    {Action:'set',Member:'Selected',Value:{Type:'bool',Value:'true'}}]);
  invoke(table+'/txtNAST-KSCHL[0,'+outputRow+']','SetFocus');
  sap.press('wnd[1]/tbar[0]/btn[37]');
  requireWindow('wnd[0]');
  invoke('wnd[0]/mbar/menu[2]/menu[1]','Select');
  requireWindow('wnd[0]');
  invoke('wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[1]','Select');
  requireWindow('wnd[1]');
  const radio='wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]';
  invoke(radio,'Select');
  if(String(sap.properties(radio,['Selected']).Selected).toLowerCase()!=='true')throw new Error('SAP did not select the text export format.');
  invoke(radio,'SetFocus');
  sap.press('wnd[1]/tbar[0]/btn[0]');
  requireWindow('wnd[1]');
  const split=optionalText('wnd[1]/usr/ctxtDY_PATH')!==null;
  if(split)setText('wnd[1]/usr/ctxtDY_PATH',directory);
  setText('wnd[1]/usr/ctxtDY_FILENAME',split?quote+'.tsv':path);
  requireWindow('wnd[1]');
  sap.press('wnd[1]/tbar[0]/btn[11]');
}

function toolboxCompareStage(input,bridge,compare){
  const invoke=(op,args)=>{
    const reply=JSON.parse(bridge(JSON.stringify({op,...args})));
    if(!reply.ok)throw new Error(reply.error.message);
    return reply.value;
  };
  if(input.quote){
    try{
      const data=compare.parse(invoke('readFile',{path:input.path}),input.quote);
      return JSON.stringify({value:{fid:data.fid},data});
    }catch(error){
      if(!error.rowError)throw error;
      return JSON.stringify({value:{rowError:error.message}});
    }
  }
  const models=input.models.map(({path,fid,warning})=>({...JSON.parse(invoke('readFile',{path})),fid,warning}));
  const align=(left,right,gap,mismatch)=>invoke('alignRows',{left,right,gap,mismatch});
  invoke('writeWorkbook',{path:input.path,plan:compare.plan(models,models.length>1,align)});
  return JSON.stringify({value:null});
}

function toolboxCompareWorkflow(ctx,api){
  const rows=resolveToolboxQuotes(ctx,api),models=new Map(),groups=new Map(),byQuote=new Map();
  const message=(row,model)=>[row.SapLookup,model.warning].filter(Boolean).join(' / ');
  rows.forEach((row,index)=>{
    if(row.Error)return;
    if(!byQuote.has(row.Quote))byQuote.set(row.Quote,[]);
    byQuote.get(row.Quote).push(index);
  });
  const exports=[...byQuote].map(([quote,indexes])=>{
    const path=api.path(ctx.work,quote+'.tsv');
    return {indexes,input:{quote,directory:ctx.work,path},file:path,label:'Exporting quotation '+quote,
      mockText:ctx.mock?'Quotation No.: '+api.mockFid(rows[indexes[0]])+'\r\nOption Type\t\tOption Selection\t\r\n\tSYSTEM\t\t\r\n\tPlatform\t\tMock '+quote:undefined};
  });
  const batch=api.sapBatch('Exporting ZBOM from SAP','VA23',toolboxCompareExportSource,exports.map(({indexes,...job})=>job));
  let exportFailure=batch.error;
  for(const [position,{indexes,input:{quote,path}}] of exports.entries()){
    const reply=batch.results[position];
    const row=rows[indexes[0]],fail=error=>indexes.forEach(index=>api.fail(index,rows[index],error));
    if(!reply){fail('Not exported: the task stopped before this quotation.');continue}
    try{
      const error=reply.error||reply.value?.rowError;
      if(error){fail(error);continue}
      if(!api.files(ctx.work,quote+'.tsv').some(file=>file.length>0)){fail('SAP created no readable TSV: '+path);continue}
      const store=quote+'-model.json';
      const model=api.invoke('scriptStage',{source:toolboxCompareStageSource,input:{path,quote},store});
      if(model.rowError){fail(model.rowError);continue}
      model.path=api.path(ctx.work,store);
      const sapRow=indexes.map(index=>rows[index]).find(item=>item.SapLookup)||row;
      model.fid=api.exportedFid(sapRow,model.fid,api.usesFidVersion&&indexes.some(index=>!rows[index].FileNameEdited),exportFailure);
      if(!model.fid){model.fid=row.Fid.slice(0,6)||quote;if(!row.FromBobj)model.warning='FID could not be read from the exported TSV. Used '+model.fid+'.'}
      else model.warning=api.exportedFidMessage(row,model.fid)||model.warning;
      models.set(quote,model);
      for(const index of indexes)api.result(index,'Formatted',message(rows[index],model),undefined,model.fid);
    }catch(error){fail(error.message);if(!error.rowError)exportFailure||=error.message}
  }
  rows.forEach((row,index)=>{
    if(row.Fields?.inputError)return;
    const key=row.Group==='__STANDALONE__'?'quote:'+row.Quote:'group:'+row.Group;
    if(!groups.has(key))groups.set(key,[]);groups.get(key).push(index);
  });
  const names=new Set(),hasDeferredNames=api.usesFidVersion||rows.some(row=>row.SapLookup||api.deferredDescription(row));
  for(const indexes of groups.values()){
    const members=indexes.map(index=>rows[index]);
    if(members.some(row=>row.Error||!models.has(row.Quote))){
      for(const index of indexes)if(!rows[index].Error&&models.has(rows[index].Quote))
        api.fail(index,rows[index],'Group incomplete. A quotation did not export.');
      continue;
    }
    const unique=new Map();
    for(const row of members)if(!unique.has(row.Quote))unique.set(row.Quote,row);
    const compared=[...unique.values()],selected=[...unique.keys()].map(quote=>models.get(quote));
    const edited=members.find(row=>row.FileNameEdited),hasLatest=members.some(row=>row.SapLookup);
    let name=edited?.FileName||compared[0].FileName;
    if(!edited&&(api.usesFidVersion||hasLatest||compared.some(api.deferredDescription)||!compared.every((row,index)=>row.Fid&&selected[index].fid.startsWith(row.Fid.slice(0,6))))){
      const customers=new Set(),descriptions=new Set(),compact=text=>text.replace(/_{2,}/g,'_').replace(/^[ ._-]+|[ ._-]+$/g,'');
      const parts=compared.map((row,index)=>compact(api.renderName(row,selected[index].fid,customers,descriptions))||row.Quote);
      name=compact(parts.join(' '+(String(ctx.spec.Form.zbomCompareSeparator||'').trim()||'vs')+' ')).replace(/[<>:"/\\|?*\u0000-\u001F]/g,',').replace(/\.xlsx$/i,'');
      name=(name||'ZBOM Compare')+'.xlsx';
    }
    if(hasDeferredNames){
      const stem=name.replace(/\.xlsx$/i,'');let suffix=0;
      while(names.has(name.toLowerCase()))name=stem+'_'+compared[0].Quote+(++suffix>1?'_'+suffix:'')+'.xlsx';
      names.add(name.toLowerCase());
    }
    const pending=api.path(ctx.work,'comparison-'+indexes[0]+'.xlsx'),destination=api.path(ctx.spec.OutputFolder,name);
    api.progress('Comparing and saving '+name);
    api.invoke('scriptStage',{source:toolboxCompareStageSource,input:{path:pending,models:selected}});
    api.publish(pending,destination);
    for(const index of indexes){
      const row=rows[index],model=models.get(row.Quote);
      api.result(index,'Completed',message(row,model),destination,model.fid);
    }
  }
  if(exportFailure)throw new Error(exportFailure);
}

function toolboxCreateQuoteWorkflow(ctx,api){
  let replacement=null;
  api.progress('Opening reference quotation');api.sap('transaction','','ZVA21M');
  api.sap('setText','wnd[0]/usr/txtGS_HEADER-VBELN',ctx.spec.Form.createQuoteQuotation);api.sap('setKey','wnd[0]/usr/cmbGS_HEADER-ACQ_ID','999');api.sap('sendVKey','wnd[0]','',{key:0});
  const reference=api.read('wnd[0]/usr/txtGS_HEADER-VERSION'),matched=/(?:^|\D)([0-9]{6})(?![0-9])/.exec(reference);
  if(!matched)api.ask('FID could not be extracted from reference version "'+reference+'". Continue?');
  else{
    const fid=matched[1],row=ctx.spec.Rows.find(item=>item.Fid===fid);
    if(row){
      const suffix=reference.slice(reference.indexOf(fid)+fid.length),version=/^(\D*)([0-9]+)(.*)$/.exec(suffix);
      if(!version)api.ask('Reference FID version "'+reference+'" could not be incremented. Continue?');
      else{replacement={...row,Version:fid+version[1]+String(BigInt(version[2])+1n).padStart(version[2].length,'0')+version[3],Crd:api.read('wnd[0]/usr/txtGS_HEADER-CRD')};}
    }
  }
  const table='wnd[0]/usr/tblZSDR0390_MASS_QUOTE_CREATE1MYTABLE',effective=original=>replacement?.Fid===original.Fid?replacement:original;
  const dates={dateFormats:['MM/dd/yyyy','M/d/yyyy','dd.MM.yyyy','d.M.yyyy','yyyy-MM-dd','yyyyMMdd']};
  ctx.spec.Rows.forEach((original,index)=>{
    const row=effective(original);api.progress('Entering item '+(index+1)+' of '+ctx.spec.Rows.length+' / FID '+row.Fid);
    if(row.Crd)api.sap('setText',table+'/ctxtGS_ITEM-REQ_DEL_DATE[1,'+index+']',row.Crd,dates);
    api.sap('setText',table+'/txtGS_ITEM-VERSION[2,'+index+']',row.Version);api.sap('setText',table+'/ctxtGS_ITEM-VALID_TO[3,'+index+']',row.ValidTo,dates);
  });
  if(ctx.spec.Rows.every(row=>effective(row).Crd)){api.focus(table+'/ctxtGS_ITEM-VALID_TO[3,0]');api.sap('sendVKey','wnd[0]','',{key:0})}
  api.progress('Checking the values retained by SAP');
  ctx.spec.Rows.forEach((original,index)=>{
    const row=effective(original);
    for(const [field,value,date] of [['ctxtGS_ITEM-REQ_DEL_DATE[1,',row.Crd,true],['txtGS_ITEM-VERSION[2,',row.Version,false],['ctxtGS_ITEM-VALID_TO[3,',row.ValidTo,true]]){
      const actual=api.read(table+'/'+field+index+']').trim(),expected=value.trim();
      if(date&&actual&&expected){if(api.date(actual)!==api.date(expected))throw new Error('SAP retained a different date for item '+(index+1)+'. Check SAP.')}
      else if(actual!==expected)throw new Error('SAP retained a different value for item '+(index+1)+'. Check SAP.');
    }
    api.result(index,'Completed','Version '+row.Version+(row.Crd?' / CRD '+row.Crd:''));
  });
}
// Document discovery and export run in the bridge's guarded SAP window.
function partChangesResultDetails(row,result){
  const message=result?.Message||'';
  // The bridge keeps the submitted input and the resolved document separately.
  if(message.startsWith('{'))try{
    const detail=JSON.parse(message);
    if(typeof detail?.message==='string'&&(Array.isArray(detail.locationFiles)||(!row?.Fields?.created&&detail.document)))return detail;
  }catch{/* Keep the original result text. */}
  return {message};
}

function partChangesModuleLocation(fileName){
  // The task already identifies the document. SAP may append line items to its number.
  const stem=fileName.replace(/\.xlsx$/i,'').split(/\s*---\s*/,1)[0];
  return stem.replace(/^.*?\s+-\s+/,'').trim()||fileName;
}

function partChangesGroups(rows){
  const groups=new Map();
  for(const row of rows){
    if(row.Group==='__STANDALONE__')continue;
    if(!groups.has(row.Group))groups.set(row.Group,[]);
    groups.get(row.Group).push(row);
  }
  return groups;
}

function partChangesDocumentType(number){
  if(/^\d{8}$/.test(number))return 'quotation';
  if(/^[A-Z][A-Z0-9]{0,9}$/.test(number))return 'so';
  return '';
}

function partChangesFidOf(version){
  return /^(\d{6})/.exec(version)?.[1]||'';
}

function partChangesNumberKey(number){return number.toUpperCase().replace(/\d+/g,digits=>digits.padStart(10,'0'))}

function partChangesRowType(row){
  // Older task records used the Quote field for SOs and had no document type.
  return row.Fields?.documentType||'so';
}

function partChangesDocumentKey(type,number){return type+':'+number}
function partChangesDocumentLabel(type){return type==='quotation'?'Quotation':'SO'}

function partChangesRowBatches(rows){
  // A latest selection resolves inside its own SAP window and can name any document of its FID,
  // so every row of that FID shares its batch; two windows never export the same document.
  const latestFids=new Set(rows.filter(row=>!row.Quote).map(row=>row.Fid));
  if(latestFids.size&&rows.some(row=>row.Quote&&!row.Fid))return null;
  const batches=new Map();
  rows.forEach((row,index)=>{
    const key=latestFids.has(row.Fid)?'fid:'+row.Fid:partChangesDocumentKey(partChangesRowType(row),row.Quote);
    if(!batches.has(key))batches.set(key,[]);
    batches.get(key).push(index);
  });
  return Object.fromEntries(batches);
}

function validatePartChangesTask(spec){
  const lookup=spec.Tool==='partChangesLookup',documents=new Map();
  for(const row of spec.Rows){
    if(lookup){
      if(row.Fid?!/^\d{6}$/.test(row.Fid)||row.Quote:!partChangesDocumentType(row.Quote))
        throw new Error('Enter a six-digit FID, an eight-digit quotation, or a sales order starting with a letter.');
      continue;
    }
    const type=partChangesRowType(row),key=partChangesDocumentKey(type,row.Quote),policy=row.Fields.lookupPolicy;
    const validNumber=type==='quotation'?/^\d{8}$/.test(row.Quote):type==='so'&&/^[A-Z0-9]{1,10}$/.test(row.Quote);
    const validFid=/^\d{6}$/.test(row.Fid),verified=Boolean(row.Version||row.Fields.created);
    const validExact=validNumber&&(!row.Fid||validFid);
    const validDocument=policy?['latest-so','latest-quote'].includes(policy)&&validFid&&!row.Quote&&!verified:
      validExact&&(!verified||validFid&&partChangesFidOf(row.Version)===row.Fid&&/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(row.Fields.created));
    if((row.Group!=='__STANDALONE__'&&!/^G\d{2,}$/.test(row.Group))||!validDocument)
      throw new Error('Enter a valid SAP document or FID for '+(row.Fid?'FID '+row.Fid:row.Quote)+'.');
    if(row.Group!=='__STANDALONE__'&&row.FileName&&
        (!/\.xlsx$/i.test(row.FileName)||/[<>:"/\\|?*\u0000-\u001f]/.test(row.FileName)||row.FileName.length>255))
      throw new Error('Invalid BOM Compare filename.');
    if(row.Quote&&row.Fid){
      if(documents.has(key)&&documents.get(key)!==row.Fid)throw new Error(row.Quote+' is linked to different FIDs.');
      documents.set(key,row.Fid);
    }
  }
  if(lookup)return;
  const groups=partChangesGroups(spec.Rows);
  if(groups.size&&!/^[A-Za-z]:[\\/]/.test(spec.OutputFolder))throw new Error('Use an absolute local output folder.');
  for(const [group,rows] of groups){
    if(rows.length<2)throw new Error(group+' needs at least two documents.');
    if(new Set(rows.map(row=>row.Quote?partChangesDocumentKey(partChangesRowType(row),row.Quote):row.Fields.lookupPolicy+':'+row.Fid)).size!==rows.length)
      throw new Error(group+' contains the same document twice.');
    if(rows.some(row=>row.FileName!==rows[0].FileName))throw new Error(group+' must have one output filename.');
  }
}

function createPartChangesWorkflow(tool,openMyLam=false){
  const lookup=tool==='partChangesLookup';
  const common=[partChangesDocumentType,partChangesRowType,partChangesDocumentKey,partChangesDocumentLabel,partChangesNumberKey].map(fn=>fn.toString()).join('\n')+'\n';
  const sources=common+'const workflowDate=('+toolboxWorkflowDate.toString()+');\n'+
    'const partChangesOrderSource='+JSON.stringify(toolboxSapScriptSource(toolboxReadSalesOrders,toolboxSearchSapVersions,toolboxWorkflowDate,partChangesNumberKey))+';\n'+
    'const toolboxSapQuoteSource='+JSON.stringify(toolboxSapScriptSource(toolboxLookupSapQuote,toolboxSearchSapVersions))+';\n'+
    findPartChangesDocuments.toString()+'\n'+
    (lookup?'':[partChangesFidOf,partChangesGroups].map(fn=>fn.toString()).join('\n')+'\nconst partChangesExportSource='+JSON.stringify(toolboxSapScriptSource(toolboxExportPartChangesBom))+';\n');
  const run=lookup?lookupPartChangesDocuments:exportPartChangesBoms;
  return {Api:1,KeepWindow:false,CompletionMessage:lookup?'SAP lookup finished':'BOM exports ready',
    OutputWorkFolder:'',AfterSuccess:openMyLam?[{Operation:'openPartChanges'}]:[],
    Operations:lookup?['sap','sapScript','progress','result']:['sap','sapScript','progress','result','path','mkdir','waitFiles','watchFileCount','writeWorkbook','mockFiles','publish'],
    Validation:'"use strict";\n'+common+partChangesGroups.toString()+'\n'+partChangesFidOf.toString()+'\n('+validatePartChangesTask.toString()+')(JSON.parse(inputJson));',
    Source:'"use strict";\n'+sources+'('+toolboxWorkflowRuntime.toString()+')(JSON.parse(contextJson),bridge,('+run.toString()+'));'};
}

function toolboxReadSalesOrders({fid,order,latestOnly},sap){
  const target=order?'SO '+order:'FID '+fid;
  function requireWindow(expected){
    const state=sap.state();
    const searchHelp=expected==='wnd[1]'&&state.statusText==='Control (amodal) search help could not be initialized';
    if(state.transaction!=='VA03'||state.window!==expected||(!['','S','I'].includes(state.statusType)&&!searchHelp))
      throw new Error('VA03 could not verify '+target+': '+(state.statusText||'unexpected SAP screen'));
    return state;
  }
  const text=control=>String(sap.properties(control,['Text']).Text??'').trim();
  const invoke=(control,Member,Arguments=[])=>sap.call(control,{Action:'invoke',Member,Arguments});
  const documentOf=value=>/^[A-Z0-9]{1,10}$/.test(value)?value:'';
  const numberField='wnd[0]/usr/ctxtVBAK-VBELN',overview='wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN';
  const salesMenu='wnd[0]/mbar/menu[2]/menu[1]/menu[0]';
  const header='wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301';
  function readOrder(document){
    requireWindow('wnd[0]');
    sap.set(numberField,'Text',document);
    if(text(numberField)!==document)throw new Error('SAP did not retain sales order '+document+'.');
    invoke('wnd[0]','SendVKey',[{Type:'int',Value:'0'}]);
    const state=sap.state();
    if(state.transaction==='VA03'&&state.window==='wnd[0]'&&state.statusType==='E')
      sap.failRow('SAP could not open SO '+document+': '+state.statusText);
    requireWindow('wnd[0]');
    if(text(overview)!==document||text(salesMenu)!=='Sales')throw new Error('SAP did not open sales order '+document+'.');
    invoke(salesMenu,'Select');
    requireWindow('wnd[0]');
    const version=text(header+'/txtVBAK-VSNMR_V'),time=text(header+'/ctxtVBAK-ERZET');
    if(!/^(?:[01]\d|2[0-3]):[0-5]\d:[0-5]\d$/.test(time))
      throw new Error('Could not read the creation time for SO '+document+'.');
    const date=toolboxWorkflowDate(text(header+'/ctxtVBAK-ERDAT').replace(/^(\d{4})[./](\d{2})[./](\d{2})$/,'$1-$2-$3'));
    return {number:document,version,created:date+' '+time};
  }
  if(order)return readOrder(order);
  const records=toolboxSearchSapVersions(sap,{fid,group:'0',label:'sales order',requireWindow,documentOf});
  // The Version search already limits every returned document to this FID.
  const orders=[...new Set(records.map(row=>row.document))].sort((a,b)=>partChangesNumberKey(b).localeCompare(partChangesNumberKey(a)));
  return (latestOnly?orders.slice(0,1):orders).map((document,index)=>{
    if(index)invoke('','StartTransaction',[{Type:'string',Value:'VA03'}]);
    return readOrder(document);
  });
}

function toolboxExportPartChangesBom({number,type,directory},sap){
  const quote=type==='quotation',transaction=quote?'ZOQUOTEBOM_CE_FMT':'ZOOBOM_CE_FMT';
  const subject=(quote?'Quotation ':'SO ')+number;
  function requireReport(state=sap.state()){
    if(state.transaction!==transaction||state.window!=='wnd[0]'||!['','S','I'].includes(state.statusType))
      throw new Error(subject+': '+(state.statusText||'Unexpected SAP BOM report screen.'));
    return state;
  }
  const input=quote?'S_QUOTE':'S_VBELN';
  const doc='wnd[0]/usr/ctxt'+input+'-LOW',folder='wnd[0]/usr/ctxtP_FOLDER',execute='wnd[0]/tbar[1]/btn[8]';
  const text=id=>String(sap.properties(id,['Text']).Text??'').trim();
  requireReport();
  if(text('wnd[0]/usr/txt%_'+input+'_%_APP_%-TEXT')!==(quote?'Quotation Number':'Sales document')||
      !String(sap.properties(execute,['Tooltip']).Tooltip).startsWith('Execute'))
    throw new Error('Cannot verify the BOM report controls for '+subject+'.');
  sap.set(doc,'Text',number);
  if(quote){
    const configured='wnd[0]/usr/chkP_ASCFG',master='wnd[0]/usr/chkP_INJCT';
    if(text(configured)!=='As-configured BOM only'||text(master)!=='Include Master BOM')
      throw new Error('Cannot verify the quotation BOM options.');
    const selected=id=>sap.properties(id,['Selected']).Selected===true;
    const choose=(id,value)=>sap.call(id,{Action:'set',Member:'Selected',Value:{Type:'bool',Value:String(value)}});
    if(selected(master))choose(master,false);
    if(!selected(configured))choose(configured,true);
    if(!selected(configured)||selected(master))throw new Error('SAP did not retain As-configured BOM only.');
  }
  const defaultDirectory=text(folder),destination=directory+'\\';
  if(!/^(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+)/.test(defaultDirectory))
    throw new Error('SAP has no absolute default BOM export folder for '+subject+'.');
  sap.set(folder,'Text',destination);
  if(text(folder).replace(/\//g,'\\').toUpperCase()!==destination.toUpperCase())
    throw new Error('SAP did not retain the temporary BOM export folder for '+subject+'.');
  if(text(doc)!==number)throw new Error('SAP did not retain the BOM export number for '+subject+'.');
  sap.press(execute);
  const state=sap.state();
  if(state.transaction===transaction&&state.window==='wnd[0]'&&state.statusType==='E')
    sap.failRow(subject+': '+state.statusText);
  return {defaultDirectory,status:requireReport(state).statusText};
}

function findPartChangesDocuments(ctx,api,row){
  const direct=row.Quote&&(row.Fields?.documentType||partChangesDocumentType(row.Quote)),policy=row.Fields?.lookupPolicy;
  const documents=[],issues=[],queried=[],types=direct?[direct]:policy==='latest-quote'?['quotation']:['so','quotation'];
  let failure='';
  for(const type of types){
    queried.push(type);
    const label=partChangesDocumentLabel(type);
    api.progress(direct?'Checking '+label+' '+row.Quote:'Finding '+label+' documents for FID '+row.Fid);
    try{
      let found;
      if(ctx.mock){
        const fixtureFid=direct?row.Quote.slice(-6):row.Fid,[first,second]=type==='so'?['P','S']:['29','28'];
        found=[{number:row.Quote||first+fixtureFid,version:fixtureFid+'-01',created:'2025-07-22 02:22:39'},
          {number:second+fixtureFid,version:fixtureFid+'-02',created:'2025-07-21 03:12:04'}];
        if(direct)found=found.slice(0,1);
      }else{
        let data;
        if(type==='quotation')data=api.lookupQuote(direct?{quoteForVersion:row.Quote,details:true}:{fid:row.Fid,list:!policy,details:true,byNumber:true});
        else{
          api.sap('transaction','','VA03');
          data=api.sapScript(partChangesOrderSource,direct?{order:row.Quote}:{fid:row.Fid,latestOnly:Boolean(policy)});
        }
        if(data?.rowError){issues.push(label+': '+data.rowError);if(policy==='latest-so'&&type==='so')break;continue}
        found=direct||type==='quotation'&&policy?(data?[data]:[]):data;
        if(type==='quotation')found=found.map(({quote,version,created})=>({number:quote,version,created}));
      }
      found.sort((a,b)=>partChangesNumberKey(b.number).localeCompare(partChangesNumberKey(a.number)));
      if(policy)found=found.slice(0,1);
      documents.push(...found.map((item,position)=>({...item,type,latest:!direct&&position===0})));
      if(policy==='latest-so'&&type==='so'&&found.length)break;
    }catch(error){const issue=label+': '+error.message;failure||=issue;issues.push(issue);if(policy==='latest-so'&&type==='so')break}
  }
  return {documents,issues,queried,failure};
}

function lookupPartChangesDocuments(ctx,api){
  ctx.spec.Rows.forEach((row,index)=>{
    const {documents,issues,queried}=findPartChangesDocuments(ctx,api,row);
    api.result(index,'Completed',JSON.stringify({documents,issues,queried}));
  });
}

function exportPartChangesBoms(ctx,api){
  const exports=new Map(),lookups=new Map(),verified=new Map();let exportFailure='';
  const documentKey=row=>partChangesDocumentKey(partChangesRowType(row),row.Quote);
  const rows=ctx.spec.Rows.map((row,index)=>{
    if(row.Fields.created)return row;
    const subject=row.Quote?partChangesDocumentLabel(partChangesRowType(row))+' '+row.Quote:'FID '+row.Fid;
    const key=row.Quote?documentKey(row):row.Fields.lookupPolicy+':'+row.Fid;
    if(!lookups.has(key)){
      const known=verified.get(key);
      if(known)lookups.set(key,{documents:[known],issues:[]});
      else{
        api.result(index,'Pending','Looking up '+subject+' in SAP.');
        lookups.set(key,findPartChangesDocuments(ctx,api,row));
      }
    }
    const {documents,issues,failure}=lookups.get(key),order=documents[0],fid=order&&partChangesFidOf(order.version);
    exportFailure||=failure||'';
    const error=!fid||!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(order.created)?
      issues.join(' · ')||'No matching SAP document with a FID and creation date for '+subject+'.':
      row.Fid&&fid!==row.Fid?'SAP links '+subject+' to FID '+fid+' (expected '+row.Fid+').':'';
    if(error){api.result(index,'Failed',error);return null}
    const resolved={...row,Quote:order.number,Fid:fid,Version:order.version,Fields:{...row.Fields,documentType:order.type,created:order.created}};
    verified.set(documentKey(resolved),order);
    return resolved;
  });
  const duplicateGroups=new Set();
  for(const [group,members] of partChangesGroups(rows.filter(Boolean))){
    if(new Set(members.map(documentKey)).size!==members.length)duplicateGroups.add(group);
  }
  // Keep grouped inputs with the task so another export cannot change the later comparison.
  const snapshots=ctx.spec.Workflow?.AfterSuccess?.some(action=>action.Operation==='openPartChanges')?
    new Set(rows.filter(row=>row&&row.Group!=='__STANDALONE__').map(documentKey)):new Set();
  const pattern='*.xlsx',isSystem=/^(?:.*?\s+-\s+)?SYSTEM(?:\s.*|---.*)?\.xlsx$/i;
  let work;
  for(const [index,row] of rows.entries()){
    if(!row)continue;
    const number=row.Quote,type=partChangesRowType(row),key=partChangesDocumentKey(type,number),subject=partChangesDocumentLabel(type)+' '+number;
    const captureLocations=snapshots.has(key);
    const result=(status,message,file,locationFiles)=>{
      const resolved=row!==ctx.spec.Rows[index];
      if(!resolved&&!locationFiles){api.result(index,status,message,file);return}
      const detail=resolved?{document:{number,type,version:row.Version,created:row.Fields.created},message}:{message};
      if(locationFiles)detail.locationFiles=locationFiles;
      api.result(index,status,JSON.stringify(detail),file,resolved?row.Fid:undefined);
    };
    if(duplicateGroups.has(row.Group)){
      result('Failed',row.Group+' resolves to the same SAP document more than once.');continue;
    }
    const known=verified.get(key),changed=subject+' details changed. Look up the document again.';
    if(!ctx.mock&&known&&(known.version!==row.Version||known.created!==row.Fields.created)){
      result('Failed',changed);continue;
    }
    if(!exports.has(key)){
      try{
        api.progress('Checking '+subject);
        if(!work){
          const folder=api.path(ctx.work,'boms',String(ctx.batchIndex));
          api.invoke('watchFileCount',{directory:folder,pattern,recursive:true});
          work=folder;
        }
        const directory=api.path(work,type+'-'+number),quote=type==='quotation';
        const sourceDirectory=quote?api.path(directory,number):directory;
        let defaultDirectory,reportStatus='';
        if(ctx.mock){
          defaultDirectory=api.path(ctx.directory,'mock-exports',type);
          api.mkdir(sourceDirectory);
          api.progress('Exporting BOMs for '+subject);
          const rows=[['MATERIAL','DESCRIPTION'],['100-000000-00','Common part'],['100-'+(row.Fid||number)+'-00','Test part for '+number]];
          if(Number(row.Fid)%2===0)rows.push(['100-999999-00','Shared across some SOs']);
          const write=(label,fileName,values)=>api.invoke('writeWorkbook',{path:api.path(sourceDirectory,fileName),plan:{Sheets:[{Name:label,Cells:values.flatMap(
            (cells,i)=>cells.map((Text,j)=>({Row:i+1,Column:j+1,Text})))}]}});
          write('SYSTEM',number+' - SYSTEM.xlsx',rows);
          // Expose the first workbook during the configured mock export delay.
          api.mockFiles([]);
          write('GB2',number+' - GB2 --- Lv mock.xlsx',rows.slice(0,3));
          write('PM2',number+(quote?'':'_165,649')+' - PM2 --- Lv mock.xlsx',[rows[0],rows[2],['100-777777-00','Module-only part']]);
          if(rows.length>3&&Number(row.Fid)%4===0)write('FIB-2',number+' - FIB-2 --- Lv mock.xlsx',[rows[0],rows[3]]);
        }else{
          let selected=known;
          if(!selected&&type==='quotation')selected=api.lookupQuote({quoteForVersion:number,details:true});
          else if(!selected){
            api.sap('transaction','','VA03');
            selected=api.sapScript(partChangesOrderSource,{fid:row.Fid,order:number});
          }
          if(selected?.rowError||selected?.version!==row.Version||selected.created!==row.Fields.created)
            throw Object.assign(new Error(selected?.rowError||changed),{rowError:true});
          verified.set(key,selected);
          api.sap('transaction','',quote?'ZOQUOTEBOM_CE_FMT':'ZOOBOM_CE_FMT');
          api.mkdir(directory);
          api.progress('Exporting BOMs for '+subject);
          const report=api.invoke('sapScript',{source:partChangesExportSource,input:{number,type,directory},timeoutMs:600000});
          if(report.rowError)throw Object.assign(new Error(report.rowError),{rowError:true});
          ({defaultDirectory,status:reportStatus}=report);
        }
        api.progress('Checking BOM workbooks for '+subject);
        const files=api.invoke('waitFiles',{directory:sourceDirectory,pattern,required:[isSystem.source],readyMs:ctx.mock?1500:3000,idleMs:25000});
        const system=files.filter(file=>file.length>0&&isSystem.test(file.name));
        if(system.length!==1)throw Object.assign(new Error((system.length?'Several':'No')+' SYSTEM workbooks for '+subject+'.'+(reportStatus?' '+reportStatus:'')),{rowError:true});
        if(files.some(file=>!file.length))throw Object.assign(new Error('SAP produced an empty BOM workbook for '+subject+'.'),{rowError:true});
        const destination=quote?api.path(defaultDirectory,number):defaultDirectory;
        const snapshot=system[0].path,file=api.path(destination,system[0].name);
        const locationFiles=captureLocations?files.filter(file=>file.path!==snapshot).map(file=>file.name):undefined;
        api.progress('Copying '+files.length+' BOM workbooks for '+subject);
        for(const file of files)api.publish(file.path,api.path(destination,file.name));
        exports.set(key,{file,snapshot,locationFiles,count:files.length});
      }catch(error){exports.set(key,{error:error.message});if(!error.rowError)exportFailure||=error.message}
    }
    const {file,snapshot,locationFiles,count,error}=exports.get(key);
    const useSnapshot=row.Group!=='__STANDALONE__'&&locationFiles!==undefined;
    result(error?'Failed':'Completed',error||count+' BOM '+(count===1?'workbook':'workbooks')+' exported',useSnapshot?snapshot:file,useSnapshot?locationFiles:undefined);
  }
  if(exportFailure)throw new Error(exportFailure);
}

function partChangesFileNameWithSuffix(name,suffix=''){
  const stem=name.replace(/\.xlsx$/i,'');
  return (suffix?stem.slice(0,250-suffix.length).replace(/[ .]+$/,''):stem)+suffix+'.xlsx';
}

function partChangesWorkbookPath(task,group,suffix='',name=task.spec.Rows.find(row=>row.Group===group)?.FileName){
  name||=group+' - '+BOM_COMPARE_LABEL+' '+task.id+'.xlsx';
  return task.spec.OutputFolder.replace(/[\\/]+$/,'')+'\\'+partChangesFileNameWithSuffix(name,suffix);
}

async function partChangesWorkbookIdentity(id,group,systems){
  const identity='Toolkit Kyra BOM Compare '+id+' '+group+' with fitted columns';
  const values=Array.isArray(systems)?systems:systems?.rows;
  if(!values?.length)return identity;
  // Explicit labels form a set. Auto depends on the BOM row that selected it.
  const lists=values.map(partChangesSystemList);
  const scope=JSON.stringify(lists.every(list=>!list.includes(''))?[...new Set(lists.flat())].sort():values);
  const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(scope));
  const key=Array.from(new Uint8Array(digest,0,16),byte=>byte.toString(16).padStart(2,'0')).join('');
  return identity+' · MyLam Systems '+key;
}

function normalizeBobjHeader(value){return String(value||'').trim().replace(/\s+/g,' ').toLocaleLowerCase();}

function pendingFidFilenameToken(token,header,omit){
  // Preview placeholders stay separate from resolved filename values.
  const field=normalizeBobjHeader(header);
  if(!field||['quote','so','today'].includes(field))return '';
  if(field==='customer'&&omit?.customer||field==='system description'&&omit?.systemDescription)return '';
  return token;
}

function partChangesTemplateValue(row,header,omit,previewValue){
  const key=normalizeBobjHeader(header),type=row.order?.type||row.type;
  if(key==='so'||key==='quote')return type===(key==='so'?'so':'quotation')?row.Quote:'';
  if(key==='fid')return row.Fid;
  if(key==='fidver')return row.Fid?(row.order?row.order.version.slice(6):'{FIDver}'):'';
  if(row.filenameValues){
    if(key==='customer'&&omit?.customer||key==='system description'&&omit?.systemDescription)return '';
    return row.filenameValues[key]||'';
  }
  return previewValue?.(row,header,omit)||'';
}

function plannedPartChangesFileNames(groups,template,separator,fixedNames=new Map(),value=partChangesTemplateValue,showPendingFid=false){
  const prefix=/^\[BOM\]\s*/i.exec(template)?.[0]||'',memberTemplate=template.slice(prefix.length);
  const names=new Map(fixedNames),used=new Set([...names.values()].map(name=>name.toLowerCase()));
  const pendingFid=row=>showPendingFid&&row.Quote&&!row.Fid;
  for(const [group,rows] of groups){
    if(names.has(group)||rows.some(row=>row.error))continue;
    const joined=joinCompareFileNames(rows,separator,value,(row,omit)=>{
      const stem=memberTemplate.replace(/\{([^{}]+)\}/g,(token,header)=>value(row,header,omit)||
        (pendingFid(row)?pendingFidFilenameToken(token,header,omit):''));
      return compactZbomCompareName(stem)||row.Quote||row.Input;
    });
    const base=finishZbomCompareFileName(prefix+joined,BOM_COMPARE_LABEL,rows.some(pendingFid));
    let name=partChangesFileNameWithSuffix(base),duplicate=0;
    while(used.has(name.toLowerCase())){
      name=partChangesFileNameWithSuffix(base,'_'+group+(duplicate?'_'+duplicate:''));duplicate++;
    }
    used.add(name.toLowerCase());names.set(group,name);
  }
  return names;
}

function runFcidGovernance(){
'use strict';
if(toolkitExpired())return;
if(document.documentElement.dataset.fcidMounted)return;
document.documentElement.dataset.fcidMounted='1';document.documentElement.lang='en'; const id=value=>document.getElementById(value); const text=value=>String(value??'').trim();
const escapeHtml=value=>String(value??'').replace(/[&<>"']/g,character=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));
const sources=fidSourceDefinitions();
const isRecord=value=>value&&typeof value==='object'&&!Array.isArray(value);

const bwData=createBwDataService();
const HOSTED_ARCHIVE_URLS=[
  'https://synalocey.github.io/ToolkitKyra/images/Stacey1.png',
  'https://raw.githubusercontent.com/synalocey/ToolkitKyra/main/images/Stacey1.png',
  'https://github.com/synalocey/ToolkitKyra/raw/refs/heads/main/images/Stacey1.png'
];
const HOSTED_PASSWORD_URL='/sap/opu/odata/sap/ZNSR_CD_SRV/GetNameDateTimeSet/?$filter=Userid%20eq%20%27LONGKY%27',HOSTED_PREFIXES={ctlNsr:'c',nce:'n'};
const store={raw:Object.fromEntries(Object.keys(sources).map(key=>[key,[]])),meta:{},busy:new Set(),refreshing:false,dirty:false,runs:{},controllers:{}};
const cacheKeys=Object.fromEntries(Object.keys(sources).map(key=>[key,key==='nsr'?'sap-nsr-search':(BW_SOURCES[key]?BW_CACHE_PREFIX:CACHE_PREFIX)+key]));
const sourceCacheRevisions=new WeakMap();
const FID_DATE_READ_KEY=CACHE_PREFIX+'fid-date-history',dateChangeDriver='Date Changes';
const fidDateLabels={mcsd:'MCSD',rd:'SRD'},fidDateKeys=Object.keys(fidDateLabels);
let fidDateChanges=new Map(),fidDateReads=new Map();
const NSR_MARKS_KEY='toolkitKyra.nsr.marks.v1',IGNORED_NSRS_KEY='toolkitKyra.nsr.ignored.v1',FID_MARKS_KEY='toolkitKyra.fid.marks.v1',TABLE_ORDERS_KEY='toolkitKyra.table.orders.v1';
const FID_CUSTOM_COLUMNS_KEY='toolkitKyra.fid.customColumns.v1';
const ignoredNsrs=new Set(),importantNsrs=new Set(),ignoredFids=new Set(),importantFids=new Set();
const markKinds={
  nsr:{key:NSR_MARKS_KEY,legacy:IGNORED_NSRS_KEY,label:'NSR#',plural:'NSRs',prefix:'Nsr',ignored:ignoredNsrs,important:importantNsrs,pattern:/^[A-Z]\d{5,9}$/,status:'fcNsrMarksStatus'},
  fid:{key:FID_MARKS_KEY,label:'FID',plural:'FIDs',prefix:'Fid',ignored:ignoredFids,important:importantFids,pattern:FID_ID_PATTERN,status:'fcFidMarksStatus'}
};
function readSavedMarkSets(kind){
  const config=markKinds[kind],ignored=new Set(),important=new Set();
  const saved=GM_getValue(config.key,null),legacy=saved===null&&config.legacy?GM_getValue(config.legacy,null):null,parsed=typeof(saved??legacy)==='string'?JSON.parse(saved??legacy):saved??legacy;
  const entries=Array.isArray(parsed)?parsed.map(number=>[number,'ignore']):Object.entries(parsed||{});
  for(const [value,mark]of entries){const number=String(value).trim().toUpperCase();if(!number||kind==='fid'&&!config.pattern.test(number))continue;if(mark==='ignore')ignored.add(number);else if(mark==='important')important.add(number)}
  return {ignored,important};
}
for(const [kind,config]of Object.entries(markKinds))try{
  const {ignored,important}=readSavedMarkSets(kind);
  for(const number of ignored)config.ignored.add(number);
  for(const number of important)config.important.add(number);
}catch{/* Keep source data available when personal marks cannot be read. */}
let fidNotes={};
try{fidNotes=readFidNotes()}catch{/* Keep source data available when notes cannot be read. */}
let tableOrders={};
try{const saved=GM_getValue(TABLE_ORDERS_KEY,null),parsed=typeof saved==='string'?JSON.parse(saved):saved;if(parsed&&typeof parsed==='object'&&!Array.isArray(parsed))tableOrders=Object.fromEntries(Object.entries(parsed).filter(([,order])=>Array.isArray(order)&&order.every(label=>typeof label==='string')).map(([name,order])=>[name,[...new Set(order.map(label=>label.replace(/^Calendar (?=Quarter$)/,'')))]]))}catch{/* Invalid preferences leave the default column order available. */}
if(tableOrders.portfolioTable)for(const [label,after]of [['CTL History','Note'],['NSR History','CTL History'],['MFG Upgrade Type','Business Unit'],['Active Quote','Active CTL']])
  if(!tableOrders.portfolioTable.includes(label)&&tableOrders.portfolioTable.includes(after))tableOrders.portfolioTable.splice(tableOrders.portfolioTable.indexOf(after)+1,0,label);
for(const order of [tableOrders.portfolioTable,tableOrders.wfAdditional])if(order?.includes('Active CTL')&&!order.includes('CTL Elapsed Days'))
  order.splice(order.indexOf('Active CTL')+1,0,'CTL Elapsed Days');
if(tableOrders.ctlTable){
  tableOrders.ctlTable=tableOrders.ctlTable.map(label=>label==='Created'?'Created Date':label);
  for(const [label,after]of [['Elapsed Days','Created Date'],['CTL Notes','Root Cause'],['Sub Reason','CTL Notes'],
    ['RD At Create','Sales Ops Approval'],['MCSD@Create','RD At Create'],['SRD @ SOPs Approval','MCSD@Create'],['MCSD At SOps Approval','SRD @ SOPs Approval']
  ])if(!tableOrders.ctlTable.includes(label)&&tableOrders.ctlTable.includes(after))tableOrders.ctlTable.splice(tableOrders.ctlTable.indexOf(after)+1,0,label);
  for(const [label,before]of [['Change Summary','Root Cause'],['Reason for Chg','Sub Reason']])if(!tableOrders.ctlTable.includes(label)&&tableOrders.ctlTable.includes(before))tableOrders.ctlTable.splice(tableOrders.ctlTable.indexOf(before),0,label);
}
const nsrData=createNsrDataService((range,count)=>sourceMessage('nsr',`Loading SAP · ${range.from} to ${range.to} · ${count} rows`,'loading')); let nsrRange=nsrData.sapRange();
const nsrRecords=createNsrRecords(),accountMapping=createFidAccountMapping(),report=createFidSourceReaders(sources);
const gridModule={exports:{}};new Function('module','exports',GM_getResourceText('tabulatorJs'))(gridModule,gridModule.exports);const Tabulator=gridModule.exports;
const filterDefs=[
  ['fAccount','Account','account','common'],['fCurrentQuarter','Current Quarter','currentQuarter','fid'],['fQuarter','Assessment Quarters','quarter','fid'],['fOwner','FID Owner','owner','fidOwner'],
  ['fCustomer','FID Customer','customer','fid'],['fSystem','System Description','system','fid'],['fBusinessUnit','Business Unit','businessUnit','fid'],['fMfg','MFG Upgrade Type','mfg','fid'],['fSalesAdmin','Sales Admin','salesAdmin','fidSalesAdmin'],
  ['fSlotInitiator','Slot Initiator','slot.initiator','fidSlot'],['fCommit','FID Commit','commit','fid'],['fFidStatus','FID Status','fidStatus','fidCtl'],['fRisk','Risk Drivers','riskType','fidRisk'],
  ['fCtlStage','CTL Stage','ctl.stage','ctl'],['fCtlInitiator','CTL Initiator','ctl.initiator','ctl'],['fCtlCreatedDate','CTL Created Date','ctl.created','ctl'],['fCtlChange','CTL Type','ctl.change','ctl'],['fCtlCustomer','CTL Customer','ctl.customer','ctl'],
  ['fNsrInitiator','NSR Initiator','nsr.initiator','nsr'],['fNsrSubmitDate','NSR Submit Date','nsr.submitted','nsr'],
  ['fNsrTaskStatus','NSR Task Status','nsr.taskStatus','nsr'],['fNsrType','NSR Type','nsr.type','nsr'],['fNsrOrg','NSR Org','nsr.org','nsr'],
  ['fNsrSubsystem','NSR Subsystem','nsr.subsystem','nsr'],['fNsrProduct','NSR Product','nsr.product','nsr'],['fNsrCustomer','NSR Customer','nsr.customer','nsr']
];
const dateFilters=['nsr.submitted','ctl.created'];
const FILTER_PRESET_KEY='toolkitKyra.filter.presets.v1.';
const filterPresetSlots=['D','1','2','3'],filterPresetName=slot=>slot==='D'?'Default':'Preset '+slot;
const filterPresetPaths=[...filterDefs.map(([,,path])=>path),'signal'];
function parseFilterPreset(stored){
  if(!stored)return null;
  try{
    const preset=JSON.parse(stored);
    return preset?.filters&&typeof preset.filters==='object'&&!Array.isArray(preset.filters)&&typeof preset.search==='string'&&typeof preset.includeIgnored==='boolean'?preset:null;
  }catch{return null}
}
const savedFilterPresets=filterPresetSlots.map(slot=>{try{return parseFilterPreset(GM_getValue(FILTER_PRESET_KEY+slot,null))}catch{return null}});
function ctlDateRange(period){
  const today=new Date(),year=today.getFullYear(),month=today.getMonth();
  if(period.endsWith('week')){const day=today.getDate()-(today.getDay()+6)%7-(period==='Last week'?7:0);return [new Date(year,month,day),new Date(year,month,day+6)].map(nsrData.localDate)}
  const offset=period==='Last month'?-1:0;return [new Date(year,month+offset,1),new Date(year,month+offset+1,0)].map(nsrData.localDate);
}
const ownerHint="Inferred from CTL/NSR records based on the highest operation count, with the Slot Initiator preferred when counts are equal. For reference only; may be inaccurate.";
const salesAdminHint="Inferred from Pull/Push CTL records based on the highest operation count. For reference only; may be inaccurate.";
const currentQuarterHint="Each FID has one current quarter from current MCSD, with SRD as fallback, on fiscal calendar.";
const quarterHint="Each FID has one or more quarters for stability assessment, including its current MCSD / SRD quarter and historical quarters supported by CTL records.";
const helpIcon=tip=>`<span class="help-icon" title="" aria-label="${escapeHtml(tip)}"></span>`;
const lifecycleNames={'open':'Pending NSR','discarded':'Cancelled / Rejected','completed-5d':'Completed 5D NSR','completed-unlinked':'Completed NSR but no CTL','ctl-pending':'CTL pending approval','implemented':'CTL approved','ctl-unavailable':'CTL status unknown',ignore:'⊘ Ignore','ignore-completed':'⊘ Ignore',important:'★ Star'};
const markLabels={all:'All',normal:'Unmarked',ignore:'⊘ Ignore',important:'★ Star'},markOptions=['important','ignore','normal'];
const defaultFidDetailMarks=['important','normal'];
const completedTasks=Object.fromEntries(['completed-unlinked','ctl-pending','implemented','completed-5d'].map(key=>['15-Completed|'+key,lifecycleNames[key]]));
const defaultCommit=['Commit'];
const defaultTypes=['Config','Slot','Drop Slot','[Unknown]'];
const filterResets={
  quarter:{attribute:'data-reset-quarter',title:'Reset to the default five quarters',values:defaultQuarters},
  commit:{attribute:'data-reset-commit',title:'Reset to Commit',values:()=>[...defaultCommit]},
  'ctl.change':{attribute:'data-reset-ctl-type',title:'Reset to Config, Slot, Drop Slot and No CTL Yet',values:()=>[...defaultTypes]}
};
const pendingTypes={'Pending Config CTL':['Slot','Config'],'Pending Pull/Push CTL':['Pull','Push']};
const overviewColors={total:'#94a3b8',completed:'#16843f',workflow:'#2563eb',draft:'#8a5c0c',rejected:'#c42f48',cancelled:'#52647d',unknown:'#52647d'};
let includeIgnored=true;
function ignoredToggle(view){return `<label class="ignore-scope" title="Exclude Ignore-marked FIDs, their linked records and Ignore NSRs from statistics and Details"><input type="checkbox" data-exclude-ignored="${view}"${includeIgnored?'':' checked'}>Exclude ignored FID/NSR</label>`}
function overviewHead(view,count,total,selected){const unit=view==='portfolio'?'FID':'NSR';return `<div class="section-head overview-head"><h2>${unit} Overview</h2><div class="overview-tools"><span class="badge">${selected?`${number(count)} / ${number(total)} selected ${unit}s`:`${number(count)} filtered ${unit}s`}</span>${ignoredToggle(view)}</div></div>`}
const pendingCtlMarker='⏳';
const ctlStageOrder=['0-Cancelled','0-Rejected','Draft','PM','Sales Ops','NSR PE','CE','MSPM','MS','15-Completed'];
function ctlStageName(value){const stage=text(value).replace(/^\d+\s*-\s*/,'');return ({cancelled:'0-Cancelled',canceled:'0-Cancelled',rejected:'0-Rejected',draft:'Draft',originator:'Draft',completed:'15-Completed'})[stage.toLowerCase()]||stage}
function compareCtlStages(a,b){const rank=value=>!value||value==='[Unknown]'?100:value==='15-Completed'?99:ctlStageOrder.includes(value)?ctlStageOrder.indexOf(value):98;const left=ctlStageName(a),right=ctlStageName(b);return rank(left)-rank(right)||numericCollator.compare(left,right)}
function taskGroup(value,ctl=false){return value==='15-Completed'?'completed':value==='0-Rejected'?'rejected':value==='0-Cancelled'?'cancelled':ctl&&value==='Draft'?'draft':ctl&&(!value||value==='[Unknown]')?'unknown':'workflow'}
function taskOverview(kind,terms,groups,total){
  const unit=kind.toUpperCase(),equation=total+' = '+terms.map(([group])=>groups[group]).join(' + ');
  const card=(group,label,count,color)=>{
    const active=taskSelected(kind,group),percent=total?(100*count/total).toFixed(1):'0.0';
    return `<button class="nsr-kpi" style="--kpi-color:${color}" type="button" data-${kind==='ctl'?'ctl-status':'nsr-group'}="${group}" data-count="${count}" aria-pressed="${active}"${group!=='all'&&!count&&!active?' disabled':''} title="${escapeHtml(group==='all'?equation:`${label} / Total ${unit} = ${count} / ${total} = ${percent}%`)}"><small>${escapeHtml(label)}</small><b>${number(count)}</b><span>${percent}%</span></button>`;
  };
  return card('all','Total '+unit,total,overviewColors.total)+'<span class="math-sign" aria-hidden="true">=</span>'+terms.map(([group,label,color],index)=>(index?'<span class="math-sign" aria-hidden="true">+</span>':'')+card(group,label,groups[group],color)).join('');
}
function taskFilterValues(kind,key){if(key==='all')return [];if(kind==='nsr'&&Object.hasOwn(completedTasks,'15-Completed|'+key))return ['15-Completed|'+key];return [...id(kind==='ctl'?'fCtlStage':'fNsrTaskStatus').options].map(option=>option.value).filter(value=>!Object.hasOwn(completedTasks,value)&&taskGroup(kind==='ctl'?ctlStageName(value):value,kind==='ctl')===key)}
function taskSelected(kind,key,types){const selected=filterValues[kind==='ctl'?'ctl.stage':'nsr.taskStatus']||[],parts=taskFilterValues(kind,key),changes=filterValues['ctl.change']||[];return key==='all'?!selected.length:parts.length>0&&parts.every(value=>selected.includes(value))&&(!types||!changes.length||types.every(value=>changes.includes(value)))}
function selectTaskFilter(kind,key,types){
  const path=kind==='ctl'?'ctl.stage':'nsr.taskStatus',selected=filterValues[path]||[],parts=taskFilterValues(kind,key),active=taskSelected(kind,key,types);
  if(types){
    if(!parts.length)return;
    const changes=taskSelected(kind,key)?filterValues['ctl.change']?.length?filterValues['ctl.change']:[...id('fCtlChange').options].map(option=>option.value):[];
    const next=active?changes.filter(value=>!types.includes(value)):[...new Set([...changes,...types])];
    changeFilter({[path]:active&&!next.length?selected.filter(value=>!parts.includes(value)):[...new Set([...selected,...parts])],'ctl.change':next});
  }else changeFilter(path,key==='all'?[]:active?selected.filter(value=>!parts.includes(value)):[...selected,...parts]);
}
const fidStatusNames={'no-slot':'No Slot','slot-ready':'Slotted FID','stable-slot':'Stable Slot','late-slot':'Late Slot','stable-ctl':'Stable CTL','late-ctl-nsr':'Late CTL','late-ctl':'Late CTL','slot-unknown':'Slot status unknown','slot-date-unknown':'Slot date unknown','ctl-unknown':'CTL status unknown','risk-fid':'Risk FID','risk-free':'Stable Slot & CTL','risk-unknown':'Risk Status Unknown','slot-no-mcsd':'Slot + no MCSD','slot-pending':'No Slot · pending approval','slot-dropped':'Slotted FID · drop slot'};
const fidCellNames={'stable-slot|stable-ctl':'Stable Slot & CTL','late-slot|stable-ctl':'Late Slot with Stable CTL','stable-slot|late-ctl-nsr':'Stable Slot with Late CTL','late-slot|late-ctl-nsr':'Late Slot with Late CTL','stable-slot|ctl-unknown':'Stable Slot · CTL unknown','late-slot|ctl-unknown':'Late Slot · CTL unknown','slot-date-unknown|stable-ctl':'Slot date unknown · Stable CTL','slot-date-unknown|late-ctl-nsr':'Late CTL · Slot unknown','slot-date-unknown|ctl-unknown':'Slot & CTL status unknown'};
const fidCauseNames={'slot-no-mcsd':'Slot + no MCSD','slot-pending':'No Slot · pending approval','slot-dropped':'Slotted FID · drop slot'};
const fidCauseKeys=Object.keys(fidCauseNames);
const fidSelection={order:[],groups:{}};
const css=GM_getResourceText('tabulatorCss')+DASHBOARD_BASE_CSS+WORKSPACE_CSS+`
body.fcid-app{--side:370px;--fid-ink:#2458a6;--fid-tint:#eaf1ff;--fid-line:#bed0ef;--nsr-ink:#b13d53;--nsr-tint:#fceef1;--nsr-line:#e6b6c1;--shadow:0 2px 12px #253c6410;font-size:13px;min-width:0}.fcid-app [hidden]{display:none!important}.fcid-app h1,.fcid-app h2,.fcid-app h3{font-weight:600;line-height:1.3;margin:0}.fcid-app h2{font-size:16px}.fcid-app h3{font-size:13px}.fcid-app .card{padding:14px;min-width:0;border-radius:11px}.muted,.small{font-size:12px;color:#53657d}.badge{font-size:12px;background:#edf2fa;color:#435a7b;border-radius:6px;padding:3px 7px;white-space:nowrap}
#fcToggleSide{font-size:21px;padding:4px 10px;line-height:1.3}
.side-collapsed .layout{grid-template-columns:minmax(0,1fr)}.side-collapsed .side{display:none}
.fc-source-help a{color:#265dbb}.fc-note{font-size:12px;color:#a73047;white-space:pre-line}.fc-note:empty{display:none}
.filter-head{display:flex;align-items:center;justify-content:space-between;margin:17px 0 10px}.fc-filter{margin-top:9px}.filter-head h2{font-size:18px;color:#30445f}.fc-filter[data-filter-needs=common]{--filter-ink:#52647d;--filter-tint:#f4f6fa;--filter-line:#cbd5e1}.fc-filter[data-filter-needs^=fid]{--filter-ink:var(--fid-ink);--filter-tint:var(--fid-tint);--filter-line:var(--fid-line)}.fc-filter[data-filter-needs=nsr]{--filter-ink:var(--nsr-ink);--filter-tint:var(--nsr-tint);--filter-line:var(--nsr-line)}.fc-filter label{display:block;font-size:12px;font-weight:600;color:var(--filter-ink,#435775);margin-bottom:4px}.fc-filter select{font-size:12px;width:100%;min-height:32px;border:1px solid var(--filter-line,#cbd6e6);border-radius:6px;padding:5px;background:#fff;color:var(--filter-ink,#263e60)}.fc-filter .chosen-container,.fc-filter .chosen-container-multi .chosen-choices li.search-field input.chosen-search-input{font-size:12px}.fc-filter .chosen-choices{border:1px solid var(--filter-line,#cbd6e6);border-radius:6px;padding:2px;background:#fff;box-shadow:none;min-height:32px}.fc-filter .chosen-choices li.search-choice{color:var(--filter-ink,#294c88);background:var(--filter-tint,#e8effd);border-color:var(--filter-line,#bacdee);padding:4px 22px 4px 6px}.fc-filter .chosen-results li{font-size:12px;line-height:1.4;padding:5px 7px}.fc-filter .chosen-results li.highlighted{background:var(--filter-ink,#2f64ca);color:#fff}.fc-filter input::placeholder{color:#526782}.fc-filter-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:4px}.fc-filter-head>label{min-width:0;margin-bottom:0}.fc-filter-actions{display:flex;flex:none;gap:6px;justify-content:flex-end}.fc-filter-actions button{border:0;background:none;color:var(--filter-ink,#365c97);font-size:12px;padding:1px 2px;cursor:pointer}.fcid-app .filter-group-title{font-size:16px;font-weight:600;color:var(--fid-ink);border-top:1px solid #dde5f1;margin-top:18px;padding-top:13px}.filter-group-title[data-filter-needs=nsr]{color:var(--nsr-ink);border-top-color:var(--nsr-line)}.fc-filter.primary-filter{border-left:3px solid var(--filter-line,#83a8e8);padding:7px 8px;background:var(--filter-tint,#f1f5fd);border-radius:0 6px 6px 0}.fc-filter.primary-filter .fc-filter-head>label{color:var(--filter-ink,#2f5590)}.quick-search{width:100%;border:1px solid #cbd6e6;border-radius:6px;padding:7px;color:#243b5c;font-size:13px}.quick-search{resize:vertical;min-height:52px}.nsr-year-picks{display:flex;gap:5px;flex-wrap:wrap;margin-top:6px}.nsr-year-picks button{font-size:12px;border:1px solid var(--filter-line,#cbd8e9);border-radius:5px;color:var(--filter-ink,#405a80);background:var(--filter-tint,#f4f7fc);padding:2px 6px;cursor:pointer}.nsr-year-picks button.active{background:var(--filter-ink,#315b96);border-color:var(--filter-ink,#315b96);color:#fff}.nsr-year-picks button.partial{background:#eaf3fc;border-color:#8db3d9;color:#285982}
.fid-lookup{display:flex;align-items:center;gap:7px}.fid-lookup textarea{flex:1;min-width:0;width:0}.fid-lookup .btn{flex:none;padding:6px 8px}.fid-lookup .btn span{display:inline-block;min-width:1.2em}
.fid-lookup .btn[data-state=loading] span{animation:fid-bw-spin .8s linear infinite}.fid-lookup .btn[data-state=done]{color:#23744b}
@keyframes fid-bw-spin{to{transform:rotate(360deg)}}@media(prefers-reduced-motion:reduce){.fid-lookup .btn[data-state=loading] span{animation:none}}
.fc-filter[data-filter-needs] .chosen-drop{border-color:var(--filter-line)}.fc-filter[data-filter-needs] .chosen-container-active .chosen-choices{border-color:var(--filter-ink);box-shadow:0 0 0 2px var(--filter-tint)}
body.fcid-app{--ctl-ink:#7950a1;--ctl-tint:#f3edfa;--ctl-line:#d2bee8}
.fc-filter[data-filter-needs=ctl]{--filter-ink:var(--ctl-ink);--filter-tint:var(--ctl-tint);--filter-line:var(--ctl-line)}.filter-group-title[data-filter-needs=ctl]{color:var(--ctl-ink);border-top-color:var(--ctl-line)}
.ctl-overview{container-type:inline-size}.ctl-flow{display:grid;grid-template-columns:minmax(0,1fr) repeat(var(--ctl-terms),14px minmax(0,1fr));gap:7px;padding:2px 0}.ctl-flow .nsr-kpi{padding:11px 9px}.ctl-flow .nsr-kpi small{overflow:hidden;text-overflow:ellipsis}
:is(#view-changes,#view-nsr) :is(.pivot-head,.pivot-count){background:var(--view-tint)}:is(#view-changes,#view-nsr) .pivot-count{color:var(--view-ink)}
#view-nsr :is(.pivot-card,.pivot-head){border-color:var(--view-line)}#view-nsr .interactive-pivot .badge{color:var(--view-ink);background:var(--view-tint)}
.fc-filter[data-filter-needs=nsr] input[type=date]{border-color:var(--nsr-line);color:var(--nsr-ink)}.fc-filter[data-filter-needs=nsr] :focus-visible{outline-color:var(--nsr-ink)}
.filter-chip:has([data-locate-filter^=fNsr]){--chip-ink:var(--nsr-ink);--chip-tint:var(--nsr-tint);--chip-line:var(--nsr-line)}.filter-chip:has([data-locate-filter^=fCtl]){--chip-ink:var(--ctl-ink);--chip-tint:var(--ctl-tint);--chip-line:var(--ctl-line)}.filter-chip:has(:is([data-locate-filter^=fNsr],[data-locate-filter^=fCtl])){color:var(--chip-ink);background:var(--chip-tint);border-color:var(--chip-line)}.filter-chip:has(:is([data-locate-filter^=fNsr],[data-locate-filter^=fCtl])):hover{border-color:var(--chip-ink)}.filter-chip:has(:is([data-locate-filter^=fNsr],[data-locate-filter^=fCtl])) .filter-chip-remove{color:var(--chip-ink)}.filter-chip:has(:is([data-locate-filter^=fNsr],[data-locate-filter^=fCtl])) .filter-chip-remove:hover{background:var(--chip-line)}
@container(max-width:850px){.ctl-flow{grid-template-columns:16px minmax(0,1fr) 16px minmax(0,1fr)}.ctl-flow>[data-ctl-status=all]{grid-column:1/-1;display:grid;grid-template-columns:1fr auto;align-items:center}.ctl-flow>[data-ctl-status=all] b{grid-column:2;grid-row:1/3;margin:0}}
@media(max-width:760px){.fcid-app .ctl-pivots{grid-template-columns:minmax(0,1fr)}}
#fNsrTaskStatus_chosen{--chosen-group-option-indent:22px}
#fFidStatus_chosen .chosen-results .fid-status-parent{font-weight:600}
#fFidStatus_chosen .chosen-results .fid-status-child{padding-inline-start:22px}

.fc-view{display:grid;gap:12px;min-width:0}.filter-context{display:flex;align-items:center;flex-wrap:wrap;gap:7px 10px;padding:9px 11px;background:#fff;border:1px solid var(--line);border-radius:9px;font-size:12px}.context-title{flex:none;font-weight:600;color:#304d77}.filter-tags{display:flex;align-items:center;min-height:27px;flex:1 1 240px;min-width:0;flex-wrap:wrap;gap:6px}.filter-chip{display:inline-flex;align-items:center;gap:0;max-width:min(360px,100%);min-width:0;border:1px solid #c8d6ed;background:#f5f8ff;color:#31528a;border-radius:16px;padding:0;font-size:12px;line-height:1.4;cursor:pointer}.filter-chip:hover{background:#eaf1ff;border-color:#94b2df}.filter-chip-label{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;border:0;background:none;color:inherit;font:inherit;text-align:left;padding:4px 0 4px 9px;cursor:pointer;border-radius:16px 0 0 16px}.filter-chip-remove{flex:none;align-self:stretch;border:0;background:none;font-size:15px;line-height:1;color:#536d95;padding:0 8px 0 6px;cursor:pointer;border-radius:0 16px 16px 0}.filter-chip-remove:hover{background:#dfe9fb;color:#244c87}.scope-note{margin-left:auto;color:#53657d}
.workspace-nav [data-view=portfolio],#view-portfolio{--view-ink:var(--fid-ink);--view-tint:var(--fid-tint);--view-line:var(--fid-line)}
.workspace-nav [data-view=nsr],#view-nsr{--view-ink:var(--nsr-ink);--view-tint:var(--nsr-tint);--view-line:var(--nsr-line)}
.workspace-nav [data-view=changes],#view-changes{--view-ink:var(--ctl-ink);--view-tint:var(--ctl-tint);--view-line:var(--ctl-line)}
.workspace-nav [data-view=cycle],#view-cycle,#wfModal{--view-ink:#a85c07;--view-tint:#fff4e4;--view-line:#e6c79e}
.workspace-nav [data-view=summary],#view-summary{--view-ink:#3d805d;--view-tint:#edf7ef;--view-line:#c1dcc7}
.workspace-nav button:not([data-view=rules]){background:var(--view-tint);color:var(--view-ink);box-shadow:inset 0 0 0 1px var(--view-line)}.workspace-nav button:not([data-view=rules]):hover{box-shadow:inset 0 0 0 1px var(--view-ink)}.workspace-nav button:not([data-view=rules]).active{background:var(--view-ink);color:#fff;box-shadow:inset 0 0 0 1px var(--view-ink),0 2px 5px #253c6420}
.fc-view :is(.section-head,.section-title) h2,#wfModal .wf-head h1{color:var(--view-ink)}

.signal-strip{display:flex;gap:7px;flex-wrap:wrap}.signal{border:1px solid #d4deef;background:#fff;color:#3e5575;padding:6px 9px;border-radius:7px;font-size:12px;cursor:pointer}.signal b{margin-right:5px;color:#34539d;font-variant-numeric:tabular-nums}
.risk-drivers .signal{border:0;border-radius:0;padding:8px 15px;background:#edf2fa;clip-path:polygon(10px 0,100% 0,100% calc(100% - 10px),calc(100% - 10px) 100%,0 100%,0 10px)}
.risk-drivers .signal b{font:inherit;margin:0;color:inherit}.risk-drivers .signal small{font-size:11px;opacity:.8;white-space:nowrap}
.risk-drivers .signal:hover{background:#e5efff}
.risk-drivers .signal:disabled{opacity:.55;cursor:default}
@media(max-width:760px){.fc-heading .title-tracking{display:none}}
.summary-heading.overview-tools{flex:1;flex-wrap:nowrap;justify-content:flex-start;max-width:100%}.summary-heading>h2{flex:none}.summary-heading .ignore-scope{margin-left:auto}
.overview-tools{display:flex;align-items:center;justify-content:flex-end;gap:8px;min-width:0}.overview-tools>.badge{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.ignore-scope{display:inline-flex;align-items:center;gap:4px;flex:none;font-size:12px;color:#758192;white-space:nowrap;cursor:pointer}.ignore-scope input{margin:0;accent-color:#758192}.section-head.overview-head{flex-wrap:nowrap}.overview-head>h2{flex:none}
.risk-drivers .signal[aria-pressed=true]{background:#3566b0;color:#fff}
.risk-drivers .signal:focus-visible{outline:0;box-shadow:inset 0 0 0 2px currentColor}
.risk-drivers .signal[data-nsr-lifecycle="open"]{background:var(--nsr-tint);color:var(--nsr-ink)}
.risk-drivers .signal[data-nsr-lifecycle="open"]:hover{background:#f8e0e6}
.risk-drivers .signal[data-nsr-lifecycle="open"][aria-pressed=true]{background:var(--nsr-ink);color:#fff}
.risk-drivers .signal[data-group-value="Pending Config CTL"][aria-pressed=true]{background:#745097;color:#fff}
.table-toolbar{display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;margin:10px 0 8px}.detail-head>.table-toolbar,summary>.table-toolbar{margin:0 0 0 auto;justify-content:flex-end}.detail-head>.table-toolbar{flex-wrap:nowrap;flex:0 1 auto}summary>.table-toolbar{float:right}.table-count{font-size:12px;color:#4d6383}.fc-table-shell{min-width:0}.fc-grid.tabulator{overflow:hidden;resize:vertical;min-height:105px;border:1px solid #d5dfed;border-radius:8px;scrollbar-width:thin;background:#fff;font:13px/20px var(--font-sans);color:#344c6b}.fid-dialog .fc-grid.tabulator{min-height:0;resize:none}.fid-dialog .tabulator-tableholder{min-height:34px}.fc-grid .tabulator-header{background:var(--table-heading-bg,#edf2fb);border-bottom:1px solid #d5dfed;color:#344c6b;font-weight:600}.fc-grid .tabulator-header .tabulator-col{background:var(--table-heading-bg,#edf2fb);border-right:1px solid #e0e7f1;font-size:12px}.fc-grid .tabulator-header .tabulator-col .tabulator-col-content{padding:7px 9px}.fc-grid .tabulator-row{background:#fff;color:#344c6b;min-height:34px}.fc-grid .tabulator-row.tabulator-row-even{background:#f8faff}.fc-grid .tabulator-row:hover{background:#f0f5fc}.fc-grid .tabulator-row .tabulator-cell{padding:6px 9px;height:34px;line-height:20px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border-right:1px solid #e8edf5;border-bottom:1px solid #e3e9f3}.fc-grid .tabulator-row .tabulator-cell.tabulator-frozen{background:inherit;border-right:1px solid #d6e0ed}.fc-grid .tabulator-tableholder{scrollbar-width:thin;overflow-anchor:none}.cell-content{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-height:22px}.fid-status-tags{display:inline-flex;gap:4px;vertical-align:top}#portfolioTable .tabulator-cell:has(.fid-status-tags){cursor:pointer}.fid-status-tags .pill{flex:none}.cell-content .pill,.cell-content .status-pill{max-width:100%;white-space:nowrap;vertical-align:top}.fc-link{border:0;background:none;color:#215fbb;text-decoration:underline;text-underline-offset:2px;cursor:pointer;padding:0;font-size:13px;white-space:nowrap}.pill{display:inline-block;font-size:12px;line-height:17px;padding:2px 6px;border-radius:5px;background:#edf2fa;color:#3e5a80}.pill.risk,.pill.alert{background:#ffe5eb;color:#a12d47}.pill.clear{background:#ddf4e7;color:#196747}.pill.unknown,.fid-detail-value.no-slot,.fid-evidence>.notice.no-slot{background:#fff1cf;color:#7c5514}.pill.info{background:#e7effb;color:#375b91}.fcid-app .status-pill{padding:2px 6px;line-height:17px}
.fc-table-dialog[open]{display:flex;flex-direction:column;gap:10px;position:fixed;inset:0;width:100%;height:100%;max-width:none;max-height:none;margin:0;padding:clamp(8px,2vw,18px);border:0;background:#fff;color:inherit;overflow:hidden}
.fc-table-dialog>.detail-head{margin:0}.fc-table-dialog>div:has(>.fc-table-shell),.fc-table-dialog .fc-table-shell{display:flex;flex:1;flex-direction:column;min-height:0}
.fc-table-dialog .col-bar{flex:none;margin-bottom:8px}.fc-table-dialog .fc-grid.tabulator{flex:1;min-height:0;height:auto!important;resize:none}body:has(.fc-table-dialog[open]){overflow:hidden}
.table-expand{display:grid;place-items:center;min-height:26px}.table-expand svg{display:block;width:15px;height:15px;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
.table-expand[aria-expanded=false] .restore-icon,.table-expand[aria-expanded=true] .expand-icon{display:none}
.record-mark-toggle{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;margin:0;padding:0;border:1px solid #a9b3c0;border-radius:3px;background:transparent;color:#7d8a9d;font:600 13px/1 var(--font-sans);vertical-align:middle;cursor:pointer}.record-mark-toggle svg{display:block;width:12px;height:12px;flex:none;fill:none;stroke:currentColor;stroke-width:1.2}.record-mark-toggle[data-record-mark=ignore]{background:#edf0f4}.record-mark-toggle[data-record-mark=important]{border-color:#d99aa5;background:#fff1f3;color:#b8324a}.record-mark-toggle:focus-visible{outline:2px solid #557cae;outline-offset:2px}.fc-grid .tabulator-row[data-record-mark=ignore] .tabulator-cell,.fc-grid .tabulator-row[data-record-mark=ignore] .tabulator-cell *{color:#929aa5}.fc-grid .tabulator-row[data-record-mark=important] .tabulator-cell,.fc-grid .tabulator-row[data-record-mark=important] .tabulator-cell *{color:#b8324a}.pill.ignore{background:#edf0f4;color:#667184}
.fid-dialog .tabulator-row:is([data-fid-risk="pending-nsr"],[data-fid-risk="pending-config"]),.pill.pending,.risk-drivers .signal[data-group-value="Pending Config CTL"]:not([aria-pressed=true]){background:#f2edfb;--fid-history-line:#aa83c6}
.pill.pending,.risk-drivers .signal[data-group-value="Pending Config CTL"]:not([aria-pressed=true]){color:#745097}
.fc-grid .tabulator-header-filter{display:none}.fc-grid.column-filters .tabulator-header .tabulator-col.tabulator-sortable .tabulator-col-title{padding-right:29px}.fc-grid.column-filters .tabulator-header .tabulator-col .tabulator-col-content .tabulator-col-sorter{right:0}.fc-grid.column-filters .tabulator-col-sorter .tabulator-arrow{border-left-width:4px;border-right-width:4px}.column-heading>span{display:block;overflow:hidden;text-overflow:ellipsis}.column-funnel,.mark-trigger{position:absolute;right:11px;top:50%;transform:translateY(-50%);display:inline-flex;align-items:center;justify-content:center;width:16px;height:18px;padding:2px;border:0;border-radius:3px;background:none;color:#8a97aa;cursor:pointer}.column-funnel svg,.mark-trigger svg{width:12px;height:12px;fill:currentColor}.column-funnel:hover,.column-funnel.active,.mark-trigger:hover,.mark-trigger.active{background:#dceaff;color:#215fbb}.column-funnel:focus-visible,.mark-trigger:focus-visible{outline:2px solid #557cae;outline-offset:-2px}
.col-bar{display:flex;align-items:center;flex-wrap:wrap;gap:4px;flex:1 1 160px;min-width:0}.col-bar:empty{display:none}.col-bar .pill{max-width:min(220px,100%);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;border:0;cursor:pointer}
.col-popup{position:fixed;position-area:bottom;position-try-fallbacks:flip-block;inset:auto;margin:8px;width:min(240px,calc(100vw - 16px));color:#344c6b;box-shadow:0 6px 22px #213a612b}.col-popup label>span{font-weight:600}.col-popup input{display:block;width:100%;margin:7px 0 10px;padding:7px;border:1px solid #c5d2e5;border-radius:5px}.col-popup .btnrow{justify-content:space-between}
#colChoices,#colNoteFilter{width:100%;max-height:160px;margin:0 0 10px;padding:3px;border:1px solid #c5d2e5;border-radius:5px;color:inherit;background:#fff;font:inherit}#colChoices option{padding:3px 6px}#colNoteFilter{padding:7px}
.mark-trigger{position:static;transform:none;display:flex;margin:auto;height:20px}.col-popup.mark-popup{position-area:bottom span-right;justify-self:start;position-try-fallbacks:flip-block,flip-inline;margin:8px 0;padding:7px;width:190px}.mark-popup label{display:flex;align-items:center;gap:9px;width:100%;padding:7px 9px;border-radius:5px;color:#344c6b;cursor:pointer}.mark-popup label:first-child{border-bottom:1px solid #e0e7f1;margin-bottom:4px}.mark-popup label:hover{background:#f0f5fc}.mark-popup label:has(input:checked){background:#e7effb;color:#215fbb}.mark-popup input[type=checkbox]{width:14px;height:14px;margin:0;padding:0;accent-color:#2863d6;flex:none}.mark-popup label>span{flex:1;font-weight:400}.mark-popup b{font-weight:400;font-variant-numeric:tabular-nums}
.filter-head .filter-presets{display:flex;gap:3px;margin-left:auto;margin-right:7px}
.filter-presets button{display:grid;place-items:center;width:23px;height:23px;padding:0;border:1px solid #cdd8e8;border-radius:4px;background:#fff;color:#526b8d;font-size:12px;font-weight:600;cursor:pointer}
.filter-presets button[data-saved=true]{background:#3566b0;border-color:#3566b0;color:#fff}
.filter-presets button:hover{border-color:#3566b0}
.col-popup.filter-preset-popup{position-area:bottom span-right;justify-self:start;position-try-fallbacks:flip-inline,flip-block,flip-inline flip-block;width:max-content;max-width:calc(100vw - 20px);padding:9px}
.filter-preset-popup p{margin:0 0 7px;font-weight:600;white-space:nowrap}
.filter-preset-popup .btnrow{justify-content:flex-end;gap:5px}
.fid-date-change{padding:0;border:0;background:none;color:inherit;font:inherit;white-space:inherit;text-align:left;text-decoration:underline dotted;text-underline-offset:2px;cursor:pointer}
.fc-grid .fid-date-change{display:block;max-width:100%;overflow:hidden;text-overflow:ellipsis}
.fid-detail-value:is([data-date-field="mcsd"],[data-date-field="rd"]){position:relative}
.fid-date-change.is-unread::after{content:"";position:absolute;inset:0 0 auto auto;border-top:8px solid #b86013;border-left:8px solid transparent}
.col-popup.fid-date-popup{position-area:bottom span-right;justify-self:start;position-try-fallbacks:flip-block,flip-inline,flip-block flip-inline;width:min(300px,calc(100vw - 16px));max-height:min(320px,80vh);padding:10px 12px;overflow:auto}
.fid-date-popup h3{margin:0 0 4px;font-size:13px}
.fid-date-popup ol{list-style:none;margin:0 0 7px;padding:0}
.fid-date-popup li{padding:5px 0;border-bottom:1px solid #e1e8f1}
.fid-date-popup li strong,.fid-date-popup li small{display:block}
.fid-date-popup li strong{font-weight:600;font-variant-numeric:tabular-nums}
.fid-date-popup li small{color:#526b8d}
.fid-date-popup [role=status]{margin:0;color:#a63436}
.fid-dialog .tabulator-row[data-fid-risk^="late-"],.fid-header-status .pill.risk{background:#fff2f5;--fid-history-line:#c75970}
.fid-dialog .tabulator-row:is([data-fid-risk="pending-nsr"],[data-fid-risk="pending-config"]):hover,.risk-drivers .signal[data-group-value="Pending Config CTL"]:not([aria-pressed=true]):hover{background:#e8ddf5}
.fid-dialog .tabulator-row[data-fid-risk^="late-"]:hover{background:#ffe0e8}
.fid-dialog .tabulator-row:is([data-fid-risk="pending-pull-push"],[data-fid-risk="drop-slot"]),.pill.neutral,.risk-drivers .signal:is([data-group-value="Pending Pull/Push CTL"],[data-group-value="MCSD pending"],[data-group-value="NCE / FA signal"]):not([aria-pressed=true]){background:#f1f3f6;--fid-history-line:#a2abb8}
.pill.neutral,.risk-drivers .signal:is([data-group-value="Pending Pull/Push CTL"],[data-group-value="MCSD pending"],[data-group-value="NCE / FA signal"]):not([aria-pressed=true]){color:#626e7c}
.pill.commit-status,.fc-grid .tabulator-row[data-record-mark] .tabulator-cell .pill.commit-status{background:#293241;color:#fff}
.fid-dialog .tabulator-row:is([data-fid-risk="pending-pull-push"],[data-fid-risk="drop-slot"]):hover,.risk-drivers .signal:is([data-group-value="Pending Pull/Push CTL"],[data-group-value="MCSD pending"],[data-group-value="NCE / FA signal"]):not([aria-pressed=true]):hover{background:#e4e8ed}
.risk-drivers .signal:is([data-group-value="Pending Pull/Push CTL"],[data-group-value="MCSD pending"],[data-group-value="NCE / FA signal"])[aria-pressed=true]{background:#697586;color:#fff}
.fid-header-status .pill:is(.pending,.risk,.neutral){border-left:3px solid var(--fid-history-line)}
.active-ctl-dot{display:inline-block;width:6px;height:6px;margin-left:4px;border-radius:50%;background:#16843f;vertical-align:middle}
.fid-dialog .tabulator-row[data-fid-risk]:not([data-fid-risk=""]) .tabulator-cell:first-child{background-image:linear-gradient(var(--fid-history-line),var(--fid-history-line));background-size:3px 100%;background-position:0 0;background-repeat:no-repeat}
.fid-dialog .tabulator-row.nsr-ctl-match{--fid-link-style:solid;outline:2px var(--fid-link-style) #3b82f6;outline-offset:-2px}
.fid-dialog .tabulator-row.nsr-ctl-note{--fid-link-style:dashed}
.fid-dialog .nsr-ctl-match .tabulator-cell:is(.tabulator-frozen,:first-child)::after{content:"";position:absolute;inset:0;pointer-events:none;border-top:2px var(--fid-link-style) #3b82f6;border-bottom:2px var(--fid-link-style) #3b82f6}
.fid-dialog .nsr-ctl-match .tabulator-cell:first-child::after{border-left:2px var(--fid-link-style) #3b82f6}
.fid-dialog>.section-head>div{min-width:0;flex:1}.fid-header-status{display:inline-flex;flex-wrap:wrap;gap:6px;min-width:0;max-width:100%;font-weight:400}.fid-header-status .pill{padding:3px 8px;font-size:13px;line-height:20px;white-space:nowrap}
.fid-dialog>.section-head>[data-export-fid]{align-self:flex-start;flex:none;margin-top:-4px}
.fid-detail-value.no-slot,.fid-evidence>.notice.no-slot{border-color:#e4c381}.fid-detail-value.no-slot small{color:inherit}
.interactive-pivot>summary{list-style:none;pointer-events:none}.interactive-pivot>summary::-webkit-details-marker{display:none}.interactive-pivot:not([open])>summary{margin-bottom:0}.pivot-toggle{display:inline-block;margin-left:7px;padding:2px 4px;color:#53657d;font-size:12px;cursor:pointer;pointer-events:auto}.pivot-toggle::before{content:"▶"}.interactive-pivot[open] .pivot-toggle::before{content:"▼"}
#nsrTable .fid-sorted .tabulator-row{background:#fff}#nsrTable .fid-sorted .tabulator-row[data-fid-band="1"]{background:#edf3fc}#nsrTable .fid-sorted .tabulator-row .tabulator-cell{border-bottom-color:transparent}#nsrTable .fid-sorted .tabulator-row[data-fid-start="true"] .tabulator-cell{box-shadow:inset 0 1px #b8c0cb}#nsrTable .fid-sorted .tabulator-row:hover{background:#e3edfc}
.nsr-overview{container-type:inline-size}.nsr-flow{display:grid;grid-template-columns:minmax(max-content,calc(20% - 110px)) 16px minmax(max-content,calc(20% - 110px)) repeat(3,16px minmax(0,1fr));gap:7px;padding:2px 0}.nsr-equation,.nsr-breakdown{display:contents}.nsr-kpi{min-width:0;text-align:left;border:1px solid #d6e1ed;border-top:3px solid var(--kpi-color);border-radius:9px;background:#fff;padding:11px;color:var(--kpi-color);cursor:pointer}.nsr-kpi small{display:block;font-size:12px;font-weight:600;color:#435774;white-space:nowrap}.nsr-kpi b{display:block;font-size:28px;line-height:1.2;font-weight:600;margin:6px 0 3px}.nsr-kpi span{font-size:12px;color:#53657d}.math-sign{align-self:center;justify-self:center;font-size:22px;color:#52657f;font-weight:600}.nsr-completed-arrow{grid-column:3;grid-row:2;position:relative;margin-top:12px;color:#78a98b}.nsr-completed-arrow::before{content:"";position:absolute;left:50%;right:0;top:-19px;bottom:50%;border-left:1.5px solid;border-bottom:1.5px solid;border-bottom-left-radius:12px}.nsr-completed-arrow::after{content:"";position:absolute;right:0;top:50%;width:6px;height:6px;border-top:1.5px solid;border-right:1.5px solid;transform:translateY(-50%) rotate(45deg)}.nsr-breakdown-equals{grid-column:4;grid-row:2;margin-top:12px;color:#16843f}.nsr-stage-sum{position:relative;grid-column:5/-1;grid-row:2;display:flex;gap:7px;min-width:0;margin-top:12px}.nsr-stage-sum::before{content:"";position:absolute;top:-10px;left:-23px;right:0;border-top:1px solid #d9dee4;pointer-events:none}.nsr-stage-sum>.math-sign{flex:none;width:16px}.nsr-stage-spacer{display:none}.nsr-stage{flex:1;display:flex;flex-direction:column}.nsr-stage:first-of-type{flex:1.3}.nsr-stage small{white-space:normal;min-height:2.8em}.nsr-stage b{margin-top:auto}.nsr-stage.alert{background:#fff2f5;border-color:#ecb0bf}.nsr-stage.alert small,.nsr-stage.alert span{color:#992f49}.nsr-stage.approved{background:#f1fbf5}.nsr-stage.pending{background:#fff9ed}.nsr-stage.unknown{background:#f3f5f9}.nsr-pivots{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}.pivot-card{border:1px solid #d6dff0;border-radius:9px;overflow:hidden;min-width:0}.pivot-head{display:flex;justify-content:space-between;gap:7px;padding:9px 10px;background:#f5f8fc;border-bottom:1px solid #dce5ef}.pivot-head span{font-size:12px;color:#52657f;white-space:nowrap}.pivot-list{height:320px;min-height:100px;padding:8px 10px;overflow:auto;resize:vertical;scrollbar-width:thin}.pivot-row{display:grid;grid-template-columns:minmax(85px,1fr) minmax(20px,.8fr) 42px;align-items:center;gap:7px;padding:5px 0;font-size:12px}.pivot-label{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#405775}.pivot-bar{height:5px;background:#eaf0fa;border-radius:4px;overflow:hidden}.pivot-bar i{display:block;height:100%;background:var(--bar);border-radius:4px}.pivot-count{border:0;border-radius:4px;background:#edf1ff;color:#3e5b9a;font-size:12px;padding:3px;cursor:pointer}
.nsr-kpi[aria-pressed=true]{outline:2px solid var(--kpi-color);outline-offset:2px}
.nsr-kpi[data-fid-status=all],.nsr-kpi[data-nsr-group=all],.nsr-kpi[data-ctl-status=all]{background:#fff;color:#344c6b;border-top:1px solid #d6e1ed}
.fid-flow{overflow-x:auto;padding:6px}
.fid-flow-inner{display:grid;grid-template-columns:120px;grid-auto-rows:100px;gap:0 8px}
.fid-flow-inner.has-breakdown{grid-template-columns:120px 16px minmax(156px,max-content) 16px minmax(180px,.85fr) 28px minmax(392px,2fr);grid-template-rows:auto 24px 136px 28px;min-width:1040px;padding-top:56px}
.fid-flow-inner.has-slot-unknown{min-width:1152px}
.fid-equation{display:contents}
.fid-equation>*{grid-row:3}
.fid-flow-inner:not(.has-breakdown) .fid-equation>*{grid-row:1}
.fid-slot-sum{grid-column:7;grid-row:3;display:grid;grid-template-columns:repeat(var(--matrix-cols),minmax(0,1fr));min-width:0;position:relative;gap:1px;border:1px solid #b9cadd;border-radius:10px;background:#cbd7e5}
.fid-axis-item{position:relative;min-width:0;display:flex}
.fid-origin-pre{grid-column:6;grid-row:3;display:grid}.fid-origin-post{grid-column:5;grid-row:2;display:grid}
.fid-breakdown{grid-column:5;grid-row:1;position:relative;display:grid;grid-template-rows:repeat(var(--matrix-rows),minmax(108px,1fr));gap:1px;min-width:0;border:1px solid #b9cadd;border-radius:10px;background:#cbd7e5}
.fid-axis-item>.nsr-kpi{width:100%;min-width:0;border:0;border-radius:0;background:#f4f7fb;padding:10px;box-shadow:none}
.fid-breakdown .fid-axis-item:first-child>.nsr-kpi{border-radius:9px 9px 0 0}.fid-breakdown .fid-axis-item:last-of-type>.nsr-kpi{border-radius:0 0 9px 9px}
.fid-slot-sum .fid-axis-item:first-child>.nsr-kpi{border-radius:9px 0 0 9px}.fid-slot-sum .fid-axis-item:nth-last-of-type(1)>.nsr-kpi{border-radius:0 9px 9px 0}
.fid-matrix{grid-column:7;grid-row:1;position:relative;min-width:0;isolation:isolate}
.fid-matrix-grid{display:grid;height:100%;grid-template-columns:repeat(var(--matrix-cols),minmax(0,1fr));grid-template-rows:repeat(var(--matrix-rows),minmax(108px,1fr));gap:1px;border:1px solid #b9cadd;border-radius:10px;overflow:hidden;background:#cbd7e5}
.fid-risk-outline{position:absolute;inset:0;width:100%;height:100%;pointer-events:none;overflow:visible;z-index:1}
.fid-risk-outline path{fill:none;stroke:#c75970;stroke-width:1.6;stroke-linejoin:round;vector-effect:non-scaling-stroke}
.fid-risk-label{position:absolute;right:14px;top:14px;transform:translateY(-100%);z-index:3;display:flex;flex-direction:column;align-items:flex-start;gap:1px;border:1px solid #dc9eac;border-radius:7px;padding:6px 12px;background:#fff5f6;color:#a93049;text-align:left;white-space:nowrap;font:600 16px/1.4 var(--font-sans);cursor:pointer}
.fid-risk-label .fid-percent{color:#53657d;font:400 14px/1.4 var(--font-sans)}
.fid-matrix-cell{min-width:0;position:relative;display:flex;flex-direction:column;justify-content:center;align-items:flex-start;padding:9px 12px;border:0;background:#fff5f6;color:#ad3048;font:inherit;text-align:left;cursor:pointer}
.fid-matrix-cell[data-assessment=risk-free]{background:#eff9f3;color:#17683e}
.fid-matrix-cell[data-assessment=risk-unknown]{background:#f3f5f8;color:#536275}
.fid-matrix-cell small{font-size:12px;line-height:1.45;font-weight:600;white-space:normal}.fid-matrix-cell b{font-size:26px;line-height:1.2;font-weight:600;margin:2px 0}.fid-matrix-cell .fid-percent{font-size:14px;line-height:1.4;color:#52657e;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%}
.fid-selection-outline{position:absolute;inset:0;width:100%;height:100%;overflow:visible;pointer-events:none;z-index:2}
.fid-overview :is(.nsr-kpi,.fid-risk-label)[aria-pressed=true],.fid-no-slot:has(>.fid-no-slot-main[aria-pressed=true]){outline:2px solid #3566b0;outline-offset:2px}
.fid-slotted:has(>.fid-slotted-main[aria-pressed=true]){outline:2px solid #c75970;outline-offset:2px}
.fid-overview button:is(.fid-aggregate,.fid-risk-label)[aria-pressed=true]{outline-color:#c75970}
.fid-overview [data-fid-status=all][aria-pressed=true]{outline-color:#94a3b8}
.fid-axis-item:has(>.nsr-kpi[aria-pressed=true]){z-index:2}
.fid-overview :is([data-fid-status],[data-fid-cell],[data-fid-region]):focus-visible{outline:2px solid #6f839a;outline-offset:-3px}
.fid-row-equals{grid-column:6;grid-row:1;display:grid;grid-template-rows:repeat(var(--matrix-rows),minmax(108px,1fr));gap:1px;padding:1px 0}.fid-col-equals{grid-column:7;grid-row:2;display:grid;grid-template-columns:repeat(var(--matrix-cols),minmax(0,1fr));gap:1px;padding:0 1px}
.fid-relation-sign{visibility:hidden;opacity:0;pointer-events:none;z-index:4;font:600 22px/28px var(--font-sans);color:#294b7b;text-align:center;min-width:28px;background:#fff;border-radius:50%;justify-self:center;align-self:center}
.fid-relation-sign.is-visible{visibility:visible;opacity:1}
.fid-relation-sign.floating{position:absolute;transform:translate(-50%,-50%)}
.fid-pre-axis{grid-column:7;grid-row:4;position:relative;display:flex;align-items:center;justify-content:center;color:#385b88;font-size:12px}
.fid-post-axis{grid-column:4;grid-row:1;position:relative;display:flex;align-items:center;justify-content:center;color:#385b88;font-size:12px}
.fid-pre-axis::before,.fid-post-axis::before{content:'';position:absolute;background:#9cb1ca}.fid-pre-axis::before{height:1px;left:0;right:4px}.fid-post-axis::before{width:1px;top:4px;bottom:0}
.fid-pre-axis::after,.fid-post-axis::after{content:'';position:absolute;width:6px;height:6px;border-top:1px solid #718fb2;border-right:1px solid #718fb2}.fid-pre-axis::after{right:2px;transform:rotate(45deg)}.fid-post-axis::after{top:2px;transform:rotate(-45deg)}
.fid-pre-axis span,.fid-post-axis span{position:relative;background:#fff;padding:0 10px;white-space:nowrap}.fid-post-axis span{writing-mode:vertical-rl;transform:rotate(180deg);padding:10px 0}
.fid-flow .nsr-kpi{position:relative;display:flex;flex-direction:column;justify-content:center;padding:10px}
.fid-flow .nsr-kpi b{font-size:26px;margin:4px 0 2px;white-space:nowrap}
.fid-flow .nsr-kpi small{white-space:normal}
.fid-flow .nsr-kpi span{display:block}
.fid-flow .nsr-kpi .fid-percent{font-size:14px;line-height:1.4}
.fid-flow [data-fid-status=all]{border-top-width:3px}
.fid-equation>.nsr-kpi{display:flex;justify-content:flex-start;padding-top:12px}
.fid-flow-inner.has-breakdown .fid-equation>:is([data-fid-status=all],.fid-no-slot){height:164px;align-self:end}
.fid-flow-inner.has-breakdown .fid-equation>:is([data-fid-status=all],.fid-no-slot) small{margin-bottom:20px}
.fid-flow [data-fid-status=slot-ready] .fid-percent{white-space:nowrap}
.fid-slot-sum .fid-percent,.fid-breakdown .fid-percent{font-size:14px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.4}
.fid-flow .fid-slotted{background:#eff5ff;border-color:#9fbae7;justify-content:center;padding-top:10px;--subcount-line:#cbdaf0;--subcount-ink:#456487}
.fid-no-slot{cursor:default;background:#fff1cf;border-color:#e4c381;border-top-color:#c58b2a}.fid-no-slot .fid-no-slot-main small,.fid-no-slot .fid-percent{color:var(--kpi-color)}
.fid-no-slot-main,.fid-slotted-main{position:relative;display:flex;flex-direction:column;align-items:stretch;width:100%;border:0;background:none;padding:0;color:inherit;text-align:left;font:inherit;cursor:pointer}
.fid-slotted-main{flex:1;justify-content:center}
.fid-flow .fid-subcount{margin-top:4px;border-top:1px solid var(--subcount-line,#e4c381);padding-top:3px;font-size:11px;line-height:1.45;color:var(--subcount-ink,#8d651e);white-space:nowrap}
.fid-slotted .fid-subcount{position:absolute;left:10px;right:10px;bottom:4px}
.fid-subcount-filter{display:block;width:100%;border:0;border-radius:3px;background:none;padding:0;color:inherit;text-align:left;font:inherit;white-space:nowrap;cursor:pointer}
.fid-subcount-filter[aria-pressed=true],.fid-subcount:has(>button[aria-pressed=true]):not(:has(>button[aria-pressed=false])){outline:2px solid #3566b0;outline-offset:1px}
.fid-subcount:not(:has(>button[aria-pressed=false]))>button[aria-pressed=true]{outline:none}
.fid-overview .fid-subcount>button[aria-pressed=true]:focus-visible{outline:2px solid #6f839a;outline-offset:-2px}
.fid-subcount-filter:hover{color:#405972;text-decoration:underline;text-underline-offset:2px}
.fid-waiting{margin-top:12px;font-size:12px;color:#7c5514}
.notice{display:flex;justify-content:space-between;align-items:center;gap:10px;padding:10px 13px;border:1px solid #e6a5b5;border-left:4px solid #c43c5c;background:#fff0f3;color:#992f48;border-radius:9px;font-size:13px}.notice small{display:block;font-size:12px;margin-top:3px}.statistics-state{margin-top:8px;color:#196b50;font-size:12px}.statistics-state.error{color:#ac3049}.ag-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}.marked-settings{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.nsr-mark-editors{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.ag-box label{display:block;font-size:13px;font-weight:600;margin:6px 0}.ag-box textarea{width:100%;min-height:150px;resize:vertical;border:1px solid #cbd7e9;border-radius:7px;padding:8px;color:#345073}.ag-notice{white-space:pre-line;background:#fff3d7;color:#805b16;border:1px solid #debf79;border-radius:7px;padding:10px;margin:10px 0;font-size:13px}
.fc-nsr-dialog,.fid-dialog,.fc-note-dialog{border:1px solid #b7c7e0;border-radius:12px;padding:17px;background:#f6f8fe;color:#2e486c;box-shadow:0 16px 70px #1e355550}.fc-nsr-dialog{width:min(450px,94vw)}.fc-nsr-dialog::backdrop,.fid-dialog::backdrop,.fc-note-dialog::backdrop{background:#152e505c}.fc-range{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:14px}.fc-range input{width:100%;padding:8px;border:1px solid #c2d0e4;border-radius:6px;margin-top:5px}.fc-actions{display:flex;justify-content:flex-end;gap:7px;margin-top:12px}.fid-dialog{width:min(1400px,96vw);height:min(860px,92vh);max-width:96vw;max-height:94vh;min-width:min(600px,94vw);min-height:300px;overflow:auto;resize:both}.fid-dialog>.section-head{position:sticky;top:-17px;background:#f6f8fe;z-index:10;padding:10px 0}.fid-dialog h1{display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-size:20px}.fid-action{display:inline-flex;align-items:center;padding:5px 9px;border:1px solid #c5d2e5;border-radius:7px;background:#fff;color:#215fbb;font-size:12px;line-height:1.3;text-decoration:none;white-space:nowrap;cursor:pointer}.fid-close{align-self:flex-start;margin:-8px -8px 0 0;width:36px;height:36px;padding:0;border:0;border-radius:7px;background:transparent;color:#526b8d;font-size:28px;line-height:1;cursor:pointer}.fid-action:hover,.fid-close:hover{background:#e7edf8}.fid-dialog-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}.fid-detail-value{padding:9px;background:#fff;border:1px solid #d6e0ef;border-radius:7px;font-size:13px;overflow-wrap:anywhere}.fid-detail-value small{display:block;font-size:12px;color:#526b8d;margin-bottom:3px}.fid-evidence{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:12px 0}.fid-evidence>.notice{display:block;min-width:0;overflow-wrap:anywhere}.fid-evidence>.notice.clear{border-color:#a8d5bb;border-left-color:#2d8b60;background:#eef9f2;color:#196747}.fid-evidence>.notice.unknown{border-color:#d2dcee;border-left-color:#91a2ba;background:#fff;color:#526b8d}.fid-evidence b{font-weight:600}.fid-evidence p{margin:7px 0 0}.fid-dialog h2{margin:0}.fid-dialog .table-toolbar{margin:16px 0 8px}.fc-toast{position:fixed;right:20px;bottom:20px;z-index:150;background:#244a79;color:#fff;padding:11px 14px;border-radius:9px;font-size:13px;max-width:min(480px,90vw);display:flex;align-items:center;gap:13px}.fc-toast button{color:#fff;border:1px solid #b7cceb;border-radius:5px;background:none;padding:4px 7px;cursor:pointer}.fcid-app .cycle-shell{display:grid;gap:12px}.fcid-app #cycleTab{display:block}.fcid-app #cycleTab:not(.active){display:none}.fcid-app .cycle-status,.fcid-app .cycle-label,.fcid-app .context-badge{font-size:12px;color:#425c7f}body:has(#workspace.wf-open){overflow:hidden}
.help-icon{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;box-sizing:border-box;margin-left:5px;vertical-align:middle;border:1px solid #b6c5da;border-radius:50%;color:#6c809d;font-size:10px;font-weight:600;line-height:1;cursor:help}.help-icon::before{content:"?"}.help-icon:hover{color:#2458a6;border-color:#6d90c2;background:#edf3ff;outline:2px solid #c0d2ed;outline-offset:2px}
.help-tip{position:fixed;inset:auto;margin:0;width:max-content;max-width:min(340px,calc(100vw - 16px));padding:9px 11px;border:1px solid #c5d2e5;border-radius:7px;background:#fff;color:#344c6b;box-shadow:0 3px 14px #253c6426;font:12px/1.5 var(--font-sans);pointer-events:none}
.detail-head{justify-content:flex-start;flex-wrap:wrap}.detail-head h2{white-space:nowrap}.fid-note-cell{display:block;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:left}.fid-note-add{display:inline-block;min-width:28px;border:1px solid #c5d2e5;border-radius:4px;background:#f5f8fd;line-height:18px;text-align:center;text-decoration:none}.fid-note-panel{margin:12px 0;border:1px solid #d2dcee;border-radius:8px;background:#fff}.fid-note-view{display:block;width:100%;padding:10px 12px;border:0;border-radius:7px;background:none;color:inherit;text-align:left;cursor:text}.fid-note-title{display:block;margin-bottom:6px;font-size:17px;font-weight:600}.fid-note-panel>form{padding:10px 12px}.fid-note-form .section-head{margin-bottom:6px}.fid-note-text{display:block;white-space:pre-wrap;overflow-wrap:anywhere;font-size:13px;line-height:1.5}.fid-personal-mark{display:inline-flex;align-items:center}.fc-note-dialog{position:fixed;inset:auto;margin:0;width:min(520px,calc(100vw - 16px));max-width:calc(100vw - 16px);max-height:calc(100dvh - 16px);overflow:auto}#fidNoteForm textarea{width:100%;min-height:180px;resize:vertical;border:1px solid #c2d0e4;border-radius:7px;padding:9px;font:13px/1.5 var(--font-sans);color:#263e60;background:#fff}.fc-note-actions{display:flex;align-items:center;justify-content:flex-end;gap:10px;margin-top:10px}
.fid-custom-grid{display:grid;grid-template-columns:minmax(0,1fr) clamp(85px,30%,120px) 60px 32px;align-items:center;gap:7px 10px;max-width:600px}
.fid-custom-grid>span{font-weight:600}.fid-custom-grid>span:last-of-type,.fid-custom-grid input[type=checkbox]{justify-self:center}
.fid-custom-grid input[type=text]{grid-column:1}
.fid-custom-grid input[type=text],.fid-custom-grid select,#fidNoteForm input[type=number]{width:100%;min-width:0;padding:6px 8px;border:1px solid #c2d0e4;border-radius:6px;background:#fff;color:inherit;font:inherit}
.fid-custom-grid [data-custom-row]{display:contents}.fid-custom-grid input[type=checkbox]{accent-color:var(--fid-ink)}
[data-add-fid-custom]{margin-top:8px}
.fid-custom-grid .btn{display:grid;place-items:center;width:32px;height:32px}
.fid-custom-grid .btn:hover{background:#fff0f3;border-color:#e6a5b5;color:#ac3049}
#fidCustomDeleteTitle{overflow-wrap:anywhere}#fidCustomDeleteDialog .btn.delete{background:#ac3049}
.fid-custom-flag{width:100%;min-height:20px;text-align:center;text-decoration:none;color:#23744b;font-weight:600;border-radius:4px}
.fid-custom-flag:hover,.fid-custom-flag:focus-visible{background:var(--fid-tint)}
.fid-custom-summary{display:flex;flex-wrap:wrap;gap:7px;margin-top:12px}
.fid-custom-tile{display:inline-flex;align-items:baseline;gap:8px;max-width:100%;min-height:32px;padding:5px 9px;border:1px solid #d2dcee;border-radius:6px;background:#fff;color:inherit;font:inherit;text-align:left;cursor:pointer}
.fid-custom-tile:hover{border-color:var(--fid-ink);background:var(--fid-tint)}.fid-custom-tile span{overflow-wrap:anywhere}
.fid-custom-tile span:first-child{font-weight:600}.fid-custom-tile span:last-child{min-width:1em;white-space:pre-wrap}
.fid-custom-tile[data-type=boolean] span:last-child{color:#23744b;font-weight:600}
.fc-note-dialog[data-custom-type=number]{width:min(330px,calc(100vw - 16px))}
.fc-note-dialog[data-custom-type=boolean]{width:min(280px,calc(100vw - 16px));padding:11px 12px;background:#fff;overflow-wrap:anywhere}
.fc-note-dialog[data-custom-type=boolean]::backdrop{background:#152e5014}
.fcid-app .fc-note-dialog[data-custom-type=boolean] button[type=submit]{min-width:0;white-space:normal;overflow-wrap:anywhere}
.date-picks{margin:0 0 6px}.date-picks:empty{display:none}.fc-filter .fc-range{grid-template-columns:minmax(0,1fr) auto minmax(0,1fr);align-items:center;gap:7px;margin-top:0}.fc-filter .fc-range input{min-width:0;min-height:32px;padding:5px 6px;margin-top:0;background:#fff;color:var(--filter-ink,#263e60);border-color:var(--filter-line,#c2d0e4);font:inherit}
.summary-heading{display:flex;align-items:center;gap:10px}.summary-controls{display:grid;grid-template-columns:170px minmax(0,1fr);gap:14px;padding:12px;margin-bottom:6px;background:#f5f7fb;border-radius:9px}.summary-controls .fc-filter{margin:0}.summary-chart{overflow:auto}.summary-table{width:100%;table-layout:fixed;border-collapse:separate;border-spacing:0 10px;color:#344c6b}.summary-table th{font-size:12px;font-weight:600}.summary-table .fc-link{color:inherit;font:inherit;text-align:left;white-space:normal;text-decoration:none}.summary-table .fc-link:hover{text-decoration:underline}.summary-table thead th{padding:4px 10px 2px;color:#5b6d85;overflow-wrap:anywhere}.summary-table th:first-child{position:sticky;left:0;z-index:1;text-align:left}.summary-table thead th:first-child{background:#fff;z-index:2}.summary-table tbody tr{--metric-color:var(--fid-ink);--metric-bg:var(--fid-tint)}.summary-table tr[data-summary-source=nsr],.summary-table tr[data-summary-source=linked]{--metric-color:var(--nsr-ink);--metric-bg:var(--nsr-tint)}.summary-table tr[data-summary-source=ctl]{--metric-color:var(--ctl-ink);--metric-bg:var(--ctl-tint)}.summary-table tr:is([data-summary-metric=slotted],[data-summary-metric=stableSlot],[data-summary-metric=stablePost]){--metric-color:#3d805d;--metric-bg:#edf7ef}.summary-table tbody th,.summary-table td{background:var(--metric-bg);border-block:1px solid #e0e6ef}.summary-table tbody th{border-left:3px solid var(--metric-color);border-radius:9px 0 0 9px;padding:16px;color:#304967;font-size:13px;line-height:1.5}.summary-table td{padding:0}.summary-table td:last-child{border-right:1px solid #e0e6ef;border-radius:0 9px 9px 0}.summary-plot{position:relative;height:112px;font-variant-numeric:tabular-nums;text-align:center}.summary-bar{position:absolute;bottom:14px;left:calc(50% - 17px);width:34px;height:var(--bar);background:var(--metric-color);border-radius:4px 4px 0 0}.summary-value{position:absolute;top:10px;left:0;width:100%;font-size:14px;font-weight:600;color:#304967;white-space:nowrap}.summary-plot.unavailable .summary-value{color:#8794a5;font-weight:400}.summary-empty{padding:22px 12px;color:#61738c}.summary-legend{margin-top:6px;font-size:12px;color:#61738c}@media(max-width:760px){.summary-controls{grid-template-columns:minmax(0,1fr)}.summary-heading{flex-wrap:wrap}.summary-table tbody th{padding:12px}}
@container(max-width:700px){.nsr-stage-sum::before{display:none}.nsr-flow{grid-template-columns:16px minmax(0,1fr) 16px minmax(0,1fr)}.nsr-equation>[data-nsr-group=all]{grid-column:1/-1;grid-row:1;display:grid;grid-template-columns:1fr auto;align-items:center}.nsr-equation>[data-nsr-group=all] b{grid-column:2;grid-row:1/3;margin:0}.nsr-equation>.math-sign:nth-child(2){grid-column:1;grid-row:2}.nsr-equation>[data-nsr-group=completed]{grid-column:2;grid-row:2}.nsr-equation>.math-sign:nth-child(4){grid-column:3;grid-row:2}.nsr-equation>[data-nsr-group=workflow]{grid-column:4;grid-row:2}.nsr-equation>.math-sign:nth-child(6){grid-column:1;grid-row:3}.nsr-equation>[data-nsr-group=rejected]{grid-column:2;grid-row:3}.nsr-equation>.math-sign:nth-child(8){grid-column:3;grid-row:3}.nsr-equation>[data-nsr-group=cancelled]{grid-column:4;grid-row:3}.has-completed .nsr-equation>[data-nsr-group=completed]{grid-column:2/-1;display:grid;grid-template-columns:1fr auto;align-items:center}.has-completed .nsr-equation>[data-nsr-group=completed] b{grid-column:2;grid-row:1/3;margin:0}.nsr-completed-arrow{grid-column:2;grid-row:3;margin:0;height:22px}.nsr-completed-arrow::before{top:-7px}.nsr-breakdown-equals{grid-column:3;grid-row:3;margin:0}.nsr-stage-sum{grid-column:1/-1;grid-row:4;display:grid;grid-template-columns:16px minmax(0,1fr) 16px minmax(0,1fr);margin:0}.nsr-stage-spacer{display:block}.nsr-stage:only-of-type{grid-column:2/-1}.has-completed .nsr-equation>.math-sign:nth-child(4){grid-column:1;grid-row:5}.has-completed .nsr-equation>[data-nsr-group=workflow]{grid-column:2;grid-row:5}.has-completed .nsr-equation>.math-sign:nth-child(6){grid-column:3;grid-row:5}.has-completed .nsr-equation>[data-nsr-group=rejected]{grid-column:4;grid-row:5}.has-completed .nsr-equation>.math-sign:nth-child(8){grid-column:1;grid-row:6}.has-completed .nsr-equation>[data-nsr-group=cancelled]{grid-column:2;grid-row:6}}

@media(max-width:1250px){body.fcid-app{--side:360px}.nsr-pivots{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:1000px){.marked-settings{grid-template-columns:minmax(0,1fr)}.ag-grid,.fid-dialog-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:760px){.fcid-app .nsr-pivots,.ag-grid,.fid-evidence{grid-template-columns:minmax(0,1fr)}.scope-note{margin-left:0}.fid-dialog>.section-head{flex-wrap:nowrap}.fid-header-status{flex-basis:100%}}
.fid-ctl-list{list-style:none;padding:0;margin:8px 0 0;overflow-x:auto}.fid-ctl-list>li{white-space:nowrap;line-height:1.8;color:#345073}.fid-ctl-list>.is-late{color:#992f48;font-weight:600}
button.fid-quarter-state{width:100%;padding:2px 4px;border:0;border-radius:4px;background:transparent;color:inherit;font:inherit;text-align:left;cursor:pointer}button.fid-quarter-state:hover{background:#f1f5fb}button.fid-quarter-state[data-selected-quarter=true]{background:#e8f0fb}button.fid-quarter-state:focus-visible{outline:2px solid #648cc2;outline-offset:1px}
.fid-quarter-state{display:flex;align-items:center;flex-wrap:wrap;gap:4px 8px;margin-top:4px;min-width:0}.fid-quarter-state>span:first-child{font-weight:400}.fid-quarter-state.is-current>span:first-child{font-weight:600}.fid-date-value{display:grid;grid-template-rows:subgrid;grid-row:span var(--fid-date-rows);row-gap:4px}.fid-date-value>small,.fid-date-value>.fid-quarter-state{margin:0}
`;
const sourceHtml=Object.entries(sources).map(([key,source],index)=>{
  const importLabel=key==='ctlNsr'?'Append':'Import';
  const help=BW_SOURCES[key]?`<strong>Auto:</strong> Connect to VPN, <a href="${bwData.loginUrl}" target="_blank" rel="noopener noreferrer">sign in to BW</a>, and click Load.<br><strong>Manual:</strong> <a href="https://synalocey.github.io/ToolkitKyra/${BW_SOURCES[key].file}" download="${BW_SOURCES[key].file}" target="_blank" rel="noopener noreferrer">Download workbook</a>, refresh in Excel, then Import.`:
    key==='ctlNsr'?'Download from <a href="https://pcazcrpap02.fremont.lamrc.net:8443/sap(bD1lbiZjPTI1MCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm" target="_blank" rel="noopener noreferrer">the CTL system</a>.':key==='nce'?'Ask Stacey for the latest NCE / FA Review file.':'';
  return `<section class="fc-source" data-source="${key}" aria-labelledby="fcSourceTitle-${key}"><div class="fc-source-head"><h3 id="fcSourceTitle-${key}"><span class="fc-source-number">${index+1}</span><span class="fc-source-name">${escapeHtml(source.label)}</span></h3><span id="fcSource-${key}" class="fc-source-state" role="status">Unavailable</span></div><div class="fc-source-toolbar"><span class="fc-source-file" id="fcMeta-${key}"></span>${BW_SOURCES[key]?`<button class="btn small light" data-source-action="${key==='fcid'?'bw-update':'ctl-bw-update'}">Load</button>`:key==='nsr'?'<button class="btn small light" data-source-action="nsr-open">Load</button>':''}<button id="fcImport-${key}" class="btn small light" data-import-source="${key}">${importLabel}</button>${help?sourceMenu('fcHelp-'+key,'Help for '+escapeHtml(source.label),help,Boolean(BW_SOURCES[key])):''}</div><input type="file" id="fcFile-${key}" data-source-file="${key}" aria-label="${importLabel} ${escapeHtml(source.label)}" accept=".xlsx,.xls,.xlsm,.csv${key==='nsr'?',.json':''}" hidden></section>`;
}).join('');
const filterHtml=([name,label,path,needs])=>dateFilters.includes(path)?`<div class="fc-filter primary-filter" data-filter-needs="${needs}"><div class="fc-filter-head"><label for="${name.slice(0,-4)}From">${label}</label><div class="fc-filter-actions"><button type="button" data-clear-filter="${name}">Clear</button></div></div><div id="${name.slice(0,-4)}${needs==='ctl'?'Periods':'Years'}" class="nsr-year-picks date-picks"></div><div id="${name}" class="fc-range" role="group" aria-label="${label}"><input id="${name.slice(0,-4)}From" type="date" aria-label="${label} from"><span aria-hidden="true">→</span><input id="${name.slice(0,-4)}To" type="date" aria-label="${label} to"></div></div>`:`<div class="fc-filter${['fAccount','fCurrentQuarter','fQuarter','fOwner','fSlotInitiator','fNsrInitiator','fCtlStage','fCtlInitiator'].includes(name)?' primary-filter':''}" data-filter-needs="${needs}"><div class="fc-filter-head"><label for="${name}" title="${path==='owner'?escapeHtml(ownerHint):path.startsWith('nsr.')?'All selected NSR conditions must match the same NSR':''}">${label}${path==='currentQuarter'?helpIcon(currentQuarterHint):path==='quarter'?helpIcon(quarterHint):path==='owner'?helpIcon(ownerHint):path==='salesAdmin'?helpIcon(salesAdminHint):''}</label><div class="fc-filter-actions">${name==='fAccount'?'<button type="button" data-mapping-rule title="Fab / Account Mapping">Rule</button>':''}${filterResets[path]?`<button type="button" ${filterResets[path].attribute} title="${filterResets[path].title}">Reset</button>`:''}<button type="button" data-select-all="${name}">All</button><button type="button" data-clear-filter="${name}">Clear</button></div></div><select id="${name}" multiple aria-label="${label}"></select></div>`;
const views=[['portfolio','FID Portfolio'],['changes','Slot / CTL'],['nsr','NSR List'],['cycle','NSR Cycle Time'],['summary','Summary Metrics'],['rules','Rules & Settings']];

setWorkspaceHead(FID_DASHBOARD_TITLE.replace(/^🏭\s*/,''),'🏭',css);document.body.className='fcid-app';
const sideToggle=`<button id="fcToggleSide" class="btn light" aria-label="Toggle Data Sources and Filters" aria-expanded="true" title="Data Sources / Filters">☰</button>`;
document.body.innerHTML=`${workspaceHeader(escapeHtml(FID_DASHBOARD_TITLE).replace(' Tracking','<span class="title-tracking"> Tracking</span>'),'fcRefresh','disabled',sideToggle)}
<div class="layout" id="workspace"><aside class="side card"><div class="section-head"><h2>Data Sources</h2><span class="badge" id="fcSourceCount">0 / 5</span></div>${sourceHtml}<div id="fcActionStatus" class="fc-note" role="status"></div>
<div class="filter-head"><h2>Filters</h2><div class="filter-presets" role="group" aria-label="Filter presets">${filterPresetSlots.map((slot,index)=>{
  const saved=Boolean(savedFilterPresets[index]),name=filterPresetName(slot);
  return `<button type="button" data-filter-preset="${slot}" data-saved="${saved}" aria-label="${name}, ${saved?'saved':'empty'}" title="${name}" aria-haspopup="dialog" aria-controls="filterPresetPopup">${slot}</button>`;
}).join('')}</div><button class="btn small light" id="fcReset" title="Restore built-in filters">Reset All</button></div>
<form id="fidLookup" class="fc-filter fid-lookup"><textarea id="fSearch" class="quick-search" placeholder="FID / CTL / NSR / Quote ID&#10;Single match → FID 360°" aria-label="FID / CTL / NSR / Quote ID&#10;Single match → FID 360°" rows="2"></textarea><button class="btn light" type="submit" aria-label="Open FID 360°" title="Open FID 360°" hidden><span aria-hidden="true">🔍</span></button></form>
${['fAccount'].map(name=>filterHtml(filterDefs.find(def=>def[0]===name))).join('')}<h3 class="filter-group-title" data-filter-needs="fid">FID Filters</h3>${['fQuarter','fCurrentQuarter','fOwner','fSlotInitiator','fCommit','fFidStatus','fRisk','fCustomer','fSystem','fBusinessUnit','fMfg','fSalesAdmin'].map(name=>filterHtml(filterDefs.find(def=>def[0]===name))).join('')}<h3 class="filter-group-title" data-filter-needs="ctl">CTL Filters</h3>${['fCtlStage','fCtlInitiator','fCtlCreatedDate','fCtlChange','fCtlCustomer'].map(name=>filterHtml(filterDefs.find(def=>def[0]===name))).join('')}<h3 class="filter-group-title" data-filter-needs="nsr">NSR Filters</h3>${['fNsrInitiator','fNsrSubmitDate','fNsrTaskStatus','fNsrType','fNsrOrg','fNsrSubsystem','fNsrProduct','fNsrCustomer'].map(name=>filterHtml(filterDefs.find(def=>def[0]===name))).join('')}</aside>
<main class="main"><nav class="workspace-nav" id="workspaceNav" aria-label="Dashboard views">${views.map(([view,label])=>`<button type="button" data-view="${view}" aria-selected="${view==='nsr'}"${view==='nsr'?' class="active"':''}${['nsr','cycle','rules'].includes(view)?'':' hidden'}>${label}</button>`).join('')}</nav><div id="filterContext" class="filter-context"></div>${views.map(([view])=>`<section id="view-${view}" class="fc-view"${view==='nsr'?'':' hidden'}></section>`).join('')}</main></div>
<div id="fidDatePopup" class="card col-popup fid-date-popup" popover role="dialog" aria-labelledby="fidDatePopupTitle" tabindex="-1" autofocus></div>
<dialog id="fcSettingsImportDialog" class="fc-nsr-dialog" aria-labelledby="fcSettingsImportTitle" aria-describedby="fcSettingsImportHelp">
  <form method="dialog"><h2 id="fcSettingsImportTitle">Import Settings</h2><p id="fcSettingsImportFileName" class="muted" style="overflow-wrap:anywhere"></p>
    <div id="fcSettingsImportHelp"><p><strong>Append</strong> adds missing entries and keeps your saved values.</p><p><strong>Replace</strong> overwrites the saved settings covered by this file.</p></div>
    <div class="fc-actions"><button class="btn light" type="submit" value="cancel" autofocus>Cancel</button><button class="btn blue" type="submit" value="append">Append</button><button class="btn light" type="submit" value="replace">Replace</button></div>
  </form>
</dialog>
<dialog id="fidCustomDeleteDialog" class="fc-nsr-dialog" aria-labelledby="fidCustomDeleteTitle" aria-describedby="fidCustomDeleteHelp">
  <form method="dialog"><h2 id="fidCustomDeleteTitle"></h2><p id="fidCustomDeleteHelp" class="muted"></p>
    <div class="fc-actions"><button class="btn light" type="submit" value="cancel" autofocus>Cancel</button><button class="btn delete" type="submit" value="delete">Delete</button></div>
  </form>
</dialog>
<div id="filterPresetPopup" class="card col-popup filter-preset-popup" popover role="dialog" aria-labelledby="filterPresetQuestion"><p id="filterPresetQuestion"></p><div class="btnrow"><button type="button" class="btn small blue" data-preset-answer="yes">Yes</button><button type="button" class="btn small light" data-preset-answer="no">No</button></div></div>
<dialog id="fcNsrDialog" class="fc-nsr-dialog" aria-labelledby="fcNsrDialogTitle"><form id="fcNsrForm"><h2 id="fcNsrDialogTitle">Load NSR</h2><div class="fc-range"><label>From<input id="fcNsrFrom" type="date" required></label><label>To<input id="fcNsrTo" type="date" required></label></div><div class="fc-actions"><button id="fcNsrCancel" class="btn light" type="button">Cancel</button><button id="fcNsrLoad" class="btn blue" type="submit">Load</button></div></form></dialog><dialog id="fidDetail" class="fid-dialog" aria-labelledby="fidDetailTitle"></dialog><dialog id="fidNoteDialog" class="fc-note-dialog" aria-labelledby="fidNoteTitle"></dialog><form id="columnFilterPopup" class="card col-popup" popover role="dialog" aria-label="Column filter"><label><span></span><input autocomplete="off" autofocus></label><select id="colNoteFilter" aria-label="Note presence" hidden><option value="">All</option><option value="has">Has note</option><option value="none">No note</option></select><select id="colChoices" size="6" aria-label="Column values" hidden></select><div class="btnrow"><button type="reset" class="fc-link">Clear</button><button id="colPendingCtl" type="button" class="fc-link" hidden>${pendingCtlMarker} Pending only</button><button type="submit" class="btn blue small">Apply</button></div></form><div id="markFilterPopup" class="card col-popup mark-popup" popover role="dialog" aria-label="Mark filter"><div role="group" aria-label="Mark status"></div></div><div id="fcToast" class="fc-toast" role="status" hidden></div>`;
const cycleTemplate=document.createElement('template');cycleTemplate.innerHTML=NSR_CYCLE_HTML;id('view-cycle').append(cycleTemplate.content.querySelector('#cycleTab'));id('workspace').append(cycleTemplate.content.querySelector('#wfModal'));
let selectedView='nsr',graph=null,current=null,revision=0,modelRevision=0,buildRun=0,dataRevision=0,toastTimer=0,searchTimer=0,sidebarAnimation=null;
let bwLookup=null;
let linkedId='';
const rendered=new Map(),tables=new Map(),filterValues={},localTables=new Map();let expandedTable=null,mappingDraft=null,portfolioCache={};const markDrafts={nsr:null,fid:null};
const workNotes=mountWorkNotes(syncFidNotes);
if(typeof GM_addValueChangeListener==='function')GM_addValueChangeListener(FID_NOTES_KEY,(_key,_old,_value,remote)=>{if(remote)syncFidNotes()});
const pendingMarks={nsr:new Set(),fid:new Set()};
const filterPresetPopup=id('filterPresetPopup');
let pendingFilterValues={commit:defaultCommit};
const pretty=(value,path='')=>path==='nsr.taskStatus'?(completedTasks[value]||value||'[Unknown]'):path==='ctl.stage'?ctlStageName(value)||'[Unknown]':path==='ctl.change'&&value==='[Unknown]'?'No CTL Yet':path==='fidStatus'?(value==='stable-slot|stable-ctl'?fidStatusNames['risk-free']:value.startsWith('cause:')?fidCauseNames[value.slice(6)]||value:fidCellNames[value]||fidStatusNames[value]||value):path==='riskType'?(value==='MCSD pending'?'Slot + no MCSD':value):value||'[Unknown]';
const number=value=>String(Number(value||0));
const pill=(value,kind='info',tip='')=>`<span class="pill ${kind}"${tip?` title="${escapeHtml(tip)}"`:''}>${escapeHtml(value||'—')}</span>`;
const fidLink=(fid,target,label=fid)=>fid?`<button type="button" class="fc-link" data-fid="${escapeHtml(fid)}"${target?.ctlId?` data-ctl-id="${escapeHtml(target.ctlId)}"`:target?.nsrId?` data-nsr-id="${escapeHtml(target.nsrId)}"`:''}>${escapeHtml(label)}</button>`:'—';
const nsrWorkflowLink=number=>number?`<button type="button" class="fc-link" data-wf-nsr="${escapeHtml(number)}" title="View full workflow for ${escapeHtml(number)}">${escapeHtml(number)}</button>`:'';
const linkSuffix=(type,fid='')=>type==='Chg'?' (Chg'+(fid?' · '+fid:'')+')':'';
const quarterYear=value=>nsrRecords.periodYear(value)||String(value).match(/'(\d{2})$/)?.[1]?.replace(/^(\d{2})$/,'20$1')||'';
function defaultQuarters(){
  const current=fidQuarterFromDate(new Date()),first=Number(current.slice(-2))*4+['Mar','Jun','Sep','Dec'].indexOf(current.slice(1,4));
  return Array.from({length:5},(_,offset)=>{const quarter=first+offset;return ['QMar','QJun','QSep','QDec'][quarter%4]+"'"+String(Math.floor(quarter/4)).slice(-2).padStart(2,'0')});
}
function toast(message){clearTimeout(toastTimer);id('fcToast').innerHTML=`<span>${escapeHtml(message)}</span><button type="button" data-toast-close aria-label="Dismiss">×</button>`;id('fcToast').hidden=false;toastTimer=setTimeout(()=>id('fcToast').hidden=true,4500)}
function captureFilterPreset(){
  const filters={};
  for(const [name,,path] of filterDefs)if(!id(name).closest('.fc-filter').hidden)filters[path]=[...(filterValues[path]||[])];
  if(filters.quarter?.length===5&&defaultQuarters().every(quarter=>filters.quarter.includes(quarter)))filters.quarter='rolling';
  if(store.raw.fcid.length&&filterValues.signal?.length)filters.signal=[...filterValues.signal];
  return {filters,search:id('fSearch').value,includeIgnored};
}
function applyFilterPreset(preset){
  const previous=graph&&JSON.stringify(queryFilters());
  pendingFilterValues={quarter:[],'ctl.change':[],commit:defaultCommit};
  for(const path of filterPresetPaths){
    const values=preset.filters[path];
    if(path==='quarter'&&values==='rolling'){pendingFilterValues.quarter=values;continue}
    if(!Array.isArray(values)||!values.every(value=>typeof value==='string'))continue;
    if(dateFilters.includes(path)&&(values.length>2||values.some(value=>value&&!nsrData.validSapDate(value))||values[0]&&values[1]&&values[0]>values[1]))continue;
    pendingFilterValues[path]=values;
  }
  for(const path of Object.keys(filterValues))delete filterValues[path];
  filterValues.signal=pendingFilterValues.signal||[];delete pendingFilterValues.signal;
  includeIgnored=preset.includeIgnored;linkedId='';id('fSearch').value=preset.search;
  if(graph){
    buildFilters();syncFilters();
    if(previous!==JSON.stringify(queryFilters())||hasPendingMarks())applyFilters(true);
  }
}
function paintFilterPreset(button,preset){
  const saved=Boolean(preset),name=filterPresetName(button.dataset.filterPreset);
  button.dataset.saved=String(saved);
  button.setAttribute('aria-label',name+(saved?', saved':', empty'));
}
function openFilterPreset(button,clear=false){
  try{
    const slot=button.dataset.filterPreset,stored=GM_getValue(FILTER_PRESET_KEY+slot,null);
    const preset=parseFilterPreset(stored),name=filterPresetName(slot);
    paintFilterPreset(button,preset);
    if(filterPresetPopup.matches(':popover-open'))filterPresetPopup.hidePopover();
    if(clear&&!preset)return;
    const action=clear?'clear':preset?'apply':'save';
    id('filterPresetQuestion').textContent=clear?`Clear ${name}?`:preset?`Apply ${name}?`:`Save filters as ${name}?`;
    filterPresetPopup.selection={button,slot,action,stored,snapshot:action==='save'?captureFilterPreset():preset};
    button.focus({preventScroll:true});
    filterPresetPopup.showPopover({source:button});
    filterPresetPopup.querySelector('[data-preset-answer="no"]').focus({preventScroll:true});
  }catch{toast('Could not open the filter preset.')}
}
function confirmFilterPreset(answer){
  const selection=filterPresetPopup.selection;if(!selection)return;
  filterPresetPopup.hidePopover();
  if(answer!=='yes')return;
  const {button,slot,action,stored,snapshot}=selection,key=FILTER_PRESET_KEY+slot;
  try{
    const current=GM_getValue(key,null);
    if(current!==stored){
      paintFilterPreset(button,parseFilterPreset(current));toast('This preset changed. Open it again.');return;
    }
    if(action==='apply'){applyFilterPreset(snapshot);return}
    const next=action==='clear'?null:snapshot;
    GM_setValue(key,next?JSON.stringify(next):null);
    paintFilterPreset(button,next);
  }catch{toast('Could not '+action+' the filter preset.')}
}
function buildFilters(paths){
  const nodes=graph.nativeFids,statNsrs=graph.nsrs.filter(nsr=>nsr.eligible),statCtls=graph.ctls;let observations;
  const fidLoaded=store.raw.fcid.length>0;
  const assessments=nodes.flatMap(node=>node.quarterStates);
  const slotUnknown=assessments.some(node=>node.portfolioCell.startsWith('slot-date-unknown|')),ctlUnknown=assessments.some(node=>node.portfolioCell.endsWith('|ctl-unknown'));
  const cells=Object.keys(fidCellNames).filter(key=>(!key.startsWith('slot-date-unknown|')||slotUnknown)&&(!key.endsWith('|ctl-unknown')||ctlUnknown));
  const normalCells=['stable-slot|stable-ctl','stable-slot|late-ctl-nsr','late-slot|late-ctl-nsr','late-slot|stable-ctl'];
  const options=[['no-slot',0],['cause:slot-pending',1],['slot-ready',0],['cause:slot-dropped',1],
    ...normalCells.map(key=>[key,1]),
    ...cells.filter(key=>!normalCells.includes(key)).map(key=>[key,1]),
    ...(nodes.some(node=>node.portfolioCell==='slot-unknown')?[['slot-unknown',0]]:[])];
  fidSelection.order=options.map(([value])=>value).filter(value=>value!=='slot-ready');
  fidSelection.groups={'slot-ready':cells,
    ...Object.fromEntries(['stable-slot','late-slot','slot-date-unknown','stable-ctl','late-ctl-nsr','ctl-unknown'].map(key=>[key,cells.filter(cell=>cell.split('|').includes(key))])),
    'risk-fid':cells.filter(key=>fidRiskAssessment(key)==='risk-fid'),
    'risk-free':['stable-slot|stable-ctl'],
    'risk-unknown':cells.filter(key=>fidRiskAssessment(key)==='risk-unknown')};
  const fidDepth=new Map(options),fidStatuses=options.map(([value])=>value);
  const riskOrder=['MCSD pending','NCE / FA signal',dateChangeDriver];
  for(const [name,,path] of filterDefs){
    if(paths&&!paths.includes(path))continue;
    if(path==='commit'&&!fidLoaded)continue;
    let requested=pendingFilterValues[path];
    if(dateFilters.includes(path)){
      if(!(path==='nsr.submitted'?statNsrs:statCtls).length)delete filterValues[path];
      else if(requested){filterValues[path]=requested;delete pendingFilterValues[path]}
      const [from='',to='']=filterValues[path]||[],[start,end]=id(name).querySelectorAll('input');start.value=from;end.value=to;start.max=to;end.min=from;start.setCustomValidity('');end.setCustomValidity('');
      id(name).previousElementSibling.innerHTML=path==='ctl.created'?['This week','Last week','This month','Last month'].map(period=>{
        const range=ctlDateRange(period),active=from===range[0]&&to===range[1];return `<button type="button" data-ctl-period="${period}" class="${active?'active':''}" aria-pressed="${active}" title="${range.join(' → ')}">${period}</button>`;
      }).join(''):[...document.querySelectorAll('[data-fid-year]')].map(button=>{
        const year=button.dataset.fidYear,active=from===year+'-01-01'&&to===year+'-12-31';
        return `<button type="button" data-nsr-year="${year}" class="${active?'active':''}" aria-pressed="${active}" title="Select January 1 through December 31, ${year}">${year}</button>`;
      }).join('');continue;
    }
    const values=path==='fidStatus'?[]:path==='commit'?nodes.map(node=>node.commit.trim().replace(/^commit$/i,'Commit')):path==='quarter'?assessments.map(state=>state.quarter):path==='account'?[...nodes.map(node=>node.account),...statNsrs.flatMap(nsr=>nsr.accounts.length?nsr.accounts:[''])]:path==='riskType'?nodes.flatMap(node=>node.risks.map(risk=>risk.label)).filter(value=>riskOrder.includes(value)):path.startsWith('ctl.')?statCtls.map(ctl=>ctl[path.slice(4)]):path.startsWith('nsr.')?(observations??=statNsrs.flatMap(nsr=>nsr.observations.length?nsr.observations:[nsr])).map(observation=>observation[path.slice(4)]):path==='slot.initiator'?nodes.map(slotInitiator):nodes.map(node=>node[path]);
    if(path==='riskType'&&fidDateChanges.size)values.push(dateChangeDriver);
    const quarterField=path==='quarter'||path==='currentQuarter';
    const extra=path==='quarter'&&fidLoaded?defaultQuarters():path==='commit'?defaultCommit:[];
    if(requested==='rolling')requested=extra;
    else if(quarterField)extra.push(...(requested&&fidLoaded?requested:filterValues[path]||[]));
    else if(path==='ctl.change'&&(store.meta.ctl||store.meta.ctlNsr))extra.push(...defaultTypes,...Object.values(pendingTypes).flat());
    const available=new Set(path==='fidStatus'?fidStatuses:[...values.map(value=>value||'[Unknown]'),...extra]);
    const ordered=path==='fidStatus'?fidStatuses:[...available].sort((a,b)=>Number(a==='[Unknown]')-Number(b==='[Unknown]')||(path==='nsr.taskStatus'?nsrRecords.compareStatus(a,b):path==='ctl.stage'?compareCtlStages(a,b):path==='riskType'?riskOrder.indexOf(a)-riskOrder.indexOf(b):quarterField?nsrRecords.periodRank(a)-nsrRecords.periodRank(b)||a.localeCompare(b):numericCollator.compare(a,b)));
    if(path==='nsr.taskStatus'&&available.has('15-Completed')){
      const completed=Object.keys(completedTasks);
      ordered.splice(ordered.indexOf('15-Completed')+1,0,...completed);
      for(const value of completed)available.add(value);
    }
    const select=id(name),desired=requested||filterValues[path]||[],saved=path==='fidStatus'?normalizeFidStatus(desired):desired;
    const selected=saved.filter(value=>available.has(value)),chosen=selected.length?new Set(selected):null;
    const option=value=>{
      const depth=fidDepth.get(value);
      const classes=path==='fidStatus'
        ? [depth===1?'fid-status-child':'',['no-slot','slot-ready'].includes(value)?'fid-status-parent':''].filter(Boolean).join(' ')
        : path==='nsr.taskStatus'&&Object.hasOwn(completedTasks,value)?'group-option':'';
      const label=path==='fidStatus'&&['cause:slot-pending','cause:slot-dropped'].includes(value)?pretty(value,path).replace(/^(?:No Slot|Slotted FID) · /,''):pretty(value,path);
      return `<option value="${escapeHtml(value)}"${classes?` class="${classes}"`:''}${chosen?.has(value)?' selected':''}>${escapeHtml(label)}</option>`;
    };
    select.innerHTML=ordered.map(option).join('');
    filterValues[path]=selected;jQuery(select).trigger('chosen:updated');
    // Keep the requested values until the relevant source and choices are available.
    if(requested&&available.size&&(!quarterField||fidLoaded)&&requested.every(value=>chosen?.has(value)))delete pendingFilterValues[path];
    if(name==='fQuarter'){
      const currentYear=new Date().getFullYear();let years=select.parentElement.querySelector('.nsr-year-picks');
      if(!years){years=document.createElement('div');years.className='nsr-year-picks';select.after(years)}
      years.innerHTML=[...new Set(values.map(quarterYear).filter(year=>year>=currentYear-1&&year<=currentYear+3))].sort().map(year=>`<button type="button" data-fid-year="${year}" title="Select / clear all quarters in ${year}">${year}</button>`).join('');
    }
  }
}
function syncFilters(){
  const fid=store.raw.fcid.length>0,nsr=store.raw.nsr.length>0,ctl=store.raw.ctl.length>0||store.raw.ctlNsr.length>0;
  const nodes=graph.nativeFids,needs={common:fid||nsr,fid,fidCtl:fid&&ctl,fidOwner:fid&&ctl&&nodes.some(node=>node.owner),fidSlot:fid&&ctl&&nodes.some(slotInitiator),fidSalesAdmin:fid&&ctl&&nodes.some(node=>node.salesAdmin),fidRisk:fid&&id('fRisk').options.length>0,ctl,nsr};document.querySelectorAll('[data-filter-needs]').forEach(group=>group.hidden=!needs[group.dataset.filterNeeds]);
}
function bwLookupIdentifiers(value){
  const identifiers=new Map();
  for(const [,label,number] of fidLookupMatches(value)){
    const kind=/^\d{6}$/.test(number)?'FID':/^\d{10}$/.test(number)?'CTL':'';
    if(!kind||label&&label.toUpperCase()!==kind)continue;
    if(label||!identifiers.has(number))identifiers.set(number,label?kind:'');
  }
  return identifiers;
}
function syncFidLookup(){
  const value=id('fSearch').value,button=id('fidLookup').querySelector('button');
  if(bwLookup&&bwLookup.value!==value)clearBwLookup();
  const identifiers=bwLookupIdentifiers(value),number=identifiers.keys().next().value,labelled=identifiers.get(number);
  let matches=graph?.lookupFids(value,true)||[];
  if(labelled)matches=matches.filter(match=>labelled==='FID'?match.fid===number:match.target?.ctlId===number);
  const match=matches.length===1&&(!identifiers.size||graph.fids.get(matches[0].fid)?.hasNativeRecord)?matches[0]:null;
  const searchBw=identifiers.size>0&&(!match||Boolean(bwLookup)),loading=Boolean(bwLookup?.controller),checked=bwLookup?.checked;
  const label=searchBw?(loading?'Searching BW':checked?'BW lookup complete':'Search in BW'):'Open FID 360°';
  button.hidden=!match&&!searchBw;
  button.disabled=Boolean(searchBw&&(loading||store.refreshing||store.busy.has('fcid')||store.busy.has('ctl')));
  button.dataset.state=searchBw?(loading?'loading':checked?'done':'idle'):'local';
  button.setAttribute('aria-label',label);button.setAttribute('aria-busy',String(loading));
  button.title=searchBw&&bwLookup?.error?label+'\n'+bwLookup.error:label;
  button.firstElementChild.textContent=searchBw?(checked?'✓':'↻'):'🔍';
  return searchBw?null:match;
}
function clearBwLookup(){
  const previous=bwLookup;bwLookup=null;previous?.controller?.abort();
}
async function readBwLookupRows(key,field,identifiers,run){
  const {signal}=run.controller,values=[...identifiers],rows=[];
  for(let offset=0;offset<values.length&&!signal.aborted;offset+=500){
    try{
      const result=await bwData.read(key,BW_SOURCES[key].query,signal,undefined,{field,values:values.slice(offset,offset+500),details:true});
      for(const row of result.rows)rows.push(row);
      run.received=true;
    }catch(error){if(!signal.aborted)run.error??=error.message}
  }
  return rows;
}
async function searchBwDetails(identifiers){
  const controller=new unsafeWindow.AbortController(),run={value:id('fSearch').value,controller},view=selectedView;
  bwLookup=run;syncFidLookup();
  const ctls=new Set(identifiers.filter(value=>value.length===10));
  const fids=new Set(identifiers.filter(value=>value.length===6));
  try{
    const directCtls=await readBwLookupRows('ctl','CTL ID',ctls,run);
    if(controller.signal.aborted)return;
    for(const ctl of graph.ctls)if(ctls.has(ctl.number)&&/^\d{6}$/.test(ctl.fidId))fids.add(ctl.fidId);
    for(const row of directCtls){const fid=text(row['Opportunity ID']);if(/^\d{6}$/.test(fid))fids.add(fid)}
    const [fidRows,history]=await Promise.all([readBwLookupRows('fcid','FID',fids,run),readBwLookupRows('ctl','Opportunity ID',fids,run)]);
    if(controller.signal.aborted)return;
    const ctlRows=new Map([...directCtls,...history].map(row=>[text(row['CTL ID']),row]));
    const appended=await appendBwLookup({fcid:fidRows,ctl:[...ctlRows.values()]});
    run.fids=fids;
    if(appended)await recalculate(true,true);
    else if(bwLookup===run)applyFilters();
    if(bwLookup===run&&selectedView===view&&view!=='portfolio'&&view!=='changes'&&[...fids].some(fid=>graph.fids.get(fid)?.hasNativeRecord))showView('portfolio');
  }catch(error){if(!controller.signal.aborted)run.error??=error.message}
  finally{
    if(bwLookup===run){
      run.controller=null;run.checked=Boolean(run.received);syncFidLookup();
      if(run.checked)setTimeout(()=>{if(bwLookup===run){run.checked=false;syncFidLookup()}},2400);
    }
  }
}
function bwDetailRows(key,rows){
  if(!bwLookup?.fids)return rows;
  const fids=[...bwLookup.fids].map(fid=>graph.fids.get(fid)).filter(node=>node?.hasNativeRecord&&(includeIgnored||node.fidMark!=='ignore'));
  const matching=key==='fcid'?fids.map(node=>graph.projectFid(node)):fids.flatMap(node=>node.ctls);
  // Exact lookup rows remain visible even when their dates fall outside the overview filters.
  const scoped=new Map(rows.map(row=>[row.id,row]));
  return matching.map(row=>scoped.get(row.id)||row);
}
function openLinkedFid(settled=false,cachedReady=false){
  if(!linkedId)return;
  const legacy=/^#fid=/i.test(unsafeWindow.location.hash);
  if(!legacy&&store.refreshing&&!cachedReady)return;
  const matches=legacy?(graph?.fids.has(linkedId)?[{fid:linkedId}]:[]):graph?.lookupFids(linkedId,true)||[];
  if(matches.length===1){
    const [{fid,target}]=matches;linkedId='';showFid(fid,false,target);
  }else if(matches.length>1&&!legacy){
    linkedId='';
  }else if(settled){
    if(legacy)toast('FID '+linkedId+' was not found in the loaded sources.');
    else if(Object.keys(store.meta).length)linkedId='';
  }
}
function readFidLink(){
  const hash=unsafeWindow.location.hash;
  linkedId=hash.match(/^#fid=(\d{5,10})$/i)?.[1]||'';
  if(/^#filter=/i.test(hash)){
    try{id('fSearch').value=decodeURIComponent(hash.slice(8))}catch{return}
    if(id('fidDetail').open)id('fidDetail').close();
    if(graph)applyFilters();
    return;
  }
  if(/^#lookup=/i.test(hash)){
    let value;
    try{value=decodeURIComponent(hash.slice(8)).toUpperCase()}catch{return}
    const identifiers=[...new Set(value.split(/[\s,;|/，、；]+/).filter(Boolean))];
    if(!identifiers.length||identifiers.some(identifier=>!/^(?:\d{5,10}|[NDL]\d{5,8})$/.test(identifier)))return;
    linkedId=identifiers.join('\n');
    if(id('fidDetail').open)id('fidDetail').close();
    id('fSearch').value=linkedId;if(graph)applyFilters();
  }
  openLinkedFid(Boolean(graph)&&!store.refreshing&&!store.busy.size);
}
async function recalculate(force=false,quiet=false){
  const started=performance.now(),run=++buildRun,inputsRevision=dataRevision,refresh=id('fcRefresh');
  if(graph){store.busy.add('model');refresh.disabled=true;if(!quiet)refresh.textContent='Processing…'}
  try{
    const dateChanges=new Map();
    const options={fidMarks:currentMarks('fid'),accountMap:accountMapping.match,ignoredNsrs,importantNsrs,dateChangeIds:dateChanges,
      loadedSources:Object.fromEntries(Object.keys(sources).map(key=>[key,Boolean(store.meta[key])])),today:nsrData.localDate(new Date()),yield:graph?async()=>{
        if(unsafeWindow.scheduler?.postTask)await unsafeWindow.scheduler.postTask(()=>{},{priority:'background'});
        else await new Promise(resolve=>setTimeout(resolve,0));
        return run===buildRun&&inputsRevision===dataRevision;
      }:null};
    // Reuse a calculated lookup only when its complete input, rules and model code still match.
    const loadedKeys=Object.keys(store.meta),cacheReady=loadedKeys.length>0&&loadedKeys.every(key=>sourceCacheRevisions.has(store.raw[key]));
    const cacheKey=CACHE_PREFIX+'model';
    const cacheSignature=async()=>{
      const revisions=Object.keys(sources).map(key=>store.meta[key]?sourceCacheRevisions.get(store.raw[key])??store.raw[key]:null);
      const input=JSON.stringify([revisions,accountMapping.getMap(),createFidAccountMapping.toString(),createFidGraphSteps.toString(),fidQuarterFromDate.toString(),fidRiskAssessment.toString()]);
      const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(input));
      return Array.from(new Uint8Array(digest),byte=>byte.toString(16).padStart(2,'0')).join('');
    };
    let signature='',next,restored=false;
    if(cacheReady&&!force)try{
      const cached=await nsrData.readDataset(cacheKey);
      if(run!==buildRun||inputsRevision!==dataRevision)return;
      if(cached?.schema===2&&cached.snapshot?.sourceRowRefs&&nsrData.validSapDate(cached.snapshot.today))signature=await cacheSignature();
      if(run!==buildRun||inputsRevision!==dataRevision)return;
      if(signature&&cached.signature===signature){
        next=await createFidGraph(store.raw,{...options,snapshot:cached.snapshot});
        restored=Boolean(next);
      }
    }catch{/* A calculated cache is optional; the source snapshots remain authoritative. */}
    if(run!==buildRun||inputsRevision!==dataRevision)return;
    if(!next){
      next=createFidGraph(store.raw,options);
      if(options.yield)next=await next;
    }
    if(!next||run!==buildRun||inputsRevision!==dataRevision)return;
    // Marks can change while a replacement is being built; apply the latest metadata before showing it.
    for(const node of next.observedPopulation){
      next.setFidMark(node.id,recordMark('fid',node.id));
      // A lone assessment is the current quarter, with a single normalized value per field.
      if(node.quarterStates.length<2)continue;
      const dates=recordedFidDates(node);
      if(dates)dateChanges.set(node.id,dates);
    }
    for(const nsr of next.nsrs)next.setNsrMark(nsr.number,recordMark('nsr',nsr.number));
    graph=next;fidDateChanges=dateChanges;modelRevision++;portfolioCache={};
    buildFilters();syncFilters();applyFilters();openLinkedFid();store.dirty=false;store.buildMs=performance.now()-started;
    if(cacheReady&&!restored){
      const save=async()=>{
        if(run!==buildRun||inputsRevision!==dataRevision)return;
        try{
          const modelSignature=signature||await cacheSignature();
          if(run!==buildRun||inputsRevision!==dataRevision)return;
          await nsrData.writeDataset({key:cacheKey,schema:2,signature:modelSignature,snapshot:packFidSnapshot(next.getSnapshot(),store.raw)});
        }catch{}
      };
      requestAnimationFrame(()=>requestAnimationFrame(()=>{
        if(unsafeWindow.requestIdleCallback)unsafeWindow.requestIdleCallback(save,{timeout:2000});else setTimeout(save,0);
      }));
    }
    return true;
  }catch(error){if(run===buildRun)toast('Dashboard update failed: '+error.message)}
  finally{if(run===buildRun){
    store.busy.delete('model');
    if(inputsRevision===dataRevision)for(const key of Object.keys(store.meta))if(id('fcSource-'+key).classList.contains('pending'))sourceMessage(key,'Loaded','done',!id('fcHelp-'+key)?.hidden);
    refresh.disabled=store.refreshing||store.busy.size>0;refresh.textContent=store.refreshing?'Refreshing…':'Refresh';
  }}
}
function queryFilters(){
  const filters={search:id('fSearch').value,fidStatus:[],signal:store.raw.fcid.length?filterValues.signal||[]:[],includeIgnored};
  for(const [name,,path] of filterDefs)if(!id(name).closest('.fc-filter').hidden&&(path==='commit'||filterValues[path]?.length)){
    const keys=path.split('.'),last=keys.pop();let target=filters;for(const key of keys)target=target[key]??={};
    target[last]=dateFilters.includes(path)?{from:filterValues[path][0],to:filterValues[path][1]}:filterValues[path];
  }
  return filters;
}
function hasPendingMarks(){return pendingMarks.fid.size>0||pendingMarks.nsr.size>0}
function applyFilters(reuseTables=false){
  clearTimeout(searchTimer);searchTimer=0;
  if(!reuseTables&&expandedTable)toggleTableExpanded();
  closeFidDateHistory();
  if(filterPresetPopup.selection)filterPresetPopup.hidePopover();
  const lookup=syncFidLookup(),dialog=id('fidDetail');
  const detailFid=dialog.open?dialog.dataset.fid:'',anchor=id('view-'+selectedView),position=reuseTables?anchor?.getBoundingClientRect():null;
  if(!reuseTables)clearFidLocationPulses(dialog);
  if(hasPendingMarks()){
    for(const number of pendingMarks.nsr)graph.setNsrMark(number,recordMark('nsr',number));
    for(const fid of pendingMarks.fid)graph.setFidMark(fid,recordMark('fid',fid));
    pendingMarks.nsr.clear();pendingMarks.fid.clear();portfolioCache={};
  }
  current=graph.query(queryFilters());revision++;summaryCache={};quarterCache={};
  for(const [name,state] of tables){if(name==='cycleTable'||name==='wfTable'||name==='wfAdditional'||reuseTables&&(state.grid.element.closest('.fc-view')?.hidden===false||state.grid.element.closest('.fid-dialog')?.open))continue;tableResize.unobserve(state.grid.element);state.grid.destroy();tables.delete(name)}
  rendered.clear();paintContext();showView(selectedView,reuseTables);
  if(detailFid){
    // Retarget the open lookup if a source update rebuilds its history tables.
    const target=!reuseTables&&lookup?.fid===detailFid&&lookup.target?.lookupId===dialog.dataset.lookupId?lookup.target:null;
    showFid(detailFid,reuseTables,target);
  }
  const additional=tables.get('wfAdditional');if(additional?.options.fid&&additional.rows.length)table(additional.name,additionalFidColumns(),additional.rows,additional.options,true);
  // Keep visible content in place when filter chips or NSR Overview cards change height.
  if(position?.top<innerHeight&&position.bottom>0)window.scrollBy({top:anchor.getBoundingClientRect().top-position.top,behavior:'instant'});
}
function quartersForYear(year){return Array.from(id('fQuarter').options,option=>option.value).filter(value=>quarterYear(value)===year)}
function paintContext(){
  document.querySelectorAll('[data-exclude-ignored]').forEach(input=>input.checked=!includeIgnored);
  document.querySelectorAll('[data-fid-year]').forEach(button=>{const values=quartersForYear(button.dataset.fidYear),selected=filterValues.quarter||[],count=values.filter(value=>selected.includes(value)).length,active=count>0&&count===values.length;
    button.classList.toggle('active',active);button.classList.toggle('partial',count>0&&count<values.length);button.setAttribute('aria-pressed',String(active))});
  id('filterContext').hidden=!Object.keys(store.meta).length;
  const chip=(value,target,clear,label,reset)=>{
    const action=reset?`${reset.attribute} title="${reset.title}"`:clear;
    return `<span class="filter-chip" title="${escapeHtml(value)}"><button class="filter-chip-label" type="button" data-locate-filter="${target}" aria-label="Locate ${escapeHtml(label)} filter">${escapeHtml(value)}</button><button class="filter-chip-remove" type="button" ${action} aria-label="${reset?'Reset':'Remove'} ${escapeHtml(label)} filter">${reset?'↺':'×'}</button></span>`;
  };
  let defaultsApplied=false;
  const chips=filterDefs.flatMap(([name,label,path])=>{
    if(id(name).closest('.fc-filter').hidden)return [];
    const selected=filterValues[path]||[],reset=filterResets[path],defaults=reset?.values();
    if(defaults?.length===selected.length&&defaults.every(value=>selected.includes(value))){defaultsApplied=true;return []}
    if(!reset&&!selected.length)return [];
    const [from,to]=selected;
    const value=!selected.length?'All':dateFilters.includes(path)?from&&to?selected.join(' → '):from?'From '+from:'Through '+to:selected.map(value=>pretty(value,path)).join(' or ');
    return [chip(label+': '+value,name,`data-clear-filter="${name}"`,label,reset)];
  });
  if(store.raw.fcid.length&&filterValues.signal?.length)chips.push(chip('Signal: '+filterValues.signal.map(value=>fidStatusNames[value]||(value==='nce'?'NCE / FA':'Pending Config CTL')).join(' or '),'fRisk','data-clear-signal','signal'));
  const search=id('fSearch').value.trim(),identifiers=fidLookupIdentifiers(search);
  if(search)chips.push(chip('Search: '+(identifiers.join(' · ')||search),'fSearch','data-clear-search','search'));
  id('filterContext').innerHTML=`<span class="context-title">Filtered By:</span><div class="filter-tags">${chips.join('')||`<span class="muted">${defaultsApplied?'Default filters':'All values'}</span>`}</div><span class="scope-note">FID portfolio: ${number(current.fids.length)} FIDs · NSR List: ${number(current.nsrs.length)}</span>`;
}
function showView(name,reuseTables=false){
  if(expandedTable&&name!==selectedView)toggleTableExpanded();
  selectedView=name;
  id('workspaceNav').querySelectorAll('[data-view]').forEach(button=>{const active=button.dataset.view===name;button.classList.toggle('active',active);button.setAttribute('aria-selected',String(active))});
  document.querySelectorAll('.fc-view').forEach(element=>element.hidden=element.id!=='view-'+name);id('cycleTab').classList.toggle('active',name==='cycle');
  if(name==='cycle'){cycle.show();return}
  if(rendered.get(name)===revision)return;
  ({portfolio:renderPortfolio,changes:renderChanges,nsr:renderNsrs,summary:renderSummary,rules:renderRules}[name])(id('view-'+name),reuseTables);rendered.set(name,revision);
}
// Persist only leaf selections. Every aggregate is the same set of cells, regardless of entry point.
function normalizeFidStatus(values){
  const selected=new Set(values.flatMap(value=>fidSelection.groups[value]||[value])),postRisk=fidSelection.groups['late-ctl-nsr'];
  if(selected.delete('cause:late-ctl'))postRisk.forEach(cell=>selected.add(cell));
  return fidSelection.order.filter(value=>selected.has(value));
}
function fidStatusSelected(value){
  const parts=fidSelection.groups[value],selected=filterValues.fidStatus||[];
  return parts?parts.length>0&&parts.length===selected.length&&parts.every(part=>selected.includes(part)):selected.includes(value);
}
function changeFilter(path,values){
  let changed=false;
  for(const [key,next] of typeof path==='string'?[[path,values]]:Object.entries(path)){
    delete pendingFilterValues[key];
    const previous=filterValues[key]||[];let selected=key==='fidStatus'?normalizeFidStatus(next):next;
    if(key==='nsr.taskStatus'){
      selected=[...new Set(next)];
      if(selected.includes('15-Completed')&&selected.some(value=>Object.hasOwn(completedTasks,value)))selected=selected.filter(value=>!previous.includes('15-Completed')?!Object.hasOwn(completedTasks,value):value!=='15-Completed');
    }
    changed||=selected.length!==previous.length||selected.some((value,index)=>value!==previous[index]);filterValues[key]=selected;const definition=filterDefs.find(def=>def[2]===key);
    if(definition){if(dateFilters.includes(key))buildFilters([key]);else{const select=id(definition[0]);for(const option of select.options)option.selected=selected.includes(option.value);jQuery(select).trigger('chosen:updated')}}
  }
  if(changed||hasPendingMarks())applyFilters(true);
}
function sourceValue(record,source,key){return graph.field(record,sources[source].fields[key]?.aliases||[key],true)}
// Slot Initiator shows and filters the source value; Owner tie-breaking keeps the normalized name.
function slotInitiator(node){return sourceValue(node.slot.active?.record,'ctl','initiator')}
function fidDateValue(node,state,key){return state.dates[key]||(state.quarter===node.currentQuarter?sourceValue(node.record,'fcid',key):'')}
function recordedFidDates(node){
  let multiple;
  for(const key of fidDateKeys){
    const dates=new Set();
    for(const state of node.quarterStates){
      const value=state.dates[key];
      if(value)for(const date of value.split(' / '))dates.add(date);
    }
    if(dates.size>1)(multiple??={})[key]=[...dates];
  }
  return multiple;
}
function hasUnreadFidDates(fid,key,dates){
  const read=fidDateReads.get(fid)?.[key];
  return dates?.some(date=>!read?.includes(date))||false;
}
function fidDateHistoryLabel(key,value,unread){return fidDateLabels[key]+' recorded dates · '+(value||'—')+(unread?' · unread':'')}
function renderFidDate(node,key,value){
  const dates=fidDateChanges.get(node.id)?.[key],date=escapeHtml(value)||'—';
  if(!dates)return date;
  const unread=hasUnreadFidDates(node.id,key,dates);
  return `<button type="button" class="fid-date-change${unread?' is-unread':''}" data-fid-date="${escapeHtml(node.id)}" data-date-change="${key}" popovertarget="fidDatePopup" aria-haspopup="dialog" aria-label="${escapeHtml(fidDateHistoryLabel(key,value,unread))}">${date}</button>`;
}
function activeCtlLabel(ctl){
  if(!ctl?.number)return '';
  return ctl.status==='Submitted'?`${ctl.number} ${pendingCtlMarker}`:ctl.number;
}
function compareCtlNumbers(a,b){
  const suffix=' '+pendingCtlMarker;
  return tableCollator.compare(a.replace(suffix,''),b.replace(suffix,''));
}
function ctlElapsedDays(ctl){
  return ctl?.status==='Submitted'&&ctl.created?(nsrRecords.dateNumber(nsrData.localDate(new Date()))-nsrRecords.dateNumber(ctl.created))/86400000:'';
}
function renderCtlElapsedDays(days){
  if(!Number.isFinite(days))return '';
  const circle=days>=60?'🔴':days>=30?'🟡':'🟢';
  return `${circle} ${days}`;
}
const fidTableWidths={FID:80,'FID Status & Risks':175,Quarter:105,Note:220};
const baseFidColumns=[markColumn('fid'),
  ['FID',node=>node.id,fidLink],['FID Status & Risks',node=>fidDetailStatus(node,false).map(([label])=>label).join(' · '),(_,node)=>`<span class="fid-status-tags">${fidDetailStatus(node).map(([label,kind,tip])=>pill(label,kind,tip)).join('')}</span>`],
  ['Quarter',node=>fidQuarterRows(node).map(state=>state.quarter||'[Unknown]').join(', ')],['System Description',node=>node.system],['CRD',node=>node.crd||sourceValue(node.record,'fcid','crd')],
  ['SRD',node=>fidQuarterRows(node).map(state=>fidDateValue(node,state,'rd')||'—').join(', '),(value,node)=>renderFidDate(node,'rd',value)],
  ['MCSD',node=>fidQuarterRows(node).map(state=>fidDateValue(node,state,'mcsd')||'—').join(', '),(value,node)=>renderFidDate(node,'mcsd',value)],
  ['Account',node=>node.account],['Customer',node=>node.customer],['FID Owner',node=>node.owner,(value,node)=>`<span title="${escapeHtml(ownerTip(node))}">${escapeHtml(value||'[Unknown]')}</span>`],['Slot Initiator',slotInitiator],
  ['Active CTL',node=>activeCtlLabel(node.activeCtl),(value,node)=>{
    if(!value)return '—';
    const {id,number}=node.activeCtl;
    return fidLink(node.id,{ctlId:id},number)+escapeHtml(value.slice(number.length));
  },{sorter:compareCtlNumbers}],['CTL Elapsed Days',node=>ctlElapsedDays(node.activeCtl),renderCtlElapsedDays],['Active Quote',node=>node.activeCtl?.quote],
  ['Note',node=>fidNotes[node.id]||'',(value,node)=>`<button type="button" class="fc-link fid-note-cell${value?'':' fid-note-add'}" data-note-fid="${escapeHtml(node.id)}" aria-label="${value?'Edit':'Add'} note for FID ${escapeHtml(node.id)}" title="${escapeHtml(value||'Add note')}">${escapeHtml(value?value.replace(/\r\n?|\n/g,' '):'…')}</button>`,{headerFilterFunc:(filter,value)=>{
    const note=String(value??'');return (!filter.presence||(filter.presence==='has')===Boolean(note.trim()))&&note.toLowerCase().includes(String(filter.text??filter).toLowerCase());
  }}],
  ['CTL History',node=>node.ctls.map(ctl=>ctl.number||ctl.id).join(', ')],['NSR History',node=>node.nsrs.map(nsr=>nsr.number).join(', '),(_,node)=>node.nsrs.map(nsr=>nsrWorkflowLink(nsr.number)).join(', ')||'—'],
  ['Commit: Field',node=>node.commit],['Business Unit',node=>node.businessUnit],['MFG Upgrade Type',node=>node.mfg],['FID Description',node=>node.fidDescription],['Fab Name',node=>node.fab]
];
const fidCustomTypes={boolean:'Boolean',number:'Number',text:'Text'};
const defaultFidCustomFields=['FIF','QCK','SHA','CAF'].map(title=>({id:title.toLowerCase(),title,type:'boolean',display:true}));
const reservedFidTitles=new Set([...baseFidColumns.map(([label])=>label),...Object.getOwnPropertyNames(Object.prototype),
  'Assessment quarters','Quarter assessments','Sales Ops RD','Sales Admin','Owner rule','Pending NSR','Completed 5D NSR','Completed NSR but no CTL'].map(label=>label.toLowerCase()));
let fidCustom={fields:defaultFidCustomFields,values:{}},fidCustomDraft=null;
try{fidCustom=readFidCustomColumns()}catch{/* Invalid personal settings leave the default columns available. */}
function normalizeFidCustomColumns(state){
  if(!isRecord(state)||!Array.isArray(state.fields)||!isRecord(state.values))throw new Error('Invalid custom column settings.');
  const ids=new Set(),titles=new Set(reservedFidTitles);
  const fields=state.fields.map(field=>{
    if(!isRecord(field)||typeof field.id!=='string'||!/^[a-z0-9-]{1,50}$/.test(field.id)||ids.has(field.id)||
      typeof field.type!=='string'||!Object.hasOwn(fidCustomTypes,field.type)||typeof field.display!=='boolean')throw new Error('Invalid custom column definition.');
    const title=typeof field.title==='string'?field.title.trim():'';
    if(!title||title.length>80||/[\r\n\t]/.test(title))throw new Error('Enter a column title of up to 80 characters.');
    if(titles.has(title.toLowerCase()))throw new Error('The column title “'+title+'” is already in use.');
    ids.add(field.id);titles.add(title.toLowerCase());return {id:field.id,title,type:field.type,display:field.display};
  });
  const values={};
  for(const [fid,row] of Object.entries(state.values)){
    if(!markKinds.fid.pattern.test(fid)||!isRecord(row))throw new Error('Invalid FID custom values.');
    const entries=Object.entries(row).map(([key,value])=>{
      if(!ids.has(key)||typeof value!=='string'&&typeof value!=='boolean'&&!Number.isFinite(value))throw new Error('Invalid FID custom value.');
      return [key,typeof value==='string'?value.replace(/\r\n?/g,'\n'):value];
    });
    if(entries.length)values[fid]=Object.fromEntries(entries);
  }
  return {fields,values};
}
function readFidCustomColumns(){
  const saved=GM_getValue(FID_CUSTOM_COLUMNS_KEY,null);
  return saved===null?{fields:defaultFidCustomFields,values:{}}:normalizeFidCustomColumns(typeof saved==='string'?JSON.parse(saved):saved);
}
function latestFidCustomColumns(){
  const saved=readFidCustomColumns();
  if(JSON.stringify(saved.fields)!==JSON.stringify(fidCustom.fields))throw new Error('Custom columns changed in another tab. Reload this page before editing.');
  return saved;
}
function fidCustomValue(fid,field){
  const value=fidCustom.values[fid]?.[field.id];
  if(value===undefined)return '';
  if(field.type==='boolean')return /^(?:true|1|✓)$/i.test(String(value).trim())?'✓':'';
  if(field.type==='number'){const numeric=Number(value);return String(value).trim()&&Number.isFinite(numeric)?numeric:''}
  return String(value);
}
function fidCustomButton(fid,field,value,summary=false){
  const empty=value==='',flag=field.type==='boolean',label='Edit '+field.title+' for FID '+fid;
  const className=summary?'fid-custom-tile':flag?'fc-link fid-custom-flag':'fc-link fid-note-cell'+(empty?' fid-note-add':'');
  const content=summary?`<span>${escapeHtml(field.title)}</span><span>${escapeHtml(value)}</span>`:escapeHtml(empty&&!flag?'…':String(value).replace(/\r\n?|\n/g,' '));
  return `<button type="button" class="${className}" data-fid-custom="${field.id}" data-custom-fid="${escapeHtml(fid)}" data-type="${field.type}" aria-haspopup="dialog" aria-label="${escapeHtml(label)}" title="${escapeHtml(empty?label:value)}">${content}</button>`;
}
function fidColumns(){
  const columns=fidCustom.fields.filter(field=>field.display).map(field=>[field.title,node=>fidCustomValue(node.id,field),
    (value,node)=>fidCustomButton(node.id,field,value),{sorter:(a,b)=>typeof a==='number'&&typeof b==='number'?a-b:tableCollator.compare(String(a),String(b))},field]);
  const note=baseFidColumns.findIndex(([label])=>label==='Note');
  return [...baseFidColumns.slice(0,note),...columns,...baseFidColumns.slice(note)];
}
function fidDetailStatus(node,withTips=true,currentOnly=false){
  const tags=[];
  const commit=node.commit?.trim().toLowerCase();
  if(commit==='upside'||commit==='gone')tags.push([commit==='upside'?'Upside':'Gone','commit-status','']);
  if(currentOnly)tags.push([(node.quarterStates.length>1?'Current Risk: ':'')+(node.portfolioStates.includes('slot-ready')?quarterStatus(node):'Risk Status Unknown'),node.portfolioStates.includes('risk-fid')?'risk':node.portfolioStates.includes('risk-free')?'clear':'unknown','']);
  else if(node.portfolioStates.includes('risk-free'))tags.push(['Stable Slot & CTL','clear',withTips?[riskTip(node.pre),riskTip(node.post)].join('\n'):'']);
  if(!node.portfolioStates.includes('slot-ready'))tags.push([fidStatusNames[node.portfolioStates[0]],'unknown',withTips?riskTip(node.pre):'']);
  else if(!currentOnly&&node.pre.state==='risk')tags.push(['Late Slot','risk',withTips?riskTip(node.pre):'']);
  else if(!currentOnly&&node.pre.state==='unknown')tags.push(['Slot date unknown','unknown',withTips?riskTip(node.pre):'']);
  if(node.slot.state==='dropped')tags.push(['Drop Slot','neutral',withTips?'Included in Slotted FID; risk uses the last approved Slot evidence.':'']);
  if(!currentOnly&&node.post.ctls.length)tags.push(['Late CTL ('+node.post.ctls.length+')','risk',withTips?riskTip(node.post):'']);
  const pending=node.post.nsrs.filter(nsr=>includeIgnored||!nsr.ignored);
  if(pending.length)tags.push(['Pending NSR ('+pending.length+')','pending',withTips?'In-work NSRs for this FID: '+pending.map(nsr=>nsr.number).join(', ')+'. Pending NSRs do not change Slot / CTL stability.':'']);
  for(const [label,kind]of [['Pending Config CTL','pending'],['Pending Pull/Push CTL','neutral']]){
    const risk=node.risks.find(risk=>risk.label===label);
    if(risk)tags.push([label+' ('+risk.count+')',kind,withTips?label+' for this FID: Submitted '+(kind==='pending'?'Slot / Config':'Pull / Push')+' CTLs.':'']);
  }
  if(currentOnly)return tags;
  if(node.portfolioStates.includes('slot-ready')&&node.post.state==='unknown')tags.push(['CTL status unknown','unknown',withTips?riskTip(node.post):'']);
  const selected=node.selectedQuarterStates||node.quarterStates,current=node.quarterStates.find(state=>state.quarter===node.currentQuarter);
  if(current&&selected?.length&&(node.pre.state!==current.pre.state||node.post.state!==current.post.state))tags.push(['Current Risk: '+quarterStatus(current),'neutral',withTips?'Current quarter '+(node.currentQuarter||'[Unknown]')+': '+quarterStatus(current):'']);
  if(withTips&&selected?.length)for(const tag of tags)tag[2]=[tag[2],selected.map(state=>(state.quarter||'[Unknown]')+': '+quarterStatus(state)).join('\n')].filter(Boolean).join('\n');
  return tags;
}
function quarterStatus(state,withCount=false){
  if(!state.portfolioStates.includes('slot-ready'))return fidStatusNames[state.portfolioStates[0]];
  return [state.pre.state==='risk'?'Late Slot':'',state.post.state==='risk'?'Late CTL'+(withCount?' ('+state.post.ctls.length+')':''):''].filter(Boolean).join(' + ')||(state.portfolioStates.includes('risk-free')?'Stable Slot & CTL':'Risk Status Unknown');
}
function riskTip(risk){return [risk.reason,risk.eventDate?`${risk.eventDateField||'Event'}: ${risk.eventDate}`:'',risk.referenceDate?`${risk.referenceField||'Reference'}: ${risk.referenceDate}`:'',risk.ctlIds?.length?'CTL: '+risk.ctlIds.slice(0,4).join(', '):''].filter(Boolean).join('\n')}
function ownerTip(node){
  const evidence=node.ownerEvidence,nsrFallback=evidence&&!evidence.ctlCandidates.length&&evidence.nsrCandidates.length;return [`Inferred owner · Account: ${node.account||'[Unknown]'} · people learned from Slot authors`,
    'CTL (Slot / Config): '+((evidence?.ctlCandidates||[]).map(candidate=>candidate.name+': '+candidate.count).join(' · ')||'No eligible CTL votes'),
    nsrFallback?'NSR fallback (all eligible CTL counts are 0): '+evidence.nsrCandidates.map(candidate=>candidate.name+': '+candidate.count).join(' · '):evidence?.tied?'NSR among tied CTL candidates: '+evidence.tiedCandidates.map(candidate=>candidate.name+': '+(evidence.nsrCandidates.find(row=>row.name===candidate.name)?.count||0)).join(' · '):'',
    evidence?.source+(evidence?.tied?(evidence.tieBreaker?' · tie resolved with '+(evidence.source==='slot-tiebreaker'?'current Slot submitter':'NSR submissions'):' · unresolved tie'):''),
    evidence?.excludedCandidates?.length?'Outside Account: '+evidence.excludedCandidates.map(candidate=>candidate.name+': '+candidate.count).join(' · '):'',
    evidence?.aliases?.map(alias=>alias.account+' → '+alias.name).join(' · ')
  ].filter(Boolean).join('\n');
}

function renderPortfolio(target,updateMarks=false){
  updateMarks=updateMarks&&tables.has('portfolioTable');
  if(!Object.keys(store.meta).length){target.innerHTML='';return}
  const riskFilters=queryFilters();delete riskFilters.riskType;
  const riskResult=filterValues.riskType?.length?graph.query(riskFilters):current,riskCounts=new Map([...(store.meta.nsr?['Pending NSR']:[]),...(store.meta.ctl||store.meta.ctlNsr?Object.keys(pendingTypes):[]),...[...id('fRisk').options].map(option=>option.value)].map(label=>[label,{count:0,fids:0}]));
  for(const [label,key]of [['Pending Config CTL','pendingCtl'],['Pending Pull/Push CTL','pendingPullPushCtl']])if(riskCounts.has(label))riskCounts.set(label,{count:riskResult.metrics[key],fids:riskResult.metrics[key+'Fids']});
  const dateChanges=riskCounts.get(dateChangeDriver);
  for(const node of riskResult.fids){
    if(dateChanges&&fidDateChanges.has(node.id)){dateChanges.count++;dateChanges.fids++}
    for(const risk of node.risks){
      const total=riskCounts.get(risk.label);if(!total||risk.label==='Pending NSR'||pendingTypes[risk.label])continue;
      total.count++;total.fids++;
    }
  }
  if(store.meta.nsr){
    const nsrFilters=queryFilters();if(nsrFilters.nsr)delete nsrFilters.nsr.taskStatus;
    const nsrs=filterValues['nsr.taskStatus']?.length?graph.query(nsrFilters,'nsrs'):current.nsrs,total=riskCounts.get('Pending NSR'),fids=new Set();
    for(const nsr of nsrs)if(taskGroup(nsr.taskStatus)==='workflow'){total.count++;for(const fid of nsr.matchedFidIds||nsr.fidIds)fids.add(fid)}
    total.fids=fids.size;
  }
  const filters=queryFilters(),selected=filterValues.fidStatus||[],causes=(filterValues.signal||[]).filter(key=>fidCauseKeys.includes(key));delete filters.fidStatus;
  filters.signal=filters.signal.filter(key=>!fidCauseKeys.includes(key));
  const overviewKey=JSON.stringify(filters);if(portfolioCache.modelRevision!==modelRevision||portfolioCache.key!==overviewKey)portfolioCache={modelRevision,key:overviewKey,rows:selected.length||causes.length?graph.query(filters,'fids'):current.fids};
  const rows=portfolioCache.rows,counts=Object.fromEntries(Object.keys(fidStatusNames).map(key=>[key,0]));
  for(const node of rows){
    for(const key of node.portfolioStates)counts[key]++;
    if(node.portfolioCell.includes('|'))counts[node.portfolioCell]=(counts[node.portfolioCell]||0)+1;
  }
  counts['late-ctl-nsr']=counts['late-ctl'];
  const ctlReady=Boolean(store.meta.ctl||store.meta.ctlNsr);
  const slotKeys=['stable-slot','late-slot',...(counts['slot-date-unknown']?['slot-date-unknown']:[])];
  const postKeys=[!counts['stable-ctl']&&counts['ctl-unknown']?'ctl-unknown':'stable-ctl','late-ctl-nsr'];if(counts['ctl-unknown']&&!postKeys.includes('ctl-unknown'))postKeys.push('ctl-unknown');postKeys.reverse();
  const selectedCells=new Set(selected);
  const card=(key,label,count,color,denominator=rows.length,tip='')=>{
    const value=fidCauseKeys.includes(key)?'cause:'+key:key,active=key==='all'?!selected.length&&!causes.length:fidStatusSelected(value),percent=denominator?(100*count/denominator).toFixed(1):'0.0',base=['all','no-slot','slot-ready'].includes(key)?'Total FID':'Slotted FID';
    const titleName=fidStatusNames[key]||label;
    const title=titleName+' / '+base+' = '+count+' / '+denominator+' = '+percent+'%'+(tip?'\n'+tip:'');
    const shareLabel={'slot-ready':'Slot Completion','stable-slot':'Pre-slot Stability','late-slot':'Pre-slot Risk','stable-ctl':'Post-slot Stability','late-ctl-nsr':'Post-slot Risk'}[key];
    const share=`${percent}%${shareLabel?' '+shareLabel:''}`;
    const buttonClass=key==='no-slot'?'fid-no-slot-main':key==='slot-ready'?'fid-slotted-main fid-aggregate':'nsr-kpi'+(fidSelection.groups[key]?' fid-aggregate':'');
    const button=`<button class="${buttonClass}" style="--kpi-color:${color}" type="button" data-fid-status="${key}" data-fid-value="${value}" data-count="${count}" aria-pressed="${active}" title="${escapeHtml(title)}"><small>${label}</small><b>${number(count)}</b><span class="fid-percent">${share}</span></button>`;
    const subcount=key==='no-slot'?['slot-pending','pending approval','No Slot with a pending Slot request.']:key==='slot-ready'?['slot-dropped','drop slot','Included in Slotted FID. The latest effective event is Drop Slot; risk uses the last approved Slot evidence.']:null;
    if(!subcount)return button;
    const [cause,name,description]=subcount;
    return `<div class="nsr-kpi ${key==='no-slot'?'fid-no-slot':'fid-slotted'}" style="--kpi-color:${color}">${button}<div class="fid-subcount"><button class="fid-subcount-filter" type="button" data-fid-status="${cause}" data-fid-value="cause:${cause}" data-count="${counts[cause]}" aria-pressed="${selected.includes('cause:'+cause)}" title="${description}">${number(counts[cause])} ${name}</button></div></div>`;
  };
  const sign=value=>`<span class="math-sign" aria-hidden="true">${value}</span>`;
  const riskLabelOffset=postKeys[0]==='late-ctl-nsr'?0:(slotKeys.length-1-slotKeys.indexOf('late-slot'))*100/slotKeys.length;
  const riskPercent=counts['slot-ready']?(100*counts['risk-fid']/counts['slot-ready']).toFixed(1):'0.0';
  const relation=(symbol,equation,position='')=>`<span class="fid-relation-sign${position?' floating':''}" data-fid-equation="${equation}"${position?` style="${position}"`:''} aria-hidden="true">${symbol}</span>`;
  const isRisk=(row,column)=>fidRiskAssessment(slotKeys[column]+'|'+postKeys[row])==='risk-fid';
  const outline=[],selection=[],cells=[],cellOperators=[];
  for(let row=0;row<postKeys.length;row++)for(let column=0;column<slotKeys.length;column++){
    const slot=slotKeys[column],post=postKeys[row],assessment=fidRiskAssessment(slot+'|'+post);
    const label=assessment==='risk-unknown'?'Risk Status Unknown':fidCellNames[slot+'|'+post];
    const count=counts[slot+'|'+post]||0,percent=counts['slot-ready']?(100*count/counts['slot-ready']).toFixed(1):'0.0';
    const active=selectedCells.has(slot+'|'+post);
    if(active)selection.push(`<rect data-selection-cell="${slot}|${post}" x="${column*100/slotKeys.length}%" y="${row*100/postKeys.length}%" width="${100/slotKeys.length}%" height="${100/postKeys.length}%"/>`);
    cells.push(`<button class="fid-matrix-cell" type="button" data-fid-cell="${slot}|${post}" data-matrix-slot="${slot}" data-matrix-post="${post}" data-assessment="${assessment}" data-count="${count}" aria-pressed="${active}" title="${escapeHtml(fidStatusNames[slot]+' ∩ '+fidStatusNames[post]+'\n'+count+' / '+counts['slot-ready']+' Slotted FID = '+percent+'%\n'+(assessment==='risk-fid'?'A known Late condition takes priority over missing evidence.':assessment==='risk-unknown'?'No known Late condition; more source data or dates are needed.':'Both Slot and CTL are confirmed Stable. Pending NSRs may still exist.'))}"><small>${escapeHtml(label)}</small><b>${number(count)}</b><span class="fid-percent">${percent}%${assessment==='risk-free'?' Stable FID':''}</span></button>`);
    if(column<slotKeys.length-1)cellOperators.push(relation('+','post:'+post+(isRisk(row,column)&&isRisk(row,column+1)?' risk-fid':''),`left:${(column+1)*100/slotKeys.length}%;top:${(row+.5)*100/postKeys.length}%`));
    if(row<postKeys.length-1)cellOperators.push(relation('+','slot:'+slot+(isRisk(row,column)&&isRisk(row+1,column)?' risk-fid':''),`left:${(column+.5)*100/slotKeys.length}%;top:${(row+1)*100/postKeys.length}%`));
    if(!isRisk(row,column))continue;
    // Draw the outside edges of the union, producing an L for the normal 2 × 2 matrix.
    const x=column*100,y=row*100;
    if(row===0||!isRisk(row-1,column))outline.push(`M${x},${y}h100`);
    if(column===slotKeys.length-1||!isRisk(row,column+1))outline.push(`M${x+100},${y}v100`);
    if(row===postKeys.length-1||!isRisk(row+1,column))outline.push(`M${x},${y+100}h100`);
    if(column===0||!isRisk(row,column-1))outline.push(`M${x},${y}v100`);
  }
  // Dilate the combined cells in screen pixels, then subtract the inner area: adjacent cells share one outer border.
  const selectionOutline=selection.length?`<svg class="fid-selection-outline" aria-hidden="true"><defs><filter id="fid-selection-border" x="-10%" y="-10%" width="120%" height="120%" primitiveUnits="userSpaceOnUse"><feMorphology in="SourceAlpha" operator="dilate" radius="4" result="outer"/><feMorphology in="SourceAlpha" operator="dilate" radius="2" result="inner"/><feComposite in="outer" in2="inner" operator="out" result="border"/><feFlood flood-color="#3566b0" result="color"/><feComposite in="color" in2="border" operator="in"/></filter></defs><g fill="#fff" filter="url(#fid-selection-border)">${selection.join('')}</g></svg>`:'';
  const tips={'stable-slot':'Slot submitted on time. SRD / RD − approved Slot effective date: >20 weeks; No Labor: >12 weeks. Approval confirms completion.','late-slot':'Slot submitted too late. SRD / RD − approved Slot effective date: ≤20 weeks; No Labor: ≤12 weeks. Approval confirms completion.','slot-date-unknown':'Approved Slot history exists, but its original Slot or SRD / RD evidence is missing.','stable-ctl':'No qualifying Late CTL. Each CTL uses its MCSD At SOps Approval, then MCSD@Create, then current MCSD / SRD. Pending NSRs do not change stability.','late-ctl-nsr':'Approved Config CTL with (its MCSD At SOps Approval − Sales Ops approval) / 7 <20 weeks; fallback: its MCSD@Create, then current MCSD / SRD. Reason = Customer Config Chg and one of the seven allowed Sub Reasons. Counted once per FID; Pending NSRs do not change stability.','ctl-unknown':'CTL sources or reference dates are insufficient. A known Late CTL still belongs to Late CTL.'};
  const axis=key=>`<div class="fid-axis-item">${card(key,fidStatusNames[key],counts[key],key.includes('unknown')?'#64748b':key.startsWith('late')?'#b3314c':'#16843f',counts['slot-ready'],tips[key])}</div>`;
  const overview=`${overviewHead('portfolio',current.fids.length,rows.length,selected.length||causes.length)}
    <div class="fid-flow"><div class="fid-flow-inner${ctlReady?' has-breakdown':''}${counts['slot-date-unknown']?' has-slot-unknown':''}" style="--matrix-cols:${slotKeys.length};--matrix-rows:${postKeys.length}"><div class="fid-equation" aria-label="${ctlReady?'Total FID − No Slot = Slotted FID':'Total FID'}">
    ${card('all','Total FID',rows.length,'#94a3b8')}${ctlReady?`${sign('−')}${card('no-slot','No Slot',counts['no-slot'],'#8d5514',rows.length,'No approved Slot or effective Drop Slot evidence.')}${sign('=')}${card('slot-ready','Slotted FID',counts['slot-ready'],'#2563eb',rows.length,'FID has an approved Slot or an effective Drop Slot. Drop Slot remains included in Pre/Post risk classification.')}`:''}</div>
    ${ctlReady?`<div class="fid-origin-pre">${relation('=','origin-pre')}</div><div class="fid-slot-sum" aria-label="Column totals: ${slotKeys.map(key=>fidStatusNames[key]).join(' + ')} = Slotted FID">${slotKeys.map(axis).join('')}${slotKeys.slice(1).map((_,index)=>relation('+','origin-pre',`left:${(index+1)*100/slotKeys.length}%;top:50%`)).join('')}</div>
    <div class="fid-origin-post">${relation('=','origin-post')}</div><div class="fid-breakdown" aria-label="Row totals: ${postKeys.map(key=>fidStatusNames[key]).join(' + ')} = Slotted FID">${postKeys.map(axis).join('')}${postKeys.slice(1).map((_,index)=>relation('+','origin-post',`left:50%;top:${(index+1)*100/postKeys.length}%`)).join('')}</div>
    <div class="fid-row-equals">${postKeys.map(key=>relation('=','post:'+key)).join('')}</div><div class="fid-col-equals">${slotKeys.map(key=>relation('=','slot:'+key)).join('')}</div>
    <div class="fid-matrix" aria-label="Pre-slot Risk × Post-slot Risk matrix"><div class="fid-matrix-grid">${cells.join('')}</div><svg class="fid-risk-outline" viewBox="0 0 ${slotKeys.length*100} ${postKeys.length*100}" preserveAspectRatio="none" aria-hidden="true"><path d="${outline.join('')}"/></svg>${selectionOutline}${cellOperators.join('')}<button class="fid-risk-label" style="right:calc(${riskLabelOffset}% + 14px)" type="button" data-fid-region="risk-fid" data-fid-value="risk-fid" aria-pressed="${fidStatusSelected('risk-fid')}" title="Risk FID = Late Slot ∪ Late CTL&#10;Risk FID / Slotted FID = ${number(counts['risk-fid'])} / ${number(counts['slot-ready'])} = ${riskPercent}%"><span>Risk FID · ${number(counts['risk-fid'])}</span><span class="fid-percent">${riskPercent}%</span></button></div>
    <div class="fid-pre-axis" data-fid-axis="pre" title="Pre: was the Slot submitted on time?"><span>Pre-slot Risk</span></div><div class="fid-post-axis" data-fid-axis="post" title="Post: any qualifying Late CTL in the delivery risk window?"><span>Post-slot Risk</span></div>
    `:''}</div></div>
    ${!ctlReady?'<div class="fid-waiting" role="status">Load a CTL source to determine Slot Completion.</div>':''}`,drivers=`<div class="section-head"><h2>Risk Drivers</h2><span class="badge">Drivers can overlap</span></div><div class="signal-strip">${[...riskCounts].map(([label,{count,fids}])=>`<button class="signal" ${label==='Pending NSR'?'data-nsr-lifecycle="open"':pendingTypes[label]?`data-ctl-status="workflow" data-ctl-types="${pendingTypes[label].join(',')}" data-group-value="${escapeHtml(label)}"${taskFilterValues('ctl','workflow').length?'':' disabled'}`:`data-group-path="riskType" data-group-value="${escapeHtml(label)}"`} aria-pressed="${label==='Pending NSR'?taskSelected('nsr','workflow'):pendingTypes[label]?taskSelected('ctl','workflow',pendingTypes[label]):(filterValues.riskType||[]).includes(label)}" title="${escapeHtml(pretty(label,'riskType'))} = ${count} ${label==='Pending NSR'?'pending NSR':pendingTypes[label]?'in-work CTL documents':'distinct FIDs'}${label.startsWith('Pending ')?' across '+fids+' distinct FID':''}"><b>${number(count)}</b> ${escapeHtml(pretty(label,'riskType'))}${label.startsWith('Pending ')?' <small>| in '+number(fids)+' FID</small>':''}</button>`).join('')||'<span class="muted">No risk drivers in this selection</span>'}</div>`;
  if(updateMarks){
    const scroll=target.querySelector('.fid-flow').scrollLeft;
    target.querySelector('.fid-overview').innerHTML=overview;
    target.querySelector('.fid-flow').scrollLeft=scroll;
    target.querySelector('.risk-drivers').innerHTML=drivers;
  }else target.innerHTML=`<section class="card fid-overview">${overview}</section><section class="card risk-drivers">${drivers}</section>${pivotHtml('fid',current.fids.length,'FIDs',target)}<section class="card"><div class="section-head detail-head"><h2>FID Details</h2></div><div id="portfolioTable"></div></section>`;
  renderPivots('fidPivots',current.fids,[['currentQuarter','Current Quarter','fCurrentQuarter','#4679cf'],['customer','FID Customer','fCustomer','#b44771'],['account','Account','fAccount','#347d50'],['businessUnit','Business Unit','fBusinessUnit','#8251bb'],['system','System Description','fSystem','#14869a']],'FIDs',updateMarks);
  table('portfolioTable',fidColumns(),bwDetailRows('fcid',current.fids),{file:'FID_Portfolio',fid:true,markFilter:'fid',defaultSort:'FID',ascending:false,frozen:['FID'],exportRows:fidDetailMatrix,copyRows:fidCopyRows,widths:fidTableWidths},updateMarks);
}
const ctlColumns=[['CTL',ctl=>ctl.number||ctl.id,(value,ctl)=>ctl.number?`<a class="fc-link" href="https://pcazcrpap02.fremont.lamrc.net:8443/sap/bc/bsp/sap/crm_ui_start/default.htm?saprole=ZSLSDIR&amp;crm-object-type=CRM_SRQM_RFC&amp;crm-object-action=B&amp;crm-object-keyname=OBJECT_ID&amp;crm-object-value=${encodeURIComponent(ctl.number)}" target="_blank" rel="noopener noreferrer">${escapeHtml(value)}</a>`:escapeHtml(value)],['FID',ctl=>ctl.fidId,(value,ctl)=>fidLink(value,{ctlId:ctl.id})],
  ['Stage',ctl=>ctlStageName(ctl.stage)||sourceValue(ctl.record,'ctl','stage'),value=>value?`<span class="status-pill ${value.startsWith('0-')?'stop':value==='15-Completed'?'terminal':value==='Draft'?'unknown':''}">${escapeHtml(value)}</span>`:'—'],
  ['Change',ctl=>sourceValue(ctl.record,'ctl','change')],['Created Date',ctl=>ctl.created||sourceValue(ctl.record,'ctl','created')],
  ['Elapsed Days',ctlElapsedDays,renderCtlElapsedDays],
  ['Sales Ops Approval',ctl=>ctl.approved||sourceValue(ctl.record,'ctl','approved')],['RD At Create',ctl=>ctl.rdCreate||sourceValue(ctl.record,'ctl','rdCreate')],['MCSD@Create',ctl=>ctl.mcsdCreate||sourceValue(ctl.record,'ctl','mcsdCreate')],
  ['SRD @ SOPs Approval',ctl=>ctl.srdApproval||sourceValue(ctl.record,'ctl','slotSrd')],['MCSD At SOps Approval',ctl=>ctl.mcsdApproval||sourceValue(ctl.record,'ctl','mcsdApproval')],['Config Frz',ctl=>ctl.configFrz||sourceValue(ctl.record,'ctl','configFrz')],
  ['Initiator',ctl=>sourceValue(ctl.record,'ctl','initiator')],['Customer',ctl=>ctl.customer],['Quote ID',ctl=>ctl.quote],
  ['NSR Reference',ctl=>ctl.nsrReferences.map(number=>number+linkSuffix(ctl.nsrReferenceTypes[number])).join(', '),(_,ctl)=>ctl.nsrReferences.map(number=>nsrWorkflowLink(number)+linkSuffix(ctl.nsrReferenceTypes[number])).join(', ')||'—'],
  ['Change Summary',ctl=>ctl.changeSummary],['Root Cause',ctl=>ctl.rootCause],['CTL Notes',ctl=>ctl.notes],['Reason for Chg',ctl=>ctl.reason],['Sub Reason',ctl=>ctl.subReason],['Sales OPs Approver',ctl=>ctl.salesOpsApprover]
];
function renderChanges(target,updateMarks=false){
  updateMarks=updateMarks&&tables.has('ctlTable');
  const selected=filterValues['ctl.stage']||[],filters=queryFilters();if(filters.ctl)delete filters.ctl.stage;
  const pendingOnly=filterValues.signal?.includes('pending-ctl'),rows=pendingOnly?current.ctls.filter(ctl=>ctl.pending):current.ctls,overviewRows=(selected.length?graph.query(filters,'ctls'):current.ctls).filter(ctl=>!pendingOnly||ctl.pending);
  const groups={completed:0,workflow:0,draft:0,rejected:0,cancelled:0,unknown:0};for(const ctl of overviewRows)groups[taskGroup(ctlStageName(ctl.stage),true)]++;
  const terms=[['completed','Completed',overviewColors.completed],['workflow','In work',overviewColors.workflow],['draft','Draft',overviewColors.draft],['rejected','Rejected',overviewColors.rejected],['cancelled','Cancelled',overviewColors.cancelled],['unknown','Stage unknown',overviewColors.unknown]].filter(([group])=>group!=='unknown'||groups.unknown||taskSelected('ctl','unknown'));
  const overview=`<div class="section-head"><h2>CTL Overview</h2><span class="badge">${selected.length?number(rows.length)+' / '+number(overviewRows.length)+' selected CTL':number(rows.length)+' filtered CTL'}</span></div><div class="ctl-flow" style="--ctl-terms:${terms.length}" aria-label="Total CTL = ${terms.map(([,label])=>label).join(' + ')}">${taskOverview('ctl',terms,groups,overviewRows.length)}</div>`;
  if(updateMarks)target.querySelector('.ctl-overview').innerHTML=overview;
  else target.innerHTML=`<section class="card ctl-overview">${overview}</section>${pivotHtml('ctl',rows.length,'CTL',target)}<section class="card"><div class="section-head detail-head"><h2>CTL Details</h2></div><div id="ctlTable"></div></section>`;
  renderPivots('ctlPivots',rows,[['quarter','Current Quarter','fCurrentQuarter','#7950a1'],['stage','CTL Stage','fCtlStage','#9468bc'],['customer','Customer','fCtlCustomer','#b491cf']],'CTL',updateMarks);
  table('ctlTable',ctlColumns,bwDetailRows('ctl',rows),{file:pendingOnly?'Pending_CTL':'CTL_Details',defaultSort:'CTL',ascending:false,widths:{CTL:115,FID:85,Stage:145,Change:105,'Created Date':100,'Elapsed Days':130,'Sales Ops Approval':100,'RD At Create':100,'MCSD@Create':100,'SRD @ SOPs Approval':100,'MCSD At SOps Approval':100,'Config Frz':100,Initiator:155,Customer:185,'Quote ID':120,'NSR Reference':180,'Root Cause':300,'CTL Notes':300,'Reason for Chg':180,'Sub Reason':190,'Sales OPs Approver':180}},updateMarks);
}
function recordMark(kind,number){const config=markKinds[kind];return config.ignored.has(number)?'ignore':config.important.has(number)?'important':''}
function markSettings({ignored,important}){return Object.fromEntries([...ignored,...important].sort().map(number=>[number,ignored.has(number)?'ignore':'important']))}
function currentMarks(kind){return markSettings(markKinds[kind])}
function markButton(kind,number){
  const value=recordMark(kind,number),label=value==='ignore'?'⊘ Ignore':value==='important'?'★ Star':'Normal',next=value==='ignore'?'★ Star':value==='important'?'Normal':'⊘ Ignore';
  return `<button type="button" class="record-mark-toggle${kind==='nsr'?' nsr-mark-toggle':''}" data-mark-id="${escapeHtml(number)}" data-ignore-${kind}="${escapeHtml(number)}" data-${kind}-mark="${value}" data-record-mark="${value}" aria-label="${escapeHtml(number)}: ${label}; next ${next}" title="${label} → ${next}\nClick to cycle: Normal → ⊘ Ignore → ★ Star → Normal">${value==='ignore'?'<svg viewBox="0 0 12 12" aria-hidden="true" focusable="false"><circle cx="6" cy="6" r="4.5"/><path d="M3 9 9 3"/></svg>':value==='important'?'<svg viewBox="0 0 12 12" aria-hidden="true" focusable="false"><path fill="currentColor" stroke="none" d="M6 1 7.55 4.3 11 4.8 8.5 7.3 9.1 11 6 9.2 2.9 11 3.5 7.3 1 4.8 4.45 4.3Z"/></svg>':''}</button>`;
}
function markColumn(kind){return ['Mark',row=>recordMark(kind,kind==='fid'?row.id:row.number),(_,row)=>markButton(kind,kind==='fid'?row.id:row.number)]}
function markListHtml(kind){
  const config=markKinds[kind],prefix=config.prefix;
  return `<section class="card" id="ignored${prefix}s"><div class="section-head"><h2>Marked ${config.plural}</h2><button type="button" class="btn blue small" data-save-${kind}-marks>Save</button></div><div class="nsr-mark-editors"><div class="ag-box"><label for="fcIgnored${prefix}s">⊘ Ignore</label><textarea id="fcIgnored${prefix}s" data-${kind}-marks="ignore" rows="8" spellcheck="false" placeholder="One ${config.label} per line"></textarea></div><div class="ag-box"><label for="fcImportant${prefix}s">★ Star</label><textarea id="fcImportant${prefix}s" data-${kind}-marks="important" rows="8" spellcheck="false" placeholder="One ${config.label} per line"></textarea></div></div><div id="${config.status}" class="statistics-state" role="status"></div></section>`;
}
function paintMarkList(kind){
  const config=markKinds[kind];
  for(const textarea of id('view-rules').querySelectorAll(`[data-${kind}-marks]`)){
    const mark=textarea.dataset[kind+'Marks'];
    textarea.value=markDrafts[kind]?.[mark]??[...(mark==='ignore'?config.ignored:config.important)].sort(tableCollator.compare).reverse().join('\n');
  }
}
const markNsrColumn=markColumn('nsr');
const nsrColumns=[
  ['NSR#',nsr=>nsr.number,nsrWorkflowLink],
  ['NSR Title',nsr=>nsr.title],['Task Status',nsr=>nsr.taskStatus,value=>value?`<span class="status-pill ${value==='15-Completed'?'terminal':value.startsWith('0-')?'stop':value.startsWith('Unmapped')?'unknown':''}">${escapeHtml(value)}</span>`:'—'],
  ['NSR Type',nsr=>nsr.type],['Impacted Sub System',nsr=>nsr.subsystem],['Initiator',nsr=>graph.field(nsr.record,['Initiator','Initiated By','Requester'],true)],['FID',nsr=>(nsr.matchedFidIds||nsr.fidIds||[nsr.fidId]).filter(Boolean).join(', '),(_,nsr)=>(nsr.matchedFidIds||nsr.fidIds||[nsr.fidId]).filter(Boolean).map(fid=>fidLink(fid,{nsrId:nsr.id})).join(', ')||'—'],
  ['Primary Product',nsr=>nsr.product],
  ['NSR → CTL',nsr=>lifecycleNames[nsr.lifecycle]||nsr.lifecycle,(_,nsr)=>pill(lifecycleNames[nsr.lifecycle]||nsr.lifecycle,nsr.ignored?'ignore':nsr.lifecycle==='completed-unlinked'?'alert':['completed-5d','implemented'].includes(nsr.lifecycle)?'clear':nsr.lifecycle==='ctl-pending'?'pending':'info')],
  ['CTL references',nsr=>(nsr.links||[]).map(link=>link.ctlNumber+linkSuffix(link.referenceType,link.kind==='reference'?link.fidId:'')).join(', '),(_,nsr)=>(nsr.links||[]).map(link=>`<button type="button" class="fc-link" data-fid="${escapeHtml(link.fidId)}" data-ctl-id="${escapeHtml(link.ctlId)}" title="Open Slot / CTL History">${escapeHtml(link.ctlNumber)}</button>${escapeHtml(linkSuffix(link.referenceType,link.kind==='reference'?link.fidId:''))}`).join(', ')||'—'],
  ['Customer',nsr=>nsr.customer],['NSR Org',nsr=>nsr.org],['Sales Rep',nsr=>sourceValue(nsr.record,'nsr','Sales Rep')],['Sales Ops',nsr=>sourceValue(nsr.record,'nsr','Sales Ops')],['Technical Contact',nsr=>sourceValue(nsr.record,'nsr','Technical Contact')],
  ['Submit Date',nsr=>nsr.submitted||sourceValue(nsr.record,'nsr','submitted')],['NSR Request',nsr=>sourceValue(nsr.record,'nsr','NSR Request')],['Temp Id#',nsr=>nsr.tempId||sourceValue(nsr.record,'nsr','tempId')]
];
function nsrCopyRows(rows,columns=nsrColumns){return nsrDetailRows(rows,columns.filter(([label])=>label!=='Mark'))}
function markedDetailValue(kind,row,[label,get]){if(label!=='Mark')return get(row)??'';const mark=recordMark(kind,kind==='fid'?row.id:row.number);return mark==='ignore'?'⊘':mark==='important'?'★':''}
function markedDetailRows(kind,rows,columns){return rows.map(row=>Object.fromEntries(columns.map(column=>[column[0],markedDetailValue(kind,row,column)])))}
function nsrDetailRows(rows,columns=[markNsrColumn,...nsrColumns]){return markedDetailRows('nsr',rows,columns)}
function fidDetailMatrix(rows,columns){return [columns.map(([label])=>label),...rows.map(row=>columns.map(column=>markedDetailValue('fid',row,column)))]}
function fidCopyRows(rows,columns){return fidDetailMatrix(rows,columns.filter(([label])=>label!=='Mark'))}
function renderNsrDetail(name,rows,options={},updateMarks=false){
  const columns=[markNsrColumn,...nsrColumns],widths={'Task Status':225,'NSR#':80,FID:80,'NSR Org':96};
  if(name==='fidNsrTable'){columns.splice(2,0,...columns.splice(columns.findIndex(([label])=>label==='NSR → CTL'),2));widths['NSR → CTL']=150}
  return table(name,columns,rows,{file:'NSR_Detail',cycle:true,nsr:true,exportRows:nsrDetailRows,copyRows:nsrCopyRows,defaultSort:'FID',ascending:false,frozen:['NSR#'],widths,...options},updateMarks);
}
function highlightHistory(name,record){
  const state=tables.get(name);if(!state?.ready)return;
  const key=name==='fidCtlTable'?'ctlId':'nsrId',linked=new Map(record?.links?.map(link=>[link[key],link.referenceType]));
  for(const row of state.grid.getRows()){const element=row.getElement(),type=linked.get(state.rows[row.getIndex()].id);element.classList.toggle('nsr-ctl-match',!!type);element.classList.toggle('nsr-ctl-note',type==='Note')}
}
function pivotHtml(kind,count,unit,target){return `<details class="card interactive-pivot"${target.querySelector('.interactive-pivot')?.open===false?'':' open'}><summary class="section-head"><h2>Interactive Pivot<span class="pivot-toggle" aria-hidden="true" title="Expand / collapse Interactive Pivot"></span></h2><span class="badge">${number(count)} filtered ${unit}</span></summary><div id="${kind}Pivots" class="nsr-pivots${kind==='ctl'?' ctl-pivots':''}"></div></details>`}
function renderPivots(name,rows,pivots,unit,updateMarks=false){
  const target=id(name);
  if(updateMarks)target.previousElementSibling.querySelector('.badge').textContent=number(rows.length)+' filtered '+unit;
  pivots.forEach(([path,label,filter,color],index)=>{
    const counts=new Map();
    for(const row of rows){const value=row[path]||'[Unknown]';counts.set(value,(counts.get(value)||0)+1)}
    const entries=[...counts].sort((a,b)=>path==='taskStatus'?nsrRecords.compareStatus(a[0],b[0]):path==='stage'?compareCtlStages(a[0],b[0]):path.toLowerCase().includes('quarter')?Number(a[0]==='[Unknown]')-Number(b[0]==='[Unknown]')||nsrRecords.periodRank(a[0])-nsrRecords.periodRank(b[0])||a[0].localeCompare(b[0]):b[1]-a[1]||a[0].localeCompare(b[0])),max=Math.max(1,...counts.values());
    const content=entries.map(([value,count])=>`<div class="pivot-row"><span class="pivot-label" title="${escapeHtml(path==='stage'?ctlStageName(value):value)}">${escapeHtml(path==='stage'?ctlStageName(value):value)}</span><span class="pivot-bar"><i style="width:${100*count/max}%;--bar:${color}"></i></span><button class="pivot-count" data-pivot-filter="${filter}" data-pivot-value="${escapeHtml(value)}" title="Filter ${escapeHtml(value)}">${number(count)}</button></div>`).join('')||'<span class="muted">No matching '+unit+'</span>';
    if(updateMarks){const card=target.children[index],list=card.querySelector('.pivot-list'),scroll=list.scrollTop;card.querySelector('.pivot-head span').textContent=counts.size+' values';list.innerHTML=content;list.scrollTop=scroll}
    else target.insertAdjacentHTML('beforeend',`<section class="pivot-card"><div class="pivot-head"><h3>${label}</h3><span>${counts.size} values</span></div><div class="pivot-list">${content}</div></section>`);
  });
}
function renderNsrs(target,updateMarks=false){
  const rows=current.nsrs,selected=filterValues['nsr.taskStatus']||[],filters=queryFilters(),previous=tables.get('nsrTable');if(filters.nsr)delete filters.nsr.taskStatus;
  const sameRows=updateMarks&&previous&&rows.length===previous.rows.length&&rows.every((row,index)=>{const before=previous.rows[index];return row.number===before.number&&row.record===before.record&&row.taskStatus===before.taskStatus&&row.fidId===before.fidId&&row.matchedFidIds?.join()===before.matchedFidIds?.join()});
  const overviewRows=selected.length?graph.query(filters,'nsrs'):rows,groups={workflow:0,completed:0,rejected:0,cancelled:0},stageCounts={'completed-5d':0,'completed-unlinked':0,'ctl-pending':0,'implemented':0,'ctl-unavailable':0};
  overviewRows.forEach(nsr=>{const group=taskGroup(nsr.taskStatus);groups[group]++;if(group==='completed')stageCounts[Object.hasOwn(stageCounts,nsr.taskLifecycle)?nsr.taskLifecycle:'ctl-unavailable']++});
  const terms=[['completed','Completed',overviewColors.completed],['workflow','In workflow',overviewColors.workflow],['rejected','Rejected',overviewColors.rejected],['cancelled','Cancelled',overviewColors.cancelled]];
  const completedStages=[['completed-unlinked','alert','#a62a49','Completed NSR → no matching CTL in loaded data'],['ctl-pending','pending','#8a5c0c','Linked CTL → awaiting Sales Ops approval'],['implemented','approved',overviewColors.completed,'Linked CTL → Sales Ops approved'],['completed-5d','approved',overviewColors.completed,'Completed D-prefix (5D) NSR → no CTL required']];
  const showCompleted=groups.completed||selected.some(value=>value==='15-Completed'||Object.hasOwn(completedTasks,value)),unknown=stageCounts['ctl-unavailable'];
  const completedEquation=`${groups.completed} = ${completedStages.map(([key])=>stageCounts[key]).join(' + ')}${unknown?' + '+unknown+' unknown':''}`;
  const overview=`${overviewHead('nsr',rows.length,overviewRows.length,selected.length)}<div class="nsr-flow${showCompleted?' has-completed':''}"><div class="nsr-equation" aria-label="Total NSR = ${terms.map(([,label])=>label).join(' + ')}">${taskOverview('nsr',terms,groups,overviewRows.length)}</div>
  ${showCompleted?`<div class="nsr-breakdown" role="group" aria-label="Completed NSR breakdown"><span class="nsr-completed-arrow" title="${escapeHtml(completedEquation)}" aria-hidden="true"></span><span class="math-sign nsr-breakdown-equals" aria-hidden="true">=</span><div class="nsr-stage-sum"><span class="nsr-stage-spacer" aria-hidden="true"></span>${completedStages.map(([key,kind,color,tip],index)=>{const count=stageCounts[key],active=taskSelected('nsr',key),percent=groups.completed?(100*count/groups.completed).toFixed(1):'0.0';return (index?'<span class="math-sign" aria-hidden="true">+</span>':'')+`<button class="nsr-kpi nsr-stage ${kind}" style="--kpi-color:${color}" type="button" data-nsr-lifecycle="${key}" data-count="${count}" aria-pressed="${active}"${!count&&!active?' disabled':''} title="${escapeHtml(`${count} / ${groups.completed} = ${percent}%\n${tip}`)}"><small>${lifecycleNames[key]}</small><b>${number(count)}</b><span>${percent}% of Completed NSRs</span></button>`}).join('')}${unknown?`<span class="math-sign" aria-hidden="true">+</span><span class="nsr-kpi nsr-stage unknown" style="--kpi-color:#52647d" data-count="${unknown}" role="status" title="CTL sources are not loaded or approval status is missing. View these records under 15-Completed."><small>CTL status unknown</small><b>${number(unknown)}</b><span>${groups.completed?(100*unknown/groups.completed).toFixed(1):'0.0'}% of Completed NSRs</span></span>`:''}</div></div>`:''}</div>`;
  if(updateMarks)target.querySelector('.nsr-overview').innerHTML=overview;
  else target.innerHTML=`<section class="card nsr-overview">${overview}</section>${pivotHtml('nsr',rows.length,'NSRs',target)}<section class="card"><div class="section-head detail-head"><h2>NSR Details</h2></div><div id="nsrTable"></div></section>`;
  const pivots=[['taskStatus','Task Status','fNsrTaskStatus','var(--nsr-ink)'],['subsystem','Impacted Sub System','fNsrSubsystem','var(--nsr-ink)'],['type','NSR Type','fNsrType','var(--nsr-ink)'],['product','Primary Product','fNsrProduct','var(--nsr-ink)'],['org','NSR Org','fNsrOrg','var(--nsr-ink)'],['customer','Customer','fNsrCustomer','var(--nsr-ink)']];
  if(!sameRows)renderPivots('nsrPivots',rows,pivots,'NSRs',updateMarks);
  renderNsrDetail('nsrTable',rows,{markFilter:'nsr'},updateMarks);
}
const summaryDimensions=[['quarter','Assessment Quarters'],['currentQuarter','Current Quarter'],['account','Account'],['customer','FID Customer'],['owner','Owner']];
const metricDefs=[
  ['Slot Completion%','slotted','fid','slot','slot-ready'],['Pre-slot Stability%','stableSlot','slotted','slot','stable-slot'],['Post-slot Stability%','stablePost','slotted','post','stable-ctl'],
  ['FID','fid','','fcid'],['Risk FID%','riskFid','slotted','slot','risk-fid'],['Stable Slot & CTL%','noLate','slotted','post','risk-free'],
  ['NSR','nsr','','nsr'],['In work NSR','open','','nsr'],['Completed NSR','completed','','nsr'],['Cancelled/Rejected NSR','discardedNsr','','nsr'],['Completed NSR but no CTL','unlinked','','linked'],
  ['CTL','ctl','','ctl'],['Submitted CTL','submitted','','ctl'],['Completed CTL','ctlCompleted','','ctl'],['Draft CTL','draft','','ctl'],['Cancelled/Rejected CTL','discardedCtl','','ctl'],['Pending Config CTL','pending','','ctl'],['Pending Pull/Push CTL','pendingPullPush','','ctl']
];
let summaryGroup='quarter',previousGroup='quarter',metricKeys=['slotted','stableSlot','stablePost','fid','nsr','ctl'],summarySort={key:'',dir:-1},summaryCache={},summaryReport=[],quarterCache={};
function summaryGroups(path){
  if(path==='quarter'&&quarterCache.revision!==revision)quarterCache={revision,result:graph.query(queryFilters(),'periods')};
  const result=path==='quarter'?quarterCache.result:current,quartersByFid=new Map();
  if(path==='quarter')for(const node of result.fids){if(!quartersByFid.has(node.id))quartersByFid.set(node.id,[]);quartersByFid.get(node.id).push(node.quarter||'[Unknown]');}
  const fidNames=fid=>path==='quarter'?quartersByFid.get(fid)||graph.fids.get(fid)?.quarterStates.map(state=>state.quarter||'[Unknown]')||['[Unknown]']:[graph.fids.get(fid)?.[path]||'[Unknown]'];
  const groups=new Map(),fidMetrics=metricDefs.filter(([,,,,state])=>state),entry=name=>{
    name=name||'[Unknown]';if(!groups.has(name))groups.set(name,{name,...Object.fromEntries(metricDefs.map(([,key])=>[key,0]))});return groups.get(name);
  };
  for(const node of result.fids){
    const row=entry(node[path]);row.fid++;
    for(const [,key,,,state]of fidMetrics)row[key]+=node.portfolioStates.includes(state);
  }
  for(const nsr of result.nsrs){
    const fids=nsr.matchedFidIds||nsr.fidIds,groupNames=fid=>{const names=fidNames(fid);return path==='account'&&names[0]==='[Unknown]'?[nsr.account||'[Unknown]']:names;};
    const names=new Set(path==='owner'?[nsr.initiator||'[Unknown]']:fids.length?fids.flatMap(groupNames):[path==='account'?nsr.account||'[Unknown]':'[Unknown]']),status=taskGroup(nsr.taskStatus);
    for(const name of names){
      const row=entry(name);row.nsr++;row.open+=status==='workflow';row.completed+=status==='completed';row.discardedNsr+=status==='cancelled'||status==='rejected';
      row.unlinked+=nsr.taskLifecycle==='completed-unlinked'&&(path==='owner'||!fids.length||fids.some(fid=>groupNames(fid).includes(name)&&nsr.implementationByFid[fid]==='completed-unlinked'));
    }
  }
  for(const ctl of result.ctls){
    for(const name of path==='owner'?[ctl.initiator]:fidNames(ctl.fidId)){
      const row=entry(name);row.ctl++;row.pending+=ctl.pending;row.pendingPullPush+=ctl.pendingPullPush;
      if(ctl.statusKind==='completed')row.ctlCompleted++;else if(ctl.statusKind==='discarded')row.discardedCtl++;else if(ctl.status==='Submitted')row.submitted++;else if(ctl.status==='Draft')row.draft++;
    }
  }
  return [...groups.values()].sort((a,b)=>Number(a.name==='[Unknown]')-Number(b.name==='[Unknown]')||(path==='owner'?b.riskFid-a.riskFid||b.fid-a.fid||b.nsr-a.nsr||b.ctl-a.ctl:['quarter','currentQuarter'].includes(path)?nsrRecords.periodRank(a.name)-nsrRecords.periodRank(b.name):0)||tableCollator.compare(a.name,b.name));
}
function renderSummary(target){
  if(!id('summaryChart')){
    target.innerHTML=`<section class="card"><div class="section-head"><div class="summary-heading overview-tools"><h2>Summary Metrics</h2><span id="summaryCount" class="badge" aria-live="polite"></span>${ignoredToggle('summary')}</div><div class="btnrow"><button class="btn light small" data-summary-export="copy">Copy</button><button class="btn light small" data-summary-export="xlsx">Export XLSX</button></div></div><div class="summary-controls"><div class="fc-filter"><label for="summaryGroup">Group by</label><select id="summaryGroup"></select></div><div class="fc-filter"><label for="summaryMetrics">Metrics</label><select id="summaryMetrics" multiple>${metricDefs.map(([label,key])=>`<option value="${key}"${metricKeys.includes(key)?' selected':''}>${label}</option>`).join('')}</select></div></div><div id="summaryChart" class="summary-chart" aria-live="polite"></div><div id="summaryLegend" class="summary-legend" hidden>FID by Owner · NSR / CTL by Initiator.</div></section>`;
    jQuery(id('summaryMetrics')).chosen({width:'100%',search_contains:true,hide_results_on_select:false,display_selected_options:false,placeholder_text_multiple:'Select metrics'});
    jQuery('#summaryGroup,#summaryMetrics').on('change',event=>{if(event.target.id==='summaryGroup'){summaryGroup=event.target.value;if(summaryGroup!=='owner')previousGroup=summaryGroup}else metricKeys=[...event.target.selectedOptions].map(option=>option.value);renderSummary(target)});
  }
  id('summaryGroup').innerHTML=summaryDimensions.filter(([value])=>value!=='owner'||summaryGroup==='owner').map(([value,label])=>`<option value="${value}"${value===summaryGroup?' selected':''}>${label}</option>`).join('');
  id('summaryLegend').hidden=summaryGroup!=='owner';
  const chart=id('summaryChart');summaryReport=[];
  if(!metricKeys.includes(summarySort.key))summarySort.key='';
  for(const button of target.querySelectorAll('[data-summary-export]'))button.disabled=true;
  if(!metricKeys.length){chart.innerHTML='<div class="summary-empty">Select at least one metric.</div>';id('summaryCount').textContent='';return}
  if(summaryCache.revision!==revision||summaryCache.group!==summaryGroup){
    summaryCache={revision,group:summaryGroup,rows:summaryGroups(summaryGroup)};
    if(['quarter','currentQuarter'].includes(summaryGroup))summarySort.key='';
  }
  const rows=summaryCache.rows.slice(),ctlReady=Boolean(store.meta.ctl||store.meta.ctlNsr),ready={fcid:!!store.meta.fcid,ctl:ctlReady,nsr:!!store.meta.nsr,slot:!!store.meta.fcid&&ctlReady,post:!!store.meta.fcid&&ctlReady,linked:ctlReady&&!!store.meta.nsr};
  id('summaryCount').textContent=rows.length+' groups';
  if(!rows.length){chart.innerHTML='<div class="summary-empty">No matching records.</div>';return}
  const metrics=metricDefs.filter(([,key])=>metricKeys.includes(key)),dimension=summaryDimensions.find(([key])=>key===summaryGroup)[1];
  const sortMetric=metrics.find(([,key])=>key===summarySort.key);
  if(sortMetric&&ready[sortMetric[3]]){
    const [,key,base]=sortMetric;
    rows.sort((a,b)=>{
      const left=base?(a[base]?a[key]/a[base]:null):a[key],right=base?(b[base]?b[key]/b[base]:null):b[key];
      return Number(left===null)-Number(right===null)||(left===null||right===null?0:summarySort.dir*(left-right));
    });
  }
  summaryReport=[['Metric',...rows.map(row=>row.name)]];
  for(const [label,key,base,source,countState]of metrics){
    summaryReport.push([label,...rows.map(row=>ready[source]&&(!base||row[base]>0)?(base?row[key]/row[base]:row[key]):'—')]);
    if(base)summaryReport.push([fidStatusNames[countState],...rows.map(row=>ready[source]?row[key]:'—')]);
  }
  for(const button of target.querySelectorAll('[data-summary-export]'))button.disabled=false;
  const content=metrics.map(([label,key,base,source])=>{
    const max=base?100:Math.max(1,...rows.map(row=>row[key]));
    const cells=rows.map(row=>{
      const available=ready[source]&&(!base||row[base]>0),value=available?(base?100*row[key]/row[base]:row[key]):null,display=value===null?'—':base?value.toFixed(1)+'%':number(value);
      const tip=!ready[source]?'Required source data is not loaded.':base?label+' = '+row[key]+'/'+row[base]+' = '+display:label+' = '+display;
      return `<td title="${escapeHtml(row.name+' · '+tip)}"><div class="summary-plot${available?'':' unavailable'}" style="--bar:${value===null?0:64*value/max}px" role="img" aria-label="${escapeHtml(row.name+' · '+tip)}"><i class="summary-bar" aria-hidden="true"></i><span class="summary-value">${display}</span></div></td>`;
    }).join('');
    return `<tr data-summary-metric="${key}" data-summary-source="${source}"><th scope="row" aria-sort="${summarySort.key===key?(summarySort.dir===1?'ascending':'descending'):'none'}"><button type="button" class="fc-link" data-summary-sort="${key}" title="Sort groups by ${label}">${label} ${summarySort.key===key?(summarySort.dir===1?'↑':'↓'):'↕'}</button></th>${cells}</tr>`;
  }).join('');
  chart.innerHTML=`<table class="summary-table" style="min-width:${180+110*rows.length}px" aria-label="Summary Metrics by ${dimension}"><colgroup><col style="width:180px">${rows.map(()=>'<col>').join('')}</colgroup><thead><tr><th scope="col">Metric</th>${rows.map(row=>`<th scope="col">${escapeHtml(row.name)}</th>`).join('')}</tr></thead><tbody>${content}</tbody></table>`;
}
const nceColumns=[['NCE',row=>row.number],['State',row=>row.status],['Title',row=>row.title],['Impact',row=>row.impact],['Opened',row=>row.opened],['Damage Group',row=>row.damageGroup],['Damage Code',row=>row.damageCode],['Part Number',row=>row.partNumber],['Part Description',row=>row.partDescription]];

const tableCollator=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'}),numericCollator=new Intl.Collator(undefined,{numeric:true});
const columnPopup=id('columnFilterPopup'),columnInput=columnPopup.querySelector('input'),columnChoices=id('colChoices'),columnNoteFilter=id('colNoteFilter'),columnPendingCtl=id('colPendingCtl'),markPopup=id('markFilterPopup');
const fidDatePopup=id('fidDatePopup');
// Hidden tabs keep their measured size, so showing an unchanged table does not rebuild its rows.
const tableResize=new ResizeObserver(entries=>{for(const {target} of entries){const state=tables.get(target.dataset.tableName);if(!state?.ready||!target.clientWidth||!target.clientHeight)continue;const size=target.clientWidth+':'+target.clientHeight;if(size!==state.size){state.size=size;state.grid.redraw()}}});
function tableColumnKey(column){return column[4]?'fidCustom:'+column[4].id:column[0]}
function orderTableColumns(definitions,columns,order,nsrHistory=false){
  if(nsrHistory){const leading=['NSR → CTL','CTL references'];order=[...leading,...(order||columns.map(([label])=>label)).filter(label=>!leading.includes(label))]}
  if(!Array.isArray(order))return definitions;
  const note=order.indexOf('Note');
  if(note>=0){
    const added=columns.filter(column=>column[4]).map(tableColumnKey).filter(key=>!order.includes(key));
    if(added.length)order=[...order.slice(0,note),...added,...order.slice(note)];
  }
  const rank=new Map(order.map((label,index)=>[label,index])),label=column=>tableColumnKey(columns[Number(column.field.slice(1))]);
  const movable=new Set(definitions.filter(column=>!column.frozen&&rank.has(label(column)))),sorted=[...movable].sort((a,b)=>rank.get(label(a))-rank.get(label(b)));let index=0;
  return definitions.map(column=>movable.has(column)?sorted[index++]:column);
}
function table(name,columns,rows,options={},reuse=false,editedIds=null){
  // Tabulator deep-clones row data; keep the full graph records in state.rows.
  const previous=tables.get(name),dataRow=(row,index)=>({id:index,...Object.fromEntries(columns.map(([,get],column)=>['c'+column,get(row)??'']))});
  if(previous&&!previous.ready&&previous.built)return previous.built.then(()=>table(name,columns,rows,options,reuse,editedIds));
  const sameColumns=!options.fid||columns.length===previous?.columns.length&&columns.every((column,index)=>column[0]===previous.columns[index][0]&&column[4]===previous.columns[index][4]);
  if(previous?.ready&&reuse&&sameColumns){
    const data=editedIds?rows.flatMap((row,index)=>editedIds.has(options.fid?row.id:row.number)?[dataRow(row,index)]:[]):rows.map(dataRow);
    if(editedIds&&!data.length)return previous.built;
    const wrap=previous.grid.element,wasHidden=wrap.hidden,active=document.activeElement,focused=wrap.contains(active)?active.dataset.markId||active.dataset.noteFid:'',focusSelector=active.matches('[data-mark-id]')?'[data-mark-id]':'[data-note-fid]';previous.rows=rows;previous.columns=columns;previous.options=options;wrap.setAttribute('aria-busy','true');
    // Keep the applied rows and sort order while the user edits cell values.
    const updated=editedIds?previous.grid.updateData(data):previous.grid.replaceData(data);
    return updated.then(()=>{
      if(editedIds)paintMarkFilter(previous);
      if(wasHidden&&!wrap.hidden)redrawTable(previous);wrap.setAttribute('aria-busy','false');if(focused)[...wrap.querySelectorAll(focusSelector)].find(button=>(button.dataset.markId||button.dataset.noteFid)===focused)?.focus({preventScroll:true});
      return previous;
    });
  }
  const retainedWidths=previous?.ready&&!sameColumns?new Map(previous.grid.getColumns().map(column=>[tableColumnKey(previous.columns[Number(column.getField().slice(1))]),column.getWidth()])):null;
  if(previous){tableResize.unobserve(previous.grid.element);previous.grid.destroy()}
  const local=localTables.get(name)||{marks:name==='portfolioTable'?defaultFidDetailMarks:null,sort:options.defaultSort?columns.findIndex(([label])=>label===options.defaultSort):-1,ascending:options.ascending??true},orderKey=name==='fidCtlTable'?'ctlTable':name==='fidNsrTable'?'nsrTable':name,state={name,columns,rows,options,local,orderKey};
  if(options.fid){
    const keys=columns.map(tableColumnKey),before=local.columnKeys;
    if(before&&(keys.length!==before.length||keys.some((key,index)=>key!==before[index]))){
      const indices=new Map(keys.map((key,index)=>[key,index])),sort=indices.get(before[local.sort]);
      local.sort=sort??columns.findIndex(([label])=>label===options.defaultSort);
      if(sort===undefined)local.ascending=options.ascending??true;
      local.headerFilters=(local.headerFilters||[]).flatMap(filter=>{
        const index=indices.get(before[Number(filter.field.slice(1))]);return index===undefined?[]:[{...filter,field:'c'+index}];
      });
    }
    local.columnKeys=keys;
  }
  if(name==='fidCtlTable'&&!id('fidDetail').open)local.showPullPush=false;
  tables.set(name,state);localTables.set(name,local);
  const host=id(name),autoHeight=Boolean(host.closest('.fid-dialog')),columnFilters=!autoHeight&&options.columnFilters!==false;if(!columnFilters)local.headerFilters=[];
  const heading=autoHeight?null:host.parentElement.querySelector(':scope > .detail-head, :scope > summary');
  host.innerHTML=`<div class="fc-table-shell"><div class="table-toolbar" data-table-tools="${name}">${autoHeight?`<div class="btnrow"><h2>${escapeHtml(options.title)} (<span data-table-count></span>)</h2>${name==='fidCtlTable'?`<button type="button" class="fc-link" data-toggle-pull-push aria-pressed="${Boolean(local.showPullPush)}">${local.showPullPush?'Hide':'Show'} Pull/Push</button>`:''}</div>`:'<span class="table-count" data-table-count></span>'}<div class="toolbar-right${autoHeight?' btnrow':''}"><button class="btn light small" data-table-copy="${name}"${options.copyId?` id="${options.copyId}"`:''}>Copy</button>${autoHeight?'':`<button class="btn light small" data-table-export="${name}"${options.exportId?` id="${options.exportId}"`:''}>Export XLSX</button>`}${heading?.matches('.detail-head')?`<button class="btn light small table-expand" type="button" data-table-expand="${name}" aria-expanded="false" aria-label="Expand table" title="Expand table"><svg viewBox="0 0 24 24" aria-hidden="true"><path class="expand-icon" d="M14 10l6-6m-5 0h5v5M10 14l-6 6m5 0H4v-5"/><path class="restore-icon" d="M20 4l-6 6m0-5v5h5M4 20l6-6m-5 0h5v5"/></svg></button>`:''}${options.cycle?`<button class="btn blue small" data-table-cycle="${name}">Cycle Time</button>`:''}</div></div>${columnFilters?'<div class="col-bar"></div>':''}<div class="fc-grid" data-table-name="${name}" aria-label="${escapeHtml(options.file||name)}" aria-busy="true"></div></div>`;
  state.toolbar=host.querySelector('.table-toolbar');
  if(heading&&state.toolbar){heading.querySelector('[data-table-tools="'+name+'"]')?.remove();heading.append(state.toolbar)}
  const wrap=host.querySelector('.fc-grid');wrap.classList.toggle('column-filters',columnFilters);
  const data=rows.map(dataRow),widths={...options.widths};
  if(options.fid){
    const measure=document.createElement('canvas').getContext('2d'),font=getComputedStyle(wrap).fontFamily;
    for(const [index,[label]]of columns.entries()){
      if(label==='Mark'||widths[label]!==undefined)continue;
      measure.font='600 12px '+font;let width=measure.measureText(label).width+44+(label==='FID Owner'?19:0);
      measure.font='13px '+font;
      for(const value of new Set(data.map(item=>String(item['c'+index])))){
        width=Math.max(width,measure.measureText(value).width+20);if(width>=220)break;
      }
      widths[label]=Math.ceil(Math.max(75,Math.min(220,width)));
    }
    widths.SRD=widths.MCSD=widths.CRD;
    if(retainedWidths)for(const column of columns){const width=retainedWidths.get(tableColumnKey(column));if(width>0)widths[column[0]]=width}
  }
  const filter=item=>(name!=='fidCtlTable'||local.showPullPush||!['pull','push'].includes(state.rows[item.id].changeKey))&&(!options.markFilter||!local.marks||local.marks.includes(recordMark(options.markFilter,options.markFilter==='fid'?state.rows[item.id].id:state.rows[item.id].number)||'normal'));
  const columnDefs=columns.map(([label,,render,settings],index)=>({title:label==='Mark'?'':label,titleFormatter:label==='Mark'&&options.markFilter?()=>markHeader(state):columnFilters&&label!=='Mark'?columnHeader:undefined,field:'c'+index,frozen:index===0||(options.frozen||[]).includes(label),minWidth:label==='Mark'?36:75,headerSort:label!=='Mark',headerFilter:columnFilters&&label!=='Mark'?'input':false,headerFilterParams:{elementAttributes:{'aria-label':'Filter '+label}},cellClick:options.fid&&label==='FID Status & Risks'?(event,cell)=>{if(event.target.closest('#wfModal')&&id('fidDetail').open)id('fidDetail').close();showFid(state.rows[cell.getRow().getIndex()].id)}:undefined,hozAlign:label==='Mark'?'center':'left',headerHozAlign:label==='Mark'?'center':'left',width:widths[label]??(label==='Mark'?36:label==='Task Status'?275:/Title|Summary|Documents|Sample|File|Reason|Description|Action/.test(label)?320:/Owner|Initiator|Candidate|Customer|Product/.test(label)?185:/^Pre$|^Post$|NSR → CTL|Application/.test(label)?180:Math.max(110,Math.min(240,label.length*7+28))),sorter:(a,b)=>label==='Stage'?compareCtlStages(a,b):/Quarter|Period/.test(label)?nsrRecords.periodRank(a)-nsrRecords.periodRank(b)||tableCollator.compare(String(a),String(b)):typeof a==='number'&&typeof b==='number'?a-b:tableCollator.compare(String(a),String(b)),...settings,formatter:cell=>{const value=cell.getValue();cell.getElement().title=label==='Mark'?'':String(value);return '<span class="cell-content">'+(render?render(value,state.rows[cell.getRow().getIndex()]):escapeHtml(value)||'—')+'</span>'}}));
  let order=tableOrders[orderKey];
  if(name==='cycleTable'&&Array.isArray(order)&&!order.includes('Additional FCID')){const index=order.indexOf('REF FCID');if(index>=0)order=[...order.slice(0,index+1),'Additional FCID',...order.slice(index+1)]}
  state.grid=new Tabulator(wrap,{data,autoResize:false,movableColumns:true,height:autoHeight?false:options.height??Math.min(620,Math.max(105,rows.length*34+68)),rowHeight:options.rowHeight??34,renderVertical:autoHeight?'basic':'virtual',renderVerticalBuffer:200,columnDefaults:{resizable:'header',headerFilterFunc:'like',headerFilterPlaceholder:'Filter…'},layout:'fitDataFill',placeholder:'',dataLoader:false,initialFilter:filter,initialHeaderFilter:local.headerFilters||[],initialSort:local.sort>=0?[{column:'c'+local.sort,dir:local.ascending?'asc':'desc'}]:[],rowFormatter:row=>{const element=row.getElement(),data=row.getData();element.dataset.rowIndex=row.getIndex();if(state.options.rowRisk)element.dataset.fidRisk=state.options.rowRisk(state.rows[row.getIndex()]);if(options.nsr||options.fid){const mark=recordMark(options.nsr?'nsr':'fid',options.nsr?state.rows[data.id].number:state.rows[data.id].id);element.dataset.recordMark=mark;if(options.nsr)element.dataset.nsrMark=mark;else element.dataset.fidMark=mark;}if(name==='nsrTable'){element.dataset.fidBand=data.fidBand;element.dataset.fidStart=data.fidStart}},columns:orderTableColumns(columnDefs,columns,order,name==='fidNsrTable')});
  state.grid.on('tableDestroyed',()=>{if(columnPopup.column?.getTable()===state.grid)columnPopup.hidePopover();if(markPopup.state===state)markPopup.hidePopover()});
  state.grid.on('columnMoved',()=>{
    if(state.ordering)return;
    const order=state.grid.getColumns().map(column=>tableColumnKey(columns[Number(column.getField().slice(1))])).filter(label=>name!=='fidNsrTable'||!['NSR → CTL','CTL references'].includes(label));let index=0;
    const base=tableOrders[orderKey]||(name==='fidNsrTable'?['Mark',...nsrColumns.map(([label])=>label)]:[]),merged=base.map(label=>order.includes(label)?order[index++]:label);merged.push(...order.slice(index));tableOrders[orderKey]=merged;
    try{GM_setValue(TABLE_ORDERS_KEY,JSON.stringify(tableOrders))}catch(error){toast('Could not save column order: '+error.message)}
    for(const target of tables.values())if(target.ready&&target.orderKey===orderKey&&(target!==state||name==='fidNsrTable')){
      target.ordering=true;
      try{orderTableColumns(target.grid.getColumns().map(column=>column.getDefinition()),target.columns,merged,target.name==='fidNsrTable').forEach((column,index)=>{const current=target.grid.getColumns()[index];if(!column.frozen&&current.getField()!==column.field)target.grid.moveColumn(column.field,current.getField(),false)})}
      finally{target.ordering=false}
    }
  });
  state.grid.on('dataFiltered',(_filters,rows)=>{
    local.headerFilters=state.grid.getHeaderFilters();if(columnFilters)paintColFilters(state);const count=state.toolbar?.querySelector('[data-table-count]');if(count)count.textContent=number(rows.length)+(autoHeight?'':' rows');
    if(autoHeight){
      wrap.hidden=!rows.length;state.toolbar.querySelector('.toolbar-right').hidden=!rows.length;
      const toggle=state.toolbar.querySelector('[data-toggle-pull-push]');if(toggle)toggle.hidden=!state.rows.length;
    }
    paintMarkFilter(state);
  });
  if(options.markFilter)state.grid.on('renderComplete',()=>{if(markPopup.state===state)markPopup.position=markMenuPos(state)});
  state.grid.on('dataSorted',(sorters,rows)=>{
    if(sorters.length){local.sort=Number(sorters[0].field.slice(1));local.ascending=sorters[0].dir==='asc'}
    if(name!=='nsrTable')return;const fidSorted=sorters.length===1&&columns[local.sort]?.[0]==='FID';wrap.classList.toggle('fid-sorted',fidSorted);if(!fidSorted)return;
    // Keep bands on row data so sorting does not create DOM for offscreen rows.
    let previous,band=0;const bands=new Map();for(const row of rows){const data=row.getData(),fid=data[sorters[0].field],start=previous!==undefined&&fid!==previous;if(start)band=1-band;data.fidBand=String(band);data.fidStart=String(start);bands.set(row.getIndex(),data);previous=fid}
    for(const element of wrap.querySelectorAll('.tabulator-row[data-row-index]')){const data=bands.get(Number(element.dataset.rowIndex));if(data){element.dataset.fidBand=data.fidBand;element.dataset.fidStart=data.fidStart}}
  });
  state.grid.on('tableBuilt',()=>state.grid.getColumns().forEach(column=>{const element=column.getElement();element.tabIndex=0;element.title=column.getDefinition().title}));
  if(name==='fidNsrTable'||name==='fidCtlTable'){
    const linkedTable=name==='fidNsrTable'?'fidCtlTable':'fidNsrTable';
    state.grid.on('rowMouseEnter',(_event,row)=>highlightHistory(linkedTable,state.rows[row.getIndex()]));
    state.grid.on('rowMouseLeave',()=>highlightHistory(linkedTable));
  }
  state.grid.on('dataProcessed',()=>{state.ready=true;state.size=wrap.clientWidth+':'+wrap.clientHeight;wrap.setAttribute('aria-busy','false')});tableResize.observe(wrap);
  return state.built=new Promise(resolve=>state.grid.on('tableBuilt',()=>resolve(state)));
}
function columnHeader(cell){
  const column=cell.getColumn(),title=column.getDefinition().title,header=document.createElement('span');header.className='column-heading';
  header.innerHTML='<span>'+escapeHtml(title)+(title==='FID Owner'?helpIcon(ownerHint):'')+'</span><button type="button" class="column-funnel" popovertarget="columnFilterPopup" aria-label="Filter '+escapeHtml(title)+'" aria-haspopup="dialog" aria-pressed="false" title="Filter '+escapeHtml(title)+'"><svg viewBox="0 0 16 16" aria-hidden="true"><path d="M1 2h14L10 8v5l-4 2V8z"/></svg></button>';
  const button=header.lastElementChild;button.onmousedown=event=>event.stopPropagation();
  button.onclick=event=>{
    event.stopPropagation();markPopup.hidePopover();columnPopup.hidePopover();columnPopup.column=column;columnPopup.querySelector('label>span').textContent=title;
    const value=column.getHeaderFilterValue()??'';columnInput.value=value.text??value;columnNoteFilter.hidden=title!=='Note';columnNoteFilter.value=value.presence||'';
    columnPendingCtl.hidden=title!=='Active CTL';
    const state=tables.get(column.getTable().element.dataset.tableName),get=state.columns[Number(column.getField().slice(1))][1],values=new Set();
    for(const row of state.rows){const value=String(get(row)??'');values.add(value);if(values.size>20||/[\r\n]/.test(value))break}
    columnPopup.choices=values.size<=20&&![...values].some(value=>/[\r\n]/.test(value))?[...values].filter(Boolean).sort(title==='Active CTL'?compareCtlNumbers:tableCollator.compare):[];paintColChoices();
  };return header;
}
function paintColChoices(){
  const search=columnInput.value.toLowerCase(),values=(columnPopup.choices||[]).filter(value=>value.toLowerCase().includes(search));
  columnChoices.replaceChildren(...values.map(value=>new Option(value,value)));columnChoices.hidden=!values.length;columnChoices.value=columnInput.value;
}
function columnFilterText(value){return typeof value==='object'?[value.presence==='has'?'Has note':'No note',value.text].filter(Boolean).join(' · '):value}
function paintColFilters(state){
  const filters=state.local.headerFilters,key=JSON.stringify(filters);if(state.columnFilterKey===key)return;state.columnFilterKey=key;
  const host=id(state.name),active=new Map(filters.map(filter=>[filter.field,filter.value]));
  for(const column of state.grid.getColumns()){
    const button=column.getElement().querySelector('.column-funnel');if(!button)continue;const selected=active.has(column.getField());button.classList.toggle('active',selected);button.setAttribute('aria-pressed',String(selected));button.title='Filter '+column.getDefinition().title+(selected?': '+columnFilterText(active.get(column.getField())):'');
  }
  host.querySelector('.col-bar').innerHTML=filters.map(({field,value})=>{const text=escapeHtml(state.grid.getColumn(field).getDefinition().title+': '+columnFilterText(value));return '<button type="button" class="pill info" data-column-remove="'+state.name+'" data-field="'+field+'" title="'+text+'" aria-label="Remove filter '+text+'">'+text+' ×</button>'}).join('')+(filters.length?'<button type="button" class="fc-link" data-column-clear="'+state.name+'">Clear</button>':'');
}
function tableResult(state){return state.grid.getRows('active').map(row=>state.rows[row.getIndex()])}
function redrawTable(state){const widths=state.grid.getColumns().map(column=>[column,column.getWidth()]);state.grid.redraw();for(const [column,width]of widths)if(column.getWidth()!==width)column.setWidth(width)}
function paintTable(state){const wasHidden=state.grid.element.hidden;state.grid.refreshFilter();if(wasHidden&&!state.grid.element.hidden)redrawTable(state)}

function toggleTableExpanded(button=expandedTable?.querySelector('[data-table-expand]')){
  const name=button.dataset.tableExpand,state=tables.get(name);
  if(!expandedTable&&!state?.ready)return;
  const host=id(name),heading=button.closest('.detail-head'),wrap=state.grid.element,holder=wrap.querySelector('.tabulator-tableholder');
  const {scrollLeft,scrollTop}=holder;
  const overlays=[columnPopup,markPopup,id('fcToast')];
  // A modal makes controls outside it inert, including fixed-position overlays.
  if(workNotes.parentElement===(expandedTable||document.body))overlays.push(workNotes);
  columnPopup.hidePopover();markPopup.hidePopover();closeFidDateHistory();
  if(expandedTable){
    const dialog=expandedTable,card=dialog.parentElement;
    dialog.before(heading,host);
    card.style.minHeight=dialog.dataset.inlineMinHeight;
    if(dialog.contains(fidDatePopup))overlays.push(fidDatePopup);
    document.body.append(...overlays);
    dialog.close();dialog.remove();expandedTable=null;
  }else{
    const card=host.parentElement,dialog=document.createElement('dialog');
    dialog.className='fc-table-dialog';
    dialog.setAttribute('aria-label',heading.querySelector('h2').textContent);
    dialog.dataset.inlineMinHeight=card.style.minHeight;
    card.style.minHeight=card.offsetHeight+'px';
    card.append(dialog);
    dialog.append(heading,host,...overlays);
    dialog.oncancel=event=>{event.preventDefault();toggleTableExpanded()};
    expandedTable=dialog;dialog.showModal();
  }
  button.title=expandedTable?'Restore table size (Esc)':'Expand table';
  button.setAttribute('aria-label',expandedTable?'Restore table size':'Expand table');
  button.setAttribute('aria-expanded',String(Boolean(expandedTable)));
  if(state.ready){
    redrawTable(state);holder.scrollLeft=scrollLeft;holder.scrollTop=scrollTop;
    state.size=wrap.clientWidth+':'+wrap.clientHeight;
  }
  button.focus({preventScroll:true});
}


function exportTableRows(state,copy=false){
  const rows=tableResult(state),columns=state.grid.getColumns().filter(column=>column.isVisible()).map(column=>state.columns[Number(column.getField().slice(1))]),format=copy?state.options.copyRows:state.options.exportRows;
  if(!rows.length)return [columns.filter(([label])=>!copy||label!=='Mark').map(([label])=>label)];
  return format?format(rows,columns):rows.map(row=>Object.fromEntries(columns.map(([label,get])=>[label,get(row)??''])));
}
function exportWorkbook(filename,sheets,rowFormats={},options={}){
  try{
    const workbook=XLSX.utils.book_new();
    for(const [name,rows]of Object.entries(sheets)){
      const matrix=Array.isArray(rows[0]),sheet=matrix?XLSX.utils.aoa_to_sheet(rows):XLSX.utils.json_to_sheet(rows);
      if(matrix)for(let row=1;row<rows.length;row++){
        const format=rowFormats[rows[row][0]];if(!format)continue;
        for(let column=1;column<rows[row].length;column++){const cell=sheet[XLSX.utils.encode_cell({r:row,c:column})];if(cell?.t==='n')cell.z=format}
      }
      XLSX.utils.book_append_sheet(workbook,sheet,name.slice(0,31));
    }
    XLSX.writeFile(workbook,filename+'.xlsx',options);
  }
  catch(error){toast('Export failed: '+error.message)}
}
function exportAllDataSources(){
  if(store.refreshing||store.busy.size){toast('Wait for data sources to finish loading before exporting.');return}
  const sheets=Object.fromEntries(Object.entries(sources).filter(([key])=>store.raw[key].length).map(([key,source])=>[key==='nce'?'NCE FA Review':source.label,store.raw[key]]));
  if(!Object.keys(sheets).length){toast('No data sources to export.');return}
  exportWorkbook('ToolkitKyra_Data_Sources_'+nsrData.localDate(new Date()),sheets,{},{compression:true});
}
async function copyRows(rows){const matrix=Array.isArray(rows[0]),fields=matrix?[]:[...new Set(rows.flatMap(Object.keys))],content=(matrix?rows:[fields,...rows.map(row=>fields.map(key=>row[key]??''))]).map(row=>row.map(value=>(/[\r\n\t"]/.test(String(value))?'"'+String(value).replace(/"/g,'""')+'"':String(value))).join('\t')).join('\r\n');try{if(navigator.clipboard&&unsafeWindow.isSecureContext)await navigator.clipboard.writeText(content);else GM_setClipboard(content,'text');toast(number(rows.length-(matrix?1:0))+' rows copied')}catch(error){toast('Copy failed: '+error.message)}}
function renderRules(target,updateMarks=false){
  const excluded=graph.query({...queryFilters(),excluded:true},'fids'),reasons=new Map();for(const node of excluded)reasons.set(node.excludedReason,(reasons.get(node.excludedReason)||0)+1);
  if(!updateMarks)target.innerHTML=`<div class="btnrow" style="justify-content:flex-end"><button class="btn light small" data-import-settings>Import Settings</button><button class="btn light small" data-export-settings>Export Settings</button><button class="btn light small" data-export-data-sources>Export All Data Sources</button><input id="fcSettingsFile" type="file" accept=".json,application/json" hidden></div><details class="card"><summary>Excluded Commit FIDs · ${number(excluded.length)}</summary><div class="signal-strip" style="margin-top:10px">${[...reasons].map(([reason,count])=>`<span class="signal"><b>${number(count)}</b>${escapeHtml(reason)}</span>`).join('')}</div><div id="excludedTable"></div></details><div class="marked-settings">${markListHtml('nsr')}${markListHtml('fid')}</div>
  <form class="card" id="fidCustomSettings" aria-labelledby="fidCustomTitle">
    <div class="section-head"><h2 id="fidCustomTitle">FID Custom Columns</h2><button type="submit" class="btn blue small">Save</button></div>
    <div class="fid-custom-grid" id="fidCustomRows"></div>
    <button type="button" class="btn light small" data-add-fid-custom><span aria-hidden="true">+</span> Add Column</button>
    <div class="statistics-state" id="fidCustomStatus" role="status"></div>
  </form>
  <section class="card" id="fabAccountMapping"><div class="section-head"><h2>Fab / Account Mapping</h2><div class="btnrow"><button class="btn blue small" data-save-mapping>Save</button><button class="btn light small" data-reset-mapping>Restore Defaults</button></div></div><p class="muted">Use Fab names and FID / NSR Customer names. Exact names take priority over * wildcards.</p><div id="agNotice"></div><div id="agGrid" class="ag-grid"></div></section>`;
  table('excludedTable',[['FID',node=>node.id,fidLink],['Reason',node=>node.excludedReason],['System',node=>node.system],['Account',node=>node.account],['Quarter',node=>node.quarter]],excluded,{file:'Excluded_Commit_FIDs'},updateMarks);
  for(const kind of Object.keys(markKinds))paintMarkList(kind);
  if(!updateMarks){paintMapping();paintFidCustomSettings()}
}
function paintFidCustomSettings(){
  const fields=fidCustomDraft||fidCustom.fields;
  id('fidCustomRows').innerHTML='<span>Title</span><span>Type</span><span>Display?</span>'+fields.map(({id,title,type,display},index)=>{
    const row=index+1,types=Object.entries(fidCustomTypes).map(([key,label])=>`<option value="${key}"${type===key?' selected':''}>${label}</option>`).join('');
    return `<div data-custom-row="${id}">
      <input type="text" maxlength="80" aria-label="Column ${row} title" value="${escapeHtml(title)}" placeholder="Column title" autocomplete="off">
      <select aria-label="Column ${row} type">${types}</select>
      <input type="checkbox" aria-label="Display column ${row}"${display?' checked':''}>
      <button type="button" class="btn light small" data-delete-fid-custom aria-label="Delete column ${row}" title="Delete column ${row}" aria-haspopup="dialog" aria-controls="fidCustomDeleteDialog">
        <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18M9 6V3h6v3M19 6l-1 15H6L5 6m5 4v7m4-7v7"/></svg>
      </button>
    </div>`;
  }).join('');
}
function readFidCustomDraft(){
  return [...id('fidCustomRows').querySelectorAll('[data-custom-row]')].map(row=>({id:row.dataset.customRow,title:row.querySelector('input[type=text]').value,type:row.querySelector('select').value,display:row.querySelector('input[type=checkbox]').checked}));
}
async function deleteFidCustomColumn(button){
  const row=button.closest('[data-custom-row]'),fieldId=row.dataset.customRow;
  const saved=fidCustom.fields.find(field=>field.id===fieldId),title=row.querySelector('input[type=text]').value.trim()||saved?.title;
  const dialog=id('fidCustomDeleteDialog');dialog.returnValue='cancel';
  id('fidCustomDeleteTitle').textContent=title?`Delete “${title}”?`:'Delete this column?';
  id('fidCustomDeleteHelp').textContent=saved?'Click Save to delete this column and its values for all FIDs.':'This column will be removed from the draft.';
  const confirmed=await new Promise(resolve=>{
    dialog.addEventListener('close',()=>resolve(dialog.returnValue==='delete'),{once:true});dialog.showModal();
  });
  if(!confirmed)return;
  const fields=readFidCustomDraft(),index=fields.findIndex(field=>field.id===fieldId);
  if(index<0)return;
  fields.splice(index,1);fidCustomDraft=fields;paintFidCustomSettings();
  const status=id('fidCustomStatus');status.className='statistics-state';status.textContent=saved?'Click Save to apply changes.':'';
  const buttons=id('fidCustomRows').querySelectorAll('[data-delete-fid-custom]');
  (buttons[Math.min(index,buttons.length-1)]||id('fidCustomSettings').querySelector('[data-add-fid-custom]')).focus({preventScroll:true});
}
function saveFidCustomSettings(){
  const status=id('fidCustomStatus');
  try{
    const stored=latestFidCustomColumns(),savedIds=new Set(stored.fields.map(field=>field.id));
    const fields=readFidCustomDraft().filter(field=>savedIds.has(field.id)||field.title.trim());
    const remainingIds=new Set(fields.map(field=>field.id)),removed=stored.fields.filter(field=>!remainingIds.has(field.id));
    if(removed.length){
      for(const row of Object.values(stored.values)){
        for(const {id} of removed)delete row[id];
      }
    }
    const next=normalizeFidCustomColumns({fields,values:stored.values});
    GM_setValue(FID_CUSTOM_COLUMNS_KEY,JSON.stringify(next));
    fidCustom=next;fidCustomDraft=null;rendered.delete('portfolio');paintFidCustomSettings();
    status.className='statistics-state';status.textContent='Saved';
  }catch(error){status.className='statistics-state error';status.textContent=error.message}
}
function paintMapping(){if(!id('agGrid'))return;const mapping=mappingDraft||accountMapping.getMap(),status=accountMapping.getStatus();id('agGrid').innerHTML=Object.entries(mapping).map(([group,fabs],index)=>`<div class="ag-box"><label for="ag-${index}">${escapeHtml(group)}</label><textarea id="ag-${index}" data-account-group="${escapeHtml(group)}">${escapeHtml(fabs)}</textarea></div>`).join('');id('agNotice').hidden=!(status.notice||status.error);id('agNotice').className='ag-notice';id('agNotice').textContent=status.error||status.notice}
function setRecordMark(kind,number){
  const mark=recordMark(kind,number),next=mark==='ignore'?'important':mark==='important'?'':'ignore',marks=currentMarks(kind);if(next)marks[number]=next;else delete marks[number];
  const result=saveMarks(kind,marks);if(!result.ok){toast('Could not save '+kind.toUpperCase()+' mark: '+result.error);return}
  const started=performance.now();
  updateSavedMarks(kind,new Set([number]));
  store.buildMs=performance.now()-started;
}
function updateSavedMarks(kind,numbers){
  for(const number of numbers){
    // A scoped view keeps the applied marks until its next explicit filter.
    if(!includeIgnored)pendingMarks[kind].add(number);
    else if(kind==='fid')graph.setFidMark(number,recordMark(kind,number));
    else graph.setNsrMark(number,recordMark(kind,number));
  }
  refreshPersonalRows(kind,numbers);
  if(!markDrafts[kind])paintMarkList(kind);
}
function saveMarks(kind,marks,persist=true){
  const config=markKinds[kind];if(persist)try{GM_setValue(config.key,JSON.stringify(marks))}catch(error){return {ok:false,error:error.message}}
  config.ignored.clear();config.important.clear();for(const [number,mark]of Object.entries(marks))(mark==='ignore'?config.ignored:config.important).add(number);
  portfolioCache={};
  return {ok:true};
}
function saveMarkLists(kind){
  const config=markKinds[kind],lists=Object.fromEntries([...document.querySelectorAll(`[data-${kind}-marks]`)].map(textarea=>[textarea.dataset[kind+'Marks'],new Set(textarea.value.toUpperCase().split(/[\s,;|，；]+/).filter(Boolean))]));
  const invalid=[...lists.ignore,...lists.important].filter(number=>!config.pattern.test(number)),overlap=[...lists.ignore].filter(number=>lists.important.has(number)),status=id(config.status);
  status.className='statistics-state';status.textContent='';
  if(invalid.length||overlap.length){status.className+=' error';status.textContent=invalid.length?'Invalid '+config.label+': '+[...new Set(invalid)].join(', '):config.label+' appears in both lists: '+overlap.join(', ');return}
  const previous=currentMarks(kind),marks=Object.fromEntries(Object.entries(lists).flatMap(([mark,numbers])=>[...numbers].map(number=>[number,mark]))),result=saveMarks(kind,marks);
  if(!result.ok){status.className+=' error';status.textContent='Could not save: '+result.error;return}
  markDrafts[kind]=null;
  updateSavedMarks(kind,new Set([...Object.keys(previous),...Object.keys(marks)].filter(number=>previous[number]!==marks[number])));
  id(config.status).textContent='Saved';
}
function markMenuPos(state){const holder=state.grid.element.querySelector('.tabulator-tableholder');return holder?[scrollX,scrollY,holder.scrollLeft,holder.scrollTop].join(':'):''}
function markHeader(state){
  const button=document.createElement('button');button.type='button';button.className='mark-trigger';button.setAttribute('popovertarget','markFilterPopup');button.setAttribute('aria-haspopup','dialog');
  button.innerHTML='<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M1 2h14L10 8v5l-4 2V8z"/></svg>';
  button.onmousedown=event=>event.stopPropagation();
  button.onclick=event=>{event.stopPropagation();columnPopup.hidePopover();markPopup.hidePopover();markPopup.state=state;markPopup.trigger=button;markPopup.position=markMenuPos(state);paintMarkFilter(state)};
  return button;
}
function paintMarkFilter(state){
  const kind=state.options.markFilter;if(!kind)return;
  const selected=state.local.marks??markOptions,active=Boolean(state.local.marks),trigger=state.grid.element.querySelector('.mark-trigger'),label='Filter marks: '+(active?(selected.length?selected.map(mark=>markLabels[mark]).join(', '):'None'):'All');
  if(trigger){trigger.classList.toggle('active',active);trigger.setAttribute('aria-pressed',String(active));trigger.setAttribute('aria-label',label);trigger.title=label}
  if(markPopup.state!==state)return;
  const host=markPopup.querySelector('[role=group]'),counts={all:0,normal:0,ignore:0,important:0};
  const filters=state.grid.getHeaderFilters().map(filter=>typeof filter.type==='function'?{field:filter.type}:filter);
  for(const {id:index}of state.grid.searchData(filters)){const record=state.rows[index];counts.all++;counts[recordMark(kind,kind==='fid'?record.id:record.number)||'normal']++}
  if(host.dataset.tableName!==state.name){host.dataset.tableName=state.name;host.innerHTML=['all',...markOptions].map(mark=>`<label><input type="checkbox" data-table-mark-filter="${state.name}" data-mark="${mark}"><span>${markLabels[mark]}</span><b aria-hidden="true"></b></label>`).join('')}
  for(const input of host.querySelectorAll('input')){const mark=input.dataset.mark;input.checked=mark==='all'?selected.length===markOptions.length:selected.includes(mark);input.indeterminate=mark==='all'&&selected.length>0&&selected.length<markOptions.length;input.dataset.count=counts[mark];input.setAttribute('aria-label',markLabels[mark]+' ('+number(counts[mark])+')');input.parentElement.querySelector('b').textContent=number(counts[mark])}
}
function paintFidPersonal(fid){
  const dialog=id('fidDetail');if(dialog.dataset.fid!==fid||!id('fidPersonalMark'))return;
  const focused=document.activeElement.dataset.ignoreFid===fid,note=fidNotes[fid]||'';id('fidPersonalMark').innerHTML=markButton('fid',fid);
  const summary=id('fidCustomSummary'),fields=fidCustom.fields.filter(field=>field.display);
  summary.hidden=!fields.length;summary.innerHTML=fields.map(field=>fidCustomButton(fid,field,fidCustomValue(fid,field),true)).join('');
  const panel=id('fidNotePanel');if(!panel.querySelector('form'))panel.innerHTML=`<button type="button" class="fid-note-view" data-note-fid="${escapeHtml(fid)}" aria-label="Edit note for FID ${escapeHtml(fid)}"><span class="fid-note-title" role="heading" aria-level="2">Note</span><span class="fid-note-text${note?'':' muted'}">${escapeHtml(note||'No note')}</span></button>`;
  if(focused)id('fidPersonalMark').querySelector('button').focus({preventScroll:true});
}
function refreshPersonalRows(kind,numbers){
  const updates=[...tables.values()].filter(state=>state.options[kind]).map(state=>table(state.name,state.columns,state.rows,state.options,true,numbers));
  const fid=id('fidDetail').dataset.fid;if(kind==='fid'&&numbers.has(fid))paintFidPersonal(fid);
  return Promise.all(updates);
}
async function syncFidNotes(){
  try{
    const next=readFidNotes(),changed=new Set([...Object.keys(fidNotes),...Object.keys(next)].filter(fid=>fidNotes[fid]!==next[fid]));
    fidNotes=next;
    if(changed.size)await refreshPersonalRows('fid',changed);
  }catch(error){toast('Could not refresh FID notes. '+error.message)}
}
function editFidNote(fid,anchor,field){
  const dialog=id('fidNoteDialog'),panel=field?null:anchor.closest('#fidNotePanel'),host=panel||dialog,rect=anchor.getBoundingClientRect();
  host.dataset.fid=fid;host.dataset.table=anchor.closest('[data-table-name]')?.dataset.tableName||'';host.dataset.field=field?.id||'';host.dataset.customType=field?.type||'';
  const title=field?.title||'Note',value=field?fidCustomValue(fid,field):fidNotes[fid]||'',label=escapeHtml(title);
  const flag=field?.type==='boolean';
  const checked=value==='✓';
  const action=flag?(checked?'Clear ':'Mark ')+title:'Save';
  const input=flag?'':
    field?.type==='number'?`<input id="fidNoteText" type="number" step="any" aria-label="${label}" value="${escapeHtml(value)}">`:
    `<textarea id="fidNoteText" rows="8" aria-label="${label}" placeholder="${field?'Enter text…':'Write a note…'}" spellcheck="false"></textarea>`;
  host.innerHTML=`<form class="fid-note-form"${flag?'':' id="fidNoteForm"'}>
    <div class="section-head"><h2${panel?'':' id="fidNoteTitle"'}>${label}${panel?'':' · FID '+escapeHtml(fid)}</h2></div>${input}
    <div class="statistics-state" role="status"${flag?'':' id="fidNoteStatus"'}></div>
    <div class="fc-note-actions">
      <button class="btn light" type="button" data-cancel-fid-note${flag?' autofocus':''}>Cancel</button>
      <button class="btn blue" type="submit"${flag?` value="${!checked}"`:''}>${escapeHtml(action)}</button>
    </div>
  </form>`;
  if(!panel){
    dialog.setAttribute('closedby',flag?'any':'closerequest');
    dialog.style.bottom='';dialog.style.maxHeight='';
    if(!dialog.open)dialog.showModal();
    const box=dialog.getBoundingClientRect();
    let left=rect.left;
    if(flag){
      // Leave the trigger outside the confirmation, even when the available space is short.
      const above=rect.top-16,below=innerHeight-rect.bottom-16;
      const openAbove=box.height>below&&above>below;
      dialog.style.maxHeight=Math.max(0,openAbove?above:below)+'px';
      dialog.style.top=openAbove?'auto':Math.max(8,rect.bottom+8)+'px';
      dialog.style.bottom=openAbove?Math.max(8,innerHeight-rect.top+8)+'px':'auto';
    }else{
      const save=dialog.querySelector('button[type="submit"]').getBoundingClientRect();
      left+=rect.width/2-(save.left-box.left+save.width/2);
      const top=rect.top+rect.height/2-(save.top-box.top+save.height/2);
      dialog.style.top=Math.max(8,Math.min(top,innerHeight-box.height-8))+'px';
    }
    dialog.style.left=Math.max(8,Math.min(left,innerWidth-box.width-8))+'px';
  }
  const focusTarget=host.querySelector(flag?'[data-cancel-fid-note]':'#fidNoteText');
  if(focusTarget.tagName==='TEXTAREA')focusTarget.defaultValue=focusTarget.value=value;
  focusTarget.focus({preventScroll:true});
}
function closeFidNote(form){
  const host=form.parentElement;
  const {fid,field,table:name}=host.dataset;if(host.id==='fidNoteDialog')host.close();host.replaceChildren();if(host.id==='fidNotePanel')paintFidPersonal(fid);
  const target=id(name)||(id('fidDetail').open?id('fidDetail'):id('portfolioTable'));
  const buttons=target?.querySelectorAll(field?'[data-fid-custom]':'[data-note-fid]')||[];
  [...buttons].find(button=>field?button.dataset.customFid===fid&&button.dataset.fidCustom===field:button.dataset.noteFid===fid)?.focus({preventScroll:true});
}
async function saveFidNote(form,submitter){
  const {fid,field:key,customType}=form.parentElement.dataset,input=form.querySelector('#fidNoteText');
  if(customType==='boolean'&&!submitter)return;
  try{
    if(key){
      const stored=latestFidCustomColumns(),field=stored.fields.find(field=>field.id===key);
      if(!field)throw new Error('This custom column is unavailable.');
      const value=field.type==='boolean'?submitter.value==='true':field.type==='number'&&input.value!==''?input.valueAsNumber:input.value;
      if(typeof value==='number'&&!Number.isFinite(value))throw new Error('Enter a valid number.');
      const values={...stored.values},row={...values[fid]};
      if(value===false||typeof value==='string'&&!value.trim())delete row[key];else row[key]=value;
      if(Object.keys(row).length)values[fid]=row;else delete values[fid];
      const next={fields:fidCustom.fields,values};GM_setValue(FID_CUSTOM_COLUMNS_KEY,JSON.stringify(next));fidCustom=next;
    }else{
      const note=input.value,next=readFidNotes();
      if((next[fid]??'')!==input.defaultValue)throw new Error('The saved note changed. Copy your draft and reopen the note.');
      if(note.trim())next[fid]=note;else delete next[fid];
      GM_setValue(FID_NOTES_KEY,JSON.stringify(next));fidNotes=next;
    }
  }catch(error){const status=form.querySelector('[role=status]');status.classList.add('error');status.textContent='Could not save. '+error.message;return}
  await refreshPersonalRows('fid',new Set([fid]));if(form.isConnected)closeFidNote(form);
}
function fidQuarterRows(node){return [...node.quarterStates].sort((a,b)=>Number(b.quarter===node.currentQuarter)-Number(a.quarter===node.currentQuarter))}
function closeFidDateHistory(){
  if(!fidDatePopup.trigger)return;
  if(fidDatePopup.matches(':popover-open'))fidDatePopup.hidePopover();
  fidDatePopup.trigger=null;fidDatePopup.selection=null;
}
function openFidDateHistory(trigger){
  if(fidDatePopup.trigger===trigger&&fidDatePopup.matches(':popover-open'))return;
  closeFidDateHistory();
  const {fidDate:fid,dateChange:key}=trigger.dataset,dates=fidDateChanges.get(fid)?.[key];
  if(!dates)return;
  const node=graph.fids.get(fid),unread=hasUnreadFidDates(fid,key,dates);
  const host=trigger.closest('dialog')||document.body;
  if(fidDatePopup.parentElement!==host)host.append(fidDatePopup);
  fidDatePopup.trigger=trigger;
  fidDatePopup.selection={fid,key,dates};
  const rows=fidQuarterRows(node).filter(state=>state.dates[key]).map(state=>{
    const quarter=(state.quarter||'[Unknown]')+(state.quarter===node.currentQuarter?' · Current':'');
    return `<li title="${escapeHtml(state.dateSource)}"><strong>${escapeHtml(state.dates[key])}</strong><small>${escapeHtml(quarter)}</small></li>`;
  }).join('');
  fidDatePopup.innerHTML=`<h3 id="fidDatePopupTitle">${fidDateLabels[key]} Recorded Dates</h3><ol>${rows}</ol>${unread?'<button type="button" class="btn small light" data-read-fid-dates>Mark as Read</button>':''}<p role="status"></p>`;
}
async function markFidDatesRead(button){
  const selection=fidDatePopup.selection;if(!selection)return;
  const {fid,key,dates}=selection;
  button.disabled=true;
  try{
    const saved=await nsrData.updateDataset(FID_DATE_READ_KEY,entry=>{
      const readDates=readFidDateMarks(entry),fields=readDates.get(fid)||{};
      fields[key]=[...new Set([...(fields[key]||[]),...dates])];
      readDates.set(fid,fields);
      return {...entry,key:FID_DATE_READ_KEY,readDates};
    });
    fidDateReads.set(fid,saved.readDates.get(fid));
    const unread=hasUnreadFidDates(fid,key,fidDateChanges.get(fid)?.[key]);
    document.querySelectorAll(`[data-fid-date="${CSS.escape(fid)}"][data-date-change="${key}"]`).forEach(trigger=>{
      trigger.classList.toggle('is-unread',unread>0);
      trigger.setAttribute('aria-label',fidDateHistoryLabel(key,trigger.textContent,unread));
    });
    if(fidDatePopup.selection===selection)closeFidDateHistory();
  }catch{
    if(fidDatePopup.selection===selection){button.disabled=false;fidDatePopup.querySelector('[role=status]').textContent='Could not mark as read.'}
  }
}
function renderFidEvidence(node,state){
  const riskLabels={'On time':'Stable Slot','Pre late risk':'Late Slot','Post config risk':'Late CTL','No post risk':'Stable CTL'};
  return [['Pre-slot',state.pre],['Post-slot',state.post]].map(([name,risk])=>{
    const post=name==='Post-slot',ctls=post?risk.assessments:[],label=post&&risk.ctls.length?'Late CTL ('+risk.ctls.length+')':riskLabels[risk.label]||risk.label.replace(/^(?:Pre|Post) /i,'');
    let content='<p>'+escapeHtml(risk.reason)+'</p>';
    if(post&&ctls.length)content='<ul class="fid-ctl-list">'+ctls.map(({ctl,weeks,referenceDate,referenceField})=>{
      const timing=weeks===null?'— weeks (date missing)':weeks.toFixed(1)+' weeks ('+(weeks<20?'<20':'>=20')+' weeks)';
      const tip='Sales Ops Approval Date: '+ctl.approved+' · '+(referenceField||'MCSD / SRD')+': '+(referenceDate||'—');
      return '<li'+(risk.ctlIds.includes(ctl.id)?' class="is-late"':'')+' title="'+escapeHtml(tip)+'">'+escapeHtml(ctl.number+' · Config · '+timing)+'</li>';
    }).join('')+'</ul>';
    else if(!post&&risk.ctlId){
      const timing=risk.weeks===null?'— weeks (date missing)':risk.weeks.toFixed(1)+' weeks ('+(risk.weeks>risk.threshold?'>':'<=')+risk.threshold+' weeks)';
      const tip=(risk.eventDateField||'Sales Ops Approval Date')+': '+(risk.eventDate||'—')+' · '+(risk.referenceField||'SRD')+': '+(risk.referenceDate||'—');
      content='<ul class="fid-ctl-list"><li'+(risk.state==='risk'?' class="is-late"':'')+' title="'+escapeHtml(tip)+'">'+escapeHtml(risk.ctlId+' · Slot · '+timing)+'</li></ul>';
    }
    return '<div class="notice '+(!post&&state.portfolioStates.includes('no-slot')?'no-slot':risk.state)+'"><b>'+escapeHtml(name+(node.quarterStates.length>1?' · '+(state.quarter||'[Unknown]'):'')+': '+label)+'</b>'+content+'</div>';
  }).join('');
}
function clearFidLocationPulses(root){
  for(const animation of root.getAnimations({subtree:true}))if(animation.id==='fid-location')animation.cancel();
}
function pulseFidLocation(element){
  clearFidLocationPulses(element);
  const animation=element.animate([{},{backgroundColor:'#fff0ad'},{}],{duration:1000,iterations:Infinity,easing:'ease-in-out'});
  animation.id='fid-location';
  const dismiss=()=>{
    element.removeEventListener('pointerenter',dismiss);
    animation.effect.updateTiming({iterations:Math.max(3,Math.ceil(Number(animation.currentTime||0)/1000))});
  };
  animation.onfinish=animation.oncancel=()=>element.removeEventListener('pointerenter',dismiss);
  element.addEventListener('pointerenter',dismiss,{once:true});
  if(element.matches(':hover'))dismiss();
}
function showFid(fid,updateMarks=false,target=null){
  const node=graph.fids.get(fid);if(!node)return;const dialog=id('fidDetail');if(dialog.dataset.fid!==fid)['fidNsrTable','fidCtlTable','fidNceTable'].forEach(name=>localTables.delete(name));
  if(dialog.contains(fidDatePopup))closeFidDateHistory();
  const viewed=dialog.open&&dialog.dataset.fid===fid?dialog.dataset.quarter:node.currentQuarter;
  const quarterState=node.quarterStates.find(state=>state.quarter===viewed)||node.quarterStates.find(state=>state.quarter===node.currentQuarter);
  const updateDetail=updateMarks&&dialog.dataset.fid===fid&&tables.has('fidNsrTable');dialog.dataset.fid=fid;dialog.dataset.quarter=quarterState.quarter;
  if(!updateDetail){
    clearFidLocationPulses(dialog);
    if(!dialog.open||dialog.dataset.lookupId!==target?.lookupId)delete dialog.dataset.workflowLookup;
    if(target?.lookupId)dialog.dataset.lookupId=target.lookupId;else delete dialog.dataset.lookupId;
  }
  const lateCtlIds=new Set(node.lateCtls.map(ctl=>ctl.id)),lateSlotIds=new Set(node.quarterStates.flatMap(state=>state.slotEvidence.filter(pre=>pre.state==='risk').map(pre=>pre.ctlId)));
  const quarters=fidQuarterRows(node),multiple=quarters.length>1;
  const quarterLines=quarters.map(state=>{
    const tag=multiple?'button':'div';
    return `<${tag}${multiple?' type="button"':''} class="fid-quarter-state${state.quarter===node.currentQuarter?' is-current':''}" data-quarter-state="${escapeHtml(state.quarter)}"${multiple?` data-selected-quarter="${state===quarterState}" aria-pressed="${state===quarterState}" aria-controls="fidEvidence"`:''}><span>${escapeHtml(state.quarter||'[Unknown]')}</span>${pill(quarterStatus(state,true),state.portfolioStates.includes('risk-fid')?'risk':state.portfolioStates.includes('risk-free')?'clear':'unknown')}</${tag}>`;
  }).join('');
  const activeCtl=node.activeCtl,activeCtlText=activeCtlLabel(activeCtl);
  const summary=`<div class="fid-dialog-grid" style="--fid-date-rows:${quarters.length+1}">${[
    ['Slot',`${node.slot.state.replace(/^./,letter=>letter.toUpperCase())} | ${node.commit||'—'}`],['Active CTL/Quote',activeCtlText||activeCtl?.quote?`${activeCtlText||'—'} | ${activeCtl.quote||'—'}`:'—'],['Account',node.account],['Business Unit',node.businessUnit],
    ['System Description',node.system],['FID Description',node.fidDescription],['Customer',node.customer],['Fab Name',node.fab],
    ['Assessment Quarter',node.quarter,'quarter'],['Sales Ops RD',node.rd||sourceValue(node.record,'fcid','rd'),'rd'],['MCSD',node.mcsd||sourceValue(node.record,'fcid','mcsd'),'mcsd'],['Customer Request Date',node.crd||sourceValue(node.record,'fcid','crd'),'crd'],
    ['MFG Upgrade Type',node.mfg],['FID Owner',node.owner||'[Unknown]'],['Slot Initiator',slotInitiator(node)],['Sales Admin',node.salesAdmin||'[Unknown]']
  ].map(([label,value,dateKey])=>{
    let dates=escapeHtml(value||'—');
    if(dateKey==='rd'||dateKey==='mcsd'){
      if(multiple)dates=quarters.map(state=>{
        const isCurrent=state.quarter===node.currentQuarter,date=fidDateValue(node,state,dateKey)||'—';
        const tip=(state.quarter||'[Unknown]')+' · '+state.dateSource+(!state.dates[dateKey]?'\nDate unavailable.':'');
        return `<div class="fid-quarter-state${isCurrent?' is-current':''}" data-quarter-date="${escapeHtml(state.quarter)}" title="${escapeHtml(tip)}"><span>${isCurrent?renderFidDate(node,dateKey,date):escapeHtml(date)}</span></div>`;
      }).join('');
      else dates=renderFidDate(node,dateKey,value);
    }else if(multiple&&dateKey==='crd')dates=`<div class="fid-quarter-state"><span>${dates}</span></div>`;
    return `<div class="fid-detail-value${multiple&&dateKey?' fid-date-value':''}${label==='Slot'&&node.portfolioStates.includes('no-slot')?' no-slot':''}"${dateKey?` data-date-field="${dateKey}"`:''}${label==='FID Owner'?` title="${escapeHtml(ownerTip(node))}"`:label.startsWith('Sales Admin')?' title="Unique Pull / Push CTL votes; ties remain unknown"':''}><small>${label}${dateKey==='quarter'?helpIcon(quarterHint):label==='FID Owner'?helpIcon(ownerHint):label==='Sales Admin'?helpIcon(salesAdminHint):''}</small>${dateKey==='quarter'?quarterLines:dates}</div>`;
  }).join('')}</div><div id="fidEvidence" class="fid-evidence">${renderFidEvidence(node,quarterState)}</div>`;
  const quoteMatch=target?.quoteId?' · <span id="fidLookupQuote" role="status">'+escapeHtml((target.sourceQuote?'Source Quote ':'Quote ID ')+target.quoteId+(target.ctlId?'':' · FID Database'))+'</span>':'';
  if(updateDetail)id('fidSummary').innerHTML=summary;
  else dialog.innerHTML=`<div class="section-head"><div><h1 id="fidDetailTitle"><span id="fidPersonalMark" class="fid-personal-mark"></span>FID ${escapeHtml(fid)} <a class="fid-action" href="https://pcazcrpap02.fremont.lamrc.net:8443/sap(bD1lbiZjPTI1MCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm?saprole=ZSLSDIR&amp;crm-object-type=BT111_OPPT&amp;crm-object-action=B&amp;crm-object-keyname=OBJECT_ID&amp;crm-object-value=${encodeURIComponent(fid)}" target="_blank" rel="noopener noreferrer">CRM ↗</a><span id="fidHeaderStatus" class="fid-header-status" role="group" aria-label="FID Status &amp; Risks"></span></h1><span class="muted">${node.inPopulation?'Commit portfolio':'Observed in source data'}${quoteMatch}</span></div><a class="fid-action" href="#" data-export-fid="${escapeHtml(fid)}">Export FID 360°</a><button class="fid-close" type="button" data-close-fid aria-label="Close FID 360" title="Close">×</button></div><div id="fidSummary">${summary}</div><div id="fidCustomSummary" class="fid-custom-summary" role="group" aria-label="FID Custom Columns" hidden></div><section id="fidNotePanel" class="fid-note-panel"></section><div id="fidNsrTable"></div><div id="fidCtlTable"></div><div id="fidNceTable"></div>`;
  id('fidHeaderStatus').innerHTML=fidDetailStatus(node,false,true).map(([label,kind,tip])=>pill(label,kind,tip)).join('');
  paintFidPersonal(fid);
  const nsrHistory=renderNsrDetail('fidNsrTable',node.nsrs,{title:'📋 NSR History',file:'FID_'+fid+'_NSR',defaultSort:'Submit Date',rowRisk:nsr=>nsr.eligible&&nsr.lifecycle==='open'?'pending-nsr':''},updateDetail);
  const ctlHistory=updateDetail?tables.get('fidCtlTable').built:table('fidCtlTable',ctlColumns.filter(([label])=>label!=='FID').map(column=>column[0]==='CTL'?[
    'CTL',ctl=>ctl.id===activeCtl?.id&&activeCtlText?activeCtlText:column[1](ctl),
    (value,ctl)=>column[2](value,ctl)+(ctl.id===activeCtl?.id&&ctl.status==='Completed'?'<span class="active-ctl-dot" title="Active CTL" aria-label="Active CTL"></span>':''),
    {sorter:compareCtlNumbers}
  ]:column),node.ctls,{title:'🔄 Slot / CTL History',file:'FID_'+fid+'_CTL',rowRisk:ctl=>lateSlotIds.has(ctl.id)?'late-slot':lateCtlIds.has(ctl.id)?'late-ctl':ctl.isDrop&&ctl.effective?'drop-slot':ctl.pending?'pending-config':ctl.pendingPullPush?'pending-pull-push':'',widths:{CTL:105,Stage:140,Change:90,'Created Date':100,'Elapsed Days':110,'Sales Ops Approval':100,'RD At Create':100,'MCSD@Create':100,'SRD @ SOPs Approval':100,'MCSD At SOps Approval':100,'Config Frz':100,Initiator:135,Customer:160,'Quote ID':105,'NSR Reference':150,'Root Cause':180,'CTL Notes':220,'Reason for Chg':160,'Sub Reason':160,'Sales OPs Approver':170}});
  if(!updateDetail)table('fidNceTable',nceColumns,node.nces,{title:'🔎 NCE / FA History',file:'FID_'+fid+'_NCE',widths:{NCE:90,State:80,Title:140,Impact:75,Opened:95,'Damage Group':95,'Damage Code':90,'Part Number':105,'Part Description':140}});
  if(!dialog.contains(workNotes))dialog.append(workNotes);
  if(!dialog.open)dialog.showModal();
  if(target?.ctlId||target?.nsrId)(target.ctlId?ctlHistory:nsrHistory).then(state=>{
    if(!dialog.open||dialog.dataset.fid!==fid||tables.get(state.name)!==state)return;
    const index=state.rows.findIndex(record=>record.id===(target.ctlId||target.nsrId)),record=state.rows[index];
    if(target.ctlId&&record&&['pull','push'].includes(record.changeKey)&&!state.local.showPullPush)id('fidCtlTable').querySelector('[data-toggle-pull-push]').click();
    const row=index>=0?state.grid.getRow(index)?.getElement():null;
    (row||state.toolbar).scrollIntoView({block:'center',inline:'nearest'});
    if(row)pulseFidLocation(row);
    // Source updates can retarget this history. Open Workflow only once for the user's lookup.
    const workflow=target.nsrId&&target.lookupId&&dialog.dataset.workflowLookup!==target.lookupId?row?.querySelector('[data-wf-nsr]'):null;
    if(workflow){dialog.dataset.workflowLookup=target.lookupId;workflow.click()}
  });
  else if(!updateDetail&&quoteMatch)pulseFidLocation(id('fidLookupQuote'));
}
function exportFid(fid){
  const original=graph.fids.get(fid),node=graph.projectFid(original,[id('fidDetail').dataset.quarter||'[Unknown]'])||original,summary=markedDetailRows('fid',[node],fidColumns())[0];
  const quarters=fidQuarterRows(original);
  summary['FID Status & Risks']=fidDetailStatus(original,false,true).map(([label])=>label).join(' · ');
  summary['Assessment quarters']=node.quarter;summary.Quarter=quarters.map(state=>state.quarter||'[Unknown]').join('\n');summary['Quarter assessments']=quarters.map(state=>(state.quarter||'[Unknown]')+': '+quarterStatus(state)).join('\n');
  if(quarters.length>1)for(const [label,key]of [['Sales Ops RD','rd'],['MCSD','mcsd']])summary[label]=quarters.map(state=>fidDateValue(original,state,key)||'—').join('\n');
  summary['Fab Name']=node.fab;summary['Business Unit']=node.businessUnit;summary['Sales Admin']=node.salesAdmin;summary['Owner rule']='Slot-derived Account membership → Slot / Config CTL counts → tied candidates\' NSR counts (all eligible CTL counts are 0: NSR fallback within Account) → current Slot submitter';
  summary['Pending NSR']=node.post.nsrs.filter(nsr=>includeIgnored||!nsr.ignored).length;summary['Completed 5D NSR']=node.nsrs.filter(nsr=>nsr.eligible&&nsr.lifecycle==='completed-5d').length;summary['Completed NSR but no CTL']=node.nsrs.filter(nsr=>nsr.eligible&&nsr.lifecycle==='completed-unlinked').length;
  exportWorkbook('FID_'+fid,{FID:[summary],NSR:exportTableRows(tables.get('fidNsrTable')),NSR_Source:node.nsrs.flatMap(nsr=>nsr.rows||[]),CTL:node.ctls.flatMap(ctl=>{const days=ctlElapsedDays(ctl);return (ctl.rows||ctl.rawRows||[]).map(row=>({...row,'Elapsed Days':days}))}),NCE:node.nces.flatMap(nce=>nce.rows||nce.rawRows||[]),FID_Source:node.rawRows||[node.record]})
}
function exportSettings(){
  try{
    const settings={format:'ToolkitKyra Settings',version:1,nsrMarks:currentMarks('nsr'),fidMarks:currentMarks('fid'),fidNotes,accountMapping:accountMapping.getMap(),fidCustomColumns:fidCustom};
    const url=URL.createObjectURL(new Blob([JSON.stringify(settings,null,2)],{type:'application/json;charset=utf-8'})),anchor=document.createElement('a');
    anchor.href=url;anchor.download='ToolkitKyra_Settings_'+nsrData.localDate(new Date())+'.json';document.body.append(anchor);anchor.click();anchor.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
  }catch(error){toast('Could not export settings: '+error.message)}
}
function appendFidCustomColumns(incoming){
  const local=readFidCustomColumns(),fields=[...local.fields],values={...local.values};
  const byId=new Map(fields.map(field=>[field.id,field])),byTitle=new Map(fields.map(field=>[field.title.toLowerCase(),field]));
  const destinations=new Map(),claimed=new Set();
  // Stable IDs win over independently created fields with a matching title.
  const ordered=[...incoming.fields].sort((a,b)=>Number(byId.has(b.id))-Number(byId.has(a.id)));
  for(const field of ordered){
    const existing=byId.get(field.id)||byTitle.get(field.title.toLowerCase()),target=existing||field;
    if(target.type!==field.type||claimed.has(target.id))continue;
    if(!existing)fields.push(field);
    destinations.set(field.id,target.id);claimed.add(target.id);
  }
  for(const [fid,imported] of Object.entries(incoming.values)){
    const row={...values[fid]};
    for(const [field,value] of Object.entries(imported)){
      const target=destinations.get(field);
      if(target&&!Object.hasOwn(row,target))row[target]=value;
    }
    if(Object.keys(row).length)values[fid]=row;
  }
  return {fields,values};
}
let settingsImporting=false;
async function importSettings(file){
  if(settingsImporting)return;
  settingsImporting=true;
  try{
    const settings=JSON.parse(await file.text());
    if(!isRecord(settings)||settings.format!=='ToolkitKyra Settings'||settings.version!==1||!['nsrMarks','fidMarks','fidNotes','accountMapping'].every(name=>isRecord(settings[name])))throw new Error('Invalid ToolkitKyra settings file.');
    const marks={nsr:settings.nsrMarks,fid:settings.fidMarks};
    for(const kind of ['nsr','fid'])if(Object.entries(marks[kind]).some(([number,mark])=>!markKinds[kind].pattern.test(number)||!['ignore','important'].includes(mark)))throw new Error('Invalid '+kind.toUpperCase()+' marks.');
    if(Object.entries(settings.fidNotes).some(([fid,note])=>!markKinds.fid.pattern.test(fid)||typeof note!=='string'))throw new Error('Invalid FID notes.');
    let notes=normalizeFidNotes(settings.fidNotes),mapping=accountMapping.normalize(settings.accountMapping);
    let custom=Object.hasOwn(settings,'fidCustomColumns')?normalizeFidCustomColumns(settings.fidCustomColumns):null;
    const dialog=id('fcSettingsImportDialog');id('fcSettingsImportFileName').textContent=file.name;dialog.returnValue='cancel';
    const mode=await new Promise(resolve=>{
      dialog.addEventListener('close',()=>resolve(dialog.returnValue),{once:true});dialog.showModal();
    });
    if(mode!=='append'&&mode!=='replace')return;
    if(mode==='append'){
      for(const kind of ['nsr','fid'])marks[kind]={...marks[kind],...markSettings(readSavedMarkSets(kind))};
      notes={...notes,...readFidNotes()};mapping=accountMapping.append(mapping);
      if(custom)custom=appendFidCustomColumns(custom);
    }
    const values={
      [NSR_MARKS_KEY]:JSON.stringify(marks.nsr),[FID_MARKS_KEY]:JSON.stringify(marks.fid),[FID_NOTES_KEY]:JSON.stringify(notes),
      [accountMapping.key]:JSON.stringify({base:accountMapping.defaults,mapping}),...(custom?{[FID_CUSTOM_COLUMNS_KEY]:JSON.stringify(custom)}:{})
    };
    const previous=Object.fromEntries(Object.keys(values).map(key=>[key,GM_getValue(key,undefined)])),written=[];
    try{for(const [key,value]of Object.entries(values)){GM_setValue(key,value);written.push(key)}}
    catch(error){
      let rollbackFailed=false;
      for(const key of written.reverse())try{
        if(previous[key]===undefined)GM_deleteValue(key);else GM_setValue(key,previous[key]);
      }catch{rollbackFailed=true}
      if(rollbackFailed)throw new Error(error.message+' Some settings could not be restored.');
      throw error;
    }
    for(const kind of ['nsr','fid']){saveMarks(kind,marks[kind],false);markDrafts[kind]=null}
    if(custom){fidCustom=custom;fidCustomDraft=null}
    fidNotes=notes;mappingDraft=null;accountMapping.reload();if(await recalculate())toast(mode==='append'?'Settings appended':'Settings imported');
  }catch(error){toast('Could not import settings: '+error.message)}
  finally{settingsImporting=false}
}
function showFidEquation(event){
  const target=event.target.closest('[data-fid-status],[data-fid-cell],[data-fid-region],[data-fid-axis]');if(!target)return;
  if(target.contains(event.relatedTarget))return;
  const leaving=event.type==='pointerout'||event.type==='focusout',key=target.dataset.fidStatus||target.dataset.fidRegion;
  for(const element of id('view-portfolio').querySelectorAll('.fid-relation-sign.is-visible'))element.classList.remove('is-visible');
  if(leaving)return;
  let equations=[];
  if(target.dataset.fidCell)equations=['post:'+target.dataset.matrixPost];
  else if(target.dataset.fidAxis)equations=['origin-'+target.dataset.fidAxis];
  else if(key==='slot-ready')equations=['origin-pre','origin-post'];
  else if(['stable-slot','late-slot','slot-date-unknown'].includes(key))equations=['slot:'+key];
  else if(['stable-ctl','late-ctl-nsr','ctl-unknown'].includes(key))equations=['post:'+key];
  else if(key==='risk-fid')equations=[key];
  for(const element of id('view-portfolio').querySelectorAll('[data-fid-equation]'))element.classList.toggle('is-visible',element.dataset.fidEquation.split(' ').some(equation=>equations.includes(equation)));
}
for(const name of ['pointerover','pointerout','focusin','focusout'])id('view-portfolio').addEventListener(name,showFidEquation);
const helpTip=document.createElement('div');helpTip.className='help-tip';helpTip.setAttribute('role','tooltip');helpTip.popover='manual';document.body.append(helpTip);
document.addEventListener('pointerover',event=>{
  const help=event.target.closest('.help-icon');if(!help)return;
  helpTip.textContent=help.getAttribute('aria-label');helpTip.showPopover();
  const rect=help.getBoundingClientRect();
  helpTip.style.left=Math.max(8,Math.min(rect.left,innerWidth-helpTip.offsetWidth-8))+'px';
  helpTip.style.top=Math.max(8,Math.min(rect.bottom+7,innerHeight-helpTip.offsetHeight-8))+'px';
});
document.addEventListener('pointerout',event=>{if(event.target.closest('.help-icon'))helpTip.hidePopover()});
// Keep Chosen focus from scrolling an option away between mouse down and mouse up.
document.addEventListener('mousedown',event=>{const option=event.target.closest('.side .chosen-results .active-result');if(event.button!==0||!option)return;option.closest('.chosen-container').querySelector('.chosen-search-input')?.focus({preventScroll:true});event.preventDefault()},true);
document.querySelector('.filter-presets').addEventListener('contextmenu',event=>{
  const button=event.target.closest('[data-filter-preset]');if(!button)return;
  event.preventDefault();openFilterPreset(button,true);
});
document.addEventListener('click',async event=>{
  const label=event.target.closest('.fc-filter-head>label[for]');if(label&&id(label.htmlFor)?.multiple&&jQuery.fn.chosen){event.preventDefault();jQuery(id(label.htmlFor)).trigger('chosen:open');return}
  closeSourceMenus(event.target);const button=event.target.closest('button,a[data-export-fid]');if(!button)return;
  if(button.dataset.filterPreset){openFilterPreset(button);return}
  if(button.dataset.presetAnswer){confirmFilterPreset(button.dataset.presetAnswer);return}
  if(button.id==='fcToggleSide'){
    const expanded=button.getAttribute('aria-expanded')!=='true';button.setAttribute('aria-expanded',String(expanded));
    if(sidebarAnimation){sidebarAnimation.reverse();return}
    if(expanded)document.body.classList.remove('side-collapsed');
    const finish=()=>{document.body.classList.toggle('side-collapsed',button.getAttribute('aria-expanded')==='false');sidebarAnimation=null};
    if(matchMedia('(prefers-reduced-motion: reduce)').matches)finish();
    else{sidebarAnimation=document.querySelector('.side').animate([{opacity:1,transform:'translateX(0)'},{opacity:0,transform:'translateX(-8px)'}],{duration:150,easing:'ease-out',direction:expanded?'reverse':'normal'});sidebarAnimation.onfinish=finish}
    return;
  }
  if(button.dataset.locateFilter){
    if(id('fcToggleSide').getAttribute('aria-expanded')==='false')id('fcToggleSide').click();
    let target=id(button.dataset.locateFilter).closest('.fc-filter');if(target.hidden)target=document.querySelector('.filter-head');
    target.scrollIntoView({block:'center',behavior:'instant'});target.tabIndex=-1;target.focus({preventScroll:true});
    target.animate([{backgroundColor:'#fff0ad',boxShadow:'0 0 0 3px #efb647'},{}],{duration:matchMedia('(prefers-reduced-motion: reduce)').matches?0:600,iterations:5,easing:'ease-out'});return;
  }
  if(button.dataset.importSource){id('fcFile-'+button.dataset.importSource).click();return}
  if(button.dataset.sourceAction){const action=button.dataset.sourceAction;if(action==='bw-update'||action==='ctl-bw-update'){const key=action==='bw-update'?'fcid':'ctl';openBwLoadDialog(bwData,key,store.meta[key],range=>loadBw(key,false,range))}else if(action==='nsr-open'){setNsrRange(nsrRange);id('fcNsrDialog').showModal()}return}
  if(button.dataset.fidDate){openFidDateHistory(button);return}
  if(button.hasAttribute('data-read-fid-dates')){await markFidDatesRead(button);return}
  if(button.dataset.noteFid){editFidNote(button.dataset.noteFid,button);return}
  if(button.dataset.fidCustom){const field=fidCustom.fields.find(field=>field.id===button.dataset.fidCustom&&field.display);if(field)editFidNote(button.dataset.customFid,button,field);return}
  if(button.hasAttribute('data-cancel-fid-note')){closeFidNote(button.form);return}
  if(button.dataset.ignoreNsr||button.dataset.ignoreFid){setRecordMark(button.dataset.ignoreFid?'fid':'nsr',button.dataset.ignoreFid||button.dataset.ignoreNsr);return}
  if(button.dataset.view){showView(button.dataset.view);return}
  if(button.hasAttribute('data-fid')){if(graph.fids.has(button.dataset.fid)){if(button.closest('#wfModal')&&id('fidDetail').open)id('fidDetail').close();showFid(button.dataset.fid,false,button.dataset)}else toast('No FID is available'+(button.dataset.nsrId?' for this NSR.':button.dataset.ctlId?' for this CTL.':'.'));return}
  if(button.hasAttribute('data-mapping-rule')){showView('rules');id('fabAccountMapping').scrollIntoView({block:'start'});return}
  if(button.hasAttribute('data-quarter-state')){
    const dialog=id('fidDetail'),node=graph.fids.get(dialog.dataset.fid),state=node?.quarterStates.find(state=>state.quarter===button.dataset.quarterState);if(!state)return;
    dialog.dataset.quarter=state.quarter;
    dialog.querySelectorAll('[data-quarter-state]').forEach(row=>{const selected=row===button;row.dataset.selectedQuarter=String(selected);row.setAttribute('aria-pressed',String(selected))});
    id('fidEvidence').innerHTML=renderFidEvidence(node,state);return;
  }
  if(button.hasAttribute('data-close-fid')){id('fidDetail').close();return}
  if(button.hasAttribute('data-toggle-pull-push')){
    const state=tables.get('fidCtlTable');state.local.showPullPush=!state.local.showPullPush;button.textContent=(state.local.showPullPush?'Hide':'Show')+' Pull/Push';button.setAttribute('aria-pressed',String(state.local.showPullPush));paintTable(state);
    // Filtering already rendered the new height; a second resize layout would reset column widths.
    const wrap=state.grid.element;state.size=wrap.clientWidth+':'+wrap.clientHeight;return;
  }
  if(button.dataset.exportFid){event.preventDefault();exportFid(button.dataset.exportFid);return}
  for(const [path,reset] of Object.entries(filterResets))if(button.hasAttribute(reset.attribute)){changeFilter(path,reset.values());return}
  if(button.dataset.clearFilter||button.dataset.selectAll){const name=button.dataset.clearFilter||button.dataset.selectAll,definition=filterDefs.find(def=>def[0]===name);const values=button.dataset.selectAll?[...id(name).options].map(option=>option.value).filter(value=>name!=='fNsrTaskStatus'||!Object.hasOwn(completedTasks,value)):[];changeFilter(definition[2],values);return}
  if(button.dataset.fidCell||button.dataset.fidStatus||button.dataset.fidRegion){
    const key=button.dataset.fidStatus||button.dataset.fidRegion;
    if(key==='all'){changeFilter({fidStatus:[],signal:(filterValues.signal||[]).filter(value=>!fidCauseKeys.includes(value))});return}
    const value=button.dataset.fidValue||button.dataset.fidCell,parts=fidSelection.groups[value]||[value],selected=filterValues.fidStatus||[];
    changeFilter('fidStatus',fidSelection.groups[value]?fidStatusSelected(value)?[]:parts:parts.every(part=>selected.includes(part))?selected.filter(item=>!parts.includes(item)):[...selected,...parts]);return;
  }
  if(button.dataset.ctlStatus){selectTaskFilter('ctl',button.dataset.ctlStatus,button.dataset.ctlTypes?.split(','));return}
  if(button.dataset.nsrGroup||button.dataset.nsrLifecycle){const key=button.dataset.nsrLifecycle||button.dataset.nsrGroup;selectTaskFilter('nsr',key==='open'?'workflow':key);return}
  if(button.dataset.ctlPeriod){changeFilter('ctl.created',ctlDateRange(button.dataset.ctlPeriod));return}
  if(button.dataset.nsrYear){const year=button.dataset.nsrYear;changeFilter('nsr.submitted',[year+'-01-01',year+'-12-31']);return}
  if(button.dataset.fidYear){const values=quartersForYear(button.dataset.fidYear),selected=filterValues.quarter||[],add=values.some(value=>!selected.includes(value));changeFilter('quarter',add?[...new Set([...selected,...values])]:selected.filter(value=>!values.includes(value)));return}
  if(button.dataset.groupPath){const path=button.dataset.groupPath,value=button.dataset.groupValue,selected=filterValues[path]||[];changeFilter(path,selected.includes(value)?selected.filter(item=>item!==value):[...selected,value]);return}
  if(button.dataset.pivotFilter){const path=filterDefs.find(def=>def[0]===button.dataset.pivotFilter)[2];changeFilter(path,[button.dataset.pivotValue]);return}

  if(button.dataset.tableExpand){toggleTableExpanded(button);return}
  if(button.dataset.tableCycle){const state=tables.get(button.dataset.tableCycle),rows=tableResult(state);if(id('fidDetail').open)id('fidDetail').close();showView('cycle');cycle.query(rows.map(row=>row.number));return}
  if(button.dataset.columnClear){tables.get(button.dataset.columnClear).grid.clearHeaderFilter();if(hasPendingMarks())applyFilters(true);return}
  if(button.dataset.columnRemove){tables.get(button.dataset.columnRemove).grid.setHeaderFilterValue(button.dataset.field,'');if(hasPendingMarks())applyFilters(true);return}
  if(button.dataset.tableCopy){const state=tables.get(button.dataset.tableCopy);if(state.options.onCopy)state.options.onCopy();else copyRows(exportTableRows(state,true));return}
  if(button.dataset.summarySort){
    const key=button.dataset.summarySort;summarySort={key,dir:summarySort.key===key?-summarySort.dir:-1};renderSummary(id('view-summary'));
    id('summaryChart').querySelector('[data-summary-sort="'+key+'"]').focus({preventScroll:true});return;
  }
  if(button.dataset.summaryExport){
    const percentMetrics=metricDefs.filter(([,,base])=>base).map(([label])=>label);
    if(button.dataset.summaryExport==='copy')copyRows(summaryReport.map(([label,...values])=>[label,...values.map(value=>typeof value==='number'&&percentMetrics.includes(label)?(100*value).toFixed(1)+'%':value)]));
    else exportWorkbook('Summary_Metrics_'+summaryGroup,{Metrics:summaryReport},Object.fromEntries(percentMetrics.map(label=>[label,'0.0%'])));
    return;
  }
  if(button.dataset.tableExport){const state=tables.get(button.dataset.tableExport);if(state.options.onExport)state.options.onExport();else exportWorkbook(state.options.file,{Detail:exportTableRows(state),...(state.options.cycle?{Source_Rows:tableResult(state).flatMap(nsr=>nsr.rows||nsr.rawRows||[])}:{})});return}
  if(button.hasAttribute('data-clear-search')){linkedId='';id('fSearch').value='';applyFilters();return}
  if(button.hasAttribute('data-clear-signal')){delete filterValues.signal;applyFilters();return}
  if(button.hasAttribute('data-toast-close')){id('fcToast').hidden=true;return}
  if(button.hasAttribute('data-save-nsr-marks')||button.hasAttribute('data-save-fid-marks')){saveMarkLists(button.hasAttribute('data-save-fid-marks')?'fid':'nsr');return}
  if(button.hasAttribute('data-save-mapping')){const value=Object.fromEntries([...document.querySelectorAll('[data-account-group]')].map(textarea=>[textarea.dataset.accountGroup,textarea.value])),result=accountMapping.save(value);if(result.ok){mappingDraft=null;if(await recalculate())toast('Account mapping saved')}else{const notice=id('agNotice');notice.hidden=false;notice.textContent=result.error}return}
  if(button.hasAttribute('data-reset-mapping')){if(!unsafeWindow.confirm('Restore the default Fab / Account mapping?'))return;const result=accountMapping.reset();if(result.ok){mappingDraft=null;await recalculate()}else toast(result.error);return}
  if(button.hasAttribute('data-import-settings')){if(!settingsImporting)id('fcSettingsFile').click();return}
  if(button.hasAttribute('data-export-settings')){exportSettings();return}
  if(button.hasAttribute('data-add-fid-custom')){
    fidCustomDraft=[...readFidCustomDraft(),{id:crypto.randomUUID(),title:'',type:'boolean',display:true}];
    paintFidCustomSettings();id('fidCustomStatus').textContent='';id('fidCustomRows').lastElementChild.querySelector('input').focus();return;
  }
  if(button.hasAttribute('data-delete-fid-custom')){await deleteFidCustomColumn(button);return}
  if(button.hasAttribute('data-export-data-sources')){exportAllDataSources();return}
  if(button.id==='fcRefresh')restoreSources(true);else if(button.id==='fcReset'){
    includeIgnored=true;pendingFilterValues={commit:defaultCommit};for(const key of Object.keys(filterValues))delete filterValues[key];
    if(store.raw.fcid.length)pendingFilterValues.quarter='rolling';
    if(store.meta.ctl||store.meta.ctlNsr)filterValues['ctl.change']=[...defaultTypes];
    linkedId='';id('fSearch').value='';for(const [name,local] of localTables){local.scroll=0;local.marks=name==='portfolioTable'?defaultFidDetailMarks:null;local.headerFilters=[]}for(const state of tables.values())if(state.ready)state.grid.clearHeaderFilter();buildFilters();applyFilters();
  }
});
document.addEventListener('keydown',event=>{
  if(event.ctrlKey&&event.shiftKey&&event.altKey&&event.code==='KeyD'){
    event.preventDefault();if(event.repeat||document.querySelector('[data-view="summary"]').hidden)return;
    summaryGroup=summaryGroup==='owner'?previousGroup:'owner';rendered.delete('summary');showView('summary');return;
  }
  if((event.key==='Enter'||event.key===' ')&&event.target.matches('.fc-grid .tabulator-col')){event.preventDefault();event.target.click()}
  if(event.key==='Escape')closeSourceMenus();
});
document.addEventListener('input',event=>{
  const kind=event.target.dataset.nsrMarks?'nsr':event.target.dataset.fidMarks?'fid':'';
  if(kind){markDrafts[kind]=Object.fromEntries([...document.querySelectorAll(`[data-${kind}-marks]`)].map(textarea=>[textarea.dataset[kind+'Marks'],textarea.value]));id(markKinds[kind].status).textContent=''}
  if(event.target.dataset.accountGroup){mappingDraft??=accountMapping.getMap();mappingDraft[event.target.dataset.accountGroup]=event.target.value}
  if(event.target.closest('#fidCustomRows')){fidCustomDraft=readFidCustomDraft();id('fidCustomStatus').textContent=''}
});
document.addEventListener('change',event=>{
  const input=event.target;
  if(input.dataset.tableMarkFilter){
    const state=tables.get(input.dataset.tableMarkFilter),marks=input.dataset.mark==='all'?(input.checked?null:[]):[...markPopup.querySelectorAll('input[data-mark]:checked')].map(input=>input.dataset.mark).filter(mark=>mark!=='all');
    state.local.marks=marks?.length===markOptions.length?null:marks;
    if(hasPendingMarks())applyFilters(true);else paintTable(state);
  }else if(input.dataset.excludeIgnored){
    const view=input.dataset.excludeIgnored,focused=document.activeElement===input;includeIgnored=!input.checked;applyFilters(true);
    if(focused)id('view-'+view).querySelector('[data-exclude-ignored]').focus({preventScroll:true});
  }else if(input.id==='fcSettingsFile'&&input.files?.[0]){importSettings(input.files[0]);input.value=''}else if(input.dataset.sourceFile&&input.files?.[0])importSource(input.files[0],input.dataset.sourceFile);
});
filterDefs.forEach(([name,,path])=>{if(dateFilters.includes(path)){id(name).addEventListener('change',event=>{const [from,to]=id(name).querySelectorAll('input');from.setCustomValidity('');to.setCustomValidity('');if(from.value&&to.value&&from.value>to.value){event.target.setCustomValidity('From date must not be later than To date.');event.target.reportValidity();return}changeFilter(path,from.value||to.value?[from.value,to.value]:[])});return}const position=()=>requestAnimationFrame(()=>{
  const chosen=id(name+'_chosen');if(!chosen?.classList.contains('chosen-with-drop'))return;
  const bounds=chosen.closest('.side').getBoundingClientRect(),box=chosen.getBoundingClientRect(),height=chosen.querySelector('.chosen-drop').getBoundingClientRect().height;
  const below=Math.min(innerHeight,bounds.bottom)-box.bottom,above=box.top-Math.max(0,bounds.top);chosen.classList.toggle('chosen-dropup',height>below&&above>below);
});jQuery(id(name)).on('change',(event,change)=>{const values=[...id(name).selectedOptions].map(option=>option.value);if(path==='commit'||path==='mfg'||path==='riskType'||path==='fidStatus'||path==='nsr.taskStatus'||path.startsWith('ctl.'))changeFilter(path,path==='fidStatus'?fidSelection.groups[change?.selected]||values:values);else{delete pendingFilterValues[path];filterValues[path]=values;applyFilters()}position()}).on('chosen:showing_dropdown',position);if(jQuery.fn.chosen)jQuery(id(name)).chosen({width:'100%',search_contains:true,hide_results_on_select:false,placeholder_text_multiple:'All',display_selected_options:false})});
columnPopup.onsubmit=columnPopup.onreset=event=>{
  event.preventDefault();const column=columnPopup.column,value=event.type==='reset'?'':!columnNoteFilter.hidden&&columnNoteFilter.value?{text:columnInput.value,presence:columnNoteFilter.value}:columnInput.value;
  if(column&&JSON.stringify(value)!==JSON.stringify(column.getHeaderFilterValue()??''))column.setHeaderFilterValue(value);
  columnPopup.hidePopover();if(column&&hasPendingMarks())applyFilters(true);
};
columnPendingCtl.onclick=()=>{columnInput.value=pendingCtlMarker;columnPopup.requestSubmit()};
columnInput.oninput=paintColChoices;
columnChoices.onchange=()=>{columnInput.value=columnChoices.value;columnInput.focus({preventScroll:true})};
columnPopup.ontoggle=event=>{if(event.newState==='open')columnInput.select();else{columnPopup.column=null;columnPopup.choices=[];columnChoices.replaceChildren();columnChoices.hidden=true}};
markPopup.onbeforetoggle=event=>{if(event.newState==='closed'){markPopup.state=null;markPopup.trigger=null}};
fidDatePopup.onbeforetoggle=event=>{if(event.newState==='closed'){fidDatePopup.trigger=null;fidDatePopup.selection=null}};
filterPresetPopup.onbeforetoggle=event=>{if(event.newState==='closed')filterPresetPopup.selection=null};
document.addEventListener('scroll',event=>{
  if(columnPopup.column&&!columnPopup.contains(event.target))columnPopup.hidePopover();
  if(markPopup.state&&!markPopup.contains(event.target)&&markPopup.position!==markMenuPos(markPopup.state))markPopup.hidePopover();
  if(fidDatePopup.trigger&&!fidDatePopup.contains(event.target))closeFidDateHistory();
},true);
window.addEventListener('resize',()=>{columnPopup.hidePopover();markPopup.hidePopover();closeFidDateHistory()});
id('fidDetail').addEventListener('close',()=>{
  const dialog=id('fidDetail');if(dialog.open)return;
   if(dialog.contains(workNotes))(expandedTable||document.body).append(workNotes);
  if(dialog.contains(fidDatePopup))closeFidDateHistory();
  clearFidLocationPulses(dialog);id('fidNotePanel')?.querySelector('form')?.remove();highlightHistory('fidCtlTable');highlightHistory('fidNsrTable');
});
id('fidNoteDialog').addEventListener('close',event=>{if(!event.currentTarget.open)event.currentTarget.replaceChildren()});
document.addEventListener('submit',event=>{
  if(event.target.classList.contains('fid-note-form')){event.preventDefault();saveFidNote(event.target,event.submitter)}
  if(event.target.id==='fidCustomSettings'){event.preventDefault();saveFidCustomSettings()}
});
id('fSearch').addEventListener('input',()=>{linkedId='';syncFidLookup();clearTimeout(searchTimer);searchTimer=setTimeout(applyFilters,180)});
id('fidLookup').addEventListener('submit',event=>{
  event.preventDefault();
  if(searchTimer)applyFilters();
  const match=syncFidLookup(),button=event.currentTarget.querySelector('button');
  if(match)showFid(match.fid,false,match.target);
  else if(!button.hidden&&!button.disabled)searchBwDetails([...bwLookupIdentifiers(id('fSearch').value).keys()]);
});
id('fcNsrFrom').addEventListener('input',()=>id('fcNsrTo').min=id('fcNsrFrom').value);id('fcNsrCancel').addEventListener('click',()=>id('fcNsrDialog').close());id('fcNsrForm').addEventListener('submit',event=>{event.preventDefault();if(store.busy.has('nsr'))return;id('fcNsrDialog').close();loadSharedNsr(false)});

function normalizeNsrRows(rows){return rows.map(raw=>{const row=nsrRecords.canonicalRow(raw);row['REF FCID']||=text(raw['Reference FID']??raw.FID);row['Submit Date']=sourceDate(row['Submit Date'],nsrData)||row['Submit Date'];return {...raw,...row,'Reference Fcid':row['REF FCID'],'Workflow Status':raw['Workflow Status']||row['NSR Status']||row['Task Name'],'Task Status':nsrRecords.deriveTaskStatus(row)}}).filter(row=>row['Reference Fcid']||row['NSR#'])}
function sourceRange(rows,key,meta){
  if(key==='nsr'&&sourceDate(meta.from,nsrData)&&sourceDate(meta.to,nsrData))return {dateFrom:meta.from,dateTo:meta.to,dateField:'NSR query range'};
  return sourceDateRange(rows,key,nsrData);
}
function sourceMessage(key,message,kind='ok',transport=false){
  paintSourceStatus(id('fcSource-'+key),store.meta[key],nsrData,message,kind,kind==='ok'&&store.dirty);
  if(BW_SOURCES[key]){const help=id('fcHelp-'+key);help.hidden=!transport;if(help.hidden)help.open=false}
}
function actionMessage(message){id('fcActionStatus').textContent=message}
function sourceMeta(key){paintSourceFile(id('fcMeta-'+key),store.meta[key],!BW_SOURCES[key]);id('fcSourceCount').textContent=Object.keys(store.meta).length+' / 5'}
function sourceBusy(key,on){
  on?store.busy.add(key):store.busy.delete(key); document.querySelectorAll(`[data-source-file="${key}"]`).forEach(input=>input.disabled=on); id('fcImport-'+key).disabled=on;
  const actions=key==='fcid'?['bw-update']:key==='ctl'?['ctl-bw-update']:key==='nsr'?['nsr-open']:[]; actions.forEach(action=>document.querySelectorAll(`[data-source-action="${action}"]`).forEach(button=>button.disabled=on));
  if(key==='nsr')id('fcNsrLoad').disabled=on;
  id('fcRefresh').disabled=store.refreshing||store.busy.size>0;
  if(BW_SOURCES[key])syncFidLookup();
}
function startSource(key){if(BW_SOURCES[key]&&bwLookup?.controller)clearBwLookup();store.controllers[key]?.abort();store.runs[key]=(store.runs[key]||0)+1;sourceBusy(key,true);return store.runs[key];}
function setSourceRows(key,rows,meta){
  if(['ctl','ctlNsr'].includes(key)&&!store.meta.ctl&&!store.meta.ctlNsr&&!pendingFilterValues['ctl.change'])filterValues['ctl.change']=[...defaultTypes];
  if(key==='fcid'&&id('workspaceNav').querySelector('[data-view="portfolio"]').hidden){
    pendingFilterValues.quarter??='rolling';
    id('workspaceNav').querySelectorAll('[data-view]').forEach(button=>button.hidden=false);
    showView('portfolio');
  }
  store.raw[key]=rows;store.meta[key]={...meta,...(meta.dateField?{}:sourceRange(rows,key,meta)),rows:rows.length,savedAt:meta.savedAt||Date.now()};
  dataRevision++;store.dirty=true;sourceMeta(key);
}
async function applySource(key,rows,meta,render=true,run){
  setSourceRows(key,rows,meta);sourceMessage(key,'Loaded');
  if(run!==undefined)await saveSourceCache(key,run);
  if(render&&!store.refreshing&&(run===undefined||run===store.runs[key]))await recalculate(run!==undefined);
}


async function downloadHostedArchive(keys,force){
  const requests=[],timers=new Set();let settled=false,releaseMirrors;
  const mirrors=new Promise(resolve=>releaseMirrors=resolve),hedge=setTimeout(releaseMirrors,3000);
  const attempts=HOSTED_ARCHIVE_URLS.map(async(url,index)=>{
    if(index){await mirrors;if(settled)throw new Error('Data download was cancelled.')}
    try{
      const saved=keys.map(key=>store.meta[key]?.hostedValidator),first=saved[0];
      // A 304 is usable only when every requested source was parsed from this exact response.
      const validator=!force&&first?.url===url&&saved.every(value=>value?.url===url&&value.etag===first.etag&&value.lastModified===first.lastModified)?first:null;
      const headers={'Cache-Control':'no-cache'};
      if(typeof validator?.etag==='string'&&/^(?:W\/)?"[^"\r\n]*"$/.test(validator.etag))headers['If-None-Match']=validator.etag;
      else if(typeof validator?.lastModified==='string'&&!/[\r\n]/.test(validator.lastModified)&&Number.isFinite(Date.parse(validator.lastModified)))headers['If-Modified-Since']=validator.lastModified;
      return await new Promise((resolve,reject)=>{
        let request,done=false;
        const finish=(error,result)=>{if(done)return;done=true;clearTimeout(timer);timers.delete(timer);error?reject(error):resolve(result)};
        // anonymous requests can use Tampermonkey's fetch transport, whose native timeout may be ignored.
        const timer=setTimeout(()=>{finish(new Error('Data download timed out (30 s).'));try{request?.abort?.()}catch{}},30000);timers.add(timer);
        try{
          request=GM_xmlhttpRequest({
            method:'GET',url:url+(force?'?t='+Date.now():''),headers,responseType:'arraybuffer',timeout:30000,nocache:force,revalidate:!force,anonymous:true,
            onload:response=>{
              const unchanged=response.status===304&&Boolean(headers['If-None-Match']||headers['If-Modified-Since']);
              if(unchanged||response.status>=200&&response.status<300&&response.response?.byteLength){
                if(!unchanged){
                  const bytes=new Uint8Array(response.response);
                  if(![137,80,78,71,13,10,26,10].every((value,index)=>bytes[index]===value)){finish(new Error('HTTP '+response.status+' (expected the hosted PNG data).'));return}
                }
                const header=name=>String(response.responseHeaders||'').match(new RegExp('^'+name+':[ \\t]*(.*)$','im'))?.[1].trim()||'';
                finish(null,{bytes:response.response,unchanged,validator:{url,etag:header('etag')||(unchanged?validator.etag:''),lastModified:header('last-modified')||(unchanged?validator.lastModified:'')}});
              }else finish(new Error('HTTP '+response.status+(response.statusText?' '+response.statusText:'')+(response.status>=200&&response.status<300?' (empty response).':response.status===304?' (no matching parsed cache).':'.')));
            },
            onerror:response=>finish(new Error((response?.status?'HTTP '+response.status+'. ':'')+(response?.error?.message||response?.error||response?.message||response?.statusText||'Network or Tampermonkey permission error.'))),
            ontimeout:()=>finish(new Error('Data download timed out (30 s).')),
            onabort:()=>finish(new Error('Data download was cancelled.'))
          });requests.push(request);
        }catch(error){finish(error)}
      });
    }catch(error){if(!index)releaseMirrors();throw error}
  });
  try{return await Promise.any(attempts)}catch(error){
    const failures=error.errors.map((problem,index)=>new URL(HOSTED_ARCHIVE_URLS[index]).hostname+': '+problem.message);
    throw new Error('All data download sources failed.\n'+failures.join('\n')+'\nCheck Tampermonkey connection permissions and network/proxy access.');
  }finally{
    settled=true;clearTimeout(hedge);releaseMirrors();timers.forEach(clearTimeout);
    for(const request of requests)try{request?.abort?.()}catch{}
  }
}
function unpackDashboardSource(key,cache){
  if(key==='nsr'&&cache?.schema===1&&!cache.columns){
    if(!Array.isArray(cache.rows)||!cache.rows.length||!nsrData.validSapDate(cache.from)||!nsrData.validSapDate(cache.to))return null;
    return {rows:normalizeNsrRows(cache.rows),meta:{file:'NSR cache',kind:'sap',sheet:'SearchNsrSet',from:cache.from,to:cache.to,pages:cache.pages,savedAt:cache.savedAt||Date.now()}};
  }
  const required=[key==='nsr'?'NSR#':sources[key].fields.fid.label,sources[key].fields.number?.label].filter(Boolean);
  if(!validSourceCache(cache,required,Boolean(BW_SOURCES[key]))||
    key==='ctlNsr'&&(!Array.isArray(cache.meta.rowTimes)||cache.meta.rowTimes.length!==cache.rows.length||cache.meta.rowTimes.some(time=>!Number.isFinite(time)||time<0)))return null;
  const columns=cache.columns;
  const rows=cache.rows.map(values=>{
    const row={};
    for(let index=0;index<columns.length;index++){
      const value=values[index];
      if(value===undefined)continue;
      const column=columns[index];
      if(column==='__proto__')Object.defineProperty(row,column,{value,enumerable:true,writable:true,configurable:true});
      else row[column]=value;
    }
    return row;
  });
  return {rows,meta:{...cache.meta,fields:columns},revision:cache.revision};
}
async function readSourceCache(key){return unpackDashboardSource(key,await nsrData.readDataset(cacheKeys[key]))}
async function restoreSourceCache(key,run,render=true){
  if(store.meta[key])return;
  const cache=await readSourceCache(key);
  if(cache&&run===store.runs[key]){
    sourceCacheRevisions.set(cache.rows,cache.revision);
    await applySource(key,cache.rows,{...cache.meta,cached:true,cacheReason:''},render);
  }
}
function sourceCacheCurrent(key,today=nsrData.localDate(new Date())){
  const meta=store.meta[key];if(!meta||meta.lookupOnly)return false;
  if(key==='nsr')return meta.kind==='file'||meta.to===today;
  return (HOSTED_PREFIXES[key]?meta.hostedDay:nsrData.localDate(new Date(meta.savedAt)))===today;
}
async function sourceFailure(key,run,message,transport=false){
  if(run!==store.runs[key])return;
  await restoreSourceCache(key,run).catch(()=>{});
  if(run!==store.runs[key])return;
  const meta=store.meta[key];if(meta){meta.cached=true;meta.cacheReason=message}
  sourceMessage(key,message,meta?'ok':'error',transport);
}
function readFidDateMarks(entry){
  if(Object.prototype.toString.call(entry?.readDates)==='[object Map]')return entry.readDates;
  const marks=new Map();
  if(entry?.schema===1&&Object.prototype.toString.call(entry.changes)==='[object Map]'){
    // Carry forward dates the user explicitly reviewed in older versions.
    for(const [fid,fields] of entry.changes){
      const read={};
      for(const key of fidDateKeys){
        const field=fields[key];
        if(!Number.isInteger(field?.readThrough)||field.readThrough<1||!Array.isArray(field.events))continue;
        const dates=new Set();
        for(const {from,to} of field.events.slice(0,field.readThrough)){
          if(from)dates.add(from);
          if(to)dates.add(to);
        }
        if(dates.size)read[key]=[...dates];
      }
      if(read.mcsd||read.rd)marks.set(fid,read);
    }
  }
  return marks;
}
function packDashboardSource(key,rows,{fields,...meta},allRows=false){
  const names=new Set(key==='ctlNsr'?fields:[]);
  for(const row of allRows||key==='nsr'?rows:rows.slice(0,1))for(const column of Object.keys(row))names.add(column);
  return packSourceCache(cacheKeys[key],rows,meta,[...names]);
}
function rememberSourceCache(rows,cache){
  // Reuse a model built from live rows only when reload will reproduce their exact values and field order.
  const positions=new Map(cache.columns.map((column,index)=>[column,index]));
  const restorable=rows.every(row=>{
    let previous=-1;
    for(const column of Object.keys(row)){
      const index=positions.get(column);
      if(!(index>previous)||row[column]===undefined)return false;
      previous=index;
    }
    return true;
  });
  if(restorable)sourceCacheRevisions.set(rows,cache.revision);
}
async function saveSourceCache(key,run,unchanged=false){
  const rows=store.raw[key],meta=store.meta[key];
  if(meta.cacheReadError){sourceMessage(key,'Cache not saved: '+meta.cacheReadError,'error');return}
  try{
    const cache=packDashboardSource(key,rows,meta);
    if(unchanged&&sourceCacheRevisions.has(rows))cache.revision=sourceCacheRevisions.get(rows);
    await nsrData.writeDataset(cache);
    rememberSourceCache(rows,cache);
  }
  catch(error){if(run===store.runs[key])sourceMessage(key,'Cache not saved: '+error.message,'error')}
}
async function appendBwLookup(incoming){
  if(!incoming.fcid.length&&!incoming.ctl.length)return false;
  const keys=['fcid','ctl'],current=keys.map(key=>({key,rows:store.raw[key],meta:store.meta[key]}));
  const rowId=(key,row)=>text(sourceValue(row,key,key==='fcid'?'fid':'number'));
  function appendMissing(key,rows,candidates){
    if(!candidates.length)return rows;
    const ids=new Set(rows.map(row=>rowId(key,row)));
    const added=candidates.filter(row=>{const number=rowId(key,row);return number&&!ids.has(number)});
    return added.length?rows.concat(added):rows;
  }
  function merge(snapshots){
    const fids=appendMissing('fcid',snapshots[0].rows,incoming.fcid),fidIds=new Set(fids.map(row=>rowId('fcid',row)));
    const ctls=appendMissing('ctl',snapshots[1].rows,incoming.ctl.filter(row=>fidIds.has(text(row['Opportunity ID']))));
    return snapshots.map((snapshot,index)=>{
      const {key}=snapshot,rows=index===0?fids:ctls;
      if(rows===snapshot.rows)return snapshot;
      const {label,query}=BW_SOURCES[key];
      const meta=snapshot.meta||{file:'BW InA · '+label+' lookup',kind:'sap',sheet:query,lookupOnly:true,savedAt:Date.now()};
      return {key,rows,meta:{...meta,...(meta.dateField?{}:sourceRange(rows,key,meta)),rows:rows.length}};
    });
  }
  keys.forEach(key=>sourceBusy(key,true));
  try{
    let next,caches=[],cacheError='';
    try{
      caches=await nsrData.updateDatasets(keys.map(key=>cacheKeys[key]),stored=>{
        // Merge against the latest snapshot so two tabs can append without replacing each other's records.
        const bases=current.map((snapshot,index)=>{
          const entry=stored[index];
          if(entry?.revision&&entry.revision===sourceCacheRevisions.get(snapshot.rows))return snapshot;
          const cached=unpackDashboardSource(snapshot.key,entry);
          if(!cached||cached.meta.savedAt<(snapshot.meta?.savedAt||0))return {...snapshot,needsSave:Boolean(snapshot.meta)};
          const rows=cached.meta.savedAt===snapshot.meta?.savedAt?appendMissing(snapshot.key,cached.rows,snapshot.rows):cached.rows;
          const needsSave=rows!==cached.rows;
          return {key:snapshot.key,rows,meta:{...cached.meta,cached:true,rows:rows.length},needsSave,revision:needsSave?undefined:cached.revision};
        });
        next=merge(bases);
        return next.flatMap((snapshot,index)=>snapshot===bases[index]&&!snapshot.needsSave?[]:[packDashboardSource(snapshot.key,snapshot.rows,snapshot.meta,true)]);
      });
    }catch(error){cacheError=error.message;next=merge(current)}
    let updated=false;
    for(const [index,{key,rows,meta,revision:cachedRevision}] of next.entries()){
      const cache=caches.find(entry=>entry.key===cacheKeys[key]);
      if(cache)rememberSourceCache(rows,cache);else if(cachedRevision)sourceCacheRevisions.set(rows,cachedRevision);
      const changed=rows!==current[index].rows;
      if(changed){setSourceRows(key,rows,meta);updated=true}
      if(changed||cache)sourceMessage(key,cacheError?'Cache not saved: '+cacheError:'Loaded',cacheError?'error':'done');
    }
    return updated;
  }finally{keys.forEach(key=>sourceBusy(key,false))}
}
function mergeCtlNsr(rows,time,append=false){
  const merged=new Map(),previous=store.meta.ctlNsr?.rowTimes||[],keyOf=row=>text(row['CTL ID'])||'\0'+JSON.stringify(row);
  store.raw.ctlNsr.forEach((row,index)=>merged.set(keyOf(row),{row,time:previous[index]||0}));
  const incoming=new Map(rows.map(row=>[keyOf(row),row]));
  for(const [key,row]of incoming){
    const existing=merged.get(key);
    if(!existing||append||time>existing.time){merged.delete(key);merged.set(key,{row,time})}
  }
  const records=[...merged.values()];return {rows:records.map(record=>record.row),rowTimes:records.map(record=>record.time),fields:[...new Set([...(store.meta.ctlNsr?.fields||[]),...Object.keys(rows[0]||{})])]};
}
async function loadHostedSources(keys=Object.keys(HOSTED_PREFIXES),force=false){
  const runs=Object.fromEntries(keys.map(key=>[key,startSource(key)])),today=nsrData.localDate(new Date());keys.forEach(key=>sourceMessage(key,'Loading data…','loading'));
  const cacheErrors={};let pending=keys,stage='SAP access',passwordController;
  const failed=(key,error)=>sourceFailure(key,runs[key],stage+' failed: '+(error.userMessage||error.message)+(stage==='SAP access'&&!store.meta[key]?'\nSign in to Fiori.':''));
  try{
    await Promise.all(keys.map(key=>restoreSourceCache(key,runs[key],false).catch(error=>{cacheErrors[key]=error.message})));
    pending=keys.filter(key=>runs[key]===store.runs[key]&&(force||!sourceCacheCurrent(key,today)));
    keys.filter(key=>runs[key]===store.runs[key]&&!pending.includes(key)).forEach(key=>sourceMessage(key,'Loaded'));
    if(!pending.length)return;
    pending.forEach(key=>sourceMessage(key,'Loading data…','loading'));
    passwordController=new unsafeWindow.AbortController();
    const passwordTimeout=setTimeout(()=>passwordController.abort(new Error('SAP request timed out (30 s).')),30000);
    // Handle password errors immediately, but let a validated parsed cache succeed without SAP access.
    const passwordRequest=(async()=>{
      try{
        const response=await unsafeWindow.fetch(HOSTED_PASSWORD_URL,{credentials:'same-origin',cache:'no-store',signal:passwordController.signal,headers:{Accept:'application/atom+xml,application/xml,text/xml'}});
        if(!response.ok)throw new Error('HTTP '+response.status);
        if((response.headers.get('Content-Type')||'').toLowerCase().includes('text/html'))throw new Error('SAP returned a sign-in page.');
        const doc=new DOMParser().parseFromString(await response.text(),'application/xml');
        if(doc.getElementsByTagName('parsererror').length)throw new Error('SAP returned invalid XML.');
        const properties=[...doc.getElementsByTagNameNS('*','properties')].find(node=>node.getElementsByTagNameNS('*','Userid')[0]?.textContent.trim()==='LONGKY');
        const password=properties?.getElementsByTagNameNS('*','FullName')[0]?.textContent.replace(/\s+/g,'')||'';
        if(!password)throw new Error('SAP returned no password.');
        return {password};
      }catch(error){return {error}}finally{clearTimeout(passwordTimeout)}
    })();
    stage='GitHub download';
    const archive=await downloadHostedArchive(pending,force);if(pending.every(key=>runs[key]!==store.runs[key]))return;
    if(archive.unchanged){
      for(const key of pending){
        if(runs[key]!==store.runs[key])continue;
        Object.assign(store.meta[key],{hostedDay:today,hostedValidator:archive.validator,cached:false,cacheReason:''});
        sourceMessage(key,'Loaded');sourceMeta(key);await saveSourceCache(key,runs[key],true);
      }
      return;
    }
    stage='SAP access';
    const {password:archivePassword,error}=await passwordRequest;if(error)throw error;
    if(pending.every(key=>runs[key]!==store.runs[key]))return;
    let bytes=archive.bytes;
    stage='Archive decryption';
    if(typeof zip==='undefined')throw new Error('Data reader is unavailable. Import a file manually.');
    // zip.js reads the encrypted ZIP appended after the PNG.
    const outer=new zip.ZipReader(new zip.Uint8ArrayReader(new Uint8Array(bytes)),{useWebWorkers:false});
    try{
      const entry=(await outer.getEntries()).find(item=>!item.directory&&item.filename.split(/[\\/]/).pop().toLowerCase()==='data.zip');
      if(!entry)throw new Error('The data file is missing.');
      bytes=await entry.getData(new zip.Uint8ArrayWriter(),{password:archivePassword,checkSignature:true});
    }finally{await outer.close()}
    stage='Archive contents';
    const reader=new zip.ZipReader(new zip.Uint8ArrayReader(new Uint8Array(bytes)),{useWebWorkers:false});
    try{
      const entries=await reader.getEntries();
      for(const key of pending){
        if(runs[key]!==store.runs[key])continue;
        try{
          stage='Read '+sources[key].label;
          const entry=entries.find(item=>!item.directory&&item.filename.split(/[\\/]/).pop().toLowerCase().startsWith(HOSTED_PREFIXES[key]));
          if(!entry)throw new Error('The data file is missing.');
          const filename=entry.filename.split(/[\\/]/).pop(),modifiedAt=entry.lastModDate?.getTime();
          if(!Number.isFinite(modifiedAt)||modifiedAt<=0)throw new Error('The data file timestamp is invalid.');
          const data=await entry.getData(new zip.Uint8ArrayWriter(),{checkSignature:true}),parsed=await report.readFile(new File([data],filename),key);
          if(runs[key]!==store.runs[key])continue;
          const {rows,...recognition}=parsed,merged=key==='ctlNsr'?mergeCtlNsr(rows,modifiedAt):{rows};
          await applySource(key,merged.rows,{file:filename,kind:'github',...recognition,modifiedAt,hostedDay:today,hostedValidator:archive.validator,savedAt:Date.now(),cacheReadError:cacheErrors[key]||'',...(merged.rowTimes?{rowTimes:merged.rowTimes,fields:merged.fields}:{})},false,runs[key]);
        }catch(error){await failed(key,error)}
      }
    }finally{await reader.close()}
  }catch(error){await Promise.all(pending.map(key=>failed(key,error)))}finally{passwordController?.abort();try{if(store.dirty&&!store.refreshing)await recalculate()}finally{keys.forEach(key=>{if(runs[key]===store.runs[key])sourceBusy(key,false)})}}
}
async function importSource(file,key){
  const run=startSource(key);sourceMessage(key,'Reading '+file.name+'…','loading');
  try{
    const parsed=await report.readFile(file,key);let range={savedAt:Date.now()};
    if(key==='nsr'){parsed.rows=normalizeNsrRows(parsed.rows);if(nsrData.validSapDate(parsed.from)&&nsrData.validSapDate(parsed.to))range={...range,from:parsed.from,to:parsed.to};delete parsed.from;delete parsed.to}
    if(run!==store.runs[key])return;
    if(key==='ctlNsr'){
      const savedAt=Date.now(),meta=store.meta[key],merged=mergeCtlNsr(parsed.rows,savedAt,true);parsed.rows=merged.rows;
      // An unchanged future-dated CSV can still outrank this upload on the next merge.
      range={rowTimes:merged.rowTimes,fields:merged.fields,hostedDay:meta?.hostedDay||'',modifiedAt:meta?.modifiedAt||0,...(meta?.hostedValidator&&meta.modifiedAt<=savedAt?{hostedValidator:meta.hostedValidator}:{}),savedAt,cacheReadError:meta?.cacheReadError||''};
    }
    if(key==='nce')range.hostedDay=nsrData.localDate(new Date());
    const {rows,...recognition}=parsed;await applySource(key,rows,{file:file.name,kind:'file',...recognition,...range},true,run);
  }catch(error){if(run===store.runs[key])sourceMessage(key,(key==='ctlNsr'?'Append':'Import')+' failed: '+(error.userMessage||error.message)+(store.meta[key]?' Previous data retained.':''),'error')}finally{if(run===store.runs[key]){sourceBusy(key,false);
    id('fcFile-'+key).value=''}}
}
async function loadBw(key='fcid',useCache=false,range){
  const run=startSource(key),controller=new unsafeWindow.AbortController();store.controllers[key]=controller;
  sourceMessage(key,useCache?'Checking BW cache…':'Loading BW InA · 0%','loading');
  try{
    if(useCache){
      await restoreSourceCache(key,run).catch(()=>{});if(run!==store.runs[key])return;
      if(store.meta[key]&&!store.meta[key].lookupOnly){store.meta[key].cached=true;sourceMessage(key,'Using cached BW data');return}
    }
    sourceMessage(key,'Loading BW InA · 0%','loading');
    const {rows,meta,warning,cleanupError}=await readBwSnapshot(bwData,key,controller.signal,message=>{if(run===store.runs[key])sourceMessage(key,message,'loading')},range||(store.meta[key]?.customRange?store.meta[key]:undefined));
    if(run!==store.runs[key])return;
    sourceMessage(key,'Loading BW InA · 100% · Processing…','loading');
    await applySource(key,rows,meta,true,run);
    const notice=[warning,cleanupError].filter(Boolean).join(' ');
    if(notice&&run===store.runs[key])sourceMessage(key,notice,'error');
  }catch(error){await sourceFailure(key,run,'BW load failed: '+error.message,Boolean(error.bwLogin))}
  finally{if(run===store.runs[key]){sourceBusy(key,false);delete store.controllers[key]}}
}
function setNsrRange(range){nsrRange=range;id('fcNsrFrom').value=range.from;id('fcNsrTo').value=range.to;id('fcNsrTo').min=range.from;id('fcNsrFrom').max=id('fcNsrTo').max=nsrData.localDate(new Date())}
async function loadSharedNsr(useCache=true){
  const key='nsr',run=startSource(key);sourceMessage(key,useCache?'Checking NSR cache…':'Loading NSR…','loading'); const today=nsrData.localDate(new Date());let range={from:id('fcNsrFrom').value,to:id('fcNsrTo').value},customRange=!useCache&&range.from!==nsrData.sapRange().from;
  try{
    if(useCache){
      await restoreSourceCache(key,run).catch(()=>{});if(run!==store.runs[key])return;
      const meta=store.meta[key];
      if(meta){
        if(meta.from&&meta.to){range={from:meta.from,to:meta.to};setNsrRange(range)}
        if(sourceCacheCurrent(key,today)){sourceMessage(key,'Using cached NSR data');return}
        // Automatic updates roll with the default window unless Load NSR chose another start date.
        customRange=Boolean(meta.customRange);range=customRange?{from:meta.from,to:today}:nsrData.sapRange();
      }else range=nsrData.sapRange();
    }
    if(!nsrData.validSapDate(range.from)||!nsrData.validSapDate(range.to)||range.from>range.to||range.to>today)throw new Error('Choose a valid NSR date range ending on or before today.');
    const controller=new unsafeWindow.AbortController();store.controllers[key]=controller; const result=await nsrData.fetchSapRows(range,controller.signal);if(run!==store.runs[key])return;
    setNsrRange(range);await applySource(key,normalizeNsrRows(result.rows),{file:'SAP live NSR',kind:'sap',sheet:'SearchNsrSet',...range,...(customRange?{customRange:true}:{}),pages:result.pages,savedAt:Date.now()},true,run);
  }catch(error){await sourceFailure(key,run,'NSR load failed: '+error.message)}
  finally{if(run===store.runs[key]){sourceBusy(key,false);delete store.controllers[key]}}
}
async function restoreSources(reset=false){
  if(store.refreshing||store.busy.size)return;
  if(bwLookup?.controller)clearBwLookup();
  store.refreshing=true;const refresh=id('fcRefresh');refresh.disabled=true;refresh.textContent='Refreshing…';
  let needsRefreshBuild=reset;
  try{
    const keys=Object.keys(sources);
    if(reset){
      for(const [name,,path] of filterDefs)if(!id(name).closest('.fc-filter').hidden)pendingFilterValues[path]??=filterValues[path]||[];
      dataRevision++;
      keys.forEach(key=>{store.runs[key]=(store.runs[key]||0)+1;store.raw[key]=[];delete store.meta[key]; sourceMessage(key,'Not loaded','');sourceMeta(key);sourceBusy(key,true);id('fcFile-'+key).value='';}); setNsrRange(nsrData.sapRange());actionMessage('');
      await recalculate(true);
    }
    // Retain hosted merged data and existing FID/CTL fallback caches; discard other legacy source caches.
    try{await nsrData.deleteDatasets([...keys.filter(key=>!HOSTED_PREFIXES[key]).map(key=>CACHE_PREFIX+key),CACHE_PREFIX+'archive'])} catch(error){if(reset)actionMessage('Local source cache could not be reset. '+error.message)} finally{if(reset)keys.forEach(key=>sourceBusy(key,false))}
    if(!reset){
      // Restore saved data first so both normal opens and lookups can use the calculated cache.
      const dateReads=nsrData.readDataset(FID_DATE_READ_KEY).then(entry=>{fidDateReads=readFidDateMarks(entry)}).catch(()=>{});
      await Promise.all([dateReads,...keys.map(async key=>{
        const run=startSource(key);
        sourceMessage(key,'Reading local cache…','loading');
        try{await restoreSourceCache(key,run,false)}catch{}
        finally{if(run===store.runs[key])sourceBusy(key,false)}
      })]);
      if(store.dirty&&(linkedId||Object.keys(store.meta).length===keys.length)&&await recalculate())openLinkedFid(false,true);
    }
    if(store.meta.nsr?.from)setNsrRange({from:store.meta.nsr.from,to:store.meta.nsr.to});
    const loads=[];
    for(const key of Object.keys(BW_SOURCES))if(!store.busy.has(key))loads.push(loadBw(key,true));
    const hostedKeys=Object.keys(HOSTED_PREFIXES).filter(key=>!store.busy.has(key));if(hostedKeys.length)loads.push(loadHostedSources(hostedKeys,reset));
    if(!store.busy.has('nsr')&&(!store.meta.nsr||store.meta.nsr.kind==='sap'))loads.push(loadSharedNsr());
    // Keep per-source progress visible; rebuild the graph once the whole load settles.
    await Promise.all(loads);
    const expired=Object.keys(BW_SOURCES).filter(key=>store.meta[key]&&!store.meta[key].lookupOnly&&!sourceCacheCurrent(key));
    if(expired.length){
      if(store.dirty&&await recalculate(reset))needsRefreshBuild=false;
      // Show the cached dashboard before asking about the slower BW update.
      await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
      if(confirmBwRefresh(expired,key=>nsrData.localDate(new Date(store.meta[key].savedAt))))await Promise.all(expired.map(key=>loadBw(key)));
    }
  }finally{store.refreshing=false;refresh.disabled=store.busy.size>0;refresh.textContent='Refresh';if(store.dirty||needsRefreshBuild)await recalculate(reset);syncFidLookup();openLinkedFid(!store.busy.size)}
}


function additionalFidColumns(){
  return fidColumns().map(([label,get,render,settings,custom])=>[label,row=>{
    const node=graph.fids.get(row.id)||(custom?row:null);return node?get(node):label==='FID'?row.id:'';
  },(value,row)=>{
    const node=graph.fids.get(row.id)||(custom?row:null);return node&&render?render(value,node):label==='FID'?fidLink(row.id):label==='Mark'||label==='CTL Elapsed Days'?'':escapeHtml(value)||'—';
  },settings,custom]);
}
const cycle=createNsrCycle(id('workspace'),{getRows:nsrs=>{
  const selected=new Set(nsrs);
  return nsrCopyRows(graph.nsrs.filter(nsr=>selected.has(nsr.number))).map(row=>({...row,'REF FCID':row.FID}));
},renderTable:(name,columns,rows,options)=>{
  if(name==='wfAdditional'){
    rows=rows.map(row=>({id:row.FID}));
    columns=additionalFidColumns();
    options={...options,fid:true,frozen:['FID'],copyRows:fidCopyRows,exportRows:fidDetailMatrix,widths:fidTableWidths};
  }
  return table(name,columns,rows,options,true);
},renderFid:fidLink,copyRows,onNotify:message=>toast(message)});
if(savedFilterPresets[0])applyFilterPreset(savedFilterPresets[0]);
unsafeWindow.addEventListener('hashchange',readFidLink);readFidLink();
setNsrRange(nsrRange);recalculate();restoreSources();
}

function sourceDate(value,dates){
  if(value instanceof Date&&!isNaN(value))return dates.localDate(value);
  const input=String(value??'').trim();let iso='';
  if(/^\d{5,7}(?:\.\d+)?$/.test(input)){const date=XLSX.SSF.parse_date_code(Number(input)); if(date)iso=`${date.y}-${String(date.m).padStart(2,'0')}-${String(date.d).padStart(2,'0')}`;}else if(/^\d{4}-\d{2}-\d{2}(?:[ T].*)?$/.test(input))iso=input.slice(0,10);
    else{const match=input.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s.*)?$/); if(match)iso=`${match[3]}-${match[1].padStart(2,'0')}-${match[2].padStart(2,'0')}`;}
  return dates.validSapDate(iso)?iso:'';
}

function sourceDateRange(rows,key,dates,columns){
  const config=SOURCE_DATES[key];if(!config)return {};
  const fields=columns?config.fields.map(field=>columns.indexOf(field)).filter(index=>index>=0):config.fields;
  let dateFrom='',dateTo='';
  for(const row of rows){let date='';for(const field of fields){date=sourceDate(row[field],dates);if(date)break}if(!date)continue; if(!dateFrom||date<dateFrom)dateFrom=date;if(!dateTo||date>dateTo)dateTo=date;}
  return {dateFrom,dateTo,dateField:config.label};
}

function paintSourceStatus(target,meta,dates,message,kind,pending=false){
  const hasRange=Boolean(meta?.dateFrom&&meta?.dateTo);
  const cutoff=new Date();cutoff.setDate(cutoff.getDate()-3);const stale=hasRange&&meta.dateTo<=dates.localDate(cutoff);
  const progress=kind==='loading'?message.match(/\b\d+%/)?.[0]:'';
  const label=pending?'Loading…':kind==='loading'?(/updat/i.test(message)?'Updating…':'Loading…')+(progress?' '+progress:''):meta?(hasRange?meta.dateFrom+' ~ '+meta.dateTo:'Loaded · no dates'):'Unavailable';
  target.textContent=label;target.className='fc-source-state '+(kind==='loading'?'loading':meta?'ok':'')+(hasRange?' has-range':'')+(kind!=='loading'&&stale?' stale':'')+(kind==='error'?' error':'')+(pending?' loading pending':'');
  target.title=kind==='loading'?message:meta?.cached?['Using cache · '+new Date(meta.savedAt).toLocaleString('en-GB'),meta.modifiedAt?(meta.kind==='bridge'?'Workbook updated · ':'Source updated · ')+new Date(meta.modifiedAt).toLocaleString('en-GB'):'',kind==='error'?message:meta.cacheReason].filter(Boolean).join('\n'):kind==='error'?message:hasRange?'Data range · '+meta.dateField:message;
  target.setAttribute('aria-label',label);target.tabIndex=kind==='error'||meta?.cached?0:-1;
}

function paintSourceFile(target,meta,fullPath){
  target.textContent=meta?String(meta.file??'').trim().split(/[\\/]/).pop():'';
  target.title=fullPath?meta?.file||'':target.textContent;
}

function validSourceCache(cache,required,legacy){
  return (cache?.schema===2||legacy&&cache?.schema===1)&&typeof cache.meta?.file==='string'&&Number.isFinite(cache.meta.savedAt)&&cache.meta.savedAt>0&&
    Array.isArray(cache.columns)&&cache.columns.every(column=>typeof column==='string')&&new Set(cache.columns).size===cache.columns.length&&
    Array.isArray(cache.rows)&&cache.rows.every(row=>Array.isArray(row)&&row.length===cache.columns.length)&&(!cache.rows.length||required.every(field=>cache.columns.includes(field)));
}

function packSourceCache(key,rows,meta,columns){
  return {key,schema:2,columns,rows:rows.map(row=>columns.map(column=>row[column])),meta,revision:globalThis.crypto?.randomUUID?.()};
}

function confirmBwRefresh(keys,savedDay){
  const dates=keys.map(key=>BW_SOURCES[key].label+' (cached '+savedDay(key)+')').join('\n');
  return unsafeWindow.confirm(dates+'\n\nThis data is from an earlier day. Update now? This may take a few minutes.');
}

async function readBwSnapshot(service,key,signal,onProgress,range){
  const defaults=service.range(key),selected=service.range(key,range),customRange=selected.from!==defaults.from||selected.to!==defaults.to;
  const {label,query}=BW_SOURCES[key],{rows,fields,pages,from,to,warning,cleanupError}=await service.read(key,query,signal,onProgress,undefined,selected);
  return {rows,meta:{file:'BW InA · '+label,kind:'sap',sheet:query,fields,pages,from,to,...(customRange?{customRange:true}:{}),savedAt:Date.now()},warning,cleanupError};
}

function openBwLoadDialog(service,key,meta,onLoad){
  const quarter=key==='fcid',label=quarter?'FID':'CTL';let range;
  try{range=service.range(key,meta?.customRange?meta:undefined)}catch{range=service.range(key)}
  let dialog=document.getElementById('bwLoadDialog');
  if(!dialog){dialog=document.createElement('dialog');dialog.id='bwLoadDialog';dialog.className='bw-load-dialog';dialog.setAttribute('aria-labelledby','bwLoadTitle');document.body.appendChild(dialog)}
  dialog.dataset.source=key;
  dialog.innerHTML=`<style>
.bw-load-dialog{width:min(450px,94vw);border:1px solid #b7c7e0;border-radius:12px;padding:17px;background:#f6f8fe;color:#2e486c;box-shadow:0 16px 70px #1e355550}.bw-load-dialog::backdrop{background:#152e505c}.bw-load-dialog h2{margin:0;font-size:20px;font-weight:600}
.bw-load-range{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:10px;margin-top:14px}.bw-load-range label{display:block}.bw-load-range input,.bw-load-range select{width:100%;min-width:0;padding:8px;border:1px solid #c2d0e4;border-radius:6px;background:#fff;color:inherit;margin-top:5px}.bw-load-quarter{display:grid;grid-template-columns:minmax(0,1fr) 64px;gap:6px}.bw-load-actions{display:flex;justify-content:flex-end;gap:7px;margin-top:12px}
</style><form id="bwLoadForm"><h2 id="bwLoadTitle">Load ${label}</h2><div class="bw-load-range">${['from','to'].map(side=>{
    const name=side==='from'?'From':'To',value=range[side];
    return quarter?`<div><label for="bwLoad${name}Year">${name} (calendar quarter)</label><div class="bw-load-quarter"><input id="bwLoad${name}Year" type="number" aria-label="${name} year" min="1000" max="9999" value="${value.slice(0,4)}" required><select id="bwLoad${name}Quarter" aria-label="${name} quarter">${[1,2,3,4].map(q=>`<option value="${q}"${Number(value[4])===q?' selected':''}>Q${q}</option>`).join('')}</select></div></div>`:
      `<label>${name} (created date)<input id="bwLoad${name}" type="date" value="${value.slice(0,4)}-${value.slice(4,6)}-${value.slice(6,8)}" max="${service.range(key).to.replace(/^(\d{4})(\d{2})(\d{2})$/,'$1-$2-$3')}" required></label>`;
  }).join('')}</div><div class="bw-load-actions"><button class="btn light" type="button">Cancel</button><button class="btn blue" type="submit">Load</button></div></form>`;
  const form=dialog.querySelector('form'),control=name=>form.querySelector('#bwLoad'+name),to=control(quarter?'ToYear':'To');
  if(!quarter)to.min=control('From').value;
  form.addEventListener('input',()=>{to.setCustomValidity('');if(!quarter)to.min=control('From').value});
  form.querySelector('[type="button"]').addEventListener('click',()=>dialog.close());
  form.addEventListener('submit',event=>{
    event.preventDefault();
    const selected=Object.fromEntries(['From','To'].map(name=>[name.toLowerCase(),quarter?control(name+'Year').value+control(name+'Quarter').value:control(name).value.replaceAll('-','')]));
    try{service.range(key,selected)}catch(error){to.setCustomValidity(error.message);to.reportValidity();return}
    dialog.close();onLoad(selected);
  });
  dialog.showModal();
}

// BW sessions, CSRF tokens and query instances stay in memory; cached data contains business rows only.
function createBwDataService(){
  const base='https://pcazbwpap04.fremont.lamrc.net:8443/sap/bw/ina/',suffix='?sap-client=010&sap-language=EN';
  // Dashboard column, row dimension, exact BW attribute, optional date conversion.
  const fields={fcid:[
    ['FID','CRHOPPTID','CRHOPPTID.DISPLAY_KEY'],
    ['Commit: Field','PHCOPPTID__PHCFCOMLV','PHCOPPTID__PHCFCOMLV.SHORT_TEXT'],
    ['Business Unit','CRHBU','CRHBU.MEDIUM_TEXT'],
    ['Fab Name','CRHOPPTID','CRHOPPTID.CRHFABID.MEDIUM_TEXT'],
    ['Fab Segment','CRHFABSGT','CRHFABSGT.SHORT_TEXT'],
    ['Customer','0CUST_GRP2','0CUST_GRP2.SHORT_TEXT'],
    ['FID Description','PHCFIDTXT','PHCFIDTXT.DISPLAY_KEY'],
    ['System Description','PHCOPPTID__CRHSYSDES','PHCOPPTID__CRHSYSDES.DISPLAY_KEY'],
    ['Date: Customer Request','PHTFCRDAT__0DATE','PHTFCRDAT__0DATE.KEY',true],
    ['Date: Target Ship','PHTPSPDAT__0DATE','PHTPSPDAT__0DATE.KEY',true],
    ['Date: Mfg Commit Ship','PHCOPPTID__CRHSCHSDT','PHCOPPTID__CRHSCHSDT.KEY',true],
    ['Date: Sales Ops Request Date (RD)','PHTBDCRD__0DATE','PHTBDCRD__0DATE.KEY',true],
    ['Mfg Upgrade Type','PHCOPPTID__CRHPROJIN','PHCOPPTID__CRHPROJIN.MEDIUM_TEXT'],
    ['Slot Actual','CRHOPPTID','CRHOPPTID.CRMT2ACT.DISPLAY_KEY',true],
    ['Time: Calendar Quarter','PHTCALQTR','PHTCALQTR.SHORT_TEXT'],
    ['Date: Initial MCSD at Slot','PHCMCSDSL__0DATE','PHCMCSDSL__0DATE.KEY',true],
    ['Initial SRD at Slot','PHCSARSOP','PHCSARSOP.KEY',true]
  ],ctl:[
    ['Opportunity ID','CRHOPPTID','CRHOPPTID.DISPLAY_KEY'],
    ['CTL ID','CRHCFGCID','CRHCFGCID.DISPLAY_KEY'],
    ['Change Cat','CRCCHCAT','CRCCHCAT.SHORT_TEXT'],
    ['MCSD @ Current','CRHSCHSDT','CRHSCHSDT.KEY',true],
    ['MCSD@Create','CRCMCSDCU','CRCMCSDCU.KEY',true],
    ['RD At Create','CRCBRDCRE','CRCBRDCRE.KEY',true],
    ['MCSD At SOps Approval','CRCMCSDSO','CRCMCSDSO.KEY',true],
    ['TSD @ SOPs Approval','CRCTSDSOA','CRCTSDSOA.KEY',true],
    ['SRD @ Current','CRCBRDCUR','CRCBRDCUR.KEY',true],
    ['SRD @ SOPs Approval','CRCBRDSOA','CRCBRDSOA.KEY',true],
    ['1st SOps Appr Dt','CRCSOPD','CRCSOPD.KEY',true],
    ['Create Date','0CRM_CRD_AT','0CRM_CRD_AT.KEY',true],
    ['Initiator','0CRM_CTD_BY','0CRM_CTD_BY.DISPLAY_KEY'],
    ['Business Unit','CRHBU','CRHBU.MEDIUM_TEXT'],
    ['Change Summary','CRCSUMTX','CRCSUMTX.LONG_TEXT'],
    ['Stage','CRCSTAGE','CRCSTAGE.LONG_TEXT'],
    ['Quote ID','CRHQUONUM','CRHQUONUM.DISPLAY_KEY'],
    ['Customer group 2','CRHOPPTID','CRHOPPTID.0CUST_GRP2.SHORT_TEXT'],
    ['Config Frz@Current','CRHCFGFRZ','CRHCFGFRZ.KEY',true],
    ['Root Cause','CRCROOTTX','CRCROOTTX.LONG_TEXT'],
    ['Sub Reason','CRHSRCCHG','CRHSRCCHG.SHORT_TEXT'],
    ['Field CRD','CRFDIRCDT','CRFDIRCDT.KEY',true],
    ['Reason for Chg','CRHRNFCHG','CRHRNFCHG.SHORT_TEXT'],
    ['Sales OPs Approver','CRCSOP','CRCSOP.SHORT_TEXT']
  ]};
  const wanted=['ClientCapabilities','DatasourceAtService','StatefulServer','ServerStructureNames','SupportsSetOperand','SupportsEncodedResultSet','ResultSetAxisType','PagingTupleCountTotal','Obtainability','ReadMode','ReadModesV2','Supplements','SAPDate','SubmitReturnsVariableValues','DefinitionReturnsVariableValues'];
  const operands=selection=>selection?.SetOperand?[selection.SetOperand]:(selection?.Operator?.SubSelections||[]).flatMap(operands);
  const values=value=>{if(Array.isArray(value))return value;if(value&&(!value.Encoding||value.Encoding==='None')&&Array.isArray(value.Values))return value.Values;throw new Error('Unsupported BW result encoding.')};
  // Returns null for an unreadable date so one bad BW value does not fail the whole read.
  function date(value){
    const text=String(value??'').trim();if(!text||/^(?:#|Not assigned|00000000|0000-00-00)$/i.test(text))return '';
    let parts=text.match(/^(\d{4})-?(\d{2})-?(\d{2})$/);
    if(!parts){const us=text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(us)parts=['',us[3],us[1].padStart(2,'0'),us[2].padStart(2,'0')]}
    if(!parts)return null;
    const result=`${parts[1]}-${parts[2]}-${parts[3]}`,parsed=new Date(result+'T00:00:00Z');
    return isNaN(parsed)||parsed.toISOString().slice(0,10)!==result?null:result;
  }
  function range(key,selected,lookup=false){
    const now=new Date(),startYear=lookup?1900:now.getFullYear()-3,today=`${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
    const {from,to}=selected||(key==='fcid'?{from:`${startYear}1`,to:'20994'}:{from:`${startYear}0101`,to:today});
    if(key==='fcid'){
      if(typeof from!=='string'||typeof to!=='string'||!/^[1-9]\d{3}[1-4]$/.test(from)||!/^[1-9]\d{3}[1-4]$/.test(to)||from>to)throw new Error('Choose a valid FID calendar-quarter range.');
    }else{
      const valid=typeof from==='string'&&typeof to==='string'&&/^\d{8}$/.test(from)&&/^\d{8}$/.test(to)&&Boolean(date(from)&&date(to));
      if(!valid||from>to||to>today)throw new Error('Choose a valid CTL creation-date range ending on or before today.');
    }
    return {from,to};
  }
  async function read(key,query,signal,onProgress,lookup,selectedRange){
    const lookupFields=['FID','Opportunity ID','Quote ID','CTL ID'],targets=lookup&&new Set(lookup.values),quoteLookup=lookup?.field==='Quote ID';
    const lookupFormat=quoteLookup?/^\d{8}$/:lookup?.field==='CTL ID'?/^\d{10}$/:/^\d{6}$/;
    const columns=lookup&&!lookup.details&&key==='ctl'?fields.ctl.filter(([column])=>['Opportunity ID','Quote ID','CTL ID','Create Date','Change Cat','Stage'].includes(column)):fields[key];
    const lookupColumn=lookup&&columns.find(([column])=>column===lookup.field&&lookupFields.includes(column));
    if(lookup&&(!lookupColumn||!targets.size||targets.size>500||[...targets].some(value=>!lookupFormat.test(value))))throw new Error('BW lookup requires 1 to 500 valid identifiers.');
    const {from,to}=range(key,selectedRange,Boolean(lookup));
    const source={ObjectName:query,Type:'Query',InstanceId:crypto.randomUUID()};let csrf='',context='',created=false,result;
    const fail=(message,bwLogin=false)=>Object.assign(new Error(message),{bwLogin});
    function request(endpoint,body,stage,timeout=60000,closing=false){
      if(!closing&&signal?.aborted)return Promise.reject(new DOMException('BW load cancelled.','AbortError'));
      const headers={Accept:'application/json','sap-contextid-accept':'header'};
      if(context)headers['sap-contextid']=context;
      if(body){headers['Content-Type']='application/json';if(csrf)headers['x-csrf-token']=csrf}else headers['x-csrf-token']='Fetch';
      return new Promise((resolve,reject)=>{
        let handle;const abort=()=>{handle?.abort?.();finish(reject,new DOMException('BW load cancelled.','AbortError'))};
        const finish=(done,value)=>{signal?.removeEventListener('abort',abort);done(value)};
        try{
          handle=GM_xmlhttpRequest({method:body?'POST':'GET',url:base+endpoint+suffix,headers,data:body?JSON.stringify(body):undefined,anonymous:false,timeout,
            onload:response=>{
              try{
                if([401,403].includes(response.status))throw fail(stage+': BW HTTP '+response.status+'. Open BW, sign in, then click Load.',true);
                if(response.finalUrl&&new URL(response.finalUrl).origin!==new URL(base).origin)throw fail(stage+': BW sign-in is required.',true);
                if(response.status<200||response.status>=300)throw fail(stage+': BW HTTP '+response.status+'.');
                const responseHeaders=response.responseHeaders||'',token=responseHeaders.match(/^x-csrf-token:\s*(.+)$/im)?.[1]?.trim();
                if(token&&!/^(fetch|required)$/i.test(token))csrf=token;
                context=responseHeaders.match(/^sap-contextid:\s*(.+)$/im)?.[1]?.trim()||context;
                let data;try{data=JSON.parse(response.responseText)}catch{throw fail(stage+': Expected BW JSON. Open BW, sign in, then click Load.',true)}
                const errors=[...(data.Messages||[]),...(data.Cube?.Messages||[]),...(data.Grids||[]).flatMap(grid=>grid.Messages||[])].filter(message=>[2,3].includes(Number(message.Type))||/^(error|fatal)$/i.test(message.Type));
                if(errors.length)throw fail(stage+': '+errors.map(message=>message.Text||'BW query error').join('\n'));
                finish(resolve,data);
              }catch(error){finish(reject,error)}
            },
            onerror:()=>finish(reject,fail(stage+': Cannot reach BW. Check the company network / VPN and Tampermonkey site access.',true)),
            ontimeout:()=>finish(reject,fail(stage+': BW request timed out.')),
            onabort:()=>finish(reject,new DOMException('BW load cancelled.','AbortError'))
          });
          if(!closing){signal?.addEventListener('abort',abort,{once:true});if(signal?.aborted)abort()}
        }catch(error){finish(reject,error)}
      });
    }
    try{
      const info=await request('GetServerInfo',null,'Connection'),advertised=new Set((info.Services?.find(service=>service.Service==='Analytics')?.Capabilities||[]).map(item=>typeof item==='string'?item:item.Capability)),caps=wanted.filter(name=>advertised.has(name));
      if(!caps.includes('DatasourceAtService'))throw new Error('BW Analytics service is unavailable.');
      const analytics=definition=>({DataSource:{...source},Capabilities:caps,Language:'EN',Definition:definition});
      created=true;
      let cube=(await request('GetResponse',{Metadata:{Context:'Analytics',DataSource:{...source},Capabilities:caps,Language:'EN',Expand:['Cube']},Options:['StatefulServer']},'Metadata')).Cube;
      if(!cube)throw new Error('BW returned no query metadata.');source.InstanceId=cube.DataSource?.InstanceId||source.InstanceId;
      const fidVariable=key==='fcid'?'ZPH_C_U_SS_OP_CRHOPPTID':'ZVARMOPPTID',bound=key==='fcid'?'ZPH_C_U_SO_OP_QTR':'ZBW_CREATE_DATE';
      const selections=key==='fcid'?{ZPH_C_C_SS_OP_I_CURR_VER:{from:'LIVE'},ZPH_MILESTONE:{from:'SHIP'},[bound]:{from,to}}:{[bound]:{from,to}};
      const inputs=(cube.Variables||[]).filter(variable=>variable.InputEnabled),expected=[];
      if(!inputs.some(variable=>variable.Name===fidVariable&&variable.InputType==='Optional'))throw new Error('BW no longer supports querying without a FID list.');
      for(const name of Object.keys(selections))if(!inputs.some(variable=>variable.Name===name))throw new Error('Missing BW variable: '+name);
      const variables=inputs.map(variable=>{
        if(variable.VariableType!=='DimensionMemberVariable')throw new Error('Unsupported BW variable: '+variable.Name);
        // Keep the query defaults; the optional FID list is deliberately left empty.
        let selected=variable.Name===fidVariable?{}:structuredClone(variable.Values||{});
        const selection=selections[variable.Name];
        if(selection){
          const dimension=cube.Dimensions.find(item=>item.Name===variable.DimensionReference?.Name),field=variable.Values?.Selection?.SetOperand?.FieldName||dimension?.KeyAttribute||dimension?.DefaultKeyAttribute;
          if(!dimension?.Attributes.some(attribute=>attribute.Name===field))throw new Error('Missing BW variable key: '+variable.Name);
          const element={Comparison:selection.to&&selection.to!==selection.from?'BETWEEN':'=',Low:selection.from};if(element.Comparison==='BETWEEN')element.High=selection.to;
          selected={Selection:{SetOperand:{FieldName:field,Elements:[element]}}};expected.push({field,element});
        }
        if(variable.InputType!=='Optional'&&!selected.Selection?.SetOperand?.Elements?.length)throw new Error('Missing mandatory BW variable: '+variable.Name);
        // This BW interprets an explicit false as exclusion; omission means inclusion.
        for(const operand of operands(selected.Selection))for(const element of operand.Elements||[])if(element.IsExcluding===false)delete element.IsExcluding;
        return {Name:variable.Name,Values:selected};
      });
      for(const step of ['VariableDefinition','VariableSubmit']){
        const response=await request('GetResponse',{Analytics:{...analytics({Variables:variables}),ProcessingDirectives:{ProcessingStep:step},Expand:['Cube']},Options:['StatefulServer']},step);
        if(step==='VariableSubmit'){cube=response.Cube;if(!cube)throw new Error('BW returned no submitted query.');source.InstanceId=cube.DataSource?.InstanceId||source.InstanceId}
      }
      const applied=[cube.FixedFilter,cube.DynamicFilter].flatMap(filter=>operands(filter?.Selection));
      for(const {field,element}of expected){
        const elements=applied.filter(operand=>operand.FieldName===field).flatMap(operand=>operand.Elements||[]);
        if(!elements.length||elements.some(value=>value.IsExcluding||value.LowIs==='Variable'||value.HighIs==='Variable'||(value.Comparison||'=')!==element.Comparison||String(value.Low).replaceAll('-','')!==element.Low||element.High!==undefined&&String(value.High).replaceAll('-','')!==element.High))throw new Error('BW did not confirm the requested range / selections.');
      }
      const rowNames=[...new Set(columns.map(column=>column[1]))];
      for(const [,name,attribute]of columns)if(!cube.Dimensions.some(dimension=>dimension.Name===name&&dimension.Attributes.some(field=>field.Name===attribute)))throw new Error('Missing required BW field: '+attribute);
      const dimensions=cube.Dimensions.filter(dimension=>['Rows','Columns','Free'].includes(dimension.AxisDefault)).map(dimension=>{
        const Axis=rowNames.includes(dimension.Name)?'Rows':dimension.AxisDefault==='Columns'?'Columns':'Free',item={Name:dimension.Name,Axis};
        if(Axis==='Free')return item;
        const attributes=Axis==='Rows'?columns.filter(column=>column[1]===dimension.Name).map(column=>column[2]):[dimension.KeyAttribute||dimension.DefaultKeyAttribute];
        item.Attributes=[...new Set(attributes)].map(Name=>({Name,Obtainability:'Always'}));item.FieldLayoutType='FieldBased';
        if(Axis==='Rows')item.ResultStructure=[{Result:'Members',Visibility:'Visible'},{Result:'Total',Visibility:'Hidden'}];
        const readMode=dimension.DefaultResultSetReadMode||dimension.ResultSetReadMode;if(readMode&&readMode!=='Undefined')item.ReadMode=item.ResultSetReadMode=readMode;
        return item;
      });
      const definition={Dimensions:dimensions,Query:{Axes:[{Axis:'Rows',Type:1},{Axis:'Columns',Type:2}].map(axis=>({...axis,ZeroSuppressionType:0,Layout:dimensions.filter(dimension=>dimension.Axis===axis.Axis).map(dimension=>dimension.Name)}))},ResultSetFeatureRequest:{ResultFormat:'Version2',ResultEncoding:'None',ResultKeepOriginalTexts:true,IgnoreAutoSubmit:true,UseDefaultAttributeKey:false,SubSetDescription:{RowFrom:0,RowTo:1000,ColumnFrom:0,ColumnTo:1}}};
      if(key==='ctl'||lookup){
        const selections=[];
        if(cube.DynamicFilter?.Selection)selections.unshift(cube.DynamicFilter.Selection);
        if(lookup){
          const [,name,field]=lookupColumn,attribute=cube.Dimensions.find(dimension=>dimension.Name===name)?.Attributes.find(item=>item.Name===field);
          if(!attribute||attribute.Filterable===false||attribute.IsFilterable===false)throw new Error('BW cannot filter '+lookup.field+' in this query.');
          selections.push({SetOperand:{FieldName:field,Elements:[...targets].map(Low=>({Comparison:'=',Low}))}});
        }else selections.push({SetOperand:{FieldName:'CRCCHCAT.KEY',Elements:['015','016','017','018','014'].map(Low=>({Comparison:'=',Low}))}});
        definition.Filter={Selection:selections.length===1?selections[0]:{Operator:{Code:'And',SubSelections:selections}}};
      }
      const rows=[];let total=null,pages=0,blankDates=0,skippedRows=0;
      for(let offset=0;;){
        onProgress?.('Loading BW InA · '+(total===null?0:Math.min(99,Math.floor(offset*100/total)))+'% · '+offset+(total===null?'':' / '+total)+' rows');
        Object.assign(definition.ResultSetFeatureRequest.SubSetDescription,{RowFrom:offset,RowTo:offset+1000});
        const response=await request('GetResponse',{Analytics:analytics(definition),Options:['StatefulServer']},'ResultSet',120000),axis=response.Grids?.[0]?.Axes?.find(axis=>String(axis.Type||axis.Name).toUpperCase()==='ROWS');
        if(!axis||!Number.isInteger(axis.TupleCount)||axis.TupleCount<0||axis.TupleCount>1000||!Number.isInteger(axis.TupleCountTotal)||axis.TupleCountTotal<offset+axis.TupleCount)throw new Error('Incomplete BW result page.');
        if(total!==null&&total!==axis.TupleCountTotal)throw new Error('BW result count changed while paging.');total=axis.TupleCountTotal;pages++;if(!total)break;
        const vectors=columns.map(([column,name,attribute,isDate])=>{
          const index=axis.Dimensions.findIndex(dimension=>dimension.Name===name),dimension=axis.Dimensions[index],field=dimension?.Attributes.find(item=>item.Name===attribute);
          if(!field||!axis.Tuples[index])throw new Error('BW omitted required field: '+column);
          return {column,isDate,indexes:values(axis.Tuples[index].MemberIndexes),types:values(dimension.MemberTypes),data:values(field.Values)};
        });
        for(let index=0;index<axis.TupleCount;index++){
          const row={};
          for(const vector of vectors){
            const member=vector.indexes[index];if(!Number.isInteger(member)||member<0||member>=vector.data.length||vector.types[member]!==0)throw new Error('Invalid BW detail row / member index.');
            const value=vector.data[member],day=vector.isDate?date(value):value;
            if(vector.isDate&&day===null)blankDates++;
            row[vector.column]=vector.isDate?day||'':value;
          }
          // Skip rows outside the requested scope instead of discarding the complete read.
          const created=String(row['Create Date']||'').replaceAll('-','');
          if(key==='ctl'&&(created&&(created<from||created>to)||!lookup&&!['Push','Config','Slot','Pull','Drop Slot'].includes(row['Change Cat']))){skippedRows++;continue}
          if(lookup){const value=String(row[lookup.field]??'').trim();if(!targets.has(quoteLookup?value.replace(/^00(?=\d{8}$)/,''):value)){skippedRows++;continue}}
          rows.push(row);
        }
        offset+=axis.TupleCount;if(offset===total)break;if(!axis.TupleCount)throw new Error('BW paging stopped before the full result.');
      }
      if(total&&!rows.length)throw new Error('BW returned no rows within the requested scope.');
      const notes=[blankDates?blankDates+' unreadable date'+(blankDates===1?'':'s')+' left blank':'',skippedRows?skippedRows+' row'+(skippedRows===1?'':'s')+' outside the requested scope skipped':''].filter(Boolean);
      result={rows,fields:columns.map(column=>column[0]),pages,from,to,...(notes.length?{warning:'Data loaded; '+notes.join('; ')+'.'}:{})};return result;
    }finally{
      if(created)try{await request('GetResponse',{Analytics:{Actions:[{Type:'Close',DataSources:[source.InstanceId]}]},Options:['StatefulServer']},'Close',10000,true)}catch{if(result)result.cleanupError='Data loaded; BW query close was not confirmed. The server may need to reclaim the session.'}
    }
  }
  return {loginUrl:base+'GetServerInfo'+suffix,read,range};
}

// Source parsing is independent of dashboard state and does not change source records.
function fidSourceDefinitions(){
  const field=(label,aliases=[])=>({label,aliases:[label,...aliases]});
  const ctlFields={
    fid:field('Opportunity ID',['FCID','FID']),number:field('CTL ID',['CTL#','CTL Number']), change:field('Change Cat',['Change Type','Change Category','CTL Change Type']),mcsd:field('MCSD @ Current',['MCSD','Current MCSD']),mcsdCreate:field('MCSD@Create',['MCSD At Create']),rdCreate:field('RD At Create'),mcsdApproval:field('MCSD At SOps Approval'),tsdApproval:field('TSD @ SOPs Approval'),
      srd:field('SRD @ Current',['Sales Ops RD','Current SRD','SRD']), slotSrd:field('SRD @ SOPs Approval',['SRD at SOPs Approval']),approved:field('1st SOps Appr Dt',['Sales Ops Approval Date','SOPS Approval Date']), created:field('Create Date',['Creation Date']),
      initiator:field('Initiator',['Created By','Initiated By','Requester']),businessUnit:field('Business Unit',['BU','Bus Unit']),modified:field('Last Modified Date',['Modified Date']), summary:field('Change Summary',['Summary']),mfg:field('Mfg Upgrade Type',['Manufacturing Upgrade Type']),
      stage:field('Stage',['CTL Stage']),quote:field('Quote ID',['Quote Number']),customer:field('Customer group 2',['Customer','Customer Name','End Customer']),configFrz:field('Config Frz@Current',['Config Frz','Config Frozen Date','Config Freeze Date']),rootCause:field('Root Cause',['Root Cause of Late Change']),notes:field('Notes',['CTL Notes']),subReason:field('Sub Reason')
  };
  return {
    fcid:{label:'FID Database',sheet:'BOBJ',fields:{
      fid:field('FID',['FCID','Field ID']),commit:field('Commit: Field',['Commit Field','Field Commit','Commit Status']),
      account:field('Account Group',['Account Team','Sales Territory','AT']),businessUnit:field('Business Unit',['BU']),fab:field('Fab Name',['Fab']),fabSegment:field('Fab Segment',['Fab_Segment']), customer:field('Customer',['Customer Name','End Customer']),
      description:field('FID Description'),system:field('System Description',['System Desc','System']),quote:field('Active Quote Number',['Quote Number','Active Quote']),crd:field('Date: Customer Request',['Customer Request Date']),targetShip:field('Date: Target Ship',['Target Ship Date','Target Ship']),mcsd:field('Date: Mfg Commit Ship',['Mfg Commit Ship',
        'Manufacturing Commit Ship Date','MCSD']), rd:field('Date: Sales Ops Request Date (RD)',['Sales Ops Request Date','Sales Ops RD','SRD','RD']),mfg:ctlFields.mfg
    }},
    ctl:{label:'CTL Report',sheet:'CTL BOBJ',fields:ctlFields},
    ctlNsr:{label:'CTL-NSR List',sheet:'export',fields:{...ctlFields, fid:field('FCID',['FID','Opportunity ID']),change:field('Change Type',['Change Cat','CTL Change Type']),mcsd:field('MCSD',['MCSD @ Current']),srd:field('Sales Ops RD',
      ['SRD','SRD @ Current','Field CRD']), approved:field('Sales Ops Approval Date',['1st SOps Appr Dt','SOPS Approval Date']),stage:field('Stage',['CTL Stage']), quote:field('Quote Number',['Quote ID']),sourceQuote:field('Source Quote',['Source Quote Number'])
    }},
    nsr:{label:'NSR Log',sheet:'Whole list',fields:{
      fid:field('Reference Fcid',['REF FCID','Reference FID','FCID','FID']),number:field('NSR#',['NSR #','NSR','NSR Number','NSR No','NSR ID']), status:field('NSR Status',['Status']),workflow:field('Workflow Status',['Workflow']),task:field('Task Name',['Task','Workflow Task']),
      title:field('NSR Title',['Title']),leadTime:field('NSR LT',['NSR_LT','LT','Lead Time']),submitted:field('Submit Date',['Submitted Date','Submission Date']), customer:field('Customer',['Customer Name']),product:field('Primary Product',['Product']),tempId:field('Temp Id#',
        ['Temp ID','TempId'])
    }},
    nce:{label:'NCE / FA Review',sheet:'Export',fields:{
      fid:field('FCID',['FID']),number:field('NCe #',['NCE Number','NCE#']),state:field('NCe State',['State','Status']), impact:field('Impact',['NCE Impact']),title:field('Title',['NCE Title']),opened:field('Open Date',['Opened Date','Creation Date']),
      damageGroup:field('Damage Group'),damageCode:field('Damage Code'),partNumber:field('Part_Number',['Part Number']), partDescription:field('Part_Description',['Part Description']),customer:field('Customer',['Customer Name'])
    }}
  };
}

function createFidSourceReaders(sources){
  const ctlWeights={fid:[35,1],number:[25,1],change:[10,0],mcsd:[5,0],srd:[5,0],approved:[5,0]};
  const weights={
    fcid:{fid:[40,1],commit:[20,1],mcsd:[5,0],customer:[5,0],system:[5,0],targetShip:[5,0],rd:[5,0],account:[5,0]},
    ctl:{...ctlWeights,summary:[5,0]},
    ctlNsr:{...ctlWeights,stage:[5,0]},
    nsr:{number:[70,1],status:[10,0],task:[10,0],title:[10,0]},
    nce:{fid:[35,1],number:[30,1],state:[10,0],impact:[10,0],title:[5,0],opened:[5,0],damageGroup:[3,0],damageCode:[2,0]}
  };

  function clean(value){let valueText=String(value??'').replace(/^\uFEFF/,'').trim().replace(/\.0$/,'');return !valueText||valueText==='#'||valueText.toLowerCase()==='nan'||valueText==='Not assigned'?'':valueText;}
  function norm(value){return String(value??'').toLowerCase().replace(/[^a-z0-9]/g,'');}

  const profiles=Object.fromEntries(Object.entries(sources).map(([key,source])=>{
    const fields=Object.entries(source.fields).map(([name,field])=>({label:field.label,aliases:field.aliases.map(norm),weight:weights[key]?.[name]?.[0]||0,required:Boolean(weights[key]?.[name]?.[1])}));
    return [key,{label:source.label,sheet:source.sheet,fields,key:source.fields.fid.label,number:source.fields.number?.label,minimum:65}];
  }));
  function profileFor(key){const profile=profiles[key];if(!profile)throw new Error('Unknown data source: '+key);return profile;}

  function scoreSourceHeader(values,key){
    const profile=profileFor(key),headers=new Set((values||[]).map(norm).filter(Boolean)),missing=[];let score=0,exactScore=0;
    profile.fields.forEach(field=>{if(field.aliases.some(alias=>headers.has(alias))){score+=field.weight;if(headers.has(norm(field.label)))exactScore+=field.weight}else if(field.required)missing.push(field.label);});
    // Keep the existing coverage threshold before rewarding native column names.
    return {score:score<profile.minimum?score:score+exactScore,missing,requiredOk:!missing.length};
  }
  function betterHeader(candidate,previous){return !previous||Number(candidate.requiredOk)>Number(previous.requiredOk)||(candidate.requiredOk===previous.requiredOk&&candidate.score>previous.score);}
  function bestHeader(matrix,key,startRow=0){
    let best=null;(matrix||[]).slice(0,50).forEach((values,index)=>{const candidate={...scoreSourceHeader(values,key),headerRow:startRow+index};if(betterHeader(candidate,best))best=candidate;});
    return best||{score:0,missing:profileFor(key).fields.filter(field=>field.required).map(field=>field.label),requiredOk:false,headerRow:startRow};
  }
  function sourceScoreError(key,candidate){const error=new Error(profileFor(key).label+': check file headers'+(candidate?.missing?.length?' ('+candidate.missing.join(', ')+')':''));error.code='LOW_SOURCE_SCORE';error.userMessage=error.message;return error;}
  function validateCandidate(key,candidate){if(!candidate||!candidate.requiredOk||candidate.score<profileFor(key).minimum)throw sourceScoreError(key,candidate);return candidate;}
  function detectSource(workbook,key){
    const profile=profileFor(key),candidates=[];
    (workbook.SheetNames||[]).forEach((sheetName,sheetOrder)=>{
      const sheet=workbook.Sheets[sheetName];if(!sheet||!sheet['!ref'])return;
      const fullRange=XLSX.utils.decode_range(sheet['!ref']),previewRange={s:{r:fullRange.s.r,c:fullRange.s.c},e:{r:Math.min(fullRange.e.r,fullRange.s.r+49),c:fullRange.e.c}};
      const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',blankrows:true,raw:false,range:previewRange}),best=bestHeader(matrix,key,fullRange.s.r);
      candidates.push({...best,sheet:sheetName,dataRows:Math.max(0,fullRange.e.r-best.headerRow),configured:sheetName===profile.sheet,sheetOrder});
    });
    candidates.sort((a,b)=>Number(b.requiredOk)-Number(a.requiredOk)||b.score-a.score||Number(b.configured)-Number(a.configured)||b.dataRows-a.dataRows||a.sheetOrder-b.sheetOrder);
    return validateCandidate(key,candidates[0]);
  }
  function canonicalize(rows,key){const profile=profileFor(key),headers=new Map(Object.keys(rows[0]||{}).map(name=>[norm(name),name])),columns=profile.fields.map(field=>[field.label,field.aliases.map(alias=>headers.get(alias)).find(Boolean)]);return rows.map(source=>{const row={...source};for(const [label,alias] of columns)if(!clean(row[label])&&String(row[label]??'').trim().toLowerCase()!=='not assigned'&&alias)row[label]=source[alias];return row}).filter(row=>clean(row[profile.key])||(key==='nsr'&&clean(row[profile.number])));}
  function parseWorkbook(bytes,key){
    const workbook=XLSX.read(bytes,{type:'array'}),candidate=detectSource(workbook,key),rows=canonicalize(XLSX.utils.sheet_to_json(workbook.Sheets[candidate.sheet],{defval:'',range:candidate.headerRow}),key);
    if(!rows.length)throw new Error('The worksheet contains no data rows.');
    return {rows,sheet:candidate.sheet,headerRow:candidate.headerRow+1,score:candidate.score,fields:Object.keys(rows[0])};
  }
  let reader=null,readSequence=0;const pendingReads=new Map();
  function readWorkbook(bytes,key){
    profileFor(key);
    // The normal userscript path always uses the shared resource-backed worker.
    if(typeof createWorkbookWorker!=='function')return Promise.resolve().then(()=>parseWorkbook(bytes,key));
    if(!reader){
      const functions=[clean,norm,profileFor,scoreSourceHeader,betterHeader,bestHeader,sourceScoreError,validateCandidate,detectSource,canonicalize,parseWorkbook];
      const code='const profiles='+JSON.stringify(profiles)+';\n'+functions.map(fn=>fn.toString()).join('\n')+';self.onmessage=event=>{const {id,bytes,key}=event.data;try{self.postMessage({id,ok:true,...parseWorkbook(bytes,key)})}catch(error){self.postMessage({id,ok:false,message:error.message,code:error.code,userMessage:error.userMessage})}};';
      reader=createWorkbookWorker(code);
      reader.onmessage=event=>{const result=event.data,waiting=pendingReads.get(result.id);if(!waiting)return;pendingReads.delete(result.id);if(result.ok){const {id,ok,...parsed}=result;waiting.resolve(parsed);}else waiting.reject(Object.assign(new Error(result.message),{code:result.code,userMessage:result.userMessage}));};
      reader.onerror=event=>{pendingReads.forEach(waiting=>waiting.reject(new Error(event.message||'Workbook reader failed.')));pendingReads.clear();reader.terminate();reader=null;};
    }
    return new Promise((resolve,reject)=>{const id=++readSequence,buffer=ArrayBuffer.isView(bytes)?bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength):bytes;pendingReads.set(id,{resolve,reject});try{reader.postMessage({id,bytes:buffer,key},[buffer]);}catch(error){pendingReads.delete(id);reject(error);}});
  }
  function csvToMatrix(content,delimiter){
    const rows=[];let row=[],value='',quoted=false;
    for(let index=0;index<content.length;index++){
      const character=content[index],next=content[index+1];
      if(character==='"'&&quoted&&next==='"'){value+='"';index++;}else if(character==='"')quoted=!quoted;
      else if(character===delimiter&&!quoted){row.push(value);value='';}
      else if((character==='\r'||character==='\n')&&!quoted){if(row.length||value){row.push(value);rows.push(row);row=[];value='';}if(character==='\r'&&next==='\n')index++;}else value+=character;
    }
    if(quoted)throw new Error('CSV contains an unclosed quoted value.');
    if(row.length||value){row.push(value);rows.push(row);}return rows;
  }
  function rowsFromMatrix(matrix,headerRow){const headers=(matrix[headerRow]||[]).map(clean);return matrix.slice(headerRow+1).filter(row=>row.some(value=>clean(value))).map(row=>Object.fromEntries(headers.map((header,index)=>[header,row[index]??''])));}
  function decodeText(buffer){const bytes=new Uint8Array(buffer);if(bytes[0]===255&&bytes[1]===254)return new TextDecoder('utf-16le').decode(bytes);if(bytes[0]===254&&bytes[1]===255)return new TextDecoder('utf-16be').decode(bytes);try{return new TextDecoder('utf-8',{fatal:true}).decode(bytes)}catch{return new TextDecoder('gb18030').decode(bytes)}}
  function csvDelimiter(text){const probe=text.split(/\r?\n/).slice(0,50);return [',','\t',';','|'].map(delimiter=>{let score=0;for(const line of probe){let quoted=false;for(const char of line){if(char==='"')quoted=!quoted;else if(char===delimiter&&!quoted)score++}}return {delimiter,score}}).sort((a,b)=>b.score-a.score)[0].delimiter}
  async function readFile(file,key){
    const extension=file.name.split('.').pop().toLowerCase();
    if(extension==='json'&&key==='nsr'){const cache=JSON.parse(decodeText(await file.arrayBuffer()));if(!Array.isArray(cache.rows))throw new Error('No NSR cache rows.');return {rows:cache.rows,sheet:'NSR cache',from:cache.from,to:cache.to,savedAt:cache.savedAt}}
    if(extension!=='csv'){if(!['xlsx','xls','xlsm'].includes(extension))throw new Error('Use Excel, CSV or NSR JSON.');return readWorkbook(await file.arrayBuffer(),key)}
    const content=(typeof file.arrayBuffer==='function'?decodeText(await file.arrayBuffer()):await file.text()).replace(/^\uFEFF/,''),matrix=csvToMatrix(content,csvDelimiter(content)),candidate=validateCandidate(key,{...bestHeader(matrix,key),sheet:'CSV',dataRows:matrix.length}),rows=canonicalize(rowsFromMatrix(matrix,candidate.headerRow),key);
    return {rows,sheet:'CSV',headerRow:candidate.headerRow+1,score:candidate.score,fields:Object.keys(rows[0]||{})};
  }
  return {readFile,readWorkbook};
}

// Storage-compatible Fab mapping; UI and conflict navigation remain with the caller.
function createFidAccountMapping(){
  const defaults={
    "AT1":`
*Alpha & Omega Semi*
*Hua Hong*
*Huahong*
*Huali*
ADVANCED SEMICONDUCTOR MANUFAC
Analog Foundries (CUMEC)
ASMC*
BYD*Semiconductor*
Cansemi*
China Resources Microelectronics
Chongqing SIMIC Semiconductor*
CR Micro*
CR Runan Technolog*
CSMC*
CSR Zhuzhou
Dingtai Jiangxin Tec
Galaxycore*
GTA Semiconductor*
Guangzhou Cansemi*
Guangzhou Zen Semico*
Hangzhou Fullsemi*
Hangzhou Hesheng*
Hangzhou Power Innovation Microelectron
HH - Fab *
HH Grace
Hikstor Technology
Hubei Sanan Optoelectronics*
Hunan Jiechuwei Semiconductor*
Hunan Sanan Semicon*
Innostar Semi*
IRay Technology*
Jiangsu Advanced Memory (AMS)
Jiangsu Nexic*
Jie Jie Microelectro*
Jiechuwei
Jinan BYD Semiconduc*
Maxscend Microelectr*
Resources West Micro*
Runpeng Semicon*
Sanan-ST Microelectronics
Shanghai Dingtai Jiangxin Tech*
Shanghai GTX Semicon*
Shanghai Huayaoxin*
Shanghai IC
Shanghai Integrated Circuit *
Shanghai Optoelectro*
Skysilicon Co. (China Resources Micro>)
SST Microelectronics
StarPower Technolog*
Sunny Omn*light*
Swaysure Technology*
TI - Chengdu
UMC - HJTC Fab 1
Union Semiconductor*
Wuyuan*
Xi'an Longwei Semico*
YOFC Advanced Semi*
Zhejiang ChipRight*
Zhejiang Hichip Micr*
Zhejiang Hikstor*
Zhejiang ICsprout Semiconductor
Zhejiang Xinsheng Semiconductor
Zhuzhou CRRC Times Electric
Zing Semiconductor*`,
    "AT2":`
Beijing Electronics*
Beijing Hongsheng Ro
Beijing Yandong Micr
BEMC
GoerPixel
Guangdong Ziyue Semi*
Hong Kong Microelec
HSRX
Ji Guang Semiconduct*
Jinan GoerPixel Tech*
MRDI
Ningbo Semiconductor
PST*
PXW Semiconductor
Quanxin Integrated Circuit Mfg-QXIC
Rong Semiconductor*
SEC Chip*
Semicon Manuf Oriental*
Semiconductor Manufacturing*
Shanghai Xin Gang In
Shenzhen Pengxinxu Technology Co., Ltd
Shenzhen Pensun Technology Co*
Shenzhen Qili Tech
SiEn*Qingdao*
Silex*
Sinovate*
SMBC B3
SMEC*
SMIC*
United Nova Technology*
UNT*
XGIC
YDMT
Yongxin*
Z*hai Hengqin Jointe*`,
    "AT3":`
*HuiXinLongGuan*
Anhui Juhe*
ASP Tech*
Beijing Jidian Holding*
Changxin*
Chengdu Hi*tech Jin*
CHJS (Chengdu High-tech Jin Science)
Chongqing lnnoevsic*
CXMT Fab2*
Dongguan Guangmao*
Hangzhou HFC Semicon*
He*fei Xinjing*
Hefei Jingyi Semicon*
Hefei Semiri Tech*
HFC Semiconductor Corporation
Innoscience*
Innotron*
Jiangsu Silicon Integrity (Forecast Only
Nexchip Semi*
Nexmemory Technology*
UNIM Intelligent Mem*
UNITED MICROELECTRONICS CORPOR
USC - Xiamen 12X
Wuhan Chuxing Techno*
Wuhan HSMC
Wuhan Xinxin Semi*
Xiamen ITRI/USC
Xinlian Micro*
XMC
Yangtze Advanced Mem*
YMTC*
Zhe Jiang Super Dim*`,
    "Hynix":`
SK hynix*`,
    "Intel":`
Intel - China F68*
INTEL SEMI STG TECH (DL)`,
    "OSAT":`
*Tongfu*
*TRS Microelectronics
Amkor*
Anhui Semi Integr Display Technology
ATX*
Beijing H*F*I*Elect*
Beijing ICLink Tech*
Beijing Xinli Tech*
BOE Technology*
Chengdu ECHINT Techn*
Chipmore*Tech*
DSHX*Shanghai*
Forehope*
Great*Bay Chipl*
Guangdong 3D-Semi*
He*fei Payton Storag*
Hsingchen Technology*
Huaian Imag*
Huatian Tech*
Hubei Hsinghui Inte*
Hubei Yangtze*
Huizhou Biwin Storage Technology(FCOnly)
Innosemi*
JCET*
Jiangsu Silicon Inte*
Jiangyin Changdian A*
Kore Semiconductor*
Kunming BOE Display
Luxshare*
Natural-Integration Advanced
NCAP*
Quliang Electronics*
Shan Dong Qi Xin Mems Corp
Shanghai Echip Semiconductor*
Shanghai Xianfang Semiconductor
Shanghai Xinyuan Inn*
Shenzhen Qili Tech F-PJT (Forecast Only)
Sidtek Anhui Semicon
SILICONWARE TECHNOLOGY (SUZHOU
SJ Semi*
Sky Chip Inter*
SPIL*
Suzhou TF-AMD Semiconductor
TF AMD Micr*
Tianfu Xinglong Lake*
Unisem*
Wuhu Micro*Display*
Xiamen Sky Semicond*
Yibu Semiconductor*
Yongjiang Lab*
Yunnan Invensight Optoelectronics
Zhejiang Hexin Integ*
Zhejiang ICTruth*
Zhuhai Tiancheng Adv
Ziyi Microelectronics*`,
    "Samsung":`
Samsung*`,
    "Training Center":`
Lam - Shanghai Training Center
Lam Research
LAM RESEARCH INTERNATIONAL SDN`,
    "TSMC":`
TSMC - *`
  };
  const key='fcid.ag.v1',legacyKey='toolkitKyra.fcid.accounts.v1';let mapping={...defaults},notice='',conflicts=[],error='';
  let matchName;
  function clean(value){let text=String(value??'').trim().replace(/\.0$/,'');return !text||text==='#'||text.toLowerCase()==='nan'||text==='Not assigned'?'':text;}
  function fabKey(value){return clean(value).toUpperCase().replace(/\s+/g,' ');}
  function normalize(value){
    if(!value||typeof value!=='object'||Array.isArray(value)||Object.values(value).some(entry=>typeof entry!=='string'))throw new Error('Invalid Fab mapping');
    const result={},seen=new Map(),duplicates=[];
    Object.entries(value).forEach(([group,lines])=>{
      group=clean(group);if(!group)return;const keep=[];
      [...new Set(lines.split(/\r?\n/).map(clean).filter(Boolean))].forEach(fab=>{const normalized=fabKey(fab),previous=seen.get(normalized);if(previous&&previous!==group)duplicates.push(fab+' ('+previous+' / '+group+')');else if(!previous){seen.set(normalized,group);keep.push(fab);}});result[group]=keep.join('\n');
    });
    if(duplicates.length)throw new Error('Duplicate Fab mappings found:\n\n'+duplicates.slice(0,12).join('\n')+'\n\nEach Fab must belong to only one Account Group.');return result;
  }
  function entries(value){return new Map(Object.entries(normalize(value)).flatMap(([group,lines])=>lines.split('\n').filter(Boolean).map(fab=>[fabKey(fab),{fab,group}])));}
  function merge(base,saved){
    const before=entries(base),user=entries(saved),current=entries(defaults),merged=Object.fromEntries(Object.keys(defaults).map(group=>[group,[]])),found=[];
    new Set([...current.keys(),...user.keys(),...before.keys()]).forEach(fab=>{
      const original=before.get(fab),custom=user.get(fab),latest=current.get(fab),edited=custom?.group!==original?.group;
      if(edited&&latest?.group!==original?.group&&custom?.group!==latest?.group)found.push((custom||latest||original).fab+': saved '+(custom?.group||'[Unmapped]')+'; updated default '+(latest?.group||'[Unmapped]'));
      const chosen=edited?custom:latest;if(chosen)(merged[chosen.group]??=[]).push(chosen.fab);
    });
    const next=Object.fromEntries(Object.entries(merged).map(([group,fabs])=>[group,fabs.join('\n')])),savedMap=normalize(saved),savedMatch=matcher(savedMap),nextMatch=matcher(next);
    // A broader default must not silently restore a name or rule the user removed.
    const restored=[...before].filter(([key,{fab}])=>!user.has(key)&&nextMatch(fab)&&savedMatch(fab)!==nextMatch(fab));
    if(restored.length){found.push(...restored.map(([,entry])=>entry.fab+': saved '+(savedMatch(entry.fab)||'[Unmapped]')+'; updated default '+nextMatch(entry.fab)));return {mapping:savedMap,conflicts:found}}
    return {mapping:next,conflicts:found};
  }
  function websiteStorage(){try{return typeof persistentStore==='function'?persistentStore():unsafeWindow.localStorage;}catch{try{return localStorage;}catch{return null;}}}
  function readLegacy(name){try{return websiteStorage()?.getItem(name)??null;}catch{return null;}}
  function persist(base,value){
    try{GM_setValue(key,JSON.stringify({base,mapping:value}));}catch(problem){error=problem.message||String(problem);return false;}
    // Only these two superseded mapping keys belong to this migration.
    try{const storage=websiteStorage();storage?.removeItem(key);storage?.removeItem(legacyKey);}catch{/* GM copy is already saved. */}return true;
  }
  function matcher(value){
    const exact=new Map(),patterns=[],matches=new Map();
    Object.entries(value).forEach(([group,lines])=>lines.split('\n').filter(Boolean).forEach(fab=>{
      if(fab.includes('*'))patterns.push({expression:new RegExp('^'+fab.split('*').map(part=>part.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('.*')+'$','i'),group,length:fab.replace(/\*/g,'').length});else exact.set(fabKey(fab),group);
    }));patterns.sort((a,b)=>b.length-a.length);
    return fab=>{
      const name=clean(fab);
      if(!matches.has(name))matches.set(name,exact.get(fabKey(fab))||patterns.find(pattern=>pattern.expression.test(name))?.group||'');
      return matches.get(name);
    };
  }
  function install(value){mapping=normalize(value);matchName=matcher(mapping)}
  function getStatus(){return {notice,conflicts:[...conflicts],error};}
  function result(ok){return {ok,map:{...mapping},...getStatus()};}
  function unpackMapping(candidate){
    const data=typeof candidate==='string'?JSON.parse(candidate):candidate,base=data?.base||null,user=data?.mapping||data;
    return {base,user,...(base?merge(base,user):{mapping:normalize(user),conflicts:[]})};
  }
  function append(value){
    const saved=GM_getValue(key,null),current=saved===null?mapping:unpackMapping(saved).mapping;
    const localRules=entries(current),groups=new Map(Object.entries(current).map(([group,lines])=>[group,lines.split('\n').filter(Boolean)]));
    for(const [name,{fab,group}] of entries(value)){
      if(localRules.has(name))continue;
      if(!groups.has(group))groups.set(group,[]);
      groups.get(group).push(fab);
    }
    return Object.fromEntries([...groups].map(([group,fabs])=>[group,fabs.join('\n')]));
  }
  function reload(){
    notice='';conflicts=[];error='';install(defaults);
    try{
      const saved=GM_getValue(key,null),candidates=saved===null?[readLegacy(key),readLegacy(legacyKey)]:[saved];
      for(const candidate of candidates){
        if(candidate===null||candidate===undefined)continue;
        try{
          const merged=unpackMapping(candidate),{base,user}=merged;
          install(merged.mapping);conflicts=merged.conflicts;
          notice=base?(conflicts.length?'Updated defaults conflict with your saved mapping. Saved assignments are kept for now. Review the mappings below, then click Save to confirm, or Restore Defaults.\n\n'+conflicts.join('\n'):''):'This saved mapping has no update baseline. Review it and click Save to keep it, or Restore Defaults to use the latest defaults.';
          // Unresolved conflicts retain the old baseline until explicit Save or Reset.
          if(saved===null)persist(base,user);else if(base&&!notice&&JSON.stringify(base)!==JSON.stringify(defaults))persist(defaults,mapping);
          return result(true);
        }catch{/* Try the next valid legacy record without overwriting malformed data. */}
      }
    }catch(problem){error=problem.message||String(problem);}return result(!error);
  }
  function save(value){
    error='';let next;try{next=normalize(value);}catch(problem){error=problem.message||String(problem);return result(false);}
    if(!persist(defaults,next))return result(false);install(next);notice='';conflicts=[];return result(true);
  }
  function reset(){return save(defaults);}
  reload();
  return {key,defaults:{...defaults},getMap:()=>({...mapping}),getStatus,match:fab=>matchName(fab),normalize,append,save,reset,reload};
}

function createFidGraph(raw, options = {}) {
  const steps=createFidGraphSteps(raw,options);
  if(options.yield)return (async()=>{
    let step,until=performance.now()+12;
    do{step=steps.next();if(!step.done&&performance.now()>=until){if(await options.yield()===false)return null;until=performance.now()+12}}while(!step.done);
    return step.value;
  })();
  let step;do{step=steps.next()}while(!step.done);return step.value;
}
function fidLookupMatches(value){
  // Match complete IDs in copied cells or prose, including the numeric part of an unprefixed NSR.
  // Keep the token boundaries in sync with KyraLookupSelection() in DefaultHotkeys.ahk.
  const pattern=/(?<![A-Z0-9_&+-])(?<![0-9]\.)(?:(FID|CTL|NSR|Quote(?:[ \t]+ID)?)[ \t]*[:：#-]?[ \t]*)?([NDL][0-9]{5,8}|[0-9]{5,10})(?![A-Z0-9_&+-]|\.[0-9])/gi;
  return String(value??'').matchAll(pattern);
}
function fidLookupIdentifiers(value){
  return [...new Set(Array.from(fidLookupMatches(value),([,,number])=>number.toUpperCase()))];
}
function packFidSnapshot(snapshot,raw){
  // Keep calculated relationships and refer back to the five source caches for record values.
  const rows=Object.values(raw).flat(),rowIds=new Map(rows.map((row,index)=>[row,index]));
  const mergedRecords=new Set([...snapshot.ctls,...snapshot.nsrs].map(entity=>entity.record).filter(record=>!rowIds.has(record)));
  // Keep shared links and evidence objects shared when IndexedDB clones the graph.
  const copies=new Map();
  function copy(value){
    if(value===null||typeof value!=='object')return value;
    const rowId=rowIds.get(value);if(rowId!==undefined)return rowId;
    if(mergedRecords.has(value))return null;
    const existing=copies.get(value);if(existing)return existing;
    const result=Array.isArray(value)?[]:{};
    copies.set(value,result);
    for(const key of Object.keys(value))result[key]=copy(value[key]);
    return result;
  }
  return {...copy(snapshot),sourceRowRefs:true};
}
function* restoreFidSnapshotRows(snapshot,raw,mergeRows){
  const rows=Object.values(raw).flat(),nodes=snapshot.fids.map(([,node])=>node);
  const entities=new Set([...nodes,...snapshot.ctls,...snapshot.nsrs,...snapshot.nces,
    ...nodes.flatMap(node=>node.nsrs),...snapshot.nsrs.flatMap(nsr=>nsr.observations)]);
  function rowAt(index){
    if(!Number.isSafeInteger(index)||index<0||index>=rows.length)throw new Error('Invalid FID cache row.');
    return rows[index];
  }
  let batch=0;
  for(const entity of entities){
    if(++batch%128===0)yield;
    for(const list of [entity.rows,entity.rawRows])if(list)for(let index=0;index<list.length;index++){
      if(typeof list[index]==='number')list[index]=rowAt(list[index]);
    }
    if(entity.record===null)entity.record=mergeRows(entity.rows);
    else if(typeof entity.record==='number')entity.record=rowAt(entity.record);
  }
}
function* createFidGraphSteps(raw, options) {
  // The synchronous API and responsive page build use the same calculation, yielding between small batches.
  let batch=0;
  const snapshot=options.snapshot;
  const sourceKeys = ['fcid','ctl','ctlNsr','nsr','nce'];
  const input = Object.fromEntries(sourceKeys.map(key => [key, !snapshot&&Array.isArray(raw?.[key]) ? raw[key] : []]));
  const sourceText=value=>{const text=String(value??''),trimmed=text.trim();return !trimmed||/^(?:#|nan|null|undefined|\[undefined\]|n\/a)$/i.test(trimmed)?'':text};
  const cleanedValues=new Map();
  const clean = value => {
    if(value===undefined||value===null||value==='')return '';
    const cached=cleanedValues.get(value);if(cached!==undefined)return cached;
    const text = String(value ?? '').trim();
    const result=/^(?:#|nan|not assigned|null|undefined|\[undefined\]|n\/a)$/i.test(text) ? '' : text;cleanedValues.set(value,result);return result;
  };
  const norm = value => String(value ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const textKey = value => clean(value).toLowerCase();
  const headerIndexes=new WeakMap(),headerSchemas=new Map(),aliasKeys=new Map();
  const field = (row, names, original=false) => {
    if (!row) return '';
    for (const name of names) if (Object.hasOwn(row,name)) {const value=original?sourceText(row[name]):clean(row[name]);if(value)return value;}
    let headers=headerIndexes.get(row);if(!headers){const keys=Object.keys(row),schema=keys.join('\0');headers=headerSchemas.get(schema);if(!headers){headers=new Map();for(const key of keys){const canonical=norm(key),variants=headers.get(canonical)||[];variants.push(key);headers.set(canonical,variants)}headerSchemas.set(schema,headers)}headerIndexes.set(row,headers)}
    for(const name of names){if(!aliasKeys.has(name))aliasKeys.set(name,norm(name));for(const key of headers.get(aliasKeys.get(name))||[]){const value=original?sourceText(row[key]):clean(row[key]);if(value)return value}}
    return '';
  };
  const validFid = value => { const id = clean(value); return id && !['NOFID','NOFCID','NA','NULL','UNKNOWN','NOTASSIGNED'].includes(norm(id).toUpperCase()) ? id : ''; };
  const loaded = Object.fromEntries(sourceKeys.map(key => [key, options.loadedSources ? (Array.isArray(options.loadedSources) ? options.loadedSources.includes(key) : !!options.loadedSources[key]) : snapshot?!!snapshot.loaded[key]:input[key].length > 0]));
  const ownerPolicy='majority';
  const effectivePolicy='approved';
  const riskPolicy='approved';
  const dropLeadWeeks=20,dropReportDay=16,pullLeadWeeks=16; // Inferred reporting windows for historical quarter membership.
  const ignoredNsrs = new Set(Array.from(options.ignoredNsrs||[],value=>clean(value).toUpperCase())),importantNsrs = new Set(Array.from(options.importantNsrs||[],value=>clean(value).toUpperCase()));
  const parsedDates=new Map(),isoDates=new Map(),quarterDates=new Map();
  const date = value => {
    const text = clean(value); if (!text) return null;
    if(parsedDates.has(text))return parsedDates.get(text);
    let result;
    if (value instanceof Date) result = new Date(value.getFullYear(), value.getMonth(), value.getDate());
    else if (/^\d{5}(?:\.\d+)?$/.test(text)) { const serial = new Date((Math.floor(Number(text)) - 25569) * 86400000); result = new Date(serial.getUTCFullYear(), serial.getUTCMonth(), serial.getUTCDate()); }
    else { const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/); result = iso ? new Date(Number(iso[1]), Number(iso[2])-1, Number(iso[3])) : new Date(text); }
    const parts=text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/),us=text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s.*)?$/);
    if (Number.isNaN(result.getTime())||parts&&(result.getFullYear()!==Number(parts[1])||result.getMonth()!==Number(parts[2])-1||result.getDate()!==Number(parts[3]))||us&&(result.getFullYear()!==Number(us[3])||result.getMonth()!==Number(us[1])-1||result.getDate()!==Number(us[2]))) {parsedDates.set(text,null);return null;}
    result.setHours(0,0,0,0); parsedDates.set(text,result);return result;
  };
  const iso = value => {const key=clean(value);if(isoDates.has(key))return isoDates.get(key);const day = date(value),result=day ? [day.getFullYear(),String(day.getMonth()+1).padStart(2,'0'),String(day.getDate()).padStart(2,'0')].join('-') : '';isoDates.set(key,result);return result;};
  const today=iso(options.today||new Date());
  const firstDate = (row, names) => { for (const name of names) { const day = iso(field(row,[name])); if (day) return {date:day,field:name}; } return {date:'',field:''}; };
  const weeks = (from, to) => { const a = date(from), b = date(to); return a && b ? (Date.UTC(b.getFullYear(),b.getMonth(),b.getDate()) - Date.UTC(a.getFullYear(),a.getMonth(),a.getDate())) / 604800000 : null; };
  const personNames=new Map(),learnedAliases=new Map(),aliasEvidence=new Map();
  const person = value => {
    const name=clean(value),map=options.personMap;if(!name)return '';
    if(map instanceof Map){const explicit=map.get(name)||map.get(norm(name));if(explicit)return clean(explicit)}
    else if(map){const key=Object.keys(map).find(key=>norm(key)===norm(name));if(key)return clean(map[key])||name}
    return learnedAliases.get(norm(name))||personNames.get(norm(name))||name.replace(/\s+/g,' ');
  };
  const completed = status => /\bCOMPLETED\b/i.test(status);
  const nsrStatus=nsr=>[nsr.status,nsr.task,nsr.taskStatus].filter(Boolean).join(' ');
  const nsrNamesByFid=new Map(),authorsByFid=new Map(),knownUserIds=new Set(),aliasCandidates=new Map();
  for(const source of ['ctl','ctlNsr'])for(const row of input[source]){
    if(++batch%128===0)yield;
    const account=field(row,['Initiator','Created By','Initiated By','Requester']),change=norm(field(row,['Change Cat','Change Type','Change Category']));
    if(!account)continue;if(/^[A-Z0-9_]+$/i.test(account))knownUserIds.add(norm(account));
    if(change&&change!=='slot'&&change!=='config')continue;
    const fid=validFid(field(row,['Opportunity ID','FCID','FID']));if(!fid)continue;
    const accounts=authorsByFid.get(fid)||new Set();accounts.add(norm(account));authorsByFid.set(fid,accounts);
  }
  for(const row of input.nsr){
    if(++batch%128===0)yield;
    if(!field(row,['NSR#','NSR Number','Nsrnum']))continue;
    const name=field(row,['Initiator','Initiated By','Requester']),fid=validFid(field(row,['Reference Fcid','REF FCID','Reference FID','FCID','FID']));
    if(name&&!/^[A-Z0-9]+$/.test(name)&&!(!/\s/.test(name)&&knownUserIds.has(norm(name)))){personNames.set(norm(name),name.replace(/\s+/g,' '));if(fid){const names=nsrNamesByFid.get(fid)||new Set();names.add(norm(name));nsrNamesByFid.set(fid,names)}}
  }
  let observation=0;
  for(const [fid,accounts] of authorsByFid){
    const names=nsrNamesByFid.get(fid);if(accounts.size!==1||names?.size!==1)continue;
    const account=[...accounts][0];if(!knownUserIds.has(account))continue;
    const name=[...names][0],candidates=aliasCandidates.get(account)||new Map(),candidate=candidates.get(name)||{name,fids:[],last:0};
    candidate.fids.push(fid);candidate.last=++observation;candidates.set(name,candidate);aliasCandidates.set(account,candidates);
  }
  for(const [account,candidates] of aliasCandidates){
    // Each FID counts once. Most frequent name wins; a tie uses the last observation.
    const ranked=[...candidates.values()].sort((a,b)=>b.fids.length-a.fids.length||b.last-a.last),winner=ranked[0];if(winner.fids.length<2)continue;
    const name=personNames.get(winner.name);learnedAliases.set(account,name);
    aliasEvidence.set(account,{account,name,fids:winner.fids,aliases:ranked.filter(candidate=>candidate.fids.length>=2&&candidate.fids.length*3>=winner.fids.length).map(candidate=>personNames.get(candidate.name)),alternatives:ranked.slice(1).map(candidate=>({name:personNames.get(candidate.name),fids:candidate.fids}))});
  }
  // Merge recurring name variants without renaming another account's primary name.
  const primaryNames=new Set([...learnedAliases.values()].map(norm));
  for(const evidence of aliasEvidence.values())for(const alias of evidence.aliases)if(!primaryNames.has(norm(alias))||norm(alias)===norm(evidence.name))personNames.set(norm(alias),evidence.name);
  const mergeRows = rows => {
    const result = {};
    for (const row of rows) for (const key of Object.keys(row)) if (!clean(result[key]) && sourceText(row[key])) result[key] = row[key];
    return result;
  };
  if(snapshot?.sourceRowRefs)yield* restoreFidSnapshotRows(snapshot,raw,mergeRows);
  const unique = values => [...new Set(values.filter(Boolean))];
  const accountMap=options.accountMap||{},accountEntries=accountMap instanceof Map?[...accountMap]:Object.entries(accountMap);
  const accountOf = record => {
    const explicit = field(record,['Account Group','Account Team','Sales Territory','AT']);
    if (explicit && !/^(?:0|00:00:00|#(?:VALUE!|REF!|NAME\?|NUM!|DIV\/0!|NULL!))$/i.test(explicit)) return explicit;
    // Fab rules are more specific; fall back to the source Customer only when no Fab rule matches.
    for(const name of unique([field(record,['Fab Name','Fab']),field(record,['Customer','Customer Name','End Customer'])])){
      if(typeof accountMap==='function'){const account=clean(accountMap(name,record));if(account)return account;continue}
      const matches=[];
      for(const [account,patterns]of accountEntries)for(const pattern of Array.isArray(patterns)?patterns:String(patterns).split(/\r?\n/)){
        const text=clean(pattern);if(!text)continue;
        const expression=new RegExp('^'+text.split('*').map(part=>part.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('.*')+'$','i');
        if(expression.test(name))matches.push({account,specificity:text.replace(/\*/g,'').length,exact:!text.includes('*')});
      }
      matches.sort((a,b)=>Number(b.exact)-Number(a.exact)||b.specificity-a.specificity);
      if(matches.length)return matches[0].account;
    }
    return '';
  };
  const exclusion = record => {
    const system = field(record,['System Description','System Desc','System']).toUpperCase();
    return /NCFUPG/.test(system) ? 'NCFUPG' : /SW/.test(system)&&/UPG|UPGRADE/.test(system) ? 'SW-UPGRADE' : /SOFTWARE/.test(system) ? 'Software' : /CUST[_ ]?BUS[_ ]?FLEX/.test(system) ? 'Cust_Bus_Flex' : /CUST[_ ]?INCENTIVE/.test(system) ? 'Cust_Incentive' : '';
  };
  const fids = new Map(snapshot?.fids), ctls = snapshot?.ctls||[], nsrs = snapshot?.nsrs||[], nces = snapshot?.nces||[];
  const orphans = snapshot?.orphans||{nsrs:[],ctls:[],nces:[],unmatchedNsrs:[]};
  const ensureFid = id => {
    if (!id) return null;
    if (!fids.has(id)) fids.set(id,{id,fidMark:['ignore','important'].includes(options.fidMarks?.[id])?options.fidMarks[id]:'normal',record:{},rawRows:[],sources:[],hasNativeRecord:false,inPopulation:false,excludedReason:'',account:'',businessUnit:'',customer:'',system:'',fidDescription:'',commit:'',quote:'',quarter:'',mcsd:'',rd:'',crd:'',targetShip:'',ctls:[],nsrs:[],nces:[],slot:null,owner:'',ownerEvidence:null,salesAdmin:'',salesAdminEvidence:null,pre:null,post:null,risks:[],category:'review'});
    return fids.get(id);
  };
  const fidRows = new Map();
  for (const row of input.fcid) {
    if(++batch%128===0)yield;
    const id = validFid(field(row,['FID','FCID','Field ID']));
    if (!id) continue;
    if (!fidRows.has(id)) fidRows.set(id,[]); fidRows.get(id).push(row);
  }
  for (const [id,rows] of fidRows) {
    if(++batch%128===0)yield;
    const node = ensureFid(id), commits = rows.filter(row=>field(row,['Commit: Field','Commit Field','Field Commit','Commit Status']).toUpperCase()==='COMMIT'), candidates = commits.length ? commits : rows;
    const selected = candidates.find(row=>field(row,['Fab Segment','Fab_Segment'])) || candidates[0];
    node.record = selected; node.rawRows = [...rows]; node.hasNativeRecord = true; node.sources.push('fcid');
    node.excludedReason = exclusion(selected); node.inPopulation = commits.length>0&&!node.excludedReason;
  }
  const ctlGroups = new Map();
  for (const source of ['ctl','ctlNsr']) for (let index=0;index<input[source].length;index++) {
    if(++batch%128===0)yield;
    const row=input[source][index], number=field(row,['CTL ID','CTL#']), key=number || '__'+source+'-row-'+index;
    if (!ctlGroups.has(key)) ctlGroups.set(key,[]); ctlGroups.get(key).push({source,row,index});
  }
    const ctlValueFields=Object.entries({
      fidId:['Opportunity ID','FCID','FID'],change:['Change Cat','Change Type','Change Category'],initiator:['Initiator','Created By','Initiated By','Requester'],businessUnit:['Business Unit','BU','Bus Unit'],
      summary:['Change Summary','Summary'],created:['Creation Date','Create Date'],approved:['1st SOps Appr Dt','Sales Ops Approval Date','SOPS Approval Date'],mcsd:['MCSD @ Current','MCSD'],mcsdCreate:['MCSD@Create','MCSD At Create'],rdCreate:['RD At Create'],mcsdApproval:['MCSD At SOps Approval'],tsdApproval:['TSD @ SOPs Approval'],
      srd:['SRD @ Current','Sales Ops RD','SRD'],srdApproval:['SRD @ SOPs Approval'],fieldCrd:['Field CRD','CRD @ Current','CRD'],mfg:['Mfg Upgrade Type','Manufacturing Upgrade Type'],stage:['Stage','CTL Stage'],quote:['Quote ID','Quote Number'],sourceQuote:['Source Quote','Source Quote Number'],
      customer:['Customer group 2','Customer','Customer Name','End Customer'],configFrz:['Config Frz@Current','Config Frz','Config Frozen Date','Config Freeze Date'],rootCause:['Root Cause','Root Cause of Late Change'],reason:['Reason for Chg','Reason for Change'],salesOpsApprover:['Sales OPs Approver']
    });
  const ctlDateFields=['created','approved','mcsd','mcsdCreate','rdCreate','mcsdApproval','tsdApproval','srd','srdApproval','fieldCrd','configFrz'];
  const ctlSourceFields=new Set(['businessUnit','customer','quote','sourceQuote','rootCause','reason','salesOpsApprover']);
  // BOBJ truncates Sub Reason at 20 characters; match stable prefixes and retain raw values.
  const lateCtlSubReasonPrefixes=['Application Change','Build Type Change (N','Fab Change – Config','Install Base Issue R','New Feature CIP','Demand Reallocation','Common SPEC/Safety'];
  for (const [key,entries] of ctlGroups) {
    if(++batch%128===0)yield;
    const rows = entries.map(entry=>entry.row), read = (names,original=false) => { for (const entry of entries) { const value=field(entry.row,names,original); if (value) return value; } return ''; };

    const reference=entries.find(entry=>entry.source==='ctlNsr')?.row;
    const ctl={id:key,number:read(['CTL ID','CTL#']),record:rows.length===1?rows[0]:mergeRows(rows),rows,rawRows:rows,sources:unique(entries.map(entry=>entry.source)),conflicts:[],links:[],changeSummary:field(reference,['Change Summary','Summary'],true)||read(['Change Summary','Summary'],true),notes:field(reference,['Notes','CTL Notes'],true)||read(['Notes','CTL Notes'],true),subReason:field(reference,['Sub Reason'],true)||read(['Sub Reason'],true)};
    for (const [name,names] of ctlValueFields) {
      const original=ctlSourceFields.has(name),alternatives=[];
      let selected=(name==='rootCause'&&field(reference,names,true))||'';
      for(const entry of entries){
        const value=field(entry.row,names,original);if(!selected)selected=value;
        if(entries.length>1){const alternative=original?field(entry.row,names):value;if(alternative&&!alternatives.includes(alternative))alternatives.push(alternative)}
      }
      ctl[name]=selected;
      if(alternatives.length>1){const comparison=alternatives.map(value=>ctlDateFields.includes(name)?iso(value)||textKey(value):textKey(value));if(unique(comparison).length>1)ctl.conflicts.push({field:name,values:alternatives})}
    }
    ctl.fidId=validFid(ctl.fidId); ctl.initiator=person(ctl.initiator); for(const name of ctlDateFields)ctl[name]=iso(ctl[name]);
    // Every CTL source uses Stage; Status columns remain only in the original source rows.
    ctl.stage=ctl.stage.replace(/^(?:\d+\s*-\s*)?(?:originator|draft)$/i,'Draft');
    const stage=String(ctl.stage||'').trim().replace(/^\d+\s*-\s*/,'').toLowerCase();
    ctl.status=stage&&stage!=='[unknown]'?({completed:'Completed',cancelled:'Cancelled',canceled:'Cancelled',rejected:'Rejected',draft:'Draft'})[stage]||'Submitted':'';
    ctl.changeKey=norm(ctl.change); ctl.statusKind=ctl.status==='Completed'?'completed':['Cancelled','Rejected'].includes(ctl.status)?'discarded':'pending';
    const subReason=field(reference,['Sub Reason'])||read(['Sub Reason']);
    ctl.lateConfigReason=ctl.changeKey.includes('config')&&read(['Reason for Chg','Reason for Change'])==='Customer Config Chg'&&lateCtlSubReasonPrefixes.some(prefix=>subReason.startsWith(prefix));
    ctl.pending=['slot','config'].includes(ctl.changeKey)&&ctl.status==='Submitted';
    ctl.pendingPullPush=['pull','push'].includes(ctl.changeKey)&&ctl.status==='Submitted';
    ctl.effective=ctl.riskEligible=ctl.applied=ctl.statusKind!=='discarded'&&!!ctl.approved;
    ctl.eventDate=ctl.approved;ctl.eventDateField=ctl.approved?'Sales Ops Approval Date':'';ctl.effectiveDate=ctl.riskDate=ctl.applied?ctl.approved:'';
    const ctlText=Object.values(entries[0].row).join(' ').toUpperCase();
    ctl.isSlot=ctl.changeKey==='slot'; ctl.isDrop=ctl.changeKey==='dropslot'||ctl.changeKey==='drop';
    ctl.isConfig=ctl.changeKey.includes('config')&&!/DOC ONLY|LOCAL NSR|LNSR/.test(ctlText);
    ctls.push(ctl);
    const node=ensureFid(ctl.fidId); if (node) { node.ctls.push(ctl); node.sources=unique([...node.sources,...ctl.sources]); } else orphans.ctls.push(ctl);
  }
  const nsrValues=(record,number)=>({
    account:accountOf(record),status:field(record,['NSR Status','NsrWorkflowStatus','Workflow Status','Status']),initiator:person(field(record,['Initiator','Initiated By','Requester','Created By'])),
    title:field(record,['NSR Title','Title','NsrTitle'],true),type:field(record,['NSR Type','Type'],true),org:field(record,['NSR Org','Org','NsrOrg'],true),product:field(record,['Primary Product','Product','PrimaryProd'],true),customer:field(record,['Customer','Customer Name'],true),subsystem:field(record,['Impacted Sub System','Impacted Subsystem','Sub System','Subsystem'],true),submitted:iso(field(record,['Submit Date','Submitted Date','Submission Date','SubmitDate'])),tempId:field(record,['Temp Id#','Temp ID','TempId']),task:field(record,['Task Name','Task','Workflow Task']),taskStatus:field(record,['Task Status']),
    searchText:[number,...Object.values(record)].join(' ').toLowerCase()
  });
  const nsrGroups = new Map();
  for (let index=0;index<input.nsr.length;index++) {
    if(++batch%128===0)yield;
    const row=input.nsr[index], number=field(row,['NSR#','NSR Number','Nsrnum']).toUpperCase(), key=number||'__nsr-row-'+index;
    if (!nsrGroups.has(key)) nsrGroups.set(key,[]); nsrGroups.get(key).push(row);
  }
  for (const [key,rows] of nsrGroups) {
    if(++batch%128===0)yield;
    const record=mergeRows(rows), number=field(record,['NSR#','NSR Number','Nsrnum']).toUpperCase(), fidIds=unique(rows.map(row=>validFid(field(row,['Reference Fcid','REF FCID','Reference FID','FCID','FID','RefFcid']))));
    const nsr={id:key,number,fidId:fidIds[0]||'',fidIds,record,rows,rawRows:rows,...nsrValues(record,number),eligible:!!number,linkAssessmentAvailable:loaded.ctl||loaded.ctlNsr,linkAssessmentProvenance:loaded.ctl&&loaded.ctlNsr?'both':loaded.ctl?'ctl':loaded.ctlNsr?'ctlNsr-partial':'unavailable',workflow:field(record,['Workflow Status','NSR Status','NsrWorkflowStatus']),links:[]};
    nsr.ignored=ignoredNsrs.has(nsr.number);nsr.important=!nsr.ignored&&importantNsrs.has(nsr.number);
    nsr.observations=rows.length<2?[]:rows.map(record=>({record,...nsrValues(record,number),fidId:validFid(field(record,['Reference Fcid','REF FCID','Reference FID','FCID','FID','RefFcid']))}));
    nsrs.push(nsr);
    if (nsr.fidIds.length) {
      let unmatched=false;
      for(const fidId of nsr.fidIds) {const node=ensureFid(fidId);node.nsrs.push(nsr);node.sources=unique([...node.sources,'nsr']);if(!node.hasNativeRecord)unmatched=true;}
      if(unmatched)orphans.unmatchedNsrs.push(nsr);
    } else orphans.nsrs.push(nsr);
  }
  for (let index=0;index<input.nce.length;index++) {
    if(++batch%128===0)yield;
    const record=input.nce[index], id=field(record,['NCe #','NCE Number'])||'__nce-row-'+index, fidId=validFid(field(record,['FCID','FID']));
    const nce={id,number:field(record,['NCe #','NCE Number']),fidId,record,rows:[record],status:field(record,['NCe State','State','Status']),impact:field(record,['Impact','NCE Impact']),title:field(record,['Title','NCE Title']),customer:field(record,['Customer','Customer Name']),opened:iso(field(record,['Open Date','Opened Date','Creation Date'])),damageGroup:field(record,['Damage Group']),damageCode:field(record,['Damage Code']),partNumber:field(record,['Part_Number','Part Number']),partDescription:field(record,['Part_Description','Part Description'])};
    nces.push(nce); const node=ensureFid(fidId); if (node) { node.nces.push(nce); node.sources=unique([...node.sources,'nce']); } else orphans.nces.push(nce);
  }
  const nsrByNumber=new Map(nsrs.filter(nsr=>nsr.number).map(nsr=>[nsr.number,nsr]));
  const tokenPattern=/(?<![A-Z0-9])([NDL]\d{5,8})(?![A-Z0-9])/gi;
  if(!snapshot)for (const ctl of ctls) {
    if(++batch%128===0)yield;
    const tokens=new Map();
    for (const row of ctl.rows) for (const name of ['Change Summary','Summary','Root Cause','Root Cause of Late Change','Notes','CTL Notes']) {
      const value=field(row,[name]); if (!value) continue;
      for (const match of value.matchAll(tokenPattern)) { const number=match[1].toUpperCase(); if (!tokens.has(number)) tokens.set(number,new Set()); tokens.get(number).add(name); }
    }
    ctl.nsrReferences=[...tokens.keys()];
    ctl.nsrReferenceTypes=Object.fromEntries([...tokens].map(([number,fields])=>[number,fields.has('Change Summary')||fields.has('Summary')?'Chg':'Note']));
    for (const [number,fields] of tokens) {
      const nsr=nsrByNumber.get(number); if (!nsr) continue;
      const kind=ctl.fidId && nsr.fidIds.includes(ctl.fidId) ? 'direct':'reference';
      const link={ctlId:ctl.id,ctlNumber:ctl.number,nsrId:nsr.id,nsrNumber:nsr.number,fidId:ctl.fidId,nsrFidId:nsr.fidId,kind,referenceType:ctl.nsrReferenceTypes[number],fields:[...fields],effective:ctl.effective,applied:ctl.applied,status:ctl.status,statusKind:ctl.statusKind};
      ctl.links.push(link); nsr.links.push(link);
    }
  }
  if(!snapshot)for(const nsr of nsrs){
    if(++batch%128===0)yield;
    nsr.implementationByFid=Object.fromEntries(nsr.fidIds.map(fidId=>[fidId,observationLifecycle(nsr,nsr,fidId)]));
    nsr.lifecycle=observationLifecycle(nsr,nsr);
  }
  const tally = actions => {
    const counts=new Map(); for (const action of actions) if (action.name) { const key=textKey(action.name), entry=counts.get(key)||{name:action.name,count:0,documents:[]}; entry.count++; entry.documents.push({id:action.id,source:action.source}); counts.set(key,entry); }
    return [...counts.values()].sort((a,b)=>b.count-a.count||a.name.localeCompare(b.name));
  };
  // Slot authors establish Account membership; approval and current Slot state do not change who authored an operation.
  const accountPeople=new Map();
  if(!snapshot)for(const node of fids.values()){if(++batch%128===0)yield;node.account=accountOf(node.record)}
  const nsrAccount=(observation,fidId=observation.fidId)=>fids.get(fidId)?.account||observation.account||'';
  if(!snapshot)for(const nsr of nsrs){if(++batch%128===0)yield;nsr.accounts=[...new Set((nsr.observations.length?nsr.observations:[nsr]).map(observation=>nsrAccount(observation)))]}
  if(!snapshot)for(const ctl of ctls){
    if(++batch%128===0)yield;
    const account=fids.get(ctl.fidId)?.account;
    if(!ctl.isSlot||!ctl.initiator||!account)continue;
    const key=textKey(account),people=accountPeople.get(key)||new Map();people.set(textKey(ctl.initiator),ctl.initiator);accountPeople.set(key,people);
  }
  const inferOwner = node => {
    const people=accountPeople.get(textKey(node.account))||new Map();
    const nsrActions=node.nsrs.filter(nsr=>nsr.eligible&&nsr.initiator).map(nsr=>({id:nsr.id,name:nsr.initiator,source:'nsr'}));
    const ctlActions=node.ctls.filter(ctl=>ctl.initiator&&(ctl.isSlot||ctl.changeKey==='config')).map(ctl=>({id:ctl.id,name:ctl.initiator,source:'ctl'}));
    const eligible=action=>people.has(textKey(action.name)),actions=[...ctlActions,...nsrActions];
    // All Account members tie at zero when no eligible CTL votes exist; use their local NSR counts.
    const nsrCandidates=tally(nsrActions.filter(eligible)),ctlCandidates=tally(ctlActions.filter(eligible)),candidates=ctlCandidates.length?ctlCandidates:nsrCandidates;
    const tied=candidates.length>1&&candidates[0].count===candidates[1].count,tiedCandidates=tied?candidates.filter(candidate=>candidate.count===candidates[0].count):[];
    let tieBreaker='',tieBreakerSource='';
    if(tied){
      const ranked=tiedCandidates.map(candidate=>({name:candidate.name,count:nsrCandidates.find(nsr=>textKey(nsr.name)===textKey(candidate.name))?.count||0})).sort((a,b)=>b.count-a.count);
      const finalists=ranked.filter(candidate=>candidate.count===ranked[0].count);
      if(finalists.length===1){tieBreaker=finalists[0].name;tieBreakerSource='nsr-tiebreaker'}
      else if(finalists.some(candidate=>textKey(candidate.name)===textKey(node.slot.initiator))){tieBreaker=node.slot.initiator;tieBreakerSource='slot-tiebreaker'}
    }
    const owner=tied?tieBreaker:candidates[0]?.name||'';
    return {owner,evidence:{policy:ownerPolicy,inferred:true,account:node.account,accountPeople:[...people.values()],source:!people.size?'account-unavailable':!candidates.length?'no-eligible-votes':tied?tieBreakerSource||'unresolved-tie':ctlCandidates.length?'ctl':'nsr-fallback',candidates,nsrCandidates,ctlCandidates,tied,tiedCandidates,tieBreaker,documents:candidates.reduce((sum,candidate)=>sum+candidate.count,0),excludedCandidates:tally(actions.filter(action=>!eligible(action))),aliases:[...aliasEvidence.values()].filter(evidence=>candidates.some(candidate=>candidate.name===evidence.name))}};
  };
  function quarterOf(day){
    if(!quarterDates.has(day))quarterDates.set(day,fidQuarterFromDate(date(day)));
    return quarterDates.get(day);
  }
  function quarterRank(value){
    const match=String(value||'').replace(/[^a-z0-9]/gi,'').match(/^Q(MAR|JUN|SEP|DEC)(\d{2})$/i);
    return match?Number(match[2])*4+['MAR','JUN','SEP','DEC'].indexOf(match[1].toUpperCase()):NaN;
  }
  function assessSlot(node,current,riskActive,nativeFallback=true){
    const ready=loaded.ctl||loaded.ctlNsr,mfg=(riskActive?.mfg||field(node.record,['Mfg Upgrade Type','Manufacturing Upgrade Type'])).toUpperCase(),threshold=mfg.includes('NO LABOR')?12:20;
    const nativeRd=nativeFallback&&!current?.isDrop?node.rd:'',referenceDate=riskActive?.srdApproval||riskActive?.srd||riskActive?.fieldCrd||nativeRd,referenceField=riskActive?.srdApproval?'SRD @ SOPs Approval':riskActive?.srd?'Sales Ops RD':riskActive?.fieldCrd?'Field CRD':nativeRd?'Date: Sales Ops Request Date (RD)':'';
    const span=weeks(riskActive?.eventDate,referenceDate),onTime=span>threshold;
    const pre={state:'unknown',label:'Pre risk unknown',weeks:span,threshold,reason:'CTL source not loaded.',eventDate:riskActive?.eventDate||'',eventDateField:riskActive?.eventDateField||'',referenceDate:referenceDate||'',referenceField,created:riskActive?.created||'',approved:riskActive?.approved||'',ctlId:riskActive?.id||'',policy:riskPolicy};
    if(ready&&!current)Object.assign(pre,{state:'risk',label:'No Slot CTL',reason:'No approved Slot CTL found.'});
    else if(ready&&!riskActive)pre.reason='Original approved Slot evidence is missing after Drop Slot.';
    else if(riskActive&&span===null)pre.reason=current?.isDrop?'Original Slot RD / SRD missing after Drop Slot.':'Slot RD / SRD missing.';
    else if(riskActive)Object.assign(pre,{state:onTime?'clear':'risk',label:onTime?'On time':'Pre late risk',reason:`Slot ${onTime?'submitted on time':'submitted too late'}; (${referenceField==='Field CRD'?'CRD':'SRD'} − Slot approved date) / 7 = ${span.toFixed(1)} weeks ${onTime?'>':'≤'}${threshold} weeks.`});
    return pre;
  }
  function assessPortfolio(target,node,slotted){
    const ready=loaded.ctl||loaded.ctlNsr;
    target.slotted=slotted;target.portfolioStatus=target.post.ctls.length?'late-ctl':target.post.state==='clear'?'stable-ctl':'ctl-unknown';
    target.portfolioStates=!ready?['slot-unknown']:!slotted?['no-slot']:['slot-ready',target.pre.state==='clear'?'stable-slot':target.pre.state==='risk'?'late-slot':'slot-date-unknown',target.portfolioStatus];
    target.portfolioCell=ready&&slotted?target.portfolioStates[1]+'|'+(target.portfolioStatus==='late-ctl'?'late-ctl-nsr':target.portfolioStatus):target.portfolioStates[0];
    if(ready&&!slotted&&node.slot.pending.some(ctl=>ctl.isSlot))target.portfolioStates.push('slot-pending');
    if(ready&&slotted){
      if(node.slot.state==='dropped')target.portfolioStates.push('slot-dropped');
      target.portfolioStates.push(fidRiskAssessment(target.portfolioCell));
      if(!node.mcsd)target.portfolioStates.push('slot-no-mcsd');
    }
    target.risks=[];
    if(target.pre.state==='risk')target.risks.push({type:'pre',label:target.pre.label,reason:target.pre.reason});
    if(target.post.ctls.length)target.risks.push({type:'post',label:'Post config risk',reason:target.post.reason});
    if(target.post.nsrs.length)target.risks.push({type:'nsr',label:'Pending NSR',count:target.post.nsrs.length});
    for(const [label,predicate]of [['Pending Config CTL',ctl=>ctl.pending],['Pending Pull/Push CTL',ctl=>ctl.pendingPullPush]]){
      const count=node.ctls.filter(predicate).length;if(count)target.risks.push({type:'ctl',label,count});
    }
    if(target.portfolioStates.includes('slot-no-mcsd'))target.risks.push({type:'mcsd',label:'MCSD pending',reason:'Slotted FID without MCSD.'});
    if(node.nces.length)target.risks.push({type:'nce',label:'NCE / FA signal',count:node.nces.length});
    target.category=target.pre.state==='risk'||target.post.state==='risk'?'risk':target.pre.state==='unknown'||target.post.state==='unknown'?'review':'clear';
    return target;
  }
  function mergeQuarterStates(node,states){
    const slotted=states.filter(state=>state.slotted),evidence=slotted.length?slotted:states;
    const worst=key=>evidence.find(state=>state[key].state==='risk')?.[key]||evidence.find(state=>state[key].state==='unknown')?.[key]||evidence[0][key];
    const pre=worst('pre'),ctls=[...new Map(evidence.flatMap(state=>state.post.ctls).map(ctl=>[ctl.id,ctl])).values()],post={...worst('post'),ctls,ctlIds:ctls.map(ctl=>ctl.id)};
    return assessPortfolio({pre,post},node,slotted.length>0);
  }
  function projectFid(node,quarters=[],systemQuarters=[]){
    const selected=node.quarterStates.filter(state=>match(state.quarter,quarters)&&match(state.quarter,systemQuarters));if(!selected.length)return null;
    if(node.quarterStates.length===1)return node;
    const assessment=selected.length===1?selected[0]:mergeQuarterStates(node,selected);
    return {...node,...assessment,quarter:selected.map(state=>state.quarter).join(', '),selectedQuarterStates:selected};
  }
  const nativeMcsdOf=record=>firstDate(record,['Date: Mfg Commit Ship','Mfg Commit Ship','Manufacturing Commit Ship Date','MCSD']).date;
  const slotEvents=node=>node.ctls.filter(ctl=>(ctl.isSlot||ctl.isDrop)&&ctl.effective).sort((a,b)=>(b.eventDate||'').localeCompare(a.eventDate||''));
  function refreshFid(node) {
    node.risks=[];
    node.nsrs=node.nsrs.map(nsr=>{
      nsr=nsrByNumber.get(nsr.number)||nsr;
      if(!nsr.observations.length)return nsr;
      const observation=nsr.observations.find(observation=>observation.fidId===node.id);
      return observation?{...nsr,...observation,fidId:node.id,lifecycle:nsr.implementationByFid[node.id]||nsr.lifecycle,globalLifecycle:nsr.lifecycle}:nsr;
    });
    const record=node.record,nativeMcsd=nativeMcsdOf(record);
    node.customer=field(record,['Customer','Customer Name','End Customer'],true)||node.nsrs.find(nsr=>nsr.customer)?.customer||'';
    node.system=field(record,['System Description','System Desc','System'],true); node.quote=field(record,['Active Quote Number','Quote Number','Active Quote']);
    node.fab=field(record,['Fab Name','Fab'],true); node.fidDescription=field(record,['FID Description'],true); node.commit=field(record,['Commit: Field','Commit Field','Field Commit','Commit Status'],true); node.crd=firstDate(record,['Date: Customer Request','Customer Request Date']).date;
    node.businessUnit=field(record,['Business Unit','BU'],true)||node.ctls.find(ctl=>ctl.businessUnit)?.businessUnit||'';
    node.mfg=field(record,['Mfg Upgrade Type','Manufacturing Upgrade Type'],true);
    node.mcsd=nativeMcsd||node.ctls.find(ctl=>ctl.mcsd)?.mcsd||'';
    node.rd=firstDate(record,['Date: Sales Ops Request Date (RD)','Sales Ops Request Date','Sales Ops RD','SRD','RD']).date; node.targetShip=firstDate(record,['Date: Target Ship','Target Ship Date','Target Ship']).date;
    const events=slotEvents(node);
    // Active CTL tracks the latest request independently of the Slot cycle.
    node.activeCtl=node.ctls.filter(ctl=>ctl.created&&(ctl.status==='Submitted'||ctl.status==='Completed'))
      .sort((a,b)=>b.created.localeCompare(a.created)||b.number.localeCompare(a.number,undefined,{numeric:true}))[0]||null;
    const current=events[0]||null,active=current?.isSlot?current:null,riskActive=active||(current?.isDrop?events.find(ctl=>ctl.isSlot):null)||null,pending=node.ctls.filter(ctl=>(ctl.isSlot||ctl.isDrop)&&ctl.statusKind!=='discarded'&&!ctl.effective);
    const quarterDate=nativeMcsd||node.rd,quarterField=nativeMcsd?'Current MCSD':'Current SRD';
    node.quarter=node.currentQuarter=quarterOf(quarterDate);
    node.quarterBasis=[quarterField+': '+(quarterDate||'[Unknown]')];
    const slotReady=loaded.ctl||loaded.ctlNsr;
    node.slot={active,current,pending,riskCurrent:current,riskActive,state:!slotReady?'unknown':current?.isDrop?'dropped':active?'active':'none',initiator:active?.initiator||'',number:active?.number||current?.number||'',quarter:node.quarter,effectivePolicy,riskPolicy,sourceReady:slotReady};
    const ownership=inferOwner(node); node.owner=ownership.owner; node.ownerEvidence=ownership.evidence;
    const adminActions=node.ctls.filter(ctl=>ctl.initiator&&(ctl.changeKey==='pull'||ctl.changeKey==='push')).map(ctl=>({id:ctl.id,name:ctl.initiator,source:'ctl'})),adminCandidates=tally(adminActions);
    const salesAdminTied=adminCandidates.length>1&&adminCandidates[0].count===adminCandidates[1].count;
    node.salesAdmin=salesAdminTied?'':adminCandidates[0]?.name||'';node.salesAdminEvidence={candidates:adminCandidates,tied:salesAdminTied,documents:adminActions.length};
    assessFidQuarters(node,events,nativeMcsd);
    node.weeksToMcsd=node.mcsd?Math.round(weeks(today,node.mcsd)):'';
    node.entitySearchText=[node.id,node.account,node.businessUnit,node.customer,node.system,node.owner,node.salesAdmin,node.quote,...Object.values(record),...node.ctls.flatMap(ctl=>[ctl.number,ctl.quote,ctl.sourceQuote,ctl.summary,ctl.initiator]),...node.nces.flatMap(nce=>[nce.number,nce.title])].join(' ').toLowerCase();
  }
  function assessFidQuarters(node,events,nativeMcsd){
    const record=node.record,{current,active,riskActive,sourceReady:slotReady}=node.slot;
    const effectiveCtls=node.ctls.filter(ctl=>ctl.effective&&ctl.eventDate<=today);
    node.pre=assessSlot(node,current,riskActive);
    const referenceDate=node.pre.referenceDate,referenceField=node.pre.referenceField;
    const postMcsd=current?.isDrop?nativeMcsd||riskActive?.mcsd:node.mcsd;
    const postDate=postMcsd||(current?.isDrop?referenceDate:active?.srd||node.rd||referenceDate),postDateField=postMcsd?'MCSD':current?.isDrop?referenceField:active?.srd?'Sales Ops RD':node.rd?'Date: Sales Ops Request Date (RD)':referenceField;
    // Retain both Stable and Late Config evidence within the existing risk scope.
    const ctlAssessments=node.ctls.filter(ctl=>ctl.lateConfigReason&&ctl.riskEligible).map(ctl=>{
      const referenceDate=ctl.mcsdApproval||ctl.mcsdCreate||postDate||'',referenceField=ctl.mcsdApproval?'MCSD At SOps Approval':ctl.mcsdCreate?'MCSD@Create':postDateField;
      return {ctl,quarter:quarterOf(ctl.mcsdCreate||ctl.rdCreate)||node.currentQuarter,weeks:weeks(ctl.approved,referenceDate),referenceDate,referenceField};
    });
    const qualifying=slotReady?ctlAssessments.filter(({weeks})=>weeks!==null&&weeks<20).map(({ctl})=>ctl):[];
    node.lateCtls=qualifying; // Keep complete event evidence for history, independently of KPI quarter membership.
    const pendingConfigs=node.ctls.filter(ctl=>ctl.isConfig&&ctl.statusKind!=='discarded'&&!ctl.effective);
    // post.nsrs is shared Pending NSR evidence; it does not determine CTL stability.
    const pendingNsrs=loaded.nsr?node.nsrs.filter(nsr=>nsr.eligible&&nsr.lifecycle==='open'):[];
    const groups=new Map([[node.quarter,{quarter:node.quarter,ctls:[],cycles:new Map([[current?.id||'',{current,riskActive}]]),basis:node.quarterBasis.slice(),rd:[],mcsd:[],sources:new Set()}]]);
    const addQuarter=(q,ctl,reason,rd=ctl.rdCreate,mcsd=ctl.mcsdCreate,at=ctl.eventDate)=>{
      if(!q)return null;
      if(!groups.has(q))groups.set(q,{quarter:q,ctls:[],cycles:new Map(),basis:[],rd:[],mcsd:[],sources:new Set()});
      const group=groups.get(q),historical=events.find(event=>event.eventDate<=at)||null,historicalSlot=historical?.isSlot?historical:historical?.isDrop?events.find(event=>event.isSlot&&event.eventDate<=historical.eventDate)||null:null;
      group.cycles.set(historical?.id||'',{current:historical,riskActive:historicalSlot});
      group.basis.push('CTL '+ctl.number+' · '+reason);group.rd.push(rd);group.mcsd.push(mcsd);group.sources.add(ctl.number);
      return group;
    };
    // Drop creation schedules retain historical quarters without replacing the native current schedule.
    for(const ctl of effectiveCtls)if(ctl.isDrop){
      const boundary=events.find(event=>event.id!==ctl.id&&event.eventDate<=(ctl.created||ctl.eventDate))?.eventDate||'';
      const retry=node.ctls.filter(event=>event.isDrop&&event.status==='Rejected'&&event.created>boundary&&event.created<=(ctl.created||ctl.eventDate)&&(event.mcsdCreate||event.rdCreate)).sort((a,b)=>a.created.localeCompare(b.created))[0];
      const original=retry||ctl,created=date(original.created);
      if(!created)continue;
      const report=iso(new Date(created.getFullYear(),created.getMonth()-(created.getDate()<dropReportDay?1:0),dropReportDay));
      const reportingCycle=events.find(event=>event.eventDate<=report),droppingCycle=events.find(event=>event.id!==ctl.id&&event.eventDate<=original.created);
      const observations=[];
      if(!original.mcsdCreate&&original.rdCreate&&reportingCycle?.isSlot&&reportingCycle.id===droppingCycle?.id)for(const event of node.ctls){
        if(event.number===original.number||event.status==='Cancelled')continue;
        if(event.created&&event.created>=reportingCycle.eventDate&&event.created<report&&quarterOf(event.created)===quarterOf(report)&&(event.mcsdCreate||event.rdCreate))observations.push({at:event.created,mcsd:event.mcsdCreate,rd:event.rdCreate,order:0,number:event.number});
        if(event.effective&&event.eventDate>=reportingCycle.eventDate&&event.eventDate<report&&quarterOf(event.eventDate)===quarterOf(report)&&(event.mcsdApproval||event.srdApproval))observations.push({at:event.eventDate,mcsd:event.mcsdApproval,rd:event.srdApproval,order:1,number:event.number});
      }
      observations.sort((a,b)=>b.at.localeCompare(a.at)||b.order-a.order||b.number.localeCompare(a.number));
      const observation=observations[0],observedMcsd=observation?.rd===original.rdCreate?observation.mcsd:'';
      const mcsd=original.mcsdCreate||observedMcsd,snapshot=mcsd||original.rdCreate;
      if(!snapshot)continue;
      const lead=weeks(report,snapshot);
      if(lead===null||lead>dropLeadWeeks)continue;
      addQuarter(quarterOf(snapshot),ctl,'Drop Slot · '+(retry?'original snapshot from CTL '+retry.number+' · ':'')+(original.mcsdCreate?'MCSD@Create':observedMcsd?'Prior MCSD observation from CTL '+observation.number+' at '+observation.at:'RD At Create')+': '+snapshot+'; created '+original.created+'; assumed report '+report+'; report lead '+lead.toFixed(1)+' weeks (<='+dropLeadWeeks+').',original.rdCreate,mcsd);
    }
    const commitRows=node.rawRows.filter(row=>field(row,['Commit: Field','Commit Field','Field Commit','Commit Status']).toUpperCase()==='COMMIT');
    const sourceQuarters=new Set((commitRows.length?commitRows:node.rawRows).map(row=>quarterRank(field(row,['Time: Calendar Quarter']))).filter(Number.isFinite));
    const initialMcsd=firstDate(record,['Date: Initial MCSD at Slot','Initial MCSD at Slot']).date,initialRd=firstDate(record,['Initial SRD at Slot']).date;
    const initialQuarters=new Set([initialMcsd,initialRd].map(day=>quarterRank(quarterOf(day))).filter(Number.isFinite));
    for(const ctl of effectiveCtls)if(ctl.changeKey==='push')for(const [before,after,label]of [[ctl.rdCreate,ctl.srdApproval,'RD At Create → SRD @ SOPs Approval'],[ctl.mcsdCreate,ctl.mcsdApproval,'MCSD@Create → MCSD At SOps Approval']]){
      const q=quarterOf(before),rank=quarterRank(q);
      if(!before||!after||quarterRank(quarterOf(after))<=rank)continue;
      // A future RD abandoned before its quarter began is insufficient when a current MCSD exists.
      if(nativeMcsd&&label.startsWith('RD')&&quarterOf(ctl.mcsdCreate)!==q&&rank>quarterRank(quarterOf(ctl.eventDate)))continue;
      if(rank-quarterRank(quarterOf(ctl.eventDate))<=1||sourceQuarters.has(rank))addQuarter(q,ctl,label+': '+before+' → '+after);
    }
    for(const [createdField,approvalField,label]of [['rdCreate','srdApproval','RD'],['mcsdCreate','mcsdApproval','MCSD']]){
      if(nativeMcsd&&label==='RD')continue;
      const observations=[];
      for(const ctl of effectiveCtls){if(ctl.created&&ctl[createdField])observations.push({at:ctl.created,day:ctl[createdField],ctl,order:0});if(ctl[approvalField])observations.push({at:ctl.eventDate,day:ctl[approvalField],ctl,order:1});}
      observations.sort((a,b)=>a.at.localeCompare(b.at)||a.order-b.order||a.ctl.number.localeCompare(b.ctl.number));
      for(let index=1;index<observations.length;index++){
        const before=observations[index-1],after=observations[index],q=quarterOf(before.day),rank=quarterRank(q);
        if(quarterRank(quarterOf(after.day))<=rank||!sourceQuarters.has(rank)||!initialQuarters.has(rank))continue;
        const rd=label==='RD'?before.day:before.order?before.ctl.srdApproval:before.ctl.rdCreate,mcsd=label==='MCSD'?before.day:before.order?before.ctl.mcsdApproval:before.ctl.mcsdCreate;
        addQuarter(q,after.ctl,'Corroborated '+label+' history: '+before.day+' → '+after.day+'; source Quarter and initial Slot date agree.',rd,mcsd,after.at);
      }
    }
    // Retain additional schedules only once approval reaches their preceding fiscal quarter.
    const initialDate=initialMcsd||initialRd;
    if(riskActive&&initialDate&&quarterRank(quarterOf(initialDate))-quarterRank(quarterOf(riskActive.eventDate))<=1){
      addQuarter(quarterOf(initialDate),riskActive,'Initial Slot '+(initialMcsd?'MCSD':'SRD')+': '+initialDate+'; Slot approved from preceding fiscal quarter onward.',initialRd,initialMcsd);
    }
    for(const ctl of effectiveCtls)if(ctl.changeKey==='pull')for(const [before,after,label]of [[ctl.rdCreate,ctl.srdApproval,'RD At Create → SRD @ SOPs Approval'],[ctl.mcsdCreate,ctl.mcsdApproval,'MCSD@Create → MCSD At SOps Approval']]){
      const q=quarterOf(before),rank=quarterRank(q);
      if(!before||!after||quarterRank(quarterOf(after))>=rank)continue;
      if(nativeMcsd&&label.startsWith('RD')&&quarterOf(ctl.mcsdCreate)!==q)continue;
      if(rank-quarterRank(quarterOf(ctl.eventDate))<=1)addQuarter(q,ctl,'Approved cross-quarter Pull · '+label+': '+before+' → '+after);
    }
    for(const ctl of effectiveCtls)if(ctl.changeKey==='push'&&ctl.mcsdCreate&&quarterRank(quarterOf(ctl.mcsdCreate))-quarterRank(quarterOf(ctl.eventDate))<=1){
      addQuarter(quarterOf(ctl.mcsdCreate),ctl,'Push MCSD@Create: '+ctl.mcsdCreate+'; approved from preceding fiscal quarter onward.');
    }
    for(const ctl of effectiveCtls)if(ctl.changeKey==='pull'&&ctl.mcsdCreate){
      const q=quarterOf(ctl.mcsdCreate),span=weeks(ctl.eventDate,ctl.mcsdCreate);
      if(span===null||span>pullLeadWeeks)continue;
      const support=effectiveCtls.find(other=>other.id!==ctl.id&&other.eventDate<=ctl.eventDate&&(quarterOf(other.mcsdApproval)===q||other.created&&other.created<=ctl.eventDate&&quarterOf(other.mcsdCreate)===q));
      if(support)addQuarter(q,ctl,'Corroborated Pull MCSD@Create: '+ctl.mcsdCreate+'; approved '+ctl.eventDate+'; lead '+span.toFixed(1)+' weeks (<='+pullLeadWeeks+'); supported by CTL '+support.number+'.');
    }
    // Only a confirmed first effective Slot can disprove earlier unslotted forecasts; missing history is not proof.
    const firstSlot=events.filter(ctl=>ctl.eventDate<=today).at(-1);
    if(firstSlot?.isSlot&&!events.some(ctl=>ctl.isDrop&&ctl.eventDate<=firstSlot.eventDate)){
      for(const [quarter,group]of groups)if(quarter!==node.currentQuarter&&[...group.cycles.values()].every(cycle=>!cycle.current))groups.delete(quarter);
    }
    // Classification uses only quarters already supported by schedule evidence; a Late CTL never adds a quarter.
    for(const ctl of qualifying){
      const snapshot=ctl.mcsdCreate||ctl.rdCreate,dateBasis=snapshot||postDate,q=snapshot?quarterOf(snapshot)||node.quarter:node.quarter;
      if(!groups.has(q))continue;
      const group=q?addQuarter(q,ctl,(ctl.mcsdCreate?'MCSD@Create':ctl.rdCreate?'RD At Create':postDateField)+': '+dateBasis):groups.get(q);
      group.ctls.push(ctl);
    }
    node.quarterStates=[...groups.values()].map(group=>{
      const cycles=[...group.cycles.values()],slotted=cycles.filter(cycle=>cycle.current),preCandidates=(slotted.length?slotted:cycles).map(cycle=>cycle.current===current&&group.quarter===node.currentQuarter?node.pre:assessSlot(node,cycle.current,cycle.riskActive,false));
      const pre=preCandidates.find(pre=>pre.state==='risk')||preCandidates.find(pre=>pre.state==='unknown')||preCandidates[0];
      const late=group.ctls,first=late[0],reference=first?.mcsdApproval||first?.mcsdCreate||postDate||'',field=first?.mcsdApproval?'MCSD At SOps Approval':first?.mcsdCreate?'MCSD@Create':postDateField,span=first?weeks(first.eventDate,reference):null;
      const post={state:late.length?'risk':slotReady&&postDate?'clear':'unknown',label:late.length?'Post config risk':slotReady&&postDate?'No post risk':'Post risk unknown',weeks:span,threshold:20,
        reason:late.length?`(${field} − CTL approved date) / 7 = ${span.toFixed(1)} weeks <20 weeks.`:slotReady&&postDate?'No qualifying Late CTL in this quarter.':!slotReady?'CTL not loaded.':'MCSD / SRD missing.',
        eventDate:first?.eventDate||'',referenceDate:reference,referenceField:field,created:first?.created||'',approved:first?.approved||'',ctlIds:late.map(ctl=>ctl.id),ctls:late,assessments:ctlAssessments.filter(assessment=>assessment.quarter===group.quarter),nsrs:pendingNsrs,pending:pendingConfigs,policy:riskPolicy};
      const dates=group.quarter===node.currentQuarter?{rd:node.rd,mcsd:node.mcsd}:{rd:unique(group.rd).sort().join(' / '),mcsd:unique(group.mcsd).sort().join(' / ')};
      const dateSource=group.sources.size?'CTL schedule evidence: '+[...group.sources].join(', ')+'\n'+group.basis.join('\n'):'Current FID / CTL dates';
      return assessPortfolio({quarter:group.quarter,pre,post,basis:group.basis,slotEvidence:preCandidates,dates,dateSource},node,slotted.length>0);
    }).sort((a,b)=>a.quarter.slice(-2).localeCompare(b.quarter.slice(-2))||['Mar','Jun','Sep','Dec'].indexOf(a.quarter.slice(1,4))-['Mar','Jun','Sep','Dec'].indexOf(b.quarter.slice(1,4)));
    Object.assign(node,node.quarterStates.find(state=>state.quarter===node.currentQuarter));
    if(node.quarterStates.length===1)node.selectedQuarterStates=node.quarterStates;else delete node.selectedQuarterStates;
  }
  if(!snapshot)for(const node of fids.values()){if(++batch%128===0)yield;refreshFid(node)}
  if(!snapshot)for(const ctl of ctls){if(++batch%128===0)yield;const node=fids.get(ctl.fidId);ctl.quarter=node?.quarter||'';ctl.quarters=node?.quarterStates.map(state=>state.quarter)||[''];if(!ctl.customer)ctl.customer=field(node?.record||{},['Customer','Customer Name','End Customer'])}
  // A new calendar day changes the displayed weeks and assessments crossed by an effective CTL.
  if(snapshot?.today&&snapshot.today!==today)for(const node of fids.values()){
    if(++batch%128===0)yield;
    node.weeksToMcsd=node.mcsd?Math.round(weeks(today,node.mcsd)):'';
    if(node.ctls.some(ctl=>ctl.effective&&(ctl.eventDate<=today)!==(ctl.eventDate<=snapshot.today))){
      assessFidQuarters(node,slotEvents(node),nativeMcsdOf(node.record));
      const quarters=node.quarterStates.map(state=>state.quarter);
      for(const ctl of node.ctls)ctl.quarters=[...quarters];
    }
  }
  function setNsrMark(number,mark) {
    const nsr=nsrByNumber.get(number);if(!nsr)return false;
    nsr.ignored=mark==='ignore';nsr.important=mark==='important';
    for(const fidId of nsr.fidIds){const row=fids.get(fidId).nsrs.find(row=>row.number===number);if(row)Object.assign(row,{ignored:nsr.ignored,important:nsr.important})}
    return true;
  }
  function setFidMark(id,mark){const node=fids.get(id);if(!node)return false;node.fidMark=['ignore','important'].includes(mark)?mark:'normal';return true}
  const allFids=[...fids.values()],nativeFids=allFids.filter(node=>node.hasNativeRecord&&!node.excludedReason),population=nativeFids.filter(node=>node.inPopulation);
  const values = value => value===undefined||value===null||value===''?[]:(Array.isArray(value)?value:[value]).map(value=>sourceText(value).trim());
  const match = (actual, expected) => { if(expected===undefined||expected===null||expected===''||Array.isArray(expected)&&!expected.length)return true;const text=sourceText(actual).trim().toLowerCase();return (Array.isArray(expected)?expected:[expected]).some(value=>{const term=sourceText(value).trim();return term==='[Unknown]'?!text:text===term.toLowerCase()}); };
  const hasPredicates = object => !!object&&Object.values(object).some(value=>values(value).length);
  const searchMatch = (text, search) => {
    const term=clean(search).toLowerCase(); if (!term) return true;
    const terms=term.split(/[\s,;|/]+/).filter(Boolean),idOnly=terms.length>1&&terms.every(part=>/^(?:\d{5,10}|[ndl]\d{5,8})$/i.test(part));
    return idOnly?terms.some(part=>text.includes(part)):text.includes(term);
  };
  function observationLifecycle(nsr,observation,fidId=''){
    const status=nsrStatus(observation);
    if(/REJECT|CANCEL/.test(status.toUpperCase()))return 'discarded';if(!completed(status))return 'open';if(nsr.number.startsWith('D'))return 'completed-5d';if(!nsr.linkAssessmentAvailable)return 'ctl-unavailable';
    const links=nsr.links.filter(link=>link.kind==='direct'&&(!fidId||link.fidId===fidId)&&link.statusKind!=='discarded');
    return !links.length?'completed-unlinked':links.some(link=>link.applied)?'implemented':links.some(link=>clean(link.status))?'ctl-pending':'ctl-unavailable';
  }
  const matchDateRange=(actual,range)=>{const from=iso(range.from),to=iso(range.to);return (!from&&!to)||!!actual&&(!from||actual>=from)&&(!to||actual<=to)};
  const matchObservation=(nsr,observation,filters={},fidId='')=>{
    if(fidId&&observation.fidId!==fidId)return false;
    return Object.entries(filters).every(([key,expected])=>{
      if(!values(expected).length)return true;
      if(key==='search')return searchMatch(observation.searchText,expected);
      if(key==='submitted')return matchDateRange(observation.submitted,expected);
      if(key==='taskStatus')return values(expected).some(value=>value.startsWith('15-Completed|')?observation.taskStatus==='15-Completed'&&observationLifecycle(nsr,observation,fidId)===value.split('|')[1]:match(observation.taskStatus,value));
      if(key==='lifecycle'){
        const lifecycle=observationLifecycle(nsr,observation,fidId);
        return values(expected).some(value=>value==='important'?nsr.important:value==='ignore'?nsr.ignored:value==='ignore-completed'?nsr.ignored&&completed(nsrStatus(observation)):value==='completed'?lifecycle!=='open'&&lifecycle!=='discarded':value==='rejected'?/REJECT/i.test(nsrStatus(observation)):value==='cancelled'?/CANCEL/i.test(nsrStatus(observation)):match(lifecycle,value));
      }
      return match(key==='number'?nsr.number:observation[key],expected);
    });
  };
  const matchNsr=(nsr,filters={},fidId='')=>nsr.eligible&&(nsr.observations.length?nsr.observations.some(observation=>matchObservation(nsr,observation,filters,fidId)):matchObservation(nsr,nsr,filters,fidId));
  const matchCtl=(ctl,filters={})=>Object.entries(filters).every(([key,value])=>key==='created'&&value&&typeof value==='object'&&!Array.isArray(value)?matchDateRange(ctl.created,value):match(ctl[key],value));
  const matchSignal=(node,signal,includeIgnored=true)=>signal==='nce'?node.nces.length>0:signal==='pending-ctl'?node.ctls.some(ctl=>ctl.pending):signal==='open-nsr'?node.portfolioStates.includes('slot-ready')&&node.post.nsrs.some(nsr=>includeIgnored||!nsr.ignored):['late-ctl','slot-no-mcsd','slot-pending','slot-dropped'].includes(signal)&&node.portfolioStates.includes(signal);
  const isNsrRisk=label=>label==='Pending NSR'||label==='Late NSR'||label==='Open NSR'||label==='Post NSR risk';
  const matchesCtlIdentifier=(value,token)=>clean(value).toLowerCase()===token;
  function matchesEntityIdentifier(node,token){
    return /^\d+$/.test(token)&&(node.id.toLowerCase()===token||node.quote.toLowerCase()===token||
      node.ctls.some(ctl=>[ctl.number,ctl.quote,ctl.sourceQuote].some(value=>matchesCtlIdentifier(value,token))));
  }
  const matchesNsrIdentifier=(nsr,token)=>nsr.number.toLowerCase()===token||nsr.number.replace(/^[NDL]/,'')===token;
  function lookupFids(value,withTarget=false){
    const identifiers=fidLookupIdentifiers(value);
    if(identifiers.length!==1)return [];
    const token=identifiers[0].toLowerCase();
    const matches=allFids.filter(node=>matchesEntityIdentifier(node,token)||node.nsrs.some(nsr=>matchesNsrIdentifier(nsr,token)));
    if(!withTarget)return matches.map(node=>node.id);
    return matches.map(node=>{
      const fid=node.id;
      if(fid.toLowerCase()===token)return {fid};
      if(/^\d+$/.test(token)){
        const ctl=node.ctls.find(ctl=>matchesCtlIdentifier(ctl.number,token));
        if(ctl)return {fid,target:{ctlId:ctl.id,lookupId:token}};
        const direct=node.ctls.filter(ctl=>matchesCtlIdentifier(ctl.quote,token));
        const related=direct.length?direct:node.ctls.filter(ctl=>matchesCtlIdentifier(ctl.sourceQuote,token));
        if(related.length){
          // Prefer the active CTL within the matching quote field, then the highest CTL number.
          const chosen=related.includes(node.activeCtl)?node.activeCtl:related.sort((a,b)=>b.number.localeCompare(a.number,undefined,{numeric:true}))[0];
          return {fid,target:{ctlId:chosen.id,quoteId:token,sourceQuote:!direct.length,lookupId:token}};
        }
        if(node.quote.toLowerCase()===token)return {fid,target:{quoteId:token,lookupId:token}};
      }
      const nsr=node.nsrs.find(nsr=>matchesNsrIdentifier(nsr,token));
      return {fid,target:nsr?{nsrId:nsr.id,lookupId:token}:null};
    });
  }
  function query(filters={},rowsOnly='') {
    const commit=values(filters.commit),defaultCommit=filters.commit===undefined||commit.length===1&&match('Commit',commit);
    const includeIgnored=filters.includeIgnored!==false;
    let base=filters.excluded?allFids.filter(node=>node.excludedReason&&match(node.commit,'Commit')):defaultCommit?population:nativeFids;
    if(!defaultCommit&&commit.length)base=base.filter(node=>match(node.commit,commit));
    if(!includeIgnored)base=base.filter(node=>node.fidMark!=='ignore');
    const ctlRows=hasPredicates(filters.ctl)?ctls.filter(ctl=>matchCtl(ctl,filters.ctl)):ctls,ctlFids=ctlRows===ctls?null:new Set(ctlRows.map(ctl=>ctl.fidId));
    // An absent CTL matches Unknown only when every selected CTL field accepts an empty value.
    const ctlUnknown=!!ctlFids&&matchCtl({},filters.ctl),matchCtlFid=fidId=>!ctlFids||ctlFids.has(fidId)||ctlUnknown&&!fids.get(fidId)?.ctls.length;
    const directKeys=['quarter','currentQuarter','systemQuarter','account','customer','system','businessUnit','mfg','salesAdmin','owner','category','preState','postState','slotState','fidStatus','fidMark','signal','riskType'];
    const hasFidPredicates=!!filters.excluded||!defaultCommit&&commit.length>0||directKeys.some(key=>key!=='account'&&values(filters[key]).length)||hasPredicates(filters.slot);
    const scalarKeys=directKeys.filter(key=>!['quarter','systemQuarter','fidStatus','signal','riskType'].includes(key)&&values(filters[key]).length),fidStatuses=values(filters.fidStatus),signals=values(filters.signal),riskTypes=values(filters.riskType).map(label=>['Post NSR risk','Late NSR','Open NSR'].includes(label)?'Pending NSR':label==='Pending CTL'?'Pending Config CTL':label),nsrRiskOnly=riskTypes.length>0&&riskTypes.every(isNsrRisk),hasNsrPredicates=hasPredicates(filters.nsr);
    const dateChanges=riskTypes.includes('Date Changes')?options.dateChangeIds:null;
    const search=clean(filters.search).toLowerCase(),tokens=fidLookupIdentifiers(search).map(token=>token.toLowerCase()),idSearch=tokens.length>0;
    const entityMatch=node=>idSearch?tokens.some(token=>matchesEntityIdentifier(node,token)):node.entitySearchText.includes(search);
    const nsrSearchMatch=(nsr,observation=nsr)=>!search||(idSearch?tokens.some(token=>matchesNsrIdentifier(nsr,token)):observation.searchText.includes(search));
    const hasAccount=values(filters.account).length>0;
    const inNsrScope=(nsr,observation,fidId='')=>matchObservation(nsr,observation,filters.nsr||{},fidId)&&(!hasAccount||match(nsrAccount(observation,fidId||observation.fidId),filters.account));
    const nsrDomainMatch=(nsr,fidId='')=>nsr.eligible&&(includeIgnored||!nsr.ignored)&&(nsr.observations.length?nsr.observations.some(observation=>inNsrScope(nsr,observation,fidId)&&nsrSearchMatch(nsr,observation)):inNsrScope(nsr,nsr,fidId)&&nsrSearchMatch(nsr));
    const entityHits=search?new Set(allFids.filter(entityMatch).map(node=>node.id)):null,entityHit=fidId=>!entityHits||entityHits.has(fidId);
    const scopedFids=base.flatMap(node=>{
      const scoped=projectFid(node,filters.quarter,filters.systemQuarter);if(!scoped)return [];
      return rowsOnly==='periods'?scoped.selectedQuarterStates.map(state=>projectFid(node,[state.quarter||'[Unknown]'])):[scoped];
    });
    const selectedFids=scopedFids.filter(node=>{
      if(fidStatuses.length&&!fidStatuses.some(status=>status.startsWith('cause:')?matchSignal(node,status.slice(6),includeIgnored):status===node.portfolioCell||node.portfolioStates.includes(status)||node.portfolioCell.endsWith('|'+status)))return false;
      for(const key of scalarKeys){const value=key==='preState'?node.pre.state:key==='postState'?node.post.state:key==='slotState'?node.slot.state:node[key];if(!match(value,filters[key]))return false}
      if(riskTypes.length&&!dateChanges?.has(node.id)&&!node.risks.some(risk=>riskTypes.includes(risk.label)))return false;
      if(signals.length&&!signals.some(signal=>matchSignal(node,signal,includeIgnored)))return false;
      if(!entityHit(node.id)&&!node.nsrs.some(nsr=>nsrDomainMatch(nsr,node.id)))return false;
      if (hasNsrPredicates&&!node.nsrs.some(nsr=>(includeIgnored||!nsr.ignored)&&matchNsr(nsr,filters.nsr,node.id))) return false;
      if (hasPredicates(filters.slot)) { const slot=node.slot.active; if (!slot||!match(field(slot.record,['Initiator','Created By','Initiated By','Requester'],true),filters.slot.initiator)||!node.selectedQuarterStates.some(state=>match(state.quarter,filters.slot.quarter))||!match(slot.status,filters.slot.status)) return false; }
      if (!matchCtlFid(node.id)) return false;
      return true;
    });
    // Overview-only queries skip unrelated projections and metric aggregation.
    if(rowsOnly==='fids')return selectedFids;
    const ids=new Set(selectedFids.map(node=>node.id));
    let selectedCtls;
    if(rowsOnly!=='nsrs'){
      const quartersByFid=new Map();for(const node of selectedFids){if(!quartersByFid.has(node.id))quartersByFid.set(node.id,new Set());for(const state of node.selectedQuarterStates)quartersByFid.get(node.id).add(state.quarter);}
      selectedCtls=ctlRows.filter(ctl=>ids.has(ctl.fidId)).map(ctl=>{const quarters=[...quartersByFid.get(ctl.fidId)];return quarters.length===ctl.quarters.length&&quarters.every(q=>ctl.quarters.includes(q))?ctl:{...ctl,quarters};});
    }
    if(rowsOnly==='ctls')return selectedCtls;
    // CTL conditions use loaded CTL relationships, even outside the loaded FID Database range.
    // The built-in Commit portfolio alone leaves the independent NSR list in scope.
    const selectedNsrs=[];
    for(const nsr of nsrs) {
      if(!nsr.eligible||!includeIgnored&&nsr.ignored)continue;
      const matchedFidIds=nsr.fidIds.filter(fidId=>{
        if(!includeIgnored&&fids.get(fidId)?.fidMark==='ignore')return false;
        if(hasFidPredicates&&!ids.has(fidId))return false;
        if(!matchCtlFid(fidId))return false;
        if(nsrRiskOnly){
          const node=fids.get(fidId),riskNsrs=node.post.nsrs;
          if(!riskNsrs.some(row=>row.number===nsr.number&&row.lifecycle==='open'))return false;
        }
        return entityHit(fidId)?(!hasNsrPredicates&&!hasAccount)||(nsr.observations.length?nsr.observations:[nsr]).some(observation=>inNsrScope(nsr,observation,fidId)):nsrDomainMatch(nsr,fidId);
      });
      const orphan=!nsr.fidId&&!hasFidPredicates&&(!ctlFids||ctlUnknown)&&nsrDomainMatch(nsr);
      if(!matchedFidIds.length&&!orphan)continue;
      if(!nsr.observations.length){selectedNsrs.push({...nsr,taskLifecycle:observationLifecycle(nsr,nsr,nsr.fidId)});continue;}
      const lifecycles=unique(matchedFidIds.map(fidId=>nsr.implementationByFid[fidId]||nsr.lifecycle)),lifecycle=['open','discarded','completed-5d','completed-unlinked','ctl-pending','implemented','ctl-unavailable'].find(state=>lifecycles.includes(state))||nsr.lifecycle;
      const fidId=matchedFidIds[0]||nsr.fidId,observation=nsr.observations.find(observation=>inNsrScope(nsr,observation,fidId)&&(entityHit(fidId)||nsrSearchMatch(nsr,observation)))||nsr.observations[0];
      let taskLifecycle=observationLifecycle(nsr,observation,fidId);
      if(lifecycles.includes('completed-unlinked')&&!['open','discarded'].includes(taskLifecycle))taskLifecycle='completed-unlinked';
      selectedNsrs.push({...nsr,...observation,fidId,lifecycle:hasNsrPredicates?taskLifecycle:lifecycle,taskLifecycle,globalLifecycle:nsr.lifecycle,matchedFidIds,scopeLifecycles:lifecycles});
    }
    if(rowsOnly==='nsrs')return selectedNsrs;
    const selectedNces=nces.filter(nce=>ids.has(nce.fidId));
    const metrics={population:selectedFids.length,atRisk:0,needsReview:0,clear:0,preRisk:0,postRisk:0,bothRisk:0,openNsr:0,completed5dNsr:0,completedUnlinked:0,pendingCtl:0,pendingPullPushCtl:0,workflowPendingCtl:0,nceFids:0,nsr:selectedNsrs.length,ignoredNsr:0,importantNsr:0,ignoredCompletedNsr:0,implementedNsr:0,ctlPendingNsr:0,discardedNsr:0,ctlUnavailableNsr:0,orphanNsr:0,ctl:selectedCtls.length,nce:selectedNces.length,riskFid:0,riskFree:0,riskPending:0,slotNoMcsd:0};
    const categoryMetrics={risk:'atRisk',review:'needsReview',clear:'clear'};
    for(const node of selectedFids){
      metrics[categoryMetrics[node.category]]++;
      metrics.preRisk+=node.pre.state==='risk';metrics.postRisk+=node.post.state==='risk';metrics.bothRisk+=node.pre.state==='risk'&&node.post.state==='risk';metrics.nceFids+=node.nces.length>0;
      metrics.riskFid+=node.portfolioStates.includes('risk-fid');metrics.riskFree+=node.portfolioStates.includes('risk-free');metrics.riskPending+=node.portfolioStates.includes('risk-unknown');metrics.slotNoMcsd+=node.portfolioStates.includes('slot-no-mcsd');
    }
    const lifecycleMetrics={open:'openNsr','completed-5d':'completed5dNsr','completed-unlinked':'completedUnlinked',implemented:'implementedNsr','ctl-pending':'ctlPendingNsr',discarded:'discardedNsr','ctl-unavailable':'ctlUnavailableNsr'};
    for(const nsr of selectedNsrs){const metric=lifecycleMetrics[nsr.lifecycle];if(metric)metrics[metric]++;metrics.ignoredNsr+=nsr.ignored;metrics.ignoredCompletedNsr+=nsr.ignored&&completed(nsrStatus(nsr));metrics.importantNsr+=nsr.important;metrics.orphanNsr+=!nsr.fidId}
    const pendingCtlFids=new Set(),pullPushFids=new Set();
    for(const ctl of selectedCtls){metrics.pendingCtl+=ctl.pending;metrics.pendingPullPushCtl+=ctl.pendingPullPush;metrics.workflowPendingCtl+=ctl.status==='Submitted';if(ctl.pending)pendingCtlFids.add(ctl.fidId);if(ctl.pendingPullPush)pullPushFids.add(ctl.fidId)}
    metrics.pendingCtlFids=pendingCtlFids.size;metrics.pendingPullPushCtlFids=pullPushFids.size;
    return {fids:selectedFids,nsrs:selectedNsrs,ctls:selectedCtls,nces:selectedNces,metrics};
  }
  return {fids,population,nativeFids,observedPopulation:allFids,nsrs,ctls,nces,orphans,query,lookupFids,projectFid,setNsrMark,setFidMark,isNsrRisk,field,getSnapshot:()=>({fids:[...fids],ctls,nsrs,nces,orphans,loaded,today})};
}

function fidQuarterFromDate(day){
  if(!day||Number.isNaN(day.getTime()))return '';
  const year=day.getFullYear(),stamp=Date.UTC(year,day.getMonth(),day.getDate()),march31=new Date(Date.UTC(year,2,31));
  // Lam's extra week belongs to the March quarter; the following quarters each span 13 weeks.
  const marchEnd=march31.getTime()-march31.getUTCDay()*86400000,quarter=Math.max(0,Math.ceil((stamp-marchEnd)/(91*86400000)));
  return ['QMar','QJun','QSep','QDec','QMar'][quarter]+"'"+String(year+(quarter===4?1:0)).slice(-2).padStart(2,'0');
}

function fidRiskAssessment(cell){
  return /late-slot|late-ctl-nsr/.test(cell)?'risk-fid':cell==='stable-slot|stable-ctl'?'risk-free':'risk-unknown';
}

// Shared SAP search and IndexedDB storage for the NSR and FCID dashboards.
function createNsrDataService(onProgress=()=>{}){
  const clean=value=>{const s=String(value??'').replace(/^\uFEFF/,'').trim();return /^(nan|null|undefined)$/i.test(s)?'':s};
const SAP_SEARCH='/sap/opu/odata/sap/ZNSR_CD_SRV/SearchNsrSet';
const SAP_FIELDS={
  'Temp Id#':['TempId'],'NSR#':['Nsrnum'],'NSR Status':['NsrWorkflowStatus','NsrStatus'],'REF FCID':['RefFcid'], 'NSR Request':['NsrRequest'],'Task Name':['DescriptionTask'],'NSR Title':['NsrTitle','EnggTitle'],'NSR Type':['NsrType'], 'Impacted Sub System':['ImpSubSys'],
    'Initiator':['InitiatedByFn','Requester','InitiatedBy'],'Sales Rep':['SalesRepFn'], 'Sales Ops':['SalesOpsFn','SalesOps'],'Technical Contact':['TechContactFn'],'Customer':['Customer'],'Submit Date':['SubmitDate'], 'NSR Org':['NsrOrg'],'Primary Product':['PrimaryProd']
};
const SOURCE_FIELDS=Object.keys(SAP_FIELDS);
const CACHE_DB='toolkitKyra',CACHE_VERSION=1,CACHE_STORE='datasets';let cacheDb;

function openCache(){
  if(!cacheDb)cacheDb=new Promise((resolve,reject)=>{
    const api=unsafeWindow.indexedDB;if(!api){reject(new Error('IndexedDB is unavailable.'));return}
    const request=api.open(CACHE_DB,CACHE_VERSION);let settled=false; request.onupgradeneeded=()=>{const db=request.result;if(!db.objectStoreNames.contains(CACHE_STORE))db.createObjectStore(CACHE_STORE,{keyPath:'key'})};
    request.onsuccess=()=>{const db=request.result;if(settled){db.close();return}settled=true;db.onversionchange=db.onclose=()=>{db.close();cacheDb=null};resolve(db)}; request.onerror=()=>{if(!settled){settled=true;reject(request.error||new Error('Could not open the data cache.'))}};
    request.onblocked=()=>{if(!settled){settled=true;reject(new Error('Data cache is blocked by another page.'))}};
  }).catch(error=>{cacheDb=null;throw error});return cacheDb;
}

function cacheDone(transaction){return new Promise((resolve,reject)=>{transaction.oncomplete=resolve;transaction.onerror=()=>reject(transaction.error||new Error('Cache write failed.'));transaction.onabort=()=>reject(transaction.error||new Error('Cache write cancelled.'))});}

function localDate(value){return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,'0')}-${String(value.getDate()).padStart(2,'0')}`;}
function sapRange(){const today=new Date(),yearAgo=new Date(today.getFullYear()-1,today.getMonth(),today.getDate()); const quarterStart=new Date(yearAgo.getFullYear(),Math.floor(yearAgo.getMonth()/3)*3,1); return {from:localDate(quarterStart),to:localDate(today)};}

function validSapDate(value){
  const match=clean(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!match)return false;
  const year=Number(match[1]),month=Number(match[2]),day=Number(match[3]),date=new Date(Date.UTC(year,month-1,day)); return date.getUTCFullYear()===year&&date.getUTCMonth()===month-1&&date.getUTCDate()===day;
}

function parseSearchXml(text){
  const doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw new Error('SAP returned invalid XML.');
  const rows=[...doc.getElementsByTagNameNS('*','entry')].map(entry=>{const props=entry.getElementsByTagNameNS('*','properties')[0]||entry,row={}; SOURCE_FIELDS.forEach(field=>row[field]=SAP_FIELDS[field].map(name=>clean(props.getElementsByTagNameNS('*',name)[0]?.textContent)).find(Boolean)||''); row['Submit Date']=clean(row['Submit Date']).match(/^(\d{4}-\d{2}-\d{2})/)?.[1]||'';
    return row;}).filter(row=>row['NSR#']);
  const next=[...doc.getElementsByTagNameNS('*','link')].find(link=>clean(link.getAttribute('rel')).toLowerCase()==='next'); return {rows,next:next?clean(next.getAttribute('href')):''};
}

async function fetchSapRows(range,signal){
  const filter=`(Nsrnum eq '' and Screen eq 'S' and Customer eq '' and Description eq '' and ImpSubSys eq '' and NsrRequest eq '' and NsrTitle eq '' and RefFcid eq '' and NsrType eq '' and SalesOpsFn eq '' and EnggTitle eq '' and `+
    `NsrWorkflowStatus eq '' and NsrStatus eq '' and InitiatedByFn eq '' and HostFidUpgrade eq '' and (ResEngGroup eq '') and PrimaryProd eq '' and TechSpec eq '' and (SubmitDate ge datetime'${range.from}T00:00:00' and SubmitDate le datetime'${range.to}T23:59:59') and (NsrOrg eq ''))`;
  const rows=[],seen=new Set();let url=`${SAP_SEARCH}?$filter=${encodeURIComponent(filter)}`,pages=0;
  while(url){
    if(seen.has(url))throw new Error('SAP returned a repeated pagination URL.');
    if(pages>=250)throw new Error('SAP returned too many result pages.');
    seen.add(url);onProgress(range,rows.length); let response;
    try{response=await unsafeWindow.fetch(url,{credentials:'same-origin',cache:'no-store',signal,headers:{Accept:'application/atom+xml,application/xml,text/xml'}})} catch{throw new Error('SAP data request failed. Check the Fiori session and network connection.')}
    const text=await response.text();
    if(nsrSignInResponse(response,text))throw new Error('Session expired. Return to NSR page and sign in again.');
    if(!response.ok)throw new Error(`SAP SearchNsrSet returned HTTP ${response.status}.`);
    const parsed=parseSearchXml(text);rows.push(...parsed.rows);pages++;
    if(parsed.next){const next=new URL(parsed.next,response.url||url);if(next.origin!==unsafeWindow.location.origin)throw new Error('SAP returned an unexpected pagination URL.');url=next.href}else url='';
  }
  return {rows,pages};
}

function cacheResult(request){return new Promise((resolve,reject)=>{request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error||new Error('Cache request failed.'))})}
async function readDataset(key){return cacheResult((await openCache()).transaction(CACHE_STORE,'readonly').objectStore(CACHE_STORE).get(key))}
async function writeDataset(entry){const transaction=(await openCache()).transaction(CACHE_STORE,'readwrite');transaction.objectStore(CACHE_STORE).put(entry);await cacheDone(transaction)}
async function updateDatasets(keys,update){
  const transaction=(await openCache()).transaction(CACHE_STORE,'readwrite'),store=transaction.objectStore(CACHE_STORE),done=cacheDone(transaction);
  try{
    const entries=update(await Promise.all(keys.map(key=>cacheResult(store.get(key)))));
    for(const entry of entries)store.put(entry);
    await done;
    return entries;
  }catch(error){
    try{transaction.abort()}catch{}
    await done.catch(()=>{});
    throw error;
  }
}
async function updateDataset(key,update){const [entry]=await updateDatasets([key],([current])=>[update(current)]);return entry}
async function deleteDatasets(keys){const transaction=(await openCache()).transaction(CACHE_STORE,'readwrite');keys.forEach(key=>transaction.objectStore(CACHE_STORE).delete(key));await cacheDone(transaction)}
return {fields:SOURCE_FIELDS,localDate,sapRange,validSapDate,fetchSapRows,readDataset,writeDataset,updateDataset,updateDatasets,deleteDatasets};
}

function nsrSignInResponse(response,text){
  return [401,403].includes(response.status)||(response.headers.get('Content-Type')||'').toLowerCase().includes('text/html')||/^\s*(?:<!doctype\s+html|<html)/i.test(text);
}

function createNsrRecords(){
const DETAIL_FIELDS=['NSR#','NSR Title','Task Status','NSR Type','Impacted Sub System','Initiator','REF FCID','NSR Org','Primary Product','Customer','Sales Rep','Sales Ops','Technical Contact','Submit Date','NSR Request','Temp Id#'];
const ALIASES={
  'Temp Id#':['Temp Id#','Temp ID','TempId','Temporary ID'], 'NSR#':['NSR#','NSR','NSR No','NSR Number','NSR ID'], 'NSR Status':['NSR Status','Status'], 'REF FCID':['REF FCID','Reference Fcid','Reference FCID','Ref FCID','FCID'], 'NSR Request':['NSR Request','Request','Request Type'],
  'Task Name':['Task Name','Task','Workflow Task'], 'NSR Title':['NSR Title','Title'], 'NSR Type':['NSR Type','Type'], 'Impacted Sub System':['Impacted Sub System','Impacted Subsystem','Sub System','Subsystem'], 'Initiator':['Initiator','Initiated By','Requester'],
  'Sales Rep':['Sales Rep','Sales Representative'], 'Sales Ops':['Sales Ops','Sales Operation','Sales Operations'], 'Technical Contact':['Technical Contact','Tech Contact'], 'Customer':['Customer','Customer Name','Account'],
  'Submit Date':['Submit Date','Submitted Date','Submission Date'], 'NSR Org':['NSR Org','NSR Organization','Organization','Org'], 'Primary Product':['Primary Product','Product','PrimaryProduct']
};
const SOURCE_FIELDS=Object.keys(ALIASES);
const TASK_STATUS={
  'CREATE NSR':'1-Create NSR','REQUEST REVIEW':'2-Request Review','SAFETY REVIEW':'5-Safety Review','ESTIMATES':'6-Estimates','COST REVIEW':'7-Cost Review',
  'PRICE CONFIRMATION':'8-Price Confirmation','DEMAND CONFIRMATION':'9-Demand Confirmation','FINAL APPROVAL':'10-Final Approval','VC VALIDATION':'11-VC Validation','ENGINEERING COMMIT AND DESIGN':'12-Engineering Commit and Design'
};
const STATUS_TASK={'PENDING WITH INITIATOR':'1-Create NSR','PENDING WITH SAFETY':'5-Safety Review','PENDING WITH PGBD':'8-Price Confirmation','PENDING WITH POA PM':'11-VC Validation','PENDING WITH DESIGN ENGINEERING':'12-Engineering Commit and Design'};
const compareLabel=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'}).compare;
const clean=v=>{if(v===null||v===undefined)return ''; if(v instanceof Date&&!isNaN(v))return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`; const s=String(v).replace(/^\uFEFF/,'').trim();
  return /^(nan|null|undefined)$/i.test(s)?'':s;};
const upper=v=>clean(v).toUpperCase().replace(/\s+/g,' '); const normHeader=v=>clean(v).toLowerCase().replace(/&/g,'and').replace(/[^a-z0-9]+/g,'');
function deriveTaskStatus(row){
  const task=upper(row['Task Name']),status=upper(row['NSR Status']);
  if(status.includes('CANCEL')||task==='CANCELLED')return '0-Cancelled';
  if(status.includes('REJECT')||task==='REJECTED')return '0-Rejected';
  if(status.includes('COMPLET')||task==='COMPLETED')return '15-Completed';
  if(task==='INITIAL ASSESSMENT')return status==='PENDING WITH PRODUCT MANAGER'?'4-Initial Assessment':'3-Initial Assessment';
  if(task==='SOBI CREATE AND APPROVAL')return status==='PENDING WITH MFE'?'14-SOBI Create and Approval':'13-SOBI Create and Approval';
  if(Object.hasOwn(TASK_STATUS,task))return TASK_STATUS[task];
  if(STATUS_TASK[status])return STATUS_TASK[status];
  return `Unmapped · ${clean(row['Task Name'])||clean(row['NSR Status'])||'Blank task/status'}`;
}

function statusRank(v){const m=clean(v).match(/^(\d+)-/); if(!m)return [99,clean(v)]; return [Number(m[1]),clean(v)];}
function compareStatus(a,b){const aa=statusRank(a),bb=statusRank(b); return aa[0]-bb[0]||compareLabel(aa[1],bb[1]);}
let cachedHeaders=null,cachedHeaderMap;
function canonicalRow(raw){
  const headers=Object.keys(raw),signature=headers.join('\0');
  if(signature!==cachedHeaders){
    cachedHeaders=signature;const byNorm=new Map(headers.map(header=>[normHeader(header),header]));
    cachedHeaderMap=Object.fromEntries(SOURCE_FIELDS.map(field=>[field,ALIASES[field].map(alias=>byNorm.get(normHeader(alias))).find(Boolean)||'']));
  }
  return Object.fromEntries(SOURCE_FIELDS.map(field=>{const value=raw[cachedHeaderMap[field]];return [field,typeof value==='string'&&clean(value)&&!['NSR#','REF FCID','Temp Id#'].includes(field)?value:clean(value)]}));
}

function dateNumber(value){
  const s=clean(value);if(!s)return Number.NEGATIVE_INFINITY;
  if(/^\d{5}(?:\.\d+)?$/.test(s))return (Number(s)-25569)*86400000;
  let m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);if(m)return Date.UTC(Number(m[3]),Number(m[1])-1,Number(m[2]));
  m=s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);if(m)return Date.UTC(Number(m[1]),Number(m[2])-1,Number(m[3]));
  const n=Date.parse(s);return Number.isNaN(n)?Number.NEGATIVE_INFINITY:n;
}

function periodRank(value){let match=clean(value).match(/^(\d{4})\s+Q([1-4])$/);if(match)return Number(match[1])*10+Number(match[2]);
  match=clean(value).match(/^Q(Mar|Jun|Sep|Dec)'(\d{2}|\d{4})$/i);return match?Number(match[2].length===2?'20'+match[2]:match[2])*10+({mar:1,jun:2,sep:3,dec:4}[match[1].toLowerCase()]):0;}
function periodYear(value){const rank=periodRank(value);return rank?String(Math.floor(rank/10)):'';}
return {detailFields:DETAIL_FIELDS,clean,upper,canonicalRow,deriveTaskStatus,compareStatus,compareLabel,dateNumber,periodRank,periodYear};
}

function createNsrCycle(root=document,{getRows=null,onNotify=null,renderTable,renderFid=null,copyRows}={}){
const ownerDocument=root.ownerDocument||root,embedded=root!==ownerDocument,modalHost=embedded?root:ownerDocument.body;
const {detailFields:DETAIL_FIELDS,clean,upper,dateNumber,compareStatus,compareLabel}=createNsrRecords();
const LIVE_FIELDS=['Submit Date','Task Status','ECD','Approver Team','Assigned Approver','Pending Days','Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'];
const DETAIL_COLUMNS=DETAIL_FIELDS.filter(field=>field!=='NSR#'&&field!=='Initiator'&&!LIVE_FIELDS.includes(field)).flatMap(field=>field==='NSR Type'?[field,'NSR Category']:field==='Customer'?[field,'Initiator']:[field]);
const FID_FIELDS=['REF FCID','Additional FCID','Primary Product','NSR → CTL','CTL references','Customer','NSR Org','Sales Rep'],fidIndex=DETAIL_COLUMNS.indexOf('REF FCID');
const CYCLE_FIELDS=['NSR#',...LIVE_FIELDS,...DETAIL_COLUMNS.slice(0,fidIndex),...FID_FIELDS,...DETAIL_COLUMNS.slice(fidIndex).filter(field=>!FID_FIELDS.includes(field))];
const WF_FIELDS=['Task Name','Approver Status','Approver Team','Approver','Pending Days','Approved Date','Comments']; const WF_EXPORT=['NSR#','Workflow Step',...WF_FIELDS];
const WF_SUMMARY=['Task Status','Assigned Approver','Pending Days','Submit Date','REF FCID','NSR Type','NSR Category','NSR Org','Customer','Primary Product','Initiator','Approver Team','ECD','Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'];
const WIDTHS=CYCLE_FIELDS.map(field=>({
  'NSR#':78,'Submit Date':116,'Approver Team':125, 'Assigned Approver':130,'Pending Days':82,'Elapsed Days':82,'Cycle Time (ECD)':110,'Cycle Time (SOBi)':116,'NSR Title':190,'Task Status':144,'ECD':96, 'NSR Category':100,'Customer':150,'Primary Product':120,'Technical Contact':108,
    'Impacted Sub System':104,'Additional FCID':180,'NSR → CTL':180,'CTL references':180
}[field]||92));
const CONCURRENCY=8,WARN_AT=100,LIMIT=500; const WF_ODATA='/sap/opu/odata/sap/ZNSR_CD_SRV/WorkflowSet/'; const DETAIL_ODATA='/sap/opu/odata/sap/ZNSR_CD_SRV/GetNsrDetSet/'; const ADDITIONAL_ODATA='/sap/opu/odata/sap/ZNSR_CD_SRV/AddtionalFcidSet/'; const FIORI_DETAIL='https://fep.lamresearch.com/flp#ZNSR-display&/NsrDetail/';
const pageFetch=(url,options)=>unsafeWindow.fetch(url,options);
const state={rows:[],cycleRows:[],cycleAbort:null,cycleRun:0,toastTimer:null};
const $id=id=>root.querySelector('#'+id),esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function statusPill(value){const cls=value==='15-Completed'?'terminal':value.startsWith('0-')?'stop':value.startsWith('Unmapped')?'unknown':'';return `<span class="status-pill ${cls}">${esc(value)}</span>`;}
function toast(message,isError=false){
  if(onNotify){onNotify(message,isError);return}
  const el=$id('toast');if(!el)return;clearTimeout(state.toastTimer);el.textContent=message;el.className=`toast show${isError?' error':''}`;state.toastTimer=setTimeout(()=>el.className='toast',3000);
}
let wfTrigger=null,wfAbort=null,workflowTable=null,additionalTable=null,cycleTable=null;
function fidLinks(value){const fids=[...new Set(clean(value).split(',').map(clean).filter(Boolean))];return fids.length?fids.map(fid=>renderFid?renderFid(fid):esc(fid)).join(', '):'—'}
function parseNsrInput(text){
  const valid=[],invalid=[],seen=new Set(); String(text||'').split(/[\s,;|]+/).map(v=>v.trim()).filter(Boolean).forEach(token=>{const nsr=token.toUpperCase(); if(!/^[A-Z]\d{5,9}$/.test(nsr)){invalid.push(token);return} if(!seen.has(nsr)){seen.add(nsr);valid.push(nsr)}});
  return {valid,invalid};
}

function xmlValue(node,name){const item=node&&node.getElementsByTagNameNS('*',name)[0];return clean(item&&item.textContent);}
function xmlFirst(node,names){for(const name of names){const value=xmlValue(node,name);if(value)return value}return '';}
function workflowDate(value){const match=clean(value).match(/^(\d{4}-\d{2}-\d{2})/);return match?match[1]:'';}
function terminalStatus(value){const status=upper(value);return status.includes('COMPLET')?'COMPLETED':status.includes('CANCEL')?'CANCELLED':status.includes('REJECT')?'REJECTED':'';}

function workflowHistory(entries){
  return entries.map(entry=>{
    const task=xmlFirst(entry,['DescriptionTask','TaskName']),approved=workflowDate(xmlFirst(entry,['ApprovalDate','ApprovedDate'])); let status=xmlFirst(entry,['ApproverStatus','Decision','Status']);if(!status&&task)status=approved?'APPROVED':'PENDING';
    return {
      'Task Name':task,'Approver Status':status,'Approver Team':xmlFirst(entry,['ApproverTeam','ApprovalTeam']), 'Approver':xmlFirst(entry,['Approver','AssignedApprover','ApproverName','NameText']), 'Pending Days':xmlFirst(entry,['PendingDays','PendingDay']),'Approved Date':approved,
      'Comments':xmlFirst(entry,['Comments','Comment','ApprovalComments','ApprovalComment','ApproverComments','DecisionComments','DecisionComment','CommentText','Remarks','Remark','Notes']).replace(/%20/gi,' ')
    };
  }).filter(stage=>WF_FIELDS.some(field=>clean(stage[field])));
}

function dateStamp(value){const match=clean(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);return match?Date.UTC(Number(match[1]),Number(match[2])-1,Number(match[3])):NaN;}
function elapsedDays(startValue,endValue=''){const start=dateStamp(startValue),now=new Date(),end=endValue?dateStamp(endValue):Date.UTC(now.getFullYear(),now.getMonth(),now.getDate()); return Number.isFinite(start)&&Number.isFinite(end)?Math.max(0,Math.floor((end-start)/
  86400000)):'';}

function parseWorkflowXml(text,statusHint=''){
  const doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw new Error('The workflow service returned invalid XML.');
  const entries=[...doc.getElementsByTagNameNS('*','entry')],hintedTerminal=terminalStatus(statusHint);
  if(!entries.length)return {'Submit Date':'','Approver Team':'','Assigned Approver':'','Pending Days':'','Elapsed Days':'','Cycle Time (ECD)':'','Cycle Time (SOBi)':'',__completed:hintedTerminal==='COMPLETED',__completedAt:'',__cycleEndAt:'',__terminal:hintedTerminal,__workflow:[],
    __error:'No workflow data'};
  const decisions=entries.map(entry=>upper(xmlValue(entry,'Decision'))),submittedIndex=decisions.indexOf('SUBMITTED'); const submittedEntry=submittedIndex>=0?entries[submittedIndex]:null; const submitted=workflowDate(xmlValue(submittedEntry,'ApprovalDate')); let lastAction=-1;
    const approvedDates=[],terminalDates={};
  entries.forEach((entry,index)=>{
    const decision=decisions[index],date=workflowDate(xmlValue(entry,'ApprovalDate'));
    if(decision&&decision!=='PENDING'&&decision!=='N/A')lastAction=index;
    if(['APPROVED','COMPLETED'].includes(decision)&&date)approvedDates.push(date);
    const terminalDecision=terminalStatus(decision);
    if(terminalDecision&&date&&(!terminalDates[terminalDecision]||dateStamp(date)>dateStamp(terminalDates[terminalDecision])))terminalDates[terminalDecision]=date;
  });
  const after=entries.slice(lastAction+1);
  const pending=after.find(entry=>{const decision=upper(xmlValue(entry,'Decision')),approved=workflowDate(xmlValue(entry,'ApprovalDate')); return (!decision||decision==='PENDING')&&!approved&&['DescriptionTask','ApproverTeam','Approver','PendingDays'].some(name=>xmlValue(entry,name));
    })||null;
  const lastDecision=lastAction>=0?decisions[lastAction]:'',xmlTerminal=terminalStatus(lastDecision),terminal=hintedTerminal||xmlTerminal; const completedAt=approvedDates.reduce((latest,date)=>!latest||dateStamp(date)>dateStamp(latest)?date:latest,'');
  const completed=terminal==='COMPLETED'||(!terminal&&lastDecision==='APPROVED'&&!pending); const cycleEndAt=completed?completedAt:(terminalDates[terminal]||'');
  return {
    'Submit Date':submitted, 'Approver Team':terminal?'':xmlValue(pending,'ApproverTeam'), 'Assigned Approver':terminal?'':xmlValue(pending,'Approver'), 'Pending Days':terminal?'':xmlValue(pending,'PendingDays'), 'Elapsed Days':'', 'Cycle Time (ECD)':'', 'Cycle Time (SOBi)':'',
    __completed:completed, __completedAt:completedAt, __cycleEndAt:cycleEndAt, __terminal:terminal, __workflow:workflowHistory(entries)
  };
}

function odataLookupUrl(endpoint,field,value){const literal=clean(value).replaceAll("'","''"),filter=`${field}  eq '${literal}'`;return `${endpoint}?$filter=${encodeURIComponent(filter)}`;}

async function fetchOdataXml(url,signal,label){
  let response;
  try{response=await pageFetch(url,{credentials:'same-origin',cache:'no-store',signal,headers:{Accept:'application/atom+xml,application/xml,text/xml'}})} catch(err){if(signal.aborted)throw err;const e=new Error('Fiori session unavailable or request blocked.');e.auth=true;throw e}
  const text=await response.text();
  if(nsrSignInResponse(response,text)){const err=new Error('Session expired. Return to NSR page and sign in again.');err.auth=true;throw err;}
  if(!response.ok)throw new Error(`${label} returned HTTP ${response.status}.`);
  if(!/<(?:[A-Za-z0-9_-]+:)?feed\b/i.test(text)&&!/<(?:[A-Za-z0-9_-]+:)?entry\b/i.test(text))throw new Error(`${label} returned an unexpected response.`);
  return text;
}


async function fetchNsrEnrichment(tempId,signal){
  if(!clean(tempId))return {category:'',ecd:''};
  const text=await fetchOdataXml(odataLookupUrl(DETAIL_ODATA,'TempId',tempId),signal,'NSR detail service'); const doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw new Error('NSR detail service returned invalid XML.');
  const entry=doc.getElementsByTagNameNS('*','entry')[0]; return {category:xmlValue(entry,'NsrCategory'),ecd:workflowDate(xmlValue(entry,'HwEngCommit'))};
}

async function fetchAdditionalFcids(tempId,signal){
  const text=await fetchOdataXml(odataLookupUrl(ADDITIONAL_ODATA,'TempId',tempId),signal,'Additional FCID service'),doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw new Error('Additional FCID service returned invalid XML.');
  return [...new Set([...doc.getElementsByTagNameNS('*','entry')].map(entry=>xmlValue(entry,'Fcid')).filter(Boolean))].join(', ');
}

function mergeCycle(nsr,live,details){
  const detail=details.get(upper(nsr))||{},row={'NSR#':nsr,...live};
  if(!row['Submit Date'])row['Submit Date']=detail['Submit Date']||'';
  CYCLE_FIELDS.forEach(field=>{if(field!=='NSR#'&&field!=='NSR Category'&&field!=='Additional FCID'&&!LIVE_FIELDS.includes(field))row[field]=detail[field]||''}); row['Task Status']=detail['Task Status']||({COMPLETED:'15-Completed',CANCELLED:'0-Cancelled',REJECTED:'0-Rejected'}[row.__terminal]||''); row['ECD']=row.__ecd||'';
  row['Elapsed Days']=row.__terminal||row.__completed?'':elapsedDays(row['Submit Date']); row['Cycle Time (ECD)']=row.__ecd?elapsedDays(row['Submit Date'],row.__ecd):''; row['Cycle Time (SOBi)']=row.__completed&&row.__completedAt?elapsedDays(row['Submit Date'],row.__completedAt):'';
  return row;
}

function setCycleBusy(busy){$id('runCycleBtn').disabled=busy;$id('cancelCycleBtn').disabled=!busy;$id('cycleInput').disabled=busy; $id('cycleBadge').textContent=busy?'Querying':'Ready';}
function cycleProgress(done,total,message){const progress=$id('cycleProgress');progress.max=Math.max(1,total);progress.value=done; $id('cycleStatus').textContent=message;$id('cycleBadge').textContent=total?`${done}/${total}`:'Ready';}

function cycleError(err,statusHint=''){
  const message=clean(err&&err.message)||'Lookup failed',terminal=terminalStatus(statusHint);
  return {'Submit Date':'','Approver Team':'','Assigned Approver':'','Pending Days':'','Elapsed Days':'','Cycle Time (ECD)':'','Cycle Time (SOBi)':'',__completed:terminal==='COMPLETED',__completedAt:'',__cycleEndAt:'',__terminal:terminal,__error:message};
}

async function runCycleLookup(nsrs){
  const parsed=parseNsrInput(Array.isArray(nsrs)?nsrs.join('\n'):$id('cycleInput').value);
  if(!parsed.valid.length){toast('Enter at least one valid NSR#.',true);return}
  if(parsed.invalid.length)toast(`${parsed.invalid.length} invalid value${parsed.invalid.length===1?' was':'s were'} ignored.`,true);
  const total=parsed.valid.length;
  if(total>LIMIT){const message=`A maximum of ${LIMIT} NSRs can be queried at once. Reduce the list and try again.`;unsafeWindow.alert(message);cycleProgress(0,0,message);return}
  if(total>WARN_AT&&!unsafeWindow.confirm(`You are about to query ${total} NSRs from SAP. This may take some time. Continue?`)){cycleProgress(0,0,`Query cancelled. ${total} NSRs were not sent to SAP.`);return;}
  if(state.cycleAbort)state.cycleAbort.abort();
  const run=++state.cycleRun,controller=new unsafeWindow.AbortController();state.cycleAbort=controller;state.cycleRows=[]; const details=new Map((getRows?getRows(parsed.valid):state.rows).map(row=>[upper(row['NSR#']),row])),results=new Array(total),workers=Math.min(CONCURRENCY,total);
  let next=0,done=0,authFailed=false,lastPaint=0; setCycleBusy(true);cycleProgress(0,total,`Starting ${total} live lookup${total===1?'':'s'} · up to ${workers} requests at once…`);await renderCycle();
  async function worker(){
    while(!controller.signal.aborted&&run===state.cycleRun&&!authFailed){
      const index=next++;if(index>=total)return;
      const nsr=parsed.valid[index],detail=details.get(upper(nsr)),statusHint=terminalStatus(detail&&detail['Task Status']); let live;
      try{live=parseWorkflowXml(await fetchOdataXml(odataLookupUrl(WF_ODATA,'Nsrnum',nsr),controller.signal,'Workflow service'),statusHint)} catch(err){if(controller.signal.aborted)return; live=cycleError(err,statusHint);if(err.auth){authFailed=true;controller.abort()}}
      let categoryError='';live['NSR Category']='';live.__ecd='';live['Additional FCID']='';
      if(!authFailed&&clean(detail&&detail['Temp Id#'])){
        try{const enrichment=await fetchNsrEnrichment(detail['Temp Id#'],controller.signal);live['NSR Category']=enrichment.category;live.__ecd=enrichment.ecd}
        catch(err){if(controller.signal.aborted)return; categoryError=clean(err&&err.message)||'NSR Category unavailable';if(err.auth){authFailed=true;controller.abort()}}
        if(!authFailed){
          try{live['Additional FCID']=await fetchAdditionalFcids(detail['Temp Id#'],controller.signal)}
          catch(err){if(controller.signal.aborted)return;live.__additionalError=clean(err&&err.message)||'Additional FCID unavailable';live['Additional FCID']='Unavailable';if(err.auth){authFailed=true;controller.abort()}}
        }
      }
      if(categoryError)live.__categoryError=categoryError;
      if(run!==state.cycleRun)return;
      results[index]=mergeCycle(nsr,live,details);done++;state.cycleRows=results.filter(Boolean); const now=Date.now();if(now-lastPaint>200||done===total||authFailed){lastPaint=now;await renderCycle()}
      if(run!==state.cycleRun)return;
      if(!authFailed)cycleProgress(done,total,`Queried ${done} of ${total} · up to ${workers} requests at once`);
    }
  }
  await Promise.all(Array.from({length:workers},()=>worker()));
  if(run!==state.cycleRun)return;
  state.cycleRows=results.filter(Boolean);await renderCycle();if(run!==state.cycleRun)return;
  const cancelled=controller.signal.aborted&&!authFailed,failed=state.cycleRows.filter(row=>row.__error||row.__categoryError||row.__additionalError).length;state.cycleAbort=null;setCycleBusy(false);
  if(authFailed){cycleProgress(done,total,'Stopped because the Fiori session is unavailable. Return to Fiori, sign in, and run the query again.');toast('Fiori session unavailable.',true)}
  else if(cancelled)cycleProgress(done,total,`Cancelled after ${done} of ${total} lookups.`);
  else{cycleProgress(done,total,`${done} live lookup${done===1?'':'s'} completed${failed?` · ${failed} unavailable`:''}.`);
    toast(`${done} cycle time row${done===1?'':'s'} loaded${failed?` · ${failed} unavailable`:''}.`,failed>0)}
}

function compareCycle(a,b,field){
  if(['Pending Days','Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'].includes(field))return (Number(a)||0)-(Number(b)||0);
  if(field==='ECD'||field.includes('Date'))return dateNumber(a)-dateNumber(b);
  if(field==='Task Status')return compareStatus(a,b);
  return compareLabel(clean(a),clean(b));
}

function cycleRows(){
  return cycleTable?.ready?cycleTable.grid.getRows('active').map(row=>cycleTable.rows[row.getIndex()]):[];
}
function cycleFields(){return cycleTable.grid.getColumns().filter(column=>column.isVisible()).map(column=>[CYCLE_FIELDS[Number(column.getField().slice(1))],column.getDefinition().title])}

function nsrDetailUrl(row){return clean(row['Temp Id#'])&&clean(row['NSR#'])?`${FIORI_DETAIL}${encodeURIComponent(clean(row['Temp Id#']))}/${encodeURIComponent(clean(row['NSR#']))}/Search`:''}

function cycleCell(row,field){
  if(field==='Task Status')return statusPill(row[field]);
  if(field==='REF FCID')return fidLinks(row[field]);
  if(field==='Additional FCID')return `<span title="${esc(row.__additionalError||row[field]||'')}">${row.__additionalError?esc(row[field]):fidLinks(row[field])}</span>`;
  if(field==='NSR Title'&&clean(row[field])&&clean(row['Temp Id#'])&&clean(row['NSR#'])){
    const href=nsrDetailUrl(row);
    return `<a class="ct-link" data-nsr-detail="${esc(row['NSR#'])}" href="${esc(href)}" target="_blank" rel="noopener noreferrer" title="Open ${esc(row['NSR#'])} in Fiori">${esc(row[field])}</a>`;
  }
  if(clean(row[field])&&(field==='NSR#'||['Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'].includes(field)&&row.__workflow?.length))return `<button class="ct-link" type="button" data-wf-nsr="${esc(row['NSR#'])}" title="View full workflow for ${esc(row['NSR#'])}">${esc(row[field])}</button>`;
  return esc(row[field]);
}

function cycleClass(row,field){
  const elapsed=clean(row['Elapsed Days']),ecdCycle=clean(row['Cycle Time (ECD)']),sobiCycle=clean(row['Cycle Time (SOBi)']); const elapsedOverEcd=elapsed&&ecdCycle&&Number(elapsed)>Number(ecdCycle); const ecdOverSobi=ecdCycle&&sobiCycle&&Number(ecdCycle)>Number(sobiCycle);
  if(field==='Elapsed Days'&&elapsedOverEcd)return 'duration-over';
  if(field==='Cycle Time (ECD)'&&(elapsedOverEcd||ecdOverSobi))return 'duration-over';
  return '';
}

async function openWorkflow(nsr,trigger){
  const modal=$id('wfModal'),title=`${clean(nsr)} Workflow Detail`;
  if(modal.open&&trigger?.closest('#fidDetail')){modal.close();modal.showModal();wfTrigger=trigger;$id('wfClose').focus({preventScroll:true})}
  if(modal.open&&$id('wfTitle').textContent===title)return;
  if(wfAbort)wfAbort.abort();const controller=new unsafeWindow.AbortController();wfAbort=controller;
  let row=state.cycleRows.find(item=>upper(item['NSR#'])===upper(nsr)&&item.__workflow?.length);
  const detail=row||(getRows?getRows([nsr]):state.rows).find(item=>upper(item['NSR#'])===upper(nsr))||{'NSR#':nsr};
  const link=$id('wfNsrLink'),href=nsrDetailUrl(detail),status=$id('wfStatus'),summary=$id('wfSummary'),additional=$id('wfAdditional'),table=$id('wfTable');
  $id('wfTitle').textContent=title;$id('wfMeta').textContent=clean(detail['NSR Title']);
  link.hidden=!href;if(href)link.href=href;else link.removeAttribute('href');
  summary.hidden=true;additional.hidden=true;table.hidden=true;status.hidden=false;status.textContent='Loading workflow…';modal.setAttribute('aria-busy','true');
  wfTrigger=trigger||ownerDocument.activeElement;if(!modal.open)modal.showModal();modalHost.classList.add('wf-open');
  $id('wfClose').focus({preventScroll:true});
  try{
    if(!row){
      const xml=await fetchOdataXml(odataLookupUrl(WF_ODATA,'Nsrnum',nsr),controller.signal,'Workflow service');
      const live=parseWorkflowXml(xml,terminalStatus(detail['Task Status']));if(!live.__workflow?.length)throw new Error(live.__error||'Workflow detail is unavailable for this NSR.');
      live['Additional FCID']='';
      if(clean(detail['Temp Id#'])){
        const [enrichment,fcids]=await Promise.allSettled([fetchNsrEnrichment(detail['Temp Id#'],controller.signal),fetchAdditionalFcids(detail['Temp Id#'],controller.signal)]);
        if(enrichment.status==='fulfilled'){live['NSR Category']=enrichment.value.category;live.__ecd=enrichment.value.ecd}else live.__categoryError=clean(enrichment.reason.message)||'NSR Category unavailable';
        if(fcids.status==='fulfilled')live['Additional FCID']=fcids.value;else live.__additionalError=clean(fcids.reason.message)||'Additional FCID unavailable';
      }
      row=mergeCycle(nsr,live,new Map([[upper(nsr),detail]]));
    }
    if(controller.signal.aborted||wfAbort!==controller||!modal.open)return;
    if(!row.__workflow?.length)throw new Error(row.__error||'Workflow detail is unavailable for this NSR.');
    summary.innerHTML=WF_SUMMARY.map(field=>`<div class="fid-detail-value"${field==='NSR Category'&&row.__categoryError?` title="${esc(row.__categoryError)}"`:''}><small>${field==='REF FCID'?'FID':field}</small>${field==='REF FCID'?fidLinks(row[field]):field==='Task Status'&&clean(row[field])?statusPill(clean(row[field])):field==='NSR Category'&&row.__categoryError?'Unavailable':esc(row[field])||'—'}</div>`).join('');
    summary.hidden=false;additional.hidden=false;table.hidden=false;status.textContent=[row.__categoryError,row.__additionalError].filter(Boolean).join(' · ');status.hidden=!status.textContent;
    const fids=row.__additionalError?[]:[...new Set(clean(row['Additional FCID']).split(',').map(clean).filter(Boolean))].map(fid=>({FID:fid}));
    additionalTable=await renderTable('wfAdditional',[['FID',row=>row.FID,fidLinks]],fids,{title:'Additional FCID',copyId:'wfAdditionalCopy',widths:{FID:120}});
    if(controller.signal.aborted||wfAbort!==controller||!modal.open)return;
    if(row.__additionalError)additionalTable.toolbar.querySelector('[data-table-count]').textContent='Unavailable';
    workflowTable=await renderTable('wfTable',WF_FIELDS.map(field=>[field,row=>row[field],value=>field==='Approver Status'?`<span class="pill ${/APPROVED|COMPLETED/.test(upper(value))?'clear':/REJECT|CANCEL/.test(upper(value))?'risk':'info'}">${esc(value)}</span>`:esc(value),{variableHeight:true,...(field==='Pending Days'?{sorter:(a,b)=>(Number(a)||0)-(Number(b)||0)}:{})}]),row.__workflow,{title:'Workflow History',rowHeight:false,copyId:'wfCopy',onCopy:copyWorkflow,widths:{'Task Name':160,'Approver Status':130,'Approver Team':140,Approver:120,'Pending Days':112,'Approved Date':125,Comments:360}});
  }catch(error){if(controller.signal.aborted||wfAbort!==controller||!modal.open)return;summary.hidden=true;additional.hidden=true;table.hidden=true;status.hidden=false;status.textContent=clean(error.message)||'Workflow detail is unavailable for this NSR.'}
  finally{if(wfAbort===controller)modal.removeAttribute('aria-busy')}
  if(modal.open&&wfAbort===controller)$id('wfClose').focus({preventScroll:true});
}

function closeWorkflow(){if($id('wfModal').open)$id('wfModal').close()}
function restoreWfFocus(){
  if(wfAbort)wfAbort.abort();wfAbort=null;$id('wfModal').removeAttribute('aria-busy');modalHost.classList.remove('wf-open');
  const nsr=wfTrigger?.dataset.wfNsr;
  const trigger=wfTrigger?.isConnected?wfTrigger:nsr?ownerDocument.querySelector('#fidDetail[open] [data-wf-nsr="'+CSS.escape(nsr)+'"]'):null;
  trigger?.focus({preventScroll:true});wfTrigger=null;
}

async function renderCycle(){
  cycleTable=await renderTable('cycleTable',CYCLE_FIELDS.map(field=>[field,row=>row[field],(_,row)=>`<span class="${cycleClass(row,field)}">${cycleCell(row,field)}</span>`,{title:embedded&&field==='REF FCID'?'FID':field,sorter:(a,b)=>compareCycle(a,b,field)}]),state.cycleRows,{file:'NSR_Cycle_Time',height:450,columnFilters:false,defaultSort:'NSR#',copyId:'copyCycleBtn',exportId:'exportCycleBtn',onCopy:copyCycle,onExport:exportCycle,widths:Object.fromEntries(CYCLE_FIELDS.map((field,index)=>[field,WIDTHS[index]]))});
}

function copyWorkflow(){if(!workflowTable?.ready)return;const fields=workflowTable.grid.getColumns().filter(column=>column.isVisible()).map(column=>WF_FIELDS[Number(column.getField().slice(1))]),rows=workflowTable.grid.getRows('active').map(row=>workflowTable.rows[row.getIndex()]);if(rows.length)return copyRows([fields,...rows.map(row=>fields.map(field=>row[field]??''))])}
function copyCycle(){const rows=cycleRows();if(!rows.length)return;const fields=cycleFields();return copyRows([fields.map(([,label])=>label),...rows.map(row=>fields.map(([field])=>row[field]??''))])}

function exportCycle(){
  const rows=cycleRows();if(!rows.length)return;
  if(!window.XLSX){toast('The XLSX library did not load.',true);return}
  const fields=cycleFields(),data=rows.map(row=>Object.fromEntries(fields.map(([field,label])=>[label,row[field]??'']))),sheet=XLSX.utils.json_to_sheet(data,{header:fields.map(([,label])=>label)});
  const workflowData=rows.flatMap(row=>(Array.isArray(row.__workflow)?row.__workflow:[]).map((stage,index)=>Object.fromEntries(WF_EXPORT.map(field=>[field,field==='NSR#'?(row['NSR#']??''):field==='Workflow Step'?index+1:(stage[field]??'')]))));
  const workflowSheet=XLSX.utils.json_to_sheet(workflowData,{header:WF_EXPORT}); sheet['!cols']=fields.map(([field])=>({wch:Math.max(10,Math.round(WIDTHS[CYCLE_FIELDS.indexOf(field)]/7))}));sheet['!autofilter']={ref:sheet['!ref']};
  workflowSheet['!cols']=[12,14,32,18,22,24,14,14,50].map(wch=>({wch}));workflowSheet['!autofilter']={ref:workflowSheet['!ref']}; const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'NSR Cycle Time');XLSX.utils.book_append_sheet(book,workflowSheet,
    'Workflow Detail'); const now=new Date(),stamp=`${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`; XLSX.writeFile(book,`NSR_Cycle_Time_${stamp}.xlsx`,{compression:true});
    toast(`${rows.length} summary rows and ${workflowData.length} workflow records exported.`);
}

function cancel(){if(!state.cycleAbort)return;state.cycleAbort.abort();$id('cycleStatus').textContent='Cancelling active requests…';}
const listeners=[];
function bind(element,event,handler){element.addEventListener(event,handler);listeners.push([element,event,handler]);}
bind($id('runCycleBtn'),'click',()=>runCycleLookup());bind($id('cancelCycleBtn'),'click',cancel);
bind(ownerDocument,'click',event=>{const trigger=event.target.closest('[data-wf-nsr]');if(trigger)openWorkflow(trigger.dataset.wfNsr,trigger)});
bind($id('wfModal'),'click',event=>{if(event.target.closest('[data-wf-close]'))closeWorkflow()});
bind($id('wfModal'),'close',()=>{if(!$id('wfModal').open)restoreWfFocus()});
setCycleBusy(false);
return {
  setRows(rows){state.rows=Array.isArray(rows)?rows:[]},
  show:()=>renderCycle().then(()=>cycleTable.grid.redraw(true)),
  query(nsrs,{run=true}={}){if(nsrs!==undefined)$id('cycleInput').value=Array.isArray(nsrs)?nsrs.join('\n'):String(nsrs);return run?runCycleLookup():parseNsrInput($id('cycleInput').value)},
  snapshot:()=>({rows:state.cycleRows,busy:Boolean(state.cycleAbort)}),cancel,
  destroy(){state.cycleRun++;if(state.cycleAbort)state.cycleAbort.abort();state.cycleAbort=null;setCycleBusy(false);clearTimeout(state.toastTimer);closeWorkflow();restoreWfFocus();cycleTable?.grid.destroy();workflowTable?.grid.destroy();additionalTable?.grid.destroy();listeners.forEach(([element,event,handler])=>element.removeEventListener(event,handler));}
};
}



// ==================== MyLam · Filtered Document Summary ====================
(function(){
'use strict';
const KMMATRIX_ORIGIN='https://kmmatrix.fremont.lamrc.net';
const KMMATRIX_API_ORIGIN='https://kmapi.fremont.lamrc.net';
const KMMATRIX_JOB_PREFIX='toolkit.mylam.partDescriptions.',KMMATRIX_TAB_KEY='toolkit.mylam.partDescriptionJob';
const KMMATRIX_SIGN_IN_TIMEOUT=120000,KMMATRIX_API_TIMEOUT=8000,KMMATRIX_RESULT_TIMEOUT=12000,KMMATRIX_REPLY_TIMEOUT=25000;
const KMMATRIX_CLOSE_TIMEOUT=1000;
const KMMATRIX_PART_NUMBER=/^[A-Z0-9][A-Z0-9._-]{4,63}$/;
if(location.origin===KMMATRIX_ORIGIN){runKmmatrixDescriptionTab();return}
if(location.hostname!=='techinfo.mylam.com'||!location.pathname.startsWith('/portal/group/mylam/contenttyperesults'))return;
const BUTTON_ID='kyra-mylam-summary',BUTTON_LABEL=PART_CHANGES_ACTION_LABEL,SEARCH_LABEL='Find Tech Articles for BOM',DIALOG_ID='kyra-mylam-summary-dialog',SUMMARY_CONCURRENCY=10,SUMMARY_REQUEST_TIMEOUT=30000;
const PREVIEW_URL=location.origin+'/portal/group/mylam/tech-article-preview?source=PE&articleId='; let cancelSummary=()=>{};

function parseDocument(row){
  const cells=row.cells,action=row.querySelector('img.hourIcon[onclick]')?.getAttribute('onclick')||''; const match=action.match(/emailURL\(\s*'((?:\\.|[^'\\])*)'\s*,\s*'((?:\\.|[^'\\])*)'\s*,\s*'((?:\\.|[^'\\])*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'(?:\s*,\s*'([^']*)')?\s*\)/);
  if(!match)return null;
  const [originalUrl,type,documentNo,phase2,articleId,spinCleanFlag]=match.slice(1).map(value=>(value||'').replace(/\\(['"\\])/g,'$1')); const hasId=articleId&&articleId!=='null',preview=[phase2,spinCleanFlag].some(value=>value.toLowerCase()==='true');
  const target=preview?(hasId?PREVIEW_URL+encodeURIComponent(articleId):''):originalUrl&&originalUrl!=='null'?originalUrl:hasId?PREVIEW_URL+encodeURIComponent(articleId):'';
  if(!target)return null;
  const url=new URL(target,location.origin+'/');if(!/^https?:$/.test(url.protocol))return null;
  return {type:type||cells[0]?.textContent.trim()||'',documentNo:documentNo||cells[1]?.textContent.trim()||'', title:cells[3]?.textContent.trim()||'',releaseDate:cells[4]?.textContent.trim()||'',articleId:hasId?articleId:'',url:url.href};
}

function waitForResultsDraw(node,trigger){
  return new Promise((resolve,reject)=>{
    const events='draw.dt.kyraSummary error.dt.kyraSummary dt-error.dt.kyraSummary';
    const finish=error=>{clearTimeout(timer);node.off(events,onEvent);error?reject(error):resolve()};
    const onEvent=(event,settings,code,message)=>finish(event.type==='draw'?null:new Error(message||'Could not load MyLam results.'));
    const timer=setTimeout(()=>finish(new Error('Loading MyLam results timed out. Please retry.')),SUMMARY_REQUEST_TIMEOUT);
    node.on(events,onEvent);
    try{trigger()}catch(error){finish(error)}
  });
}

function drawPage(jq,table,length,page){
  return waitForResultsDraw(jq(table.table().node()),()=>table.page.len(length).page(page).draw('page'));
}

async function collectDocuments(){
  // DataTables belongs to the page's jQuery, not the userscript's bundled jQuery.
  const jq=[unsafeWindow.jQuery,unsafeWindow.$].find(value=>value?.fn?.dataTable?.isDataTable('#searchResults'));
  if(!jq)throw new Error('MyLam search results are not ready. Please wait and retry.');
  const apply=document.querySelector('button#apply.applyFilter');
  if(!apply||apply.disabled)throw new Error('MyLam filters are not ready. Please wait and retry.');
  const results=jq('#searchResults');
  await waitForResultsDraw(results,()=>apply.click());
  const table=results.DataTable(),original=table.page.info(),items=new Map();let read=0,missing=0;
  try{
    if(original.length!==500||original.page!==0)await drawPage(jq,table,500,0);
    const first=table.page.info();
    for(let page=0;page<first.pages;page++){
      if(page)await drawPage(jq,table,500,page);
      const current=table.page.info();
      if(current.page!==page||current.recordsDisplay!==first.recordsDisplay)throw new Error('Search results changed. Please retry Summary.');
      for(const row of table.rows({page:'current',search:'applied'}).nodes().toArray()){read++;const item=parseDocument(row);if(item)items.set(item.articleId||item.url,item);else missing++;}
    }
    if(read!==first.recordsDisplay||missing)throw new Error(`Could not read all filtered documents (${read-missing}/${first.recordsDisplay}). Please retry.`);
    return Array.from(items.values(),(item,documentIndex)=>({...item,documentIndex}));
  }finally{const current=table.page.info(); if(current.length!==original.length||current.page!==original.page)await drawPage(jq,table,original.length,original.page);}
}

function extractPartNumbers(html){
  const parse=text=>{const template=document.createElement('template');template.innerHTML=text;return template.content};
  const page=parse(html),root=page.querySelector('.tech-document-article__container');
  const partHeading=/\b(?:parts?\s*(?:numbers?\b|nos?\b\.?|#)|p\s*\/\s*n\b|p\.?n\.?s?\b)/i;
  const header=text=>text.length<100&&!/\d/.test(text)&&partHeading.test(text);
  const descriptionHeading=/descriptions?|\bdesc\b/i;
  const legacyTables=new Set();
  if(!root){
    const tables=[...page.querySelectorAll('table')].map(table=>({table,labels:[...table.rows[0]?.cells||[]].map(cell=>
      cell.textContent.toLowerCase().replace(/\s+/g,' ').trim())}));
    // Older MyLam HTML articles have paired part and document tables instead of the portal container.
    const hasPartTable=tables.some(({labels})=>labels.includes('part number')&&labels.includes('description')&&(labels.includes('rev')||labels.includes('qty')));
    const hasDocumentTable=tables.some(({labels})=>labels.includes('document number')&&labels.includes('description'));
    if(hasPartTable&&hasDocumentTable){
      for(const {table,labels} of tables){
        if(labels.some(label=>!descriptionHeading.test(label)&&header(label)))legacyTables.add(table);
      }
    }
  }
  if(root?!root.querySelector('.tech-document-article__fields,.tech-document-article__main-info'):!legacyTables.size){
    throw new Error('Article content not found (login, access restriction or unsupported page).');
  }
  const fragments=[root||page],occurrences=[],numbers=new Set(),partDescriptions=new Map();
  const normalize=text=>text.replace(/[\u2010-\u2015\u2212]/g,'-').replace(/\u00ad/g,'').replace(/\u00a0/g,' ');
  // Rich-text tables are serialized as raw*Data strings; decode literals without executing scripts.
  for(const script of root?.querySelectorAll('script')||[]){
    const declarations=[...script.textContent.matchAll(/\b(?:const|let|var)\s+(raw\w*Data)\s*=/g)];
    for(const declaration of declarations){
      const rest=script.textContent.slice(declaration.index+declaration[0].length),literal=rest.match(/^\s*("(?:\\[\s\S]|[^"\\])*"|'(?:\\[\s\S]|[^'\\])*')\s*;/);
      if(!literal)throw new Error(`Could not read embedded article content: ${declaration[1]}.`);
      const decoded=literal[1].slice(1,-1).replace(/\\(u[\da-f]{4}|x[\da-f]{2}|\r?\n|[\s\S])/gi,(_,escape)=>{if(/^[ux][\da-f]+$/i.test(escape))return String.fromCharCode(parseInt(escape.slice(1),16)); return ({n:'\n',r:'\r',t:'\t',b:'\b',f:'\f',v:'\v','0':'\0','\n':'',
        '\r\n':''})[escape]??escape;});
      const fragment=parse(decoded);fragments.push(!fragment.querySelector('table')&&/&lt;table\b/i.test(decoded)?parse(fragment.textContent):fragment);
    }
  }
  const metadataDescription=/^(?:(?:document|procedure|rev(?:ision)?|change)\s+(?:descriptions?|desc)|(?:descriptions?|desc)\s+of\s+(?:the\s+)?(?:document|procedure|rev(?:ision)?|change))\.?$/i;
  const partRoles=['old','new','current','existing','original','previous','obsolete','affected','alternate','replacement','revised','before','after','pre','post'];
  const headingWords=text=>(text.toLowerCase().replace(partHeading,' ').replace(descriptionHeading,' ').match(/[a-z]+/g)||[])
    .filter(word=>!/^(?:parts?|of|for|the)$/.test(word));
  const referenceHeader=/^(?:doc(?:ument)?|procedure)\s+(?:numbers?|nos?\.?|#|id)$/i;
  let tableIndex=0;
  for(const fragment of fragments)for(const table of fragment.querySelectorAll('table')){
    tableIndex++;
    if(!root&&!legacyTables.has(table))continue;
    const grid=[],columns=new Set(),cellText=new Map(),descriptionColumns=new Map(),referenceColumns=new Set();
    for(const [r,row] of [...table.rows].entries()){
      grid[r]??=[];let column=0;
      for(const cell of row.cells){while(grid[r][column])column++; const rowSpan=cell.rowSpan||table.rows.length-r; for(let y=r;y<Math.min(table.rows.length,r+rowSpan);y++)for(let x=column;x<column+cell.colSpan;x++){grid[y]??=[];grid[y][x]=cell} column+=cell.colSpan;}
    }
    for(const [r,row] of grid.entries()){
      const seen=new Set(),rowCells=[],partHeaders=[];
      for(const [c,cell] of row.entries()){
        if(!cell||seen.has(cell))continue;
        seen.add(cell);
        if(!cellText.has(cell)){
          const clone=cell.cloneNode(true);
          clone.querySelectorAll('table,script,style').forEach(node=>node.remove());
          clone.querySelectorAll('br').forEach(node=>node.replaceWith('\n'));
          clone.querySelectorAll('p,div,li').forEach(node=>node.append('\n'));
          cellText.set(cell,clone.textContent.trim());
        }
        const text=normalize(cellText.get(cell)).trim(),label=text.replace(/_/g,' ').replace(/\s+/g,' ');
        const isDescriptionHeader=label.length<100&&descriptionHeading.test(label);
        const isPartHeader=!isDescriptionHeader&&header(label);
        const entry={c,cell,text,label,isPartHeader,isDescriptionHeader};
        rowCells.push(entry);
        if(isPartHeader)partHeaders.push(entry);
      }
      // A data cell can mention P/N without starting a new set of table headings.
      const hasPartValue=rowCells.some(({c,text})=>columns.has(c)&&/\d/.test(text));
      const isHeadingRow=partHeaders.length>0&&!hasPartValue;
      const references=hasPartValue?[]:rowCells.filter(({label})=>referenceHeader.test(label)||/\bdocument$/i.test(label));
      // Document and procedure references can have the same format as Lam part numbers.
      if(r===0&&!partHeaders.length&&references.some(({label})=>referenceHeader.test(label)))break;
      for(const {c,cell} of references)for(let x=c;x<c+cell.colSpan;x++)referenceColumns.add(x);
      if(isHeadingRow){
        descriptionColumns.clear();
        const parts=partHeaders.map(({c,label})=>({c,words:headingWords(label)}));
        const qualifiers=new Set([...partRoles,...parts.flatMap(part=>part.words)]);
        const descriptions=rowCells.filter(entry=>entry.isDescriptionHeader&&!metadataDescription.test(entry.label)).map(({c,label})=>{
          const words=headingWords(label).filter(word=>qualifiers.has(word));
          const owners=words.length?parts.filter(part=>words.every(word=>part.words.includes(word))):[];
          return {c,owner:owners.length===1?owners[0].c:null,qualified:words.length>0};
        });
        const separateColumns=descriptions.some(description=>description.c>parts[0].c&&description.c<parts.at(-1).c);
        parts.forEach(({c},index)=>{
          const next=parts[index+1]?.c??Infinity;
          const qualified=descriptions.filter(description=>description.owner===c);
          // A single generic description after several part columns may be shared.
          const matches=qualified.length?qualified:descriptions.filter(description=>!description.qualified&&
            (parts.length===1||(description.c>c&&description.c<next&&(next!==Infinity||separateColumns))));
          if(matches.length===1)descriptionColumns.set(c,matches[0].c);
        });
      }
      for(const {c,cell,text,isPartHeader} of rowCells){
        if(referenceColumns.has(c))continue;
        if(isPartHeader){
          if(isHeadingRow){
            for(let x=c;x<c+cell.colSpan;x++)columns.add(x);
            // Also support key/value rows such as "Part Number | 123-456789-00".
            if(row[c+cell.colSpan]&&!header(row[c+cell.colSpan].textContent)&&/\d/.test(row[c+cell.colSpan].textContent))columns.add(c+cell.colSpan);
          }
          continue;
        }
        const candidates=new Set(text.match(/(?<![A-Z0-9._-])\d{2,4}-[A-Z0-9]{4,8}(?:-[A-Z0-9]{2,4})?(?![A-Z0-9._-])/gi)||[]);
        if(columns.has(c)&&cell.colSpan===1)for(const token of text.match(/\b[A-Z0-9]+(?:[-._][A-Z0-9]+)*\b/gi)||[]){if(token.length>=5&&/\d/.test(token)&&!/^\d{1,3}(?:\.\d+)?$/.test(token))candidates.add(token);}
        const descriptionCell=isHeadingRow?null:row[descriptionColumns.get(c)];
        const description=descriptionCell&&descriptionCell!==cell?cellText.get(descriptionCell).replace(/\s+/g,' ').trim():'';
        for(const value of candidates){
          const partNumber=value.toUpperCase();
          numbers.add(partNumber);
          occurrences.push({partNumber,table:tableIndex,row:r+1,column:c+1});
          if(description){
            if(!partDescriptions.has(partNumber))partDescriptions.set(partNumber,new Set());
            partDescriptions.get(partNumber).add(description);
          }
        }
      }
    }
  }
  return {partNumbers:[...numbers].sort(),occurrences,partDescriptions};
}

async function readArticle(item,signal){
  const url=new URL(item.url);if(url.origin!==location.origin)throw new Error('This document is outside MyLam; open its link to inspect it.');
  const response=await unsafeWindow.fetch(url.href,{credentials:'include',signal});
  if(!response.ok){const error=new Error(`HTTP ${response.status}`);error.pause=[401,403,429,503].includes(response.status);throw error}
  const type=response.headers.get('content-type')||'';if(!/html/i.test(type))throw new Error('This link is not an HTML article (it may be a PDF).');
  if(response.redirected&&/login|sign[-_]?in|saml|oauth/i.test(response.url)){const error=new Error('Session expired. Sign in to MyLam and retry.');error.pause=true;throw error}
  const html=await response.text();
  if(!html.includes('tech-document-article__container')&&/<input\b[^>]*\btype\s*=\s*["']?password\b/i.test(html)){const error=new Error('Session expired. Sign in to MyLam and retry.');error.pause=true;throw error}
  return extractPartNumbers(html);
}

async function readSummaryArticle(item,controller){
  let timedOut=false;
  const timer=setTimeout(()=>{timedOut=true;controller.abort()},SUMMARY_REQUEST_TIMEOUT);
  try{return await readArticle(item,controller.signal)}
  catch(error){if(timedOut)throw new Error('Request timed out.');throw error}
  finally{clearTimeout(timer)}
}

// An article can repeat one description with and without its "| CE!" mark, or with different spacing.
// Those count as one description and the most complete spelling is kept. Different wording stays separate.
function mergePartDescriptions(descriptions){
  const merged=new Map();
  for(const description of descriptions){
    const key=description.replace(/\s*\|\s*CE!?\s*$/i,'').replace(/\s+/g,'').toUpperCase();
    if(!merged.has(key)||description.length>merged.get(key).length)merged.set(key,description);
  }
  return [...merged.values()];
}

function* articlePartRows(item,result){
  const {documentNo,title,url,releaseDate,type,articleId,documentIndex}=item;
  for(const partNumber of result.partNumbers){
    yield {partNumber,description:mergePartDescriptions(result.partDescriptions.get(partNumber)||[]).join('\n'),documentNo,title,url,releaseDate,type,articleId,documentIndex};
  }
}

async function runSummaryQueue(queue,read,stopped){let next=0;const worker=async()=>{while(!stopped()&&next<queue.length)await read(queue[next++])}; await Promise.all(Array.from({length:Math.min(SUMMARY_CONCURRENCY,queue.length)},worker));}
function partNumberSearchUrl(partNumber){return KMMATRIX_ORIGIN+'/Search?q='+encodeURIComponent(partNumber);}

function lookupKmmatrixDescriptions(parts,signal,onResult){
  return new Promise((resolve,reject)=>{
    const id=crypto.randomUUID(),key=KMMATRIX_JOB_PREFIX+id;
    let index=0,tab,timer,closeAfterCancel,finished=false;
    function finish(error,waitForHelper=true){
      if(finished)return;
      finished=true;
      clearTimeout(timer);
      signal.removeEventListener('abort',cancel);
      window.removeEventListener('pagehide',cancel);
      const close=()=>{
        closeAfterCancel=null;
        clearTimeout(timer);
        GM_removeValueChangeListener(listener);
        GM_deleteValue(key);
        if(tab&&!tab.closed)tab.close();
        if(error)reject(error);else resolve();
      };
      if(waitForHelper&&error&&tab&&!tab.closed){
        // Let the helper remove its leave prompt before MyLam closes the tab.
        closeAfterCancel=close;
        timer=setTimeout(close,KMMATRIX_CLOSE_TIMEOUT);
        GM_setValue(key,{type:'cancel',expires:Date.now()+KMMATRIX_CLOSE_TIMEOUT});
      }else close();
    }
    const cancel=event=>finish(signal.reason||new DOMException('Lookup cancelled','AbortError'),event.type!=='pagehide');
    const listener=GM_addValueChangeListener(key,(_key,_old,reply,remote)=>{
      if(!remote)return;
      if(finished){if(reply?.type==='cancel-ready')closeAfterCancel?.();return}
      if(reply?.type!=='result'||reply.index!==index||reply.part!==parts[index])return;
      clearTimeout(timer);
      try{
        onResult(parts[index],typeof reply.description==='string'?reply.description:'',index+1,reply.timedOut===true);
        index++;
        if(index===parts.length)finish();else request().catch(finish);
      }catch(error){finish(error)}
    });
    async function request(){
      const timeout=index===0?KMMATRIX_SIGN_IN_TIMEOUT:KMMATRIX_REPLY_TIMEOUT;
      timer=setTimeout(()=>finish(new Error(index===0
        ?'KMMATRIX did not respond. Try again and complete sign-in if prompted.'
        :'The KMMATRIX tab stopped responding. Try again to continue.')),timeout);
      await GM.setValue(key,{type:'lookup',index,total:parts.length,part:parts[index],expires:Date.now()+timeout});
      if(finished){GM_deleteValue(key);return}
      if(!tab){
        tab=GM_openInTab(partNumberSearchUrl(parts[0])+'#kyra-part-descriptions='+id,{active:true,setParent:true});
        tab.onclose=()=>finish(new Error('KMMATRIX lookup stopped because its tab was closed. Run Find Tech Articles for BOM to continue.'));
      }
    }
    signal.addEventListener('abort',cancel,{once:true});
    window.addEventListener('pagehide',cancel,{once:true});
    try{
      signal.throwIfAborted();
      request().catch(finish);
    }catch(error){finish(error)}
  });
}

function readKmmatrixDescription(part){
  const descriptions=new Set();
  for(const cell of document.querySelectorAll('td.partnumber-column-table-search-page')){
    if(normalizedBomPartNumber(cell.textContent)!==part)continue;
    const headings=[...cell.closest('.k-grid')?.querySelectorAll('th')||[]];
    const number=headings.find(heading=>/^part number$/i.test(heading.textContent.trim()));
    const description=headings.find(heading=>/^description$/i.test(heading.textContent.trim()));
    if(!number||!description||number.parentElement!==description.parentElement)continue;
    const column=cell.cellIndex+description.cellIndex-number.cellIndex;
    const value=cell.parentElement.cells[column]?.textContent.replace(/\s+/g,' ').trim();
    if(value)descriptions.add(value);
  }
  if(!descriptions.size)return null;
  return descriptions.size===1?[...descriptions][0]:'';
}

function readKmmatrixApiDescription(payload,part){
  const records=[payload],descriptions=new Set();
  while(records.length){
    const record=records.pop();
    if(!record||typeof record!=='object')continue;
    if(typeof record.partnumber==='string'&&normalizedBomPartNumber(record.partnumber)===part&&typeof record.partdescription==='string'){
      const description=record.partdescription.replace(/\s+/g,' ').trim();
      if(description)descriptions.add(description);
    }
    for(const value of Object.values(record))if(value&&typeof value==='object')records.push(value);
  }
  return descriptions.size===1?[...descriptions][0]:'';
}

function runKmmatrixDescriptionTab(){
  const marker=new URLSearchParams(location.hash.slice(1)).get('kyra-part-descriptions');
  const id=marker||sessionStorage.getItem(KMMATRIX_TAB_KEY);
  if(!id||!/^[-a-f0-9]{36}$/i.test(id))return;
  const key=KMMATRIX_JOB_PREFIX+id,initial=GM_getValue(key);
  if(!initial||initial.expires<Date.now()){sessionStorage.removeItem(KMMATRIX_TAB_KEY);return}
  // The tab keeps its own marker while the site's SSO redirects back to KMMATRIX.
  sessionStorage.setItem(KMMATRIX_TAB_KEY,id);
  const notice=document.createElement('div');
  notice.id='kyra-kmmatrix-lookup-status';
  notice.setAttribute('role','status');
  notice.style.cssText='position:fixed;right:12px;bottom:12px;z-index:99999;padding:8px;border:1px solid #999;background:#fff;color:#222;font:13px sans-serif';
  const mountNotice=()=>document.body?.append(notice);
  let timer,pending,api,apiController;
  const xhr=unsafeWindow.XMLHttpRequest.prototype,open=xhr.open,setHeader=xhr.setRequestHeader,nativeRequests=new WeakMap();
  // Reuse the site's authorized part lookup in this tab. Only descriptions are sent back to MyLam.
  function captureOpen(method,target,...args){
    const result=open.call(this,method,target,...args);
    nativeRequests.delete(this);
    const request=pending;
    if(!request||!/^GET$/i.test(method))return result;
    const url=new URL(target,location.href);
    if(url.origin!==KMMATRIX_API_ORIGIN||url.pathname.toLowerCase()!=='/api/getpartinformation')return result;
    const parameter=[...url.searchParams].find(([,value])=>normalizedBomPartNumber(value)===request.part)?.[0];
    if(!parameter)return result;
    const native={url,parameter};
    nativeRequests.set(this,native);
    this.addEventListener('load',()=>{
      if(pending!==request||nativeRequests.get(this)!==native||!native.authorization||this.status!==200)return;
      try{
        const payload=this.responseType==='json'?this.response:JSON.parse(this.responseText);
        api=native;
        reply(request,readKmmatrixApiDescription(payload,request.part));
      }catch{/* The site can still provide the result in its table. */}
    },{once:true});
    return result;
  }
  function captureHeader(name,value){
    const result=setHeader.call(this,name,value);
    const native=nativeRequests.get(this);
    if(native&&/^authorization$/i.test(name)&&typeof value==='string')native.authorization=value;
    return result;
  }
  xhr.open=captureOpen;
  xhr.setRequestHeader=captureHeader;

  function confirmLeave(event){
    // KMMATRIX can redirect to sign-in before its search input is ready.
    if(!pending||pending.expires<=Date.now()||!document.getElementById('input-search-home'))return;
    event.preventDefault();
    event.returnValue=true;
  }

  function cleanup(){
    pending=null;api=null;
    clearInterval(timer);
    apiController?.abort();
    if(xhr.open===captureOpen)xhr.open=open;
    if(xhr.setRequestHeader===captureHeader)xhr.setRequestHeader=setHeader;
    GM_removeValueChangeListener(listener);
    document.removeEventListener('DOMContentLoaded',mountNotice);
    window.removeEventListener('beforeunload',confirmLeave);
    window.removeEventListener('pagehide',cleanup);
    notice.remove();
  }

  function reply(request,description,timedOut=false){
    if(pending!==request)return;
    pending=null;
    clearInterval(timer);
    const current=GM_getValue(key);
    if(Date.now()<request.expires&&current?.type==='lookup'&&current.index===request.index&&current.part===request.part){
      GM_setValue(key,{type:'result',index:request.index,part:request.part,description,timedOut,expires:request.expires});
    }
  }

  function searchWithSite(request){
    let started=0,previousEmpty;
    function check(){
      if(Date.now()>=request.expires){clearInterval(timer);return}
      const input=document.querySelector('#input-search-home');
      if(!input)return;
      const url=new URL(location.href);
      const empty=document.querySelector('.k-grid-norecords');
      if(!started){
        started=Date.now();
        if(url.pathname.toLowerCase()!=='/search'||url.searchParams.get('q')!==request.part){
          previousEmpty=empty;
          const target=partNumberSearchUrl(request.part);
          history.replaceState(history.state,'',target);
          unsafeWindow.dispatchEvent(new PopStateEvent('popstate',{state:history.state}));
          return;
        }
      }
      if(url.searchParams.get('q')!==request.part)return;
      const description=readKmmatrixDescription(request.part);
      if(!empty)previousEmpty=null;
      const noRecords=empty&&empty!==previousEmpty&&empty.getClientRects().length&&normalizedBomPartNumber(input.value)===request.part;
      if(description===null&&!noRecords&&Date.now()-started<KMMATRIX_RESULT_TIMEOUT)return;
      reply(request,description||'',description===null&&!noRecords);
    }
    timer=setInterval(check,250);
    check();
  }

  async function searchApi(request){
    const {url,parameter,authorization}=api;
    const target=new URL(url);
    target.searchParams.set(parameter,request.part);
    const controller=new unsafeWindow.AbortController();
    apiController=controller;
    const deadline=setTimeout(()=>controller.abort(),KMMATRIX_API_TIMEOUT);
    try{
      const response=await unsafeWindow.fetch(target.href,{
        credentials:'omit',redirect:'error',headers:{Authorization:authorization},signal:controller.signal,
      });
      if(pending!==request)return;
      if(response.status===401||response.status===403){
        api=null;
        searchWithSite(request);
      }else{
        const description=response.ok&&response.status!==204?readKmmatrixApiDescription(await response.json(),request.part):'';
        reply(request,description);
      }
    }catch{
      if(pending!==request)return;
      if(controller.signal.aborted){reply(request,'',true);return}
      api=null;
      searchWithSite(request);
    }finally{
      clearTimeout(deadline);
      if(apiController===controller)apiController=null;
    }
  }

  function receive(request){
    if(!request||request.type==='cancel'){
      sessionStorage.removeItem(KMMATRIX_TAB_KEY);
      cleanup();
      if(request?.type==='cancel'&&request.expires>Date.now())GM_setValue(key,{type:'cancel-ready',expires:request.expires});
      return;
    }
    if(request.type!=='lookup'||!KMMATRIX_PART_NUMBER.test(request.part)||request.expires<=Date.now())return;
    clearInterval(timer);
    pending=request;
    notice.textContent=`Checking part ${request.part}${request.total?` (${request.index+1}/${request.total})`:''}. Keep this tab open. It closes automatically.`;
    apiController?.abort();
    if(api)searchApi(request);else searchWithSite(request);
  }
  const listener=GM_addValueChangeListener(key,(_key,_old,request,remote)=>{if(remote)receive(request)});
  if(document.body)mountNotice();else document.addEventListener('DOMContentLoaded',mountNotice,{once:true});
  window.addEventListener('beforeunload',confirmLeave);
  window.addEventListener('pagehide',cleanup,{once:true});
  receive(initial);
}
function bomFileStem(fileName){let stem=String(fileName||'BOM').replace(/\.[^.]+$/,'').trim(),system=stem.search(/\s*-\s*SYSTEM\b/i); if(system>=0)stem=stem.slice(0,system).trim(); return (stem||'BOM').replace(/[<>:"/\\|?*]/g,'_').replace(/[. ]+$/g,'').slice(0,100)||'BOM';}

function analyzeBomMatches(boms,summaryRows){
  const differences=bomPartDifferences(boms),counts=boms.map(()=>0),unmatched=boms.map(()=>[]);
  const documents=new Map(),matched=new Set();
  for(const row of summaryRows){
    const part=normalizedBomPartNumber(row.partNumber),indexes=differences.get(part);
    if(!indexes)continue;
    matched.add(part);
    const key=row.articleId||row.url||`${row.documentNo}\n${row.title}`;
    let document=documents.get(key);
    if(!document){
      document={order:Number(row.documentIndex)||0,title:row.title||'',url:row.url||'',documentNo:row.documentNo||'',releaseDate:row.releaseDate||'',
        type:row.type||'',partsByBom:boms.map(()=>new Map())};
      documents.set(key,document);
    }
    for(const index of indexes)document.partsByBom[index].set(part,row);
  }
  for(const [part,indexes] of differences)for(const index of indexes){
    counts[index]++;
    if(!matched.has(part))unmatched[index].push(part);
  }
  return {counts,unmatched,matched,documents:[...documents.values()].sort((left,right)=>left.order-right.order||left.title.localeCompare(right.title))};
}

const compareBomLocations=new Intl.Collator('en',{numeric:true,sensitivity:'base'}).compare;
const systemBomLocation=['System'];
function bomLocations(bom,part){return bom.locations?.get(part)||systemBomLocation}
function bomLocationText(bom,part){return [...bomLocations(bom,part)].join(', ')}
function combinedBomLocations(positions){
  const unique=new Map();
  for(const position of positions)unique.set(position.toUpperCase(),position);
  return [...unique.values()].sort(compareBomLocations).join(', ');
}

function addBomSourceSheet(workbook,bom,name){
  const worksheet=workbook.addWorksheet(name),clean=value=>value===undefined||value===null||value===''?null:value instanceof Date?value:typeof value==='number'&&Number.isInteger(value)&&Math.abs(value)>=1e15?String(value):['string','number',
    'boolean'].includes(typeof value)?value:String(value);
  const merges=bom.merges.map(range=>XLSX.utils.decode_range(range)),mergedEnds=new Map();
  for(const {s,e} of merges)if(s.r<e.r)mergedEnds.set(s.r,Math.max(e.r,mergedEnds.get(s.r)||0));
  const blocks=[],locations=bom.partKeys.map(key=>key?bomLocationText(bom,key):''),lastRow=bom.rows.length-1;
  // Move rows connected by a vertical merge together. Headers and non-part blocks stay in place.
  for(let first=0;first<=lastRow;){
    let last=Math.min(lastRow,mergedEnds.get(first)??first);
    for(let row=first+1;row<=last;row++)last=Math.max(last,Math.min(lastRow,mergedEnds.get(row)??row));
    let location='';
    if(first>bom.headerRow){
      if(first===last)location=locations[first]||'';
      else{
        const positions=[];
        for(let row=first;row<=last;row++)if(bom.partKeys[row])positions.push(...bomLocations(bom,bom.partKeys[row]));
        location=combinedBomLocations(positions);
      }
    }
    blocks.push({first,last,location});first=last+1;
  }
  const sorted=blocks.filter(block=>block.location).sort((a,b)=>compareBomLocations(a.location,b.location));
  const order=[],destination=[];
  let next=0;
  for(const block of blocks){
    const {first,last}=block.location?sorted[next++]:block;
    for(let row=first;row<=last;row++){destination[row]=order.length;order.push(row)}
  }
  for(let start=0;start<order.length;start+=3000)worksheet.addRows(order.slice(start,start+3000).map(index=>[
    index===bom.headerRow?'Locations':locations[index]||null,...Array.from(bom.rows[index]||[],clean)
  ]));
  for(const {s,e} of merges)try{
    const row=destination[s.r]??s.r;
    worksheet.mergeCells(row+1,s.c+2,row+e.r-s.r+1,e.c+2);
  }catch(_){/* keep the remaining source table when a legacy merge is invalid */}
  const widths=[];
  measureWorkbookRow(widths,['Locations']);
  for(const location of new Set(locations))measureWorkbookRow(widths,[location]);
  for(const row of bom.rows)if(row)measureWorkbookRow(widths,row,1);
  fitWorkbookColumns(worksheet,widths);
  bom.rowHeights.forEach((height,index)=>{if(height)worksheet.getRow((destination[index]??index)+1).height=height}); return worksheet;
}

function addBomSourceSheets(workbook,boms){
  boms.forEach(bom=>{
    const base=bomFileStem(bom.fileName).replace(/[\[\]:*?/\\]/g,'_').replace(/^'+|'+$/g,'').slice(0,31)||'BOM';let name=base,index=2;
    while(name.toLowerCase()==='tech articles'||workbook.worksheets.some(sheet=>sheet.name.toLowerCase()===name.toLowerCase())){const suffix=` (${index++})`;name=base.slice(0,31-suffix.length)+suffix}
    addBomSourceSheet(workbook,bom,name);
  });
}

function setWorkbookLink(cell,text,hyperlink){
  cell.value={text,hyperlink};
  cell.font={name:'Arial',size:10,color:{argb:'FF0369A1'},underline:true};
}

function setWorkbookArticleLink(cell,{documentNo,url}){
  if(url)setWorkbookLink(cell,documentNo||'Open article',url);
}

// Columns fit their longest text up to a limit, like double-clicking a column border in Excel.
// Cells never wrap: longer text is cut off at the next cell.
const WORKBOOK_COLUMN_LIMIT=60,WORKBOOK_NARROW_CHARACTERS=' ,.:;!\'"|/\\()[]-Iijlftr',WORKBOOK_WIDE_CHARACTERS='MWmw@%';
function workbookTextWidth(value){
  // Arial 10 in Excel column units.
  let width=0;
  for(const char of String(value)){
    width+=WORKBOOK_NARROW_CHARACTERS.includes(char)?.55:WORKBOOK_WIDE_CHARACTERS.includes(char)?1.6:char>='A'&&char<='Z'?1.35:char>'⹿'?2.1:1.06;
    if(width>WORKBOOK_COLUMN_LIMIT)break;
  }
  return width;
}
function measureWorkbookRow(widths,values,first=0){
  values.forEach((value,index)=>{
    if(value===null||value===undefined||value==='')return;
    const column=first+index;
    if(!(widths[column]>WORKBOOK_COLUMN_LIMIT))widths[column]=Math.max(widths[column]||0,value instanceof Date?11:workbookTextWidth(value));
  });
}
function fitWorkbookColumns(worksheet,widths){
  for(let column=0;column<widths.length;column++)
    worksheet.getColumn(column+1).width=Math.min(WORKBOOK_COLUMN_LIMIT,Math.max(8,Math.ceil((widths[column]||0)+2.5)));
}
// Several descriptions of one part share a line.
function workbookLine(value){return typeof value==='string'?value.replace(/\s*[\r\n]+\s*/g,'; '):value}

function setWorkbookColumns(page,row,headers){
  headers.forEach((header,index)=>{
    const cell=page.getCell(row,index+1);cell.value=header;
    cell.font={name:'Arial',size:10,bold:true,color:{argb:'FFFFFFFF'}};
    cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF526B78'}};
    cell.alignment={vertical:'middle',horizontal:'center'};
    cell.border={right:{style:'thin',color:{argb:'FFFFFFFF'}}};
  });
  // Bold headers are slightly wider than the same text in a data row.
  const widths=[];
  measureWorkbookRow(widths,headers.map(header=>header+'0'));
  return widths;
}

function addTechArticlesSheet(workbook,summaryRows){
  const columns=['Title','Document No.','Release Date','Part Number','Description'],worksheet=workbook.addWorksheet('Tech Articles',{views:[{state:'frozen',ySplit:4,showGridLines:false}]});worksheet.properties.defaultRowHeight=18;
  worksheet.getCell(1,1).value='Tech Articles';worksheet.getCell(1,1).font={name:'Arial',size:15,bold:true,color:{argb:'FF1F2937'}};worksheet.getCell(1,1).alignment={vertical:'middle'};worksheet.getRow(1).height=25;
  for(let column=1;column<=columns.length;column++)worksheet.getCell(1,column).border={bottom:{style:'thin',color:{argb:'FF94A3B8'}}};
  const documentCount=new Set(summaryRows.map(row=>row.articleId||row.url||`${row.documentNo}\n${row.title}`)).size,partNumberCount=new Set(summaryRows.map(row=>normalizedBomPartNumber(row.partNumber)).filter(Boolean)).size;
  [['Documents',documentCount],['Unique Part Numbers',partNumberCount]].forEach((item,index)=>{const start=index*2+1,label=worksheet.getCell(2,start),value=worksheet.getCell(2,start+1);label.value=item[0];label.font={name:'Arial',size:10,bold:true,color:{argb:'FF475569'}};label.alignment={vertical:'middle'};
    value.value=item[1];value.font={name:'Arial',size:10,color:{argb:'FF475569'}}});worksheet.getCell(2,5).value=`Links: ${summaryRows.length}`;worksheet.getCell(2,5).font={name:'Arial',size:10,color:{argb:'FF475569'}};
  worksheet.getRow(2).height=28;const widths=setWorkbookColumns(worksheet,4,columns);worksheet.getRow(4).height=28;const edge={style:'thin',color:{argb:'FFD6E0EA'}};
  measureWorkbookRow(widths,['Documents0',documentCount,'Unique Part Numbers0',partNumberCount]);
  [...summaryRows].sort((left,right)=>(Number(left.documentIndex)||0)-(Number(right.documentIndex)||0)||String(left.partNumber).localeCompare(String(right.partNumber))).forEach(document=>{
    // An empty text cell stops the text before it from running across the row.
    const values=[document.title,document.documentNo,document.releaseDate,document.partNumber,document.description].map(value=>workbookLine(value)||''),row=worksheet.addRow(values);
    measureWorkbookRow(widths,values);
    row.eachCell({includeEmpty:true},cell=>{cell.font={name:'Arial',size:10,color:{argb:'FF1F2937'}};cell.alignment={vertical:'middle',horizontal:'left'};cell.border={bottom:edge}});
    setWorkbookArticleLink(row.getCell(2),document);
    if(document.partNumber)setWorkbookLink(row.getCell(4),document.partNumber,partNumberSearchUrl(document.partNumber));
  });
  fitWorkbookColumns(worksheet,widths);
  return worksheet;
}

async function createDownloadWorkbook(boms,analysis,summaryRows,{identity,unread=0,unreadLocations=0}={}){
  const ExcelJS=loadExcelJs();
  const workbook=new ExcelJS.Workbook();workbook.creator='Toolkit Kyra';workbook.company='Lam Research';workbook.created=new Date();
  if(identity){workbook.description=identity;workbook.keywords=`${unread} ${unreadLocations}`}
  if(analysis){
    const articleColumns=['Document No.','Release Date','Type','Title'];
    const bomColumns=boms.map(bom=>[bomFileStem(bom.fileName),'Description']),locationColumn='Locations';
    const edge={style:'thin',color:{argb:'FFD6E0EA'}};
    function writeParts(worksheet,widths,keysByBom,metadata,memo){
      // An empty text cell stops the text before it from running across the row.
      const partColumns=memo?3:2,emptyPart=Array(partColumns).fill(''),emptyMetadata=metadata.map(()=>null);
      const parts=keysByBom.map((keys,index)=>{
        const bom=boms[index];
        return Array.from(keys,key=>{
          const description=memo?memo.partsByBom[index].get(key).description:bom.descriptions.get(key);
          const values=[bom.materials.get(key),workbookLine(description)||''];
          if(memo)values.push(bomLocationText(bom,key));
          return values;
        }).sort(([a],[b])=>a<b?-1:a>b?1:0);
      });
      const count=Math.max(0,...parts.map(list=>list.length)),start=worksheet.rowCount+1;
      for(let index=0;index<count;index++){
        const values=[...(index===0?metadata:emptyMetadata),...parts.flatMap(list=>list[index]||emptyPart)],row=worksheet.addRow(values);
        measureWorkbookRow(widths,values);
        row.eachCell({includeEmpty:true},cell=>{cell.font={name:'Arial',size:10,color:{argb:'FF1F2937'}};cell.alignment={vertical:'middle',horizontal:'left'};cell.border={bottom:edge}});
        parts.forEach((list,offset)=>{
          const column=metadata.length+1+offset*partColumns,part=list[index]?.[0],cell=row.getCell(column);
          const fill={type:'pattern',pattern:'solid',fgColor:{argb:offset%2?'FFECFDF5':'FFEFF6FF'}};
          for(let extra=0;extra<partColumns;extra++)row.getCell(column+extra).fill=fill;
          if(part)setWorkbookLink(cell,part,partNumberSearchUrl(normalizedBomPartNumber(part)));
        });
      }
      if(count>1)for(let column=1;column<=metadata.length;column++)worksheet.mergeCells(start,column,start+count-1,column);
      if(memo)setWorkbookArticleLink(worksheet.getCell(start,1),memo);
    }
    const related=workbook.addWorksheet('Part Changes related to TM_PCN',{views:[{state:'frozen',ySplit:1,showGridLines:false}]});
    const relatedWidths=setWorkbookColumns(related,1,[...articleColumns,...bomColumns.flatMap(columns=>[...columns,locationColumn])]);
    const articles=analysis.documents.map(memo=>{
      const positions=[];
      memo.partsByBom.forEach((parts,index)=>{
        for(const key of parts.keys())positions.push(...bomLocations(boms[index],key));
      });
      return {memo,location:combinedBomLocations(positions)};
    }).sort((a,b)=>compareBomLocations(a.location,b.location));
    for(const {memo} of articles)writeParts(related,relatedWidths,memo.partsByBom.map(parts=>parts.keys()),[memo.documentNo,memo.releaseDate,memo.type,workbookLine(memo.title)],memo);
    fitWorkbookColumns(related,relatedWidths);
    const unrelated=workbook.addWorksheet('Part Changes without TM_PCN',{views:[{state:'frozen',ySplit:1,showGridLines:false}]});
    const unrelatedWidths=setWorkbookColumns(unrelated,1,[locationColumn,...bomColumns.flat()]);
    const locationGroups=new Map();
    analysis.unmatched.forEach((parts,index)=>{
      for(const key of parts){
        const location=bomLocationText(boms[index],key),groupKey=location.toUpperCase();
        if(!locationGroups.has(groupKey))locationGroups.set(groupKey,{location,keysByBom:boms.map(()=>[])});
        locationGroups.get(groupKey).keysByBom[index].push(key);
      }
    });
    for(const {location,keysByBom} of [...locationGroups.values()].sort((a,b)=>compareBomLocations(a.location,b.location)))
      writeParts(unrelated,unrelatedWidths,keysByBom,[location]);
    fitWorkbookColumns(unrelated,unrelatedWidths);
  }
  addBomSourceSheets(workbook,boms);
  addTechArticlesSheet(workbook,summaryRows||[]); return workbook.xlsx.writeBuffer();
}

function downloadBomWorkbook(buffer,fileName){
  const url=URL.createObjectURL(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})),anchor=document.createElement('a'); anchor.href=url;anchor.download=fileName;document.body.append(anchor);anchor.click();anchor.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
}

function showSummary(items){
  cancelSummary();document.getElementById(DIALOG_ID)?.remove(); const dialog=document.createElement('dialog'),rows=[],controllers=new Set(),missingDescriptions=new Map();
  let lookupController,running=false,scanComplete=false,stopped=false,message='',comparisonQueued=false,comparisonRunning=false,downloadRunning=false,
    latestBomResult=null; items.forEach(item=>Object.assign(item,{status:'pending',partNumbers:[],occurrences:[],error:''})); unsafeWindow.myLamSummary=items;unsafeWindow.myLamPartNumbers=rows; dialog.id=DIALOG_ID;dialog.style.cssText=
    'width:96vw;max-width:1500px;height:88vh;padding:20px;border:1px solid #cbd5e1;border-radius:10px;background:#fff;color:#222;box-shadow:0 4px 24px #0005';
  dialog.innerHTML=(`<style>
    #${DIALOG_ID}{box-sizing:border-box;overflow:hidden}#${DIALOG_ID}::backdrop{background:#0006}#${DIALOG_ID} [hidden]{display:none!important}#${DIALOG_ID} button{cursor:pointer}#${DIALOG_ID} button:disabled{cursor:not-allowed;opacity:.55}
    #${DIALOG_ID} table{width:100%;border-collapse:collapse;font:13px Arial,sans-serif}#${DIALOG_ID} th,#${DIALOG_ID} td{padding:8px;text-align:left;border-bottom:1px solid #dce3e8;overflow-wrap:anywhere}#${DIALOG_ID} th{position:sticky;top:0;background:#eef5f7}#${DIALOG_ID} a{color:#007a95}
    #${DIALOG_ID} .kyra-bom-panel{flex:none;padding:11px;border:1px solid #cbd5e1;border-radius:8px;background:#f8fafc}#${DIALOG_ID} .kyra-bom-panel>summary{cursor:pointer;color:#334155;font:600 13px Arial,sans-serif}#${DIALOG_ID} .kyra-bom-inputs{display:grid;grid-template-columns:minmax(220px,1fr) minmax(220px,1fr) auto;gap:10px;align-items:end;margin-top:12px}
    #${DIALOG_ID} .kyra-bom-file{display:grid;gap:5px;color:#334155;font:600 12px Arial,sans-serif}#${DIALOG_ID} .kyra-bom-file input{min-width:0;font:12px Arial,sans-serif}
    #${DIALOG_ID} [data-action="compare"],#${DIALOG_ID} [data-action="download"]{min-height:32px;border:1px solid #526b78;border-radius:6px;background:#526b78;color:#fff;font-weight:700;padding:6px 12px;white-space:nowrap}#${DIALOG_ID} [data-action="download"]{border-color:#0f766e;background:#0f766e}#${DIALOG_ID} [data-bom-error]{color:#b91c1c;font:12px Arial,sans-serif}#${DIALOG_ID} [data-bom-error]:empty{display:none}
    #${DIALOG_ID} .kyra-comparison{display:flex;flex-direction:column;flex:2;min-height:0;border:1px solid #cbd5e1;border-radius:8px;background:#fff;overflow:hidden}#${DIALOG_ID} .kyra-comparison-heading{padding:10px 12px;background:#f1f5f9;color:#334155;font:700 14px Arial,sans-serif}#${DIALOG_ID} .kyra-comparison-content{display:flex;flex-direction:column;flex:1;min-height:0;gap:8px;padding:9px}#${DIALOG_ID} .kyra-comparison-metrics{display:flex;gap:7px;flex-wrap:wrap}#${DIALOG_ID} .kyra-comparison-metrics span{border-radius:999px;background:#e2e8f0;color:#334155;padding:4px 8px;font:11px Arial,sans-serif}#${DIALOG_ID} .kyra-comparison-table{flex:1;min-height:0;overflow:auto;border:1px solid #dce3e8;border-radius:6px}#${DIALOG_ID} .kyra-comparison-table table{min-width:1050px}#${DIALOG_ID} .kyra-comparison-table th{white-space:nowrap}#${DIALOG_ID} .kyra-comparison-table th:first-child{width:50%;min-width:350px}#${DIALOG_ID} .kyra-comparison-table th:nth-child(2){min-width:105px}#${DIALOG_ID} .kyra-comparison-table th:nth-child(3){min-width:120px}#${DIALOG_ID} .kyra-comparison-table th:nth-child(n+5){min-width:210px}#${DIALOG_ID} .kyra-comparison-table td{vertical-align:top}#${DIALOG_ID} .kyra-comparison-table td:not(:first-child){white-space:pre;overflow-wrap:normal;word-break:normal}#${DIALOG_ID} .kyra-bom-first{background:#eff6ff}#${DIALOG_ID} .kyra-bom-second{background:#ecfdf5}#${DIALOG_ID} [data-comparison-empty]{padding:10px;color:#64748b;font:12px Arial,sans-serif}
    #${DIALOG_ID} .kyra-documents{flex:none;border:1px solid #dce3e8;border-radius:6px;overflow:auto}#${DIALOG_ID} .kyra-documents[open]{flex:1;min-height:150px}#${DIALOG_ID} .kyra-documents>summary{position:sticky;top:0;z-index:1;padding:10px;background:#f8fafc;color:#334155;cursor:pointer;font:600 13px Arial,sans-serif}
    #${DIALOG_ID} .kyra-documents td{vertical-align:top}#${DIALOG_ID} .kyra-documents td:nth-child(2){white-space:pre-line}
    #${DIALOG_ID} .kyra-article-status{padding:8px 10px;font-size:13px}
    #${DIALOG_ID} .kyra-article-actions{display:flex;flex-wrap:wrap;align-items:center;gap:7px}
    #${DIALOG_ID} .kyra-article-actions button{border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;background:#f8fafc;color:#334155;font-weight:600}
    #${DIALOG_ID} [data-read-errors],#${DIALOG_ID} [data-description-status]{color:#475569;margin-top:4px;font-size:12px}
    #${DIALOG_ID} [data-read-errors]:empty,#${DIALOG_ID} [data-description-status]:empty{display:none}
    @media(max-width:1050px){#${DIALOG_ID} .kyra-bom-inputs{grid-template-columns:1fr 1fr}#${DIALOG_ID} [data-action="compare"],#${DIALOG_ID} [data-action="download"]{justify-self:start}}@media(max-width:700px){#${DIALOG_ID} .kyra-bom-inputs{grid-template-columns:1fr}}
    </style>
    <div style="display:flex;flex-direction:column;gap:12px;height:100%"><div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap"><strong style="flex:1;font-size:20px">BOM - Part Number Analysis</strong><button ` +
  `type="button" data-action="download" disabled>Download Excel</button><button type="button" data-action="close">Close</button></div>
    <details data-bom-panel class="kyra-bom-panel" open><summary>BOM files <span data-bom-names></span></summary><div class="kyra-bom-inputs"><label class="kyra-bom-file"><span data-bom-label>BOM1 - SYSTEM</span><input ` +
  `data-bom="1" type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv"></label><label class="kyra-bom-file"><span data-bom-label>BOM2 - SYSTEM</span><input data-bom="2" type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv"></label><button ` +
  `type="button" data-action="compare">${SEARCH_LABEL}</button></div></details><div data-bom-error role="alert"></div>
    <section data-comparison class="kyra-comparison" hidden><div class="kyra-comparison-heading">Comparison result <span data-comparison-summary></span></div><div class="kyra-comparison-content"><div data-comparison-metrics ` +
  `class="kyra-comparison-metrics"></div><div data-description-status role="status"></div><div data-comparison-empty hidden></div><div data-comparison-table class="kyra-comparison-table" hidden><table><thead data-comparison-head></thead><tbody data-comparison-body></tbody>` +
  `</table></div></div></section>
    <details data-summary-list class="kyra-documents" open><summary>Tech Articles</summary><div class="kyra-article-status"><div class="kyra-article-actions"><span data-progress role="status"></span>
    <button type="button" data-action="retry" hidden>Retry failed articles</button></div>
    <div data-read-errors></div></div><table><thead><tr><th>Part Number</th><th>Description</th><th>Document ` +
  `No.</th><th>Title</th><th>Release Date</th></tr></thead><tbody data-summary-body></tbody></table></details></div>`);
  const button=name=>dialog.querySelector(`[data-action="${name}"]`),progress=dialog.querySelector('[data-progress]'),body=dialog.querySelector('[data-summary-body]'),bomInputs=[...dialog.querySelectorAll('[data-bom]')],bomPanel=dialog.querySelector('[data-bom-panel]'),
    summaryList=dialog.querySelector('[data-summary-list]'),bomError=dialog.querySelector('[data-bom-error]'),comparisonPanel=dialog.querySelector('[data-comparison]'),comparisonSummary=dialog.querySelector('[data-comparison-summary]'),
    comparisonMetrics=dialog.querySelector('[data-comparison-metrics]'),comparisonEmpty=dialog.querySelector('[data-comparison-empty]'),comparisonTable=dialog.querySelector('[data-comparison-table]'),comparisonHead=dialog.querySelector('[data-comparison-head]'),
    comparisonBody=dialog.querySelector('[data-comparison-body]');
  const retryButton=button('retry'),readErrors=dialog.querySelector('[data-read-errors]'),descriptionStatus=dialog.querySelector('[data-description-status]');
  const link=(url,label)=>{const a=document.createElement('a');a.href=url;a.target='_blank';a.rel='noopener noreferrer';a.textContent=label||url;return a}; const selectedBomFiles=()=>bomInputs.map(input=>input.files?.[0]).filter(Boolean);
  const setBomError=text=>{bomError.textContent=text};
  function clearBomComparison(){
    descriptionStatus.textContent='';
    latestBomResult=null;setBomError('');comparisonPanel.hidden=true;comparisonSummary.textContent='';comparisonMetrics.replaceChildren();comparisonHead.replaceChildren();comparisonBody.replaceChildren();comparisonEmpty.hidden=true;comparisonTable.hidden=true;bomPanel.open=true;
      summaryList.open=true;
  }
  function renderBomComparison(result){
    const {boms,analysis}=result,two=boms.length===2,names=boms.map(bom=>bomFileStem(bom.fileName)),headers=['Title','Document No.','Release Date','Type',...names]; comparisonPanel.hidden=false;bomPanel.open=false;summaryList.open=false;comparisonSummary.textContent=
      `— ${analysis.documents.length} Tech Article${analysis.documents.length===1?'':'s'}`;comparisonMetrics.replaceChildren();
    const metrics=names.map((name,index)=>`${name} ${two?'differences':'materials'}: ${analysis.counts[index]}`);
    metrics.push(`Matching Part Numbers: ${analysis.matched.size}`);
    metrics.forEach(text=>{const span=document.createElement('span');span.textContent=text;comparisonMetrics.append(span)}); const headerRow=document.createElement('tr');headers.forEach(text=>{const th=document.createElement('th');th.textContent=text;headerRow.append(th)});
      comparisonHead.replaceChildren(headerRow);comparisonBody.replaceChildren();
    analysis.documents.forEach(memo=>{
      const tr=document.createElement('tr');
      tr.insertCell().append(link(memo.url,memo.title));
      for(const value of [memo.documentNo,memo.releaseDate,memo.type])tr.insertCell().textContent=value;
      memo.partsByBom.forEach((parts,index)=>{
        const td=tr.insertCell();
        td.className=index?'kyra-bom-second':'kyra-bom-first';
        Array.from(parts.keys(),key=>boms[index].materials.get(key)).sort().forEach((part,partIndex)=>{
          if(partIndex)td.append('\n');
          td.append(link(partNumberSearchUrl(normalizedBomPartNumber(part)),part));
        });
      });
      comparisonBody.append(tr);
    });
    comparisonEmpty.textContent=analysis.documents.length?'':'No Tech Articles matched the relevant BOM Part Numbers.';comparisonEmpty.hidden=Boolean(analysis.documents.length);comparisonTable.hidden=!analysis.documents.length;
  }
  function update(){
    const done=items.filter(item=>['done','empty','error'].includes(item.status)).length,failures=items.filter(item=>item.status==='error'),failed=failures.length;
    progress.textContent=`${done}/${items.length} documents · ${new Set(rows.map(row=>row.partNumber)).size} unique part numbers · ${failed} failed${message?' — '+message:''}`; const files=selectedBomFiles(),compareButton=button('compare'),downloadButton=button('download');
    const busy=comparisonRunning||downloadRunning;
    compareButton.disabled=busy||comparisonQueued||!files.length||(!running&&!scanComplete);
    compareButton.textContent=SEARCH_LABEL;
    downloadButton.disabled=busy||running||!scanComplete||Boolean(files.length&&!latestBomResult);
    downloadButton.textContent=downloadRunning?'Creating Excel...':'Download Excel';
    bomInputs.forEach(input=>input.disabled=busy);
    retryButton.hidden=running||(!failed&&scanComplete);
    retryButton.disabled=busy;
    retryButton.textContent=scanComplete?'Retry failed articles':'Resume MyLam scan';
    readErrors.textContent=running?'':[...new Set(failures.map(item=>item.error))].join(' · ');
  }
  function appendResults(item){
    const fragment=document.createDocumentFragment();
    for(const result of articlePartRows(item,item)){
      const {partNumber,description}=result;
      rows.push(result);
      const tr=document.createElement('tr');
      for(const [value,url] of [[partNumber,partNumberSearchUrl(partNumber)],[description],[item.documentNo],[item.title,item.url],[item.releaseDate]]){
        const td=document.createElement('td');
        if(url)td.append(link(url,value));else td.textContent=value;
        tr.append(td);
      }
      const key=normalizedBomPartNumber(partNumber);
      if(!description&&KMMATRIX_PART_NUMBER.test(key)){
        if(!missingDescriptions.has(key))missingDescriptions.set(key,[]);
        missingDescriptions.get(key).push({result,cell:tr.cells[1]});
      }
      fragment.append(tr);
    }
    body.append(fragment);
  }
  const stop=()=>{stopped=true;controllers.forEach(controller=>controller.abort());lookupController?.abort()};cancelSummary=stop;
  async function fillMissingDescriptions(parts){
    const total=parts.length;
    let filled=0,timedOut=0;
    lookupController=new AbortController();
    descriptionStatus.textContent='Opening KMMATRIX. Complete sign-in in its tab if prompted.';
    try{
      await lookupKmmatrixDescriptions(parts,lookupController.signal,(part,description,checked,expired)=>{
        if(description){
          for(const {result,cell} of missingDescriptions.get(part)){
            result.description=description;
            cell.textContent=description;
            cell.title='From KMMATRIX';
          }
          missingDescriptions.delete(part);
          filled++;
        }
        if(expired)timedOut++;
        descriptionStatus.textContent=`KMMATRIX ${checked}/${total} checked · ${filled} filled${timedOut?` · ${timedOut} timed out`:''}`;
      });
      descriptionStatus.textContent=`KMMATRIX finished · ${filled}/${total} filled${timedOut?` · ${timedOut} timed out`:''}`;
    }catch(error){
      descriptionStatus.textContent=error.name==='AbortError'?'KMMATRIX lookup cancelled.':error.message;
    }finally{lookupController=null}
  }
  async function startBomComparison(lookupDescriptions){
    const files=selectedBomFiles();if(!files.length||comparisonRunning||downloadRunning||stopped)return;
    if(running){comparisonQueued=true;update();return}
    if(!scanComplete){setBomError('Resume the MyLam scan before comparing BOM files.');update();return}
    comparisonQueued=false;comparisonRunning=true;clearBomComparison();update();
    try{
      const boms=[];
      for(const file of files)boms.push(await readBomFile(file));
      if(stopped||!dialog.open)return;
      const analysis=analyzeBomMatches(boms,rows),outputName=boms.map(bom=>bomFileStem(bom.fileName)).join(' vs ')+'.xlsx';
      latestBomResult={boms,analysis,outputName};
      renderBomComparison(latestBomResult);
      const parts=lookupDescriptions&&boms.length===2?[...analysis.matched].filter(part=>missingDescriptions.has(part)).sort():[];
      if(parts.length)await fillMissingDescriptions(parts);
    }
      catch(error){clearBomComparison();console.error('[MyLam BOM Analysis]',error);setBomError(error.message||String(error))}
    finally{comparisonRunning=false;update()}
  }
  async function downloadBomResult(){
    const files=selectedBomFiles(),result=files.length?latestBomResult:null;if(!scanComplete||(files.length&&!result))return;const outputName=result?result.outputName:'MyLam_Tech_Articles.xlsx';downloadRunning=true;setBomError('');update();
    try{downloadBomWorkbook(await createDownloadWorkbook(result?result.boms:[],result?result.analysis:null,rows),outputName)} catch(error){console.error('[MyLam BOM Analysis]',error);setBomError(error.message||String(error))} finally{downloadRunning=false;update()}
  }
  async function readOne(item){
    if(stopped)return;item.status='reading';item.error='';const controller=new unsafeWindow.AbortController();controllers.add(controller);update();
    try{Object.assign(item,await readSummaryArticle(item,controller));item.status=item.partNumbers.length?'done':'empty';appendResults(item)}
    catch(error){
      const paused=Boolean(error.pause);item.status=stopped&&!paused?'pending':'error';item.error=item.status==='error'?error.message:'';
      if(paused&&!stopped){stopped=true;message='Check MyLam access, then resume the scan.';controllers.forEach(active=>active.abort())}
    }finally{controllers.delete(controller);update()}
  }
  async function run(){
    if(running||comparisonRunning||downloadRunning)return;
    const refreshComparison=Boolean(latestBomResult);
    if(refreshComparison)clearBomComparison();
    scanComplete=false;running=true;stopped=false;message='';descriptionStatus.textContent='';update();
    try{await runSummaryQueue(items.filter(item=>['pending','error'].includes(item.status)),readOne,()=>stopped);}finally{
      running=false;scanComplete=!stopped&&!items.some(item=>['pending','reading'].includes(item.status));
      if(!scanComplete&&(comparisonQueued||refreshComparison)){comparisonQueued=false;setBomError('Resume the MyLam scan before comparing BOM files.')}
      update();
      if(scanComplete&&(comparisonQueued||refreshComparison))await startBomComparison(comparisonQueued);
    }
  }
  retryButton.onclick=run;
  button('compare').onclick=()=>startBomComparison(true);button('download').onclick=downloadBomResult;bomInputs.forEach(input=>input.addEventListener('change',()=>{
    clearBomComparison();const files=selectedBomFiles();if(!files.length)comparisonQueued=false;
    bomInputs.forEach((field,index)=>{field.previousElementSibling.textContent=field.files?.[0]?bomFileStem(field.files[0].name):`BOM${index+1} - SYSTEM`});dialog.querySelector('[data-bom-names]').textContent=files.map(file=>bomFileStem(file.name)).join(' vs '); update();
  }));
  button('close').onclick=()=>dialog.close();dialog.addEventListener('close',()=>{stop();dialog.remove()}); document.body.append(dialog);dialog.showModal();update();run();
}

function addSummaryButton(){
  if(document.getElementById(BUTTON_ID))return;
  const clear=document.querySelector('button.clearFilter[onclick="clearFilter()"]');if(!clear)return;
  const button=document.createElement('button');button.id=BUTTON_ID;button.type='button';button.className='primary-button'; button.textContent=BUTTON_LABEL;button.style.cssText=
    'display:inline-flex;align-items:center;justify-content:center;margin:0 0 0 8px;padding:0 24px;min-height:40px;white-space:nowrap;box-sizing:border-box;color:#fff';button.style.setProperty('background','#5b6f7a','important');
    button.style.setProperty('border-color','#5b6f7a','important');
  button.addEventListener('click',async()=>{button.disabled=true;button.textContent='Loading...'; try{showSummary(await collectDocuments())}catch(error){console.error('[MyLam BOM Analysis]',error);alert(`BOM - Part Number Analysis failed:\n${error.message}`)} finally{button.disabled=
    false;button.textContent=BUTTON_LABEL}});
  const row=clear.parentElement;row.style.display='flex';row.style.alignItems='center';row.style.flexWrap='wrap';row.style.rowGap='8px'; clear.insertAdjacentElement('afterend',button);
  row.querySelectorAll('button').forEach(control=>{control.style.marginTop='0';control.style.marginBottom='0'});
  requestAnimationFrame(()=>{
    const apply=row.querySelector('button.applyFilter'),height=apply?.getBoundingClientRect().height;
    if(apply)button.style.font=getComputedStyle(apply).font;
    if(height){button.style.setProperty('height',`${height}px`,'important');button.style.setProperty('min-height',`${height}px`,'important')}
  });
}

function partChangesMyLamPanel(){
  const panel=document.createElement('section');
  panel.id='kyra-part-changes-mylam';
  panel.style.cssText='position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:2147483647;box-sizing:border-box;'+
    'width:min(600px,calc(100vw - 32px));max-height:calc(100vh - 32px);overflow:auto;overflow-wrap:anywhere;padding:clamp(18px,4vw,30px);'+
    'border:1px solid #cbd5e1;border-top:5px solid #215fbb;border-radius:14px;background:#fff;color:#24364b;box-shadow:0 16px 52px #17272f42;font:16px/1.5 Segoe UI,sans-serif';
  const label=document.createElement('div'),title=document.createElement('h2'),status=document.createElement('p'),note=document.createElement('p'),dismiss=document.createElement('button');
  label.textContent='SAP Toolbox · '+BOM_COMPARE_LABEL;label.style.cssText='font-size:14px;font-weight:600;color:#315a94;margin-bottom:6px';
  title.id=panel.id+'-title';panel.setAttribute('aria-labelledby',title.id);
  title.textContent='BOM comparison in progress';title.style.cssText='margin:0;font:600 24px/1.25 Segoe UI,sans-serif;color:#1e3658';
  status.setAttribute('role','status');status.style.margin='16px 0 10px';
  note.textContent='Keep this MyLam tab open until the comparison finishes.';note.style.cssText='margin:0;font-weight:600';
  dismiss.type='button';dismiss.hidden=true;dismiss.textContent='Dismiss';
  dismiss.style.cssText='margin-top:16px;padding:7px 14px;border:1px solid #cbd5e1;border-radius:6px;background:#f6f8fe;color:#24364b;font:inherit;cursor:pointer';
  dismiss.addEventListener('click',()=>panel.remove());
  panel.append(label,title,status,note,dismiss);document.body.append(panel);
  return {
    update(message){status.textContent=message},
    finish(message,complete=false){
      title.textContent=complete?'BOM comparison finished':'BOM comparison needs attention';
      status.textContent=message;note.textContent='You can close this MyLam tab.';dismiss.hidden=false;
    }
  };
}

function selectPartChangesSystems(list,hints,override){
  const options=[...list.querySelectorAll('input.inner:not(:disabled)')].map(input=>({
    id:input.id,label:(input.name||input.nextElementSibling?.textContent||input.value).trim()
  }));
  const labels=[...new Set(options.map(option=>option.label).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'en',{numeric:true}));
  if(!labels.length)throw new Error('MyLam returned no System options.');
  if(JSON.stringify(labels)!==JSON.stringify(GM_getValue(PART_CHANGES_MYLAM_SYSTEM_KEY)))GM_setValue(PART_CHANGES_MYLAM_SYSTEM_KEY,labels);
  const suggest=createPartChangesSystemGuesser(labels);
  // Every BOM row adds its own Systems. Auto suggests one for each code in the System Description.
  const selected=new Set(override||hints.flatMap(hint=>partChangesSystemList(hint.myLamSystem).flatMap(label=>label?[label]:suggest(hint.system,hint.product))));
  const unavailable=[...selected].filter(label=>!labels.includes(label));
  if(unavailable.length)throw new Error('MyLam does not offer '+unavailable.join(', ')+'. Set FID System in Tasks and run the comparison again.');
  if(!selected.size)throw new Error('No MyLam System suggestion. Set FID System in Tasks and run the comparison again.');
  return options.filter(option=>selected.has(option.label));
}

function partChangesSetMyLamFilters(list,systems,from,to){
  const ids=new Set(systems.map(system=>system.id)),all=list.querySelector('input.checkAll');
  if(all?.checked)all.click();
  for(const input of list.querySelectorAll('input.inner:not(:disabled)')){
    if(input.checked!==ids.has(input.id))input.click();
  }
  const selected=[...list.querySelectorAll('input.inner:checked:not(:disabled)')];
  if(selected.length!==ids.size||selected.some(input=>!ids.has(input.id)))
    throw new Error('Could not apply the selected MyLam Systems.');
  for(const [id,date] of [['startDate',from],['endDate',to]]){
    const input=document.getElementById(id),match=/^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
    if(!input||!match)throw new Error('MyLam release-date filters are unavailable.');
    input.value=match[2]+'/'+match[3]+'/'+match[1];
    input.dispatchEvent(new unsafeWindow.Event('change',{bubbles:true}));
  }
}

async function scanPartChangesArticles(documents,from,to,progress,cache){
  const rows=[],controllers=new Set();
  let done=0,unread=0,paused=null;
  progress('Reading 0 / '+documents.length+' Tech Articles');
  await runSummaryQueue(documents,async item=>{
    let controller;
    try{
      const match=/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(item.releaseDate);
      const date=match?match[3]+'-'+match[1].padStart(2,'0')+'-'+match[2].padStart(2,'0'):
        /^\d{4}-\d{2}-\d{2}$/.test(item.releaseDate)?item.releaseDate:'';
      if(!date)throw new Error('Unrecognized release date: '+item.releaseDate);
      if(date<from||date>to)return;
      let result=cache.get(item.url);
      if(!result){
        controller=new unsafeWindow.AbortController();
        controllers.add(controller);
        result=await readSummaryArticle(item,controller);
        cache.set(item.url,result);
      }
      for(const row of articlePartRows(item,result))rows.push(row);
    }catch(error){
      if(!paused&&error.pause){paused=error;controllers.forEach(active=>active.abort())}
      if(!paused)unread++;
    }finally{
      controllers.delete(controller);
      done++;
      if(done===documents.length||done%10===0)progress('Reading '+done+' / '+documents.length+' Tech Articles',done/documents.length);
    }
  },()=>Boolean(paused));
  if(paused)throw paused;
  return {rows,unread};
}

async function fillPartChangesDescriptions(analysis,progress){
  const missing=new Map();
  for(const memo of analysis.documents)for(const parts of memo.partsByBom)for(const [key,row] of parts){
    if(row.description||!KMMATRIX_PART_NUMBER.test(key))continue;
    if(!missing.has(key))missing.set(key,new Set());
    missing.get(key).add(row);
  }
  if(!missing.size)return '';
  const parts=[...missing.keys()].sort();
  let timedOut=0;
  progress('Opening KMMATRIX. Complete sign-in if prompted.');
  try{
    await lookupKmmatrixDescriptions(parts,new AbortController().signal,(part,description,checked,expired)=>{
      if(description)for(const row of missing.get(part))row.description=description;
      if(expired)timedOut++;
      progress('KMMATRIX '+checked+' / '+parts.length+' checked',checked/parts.length);
    });
    return timedOut?'KMMATRIX: '+timedOut+' descriptions timed out':'';
  }catch(error){return error.message}
}

async function partChangesWorkbookInfo(files,path,identity){
  try{
    const data=await files('read',{path});
    const props=XLSX.read(data.contentBase64,{type:'base64',bookProps:true}).Props;
    if(props?.Comments===identity){
      const [unread=0,unreadLocations=0]=String(props.Keywords||'0').split(' ').map(Number);
      return {unread,unreadLocations};
    }
  }catch(_){/* The existing file cannot be identified as an output of this task. */}
  return null;
}

async function partChangesWorkbookOutput(files,task,group,name,identity){
  const {entries=[]}=await files('list',{path:task.spec.OutputFolder,pattern:'*.xlsx'});
  const byName=new Map(entries.map(entry=>[entry.name.toLowerCase(),entry]));
  const base=partChangesWorkbookPath(task,group,'',name).split(/[\\/]/).pop().toLowerCase();
  const suffix=new RegExp('_'+task.id+'(?:_\\d+)?\\.xlsx$','i');
  for(const entry of entries){
    if(entry.kind!=='file'||!entry.size)continue;
    const legacy=entry.name.startsWith(group+' - ')&&entry.name.endsWith(' '+task.id+'.xlsx');
    if(entry.name.toLowerCase()!==base&&!legacy&&!suffix.test(entry.name))continue;
    const existing=await partChangesWorkbookInfo(files,entry.path,identity);
    if(existing)return {path:entry.path,existing};
  }
  for(let index=0;;index++){
    const tail=index?'_'+task.id+(index>1?'_'+index:''):'';
    const path=partChangesWorkbookPath(task,group,tail,name),basename=path.split(/[\\/]/).pop();
    if(!byName.has(basename.toLowerCase()))return {path};
  }
}

function resolvePartChangesRows(task){
  const [fields=[],snapshot={}]=task.spec.Form.exportFidSource?JSON.parse(task.spec.Form.exportFidSource):[];
  return task.spec.Rows.map((row,index)=>{
    const result=task.results[index],{document,locationFiles}=partChangesResultDetails(row,result);
    const resolved={...row,file:result?.Status==='Completed'&&result.Files?.length===1?result.Files[0]:'',locationFiles};
    if(document){
      resolved.Quote=document.number;resolved.Fid=partChangesFidOf(document.version);resolved.Version=document.version;
      resolved.Fields={...row.Fields,documentType:document.type,created:document.created};
    }
    if(row.Fields.filenameValues){
      const values=JSON.parse(row.Fields.filenameValues),source=!row.Fid&&snapshot[resolved.Fid];
      if(source)fields.forEach((field,index)=>{if(source[index])values[field]=source[index]});
      resolved.filenameValues=values;
      resolved.order={type:partChangesRowType(resolved),version:resolved.Version};
      resolved.Fields={...resolved.Fields,customer:values.customer,system:values['system description'],
        product:resolved.Fields.product||values['primary product']||values['fid description']||''};
    }
    return resolved;
  });
}

async function readPartChangesBom(row,files){
  async function read(path,options){
    const data=await files('read',{path});
    const bytes=Uint8Array.from(atob(data.contentBase64),char=>char.charCodeAt(0));
    return readBomFile(new File([bytes],path.split(/[\\/]/).pop()),options);
  }
  const bom=await read(row.file,{requireDescription:true}),directory=row.file.replace(/[^\\/]+$/,''),byPart=new Map(),labels=new Map();
  const modules=(row.locationFiles||[]).map(name=>({name,label:partChangesModuleLocation(name)}));
  modules.sort((a,b)=>compareBomLocations(a.label,b.label));
  let unreadLocations=0;
  for(const {name,label} of modules)try{
    const module=await read(directory+name,{materialsOnly:true}),labelKey=label.toUpperCase();
    const position=labels.get(labelKey)||label;
    labels.set(labelKey,position);
    for(const part of module.materials.keys()){
      if(!bom.materials.has(part))continue;
      let positions=byPart.get(part);
      if(!positions)byPart.set(part,positions=new Set());
      positions.add(position);
    }
  }catch(error){unreadLocations++;console.warn('[SAP Toolbox BOM Compare] '+name,error.message)}
  return {...bom,locations:byPart,unreadLocations};
}

async function finishPartChangesTask(task,files,update,systemOverrides={}){
  const rows=resolvePartChangesRows(task),groups=partChangesGroups(rows),standalone=rows.filter(row=>row.Group===ZBOM_COMPARE_STANDALONE);
  const deferred=new Map([...groups].filter(([,members])=>members.every(row=>row.file)&&members.some(row=>row.filenameValues)));
  const fixedNames=new Map([...groups].filter(([,members])=>members[0].FileName).map(([group,members])=>[group,members[0].FileName]));
  const names=plannedPartChangesFileNames(deferred,task.spec.Template||PART_CHANGES_DEFAULT_TEMPLATE,
    task.spec.Form.partChangesSeparator||ZBOM_COMPARE_DEFAULT_SEPARATOR,fixedNames);
  const boms=new Map(),cache=new Map(),reports=new Map(),issues=[];
  let skipped=0,unread=0,unreadLocations=0,position=0;
  for(const [group,members] of groups){
    const started=position++;
    if(members.some(row=>!row.file)){skipped++;continue}
    // Every group takes an equal share of the task: BOMs, articles, descriptions, then the workbook.
    const progress=(message,share)=>update(group+' · '+message,(started+share)/groups.size);
    try{
      const name=names.get(group),filenameError=deferred.has(group)&&zbomFileNameError(name,'partChanges');
      if(filenameError)throw new Error(filenameError);
      const override=systemOverrides[group];
      if(override?.rows&&override.rows.length!==members.length)throw new Error('Select FID Systems for the current BOM rows in Tasks.');
      const identity=await partChangesWorkbookIdentity(task.id,group,override);
      const {path,existing}=await partChangesWorkbookOutput(files,task,group,name,identity);
      if(existing){reports.set(group,path);unread+=existing.unread;unreadLocations+=existing.unreadLocations;continue}
      const groupBoms=[];
      for(const row of members){
        if(!boms.has(row.file)){
          progress('Reading '+partChangesDocumentLabel(partChangesRowType(row))+' '+row.Quote,0);
          boms.set(row.file,await readPartChangesBom(row,files));
        }
        groupBoms.push(boms.get(row.file));
      }
      const from=members.map(row=>row.Fields.created.slice(0,10)).sort()[0],to=task.spec.Form.partChangesThrough;
      if(!/^\d{4}-\d{2}-\d{2}$/.test(from)||!/^\d{4}-\d{2}-\d{2}$/.test(to)||from>to)
        throw new Error('The SAP creation dates do not give a valid article search range.');
      progress('Waiting for Tech Articles filters',.05);
      const ready=()=>document.querySelector('#slinesSelections input.inner')&&document.querySelector('button#apply.applyFilter:not(:disabled)')&&
        [unsafeWindow.jQuery,unsafeWindow.$].some(jq=>jq?.fn?.dataTable?.isDataTable('#searchResults'));
      for(let attempt=0;attempt<60&&!ready();attempt++)await new Promise(resolve=>setTimeout(resolve,500));
      if(!ready()||document.getElementById('contentList')?.selectedOptions[0]?.textContent.trim()!=='Tech Articles')
        throw Object.assign(new Error('Tech Articles filters are unavailable. Sign in to MyLam and retry.'),{pause:true});
      const hints=members.map(({Fields},index)=>override?.rows?{...Fields,myLamSystem:override.rows[index]}:Fields);
      const list=document.getElementById('slinesSelections'),systems=selectPartChangesSystems(list,hints,Array.isArray(override)?override:undefined);
      partChangesSetMyLamFilters(list,systems,from,to);
      const labels=[...new Set(systems.map(system=>system.label))];
      progress('Finding Tech Articles for '+labels.join(', '),.1);
      const documents=await collectDocuments(),scan=await scanPartChangesArticles(documents,from,to,(message,read=0)=>progress(message,.1+.65*read),cache);

      const analysis=analyzeBomMatches(groupBoms,scan.rows);
      const descriptionIssue=await fillPartChangesDescriptions(analysis,(message,checked=0)=>progress(message,.75+.2*checked));
      const locationReadErrors=groupBoms.reduce((sum,bom)=>sum+bom.unreadLocations,0);
      progress('Writing comparison Excel',.95);
      const buffer=await createDownloadWorkbook(groupBoms,analysis,scan.rows,{identity,unread:scan.unread,unreadLocations:locationReadErrors});
      const contentBase64=await new Promise((resolve,reject)=>{
        const reader=new FileReader();
        reader.onload=()=>resolve(String(reader.result).split(',')[1]);
        reader.onerror=()=>reject(reader.error);
        reader.readAsDataURL(new Blob([buffer]));
      });
      const staged=task.spec.OutputFolder.replace(/[\\/]+$/,'')+'\\.kyra-'+crypto.randomUUID()+'.pending';
      try{
        await files('write',{path:staged,contentBase64});
        await files('move',{sourcePath:staged,destinationPath:path});
      }catch(error){
        if(!await partChangesWorkbookInfo(files,path,identity))throw error;
      }finally{try{await files('deleteFile',{path:staged})}catch(_){/* A successful move already removed the temporary file. */}}
      reports.set(group,path);unread+=scan.unread;unreadLocations+=locationReadErrors;
      if(descriptionIssue)issues.push(group+': '+descriptionIssue);
    }catch(error){
      issues.push(group+': '+error.message);
      if(error.pause)break;
    }
  }
  const saved=reports.size,exported=standalone.filter(row=>row.file).length;
  return {saved,files:Object.fromEntries(reports),complete:saved===groups.size,message:[saved+' / '+groups.size+' comparison files ready',
    standalone.length?exported+' / '+standalone.length+' standalone BOMs exported':'',
    skipped?skipped+' groups need a successful BOM export':'',unread?unread+' articles could not be read':'',
    unreadLocations?unreadLocations+' location files could not be read':'',...issues].filter(Boolean).join(' · ')};
}

async function runPartChangesMyLam({id,port,systems}){
  const {update,finish}=partChangesMyLamPanel(),bridge=createSapGuiControlClient(['http://127.0.0.1:'+port]);
  const route='/toolbox/tasks/'+id;
  update('Reading the SAP export task');
  try{
    await navigator.locks.request('toolkit-bom-compare:'+id,{ifAvailable:true},async lock=>{
      if(!lock){finish('This task is already running in another MyLam tab.');return}
      const {task}=await bridge.request('GET',route,undefined,10000);
      if(task?.id!==id||task.tool!=='partChanges'||!task.finishedAt||!task.completed||!partChangesGroups(task.spec.Rows).size)
        throw new Error('Wait for the SAP BOM exports to finish.');
      // Tasks follows the comparison step by step until the final report replaces it.
      let sent=0,reports=Promise.resolve();
      const report=text=>bridge.request('POST',route+'/comparison',text.slice(0,2000),10000).catch(()=>{});
      const progress=(text,share=0)=>{
        update(text);
        if(Date.now()-sent<3000)return;
        sent=Date.now();
        reports=reports.then(()=>report(PART_CHANGES_COMPARING+Math.floor(share*100)+'% · '+text));
      };
      const closing=()=>{void report('Comparison stopped because the MyLam tab closed. Retry from Tasks → '+PART_CHANGES_ACTION_LABEL+'.')};
      window.addEventListener('pagehide',closing);
      let result;
      try{result=await finishPartChangesTask(task,bridge.files,progress,systems)}
      catch(error){result={saved:0,message:error.message}}
      finally{window.removeEventListener('pagehide',closing)}
      await reports;
      const {saved,message,files={}}=result,phase=message.slice(0,2000);
      const summary=saved?message+' · '+task.spec.OutputFolder:message+' · Retry from Tasks → '+PART_CHANGES_ACTION_LABEL+'.';
      // Keep reporting the outcome to older SAPGUIcontrol releases that only accept text.
      try{await bridge.request('POST',route+'/comparison',task.comparisonFiles?{Message:phase,Files:files}:phase,10000)}
      catch(error){finish(summary+' · Could not update Tasks. '+error.message);return}
      finish(saved&&!task.comparisonFiles?summary+' · Update SAPGUIcontrol to see the comparison files in Tasks.':summary,result.complete);
      if(result.complete){
        GM_getTab(tab=>{
          if(tab?.[PART_CHANGES_MYLAM_TAB]?.id!==id)return;
          delete tab[PART_CHANGES_MYLAM_TAB];GM_saveTab(tab);
        });
        if(partChangesHandoff(location.hash)?.id===id)history.replaceState(history.state,'',location.pathname+location.search);
      }
    });
  }catch(error){finish(error.message+' · Retry from Tasks → '+PART_CHANGES_ACTION_LABEL+'.')}
}

function initManualSummary(){
  let dateInitialized=false;
  function updatePage(){if(!dateInitialized){const input=document.querySelector('#startDate.textInput.excludeXLS');if(input){dateInitialized=true;if(!input.value)input.value='01/01/'+(new Date().getFullYear()-3)}} addSummaryButton();}
  updatePage();new MutationObserver(updatePage).observe(document.body,{childList:true,subtree:true});
}

function init(){
  const handoff=partChangesHandoff(location.hash);
  if(handoff){void runPartChangesMyLam(handoff).finally(initManualSummary);return}
  GM_getTab(tab=>{
    const saved=tab?.[PART_CHANGES_MYLAM_TAB],request=saved&&partChangesHandoff('#'+PART_CHANGES_MYLAM_TAB+'='+saved.id+'&port='+saved.port);
    if(request)void runPartChangesMyLam({...request,systems:saved.systems}).finally(initManualSummary);
    else initManualSummary();
  });
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();

// ==================== MyLam · Chosen Filters ====================
(function(){
'use strict';
if(location.hostname!=='techinfo.mylam.com'||!location.pathname.startsWith('/portal/group/mylam/contenttyperesults'))return;
const jq=jQuery,filters=new Map();
if(!jq.fn.chosen)return;
const selector='.filterContainerClass ul.multiSelect,.filterContainerClass select:not(.kyra-chosen-input),select[id="contentList"],select[id="savedFilters"]';
const css=(`
.kyra-chosen-source,.kyra-chosen-source~.select2{display:none!important}.kyra-mylam-chosen{min-width:0;width:100%;margin-top:9px}.searchbox .kyra-mylam-chosen{width:220px;flex:0 1 220px;margin-top:0}
.kyra-mylam-chosen>select.kyra-chosen-input,.kyra-mylam-chosen>.select2-container{display:none!important}.kyra-mylam-chosen .chosen-container{width:100%!important}.kyra-mylam-chosen .chosen-with-drop{z-index:1100}
.kyra-mylam-chosen .chosen-container-single .chosen-single div b,.kyra-mylam-chosen .search-choice-close{background:none!important}.kyra-mylam-chosen .chosen-container-single .chosen-single div b:after{content:'▾';` +
  `display:block;text-align:center;line-height:24px}.kyra-mylam-chosen .search-choice-close:after{content:'×';display:block;color:#666;font-size:15px;line-height:11px;text-align:center}
`);

function readFilter(source){
  if(source.tagName==='SELECT')return {multiple:source.multiple,disabled:source.disabled,options:[...source.options].map(option=>({value:option.value,text:option.text,selected:option.selected,disabled:option.disabled}))};
  const wrapper=source.parentElement,label=wrapper.querySelector('.docLabels')?.textContent.trim()||'',selected=new Set(label.split(',').filter(Boolean));
  const all=label==='Select All'||!label,disabled=wrapper.classList.contains('select-disabled')||!!wrapper.querySelector('a.multi.disabled');
  return {multiple:true,disabled,options:[...source.querySelectorAll('input.inner')].map(input=>({value:input.value,text:input.name||input.nextElementSibling?.textContent.trim()||input.value,selected:!all&&selected.has(input.id),disabled:input.disabled}))};
}

function applySelection(source,select){
  if(source.tagName==='SELECT'){const values=new Set([...select.selectedOptions].map(option=>option.value));for(const option of source.options)option.selected=values.has(option.value); source.dispatchEvent(new unsafeWindow.Event('change',{bubbles:true}));return;}
  const values=new Set([...select.selectedOptions].map(option=>option.value)),inputs=[...source.querySelectorAll('input.inner')],all=source.querySelector('input.checkAll');
  if(!values.size){if(all&&!all.disabled){all.checked=false;all.click()}return}
  if(all?.checked)all.click();
  for(const input of inputs)if(!input.disabled&&input.checked!==values.has(input.value))input.click();
}

function syncFilters(){
  for(const [source,entry] of filters)if(!source.isConnected){jq(entry.select).chosen('destroy');entry.host.remove();filters.delete(source)}
  for(const source of document.querySelectorAll(selector)){
    const state=readFilter(source),signature=JSON.stringify(state);let entry=filters.get(source);
    if(!entry){
      const host=document.createElement('div'),select=document.createElement('select'),wrapper=source.tagName==='SELECT'?source.closest('.select-dropdown')||source:source.parentElement; host.className='kyra-mylam-chosen';select.className='kyra-chosen-input';select.multiple=
        state.multiple; const label=source.closest('.filterContainerClass');select.setAttribute('aria-label',label?[...label.childNodes].filter(node=>node.nodeType===3).map(node=>node.textContent.trim()).join(' ').trim():source.id==='contentList'?'Content Type':'Saved Filters');
        host.append(select);wrapper.insertAdjacentElement('afterend',host);entry={host,select,wrapper,signature:null};filters.set(source,entry); jq(select).on('change.kyraFilters',()=>{entry.signature=null;applySelection(source,select)});
    }
    if(entry.signature===signature)continue;
    const {select,wrapper}=entry;select.multiple=state.multiple;select.disabled=state.disabled; select.replaceChildren(...state.options.map(item=>{const option=new Option(item.text,item.value,false,item.selected);option.disabled=item.disabled;return option}));
    if(jq(select).data('chosen'))jq(select).trigger('chosen:updated'); else jq(select).chosen({width:'100%',search_contains:true});
    if(jq(select).data('chosen')){wrapper.classList.add('kyra-chosen-source'); if(source.tagName==='SELECT'&&wrapper===source){const old=source.previousElementSibling;if(old?.classList.contains('select-dropdown-handle'))old.classList.add('kyra-chosen-source')}}
    entry.signature=signature;
  }
}
function init(){
  const style=document.createElement('style');style.textContent=chosenBaseCss+css;document.head.append(style);syncFilters();
  // MyLam also changes checked/selected properties without DOM events (saved filters and resets).
  setInterval(()=>{if(!document.hidden)syncFilters()},2000);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();
})();
