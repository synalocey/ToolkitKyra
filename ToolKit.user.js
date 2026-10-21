// ==UserScript==
// @name         Toolkit Kyra
// @version      26.10.21
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
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @require      https://cdn.jsdelivr.net/npm/@zip.js/zip.js@2.24.0/dist/zip.min.js
// @resource     chosenCss https://cdn.jsdelivr.net/gh/JJJ/chosen@4.0.0/dist/css/chosen.min.css
// @resource     xlsxJs https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @match        *://*.lamresearch.com/*
// @match        *://*.mylam.com/*
// @match        *://*.github.io/*
// @run-at       document-start
// @noframes
// @grant        GM_registerMenuCommand
// @grant        GM_setClipboard
// @grant        GM_xmlhttpRequest
// @grant        GM_getResourceText
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// @connect      127.0.0.1
// @connect      raw.githubusercontent.com
// ==/UserScript==
/* globals jQuery, XLSX, ExcelJS, zip */
const chosenBaseCss=GM_getResourceText('chosenCss');

(function(){
'use strict';
const date_v = new Date('20' + GM_info.script.version); const APP_PATH='/k1.png'; const BRIDGE_PATH='/k2.png'; const FCID_PATH='/stacey1.png'; const NSR_HEADER_BUTTON_ID='toolkitKyraControlTowerButton';
const BRIDGE_URLS=[25972,25973,25974,8765].map(port=>'http://127.0.0.1:'+port); let bridgeUrl='',bridgeDiscovery=null;
const SAPGUICONTROL_DOWNLOAD_URL='https://github.com/synalocey/ToolkitKyra/raw/main/SAPGUIcontrol.zip'; const SAPGUICONTROL_VERSION_URL='https://raw.githubusercontent.com/synalocey/ToolkitKyra/main/SAPGUIcontrol.version.json';
const GUIDE_URL='https://synalocey.github.io/ToolkitKyra/guide';
const APP_CSS=chosenBaseCss+(`
:root{--bg:#f3f6f9;--panel:#fff;--ink:#1e293b;--muted:#64748b;--line:#dce5ed;--navy:#10384b;--blue:#2563eb;--teal:#0f766e;--cyan:#0891b2;--green:#16a34a;--amber:#d97706;--red:#dc2626;--purple:#7c3aed;
--shadow:0 2px 6px rgba(15,23,42,.025),0 10px 28px rgba(15,23,42,.035);--radius:14px;--side:350px;
--font-sans:"Segoe UI Variable Text","Segoe UI","Microsoft YaHei UI","Microsoft YaHei",Arial,sans-serif}*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--bg);color:var(--ink);` +
  `font-family:var(--font-sans);font-size:13px;font-weight:400;line-height:1.45}
h1,h2,h3,b,strong{font-weight:600}button,input,select,textarea{font:inherit}input::placeholder,textarea::placeholder{color:#8493a5}
button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible,summary:focus-visible{outline:3px solid rgba(37,99,235,.22);outline-offset:2px}.hero{background:linear-gradient(115deg,#10384b 0,#15556c ` +
  `58%,#176f6b 100%);color:#fff;padding:24px 30px;box-shadow:0 3px 12px rgba(16,56,75,.1)}
.hero-inner{max-width:1900px;margin:auto;display:flex;justify-content:space-between;align-items:center;gap:24px}.hero h1{display:flex;align-items:center;gap:10px;margin:0 0 7px;font-size:28px;line-height:1.25;` +
  `letter-spacing:-.6px}
.hero h1>span{display:inline-flex;align-items:center;justify-content:center;width:37px;height:37px;border:1px solid #ffffff24;border-radius:10px;background:#ffffff12;font-size:25px;letter-spacing:0}
.hero p{margin:0;color:#d0e4e8;font-size:12px;line-height:1.6;max-width:1050px}.tag{display:inline-flex;align-items:center;border:1px solid #ffffff40;background:#ffffff14;border-radius:999px;padding:4px 10px;margin:0 6px 4px ` +
  `0;font-size:11px;font-weight:600;letter-spacing:.2px}
.hero-source{min-width:270px;background:#ffffff0d;border:1px solid #ffffff26;border-radius:11px;padding:12px 15px;text-align:right}.hero-source small{display:block;color:#c1dce1;font-size:10px;text-transform:uppercase;` +
  `letter-spacing:.8px;font-weight:600}
.hero-source strong{display:block;font-size:13px;margin-top:5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:390px}
.layout{display:grid;grid-template-columns:var(--side) minmax(0,1fr);gap:16px;max-width:1900px;margin:auto;padding:16px}.card,.side{background:var(--panel);border:1px solid var(--line);border-radius:
var(--radius);box-shadow:var(--shadow)}.side{align-self:start;padding:16px}.main{display:grid;align-content:start;gap:16px;
min-width:0}.section-title{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:14px}.section-title h2,.side h2{font-size:16px;margin:0;letter-spacing:-.2px}
.section-title p{margin:4px 0 0;color:var(--muted);font-size:11px;line-height:1.5}.side-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin:18px 0 8px;padding-top:16px;border-top:1px solid ` +
  `#e8eef3}.upload-box{border:1px solid #dce7ed;background:linear-gradient(180deg,#f6fafc,#fbfdfe);border-radius:11px;padding:12px;margin-top:12px}.upload-box label{display:block;font-size:11px;font-weight:600;color:#334155;` +
  `text-transform:
uppercase;letter-spacing:.45px;margin-bottom:7px}.upload-manual{margin-top:12px;padding-top:12px;border-top:1px solid #dbe5ee}.upload-manual summary{color:#334155;font-size:11px;font-weight:600;text-transform:uppercase;` +
  `letter-spacing:.45px;cursor:pointer}.upload-manual-content{padding-top:10px}.upload-box input{display:block;width:100%;font-size:11px;color:#475569}.upload-box input::file-selector-button{border:1px solid #d4e0e8;` +
  `border-radius:7px;background:#fff;color:#334155;font:inherit;font-weight:600;padding:7px 9px;margin-right:7px;cursor:pointer}.source-status{margin-top:10px;border-left:3px solid var(--blue);background:#eff6ff;color:#1e3a8a;` +
  `border-radius:8px;padding:9px 10px;font-size:11px;line-height:1.5}.source-status.loading{border-color:var(--amber);background:#fffbeb;color:#78350f}.source-status.error{border-color:var(--red);background:#fef2f2;
color:#991b1b}.source-status.ok{border-color:var(--teal);background:#ecfdf5;color:#14532d}.btn{border:1px solid transparent;border-radius:8px;background:var(--teal);color:#fff;font-size:13px;font-weight:600;line-height:1.35;` +
  `padding:9px 13px;cursor:pointer;box-shadow:0 1px 2px rgba(15,23,42,.06);
transition:background .15s ease,border-color .15s ease,box-shadow .15s ease,transform .15s ease}.btn:hover:not(:disabled){background:#0d665f;box-shadow:0 3px 7px rgba(15,23,42,.1);transform:translateY(-1px)}
.btn.light{border-color:#d5e0e8;background:#fff;color:#334155;box-shadow:0 1px 2px rgba(15,23,42,.025)}.btn.light:hover:not(:disabled){border-color:#b9cbd8;background:#f1f6f9}.btn.blue{background:var(--blue)}` +
  `.btn.blue:hover:not(:disabled){background:#1d4ed8}
.btn.small{padding:5px 9px;font-size:11px}.btn.wide{width:100%;margin-top:0}
.filter-group{border-top:1px solid #edf2f6;padding:11px 0 3px}.filter-label{display:flex;align-items:center;justify-content
:space-between;gap:8px;margin-bottom:5px}.filter-label label{font-size:10px;color:#64748b;text-transform:uppercase;letter-spacing:.45px;font-weight:600;cursor:pointer}.filter-label button{border:0;
background:0 0;color:var(--blue);padding:2px;font-size:10px;font-weight:600;cursor:pointer}.filter-select{width:100%;min-height:38px;border:1px solid var(--line);border-radius:9px;background:#fff;
padding:6px}.quick-pick{display:flex;align-items:center;gap:5px;flex-wrap:wrap;margin-top:6px}.quick-pick span{font-size:9px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:
.35px}.quick-pick button{border:1px solid #dce5ed;background:#f7f9fb;color:#475569;border-radius:6px;padding:4px 7px;font-size:10px;font-weight:600;cursor:pointer}.quick-pick button:hover{border-color:#8ebdbb;color:#0f766e}` +
  `.quick-pick button.active{border-color:#83bcb6;background:#e6f4f1;color:#115e59}.quick-pick button.partial{border-color:#a7cbe0;background:#edf6fc;color:#075985}
.quick-search{width:100%;border:1px solid var(--line);border-radius:8px;background:#fff;padding:9px 10px;font-size:12px;color:var(--ink)}.chosen-container{font-family:var(--font-sans)!important;font-size:12px!important;width:
100%!important}.chosen-container-multi .chosen-choices{min-height:38px!important;border:1px solid var(--line)!important;border-radius:8px!important;background:#fff!important;background-image:
none!important;box-shadow:none!important;padding:3px 5px!important}.chosen-container-active .chosen-choices{border-color:#38bdf8!important;box-shadow:0 0 0 3px rgba(56,189,248,.13)!important}
.chosen-container-multi .chosen-choices li.search-choice{border:0!important;background:#eaf4f7!important;color:#195b70!important;border-radius:6px!important;box-shadow:none!important;padding:5px 22px
5px 7px!important;font-weight:600!important}.chosen-container-multi .chosen-choices li.search-choice .search-choice-close{top:50%!important;right:5px!important;width:14px!important;height:
14px!important;margin-top:-7px!important;background:0 0!important;font-size:0!important}.chosen-container-multi .chosen-choices li.search-choice .search-choice-close::after{content:"×";display:block;
color:#0369a1;font-size:15px;line-height:13px;text-align:center;font-weight:400;pointer-events:none}.chosen-container-multi .chosen-choices li.search-choice .search-choice-close:hover::after{color:
var(--red)}.chosen-container .chosen-drop{border:1px solid #b8c8d8!important;border-radius:0 0 11px 11px!important;box-shadow:0 12px 26px rgba(15,23,42,.16)!important}
.chosen-container .chosen-results li.highlighted{background:#0f766e!important;background-image:none!important}.chosen-container .chosen-results{max-height:250px!important}
.chosen-container-multi .chosen-choices li.search-field input{height:28px!important;color:#64748b!important;font-family:inherit!important;font-size:12px!important;line-height:28px!important}.active-filters{display:flex;` +
  `flex-wrap:wrap;gap:6px
;margin-top:9px}.filter-chip{display:inline-flex;align-items:center;gap:6px;border:1px solid #bae6fd;background:#f0f9ff;color:#075985;border-radius:999px;padding:4px 7px 4px 9px;font-size:10px;
font-weight:600;max-width:100%}.filter-chip span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.filter-chip button{border:0;background:0 0;color:#075985;cursor:pointer;padding:0;font-size:
14px;line-height:1}.all-chip{font-size:11px;color:var(--muted);padding:3px 0}.hint{border-left:3px solid #81b7b1;background:#f0f7f7;color:#486570;border-radius:8px;padding:9px 10px;font-size:11px;line-height:1.5;` +
  `margin-top:12px}.kpis{display:grid;grid-template-columns:repeat(5,minmax(125px,1fr));gap:12px}.kpis:empty{display:none}.kpi{background:#fff;border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow);` +
  `padding:15px 16px;min-height:105px;position:relative;overflow:hidden}.kpi::after{content:"";position:absolute;height:3px;left:16px;right:16px;top:0;border-radius:0 0 3px 3px;background:var(--kpi-color,#2563eb);opacity:.65}` +
  `.kpi small{display:block;color:var(--muted);font-size:10px;text-transform:uppercase;letter-spacing:.5px;font-weight:600}.kpi strong,.kpi-num{display:block;font-size:28px;font-weight:700;font-variant-numeric:tabular-nums;` +
  `letter-spacing:-.7px;line-height:1.1;margin:9px 0 5px;color:var(--kpi-color,#2563eb)}.kpi-num{border:0;background:0 0;padding:0;font-weight:700;cursor:pointer;text-align:left}.kpi-num:hover{text-decoration:
underline;text-underline-offset:3px}.kpi span{font-size:11px;color:#64748b}.card{padding:18px;min-width:0}.context-badge{display:inline-flex;align-items:center;gap:7px;background:#edf8f3;color:#22604b;border:1px solid ` +
  `#cde7d9;border-radius:999px;padding:5px 10px;font-size:11px;font-weight:600;white-space:nowrap}.context-badge::before{content:"";width:6px;height:6px;border-radius:50%;background:#25a568;box-shadow:0 0 0 3px ` +
  `rgba(37,165,104,.09)}.pivot-grid{display:grid;grid-template-columns:repeat(3,minmax(260px,1fr));gap:12px}.pivot-card{border:1px solid #e0e8ef;border-radius:10px;background:#fff;min-width:0;overflow:hidden}` +
  `.pivot-head{display:flex;justify-content:space-between;align-items:center;gap:9px;padding:11px 12px;border-bottom:1px solid #e7edf2;background:#f7fafc}
.pivot-head h3{font-size:12px;margin:0;color:#334155}.pivot-head span{font-size:10px;color:var(--muted);white-space:nowrap}.pivot-list{height:310px;min-height:120px;padding:7px 9px 9px;overflow:auto;resize:vertical;
scrollbar-width:thin}.pivot-row{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(55px,1fr) 42px;gap:8px;align-items:center;padding:6px 3px;border-radius:6px}.pivot-row:hover{background:
#f1f5f9}.pivot-label{font-size:11px;color:#334155;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.pivot-bar{height:6px;background:#edf2f6;border-radius:999px;overflow:hidden}.pivot-bar i{
display:block;height:100%;min-width:3px;border-radius:999px;background:var(--bar,#0f766e)}.pivot-count{border:0;background:#edf5fa;color:#195b70;border-radius:6px;padding:4px 6px;font-size:11px;font-weight:700;cursor:pointer;` +
  `text-align:center;font-variant-numeric:tabular-nums}.pivot-count:hover{background:#0f766e;color:#fff}.empty{padding:28px 14px;text-align:center;color:var(--muted);
font-size:12px}.table-toolbar{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;margin-bottom:11px}.toolbar-left,.toolbar-right{display:flex;align-items:center;gap:
8px;flex-wrap:wrap}.table-search{width:290px;max-width:60vw;border:1px solid var(--line);border-radius:8px;padding:9px 10px;font-size:12px}.row-count{font-size:11px;color:#526579;background:#f3f6f9;border:1px solid #e5ebf1;` +
  `border-radius:6px;padding:4px 7px;font-variant-numeric:tabular-nums}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:10px;max-height:660px;background:#fff;scrollbar-width:thin}` +
  `table{border-collapse:separate;
border-spacing:0;table-layout:fixed;width:1658px;min-width:1658px;font-size:11px}td,th{border-right:1px solid #edf1f5;border-bottom:1px solid #e7edf4;padding:7px 8px;text-align:left;white-space:nowrap
;overflow:hidden;text-overflow:ellipsis}th{position:sticky;top:0;z-index:4;background:#f2f6f9;color:#475569;font-size:10px;text-transform:uppercase;letter-spacing:.3px;font-weight:600;cursor:pointer;
user-select:none}td:first-child,th:first-child{position:sticky;left:0;box-shadow:2px 0 0 #dbe5ee}th:first-child{z-index:6}td:first-child{z-index:2;background:#fff}th:hover{background:#e8f0f7}
th .sort-mark{margin-left:5px;color:#94a3b8}th.sorted .sort-mark{color:var(--blue)}tbody tr:nth-child(2n) td{background:#fbfdff}tbody tr:hover td{background:#f0fdfa}.status-pill{display:inline-flex;
align-items:center;border-radius:999px;padding:3px 7px;font-size:11px;font-weight:400;line-height:1.2;background:#dbeafe;color:#1d4ed8}.status-pill.terminal{background:#dcfce7;color:#166534}
.status-pill.stop{background:#fee2e2;color:#991b1b}.status-pill.unknown{background:#fef3c7;color:#92400e}.footer{max-width:1900px;margin:auto;padding:0 18px 20px;color:#64748b;font-size:10px;
line-height:1.5}.toast{position:fixed;right:20px;bottom:20px;z-index:50;max-width:390px;background:#0f172a;color:#fff;border-radius:12px;padding:11px 14px;box-shadow:0 18px 36px rgba(15,23,42,.26);
font-size:12px;line-height:1.45;opacity:0;transform:translateY(12px);pointer-events:none;transition:.2s ease}.toast.show{opacity:1;transform:translateY(0)}.toast.error{background:#991b1b}.loading-line
{height:3px;position:fixed;left:0;top:0;z-index:99;background:#2dd4bf;width:0;transition:width .25s ease}.loading-line.on{width:72%;animation:loading 1.3s ease-in-out infinite}.loading-line.done{width
:100%}@keyframes loading{0%{opacity:.55}50%{opacity:1}100%{opacity:.55}}@media(max-width:1450px){.pivot-grid{grid-template-columns:repeat(2,minmax(260px,1fr))}.kpis{grid-template-columns:repeat(3,1fr)
}}@media(max-width:980px){body{min-width:1030px}.hero-inner{display:block}.hero-source{margin-top:12px;text-align:left}.layout{grid-template-columns:320px minmax(680px,1fr);min-width:1030px}` +
  `.pivot-grid{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,1fr)}}@media(max-width:560px){.hero{padding:20px 16px}
.hero h1{font-size:24px}.layout{padding:10px}.kpis{grid-template-columns:1fr 1fr;gap:8px}.kpi{min-height:90px;padding:11px}.kpi strong,.kpi-num{font-size:23px}.table-search{width:100%;max-width:none}}
@media print{.footer,.pivot-grid,.side,.table-toolbar{display:none}.layout{display:block}.hero{background:#075985!important;-webkit-print-color-adjust:exact}.table-wrap{max-height:none;overflow:
visible}table{width:100%;font-size:7px}td,th{position:static!important;white-space:normal}}.tabs{position:sticky;top:0;z-index:40;display:flex;align-items:center;justify-content:space-between;gap:12px;background:#fff;` +
  `border-bottom:1px solid var(--line);box-shadow:0 2px 5px rgba(15,23,42,.025);padding:9px max(16px,calc((100vw - 1900px)/ 2 + 16px))}
.tab-list{display:flex;gap:3px;padding:3px;border:1px solid #e1e8ef;border-radius:11px;background:#f0f4f7}.tab-btn{border:1px solid transparent;background:transparent;color:#64748b;border-radius:8px;padding:7px 14px;` +
  `font-size:12px;font-weight:600;cursor:pointer;transition:background .15s ease,color .15s ease,box-shadow .15s ease}
.tab-btn:hover{background:#e7eef3;color:#195b70}.tab-btn.active{border-color:#d9e5e7;background:#fff;color:#0f766e;box-shadow:0 1px 3px rgba(15,23,42,.07)}
.tab-page{display:none}.tab-page.active{display:block}.cycle-shell{display:grid;gap:16px;max-width:1900px;margin:auto;padding:16px}
.cycle-label{display:block;color:#334155;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.45px;margin-bottom:6px}.cycle-input{display:block;width:100%;resize:vertical;min-height
:120px;border:1px solid var(--line);border-radius:12px;background:#fff;color:var(--ink);padding:11px 12px;font:12px/1.55 Consolas,"Segoe UI",sans-serif}.cycle-input:focus{outline:3px solid 
rgba(14,165,233,.2);border-color:#38bdf8}.cycle-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:10px}.cycle-progress{display:
block;width:100%;height:6px;margin-top:14px;accent-color:var(--teal);appearance:none;border:0;border-radius:999px;overflow:hidden;background:#e9eff3}.cycle-progress::-webkit-progress-bar{background:#e9eff3;` +
  `border-radius:999px}
.cycle-progress::-webkit-progress-value{background:var(--teal);border-radius:999px}.cycle-progress::-moz-progress-bar{background:var(--teal);border-radius:999px}.cycle-status{margin-top:7px;color:var(--muted);font-size:11px;` +
  `line-height:1.45}.cycle-status:empty{display:none}.cycle-wrap{max-height:690px}#cycleTable{width:
2422px;min-width:2422px}.ct-link{border:0;background:0 0;color:#0369a1;padding:0;font:inherit;font-weight:inherit;text-decoration:underline;text-underline-offset:2px;cursor:pointer}.ct-link:hover{
color:#0f766e}.ct-link:focus-visible{outline:2px solid #38bdf8;outline-offset:3px;border-radius:2px}.duration-over,.duration-over .ct-link{color:var(--red)!important}.wf-modal[hidden]{display:none}
.wf-modal{position:fixed;inset:0;z-index:1000;display:grid;place-items:center;padding:20px}.wf-bg{position:absolute;inset:0;width:100%;height:100%;border:0;background:rgba(15,23,42,.48);backdrop-filter:blur(2px);` +
  `cursor:default
}.wf-box{position:relative;display:flex;flex-direction:column;width:min(1180px,calc(100vw - 40px));max-height:calc(100vh - 40px);overflow:hidden;border:1px solid #cbd5e1;border-radius:16px;background:
#fff;box-shadow:0 24px 80px rgba(15,23,42,.35)}.wf-head{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:15px 17px;border-bottom:1px solid var(--line);background:#f8fafc}
.wf-head h2{margin:0;color:#0f172a;font-size:17px}.wf-head p{margin:3px 0 0;color:var(--muted);font-size:11px}.wf-actions{display:flex;align-items:center;gap:8px}.wf-wrap{overflow:auto;max-height:
calc(100vh - 120px)}#wfTable{width:1120px;min-width:1120px;table-layout:fixed}#wfTable th:first-child{width:190px}#wfTable th:nth-child(2){width:120px}#wfTable th:nth-child(3){width:160px}
#wfTable th:nth-child(4){width:95px}#wfTable th:nth-child(5){width:120px}#wfTable th:nth-child(6){width:435px}#wfTable td{white-space:normal;overflow-wrap:anywhere;vertical-align:top;line-height:1.45}
#wfTable td:nth-child(6){white-space:pre-wrap}body.wf-open,body.sap-range-open{overflow:hidden}.toast{z-index:1100}#wfTable th{white-space:normal;overflow:visible;text-overflow:clip;line-height:1.25;cursor:default}
.wf-box{width:min(1320px,calc(100vw - 40px))}#wfTable{width:1280px;min-width:1280px}#wfTable th:nth-child(4){width:160px}#wfTable th:nth-child(5){width:95px}#wfTable th:nth-child(6){width:120px}
#wfTable th:nth-child(7){width:435px}#wfTable td:nth-child(6){white-space:normal}#wfTable td:nth-child(7){white-space:pre-wrap}.wf-box.sap-range-box{width:min(460px,calc(100vw - 40px))}
.sap-range-form{display:grid;gap:14px;padding:17px}.sap-range-fields{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.sap-range-field{display:grid;gap:6px}.sap-range-field label{color:#475569;` +
  `font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.4px}.sap-range-field input{width:100%;border:1px solid var(--line);border-radius:8px;background:#fff;color:var(--ink);padding:10px;font-size:12px}` +
  `.sap-range-field input:focus{outline:3px solid rgba(14,165,233,.2);border-color:#38bdf8}.sap-range-note,.sap-range-default{color:var(--muted);font-size:10px;line-height:1.45}.sap-range-default{margin-top:9px}` +
  `.sap-range-error{min-height:16px;color:#b91c1c;font-size:11px}.sap-range-form .wf-actions{justify-content:flex-end}.btn:disabled{cursor:not-allowed;box-shadow:none;transform:none;opacity:.48}
@media(max-width:980px){.tabs{min-width:1030px}.cycle-shell{min-width:1030px}}@media(prefers-reduced-motion:reduce){html{scroll-behavior:auto}*,*::before,*::after{animation:none!important;transition:none!important}}
`);
const APP_HTML=`<div id="loadingLine" class="loading-line"></div>
<header class="hero">
  <div class="hero-inner">
    <div>
      <h1><span aria-hidden="true">📊</span> NSR Tracking Dashboard</h1>
      <p>One-page NSR workflow view. Every pivot uses the same filtered NSR population; click any count to drill into that value and update the filter panel.</p>
    </div>
    <div class="hero-source"><small>Current data source</small><strong id="heroSource">Checking local cache…</strong></div>
  </div>
</header>
<nav class="tabs" aria-label="NSR Tracking Dashboard pages">
  <div class="tab-list" role="tablist">
    <button id="overviewBtn" class="tab-btn active" type="button" role="tab" aria-selected="true" data-tab="overview">NSR Overview</button>
    <button id="cycleBtn" class="tab-btn" type="button" role="tab" aria-selected="false" data-tab="cycle">NSR Cycle Time</button>
  </div>
</nav>

<section id="overviewTab" class="tab-page active" role="tabpanel" aria-labelledby="overviewBtn"><div class="layout">
  <aside class="side">
    <h2>Data Source</h2>
    <div class="upload-box">
      <button id="sapRangeOpen" class="btn blue wide" type="button">Load / Refresh NSR</button>
      <div id="sapRangeDefault" class="sap-range-default"></div>
      <details class="upload-manual">
        <summary>Upload raw data</summary>
        <div class="upload-manual-content">
          <label for="rawFile">Choose CSV or Excel file</label>
          <input id="rawFile" type="file" accept=".xlsx,.xls,.csv">
        </div>
      </details>
      <div id="sourceStatus" class="source-status loading">Checking local NSR cache…</div>
    </div>

    <div class="side-head"><h2>Filters</h2><button id="resetBtn" class="btn light small" type="button">Reset all</button></div>
    <div class="filter-group">
      <div class="filter-label"><label for="globalSearch">NSR / title quick search</label><button type="button" data-clear-search>Clear</button></div>
      <input id="globalSearch" class="quick-search" placeholder="NSR#, FCID, title…" title="Paste multiple FIDs / NSRs from Excel, or separate with whitespace, commas, semicolons, / or |. Matches any ID." autocomplete="off">
    </div>
    <div id="filterHost"></div>
    <div id="activeFilters" class="active-filters"><span class="all-chip">All NSRs are selected.</span></div>
  </aside>

  <main class="main">
    <section id="kpis" class="kpis" aria-label="NSR summary"></section>

    <section class="card">
      <div class="section-title">
        <div><h2>Interactive Pivot</h2></div>
        <span id="pivotContext" class="context-badge">0 filtered NSRs</span>
      </div>
      <div id="pivotGrid" class="pivot-grid"></div>
    </section>

    <section class="card" id="detailSection">
      <div class="section-title">
        <div><h2>NSR Detail</h2></div>
      </div>
      <div class="table-toolbar">
        <div class="toolbar-left">
          <input id="tableSearch" class="table-search" placeholder="Search within current result…" autocomplete="off">
          <span id="rowCount" class="row-count">0 rows</span>
        </div>
        <div class="toolbar-right">
          <button id="copyBtn" class="btn light" type="button">Copy</button>
          <button id="sendCycleBtn" class="btn blue" type="button">Cycle Time</button>
        </div>
      </div>
      <div class="table-wrap">
        <table id="detailTable" aria-label="Filtered NSR detail"></table>
      </div>
    </section>
  </main>
</div>
</section>
<section id="cycleTab" class="tab-page" role="tabpanel" aria-labelledby="cycleBtn">
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
      <div class="section-title">
        <div><h2>Cycle Time Detail</h2></div>
      </div>
      <div class="table-toolbar">
        <div class="toolbar-left">
          <input id="cycleSearch" class="table-search" placeholder="Search within lookup result…" autocomplete="off">
          <span id="cycleCount" class="row-count">0 rows</span>
        </div>
        <div class="toolbar-right">
          <button id="copyCycleBtn" class="btn light" type="button" disabled>Copy</button>
          <button id="exportCycleBtn" class="btn light" type="button" title="Export the current Cycle Time rows and their workflow details" disabled>Export XLSX</button>
        </div>
      </div>
      <div class="table-wrap cycle-wrap">
        <table id="cycleTable" aria-label="NSR cycle time detail"></table>
      </div>
    </section>
  </div>
</section>
<div id="sapRangeModal" class="wf-modal" hidden>
  <button class="wf-bg" type="button" data-sap-range-close aria-label="Close SAP date range"></button>
  <section class="wf-box sap-range-box" role="dialog" aria-modal="true" aria-label="Load or refresh NSR date range">
    <form id="sapRangeForm" class="sap-range-form">
      <div class="sap-range-fields">
        <div class="sap-range-field"><label for="sapRangeFrom">From date</label><input id="sapRangeFrom" type="date" required></div>
        <div class="sap-range-field"><label for="sapRangeTo">To date</label><input id="sapRangeTo" type="date" required></div>
      </div>
      <div class="sap-range-note">Larger ranges may take 1-2 minutes.</div>
      <div id="sapRangeError" class="sap-range-error" role="alert" aria-live="polite"></div>
      <div class="wf-actions">
        <button class="btn light" type="button" data-sap-range-close>Cancel</button>
        <button id="sapRangeLoad" class="btn blue" type="submit">Load</button>
      </div>
    </form>
  </section>
</div>
<div id="wfModal" class="wf-modal" hidden>
  <button class="wf-bg" type="button" data-wf-close aria-label="Close workflow detail"></button>
  <section class="wf-box" role="dialog" aria-modal="true" aria-labelledby="wfTitle">
    <div class="wf-head">
      <div><h2 id="wfTitle">Workflow Detail</h2><p id="wfMeta"></p></div>
      <div class="wf-actions">
        <button id="wfCopy" class="btn light" type="button">Copy</button>
        <button id="wfClose" class="btn light" type="button" data-wf-close aria-label="Close workflow detail">Close</button>
      </div>
    </div>
    <div class="wf-wrap"><table id="wfTable" aria-label="Full NSR workflow detail"></table></div>
  </section>
</div>
<div id="toast" class="toast" role="status" aria-live="polite"></div>`;
const BRIDGE_CSS=(`
:root{--bg:#f3f7fb;--panel:#fff;--ink:#0f172a;--muted:#64748b;--line:#d8e3ee;--blue:#2563eb;--teal:#0f766e;--green:#16a34a;--amber:#d97706;--red:#dc2626;--shadow:0 14px 34px rgba(15,23,42,.08);
--font-sans:"Segoe UI Variable Text","Segoe UI","Microsoft YaHei UI","Microsoft YaHei",Arial,sans-serif}*{box-sizing:border-box}[hidden]{display:none!important}body{margin:0;background:var(--bg);color
:var(--ink);font-family:var(--font-sans);font-weight:400;line-height:1.4}b,h1,h2,strong{font-weight:600}button,input,select,textarea{font:inherit}button:focus-visible,input:focus-visible,
select:focus-visible,textarea:focus-visible{outline:3px solid rgba(14,165,233,.3);outline-offset:2px}.bridge-hero{background:linear-gradient(128deg,#052e44,#075985 52%,#0f766e);color:#fff;padding:24px
 28px;box-shadow:0 9px 24px rgba(3,46,70,.18)}.bridge-hero-inner{max-width:1180px;margin:auto;display:flex;align-items:flex-end;justify-content:space-between;gap:20px}.bridge-hero h1{margin:0 0 6px;
font-size:29px;line-height:1.15;letter-spacing:-.3px}.bridge-hero p{max-width:760px;margin:0;color:#d8fbf5;font-size:13px;line-height:1.5}.toolbox-shell{max-width:1180px;margin:auto;padding:18px}
.bridge-shell,.zbom-grid,.create-quote-grid{display:grid;gap:16px}.bridge-shell{grid-template-columns:repeat(2,minmax(0,1fr))}.zbom-grid{grid-template-columns:minmax(0,1.3fr) minmax(300px,.9fr)}` +
  `.create-quote-grid{grid-template-columns:240px minmax(240px,1fr) 320px}.bridge-card{min-width:0;
background:var(--panel);border:1px solid var(--line);border-radius:16px;box-shadow:var(--shadow);padding:17px}.bridge-card.wide{grid-column:1/-1}.bridge-card h2{margin:0;font-size:17px;line-height:1.3
}.bridge-card p{margin:4px 0 0;color:var(--muted);font-size:11px;line-height:1.5}.bridge-card-head{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;margin-bottom:14px}
.bridge-card-tools{display:flex;align-items:center;gap:7px}.connection-card{margin-bottom:15px}.bridge-btn{border:0;border-radius:10px;background:var(--teal);color:#fff;font-weight:600;padding:9px 
12px;cursor:pointer;transition:filter .12s ease,transform .12s ease}.bridge-btn:hover:not(:disabled){filter:brightness(1.05);transform:translateY(-1px)}.bridge-btn.blue{background:var(--blue)}
.bridge-btn.light{background:#e2e8f0;color:#0f172a}.bridge-btn.ghost{border:1px solid #ffffff66;background:#ffffff18}.bridge-btn:disabled{cursor:not-allowed;opacity:.5}.bridge-btn.small{padding:6px 
9px;font-size:10px}.bridge-badge{display:inline-flex;align-items:center;gap:7px;border-radius:999px;background:#f1f5f9;color:#475569;padding:6px 10px;font-size:11px;font-weight:600;white-space:nowrap}
.bridge-badge::before{content:"";width:8px;height:8px;border-radius:50%;background:#94a3b8}.bridge-badge.ok{background:#dcfce7;color:#166534}.bridge-badge.ok::before{background:var(--green)}
.bridge-badge.warn{background:#fef3c7;color:#92400e}.bridge-badge.warn::before{background:var(--amber)}.bridge-badge.error{background:#fee2e2;color:#991b1b}.bridge-badge.error::before{background:
var(--red)}.bridge-status{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}.bridge-status-item{border:1px solid #e2e8f0;border-radius:12px;background:#f8fafc;padding:11px}
.bridge-status-item small{display:block;color:var(--muted);font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.35px}.bridge-status-item strong{display:block;margin-top:5px;font-size
:12px;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.bridge-download{display:inline-flex;align-items:center;margin-top:9px;text-decoration:none}.bobj-status-head{display:flex;align-items:center;` +
  `justify-content:space-between;gap:8px}.bobj-status-actions{display:flex;align-items:center;gap:5px}.bobj-update{flex:none;padding:4px 8px}.bobj-data-source{font-variant-numeric:tabular-nums}.bobj-refresh-state{display:block;` +
  `margin-top:4px;white-space:normal}.bobj-refresh-state.running{color:#92400e}.bobj-refresh-state.error{color:#b91c1c}.bridge-field{display:grid;gap:6px;margin-top:11px}.bridge-field label{color:#475569;font-size:10px;` +
  `font-weight:600;
text-transform:uppercase;letter-spacing:.35px}.bridge-field input,.bridge-field select,.bridge-field textarea{display:block;width:100%;border:1px solid #cbd5e1;border-radius:10px;background:#fff;color
:var(--ink);padding:10px 11px;font-size:12px}.bridge-field textarea{min-height:88px;resize:vertical;font:11px/1.55 Consolas,"Courier New",monospace;tab-size:4}.bridge-field textarea.quotation-input{
white-space:pre;overflow:auto;overflow-wrap:normal;word-break:normal}.bridge-field input:disabled,.bridge-field select:disabled,.bridge-field textarea:disabled{background:#f1f5f9;color:#94a3b8}
.bridge-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:12px}.bridge-inline{display:grid;grid-template-columns:1fr 110px;gap:10px}.bridge-divider{height:1px;background:
#e2e8f0;margin:15px 0}.bridge-note{color:var(--muted);font-size:10px;line-height:1.45}.bridge-result{min-height:130px;max-height:320px;overflow:auto;margin:0;border:1px solid #dbe5ee;border-radius:
12px;background:#071521;color:#d8fbf5;padding:12px;white-space:pre-wrap;overflow-wrap:anywhere;font:11px/1.55 Consolas,"Courier New",monospace}.tool-tabs{display:flex;gap:6px;margin:0 0 15px;padding:
5px;border:1px solid var(--line);border-radius:13px;background:#e8f0f7;box-shadow:0 8px 20px rgba(15,23,42,.05)}.tool-tab{flex:0 0 auto;border:0;border-radius:9px;background:0 0;color:#475569;padding:
9px 18px;font-weight:600;cursor:pointer}.tool-tab[aria-selected=true]{background:#fff;color:#0f4c67;box-shadow:0 3px 10px rgba(15,23,42,.1)}.token-list{display:flex;flex-wrap:wrap;gap:6px;margin-top:
8px}.token-chip{border:1px solid #bfdbfe;border-radius:999px;background:#eff6ff;color:#1d4ed8;padding:4px 8px;font-size:10px;cursor:pointer}.folder-line{display:flex;align-items:center;gap:9px;
flex-wrap:wrap;margin-top:12px}.folder-name{min-width:0;color:#475569;font-size:11px;overflow-wrap:anywhere}.history-button-count{font-variant-numeric:tabular-nums}.toolbox-modal{position:fixed;inset:
0;z-index:1000;display:grid;place-items:center;padding:24px}.toolbox-modal-backdrop{position:absolute;inset:0;background:rgba(2,16,27,.56);backdrop-filter:blur(2px)}.toolbox-modal-card{position:
relative;display:grid;grid-template-rows:auto minmax(0,1fr);width:min(920px,100%);max-height:calc(100vh - 48px);overflow:hidden;border:1px solid var(--line);border-radius:16px;background:#fff;
box-shadow:0 24px 70px rgba(2,16,27,.3);padding:17px}.toolbox-modal-head{display:flex;align-items:center;justify-content:space-between;gap:14px;padding-bottom:12px;border-bottom:1px solid #e2e8f0}
.toolbox-modal-title{display:flex;align-items:baseline;gap:10px}.toolbox-modal-title h2{margin:0;font-size:17px}.history-count{color:#64748b;font-size:10px}.history-list{display:grid;gap:8px;
overflow-y:auto;overflow-x:hidden;padding:12px 2px 2px}.history-empty{color:#94a3b8;font-size:11px;padding:14px 5px}.history-entry{display:grid;grid-template-columns:165px minmax(0,1fr);gap:12px;width
:100%;min-width:0;border:1px solid #dbe5ee;border-radius:10px;background:#fff;color:#334155;padding:10px 11px;text-align:left;cursor:pointer}.history-entry:hover{border-color:#93c5fd;background:
#eff6ff}.history-entry-time{display:block;color:#64748b;font-size:10px;line-height:1.4;font-variant-numeric:tabular-nums}.history-entry-text{display:block;min-width:0;white-space:pre-wrap;
overflow-wrap:anywhere;word-break:break-word;font:10px/1.45 Consolas,"Courier New",monospace}body.history-open{overflow:hidden}.preview-toolbar{display:flex;align-items:center;justify-content:
space-between;gap:10px;flex-wrap:wrap}.preview-count{color:#475569;font-size:11px}.zbom-table-wrap{margin-top:12px;overflow:auto;max-height:360px;border:1px solid #dbe5ee;border-radius:12px}
.zbom-table{width:100%;border-collapse:separate;border-spacing:0;font-size:11px}.zbom-table th{position:sticky;top:0;z-index:1;background:#edf5fa;color:#334155;text-align:left;font-size:10px;
white-space:nowrap}.zbom-table td,.zbom-table th{padding:8px 9px;border-right:1px solid #e5edf4;border-bottom:1px solid #e5edf4;vertical-align:top}.zbom-table td:last-child,.zbom-table th:last-child{
border-right:0}.zbom-table tr:last-child td{border-bottom:0}.zbom-table .missing{color:#94a3b8}.row-ok{color:#166534}.row-warning{color:#92400e}.row-error{color:#b91c1c}.planned-name{min-width:270px;max-width:440px;
overflow-wrap:anywhere}.compare-template-control{display:grid;grid-template-columns:minmax(0,1fr) 64px;gap:8px;align-items:center}.compare-template-separator{text-align:center;font-weight:700}.compare-table td,.compare-table ` +
  `th{padding:4px 7px;vertical-align:middle}
.compare-fid{min-width:78px}.compare-customer{width:110px;min-width:100px;max-width:145px;overflow-wrap:anywhere}.compare-system-description{width:155px;min-width:140px;max-width:230px;overflow-wrap:anywhere}
.compare-group-cell{width:92px;min-width:92px}.compare-group-select{display:block;width:100%;min-width:0;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:var(--ink);padding:3px 6px;font-size:11px}` +
  `.compare-file-name{min-width:460px;max-width:900px;
overflow-wrap:anywhere}.compare-file-link{border:0;background:0 0;color:#0369a1;padding:0;text-align:left;font:inherit;text-decoration:underline;text-underline-offset:2px;cursor:pointer;overflow-wrap:anywhere}` +
  `.compare-file-link:hover{color:#0f766e}.compare-file-link:disabled{color:#94a3b8;cursor:wait}.compare-output-control{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;align-items:center}` +
  `.compare-keep-temp{white-space:nowrap}.compare-keep-temp[aria-pressed=true]{background:var(--teal);color:#fff}.create-quote-reference-title{white-space:nowrap}.create-quote-number{max-width:112px;` +
  `font-variant-numeric:tabular-nums}.create-quote-setting-row{grid-template-columns:max-content minmax(0,1fr);align-items:center;column-gap:10px}.create-quote-days-control{display:grid;grid-template-columns:auto 62px auto;` +
  `align-items:center;justify-content:start;gap:7px}.create-quote-days-control input{text-align:center;font-variant-numeric:tabular-nums}.create-quote-unit{color:#475569;font-size:11px;font-weight:600}` +
  `.create-quote-suffix{max-width:76px}.create-quote-crd-input{display:block;width:150px;max-width:100%;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:var(--ink);padding:5px 7px;font-size:11px}` +
  `.create-quote-crd-input.invalid{border-color:var(--red);background:#fff1f2}.create-quote-status{margin-top:10px}.create-quote-version-note{font-weight:400;white-space:nowrap}.create-quote-table{table-layout:fixed}` +
  `.create-quote-table th:first-child,.create-quote-table td:first-child{width:58px}.create-quote-table th:not(:first-child),.create-quote-table td:not(:first-child){width:calc((100% - 58px)/3)}.compare-status{width:74px;` +
  `min-width:74px;white-space:nowrap}.compare-result-preview{margin-top:14px;border:1px solid #dbe5ee;border-radius:12px;background:#fff;overflow:hidden}.compare-result-tabs{display:flex;gap:5px;overflow-x:auto;padding:8px;` +
  `border-bottom:1px solid #dbe5ee;background:#f8fafc}.compare-result-tab{flex:0 0 auto;border:1px solid #cbd5e1;border-radius:8px;background:#fff;color:#475569;padding:6px 10px;font-size:10px;font-weight:700;cursor:pointer}` +
  `.compare-result-tab[aria-selected=true]{border-color:#0f766e;background:#0f766e;color:#fff}.compare-result-sheet-wrap{max-height:520px;overflow:auto}.compare-result-sheet{border-collapse:separate;border-spacing:0;` +
  `min-width:100%;table-layout:fixed;font-size:10px}.compare-result-sheet td{padding:4px 5px;border-right:1px solid #d6e0ea;border-bottom:1px solid #d6e0ea;vertical-align:top;white-space:pre-wrap;overflow-wrap:anywhere}` +
  `.compare-result-sheet tr:first-child td,.compare-result-sheet tr:nth-child(2) td{position:sticky;z-index:2}.compare-result-sheet tr:first-child td{top:0}.compare-result-sheet tr:nth-child(2) td{top:25px}` +
  `.compare-result-empty{padding:24px;color:#94a3b8;font-size:11px}.progress-track{height:9px;margin-top:12px;border-radius:999px;background:#e2e8f0;overflow:hidden}.progress-bar{height:100%;width:0;background:
linear-gradient(90deg,var(--blue),var(--teal));transition:width .2s ease}.progress-line{display:flex;align-items:center;justify-content:flex-end;gap:12px;margin-top:8px;color:#475569;font-size:11px}
.result-summary{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}.result-pill{border-radius:999px;background:#e2e8f0;color:#334155;padding:5px 9px;font-size:10px}.result-pill.ok{background:#dcfce7;
color:#166534}.result-pill.warn{background:#fef3c7;color:#92400e}@media(max-width:800px){.bridge-hero-inner{align-items:flex-start;flex-direction:column}.bridge-inline,.bridge-shell,.bridge-status,
.zbom-grid,.create-quote-grid{grid-template-columns:1fr}.bridge-card.wide{grid-column:auto}.toolbox-shell{padding:12px}.tool-tabs{overflow:auto}.planned-name{min-width:220px}.toolbox-modal{padding:12px}
.toolbox-modal-card{max-height:calc(100vh - 24px)}.history-entry{grid-template-columns:1fr;gap:4px}}`);
const BRIDGE_HTML=(`
<header class="bridge-hero">
  <div class="bridge-hero-inner">
    <div><h1><span aria-hidden="true">🧰</span> SAP Toolbox</h1></div>
    <button id="bridgeRefresh" class="bridge-btn ghost" type="button" data-bridge-action>Refresh</button>
  </div>
</header>
<main class="toolbox-shell">
  <section class="bridge-card connection-card">
    <div class="bridge-card-head"><div><h2>Connection</h2></div><span id="bridgeBadge" class="bridge-badge">Checking</span></div>
    <div class="bridge-status">
      <div class="bridge-status-item">
        <small>Local service</small>
        <strong id="bridgeService">Checking SAPGUIcontrol…</strong>
        <div id="bridgeUpdateStatus" class="bridge-note" style="margin-top:6px" aria-live="polite">Checking for updates…</div>
        <a id="bridgeDownload" class="bridge-btn light small bridge-download" href="${SAPGUICONTROL_DOWNLOAD_URL}" download="SAPGUIcontrol.zip" target="_blank" rel="noopener noreferrer"` +
  ` title="Download SAPGUIcontrol.zip, then extract it and run SAPGUIcontrol.exe." hidden>Download SAPGUIcontrol</a>
      </div>
      <div class="bridge-status-item"><small>SAP GUI</small><strong id="bridgeSap">Not checked</strong></div>
      <div class="bridge-status-item">
        <div class="bobj-status-head"><small>FID BOBJ Data</small><div class="bobj-status-actions"><button id="bobjUpdate" class="bridge-btn light small bobj-update" type="button" disabled>Update</button><button ` +
  `id="bobjOpen" class="bridge-btn light small bobj-update" type="button" disabled>Open</button></div></div>
        <strong id="bobjDataSource" class="bobj-data-source">Current Data Source: Checking…</strong>
        <span id="bobjRefreshState" class="bridge-note bobj-refresh-state">Checking refresh status…</span>
      </div>
      <div class="bridge-status-item">
        <div class="bobj-status-head"><small>CTL BOBJ</small><div class="bobj-status-actions"><button id="ctlBobjUpdate" class="bridge-btn light small bobj-update" type="button" disabled>Update</button><button ` +
  `id="ctlBobjOpen" class="bridge-btn light small bobj-update" type="button" disabled>Open</button></div></div>
        <strong id="ctlBobjDataSource" class="bobj-data-source">Current Data Source: Checking…</strong>
        <span id="ctlBobjRefreshState" class="bridge-note bobj-refresh-state">Checking refresh status…</span>
      </div>
    </div>
    <div class="bridge-field"><label id="sapSessionLabel" for="sapSession">SAP session · ZBOM PDF</label><select id="sapSession" disabled><option value="">Connecting to SAPGUIcontrol…</option></select><div ` +
  `class="bridge-actions"><span id="sapSessionState" class="bridge-note">Choose a window for this function.</span><button id="sapReleaseSession" class="bridge-btn light small" type="button" hidden>Release Retained ` +
  `Sessions</button></div></div>
  </section>

  <nav class="tool-tabs" role="tablist" aria-label="SAP Toolbox functions">
    <button id="tabZbom" class="tool-tab" type="button" role="tab" aria-selected="true" aria-controls="panelZbom" data-tool-tab="zbom">ZBOM PDF</button>
    <button id="tabZbomCompare" class="tool-tab" type="button" role="tab" aria-selected="false" aria-controls="panelZbomCompare" data-tool-tab="zbomCompare">ZBOM Compare</button>
    <button id="tabCreateQuote" class="tool-tab" type="button" role="tab" aria-selected="false" aria-controls="panelCreateQuote" data-tool-tab="createQuote">Quote Create</button>
    <button id="tabDebug" class="tool-tab" type="button" role="tab" aria-selected="false" aria-controls="panelDebug" data-tool-tab="debug" hidden>Debug</button>
  </nav>

  <section id="panelZbom" class="tool-panel" role="tabpanel" aria-labelledby="tabZbom" data-tool-panel="zbom">
    <div class="zbom-grid">
      <section class="bridge-card">
        <div class="bridge-card-head">
          <div><h2>1. Quotation Data</h2></div>
          <div class="bridge-card-tools">
            <button id="zbomOpenHistory" class="bridge-btn light small" type="button">History (<span id="zbomHistoryButtonCount" class="history-button-count">0</span>)</button>
            <button id="zbomResetInput" class="bridge-btn light small" type="button">Reset</button>
          </div>
        </div>
        <div class="bridge-field"><textarea id="zbomInput" class="quotation-input" rows="4" wrap="off" spellcheck="false" aria-label="Quotation data" placeholder="FID    System Description    Quote&#10;225822-4    ` +
  `WTSC_MAX/HALT_MAXx2    20736497"></textarea></div>
      </section>

      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2>2. Filename</h2></div><button id="zbomResetTemplate" class="bridge-btn light small" type="button">Reset</button></div>
        <div class="bridge-field"><input id="zbomTemplate" type="text" maxlength="240" value="{FID}_{System Description}_{Quote}_{Today}" autocomplete="off"></div>
        <div class="token-list" aria-label="Available file-name tokens">
          <button class="token-chip" type="button" data-zbom-token="{FID}">{FID}</button>
          <button class="token-chip" type="button" data-zbom-token="{System Description}">{System Description}</button>
          <button class="token-chip" type="button" data-zbom-token="{Quote}">{Quote}</button>
          <button class="token-chip" type="button" data-zbom-token="{Today}">{Today}</button>
        </div>
        <div class="folder-line"><span class="folder-name">Default output folder: C:\\PDFFILES</span></div>
      </section>

      <section class="bridge-card wide">
        <div class="preview-toolbar">
          <div><h2>3. Preview</h2></div>
          <span id="zbomCount" class="preview-count">0 rows</span>
        </div>
        <div class="zbom-table-wrap">
          <table class="zbom-table" aria-label="Detected ZBOM PDF input">
            <thead><tr><th>FID</th><th>System Description</th><th>Quote</th><th>Planned file name</th><th>Status</th></tr></thead>
            <tbody id="zbomPreviewBody"><tr><td colspan="5" class="missing">No preview yet.</td></tr></tbody>
          </table>
        </div>
      </section>

      <section class="bridge-card wide">
        <div class="bridge-card-head"><div><h2>4. Export PDF Files</h2></div></div>
        <div class="bridge-actions">
          <button id="zbomStart" class="bridge-btn blue" type="button" disabled>Start PDF Export</button>
          <button id="zbomOpenFolder" class="bridge-btn light" type="button">Open Folder</button>
        </div>
        <div class="progress-track" aria-hidden="true"><div id="zbomProgressBar" class="progress-bar"></div></div>
        <div class="progress-line"><span id="zbomProgressNumber">0%</span></div>
        <div id="zbomResultSummary" class="result-summary"></div>
      </section>
    </div>
  </section>

  <section id="panelZbomCompare" class="tool-panel" role="tabpanel" aria-labelledby="tabZbomCompare" data-tool-panel="zbomCompare" hidden>
    <div class="zbom-grid">
      <section class="bridge-card">
        <div class="bridge-card-head">
          <div><h2>1. Quotation Data</h2></div>
          <div class="bridge-card-tools">
            <button id="zbomCompareOpenHistory" class="bridge-btn light small" type="button">History (<span id="zbomCompareHistoryButtonCount" class="history-button-count">0</span>)</button>
            <button id="zbomCompareResetInput" class="bridge-btn light small" type="button">Reset</button>
          </div>
        </div>
        <div class="bridge-field"><textarea id="zbomCompareInput" class="quotation-input" rows="6" wrap="off" spellcheck="false" aria-label="ZBOM Compare quotation data" placeholder="Customer    FID    System Description    ` +
  `Quotation    Group&#10;Customer A    225822    WTSC_MAX/HALT_MAXx2    21314937    G1"></textarea></div>
        <div class="bridge-note">Paste one Quotation per line or multiple Excel columns.</div>
      </section>

      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2>2. Filename</h2></div><button id="zbomCompareResetTemplate" class="bridge-btn light small" type="button">Reset</button></div>
        <div class="bridge-field">
          <div class="compare-template-control">
            <input id="zbomCompareTemplate" type="text" maxlength="240" value="{Customer}_{FID}_{System Description}" autocomplete="off" aria-label="ZBOM Compare filename template">
            <input id="zbomCompareSeparator" class="compare-template-separator" type="text" maxlength="24" value="vs" autocomplete="off" spellcheck="false" aria-label="Text between compared file names" title="Text between ` +
  `compared file names">
          </div>
        </div>
        <div class="token-list" aria-label="Available ZBOM Compare file-name tokens">
          <button class="token-chip" type="button" data-zbom-compare-token="{Customer}">{Customer}</button>
          <button class="token-chip" type="button" data-zbom-compare-token="{FID}">{FID}</button>
          <button class="token-chip" type="button" data-zbom-compare-token="{System Description}">{System Description}</button>
          <button class="token-chip" type="button" data-zbom-compare-token="{Quote}">{Quote}</button>
        </div>
        <div class="bridge-note">{FID} stays as a placeholder when it is missing. The short field joins the two compared file names.</div>
        <div class="bridge-field">
          <label for="zbomCompareOutputFolder">Output Folder</label>
          <div class="compare-output-control">
            <input id="zbomCompareOutputFolder" type="text" maxlength="1000" value="C:\\PDFFILES\\Compare" autocomplete="off" spellcheck="false" placeholder="C:\\PDFFILES\\Compare">
            <button id="zbomCompareKeepTemp" class="bridge-btn light small compare-keep-temp" type="button" aria-pressed="false" title="Keep the temp folder after Compare finishes">SaveTemp</button>
          </div>
        </div>
      </section>

      <section class="bridge-card wide">
        <div class="preview-toolbar">
          <div><h2>3. Preview</h2></div>
          <span id="zbomCompareCount" class="preview-count">0 quotations</span>
        </div>
        <div class="zbom-table-wrap">
          <table class="zbom-table compare-table" aria-label="Detected ZBOM Compare input">
            <thead><tr><th>Quotation</th><th>FID</th><th>Customer</th><th>System Description</th><th>Compare Group</th><th>File Name</th><th class="compare-status">Status</th></tr></thead>
            <tbody id="zbomComparePreviewBody"><tr><td colspan="7" class="missing">No preview yet.</td></tr></tbody>
          </table>
        </div>
      </section>

      <section class="bridge-card wide">
        <div class="bridge-card-head"><div><h2>4. Compare</h2></div></div>
        <div class="bridge-actions">
          <button id="zbomCompareStart" class="bridge-btn blue" type="button" data-bridge-action disabled>Start Compare</button>
          <button id="zbomCompareOpenFolder" class="bridge-btn light" type="button">Open Folder</button>
        </div>
        <div class="progress-track" aria-hidden="true"><div id="zbomCompareProgressBar" class="progress-bar"></div></div>
        <div class="progress-line"><span id="zbomCompareProgressText">Ready.</span><span id="zbomCompareProgressNumber">0%</span></div>
        <div id="zbomCompareResultSummary" class="result-summary"></div>
        <div id="zbomCompareResultPreview" class="compare-result-preview" hidden>
          <div id="zbomCompareResultTabs" class="compare-result-tabs" role="tablist" aria-label="ZBOM Compare result sheets"></div>
          <div id="zbomCompareResultSheet" class="compare-result-sheet-wrap"></div>
        </div>
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
            <button id="createQuoteOpenHistory" class="bridge-btn light small" type="button" data-bridge-action>History (<span id="createQuoteHistoryButtonCount" class="history-button-count">0</span>)</button>
            <button id="createQuoteResetInput" class="bridge-btn light small" type="button" data-bridge-action>Reset</button>
          </div>
        </div>
        <div class="bridge-field"><textarea id="createQuoteInput" class="quotation-input" rows="4" wrap="off" spellcheck="false" aria-label="Create Quote FID and date data" placeholder="FID    CRD    ValidTo&#10;264072    ` +
  `08/01/2026    09/15/2026"></textarea></div>
      </section>

      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2>3. Settings</h2></div><button id="createQuoteResetSettings" class="bridge-btn light small" type="button" data-bridge-action>Reset</button></div>
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
        <div class="bridge-actions">
          <button id="createQuoteStart" class="bridge-btn blue" type="button" disabled>Create Quote</button>
        </div>
        <div id="createQuoteStatus" class="bridge-note create-quote-status">Ready.</div>
      </section>
    </div>
  </section>

  <section id="panelDebug" class="tool-panel" role="tabpanel" aria-labelledby="tabDebug" data-tool-panel="debug" hidden>
    <div class="bridge-shell">
      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2>Open Transaction</h2></div></div>
        <div class="bridge-field"><label for="sapTransaction">Transaction code</label><input id="sapTransaction" type="text" placeholder="For example: VA03 or /SCWM/MON" autocomplete="off"></div>
        <div class="bridge-actions"><button id="sapRunTransaction" class="bridge-btn blue" type="button" data-bridge-action disabled>Open transaction</button></div>
      </section>

      <section class="bridge-card">
        <div class="bridge-card-head"><div><h2>Control Inspector</h2></div></div>
        <div class="bridge-field"><label for="sapControlId">SAP control ID</label><input id="sapControlId" type="text" placeholder="wnd[0]/usr/ctxt…" autocomplete="off"></div>
        <div class="bridge-field"><label for="sapControlValue">Value</label><input id="sapControlValue" type="text" placeholder="Text or key to write" autocomplete="off"></div>
        <div class="bridge-actions">
          <button id="sapReadText" class="bridge-btn light" type="button" data-bridge-action disabled>Read text</button>
          <button id="sapSetText" class="bridge-btn blue" type="button" data-bridge-action disabled>Set text</button>
          <button id="sapSetKey" class="bridge-btn light" type="button" data-bridge-action disabled>Set key</button>
          <button id="sapPress" class="bridge-btn light" type="button" data-bridge-action disabled>Press</button>
          <button id="sapSelect" class="bridge-btn light" type="button" data-bridge-action disabled>Select</button>
        </div>
        <div class="bridge-divider"></div>
        <div class="bridge-inline">
          <div class="bridge-field"><label for="sapVKeyTarget">Window ID</label><input id="sapVKeyTarget" type="text" value="wnd[0]" autocomplete="off"></div>
          <div class="bridge-field"><label for="sapVKey">VKey</label><input id="sapVKey" type="number" value="0"></div>
        </div>
        <div class="bridge-actions"><button id="sapSendVKey" class="bridge-btn light" type="button" data-bridge-action disabled>Send VKey</button></div>
      </section>

      <section class="bridge-card wide">
        <div class="bridge-card-head"><div><h2>Run Script</h2></div></div>
        <div class="bridge-inline">
          <div class="bridge-field"><label for="nativeLanguage">Language</label><select id="nativeLanguage"><option value="powershell">PowerShell</option><option value="vba">VBA (Excel)</option></select></div>
          <div id="nativeEntryPointField" class="bridge-field" hidden><label for="nativeEntryPoint">VBA entry point</label><input id="nativeEntryPoint" value="Main" autocomplete="off"></div>
        </div>
        <div class="bridge-field"><label for="nativeScript">Script</label><textarea id="nativeScript" class="zbom-input" rows="9" spellcheck="false" placeholder="PowerShell script or VBA module"></textarea></div>
        <div class="bridge-actions"><button id="nativeRun" class="bridge-btn blue" type="button">Run Script</button></div>
      </section>

      <section class="bridge-card wide">
        <div class="bridge-card-head"><div><h2>Result</h2></div></div>
        <pre id="sapResult" class="bridge-result" aria-live="polite">Waiting for SAPGUIcontrol…</pre>
      </section>
    </div>
  </section>
</main>
<div id="zbomHistoryDialog" class="toolbox-modal" hidden>
  <div class="toolbox-modal-backdrop" data-zbom-history-close></div>
  <section class="toolbox-modal-card" role="dialog" aria-modal="true" aria-labelledby="zbomHistoryTitle">
    <div class="toolbox-modal-head">
      <div class="toolbox-modal-title"><h2 id="zbomHistoryTitle">Quotation History</h2><span id="zbomHistoryCount" class="history-count">0/50</span></div>
      <button id="zbomCloseHistory" class="bridge-btn light small" type="button">Close</button>
    </div>
    <div id="zbomHistoryList" class="history-list"><div class="history-empty">No export history.</div></div>
  </section>
</div>
<div id="zbomCompareHistoryDialog" class="toolbox-modal" hidden>
  <div class="toolbox-modal-backdrop" data-zbom-compare-history-close></div>
  <section class="toolbox-modal-card" role="dialog" aria-modal="true" aria-labelledby="zbomCompareHistoryTitle">
    <div class="toolbox-modal-head">
      <div class="toolbox-modal-title"><h2 id="zbomCompareHistoryTitle">ZBOM Compare History</h2><span id="zbomCompareHistoryCount" class="history-count">0/50</span></div>
      <button id="zbomCompareCloseHistory" class="bridge-btn light small" type="button">Close</button>
    </div>
    <div id="zbomCompareHistoryList" class="history-list"><div class="history-empty">No compare history.</div></div>
  </section>
</div>
<div id="createQuoteHistoryDialog" class="toolbox-modal" hidden>
  <div class="toolbox-modal-backdrop" data-create-quote-history-close></div>
  <section class="toolbox-modal-card" role="dialog" aria-modal="true" aria-labelledby="createQuoteHistoryTitle">
    <div class="toolbox-modal-head">
      <div class="toolbox-modal-title"><h2 id="createQuoteHistoryTitle">Quote Create History</h2><span id="createQuoteHistoryCount" class="history-count">0/50</span></div>
      <button id="createQuoteCloseHistory" class="bridge-btn light small" type="button">Close</button>
    </div>
    <div id="createQuoteHistoryList" class="history-list"><div class="history-empty">No Quote Create history.</div></div>
  </section>
</div>`);
const ZBOM_PDF_FOLDER='C:\\PDFFILES'; const ZBOM_DEFAULT_TEMPLATE='{FID}_{System Description}_{Quote}_{Today}'; const ZBOM_HISTORY_LIMIT=50; const ZBOM_STORAGE_KEYS={input:'nsrSapToolbox.zbomInput.v1',template:'nsrSapToolbox.zbomTemplate.v1',history:'nsrSapToolbox.zbomHistory.v1'};
const ZBOM_COMPARE_DEFAULT_TEMPLATE='{Customer}_{FID}_{System Description}'; const ZBOM_COMPARE_DEFAULT_SEPARATOR='vs'; const ZBOM_COMPARE_DEFAULT_GROUP='G01'; const ZBOM_COMPARE_DEFAULT_FOLDER='C:\\PDFFILES\\Compare'; const ZBOM_COMPARE_STANDALONE='__STANDALONE__';
const ZBOM_COMPARE_FID_ERROR='FID Read Error'; const ZBOM_COMPARE_HEADER_FILL='FF0070C0'; const ZBOM_COMPARE_BORDER_COLOR='FFD6E0EA'; const ZBOM_COMPARE_FILE_COLORS=['FFFFC7CE','FFC6EFCE','FFFFEB9C','FFBDD7EE','FFE4DFEC','FFFCE4D6','FFDDEBF7','FFE2F0D9','FFF4B183','FFD9E1F2'];
const ZBOM_COMPARE_STORAGE_KEYS={input:'nsrSapToolbox.zbomCompareInput.v1',template:'nsrSapToolbox.zbomCompareTemplate.v1',separator:'nsrSapToolbox.zbomCompareSeparator.v1',folder:'nsrSapToolbox.zbomCompareFolder.v1',groups:'nsrSapToolbox.zbomCompareGroups.v1',
  history:'nsrSapToolbox.zbomCompareHistory.v1'};
const SAP_ACTION_TIMEOUT=600000,SAP_BUSY_TIMEOUT=120000,SAP_FILE_WAIT_TIMEOUT=120000,SAP_PDF_FIRST_FILE_WAIT_MS=15000,SAP_FILE_SETTLE_MS=5000,SAP_BATCH_SETTLE_MS=15000,SAP_FILE_POLL_MS=400,BRIDGE_FILE_TIMEOUT=600000; const CREATE_QUOTE_DEFAULT_VALID_TO_DAYS='30';
const CREATE_QUOTE_DEFAULT_FID_SUFFIX='-01'; const CREATE_QUOTE_HISTORY_LIMIT=50;
const CREATE_QUOTE_STORAGE_KEYS={quotation:'nsrSapToolbox.createQuoteQuotation.v1',input:'nsrSapToolbox.createQuoteInput.v1',validToDays:'nsrSapToolbox.createQuoteValidToDays.v1',fidSuffix:'nsrSapToolbox.createQuoteFidSuffix.v1',crdOverrides:'nsrSapToolbox.createQuoteCrdOverrides.v1',
  history:'nsrSapToolbox.createQuoteHistory.v1'};
const BRIDGE_DEBUG_STORAGE_KEY='nsrSapToolbox.debugUnlocked.v1'; const SAP_TOOL_LABELS={zbom:'ZBOM PDF',zbomCompare:'ZBOM Compare',createQuote:'Quote Create',debug:'Debug'}; const sapGuiControlUpdate={online:null,current:'',release:null,checking:true};
const bridgePageId=crypto.randomUUID(); const sapTools=Object.fromEntries(Object.keys(SAP_TOOL_LABELS).map(name=>[name,{selection:'',starting:false,job:null}])); let bobjTaskRefreshPromise=null;
let activeSapTool='zbom',sharedSapTasks=[],bridgeSupportsTasks=false,bridgeTaskOnline=false,bridgeTaskPollTimer=0,bridgeTaskPolling=false,bridgeSessionRefreshPromise=null,bridgeDebugLog='';
let bridgeBusy=false,bridgeRefreshPending=false,bridgeReconnectTimer=0,bridgeSessions=[],bridgeMock=false,zbomRows=[],zbomErrors=[],zbomRunning=false,zbomHasRun=false,zbomLogLines=[],zbomOutputClaims=new Map(),zbomResults=new Map(),zbomHistory=[],zbomHistoryReturnFocus=null;
let zbomCompareRows=[],zbomCompareStats={duplicates:0,ignored:0},zbomCompareSavedGroups=new Map(),zbomCompareRunning=false,zbomCompareHasRun=false,zbomCompareResults=new Map(),zbomCompareLogLines=[],zbomCompareOutputRowHint=4,zbomCompareHistory=[],zbomCompareHistoryReturnFocus=null,
  zbomCompareResultTabs=[],zbomCompareActiveResultTab='';
let createQuoteRows=[],createQuoteStats={duplicates:0,ignored:0},createQuoteCrdOverrides=new Map(),createQuoteManualCrd=new Set(),createQuoteHistory=[],createQuoteHistoryReturnFocus=null; let bridgeDebugUnlocked=false;
let ctlBobjStatus=null,ctlBobjBusy=false,ctlBobjTimer=0,ctlBobjRun=0;
let bobjStatus=null,bobjRefreshing=false,bobjOpening=false,bobjPollTimer=0,bobjStatusRun=0,bobjDataRows=new Map(),bobjDataHeaderIndexes=new Map(),bobjDataMissing=new Set(),bobjDataPath='',bobjDataModified=0,bobjDataCount=0,bobjDataReady=false,bobjDataLoadRun=0,
  bobjConsumersPending=false,bobjSourcePath='',bobjSourceModified=0,bobjDataLoading=false,bobjDataPromise=null,bobjWorker=null,bobjWorkerRequest=0,bobjWorkerPending=new Map();

GM_registerMenuCommand('📊 NSR Tracking Dashboard',()=>unsafeWindow.open(APP_PATH+'?nsrct=1','_blank','noopener')); GM_registerMenuCommand('🏭 FCID Governance Control Tower',()=>unsafeWindow.open('https://fep.lamresearch.com'+FCID_PATH,'_blank','noopener'));
GM_registerMenuCommand('🧰 SAP Toolbox',()=>unsafeWindow.open(BRIDGE_PATH+'?nsrbridge=1','_blank','noopener')); GM_registerMenuCommand('📖 Help',()=>unsafeWindow.open(GUIDE_URL,'_blank','noopener'));

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
if(currentPath===APP_PATH){whenBody(mountApp);}else if(currentPath===BRIDGE_PATH){whenBody(mountBridgeApp);}else if(currentPath===FCID_PATH&&unsafeWindow.location.hostname.toLowerCase()==='fep.lamresearch.com'){whenBody(runFcidGovernance);}
  else if(unsafeWindow.location.hostname.toLowerCase()==='fep.lamresearch.com')addNsrHeaderButton();

function whenBody(callback){if(document.body)callback();else document.addEventListener('DOMContentLoaded',callback,{once:true});}

async function addNsrHeaderButton(){
  while(true){
    const target=document.getElementById('shellAppTitle-button'),title=target&&target.querySelector('.sapUshellAppTitleText');
    if(target&&String(title&&title.textContent||'').trim().toUpperCase()==='NSR'&&!document.getElementById(NSR_HEADER_BUTTON_ID)){
      const button=document.createElement('button');button.id=NSR_HEADER_BUTTON_ID;button.type='button';button.title='Open NSR Tracking Dashboard';button.textContent='📊 NSR Tracking'; button.style.cssText='margin-left:8px;padding:4px 8px;cursor:pointer';
      button.addEventListener('click',event=>{event.preventDefault();event.stopPropagation();unsafeWindow.open(APP_PATH+'?nsrct=1','_blank','noopener')}); target.insertAdjacentElement('afterend',button);
    }
    await new Promise(resolve=>setTimeout(resolve,500));
  }
}

function persistentStore(){try{return unsafeWindow.localStorage}catch(_){try{return localStorage}catch(__){return null}}}
function readPersistent(key,fallback){const storage=persistentStore();if(!storage)return fallback; try{const value=storage.getItem(key);return value===null?fallback:value}catch(_){return fallback}}
function writePersistent(key,value){const storage=persistentStore();if(!storage)return; try{storage.setItem(key,String(value))}catch(_){/* ignore unavailable or full storage */}}

function loadZbomPreferences(){
  const input=bridgeId('zbomInput'),template=bridgeId('zbomTemplate'),savedInput=readPersistent(ZBOM_STORAGE_KEYS.input,null),savedTemplate=readPersistent(ZBOM_STORAGE_KEYS.template,null);
  if(savedInput!==null)input.value=savedInput;
  if(savedTemplate!==null)template.value=savedTemplate;
  const savedHistory=readPersistent(ZBOM_STORAGE_KEYS.history,'[]');
  try{const parsed=JSON.parse(savedHistory); zbomHistory=Array.isArray(parsed)?parsed.filter(item=>item&&typeof item.text==='string'&&typeof item.runAt==='string').slice(0,ZBOM_HISTORY_LIMIT):[];}catch(_){zbomHistory=[]}
  renderZbomHistory();
}

function saveZbomInput(){const input=bridgeId('zbomInput');if(input)writePersistent(ZBOM_STORAGE_KEYS.input,input.value);}
function saveZbomTemplate(){const input=bridgeId('zbomTemplate');if(input)writePersistent(ZBOM_STORAGE_KEYS.template,input.value);}

function loadZbomComparePreferences(){
  const input=bridgeId('zbomCompareInput'),template=bridgeId('zbomCompareTemplate'),separator=bridgeId('zbomCompareSeparator'),folder=bridgeId('zbomCompareOutputFolder'),savedInput=readPersistent(ZBOM_COMPARE_STORAGE_KEYS.input,null),
    savedTemplate=readPersistent(ZBOM_COMPARE_STORAGE_KEYS.template,null),savedSeparator=readPersistent(ZBOM_COMPARE_STORAGE_KEYS.separator,null),savedFolder=readPersistent(ZBOM_COMPARE_STORAGE_KEYS.folder,null);
  if(savedInput!==null)input.value=savedInput;
  if(savedTemplate!==null){template.value=/^\{fid\}$/i.test(String(savedTemplate).trim())?ZBOM_COMPARE_DEFAULT_TEMPLATE:savedTemplate;if(template.value!==savedTemplate)writePersistent(ZBOM_COMPARE_STORAGE_KEYS.template,template.value)}
  if(savedSeparator!==null)separator.value=savedSeparator;
  folder.value=savedFolder!==null&&String(savedFolder).trim()?savedFolder:ZBOM_COMPARE_DEFAULT_FOLDER; zbomCompareSavedGroups=new Map();
  try{
    const parsed=JSON.parse(readPersistent(ZBOM_COMPARE_STORAGE_KEYS.groups,'{}'));
    if(parsed&&typeof parsed==='object'&&!Array.isArray(parsed))Object.entries(parsed).forEach(([quote,value])=>{if(!/^2\d{7}$/.test(quote)||typeof value!=='string')return; const group=value===ZBOM_COMPARE_STANDALONE||value===
      '__DO_NOT_COMPARE__'?ZBOM_COMPARE_STANDALONE:normalizeZbomCompareGroup(value,true); if(group)zbomCompareSavedGroups.set(quote,group);});
  }catch(_){zbomCompareSavedGroups=new Map()}
  try{const parsed=JSON.parse(readPersistent(ZBOM_COMPARE_STORAGE_KEYS.history,'[]')); zbomCompareHistory=Array.isArray(parsed)?parsed.filter(item=>item&&typeof item.text==='string'&&typeof item.runAt==='string').slice(0,ZBOM_HISTORY_LIMIT):[];}catch(_){zbomCompareHistory=[]}
  renderZbomCompareHistory();
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
  try{const parsed=JSON.parse(readPersistent(CREATE_QUOTE_STORAGE_KEYS.history,'[]')); createQuoteHistory=Array.isArray(parsed)?parsed.filter(item=>item&&typeof item.quotation==='string'&&typeof item.input==='string'&&typeof item.runAt==='string').slice(0,
    CREATE_QUOTE_HISTORY_LIMIT):[];}catch(_){createQuoteHistory=[]}
  renderCreateQuoteHistory();
}

function saveCreateQuoteQuotation(){const input=bridgeId('createQuoteQuotation');if(input)writePersistent(CREATE_QUOTE_STORAGE_KEYS.quotation,input.value);}
function saveCreateQuoteInput(){const input=bridgeId('createQuoteInput');if(input)writePersistent(CREATE_QUOTE_STORAGE_KEYS.input,input.value);}
function saveCreateQuoteValidToDays(){const input=bridgeId('createQuoteDefaultValidTo');if(input)writePersistent(CREATE_QUOTE_STORAGE_KEYS.validToDays,input.value);}
function saveCreateQuoteFidSuffix(){const input=bridgeId('createQuoteFidSuffix');if(input)writePersistent(CREATE_QUOTE_STORAGE_KEYS.fidSuffix,input.value);}
function saveCreateQuoteCrdOverrides(){const values={};createQuoteCrdOverrides.forEach((value,fid)=>{values[fid]=value});writePersistent(CREATE_QUOTE_STORAGE_KEYS.crdOverrides,JSON.stringify(values));}
function keepZbomCompareTemp(){return bridgeId('zbomCompareKeepTemp').getAttribute('aria-pressed')==='true'}
function saveZbomCompareGroups(){const values={};zbomCompareSavedGroups=new Map(); zbomCompareRows.forEach(row=>{if(row.quote&&row.group){values[row.quote]=row.group;zbomCompareSavedGroups.set(row.quote,row.group)}}); writePersistent(ZBOM_COMPARE_STORAGE_KEYS.groups,
  JSON.stringify(values));}

async function openZbomCompareOutputFolder(){
  const input=bridgeId('zbomCompareOutputFolder');let path;
  try{path=normalizedZbomCompareFolder()}catch(error){showBridgeResult(bridgeError(error),'Unable to open the ZBOM Compare output folder.');input.focus();return}
  input.value=path;saveZbomCompareFolder();
  try{await bridgeFileRequest('createDirectory',{path:path},30000);const data=await openBridgeFolder(path);showBridgeResult(data,'ZBOM Compare output folder opened.');}catch(error){showBridgeResult(bridgeError(error),'Unable to open the ZBOM Compare output folder.')}
}

async function openZbomCompareWorkbook(path,button){
  if(!path)return;
  if(button)button.disabled=true;
  try{await bridgeFileRequest('openFile',{path:path},30000);setZbomCompareProgress(100,'Opened '+String(path).split(/[\\/]/).pop()+'.');}catch(error){setZbomCompareProgress(100,'Unable to open workbook: '+bridgeError(error))} finally{if(button)button.disabled=false}
}

function clearZbomCompareResultPreview(){zbomCompareResultTabs=[];zbomCompareActiveResultTab=''; const preview=bridgeId('zbomCompareResultPreview'),tabs=bridgeId('zbomCompareResultTabs'),sheet=bridgeId('zbomCompareResultSheet'); if(preview)preview.hidden=true;
  if(tabs)tabs.replaceChildren();if(sheet)sheet.replaceChildren();}
function invalidateZbomCompareResults(){zbomCompareHasRun=false;zbomCompareResults=new Map();bridgeId('zbomCompareResultSummary').replaceChildren();clearZbomCompareResultPreview();}

function renderZbomCompareResultSheet(){
  const container=bridgeId('zbomCompareResultSheet'),tab=zbomCompareResultTabs.find(item=>item.key===zbomCompareActiveResultTab);if(!container)return;
  container.replaceChildren();
  if(!tab||!tab.sheet||!tab.sheet.rows.length){const empty=document.createElement('div');empty.className='compare-result-empty';empty.textContent='No table content.';container.appendChild(empty);return}
  const table=document.createElement('table'),colgroup=document.createElement('colgroup'),body=document.createElement('tbody');table.className='compare-result-sheet';
  const widths=tab.sheet.widths.map(width=>{const size=Number(width)||12;return size<=5?20:Math.max(56,Math.min(140,Math.round(size*3.8)))}); widths.forEach(width=>{const column=document.createElement('col');column.style.width=width+'px';colgroup.appendChild(column)});
  table.style.width=widths.reduce((total,width)=>total+width,0)+'px';
  tab.sheet.rows.forEach(row=>{
    const tr=document.createElement('tr');row.forEach(item=>{const cell=document.createElement('td');cell.textContent=item.text; if(item.fill)cell.style.backgroundColor='#'+item.fill.slice(-6); if(item.fontColor)cell.style.color='#'+item.fontColor.slice(-6);
      if(item.bold)cell.style.fontWeight='700'; tr.appendChild(cell);});body.appendChild(tr);
  });
  table.append(colgroup,body);container.appendChild(table);
}

function renderZbomCompareResultTabs(preferredLabel){
  const preview=bridgeId('zbomCompareResultPreview'),container=bridgeId('zbomCompareResultTabs');if(!preview||!container)return;
  if(!zbomCompareResultTabs.length){clearZbomCompareResultPreview();return}
  if(preferredLabel){const preferred=zbomCompareResultTabs.find(item=>item.label===preferredLabel);if(preferred)zbomCompareActiveResultTab=preferred.key}
  if(!zbomCompareResultTabs.some(item=>item.key===zbomCompareActiveResultTab))zbomCompareActiveResultTab=zbomCompareResultTabs[0].key;
  preview.hidden=false;container.replaceChildren();
  zbomCompareResultTabs.forEach(tab=>{
    const button=document.createElement('button');button.type='button';button.className='compare-result-tab';button.setAttribute('role','tab');button.setAttribute('aria-selected',tab.key===zbomCompareActiveResultTab?'true':'false');button.textContent=tab.label;
    button.addEventListener('click',()=>{zbomCompareActiveResultTab=tab.key;renderZbomCompareResultTabs()});container.appendChild(button);
  });
  renderZbomCompareResultSheet();
}

function showZbomCompareResultPreview(groupPreviews,fidPreviews){
  const groups=(groupPreviews||[]).slice().sort((left,right)=>{
    const leftMatch=String(left.group).match(/^G(\d+)$/i),rightMatch=String(right.group).match(/^G(\d+)$/i);
    if(leftMatch&&rightMatch)return Number(leftMatch[1])-Number(rightMatch[1]);
    if(leftMatch)return -1;if(rightMatch)return 1;return String(left.group).localeCompare(String(right.group));
  });
  const tabs=[]; groups.forEach(item=>{if(item.options)tabs.push({key:'group:'+item.group,label:item.group,sheet:item.options});if(item.gas)tabs.push({key:'group-gas:'+item.group,label:item.group+'_Gas',sheet:item.gas})});
  Array.from((fidPreviews||new Map()).values()).sort((left,right)=>left.order-right.order).forEach(item=>{if(item.sheet)tabs.push({key:'fid:'+item.fid,label:item.fid,sheet:item.sheet})}); zbomCompareResultTabs=tabs;zbomCompareActiveResultTab='';renderZbomCompareResultTabs('G01');
}

function historyDisplayTime(value){const date=new Date(value);return Number.isNaN(date.getTime())?String(value):date.toLocaleString([], {hour12:false});}
function historyDisplayText(value){const lines=String(value||'').split(/\r?\n/).map(line=>line.trim()).filter(Boolean); if(!lines.length)return '(empty)'; const preview=lines.slice(0,4); if(lines.length>4)preview.push('… '+lines.length+' lines total'); return preview.join('\n');}
function openZbomHistory(){const dialog=bridgeId('zbomHistoryDialog');if(!dialog)return; closeZbomCompareHistory();closeCreateQuoteHistory(); zbomHistoryReturnFocus=document.activeElement;renderZbomHistory();dialog.hidden=false;document.body.classList.add('history-open');
  bridgeId('zbomCloseHistory').focus();}

function closeZbomHistory(){
  const dialog=bridgeId('zbomHistoryDialog');if(!dialog||dialog.hidden)return;
  dialog.hidden=true;document.body.classList.remove('history-open'); const returnFocus=zbomHistoryReturnFocus;zbomHistoryReturnFocus=null;
  if(returnFocus&&typeof returnFocus.focus==='function')returnFocus.focus();
}

function renderZbomHistory(){
  const list=bridgeId('zbomHistoryList'),count=bridgeId('zbomHistoryCount'),buttonCount=bridgeId('zbomHistoryButtonCount');
  if(buttonCount)buttonCount.textContent=String(zbomHistory.length);
  if(!list||!count)return;
  count.textContent=zbomHistory.length+'/'+ZBOM_HISTORY_LIMIT;list.replaceChildren();
  if(!zbomHistory.length){const empty=document.createElement('div');empty.className='history-empty';empty.textContent='No export history.';list.appendChild(empty);return}
  zbomHistory.forEach((entry,index)=>{
    const displayTime=historyDisplayTime(entry.runAt),button=document.createElement('button');button.type='button';button.className='history-entry';button.dataset.historyIndex=String(index);button.setAttribute('aria-label','Restore quotation data from '+displayTime);
    const time=document.createElement('span');time.className='history-entry-time';time.textContent=displayTime; const text=document.createElement('span');text.className='history-entry-text';text.textContent=historyDisplayText(entry.text); button.append(time,text);
      button.addEventListener('click',()=>{const current=zbomHistory[index];if(!current)return; const input=bridgeId('zbomInput');input.value=current.text;saveZbomInput();previewZbomInput();closeZbomHistory();input.focus();}); list.appendChild(button);
  });
}

function recordZbomHistory(text){zbomHistory=[{text:String(text||''),runAt:new Date().toISOString()},...zbomHistory].slice(0,ZBOM_HISTORY_LIMIT); writePersistent(ZBOM_STORAGE_KEYS.history,JSON.stringify(zbomHistory));renderZbomHistory();}

function openZbomCompareHistory(){
  const dialog=bridgeId('zbomCompareHistoryDialog');if(!dialog)return;
  closeZbomHistory();closeCreateQuoteHistory(); zbomCompareHistoryReturnFocus=document.activeElement;renderZbomCompareHistory();dialog.hidden=false;document.body.classList.add('history-open');bridgeId('zbomCompareCloseHistory').focus();
}

function closeZbomCompareHistory(){
  const dialog=bridgeId('zbomCompareHistoryDialog');if(!dialog||dialog.hidden)return;
  dialog.hidden=true;document.body.classList.remove('history-open'); const returnFocus=zbomCompareHistoryReturnFocus;zbomCompareHistoryReturnFocus=null;
  if(returnFocus&&typeof returnFocus.focus==='function')returnFocus.focus();
}

function renderZbomCompareHistory(){
  const list=bridgeId('zbomCompareHistoryList'),count=bridgeId('zbomCompareHistoryCount'),buttonCount=bridgeId('zbomCompareHistoryButtonCount');
  if(buttonCount)buttonCount.textContent=String(zbomCompareHistory.length);
  if(!list||!count)return;
  count.textContent=zbomCompareHistory.length+'/'+ZBOM_HISTORY_LIMIT;list.replaceChildren();
  if(!zbomCompareHistory.length){const empty=document.createElement('div');empty.className='history-empty';empty.textContent='No compare history.';list.appendChild(empty);return}
  zbomCompareHistory.forEach((entry,index)=>{
    const displayTime=historyDisplayTime(entry.runAt),button=document.createElement('button');button.type='button';button.className='history-entry';button.setAttribute('aria-label','Restore ZBOM Compare data from '+displayTime); const time=document.createElement('span');
      time.className='history-entry-time';time.textContent=displayTime; const text=document.createElement('span');text.className='history-entry-text';text.textContent=historyDisplayText(entry.text); button.append(time,text);button.addEventListener('click',
      ()=>{const current=zbomCompareHistory[index];if(!current)return; const input=bridgeId('zbomCompareInput');input.value=current.text;saveZbomCompareInput();previewZbomCompareInput();closeZbomCompareHistory();input.focus();}); list.appendChild(button);
  });
}

function recordZbomCompareHistory(text){zbomCompareHistory=[{text:String(text||''),runAt:new Date().toISOString()},...zbomCompareHistory].slice(0,ZBOM_HISTORY_LIMIT); writePersistent(ZBOM_COMPARE_STORAGE_KEYS.history,JSON.stringify(zbomCompareHistory));renderZbomCompareHistory();}

function openCreateQuoteHistory(){
  const dialog=bridgeId('createQuoteHistoryDialog');if(!dialog)return;
  closeZbomHistory();closeZbomCompareHistory();createQuoteHistoryReturnFocus=document.activeElement;renderCreateQuoteHistory();dialog.hidden=false;document.body.classList.add('history-open');bridgeId('createQuoteCloseHistory').focus();
}

function closeCreateQuoteHistory(){
  const dialog=bridgeId('createQuoteHistoryDialog');if(!dialog||dialog.hidden)return;
  dialog.hidden=true;document.body.classList.remove('history-open'); const returnFocus=createQuoteHistoryReturnFocus;createQuoteHistoryReturnFocus=null;if(returnFocus&&typeof returnFocus.focus==='function')returnFocus.focus();
}

function renderCreateQuoteHistory(){
  const list=bridgeId('createQuoteHistoryList'),count=bridgeId('createQuoteHistoryCount'),buttonCount=bridgeId('createQuoteHistoryButtonCount');
  if(buttonCount)buttonCount.textContent=String(createQuoteHistory.length);if(!list||!count)return;
  count.textContent=createQuoteHistory.length+'/'+CREATE_QUOTE_HISTORY_LIMIT;list.replaceChildren();
  if(!createQuoteHistory.length){const empty=document.createElement('div');empty.className='history-empty';empty.textContent='No Quote Create history.';list.appendChild(empty);return}
  createQuoteHistory.forEach((entry,index)=>{
    const displayTime=historyDisplayTime(entry.runAt),button=document.createElement('button');button.type='button';button.className='history-entry';button.setAttribute('aria-label','Restore Quote Create data from '+displayTime); const time=document.createElement('span');
      time.className='history-entry-time';time.textContent=displayTime; const text=document.createElement('span');text.className='history-entry-text';text.textContent='Reference Quotation: '+(String(entry&&entry.quotation||'').trim()||'(empty)')+'\nFID:\n'+historyDisplayText(entry&&
      entry.input);button.append(time,text);button.addEventListener('click',
      ()=>{
      const current=createQuoteHistory[index];if(!current)return;
      const quotation=bridgeId('createQuoteQuotation'),input=bridgeId('createQuoteInput');quotation.value=current.quotation;input.value=current.input;sanitizeCreateQuoteQuotation();saveCreateQuoteQuotation();saveCreateQuoteInput();previewCreateQuoteInput();
        setCreateQuoteStatus('Ready.');ensureBobjForCurrentInput();closeCreateQuoteHistory();input.focus();
    });list.appendChild(button);
  });
}

function recordCreateQuoteHistory(quotation,input){createQuoteHistory=[{quotation:String(quotation||''),input:String(input||''),runAt:new Date().toISOString()},...createQuoteHistory].slice(0,CREATE_QUOTE_HISTORY_LIMIT); writePersistent(CREATE_QUOTE_STORAGE_KEYS.history,
  JSON.stringify(createQuoteHistory));renderCreateQuoteHistory();}

function emojiFavicon(emoji){
  const icon=document.createElement('link');icon.rel='icon';icon.type='image/svg+xml';icon.sizes='any';
  icon.href='data:image/svg+xml,'+encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><text x="32" y="52" text-anchor="middle" font-size="54" font-family="Segoe UI Emoji, Apple Color Emoji, Noto Color Emoji, sans-serif">${emoji}</text></svg>`); return icon;
}

function mountApp(){
  if(toolkitExpired())return;
  if(document.documentElement.dataset.nsrCtMounted)return;
  document.documentElement.dataset.nsrCtMounted='1';document.documentElement.lang='en'; const charset=document.createElement('meta');charset.charset='utf-8'; const viewport=document.createElement('meta');viewport.name='viewport';viewport.content='width=device-width,initial-scale=1';
  const title=document.createElement('title');title.textContent='NSR Tracking Dashboard'; const style=document.createElement('style');style.textContent=APP_CSS; document.head.replaceChildren(charset,viewport,title,emojiFavicon('📊'),style);document.body.innerHTML=APP_HTML;
  runDashboard();
}

function mountBridgeApp(){
  if(toolkitExpired())return;
  if(document.documentElement.dataset.nsrSapBridgeMounted)return;
  document.documentElement.dataset.nsrSapBridgeMounted='1';document.documentElement.lang='en'; const charset=document.createElement('meta');charset.charset='utf-8'; const viewport=document.createElement('meta');viewport.name='viewport';viewport.content=
    'width=device-width,initial-scale=1'; const title=document.createElement('title');title.textContent='SAP Toolbox'; const style=document.createElement('style');style.textContent=BRIDGE_CSS;
  document.head.replaceChildren(charset,viewport,title,emojiFavicon('🧰'),style);document.body.innerHTML=BRIDGE_HTML; bindBridgeEvents();refreshBridge();
}

function bridgeId(id){return document.getElementById(id)}
function runBridgePowerShell(script,data,options={}){return bridgeRequest('POST','/execute',{script:script,data:data,workingDirectory:options.workingDirectory},options.timeout??0);}

function runBridgeVba(code,options={}){
  return runBridgePowerShell(String.raw`$ErrorActionPreference = 'Stop'
$app = $null; $document = $null
try {
  $app = New-Object -ComObject $data.application
  $app.Visible = [bool]$data.visible
  $app.DisplayAlerts = 0
  $document = $app.($data.collection).Add()
  $module = $document.VBProject.VBComponents.Add(1)
  $module.CodeModule.AddFromString([string]$data.code)
  $arguments = [object[]]@($data.entryPoint)
  if ($null -ne $data.arguments) { $arguments += [object[]]@($data.arguments) }
  $app.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, $arguments)
} finally {
  if (-not $data.keepOpen) {
    if ($document) { $document.Close($false) }
    if ($app) { $app.Quit() }
  }
}`,{code:code,entryPoint:options.entryPoint||'Main',arguments:options.arguments,application:options.application||'Excel.Application',collection:options.collection||'Workbooks',visible:options.visible??Boolean(options.keepOpen),keepOpen:Boolean(options.keepOpen)},options);
}

async function openBridgeFolder(path){
  try{const result=await bridgeFileRequest('openDirectory',{path:path,activate:true},30000); if(result.foreground!==true)alert('The folder has been opened in File Explorer.\nPlease switch to it manually.\n\n'+path); return result;}catch(error){alert('Could not open the folder.\n\n'+
    bridgeError(error));throw error}
}

async function runNativeScriptFromEditor(){
  const button=bridgeId('nativeRun');button.disabled=true;
  try{
    const code=bridgeId('nativeScript').value,result=bridgeId('nativeLanguage').value==='vba'?await runBridgeVba(code,{entryPoint:bridgeId('nativeEntryPoint').value}):await runBridgePowerShell(code); showBridgeResult(result,result.ok?'Script completed.':'Script failed.');
      setBridgeBadge(result.ok?'Ready':'Action failed',result.ok?'ok':'error');
  }catch(error){showBridgeResult(bridgeError(error),'Script failed.');setBridgeBadge('Action failed','error')}
  finally{button.disabled=false}
}

function bridgeTransportError(){const error=new Error('Cannot reach SAPGUIcontrol.exe');error.transport=true;return error;}
function bridgeHttpRequest(method,url,payload,timeout){
  return new Promise((resolve,reject)=>{
    if(typeof GM_xmlhttpRequest!=='function'){reject(new Error('Tampermonkey permission for SAPGUIcontrol is unavailable. Reinstall or update this userscript.'));return}
    const headers={Accept:'application/json'};
    if(payload!==undefined)headers['Content-Type']='application/json';
    const options={method:method,url,headers:headers,timeout:timeout??20000,nocache:true,fetch:false,
      onload:response=>{
        if(!response||!Number(response.status)){reject(bridgeTransportError());return}
        let data={};
        try{data=response.responseText?JSON.parse(response.responseText):{}}catch(_){data={message:response.responseText||'Invalid SAPGUIcontrol response.'}}
        if(response.status>=200&&response.status<300){resolve(data);return}
        const error=new Error(data.message||('SAPGUIcontrol returned HTTP '+response.status+'.'));error.status=response.status;error.data=data; reject(error);
      },
      onerror:()=>reject(bridgeTransportError()), ontimeout:()=>reject(bridgeTransportError()), onabort:()=>reject(bridgeTransportError())
    };
    if(options.timeout===0)delete options.timeout;
    if(payload!==undefined)options.data=JSON.stringify(payload);
    try{GM_xmlhttpRequest(options)}catch(error){options.onerror(error&&error.message)}
  });
}

async function bridgeRequest(method,path,payload,timeout){
  if(typeof GM_xmlhttpRequest!=='function')throw new Error('Tampermonkey permission for SAPGUIcontrol is unavailable. Reinstall or update this userscript.');
  if(/^https?:\/\//i.test(path))return bridgeHttpRequest(method,path,payload,timeout);
  // Recheck identity before operations, including after an EXE restart. Concurrent callers share probes.
  if(!bridgeDiscovery)bridgeDiscovery=(async()=>{
    const candidates=bridgeUrl?[bridgeUrl,...BRIDGE_URLS.filter(url=>url!==bridgeUrl)]:BRIDGE_URLS;
    for(const url of candidates){
      try{const service=await bridgeHttpRequest('GET',url+'/',undefined,1500);if(service&&service.ok===true&&service.service==='SAPGUIcontrol'){bridgeUrl=url;return url}}catch(_){/* try the next candidate */}
    }
    bridgeUrl='';throw bridgeTransportError();
  })().finally(()=>{bridgeDiscovery=null});
  const url=await bridgeDiscovery;
  // Do not replay a failed operation: an Update may already have started in the EXE.
  return bridgeHttpRequest(method,new URL(path,url).href,payload,timeout);
}

function setBridgeBadge(text,kind){const badge=bridgeId('bridgeBadge');badge.textContent=text;badge.className='bridge-badge'+(kind?' '+kind:'');}
function showBridgeResult(value,label){const prefix=label?label+'\n':''; bridgeDebugLog=prefix+(typeof value==='string'?value:JSON.stringify(value,null,2));renderSapLog();}
function bridgeError(error){return error&&error.message?error.message:String(error||'Unknown SAPGUIcontrol error.');}
function normalizeBobjHeader(value){return String(value||'').trim().replace(/\s+/g,' ').toLocaleLowerCase();}
function bobjFidKey(value){const match=String(value||'').trim().match(/(?:^|\D)(\d{6})(?!\d)/);return match?match[1]:'';}
function bobjValue(fid,header){const row=bobjDataRows.get(bobjFidKey(fid)),index=bobjDataHeaderIndexes.get(normalizeBobjHeader(header)),value=row&&index!==undefined?String(row[index]??'').trim():'';return value==='#'?'':value;}

function bobjTemplateValue(row,header,omit){
  const key=normalizeBobjHeader(header),options=omit||{};
  if(key==='quote')return String(row&&row.quote||'');
  if(key==='today')return todayZbomStamp();
  if(options.customer&&key==='customer')return '';
  if(options.systemDescription&&key==='system description')return '';
  const source=bobjValue(row&&row.fid,header);if(source)return source;
  if(key==='fid')return String(row&&row.fid||'');
  if(key==='customer')return String(row&&row.customer||'');
  if(key==='system description')return String(row&&row.systemDescription||'');
  return '';
}

function renderBobjTemplate(template,row,omit){return String(template||'').replace(/\{([^{}]+)\}/g,(_token,header)=>bobjTemplateValue(row,header,omit));}

function missingBobjFileNameFields(rows,template){
  const headers=Array.from(new Set(Array.from(String(template||'').matchAll(/\{([^{}]+)\}/g),match=>match[1].trim()))) .filter(header=>!['fid','quote','today'].includes(normalizeBobjHeader(header)));
  return headers.filter(header=>rows.some(row=>{if(row._error)return false; const value=bobjTemplateValue(row,header);return !value||value==='#';}));
}

async function refreshBobjAndWait(){
  if(!bobjTaskRefreshPromise)bobjTaskRefreshPromise=(async()=>{
    bobjStatusRun++;clearTimeout(bobjPollTimer);bobjPollTimer=0;bobjRefreshing=true;updateBobjUpdate(); let status=await bridgeRequest('GET','/excel/status',undefined,10000);
    if(!status.running){try{status=await bridgeRequest('POST','/excel/refresh',{visible:false},20000)} catch(error){if(error.status!==409)throw error;status=await bridgeRequest('GET','/excel/status',undefined,10000)}}
    const deadline=Date.now()+660000;
    while(status.running){bobjStatus=status;setBobjRefreshState('Updating · '+String(status.stage||'running'),'running'); if(Date.now()>deadline)throw new Error('FID BOBJ refresh did not finish within 11 minutes.'); await waitMs(1000);status=await bridgeRequest('GET','/excel/status',
      undefined,10000);}
    bobjStatus=status;bobjRefreshing=false;
    if(status.stage!=='completed')throw new Error(status.message||'FID BOBJ update did not complete.');
    const source=await readBobjDataSource(status);
    if(!source||!source.available)throw new Error('The updated FID BOBJ workbook is unavailable.');
    if(bobjDataPromise)await bobjDataPromise;
    bobjDataReady=false; // Re-index once even if the filesystem timestamp has not changed.
    if(!await ensureBobjDataSource()||!bobjDataIsCurrent())throw new Error('The updated FID BOBJ Data could not be loaded.');
    setBobjRefreshState('Update completed · '+bobjDataCount+' FIDs indexed.','');
  })().finally(()=>{bobjRefreshing=false;bobjTaskRefreshPromise=null;updateBobjUpdate()});
  return bobjTaskRefreshPromise;
}

async function prepareBobjFileNames(tool,rows,template){
  await loadBobjRows(rows.map(row=>row.fid)); const missing=missingBobjFileNameFields(rows,template);
  if(!missing.length||sapTools[tool].bobjPrompted)return;
  sapTools[tool].bobjPrompted=true;
  if(!bobjTaskRefreshPromise&&!unsafeWindow.confirm('FID BOBJ Data may be outdated. Update it first?'))return;
  const progress=message=>{if(tool==='zbom')setZbomProgress(2,message);else setZbomCompareProgress(2,message)}; progress('Waiting for the FID BOBJ update before continuing…');
  try{await refreshBobjAndWait();await loadBobjRows(rows.map(row=>row.fid)); rows.forEach(enrichZbomCompareRowFromBobj); progress(missingBobjFileNameFields(rows,template).length?'FID BOBJ updated. Continuing with the available information…':'FID BOBJ updated. Continuing…');}
    catch(error){setBobjRefreshState('Update failed: '+bridgeError(error),'error'); progress('FID BOBJ update failed. Continuing with the available information…');}
}

function enrichZbomCompareRowFromBobj(row){if(!row||!row.fid)return row; if(!row.customer)row.customer=bobjValue(row.fid,'Customer'); if(!row.systemDescription)row.systemDescription=bobjValue(row.fid,'System Description'); return row;}

// Workers have separate globals; bootstrap them with the cached XLSX resource.
function createWorkbookWorker(code){
  const url=URL.createObjectURL(new Blob([GM_getResourceText('xlsxJs'),'\n;\n',code],{type:'text/javascript'})),worker=new Worker(url);
  URL.revokeObjectURL(url);return worker;
}
function bobjWorkerMain(){
  'use strict';
  let rows=new Map(); const normalizeHeader=value=>String(value||'').trim().replace(/\s+/g,' ').toLocaleLowerCase(); const sourceFid=value=>{const text=String(value||'').trim();return /^\d{6}$/.test(text)?text:''};
  const base64Bytes=value=>{const binary=atob(String(value||'')),bytes=new Uint8Array(binary.length);for(let index=0;index<binary.length;index++)bytes[index]=binary.charCodeAt(index);return bytes};
  const findTable=(workbook,firstMatch,ctl=false)=>{
    let best=null;
    (workbook&&workbook.SheetNames||[]).forEach((sheetName,sheetIndex)=>{
      const sheet=workbook.Sheets[sheetName],reference=sheet&&sheet['!ref'];if(!reference)return;
      const range=XLSX.utils.decode_range(reference),options={header:1,defval:'',raw:false,dateNF:'mm/dd/yyyy',blankrows:true},scanRange={s:{r:range.s.r,c:range.s.c},e:{r:Math.min(range.e.r,range.s.r+49),c:range.e.c}},preview=XLSX.utils.sheet_to_json(sheet,Object.assign({},options,
        {range:XLSX.utils.encode_range(scanRange)})),headerOffset=preview.findIndex(row=>ctl?['opportunity id','ctl id','change cat'].every(name=>(row||[]).some(value=>normalizeHeader(value)===name)):normalizeHeader(row&&row[0])==='fid');if(headerOffset<0)return;
      const dataRange={s:{r:range.s.r+headerOffset,c:range.s.c},e:{r:range.e.r,c:range.e.c}},matrix=XLSX.utils.sheet_to_json(sheet,Object.assign({},options,{range:XLSX.utils.encode_range(dataRange)}));
      if(firstMatch){best={matrix:matrix,sheet:sheetName,headerRow:range.s.r+headerOffset+1};return}
      const fidIndex=ctl?matrix[0].findIndex(value=>normalizeHeader(value)==='opportunity id'):0; const fidCount=matrix.reduce((count,row,index)=>count+(index&&sourceFid(row&&row[fidIndex])?1:0),0);if(!best||fidCount>best.fidCount)best={matrix:matrix,fidCount:fidCount,sheet:sheetName,
        headerRow:range.s.r+headerOffset+1};
    });
    return best;
  };
  self.onmessage=event=>{
    const message=event.data||{},reply=data=>self.postMessage(Object.assign({id:message.id,ok:true},data));
    try{
      if(message.type==='load'||message.type==='governance'||message.type==='governance-ctl'){
        const ctl=message.type==='governance-ctl',options={type:'array',cellDates:true,dense:true,sheets:[ctl?'CTL BOBJ':'BOBJ'],cellFormula:false,cellHTML:false,cellStyles:false},bytes=base64Bytes(message.contentBase64);
        let workbook=XLSX.read(bytes,options),table=findTable(workbook,true,ctl);if(!table){const fallback=Object.assign({},options);delete fallback.sheets;workbook=XLSX.read(bytes,fallback);table=findTable(workbook,false,ctl)}
        if(!table)throw new Error(ctl?'No worksheet with Opportunity ID, CTL ID and Change Cat was found.':'No worksheet with FID in its first column was found.');
        const headers=(table.matrix[0]||[]).map(value=>String(value||'').trim()),headerIndexes=new Map();headers.forEach((header,index)=>{const key=normalizeHeader(header);if(key&&!headerIndexes.has(key))headerIndexes.set(key,index)});
        if(message.type!=='load'){reply({headers:headers,values:table.matrix.slice(1),sheet:table.sheet,headerRow:table.headerRow});return}
        const nextRows=new Map();let blank=0,invalid=0,duplicates=0;
        for(let index=1;index<table.matrix.length;index++){const values=table.matrix[index],text=String(values&&values[0]||'').trim(),fid=sourceFid(text); if(!text){blank++;continue}if(!fid){invalid++;continue}if(nextRows.has(fid)){duplicates++;continue}nextRows.set(fid,values);}
        if(!nextRows.size)throw new Error('The FID BOBJ worksheet contains no readable FID rows.');
        rows=nextRows;reply({type:'loaded',count:rows.size,headerIndexes:Array.from(headerIndexes),ignored:blank+invalid,duplicates:duplicates});return;
      }
      if(message.type==='lookup'){const found=[];(message.fids||[]).forEach(fid=>{const row=rows.get(fid);if(row)found.push([fid,row])});reply({type:'lookup',rows:found});return;}
      throw new Error('Unknown FID BOBJ worker request.');
    }catch(error){self.postMessage({id:message.id,ok:false,error:error&&error.message?error.message:String(error)})}
  };
}

function createBobjWorker(){
  if(bobjWorker)return bobjWorker;
  if(typeof Worker==='undefined')throw new Error('This browser does not support background workbook reading.');
  const worker=createWorkbookWorker('('+bobjWorkerMain.toString()+')()');
  worker.onmessage=event=>{const data=event.data||{},pending=bobjWorkerPending.get(data.id);if(!pending)return;bobjWorkerPending.delete(data.id);if(data.ok)pending.resolve(data);else pending.reject(new Error(data.error||'FID BOBJ background worker failed.'))};
  worker.onerror=event=>{const error=new Error(event&&event.message||'FID BOBJ background worker failed.');bobjWorkerPending.forEach(pending=>pending.reject(error));bobjWorkerPending.clear();worker.terminate();if(bobjWorker===worker)bobjWorker=null}; bobjWorker=worker;return worker;
}

function callBobjWorker(type,payload){return new Promise((resolve,reject)=>{const id=++bobjWorkerRequest;bobjWorkerPending.set(id,{resolve:resolve,reject:reject});try{createBobjWorker().postMessage(Object.assign({id:id,type:type},payload||{}))}
  catch(error){bobjWorkerPending.delete(id);reject(error)}});}

function currentBobjFids(){
  const fids=new Set(),pattern=/(?:^|\D)(\d{6})(?!\d)/g; ['zbomInput','zbomCompareInput','createQuoteInput'].forEach(id=>{const input=bridgeId(id),text=String(input&&input.value||'');let match;while((match=pattern.exec(text)))fids.add(match[1]);pattern.lastIndex=0});
  return Array.from(fids);
}

function refreshBobjConsumers(){
  if(!bridgeId('zbomInput'))return;
  if(bridgeBusy){bobjConsumersPending=true;return}
  bobjConsumersPending=false;
  if(!sapToolRunning('zbom')&&!zbomHasRun)previewZbomInput();
  if(!sapToolRunning('zbomCompare')&&!zbomCompareHasRun)previewZbomCompareInput();
  const retainedCreate=sapSessionTask(sapTools.createQuote.selection);
  if(!sapToolRunning('createQuote')&&!(retainedCreate&&retainedCreate.tool==='createQuote'))previewCreateQuoteInput();
}

async function loadBobjDataSource(path,lastModified){
  const modified=Number(lastModified)||0;
  if(path===bobjDataPath&&modified===bobjDataModified&&bobjDataReady)return {count:bobjDataCount,changed:false};
  const run=++bobjDataLoadRun,data=await bridgeFileRequest('read',{path:path},BRIDGE_FILE_TIMEOUT);if(!data||!data.contentBase64)throw new Error('SAPGUIcontrol returned no FID BOBJ workbook content.');
  const loading=callBobjWorker('load',{contentBase64:data.contentBase64});data.contentBase64='';const loaded=await loading;
  if(run!==bobjDataLoadRun)return {count:bobjDataCount,changed:false};
  bobjDataRows=new Map();bobjDataHeaderIndexes=new Map(loaded.headerIndexes||[]);bobjDataMissing=new Set();bobjDataPath=path;bobjDataModified=modified;bobjDataCount=Number(loaded.count)||0;bobjDataReady=true;return {count:bobjDataCount,changed:true};
}

async function loadBobjRows(fids){
  if(!bobjDataIsCurrent())return false;
  const missing=Array.from(new Set((fids||[]).map(bobjFidKey).filter(fid=>fid&&!bobjDataRows.has(fid)&&!bobjDataMissing.has(fid))));if(!missing.length)return false;
  const path=bobjDataPath,modified=bobjDataModified,data=await callBobjWorker('lookup',{fids:missing});if(path!==bobjDataPath||modified!==bobjDataModified)return false;
  const found=new Set();(data.rows||[]).forEach(item=>{if(!Array.isArray(item)||!item[0])return;found.add(item[0]);bobjDataRows.set(item[0],item[1]||[])});missing.forEach(fid=>{if(!found.has(fid))bobjDataMissing.add(fid)});return Boolean(found.size);
}

function bobjDataIsCurrent(){return Boolean(bobjDataReady&&bobjSourcePath&&bobjDataPath===bobjSourcePath&&bobjDataModified===bobjSourceModified);}

async function ensureBobjDataSource(){
  const path=bobjSourcePath,modified=bobjSourceModified;if(!path)return false;
  if(!bobjDataIsCurrent()){
    if(!bobjDataPromise){
      bobjDataLoading=true;if(!bobjRefreshing)setBobjRefreshState('Indexing FID BOBJ Data in the background…','running');updateBobjUpdate();
      bobjDataPromise=loadBobjDataSource(path,modified).then(loaded=>{if(!bobjRefreshing&&path===bobjSourcePath&&modified===bobjSourceModified)setBobjRefreshState('Ready · '+loaded.count+' FIDs indexed.','');return true}).catch(error=>{bobjDataReady=false;
        if(!bobjRefreshing)setBobjRefreshState('Data source read failed: '+bridgeError(error),'error');return false}).finally(()=>{bobjDataLoading=false;bobjDataPromise=null;updateBobjUpdate()});
    }
    if(!await bobjDataPromise)return false;
  }
  try{await loadBobjRows(currentBobjFids());refreshBobjConsumers();return true}catch(error){if(!bobjRefreshing)setBobjRefreshState('Data source lookup failed: '+bridgeError(error),'error');return false}
}

function ensureBobjForCurrentInput(){if(currentBobjFids().length)ensureBobjDataSource();}

function formatBobjDataSourceTime(value){
  const date=new Date(Number(value));if(!Number.isFinite(date.getTime()))return '';
  const part=number=>String(number).padStart(2,'0'); return part(date.getMonth()+1)+'/'+part(date.getDate())+'/'+date.getFullYear()+' '+part(date.getHours())+':'+part(date.getMinutes())+':'+part(date.getSeconds());
}

function setBobjRefreshState(message,kind){const state=bridgeId('bobjRefreshState');if(!state)return; state.textContent=message;state.className='bridge-note bobj-refresh-state'+(kind?' '+kind:'');}

function updateBobjUpdate(){
  updateCtlBobjControls(); const button=bridgeId('bobjUpdate'),openButton=bridgeId('bobjOpen');if(!button||!openButton)return;
  button.disabled=bridgeBusy||bobjRefreshing||bobjOpening||bobjDataLoading||!bobjStatus||!bobjStatus.launcherFound;
  button.title=bobjRefreshing?'FID BOBJ refresh is running.':bobjDataLoading?'FID BOBJ Data is loading.':bobjStatus&&!bobjStatus.launcherFound?'SAP Analysis Launcher was not found.':'Refresh FID BOBJ Data in the background.';
  openButton.disabled=bridgeBusy||bobjRefreshing||bobjOpening||bobjDataLoading||!bobjStatus||!bobjStatus.hasWorkbook||!bobjStatus.outputPath;
  openButton.title=bobjRefreshing?'Wait for the FID BOBJ refresh to finish.':bobjDataLoading?'Wait for the FID BOBJ Data to finish loading.':'Open the current FID BOBJ workbook in Excel.';
}

async function readBobjDataSource(status){
  const source=bridgeId('bobjDataSource'),outputPath=String(status&&status.outputPath||'');if(!source)return;
  source.title=outputPath;
  if(!status||!status.hasWorkbook||!outputPath){bobjSourcePath='';bobjSourceModified=0;source.textContent='Current Data Source: Not available';return {available:false,count:0,loaded:false}}
  try{
    const data=await bridgeFileRequest('stat',{path:outputPath},10000),entry=Array.isArray(data.entries)?data.entries[0]:null,modified=entry&&formatBobjDataSourceTime(entry.lastModified);
    source.textContent=modified?'Current Data Source: '+modified:'Current Data Source: Not available';
    if(!entry){bobjSourcePath='';bobjSourceModified=0;return {available:false,count:0,loaded:false}}
    bobjSourcePath=outputPath;bobjSourceModified=Number(entry.lastModified)||0;setTimeout(ensureBobjForCurrentInput,250); return {available:true,count:bobjDataIsCurrent()?bobjDataCount:0,loaded:bobjDataIsCurrent()};
  }catch(error){source.textContent='Current Data Source: Unable to read';source.title=outputPath+'\n'+bridgeError(error);return {available:false,count:0,loaded:false,error:bridgeError(error)}}
}

async function refreshBobjStatus(){
  if(bobjTaskRefreshPromise){try{await bobjTaskRefreshPromise;return bobjStatus}catch(_){return null}}
  const run=++bobjStatusRun,wasRefreshing=bobjRefreshing;clearTimeout(bobjPollTimer);bobjPollTimer=0;
  try{
    const status=await bridgeRequest('GET','/excel/status',undefined,10000);if(run!==bobjStatusRun)return null;
    bobjStatus=status;bobjRefreshing=Boolean(status.running); const state=bridgeId('bobjRefreshState');if(state)state.title=String(status.message||'');
    if(bobjRefreshing){setBobjRefreshState('Updating · '+String(status.stage||'running'),'running'); if(!wasRefreshing)await readBobjDataSource(status); updateBobjUpdate();if(run===bobjStatusRun)bobjPollTimer=setTimeout(refreshBobjStatus,1000);}else{
      const sourceState=await readBobjDataSource(status),loaded=sourceState&&sourceState.loaded?' · '+sourceState.count+' FIDs indexed':' · Data loads when needed';
      if(status.stage==='failed')setBobjRefreshState('Update failed: '+String(status.message||'Unknown error'),'error');
      else if(!status.launcherFound)setBobjRefreshState('SAP Analysis Launcher not found.','error');
      else if(sourceState&&sourceState.error)setBobjRefreshState('Data source read failed: '+sourceState.error,'error');
      else if(sourceState&&!sourceState.available)setBobjRefreshState('Ready to update.',''); else setBobjRefreshState((status.stage==='completed'?'Update completed':'Ready')+loaded+'.','');
      updateBobjUpdate();
    }
    return status;
  }catch(error){if(run!==bobjStatusRun)return null; bobjStatus=null;bobjRefreshing=false;bridgeId('bobjDataSource').textContent='Current Data Source: Unavailable';setBobjRefreshState(bridgeError(error),'error');updateBobjUpdate();return null;}
}

async function startBobjUpdate(){
  if(bobjRefreshing||bobjTaskRefreshPromise)return;
  bobjStatusRun++;clearTimeout(bobjPollTimer);bobjPollTimer=0;bobjRefreshing=true;setBobjRefreshState('Starting background refresh…','running');updateBobjUpdate();
  try{bobjStatus=await bridgeRequest('POST','/excel/refresh',{visible:false},20000);bobjRefreshing=Boolean(bobjStatus.running);setBobjRefreshState('Updating · '+String(bobjStatus.stage||'queued'),'running');updateBobjUpdate();bobjPollTimer=setTimeout(refreshBobjStatus,500);}
    catch(error){bobjRefreshing=false;setBobjRefreshState('Update failed: '+bridgeError(error),'error');updateBobjUpdate()}
}

async function openBobjData(){
  const path=String(bobjStatus&&bobjStatus.outputPath||'');if(!path||!bobjStatus.hasWorkbook)return;
  bobjOpening=true;updateBobjUpdate();
  try{await bridgeFileRequest('openFile',{path:path},30000);setBobjRefreshState('Current data source opened in Excel.','')} catch(error){setBobjRefreshState('Open failed: '+bridgeError(error),'error')} finally{bobjOpening=false;updateBobjUpdate()}
}

function updateCtlBobjControls(){
  const update=bridgeId('ctlBobjUpdate'),open=bridgeId('ctlBobjOpen');if(!update||!open)return;
  const busy=bridgeBusy||ctlBobjBusy||Boolean(ctlBobjStatus&&ctlBobjStatus.running); update.disabled=busy||!ctlBobjStatus||!ctlBobjStatus.launcherFound; open.disabled=busy||!ctlBobjStatus||!ctlBobjStatus.hasWorkbook||!ctlBobjStatus.outputPath;
}
function ctlBobjMessage(message,kind='',details=message){const note=bridgeId('ctlBobjRefreshState');note.textContent=message;note.title=details;note.className='bridge-note bobj-refresh-state'+(kind?' '+kind:'');}
async function refreshCtlBobjStatus(){
  const run=++ctlBobjRun;clearTimeout(ctlBobjTimer);ctlBobjTimer=0;
  try{
    const status=await bridgeRequest('GET','/excel/ctl/status',undefined,10000);if(run!==ctlBobjRun)return;
    ctlBobjStatus=status; const source=bridgeId('ctlBobjDataSource');source.title=status.outputPath||'';
    if(status.hasWorkbook&&status.outputPath){const data=await bridgeFileRequest('stat',{path:status.outputPath},10000);if(run!==ctlBobjRun)return; const entry=Array.isArray(data.entries)?data.entries[0]:null; source.textContent=
      'Current Data Source: '+(entry?formatBobjDataSourceTime(entry.lastModified):'Not available');}else source.textContent='Current Data Source: Not available';
    if(status.running){ctlBobjMessage('Updating · '+(status.stage||'running'),'running',status.message);ctlBobjTimer=setTimeout(refreshCtlBobjStatus,1000)}
    else if(status.stage==='failed')ctlBobjMessage('Update failed','error',status.message);
    else if(!status.launcherFound)ctlBobjMessage('SAP Analysis Launcher not found.','error'); else ctlBobjMessage(status.stage==='completed'?'Update completed.':status.hasWorkbook?'Ready.':'Ready to update.');
  }catch(error){if(run!==ctlBobjRun)return; ctlBobjStatus=null;bridgeId('ctlBobjDataSource').textContent='Current Data Source: Unavailable';ctlBobjMessage('CTL BOBJ unavailable.','error',bridgeError(error));}
  updateCtlBobjControls();
}
async function startCtlBobjUpdate(){
  if(ctlBobjBusy||ctlBobjStatus?.running)return;
  ctlBobjRun++;clearTimeout(ctlBobjTimer);ctlBobjBusy=true;ctlBobjMessage('Starting background refresh…','running');updateCtlBobjControls();
  try{ctlBobjStatus=await bridgeRequest('POST','/excel/ctl/refresh',{visible:false},20000);await refreshCtlBobjStatus()} catch(error){ctlBobjMessage('Update failed','error',bridgeError(error))} finally{ctlBobjBusy=false;updateCtlBobjControls()}
}
async function openCtlBobjData(){
  const path=String(ctlBobjStatus?.outputPath||'');if(!path||!ctlBobjStatus.hasWorkbook)return;
  ctlBobjBusy=true;updateCtlBobjControls();
  try{await bridgeFileRequest('openFile',{path},30000);ctlBobjMessage('Current data source opened in Excel.')} catch(error){ctlBobjMessage('Open failed','error',bridgeError(error))} finally{ctlBobjBusy=false;updateCtlBobjControls()}
}

function sapToolRunning(tool){return Boolean(sapTools[tool].starting||sapTools[tool].job)}
function selectedBridgeSession(tool=activeSapTool){return bridgeSessions.find(session=>session.sessionKey===sapTools[tool].selection)||null;}
function sapSessionTask(key){return sharedSapTasks.find(task=>task.sessionKey===key)}
function canStartSapTool(tool){const session=selectedBridgeSession(tool); return bridgeSupportsTasks&&bridgeTaskOnline&&!bridgeBusy&&!sapToolRunning(tool)&&session&&!session.busy&&!sapSessionTask(session.sessionKey);}
function hasFreeSapSession(){return bridgeSessions.some(session=>!session.busy&&!sapSessionTask(session.sessionKey))}
function selectFreeSapSession(tool){if(sapToolRunning(tool))return; const free=bridgeSessions.filter(session=>!session.busy&&!sapSessionTask(session.sessionKey)); if(!free.some(session=>session.sessionKey===sapTools[tool].selection))sapTools[tool].selection=free[0]?.sessionKey||'';}

function setBridgeBusy(value){
  bridgeBusy=Boolean(value);
  document.querySelectorAll('[data-bridge-action]').forEach(button=>{const panel=button.closest('[data-tool-panel]'),tool=panel&&panel.dataset.toolPanel; button.disabled=bridgeBusy||Boolean(tool&&sapToolRunning(tool))||(button.id.indexOf('sap')===0&&!canStartSapTool('debug'));});
  document.querySelectorAll('[data-tool-panel] input,[data-tool-panel] textarea,[data-tool-panel] select').forEach(input=>{const tool=input.closest('[data-tool-panel]').dataset.toolPanel;input.disabled=sapToolRunning(tool);});
  document.querySelectorAll('[data-zbom-token],[data-zbom-compare-token],#zbomOpenHistory,#zbomResetInput,#zbomResetTemplate,#zbomCompareOpenHistory,#zbomCompareResetInput,#zbomCompareResetTemplate,#zbomCompareKeepTemp').forEach(button=>{button.disabled=
    sapToolRunning(button.closest('[data-tool-panel]').dataset.toolPanel);});
  renderSapSessionSelect(); updateZbomStart();updateZbomCompareStart();updateCreateQuoteStart();updateBobjUpdate();
}

function renderSapSessionSelect(){
  const select=bridgeId('sapSession');if(!select)return;
  const state=sapTools[activeSapTool],sessions=bridgeSessions.slice(); sharedSapTasks.forEach(task=>{if(task.session&&!sessions.some(session=>session.sessionKey===task.sessionKey))sessions.push({...task.session,missing:true})});
  if(!state.selection&&sessions.length){const first=sessions.find(session=>!session.missing&&!session.busy&&!sapSessionTask(session.sessionKey));if(first)state.selection=first.sessionKey;}
  const signature=JSON.stringify([state.selection,sessions,sharedSapTasks.map(task=>[task.taskId,task.state,task.ownerId])]);
  if(select.dataset.sessionOptions!==signature){
    select.replaceChildren();select.dataset.sessionOptions=signature;
    if(!sessions.some(session=>session.sessionKey===state.selection)){
      const option=document.createElement('option');option.value=state.selection;option.textContent=state.selection?'Selected SAP window is no longer available':sessions.length?'No free SAP sessions — open another SAP window and refresh':'No SAP GUI sessions found';
        option.disabled=true;select.appendChild(option);
    }
    sessions.forEach(session=>{
      const task=sapSessionTask(session.sessionKey),option=document.createElement('option');option.value=session.sessionKey;
      const status=task?(SAP_TOOL_LABELS[task.tool]||task.tool)+' · '+(task.state==='running'?'Running':task.state==='held'?'Retained':'Interrupted')+(task.ownerId!==bridgePageId?' · Other K2 page':''):session.missing?'Unavailable':session.busy?'SAP busy':'Idle';
      option.textContent=(task||session.busy||session.missing?'🔴 ':'🟢 ')+'Window '+(session.connectionIndex+1)+'.'+(session.sessionIndex+1)+' · '+(session.system||'SAP')+' '+(session.client||'')+' · '+(session.user||'')+' · '+(session.transaction||
        'No transaction')+' · '+(session.windowTitle||'')+' — '+status;
      option.disabled=Boolean(task||session.busy||session.missing);if(option.disabled)option.style.color='#94a3b8';
      select.appendChild(option);
    });
  }
  select.value=state.selection;select.disabled=sapToolRunning(activeSapTool)||!sessions.length; bridgeId('sapSessionLabel').textContent='SAP session · '+SAP_TOOL_LABELS[activeSapTool];
  const task=sapSessionTask(state.selection),note=bridgeId('sapSessionState'),release=bridgeId('sapReleaseSession');
  note.textContent=!bridgeSupportsTasks?'Update and restart SAPGUIcontrol to use independent sessions.':!bridgeTaskOnline?'Task status unavailable. Reconnecting…':task?(SAP_TOOL_LABELS[task.tool]||task.tool)+(task.state==='running'?' is running in this window.':task.state===
    'held'?' retained this window. Finish or discard its SAP work before releasing.':' was interrupted. Check SAP before releasing this window.'):'This selection belongs to '+SAP_TOOL_LABELS[activeSapTool]+'. Check the SAP window before starting.';
  if(task&&task.releaseInSeconds!=null)note.textContent='Quote Create retained this window. Automatically available to K2 again in '+task.releaseInSeconds+' seconds.';
  if(bridgeSessionRefreshPromise&&!hasFreeSapSession())note.textContent='Refreshing SAP windows…';
  const retained=sharedSapTasks.filter(item=>item.tool===activeSapTool&&item.state!=='running'); const retainedTask=retained.find(item=>item.releaseInSeconds!=null);
  if(!task&&retainedTask&&retainedTask.releaseInSeconds!=null)note.textContent+=' Quote Create window releases in '+retainedTask.releaseInSeconds+' seconds.';
  release.hidden=!retained.length;release.disabled=!retained.some(item=>!item.inFlight)||!bridgeTaskOnline;
  document.querySelectorAll('[data-tool-tab]').forEach(button=>{const tool=button.dataset.toolTab;button.textContent=SAP_TOOL_LABELS[tool]+(sapToolRunning(tool)?' · Running':'')});
}

function renderBridgeSessions(sessions){bridgeSessions=Array.isArray(sessions)?sessions:[];setBridgeBusy(bridgeBusy);}
function acceptSapTaskStatus(data){sharedSapTasks=Array.isArray(data.tasks)?data.tasks:[];bridgeTaskOnline=true;setBridgeBusy(bridgeBusy);}

async function pollSapTasks(){
  if(bridgeTaskPolling||!bridgeSupportsTasks)return;
  bridgeTaskPolling=true;clearTimeout(bridgeTaskPollTimer);
  try{acceptSapTaskStatus(await bridgeRequest('POST','/tasks/heartbeat',{ownerId:bridgePageId},10000))} catch(_){bridgeTaskOnline=false;setBridgeBusy(bridgeBusy)} finally{bridgeTaskPolling=false;bridgeTaskPollTimer=setTimeout(pollSapTasks,3000)}
}

async function refreshSapSessionChoices(){
  if(!bridgeSupportsTasks)return;
  if(hasFreeSapSession()){selectFreeSapSession(activeSapTool);setBridgeBusy(bridgeBusy);return;}
  if(!bridgeSessionRefreshPromise){
    bridgeSessionRefreshPromise=(async()=>{
      const status=await bridgeRequest('GET','/tasks',undefined,10000); sharedSapTasks=status.tasks||[];bridgeTaskOnline=true;selectFreeSapSession(activeSapTool);setBridgeBusy(bridgeBusy);
      if(hasFreeSapSession())return;
      try{const health=await bridgeRequest('GET','/health',undefined,10000); bridgeSupportsTasks=health.sessionTasks===true;bridgeSessions=health.sessions||[]; selectFreeSapSession(activeSapTool);}catch(_){/* Keep the last known windows if SAP is busy. */}
    })().catch(()=>{bridgeTaskOnline=false}).finally(()=>{bridgeSessionRefreshPromise=null;setBridgeBusy(bridgeBusy)});
    setBridgeBusy(bridgeBusy);
  }
  await bridgeSessionRefreshPromise;
}

async function acquireSapTask(tool,options={}){
  if(!bridgeSupportsTasks||!bridgeTaskOnline)throw new Error('Update/restart SAPGUIcontrol and refresh the connection first.');
  const session=selectedBridgeSession(tool);if(!session)throw new Error('Select an available SAP GUI session for '+SAP_TOOL_LABELS[tool]+'.');
  const job={taskId:crypto.randomUUID(),ownerId:bridgePageId,sessionKey:session.sessionKey,keep:true}; sapTools[tool].job=job;
  try{acceptSapTaskStatus(await bridgeRequest('POST','/tasks/acquire',{...job,tool,...options},15000))} catch(error){if(error.status)sapTools[tool].job=null;throw error}
  return job;
}

async function runSapTool(tool,run){
  if(sapToolRunning(tool))return;
  sapTools[tool].starting=true;sapTools[tool].bobjPrompted=false;setBridgeBusy(bridgeBusy);
  try{await run()}
  catch(error){const message=bridgeError(error); if(tool==='zbom')setZbomProgress(0,message); else if(tool==='zbomCompare')setZbomCompareProgress(0,message); else if(tool==='createQuote')setCreateQuoteStatus(message,'error'); else showBridgeResult(message,'SAP action failed.');}
    finally{
    const job=sapTools[tool].job;
    if(job)try{acceptSapTaskStatus(await bridgeRequest('POST','/tasks/finish',{taskId:job.taskId,ownerId:bridgePageId,keep:job.keep,autoRelease:job.autoRelease===true},10000))} catch(_){bridgeTaskOnline=false}
    sapTools[tool].job=null;sapTools[tool].starting=false;setBridgeBusy(bridgeBusy);if(bobjConsumersPending)refreshBobjConsumers();pollSapTasks();
  }
}

async function releaseSapSession(){
  const retained=sharedSapTasks.filter(item=>item.tool===activeSapTool&&item.state!=='running'&&!item.inFlight);
  try{for(const task of retained)acceptSapTaskStatus(await bridgeRequest('POST','/tasks/release',{taskId:task.taskId},10000)); await refreshSapSessionChoices();} catch(error){bridgeId('sapSessionState').textContent=bridgeError(error)}
}

function renderSapLog(){
  const result=bridgeId('sapResult');if(!result)return;
  result.textContent=activeSapTool==='zbom'?'ZBOM PDF log\n'+zbomLogLines.join('\n'):activeSapTool==='zbomCompare'?'ZBOM Compare log\n'+zbomCompareLogLines.join('\n'):activeSapTool==='debug'?bridgeDebugLog:''; result.scrollTop=result.scrollHeight;
}

function isSapGuiControlUpdateAvailable(current,latest){const local=current.replace(/^v/i,'').split('.'),remote=latest.replace(/^v/i,'').split('.'); for(let index=0;index<Math.max(local.length,remote.length);index++){const difference=Number(remote[index]||0)-Number(local[index]||0);
  if(difference)return difference>0;} return false;}

function renderSapGuiControlUpdate(){
  const state=sapGuiControlUpdate,release=state.release,download=bridgeId('bridgeDownload'); const newer=Boolean(release&&state.current&&isSapGuiControlUpdateAvailable(state.current,release.version)); download.href=release&&release.downloadUrl||SAPGUICONTROL_DOWNLOAD_URL;
  download.hidden=state.online===null||Boolean(state.online&&!newer); download.textContent=state.online?'Update SAPGUIcontrol':'Download SAPGUIcontrol';
  download.title=(state.online&&release?'Download SAPGUIcontrol '+release.version+' (ZIP).':'Download SAPGUIcontrol.zip.')+' Extract it and run SAPGUIcontrol.exe.';
  bridgeId('bridgeUpdateStatus').textContent=state.checking?'Checking for updates…':release?(state.online&&!newer?'Up to date':'Latest version: '+release.version):'Update check unavailable';
}

async function checkSapGuiControlUpdate(){
  sapGuiControlUpdate.checking=true;renderSapGuiControlUpdate();
  try{const release=await bridgeRequest('GET',SAPGUICONTROL_VERSION_URL+'?t='+Date.now(),undefined,8000); sapGuiControlUpdate.release={version:release.version.trim(),downloadUrl:release.downloadUrl};}catch(_){sapGuiControlUpdate.release=null}
  finally{sapGuiControlUpdate.checking=false;renderSapGuiControlUpdate()}
}

async function refreshBridge(background=false){
  if(bridgeRefreshPending)return;
  bridgeRefreshPending=true;clearTimeout(bridgeReconnectTimer);bridgeReconnectTimer=0;
  if(background!==true){setBridgeBusy(true);setBridgeBadge('Checking','');bridgeId('bridgeService').textContent='Checking SAPGUIcontrol…';sapGuiControlUpdate.online=null;checkSapGuiControlUpdate()}
  try{
    const health=await bridgeRequest('GET','/health',undefined,background===true?3000:8000),mock=health.mode==='mock';
    if(background===true)setBridgeBusy(true);
    bridgeMock=mock;bridgeSupportsTasks=health.sessionTasks===true;await pollSapTasks(); bridgeId('bridgeService').textContent=(health.service||'SAPGUIcontrol')+' '+(health.version||'')+(mock?' · MOCK':''); sapGuiControlUpdate.online=true;sapGuiControlUpdate.current=health.version||
      '';renderSapGuiControlUpdate(); bridgeId('bridgeSap').textContent=mock?'Simulated SAP session':health.sapAvailable?'SAP GUI detected':'SAP GUI not detected'; renderBridgeSessions(health.sessions); selectFreeSapSession(activeSapTool);setBridgeBusy(bridgeBusy);
    if(bridgeSessions.length)setBridgeBadge(mock?'Mock ready':'Ready',mock?'warn':'ok'); else setBridgeBadge('SAP unavailable','warn');
    showBridgeResult(health,'Connection refreshed.'); await Promise.all([refreshBobjStatus(),refreshCtlBobjStatus()]);
  }catch(error){
    if(background===true)return;
    bridgeMock=false;bridgeTaskOnline=false; renderBridgeSessions([]); bridgeId('bridgeService').textContent=error.status?'SAPGUIcontrol responded with an error':'Offline';bridgeId('bridgeSap').textContent='Unavailable';
      setBridgeBadge(error.status?'SAPGUIcontrol error':'SAPGUIcontrol offline','error');sapGuiControlUpdate.online=false; sapGuiControlUpdate.current='';renderSapGuiControlUpdate(); bobjStatusRun++;clearTimeout(bobjPollTimer);bobjPollTimer=0;bobjStatus=null;bobjRefreshing=false;
      bridgeId('bobjDataSource').textContent='Current Data Source: Unavailable';setBobjRefreshState('SAPGUIcontrol unavailable.','error');updateBobjUpdate(); ctlBobjRun++;clearTimeout(ctlBobjTimer);ctlBobjTimer=0;ctlBobjStatus=null;bridgeId('ctlBobjDataSource').textContent=
      'Current Data Source: Unavailable';ctlBobjMessage('SAPGUIcontrol unavailable.','error');updateCtlBobjControls(); showBridgeResult(bridgeError(error),'Connection failed.');
  }finally{bridgeRefreshPending=false;setBridgeBusy(false);if(bobjConsumersPending)refreshBobjConsumers(); if(sapGuiControlUpdate.online===false)bridgeReconnectTimer=setTimeout(()=>refreshBridge(true),10000);}
}

function setToolTab(name){
  if(name==='debug'&&!bridgeDebugUnlocked)name='zbom';
  activeSapTool=name; document.querySelectorAll('[data-tool-tab]').forEach(button=>{const active=button.dataset.toolTab===name;button.setAttribute('aria-selected',active?'true':'false');});
  document.querySelectorAll('[data-tool-panel]').forEach(panel=>{panel.hidden=panel.dataset.toolPanel!==name}); setBridgeBusy(bridgeBusy);renderSapLog();refreshSapSessionChoices();
}

function splitZbomLine(line){
  if(line.indexOf('\t')>=0)return line.split('\t').map(value=>value.trim());
  if(line.indexOf(',')>=0){
    const values=[];let value='',quoted=false;
    for(let index=0;index<line.length;index++){const character=line[index]; if(character==='"'){if(quoted&&line[index+1]==='"'){value+='"';index++}else quoted=!quoted;}else if(character===','&&!quoted){values.push(value.trim());value=''} else value+=character;}
    values.push(value.trim());return values;
  }
  if(/\s{2,}/.test(line.trim()))return line.trim().split(/\s{2,}/).map(value=>value.trim());
  return [line.trim()];
}

function normalizedZbomHeader(value){return String(value||'').trim().toLowerCase().replace(/[\s_.\-/#]+/g,'');}

function zbomHeaderRole(value){
  const key=normalizedZbomHeader(value);
  if(['fid','facilityid'].includes(key))return 'fid';
  if(['systemdescription','systemdesc','sysdescription','sysdesc','description'].includes(key))return 'systemDescription';
  if(['quote','quotation','quotenumber','quotationnumber','vbeln'].includes(key))return 'quote';
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
  const parsed=Date.parse(text);return Number.isFinite(parsed)?parsed:null;
}

function parseCreateQuoteInput(text){
  const source=String(text||'').replace(/\r/g,'').split('\n').map((line,index)=>{const values=splitZbomLine(line);return {line:index+1,values:values.length===1&&/\s/.test(line.trim())?line.trim().split(/\s+/).map(value=>value.trim()):values};}).filter(item=>item.values.some(Boolean));
  if(!source.length)return {rows:[],duplicates:0,ignored:0};
  const hasHeader=source[0].values.some(value=>Boolean(createQuoteHeaderRole(value))),data=hasHeader?source.slice(1):source,rows=[],seen=new Map();let duplicates=0,ignored=0;
  data.forEach(item=>{
    const cells=item.values.map((value,index)=>({value:String(value||'').trim(),index:index})).filter(cell=>cell.value),fidCell=cells.find(cell=>/^\d{6}$/.test(cell.value));
    if(!fidCell){ignored++;return}
    const dates=cells.filter(cell=>cell.index!==fidCell.index&&createQuoteDateValue(cell.value)!==null); const row={fid:fidCell.value,crd:dates[0]?dates[0].value:'',validTo:dates[1]?dates[1].value:''},existing=seen.get(fidCell.value);
    if(existing){duplicates++;if(!existing.crd&&row.crd)existing.crd=row.crd;if(!existing.validTo&&row.validTo)existing.validTo=row.validTo;return}
    seen.set(fidCell.value,row);rows.push(row);
  });
  return {rows:rows,duplicates:duplicates,ignored:ignored};
}

function formatCreateQuoteDate(date){return String(date.getMonth()+1).padStart(2,'0')+'/'+String(date.getDate()).padStart(2,'0')+'/'+date.getFullYear();}
function formatCreateQuoteDateParts(year,month,day){return createQuoteUtcDate(year,month,day)===null?'':String(Number(month)).padStart(2,'0')+'/'+String(Number(day)).padStart(2,'0')+'/'+String(Number(year)).padStart(4,'0');}

function normalizeCreateQuoteDate(value){
  const text=String(value||'').trim();if(!text)return '';
  let match=text.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if(match)return formatCreateQuoteDateParts(match[1],match[2],match[3]);
  match=text.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if(match)return formatCreateQuoteDateParts(match[1],match[2],match[3]);
  match=text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if(match)return formatCreateQuoteDateParts(match[1],match[2],match[3]);
  match=text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if(match){const first=Number(match[1]),second=Number(match[2]),month=first>12?second:first,day=first>12?first:second; return formatCreateQuoteDateParts(match[3],month,day);}
  if(/^\d{4,5}(?:\.\d+)?$/.test(text)){const serial=Number(text); if(serial>=20000&&serial<=80000){const date=new Date(Date.UTC(1899,11,30)+Math.round(serial*86400000));return formatCreateQuoteDateParts(date.getUTCFullYear(),date.getUTCMonth()+1,date.getUTCDate())}}
  const parsed=Date.parse(text);return Number.isFinite(parsed)?formatCreateQuoteDate(new Date(parsed)):'';
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
  bridgeId('createQuoteCount').textContent=parts.join(' · '); updateCreateQuoteStart();
}

function previewCreateQuoteInput(){
  const parsed=parseCreateQuoteInput(bridgeId('createQuoteInput').value);createQuoteRows=parsed.rows;createQuoteRows.forEach(row=>{
    const override=createQuoteCrdOverrides.get(row.fid);
    if(createQuoteManualCrd.has(row.fid))row.crd=override||''; else if(String(override||'').trim())row.crd=override; else if(!row.crd){const source=bobjValue(row.fid,'Date: Customer Request');if(source&&source!=='#')row.crd=normalizeCreateQuoteDate(source)||source}
  });createQuoteStats={duplicates:parsed.duplicates,ignored:parsed.ignored};renderCreateQuotePreview();
}

function setCreateQuoteStatus(message,kind){const status=bridgeId('createQuoteStatus');if(!status)return; status.textContent=message;status.className='bridge-note create-quote-status'+(kind==='error'?' row-error':kind==='ok'?' row-ok':'');}
function updateCreateQuoteStart(){const button=bridgeId('createQuoteStart');if(!button)return; button.disabled=!canStartSapTool('createQuote')||!sanitizeCreateQuoteQuotation()||!createQuoteRows.length;}

function validateCreateQuotePreview(){
  const inputs=Array.from(document.querySelectorAll('#createQuotePreviewBody .create-quote-crd-input')),normalized=[];let firstInvalid=null,error='';
  createQuoteRows.forEach((row,index)=>{
    const input=inputs[index],crdText=String(input?input.value:row.crd).trim(),crd=crdText?normalizeCreateQuoteDate(crdText):'',validTo=normalizeCreateQuoteDate(row.validTo||defaultCreateQuoteValidTo());
    if(input)input.classList.remove('invalid');
    if(crdText&&!crd){if(input){input.classList.add('invalid');if(!firstInvalid)firstInvalid=input} if(!error)error='Item '+(index+1)+': CRD is not a recognized date.'; return;}
    if(!validTo){if(!error)error='Item '+(index+1)+': Valid to is not a recognized date.';return}
    normalized[index]={crd:crd,validTo:validTo};
  });
  if(error){setCreateQuoteStatus(error,'error');if(firstInvalid)firstInvalid.focus();return false}
  normalized.forEach((value,index)=>{createQuoteRows[index].crd=value.crd;createQuoteRows[index].validTo=value.validTo;createQuoteCrdOverrides.set(createQuoteRows[index].fid,value.crd)}); saveCreateQuoteCrdOverrides();renderCreateQuotePreview();return true;
}

async function startCreateQuote(){
  await ensureBobjDataSource();previewCreateQuoteInput();
  if(!validateCreateQuotePreview())return;
  recordCreateQuoteHistory(sanitizeCreateQuoteQuotation(),bridgeId('createQuoteInput').value); await acquireSapTask('createQuote');setBridgeBusy(bridgeBusy);setCreateQuoteStatus('Opening SAP transaction ZVA21M…');
  try{
    await sendZbomAction('createQuote','transaction',{code:'ZVA21M'}); await sendZbomAction('createQuote','setText',{controlId:'wnd[0]/usr/txtGS_HEADER-VBELN',value:sanitizeCreateQuoteQuotation()});
    await sendZbomAction('createQuote','setKey',{controlId:'wnd[0]/usr/cmbGS_HEADER-ACQ_ID',value:'999'}); await sendZbomAction('createQuote','sendVKey',{controlId:'wnd[0]',key:0});
    const data=await sendZbomAction('createQuote','readText',{controlId:'wnd[0]/usr/txtGS_HEADER-VERSION'}),referenceVersion=String(data&&data.value||'').trim(),fidMatch=referenceVersion.match(/(?:^|\D)(\d{6})(?!\d)/),referenceFid=fidMatch?fidMatch[1]:'';
    if(!referenceFid){if(!unsafeWindow.confirm('FID could not be extracted from "'+(referenceVersion||'(blank)')+'". Continue?')){setCreateQuoteStatus('Create Quote stopped.');return}}else{
      const row=createQuoteRows.find(item=>item.fid===referenceFid);
      if(row){
        const tail=referenceVersion.slice(referenceVersion.indexOf(referenceFid)+referenceFid.length),versionMatch=tail.match(/^(\D*)(\d+)(.*)$/);
        if(versionMatch){
          const crdData=await sendZbomAction('createQuote','readText',{controlId:'wnd[0]/usr/txtGS_HEADER-CRD'}),referenceCrd=String(crdData&&crdData.value||'').trim();
          row.version=referenceFid+versionMatch[1]+String(Number(versionMatch[2])+1).padStart(versionMatch[2].length,'0')+versionMatch[3];row.crd=referenceCrd;createQuoteCrdOverrides.set(row.fid,row.crd);saveCreateQuoteCrdOverrides();renderCreateQuotePreview();
        }else if(!unsafeWindow.confirm('FID Version "'+referenceVersion+'" could not be incremented. Continue?')){setCreateQuoteStatus('Create Quote stopped.');return}
      }
    }
    setCreateQuoteStatus('Entering '+createQuoteRows.length+' item(s) into SAP…'); const tableId='wnd[0]/usr/tblZSDR0390_MASS_QUOTE_CREATE1MYTABLE';
    for(let index=0;index<createQuoteRows.length;index++){
      const row=createQuoteRows[index];
      if(row.crd)await sendZbomAction('createQuote','setText',{controlId:tableId+'/ctxtGS_ITEM-REQ_DEL_DATE[1,'+index+']',value:row.crd});
      await sendZbomAction('createQuote','setText',{controlId:tableId+'/txtGS_ITEM-VERSION[2,'+index+']',value:row.version||createQuotePreviewFid(row.fid)});
      await sendZbomAction('createQuote','setText',{controlId:tableId+'/ctxtGS_ITEM-VALID_TO[3,'+index+']',value:row.validTo||defaultCreateQuoteValidTo()});
    }
    if(createQuoteRows.every(row=>String(row.crd||'').trim())){await sendZbomAction('createQuote','sapCall',{controlId:tableId+'/ctxtGS_ITEM-VALID_TO[3,0]',steps:[{action:'invoke',member:'SetFocus',arguments:[]}]}); await sendZbomAction('createQuote','sendVKey',{controlId:'wnd[0]',
      key:0});}
    sapTools.createQuote.job.autoRelease=true; setCreateQuoteStatus(createQuoteRows.length+' item(s) entered. Ready for the next step. Session releases automatically after 1 minute.','ok');
  }catch(error){setCreateQuoteStatus('Unable to prepare ZVA21M: '+bridgeError(error),'error')}
  finally{setBridgeBusy(bridgeBusy)}
}

function zbomCompareHeaderRole(value){
  const key=normalizedZbomHeader(value);
  if(['quote','quotation','quoteno','quotationno','quotenumber','quotationnumber','vbeln'].includes(key))return 'quote';
  if(['fid','facilityid'].includes(key))return 'fid';
  if(['group','groupid','comparegroup','comparisongroup','zbomgroup'].includes(key))return 'group';
  if(['systemdescription','systemdesc','sysdescription','sysdesc','description'].includes(key))return 'systemDescription';
  if(['customer','customername','account','accountname'].includes(key))return 'customer';
  return '';
}

function extractZbomCompareQuote(value){const text=String(value||'').trim(); if(/^2\d{7}$/.test(text))return text; const match=text.match(/(?:^|\D)(2\d{7})(?!\d)/); return match?match[1]:'';}
function normalizeZbomCompareGroup(value,allowCustom){const text=String(value||'').trim(); if(!text)return ''; if(/^standalone$/i.test(text)||/^do[\s_-]*not[\s_-]*compare$/i.test(text))return ZBOM_COMPARE_STANDALONE; if(/^g\d+$/i.test(text))return 'G'+text.slice(1);
  return allowCustom&&/^[A-Za-z0-9]+$/.test(text)?text:'';}
function mappedZbomCompareValue(values,index){return index>=0&&index<values.length?String(values[index]||'').trim():'';}

function inferZbomCompareColumns(rows,mapping){
  if(mapping.systemDescription>=0)return mapping;
  const width=Math.max(0,...rows.map(values=>values.length)),reserved=new Set(Object.values(mapping).filter(index=>index>=0));let best=-1,bestMatches=0;
  for(let column=0;column<width;column++){if(reserved.has(column))continue; const matches=rows.reduce((total,values)=>total+(/[\/_]/.test(String(values[column]||'').trim())?1:0),0); if(matches>bestMatches){best=column;bestMatches=matches}}
  if(best>=0)mapping.systemDescription=best;
  return mapping;
}

function classifyZbomCompareRow(values,mapping){
  const cells=values.map((value,index)=>({value:String(value||'').trim(),index:index})).filter(cell=>cell.value),used=new Set(); Object.values(mapping).forEach(index=>{if(index>=0)used.add(index)});
  let quote=extractZbomCompareQuote(mappedZbomCompareValue(values,mapping.quote)),quoteIndex=quote?mapping.quote:-1;
  if(!quote){
    const candidates=cells.filter(cell=>!used.has(cell.index)).map(cell=>Object.assign({},cell,{quote:extractZbomCompareQuote(cell.value)})).filter(cell=>cell.quote); const candidate=candidates.find(cell=>cell.value===cell.quote)||candidates[0];
    if(candidate){quote=candidate.quote;quoteIndex=candidate.index;used.add(candidate.index)}
  }
  let fid=mappedZbomCompareValue(values,mapping.fid);
  if(!/^\d{6}$/.test(fid))fid='';
  if(!fid){const candidate=cells.find(cell=>!used.has(cell.index)&&/^\d{6}$/.test(cell.value)); if(candidate){fid=candidate.value;used.add(candidate.index)}}
  let group=normalizeZbomCompareGroup(mappedZbomCompareValue(values,mapping.group),mapping.group>=0);
  if(!group){const candidate=cells.find(cell=>!used.has(cell.index)&&/^g\d+$/i.test(cell.value)); if(candidate){group=normalizeZbomCompareGroup(candidate.value,false);used.add(candidate.index)}}
  let customer=mappedZbomCompareValue(values,mapping.customer),systemDescription=mappedZbomCompareValue(values,mapping.systemDescription); const unknown=cells.filter(cell=>!used.has(cell.index));
  if(!systemDescription){const marked=unknown.find(cell=>/[\/_]/.test(cell.value)); if(marked){systemDescription=marked.value;used.add(marked.index)}}
  const remaining=unknown.filter(cell=>!used.has(cell.index));
  if(!customer&&!systemDescription){if(remaining.length===1)systemDescription=remaining[0].value; else if(remaining.length>1){customer=remaining[0].value;systemDescription=remaining[1].value}}else if(!customer&&remaining.length)customer=remaining[0].value; else if(!systemDescription&&
    remaining.length)systemDescription=remaining[0].value;
  return {quote:quote,fid:fid,group:group,customer:customer,systemDescription:systemDescription,_groupExplicit:Boolean(group),_quoteIndex:quoteIndex};
}

function parseZbomCompareInput(text){
  const source=String(text||'').replace(/\r/g,'').split('\n').map((line,index)=>({line:index+1,values:splitZbomLine(line)})).filter(item=>item.values.some(Boolean));
  if(!source.length)return {rows:[],duplicates:0,ignored:0};
  const firstRoles=source[0].values.map(zbomCompareHeaderRole),hasHeader=firstRoles.includes('quote'),mapping={quote:-1,fid:-1,group:-1,customer:-1,systemDescription:-1};
  if(hasHeader)firstRoles.forEach((role,index)=>{if(role&&mapping[role]<0)mapping[role]=index});
  const rows=[],seen=new Map(),data=hasHeader?source.slice(1):source;let duplicates=0,ignored=0; inferZbomCompareColumns(data.map(item=>item.values),mapping);
  data.forEach(item=>{
    const row=classifyZbomCompareRow(item.values,mapping);
    if(!row.quote){ignored++;return}
    if(!row.group)row.group=zbomCompareSavedGroups.get(row.quote)||ZBOM_COMPARE_DEFAULT_GROUP;
    const existing=seen.get(row.quote);
    if(existing){duplicates++; ['fid','customer','systemDescription'].forEach(field=>{if(!existing[field]&&row[field])existing[field]=row[field]}); if(!existing._groupExplicit&&row._groupExplicit){existing.group=row.group;existing._groupExplicit=true} return;}
    seen.set(row.quote,row);rows.push(row);
  });
  return {rows:rows,duplicates:duplicates,ignored:ignored};
}

function zbomCompareGroupLabel(number){return 'G'+String(number).padStart(2,'0');}

function zbomCompareGroupOptions(){
  const selected=new Set(zbomCompareRows.map(row=>row.group)),groups=[zbomCompareGroupLabel(1)],custom=[],customKeys=new Set();let number=1;
  while(number<500&&selected.has(zbomCompareGroupLabel(number))){number++;groups.push(zbomCompareGroupLabel(number))}
  zbomCompareRows.forEach(row=>{if(row.group&&row.group!==ZBOM_COMPARE_STANDALONE&&!groups.includes(row.group)){const key=row.group;if(!customKeys.has(key)){customKeys.add(key);custom.push(row.group)}}}); return groups.concat(custom,ZBOM_COMPARE_STANDALONE);
}

function compactZbomCompareName(value){return String(value||'').trim().replace(/_{2,}/g,'_').replace(/^[\s_.-]+|[\s_.-]+$/g,'');}
function renderZbomCompareFileName(row,omit){const input=bridgeId('zbomCompareTemplate'),template=String(input.value||'').trim()||ZBOM_COMPARE_DEFAULT_TEMPLATE; return compactZbomCompareName(renderBobjTemplate(template,Object.assign({},row,{fid:row.fid||'{FID}'}),omit));}
function finishZbomCompareFileName(value){const safe=compactZbomCompareName(value).replace(/[<>:"/\\|?*\u0000-\u001F]/g,',').replace(/\.xlsx$/i,''); return (safe||'ZBOM Compare')+'.xlsx';}
function zbomCompareDuplicateKey(value){return String(value||'').trim().toLocaleLowerCase()}

function renderZbomCompareGroupFileName(rows){
  const separator=String(bridgeId('zbomCompareSeparator').value||'').trim()||ZBOM_COMPARE_DEFAULT_SEPARATOR; const seenCustomers=new Set(),seenDescriptions=new Set(),parts=[];
  rows.forEach(row=>{
    const customerKey=zbomCompareDuplicateKey(bobjTemplateValue(row,'Customer')),descriptionKey=zbomCompareDuplicateKey(bobjTemplateValue(row,'System Description'));
    parts.push(renderZbomCompareFileName(row,{customer:Boolean(customerKey&&seenCustomers.has(customerKey)),systemDescription:Boolean(descriptionKey&&seenDescriptions.has(descriptionKey))}));
    if(customerKey)seenCustomers.add(customerKey);
    if(descriptionKey)seenDescriptions.add(descriptionKey);
  });
  return finishZbomCompareFileName(parts.join(' '+separator+' '));
}

function plannedZbomCompareFileNames(){
  const names=new Map(),groups=new Map();
  zbomCompareRows.forEach(row=>{if(row.group===ZBOM_COMPARE_STANDALONE){names.set(row.quote,finishZbomCompareFileName(renderZbomCompareFileName(row)));return} if(!groups.has(row.group))groups.set(row.group,[]); groups.get(row.group).push(row);});
  groups.forEach(rows=>{const name=renderZbomCompareGroupFileName(rows); rows.forEach(row=>names.set(row.quote,name));}); return names;
}

function renderZbomComparePreview(){
  const body=bridgeId('zbomComparePreviewBody');body.replaceChildren();
  if(!zbomCompareRows.length){const tr=document.createElement('tr'),cell=document.createElement('td');cell.colSpan=7;cell.className='missing';cell.textContent='No preview yet.';tr.appendChild(cell);body.appendChild(tr);}else{
    const options=zbomCompareGroupOptions(),fileNames=plannedZbomCompareFileNames();
    zbomCompareRows.forEach(row=>{
      const tr=document.createElement('tr'),quoteCell=document.createElement('td'),fidCell=document.createElement('td'),customerCell=document.createElement('td'),systemCell=document.createElement('td'),groupCell=document.createElement('td'),fileCell=document.createElement('td'),
        statusCell=document.createElement('td'),select=document.createElement('select'),result=zbomCompareResults.get(row.quote);
      quoteCell.textContent=row.quote;fidCell.textContent=row.fid||'{FID}';fidCell.className='compare-fid';customerCell.textContent=row.customer||'';customerCell.className='compare-customer';systemCell.textContent=row.systemDescription||'';
        systemCell.className='compare-system-description'; groupCell.className='compare-group-cell';select.className='compare-group-select';select.dataset.quote=row.quote;select.setAttribute('aria-label','Compare group for quotation '+row.quote);
      options.forEach(value=>{const option=document.createElement('option');option.value=value;option.textContent=value===ZBOM_COMPARE_STANDALONE?'Standalone':value;select.appendChild(option)}); select.value=row.group;select.disabled=zbomCompareRunning;
        select.addEventListener('change',()=>{row.group=select.value;saveZbomCompareGroups();invalidateZbomCompareResults();renderZbomComparePreview(); const next=Array.from(body.querySelectorAll('select')).find(item=>item.dataset.quote===row.quote);if(next)next.focus()});
      groupCell.appendChild(select);fileCell.className='compare-file-name';const fileName=fileNames.get(row.quote)||'',outputPath=result&&result.completed?String(result.outputPath||''):'';if(outputPath){const link=document.createElement('button');link.type='button'; link.className=
        'compare-file-link';link.textContent=fileName;link.title='Open '+outputPath;link.addEventListener('click',()=>openZbomCompareWorkbook(outputPath,link));fileCell.appendChild(link)}else fileCell.textContent=fileName;
        statusCell.textContent=result?result.status:'Ready';statusCell.className='compare-status'+(result&&['Exported','Formatted','Completed'].includes(result.status)?' row-ok':result&&result.status===ZBOM_COMPARE_FID_ERROR?' row-warning':result&&['Failed',
          'Group incomplete'].includes(result.status)?' row-error':'');if(result&&(result.error||result.warning))statusCell.title=result.error||result.warning;tr.append(quoteCell,fidCell,customerCell,systemCell,groupCell,fileCell,statusCell);body.appendChild(tr);
    });
  }
  const count=zbomCompareRows.length,parts=[count+' quotation'+(count===1?'':'s')];
  if(zbomCompareStats.duplicates)parts.push(zbomCompareStats.duplicates+' duplicate'+(zbomCompareStats.duplicates===1?'':'s')+' removed');
  if(zbomCompareStats.ignored)parts.push(zbomCompareStats.ignored+' line'+(zbomCompareStats.ignored===1?'':'s')+' ignored');
  bridgeId('zbomCompareCount').textContent=parts.join(' · '); updateZbomCompareStart();
}

function previewZbomCompareInput(){
  const parsed=parseZbomCompareInput(bridgeId('zbomCompareInput').value); zbomCompareRows=parsed.rows;zbomCompareRows.forEach(enrichZbomCompareRowFromBobj);zbomCompareStats={duplicates:parsed.duplicates,ignored:parsed.ignored};invalidateZbomCompareResults();setZbomCompareProgress(0,
    'Ready.');saveZbomCompareGroups();renderZbomComparePreview();
}

function insertZbomCompareToken(token){
  const input=bridgeId('zbomCompareTemplate'),start=input.selectionStart==null?input.value.length:input.selectionStart,end=input.selectionEnd==null?start:input.selectionEnd; input.value=input.value.slice(0,start)+token+input.value.slice(end);input.focus();
    input.setSelectionRange(start+token.length,start+token.length);saveZbomCompareTemplate();invalidateZbomCompareResults();renderZbomComparePreview();
}

function setZbomCompareProgress(percent,message){
  const value=Math.max(0,Math.min(100,Number(percent)||0)),bar=bridgeId('zbomCompareProgressBar'),number=bridgeId('zbomCompareProgressNumber'),text=bridgeId('zbomCompareProgressText');
  if(bar)bar.style.width=value+'%';
  if(number)number.textContent=Math.round(value)+'%';
  if(text&&message)text.textContent=String(message);
}

function setZbomCompareRunning(value){
  zbomCompareRunning=Boolean(value);
  ['zbomCompareInput','zbomCompareTemplate','zbomCompareSeparator','zbomCompareOutputFolder','zbomCompareKeepTemp','zbomCompareOpenHistory','zbomCompareResetInput','zbomCompareResetTemplate'].forEach(id=>{const element=bridgeId(id);if(element)element.disabled=zbomCompareRunning});
  document.querySelectorAll('[data-zbom-compare-token]').forEach(button=>button.disabled=zbomCompareRunning); document.querySelectorAll('.compare-group-select').forEach(select=>select.disabled=zbomCompareRunning); setBridgeBusy(bridgeBusy);
}

function updateZbomCompareStart(){const button=bridgeId('zbomCompareStart');if(!button)return; const folder=bridgeId('zbomCompareOutputFolder'),path=String(folder&&folder.value||'').trim(),validPath=/^[A-Za-z]:[\\/]/.test(path); button.disabled=!canStartSapTool('zbomCompare')||
  !zbomCompareRows.length||!validPath;}
function normalizedZbomCompareFolder(){const input=bridgeId('zbomCompareOutputFolder'),raw=String(input&&input.value||'').trim().replace(/\//g,'\\'); if(!raw||!/^[A-Za-z]:\\/.test(raw))throw new Error('Use an absolute local output folder such as C:\\PDFFILES\\Compare.');
  return raw.length>3?raw.replace(/\\+$/,''):raw;}

async function cleanupZbomCompareTemp(tempFolder){
  if(keepZbomCompareTemp()){addZbomCompareLog('Temp folder kept: '+tempFolder);return ''}
  try{await bridgeFileRequest('deleteDirectory',{path:tempFolder,recursive:true},BRIDGE_FILE_TIMEOUT);addZbomCompareLog('Temp folder deleted.');return ''} catch(error){const message=bridgeError(error);addZbomCompareLog('TEMP CLEANUP FAILED: '+message);return message}
}

function sapCallArgument(type,value){return {type:type,value:value===null||value===undefined?'':String(value)}}
function addZbomCompareLog(message){zbomCompareLogLines.push('['+new Date().toLocaleTimeString([], {hour12:false})+'] '+message); if(zbomCompareLogLines.length>160)zbomCompareLogLines.shift(); const result=bridgeId('sapResult');if(!result)return; if(activeSapTool===
  'zbomCompare')renderSapLog();}
function isZbomCompareAccessError(text){return /(?:do(?:es)?\s+not|doesn't|don't)\s+have\s+access|no\s+access|not\s+authori[sz]ed|no\s+authori[sz]ation/i.test(String(text||''));}

async function selectZbomCompareOutputType(row){
  const tableId='wnd[1]/usr/tblSAPLVMSGTABCONTROL',cellId=tableId+'/txtNAST-KSCHL[0,'+row+']';
  try{
    await sendZbomAction('zbomCompare','sapCall',{steps:[{action:'invoke',member:'Table',arguments:[sapCallArgument('string',tableId)]}, {action:'invoke',member:'SelectRow',arguments:[sapCallArgument('int',row)]}, {action:'findById',arguments:[sapCallArgument('string',cellId)]},
      {action:'invoke',member:'SetFocus',arguments:[]}]});
  }catch(_){
    await sendZbomAction('zbomCompare','sapCall',{steps:[
      {action:'findById',arguments:[sapCallArgument('string',tableId)]}, {action:'invoke',member:'GetAbsoluteRow',arguments:[sapCallArgument('int',row)]}, {action:'set',member:'Selected',value:sapCallArgument('bool',true)},
      {action:'findById',arguments:[sapCallArgument('string',cellId)]}, {action:'invoke',member:'SetFocus',arguments:[]}
    ]});
  }
}

async function findZbomCompareOutputRow(){
  const tableId='wnd[1]/usr/tblSAPLVMSGTABCONTROL',rows=[zbomCompareOutputRowHint,...Array.from({length:31},(_,row)=>row).filter(row=>row!==zbomCompareOutputRowHint)];let tableOpened=false;
  if(bridgeMock)return {row:4,tableOpened:true};
  for(const row of rows){
    try{const data=await sendZbomAction('zbomCompare','readText',{controlId:tableId+'/txtNAST-KSCHL[0,'+row+']'});tableOpened=true; if(String(data&&data.value||'').trim().toUpperCase()==='ZBOM'){zbomCompareOutputRowHint=row;return {row:row,tableOpened:true}}}
      catch(_){/* an absent row is expected near the end of the table */}
  }
  return {row:-1,tableOpened:tableOpened};
}

async function setZbomCompareSaveTarget(tempFolder,fileName){
  const fullPath=tempFolder+'\\'+fileName;let splitPath=false;
  try{await sendZbomAction('zbomCompare','setText',{controlId:'wnd[1]/usr/ctxtDY_PATH',value:tempFolder}); await sendZbomAction('zbomCompare','setText',{controlId:'wnd[1]/usr/ctxtDY_FILENAME',value:fileName}); splitPath=true;}catch(_){/* older SAP dialogs provide only DY_FILENAME */}
  if(!splitPath)await sendZbomAction('zbomCompare','setText',{controlId:'wnd[1]/usr/ctxtDY_FILENAME',value:fullPath});
  return fullPath;
}

async function waitForZbomCompareFile(path,timeout){
  const started=Date.now(),limit=timeout||SAP_FILE_WAIT_TIMEOUT;let deadline=started+limit,signature='',stableSince=0;
  while(Date.now()<deadline){
    try{
      const result=await bridgeFileRequest('stat',{path:path},10000),file=result&&result.exists&&Array.isArray(result.entries)?result.entries[0]:null;
      if(file){const next=String(file.size)+':'+String(file.lastModified); if(next!==signature){signature=next;stableSince=Date.now();deadline=Math.max(deadline,stableSince+SAP_FILE_SETTLE_MS+SAP_FILE_POLL_MS)} else if(Date.now()-stableSince>=SAP_FILE_SETTLE_MS)return result;}
        else{signature='';stableSince=0}
    }catch(_){signature='';stableSince=0}
    await waitMs(SAP_FILE_POLL_MS);
  }
  throw new Error('SAP did not create '+path+' within '+Math.round(limit/1000)+' seconds.');
}

async function exportOneZbomCompareQuotation(row,tempFolder){
  const quote=row.quote,fileName=quote+'.tsv',fullPath=tempFolder+'\\'+fileName; await sendZbomAction('zbomCompare','transaction',{code:'VA23'}); await sendZbomAction('zbomCompare','setText',{controlId:'wnd[0]/usr/ctxtVBAK-VBELN',value:quote});
  await sendZbomAction('zbomCompare','select',{controlId:'wnd[0]/mbar/menu[0]/menu[5]'}); const popup=await sendZbomAction('zbomCompare','readText',{controlId:'wnd[1]/usr/txtMESSTXT1'}).catch(()=>null),popupText=String(popup?.value||'').trim();
  if(popupText){await sendZbomAction('zbomCompare','press',{controlId:'wnd[1]/tbar[0]/btn[0]'});addZbomCompareLog(quote+' — information popup confirmed: '+popupText); if(isZbomCompareAccessError(popupText))throw new Error(quote+': '+popupText);}
  const {row:outputRow,tableOpened}=await findZbomCompareOutputRow();
  if(outputRow<0)throw new Error(quote+': '+(tableOpened?'ZBOM output type was not found.':popupText||'SAP output dialog did not open.'));
  await selectZbomCompareOutputType(outputRow);addZbomCompareLog(quote+' — ZBOM output selected.'); await sendZbomAction('zbomCompare','press',{controlId:'wnd[1]/tbar[0]/btn[37]'}); await sendZbomAction('zbomCompare','select',{controlId:'wnd[0]/mbar/menu[2]/menu[1]'});
  await sendZbomAction('zbomCompare','select',{controlId:'wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[1]'}); const radioId='wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]'; await sendZbomAction('zbomCompare','select',{controlId:radioId});
  await sendZbomAction('zbomCompare','sapCall',{controlId:radioId,steps:[{action:'invoke',member:'SetFocus',arguments:[]}]}); await sendZbomAction('zbomCompare','press',{controlId:'wnd[1]/tbar[0]/btn[0]'});addZbomCompareLog(quote+' — local save dialog opened.');
  await setZbomCompareSaveTarget(tempFolder,fileName); await sendZbomAction('zbomCompare','press',{controlId:'wnd[1]/tbar[0]/btn[11]'});
  for(const controlId of ['wnd[1]/tbar[0]/btn[11]','wnd[1]/usr/btnSPOP-OPTION1']){try{await sendZbomAction('zbomCompare','press',{controlId:controlId})}catch(_){/* optional overwrite/confirmation dialog */}}
  if(bridgeMock)return {path:fullPath,simulated:true};
  addZbomCompareLog(quote+' — save requested; waiting for '+fullPath+'.'); await waitForZbomCompareFile(fullPath);return {path:fullPath,simulated:false};
}

function decodeZbomCompareBase64(contentBase64){
  const binary=atob(String(contentBase64||'')),bytes=new Uint8Array(binary.length);
  for(let index=0;index<binary.length;index++)bytes[index]=binary.charCodeAt(index);
  if(bytes[0]===0xff&&bytes[1]===0xfe)return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if(bytes[0]===0xfe&&bytes[1]===0xff){const swapped=new Uint8Array(bytes.length-2); for(let index=2;index+1<bytes.length;index+=2){swapped[index-2]=bytes[index+1];swapped[index-1]=bytes[index]} return new TextDecoder('utf-16le').decode(swapped);}
  const offset=bytes[0]===0xef&&bytes[1]===0xbb&&bytes[2]===0xbf?3:0; return new TextDecoder('utf-8').decode(bytes.subarray(offset));
}

function splitZbomCompareTsvLine(line){
  const values=[];let value='',quoted=false;
  for(let index=0;index<line.length;index++){const character=line[index]; if(character==='"'){if(quoted&&line[index+1]==='"'){value+='"';index++}else quoted=!quoted;}else if(character==='\t'&&!quoted){values.push(value);value=''}else value+=character;}
  values.push(value);return values;
}

function parseZbomCompareTsvRows(text){return String(text||'').replace(/^\uFEFF/,'').replace(/\r\n?/g,'\n').split('\n').map(line=>splitZbomCompareTsvLine(line).map(value=>String(value||'').replace(/\u00a0/g,' ').trim()));}
function extractZbomCompareFid(rows){const pattern=/Quotation\s*No\.?\s*[:#]?\s*(\d{6}(?:[-_][A-Za-z0-9]+)*)/i; for(const row of rows){const match=row.join(' ').match(pattern); if(match)return match[1];} return '';}

function parseZbomCompareOptions(rows,gasStart){
  const options=[];let optionTypeColumn=-1,selectionColumn=-1,effectiveType=''; const end=gasStart>=0?gasStart:rows.length;
  for(let rowIndex=0;rowIndex<end;rowIndex++){
    const row=rows[rowIndex],typeHeader=row.findIndex(value=>normalizedZbomHeader(value)==='optiontype'),selectionHeader=row.findIndex(value=>normalizedZbomHeader(value)==='optionselection');
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

function parseZbomCompareTsv(text,path){
  const rows=parseZbomCompareTsvRows(text),fid=extractZbomCompareFid(rows),gasMarker=rows.findIndex(row=>row.some(value=>normalizeZbomGasHeader(value).includes('gasboxconfiguration'))),gasStart=findZbomCompareGasHeaderRow(rows,gasMarker)>=0?gasMarker:-1;
  const options=parseZbomCompareOptions(rows,gasStart),gas=parseZbomCompareGas(rows,gasStart);
  if(!options.length)throw new Error('No option rows were found in '+path+'.');
  return {fid:fid,fidReadError:!fid,options:options,gas:gas,sourcePath:path};
}

async function readZbomCompareTsv(path){const data=await bridgeFileRequest('read',{path:path},BRIDGE_FILE_TIMEOUT); if(!data||!data.contentBase64)throw new Error('SAPGUIcontrol returned no file content for '+path+'.');
  return parseZbomCompareTsv(decodeZbomCompareBase64(data.contentBase64),path);}

function applyZbomCompareFidFallback(model,row){
  if(!model.fidReadError)return '';
  const providedFid=String(row.fid||'').trim(); model.fid=providedFid||row.quote; model.fidWarning='FID could not be read from the exported TSV; using '+(providedFid?'provided FID '+providedFid:'Quotation '+row.quote+' as FID')+'.'; return model.fidWarning;
}

function completedZbomCompareResult(result,outputPath){return {status:result&&result.warning?ZBOM_COMPARE_FID_ERROR:'Completed',outputPath:outputPath,completed:true};}

function alignZbomCompareRecordSets(recordSets,keyOf,similar){
  if(!recordSets.length)return [];
  let slots=(recordSets[0]||[]).map(record=>({records:[record]}));
  for(let fileIndex=1;fileIndex<recordSets.length;fileIndex++){
    const incoming=recordSets[fileIndex]||[],rowCount=slots.length,columnCount=incoming.length,gap=-2,dp=Array.from({length:rowCount+1},()=>Array(columnCount+1).fill(0));
    for(let row=1;row<=rowCount;row++)dp[row][0]=dp[row-1][0]+gap;
    for(let column=1;column<=columnCount;column++)dp[0][column]=dp[0][column-1]+gap;
    const matchScore=(slot,record)=>{const existing=slot.records.filter(Boolean),key=keyOf(record); if(existing.some(item=>keyOf(item)===key))return 12; if(similar&&existing.some(item=>similar(item,record)))return 4; return -5;};
    for(let row=1;row<=rowCount;row++){for(let column=1;column<=columnCount;column++){const diagonal=dp[row-1][column-1]+matchScore(slots[row-1],incoming[column-1]),above=dp[row-1][column]+gap,left=dp[row][column-1]+gap; dp[row][column]=Math.max(diagonal,above,left);}}
    const actions=[];let row=rowCount,column=columnCount;
    while(row>0||column>0){
      const score=row>0&&column>0?matchScore(slots[row-1],incoming[column-1]):-Infinity,diagonal=row>0&&column>0?dp[row-1][column-1]+score:-Infinity,above=row>0?dp[row-1][column]+gap:-Infinity,left=column>0?dp[row][column-1]+gap:-Infinity;
      if(row>0&&column>0&&diagonal===dp[row][column]&&(score>0||(diagonal>above&&diagonal>left))){actions.push(['match',row-1,column-1]);row--;column--}
      else if(row>0&&above===dp[row][column]){actions.push(['missing',row-1,-1]);row--} else{actions.push(['insert',-1,column-1]);column--}
    }
    actions.reverse();
    slots=actions.map(action=>{
      if(action[0]==='insert')return {records:Array(fileIndex).fill(null).concat(incoming[action[2]])};
      const slot=slots[action[1]],records=slot.records.slice();while(records.length<fileIndex)records.push(null);
      records.push(action[0]==='match'?incoming[action[2]]:null);return {records:records};
    });
  }
  slots.forEach(slot=>{while(slot.records.length<recordSets.length)slot.records.push(null)}); return slots;
}

function zbomCompareFileColor(index){
  if(index<ZBOM_COMPARE_FILE_COLORS.length)return ZBOM_COMPARE_FILE_COLORS[index];
  const hue=(index*137.508)%360,saturation=.48,lightness=.84,component=value=>{const amount=saturation*Math.min(lightness,1-lightness),part=(value+hue/30)%12,color=lightness-amount*Math.max(-1,Math.min(part-3,9-part,1)); return Math.round(255*color).toString(16).padStart(2,'0');};
  return 'FF'+component(0)+component(8)+component(4);
}

function normalizedZbomCompareCell(value){return String(value===null||value===undefined?'':value).trim()}
function zbomCompareOptionKey(option){return [option.optionType,option.selection,option.comment].map(normalizedZbomCompareCell).join('\u0000');}

function similarZbomCompareOption(left,right){
  const leftSelection=normalizedZbomCompareCell(left.selection),rightSelection=normalizedZbomCompareCell(right.selection);
  if(leftSelection&&leftSelection===rightSelection)return true;
  const leftType=normalizedZbomCompareCell(left.effectiveType),rightType=normalizedZbomCompareCell(right.effectiveType); return Boolean(leftType&&leftType===rightType);
}

function isZbomCompareSectionOption(option){const title=normalizedZbomCompareCell(option&&option.optionType),selection=normalizedZbomCompareCell(option&&option.selection); return Boolean(title&&!selection&&/[A-Z]/.test(title)&&title===title.toUpperCase());}

function splitZbomCompareOptionSections(options){
  const sections=[];let section=null;
  (options||[]).forEach(option=>{
    if(isZbomCompareSectionOption(option)){const title=normalizedZbomCompareCell(option.optionType).replace(/\s+/g,' ').toUpperCase(); section={key:'SECTION:'+title,records:[]};sections.push(section);}else if(!section){section={key:'PREAMBLE',records:[],preamble:true};
      sections.push(section)}
    section.records.push(option);
  });
  return sections;
}

function scoreZbomCompareSectionContent(left,right){
  const records=section=>section.records.slice(section.preamble?0:1).map(zbomCompareOptionKey),leftKeys=records(left),rightKeys=records(right),counts=new Map(); leftKeys.forEach(key=>counts.set(key,(counts.get(key)||0)+1));let matches=0;
  rightKeys.forEach(key=>{const count=counts.get(key)||0;if(count){matches++;counts.set(key,count-1)}}); return leftKeys.length+rightKeys.length?Math.round(200*matches/(leftKeys.length+rightKeys.length)):100;
}

function scoreZbomCompareSectionMatch(slot,section){const existing=slot.sections.filter(Boolean).filter(item=>item.key===section.key); if(!existing.length)return -1; return 1000+Math.max(...existing.map(item=>scoreZbomCompareSectionContent(item,section)));}

function orderedZbomCompareSectionMatches(slots,incoming){
  const rowCount=slots.length,columnCount=incoming.length,dp=Array.from({length:rowCount+1},()=>Array(columnCount+1).fill(0));
  for(let row=1;row<=rowCount;row++)for(let column=1;column<=columnCount;column++){const score=scoreZbomCompareSectionMatch(slots[row-1],incoming[column-1]),diagonal=score>=0?dp[row-1][column-1]+score:-1; dp[row][column]=Math.max(diagonal,dp[row-1][column],dp[row][column-1]);}
  const matches=[];let row=rowCount,column=columnCount;
  while(row&&column){const score=scoreZbomCompareSectionMatch(slots[row-1],incoming[column-1]),diagonal=score>=0?dp[row-1][column-1]+score:-1; if(score>=0&&diagonal===dp[row][column]){matches.push([row-1,column-1]);row--;column--} else if(dp[row-1][column]>=dp[row][column-1])row--;
    else column--;}
  return matches.reverse();
}

function alignZbomCompareOptionSections(optionSets){
  if(!optionSets.length)return [];
  let slots=splitZbomCompareOptionSections(optionSets[0]).map(section=>({sections:[section]}));
  for(let fileIndex=1;fileIndex<optionSets.length;fileIndex++){
    const incoming=splitZbomCompareOptionSections(optionSets[fileIndex]),assignments=new Map(),usedSlots=new Set(); slots.forEach(slot=>{while(slot.sections.length<=fileIndex)slot.sections.push(null)});
    // Preserve the original Section order wherever an ordered match is possible.
    orderedZbomCompareSectionMatches(slots,incoming).forEach(([slotIndex,incomingIndex])=>{slots[slotIndex].sections[fileIndex]=incoming[incomingIndex];assignments.set(incomingIndex,slots[slotIndex]);usedSlots.add(slotIndex);});

    // Only unmatched Sections with the same title may move to a different position.
    const moved=[];
    slots.forEach((slot,slotIndex)=>{
      if(usedSlots.has(slotIndex))return;
      incoming.forEach((section,incomingIndex)=>{if(assignments.has(incomingIndex)||section.preamble)return; const score=scoreZbomCompareSectionMatch(slot,section); if(score>=0)moved.push({slot:slot,slotIndex:slotIndex,incomingIndex:incomingIndex,score:score,
        distance:Math.abs(slotIndex-incomingIndex)});});
    });
    moved.sort((left,right)=>right.score-left.score||left.distance-right.distance);
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
  return slots.flatMap(slot=>alignZbomCompareRecordSets(slot.sections.map(section=>section?section.records:[]),zbomCompareOptionKey,similarZbomCompareOption));
}

function applyZbomCompareDifference(cells,values,fileIndexes){
  const keys=values.map(item=>{const value=normalizedZbomCompareCell(item&&item.value);return value?'VALUE:'+value:'__BLANK__'});
  if(new Set(keys).size<=1)return;
  const representative=new Map();keys.forEach((key,index)=>{if(!representative.has(key))representative.set(key,fileIndexes[index])}); cells.forEach((cell,index)=>{if(cell)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:zbomCompareFileColor(representative.get(keys[index]))}}});
}

function commonZbomGasHeaders(models){
  if(!models.length)return [];
  let common=new Set(models[0].gas.headers.map(header=>header.compareKey)); models.slice(1).forEach(model=>{const available=new Set(model.gas.headers.map(header=>header.compareKey));common=new Set(Array.from(common).filter(key=>available.has(key)))}); return Array.from(common);
}

function setZbomCompareHeaderStyle(cell){cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:ZBOM_COMPARE_HEADER_FILL}}; cell.font={bold:true,color:{argb:'FFFFFFFF'}}; cell.alignment={vertical:'middle',horizontal:'left',wrapText:true}; setZbomCompareBorder(cell);}
function setZbomCompareBorder(cell){const edge={style:'thin',color:{argb:ZBOM_COMPARE_BORDER_COLOR}}; cell.border={top:edge,left:edge,bottom:edge,right:edge};}
function setZbomCompareDataStyle(cell){cell.alignment={vertical:'top',horizontal:'left',wrapText:true};setZbomCompareBorder(cell);}

function isZbomCompareSectionRow(worksheet,rowNumber,firstColumn){
  firstColumn=firstColumn||1; const title=normalizedZbomCompareCell(worksheet.getCell(rowNumber,firstColumn).value),selection=normalizedZbomCompareCell(worksheet.getCell(rowNumber,firstColumn+1).value);
  return Boolean(title&&!selection&&/[A-Z]/.test(title)&&title===title.toUpperCase());
}

function boldZbomCompareSectionRows(worksheet,lastRow,lastColumn,firstColumn){
  firstColumn=firstColumn||1;
  for(let rowNumber=1;rowNumber<=lastRow;rowNumber++){if(!isZbomCompareSectionRow(worksheet,rowNumber,firstColumn))continue; for(let column=firstColumn;column<=lastColumn;column++){const cell=worksheet.getCell(rowNumber,column);cell.font=Object.assign({},cell.font||{}, {bold:true});}}
}

function applyZbomCompareAllBorders(worksheet,firstRow,lastRow,firstColumn,lastColumn){for(let rowNumber=firstRow;rowNumber<=lastRow;rowNumber++)for(let column=firstColumn;column<=lastColumn;column++)setZbomCompareBorder(worksheet.getCell(rowNumber,column));}

function safeZbomCompareSheetName(workbook,value){
  const original=String(value||'Sheet').replace(/[\[\]:*?/\\]/g,',').replace(/^'+|'+$/g,'').trim()||'Sheet';let name=original.slice(0,31),number=2; const exists=current=>workbook.worksheets.some(sheet=>sheet.name.toLowerCase()===current.toLowerCase());
  while(exists(name)){const suffix=' ('+number+++')';name=original.slice(0,31-suffix.length)+suffix;}
  return name;
}

function prepareZbomCompareWorksheet(worksheet){worksheet.views=[{state:'frozen',ySplit:2,showGridLines:false}];worksheet.properties.defaultRowHeight=16;}
function zbomCompareExcelValue(value){return value===null||value===undefined||String(value)===''?null:value}

function addZbomCompareOptionSheet(workbook,model){
  const worksheet=workbook.addWorksheet(safeZbomCompareSheetName(workbook,model.fid));prepareZbomCompareWorksheet(worksheet); worksheet.getCell(1,1).value='FID No.';worksheet.getCell(1,2).value=model.fid;
  ['Option Type','Option Selection','Comment'].forEach((label,index)=>worksheet.getCell(2,index+1).value=label);
  for(let column=1;column<=3;column++){setZbomCompareHeaderStyle(worksheet.getCell(1,column));setZbomCompareHeaderStyle(worksheet.getCell(2,column))}
  model.options.forEach((option,index)=>{const values=[option.optionType,option.selection,option.comment]; for(let column=1;column<=3;column++){const cell=worksheet.getCell(index+3,column);cell.value=zbomCompareExcelValue(values[column-1]);setZbomCompareDataStyle(cell)}});
  applyZbomCompareAllBorders(worksheet,1,model.options.length+2,1,3);boldZbomCompareSectionRows(worksheet,model.options.length+2,3); worksheet.getColumn(1).width=30;worksheet.getColumn(2).width=38;worksheet.getColumn(3).width=22;worksheet.autoFilter='A2:C2'; return worksheet;
}

function addZbomCompareGasSheet(workbook,model){
  if(!model.gas)return null;
  const worksheet=workbook.addWorksheet(safeZbomCompareSheetName(workbook,model.fid+'_Gas'));prepareZbomCompareWorksheet(worksheet); const width=model.gas.headers.length;worksheet.getCell(1,1).value='FID No.';if(width>1)worksheet.getCell(1,2).value=model.fid;
  model.gas.headers.forEach((header,index)=>worksheet.getCell(2,index+1).value=header.label);
  for(let column=1;column<=width;column++){setZbomCompareHeaderStyle(worksheet.getCell(1,column));setZbomCompareHeaderStyle(worksheet.getCell(2,column))}
  model.gas.records.forEach((record,rowIndex)=>{model.gas.headers.forEach((header,columnIndex)=>{const cell=worksheet.getCell(rowIndex+3,columnIndex+1);cell.value=zbomCompareExcelValue(record.values[header.key]);setZbomCompareDataStyle(cell)});});
  applyZbomCompareAllBorders(worksheet,1,model.gas.records.length+2,1,width);
  model.gas.headers.forEach((header,index)=>{const values=model.gas.records.map(record=>String(record.values[header.key]||'').length),longest=Math.max(header.label.length,...values); worksheet.getColumn(index+1).width=Math.max(11,Math.min(35,longest+2));}); return worksheet;
}

function addZbomCompareOptionComparison(workbook,models){
  const worksheet=workbook.addWorksheet(safeZbomCompareSheetName(workbook,'Compare'));prepareZbomCompareWorksheet(worksheet); const slots=alignZbomCompareOptionSections(models.map(model=>model.options));
  models.forEach((model,fileIndex)=>{
    const start=1+fileIndex*4;worksheet.getCell(1,start).value='FID No.';worksheet.getCell(1,start+1).value=model.fid; ['Option Type','Option Selection','Comment'].forEach((label,index)=>worksheet.getCell(2,start+index).value=label);
    for(let column=0;column<3;column++){setZbomCompareHeaderStyle(worksheet.getCell(1,start+column));setZbomCompareHeaderStyle(worksheet.getCell(2,start+column))}
    worksheet.getColumn(start).width=30;worksheet.getColumn(start+1).width=38;worksheet.getColumn(start+2).width=22;if(fileIndex<models.length-1)worksheet.getColumn(start+3).width=3;
  });
  slots.forEach((slot,rowIndex)=>{
    models.forEach((model,fileIndex)=>{
      const start=1+fileIndex*4,record=slot.records[fileIndex],values=record?[record.optionType,record.selection,record.comment]:[null,null,null];
      values.forEach((value,columnIndex)=>{const cell=worksheet.getCell(rowIndex+3,start+columnIndex);cell.value=zbomCompareExcelValue(value);setZbomCompareDataStyle(cell)});
    });
    const fields=['optionType','selection','comment'];
    fields.forEach((field,columnIndex)=>{
      const cells=[],values=[],fileIndexes=[];
      models.forEach((model,fileIndex)=>{const record=slot.records[fileIndex],start=1+fileIndex*4;cells.push(worksheet.getCell(rowIndex+3,start+columnIndex));values.push(record?{value:record[field],missing:false}:{value:'',missing:true});fileIndexes.push(fileIndex);});
      applyZbomCompareDifference(cells,values,fileIndexes);
    });
  });
  const lastRow=slots.length+2;models.forEach((model,fileIndex)=>{const start=1+fileIndex*4;applyZbomCompareAllBorders(worksheet,1,lastRow,start,start+2);boldZbomCompareSectionRows(worksheet,lastRow,start+2,start)}); return worksheet;
}

function gasRecordValue(model,record,compareKey){const header=model.gas.headers.find(item=>item.compareKey===compareKey);return header?record.values[header.key]||'':'';}

function addZbomCompareGasComparison(workbook,models){
  const gasModels=models.map((model,fileIndex)=>({model:model,fileIndex:fileIndex})).filter(item=>item.model.gas);
  if(!gasModels.length)return null;
  const worksheet=workbook.addWorksheet(safeZbomCompareSheetName(workbook,'Compare_Gas'));prepareZbomCompareWorksheet(worksheet); const blocks=[];let start=1;
  models.forEach((model,fileIndex)=>{
    const width=model.gas?Math.max(2,model.gas.headers.length):3;blocks.push({start:start,width:width,model:model,fileIndex:fileIndex}); worksheet.getCell(1,start).value='FID No.';worksheet.getCell(1,start+1).value=model.fid;
    for(let column=0;column<width;column++)setZbomCompareHeaderStyle(worksheet.getCell(1,start+column));
    if(model.gas){model.gas.headers.forEach((header,index)=>{const cell=worksheet.getCell(2,start+index);cell.value=header.label;setZbomCompareHeaderStyle(cell);worksheet.getColumn(start+index).width=Math.max(11,Math.min(28,header.label.length+4))});}else{
      worksheet.getColumn(start).width=11;worksheet.getColumn(start+1).width=14;worksheet.getColumn(start+2).width=5; worksheet.mergeCells(2,start,2,start+width-1);const cell=worksheet.getCell(2,start);cell.value='No Gas';setZbomCompareHeaderStyle(cell);cell.fill={type:'pattern',
        pattern:'solid',fgColor:{argb:zbomCompareFileColor(fileIndex)}};
    }
    start+=width+1;if(fileIndex<models.length-1)worksheet.getColumn(start-1).width=3;
  });
  const slots=alignZbomCompareRecordSets(gasModels.map(item=>item.model.gas.records),record=>record.key,(left,right)=>Boolean(left.module&&left.module===right.module));
  slots.forEach((slot,rowIndex)=>{
    gasModels.forEach((item,gasIndex)=>{
      const block=blocks[item.fileIndex],record=slot.records[gasIndex];
      item.model.gas.headers.forEach((header,columnIndex)=>{const cell=worksheet.getCell(rowIndex+3,block.start+columnIndex);cell.value=record?zbomCompareExcelValue(record.values[header.key]):null;setZbomCompareDataStyle(cell)});
    });
  });
  const commonHeaders=commonZbomGasHeaders(gasModels.map(item=>item.model));
  slots.forEach((slot,rowIndex)=>{
    commonHeaders.forEach(compareKey=>{
      const cells=[],values=[],fileIndexes=[];
      gasModels.forEach((item,gasIndex)=>{
        const headerIndex=item.model.gas.headers.findIndex(header=>header.compareKey===compareKey),block=blocks[item.fileIndex],record=slot.records[gasIndex]; cells.push(worksheet.getCell(rowIndex+3,block.start+headerIndex));values.push(record?{value:gasRecordValue(item.model,record,
          compareKey),missing:false}:{value:'',missing:true});fileIndexes.push(item.fileIndex);
      });
      applyZbomCompareDifference(cells,values,fileIndexes);
    });
  });
  const lastRow=slots.length+2;gasModels.forEach(item=>{const block=blocks[item.fileIndex];applyZbomCompareAllBorders(worksheet,1,lastRow,block.start,block.start+item.model.gas.headers.length-1)}); return worksheet;
}

function snapshotZbomCompareWorksheet(worksheet){
  if(!worksheet)return null;
  const rows=[];
  for(let rowNumber=1;rowNumber<=worksheet.rowCount;rowNumber++){
    const row=[];
    for(let column=1;column<=worksheet.columnCount;column++){
      const cell=worksheet.getCell(rowNumber,column),fill=cell.fill&&cell.fill.fgColor&&cell.fill.fgColor.argb,fontColor=cell.font&&cell.font.color&&cell.font.color.argb;
      row.push({text:cell.text==null?'':String(cell.text),fill:fill||'',fontColor:fontColor||'',bold:Boolean(cell.font&&cell.font.bold)});
    }
    rows.push(row);
  }
  return {rows:rows,widths:Array.from({length:worksheet.columnCount},(_,index)=>Number(worksheet.getColumn(index+1).width)||12)};
}

function buildZbomCompareWorkbook(models,grouped){
  if(typeof ExcelJS==='undefined'||!ExcelJS.Workbook)throw new Error('Excel workbook library is unavailable. Refresh the page and try again.');
  const workbook=new ExcelJS.Workbook();workbook.creator='NSR Tracking Dashboard';workbook.company='Lam Research';workbook.created=new Date();
  const compareSheet=grouped&&models.length>1?addZbomCompareOptionComparison(workbook,models):null,compareGasSheet=grouped&&models.length>1?addZbomCompareGasComparison(workbook,models):null,modelSheets=[];
  models.forEach(model=>modelSheets.push({fid:model.fid,quote:model.quote,options:addZbomCompareOptionSheet(workbook,model),gas:addZbomCompareGasSheet(workbook,model)}));
  return {workbook:workbook,preview:{compare:snapshotZbomCompareWorksheet(compareSheet),compareGas:snapshotZbomCompareWorksheet(compareGasSheet),models:modelSheets.map(item=>({fid:item.fid,quote:item.quote,options:snapshotZbomCompareWorksheet(item.options),
    gas:snapshotZbomCompareWorksheet(item.gas)}))}};
}

function zbomCompareBufferBase64(buffer){
  const bytes=buffer instanceof ArrayBuffer?new Uint8Array(buffer):new Uint8Array(buffer.buffer,buffer.byteOffset||0,buffer.byteLength),chunkSize=0x8000;let binary='';
  for(let offset=0;offset<bytes.length;offset+=chunkSize)binary+=String.fromCharCode.apply(null,bytes.subarray(offset,Math.min(offset+chunkSize,bytes.length)));
  return btoa(binary);
}

async function writeZbomCompareWorkbook(models,grouped,path){
  const built=buildZbomCompareWorkbook(models,grouped),buffer=await built.workbook.xlsx.writeBuffer(); await bridgeFileRequest('write',{path:path,contentBase64:zbomCompareBufferBase64(buffer),overwrite:true,createParents:true},BRIDGE_FILE_TIMEOUT);
  return {path:path,preview:built.preview};
}

function renderZbomCompareResultSummary(exported,outputs,failed,mock){
  const container=bridgeId('zbomCompareResultSummary');if(!container)return;
  container.replaceChildren(); const items=[];
  if(mock)items.push(['MOCK: no real TSV or XLSX written','warn']);
  items.push(['TSV exported: '+exported,exported?'ok':'warn'],['XLSX created: '+outputs,outputs?'ok':mock?'warn':''],['Failed: '+failed,failed?'warn':'ok']);
  items.forEach(item=>{const pill=document.createElement('span');pill.className='result-pill'+(item[1]?' '+item[1]:'');pill.textContent=item[0];container.appendChild(pill);});
}

async function startZbomCompare(){
  await ensureBobjDataSource();previewZbomCompareInput(); const rows=zbomCompareRows.slice(),folderInput=bridgeId('zbomCompareOutputFolder');let folder,tempFolder;
  try{folder=normalizedZbomCompareFolder();tempFolder=folder+'\\temp'}catch(error){setZbomCompareProgress(0,bridgeError(error));if(folderInput)folderInput.focus();return}
  if(!rows.length){setZbomCompareProgress(0,'Paste at least one valid Quotation first.');return}
  if(!selectedBridgeSession('zbomCompare')){setZbomCompareProgress(0,'Select an available SAP GUI session first.');return}
  if(rows.length>100&&!unsafeWindow.confirm('You are about to export '+rows.length+' quotations through SAP. Continue?'))return;
  if(folderInput){folderInput.value=folder;saveZbomCompareFolder()}
  await acquireSapTask('zbomCompare',{outputFolder:folder}); tempFolder=folder+'\\temp-'+sapTools.zbomCompare.job.taskId; recordZbomCompareHistory(bridgeId('zbomCompareInput').value);invalidateZbomCompareResults();zbomCompareLogLines=[];renderSapLog();setZbomCompareRunning(true);
    renderZbomComparePreview();setZbomCompareProgress(2,'Preparing the temp folder…'); const formatted=new Map(),updateResult=(row,patch)=>{zbomCompareResults.set(row.quote,Object.assign({},zbomCompareResults.get(row.quote)||{},patch));renderZbomComparePreview();};
    const groupPreviews=[],fidPreviews=new Map();let exported=0,outputs=0,tempReady=false,sapFailed=false;
  try{
    await bridgeFileRequest('createDirectory',{path:tempFolder},30000);tempReady=true;
    for(let index=0;index<rows.length;index++){
      const row=rows[index],percent=5+(index/rows.length)*52;setZbomCompareProgress(percent,'Exporting '+row.quote+' ('+(index+1)+'/'+rows.length+')…'); updateResult(row,{status:'Exporting',error:''});
      try{addZbomCompareLog(row.quote+' — export started.');const result=await exportOneZbomCompareQuotation(row,tempFolder);updateResult(row,{status:'Exported',path:result.path,simulated:result.simulated});exported++;addZbomCompareLog(row.quote+' — TSV exported.');}
        catch(error){sapFailed=true;const message=bridgeError(error);updateResult(row,{status:'Failed',error:message});addZbomCompareLog(row.quote+' — FAILED: '+message);}
    }
    if(bridgeMock){sapTools.zbomCompare.job.keep=sapFailed; zbomCompareHasRun=true;renderZbomCompareResultSummary(exported,0,rows.length-exported,true);setZbomCompareProgress(100,'Mock export completed. No files were formatted.');setBridgeBadge('Mock ready','warn');return;}
    sapTools.zbomCompare.job.keep=sapFailed; const exportedRows=rows.filter(row=>zbomCompareResults.get(row.quote)&&zbomCompareResults.get(row.quote).status==='Exported');
    for(let index=0;index<exportedRows.length;index++){
      const row=exportedRows[index],result=zbomCompareResults.get(row.quote);setZbomCompareProgress(58+(index/Math.max(1,exportedRows.length))*16,'Formatting '+row.quote+' ('+(index+1)+'/'+exportedRows.length+')…');updateResult(row,{status:'Formatting'});
      try{
        const model=await readZbomCompareTsv(result.path),providedFid=row.fid;model.quote=row.quote; const fidWarning=applyZbomCompareFidFallback(model,row);
        if(fidWarning)addZbomCompareLog(row.quote+' — FID READ ERROR: '+fidWarning); else if(providedFid&&providedFid!==model.fid)addZbomCompareLog(row.quote+' — FID updated from '+providedFid+' to '+model.fid+'.');
        formatted.set(row.quote,model);row.fid=model.fid;await loadBobjRows([model.fid]);enrichZbomCompareRowFromBobj(row);updateResult(row,{status:fidWarning?ZBOM_COMPARE_FID_ERROR:'Formatted',fid:model.fid,warning:fidWarning});
          addZbomCompareLog(row.quote+' — formatted as '+model.fid+(model.gas?' with Gas.':' without Gas.'));
      }catch(error){const message=bridgeError(error);updateResult(row,{status:'Failed',error:message});addZbomCompareLog(row.quote+' — FORMAT FAILED: '+message);}
    }
    // Check FID BOBJ only after SAP has supplied the FIDs used for the output file names.
    await prepareBobjFileNames('zbomCompare',rows.filter(row=>formatted.has(row.quote)&&zbomCompareResults.get(row.quote)?.status!=='Failed'&&bobjFidKey(row.fid)),bridgeId('zbomCompareTemplate').value||ZBOM_COMPARE_DEFAULT_TEMPLATE); renderZbomComparePreview();
    const standaloneRows=rows.filter(row=>row.group===ZBOM_COMPARE_STANDALONE),groups=new Map(); rows.filter(row=>row.group!==ZBOM_COMPARE_STANDALONE).forEach(row=>{if(!groups.has(row.group))groups.set(row.group,[]);groups.get(row.group).push(row)});
    const outputJobs=Math.max(1,standaloneRows.length+groups.size);let completedJobs=0;
    for(const row of standaloneRows){
      const model=formatted.get(row.quote);if(!model)continue;
      const fileName=finishZbomCompareFileName(renderZbomCompareFileName(row)),outputPath=folder+'\\'+fileName; setZbomCompareProgress(75+(completedJobs/outputJobs)*23,'Saving '+fileName+'…');updateResult(row,{status:'Saving'});
      try{
        const saved=await writeZbomCompareWorkbook([model],false,outputPath),modelPreview=saved.preview.models[0];outputs++;if(modelPreview&&!fidPreviews.has(model.fid))fidPreviews.set(model.fid,{fid:model.fid,sheet:modelPreview.options, order:rows.findIndex(item=>item.quote===
          row.quote)});updateResult(row,completedZbomCompareResult(zbomCompareResults.get(row.quote),outputPath));addZbomCompareLog(row.quote+' — workbook saved: '+outputPath);
      }catch(error){const message=bridgeError(error);updateResult(row,{status:'Failed',error:message});addZbomCompareLog(row.quote+' — XLSX FAILED: '+message);}
      completedJobs++;
    }
    for(const [group,groupRows] of groups){
      const missing=groupRows.filter(row=>!formatted.has(row.quote));
      if(missing.length){
        const message=group+' was not created because '+missing.map(row=>row.quote).join(', ')+' failed before comparison.'; groupRows.filter(row=>formatted.has(row.quote)).forEach(row=>updateResult(row,{status:'Group incomplete',error:message}));addZbomCompareLog(message);
          completedJobs++;continue;
      }
      const models=groupRows.map(row=>formatted.get(row.quote)),fileName=renderZbomCompareGroupFileName(groupRows),outputPath=folder+'\\'+fileName; const comparing=groupRows.length>1;
      setZbomCompareProgress(75+(completedJobs/outputJobs)*23,(comparing?'Comparing '+group+' and saving ':'Saving ')+fileName+'…');groupRows.forEach(row=>updateResult(row,{status:comparing?'Comparing':'Saving'}));
      try{
        const saved=await writeZbomCompareWorkbook(models,comparing,outputPath);outputs++;if(saved.preview.models)models.forEach(model=>{const modelPreview=saved.preview.models.find(item=>item.quote===model.quote)||saved.preview.models.find(item=>item.fid===model.fid);
          if(modelPreview&&!fidPreviews.has(model.fid))fidPreviews.set(model.fid,{fid:model.fid,sheet:modelPreview.options,order:rows.findIndex(item=>item.quote===model.quote)})});groupPreviews.push({group:group,options:comparing?saved.preview.compare:saved.preview.models[0]?.options,
            gas:comparing?saved.preview.compareGas:saved.preview.models[0]?.gas});groupRows.forEach(row=>updateResult(row,completedZbomCompareResult(zbomCompareResults.get(row.quote),outputPath)));addZbomCompareLog(group+' — workbook saved: '+outputPath);
      }catch(error){const message=bridgeError(error);groupRows.forEach(row=>updateResult(row,{status:'Failed',error:message}));addZbomCompareLog(group+' — XLSX FAILED: '+message);}
      completedJobs++;
    }
    const failed=rows.filter(row=>{const result=zbomCompareResults.get(row.quote);return !result||!result.completed}).length; zbomCompareHasRun=true;renderZbomCompareResultSummary(exported,outputs,failed,false);showZbomCompareResultPreview(groupPreviews,fidPreviews);
      setZbomCompareProgress(100,failed?'Completed with '+failed+' failed or incomplete quotation(s).':'Compare completed.'); setBridgeBadge(failed?'Completed with errors':'Ready',failed?'warn':'ok');
  }catch(error){
    const failed=rows.filter(row=>{const result=zbomCompareResults.get(row.quote);return !result||!result.completed}).length; zbomCompareHasRun=true;renderZbomCompareResultSummary(exported,outputs,failed,false);showZbomCompareResultPreview(groupPreviews,fidPreviews);
      setZbomCompareProgress(100,'Compare stopped: '+bridgeError(error));setBridgeBadge('Action failed','error');
  }finally{if(tempReady){const cleanupError=await cleanupZbomCompareTemp(tempFolder);if(cleanupError)setZbomCompareProgress(100,'Compare finished, but the temp folder could not be deleted: '+cleanupError)} setZbomCompareRunning(false);}
}

function bestZbomColumn(rows,used,score,minScore){
  const width=Math.max(0,...rows.map(row=>row.length));let best=-1,bestScore=minScore;
  for(let column=0;column<width;column++){if(used.has(column))continue; const values=rows.map(row=>String(row[column]||'').trim()).filter(Boolean); if(!values.length)continue; const current=score(values); if(current>bestScore){best=column;bestScore=current}}
  if(best>=0)used.add(best);
  return best;
}

function inferZbomColumns(rows){
  const used=new Set(),mapping={fid:-1,systemDescription:-1,quote:-1}; mapping.quote=bestZbomColumn(rows,used,values=>values.reduce((total,value)=>total+(/^\d{8}$/.test(value)?4+(value[0]==='2'?1:0):0),0)/values.length,1.5);
  mapping.fid=bestZbomColumn(rows,used,values=>values.reduce((total,value)=>total+(isZbomFid(value)?4:0),0)/values.length,1.5);
  mapping.systemDescription=bestZbomColumn(rows,used,values=>values.reduce((total,value)=>total+(/[A-Za-z]/.test(value)?2+(/[\/_]/.test(value)?2:0):0),0)/values.length,.8); return mapping;
}

function isZbomFid(value){return /^\d{6}(?:-[A-Za-z0-9-]+)?$/.test(String(value||'').trim());}
function pickZbomValue(values,index){return index>=0&&index<values.length?String(values[index]||'').trim():'';}

function inferZbomRow(values,mapping){
  const used=new Set(); let quote=pickZbomValue(values,mapping.quote);if(mapping.quote>=0&&quote)used.add(mapping.quote);
  if(!/^\d{8}$/.test(quote)){const candidates=values.map((value,index)=>({value:String(value||'').trim(),index:index})).filter(item=>/^\d{8}$/.test(item.value)); if(candidates.length===1){quote=candidates[0].value;used.add(candidates[0].index)}}
  let fid=pickZbomValue(values,mapping.fid);if(mapping.fid>=0&&fid)used.add(mapping.fid);
  if(!fid){const candidate=values.map((value,index)=>({value:String(value||'').trim(),index:index})).find(item=>!used.has(item.index)&&isZbomFid(item.value)); if(candidate){fid=candidate.value;used.add(candidate.index)}}
  let systemDescription=pickZbomValue(values,mapping.systemDescription);if(mapping.systemDescription>=0&&systemDescription)used.add(mapping.systemDescription);
  if(!systemDescription){const candidate=values.map((value,index)=>({value:String(value||'').trim(),index:index})).find(item=>!used.has(item.index)&&/[A-Za-z]/.test(item.value)); if(candidate){systemDescription=candidate.value;used.add(candidate.index)}}
  return {fid:fid,systemDescription:systemDescription,quote:quote};
}

function parseZbomInput(text){
  const source=String(text||'').replace(/\r/g,'').split('\n').map((line,index)=>({line:index+1,values:splitZbomLine(line)})).filter(item=>item.values.some(Boolean));
  if(!source.length)return {rows:[],errors:['No data was pasted.'],duplicates:0,detection:'No columns detected.'};
  const firstRoles=source[0].values.map(zbomHeaderRole),hasHeader=firstRoles.includes('quote')&&firstRoles.some(Boolean); let mapping={fid:-1,systemDescription:-1,quote:-1};
  if(hasHeader)firstRoles.forEach((role,index)=>{if(role&&mapping[role]<0)mapping[role]=index});
  const data=hasHeader?source.slice(1):source;
  if(!hasHeader)mapping=inferZbomColumns(data.map(item=>item.values));
  const rows=[],errors=[],seen=new Map();let duplicates=0;
  data.forEach(item=>{
    const row=Object.assign(inferZbomRow(item.values,mapping),{line:item.line,_error:'',_status:'Ready'});
    if(!row.quote)row._error='Quote was not detected.'; else if(!/^\d{8}$/.test(row.quote))row._error='Quote must contain exactly 8 digits.';
    if(row._error){row._status=row._error;errors.push('Row '+item.line+': '+row._error);rows.push(row);return}
    const previous=seen.get(row.quote);
    if(previous){
      const conflicts=[];
      [['fid','FID'],['systemDescription','System Description']].forEach(pair=>{const field=pair[0],label=pair[1]; if(!previous[field]&&row[field])previous[field]=row[field]; else if(previous[field]&&row[field]&&previous[field]!==row[field])conflicts.push(label);});
      if(conflicts.length){row._error='Duplicate Quote has conflicting '+conflicts.join(', ')+'.';row._status=row._error; errors.push('Row '+item.line+': '+row._error);rows.push(row);}else duplicates++;
      return;
    }
    seen.set(row.quote,row);rows.push(row);
  });
  const validCount=rows.filter(row=>!row._error).length;
  if(validCount>500)errors.push('A maximum of 500 unique quotations can be exported at once.');
  const labels={fid:'FID',systemDescription:'System Description',quote:'Quote'}; const detected=Object.keys(mapping).filter(role=>mapping[role]>=0).map(role=>labels[role]+' = column '+(mapping[role]+1));
  const detection=(hasHeader?'Header mapping: ':'Inferred mapping: ')+(detected.length?detected.join(' · '):'Quote will be detected row by row.'); return {rows:rows,errors:errors,duplicates:duplicates,detection:detection};
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

function renderZbomFileName(row,template){
  let stem=renderBobjTemplate(template||ZBOM_DEFAULT_TEMPLATE,row); stem=stem.replace(/[<>:"/\\|?*\u0000-\u001F]/g,'-').replace(/_{2,}/g,'_').replace(/-{2,}/g,'-').replace(/\s{2,}/g,' ').replace(/^[\s._-]+|[\s._-]+$/g,'');
  if(!stem)stem=row.quote||'ZBOM_PDF';
  if(stem.length>190)stem=stem.slice(0,190).replace(/[\s._-]+$/g,'');
  return stem+'.PDF';
}

function appendQuoteToName(name,quote){const dot=name.toLowerCase().lastIndexOf('.pdf'),stem=dot>=0?name.slice(0,dot):name; return stem+'_'+quote+'.PDF';}

function plannedZbomNames(rows,template){
  const owners=new Map(),names=new Map();
  rows.filter(row=>!row._error).forEach(row=>{let name=renderZbomFileName(row,template),key=name.toLowerCase(),owner=owners.get(key); if(owner&&owner!==row.quote){name=appendQuoteToName(name,row.quote);key=name.toLowerCase()} owners.set(key,row.quote);names.set(row.quote,name);});
  return names;
}

function renderZbomPreview(){
  const body=bridgeId('zbomPreviewBody');body.replaceChildren(); const templateInfo=zbomTemplateInfo(),planned=plannedZbomNames(zbomRows.map(row=>Object.assign({},row,{fid:row.fid||'[SAP FID]'})),templateInfo.template);
  if(!zbomRows.length){const row=document.createElement('tr'),cell=document.createElement('td');cell.colSpan=5;cell.className='missing';cell.textContent='No preview yet.';row.appendChild(cell);body.appendChild(row);}else{
    zbomRows.forEach(row=>{
      const tr=document.createElement('tr'),result=zbomResults.get(row.quote); const actual=result&&result.files.size?Array.from(result.files).join('; '):planned.get(row.quote)||''; const values=[row.fid||'From SAP',row.systemDescription,row.quote,actual];
      values.forEach((value,index)=>{const cell=document.createElement('td');cell.textContent=value||'—';if(!value)cell.className='missing';if(index===3)cell.classList.add('planned-name');tr.appendChild(cell);}); const status=document.createElement('td');
      status.textContent=row._error||templateInfo.error||(result?result.status:'Ready'); status.className=row._error||templateInfo.error||(result&&result.status==='No PDF generated')?'row-error':'row-ok'; tr.appendChild(status);body.appendChild(tr);
    });
  }
  const valid=zbomRows.filter(row=>!row._error).length; bridgeId('zbomCount').textContent=valid+' valid row'+(valid===1?'':'s')+(zbomRows.length!==valid?' · '+(zbomRows.length-valid)+' invalid':''); updateZbomStart();
}

function previewZbomInput(){
  const parsed=parseZbomInput(bridgeId('zbomInput').value); zbomRows=parsed.rows;zbomRows.forEach(row=>{if(row&&row.fid&&!row.systemDescription)row.systemDescription=bobjValue(row.fid,'System Description')});zbomErrors=parsed.errors;zbomHasRun=false;zbomResults=new Map();
  bridgeId('zbomResultSummary').replaceChildren();renderZbomPreview();
}

function insertZbomToken(token){
  const input=bridgeId('zbomTemplate'),start=input.selectionStart==null?input.value.length:input.selectionStart,end=input.selectionEnd==null?start:input.selectionEnd; input.value=input.value.slice(0,start)+token+input.value.slice(end);input.focus();input.setSelectionRange(start+
    token.length,start+token.length); saveZbomTemplate();zbomHasRun=false;zbomResults=new Map();bridgeId('zbomResultSummary').replaceChildren();renderZbomPreview();
}

function updateZbomStart(){const button=bridgeId('zbomStart');if(!button)return; const valid=zbomRows.filter(row=>!row._error).length,templateError=zbomTemplateInfo().error; button.disabled=!canStartSapTool('zbom')||!valid||Boolean(zbomErrors.length)||Boolean(templateError);}
function bridgeFileRequest(operation,options,timeout){return bridgeRequest('POST','/files',Object.assign({operation:operation},options||{}),timeout||BRIDGE_FILE_TIMEOUT);}

async function openZbomFolder(){
  try{const result=await openBridgeFolder(ZBOM_PDF_FOLDER);addZbomLog('Opened '+ZBOM_PDF_FOLDER+' in Windows Explorer.'+(result.foreground?' Window activated.':' Please switch to it manually.'));setBridgeBadge('Ready','ok');}catch(error){addZbomLog('Could not open '+ZBOM_PDF_FOLDER+
    ': '+bridgeError(error));setBridgeBadge('Action failed','error');}
}

async function listZbomRawPdfs(){
  const data=await bridgeFileRequest('list',{path:ZBOM_PDF_FOLDER,pattern:'00*.pdf',recursive:false},30000),files=new Map();
  (Array.isArray(data.entries)?data.entries:[]).forEach(entry=>{const name=String(entry.name||''); if(entry.kind!=='file'||!/^00.*\.pdf$/i.test(name)||zbomOutputClaims.has(name.toLowerCase()))return; files.set(name,{name:name,path:String(entry.path||''),size:Number(entry.size)||0,
    lastModified:Number(entry.lastModified)||0});});
  return files;
}

function changedZbomFiles(before,current){const changed=[]; current.forEach((file,name)=>{const old=before.get(name); if(!old||old.size!==file.size||old.lastModified!==file.lastModified)changed.push(file);}); return changed;}
function waitMs(milliseconds){return new Promise(resolve=>setTimeout(resolve,milliseconds))}
function zbomFilesCoverRows(files,rows){const found=new Set(files.map(file=>matchZbomRow(file.name)?.quote).filter(Boolean)); return rows.every(row=>zbomResults.get(row.quote)?.files.size||found.has(row.quote));}

async function waitForZbomFiles(before,rows){
  const firstFileDeadline=Date.now()+SAP_PDF_FIRST_FILE_WAIT_MS; let deadline=Date.now()+SAP_FILE_WAIT_TIMEOUT,lastSignature='',stableSince=0,latest=[];
  while(Date.now()<deadline){
    const current=await listZbomRawPdfs();latest=changedZbomFiles(before,current);
    if(!latest.length){if(Date.now()>=firstFileDeadline)return []; lastSignature='';stableSince=0; await waitMs(SAP_FILE_POLL_MS);continue;}
    const signature=latest.slice().sort((a,b)=>a.name.localeCompare(b.name)).map(file=>file.name+':'+file.size+':'+file.lastModified).join('|');
    if(signature===lastSignature){const settle=zbomFilesCoverRows(latest,rows)?SAP_FILE_SETTLE_MS:SAP_BATCH_SETTLE_MS; if(Date.now()-stableSince>=settle)return latest;}else{lastSignature=signature;stableSince=Date.now();deadline=Math.max(deadline,stableSince+SAP_BATCH_SETTLE_MS+
      SAP_FILE_POLL_MS)}
    await waitMs(SAP_FILE_POLL_MS);
  }
  return latest;
}

function addZbomLog(message){zbomLogLines.push('['+new Date().toLocaleTimeString([], {hour12:false})+'] '+message); if(zbomLogLines.length>120)zbomLogLines.shift(); const result=bridgeId('sapResult'); if(activeSapTool==='zbom')renderSapLog();}
function setZbomProgress(percent,message){const value=Math.max(0,Math.min(100,Number(percent)||0)); bridgeId('zbomProgressBar').style.width=value+'%';bridgeId('zbomProgressNumber').textContent=Math.round(value)+'%';}

async function sendZbomAction(tool,operation,options){
  const job=sapTools[tool].job;if(!job)throw new Error('No SAP window is assigned to this task.');
  const payload=Object.assign({},options||{},{operation:operation,sessionKey:job.sessionKey,taskId:job.taskId,ownerId:job.ownerId}),started=Date.now();
  while(Date.now()-started<SAP_BUSY_TIMEOUT){try{return await bridgeRequest('POST','/action',payload,SAP_ACTION_TIMEOUT)}catch(error){if(!error||error.status!==409||!/SAP is busy/i.test(bridgeError(error)))throw error; await waitMs(1000);}}
  throw new Error('SAP remained busy for 2 minutes.');
}

function sapStatusSummary(data){return data&&data.statusText?' SAP '+(data.statusType||'status')+': '+data.statusText:'';}
async function readClipboardBestEffort(){try{const clipboard=unsafeWindow.navigator&&unsafeWindow.navigator.clipboard; return clipboard&&typeof clipboard.readText==='function'?await Promise.race([clipboard.readText(),waitMs(1000).then(()=>null)]):null;}catch(_){return null}}
function writeQuotationClipboard(text){if(typeof GM_setClipboard!=='function')throw new Error('Tampermonkey clipboard permission is unavailable. Reinstall or update this userscript.'); GM_setClipboard(text,'text');}

function matchZbomRow(rawName){
  const base=rawName.replace(/\.pdf$/i,''),metadata=rawZbomMetadata(rawName);
  if(metadata.quote){const exact=zbomRows.find(row=>!row._error&&row.quote===metadata.quote);if(exact)return exact}
  const matches=zbomRows.filter(row=>!row._error&&base.indexOf(row.quote)>=0); return matches.length===1?matches[0]:null;
}

function rawZbomMetadata(rawName){
  const base=String(rawName||'').replace(/\.pdf$/i,''),parts=base.split('_'),quoteMatch=String(parts[0]||'').match(/^00(\d{8})$/),timestamp=parts.find(part=>/^\d{14}$/.test(part))||'',fid=parts.length>3?parts.slice(3).join('_').trim():'';
  return {quote:quoteMatch?quoteMatch[1]:'',timestamp:timestamp,fid:fid};
}

function safeUnmatchedName(name){let result=String(name||'Unmatched.pdf').replace(/[<>:"/\\|?*\u0000-\u001F]/g,'-').replace(/-{2,}/g,'-').replace(/^[\s._-]+|[\s._-]+$/g,''); if(!/\.pdf$/i.test(result))result+='.PDF'; return result||'Unmatched.PDF';}

function claimZbomOutputName(baseName,quote){
  let name=baseName,key=name.toLowerCase(),owner=zbomOutputClaims.get(key);
  if(owner&&owner!==quote){name=appendQuoteToName(name,quote);key=name.toLowerCase();let suffix=2; while(zbomOutputClaims.has(key)&&zbomOutputClaims.get(key)!==quote){name=appendQuoteToName(baseName,quote+'_'+suffix++);key=name.toLowerCase()}}
  zbomOutputClaims.set(key,quote);return name;
}

function zbomFilePath(name){return ZBOM_PDF_FOLDER+'\\'+name}

async function moveZbomFile(descriptor,name){
  const options={sourcePath:descriptor.path||zbomFilePath(descriptor.name),destinationPath:zbomFilePath(name),overwrite:true};let error;
  for(let attempt=0;attempt<60;attempt++)try{return await bridgeFileRequest('move',options,BRIDGE_FILE_TIMEOUT)}catch(current){error=current;if(attempt<59)await waitMs(500)}
  throw error;
}

async function collectZbomFiles(files,version){
  if(!files.length){addZbomLog('Version '+version+' created no new PDF. This is normal when that print path does not apply.');return 0}
  let moved=0;
  for(const descriptor of files){
    const row=matchZbomRow(descriptor.name),metadata=rawZbomMetadata(descriptor.name);
    if(!row){const name='UNMATCHED_'+safeUnmatchedName(descriptor.name); await moveZbomFile(descriptor,name); addZbomLog('Could not match '+descriptor.name+' to exactly one input Quote; renamed to '+name+'.');moved++;continue;}
    const actual={fid:metadata.fid||row.fid,systemDescription:row.systemDescription,quote:metadata.quote||row.quote};
    if(bobjFidKey(actual.fid))await prepareBobjFileNames('zbom',[actual],zbomTemplateInfo().template);
    const template=zbomTemplateInfo().template,baseName=renderZbomFileName(actual,template),outputName=claimZbomOutputName(baseName,actual.quote); await moveZbomFile(descriptor,outputName); let result=zbomResults.get(row.quote);
    if(!result){result={files:new Set(),notes:[],status:'Exported'};zbomResults.set(row.quote,result)}
    result.files.add(outputName);result.status='Exported';
    if(row.fid&&metadata.fid&&row.fid!==metadata.fid){const note='Input FID '+row.fid+' differs from SAP filename FID '+metadata.fid+'; SAP FID was used.'; if(!result.notes.includes(note))result.notes.push(note);addZbomLog(row.quote+': '+note);}
    addZbomLog('Version '+version+': '+descriptor.name+' → '+outputName);moved++;
  }
  return moved;
}

async function runZbomVersion(version,rows,previousClipboard,startProgress,collectProgress){
  const before=bridgeMock?new Map():await listZbomRawPdfs();let actionError=null,clipboardWritten=false; const quotationText=rows.map(row=>row.quote).join('\r\n'); setZbomProgress(startProgress,'Running SAP Version '+version+'…');addZbomLog('Starting Version '+version+' for '+
    rows.length+' quotation(s).');
  try{
    let data=await sendZbomAction('zbom','transaction',{code:'ZMASSPRINT'});addZbomLog('Version '+version+': opened ZMASSPRINT.'+sapStatusSummary(data)); await sendZbomAction('zbom','setText',{controlId:'wnd[0]/usr/ctxtRG_KSCHL-LOW',value:'ZBOM'});
    await sendZbomAction('zbom','setText',{controlId:'wnd[0]/usr/ctxtP_FILE',value:ZBOM_PDF_FOLDER}); await sendZbomAction('zbom','press',{controlId:'wnd[0]/usr/btn%_RG_VBELN_%_APP_%-VALU_PUSH'}); writeQuotationClipboard(quotationText);clipboardWritten=true;
    await sendZbomAction('zbom','press',{controlId:'wnd[1]/tbar[0]/btn[24]'}); await sendZbomAction('zbom','press',{controlId:'wnd[1]/tbar[0]/btn[8]'}); await sendZbomAction('zbom','setText',{controlId:'wnd[0]/usr/ctxtPM_VERMO',value:String(version)});
    data=await sendZbomAction('zbom','press',{controlId:'wnd[0]/tbar[1]/btn[8]'});addZbomLog('Version '+version+': execute completed.'+sapStatusSummary(data));
    for(const controlId of ['wnd[0]/tbar[1]/btn[5]','wnd[0]/tbar[1]/btn[14]']){data=await sendZbomAction('zbom','pressIfPresent',{controlId:controlId}); if(data.value)addZbomLog('Version '+version+': pressed '+controlId+'.'+sapStatusSummary(data));}
  }catch(error){actionError=error;addZbomLog('Version '+version+' did not complete all SAP steps: '+bridgeError(error));}finally{if(clipboardWritten&&previousClipboard!==null)try{writeQuotationClipboard(previousClipboard)}catch(_){/* ignore */}}
  if(bridgeMock){addZbomLog('Version '+version+': mock mode creates no PDF files.');return {count:0,error:actionError}}
  setZbomProgress(collectProgress,'Checking files created by Version '+version+'…'); const changed=await waitForZbomFiles(before,rows),count=await collectZbomFiles(changed,version); return {count:count,error:actionError};
}

function renderZbomResultSummary(){
  const container=bridgeId('zbomResultSummary');container.replaceChildren(); const valid=zbomRows.filter(row=>!row._error),success=valid.filter(row=>{const result=zbomResults.get(row.quote);return result&&result.files.size}).length,missing=valid.length-success;
  [['Exported: '+success,'ok'],['No PDF: '+missing,missing?'warn':'ok'],['Output: '+ZBOM_PDF_FOLDER,'']].forEach(item=>{const pill=document.createElement('span');pill.className='result-pill'+(item[1]?' '+item[1]:'');pill.textContent=item[0];container.appendChild(pill);});
  renderZbomPreview();return {success:success,missing:missing};
}

function setZbomRunning(value){
  zbomRunning=Boolean(value); ['zbomInput','zbomTemplate','zbomOpenHistory','zbomResetInput','zbomResetTemplate'].forEach(id=>{const element=bridgeId(id);if(element)element.disabled=zbomRunning});
  document.querySelectorAll('[data-zbom-token]').forEach(button=>button.disabled=zbomRunning); setBridgeBusy(bridgeBusy);
}

async function startZbomExport(){
  await ensureBobjDataSource();previewZbomInput(); const validRows=zbomRows.filter(row=>!row._error),templateInfo=zbomTemplateInfo();
  if(!validRows.length||zbomErrors.length||templateInfo.error){setZbomProgress(0,'Fix the detected data or file-name template before export.');return}
  if(!selectedBridgeSession('zbom')){setZbomProgress(0,'Select an available SAP GUI session first.');return}
  if(validRows.length>100&&!unsafeWindow.confirm('You are about to print '+validRows.length+' quotations through SAP. Continue?'))return;
  await acquireSapTask('zbom'); recordZbomHistory(bridgeId('zbomInput').value); const previousClipboardPromise=readClipboardBestEffort(); zbomHasRun=false;setZbomRunning(true);zbomLogLines=[];zbomOutputClaims=new Map();zbomResults=new Map();
    bridgeId('zbomResultSummary').replaceChildren();renderSapLog(); setZbomProgress(2,'Preparing the ZBOM PDF export…');
  try{
    addZbomLog('Output folder: '+ZBOM_PDF_FOLDER+' (managed by SAPGUIcontrol).'); const previousClipboard=await previousClipboardPromise; const version2=await runZbomVersion(2,validRows,previousClipboard,10,42);
    const remaining=validRows.filter(row=>!zbomResults.get(row.quote)?.files.size); const version1=remaining.length?await runZbomVersion(1,remaining,previousClipboard,52,84):{count:0,error:null};
    if(!remaining.length)addZbomLog('All preview quotations have PDFs. Version 1 was not needed.');
    sapTools.zbom.job.keep=Boolean(version1.error||version2.error); const summary=renderZbomResultSummary(); setZbomProgress(100,summary.missing?'Completed with '+summary.missing+' quotation(s) to review.':'Export completed successfully.'); zbomHasRun=true;
    addZbomLog('Completed: '+summary.success+' quotation(s) exported, '+summary.missing+' produced no PDF.');
    if(version1.error&&version2.error)addZbomLog('Both SAP paths reported an action error. Review the SAP messages and the rows with no PDF.');
  }catch(error){setZbomProgress(100,'Export stopped: '+bridgeError(error));zbomHasRun=true;addZbomLog('Export stopped: '+bridgeError(error));setBridgeBadge('Action failed','error');}finally{const previousClipboard=await previousClipboardPromise; if(previousClipboard!==
    null)try{writeQuotationClipboard(previousClipboard)}catch(_){/* ignore */} setZbomRunning(false);}
}

async function runBridgeAction(operation){
  return runSapTool('debug',async()=>{
    const options={};
    if(operation==='transaction')options.code=bridgeId('sapTransaction').value.trim();
    else if(operation==='sendVKey'){options.controlId=bridgeId('sapVKeyTarget').value.trim()||'wnd[0]';options.key=Number(bridgeId('sapVKey').value)} else{options.controlId=bridgeId('sapControlId').value.trim();options.value=bridgeId('sapControlValue').value}
    await acquireSapTask('debug'); const data=await sendZbomAction('debug',operation,options);
    if(operation==='readText'&&Object.prototype.hasOwnProperty.call(data,'value'))bridgeId('sapControlValue').value=data.value==null?'':String(data.value);
    sapTools.debug.job.keep=false;showBridgeResult(data,'SAP action completed.');
  });
}

function bindBridgeEvents(){
  unsafeWindow.KyraBridge={runPowerShell:runBridgePowerShell,runVba:runBridgeVba,openFolder:openBridgeFolder,activateWindow:handle=>bridgeRequest('POST','/windows/activate',{handle:handle},20000)}; bridgeId('nativeRun').addEventListener('click',runNativeScriptFromEditor);
  bridgeId('nativeLanguage').addEventListener('change',()=>{bridgeId('nativeEntryPointField').hidden=bridgeId('nativeLanguage').value!=='vba'}); bridgeId('bridgeRefresh').addEventListener('click',refreshBridge); bridgeId('bobjUpdate').addEventListener('click',startBobjUpdate);
  bridgeId('bobjOpen').addEventListener('click',openBobjData); bridgeId('ctlBobjUpdate').addEventListener('click',startCtlBobjUpdate); bridgeId('ctlBobjOpen').addEventListener('click',openCtlBobjData);
  bridgeId('sapSession').addEventListener('change',()=>{sapTools[activeSapTool].selection=bridgeId('sapSession').value;setBridgeBusy(bridgeBusy)}); bridgeId('sapReleaseSession').addEventListener('click',releaseSapSession);
  window.addEventListener('beforeunload',event=>{if(Object.keys(sapTools).some(sapToolRunning)){event.preventDefault();event.returnValue=''}});
  document.querySelectorAll('[data-tool-tab]').forEach(button=>button.addEventListener('click',()=>{setToolTab(button.dataset.toolTab);if(button.dataset.toolTab!=='debug')ensureBobjForCurrentInput()}));
  bridgeId('sapRunTransaction').addEventListener('click',()=>runBridgeAction('transaction')); bridgeId('sapReadText').addEventListener('click',()=>runBridgeAction('readText')); bridgeId('sapSetText').addEventListener('click',()=>runBridgeAction('setText'));
  bridgeId('sapSetKey').addEventListener('click',()=>runBridgeAction('setKey')); bridgeId('sapPress').addEventListener('click',()=>runBridgeAction('press')); bridgeId('sapSelect').addEventListener('click',()=>runBridgeAction('select'));
  bridgeId('sapSendVKey').addEventListener('click',()=>runBridgeAction('sendVKey')); bridgeId('zbomInput').addEventListener('input',()=>{saveZbomInput();previewZbomInput();ensureBobjForCurrentInput()});
  bridgeId('zbomTemplate').addEventListener('input',()=>{saveZbomTemplate();zbomHasRun=false;zbomResults=new Map();bridgeId('zbomResultSummary').replaceChildren();renderZbomPreview();ensureBobjForCurrentInput()});
  document.querySelectorAll('[data-zbom-token]').forEach(button=>button.addEventListener('click',()=>insertZbomToken(button.dataset.zbomToken)));
  bridgeId('createQuoteQuotation').addEventListener('input',()=>{sanitizeCreateQuoteQuotation();saveCreateQuoteQuotation();updateCreateQuoteStart()});
  bridgeId('createQuoteInput').addEventListener('input',()=>{saveCreateQuoteInput();previewCreateQuoteInput();ensureBobjForCurrentInput()});
  bridgeId('createQuoteDefaultValidTo').addEventListener('input',()=>{sanitizeCreateQuoteDays();saveCreateQuoteValidToDays();renderCreateQuotePreview()}); bridgeId('createQuoteFidSuffix').addEventListener('input',()=>{saveCreateQuoteFidSuffix();renderCreateQuotePreview()});
  bridgeId('zbomCompareInput').addEventListener('input',()=>{saveZbomCompareInput();previewZbomCompareInput();ensureBobjForCurrentInput()});
  bridgeId('zbomCompareTemplate').addEventListener('input',()=>{saveZbomCompareTemplate();invalidateZbomCompareResults();renderZbomComparePreview();ensureBobjForCurrentInput()});
  bridgeId('zbomCompareSeparator').addEventListener('input',()=>{saveZbomCompareSeparator();invalidateZbomCompareResults();renderZbomComparePreview()});
  bridgeId('zbomCompareOutputFolder').addEventListener('input',()=>{saveZbomCompareFolder();invalidateZbomCompareResults();renderZbomComparePreview();updateZbomCompareStart()});
  bridgeId('zbomCompareKeepTemp').addEventListener('click',()=>{const button=bridgeId('zbomCompareKeepTemp'),keep=!keepZbomCompareTemp();button.setAttribute('aria-pressed',keep?'true':'false');button.title=
    keep?'The temp folder will be kept.':'The temp folder will be deleted after Compare finishes.';});
  document.querySelectorAll('[data-zbom-compare-token]').forEach(button=>button.addEventListener('click',()=>insertZbomCompareToken(button.dataset.zbomCompareToken))); bridgeId('zbomCompareOpenFolder').addEventListener('click',openZbomCompareOutputFolder);
  bridgeId('zbomCompareOpenHistory').addEventListener('click',openZbomCompareHistory); bridgeId('zbomCompareCloseHistory').addEventListener('click',closeZbomCompareHistory); document.querySelector('[data-zbom-compare-history-close]').addEventListener('click',closeZbomCompareHistory);
  bridgeId('createQuoteOpenHistory').addEventListener('click',openCreateQuoteHistory); bridgeId('createQuoteCloseHistory').addEventListener('click',closeCreateQuoteHistory); document.querySelector('[data-create-quote-history-close]').addEventListener('click',closeCreateQuoteHistory);
  bridgeId('zbomOpenHistory').addEventListener('click',openZbomHistory); bridgeId('zbomCloseHistory').addEventListener('click',closeZbomHistory); document.querySelector('[data-zbom-history-close]').addEventListener('click',closeZbomHistory);
  document.addEventListener('keydown',event=>{if(event.key!=='Escape')return;if(!bridgeId('zbomHistoryDialog').hidden)closeZbomHistory();if(!bridgeId('zbomCompareHistoryDialog').hidden)closeZbomCompareHistory();
    if(!bridgeId('createQuoteHistoryDialog').hidden)closeCreateQuoteHistory()});
  document.addEventListener('keydown',event=>{if(event.ctrlKey&&event.altKey&&event.shiftKey&&event.code==='KeyD'){event.preventDefault();bridgeDebugUnlocked=!bridgeDebugUnlocked;writePersistent(BRIDGE_DEBUG_STORAGE_KEY,bridgeDebugUnlocked?'1':'0'); bridgeId('tabDebug').hidden=
    !bridgeDebugUnlocked;setToolTab(bridgeDebugUnlocked?'debug':'zbom');}});
  bridgeId('zbomResetInput').addEventListener('click',()=>{const input=bridgeId('zbomInput');input.value='';saveZbomInput();previewZbomInput();input.focus();});
  bridgeId('zbomResetTemplate').addEventListener('click',()=>{const input=bridgeId('zbomTemplate');input.value=ZBOM_DEFAULT_TEMPLATE;saveZbomTemplate();zbomHasRun=false;zbomResults=new Map();bridgeId('zbomResultSummary').replaceChildren();renderZbomPreview();input.focus();});
  bridgeId('zbomCompareResetInput').addEventListener('click',()=>{const input=bridgeId('zbomCompareInput');input.value='';saveZbomCompareInput();previewZbomCompareInput();input.focus();});
  bridgeId('zbomCompareResetTemplate').addEventListener('click',()=>{
    const input=bridgeId('zbomCompareTemplate'),separator=bridgeId('zbomCompareSeparator'),folder=bridgeId('zbomCompareOutputFolder');input.value=ZBOM_COMPARE_DEFAULT_TEMPLATE;separator.value=ZBOM_COMPARE_DEFAULT_SEPARATOR;folder.value=ZBOM_COMPARE_DEFAULT_FOLDER;
      saveZbomCompareTemplate();saveZbomCompareSeparator();saveZbomCompareFolder();invalidateZbomCompareResults();renderZbomComparePreview();updateZbomCompareStart();input.focus();
  });
  bridgeId('createQuoteResetInput').addEventListener('click',()=>{const input=bridgeId('createQuoteInput');input.value='';createQuoteCrdOverrides=new Map();createQuoteManualCrd=new Set();saveCreateQuoteInput();saveCreateQuoteCrdOverrides();previewCreateQuoteInput();
    setCreateQuoteStatus('Ready.');input.focus();});
  bridgeId('createQuoteResetSettings').addEventListener('click',()=>{
    const validToDays=bridgeId('createQuoteDefaultValidTo'),fidSuffix=bridgeId('createQuoteFidSuffix');validToDays.value=CREATE_QUOTE_DEFAULT_VALID_TO_DAYS;fidSuffix.value=CREATE_QUOTE_DEFAULT_FID_SUFFIX;saveCreateQuoteValidToDays();saveCreateQuoteFidSuffix();
      renderCreateQuotePreview();validToDays.focus();
  });
  bridgeId('zbomStart').addEventListener('click',()=>runSapTool('zbom',startZbomExport)); bridgeId('zbomOpenFolder').addEventListener('click',openZbomFolder); bridgeId('zbomCompareStart').addEventListener('click',()=>runSapTool('zbomCompare',startZbomCompare));
  bridgeId('createQuoteStart').addEventListener('click',()=>runSapTool('createQuote',startCreateQuote)); bridgeDebugUnlocked=readPersistent(BRIDGE_DEBUG_STORAGE_KEY,'0')==='1';bridgeId('tabDebug').hidden=!bridgeDebugUnlocked; loadZbomPreferences();previewZbomInput();
    loadZbomComparePreferences();previewZbomCompareInput();loadCreateQuotePreferences();previewCreateQuoteInput();setToolTab('zbom');
}

function runFcidGovernance(){
'use strict';
if(toolkitExpired())return;
if(document.documentElement.dataset.fcidMounted)return;
document.documentElement.dataset.fcidMounted='1';document.documentElement.lang='en'; const id=value=>document.getElementById(value); const text=value=>String(value??'').trim();
const escapeHtml=value=>String(value??'').replace(/[&<>"']/g,character=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character])); const field=(label,aliases=[],required=false)=>({label,aliases:[label,...aliases],required});
const ctlFields={
  fid:field('Opportunity ID',['FCID','FID'],true),number:field('CTL ID',['CTL#','CTL Number'],true), change:field('Change Cat',['Change Type','Change Category','CTL Change Type'],true),status:field('Status',['CTL Status']), mcsd:field('MCSD @ Current',['MCSD','Current MCSD']),
    srd:field('SRD @ Current',['Sales Ops RD','Current SRD','SRD']), slotSrd:field('SRD @ SOPs Approval',['SRD at SOPs Approval']),approved:field('1st SOps Appr Dt',['Sales Ops Approval Date','SOPS Approval Date']), created:field('Create Date',['Creation Date']),
    modified:field('Last Modified Date',['Modified Date']), summary:field('Change Summary',['Summary']),mfg:field('Mfg Upgrade Type',['Manufacturing Upgrade Type'])
};
const sources={
  fcid:{label:'FCID Database',sheet:'BOBJ',fields:{
    fid:field('FID',['FCID','Field ID'],true),commit:field('Commit: Field',['Commit Field','Field Commit','Commit Status'],true), calendarQuarter:field('Time: Calendar Quarter',['Calendar Quarter','Cal Quarter','Quarter']),
    account:field('Account Group',['Account Team','Sales Territory','AT']),fab:field('Fab Name',['Fab']),fabSegment:field('Fab Segment',['Fab_Segment']), customer:field('Customer',['Customer Name','End Customer']),product:field('Primary Product',['Product','Product Name',
      'FID Description']), system:field('System Description',['System Desc','System']),quote:field('Active Quote Number',['Quote Number','Active Quote']), targetShip:field('Date: Target Ship',['Target Ship Date','Target Ship']),mcsd:field('Date: Mfg Commit Ship',['Mfg Commit Ship',
      'Manufacturing Commit Ship Date','MCSD']), rd:field('Date: Sales Ops Request Date (RD)',['Sales Ops Request Date','Sales Ops RD','RD']),mfg:ctlFields.mfg
  }},
  ctl:{label:'CTL BOBJ',sheet:'CTL BOBJ',fields:ctlFields},
  ctlNsr:{label:'CTL-NSR List',sheet:'export',fields:{...ctlFields, fid:field('FCID',['FID','Opportunity ID'],true),change:field('Change Type',['Change Cat','CTL Change Type'],true), status:field('CTL Status',['Status']),mcsd:field('MCSD',['MCSD @ Current']),srd:field('Sales Ops RD',
    ['SRD','SRD @ Current','Field CRD']), approved:field('Sales Ops Approval Date',['1st SOps Appr Dt','SOPS Approval Date']),stage:field('Stage',['CTL Stage']), quote:field('Quote Number',['Quote ID']),sourceQuote:field('Source Quote',['Source Quote Number'])
  }},
  nsr:{label:'NSR Log',sheet:'Whole list',fields:{
    fid:field('Reference Fcid',['REF FCID','Reference FID','FCID','FID'],true),number:field('NSR#',['NSR #','NSR Number','NSR No'],true), status:field('NSR Status',['Status']),workflow:field('Workflow Status',['Workflow']),task:field('Task Name',['Workflow Task']),
    title:field('NSR Title',['Title']),leadTime:field('NSR LT',['NSR_LT','LT','Lead Time']),submitted:field('Submit Date',['Submitted Date','Submission Date']), customer:field('Customer',['Customer Name']),product:field('Primary Product',['Product']),tempId:field('Temp Id#',
      ['Temp ID','TempId'])
  }},
  nce:{label:'NCE / FA Review',sheet:'Export',fields:{
    fid:field('FCID',['FID'],true),number:field('NCe #',['NCE Number','NCE#'],true),state:field('NCe State',['State','Status']), impact:field('Impact',['NCE Impact']),title:field('Title',['NCE Title']),opened:field('Open Date',['Opened Date','Creation Date']),
    damageGroup:field('Damage Group'),damageCode:field('Damage Code'),partNumber:field('Part_Number',['Part Number']), partDescription:field('Part_Description',['Part Description']),customer:field('Customer',['Customer Name'])
  }}
};

sources.ctl.label='CTL Report'; const BOBJ_SOURCES={fcid:{label:'FID BOBJ',endpoint:'/excel',file:'BOBJ_FID.xlsm'}, ctl:{label:'CTL BOBJ',endpoint:'/excel/ctl',file:'BOBJ_CTL.xlsm'}}; const HOSTED_ARCHIVE_URL='https://raw.githubusercontent.com/synalocey/ToolkitKyra/main/Stacey1.zip';
const HOSTED_PASSWORD_URL='/sap/opu/odata/sap/ZNSR_CD_SRV/GetNameDateTimeSet/?$filter=Userid%20eq%20%27LONGKY%27',HOSTED_SOURCE_PREFIXES={ctlNsr:'c',nce:'n'};
const SOURCE_DATES={
  fcid:{fields:['Slot Actual'],label:'Slot Actual'}, ctl:{fields:['Creation Date','Create Date'],label:'Creation Date'}, ctlNsr:{fields:['Create Date','Creation Date'],label:'Create Date'},
  nsr:{fields:['Submit Date','Submitted Date','Submission Date'],label:'Submit Date'}, nce:{fields:['Open Date'],label:'Open Date'}
};
const CACHE_PREFIX='fcid-governance:'; const store={raw:Object.fromEntries(Object.keys(sources).map(key=>[key,[]])),meta:{},busy:new Set(),refreshing:false,runs:{},controllers:{},messages:{}};
const nsrData=createNsrDataService((range,count)=>sourceMessage('nsr',`Loading SAP · ${range.from} to ${range.to} · ${count.toLocaleString()} rows`,'loading')); let nsrRange=nsrData.sapRange();
const currentYear=new Date().getFullYear(),quarterYears=Array.from({length:5},(_,index)=>currentYear+index-2); const filterDefs=[['fQuarter','Quarter'],['fAccount','Account'],['fInitiator','Initiator'],['fCustomer','Customer'],['fRisk','Status / Risk']];
const riskNames=['Pre Late Risk','No Slot CTL','Post Config CTL Risk','MCSD Pending','NSR Appeal Watch','OOBQ / NCE Signal'];
const css=chosenBaseCss+(`
:root{--bg:#f4f7fb;--panel:#fff;--ink:#0f172a;--muted:#64748b;--line:#d9e3ef;--teal:#0f766e;--blue:#2563eb;--green:#16a34a;--amber:#f59e0b;--red:#dc2626;--purple:#7c3aed;--shadow:0 14px 34px rgba(15,23,42,.08);--r:18px}` +
  `*{box-sizing:border-box}body{margin:0;min-width:1980px;background:var(--bg);color:var(--ink);font-family:Segoe UI,Arial,"Microsoft YaHei",sans-serif}.hero{background:linear-gradient(135deg,#052e44,#075985 54%,#0f766e);` +
  `color:#fff;padding:24px 30px}.hero h1{margin:8px 0 6px;font-size:30px}.hero p{margin:0;color:#dffcf5;max-width:1650px;line-height:1.55}.tag{display:inline-block;border:1px solid #ffffff55;background:#ffffff22;` +
  `border-radius:999px;padding:4px 10px;margin-right:7px;font-size:12px;font-weight:800}.layout{display:grid;grid-template-columns:410px 1fr;gap:16px;max-width:1980px;margin:auto;padding:16px}.side,.card{background:var(--panel);` +
  `border:1px solid var(--line);border-radius:var(--r);box-shadow:var(--shadow);padding:16px}.side{position:sticky;top:14px;align-self:start;max-height:calc(100vh - 28px);overflow:auto}.main{display:grid;gap:16px}` +
  `h2{font-size:18px;margin:0 0 12px}h3{font-size:14px;margin:0 0 8px;color:#334155}.small{font-size:12px;color:var(--muted);line-height:1.55}.upload{border:1.5px dashed #a9bbca;background:#fbfdff;border-radius:14px;` +
  `padding:12px;margin-bottom:14px}.upload label{display:block;font-weight:900;color:#334155;font-size:12px;margin:9px 0 4px}.required-star{color:var(--red);font-size:14px;margin-left:2px}.upload input{width:100%;font-size:12px}` +
  `.btn{border:0;border-radius:11px;background:var(--teal);color:#fff;font-weight:900;padding:10px 12px;cursor:pointer}.btn.light{background:#e2e8f0;color:#0f172a}.btn.blue{background:var(--blue)}` +
  `.btn.warn{background:var(--amber);color:#111827}.btn.red{background:var(--red)}.btnrow{display:flex;gap:8px;flex-wrap:wrap}.fg{margin:10px 0}.fg label{display:block;font-size:11px;color:#64748b;text-transform:uppercase;` +
  `letter-spacing:.4px;font-weight:900;margin-bottom:5px}.fg label[for]{cursor:pointer}.fg select,.fg input{width:100%;padding:8px;border:1px solid var(--line);border-radius:10px;background:#fff}.fg select[multiple]{width:100%;` +
  `max-width:100%;height:92px;min-height:56px;resize:vertical;overflow:auto}.hint{border-left:4px solid var(--teal);background:#ecfeff;color:#164e63;border-radius:10px;padding:10px;font-size:12px;line-height:1.55;` +
  `margin-top:10px}.warnbox{border-left:4px solid var(--amber);background:#fffbeb;color:#78350f;border-radius:10px;padding:10px;font-size:12px;line-height:1.55;margin-top:10px}.msg{border-left:4px solid var(--blue);` +
  `background:#eff6ff;color:#1e3a8a;border-radius:10px;padding:10px;font-size:12px;line-height:1.55;margin-top:10px;max-height:190px;overflow:auto}.tabs{display:flex;gap:8px;flex-wrap:wrap}.tab{border:1px solid var(--line);` +
  `background:#fff;border-radius:999px;padding:9px 13px;font-weight:900;color:#334155;cursor:pointer}.tab.active{background:var(--teal);border-color:var(--teal);color:#fff}.page{display:grid;gap:16px}` +
  `.hidden{display:none!important}.sectionTitle{display:flex;justify-content:space-between;align-items:center;gap:12px}.badge{font-size:11px;color:#fff;background:#134e4a;border-radius:999px;padding:4px 10px;font-weight:900}` +
  `.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.kpi{min-height:116px}.kpi small{display:block;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.55px;font-weight:900}.kpi ` +
  `b{display:block;font-size:28px;margin:8px 0 2px}.kpiJump{cursor:pointer;border-radius:7px;width:max-content;max-width:100%;transition:opacity .15s ease}.kpiJump:hover{text-decoration:underline;text-underline-offset:4px}` +
  `.kpiJump:focus-visible{outline:3px solid #38bdf8;outline-offset:3px}.blue b{color:var(--blue)}.green b{color:var(--green)}.amber b{color:var(--amber)}.red b{color:var(--red)}.purple b{color:var(--purple)}.grid2{display:grid;` +
  `grid-template-columns:1fr 1fr;gap:16px}.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}.grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}.miniGrid{display:grid;` +
  `grid-template-columns:repeat(4,1fr);gap:10px}.mini{background:#f8fafc;border:1px solid #e5e7eb;border-radius:12px;padding:10px}.mini small{display:block;color:#64748b;font-size:11px;text-transform:uppercase;font-weight:900}` +
  `.mini b{font-size:18px}.tableWrap{overflow:auto;border:1px solid var(--line);border-radius:14px;max-height:680px}table{border-collapse:collapse;width:100%;min-width:1500px;font-size:12px}th,td{
    padding:9px 10px;
    border-bottom:1px solid #edf2f7;
    text-align:left;
    vertical-align:top;
}th{background:#f8fafc;color:#334155;position:sticky;top:0;z-index:1}tr:hover{background:#fbfdff}.pill{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border-radius:999px;font-size:11px;font-weight:900}` +
  `.ok{background:#dcfce7;color:#166534}.bad{background:#fee2e2;color:#991b1b}.warnp{background:#fef3c7;color:#92400e}.info{background:#dbeafe;color:#1d4ed8}.purplep{background:#ede9fe;color:#5b21b6}.barRow{display:grid;` +
  `grid-template-columns:205px 1fr 70px;gap:9px;align-items:center;margin:10px 0}.bar{height:13px;background:#e2e8f0;border-radius:999px;overflow:hidden}.bar span{display:block;height:100%;border-radius:999px}.heat{display:grid;` +
  `grid-template-columns:160px repeat(5,1fr);border:1px solid var(--line);border-radius:14px;overflow:hidden}.heat div{padding:10px;border-right:1px solid #e5e7eb;border-bottom:1px solid #e5e7eb;font-size:12px}.heat ` +
  `.head{background:#f8fafc;font-weight:900;color:#334155}.heat .g{background:#dcfce7;color:#166534;font-weight:900}.heat .y{background:#fef3c7;color:#92400e;font-weight:900}.heat .r{background:#fee2e2;color:#991b1b;` +
  `font-weight:900}.funnel{display:flex;gap:10px;align-items:stretch;flex-wrap:wrap}.step{border:1px solid var(--line);border-radius:14px;background:#fbfdff;min-width:155px;flex:1;padding:13px;text-align:center}.step ` +
  `b{display:block;font-size:22px;margin:3px 0;color:#0f766e}.arrow{font-weight:900;color:#94a3b8;align-self:center}.code{font-family:Consolas,monospace;background:#f8fafc;border:1px solid #e5e7eb;border-radius:12px;` +
  `padding:10px;white-space:pre-wrap;font-size:12px;color:#334155}.searchBox{display:flex;gap:8px;flex-wrap:wrap}.searchBox input{flex:1;min-width:260px;border:1px solid var(--line);border-radius:11px;padding:10px}` +
  `.footer{font-size:11px;color:#64748b;line-height:1.45;max-width:1980px;margin:auto;padding:0 16px 18px}.sectionFocusPulse{position:relative;isolation:isolate}.sectionFocusPulse::after{content:"";position:absolute;inset:-4px;` +
  `border-radius:calc(var(--r) + 4px);padding:4px;background:linear-gradient(115deg,#14b8a6,#38bdf8,#7c3aed,#f59e0b,#14b8a6);-webkit-mask:linear-gradient(#000 0 0) content-box,linear-gradient(#000 0 0);` +
  `-webkit-mask-composite:xor;mask-composite:exclude;filter:drop-shadow(0 0 10px rgba(37,99,235,.5));pointer-events:none;animation:sectionFocusBorder 5s cubic-bezier(.22,.61,.36,1) forwards}@keyframes ` +
  `sectionFocusBorder{0%{opacity:0}8%{opacity:1}22%{opacity:.95}55%{opacity:.55}100%{opacity:0}}
.btn[href]{display:inline-block;text-decoration:none}.agCtl{font-size:14px}.agGrid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}.agBox{border:1px solid var(--line);border-radius:12px;padding:10px;` +
  `background:#f8fafc}.agBox label{display:block;font-weight:900;color:#334155;margin-bottom:6px}.agBox textarea{width:100%;min-height:165px;resize:vertical;border:1px solid var(--line);border-radius:9px;padding:8px;` +
  `font:12px/1.45 Consolas,monospace;background:#fff}
`)+(`.fc-source{position:relative;border:1px solid #dce7ed;background:#fbfdfe;border-radius:10px;padding:10px 8px;margin-top:9px}.fc-source-head{display:flex;align-items:center;justify-content:space-between;gap:5px;` +
  `margin-bottom:7px}.fc-source h3{display:flex;align-items:center;gap:3px;min-width:0;margin:0;color:#24364b;font-size:12px;line-height:1.35}.fc-source-name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}` +
  `.fc-source-number{display:inline-flex;align-items:center;justify-content:center;flex:none;width:22px;height:22px;background:var(--teal);color:#fff;border-radius:6px;font-size:12px;font-weight:600}.fc-source-bobj{flex:none;border:1px solid #d5e0e8;border-radius:4px;padding:1px 4px;color:#64748b;font-size:9px;font-weight:600}
.fc-source-state{display:inline-flex;align-items:center;gap:5px;flex:none;max-width:105px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#64748b;font-size:10px;line-height:1.5}.fc-source-state.has-range{max-width:none}` +
  `.fc-source-state::before{content:"";display:inline-block;flex:none;width:5px;height:5px;border-radius:50%;background:#94a3b8}.fc-source-state.ok{color:#166534}.fc-source-state.ok::before{background:var(--green)}` +
  `.fc-source-state.loading{color:#92400e}.fc-source-state.loading::before{background:var(--amber)}.fc-source-state.error{color:#991b1b;cursor:help}.fc-source-state.error::before{background:var(--red)}
.fc-source-toolbar{display:flex;align-items:center;gap:6px;min-width:0}.fc-source-file{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#64748b;font-size:10px;line-height:1.5}` +
  `.fc-source-toolbar>.btn,.fc-source-help>.btn{flex:none;padding:4px 6px;font-size:10px;line-height:1.4}.fc-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.fc-source-help{position:relative;flex:none}.fc-source-help>summary{list-style:none;font:inherit}.fc-source-help>summary::-webkit-details-marker{display:none}
.fc-source-help-bubble{position:absolute;z-index:20;right:0;top:calc(100% + 7px);width:230px;max-width:calc(100vw - 48px);padding:9px 10px;border:1px solid #dce7ed;border-radius:8px;background:#fff;color:#475569;box-shadow:0 5px 18px #0f172a1a;font-size:11px;line-height:1.5}
.fc-source-help-bubble::before{content:"";position:absolute;right:7px;top:-5px;width:8px;height:8px;border-top:1px solid #dce7ed;border-left:1px solid #dce7ed;background:#fff;transform:rotate(45deg)}.fc-source-help-bubble a{color:#0f766e}
.fc-nsr-dialog{width:380px;max-width:calc(100vw - 32px);padding:20px;border:1px solid #d5e0e8;border-radius:12px;background:#fff;color:var(--ink);box-shadow:0 18px 60px #0f172a33;font:13px/1.5 var(--font-sans)}
.fc-nsr-dialog::backdrop{background:#0f172a66}.fc-nsr-dialog h2{margin:0 0 16px;font-size:16px}.fc-nsr-dialog .fc-actions{justify-content:flex-end;margin-top:20px}
.fc-nsr-help{margin:12px 0 0;color:#64748b;font-size:11px}.fc-nsr-help a{color:inherit;text-underline-offset:2px}
.fc-range{display:grid;grid-template-columns:1fr 1fr;gap:12px}.fc-range label{font-size:12px;color:#475569}.fc-range input{display:block;width:100%;min-width:0;margin-top:5px;padding:8px;border:1px solid var(--line);border-radius:7px;` +
  `background:#fff;font:inherit;color:#334155}
.side-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin:19px 0 10px;padding-top:17px;border-top:1px solid #e5ecf2}.side-head h2{margin:0}.quick-search{display:block;width:100%;padding:9px 10px;` +
  `border:1px solid var(--line);border-radius:8px;background:#fff;color:var(--ink);font-size:12px}.quick-search::placeholder{color:#8493a5}
.btn{border:1px solid transparent;border-radius:8px;background:var(--teal);color:#fff;font-size:12px;font-weight:600;line-height:1.4;padding:8px 12px;cursor:pointer;box-shadow:0 1px 2px #0f172a0a;transition:background ` +
  `.15s,border-color .15s}.fc-side .btn:hover:not(:disabled){background:#0d665f}.fc-side .btn.ghost{border-color:#ffffff66;background:#ffffff18;flex:none}.fc-side .btn.ghost:hover:not(:disabled){background:#ffffff28}.fc-side ` +
  `.btn.light{border-color:#d5e0e8;background:#fff;color:#334155;box-shadow:0 1px 2px #0f172a05}.fc-side .btn.light:hover:not(:disabled){border-color:#b9cbd8;background:#f1f6f9}.fc-side .btn.blue{background:var(--blue)}.fc-side ` +
  `.btn.blue:hover:not(:disabled){background:#1d4ed8}.fc-side .btn.warn{border-color:#f2d28b;background:#fff6dc;color:#8f4f00}.fc-side .btn.warn:hover:not(:disabled){background:#ffedb8}.fc-side .btn.red{background:var(--red)}` +
  `.fc-side .btn.small{padding:5px 8px;font-size:11px}
.fc-filter{margin-top:14px}.fc-filter>label{display:block;font-size:11px;color:#475569;font-weight:600;margin-bottom:5px}.fc-filter select{width:100%;height:80px;min-height:56px;resize:vertical;border:1px solid var(--line);` +
  `border-radius:8px;background:#fff;padding:6px}.fc-filter .fc-actions{justify-content:flex-end;margin-top:5px}.chosen-container{font-family:var(--font-sans);font-size:12px;width:100%!important;max-width:100%}` +
  `.chosen-container-multi .chosen-choices{min-height:38px;border:1px solid var(--line);border-radius:8px;background:#fff;background-image:none;box-shadow:none;padding:3px 5px}.chosen-container-active ` +
  `.chosen-choices{border-color:#38bdf8;box-shadow:0 0 0 3px #38bdf822}.chosen-container-multi .chosen-choices li.search-choice{border:1px solid #b5dcd6;border-radius:5px;background:#e6f5f1;background-image:none;color:#115e59;` +
  `box-shadow:none}.fc-filter .chosen-container-multi .chosen-choices li.search-field input.chosen-search-input{font-size:12px}
.fc-years{display:flex;align-items:center;gap:4px;white-space:nowrap;margin-top:6px}.fc-years span{font-size:9px;color:var(--muted);font-weight:600}.fc-years button{flex:none;border:1px solid #dce5ed;border-radius:6px;` +
  `background:#f7f9fb;color:#475569;padding:3px 5px;font-size:10px;font-weight:600}.fc-years button:hover:not(:disabled){border-color:#8ebdbb;color:#0f766e}.fc-years button.active{border-color:#83bcb6;background:#e6f4f1;` +
  `color:#115e59}.fc-years button.partial{border-color:#a7cbe0;background:#edf6fc;color:#075985}
`)+(`
body.fcid-app{min-width:0;font-size:13px}.hero{padding:23px 28px}.hero-inner{max-width:1900px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:20px}.hero-inner>div{min-width:0;flex:1}.hero ` +
  `.tag{font-size:10px;font-weight:600}.hero h1{font-size:28px}.hero p{font-size:12px}.hero .fc-refresh{flex:none;border:1px solid #ffffff66;background:#ffffff18;font-weight:600;font-size:12px;padding:8px 12px}` +
  `.layout{grid-template-columns:340px minmax(0,1fr);max-width:1900px}.fc-side{position:static;max-height:none;overflow:visible;padding:16px;align-self:start}.main{min-width:0;align-content:start}.main>.card,.page>.card,.grid2>` +
  `.card{min-width:0}.main .kpis{grid-template-columns:repeat(4,minmax(0,1fr))}.main .miniGrid{grid-template-columns:repeat(4,minmax(0,1fr))}.main table{min-width:1000px}.main ` +
  `.grid2{grid-template-columns:repeat(2,minmax(0,1fr))}.main .grid2 .tableWrap{min-width:0}.main .heat{grid-template-columns:110px repeat(5,minmax(45px,1fr));overflow:auto}.main .heat div{font-size:11px;padding:8px}.main ` +
  `.barRow{grid-template-columns:155px minmax(0,1fr) 42px;font-size:11px}.main .card{box-shadow:0 2px 6px #0f172a07,0 10px 28px #0f172a09}.main .tab,.main .btn{font-weight:600;font-size:12px}.fc-side h2{font-size:16px}.fc-side ` +
  `button,.fc-side input,.fc-side select{font:inherit}.fc-side .fc-years button{font-size:10px}.fc-side .btn{font-size:12px}.fc-side .btn.small{font-size:11px}.fc-side .chosen-container{font-size:12px}.fc-side ` +
  `[hidden]{display:none!important}button:disabled,input:disabled{cursor:not-allowed;opacity:.48}.main .agGrid{grid-template-columns:repeat(3,minmax(0,1fr))}.main .agBox{min-width:0}.footer{max-width:1900px}` +
  `.fc-note:empty{display:none}
@media(max-width:1250px){.layout{grid-template-columns:300px minmax(0,1fr)}.main .grid2{grid-template-columns:minmax(0,1fr)}.main .agGrid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:800px){.layout{grid-template-columns:minmax(0,1fr);padding:10px}.hero{padding:20px 16px}.hero h1{font-size:22px}.main .kpis,.main .miniGrid,.main .grid3,.main ` +
  `.grid4{grid-template-columns:repeat(2,minmax(0,1fr))}.main .heat{grid-template-columns:100px repeat(5,55px)}.main .heatmap-scroll{overflow:auto}.main .agGrid{grid-template-columns:minmax(0,1fr)}.main .searchBox ` +
  `input{min-width:0}}
`);
const sourceHtml=Object.entries(sources).map(([key,source],index)=>{
  const help=BOBJ_SOURCES[key]?`Open SAPGUIcontrol.exe, then try again. <a href="${SAPGUICONTROL_DOWNLOAD_URL}" download="SAPGUIcontrol.zip" target="_blank" rel="noopener noreferrer">Download SAPGUIcontrol</a>.`:key==='ctlNsr'?'Download from the CTL system: <a href="https://pcazcrpap02.fremont.lamrc.net:8443/sap(bD1lbiZjPTI1MCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm" target="_blank" rel="noopener noreferrer">Open download page</a>.':key==='nce'?'Ask Stacey for the latest NCE / FA Review file.':'';
  return `<section class="fc-source" aria-labelledby="fcSourceTitle-${key}"><div class="fc-source-head"><h3 id="fcSourceTitle-${key}"><span class="fc-source-number">${index+
    1}</span><span class="fc-source-name">${escapeHtml(source.label)}</span>${BOBJ_SOURCES[key]?'<small class="fc-source-bobj">BOBJ</small>':''}</h3><span class="fc-source-state" id="fcSource-${key}" role="status">Not loaded</span></div>
    <div class="fc-source-toolbar"><span class="fc-source-file" id="fcMeta-${key}"></span>
      ${key==='fcid'?'<button class="btn small" data-source-action="bobj-update" aria-label="Update BOBJ">Update</button>':''}
      ${key==='ctl'?'<button class="btn small" data-source-action="ctl-bobj-update" aria-label="Update CTL BOBJ">Update</button>':''}
      ${key==='nsr'?'<button class="btn small" data-source-action="nsr-open">Load NSR</button>':''}
      <button id="fcImport-${key}" class="btn light small" data-import-source="${key}">Import</button>
      ${help?`<details class="fc-source-help" id="fcHelp-${key}"${BOBJ_SOURCES[key]?' hidden':''}><summary class="btn light small" aria-label="Download help for ${escapeHtml(source.label)}">...</summary><div class="fc-source-help-bubble" role="note">${help}</div></details>`:''}
    </div><input id="fcFile-${key}" data-source-file="${key}" type="file" aria-label="Import ${escapeHtml(source.label)}" accept=".xlsx,.xls,.xlsm,.csv${key==='nsr'?',.json':''}" hidden>
    ${key==='nsr'?'<dialog id="fcNsrDialog" class="fc-nsr-dialog" aria-labelledby="fcNsrDialogTitle"><form id="fcNsrForm"><h2 id="fcNsrDialogTitle">Load NSR</h2><div class="fc-range"><label>From<input id="fcNsrFrom" type="date" required></label><label>To<input id="fcNsrTo" type="date" required></label></div><p class="fc-nsr-help">Manual download: <a href="https://fep.lamresearch.com/flp#ZNSR-display&amp;/Search/Search/%20" target="_blank" rel="noopener noreferrer">Open NSR search</a></p><div class="fc-actions"><button id="fcNsrCancel" class="btn light" type="button">Cancel</button><button id="fcNsrLoad" class="btn" type="submit">Load</button></div></form></dialog>':''}</section>`;
}).join('');

const filterHtml=filterDefs.map(([name,label])=>`<div class="fc-filter"><label for="${name}">${label}</label><select id="${name}" multiple>${name==='fRisk'?riskNames.map(value=>'<option>'+value+'</option>').join(''):''}</select>${name===
  'fQuarter'?'<div class="fc-years"><span>Select Year</span>'+quarterYears.map(year=>'<button type="button" data-quarter-year="'+ year+'" aria-pressed="false">'+year+'</button>').join('')+
  '</div>':''}<div class="fc-actions">${name==='fAccount'?'<button class="btn light small" type="button" data-reference-click="28" aria-controls="fabAccountMapping" title="Fab / Account Group Mapping">Rule</button>':''}<button class="btn light small" data-select-all="${name}">All</button><button class="btn light small" data-clear-filter="${name}">Clear</button></div></div>`).join('');
const charset=document.createElement('meta');charset.charset='utf-8'; const viewport=document.createElement('meta');viewport.name='viewport';viewport.content='width=device-width,initial-scale=1'; const title=document.createElement('title');title.textContent=
  'FCID Governance Control Tower'; const style=document.createElement('style'); style.textContent=css; document.head.replaceChildren(charset,viewport,
  title,emojiFavicon('🏭'),style);document.body.className='fcid-app';
document.body.innerHTML=`<header class="hero"><div class="hero-inner"><div>`+(
  `<h1>🏭 FCID Governance Control Tower</h1><p>Quarter priority = MCSD → Target Ship → Sales Ops RD → Calendar Quarter. Governance population excludes NCFUPG, Software, SW-UPGRADE, ` +
  `and Cust_Bus_Flex. All exports and KPIs follow the current multi-select filters.</p>`)+
    (`</div><button id="fcRefresh" class="btn fc-refresh" type="button" title="Reset all data sources and reload the default sources" disabled>Refresh</button></div></header><div class="layout"><aside class="side fc-side"><h2>` +
  `Data Sources</h2>`)+
    sourceHtml+`<div id="fcActionStatus" class="fc-note" role="status"></div><div id="loadStatus" hidden></div><div class="side-head"><h2>Filters</h2><button id="fcReset" class="btn light small" type="button">Reset Filters</button></div>`+
    filterHtml+`<div class="fc-filter"><label for="fSearch">Search FCID / CTL / NSR</label><input id="fSearch" class="quick-search" placeholder="e.g. 257020 / 8000295092 / N515783"></div></aside>`+(
  `<main class="main"><section class="card"><div class="tabs"><button class="tab active" data-reference-click="0">Executive</button><button class="tab" data-reference-click="1">Stability Engine</button><button class="tab" ` +
  `data-reference-click="2">Slot Governance</button><button class="tab" data-page="Drivers" data-reference-click="3">Risk Drivers</button><button class="tab" data-page="NSR"\ndata-reference-click="4">\nNSR Governance\n</button>` +
  `<button class="tab" data-reference-click="5">Quarter Governance</button><button class="tab" data-reference-click="6">FCID 360°</button><button class="tab" data-reference-click="7">Mapping Diagnostics</button><button ` +
  `class="tab" data-reference-click="8">UAT Console</button><button class="tab" data-page="Rules" data-reference-click="9">Rules</button></div></section>\n<section class="card"><div class="sectionTitle"><h2>Export Center</h2><span class="badge">` +
  `Current Filter Only</span></div><div class="btnrow"><button class="btn blue" data-reference-click="10">Export Population</button><button class="btn blue" data-reference-click="11">Export Excluded</button><button class="btn ` +
  `blue" data-reference-click="12">Export Stability</button><button class="btn blue" data-reference-click="13">Export Slot</button><button class="btn blue" data-reference-click="14">Export Risk</button><button class="btn blue" ` +
  `data-reference-click="15">Export Quarter</button><button class="btn warn" data-reference-click="16">Export All Tabs</button></div></section>\n<section class="card" id="filterContextCard"><div class="sectionTitle"><h2>Current ` +
  `Population Context</h2><span class="badge" id="filterContextBadge">All Population</span></div><div class="miniGrid" id="filterContextMini"></div></section>\n<section id="pageExecutive" class="page"><div class="card"><div ` +
  `class="sectionTitle"><h2>Executive Command Center</h2><span class="badge">Filtered KPI</span></div><div class="kpis"><div class="card kpi blue"><small>Commit FCID</small><b id="kPop">0</b><div class="small">Filtered Commit ` +
  `Population</div></div><div class="card kpi green"><small>Projected Pre-Slot</small><b id="kPre">0%</b><div class="small">On-time slot CTL</div></div><div class="card kpi green"><small>Projected Post-Slot</small><b id="kPost">` +
  `0%</b><div class="small">No config CTL in 20w window</div></div><div class="card kpi red"><small>Risk FCID</small><b id="kRisk" class="kpiJump" role="button" tabindex="0" aria-label="Open Risk Driver Detail" ` +
  `data-reference-click="17" data-reference-keydown="18">0</b><div class="small">P1/P2/P3 distinct</div></div><div class="card kpi green"><small>Slot Readiness</small><b id="kSlotReady">0%</b><div class="small">Slot Ready / ` +
  `Governance Population</div></div><div class="card kpi purple"><small>Open NSR Qty</small><b id="kNSRWatch" class="kpiJump" role="button" tabindex="0" aria-label="Open NSR Population" data-reference-click="19" ` +
  `data-reference-keydown="20">0</b><div class="small">Management Attention Items</div></div><div class="card kpi blue">\n<small>FCIDs Under Open NSR</small>\n<b id="kNSRImpact" class="kpiJump" role="button" tabindex="0" ` +
  `aria-label="Open NSR Population" data-reference-click="19" data-reference-keydown="20">0</b>\n<div class="small">\nDistinct FCIDs with Open NSRs\n</div>\n</div></div></div><div class="card"><h2>Governance ` +
  `Funnel</h2><div id="funnel" class="funnel"></div></div><div class="grid2"><div class="card"><h2>Quarter Stability Heatmap</h2><div id="heatmap" class="heat"></div></div><div class="card"><h2>Top Risk Drivers</h2><div ` +
  `id="driverBars"></div></div></div><div class="card"><h2>Executive Watchlist</h2><div class="tableWrap"><table id="watchTbl"></table></div></div></section>\n<section id="pageStability" class="page hidden"><div class="grid2">` +
  `<div class="card"><h2>Official-aligned Projection Logic</h2><div class="code">Pre-Slot Projection = FIDs with slot CTL submitted on time / filtered governance population\\nPost-Slot Projection = FIDs without qualifying config ` +
  `CTL inside 20w to MCSD / filtered governance population\\nQuarter = derived by MCSD, Target Ship, RD, Calendar Quarter priority.</div></div><div class="card"><h2>Projection Summary</h2><div id="projectionMini" ` +
  `class="miniGrid"></div></div></div><div class="card"><h2>Stability Detail by FCID</h2><div class="tableWrap"><table id="stabilityTbl"></table></div></div></section>\n<section id="pageSlot" class="page hidden"><div ` +
  `class="grid3"><div class="card kpi green"><small>Slot + MCSD</small><b id="kSlotMcsd">0</b></div><div class="card kpi amber"><small>Slot + No MCSD</small><b id="kSlotNoMcsd">0</b></div><div class="card kpi red"><small>No ` +
  `Slot CTL</small><b id="kNoSlotCtl">0</b></div></div><div class="card"><h2>Slot Governance Population</h2><div class="tableWrap"><table id="slotTbl"></table></div></div></section>\n<section id="pageDrivers" class="page hidden">` +
  `<div class="card"><h2>Risk Driver Cards</h2><div id="driverCards" class="grid4"></div></div><div class="card" id="riskDriverDetail"><h2>Risk Driver Detail</h2><div class="tableWrap"><table id="riskTbl"></table></div></div>` +
  `</section>\n<section id="pageNSR"\nclass="page hidden">\n\n<div class="card" id="openNsrPopulation">\n<h2>Open NSR Population</h2>\n<div class="tableWrap">\n<table id="nsrTbl"></table>\n</div>\n` +
  `    </div>\n\n</section><section id="pageQuarter" class="page hidden"><div class="card"><h2>Quarter Governance View</h2><div class="tableWrap"><table id="quarterTbl"></table></div></div><div class="card"><h2>Quarter Source ` +
  `Detail</h2><div class="tableWrap"><table id="quarterSourceTbl"></table></div></div></section>\n<section id="pageFCID" class="page hidden"><div class="card"><h2>FCID 360°</h2><div class="searchBox"><input id="fcidInput" ` +
  `placeholder="输入 FCID / FID"><button class="btn" data-reference-click="21">View FCID</button><button class="btn blue" data-reference-click="22">Export FCID360</button></div></div><div class="grid2"><div class="card"><h2>Basic ` +
  `Info</h2><div id="fcidBasic" class="miniGrid"></div></div><div class="card"><h2>Recommended Action</h2><div id="fcidAction" class="hint">输入 FCID 查看。</div></div></div><div class="card"><h2>FCID360 History & Evidence</h2><div ` +
  `class="tableWrap"><table id="fcidTbl"></table></div></div></section>\n<section id="pageMapping" class="page hidden"><div class="grid2"><div class="card"><h2>Workbook Loading Summary</h2><div class="tableWrap"><table ` +
  `id="loadTbl"></table></div></div><div class="card"><h2>Key Diagnostics</h2><div class="tableWrap"><table id="overlapTbl"></table></div></div></div><div class="card"><h2>Field Mapping Matrix</h2><div class="tableWrap"><table ` +
  `id="mappingTbl"></table></div></div></section>\n<section id="pageUAT" class="page hidden"><div class="card"><h2>UAT Console</h2><div class="tableWrap"><table id="uatTbl"></table></div></div></section>\n<section id="pageRules" ` +
  `class="page hidden"><div class="card"><h2>Rule Book</h2><div class="tableWrap"><table id="rulesTbl"></table></div></div><div class="card" id="fabAccountMapping"><h2>Fab / Account Group Mapping</h2><div id="agNotice" class="warnbox" role="status" style="white-space:pre-line;margin-bottom:12px" hidden></div><div id="agGrid" class="agGrid"></div>` +
  `<div class="btnrow" style="margin-top:12px"><button class="btn blue agCtl" data-reference-click="23">Save</button><button class="btn light agCtl" data-reference-click="26">Restore Defaults</button></div></div></section>\n</main>`)+`</div>`;
// Keep the reference export actions at the bottom of every tab.
const exportCenter=[...document.querySelectorAll('main > .card')].find(card=>card.querySelector('h2')?.textContent==='Export Center'); document.querySelectorAll('main > .page').forEach(page=>page.append(exportCenter.cloneNode(true))); exportCenter.remove();
const report=createReferenceReport();

function recalculate(){report.setSources(store.raw,store.meta);paintQuarterYears()}
function normalizeNsrRows(rows){return report.canonicalize(rows.map(row=>({...row,'Reference Fcid':row['Reference Fcid']??row['REF FCID']??row['Reference FID']??row.FCID??row.FID??'','Workflow Status':row['Workflow Status']??row['NSR Status']??row['Task Name']??''})),'nsr');}
function sourceDate(value){
  if(value instanceof Date&&!isNaN(value))return nsrData.localDate(value);
  const input=text(value);let iso='';
  if(/^\d{5,7}(?:\.\d+)?$/.test(input)){const date=XLSX.SSF.parse_date_code(Number(input)); if(date)iso=`${date.y}-${String(date.m).padStart(2,'0')}-${String(date.d).padStart(2,'0')}`;}else if(/^\d{4}-\d{2}-\d{2}(?:[ T].*)?$/.test(input))iso=input.slice(0,10);
    else{const match=input.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s.*)?$/); if(match)iso=`${match[3]}-${match[1].padStart(2,'0')}-${match[2].padStart(2,'0')}`;}
  return nsrData.validSapDate(iso)?iso:'';
}
function sourceRange(rows,key,meta){
  if(key==='nsr'&&sourceDate(meta.from)&&sourceDate(meta.to))return {dateFrom:meta.from,dateTo:meta.to,dateField:'NSR query range'};
  const config=SOURCE_DATES[key];if(!config)return {};
  let dateFrom='',dateTo='';
  for(const row of rows){const date=config.fields.map(field=>sourceDate(row[field])).find(Boolean);if(!date)continue; if(!dateFrom||date<dateFrom)dateFrom=date;if(!dateTo||date>dateTo)dateTo=date;}
  return {dateFrom,dateTo,dateField:config.label};
}
function sourceMessage(key,message,kind='ok',transport=false){
  store.messages[key]={message,kind};const target=id('fcSource-'+key),meta=store.meta[key]; const hasRange=kind==='ok'&&meta?.dateFrom&&meta?.dateTo;
  let label=meta?(SOURCE_DATES[key]?(hasRange?`${meta.dateFrom} ~ ${meta.dateTo}`:'Loaded · no dates'):meta.rows.toLocaleString()+(meta.rows===1?' row':' rows')):'Not loaded';
  if(kind==='loading')label=/updat/i.test(message)?'Updating…':'Loading…'; else if(kind==='error'){if(/^Loaded;/.test(message))label='Cache not saved'; else if(/^Import failed:/.test(message))label='Import failed'; else label=meta?'Update failed':'Unavailable';}
  target.textContent=label;target.title=kind==='error'?message:hasRange?'Data range · '+meta.dateField:message; target.className='fc-source-state '+kind+(hasRange?' has-range':'');
  target.setAttribute('aria-label',kind==='error'?message:label);target.tabIndex=kind==='error'?0:-1;
  if(BOBJ_SOURCES[key]){const help=id('fcHelp-'+key);help.hidden=kind!=='error'||!transport;if(help.hidden)help.open=false;}
}
function actionMessage(message){id('fcActionStatus').textContent=message}
function sourceMeta(key){
  const meta=store.meta[key];
  id('fcMeta-'+key).textContent=meta?(meta.kind==='github'?sources[key].label:text(meta.file).split(/[\\/]/).pop()):'';
}
function sourceBusy(key,on){
  on?store.busy.add(key):store.busy.delete(key); document.querySelectorAll(`[data-source-file="${key}"]`).forEach(input=>input.disabled=on); id('fcImport-'+key).disabled=on;
  const actions=key==='fcid'?['bobj-update']:key==='ctl'?['ctl-bobj-update']:key==='nsr'?['nsr-open']:[]; actions.forEach(action=>document.querySelectorAll(`[data-source-action="${action}"]`).forEach(button=>button.disabled=on));
  if(key==='nsr')id('fcNsrLoad').disabled=on;
  id('fcRefresh').disabled=store.refreshing||store.busy.size>0;
}
function startSource(key){store.controllers[key]?.abort();store.runs[key]=(store.runs[key]||0)+1;sourceBusy(key,true);return store.runs[key];}
async function applySource(key,rows,meta){
  store.raw[key]=rows;store.meta[key]={...meta,...sourceRange(rows,key,meta),rows:rows.length,savedAt:meta.savedAt||Date.now()}; sourceMessage(key,'Loaded');sourceMeta(key);recalculate();
}
function sourceCsvText(bytes){
  const data=new Uint8Array(bytes);let content;
  if(data[0]===255&&data[1]===254)content=new TextDecoder('utf-16le').decode(data);
  else if(data[0]===254&&data[1]===255)content=new TextDecoder('utf-16be').decode(data); else{try{content=new TextDecoder('utf-8',{fatal:true}).decode(data)}catch(_){content=new TextDecoder('gb18030').decode(data)}}
  content=content.replace(/^\uFEFF/,''); const header=content.split(/\r?\n/,1)[0],delimiter=header.includes('\t')?'\t':header.split(';').length>header.split(',').length?';':',';
  if(delimiter===',')return content;
  const rows=[];let row=[],value='',quoted=false;
  for(let index=0;index<content.length;index++){
    const char=content[index],next=content[index+1];
    if(char==='"'&&quoted&&next==='"'){value+='"';index++}
    else if(char==='"')quoted=!quoted; else if(char===delimiter&&!quoted){row.push(value);value=''} else if((char==='\r'||char==='\n')&&!quoted){if(row.length||value){row.push(value);rows.push(row);row=[];value=''} if(char==='\r'&&next==='\n')index++;}else value+=char;
  }
  if(quoted)throw new Error('CSV contains an unclosed quoted value.');
  if(row.length||value){row.push(value);rows.push(row)}
  return rows.map(cells=>cells.map(cell=>'"'+cell.replace(/"/g,'""')+'"').join(',')).join('\r\n');
}
async function readSourceFile(file,key){if(!file.name.toLowerCase().endsWith('.csv'))return report.readFile(file,key); const content=sourceCsvText(await file.arrayBuffer()); return report.readFile({name:file.name,text:async()=>content},key);}
function downloadHostedArchive(){
  return new Promise((resolve,reject)=>{
    GM_xmlhttpRequest({
      method:'GET',url:HOSTED_ARCHIVE_URL+'?t='+Date.now(),responseType:'arraybuffer',timeout:30000,nocache:true,anonymous:true,
      onload:response=>{if(response.status>=200&&response.status<300&&response.response)resolve(response.response); else reject(new Error('Data download failed (HTTP '+response.status+').'));}, onerror:()=>reject(new Error('Could not download data.')),
      ontimeout:()=>reject(new Error('Data download timed out.'))
    });
  });
}
async function loadHostedSources(keys=Object.keys(HOSTED_SOURCE_PREFIXES)){
  const runs=Object.fromEntries(keys.map(key=>[key,startSource(key)]));keys.forEach(key=>sourceMessage(key,'Loading data…','loading'));
  let archivePassword='';
  const failed=key=>{if(runs[key]===store.runs[key])sourceMessage(key,(archivePassword?'Data could not be loaded.':'SAP access failed. Sign in to Fiori.')+(store.meta[key]?' Previous data retained.':' Retry or import a file.'),'error');};
  try{
    const response=await unsafeWindow.fetch(HOSTED_PASSWORD_URL,{credentials:'same-origin',cache:'no-store',signal:unsafeWindow.AbortSignal.timeout(30000),headers:{Accept:'application/atom+xml,application/xml,text/xml'}});
    if(!response.ok||(response.headers.get('Content-Type')||'').toLowerCase().includes('text/html'))throw new Error('SAP access failed.');
    const doc=new DOMParser().parseFromString(await response.text(),'application/xml');
    if(doc.getElementsByTagName('parsererror').length)throw new Error('SAP returned invalid XML.');
    const properties=[...doc.getElementsByTagNameNS('*','properties')].find(node=>node.getElementsByTagNameNS('*','Userid')[0]?.textContent.trim()==='LONGKY');
    archivePassword=properties?.getElementsByTagNameNS('*','FullName')[0]?.textContent.replace(/\s+/g,'')||'';
    if(!archivePassword)throw new Error('SAP returned no password.');
    if(keys.every(key=>runs[key]!==store.runs[key]))return;
    let bytes=await downloadHostedArchive();if(keys.every(key=>runs[key]!==store.runs[key]))return;
    if(typeof zip==='undefined')throw new Error('Data reader is unavailable. Import a file manually.');
    // Keep report filenames inside the encrypted outer ZIP.
    const outer=new zip.ZipReader(new zip.Uint8ArrayReader(new Uint8Array(bytes)),{useWebWorkers:false});
    try{
      const entry=(await outer.getEntries()).find(item=>!item.directory&&item.filename.split(/[\\/]/).pop().toLowerCase()==='data.zip');
      if(!entry)throw new Error('The data file is missing.');
      bytes=await entry.getData(new zip.Uint8ArrayWriter(),{password:archivePassword,checkSignature:true});
    }finally{await outer.close()}
    const reader=new zip.ZipReader(new zip.Uint8ArrayReader(new Uint8Array(bytes)),{useWebWorkers:false});
    try{
      const entries=await reader.getEntries();
      for(const key of keys){
        if(runs[key]!==store.runs[key])continue;
        try{
          const entry=entries.find(item=>!item.directory&&item.filename.split(/[\\/]/).pop().toLowerCase().startsWith(HOSTED_SOURCE_PREFIXES[key]));
          if(!entry)throw new Error('The data file is missing.');
          const filename=entry.filename.split(/[\\/]/).pop();
          const data=await entry.getData(new zip.Uint8ArrayWriter(),{checkSignature:true}),parsed=await readSourceFile(new File([data],filename),key);
          if(runs[key]!==store.runs[key])continue;
          const {rows,...recognition}=parsed;await applySource(key,rows,{file:sources[key].label,kind:'github',...recognition});
        }catch(error){failed(key)}
      }
    }finally{await reader.close()}
  }catch(error){keys.forEach(failed)}finally{keys.forEach(key=>{if(runs[key]===store.runs[key])sourceBusy(key,false)})}
}
async function importSource(file,key){
  const run=startSource(key);sourceMessage(key,'Reading '+file.name+'…','loading');
  try{
    let parsed,range={};
    if(file.name.toLowerCase().endsWith('.json')&&key==='nsr'){
      const cache=JSON.parse(await file.text());if(!Array.isArray(cache.rows))throw new Error('This JSON does not contain NSR cache rows.');
      parsed={rows:normalizeNsrRows(cache.rows),sheet:'NSR cache',fields:Object.keys(sources.nsr.fields)};
      if(nsrData.validSapDate(cache.from)&&nsrData.validSapDate(cache.to))range={from:cache.from,to:cache.to};
    }else parsed=await readSourceFile(file,key);
    if(run!==store.runs[key])return;
    const {rows,...recognition}=parsed;await applySource(key,rows,{file:file.name,kind:'file',...recognition,...range});
  }catch(error){if(run===store.runs[key]){const message=error.userMessage?error.message+'\n\n'+error.userMessage:error.message;sourceMessage(key,'Import failed: '+message+(store.meta[key]?' Previous data retained.':''),'error')}}finally{if(run===store.runs[key]){sourceBusy(key,false);
    id('fcFile-'+key).value=''}}
}
async function loadBobj(update=false,key='fcid'){
  const config=BOBJ_SOURCES[key],run=startSource(key),{endpoint,label}=config;sourceMessage(key,update?'Updating '+label+'…':'Reading current '+label+' workbook…','loading');
  try{
    let status=await bridgeRequest('GET',endpoint+'/status',undefined,10000);
    if(!status.hasWorkbook||!status.outputPath)update=true;
    if(update&&!status.running)status=await bridgeRequest('POST',endpoint+'/refresh',{visible:false},20000);
    const deadline=Date.now()+SAP_ACTION_TIMEOUT;
    while(status.running){
      if(Date.now()>deadline)throw new Error('BOBJ update timed out.');
      sourceMessage(key,'Updating BOBJ · '+(status.stage||'running'),'loading'); await new Promise(resolve=>setTimeout(resolve,1000));if(run!==store.runs[key])return;
      status=await bridgeRequest('GET',endpoint+'/status',undefined,10000);
    }
    if(update&&status.stage==='failed')throw new Error(status.message||'BOBJ update failed.');
    if(!status.hasWorkbook||!status.outputPath)throw new Error('No '+label+' workbook. Click Update or Import.');
    const result=await bridgeFileRequest('read',{path:status.outputPath},BRIDGE_FILE_TIMEOUT);
    if(!result.contentBase64)throw new Error('The BOBJ workbook is empty.');
    sourceMessage(key,'Reading BOBJ in the background…','loading'); const bytes=Uint8Array.from(atob(result.contentBase64),character=>character.charCodeAt(0)),parsed=await report.readWorkbook(bytes.buffer,key);
    if(run!==store.runs[key])return;
    const {rows,...recognition}=parsed;await applySource(key,rows,{file:status.outputPath,kind:'bridge',...recognition});
  }catch(error){if(run===store.runs[key])sourceMessage(key,error.message+(store.meta[key]&&!error.transport?' Previous data retained.':''),'error',error.transport)}
  finally{if(run===store.runs[key])sourceBusy(key,false)}
}
function setNsrRange(range){nsrRange=range;id('fcNsrFrom').value=range.from;id('fcNsrTo').value=range.to;id('fcNsrTo').min=range.from;id('fcNsrFrom').max=id('fcNsrTo').max=nsrData.localDate(new Date())}
async function loadSharedNsr(useCache=true){
  const key='nsr',run=startSource(key);sourceMessage(key,useCache?'Checking NSR cache…':'Loading NSR…','loading'); const today=nsrData.localDate(new Date());let range={from:id('fcNsrFrom').value,to:id('fcNsrTo').value};
  try{
    if(useCache){
      const cache=await nsrData.readNsrCache();if(run!==store.runs[key])return;
      if(cache){
        range={from:cache.from,to:cache.to};setNsrRange(range);
        await applySource(key,normalizeNsrRows(cache.rows),{file:'NSR cache',kind:'sap',sheet:'SearchNsrSet',...range,savedAt:cache.savedAt,fields:Object.keys(sources.nsr.fields).filter(name=>name!=='leadTime'&&name!=='workflow')});
        if(cache.to===today)return;
        range={from:cache.from,to:today};
      }else range=nsrData.sapRange();
    }
    if(!nsrData.validSapDate(range.from)||!nsrData.validSapDate(range.to)||range.from>range.to||range.to>today)throw new Error('Choose a valid NSR date range ending on or before today.');
    const controller=new unsafeWindow.AbortController();store.controllers[key]=controller; const result=await nsrData.fetchSapRows(range,controller.signal);if(run!==store.runs[key])return;
    let savedAt=Date.now(),cacheError='';
    try{savedAt=await nsrData.writeNsrCache(result.rows,range,result.pages)}catch(error){cacheError=error.message}
    if(run!==store.runs[key])return;setNsrRange(range); await applySource(key,normalizeNsrRows(result.rows),{file:'SAP live NSR',kind:'sap',sheet:'SearchNsrSet',...range,savedAt,fields:Object.keys(sources.nsr.fields).filter(name=>name!=='leadTime'&&name!=='workflow')});
    if(cacheError)sourceMessage(key,'Loaded; shared cache could not be saved. '+cacheError,'error');
  }catch(error){if(run===store.runs[key])sourceMessage(key,'NSR load failed: '+error.message+(store.meta[key]?' Previous data retained.':''),'error')}
  finally{if(run===store.runs[key]){sourceBusy(key,false);delete store.controllers[key]}}
}
async function restoreSources(reset=false){
  if(store.refreshing||store.busy.size)return;
  store.refreshing=true;const refresh=id('fcRefresh');refresh.disabled=true;refresh.textContent='Refreshing…';
  try{
    const keys=Object.keys(sources);
    if(reset){
      keys.forEach(key=>{store.runs[key]=(store.runs[key]||0)+1;store.raw[key]=[];delete store.meta[key]; sourceMessage(key,'Not loaded','');sourceMeta(key);sourceBusy(key,true);id('fcFile-'+key).value='';}); setNsrRange(nsrData.sapRange());store.pages={};actionMessage('');
        recalculate();
    }
    // Manual imports are page-local; discard source caches left by earlier versions.
    try{await nsrData.deleteDatasets(keys.map(key=>CACHE_PREFIX+key))} catch(error){if(reset)actionMessage('Local source cache could not be reset. '+error.message)} finally{if(reset)keys.forEach(key=>sourceBusy(key,false))}
    recalculate();
    if(store.meta.nsr?.from)setNsrRange({from:store.meta.nsr.from,to:store.meta.nsr.to});
    const loads=[];
    for(const key of Object.keys(BOBJ_SOURCES))if(!store.busy.has(key))loads.push(loadBobj(false,key));
    const hostedKeys=Object.keys(HOSTED_SOURCE_PREFIXES).filter(key=>!store.busy.has(key));if(hostedKeys.length)loads.push(loadHostedSources(hostedKeys));
    if(!store.busy.has('nsr')&&(!store.meta.nsr||store.meta.nsr.kind==='sap'))loads.push(loadSharedNsr());
    await Promise.all(loads);
  }finally{store.refreshing=false;refresh.disabled=store.busy.size>0;refresh.textContent='Refresh'}
}


function quarterYear(value){const match=text(value).match(/^Q(?:Mar|Jun|Sep|Dec)'(\d{2}|\d{4})$/i);return match?(match[1].length===2?'20'+match[1]:match[1]):''}
function toggleQuarterYear(year){const options=[...id('fQuarter').options].filter(option=>quarterYear(option.value)===year),selectAll=options.some(option=>!option.selected); options.forEach(option=>option.selected=selectAll);report.syncChosen('fQuarter');report.applyFilters();
  paintQuarterYears();}
function paintQuarterYears(){
  document.querySelectorAll('[data-quarter-year]').forEach(button=>{
    const options=[...id('fQuarter').options].filter(option=>quarterYear(option.value)===button.dataset.quarterYear),count=options.filter(option=>option.selected).length,all=options.length>0&&count===options.length; button.disabled=!options.length;button.classList.toggle('active',
      all);button.classList.toggle('partial',count>0&&!all);button.setAttribute('aria-pressed',String(all));
  });
}
document.addEventListener('click',event=>{
  document.querySelectorAll('.fc-source-help[open]').forEach(help=>{if(!help.contains(event.target))help.open=false});
  const label=event.target.closest('.fc-filter>label[for]');if(label&&label.htmlFor!=='fSearch'){event.preventDefault();jQuery(id(label.htmlFor)).trigger('chosen:open');return}
  const button=event.target.closest('button');if(!button)return;
  if(button.dataset.importSource){id('fcFile-'+button.dataset.importSource).click();return}
  if(button.dataset.quarterYear){toggleQuarterYear(button.dataset.quarterYear);return}
  if(button.dataset.clearFilter){report.clearSelection(button.dataset.clearFilter);paintQuarterYears();return}
  if(button.dataset.selectAll){report.selectAll(button.dataset.selectAll);paintQuarterYears();return}
  if(button.dataset.sourceAction){const action=button.dataset.sourceAction; if(action==='bobj-update')loadBobj(true); else if(action==='ctl-bobj-update')loadBobj(true,'ctl');
    else if(action==='nsr-open'){setNsrRange(nsrRange);id('fcNsrDialog').showModal()}return;}
  if(button.id==='fcRefresh')restoreSources(true); else if(button.id==='fcReset'){report.resetFilters();paintQuarterYears()}
});
document.addEventListener('keydown',event=>{if(event.key==='Escape')document.querySelectorAll('.fc-source-help[open]').forEach(help=>help.open=false)});
id('fcNsrFrom').addEventListener('input',()=>{id('fcNsrTo').min=id('fcNsrFrom').value});
id('fcNsrCancel').addEventListener('click',()=>id('fcNsrDialog').close());
id('fcNsrForm').addEventListener('submit',event=>{event.preventDefault();if(store.busy.has('nsr'))return;id('fcNsrDialog').close();loadSharedNsr(false)});
document.addEventListener('change',event=>{const input=event.target;if(input.dataset.sourceFile&&input.files?.[0])importSource(input.files[0],input.dataset.sourceFile)}); id('fSearch').addEventListener('input',()=>report.applyFilters());
filterDefs.forEach(([name])=>{jQuery(id(name)).on('change',()=>{report.applyFilters();paintQuarterYears()}); if(jQuery.fn.chosen){jQuery(id(name)).chosen({width:'100%',search_contains:true,hide_results_on_select:false,placeholder_text_multiple:'All',display_selected_options:false});
  report.registerChosen(name)}});
// Source initialization is deliberately separate from the original report engine.
setNsrRange(nsrRange);recalculate();restoreSources();

function createReferenceReport(){
  // The supplied HTML governs every metric, rule, term, view and export below.
  // Deviating from that reference requires the user's confirmation.
  const window={XLSX,jQuery,alert:message=>unsafeWindow.alert(message),confirm:message=>unsafeWindow.confirm(message)}; const alert=window.alert;
  const {loadStatus,fQuarter,fAccount,fCustomer,fRisk,fSearch,filterContextCard,filterContextBadge,filterContextMini,pageExecutive,kPop,kPre,kPost,kRisk,kSlotReady,kNSRWatch,kNSRImpact,funnel,heatmap,driverBars,watchTbl,pageStability,projectionMini,stabilityTbl,
    pageSlot,kSlotMcsd,kSlotNoMcsd,kNoSlotCtl,slotTbl,pageDrivers,driverCards,riskDriverDetail,riskTbl,pageNSR,openNsrPopulation,nsrTbl,pageQuarter,quarterTbl,quarterSourceTbl,pageFCID,fcidInput,fcidBasic,fcidAction,fcidTbl,pageMapping,loadTbl,overlapTbl,mappingTbl,pageUAT,uatTbl,
    pageRules,rulesTbl,agNotice,agGrid}=Object.fromEntries([...document.querySelectorAll('[id]')].map(element=>[element.id,element]));
// BEGIN FCID REPORT ENGINE
const cfg={fcid:{label:'FCID Database',sheet:'Sheet1',key:'FID',commit:'Commit: Field',slotFlag:'Flag: Slotted',quarter:'Time: Calendar Quarter',account:'Account Group',fab:'Fab Name',customer:'Customer',product:'Primary Product',system:'System Description',
  quote:'Active Quote Number',targetShip:'Date: Target Ship',ship:'Date: Mfg Commit Ship',rd:'Date: Sales Ops Request Date (RD)'},ctl:{label:'CTL Report',sheet:'CTL BOBJ',key:'Opportunity ID',ctlId:'CTL ID',change:'Change Cat',status:'Status',mcsd:'MCSD @ Current',srd:'SRD @ Current',
  srdSops:'SRD @ SOPs Approval',sops:'1st SOps Appr Dt',summary:'Change Summary',mfg:'Mfg Upgrade Type'},ctlNsr:{label:'CTL-NSR List',sheet:'export',key:'FCID',ctlId:'CTL ID',change:'Change Type',status:'CTL Status',mcsd:'MCSD',srd:'Sales Ops RD',fieldCrd:'Field CRD',
  sops:'Sales Ops Approval Date',quote:'Quote Number',sourceQuote:'Source Quote',summary:'Change Summary',stage:'Stage',mfg:'Mfg Upgrade Type'},nsr:{label:'NSR Log', sheet:'Whole list', key:'Reference Fcid', nsr:'NSR#', status:'NSR Status', workflow:'Workflow Status',
  customer:'Customer', product:'Primary Product', title:'NSR Title', lt:'NSR LT'},nce:{label:'NCE / FA Review', sheet:'Export', key:'FCID', nce:'NCe #', state:'NCe State', impact:'Impact', title:'Title', customer:'Customer', openDate:'Open Date', damageGroup:'Damage Group',
  damageCode:'Damage Code', partNumber:'Part_Number', partDescription:'Part_Description'}};
const raw={fcid:[],ctl:[],ctlNsr:[],nsr:[],nce:[]};let meta={},M={rows:[],excluded:[],risks:[],idx:{},originalCommit:0,fcidAudit:{rawCommitRows:0,validCommitRows:0,invalidFidRows:0,distinctCommitFcids:0,duplicateFcidGroups:0,duplicateRowsRemoved:0,fabSegmentPreferred:0}};
const AG0={"AT1":(`Analog Foundries (CUMEC)\nASMC - Fab 3\nBYD Semiconductor Chengdu*\nCansemi*\nChina Resources Microelectronics\nChongqing SIMIC Semiconductor Limited\nCR Micro*\nCR Runan Technologies (Chongqing)\nCSMC - Fab *\nGalaxycore *\nGTA Semico` +
  `nductor\nGuangzhou Cansemi*\nGuangzhou Zen Semiconductor\nHangzhou Fullsemi Technology\nHangzhou Hesheng MicroElectronics\nHangzhou Power Innovation Microelectron\nHH - Fab *\nHikstor Technology\nHuBei Sanan Optoelectronics co.ltd\nHun` +
  `an Jiechuwei Semiconductor Technology\nHunan Sanan Semiconductor\nInnostar Semiconductor - *\nIRay Technology Hefei\nJiangsu Advanced Memory (AMS)\nJiangSu Nexic Technology\nJie Jie Microelectronics (Nan Tong) Tech\nJinan BYD Semicondu` +
  `ctor\nMaxscend Microelectronics Company\nResources West Microelectronics (RWME)\nRunpeng Semiconductor (Shenzhen)\nSanan-ST Microelectronics\nShanghai Dingtai Jiangxin Tech.\nShanghai GTX Semiconductor*\nShanghai Huayaoxin Semi (Foreca` +
  `st Only)\nShanghai Integrated Circuit *\nShanghai Optoelectronic Science\nSkysilicon Co. (China Resources Micro>)\nStarPower Technologies\nSunny OmniLight*\nSwaysure Technology*\nTI - Chengdu\nUMC - HJTC Fab 1\nUnion Semiconductor (Hef` +
  `ei) Co.,Ltd\nWUYUAN\nXi'an Longwei Semiconductor\nXiamen ITRI/USC\nYOFC Advanced Semiconductor\nZhejiang ChipRight Semiconductor\nZheJiang HiChip Micro Tech\nZhejiang ICsprout Semiconductor\nZhejiang Xinsheng Semiconductor\nZhuzhou CRR` +
  `C Times Electric\nZing Semiconductor Corporation`),
    "AT2":(`BEMC\nGoerPixel\nGuangdong Ziyue Semiconductor (Forecast)\nJi Guang Semiconductor (Shaoxing) Corp.\nMRDI\nNingbo Semiconductor\nPST*\nPXW Semiconductor\nQuanxin Integrated Circuit Mfg-QXIC\nRong Semiconductor*\nSEC Chip*\nSemicon Manuf` +
  ` Oriental*\nShenzhen Pengxinxu Technology Co., Ltd\nShenzhen Pensun Technology Co., Ltd\nShenzhen Qili Tech\nSiEn (Qingdao)*\nSilex Microsystems Beijing Co\nSMBC B3\nSMEC Shaoxing*\nSMIC*\nUNT*\nYDMT\nYongxin*\nZhuhai Hengqin Jointed F` +
  `oundry Semicon`),"AT3":(`ANHUI JUHE MICROELECTRONICS*\nASP Tech (Chengdu)\nBeijing Jidian Holding*\nChangxin*\nChengdu High-tech Jin Science & Tech\nCHJS (Chengdu High-tech Jin Science)\nChongqing lnnoevsic Electronics\nCXMT Fab2*\nDongguan Guangmao Technology` +
  `*\nGuangdong 3D-Semi\nHangzhou HFC Semiconductor\nHefei Xinjing\nHFC Semiconductor Corporation\nInnoscience (Suzhou) Semiconductor\nInnotron Memory Co.*\nJiangsu Silicon Integrity (Forecast Only\nNexchip Semiconductor\nNexmemory Techno` +
  `logy Co., Ltd.\nUSC - Xiamen 12X\nWuhan Chuxing Technology\nWuhan HSMC\nWuhan Xinxin Semi Mfg Co. Ltd*\nXinlian Microelectronics\nYangtze Advanced Memory*\nYMTC*\nZhe Jiang Super Dimension Semiconductor`),"Hynix":"SK hynix - *","Intel":"Intel - China F68*",
    "OSAT":(`Amkor - Assembly & Test (Shanghai)\nAnhui Semi Integr Display Technology\nATX*\nBeijing H.F.I. Elect*\nBeijing ICLink Technology (ForecastOnly)\nBeijing Xinli Technology Innovation*\nBOE Technology*\nChengdu ECHINT Technology Co.,Ltd\n` +
  `DSHX (Shanghai) Corp.\nForehope*\nGreat Bay Chiplet*\nHefei Payton Storage (Forecast Only)\nHsingchen Technology*\nHuaian Imaging (HiDM)\nHuatian Tech*\nHubei Yangtze*\nHuizhou Biwin Storage Technology(FCOnly)\nInnosemi*\nJCET*\nJiangs` +
  `u Silicon Integrity\nJiangyin Changdian Advanced Pkg (JCAP)\nKore Semiconductor 200MM\nKunming BOE Display\nLuxshare*\nNatural-Integration Advanced\nNCAP*\nNingbo TRS Microelectronics\nQuliang Electronics*\nShan Dong Qi Xin Mems Corp\n` +
  `Shanghai Echip Semiconductor\nShanghai Xianfang Semiconductor\nShanghai Xinyuan Innovation Center\nShenzhen Qili Tech F-PJT (Forecast Only)\nSJ Semi - Jiangyin*\nSky Chip Inter Tech (Forecast Only)\nSPIL Suzhou\nSuzhou TF-AMD Semicondu` +
  `ctor\nTF AMD Microelectronics (Penang)\nTianfu Xinglong Lake Laboratory\n*Tongfu*\nUNISEM CHENGDU (Forecast Only)\nWuhu Micro Display Intelligent Technolog\nXiamen Sky Semiconductor Technology\nYibu Semiconductor Co., LTD\nYongjiang La` +
  `b*\nYunnan Invensight Optoelectronics\nZhejiang Hexin Integrated Circuit (FC)\nZhejiang ICTruth Semiconductor\nZiyi Microelectronics Ltd.`),"Samsung":"Samsung - *","Training Center":"Lam - Shanghai Training Center","TSMC":"TSMC - *"};let AG={...AG0},AGI=new Map,AGP=[];
const demo={fcid:[{'FID':'257020','Commit: Field':'Commit','Flag: Slotted':'Y','Time: Calendar Quarter':"QMar'26",'Account Group':'AT3',Customer:'Dongguan Guangmao','Primary Product':'CoronusHP','System Description':'System','Date: Target Ship':'11/20/2026',
  'Date: Sales Ops Request Date (RD)':'11/20/2026'},{'FID':'233525','Commit: Field':'Commit','Flag: Slotted':'#','Time: Calendar Quarter':"QSep'26",'Account Group':'AT3',Customer:'Nexchip','Primary Product':'E4/Kiyox3','System Description':'System','Date: Target Ship':'',
  'Date: Sales Ops Request Date (RD)':'11/05/2026'},{'FID':'250544','Commit: Field':'Commit','Flag: Slotted':'Y','Time: Calendar Quarter':"QSep'26",'Account Group':'AT1',Customer:'SK hynix','Primary Product':'EXT_AHM','System Description':'System','Date: Target Ship':'08/28/2026',
  'Date: Sales Ops Request Date (RD)':'08/28/2026','Date: Mfg Commit Ship':'06/12/2026'},{'FID':'999001','Commit: Field':'Commit','Flag: Slotted':'Y','Time: Calendar Quarter':"QSep'26",'Account Group':'AT1',Customer:'Demo','Primary Product':'SW',
  'System Description':'KIYOFXE SW-UPGRADE','Date: Target Ship':'09/01/2026'}],ctl:[{'Opportunity ID':'257020','CTL ID':'8000295092','Change Cat':'Slot','Status':'Completed','SRD @ Current':'11/20/2026','SRD @ SOPs Approval':'11/20/2026','1st SOps Appr Dt':'06/01/2026',
  'Change Summary':'Slot CTL'},{'Opportunity ID':'250544','CTL ID':'8000222454','Change Cat':'Slot','Status':'Completed','MCSD @ Current':'06/12/2026','SRD @ SOPs Approval':'08/28/2026','1st SOps Appr Dt':'01/10/2026','Change Summary':'Slot CTL'}],ctlNsr:[{FCID:'257020',
  'CTL ID':'8000295092','Change Type':'Slot','CTL Status':'Completed','Sales Ops RD':'11/20/2026','Sales Ops Approval Date':'06/01/2026'},{FCID:'250544','CTL ID':'8000222454','Change Type':'Slot','CTL Status':'Completed','MCSD':'06/12/2026','Sales Ops RD':'08/28/2026',
  'Sales Ops Approval Date':'01/10/2026'}],nsr:[{'Reference Fcid':'233525','NSR#':'N304262','NSR Status':'PENDING WITH DESIGN ENGINEERING','Workflow Status':'Pending','NSR Title':'Nexchip Etch sys'}],nce:[]};
function clean(v){let s=String(v??'').trim();s=s.replace(/\.0$/,'');return (s===''||s==='#'||s.toLowerCase()==='nan'||s==='Not assigned')?'':s}function norm(s){return String(s??'').toLowerCase().replace(/[^a-z0-9]/g,'')}function get(row,names){for(const n of names){if(row&&
  Object.prototype.hasOwnProperty.call(row,n))return clean(row[n])}let keys=Object.keys(row||{});for(const n of names){let nn=norm(n);let k=keys.find(k=>norm(k)===nn)||keys.find(k=>norm(k).includes(nn)||nn.includes(norm(k)));
  if(k)return clean(row[k])}return ''}
function toDate(v){
  const s=clean(v);if(!s)return null;let date;
  // Exports contain calendar dates: preserve the same local date for Excel and CSV.
  if(/^\d{5}(?:\.\d+)?$/.test(s)){
    const serialDate=new Date((Math.floor(Number(s))-25569)*86400000);
    date=new Date(serialDate.getUTCFullYear(),serialDate.getUTCMonth(),serialDate.getUTCDate());
  }else{
    const iso=s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/);
    date=iso?new Date(Number(iso[1]),Number(iso[2])-1,Number(iso[3])):new Date(s);
  }
  if(isNaN(date))return null;date.setHours(0,0,0,0);return date;
}
function weeks(a,b){
  if(!a||!b)return null;
  // Count calendar days in UTC so daylight-saving changes cannot add or remove hours.
  return (Date.UTC(b.getFullYear(),b.getMonth(),b.getDate())-Date.UTC(a.getFullYear(),a.getMonth(),a.getDate()))/604800000;
}
function qFromDate(d){if(!d)return '';let y=String(d.getFullYear()).slice(-2),m=d.getMonth()+1;return (m<=3?'QMar':m<=6?'QJun':m<=9?'QSep':'QDec')+"'"+y}function upper(v){return clean(v).toUpperCase()}function pct(a,b){return b?(a/b*100).toFixed(1)+
  '%':'0.0%'}function uniq(a){return [...new Set(a.filter(x=>x!==''&&x!==undefined&&x!==null))]}
const AGK='fcid.ag.v1',AGO=['AT1','AT2','AT3','OSAT','Hynix','Intel','Samsung','Training Center','TSMC'];
function agKey(v){return upper(v).replace(/\s+/g,' ')}
function agRx(s){return new RegExp('^'+s.split('*').map(x=>x.replace(/[.*+?^$(){}|[\]\\]/g,c=>'\\'+c)).join('.*')+'$','i')}
function normAg(x){
    let o={},seen=new Map,dups=[];
    Object.entries(x||{}).forEach(([g,s])=>{g=clean(g);if(!g)return; let keep=[]; uniq(String(s??'').split(/\r?\n/).map(clean)).forEach(f=>{let k=agKey(f),old=seen.get(k);if(!k)return; if(old&&old!==g)dups.push(f+' ('+old+' / '+g+')'); else if(!old){seen.set(k,g);keep.push(f)}});
      o[g]=keep.join('\n');});
    return {o,dups}
}
function setAg(x,warn=1){
    let r=normAg(x);
    if(r.dups.length){if(warn)alert('Duplicate Fab mappings found:\n\n'+r.dups.slice(0,12).join('\n')+'\n\nEach Fab must belong to only one Account Group.'); return false;}
    AG=r.o;AGI=new Map;AGP=[]; Object.entries(AG).forEach(([g,s])=>s.split('\n').forEach(f=>{let k=agKey(f);if(!k)return; if(f.includes('*'))AGP.push([agRx(f),g,f.replace(/\*/g,'').length]); else AGI.set(k,g);})); AGP.sort((a,b)=>b[2]-a[2]); return true
}
function accountOf(r){let a=get(r,[cfg.fcid.account]); if(a&&!/^(NOT ASSIGNED|#?N\/A|NA|NULL|UNKNOWN|0|00:00:00|#(?:VALUE!|REF!|NAME\?|NUM!|DIV\/0!|NULL!))$/i.test(a))return a; let f=get(r,[cfg.fcid.fab]),m=AGI.get(agKey(f)),p; if(!m&&(p=AGP.find(x=>x[0].test(f))))m=p[1]; return m||
  '[Undefined]';}
function readAg(){let o={};document.querySelectorAll('#agGrid textarea').forEach(t=>o[t.dataset.g]=t.value);return o}
function renderAg(){agGrid.innerHTML=Object.entries(AG).sort((a,b)=>(AGO.indexOf(a[0])+1||99)-(AGO.indexOf(b[0])+1||99)).map(([g,s])=>`<div class="agBox"><label>`+escHtml(g)+`</label><textarea data-g="`+escHtml(g)+'" spellcheck="false">'+escHtml(s)+`</textarea></div>`).join('')}
function agEntries(mapping){
    if(!mapping||typeof mapping!=='object'||Array.isArray(mapping)||Object.values(mapping).some(value=>typeof value!=='string'))throw Error('Invalid Fab mapping');
    const normalized=normAg(mapping);if(normalized.dups.length)throw Error('Duplicate Fab mappings');
    return new Map(Object.entries(normalized.o).flatMap(([group,value])=>value.split('\n').filter(Boolean).map(fab=>[agKey(fab),{fab,group}])));
}
function mergeAg(base,saved){
    const before=agEntries(base),user=agEntries(saved),current=agEntries(AG0),merged=Object.fromEntries(Object.keys(AG0).map(group=>[group,[]])),conflicts=[];
    // Compare ownership per Fab, so reordering lines or changing case is not an edit.
    new Set([...current.keys(),...user.keys(),...before.keys()]).forEach(key=>{
        const original=before.get(key),custom=user.get(key),latest=current.get(key),edited=custom?.group!==original?.group;
        if(edited&&latest?.group!==original?.group&&custom?.group!==latest?.group)
            conflicts.push((custom||latest||original).fab+': saved '+(custom?.group||'[Unmapped]')+'; updated default '+(latest?.group||'[Unmapped]'));
        const chosen=edited?custom:latest;if(chosen)(merged[chosen.group]??=[]).push(chosen.fab);
    });
    return {mapping:Object.fromEntries(Object.entries(merged).map(([group,fabs])=>[group,fabs.join('\n')])),conflicts};
}
function storeAg(base=AG0,mapping=AG){
    try{GM_setValue(AGK,JSON.stringify({base,mapping}))}catch(e){console.warn('Tampermonkey settings unavailable',e);return false}
    try{const storage=persistentStore();storage?.removeItem(AGK);storage?.removeItem('toolkitKyra.fcid.accounts.v1')}catch(e){/* saved in Tampermonkey; old website data can remain */}
    return true
}
function loadAg(){
    try{
        const saved=GM_getValue(AGK,null),candidates=saved===null?[readPersistent(AGK,null),readPersistent('toolkitKyra.fcid.accounts.v1',null)]:[saved];
        for(const value of candidates){
            try{
                const data=typeof value==='string'?JSON.parse(value):value,base=data?.base||null,user=data?.mapping||data;
                const result=base?mergeAg(base,user):{mapping:user,conflicts:[]};if(!base)agEntries(user);
                if(!setAg(result.mapping,0))continue;
                const notice=base?(result.conflicts.length?'Updated defaults conflict with your saved mapping. Saved assignments are kept for now. Review the mappings below, then click Save to confirm, or Restore Defaults.\n\n'+result.conflicts.join('\n'):''):
                    'This saved mapping has no update baseline. Review it and click Save to keep it, or Restore Defaults to use the latest defaults.';
                if(notice){agNotice.textContent=notice;agNotice.hidden=false;requestAnimationFrame(()=>jumpToSection('Rules','fabAccountMapping'))}
                // Keep the original baseline until conflicts are explicitly confirmed.
                if(saved===null)storeAg(base,user);else if(base&&!notice&&JSON.stringify(base)!==JSON.stringify(AG0))storeAg();
                return;
            }catch(e){/* try the next legacy setting */}
        }
    }catch(e){console.warn('Tampermonkey settings unavailable',e)}
    setAg(AG0,0)
}
function restoreAg(){setAg(AG0,0);renderAg();recalculate();if(storeAg())agNotice.hidden=true;else alert('Default settings could not be saved in Tampermonkey. Please try again.')}
function saveAg(btn){
    if(!setAg(readAg()))return false;
    recalculate();
    if(!storeAg()){alert('Settings could not be saved in Tampermonkey. Please try again.');return false}
    agNotice.hidden=true;
    if(btn){clearTimeout(btn._st);btn.textContent='Saved';btn._st=setTimeout(()=>btn.textContent='Save',1200)}
    return true
}
function csvParse(str){let rows=[],row=[],v='',q=false;for(let i=0;i<str.length;i++){let c=str[i],n=str[i+1];if(c==='"'&&q&&n==='"'){v+='"';i++}else if(c==='"'){q=!q}else if(c===','&&!q){row.push(v);v=''}else if((c==='\n'||c==='\r')&&!q){if(row.length||v){row.push(v);rows.push(row);
  row=[];v=''}if(c==='\r'&&n==='\n')i++}else v+=c}if(row.length||v){row.push(v);rows.push(row)}let h=(rows.shift()||[]).map(x=>clean(x));return rows.filter(r=>r.some(x=>clean(x))).map(r=>Object.fromEntries(h.map((k,i)=>[k,clean(r[i])])))}
const SP={
fcid:{m:65,k:cfg.fcid.key,f:[[cfg.fcid.key,40,1,'FCID','Field ID'],[cfg.fcid.commit,20,1,'Commit Field','Field Commit','Commit Status'],[cfg.fcid.quarter,10,0,'Calendar Quarter','Cal Quarter','Quarter'],[cfg.fcid.customer,5,0,'Customer Name','End Customer'],[cfg.fcid.system,5,0,
  'System Desc','System'],[cfg.fcid.targetShip,5,0,'Target Ship Date','Target Ship'],[cfg.fcid.rd,5,0,'Sales Ops Request Date','Sales Ops RD','RD'],[cfg.fcid.account,5,0,'Account Team','Sales Territory','AT'],[cfg.fcid.product,5,0,'Product','Product Name','FID Description'],
  [cfg.fcid.fab,0,0,'Fab'],[cfg.fcid.slotFlag,0,0,'Flag Slotted','Slotted Flag','Slot Flag'],[cfg.fcid.ship,0,0,'Mfg Commit Ship','Manufacturing Commit Ship Date','MCSD'],[cfg.fcid.quote,0,0,'Quote Number','Active Quote']]},
ctl:{m:65,k:cfg.ctl.key,f:[[cfg.ctl.key,35,1],[cfg.ctl.ctlId,25,1,'CTL#'],[cfg.ctl.change,10,0,'Change Category'],[cfg.ctl.status,10,0,'CTL Status'],[cfg.ctl.mcsd,5,0,'Current MCSD','MCSD'],[cfg.ctl.srd,5,0,'Current SRD','Sales Ops RD','SRD'],[cfg.ctl.sops,5,0,
  'Sales Ops Approval Date','SOPS Approval Date'],[cfg.ctl.summary,5,0,'Summary'],[cfg.ctl.srdSops,0,0,'SRD at SOPs Approval'],[cfg.ctl.mfg,0,0,'Manufacturing Upgrade Type']]},
ctlNsr:{m:65,k:cfg.ctlNsr.key,f:[[cfg.ctlNsr.key,35,1,'FID'],[cfg.ctlNsr.ctlId,25,1,'CTL#'],[cfg.ctlNsr.change,10,0,'CTL Change Type'],[cfg.ctlNsr.status,10,0,'Status'],[cfg.ctlNsr.mcsd,5,0,'MCSD @ Current'],[cfg.ctlNsr.srd,5,0,'SRD','SRD @ Current'],[cfg.ctlNsr.sops,5,0,
  '1st SOps Appr Dt'],[cfg.ctlNsr.stage,5,0,'CTL Stage'],[cfg.ctlNsr.fieldCrd,0,0,'CRD'],[cfg.ctlNsr.quote,0,0,'Quote ID'],[cfg.ctlNsr.sourceQuote,0,0,'Source Quote Number'],[cfg.ctlNsr.summary,0,0,'Summary'],[cfg.ctlNsr.mfg,0,0,'Manufacturing Upgrade Type']]},
nsr:{m:65,k:cfg.nsr.key,f:[[cfg.nsr.key,35,1,'Reference FID','FCID'],[cfg.nsr.nsr,30,1,'NSR Number'],[cfg.nsr.status,10,0,'Status'],[cfg.nsr.workflow,10,0,'Workflow'],[cfg.nsr.title,10,0,'Task Name','Title'],[cfg.nsr.lt,5,0,'Lead Time','LT'],[cfg.nsr.customer,0,0,'Customer Name'],
  [cfg.nsr.product,0,0,'Product']]},
nce:{m:65,k:cfg.nce.key,f:[[cfg.nce.key,35,1,'FID'],[cfg.nce.nce,30,1,'NCE Number'],[cfg.nce.state,10,0,'State','Status'],[cfg.nce.impact,10,0,'NCE Impact'],[cfg.nce.title,5,0,'NCE Title'],[cfg.nce.openDate,5,0,'Opened Date','Creation Date'],[cfg.nce.damageGroup,3],
  [cfg.nce.damageCode,2],[cfg.nce.partNumber,0],[cfg.nce.partDescription,0],[cfg.nce.customer,0,0,'Customer Name']]}}
function aliasesFor(f){return uniq([f[0],...f.slice(3)])}
function exactAliasKey(row,aliases){let keys=Object.keys(row||{}),wanted=new Set(aliases.map(norm).filter(Boolean));return keys.find(key=>wanted.has(norm(key)))||''}
function scoreSourceHeader(values,k){let p=SP[k],headers=new Set((values||[]).map(norm).filter(Boolean)),score=0,missing=[];p.f.forEach(f=>{let hit=aliasesFor(f).some(a=>headers.has(norm(a)));if(hit)score+=f[1];else if(f[2])missing.push(f[0])});return {score,missing,
  requiredOk:!missing.length}}
function betterHeader(a,b){if(!b)return true;return Number(a.requiredOk)>Number(b.requiredOk)||(a.requiredOk===b.requiredOk&&a.score>b.score)}
function bestHeaderFromMatrix(matrix,k,startRow=0){let best=null;(matrix||[]).slice(0,50).forEach((values,index)=>{let scored=scoreSourceHeader(values,k),candidate={...scored,headerRow:startRow+index};if(betterHeader(candidate,best))best=candidate});return best||{score:0,
  missing:SP[k].f.filter(x=>x[2]).map(x=>x[0]),requiredOk:false,headerRow:startRow}}
function sourceScoreError(k,candidate){let label=cfg[k].label,sheet=candidate?.sheet||'Not identified',score=candidate?.score||0,missing=(candidate?.missing||[]).join(', ')||'None';let err=new Error(label+' recognition score is too low: '+score+'/100');err.code='LOW_SOURCE_SCORE';
  err.userMessage='File type validation failed!\n\nUpload slot: '+label+'\nBest candidate sheet: '+sheet+'\nHighest score: '+score+'/100\nMissing required fields: '+missing+
  '\n\nThe wrong file may have been selected, or the header may not be within the first 50 rows. Please check the file and try again.';return err}
function validateSourceCandidate(k,candidate){if(!candidate||!candidate.requiredOk||candidate.score<SP[k].m)throw sourceScoreError(k,candidate);return candidate}
function detectWorkbookSource(wb,k){let candidates=[];wb.SheetNames.forEach((sheetName,sheetOrder)=>{let ws=wb.Sheets[sheetName];if(!ws||!ws['!ref'])return;let fullRange=XLSX.utils.decode_range(ws['!ref']),previewRange={s:{r:fullRange.s.r,c:fullRange.s.c},e:{r:Math.min(fullRange.e.r,
  fullRange.s.r+49),c:fullRange.e.c}},matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',blankrows:true,raw:false,range:previewRange}),best=bestHeaderFromMatrix(matrix,k,fullRange.s.r);candidates.push({...best,sheet:sheetName,dataRows:Math.max(0,fullRange.e.r-best.headerRow),
  configured:sheetName===cfg[k].sheet,sheetOrder})});candidates.sort((a,b)=>Number(b.requiredOk)-Number(a.requiredOk)||b.score-a.score||b.dataRows-a.dataRows||Number(b.configured)-Number(a.configured)||a.sheetOrder-b.sheetOrder);return validateSourceCandidate(k,candidates[0])}
function csvToMatrix(str){let rows=[],row=[],value='',quoted=false;for(let i=0;i<str.length;i++){let char=str[i],next=str[i+1];if(char==='"'&&quoted&&next==='"'){value+='"';i++}else if(char==='"'){quoted=!quoted}else if(char===','&&!quoted){row.push(value);value=''}else if((char===
  '\n'||char==='\r')&&!quoted){if(row.length||value){row.push(value);rows.push(row);row=[];value=''}if(char==='\r'&&next==='\n')i++}else value+=char}if(row.length||value){row.push(value);rows.push(row)}return rows}
function rowsFromMatrix(matrix,headerRow){let headers=(matrix[headerRow]||[]).map(clean);return matrix.slice(headerRow+1).filter(row=>row.some(value=>clean(value))).map(row=>Object.fromEntries(headers.map((header,index)=>[header,clean(row[index])])))}
function canonicalizeSourceRows(rows,k){let p=SP[k];return rows.map(source=>{let row={...source};p.f.forEach(f=>{if(clean(row[f[0]]))return;let key=exactAliasKey(source,aliasesFor(f));if(key)row[f[0]]=source[key]});return row}).filter(row=>clean(row[p.k]))}
async function loadFile(e,k){let f=e.target.files[0],pm=meta[k],dm=dataMode,done=false;if(!f)return;try{let ext=f.name.split('.').pop().toLowerCase(),rows,candidate;if(ext==='csv'){let matrix=csvToMatrix(await f.text());candidate=validateSourceCandidate(k,
  {...bestHeaderFromMatrix(matrix,k,0),sheet:'CSV',dataRows:matrix.length});rows=rowsFromMatrix(matrix,candidate.headerRow)}else{if(!window.XLSX)throw new Error('XLSX parser not loaded');let wb=XLSX.read(await f.arrayBuffer(),{type:'array'});candidate=detectWorkbookSource(wb,k);rows=
  XLSX.utils.sheet_to_json(wb.Sheets[candidate.sheet],{defval:'',range:candidate.headerRow})}let normalized=canonicalizeSourceRows(rows,k);if(!normalized.length)throw sourceScoreError(k,{...candidate,score:0,missing:[SP[k].k]});raw[k]=normalized;meta[k]={file:f.name,
  sheet:candidate.sheet,headerRow:candidate.headerRow+1,rows:raw[k].length,score:candidate.score,mode:'header score '+candidate.score+'/100 | row '+(candidate.headerRow+1)};if(dm==='demo'&&meta[k]!==pm){let rr=raw[k],mm=meta[k];Object.keys(raw).forEach(x=>raw[x]=[]);raw[k]=rr;meta={};
  meta[k]=mm;dataMode='live'}setStatus();done=true;recalculate()}catch(err){if(!done)e.target.value='';if(err.code==='LOW_SOURCE_SCORE')window.alert(err.userMessage);loadStatus.innerHTML=`<b>`+(done?'Calculation failed after loading: ':'Load failed: ')+escHtml(f.name)+`</b><br>`+
  escHtml(err.message)}}
function setStatus(){loadStatus.innerHTML=Object.keys(cfg).map(k=>`<b>`+escHtml(cfg[k].label)+`</b>: `+(meta[k]?escHtml(meta[k].file)+' | sheet: '+escHtml(meta[k].sheet)+' | rows: '+meta[k].rows+' | '+escHtml(meta[k].mode):'not loaded')).join(`<br>`)}
function loadDemo(){document.querySelectorAll('.upload input[type=file]').forEach(x=>x.value='');dataMode='demo';Object.keys(demo).forEach(k=>{raw[k]=JSON.parse(JSON.stringify(demo[k]));meta[k]={file:'demo '+k,sheet:cfg[k].sheet,rows:raw[k].length,score:100,mode:'demo'}});setStatus();
  recalculate()}
function indexBy(rows,key){let m=new Map();rows.forEach(r=>{let id=clean(get(r,[key]));if(!id)return;if(!m.has(id))m.set(id,[]);m.get(id).push(r)});return m}function isCommit(r){return upper(get(r,[cfg.fcid.commit]))==='COMMIT'}
const invalidFcidKeys=new Set(['NOFID','NOFCID','NA','NULL','UNKNOWN','NOTASSIGNED']);
function isValidFcidKey(value){let key=norm(clean(value)).toUpperCase();return Boolean(key)&&!invalidFcidKeys.has(key)}
function hasAssignedFabSegment(row){let value=upper(get(row,['Fab Segment','Fab_Segment']));return Boolean(value)&&value!=='NOT ASSIGNED'}
function prepareCommitFcids(rows){let selected=new Map(),audit={rawCommitRows:0,validCommitRows:0,invalidFidRows:0,distinctCommitFcids:0,duplicateFcidGroups:0,duplicateRowsRemoved:0,fabSegmentPreferred:0};for(const row of rows){if(!isCommit(row))continue;audit.rawCommitRows++;
  let fid=get(row,[cfg.fcid.key]);if(!isValidFcidKey(fid)){audit.invalidFidRows++;continue}audit.validCommitRows++;let entry=selected.get(fid);if(!entry){selected.set(fid,{row,preferred:null,count:1});continue}audit.duplicateRowsRemoved++;entry.count++;if(entry.count===
  2){audit.duplicateFcidGroups++;entry.preferred=hasAssignedFabSegment(entry.row)}if(!entry.preferred&&hasAssignedFabSegment(row)){entry.row=row;entry.preferred=true;audit.fabSegmentPreferred++}}let prepared=[];selected.forEach(entry=>prepared.push(entry.row));
  audit.distinctCommitFcids=prepared.length;return {rows:prepared,audit}}
function exclusionReason(row){let s=upper(get(row,[cfg.fcid.system]));if(/NCFUPG/.test(s))return 'NCFUPG';if((s.includes('SW')&&(s.includes('UPG')||s.includes('UPGRADE'))))return 'SW-UPGRADE';if(/SOFTWARE/.test(s))return 'Software';
  if(/CUST[_ ]?BUS[_ ]?FLEX/.test(s))return 'Cust_Bus_Flex';if(/CUST[_ ]?INCENTIVE/.test(s))return 'Cust_Incentive';return ''}function rowText(r){return Object.values(r||{}).join(' ')}function isDropSlotCtl(r,src){
    let change = norm(get(r,[src === "ctl" ? cfg.ctl.change : cfg.ctlNsr.change, "Change Type", "Change Cat"])); let status = upper(get(r,[src === "ctl" ? cfg.ctl.status : cfg.ctlNsr.status, "Status", "CTL Status"]));
    if(status.includes("REJECT") || status.includes("CANCEL"))return false;
    return change === 'dropslot' || change === 'drop' || status.includes("DROP");
}function isSlotCtl(r,src){
    let change = norm(get(r,[src === "ctl" ? cfg.ctl.change : cfg.ctlNsr.change, "Change Type", "Change Cat"])); let status = upper(get(r,[src === "ctl" ? cfg.ctl.status : cfg.ctlNsr.status, "Status", "CTL Status"]));
    if(status.includes("DROP") || status.includes("REJECT") || status.includes("CANCEL"))return false;
    return change === 'slot';
}function isConfigCtl(r,src){let change=upper(get(r,[src==='ctl'?cfg.ctl.change:cfg.ctlNsr.change,'Change Type','Change Cat']));let text=rowText(r).toUpperCase();let status=upper(get(r,[src==='ctl'?cfg.ctl.status:cfg.ctlNsr.status,'Status','CTL Status']));if(['REJECTED','CANCELLED',
  'CANCELED'].includes(status))return false;let docLocal=/DOC ONLY|LOCAL NSR|LNSR/.test(text);return (change.includes('CONFIG')||/CONFIG CHG|CUSTOMER CONFIG|PG CONFIG|GOPS CONFIG/.test(text))&&!docLocal}function firstDate(row,names){for(const n of names){let d=toDate(get(row,[n]));
  if(d)return d}return null}function slotSubmitDate(r,src){return firstDate(r,src==='ctl'?[cfg.ctl.sops,'Sales Ops Approval Date','Create Date','Creation Date']:[cfg.ctlNsr.sops,'Sales Ops Approval Date','Create Date'])}function srdAtSlot(r,src){return firstDate(r,src===
  'ctl'?[cfg.ctl.srdSops,cfg.ctl.srd,'Sales Ops RD','SRD @ Current']:[cfg.ctlNsr.srd,cfg.ctlNsr.fieldCrd,'SRD','Sales Ops RD'])}function thresholdWeeks(base,slotRow){let t=(get(slotRow,[cfg.ctl.mfg,cfg.ctlNsr.mfg,'Mfg Upgrade Type'])||get(base,['Mfg Upgrade Type'])||'').toUpperCase();
  if(!t||t.includes('NOT ASSIGNED'))return 20;if(t.includes('NO LABOR'))return 12;if(t.includes('LABOR')||t.includes('ROUTER'))return 20;return 20}
function getQuarterFromMcsd(mcsd){
    const date=new Date(mcsd),year=date.getFullYear(),endMonth=Math.floor(date.getMonth()/3)*3+3,cutoff=new Date(year,endMonth,0);
    // The quarter ends on its last Friday, inclusive; later MCSD dates belong to the next quarter.
    cutoff.setDate(cutoff.getDate()-(cutoff.getDay()+2)%7); const carried=date.getMonth()===cutoff.getMonth()&&date.getDate()>cutoff.getDate(); return {quarter:qFromDate(carried?new Date(year,endMonth,1):date),adjustment:carried?'Quarter End Carry-Over':''};
}
function deriveQuarter(base,mcsd){
    let cal = get(base,[cfg.fcid.quarter]);
    if(mcsd){let result = getQuarterFromMcsd(mcsd); return {quarter: result.quarter, source: "MCSD", adjustment: result.adjustment};}
    let target = firstDate(base,[cfg.fcid.targetShip]);
    if(target){let tq = qFromDate(target); if(tq !== cal){return {quarter: tq, source: "Target Ship", adjustment: ""};} return {quarter: cal, source: "Calendar Quarter", adjustment: ""};}
    let rd = firstDate(base,[cfg.fcid.rd]);
    if(rd){return {quarter: qFromDate(rd), source: "Sales Ops RD", adjustment: ""};}
    return {quarter: cal, source: "Calendar Quarter", adjustment: ""};
}
let initiatorIndex=new Map;
function initiatorsOf(fid){return initiatorIndex.get(fid)||['[Undefined]']}
function initiatorValue(record){
  const key=exactAliasKey(record,['Initiator','Initiated By','Requester','Created By']),name=key?clean(record[key]):'';
  return /^(undefined|\[undefined\]|null|n\/a)$/i.test(name)?'':name;
}
function primaryInitiator(actions){
  const documents=new Map,counts=new Map;
  for(const [document,name] of actions){
    if(!name)continue;
    const previous=documents.get(document);
    if(previous&&norm(previous)!==norm(name))return '';
    documents.set(document,name);
  }
  for(const name of documents.values()){
    const key=norm(name),entry=counts.get(key)||{name,count:0};entry.count++;counts.set(key,entry);
  }
  let highest=0,owner='';
  for(const {name,count} of counts.values()){
    if(count>highest){highest=count;owner=name}else if(count===highest)owner='';
  }
  return owner;
}
function recalculate(){
  M.idx={ctl:indexBy(raw.ctl,cfg.ctl.key),ctlNsr:indexBy(raw.ctlNsr,cfg.ctlNsr.key),nsr:indexBy(raw.nsr,cfg.nsr.key),nce:indexBy(raw.nce,cfg.nce.key)};
  const nsrNames=new Map,accountNames=new Map;
  for(const name of raw.nsr.map(initiatorValue))if(name&&!/^[A-Z0-9]+$/.test(name)&&!nsrNames.has(norm(name)))nsrNames.set(norm(name),name);
  // Learn aliases only from single-initiator FCIDs, corroborated by other FCIDs without conflicts.
  for(const [fid,records] of M.idx.ctlNsr){
    const accounts=uniq(records.map(record=>norm(initiatorValue(record))));
    const names=uniq((M.idx.nsr.get(fid)||[]).map(record=>nsrNames.get(norm(initiatorValue(record)))));
    if(accounts.length!==1||names.length!==1)continue;
    const account=accounts[0],name=names[0],previous=accountNames.get(account);
    accountNames.set(account,{name:previous&&previous.name!==name?'':name,count:(previous?.count||0)+1});
  }
  const nameOf=name=>{
    const key=norm(name),match=accountNames.get(key);
    return nsrNames.get(key)||(match?.count>1?match.name:'')||(/^[a-z0-9]+$/i.test(name)?'':name);
  };
  initiatorIndex=new Map;
  const fids=new Set([...M.idx.ctl.keys(),...M.idx.ctlNsr.keys(),...M.idx.nsr.keys()]);
  for(const fid of fids){
    const ctls=[...(M.idx.ctlNsr.get(fid)||[]),...(M.idx.ctl.get(fid)||[])];
    const actions=[...ctls.map((record,index)=>['ctl:'+(get(record,['CTL ID','CTL#'])||'row:'+index),initiatorValue(record)]),
      ...(M.idx.nsr.get(fid)||[]).map((record,index)=>['nsr:'+(nsrNumber(record)||'row:'+index),initiatorValue(record)])];
    // Keep unmapped accounts distinct while counting; only the final owner may become Undefined.
    const owner=primaryInitiator(actions.map(([document,name])=>[document,nameOf(name)||name]));
    const name=nameOf(owner);if(name)initiatorIndex.set(fid,[name]);
  }
  let prepared=prepareCommitFcids(raw.fcid),commit=prepared.rows;M.fcidAudit=prepared.audit;M.originalCommit=
  prepared.audit.rawCommitRows;M.excluded=[];let base=[];commit.forEach(b=>{let reason=exclusionReason(b);if(reason){M.excluded.push({fid:get(b,[cfg.fcid.key]),quarter:get(b,[cfg.fcid.quarter]),account:accountOf(b),customer:get(b,[cfg.fcid.customer]),product:get(b,[cfg.fcid.product]),
  system:get(b,[cfg.fcid.system]),excludedReason:reason});}else base.push(b)});let rows=[],risks=[];base.forEach(b=>{let fid=get(b,[cfg.fcid.key]);let ctls=M.idx.ctl.get(fid)||[],bridges=M.idx.ctlNsr.get(fid)||[];let mc=firstDate(b,[cfg.fcid.ship])||
  firstDate((ctls.find(r=>firstDate(r,[cfg.ctl.mcsd]))||{}),[cfg.ctl.mcsd])||firstDate((bridges.find(r=>firstDate(r,[cfg.ctlNsr.mcsd]))||{}),[cfg.ctlNsr.mcsd]);let q=deriveQuarter(b,mc);let rd=firstDate(b,[cfg.fcid.rd]);let target=firstDate(b,[cfg.fcid.targetShip]);
  let slotEvents = [...ctls.map(r=>({r,src:"ctl"})), ...bridges.map(r=>({r,src:"ctlNsr"}))] .filter(x => isSlotCtl(x.r,x.src) || isDropSlotCtl(x.r,x.src));
slotEvents.sort((a,b)=>{
    let da = slotSubmitDate(a.r,a.src) || firstDate(a.r,["Last Modified Date", "Modified Date", "Create Date", "Creation Date"]) || new Date(0); let db = slotSubmitDate(b.r,b.src) || firstDate(b.r,["Last Modified Date", "Modified Date", "Create Date", "Creation Date"]) || new Date(0);
    return db - da;
});
let currentSlotEvent = slotEvents[0]; let currentSlotDropped = currentSlotEvent && isDropSlotCtl(currentSlotEvent.r,currentSlotEvent.src); let effectiveSlot = currentSlotEvent && !currentSlotDropped && isSlotCtl(currentSlotEvent.r,currentSlotEvent.src) ? currentSlotEvent : null;
let hasSlot = Boolean(effectiveSlot); let slotState = currentSlotDropped ? "Dropped Slot" : (hasSlot ? "Active Slot" : "No Slot");let preStatus='No Slot CTL',preWeeks=null,threshold=20,slotId='',ctlSrd='';if(effectiveSlot){let sDate=slotSubmitDate(effectiveSlot.r,effectiveSlot.src),
  srd=srdAtSlot(effectiveSlot.r,effectiveSlot.src)||rd;ctlSrd=srd?srd.toLocaleDateString():'';threshold=thresholdWeeks(b,effectiveSlot.r);preWeeks=weeks(sDate,srd);preStatus=(preWeeks!==null&&preWeeks>=threshold)?'On Time':'Pre Late Risk';slotId=get(effectiveSlot.r,[cfg.ctl.ctlId,
  cfg.ctlNsr.ctlId,'CTL ID'])}if(currentSlotDropped){preStatus = "No Slot CTL"; slotId = get(currentSlotEvent.r, [cfg.ctl.ctlId, cfg.ctlNsr.ctlId, "CTL ID"]);}let configRows=[...ctls.map(r=>({r,src:'ctl'})),...bridges.map(r=>({r,src:'ctlNsr'}))].filter(x=>isConfigCtl(x.r,x.src));
  let postRiskRows=configRows.filter(x=>{let d=slotSubmitDate(x.r,x.src),w=weeks(d,mc);return d&&mc&&w!==null&&w>=0&&w<=20});let postStatus= postRiskRows.length ? 'Post Config CTL Risk' : (mc ? 'Projected Stable' : 'Projected Stable + MCSD Pending');
let pendingNsr =
(M.idx.nsr.get(fid) || []) .filter(n => {
    let status = (get(n,[cfg.nsr.status]) + ' ' + get(n,[cfg.nsr.workflow])).toUpperCase();
    // Closed NSR
    if(status.includes('COMPLETED') || status.includes('REJECTED') || status.includes('CANCELLED') || status.includes('CANCELED')){return false;}
    // Everything else is considered Open NSR
    return true;
});let watchNsr = pendingNsr.filter(n=>{let nsrNo = nsrNumber(n).toUpperCase(); return !nsrNo.startsWith('D');}); let nces = M.idx.nce.get(fid) || [];let row={fid,quarter:q.quarter,quarterSource:q.source,quarterAdjustment:q.adjustment || '',calendarQuarter:get(b,
  [cfg.fcid.quarter]),targetShip:target?target.toLocaleDateString():'',rd:rd?rd.toLocaleDateString():'',mcsd:mc?mc.toLocaleDateString():'',ctlSrd,slotState,account:accountOf(b),customer:get(b,[cfg.fcid.customer]),product:get(b,[cfg.fcid.product]),system:get(b,[cfg.fcid.system]),
  hasSlot,slotId,preStatus,preWeeks:preWeeks===null?'':preWeeks.toFixed(1), weeksToPreStability: preWeeks===null ? '' : Math.round(preWeeks),
weeksToPostStability: mc ? Math.round(weeks(new Date(), mc)) : '', threshold, postStatus,postRiskCtl:postRiskRows.map(x=>get(x.r,[cfg.ctl.ctlId,cfg.ctlNsr.ctlId,'CTL ID'])).join(', '),pendingNsr: pendingNsr.length, watchNsrCount: watchNsr.length,
openNsrTitle: pendingNsr.map(x => get(x,[cfg.nsr.title])), openNsrNumbers: pendingNsr.map(nsrNumber),openNsrStatus: pendingNsr.map(x => get(x,[cfg.nsr.workflow])),nce:nces.length};rows.push(row);function push(type,priority,
  msg){risks.push({fid,quarter:row.quarter,quarterSource:row.quarterSource,account:row.account,rowCustomer:row.customer,product:row.product,type,priority,msg})}if(preStatus==='No Slot CTL')push('No Slot CTL','P1','No valid Slot CTL found.');if(preStatus===
  'Pre Late Risk')push('Pre Late Risk','P1','Slot CTL submitted '+row.preWeeks+' weeks before RD/SRD. Threshold is '+threshold+' weeks.');if(!mc&&hasSlot)push('MCSD Pending','P2','Valid slot exists but MCSD is blank.');if(postRiskRows.length)push('Post Config CTL Risk','P1',
  'Qualifying config CTL submitted inside 20-week MCSD window: '+row.postRiskCtl);if(watchNsr.length)push('NSR Appeal Watch','P3','Pending NSR may become appeal/watch driver.');if(nces.length)push('OOBQ / NCE Signal','P3','Matched NCE / FA quality signal found.');});M.rows=rows;
  M.risks=risks;sanitizeModel();buildFilters();renderAll()}
function selected(id){return [...document.getElementById(id).selectedOptions].map(o=>o.value)}function matchMulti(v,a){return !a.length||a.includes(v)}
function matchesSearch(row,term){
  if(!term||Object.values(row).join(' ').toLowerCase().includes(term))return true;
  return ['ctl','ctlNsr'].some(source=>(M.idx[source]?.get(row.fid)||[]).some(record=>get(record,[cfg[source].ctlId]).toLowerCase().includes(term)))||
    (M.idx.nsr?.get(row.fid)||[]).some(record=>nsrNumber(record).toLowerCase().includes(term));
}
function filteredRows(source=M.rows){let i=selected('fInitiator'),q=selected('fQuarter'),a=selected('fAccount'),c=selected('fCustomer'),s=(fSearch.value||'').toLowerCase();
  return source.filter(r=>(!i.length||initiatorsOf(r.fid).some(name=>i.includes(name)))&&matchMulti(r.quarter,q)&&matchMulti(r.account,a)&&matchMulti(r.customer,c)&&matchesSearch(r,s))}function filteredRisks(){let t=selected('fRisk'),ids=new Set(filteredRows().map(r=>r.fid));
  return M.risks.filter(r=>ids.has(r.fid)&&matchMulti(r.type,t))}
const chOn=new Set;
function chSync(id){if(chOn.has(id))jQuery(document.getElementById(id)).trigger('chosen:updated')}
function fill(id,arr,compareFn){let el=document.getElementById(id),cur=selected(id),values=uniq(arr).filter(x=>x);compareFn?values.sort(compareFn):values.sort();el.innerHTML=values.map(x=>`<option>`+escHtml(x)+`</option>`).join('');[...el.options].forEach(o=>o.selected=cur.includes(o.value));
  chSync(id)}
function quarterTimeComparator(a,b){return quarterSortValue(a)-quarterSortValue(b)||String(a).localeCompare(String(b))}
function buildFilters(){fill('fInitiator',[...M.rows,...M.excluded].flatMap(r=>initiatorsOf(r.fid)));fill('fQuarter',M.rows.map(r=>r.quarter),quarterTimeComparator);fill('fAccount',M.rows.map(r=>r.account));fill('fCustomer',M.rows.map(r=>r.customer))}
function selSet(id,on,run=1){[...document.getElementById(id).options].forEach(o=>o.selected=on);chSync(id);if(run)applyFilters()}function selectAll(id){selSet(id,true)}function clearSel(id){selSet(id,false)}function resetFilters(){['fInitiator','fQuarter','fAccount','fCustomer',
  'fRisk'].forEach(id=>selSet(id,false,0));fSearch.value='';applyFilters()}function h(a){return `<thead><tr>`+a.map(x=>`<th>`+x+`</th>`).join('')+`</tr></thead><tbody>`}function tr(a){return `<tr>`+a.map(x=>`<td>`+String(x??'')+`</td>`).join('')+
  `</tr>`}function pill(v){let cls=/P1|Risk|No Slot/.test(v)?'bad':/P2|Pending|Late/.test(v)?'warnp':/Stable|On Time|Yes/.test(v)?'ok':/P3|Watch|NCE/.test(v)?'purplep':'info';return '<span class="pill '+cls+'">'+v+`</span>`}
function renderAll(){applyFilters();renderMapping();renderRules();renderUAT()}
function labelSel(id,all){let a=selected(id);return a.length?a.length+' '+all+' selected':'All '+all}
function quarterSortValue(q){let m=String(q||"").match(/(QMar|QJun|QSep|QDec)'(\d+)/);return m?+m[2]*10+{QMar:1,QJun:2,QSep:3,QDec:4}[m[1]]:999999}
function riskCounts(){let rows=filteredRows(),ids=new Set(rows.map(r=>r.fid)),rks=filteredRisks().filter(r=>ids.has(r.fid)),types=['Pre Late Risk','No Slot CTL','Post Config CTL Risk','MCSD Pending','NSR Appeal Watch','OOBQ / NCE Signal'];return types.map(t=>[t,
  uniq(rks.filter(r=>r.type===t).map(r=>r.fid)).length,t.includes('Post')||t.includes('No Slot')?'#dc2626':t.includes('MCSD')?'#f59e0b':t.includes('NCE')?'#7c3aed':'#2563eb'])}function renderExcludeBars(){let counts={};M.excluded.forEach(x=>counts[x.excludedReason]=
  (counts[x.excludedReason]||0)+1);let max=Math.max(1,...Object.values(counts));excludeBars.innerHTML=Object.entries(counts).map(x=>`<div class="barRow"><b>`+x[0]+`</b><div class="bar"><span style="width:`+(x[1]/max*100)+`%;background:#dc2626"></span></div><strong>`+x[1]+
  `</strong></div>`).join('')||`<div class="small">No exclusions detected.</div>`}
function renderRules(){rulesTbl.innerHTML=h(['Rule','Description'])+[['Quarter Governance','Priority: MCSD -> Target Ship -> Sales Ops RD -> Calendar Quarter.'],['Quarter Source','All rows retain Quarter Source traceability.'],['Population',
  'Commit FCIDs only; exclude NCFUPG, Software, SW-UPGRADE, Cust_Bus_Flex.'],['FCID Deduplication','One Commit row per valid FCID; prefer the first row with an assigned Fab Segment, otherwise keep the first row.'],['Account Group',
  'Retain a valid source Account Group; use exact or wildcard Fab mapping only when the source value is missing.'],['Pre-Slot Projection','Slot CTL submit date compared with RD / SRD at slot.'],['Post-Slot Projection','Config CTL inside 20-week MCSD window.'],['Filters',
  'Initiator, Quarter, Account, Customer, and Risk Type use multi-select logic.'],['Export','All exports use current filtered population.']].map(r=>tr(r)).join('')+`</tbody>`}
function renderUAT(){uatTbl.innerHTML=h(['UAT','Condition','Expected','Current Support'])+[['UAT-09','MCSD blank; Calendar Q = QSep; Target Ship = Nov 20','Quarter = QDec; Source = Target Ship','Supported'],['UAT-10','MCSD blank; Target Ship blank; RD = Nov 05',
  'Quarter = QDec; Source = Sales Ops RD','Supported'],['UAT-11','MCSD = Jan 15 2027','Quarter = QMar 27; Source = MCSD','Supported'],['UAT-12','System Description contains SW-UPGRADE','Excluded; reason = SW-UPGRADE','Supported'],['UAT-13','Multi-select AT1 + AT3',
  'All KPI recalc on combined population','Supported'],['UAT-14','Duplicate Commit rows for one FCID','One FCID retained; assigned Fab Segment is preferred','Supported'],['UAT-15','FID is blank, NO_FID, N/A, NULL, or UNKNOWN','Excluded before governance calculation',
  'Supported']].map(r=>tr(r)).join('')+`</tbody>`}
function escHtml(v){return String(v ?? '') .replace(/&/g,'&amp;') .replace(/</g,'&lt;') .replace(/>/g,'&gt;') .replace(/"/g,'&quot;') .replace(/'/g,'&#39;');}
function isOpenNsrRecord(n){let status = (get(n,[cfg.nsr.status]) + ' ' + get(n,[cfg.nsr.workflow])).toUpperCase(); if(status.includes('COMPLETED') || status.includes('REJECTED') || status.includes('CANCELLED') || status.includes('CANCELED')){return false;} return true;}
function nsrWorkflow(n){return get(n,[cfg.nsr.workflow]) || get(n,[cfg.nsr.status]);}
function nsrTitle(n){return get(n,[cfg.nsr.title, 'NSR Title', 'Task Name']);}
function nsrNumber(n){let key=exactAliasKey(n,[cfg.nsr.nsr,'NSR #','NSR Number','NSR No']);return key?clean(n[key]):'';}
function nceTitle(n){return get(n,[cfg.nce.title,'Title']);}
function nceState(n){return get(n,[cfg.nce.state,'NCe State']);}
function nceImpact(n){return get(n,[cfg.nce.impact,'Impact']);}
function nceOpenDate(n){return get(n,[cfg.nce.openDate,'Open Date']);}
function nceAccuracyType(n){let group = get(n,[cfg.nce.damageGroup,'Damage Group']); let code = get(n,[cfg.nce.damageCode,'Damage Code']); if(group && code) return group + ' / ' + code; return group || code || 'Config Accuracy Issue';}
function ncePartNumber(n){return get(n,[cfg.nce.partNumber,'Part_Number','Part Number']);}
function ncePartDescription(n){return get(n,[cfg.nce.partDescription,'Part_Description','Part Description']);}
function renderNsrCards(records,emptyText){
    if(!records.length){return `<div class="small">`+escHtml(emptyText)+`</div>`;}
    return records.map(n=>{
        let no = nsrNumber(n) || '-'; let title = nsrTitle(n) || '-'; let workflow = nsrWorkflow(n) || '-';
        return (`\n<div class="mini" style="margin-bottom:10px">\n<small>${escHtml(no)}</small>\n<b style="font-size:14px;display:block;white-space:normal;line-height:1.35">\n${escHtml(title)}` +
  `\n</b>\n<div class="small">\n${escHtml(workflow)}\n</div>\n</div>\n`);
    }).join('');
}
function renderAccuracyCards(records){
    if(!records.length){return `<div class="small">No historical configuration accuracy issues found for this FCID.</div>`;}
    return records.map(n=>{
        let date = nceOpenDate(n) || '-'; let impact = nceImpact(n) || '-'; let state = nceState(n) || '-'; let title = nceTitle(n) || '-'; let accuracyType = nceAccuracyType(n) || '-'; let pn = ncePartNumber(n) || '-'; let desc = (ncePartDescription(n) || '-') .slice(0,80);
        return (`\n<div class="mini" style="margin-bottom:10px">\n<small>${escHtml(date)} | ${escHtml(impact)} | ${escHtml(state)}` +
  `</small>\n<b style="font-size:14px;display:block;white-space:normal;line-height:1.35">\n${escHtml(title)}` +
  `\n</b>\n<div class="small" style="margin-top:6px">\n<b>Accuracy Type:</b> ${escHtml(accuracyType)}\n<br>\n<b>Part Number:</b> ${escHtml(pn)}` +
  `\n<br>\n<b>Part Description:</b> ${escHtml(desc)}\n</div>\n</div>\n`);
    }).join('');
}
function fileSuffix(){let q=selected('fQuarter').join('_')||'AllQ',a=selected('fAccount').join('_')||'AllAcct',c=selected('fCustomer').join('_')||'AllCust';return (q+'_'+a+'_'+c).replace(/[^a-z0-9_'-]/ig,'_').slice(0,120)}
function exportWorkbook(name,sheets){
  if(!window.XLSX){alert('XLSX library not available. Please save source files as CSV or enable CDN.');return}
  const wb=XLSX.utils.book_new();
  Object.entries(sheets).forEach(([sheet,rows])=>{
    const data=rows.length?rows.map(row=>Object.fromEntries(Object.entries(row).map(([key,value])=>[key==='openNsrNumbers'?'NSR#':key,Array.isArray(value)?value.join('\n'):value]))):[{Notice:'No data'}];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),sheet.slice(0,31));
  });
  XLSX.writeFile(wb,name+'.xlsx');
}
/* Data readiness and confidence layer */
let dataMode='demo';
const readyRules={
    population:{required:['fcid'],minimumAll:['fcid']}, core:{required:['fcid','ctl','ctlNsr'],minimumAll:['fcid'],minimumAny:['ctl','ctlNsr']}, risk:{required:['fcid','ctl','ctlNsr','nsr','nce'],minimumAll:['fcid'],minimumAny:['ctl','ctlNsr','nsr','nce']},
    nsr:{required:['fcid','nsr'],minimumAll:['fcid','nsr']}, nce:{required:['fcid','nce'],minimumAll:['fcid','nce']}
};
const coreRisks=['Pre Late Risk','No Slot CTL','Post Config CTL Risk','MCSD Pending']; const sourceNames={fcid:'FCID Database',ctl:'CTL Report',ctlNsr:'CTL-NSR List',nsr:'NSR Log',nce:'NCE / FA Review'};
function loaded(k){return Boolean(meta[k])&&(dataMode==='demo'||meta[k].mode!=='demo')}
function ready(key){
    let rule=readyRules[key];
    if(dataMode==='demo')return {key,level:'demo',missing:[],label:'Demo Data'};
    let missing=rule.required.filter(k=>!loaded(k)); let minimumAll=(rule.minimumAll||[]).every(loaded); let minimumAny=!rule.minimumAny||(rule.minimumAny||[]).some(loaded);
    if(!minimumAll||!minimumAny)return {key,level:'unavailable',missing,label:'Waiting for data'};
    if(missing.length)return {key,level:'partial',missing,label:'Provisional'};
    return {key,level:'complete',missing:[],label:'Complete'};
}
function missingText(state){return state.missing.map(k=>sourceNames[k]).join(', ')}
function readyText(state){if(state.level==='complete')return 'Required sources are complete.'; if(state.level==='demo')return 'Demo data only.'; if(state.level==='partial')return 'Provisional; missing '+missingText(state)+'.'; return 'Unavailable; load '+missingText(state)+'.';}
function metric(value,stateOrKey){let state=typeof stateOrKey==='string'?ready(stateOrKey):stateOrKey;if(state.level==='unavailable')return '—';return String(value)}
function readyPill(state){let cls=state.level==='complete'?'ok':state.level==='partial'?'warnp':state.level==='demo'?'purplep':'info';return '<span class="pill '+cls+'">'+escHtml(state.label)+`</span>`}
function readyMini(label,value,key){let state=ready(key);return '<div class="mini" title="'+escHtml(readyText(state))+`"><small>`+escHtml(label)+`</small><b>`+escHtml(metric(value,state))+`</b><div class="small" style="margin-top:5px">`+readyPill(state)+`</div></div>`}
function setMetric(id,value,key){let el=document.getElementById(id);if(!el)return;let state=ready(key),description=readyText(state);el.innerText=metric(value,state);el.title=description;let note=document.getElementById(id+'Readiness');if(!note){note=document.createElement('div');
  note.id=id+'Readiness';note.className='small';note.style.margin='3px 0 5px';el.insertAdjacentElement('afterend',note)}note.innerHTML=readyPill(state);note.title=description}
function sanitizeModel(){
    let coreEvidence=dataMode==='demo'||loaded('ctl')||loaded('ctlNsr');
    if(!coreEvidence){
        M.rows.forEach(row=>{row.hasSlot=null;row.slotState='Unavailable';row.preStatus='Unavailable';row.postStatus='Unavailable';row.preWeeks='';row.weeksToPreStability='';row.weeksToPostStability='';row.slotId='';row.ctlSrd='';row.postRiskCtl=''});
        M.risks=M.risks.filter(r=>!coreRisks.includes(r.type));
    }
    if(!loaded('nsr')){M.rows.forEach(row=>{row.pendingNsr=null;row.watchNsrCount=null;row.openNsrTitle=[];row.openNsrNumbers=[];row.openNsrStatus=[]});M.risks=M.risks.filter(r=>r.type!=='NSR Appeal Watch')}
    if(!loaded('nce')){M.rows.forEach(row=>{row.nce=null});M.risks=M.risks.filter(r=>r.type!=='OOBQ / NCE Signal')}
}
function filteredMetrics(){let rows=filteredRows(),ids=new Set(rows.map(r=>r.fid)),risks=filteredRisks().filter(r=>ids.has(r.fid)),pop=rows.length,preOk=rows.filter(r=>r.preStatus==='On Time').length,postOk=rows.filter(r=>r.postStatus==='Projected Stable').length,
  slotReady=rows.filter(r=>r.hasSlot===true).length;return {pop,preOk,postOk,slotReady,preRisk:uniq(risks.filter(r=>['No Slot CTL','Pre Late Risk'].includes(r.type)).map(r=>r.fid)).length,postRisk:uniq(risks.filter(r=>r.type==='Post Config CTL Risk').map(r=>r.fid)).length,
  riskFcid:uniq(risks.map(r=>r.fid)).length,slotMcsd:rows.filter(r=>r.hasSlot===true&&r.mcsd).length,slotNoMcsd:rows.filter(r=>r.hasSlot===true&&!r.mcsd).length,noSlotCtl:rows.filter(r=>r.preStatus==='No Slot CTL').length,rows,risks}}
function renderFilterContext(){let m=filteredMetrics();filterContextBadge.innerText=[labelSel('fInitiator','Initiators'),labelSel('fQuarter','Quarters'),labelSel('fAccount','Accounts'),labelSel('fCustomer','Customers'),labelSel('fRisk','Risks')].join(' | ');filterContextMini.innerHTML=[readyMini('Current Population',
  m.pop,'population'),readyMini('Projected Pre-Slot',pct(m.preOk,m.pop),'core'),readyMini('Projected Post-Slot',pct(m.postOk,m.pop),'core'),readyMini('Slot Readiness',pct(m.slotReady,m.pop),'core'),readyMini('Pre Risk FCID',m.preRisk,'core'),readyMini('Post Risk FCID',m.postRisk,
  'core'),readyMini('Slot + No MCSD',m.slotNoMcsd,'core'),readyMini('No Slot CTL',m.noSlotCtl,'core')].join('')}
function renderFunnel(){let m=filteredMetrics(),items=[['Governance Population',m.pop,'population','baseline'],['Slot Ready',m.slotReady,'core','baseline'],['Pre On Time',m.preOk,'core',pct(m.preOk,m.pop)],['Post Stable',m.postOk,'core',pct(m.postOk,m.pop)],['Risk Free',
  m.pop-m.riskFcid,'risk',pct(m.pop-m.riskFcid,m.pop)]];funnel.innerHTML=items.map(x=>{let s=ready(x[2]);return '<div class="step" title="'+escHtml(readyText(s))+`"><small>`+x[0]+`</small><b>`+escHtml(metric(x[1],s))+`</b><span class="small">`+escHtml(metric(x[3],s))+`</span><div>`+
  readyPill(s)+`</div></div>`}).join(`<div class="arrow">›</div>`)}
function annotateTable(id,key){let table=document.getElementById(id);if(!table)return;let state=ready(key);if(table.caption)table.deleteCaption();if(state.level==='complete')return;let caption=table.createCaption();caption.style.cssText=
  'caption-side:top;text-align:left;padding:9px 10px;font-weight:800;background:'+(state.level==='partial'?'#fffbeb':state.level==='demo'?'#f5f3ff':'#eff6ff')+';color:#334155';caption.innerHTML=(state.level==='unavailable'?'— ':'')+escHtml(readyText(state));if(state.level===
  'unavailable'){let tbody=table.tBodies[0]||table.createTBody(),cols=table.tHead?.rows[0]?.cells.length||1;tbody.innerHTML=`<tr><td colspan="`+cols+'" class="small">— '+escHtml(readyText(state))+`</td></tr>`}}
function renderNsrTable(){let rows=filteredRows().filter(r=>r.pendingNsr>0);nsrTbl.innerHTML=h(['FCID','Quarter','Customer','Product','Open NSR Count','NSR Title','NSR Status','NSR#'])+rows.map(r=>tr([r.fid,r.quarter,r.customer,r.product,r.pendingNsr,r.openNsrTitle.join(`<br>`),
  r.openNsrStatus.join(`<br>`),r.openNsrNumbers.join(`<br>`)])).join('')+`</tbody>`;annotateTable('nsrTbl','nsr')}
function renderQuarterReady(){
    let rows=filteredRows(),risks=filteredRisks(),ids=new Set(rows.map(r=>r.fid));risks=risks.filter(r=>ids.has(r.fid));let qsum={},sourceRows={},state=ready('core');
    rows.forEach(r=>{let q=r.quarter||'(blank)';qsum[q]=qsum[q]||{q,pop:0,slotMcsd:0,slotNoMcsd:0,preOk:0,postOk:0,preRisk:[],postRisk:[]};qsum[q].pop++;if(r.hasSlot===true&&r.mcsd)qsum[q].slotMcsd++;if(r.hasSlot===true&&!r.mcsd)qsum[q].slotNoMcsd++;if(r.preStatus===
      'On Time')qsum[q].preOk++;if(r.postStatus==='Projected Stable')qsum[q].postOk++;let sourceKey=q+'|'+r.quarterSource;sourceRows[sourceKey]=sourceRows[sourceKey]||{Quarter:q,QuarterSource:r.quarterSource,Count:0};sourceRows[sourceKey].Count++});
    risks.forEach(r=>{let q=r.quarter||'(blank)';if(!qsum[q])return;if(['No Slot CTL','Pre Late Risk'].includes(r.type))qsum[q].preRisk.push(r.fid);if(r.type==='Post Config CTL Risk')qsum[q].postRisk.push(r.fid)});
    quarterTbl.innerHTML=h(['Quarter','Population','Slot + MCSD','Slot + No MCSD','Projected Pre-Slot','Projected Post-Slot','Pre Risk FCID','Post Risk FCID'])+Object.values(qsum).sort((a,b)=>quarterSortValue(a.q)-quarterSortValue(b.q)).map(x=>tr([x.q,metric(x.pop,state),
      metric(x.slotMcsd,state),metric(x.slotNoMcsd,state),metric(pct(x.preOk,x.pop),state),metric(pct(x.postOk,x.pop),state),metric(uniq(x.preRisk).length,state),metric(uniq(x.postRisk).length,state)])).join('')+`</tbody>`;
    quarterSourceTbl.innerHTML=h(['Quarter','Quarter Source','FCID Count'])+Object.values(sourceRows).map(x=>tr([x.Quarter,x.QuarterSource,metric(x.Count,state)])).join('')+`</tbody>`; let m=filteredMetrics();projectionMini.innerHTML=[readyMini('Population',m.pop,'population'),
      readyMini('Pre On Time',m.preOk+' / '+m.pop,'core'),readyMini('Post Stable',m.postOk+' / '+m.pop,'core'),readyMini('Slot Readiness',m.slotReady+' / '+m.pop,'core')].join('');
}
function renderTables(){let rows=filteredRows(),risks=filteredRisks(),ids=new Set(rows.map(r=>r.fid));risks=risks.filter(r=>ids.has(r.fid));watchTbl.innerHTML=h(['Priority','FCID','Quarter','Quarter Source','Account','Customer','Product','Risk','Message'])+
  risks.filter(r=>r.priority==='P1').map(r=>tr([pill(r.priority),r.fid,r.quarter,r.quarterSource,r.account||'',r.rowCustomer,r.product,pill(r.type),r.msg])).join('')+`</tbody>`;stabilityTbl.innerHTML=h(['FCID','Quarter','Target Ship','RD','MCSD','Account','Customer','Product',
  'Slot CTL','Weeks to Pre-Stability','Pre Status','Weeks to Post-Stability','Post Status','Config CTL Risk'])+rows.map(r=>tr([r.fid,r.quarter,r.targetShip,r.rd,r.mcsd,r.account,r.customer,r.product,r.slotId,r.weeksToPreStability,pill(r.preStatus),r.weeksToPostStability,
  pill(r.postStatus),r.postRiskCtl])).join('')+`</tbody>`;slotTbl.innerHTML=h(['FCID','Quarter','Source','Adjustment','Account','Customer','Product','Slot State','Slot Ready','Slot CTL','MCSD','RD','CTL SRD','Pre Status'])+rows.map(r=>tr([r.fid,r.quarter,r.quarterSource,
  r.quarterAdjustment||'',r.account||'',r.customer,r.product,r.slotState||'',r.hasSlot?'Y':'N',r.slotId,r.mcsd,r.rd,r.ctlSrd,pill(r.preStatus)])).join('')+`</tbody>`;riskTbl.innerHTML=h(['Priority','FCID','Quarter','Source','Account','Customer','Product','Risk Type','Message'])+
  risks.map(r=>tr([pill(r.priority),r.fid,r.quarter,r.quarterSource,r.account||'',r.rowCustomer,r.product,pill(r.type),r.msg])).join('')+`</tbody>`;renderQuarterReady();[['watchTbl','core'],['stabilityTbl','core'],['slotTbl','core'],['riskTbl','risk'],['quarterTbl','core'],
  ['quarterSourceTbl','core']].forEach(x=>annotateTable(...x))}
function renderHeat(){let state=ready('core'),rows=filteredRows();if(state.level==='unavailable'){heatmap.innerHTML=`<div class="head" style="grid-column:1/-1">— `+escHtml(readyText(state))+`</div>`;return}let qs=uniq(rows.map(r=>r.quarter)).sort((a,
  b)=>quarterSortValue(a)-quarterSortValue(b)).slice(0,5);while(qs.length<5)qs.push('-');let html=(state.level!=='complete'?`<div class="head" style="grid-column:1/-1">`+escHtml(readyText(state))+`</div>`:'')+`<div class="head">Metric</div>`+qs.map(q=>`<div class="head">`+q+
  `</div>`).join('');[['Pre Projection','pre'],['Post Projection','post'],['Slot Readiness','slot'],['MCSD Pending','pending']].forEach(m=>{html+=`<div class="head">`+m[0]+`</div>`;qs.forEach(q=>{if(q==='-'){html+=`<div class="head">-</div>`;return}let rr=rows.filter(r=>r.quarter===
  q),den=Math.max(1,rr.length),val=0;if(m[1]==='pre')val=rr.filter(r=>r.preStatus==='On Time').length/den*100;if(m[1]==='post')val=rr.filter(r=>r.postStatus==='Projected Stable').length/den*100;if(m[1]==='slot')val=rr.filter(r=>r.hasSlot===true).length/den*100;if(m[1]==='pending')val=
  rr.filter(r=>r.hasSlot===true&&!r.mcsd).length/den*100;let cls=m[1]==='pending'?(val<=5?'g':val<=10?'y':'r'):(val>=95?'g':val>=90?'y':'r');html+='<div class="'+cls+'">'+escHtml(metric(val.toFixed(1)+'%',state))+`</div>`})});heatmap.innerHTML=html}
function riskReady(type){if(coreRisks.includes(type))return ready('core');if(type==='NSR Appeal Watch')return ready('nsr');return ready('nce')}
function renderDriverBars(){let items=riskCounts(),max=Math.max(1,...items.map(x=>x[1]));driverBars.innerHTML=items.map(x=>{let s=riskReady(x[0]),w=s.level==='unavailable'?0:x[1]/max*100;return '<div class="barRow" title="'+escHtml(readyText(s))+`"><b>`+x[0]+
  `</b><div class="bar"><span style="width:`+w+'%;background:'+x[2]+`"></span></div><strong>`+escHtml(metric(x[1],s))+`</strong></div>`}).join('')}
function renderDriverCards(){driverCards.innerHTML=riskCounts().map(x=>{let s=riskReady(x[0]),status=s.level==='unavailable'?'Waiting':s.level==='partial'?'Provisional':s.level==='demo'?'Demo':x[1]?'Open':'Clear',cls=s.level==='unavailable'?'info':s.level===
  'partial'?'warnp':s.level==='demo'?'purplep':x[1]?'warnp':'ok';return '<div class="card" title="'+escHtml(readyText(s))+`"><h3>`+x[0]+`</h3><div style="font-size:28px;font-weight:900;color:`+x[2]+'">'+escHtml(metric(x[1],s))+`</div><span class="pill `+cls+'">'+status+
  `</span></div>`}).join('')}
function renderFcid360(){let rows=filteredRows();if(!rows.length){fcidBasic.innerHTML='';fcidAction.innerHTML=metric('No FCID','population');fcidTbl.innerHTML='';return}if(!fcidInput.value||!rows.some(x=>x.fid===clean(fcidInput.value)))fcidInput.value=rows[0].fid;
  let fid=clean(fcidInput.value),r=rows.find(x=>x.fid===fid)||rows[0],nsrRecords=M.idx.nsr.get(r.fid)||[],openNsrs=nsrRecords.filter(isOpenNsrRecord),historicalNsrs=nsrRecords.filter(n=>!isOpenNsrRecord(n)),nceRecords=M.idx.nce.get(r.fid)||[],
  historicalAccuracy=nceRecords.filter(n=>{let state=nceState(n).toUpperCase();return state.includes('CLOSED')||state.includes('CANCEL')}),basic=[['FCID',r.fid,'population'],['Quarter',r.quarter,'core'],['Customer',r.customer,'population'],['Product',r.product,'population'],
  ['Target Ship',r.targetShip||'-','population'],['MCSD',r.mcsd||'-','core'],['Slot State',r.slotState||'-','core'],['Pre Status',r.preStatus,'core'],['Post Status',r.postStatus,'core'],['Open NSR',openNsrs.length,'nsr'],['Historical Accuracy Issues',historicalAccuracy.length,'nce']];
  fcidBasic.innerHTML=basic.map(x=>readyMini(x[0],x[1],x[2])).join('');let risks=filteredRisks().filter(x=>x.fid===r.fid),riskState=ready('risk');if(riskState.level==='unavailable')fcidAction.innerHTML='— '+escHtml(readyText(riskState));else{let warning=riskState.level===
  'partial'?`<div class="warnbox">`+escHtml(readyText(riskState))+`</div>`:riskState.level==='demo'?`<div class="warnbox">Demo data only.</div>`:'';fcidAction.innerHTML=warning+(risks.length?`<b>Action focus:</b><br>`+risks.map(x=>'• '+escHtml(x.type)+': '+
  escHtml(x.msg)).join(`<br>`):(riskState.level==='complete'?'No current early warning risk found for this FCID.':'No risk found in currently loaded sources; this is not a final result.'))}let nsrState=ready('nsr'),nceReady=ready('nce'),nsrUnavailable=`<div class="small">`+
  escHtml(metric('',nsrState))+' '+escHtml(readyText(nsrState))+`</div>`,nceUnavailable=`<div class="small">`+escHtml(metric('',nceReady))+' '+escHtml(readyText(nceReady))+`</div>`;fcidTbl.innerHTML=
  `<thead><tr><th style="width:22%">Section</th><th>Details</th></tr></thead><tbody><tr><td><b>Open NSR (`+escHtml(metric(openNsrs.length,nsrState))+`)</b></td><td>`+(nsrState.level==='unavailable'?nsrUnavailable:renderNsrCards(openNsrs,'No open NSR found for this FCID.'))+
  `</td></tr>`+(nsrState.level!=='unavailable'&&historicalNsrs.length?`<tr><td><b>Historical NSR (`+historicalNsrs.length+`)</b></td><td>`+renderNsrCards(historicalNsrs,'')+`</td></tr>`:'')+`<tr><td><b>Configuration Accuracy History (`+escHtml(metric(historicalAccuracy.length,
  nceReady))+`)</b></td><td>`+(nceReady.level==='unavailable'?nceUnavailable:renderAccuracyCards(historicalAccuracy))+`</td></tr></tbody>`}
function renderMapping(){let popState=ready('population'),a=M.fcidAudit||{};loadTbl.innerHTML=h(['Source','Readiness','Expected Sheet','Actual Sheet','Score','Rows','Key Field','Unique Key'])+Object.keys(cfg).map(k=>{let isLoaded=loaded(k),status=dataMode===
  'demo'?'Demo':isLoaded?'Loaded':'Missing',uniqueCount=isLoaded?(k==='fcid'?(a.distinctCommitFcids||0):(M.idx[k]?.size||0)):'—';return tr([cfg[k].label,pill(status),cfg[k].sheet,meta[k]?.sheet||'-',meta[k]?.score!==undefined?meta[k].score+'/100':'-',isLoaded?raw[k].length:'—',
  cfg[k].key,uniqueCount])}).join('')+`</tbody>`;overlapTbl.innerHTML=h(['Diagnostic','Value','Note'])+tr(['Quarter source engine','MCSD -> Target Ship -> RD -> Calendar Quarter','MCSD cutoff: last Friday (inclusive)'])+tr(['Raw Commit Rows',metric(a.rawCommitRows||0,popState),
  'Before FCID validation and deduplication'])+tr(['Invalid FID Rows',metric(a.invalidFidRows||0,popState),'Blank and placeholder FIDs removed'])+tr(['Valid Commit Rows',metric(a.validCommitRows||0,popState),'Before FCID deduplication'])+tr(['Distinct Commit FCIDs',
  metric(a.distinctCommitFcids||0,popState),'One retained row per FCID'])+tr(['Duplicate FCID Groups',metric(a.duplicateFcidGroups||0,popState),'FCIDs with multiple Commit rows'])+tr(['Duplicate Rows Removed',metric(a.duplicateRowsRemoved||0,popState),
  'Not processed by downstream calculations'])+tr(['Fab Segment Preferred',metric(a.fabSegmentPreferred||0,popState),'Assigned Fab Segment replaced an earlier unassigned row'])+tr(['Excluded',metric(M.excluded.length,popState),'System Description exclusions'])+
  tr(['Effective Population',metric(M.rows.length,popState),'Distinct FCIDs after exclusions'])+`</tbody>`;mappingTbl.innerHTML=h(['Standard Field','FCID Database','CTL Report','CTL-NSR List','NSR Log','NCE / FA'])+[['Unified FCID','FID','Opportunity ID','FCID','Reference Fcid',
  'FCID'],['Quarter Engine','Calendar Quarter, Target Ship, RD','MCSD','MCSD','Not primary','Not primary'],['Pre Anchor','RD fallback','SRD at slot reference','Sales Ops RD reference','Not primary','Not primary'],['Post Anchor','Date: Mfg Commit Ship fallback','MCSD @ Current','MCSD',
  'Not primary','Not primary'],['Exclusion','System Description','Not primary','Not primary','Not primary','Not primary']].map(r=>tr(r)).join('')+`</tbody>`}
function applyFilters(){let m=filteredMetrics(),n=m.rows.reduce((s,r)=>s+(r.watchNsrCount||0),0),i=uniq(m.rows.filter(r=>r.watchNsrCount>0).map(r=>r.fid)).length;[['kPop',m.pop,'population'],['kPre',pct(m.preOk,m.pop),'core'],['kPost',pct(m.postOk,m.pop),'core'],['kRisk',m.riskFcid,
  'risk'],['kSlotReady',m.slotReady+' / '+m.pop+' ('+pct(m.slotReady,m.pop)+')','core'],['kNSRWatch',n,'nsr'],['kNSRImpact',i,'nsr'],['kSlotMcsd',m.slotMcsd,'core'],['kSlotNoMcsd',m.slotNoMcsd,'core'],['kNoSlotCtl',m.noSlotCtl,'core']].forEach(x=>setMetric(...x));
  renderFilterContext();renderFunnel();renderTables();renderNsrTable();renderHeat();renderDriverBars();renderDriverCards();renderFcid360()}
function tableRows(type){let rows=filteredRows(),risks=filteredRisks(),ids=new Set(rows.map(r=>r.fid));risks=risks.filter(r=>ids.has(r.fid));if(['population','stability','slot'].includes(type))return rows;if(type==='excluded')return filteredRows(M.excluded);if(type==='risk')return risks;if(type===
  'watch')return risks.filter(r=>r.priority==='P1');if(type==='fcid360')return rows.filter(r=>r.fid===clean(fcidInput.value));if(type==='quarter'){let out={};rows.forEach(r=>{let key=r.quarter+'|'+r.quarterSource,x=out[key]??={Quarter:r.quarter,QuarterSource:r.quarterSource,
  Population:0,PreOnTime:0,PostStable:0,SlotReady:0};x.Population++;if(r.preStatus==='On Time')x.PreOnTime++;if(r.postStatus==='Projected Stable')x.PostStable++;if(r.hasSlot===true)x.SlotReady++});return Object.values(out)}return rows}
function exportKey(type){if(['population','excluded'].includes(type))return 'population';if(type==='risk'||type==='fcid360')return 'risk';return 'core'}
function allowExport(key){let state=ready(key);if(state.level==='unavailable'){window.alert('Export blocked. '+readyText(state));return false}if((state.level==='partial'||state.level==='demo')&&!window.confirm((state.level===
  'partial'?'This export is provisional. ':'This export contains demo data. ')+readyText(state)+' Continue?'))return false;return true}
function exportCurrent(type){let map={population:'Filtered_Population',excluded:'Excluded_Population',stability:'Stability_Detail',slot:'Slot_Governance',risk:'Risk_Drivers',quarter:'Quarter_Governance',watch:'Executive_Watchlist',fcid360:'FCID360'},key=exportKey(type);
  if(!allowExport(key))return;exportWorkbook(map[type]+'_'+fileSuffix(),{[map[type]]:tableRows(type)})}
function exportAll(){if(!allowExport('risk'))return;let m=filteredMetrics();exportWorkbook('V15_3_4_Governance_Report_'+fileSuffix(),{Executive_KPI:[{Population:m.pop,PopulationStatus:ready('population').label,PreSlot:pct(m.preOk,m.pop),PostSlot:pct(m.postOk,m.pop),
  SlotReadiness:pct(m.slotReady,m.pop),CoreStatus:ready('core').label,RiskFCID:m.riskFcid,RiskStatus:ready('risk').label,OpenNSR:m.rows.reduce((s,r)=>s+(r.watchNsrCount||0),0),NSRStatus:ready('nsr').label}],Population:tableRows('population'),Excluded:tableRows('excluded'),
  Watchlist:tableRows('watch'),Stability:tableRows('stability'),Slot:tableRows('slot'),Risk:tableRows('risk'),Quarter:tableRows('quarter'),FCID360:tableRows('fcid360')})}
function showSectionFocus(t){clearTimeout(t._focusTimer);t.classList.remove('sectionFocusPulse');void t.offsetWidth;t.classList.add('sectionFocusPulse');t._focusTimer=setTimeout(()=>t.classList.remove('sectionFocusPulse'),5000)}
function jumpToSection(page,targetId){let btn=document.querySelector('.tab[data-page="'+page+'"]');if(!btn)return;tabTo(page,btn);requestAnimationFrame(()=>{let target=document.getElementById(targetId);if(!target)return;target.scrollIntoView({behavior:'smooth',block:'start'});
  showSectionFocus(target)})}
function handleKpiJumpKey(event,page,targetId){if(event.key==='Enter'||event.key===' '||event.key==='Spacebar'){event.preventDefault();jumpToSection(page,targetId)}}
function tabTo(n,btn){document.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));btn.classList.add('active');document.querySelectorAll('.page').forEach(p=>p.classList.add('hidden'));document.getElementById('page'+n).classList.remove('hidden')}
// END FCID REPORT ENGINE

  // Only the original demo/Chosen startup is replaced by the authorized readers.
  dataMode='live';loadAg();
  renderAg();
  function bindReferenceAction(key,eventName,handler){document.querySelectorAll('[data-reference-'+eventName+'="'+key+'"]').forEach(element=>element.addEventListener(eventName,function(event){if(handler.call(this,event)===false)event.preventDefault()}));}
  bindReferenceAction('0','click',function(event){tabTo('Executive',this)}); bindReferenceAction('1','click',function(event){tabTo('Stability',this)}); bindReferenceAction('2','click',function(event){tabTo('Slot',this)});
  bindReferenceAction('3','click',function(event){tabTo('Drivers',this)}); bindReferenceAction('4','click',function(event){tabTo('NSR',this)}); bindReferenceAction('5','click',function(event){tabTo('Quarter',this)});
  bindReferenceAction('6','click',function(event){tabTo('FCID',this)}); bindReferenceAction('7','click',function(event){tabTo('Mapping',this)}); bindReferenceAction('8','click',function(event){tabTo('UAT',this)}); bindReferenceAction('9','click',function(event){tabTo('Rules',this)});
  bindReferenceAction('10','click',function(event){exportCurrent('population')}); bindReferenceAction('11','click',function(event){exportCurrent('excluded')}); bindReferenceAction('12','click',function(event){exportCurrent('stability')});
  bindReferenceAction('13','click',function(event){exportCurrent('slot')}); bindReferenceAction('14','click',function(event){exportCurrent('risk')}); bindReferenceAction('15','click',function(event){exportCurrent('quarter')});
  bindReferenceAction('16','click',function(event){exportAll()}); bindReferenceAction('17','click',function(event){jumpToSection('Drivers','riskDriverDetail')}); bindReferenceAction('18','keydown',function(event){handleKpiJumpKey(event,'Drivers','riskDriverDetail')});
  bindReferenceAction('19','click',function(event){jumpToSection('NSR','openNsrPopulation')}); bindReferenceAction('20','keydown',function(event){handleKpiJumpKey(event,'NSR','openNsrPopulation')}); bindReferenceAction('21','click',function(event){renderFcid360()});
  bindReferenceAction('22','click',function(event){exportCurrent('fcid360')}); bindReferenceAction('23','click',function(event){saveAg(this)});
  bindReferenceAction('26','click',function(event){restoreAg()}); bindReferenceAction('28','click',function(event){jumpToSection('Rules','fabAccountMapping')}); let reader=null,readSequence=0;const pendingReads=new Map();
  function readWorkbook(bytes,key){
    if(!reader){
      // The worker calls the reference's own header scoring and canonicalization.
      const functions=[clean,norm,uniq,aliasesFor,exactAliasKey,scoreSourceHeader,betterHeader,bestHeaderFromMatrix,sourceScoreError,validateSourceCandidate,detectWorkbookSource,canonicalizeSourceRows];
      const code=`const cfg=${JSON.stringify(cfg)},SP=${JSON.stringify(SP)};`+functions.map(fn=>fn.toString()).join('\n')+
        `;self.onmessage=event=>{const {id,bytes,key}=event.data;try{const wb=XLSX.read(bytes,{type:'array'}),candidate=detectWorkbookSource(wb,key);`+
        `const rows=canonicalizeSourceRows(XLSX.utils.sheet_to_json(wb.Sheets[candidate.sheet],{defval:'',range:candidate.headerRow}),key);`+
        `if(!rows.length)throw sourceScoreError(key,{...candidate,score:0,missing:[SP[key].k]});`+
        `self.postMessage({id,ok:true,rows,sheet:candidate.sheet,headerRow:candidate.headerRow+1,score:candidate.score,fields:Object.keys(rows[0])})}`+
        `catch(error){self.postMessage({id,ok:false,message:error.message,code:error.code,userMessage:error.userMessage})}}`;
      reader=createWorkbookWorker(code);
      reader.onmessage=event=>{const result=event.data,waiting=pendingReads.get(result.id);if(!waiting)return;pendingReads.delete(result.id);if(result.ok){const {id,ok,...parsed}=result;waiting.resolve(parsed)}else waiting.reject(Object.assign(new Error(result.message),
        {code:result.code,userMessage:result.userMessage}))};
      reader.onerror=event=>{pendingReads.forEach(waiting=>waiting.reject(new Error(event.message||'Workbook reader failed.')));pendingReads.clear();reader.terminate();reader=null};
    }
    return new Promise((resolve,reject)=>{const id=++readSequence;pendingReads.set(id,{resolve,reject});reader.postMessage({id,bytes,key},[bytes])});
  }
  async function readFile(file,key){
    const ext=file.name.split('.').pop().toLowerCase();
    if(ext!=='csv'){if(!['xlsx','xls','xlsm'].includes(ext))throw new Error('Use an Excel or CSV file.');return readWorkbook(await file.arrayBuffer(),key)}
    const matrix=csvToMatrix(await file.text()),candidate=validateSourceCandidate(key,{...bestHeaderFromMatrix(matrix,key,0),sheet:'CSV',dataRows:matrix.length}),rows=canonicalizeSourceRows(rowsFromMatrix(matrix,candidate.headerRow),key);
    if(!rows.length)throw sourceScoreError(key,{...candidate,score:0,missing:[SP[key].k]});
    return {rows,sheet:'CSV',headerRow:candidate.headerRow+1,score:candidate.score,fields:Object.keys(rows[0])};
  }
  return {
    setSources(rows,metadata){Object.keys(raw).forEach(key=>raw[key]=rows[key]);meta=Object.fromEntries(Object.entries(metadata).map(([key,value])=>[key,{...value,mode:value.kind||'live'}]));dataMode='live';recalculate()}, applyFilters,resetFilters,selectAll,clearSelection:clearSel,
      syncChosen:chSync,registerChosen:name=>chOn.add(name), canonicalize:canonicalizeSourceRows,readWorkbook,readFile, snapshot:()=>({raw,meta,M,metrics:filteredMetrics(),ready:readyRules,accounts:AG}),tableRows
  };
}
}

// Shared SAP search and IndexedDB storage for the NSR and FCID dashboards.
function createNsrDataService(onProgress=()=>{}){
  const clean=value=>{const s=String(value??'').replace(/^\uFEFF/,'').trim();return /^(nan|null|undefined)$/i.test(s)?'':s}; const pageFetch=(url,options)=>unsafeWindow.fetch(url,options); const workflowDate=value=>clean(value).match(/^(\d{4}-\d{2}-\d{2})/)?.[1]||'';
  const xmlFirst=(node,names)=>names.map(name=>clean(node.getElementsByTagNameNS('*',name)[0]?.textContent)).find(Boolean)||'';
const SOURCE_FIELDS=['Temp Id#','NSR#','NSR Status','REF FCID','NSR Request','Task Name','NSR Title','NSR Type','Impacted Sub System','Initiator','Sales Rep','Sales Ops','Technical Contact','Customer','Submit Date','NSR Org','Primary Product'];
const SAP_SEARCH='/sap/opu/odata/sap/ZNSR_CD_SRV/SearchNsrSet';
const SAP_FIELDS={
  'Temp Id#':['TempId'],'NSR#':['Nsrnum'],'NSR Status':['NsrWorkflowStatus','NsrStatus'],'REF FCID':['RefFcid'], 'NSR Request':['NsrRequest'],'Task Name':['DescriptionTask'],'NSR Title':['NsrTitle','EnggTitle'],'NSR Type':['NsrType'], 'Impacted Sub System':['ImpSubSys'],
    'Initiator':['InitiatedByFn','Requester','InitiatedBy'],'Sales Rep':['SalesRepFn'], 'Sales Ops':['SalesOpsFn','SalesOps'],'Technical Contact':['TechContactFn'],'Customer':['Customer'],'Submit Date':['SubmitDate'], 'NSR Org':['NsrOrg'],'Primary Product':['PrimaryProd']
};
const NSR_CACHE_DB='toolkitKyra',NSR_CACHE_DB_VERSION=1,NSR_CACHE_STORE='datasets',NSR_CACHE_KEY='sap-nsr-search',NSR_CACHE_SCHEMA=1;

function openNsrCache(){
  return new Promise((resolve,reject)=>{
    const api=unsafeWindow.indexedDB;if(!api){reject(new Error('IndexedDB is unavailable.'));return}
    const request=api.open(NSR_CACHE_DB,NSR_CACHE_DB_VERSION);let settled=false; request.onupgradeneeded=()=>{const db=request.result;if(!db.objectStoreNames.contains(NSR_CACHE_STORE))db.createObjectStore(NSR_CACHE_STORE,{keyPath:'key'})};
    request.onsuccess=()=>{if(settled){request.result.close();return}settled=true;resolve(request.result)}; request.onerror=()=>{if(!settled){settled=true;reject(request.error||new Error('Could not open the NSR cache.'))}};
    request.onblocked=()=>{if(!settled){settled=true;reject(new Error('The NSR cache is blocked by another page.'))}};
  });
}

function cachedRequest(request){return new Promise((resolve,reject)=>{request.onsuccess=()=>resolve(request.result); request.onerror=()=>reject(request.error||new Error('The NSR cache request failed.'));});}
function cacheTransactionDone(transaction){return new Promise((resolve,reject)=>{transaction.oncomplete=()=>resolve(); transaction.onerror=()=>reject(transaction.error||new Error('The NSR cache transaction failed.')); transaction.onabort=()=>reject(transaction.error||
  new Error('The NSR cache transaction was cancelled.'));});}

async function readNsrCache(){
  const db=await openNsrCache();
  try{
    const entry=await cachedRequest(db.transaction(NSR_CACHE_STORE,'readonly').objectStore(NSR_CACHE_STORE).get(NSR_CACHE_KEY));
    if(!entry||entry.schema!==NSR_CACHE_SCHEMA||!Array.isArray(entry.rows)||!entry.rows.length||!validSapDate(entry.from)||!validSapDate(entry.to))return null;
    return entry;
  }finally{db.close()}
}

async function writeNsrCache(rows,range,pages){
  const db=await openNsrCache(),savedAt=Date.now();
  try{
    const transaction=db.transaction(NSR_CACHE_STORE,'readwrite'); transaction.objectStore(NSR_CACHE_STORE).put({key:NSR_CACHE_KEY,schema:NSR_CACHE_SCHEMA,savedAt:savedAt,from:range.from,to:range.to,pages:Number(pages)||0,rows:rows}); await cacheTransactionDone(transaction);
      return savedAt;
  }finally{db.close()}
}

function localDate(value){return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,'0')}-${String(value.getDate()).padStart(2,'0')}`;}
function sapRange(){const today=new Date(),yearAgo=new Date(today.getFullYear()-1,today.getMonth(),today.getDate()); const quarterStart=new Date(yearAgo.getFullYear(),Math.floor(yearAgo.getMonth()/3)*3,1); return {from:localDate(quarterStart),to:localDate(today)};}

function validSapDate(value){
  const match=clean(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!match)return false;
  const year=Number(match[1]),month=Number(match[2]),day=Number(match[3]),date=new Date(Date.UTC(year,month-1,day)); return date.getUTCFullYear()===year&&date.getUTCMonth()===month-1&&date.getUTCDate()===day;
}

function sapFilter(from,to){
  return (`(Nsrnum eq '' and Screen eq 'S' and Customer eq '' and Description eq '' and ImpSubSys eq '' and NsrRequest eq '' and NsrTitle eq '' and RefFcid eq '' and NsrType eq '' and SalesOpsFn eq '' and EnggTitle eq '' and ` +
  `NsrWorkflowStatus eq '' and NsrStatus eq '' and InitiatedByFn eq '' and HostFidUpgrade eq '' and (ResEngGroup eq '') and PrimaryProd eq '' and TechSpec eq '' and (SubmitDate ge datetime'${from}T00:00:00' and SubmitDate le datetime'${to}T23:59:59') and (NsrOrg eq ''))`);
}

function sapUrl(range){return `${SAP_SEARCH}?$filter=${encodeURIComponent(sapFilter(range.from,range.to))}`;}

function parseSearchXml(text){
  const doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw new Error('SAP returned invalid XML.');
  const rows=[...doc.getElementsByTagNameNS('*','entry')].map(entry=>{const props=entry.getElementsByTagNameNS('*','properties')[0]||entry,row={}; SOURCE_FIELDS.forEach(field=>row[field]=xmlFirst(props,SAP_FIELDS[field]||[field])); row['Submit Date']=workflowDate(row['Submit Date']);
    return row;}).filter(row=>row['NSR#']);
  const next=[...doc.getElementsByTagNameNS('*','link')].find(link=>clean(link.getAttribute('rel')).toLowerCase()==='next'); return {rows,next:next?clean(next.getAttribute('href')):''};
}

function nextSapUrl(href,current){const url=new URL(href,current); if(url.origin!==unsafeWindow.location.origin)throw new Error('SAP returned an unexpected pagination URL.'); return url.href;}

async function fetchSapRows(range,signal){
  const rows=[],seen=new Set();let url=sapUrl(range),pages=0;
  while(url){
    if(seen.has(url))throw new Error('SAP returned a repeated pagination URL.');
    if(pages>=250)throw new Error('SAP returned too many result pages.');
    seen.add(url);onProgress(range,rows.length); let response;
    try{response=await pageFetch(url,{credentials:'same-origin',cache:'no-store',signal,headers:{Accept:'application/atom+xml,application/xml,text/xml'}})} catch(_){throw new Error('SAP data request failed. Check the Fiori session and network connection.')}
    const text=await response.text(),type=(response.headers.get('Content-Type')||'').toLowerCase();
    if(response.status===401||response.status===403||type.includes('text/html')||/^\s*<!doctype\s+html/i.test(text)||/^\s*<html/i.test(text))throw new Error('Session expired. Return to NSR page and sign in again.');
    if(!response.ok)throw new Error(`SAP SearchNsrSet returned HTTP ${response.status}.`);
    const parsed=parseSearchXml(text);rows.push(...parsed.rows);pages++; url=parsed.next?nextSapUrl(parsed.next,response.url||url):'';
  }
  return {rows,pages};
}

async function readDataset(key){const db=await openNsrCache(); try{return await cachedRequest(db.transaction(NSR_CACHE_STORE,'readonly').objectStore(NSR_CACHE_STORE).get(key))}finally{db.close()}}
async function writeDataset(entry){const db=await openNsrCache(); try{const transaction=db.transaction(NSR_CACHE_STORE,'readwrite');transaction.objectStore(NSR_CACHE_STORE).put(entry); await cacheTransactionDone(transaction);}finally{db.close()}}
async function deleteDatasets(keys){const db=await openNsrCache(); try{const transaction=db.transaction(NSR_CACHE_STORE,'readwrite');keys.forEach(key=>transaction.objectStore(NSR_CACHE_STORE).delete(key)); await cacheTransactionDone(transaction);}finally{db.close()}}
return {fields:SOURCE_FIELDS,readNsrCache,writeNsrCache,localDate,sapRange,validSapDate,fetchSapRows,readDataset,writeDataset,deleteDatasets};
}

function runDashboard(){
'use strict';

const DETAIL_FIELDS=['NSR#','NSR Title','Task Status','NSR Type','Impacted Sub System','Initiator','REF FCID','NSR Org','Primary Product','Customer','Sales Rep','Sales Ops','Technical Contact','Submit Date','NSR Request','Temp Id#'];
const DETAIL_WIDTHS=[76,190,144,82,102,94,78,64,120,150,90,90,108,84,96,90]; const LIVE_FIELDS=['Submit Date','Task Status','ECD','Approver Team','Assigned Approver','Pending Days','Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'];
const CYCLE_DETAIL_FIELDS=DETAIL_FIELDS.filter(field=>field!=='NSR#'&&field!=='Initiator'&&!LIVE_FIELDS.includes(field));
const CYCLE_FIELDS=['NSR#',...LIVE_FIELDS,...CYCLE_DETAIL_FIELDS.flatMap(field=>field==='NSR Type'?[field,'NSR Category']:field==='Customer'?[field,'Initiator']:[field])];
const WF_FIELDS=['Task Name','Approver Status','Approver Team','Approver','Pending Days','Approved Date','Comments']; const WORKFLOW_EXPORT_FIELDS=['NSR#','Workflow Step',...WF_FIELDS];
const CYCLE_WIDTHS=CYCLE_FIELDS.map(field=>({
  'NSR#':78,'Submit Date':116,'Approver Team':125, 'Assigned Approver':130,'Pending Days':82,'Elapsed Days':82,'Cycle Time (ECD)':110,'Cycle Time (SOBi)':116,'NSR Title':190,'Task Status':144,'ECD':96, 'NSR Category':100,'Customer':150,'Primary Product':120,'Technical Contact':108,
    'Impacted Sub System':104
}[field]||92));
const CYCLE_CONCURRENCY=8,CYCLE_WARN_AT=100,CYCLE_LIMIT=500; const WORKFLOW_ODATA='/sap/opu/odata/sap/ZNSR_CD_SRV/WorkflowSet/'; const NSR_DETAIL_ODATA='/sap/opu/odata/sap/ZNSR_CD_SRV/GetNsrDetSet/'; const FIORI_DETAIL='/flp#ZNSR-display&/NsrDetail/';
const pageFetch=(url,options)=>unsafeWindow.fetch(url,options);
const {fields:SOURCE_FIELDS,readNsrCache,writeNsrCache,localDate,sapRange,validSapDate,fetchSapRows}=createNsrDataService((range,count)=>{setSourceStatus(`Loading SAP data…<br>${range.from} to ${range.to}${count?` · ${count.toLocaleString()} rows`:''}`,'loading');});
const FILTERS=[
  {id:'fInit',field:'Initiator',label:'Initiator',color:'#0f766e',filterOnly:true}, {id:'fStatus',field:'Task Status',label:'Task Status',color:'#2563eb'}, {id:'fDate',field:'Submit Period',label:'Submit Date',color:'#475569',filterOnly:true},
  {id:'fSys',field:'Impacted Sub System',label:'Impacted Sub System',color:'#0891b2'}, {id:'fType',field:'NSR Type',label:'NSR Type',color:'#7c3aed'}, {id:'fProd',field:'Primary Product',label:'Primary Product',color:'#d97706'},
  {id:'fOrg',field:'NSR Org',label:'NSR Org',color:'#16a34a'}, {id:'fCust',field:'Customer',label:'Customer',color:'#dc2626'}
];
const PIVOT_FILTERS=FILTERS.filter(f=>!f.filterOnly);
const ALIASES={
  'Temp Id#':['Temp Id#','Temp ID','TempId','Temporary ID'], 'NSR#':['NSR#','NSR','NSR No','NSR Number','NSR ID'], 'NSR Status':['NSR Status','Status'], 'REF FCID':['REF FCID','Reference Fcid','Reference FCID','Ref FCID','FCID'], 'NSR Request':['NSR Request','Request','Request Type'],
  'Task Name':['Task Name','Task','Workflow Task'], 'NSR Title':['NSR Title','Title'], 'NSR Type':['NSR Type','Type'], 'Impacted Sub System':['Impacted Sub System','Impacted Subsystem','Sub System','Subsystem'], 'Initiator':['Initiator','Initiated By','Requester'],
  'Sales Rep':['Sales Rep','Sales Representative'], 'Sales Ops':['Sales Ops','Sales Operation','Sales Operations'], 'Technical Contact':['Technical Contact','Tech Contact'], 'Customer':['Customer','Customer Name','Account'],
  'Submit Date':['Submit Date','Submitted Date','Submission Date'], 'NSR Org':['NSR Org','NSR Organization','Organization','Org'], 'Primary Product':['Primary Product','Product','PrimaryProduct']
};
const TASK_STEPS={'CREATE NSR':1,'REQUEST REVIEW':2,'SAFETY REVIEW':5,'ESTIMATES':6,'COST REVIEW':7, 'PRICE CONFIRMATION':8,'DEMAND CONFIRMATION':9,'FINAL APPROVAL':10,'VC VALIDATION':11, 'ENGINEERING COMMIT AND DESIGN':12};
const collator=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'});
const state={rows:[],filtered:[],source:null,stats:null,sort:{field:'Submit Date',dir:'desc'},chosenReady:false,toastTimer:null,sourceRun:0,sourceAbort:null,cycleRows:[],cycleSort:{field:'NSR#',dir:'asc'},cycleAbort:null,cycleRun:0}; let sapRangeTrigger=null;

const $id=id=>document.getElementById(id);
const clean=v=>{if(v===null||v===undefined)return ''; if(v instanceof Date&&!isNaN(v))return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`; const s=String(v).replace(/^\uFEFF/,'').trim();
  return /^(nan|null|undefined)$/i.test(s)?'':s;};
const upper=v=>clean(v).toUpperCase().replace(/\s+/g,' '); const normHeader=v=>clean(v).toLowerCase().replace(/&/g,'and').replace(/[^a-z0-9]+/g,''); const esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const displayValue=v=>clean(v)||'（Blank）'; const pct=(a,b)=>b?`${(a/b*100).toFixed(1)}%`:'0.0%';

function deriveTaskStatus(row){
  const task=upper(row['Task Name']),status=upper(row['NSR Status']);
  if(status.includes('CANCEL')||task==='CANCELLED')return '0-Cancelled';
  if(status.includes('REJECT')||task==='REJECTED')return '0-Rejected';
  if(status.includes('COMPLET')||task==='COMPLETED')return '15-Completed';
  if(task==='INITIAL ASSESSMENT')return status==='PENDING WITH PRODUCT MANAGER'?'4-Initial Assessment':'3-Initial Assessment';
  if(task==='SOBI CREATE AND APPROVAL')return status==='PENDING WITH MFE'?'14-SOBI Create and Approval':'13-SOBI Create and Approval';
  if(Object.prototype.hasOwnProperty.call(TASK_STEPS,task))return `${TASK_STEPS[task]}-${titleForTask(task)}`;
  const uniqueByStatus={'PENDING WITH INITIATOR':'1-Create NSR','PENDING WITH SAFETY':'5-Safety Review', 'PENDING WITH PGBD':'8-Price Confirmation','PENDING WITH POA PM':'11-VC Validation', 'PENDING WITH DESIGN ENGINEERING':'12-Engineering Commit and Design'};
  if(uniqueByStatus[status])return uniqueByStatus[status];
  return `Unmapped · ${clean(row['Task Name'])||clean(row['NSR Status'])||'Blank task/status'}`;
}

function titleForTask(task){
  const names={
    'CREATE NSR':'Create NSR','REQUEST REVIEW':'Request Review','SAFETY REVIEW':'Safety Review','ESTIMATES':'Estimates', 'COST REVIEW':'Cost Review','PRICE CONFIRMATION':'Price Confirmation','DEMAND CONFIRMATION':'Demand Confirmation', 'FINAL APPROVAL':'Final Approval',
      'VC VALIDATION':'VC Validation','ENGINEERING COMMIT AND DESIGN':'Engineering Commit and Design'
  };
  return names[task]||task;
}

function statusRank(v){const m=clean(v).match(/^(\d+)-/); if(!m)return [99,clean(v)]; return [Number(m[1]),clean(v)];}
function compareStatus(a,b){const aa=statusRank(a),bb=statusRank(b); return aa[0]-bb[0]||collator.compare(aa[1],bb[1]);}
function buildHeaderMap(headers){const byNorm=new Map(headers.map(h=>[normHeader(h),h])); const map={}; for(const field of SOURCE_FIELDS){const aliases=(ALIASES[field]||[field]).map(normHeader); map[field]=aliases.map(a=>byNorm.get(a)).find(Boolean)||'';} return map;}

function normalizeRows(rawRows){
  if(!rawRows.length)throw new Error('No data rows were found.');
  const headers=Object.keys(rawRows[0]||{}),headerMap=buildHeaderMap(headers);
  if(!headerMap['NSR#'])throw new Error('Required header “NSR#” was not found.');
  const map=new Map();let missingNsr=0,duplicates=0;
  rawRows.forEach(raw=>{
    const row={}; SOURCE_FIELDS.forEach(field=>row[field]=clean(raw[headerMap[field]])); const nsr=row['NSR#'];
    if(!nsr){missingNsr++;return}
    if(map.has(nsr)){duplicates++; const existing=map.get(nsr); SOURCE_FIELDS.forEach(field=>{if(!existing[field]&&row[field])existing[field]=row[field]});}else map.set(nsr,row);
  });
  const rows=[...map.values()].map(row=>({...row,'Task Status':deriveTaskStatus(row),'Submit Period':submitPeriod(row['Submit Date'])})); const unmapped=rows.filter(r=>r['Task Status'].startsWith('Unmapped')).length;
  return {rows,stats:{inputRows:rawRows.length,validRows:rows.length,missingNsr,duplicates,unmapped}};
}

function countRecognizedHeaders(values){const norms=new Set(values.map(normHeader).filter(Boolean)); let score=0; for(const field of SOURCE_FIELDS){if((ALIASES[field]||[field]).some(a=>norms.has(normHeader(a))))score++;} const hasNsr=(ALIASES['NSR#']||
  []).some(a=>norms.has(normHeader(a))); return {score,hasNsr};}

function findHeader(matrix){
  let best=null; matrix.slice(0,50).forEach((row,index)=>{const hit=countRecognizedHeaders(row||[]),candidate={index,score:hit.score,hasNsr:hit.hasNsr}; if(candidate.hasNsr&&(!best||candidate.score>best.score))best=candidate;});
  if(!best)throw new Error('Required header “NSR#” was not found in the first 50 rows.');
  return best;
}

function rowsFromMatrix(matrix,headerIndex){
  const rawHeaders=(matrix[headerIndex]||[]).map((v,i)=>clean(v)||`Column ${i+1}`),seen={}; const headers=rawHeaders.map(h=>{seen[h]=(seen[h]||0)+1;return seen[h]===1?h:`${h}_${seen[h]}`});
  return matrix.slice(headerIndex+1).filter(row=>(row||[]).some(v=>clean(v))).map(row=>{const out={};headers.forEach((h,i)=>out[h]=row[i]??'');return out;});
}

function delimiterScore(line,delimiter){let count=0,quoted=false; for(let i=0;i<line.length;i++){if(line[i]==='"')quoted=!quoted; else if(line[i]===delimiter&&!quoted)count++;} return count;}

function parseDelimited(text){
  text=String(text||'').replace(/^\uFEFF/,''); const probe=text.split(/\r?\n/).filter(Boolean).slice(0,8); const candidates=[',','\t',';','|']; const delimiter=candidates.map(d=>[d,probe.reduce((s,l)=>s+delimiterScore(l,d),0)]).sort((a,b)=>b[1]-a[1])[0][0]; const rows=[];let row=[],
    value='',quoted=false;
  for(let i=0;i<text.length;i++){
    const c=text[i],next=text[i+1];
    if(c==='"'&&quoted&&next==='"'){value+='"';i++}
    else if(c==='"')quoted=!quoted; else if(c===delimiter&&!quoted){row.push(value);value=''} else if((c==='\r'||c==='\n')&&!quoted){if(c==='\r'&&next==='\n')i++; row.push(value);value=''; if(row.some(v=>clean(v)))rows.push(row); row=[];}else value+=c;
  }
  if(row.length||value){row.push(value);if(row.some(v=>clean(v)))rows.push(row)}
  return rows;
}

function decodeText(buffer){const bytes=buffer instanceof Uint8Array?buffer:new Uint8Array(buffer); try{return new TextDecoder('utf-8',{fatal:true}).decode(bytes)}catch(_){try{return new TextDecoder('gb18030').decode(bytes)}
  catch(__){return new TextDecoder('windows-1252').decode(bytes)}}}

async function parseFile(file){
  const ext=(file.name.split('.').pop()||'').toLowerCase();
  if(ext==='csv'){const matrix=parseDelimited(decodeText(await file.arrayBuffer())),header=findHeader(matrix); return {rawRows:rowsFromMatrix(matrix,header.index),sheet:'CSV',headerRow:header.index+1,score:header.score};}
  if(!['xlsx','xls'].includes(ext))throw new Error('Please upload an .xlsx, .xls, or .csv file.');
  if(!window.XLSX)throw new Error('The XLSX parser did not load. Check the network connection and try again.');
  let wb;
  try{wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true})}
  catch(err){if(/encrypt|password|protected/i.test(String(err&&err.message||err)))throw new Error('Password-protected or encrypted Excel files are not supported. Save an unencrypted copy and upload it again.'); throw err;}
  let best=null;
  wb.SheetNames.forEach(sheet=>{
    const ws=wb.Sheets[sheet];if(!ws||!ws['!ref'])return;
    const range=XLSX.utils.decode_range(ws['!ref']);range.e.r=Math.min(range.e.r,49); const preview=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false,range:XLSX.utils.encode_range(range)});
    try{const hit=findHeader(preview),candidate={sheet,index:hit.index,score:hit.score}; if(!best||candidate.score>best.score)best=candidate;}catch(_){/* try next sheet */}
  });
  if(!best)throw new Error('No worksheet contains the required “NSR#” header.');
  const matrix=XLSX.utils.sheet_to_json(wb.Sheets[best.sheet],{header:1,defval:'',raw:false}); return {rawRows:rowsFromMatrix(matrix,best.index),sheet:best.sheet,headerRow:best.index+1,score:best.score};
}

function openSapRange(){
  const range=state.source&&state.source.isSap&&validSapDate(state.source.from)&&validSapDate(state.source.to)?{from:state.source.from,to:state.source.to}:sapRange(); const today=localDate(new Date()),from=$id('sapRangeFrom'),to=$id('sapRangeTo'); from.value=range.from;to.value=
    range.to;from.max=today;to.max=today;$id('sapRangeError').textContent=''; sapRangeTrigger=document.activeElement;$id('sapRangeModal').hidden=false;document.body.classList.add('sap-range-open');from.focus();
}

function closeSapRange(){if($id('sapRangeModal').hidden)return;$id('sapRangeModal').hidden=true;document.body.classList.remove('sap-range-open');$id('sapRangeError').textContent=''; if(sapRangeTrigger&&sapRangeTrigger.isConnected)sapRangeTrigger.focus();sapRangeTrigger=null;}

function selectedSapRange(){
  const from=clean($id('sapRangeFrom').value),to=clean($id('sapRangeTo').value),today=localDate(new Date());let error='';
  if(!validSapDate(from)||!validSapDate(to))error='Choose both From date and To date.'; else if(from>to)error='From date must be on or before To date.'; else if(to>today)error='To date cannot be after today.';
  $id('sapRangeError').textContent=error;return error?null:{from,to};
}

function setLoading(on){const line=$id('loadingLine');line.className='loading-line'+(on?' on':' done'); if(!on)setTimeout(()=>line.className='loading-line',350);}
function setSourceStatus(message,type='ok'){const el=$id('sourceStatus');el.className=`source-status ${type}`;el.innerHTML=message;}

function commitRaw(rawRows,meta){
  const normalized=normalizeRows(rawRows);state.rows=normalized.rows;state.stats=normalized.stats;state.source=meta; state.sort={field:'Submit Date',dir:'desc'};$id('globalSearch').value='';$id('tableSearch').value=''; buildFilterControls();applyFilters();
  if(meta.isSap){
    const savedAt=Number(meta.savedAt),savedLabel=Number.isFinite(savedAt)&&savedAt>0?new Date(savedAt).toLocaleString():'';
    if(meta.fromCache){$id('heroSource').textContent=`SAP cache · ${meta.from} to ${meta.to}`; setSourceStatus(`Cached SAP data loaded · ${esc(meta.from)} to ${esc(meta.to)}${savedLabel?`<br>Saved ${esc(savedLabel)}`:''}`,'ok');}else{$id('heroSource').textContent=
      `SAP live · ${meta.from} to ${meta.to}`; setSourceStatus(`SAP data loaded · ${esc(meta.from)} to ${esc(meta.to)}${meta.cacheSaved?' · cached locally':''}`,'ok');}
  }else{
    const s=state.stats,duplicates=s.duplicates+(meta.preprocessedDuplicates||0),details=[`${s.validRows.toLocaleString()} NSR#`,`${duplicates} duplicate${duplicates===1?'':'s'} consolidated`,`${s.unmapped} unmapped`]; $id('heroSource').textContent=meta.file;
    setSourceStatus(`<b>${esc(meta.file)}</b><br>${esc(meta.sheet)} · header row ${meta.headerRow} · ${details.join(' · ')}`,'ok');
  }
}

async function loadSapData(range=sapRange()){
  const run=++state.sourceRun,controller=new unsafeWindow.AbortController();
  if(state.sourceAbort)state.sourceAbort.abort();state.sourceAbort=controller; setLoading(true);setSourceStatus(`Loading SAP data…<br>${range.from} to ${range.to}`,'loading');
  try{
    $id('rawFile').value='';const result=await fetchSapRows(range,controller.signal);if(run!==state.sourceRun)return;
    let savedAt=0,cacheError='';
    try{savedAt=await writeNsrCache(result.rows,range,result.pages)}catch(error){cacheError=clean(error&&error.message)||'Local cache unavailable';console.warn('[Toolkit Kyra] Could not update the NSR cache.',error)}
    if(run!==state.sourceRun)return;
    commitRaw(result.rows,{file:'SAP live data',isSap:true,from:range.from,to:range.to,pages:result.pages,savedAt:savedAt,cacheSaved:Boolean(savedAt)});
    if(cacheError)setSourceStatus(`SAP data loaded · ${esc(range.from)} to ${esc(range.to)}<br>${esc(cacheError)}; this page will fetch SAP again next time.`,'ok');
    toast(`${state.rows.length.toLocaleString()} SAP NSR${state.rows.length===1?'':'s'} loaded.`);
  }catch(err){if(run!==state.sourceRun)return; setSourceStatus(`<b>SAP data load failed</b><br>${esc(err.message)} Upload a raw data file to continue.`,'error');toast(err.message,true);}finally{if(run===state.sourceRun){state.sourceAbort=null;setLoading(false)}}
}

async function loadInitialNsrData(){
  const run=++state.sourceRun;if(state.sourceAbort)state.sourceAbort.abort();state.sourceAbort=null; setLoading(true);setSourceStatus('Checking local NSR cache…','loading');
  try{
    const cached=await readNsrCache();if(run!==state.sourceRun)return;
    if(cached){
      $id('rawFile').value=''; commitRaw(cached.rows,{file:'SAP cached data',isSap:true,fromCache:true,from:cached.from,to:cached.to,pages:cached.pages,savedAt:cached.savedAt}); const today=localDate(new Date());
      if(cached.to!==today)await loadSapData({from:cached.from,to:today}); else setLoading(false);
      return;
    }
  }catch(error){console.warn('[Toolkit Kyra] Could not read the NSR cache; loading from SAP.',error)}
  if(run===state.sourceRun)await loadSapData(sapRange());
}

async function handleUpload(event){
  const file=event.target.files&&event.target.files[0];if(!file)return;
  const run=++state.sourceRun;if(state.sourceAbort)state.sourceAbort.abort();state.sourceAbort=null; setLoading(true);setSourceStatus(`Reading <b>${esc(file.name)}</b>…`,'loading');
  try{const parsed=await parseFile(file);if(run!==state.sourceRun)return; commitRaw(parsed.rawRows,{file:file.name,sheet:parsed.sheet,headerRow:parsed.headerRow,score:parsed.score});toast(`${file.name} loaded successfully.`);}catch(err){if(run!==state.sourceRun)return;
    setSourceStatus(`<b>Load failed: ${esc(file.name)}</b><br>${esc(err.message)}`,'error');toast(err.message,true);event.target.value='';}finally{if(run===state.sourceRun)setLoading(false)}
}

function buildFilterControls(){
  const host=$id('filterHost');
  if(!host.children.length){
    host.innerHTML=FILTERS.map(f=>`<div class="filter-group"><div class="filter-label"><label for="${f.id}">${esc(f.label)}</label><button type="button" data-clear-filter="${f.id}">Clear</button></div><select id="${f.id}" class="filter-select" multiple data-placeholder="All ${esc(f.label)}"></select>${f.field==='Task Status'?'<div class="quick-pick"><span>Quick select</span><button type="button" data-status="flow" aria-pressed="false">1–14 In workflow</button></div>':f.field==='Submit Period'?'<div id="dateYears" class="quick-pick"></div>':''}</div>`).join('');
    host.addEventListener('click',e=>{const year=e.target.dataset.year;if(year){toggleYear(year);return} const id=e.target.dataset.clearFilter;if(id)clearFilter(id);}); host.addEventListener('change',e=>{if(e.target.matches('.filter-select'))applyFilters()});
  }
  FILTERS.forEach(f=>{
    const select=$id(f.id);let values=[...new Set(state.rows.map(r=>r[f.field]))];
    if(f.field==='Submit Period'){values.sort((a,b)=>collator.compare(b,a)); const years=[...new Set(values.map(periodYear).filter(Boolean))]; renderYears(years);}else values.sort(f.field==='Task Status'?compareStatus:collator.compare);
    select.innerHTML=values.map(v=>`<option value="${esc(v)}">${esc(displayValue(v))}</option>`).join('');
  });
  initOrRefreshChosen();
}

function initOrRefreshChosen(){
  if(window.jQuery&&jQuery.fn&&jQuery.fn.chosen){
    FILTERS.forEach(f=>{const jq=jQuery(`#${f.id}`); if(jq.data('chosen'))jq.trigger('chosen:updated'); else jq.chosen({width:'100%',search_contains:true,hide_results_on_select:false,no_results_text:'No match'}); jq.off('change.nsrControlTower').on('change.nsrControlTower',
      applyFilters);});
    state.chosenReady=true;
  }
}

function renderYears(years){const host=$id('dateYears');if(!host)return; host.innerHTML=years.length?`<span>Select year</span>${years.map(year=>`<button type="button" data-year="${esc(year)}" aria-pressed="false">${esc(year)}</button>`).join('')}`:'';}
function toggleYear(year){const options=[...$id('fDate').options].filter(o=>periodYear(o.value)===year),selectAll=!options.length||options.some(o=>!o.selected); options.forEach(o=>o.selected=selectAll);syncSelect('fDate');applyFilters();}

function paintYears(){
  document.querySelectorAll('[data-year]').forEach(button=>{
    const options=[...$id('fDate').options].filter(o=>periodYear(o.value)===button.dataset.year),selected=options.filter(o=>o.selected).length; button.classList.toggle('active',options.length>0&&selected===options.length);button.classList.toggle('partial',selected>0&&selected<
      options.length);button.setAttribute('aria-pressed',selected===options.length?'true':'false');
  });
}

function statusSet(group){const single={done:'15-Completed',rejected:'0-Rejected',cancelled:'0-Cancelled'}[group]; return [...$id('fStatus').options].filter(o=>{if(group!=='flow')return o.value===single;const n=statusRank(o.value)[0];return n>=1&&n<=14}).map(o=>o.value);}
function toggleStatus(group){const select=$id('fStatus'),values=statusSet(group),selected=selectedValues('fStatus'); const same=values.length===selected.length&&values.every(v=>selected.includes(v)); [...select.options].forEach(o=>o.selected=!same&&values.includes(o.value));
  syncSelect('fStatus');applyFilters();}

function paintFlow(){
  const button=document.querySelector('#filterHost [data-status="flow"]');if(!button)return;
  const values=statusSet('flow'),selected=selectedValues('fStatus'),hits=selected.filter(v=>values.includes(v)).length; const active=values.length>0&&hits===values.length&&hits===selected.length; button.classList.toggle('active',active);button.classList.toggle('partial',hits>0&&
    !active);button.setAttribute('aria-pressed',active?'true':'false');
}

function selectedValues(id){return [...$id(id).selectedOptions].map(o=>o.value)}
function syncSelect(id){if(state.chosenReady)jQuery(`#${id}`).trigger('chosen:updated')}
function clearFilter(id,render=true){[...$id(id).options].forEach(o=>o.selected=false);syncSelect(id);if(render)applyFilters();}

function drillFilter(id,value,additive=false){
  const select=$id(id); [...select.options].forEach(o=>o.selected=o.value===value||(additive&&o.selected)); syncSelect(id);applyFilters(); const f=FILTERS.find(x=>x.id===id);toast(`${f.label}: ${displayValue(value)}`);
  $id('activeFilters').scrollIntoView({behavior:'smooth',block:'nearest'});
}

function filterRows(){
  const selections=Object.fromEntries(FILTERS.map(f=>[f.field,selectedValues(f.id)])); const query=clean($id('globalSearch').value).toLowerCase(); const tokens=query.split(/[\s,，;；、|/]+/).filter(Boolean);
  const queries=tokens.length>1&&tokens.every(token=>/^[a-z]?\d+(?:[-_][a-z0-9]+)*$/.test(token))?tokens:[query];
  return state.rows.filter(row=>{
    if(!FILTERS.every(f=>!selections[f.field].length||selections[f.field].includes(row[f.field])))return false;
    if(!query)return true;
    return queries.some(term=>DETAIL_FIELDS.some(field=>clean(row[field]).toLowerCase().includes(term))||clean(row['NSR Status']).toLowerCase().includes(term)||clean(row['Task Name']).toLowerCase().includes(term));
  });
}

function applyFilters(){state.filtered=filterRows(); paintYears();paintFlow();renderActiveFilters();renderKpis();renderPivots();renderTable();}

function renderActiveFilters(){
  const parts=[];
  FILTERS.forEach(f=>selectedValues(f.id).forEach(v=>parts.push(`<span class="filter-chip"><span>${esc(f.label)}: ${esc(displayValue(v))}</span><button type="button" aria-label="Remove filter" data-remove-filter="${f.id}" data-remove-value="${esc(v)}">×</button></span>`)));
  const q=clean($id('globalSearch').value);if(q)parts.push(`<span class="filter-chip"><span>Search: ${esc(q)}</span><button type="button" aria-label="Clear search" data-remove-search>×</button></span>`);
  $id('activeFilters').innerHTML=parts.join('')||'<span class="all-chip">All NSRs are selected.</span>';
}

function renderKpis(){
  const rows=state.filtered,total=rows.length;
  const counts={
    active:rows.filter(r=>{const n=statusRank(r['Task Status'])[0];return n>=1&&n<=14}).length, completed:rows.filter(r=>r['Task Status']==='15-Completed').length, rejected:rows.filter(r=>r['Task Status']==='0-Rejected').length,
    cancelled:rows.filter(r=>r['Task Status']==='0-Cancelled').length
  };
  const items=[
    ['Filtered NSR',total,'Current population','#0f766e'],['In workflow',counts.active,pct(counts.active,total),'#2563eb','flow'], ['Completed',counts.completed,pct(counts.completed,total),'#16a34a','done'],['Rejected',counts.rejected,pct(counts.rejected,total),'#dc2626','rejected'],
    ['Cancelled',counts.cancelled,pct(counts.cancelled,total),'#64748b','cancelled']
  ];
  $id('kpis').innerHTML=items.map(x=>`<div class="kpi" style="--kpi-color:${x[3]}"><small>${x[0]}</small>${x[4]?`<button class="kpi-num" type="button" data-status="${x[4]}" title="Filter Task Status">${Number(x[1]).toLocaleString()}</button>`:`<strong>${Number(x[1]).toLocaleString()}</strong>`}<span>${x[2]}</span></div>`).join('');
}

function renderPivots(){
  const rows=state.filtered,total=rows.length; $id('pivotContext').textContent=`${total.toLocaleString()} filtered NSR${total===1?'':'s'}`;
  $id('pivotGrid').innerHTML=PIVOT_FILTERS.map(f=>{
    const counts=new Map();rows.forEach(r=>counts.set(r[f.field],(counts.get(r[f.field])||0)+1)); let entries=[...counts.entries()]; entries.sort(f.field==='Task Status'?(a,b)=>compareStatus(a[0],b[0]):(a,b)=>b[1]-a[1]||collator.compare(a[0],b[0]));
    const max=Math.max(1,...entries.map(x=>x[1]));
    const list=entries.length?entries.map(([value,count])=>`<div class="pivot-row" title="${esc(displayValue(value))}"><span class="pivot-label">${esc(displayValue(value))}</span><span class="pivot-bar"><i style="width:${(count/ max*
      100).toFixed(1)}%;--bar:${f.color}"></i></span><button class="pivot-count" type="button" data-drill-id="${f.id}" data-drill-value="${esc(value)}" title="Filter by ${esc(displayValue(value))}">${count}</button></div>`).join(''):'<div class="empty">No matching NSRs</div>';
    return `<section class="pivot-card"><div class="pivot-head"><h3>${esc(f.label)}</h3><span>${counts.size} values · ${total.toLocaleString()} NSR</span></div><div class="pivot-list">${list}</div></section>`;
  }).join('');
}

function dateNumber(value){
  const s=clean(value);if(!s)return Number.NEGATIVE_INFINITY;
  if(/^\d{5}(?:\.\d+)?$/.test(s))return (Number(s)-25569)*86400000;
  let m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);if(m)return Date.UTC(Number(m[3]),Number(m[1])-1,Number(m[2]));
  m=s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);if(m)return Date.UTC(Number(m[1]),Number(m[2])-1,Number(m[3]));
  const n=Date.parse(s);return Number.isNaN(n)?Number.NEGATIVE_INFINITY:n;
}

function submitPeriod(value){const n=dateNumber(value);if(!Number.isFinite(n))return ''; const date=new Date(n);return `${date.getUTCFullYear()} Q${Math.floor(date.getUTCMonth()/3)+1}`;}
function periodYear(value){const match=clean(value).match(/^(\d{4})\s+Q[1-4]$/);return match?match[1]:'';}
function compareRows(a,b,field){if(field==='Task Status')return compareStatus(a[field],b[field]); if(field==='Submit Date')return dateNumber(a[field])-dateNumber(b[field]); return collator.compare(clean(a[field]),clean(b[field]));}
function currentTableRows(){const q=clean($id('tableSearch').value).toLowerCase(); const rows=q?state.filtered.filter(r=>DETAIL_FIELDS.some(f=>clean(r[f]).toLowerCase().includes(q))):[...state.filtered]; const {field,dir}=state.sort; return rows.sort((a,b)=>compareRows(a,b,
  field)*(dir==='asc'?1:-1));}
function statusPill(value){const cls=value==='15-Completed'?'terminal':value.startsWith('0-')?'stop':value.startsWith('Unmapped')?'unknown':''; return `<span class="status-pill ${cls}">${esc(value)}</span>`;}

function renderTable(){
  const rows=currentTableRows(),sort=state.sort; const cols=`<colgroup>${DETAIL_WIDTHS.map(width=>`<col style="width:${width}px">`).join('')}</colgroup>`;
  const head=`<thead><tr>${DETAIL_FIELDS.map(field=>`<th scope="col" title="${esc(field)}" data-sort="${esc(field)}" class="${sort.field===field?'sorted':''}" aria-sort="${sort.field===field?(sort.dir===
    'asc'?'ascending':'descending'):'none'}">${esc(field)}<span class="sort-mark">${sort.field===field?(sort.dir==='asc'?'▲':'▼'):'↕'}</span></th>`).join('')}</tr></thead>`;
  const body=rows.length?rows.map(row=>`<tr>${DETAIL_FIELDS.map(field=>`<td title="${esc(row[field])}">${field===
    'Task Status'?statusPill(row[field]):esc(row[field])}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="${DETAIL_FIELDS.length}"><div class="empty">No NSRs match the current filters.</div></td></tr>`;
  $id('detailTable').innerHTML=cols+head+`<tbody>${body}</tbody>`; $id('rowCount').textContent=`${rows.length.toLocaleString()} shown · ${state.filtered.length.toLocaleString()} filtered`;
}

function sortTable(field){if(state.sort.field===field)state.sort.dir=state.sort.dir==='asc'?'desc':'asc'; else state.sort={field,dir:field==='Submit Date'?'desc':'asc'}; renderTable();}
function tsvValue(v){return clean(v).replace(/[\t\r\n]+/g,' ')}
async function copyRows(){
  const rows=currentTableRows(),text=[DETAIL_FIELDS.join('\t'),...rows.map(r=>DETAIL_FIELDS.map(f=>tsvValue(r[f])).join('\t'))].join('\r\n');
  try{
    if(navigator.clipboard&&window.isSecureContext)await navigator.clipboard.writeText(text);
    else{const ta=document.createElement('textarea');ta.value=text;ta.style.cssText='position:fixed;left:-9999px;top:0';document.body.appendChild(ta);ta.select(); if(!document.execCommand('copy'))throw new Error('Copy was blocked by the browser.');ta.remove();}
    toast(`${rows.length.toLocaleString()} rows copied as tab-separated values.`);
  }catch(err){toast(err.message||'Copy failed.',true)}
}

function toast(message,isError=false){const el=$id('toast');clearTimeout(state.toastTimer);el.textContent=message;el.className=`toast show${isError?' error':''}`; state.toastTimer=setTimeout(()=>el.className='toast',3000);}

let wfTrigger=null,wfRows=[];

function switchTab(name){
  const cycle=name==='cycle'; $id('overviewTab').classList.toggle('active',!cycle);$id('cycleTab').classList.toggle('active',cycle); $id('overviewBtn').classList.toggle('active',!cycle);$id('cycleBtn').classList.toggle('active',cycle);
  $id('overviewBtn').setAttribute('aria-selected',cycle?'false':'true');$id('cycleBtn').setAttribute('aria-selected',cycle?'true':'false'); unsafeWindow.scrollTo({top:0,behavior:'smooth'});
}

function parseNsrInput(text){
  const valid=[],invalid=[],seen=new Set(); String(text||'').split(/[\s,;|]+/).map(v=>v.trim()).filter(Boolean).forEach(token=>{const nsr=token.toUpperCase(); if(!/^[A-Z]\d{5,9}$/.test(nsr)){invalid.push(token);return} if(!seen.has(nsr)){seen.add(nsr);valid.push(nsr)}});
  return {valid,invalid};
}

function xmlValue(node,name){const item=node&&node.getElementsByTagNameNS('*',name)[0];return clean(item&&item.textContent);}
function xmlFirst(node,names){for(const name of names){const value=xmlValue(node,name);if(value)return value}return '';}
function workflowDate(value){const match=clean(value).match(/^(\d{4}-\d{2}-\d{2})/);return match?match[1]:'';}
function workflowStatusHint(value){const status=upper(value);return status.includes('COMPLET')?'COMPLETED':status.includes('CANCEL')?'CANCELLED':status.includes('REJECT')?'REJECTED':'';}

function workflowHistory(entries){
  return entries.map(entry=>{
    const task=xmlFirst(entry,['DescriptionTask','TaskName']),approved=workflowDate(xmlFirst(entry,['ApprovalDate','ApprovedDate'])); let status=xmlFirst(entry,['ApproverStatus','Decision','Status']);if(!status&&task)status=approved?'APPROVED':'PENDING';
    return {
      'Task Name':task,'Approver Status':status,'Approver Team':xmlFirst(entry,['ApproverTeam','ApprovalTeam']), 'Approver':xmlFirst(entry,['Approver','AssignedApprover','ApproverName','NameText']), 'Pending Days':xmlFirst(entry,['PendingDays','PendingDay']),'Approved Date':approved,
      'Comments':xmlFirst(entry,['Comments','Comment','ApprovalComments','ApprovalComment','ApproverComments','DecisionComments','DecisionComment','CommentText','Remarks','Remark','Notes']).replace(/%20/gi,' ')
    };
  }).filter(stage=>WF_FIELDS.some(field=>clean(stage[field])));
}

function workflowStamp(value){const match=clean(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);return match?Date.UTC(Number(match[1]),Number(match[2])-1,Number(match[3])):NaN;}
function elapsedDays(startValue,endValue=''){const start=workflowStamp(startValue),now=new Date(),end=endValue?workflowStamp(endValue):Date.UTC(now.getFullYear(),now.getMonth(),now.getDate()); return Number.isFinite(start)&&Number.isFinite(end)?Math.max(0,Math.floor((end-start)/
  86400000)):'';}

function parseWorkflowXml(text,statusHint=''){
  const doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw new Error('The workflow service returned invalid XML.');
  const entries=[...doc.getElementsByTagNameNS('*','entry')],hintedTerminal=workflowStatusHint(statusHint);
  if(!entries.length)return {'Submit Date':'','Approver Team':'','Assigned Approver':'','Pending Days':'','Elapsed Days':'','Cycle Time (ECD)':'','Cycle Time (SOBi)':'',__completed:hintedTerminal==='COMPLETED',__completedAt:'',__cycleEndAt:'',__terminal:hintedTerminal,__workflow:[],
    __error:'No workflow data'};
  const decisions=entries.map(entry=>upper(xmlValue(entry,'Decision'))),submittedIndex=decisions.indexOf('SUBMITTED'); const submittedEntry=submittedIndex>=0?entries[submittedIndex]:null; const submitted=workflowDate(xmlValue(submittedEntry,'ApprovalDate')); let lastAction=-1;
    const approvedDates=[],terminalDates={};
  entries.forEach((entry,index)=>{
    const decision=decisions[index],date=workflowDate(xmlValue(entry,'ApprovalDate'));
    if(decision&&decision!=='PENDING'&&decision!=='N/A')lastAction=index;
    if(['APPROVED','COMPLETED'].includes(decision)&&date)approvedDates.push(date);
    const terminalDecision=workflowStatusHint(decision);
    if(terminalDecision&&date&&(!terminalDates[terminalDecision]||workflowStamp(date)>workflowStamp(terminalDates[terminalDecision])))terminalDates[terminalDecision]=date;
  });
  const after=entries.slice(lastAction+1);
  const pending=after.find(entry=>{const decision=upper(xmlValue(entry,'Decision')),approved=workflowDate(xmlValue(entry,'ApprovalDate')); return (!decision||decision==='PENDING')&&!approved&&['DescriptionTask','ApproverTeam','Approver','PendingDays'].some(name=>xmlValue(entry,name));
    })||null;
  const lastDecision=lastAction>=0?decisions[lastAction]:'',xmlTerminal=workflowStatusHint(lastDecision),terminal=hintedTerminal||xmlTerminal; const completedAt=approvedDates.reduce((latest,date)=>!latest||workflowStamp(date)>workflowStamp(latest)?date:latest,'');
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
  const text=await response.text(),type=(response.headers.get('Content-Type')||'').toLowerCase();
  if(response.status===401||response.status===403||type.includes('text/html')||/^\s*<!doctype\s+html/i.test(text)||/^\s*<html/i.test(text)){const err=new Error('Session expired. Return to NSR page and sign in again.');err.auth=true;throw err;}
  if(!response.ok)throw new Error(`${label} returned HTTP ${response.status}.`);
  if(!/<(?:[A-Za-z0-9_-]+:)?feed\b/i.test(text)&&!/<(?:[A-Za-z0-9_-]+:)?entry\b/i.test(text))throw new Error(`${label} returned an unexpected response.`);
  return text;
}

async function fetchWorkflow(nsr,signal,statusHint){const text=await fetchOdataXml(odataLookupUrl(WORKFLOW_ODATA,'Nsrnum',nsr),signal,'Workflow service'); return parseWorkflowXml(text,statusHint);}

async function fetchNsrEnrichment(tempId,signal){
  if(!clean(tempId))return {category:'',ecd:''};
  const text=await fetchOdataXml(odataLookupUrl(NSR_DETAIL_ODATA,'TempId',tempId),signal,'NSR detail service'); const doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw new Error('NSR detail service returned invalid XML.');
  const entry=doc.getElementsByTagNameNS('*','entry')[0]; return {category:xmlValue(entry,'NsrCategory'),ecd:workflowDate(xmlValue(entry,'HwEngCommit'))};
}

function mergeCycle(nsr,live,details){
  const detail=details.get(upper(nsr))||{},row={'NSR#':nsr,...live};
  if(!row['Submit Date'])row['Submit Date']=detail['Submit Date']||'';
  DETAIL_FIELDS.forEach(field=>{if(field!=='NSR#'&&!LIVE_FIELDS.includes(field))row[field]=detail[field]||''}); row['Task Status']=detail['Task Status']||({COMPLETED:'15-Completed',CANCELLED:'0-Cancelled',REJECTED:'0-Rejected'}[row.__terminal]||''); row['ECD']=row.__ecd||'';
  row['Elapsed Days']=row.__terminal||row.__completed?'':elapsedDays(row['Submit Date']); row['Cycle Time (ECD)']=row.__ecd?elapsedDays(row['Submit Date'],row.__ecd):''; row['Cycle Time (SOBi)']=row.__completed&&row.__completedAt?elapsedDays(row['Submit Date'],row.__completedAt):'';
  return row;
}

function setCycleBusy(busy){$id('runCycleBtn').disabled=busy;$id('cancelCycleBtn').disabled=!busy;$id('cycleInput').disabled=busy; $id('cycleBadge').textContent=busy?'Querying':'Ready';}
function setCycleProgress(done,total,message){const progress=$id('cycleProgress');progress.max=Math.max(1,total);progress.value=done; $id('cycleStatus').textContent=message;$id('cycleBadge').textContent=total?`${done}/${total}`:'Ready';}

function cycleError(err,statusHint=''){
  const message=clean(err&&err.message)||'Lookup failed',terminal=workflowStatusHint(statusHint);
  return {'Submit Date':'','Approver Team':'','Assigned Approver':'','Pending Days':'','Elapsed Days':'','Cycle Time (ECD)':'','Cycle Time (SOBi)':'',__completed:terminal==='COMPLETED',__completedAt:'',__cycleEndAt:'',__terminal:terminal,__error:message};
}

async function runCycleLookup(nsrs){
  const parsed=parseNsrInput(Array.isArray(nsrs)?nsrs.join('\n'):$id('cycleInput').value);
  if(!parsed.valid.length){toast('Enter at least one valid NSR#.',true);return}
  if(parsed.invalid.length)toast(`${parsed.invalid.length} invalid value${parsed.invalid.length===1?' was':'s were'} ignored.`,true);
  const total=parsed.valid.length;
  if(total>CYCLE_LIMIT){const message=`A maximum of ${CYCLE_LIMIT} NSRs can be queried at once. Reduce the list and try again.`;unsafeWindow.alert(message);setCycleProgress(0,0,message);return}
  if(total>CYCLE_WARN_AT&&!unsafeWindow.confirm(`You are about to query ${total.toLocaleString()} NSRs from SAP. This may take some time. Continue?`)){setCycleProgress(0,0,`Query cancelled. ${total.toLocaleString()} NSRs were not sent to SAP.`);return;}
  if(state.cycleAbort)state.cycleAbort.abort();
  const run=++state.cycleRun,controller=new unsafeWindow.AbortController();state.cycleAbort=controller;state.cycleRows=[]; const details=new Map(state.rows.map(row=>[upper(row['NSR#']),row])),results=new Array(total),workers=Math.min(CYCLE_CONCURRENCY,total);
  let next=0,done=0,stoppedForAuth=false,lastPaint=0; setCycleBusy(true);setCycleProgress(0,total,`Starting ${total.toLocaleString()} live lookup${total===1?'':'s'} · up to ${workers} requests at once…`);renderCycleTable();
  async function worker(){
    while(!controller.signal.aborted&&run===state.cycleRun&&!stoppedForAuth){
      const index=next++;if(index>=total)return;
      const nsr=parsed.valid[index],detail=details.get(upper(nsr)),statusHint=workflowStatusHint(detail&&detail['Task Status']); let live;
      try{live=await fetchWorkflow(nsr,controller.signal,statusHint)} catch(err){if(controller.signal.aborted)return; live=cycleError(err,statusHint);if(err.auth){stoppedForAuth=true;controller.abort()}}
      let categoryError='';live['NSR Category']='';live.__ecd='';
      if(!stoppedForAuth&&clean(detail&&detail['Temp Id#'])){
        try{const enrichment=await fetchNsrEnrichment(detail['Temp Id#'],controller.signal);live['NSR Category']=enrichment.category;live.__ecd=enrichment.ecd}
        catch(err){if(controller.signal.aborted)return; categoryError=clean(err&&err.message)||'NSR Category unavailable';if(err.auth){stoppedForAuth=true;controller.abort()}}
      }
      if(categoryError)live.__categoryError=categoryError;
      if(run!==state.cycleRun)return;
      results[index]=mergeCycle(nsr,live,details);done++;state.cycleRows=results.filter(Boolean); const now=Date.now();if(now-lastPaint>200||done===total||stoppedForAuth){renderCycleTable();lastPaint=now}
      if(!stoppedForAuth)setCycleProgress(done,total,`Queried ${done.toLocaleString()} of ${total.toLocaleString()} · up to ${workers} requests at once`);
    }
  }
  await Promise.all(Array.from({length:workers},()=>worker()));
  if(run!==state.cycleRun)return;
  state.cycleRows=results.filter(Boolean); const cancelled=controller.signal.aborted&&!stoppedForAuth,failed=state.cycleRows.filter(row=>row.__error||row.__categoryError).length; state.cycleAbort=null;setCycleBusy(false);renderCycleTable();
  if(stoppedForAuth){setCycleProgress(done,total,'Stopped because the Fiori session is unavailable. Return to Fiori, sign in, and run the query again.');toast('Fiori session unavailable.',true)}
  else if(cancelled)setCycleProgress(done,total,`Cancelled after ${done.toLocaleString()} of ${total.toLocaleString()} lookups.`);
  else{setCycleProgress(done,total,`${done.toLocaleString()} live lookup${done===1?'':'s'} completed${failed?` · ${failed.toLocaleString()} unavailable`:''}.`);
    toast(`${done.toLocaleString()} cycle time row${done===1?'':'s'} loaded${failed?` · ${failed.toLocaleString()} unavailable`:''}.`,failed>0)}
}

function compareCycle(a,b,field){
  if(['Pending Days','Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'].includes(field))return (Number(a[field])||0)-(Number(b[field])||0);
  if(field==='ECD'||field.includes('Date'))return dateNumber(a[field])-dateNumber(b[field]);
  if(field==='Task Status')return compareStatus(a[field],b[field]);
  return collator.compare(clean(a[field]),clean(b[field]));
}

function cycleRows(){
  const query=clean($id('cycleSearch').value).toLowerCase(); const rows=query?state.cycleRows.filter(row=>CYCLE_FIELDS.some(field=>clean(row[field]).toLowerCase().includes(query))):[...state.cycleRows]; const {field,dir}=state.cycleSort;return rows.sort((a,b)=>compareCycle(a,b,
    field)*(dir==='asc'?1:-1));
}

function cycleCell(row,field){
  if(field==='Task Status')return statusPill(row[field]);
  if(field==='NSR Title'&&clean(row[field])&&clean(row['Temp Id#'])&&clean(row['NSR#'])){
    const href=`${unsafeWindow.location.origin}${FIORI_DETAIL}${encodeURIComponent(clean(row['Temp Id#']))}/${encodeURIComponent(clean(row['NSR#']))}/Search`;
    return `<a class="ct-link" data-nsr-detail="${esc(row['NSR#'])}" href="${esc(href)}" target="_blank" rel="noopener noreferrer" title="Open ${esc(row['NSR#'])} in Fiori">${esc(row[field])}</a>`;
  }
  if(['Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'].includes(field)&&clean(row[field])&&Array.isArray(row.__workflow)&&
    row.__workflow.length){return `<button class="ct-link" type="button" data-wf-nsr="${esc(row['NSR#'])}" title="View full workflow for ${esc(row['NSR#'])}">${esc(row[field])}</button>`;}
  return esc(row[field]);
}

function cycleCellClass(row,field){
  const elapsed=clean(row['Elapsed Days']),ecdCycle=clean(row['Cycle Time (ECD)']),sobiCycle=clean(row['Cycle Time (SOBi)']); const elapsedOverEcd=elapsed&&ecdCycle&&Number(elapsed)>Number(ecdCycle); const ecdOverSobi=ecdCycle&&sobiCycle&&Number(ecdCycle)>Number(sobiCycle);
  if(field==='Elapsed Days'&&elapsedOverEcd)return 'duration-over';
  if(field==='Cycle Time (ECD)'&&(elapsedOverEcd||ecdOverSobi))return 'duration-over';
  return '';
}

function openWorkflow(nsr,trigger){
  const row=state.cycleRows.find(item=>upper(item['NSR#'])===upper(nsr)),stages=row&&Array.isArray(row.__workflow)?row.__workflow:[];
  if(!stages.length){toast('Workflow detail is unavailable for this NSR.',true);return}
  const head=`<thead><tr>${WF_FIELDS.map(field=>`<th scope="col">${esc(field)}</th>`).join('')}</tr></thead>`; const body=`<tbody>${stages.map(stage=>`<tr>${WF_FIELDS.map(field=>`<td>${esc(stage[field])}</td>`).join('')}</tr>`).join('')}</tbody>`;
  $id('wfTitle').textContent=`${clean(nsr)} Workflow Detail`;$id('wfMeta').textContent=`${stages.length.toLocaleString()} workflow record${stages.length===1?'':'s'}`; $id('wfTable').innerHTML=head+body;wfRows=stages;wfTrigger=trigger||document.activeElement;$id('wfModal').hidden=
    false;document.body.classList.add('wf-open');$id('wfClose').focus();
}

function closeWorkflow(){if($id('wfModal').hidden)return;$id('wfModal').hidden=true;document.body.classList.remove('wf-open'); if(wfTrigger&&wfTrigger.isConnected)wfTrigger.focus();wfTrigger=null;wfRows=[];}

function renderCycleTable(){
  const rows=cycleRows(),sort=state.cycleSort; const cols=`<colgroup>${CYCLE_WIDTHS.map(width=>`<col style="width:${width}px">`).join('')}</colgroup>`;
  const head=`<thead><tr>${CYCLE_FIELDS.map(field=>`<th scope="col" title="${esc(field)}" data-cycle-sort="${esc(field)}" class="${sort.field===field?'sorted':''}" aria-sort="${sort.field===field?(sort.dir===
    'asc'?'ascending':'descending'):'none'}">${esc(field)}<span class="sort-mark">${sort.field===field?(sort.dir==='asc'?'▲':'▼'):'↕'}</span></th>`).join('')}</tr></thead>`;
  const body=rows.length?rows.map(row=>`<tr>${CYCLE_FIELDS.map(field=>`<td class="${cycleCellClass(row,field)}" title="${esc(row[field])}">${cycleCell(row,
    field)}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="${CYCLE_FIELDS.length}"><div class="empty">No cycle time results yet.</div></td></tr>`;
  $id('cycleTable').innerHTML=cols+head+`<tbody>${body}</tbody>`; const failed=state.cycleRows.filter(row=>row.__error||row.__categoryError).length;
  $id('cycleCount').textContent=`${rows.length.toLocaleString()} shown · ${state.cycleRows.length.toLocaleString()} queried${failed?` · ${failed.toLocaleString()} unavailable`:''}`; $id('copyCycleBtn').disabled=!rows.length;$id('exportCycleBtn').disabled=!rows.length;
}

function sortCycle(field){if(state.cycleSort.field===field)state.cycleSort.dir=state.cycleSort.dir==='asc'?'desc':'asc'; else state.cycleSort={field,dir:field==='ECD'||field.includes('Date')?'desc':'asc'}; renderCycleTable();}

async function writeClipboard(text){
  if(navigator.clipboard&&window.isSecureContext)return navigator.clipboard.writeText(text);
  const ta=document.createElement('textarea');ta.value=text;ta.style.cssText='position:fixed;left:-9999px;top:0';document.body.appendChild(ta);ta.select();
  try{if(!document.execCommand('copy'))throw new Error('Copy was blocked by the browser.')}finally{ta.remove()}
}

async function copyWorkflow(){
  if(!wfRows.length)return;
  const text=[WF_FIELDS.join('\t'),...wfRows.map(row=>WF_FIELDS.map(field=>tsvValue(row[field])).join('\t'))].join('\r\n');
  try{await writeClipboard(text);toast(`${wfRows.length.toLocaleString()} workflow record${wfRows.length===1?'':'s'} copied.`)}catch(err){toast(err.message||'Copy failed.',true)}
}

async function copyCycleRows(){
  const rows=cycleRows(),text=[CYCLE_FIELDS.join('\t'),...rows.map(row=>CYCLE_FIELDS.map(field=>tsvValue(row[field])).join('\t'))].join('\r\n');
  try{await writeClipboard(text);toast(`${rows.length.toLocaleString()} cycle time row${rows.length===1?'':'s'} copied.`)}catch(err){toast(err.message||'Copy failed.',true)}
}

function exportCycleRows(){
  const rows=cycleRows();if(!rows.length)return;
  if(!window.XLSX){toast('The XLSX library did not load.',true);return}
  const data=rows.map(row=>Object.fromEntries(CYCLE_FIELDS.map(field=>[field,row[field]??'']))),sheet=XLSX.utils.json_to_sheet(data,{header:CYCLE_FIELDS});
  const workflowData=rows.flatMap(row=>(Array.isArray(row.__workflow)?row.__workflow:[]).map((stage,index)=>Object.fromEntries(WORKFLOW_EXPORT_FIELDS.map(field=>[field,field==='NSR#'?(row['NSR#']??''):field==='Workflow Step'?index+1:(stage[field]??'')]))));
  const workflowSheet=XLSX.utils.json_to_sheet(workflowData,{header:WORKFLOW_EXPORT_FIELDS}); sheet['!cols']=CYCLE_WIDTHS.map(width=>({wch:Math.max(10,Math.round(width/7))}));sheet['!autofilter']={ref:sheet['!ref']};
  workflowSheet['!cols']=[12,14,32,18,22,24,14,14,50].map(wch=>({wch}));workflowSheet['!autofilter']={ref:workflowSheet['!ref']}; const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'NSR Cycle Time');XLSX.utils.book_append_sheet(book,workflowSheet,
    'Workflow Detail'); const now=new Date(),stamp=`${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`; XLSX.writeFile(book,`NSR_Cycle_Time_${stamp}.xlsx`,{compression:true});
    toast(`${rows.length.toLocaleString()} summary rows and ${workflowData.length.toLocaleString()} workflow records exported.`);
}

function sendDetailToCycle(){const nsrs=currentTableRows().map(row=>clean(row['NSR#'])).filter(Boolean);if(!nsrs.length){toast('No visible NSR Detail rows to query.',true);return} $id('cycleInput').value=nsrs.join('\n');switchTab('cycle');runCycleLookup(nsrs);}
function bindEvents(){
  const defaultRange=sapRange();$id('sapRangeDefault').textContent=`Default SAP range: ${defaultRange.from} to ${defaultRange.to}`; $id('rawFile').addEventListener('change',handleUpload); $id('sapRangeOpen').addEventListener('click',openSapRange);
  $id('sapRangeForm').addEventListener('submit',(event)=>{event.preventDefault();const range=selectedSapRange();if(!range)return; closeSapRange();loadSapData(range);}); $id('sapRangeModal').addEventListener('click',e=>{if(e.target.closest('[data-sap-range-close]'))closeSapRange()});
  $id('resetBtn').addEventListener('click',()=>{FILTERS.forEach(f=>clearFilter(f.id,false)); $id('globalSearch').value='';$id('tableSearch').value='';applyFilters();toast('All filters cleared.');}); $id('globalSearch').addEventListener('input',applyFilters);
  $id('globalSearch').addEventListener('paste',event=>{const input=event.currentTarget,text=event.clipboardData.getData('text/plain'),normalized=text.replace(/\s+/g,' '); if(text===normalized)return; event.preventDefault();input.setRangeText(normalized,input.selectionStart,
    input.selectionEnd,'end');applyFilters();});
  $id('tableSearch').addEventListener('input',renderTable); $id('copyBtn').addEventListener('click',copyRows); $id('sendCycleBtn').addEventListener('click',sendDetailToCycle); $id('runCycleBtn').addEventListener('click',()=>runCycleLookup());
  $id('cancelCycleBtn').addEventListener('click',()=>{if(!state.cycleAbort)return;state.cycleAbort.abort();$id('cycleStatus').textContent='Cancelling active requests…';}); $id('copyCycleBtn').addEventListener('click',copyCycleRows);
  $id('exportCycleBtn').addEventListener('click',exportCycleRows); $id('cycleSearch').addEventListener('input',renderCycleTable); $id('cycleTable').addEventListener('click',e=>{const trigger=e.target.closest('[data-wf-nsr]');if(trigger)openWorkflow(trigger.dataset.wfNsr,trigger)});
  $id('wfCopy').addEventListener('click',copyWorkflow); $id('wfModal').addEventListener('click',e=>{if(e.target.closest('[data-wf-close]'))closeWorkflow()});
  document.addEventListener('keydown',e=>{if(e.key!=='Escape')return;if(!$id('sapRangeModal').hidden){closeSapRange();return}if(!$id('wfModal').hidden)closeWorkflow()});
  document.querySelectorAll('[data-tab]').forEach(button=>button.addEventListener('click',()=>switchTab(button.dataset.tab)));
  document.addEventListener('click',e=>{
    const status=e.target.closest('[data-status]');if(status){toggleStatus(status.dataset.status);return}
    const drill=e.target.closest('[data-drill-id]');if(drill){drillFilter(drill.dataset.drillId,drill.dataset.drillValue,e.ctrlKey||e.metaKey);return}
    const cycleTh=e.target.closest('th[data-cycle-sort]');if(cycleTh){sortCycle(cycleTh.dataset.cycleSort);return}
    const th=e.target.closest('th[data-sort]');if(th){sortTable(th.dataset.sort);return}
    const remove=e.target.closest('[data-remove-filter]');if(remove){const select=$id(remove.dataset.removeFilter);[...select.options].forEach(o=>{if(o.value===remove.dataset.removeValue)o.selected=false});syncSelect(select.id);applyFilters();return;}
    if(e.target.closest('[data-remove-search]')||e.target.closest('[data-clear-search]')){$id('globalSearch').value='';applyFilters()}
  });
}

bindEvents();renderCycleTable();loadInitialNsrData();

}
})();

// ==================== MyLam · Filtered Document Summary ====================
(function(){
'use strict';
if(location.hostname!=='techinfo.mylam.com'||!location.pathname.startsWith('/portal/group/mylam/contenttyperesults'))return;
const BUTTON_ID='kyra-mylam-summary',BUTTON_LABEL='🔩 Find Part Changes Related to TM/PCN',SEARCH_LABEL='Find Tech Articles for BOM',DIALOG_ID='kyra-mylam-summary-dialog',SUMMARY_CONCURRENCY=10,SUMMARY_REQUEST_TIMEOUT=30000;
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

function drawPage(jq,table,length,page){
  return new Promise((resolve,reject)=>{
    const node=jq(table.table().node()),events='draw.dt.kyraSummary error.dt.kyraSummary dt-error.dt.kyraSummary'; const finish=error=>{clearTimeout(timer);node.off(events,onEvent);error?reject(error):resolve()};
    const onEvent=(event,settings,code,message)=>finish(event.type==='draw'?null:new Error(message||'Could not load MyLam results.')); const timer=setTimeout(()=>finish(new Error('Loading MyLam results timed out. Please retry.')),30000); node.on(events,onEvent);
    try{table.page.len(length).page(page).draw('page')}catch(error){finish(error)}
  });
}

async function collectDocuments(){
  // DataTables belongs to the page's jQuery, not the userscript's bundled jQuery.
  const jq=[unsafeWindow.jQuery,unsafeWindow.$].find(value=>value?.fn?.dataTable?.isDataTable('#searchResults'));
  if(!jq)throw new Error('MyLam search results are not ready. Please wait and retry.');
  const table=jq('#searchResults').DataTable(),original=table.page.info(),items=new Map();let read=0,missing=0;
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
    return [...items.values()];
  }finally{const current=table.page.info(); if(current.length!==original.length||current.page!==original.page)await drawPage(jq,table,original.length,original.page);}
}

function extractPartNumbers(html){
  const parse=text=>{const template=document.createElement('template');template.innerHTML=text;return template.content}; const root=parse(html).querySelector('.tech-document-article__container');
  if(!root?.querySelector('.tech-document-article__fields,.tech-document-article__main-info'))throw new Error('Article content not found (login, access restriction or unsupported page).');
  const fragments=[root],occurrences=[],numbers=new Set(),normalize=text=>text.replace(/[\u2010-\u2015\u2212]/g,'-').replace(/\u00ad/g,'').replace(/\u00a0/g,' ');
  // Rich-text tables are serialized as raw*Data strings; decode literals without executing scripts.
  for(const script of root.querySelectorAll('script')){
    const declarations=[...script.textContent.matchAll(/\b(?:const|let|var)\s+(raw\w*Data)\s*=/g)];
    for(const declaration of declarations){
      const rest=script.textContent.slice(declaration.index+declaration[0].length),literal=rest.match(/^\s*("(?:\\[\s\S]|[^"\\])*"|'(?:\\[\s\S]|[^'\\])*')\s*;/);
      if(!literal)throw new Error(`Could not read embedded article content: ${declaration[1]}.`);
      const decoded=literal[1].slice(1,-1).replace(/\\(u[\da-f]{4}|x[\da-f]{2}|\r?\n|[\s\S])/gi,(_,escape)=>{if(/^[ux][\da-f]+$/i.test(escape))return String.fromCharCode(parseInt(escape.slice(1),16)); return ({n:'\n',r:'\r',t:'\t',b:'\b',f:'\f',v:'\v','0':'\0','\n':'',
        '\r\n':''})[escape]??escape;});
      const fragment=parse(decoded);fragments.push(!fragment.querySelector('table')&&/&lt;table\b/i.test(decoded)?parse(fragment.textContent):fragment);
    }
  }
  const header=text=>text.length<100&&!/\d/.test(text)&&/\bpart\s*(?:numbers?\b|nos?\b\.?|#)|\bp\s*\/\s*n\b|\bp\.?n\.?s?\b/i.test(text); let tableIndex=0;
  for(const fragment of fragments)for(const table of fragment.querySelectorAll('table')){
    tableIndex++;const grid=[],columns=new Set();
    for(const [r,row] of [...table.rows].entries()){
      grid[r]??=[];let column=0;
      for(const cell of row.cells){while(grid[r][column])column++; const rowSpan=cell.rowSpan||table.rows.length-r; for(let y=r;y<Math.min(table.rows.length,r+rowSpan);y++)for(let x=column;x<column+cell.colSpan;x++){grid[y]??=[];grid[y][x]=cell} column+=cell.colSpan;}
    }
    for(const [r,row] of grid.entries()){
      const seen=new Set();
      for(const [c,cell] of row.entries()){
        if(!cell||seen.has(cell))continue;seen.add(cell); const clone=cell.cloneNode(true);clone.querySelectorAll('table,script,style').forEach(node=>node.remove()); clone.querySelectorAll('br').forEach(node=>node.replaceWith('\n'));
          clone.querySelectorAll('p,div,li').forEach(node=>node.append('\n')); const text=normalize(clone.textContent).trim();
        if(header(text)){
          for(let x=c;x<c+cell.colSpan;x++)columns.add(x);
          // Also support key/value rows such as "Part Number | 123-456789-00".
          if(row[c+cell.colSpan]&&!header(row[c+cell.colSpan].textContent)&&/\d/.test(row[c+cell.colSpan].textContent))columns.add(c+cell.colSpan);
          continue;
        }
        const candidates=new Set(text.match(/(?<![A-Z0-9._-])\d{2,4}-[A-Z0-9]{4,8}(?:-[A-Z0-9]{2,4})?(?![A-Z0-9._-])/gi)||[]);
        if(columns.has(c)&&cell.colSpan===1)for(const token of text.match(/\b[A-Z0-9]+(?:[-._][A-Z0-9]+)*\b/gi)||[]){if(token.length>=5&&/\d/.test(token)&&!/^\d{1,3}(?:\.\d+)?$/.test(token))candidates.add(token);}
        for(const value of candidates){const partNumber=value.toUpperCase();numbers.add(partNumber);occurrences.push({partNumber,table:tableIndex,row:r+1,column:c+1})}
      }
    }
  }
  return {partNumbers:[...numbers].sort(),occurrences};
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

async function runSummaryQueue(queue,read,stopped){let next=0;const worker=async()=>{while(!stopped()&&next<queue.length)await read(queue[next++])}; await Promise.all(Array.from({length:Math.min(SUMMARY_CONCURRENCY,queue.length)},worker));}
function normalizedBomPartNumber(value){return String(value??'').normalize('NFKC').replace(/[\u2010-\u2015\u2212]/g,'-').replace(/\u00ad/g,'').replace(/\s+/g,'').toUpperCase();}
function bomFileStem(fileName){let stem=String(fileName||'BOM').replace(/\.[^.]+$/,'').trim(),system=stem.search(/\s*-\s*SYSTEM\b/i); if(system>=0)stem=stem.slice(0,system).trim(); return (stem||'BOM').replace(/[<>:"/\\|?*]/g,'_').replace(/[. ]+$/g,'').slice(0,100)||'BOM';}

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

async function readBomFile(file){
  if(typeof XLSX==='undefined'||!XLSX.read)throw new Error('Excel reader is unavailable. Refresh the page and try again.');
  let workbook,buffer;
  try{buffer=await file.arrayBuffer();workbook=XLSX.read(buffer,{type:'array',cellDates:true,cellStyles:true})} catch(error){throw new Error(`${file.name}: could not read the workbook (${error.message}).`)}
  const sourceSheetName=workbook.SheetNames[0],source=workbook.Sheets[sourceSheetName];
  if(!source)throw new Error(`${file.name}: the workbook has no worksheet.`);
  const decoded=XLSX.utils.decode_range(source['!ref']||'A1'),range={s:{r:0,c:0},e:decoded.e};
  let rows=XLSX.utils.sheet_to_json(source,{header:1,range,raw:true,defval:'',blankrows:true}),display=(row,column)=>{const cell=source[XLSX.utils.encode_cell({r:row,c:column})]; return !cell?'':typeof cell.v==='number'&&Number.isInteger(cell.v)&&Math.abs(cell.v)>=
    1e15?String(cell.v):XLSX.utils.format_cell(cell)};
  let lastRow=decoded.e.r,lastColumn=decoded.e.c,headerRow=-1,materialColumn=-1;
  function findHeader(){findMaterial:for(let row=0;row<=Math.min(lastRow,199);row++)for(let column=0;column<=lastColumn;column++){if(String(display(row,column)).replace(/\s+/g,' ').trim().toUpperCase()==='MATERIAL'){headerRow=row;materialColumn=column;break findMaterial}}}
  findHeader();
  if(headerRow<0){const recovered=recoverMalformedBomRows(buffer); if(recovered){rows=recovered;lastRow=rows.length-1;lastColumn=Math.max(0,...rows.slice(0,200).map(row=>row?.length||0))-1;display=(row,column)=>rows[row]?.[column]??'';findHeader()}}
  if(headerRow<0)throw new Error(`${file.name}: no MATERIAL column was found in the first worksheet.`);
  const materials=new Map();
  for(let row=headerRow+1;row<=lastRow;row++){const shown=String(display(row,materialColumn)).replace(/[\u2010-\u2015\u2212]/g,'-').replace(/\u00ad/g,'').trim(),key=normalizedBomPartNumber(shown); if(key&&!materials.has(key))materials.set(key,shown);}
  return {file,fileName:file.name,sourceSheetName,rows,headerRow,materialColumn,materials, merges:(source['!merges']||[]).map(rangeItem=>XLSX.utils.encode_range(rangeItem)), columnWidths:(source['!cols']||[]).map(column=>Number(column.wch)||(Number(column.wpx)?Number(column.wpx)/
    7:0)), rowHeights:(source['!rows']||[]).map(row=>Number(row.hpt)||(Number(row.hpx)?Number(row.hpx)*0.75:0))};
}

function analyzeBomMatches(boms,summaryRows){
  const first=boms[0],second=boms[1]||null,firstOnly=new Set(),secondOnly=new Set(),scope=new Set();
  if(second){first.materials.forEach((_value,key)=>{if(!second.materials.has(key)){firstOnly.add(key);scope.add(key)}}); second.materials.forEach((_value,key)=>{if(!first.materials.has(key)){secondOnly.add(key);scope.add(key)}});}else first.materials.forEach((_value,
    key)=>scope.add(key));
  const documents=new Map(),matched=new Set();
  summaryRows.forEach(row=>{
    const partNumber=normalizedBomPartNumber(row.partNumber),inFirst=second?firstOnly.has(partNumber):scope.has(partNumber),inSecond=Boolean(second&&secondOnly.has(partNumber));
    if(!partNumber||(!inFirst&&!inSecond))return;
    matched.add(partNumber);const key=row.articleId||row.url||`${row.documentNo}\n${row.title}`; let document=documents.get(key);
    if(!document){document={order:Number(row.documentIndex)||0,title:row.title||'',url:row.url||'',documentNo:row.documentNo||'',releaseDate:row.releaseDate||'',first:new Set(),second:new Set()};documents.set(key,document)}
    if(inFirst)document.first.add(first.materials.get(partNumber)||row.partNumber);
    if(inSecond)document.second.add(second.materials.get(partNumber)||row.partNumber);
  });
  return {second,firstOnly,secondOnly,matched,documents:[...documents.values()].sort((left,right)=>left.order-right.order||left.title.localeCompare(right.title))};
}

function addBomSourceSheet(workbook,bom,name){
  const worksheet=workbook.addWorksheet(name),clean=value=>value===undefined||value===null||value===''?null:value instanceof Date?value:typeof value==='number'&&Number.isInteger(value)&&Math.abs(value)>=1e15?String(value):['string','number',
    'boolean'].includes(typeof value)?value:String(value);
  for(let offset=0;offset<bom.rows.length;offset+=3000)worksheet.addRows(bom.rows.slice(offset,offset+3000).map(row=>Array.from(row,clean)));
  bom.merges.forEach(range=>{try{worksheet.mergeCells(range)}catch(_){/* keep the remaining source table when a legacy merge is invalid */}}); const columnCount=Math.max(bom.materialColumn+1,...bom.rows.slice(0,500).map(row=>row.length));
  for(let column=0;column<columnCount;column++){const measured=Math.max(8,...bom.rows.slice(0,500).map(row=>String(row[column]??'').split(/\r?\n/).reduce((max,line)=>Math.max(max,line.length),0)+2)); worksheet.getColumn(column+1).width=Math.min(60,bom.columnWidths[column]||measured);}
  bom.rowHeights.forEach((height,index)=>{if(height)worksheet.getRow(index+1).height=height}); return worksheet;
}

function addTechArticlesSheet(workbook,summaryRows){
  const columns=[['Title',50],['URL',46],['Document No.',18],['Release Date',16],['Part Number',28]],worksheet=workbook.addWorksheet('Tech Articles',{views:[{state:'frozen',ySplit:4,showGridLines:false}]});worksheet.properties.defaultRowHeight=18;
  worksheet.getCell(1,1).value='Tech Articles';worksheet.getCell(1,1).font={name:'Arial',size:15,bold:true,color:{argb:'FF1F2937'}};worksheet.getCell(1,1).alignment={vertical:'middle'};worksheet.getRow(1).height=25;
  for(let column=1;column<=columns.length;column++)worksheet.getCell(1,column).border={bottom:{style:'thin',color:{argb:'FF94A3B8'}}};
  const documentCount=new Set(summaryRows.map(row=>row.articleId||row.url||`${row.documentNo}\n${row.title}`)).size,partNumberCount=new Set(summaryRows.map(row=>normalizedBomPartNumber(row.partNumber)).filter(Boolean)).size;
  [['Documents',documentCount],['Unique Part Numbers',partNumberCount]].forEach((item,index)=>{const start=index*2+1,label=worksheet.getCell(2,start),value=worksheet.getCell(2,start+1);label.value=item[0];label.font={name:'Arial',size:10,bold:true,color:{argb:'FF475569'}};
    value.value=item[1];value.font={name:'Arial',size:10,color:{argb:'FF475569'}}});worksheet.getCell(2,5).value=`Links: ${summaryRows.length}`;worksheet.getCell(2,5).font={name:'Arial',size:10,color:{argb:'FF475569'}};
  columns.forEach((column,index)=>{worksheet.getColumn(index+1).width=column[1];const cell=worksheet.getCell(4,index+1);cell.value=column[0];cell.font={name:'Arial',size:10,bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF526B78'}};
    cell.alignment={vertical:'middle',horizontal:'center',wrapText:true};cell.border={right:{style:'thin',color:{argb:'FFFFFFFF'}}}});worksheet.getRow(4).height=28; const edge={style:'thin',color:{argb:'FFD6E0EA'}};
  [...summaryRows].sort((left,right)=>(Number(left.documentIndex)||0)-(Number(right.documentIndex)||0)||String(left.partNumber).localeCompare(String(right.partNumber))).forEach(document=>{
    const values=[document.title,document.url,document.documentNo,document.releaseDate,document.partNumber],row=worksheet.addRow(values);
    row.eachCell({includeEmpty:true},cell=>{cell.font={name:'Arial',size:10,color:{argb:'FF1F2937'}};cell.alignment={vertical:'top',horizontal:'left',wrapText:true};cell.border={bottom:edge}});
    if(document.url){const linkCell=row.getCell(2);linkCell.value={text:document.url,hyperlink:document.url};linkCell.font={name:'Arial',size:10,color:{argb:'FF0369A1'},underline:true}}
    row.height=Math.min(75,Math.max(32,15*Math.max(2,...values.map(value=>String(value||'').split('\n').length))));
  });
  worksheet.autoFilter={from:{row:4,column:1},to:{row:Math.max(4,worksheet.rowCount),column:columns.length}};return worksheet;
}

async function createDownloadWorkbook(boms,analysis,summaryRows){
  if(typeof ExcelJS==='undefined'||!ExcelJS.Workbook)throw new Error('Excel workbook writer is unavailable. Refresh the page and try again.');
  const workbook=new ExcelJS.Workbook();workbook.creator='Toolkit Kyra';workbook.company='Lam Research';workbook.created=new Date();
  if(analysis){
    const two=Boolean(analysis.second),columns=[['Title',60],['URL',46],['Document No.',18],['Release Date',16],...boms.map(bom=>[bomFileStem(bom.fileName),30])]; const worksheet=workbook.addWorksheet('Part Number Analysis',{views:[{state:'frozen',ySplit:1,showGridLines:false}]});
    columns.forEach(([header,width],index)=>{worksheet.getColumn(index+1).width=width;const cell=worksheet.getCell(1,index+1);cell.value=header;cell.font={name:'Arial',size:10,bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF526B78'}};
      cell.alignment={vertical:'middle',horizontal:'center'};cell.border={right:{style:'thin',color:{argb:'FFFFFFFF'}}}});
    const edge={style:'thin',color:{argb:'FFD6E0EA'}};
    analysis.documents.forEach(memo=>{
      const parts=[[...memo.first].sort(),[...memo.second].sort()],count=Math.max(1,parts[0].length,parts[1].length),start=worksheet.rowCount+1;
      for(let index=0;index<count;index++){
        const metadata=index===0?[memo.title,memo.url,memo.documentNo,memo.releaseDate]:[null,null,null,null],values=[...metadata,parts[0][index]||null];if(two)values.push(parts[1][index]||null);
        const row=worksheet.addRow(values); row.eachCell({includeEmpty:true},(cell,column)=>{cell.font={name:'Arial',size:10,color:{argb:'FF1F2937'}};cell.alignment={vertical:'middle',horizontal:'left',wrapText:column<=2};cell.border={bottom:edge}});
        row.getCell(5).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEFF6FF'}};
        if(two)row.getCell(6).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFECFDF5'}};
      }
      if(count>1)for(let column=1;column<=4;column++)worksheet.mergeCells(start,column,start+count-1,column);
      if(memo.url){const cell=worksheet.getCell(start,2);cell.value={text:memo.url,hyperlink:memo.url};cell.font={name:'Arial',size:10,color:{argb:'FF0369A1'},underline:true}}
      parts.slice(0,boms.length).forEach((values,index)=>{for(const value of values)worksheet.getColumn(index+5).width=Math.max(worksheet.getColumn(index+5).width,value.length+2)});
    });
    worksheet.autoFilter={from:{row:1,column:1},to:{row:Math.max(1,worksheet.rowCount),column:columns.length}};
  }
  boms.forEach(bom=>{
    const base=bomFileStem(bom.fileName).replace(/[\[\]:*?/\\]/g,'_').replace(/^'+|'+$/g,'').slice(0,31)||'BOM';let name=base,index=2;
    while(name.toLowerCase()==='tech articles'||workbook.worksheets.some(sheet=>sheet.name.toLowerCase()===name.toLowerCase())){const suffix=` (${index++})`;name=base.slice(0,31-suffix.length)+suffix}
    addBomSourceSheet(workbook,bom,name);
  });
  addTechArticlesSheet(workbook,summaryRows||[]); return workbook.xlsx.writeBuffer();
}

function downloadBomWorkbook(buffer,fileName){
  const url=URL.createObjectURL(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})),anchor=document.createElement('a'); anchor.href=url;anchor.download=fileName;document.body.append(anchor);anchor.click();anchor.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
}

function showSummary(items){
  cancelSummary();document.getElementById(DIALOG_ID)?.remove(); const dialog=document.createElement('dialog'),rows=[],controllers=new Set();let running=false,scanComplete=false,stopped=false,message='',comparisonQueued=false,comparisonRunning=false,downloadRunning=false,
    latestBomResult=null; items.forEach((item,index)=>Object.assign(item,{status:'pending',partNumbers:[],occurrences:[],error:'',documentIndex:index})); unsafeWindow.myLamSummary=items;unsafeWindow.myLamPartNumbers=rows; dialog.id=DIALOG_ID;dialog.style.cssText=
    'width:96vw;max-width:1500px;height:88vh;padding:20px;border:1px solid #cbd5e1;border-radius:10px;background:#fff;color:#222;box-shadow:0 4px 24px #0005';
  dialog.innerHTML=(`<style>
    #${DIALOG_ID}{box-sizing:border-box;overflow:hidden}#${DIALOG_ID}::backdrop{background:#0006}#${DIALOG_ID} [hidden]{display:none!important}#${DIALOG_ID} button{cursor:pointer}#${DIALOG_ID} button:disabled{cursor:not-allowed;opacity:.55}
    #${DIALOG_ID} table{width:100%;border-collapse:collapse;font:13px Arial,sans-serif}#${DIALOG_ID} th,#${DIALOG_ID} td{padding:8px;text-align:left;border-bottom:1px solid #dce3e8;overflow-wrap:anywhere}#${DIALOG_ID} th{position:sticky;top:0;background:#eef5f7}#${DIALOG_ID} a{color:#007a95}
    #${DIALOG_ID} .kyra-bom-panel{flex:none;padding:11px;border:1px solid #cbd5e1;border-radius:8px;background:#f8fafc}#${DIALOG_ID} .kyra-bom-panel>summary{cursor:pointer;color:#334155;font:600 13px Arial,sans-serif}#${DIALOG_ID} .kyra-bom-inputs{display:grid;grid-template-columns:minmax(220px,1fr) minmax(220px,1fr) auto;gap:10px;align-items:end;margin-top:12px}
    #${DIALOG_ID} .kyra-bom-file{display:grid;gap:5px;color:#334155;font:600 12px Arial,sans-serif}#${DIALOG_ID} .kyra-bom-file input{min-width:0;font:12px Arial,sans-serif}
    #${DIALOG_ID} [data-action="compare"],#${DIALOG_ID} [data-action="download"]{min-height:32px;border:1px solid #526b78;border-radius:6px;background:#526b78;color:#fff;font-weight:700;padding:6px 12px;white-space:nowrap}#${DIALOG_ID} [data-action="download"]{border-color:#0f766e;background:#0f766e}#${DIALOG_ID} [data-bom-error]{color:#b91c1c;font:12px Arial,sans-serif}#${DIALOG_ID} [data-bom-error]:empty{display:none}
    #${DIALOG_ID} .kyra-comparison{display:flex;flex-direction:column;flex:2;min-height:0;border:1px solid #cbd5e1;border-radius:8px;background:#fff;overflow:hidden}#${DIALOG_ID} .kyra-comparison-heading{padding:10px 12px;background:#f1f5f9;color:#334155;font:700 14px Arial,sans-serif}#${DIALOG_ID} .kyra-comparison-content{display:flex;flex-direction:column;flex:1;min-height:0;gap:8px;padding:9px}#${DIALOG_ID} .kyra-comparison-metrics{display:flex;gap:7px;flex-wrap:wrap}#${DIALOG_ID} .kyra-comparison-metrics span{border-radius:999px;background:#e2e8f0;color:#334155;padding:4px 8px;font:11px Arial,sans-serif}#${DIALOG_ID} .kyra-comparison-table{flex:1;min-height:0;overflow:auto;border:1px solid #dce3e8;border-radius:6px}#${DIALOG_ID} .kyra-comparison-table table{min-width:1050px}#${DIALOG_ID} .kyra-comparison-table th{white-space:nowrap}#${DIALOG_ID} .kyra-comparison-table th:first-child{width:50%;min-width:350px}#${DIALOG_ID} .kyra-comparison-table th:nth-child(2){min-width:105px}#${DIALOG_ID} .kyra-comparison-table th:nth-child(3){min-width:120px}#${DIALOG_ID} .kyra-comparison-table th:nth-child(n+4){min-width:210px}#${DIALOG_ID} .kyra-comparison-table td{vertical-align:top}#${DIALOG_ID} .kyra-comparison-table td:not(:first-child){white-space:pre;overflow-wrap:normal;word-break:normal}#${DIALOG_ID} .kyra-bom-first{background:#eff6ff}#${DIALOG_ID} .kyra-bom-second{background:#ecfdf5}#${DIALOG_ID} [data-comparison-empty]{padding:10px;color:#64748b;font:12px Arial,sans-serif}
    #${DIALOG_ID} .kyra-documents{flex:none;border:1px solid #dce3e8;border-radius:6px;overflow:auto}#${DIALOG_ID} .kyra-documents[open]{flex:1;min-height:150px}#${DIALOG_ID} .kyra-documents>summary{position:sticky;top:0;z-index:1;padding:10px;background:#f8fafc;color:#334155;cursor:pointer;font:600 13px Arial,sans-serif}
    @media(max-width:1050px){#${DIALOG_ID} .kyra-bom-inputs{grid-template-columns:1fr 1fr}#${DIALOG_ID} [data-action="compare"],#${DIALOG_ID} [data-action="download"]{justify-self:start}}@media(max-width:700px){#${DIALOG_ID} .kyra-bom-inputs{grid-template-columns:1fr}}
    </style>
    <div style="display:flex;flex-direction:column;gap:12px;height:100%"><div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap"><strong style="flex:1;font-size:20px">BOM - Part Number Analysis</strong><button ` +
  `type="button" data-action="download" disabled>Download Excel</button><button type="button" data-action="close">Close</button></div>
    <details data-bom-panel class="kyra-bom-panel" open><summary>BOM files <span data-bom-names></span></summary><div class="kyra-bom-inputs"><label class="kyra-bom-file"><span data-bom-label>BOM1 - SYSTEM</span><input ` +
  `data-bom="1" type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv"></label><label class="kyra-bom-file"><span data-bom-label>BOM2 - SYSTEM</span><input data-bom="2" type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv"></label><button ` +
  `type="button" data-action="compare">${SEARCH_LABEL}</button></div></details><div data-bom-error role="alert"></div>
    <section data-comparison class="kyra-comparison" hidden><div class="kyra-comparison-heading">Comparison result <span data-comparison-summary></span></div><div class="kyra-comparison-content"><div data-comparison-metrics ` +
  `class="kyra-comparison-metrics"></div><div data-comparison-empty hidden></div><div data-comparison-table class="kyra-comparison-table" hidden><table><thead data-comparison-head></thead><tbody data-comparison-body></tbody>` +
  `</table></div></div></section>
    <details data-summary-list class="kyra-documents" open><summary>Tech Articles</summary><div data-progress role="status" style="padding:8px 10px;font-size:13px"></div><table><thead><tr><th>Part Number</th><th>Document ` +
  `No.</th><th>Title</th><th>Release Date</th></tr></thead><tbody data-summary-body></tbody></table></details></div>`);
  const button=name=>dialog.querySelector(`[data-action="${name}"]`),progress=dialog.querySelector('[data-progress]'),body=dialog.querySelector('[data-summary-body]'),bomInputs=[...dialog.querySelectorAll('[data-bom]')],bomPanel=dialog.querySelector('[data-bom-panel]'),
    summaryList=dialog.querySelector('[data-summary-list]'),bomError=dialog.querySelector('[data-bom-error]'),comparisonPanel=dialog.querySelector('[data-comparison]'),comparisonSummary=dialog.querySelector('[data-comparison-summary]'),
    comparisonMetrics=dialog.querySelector('[data-comparison-metrics]'),comparisonEmpty=dialog.querySelector('[data-comparison-empty]'),comparisonTable=dialog.querySelector('[data-comparison-table]'),comparisonHead=dialog.querySelector('[data-comparison-head]'),
    comparisonBody=dialog.querySelector('[data-comparison-body]');
  const link=(item,label)=>{const a=document.createElement('a');a.href=item.url;a.target='_blank';a.rel='noopener noreferrer';a.textContent=label||item.url;return a}; const selectedBomFiles=()=>bomInputs.map(input=>input.files?.[0]).filter(Boolean);
  const setBomError=text=>{bomError.textContent=text};
  function clearBomComparison(){
    latestBomResult=null;setBomError('');comparisonPanel.hidden=true;comparisonSummary.textContent='';comparisonMetrics.replaceChildren();comparisonHead.replaceChildren();comparisonBody.replaceChildren();comparisonEmpty.hidden=true;comparisonTable.hidden=true;bomPanel.open=true;
      summaryList.open=true;
  }
  function renderBomComparison(result){
    const {boms,analysis}=result,two=boms.length===2,names=boms.map(bom=>bomFileStem(bom.fileName)),headers=['Title','Document No.','Release Date',...names]; comparisonPanel.hidden=false;bomPanel.open=false;summaryList.open=false;comparisonSummary.textContent=
      `— ${analysis.documents.length} Tech Article${analysis.documents.length===1?'':'s'}`;comparisonMetrics.replaceChildren();
    const metrics=two?[`${names[0]} differences: ${analysis.firstOnly.size}`,`${names[1]} differences: ${analysis.secondOnly.size}`,`Matching Part Numbers: ${analysis.matched.size}`]:[`${names[0]} materials: ${boms[0].materials.size}`,
      `Matching Part Numbers: ${analysis.matched.size}`];
    metrics.forEach(text=>{const span=document.createElement('span');span.textContent=text;comparisonMetrics.append(span)}); const headerRow=document.createElement('tr');headers.forEach(text=>{const th=document.createElement('th');th.textContent=text;headerRow.append(th)});
      comparisonHead.replaceChildren(headerRow);comparisonBody.replaceChildren();
    analysis.documents.forEach(memo=>{
      const values=[memo.title,memo.documentNo,memo.releaseDate,[...memo.first].sort().join('\n')];if(two)values.push([...memo.second].sort().join('\n'));const tr=document.createElement('tr');
      values.forEach((value,index)=>{const td=document.createElement('td');if(index===0)td.append(link(memo,memo.title));else td.textContent=value;if(index===3)td.className='kyra-bom-first';if(index===4)td.className='kyra-bom-second';tr.append(td)});comparisonBody.append(tr);
    });
    comparisonEmpty.textContent=analysis.documents.length?'':'No Tech Articles matched the relevant BOM Part Numbers.';comparisonEmpty.hidden=Boolean(analysis.documents.length);comparisonTable.hidden=!analysis.documents.length;
  }
  function update(){
    const done=items.filter(item=>['done','empty','error'].includes(item.status)).length,failed=items.filter(item=>item.status==='error').length;
    progress.textContent=`${done}/${items.length} documents · ${new Set(rows.map(row=>row.partNumber)).size} unique part numbers · ${failed} failed${message?' — '+message:''}`; const files=selectedBomFiles(),compareButton=button('compare'),downloadButton=button('download');
      compareButton.disabled=comparisonRunning||comparisonQueued||downloadRunning||!files.length||(!running&&!scanComplete);compareButton.textContent=SEARCH_LABEL;
      downloadButton.disabled=downloadRunning||comparisonRunning||running||!scanComplete||Boolean(files.length&&!latestBomResult); downloadButton.textContent=downloadRunning?'Creating Excel...':'Download Excel';bomInputs.forEach(input=>input.disabled=comparisonRunning||
        downloadRunning);
  }
  function appendResults(item){
    const fragment=document.createDocumentFragment();
    for(const partNumber of item.partNumbers){
      const result={partNumber,documentNo:item.documentNo,title:item.title,url:item.url,releaseDate:item.releaseDate,articleId:item.articleId,documentIndex:item.documentIndex};rows.push(result); const tr=document.createElement('tr');for(const value of [partNumber,
        item.documentNo]){const td=document.createElement('td');td.textContent=value;tr.append(td)}
      const td=document.createElement('td');td.append(link(item,item.title));tr.append(td);const date=document.createElement('td');date.textContent=item.releaseDate;tr.append(date);fragment.append(tr);
    }
    body.append(fragment);
  }
  const stop=()=>{stopped=true;controllers.forEach(controller=>controller.abort())};cancelSummary=stop;
  async function startBomComparison(){
    const files=selectedBomFiles();if(!files.length)return;
    if(running){comparisonQueued=true;update();return}
    if(!scanComplete){setBomError('Document loading was interrupted. Close this window and open the analysis again to refresh the current filtered results.');update();return}
    comparisonQueued=false;comparisonRunning=true;clearBomComparison();update();
    try{const boms=[];for(const file of files)boms.push(await readBomFile(file)); const analysis=analyzeBomMatches(boms,rows),outputName=boms.map(bom=>bomFileStem(bom.fileName)).join(' vs ')+'.xlsx';latestBomResult={boms,analysis,outputName};renderBomComparison(latestBomResult);}
      catch(error){clearBomComparison();console.error('[MyLam BOM Analysis]',error);setBomError(error.message||String(error))}
    finally{comparisonRunning=false;update()}
  }
  async function downloadBomResult(){
    const files=selectedBomFiles(),result=files.length?latestBomResult:null;if(!scanComplete||(files.length&&!result))return;const outputName=result?result.outputName:'MyLam_Tech_Articles.xlsx';downloadRunning=true;setBomError('');update();
    try{downloadBomWorkbook(await createDownloadWorkbook(result?result.boms:[],result?result.analysis:null,rows),outputName)} catch(error){console.error('[MyLam BOM Analysis]',error);setBomError(error.message||String(error))} finally{downloadRunning=false;update()}
  }
  async function readOne(item){
    if(stopped)return;item.status='reading';item.error='';const controller=new unsafeWindow.AbortController();controllers.add(controller);let timedOut=false; const timer=setTimeout(()=>{timedOut=true;controller.abort()},SUMMARY_REQUEST_TIMEOUT);update();
    try{Object.assign(item,await readArticle(item,controller.signal));item.status=item.partNumbers.length?'done':'empty';appendResults(item)}
    catch(error){
      const paused=Boolean(error.pause);item.status=stopped&&!paused?'pending':'error';item.error=item.status==='error'?(timedOut?'Request timed out.':error.message):'';
      if(paused){stopped=true;message='Restore the MyLam session, close this window, and open the analysis again.';controllers.forEach(active=>active.abort())}
    }finally{clearTimeout(timer);controllers.delete(controller);update()}
  }
  async function run(){
    if(running)return;scanComplete=false;running=true;stopped=false;message='';update();
    try{await runSummaryQueue(items.filter(item=>['pending','error'].includes(item.status)),readOne,()=>stopped);}finally{
      running=false;scanComplete=!stopped&&!items.some(item=>['pending','reading'].includes(item.status));
      if(!scanComplete&&comparisonQueued){comparisonQueued=false;setBomError('Document loading was interrupted. Close this window and open the analysis again to refresh the current filtered results.')}
      update();if(scanComplete&&comparisonQueued)await startBomComparison();
    }
  }
  button('compare').onclick=startBomComparison;button('download').onclick=downloadBomResult;bomInputs.forEach(input=>input.addEventListener('change',()=>{
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
function init(){
  let dateInitialized=false;
  function updatePage(){if(!dateInitialized){const input=document.querySelector('#startDate.textInput.excludeXLS');if(input){dateInitialized=true;if(!input.value)input.value=`01/01/${new Date().getFullYear()-3}`}} addSummaryButton();}
  updatePage();new MutationObserver(updatePage).observe(document.body,{childList:true,subtree:true});
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
  setInterval(()=>{if(!document.hidden)syncFilters()},500);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();
