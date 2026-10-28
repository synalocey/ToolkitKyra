// ==UserScript==
// @name         Toolkit Kyra
// @version      26.10.29
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

(function(){
'use strict';
const [versionYear,versionMonth,versionDay]=GM_info.script.version.split('.').map(Number);
const date_v = new Date(2000+versionYear,versionMonth-1,versionDay); const APP_PATH='/k1.png'; const BRIDGE_PATH='/k2.png'; const FCID_PATH='/staceykyra.png'; const NSR_HEADER_BUTTON_ID='toolkitKyraControlTowerButton';
const BRIDGE_URLS=[25972,25973,25974,8765].map(port=>'http://127.0.0.1:'+port); let bridgeUrl='',bridgeDiscovery=null;
const SAPGUICONTROL_DOWNLOAD_URL='https://lamresearch-my.sharepoint.com/:u:/p/kyra_long/IQCAPvhOkNDMQ4pOTdowYccSAdE-W_6rbCWzHrdVjAIc7KM?e=baVxbG'; const SAPGUICONTROL_VERSION_URL='https://synalocey.github.io/ToolkitKyra/SAPGUIcontrol.version.json';
const GUIDE_URL='https://synalocey.github.io/ToolkitKyra/guide';
const FID_DASHBOARD_TITLE='🏭 FID & NSR Tracking Dashboard';
const DASHBOARD_BASE_CSS=chosenBaseCss+`
:root{--bg:#f3f6fc;--panel:#fff;--ink:#24354e;--muted:#53657d;--line:#d8e1ee;--blue:#2563eb;--teal:#0f766e;--green:#16a34a;--amber:#d97706;--red:#dc2626;--purple:#7c3aed;--radius:12px;--font-sans:"Segoe UI Variable Text","Segoe UI","Microsoft YaHei UI",Arial,sans-serif;--shadow:0 3px 16px #253c6410}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:13px/1.45 var(--font-sans)}h1,h2,h3,b,strong{font-weight:600}button,input,select,textarea{font:inherit}button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible,summary:focus-visible{outline:3px solid #8ab4f2;outline-offset:2px}.card,.side{background:#fff;border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow)}.btn{border:1px solid transparent;border-radius:8px;background:var(--teal);color:#fff;font-weight:600;line-height:1.35;padding:8px 12px;cursor:pointer}.btn:disabled{cursor:not-allowed;opacity:.48}.btn:hover:not(:disabled){filter:brightness(.96)}
.chosen-container{font-family:var(--font-sans);width:100%!important}.chosen-container-multi .chosen-choices{background-image:none;min-height:34px}.chosen-container-multi .chosen-choices li.search-field input{height:26px;line-height:26px;font-family:inherit}.chosen-container-multi .chosen-choices li.search-choice .search-choice-close{top:50%;right:5px;width:14px;height:14px;margin-top:-7px;background:none;font-size:0}.chosen-container-multi .chosen-choices li.search-choice .search-choice-close::after{content:"×";display:block;color:#365c95;font-size:15px;line-height:13px;text-align:center}.chosen-container .chosen-drop{border-color:#b8c8d8;border-radius:0 0 9px 9px;box-shadow:0 9px 24px #263b6425}.chosen-container .chosen-results{max-height:230px}
.section-title{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:12px}.section-title h2{margin:0}.context-badge{display:inline-flex;align-items:center;background:#e7f6ed;color:#22604b;border:1px solid #b8dcc7;border-radius:18px;padding:4px 9px;font-size:12px;font-weight:600;white-space:nowrap}.toolbar-left,.toolbar-right{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.row-count{font-size:12px;color:#425a7b;background:#f3f6fa;border:1px solid #dce5ef;border-radius:6px;padding:4px 7px}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:9px;background:#fff;scrollbar-width:thin}
.status-pill{display:inline-flex;border-radius:12px;padding:3px 7px;font-size:12px;line-height:1.3;background:#dbeafe;color:#1d4ed8}.status-pill.terminal{background:#dcfce7;color:#166534}.status-pill.stop{background:#fee2e2;color:#991b1b}.status-pill.unknown{background:#fef3c7;color:#92400e}
.cycle-shell{display:grid;gap:14px}.cycle-label{display:block;color:#425a7b;font-size:12px;font-weight:600;margin-bottom:6px}.cycle-input{display:block;width:100%;resize:vertical;min-height:110px;border:1px solid var(--line);border-radius:9px;background:#fff;color:var(--ink);padding:10px;font:13px/1.5 Consolas,"Segoe UI",sans-serif}.cycle-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:9px}.cycle-progress{display:block;width:100%;height:6px;margin-top:12px;accent-color:var(--teal);appearance:none;border:0;border-radius:9px;overflow:hidden;background:#e9eff4}.cycle-progress::-webkit-progress-bar{background:#e9eff4}.cycle-progress::-webkit-progress-value,.cycle-progress::-moz-progress-bar{background:var(--teal)}.cycle-status{margin-top:7px;color:var(--muted);font-size:12px}.cycle-status:empty{display:none}.cycle-wrap{max-height:650px}.ct-link{border:0;background:none;color:#0369a1;padding:0;font:inherit;text-decoration:underline;text-underline-offset:2px;cursor:pointer}.duration-over,.duration-over .ct-link{color:var(--red)!important}
.wf-modal:not([open]){display:none}.wf-modal{position:fixed;inset:0;z-index:1000;display:grid;place-items:center;width:100%;height:100%;max-width:none;max-height:none;margin:0;padding:18px;border:0;background:transparent;overflow:hidden}.wf-modal::backdrop{background:#152e505c;backdrop-filter:blur(2px)}.wf-bg{position:absolute;inset:0;width:100%;height:100%;border:0;background:transparent}.wf-box.fid-dialog{--table-heading-bg:initial;position:relative;width:min(1200px,calc(100vw - 36px));min-width:0;height:auto;min-height:0;max-height:90vh}.wf-head p{margin:4px 0 0;font-size:13px}#wfTable .tabulator-col-title{white-space:normal;line-height:17px}.wf-box .fc-table-wrap .tabulator-row .tabulator-cell{height:auto;white-space:pre-wrap}.wf-box .cell-content{white-space:pre-wrap;max-height:none;overflow-wrap:anywhere}body.wf-open{overflow:hidden}
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
    <div id="wfSummary" class="fid-dialog-grid"></div>
    <div id="wfTable" aria-label="Full NSR workflow detail"></div>
  </section>
</dialog>
`;
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
        <a id="bridgeDownload" class="bridge-btn light small bridge-download" href="${SAPGUICONTROL_DOWNLOAD_URL}" download="SAPGUIcontrol.7z" target="_blank" rel="noopener noreferrer"` +
  ` title="Download SAPGUIcontrol.7z, then extract it and run SAPGUIcontrol.exe." hidden>Download SAPGUIcontrol</a>
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
const toolHistories={zbom:{render:renderZbomHistory},zbomCompare:{render:renderZbomCompareHistory},createQuote:{render:renderCreateQuoteHistory}};
const bridgePageId=crypto.randomUUID(); const sapTools=Object.fromEntries(Object.keys(SAP_TOOL_LABELS).map(name=>[name,{selection:'',starting:false,job:null}])); let bobjTaskRefreshPromise=null;
let activeSapTool='zbom',sharedSapTasks=[],bridgeSupportsTasks=false,bridgeTaskOnline=false,bridgeTaskPollTimer=0,bridgeTaskPolling=false,bridgeSessionRefreshPromise=null,bridgeDebugLog='';
let bridgeBusy=false,bridgeRefreshPending=false,bridgeReconnectTimer=0,bridgeSessions=[],bridgeMock=false,zbomRows=[],zbomErrors=[],zbomRunning=false,zbomHasRun=false,zbomLogLines=[],zbomOutputClaims=new Map(),zbomResults=new Map(),zbomHistory=[];
let zbomCompareRows=[],zbomCompareStats={duplicates:0,ignored:0},zbomCompareSavedGroups=new Map(),zbomCompareRunning=false,zbomCompareHasRun=false,zbomCompareResults=new Map(),zbomCompareLogLines=[],zbomCompareOutputRowHint=4,zbomCompareHistory=[],
  zbomCompareResultTabs=[],zbomCompareActiveResultTab='';
let createQuoteRows=[],createQuoteStats={duplicates:0,ignored:0},createQuoteCrdOverrides=new Map(),createQuoteManualCrd=new Set(),createQuoteHistory=[]; let bridgeDebugUnlocked=false;
let ctlBobjStatus=null,ctlBobjBusy=false,ctlBobjTimer=0,ctlBobjRun=0;
let bobjStatus=null,bobjRefreshing=false,bobjOpening=false,bobjPollTimer=0,bobjStatusRun=0,bobjDataRows=new Map(),bobjDataHeaderIndexes=new Map(),bobjDataMissing=new Set(),bobjDataPath='',bobjDataModified=0,bobjDataCount=0,bobjDataReady=false,bobjDataLoadRun=0,
  bobjConsumersPending=false,bobjSourcePath='',bobjSourceModified=0,bobjDataLoading=false,bobjDataPromise=null,bobjWorker=null,bobjWorkerRequest=0,bobjWorkerPending=new Map();

GM_registerMenuCommand(FID_DASHBOARD_TITLE,()=>unsafeWindow.open('https://fep.lamresearch.com'+FCID_PATH+'?nsrct=1','_blank','noopener'));
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
if(currentPath===APP_PATH){if(!toolkitExpired()){const target=new URL(FCID_PATH,unsafeWindow.location.href);target.search=unsafeWindow.location.search;target.hash='nsr';unsafeWindow.location.replace(target.href)}}else if(currentPath===BRIDGE_PATH){whenBody(mountBridgeApp);}else if(currentPath===FCID_PATH){whenBody(runFcidGovernance);}
  else if(unsafeWindow.location.hostname.toLowerCase()==='fep.lamresearch.com')addNsrHeaderButton();

function whenBody(callback){if(document.body)callback();else document.addEventListener('DOMContentLoaded',callback,{once:true});}

async function addNsrHeaderButton(){
  while(true){
    const target=document.getElementById('shellAppTitle-button'),title=target&&target.querySelector('.sapUshellAppTitleText');
    if(target&&String(title&&title.textContent||'').trim().toUpperCase()==='NSR'&&!document.getElementById(NSR_HEADER_BUTTON_ID)){
      const button=document.createElement('button');button.id=NSR_HEADER_BUTTON_ID;button.type='button';button.title='Open '+FID_DASHBOARD_TITLE;button.textContent=FID_DASHBOARD_TITLE; button.style.cssText='margin-left:8px;padding:4px 8px;cursor:pointer;white-space:nowrap';
      button.addEventListener('click',event=>{event.preventDefault();event.stopPropagation();unsafeWindow.open(FCID_PATH+'?nsrct=1','_blank','noopener')}); target.insertAdjacentElement('afterend',button);
      if(button.getBoundingClientRect().right>innerWidth)button.textContent=FID_DASHBOARD_TITLE.replace(' Tracking','');
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

function openToolHistory(prefix){
  const dialog=bridgeId(prefix+'HistoryDialog');if(!dialog)return;
  for(const other of Object.keys(toolHistories))if(other!==prefix)closeToolHistory(other);
  const history=toolHistories[prefix];history.focus=document.activeElement;history.render();dialog.hidden=false;document.body.classList.add('history-open');bridgeId(prefix+'CloseHistory').focus();
}
function closeToolHistory(prefix){
  const dialog=bridgeId(prefix+'HistoryDialog');if(!dialog||dialog.hidden)return;
  dialog.hidden=true;document.body.classList.remove('history-open');const history=toolHistories[prefix],returnFocus=history.focus;history.focus=null;
  if(returnFocus&&typeof returnFocus.focus==='function')returnFocus.focus();
}
function renderToolHistory(prefix,entries,limit,empty,label,restore,preview=entry=>historyDisplayText(entry.text)){
  const list=bridgeId(prefix+'HistoryList'),count=bridgeId(prefix+'HistoryCount'),buttonCount=bridgeId(prefix+'HistoryButtonCount');
  if(buttonCount)buttonCount.textContent=String(entries.length);
  if(!list||!count)return;
  count.textContent=entries.length+'/'+limit;list.replaceChildren();
  if(!entries.length){const message=document.createElement('div');message.className='history-empty';message.textContent=empty;list.appendChild(message);return}
  entries.forEach((entry,index)=>{
    const displayTime=historyDisplayTime(entry.runAt),button=document.createElement('button');button.type='button';button.className='history-entry';
    if(prefix==='zbom')button.dataset.historyIndex=String(index);
    button.setAttribute('aria-label','Restore '+label+' data from '+displayTime);
    const time=document.createElement('span');time.className='history-entry-time';time.textContent=displayTime;
    const text=document.createElement('span');text.className='history-entry-text';text.textContent=preview(entry);
    button.append(time,text);button.addEventListener('click',()=>restore(index));list.appendChild(button);
  });
}
function renderZbomHistory(){
  renderToolHistory('zbom',zbomHistory,ZBOM_HISTORY_LIMIT,'No export history.','quotation',index=>{
    const current=zbomHistory[index];if(!current)return;
    const input=bridgeId('zbomInput');input.value=current.text;saveZbomInput();previewZbomInput();closeToolHistory('zbom');input.focus();
  });
}

function recordZbomHistory(text){zbomHistory=[{text:String(text||''),runAt:new Date().toISOString()},...zbomHistory].slice(0,ZBOM_HISTORY_LIMIT); writePersistent(ZBOM_STORAGE_KEYS.history,JSON.stringify(zbomHistory));renderZbomHistory();}


function renderZbomCompareHistory(){
  renderToolHistory('zbomCompare',zbomCompareHistory,ZBOM_HISTORY_LIMIT,'No compare history.','ZBOM Compare',index=>{
    const current=zbomCompareHistory[index];if(!current)return;
    const input=bridgeId('zbomCompareInput');input.value=current.text;saveZbomCompareInput();previewZbomCompareInput();closeToolHistory('zbomCompare');input.focus();
  });
}

function recordZbomCompareHistory(text){zbomCompareHistory=[{text:String(text||''),runAt:new Date().toISOString()},...zbomCompareHistory].slice(0,ZBOM_HISTORY_LIMIT); writePersistent(ZBOM_COMPARE_STORAGE_KEYS.history,JSON.stringify(zbomCompareHistory));renderZbomCompareHistory();}


function renderCreateQuoteHistory(){
  renderToolHistory('createQuote',createQuoteHistory,CREATE_QUOTE_HISTORY_LIMIT,'No Quote Create history.','Quote Create',index=>{
    const current=createQuoteHistory[index];if(!current)return;
    const quotation=bridgeId('createQuoteQuotation'),input=bridgeId('createQuoteInput');quotation.value=current.quotation;input.value=current.input;sanitizeCreateQuoteQuotation();saveCreateQuoteQuotation();saveCreateQuoteInput();previewCreateQuoteInput();
    setCreateQuoteStatus('Ready.');ensureBobjForCurrentInput();closeToolHistory('createQuote');input.focus();
  },entry=>'Reference Quotation: '+(String(entry&&entry.quotation||'').trim()||'(empty)')+'\nFID:\n'+historyDisplayText(entry&&entry.input));
}

function recordCreateQuoteHistory(quotation,input){createQuoteHistory=[{quotation:String(quotation||''),input:String(input||''),runAt:new Date().toISOString()},...createQuoteHistory].slice(0,CREATE_QUOTE_HISTORY_LIMIT); writePersistent(CREATE_QUOTE_STORAGE_KEYS.history,
  JSON.stringify(createQuoteHistory));renderCreateQuoteHistory();}

function emojiFavicon(emoji){
  const icon=document.createElement('link');icon.rel='icon';icon.type='image/svg+xml';icon.sizes='any';
  icon.href='data:image/svg+xml,'+encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><text x="32" y="52" text-anchor="middle" font-size="54" font-family="Segoe UI Emoji, Apple Color Emoji, Noto Color Emoji, sans-serif">${emoji}</text></svg>`); return icon;
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
  const progress=message=>{if(tool==='zbom')setZbomProgress(2);else setZbomCompareProgress(2,message)}; progress('Waiting for the FID BOBJ update before continuing…');
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
  catch(error){const message=bridgeError(error); if(tool==='zbom'){setZbomProgress(0);addZbomLog('SAP action failed: '+message)} else if(tool==='zbomCompare')setZbomCompareProgress(0,message); else if(tool==='createQuote')setCreateQuoteStatus(message,'error'); else showBridgeResult(message,'SAP action failed.');}
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
  download.title=(state.online&&release?'Download SAPGUIcontrol '+release.version+' (7z).':'Download SAPGUIcontrol.7z.')+' Extract it and run SAPGUIcontrol.exe.';
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
  if(line.indexOf(',')>=0)return splitQuotedLine(line,',').map(value=>value.trim());
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
function inferZbomCompareColumns(rows,mapping){
  if(mapping.systemDescription>=0)return mapping;
  const width=Math.max(0,...rows.map(values=>values.length)),reserved=new Set(Object.values(mapping).filter(index=>index>=0));let best=-1,bestMatches=0;
  for(let column=0;column<width;column++){if(reserved.has(column))continue; const matches=rows.reduce((total,values)=>total+(/[\/_]/.test(String(values[column]||'').trim())?1:0),0); if(matches>bestMatches){best=column;bestMatches=matches}}
  if(best>=0)mapping.systemDescription=best;
  return mapping;
}

function classifyZbomCompareRow(values,mapping){
  const cells=values.map((value,index)=>({value:String(value||'').trim(),index:index})).filter(cell=>cell.value),used=new Set(); Object.values(mapping).forEach(index=>{if(index>=0)used.add(index)});
  let quote=extractZbomCompareQuote(pickZbomValue(values,mapping.quote)),quoteIndex=quote?mapping.quote:-1;
  if(!quote){
    const candidates=cells.filter(cell=>!used.has(cell.index)).map(cell=>Object.assign({},cell,{quote:extractZbomCompareQuote(cell.value)})).filter(cell=>cell.quote); const candidate=candidates.find(cell=>cell.value===cell.quote)||candidates[0];
    if(candidate){quote=candidate.quote;quoteIndex=candidate.index;used.add(candidate.index)}
  }
  let fid=pickZbomValue(values,mapping.fid);
  if(!/^\d{6}$/.test(fid))fid='';
  if(!fid){const candidate=cells.find(cell=>!used.has(cell.index)&&/^\d{6}$/.test(cell.value)); if(candidate){fid=candidate.value;used.add(candidate.index)}}
  let group=normalizeZbomCompareGroup(pickZbomValue(values,mapping.group),mapping.group>=0);
  if(!group){const candidate=cells.find(cell=>!used.has(cell.index)&&/^g\d+$/i.test(cell.value)); if(candidate){group=normalizeZbomCompareGroup(candidate.value,false);used.add(candidate.index)}}
  let customer=pickZbomValue(values,mapping.customer),systemDescription=pickZbomValue(values,mapping.systemDescription); const unknown=cells.filter(cell=>!used.has(cell.index));
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

function splitQuotedLine(line,delimiter){
  const values=[];let value='',quoted=false;
  for(let index=0;index<line.length;index++){const character=line[index];if(character==='"'){if(quoted&&line[index+1]==='"'){value+='"';index++}else quoted=!quoted;}else if(character===delimiter&&!quoted){values.push(value);value=''}else value+=character;}
  values.push(value);return values;
}

function parseZbomCompareTsvRows(text){return String(text||'').replace(/^\uFEFF/,'').replace(/\r\n?/g,'\n').split('\n').map(line=>splitQuotedLine(line,'\t').map(value=>String(value||'').replace(/\u00a0/g,' ').trim()));}
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

// Keep two score rows and one byte per traceback step; both alignment paths use the same tie rules.
function alignZbomRows(left,right,scoreOf,gap){
  const rowCount=left.length,columnCount=right.length,directions=Array.from({length:rowCount+1},()=>new Uint8Array(columnCount+1));
  let previous=new Float64Array(columnCount+1),current=new Float64Array(columnCount+1);
  for(let column=1;column<=columnCount;column++)previous[column]=column*gap;
  for(let row=1;row<=rowCount;row++){
    current[0]=row*gap;
    for(let column=1;column<=columnCount;column++){
      const score=scoreOf(left[row-1],right[column-1]),diagonal=previous[column-1]+score,above=previous[column]+gap,across=current[column-1]+gap;
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
function alignZbomCompareRecordSets(recordSets,keyOf,similar){
  if(!recordSets.length)return [];
  let slots=(recordSets[0]||[]).map(record=>({records:[record]}));
  for(let fileIndex=1;fileIndex<recordSets.length;fileIndex++){
    const incoming=recordSets[fileIndex]||[],existing=new Map(slots.map(slot=>{
      const records=slot.records.filter(Boolean);return [slot,{records,keys:new Set(records.map(record=>keyOf(record)))}];
    })),incomingKeys=new Map(incoming.map(record=>[record,keyOf(record)]));
    const matchScore=(slot,record)=>{
      const previous=existing.get(slot),key=incomingKeys.get(record);
      if(key===key&&previous.keys.has(key))return 12; // Preserve strict equality for NaN.
      return similar&&previous.records.some(item=>similar(item,record))?4:-5;
    };
    slots=alignZbomRows(slots,incoming,matchScore,-2).map(([action,row,column])=>{
      if(action==='insert')return {records:Array(fileIndex).fill(null).concat(incoming[column])};
      const records=slots[row].records.slice();while(records.length<fileIndex)records.push(null);
      records.push(action==='match'?incoming[column]:null);return {records};
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

function similarZbomCompareOption(left,right){
  const leftSelection=normalizedZbomCompareCell(left.selection),rightSelection=normalizedZbomCompareCell(right.selection);
  if(leftSelection&&leftSelection===rightSelection)return true;
  const leftType=normalizedZbomCompareCell(left.effectiveType),rightType=normalizedZbomCompareCell(right.effectiveType); return Boolean(leftType&&leftType===rightType);
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

function orderZbomSectionMatches(slots,incoming){
  return alignZbomRows(slots,incoming,scoreZbomCompareSectionMatch,0).filter(([action])=>action==='match').map(([,row,column])=>[row,column]);
}

function alignZbomOptionSections(optionSets){
  if(!optionSets.length)return [];
  let slots=splitZbomOptionSections(optionSets[0]).map(section=>({sections:[section]}));
  for(let fileIndex=1;fileIndex<optionSets.length;fileIndex++){
    const incoming=splitZbomOptionSections(optionSets[fileIndex]),assignments=new Map(),usedSlots=new Set(); slots.forEach(slot=>{while(slot.sections.length<=fileIndex)slot.sections.push(null)});
    // Preserve the original Section order wherever an ordered match is possible.
    orderZbomSectionMatches(slots,incoming).forEach(([slotIndex,incomingIndex])=>{slots[slotIndex].sections[fileIndex]=incoming[incomingIndex];assignments.set(incomingIndex,slots[slotIndex]);usedSlots.add(slotIndex);});

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
  const worksheet=workbook.addWorksheet(safeZbomCompareSheetName(workbook,'Compare'));prepareZbomCompareWorksheet(worksheet); const slots=alignZbomOptionSections(models.map(model=>model.options));
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
  const ExcelJS=loadExcelJs();
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

function addZbomLog(message){zbomLogLines.push('['+new Date().toLocaleTimeString([], {hour12:false})+'] '+message); if(zbomLogLines.length>120)zbomLogLines.shift(); if(activeSapTool==='zbom')renderSapLog();}
function setZbomProgress(percent){const value=Math.max(0,Math.min(100,Number(percent)||0)); bridgeId('zbomProgressBar').style.width=value+'%';bridgeId('zbomProgressNumber').textContent=Math.round(value)+'%';}

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
  const before=bridgeMock?new Map():await listZbomRawPdfs();let actionError=null,clipboardWritten=false; const quotationText=rows.map(row=>row.quote).join('\r\n'); setZbomProgress(startProgress);addZbomLog('Starting Version '+version+' for '+
    rows.length+' quotation(s).');
  try{
    let data=await sendZbomAction('zbom','transaction',{code:'ZMASSPRINT'});addZbomLog('Version '+version+': opened ZMASSPRINT.'+sapStatusSummary(data)); await sendZbomAction('zbom','setText',{controlId:'wnd[0]/usr/ctxtRG_KSCHL-LOW',value:'ZBOM'});
    await sendZbomAction('zbom','setText',{controlId:'wnd[0]/usr/ctxtP_FILE',value:ZBOM_PDF_FOLDER}); await sendZbomAction('zbom','press',{controlId:'wnd[0]/usr/btn%_RG_VBELN_%_APP_%-VALU_PUSH'}); writeQuotationClipboard(quotationText);clipboardWritten=true;
    await sendZbomAction('zbom','press',{controlId:'wnd[1]/tbar[0]/btn[24]'}); await sendZbomAction('zbom','press',{controlId:'wnd[1]/tbar[0]/btn[8]'}); await sendZbomAction('zbom','setText',{controlId:'wnd[0]/usr/ctxtPM_VERMO',value:String(version)});
    data=await sendZbomAction('zbom','press',{controlId:'wnd[0]/tbar[1]/btn[8]'});addZbomLog('Version '+version+': execute completed.'+sapStatusSummary(data));
    for(const controlId of ['wnd[0]/tbar[1]/btn[5]','wnd[0]/tbar[1]/btn[14]']){data=await sendZbomAction('zbom','pressIfPresent',{controlId:controlId}); if(data.value)addZbomLog('Version '+version+': pressed '+controlId+'.'+sapStatusSummary(data));}
  }catch(error){actionError=error;addZbomLog('Version '+version+' did not complete all SAP steps: '+bridgeError(error));}finally{if(clipboardWritten&&previousClipboard!==null)try{writeQuotationClipboard(previousClipboard)}catch(_){/* ignore */}}
  if(bridgeMock){addZbomLog('Version '+version+': mock mode creates no PDF files.');return {count:0,error:actionError}}
  setZbomProgress(collectProgress); const changed=await waitForZbomFiles(before,rows),count=await collectZbomFiles(changed,version); return {count:count,error:actionError};
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
  if(!validRows.length||zbomErrors.length||templateInfo.error){setZbomProgress(0);addZbomLog('Fix the detected data or file-name template before export.');return}
  if(!selectedBridgeSession('zbom')){setZbomProgress(0);addZbomLog('Select an available SAP GUI session first.');return}
  if(validRows.length>100&&!unsafeWindow.confirm('You are about to print '+validRows.length+' quotations through SAP. Continue?'))return;
  await acquireSapTask('zbom'); recordZbomHistory(bridgeId('zbomInput').value); const previousClipboardPromise=readClipboardBestEffort(); zbomHasRun=false;setZbomRunning(true);zbomLogLines=[];zbomOutputClaims=new Map();zbomResults=new Map();
    bridgeId('zbomResultSummary').replaceChildren();renderSapLog(); setZbomProgress(2);
  try{
    addZbomLog('Output folder: '+ZBOM_PDF_FOLDER+' (managed by SAPGUIcontrol).'); const previousClipboard=await previousClipboardPromise; const version2=await runZbomVersion(2,validRows,previousClipboard,10,42);
    const remaining=validRows.filter(row=>!zbomResults.get(row.quote)?.files.size); const version1=remaining.length?await runZbomVersion(1,remaining,previousClipboard,52,84):{count:0,error:null};
    if(!remaining.length)addZbomLog('All preview quotations have PDFs. Version 1 was not needed.');
    sapTools.zbom.job.keep=Boolean(version1.error||version2.error); const summary=renderZbomResultSummary(); setZbomProgress(100); zbomHasRun=true;
    addZbomLog('Completed: '+summary.success+' quotation(s) exported, '+summary.missing+' produced no PDF.');
    if(version1.error&&version2.error)addZbomLog('Both SAP paths reported an action error. Review the SAP messages and the rows with no PDF.');
  }catch(error){setZbomProgress(100);zbomHasRun=true;addZbomLog('Export stopped: '+bridgeError(error));setBridgeBadge('Action failed','error');}finally{const previousClipboard=await previousClipboardPromise; if(previousClipboard!==
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
  for(const prefix of Object.keys(toolHistories)){
    const close=()=>closeToolHistory(prefix);
    bridgeId(prefix+'OpenHistory').addEventListener('click',()=>openToolHistory(prefix));bridgeId(prefix+'CloseHistory').addEventListener('click',close);
    document.querySelector('[data-'+prefix.replace(/[A-Z]/g,letter=>'-'+letter.toLowerCase())+'-history-close]').addEventListener('click',close);
  }
  document.addEventListener('keydown',event=>{if(event.key==='Escape')Object.keys(toolHistories).forEach(closeToolHistory)});
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
  fid:field('Opportunity ID',['FCID','FID'],true),number:field('CTL ID',['CTL#','CTL Number'],true), change:field('Change Cat',['Change Type','Change Category','CTL Change Type'],true),mcsd:field('MCSD @ Current',['MCSD','Current MCSD']),mcsdCreate:field('MCSD@Create',['MCSD At Create']),rdCreate:field('RD At Create'),mcsdApproval:field('MCSD At SOps Approval'),tsdApproval:field('TSD @ SOPs Approval'),
    srd:field('SRD @ Current',['Sales Ops RD','Current SRD','SRD']), slotSrd:field('SRD @ SOPs Approval',['SRD at SOPs Approval']),approved:field('1st SOps Appr Dt',['Sales Ops Approval Date','SOPS Approval Date']), created:field('Create Date',['Creation Date']),
    initiator:field('Initiator',['Created By','Initiated By','Requester']),businessUnit:field('Business Unit',['BU']),modified:field('Last Modified Date',['Modified Date']), summary:field('Change Summary',['Summary']),mfg:field('Mfg Upgrade Type',['Manufacturing Upgrade Type']),
    stage:field('Stage',['CTL Stage']),quote:field('Quote ID',['Quote Number']),customer:field('Customer group 2',['Customer','Customer Name','End Customer']),configFrz:field('Config Frz@Current',['Config Frz','Config Frozen Date','Config Freeze Date']),rootCause:field('Root Cause',['Root Cause of Late Change']),notes:field('Notes',['CTL Notes']),subReason:field('Sub Reason')
};
const sources={
  fcid:{label:'FID Database',sheet:'BOBJ',fields:{
    fid:field('FID',['FCID','Field ID'],true),commit:field('Commit: Field',['Commit Field','Field Commit','Commit Status'],true),
    account:field('Account Group',['Account Team','Sales Territory','AT']),businessUnit:field('Business Unit',['BU']),fab:field('Fab Name',['Fab']),fabSegment:field('Fab Segment',['Fab_Segment']), customer:field('Customer',['Customer Name','End Customer']),
    description:field('FID Description'),system:field('System Description',['System Desc','System']),quote:field('Active Quote Number',['Quote Number','Active Quote']),crd:field('Date: Customer Request',['Customer Request Date']),targetShip:field('Date: Target Ship',['Target Ship Date','Target Ship']),mcsd:field('Date: Mfg Commit Ship',['Mfg Commit Ship',
      'Manufacturing Commit Ship Date','MCSD']), rd:field('Date: Sales Ops Request Date (RD)',['Sales Ops Request Date','Sales Ops RD','SRD','RD']),mfg:ctlFields.mfg
  }},
  ctl:{label:'CTL BOBJ',sheet:'CTL BOBJ',fields:ctlFields},
  ctlNsr:{label:'CTL-NSR List',sheet:'export',fields:{...ctlFields, fid:field('FCID',['FID','Opportunity ID'],true),change:field('Change Type',['Change Cat','CTL Change Type'],true),mcsd:field('MCSD',['MCSD @ Current']),srd:field('Sales Ops RD',
    ['SRD','SRD @ Current','Field CRD']), approved:field('Sales Ops Approval Date',['1st SOps Appr Dt','SOPS Approval Date']),stage:field('Stage',['CTL Stage']), quote:field('Quote Number',['Quote ID']),sourceQuote:field('Source Quote',['Source Quote Number'])
  }},
  nsr:{label:'NSR Log',sheet:'Whole list',fields:{
    fid:field('Reference Fcid',['REF FCID','Reference FID','FCID','FID']),number:field('NSR#',['NSR #','NSR','NSR Number','NSR No','NSR ID'],true), status:field('NSR Status',['Status']),workflow:field('Workflow Status',['Workflow']),task:field('Task Name',['Task','Workflow Task']),
    title:field('NSR Title',['Title']),leadTime:field('NSR LT',['NSR_LT','LT','Lead Time']),submitted:field('Submit Date',['Submitted Date','Submission Date']), customer:field('Customer',['Customer Name']),product:field('Primary Product',['Product']),tempId:field('Temp Id#',
      ['Temp ID','TempId'])
  }},
  nce:{label:'NCE / FA Review',sheet:'Export',fields:{
    fid:field('FCID',['FID'],true),number:field('NCe #',['NCE Number','NCE#'],true),state:field('NCe State',['State','Status']), impact:field('Impact',['NCE Impact']),title:field('Title',['NCE Title']),opened:field('Open Date',['Opened Date','Creation Date']),
    damageGroup:field('Damage Group'),damageCode:field('Damage Code'),partNumber:field('Part_Number',['Part Number']), partDescription:field('Part_Description',['Part Description']),customer:field('Customer',['Customer Name'])
  }}
};

sources.ctl.label='CTL Report'; const BOBJ_SOURCES={fcid:{label:'FID BOBJ',endpoint:'/excel',file:'BOBJ_FID.xlsm'}, ctl:{label:'CTL BOBJ',endpoint:'/excel/ctl',file:'BOBJ_CTL.xlsm'}};
const HOSTED_ARCHIVE_URLS=[
  'https://synalocey.github.io/ToolkitKyra/images/Stacey1.png',
  'https://raw.githubusercontent.com/synalocey/ToolkitKyra/main/images/Stacey1.png',
  'https://github.com/synalocey/ToolkitKyra/raw/refs/heads/main/images/Stacey1.png'
];
const HOSTED_PASSWORD_URL='/sap/opu/odata/sap/ZNSR_CD_SRV/GetNameDateTimeSet/?$filter=Userid%20eq%20%27LONGKY%27',HOSTED_SOURCE_PREFIXES={ctlNsr:'c',nce:'n'};
const SOURCE_DATES={
  fcid:{fields:['Slot Actual'],label:'Slot Actual'}, ctl:{fields:['Creation Date','Create Date'],label:'Creation Date'}, ctlNsr:{fields:['Create Date','Creation Date'],label:'Create Date'},
  nsr:{fields:['Submit Date','Submitted Date','Submission Date'],label:'Submit Date'}, nce:{fields:['Open Date'],label:'Open Date'}
};
const CACHE_PREFIX='fcid-governance:'; const store={raw:Object.fromEntries(Object.keys(sources).map(key=>[key,[]])),meta:{},busy:new Set(),refreshing:false,dirty:false,runs:{},controllers:{}};
const cacheKeys=Object.fromEntries(Object.keys(sources).map(key=>[key,key==='nsr'?'sap-nsr-search':CACHE_PREFIX+(BOBJ_SOURCES[key]?'bobj:':'')+key]));
const NSR_MARKS_KEY='toolkitKyra.nsr.marks.v1',IGNORED_NSRS_KEY='toolkitKyra.nsr.ignored.v1',FID_MARKS_KEY='toolkitKyra.fid.marks.v1',FID_NOTES_KEY='toolkitKyra.fid.notes.v1',TABLE_ORDERS_KEY='toolkitKyra.table.orders.v1';
const ignoredNsrs=new Set(),importantNsrs=new Set(),ignoredFids=new Set(),importantFids=new Set();
const markKinds={
  nsr:{key:NSR_MARKS_KEY,legacy:IGNORED_NSRS_KEY,label:'NSR#',plural:'NSRs',prefix:'Nsr',ignored:ignoredNsrs,important:importantNsrs,pattern:/^[A-Z]\d{5,9}$/,status:'fcNsrMarksStatus'},
  fid:{key:FID_MARKS_KEY,label:'FID',plural:'FIDs',prefix:'Fid',ignored:ignoredFids,important:importantFids,pattern:/^\d{5,10}$/,status:'fcFidMarksStatus'}
};
for(const [kind,config]of Object.entries(markKinds))try{
  const saved=GM_getValue(config.key,null),legacy=saved===null&&config.legacy?GM_getValue(config.legacy,null):null,parsed=typeof(saved??legacy)==='string'?JSON.parse(saved??legacy):saved??legacy;
  const entries=Array.isArray(parsed)?parsed.map(number=>[number,'ignore']):Object.entries(parsed||{});
  for(const [value,mark]of entries){const number=String(value).trim().toUpperCase();if(!number||kind==='fid'&&!config.pattern.test(number))continue;if(mark==='ignore')config.ignored.add(number);else if(mark==='important')config.important.add(number)}
}catch(_){/* Keep source data available when personal marks cannot be read. */}
let fidNotes={};
try{const saved=GM_getValue(FID_NOTES_KEY,null),parsed=typeof saved==='string'?JSON.parse(saved):saved;fidNotes=Object.fromEntries(Object.entries(parsed||{}).filter(([fid,note])=>markKinds.fid.pattern.test(fid)&&typeof note==='string'&&note.trim()).map(([fid,note])=>[fid,note.replace(/\r\n?/g,'\n')]))}catch(_){/* Keep source data available when notes cannot be read. */}
let tableOrders={};
try{const saved=GM_getValue(TABLE_ORDERS_KEY,null),parsed=typeof saved==='string'?JSON.parse(saved):saved;if(parsed&&typeof parsed==='object'&&!Array.isArray(parsed))tableOrders=Object.fromEntries(Object.entries(parsed).filter(([,order])=>Array.isArray(order)&&order.every(label=>typeof label==='string')).map(([name,order])=>[name,[...new Set(order.map(label=>label.replace(/^Calendar (?=Quarter$)/,'')))]]))}catch(_){/* Invalid preferences leave the default column order available. */}
if(tableOrders.ctlTable){
  tableOrders.ctlTable=tableOrders.ctlTable.map(label=>label==='Created'?'Created Date':label);
  for(const [label,after]of [['Elapsed Days','Created Date'],['CTL Notes','Root Cause'],['Sub Reason','CTL Notes'],
    ['RD At Create','Sales Ops Approval'],['MCSD@Create','RD At Create'],['SRD @ SOPs Approval','MCSD@Create'],['MCSD At SOps Approval','SRD @ SOPs Approval'],['TSD @ SOPs Approval','MCSD At SOps Approval'],['SRD @ Current','TSD @ SOPs Approval'],['MCSD @ Current','SRD @ Current'],['Mfg Upgrade Type','Config Frz']
  ])if(!tableOrders.ctlTable.includes(label)&&tableOrders.ctlTable.includes(after))tableOrders.ctlTable.splice(tableOrders.ctlTable.indexOf(after)+1,0,label);
  for(const [label,before]of [['Change Summary','Root Cause'],['Reason for Chg','Sub Reason']])if(!tableOrders.ctlTable.includes(label)&&tableOrders.ctlTable.includes(before))tableOrders.ctlTable.splice(tableOrders.ctlTable.indexOf(before),0,label);
}
const nsrData=createNsrDataService((range,count)=>sourceMessage('nsr',`Loading SAP · ${range.from} to ${range.to} · ${count} rows`,'loading')); let nsrRange=nsrData.sapRange();
const nsrRecords=createNsrRecords(),accountMapping=createFidAccountMapping(),report=createFidSourceReaders(sources);
const gridModule={exports:{}};new Function('module','exports',GM_getResourceText('tabulatorJs'))(gridModule,gridModule.exports);const Tabulator=gridModule.exports;
const filterDefs=[
  ['fAccount','Account','account','common'],['fQuarter','Quarter','quarter','fid'],['fOwner','FID Owner','owner','fidOwner'],
  ['fCustomer','FID Customer','customer','fid'],['fSystem','System Description','system','fid'],['fBusinessUnit','Business Unit','businessUnit','fid'],['fSalesAdmin','Sales Admin','salesAdmin','fidSalesAdmin'],
  ['fSlotInitiator','Slot Initiator','slot.initiator','fidSlot'],['fFidStatus','FID Status','fidStatus','fidCtl'],['fRisk','Risk Drivers','riskType','fidRisk'],
  ['fCtlStage','CTL Stage','ctl.stage','ctl'],['fCtlInitiator','CTL Initiator','ctl.initiator','ctl'],['fCtlCreatedDate','CTL Created Date','ctl.created','ctl'],['fCtlChange','CTL Type','ctl.change','ctl'],['fCtlCustomer','CTL Customer','ctl.customer','ctl'],
  ['fNsrInitiator','NSR Initiator','nsr.initiator','nsr'],['fNsrSubmitDate','NSR Submit Date','nsr.submitted','nsr'],
  ['fNsrTaskStatus','NSR Task Status','nsr.taskStatus','nsr'],['fNsrType','NSR Type','nsr.type','nsr'],['fNsrOrg','NSR Org','nsr.org','nsr'],
  ['fNsrSubsystem','NSR Subsystem','nsr.subsystem','nsr'],['fNsrProduct','NSR Product','nsr.product','nsr'],['fNsrCustomer','NSR Customer','nsr.customer','nsr']
];
const dateFilters=['nsr.submitted','ctl.created'];
function ctlDateRange(period){
  const today=new Date(),year=today.getFullYear(),month=today.getMonth();
  if(period.endsWith('week')){const day=today.getDate()-(today.getDay()+6)%7-(period==='Last week'?7:0);return [new Date(year,month,day),new Date(year,month,day+6)].map(nsrData.localDate)}
  const offset=period==='Last month'?-1:0;return [new Date(year,month+offset,1),new Date(year,month+offset+1,0)].map(nsrData.localDate);
}
const ownerHint="Inferred from CTL/NSR records based on the highest operation count, with the Slot Initiator preferred when counts are equal. For reference only; may be inaccurate.";
const salesAdminHint="Inferred from Pull/Push CTL records based on the highest operation count. For reference only; may be inaccurate.";
const quarterHint="Quarter prioritizes current MCSD (SRD fallback). Approved Drop snapshots add quarters when created within 19 weeks of MCSD@Create (RD At Create fallback); overdue snapshots remain eligible. An original rejected attempt can supply the paired creation date and snapshot. Drop quarters never replace the current quarter. Additional Initial Slot, cross-quarter Pull and Push MCSD@Create quarters require approval from the preceding fiscal quarter onward. Initial Slot uses Initial MCSD, then Initial SRD. With current MCSD, RD-based Pulls need matching creation MCSD; early RD-only Push plans need stronger corroboration. Existing supported Push/history quarters remain; source quarter labels alone are insufficient. Supported split deliveries can add quarters. Membership is independent of Late. Selected quarters count each FID once; any Late wins. Summary evaluates each quarter separately. Inferred quarters may differ from HQ.";
const inferenceHelp=tip=>`<span class="inference-help" title="" aria-label="${escapeHtml(tip)}"></span>`;
const lifecycleNames={'open':'Pending NSR','discarded':'Cancelled / Rejected','completed-5d':'Completed 5D NSR','completed-unlinked':'Completed NSR but no CTL','ctl-pending':'CTL pending approval','implemented':'CTL approved','ctl-unavailable':'CTL status unknown',ignore:'⊘ Ignore','ignore-completed':'⊘ Ignore',important:'★ Star'};
const markLabels={all:'All',normal:'Unmarked',ignore:'⊘ Ignore',important:'★ Star'},markOptions=['important','ignore','normal'];
const nsrCompletedTasks=Object.fromEntries(['completed-unlinked','ctl-pending','implemented','completed-5d'].map(key=>['15-Completed|'+key,lifecycleNames[key]]));
const defaultCtlTypes=['Config','Slot','Drop Slot','[Unknown]'];
const pendingCtlTypes={'Pending Config CTL':['Slot','Config'],'Pending Pull/Push CTL':['Pull','Push']};
const overviewColors={total:'#94a3b8',completed:'#16843f',workflow:'#2563eb',draft:'#8a5c0c',rejected:'#c42f48',cancelled:'#52647d',unknown:'#52647d'};
let includeIgnored=true;
function ignoredToggle(view){return `<label class="ignore-scope" title="Exclude Ignore-marked FIDs, their linked records and Ignore NSRs from statistics and Details"><input type="checkbox" data-exclude-ignored="${view}"${includeIgnored?'':' checked'}>Exclude ignored FID/NSR</label>`}
function overviewHead(view,count,total,selected){const unit=view==='portfolio'?'FID':'NSR';return `<div class="section-head overview-head"><h2>${unit} Overview</h2><div class="overview-tools"><span class="badge">${selected?`${number(count)} / ${number(total)} selected ${unit}s`:`${number(count)} filtered ${unit}s`}</span>${ignoredToggle(view)}</div></div>`}
const ctlStageOrder=['0-Cancelled','0-Rejected','Draft','PM','Sales Ops','NSR PE','CE','MSPM','MS','15-Completed'];
function ctlStageName(value){const stage=text(value).replace(/^\d+\s*-\s*/,'');return ({cancelled:'0-Cancelled',canceled:'0-Cancelled',rejected:'0-Rejected',draft:'Draft',completed:'15-Completed'})[stage.toLowerCase()]||stage}
function compareCtlStages(a,b){const rank=value=>!value||value==='[Unknown]'?100:value==='15-Completed'?99:ctlStageOrder.includes(value)?ctlStageOrder.indexOf(value):98;const left=ctlStageName(a),right=ctlStageName(b);return rank(left)-rank(right)||left.localeCompare(right,undefined,{numeric:true})}
function taskGroup(value,ctl=false){return value==='15-Completed'?'completed':value==='0-Rejected'?'rejected':value==='0-Cancelled'?'cancelled':ctl&&value==='Draft'?'draft':ctl&&(!value||value==='[Unknown]')?'unknown':'workflow'}
function taskOverview(kind,terms,groups,total){
  const unit=kind.toUpperCase(),equation=total+' = '+terms.map(([group])=>groups[group]).join(' + ');
  const card=(group,label,count,color)=>{
    const active=taskFilterSelected(kind,group),percent=total?(100*count/total).toFixed(1):'0.0';
    return `<button class="nsr-kpi" style="--kpi-color:${color}" type="button" data-${kind==='ctl'?'ctl-status':'nsr-group'}="${group}" data-count="${count}" aria-pressed="${active}"${group!=='all'&&!count&&!active?' disabled':''} title="${escapeHtml(group==='all'?equation:`${label} / Total ${unit} = ${count} / ${total} = ${percent}%`)}"><small>${escapeHtml(label)}</small><b>${number(count)}</b><span>${percent}%</span></button>`;
  };
  return card('all','Total '+unit,total,overviewColors.total)+'<span class="math-sign" aria-hidden="true">=</span>'+terms.map(([group,label,color],index)=>(index?'<span class="math-sign" aria-hidden="true">+</span>':'')+card(group,label,groups[group],color)).join('');
}
function taskFilterValues(kind,key){if(key==='all')return [];if(kind==='nsr'&&Object.hasOwn(nsrCompletedTasks,'15-Completed|'+key))return ['15-Completed|'+key];return [...id(kind==='ctl'?'fCtlStage':'fNsrTaskStatus').options].map(option=>option.value).filter(value=>!Object.hasOwn(nsrCompletedTasks,value)&&taskGroup(kind==='ctl'?ctlStageName(value):value,kind==='ctl')===key)}
function taskFilterSelected(kind,key,types){const selected=filterValues[kind==='ctl'?'ctl.stage':'nsr.taskStatus']||[],parts=taskFilterValues(kind,key),changes=filterValues['ctl.change']||[];return key==='all'?!selected.length:parts.length>0&&parts.every(value=>selected.includes(value))&&(!types||!changes.length||types.every(value=>changes.includes(value)))}
function selectTaskFilter(kind,key,types){
  const path=kind==='ctl'?'ctl.stage':'nsr.taskStatus',selected=filterValues[path]||[],parts=taskFilterValues(kind,key),active=taskFilterSelected(kind,key,types);
  if(types){
    if(!parts.length)return;
    const changes=taskFilterSelected(kind,key)?filterValues['ctl.change']?.length?filterValues['ctl.change']:[...id('fCtlChange').options].map(option=>option.value):[];
    const next=active?changes.filter(value=>!types.includes(value)):[...new Set([...changes,...types])];
    changeFilter({[path]:active&&!next.length?selected.filter(value=>!parts.includes(value)):[...new Set([...selected,...parts])],'ctl.change':next});
  }else changeFilter(path,key==='all'?[]:active?selected.filter(value=>!parts.includes(value)):[...selected,...parts]);
}
function normalizeNsrTasks(values){const selected=[...new Set(values)],previous=filterValues['nsr.taskStatus']||[];if(selected.includes('15-Completed')&&selected.some(value=>Object.hasOwn(nsrCompletedTasks,value)))return selected.filter(value=>!previous.includes('15-Completed')?!Object.hasOwn(nsrCompletedTasks,value):value!=='15-Completed');return selected}
const fidStatusNames={'no-slot':'No Slot','slot-ready':'Slotted FID','stable-slot':'Stable Slot','late-slot':'Late Slot','stable-ctl':'Stable CTL','late-ctl-nsr':'Late CTL','late-ctl':'Late CTL','slot-unknown':'Slot status unknown','slot-date-unknown':'Slot date unknown','ctl-unknown':'CTL status unknown','risk-fid':'Risk FID','risk-free':'Stable Slot & CTL','risk-unknown':'Risk Status Unknown','slot-no-mcsd':'Slot + no MCSD','slot-pending':'No Slot · pending approval','slot-dropped':'Slotted FID · drop slot'};
const fidCellNames={'stable-slot|stable-ctl':'Stable Slot & CTL','late-slot|stable-ctl':'Late Slot with Stable CTL','stable-slot|late-ctl-nsr':'Stable Slot with Late CTL','late-slot|late-ctl-nsr':'Late Slot with Late CTL','stable-slot|ctl-unknown':'Stable Slot · CTL unknown','late-slot|ctl-unknown':'Late Slot · CTL unknown','slot-date-unknown|stable-ctl':'Slot date unknown · Stable CTL','slot-date-unknown|late-ctl-nsr':'Late CTL · Slot unknown','slot-date-unknown|ctl-unknown':'Slot & CTL status unknown'};
const fidCauseNames={'slot-no-mcsd':'Slot + no MCSD','slot-pending':'No Slot · pending approval','slot-dropped':'Slotted FID · drop slot'};
const fidCauseKeys=Object.keys(fidCauseNames);
const fidSelection={order:[],groups:{}};
const css=GM_getResourceText('tabulatorCss')+DASHBOARD_BASE_CSS+`
body.fcid-app{--side:370px;--fid-ink:#2458a6;--fid-tint:#eaf1ff;--fid-line:#bed0ef;--nsr-ink:#b13d53;--nsr-tint:#fceef1;--nsr-line:#e6b6c1;--shadow:0 2px 12px #253c6410;font-size:13px;min-width:0}.fcid-app [hidden]{display:none!important}.fcid-app h1,.fcid-app h2,.fcid-app h3{font-weight:600;line-height:1.3;margin:0}.fcid-app h2{font-size:16px}.fcid-app h3{font-size:13px}.fcid-app .card{padding:14px;min-width:0;border-radius:11px}.fcid-app .btn{color:#fff;font-size:12px;padding:7px 10px;font-weight:600;white-space:nowrap}.fcid-app .btn.small{padding:4px 7px}.fcid-app .btn.light{background:#edf1f6;border:1px solid #ccd6e3;color:#354d6d}.fcid-app .btn.blue{background:#2863d6;color:#fff}.muted,.small{font-size:12px;color:#53657d}.btnrow,.fc-hero-actions{display:flex;gap:7px;align-items:center;flex-wrap:wrap}.section-head{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:12px}.badge{font-size:12px;background:#edf2fa;color:#435a7b;border-radius:6px;padding:3px 7px;white-space:nowrap}
.fc-hero{padding:15px 18px;background:linear-gradient(110deg,#e8f3ff,#f0ebff 57%,#fff0e4);border-bottom:1px solid #ccd9ed}.fc-hero-inner{max-width:1900px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:14px}.fc-brand{font-size:12px;letter-spacing:1.5px;color:#4d5b89}.fc-hero h1{font-size:27px;margin-top:3px;letter-spacing:-.5px}.fc-heading{display:flex;align-items:center;gap:13px}#fcToggleSide{font-size:21px;padding:4px 10px;line-height:1.3}
.fcid-app .layout{max-width:1948px;margin:auto;display:grid;grid-template-columns:var(--side) minmax(0,1fr);gap:14px;padding:14px;align-items:start}.fcid-app .main{min-width:0;display:grid;grid-template-columns:minmax(0,1fr);gap:12px}.fcid-app .side{position:sticky;top:12px;max-height:calc(100vh - 24px);overflow:auto;scrollbar-width:thin;padding:13px;min-width:0}.side-collapsed .layout{grid-template-columns:minmax(0,1fr)}.side-collapsed .side{display:none}
.fc-source{position:relative;border:1px solid #dce7ed;background:#fbfdfe;border-radius:9px;padding:9px 8px;margin-top:8px}.fc-source-head{display:flex;align-items:center;justify-content:space-between;gap:5px;margin-bottom:7px}.fc-source h3{display:flex;align-items:center;gap:5px;min-width:0;font-size:12px;color:#24364b}.fc-source-name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.fc-source-bobj{flex:none;border:1px solid #d5e0e8;border-radius:4px;padding:1px 4px;color:#64748b;font-size:9px;font-weight:600}.fc-source-number{display:inline-flex;align-items:center;justify-content:center;flex:none;width:22px;height:22px;background:#376bd2;color:#fff;border-radius:6px;font-size:12px;font-weight:600}.fc-source[data-source=ctl] .fc-source-number{background:#8152bf}.fc-source[data-source=ctlNsr] .fc-source-number{background:#a85c07}.fc-source[data-source=nsr] .fc-source-number{background:#bf3d74}.fc-source[data-source=nce] .fc-source-number{background:#087e70}.fc-source-state{display:inline-flex;align-items:center;gap:5px;flex:none;max-width:105px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#53657d;font-size:12px;line-height:1.5;font-variant-numeric:tabular-nums}.fc-source-state.has-range{max-width:none}.fc-source-state::before{content:"";display:inline-block;flex:none;width:5px;height:5px;border-radius:50%;background:#94a3b8}.fc-source-state.ok{color:#166534}.fc-source-state.ok::before{background:#158653}.fc-source-state.stale{color:#d97706}.fc-source-state.stale::before{background:#d97706}.fc-source-state.loading{color:#92400e}.fc-source-state.loading::before{background:#d97706}.fc-source-state.error{cursor:help}.fc-source-toolbar{display:flex;align-items:center;gap:6px;min-width:0}.fc-source-toolbar>.btn{flex:none}.fc-source-file{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#53657d;font-size:12px;line-height:1.5}.fc-source-help{position:relative}.fc-source-help summary{list-style:none}.fc-source-help-bubble{position:absolute;right:0;top:30px;width:235px;z-index:30;padding:10px;background:#fff;border:1px solid #c4d0e4;border-radius:8px;box-shadow:0 8px 26px #23395630;font-size:12px}.fc-source-help a,.fc-nsr-help a{color:#265dbb}.fc-note{font-size:12px;color:#a73047;white-space:pre-line}.fc-note:empty{display:none}
.filter-head{display:flex;align-items:center;justify-content:space-between;margin:17px 0 10px}.fc-filter{margin-top:9px}.filter-head h2{font-size:18px;color:#30445f}.fc-filter[data-filter-needs=common]{--filter-ink:#52647d;--filter-tint:#f4f6fa;--filter-line:#cbd5e1}.fc-filter[data-filter-needs^=fid]{--filter-ink:var(--fid-ink);--filter-tint:var(--fid-tint);--filter-line:var(--fid-line)}.fc-filter[data-filter-needs=nsr]{--filter-ink:var(--nsr-ink);--filter-tint:var(--nsr-tint);--filter-line:var(--nsr-line)}.fc-filter label{display:block;font-size:12px;font-weight:600;color:var(--filter-ink,#435775);margin-bottom:4px}.fc-filter select{font-size:12px;width:100%;min-height:32px;border:1px solid var(--filter-line,#cbd6e6);border-radius:6px;padding:5px;background:#fff;color:var(--filter-ink,#263e60)}.fc-filter .chosen-container,.fc-filter .chosen-container-multi .chosen-choices li.search-field input.chosen-search-input{font-size:12px}.fc-filter .chosen-choices{border:1px solid var(--filter-line,#cbd6e6);border-radius:6px;padding:2px;background:#fff;box-shadow:none;min-height:32px}.fc-filter .chosen-choices li.search-choice{color:var(--filter-ink,#294c88);background:var(--filter-tint,#e8effd);border-color:var(--filter-line,#bacdee);padding:4px 22px 4px 6px}.fc-filter .chosen-results li{font-size:12px;line-height:1.4;padding:5px 7px}.fc-filter .chosen-results li.highlighted{background:var(--filter-ink,#2f64ca);color:#fff}.fc-filter input::placeholder{color:#526782}.fc-filter-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:4px}.fc-filter-head>label{min-width:0;margin-bottom:0}.fc-filter-actions{display:flex;flex:none;gap:6px;justify-content:flex-end}.fc-filter-actions button{border:0;background:none;color:var(--filter-ink,#365c97);font-size:12px;padding:1px 2px;cursor:pointer}.fcid-app .filter-group-title{font-size:16px;font-weight:600;color:var(--fid-ink);border-top:1px solid #dde5f1;margin-top:18px;padding-top:13px}.filter-group-title[data-filter-needs=nsr]{color:var(--nsr-ink);border-top-color:var(--nsr-line)}.fc-filter.primary-filter{border-left:3px solid var(--filter-line,#83a8e8);padding:7px 8px;background:var(--filter-tint,#f1f5fd);border-radius:0 6px 6px 0}.fc-filter.primary-filter .fc-filter-head>label{color:var(--filter-ink,#2f5590)}.quick-search{width:100%;border:1px solid #cbd6e6;border-radius:6px;padding:7px;color:#243b5c;font-size:13px}.quick-search{resize:vertical;min-height:52px}.nsr-year-picks{display:flex;gap:5px;flex-wrap:wrap;margin-top:6px}.nsr-year-picks button{font-size:12px;border:1px solid var(--filter-line,#cbd8e9);border-radius:5px;color:var(--filter-ink,#405a80);background:var(--filter-tint,#f4f7fc);padding:2px 6px;cursor:pointer}.nsr-year-picks button.active{background:var(--filter-ink,#315b96);border-color:var(--filter-ink,#315b96);color:#fff}.nsr-year-picks button.partial{background:#eaf3fc;border-color:#8db3d9;color:#285982}
.fid-lookup{display:flex;align-items:center;gap:7px}.fid-lookup textarea{flex:1;min-width:0;width:0}.fid-lookup .btn{flex:none;padding:6px 8px}
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
#fFidStatus_chosen .chosen-results .fid-status-detail{padding-inline-start:37px}
.workspace-nav{display:flex;gap:12px;flex-wrap:wrap;padding:6px;background:#fff;border:1px solid var(--line);border-radius:10px}.workspace-nav button{border:0;background:transparent;color:#485d7b;font-size:13px;font-weight:600;padding:8px 10px;border-radius:7px;cursor:pointer}.workspace-nav button.active{background:#e4edff;color:#1e52ad}.workspace-nav button:hover{background:#eef3fd}.workspace-nav button[data-view=rules]{margin-left:auto}.fc-view{display:grid;gap:12px;min-width:0}.filter-context{display:flex;align-items:center;flex-wrap:wrap;gap:7px 10px;padding:9px 11px;background:#fff;border:1px solid var(--line);border-radius:9px;font-size:12px}.context-title{flex:none;font-weight:600;color:#304d77}.filter-tags{display:flex;align-items:center;min-height:27px;flex:1 1 240px;min-width:0;flex-wrap:wrap;gap:6px}.filter-chip{display:inline-flex;align-items:center;gap:0;max-width:min(360px,100%);min-width:0;border:1px solid #c8d6ed;background:#f5f8ff;color:#31528a;border-radius:16px;padding:0;font-size:12px;line-height:1.4;cursor:pointer}.filter-chip:hover{background:#eaf1ff;border-color:#94b2df}.filter-chip-label{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;border:0;background:none;color:inherit;font:inherit;text-align:left;padding:4px 0 4px 9px;cursor:pointer;border-radius:16px 0 0 16px}.filter-chip-remove{flex:none;align-self:stretch;border:0;background:none;font-size:15px;line-height:1;color:#536d95;padding:0 8px 0 6px;cursor:pointer;border-radius:0 16px 16px 0}.filter-chip-remove:hover{background:#dfe9fb;color:#244c87}.scope-note{margin-left:auto;color:#53657d}
.workspace-nav [data-view=portfolio],#view-portfolio{--view-ink:var(--fid-ink);--view-tint:var(--fid-tint);--view-line:var(--fid-line)}
.workspace-nav [data-view=nsr],#view-nsr{--view-ink:var(--nsr-ink);--view-tint:var(--nsr-tint);--view-line:var(--nsr-line)}
.workspace-nav [data-view=changes],#view-changes{--view-ink:var(--ctl-ink);--view-tint:var(--ctl-tint);--view-line:var(--ctl-line)}
.workspace-nav [data-view=cycle],#view-cycle,#wfModal{--view-ink:#a85c07;--view-tint:#fff4e4;--view-line:#e6c79e}
.workspace-nav [data-view=summary],#view-summary{--view-ink:#3d805d;--view-tint:#edf7ef;--view-line:#c1dcc7}
.workspace-nav button:not([data-view=rules]){background:var(--view-tint);color:var(--view-ink);box-shadow:inset 0 0 0 1px var(--view-line)}.workspace-nav button:not([data-view=rules]):hover{box-shadow:inset 0 0 0 1px var(--view-ink)}.workspace-nav button:not([data-view=rules]).active{background:var(--view-ink);color:#fff;box-shadow:inset 0 0 0 1px var(--view-ink),0 2px 5px #253c6420}
.fc-view :is(.section-head,.section-title) h2,#wfModal .wf-head h1{color:var(--view-ink)}

.signal-strip{display:flex;gap:7px;flex-wrap:wrap}.signal{border:1px solid #d4deef;background:#fff;color:#3e5575;padding:6px 9px;border-radius:7px;font-size:12px;cursor:pointer}.signal b{margin-right:5px;color:#34539d;font-variant-numeric:tabular-nums}
.fid-risk-drivers .signal{border:0;border-radius:0;padding:8px 15px;background:#edf2fa;clip-path:polygon(10px 0,100% 0,100% calc(100% - 10px),calc(100% - 10px) 100%,0 100%,0 10px)}
.fid-risk-drivers .signal b{font:inherit;margin:0;color:inherit}.fid-risk-drivers .signal small{font-size:11px;opacity:.8;white-space:nowrap}
.fid-risk-drivers .signal:hover{background:#e5efff}
.fid-risk-drivers .signal:disabled{opacity:.55;cursor:default}
@media(max-width:760px){.fc-heading .title-tracking{display:none}}
.summary-heading.overview-tools{flex:1;flex-wrap:nowrap;justify-content:flex-start;max-width:100%}.summary-heading>h2{flex:none}.summary-heading .ignore-scope{margin-left:auto}
.overview-tools{display:flex;align-items:center;justify-content:flex-end;gap:8px;min-width:0}.overview-tools>.badge{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.ignore-scope{display:inline-flex;align-items:center;gap:4px;flex:none;font-size:12px;color:#758192;white-space:nowrap;cursor:pointer}.ignore-scope input{margin:0;accent-color:#758192}.section-head.overview-head{flex-wrap:nowrap}.overview-head>h2{flex:none}
.fid-risk-drivers .signal[aria-pressed=true]{background:#3566b0;color:#fff}
.fid-risk-drivers .signal:focus-visible{outline:0;box-shadow:inset 0 0 0 2px currentColor}
.fid-risk-drivers .signal[data-nsr-lifecycle="open"]{background:var(--nsr-tint);color:var(--nsr-ink)}
.fid-risk-drivers .signal[data-nsr-lifecycle="open"]:hover{background:#f8e0e6}
.fid-risk-drivers .signal[data-nsr-lifecycle="open"][aria-pressed=true]{background:var(--nsr-ink);color:#fff}
.fid-risk-drivers .signal[data-group-value="Pending Config CTL"][aria-pressed=true]{background:#745097;color:#fff}
.table-toolbar{display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;margin:10px 0 8px}.detail-head>.table-toolbar,summary>.table-toolbar{margin:0 0 0 auto;justify-content:flex-end}.detail-head>.table-toolbar{flex-wrap:nowrap;flex:0 1 auto}summary>.table-toolbar{float:right}.detail-head{flex-wrap:wrap}.table-count{font-size:12px;color:#4d6383}.fc-table-shell{min-width:0}.fc-table-wrap.tabulator{overflow:hidden;resize:vertical;min-height:105px;border:1px solid #d5dfed;border-radius:8px;scrollbar-width:thin;background:#fff;font:13px/20px var(--font-sans);color:#344c6b}.fid-dialog .fc-table-wrap.tabulator{min-height:0;resize:none}.fid-dialog .tabulator-tableholder{min-height:34px}.fc-table-wrap .tabulator-header{background:var(--table-heading-bg,#edf2fb);border-bottom:1px solid #d5dfed;color:#344c6b;font-weight:600}.fc-table-wrap .tabulator-header .tabulator-col{background:var(--table-heading-bg,#edf2fb);border-right:1px solid #e0e7f1;font-size:12px}.fc-table-wrap .tabulator-header .tabulator-col .tabulator-col-content{padding:7px 9px}.fc-table-wrap .tabulator-row{background:#fff;color:#344c6b;min-height:34px}.fc-table-wrap .tabulator-row.tabulator-row-even{background:#f8faff}.fc-table-wrap .tabulator-row:hover{background:#f0f5fc}.fc-table-wrap .tabulator-row .tabulator-cell{padding:6px 9px;height:34px;line-height:20px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border-right:1px solid #e8edf5;border-bottom:1px solid #e3e9f3}.fc-table-wrap .tabulator-row .tabulator-cell.tabulator-frozen{background:inherit;border-right:1px solid #d6e0ed}.fc-table-wrap .tabulator-tableholder{scrollbar-width:thin;overflow-anchor:none}.cell-content{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-height:22px}.fid-status-tags{display:inline-flex;gap:4px;vertical-align:top}#portfolioTable .tabulator-cell:has(.fid-status-tags){cursor:pointer}.fid-status-tags .pill{flex:none}.cell-content .pill,.cell-content .status-pill{max-width:100%;white-space:nowrap;vertical-align:top}.fc-link{border:0;background:none;color:#215fbb;text-decoration:underline;text-underline-offset:2px;cursor:pointer;padding:0;font-size:13px;white-space:nowrap}.pill{display:inline-block;font-size:12px;line-height:17px;padding:2px 6px;border-radius:5px;background:#edf2fa;color:#3e5a80}.pill.risk,.pill.alert{background:#ffe5eb;color:#a12d47}.pill.clear{background:#ddf4e7;color:#196747}.pill.unknown,.fid-detail-value.no-slot,.fid-evidence>.notice.no-slot{background:#fff1cf;color:#7c5514}.pill.info{background:#e7effb;color:#375b91}.fcid-app .status-pill{padding:2px 6px;line-height:17px}
.record-mark-toggle{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;margin:0;padding:0;border:1px solid #a9b3c0;border-radius:3px;background:transparent;color:#7d8a9d;font:600 13px/1 var(--font-sans);vertical-align:middle;cursor:pointer}.record-mark-toggle svg{display:block;width:12px;height:12px;flex:none;fill:none;stroke:currentColor;stroke-width:1.2}.record-mark-toggle[data-record-mark=ignore]{background:#edf0f4}.record-mark-toggle[data-record-mark=important]{border-color:#d99aa5;background:#fff1f3;color:#b8324a}.record-mark-toggle:focus-visible{outline:2px solid #557cae;outline-offset:2px}.fc-table-wrap .tabulator-row[data-record-mark=ignore] .tabulator-cell,.fc-table-wrap .tabulator-row[data-record-mark=ignore] .tabulator-cell *{color:#929aa5}.fc-table-wrap .tabulator-row[data-record-mark=important] .tabulator-cell,.fc-table-wrap .tabulator-row[data-record-mark=important] .tabulator-cell *{color:#b8324a}.pill.ignore{background:#edf0f4;color:#667184}
.fid-dialog .tabulator-row:is([data-fid-risk="pending-nsr"],[data-fid-risk="pending-config"]),.pill.pending,.fid-risk-drivers .signal[data-group-value="Pending Config CTL"]:not([aria-pressed=true]){background:#f2edfb;--fid-history-line:#aa83c6}
.pill.pending,.fid-risk-drivers .signal[data-group-value="Pending Config CTL"]:not([aria-pressed=true]){color:#745097}
.fc-table-wrap .tabulator-header-filter{display:none}.fc-table-wrap.column-filters .tabulator-header .tabulator-col.tabulator-sortable .tabulator-col-title{padding-right:29px}.fc-table-wrap.column-filters .tabulator-header .tabulator-col .tabulator-col-content .tabulator-col-sorter{right:0}.fc-table-wrap.column-filters .tabulator-col-sorter .tabulator-arrow{border-left-width:4px;border-right-width:4px}.column-heading>span{display:block;overflow:hidden;text-overflow:ellipsis}.column-funnel,.mark-filter-trigger{position:absolute;right:11px;top:50%;transform:translateY(-50%);display:inline-flex;align-items:center;justify-content:center;width:16px;height:18px;padding:2px;border:0;border-radius:3px;background:none;color:#8a97aa;cursor:pointer}.column-funnel svg,.mark-filter-trigger svg{width:12px;height:12px;fill:currentColor}.column-funnel:hover,.column-funnel.active,.mark-filter-trigger:hover,.mark-filter-trigger.active{background:#dceaff;color:#215fbb}.column-funnel:focus-visible,.mark-filter-trigger:focus-visible{outline:2px solid #557cae;outline-offset:-2px}
.column-filter-bar{display:flex;align-items:center;flex-wrap:wrap;gap:4px;flex:1 1 160px;min-width:0}.column-filter-bar:empty{display:none}.column-filter-bar .pill{max-width:min(220px,100%);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;border:0;cursor:pointer}
.column-filter-popup{position:fixed;position-area:bottom;position-try-fallbacks:flip-block;inset:auto;margin:8px;width:min(240px,calc(100vw - 16px));color:#344c6b;box-shadow:0 6px 22px #213a612b}.column-filter-popup label>span{font-weight:600}.column-filter-popup input{display:block;width:100%;margin:7px 0 10px;padding:7px;border:1px solid #c5d2e5;border-radius:5px}.column-filter-popup .btnrow{justify-content:space-between}
.mark-filter-trigger{position:static;transform:none;display:flex;margin:auto;height:20px}.column-filter-popup.mark-filter-popup{position-area:bottom span-right;justify-self:start;position-try-fallbacks:flip-block,flip-inline;margin:8px 0;padding:7px;width:190px}.mark-filter-popup label{display:flex;align-items:center;gap:9px;width:100%;padding:7px 9px;border-radius:5px;color:#344c6b;cursor:pointer}.mark-filter-popup label:first-child{border-bottom:1px solid #e0e7f1;margin-bottom:4px}.mark-filter-popup label:hover{background:#f0f5fc}.mark-filter-popup label:has(input:checked){background:#e7effb;color:#215fbb}.mark-filter-popup input[type=checkbox]{width:14px;height:14px;margin:0;padding:0;accent-color:#2863d6;flex:none}.mark-filter-popup label>span{flex:1;font-weight:400}.mark-filter-popup b{font-weight:400;font-variant-numeric:tabular-nums}
.fid-dialog .tabulator-row[data-fid-risk^="late-"],.fid-header-status .pill.risk{background:#fff2f5;--fid-history-line:#c75970}
.fid-dialog .tabulator-row:is([data-fid-risk="pending-nsr"],[data-fid-risk="pending-config"]):hover,.fid-risk-drivers .signal[data-group-value="Pending Config CTL"]:not([aria-pressed=true]):hover{background:#e8ddf5}
.fid-dialog .tabulator-row[data-fid-risk^="late-"]:hover{background:#ffe0e8}
.fid-dialog .tabulator-row[data-fid-risk="pending-pull-push"],.pill.neutral,.fid-risk-drivers .signal:is([data-group-value="Pending Pull/Push CTL"],[data-group-value="MCSD pending"],[data-group-value="NCE / FA signal"]):not([aria-pressed=true]){background:#f1f3f6;--fid-history-line:#a2abb8}
.pill.neutral,.fid-risk-drivers .signal:is([data-group-value="Pending Pull/Push CTL"],[data-group-value="MCSD pending"],[data-group-value="NCE / FA signal"]):not([aria-pressed=true]){color:#626e7c}
.fid-dialog .tabulator-row[data-fid-risk="pending-pull-push"]:hover,.fid-risk-drivers .signal:is([data-group-value="Pending Pull/Push CTL"],[data-group-value="MCSD pending"],[data-group-value="NCE / FA signal"]):not([aria-pressed=true]):hover{background:#e4e8ed}
.fid-risk-drivers .signal:is([data-group-value="Pending Pull/Push CTL"],[data-group-value="MCSD pending"],[data-group-value="NCE / FA signal"])[aria-pressed=true]{background:#697586;color:#fff}
.fid-header-status .pill:is(.pending,.risk,.neutral){border-left:3px solid var(--fid-history-line)}
.fid-dialog .tabulator-row[data-fid-risk]:not([data-fid-risk=""]) .tabulator-cell:first-child{background-image:linear-gradient(var(--fid-history-line),var(--fid-history-line));background-size:3px 100%;background-position:0 0;background-repeat:no-repeat}
.fid-dialog .tabulator-row.nsr-ctl-match{outline:2px solid #3b82f6;outline-offset:-2px}.fid-dialog .nsr-ctl-match .tabulator-frozen{box-shadow:inset 0 2px #3b82f6,inset 0 -2px #3b82f6}.fid-dialog .nsr-ctl-match .tabulator-cell:first-child{box-shadow:inset 2px 0 #3b82f6,inset 0 2px #3b82f6,inset 0 -2px #3b82f6}
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
.fc-nsr-dialog,.fid-dialog,.fc-note-dialog{border:1px solid #b7c7e0;border-radius:12px;padding:17px;background:#f6f8fe;color:#2e486c;box-shadow:0 16px 70px #1e355550}.fc-nsr-dialog{width:min(450px,94vw)}.fc-nsr-dialog::backdrop,.fid-dialog::backdrop,.fc-note-dialog::backdrop{background:#152e505c}.fc-range{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:14px}.fc-range input{width:100%;padding:8px;border:1px solid #c2d0e4;border-radius:6px;margin-top:5px}.fc-nsr-help{font-size:12px;margin:12px 0}.fc-actions{display:flex;justify-content:flex-end;gap:7px;margin-top:12px}.fid-dialog{width:min(1400px,96vw);height:min(860px,92vh);max-width:96vw;max-height:94vh;min-width:min(600px,94vw);min-height:300px;overflow:auto;resize:both}.fid-dialog>.section-head{position:sticky;top:-17px;background:#f6f8fe;z-index:10;padding:10px 0}.fid-dialog h1{display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-size:20px}.fid-action{display:inline-flex;align-items:center;padding:5px 9px;border:1px solid #c5d2e5;border-radius:7px;background:#fff;color:#215fbb;font-size:12px;line-height:1.3;text-decoration:none;white-space:nowrap;cursor:pointer}.fid-close{align-self:flex-start;margin:-8px -8px 0 0;width:36px;height:36px;padding:0;border:0;border-radius:7px;background:transparent;color:#526b8d;font-size:28px;line-height:1;cursor:pointer}.fid-action:hover,.fid-close:hover{background:#e7edf8}.fid-dialog-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}.fid-detail-value{padding:9px;background:#fff;border:1px solid #d6e0ef;border-radius:7px;font-size:13px;overflow-wrap:anywhere}.fid-detail-value small{display:block;font-size:12px;color:#526b8d;margin-bottom:3px}.fid-evidence{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:12px 0}.fid-evidence>.notice{display:block;min-width:0;overflow-wrap:anywhere}.fid-evidence>.notice.clear{border-color:#a8d5bb;border-left-color:#2d8b60;background:#eef9f2;color:#196747}.fid-evidence>.notice.unknown{border-color:#d2dcee;border-left-color:#91a2ba;background:#fff;color:#526b8d}.fid-evidence b{font-weight:600}.fid-evidence p{margin:7px 0 0}.fid-dialog h2{margin:0}.fid-dialog .table-toolbar{margin:16px 0 8px}.fc-toast{position:fixed;right:20px;bottom:20px;z-index:150;background:#244a79;color:#fff;padding:11px 14px;border-radius:9px;font-size:13px;max-width:min(480px,90vw);display:flex;align-items:center;gap:13px}.fc-toast button{color:#fff;border:1px solid #b7cceb;border-radius:5px;background:none;padding:4px 7px;cursor:pointer}.fcid-app .cycle-shell{display:grid;gap:12px}.fcid-app #cycleTab{display:block}.fcid-app #cycleTab:not(.active){display:none}.fcid-app .cycle-status,.fcid-app .cycle-label,.fcid-app .row-count,.fcid-app .context-badge{font-size:12px;color:#425c7f}body:has(#dashboardWorkspace.wf-open){overflow:hidden}
.inference-help{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;box-sizing:border-box;margin-left:5px;vertical-align:middle;border:1px solid #b6c5da;border-radius:50%;color:#6c809d;font-size:10px;font-weight:600;line-height:1;cursor:help}.inference-help::before{content:"?"}.inference-help:hover{color:#2458a6;border-color:#6d90c2;background:#edf3ff;outline:2px solid #c0d2ed;outline-offset:2px}
.inference-tooltip{position:fixed;inset:auto;margin:0;width:max-content;max-width:min(340px,calc(100vw - 16px));padding:9px 11px;border:1px solid #c5d2e5;border-radius:7px;background:#fff;color:#344c6b;box-shadow:0 3px 14px #253c6426;font:12px/1.5 var(--font-sans);pointer-events:none}
.detail-head{justify-content:flex-start;flex-wrap:wrap}.detail-head h2{white-space:nowrap}.fid-note-cell{display:block;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:left}.fid-note-add{display:inline-block;min-width:28px;border:1px solid #c5d2e5;border-radius:4px;background:#f5f8fd;line-height:18px;text-align:center;text-decoration:none}.fid-note-panel{margin:12px 0;border:1px solid #d2dcee;border-radius:8px;background:#fff}.fid-note-view{display:block;width:100%;padding:10px 12px;border:0;border-radius:7px;background:none;color:inherit;text-align:left;cursor:text}.fid-note-title{display:block;margin-bottom:6px;font-size:17px;font-weight:600}.fid-note-panel>form{padding:10px 12px}#fidNoteForm .section-head{margin-bottom:6px}.fid-note-text{display:block;white-space:pre-wrap;overflow-wrap:anywhere;font-size:13px;line-height:1.5}.fid-personal-mark{display:inline-flex;align-items:center}.fc-note-dialog{position:fixed;inset:auto;margin:0;width:min(520px,calc(100vw - 16px));max-width:calc(100vw - 16px);max-height:calc(100dvh - 16px);overflow:auto}#fidNoteForm textarea{width:100%;min-height:180px;resize:vertical;border:1px solid #c2d0e4;border-radius:7px;padding:9px;font:13px/1.5 var(--font-sans);color:#263e60;background:#fff}.fc-note-actions{display:flex;align-items:center;justify-content:flex-end;gap:10px;margin-top:10px}
.date-picks{margin:0 0 6px}.date-picks:empty{display:none}.fc-filter .fc-range{grid-template-columns:repeat(2,minmax(0,1fr));gap:7px;margin-top:0}.fc-filter .fc-range label{min-width:0;font-weight:400}.fc-filter .fc-range input{min-width:0;min-height:32px;padding:5px 6px;margin-top:0;background:#fff;color:var(--filter-ink,#263e60);border-color:var(--filter-line,#c2d0e4);font:inherit}
.summary-heading{display:flex;align-items:center;gap:10px}.summary-controls{display:grid;grid-template-columns:170px minmax(0,1fr);gap:14px;padding:12px;margin-bottom:6px;background:#f5f7fb;border-radius:9px}.summary-controls .fc-filter{margin:0}.summary-chart{overflow:auto}.summary-table{width:100%;table-layout:fixed;border-collapse:separate;border-spacing:0 10px;color:#344c6b}.summary-table th{font-size:12px;font-weight:600}.summary-table .fc-link{color:inherit;font:inherit;text-align:left;white-space:normal;text-decoration:none}.summary-table .fc-link:hover{text-decoration:underline}.summary-table thead th{padding:4px 10px 2px;color:#5b6d85;overflow-wrap:anywhere}.summary-table th:first-child{position:sticky;left:0;z-index:1;text-align:left}.summary-table thead th:first-child{background:#fff;z-index:2}.summary-table tbody tr{--metric-color:var(--fid-ink);--metric-bg:var(--fid-tint)}.summary-table tr[data-summary-source=nsr],.summary-table tr[data-summary-source=linked]{--metric-color:var(--nsr-ink);--metric-bg:var(--nsr-tint)}.summary-table tr[data-summary-source=ctl]{--metric-color:var(--ctl-ink);--metric-bg:var(--ctl-tint)}.summary-table tr:is([data-summary-metric=slotted],[data-summary-metric=stableSlot],[data-summary-metric=stablePost]){--metric-color:#3d805d;--metric-bg:#edf7ef}.summary-table tbody th,.summary-table td{background:var(--metric-bg);border-block:1px solid #e0e6ef}.summary-table tbody th{border-left:3px solid var(--metric-color);border-radius:9px 0 0 9px;padding:16px;color:#304967;font-size:13px;line-height:1.5}.summary-table td{padding:0}.summary-table td:last-child{border-right:1px solid #e0e6ef;border-radius:0 9px 9px 0}.summary-plot{position:relative;height:112px;font-variant-numeric:tabular-nums;text-align:center}.summary-bar{position:absolute;bottom:14px;left:calc(50% - 17px);width:34px;height:var(--bar);background:var(--metric-color);border-radius:4px 4px 0 0}.summary-value{position:absolute;top:10px;left:0;width:100%;font-size:14px;font-weight:600;color:#304967;white-space:nowrap}.summary-plot.unavailable .summary-value{color:#8794a5;font-weight:400}.summary-empty{padding:22px 12px;color:#61738c}.summary-legend{margin-top:6px;font-size:12px;color:#61738c}@media(max-width:760px){.summary-controls{grid-template-columns:minmax(0,1fr)}.summary-heading{flex-wrap:wrap}.summary-table tbody th{padding:12px}}
@container(max-width:700px){.nsr-stage-sum::before{display:none}.nsr-flow{grid-template-columns:16px minmax(0,1fr) 16px minmax(0,1fr)}.nsr-equation>[data-nsr-group=all]{grid-column:1/-1;grid-row:1;display:grid;grid-template-columns:1fr auto;align-items:center}.nsr-equation>[data-nsr-group=all] b{grid-column:2;grid-row:1/3;margin:0}.nsr-equation>.math-sign:nth-child(2){grid-column:1;grid-row:2}.nsr-equation>[data-nsr-group=completed]{grid-column:2;grid-row:2}.nsr-equation>.math-sign:nth-child(4){grid-column:3;grid-row:2}.nsr-equation>[data-nsr-group=workflow]{grid-column:4;grid-row:2}.nsr-equation>.math-sign:nth-child(6){grid-column:1;grid-row:3}.nsr-equation>[data-nsr-group=rejected]{grid-column:2;grid-row:3}.nsr-equation>.math-sign:nth-child(8){grid-column:3;grid-row:3}.nsr-equation>[data-nsr-group=cancelled]{grid-column:4;grid-row:3}.has-completed .nsr-equation>[data-nsr-group=completed]{grid-column:2/-1;display:grid;grid-template-columns:1fr auto;align-items:center}.has-completed .nsr-equation>[data-nsr-group=completed] b{grid-column:2;grid-row:1/3;margin:0}.nsr-completed-arrow{grid-column:2;grid-row:3;margin:0;height:22px}.nsr-completed-arrow::before{top:-7px}.nsr-breakdown-equals{grid-column:3;grid-row:3;margin:0}.nsr-stage-sum{grid-column:1/-1;grid-row:4;display:grid;grid-template-columns:16px minmax(0,1fr) 16px minmax(0,1fr);margin:0}.nsr-stage-spacer{display:block}.nsr-stage:only-of-type{grid-column:2/-1}.has-completed .nsr-equation>.math-sign:nth-child(4){grid-column:1;grid-row:5}.has-completed .nsr-equation>[data-nsr-group=workflow]{grid-column:2;grid-row:5}.has-completed .nsr-equation>.math-sign:nth-child(6){grid-column:3;grid-row:5}.has-completed .nsr-equation>[data-nsr-group=rejected]{grid-column:4;grid-row:5}.has-completed .nsr-equation>.math-sign:nth-child(8){grid-column:1;grid-row:6}.has-completed .nsr-equation>[data-nsr-group=cancelled]{grid-column:2;grid-row:6}}

@media(max-width:1250px){body.fcid-app{--side:360px}.nsr-pivots{grid-template-columns:repeat(2,minmax(0,1fr))}.fc-hero h1{font-size:25px}}@media(max-width:1000px){.marked-settings{grid-template-columns:minmax(0,1fr)}.ag-grid,.fid-dialog-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:760px){.fcid-app .layout{grid-template-columns:minmax(0,1fr);padding:10px;gap:10px}.fcid-app .side{position:static;max-height:none}.fc-hero{padding:12px}.fc-hero h1{font-size:21px}.fc-hero-inner{align-items:flex-start;flex-wrap:wrap}.fc-hero-actions{margin-left:auto}.fcid-app .nsr-pivots,.ag-grid,.fid-evidence{grid-template-columns:minmax(0,1fr)}.workspace-nav button{padding:7px}.scope-note{margin-left:0}.fid-dialog>.section-head{flex-wrap:nowrap}.fid-header-status{flex-basis:100%}.section-head{flex-wrap:wrap}}
.fid-quarter-state{display:flex;align-items:center;flex-wrap:wrap;gap:4px 8px;margin-top:4px;min-width:0}.fid-quarter-state>span:first-child{font-weight:400}.fid-quarter-state.is-current>span:first-child{font-weight:600}.fid-date-value{display:grid;grid-template-rows:subgrid;grid-row:span var(--fid-date-rows);row-gap:4px}.fid-date-value>small,.fid-date-value>.fid-quarter-state{margin:0}
`;
const sourceHtml=Object.entries(sources).map(([key,source],index)=>{
  const importLabel=key==='ctlNsr'?'Append':'Import';
  const help=BOBJ_SOURCES[key]?`Open SAPGUIcontrol.exe. <a href="${SAPGUICONTROL_DOWNLOAD_URL}" download="SAPGUIcontrol.7z" target="_blank" rel="noopener noreferrer">Download SAPGUIcontrol</a>.`:key==='ctlNsr'?'Download from the CTL system: <a href="https://pcazcrpap02.fremont.lamrc.net:8443/sap(bD1lbiZjPTI1MCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm" target="_blank" rel="noopener noreferrer">Open download page</a>.':key==='nce'?'Ask Stacey for the latest NCE / FA Review file.':'';
  return `<section class="fc-source" data-source="${key}" aria-labelledby="fcSourceTitle-${key}"><div class="fc-source-head"><h3 id="fcSourceTitle-${key}"><span class="fc-source-number">${index+1}</span><span class="fc-source-name">${escapeHtml(source.label)}</span>${BOBJ_SOURCES[key]?'<small class="fc-source-bobj">BOBJ</small>':''}</h3><span id="fcSource-${key}" class="fc-source-state" role="status">Unavailable</span></div><div class="fc-source-toolbar"><span class="fc-source-file" id="fcMeta-${key}"></span>${BOBJ_SOURCES[key]?`<button class="btn small" data-source-action="${key==='fcid'?'bobj-update':'ctl-bobj-update'}">Update</button>`:key==='nsr'?'<button class="btn small light" data-source-action="nsr-open">Load NSR</button>':''}<button id="fcImport-${key}" class="btn small light" data-import-source="${key}">${importLabel}</button>${help?`<details class="fc-source-help" id="fcHelp-${key}"${BOBJ_SOURCES[key]?' hidden':''}><summary class="btn light small" aria-label="Help for ${escapeHtml(source.label)}">...</summary><div class="fc-source-help-bubble">${help}</div></details>`:''}</div><input type="file" id="fcFile-${key}" data-source-file="${key}" aria-label="${importLabel} ${escapeHtml(source.label)}" accept=".xlsx,.xls,.xlsm,.csv${key==='nsr'?',.json':''}" hidden></section>`;
}).join('');
const filterHtml=([name,label,path,needs])=>dateFilters.includes(path)?`<div class="fc-filter primary-filter" data-filter-needs="${needs}"><div class="fc-filter-head"><label for="${name.slice(0,-4)}From">${label}</label><div class="fc-filter-actions"><button type="button" data-clear-filter="${name}">Clear</button></div></div><div id="${name.slice(0,-4)}${needs==='ctl'?'Periods':'Years'}" class="nsr-year-picks date-picks"></div><div id="${name}" class="fc-range" role="group" aria-label="${label}"><label>From<input id="${name.slice(0,-4)}From" type="date" aria-label="${label} from"></label><label>To<input id="${name.slice(0,-4)}To" type="date" aria-label="${label} to"></label></div></div>`:`<div class="fc-filter${['fAccount','fQuarter','fOwner','fSlotInitiator','fNsrInitiator','fCtlStage','fCtlInitiator'].includes(name)?' primary-filter':''}" data-filter-needs="${needs}"><div class="fc-filter-head"><label for="${name}" title="${path==='owner'?escapeHtml(ownerHint):path.startsWith('nsr.')?'All selected NSR conditions must match the same NSR':''}">${label}${path==='quarter'?inferenceHelp(quarterHint):path==='owner'?inferenceHelp(ownerHint):path==='salesAdmin'?inferenceHelp(salesAdminHint):''}</label><div class="fc-filter-actions">${name==='fAccount'?'<button type="button" data-mapping-rule title="Fab / Account Mapping">Rule</button>':''}${name==='fQuarter'?'<button type="button" data-reset-quarter>Reset</button>':''}<button type="button" data-select-all="${name}">All</button><button type="button" data-clear-filter="${name}">Clear</button></div></div><select id="${name}" multiple aria-label="${label}"></select></div>`;
const views=[['portfolio','FID Portfolio'],['changes','Slot / CTL'],['nsr','NSR List'],['cycle','NSR Cycle Time'],['summary','Summary Metrics'],['rules','Rules & Settings']];

const charset=document.createElement('meta');charset.charset='utf-8';const viewport=document.createElement('meta');viewport.name='viewport';viewport.content='width=device-width,initial-scale=1';const title=document.createElement('title');title.textContent=FID_DASHBOARD_TITLE.replace(/^🏭\s*/,'');const style=document.createElement('style');style.textContent=css;
document.head.replaceChildren(charset,viewport,title,emojiFavicon('🏭'),style);document.body.className='fcid-app';
document.body.innerHTML=`<header class="fc-hero"><div class="fc-hero-inner"><div class="fc-heading"><button id="fcToggleSide" class="btn light" aria-label="Toggle Data Sources and Filters" aria-expanded="true" title="Data Sources / Filters">☰</button><div><div class="fc-brand">TOOLKIT</div><h1>${escapeHtml(FID_DASHBOARD_TITLE).replace(' Tracking','<span class="title-tracking"> Tracking</span>')}</h1></div></div><div class="fc-hero-actions"><button id="fcRefresh" class="btn light" disabled>Refresh</button><a class="btn light" href="${GUIDE_URL}" target="_blank" rel="noopener noreferrer">Help</a></div></div></header>
<div class="layout" id="dashboardWorkspace"><aside class="side card"><div class="section-head"><h2>Data Sources</h2><span class="badge" id="fcSourceCount">0 / 5</span></div>${sourceHtml}<div id="fcActionStatus" class="fc-note" role="status"></div>
<div class="filter-head"><h2>Filters</h2><button class="btn small light" id="fcReset" title="Restore default filters">Reset</button></div>
<form id="fidLookup" class="fc-filter fid-lookup"><textarea id="fSearch" class="quick-search" placeholder="Search multiple FIDs, CTLs, or NSRs&#10;Single FID/CTL enables FID 360" aria-label="Search multiple FIDs, CTLs, or NSRs&#10;Single FID/CTL enables FID 360" rows="2"></textarea><button class="btn light" type="submit" aria-label="Open FID 360°" title="Open FID 360°" hidden><span aria-hidden="true">🔍</span></button></form>
${['fAccount'].map(name=>filterHtml(filterDefs.find(def=>def[0]===name))).join('')}<h3 class="filter-group-title" data-filter-needs="fid">FID Filters</h3>${['fQuarter','fOwner','fSlotInitiator','fFidStatus','fRisk','fCustomer','fSystem','fBusinessUnit','fSalesAdmin'].map(name=>filterHtml(filterDefs.find(def=>def[0]===name))).join('')}<h3 class="filter-group-title" data-filter-needs="ctl">CTL Filters</h3>${['fCtlStage','fCtlInitiator','fCtlCreatedDate','fCtlChange','fCtlCustomer'].map(name=>filterHtml(filterDefs.find(def=>def[0]===name))).join('')}<h3 class="filter-group-title" data-filter-needs="nsr">NSR Filters</h3>${['fNsrInitiator','fNsrSubmitDate','fNsrTaskStatus','fNsrType','fNsrOrg','fNsrSubsystem','fNsrProduct','fNsrCustomer'].map(name=>filterHtml(filterDefs.find(def=>def[0]===name))).join('')}</aside>
<main class="main"><nav class="workspace-nav" id="workspaceNav" aria-label="Dashboard views">${views.map(([view,label])=>`<button type="button" data-view="${view}" aria-selected="${view==='nsr'}"${view==='nsr'?' class="active"':''}${['nsr','cycle','rules'].includes(view)?'':' hidden'}>${label}</button>`).join('')}</nav><div id="filterContext" class="filter-context"></div>${views.map(([view])=>`<section id="view-${view}" class="fc-view"${view==='nsr'?'':' hidden'}></section>`).join('')}</main></div>
<dialog id="fcNsrDialog" class="fc-nsr-dialog" aria-labelledby="fcNsrDialogTitle"><form id="fcNsrForm"><h2 id="fcNsrDialogTitle">Load NSR</h2><div class="fc-range"><label>From<input id="fcNsrFrom" type="date" required></label><label>To<input id="fcNsrTo" type="date" required></label></div><p class="fc-nsr-help"><a href="https://fep.lamresearch.com/flp#ZNSR-display&amp;/Search/Search/%20" target="_blank" rel="noopener noreferrer">Open NSR search</a></p><div class="fc-actions"><button id="fcNsrCancel" class="btn light" type="button">Cancel</button><button id="fcNsrLoad" class="btn blue" type="submit">Load</button></div></form></dialog><dialog id="fidDetail" class="fid-dialog" aria-labelledby="fidDetailTitle"></dialog><dialog id="fidNoteDialog" class="fc-note-dialog" aria-labelledby="fidNoteTitle"></dialog><form id="columnFilterPopup" class="card column-filter-popup" popover role="dialog" aria-label="Column filter"><label><span></span><input autocomplete="off" autofocus></label><div class="btnrow"><button type="reset" class="fc-link">Clear</button><button type="submit" class="btn blue small">Apply</button></div></form><div id="markFilterPopup" class="card column-filter-popup mark-filter-popup" popover role="dialog" aria-label="Mark filter"><div role="group" aria-label="Mark status"></div></div><div id="fcToast" class="fc-toast" role="status" hidden></div>`;
const cycleTemplate=document.createElement('template');cycleTemplate.innerHTML=NSR_CYCLE_HTML;id('view-cycle').append(cycleTemplate.content.querySelector('#cycleTab'));id('dashboardWorkspace').append(cycleTemplate.content.querySelector('#wfModal'));
let selectedView='nsr',graph=null,current=null,revision=0,modelRevision=0,toastTimer=0,searchTimer=0,sidebarAnimation=null;
const rendered=new Map(),tables=new Map(),filterValues={},localTables=new Map();let mappingDraft=null;const markDrafts={nsr:null,fid:null};
const pretty=(value,path='')=>path==='nsr.taskStatus'?(nsrCompletedTasks[value]||value||'[Unknown]'):path==='ctl.stage'?ctlStageName(value)||'[Unknown]':path==='ctl.change'&&value==='[Unknown]'?'No CTL Yet':path==='fidStatus'?(value==='stable-slot|stable-ctl'?fidStatusNames['risk-free']:value.startsWith('cause:')?fidCauseNames[value.slice(6)]||value:fidCellNames[value]||fidStatusNames[value]||value):path==='riskType'?(value==='MCSD pending'?'Slot + no MCSD':value):value||'[Unknown]';
const number=value=>String(Number(value||0));
const pill=(value,kind='info',tip='')=>`<span class="pill ${kind}"${tip?` title="${escapeHtml(tip)}"`:''}>${escapeHtml(value||'—')}</span>`;
const fidLink=value=>value?`<button type="button" class="fc-link" data-fid="${escapeHtml(value)}">${escapeHtml(value)}</button>`:'—';
const quarterYear=value=>nsrRecords.periodYear(value)||String(value).match(/'(\d{2})$/)?.[1]?.replace(/^(\d{2})$/,'20$1')||'';
function defaultQuarters(){
  const current=fidQuarterFromDate(new Date()),first=Number(current.slice(-2))*4+['Mar','Jun','Sep','Dec'].indexOf(current.slice(1,4));
  return Array.from({length:5},(_,offset)=>{const quarter=first+offset;return ['QMar','QJun','QSep','QDec'][quarter%4]+"'"+String(Math.floor(quarter/4)).slice(-2).padStart(2,'0')});
}
const setPath=(value,path,next)=>{const keys=path.split('.'),last=keys.pop();let target=value;keys.forEach(key=>target=target[key]??={});target[last]=next;};
function toast(message){clearTimeout(toastTimer);id('fcToast').innerHTML=`<span>${escapeHtml(message)}</span><button type="button" data-toast-close aria-label="Dismiss">×</button>`;id('fcToast').hidden=false;toastTimer=setTimeout(()=>id('fcToast').hidden=true,4500)}
function buildFilters(paths){
  const nodes=graph.population,statNsrs=graph.nsrs.filter(nsr=>nsr.eligible),statCtls=graph.ctls;let observations;
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
  const riskOrder=['MCSD pending','NCE / FA signal'];
  for(const [name,,path] of filterDefs){
    if(paths&&!paths.includes(path))continue;
    if(dateFilters.includes(path)){
      if(!(path==='nsr.submitted'?statNsrs:statCtls).length)delete filterValues[path];const [from='',to='']=filterValues[path]||[],[start,end]=id(name).querySelectorAll('input');start.value=from;end.value=to;start.max=to;end.min=from;start.setCustomValidity('');end.setCustomValidity('');
      id(name).previousElementSibling.innerHTML=path==='ctl.created'?['This week','Last week','This month','Last month'].map(period=>{
        const range=ctlDateRange(period),active=from===range[0]&&to===range[1];return `<button type="button" data-ctl-period="${period}" class="${active?'active':''}" aria-pressed="${active}" title="${range.join(' → ')}">${period}</button>`;
      }).join(''):[...document.querySelectorAll('[data-fid-year]')].map(button=>{
        const year=button.dataset.fidYear,active=from===year+'-01-01'&&to===year+'-12-31';
        return `<button type="button" data-nsr-year="${year}" class="${active?'active':''}" aria-pressed="${active}" title="Select January 1 through December 31, ${year}">${year}</button>`;
      }).join('');continue;
    }
    const values=path==='fidStatus'?[]:path==='quarter'?assessments.map(state=>state.quarter):path==='account'?[...nodes.map(node=>node.account),...statNsrs.flatMap(nsr=>nsr.accounts.length?nsr.accounts:[''])]:path==='riskType'?nodes.flatMap(node=>node.risks.map(risk=>risk.label)).filter(value=>riskOrder.includes(value)):path.startsWith('ctl.')?statCtls.map(ctl=>ctl[path.slice(4)]):path.startsWith('nsr.')?(observations??=statNsrs.flatMap(nsr=>nsr.observations.length?nsr.observations:[nsr])).map(observation=>observation[path.slice(4)]):path==='slot.initiator'?nodes.map(node=>node.slot.initiator):nodes.map(node=>node[path]);
    const quarters=path==='quarter'?[...(store.raw.fcid.length?defaultQuarters():[]),...(filterValues.quarter||[])]:path==='ctl.change'&&(store.meta.ctl||store.meta.ctlNsr)?[...defaultCtlTypes,...Object.values(pendingCtlTypes).flat()]:[];
    const available=path==='fidStatus'?fidStatuses:[...new Set([...values.map(value=>value||'[Unknown]'),...quarters])].sort((a,b)=>Number(a==='[Unknown]')-Number(b==='[Unknown]')||(path==='nsr.taskStatus'?nsrRecords.compareStatus(a,b):path==='ctl.stage'?compareCtlStages(a,b):path==='riskType'?riskOrder.indexOf(a)-riskOrder.indexOf(b):path.toLowerCase().includes('quarter')?nsrRecords.periodRank(a)-nsrRecords.periodRank(b)||a.localeCompare(b):a.localeCompare(b,undefined,{numeric:true})));
    if(path==='nsr.taskStatus'&&available.includes('15-Completed'))available.splice(available.indexOf('15-Completed')+1,0,...Object.keys(nsrCompletedTasks));
    const select=id(name),saved=path==='fidStatus'?normalizeFidStatus(filterValues[path]||[]):filterValues[path]||[];
    const option=value=>{
      const depth=fidDepth.get(value);
      const classes=path==='fidStatus'
        ? [depth===1?'fid-status-child':depth===2?'fid-status-detail':'',['no-slot','slot-ready'].includes(value)?'fid-status-parent':''].filter(Boolean).join(' ')
        : path==='nsr.taskStatus'&&Object.hasOwn(nsrCompletedTasks,value)?'group-option':'';
      const label=path==='fidStatus'&&['cause:slot-pending','cause:slot-dropped'].includes(value)?pretty(value,path).replace(/^(?:No Slot|Slotted FID) · /,''):pretty(value,path);
      return `<option value="${escapeHtml(value)}"${classes?` class="${classes}"`:''}${saved.includes(value)?' selected':''}>${escapeHtml(label)}</option>`;
    };
    select.innerHTML=available.map(option).join('');
    filterValues[path]=saved.filter(value=>available.includes(value));jQuery(select).trigger('chosen:updated');
    if(name==='fQuarter'){
      const currentYear=new Date().getFullYear();let years=select.parentElement.querySelector('.nsr-year-picks');
      if(!years){years=document.createElement('div');years.className='nsr-year-picks';select.after(years)}
      years.innerHTML=[...new Set(values.map(quarterYear).filter(year=>year>=currentYear-1&&year<=currentYear+3))].sort().map(year=>`<button type="button" data-fid-year="${year}" title="Select / clear all quarters in ${year}">${year}</button>`).join('');
    }
  }
}
function syncFilters(){
  const fid=store.raw.fcid.length>0,nsr=store.raw.nsr.length>0,ctl=store.raw.ctl.length>0||store.raw.ctlNsr.length>0;
  const needs={common:fid||nsr,fid,fidCtl:fid&&ctl,fidOwner:fid&&ctl&&graph.population.some(node=>node.owner),fidSlot:fid&&ctl&&graph.population.some(node=>node.slot.initiator),fidSalesAdmin:fid&&ctl&&graph.population.some(node=>node.salesAdmin),fidRisk:fid&&graph.population.some(node=>node.risks.some(risk=>['MCSD pending','NCE / FA signal'].includes(risk.label))),ctl,nsr};document.querySelectorAll('[data-filter-needs]').forEach(group=>group.hidden=!needs[group.dataset.filterNeeds]);
}
function syncFidLookup(){
  const tokens=id('fSearch').value.trim().split(/[\s,;|/]+/).filter(Boolean),value=tokens.length===1&&/^\d{5,10}$/.test(tokens[0])?tokens[0]:'';
  const fid=value&&(graph.fids.has(value)?value:graph.ctls.find(ctl=>ctl.number===value)?.fidId||'');
  id('fidLookup').querySelector('button').hidden=!fid;return fid;
}
function recalculate(updateMarks=false){
  const started=performance.now();graph=createFidGraph(store.raw,{fidMarks:currentMarks('fid'),accountMap:accountMapping.match,ignoredNsrs,importantNsrs,loadedSources:Object.fromEntries(Object.keys(store.meta).map(key=>[key,true])),today:nsrData.localDate(new Date())});modelRevision++;
  buildFilters();syncFilters();applyFilters(updateMarks);store.dirty=false;store.buildMs=performance.now()-started;
}
function queryFilters(){const filters={search:id('fSearch').value,fidStatus:[],signal:store.raw.fcid.length?filterValues.signal||[]:[],includeIgnored};for(const [name,,path] of filterDefs)if(!id(name).closest('.fc-filter').hidden&&filterValues[path]?.length)setPath(filters,path,dateFilters.includes(path)?{from:filterValues[path][0],to:filterValues[path][1]}:filterValues[path]);return filters}
function applyFilters(updateMarks=false,changedNsr=''){
  syncFidLookup();
  const detailFid=id('fidDetail').open?id('fidDetail').dataset.fid:'',view=id('view-'+selectedView),anchor=changedNsr&&!detailFid&&selectedView==='nsr'?id('nsrTable'):view,position=updateMarks?anchor?.getBoundingClientRect():null;
  current=graph.query(queryFilters());revision++;
  for(const [name,state] of tables){if(name==='cycleTable'||name==='wfTable'||updateMarks&&(changedNsr||state.grid.element.closest('.fc-view')?.hidden===false||state.grid.element.closest('.fid-dialog')?.open))continue;tableResize.unobserve(state.grid.element);state.grid.destroy();tables.delete(name)}
  rendered.clear();paintContext();showView(selectedView,updateMarks,changedNsr);if(detailFid)showFid(detailFid,updateMarks,changedNsr);
  // Keep visible content in place when filter chips or NSR Overview cards change height.
  if(position?.top<innerHeight&&position.bottom>0)window.scrollBy({top:anchor.getBoundingClientRect().top-position.top,behavior:'instant'});
}
function paintContext(){
  document.querySelectorAll('[data-exclude-ignored]').forEach(input=>input.checked=!includeIgnored);
  document.querySelectorAll('[data-fid-year]').forEach(button=>{const values=[...id('fQuarter').options].map(option=>option.value).filter(value=>quarterYear(value)===button.dataset.fidYear),selected=filterValues.quarter||[],count=values.filter(value=>selected.includes(value)).length;button.classList.toggle('active',count>0&&count===values.length);button.classList.toggle('partial',count>0&&count<values.length);button.setAttribute('aria-pressed',String(count>0&&count===values.length))});
  id('filterContext').hidden=!Object.keys(store.meta).length;
  const chip=(value,target,clear,label)=>{
    const reset=target==='fQuarter'||target==='fCtlChange',action=target==='fQuarter'?'data-reset-quarter title="Reset to the default five quarters"':target==='fCtlChange'?'data-reset-ctl-type title="Reset to Config, Slot, Drop Slot and No CTL Yet"':clear;
    return `<span class="filter-chip" title="${escapeHtml(value)}"><button class="filter-chip-label" type="button" data-locate-filter="${target}" aria-label="Locate ${escapeHtml(label)} filter">${escapeHtml(value)}</button><button class="filter-chip-remove" type="button" ${action} aria-label="${reset?'Reset':'Remove'} ${escapeHtml(label)} filter">${reset?'↺':'×'}</button></span>`;
  };
  const chips=filterDefs.filter(([name,,path])=>!id(name).closest('.fc-filter').hidden&&filterValues[path]?.length).map(([name,label,path])=>chip(label+': '+(dateFilters.includes(path)?(filterValues[path][0]&&filterValues[path][1]?filterValues[path].join(' → '):filterValues[path][0]?'From '+filterValues[path][0]:'Through '+filterValues[path][1]):filterValues[path].map(value=>pretty(value,path)).join(' or ')),name,`data-clear-filter="${name}"`,label));
  if(store.raw.fcid.length&&filterValues.signal?.length)chips.push(chip('Signal: '+filterValues.signal.map(value=>fidStatusNames[value]||(value==='nce'?'NCE / FA':'Pending Config CTL')).join(' or '),'fRisk','data-clear-signal','signal'));
  const search=id('fSearch').value.trim();if(search)chips.push(chip('Search: '+search,'fSearch','data-clear-search','search'));
  id('filterContext').innerHTML=`<span class="context-title">Filtered By:</span><div class="filter-tags">${chips.join('')||'<span class="muted">All values</span>'}</div><span class="scope-note">Commit portfolio: ${number(current.fids.length)} FIDs · NSR List: ${number(current.nsrs.length)}</span>`;
}
function showView(name,updateMarks=false,changedNsr=''){
  selectedView=name;
  id('workspaceNav').querySelectorAll('[data-view]').forEach(button=>{const active=button.dataset.view===name;button.classList.toggle('active',active);button.setAttribute('aria-selected',String(active))});
  document.querySelectorAll('.fc-view').forEach(element=>element.hidden=element.id!=='view-'+name);id('cycleTab').classList.toggle('active',name==='cycle');
  if(name==='cycle'){cycle.show();return}
  if(rendered.get(name)===revision)return;
  ({portfolio:renderPortfolio,changes:renderChanges,nsr:renderNsrs,summary:renderSummaryMetrics,rules:renderRules}[name])(id('view-'+name),updateMarks,changedNsr);rendered.set(name,revision);
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
function selectFidStatus(value){
  const parts=fidSelection.groups[value]||[value],selected=filterValues.fidStatus||[];
  if(fidSelection.groups[value]){changeFilter('fidStatus',fidStatusSelected(value)?[]:parts);return}
  changeFilter('fidStatus',parts.every(part=>selected.includes(part))
    ?selected.filter(item=>!parts.includes(item)):[...selected,...parts]);
}
function changeFilter(path,values){
  let changed=false;
  for(const [key,next] of typeof path==='string'?[[path,values]]:Object.entries(path)){
    const selected=key==='fidStatus'?normalizeFidStatus(next):key==='nsr.taskStatus'?normalizeNsrTasks(next):next,previous=filterValues[key]||[];changed||=selected.length!==previous.length||selected.some((value,index)=>value!==previous[index]);filterValues[key]=selected;const definition=filterDefs.find(def=>def[2]===key);
    if(definition){if(dateFilters.includes(key))buildFilters([key]);else{const select=id(definition[0]);for(const option of select.options)option.selected=selected.includes(option.value);jQuery(select).trigger('chosen:updated')}}
  }
  if(changed)applyFilters(true);
}
function positionFilterChoices(select){
  const chosen=id(select.id+'_chosen');if(!chosen?.classList.contains('chosen-with-drop'))return;
  const bounds=chosen.closest('.side').getBoundingClientRect(),box=chosen.getBoundingClientRect(),height=chosen.querySelector('.chosen-drop').getBoundingClientRect().height;
  const below=Math.min(innerHeight,bounds.bottom)-box.bottom,above=box.top-Math.max(0,bounds.top);chosen.classList.toggle('chosen-dropup',height>below&&above>below);
}
const portfolioColumns=[markColumn('fid'),
  ['FID',node=>node.id,fidLink],['FID Status & Risks',node=>fidDetailStatus(node,false).map(([label])=>label).join(' · '),(_,node)=>`<span class="fid-status-tags">${fidDetailStatus(node).map(([label,kind,tip])=>pill(label,kind,tip)).join('')}</span>`],
  ['Quarter',node=>fidQuarterRows(node).map(state=>state.quarter||'[Unknown]').join(', ')],['System Description',node=>node.system],['CRD',node=>node.crd],['SRD',node=>fidQuarterRows(node).map(state=>state.dates.rd||'—').join(', ')],['MCSD',node=>fidQuarterRows(node).map(state=>state.dates.mcsd||'—').join(', ')],
  ['Account',node=>node.account],['Customer',node=>node.customer],['FID Owner',node=>node.owner,(value,node)=>`<span title="${escapeHtml(ownerTip(node))}">${escapeHtml(value||'[Unknown]')}</span>`],['Slot Initiator',node=>node.slot.initiator],['Active CTL',node=>node.slot.active?.number],
  ['Note',node=>fidNotes[node.id]||'',(value,node)=>`<button type="button" class="fc-link fid-note-cell${value?'':' fid-note-add'}" data-note-fid="${escapeHtml(node.id)}" aria-label="${value?'Edit':'Add'} note for FID ${escapeHtml(node.id)}" title="${escapeHtml(value||'Add note')}">${escapeHtml(value?value.replace(/\r\n?|\n/g,' '):'…')}</button>`]
];
function fidDetailStatus(node,withTips=true){
  const tags=[];
  if(node.portfolioStates.includes('risk-free'))tags.push(['Stable Slot & CTL','clear',withTips?[riskTip(node.pre),riskTip(node.post)].join('\n'):'']);
  if(!node.portfolioStates.includes('slot-ready'))tags.push([fidStatusNames[node.portfolioStates[0]],'unknown',withTips?riskTip(node.pre):'']);
  else if(node.pre.state==='risk')tags.push(['Late Slot','risk',withTips?riskTip(node.pre):'']);
  else if(node.pre.state==='unknown')tags.push(['Slot date unknown','unknown',withTips?riskTip(node.pre):'']);
  if(node.slot.state==='dropped')tags.push(['Drop Slot','neutral',withTips?'Included in Slotted FID; risk uses the last approved Slot evidence.':'']);
  if(node.post.ctls.length)tags.push(['Late CTL ('+node.post.ctls.length+')','risk',withTips?riskTip(node.post):'']);
  const pending=node.post.nsrs.filter(nsr=>includeIgnored||!nsr.ignored);
  if(pending.length)tags.push(['Pending NSR ('+pending.length+')','pending',withTips?'In-work NSRs for this FID: '+pending.map(nsr=>nsr.number).join(', ')+'. Pending NSRs do not change Slot / CTL stability.':'']);
  for(const [label,kind]of [['Pending Config CTL','pending'],['Pending Pull/Push CTL','neutral']]){
    const risk=node.risks.find(risk=>risk.label===label);
    if(risk)tags.push([label+' ('+risk.count+')',kind,withTips?label+' for this FID: Submitted '+(kind==='pending'?'Slot / Config':'Pull / Push')+' CTLs.':'']);
  }
  if(node.portfolioStates.includes('slot-ready')&&node.post.state==='unknown')tags.push(['CTL status unknown','unknown',withTips?riskTip(node.post):'']);
  const selected=node.selectedQuarterStates||node.quarterStates,current=node.quarterStates.find(state=>state.quarter===node.currentQuarter);
  if(current&&selected?.length&&(node.pre.state!==current.pre.state||node.post.state!==current.post.state))tags.push(['Current: '+fidQuarterStatus(current),'neutral',withTips?'Current quarter '+(node.currentQuarter||'[Unknown]')+': '+fidQuarterStatus(current):'']);
  if(withTips&&selected?.length)for(const tag of tags)tag[2]=[tag[2],selected.map(state=>(state.quarter||'[Unknown]')+': '+fidQuarterStatus(state)).join('\n')].filter(Boolean).join('\n');
  return tags;
}
function fidQuarterStatus(state){
  if(!state.portfolioStates.includes('slot-ready'))return fidStatusNames[state.portfolioStates[0]];
  return [state.pre.state==='risk'?'Late Slot':'',state.post.state==='risk'?'Late CTL':''].filter(Boolean).join(' + ')||(state.portfolioStates.includes('risk-free')?'Stable Slot & CTL':'Risk Status Unknown');
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
  const m=current.metrics;if(!Object.keys(store.meta).length){target.innerHTML='';return}
  const riskFilters=queryFilters();delete riskFilters.riskType;
  const riskResult=filterValues.riskType?.length?graph.query(riskFilters):current,riskCounts=new Map([...(store.meta.nsr?['Pending NSR']:[]),...(store.meta.ctl||store.meta.ctlNsr?Object.keys(pendingCtlTypes):[]),...[...id('fRisk').options].map(option=>option.value)].map(label=>[label,{count:0,fids:0}]));
  for(const [label,key]of [['Pending Config CTL','pendingCtl'],['Pending Pull/Push CTL','pendingPullPushCtl']])if(riskCounts.has(label))riskCounts.set(label,{count:riskResult.metrics[key],fids:riskResult.metrics[key+'Fids']});
  for(const node of riskResult.fids)for(const risk of node.risks){
    const total=riskCounts.get(risk.label);if(!total||risk.label==='Pending NSR'||pendingCtlTypes[risk.label])continue;
    total.count++;total.fids++;
  }
  if(store.meta.nsr){
    const nsrFilters=queryFilters();if(nsrFilters.nsr)delete nsrFilters.nsr.taskStatus;
    const nsrs=filterValues['nsr.taskStatus']?.length?graph.query(nsrFilters,'nsrs'):current.nsrs,total=riskCounts.get('Pending NSR'),fids=new Set();
    for(const nsr of nsrs)if(taskGroup(nsr.taskStatus)==='workflow'){total.count++;for(const fid of nsr.matchedFidIds||nsr.fidIds)fids.add(fid)}
    total.fids=fids.size;
  }
  const filters=queryFilters(),selected=filterValues.fidStatus||[],causes=(filterValues.signal||[]).filter(key=>fidCauseKeys.includes(key));delete filters.fidStatus;
  filters.signal=filters.signal.filter(key=>!fidCauseKeys.includes(key));
  const rows=selected.length||causes.length?graph.query(filters,'fids'):current.fids,counts=Object.fromEntries(Object.keys(fidStatusNames).map(key=>[key,0]));
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
    ${!ctlReady?'<div class="fid-waiting" role="status">Load a CTL source to determine Slot Completion.</div>':''}`,drivers=`<div class="section-head"><h2>Risk Drivers</h2><span class="badge">Drivers can overlap</span></div><div class="signal-strip">${[...riskCounts].map(([label,{count,fids}])=>`<button class="signal" ${label==='Pending NSR'?'data-nsr-lifecycle="open"':pendingCtlTypes[label]?`data-ctl-status="workflow" data-ctl-types="${pendingCtlTypes[label].join(',')}" data-group-value="${escapeHtml(label)}"${taskFilterValues('ctl','workflow').length?'':' disabled'}`:`data-group-path="riskType" data-group-value="${escapeHtml(label)}"`} aria-pressed="${label==='Pending NSR'?taskFilterSelected('nsr','workflow'):pendingCtlTypes[label]?taskFilterSelected('ctl','workflow',pendingCtlTypes[label]):(filterValues.riskType||[]).includes(label)}" title="${escapeHtml(pretty(label,'riskType'))} = ${count} ${label==='Pending NSR'?'pending NSR':pendingCtlTypes[label]?'in-work CTL documents':'distinct FIDs'}${label.startsWith('Pending ')?' across '+fids+' distinct FID':''}"><b>${number(count)}</b> ${escapeHtml(pretty(label,'riskType'))}${label.startsWith('Pending ')?' <small>| in '+number(fids)+' FID</small>':''}</button>`).join('')||'<span class="muted">No risk drivers in this selection</span>'}</div>`,signals=`<button class="signal" data-nce-signal${store.meta.nce?'':' hidden'}><b>${number(m.nceFids)}</b>FIDs with NCE / FA</button>`;
  if(updateMarks){
    const scroll=target.querySelector('.fid-flow').scrollLeft;
    target.querySelector('.fid-overview').innerHTML=overview;
    target.querySelector('.fid-flow').scrollLeft=scroll;
    target.querySelector('.fid-risk-drivers').innerHTML=drivers;
    target.querySelector('[data-nce-signal]').parentElement.innerHTML=signals;
  }else target.innerHTML=`<section class="card fid-overview">${overview}</section><section class="card fid-risk-drivers">${drivers}</section>${pivotHtml('fid',current.fids.length,'FIDs',target)}<div class="signal-strip">${signals}</div><section class="card"><div class="section-head detail-head"><h2>FID Details</h2></div><div id="portfolioTable"></div></section>`;
  renderPivots('fidPivots',current.fids,[['quarter','Quarter','fQuarter','#4679cf'],['customer','FID Customer','fCustomer','#b44771'],['account','Account','fAccount','#347d50'],['businessUnit','Business Unit','fBusinessUnit','#8251bb'],['system','System Description','fSystem','#14869a']],'FIDs',updateMarks);
  table('portfolioTable',portfolioColumns,current.fids,{file:'FID_Portfolio',fid:true,markFilter:'fid',defaultSort:'FID',ascending:false,frozen:['FID'],exportRows:fidDetailRows,copyRows:fidCopyRows,widths:{FID:80,'FID Status & Risks':175,Note:220}},updateMarks);
}
const ctlColumns=[['CTL',ctl=>ctl.number||ctl.id,(value,ctl)=>ctl.number?`<a class="fc-link" href="https://pcazcrpap02.fremont.lamrc.net:8443/sap/bc/bsp/sap/crm_ui_start/default.htm?saprole=ZSLSDIR&amp;crm-object-type=CRM_SRQM_RFC&amp;crm-object-action=B&amp;crm-object-keyname=OBJECT_ID&amp;crm-object-value=${encodeURIComponent(ctl.number)}" target="_blank" rel="noopener noreferrer">${escapeHtml(value)}</a>`:escapeHtml(value)],['FID',ctl=>ctl.fidId,fidLink],
  ['Stage',ctl=>ctlStageName(ctl.stage),value=>value?`<span class="status-pill ${value.startsWith('0-')?'stop':value==='15-Completed'?'terminal':value==='Draft'?'unknown':''}">${escapeHtml(value)}</span>`:'—'],
  ['Change',ctl=>ctl.change],['Created Date',ctl=>ctl.created],
  ['Elapsed Days',ctl=>taskGroup(ctlStageName(ctl.stage),true)==='workflow'&&ctl.created?(nsrRecords.dateNumber(nsrData.localDate(new Date()))-nsrRecords.dateNumber(ctl.created))/86400000:''],
  ['Sales Ops Approval',ctl=>ctl.approved],['RD At Create',ctl=>ctl.rdCreate],['MCSD@Create',ctl=>ctl.mcsdCreate],
  ['SRD @ SOPs Approval',ctl=>ctl.srdApproval],['MCSD At SOps Approval',ctl=>ctl.mcsdApproval],['TSD @ SOPs Approval',ctl=>ctl.tsdApproval],
  ['SRD @ Current',ctl=>ctl.srd],['MCSD @ Current',ctl=>ctl.mcsd],['Config Frz',ctl=>ctl.configFrz],['Mfg Upgrade Type',ctl=>ctl.mfg],
  ['Initiator',ctl=>ctl.initiator],['Customer',ctl=>ctl.customer],['Quote ID',ctl=>ctl.quote],['NSR Reference',ctl=>ctl.nsrReferences.join(', ')],['Change Summary',ctl=>ctl.changeSummary],['Root Cause',ctl=>ctl.rootCause],['CTL Notes',ctl=>ctl.notes],['Reason for Chg',ctl=>ctl.reason],['Sub Reason',ctl=>ctl.subReason],['Sales OPs Approver',ctl=>ctl.salesOpsApprover]
];
function renderChanges(target,updateMarks=false){
  updateMarks=updateMarks&&tables.has('ctlTable');
  const selected=filterValues['ctl.stage']||[],filters=queryFilters();if(filters.ctl)delete filters.ctl.stage;
  const pendingOnly=filterValues.signal?.includes('pending-ctl'),rows=pendingOnly?current.ctls.filter(ctl=>ctl.pending):current.ctls,overviewRows=(selected.length?graph.query(filters,'ctls'):current.ctls).filter(ctl=>!pendingOnly||ctl.pending);
  const groups={completed:0,workflow:0,draft:0,rejected:0,cancelled:0,unknown:0};for(const ctl of overviewRows)groups[taskGroup(ctlStageName(ctl.stage),true)]++;
  const terms=[['completed','Completed',overviewColors.completed],['workflow','In work',overviewColors.workflow],['draft','Draft',overviewColors.draft],['rejected','Rejected',overviewColors.rejected],['cancelled','Cancelled',overviewColors.cancelled],['unknown','Stage unknown',overviewColors.unknown]].filter(([group])=>group!=='unknown'||groups.unknown||taskFilterSelected('ctl','unknown'));
  const overview=`<div class="section-head"><h2>CTL Overview</h2><span class="badge">${selected.length?number(rows.length)+' / '+number(overviewRows.length)+' selected CTL':number(rows.length)+' filtered CTL'}</span></div><div class="ctl-flow" style="--ctl-terms:${terms.length}" aria-label="Total CTL = ${terms.map(([,label])=>label).join(' + ')}">${taskOverview('ctl',terms,groups,overviewRows.length)}</div>`;
  if(updateMarks)target.querySelector('.ctl-overview').innerHTML=overview;
  else target.innerHTML=`<section class="card ctl-overview">${overview}</section>${pivotHtml('ctl',rows.length,'CTL',target)}<section class="card"><div class="section-head detail-head"><h2>CTL Details</h2></div><div id="ctlTable"></div></section>`;
  renderPivots('ctlPivots',rows,[['quarter','Quarter','fQuarter','#7950a1'],['stage','CTL Stage','fCtlStage','#9468bc'],['customer','Customer','fCtlCustomer','#b491cf']],'CTL',updateMarks);
  table('ctlTable',ctlColumns,rows,{file:pendingOnly?'Pending_CTL':'CTL_Details',defaultSort:'CTL',ascending:false,widths:{CTL:115,FID:85,Stage:145,Change:105,'Created Date':130,'Elapsed Days':130,'Sales Ops Approval':170,'RD At Create':135,'MCSD@Create':135,'SRD @ SOPs Approval':180,'MCSD At SOps Approval':190,'TSD @ SOPs Approval':180,'SRD @ Current':145,'MCSD @ Current':145,'Config Frz':115,'Mfg Upgrade Type':170,Initiator:155,Customer:185,'Quote ID':120,'NSR Reference':180,'Root Cause':300,'CTL Notes':300,'Reason for Chg':180,'Sub Reason':190,'Sales OPs Approver':180}},updateMarks);
}
function recordMark(kind,number){const config=markKinds[kind];return config.ignored.has(number)?'ignore':config.important.has(number)?'important':''}
function currentMarks(kind){const config=markKinds[kind];return Object.fromEntries([...config.ignored,...config.important].sort().map(number=>[number,recordMark(kind,number)]))}
function markSymbol(mark){return mark==='ignore'?'⊘':mark==='important'?'★':''}
function markButton(kind,number){
  const value=recordMark(kind,number),label=value==='ignore'?'⊘ Ignore':value==='important'?'★ Star':'Normal',next=value==='ignore'?'★ Star':value==='important'?'Normal':'⊘ Ignore';
  return `<button type="button" class="record-mark-toggle${kind==='nsr'?' nsr-mark-toggle':''}" data-mark-id="${escapeHtml(number)}" data-ignore-${kind}="${escapeHtml(number)}" data-${kind}-mark="${value}" data-record-mark="${value}" aria-label="${escapeHtml(number)}: ${label}; next ${next}" title="${label} → ${next}\nClick to cycle: Normal → ⊘ Ignore → ★ Star → Normal">${value==='ignore'?'<svg viewBox="0 0 12 12" aria-hidden="true" focusable="false"><circle cx="6" cy="6" r="4.5"/><path d="M3 9 9 3"/></svg>':value==='important'?'<svg viewBox="0 0 12 12" aria-hidden="true" focusable="false"><path fill="currentColor" stroke="none" d="M6 1 7.55 4.3 11 4.8 8.5 7.3 9.1 11 6 9.2 2.9 11 3.5 7.3 1 4.8 4.45 4.3Z"/></svg>':''}</button>`;
}
function markColumn(kind){return ['Mark',row=>recordMark(kind,kind==='fid'?row.id:row.number),(_,row)=>markButton(kind,kind==='fid'?row.id:row.number)]}
function markListHtml(kind){
  const config=markKinds[kind],prefix=config.prefix;
  return `<section class="card" id="ignored${prefix}s"><div class="section-head"><h2>Marked ${config.plural}</h2><button type="button" class="btn blue small" data-save-${kind}-marks>Save</button></div><div class="nsr-mark-editors"><div class="ag-box"><label for="fcIgnored${prefix}s">⊘ Ignore</label><textarea id="fcIgnored${prefix}s" data-${kind}-marks="ignore" rows="8" spellcheck="false" placeholder="One ${config.label} per line"></textarea></div><div class="ag-box"><label for="fcImportant${prefix}s">★ Star</label><textarea id="fcImportant${prefix}s" data-${kind}-marks="important" rows="8" spellcheck="false" placeholder="One ${config.label} per line"></textarea></div></div><div id="${config.status}" class="statistics-state" role="status"></div></section>`;
}
const markNsrColumn=markColumn('nsr');
const nsrColumns=[
  ['NSR#',nsr=>nsr.number,value=>`<button class="fc-link" data-cycle-nsr="${escapeHtml(value)}" title="Query Cycle Time">${escapeHtml(value)}</button>`],
  ['NSR Title',nsr=>nsr.title],['Task Status',nsr=>nsr.taskStatus,value=>value?`<span class="status-pill ${value==='15-Completed'?'terminal':value.startsWith('0-')?'stop':value.startsWith('Unmapped')?'unknown':''}">${escapeHtml(value)}</span>`:'—'],
  ['NSR Type',nsr=>nsr.type],['Impacted Sub System',nsr=>nsr.subsystem],['Initiator',nsr=>nsr.initiator],['FID',nsr=>(nsr.matchedFidIds||nsr.fidIds||[nsr.fidId]).filter(Boolean).join(', '),(_,nsr)=>(nsr.matchedFidIds||nsr.fidIds||[nsr.fidId]).filter(Boolean).map(fidLink).join(', ')||'—'],
  ['Primary Product',nsr=>nsr.product],
  ['NSR → CTL',nsr=>lifecycleNames[nsr.lifecycle]||nsr.lifecycle,(_,nsr)=>pill(lifecycleNames[nsr.lifecycle]||nsr.lifecycle,nsr.ignored?'ignore':nsr.lifecycle==='completed-unlinked'?'alert':['completed-5d','implemented'].includes(nsr.lifecycle)?'clear':nsr.lifecycle==='ctl-pending'?'pending':'info')],
  ['CTL references',nsr=>(nsr.links||[]).map(link=>link.ctlNumber+(link.kind==='reference'?' ('+link.fidId+')':'')).join(', ')],
  ['Customer',nsr=>nsr.customer],['NSR Org',nsr=>nsr.org],['Sales Rep',nsr=>nsr.record['Sales Rep']],['Sales Ops',nsr=>nsr.record['Sales Ops']],['Technical Contact',nsr=>nsr.record['Technical Contact']],
  ['Submit Date',nsr=>nsr.submitted],['NSR Request',nsr=>nsr.record['NSR Request']],['Temp Id#',nsr=>nsr.tempId]
];
function nsrCopyRows(rows,columns=nsrColumns){return nsrDetailRows(rows,columns.filter(([label])=>label!=='Mark'))}
function markedDetailRows(kind,rows,columns){return rows.map(row=>Object.fromEntries(columns.map(([label,get])=>[label,label==='Mark'?markSymbol(recordMark(kind,kind==='fid'?row.id:row.number)):get(row)??''])))}
function nsrDetailRows(rows,columns=[markNsrColumn,...nsrColumns]){return markedDetailRows('nsr',rows,columns)}
function fidDetailRows(rows,columns=portfolioColumns){return markedDetailRows('fid',rows,columns)}
function fidCopyRows(rows,columns=portfolioColumns){return fidDetailRows(rows,columns.filter(([label])=>label!=='Mark'))}
function renderNsrDetail(name,rows,options={},updateMarks=false,changedNsr=''){
  const columns=[markNsrColumn,...nsrColumns],widths={'Task Status':225,'NSR#':80,FID:80,'NSR Org':96};
  if(name==='fidNsrTable'){columns.splice(2,0,...columns.splice(columns.findIndex(([label])=>label==='NSR → CTL'),2));widths['NSR → CTL']=150}
  table(name,columns,rows,{file:'NSR_Detail',cycle:true,nsr:true,exportRows:nsrDetailRows,copyRows:nsrCopyRows,defaultSort:'FID',ascending:false,frozen:['NSR#'],widths,...options},updateMarks,changedNsr);
}
function highlightLinkedHistory(name,record){
  const state=tables.get(name);if(!state?.ready)return;
  const key=name==='fidCtlTable'?'ctlId':'nsrId',linked=new Set(record?.links?.map(link=>link[key]));
  for(const row of state.grid.getRows())row.getElement().classList.toggle('nsr-ctl-match',linked.has(state.rows[row.getIndex()].id));
}
function pivotHtml(kind,count,unit,target){return `<details class="card interactive-pivot"${target.querySelector('.interactive-pivot')?.open===false?'':' open'}><summary class="section-head"><h2>Interactive Pivot<span class="pivot-toggle" aria-hidden="true" title="Expand / collapse Interactive Pivot"></span></h2><span class="badge">${number(count)} filtered ${unit}</span></summary><div id="${kind}Pivots" class="nsr-pivots${kind==='ctl'?' ctl-pivots':''}"></div></details>`}
function renderPivots(name,rows,pivots,unit,updateMarks=false){
  const target=id(name);
  if(updateMarks)target.previousElementSibling.querySelector('.badge').textContent=number(rows.length)+' filtered '+unit;
  pivots.forEach(([path,label,filter,color],index)=>{
    const counts=new Map(),pivotRows=path==='quarter'&&name==='fidPivots'?quarterResult().fids:path==='quarter'&&name==='ctlPivots'?quarterResult().ctls:rows;
    pivotRows.forEach(row=>{for(const value of new Set(path==='quarter'?(row.selectedQuarterStates?.map(state=>state.quarter)||row.quarters||[row.quarter]).map(value=>value||'[Unknown]'):[row[path]||'[Unknown]']))counts.set(value,(counts.get(value)||0)+1)});
    const entries=[...counts].sort((a,b)=>path==='taskStatus'?nsrRecords.compareStatus(a[0],b[0]):path==='stage'?compareCtlStages(a[0],b[0]):path==='quarter'?Number(a[0]==='[Unknown]')-Number(b[0]==='[Unknown]')||nsrRecords.periodRank(a[0])-nsrRecords.periodRank(b[0])||a[0].localeCompare(b[0]):b[1]-a[1]||a[0].localeCompare(b[0])),max=Math.max(1,...counts.values());
    const content=entries.map(([value,count])=>`<div class="pivot-row"><span class="pivot-label" title="${escapeHtml(path==='stage'?ctlStageName(value):value)}">${escapeHtml(path==='stage'?ctlStageName(value):value)}</span><span class="pivot-bar"><i style="width:${100*count/max}%;--bar:${color}"></i></span><button class="pivot-count" data-pivot-filter="${filter}" data-pivot-value="${escapeHtml(value)}" title="Filter ${escapeHtml(value)}">${number(count)}</button></div>`).join('')||'<span class="muted">No matching '+unit+'</span>';
    if(updateMarks){const card=target.children[index],list=card.querySelector('.pivot-list'),scroll=list.scrollTop;card.querySelector('.pivot-head span').textContent=counts.size+' values';list.innerHTML=content;list.scrollTop=scroll}
    else target.insertAdjacentHTML('beforeend',`<section class="pivot-card"><div class="pivot-head"><h3>${label}</h3><span>${counts.size} values</span></div><div class="pivot-list">${content}</div></section>`);
  });
}
function renderNsrs(target,updateMarks=false,changedNsr=''){
  const rows=current.nsrs,selected=filterValues['nsr.taskStatus']||[],filters=queryFilters(),previous=tables.get('nsrTable');if(filters.nsr)delete filters.nsr.taskStatus;
  const sameRows=updateMarks&&previous&&rows.length===previous.rows.length&&rows.every((row,index)=>{const before=previous.rows[index];return row.number===before.number&&row.record===before.record&&row.taskStatus===before.taskStatus&&row.fidId===before.fidId&&row.matchedFidIds?.join()===before.matchedFidIds?.join()});
  const overviewRows=selected.length?graph.query(filters,'nsrs'):rows,groups={workflow:0,completed:0,rejected:0,cancelled:0},stageCounts={'completed-5d':0,'completed-unlinked':0,'ctl-pending':0,'implemented':0,'ctl-unavailable':0};
  overviewRows.forEach(nsr=>{const group=taskGroup(nsr.taskStatus);groups[group]++;if(group==='completed')stageCounts[Object.hasOwn(stageCounts,nsr.taskLifecycle)?nsr.taskLifecycle:'ctl-unavailable']++});
  const terms=[['completed','Completed',overviewColors.completed],['workflow','In workflow',overviewColors.workflow],['rejected','Rejected',overviewColors.rejected],['cancelled','Cancelled',overviewColors.cancelled]];
  const completedStages=[['completed-unlinked','alert','#a62a49','Completed NSR → no matching CTL in loaded data'],['ctl-pending','pending','#8a5c0c','Linked CTL → awaiting Sales Ops approval'],['implemented','approved',overviewColors.completed,'Linked CTL → Sales Ops approved'],['completed-5d','approved',overviewColors.completed,'Completed D-prefix (5D) NSR → no CTL required']];
  const showCompleted=groups.completed||selected.some(value=>value==='15-Completed'||Object.hasOwn(nsrCompletedTasks,value)),unknown=stageCounts['ctl-unavailable'];
  const completedEquation=`${groups.completed} = ${completedStages.map(([key])=>stageCounts[key]).join(' + ')}${unknown?' + '+unknown+' unknown':''}`;
  const overview=`${overviewHead('nsr',rows.length,overviewRows.length,selected.length)}<div class="nsr-flow${showCompleted?' has-completed':''}"><div class="nsr-equation" aria-label="Total NSR = ${terms.map(([,label])=>label).join(' + ')}">${taskOverview('nsr',terms,groups,overviewRows.length)}</div>
  ${showCompleted?`<div class="nsr-breakdown" role="group" aria-label="Completed NSR breakdown"><span class="nsr-completed-arrow" title="${escapeHtml(completedEquation)}" aria-hidden="true"></span><span class="math-sign nsr-breakdown-equals" aria-hidden="true">=</span><div class="nsr-stage-sum"><span class="nsr-stage-spacer" aria-hidden="true"></span>${completedStages.map(([key,kind,color,tip],index)=>{const count=stageCounts[key],active=taskFilterSelected('nsr',key),percent=groups.completed?(100*count/groups.completed).toFixed(1):'0.0';return (index?'<span class="math-sign" aria-hidden="true">+</span>':'')+`<button class="nsr-kpi nsr-stage ${kind}" style="--kpi-color:${color}" type="button" data-nsr-lifecycle="${key}" data-count="${count}" aria-pressed="${active}"${!count&&!active?' disabled':''} title="${escapeHtml(`${count} / ${groups.completed} = ${percent}%\n${tip}`)}"><small>${lifecycleNames[key]}</small><b>${number(count)}</b><span>${percent}% of Completed NSRs</span></button>`}).join('')}${unknown?`<span class="math-sign" aria-hidden="true">+</span><span class="nsr-kpi nsr-stage unknown" style="--kpi-color:#52647d" data-count="${unknown}" role="status" title="CTL sources are not loaded or approval status is missing. View these records under 15-Completed."><small>CTL status unknown</small><b>${number(unknown)}</b><span>${groups.completed?(100*unknown/groups.completed).toFixed(1):'0.0'}% of Completed NSRs</span></span>`:''}</div></div>`:''}</div>`;
  if(updateMarks)target.querySelector('.nsr-overview').innerHTML=overview;
  else target.innerHTML=`<section class="card nsr-overview">${overview}</section>${pivotHtml('nsr',rows.length,'NSRs',target)}<section class="card"><div class="section-head detail-head"><h2>NSR Details</h2></div><div id="nsrTable"></div></section>`;
  const pivots=[['taskStatus','Task Status','fNsrTaskStatus','var(--nsr-ink)'],['subsystem','Impacted Sub System','fNsrSubsystem','var(--nsr-ink)'],['type','NSR Type','fNsrType','var(--nsr-ink)'],['product','Primary Product','fNsrProduct','var(--nsr-ink)'],['org','NSR Org','fNsrOrg','var(--nsr-ink)'],['customer','Customer','fNsrCustomer','var(--nsr-ink)']];
  if(!sameRows)renderPivots('nsrPivots',rows,pivots,'NSRs',updateMarks);
  renderNsrDetail('nsrTable',rows,{markFilter:'nsr'},updateMarks,sameRows?changedNsr:'');
}
const summaryDimensions=[['quarter','Quarter'],['account','Account'],['customer','FID Customer'],['owner','Owner']];
const summaryMetricDefs=[
  ['Slot Completion%','slotted','fid','slot','slot-ready'],['Pre-slot Stability%','stableSlot','slotted','slot','stable-slot'],['Post-slot Stability%','stablePost','slotted','post','stable-ctl'],
  ['FID','fid','','fcid'],['Risk FID%','riskFid','slotted','slot','risk-fid'],['Stable Slot & CTL%','noLate','slotted','post','risk-free'],
  ['NSR','nsr','','nsr'],['In work NSR','open','','nsr'],['Completed NSR','completed','','nsr'],['Cancelled/Rejected NSR','discardedNsr','','nsr'],['Completed NSR but no CTL','unlinked','','linked'],
  ['CTL','ctl','','ctl'],['Submitted CTL','submitted','','ctl'],['Completed CTL','ctlCompleted','','ctl'],['Draft CTL','draft','','ctl'],['Cancelled/Rejected CTL','discardedCtl','','ctl'],['Pending Config CTL','pending','','ctl'],['Pending Pull/Push CTL','pendingPullPush','','ctl']
];
let summaryGroup='quarter',summaryPreviousGroup='quarter',summarySelected=['slotted','stableSlot','stablePost','fid','nsr','ctl'],summarySort={key:'',dir:-1},summaryCache={},summaryReport=[],quarterCache={};
function quarterResult(){
  if(quarterCache.revision!==revision)quarterCache={revision,result:graph.query(queryFilters(),'periods')};
  return quarterCache.result;
}
function summaryGroups(path){
  const result=path==='quarter'?quarterResult():current,quartersByFid=new Map();
  if(path==='quarter')for(const node of result.fids){if(!quartersByFid.has(node.id))quartersByFid.set(node.id,[]);quartersByFid.get(node.id).push(node.quarter||'[Unknown]');}
  const fidNames=fid=>path==='quarter'?quartersByFid.get(fid)||graph.fids.get(fid)?.quarterStates.map(state=>state.quarter||'[Unknown]')||['[Unknown]']:[graph.fids.get(fid)?.[path]||'[Unknown]'];
  const groups=new Map(),fidMetrics=summaryMetricDefs.filter(([,,,,state])=>state),entry=name=>{
    name=name||'[Unknown]';if(!groups.has(name))groups.set(name,{name,...Object.fromEntries(summaryMetricDefs.map(([,key])=>[key,0]))});return groups.get(name);
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
  return [...groups.values()].sort((a,b)=>Number(a.name==='[Unknown]')-Number(b.name==='[Unknown]')||(path==='owner'?b.riskFid-a.riskFid||b.fid-a.fid||b.nsr-a.nsr||b.ctl-a.ctl:path==='quarter'?nsrRecords.periodRank(a.name)-nsrRecords.periodRank(b.name):0)||tableCollator.compare(a.name,b.name));
}
function renderSummaryMetrics(target){
  if(!id('summaryChart')){
    target.innerHTML=`<section class="card"><div class="section-head"><div class="summary-heading overview-tools"><h2>Summary Metrics</h2><span id="summaryCount" class="badge" aria-live="polite"></span>${ignoredToggle('summary')}</div><div class="btnrow"><button class="btn light small" data-summary-export="copy">Copy</button><button class="btn light small" data-summary-export="xlsx">Export XLSX</button></div></div><div class="summary-controls"><div class="fc-filter"><label for="summaryGroup">Group by</label><select id="summaryGroup"></select></div><div class="fc-filter"><label for="summaryMetrics">Metrics</label><select id="summaryMetrics" multiple>${summaryMetricDefs.map(([label,key])=>`<option value="${key}"${summarySelected.includes(key)?' selected':''}>${label}</option>`).join('')}</select></div></div><div id="summaryChart" class="summary-chart" aria-live="polite"></div><div id="summaryLegend" class="summary-legend" hidden>FID by Owner · NSR / CTL by Initiator.</div></section>`;
    jQuery(id('summaryMetrics')).chosen({width:'100%',search_contains:true,hide_results_on_select:false,display_selected_options:false,placeholder_text_multiple:'Select metrics'});
    jQuery('#summaryGroup,#summaryMetrics').on('change',event=>{if(event.target.id==='summaryGroup'){summaryGroup=event.target.value;if(summaryGroup!=='owner')summaryPreviousGroup=summaryGroup}else summarySelected=[...event.target.selectedOptions].map(option=>option.value);renderSummaryMetrics(target)});
  }
  id('summaryGroup').innerHTML=summaryDimensions.filter(([value])=>value!=='owner'||summaryGroup==='owner').map(([value,label])=>`<option value="${value}"${value===summaryGroup?' selected':''}>${label}</option>`).join('');
  id('summaryLegend').hidden=summaryGroup!=='owner';
  const chart=id('summaryChart');summaryReport=[];
  if(!summarySelected.includes(summarySort.key))summarySort.key='';
  for(const button of target.querySelectorAll('[data-summary-export]'))button.disabled=true;
  if(!summarySelected.length){chart.innerHTML='<div class="summary-empty">Select at least one metric.</div>';id('summaryCount').textContent='';return}
  if(summaryCache.revision!==revision||summaryCache.group!==summaryGroup){
    summaryCache={revision,group:summaryGroup,rows:summaryGroups(summaryGroup)};
    if(summaryGroup==='quarter')summarySort.key='';
  }
  const rows=summaryCache.rows.slice(),ctlReady=Boolean(store.meta.ctl||store.meta.ctlNsr),ready={fcid:!!store.meta.fcid,ctl:ctlReady,nsr:!!store.meta.nsr,slot:!!store.meta.fcid&&ctlReady,post:!!store.meta.fcid&&ctlReady,linked:ctlReady&&!!store.meta.nsr};
  id('summaryCount').textContent=rows.length+' groups';
  if(!rows.length){chart.innerHTML='<div class="summary-empty">No matching records.</div>';return}
  const metrics=summaryMetricDefs.filter(([,key])=>summarySelected.includes(key)),dimension=summaryDimensions.find(([key])=>key===summaryGroup)[1];
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

const tableCollator=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'});
const columnPopup=id('columnFilterPopup'),columnInput=columnPopup.querySelector('input'),markPopup=id('markFilterPopup');
// Hidden tabs keep their measured size, so showing an unchanged table does not rebuild its rows.
const tableResize=new ResizeObserver(entries=>{for(const {target} of entries){const state=tables.get(target.dataset.tableName);if(!state?.ready||!target.clientWidth||!target.clientHeight)continue;const size=target.clientWidth+':'+target.clientHeight;if(size!==state.size){state.size=size;state.grid.redraw()}}});
function orderTableColumns(definitions,columns,order,nsrHistory=false){
  if(nsrHistory){const leading=['NSR → CTL','CTL references'];order=[...leading,...(order||columns.map(([label])=>label)).filter(label=>!leading.includes(label))]}
  if(!Array.isArray(order))return definitions;
  const rank=new Map(order.map((label,index)=>[label,index])),label=column=>columns[Number(column.field.slice(1))][0];
  const movable=new Set(definitions.filter(column=>!column.frozen&&rank.has(label(column)))),sorted=[...movable].sort((a,b)=>rank.get(label(a))-rank.get(label(b)));let index=0;
  return definitions.map(column=>movable.has(column)?sorted[index++]:column);
}
function table(name,columns,rows,options={},updateMarks=false,changedNsr=''){
  // Tabulator deep-clones row data; keep the full graph records in state.rows.
  const previous=tables.get(name),dataRow=(row,index)=>({id:index,...Object.fromEntries(columns.map(([,get],column)=>['c'+column,get(row)??'']))});
  if(previous&&!previous.ready&&previous.built)return previous.built.then(()=>table(name,columns,rows,options,updateMarks,changedNsr));
  if(previous?.ready&&updateMarks){
    const wrap=previous.grid.element,active=document.activeElement,focused=wrap.contains(active)?active.dataset.markId||active.dataset.noteFid:'',focusSelector=active.matches('[data-mark-id]')?'[data-mark-id]':'[data-note-fid]';previous.rows=rows;previous.columns=columns;previous.options=options;wrap.setAttribute('aria-busy','true');
    const index=changedNsr?rows.findIndex(row=>(options.fid?row.id:row.number)===changedNsr):-1,updated=index>=0?previous.grid.updateData([dataRow(rows[index],index)]):previous.grid.replaceData(rows.map(dataRow));
    return updated.then(()=>{
      if(index>=0){if(previous.local.marks?.length||previous.local.headerFilters?.length)previous.grid.refreshFilter();else paintMarkFilter(previous);const sorters=previous.grid.getSorters();if(sorters.some(sorter=>['NSR → CTL','Note'].includes(columns[Number(sorter.field.slice(1))]?.[0])))previous.grid.setSort(sorters.map(sorter=>({column:sorter.field,dir:sorter.dir})))}
      wrap.setAttribute('aria-busy','false');if(focused)[...wrap.querySelectorAll(focusSelector)].find(button=>(button.dataset.markId||button.dataset.noteFid)===focused)?.focus({preventScroll:true});
      return previous;
    });
  }
  if(previous){tableResize.unobserve(previous.grid.element);previous.grid.destroy()}
  const local=localTables.get(name)||{marks:null,sort:options.defaultSort?columns.findIndex(([label])=>label===options.defaultSort):-1,ascending:options.ascending??true},orderKey=name==='fidCtlTable'?'ctlTable':name==='fidNsrTable'?'nsrTable':name,state={name,columns,rows,options,local,orderKey};
  if(name==='fidCtlTable'&&!id('fidDetail').open)local.showPullPush=false;
  tables.set(name,state);localTables.set(name,local);
  const host=id(name),autoHeight=Boolean(host.closest('.fid-dialog')),columnFilters=!autoHeight&&options.columnFilters!==false;if(!columnFilters)local.headerFilters=[];
  host.innerHTML=`<div class="fc-table-shell"><div class="table-toolbar" data-table-tools="${name}">${autoHeight?`<div class="btnrow"><h2>${escapeHtml(options.title)} (<span data-table-count></span>)</h2>${name==='fidCtlTable'?`<button type="button" class="fc-link" data-toggle-pull-push aria-pressed="${Boolean(local.showPullPush)}">${local.showPullPush?'Hide':'Show'} Pull/Push</button>`:''}</div>`:'<span class="table-count" data-table-count></span>'}<div class="toolbar-right${autoHeight?' btnrow':''}"><button class="btn light small" data-table-copy="${name}"${options.copyId?` id="${options.copyId}"`:''}>Copy</button>${autoHeight?'':`<button class="btn light small" data-table-export="${name}"${options.exportId?` id="${options.exportId}"`:''}>Export XLSX</button>`}${options.cycle?`<button class="btn blue small" data-table-cycle="${name}">Cycle Time</button>`:''}</div></div>${columnFilters?'<div class="column-filter-bar"></div>':''}<div class="fc-table-wrap" data-table-name="${name}" aria-label="${escapeHtml(options.file||name)}" aria-busy="true"></div></div>`;
  state.toolbar=host.querySelector('.table-toolbar');
  const heading=autoHeight?null:host.parentElement.querySelector(':scope > .detail-head, :scope > summary');
  if(heading&&state.toolbar){heading.querySelector('[data-table-tools="'+name+'"]')?.remove();heading.append(state.toolbar)}
  const wrap=host.querySelector('.fc-table-wrap');wrap.classList.toggle('column-filters',columnFilters);
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
  }
  const filter=item=>(name!=='fidCtlTable'||local.showPullPush||!['pull','push'].includes(state.rows[item.id].changeKey))&&(!options.markFilter||!local.marks||local.marks.includes(recordMark(options.markFilter,options.markFilter==='fid'?state.rows[item.id].id:state.rows[item.id].number)||'normal'));
  const columnDefs=columns.map(([label,,render,settings],index)=>({title:label==='Mark'?'':label,titleFormatter:label==='Mark'&&options.markFilter?()=>markHeader(state):columnFilters&&label!=='Mark'?columnHeader:undefined,field:'c'+index,frozen:index===0||(options.frozen||[]).includes(label),minWidth:label==='Mark'?36:75,headerSort:label!=='Mark',headerFilter:columnFilters&&label!=='Mark'?'input':false,headerFilterParams:{elementAttributes:{'aria-label':'Filter '+label}},cellClick:name==='portfolioTable'&&label==='FID Status & Risks'?(_event,cell)=>showFid(state.rows[cell.getRow().getIndex()].id):undefined,hozAlign:label==='Mark'?'center':'left',headerHozAlign:label==='Mark'?'center':'left',width:widths[label]??(label==='Mark'?36:label==='Task Status'?275:/Title|Summary|Documents|Sample|File|Reason|Description|Action/.test(label)?320:/Owner|Initiator|Candidate|Customer|Product/.test(label)?185:/^Pre$|^Post$|NSR → CTL|Application/.test(label)?180:Math.max(110,Math.min(240,label.length*7+28))),sorter:(a,b)=>label==='Stage'?compareCtlStages(a,b):/Quarter|Period/.test(label)?nsrRecords.periodRank(a)-nsrRecords.periodRank(b)||tableCollator.compare(String(a),String(b)):typeof a==='number'&&typeof b==='number'?a-b:tableCollator.compare(String(a),String(b)),...settings,formatter:cell=>{const value=cell.getValue();cell.getElement().title=label==='Mark'?'':String(value);return '<span class="cell-content">'+(render?render(value,state.rows[cell.getRow().getIndex()]):escapeHtml(value)||'—')+'</span>'}}));
  state.grid=new Tabulator(wrap,{data,autoResize:false,movableColumns:true,height:autoHeight?false:options.height??Math.min(620,Math.max(105,rows.length*34+68)),rowHeight:options.rowHeight??34,renderVertical:autoHeight?'basic':'virtual',renderVerticalBuffer:200,columnDefaults:{resizable:'header',headerFilterFunc:'like',headerFilterPlaceholder:'Filter…'},layout:'fitDataFill',placeholder:'',dataLoader:false,initialFilter:filter,initialHeaderFilter:local.headerFilters||[],initialSort:local.sort>=0?[{column:'c'+local.sort,dir:local.ascending?'asc':'desc'}]:[],rowFormatter:row=>{const element=row.getElement(),data=row.getData();element.dataset.rowIndex=row.getIndex();if(state.options.rowRisk)element.dataset.fidRisk=state.options.rowRisk(state.rows[row.getIndex()]);if(options.nsr||options.fid){const mark=recordMark(options.nsr?'nsr':'fid',options.nsr?state.rows[data.id].number:state.rows[data.id].id);element.dataset.recordMark=mark;if(options.nsr)element.dataset.nsrMark=mark;else element.dataset.fidMark=mark;}if(name==='nsrTable'){element.dataset.fidBand=data.fidBand;element.dataset.fidStart=data.fidStart}},columns:orderTableColumns(columnDefs,columns,tableOrders[orderKey],name==='fidNsrTable')});
  state.grid.on('tableDestroyed',()=>{if(columnPopup.column?.getTable()===state.grid)columnPopup.hidePopover();if(markPopup.state===state)markPopup.hidePopover()});
  state.grid.on('columnMoved',()=>{
    if(state.ordering)return;
    const order=state.grid.getColumns().map(column=>columns[Number(column.getField().slice(1))][0]).filter(label=>name!=='fidNsrTable'||!['NSR → CTL','CTL references'].includes(label));let index=0;
    const base=tableOrders[orderKey]||(name==='fidNsrTable'?['Mark',...nsrColumns.map(([label])=>label)]:[]),merged=base.map(label=>order.includes(label)?order[index++]:label);merged.push(...order.slice(index));tableOrders[orderKey]=merged;
    try{GM_setValue(TABLE_ORDERS_KEY,JSON.stringify(tableOrders))}catch(error){toast('Could not save column order: '+error.message)}
    for(const target of tables.values())if(target.ready&&target.orderKey===orderKey&&(target!==state||name==='fidNsrTable')){
      target.ordering=true;
      try{orderTableColumns(target.grid.getColumns().map(column=>column.getDefinition()),target.columns,merged,target.name==='fidNsrTable').forEach((column,index)=>{const current=target.grid.getColumns()[index];if(!column.frozen&&current.getField()!==column.field)target.grid.moveColumn(column.field,current.getField(),false)})}
      finally{target.ordering=false}
    }
  });
  state.grid.on('dataFiltered',(_filters,rows)=>{local.headerFilters=state.grid.getHeaderFilters();if(columnFilters)paintColumnFilters(state);const count=state.toolbar?.querySelector('[data-table-count]');if(count)count.textContent=number(rows.length)+(autoHeight?'':' rows');paintMarkFilter(state)});
  if(options.markFilter)state.grid.on('renderComplete',()=>{if(markPopup.state===state)markPopup.position=markMenuPosition(state)});
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
    state.grid.on('rowMouseEnter',(_event,row)=>highlightLinkedHistory(linkedTable,state.rows[row.getIndex()]));
    state.grid.on('rowMouseLeave',()=>highlightLinkedHistory(linkedTable));
  }
  state.grid.on('dataProcessed',()=>{state.ready=true;state.size=wrap.clientWidth+':'+wrap.clientHeight;wrap.setAttribute('aria-busy','false')});tableResize.observe(wrap);
  return state.built=new Promise(resolve=>state.grid.on('tableBuilt',()=>resolve(state)));
}
function columnHeader(cell){
  const column=cell.getColumn(),title=column.getDefinition().title,header=document.createElement('span');header.className='column-heading';
  header.innerHTML='<span>'+escapeHtml(title)+(title==='FID Owner'?inferenceHelp(ownerHint):'')+'</span><button type="button" class="column-funnel" popovertarget="columnFilterPopup" aria-label="Filter '+escapeHtml(title)+'" aria-haspopup="dialog" aria-pressed="false" title="Filter '+escapeHtml(title)+'"><svg viewBox="0 0 16 16" aria-hidden="true"><path d="M1 2h14L10 8v5l-4 2V8z"/></svg></button>';
  const button=header.lastElementChild;button.onmousedown=event=>event.stopPropagation();
  button.onclick=event=>{event.stopPropagation();markPopup.hidePopover();columnPopup.hidePopover();columnPopup.column=column;columnPopup.querySelector('label>span').textContent=title;columnInput.value=column.getHeaderFilterValue()??''};return header;
}
function paintColumnFilters(state){
  const filters=state.local.headerFilters,key=JSON.stringify(filters);if(state.columnFilterKey===key)return;state.columnFilterKey=key;
  const host=id(state.name),active=new Map(filters.map(filter=>[filter.field,filter.value]));
  for(const column of state.grid.getColumns()){
    const button=column.getElement().querySelector('.column-funnel');if(!button)continue;const selected=active.has(column.getField());button.classList.toggle('active',selected);button.setAttribute('aria-pressed',String(selected));button.title='Filter '+column.getDefinition().title+(selected?': '+active.get(column.getField()):'');
  }
  host.querySelector('.column-filter-bar').innerHTML=filters.map(({field,value})=>{const text=escapeHtml(state.grid.getColumn(field).getDefinition().title+': '+value);return '<button type="button" class="pill info" data-column-remove="'+state.name+'" data-field="'+field+'" title="'+text+'" aria-label="Remove filter '+text+'">'+text+' ×</button>'}).join('')+(filters.length?'<button type="button" class="fc-link" data-column-clear="'+state.name+'">Clear</button>':'');
}
function tableResult(state){return state.grid.getRows('active').map(row=>state.rows[row.getIndex()])}
function paintTable(state){state.grid.refreshFilter()}


function exportTableRows(state,copy=false){const rows=tableResult(state),columns=state.grid.getColumns().filter(column=>column.isVisible()).map(column=>state.columns[Number(column.getField().slice(1))]),format=copy?state.options.copyRows:state.options.exportRows;return format?format(rows,columns):rows.map(row=>Object.fromEntries(columns.map(([label,get])=>[label,get(row)??''])))}
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
  if(!updateMarks)target.innerHTML=`<div class="btnrow" style="justify-content:flex-end"><button class="btn light small" data-import-settings>Import Settings</button><button class="btn light small" data-export-settings>Export Settings</button><button class="btn light small" data-export-data-sources>Export All Data Sources</button><input id="fcSettingsFile" type="file" accept=".json,application/json" hidden></div><details class="card"><summary>Excluded Commit FIDs · ${number(excluded.length)}</summary><div class="signal-strip" style="margin-top:10px">${[...reasons].map(([reason,count])=>`<span class="signal"><b>${number(count)}</b>${escapeHtml(reason)}</span>`).join('')}</div><div id="excludedTable"></div></details><div class="marked-settings">${markListHtml('nsr')}${markListHtml('fid')}</div><section class="card" id="fabAccountMapping"><div class="section-head"><h2>Fab / Account Mapping</h2><div class="btnrow"><button class="btn blue small" data-save-mapping>Save</button><button class="btn light small" data-reset-mapping>Restore Defaults</button></div></div><p class="muted">Use Fab names and FID / NSR Customer names. Exact names take priority over * wildcards.</p><div id="agNotice"></div><div id="agGrid" class="ag-grid"></div></section>`;
  table('excludedTable',[['FID',node=>node.id,fidLink],['Reason',node=>node.excludedReason],['System',node=>node.system],['Account',node=>node.account],['Quarter',node=>node.quarter]],excluded,{file:'Excluded_Commit_FIDs'},updateMarks);
  for(const [kind,config]of Object.entries(markKinds))for(const textarea of target.querySelectorAll(`[data-${kind}-marks]`)){const mark=textarea.dataset[kind+'Marks'];textarea.value=markDrafts[kind]?.[mark]??[...(mark==='ignore'?config.ignored:config.important)].sort(tableCollator.compare).reverse().join('\n')}
  if(!updateMarks)paintMapping();
}
function paintMapping(){if(!id('agGrid'))return;const mapping=mappingDraft||accountMapping.getMap(),status=accountMapping.getStatus();id('agGrid').innerHTML=Object.entries(mapping).map(([group,fabs],index)=>`<div class="ag-box"><label for="ag-${index}">${escapeHtml(group)}</label><textarea id="ag-${index}" data-account-group="${escapeHtml(group)}">${escapeHtml(fabs)}</textarea></div>`).join('');id('agNotice').hidden=!(status.notice||status.error);id('agNotice').className='ag-notice';id('agNotice').textContent=status.error||status.notice}
function setRecordMark(kind,number){
  const mark=recordMark(kind,number),next=mark==='ignore'?'important':mark==='important'?'':'ignore',marks=currentMarks(kind);if(next)marks[number]=next;else delete marks[number];
  const result=saveMarks(kind,marks);if(!result.ok){toast('Could not save '+kind.toUpperCase()+' mark: '+result.error);return}
  const started=performance.now();modelRevision++;
  if(kind==='nsr'){graph.setNsrMark(number,next);applyFilters(true,number)}
  else{graph.setFidMark(number,next);if(includeIgnored)refreshFidPersonal(number);else applyFilters(true)}
  store.buildMs=performance.now()-started;
}
function saveMarks(kind,marks,persist=true){
  const config=markKinds[kind];if(persist)try{GM_setValue(config.key,JSON.stringify(marks))}catch(error){return {ok:false,error:error.message}}
  config.ignored.clear();config.important.clear();for(const [number,mark]of Object.entries(marks))(mark==='ignore'?config.ignored:config.important).add(number);
  return {ok:true};
}
function saveMarkLists(kind){
  const config=markKinds[kind],lists=Object.fromEntries([...document.querySelectorAll(`[data-${kind}-marks]`)].map(textarea=>[textarea.dataset[kind+'Marks'],new Set(textarea.value.toUpperCase().split(/[\s,;|，；]+/).filter(Boolean))]));
  const invalid=[...lists.ignore,...lists.important].filter(number=>!config.pattern.test(number)),overlap=[...lists.ignore].filter(number=>lists.important.has(number)),status=id(config.status);
  status.className='statistics-state';status.textContent='';
  if(invalid.length||overlap.length){status.className+=' error';status.textContent=invalid.length?'Invalid '+config.label+': '+[...new Set(invalid)].join(', '):config.label+' appears in both lists: '+overlap.join(', ');return}
  const marks=Object.fromEntries(Object.entries(lists).flatMap(([mark,numbers])=>[...numbers].map(number=>[number,mark]))),result=saveMarks(kind,marks);
  if(!result.ok){status.className+=' error';status.textContent='Could not save: '+result.error;return}
  markDrafts[kind]=null;
  if(kind==='nsr'){for(const nsr of graph.nsrs)graph.setNsrMark(nsr.number,recordMark(kind,nsr.number))}else{for(const node of graph.fids.values())graph.setFidMark(node.id,recordMark(kind,node.id))}
  applyFilters(true);
  id(config.status).textContent='Saved';
}
function markMenuPosition(state){const holder=state.grid.element.querySelector('.tabulator-tableholder');return holder?[scrollX,scrollY,holder.scrollLeft,holder.scrollTop].join(':'):''}
function markHeader(state){
  const button=document.createElement('button');button.type='button';button.className='mark-filter-trigger';button.setAttribute('popovertarget','markFilterPopup');button.setAttribute('aria-haspopup','dialog');
  button.innerHTML='<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M1 2h14L10 8v5l-4 2V8z"/></svg>';
  button.onmousedown=event=>event.stopPropagation();
  button.onclick=event=>{event.stopPropagation();columnPopup.hidePopover();markPopup.hidePopover();markPopup.state=state;markPopup.trigger=button;markPopup.position=markMenuPosition(state);paintMarkFilter(state)};
  return button;
}
function paintMarkFilter(state){
  const kind=state.options.markFilter;if(!kind)return;
  const selected=state.local.marks??markOptions,active=Boolean(state.local.marks),trigger=state.grid.element.querySelector('.mark-filter-trigger'),label='Filter marks: '+(active?(selected.length?selected.map(mark=>markLabels[mark]).join(', '):'None'):'All');
  if(trigger){trigger.classList.toggle('active',active);trigger.setAttribute('aria-pressed',String(active));trigger.setAttribute('aria-label',label);trigger.title=label}
  if(markPopup.state!==state)return;
  const host=markPopup.querySelector('[role=group]'),counts={all:0,normal:0,ignore:0,important:0};
  for(const {id:index}of state.grid.searchData(state.grid.getHeaderFilters())){const record=state.rows[index];counts.all++;counts[recordMark(kind,kind==='fid'?record.id:record.number)||'normal']++}
  if(host.dataset.tableName!==state.name){host.dataset.tableName=state.name;host.innerHTML=['all',...markOptions].map(mark=>`<label><input type="checkbox" data-table-mark-filter="${state.name}" data-mark="${mark}"><span>${markLabels[mark]}</span><b aria-hidden="true"></b></label>`).join('')}
  for(const input of host.querySelectorAll('input')){const mark=input.dataset.mark;input.checked=mark==='all'?selected.length===markOptions.length:selected.includes(mark);input.indeterminate=mark==='all'&&selected.length>0&&selected.length<markOptions.length;input.dataset.count=counts[mark];input.setAttribute('aria-label',markLabels[mark]+' ('+number(counts[mark])+')');input.parentElement.querySelector('b').textContent=number(counts[mark])}
}
function paintFidPersonal(fid){
  const dialog=id('fidDetail');if(dialog.dataset.fid!==fid||!id('fidPersonalMark'))return;
  const focused=document.activeElement.dataset.ignoreFid===fid,note=fidNotes[fid]||'';id('fidPersonalMark').innerHTML=markButton('fid',fid);
  const panel=id('fidNotePanel');if(!panel.querySelector('form'))panel.innerHTML=`<button type="button" class="fid-note-view" data-note-fid="${escapeHtml(fid)}" aria-label="Edit note for FID ${escapeHtml(fid)}"><span class="fid-note-title" role="heading" aria-level="2">Note</span><span class="fid-note-text${note?'':' muted'}">${escapeHtml(note||'No note')}</span></button>`;
  if(focused)id('fidPersonalMark').querySelector('button').focus({preventScroll:true});
}
function refreshFidPersonal(fid){
  const state=tables.get('portfolioTable'),update=state?.rows.some(node=>node.id===fid)?table('portfolioTable',portfolioColumns,state.rows,state.options,true,fid):null;
  paintFidPersonal(fid);return update;
}
function editFidNote(fid,anchor){
  const dialog=id('fidNoteDialog'),panel=anchor.closest('#fidNotePanel'),host=panel||dialog,rect=anchor.getBoundingClientRect();host.dataset.fid=fid;
  host.innerHTML=`<form id="fidNoteForm"><div class="section-head"><h2 id="fidNoteTitle">${panel?'Note':'Note · FID '+escapeHtml(fid)}</h2></div><textarea id="fidNoteText" rows="8" aria-label="Note" placeholder="Write a note…" spellcheck="false">${escapeHtml(fidNotes[fid]||'')}</textarea><div id="fidNoteStatus" class="statistics-state" role="status"></div><div class="fc-note-actions"><button class="btn light" type="button" data-cancel-fid-note>Cancel</button><button class="btn blue" type="submit">Save</button></div></form>`;
  if(!panel){
    if(!dialog.open)dialog.showModal();
    const box=dialog.getBoundingClientRect(),save=dialog.querySelector('button[type="submit"]').getBoundingClientRect();
    const left=rect.left+rect.width/2-(save.left-box.left+save.width/2),top=rect.top+rect.height/2-(save.top-box.top+save.height/2);
    dialog.style.left=Math.max(8,Math.min(left,innerWidth-box.width-8))+'px';
    dialog.style.top=Math.max(8,Math.min(top,innerHeight-box.height-8))+'px';
  }
  id('fidNoteText').focus({preventScroll:true});
}
function closeFidNote(){
  const host=id('fidNoteForm').parentElement,fid=host.dataset.fid;if(host.id==='fidNoteDialog')host.close();host.replaceChildren();if(host.id==='fidNotePanel')paintFidPersonal(fid);
  const target=id('fidDetail').open?id('fidDetail'):id('portfolioTable');[...target.querySelectorAll('[data-note-fid]')].find(button=>button.dataset.noteFid===fid)?.focus({preventScroll:true});
}
async function saveFidNote(){
  const fid=id('fidNoteForm').parentElement.dataset.fid,note=id('fidNoteText').value,next={...fidNotes};if(note.trim())next[fid]=note;else delete next[fid];
  try{GM_setValue(FID_NOTES_KEY,JSON.stringify(next))}catch(error){id('fidNoteStatus').className='statistics-state error';id('fidNoteStatus').textContent='Could not save: '+error.message;return}
  fidNotes=next;await refreshFidPersonal(fid);closeFidNote();
}
function fidQuarterRows(node){return [...node.quarterStates].sort((a,b)=>Number(b.quarter===node.currentQuarter)-Number(a.quarter===node.currentQuarter))}
function showFid(fid,updateMarks=false,changedNsr=''){
  const original=graph.fids.get(fid);if(!original)return;const node=graph.projectFid(original,queryFilters().quarter)||graph.projectFid(original),dialog=id('fidDetail');if(dialog.dataset.fid!==fid)['fidNsrTable','fidCtlTable','fidNceTable'].forEach(name=>localTables.delete(name));
  const updateDetail=updateMarks&&dialog.dataset.fid===fid&&tables.has('fidNsrTable');dialog.dataset.fid=fid;
  const lateCtlIds=new Set(original.lateCtls.map(ctl=>ctl.id)),lateSlotIds=new Set(original.quarterStates.flatMap(state=>state.slotEvidence.filter(pre=>pre.state==='risk').map(pre=>pre.ctlId)));
  const riskLabels={'On time':'Stable Slot','Pre late risk':'Late Slot','Post config risk':'Late CTL','No post risk':'Stable CTL'};
  const quarters=fidQuarterRows(original),multiple=quarters.length>1;
  const quarterLines=quarters.map(state=>`<div class="fid-quarter-state${state.quarter===original.currentQuarter?' is-current':''}" data-quarter-state="${escapeHtml(state.quarter)}" data-selected-quarter="${node.selectedQuarterStates.includes(state)}" title="${escapeHtml(state.basis.join('\n'))}"><span>${escapeHtml(state.quarter||'[Unknown]')}</span>${multiple?pill(fidQuarterStatus(state),state.portfolioStates.includes('risk-fid')?'risk':state.portfolioStates.includes('risk-free')?'clear':'unknown',[riskTip(state.pre),riskTip(state.post)].join('\n')):''}</div>`).join('');
  const summary=`<div class="fid-dialog-grid" style="--fid-date-rows:${quarters.length+1}">${[
    ['Slot',node.slot.state.replace(/^./,letter=>letter.toUpperCase())],['Commit: Field',node.commit],['Account',node.account],['Business Unit',node.businessUnit],
    ['System Description',node.system],['FID Description',node.fidDescription],['Customer',node.customer],['Fab Name',node.fab],
    ['Quarter',node.quarter,'quarter'],['Sales Ops RD',node.rd,'rd'],['MCSD',node.mcsd,'mcsd'],['Customer Request Date',node.crd,'crd'],
    ['FID Owner',node.owner],['Slot Initiator',node.slot.initiator],['Sales Admin',node.salesAdmin]
  ].map(([label,value,dateKey])=>{
    const lines=multiple&&dateKey&&dateKey!=='quarter'?(dateKey==='crd'?`<div class="fid-quarter-state"><span>${escapeHtml(value||'[Unknown]')}</span></div>`:quarters.map(state=>`<div class="fid-quarter-state${state.quarter===original.currentQuarter?' is-current':''}" data-quarter-date="${escapeHtml(state.quarter)}" title="${escapeHtml((state.quarter||'[Unknown]')+' · '+state.dateSource+(!state.dates[dateKey]?'\nDate unavailable.':''))}"><span>${escapeHtml(state.dates[dateKey]||'—')}</span></div>`).join('')):'';
    return `<div class="fid-detail-value${multiple&&dateKey?' fid-date-value':''}${label==='Slot'&&node.portfolioStates.includes('no-slot')?' no-slot':''}"${dateKey?` data-date-field="${dateKey}"`:''}${label==='FID Owner'?` title="${escapeHtml(ownerTip(node))}"`:label.startsWith('Sales Admin')?' title="Unique Pull / Push CTL votes; ties remain unknown"':''}><small>${label}${label==='Quarter'?inferenceHelp(quarterHint):label==='FID Owner'?inferenceHelp(ownerHint):label==='Sales Admin'?inferenceHelp(salesAdminHint):''}</small>${label==='Quarter'?quarterLines:lines||escapeHtml(value||'[Unknown]')}</div>`;
  }).join('')}</div><div class="fid-evidence">${[['Pre-slot',node.pre],['Post-slot',node.post]].map(([name,risk])=>`<div class="notice ${name==='Pre-slot'&&node.portfolioStates.includes('no-slot')?'no-slot':risk.state}"><b>${escapeHtml(name+(original.quarterStates.length>1?' · '+node.quarter:'')+': '+(riskLabels[risk.label]||risk.label.replace(/^(?:Pre|Post) /i,'')))}</b><p>${escapeHtml(risk.reason)}</p><p class="muted">${escapeHtml([risk.eventDate?`${risk.eventDateField||'Event'}: ${risk.eventDate}`:'',risk.referenceDate?`${risk.referenceField||'Reference'}: ${risk.referenceDate}`:''].filter(Boolean).join(' · '))}</p></div>`).join('')}</div>`;
  if(updateDetail)id('fidSummary').innerHTML=summary;
  else dialog.innerHTML=`<div class="section-head"><div><h1 id="fidDetailTitle"><span id="fidPersonalMark" class="fid-personal-mark"></span>FID ${escapeHtml(fid)} <a class="fid-action" href="https://pcazcrpap02.fremont.lamrc.net:8443/sap(bD1lbiZjPTI1MCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm?saprole=ZSLSDIR&amp;crm-object-type=BT111_OPPT&amp;crm-object-action=B&amp;crm-object-keyname=OBJECT_ID&amp;crm-object-value=${encodeURIComponent(fid)}" target="_blank" rel="noopener noreferrer">CRM ↗</a><span id="fidHeaderStatus" class="fid-header-status" role="group" aria-label="FID Status &amp; Risks"></span></h1><span class="muted">${node.inPopulation?'Commit portfolio':'Observed in source data'}</span></div><a class="fid-action" href="#" data-export-fid="${escapeHtml(fid)}">Export FID 360°</a><button class="fid-close" type="button" data-close-fid aria-label="Close FID 360" title="Close">×</button></div><div id="fidSummary">${summary}</div><section id="fidNotePanel" class="fid-note-panel"></section><div id="fidNsrTable"></div><div id="fidCtlTable"></div><div id="fidNceTable"></div>`;
  id('fidHeaderStatus').innerHTML=fidDetailStatus(node).map(([label,kind,tip])=>pill(label,kind,tip)).join('');
  paintFidPersonal(fid);
  renderNsrDetail('fidNsrTable',node.nsrs,{title:'📋 NSR History',file:'FID_'+fid+'_NSR',defaultSort:'Submit Date',rowRisk:nsr=>nsr.eligible&&nsr.lifecycle==='open'?'pending-nsr':''},updateDetail,changedNsr);
  if(!updateDetail){table('fidCtlTable',ctlColumns.filter(([label])=>label!=='FID'),node.ctls,{title:'🔄 Slot / CTL History',file:'FID_'+fid+'_CTL',rowRisk:ctl=>lateSlotIds.has(ctl.id)?'late-slot':lateCtlIds.has(ctl.id)?'late-ctl':ctl.pending?'pending-config':ctl.pendingPullPush?'pending-pull-push':'',widths:{CTL:105,Stage:140,Change:90,'Created Date':115,'Elapsed Days':110,'Sales Ops Approval':135,'RD At Create':125,'MCSD@Create':125,'SRD @ SOPs Approval':165,'MCSD At SOps Approval':175,'TSD @ SOPs Approval':165,'SRD @ Current':135,'MCSD @ Current':135,'Config Frz':100,'Mfg Upgrade Type':155,Initiator:135,Customer:160,'Quote ID':105,'NSR Reference':150,'Root Cause':180,'CTL Notes':220,'Reason for Chg':160,'Sub Reason':160,'Sales OPs Approver':170}});table('fidNceTable',nceColumns,node.nces,{title:'🔎 NCE / FA History',file:'FID_'+fid+'_NCE',widths:{NCE:90,State:80,Title:140,Impact:75,Opened:95,'Damage Group':95,'Damage Code':90,'Part Number':105,'Part Description':140}})}
  if(!dialog.open)dialog.showModal();
}
function exportFid(fid){
  const original=graph.fids.get(fid),node=graph.projectFid(original,queryFilters().quarter)||graph.projectFid(original),summary=fidDetailRows([node])[0],elapsedDays=ctlColumns.find(([label])=>label==='Elapsed Days')[1];
  const quarters=fidQuarterRows(original);
  summary['Assessment quarters']=node.quarter;summary.Quarter=quarters.map(state=>state.quarter||'[Unknown]').join('\n');summary['Quarter assessments']=quarters.map(state=>(state.quarter||'[Unknown]')+': '+fidQuarterStatus(state)).join('\n');
  if(quarters.length>1)for(const [label,key]of [['Sales Ops RD','rd'],['MCSD','mcsd']])summary[label]=quarters.map(state=>state.dates[key]||'—').join('\n');
  summary['Fab Name']=node.fab;summary['Business Unit']=node.businessUnit;summary['Sales Admin']=node.salesAdmin;summary['Owner rule']='Slot-derived Account membership → Slot / Config CTL counts → tied candidates\' NSR counts (all eligible CTL counts are 0: NSR fallback within Account) → current Slot submitter';
  summary['Pending NSR']=node.post.nsrs.filter(nsr=>includeIgnored||!nsr.ignored).length;summary['Completed 5D NSR']=node.nsrs.filter(nsr=>nsr.eligible&&nsr.lifecycle==='completed-5d').length;summary['Completed NSR but no CTL']=node.nsrs.filter(nsr=>nsr.eligible&&nsr.lifecycle==='completed-unlinked').length;
  exportWorkbook('FID_'+fid,{FID:[summary],NSR:exportTableRows(tables.get('fidNsrTable')),NSR_Source:node.nsrs.flatMap(nsr=>nsr.rows||[]),CTL:node.ctls.flatMap(ctl=>{const days=elapsedDays(ctl);return (ctl.rows||ctl.rawRows||[]).map(row=>({...row,'Elapsed Days':days}))}),NCE:node.nces.flatMap(nce=>nce.rows||nce.rawRows||[]),FID_Source:node.rawRows||[node.record]})
}
function exportSettings(){
  try{
    const settings={format:'ToolkitKyra Settings',version:1,nsrMarks:currentMarks('nsr'),fidMarks:currentMarks('fid'),fidNotes,accountMapping:accountMapping.getMap()};
    const url=URL.createObjectURL(new Blob([JSON.stringify(settings,null,2)],{type:'application/json;charset=utf-8'})),anchor=document.createElement('a');
    anchor.href=url;anchor.download='ToolkitKyra_Settings_'+nsrData.localDate(new Date())+'.json';document.body.append(anchor);anchor.click();anchor.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
  }catch(error){toast('Could not export settings: '+error.message)}
}
async function importSettings(file){
  try{
    const settings=JSON.parse(await file.text()),isMap=value=>value&&typeof value==='object'&&!Array.isArray(value);
    if(!isMap(settings)||settings.format!=='ToolkitKyra Settings'||settings.version!==1||!['nsrMarks','fidMarks','fidNotes','accountMapping'].every(name=>isMap(settings[name])))throw new Error('Invalid ToolkitKyra settings file.');
    for(const kind of ['nsr','fid'])if(Object.entries(settings[kind+'Marks']).some(([number,mark])=>!markKinds[kind].pattern.test(number)||!['ignore','important'].includes(mark)))throw new Error('Invalid '+kind.toUpperCase()+' marks.');
    if(Object.entries(settings.fidNotes).some(([fid,note])=>!markKinds.fid.pattern.test(fid)||typeof note!=='string'))throw new Error('Invalid FID notes.');
    const notes=Object.fromEntries(Object.entries(settings.fidNotes).filter(([,note])=>note.trim()).map(([fid,note])=>[fid,note.replace(/\r\n?/g,'\n')]));
    const mapping=accountMapping.normalize(settings.accountMapping),values={
      [NSR_MARKS_KEY]:JSON.stringify(settings.nsrMarks),[FID_MARKS_KEY]:JSON.stringify(settings.fidMarks),[FID_NOTES_KEY]:JSON.stringify(notes),
      [accountMapping.key]:JSON.stringify({base:accountMapping.defaults,mapping})
    };
    const previous=Object.fromEntries(Object.keys(values).map(key=>[key,GM_getValue(key,null)])),written=[];
    try{for(const [key,value]of Object.entries(values)){GM_setValue(key,value);written.push(key)}}
    catch(error){for(const key of written.reverse())GM_setValue(key,previous[key]);throw error}
    for(const kind of ['nsr','fid']){saveMarks(kind,settings[kind+'Marks'],false);markDrafts[kind]=null}
    fidNotes=notes;mappingDraft=null;accountMapping.reload();recalculate();toast('Settings imported');
  }catch(error){toast('Could not import settings: '+error.message)}
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
const inferenceTooltip=document.createElement('div');inferenceTooltip.className='inference-tooltip';inferenceTooltip.setAttribute('role','tooltip');inferenceTooltip.popover='manual';document.body.append(inferenceTooltip);
document.addEventListener('pointerover',event=>{
  const help=event.target.closest('.inference-help');if(!help)return;
  inferenceTooltip.textContent=help.getAttribute('aria-label');inferenceTooltip.showPopover();
  const rect=help.getBoundingClientRect();
  inferenceTooltip.style.left=Math.max(8,Math.min(rect.left,innerWidth-inferenceTooltip.offsetWidth-8))+'px';
  inferenceTooltip.style.top=Math.max(8,Math.min(rect.bottom+7,innerHeight-inferenceTooltip.offsetHeight-8))+'px';
});
document.addEventListener('pointerout',event=>{if(event.target.closest('.inference-help'))inferenceTooltip.hidePopover()});
// Keep Chosen focus from scrolling an option away between mouse down and mouse up.
document.addEventListener('mousedown',event=>{const option=event.target.closest('.side .chosen-results .active-result');if(event.button!==0||!option)return;option.closest('.chosen-container').querySelector('.chosen-search-input')?.focus({preventScroll:true});event.preventDefault()},true);
document.addEventListener('click',event=>{
  const label=event.target.closest('.fc-filter-head>label[for]');if(label&&id(label.htmlFor)?.multiple&&jQuery.fn.chosen){event.preventDefault();jQuery(id(label.htmlFor)).trigger('chosen:open');return}
  document.querySelectorAll('.fc-source-help[open]').forEach(help=>{if(!help.contains(event.target))help.open=false});const button=event.target.closest('button,a[data-export-fid]');if(!button)return;
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
    target.animate([{backgroundColor:'#fff0ad',boxShadow:'0 0 0 3px #efb647'},{}],{duration:matchMedia('(prefers-reduced-motion: reduce)').matches?0:600,easing:'ease-out'});return;
  }
  if(button.dataset.importSource){id('fcFile-'+button.dataset.importSource).click();return}
  if(button.dataset.sourceAction){const action=button.dataset.sourceAction;if(action==='bobj-update')loadBobj(true);else if(action==='ctl-bobj-update')loadBobj(true,'ctl');else if(action==='nsr-open'){setNsrRange(nsrRange);id('fcNsrDialog').showModal()}return}
  if(button.hasAttribute('data-nce-signal')){filterValues.signal=['nce'];applyFilters();return}
  if(button.dataset.noteFid){editFidNote(button.dataset.noteFid,button);return}
  if(button.hasAttribute('data-cancel-fid-note')){closeFidNote();return}
  if(button.dataset.ignoreNsr||button.dataset.ignoreFid){setRecordMark(button.dataset.ignoreFid?'fid':'nsr',button.dataset.ignoreFid||button.dataset.ignoreNsr);return}
  if(button.dataset.view){showView(button.dataset.view);return}
  if(button.dataset.fid){showFid(button.dataset.fid);return}
  if(button.hasAttribute('data-mapping-rule')){showView('rules');id('fabAccountMapping').scrollIntoView({block:'start'});return}
  if(button.hasAttribute('data-close-fid')){id('fidDetail').close();return}
  if(button.hasAttribute('data-toggle-pull-push')){
    const state=tables.get('fidCtlTable');state.local.showPullPush=!state.local.showPullPush;button.textContent=(state.local.showPullPush?'Hide':'Show')+' Pull/Push';button.setAttribute('aria-pressed',String(state.local.showPullPush));paintTable(state);
    // Filtering already rendered the new height; a second resize layout would reset column widths.
    const wrap=state.grid.element;state.size=wrap.clientWidth+':'+wrap.clientHeight;return;
  }
  if(button.dataset.exportFid){event.preventDefault();exportFid(button.dataset.exportFid);return}
  if(button.hasAttribute('data-reset-ctl-type')){changeFilter('ctl.change',[...defaultCtlTypes]);return}
  if(button.hasAttribute('data-reset-quarter')){changeFilter('quarter',defaultQuarters());return}
  if(button.dataset.clearFilter||button.dataset.selectAll){const name=button.dataset.clearFilter||button.dataset.selectAll,definition=filterDefs.find(def=>def[0]===name);const values=button.dataset.selectAll?[...id(name).options].map(option=>option.value).filter(value=>name!=='fNsrTaskStatus'||!Object.hasOwn(nsrCompletedTasks,value)):[];changeFilter(definition[2],values);return}
  if(button.dataset.fidCell||button.dataset.fidStatus||button.dataset.fidRegion){
    const key=button.dataset.fidStatus||button.dataset.fidRegion;
    if(key==='all'){changeFilter({fidStatus:[],signal:(filterValues.signal||[]).filter(value=>!fidCauseKeys.includes(value))});return}
    selectFidStatus(button.dataset.fidValue||button.dataset.fidCell);return;
  }
  if(button.dataset.ctlStatus){selectTaskFilter('ctl',button.dataset.ctlStatus,button.dataset.ctlTypes?.split(','));return}
  if(button.dataset.nsrGroup||button.dataset.nsrLifecycle){const key=button.dataset.nsrLifecycle||button.dataset.nsrGroup;selectTaskFilter('nsr',key==='open'?'workflow':key);return}
  if(button.dataset.ctlPeriod){changeFilter('ctl.created',ctlDateRange(button.dataset.ctlPeriod));return}
  if(button.dataset.nsrYear){const year=button.dataset.nsrYear;changeFilter('nsr.submitted',[year+'-01-01',year+'-12-31']);return}
  if(button.dataset.fidYear){const values=[...id('fQuarter').options].map(option=>option.value).filter(value=>quarterYear(value)===button.dataset.fidYear),selected=filterValues.quarter||[],add=values.some(value=>!selected.includes(value));changeFilter('quarter',add?[...new Set([...selected,...values])]:selected.filter(value=>!values.includes(value)));return}
  if(button.dataset.groupPath){const path=button.dataset.groupPath,value=button.dataset.groupValue,selected=filterValues[path]||[];changeFilter(path,selected.includes(value)?selected.filter(item=>item!==value):[...selected,value]);return}
  if(button.dataset.pivotFilter){const path=filterDefs.find(def=>def[0]===button.dataset.pivotFilter)[2];changeFilter(path,[button.dataset.pivotValue]);return}

  if(button.dataset.cycleNsr){
    if(!unsafeWindow.confirm(`Query Cycle Time and Workflow for NSR ${button.dataset.cycleNsr}?`))return;
    if(id('fidDetail').open)id('fidDetail').close();showView('cycle');cycle.query([button.dataset.cycleNsr]);return;
  }
  if(button.dataset.tableCycle){const state=tables.get(button.dataset.tableCycle),rows=tableResult(state);if(id('fidDetail').open)id('fidDetail').close();showView('cycle');cycle.query(rows.map(row=>row.number));return}
  if(button.dataset.columnClear){tables.get(button.dataset.columnClear).grid.clearHeaderFilter();return}
  if(button.dataset.columnRemove){tables.get(button.dataset.columnRemove).grid.setHeaderFilterValue(button.dataset.field,'');return}
  if(button.dataset.tableCopy){const state=tables.get(button.dataset.tableCopy);if(state.options.onCopy)state.options.onCopy();else copyRows(exportTableRows(state,true));return}
  if(button.dataset.summarySort){
    const key=button.dataset.summarySort;summarySort={key,dir:summarySort.key===key?-summarySort.dir:-1};renderSummaryMetrics(id('view-summary'));
    id('summaryChart').querySelector('[data-summary-sort="'+key+'"]').focus({preventScroll:true});return;
  }
  if(button.dataset.summaryExport){
    const percentMetrics=summaryMetricDefs.filter(([,,base])=>base).map(([label])=>label);
    if(button.dataset.summaryExport==='copy')copyRows(summaryReport.map(([label,...values])=>[label,...values.map(value=>typeof value==='number'&&percentMetrics.includes(label)?(100*value).toFixed(1)+'%':value)]));
    else exportWorkbook('Summary_Metrics_'+summaryGroup,{Metrics:summaryReport},Object.fromEntries(percentMetrics.map(label=>[label,'0.0%'])));
    return;
  }
  if(button.dataset.tableExport){const state=tables.get(button.dataset.tableExport);if(state.options.onExport)state.options.onExport();else exportWorkbook(state.options.file,{Detail:exportTableRows(state),...(state.options.cycle?{Source_Rows:tableResult(state).flatMap(nsr=>nsr.rows||nsr.rawRows||[])}:{})});return}
  if(button.hasAttribute('data-clear-search')){id('fSearch').value='';applyFilters();return}
  if(button.hasAttribute('data-clear-signal')){delete filterValues.signal;applyFilters();return}
  if(button.hasAttribute('data-toast-close')){id('fcToast').hidden=true;return}
  if(button.hasAttribute('data-save-nsr-marks')||button.hasAttribute('data-save-fid-marks')){saveMarkLists(button.hasAttribute('data-save-fid-marks')?'fid':'nsr');return}
  if(button.hasAttribute('data-save-mapping')){const value=Object.fromEntries([...document.querySelectorAll('[data-account-group]')].map(textarea=>[textarea.dataset.accountGroup,textarea.value])),result=accountMapping.save(value);if(result.ok){mappingDraft=null;recalculate();toast('Account mapping saved')}else{const notice=id('agNotice');notice.hidden=false;notice.textContent=result.error}return}
  if(button.hasAttribute('data-reset-mapping')){if(!unsafeWindow.confirm('Restore the default Fab / Account mapping?'))return;const result=accountMapping.reset();if(result.ok){mappingDraft=null;recalculate()}else toast(result.error);return}
  if(button.hasAttribute('data-import-settings')){id('fcSettingsFile').click();return}
  if(button.hasAttribute('data-export-settings')){exportSettings();return}
  if(button.hasAttribute('data-export-data-sources')){exportAllDataSources();return}
  if(button.id==='fcRefresh')restoreSources(true);else if(button.id==='fcReset'){
    includeIgnored=true;clearTimeout(searchTimer);for(const key of Object.keys(filterValues))delete filterValues[key];
    if(store.raw.fcid.length)filterValues.quarter=defaultQuarters();
    if(store.meta.ctl||store.meta.ctlNsr)filterValues['ctl.change']=[...defaultCtlTypes];
    id('fSearch').value='';for(const local of localTables.values()){local.scroll=0;local.marks=null;local.headerFilters=[]}for(const state of tables.values())if(state.ready)state.grid.clearHeaderFilter();buildFilters();applyFilters();
  }
});
document.addEventListener('keydown',event=>{
  if(event.ctrlKey&&event.shiftKey&&event.altKey&&event.code==='KeyD'){
    event.preventDefault();if(event.repeat||document.querySelector('[data-view="summary"]').hidden)return;
    summaryGroup=summaryGroup==='owner'?summaryPreviousGroup:'owner';rendered.delete('summary');showView('summary');return;
  }
  if((event.key==='Enter'||event.key===' ')&&event.target.matches('.fc-table-wrap .tabulator-col')){event.preventDefault();event.target.click()}
  if(event.key==='Escape')document.querySelectorAll('.fc-source-help[open]').forEach(help=>help.open=false);
});
document.addEventListener('input',event=>{const kind=event.target.dataset.nsrMarks?'nsr':event.target.dataset.fidMarks?'fid':'';if(kind){markDrafts[kind]=Object.fromEntries([...document.querySelectorAll(`[data-${kind}-marks]`)].map(textarea=>[textarea.dataset[kind+'Marks'],textarea.value]));id(markKinds[kind].status).textContent=''}if(event.target.dataset.accountGroup){mappingDraft??=accountMapping.getMap();mappingDraft[event.target.dataset.accountGroup]=event.target.value}});
document.addEventListener('change',event=>{
  const input=event.target;
  if(input.dataset.tableMarkFilter){
    const state=tables.get(input.dataset.tableMarkFilter),marks=input.dataset.mark==='all'?(input.checked?null:[]):[...markPopup.querySelectorAll('input[data-mark]:checked')].map(input=>input.dataset.mark).filter(mark=>mark!=='all');
    state.local.marks=marks?.length===markOptions.length?null:marks;paintTable(state);
  }else if(input.dataset.excludeIgnored){
    const view=input.dataset.excludeIgnored,focused=document.activeElement===input;includeIgnored=!input.checked;applyFilters(true);
    if(focused)id('view-'+view).querySelector('[data-exclude-ignored]').focus({preventScroll:true});
  }else if(input.id==='fcSettingsFile'&&input.files?.[0]){importSettings(input.files[0]);input.value=''}else if(input.dataset.sourceFile&&input.files?.[0])importSource(input.files[0],input.dataset.sourceFile);
});
filterDefs.forEach(([name,,path])=>{if(dateFilters.includes(path)){id(name).addEventListener('change',event=>{const [from,to]=id(name).querySelectorAll('input');from.setCustomValidity('');to.setCustomValidity('');if(from.value&&to.value&&from.value>to.value){event.target.setCustomValidity('From date must not be later than To date.');event.target.reportValidity();return}changeFilter(path,from.value||to.value?[from.value,to.value]:[])});return}const position=()=>requestAnimationFrame(()=>positionFilterChoices(id(name)));jQuery(id(name)).on('change',(event,change)=>{const values=[...id(name).selectedOptions].map(option=>option.value);if(path==='riskType'||path==='fidStatus'||path==='nsr.taskStatus'||path.startsWith('ctl.'))changeFilter(path,path==='fidStatus'?fidSelection.groups[change?.selected]||values:values);else{filterValues[path]=values;applyFilters()}position()}).on('chosen:showing_dropdown',position);if(jQuery.fn.chosen)jQuery(id(name)).chosen({width:'100%',search_contains:true,hide_results_on_select:false,placeholder_text_multiple:'All',display_selected_options:false})});
columnPopup.onsubmit=columnPopup.onreset=event=>{
  event.preventDefault();const column=columnPopup.column,value=event.type==='reset'?'':columnInput.value;
  if(column&&value!==(column.getHeaderFilterValue()??''))column.setHeaderFilterValue(value);columnPopup.hidePopover();
};
columnPopup.ontoggle=event=>{if(event.newState==='open')columnInput.select();else columnPopup.column=null};
markPopup.onbeforetoggle=event=>{if(event.newState==='closed'){markPopup.state=null;markPopup.trigger=null}};
document.addEventListener('scroll',event=>{if(columnPopup.column&&!columnPopup.contains(event.target))columnPopup.hidePopover();if(markPopup.state&&!markPopup.contains(event.target)){if(markPopup.position!==markMenuPosition(markPopup.state))markPopup.hidePopover()}},true);window.addEventListener('resize',()=>{columnPopup.hidePopover();markPopup.hidePopover()});
id('fidDetail').addEventListener('close',()=>{id('fidNotePanel')?.querySelector('form')?.remove();highlightLinkedHistory('fidCtlTable');highlightLinkedHistory('fidNsrTable')});
id('fidNoteDialog').addEventListener('close',()=>id('fidNoteDialog').replaceChildren());
document.addEventListener('submit',event=>{if(event.target.id==='fidNoteForm'){event.preventDefault();saveFidNote()}});
id('fSearch').addEventListener('input',()=>{syncFidLookup();clearTimeout(searchTimer);searchTimer=setTimeout(applyFilters,180)});
id('fidLookup').addEventListener('submit',event=>{event.preventDefault();const fid=syncFidLookup();if(fid)showFid(fid)});
id('fcNsrFrom').addEventListener('input',()=>id('fcNsrTo').min=id('fcNsrFrom').value);id('fcNsrCancel').addEventListener('click',()=>id('fcNsrDialog').close());id('fcNsrForm').addEventListener('submit',event=>{event.preventDefault();if(store.busy.has('nsr'))return;id('fcNsrDialog').close();loadSharedNsr(false)});

function normalizeNsrRows(rows){return rows.map(raw=>{const row=nsrRecords.canonicalRow(raw);row['REF FCID']||=text(raw['Reference FID']??raw.FID);row['Submit Date']=sourceDate(row['Submit Date'])||row['Submit Date'];return {...raw,...row,'Reference Fcid':row['REF FCID'],'Workflow Status':raw['Workflow Status']||row['NSR Status']||row['Task Name'],'Task Status':nsrRecords.deriveTaskStatus(row)}}).filter(row=>row['Reference Fcid']||row['NSR#'])}
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
  for(const row of rows){let date='';for(const field of config.fields){date=sourceDate(row[field]);if(date)break}if(!date)continue; if(!dateFrom||date<dateFrom)dateFrom=date;if(!dateTo||date>dateTo)dateTo=date;}
  return {dateFrom,dateTo,dateField:config.label};
}
function sourceMessage(key,message,kind='ok',transport=false){
  const target=id('fcSource-'+key),meta=store.meta[key],hasRange=Boolean(meta?.dateFrom&&meta?.dateTo);
  const cutoff=new Date();cutoff.setDate(cutoff.getDate()-3);const stale=hasRange&&meta.dateTo<=nsrData.localDate(cutoff);
  const label=kind==='loading'?(/updat/i.test(message)?'Updating…':'Loading…'):meta?(hasRange?meta.dateFrom+' ~ '+meta.dateTo:'Loaded · no dates'):'Unavailable';
  target.textContent=label;target.className='fc-source-state '+(kind==='loading'?'loading':meta?'ok':'')+(hasRange?' has-range':'')+(kind!=='loading'&&stale?' stale':'')+(kind==='error'?' error':'');
  target.title=meta?.cached&&kind!=='loading'?['Using cache · '+new Date(meta.savedAt).toLocaleString('en-GB'),meta.modifiedAt?(BOBJ_SOURCES[key]?'Workbook updated · ':'Source updated · ')+new Date(meta.modifiedAt).toLocaleString('en-GB'):'',kind==='error'?message:meta.cacheReason].filter(Boolean).join('\n'):kind==='error'?message:hasRange?'Data range · '+meta.dateField:message;
  target.setAttribute('aria-label',label);target.tabIndex=kind==='error'||meta?.cached?0:-1;
  if(BOBJ_SOURCES[key]){const help=id('fcHelp-'+key);help.hidden=!transport;if(help.hidden)help.open=false}
}
function actionMessage(message){id('fcActionStatus').textContent=message}
function sourceMeta(key){const meta=store.meta[key],target=id('fcMeta-'+key);target.textContent=meta?text(meta.file).split(/[\\/]/).pop():'';target.title=meta?.file||'';id('fcSourceCount').textContent=Object.keys(store.meta).length+' / 5'}
function sourceBusy(key,on){
  on?store.busy.add(key):store.busy.delete(key); document.querySelectorAll(`[data-source-file="${key}"]`).forEach(input=>input.disabled=on); id('fcImport-'+key).disabled=on;
  const actions=key==='fcid'?['bobj-update']:key==='ctl'?['ctl-bobj-update']:key==='nsr'?['nsr-open']:[]; actions.forEach(action=>document.querySelectorAll(`[data-source-action="${action}"]`).forEach(button=>button.disabled=on));
  if(key==='nsr')id('fcNsrLoad').disabled=on;
  id('fcRefresh').disabled=store.refreshing||store.busy.size>0;
}
function startSource(key){store.controllers[key]?.abort();store.runs[key]=(store.runs[key]||0)+1;sourceBusy(key,true);return store.runs[key];}
async function applySource(key,rows,meta,render=true){
  if(['ctl','ctlNsr'].includes(key)&&!store.meta.ctl&&!store.meta.ctlNsr)filterValues['ctl.change']=[...defaultCtlTypes];
  if(key==='fcid'&&id('workspaceNav').querySelector('[data-view="portfolio"]').hidden){
    filterValues.quarter=defaultQuarters();
    id('workspaceNav').querySelectorAll('[data-view]').forEach(button=>button.hidden=false);
    showView('portfolio');
  }
  store.raw[key]=rows;store.meta[key]={...meta,...(meta.cached&&meta.dateField?{}:sourceRange(rows,key,meta)),rows:rows.length,savedAt:meta.savedAt||Date.now()};store.dirty=true;sourceMessage(key,'Loaded');sourceMeta(key);if(render&&!store.refreshing)recalculate();
}


async function downloadHostedArchive(){
  const failures=[];
  for(const url of HOSTED_ARCHIVE_URLS){
    try{
      return await new Promise((resolve,reject)=>{
        GM_xmlhttpRequest({
          method:'GET',url:url+'?t='+Date.now(),responseType:'arraybuffer',timeout:30000,nocache:true,anonymous:true,
          onload:response=>{if(response.status>=200&&response.status<300&&response.response?.byteLength)resolve(response.response);else reject(new Error('HTTP '+response.status+(response.statusText?' '+response.statusText:'')+(response.status>=200&&response.status<300?' (empty response).':'.')));},
          onerror:response=>reject(new Error((response?.status?'HTTP '+response.status+'. ':'')+(response?.error?.message||response?.error||response?.message||response?.statusText||'Network or Tampermonkey permission error.'))),
          ontimeout:()=>reject(new Error('Data download timed out (30 s).')),
          onabort:()=>reject(new Error('Data download was cancelled.'))
        });
      });
    }catch(error){failures.push(new URL(url).hostname+': '+error.message)}
  }
  throw new Error('All data download sources failed.\n'+failures.join('\n')+'\nCheck Tampermonkey connection permissions and network/proxy access.');
}
async function readSourceCache(key){
  const cache=await nsrData.readDataset(cacheKeys[key]);
  if(key==='nsr'&&cache?.schema===1&&!cache.columns){
    if(!Array.isArray(cache.rows)||!cache.rows.length||!nsrData.validSapDate(cache.from)||!nsrData.validSapDate(cache.to))return null;
    return {rows:normalizeNsrRows(cache.rows),meta:{file:'NSR cache',kind:'sap',sheet:'SearchNsrSet',from:cache.from,to:cache.to,pages:cache.pages,savedAt:cache.savedAt||Date.now()}};
  }
  if(cache?.schema!==2&&(!BOBJ_SOURCES[key]||cache?.schema!==1)||typeof cache.meta?.file!=='string'||!Number.isFinite(cache.meta.savedAt)||cache.meta.savedAt<=0||
    !Array.isArray(cache.columns)||cache.columns.some(column=>typeof column!=='string')||new Set(cache.columns).size!==cache.columns.length||
    !Array.isArray(cache.rows)||cache.rows.some(row=>!Array.isArray(row)||row.length!==cache.columns.length)||
    cache.rows.length&&[key==='nsr'?'NSR#':sources[key].fields.fid.label,sources[key].fields.number?.label].filter(Boolean).some(field=>!cache.columns.includes(field))||
    key==='ctlNsr'&&(!Array.isArray(cache.meta.rowTimes)||cache.meta.rowTimes.length!==cache.rows.length||cache.meta.rowTimes.some(time=>!Number.isFinite(time)||time<0)))return null;
  return {rows:cache.rows.map(row=>Object.fromEntries(cache.columns.map((column,index)=>[column,row[index]]).filter(([,value])=>value!==undefined))),meta:{...cache.meta,fields:cache.columns}};
}
async function restoreSourceCache(key,run,render=true){
  if(store.meta[key])return;
  const cache=await readSourceCache(key);
  if(cache&&run===store.runs[key])await applySource(key,cache.rows,{...cache.meta,cached:true,cacheReason:''},render);
}
async function sourceFailure(key,run,message,transport=false){
  if(run!==store.runs[key])return;
  await restoreSourceCache(key,run).catch(()=>{});
  if(run!==store.runs[key])return;
  const meta=store.meta[key];if(meta){meta.cached=true;meta.cacheReason=message}
  sourceMessage(key,message,meta?'ok':'error',transport);
}
async function saveSourceCache(key,run){
  const rows=store.raw[key],{fields,...meta}=store.meta[key],names=new Set(key==='ctlNsr'?fields:[]);
  if(meta.cacheReadError){sourceMessage(key,'Cache not saved: '+meta.cacheReadError,'error');return}
  for(const row of key==='nsr'?rows:rows.slice(0,1))for(const column of Object.keys(row))names.add(column);
  const columns=[...names];
  try{await nsrData.writeDataset({key:cacheKeys[key],schema:2,columns,rows:rows.map(row=>columns.map(column=>row[column])),meta})}
  catch(error){if(run===store.runs[key])sourceMessage(key,'Cache not saved: '+error.message,'error')}
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
async function loadHostedSources(keys=Object.keys(HOSTED_SOURCE_PREFIXES),force=false){
  const runs=Object.fromEntries(keys.map(key=>[key,startSource(key)])),today=nsrData.localDate(new Date());keys.forEach(key=>sourceMessage(key,'Loading data…','loading'));
  const cacheErrors={};let pending=keys,stage='SAP access';
  const failed=(key,error)=>sourceFailure(key,runs[key],stage+' failed: '+(error.userMessage||error.message)+(stage==='SAP access'&&!store.meta[key]?'\nSign in to Fiori.':''));
  try{
    await Promise.all(keys.map(key=>restoreSourceCache(key,runs[key],false).catch(error=>{cacheErrors[key]=error.message})));
    pending=keys.filter(key=>runs[key]===store.runs[key]&&(force||store.meta[key]?.hostedDay!==today));
    keys.filter(key=>runs[key]===store.runs[key]&&!pending.includes(key)).forEach(key=>sourceMessage(key,'Loaded'));
    if(!pending.length)return;
    pending.forEach(key=>sourceMessage(key,'Loading data…','loading'));
    const response=await unsafeWindow.fetch(HOSTED_PASSWORD_URL,{credentials:'same-origin',cache:'no-store',signal:unsafeWindow.AbortSignal.timeout(30000),headers:{Accept:'application/atom+xml,application/xml,text/xml'}});
    if(!response.ok)throw new Error('HTTP '+response.status);
    if((response.headers.get('Content-Type')||'').toLowerCase().includes('text/html'))throw new Error('SAP returned a sign-in page.');
    const doc=new DOMParser().parseFromString(await response.text(),'application/xml');
    if(doc.getElementsByTagName('parsererror').length)throw new Error('SAP returned invalid XML.');
    const properties=[...doc.getElementsByTagNameNS('*','properties')].find(node=>node.getElementsByTagNameNS('*','Userid')[0]?.textContent.trim()==='LONGKY');
    const archivePassword=properties?.getElementsByTagNameNS('*','FullName')[0]?.textContent.replace(/\s+/g,'')||'';
    if(!archivePassword)throw new Error('SAP returned no password.');
    if(pending.every(key=>runs[key]!==store.runs[key]))return;
    stage='GitHub download';
    let bytes=await downloadHostedArchive();if(pending.every(key=>runs[key]!==store.runs[key]))return;
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
          const entry=entries.find(item=>!item.directory&&item.filename.split(/[\\/]/).pop().toLowerCase().startsWith(HOSTED_SOURCE_PREFIXES[key]));
          if(!entry)throw new Error('The data file is missing.');
          const filename=entry.filename.split(/[\\/]/).pop(),modifiedAt=entry.lastModDate?.getTime();
          if(!Number.isFinite(modifiedAt)||modifiedAt<=0)throw new Error('The data file timestamp is invalid.');
          const data=await entry.getData(new zip.Uint8ArrayWriter(),{checkSignature:true}),parsed=await report.readFile(new File([data],filename),key);
          if(runs[key]!==store.runs[key])continue;
          const {rows,...recognition}=parsed,merged=key==='ctlNsr'?mergeCtlNsr(rows,modifiedAt):{rows};
          await applySource(key,merged.rows,{file:filename,kind:'github',...recognition,modifiedAt,hostedDay:today,savedAt:Date.now(),cacheReadError:cacheErrors[key]||'',...(merged.rowTimes?{rowTimes:merged.rowTimes,fields:merged.fields}:{})},false);
          await saveSourceCache(key,runs[key]);
        }catch(error){await failed(key,error)}
      }
    }finally{await reader.close()}
  }catch(error){await Promise.all(pending.map(key=>failed(key,error)))}finally{if(store.dirty&&!store.refreshing)recalculate();keys.forEach(key=>{if(runs[key]===store.runs[key])sourceBusy(key,false)})}
}
async function importSource(file,key){
  const run=startSource(key);sourceMessage(key,'Reading '+file.name+'…','loading');
  try{
    const parsed=await report.readFile(file,key);let range={savedAt:Date.now()};
    if(key==='nsr'){parsed.rows=normalizeNsrRows(parsed.rows);if(nsrData.validSapDate(parsed.from)&&nsrData.validSapDate(parsed.to))range={...range,from:parsed.from,to:parsed.to};delete parsed.from;delete parsed.to}
    if(run!==store.runs[key])return;
    if(key==='ctlNsr'){
      const savedAt=Date.now(),meta=store.meta[key],merged=mergeCtlNsr(parsed.rows,savedAt,true);parsed.rows=merged.rows;
      range={rowTimes:merged.rowTimes,fields:merged.fields,hostedDay:meta?.hostedDay||'',modifiedAt:meta?.modifiedAt||0,savedAt,cacheReadError:meta?.cacheReadError||''};
    }
    if(key==='nce')range.hostedDay=nsrData.localDate(new Date());
    const {rows,...recognition}=parsed;await applySource(key,rows,{file:file.name,kind:'file',...recognition,...range});
    await saveSourceCache(key,run);
  }catch(error){if(run===store.runs[key])sourceMessage(key,(key==='ctlNsr'?'Append':'Import')+' failed: '+(error.userMessage||error.message)+(store.meta[key]?' Previous data retained.':''),'error')}finally{if(run===store.runs[key]){sourceBusy(key,false);
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
    let modified=0;
    try{const data=await bridgeFileRequest('stat',{path:status.outputPath},10000);modified=Number(data.entries?.[0]?.lastModified)||0}catch(_){/* File dates are optional; keep successfully read data. */}
    if(run!==store.runs[key])return;
    store.meta[key].modifiedAt=modified;
    await saveSourceCache(key,run);
    if(run!==store.runs[key])return;
    if(!update){
      const date=modified>0?nsrData.localDate(new Date(modified)):'';
      if(date&&date<nsrData.localDate(new Date())&&unsafeWindow.confirm(`${sources[key].label} data is not up to date (last updated: ${date}). Update now? This may take 2-3 minutes.`))setTimeout(()=>document.querySelector(`[data-source-action="${key==='ctl'?'ctl-bobj-update':'bobj-update'}"]`).click(),0);
    }
  }catch(error){await sourceFailure(key,run,error.message,error.transport)}
  finally{if(run===store.runs[key])sourceBusy(key,false)}
}
function setNsrRange(range){nsrRange=range;id('fcNsrFrom').value=range.from;id('fcNsrTo').value=range.to;id('fcNsrTo').min=range.from;id('fcNsrFrom').max=id('fcNsrTo').max=nsrData.localDate(new Date())}
async function loadSharedNsr(useCache=true){
  const key='nsr',run=startSource(key);sourceMessage(key,useCache?'Checking NSR cache…':'Loading NSR…','loading'); const today=nsrData.localDate(new Date());let range={from:id('fcNsrFrom').value,to:id('fcNsrTo').value};
  try{
    if(useCache){
      await restoreSourceCache(key,run).catch(()=>{});if(run!==store.runs[key])return;
      const meta=store.meta[key];
      if(meta){
        if(meta.from&&meta.to){range={from:meta.from,to:meta.to};setNsrRange(range)}
        if(meta.kind==='file'||meta.to===today)return;
        range={from:meta.from,to:today};
      }else range=nsrData.sapRange();
    }
    if(!nsrData.validSapDate(range.from)||!nsrData.validSapDate(range.to)||range.from>range.to||range.to>today)throw new Error('Choose a valid NSR date range ending on or before today.');
    const controller=new unsafeWindow.AbortController();store.controllers[key]=controller; const result=await nsrData.fetchSapRows(range,controller.signal);if(run!==store.runs[key])return;
    setNsrRange(range);await applySource(key,normalizeNsrRows(result.rows),{file:'SAP live NSR',kind:'sap',sheet:'SearchNsrSet',...range,pages:result.pages,savedAt:Date.now()});await saveSourceCache(key,run);
  }catch(error){await sourceFailure(key,run,'NSR load failed: '+error.message)}
  finally{if(run===store.runs[key]){sourceBusy(key,false);delete store.controllers[key]}}
}
async function restoreSources(reset=false){
  if(store.refreshing||store.busy.size)return;
  store.refreshing=true;const refresh=id('fcRefresh');refresh.disabled=true;refresh.textContent='Refreshing…';
  try{
    const keys=Object.keys(sources);
    if(reset){
      keys.forEach(key=>{store.runs[key]=(store.runs[key]||0)+1;store.raw[key]=[];delete store.meta[key]; sourceMessage(key,'Not loaded','');sourceMeta(key);sourceBusy(key,true);id('fcFile-'+key).value='';}); setNsrRange(nsrData.sapRange());actionMessage('');
        recalculate();
    }
    // Retain hosted merged data and BOBJ fallback caches; discard other legacy source caches.
    try{await nsrData.deleteDatasets([...keys.filter(key=>!HOSTED_SOURCE_PREFIXES[key]).map(key=>CACHE_PREFIX+key),CACHE_PREFIX+'archive'])} catch(error){if(reset)actionMessage('Local source cache could not be reset. '+error.message)} finally{if(reset)keys.forEach(key=>sourceBusy(key,false))}
    if(store.meta.nsr?.from)setNsrRange({from:store.meta.nsr.from,to:store.meta.nsr.to});
    const loads=[];
    for(const key of Object.keys(BOBJ_SOURCES))if(!store.busy.has(key))loads.push(loadBobj(false,key));
    const hostedKeys=Object.keys(HOSTED_SOURCE_PREFIXES).filter(key=>!store.busy.has(key));if(hostedKeys.length)loads.push(loadHostedSources(hostedKeys,reset));
    if(!store.busy.has('nsr')&&(!store.meta.nsr||store.meta.nsr.kind==='sap'))loads.push(loadSharedNsr());
    // Keep per-source progress visible; rebuild the graph once the whole load settles.
    await Promise.all(loads);
  }finally{store.refreshing=false;refresh.disabled=store.busy.size>0;refresh.textContent='Refresh';if(store.dirty||reset)recalculate()}
}


const cycle=createNsrCycle(id('dashboardWorkspace'),{getRows:nsrs=>{
  const selected=new Set(nsrs);
  return nsrCopyRows(graph.nsrs.filter(nsr=>selected.has(nsr.number))).map(row=>({...row,'REF FCID':row.FID}));
},renderTable:(name,columns,rows,options)=>table(name,columns,rows,options,true),copyRows,onNotify:message=>toast(message)});
setNsrRange(nsrRange);recalculate();restoreSources();
}

// Source parsing is independent of dashboard state and does not change source records.
function createFidSourceReaders(sources){
  const weights={
    fcid:{fid:[40,1],commit:[20,1],mcsd:[5,0],customer:[5,0],system:[5,0],targetShip:[5,0],rd:[5,0],account:[5,0]},
    ctl:{fid:[35,1],number:[25,1],change:[10,0],mcsd:[5,0],srd:[5,0],approved:[5,0],summary:[5,0]},
    ctlNsr:{fid:[35,1],number:[25,1],change:[10,0],mcsd:[5,0],srd:[5,0],approved:[5,0],stage:[5,0]},
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
  function bestHeaderFromMatrix(matrix,key,startRow=0){
    let best=null;(matrix||[]).slice(0,50).forEach((values,index)=>{const candidate={...scoreSourceHeader(values,key),headerRow:startRow+index};if(betterHeader(candidate,best))best=candidate;});
    return best||{score:0,missing:profileFor(key).fields.filter(field=>field.required).map(field=>field.label),requiredOk:false,headerRow:startRow};
  }
  function sourceScoreError(key,candidate){const error=new Error(profileFor(key).label+': check file headers'+(candidate?.missing?.length?' ('+candidate.missing.join(', ')+')':''));error.code='LOW_SOURCE_SCORE';error.userMessage=error.message;return error;}
  function validateSourceCandidate(key,candidate){if(!candidate||!candidate.requiredOk||candidate.score<profileFor(key).minimum)throw sourceScoreError(key,candidate);return candidate;}
  function detectWorkbookSource(workbook,key){
    const profile=profileFor(key),candidates=[];
    (workbook.SheetNames||[]).forEach((sheetName,sheetOrder)=>{
      const sheet=workbook.Sheets[sheetName];if(!sheet||!sheet['!ref'])return;
      const fullRange=XLSX.utils.decode_range(sheet['!ref']),previewRange={s:{r:fullRange.s.r,c:fullRange.s.c},e:{r:Math.min(fullRange.e.r,fullRange.s.r+49),c:fullRange.e.c}};
      const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',blankrows:true,raw:false,range:previewRange}),best=bestHeaderFromMatrix(matrix,key,fullRange.s.r);
      candidates.push({...best,sheet:sheetName,dataRows:Math.max(0,fullRange.e.r-best.headerRow),configured:sheetName===profile.sheet,sheetOrder});
    });
    candidates.sort((a,b)=>Number(b.requiredOk)-Number(a.requiredOk)||b.score-a.score||Number(b.configured)-Number(a.configured)||b.dataRows-a.dataRows||a.sheetOrder-b.sheetOrder);
    return validateSourceCandidate(key,candidates[0]);
  }
  function canonicalize(rows,key){const profile=profileFor(key),headers=new Map(Object.keys(rows[0]||{}).map(name=>[norm(name),name])),columns=profile.fields.map(field=>[field.label,field.aliases.map(alias=>headers.get(alias)).find(Boolean)]);return rows.map(source=>{const row={...source};for(const [label,alias] of columns)if(!clean(row[label])&&alias)row[label]=source[alias];return row}).filter(row=>clean(row[profile.key])||(key==='nsr'&&clean(row[profile.number])));}
  function parseWorkbook(bytes,key){
    const workbook=XLSX.read(bytes,{type:'array'}),candidate=detectWorkbookSource(workbook,key),rows=canonicalize(XLSX.utils.sheet_to_json(workbook.Sheets[candidate.sheet],{defval:'',range:candidate.headerRow}),key);
    return {rows,sheet:candidate.sheet,headerRow:candidate.headerRow+1,score:candidate.score,fields:Object.keys(rows[0])};
  }
  let reader=null,readSequence=0;const pendingReads=new Map();
  function readWorkbook(bytes,key){
    profileFor(key);
    // The normal userscript path always uses the shared resource-backed worker.
    if(typeof createWorkbookWorker!=='function')return Promise.resolve().then(()=>parseWorkbook(bytes,key));
    if(!reader){
      const functions=[clean,norm,profileFor,scoreSourceHeader,betterHeader,bestHeaderFromMatrix,sourceScoreError,validateSourceCandidate,detectWorkbookSource,canonicalize,parseWorkbook];
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
  function decodeText(buffer){const bytes=new Uint8Array(buffer);if(bytes[0]===255&&bytes[1]===254)return new TextDecoder('utf-16le').decode(bytes);if(bytes[0]===254&&bytes[1]===255)return new TextDecoder('utf-16be').decode(bytes);try{return new TextDecoder('utf-8',{fatal:true}).decode(bytes)}catch(_){return new TextDecoder('gb18030').decode(bytes)}}
  function csvDelimiter(text){const probe=text.split(/\r?\n/).slice(0,50);return [',','\t',';','|'].map(delimiter=>{let score=0;for(const line of probe){let quoted=false;for(const char of line){if(char==='"')quoted=!quoted;else if(char===delimiter&&!quoted)score++}}return {delimiter,score}}).sort((a,b)=>b.score-a.score)[0].delimiter}
  async function readFile(file,key){
    const extension=file.name.split('.').pop().toLowerCase();
    if(extension==='json'&&key==='nsr'){const cache=JSON.parse(decodeText(await file.arrayBuffer()));if(!Array.isArray(cache.rows))throw new Error('No NSR cache rows.');return {rows:cache.rows,sheet:'NSR cache',from:cache.from,to:cache.to,savedAt:cache.savedAt}}
    if(extension!=='csv'){if(!['xlsx','xls','xlsm'].includes(extension))throw new Error('Use Excel, CSV or NSR JSON.');return readWorkbook(await file.arrayBuffer(),key)}
    const content=(typeof file.arrayBuffer==='function'?decodeText(await file.arrayBuffer()):await file.text()).replace(/^\uFEFF/,''),matrix=csvToMatrix(content,csvDelimiter(content)),candidate=validateSourceCandidate(key,{...bestHeaderFromMatrix(matrix,key),sheet:'CSV',dataRows:matrix.length}),rows=canonicalize(rowsFromMatrix(matrix,candidate.headerRow),key);
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
  let matchName=()=>'';
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
  function websiteStorage(){try{return typeof persistentStore==='function'?persistentStore():unsafeWindow.localStorage;}catch(_){try{return localStorage;}catch(__){return null;}}}
  function readLegacy(name){try{return websiteStorage()?.getItem(name)??null;}catch(_){return null;}}
  function persist(base,value){
    try{GM_setValue(key,JSON.stringify({base,mapping:value}));}catch(problem){error=problem.message||String(problem);return false;}
    // Only these two superseded mapping keys belong to this migration.
    try{const storage=websiteStorage();storage?.removeItem(key);storage?.removeItem(legacyKey);}catch(_){/* GM copy is already saved. */}return true;
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
  function reload(){
    notice='';conflicts=[];error='';install(defaults);
    try{
      const saved=GM_getValue(key,null),candidates=saved===null?[readLegacy(key),readLegacy(legacyKey)]:[saved];
      for(const candidate of candidates){
        if(candidate===null||candidate===undefined)continue;
        try{
          const data=typeof candidate==='string'?JSON.parse(candidate):candidate,base=data?.base||null,user=data?.mapping||data,merged=base?merge(base,user):{mapping:normalize(user),conflicts:[]};
          install(merged.mapping);conflicts=merged.conflicts;
          notice=base?(conflicts.length?'Updated defaults conflict with your saved mapping. Saved assignments are kept for now. Review the mappings below, then click Save to confirm, or Restore Defaults.\n\n'+conflicts.join('\n'):''):'This saved mapping has no update baseline. Review it and click Save to keep it, or Restore Defaults to use the latest defaults.';
          // Unresolved conflicts retain the old baseline until explicit Save or Reset.
          if(saved===null)persist(base,user);else if(base&&!notice&&JSON.stringify(base)!==JSON.stringify(defaults))persist(defaults,mapping);
          return result(true);
        }catch(_){/* Try the next valid legacy record without overwriting malformed data. */}
      }
    }catch(problem){error=problem.message||String(problem);}return result(!error);
  }
  function save(value){
    error='';let next;try{next=normalize(value);}catch(problem){error=problem.message||String(problem);return result(false);}
    if(!persist(defaults,next))return result(false);install(next);notice='';conflicts=[];return result(true);
  }
  function reset(){return save(defaults);}
  reload();
  return {key,defaults:{...defaults},getMap:()=>({...mapping}),getStatus,match:fab=>matchName(fab),normalize,save,reset,reload};
}

function createFidGraph(raw, options = {}) {
  const sourceKeys = ['fcid','ctl','ctlNsr','nsr','nce'];
  const input = Object.fromEntries(sourceKeys.map(key => [key, Array.isArray(raw?.[key]) ? raw[key] : []]));
  const cleanedValues=new Map();
  const clean = value => {
    if(cleanedValues.has(value))return cleanedValues.get(value);
    const text = String(value ?? '').trim();
    const result=/^(?:#|nan|not assigned|null|undefined|\[undefined\]|n\/a)$/i.test(text) ? '' : text;cleanedValues.set(value,result);return result;
  };
  const norm = value => String(value ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const textKey = value => clean(value).toLowerCase();
  const headerIndexes=new WeakMap(),headerSchemas=new Map(),aliasKeys=new Map();
  const field = (row, names) => {
    if (!row) return '';
    for (const name of names) if (Object.hasOwn(row,name)) {const value=clean(row[name]);if(value)return value;}
    let headers=headerIndexes.get(row);if(!headers){const keys=Object.keys(row),schema=keys.join('\0');headers=headerSchemas.get(schema);if(!headers){headers=new Map();for(const key of keys){const canonical=norm(key),variants=headers.get(canonical)||[];variants.push(key);headers.set(canonical,variants)}headerSchemas.set(schema,headers)}headerIndexes.set(row,headers)}
    for(const name of names){if(!aliasKeys.has(name))aliasKeys.set(name,norm(name));for(const key of headers.get(aliasKeys.get(name))||[]){const value=clean(row[key]);if(value)return value}}
    return '';
  };
  const validFid = value => { const id = clean(value); return id && !['NOFID','NOFCID','NA','NULL','UNKNOWN','NOTASSIGNED'].includes(norm(id).toUpperCase()) ? id : ''; };
  const loaded = Object.fromEntries(sourceKeys.map(key => [key, options.loadedSources ? (Array.isArray(options.loadedSources) ? options.loadedSources.includes(key) : !!options.loadedSources[key]) : input[key].length > 0]));
  const ownerPolicy='majority';
  const effectivePolicy='approved';
  const riskPolicy='approved';
  const dropQuarterLeadWeeks=19; // Inferred reporting window; use the Drop's creation snapshot, including overdue schedules.
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
  const initiator = record => person(field(record,['Initiator','Initiated By','Requester','Created By']));
  const discarded = status => /REJECT|CANCEL/.test(String(status).toUpperCase());
  const completed = status => /\bCOMPLETED\b/i.test(status);
  const nsrStatus=nsr=>[nsr.status,nsr.task,nsr.taskStatus].filter(Boolean).join(' ');
  const nsrNamesByFid=new Map(),authorsByFid=new Map(),knownUserIds=new Set(),aliasCandidates=new Map();
  for(const source of ['ctl','ctlNsr'])for(const row of input[source]){
    const account=field(row,['Initiator','Created By','Initiated By','Requester']),change=norm(field(row,['Change Cat','Change Type','Change Category']));
    if(!account)continue;if(/^[A-Z0-9_]+$/i.test(account))knownUserIds.add(norm(account));
    if(change&&change!=='slot'&&change!=='config')continue;
    const fid=validFid(field(row,['Opportunity ID','FCID','FID']));if(!fid)continue;
    const accounts=authorsByFid.get(fid)||new Set();accounts.add(norm(account));authorsByFid.set(fid,accounts);
  }
  for(const row of input.nsr){
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
    for (const row of rows) for (const [key,value] of Object.entries(row)) if (!clean(result[key]) && clean(value)) result[key] = value;
    return result;
  };
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
  const fids = new Map(), ctls = [], nsrs = [], nces = [];
  const orphans = {nsrs:[],ctls:[],nces:[],unmatchedNsrs:[]};
  const ensureFid = id => {
    if (!id) return null;
    if (!fids.has(id)) fids.set(id,{id,fidMark:['ignore','important'].includes(options.fidMarks?.[id])?options.fidMarks[id]:'normal',record:{},rawRows:[],sources:[],hasNativeRecord:false,inPopulation:false,excludedReason:'',account:'',businessUnit:'',customer:'',system:'',fidDescription:'',commit:'',quote:'',quarter:'',mcsd:'',rd:'',crd:'',targetShip:'',ctls:[],nsrs:[],nces:[],slot:null,owner:'',ownerEvidence:null,salesAdmin:'',salesAdminEvidence:null,pre:null,post:null,risks:[],category:'review'});
    return fids.get(id);
  };
  const sourceFidRows = new Map();
  for (const row of input.fcid) {
    const id = validFid(field(row,['FID','FCID','Field ID']));
    if (!id) continue;
    if (!sourceFidRows.has(id)) sourceFidRows.set(id,[]); sourceFidRows.get(id).push(row);
  }
  for (const [id,rows] of sourceFidRows) {
    const node = ensureFid(id), commits = rows.filter(row=>field(row,['Commit: Field','Commit Field','Field Commit','Commit Status']).toUpperCase()==='COMMIT'), candidates = commits.length ? commits : rows;
    const selected = candidates.find(row=>field(row,['Fab Segment','Fab_Segment'])) || candidates[0];
    node.record = {...selected}; node.rawRows = [...rows]; node.hasNativeRecord = true; node.sources.push('fcid');
    if (commits.length) {
      node.excludedReason = exclusion(selected); node.inPopulation = !node.excludedReason;
    }
  }
  const ctlGroups = new Map();
  for (const source of ['ctl','ctlNsr']) for (let index=0;index<input[source].length;index++) {
    const row=input[source][index], number=field(row,['CTL ID','CTL#']), key=number || '__'+source+'-row-'+index;
    if (!ctlGroups.has(key)) ctlGroups.set(key,[]); ctlGroups.get(key).push({source,row,index});
  }
    const ctlValueFields=Object.entries({
      fidId:['Opportunity ID','FCID','FID'],change:['Change Cat','Change Type','Change Category'],initiator:['Initiator','Created By','Initiated By','Requester'],businessUnit:['Business Unit','BU'],
      summary:['Change Summary','Summary'],created:['Creation Date','Create Date'],approved:['1st SOps Appr Dt','Sales Ops Approval Date','SOPS Approval Date'],mcsd:['MCSD @ Current','MCSD'],mcsdCreate:['MCSD@Create','MCSD At Create'],rdCreate:['RD At Create'],mcsdApproval:['MCSD At SOps Approval'],tsdApproval:['TSD @ SOPs Approval'],
      srd:['SRD @ Current','Sales Ops RD','SRD'],srdApproval:['SRD @ SOPs Approval'],fieldCrd:['Field CRD','CRD @ Current','CRD'],mfg:['Mfg Upgrade Type','Manufacturing Upgrade Type'],stage:['Stage','CTL Stage'],quote:['Quote ID','Quote Number'],sourceQuote:['Source Quote','Source Quote Number'],
      customer:['Customer group 2','Customer','Customer Name','End Customer'],configFrz:['Config Frz@Current','Config Frz','Config Frozen Date','Config Freeze Date'],rootCause:['Root Cause','Root Cause of Late Change'],reason:['Reason for Chg','Reason for Change'],salesOpsApprover:['Sales OPs Approver']
    });
  const ctlDateFields=['created','approved','mcsd','mcsdCreate','rdCreate','mcsdApproval','tsdApproval','srd','srdApproval','fieldCrd','configFrz'];
  // BOBJ truncates Sub Reason at 20 characters; match stable prefixes and retain raw values.
  const lateCtlSubReasonPrefixes=['Application Change','Build Type Change (N','Fab Change – Config','Install Base Issue R','New Feature CIP','Demand Reallocation','Common SPEC/Safety'];
  for (const [key,entries] of ctlGroups) {
    const rows = entries.map(entry=>entry.row), read = names => { for (const entry of entries) { const value=field(entry.row,names); if (value) return value; } return ''; };

    const reference=entries.find(entry=>entry.source==='ctlNsr')?.row;
    const ctl={id:key,number:read(['CTL ID','CTL#']),record:rows.length===1?rows[0]:mergeRows(rows),rows,rawRows:rows,sources:unique(entries.map(entry=>entry.source)),conflicts:[],links:[],changeSummary:field(reference,['Change Summary','Summary'])||read(['Change Summary','Summary']),notes:field(reference,['Notes','CTL Notes'])||read(['Notes','CTL Notes']),subReason:field(reference,['Sub Reason'])||read(['Sub Reason'])};
    for (const [name,names] of ctlValueFields) {
      ctl[name]=(name==='rootCause'&&field(reference,names))||read(names);
      if(entries.length>1){const alternatives=unique(entries.map(entry=>field(entry.row,names)).filter(Boolean)),comparison=alternatives.map(value=>ctlDateFields.includes(name)?iso(value)||textKey(value):textKey(value));if(unique(comparison).length>1)ctl.conflicts.push({field:name,values:alternatives})}
    }
    ctl.fidId=validFid(ctl.fidId); ctl.initiator=person(ctl.initiator); for(const name of ctlDateFields)ctl[name]=iso(ctl[name]);
    // Every CTL source uses Stage; Status columns remain only in the original source rows.
    ctl.stage=ctl.stage.replace(/^(?:\d+\s*-\s*)?(?:originator|draft)$/i,'Draft');
    const stage=String(ctl.stage||'').trim().replace(/^\d+\s*-\s*/,'').toLowerCase();
    ctl.status=stage&&stage!=='[unknown]'?({completed:'Completed',cancelled:'Cancelled',canceled:'Cancelled',rejected:'Rejected',draft:'Draft'})[stage]||'Submitted':'';
    ctl.changeKey=norm(ctl.change); ctl.statusKind=ctl.status==='Completed'?'completed':['Cancelled','Rejected'].includes(ctl.status)?'discarded':'pending';
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
    account:accountOf(record),status:field(record,['NSR Status','NsrWorkflowStatus','Workflow Status','Status']),initiator:initiator(record),
    title:field(record,['NSR Title','Title','NsrTitle']),type:field(record,['NSR Type','Type']),org:field(record,['NSR Org','Org','NsrOrg']),product:field(record,['Primary Product','Product','PrimaryProd']),customer:field(record,['Customer','Customer Name']),subsystem:field(record,['Impacted Sub System','Impacted Subsystem','Sub System','Subsystem']),submitted:iso(field(record,['Submit Date','Submitted Date','Submission Date','SubmitDate'])),tempId:field(record,['Temp Id#','Temp ID','TempId']),task:field(record,['Task Name','Task','Workflow Task']),taskStatus:field(record,['Task Status']),
    searchText:[number,...Object.values(record)].join(' ').toLowerCase()
  });
  const nsrGroups = new Map();
  for (let index=0;index<input.nsr.length;index++) {
    const row=input.nsr[index], number=field(row,['NSR#','NSR Number','Nsrnum']).toUpperCase(), key=number||'__nsr-row-'+index;
    if (!nsrGroups.has(key)) nsrGroups.set(key,[]); nsrGroups.get(key).push(row);
  }
  for (const [key,rows] of nsrGroups) {
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
    const record=input.nce[index], id=field(record,['NCe #','NCE Number'])||'__nce-row-'+index, fidId=validFid(field(record,['FCID','FID']));
    const nce={id,number:field(record,['NCe #','NCE Number']),fidId,record,rows:[record],status:field(record,['NCe State','State','Status']),impact:field(record,['Impact','NCE Impact']),title:field(record,['Title','NCE Title']),customer:field(record,['Customer','Customer Name']),opened:iso(field(record,['Open Date','Opened Date','Creation Date'])),damageGroup:field(record,['Damage Group']),damageCode:field(record,['Damage Code']),partNumber:field(record,['Part_Number','Part Number']),partDescription:field(record,['Part_Description','Part Description'])};
    nces.push(nce); const node=ensureFid(fidId); if (node) { node.nces.push(nce); node.sources=unique([...node.sources,'nce']); } else orphans.nces.push(nce);
  }
  const nsrByNumber=new Map(nsrs.filter(nsr=>nsr.number).map(nsr=>[nsr.number,nsr]));
  const tokenPattern=/(?<![A-Z0-9])([NDL]\d{5,8})(?![A-Z0-9])/gi;
  for (const ctl of ctls) {
    const tokens=new Map();
    for (const row of ctl.rows) for (const name of ['Change Summary','Summary','Root Cause','Root Cause of Late Change','Notes','CTL Notes']) {
      const value=field(row,[name]); if (!value) continue;
      for (const match of value.matchAll(tokenPattern)) { const number=match[1].toUpperCase(); if (!tokens.has(number)) tokens.set(number,new Set()); tokens.get(number).add(name); }
    }
    ctl.nsrReferences=[...tokens.keys()];
    for (const [number,fields] of tokens) {
      const nsr=nsrByNumber.get(number); if (!nsr) continue;
      const kind=ctl.fidId && nsr.fidIds.includes(ctl.fidId) ? 'direct':'reference';
      const link={ctlId:ctl.id,ctlNumber:ctl.number,nsrId:nsr.id,nsrNumber:nsr.number,fidId:ctl.fidId,nsrFidId:nsr.fidId,kind,fields:[...fields],effective:ctl.effective,applied:ctl.applied,status:ctl.status,statusKind:ctl.statusKind};
      ctl.links.push(link); nsr.links.push(link);
    }
  }
  const linkedLifecycle=links=>!links.length?'completed-unlinked':links.some(link=>link.applied)?'implemented':links.some(link=>clean(link.status))?'ctl-pending':'ctl-unavailable';
  function refreshNsrLifecycle(nsr) {
    nsr.implementationByFid=Object.fromEntries(nsr.fidIds.map(fidId=>[fidId,observationLifecycle(nsr,nsr,fidId)]));
    nsr.lifecycle=observationLifecycle(nsr,nsr);
  }
  nsrs.forEach(refreshNsrLifecycle);
  const tally = actions => {
    const counts=new Map(); for (const action of actions) if (action.name) { const key=textKey(action.name), entry=counts.get(key)||{name:action.name,count:0,documents:[]}; entry.count++; entry.documents.push({id:action.id,source:action.source}); counts.set(key,entry); }
    return [...counts.values()].sort((a,b)=>b.count-a.count||a.name.localeCompare(b.name));
  };
  // Slot authors establish Account membership; approval and current Slot state do not change who authored an operation.
  const accountPeople=new Map();
  for(const node of fids.values())node.account=accountOf(node.record);
  const nsrAccount=(observation,fidId=observation.fidId)=>fids.get(fidId)?.account||observation.account||'';
  for(const nsr of nsrs)nsr.accounts=[...new Set((nsr.observations.length?nsr.observations:[nsr]).map(observation=>nsrAccount(observation)))];
  for(const ctl of ctls){
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
    return {owner,evidence:{policy:ownerPolicy,inferred:true,account:node.account,accountPeople:[...people.values()],source:!people.size?'account-unavailable':!candidates.length?'no-eligible-votes':tied?tieBreakerSource||'unresolved-tie':ctlCandidates.length?'ctl':'nsr-fallback',confidence:!owner?'unknown':!ctlCandidates.length||tieBreakerSource==='nsr-tiebreaker'?'medium':'low',candidates,nsrCandidates,ctlCandidates,tied,tiedCandidates,tieBreaker,documents:candidates.reduce((sum,candidate)=>sum+candidate.count,0),excludedCandidates:tally(actions.filter(action=>!eligible(action))),aliases:[...aliasEvidence.values()].filter(evidence=>candidates.some(candidate=>candidate.name===evidence.name))}};
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
  function refreshFid(node) {
    node.risks=[];
    node.nsrs=node.nsrs.map(nsr=>{
      nsr=nsrByNumber.get(nsr.number)||nsr;
      if(!nsr.observations.length)return nsr;
      const observation=nsr.observations.find(observation=>observation.fidId===node.id);
      return observation?{...nsr,...observation,fidId:node.id,lifecycle:nsr.implementationByFid[node.id]||nsr.lifecycle,globalLifecycle:nsr.lifecycle}:nsr;
    });
    const record=node.record, nativeMcsd=firstDate(record,['Date: Mfg Commit Ship','Mfg Commit Ship','Manufacturing Commit Ship Date','MCSD']);
    node.customer=field(record,['Customer','Customer Name','End Customer'])||node.nsrs.find(nsr=>nsr.customer)?.customer||'';
    node.system=field(record,['System Description','System Desc','System']); node.quote=field(record,['Active Quote Number','Quote Number','Active Quote']);
    node.fab=field(record,['Fab Name','Fab']); node.fidDescription=field(record,['FID Description']); node.commit=field(record,['Commit: Field','Commit Field','Field Commit','Commit Status']); node.crd=firstDate(record,['Date: Customer Request','Customer Request Date']).date;
    node.businessUnit=field(record,['Business Unit','BU'])||node.ctls.find(ctl=>ctl.businessUnit)?.businessUnit||'';
    node.mcsd=nativeMcsd.date||node.ctls.find(ctl=>ctl.mcsd)?.mcsd||'';
    node.rd=firstDate(record,['Date: Sales Ops Request Date (RD)','Sales Ops Request Date','Sales Ops RD','SRD','RD']).date; node.targetShip=firstDate(record,['Date: Target Ship','Target Ship Date','Target Ship']).date;
    const events=node.ctls.filter(ctl=>(ctl.isSlot||ctl.isDrop)&&ctl.effective).sort((a,b)=>(b.eventDate||'').localeCompare(a.eventDate||''));
    const current=events[0]||null,active=current?.isSlot?current:null,riskActive=active||(current?.isDrop?events.find(ctl=>ctl.isSlot):null)||null,pending=node.ctls.filter(ctl=>(ctl.isSlot||ctl.isDrop)&&ctl.statusKind!=='discarded'&&!ctl.effective);
    const effectiveCtls=node.ctls.filter(ctl=>ctl.effective&&ctl.eventDate<=today);
    const quarterDate=nativeMcsd.date||node.rd,quarterField=nativeMcsd.date?'Current MCSD':'Current SRD';
    node.quarter=node.currentQuarter=quarterOf(quarterDate);
    node.quarterBasis=[quarterField+': '+(quarterDate||'[Unknown]')];
    const slotReady=loaded.ctl||loaded.ctlNsr;
    node.slot={active,current,pending,riskCurrent:current,riskActive,state:!slotReady?'unknown':current?.isDrop?'dropped':active?'active':'none',initiator:active?.initiator||'',number:active?.number||current?.number||'',quarter:node.quarter,effectivePolicy,riskPolicy,sourceReady:slotReady};
    const ownership=inferOwner(node); node.owner=ownership.owner; node.ownerEvidence=ownership.evidence;
    const salesAdminActions=node.ctls.filter(ctl=>ctl.initiator&&(ctl.changeKey==='pull'||ctl.changeKey==='push')).map(ctl=>({id:ctl.id,name:ctl.initiator,source:'ctl'})),salesAdminCandidates=tally(salesAdminActions);
    const salesAdminTied=salesAdminCandidates.length>1&&salesAdminCandidates[0].count===salesAdminCandidates[1].count;
    node.salesAdmin=salesAdminTied?'':salesAdminCandidates[0]?.name||'';node.salesAdminEvidence={candidates:salesAdminCandidates,tied:salesAdminTied,documents:salesAdminActions.length};
    node.pre=assessSlot(node,current,riskActive);
    const referenceDate=node.pre.referenceDate,referenceField=node.pre.referenceField;
    const postMcsd=current?.isDrop?nativeMcsd.date||riskActive?.mcsd:node.mcsd;
    const postReference=postMcsd||(current?.isDrop?referenceDate:active?.srd||node.rd||referenceDate),postReferenceField=postMcsd?'MCSD':current?.isDrop?referenceField:active?.srd?'Sales Ops RD':node.rd?'Date: Sales Ops Request Date (RD)':referenceField;
    const qualifying=slotReady?node.ctls.filter(ctl=>{
      if(!ctl.changeKey.includes('config')||!ctl.riskEligible||ctl.reason!=='Customer Config Chg'||!lateCtlSubReasonPrefixes.some(prefix=>ctl.subReason.startsWith(prefix)))return false;
      const span=weeks(ctl.eventDate,ctl.mcsdApproval||ctl.mcsdCreate||postReference);return span!==null&&span<20;
    }):[];
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
      const original=retry||ctl,snapshot=original.mcsdCreate||original.rdCreate;
      if(!snapshot)continue;
      const lead=weeks(original.created,snapshot);
      if(lead===null||lead>dropQuarterLeadWeeks)continue;
      addQuarter(quarterOf(snapshot),ctl,'Drop Slot · '+(retry?'original snapshot from CTL '+retry.number+' · ':'')+(original.mcsdCreate?'MCSD@Create':'RD At Create')+': '+snapshot+'; created '+original.created+'; lead '+lead.toFixed(1)+' weeks (<='+dropQuarterLeadWeeks+').',original.rdCreate,original.mcsdCreate);
    }
    const commitRows=node.rawRows.filter(row=>field(row,['Commit: Field','Commit Field','Field Commit','Commit Status']).toUpperCase()==='COMMIT');
    const sourceQuarters=new Set((commitRows.length?commitRows:node.rawRows).map(row=>quarterRank(field(row,['Time: Calendar Quarter']))).filter(Number.isFinite));
    const initialMcsd=firstDate(record,['Date: Initial MCSD at Slot','Initial MCSD at Slot']).date,initialRd=firstDate(record,['Initial SRD at Slot']).date;
    const initialQuarters=new Set([initialMcsd,initialRd].map(day=>quarterRank(quarterOf(day))).filter(Number.isFinite));
    for(const ctl of effectiveCtls)if(ctl.changeKey==='push')for(const [before,after,label]of [[ctl.rdCreate,ctl.srdApproval,'RD At Create → SRD @ SOPs Approval'],[ctl.mcsdCreate,ctl.mcsdApproval,'MCSD@Create → MCSD At SOps Approval']]){
      const q=quarterOf(before),rank=quarterRank(q);
      if(!before||!after||quarterRank(quarterOf(after))<=rank)continue;
      // A future RD abandoned before its quarter began is insufficient when a current MCSD exists.
      if(nativeMcsd.date&&label.startsWith('RD')&&quarterOf(ctl.mcsdCreate)!==q&&rank>quarterRank(quarterOf(ctl.eventDate)))continue;
      if(rank-quarterRank(quarterOf(ctl.eventDate))<=1||sourceQuarters.has(rank))addQuarter(q,ctl,label+': '+before+' → '+after);
    }
    for(const [createdField,approvalField,label]of [['rdCreate','srdApproval','RD'],['mcsdCreate','mcsdApproval','MCSD']]){
      if(nativeMcsd.date&&label==='RD')continue;
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
    const splitFlag=/^(?:Y|YES|X|TRUE|1)$/i.test(field(record,['Split Ship Flag']));
    for(const ctl of effectiveCtls){
      const partial=splitFlag||/\b(?:split|partial)\s+(?:ship(?:ment)?|deliver)|\banother\s+PM(?:x|\b)|\bship\b[\s\S]{0,120}\bfirst\b[\s\S]{0,120}\b(?:later|second)\b/i.test([ctl.rootCause,ctl.summary,ctl.notes].join(' '));
      if(partial&&ctl.tsdApproval&&ctl.srdApproval&&quarterOf(ctl.tsdApproval)!==quarterOf(ctl.srdApproval))addQuarter(quarterOf(ctl.tsdApproval),ctl,'Split delivery · TSD @ SOPs Approval: '+ctl.tsdApproval+'; SRD @ SOPs Approval: '+ctl.srdApproval);
    }
    // Retain additional schedules only once approval reaches their preceding fiscal quarter.
    const initialDate=initialMcsd||initialRd;
    if(riskActive&&initialDate&&quarterRank(quarterOf(initialDate))-quarterRank(quarterOf(riskActive.eventDate))<=1){
      addQuarter(quarterOf(initialDate),riskActive,'Initial Slot '+(initialMcsd?'MCSD':'SRD')+': '+initialDate+'; Slot approved from preceding fiscal quarter onward.',initialRd,initialMcsd);
    }
    for(const ctl of effectiveCtls)if(ctl.changeKey==='pull')for(const [before,after,label]of [[ctl.rdCreate,ctl.srdApproval,'RD At Create → SRD @ SOPs Approval'],[ctl.mcsdCreate,ctl.mcsdApproval,'MCSD@Create → MCSD At SOps Approval']]){
      const q=quarterOf(before),rank=quarterRank(q);
      if(!before||!after||quarterRank(quarterOf(after))>=rank)continue;
      if(nativeMcsd.date&&label.startsWith('RD')&&quarterOf(ctl.mcsdCreate)!==q)continue;
      if(rank-quarterRank(quarterOf(ctl.eventDate))<=1)addQuarter(q,ctl,'Approved cross-quarter Pull · '+label+': '+before+' → '+after);
    }
    for(const ctl of effectiveCtls)if(ctl.changeKey==='push'&&ctl.mcsdCreate&&quarterRank(quarterOf(ctl.mcsdCreate))-quarterRank(quarterOf(ctl.eventDate))<=1){
      addQuarter(quarterOf(ctl.mcsdCreate),ctl,'Push MCSD@Create: '+ctl.mcsdCreate+'; approved from preceding fiscal quarter onward.');
    }
    // Classification uses only quarters already supported by schedule evidence; a Late CTL never adds a quarter.
    for(const ctl of qualifying){
      const snapshot=ctl.mcsdCreate||ctl.rdCreate,dateBasis=snapshot||postReference,q=snapshot?quarterOf(snapshot)||node.quarter:node.quarter;
      if(!groups.has(q))continue;
      const group=q?addQuarter(q,ctl,(ctl.mcsdCreate?'MCSD@Create':ctl.rdCreate?'RD At Create':postReferenceField)+': '+dateBasis):groups.get(q);
      group.ctls.push(ctl);
    }
    node.quarterStates=[...groups.values()].map(group=>{
      const cycles=[...group.cycles.values()],slotted=cycles.filter(cycle=>cycle.current),preCandidates=(slotted.length?slotted:cycles).map(cycle=>cycle.current===current&&group.quarter===node.currentQuarter?node.pre:assessSlot(node,cycle.current,cycle.riskActive,false));
      const pre=preCandidates.find(pre=>pre.state==='risk')||preCandidates.find(pre=>pre.state==='unknown')||preCandidates[0];
      const late=group.ctls,first=late[0],reference=first?.mcsdApproval||first?.mcsdCreate||postReference||'',field=first?.mcsdApproval?'MCSD At SOps Approval':first?.mcsdCreate?'MCSD@Create':postReferenceField,span=first?weeks(first.eventDate,reference):null;
      const post={state:late.length?'risk':slotReady&&postReference?'clear':'unknown',label:late.length?'Post config risk':slotReady&&postReference?'No post risk':'Post risk unknown',weeks:span,threshold:20,
        reason:late.length?`(${field} − CTL approved date) / 7 = ${span.toFixed(1)} weeks <20 weeks.`:slotReady&&postReference?'No qualifying Late CTL in this quarter.':!slotReady?'CTL not loaded.':'MCSD / SRD missing.',
        eventDate:first?.eventDate||'',referenceDate:reference,referenceField:field,created:first?.created||'',approved:first?.approved||'',ctlIds:late.map(ctl=>ctl.id),ctls:late,nsrs:pendingNsrs,pending:pendingConfigs,policy:riskPolicy};
      const dates=group.quarter===node.currentQuarter?{rd:node.rd,mcsd:node.mcsd}:{rd:unique(group.rd).sort().join(' / '),mcsd:unique(group.mcsd).sort().join(' / ')};
      const dateSource=group.sources.size?'CTL schedule evidence: '+[...group.sources].join(', ')+'\n'+group.basis.join('\n'):'Current FID / CTL dates';
      return assessPortfolio({quarter:group.quarter,pre,post,basis:group.basis,slotEvidence:preCandidates,dates,dateSource},node,slotted.length>0);
    }).sort((a,b)=>a.quarter.slice(-2).localeCompare(b.quarter.slice(-2))||['Mar','Jun','Sep','Dec'].indexOf(a.quarter.slice(1,4))-['Mar','Jun','Sep','Dec'].indexOf(b.quarter.slice(1,4)));
    Object.assign(node,node.quarterStates.find(state=>state.quarter===node.currentQuarter));
    if(node.quarterStates.length===1)node.selectedQuarterStates=node.quarterStates;
    node.weeksToMcsd=node.mcsd?Math.round(weeks(today,node.mcsd)):'';
    node.entitySearchText=[node.id,node.account,node.businessUnit,node.customer,node.system,node.owner,node.salesAdmin,node.quote,...Object.values(record),...node.ctls.flatMap(ctl=>[ctl.number,ctl.summary,ctl.initiator]),...node.nces.flatMap(nce=>[nce.number,nce.title])].join(' ').toLowerCase();
  }
  for(const node of fids.values())refreshFid(node);
  for(const ctl of ctls){const node=fids.get(ctl.fidId);ctl.quarter=node?.quarter||'';ctl.quarters=node?.quarterStates.map(state=>state.quarter)||[''];if(!ctl.customer)ctl.customer=field(node?.record||{},['Customer','Customer Name','End Customer'])}
  function setNsrMark(number,mark) {
    const nsr=nsrByNumber.get(number);if(!nsr)return false;
    nsr.ignored=mark==='ignore';nsr.important=mark==='important';
    for(const fidId of nsr.fidIds){const row=fids.get(fidId).nsrs.find(row=>row.number===number);if(row)Object.assign(row,{ignored:nsr.ignored,important:nsr.important})}
    return true;
  }
  function setFidMark(id,mark){const node=fids.get(id);if(!node)return false;node.fidMark=['ignore','important'].includes(mark)?mark:'normal';return true}
  const observedPopulation=[...fids.values()],population=observedPopulation.filter(node=>node.inPopulation);
  const values = value => value===undefined||value===null||value===''?[]:(Array.isArray(value)?value:[value]).map(clean);
  const match = (actual, expected) => { if(expected===undefined||expected===null||expected===''||Array.isArray(expected)&&!expected.length)return true;const selected=values(expected); return !selected.length||selected.some(value=>value==='[Unknown]'?!clean(actual):textKey(actual)===textKey(value)); };
  const hasPredicates = object => !!object&&Object.values(object).some(value=>values(value).length);
  const searchMatch = (text, search) => {
    const term=clean(search).toLowerCase(); if (!term) return true;
    const terms=term.split(/[\s,;|/]+/).filter(Boolean),idOnly=terms.length>1&&terms.every(part=>/^(?:\d{5,10}|[ndl]\d{5,8})$/i.test(part));
    return idOnly?terms.some(part=>text.includes(part)):text.includes(term);
  };
  function observationLifecycle(nsr,observation,fidId=''){
    const status=nsrStatus(observation);
    if(discarded(status))return 'discarded';if(!completed(status))return 'open';if(nsr.number.startsWith('D'))return 'completed-5d';if(!nsr.linkAssessmentAvailable)return 'ctl-unavailable';
    const links=nsr.links.filter(link=>link.kind==='direct'&&(!fidId||link.fidId===fidId)&&link.statusKind!=='discarded');return linkedLifecycle(links);
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
  function query(filters={},rowsOnly='') {
    const includeIgnored=filters.includeIgnored!==false;let base=filters.excluded?observedPopulation.filter(node=>node.excludedReason):population;
    if(!includeIgnored)base=base.filter(node=>node.fidMark!=='ignore');
    const ctlRows=hasPredicates(filters.ctl)?ctls.filter(ctl=>matchCtl(ctl,filters.ctl)):ctls,ctlFids=ctlRows===ctls?null:new Set(ctlRows.map(ctl=>ctl.fidId));
    // An absent CTL matches Unknown only when every selected CTL field accepts an empty value.
    const ctlUnknown=!!ctlFids&&matchCtl({},filters.ctl),matchCtlFid=fidId=>!ctlFids||ctlFids.has(fidId)||ctlUnknown&&!fids.get(fidId)?.ctls.length;
    const directKeys=['quarter','systemQuarter','account','customer','system','businessUnit','salesAdmin','owner','category','preState','postState','slotState','fidStatus','fidMark','signal','riskType'];
    const hasFidPredicates=!!filters.excluded||directKeys.some(key=>key!=='account'&&values(filters[key]).length)||hasPredicates(filters.slot);
    const scalarKeys=directKeys.filter(key=>!['quarter','systemQuarter','fidStatus','signal','riskType'].includes(key)&&values(filters[key]).length),fidStatuses=values(filters.fidStatus),signals=values(filters.signal),riskTypes=values(filters.riskType).map(label=>['Post NSR risk','Late NSR','Open NSR'].includes(label)?'Pending NSR':label==='Pending CTL'?'Pending Config CTL':label),nsrRiskOnly=riskTypes.length>0&&riskTypes.every(isNsrRisk),hasNsrPredicates=hasPredicates(filters.nsr);
    const search=clean(filters.search).toLowerCase(),tokens=search.split(/[\s,;|/]+/).filter(Boolean),idSearch=!!tokens.length&&tokens.every(token=>/^(?:\d{5,10}|[ndl]\d{5,8})$/i.test(token));
    const entityMatch=node=>idSearch?tokens.some(token=>/^\d+$/.test(token)&&(node.id.toLowerCase()===token||node.ctls.some(ctl=>ctl.number.toLowerCase()===token))):node.entitySearchText.includes(search);
    const nsrSearchMatch=(nsr,observation=nsr)=>!search||(idSearch?tokens.some(token=>nsr.number.toLowerCase()===token||nsr.number.replace(/^[NDL]/,'')===token):observation.searchText.includes(search));
    const hasAccount=values(filters.account).length>0;
    const observationMatches=(nsr,observation,fidId='')=>matchObservation(nsr,observation,filters.nsr||{},fidId)&&(!hasAccount||match(nsrAccount(observation,fidId||observation.fidId),filters.account));
    const nsrDomainMatch=(nsr,fidId='')=>nsr.eligible&&(includeIgnored||!nsr.ignored)&&(nsr.observations.length?nsr.observations.some(observation=>observationMatches(nsr,observation,fidId)&&nsrSearchMatch(nsr,observation)):observationMatches(nsr,nsr,fidId)&&nsrSearchMatch(nsr));
    const entityHits=search?new Set(observedPopulation.filter(entityMatch).map(node=>node.id)):null,entityHit=fidId=>!entityHits||entityHits.has(fidId);
    const scopedFids=base.flatMap(node=>{
      const scoped=projectFid(node,filters.quarter,filters.systemQuarter);if(!scoped)return [];
      return rowsOnly==='periods'?scoped.selectedQuarterStates.map(state=>projectFid(node,[state.quarter||'[Unknown]'])):[scoped];
    });
    const selectedFids=scopedFids.filter(node=>{
      if(fidStatuses.length&&!fidStatuses.some(status=>status.startsWith('cause:')?matchSignal(node,status.slice(6),includeIgnored):status===node.portfolioCell||node.portfolioStates.includes(status)||node.portfolioCell.endsWith('|'+status)))return false;
      for(const key of scalarKeys){const value=key==='preState'?node.pre.state:key==='postState'?node.post.state:key==='slotState'?node.slot.state:node[key];if(!match(value,filters[key]))return false}
      if(riskTypes.length&&!node.risks.some(risk=>riskTypes.includes(risk.label)))return false;
      if(signals.length&&!signals.some(signal=>matchSignal(node,signal,includeIgnored)))return false;
      if(!entityHit(node.id)&&!node.nsrs.some(nsr=>nsrDomainMatch(nsr,node.id)))return false;
      if (hasNsrPredicates&&!node.nsrs.some(nsr=>(includeIgnored||!nsr.ignored)&&matchNsr(nsr,filters.nsr,node.id))) return false;
      if (hasPredicates(filters.slot)) { const slot=node.slot.active; if (!slot||!match(slot.initiator,filters.slot.initiator)||!node.selectedQuarterStates.some(state=>match(state.quarter,filters.slot.quarter))||!match(slot.status,filters.slot.status)) return false; }
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
    // Only explicit FID conditions restrict NSRs to matching Commit FIDs.
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
        return entityHit(fidId)?(!hasNsrPredicates&&!hasAccount)||(nsr.observations.length?nsr.observations:[nsr]).some(observation=>observationMatches(nsr,observation,fidId)):nsrDomainMatch(nsr,fidId);
      });
      const orphan=!nsr.fidId&&!hasFidPredicates&&(!ctlFids||ctlUnknown)&&nsrDomainMatch(nsr);
      if(!matchedFidIds.length&&!orphan)continue;
      if(!nsr.observations.length){selectedNsrs.push({...nsr,taskLifecycle:observationLifecycle(nsr,nsr,nsr.fidId)});continue;}
      const scopeLifecycles=unique(matchedFidIds.map(fidId=>nsr.implementationByFid[fidId]||nsr.lifecycle)),lifecycle=['open','discarded','completed-5d','completed-unlinked','ctl-pending','implemented','ctl-unavailable'].find(state=>scopeLifecycles.includes(state))||nsr.lifecycle;
      const fidId=matchedFidIds[0]||nsr.fidId,observation=nsr.observations.find(observation=>observationMatches(nsr,observation,fidId)&&(entityHit(fidId)||nsrSearchMatch(nsr,observation)))||nsr.observations[0];
      let taskLifecycle=observationLifecycle(nsr,observation,fidId);
      if(scopeLifecycles.includes('completed-unlinked')&&!['open','discarded'].includes(taskLifecycle))taskLifecycle='completed-unlinked';
      selectedNsrs.push({...nsr,...observation,fidId,lifecycle:hasNsrPredicates?taskLifecycle:lifecycle,taskLifecycle,globalLifecycle:nsr.lifecycle,matchedFidIds,scopeLifecycles});
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
    const pendingCtlFids=new Set(),pendingPullPushCtlFids=new Set();
    for(const ctl of selectedCtls){metrics.pendingCtl+=ctl.pending;metrics.pendingPullPushCtl+=ctl.pendingPullPush;metrics.workflowPendingCtl+=ctl.status==='Submitted';if(ctl.pending)pendingCtlFids.add(ctl.fidId);if(ctl.pendingPullPush)pendingPullPushCtlFids.add(ctl.fidId)}
    metrics.pendingCtlFids=pendingCtlFids.size;metrics.pendingPullPushCtlFids=pendingPullPushCtlFids.size;
    return {fids:selectedFids,nsrs:selectedNsrs,ctls:selectedCtls,nces:selectedNces,metrics};
  }
  return {fids,population,observedPopulation,nsrs,ctls,nces,orphans,query,projectFid,setNsrMark,setFidMark,isNsrRisk};
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
  const clean=value=>{const s=String(value??'').replace(/^\uFEFF/,'').trim();return /^(nan|null|undefined)$/i.test(s)?'':s}; const pageFetch=(url,options)=>unsafeWindow.fetch(url,options); const workflowDate=value=>clean(value).match(/^(\d{4}-\d{2}-\d{2})/)?.[1]||'';
  const xmlFirst=(node,names)=>names.map(name=>clean(node.getElementsByTagNameNS('*',name)[0]?.textContent)).find(Boolean)||'';
const SOURCE_FIELDS=['Temp Id#','NSR#','NSR Status','REF FCID','NSR Request','Task Name','NSR Title','NSR Type','Impacted Sub System','Initiator','Sales Rep','Sales Ops','Technical Contact','Customer','Submit Date','NSR Org','Primary Product'];
const SAP_SEARCH='/sap/opu/odata/sap/ZNSR_CD_SRV/SearchNsrSet';
const SAP_FIELDS={
  'Temp Id#':['TempId'],'NSR#':['Nsrnum'],'NSR Status':['NsrWorkflowStatus','NsrStatus'],'REF FCID':['RefFcid'], 'NSR Request':['NsrRequest'],'Task Name':['DescriptionTask'],'NSR Title':['NsrTitle','EnggTitle'],'NSR Type':['NsrType'], 'Impacted Sub System':['ImpSubSys'],
    'Initiator':['InitiatedByFn','Requester','InitiatedBy'],'Sales Rep':['SalesRepFn'], 'Sales Ops':['SalesOpsFn','SalesOps'],'Technical Contact':['TechContactFn'],'Customer':['Customer'],'Submit Date':['SubmitDate'], 'NSR Org':['NsrOrg'],'Primary Product':['PrimaryProd']
};
const NSR_CACHE_DB='toolkitKyra',NSR_CACHE_DB_VERSION=1,NSR_CACHE_STORE='datasets';let cacheDb;

function openCache(){
  if(!cacheDb)cacheDb=new Promise((resolve,reject)=>{
    const api=unsafeWindow.indexedDB;if(!api){reject(new Error('IndexedDB is unavailable.'));return}
    const request=api.open(NSR_CACHE_DB,NSR_CACHE_DB_VERSION);let settled=false; request.onupgradeneeded=()=>{const db=request.result;if(!db.objectStoreNames.contains(NSR_CACHE_STORE))db.createObjectStore(NSR_CACHE_STORE,{keyPath:'key'})};
    request.onsuccess=()=>{const db=request.result;if(settled){db.close();return}settled=true;db.onversionchange=db.onclose=()=>{db.close();cacheDb=null};resolve(db)}; request.onerror=()=>{if(!settled){settled=true;reject(request.error||new Error('Could not open the data cache.'))}};
    request.onblocked=()=>{if(!settled){settled=true;reject(new Error('Data cache is blocked by another page.'))}};
  }).catch(error=>{cacheDb=null;throw error});return cacheDb;
}

function cacheRequest(request){return new Promise((resolve,reject)=>{request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error||new Error('Cache request failed.'))});}
function cacheDone(transaction){return new Promise((resolve,reject)=>{transaction.oncomplete=resolve;transaction.onerror=()=>reject(transaction.error||new Error('Cache write failed.'));transaction.onabort=()=>reject(transaction.error||new Error('Cache write cancelled.'))});}

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

async function readDataset(key){return cacheRequest((await openCache()).transaction(NSR_CACHE_STORE,'readonly').objectStore(NSR_CACHE_STORE).get(key))}
async function writeDataset(entry){const transaction=(await openCache()).transaction(NSR_CACHE_STORE,'readwrite');transaction.objectStore(NSR_CACHE_STORE).put(entry);await cacheDone(transaction)}
async function deleteDatasets(keys){const transaction=(await openCache()).transaction(NSR_CACHE_STORE,'readwrite');keys.forEach(key=>transaction.objectStore(NSR_CACHE_STORE).delete(key));await cacheDone(transaction)}
return {fields:SOURCE_FIELDS,localDate,sapRange,validSapDate,fetchSapRows,readDataset,writeDataset,deleteDatasets};
}

function createNsrRecords(){
const {fields:SOURCE_FIELDS}=createNsrDataService();
const DETAIL_FIELDS=['NSR#','NSR Title','Task Status','NSR Type','Impacted Sub System','Initiator','REF FCID','NSR Org','Primary Product','Customer','Sales Rep','Sales Ops','Technical Contact','Submit Date','NSR Request','Temp Id#'];
const ALIASES={
  'Temp Id#':['Temp Id#','Temp ID','TempId','Temporary ID'], 'NSR#':['NSR#','NSR','NSR No','NSR Number','NSR ID'], 'NSR Status':['NSR Status','Status'], 'REF FCID':['REF FCID','Reference Fcid','Reference FCID','Ref FCID','FCID'], 'NSR Request':['NSR Request','Request','Request Type'],
  'Task Name':['Task Name','Task','Workflow Task'], 'NSR Title':['NSR Title','Title'], 'NSR Type':['NSR Type','Type'], 'Impacted Sub System':['Impacted Sub System','Impacted Subsystem','Sub System','Subsystem'], 'Initiator':['Initiator','Initiated By','Requester'],
  'Sales Rep':['Sales Rep','Sales Representative'], 'Sales Ops':['Sales Ops','Sales Operation','Sales Operations'], 'Technical Contact':['Technical Contact','Tech Contact'], 'Customer':['Customer','Customer Name','Account'],
  'Submit Date':['Submit Date','Submitted Date','Submission Date'], 'NSR Org':['NSR Org','NSR Organization','Organization','Org'], 'Primary Product':['Primary Product','Product','PrimaryProduct']
};
const TASK_STATUS={
  'CREATE NSR':'1-Create NSR','REQUEST REVIEW':'2-Request Review','SAFETY REVIEW':'5-Safety Review','ESTIMATES':'6-Estimates','COST REVIEW':'7-Cost Review',
  'PRICE CONFIRMATION':'8-Price Confirmation','DEMAND CONFIRMATION':'9-Demand Confirmation','FINAL APPROVAL':'10-Final Approval','VC VALIDATION':'11-VC Validation','ENGINEERING COMMIT AND DESIGN':'12-Engineering Commit and Design'
};
const STATUS_TASK={'PENDING WITH INITIATOR':'1-Create NSR','PENDING WITH SAFETY':'5-Safety Review','PENDING WITH PGBD':'8-Price Confirmation','PENDING WITH POA PM':'11-VC Validation','PENDING WITH DESIGN ENGINEERING':'12-Engineering Commit and Design'};
const collator=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'});
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
function compareStatus(a,b){const aa=statusRank(a),bb=statusRank(b); return aa[0]-bb[0]||collator.compare(aa[1],bb[1]);}
function buildHeaderMap(headers){const byNorm=new Map(headers.map(h=>[normHeader(h),h])); const map={}; for(const field of SOURCE_FIELDS){const aliases=(ALIASES[field]||[field]).map(normHeader); map[field]=aliases.map(a=>byNorm.get(a)).find(Boolean)||'';} return map;}

let cachedHeaders=null,cachedHeaderMap;
function canonicalRow(raw,headerMap){if(!headerMap){const headers=Object.keys(raw),signature=headers.join('\0');if(signature!==cachedHeaders){cachedHeaders=signature;cachedHeaderMap=buildHeaderMap(headers)}headerMap=cachedHeaderMap}return Object.fromEntries(SOURCE_FIELDS.map(field=>[field,clean(raw[headerMap[field]])]))}
function normalizeRows(rawRows){
  if(!rawRows.length)throw new Error('No data rows were found.');
  const headers=Object.keys(rawRows[0]||{}),headerMap=buildHeaderMap(headers);
  if(!headerMap['NSR#'])throw new Error('Required header “NSR#” was not found.');
  const map=new Map();let missingNsr=0,duplicates=0;
  rawRows.forEach(raw=>{
    const row=canonicalRow(raw,headerMap),nsr=row['NSR#'];
    if(!nsr){missingNsr++;return}
    if(map.has(nsr)){duplicates++; const existing=map.get(nsr); SOURCE_FIELDS.forEach(field=>{if(!existing[field]&&row[field])existing[field]=row[field]});}else map.set(nsr,row);
  });
  const rows=[...map.values()].map(row=>({...row,'Task Status':deriveTaskStatus(row)})); const unmapped=rows.filter(r=>r['Task Status'].startsWith('Unmapped')).length;
  return {rows,stats:{inputRows:rawRows.length,validRows:rows.length,missingNsr,duplicates,unmapped}};
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
return {detailFields:DETAIL_FIELDS,clean,upper,canonicalRow,normalizeRows,deriveTaskStatus,compareStatus,dateNumber,periodRank,periodYear};
}

function createNsrCycle(root=document,{getRows=null,onNotify=null,renderTable,copyRows}={}){
const ownerDocument=root.ownerDocument||root,embedded=root!==ownerDocument,modalHost=embedded?root:ownerDocument.body;
const {detailFields:DETAIL_FIELDS,clean,upper,dateNumber,compareStatus}=createNsrRecords();
const LIVE_FIELDS=['Submit Date','Task Status','ECD','Approver Team','Assigned Approver','Pending Days','Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'];
const CYCLE_DETAIL_FIELDS=DETAIL_FIELDS.filter(field=>field!=='NSR#'&&field!=='Initiator'&&!LIVE_FIELDS.includes(field)).flatMap(field=>field==='NSR Type'?[field,'NSR Category']:field==='Customer'?[field,'Initiator']:[field]);
const FID_FIELDS=['REF FCID','Primary Product','NSR → CTL','CTL references','Customer','NSR Org','Sales Rep'],fidIndex=CYCLE_DETAIL_FIELDS.indexOf('REF FCID');
const CYCLE_FIELDS=['NSR#',...LIVE_FIELDS,...CYCLE_DETAIL_FIELDS.slice(0,fidIndex),...FID_FIELDS,...CYCLE_DETAIL_FIELDS.slice(fidIndex).filter(field=>!FID_FIELDS.includes(field))];
const WF_FIELDS=['Task Name','Approver Status','Approver Team','Approver','Pending Days','Approved Date','Comments']; const WORKFLOW_EXPORT_FIELDS=['NSR#','Workflow Step',...WF_FIELDS];
const CYCLE_WIDTHS=CYCLE_FIELDS.map(field=>({
  'NSR#':78,'Submit Date':116,'Approver Team':125, 'Assigned Approver':130,'Pending Days':82,'Elapsed Days':82,'Cycle Time (ECD)':110,'Cycle Time (SOBi)':116,'NSR Title':190,'Task Status':144,'ECD':96, 'NSR Category':100,'Customer':150,'Primary Product':120,'Technical Contact':108,
    'Impacted Sub System':104,'NSR → CTL':180,'CTL references':180
}[field]||92));
const CYCLE_CONCURRENCY=8,CYCLE_WARN_AT=100,CYCLE_LIMIT=500; const WORKFLOW_ODATA='/sap/opu/odata/sap/ZNSR_CD_SRV/WorkflowSet/'; const NSR_DETAIL_ODATA='/sap/opu/odata/sap/ZNSR_CD_SRV/GetNsrDetSet/'; const FIORI_DETAIL='https://fep.lamresearch.com/flp#ZNSR-display&/NsrDetail/';
const pageFetch=(url,options)=>unsafeWindow.fetch(url,options);
const collator=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'});
const state={rows:[],cycleRows:[],cycleAbort:null,cycleRun:0,toastTimer:null};
const $id=id=>root.querySelector('#'+id),esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function statusPill(value){const cls=value==='15-Completed'?'terminal':value.startsWith('0-')?'stop':value.startsWith('Unmapped')?'unknown':'';return `<span class="status-pill ${cls}">${esc(value)}</span>`;}
function toast(message,isError=false){
  if(onNotify){onNotify(message,isError);return}
  const el=$id('toast');if(!el)return;clearTimeout(state.toastTimer);el.textContent=message;el.className=`toast show${isError?' error':''}`;state.toastTimer=setTimeout(()=>el.className='toast',3000);
}
let wfTrigger=null,workflowTable=null,cycleTable=null;
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
  CYCLE_FIELDS.forEach(field=>{if(field!=='NSR#'&&field!=='NSR Category'&&!LIVE_FIELDS.includes(field))row[field]=detail[field]||''}); row['Task Status']=detail['Task Status']||({COMPLETED:'15-Completed',CANCELLED:'0-Cancelled',REJECTED:'0-Rejected'}[row.__terminal]||''); row['ECD']=row.__ecd||'';
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
  if(total>CYCLE_WARN_AT&&!unsafeWindow.confirm(`You are about to query ${total} NSRs from SAP. This may take some time. Continue?`)){setCycleProgress(0,0,`Query cancelled. ${total} NSRs were not sent to SAP.`);return;}
  if(state.cycleAbort)state.cycleAbort.abort();
  const run=++state.cycleRun,controller=new unsafeWindow.AbortController();state.cycleAbort=controller;state.cycleRows=[]; const details=new Map((getRows?getRows(parsed.valid):state.rows).map(row=>[upper(row['NSR#']),row])),results=new Array(total),workers=Math.min(CYCLE_CONCURRENCY,total);
  let next=0,done=0,stoppedForAuth=false,lastPaint=0; setCycleBusy(true);setCycleProgress(0,total,`Starting ${total} live lookup${total===1?'':'s'} · up to ${workers} requests at once…`);await renderCycleTable();
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
      results[index]=mergeCycle(nsr,live,details);done++;state.cycleRows=results.filter(Boolean); const now=Date.now();if(now-lastPaint>200||done===total||stoppedForAuth){lastPaint=now;await renderCycleTable()}
      if(run!==state.cycleRun)return;
      if(!stoppedForAuth)setCycleProgress(done,total,`Queried ${done} of ${total} · up to ${workers} requests at once`);
    }
  }
  await Promise.all(Array.from({length:workers},()=>worker()));
  if(run!==state.cycleRun)return;
  state.cycleRows=results.filter(Boolean);await renderCycleTable();if(run!==state.cycleRun)return;
  const cancelled=controller.signal.aborted&&!stoppedForAuth,failed=state.cycleRows.filter(row=>row.__error||row.__categoryError).length;state.cycleAbort=null;setCycleBusy(false);
  if(stoppedForAuth){setCycleProgress(done,total,'Stopped because the Fiori session is unavailable. Return to Fiori, sign in, and run the query again.');toast('Fiori session unavailable.',true)}
  else if(cancelled)setCycleProgress(done,total,`Cancelled after ${done} of ${total} lookups.`);
  else{setCycleProgress(done,total,`${done} live lookup${done===1?'':'s'} completed${failed?` · ${failed} unavailable`:''}.`);
    toast(`${done} cycle time row${done===1?'':'s'} loaded${failed?` · ${failed} unavailable`:''}.`,failed>0)}
}

function compareCycle(a,b,field){
  if(['Pending Days','Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'].includes(field))return (Number(a)||0)-(Number(b)||0);
  if(field==='ECD'||field.includes('Date'))return dateNumber(a)-dateNumber(b);
  if(field==='Task Status')return compareStatus(a,b);
  return collator.compare(clean(a),clean(b));
}

function cycleRows(){
  return cycleTable?.ready?cycleTable.grid.getRows('active').map(row=>cycleTable.rows[row.getIndex()]):[];
}
function cycleFields(){return cycleTable.grid.getColumns().filter(column=>column.isVisible()).map(column=>CYCLE_FIELDS[Number(column.getField().slice(1))])}

function nsrDetailUrl(row){return clean(row['Temp Id#'])&&clean(row['NSR#'])?`${FIORI_DETAIL}${encodeURIComponent(clean(row['Temp Id#']))}/${encodeURIComponent(clean(row['NSR#']))}/Search`:''}

function cycleCell(row,field){
  if(field==='Task Status')return statusPill(row[field]);
  if(field==='NSR Title'&&clean(row[field])&&clean(row['Temp Id#'])&&clean(row['NSR#'])){
    const href=nsrDetailUrl(row);
    return `<a class="ct-link" data-nsr-detail="${esc(row['NSR#'])}" href="${esc(href)}" target="_blank" rel="noopener noreferrer" title="Open ${esc(row['NSR#'])} in Fiori">${esc(row[field])}</a>`;
  }
  if(['NSR#','Elapsed Days','Cycle Time (ECD)','Cycle Time (SOBi)'].includes(field)&&clean(row[field])&&Array.isArray(row.__workflow)&&
    row.__workflow.length){return `<button class="ct-link" type="button" data-wf-nsr="${esc(row['NSR#'])}" title="View full workflow for ${esc(row['NSR#'])}">${esc(row[field])}</button>`;}
  return esc(row[field]);
}

function cycleCellClass(row,field){
  const elapsed=clean(row['Elapsed Days']),ecdCycle=clean(row['Cycle Time (ECD)']),sobiCycle=clean(row['Cycle Time (SOBi)']); const elapsedOverEcd=elapsed&&ecdCycle&&Number(elapsed)>Number(ecdCycle); const ecdOverSobi=ecdCycle&&sobiCycle&&Number(ecdCycle)>Number(sobiCycle);
  if(field==='Elapsed Days'&&elapsedOverEcd)return 'duration-over';
  if(field==='Cycle Time (ECD)'&&(elapsedOverEcd||ecdOverSobi))return 'duration-over';
  return '';
}

async function openWorkflow(nsr,trigger){
  const row=state.cycleRows.find(item=>upper(item['NSR#'])===upper(nsr)),stages=row&&Array.isArray(row.__workflow)?row.__workflow:[];
  if(!stages.length){toast('Workflow detail is unavailable for this NSR.',true);return}
  const modal=$id('wfModal'),link=$id('wfNsrLink'),href=nsrDetailUrl(row);
  $id('wfTitle').textContent=`${clean(nsr)} Workflow Detail`;$id('wfMeta').textContent=clean(row['NSR Title']);
  link.hidden=!href;if(href)link.href=href;else link.removeAttribute('href');
  $id('wfSummary').innerHTML=['Task Status','Assigned Approver','Pending Days','Submit Date'].map(field=>`<div class="fid-detail-value"><small>${field}</small>${field==='Task Status'&&clean(row[field])?statusPill(clean(row[field])):esc(row[field])||'—'}</div>`).join('');
  wfTrigger=trigger||ownerDocument.activeElement;if(!modal.open)modal.showModal();modalHost.classList.add('wf-open');
  workflowTable=await renderTable('wfTable',WF_FIELDS.map(field=>[field,row=>row[field],value=>field==='Approver Status'?`<span class="pill ${/APPROVED|COMPLETED/.test(upper(value))?'clear':/REJECT|CANCEL/.test(upper(value))?'risk':'info'}">${esc(value)}</span>`:esc(value),{variableHeight:true,...(field==='Pending Days'?{sorter:(a,b)=>(Number(a)||0)-(Number(b)||0)}:{})}]),stages,{title:'Workflow History',rowHeight:false,copyId:'wfCopy',onCopy:copyWorkflow,widths:{'Task Name':160,'Approver Status':130,'Approver Team':140,Approver:120,'Pending Days':112,'Approved Date':125,Comments:360}});
  if(modal.open)$id('wfClose').focus({preventScroll:true});
}

function closeWorkflow(){if($id('wfModal').open)$id('wfModal').close()}
function restoreWorkflowFocus(){modalHost.classList.remove('wf-open');if(wfTrigger?.isConnected)wfTrigger.focus({preventScroll:true});wfTrigger=null}

async function renderCycleTable(){
  cycleTable=await renderTable('cycleTable',CYCLE_FIELDS.map(field=>[field,row=>row[field],(_,row)=>`<span class="${cycleCellClass(row,field)}">${cycleCell(row,field)}</span>`,{title:embedded&&field==='REF FCID'?'FID':field,sorter:(a,b)=>compareCycle(a,b,field)}]),state.cycleRows,{file:'NSR_Cycle_Time',height:450,columnFilters:false,defaultSort:'NSR#',copyId:'copyCycleBtn',exportId:'exportCycleBtn',onCopy:copyCycleRows,onExport:exportCycleRows,widths:Object.fromEntries(CYCLE_FIELDS.map((field,index)=>[field,CYCLE_WIDTHS[index]]))});
}

function copyWorkflow(){if(!workflowTable?.ready)return;const fields=workflowTable.grid.getColumns().filter(column=>column.isVisible()).map(column=>WF_FIELDS[Number(column.getField().slice(1))]),rows=workflowTable.grid.getRows('active').map(row=>workflowTable.rows[row.getIndex()]);if(rows.length)return copyRows([fields,...rows.map(row=>fields.map(field=>row[field]??''))])}
function copyCycleRows(){const rows=cycleRows();if(!rows.length)return;const fields=cycleFields();return copyRows([fields,...rows.map(row=>fields.map(field=>row[field]??''))])}

function exportCycleRows(){
  const rows=cycleRows();if(!rows.length)return;
  if(!window.XLSX){toast('The XLSX library did not load.',true);return}
  const fields=cycleFields(),data=rows.map(row=>Object.fromEntries(fields.map(field=>[field,row[field]??'']))),sheet=XLSX.utils.json_to_sheet(data,{header:fields});
  const workflowData=rows.flatMap(row=>(Array.isArray(row.__workflow)?row.__workflow:[]).map((stage,index)=>Object.fromEntries(WORKFLOW_EXPORT_FIELDS.map(field=>[field,field==='NSR#'?(row['NSR#']??''):field==='Workflow Step'?index+1:(stage[field]??'')]))));
  const workflowSheet=XLSX.utils.json_to_sheet(workflowData,{header:WORKFLOW_EXPORT_FIELDS}); sheet['!cols']=fields.map(field=>({wch:Math.max(10,Math.round(CYCLE_WIDTHS[CYCLE_FIELDS.indexOf(field)]/7))}));sheet['!autofilter']={ref:sheet['!ref']};
  workflowSheet['!cols']=[12,14,32,18,22,24,14,14,50].map(wch=>({wch}));workflowSheet['!autofilter']={ref:workflowSheet['!ref']}; const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'NSR Cycle Time');XLSX.utils.book_append_sheet(book,workflowSheet,
    'Workflow Detail'); const now=new Date(),stamp=`${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`; XLSX.writeFile(book,`NSR_Cycle_Time_${stamp}.xlsx`,{compression:true});
    toast(`${rows.length} summary rows and ${workflowData.length} workflow records exported.`);
}

function cancel(){if(!state.cycleAbort)return;state.cycleAbort.abort();$id('cycleStatus').textContent='Cancelling active requests…';}
const listeners=[];
function bind(element,event,handler){element.addEventListener(event,handler);listeners.push([element,event,handler]);}
bind($id('runCycleBtn'),'click',()=>runCycleLookup());bind($id('cancelCycleBtn'),'click',cancel);
bind($id('cycleTable'),'click',event=>{const trigger=event.target.closest('[data-wf-nsr]');if(trigger)openWorkflow(trigger.dataset.wfNsr,trigger)});
bind($id('wfModal'),'click',event=>{if(event.target.closest('[data-wf-close]'))closeWorkflow()});
bind($id('wfModal'),'close',()=>{if(!$id('wfModal').open)restoreWorkflowFocus()});
setCycleBusy(false);
return {
  setRows(rows){state.rows=Array.isArray(rows)?rows:[]},
  show:()=>renderCycleTable().then(()=>cycleTable.grid.redraw(true)),
  query(nsrs,{run=true}={}){if(nsrs!==undefined)$id('cycleInput').value=Array.isArray(nsrs)?nsrs.join('\n'):String(nsrs);return run?runCycleLookup():parseNsrInput($id('cycleInput').value)},
  snapshot:()=>({rows:state.cycleRows,busy:Boolean(state.cycleAbort)}),cancel,
  destroy(){state.cycleRun++;if(state.cycleAbort)state.cycleAbort.abort();state.cycleAbort=null;setCycleBusy(false);clearTimeout(state.toastTimer);closeWorkflow();restoreWorkflowFocus();cycleTable?.grid.destroy();workflowTable?.grid.destroy();listeners.forEach(([element,event,handler])=>element.removeEventListener(event,handler));}
};
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
  const first=boms[0],second=boms[1]||null,firstOnly=new Set(),secondOnly=new Set();
  if(second){for(const key of first.materials.keys())if(!second.materials.has(key))firstOnly.add(key);for(const key of second.materials.keys())if(!first.materials.has(key))secondOnly.add(key);}
  const documents=new Map(),matched=new Set();
  summaryRows.forEach(row=>{
    const partNumber=normalizedBomPartNumber(row.partNumber),inFirst=second?firstOnly.has(partNumber):first.materials.has(partNumber),inSecond=Boolean(second&&secondOnly.has(partNumber));
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
  const ExcelJS=loadExcelJs();
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
