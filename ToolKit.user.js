// ==UserScript==
// @name         Toolkit Kyra
// @version      26.11.22.1
// @namespace    https://github.com/synalocey/ToolkitKyra
// @description  Toolkit Kyra
// @author       Syna
// @license      GPL-3.0
// @icon64       data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxODAgMTgwIj48cGF0aCBmaWxsPSIjMjQyNDM3IiBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik04OS45IDEwLjYuNSAxNjloMTc4Ljd6TTU4LjEgODcuNCAxNTcuOCAxNjloMjAuN3ptMCAxNS4zIDc4LjggNjYuM2gyMC43em0wIDE1LjQgNTcuOCA1MC45aDIwLjd6bTAgMTUuNEw5NSAxNjloMjAuNnoiLz48L3N2Zz4=
// @downloadURL  https://lamresearch.sharepoint.com/sites/OFMAutomationTools/FID_NSR_Dashboard/FID_NSR_Toolkit.user.js
// @updateURL    https://lamresearch.sharepoint.com/sites/OFMAutomationTools/FID_NSR_Dashboard/FID_NSR_Toolkit.user.js
// @match        *://fep.lamresearch.com/*
// @run-at       document-start
// @noframes
// ==/UserScript==
(function(){
  'use strict';
  if(location.hostname.toLowerCase()!=='fep.lamresearch.com')return;
  const APP_PATH='/k1.png',FCID_PATH='/staceykyra.png',TOOLBOX_PATH='/k2.png';
  const UPGRADE_URL='https://lamresearch.sharepoint.com/sites/OFMAutomationTools/FID_NSR_Dashboard/FID_NSR_Toolkit.user.js';
  const path=location.pathname.toLowerCase();
  if(path===APP_PATH){
    const target=new URL(FCID_PATH,location.href);
    target.search=location.search;
    target.hash='nsr';
    location.replace(target.href);
    return;
  }
  if(path!==FCID_PATH&&path!==TOOLBOX_PATH)return;
  const render=()=>{
    document.documentElement.lang='en';
    document.documentElement.innerHTML='<head><meta charset="utf-8"><title>FID &amp; NSR Dashboard Upgrade Required</title></head>'
      +'<body style="margin:24px;font:14px/1.6 Segoe UI,Arial,sans-serif;color:#1e293b">'
      +'<p style="margin:0">FID &amp; NSR Dashboard has been updated. Please <a id="tkUpgradeLink" href="'+UPGRADE_URL+'" target="_blank" rel="noopener noreferrer" style="color:#2563eb;font-weight:600">click here to update to the latest version</a>, then refresh this page.</p>'
      +'</body>';
  };
  if(document.body)render();
  else document.addEventListener('DOMContentLoaded',render,{once:true});
})();
