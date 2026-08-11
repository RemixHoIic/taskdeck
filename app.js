(async function(){
  "use strict";

  var STORAGE_KEY = "apple_task_manager_v1";
  var COLORS = ["#E24B4A","#F0A23D","#F5CC3E","#34B87A","#2E5BFF","#6C5DD3","#B45DE0","#E0548C","#8A90A6"];

  /* =========================================================
     コンテキスト自動判定（同じ app.html を PCブラウザ／スマホブラウザ／
     拡張機能オーバーレイ(iframe内)の3箇所から使い回すための判定）。
     - IS_OVERLAY: content.js が生成する iframe 内で開かれているか
     - isNarrow(): 画面幅が狭い（スマホ or オーバーレイパネル幅）か
     - isMobilePhoneContext(): iframeではない状態で幅が狭い＝スマホ本体で
       直接開いているケース。この場合だけ Excel/対応履歴/KPT を隠す。
  ========================================================= */
  var IS_OVERLAY = (function(){
    try { return window.self !== window.top; } catch(e){ return true; }
  })();
  function isNarrow(){ return window.innerWidth <= 900; }
  function isMobilePhoneContext(){ return !IS_OVERLAY && isNarrow(); }
  var MOBILE_HIDDEN_MODES = ["excel","history","kpt"];
  function visibleModeIds(){
    if(isMobilePhoneContext()) return MODE_IDS.filter(function(m){ return MOBILE_HIDDEN_MODES.indexOf(m) === -1; });
    return MODE_IDS;
  }

  function uid(){ return Date.now().toString(36) + Math.random().toString(36).slice(2,8); }
  function todayStr(){
    var d = new Date();
    return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
  }
  function getISOWeek(d){
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    var dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
    var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1)/7);
    return d.getUTCFullYear() + "-W" + String(weekNo).padStart(2,"0");
  }
  function deepClone(o){ return JSON.parse(JSON.stringify(o)); }
  function escapeHtml(s){
    return String(s).replace(/[&<>"']/g, function(c){
      return { "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" }[c];
    });
  }
  function ic(path){ return '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'+path+'</svg>'; }

  /* ---------- Excel formula syntax highlighting ----------
     "文字列" -> tok-str / 関数名( -> tok-fn / セル参照・範囲 -> tok-ref / 数値 -> tok-num
     コピー時は t.fields の生テキストを使うので、ハイライトは表示専用（安全にHTML化するだけ）。 */
  var FORMULA_TOKEN_RE = /("[^"]*")|(\b[A-Za-z_][A-Za-z0-9_.]*(?=\())|(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})|(\b\d+(?:\.\d+)?\b)/g;
  function highlightFormula(str){
    if(!str) return "";
    var out = "", lastIndex = 0, m;
    FORMULA_TOKEN_RE.lastIndex = 0;
    while((m = FORMULA_TOKEN_RE.exec(str))){
      out += escapeHtml(str.slice(lastIndex, m.index));
      if(m[1] !== undefined) out += '<span class="tok-str">' + escapeHtml(m[1]) + '</span>';
      else if(m[2] !== undefined) out += '<span class="tok-fn">' + escapeHtml(m[2]) + '</span>';
      else if(m[3] !== undefined) out += '<span class="tok-ref">' + escapeHtml(m[3]) + '</span>';
      else if(m[4] !== undefined) out += '<span class="tok-num">' + escapeHtml(m[4]) + '</span>';
      lastIndex = FORMULA_TOKEN_RE.lastIndex;
    }
    out += escapeHtml(str.slice(lastIndex));
    return out;
  }

  /* =========================================================
     MODE_CONFIG — 各モードの見た目・入力項目・並び方を定義する。
     ここを増やすだけで新モードを追加できるようにするのが狙い。
  ========================================================= */
  var MODE_IDS = ["task","idea","excel","history","kpt"];

  var TAB_ICON = {
    task: ic('<path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/>'),
    idea: ic('<path d="M9 18h6"/><path d="M10 22h4"/><path d="M12 2a7 7 0 00-4 12.7c.5.4.8 1 .8 1.7V17h6.4v-.6c0-.7.3-1.3.8-1.7A7 7 0 0012 2z"/>'),
    excel: ic('<rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/>'),
    history: ic('<rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v10a2 2 0 002 2h10a2 2 0 002-2V8"/><line x1="10" y1="12" x2="14" y2="12"/>'),
    kpt: ic('<polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 014-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 01-4 4H3"/>')
  };
  var ICONS = {
    today: ic('<rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><path d="M8 15l2 2 4-4"/>'),
    upcoming: ic('<rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>'),
    flagged: '<svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="1"><path d="M3 3l3 9-3 9 19-9-19-9z"/></svg>',
    all: ic('<polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11L2 12v6a2 2 0 002 2h16a2 2 0 002-2v-6l-3.45-6.89A2 2 0 0016.76 4H7.24a2 2 0 00-1.79 1.11z"/>')
  };

  var IDEA_STATUS = [
    { val:"", label:"検討中", chipClass:"" },
    { val:"todo", label:"TODO昇格済み", chipClass:"status-todo" },
    { val:"reject", label:"ボツ", chipClass:"status-reject" }
  ];

  var MODE_CONFIG = {
    task: {
      label:"タスク", tabLabel:"タスク", icon:TAB_ICON.task,
      listLabel:"リスト", listSectionLabel:"マイリスト", addListLabel:"リストを追加",
      titlePlaceholder:"タスク名", showNotes:true, notesPlaceholder:"メモを追加",
      emptyHint:"タスクを選択してください", emptyListHint:"タスクはありません",
      addPlaceholder:"新規タスク", autoSelectOnAdd:false,
      hasCheckbox:true, hasDue:true, hasPriority:true, hasFlag:true,
      hasSmartViews:true, defaultView:"today",
      searchFieldsExtra:[]
    },
    idea: {
      label:"アイデア", tabLabel:"アイデア", icon:TAB_ICON.idea,
      listLabel:"カテゴリ", listSectionLabel:"カテゴリ", addListLabel:"カテゴリを追加",
      titlePlaceholder:"アイデア名", showNotes:true, notesPlaceholder:"メモ・詳細を追加",
      emptyHint:"アイデアを選択してください", emptyListHint:"アイデアはありません",
      addPlaceholder:"新規アイデア", autoSelectOnAdd:true,
      hasCheckbox:false, hasDue:false, hasPriority:false, hasFlag:false,
      hasSmartViews:false, defaultView:null,
      sectionBy:"status", sectionOrder:["","todo","reject"],
      sectionLabels:{"":"検討中","todo":"TODO昇格済み","reject":"ボツ"},
      detailFields:[
        { key:"tags", label:"タグ", type:"tags", icon: ic('<path d="M20.59 13.41L11 22.99 1.01 13 10.59 3.41A2 2 0 0112 3H20a2 2 0 012 2v8a2 2 0 01-.59 1.41z"/><circle cx="16" cy="8" r="1.5"/>') }
      ],
      searchFieldsExtra:["fields.tags"]
    },
    excel: {
      label:"Excel関数チートシート", tabLabel:"Excel", icon:TAB_ICON.excel,
      listLabel:"カテゴリ", listSectionLabel:"カテゴリ", addListLabel:"カテゴリを追加",
      titlePlaceholder:"関数名 (例: SUMIFS)", showNotes:false, notesPlaceholder:"",
      emptyHint:"関数を選択してください", emptyListHint:"関数が登録されていません",
      addPlaceholder:"新規関数名", autoSelectOnAdd:true,
      hasCheckbox:false, hasDue:false, hasPriority:false, hasFlag:false,
      hasSmartViews:true, defaultView:null,
      groupByListInAll:true,
      blockFields:[
        { key:"syntax", label:"構文", mono:true, copyable:true, highlight:true, placeholder:"=SUMIFS(合計範囲,条件範囲1,条件1,...)" },
        { key:"example", label:"実例", mono:true, copyable:true, highlight:true, placeholder:"=SUMIFS(C:C,A:A,\"東京\",B:B,\">=2026/01/01\")" }
      ],
      searchFieldsExtra:["fields.syntax","fields.example"]
    },
    history: {
      label:"対応履歴クイック検索", tabLabel:"対応履歴", icon:TAB_ICON.history,
      listLabel:"案件カテゴリ", listSectionLabel:"案件カテゴリ", addListLabel:"カテゴリを追加",
      titlePlaceholder:"対応タイトル", showNotes:false, notesPlaceholder:"",
      emptyHint:"対応履歴を選択してください", emptyListHint:"対応履歴はありません",
      addPlaceholder:"新規対応履歴", autoSelectOnAdd:true,
      hasCheckbox:false, hasDue:false, hasPriority:false, hasFlag:false,
      hasSmartViews:true, defaultView:null,
      groupByListInAll:true,
      detailFields:[
        { key:"tags", label:"タグ", type:"tags", icon: ic('<path d="M20.59 13.41L11 22.99 1.01 13 10.59 3.41A2 2 0 0112 3H20a2 2 0 012 2v8a2 2 0 01-.59 1.41z"/><circle cx="16" cy="8" r="1.5"/>') }
      ],
      blockFields:[
        { key:"content", label:"対応内容メモ", mono:false, copyable:false, placeholder:"対応内容・経緯・結果などを記録" }
      ],
      searchFieldsExtra:["fields.tags","fields.content"]
    },
    kpt: {
      label:"KPT振り返り", tabLabel:"KPT", icon:TAB_ICON.kpt,
      listLabel:"プロジェクト", listSectionLabel:"プロジェクト", addListLabel:"プロジェクトを追加",
      titlePlaceholder:"週 (例: 2026-W32)", showNotes:false, notesPlaceholder:"",
      emptyHint:"週を選択してください", emptyListHint:"振り返りがまだありません",
      addPlaceholder:"新規: 例) 2026-W32", autoSelectOnAdd:true,
      hasCheckbox:false, hasDue:false, hasPriority:false, hasFlag:false,
      hasSmartViews:true, defaultView:null,
      groupByListInAll:true, isWeekTitle:true, sortDesc:true,
      blockFields:[
        { key:"keep", label:"Keep（続けること）", mono:false, copyable:false, placeholder:"うまくいったこと・続けたいこと", accent:"var(--green)" },
        { key:"problem", label:"Problem（課題）", mono:false, copyable:false, placeholder:"うまくいかなかったこと・課題", accent:"var(--red)" },
        { key:"try", label:"Try（次に試すこと）", mono:false, copyable:false, placeholder:"次のアクション・試すこと", accent:"var(--accent)" }
      ],
      searchFieldsExtra:["fields.keep","fields.problem","fields.try"]
    }
  };

  /* =========================================================
     State + migration
  ========================================================= */
  function defaultModeState(mode){
    if(mode === "task") return { lists:[{id:"inbox",name:"リマインダー",color:"#2E5BFF"}], items:[], selectedView:"today", selectedItemId:null };
    if(mode === "idea") return { lists:[{id:"general",name:"一般",color:"#6C5DD3"}], items:[], selectedView:"general", selectedItemId:null };
    if(mode === "excel") return { lists:[{id:"agg",name:"集計系",color:"#34B87A"},{id:"lookup",name:"検索系",color:"#2E5BFF"},{id:"text",name:"文字列系",color:"#F0A23D"}], items:[], selectedView:"all", selectedItemId:null };
    if(mode === "history") return { lists:[{id:"general",name:"一般",color:"#F0A23D"}], items:[], selectedView:"all", selectedItemId:null };
    if(mode === "kpt") return { lists:[{id:"team",name:"チーム",color:"#B45DE0"}], items:[], selectedView:"all", selectedItemId:null };
  }
  function defaultState(){
    var modes = {};
    MODE_IDS.forEach(function(m){ modes[m] = defaultModeState(m); });
    return { theme:"light", activeMode:"task", modes:modes };
  }

  function migrateState(raw){
    if(!raw) return defaultState();
    if(raw.modes){
      MODE_IDS.forEach(function(m){ if(!raw.modes[m]) raw.modes[m] = defaultModeState(m); });
      if(!raw.activeMode) raw.activeMode = "task";
      if(!raw.theme) raw.theme = "light";
      return raw;
    }
    // 旧バージョン（タスクのみ）のデータを新形式へ変換
    if(raw.lists && raw.tasks){
      var s = defaultState();
      s.theme = raw.theme || "light";
      s.modes.task.lists = raw.lists;
      s.modes.task.items = raw.tasks;
      s.modes.task.selectedView = raw.selectedView || "today";
      s.modes.task.selectedItemId = raw.selectedTaskId || null;
      s.activeMode = "task";
      return s;
    }
    return defaultState();
  }

  async function loadState(){
    try{
      var result = await chrome.storage.local.get(STORAGE_KEY);
      var raw = result[STORAGE_KEY];
      if(!raw) return defaultState();
      return migrateState(raw); // chrome.storage.localはJSON文字列化不要（オブジェクトのまま保存/取得される）
    }catch(e){ console.warn("load failed", e); return defaultState(); }
  }
  function saveState(){
    // fire-and-forgetでOK。呼び出し側は同期のままawaitしない（既存コードとの互換のため）。
    chrome.storage.local.set({ [STORAGE_KEY]: state }).catch(function(e){ console.warn("save failed", e); });
  }
  var state = await loadState();

  // 他のウィンドウ（quick-add小窓など）が保存した変更をside panelにも反映する
  chrome.storage.onChanged.addListener(function(changes, area){
    if(area === "local" && changes[STORAGE_KEY]){
      state = changes[STORAGE_KEY].newValue || state;
      render();
    }
  });

  document.documentElement.setAttribute("data-theme", state.theme || "light");

  var searchQuery = "";
  var completedCollapsed = true;

  var $ = function(id){ return document.getElementById(id); };
  var smartListsEl = $("smartLists");
  var myListsEl = $("myLists");
  var taskListScroll = $("taskListScroll");
  var viewTitle = $("viewTitle");
  var viewSub = $("viewSub");
  var detailEmpty = $("detailEmpty");
  var detailScroll = $("detailScroll");
  var detailCard = $("detailCard");

  function cfg(){ return MODE_CONFIG[state.activeMode]; }
  function ms(){ return state.modes[state.activeMode]; }
  function currentItems(){ return ms().items; }

  /* =========================================================
     View / filtering helpers
  ========================================================= */
  function isOverdue(t){ return t.due && !t.completed && t.due < todayStr(); }
  function isToday(t){ return t.due === todayStr() && !t.completed; }
  function isShoppingList(listId){
    var l = ms().lists.find(function(x){ return x.id === listId; });
    return !!(l && l.kind === "shopping");
  }

  function itemsForView(viewId){
    var all = currentItems();
    if(state.activeMode === "task"){
      if(viewId === "today") return all.filter(function(t){ return !t.completed && t.due && t.due <= todayStr(); });
      if(viewId === "upcoming") return all.filter(function(t){ return !t.completed && t.due && t.due > todayStr(); });
      if(viewId === "flagged") return all.filter(function(t){ return t.flagged; });
      if(viewId === "all") return all.slice();
      return all.filter(function(t){ return t.listId === viewId; });
    }
    if(viewId === "all") return all.slice();
    return all.filter(function(t){ return t.listId === viewId; });
  }
  function countForView(viewId){
    var list = itemsForView(viewId);
    if(state.activeMode === "task") return list.filter(function(t){ return !t.completed; }).length;
    return list.length;
  }

  /* =========================================================
     Render: overall
  ========================================================= */
  function render(){
    var vis = visibleModeIds();
    if(vis.indexOf(state.activeMode) === -1) state.activeMode = "task";
    renderModeTabs();
    renderSidebar();
    renderTaskList();
    renderDetail();
    renderStatusBar();
    renderDueBanner();
    renderQuickNav();
  }

  function switchMode(m){
    if(state.activeMode === m) return;
    state.activeMode = m;
    searchQuery = ""; $("searchInput").value = "";
    completedCollapsed = true;
    saveState(); render();
  }

  function renderModeTabs(){
    var el = $("modeTabs");
    el.innerHTML = "";
    visibleModeIds().forEach(function(m){
      var c = MODE_CONFIG[m];
      var tab = document.createElement("div");
      tab.className = "mode-tab" + (state.activeMode === m ? " active" : "");
      tab.innerHTML = c.icon + "<span>"+escapeHtml(c.tabLabel)+"</span>";
      tab.addEventListener("click", function(){
        switchMode(m);
        closeSidebarOnMobile();
      });
      el.appendChild(tab);
    });
  }

  /* =========================================================
     Quick nav bar（スマホ／拡張機能オーバーレイ用の常設ナビ）
     三本線のサイドバーを開かなくても、モード切替とビュー切替が
     その場でできるようにするための簡易バー。
  ========================================================= */
  function renderQuickNav(){
    var qnModes = $("quickNavModes");
    if(!qnModes) return; // 古いapp.htmlとの後方互換
    qnModes.innerHTML = "";
    visibleModeIds().forEach(function(m){
      var c = MODE_CONFIG[m];
      var pill = document.createElement("div");
      pill.className = "qn-mode-pill" + (state.activeMode === m ? " active" : "");
      pill.innerHTML = c.icon + "<span>" + escapeHtml(c.tabLabel) + "</span>";
      pill.addEventListener("click", function(){ switchMode(m); closeQuickNavDropdown(); });
      qnModes.appendChild(pill);
    });

    var meta = currentViewMeta();
    var countTxt = "";
    if(ms().selectedView){
      var n = countForView(ms().selectedView);
      if(n) countTxt = " (" + n + ")";
    }
    $("quickNavViewLabel").textContent = meta.name + countTxt;

    renderQuickNavDropdown();
  }

  function renderQuickNavDropdown(){
    var el = $("quickNavDropdown");
    if(!el) return;
    el.innerHTML = "";
    var rows = [];
    if(state.activeMode === "task"){
      rows.push({ id:"today", name:"今日", icon:ICONS.today });
      rows.push({ id:"upcoming", name:"予定", icon:ICONS.upcoming });
      rows.push({ id:"flagged", name:"フラグ付き", icon:ICONS.flagged });
      rows.push({ id:"all", name:"すべて", icon:ICONS.all });
    } else {
      rows.push({ id:"all", name:"すべて", icon:ICONS.all });
    }
    rows.forEach(function(r){
      var d = document.createElement("div");
      d.className = "qn-drop-item" + (ms().selectedView === r.id ? " active" : "");
      d.innerHTML = r.icon + "<span>" + escapeHtml(r.name) + "</span>";
      d.addEventListener("click", function(){
        ms().selectedView = r.id; ms().selectedItemId = null; saveState(); closeQuickNavDropdown(); render();
      });
      el.appendChild(d);
    });

    if(ms().lists.length){
      var divider = document.createElement("div");
      divider.className = "qn-drop-divider";
      el.appendChild(divider);
    }
    ms().lists.forEach(function(l){
      var d = document.createElement("div");
      d.className = "qn-drop-item" + (ms().selectedView === l.id ? " active" : "");
      d.innerHTML = '<span class="color-dot" style="background:'+l.color+'"></span><span>'+escapeHtml(l.name)+(l.kind==="shopping" ? " 🛒" : "")+'</span>';
      d.addEventListener("click", function(){
        ms().selectedView = l.id; ms().selectedItemId = null; saveState(); closeQuickNavDropdown(); render();
      });
      el.appendChild(d);
    });
  }

  function closeQuickNavDropdown(){
    var dd = $("quickNavDropdown"), vb = $("quickNavView");
    if(dd) dd.classList.remove("show");
    if(vb) vb.classList.remove("open");
  }

  function renderSidebar(){
    var c = cfg();
    $("listSectionLabel").textContent = c.listSectionLabel;
    $("addListLabel").textContent = c.addListLabel;

    smartListsEl.innerHTML = "";
    if(state.activeMode === "task"){
      [
        { id:"today", name:"今日", icon:ICONS.today },
        { id:"upcoming", name:"予定", icon:ICONS.upcoming },
        { id:"flagged", name:"フラグ付き", icon:ICONS.flagged },
        { id:"all", name:"すべて", icon:ICONS.all }
      ].forEach(appendSmartRow);
    } else {
      appendSmartRow({ id:"all", name:"すべて", icon:ICONS.all });
    }

    function appendSmartRow(s){
      var row = document.createElement("div");
      row.className = "nav-item" + (ms().selectedView === s.id ? " active" : "");
      row.innerHTML = s.icon + '<span class="nav-name">'+escapeHtml(s.name)+'</span><span class="nav-count">'+(countForView(s.id) || "")+'</span>';
      row.addEventListener("click", function(){
        ms().selectedView = s.id; ms().selectedItemId = null; saveState(); render();
        closeSidebarOnMobile();
      });
      smartListsEl.appendChild(row);
    }

    myListsEl.innerHTML = "";
    ms().lists.forEach(function(l){
      var row = document.createElement("div");
      row.className = "nav-item" + (ms().selectedView === l.id ? " active" : "");
      row.innerHTML =
        '<span class="color-dot" style="background:'+l.color+'"></span>' +
        '<span class="nav-name">'+escapeHtml(l.name)+(l.kind==="shopping" ? " 🛒" : "")+'</span>' +
        '<span class="nav-count">'+(countForView(l.id) || "")+'</span>' +
        '<button class="nav-edit-btn" title="編集・削除" aria-label="編集・削除"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/></svg></button>';
      row.addEventListener("click", function(){
        ms().selectedView = l.id; ms().selectedItemId = null; saveState(); render();
        closeSidebarOnMobile();
      });
      row.querySelector(".nav-edit-btn").addEventListener("click", function(e){
        e.stopPropagation();
        openListModal(l);
      });
      myListsEl.appendChild(row);
    });
  }

  function currentViewMeta(){
    var c = cfg();
    if(state.activeMode === "task"){
      var smart = { today:"今日", upcoming:"予定", flagged:"フラグ付き", all:"すべて" };
      if(smart[ms().selectedView]) return { name: smart[ms().selectedView] };
    } else if(ms().selectedView === "all"){
      return { name: "すべて" };
    }
    var l = ms().lists.find(function(x){ return x.id === ms().selectedView; });
    if(l) return { name:l.name };
    return { name:c.label };
  }

  /* =========================================================
     Search matching across generic fields
  ========================================================= */
  function itemSearchText(t){
    var parts = [t.title || "", t.notes || ""];
    if(t.fields){
      Object.keys(t.fields).forEach(function(k){
        var v = t.fields[k];
        if(Array.isArray(v)) parts.push(v.join(" "));
        else if(v) parts.push(String(v));
      });
    }
    return parts.join(" ").toLowerCase();
  }

  /* =========================================================
     Render: item list (center column)
  ========================================================= */
  function renderTaskList(){
    var c = cfg();
    var meta = currentViewMeta();
    viewTitle.textContent = meta.name;

    var list = itemsForView(ms().selectedView);
    if(searchQuery.trim()){
      var q = searchQuery.trim().toLowerCase();
      list = currentItems().filter(function(t){ return itemSearchText(t).indexOf(q) > -1; });
      viewTitle.textContent = "検索結果";
    }

    if(state.activeMode === "task"){
      var openCount = list.filter(function(t){ return !t.completed; }).length;
      viewSub.textContent = openCount + " 件の未完了タスク";
    } else {
      viewSub.textContent = list.length + " 件";
    }

    taskListScroll.innerHTML = "";

    var isCustomList = ms().lists.some(function(l){ return l.id === ms().selectedView; });
    if(isCustomList && !searchQuery.trim() && state.activeMode === "task" && isShoppingList(ms().selectedView)){
      var doneCount = currentItems().filter(function(t){ return t.listId === ms().selectedView && t.completed; }).length;
      var resetRow = document.createElement("div");
      resetRow.className = "shopping-reset-row";
      resetRow.innerHTML = '<span>🛒 買い物リスト</span><button type="button" id="shoppingResetBtn"'+(doneCount===0?" disabled":"")+'>チェックをすべて外す'+(doneCount?" ("+doneCount+")":"")+'</button>';
      taskListScroll.appendChild(resetRow);
      resetRow.querySelector("#shoppingResetBtn").addEventListener("click", function(){
        currentItems().forEach(function(t){
          if(t.listId === ms().selectedView && t.completed){ t.completed = false; t.completedAt = null; }
        });
        saveState(); render();
      });
    }

    if(isCustomList && !searchQuery.trim()){
      var quickAdd = document.createElement("div");
      quickAdd.className = "quick-add";
      var placeholder = c.isWeekTitle ? (c.addPlaceholder + " (今週:" + getISOWeek(new Date()) + ")") : c.addPlaceholder;
      if(state.activeMode === "task" && isShoppingList(ms().selectedView)) placeholder = "買うものを追加";
      quickAdd.innerHTML = '<div class="plus-circle">+</div><input type="text" placeholder="'+escapeHtml(placeholder)+'" id="quickAddInput">';
      taskListScroll.appendChild(quickAdd);
      var qi = quickAdd.querySelector("input");
      qi.addEventListener("keydown", function(e){
        if(e.key === "Enter"){
          var val = qi.value.trim();
          if(!val && c.isWeekTitle) val = getISOWeek(new Date());
          if(val){ addItem(val, ms().selectedView); qi.value=""; }
        }
      });
    }

    if(list.length === 0){
      var empty = document.createElement("div");
      empty.className = "empty-hint";
      empty.textContent = c.emptyListHint;
      taskListScroll.appendChild(empty);
      return;
    }

    if(state.activeMode === "task"){
      renderTaskModeList(list);
    } else if(c.sectionBy === "status"){
      renderStatusGroupedList(list, c);
    } else if(c.groupByListInAll && ms().selectedView === "all" && !searchQuery.trim()){
      renderListGroupedList(list, c);
    } else {
      var sorted = sortGeneric(list, c);
      sorted.forEach(function(t){ taskListScroll.appendChild(buildItemRow(t)); });
    }
  }

  function sortGeneric(list, c){
    var copy = list.slice();
    if(c.isWeekTitle){
      copy.sort(function(a,b){ return c.sortDesc ? (a.title < b.title ? 1 : (a.title > b.title ? -1 : 0)) : (a.title < b.title ? -1 : 1); });
    } else {
      copy.sort(function(a,b){ return (b.createdAt||0) - (a.createdAt||0); });
    }
    return copy;
  }

  function renderTaskModeList(list){
    if(isShoppingList(ms().selectedView)){
      // 買い物リスト: チェック済みも消さずそのまま並べる（取り消し線のみ）。まとめてリセットは上のボタンで。
      var sorted = list.slice().sort(function(a,b){ return (a.createdAt||0) - (b.createdAt||0); });
      sorted.forEach(function(t){ taskListScroll.appendChild(buildItemRow(t)); });
      return;
    }

    var open = list.filter(function(t){ return !t.completed; });
    var done = list.filter(function(t){ return t.completed; });

    open.sort(function(a,b){
      if(a.due && b.due) return a.due < b.due ? -1 : (a.due > b.due ? 1 : 0);
      if(a.due) return -1;
      if(b.due) return 1;
      return (b.createdAt||0) - (a.createdAt||0);
    });

    open.forEach(function(t){ taskListScroll.appendChild(buildItemRow(t)); });

    if(done.length > 0){
      var label = document.createElement("div");
      label.className = "section-label" + (completedCollapsed ? " collapsed" : "");
      label.innerHTML = '<span class="chev">▾</span><span>完了済み ('+done.length+')</span>';
      label.addEventListener("click", function(){ completedCollapsed = !completedCollapsed; renderTaskList(); });
      taskListScroll.appendChild(label);
      if(!completedCollapsed){
        done.sort(function(a,b){ return (b.completedAt||0) - (a.completedAt||0); });
        done.forEach(function(t){ taskListScroll.appendChild(buildItemRow(t)); });
      }
    }
  }

  function renderStatusGroupedList(list, c){
    c.sectionOrder.forEach(function(statusVal){
      var group = list.filter(function(t){ return (t.status||"") === statusVal; });
      if(group.length === 0) return;
      var label = document.createElement("div");
      label.className = "section-label static";
      label.innerHTML = '<span>'+escapeHtml(c.sectionLabels[statusVal])+' ('+group.length+')</span>';
      taskListScroll.appendChild(label);
      group.sort(function(a,b){ return (b.createdAt||0) - (a.createdAt||0); });
      group.forEach(function(t){ taskListScroll.appendChild(buildItemRow(t)); });
    });
  }

  function renderListGroupedList(list, c){
    ms().lists.forEach(function(l){
      var group = list.filter(function(t){ return t.listId === l.id; });
      if(group.length === 0) return;
      var label = document.createElement("div");
      label.className = "section-label static";
      label.innerHTML = '<span class="color-dot" style="background:'+l.color+';display:inline-block;"></span> <span>'+escapeHtml(l.name)+' ('+group.length+')</span>';
      taskListScroll.appendChild(label);
      sortGeneric(group, c).forEach(function(t){ taskListScroll.appendChild(buildItemRow(t)); });
    });
  }

  function buildItemRow(t){
    var c = cfg();
    var row = document.createElement("div");
    row.className = "task-row" + (t.completed ? " done" : "") + (ms().selectedItemId === t.id ? " selected" : "");

    if(c.hasCheckbox){
      var cb = document.createElement("div");
      cb.className = "checkbox" + (t.completed ? " done" : "") + (t.priority === "high" && !t.completed ? " priority" : "");
      cb.innerHTML = '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>';
      cb.addEventListener("click", function(e){ e.stopPropagation(); toggleComplete(t.id); });
      row.appendChild(cb);
    } else {
      var l = ms().lists.find(function(x){ return x.id === t.listId; });
      var dotWrap = document.createElement("div");
      dotWrap.className = "list-dot-wrap";
      dotWrap.innerHTML = '<span class="dot" style="background:'+(l?l.color:"#9DA1B5")+'"></span>';
      row.appendChild(dotWrap);
    }

    var body = document.createElement("div");
    body.className = "task-body";
    var titleDiv = document.createElement("div");
    titleDiv.className = "task-title";
    titleDiv.textContent = t.title;
    body.appendChild(titleDiv);

    var meta = document.createElement("div");
    meta.className = "task-meta";

    if(state.activeMode === "task"){
      if(t.flagged){
        var fd = document.createElement("span");
        fd.className = "flag-dot";
        fd.innerHTML = '<svg width="11" height="11" viewBox="0 0 24 24" fill="var(--orange)" stroke="var(--orange)" stroke-width="1"><path d="M3 3l3 9-3 9 19-9-19-9z"/></svg>';
        meta.appendChild(fd);
      }
      if(t.due){
        var chip = document.createElement("span");
        chip.className = "due-chip" + (isOverdue(t) ? " overdue" : (isToday(t) ? " today" : ""));
        var chipText = formatDue(t.due);
        if(t.dueTime) chipText += " " + t.dueTime;
        if(t.repeat) chipText += " ↻";
        chip.textContent = chipText;
        meta.appendChild(chip);
      }
      if(t.notes){
        var np = document.createElement("span");
        np.className = "task-notes-preview";
        np.textContent = t.notes;
        meta.appendChild(np);
      }
    } else if(state.activeMode === "idea"){
      var stDef = IDEA_STATUS.find(function(s){ return s.val === (t.status||""); }) || IDEA_STATUS[0];
      var sChip = document.createElement("span");
      sChip.className = "due-chip" + (stDef.chipClass ? " "+stDef.chipClass : "");
      sChip.textContent = stDef.label;
      meta.appendChild(sChip);
      var tags = (t.fields && t.fields.tags) || [];
      if(tags.length){
        var tagP = document.createElement("span");
        tagP.className = "task-notes-preview";
        tagP.textContent = tags.map(function(x){ return "#"+x; }).join(" ");
        meta.appendChild(tagP);
      }
    } else if(state.activeMode === "excel"){
      var syn = (t.fields && t.fields.syntax) || "";
      if(syn){
        var synP = document.createElement("span");
        synP.className = "task-notes-preview";
        synP.textContent = syn;
        meta.appendChild(synP);
      }
    } else if(state.activeMode === "history"){
      var htags = (t.fields && t.fields.tags) || [];
      if(htags.length){
        var htagP = document.createElement("span");
        htagP.className = "due-chip accent";
        htagP.textContent = htags.map(function(x){ return "#"+x; }).join(" ");
        meta.appendChild(htagP);
      }
      var content = (t.fields && t.fields.content) || "";
      if(content){
        var cp = document.createElement("span");
        cp.className = "task-notes-preview";
        cp.textContent = content;
        meta.appendChild(cp);
      }
    } else if(state.activeMode === "kpt"){
      var keep = (t.fields && t.fields.keep) || "";
      if(keep){
        var kp = document.createElement("span");
        kp.className = "task-notes-preview";
        kp.textContent = "Keep: " + keep;
        meta.appendChild(kp);
      }
    }

    if(meta.children.length) body.appendChild(meta);

    row.appendChild(body);
    row.addEventListener("click", function(){
      ms().selectedItemId = t.id; saveState(); renderTaskList(); renderDetail();
      if(window.innerWidth <= 680){ detailCard.classList.add("show"); }
    });
    return row;
  }

  function formatDue(dateStr){
    var d = new Date(dateStr+"T00:00:00");
    var t = new Date(todayStr()+"T00:00:00");
    var diffDays = Math.round((d-t)/86400000);
    if(diffDays === 0) return "今日";
    if(diffDays === 1) return "明日";
    if(diffDays === -1) return "昨日";
    return (d.getMonth()+1)+"/"+d.getDate();
  }

  /* =========================================================
     Render: detail panel (right column) — config-driven
  ========================================================= */
  function renderDetail(){
    var c = cfg();
    var t = currentItems().find(function(x){ return x.id === ms().selectedItemId; });
    if(!t){
      $("detailEmptyText").textContent = c.emptyHint;
      detailEmpty.style.display = "flex";
      detailScroll.style.display = "none";
      detailCard.classList.remove("show");
      return;
    }
    detailEmpty.style.display = "none";
    detailScroll.style.display = "block";

    var cbWrap = $("detailCheckbox");
    if(c.hasCheckbox){
      cbWrap.style.display = "flex";
      cbWrap.className = "checkbox" + (t.completed ? " done" : "");
      cbWrap.onclick = function(){ toggleComplete(t.id); };
    } else {
      cbWrap.style.display = "none";
      cbWrap.onclick = null;
    }

    var titleEl = $("detailTitle");
    titleEl.value = t.title;
    titleEl.placeholder = c.titlePlaceholder;
    titleEl.className = "detail-title" + (t.completed ? " done" : "");
    autoGrow(titleEl);
    titleEl.oninput = function(){ t.title = titleEl.value; autoGrow(titleEl); saveState(); renderTaskList(); renderSidebar(); };

    var notesEl = $("detailNotes");
    if(c.showNotes){
      notesEl.style.display = "block";
      notesEl.value = t.notes || "";
      notesEl.placeholder = c.notesPlaceholder;
      autoGrow(notesEl);
      notesEl.oninput = function(){ t.notes = notesEl.value; autoGrow(notesEl); saveState(); renderTaskList(); };
    } else {
      notesEl.style.display = "none";
      notesEl.oninput = null;
    }

    renderFieldBlocks(t, c);
    renderFieldGroup(t, c);

    $("deleteTaskBtn").onclick = function(){
      ms().items = ms().items.filter(function(x){ return x.id !== t.id; });
      ms().selectedItemId = null; saveState(); render();
    };

    $("mobileBack").onclick = function(){ detailCard.classList.remove("show"); };
  }

  function renderFieldBlocks(t, c){
    var wrap = $("fieldBlocks");
    wrap.innerHTML = "";
    if(!c.blockFields) return;
    if(!t.fields) t.fields = {};
    c.blockFields.forEach(function(fdef){
      var block = document.createElement("div");
      block.className = "field-block";
      if(fdef.accent) block.style.borderLeftColor = fdef.accent;
      var head = document.createElement("div");
      head.className = "field-block-head";
      var lab = document.createElement("span");
      lab.className = "field-block-label";
      lab.textContent = fdef.label;
      head.appendChild(lab);
      if(fdef.copyable){
        var btn = document.createElement("button");
        btn.className = "copy-btn";
        btn.type = "button";
        btn.textContent = "コピー";
        btn.addEventListener("click", function(){
          copyText(t.fields[fdef.key] || "");
          btn.textContent = "コピーしました";
          btn.classList.add("copied");
          setTimeout(function(){ btn.textContent = "コピー"; btn.classList.remove("copied"); }, 1400);
        });
        head.appendChild(btn);
      }
      block.appendChild(head);

      var ta = document.createElement("textarea");
      ta.rows = 1;
      ta.placeholder = fdef.placeholder || "";
      ta.value = t.fields[fdef.key] || "";

      if(fdef.highlight){
        // シンタックスハイライト: textareaを透明文字にして、色付きの<pre>を裏に重ねる古典的手法。
        var editorWrap = document.createElement("div");
        editorWrap.className = "code-editor";
        var pre = document.createElement("pre");
        pre.className = "code-highlight";
        pre.innerHTML = highlightFormula(ta.value) || "&nbsp;";
        ta.className = "code-input";
        editorWrap.appendChild(pre);
        editorWrap.appendChild(ta);
        block.appendChild(editorWrap);
        wrap.appendChild(block);
        autoGrow(ta);
        ta.addEventListener("input", function(){
          t.fields[fdef.key] = ta.value; autoGrow(ta);
          pre.innerHTML = highlightFormula(ta.value) || "&nbsp;";
          saveState(); renderTaskList();
        });
        ta.addEventListener("scroll", function(){ pre.scrollTop = ta.scrollTop; pre.scrollLeft = ta.scrollLeft; });
      } else {
        ta.className = fdef.mono ? "mono" : "";
        block.appendChild(ta);
        wrap.appendChild(block);
        autoGrow(ta);
        ta.addEventListener("input", function(){
          t.fields[fdef.key] = ta.value; autoGrow(ta); saveState(); renderTaskList();
        });
      }
    });
  }

  function fieldRow(iconSvg, label){
    var row = document.createElement("div");
    row.className = "field-row";
    var ic1 = document.createElement("div");
    ic1.className = "f-icon";
    ic1.innerHTML = iconSvg;
    var lab = document.createElement("div");
    lab.className = "f-label";
    lab.textContent = label;
    row.appendChild(ic1);
    row.appendChild(lab);
    return row;
  }

  var ICON_FLAG = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="var(--red)" stroke-width="2.2"><path d="M3 3l3 9-3 9 19-9-19-9z"/></svg>';
  var ICON_DUE = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>';
  var ICON_PRIORITY = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>';
  var ICON_LIST = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="7" width="18" height="13" rx="2"/><path d="M8 7V4h8v3"/></svg>';
  var ICON_STATUS = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3.2"/><path d="M12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1"/></svg>';
  var ICON_REPEAT = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 014-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 01-4 4H3"/></svg>';

  function renderFieldGroup(t, c){
    var group = $("fieldGroup");
    group.innerHTML = "";
    if(!t.fields) t.fields = {};
    var shopping = state.activeMode === "task" && isShoppingList(t.listId);

    if(c.hasFlag && !shopping){
      var row = fieldRow(ICON_FLAG, "フラグ");
      var toggle = document.createElement("div");
      toggle.className = "toggle" + (t.flagged ? " on" : "");
      toggle.innerHTML = '<div class="knob"></div>';
      toggle.onclick = function(){ t.flagged = !t.flagged; saveState(); render(); };
      row.appendChild(toggle);
      group.appendChild(row);
    }

    if(c.hasDue && !shopping){
      var rowD = fieldRow(ICON_DUE, "期限");
      var dueWrap = document.createElement("div");
      dueWrap.className = "due-time-wrap";
      var dueInput = document.createElement("input");
      dueInput.type = "date";
      dueInput.value = t.due || "";
      var timeInput = document.createElement("input");
      timeInput.type = "time";
      timeInput.value = t.dueTime || "";
      timeInput.disabled = !t.due;
      dueInput.onchange = function(){
        t.due = dueInput.value || null;
        if(!t.due){ t.dueTime = null; }
        saveState(); render();
      };
      timeInput.onchange = function(){ t.dueTime = timeInput.value || null; saveState(); render(); };
      dueWrap.appendChild(dueInput);
      dueWrap.appendChild(timeInput);
      rowD.appendChild(dueWrap);
      group.appendChild(rowD);

      var rowR = fieldRow(ICON_REPEAT, "繰り返し");
      var segR = document.createElement("div");
      segR.className = "segmented";
      [["","なし"],["daily","毎日"],["weekly","毎週"],["monthly","毎月"]].forEach(function(pair){
        var btn = document.createElement("button");
        btn.type = "button";
        btn.textContent = pair[1];
        btn.classList.toggle("active", pair[0] === (t.repeat||""));
        btn.onclick = function(){ t.repeat = pair[0] || null; saveState(); renderTaskList(); renderDetail(); };
        segR.appendChild(btn);
      });
      rowR.appendChild(segR);
      group.appendChild(rowR);
    }

    if(c.hasPriority && !shopping){
      var rowP = fieldRow(ICON_PRIORITY, "優先度");
      var seg = document.createElement("div");
      seg.className = "segmented";
      [["","なし"],["low","低"],["mid","中"],["high","高"]].forEach(function(pair){
        var btn = document.createElement("button");
        btn.textContent = pair[1];
        btn.type = "button";
        btn.classList.toggle("active", pair[0] === (t.priority||""));
        btn.onclick = function(){ t.priority = pair[0] || null; saveState(); renderTaskList(); renderDetail(); };
        seg.appendChild(btn);
      });
      rowP.appendChild(seg);
      group.appendChild(rowP);
    }

    if(state.activeMode === "idea"){
      var rowS = fieldRow(ICON_STATUS, "ステータス");
      var sel = document.createElement("select");
      sel.className = "plain-select";
      IDEA_STATUS.forEach(function(s){
        var opt = document.createElement("option");
        opt.value = s.val; opt.textContent = s.label;
        if((t.status||"") === s.val) opt.selected = true;
        sel.appendChild(opt);
      });
      sel.onchange = function(){ t.status = sel.value || ""; saveState(); renderTaskList(); renderDetail(); };
      rowS.appendChild(sel);
      group.appendChild(rowS);
    }

    if(c.detailFields){
      c.detailFields.forEach(function(fdef){
        if(fdef.type === "tags"){
          var rowT = fieldRow(fdef.icon || ICON_STATUS, fdef.label);
          var input = document.createElement("input");
          input.type = "text";
          input.placeholder = "カンマ区切り";
          input.value = (t.fields[fdef.key] || []).join(", ");
          input.onchange = function(){
            t.fields[fdef.key] = input.value.split(",").map(function(s){ return s.trim(); }).filter(Boolean);
            saveState(); renderTaskList();
          };
          rowT.appendChild(input);
          group.appendChild(rowT);
        }
      });
    }

    var rowL = fieldRow(ICON_LIST, c.listLabel);
    var listSelect = document.createElement("select");
    listSelect.className = "plain-select";
    ms().lists.forEach(function(l){
      var opt = document.createElement("option");
      opt.value = l.id; opt.textContent = l.name;
      if(l.id === t.listId) opt.selected = true;
      listSelect.appendChild(opt);
    });
    listSelect.onchange = function(){ t.listId = listSelect.value; saveState(); render(); };
    rowL.appendChild(listSelect);
    group.appendChild(rowL);
  }

  /* =========================================================
     Status bar
  ========================================================= */
  function renderStatusBar(){
    var c = cfg();
    var listCount = ms().lists.length;
    if(state.activeMode === "task"){
      var openCount = ms().items.filter(function(t){ return !t.completed; }).length;
      $("statusText").textContent = listCount + "件のリスト ・ " + openCount + "件の未完了タスク";
    } else {
      $("statusText").textContent = listCount + "件の" + c.listLabel + " ・ " + ms().items.length + "件のアイテム";
    }
  }

  function autoGrow(el){ el.style.height = "auto"; el.style.height = el.scrollHeight + "px"; }

  /* =========================================================
     Mutations
  ========================================================= */
  function nextRepeatDate(dueStr, repeat){
    var base = dueStr ? new Date(dueStr+"T00:00:00") : new Date(todayStr()+"T00:00:00");
    var today = new Date(todayStr()+"T00:00:00");
    if(base < today) base = today;
    if(repeat === "daily") base.setDate(base.getDate()+1);
    else if(repeat === "weekly") base.setDate(base.getDate()+7);
    else if(repeat === "monthly") base.setMonth(base.getMonth()+1);
    return base.getFullYear()+"-"+String(base.getMonth()+1).padStart(2,"0")+"-"+String(base.getDate()).padStart(2,"0");
  }

  function toggleComplete(id){
    var t = currentItems().find(function(x){ return x.id === id; });
    if(!t) return;
    if(!t.completed && t.repeat && state.activeMode === "task"){
      // 繰り返しタスク: 完了にせず次回の期限へ進める
      t.due = nextRepeatDate(t.due, t.repeat);
      saveState(); render();
      return;
    }
    t.completed = !t.completed;
    t.completedAt = t.completed ? Date.now() : null;
    saveState(); render();
  }

  function addItem(title, listId){
    var c = cfg();
    var m = ms();
    var finalTitle = title;
    if(c.isWeekTitle){
      finalTitle = title || getISOWeek(new Date());
      var dup = m.items.find(function(x){ return x.listId === (listId||m.lists[0].id) && x.title === finalTitle; });
      if(dup){ m.selectedItemId = dup.id; saveState(); renderTaskList(); renderDetail(); return; }
    }
    var t = {
      id: uid(), title: finalTitle, notes: "",
      listId: listId || (m.lists[0] && m.lists[0].id),
      createdAt: Date.now()
    };
    if(state.activeMode === "task"){
      t.due = m.selectedView === "today" ? todayStr() : null;
      t.dueTime = null; t.repeat = null;
      t.flagged = m.selectedView === "flagged";
      t.priority = null; t.completed = false; t.completedAt = null;
    } else if(state.activeMode === "idea"){
      t.status = ""; t.fields = { tags: [] };
    } else if(state.activeMode === "excel"){
      t.fields = { syntax:"", example:"" };
    } else if(state.activeMode === "history"){
      t.fields = { content:"", tags:[] };
    } else if(state.activeMode === "kpt"){
      t.fields = { keep:"", problem:"", "try":"" };
    }
    m.items.push(t);
    if(c.autoSelectOnAdd) m.selectedItemId = t.id;
    saveState(); renderTaskList(); renderSidebar(); renderStatusBar();
    if(c.autoSelectOnAdd) renderDetail();
  }

  function copyText(str){
    if(navigator.clipboard && navigator.clipboard.writeText){
      navigator.clipboard.writeText(str).catch(function(){ fallbackCopy(str); });
    } else { fallbackCopy(str); }
  }
  function fallbackCopy(str){
    var ta = document.createElement("textarea");
    ta.value = str; ta.style.position = "fixed"; ta.style.left = "-9999px";
    document.body.appendChild(ta); ta.focus(); ta.select();
    try{ document.execCommand("copy"); }catch(e){}
    document.body.removeChild(ta);
  }

  /* =========================================================
     List (category) modal — shared across modes
  ========================================================= */
  var modalOverlay = null;
  function openListModal(editingList){
    closeModal();
    var c = cfg();
    modalOverlay = document.createElement("div");
    modalOverlay.className = "modal-overlay";
    var chosenColor = editingList ? editingList.color : COLORS[4];
    var chosenKind = editingList ? (editingList.kind || "todo") : "todo";
    var shoppingToggleHtml = "";
    if(state.activeMode === "task"){
      shoppingToggleHtml =
        '<div class="shopping-toggle-row" id="shoppingToggleRow">' +
          '<div class="shopping-toggle-text">' +
            '<div class="shopping-toggle-title">🛒 買い物リストにする</div>' +
            '<div class="shopping-toggle-sub">チェックした項目は消えずに残り、ボタン一つで全部リセットして次回また使えます</div>' +
          '</div>' +
          '<div class="toggle' + (chosenKind === "shopping" ? " on" : "") + '" id="shoppingKindToggle"><div class="knob"></div></div>' +
        '</div>';
    }
    modalOverlay.innerHTML =
      '<div class="modal-card">' +
        '<h3>'+(editingList ? (c.listLabel+"を編集") : (c.listLabel+"を新規作成"))+'</h3>' +
        '<input type="text" id="listNameInput" placeholder="'+escapeHtml(c.listLabel)+'名" value="'+(editingList?escapeHtml(editingList.name):"")+'">' +
        '<div class="color-swatches" id="colorSwatches"></div>' +
        shoppingToggleHtml +
        '<div class="modal-buttons">' +
          (editingList ? '<button class="btn btn-danger" id="deleteListBtn" style="margin-right:auto;">削除</button>' : '') +
          '<button class="btn btn-secondary" id="cancelListBtn">キャンセル</button>' +
          '<button class="btn btn-primary" id="saveListBtn">保存</button>' +
        '</div>' +
      '</div>';
    document.body.appendChild(modalOverlay);

    var swatchWrap = modalOverlay.querySelector("#colorSwatches");
    COLORS.forEach(function(cColor){
      var sw = document.createElement("div");
      sw.className = "swatch" + (cColor === chosenColor ? " selected" : "");
      sw.style.background = cColor;
      sw.addEventListener("click", function(){
        swatchWrap.querySelectorAll(".swatch").forEach(function(s){ s.classList.remove("selected"); });
        sw.classList.add("selected");
        chosenColor = cColor;
      });
      swatchWrap.appendChild(sw);
    });

    modalOverlay.querySelector("#cancelListBtn").addEventListener("click", closeModal);
    modalOverlay.addEventListener("click", function(e){ if(e.target === modalOverlay) closeModal(); });

    var kindToggleEl = modalOverlay.querySelector("#shoppingKindToggle");
    if(kindToggleEl){
      kindToggleEl.addEventListener("click", function(){
        chosenKind = (chosenKind === "shopping") ? "todo" : "shopping";
        kindToggleEl.classList.toggle("on", chosenKind === "shopping");
      });
    }

    modalOverlay.querySelector("#saveListBtn").addEventListener("click", function(){
      var name = modalOverlay.querySelector("#listNameInput").value.trim();
      if(!name) return;
      if(editingList){ editingList.name = name; editingList.color = chosenColor; editingList.kind = chosenKind; }
      else{
        var newList = { id: uid(), name: name, color: chosenColor, kind: chosenKind };
        ms().lists.push(newList);
        ms().selectedView = newList.id;
      }
      saveState(); closeModal(); render();
    });

    if(editingList){
      modalOverlay.querySelector("#deleteListBtn").addEventListener("click", function(){
        if(ms().lists.length <= 1){ alert("最後の"+c.listLabel+"は削除できません"); return; }
        ms().lists = ms().lists.filter(function(l){ return l.id !== editingList.id; });
        ms().items = ms().items.filter(function(t){ return t.listId !== editingList.id; });
        if(ms().selectedView === editingList.id) ms().selectedView = c.hasSmartViews ? (c.defaultView || "all") : ms().lists[0].id;
        saveState(); closeModal(); render();
      });
    }
    setTimeout(function(){
      var input = modalOverlay && modalOverlay.querySelector("#listNameInput");
      if(input) input.focus();
    }, 10);
  }
  function closeModal(){ if(modalOverlay){ modalOverlay.remove(); modalOverlay = null; } }

  function updateThemeLabel(){
    var isDark = state.theme === "dark";
    $("themeLabel").textContent = isDark ? "ダークモード" : "ライトモード";
    $("themeIcon").innerHTML = isDark
      ? '<path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/>'
      : '<circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>';
  }
  updateThemeLabel();

  /* =========================================================
     Due notifications (task mode only)
  ========================================================= */
  var bannerDismissedDate = null;

  function dueTasksNow(){
    if(state.activeMode !== "task") return [];
    return ms().items.filter(function(t){ return !t.completed && t.due && t.due <= todayStr(); });
  }

  function renderDueBanner(){
    var banner = $("dueBanner");
    if(state.activeMode !== "task"){ banner.style.display = "none"; return; }
    var due = dueTasksNow();
    if(due.length === 0 || bannerDismissedDate === todayStr()){
      banner.style.display = "none";
      return;
    }
    banner.style.display = "flex";
    var overdueCount = due.filter(function(t){ return t.due < todayStr(); }).length;
    var text = due.length + "件のタスクが期限です";
    if(overdueCount > 0) text += "（うち期限切れ " + overdueCount + "件）";
    $("dueBannerText").textContent = text;
  }

  $("addListBtn").addEventListener("click", function(){ openListModal(null); closeSidebarOnMobile(); });

  $("dueBannerClose").addEventListener("click", function(){
    bannerDismissedDate = todayStr();
    $("dueBanner").style.display = "none";
  });
  $("dueBannerLink").addEventListener("click", function(){
    ms().selectedView = "today"; ms().selectedItemId = null; saveState(); render();
    closeSidebarOnMobile();
  });

  $("themeToggle").addEventListener("click", function(){
    state.theme = (state.theme === "dark") ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", state.theme);
    updateThemeLabel();
    saveState();
  });

  function isMobile(){ return window.innerWidth <= 900; }
  function openSidebarOnMobile(){
    $("sidebar").classList.remove("collapsed");
    if(isMobile()) $("sidebarBackdrop").classList.add("show");
  }
  function closeSidebarOnMobile(){
    if(isMobile()){
      $("sidebar").classList.add("collapsed");
      $("sidebarBackdrop").classList.remove("show");
    }
  }
  if(isMobile()) $("sidebar").classList.add("collapsed");

  $("sidebarToggle").addEventListener("click", function(){
    if($("sidebar").classList.contains("collapsed")) openSidebarOnMobile();
    else closeSidebarOnMobile();
  });
  $("sidebarBackdrop").addEventListener("click", closeSidebarOnMobile);
  window.addEventListener("resize", function(){
    if(!isMobile()){
      $("sidebar").classList.remove("collapsed");
      $("sidebarBackdrop").classList.remove("show");
    }else if(!$("sidebar").classList.contains("collapsed") && !$("sidebarBackdrop").classList.contains("show")){
      $("sidebar").classList.add("collapsed");
    }
  });

  $("searchInput").addEventListener("input", function(e){ searchQuery = e.target.value; renderTaskList(); });

  if($("quickNavView")){
    $("quickNavView").addEventListener("click", function(e){
      e.stopPropagation();
      $("quickNavDropdown").classList.toggle("show");
      $("quickNavView").classList.toggle("open");
    });
    document.addEventListener("click", function(e){
      var dd = $("quickNavDropdown");
      if(dd && dd.classList.contains("show") && !dd.contains(e.target)){
        closeQuickNavDropdown();
      }
    });
  }

  $("exportBtn").addEventListener("click", function(){
    var blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = "taskdeck_backup_" + todayStr() + ".json";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  $("importBtn").addEventListener("click", function(){ $("importFile").click(); });
  $("importFile").addEventListener("change", function(e){
    var file = e.target.files[0];
    if(!file) return;
    var reader = new FileReader();
    reader.onload = function(evt){
      try{
        var imported = JSON.parse(evt.target.result);
        var migrated = migrateState(imported);
        if(migrated){
          state = migrated;
          document.documentElement.setAttribute("data-theme", state.theme || "light");
          updateThemeLabel();
          saveState(); render();
          alert("読み込みが完了しました");
        }else{ alert("形式が正しくありません"); }
      }catch(err){ alert("読み込みに失敗しました: " + err.message); }
      $("importFile").value = "";
    };
    reader.readAsText(file);
  });

  render();
})();
