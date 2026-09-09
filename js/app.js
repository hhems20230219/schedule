$(function(){

  /* =========================================================
     01. 系統初始化與共用物件
     ---------------------------------------------------------
     - 網頁本體放在 GitHub Pages。
     - Google Apps Script 只當 API。
     - Google Sheet 只使用一張 BoardData 工作表。
     - dataSource.useLocalData 用 true / false 切換本地與線上資料。
  ========================================================= */
  const importModal = new bootstrap.Modal($('#importModal')[0]);
  const masterDataModal = new bootstrap.Modal($('#masterDataModal')[0]);
  const dataEditorModal = new bootstrap.Modal($('#dataEditorModal')[0]);
  const volunteerBatchModal = new bootstrap.Modal($('#volunteerBatchModal')[0]);
  const appToast = new bootstrap.Toast($('#appToast')[0], {delay:1800});

  let excelRows = [];
  let excelHeaders = [];
  let todayRoster = [];
  let vehicleMaster = [];
  let dutySchedule = [];
  let dutyStatusByNo = new Map();
  // 前端自行維護的人員（例如中隊長、義消）會直接合併到人員池。
  let manualPersonnel = [];
  // 人員本身的職稱／身分，用於顯示與顏色；不是職務池。
  let roleMaster = [];
  // 職務池只放可拖曳的職務文字，例如「火警值班」。
  let dutyMaster = [];
  // 每次匯入 Excel 時，直接以該檔案內的「火警出動人員車輛分配表」作為當日基礎配置。
  let importedBaseAssignments = [];
  let appConfig = null;
  let currentDutyKey = '';
  // UI-only 測試時段；空字串代表依電腦現在時間自動切換。
  // 不寫入 boardState，避免測試時段影響其他使用者。
  let dutyPeriodOverrideKey = '';
  let hasDetailedDutyData = false;

  // API / 即時同步狀態。
  let autoSaveTimer = null;
  let pollTimer = null;
  let isSaving = false;
  let saveQueued = false;
  let isDragging = false;
  let lastServerStateJson = '';

  /* =========================================================
     02. 基本 UI 訊息工具
     ---------------------------------------------------------
     所有短暫操作結果統一由 Toast 顯示，避免散落 alert。
  ========================================================= */
  function toast(message){
    $('#toastBody').text(message);
    appToast.show();
  }

  function setSyncStatus(text,cssClass){
    const $target = $('#syncStatus');
    $target
      .removeClass('text-success text-danger text-warning text-secondary')
      .addClass(cssClass || 'text-secondary')
      .text(text);
  }

  /* =========================================================
     03. 本地 / 線上資料來源切換
     ---------------------------------------------------------
     data/board-data.json：
     - useLocalData = true  → 只使用瀏覽器 localStorage。
     - useLocalData = false → 只使用 GAS + Google Sheet。

     這個開關只決定「資料存在哪裡」，不改變看板操作流程。
     本地模式適合離線開發與測試；正式 GitHub Pages 請改 false。
  ========================================================= */
  function useLocalData(){
    return appConfig?.dataSource?.useLocalData === true;
  }

  function gasApiUrl(){
    const url = String(appConfig?.api?.gasWebAppUrl || '').trim();
    if(!/^https:\/\/script\.google\.com\/macros\/s\//i.test(url)) return '';
    return url;
  }

  function isGasApiConfigured(){
    return !!gasApiUrl();
  }

  function requestTimeoutMs(){
    const value = Number(appConfig?.api?.requestTimeoutMs || 15000);
    return Number.isFinite(value) && value >= 3000 ? value : 15000;
  }

  function pollIntervalMs(){
    const value = Number(appConfig?.api?.pollIntervalMs || 10000);
    return Number.isFinite(value) && value >= 3000 ? value : 10000;
  }

  function localStorageKey(){
    return 'fireBoard.boardState';
  }

  /* =========================================================
     03-1. 開頁自動全螢幕
     ---------------------------------------------------------
     瀏覽器基於安全限制，requestFullscreen 通常必須由使用者操作觸發。
     因此先在開頁時直接嘗試；若被瀏覽器阻擋，第一次點擊／按鍵時
     立即補進全螢幕，不需要再特別按「全螢幕」按鈕。
  ========================================================= */
  let autoFullscreenRetryBound = false;

  function requestAppFullscreen(){
    if(document.fullscreenElement || !document.documentElement.requestFullscreen){
      return Promise.resolve();
    }
    return document.documentElement.requestFullscreen().catch(()=>{});
  }

  function setupAutoFullscreen(){
    if(appConfig?.app?.autoFullscreen === false) return;

    requestAppFullscreen();

    if(autoFullscreenRetryBound) return;
    autoFullscreenRetryBound = true;

    const retry = function(){
      requestAppFullscreen();
      document.removeEventListener('pointerdown',retry,true);
      document.removeEventListener('keydown',retry,true);
      autoFullscreenRetryBound = false;
    };

    document.addEventListener('pointerdown',retry,true);
    document.addEventListener('keydown',retry,true);
  }

  /* =========================================================
     04. API 共用 fetch
     ---------------------------------------------------------
     POST 使用 text/plain，避免 GitHub Pages 跨網域送到 GAS 時
     因 application/json 產生額外的 CORS preflight。
  ========================================================= */
  async function fetchWithTimeout(url,options={}){
    const controller = new AbortController();
    const timer = setTimeout(()=>controller.abort(),requestTimeoutMs());

    try{
      const response = await fetch(url,{
        ...options,
        cache:'no-store',
        signal:controller.signal
      });

      if(!response.ok){
        throw new Error(`HTTP ${response.status}`);
      }

      return await response.json();
    }finally{
      clearTimeout(timer);
    }
  }

  async function apiGetBoardState(){
    const url = `${gasApiUrl()}?action=get&t=${Date.now()}`;
    return await fetchWithTimeout(url,{method:'GET'});
  }

  async function apiSaveBoardState(state){
    return await fetchWithTimeout(gasApiUrl(),{
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'},
      body:JSON.stringify({action:'save',state})
    });
  }

  /* =========================================================
     05. 看板狀態序列化
     ---------------------------------------------------------
     Google Sheet 不拆多張資料表，直接把目前完整看板存成一份
     boardState JSON，包含人員名冊、勤務時段與每一個配置格。
  ========================================================= */
  function getBoardTargets(){
    return $('.drop-target').toArray();
  }

  function serializeDragItem($item){
    if(!$item || !$item.length) return null;

    const type = String($item.attr('data-drag-type') || '').trim();
    const value = String($item.attr('data-value') || $item.text().trim()).trim();
    if(!type || !value) return null;

    const result = {type,value};
    const autoSource = String($item.attr('data-auto-source') || '').trim();
    if(autoSource) result.autoSource = autoSource;

    if(type === 'person'){
      result.no = $item.attr('data-no') ?? '';
      result.role = normalizeRole($item.attr('data-role') || findRosterRole(value));
      const dutyRole = String($item.attr('data-duty-role') || '').trim();
      if(dutyRole) result.dutyRole = dutyRole;
      const personSource = String($item.attr('data-person-source') || '').trim();
      if(personSource) result.personSource = personSource;
    }

    return result;
  }

  function localDateText(){
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth()+1).padStart(2,'0');
    const day = String(now.getDate()).padStart(2,'0');
    return `${year}-${month}-${day}`;
  }

  function collectBoardState(){
    const slots = getBoardTargets().map((target,index)=>{
      const $target = $(target);
      const item = serializeDragItem($target.children('[data-drag-type]').first());

      return {
        index,
        kind:String($target.attr('data-kind') || ''),
        item
      };
    });

    const rest = $('#restingBody [data-drag-type="person"]').map(function(){
      return serializeDragItem($(this));
    }).get().filter(Boolean);

    const official = $('#officialBody [data-drag-type="person"]').map(function(){
      return serializeDragItem($(this));
    }).get().filter(Boolean);

    return {
      version:13,
      date:localDateText(),
      todayRoster:todayRoster.map(item=>({
        no:item.no ?? '',
        name:item.name,
        role:normalizeRole(item.role),
        start:item.start || '',
        end:item.end || ''
      })),
      dutySchedule:dutySchedule.map(period=>({
        start:period.start,
        end:period.end,
        '備勤91':[...(period['備勤91'] || [])],
        '備勤救災':[...(period['備勤救災'] || [])],
        '值班':[...(period['值班'] || [])],
        '在隊備勤':[...(period['在隊備勤'] || [])],
        '休息時間':[...(period['休息時間'] || [])],
        allDutyNumbers:[...(period.allDutyNumbers || [])]
      })),
      manualPersonnel:manualPersonnel.map(item=>({id:item.id,no:item.no ?? '',name:item.name,role:item.role})),
      roleMaster:[...roleMaster],
      dutyMaster:[...dutyMaster],
      importedBaseAssignments:importedBaseAssignments.map(item=>({...item})),
      vehicleMaster:vehicleMaster.map(item=>typeof item === 'string' ? {name:item,type:'車輛'} : {name:item.name,type:item.type || '車輛'}),
      dutyStatuses:Object.fromEntries(
        [...dutyStatusByNo.entries()].map(([no,statuses])=>[String(no),[...(statuses || [])]])
      ),
      currentDutyKey,
      board:{slots,rest,official}
    };
  }

  /* =========================================================
     06. 從 Google Sheet 還原看板
     ---------------------------------------------------------
     第一次完全沒有 boardState 時才是全白；之後任何人開啟
     GitHub Pages 都先讀取值班台最後一次同步的完整狀態。
  ========================================================= */
  function createRestoredItem(item){
    if(!item || !item.type || !item.value) return null;

    if(item.type === 'vehicle'){
      const $vehicle = $('<div class="vehicle-chip"></div>')
        .attr('data-drag-type','vehicle')
        .attr('data-value',item.value)
        .text(item.value);
      if(item.autoSource) $vehicle.attr('data-auto-source',item.autoSource);
      return $vehicle;
    }

    const restoredSource = String(item.personSource || (item.dutyRole ? 'duty' : 'daily')).trim();
    if(restoredSource === 'daily') ensureRosterPerson(item.value,item.role,item.no);

    const $chip = $('<div class="person-chip"></div>')
      .attr('data-drag-type','person')
      .attr('data-value',item.value)
      .attr('data-no',item.no ?? '')
      .attr('data-person-source',restoredSource)
      .text(item.value);

    if(item.dutyRole) $chip.attr('data-duty-role',item.dutyRole);
    if(restoredSource === 'duty') $chip.addClass('duty-assignment-chip');

    if(item.autoSource) $chip.attr('data-auto-source',item.autoSource);
    if(restoredSource !== 'duty') applyPersonRole($chip,item.value,item.role);
    return $chip;
  }

  function createRestoredStatusItem(item){
    if(!item || item.type !== 'person' || !item.value) return null;

    const restoredSource = String(item.personSource || (item.dutyRole ? 'duty' : 'daily')).trim();
    if(restoredSource === 'daily') ensureRosterPerson(item.value,item.role,item.no);

    const $chip = $('<div class="status-chip"></div>')
      .attr('data-drag-type','person')
      .attr('data-value',item.value)
      .attr('data-no',item.no ?? '')
      .attr('data-person-source',restoredSource)
      .append($('<span></span>').text(item.value));
    if(item.dutyRole) $chip.attr('data-duty-role',item.dutyRole);

    applyPersonRole($chip,item.value,item.role);
    return $chip;
  }

  function restoreBoardState(state){
    clearBoardAssignments();

    if(!state || typeof state !== 'object'){
      todayRoster = [];
      dutySchedule = [];
      dutyStatusByNo = new Map();
      currentDutyKey = '';
      hasDetailedDutyData = false;

      // 尚未匯入勤務表時，先顯示固定的基礎火警配置。
      // 匯入後會再依當日勤務、休息與請假狀態覆蓋。
      applyBaseAssignments({onlyEmpty:true,ignoreDuty:true});
      $('#currentDutyPeriod').text('尚未匯入');
      syncAll();
      return;
    }

    // v34 以前的已儲存狀態沒有「在隊備勤／休息時間／全勤務番號」。
    // 舊資料仍可正常開啟，但在重新匯入新版勤務表前，不把整池人員誤判成未列勤務。
    hasDetailedDutyData = Array.isArray(state.dutySchedule)
      && state.dutySchedule.some(period=>Object.prototype.hasOwnProperty.call(period || {},'allDutyNumbers'));

    todayRoster = Array.isArray(state.todayRoster)
      ? state.todayRoster.map(item=>({
          no:item.no ?? '',
          name:String(item.name || '').trim(),
          role:normalizeRole(item.role),
          start:item.start || '',
          end:item.end || ''
        })).filter(item=>item.name)
      : [];

    dutySchedule = Array.isArray(state.dutySchedule)
      ? state.dutySchedule.map(period=>({
          start:period.start,
          end:period.end,
          '備勤91':[...(period['備勤91'] || [])],
          '備勤救災':[...(period['備勤救災'] || [])],
          '值班':[...(period['值班'] || [])],
          '在隊備勤':[...(period['在隊備勤'] || [])],
          '休息時間':[...(period['休息時間'] || [])],
          allDutyNumbers:[...(period.allDutyNumbers || [])]
        }))
      : [];

    dutyStatusByNo = new Map();
    if(state.dutyStatuses && typeof state.dutyStatuses === 'object'){
      Object.entries(state.dutyStatuses).forEach(([no,statuses])=>{
        const list = Array.isArray(statuses) ? statuses : [statuses];
        dutyStatusByNo.set(String(no),list.map(item=>String(item || '').trim()).filter(Boolean));
      });
    }

    manualPersonnel = Array.isArray(state.manualPersonnel)
      ? state.manualPersonnel
          .filter(item=>item && item.name)
          .map(item=>({
            id:item.id || `manual-${Date.now()}-${Math.random()}`,
            no:item.no ?? '',
            name:String(item.name).trim(),
            role:String(item.role || '隊員').trim(),
            isAttending:item.isAttending !== false
          }))
      : [];

    roleMaster = Array.isArray(state.roleMaster) && state.roleMaster.length
      ? state.roleMaster.map(item=>String(item || '').trim()).filter(Boolean)
      : (Array.isArray(appConfig?.roles) ? appConfig.roles.map(item=>String(item || '').trim()).filter(Boolean) : ['隊員','小隊長','役男','分隊長','中隊長','義消']);

    dutyMaster = Array.isArray(state.dutyMaster) && state.dutyMaster.length
      ? state.dutyMaster.map(item=>String(item || '').trim()).filter(Boolean)
      : (Array.isArray(appConfig?.dutyPool) ? appConfig.dutyPool.map(item=>String(item || '').trim()).filter(Boolean) : ['火警值班']);

    importedBaseAssignments = Array.isArray(state.importedBaseAssignments) && state.importedBaseAssignments.length
      ? state.importedBaseAssignments.map(item=>({...item}))
      : (Array.isArray(appConfig?.baseAssignments) ? appConfig.baseAssignments.map(item=>({...item})) : []);

    // 相容 v41：舊版把中隊長／義消／火警值班當成「特殊職務人員」。
    // v42 起中隊長、義消都是一般人員職稱；火警值班則是職務池文字。
    if(Array.isArray(state.specialDutyPersonnel)){
      state.specialDutyPersonnel.forEach(item=>{
        if(!item || !item.name || !item.duty) return;
        const duty=String(item.duty).trim();
        const name=String(item.name).trim();
        if(duty==='火警值班'){
          if(!dutyMaster.includes('火警值班')) dutyMaster.push('火警值班');
          return;
        }
        if(!manualPersonnel.some(x=>x.name===name)){
          manualPersonnel.push({id:item.id || `manual-${Date.now()}-${Math.random()}`,no:'',name,role:duty});
        }
        if(!roleMaster.includes(duty)) roleMaster.push(duty);
      });
    }

    if(Array.isArray(state.vehicleMaster) && state.vehicleMaster.length){
      vehicleMaster = state.vehicleMaster
        .filter(item=>item && (typeof item === 'string' || item.name))
        .map(item=>typeof item === 'string' ? {name:item,type:'車輛'} : {name:String(item.name),type:String(item.type || '車輛')});
      renderVehiclePool();
    }

    currentDutyKey = String(state.currentDutyKey || '');

    const targets = getBoardTargets();
    const savedSlots = state.board && Array.isArray(state.board.slots)
      ? state.board.slots
      : [];

    savedSlots.forEach(slot=>{
      const target = targets[Number(slot.index)];
      if(!target || !slot.item) return;

      // 舊資料若曾把車輛存進人員格，這裡直接忽略，避免錯誤資料再次出現。
      if(!isSavedItemAllowedInTarget(slot.item,target)){
        console.warn('略過型別不相符的舊看板資料',slot);
        return;
      }

      const $item = createRestoredItem(slot.item);
      if($item) $(target).append($item);
    });

    const rest = state.board && Array.isArray(state.board.rest) ? state.board.rest : [];
    rest.forEach(item=>{
      const $item = createRestoredStatusItem(item);
      if($item) $('#restingBody').append($item);
    });

    const official = state.board && Array.isArray(state.board.official) ? state.board.official : [];
    official.forEach(item=>{
      const $item = createRestoredStatusItem(item);
      if($item) $('#officialBody').append($item);
    });

    syncAll();

    // v37：boardState 還原後，不論目前時段 key 是否與已儲存狀態相同，
    // 都要先更新勤務測試編輯器，避免 HTML 初始 disabled 狀態一直保留。
    updateDutyPeriodEditorUi();

    // 正式模式依現在時間；若本機正在測試，則維持測試指定時段。
    const period = getActiveDutyPeriod();
    if(period){
      $('#currentDutyPeriod').text(`${period.start}–${period.end}`);
      // v44：即使儲存的 currentDutyKey 與目前時段相同，
      // 還原後也要補一次火警值班右側連動，避免舊版狀態要等下一次換時段才生效。
      applyFireDutyWatch(period,{onlyEmpty:true});
      refreshDutySchedule(false);
    }else{
      $('#currentDutyPeriod').text(dutySchedule.length ? '無符合時段' : '尚未匯入');
      updateDutyPeriodEditorUi();
    }
  }

  /* =========================================================
     06-1. 拖曳資料型別與目標格驗證
     ---------------------------------------------------------
     規則只有兩種：
     - vehicle 只能進 data-kind="vehicle" 的車輛格。
     - person 只能進 data-kind="person" 或人員狀態區。

     這一層是保護機制，即使未來 HTML 欄位增加，也不會因為
     Sortable group 設定疏漏，把 11車之類的車輛丟到破壞小組。
  ========================================================= */
  function dragItemType(item){
    return String($(item).attr('data-drag-type') || '').trim();
  }

  function targetKind(target){
    return String($(target).attr('data-kind') || '').trim();
  }

  function isItemAllowedInDropTarget(item,target){
    const itemType = dragItemType(item);
    const kind = targetKind(target);
    return !!itemType && !!kind && itemType === kind;
  }

  function isSavedItemAllowedInTarget(item,target){
    if(!item || !item.type) return false;
    return String(item.type) === targetKind(target);
  }

  /* =========================================================
     07. 讀取最新看板與開啟頁面同步
     ---------------------------------------------------------
     - API 已設定：從 Google Sheet 讀。
     - API 尚未設定：用 localStorage 測試。
     - 輪詢時只有伺服器內容真的不同才重畫畫面。
  ========================================================= */
  async function loadPublishedState(silent=false){
    if(isDragging || isSaving || autoSaveTimer) return;

    // 測試模式只在目前瀏覽器預覽。背景輪詢不得把測試畫面蓋回正式資料。
    if(silent && dutyPeriodOverrideKey) return;

    // true：固定使用本機資料，不會呼叫 GAS。
    if(useLocalData()){
      try{
        const raw = localStorage.getItem(localStorageKey());
        const state = raw ? JSON.parse(raw) : null;
        lastServerStateJson = state ? JSON.stringify(state) : '';
        restoreBoardState(state);
        setSyncStatus('本地資料','text-warning');
      }catch(error){
        console.error(error);
        restoreBoardState(null);
        setSyncStatus('本地資料讀取失敗','text-danger');
      }
      return;
    }

    // false：固定使用線上資料。URL 沒設定時直接報錯，不偷偷退回本機。
    if(!isGasApiConfigured()){
      restoreBoardState(null);
      setSyncStatus('線上模式未設定 GAS URL','text-danger');
      if(!silent) toast('目前為線上資料模式，請先設定 GAS Web App URL');
      return;
    }

    try{
      if(!silent) setSyncStatus('讀取中…','text-secondary');
      const result = await apiGetBoardState();

      if(!result || result.ok !== true){
        throw new Error(result?.message || 'API 回傳格式錯誤');
      }

      const state = result.state || null;
      const remoteJson = state ? JSON.stringify(state) : '';

      if(!silent || remoteJson !== lastServerStateJson){
        lastServerStateJson = remoteJson;
        restoreBoardState(state);
      }

      setSyncStatus(result.updatedAt ? `已同步 ${result.updatedAt}` : '已同步','text-success');
    }catch(error){
      console.error(error);
      setSyncStatus('Google Sheet 連線失敗','text-danger');
      if(!silent) toast('無法讀取 Google Sheet 最新看板');
    }
  }

  /* =========================================================
     08. 自動儲存
     ---------------------------------------------------------
     匯入成功或拖曳完成後就自動寫入 Google Sheet。
     多個連續 Sortable 事件先 debounce，再只寫一次，避免浪費額度。
  ========================================================= */
  function queueAutoSave(reason='change',delay=250){
    // 測試時段只是本機預覽，任何拖曳都不能污染正式 boardState。
    if(dutyPeriodOverrideKey){
      clearTimeout(autoSaveTimer);
      autoSaveTimer = null;
      setSyncStatus('測試模式：不自動同步','text-warning');
      return;
    }

    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(()=>{
      autoSaveTimer = null;
      saveBoardStateNow(reason);
    },delay);
  }

  async function saveBoardStateNow(reason='change'){
    if(isSaving){
      saveQueued = true;
      return;
    }

    const state = collectBoardState();
    const stateJson = JSON.stringify(state);
    isSaving = true;
    setSyncStatus('同步中…','text-secondary');

    try{
      // true：只寫入目前瀏覽器 localStorage。
      if(useLocalData()){
        localStorage.setItem(localStorageKey(),stateJson);
        lastServerStateJson = stateJson;
        setSyncStatus('本地已自動儲存','text-warning');
      }else{
        // false：只寫入 GAS / Google Sheet；URL 未設定就視為錯誤。
        if(!isGasApiConfigured()){
          throw new Error('線上資料模式尚未設定 GAS Web App URL');
        }

        const result = await apiSaveBoardState(state);
        if(!result || result.ok !== true){
          throw new Error(result?.message || 'API 儲存失敗');
        }

        lastServerStateJson = stateJson;
        setSyncStatus(result.updatedAt ? `已同步 ${result.updatedAt}` : '已同步','text-success');
      }

      if(reason === 'import'){
        toast('勤務表已匯入並自動同步');
      }
    }catch(error){
      console.error(error);
      setSyncStatus('自動同步失敗','text-danger');
      toast('自動同步 Google Sheet 失敗，請確認 GAS Web App URL 與部署權限');
    }finally{
      isSaving = false;

      if(saveQueued){
        saveQueued = false;
        queueAutoSave('queued',100);
      }
    }
  }

  /* =========================================================
     09. 啟動遠端更新輪詢
     ---------------------------------------------------------
     已開啟網頁的人每隔一段時間讀一次最新 boardState。
     間隔可在 board-data.json 的 pollIntervalMs 調整。
  ========================================================= */
  function startRemotePolling(){
    clearInterval(pollTimer);

    // 本地模式不需要輪詢；只有線上模式才向 Google Sheet 讀最新狀態。
    if(useLocalData() || !isGasApiConfigured()) return;

    pollTimer = setInterval(()=>{
      loadPublishedState(true);
    },pollIntervalMs());
  }

  /* =========================================================
     10. 載入前端設定並開始系統
     ---------------------------------------------------------
     board-data.json 只放固定設定，不保存當日姓名或勤務資料。
  ========================================================= */
  function loadBoardData(){
    $.getJSON('data/board-data.json')
      .done(function(data){
        appConfig = data || {};
        todayRoster = [];
        vehicleMaster = Array.isArray(appConfig.vehicles) ? appConfig.vehicles : [];
        roleMaster = Array.isArray(appConfig.roles) ? appConfig.roles.map(item=>String(item || '').trim()).filter(Boolean) : ['隊員','小隊長','役男','分隊長','中隊長','義消'];
        dutyMaster = Array.isArray(appConfig.dutyPool) ? appConfig.dutyPool.map(item=>String(item || '').trim()).filter(Boolean) : ['火警值班'];
        importedBaseAssignments = Array.isArray(appConfig.baseAssignments) ? appConfig.baseAssignments.map(item=>({...item})) : [];
        manualPersonnel = [];
        dutySchedule = [];
        dutyStatusByNo = new Map();

        if(appConfig.app && appConfig.app.title){
          $('#topbarTitle').text(appConfig.app.title);
          document.title = appConfig.app.title;
        }

        setupAutoFullscreen();
        renderVehiclePool();
        rebuildSpecialDutyPool();
        loadPublishedState(false);
        startRemotePolling();
      })
      .fail(function(){
        toast('無法讀取 data/board-data.json，請確認 GitHub Pages 路徑。');
      });
  }

  /* =========================================================
     11. 現在時間與勤務時段定時檢查
  ========================================================= */
  function updateClock(){
    const now = new Date();
    $('#clockTime').text(now.toLocaleTimeString('zh-TW',{hour12:false}));
    $('#clockDate').text(now.toLocaleDateString('zh-TW',{
      year:'numeric',month:'2-digit',day:'2-digit',weekday:'short'
    }));
  }
  updateClock();
  setInterval(updateClock,1000);

  // 每 10 秒確認是否跨到下一個勤務時段。
  // 只有時段真正改變時才重新套用，不會每秒重畫看板。
  setInterval(function(){
    refreshDutySchedule(false);
  },10000);

  /* =========================================================
     12. 看板配置狀態與勤務時間共用工具
  ========================================================= */
  function setEmpty($target){
    if($target.find('[data-drag-type]').length) return;
    $target.html('');
  }

  function assignedSet(kind){
    const set = new Set();
    $(`.drop-target[data-kind="${kind}"] [data-drag-type="${kind}"]`).each(function(){
      const value = $(this).data('value') || $(this).text().trim();
      if(value) set.add(String(value));
    });
    return set;
  }

  // =========================================================
  // v47：全看板唯一性
  // ---------------------------------------------------------
  // 同一個人員、職務、車輛在網頁上只能存在一份。
  // 也就是：要嘛留在池中，要嘛已配置在看板；不允許兩邊或多格重複。
  // =========================================================
  function boardEntityKey($item){
    if(!$item || !$item.length) return '';
    const type=String($item.attr('data-drag-type') || '').trim();
    const value=String($item.attr('data-value') || $item.text().trim()).trim();
    if(!type || !value) return '';

    if(type==='vehicle') return `vehicle:${value}`;

    const source=String($item.attr('data-person-source') || '').trim();
    const dutyRole=String($item.attr('data-duty-role') || '').trim();
    if(source==='duty' || dutyRole) return `duty:${dutyRole || value}`;

    const no=String($item.attr('data-no') || '').trim();
    return no ? `person-no:${no}` : `person-name:${value}`;
  }

  function boardEntityPriority($item){
    const $parent=$item.parent();
    const type=String($item.attr('data-drag-type') || '').trim();
    const value=String($item.attr('data-value') || $item.text().trim()).trim();

    // 91 / 92 車固定以「專責救護」為唯一位置。
    if(type==='vehicle' && (value==='91車' || value==='92車') && String($parent.attr('data-duty-vehicle') || '')===value) return 1000;
    // 勤務表自動休息優先保留。
    if($parent.is('#restingBody') || $parent.closest('#restingBody').length) return 900;
    // 91 / 92 當前救護人員優先於一般火警基礎配置。
    if($parent.attr('data-duty-source')==='備勤91' || $parent.attr('data-duty-source')==='備勤救災') return 800;
    // 火警值班右側由值班欄自動帶入，優先於一般基礎配置。
    if(String($item.attr('data-auto-source') || '')==='fire-duty-watch') return 700;
    return 100;
  }

  function enforceUniqueBoardAssignments(){
    const items=$('.drop-target [data-drag-type], #restingBody [data-drag-type], #officialBody [data-drag-type]').toArray();
    const ordered=items.map((el,index)=>({el,index,$item:$(el),priority:boardEntityPriority($(el))}))
      .sort((a,b)=>b.priority-a.priority || a.index-b.index);
    const kept=new Set();

    ordered.forEach(entry=>{
      const key=boardEntityKey(entry.$item);
      if(!key) return;
      if(kept.has(key)){
        entry.$item.remove();
        return;
      }
      kept.add(key);
    });
  }

  function statusSet(){
    const set = new Set();
    $('#restingBody [data-drag-type="person"], #officialBody [data-drag-type="person"]').each(function(){
      const value = $(this).data('value') || $(this).text().trim();
      if(value) set.add(String(value));
    });
    return set;
  }

  function minutesOf(text){
    const m = String(text || '').match(/(\d{1,2})[:：](\d{2})/);
    if(!m) return null;
    return Number(m[1]) * 60 + Number(m[2]);
  }

  function isActiveShift(item){
    const start = minutesOf(item.start);
    const end = minutesOf(item.end);
    if(start === null || end === null) return true;

    const now = new Date();
    const current = now.getHours()*60 + now.getMinutes();

    if(end >= start) return current >= start && current <= end;
    return current >= start || current <= end;
  }

  /* =========================================================
     13. 人員身分標準化與顏色分類
  ========================================================= */
  function normalizeRole(role){
    const raw = String(role || '').trim();

    // JSON 與系統正式資料一律使用中文。
    // 英文只保留為舊資料相容，不會再寫回英文值。
    const compatibilityMap = {
      'member':'隊員',
      'squadleader':'小隊長',
      'squad_leader':'小隊長',
      'alternative':'役男',
      '替代役':'役男',
      'captain':'分隊長'
    };

    const supportedRoles = roleMaster.length ? roleMaster : ['隊員','小隊長','役男','分隊長'];

    if(supportedRoles.includes(raw)){
      return raw;
    }

    const compatible = compatibilityMap[raw.toLowerCase()];
    if(compatible) return compatible;
    return raw || '隊員';
  }

  function roleClass(role){
    const value = normalizeRole(role);

    return {
      '隊員':'role-member',
      '小隊長':'role-squad-leader',
      '役男':'role-alternative',
      '分隊長':'role-captain',
      '中隊長':'role-company-leader'
    }[value] || 'role-member';
  }

  function roleLabel(role){
    return normalizeRole(role);
  }

  function findRosterRole(name){
    const item = todayRoster.find(x => x.name === name) || manualPersonnel.find(x=>x.name === name);
    return item ? normalizeRole(item.role) : '隊員';
  }

  function applyPersonRole($element, name, role){
    const normalized = normalizeRole(role || findRosterRole(name));
    $element
      .removeClass('role-member role-squad-leader role-alternative role-captain role-company-leader')
      .addClass(roleClass(normalized))
      .attr('data-role', normalized);
    return $element;
  }

  function findPersonByNo(no){
    const normalized = String(no ?? '').trim();
    if(!normalized) return null;

    return todayRoster.find(item => String(item.no ?? '').trim() === normalized) || null;
  }

  function currentMinutes(){
    const now = new Date();
    return now.getHours() * 60 + now.getMinutes();
  }

  function isMinuteInPeriod(current,start,end){
    if(start === null || end === null) return false;

    // 一般時段，例如 08:00-09:00
    if(end > start){
      return current >= start && current < end;
    }

    // 跨午夜，例如 23:00-00:00
    if(end < start){
      return current >= start || current < end;
    }

    // start == end 視為整天，不建議一般勤務這樣設定
    return true;
  }

  function dutyPeriodKey(period){
    if(!period) return '';
    return `${period.start || ''}|${period.end || ''}`;
  }

  function getCurrentDutyPeriod(){
    const current = currentMinutes();

    return dutySchedule.find(item=>{
      return isMinuteInPeriod(
        current,
        minutesOf(item.start),
        minutesOf(item.end)
      );
    }) || null;
  }

  function getDutyPeriodByKey(key){
    const normalized = String(key || '').trim();
    if(!normalized) return null;
    return dutySchedule.find(item=>dutyPeriodKey(item) === normalized) || null;
  }

  // 正式模式使用現在時間；測試模式使用側欄手動指定的勤務時段。
  function getActiveDutyPeriod(){
    if(dutyPeriodOverrideKey){
      return getDutyPeriodByKey(dutyPeriodOverrideKey);
    }
    return getCurrentDutyPeriod();
  }

  function updateDutyPeriodEditorUi(){
    const $editBtn = $('#dutyPeriodEditBtn');
    const $select = $('#dutyPeriodSelect');
    const $badge = $('#dutyPeriodTestBadge');

    $editBtn.prop('disabled',!dutySchedule.length);
    $badge.toggleClass('d-none',!dutyPeriodOverrideKey);

    if(!$select.length) return;
    const selectedKey = dutyPeriodOverrideKey || dutyPeriodKey(getCurrentDutyPeriod());
    $select.empty();

    dutySchedule.forEach(period=>{
      const key = dutyPeriodKey(period);
      const label = `${period.start || '--:--'}–${period.end || '--:--'}`;
      $select.append($('<option></option>').val(key).text(label));
    });

    if(selectedKey && $select.find(`option[value="${selectedKey}"]`).length){
      $select.val(selectedKey);
    }
  }

  function createPersonChip(person,cssClass){
    const $chip = $(`<div class="${cssClass}"></div>`)
      .attr('data-drag-type','person')
      .attr('data-value',person.name)
      .attr('data-role',normalizeRole(person.role))
      .attr('data-no',person.no ?? '');

    if(cssClass === 'status-chip'){
      $chip.append($('<span></span>').text(person.name));
    }else{
      $chip.text(person.name);
    }

    applyPersonRole($chip,person.name,person.role);
    return $chip;
  }

  function statusesForNo(no){
    const key = String(no ?? '').trim();
    return key ? [...(dutyStatusByNo.get(key) || [])] : [];
  }

  function blockingStatusSet(){
    const settings = (appConfig && appConfig.importSettings) || {};
    return new Set((settings.blockingStatusKeywords || [
      '輪休','請休','補休','公假','連續補休','休假役男'
    ]).map(item=>cleanText(item)));
  }

  function hasBlockingDutyStatus(no){
    const blocking = blockingStatusSet();
    return statusesForNo(no).some(status=>blocking.has(cleanText(status)));
  }

  function currentDutyNumberSet(){
    const period = getActiveDutyPeriod();
    return new Set((period?.allDutyNumbers || []).map(no=>String(no)));
  }

  function currentRestNumberSet(){
    const period = getActiveDutyPeriod();
    return new Set((period?.['休息時間'] || []).map(no=>String(no)));
  }

  function assignedPersonNoSet(){
    const result = new Set();
    // 只計算「已經放在看板上的人員」，不能把人員池也算進去。
    // v47 若掃描整頁，人員池本身就會讓所有 Excel 基礎人員被誤判為已配置，
    // 因而造成匯入後整張基礎火警編制都沒有帶進來。
    $('.drop-target [data-drag-type="person"][data-no], #restingBody [data-drag-type="person"][data-no], #officialBody [data-drag-type="person"][data-no]').each(function(){
      const no = String($(this).attr('data-no') || '').trim();
      if(no) result.add(no);
    });
    return result;
  }

  function resolveBasePerson(assignment){
    const no = String(assignment?.no ?? '').trim();
    const found = no ? findPersonByNo(no) : null;
    if(found) return found;

    const name = String(assignment?.name || '').trim();
    if(!name) return null;

    return {
      no:assignment?.no ?? '',
      name,
      role:normalizeRole(assignment?.role)
    };
  }

  function basePersonAllowedNow(person,ignoreDuty=false){
    if(!person) return false;
    if(ignoreDuty || !hasDetailedDutyData) return true;

    const no = String(person.no ?? '').trim();
    if(!no) return false;

    // v50：火警出動人員只從「目前時段在隊備勤」的人員中產生。
    // 91 / 92、休息、請假或其他非在隊備勤勤務都不能被基礎配置帶回火警表。
    if(!currentAtStationNumberSet().has(no)) return false;
    if(hasBlockingDutyStatus(no)) return false;
    if(currentRestNumberSet().has(no)) return false;
    return true;
  }

  function clearAutoBaseAssignments(){
    // 新版自動配置有 data-auto-source，可直接清除。
    $('[data-auto-source="base"]').remove();

    // 相容 v35 以前已儲存的 boardState：舊資料沒有 autoSource。
    // 若「基礎指定 slot」裡仍放著與 baseAssignments 完全相同的人/車，
    // 視為舊版自動基礎配置並清除後重算；不同內容代表人工調整，保留不動。
    const assignments = importedBaseAssignments.length ? importedBaseAssignments : (Array.isArray(appConfig?.baseAssignments) ? appConfig.baseAssignments : []);
    assignments.forEach(assignment=>{
      const slotId = String(assignment?.slotId || '').trim();
      if(!slotId) return;

      const $target = $(`.drop-target[data-slot-id="${slotId}"]`).first();
      const $item = $target.children('[data-drag-type]').first();
      if(!$target.length || !$item.length) return;

      if(assignment.type === 'vehicle'){
        const expected = String(assignment.value || '').trim();
        const actual = String($item.attr('data-value') || $item.text().trim()).trim();
        if($item.attr('data-drag-type') === 'vehicle' && expected && actual === expected){
          $item.remove();
        }
        return;
      }

      if(assignment.type === 'person'){
        const expectedNo = String(assignment.no ?? '').trim();
        const expectedName = String(assignment.name || '').trim();
        const actualNo = String($item.attr('data-no') || '').trim();
        const actualName = String($item.attr('data-value') || $item.text().trim()).trim();
        const samePerson = (expectedNo && actualNo === expectedNo) || (!expectedNo && expectedName && actualName === expectedName);
        if($item.attr('data-drag-type') === 'person' && samePerson){
          $item.remove();
        }
      }
    });
  }

  function applyBaseAssignments(options={}){
    const assignments = importedBaseAssignments.length ? importedBaseAssignments : (Array.isArray(appConfig?.baseAssignments) ? appConfig.baseAssignments : []);
    const onlyEmpty = options.onlyEmpty !== false;
    const ignoreDuty = options.ignoreDuty === true;
    const assignedNos = assignedPersonNoSet();
    const assignedNames = new Set();
    $('.drop-target [data-drag-type="person"]').not('[data-person-source="duty"]')
      .add('#restingBody [data-drag-type="person"]').not('[data-person-source="duty"]')
      .add('#officialBody [data-drag-type="person"]').not('[data-person-source="duty"]')
      .each(function(){
        const name=String($(this).attr('data-value') || $(this).text().trim()).trim();
        if(name) assignedNames.add(name);
      });
    const assignedVehicles = assignedSet('vehicle');
    const assignedDuties = new Set();
    // 職務池中的「火警值班」只是尚未配置的來源，不代表它已經在看板上。
    // 因此唯一性判斷只能掃描看板/狀態區，不能掃描職務池。
    $('.drop-target [data-person-source="duty"], #restingBody [data-person-source="duty"], #officialBody [data-person-source="duty"]').each(function(){
      const duty=String($(this).attr('data-duty-role') || $(this).attr('data-value') || '').trim();
      if(duty) assignedDuties.add(duty);
    });

    assignments.forEach(assignment=>{
      const slotId = String(assignment?.slotId || '').trim();
      if(!slotId) return;

      const $target = $(`.drop-target[data-slot-id="${slotId}"]`).first();
      if(!$target.length) return;
      if(onlyEmpty && $target.children('[data-drag-type]').length) return;

      if(assignment.type === 'vehicle'){
        const value = String(assignment.value || '').trim();
        if(!value || targetKind($target[0]) !== 'vehicle') return;
        if(assignedVehicles.has(value)) return;
        $target.empty().append(
          $('<div class="vehicle-chip"></div>')
            .attr('data-drag-type','vehicle')
            .attr('data-value',value)
            .attr('data-auto-source','base')
            .text(value)
        );
        assignedVehicles.add(value);
        return;
      }

      if(assignment.type === 'duty'){
        const value = String(assignment.value || '').trim();
        if(!value || targetKind($target[0]) !== 'person') return;
        if(assignedDuties.has(value)) return;
        $target.empty().append(
          $('<div class="person-chip duty-assignment-chip"></div>')
            .attr('data-drag-type','person')
            .attr('data-value',value)
            .attr('data-no','')
            .attr('data-person-source','duty')
            .attr('data-duty-role',value)
            .attr('data-auto-source','base')
            .text(value)
        );
        assignedDuties.add(value);
        return;
      }

      if(assignment.type !== 'person' || targetKind($target[0]) !== 'person') return;

      const person = resolveBasePerson(assignment);

      // v53：帶隊官屬於 Excel 的「基礎火警編制」，不是目前時段的在隊備勤隨機人員。
      // 因此 leader 格不可套用「必須在在隊備勤」的過濾，否則 Excel 明明有帶隊官，
      // 網頁卻會因該番號不在在隊備勤欄而直接略過。
      // 其他基礎人員仍維持目前勤務限制，後續駕駛／瞄子／副瞄子再由隨機編組接手。
      const isBaseLeader = /-leader$/.test(slotId);
      if(!isBaseLeader && !basePersonAllowedNow(person,ignoreDuty)) return;
      if(isBaseLeader && !person) return;

      const no = String(person.no ?? '').trim();
      if(no && assignedNos.has(no)) return;
      if(!no && assignedNames.has(String(person.name || '').trim())) return;

      ensureRosterPerson(person.name,person.role,person.no);
      const $chip = createPersonChip(person,'person-chip').attr('data-auto-source','base');
      $target.empty().append($chip);
      if(no) assignedNos.add(no);
      else if(person.name) assignedNames.add(String(person.name).trim());
    });
  }

  function applyDutyToFixedSlots(source,numbers){
    const list = Array.isArray(numbers) ? numbers : [];
    const $targets = $(`.drop-target[data-duty-source="${source}"]`).sort(function(a,b){
      return Number($(a).data('duty-index')) - Number($(b).data('duty-index'));
    });

    $targets.each(function(index){
      const $target = $(this);
      $target.empty();

      const no = list[index];
      if(no === undefined || no === null || String(no).trim() === ''){
        return;
      }

      const person = findPersonByNo(no);
      if(!person){
        console.warn(`勤務番號 ${no} 找不到對應人員`);
        return;
      }

      // 91 / 92 專責救護屬於當前勤務，優先於基礎火警配置。
      removePersonFromBoardByNo(no,$target[0]);
      $target.append(createPersonChip(person,'person-chip'));
    });
  }

  function removePersonFromBoardByNo(no,exceptTarget){
    const normalized = String(no ?? '').trim();
    if(!normalized) return;

    $(`[data-drag-type="person"][data-no="${normalized}"]`).each(function(){
      if(exceptTarget && $.contains(exceptTarget,this)) return;
      $(this).remove();
    });
  }

  function removeBlockingStatusPeopleFromBoard(){
    dutyStatusByNo.forEach((statuses,no)=>{
      if(!hasBlockingDutyStatus(no)) return;
      removePersonFromBoardByNo(no,null);
    });
  }

  function applyDutyToStatus(source,numbers){
    const list = Array.isArray(numbers) ? numbers : [];
    const $target = $(`.panel-body-drop[data-duty-source="${source}"]`);

    if(!$target.length) return;
    $target.empty();

    list.forEach(no=>{
      const person = findPersonByNo(no);
      if(!person){
        console.warn(`勤務番號 ${no} 找不到對應人員`);
        return;
      }

      // 勤務表明確標示「休息時間」時，以勤務表為最高優先：
      // 基礎配置、91/92 或人工配置只要是同一人，都先移除。
      removePersonFromBoardByNo(no,$target[0]);
      $target.append(createPersonChip(person,'status-chip'));
    });
  }

  /* =========================================================
     14. 基礎配置 + 當前勤務優先順序
     ---------------------------------------------------------
     1. 基礎火警配置只補空格，不覆蓋人工調整。
     2. 91 / 92 當前專責救護優先於基礎配置。
     3. 輪休 / 公假 / 請休等不可排狀態會從看板移除。
     4. 「休息時間」優先級最高，該人原本的基礎格保持空白。
  ========================================================= */
  function ensureDutyVehicle(vehicleName){
    const $target = $(`.drop-target[data-duty-vehicle="${vehicleName}"]`).first();
    if(!$target.length) return;

    $target.empty().append(
      $('<div class="vehicle-chip"></div>')
        .attr('data-drag-type','vehicle')
        .attr('data-value',vehicleName)
        .text(vehicleName)
    );
  }

  // =========================================================
  // v45 火警值班連動
  // ---------------------------------------------------------
  // 「火警值班」本身是職務池卡片；當它被放進火警配置表後，
  // 其右側相鄰的人員格會依目前勤務表「值班」欄自動帶入當班番號。
  // 若名冊可找到該番號，就顯示該人姓名；找不到時至少顯示「XX號」。
  // 自動帶入的人員以 data-auto-source="fire-duty-watch" 標記，
  // 只會清除自己產生的內容，不會覆蓋值班台人工配置。
  // =========================================================
  function currentWatchNumbers(period){
    return [...(period?.['值班'] || [])].map(no=>String(no ?? '').trim()).filter(Boolean);
  }

  function fireDutyRoleTargets(){
    return $('.fire-table .drop-target').filter(function(){
      const $chip=$(this).children('[data-person-source="duty"][data-duty-role="火警值班"]').first();
      return $chip.length>0;
    });
  }

  function clearAutoFireDutyWatch(){
    $('.fire-table [data-auto-source="fire-duty-watch"]').remove();
  }

  function applyFireDutyWatch(period,{onlyEmpty=true}={}){
    clearAutoFireDutyWatch();
    const watchNos=currentWatchNumbers(period);
    if(!watchNos.length) return;

    fireDutyRoleTargets().each(function(index){
      const $roleCell=$(this);
      const $right=$roleCell.next('td.drop-target[data-kind="person"]');
      if(!$right.length) return;
      if(onlyEmpty && $right.children('[data-drag-type]').length) return;

      // 若勤務表值班欄有多個番號，依火警值班卡片由上到下依序配對。
      const no=watchNos[Math.min(index,watchNos.length-1)];
      const person=findPersonByNo(no) || manualPersonnel.find(item=>String(item.no ?? '').trim()===no) || null;
      const display=person?.name || `${no}號`;
      const role=person?.role || '隊員';

      // 同一番號不得重複。若此人目前已在 91 / 92 或休息，這些勤務優先，
      // 火警值班右側保持空白；若只是在一般火警格，則移到值班右側。
      const $samePerson=$(`[data-drag-type="person"][data-no="${no}"]`).filter(function(){
        return !$.contains($right[0],this);
      });
      const hasHigherPriority=$samePerson.toArray().some(el=>{
        const $parent=$(el).parent();
        return $parent.is('#restingBody')
          || $parent.closest('#restingBody').length>0
          || $parent.attr('data-duty-source')==='備勤91'
          || $parent.attr('data-duty-source')==='備勤救災';
      });
      if(hasHigherPriority) return;
      removePersonFromBoardByNo(no,$right[0]);

      const $chip=$('<div class="person-chip"></div>')
        .attr('data-drag-type','person')
        .attr('data-value',display)
        .attr('data-no',no)
        .attr('data-person-source',person ? (todayRoster.includes(person)?'daily':'manual') : 'daily')
        .attr('data-auto-source','fire-duty-watch')
        .text(display);
      applyPersonRole($chip,display,role);
      $right.empty().append($chip);
    });
  }

  // =========================================================
  // v52 在隊備勤 → 火警出動隨機編組
  // ---------------------------------------------------------
  // 只有目前時段「在隊備勤」的人才會自動補入火警主表。
  // 排除：91 / 92、休息、請假等不可排狀態、義消，以及目前值班人員。
  // 配置優先順序：先依車序填駕駛，再依車序填瞄子手，再依車序填副瞄子手。
  // 駕駛車序：11車 → 31車 → 61車 → 中隊指揮車；瞄子手／副瞄子手只有 11、31、61 車。
  // 帶隊官完全由 Excel 基礎編制帶入，不參與隨機；中隊長僅能由人員池手動拉入。
  // =========================================================
  function shuffleCopy(items){
    const result=[...(items || [])];
    for(let i=result.length-1;i>0;i--){
      const j=Math.floor(Math.random()*(i+1));
      [result[i],result[j]]=[result[j],result[i]];
    }
    return result;
  }

  function clearAutoFireRandomAssignments(){
    $('[data-auto-source="fire-random"]').remove();
  }

  function fireRandomTargetSlotIds(){
    // v52：中隊指揮車只有駕駛，不配置瞄子手／副瞄子手。
    // 11駕駛 → 31駕駛 → 61駕駛 → 中指車駕駛
    // → 11瞄子 → 31瞄子 → 61瞄子
    // → 11副瞄 → 31副瞄 → 61副瞄。
    return [
      'first-1-driver','first-3-driver','first-2-driver','second-3-driver',
      'first-1-nozzle','first-3-nozzle','first-2-nozzle',
      'first-1-assistant-nozzle','first-3-assistant-nozzle','first-2-assistant-nozzle'
    ];
  }

  function ensureMainFireVehicle(slotId,vehicleName){
    const $target=$(`.drop-target[data-slot-id="${slotId}"]`).first();
    if(!$target.length || targetKind($target[0])!=='vehicle') return;
    const $current=$target.children('[data-drag-type="vehicle"]').first();
    if($current.length){
      const current=String($current.attr('data-value') || $current.text().trim()).trim();
      if(current===vehicleName) return;
      // 人工放置的其他車輛不覆蓋。
      if(!$current.attr('data-auto-source')) return;
      $current.remove();
    }
    // 同一車輛只能存在一份；若別處已有人工配置則不再複製。
    const $duplicate=$(`[data-drag-type="vehicle"][data-value="${vehicleName}"]`).filter(function(){
      return !$.contains($target[0],this);
    }).first();
    if($duplicate.length) return;
    $target.append(
      $('<div class="vehicle-chip"></div>')
        .attr('data-drag-type','vehicle')
        .attr('data-value',vehicleName)
        .attr('data-auto-source','fire-random')
        .text(vehicleName)
    );
  }

  function eligibleAtStationPeople(period){
    const atStation=new Set((period?.['在隊備勤'] || []).map(no=>String(no ?? '').trim()).filter(Boolean));
    const rescue=new Set([
      ...(period?.['備勤91'] || []),
      ...(period?.['備勤救災'] || [])
    ].map(no=>String(no ?? '').trim()).filter(Boolean));
    const resting=new Set((period?.['休息時間'] || []).map(no=>String(no ?? '').trim()).filter(Boolean));
    const watch=new Set(currentWatchNumbers(period));

    return todayRoster.filter(person=>{
      const no=String(person?.no ?? '').trim();
      if(!no || !atStation.has(no)) return false;
      if(rescue.has(no) || resting.has(no) || watch.has(no)) return false;
      if(hasBlockingDutyStatus(no)) return false;
      if(normalizeRole(person?.role)==='義消') return false;
      return true;
    });
  }

  function applyAtStationFireRandom(period){
    if(!period || !hasDetailedDutyData) return;

    // 車輛本身由 Excel 基礎編制帶入；這裡只確保既有標準車名不因舊狀態遺失。
    ensureMainFireVehicle('first-1-vehicle','11車');
    ensureMainFireVehicle('first-3-vehicle','31車');
    ensureMainFireVehicle('first-2-vehicle','61車');
    ensureMainFireVehicle('second-3-vehicle','中隊指揮車');

    // 共 10 格：4 個駕駛 + 3 個瞄子手 + 3 個副瞄子手，由目前在隊備勤人員隨機編組。
    // 只清掉系統從 Excel 自動帶入的基礎人員；人工拖曳內容保留。
    const targetIds=fireRandomTargetSlotIds();
    targetIds.forEach(slotId=>{
      const $target=$(`.drop-target[data-slot-id="${slotId}"]`).first();
      $target.children('[data-auto-source="base"][data-drag-type="person"]').remove();
    });

    // 帶隊官不在這裡處理：一律保留 Excel 基礎編制。
    // 中隊長也不讀勤務表；如需配置，由人員池手動拖入。
    const assigned=assignedPersonNoSet();
    const candidates=shuffleCopy(eligibleAtStationPeople(period).filter(person=>{
      const no=String(person.no ?? '').trim();
      if(!no) return false;
      return !assigned.has(no);
    }));

    let candidateIndex=0;
    targetIds.forEach(slotId=>{
      const $target=$(`.drop-target[data-slot-id="${slotId}"]`).first();
      if(!$target.length || $target.children('[data-drag-type]').length) return;

      const person=candidates[candidateIndex++];
      if(!person) return;
      const no=String(person.no ?? '').trim();
      removePersonFromBoardByNo(no,$target[0]);
      $target.append(createPersonChip(person,'person-chip').attr('data-auto-source','fire-random'));
    });
  }

  function applyDutyPeriod(period){
    if(!period) return;

    // 清掉上一時段的『自動休息』與『自動基礎配置』，再依目前/測試時段重算。
    // 人工拖曳的配置不會被 clearAutoBaseAssignments() 清除。
    $('#restingBody').empty();
    clearAutoBaseAssignments();
    clearAutoFireRandomAssignments();

    // Excel 右上角仍提供當日基礎車輛、帶隊官、火警值班等設定；
    // 但人員只有目前時段「在隊備勤」才允許自動帶入。
    applyBaseAssignments({onlyEmpty:true,ignoreDuty:false});

    ensureDutyVehicle('91車');
    ensureDutyVehicle('92車');
    applyDutyToFixedSlots('備勤91',period['備勤91']);
    applyDutyToFixedSlots('備勤救災',period['備勤救災']);

    // 當日請假/輪休等狀態不可出現在出動看板。
    removeBlockingStatusPeopleFromBoard();

    // 休息時間最後處理，優先級最高。
    applyDutyToStatus('休息時間',period['休息時間']);

    // 火警值班卡片若已放在表內，依本時段「值班」番號同步右側人員。
    applyFireDutyWatch(period,{onlyEmpty:true});

    // 其餘「在隊備勤」人員依 11駕駛→31駕駛→61駕駛→中指車駕駛→11/31/61瞄子手→11/31/61副瞄子手隨機補入。
    applyAtStationFireRandom(period);

    $('#currentDutyPeriod').text(`${period.start || '--:--'}–${period.end || '--:--'}`);
    updateDutyPeriodEditorUi();
    syncAll();
  }

  function refreshDutySchedule(force){
    // v37：即使時段沒有變，也先刷新「編輯」按鈕與下拉選單可用狀態。
    updateDutyPeriodEditorUi();
    if(!dutySchedule.length) return;

    const period = getActiveDutyPeriod();
    if(!period) return;

    const key = `${period.start}|${period.end}`;
    if(!force && key === currentDutyKey) return;

    currentDutyKey = key;
    applyDutyPeriod(period);
  }

  /* =========================================================
     15. 當日人員名冊與人員 / 車輛池同步
  ========================================================= */
  function ensureRosterPerson(name,role,no){
    const cleanName = String(name || '').trim();
    if(!cleanName) return null;

    let item = todayRoster.find(x=>x.name === cleanName);

    if(!item){
      item = {
        no:no ?? '',
        name:cleanName,
        role:normalizeRole(role),
        start:'',
        end:''
      };
      todayRoster.push(item);
    }else{
      if(role) item.role = normalizeRole(role);
      if((item.no === '' || item.no === null || item.no === undefined) && no !== undefined){
        item.no = no;
      }
    }

    return item;
  }

  function clearBoardAssignments(){
    $('.drop-target [data-drag-type]').remove();
    $('#restingBody [data-drag-type="person"], #officialBody [data-drag-type="person"]').remove();
    currentDutyKey = '';
    syncVehiclePool();
    rebuildPersonPool();
    $('#currentDutyPeriod').text(dutySchedule.length ? '--:--–--:--' : '尚未匯入');
    updateDutyPeriodEditorUi();
  }

  function todayDutyNumberSet(){
    const result = new Set();
    dutySchedule.forEach(period=>{
      (period.allDutyNumbers || []).forEach(no=>result.add(String(no)));
    });
    return result;
  }

  function currentAtStationNumberSet(){
    if(!hasDetailedDutyData) return new Set();
    const period = getActiveDutyPeriod();
    return new Set((period?.['在隊備勤'] || []).map(no=>String(no)));
  }

  function rebuildPersonPool(){
    const assigned = assignedSet('person');
    const statuses = statusSet();
    const scheduledToday = todayDutyNumberSet();
    const atStationNow = currentAtStationNumberSet();
    const currentDuty = currentDutyNumberSet();
    const keyword = $('#personSearch').val().trim().toLowerCase();
    const $pool = $('#personPool').empty();

    // v43：Excel 人員只有「今天勤務表有出勤」才進人員池。
    // 前端新增的中隊長／義消等勤務表外人員，設定完成後直接供調配，
    // 不再另外維護「今日出勤」欄位。
    const merged = [];
    const seen = new Set();

    todayRoster.forEach(item=>{
      if(!item || !item.name || !isActiveShift(item)) return;
      const noText=String(item.no ?? '').trim();
      const isAttending=!hasDetailedDutyData || (!!noText && scheduledToday.has(noText));
      if(!isAttending || hasBlockingDutyStatus(noText)) return;
      if(seen.has(item.name)) return;
      seen.add(item.name);
      merged.push({...item,_source:'daily'});
    });

    manualPersonnel.forEach(item=>{
      if(!item || !item.name) return;
      if(seen.has(item.name)) return;
      seen.add(item.name);
      merged.push({...item,_source:'manual'});
    });

    const candidates = merged.filter(item=>{
      if(assigned.has(item.name) || statuses.has(item.name)) return false;
      if(keyword && !item.name.toLowerCase().includes(keyword) && !String(item.no ?? '').toLowerCase().includes(keyword) && !String(item.role ?? '').toLowerCase().includes(keyword)) return false;
      return true;
    });

    // v43：先依職務階級，再於同職務內依番號排序。
    // 無番號的中隊長仍會因職務優先權排在整個人員池最上方。
    const rolePriority={'中隊長':0,'分隊長':1,'小隊長':2,'隊員':3,'役男':3,'義消':4};
    candidates.sort((a,b)=>{
      const ar=rolePriority[normalizeRole(a.role)] ?? 99;
      const br=rolePriority[normalizeRole(b.role)] ?? 99;
      if(ar!==br) return ar-br;
      const aRaw=String(a.no ?? '').trim();
      const bRaw=String(b.no ?? '').trim();
      const aNo=/^\d+$/.test(aRaw) ? Number(aRaw) : Number.POSITIVE_INFINITY;
      const bNo=/^\d+$/.test(bRaw) ? Number(bRaw) : Number.POSITIVE_INFINITY;
      if(aNo!==bNo) return aNo-bNo;
      return String(a.name).localeCompare(String(b.name),'zh-Hant');
    });

    candidates.forEach(item=>{
      const noText=String(item.no ?? '').trim();
      const isManual=item._source==='manual';
      const isAtStationNow=!isManual && hasDetailedDutyData && !!noText && atStationNow.has(noText);
      const isOnDutyNow=!isManual && hasDetailedDutyData && !!noText && currentDuty.has(noText);

      const $item = $('<div class="pool-item"></div>')
        .attr('data-drag-type','person')
        .attr('data-value',item.name)
        .attr('data-role',normalizeRole(item.role))
        .attr('data-no',item.no ?? '')
        .attr('data-person-source',item._source)
        .toggleClass('is-at-station',isAtStationNow)
        .toggleClass('is-on-duty-now',isOnDutyNow && !isAtStationNow)
        .addClass(roleClass(item.role))
        .append($('<span></span>').text(item.name));

      const $meta=$('<small class="pool-meta"></small>');
      $meta.append($('<span></span>').text(`${noText ? noText+'號 · ' : ''}${roleLabel(item.role)}`));

      // 「請替」等附加狀態不取代勤務項目；人員池只顯示目前勤務位置。
      // 例如：22號 · 隊員 [在隊備勤]
      if(isAtStationNow){
        $meta.append($('<span class="duty-badge duty-badge-ready"></span>').text('在隊備勤'));
      }else if(isOnDutyNow){
        $meta.append($('<span class="duty-badge duty-badge-working"></span>').text('勤務中'));
      }

      $item.append($meta);
      $pool.append($item);
    });

    $('#personCount').text(candidates.length);
    initPersonPoolSortable();
  }

  /* =========================================================
     14-1. 職務池
     ---------------------------------------------------------
     v42：職務池只放「職務文字」，例如火警值班。
     中隊長、義消都屬於人員本身的職稱，設定好且今日出勤後
     直接出現在人員池，不再放到職務池。
  ========================================================= */
  function rebuildSpecialDutyPool(){
    const $pool=$('#specialDutyPool').empty();
    if(!$pool.length) return;

    const placed=new Set();
    $('.drop-target [data-person-source="duty"]').each(function(){
      const value=String($(this).attr('data-value') || '').trim();
      if(value) placed.add(value);
    });

    const duties=[...new Set(dutyMaster.map(item=>String(item || '').trim()).filter(Boolean))]
      .filter(item=>!placed.has(item))
      .sort((a,b)=>a.localeCompare(b,'zh-Hant'));

    duties.forEach(duty=>{
      $pool.append(
        $('<div class="pool-item duty-pool-item"></div>')
          .attr('data-drag-type','person')
          .attr('data-value',duty)
          .attr('data-role','職務')
          .attr('data-no','')
          .attr('data-duty-role',duty)
          .attr('data-person-source','duty')
          .append($('<span></span>').text(duty))
          .append($('<small></small>').text('職務'))
      );
    });

    $('#specialDutyCount').text(duties.length);
    initSpecialDutySortable();
  }

  let specialDutySortable=null;
  function initSpecialDutySortable(){
    if(specialDutySortable){specialDutySortable.destroy();specialDutySortable=null;}
    const el=$('#specialDutyPool')[0];
    if(!el) return;
    specialDutySortable=new Sortable(el,{
      group:{name:'people',pull:'clone',put:false},
      sort:false,
      draggable:'.duty-pool-item',
      animation:120,
      delay:300,
      delayOnTouchOnly:true,
      touchStartThreshold:4,
      fallbackOnBody:true,
      fallbackTolerance:5
    });
  }

  /* =========================================================
     14-2. 前端資料管理：人員／職務／車輛
     ---------------------------------------------------------
     - 主視窗使用三個標籤頁，一次只顯示一類。
     - 人員的「職務」是人員職稱／身分，可自由輸入，例如中隊長、義消。
     - 職務標籤管理的是職務池文字，例如火警值班。
     - 前端新增的勤務表外人員設定完成後直接進人員池，不另設「今日出勤」。
  ========================================================= */
  let dataEditorContext=null;
  let dataEditorReturnToMaster=false;

  function refreshMasterDataUi(){
    renderPersonManageTable();
    renderRoleManageTable();
    renderVehicleManageTable();
  }

  function personManageRows(){
    const scheduledToday=todayDutyNumberSet();
    const rows=[
      ...todayRoster.map((x,i)=>({...x,_source:'Excel／當日',_kind:'daily',_index:i,_attending:(!hasDetailedDutyData || (!!String(x.no ?? '').trim() && scheduledToday.has(String(x.no)) && !hasBlockingDutyStatus(String(x.no))))})),
      ...manualPersonnel.map((x,i)=>({...x,_source:'前端設定',_kind:'manual',_index:i}))
    ];
    rows.sort((a,b)=>{
      const ar=String(a.no ?? '').trim(),br=String(b.no ?? '').trim();
      const an=/^\d+$/.test(ar)?Number(ar):Number.POSITIVE_INFINITY;
      const bn=/^\d+$/.test(br)?Number(br):Number.POSITIVE_INFINITY;
      if(an!==bn) return an-bn;
      return String(a.name).localeCompare(String(b.name),'zh-Hant');
    });
    return rows;
  }

  function renderPersonManageTable(){
    const $body=$('#personManageBody').empty();
    personManageRows().forEach(item=>{
      const $tr=$('<tr></tr>');
      $tr.append(
        $('<td></td>').text(item.no ?? ''),
        $('<td class="fw-bold"></td>').text(item.name),
        $('<td></td>').text(item.role || ''),
        $('<td></td>').append($('<span class="badge"></span>').addClass(item._kind==='manual'?'text-bg-primary':'text-bg-light border').text(item._source))
      );
      const $a=$('<td class="text-end text-nowrap"></td>');
      $a.append(
        $('<button class="btn btn-outline-primary btn-sm me-1" type="button">編輯</button>').on('click',()=>openPersonEditor(item._kind,item._index)),
        $('<button class="btn btn-outline-danger btn-sm" type="button">刪除</button>').on('click',()=>deleteManagedPerson(item._kind,item._index))
      );
      $tr.append($a);$body.append($tr);
    });
  }

  function editorField(id,label,type='text',value='',options=null,help=''){
    const $wrap=$('<div class="col-12"></div>');
    $wrap.append($('<label class="form-label fw-bold"></label>').attr('for',id).text(label));
    let $control;
    if(Array.isArray(options)){
      $control=$('<select class="form-select"></select>').attr('id',id);
      options.forEach(option=>$control.append($('<option></option>').val(option).text(option)));
      $control.val(String(value ?? ''));
    }else{
      $control=$('<input class="form-control">').attr({id,type}).val(value ?? '');
    }
    $wrap.append($control);
    if(help) $wrap.append($('<div class="form-text"></div>').text(help));
    return $wrap;
  }

  function editorCheckField(id,label,checked,help=''){
    const $wrap=$('<div class="col-12"></div>');
    const $box=$('<div class="form-check form-switch"></div>');
    $box.append($('<input class="form-check-input" type="checkbox">').attr('id',id).prop('checked',!!checked));
    $box.append($('<label class="form-check-label fw-bold"></label>').attr('for',id).text(label));
    $wrap.append($box);
    if(help) $wrap.append($('<div class="form-text"></div>').text(help));
    return $wrap;
  }

  function openDataEditor(title,fields,context,hint=''){
    dataEditorContext=context;
    $('#dataEditorTitle').text(title);
    const $form=$('#dataEditorForm').empty();
    fields.forEach(field=>{
      if(field.kind==='check') $form.append(editorCheckField(field.id,field.label,field.checked,field.help));
      else $form.append(editorField(field.id,field.label,field.type||'text',field.value??'',field.options||null,field.help||''));
    });
    $('#dataEditorHint').toggleClass('d-none',!hint).text(hint || '');
    dataEditorReturnToMaster=$('#masterDataModal').hasClass('show');
    if(dataEditorReturnToMaster){
      $('#masterDataModal').one('hidden.bs.modal',function(){
        dataEditorModal.show();
        setTimeout(()=>$('#dataEditorForm input, #dataEditorForm select').first().trigger('focus'),150);
      });
      masterDataModal.hide();
    }else{
      dataEditorModal.show();
      setTimeout(()=>$('#dataEditorForm input, #dataEditorForm select').first().trigger('focus'),150);
    }
  }

  function openPersonEditor(kind,index){
    let current={no:'',name:'',role:'隊員'};
    if(kind==='daily') current={...(todayRoster[index] || current)};
    else if(kind==='manual') current=manualPersonnel[index] || current;

    openDataEditor(
      kind==='new'?'新增人員':'編輯人員',
      [
        {id:'editorPersonNo',label:'番號',value:current.no ?? '',help:'勤務表外人員可不填番號；排序仍以人員職務優先。'},
        {id:'editorPersonName',label:'姓名',value:current.name ?? ''},
        {id:'editorPersonRole',label:'人員職務',value:current.role ?? '',help:'自由輸入，例如：中隊長、分隊長、小隊長、隊員、義消。'}
      ],
      {type:'person',kind,index},
      kind==='daily'?'Excel 人員是否有出勤由每日勤務表自動判斷，不需要另外設定。':'勤務表外人員設定完成後會直接進入人員池供調配。'
    );
  }

  function removePersonFromPlaced(name){
    $('[data-drag-type="person"]').filter(function(){return String($(this).attr('data-value')||'')===String(name) && String($(this).attr('data-person-source')||'')!=='duty';}).remove();
  }

  function updatePlacedPerson(oldName,value,source){
    $('[data-drag-type="person"]').filter(function(){return String($(this).attr('data-value')||'')===oldName && String($(this).attr('data-person-source')||'')!=='duty';}).each(function(){
      const $el=$(this);
      $el.attr('data-value',value.name).attr('data-no',value.no ?? '').attr('data-person-source',source || 'manual').removeAttr('data-duty-role').text(value.name);
      applyPersonRole($el,value.name,value.role);
    });
  }

  function deleteManagedPerson(kind,index){
    const list=kind==='daily'?todayRoster:manualPersonnel;
    const current=list[index];if(!current)return;
    const name=current.name;
    if(!window.confirm(`確定刪除「${name}」？`)) return;
    list.splice(index,1);
    removePersonFromPlaced(name);
    syncAll();refreshMasterDataUi();queueAutoSave('delete-person');
  }

  function addOtherPerson(){openPersonEditor('new',-1);}

  /* =========================================================
     14-3. 義消批次新增
     ---------------------------------------------------------
     每行可直接貼一個姓名；也支援「番號,姓名」或從 Excel 貼上
     「番號<TAB>姓名」。批次新增的人員職務一律為「義消」。
  ========================================================= */
  function openVolunteerBatch(){
    $('#volunteerBatchText').val('');
    $('#volunteerBatchResult').addClass('d-none').text('');

    if($('#masterDataModal').hasClass('show')){
      $('#masterDataModal').one('hidden.bs.modal',function(){
        volunteerBatchModal.show();
        setTimeout(()=>$('#volunteerBatchText').trigger('focus'),150);
      });
      masterDataModal.hide();
    }else{
      volunteerBatchModal.show();
    }
  }

  function parseVolunteerBatchLine(rawLine){
    const line=String(rawLine || '').trim();
    if(!line) return null;

    let parts=line.split(/\t|,|，|;|；/).map(x=>x.trim()).filter(Boolean);
    if(parts.length>=2){
      const first=parts[0];
      const second=parts[1];
      if(/^\d+$/.test(first)) return {no:first,name:second};
      return {no:'',name:first};
    }

    const spaced=line.match(/^(\d+)\s+(.+)$/);
    if(spaced) return {no:spaced[1],name:spaced[2].trim()};

    return {no:'',name:line};
  }

  function saveVolunteerBatch(){
    const lines=String($('#volunteerBatchText').val() || '').split(/\r?\n/);
    let added=0, skipped=0;
    const existingNames=new Set([
      ...todayRoster.map(x=>String(x.name || '').trim()),
      ...manualPersonnel.map(x=>String(x.name || '').trim())
    ].filter(Boolean));
    const existingNos=new Set([
      ...todayRoster.map(x=>String(x.no ?? '').trim()),
      ...manualPersonnel.map(x=>String(x.no ?? '').trim())
    ].filter(Boolean));

    lines.forEach(line=>{
      const parsed=parseVolunteerBatchLine(line);
      if(!parsed || !parsed.name) return;
      const parsedNo=String(parsed.no ?? '').trim();
      if(existingNames.has(parsed.name) || (parsedNo && existingNos.has(parsedNo))){skipped++;return;}

      manualPersonnel.push({
        id:`manual-${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
        no:parsed.no,
        name:parsed.name,
        role:'義消'
      });
      existingNames.add(parsed.name);
      if(parsedNo) existingNos.add(parsedNo);
      added++;
    });

    if(!roleMaster.includes('義消')) roleMaster.push('義消');

    if(!added){
      $('#volunteerBatchResult').removeClass('d-none').text(skipped ? `沒有新增人員；${skipped} 筆姓名已存在。` : '沒有可新增的姓名。');
      return;
    }

    syncAll();
    refreshMasterDataUi();
    queueAutoSave('batch-volunteers',50);
    volunteerBatchModal.hide();
    toast(`已批次新增 ${added} 位義消${skipped ? `，略過 ${skipped} 筆重複姓名／番號` : ''}`);

    setTimeout(()=>{
      refreshMasterDataUi();
      masterDataModal.show();
      const trigger=document.querySelector('[data-bs-target="#personManagePane"]');
      if(trigger) bootstrap.Tab.getOrCreateInstance(trigger).show();
    },180);
  }

  // 「職務」標籤管理的是職務池項目，不是人員職稱。
  function renderRoleManageTable(){
    const $body=$('#roleManageBody').empty();
    dutyMaster.forEach((duty,index)=>{
      const $tr=$('<tr></tr>').append($('<td class="fw-bold"></td>').text(duty));
      const $a=$('<td class="text-end text-nowrap"></td>');
      $a.append(
        $('<button class="btn btn-outline-primary btn-sm me-1" type="button">編輯</button>').on('click',()=>openRoleEditor(index)),
        $('<button class="btn btn-outline-danger btn-sm" type="button">刪除</button>').on('click',()=>deleteRole(index))
      );
      $tr.append($a);$body.append($tr);
    });
  }

  function openRoleEditor(index=-1){
    const current=index>=0?dutyMaster[index]:'';
    openDataEditor(index>=0?'編輯職務池項目':'新增職務池項目',[{id:'editorRoleName',label:'職務名稱',value:current,help:'例如：火警值班。儲存後會出現在職務池，可直接拖到火警配置表。'}],{type:'role',index});
  }
  function addRole(){openRoleEditor(-1);}

  function deleteRole(index){
    const duty=dutyMaster[index];if(!duty)return;
    const used=$('.drop-target [data-person-source="duty"]').filter(function(){return String($(this).attr('data-value')||'')===duty;}).length>0;
    if(used){toast('此職務目前已放在看板上，請先放回職務池');return;}
    if(!window.confirm(`確定刪除職務「${duty}」？`)) return;
    dutyMaster.splice(index,1);rebuildSpecialDutyPool();renderRoleManageTable();queueAutoSave('delete-duty');
  }

  function renderVehicleManageTable(){
    const $body=$('#vehicleManageBody').empty();
    vehicleMaster.forEach((raw,index)=>{
      const item=typeof raw==='string'?{name:raw,type:'車輛'}:raw;
      const $tr=$('<tr></tr>').append($('<td class="fw-bold"></td>').text(item.name),$('<td></td>').text(item.type||'車輛'));
      const $a=$('<td class="text-end text-nowrap"></td>');
      $a.append(
        $('<button class="btn btn-outline-primary btn-sm me-1" type="button">編輯</button>').on('click',()=>openVehicleEditor(index)),
        $('<button class="btn btn-outline-danger btn-sm" type="button">刪除</button>').on('click',()=>deleteVehicle(index))
      );
      $tr.append($a);$body.append($tr);
    });
  }

  function openVehicleEditor(index=-1){
    const raw=index>=0?vehicleMaster[index]:null;
    const current=raw?(typeof raw==='string'?{name:raw,type:'車輛'}:raw):{name:'',type:'消防車'};
    openDataEditor(index>=0?'編輯車輛':'新增車輛',[
      {id:'editorVehicleName',label:'車輛名稱',value:current.name||''},
      {id:'editorVehicleType',label:'車輛類型',value:current.type||'車輛'}
    ],{type:'vehicle',index});
  }
  function addVehicle(){openVehicleEditor(-1);}

  function deleteVehicle(index){
    const raw=vehicleMaster[index],cur=typeof raw==='string'?{name:raw}:raw;if(!cur)return;
    if(!window.confirm(`確定刪除車輛「${cur.name}」？`))return;
    vehicleMaster.splice(index,1);
    $('[data-drag-type="vehicle"]').filter(function(){return String($(this).attr('data-value')||'')===cur.name;}).remove();
    renderVehiclePool();syncVehiclePool();renderVehicleManageTable();queueAutoSave('delete-vehicle');
  }

  function saveDataEditor(){
    const ctx=dataEditorContext;if(!ctx)return;

    if(ctx.type==='person'){
      const no=String($('#editorPersonNo').val()||'').trim();
      const name=String($('#editorPersonName').val()||'').trim();
      const role=String($('#editorPersonRole').val()||'').trim();
      if(!name){toast('姓名不可空白');return;}
      if(!role){toast('人員職務不可空白');return;}
      if(!roleMaster.includes(role)) roleMaster.push(role);

      if(ctx.kind==='new'){
        if(todayRoster.some(x=>x.name===name)||manualPersonnel.some(x=>x.name===name)){toast('此姓名已存在');return;}
        manualPersonnel.push({id:`manual-${Date.now()}`,no,name,role});
      }else if(ctx.kind==='daily'){
        const current=todayRoster[ctx.index];if(!current)return;
        const oldName=current.name;
        todayRoster[ctx.index]={...current,no,name,role};
        updatePlacedPerson(oldName,{no,name,role},'daily');
      }else{
        const current=manualPersonnel[ctx.index];if(!current)return;
        const oldName=current.name;
        manualPersonnel[ctx.index]={...current,no,name,role};
        updatePlacedPerson(oldName,{no,name,role},'manual');
      }
      dataEditorModal.hide();syncAll();refreshMasterDataUi();queueAutoSave('person-data');toast('人員資料已更新');return;
    }

    if(ctx.type==='role'){
      const newDuty=String($('#editorRoleName').val()||'').trim();
      if(!newDuty){toast('職務名稱不可空白');return;}
      if(ctx.index<0){
        if(dutyMaster.includes(newDuty)){toast('職務已存在');return;}
        dutyMaster.push(newDuty);
      }else{
        const oldDuty=dutyMaster[ctx.index];if(!oldDuty)return;
        if(newDuty!==oldDuty && dutyMaster.includes(newDuty)){toast('職務已存在');return;}
        dutyMaster[ctx.index]=newDuty;
        $('.drop-target [data-person-source="duty"]').filter(function(){return String($(this).attr('data-value')||'')===oldDuty;}).each(function(){
          $(this).attr('data-value',newDuty).attr('data-duty-role',newDuty).text(newDuty);
        });
      }
      dataEditorModal.hide();rebuildSpecialDutyPool();refreshMasterDataUi();queueAutoSave('duty-data');toast('職務資料已更新');return;
    }

    if(ctx.type==='vehicle'){
      const name=String($('#editorVehicleName').val()||'').trim();
      const type=String($('#editorVehicleType').val()||'').trim()||'車輛';
      if(!name){toast('車輛名稱不可空白');return;}
      const duplicate=vehicleMaster.some((x,i)=>i!==ctx.index && (typeof x==='string'?x:x.name)===name);
      if(duplicate){toast('車輛名稱已存在');return;}
      if(ctx.index<0){vehicleMaster.push({name,type});}
      else{
        const raw=vehicleMaster[ctx.index],old=typeof raw==='string'?raw:raw.name;
        vehicleMaster[ctx.index]={name,type};
        $('[data-drag-type="vehicle"]').filter(function(){return String($(this).attr('data-value')||'')===old;}).attr('data-value',name).text(name);
      }
      dataEditorModal.hide();renderVehiclePool();syncVehiclePool();refreshMasterDataUi();queueAutoSave('vehicle-data');toast('車輛資料已更新');
    }
  }

  function renderVehiclePool(){
    const $pool = $('#vehiclePool').empty();

    vehicleMaster.forEach(item=>{
      const name = typeof item === 'string' ? item : item.name;
      const type = typeof item === 'string' ? '車輛' : (item.type || '車輛');

      $pool.append(
        $('<div class="pool-item vehicle"></div>')
          .attr('data-drag-type','vehicle')
          .attr('data-value',name)
          .append($('<span></span>').text(name))
          .append($('<small></small>').text(type))
      );
    });
  }

  function syncVehiclePool(){
    const assigned = assignedSet('vehicle');
    let count = 0;

    $('#vehiclePool .pool-item').each(function(){
      const value = String($(this).data('value'));
      const show = !assigned.has(value);
      $(this).toggle(show);
      if(show) count++;
    });

    $('#vehicleCount').text(count);
  }

  /* =========================================================
     15-1. 看板姓名 / 車輛字體自動放到「不換行的最大值」
     ---------------------------------------------------------
     CSS clamp 只能依 viewport 猜字級，無法知道某一格實際有多寬、
     姓名有幾個字。這裡直接量測每一個主看板 / 專責救護格子的
     可用寬高，以二分搜尋找出仍可完整顯示且不換行的最大字級。

     適用：
     - 第一梯次 / 第二梯次的人員姓名
     - 第一梯次 / 第二梯次的車輛
     - 專責救護的人員姓名 / 車輛

     不處理休息 / 因公外出，避免狀態區多人排列時被單一卡片搶滿。
  ========================================================= */
  let fitBoardTextFrame = null;

  function getFitTextNode(element){
    // 目前 person-chip / vehicle-chip 為純文字；預留 span 結構相容性。
    return element.querySelector('span') || element;
  }

  function fitBoardChipText(element){
    if(!element || !element.parentElement) return;

    const parent = element.parentElement;
    const value = String(element.dataset.value || element.textContent || '').trim();
    if(!value) return;

    const parentStyle = window.getComputedStyle(parent);
    const horizontalPadding = parseFloat(parentStyle.paddingLeft || 0) + parseFloat(parentStyle.paddingRight || 0);
    const verticalPadding = parseFloat(parentStyle.paddingTop || 0) + parseFloat(parentStyle.paddingBottom || 0);
    const availableWidth = Math.max(0,parent.clientWidth - horizontalPadding - 6);
    const availableHeight = Math.max(0,parent.clientHeight - verticalPadding - 4);

    if(availableWidth <= 0 || availableHeight <= 0) return;

    const textNode = getFitTextNode(element);
    const computed = window.getComputedStyle(element);
    const measure = document.createElement('span');
    measure.textContent = value;
    measure.style.position = 'fixed';
    measure.style.left = '-10000px';
    measure.style.top = '-10000px';
    measure.style.visibility = 'hidden';
    measure.style.pointerEvents = 'none';
    measure.style.whiteSpace = 'nowrap';
    measure.style.fontFamily = computed.fontFamily;
    measure.style.fontWeight = computed.fontWeight;
    measure.style.fontStyle = computed.fontStyle;
    measure.style.letterSpacing = computed.letterSpacing;
    measure.style.lineHeight = '1';
    document.body.appendChild(measure);

    // 上限刻意給得比目前 CSS 大；真正上限由該格寬、高共同決定。
    let low = 8;
    let high = Math.max(low,Math.min(64,Math.floor(availableHeight)));
    let best = low;

    while(low <= high){
      const mid = Math.floor((low + high) / 2);
      measure.style.fontSize = `${mid}px`;
      const rect = measure.getBoundingClientRect();

      if(rect.width <= availableWidth && rect.height <= availableHeight){
        best = mid;
        low = mid + 1;
      }else{
        high = mid - 1;
      }
    }

    measure.remove();

    element.style.setProperty('font-size',`${best}px`,'important');
    element.style.setProperty('line-height','1','important');
    element.style.setProperty('white-space','nowrap','important');
    element.style.setProperty('overflow','hidden','important');
    element.style.setProperty('text-overflow','clip','important');

    if(textNode !== element){
      textNode.style.whiteSpace = 'nowrap';
    }
  }

  function fitAllBoardText(){
    /*
       v59：主表是字級基準，專責救護只能跟隨，不能自己放大。
       原本 rescue-table 的儲存格比火警主表單列高，因此 fitBoardChipText()
       會替 91/92 姓名算出更大的字級；即使 CSS 寫成一樣，inline !important
       仍會讓專責救護看起來大一截。

       正確規則：
       1. 只讓 fire-table 的姓名／車號自行計算最大可用字級。
       2. 取 fire-table 同類型中最小值作為整張看板共同字級。
       3. rescue-table 直接套用這個值，不再依自己的列高計算。
    */
    const firePersons=$('.fire-table .person-chip').toArray();
    const fireVehicles=$('.fire-table .vehicle-chip').toArray();
    const rescuePersons=$('.rescue-table .person-chip').toArray();
    const rescueVehicles=$('.rescue-table .vehicle-chip').toArray();

    const clearInlineFont=function(items){
      items.forEach(el=>{
        el.style.removeProperty('font-size');
        el.style.removeProperty('line-height');
      });
    };

    const fitFireAndGetCommonSize=function(items){
      if(!items.length) return null;
      clearInlineFont(items);
      items.forEach(el=>fitBoardChipText(el));
      const sizes=items
        .map(el=>parseFloat(window.getComputedStyle(el).fontSize || '0'))
        .filter(v=>Number.isFinite(v) && v>0);
      return sizes.length ? Math.min(...sizes) : null;
    };

    const applyCommonSize=function(items,size){
      if(!size) return;
      items.forEach(el=>{
        el.style.setProperty('font-size',`${size}px`,'important');
        el.style.setProperty('line-height','1','important');
        el.style.setProperty('white-space','nowrap','important');
        el.style.setProperty('overflow','hidden','important');
        el.style.setProperty('text-overflow','clip','important');
      });
    };

    const personSize=fitFireAndGetCommonSize(firePersons);
    const vehicleSize=fitFireAndGetCommonSize(fireVehicles);

    // 主表全部統一到共同最小值。
    applyCommonSize(firePersons,personSize);
    applyCommonSize(fireVehicles,vehicleSize);

    // 專責救護完全跟隨主表；不得再自己依較高的儲存格放大。
    clearInlineFont(rescuePersons);
    clearInlineFont(rescueVehicles);
    applyCommonSize(rescuePersons,personSize);
    applyCommonSize(rescueVehicles,vehicleSize);

    // 若主表暫時沒有同類卡片，才以救護區自身為備援基準。
    if(!personSize && rescuePersons.length){
      const fallback=fitFireAndGetCommonSize(rescuePersons);
      applyCommonSize(rescuePersons,fallback);
    }
    if(!vehicleSize && rescueVehicles.length){
      const fallback=fitFireAndGetCommonSize(rescueVehicles);
      applyCommonSize(rescueVehicles,fallback);
    }
  }

  let fitBoardTextTimer = null;

  function scheduleFitBoardText(){
    if(fitBoardTextFrame){
      cancelAnimationFrame(fitBoardTextFrame);
    }
    if(fitBoardTextTimer){
      clearTimeout(fitBoardTextTimer);
      fitBoardTextTimer = null;
    }

    // Sortable 放下卡片時，DOM 已移動但儲存格尺寸 / 動畫可能尚未穩定。
    // 連續等兩個 animation frame 後再量測，避免用到拖曳中的暫時尺寸。
    fitBoardTextFrame = requestAnimationFrame(function(){
      fitBoardTextFrame = requestAnimationFrame(function(){
        fitBoardTextFrame = null;
        fitAllBoardText();
      });
    });

    // 再補一次延遲校正，涵蓋 Sortable animation:120ms 與行動裝置 reflow。
    fitBoardTextTimer = setTimeout(function(){
      fitBoardTextTimer = null;
      fitAllBoardText();
    },180);
  }

  function syncAll(){
    enforceUniqueBoardAssignments();
    $('.drop-target').each(function(){ setEmpty($(this)); });

    $('[data-drag-type="person"]').each(function(){
      const $item = $(this);
      if(String($item.attr('data-person-source') || '') === 'duty') return;
      const name = String($item.data('value') || $item.text().trim());
      applyPersonRole($item,name,$item.attr('data-role') || findRosterRole(name));
    });

    syncVehiclePool();
    rebuildPersonPool();
    rebuildSpecialDutyPool();
    scheduleFitBoardText();
  }

  /* =========================================================
     16. 拖曳返回區顯示控制
  ========================================================= */
  function showReturnZone(kind,source=''){
    const text = kind === 'vehicle' ? '拖到這裡：放回車輛池' : (source === 'duty' ? '拖到這裡：放回職務池' : '拖到這裡：放回人員池');
    $('#returnZone').text(text).addClass('show');
  }

  function hideReturnZone(){
    $('#returnZone').removeClass('show');
  }

  /* =========================================================
     17. SortableJS：人員池、車輛池與主看板拖曳
     ---------------------------------------------------------
     每次放開後 queueAutoSave()，所以不需要儲存按鈕。
  ========================================================= */
  let personPoolSortable = null;

  function initPersonPoolSortable(){
    if(personPoolSortable){
      personPoolSortable.destroy();
      personPoolSortable = null;
    }

    personPoolSortable = new Sortable($('#personPool')[0],{
      group:{name:'people',pull:'clone',put:false},
      sort:false,
      filter:'.pool-item.is-unavailable',
      preventOnFilter:true,
      animation:120,
      delay:300,
      delayOnTouchOnly:true,
      touchStartThreshold:4,
      fallbackOnBody:true,
      fallbackTolerance:5
    });
  }

  new Sortable($('#vehiclePool')[0],{
    group:{name:'vehicles',pull:'clone',put:false},
    sort:false,
    animation:120,
    delay:300,
    delayOnTouchOnly:true,
    touchStartThreshold:4,
    fallbackOnBody:true,
    fallbackTolerance:5
  });

  $('.drop-target').each(function(){
    const el = this;
    const $target = $(this);
    const kind = $target.data('kind');
    const groupName = kind === 'vehicle' ? 'vehicles' : 'people';

    new Sortable(el,{
      group:{name:groupName,pull:true,put:[groupName]},
      sort:false,
      animation:120,
      delay:300,
      delayOnTouchOnly:true,
      touchStartThreshold:4,
      fallbackOnBody:true,
      fallbackTolerance:5,
      draggable:'[data-drag-type]',
      ghostClass:'sortable-ghost',
      chosenClass:'sortable-chosen',

      // 不只依賴 Sortable group，再檢查一次實際 item / target 型別。
      // 例如 11車(type=vehicle) 只能進入 data-kind=vehicle 的車輛欄。
      onMove:function(evt){
        // 返回區不是正式配置格，因此沒有 data-kind。
        // 若拖曳目標是返回區，直接允許人員 / 車輛離開目前配置格；
        // 其他目標才依 data-kind 嚴格限制型別。
        if(evt.to && evt.to.id === 'returnZone') return true;
        return isItemAllowedInDropTarget(evt.dragged,evt.to);
      },

      onStart:function(evt){
        // 開始拖曳時暫停遠端重畫，避免操作中被輪詢結果覆蓋。
        isDragging = true;
        showReturnZone(kind,String($(evt.item).attr('data-person-source') || ''));
      },

      onAdd:function(evt){
        const $item = $(evt.item);

        // 最後一道保護：若來源型別與欄位型別不同，立即退回並不保存。
        if(!isItemAllowedInDropTarget(evt.item,$target[0])){
          $item.remove();
          toast(kind === 'vehicle' ? '這一格只能放車輛' : '車輛只能放在「車輛」欄');
          setTimeout(syncAll,0);
          return;
        }

        const value = String($item.data('value') || $item.text().trim());

        $target.find(`[data-drag-type="${kind}"]`).not(evt.item).remove();

        let $replacement;
        if(kind === 'vehicle'){
          $replacement = $('<div class="vehicle-chip"></div>')
            .attr('data-drag-type','vehicle')
            .attr('data-value',value)
            .text(value);
        }else{
          const role = $item.attr('data-role') || findRosterRole(value);
          const no = $item.attr('data-no') ?? '';
          const dutyRole = String($item.attr('data-duty-role') || '').trim();
          const personSource = String($item.attr('data-person-source') || (dutyRole ? 'duty' : 'daily')).trim();
          if(personSource === 'daily') ensureRosterPerson(value,role,no);

          $replacement = $('<div class="person-chip"></div>')
            .attr('data-drag-type','person')
            .attr('data-value',value)
            .attr('data-no',no)
            .attr('data-person-source',personSource)
            .text(value);
          if(dutyRole) $replacement.attr('data-duty-role',dutyRole);

          if(personSource === 'duty') $replacement.addClass('duty-assignment-chip');
          else applyPersonRole($replacement,value,role);
        }

        $item.replaceWith($replacement);
        $target.find('.empty-slot').remove();

        // 新卡片不可沿用池子的字級；等 Sortable 完成版面配置後，
        // 一律重新依看板儲存格尺寸計算，並與同區卡片統一字級。
        $replacement[0].style.removeProperty('font-size');
        scheduleFitBoardText();

        setTimeout(function(){
          // 放入「火警值班」後立即把目前值班番號對應人員帶到右邊那一格。
          if(personSource === 'duty' && dutyRole === '火警值班') applyFireDutyWatch(getActiveDutyPeriod(),{onlyEmpty:true});
          syncAll();
          queueAutoSave('drag');
        },0);
      },

      onRemove:function(evt){
        setTimeout(function(){
          const dutyRole=String($(evt.item).attr('data-duty-role') || '').trim();
          if(dutyRole === '火警值班') applyFireDutyWatch(getActiveDutyPeriod(),{onlyEmpty:true});
          syncAll();
          queueAutoSave('drag');
        },0);
      },

      onEnd:function(){
        isDragging = false;
        hideReturnZone();
        setTimeout(syncAll,0);
      }
    });
  });

  /* =========================================================
     18. SortableJS：休息與因公外出
  ========================================================= */
  function initStatusSortable(selector){
    new Sortable($(selector)[0],{
      group:{name:'people',pull:true,put:['people']},
      sort:true,
      animation:120,
      delay:300,
      delayOnTouchOnly:true,
      touchStartThreshold:4,
      fallbackOnBody:true,
      fallbackTolerance:5,
      draggable:'[data-drag-type="person"]',
      ghostClass:'sortable-ghost',
      chosenClass:'sortable-chosen',

      // 休息 / 因公外出只接受人員，車輛一律拒絕。
      onMove:function(evt){
        return dragItemType(evt.dragged) === 'person' && String($(evt.dragged).attr('data-person-source') || '') !== 'duty';
      },

      onStart:function(){
        // 休息 / 因公外出同樣屬於正式看板資料，拖曳後立即同步。
        isDragging = true;
        showReturnZone('person');
      },

      onAdd:function(evt){
        const $item = $(evt.item);
        const value = String($item.data('value') || $item.text().trim());

        if(!$item.hasClass('status-chip')){
          const role = $item.attr('data-role') || findRosterRole(value);
          const no = $item.attr('data-no') ?? '';
          const personSource = String($item.attr('data-person-source') || 'daily').trim();
          if(personSource === 'daily') ensureRosterPerson(value,role,no);
          const $replacement = $('<div class="status-chip"></div>')
            .attr('data-drag-type','person')
            .attr('data-value',value)
            .attr('data-no',no)
            .attr('data-person-source',personSource)
            .append($('<span></span>').text(value));

          applyPersonRole($replacement,value,role);

          $item.replaceWith($replacement);
        }

        setTimeout(function(){
          syncAll();
          queueAutoSave('status');
        },0);
      },

      onRemove:function(){
        setTimeout(function(){
          syncAll();
          queueAutoSave('status');
        },0);
      },

      onEnd:function(){
        isDragging = false;
        hideReturnZone();
        setTimeout(function(){
          syncAll();
          queueAutoSave('status-order');
        },0);
      }
    });
  }

  initStatusSortable('#restingBody');
  initStatusSortable('#officialBody');

  new Sortable($('#returnZone')[0],{
    group:{name:'returnPool',put:['people','vehicles'],pull:false},
    sort:false,
    onAdd:function(evt){
      const $item = $(evt.item);
      const kind = $item.attr('data-drag-type');

      if(kind === 'person'){
        const name = String($item.data('value') || $item.text().trim());
        const role = $item.attr('data-role') || findRosterRole(name);
        const no = $item.attr('data-no') ?? '';
        const source = String($item.attr('data-person-source') || '').trim();
        // 前端新增人員與職務文字回池時，不要誤寫進 Excel 當日名冊。
        if(source !== 'manual' && source !== 'duty'){
          ensureRosterPerson(name,role,no);
        }
      }

      $item.remove();
      hideReturnZone();

      // 等 Sortable 完成 DOM 移動後再重建池，避免 assignedSet 還讀到舊位置
      setTimeout(function(){
        syncAll();
        queueAutoSave('return-pool');
      },20);
    }
  });

  /* =========================================================
     19. 人員池搜尋
  ========================================================= */
  $('#personSearch').on('input',rebuildPersonPool);
  $('#masterDataBtn').on('click',function(){refreshMasterDataUi();masterDataModal.show();});
  $('#addOtherPersonBtn').on('click',addOtherPerson);
  $('#batchVolunteerBtn').on('click',openVolunteerBatch);
  $('#saveVolunteerBatchBtn').on('click',saveVolunteerBatch);
  $('#addRoleBtn').on('click',addRole);
  $('#addVehicleBtn').on('click',addVehicle);
  $('#dataEditorSaveBtn').on('click',saveDataEditor);
  $('#dataEditorForm').on('submit',function(e){e.preventDefault();saveDataEditor();});
  $('#dataEditorModal').on('hidden.bs.modal',function(){
    if(dataEditorReturnToMaster){
      dataEditorReturnToMaster=false;
      refreshMasterDataUi();
      masterDataModal.show();
    }
  });

  /* =========================================================
     19-1. 勤務時段測試編輯
     ---------------------------------------------------------
     - 「編輯」可從已匯入的勤務時段中手動選一段。
     - 測試模式只影響目前瀏覽器，不寫入 Google Sheet / localStorage。
     - 「依現在時間」會重新讀取正式 boardState，再套用目前真實時段。
  ========================================================= */
  $('#dutyPeriodEditBtn').on('click',function(){
    if(!dutySchedule.length){
      toast('請先匯入每日勤務表');
      return;
    }
    updateDutyPeriodEditorUi();
    $('#dutyPeriodEditor').toggleClass('d-none');
  });

  $('#dutyPeriodApplyBtn').on('click',function(){
    const key = String($('#dutyPeriodSelect').val() || '').trim();
    const period = getDutyPeriodByKey(key);
    if(!period){
      toast('找不到選擇的勤務時段');
      return;
    }

    dutyPeriodOverrideKey = key;
    currentDutyKey = '';
    refreshDutySchedule(true);
    setSyncStatus('測試模式：不自動同步','text-warning');
    $('#dutyPeriodEditor').addClass('d-none');
    toast(`測試勤務已切換為 ${period.start}–${period.end}`);
  });

  $('#dutyPeriodNowBtn').on('click',async function(){
    dutyPeriodOverrideKey = '';
    currentDutyKey = '';
    $('#dutyPeriodEditor').addClass('d-none');
    updateDutyPeriodEditorUi();

    // 測試過程的畫面不應成為正式資料；先重新讀取最後一次正式 boardState。
    await loadPublishedState(false);
    currentDutyKey = '';
    refreshDutySchedule(true);
    updateDutyPeriodEditorUi();
    toast('已恢復依現在時間自動切換勤務');
  });

  /* =========================================================
     20. 全螢幕顯示
     ---------------------------------------------------------
     瀏覽器安全規則通常禁止「完全沒有使用者操作」就進入 Fullscreen。
     因此：
     1. 頁面載入時先直接嘗試一次。
     2. 若瀏覽器阻擋，第一次點擊／觸控／按鍵時立即自動進全螢幕。
     使用者不需要再另外按「全螢幕」按鈕。
  ========================================================= */
  async function requestPageFullscreen(showError=false){
    if(document.fullscreenElement || !document.documentElement.requestFullscreen) return true;
    try{
      await document.documentElement.requestFullscreen({navigationUI:'hide'});
      return true;
    }catch(err){
      if(showError){
        console.error(err);
        toast('此瀏覽器無法進入全螢幕');
      }
      return false;
    }
  }

  function enableAutoFullscreen(){
    if(appConfig?.app?.autoFullscreen === false) return;

    // 有些 kiosk / 已授權環境可在載入時直接成功。
    requestPageFullscreen(false);

    // 一般 Chrome / Edge 需要 user gesture；第一次操作就自動完成。
    const activate=async function(event){
      // 全螢幕按鈕本身保留原本的切換行為，避免 pointerdown 先進入、click 又立刻退出。
      if($(event?.target).closest('#fullscreenBtn').length) return;
      if(document.fullscreenElement) return;
      const ok=await requestPageFullscreen(false);
      if(ok){
        document.removeEventListener('pointerdown',activate,true);
        document.removeEventListener('touchstart',activate,true);
        document.removeEventListener('keydown',activate,true);
      }
    };
    document.addEventListener('pointerdown',activate,true);
    document.addEventListener('touchstart',activate,true);
    document.addEventListener('keydown',activate,true);
  }

  $('#fullscreenBtn').on('click',async function(){
    if(!document.fullscreenElement){
      await requestPageFullscreen(true);
      return;
    }
    try{ await document.exitFullscreen(); }
    catch(err){ console.error(err); }
  });

  $(document).on('fullscreenchange',function(){
    const active = !!document.fullscreenElement;
    $('body').toggleClass('is-fullscreen',active);
    $('#fullscreenText').text(active ? '退出全螢幕' : '全螢幕');
    scheduleFitBoardText();
  });

  // 視窗縮放時，每一格實際寬高都會變化；重新計算最大不換行字級。
  let fitResizeTimer = null;
  $(window).on('resize',function(){
    clearTimeout(fitResizeTimer);
    fitResizeTimer = setTimeout(scheduleFitBoardText,60);
  });

  /* =========================================================
     21. Excel：依勤務分配表版型自動解析
     - 人員番號 / 姓名 / 身分
     - 備勤91 -> 91車
     - 備勤救災 -> 92車
     - 在隊備勤 -> 人員池標註
     - 休息時間 -> 自動放入休息區
     - 全勤務番號 -> 判斷未列勤務人員反灰
     - 輪休 / 公假 / 請休等 -> 人員池顯示狀態並控制可排性
  ========================================================= */

  let pendingImport = null;

  $('#importBtn').on('click',function(){
    $('#excelFile').trigger('click');
  });

  function cleanText(value){
    return String(value ?? '')
      .replace(/\r?\n/g,'')
      .replace(/\s+/g,'')
      .replace(/[（）()【】\[\]、，,。．.]/g,'')
      .trim();
  }

  /* =========================================================
     21-1. 人員身分 / 番號 / 姓名基本解析
     ---------------------------------------------------------
     實際勤務表的人員名冊不是固定三欄格式，常見至少三種：
     1. 番號 | 「隊員 王小明」
     2. 「役男」 | 番號 | 姓名
     3. 番號 | 身分 | 姓名

     因此不能只找「身分獨立一格」；這裡同時支援
     「身分＋姓名在同一格」與「身分、姓名分開」兩種版型。
  ========================================================= */
  function roleFromText(value){
    const text = cleanText(value);
    const supported = (appConfig && appConfig.importSettings && appConfig.importSettings.supportedRoles)
      || ['隊員','小隊長','役男','分隊長'];

    return supported.includes(text) ? text : '';
  }

  function parseRoleAndName(value){
    const original = String(value ?? '').replace(/\r?\n/g,' ').trim();
    const compact = original.replace(/\s+/g,'').trim();
    const supported = (appConfig && appConfig.importSettings && appConfig.importSettings.supportedRoles)
      || ['隊員','小隊長','役男','分隊長'];

    // 角色名稱長短不同，先比對較長者，避免未來新增相近名稱時誤判。
    const roles = [...supported].sort((a,b)=>b.length-a.length);

    for(const role of roles){
      const normalizedRole = String(role).replace(/\s+/g,'');
      if(!compact.startsWith(normalizedRole)) continue;

      const name = compact.slice(normalizedRole.length).trim();
      if(nameFromCell(name)){
        return {role:normalizedRole,name};
      }
    }

    return {role:'',name:''};
  }

  function numberFromCell(value){
    const text = String(value ?? '').trim();
    const match = text.match(/^\s*(\d{1,3})\s*$/);
    return match ? Number(match[1]) : null;
  }

  function nameFromCell(value){
    const text = String(value ?? '').replace(/\s+/g,'').trim();
    if(!text) return '';
    if(/^\d+$/.test(text)) return '';
    if(roleFromText(text)) return '';

    // 避免把表頭 / 勤務文字誤判為姓名；中文姓名通常 2~6 字，
    // 但保留到 10 字以兼容較長姓名或特殊姓名。
    if(text.length < 2 || text.length > 10) return '';
    if(/時間|勤務|備勤|休息|服勤|支援|梯次|車輛|駕駛|瞄子|搜救|照相|火警|公差|請假|輪休/.test(text)) return '';
    return text;
  }

  function parseNumberList(values){
    const result = [];

    (Array.isArray(values) ? values : [values]).forEach(value=>{
      const text = String(value ?? '');
      const nums = text.match(/\d{1,3}/g) || [];

      nums.forEach(n=>{
        const valueNo = Number(n);
        if(valueNo > 0 && valueNo <= 999 && !result.includes(valueNo)){
          result.push(valueNo);
        }
      });
    });

    return result;
  }

  function parseTimeRange(value){
    const text = String(value ?? '')
      .replace(/\s+/g,'')
      .replace(/[～~至]/g,'-')
      .replace(/[－—–]/g,'-');

    const match = text.match(/^(\d{1,2})\s*-\s*(\d{1,2})(?:時)?$/);
    if(!match) return null;

    const startHour = Number(match[1]);
    const endHour = Number(match[2]);

    if(startHour < 0 || startHour > 23 || endHour < 0 || endHour > 24){
      return null;
    }

    return {
      start:`${String(startHour % 24).padStart(2,'0')}:00`,
      end:`${String(endHour % 24).padStart(2,'0')}:00`
    };
  }

  function getMergeSpan(sheet,row,col){
    const merges = sheet['!merges'] || [];

    const merge = merges.find(item=>
      row >= item.s.r && row <= item.e.r &&
      col >= item.s.c && col <= item.e.c
    );

    if(!merge){
      return {startRow:row,endRow:row,startCol:col,endCol:col};
    }

    return {
      startRow:merge.s.r,
      endRow:merge.e.r,
      startCol:merge.s.c,
      endCol:merge.e.c
    };
  }

  // Excel「合併後置中」只在合併區左上角保存值。
  // 讀取勤務資料時，任何位於合併範圍內的格都應取得左上角值，
  // 才能正確表示例如休息番號跨 10-11、11-12 兩個時段，
  // 或只跨單一時段的情況。
  function mergedCellValue(sheet,matrix,row,col){
    const span=getMergeSpan(sheet,row,col);
    return matrix?.[span.startRow]?.[span.startCol] ?? '';
  }

  function sheetMatrix(sheet){
    return XLSX.utils.sheet_to_json(sheet,{
      header:1,
      defval:'',
      raw:false,
      blankrows:true
    });
  }

  function findExactKeywordCell(matrix,keywords){
    const normalizedKeywords = keywords.map(cleanText).filter(Boolean);

    for(let r=0;r<matrix.length;r++){
      const row = matrix[r] || [];
      for(let c=0;c<row.length;c++){
        const text = cleanText(row[c]);
        if(!text) continue;
        if(normalizedKeywords.includes(text)){
          return {row:r,col:c,text:String(row[c] ?? '')};
        }
      }
    }
    return null;
  }

  function findKeywordCell(matrix,keywords){
    const normalizedKeywords = keywords.map(cleanText);

    for(let r=0;r<matrix.length;r++){
      const row = matrix[r] || [];

      for(let c=0;c<row.length;c++){
        const text = cleanText(row[c]);
        if(!text) continue;

        if(normalizedKeywords.some(keyword=>{
          return text === keyword || text.includes(keyword) || keyword.includes(text);
        })){
          return {row:r,col:c,text:String(row[c] ?? '')};
        }
      }
    }

    return null;
  }

  /* =========================================================
     21-2. 人員名冊偵測
     ---------------------------------------------------------
     以「番號」作為穩定錨點，再向左右搜尋身分與姓名。
     這樣可直接讀取目前提供的勤務表 Excel：
     - R19=2、S19="小隊長 周文信"
     - X19=1、Y19="分隊長 陳宗揚"
     - Q36="役男"、R36=37、S36="曾得安"

     同一番號若在勤務內容出現很多次，只會保留真正找到
     「姓名 + 身分」的完整人員資料，不會把勤務數字誤當名冊。
  ========================================================= */
  function detectPersonnel(matrix){
    const byNo = new Map();

    for(let r=0;r<matrix.length;r++){
      const row = matrix[r] || [];

      for(let c=0;c<row.length;c++){
        const no = numberFromCell(row[c]);
        if(no === null || no <= 0) continue;

        let role = '';
        let name = '';

        // A. 最常見：番號右邊緊鄰就是「身分 姓名」。
        // 只看右邊第一格，避免把左側勤務區的數字誤配到
        // 同一列較遠的人員名冊，例如 P 欄的勤務數字誤配 S 欄姓名。
        if(c+1 < row.length){
          const parsed = parseRoleAndName(row[c+1]);
          if(parsed.role && parsed.name){
            role = parsed.role;
            name = parsed.name;
          }
        }

        // B. 身分在番號左邊、姓名在右邊，例如：役男 | 37 | 曾得安。
        if(!name){
          // 實際勤務表的「役男 | 番號 | 姓名」就是左右各一格。
          // 限定相鄰可大幅降低勤務區數字被誤判成人員番號。
          const candidateRole = c-1 >= 0 ? roleFromText(row[c-1]) : '';
          const candidateName = c+1 < row.length ? nameFromCell(row[c+1]) : '';

          if(candidateRole && candidateName){
            role = candidateRole;
            name = candidateName;
          }
        }

        // C. 番號右邊先是身分，再下一格才是姓名。
        if(!name && c+2 < row.length){
          // 兼容「番號 | 身分 | 姓名」三格格式，同樣只接受相鄰欄位。
          const candidateRole = roleFromText(row[c+1]);
          const candidateName = nameFromCell(row[c+2]);
          if(candidateRole && candidateName){
            role = candidateRole;
            name = candidateName;
          }
        }

        if(role && name){
          byNo.set(no,{
            no,
            name,
            role:normalizeRole(role),
            start:'',
            end:''
          });
        }
      }
    }

    return [...byNo.values()].sort((a,b)=>a.no-b.no);
  }

  function detectTimeColumn(matrix){
    const counts = new Map();

    for(let r=0;r<matrix.length;r++){
      const row = matrix[r] || [];

      for(let c=0;c<row.length;c++){
        if(parseTimeRange(row[c])){
          counts.set(c,(counts.get(c) || 0) + 1);
        }
      }
    }

    let bestCol = null;
    let bestCount = 0;

    counts.forEach((count,col)=>{
      if(count > bestCount){
        bestCount = count;
        bestCol = col;
      }
    });

    return bestCount >= 4 ? bestCol : null;
  }

  function valuesAcross(row,startCol,endCol){
    const values = [];
    for(let c=startCol;c<=endCol;c++){
      values.push((row || [])[c] ?? '');
    }
    return values;
  }

  function numbersFromColumnState(columnState,startCol,endCol){
    const result = [];
    for(let c=startCol;c<=endCol;c++){
      (columnState.get(c) || []).forEach(no=>{
        if(!result.includes(no)) result.push(no);
      });
    }
    return result;
  }

  function addStatusForNumber(map,no,status){
    const key = String(no ?? '').trim();
    const label = String(status || '').trim();
    if(!key || !label) return;

    const list = map.get(key) || [];
    if(!list.includes(label)) list.push(label);
    map.set(key,list);
  }

  /* =========================================================
     21-3. 每日人員狀態解析
     ---------------------------------------------------------
     勤務表右側會列出輪休、請休、補休、公假、休假役男、請替等狀態。
     狀態欄的實際排版不是固定欄數，因此以「狀態文字」作為錨點，
     再讀取同格及右側直到下一個狀態錨點前的番號。
  ========================================================= */
  function detectDailyStatuses(matrix){
    const settings = (appConfig && appConfig.importSettings) || {};
    const keywords = settings.personStatusKeywords || [
      '輪休','請休','補休','公假','連續補休','休假役男','請替'
    ];
    const normalized = keywords.map(item=>({raw:item,clean:cleanText(item)}));
    const result = new Map();

    for(let r=0;r<matrix.length;r++){
      const row = matrix[r] || [];

      for(let c=0;c<row.length;c++){
        const cellText = cleanText(row[c]);
        if(!cellText) continue;

        const matched = normalized.find(item=>cellText === item.clean || cellText.startsWith(item.clean));
        if(!matched) continue;

        const values = [row[c] ?? ''];

        // 實際版型通常是「狀態 | 空格 | 番號」，最多向右抓 5 格；
        // 遇到下一個狀態名稱立即停止，避免把別組狀態混在一起。
        for(let c2=c+1;c2<row.length && c2<=c+5;c2++){
          const nextText = cleanText(row[c2]);
          const isNextStatus = normalized.some(item=>
            nextText === item.clean || nextText.startsWith(item.clean)
          );
          if(isNextStatus) break;
          values.push(row[c2] ?? '');
        }

        parseNumberList(values).forEach(no=>addStatusForNumber(result,no,matched.raw));
      }
    }

    return result;
  }

  function detectDutySchedule(sheet,matrix){
    const settings = (appConfig && appConfig.importSettings) || {};
    const key91 = settings.duty91Keywords || ['備勤91','備勤(91)','備勤（91）'];
    const key92 = settings.duty92Keywords || ['備勤救災','備勤(救災)','備勤（救災）'];
    const keyWatch = settings.dutyWatchKeywords || ['值班'];
    const keyAtStation = settings.atStationKeywords || ['在隊備勤'];
    const keyRest = settings.restKeywords || ['休息時間'];
    const keyDutyEnd = settings.dutyRegionEndKeywords || ['服勤編組'];

    const header91 = findKeywordCell(matrix,key91);
    const header92 = findKeywordCell(matrix,key92);
    // 「值班」與「值班指導員」是兩個不同欄位。
    // 火警值班只能讀取欄名完全等於「值班」的那一欄，不能用模糊比對抓到「值班指導員」。
    const headerWatch = findExactKeywordCell(matrix,keyWatch);
    const headerAtStation = findKeywordCell(matrix,keyAtStation);
    const headerRest = findKeywordCell(matrix,keyRest);
    const headerDutyEnd = findKeywordCell(matrix,keyDutyEnd);
    const timeCol = detectTimeColumn(matrix);

    if(!header91 || !header92 || !headerAtStation || !headerRest || timeCol === null){
      return {
        schedule:[],
        header91,
        header92,
        headerWatch,
        headerAtStation,
        headerRest,
        timeCol,
        error:'找不到「時間／備勤91／備勤救災／在隊備勤／休息時間」欄位'
      };
    }

    const span91 = getMergeSpan(sheet,header91.row,header91.col);
    const span92 = getMergeSpan(sheet,header92.row,header92.col);
    const spanWatch = headerWatch ? getMergeSpan(sheet,headerWatch.row,headerWatch.col) : null;
    const spanAtStation = getMergeSpan(sheet,headerAtStation.row,headerAtStation.col);
    const spanRest = getMergeSpan(sheet,headerRest.row,headerRest.col);

    // 勤務區從時間欄右側開始，到「服勤編組」前一欄為止。
    // 「服勤編組」本身是人數（例如 10、11），不是人員番號；不可納入 allDutyNumbers。
    // 若版型沒有服勤編組欄，至少涵蓋目前已辨識的 91 / 92 / 在隊備勤 / 休息時間。
    const dutyStartCol = timeCol + 1;
    const fallbackDutyEnd = Math.max(
      span91.endCol,
      span92.endCol,
      ...(spanWatch ? [spanWatch.endCol] : []),
      spanAtStation.endCol,
      spanRest.endCol
    );
    const dutyEndCol = headerDutyEnd
      ? Math.max(dutyStartCol,getMergeSpan(sheet,headerDutyEnd.row,headerDutyEnd.col).startCol - 1)
      : fallbackDutyEnd;

    const currentByCol = new Map();
    const schedule = [];

    // 休息時間通常是連續兩個 1 小時時段，例如 10-11 + 11-12。
    // Excel 實務上常只在第一個時段寫番號，下一列留白，所以不能單純把空白視為「休息結束」。
    // 規則：休息番號最多向後延續 1 個時段；若下一時段該番號已明確出現在其他勤務，則立即停止，
    // 因此同時支援常見的 2 小時休息，以及只有 1 小時後就回到其他勤務的情況。
    let restCarryNumbers = [];
    let restCarryRemaining = 0;

    for(let r=0;r<matrix.length;r++){
      const row = matrix[r] || [];
      const period = parseTimeRange(row[timeCol]);
      if(!period) continue;

      const parsedByCol = new Map();
      let rowHasDutyNumbers = false;

      for(let c=dutyStartCol;c<=dutyEndCol;c++){
        const nums = parseNumberList(mergedCellValue(sheet,matrix,r,c));
        parsedByCol.set(c,nums);
        if(nums.length) rowHasDutyNumbers = true;
      }

      /*
         一般勤務維持原本的時段沿用：
         - 本列有勤務番號：以本列為新的狀態。
         - 整列無勤務番號：沿用上一列。
      */
      if(rowHasDutyNumbers || currentByCol.size === 0){
        for(let c=dutyStartCol;c<=dutyEndCol;c++){
          currentByCol.set(c,[...(parsedByCol.get(c) || [])]);
        }
      }

      // 休息欄位獨立處理，不讓其他勤務欄位的更新誤把第二個休息時段清掉。
      const explicitRest=[];
      for(let c=spanRest.startCol;c<=spanRest.endCol;c++){
        // 只看這一列實際填寫的儲存格，不把 mergedCellValue 的延伸值再當成新的起點，
        // 否則兩列合併的休息會被錯誤延長成第三個時段。
        explicitRest.push(...parseNumberList(row[c] ?? ''));
      }
      const uniqueExplicitRest=[...new Set(explicitRest.map(no=>String(no)))];

      let effectiveRest=[];
      if(uniqueExplicitRest.length){
        effectiveRest=[...uniqueExplicitRest];
        restCarryNumbers=[...uniqueExplicitRest];
        restCarryRemaining=1;
      }else if(restCarryRemaining>0 && restCarryNumbers.length){
        // 下一時段若已明確排到其他勤務，該番號就不再延續休息。
        const explicitOtherDuty=new Set();
        for(let c=dutyStartCol;c<=dutyEndCol;c++){
          if(c>=spanRest.startCol && c<=spanRest.endCol) continue;
          (parsedByCol.get(c) || []).forEach(no=>explicitOtherDuty.add(String(no)));
        }
        effectiveRest=restCarryNumbers.filter(no=>!explicitOtherDuty.has(String(no)));
        restCarryNumbers=[...effectiveRest];
        restCarryRemaining=0;
      }else{
        effectiveRest=[];
        restCarryNumbers=[];
        restCarryRemaining=0;
      }

      // 寫回統一勤務狀態，後面的 allDutyNumbers / 休息區都使用同一份結果。
      for(let c=spanRest.startCol;c<=spanRest.endCol;c++) currentByCol.set(c,[]);
      currentByCol.set(spanRest.startCol,[...effectiveRest]);

      const nums91 = numbersFromColumnState(currentByCol,span91.startCol,span91.endCol);
      const nums92 = numbersFromColumnState(currentByCol,span92.startCol,span92.endCol);
      const numsWatch = spanWatch ? numbersFromColumnState(currentByCol,spanWatch.startCol,spanWatch.endCol) : [];
      const numsAtStation = numbersFromColumnState(currentByCol,spanAtStation.startCol,spanAtStation.endCol);
      const numsRest = numbersFromColumnState(currentByCol,spanRest.startCol,spanRest.endCol);
      const allDutyNumbers = numbersFromColumnState(currentByCol,dutyStartCol,dutyEndCol);

      schedule.push({
        start:period.start,
        end:period.end,
        '備勤91':nums91,
        '備勤救災':nums92,
        '值班':numsWatch,
        '在隊備勤':numsAtStation,
        '休息時間':numsRest,
        allDutyNumbers
      });
    }

    return {
      schedule,
      header91,
      header92,
      headerWatch,
      headerAtStation,
      headerRest,
      headerDutyEnd,
      timeCol,
      dutyStartCol,
      dutyEndCol,
      error:''
    };
  }


  /* =========================================================
     Excel 內建火警出動基礎配置
     ---------------------------------------------------------
     直接讀取每日 Excel 右上角「火警出動人員車輛分配表」。
     這份配置是當日基礎，不再只依賴 board-data.json 的固定範例。
     人員是否實際放入仍由目前 91 / 92、休息、請假狀態決定。
  ========================================================= */
  function detectFireBoardBase(matrix,roster){
    const title=findExactKeywordCell(matrix,['火警出動人員車輛分配表']);
    if(!title) return [];

    let headerRow=-1;
    for(let r=title.row;r<Math.min(matrix.length,title.row+5);r++){
      const row=matrix[r] || [];
      if(row.some(v=>cleanText(v)==='梯次') && row.some(v=>cleanText(v)==='車輛')){
        headerRow=r; break;
      }
    }
    if(headerRow<0) return [];

    const row=matrix[headerRow] || [];
    const exactCol=labels=>{
      const list=Array.isArray(labels)?labels:[labels];
      for(let c=title.col;c<row.length;c++){
        if(list.includes(cleanText(row[c]))) return c;
      }
      return -1;
    };
    const cols={
      leader:exactCol('帶隊官'),
      vehicle:exactCol('車輛'),
      driver:exactCol('駕駛員'),
      nozzle:exactCol('瞄子手'),
      assistantNozzle:exactCol('副瞄子手'),
      searchPhoto:exactCol(['搜救小組、照相','搜救小組/照相','搜救小組／照相'])
    };
    if(cols.vehicle<0) return [];

    const normalizeVehicle=value=>{
      const text=String(value ?? '').replace(/\s+/g,'').trim();
      if(!text) return '';
      if(text==='中指車') return '中隊指揮車';
      return text.replace(/^(\d+)\([^)]*\)車$/,'$1車');
    };
    const personFromValue=value=>{
      const raw=String(value ?? '').replace(/\r?\n/g,' ').trim();
      if(!raw) return null;
      const parsed=parseRoleAndName(raw);
      const name=(parsed.name || raw.replace(/^(中隊長|分隊長|小隊長|隊員|役男|義消)\s*/,'')).replace(/\s+/g,'').trim();
      if(!name || /火警值班|值班人員/.test(name)) return null;
      const found=roster.find(item=>String(item.name || '').replace(/\s+/g,'')===name);
      return found ? {no:found.no,name:found.name,role:found.role} : {no:'',name,role:parsed.role || '隊員'};
    };

    const assignments=[];
    const fields=[
      ['leader',cols.leader],
      ['vehicle',cols.vehicle],
      ['driver',cols.driver],
      ['nozzle',cols.nozzle],
      ['assistant-nozzle',cols.assistantNozzle]
    ];

    for(let offset=1;offset<=6;offset++){
      const r=headerRow+offset;
      const source=matrix[r] || [];
      const echelon=offset<=3?'first':'second';
      const index=((offset-1)%3)+1;
      const rowVehicle=normalizeVehicle(source[cols.vehicle]);

      // Excel 第二梯次中的 91 / 92 是專責救護基礎資料。
      // 網頁已有獨立「專責救護」91 / 92 區，這兩列不可再複製進火警主表。
      if(echelon==='second' && (rowVehicle==='91車' || rowVehicle==='92車')) continue;

      fields.forEach(([slot,col])=>{
        if(col<0) return;
        const raw=String(source[col] ?? '').trim();
        if(!raw) return;
        const slotId=`${echelon}-${index}-${slot}`;

        if(slot==='vehicle'){
          const value=normalizeVehicle(raw);
          if(value && value!=='91車' && value!=='92車') assignments.push({slotId,type:'vehicle',value});
          return;
        }

        // 火警值班在這份 Excel 位於第二梯次最後一列。
        // 對應網頁「第 8 欄、第 7 行」＝攝影照相格；右邊破壞小組格放目前「值班」番號。
        if(cleanText(raw)==='火警值班'){
          const fireDutySlotId=(echelon==='second' && index===3) ? 'second-3-photo' : slotId;
          if(!assignments.some(x=>x.type==='duty' && x.value==='火警值班')){
            assignments.push({slotId:fireDutySlotId,type:'duty',value:'火警值班'});
          }
          return;
        }

        if(cleanText(raw)==='值班人員') return;
        const person=personFromValue(raw);
        if(person) assignments.push({slotId,type:'person',...person});
      });

      // 範例 Excel 的「值班人員」位於火警值班右側欄位；它只是提示字，
      // 網頁右側實際人員由目前勤務時段的「值班」欄自動帶入，所以不建立固定人員。
      if(cols.searchPhoto>=0 && cleanText(source[cols.searchPhoto])==='值班人員'){
        // no-op
      }
    }

    // 特別確認 Excel 第二梯次最後一列是否含「火警值班」。
    // 不依它在 Excel 的實體欄位置硬映射；網頁固定落在第8欄第7行(second-3-photo)。
    const lastRow=matrix[headerRow+6] || [];
    if(lastRow.some(v=>cleanText(v)==='火警值班') && !assignments.some(x=>x.type==='duty' && x.value==='火警值班')){
      assignments.push({slotId:'second-3-photo',type:'duty',value:'火警值班'});
    }

    return assignments;
  }

  function collectMissingNumbers(roster,schedule){
    const known = new Set(roster.map(item=>Number(item.no)));
    const missing = new Set();

    schedule.forEach(period=>{
      (period.allDutyNumbers || []).forEach(no=>{
        if(!known.has(Number(no))){
          missing.add(Number(no));
        }
      });
    });

    return [...missing].sort((a,b)=>a-b);
  }

  function previewImport(result){
    const missingText = result.missing.length
      ? `<span class="text-danger fw-bold">找不到姓名的番號：${result.missing.join('、')}</span>`
      : '<span class="text-success fw-bold">91／92 使用到的番號皆可對應姓名</span>';

    $('#importDetectStatus').html(`
      <div class="row g-2">
        <div class="col-md-3"><div class="border rounded p-2"><b>工作表</b><br>${escapeHtml(result.sheetName)}</div></div>
        <div class="col-md-3"><div class="border rounded p-2"><b>人員</b><br>${result.roster.length} 人</div></div>
        <div class="col-md-3"><div class="border rounded p-2"><b>勤務時段</b><br>${result.schedule.length} 段<br><small>狀態 ${result.statuses instanceof Map ? result.statuses.size : 0} 人</small></div></div>
        <div class="col-md-3"><div class="border rounded p-2"><b>火警基礎配置</b><br>${Array.isArray(result.baseAssignments) ? result.baseAssignments.length : 0} 格</div></div>
      </div>
      <div class="mt-2">${missingText}</div>
    `);

    let html = `
      <table class="table table-sm table-bordered align-middle mb-0">
        <thead>
          <tr>
            <th>時間</th>
            <th>備勤91 → 91車</th>
            <th>備勤救災 → 92車</th>
            <th>值班</th>
            <th>在隊備勤</th>
            <th>休息時間</th>
          </tr>
        </thead>
        <tbody>
    `;

    result.schedule.slice(0,12).forEach(period=>{
      const names91 = (period['備勤91'] || []).map(no=>{
        const p = result.roster.find(x=>Number(x.no) === Number(no));
        return p ? `${no} ${p.name}` : `${no}`;
      }).join('、');

      const names92 = (period['備勤救災'] || []).map(no=>{
        const p = result.roster.find(x=>Number(x.no) === Number(no));
        return p ? `${no} ${p.name}` : `${no}`;
      }).join('、');

      const namesWatch = (period['值班'] || []).map(no=>{
        const p = result.roster.find(x=>Number(x.no) === Number(no));
        return p ? `${no} ${p.name}` : `${no}`;
      }).join('、');

      const namesAtStation = (period['在隊備勤'] || []).map(no=>{
        const p = result.roster.find(x=>Number(x.no) === Number(no));
        return p ? `${no} ${p.name}` : `${no}`;
      }).join('、');

      const namesRest = (period['休息時間'] || []).map(no=>{
        const p = result.roster.find(x=>Number(x.no) === Number(no));
        return p ? `${no} ${p.name}` : `${no}`;
      }).join('、');

      html += `
        <tr>
          <td>${escapeHtml(period.start)}–${escapeHtml(period.end)}</td>
          <td>${escapeHtml(names91)}</td>
          <td>${escapeHtml(names92)}</td>
          <td>${escapeHtml(namesWatch)}</td>
          <td>${escapeHtml(namesAtStation)}</td>
          <td>${escapeHtml(namesRest)}</td>
        </tr>
      `;
    });

    html += '</tbody></table>';
    $('#previewWrap').html(html);
  }

  /* =========================================================
     22. Excel 預覽輸出安全處理
  ========================================================= */
  function escapeHtml(value){
    return $('<div>').text(String(value ?? '')).html();
  }

  /* =========================================================
     23. Excel 檔案讀取與自動偵測最佳工作表
  ========================================================= */
  $('#excelFile').on('change',async function(){
    const file = this.files && this.files[0];
    if(!file) return;

    pendingImport = null;
    $('#confirmImport').prop('disabled',true);
    $('#previewWrap').empty();
    $('#importDetectStatus').text('正在解析勤務表…');

    try{
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer,{type:'array',cellDates:true});

      let best = null;

      workbook.SheetNames.forEach(sheetName=>{
        const sheet = workbook.Sheets[sheetName];
        const matrix = sheetMatrix(sheet);
        const roster = detectPersonnel(matrix);
        const duty = detectDutySchedule(sheet,matrix);
        const statuses = detectDailyStatuses(matrix);
        const baseAssignments = detectFireBoardBase(matrix,roster);

        const score = roster.length * 3 + duty.schedule.length * 5 + statuses.size;

        if(!best || score > best.score){
          best = {
            score,
            sheetName,
            roster,
            schedule:duty.schedule,
            statuses,
            baseAssignments,
            error:duty.error
          };
        }
      });

      if(!best || !best.roster.length || !best.schedule.length){
        $('#importDetectStatus').html(
          '<span class="text-danger fw-bold">無法從這份 Excel 辨識人員名冊或勤務時段。</span>'
        );
        $('#previewWrap').html(
          '<div class="p-3 small">目前解析器是依你提供的勤務表圖片版型設計；取得實際 Excel 後，如果儲存格結構不同，再調整一次欄位偵測即可。</div>'
        );
        importModal.show();
        this.value = '';
        return;
      }

      best.missing = collectMissingNumbers(best.roster,best.schedule);
      pendingImport = best;

      previewImport(best);
      $('#confirmImport').prop('disabled',false);
      importModal.show();

    }catch(err){
      console.error(err);
      $('#importDetectStatus').html('<span class="text-danger fw-bold">Excel 讀取失敗</span>');
      $('#previewWrap').empty();
      importModal.show();
    }

    this.value = '';
  });

  /* =========================================================
     24. 確認匯入：立即成為最新看板並自動同步
  ========================================================= */
  $('#confirmImport').on('click',function(){
    if(!pendingImport) return;

    todayRoster = pendingImport.roster.map(item=>({
      no:item.no,
      name:item.name,
      role:normalizeRole(item.role),
      start:'',
      end:''
    }));

    hasDetailedDutyData = true;
    importedBaseAssignments = Array.isArray(pendingImport.baseAssignments) && pendingImport.baseAssignments.length
      ? pendingImport.baseAssignments.map(item=>({...item}))
      : (Array.isArray(appConfig?.baseAssignments) ? appConfig.baseAssignments.map(item=>({...item})) : []);
    dutyStatusByNo = new Map();
    if(pendingImport.statuses instanceof Map){
      pendingImport.statuses.forEach((statuses,no)=>{
        dutyStatusByNo.set(String(no),[...(statuses || [])]);
      });
    }

    dutySchedule = pendingImport.schedule.map(item=>({
      start:item.start,
      end:item.end,
      '備勤91':[...(item['備勤91'] || [])],
      '備勤救災':[...(item['備勤救災'] || [])],
      '值班':[...(item['值班'] || [])],
      '在隊備勤':[...(item['在隊備勤'] || [])],
      '休息時間':[...(item['休息時間'] || [])],
      allDutyNumbers:[...(item.allDutyNumbers || [])]
    }));

    dutyPeriodOverrideKey = '';
    currentDutyKey = '';
    updateDutyPeriodEditorUi();

    // 確認匯入後，這份畫面就是目前最新版本；不需要再按儲存。
    // 先套用基礎火警配置，再依當前 91 / 92、休息與請假狀態覆蓋，最後立即寫入 Google Sheet。
    clearBoardAssignments();
    renderVehiclePool();
    syncAll();
    refreshDutySchedule(true);

    importModal.hide();

    const missingText = pendingImport.missing.length
      ? `；另有 ${pendingImport.missing.length} 個番號尚未對應姓名`
      : '';

    toast(`已匯入 ${todayRoster.length} 位人員、${dutySchedule.length} 個勤務時段${missingText}`);

    // 匯入完成後自動同步，不需要額外按任何儲存按鈕。
    queueAutoSave('import',50);

    pendingImport = null;
  });

  /* =========================================================
     25. 程式進入點
     ---------------------------------------------------------
     所有 UI / 拖曳事件先建立，再讀設定與 Google Sheet。
  ========================================================= */
  // 所有畫面與拖曳事件建立完成後，再載入 JSON。
  loadBoardData();
  enableAutoFullscreen();
  scheduleFitBoardText();

});
