<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no, viewport-fit=cover">
    <title>沐沐的週邊✨</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.0/Sortable.min.js"></script>
    <style>
        :root { --primary: #6c5ce7; --primary-light: #efeeff; --bg: #f8f9ff; --gray-text: #636e72; --dash-color: #d1d1f0; --deep-gray: #4a4a4a; }
        * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
        body { font-family: -apple-system, sans-serif; background: var(--bg); margin: 0; padding: 12px; padding-bottom: 80px; }
        
        body.modal-open { overflow: hidden; position: fixed; width: 100%; }

        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
        .header h2 { margin: 0; color: var(--primary); font-size: 22px; font-weight: 800; cursor: pointer; }
        .settings-icon { font-size: 26px; cursor: pointer; color: #ccc; padding: 5px; }
        
        .card { background: white; padding: 12px; border-radius: 12px; box-shadow: 0 3px 10px rgba(0,0,0,0.06); margin-bottom: 12px; width: 100%; }
        .order-item { border-left: 5px solid var(--primary); padding: 12px; position: relative; }
        .order-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
        .title-text { font-size: 17px; font-weight: 800; color: #333; flex: 1; margin-left: 6px; }
        
        .tag { padding: 3px 8px; border-radius: 6px; font-size: 13px; margin-right: 4px; font-weight: 600; color: #444; }
        .tag-yellow { background: #fff9db; }
        .tag-pink { background: #ffe3e3; }
        .tag-blue { background: #e7f5ff; }
        .tag-green { background: #ebfbee; }
        .tag-default { background: #f1f2f6; color: #636e72; }

        .summary-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; margin-top: 10px; padding-top: 10px; border-top: 1px dashed var(--dash-color); }
        .summary-cell { font-size: 12px; font-weight: 700; color: #444; }
        .align-right { text-align: right; }
        .text-purple { color: #6c5ce7; }
        .text-red { color: #ff7675; }

        input, select { height: 38px; width: 100%; padding: 6px; margin: 3px 0; border: 1px solid #e0e0e0; border-radius: 8px; font-size: 14px; outline: none; }
        .row { display: flex; gap: 6px; align-items: flex-end; width: 100%; margin-bottom: 4px; }
        button { border: none; border-radius: 8px; font-weight: bold; cursor: pointer; }
        .btn-sub { background: var(--primary-light); color: var(--primary); width: 100%; padding: 10px; font-size: 13px; }
        .btn-main { background: var(--primary); color: white; width: 100%; padding: 14px; border-radius: 12px; font-size: 16px; margin-top: 8px; font-weight: 800; }
        
        #toast { display: none; position: fixed; top: 20px; left: 50%; transform: translateX(-50%); background: rgba(0,0,0,0.8); color: white; padding: 8px 20px; border-radius: 20px; z-index: 11000; }

        .modal { 
            display: none; 
            position: fixed; 
            top: 0; left: 0; right: 0; bottom: 0; 
            width: 100vw; height: 100vh; 
            background: #f8f9ff !important;
            z-index: 9999; 
            overflow-y: auto; 
            padding: 20px; 
            -webkit-overflow-scrolling: touch;
        }

        .item-input-group { background: white; border: 1px dashed var(--dash-color); padding: 12px; border-radius: 10px; margin-bottom: 12px; position: relative; }
        .btn-del-sub { position: absolute; right: -8px; top: -8px; background: #ff7675; color: white; width: 24px; height: 24px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 2px solid white; z-index: 5; }
        .btn-drag-sub { position: absolute; left: -8px; top: -8px; background: #eee; color: #888; width: 24px; height: 24px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 2px solid white; z-index: 5; }

        .tab-container { display: flex; gap: 8px; margin: 10px 0; }
        .tab { flex: 1; text-align: center; padding: 12px; background: #e9ecef; border-radius: 10px; color: #6c757d; font-size: 14px; font-weight: bold; }
        .tab.active { background: var(--primary); color: white; }

        .section-header { font-size: 16px; font-weight: 800; padding: 12px; background: white; border-radius: 10px; margin-top: 10px; display: flex; justify-content: space-between; }
        .collapse-content { display: none; padding: 10px 0; }
        
        .btn-export { background: #ffeaa7; color: #d6a312; }
        .btn-import { background: #d1f7d1; color: #2d8a2d; margin-top: 10px; }
    </style>
</head>
<body>

<datalist id="list_sellers"></datalist>
<datalist id="list_series_all"></datalist>
<div id="toast">✅ 操作成功！</div>
<input type="file" id="importExcelFile" style="display:none;" accept=".xlsx, .xls" onchange="processExcelImport(event)">

<div id="main-container">
    <div class="header">
        <h2 id="mainTitle" contenteditable="true" onblur="saveTitle()">📦 沐沐的週邊✨</h2>
        <span class="settings-icon" onclick="toggleModal('settingsPage', true)">⚙️</span>
    </div>

    <div class="card" id="mainInputForm">
        <div class="row"><select id="sel_platform"></select><input type="text" id="seller" placeholder="賣家" list="list_sellers"></div>
        <div class="row"><select id="sel_status" onchange="calcPickup('mainInputForm')">
            <option>已匯款/未下單</option><option>已匯款/已下單</option>
            <option>貨到付款/未下單</option><option>貨到付款/已下單</option>
        </select><input type="text" id="order_note" placeholder="備註"></div>
        <div class="row">
            <input type="number" id="unpaid_fee" placeholder="未付" oninput="calcPickup('mainInputForm')">
            <input type="number" id="shipping" placeholder="運費" oninput="calcPickup('mainInputForm')">
            <input type="number" id="packing_fee" placeholder="包手" oninput="calcPickup('mainInputForm')">
            <input type="number" id="pickup_total" placeholder="取貨" readonly style="background:#eee">
        </div>
        <div class="subContainer" id="mainSubContainer" style="margin-top:10px;"></div>
        <button class="btn-sub" onclick="addSubItem('mainInputForm')">＋增加項目</button>
        <button class="btn-main" onclick="handleSave('mainInputForm')">💾 儲存整筆訂單</button>
    </div>

    <div class="tab-container">
        <div class="tab active" id="tab-unarrived" onclick="switchTab(false)">🎁 未到貨</div>
        <div class="tab" id="tab-arrived" onclick="switchTab(true)">🧸 已到貨</div>
    </div>
    <div id="mainList"></div>
</div>

<div id="editModal" class="modal">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
        <span onclick="toggleModal('editModal', false)" style="font-size:28px;">✖</span>
        <h3 style="margin:0;">編輯訂單</h3>
        <span onclick="handleSave('editModal')" style="font-size:28px;">💾</span>
    </div>
    <div id="editForm">
        <div class="row"><select id="edit_platform"></select><input type="text" id="edit_seller" list="list_sellers"></div>
        <div class="row"><select id="edit_status" onchange="calcPickup('editModal')">
            <option>已匯款/未下單</option><option>已匯款/已下單</option>
            <option>貨到付款/未下單</option><option>貨到付款/已下單</option>
        </select><input type="text" id="edit_note"></div>
        <div class="row">
            <input type="number" id="edit_unpaid" oninput="calcPickup('editModal')">
            <input type="number" id="edit_shipping" oninput="calcPickup('editModal')">
            <input type="number" id="edit_packing" oninput="calcPickup('editModal')">
            <input type="number" id="edit_pickup" readonly style="background:#eee">
        </div>
        <div class="row">
            <div style="flex:1"><small>🕒 建立時間</small><input type="datetime-local" id="edit_time_picker"></div>
            <div style="flex:1"><small>✅ 到貨時間</small><input type="datetime-local" id="edit_arrive_picker"></div>
        </div>
        <div class="subContainer" id="editSubContainer" style="margin-top:15px;"></div>
        <button class="btn-sub" onclick="addSubItem('editModal')">＋增加項目</button>
        <div style="height: 150px;"></div>
    </div>
</div>

<div id="settingsPage" class="modal">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
        <h2 style="margin:0;">⚙️ 設定</h2>
        <span onclick="toggleModal('settingsPage', false)" style="font-size:32px;">✖</span>
    </div>
    <div class="section-header" onclick="toggleCollapse('optSection', this)"><span>📋 選項管理</span><span>⬇️</span></div>
    <div id="optSection" class="collapse-content">
        <input type="text" id="new_opt_val" placeholder="新增名稱..." style="margin-bottom:10px;">
        <div class="row" style="flex-wrap:wrap">
            <button class="btn-sub" style="flex:1; min-width:80px" onclick="addOption('platform')">加平台</button>
            <button class="btn-sub" style="flex:1; min-width:80px" onclick="addOption('series')">加系列</button>
            <button class="btn-sub" style="flex:1; min-width:80px" onclick="addOption('char')">加人物</button>
            <button class="btn-sub" style="flex:1; min-width:80px" onclick="addOption('type')">加類別</button>
        </div>
        <div id="configDisplay" style="margin-top:15px;"></div>
    </div>
    <button class="btn-main btn-export" onclick="exportToExcel()">📥 匯出 Excel 備份</button>
    <button class="btn-main btn-import" onclick="document.getElementById('importExcelFile').click()">📤 匯入 Excel 還原</button>
</div>

<script>
let orders = JSON.parse(localStorage.getItem('MU_DATA_V9') || '[]');
let config = JSON.parse(localStorage.getItem('m_conf_v9') || '{"platform":[],"series":[],"char":[],"type":[]}');
let editingId = null;
let currentTabArrived = false;

function saveTitle() { localStorage.setItem('MU_APP_TITLE', document.getElementById('mainTitle').innerText); }
function initTitle() { document.getElementById('mainTitle').innerText = localStorage.getItem('MU_APP_TITLE') || '📦 沐沐的週邊✨'; }

function toggleModal(id, show) {
    const m = document.getElementById(id); m.style.display = show ? 'block' : 'none';
    if(show) { document.body.classList.add('modal-open'); m.scrollTo(0,0); }
    else { document.body.classList.remove('modal-open'); }
}

function addSubItem(cid, data=null, anchorEl=null) {
    const sc = document.querySelector(`#${cid} .subContainer`);
    const div = document.createElement('div'); div.className = 'item-input-group';
    div.innerHTML = `<div class="btn-drag-sub" onclick="copySubItem(this, '${cid}')">📋</div><div class="btn-del-sub" onclick="removeSubItem(this, '${cid}')">✖</div>
        <div class="row"><input type="text" class="p-series" placeholder="系列" list="list_series_all"><select class="p-char"></select><select class="p-type"></select></div>
        <div class="row" style="margin-top:8px;"><input type="text" class="p-name" placeholder="名稱" style="flex:2"><input type="number" class="p-price" placeholder="單價" style="flex:1" oninput="calcPickup('${cid}')"><input type="number" class="p-qty" style="flex:0.8" value="1" oninput="calcPickup('${cid}')"></div>`;
    if(anchorEl) anchorEl.insertAdjacentElement('afterend', div); else sc.appendChild(div);
    const setOpt = (el, k) => { el.innerHTML = `<option value="">${k}</option>` + config[k==='人物'?'char':'type'].map(x=>`<option value="${x}">${x}</option>`).join(''); };
    setOpt(div.querySelector('.p-char'), '人物'); setOpt(div.querySelector('.p-type'), '類別');
    if(data) { div.querySelector('.p-series').value=data.s; div.querySelector('.p-char').value=data.c; div.querySelector('.p-type').value=data.t; div.querySelector('.p-name').value=data.n; div.querySelector('.p-price').value=data.p; div.querySelector('.p-qty').value=data.q; }
    calcPickup(cid);
}

function copySubItem(btn, cid) {
    const p = btn.parentElement;
    const data = { s: p.querySelector('.p-series').value, c: p.querySelector('.p-char').value, t: p.querySelector('.p-type').value, n: p.querySelector('.p-name').value, p: p.querySelector('.p-price').value, q: p.querySelector('.p-qty').value };
    addSubItem(cid, data, p);
}

function removeSubItem(btn, cid) { if(document.querySelectorAll(`#${cid} .item-input-group`).length > 1) { btn.parentElement.remove(); calcPickup(cid); } }

function calcPickup(cid) { 
    const box = document.getElementById(cid); const isE = cid === 'editModal';
    const status = box.querySelector(isE ? '#edit_status' : '#sel_status').value;
    let itemSum = 0;
    box.querySelectorAll('.item-input-group').forEach(r => { itemSum += (Number(r.querySelector('.p-price').value) || 0) * (Number(r.querySelector('.p-qty').value) || 1); });
    const unpaidInput = box.querySelector(isE ? '#edit_unpaid' : '#unpaid_fee');
    if (status.includes('貨到付款')) unpaidInput.value = itemSum;
    const u = Number(unpaidInput.value)||0, s = Number(box.querySelector(isE?'#edit_shipping':'#shipping')?.value)||0, p = Number(box.querySelector(isE?'#edit_packing':'#packing_fee')?.value)||0;
    box.querySelector(isE ? '#edit_pickup' : '#pickup_total').value = u + s + p;
}

function handleSave(cid) {
    const isE = cid === 'editModal'; const box = document.getElementById(cid);
    const items = []; let iTotal = 0;
    box.querySelectorAll('.item-input-group').forEach(r => {
        const s = r.querySelector('.p-series').value.trim(); if(s && !config.series.includes(s)) config.series.push(s);
        const p = Number(r.querySelector('.p-price').value)||0, q = Number(r.querySelector('.p-qty').value)||1;
        items.push({ s, c: r.querySelector('.p-char').value, t: r.querySelector('.p-type').value, n: r.querySelector('.p-name').value.trim(), p, q });
        iTotal += p * q;
    });
    localStorage.setItem('m_conf_v9', JSON.stringify(config)); updateDropdowns();
    const unpaid = Number(box.querySelector(isE ? '#edit_unpaid' : '#unpaid_fee').value)||0, shipping = Number(box.querySelector(isE ? '#edit_shipping' : '#shipping').value)||0, packing = Number(box.querySelector(isE ? '#edit_packing' : '#packing_fee').value)||0;
    const data = {
        id: isE ? editingId : Date.now(), platform: box.querySelector(isE ? '#edit_platform' : '#sel_platform').value, seller: box.querySelector(isE ? '#edit_seller' : '#seller').value.trim(), status: box.querySelector(isE ? '#edit_status' : '#sel_status').value, note: box.querySelector(isE ? '#edit_note' : '#order_note').value.trim(),
        unpaid, shipping, packing, pickup: unpaid + shipping + packing, itemTotal: iTotal, items, time: isE ? new Date(box.querySelector('#edit_time_picker').value).toISOString() : new Date().toISOString(), arrived: isE ? !!box.querySelector('#edit_arrive_picker').value : false, arriveTime: isE && box.querySelector('#edit_arrive_picker').value ? new Date(box.querySelector('#edit_arrive_picker').value).toISOString() : null
    };
    if(isE) { orders[orders.findIndex(o => o.id === editingId)] = data; toggleModal('editModal', false); }
    else { orders.unshift(data); clearMainForm(); }
    saveToLocal(); renderOrders(currentTabArrived); showToast();
}

function clearMainForm() { ['seller', 'order_note', 'unpaid_fee', 'shipping', 'packing_fee', 'pickup_total'].forEach(id => document.getElementById(id).value = ''); document.querySelector('#mainSubContainer').innerHTML = ''; addSubItem('mainInputForm'); }

function renderOrders(isArrived) {
    currentTabArrived = isArrived; let filtered = orders.filter(o => o.arrived === isArrived);
    const list = document.getElementById('mainList');
    if(!filtered.length) { list.innerHTML = `<div style="text-align:center;color:#ccc;margin-top:50px;">尚無訂單</div>`; return; }
    list.innerHTML = filtered.map(o => `
        <div class="card order-item">
            <div class="order-header">
                <input type="checkbox" style="width:20px;height:20px;" ${o.arrived?'checked':''} onclick="toggleArrive(${o.id})">
                <div class="title-text">${o.platform} - ${o.seller}</div>
                <div style="font-size:13px;"><span class="btn-action-text" onclick="editOrder(${o.id})">編輯</span> | <span class="btn-action-text" onclick="del(${o.id})">刪除</span></div>
            </div>
            <div style="margin:5px 0;"><span class="tag ${getTagClass(o.status)}">${o.status}</span></div>
            <div style="font-size:14px;color:#444;">${o.items.map(i=>`<div>· [${i.s}] ${i.c} ${i.n} $${i.p}x${i.q}</div>`).join('')}</div>
            <div class="summary-grid"><div class="summary-cell text-purple">取貨：$${o.pickup}</div><div class="summary-cell align-right text-red">總計：$${o.itemTotal - o.unpaid + o.pickup}</div></div>
        </div>`).join('');
}

function editOrder(id) {
    const o = orders.find(x => x.id === id); editingId = id; toggleModal('editModal', true);
    document.getElementById('edit_platform').value = o.platform; document.getElementById('edit_seller').value = o.seller; document.getElementById('edit_status').value = o.status; document.getElementById('edit_note').value = o.note; document.getElementById('edit_unpaid').value = o.unpaid; document.getElementById('edit_shipping').value = o.shipping; document.getElementById('edit_packing').value = o.packing; document.getElementById('edit_pickup').value = o.pickup;
    document.getElementById('edit_time_picker').value = o.time.slice(0,16); document.getElementById('edit_arrive_picker').value = o.arriveTime ? o.arriveTime.slice(0,16) : '';
    const sc = document.getElementById('editSubContainer'); sc.innerHTML = ''; o.items.forEach(i => addSubItem('editModal', i));
}

function toggleArrive(id) { const o = orders.find(x => x.id === id); o.arrived = !o.arrived; o.arriveTime = o.arrived ? new Date().toISOString() : null; saveToLocal(); renderOrders(currentTabArrived); }
function switchTab(ar) { document.getElementById('tab-arrived').className = ar ? 'tab active' : 'tab'; document.getElementById('tab-unarrived').className = ar ? 'tab' : 'tab active'; renderOrders(ar); }
function saveToLocal() { localStorage.setItem('MU_DATA_V9', JSON.stringify(orders)); }
function getTagClass(s) { if(s.includes('已匯款/已下單')) return 'tag-pink'; if(s.includes('已匯款')) return 'tag-yellow'; if(s.includes('已下單')) return 'tag-green'; return 'tag-blue'; }
function del(id) { if(confirm('確定刪除？')) { orders = orders.filter(x=>x.id!==id); saveToLocal(); renderOrders(currentTabArrived); } }

function updateDropdowns() {
    const opt = (arr, h) => `<option value="">${h}</option>` + arr.map(x=>`<option value="${x}">${x}</option>`).join('');
    ['sel_platform', 'edit_platform'].forEach(id => document.getElementById(id).innerHTML = opt(config.platform, '選擇平台'));
    document.getElementById('list_series_all').innerHTML = config.series.map(s => `<option value="${s}">`).join('');
    renderConfig();
}

function addOption(k) { const v = document.getElementById('new_opt_val').value.trim(); if(v) { config[k].push(v); localStorage.setItem('m_conf_v9', JSON.stringify(config)); document.getElementById('new_opt_val').value = ''; updateDropdowns(); } }
function renderConfig() { document.getElementById('configDisplay').innerHTML = Object.keys(config).map(k => `<div style="background:white;padding:8px;border-radius:8px;margin-bottom:5px;"><small>${k}</small><div style="display:flex;flex-wrap:wrap;gap:5px;">${config[k].map((v, i) => `<span style="background:#eee;padding:2px 8px;border-radius:5px;font-size:12px;">${v} <b onclick="delOpt('${k}',${i})">×</b></span>`).join('')}</div></div>`).join(''); }
function delOpt(k, i) { config[k].splice(i,1); updateDropdowns(); }
function toggleCollapse(id, el) { const c = document.getElementById(id); const s = c.style.display === 'block'; c.style.display = s ? 'none' : 'block'; el.querySelector('span:last-child').innerText = s ? '⬇️' : '⬆️'; }
function showToast() { const t = document.getElementById('toast'); t.style.display = 'block'; setTimeout(() => t.style.display = 'none', 1500); }

// --- Excel 匯出與匯入 ---
function exportToExcel() {
    const data = [];
    orders.forEach(o => o.items.forEach(i => data.push({ "日期": o.time.slice(0,10), "平台": o.platform, "賣家": o.seller, "系列": i.s, "人物": i.c, "名稱": i.n, "單價": i.p, "數量": i.q, "狀態": o.status, "備註": o.note, "未付": o.unpaid, "運費": o.shipping, "包手": o.packing })));
    const ws = XLSX.utils.json_to_sheet(data), wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "週邊");
    XLSX.writeFile(wb, `沐沐備份_${new Date().toLocaleDateString()}.xlsx`);
}

function processExcelImport(e) {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        
        if(!json.length) { alert("檔案內無資料！"); return; }
        if(!confirm(`偵測到 ${json.length} 筆資料項目，是否覆蓋現有資料並還原？`)) return;

        // 重新構建訂單格式
        const newOrdersMap = {};
        json.forEach((row, index) => {
            // 以「日期+賣家+平台」作為同筆訂單的群組鍵
            const key = `${row['日期']}_${row['賣家']}_${row['平台']}`;
            if(!newOrdersMap[key]) {
                newOrdersMap[key] = {
                    id: Date.now() + index,
                    time: row['日期'] ? new Date(row['日期']).toISOString() : new Date().toISOString(),
                    platform: row['平台'] || "",
                    seller: row['賣家'] || "",
                    status: row['狀態'] || "已下單",
                    note: row['備註'] || "",
                    unpaid: Number(row['未付']) || 0,
                    shipping: Number(row['運費']) || 0,
                    packing: Number(row['包手']) || 0,
                    items: [],
                    arrived: (row['狀態'] || "").includes("已到貨"),
                    arriveTime: null
                };
            }
            newOrdersMap[key].items.push({
                s: row['系列'] || "",
                c: row['人物'] || "",
                t: "", // 類別 Excel 通常沒記，預設為空
                n: row['名稱'] || "",
                p: Number(row['單價']) || 0,
                q: Number(row['數量']) || 1
            });

            // 自動補充 config 選項
            if(row['平台'] && !config.platform.includes(row['平台'])) config.platform.push(row['平台']);
            if(row['系列'] && !config.series.includes(row['系列'])) config.series.push(row['系列']);
            if(row['人物'] && !config.char.includes(row['人物'])) config.char.push(row['人物']);
        });

        // 轉換回陣列並計算各項總額
        orders = Object.values(newOrdersMap).map(o => {
            let itemTotal = 0;
            o.items.forEach(it => itemTotal += it.p * it.q);
            o.itemTotal = itemTotal;
            o.pickup = o.unpaid + o.shipping + o.packing;
            return o;
        });

        saveToLocal(); localStorage.setItem('m_conf_v9', JSON.stringify(config));
        updateDropdowns(); renderOrders(false);
        alert("還原成功！");
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ""; // 清空 input 以便下次觸發
}

initTitle(); updateDropdowns(); addSubItem('mainInputForm'); renderOrders(false);
Sortable.create(document.getElementById('mainSubContainer'), { handle: '.btn-drag-sub', animation: 150 });
Sortable.create(document.getElementById('editSubContainer'), { handle: '.btn-drag-sub', animation: 150 });
</script>
</body>
</html>
