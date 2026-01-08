/*********************************
 * HTML Web App Router
 *********************************/
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';
  try {
    return HtmlService
      .createTemplateFromFile(page)
      .evaluate()
      .setTitle('實物捐贈暨資產管理系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput("頁面不存在：" + page);
  }
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// 定義工作表名稱
const SHEET_NAME = 'Donations';       
const TRANS_SHEET_NAME = 'Transactions'; 
const ASSET_SHEET_NAME = 'Assets';     
const REQUIRED_FIELDS = ['donorName', 'itemName', 'quantity'];

const _sheets = {};

function getSheet(name) {
  if (_sheets[name]) return _sheets[name];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_NAME) {
      sheet.appendRow(['系統時間', '日期', '來源/捐贈者', '物品名稱', '單位', '數量(標準化)', '位置', '經辦人', '效期', '照片網址', '分類', '顏色', '庫存狀態']);
    } else if (name === TRANS_SHEET_NAME) {
      sheet.appendRow(['紀錄時間', '異動類型', '物品名稱', '資產編號', '數量', '領用/借用人', '預計歸還日', '狀態', '經手人', '備註']);
    } else if (name === ASSET_SHEET_NAME) {
      sheet.appendRow(['建檔日期', '資產編號', '物品名稱', '顏色規格', '來源類別', '型號規格', '存放位置', '目前狀態', '固定保管人', '目前借用人', '備註', '單位', '照片網址']);
    }
    sheet.setFrozenRows(1);
  }
  _sheets[name] = sheet;
  return sheet;
}

/*********************************
 * API 功能函式
 *********************************/

function getSummaryData(includeSangha = false) {
  try {
    return {
      success: true,
      inventory: getInventorySummary(includeSangha).data || [],
      assets: getAggregatedAssets() || [],
      recent: getRecentDonations(20).data || [],
      expiry: getNearExpiry(7).data || [] 
    };
  } catch (err) { return { success: false, message: "資料載入失敗: " + err.toString() }; }
}

/** 🚀 核心修復：即期品偵測邏輯 */
function getNearExpiry(days) {
  try {
    const sheet = getSheet(SHEET_NAME);
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return { success: true, data: [] };

    const today = new Date();
    today.setHours(0, 0, 0, 0); 
    const limitDate = new Date();
    limitDate.setDate(today.getDate() + days);
    limitDate.setHours(23, 59, 59, 999);

    const expiryList = rows.slice(1).filter(r => {
      const expiryDate = r[8]; 
      if (!expiryDate || !(expiryDate instanceof Date)) return false;
      const checkDate = new Date(expiryDate);
      return checkDate >= today && checkDate <= limitDate;
    }).map(r => ({
      itemName: r[3], quantity: r[5], unit: r[4], 
      expiryDate: Utilities.formatDate(r[8], "GMT+8", "yyyy-MM-dd"),
      location: r[6], category: r[10], color: r[11], photoUrl: r[9]
    }));
    return { success: true, data: expiryList };
  } catch (err) { return { success: false, message: "即期品抓取失敗" }; }
}

/** 【入庫】消耗品 */
function addDonation(p) {
  try {
    const missing = REQUIRED_FIELDS.filter(f => !String(p[f] || '').trim());
    if (missing.length) return { success: false, message: '必填缺失：' + missing.join('、') };
    const sheet = getSheet(SHEET_NAME);
    const ratio = Number(p.unitRatio) || 1; 
    const totalQty = Number(p.quantity) * ratio;
    const category = autoCategory(p.itemName);
    
    sheet.appendRow([new Date(), p.donationDate ? new Date(p.donationDate) : new Date(), p.donorName, p.itemName, p.unit || '個', totalQty, p.location || '', p.handler || '', p.expiryDate ? new Date(p.expiryDate) : '', p.photoUrl || '', category, p.color || '無', p.itemStatus || '可用']);
    return { success: true, category: category };
  } catch (err) { return { success: false, message: err.toString() }; }
}

/** 【建檔】固定資產 */
function importAsset(p) {
  try {
    const sheet = getSheet(ASSET_SHEET_NAME);
    const count = parseInt(p.assetQty, 10) || 1; 
    const fullData = sheet.getDataRange().getValues();
    const category = autoCategory(p.itemName);
    const yearShort = Utilities.formatDate(new Date(), "GMT+8", "yy"); 
    
    const prefixMap = { '佛事用具': 'BT', '家具類': 'FUR', '防疫/醫療': 'MED', '蔬果類': 'VEG', '五穀糧食': 'GRN', '豆奶': 'PRO', '調味油品': 'OIL', '加工食品': 'PRO', '飲品飲料': 'DRK', '民生用品': 'LIF', '衣物寢具': 'CLO', '圖書影音': 'LIB', '文具辦公': 'OFF', '資訊耗材': 'IT', '五金工具': 'TLS' };
    const prefix = prefixMap[category] || 'AST';
    const searchPrefix = prefix + yearShort; 

    let maxSerial = 0;
    for (let i = 1; i < fullData.length; i++) {
      const idCell = String(fullData[i][1]);
      const matches = idCell.match(/\d{3}$/); 
      if (matches) {
        const lastNum = parseInt(matches[0], 10);
        if (!isNaN(lastNum) && lastNum > maxSerial) maxSerial = lastNum;
      }
    }
    let assetIds = [];
    for (let i = 1; i <= count; i++) { assetIds.push(searchPrefix + ("00" + (maxSerial + i)).slice(-3)); }
    const idString = assetIds.join(', ');
    const fixedHolder = (p.keeper && p.keeper.trim() !== "") ? p.keeper : "庫房";
    sheet.appendRow([new Date(), idString, p.itemName, p.color || '無', p.sourceType, p.spec || '', p.location || '', '在庫', fixedHolder, '', p.note || '', p.unit || '個', p.photoUrl || '']);
    return { success: true, message: `建檔成功`, id: idString };
  } catch (err) { return { success: false, message: err.toString() }; }
}

/** 🚀 核心優化：消耗品領用 + 固定資產精確借出 (高穩定、高容錯比對) */
function withdrawItem(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 鎖定30秒，確保併發安全
    const requestQty = Math.abs(Number(p.quantity));
    const now = new Date();
    const tSheet = getSheet(TRANS_SHEET_NAME);
    const aSheet = getSheet(ASSET_SHEET_NAME);
    const dSheet = getSheet(SHEET_NAME);

    // 🚀 優化：前端傳入的是 "品名 (顏色)" 格式，在此進行解析或組合比對
    const targetFullName = String(p.itemName).trim();

    // 1. 判斷是否為固定資產
    const assetValues = aSheet.getDataRange().getValues();
    let isAsset = false;
    for (let i = 1; i < assetValues.length; i++) {
      const aName = String(assetValues[i][2]).trim();
      const aColor = String(assetValues[i][3]).trim();
      const aFullName = (aColor && aColor !== '無' && aColor !== '') ? `${aName} (${aColor})` : aName;
      if (aFullName === targetFullName) { isAsset = true; break; }
    }
    
    if (isAsset) {
      // --- 【固定資產拆分借出邏輯】 ---
      let selectedIds = [];
      let leftToBorrow = requestQty;
      const currentAData = aSheet.getDataRange().getValues();
      for (let i = 1; i < currentAData.length; i++) {
        if (leftToBorrow <= 0) break;
        const aName = String(currentAData[i][2]).trim();
        const aColor = String(currentAData[i][3]).trim();
        const aFullName = (aColor && aColor !== '無' && aColor !== '') ? `${aName} (${aColor})` : aName;

        if (aFullName === targetFullName && currentAData[i][7] === '在庫') {
          const rowIds = String(currentAData[i][1]).split(', ').map(s => s.trim());
          if (rowIds.length <= leftToBorrow) {
            selectedIds = selectedIds.concat(rowIds);
            aSheet.getRange(i + 1, 8).setValue('借出中');
            aSheet.getRange(i + 1, 10).setValue(p.receiver); // 更新 J 欄
            leftToBorrow -= rowIds.length;
          } else {
            const toBorrow = rowIds.slice(0, leftToBorrow);
            const toKeep = rowIds.slice(leftToBorrow);
            aSheet.getRange(i + 1, 2).setValue(toKeep.join(', '));
            const newRow = [...currentAData[i]];
            newRow[0] = now; newRow[1] = toBorrow.join(', '); newRow[7] = '借出中'; newRow[9] = p.receiver; 
            aSheet.appendRow(newRow);
            selectedIds = selectedIds.concat(toBorrow);
            leftToBorrow = 0;
          }
        }
      }
      if (leftToBorrow > 0) throw new Error("資產在庫不足！");
      tSheet.appendRow([now, '借出', targetFullName, selectedIds.join(', '), requestQty * -1, p.receiver, p.returnDate || '', '待歸還', p.handler || '', '資產借出']);

    } else {
      // --- 【消耗品跨列累計扣除邏輯】 ---
      const dData = dSheet.getDataRange().getValues();
      let remainingToDeduct = requestQty;
      let found = false;
      for (let i = 1; i < dData.length; i++) {
        if (remainingToDeduct <= 0) break;
        const itemName = String(dData[i][3]).trim();
        const itemColor = String(dData[i][11]).trim();
        const fullName = (itemColor && itemColor !== '無' && itemColor !== '') ? `${itemName} (${itemColor})` : itemName;

        if (fullName === targetFullName) {
          found = true;
          const currentStock = Number(dData[i][5]);
          if (currentStock > 0) {
            const deduct = Math.min(currentStock, remainingToDeduct);
            dSheet.getRange(i + 1, 6).setValue(currentStock - deduct);
            remainingToDeduct -= deduct;
          }
        }
      }
      if (!found) throw new Error("庫存表中找不到該物品品名: " + targetFullName);
      if (remainingToDeduct > 0) throw new Error("庫存總量不足，尚缺：" + remainingToDeduct);
      tSheet.appendRow([now, '領用', targetFullName, '', requestQty * -1, p.receiver, '', '完成', p.handler || '', '消耗品領用']);
    }

    SpreadsheetApp.flush(); // 🚀 強制同步
    return { success: true };

  } catch (err) { return { success: false, message: err.toString() }; }
  finally { lock.releaseLock(); }
}

/** 【歸還/報損】自動清空借用人，恢復庫房權限 */
function returnAsset(p) {
  try {
    const aSheet = getSheet(ASSET_SHEET_NAME);
    const tSheet = getSheet(TRANS_SHEET_NAME);
    const assetIdsToReturn = Array.isArray(p.assetIds) ? p.assetIds : [p.assetIds]; 
    const now = new Date();
    const targetStatus = p.targetStatus || '在庫'; 
    const handler = p.handler || '系統紀錄';
    let recordedItemName = "";
    assetIdsToReturn.forEach(returnId => {
      const aData = aSheet.getDataRange().getValues();
      for (let i = aData.length - 1; i >= 1; i--) {
        let rowIds = String(aData[i][1]).split(', ').map(s => s.trim());
        if (rowIds.includes(returnId)) {
          if (!recordedItemName) recordedItemName = aData[i][2];
          if (rowIds.length === 1) { aSheet.getRange(i + 1, 8).setValue(targetStatus); aSheet.getRange(i + 1, 10).setValue(targetStatus === '在庫' ? '' : handler); } 
          else {
            const remainingIds = rowIds.filter(id => id !== returnId);
            aSheet.getRange(i + 1, 2).setValue(remainingIds.join(', '));
            const newRow = [...aData[i]];
            newRow[0] = now; newRow[1] = returnId; newRow[7] = targetStatus; newRow[9] = targetStatus === '在庫' ? '' : handler;
            aSheet.appendRow(newRow);
          }
          break;
        }
      }
    });
    tSheet.appendRow([now, targetStatus === '在庫' ? '歸還' : '資產異動', recordedItemName || "批次項目", assetIdsToReturn.join(', '), assetIdsToReturn.length, handler, '', targetStatus, handler, p.note || '']);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (err) { return { success: false, message: err.toString() }; }
}

/** 彙整庫存摘要：核心過濾供僧邏輯 (排除0庫存) */
function getInventorySummary(includeSangha = false) {
  try {
    const invMap = {};
    const dRows = getSheet(SHEET_NAME).getDataRange().getValues();
    const tRows = getSheet(TRANS_SHEET_NAME).getDataRange().getValues();
    for (let i = 1; i < dRows.length; i++) {
      const stockStatus = dRows[i][12]; 
      if (!includeSangha && stockStatus === '供僧') continue;
      const key = dRows[i][3] + (dRows[i][11] !== '無' ? " (" + dRows[i][11] + ")" : "");
      if (!invMap[key]) { invMap[key] = { name: dRows[i][3], color: dRows[i][11], qty: 0, unit: dRows[i][4], category: dRows[i][10], photoUrl: dRows[i][9], location: dRows[i][6], isSangha: (stockStatus === '供僧') }; }
      invMap[key].qty += Number(dRows[i][5]);
    }
    for (let i = 1; i < tRows.length; i++) {
      const targetName = tRows[i][2];
      for (let key in invMap) { if (key === targetName || key.startsWith(targetName + " (")) { invMap[key].qty += Number(tRows[i][4]); } }
    }
    const result = Object.values(invMap).filter(item => item.qty > 0);
    return { success: true, data: result };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function getBorrowedAssets() {
  try {
    const aSheet = getSheet(ASSET_SHEET_NAME);
    const tSheet = getSheet(TRANS_SHEET_NAME);
    const aData = aSheet.getDataRange().getValues();
    const tData = tSheet.getDataRange().getValues();
    const borrowDateMap = {};
    for (let i = 1; i < tData.length; i++) { if (tData[i][1] === '借出') { const dateStr = tData[i][0] instanceof Date ? Utilities.formatDate(tData[i][0], "GMT+8", "yyyy-MM-dd") : "2026-01-08"; String(tData[i][3]).split(', ').forEach(id => borrowDateMap[id.trim()] = dateStr); } }
    let results = [];
    aData.slice(1).forEach(r => { if (r[7] === '借出中') { String(r[1]).split(', ').forEach(id => { const tid = id.trim(); results.push({ id: tid, name: r[2], receiver: r[9] || "未知", location: r[6], photoUrl: r[12], borrowDate: borrowDateMap[tid] || "2026-01-08" }); }); } });
    return results;
  } catch (err) { return []; }
}

function getAvailableAssetsFull() {
  let results = [];
  getSheet(ASSET_SHEET_NAME).getDataRange().getValues().slice(1).forEach(r => { String(r[1]).split(', ').forEach(id => { const holderDisplay = r[9] ? `${r[8]} (借予: ${r[9]})` : r[8]; results.push({ id: id.trim(), name: r[2], color: r[3], status: r[7], keeper: holderDisplay, photoUrl: r[12], location: r[6] }); }); });
  return results;
}

function getAggregatedAssets() {
  const rows = getSheet(ASSET_SHEET_NAME).getDataRange().getValues();
  if (rows.length < 2) return [];
  const map = {};
  rows.slice(1).forEach(r => {
    const ids = String(r[1]).split(', ');
    const key = `${r[2]}|${r[5]}|${r[3]}`;
    if (!map[key]) map[key] = { name: r[2], color: r[3], spec: r[5], total: 0, inStock: 0, borrowed: 0, locations: new Set(), photoUrl: r[12] };
    map[key].total += ids.length;
    if (r[7] === '在庫') map[key].inStock += ids.length; else if (r[7] === '借出中') map[key].borrowed += ids.length;
    if (r[6]) map[key].locations.add(r[6]);
  });
  return Object.values(map).map(item => ({ ...item, location: Array.from(item.locations).join(', '), status: `在庫:${item.inStock} / 借出:${item.borrowed}` }));
}

function autoCategory(name) {
  if (!name) return '其他';
  const rules = [
    { cat: '佛事用具', regex: /香|燭|佛|經|僧|法器|蓮|燈|供|檀|拜|跪|幡|幢|鈸|鈴|木魚|淨水|香爐|金紙|平灰器/ },
    { cat: '家具類', regex: /桌|椅|床|櫃|架|凳|沙發|几|案|櫥|斗櫃|衣架/ },
    { cat: '防疫/醫療', regex: /罩|酒精|藥|護|檢測|貼|棉片|紗布|手套|消毒/ },
    { cat: '蔬果類', regex: /菜|菇|瓜|果|蕉|柑|橘|桃|李|莓|棗|筍|椒|薑|蘿蔔|芹|苗|玉米|茄/ },
    { cat: '五穀糧食', regex: /米|麵|粉|糧|薯|芋|麥|燕麥|穀|米粉|冬粉|糙米/ },
    { cat: '豆奶', regex: /豆|乳|奶|豆腐|豆干|豆漿|起司|植物奶|豆皮/ },
    { cat: '調味油品', regex: /油|鹽|糖|醬|醋|蜜|膏|味精|芡|麻油|胡椒|咖哩/ },
    { cat: '加工食品', regex: /罐|乾|餅|零食|包裝|冷凍|泡麵|即食|素料|糖果|巧克力|酥|條/ },
    { cat: '飲品飲料', regex: /水|茶|咖啡|汁|飲|奶粉|可可|沖泡|麥片|汽水/ },
    { cat: '民生用品', regex: /紙|洗|潔|皂|巾|袋|牙膏|刷|沐浴|洗髮|柔順|抹布|垃圾桶|雨具/ },
    { cat: '衣物寢具', regex: /衣|褲|鞋|襪|被|枕|毯|帽|袍|衫|床單|圍巾/ },
    { cat: '圖書影音', regex: /書|影音|CD|DVD|雜誌|刊物|冊|報|講義|光碟/ },
    { cat: '文具辦公', regex: /筆|膠|夾|剪|釘|尺|墨|印|章|資料夾/ },
    { cat: '資訊耗材', regex: /電腦|鼠|碟|線|電池|充電|usb|網路|螢幕|主機|鍵盤|硬碟/ },
    { cat: '五金工具', regex: /起子|鉗|梯|鑽|鎖|扳手|鎚|釘|鋸|捲尺|膠帶/ }
  ];
  const match = rules.find(r => r.regex.test(name));
  return match ? match.cat : '其他';
}

function getRecentDonations(limit) {
  const rows = getSheet(SHEET_NAME).getDataRange().getValues();
  return { success: true, data: rows.slice(1).reverse().slice(0, limit).map(r => ({ donationDate: Utilities.formatDate(r[1] instanceof Date ? r[1] : new Date(), "GMT+8", "yyyy-MM-dd"), donorName: r[2], itemName: r[3], unit: r[4], quantity: r[5], location: r[6], category: r[10], color: r[11], photoUrl: r[9], stockStatus: r[12] })) };
}

function exportInventoryToHtml(type) {
  const invRes = getInventorySummary(true); 
  const assetList = getAvailableAssetsFull();
  const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
  let html = `<style>table{width:100%;border-collapse:collapse;} th,td{border:1px solid #ddd;padding:8px;} th{background:#f4f4f4;}</style><h2>📊 報表 (${nowStr})</h2>`;
  if (type === 'all' || type === 'inventory') {
    html += `<h3>📦 消耗品清單</h3><table><tr><th>品名規格</th><th>存放位置</th><th>庫存數量</th></tr>`;
    invRes.data.forEach(i => html += `<tr><td><b>${i.name}</b> (${i.color||'無'})</td><td>${i.location||'庫房'}</td><td>${i.qty} ${i.unit}</td></tr>`);
    html += `</table>`;
  }
  if (type === 'all' || type === 'asset') {
    html += `<h3>🛠️ 固定資產清冊</h3><table><tr><th>編號</th><th>品名</th><th>保管/借用人</th><th>狀態</th></tr>`;
    assetList.forEach(a => html += `<tr><td>${a.id}</td><td>${a.name}</td><td>${a.keeper}</td><td>${a.status}</td></tr>`);
    html += `</table>`;
  }
  return html;
}
