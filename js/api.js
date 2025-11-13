// ⚠️ 這裡請換成你部署好的 Apps Script Web App URL（一定要是 /exec 結尾）
const API_BASE = 'https://script.google.com/macros/s/AKfycbz20qX-iLdd0ZAdLsa1pm4HSUbyQR8Dkmfb_tseEcqNu_yiiUFCS6e4MuwvBpiK4KIIoQ/exec';

/**
 * 新增捐贈紀錄
 * @param {Object} payload - 捐贈資料物件
 * @returns {Promise<Object>} - { success: boolean, message?: string }
 */
async function apiAddDonation(payload) {
  const body = new URLSearchParams({
    action: 'add',   // 後端用這個判斷要跑 addDonation
    ...payload
  });

  const res = await fetch(API_BASE, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
    },
    body
  });

  return res.json();
}

/**
 * 取得最近捐贈清單
 * @param {number} limit - 要抓幾筆
 * @returns {Promise<Object>} - { success: boolean, data?: Array }
 */
async function apiGetRecent(limit = 20) {
  const url = `${API_BASE}?action=recent&limit=${encodeURIComponent(limit)}`;
  const res = await fetch(url);
  return res.json();
}

/**
 * 取得即期品清單
 * @param {number} days - 未來幾天內到期
 * @returns {Promise<Object>} - { success: boolean, data?: Array }
 */
async function apiGetNearExpiry(days = 7) {
  const url = `${API_BASE}?action=nearexpiry&days=${encodeURIComponent(days)}`;
  const res = await fetch(url);
  return res.json();
}
