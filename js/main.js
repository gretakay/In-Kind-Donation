document.addEventListener('DOMContentLoaded', () => {
  // 判斷是哪一頁
  const donationForm = document.getElementById('donation-form');
  const recentList = document.getElementById('recent-list');
  const nearExpiryList = document.getElementById('near-expiry-list');

  if (donationForm) {
    initDonationPage();
  }

  if (recentList) {
    initRecentPage();
  }

  if (nearExpiryList) {
    initNearExpiryPage();
  }
});

/* ========= index.html：捐贈登錄頁 ========= */

function initDonationPage() {
  const form = document.getElementById('donation-form');
  const msg = document.getElementById('message');
  const donationDateInput = document.getElementById('donationDate');

  // 預設填入今天日期
  function setToday() {
    const today = new Date().toISOString().slice(0, 10);
    if (donationDateInput) {
      donationDateInput.value = today;
    }
  }
  setToday();

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    msg.textContent = '送出中…';

    const formData = new FormData(form);
    const payload = Object.fromEntries(formData.entries());

    // 若未填日期，預設今天
    if (!payload.donationDate) {
      payload.donationDate = new Date().toISOString().slice(0, 10);
    }

    // 先處理照片上傳
    const photoFile = form.photo && form.photo.files && form.photo.files[0];
    if (photoFile) {
      try {
        const photoForm = new FormData();
        photoForm.append('action', 'uploadPhoto');
        photoForm.append('photo', photoFile);
        // 傳送品項名稱給後端，讓後端可用於檔名
        photoForm.append('itemName', form.itemName ? form.itemName.value : '');
        const res = await fetch(API_BASE, {
          method: 'POST',
          body: photoForm
        });
        const photoResult = await res.json();
        if (photoResult.success && photoResult.url) {
          payload.photoUrl = photoResult.url;
        } else {
          msg.textContent = '照片上傳失敗：' + (photoResult.message || 'unknown');
          return;
        }
      } catch (err) {
        msg.textContent = '照片上傳失敗：' + err;
        return;
      }
    }

    try {
      const result = await apiAddDonation(payload);
      if (result.success) {
        msg.textContent = '已完成登錄 ✔';
        form.reset();
        setToday(); // reset 後再帶回今天
      } else {
        msg.textContent = '發生錯誤：' + (result.message || 'unknown');
      }
    } catch (err) {
      msg.textContent = '連線失敗：' + err;
    }
  });
}

/* ========= recent.html：最近捐贈頁 ========= */

function initRecentPage() {
  const listEl = document.getElementById('recent-list');
  const limitSelect = document.getElementById('limit-select');
  const reloadBtn = document.getElementById('reload-btn');

  function formatDate(value) {
    if (!value) return '';
    const d = new Date(value);
    if (isNaN(d)) return value;
    return d.toLocaleDateString('zh-TW');
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  async function loadRecent() {
    listEl.textContent = '載入中…';
    const limit = parseInt(limitSelect.value || '20', 10);

    try {
      const result = await apiGetRecent(limit);
      if (!result.success) {
        listEl.textContent = '載入失敗：' + (result.message || 'unknown');
        return;
      }

      const data = result.data || [];
      if (!data.length) {
        listEl.textContent = '目前尚無捐贈記錄。';
        return;
      }

      const html = data.map(item => {
        const donationDate = formatDate(item.donationDate);
        const expiryDate = item.expiryDate ? formatDate(item.expiryDate) : '—';
        const qty = item.quantity || 0;
        const unit = item.unit || '';
        const location = item.location || '（未填）';
        const handler = item.handler || '（未填）';
        const photo = item.photoUrl ? `<img src="${escapeHtml(item.photoUrl)}" alt="${escapeHtml(item.itemName)}" style="max-width:48px;max-height:48px;margin-right:8px;vertical-align:middle;border-radius:6px;object-fit:cover;">` : '';

        return `
          <div class="card">
            <div class="card-main">
              ${photo}<strong>${escapeHtml(item.itemName || '（未填品項）')}</strong>
              <span> × ${qty} ${escapeHtml(unit)}</span>
            </div>
            <div>捐贈日期：${donationDate}</div>
            <div>放置位置：${escapeHtml(location)}</div>
            <div>經手人：${escapeHtml(handler)}</div>
            <div>有效期限：${expiryDate}</div>
          </div>
        `;
      }).join('');

      listEl.innerHTML = html;
    } catch (err) {
      listEl.textContent = '載入時發生錯誤：' + err;
    }
  }

  reloadBtn.addEventListener('click', loadRecent);
  limitSelect.addEventListener('change', loadRecent);

  // 初次載入
  loadRecent();
}

/* ========= near-expiry.html：即期品頁 ========= */

function initNearExpiryPage() {
  const listEl = document.getElementById('near-expiry-list');
  const daysSelect = document.getElementById('days-select');
  const reloadBtn = document.getElementById('reload-btn');

  function formatDate(value) {
    if (!value) return '';
    const d = new Date(value);
    if (isNaN(d)) return value;
    return d.toLocaleDateString('zh-TW');
  }

  function diffDays(from, to) {
    const ms = to.getTime() - from.getTime();
    return Math.floor(ms / (1000 * 60 * 60 * 24));
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  async function loadNearExpiry() {
    listEl.textContent = '載入中…';
    const days = parseInt(daysSelect.value || '7', 10);

    try {
      const result = await apiGetNearExpiry(days);
      if (!result.success) {
        listEl.textContent = '載入失敗：' + (result.message || 'unknown');
        return;
      }

      let data = result.data || [];
      if (!data.length) {
        listEl.textContent = `未來 ${days} 天內沒有即期品。`;
        return;
      }

      // 依到期日排序（最近到期的在前）
      data = data.sort((a, b) => new Date(a.expiryDate) - new Date(b.expiryDate));

      const today = new Date();

      const html = data.map(item => {
        const expiry = item.expiryDate ? new Date(item.expiryDate) : null;
        const expiryText = expiry ? formatDate(expiry) : '—';
        let remainText = '';
        if (expiry) {
          const remain = diffDays(today, expiry);
          if (remain >= 0) {
            remainText = `（還有 ${remain} 天）`;
          } else {
            remainText = `（已過期 ${Math.abs(remain)} 天）`;
          }
        }

        const qty = item.quantity || 0;
        const unit = item.unit || '';
        const location = item.location || '（未填）';
        const handler = item.handler || '（未填）';
        const photo = item.photoUrl ? `<img src="${escapeHtml(item.photoUrl)}" alt="${escapeHtml(item.itemName)}" style="max-width:48px;max-height:48px;margin-right:8px;vertical-align:middle;border-radius:6px;object-fit:cover;">` : '';

        return `
          <div class="card card-expiry">
            <div class="card-main">
              ${photo}<strong>${escapeHtml(item.itemName || '（未填品項）')}</strong>
              <span> × ${qty} ${escapeHtml(unit)}</span>
            </div>
            <div>放置位置：${escapeHtml(location)}</div>
            <div>經手人：${escapeHtml(handler)}</div>
            <div class="expiry-line">
              有效期限：${expiryText} ${remainText}
            </div>
          </div>
        `;
      }).join('');

      listEl.innerHTML = html;
    } catch (err) {
      listEl.textContent = '載入時發生錯誤：' + err;
    }
  }

  reloadBtn.addEventListener('click', loadNearExpiry);
  daysSelect.addEventListener('change', loadNearExpiry);

  // 初次載入
  loadNearExpiry();
}
