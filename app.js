/* =========================
   配置區：請確認試算表 ID
========================= */
const CONFIG = {
  CLIENT_ID: "760923439271-cechi85qk63kpq5ts1e3h0n0v55249s0.apps.googleusercontent.com",
  SPREADSHEET_ID: "151gNjbFwYjTfA7fy_WjXAnCymySYmLy-INBJuQjfdMk",
  SHEET_NAME: "公司帳務", // 請確認 Google 試算表中有這個名稱的工作表
  SCOPES: "https://www.googleapis.com/auth/spreadsheets"
};

let accessToken = "";
let tokenClient = null;
let allRecords = [];
let chartInstance = null;

const $ = (id) => document.getElementById(id);

// GIS 載入完成，掛載到 Window
window.initGis = function() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.CLIENT_ID,
    scope: CONFIG.SCOPES,
    callback: (resp) => {
      if (resp.error) {
        setStatus("❌ 登入失敗: " + resp.error);
        return;
      }
      accessToken = resp.access_token;
      afterSignedIn();
    }
  });
  $("btnSignIn").onclick = () => tokenClient.requestAccessToken();
  setStatus("✨ 準備就緒，請點擊登入");
};

function setStatus(msg) { $("statusText").textContent = msg; }

// 登入後初始化
async function afterSignedIn() {
  $("btnSignIn").classList.add("hidden");
  $("btnSignOut").classList.remove("hidden");
  $("btnSubmit").disabled = false;
  
  $("fDate").value = new Date().toISOString().split('T')[0];
  setStatus("🔍 正在同步帳本資料...");
  
  await fetchData();
  bindEvents();
}

// 綁定事件
function bindEvents() {
  $("btnSignOut").onclick = () => { accessToken = ""; location.reload(); };
  $("filterAccount").onchange = updateDashboard;
  $("filterItem").onchange = updateDashboard;
  
  $("recordForm").onsubmit = async (e) => {
    e.preventDefault();
    await submitRecord();
  };
}

// 呼叫 Google Sheets API
async function callSheetsAPI(endpoint, method = "GET", body = null) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/${endpoint}`;
  const options = { method, headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } };
  if (body) options.body = JSON.stringify(body);
  const res = await fetch(url, options);
  if (!res.ok) throw new Error("API Request Failed");
  return res.json();
}

// 獲取與解析資料
async function fetchData() {
  try {
    const data = await callSheetsAPI(`values/${encodeURIComponent(CONFIG.SHEET_NAME)}!A:G`);
    const rows = data.values || [];
    
    allRecords = rows.slice(1).map(r => ({
      date: r[0] || "",
      account: r[1] || "",
      item: r[2] || "",
      income: Number(r[3]) || 0,
      expense: Number(r[4]) || 0,
      balance: Number(r[5]) || 0,
      note: r[6] || ""
    }));

    updateDatalists();
    updateDashboard();
    setStatus("✅ 帳本同步完成！");
  } catch (e) {
    setStatus("❌ 讀取失敗，請確認是否建立了「公司帳務」工作表");
  }
}

// 更新下拉選單清單
function updateDatalists() {
  const accounts = [...new Set(allRecords.map(r => r.account).filter(Boolean))];
  const items = [...new Set(allRecords.map(r => r.item).filter(Boolean))];

  $("accountList").innerHTML = accounts.map(a => `<option value="${a}">`).join("");
  $("itemList").innerHTML = items.map(i => `<option value="${i}">`).join("");

  // 保留"ALL"選項，再加入動態產生的選項
  $("filterAccount").innerHTML = `<option value="ALL">🏢 全部公司/帳戶</option>` + accounts.map(a => `<option value="${a}">${a}</option>`).join("");
  $("filterItem").innerHTML = `<option value="ALL">📂 全部項目</option>` + items.map(i => `<option value="${i}">${i}</option>`).join("");
}

// 更新看板、表格與圖表
function updateDashboard() {
  const selAcc = $("filterAccount").value;
  const selItem = $("filterItem").value;

  let filtered = allRecords;
  if (selAcc !== "ALL") filtered = filtered.filter(r => r.account === selAcc);
  
  // 計算帳戶餘額：抓取該帳戶「最新日期」的餘額
  let currentBalance = "-";
  if (selAcc !== "ALL") {
    const accRecords = allRecords.filter(r => r.account === selAcc).sort((a,b) => new Date(a.date) - new Date(b.date));
    if (accRecords.length > 0) currentBalance = accRecords[accRecords.length - 1].balance.toLocaleString();
  } else {
    currentBalance = "請選擇單一帳戶";
  }
  $("dispBalance").textContent = currentBalance;

  // 篩選項目算收支
  if (selItem !== "ALL") filtered = filtered.filter(r => r.item === selItem);
  
  const totalInc = filtered.reduce((sum, r) => sum + r.income, 0);
  const totalExp = filtered.reduce((sum, r) => sum + r.expense, 0);
  
  $("dispIncome").textContent = totalInc.toLocaleString();
  $("dispExpense").textContent = totalExp.toLocaleString();

  renderTable(filtered);
  renderChart(filtered);
}

// 繪製表格
function renderTable(data) {
  const html = data.sort((a,b) => new Date(b.date) - new Date(a.date)).slice(0, 50).map(r => `
    <tr>
      <td>${r.date}</td>
      <td>${r.account}</td>
      <td>${r.item}</td>
      <td class="txt-right" style="color:#2a9d8f">${r.income > 0 ? r.income.toLocaleString() : '-'}</td>
      <td class="txt-right" style="color:#e76f51">${r.expense > 0 ? r.expense.toLocaleString() : '-'}</td>
      <td class="txt-right font-weight-bold">${r.balance.toLocaleString()}</td>
      <td><small>${r.note}</small></td>
    </tr>
  `).join("");
  $("recordsTbody").innerHTML = html || `<tr><td colspan="7" class="empty-msg">尚未有紀錄喔 🍃</td></tr>`;
}

// 繪製近六個月趨勢圖
function renderChart(data) {
  const ctx = $("trendChart").getContext("2d");
  const labels = [], incData = [], expData = [];
  const today = new Date();

  for (let i = 5; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const mStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    labels.push(mStr);

    const mRecords = data.filter(r => r.date.startsWith(mStr));
    incData.push(mRecords.reduce((sum, r) => sum + r.income, 0));
    expData.push(mRecords.reduce((sum, r) => sum + r.expense, 0));
  }

  if (chartInstance) chartInstance.destroy();
  chartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: '收入 💰', data: incData, backgroundColor: '#B5EAD7', borderRadius: 5 },
        { label: '支出 💸', data: expData, backgroundColor: '#FFB7B2', borderRadius: 5 }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false }
  });
}

// 寫入新紀錄
async function submitRecord() {
  $("btnSubmit").disabled = true;
  $("btnSubmit").innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 寫入中...';

  const date = $("fDate").value;
  const account = $("fAccount").value;
  const item = $("fItem").value;
  const type = $("fType").value;
  const amount = Number($("fAmount").value);
  const note = $("fNote").value;

  const income = type === "收入" ? amount : "";
  const expense = type === "支出" ? amount : "";

  // 自動計算餘額
  const accRecords = allRecords.filter(r => r.account === account).sort((a,b) => new Date(a.date) - new Date(b.date));
  let lastBalance = accRecords.length > 0 ? accRecords[accRecords.length - 1].balance : 0;
  let newBalance = type === "收入" ? lastBalance + amount : lastBalance - amount;

  const row = [date, account, item, income, expense, newBalance, note];

  try {
    await callSheetsAPI(`values/${encodeURIComponent(CONFIG.SHEET_NAME)}!A:G:append?valueInputOption=USER_ENTERED`, "POST", { values: [row] });
    
    // 重設表單金額與備註，保留日期與公司
    $("fAmount").value = ""; $("fNote").value = "";
    
    await fetchData(); // 重新整理資料
    showCuteToast();
  } catch (e) {
    alert("寫入失敗，請確認權限或網路狀態！");
  } finally {
    $("btnSubmit").disabled = false;
    $("btnSubmit").innerHTML = '<i class="fa-solid fa-paper-plane"></i> 送出進帳本';
  }
}

// 隨機鼓勵視窗
const encourageMsgs = [
  "老闆太神啦！金幣叮噹響 💰", 
  "記帳完成！喝杯好茶休息一下吧 🍵", 
  "太棒了！每一筆都是成長的印記 🐾", 
  "順利存檔！公司營運有你真放心 ✨", 
  "好棒！清楚的帳務是成功的基石 🏰"
];

function showCuteToast() {
  $("toastMsg").textContent = encourageMsgs[Math.floor(Math.random() * encourageMsgs.length)];
  const toast = $("toast");
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 3500);
}
