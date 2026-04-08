/* =========================
   配置：已更新為您的新 ID 與結構
========================= */
const CONFIG = {
  CLIENT_ID: "760923439271-cechi85qk63kpq5ts1e3h0n0v55249s0.apps.googleusercontent.com",
  SPREADSHEET_ID: "151gNjbFwYjTfA7fy_WjXAnCymySYmLy-INBJuQjfdMk",
  SHEET_RECORDS: "公司帳務",
  SHEET_PARAMS: "參數",
  SCOPES: "https://www.googleapis.com/auth/spreadsheets"
};

let accessToken = "";
let tokenClient = null;
let allRecords = [];
let chartInstance = null;

const $ = (id) => document.getElementById(id);

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

async function afterSignedIn() {
  $("btnSignIn").classList.add("hidden");
  $("btnSignOut").classList.remove("hidden");
  $("btnSubmit").disabled = false;
  
  $("fDate").value = new Date().toISOString().split('T')[0];
  setStatus("🔍 正在同步帳本與參數...");
  
  await fetchData();
  bindEvents();
}

function bindEvents() {
  $("btnSignOut").onclick = () => { accessToken = ""; location.reload(); };
  $("filterAccount").onchange = updateDashboard;
  $("filterCategory").onchange = updateDashboard;
  
  $("recordForm").onsubmit = async (e) => {
    e.preventDefault();
    await submitRecord();
  };
}

async function callSheetsAPI(endpoint, method = "GET", body = null) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/${endpoint}`;
  const options = { method, headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } };
  if (body) options.body = JSON.stringify(body);
  const res = await fetch(url, options);
  if (!res.ok) throw new Error("API Error");
  return res.json();
}

// 同時讀取「公司帳務」與「參數」表
async function fetchData() {
  try {
    // 讀取紀錄 (A:H 共8欄)
    const dataRes = await callSheetsAPI(`values/${encodeURIComponent(CONFIG.SHEET_RECORDS)}!A:H`);
    const rows = dataRes.values || [];
    
    allRecords = rows.slice(1).map(r => ({
      date: r[0] || "",
      account: r[1] || "",
      category: r[2] || "",
      subItem: r[3] || "",
      income: Number(r[4]) || 0,
      expense: Number(r[5]) || 0,
      balance: Number(r[6]) || 0,
      note: r[7] || ""
    }));

    // 讀取參數 (A:C 共3欄：帳戶、分類、細項)
    let paramAccounts = new Set();
    let paramCategories = new Set();
    let paramSubItems = new Set();
    
    try {
      const paramRes = await callSheetsAPI(`values/${encodeURIComponent(CONFIG.SHEET_PARAMS)}!A:C`);
      const pRows = paramRes.values || [];
      pRows.slice(1).forEach(r => {
        if(r[0]) paramAccounts.add(r[0].trim());
        if(r[1]) paramCategories.add(r[1].trim());
        if(r[2]) paramSubItems.add(r[2].trim());
      });
    } catch(e) { console.log("參數表讀取失敗或不存在，將使用歷史紀錄建立選單"); }

    // 把歷史紀錄裡出現過的名詞也加進提示中
    allRecords.forEach(r => {
      if(r.account) paramAccounts.add(r.account);
      if(r.category) paramCategories.add(r.category);
      if(r.subItem) paramSubItems.add(r.subItem);
    });

    updateDatalists([...paramAccounts], [...paramCategories], [...paramSubItems]);
    updateDashboard();
    setStatus("✅ 帳本同步完成！");
  } catch (e) {
    console.error(e);
    setStatus("❌ 讀取失敗，請確認試算表名稱是否正確");
  }
}

function updateDatalists(accounts, categories, subItems) {
  $("accountList").innerHTML = accounts.map(v => `<option value="${v}">`).join("");
  $("categoryList").innerHTML = categories.map(v => `<option value="${v}">`).join("");
  $("subItemList").innerHTML = subItems.map(v => `<option value="${v}">`).join("");

  // 更新分析看板的下拉選單
  $("filterAccount").innerHTML = `<option value="ALL">🏢 全部帳戶</option>` + accounts.map(a => `<option value="${a}">${a}</option>`).join("");
  $("filterCategory").innerHTML = `<option value="ALL">📂 全部分類</option>` + categories.map(c => `<option value="${c}">${c}</option>`).join("");
}

function updateDashboard() {
  const selAcc = $("filterAccount").value;
  const selCat = $("filterCategory").value;

  let filtered = allRecords;
  if (selAcc !== "ALL") filtered = filtered.filter(r => r.account === selAcc);
  
  // 計算當前餘額 (抓取該帳戶的最後一筆紀錄)
  let currentBalance = "-";
  if (selAcc !== "ALL") {
    const accRecords = allRecords.filter(r => r.account === selAcc).sort((a,b) => new Date(a.date) - new Date(b.date));
    if (accRecords.length > 0) currentBalance = accRecords[accRecords.length - 1].balance.toLocaleString();
  } else {
    currentBalance = "請選擇單一帳戶";
  }
  $("dispBalance").textContent = currentBalance;

  // 篩選分類
  if (selCat !== "ALL") filtered = filtered.filter(r => r.category === selCat);
  
  const totalInc = filtered.reduce((sum, r) => sum + r.income, 0);
  const totalExp = filtered.reduce((sum, r) => sum + r.expense, 0);
  
  $("dispIncome").textContent = totalInc.toLocaleString();
  $("dispExpense").textContent = totalExp.toLocaleString();

  renderTable(filtered);
  renderChart(filtered);
}

function renderTable(data) {
  const html = data.sort((a,b) => new Date(b.date) - new Date(a.date)).slice(0, 50).map(r => `
    <tr>
      <td>${r.date}</td>
      <td>${r.account}</td>
      <td>${r.category}</td>
      <td>${r.subItem}</td>
      <td class="txt-right font-bold" style="color:#2a9d8f">${r.income > 0 ? r.income.toLocaleString() : '-'}</td>
      <td class="txt-right font-bold" style="color:#e76f51">${r.expense > 0 ? r.expense.toLocaleString() : '-'}</td>
      <td class="txt-right font-bold">${r.balance.toLocaleString()}</td>
      <td><small>${r.note}</small></td>
    </tr>
  `).join("");
  $("recordsTbody").innerHTML = html || `<tr><td colspan="8" class="empty-msg">尚未有紀錄喔 🍃</td></tr>`;
}

function renderChart(data) {
  const ctx = $("trendChart").getContext("2d");
  const labels = [], incData = [], expData = [];
  const today = new Date();

  for (let i = 5; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const mStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    // 支援 yyyy/mm/dd 或 yyyy-mm-dd 格式
    const mStrAlt = `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}`;
    const mStrAlt2 = `${d.getFullYear()}/${d.getMonth() + 1}`; 
    
    labels.push(mStr);

    const mRecords = data.filter(r => r.date.startsWith(mStr) || r.date.startsWith(mStrAlt) || r.date.startsWith(mStrAlt2));
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

async function submitRecord() {
  $("btnSubmit").disabled = true;
  $("btnSubmit").innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 寫入中...';

  const date = $("fDate").value;
  const account = $("fAccount").value;
  const category = $("fCategory").value;
  const subItem = $("fSubItem").value;
  const type = $("fType").value;
  const amount = Number($("fAmount").value);
  const note = $("fNote").value;

  const income = type === "收入" ? amount : "";
  const expense = type === "支出" ? amount : "";

  // 自動計算餘額
  const accRecords = allRecords.filter(r => r.account === account).sort((a,b) => new Date(a.date) - new Date(b.date));
  let lastBalance = accRecords.length > 0 ? accRecords[accRecords.length - 1].balance : 0;
  let newBalance = type === "收入" ? lastBalance + amount : lastBalance - amount;

  // 注意：這裡寫入 8 個欄位 (A~H)
  const row = [date, account, category, subItem, income, expense, newBalance, note];

  try {
    await callSheetsAPI(`values/${encodeURIComponent(CONFIG.SHEET_RECORDS)}!A:H:append?valueInputOption=USER_ENTERED`, "POST", { values: [row] });
    
    // 清空金額與備註，保留日期等欄位方便連續記帳
    $("fAmount").value = ""; 
    $("fNote").value = "";
    
    await fetchData(); 
    showCuteToast();
  } catch (e) {
    alert("寫入失敗，請確認權限或網路狀態！");
  } finally {
    $("btnSubmit").disabled = false;
    $("btnSubmit").innerHTML = '<i class="fa-solid fa-paper-plane"></i> 送出進帳本';
  }
}

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
