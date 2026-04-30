// 任務管理腳本
// 透過瀏覽器 localStorage 保存資料來模擬資料庫

(function () {
  "use strict";

  /**
   * 國定假日資料。以年份為 key，值為假日日期字串陣列 (YYYY-MM-DD)。
   * 如果需要更多年份，可依需求擴充。
   */
  const holidayMap = {
    2026: [
      "2026-01-01", // 新年、開國紀念日
      "2026-02-15", // 春節前日 (補假)
      "2026-02-16", // 除夕
      "2026-02-17", // 春節
      "2026-02-18", // 春節假期
      "2026-02-19", // 春節假期
      "2026-02-20", // 春節假期
      "2026-02-27", // 228紀念日補假
      "2026-02-28", // 228紀念日
      "2026-04-03", // 兒童節補假
      "2026-04-04", // 兒童節
      "2026-04-05", // 清明節
      "2026-04-06", // 清明節補假
      "2026-05-01", // 勞動節
      "2026-06-19", // 端午節
      "2026-09-25", // 中秋節
      "2026-09-28", // 孔子誕辰紀念日 (教師節)
      "2026-10-09", // 國慶日補假
      "2026-10-10", // 國慶日
      "2026-10-25", // 光復節
      "2026-10-26", // 光復節補假
      "2026-12-25"  // 行憲紀念日
    ],
    2027: [
      "2027-01-01", // 新年、開國紀念日
      "2027-02-05", "2027-02-06", "2027-02-07", "2027-02-08", "2027-02-09", "2027-02-10", // 春節假期
      "2027-02-28", // 228紀念日
      "2027-03-01", // 228補假
      "2027-04-04", // 兒童節
      "2027-04-05", // 清明節及兒童節補假
      "2027-05-01", // 勞動節
      "2027-06-09", // 端午節
      "2027-09-15", // 中秋節
      "2027-09-28", // 孔子誕辰紀念日
      "2027-10-10", // 國慶日
      "2027-10-11", // 國慶日補假
      "2027-10-25", // 光復節
      "2027-12-25", // 行憲紀念日
      "2027-12-31"  // 開國紀念日補假 (隔年跨年放假)
    ],
    2028: [
      "2028-01-01", // 新年、開國紀念日
      "2028-01-25", "2028-01-26", "2028-01-27", "2028-01-28", "2028-01-29", "2028-01-30", // 春節假期
      "2028-02-28", // 228紀念日
      "2028-04-04", // 兒童節及清明節
      "2028-05-01", // 勞動節
      "2028-05-28", // 端午節
      "2028-10-03", // 中秋節
      "2028-10-10"  // 國慶日
    ]
  };

  /**
   * API 端點 URL（Google Apps Script 部署的 Web App）
   * 請將此處替換為您部署的 Web App URL，例如：https://script.google.com/macros/s/XXXXXXXX/exec
   */
  const API_URL = "https://script.google.com/macros/s/AKfycbwaeouIR-n4HMDYAnWH4T2tg9hh78iWVRvy7V25VZHSxlBSBXnw1fGy0WJH78neAXI/exec";

  /**
   * 任務快取。所有操作皆在此陣列處理並同步至 Google Sheets
   */
  let cachedTasks = [];

  /**
   * 從快取取得任務清單
   * @returns {Array}
   */
  function getTasks() {
    return cachedTasks;
  }

  /**
   * 從遠端 Google Sheets 讀取任務資料，並更新快取
   * @returns {Promise<Array>}
   */
  async function fetchTasks() {
    try {
      const response = await fetch(API_URL);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      const data = await response.json();
      cachedTasks = Array.isArray(data) ? data : [];
    } catch (err) {
      console.error("讀取任務資料失敗：", err);
      cachedTasks = [];
    }
    return cachedTasks;
  }

  /**
   * 將任務資料同步到遠端 Google Sheets
   * @param {Array} tasks 任務陣列
   * @returns {Promise<void>}
   */
  async function saveTasks(tasks) {
    cachedTasks = tasks;
    try {
      await fetch(API_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ tasks: cachedTasks })
      });
    } catch (err) {
      console.error("儲存任務資料失敗：", err);
    }
  }

  /**
   * 判斷是否為國定假日
   * @param {Date} date 日期
   * @returns {boolean}
   */
  function isHoliday(date) {
    const year = date.getFullYear();
    const formatted = date.toISOString().split("T")[0];
    const list = holidayMap[year] || [];
    return list.includes(formatted);
  }

  /**
   * 判斷是否為工作日（非週末且非國定假日）
   * @param {Date} date 日期
   * @returns {boolean}
   */
  function isBusinessDay(date) {
    const day = date.getDay();
    // 0: Sunday, 6: Saturday
    if (day === 0 || day === 6) {
      return false;
    }
    return !isHoliday(date);
  }

  /**
   * 計算預計完成日期。從指定的起始日期開始，根據開發天數排除週末及台灣假日。
   * @param {string} startDateStr 起始日期字串（YYYY-MM-DD）
   * @param {number} days 開發天數（必須 >= 1）
   * @returns {Date} 完成日期
   */
  function computeCompletionDate(startDateStr, days) {
    let date = new Date(startDateStr);
    // 若起始日不是工作日，找出下一個工作日作為開始
    while (!isBusinessDay(date)) {
      date.setDate(date.getDate() + 1);
    }
    let count = 0;
    let result = new Date(date);
    while (count < days) {
      if (isBusinessDay(result)) {
        count++;
        if (count === days) {
          break;
        }
      }
      result.setDate(result.getDate() + 1);
    }
    return result;
  }

  /**
   * 計算從今天到指定日期的剩餘天數。
   * 若日期已過，回傳負值。
   * @param {string} dateStr 目標日期字串
   * @returns {number} 天數
   */
  function daysUntil(dateStr) {
    const today = new Date();
    // 將時間設為 00:00:00 以避免時區與夏令時間影響
    const todayStart = new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate()
    );
    const target = new Date(dateStr);
    const diff = target - todayStart;
    return Math.floor(diff / (24 * 60 * 60 * 1000));
  }

  /**
   * 將日期格式化為 YYYY-MM-DD 字串
   * @param {Date} date 日期物件
   * @returns {string}
   */
  function formatDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  /**
   * 處理表單送出以新增任務
   */
  async function handleAddTask(event) {
    event.preventDefault();
    const name = document.getElementById("taskName").value.trim();
    const desc = document.getElementById("taskDesc").value.trim();
    const missing = document.getElementById("missingContent").value.trim();
    const dueDate = document.getElementById("dueDate").value;
    const devDays = parseInt(document.getElementById("devDays").value, 10);

    if (!name || !desc || !dueDate || !devDays || devDays <= 0) {
      showMessage("請完整填寫必填欄位", true);
      return;
    }
    // 先取得最新任務資料
    await fetchTasks();
    const tasks = getTasks();
    const id = Date.now();
    tasks.push({
      id,
      name,
      desc,
      missing,
      dueDate,
      devDays,
      completed: false
    });
    await saveTasks(tasks);
    showMessage("任務已新增！");
    // 清除輸入
    document.getElementById("taskForm").reset();
  }

  /**
   * 顯示訊息
   * @param {string} text 顯示文字
   * @param {boolean} isError 是否為錯誤訊息
   */
  function showMessage(text, isError = false) {
    const msgEl = document.getElementById("message");
    if (msgEl) {
      msgEl.textContent = text;
      msgEl.style.color = isError ? "#e36262" : "#28a745";
      // 設定幾秒後自動消失
      setTimeout(() => {
        msgEl.textContent = "";
      }, 3000);
    }
  }

  /**
   * 渲染任務列表
   */
  function renderTasks() {
    const container = document.getElementById("tasksContainer");
    if (!container) return;
    const tasks = getTasks();
    // 清空內容
    container.innerHTML = "";
    if (!tasks || tasks.length === 0) {
      const p = document.createElement("p");
      p.textContent = "目前沒有任何任務";
      container.appendChild(p);
      return;
    }
    // 建立表格
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    const headers = [
      "項目名稱",
      "內容敘述",
      "客戶缺少提供內容",
      "客戶提供日期",
      "開發天數",
      "預計完成日期",
      "客戶提供倒數天數",
      "完成與否",
      "操作"
    ];
    headers.forEach((h) => {
      const th = document.createElement("th");
      th.textContent = h;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    const tbody = document.createElement("tbody");
    tasks.forEach((task) => {
      tbody.appendChild(buildTaskRow(task));
    });
    table.appendChild(tbody);
    container.appendChild(table);
  }

  /**
   * 建立單一任務列
   * @param {Object} task 任務資料
   * @returns {HTMLTableRowElement}
   */
  function buildTaskRow(task) {
    const tr = document.createElement("tr");
    // 計算預計完成日期
    const completionDate = computeCompletionDate(task.dueDate, task.devDays);
    const completionDateStr = formatDate(completionDate);
    // 計算倒數天數
    const daysDiff = daysUntil(task.dueDate);
    const countdown = daysDiff < 0 ? 0 : daysDiff;
    // 判斷是否逾期
    const today = new Date();
    const isOverdue = completionDate < new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate()
    );
    if (task.completed) {
      tr.classList.add("completed");
    } else if (isOverdue) {
      tr.classList.add("overdue");
    }
    // 各欄位
    const cols = [
      task.name,
      task.desc,
      task.missing || "",
      task.dueDate,
      String(task.devDays),
      completionDateStr,
      String(countdown),
      task.completed ? "完成" : "未完成"
    ];
    cols.forEach((text) => {
      const td = document.createElement("td");
      td.textContent = text;
      tr.appendChild(td);
    });
    // 操作欄：直接顯示各個動作按鈕，而非下拉選單
    const actionTd = document.createElement("td");
    actionTd.className = "actions";

    // 標記完成或未完成按鈕
    const toggleBtn = document.createElement("button");
    toggleBtn.type = "button";
    toggleBtn.className = "action-toggle";
    toggleBtn.textContent = task.completed ? "未完成" : "完成";
    toggleBtn.addEventListener("click", () => {
      if (task.completed) {
        markIncomplete(task.id);
      } else {
        markComplete(task.id);
      }
    });
    actionTd.appendChild(toggleBtn);

    // 編輯按鈕
    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "action-edit";
    editBtn.textContent = "修改";
    editBtn.addEventListener("click", () => {
      openEditModal(task.id);
    });
    actionTd.appendChild(editBtn);

    // 刪除按鈕
    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "action-delete";
    deleteBtn.textContent = "刪除";
    deleteBtn.addEventListener("click", () => {
      if (confirm("確定要刪除此項目嗎？")) {
        deleteTask(task.id);
      }
    });
    actionTd.appendChild(deleteBtn);

    tr.appendChild(actionTd);
    return tr;
  }

  /**
   * 標記任務為完成
   * @param {number} id 任務 id
   */
  async function markComplete(id) {
    // 先載入最新資料
    await fetchTasks();
    const tasks = getTasks();
    const idx = tasks.findIndex((t) => t.id === id);
    if (idx !== -1) {
      tasks[idx].completed = true;
      await saveTasks(tasks);
      renderTasks();
    }
  }

  /**
   * 標記任務為未完成
   * 有時已標記完成的任務後續出現問題，需要恢復為未完成
   * @param {number} id 任務 id
   */
  async function markIncomplete(id) {
    await fetchTasks();
    const tasks = getTasks();
    const idx = tasks.findIndex((t) => t.id === id);
    if (idx !== -1) {
      tasks[idx].completed = false;
      await saveTasks(tasks);
      renderTasks();
    }
  }

  /**
   * 匯出任務資料為 Excel 檔案
   * 透過產生 HTML table 並以 data URI 方式下載
   */
  function exportToExcel() {
    const tasks = getTasks();
    if (!tasks || tasks.length === 0) {
      alert("沒有任務可輸出");
      return;
    }
    // 標題列
    const headers = [
      "項目名稱",
      "內容敘述",
      "客戶缺少提供內容",
      "客戶提供日期",
      "預計開發天數",
      "預計完成日期",
      "客戶提供倒數天數",
      "完成與否"
    ];
    // 產生資料列
    const rows = tasks.map((task) => {
      const completionDate = computeCompletionDate(task.dueDate, task.devDays);
      const completionDateStr = formatDate(completionDate);
      const daysDiff = daysUntil(task.dueDate);
      const countdown = daysDiff < 0 ? 0 : daysDiff;
      const status = task.completed ? "完成" : "未完成";
      return [
        task.name,
        task.desc,
        task.missing || "",
        task.dueDate,
        String(task.devDays),
        completionDateStr,
        String(countdown),
        status
      ];
    });
    // 建立 HTML 表格字串
    let tableHTML = '<table><thead><tr>';
    headers.forEach((h) => {
      tableHTML += `<th>${h}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';
    rows.forEach((row) => {
      tableHTML += '<tr>';
      row.forEach((cell) => {
        tableHTML += `<td>${cell}</td>`;
      });
      tableHTML += '</tr>';
    });
    tableHTML += '</tbody></table>';
    // 建立下載連結
    const dataType = 'application/vnd.ms-excel';
    const dataUri = 'data:' + dataType + ';charset=utf-8,' + encodeURIComponent(tableHTML);
    const fileName = 'tasks_' + new Date().toISOString().slice(0, 10) + '.xls';
    const link = document.createElement('a');
    link.href = dataUri;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  /**
   * 開啟編輯對話框
   * @param {number} id 任務 id
   */
  function openEditModal(id) {
    const modal = document.getElementById("editModal");
    if (!modal) return;
    const tasks = getTasks();
    const task = tasks.find((t) => t.id === id);
    if (!task) return;
    // 填入原始資料
    document.getElementById("editId").value = task.id;
    document.getElementById("editDesc").value = task.desc;
    document.getElementById("editMissing").value = task.missing || "";
    document.getElementById("editDueDate").value = task.dueDate;
    document.getElementById("editDevDays").value = task.devDays;
    // 顯示對話框
    modal.classList.remove("hidden");
    modal.setAttribute("aria-hidden", "false");
  }

  /**
   * 關閉編輯對話框
   */
  function closeEditModal() {
    const modal = document.getElementById("editModal");
    if (modal) {
      modal.classList.add("hidden");
      modal.setAttribute("aria-hidden", "true");
    }
  }

  /**
   * 更新任務資料
   * @param {Event} event 表單提交事件
   */
  async function handleEditSubmit(event) {
    event.preventDefault();
    const id = parseInt(document.getElementById("editId").value, 10);
    const desc = document.getElementById("editDesc").value.trim();
    const missing = document.getElementById("editMissing").value.trim();
    const dueDate = document.getElementById("editDueDate").value;
    const devDays = parseInt(document.getElementById("editDevDays").value, 10);
    if (!desc || !dueDate || !devDays || devDays <= 0) {
      alert("請完整填寫必填欄位");
      return;
    }
    await fetchTasks();
    const tasks = getTasks();
    const idx = tasks.findIndex((t) => t.id === id);
    if (idx !== -1) {
      tasks[idx].desc = desc;
      tasks[idx].missing = missing;
      tasks[idx].dueDate = dueDate;
      tasks[idx].devDays = devDays;
      await saveTasks(tasks);
      renderTasks();
      closeEditModal();
    }
  }

  /**
   * 刪除任務
   * @param {number} id 任務 id
   */
  async function deleteTask(id) {
    await fetchTasks();
    let tasks = getTasks();
    tasks = tasks.filter((t) => t.id !== id);
    await saveTasks(tasks);
    renderTasks();
  }

  /**
   * 初始化輸入頁
   */
  function initInputPage() {
    const form = document.getElementById("taskForm");
    if (form) {
      form.addEventListener("submit", handleAddTask);
    }
    const viewBtn = document.getElementById("viewTasks");
    if (viewBtn) {
      viewBtn.addEventListener("click", () => {
        // 導向顯示頁面
        window.location.href = "tasks.html";
      });
    }
  }

  /**
   * 初始化顯示頁
   */
  function initTasksPage() {
    // 載入遠端資料後再渲染
    fetchTasks().then(() => {
      renderTasks();
    });
    const backBtn = document.getElementById("backToInput");
    if (backBtn) {
      backBtn.addEventListener("click", () => {
        window.location.href = "index.html";
      });
    }
    // 編輯表單提交
    const editForm = document.getElementById("editForm");
    if (editForm) {
      editForm.addEventListener("submit", handleEditSubmit);
    }
    // 編輯表單取消
    const cancelBtn = document.getElementById("cancelEdit");
    if (cancelBtn) {
      cancelBtn.addEventListener("click", () => {
        closeEditModal();
      });
    }
    // 匯出 Excel 按鈕事件
    const exportBtn = document.getElementById("exportExcel");
    if (exportBtn) {
      exportBtn.addEventListener("click", () => {
        exportToExcel();
      });
    }
  }

  // 根據頁面不同做初始化
  document.addEventListener("DOMContentLoaded", () => {
    if (document.getElementById("taskForm")) {
      initInputPage();
    }
    if (document.getElementById("tasksContainer")) {
      initTasksPage();
    }
  });
})();