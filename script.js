const SHEET_ID = "1VquQU7_qbx-MHhUjz2w-yuIJs1sgZbJbVdXmclio80Y";
const SHEET_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv`;
const SHEET_JSONP_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq`;
const GAS_ID = "AKfycbwkyBaRuG3W9BpptBkGV07owXtwZaSz4kG0-rVi613GSNYcLV4VjGao69LxPpYYMjN_0A";
const GAS_URL = `https://script.google.com/macros/s/${GAS_ID}/exec`;
const CP_SHEET_NAME = "CP";
const IDEA_SHEET_NAME = "アイデア";
const TODO_SHEET_NAME = "やりたいこと";
const LOCAL_IDEAS_KEY = "mycee-app-ideas-v1";
const LOCAL_TODOS_KEY = "mycee-app-todos-v1";

const productCatalog = [
  { name: "ブランド全体", color: "#ffffff" },
  { name: "マイシーブランド", color: "#f4d44d" },
  { name: "スキニティブランド", color: "#78c9a2" },
  { name: "マイシー", color: "#ffd84d" },
  { name: "ホワイトプラス", color: "#6bb6ff" },
  { name: "セラム", color: "#ff8a3d" },
  { name: "ジェルバーム", color: "#60d5b5" },
  { name: "化粧水", color: "#79aee8" },
  { name: "パック", color: "#ff5d6c" },
  { name: "ナノビタムース", color: "#a7d948" },
  { name: "リポレモンティ", color: "#ffcf6e" },
  { name: "ビアス", color: "#c89bff" },
  { name: "酵素", color: "#7ad66d" },
  { name: "その他", color: "#c8c2b6" },
];

const sampleCampaigns = [
  {
    id: "sample-cp-1",
    name: "シート接続待ち：春のビタミン集中CP",
    product: "マイシー",
    startDate: "2026-05-01",
    endDate: "2026-05-31",
    detail: "Googleシートが公開されていない場合の表示例です。共有設定または公開範囲を確認すると実データに切り替わります。",
    channel: "Instagram / EC",
    owner: "My+Cee Team",
    status: "進行予定",
    needs: "シートの閲覧権限",
    source: "sample",
  },
  {
    id: "sample-cp-2",
    name: "代理店確認中：ホワイトプラスLP改善",
    product: "ホワイトプラス",
    startDate: "2026-04-15",
    endDate: "2026-05-15",
    detail: "担当、ステータス、確認事項が埋まると管理しやすくなります。",
    channel: "広告 / LP",
    owner: "Agency",
    status: "確認中",
    needs: "クリエイティブ最終案",
    source: "sample",
  },
];

const state = {
  campaigns: [],
  ideas: [],
  todos: [],
  selectedDetail: null,
  editingDetail: null,
  expanded: {
    active: false,
    upcoming: false,
    unscheduled: false,
    ideas: false,
    todos: false,
  },
  calendarDate: startOfMonth(new Date()),
};

const els = {
  connectionDot: document.querySelector("#connectionDot"),
  connectionStatus: document.querySelector("#connectionStatus"),
  activeCampaignList: document.querySelector("#activeCampaignList"),
  upcomingCampaignList: document.querySelector("#upcomingCampaignList"),
  unscheduledCampaignList: document.querySelector("#unscheduledCampaignList"),
  monthlyCalendar: document.querySelector("#monthlyCalendar"),
  calendarLabel: document.querySelector("#calendarLabel"),
  ideaGrid: document.querySelector("#ideaGrid"),
  todoGrid: document.querySelector("#todoGrid"),
  refreshButton: document.querySelector("#refreshButton"),
  cpButton: document.querySelector("#cpButton"),
  cpButtonInline: document.querySelector("#cpButtonInline"),
  prevMonthButton: document.querySelector("#prevMonthButton"),
  nextMonthButton: document.querySelector("#nextMonthButton"),
  todayButton: document.querySelector("#todayButton"),
  ideaButton: document.querySelector("#ideaButton"),
  ideaButtonInline: document.querySelector("#ideaButtonInline"),
  todoButton: document.querySelector("#todoButton"),
  todoButtonInline: document.querySelector("#todoButtonInline"),
  cpDialog: document.querySelector("#cpDialog"),
  closeCpDialog: document.querySelector("#closeCpDialog"),
  cpForm: document.querySelector("#cpForm"),
  cpDateTbd: document.querySelector('input[name="dateTbd"]'),
  cpStartDate: document.querySelector('input[name="startDate"]'),
  cpEndDate: document.querySelector('input[name="endDate"]'),
  ideaDialog: document.querySelector("#ideaDialog"),
  closeDialog: document.querySelector("#closeDialog"),
  ideaForm: document.querySelector("#ideaForm"),
  exportIdeas: document.querySelector("#exportIdeas"),
  todoDialog: document.querySelector("#todoDialog"),
  closeTodoDialog: document.querySelector("#closeTodoDialog"),
  todoForm: document.querySelector("#todoForm"),
  detailDialog: document.querySelector("#detailDialog"),
  closeDetailDialog: document.querySelector("#closeDetailDialog"),
  detailKind: document.querySelector("#detailKind"),
  detailTitle: document.querySelector("#detailTitle"),
  detailBody: document.querySelector("#detailBody"),
  editDetailButton: document.querySelector("#editDetailButton"),
  deleteDetailButton: document.querySelector("#deleteDetailButton"),
};

function init() {
  fillProductOptions();
  bindEvents();
  loadIdeas();
  loadTodos();
  refreshSheetData();
}

function bindEvents() {
  els.refreshButton.addEventListener("click", refreshSheetData);
  els.prevMonthButton.addEventListener("click", () => {
    state.calendarDate = addMonths(state.calendarDate, -1);
    render();
  });
  els.nextMonthButton.addEventListener("click", () => {
    state.calendarDate = addMonths(state.calendarDate, 1);
    render();
  });
  els.todayButton.addEventListener("click", () => {
    state.calendarDate = startOfMonth(new Date());
    render();
  });
  [els.cpButton, els.cpButtonInline].filter(Boolean).forEach((button) => {
    button.addEventListener("click", () => els.cpDialog.showModal());
  });
  [els.ideaButton, els.ideaButtonInline].filter(Boolean).forEach((button) => {
    button.addEventListener("click", () => els.ideaDialog.showModal());
  });
  [els.todoButton, els.todoButtonInline].filter(Boolean).forEach((button) => {
    button.addEventListener("click", () => els.todoDialog.showModal());
  });
  els.closeCpDialog?.addEventListener("click", () => closeEntryDialog("cp"));
  els.closeDialog.addEventListener("click", () => closeEntryDialog("idea"));
  els.closeTodoDialog?.addEventListener("click", () => closeEntryDialog("todo"));
  els.cpDateTbd?.addEventListener("change", toggleCpDateFields);
  els.cpForm?.addEventListener("submit", saveCampaign);
  els.ideaForm.addEventListener("submit", saveIdea);
  els.todoForm?.addEventListener("submit", saveTodo);
  els.exportIdeas.addEventListener("click", exportIdeasAsCsv);
  els.activeCampaignList.addEventListener("click", handleDetailClick);
  els.upcomingCampaignList.addEventListener("click", handleDetailClick);
  els.unscheduledCampaignList?.addEventListener("click", handleDetailClick);
  els.monthlyCalendar.addEventListener("click", handleDetailClick);
  els.ideaGrid.addEventListener("click", handleDetailClick);
  els.todoGrid?.addEventListener("click", handleDetailClick);
  els.closeDetailDialog.addEventListener("click", () => els.detailDialog.close());
  els.editDetailButton.addEventListener("click", editSelectedDetail);
  els.deleteDetailButton.addEventListener("click", deleteSelectedDetail);
}

function closeEntryDialog(type) {
  if (type === "cp") {
    els.cpDialog.close();
    resetEntryForm("cp");
  } else if (type === "idea") {
    els.ideaDialog.close();
    resetEntryForm("idea");
  } else if (type === "todo") {
    els.todoDialog.close();
    resetEntryForm("todo");
  }
}

function refreshSheetData() {
  loadCampaigns();
  loadSheetIdeas();
  loadSheetTodos();
}

function toggleCpDateFields() {
  const isTbd = Boolean(els.cpDateTbd?.checked);
  [els.cpStartDate, els.cpEndDate].forEach((input) => {
    if (!input) return;
    input.required = !isTbd;
    input.disabled = isTbd;
    if (isTbd) input.value = "";
  });
}

async function loadCampaigns() {
  setConnection("loading", "Googleシートを確認中");
  try {
    const rows = await loadSheetRows();
    if (!hasCampaignHeaders(rows[0] || [])) throw new Error("Expected campaign headers not found");
    const campaigns = dedupeCampaigns(rowsToCampaigns(rows));
    state.campaigns = campaigns;
    setConnection("ok", `Googleシート接続中：${campaigns.length}件`);
  } catch (error) {
    state.campaigns = sampleCampaigns;
    setConnection("warn", "シート未接続：サンプル表示中");
    console.warn("Sheet load failed", error);
  }
  render();
}

async function loadSheetRows() {
  try {
    const cpRows = await loadRowsWithJsonp(CP_SHEET_NAME);
    if (!hasCampaignHeaders(cpRows[0] || [])) throw new Error("CP sheet has no campaign headers");
    return cpRows;
  } catch (error) {
    console.warn("CP sheet load failed, trying default sheet", error);
    try {
      const defaultRows = await loadRowsWithJsonp();
      if (!hasCampaignHeaders(defaultRows[0] || [])) throw new Error("Default sheet has no campaign headers");
      return defaultRows;
    } catch (fallbackError) {
      console.warn("JSONP sheet load failed, trying CSV fetch", fallbackError);
      const response = await fetch(`${SHEET_CSV_URL}&cacheBust=${Date.now()}`, {
        cache: "no-store",
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return parseCsv(await response.text());
    }
  }
}

function loadRowsWithJsonp(sheetName = "") {
  return new Promise((resolve, reject) => {
    const callbackName = `__myceeSheet${Date.now()}${Math.round(Math.random() * 1000)}`;
    const script = document.createElement("script");
    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("Google Sheet request timed out"));
    }, 8000);

    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (response) => {
      cleanup();
      if (response?.status === "error") {
        reject(new Error(response.errors?.[0]?.detailed_message || "Google Sheet returned an error"));
        return;
      }
      resolve(tableToRows(response.table));
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Google Sheet script failed to load"));
    };
    const sheetParam = sheetName ? `&sheet=${encodeURIComponent(sheetName)}` : "";
    script.src = `${SHEET_JSONP_URL}?tqx=out:json;responseHandler:${callbackName}${sheetParam}&tq=${encodeURIComponent("select *")}&cacheBust=${Date.now()}`;
    document.head.append(script);
  });
}

function tableToRows(table) {
  const headers = (table?.cols || []).map((column) => cleanCell(column.label || column.id));
  const rows = (table?.rows || []).map((row) => (row.c || []).map(cellToValue));
  if (hasKnownHeaders(headers)) {
    return [headers, ...rows].filter((row) => row.some((value) => cleanCell(value)));
  }
  if (hasKnownHeaders(rows[0] || [])) {
    return rows.filter((row) => row.some((value) => cleanCell(value)));
  }
  return [headers, ...rows].filter((row) => row.some((value) => cleanCell(value)));
}

function hasKnownHeaders(headerRow) {
  return hasCampaignHeaders(headerRow) || hasIdeaHeaders(headerRow) || hasTodoHeaders(headerRow);
}

function cellToValue(cell) {
  if (!cell) return "";
  return cleanCell(cell.f ?? cell.v ?? "");
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"' && inQuotes && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      if (row.some((value) => value.trim() !== "")) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  if (cell || row.length) {
    row.push(cell);
    if (row.some((value) => value.trim() !== "")) rows.push(row);
  }
  return rows;
}

function rowsToCampaigns(rows) {
  if (rows.length < 1) return [];
  const headers = rows[0].map(normalizeHeader);
  if (!hasCampaignHeaders(rows[0])) return [];
  return rows
    .slice(1)
    .map((row, index) => {
      const get = (...keys) => {
        const normalizedKeys = keys.map(normalizeHeader);
        const index = headers.findIndex((header) => normalizedKeys.includes(header));
        return index >= 0 ? cleanCell(row[index]) : "";
      };
      return {
        name: get("CP名", "キャンペーン名", "CP", "Campaign"),
        product: get("商材名", "商材", "商品", "Product") || "その他",
        startDate: normalizeDate(get("開始日", "Start", "Start Date")),
        endDate: normalizeDate(get("終了日", "End", "End Date")),
        dateTbd: /未定|tbd/i.test([get("開始日", "Start", "Start Date"), get("終了日", "End", "End Date")].join(" ")),
        detail: get("内容・詳細", "内容", "詳細", "Detail", "Description"),
        channel: get("対象チャネル", "チャネル", "Channel"),
        owner: get("担当者", "担当", "Owner"),
        status: get("ステータス", "Status"),
        cpUrl: get("CP URL", "CPURL", "CPリンク", "Campaign URL"),
        otherUrl: get("その他URL", "その他 URL", "Other URL", "Reference URL"),
        needs: get("確認事項", "不足", "不足項目", "Needs", "Blocker"),
        source: "sheet",
        sheetRow: index + 2,
        id: `cp-${index + 2}-${get("CP名", "キャンペーン名", "CP", "Campaign")}`,
      };
    })
    .filter((campaign) => [campaign.name, campaign.product, campaign.startDate, campaign.endDate, campaign.detail, campaign.channel, campaign.owner, campaign.status, campaign.cpUrl, campaign.otherUrl, campaign.needs].some((value) => String(value).trim()));
}

function dedupeCampaigns(campaigns) {
  const seen = new Set();
  return [...campaigns].reverse().filter((campaign) => {
    const key = [
      campaign.name,
      campaign.product,
      campaign.startDate,
      campaign.endDate,
      campaign.owner,
    ]
      .map((value) => cleanCell(value))
      .join("|");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).reverse();
}

function hasCampaignHeaders(headerRow) {
  const headers = headerRow.map(normalizeHeader);
  const expectedHeaders = ["CP名", "商材名", "開始日", "終了日", "内容・詳細", "対象チャネル", "担当者", "ステータス"].map(normalizeHeader);
  return headers.some((header) => expectedHeaders.includes(header));
}

function rowsToIdeas(rows) {
  if (rows.length < 1) return [];
  const headers = rows[0].map(normalizeHeader);
  if (!hasIdeaHeaders(rows[0])) return [];
  return rows
    .slice(1)
    .map((row, index) => {
      const get = (...keys) => {
        const normalizedKeys = keys.map(normalizeHeader);
        const headerIndex = headers.findIndex((header) => normalizedKeys.includes(header));
        return headerIndex >= 0 ? cleanCell(row[headerIndex]) : "";
      };
      return {
        id: `idea-${index + 2}-${get("アイデア名", "Idea")}`,
        title: get("アイデア名", "Idea", "Title"),
        product: get("商材名", "商材", "商品", "Product") || "その他",
        detail: get("内容・詳細", "内容", "詳細", "Detail", "Description"),
        writer: get("記入者", "Writer", "Author"),
        url: get("URL", "Url", "リンク", "Link"),
        source: "sheet",
        syncStatus: "synced",
        sheetRow: index + 2,
      };
    })
    .filter((idea) => [idea.title, idea.product, idea.detail, idea.writer, idea.url].some((value) => String(value).trim()));
}

function hasIdeaHeaders(headerRow) {
  const headers = headerRow.map(normalizeHeader);
  const expectedHeaders = ["アイデア名", "商材名", "内容・詳細", "記入者", "URL"].map(normalizeHeader);
  return headers.some((header) => expectedHeaders.includes(header));
}

function rowsToTodos(rows) {
  if (rows.length < 1) return [];
  const headers = rows[0].map(normalizeHeader);
  if (!hasTodoHeaders(rows[0])) return [];
  return rows
    .slice(1)
    .map((row, index) => {
      const get = (...keys) => {
        const normalizedKeys = keys.map(normalizeHeader);
        const headerIndex = headers.findIndex((header) => normalizedKeys.includes(header));
        return headerIndex >= 0 ? cleanCell(row[headerIndex]) : "";
      };
      return {
        id: `todo-${index + 2}-${get("やりたいこと", "Want", "ToDo", "Todo")}`,
        title: get("やりたいこと", "Want", "ToDo", "Todo", "Title"),
        product: get("商材名", "商材", "商品", "Product") || "その他",
        detail: get("詳細", "Detail", "Description"),
        writer: get("記入者", "Writer", "Author"),
        source: "sheet",
        syncStatus: "synced",
        sheetRow: index + 2,
      };
    })
    .filter((todo) => [todo.title, todo.product, todo.detail, todo.writer].some((value) => String(value).trim()));
}

function hasTodoHeaders(headerRow) {
  const headers = headerRow.map(normalizeHeader);
  const expectedHeaders = ["やりたいこと", "商材名", "詳細", "記入者"].map(normalizeHeader);
  return headers.some((header) => expectedHeaders.includes(header));
}

function normalizeHeader(value) {
  return cleanCell(value).replace(/\s+/g, "").toLowerCase();
}

function cleanCell(value) {
  return String(value ?? "").trim();
}

function normalizeDate(value) {
  if (!value) return "";
  const cleaned = String(value).replace(/\./g, "/").replace(/年|月/g, "/").replace(/日/g, "").trim();
  const parsed = new Date(cleaned);
  if (Number.isNaN(parsed.getTime())) return value;
  return toLocalIsoDate(parsed);
}

function toLocalIsoDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function render() {
  const campaigns = state.campaigns;
  renderSpotlightCampaigns(campaigns);
  renderMonthlyCalendar(campaigns);
  renderIdeas();
  renderTodos();
}

function renderSpotlightCampaigns(campaigns) {
  const activeCampaigns = campaigns.filter((campaign) => !isUnscheduledCampaign(campaign) && getStatus(campaign).type === "active");
  const upcomingCampaigns = campaigns
    .filter((campaign) => !isUnscheduledCampaign(campaign) && getStatus(campaign).type === "upcoming")
    .sort((a, b) => (toDate(a.startDate) || new Date(8640000000000000)) - (toDate(b.startDate) || new Date(8640000000000000)));
  const unscheduledCampaigns = campaigns.filter(isUnscheduledCampaign);

  els.activeCampaignList.innerHTML = activeCampaigns.length
    ? renderLimitedItems(activeCampaigns, "active", (campaign) => renderCampaignCard(campaign, "active"))
    : `<div class="empty-state">進行中のCPはありません。</div>`;
  els.upcomingCampaignList.innerHTML = upcomingCampaigns.length
    ? renderLimitedItems(upcomingCampaigns, "upcoming", (campaign) => renderCampaignCard(campaign, "upcoming"))
    : `<div class="empty-state">今後の予定はありません。</div>`;
  if (els.unscheduledCampaignList) {
    els.unscheduledCampaignList.innerHTML = unscheduledCampaigns.length
      ? renderLimitedItems(unscheduledCampaigns, "unscheduled", (campaign) => renderCampaignCard(campaign, "upcoming"))
      : `<div class="empty-state">日程未定のCPはありません。</div>`;
  }
}

function renderLimitedItems(items, key, renderer) {
  const visibleItems = state.expanded[key] ? items : items.slice(0, 3);
  return `${visibleItems.map(renderer).join("")}${renderMoreButton(key, items.length)}`;
}

function renderMoreButton(key, total) {
  if (total <= 3) return "";
  return `
    <div class="more-row">
      <button class="secondary-button more-button" type="button" data-expand-key="${escapeHtml(key)}">
        ${state.expanded[key] ? "閉じる" : "その他"}
      </button>
    </div>
  `;
}

function renderCampaignCard(campaign, tone = "active") {
  const product = getProduct(campaign.product);
  const progress = getDateProgress(campaign);
  return `
    <article class="campaign-card is-${tone}" data-detail-type="cp" data-detail-id="${escapeHtml(campaign.id)}" style="--product-color: ${product.color}">
      <div class="product-band">${escapeHtml(campaign.product || "その他")}</div>
      <div class="campaign-card-body">
        <div class="campaign-title-row">
          <h3>${escapeHtml(campaign.name || "")}</h3>
        </div>
        <div class="campaign-meta">
          ${renderTextTag("担当", campaign.owner)}
          ${renderTextTag("チャネル", campaign.channel)}
          ${renderUrlTag(campaign.cpUrl, "CP URL")}
          ${renderUrlTag(campaign.otherUrl, "その他URL")}
          ${campaign.syncStatus ? `<span class="tag">${escapeHtml(getSyncLabel(campaign))}</span>` : ""}
          ${campaign.needs ? `<span class="tag">確認 ${escapeHtml(campaign.needs)}</span>` : ""}
        </div>
      </div>
      <div class="campaign-side">
        <div class="date-range">${formatDateRange(campaign)}</div>
        <div class="progress-track" aria-label="期間の進捗">
          <span class="progress-fill" style="--progress: ${progress}%"></span>
        </div>
      </div>
    </article>
  `;
}

function renderUrlTag(url, label) {
  if (!url) return "";
  const href = escapeHtml(url);
  return `<a class="tag tag-link" href="${href}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)}</a>`;
}

function renderTextTag(label, value) {
  if (!cleanCell(value)) return "";
  const text = label ? `${label} ${value}` : value;
  return `<span class="tag">${escapeHtml(text)}</span>`;
}

function renderMonthlyCalendar(campaigns) {
  const monthStart = startOfMonth(state.calendarDate);
  const monthEnd = endOfMonth(monthStart);
  const firstGridDate = addDays(monthStart, -monthStart.getDay());
  const weeks = Array.from({ length: 6 }, (_, index) => addDays(firstGridDate, index * 7));
  const datedCampaigns = campaigns.filter((campaign) => !isUnscheduledCampaign(campaign) && (campaign.startDate || campaign.endDate));

  els.calendarLabel.textContent = `${monthStart.getFullYear()}年${monthStart.getMonth() + 1}月`;
  els.monthlyCalendar.innerHTML = `
    ${["日", "月", "火", "水", "木", "金", "土"].map((day) => `<div class="calendar-weekday">${day}</div>`).join("")}
    ${weeks.map((weekStart) => renderCalendarWeek(weekStart, monthStart, datedCampaigns)).join("")}
  `;

  if (!datedCampaigns.some((campaign) => isCampaignInRange(campaign, monthStart, monthEnd))) {
    els.monthlyCalendar.insertAdjacentHTML("beforeend", `<div class="calendar-empty">この月に表示できるCPはありません</div>`);
  }
}

function renderCalendarWeek(weekStart, monthStart, campaigns) {
  const maxLanes = 4;
  const weekEnd = addDays(weekStart, 6);
  const days = Array.from({ length: 7 }, (_, index) => addDays(weekStart, index));
  const segments = campaigns
    .map((campaign) => getCalendarSegment(campaign, weekStart, weekEnd))
    .filter(Boolean)
    .sort((a, b) => a.columnStart - b.columnStart || b.columnSpan - a.columnSpan);
  const arrangedSegments = arrangeCalendarSegments(segments, maxLanes);

  return `
    <div class="calendar-week">
      ${days
        .map(
          (day, index) => `
            <div class="calendar-day-bg ${isSameMonth(day, monthStart) ? "" : "is-muted"} ${isSameDate(day, new Date()) ? "is-today" : ""}" style="--calendar-column: ${index + 1}">
              <span class="calendar-date">${day.getDate()}</span>
            </div>
          `,
        )
        .join("")}
      ${arrangedSegments.visible
        .map((segment) => {
          const product = getProduct(segment.campaign.product);
          const statusType = getStatus(segment.campaign).type;
          return `
            <span
              class="calendar-event is-${statusType}"
              data-detail-type="cp"
              data-detail-id="${escapeHtml(segment.campaign.id)}"
              style="--calendar-column: ${segment.columnStart}; --calendar-span: ${segment.columnSpan}; --calendar-row: ${segment.lane + 2}; --event-color: ${product.color}"
              title="${escapeHtml(`${segment.campaign.name || ""}｜${formatDateRange(segment.campaign)}`)}"
            >
              ${escapeHtml(segment.campaign.name || "")}
            </span>
          `;
        })
        .join("")}
      ${
        arrangedSegments.hiddenCount
          ? `
            <span class="calendar-more" style="--calendar-column: 1; --calendar-span: 7; --calendar-row: 6">
              ほか${arrangedSegments.hiddenCount}件
            </span>
          `
          : ""
      }
    </div>
  `;
}

function getCalendarSegment(campaign, weekStart, weekEnd) {
  const campaignStart = toDate(campaign.startDate) || toDate(campaign.endDate);
  const campaignEnd = toDate(campaign.endDate) || campaignStart;
  if (!campaignStart || !campaignEnd || campaignStart > weekEnd || campaignEnd < weekStart) return null;
  const start = campaignStart < weekStart ? weekStart : campaignStart;
  const end = campaignEnd > weekEnd ? weekEnd : campaignEnd;
  const columnStart = diffInDays(weekStart, start) + 1;
  const columnSpan = diffInDays(start, end) + 1;
  return {
    campaign,
    columnStart,
    columnSpan,
    columnEnd: columnStart + columnSpan - 1,
  };
}

function arrangeCalendarSegments(segments, maxLanes) {
  const laneEnds = [];
  const visible = [];
  let hiddenCount = 0;

  segments.forEach((segment) => {
    const lane = laneEnds.findIndex((endColumn) => endColumn < segment.columnStart);
    const assignedLane = lane >= 0 ? lane : laneEnds.length;
    if (assignedLane >= maxLanes) {
      hiddenCount += 1;
      return;
    }
    laneEnds[assignedLane] = segment.columnEnd;
    visible.push({ ...segment, lane: assignedLane });
  });

  return { visible, hiddenCount };
}

function diffInDays(start, end) {
  return Math.round((startOfDay(end) - startOfDay(start)) / 86400000);
}

function renderIdeas() {
  if (!state.ideas.length) {
    els.ideaGrid.innerHTML = `<div class="empty-state">まだアイデアがありません。「＋ アイデア」から蓄積できます。</div>`;
    return;
  }
  els.ideaGrid.innerHTML = renderLimitedItems(state.ideas, "ideas", (idea) => {
    const product = getProduct(idea.product);
    return `
      <article class="idea-card" data-detail-type="idea" data-detail-id="${escapeHtml(idea.id)}" style="border-top: 4px solid ${product.color}">
        <div>
          <strong>${escapeHtml(idea.title || "")}</strong>
        </div>
        <div>
          <div class="campaign-meta">
            ${renderTextTag("", idea.product)}
            ${renderTextTag("記入者", idea.writer)}
            ${renderUrlTag(idea.url, "URL")}
          </div>
          <div class="idea-footer">
            <span class="tag">${escapeHtml(getIdeaSyncLabel(idea))}</span>
          </div>
        </div>
      </article>
    `;
  });
}

function renderTodos() {
  if (!els.todoGrid) return;
  if (!state.todos.length) {
    els.todoGrid.innerHTML = `<div class="empty-state">まだやりたいことがありません。「＋ 追加」から蓄積できます。</div>`;
    return;
  }
  els.todoGrid.innerHTML = renderLimitedItems(
    state.todos,
    "todos",
    (todo) => {
      const product = getProduct(todo.product);
      return `
      <article class="todo-card" data-detail-type="todo" data-detail-id="${escapeHtml(todo.id)}" style="border-top: 4px solid ${product.color}">
        <div>
          <strong>${escapeHtml(todo.title || "")}</strong>
        </div>
        <div>
          <div class="campaign-meta">
            ${renderTextTag("", todo.product)}
            ${renderTextTag("記入者", todo.writer)}
          </div>
          <div class="idea-footer">
            <span class="tag">${escapeHtml(getSyncLabel(todo))}</span>
          </div>
        </div>
      </article>
    `;
    },
  );
}

function getIdeaSyncLabel(idea) {
  return getSyncLabel(idea);
}

function getSyncLabel(item) {
  if (item.syncStatus === "synced") return "シート同期済み";
  if (item.syncStatus === "failed") return "シート未同期";
  return "同期中";
}

function handleDetailClick(event) {
  const expandButton = event.target.closest("[data-expand-key]");
  if (expandButton) {
    const key = expandButton.dataset.expandKey;
    state.expanded[key] = !state.expanded[key];
    render();
    return;
  }
  if (event.target.closest("a, button")) return;
  const target = event.target.closest("[data-detail-type][data-detail-id]");
  if (!target) return;
  openDetail(target.dataset.detailType, target.dataset.detailId);
}

function openDetail(type, id) {
  const item = findDetailItem(type, id);
  if (!item) return;
  state.selectedDetail = { type, id };
  els.detailKind.textContent = type === "cp" ? "Campaign Detail" : type === "idea" ? "Idea Detail" : "Want Detail";
  els.detailTitle.textContent = type === "cp" ? item.name || "" : item.title || "";
  els.detailBody.innerHTML = type === "cp" ? renderCampaignDetail(item) : type === "idea" ? renderIdeaDetail(item) : renderTodoDetail(item);
  els.detailDialog.showModal();
}

function findDetailItem(type, id) {
  if (type === "cp") return state.campaigns.find((campaign) => campaign.id === id);
  if (type === "idea") return state.ideas.find((idea) => idea.id === id);
  if (type === "todo") return state.todos.find((todo) => todo.id === id);
  return null;
}

function renderCampaignDetail(campaign) {
  return `
    ${renderDetailRow("商材名", campaign.product)}
    ${renderDetailRow("開始日 / 終了日", formatDateRange(campaign))}
    ${renderDetailRow("対象チャネル", campaign.channel)}
    ${renderDetailRow("担当者", campaign.owner)}
    ${renderDetailRow("内容・詳細", campaign.detail)}
    ${renderDetailLink("CP URL", campaign.cpUrl)}
    ${renderDetailLink("その他URL", campaign.otherUrl)}
    ${renderDetailRow("確認事項", campaign.needs)}
  `;
}

function renderIdeaDetail(idea) {
  return `
    ${renderDetailRow("商材名", idea.product)}
    ${renderDetailRow("記入者", idea.writer)}
    ${renderDetailRow("内容・詳細", idea.detail)}
    ${renderDetailLink("URL", idea.url)}
  `;
}

function renderTodoDetail(todo) {
  return `
    ${renderDetailRow("商材名", todo.product)}
    ${renderDetailRow("詳細", todo.detail)}
    ${renderDetailRow("記入者", todo.writer)}
  `;
}

function renderDetailRow(label, value) {
  return `
    <div class="detail-row">
      <span>${escapeHtml(label)}</span>
      <p>${escapeHtml(value || "")}</p>
    </div>
  `;
}

function renderDetailLink(label, url) {
  if (!url) return renderDetailRow(label, "");
  const href = escapeHtml(url);
  return `
    <div class="detail-row">
      <span>${escapeHtml(label)}</span>
      <p><a href="${href}" target="_blank" rel="noopener noreferrer">${href}</a></p>
    </div>
  `;
}

async function deleteSelectedDetail() {
  if (!state.selectedDetail) return;
  if (!window.confirm("削除しますか")) return;
  const { type, id } = state.selectedDetail;
  const item = findDetailItem(type, id);
  if (!item) return;

  if (type === "cp") {
    state.campaigns = state.campaigns.filter((campaign) => campaign.id !== id);
  } else if (type === "idea") {
    state.ideas = state.ideas.filter((idea) => idea.id !== id);
    persistIdeas();
  } else if (type === "todo") {
    state.todos = state.todos.filter((todo) => todo.id !== id);
    persistTodos();
  }
  els.detailDialog.close();
  render();

  try {
    await syncDeleteToSheet(type, item);
  } catch (error) {
    console.warn("Delete sync failed", error);
  }
}

function getStatus(item) {
  const raw = String(item.status || "").trim();
  const today = startOfDay(new Date());
  const start = toDate(item.startDate);
  const end = toDate(item.endDate);
  const lowered = raw.toLowerCase();

  if (/完了|終了|done|closed/.test(lowered)) return { label: raw || "完了", type: "done" };
  if (/停止|保留|確認|待ち|不足|hold|stop|pending|blocked/.test(lowered)) return { label: raw || "確認中", type: "paused" };
  if (/アイデア|検討|idea/.test(lowered)) return { label: raw || "アイデア", type: "idea" };
  if (/予定|plan|upcoming/.test(lowered)) return { label: raw || "進行予定", type: "upcoming" };
  if (/進行中|active|running/.test(lowered)) return { label: raw || "進行中", type: "active" };
  if (start && start > today) return { label: raw || "進行予定", type: "upcoming" };
  if (start && end && start <= today && today <= end) return { label: raw || "進行中", type: "active" };
  if (!end || end >= today) return { label: raw || "進行中", type: "active" };
  return { label: raw || "終了", type: "done" };
}

function isUnscheduledCampaign(campaign) {
  return Boolean(campaign.dateTbd) || campaign.startDate === "未定" || campaign.endDate === "未定" || (!toDate(campaign.startDate) && !toDate(campaign.endDate));
}

function hasNeeds(item) {
  return Boolean(item.needs) || getMissingFields(item).length > 0;
}

function getMissingFields(item) {
  const missing = [];
  if (!item.name) missing.push("CP名");
  if (!item.product) missing.push("商材名");
  if (!item.startDate) missing.push("開始日");
  if (!item.endDate) missing.push("終了日");
  if (!item.owner) missing.push("担当者");
  return missing;
}

function getDateProgress(item) {
  const start = toDate(item.startDate);
  const end = toDate(item.endDate);
  if (!start || !end) return 0;
  const today = startOfDay(new Date());
  if (today <= start) return 0;
  if (today >= end) return 100;
  return Math.round(((today - start) / (end - start)) * 100);
}

function toDate(value) {
  if (!value || value === "未定") return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : startOfDay(date);
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function endOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function addMonths(date, amount) {
  return new Date(date.getFullYear(), date.getMonth() + amount, 1);
}

function addDays(date, amount) {
  const nextDate = new Date(date);
  nextDate.setDate(nextDate.getDate() + amount);
  return nextDate;
}

function isSameMonth(date, monthDate) {
  return date.getFullYear() === monthDate.getFullYear() && date.getMonth() === monthDate.getMonth();
}

function isSameDate(left, right) {
  return startOfDay(left).getTime() === startOfDay(right).getTime();
}

function isCampaignOnDate(campaign, date) {
  const start = toDate(campaign.startDate) || toDate(campaign.endDate);
  const end = toDate(campaign.endDate) || start;
  if (!start || !end) return false;
  const target = startOfDay(date);
  return start <= target && target <= end;
}

function isCampaignInRange(campaign, rangeStart, rangeEnd) {
  const start = toDate(campaign.startDate) || toDate(campaign.endDate);
  const end = toDate(campaign.endDate) || start;
  if (!start || !end) return false;
  return start <= rangeEnd && end >= rangeStart;
}

function formatDateRange(item) {
  if (item.dateTbd || item.startDate === "未定" || item.endDate === "未定") return "未定";
  const start = formatDate(item.startDate);
  const end = formatDate(item.endDate);
  if (start && end) return `${start} - ${end}`;
  return start || end || "";
}

function formatDate(value) {
  const date = toDate(value);
  if (!date) return value || "";
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

function loadIdeas() {
  try {
    state.ideas = normalizeIdeas(JSON.parse(localStorage.getItem(LOCAL_IDEAS_KEY) || "[]"));
  } catch {
    state.ideas = [];
  }
}

function loadTodos() {
  try {
    state.todos = normalizeTodos(JSON.parse(localStorage.getItem(LOCAL_TODOS_KEY) || "[]"));
  } catch {
    state.todos = [];
  }
}

async function loadSheetIdeas() {
  try {
    const rows = await loadRowsWithJsonp(IDEA_SHEET_NAME);
    if (!hasIdeaHeaders(rows[0] || [])) throw new Error("Idea sheet has no idea headers");
    state.ideas = mergeIdeas(rowsToIdeas(rows), state.ideas.filter((idea) => idea.source !== "sheet"));
    persistIdeas();
    render();
  } catch (error) {
    console.warn("Idea sheet load failed", error);
  }
}

async function loadSheetTodos() {
  try {
    const rows = await loadRowsWithJsonp(TODO_SHEET_NAME);
    if (!hasTodoHeaders(rows[0] || [])) throw new Error("Todo sheet has no todo headers");
    state.todos = mergeTodos(rowsToTodos(rows), state.todos.filter((todo) => todo.source !== "sheet"));
    persistTodos();
    render();
  } catch (error) {
    console.warn("Todo sheet load failed", error);
  }
}

function normalizeIdeas(ideas) {
  return ideas.map((idea) => ({
    ...idea,
    writer: idea.writer || idea.owner || "",
    url: idea.url || "",
  }));
}

function normalizeTodos(todos) {
  return todos.map((todo) => ({
    ...todo,
    title: todo.title || todo.name || "",
    product: todo.product || "その他",
    detail: todo.detail || "",
    writer: todo.writer || todo.owner || "",
  }));
}

function mergeIdeas(sheetIdeas, localIdeas) {
  const seen = new Set();
  return [...sheetIdeas, ...localIdeas].filter((idea) => {
    const key = [idea.title, idea.product, idea.detail, idea.writer, idea.url].map((value) => cleanCell(value)).join("|");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function mergeTodos(sheetTodos, localTodos) {
  const seen = new Set();
  return [...sheetTodos, ...localTodos].filter((todo) => {
    const key = [todo.title, todo.product, todo.detail, todo.writer].map((value) => cleanCell(value)).join("|");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function createLocalId() {
  return crypto.randomUUID ? crypto.randomUUID() : String(Date.now());
}

function editSelectedDetail() {
  if (!state.selectedDetail) return;
  const { type, id } = state.selectedDetail;
  const item = findDetailItem(type, id);
  if (!item) return;
  state.editingDetail = { type, id };
  els.detailDialog.close();
  if (type === "cp") {
    populateCampaignForm(item);
    els.cpDialog.showModal();
  } else if (type === "idea") {
    populateIdeaForm(item);
    els.ideaDialog.showModal();
  } else if (type === "todo") {
    populateTodoForm(item);
    els.todoDialog.showModal();
  }
}

function populateCampaignForm(campaign) {
  els.cpForm.querySelector(".dialog-header h2").textContent = "CPを編集";
  els.cpForm.querySelector('button[type="submit"]').textContent = "更新";
  setFormValue(els.cpForm, "name", campaign.name);
  setFormValue(els.cpForm, "product", campaign.product || "その他");
  setFormValue(els.cpForm, "channel", campaign.channel);
  setFormValue(els.cpForm, "startDate", campaign.startDate === "未定" ? "" : campaign.startDate);
  setFormValue(els.cpForm, "endDate", campaign.endDate === "未定" ? "" : campaign.endDate);
  setFormValue(els.cpForm, "owner", campaign.owner);
  setFormValue(els.cpForm, "detail", campaign.detail);
  setFormValue(els.cpForm, "cpUrl", campaign.cpUrl);
  setFormValue(els.cpForm, "otherUrl", campaign.otherUrl);
  setFormValue(els.cpForm, "needs", campaign.needs);
  if (els.cpDateTbd) els.cpDateTbd.checked = Boolean(campaign.dateTbd || campaign.startDate === "未定" || campaign.endDate === "未定");
  toggleCpDateFields();
}

function populateIdeaForm(idea) {
  els.ideaForm.querySelector(".dialog-header h2").textContent = "アイデアを編集";
  els.ideaForm.querySelector('button[type="submit"]').textContent = "更新";
  setFormValue(els.ideaForm, "title", idea.title);
  setFormValue(els.ideaForm, "product", idea.product || "その他");
  setFormValue(els.ideaForm, "writer", idea.writer);
  setFormValue(els.ideaForm, "detail", idea.detail);
  setFormValue(els.ideaForm, "url", idea.url);
}

function populateTodoForm(todo) {
  els.todoForm.querySelector(".dialog-header h2").textContent = "やりたいことを編集";
  els.todoForm.querySelector('button[type="submit"]').textContent = "更新";
  setFormValue(els.todoForm, "title", todo.title);
  setFormValue(els.todoForm, "product", todo.product || "その他");
  setFormValue(els.todoForm, "detail", todo.detail);
  setFormValue(els.todoForm, "writer", todo.writer);
}

function setFormValue(form, name, value) {
  const field = form.querySelector(`[name="${name}"]`);
  if (field) field.value = value || "";
}

function resetEntryForm(type) {
  state.editingDetail = null;
  if (type === "cp") {
    els.cpForm.reset();
    els.cpForm.querySelector(".dialog-header h2").textContent = "CPを追加";
    els.cpForm.querySelector('button[type="submit"]').textContent = "保存";
    toggleCpDateFields();
  } else if (type === "idea") {
    els.ideaForm.reset();
    els.ideaForm.querySelector(".dialog-header h2").textContent = "アイデアを追加";
    els.ideaForm.querySelector('button[type="submit"]').textContent = "保存";
  } else if (type === "todo") {
    els.todoForm.reset();
    els.todoForm.querySelector(".dialog-header h2").textContent = "やりたいことを追加";
    els.todoForm.querySelector('button[type="submit"]').textContent = "保存";
  }
}

async function saveCampaign(event) {
  event.preventDefault();
  const submitButton = els.cpForm.querySelector('button[type="submit"]');
  const editing = state.editingDetail?.type === "cp" ? state.editingDetail : null;
  const original = editing ? findDetailItem("cp", editing.id) : null;
  submitButton.disabled = true;
  submitButton.textContent = "保存中";
  const formData = new FormData(els.cpForm);
  const dateTbd = formData.get("dateTbd") === "on";
  const campaign = {
    id: original?.id || createLocalId(),
    name: cleanCell(formData.get("name")),
    product: cleanCell(formData.get("product")),
    startDate: dateTbd ? "未定" : normalizeDate(cleanCell(formData.get("startDate"))),
    endDate: dateTbd ? "未定" : normalizeDate(cleanCell(formData.get("endDate"))),
    dateTbd,
    detail: cleanCell(formData.get("detail")),
    channel: cleanCell(formData.get("channel")),
    owner: cleanCell(formData.get("owner")),
    status: "",
    cpUrl: cleanCell(formData.get("cpUrl")),
    otherUrl: cleanCell(formData.get("otherUrl")),
    needs: cleanCell(formData.get("needs")),
    source: original?.source || "local",
    sheetRow: original?.sheetRow || "",
    createdAt: original?.createdAt || new Date().toISOString(),
    syncStatus: "pending",
  };

  if (original) {
    state.campaigns = state.campaigns.map((item) => (item.id === original.id ? campaign : item));
  } else {
    state.campaigns.unshift(campaign);
  }
  els.cpForm.reset();
  toggleCpDateFields();
  els.cpDialog.close();
  resetEntryForm("cp");
  render();

  try {
    if (original) {
      await syncReplaceInSheet("cp", original, campaign);
    } else {
      await syncEntryToSheet("cp", campaign);
    }
    campaign.syncStatus = "synced";
    campaign.syncedAt = new Date().toISOString();
  } catch (error) {
    campaign.syncStatus = "failed";
    campaign.syncError = error.message;
    console.warn("CP sync failed", error);
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "保存";
    render();
  }
}

async function saveIdea(event) {
  event.preventDefault();
  const submitButton = els.ideaForm.querySelector('button[type="submit"]');
  const editing = state.editingDetail?.type === "idea" ? state.editingDetail : null;
  const original = editing ? findDetailItem("idea", editing.id) : null;
  submitButton.disabled = true;
  submitButton.textContent = "保存中";
  const formData = new FormData(els.ideaForm);
  const idea = {
    id: original?.id || createLocalId(),
    title: cleanCell(formData.get("title")),
    product: cleanCell(formData.get("product")),
    detail: cleanCell(formData.get("detail")),
    writer: cleanCell(formData.get("writer")),
    url: cleanCell(formData.get("url")),
    source: original?.source || "local",
    sheetRow: original?.sheetRow || "",
    createdAt: original?.createdAt || new Date().toISOString(),
    syncStatus: "pending",
  };
  if (original) {
    state.ideas = state.ideas.map((item) => (item.id === original.id ? idea : item));
  } else {
    state.ideas.unshift(idea);
  }
  persistIdeas();
  els.ideaForm.reset();
  els.ideaDialog.close();
  resetEntryForm("idea");
  render();

  try {
    if (original) {
      await syncReplaceInSheet("idea", original, idea);
    } else {
      await syncEntryToSheet("idea", idea);
    }
    idea.syncStatus = "synced";
    idea.syncedAt = new Date().toISOString();
  } catch (error) {
    idea.syncStatus = "failed";
    idea.syncError = error.message;
    console.warn("Idea sync failed", error);
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "保存";
    persistIdeas();
    render();
  }
}

async function saveTodo(event) {
  event.preventDefault();
  const submitButton = els.todoForm.querySelector('button[type="submit"]');
  const editing = state.editingDetail?.type === "todo" ? state.editingDetail : null;
  const original = editing ? findDetailItem("todo", editing.id) : null;
  submitButton.disabled = true;
  submitButton.textContent = "保存中";
  const formData = new FormData(els.todoForm);
  const todo = {
    id: original?.id || createLocalId(),
    title: cleanCell(formData.get("title")),
    product: cleanCell(formData.get("product")),
    detail: cleanCell(formData.get("detail")),
    writer: cleanCell(formData.get("writer")),
    source: original?.source || "local",
    sheetRow: original?.sheetRow || "",
    createdAt: original?.createdAt || new Date().toISOString(),
    syncStatus: "pending",
  };
  if (original) {
    state.todos = state.todos.map((item) => (item.id === original.id ? todo : item));
  } else {
    state.todos.unshift(todo);
  }
  persistTodos();
  els.todoForm.reset();
  els.todoDialog.close();
  resetEntryForm("todo");
  render();

  try {
    if (original) {
      await syncReplaceInSheet("todo", original, todo);
    } else {
      await syncEntryToSheet("todo", todo);
    }
    todo.syncStatus = "synced";
    todo.syncedAt = new Date().toISOString();
  } catch (error) {
    todo.syncStatus = "failed";
    todo.syncError = error.message;
    console.warn("Todo sync failed", error);
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "保存";
    persistTodos();
    render();
  }
}

async function syncEntryToSheet(type, item) {
  const payload = type === "cp" ? campaignToPayload(item) : type === "idea" ? ideaToPayload(item) : todoToPayload(item);
  await sendGasPayload(payload);
}

async function syncUpdateToSheet(type, oldItem, newItem) {
  const payload = {
    action: "update",
    type,
    sheetName: type === "cp" ? CP_SHEET_NAME : type === "idea" ? IDEA_SHEET_NAME : TODO_SHEET_NAME,
    rowNumber: oldItem.sheetRow || "",
    key: type === "cp" ? campaignToPayload(oldItem) : type === "idea" ? ideaToPayload(oldItem) : todoToPayload(oldItem),
    values: type === "cp" ? campaignToPayload(newItem) : type === "idea" ? ideaToPayload(newItem) : todoToPayload(newItem),
  };
  await sendGasPayload(payload);
}

async function syncReplaceInSheet(type, oldItem, newItem) {
  await syncDeleteToSheet(type, oldItem);
  await wait(700);
  await syncEntryToSheet(type, newItem);
  newItem.sheetRow = "";
  newItem.source = "local";
}

async function syncDeleteToSheet(type, item) {
  const payload = {
    action: "delete",
    type,
    sheetName: type === "cp" ? CP_SHEET_NAME : type === "idea" ? IDEA_SHEET_NAME : TODO_SHEET_NAME,
    rowNumber: item.sheetRow || "",
    key: type === "cp" ? campaignToPayload(item) : type === "idea" ? ideaToPayload(item) : todoToPayload(item),
  };
  await sendGasPayload(payload);
}

function wait(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

async function sendGasPayload(payload) {
  await fetch(GAS_URL, {
    method: "POST",
    mode: "no-cors",
    headers: {
      "Content-Type": "text/plain;charset=utf-8",
    },
    body: JSON.stringify(payload),
  });
}

function campaignToPayload(campaign) {
  return {
    type: "cp",
    sheetName: CP_SHEET_NAME,
    name: campaign.name,
    title: campaign.name,
    product: campaign.product,
    startDate: campaign.startDate,
    endDate: campaign.endDate,
    dateTbd: campaign.dateTbd,
    detail: campaign.detail,
    channel: campaign.channel,
    owner: campaign.owner,
    status: campaign.status,
    cpUrl: campaign.cpUrl,
    otherUrl: campaign.otherUrl,
    needs: campaign.needs,
    createdAt: campaign.createdAt,
  };
}

function ideaToPayload(idea) {
  return {
    type: "idea",
    sheetName: IDEA_SHEET_NAME,
    title: idea.title,
    name: idea.title,
    product: idea.product,
    detail: idea.detail,
    writer: idea.writer,
    url: idea.url,
  };
}

function todoToPayload(todo) {
  return {
    type: "todo",
    sheetName: TODO_SHEET_NAME,
    title: todo.title,
    name: todo.title,
    product: todo.product,
    detail: todo.detail,
    writer: todo.writer,
  };
}

function persistIdeas() {
  localStorage.setItem(LOCAL_IDEAS_KEY, JSON.stringify(state.ideas));
}

function persistTodos() {
  localStorage.setItem(LOCAL_TODOS_KEY, JSON.stringify(state.todos));
}

function ideaToCampaign(idea) {
  return {
    name: idea.title,
    product: idea.product,
    detail: idea.detail,
    channel: "",
    owner: idea.writer,
    status: "アイデア",
    cpUrl: "",
    otherUrl: idea.url,
    needs: "",
    source: "idea",
  };
}

function exportIdeasAsCsv() {
  const headers = ["アイデア名", "商材名", "内容・詳細", "記入者", "URL"];
  const rows = state.ideas.map((idea) => [
    idea.title,
    idea.product,
    idea.detail,
    idea.writer,
    idea.url,
  ]);
  const csv = [headers, ...rows].map((row) => row.map(csvEscape).join(",")).join("\n");
  const blob = new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `mycee-ideas-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
}

function fillProductOptions() {
  [els.ideaForm, els.cpForm, els.todoForm].filter(Boolean).forEach((form) => {
    const productSelect = form.querySelector('select[name="product"]');
    if (!productSelect) return;
    productSelect.innerHTML = productCatalog.map((product) => `<option>${escapeHtml(product.name)}</option>`).join("");
  });
}

function setConnection(type, message) {
  els.connectionDot.className = `pulse-dot ${type === "ok" ? "ok" : type === "warn" ? "warn" : ""}`;
  els.connectionStatus.textContent = message;
}

function getProduct(productName) {
  return productCatalog.find((product) => product.name === productName) || productCatalog.at(-1);
}

function countBy(items, key) {
  const counts = new Map();
  items.forEach((item) => {
    const value = item[key] || "";
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1]);
}

function unique(values) {
  return [...new Set(values)].sort((a, b) => String(a).localeCompare(String(b), "ja"));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

init();
