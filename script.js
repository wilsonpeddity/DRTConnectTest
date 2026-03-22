const themeToggle = document.getElementById("themeToggle");
const themeLabel = themeToggle?.querySelector(".theme-label");
const modalBackdrop = document.getElementById("modalBackdrop");
const modalClose = document.getElementById("modalClose");
const modalEyebrow = document.getElementById("modalEyebrow");
const modalTitle = document.getElementById("modalTitle");
const modalCopy = document.getElementById("modalCopy");
const modalInput = document.getElementById("modalInput");
const modalSecondary = document.getElementById("modalSecondary");
const modalLayout = document.getElementById("modalLayout");
const createDrtModalContent = document.getElementById("createDrtModalContent");
const updatesModalContent = document.getElementById("updatesModalContent");
const updatesModalBody = document.getElementById("updatesModalBody");
const updatesSearchInput = document.getElementById("updatesSearchInput");
const searchResults = document.getElementById("searchResults");
const secondaryAction = document.getElementById("secondaryAction");
const primaryAction = document.getElementById("primaryAction");
const navActions = document.querySelectorAll(".nav-action");
const createNewBtn = document.querySelector(".create-new-btn");
const reportsToggle = document.getElementById("reportsToggle");
const reportsSection = document.querySelector(".reports-section");
const tabPills = document.querySelectorAll(".tab-pill");
const drtRowsContainer = document.getElementById("drtRows");
const queueSearchInput = document.getElementById("queueSearchInput");
const unreadUpdatesList = document.getElementById("unreadUpdatesList");
const sortableHeaders = document.querySelectorAll(".sortable");
const savedTheme = window.localStorage.getItem("drt-theme");

const CURRENT_USER = "Wilson Peddity";
let currentQueue = "highPriority";
let searchQuery = "";
let updatesSearchQuery = "";
let sortColumn = null;
let sortDirection = 1;

const modalConfigs = {
  create: {
    eyebrow: "Create",
    title: "Create New DRT",
    copy: "Capture the request, set the right priority, and route it into the DRT flow without leaving the dashboard.",
    primary: "Close",
    secondary: "Cancel",
    placeholder: "Enter account, opportunity, or request name",
    notes: "Add business context, stakeholders, and required approvals",
    searchMode: false
  },
  updates: {
    eyebrow: "Updates",
    title: "Deal Desk Updates",
    copy: "View and manage your unread deal desk notifications, pricing approvals, and workflow updates.",
    primary: "Close",
    secondary: "Cancel",
    placeholder: "Filter updates",
    notes: "Mark items as read or take action",
    searchMode: false
  },
  settings: {
    eyebrow: "Settings",
    title: "Workspace Settings",
    copy: "Adjust notifications, review thresholds, and dashboard defaults for the team operating model.",
    primary: "Save Settings",
    secondary: "Cancel",
    placeholder: "Set default workspace or notification rule",
    notes: "Capture any admin notes or policy changes",
    searchMode: false
  },
  search: {
    eyebrow: "Search",
    title: "Search DRT Workspace",
    copy: "Search live DRT records by ID, owner, account, or status. Matching rows update directly in the table below.",
    primary: "Apply Search",
    secondary: "Reset",
    placeholder: "Try Apex, Wilson, pricing, or DRT-1042",
    notes: "Optional notes for what you are looking for",
    searchMode: true
  }
};

function getQueueCount(queue) {
  const data = window.DRT_QUEUE_DATA || [];
  return data.filter((row) => {
    const isClosed = ["Closed Won", "Closed Lost"].includes(row.status);
    const isDraft = row.status === "Draft";
    switch (queue) {
      case "highPriority": return row.priority === "High" && !isClosed && !isDraft;
      case "assigned": return row.assigned && !isClosed && !isDraft;
      case "myDrt": return row.owner === CURRENT_USER && !isClosed && !isDraft;
      case "allOpen": return !isClosed && !isDraft;
      case "drafts": return isDraft;
      case "closed": return isClosed;
      default: return true;
    }
  }).length;
}

function updateTabCounts() {
  document.querySelectorAll(".tab-count").forEach((el) => {
    const queue = el.dataset.queue;
    if (queue) el.textContent = getQueueCount(queue);
  });
}

function getQueueFilteredData() {
  const data = window.DRT_QUEUE_DATA || [];
  return data.filter((row) => {
    const isClosed = ["Closed Won", "Closed Lost"].includes(row.status);
    const isDraft = row.status === "Draft";
    switch (currentQueue) {
      case "highPriority": return row.priority === "High" && !isClosed && !isDraft;
      case "assigned": return row.assigned && !isClosed && !isDraft;
      case "myDrt": return row.owner === CURRENT_USER && !isClosed && !isDraft;
      case "allOpen": return !isClosed && !isDraft;
      case "drafts": return isDraft;
      case "closed": return isClosed;
      default: return true;
    }
  });
}

function getSearchFilteredData(data) {
  if (!searchQuery.trim()) return data;
  const q = searchQuery.trim().toLowerCase();
  return data.filter((row) => {
    const searchStr = [row.id, row.account, row.owner].join(" ").toLowerCase();
    return searchStr.includes(q);
  });
}

function getSortedData(data) {
  if (!sortColumn) return data;
  const key = sortColumn;
  const dir = sortDirection;
  return [...data].sort((a, b) => {
    let va = a[key];
    let vb = b[key];
    if (key === "id") {
      va = parseInt((va || "").replace(/\D/g, ""), 10) || 0;
      vb = parseInt((vb || "").replace(/\D/g, ""), 10) || 0;
      return dir * (va - vb);
    }
    if (key === "priority") {
      const order = { High: 3, Medium: 2, Low: 1 };
      return dir * ((order[va] || 0) - (order[vb] || 0));
    }
    if (key === "aov") {
      va = parseFloat((va || "0").replace(/[^0-9.]/g, "")) || 0;
      vb = parseFloat((vb || "0").replace(/[^0-9.]/g, "")) || 0;
      return dir * (va - vb);
    }
    if (key === "created" && (a.createdTimestamp || b.createdTimestamp)) {
      va = new Date(a.createdTimestamp || 0).getTime();
      vb = new Date(b.createdTimestamp || 0).getTime();
      return dir * (va - vb);
    }
    if (key === "timeToTarget") {
      const order = { "Overdue": 0, "Today": 1, "1 day left": 2, "2 days left": 3, "3 days left": 4, "4 days left": 5, "5 days left": 6, "6 days left": 7, "7 days left": 8, "—": 99 };
      return dir * ((order[va] ?? 9) - (order[vb] ?? 9));
    }
    if (typeof va === "string" && typeof vb === "string") {
      return dir * va.localeCompare(vb);
    }
    return dir * (String(va).localeCompare(String(vb)));
  });
}

function priorityClass(p) {
  const c = (p || "").toLowerCase();
  return c === "high" ? "high" : c === "medium" ? "medium" : "low";
}

function timeToTargetClass(val) {
  if (!val || val === "—") return "time-na";
  const v = (val || "").toLowerCase();
  if (v === "overdue") return "time-overdue";
  if (v === "today") return "time-urgent";
  if (v === "1 day left") return "time-urgent";
  if (v === "2 days left" || v === "3 days left") return "time-warning";
  if (v === "4 days left" || v === "5 days left" || v === "6 days left" || v === "7 days left" || v.includes("week")) return "time-ok";
  return "time-ok";
}

function formatCreatedDate(createdTimestamp) {
  if (!createdTimestamp) return "";
  const created = new Date(createdTimestamp);
  const now = new Date();
  const diffMs = now - created;
  const diffHours = diffMs / (1000 * 60 * 60);
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  if (diffDays < 1) {
    if (diffHours < 1) {
      const mins = Math.floor(diffMs / (1000 * 60));
      return mins <= 1 ? "Just now" : `${mins} minutes ago`;
    }
    const hours = Math.floor(diffHours);
    return hours === 1 ? "1 hour ago" : `${hours} hours ago`;
  }
  const mon = created.toLocaleString("en-US", { month: "short" });
  const day = created.getDate();
  const hrs = created.getHours();
  const mins = created.getMinutes();
  const time = `${String(hrs).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;
  return `${mon} ${day}, ${time}`;
}

function getNextDrtId() {
  const data = window.DRT_QUEUE_DATA || [];
  let max = 0;
  data.forEach((r) => {
    const n = parseInt((r.id || "").replace(/\D/g, ""), 10) || 0;
    if (n > max) max = n;
  });
  return "DRT-" + (max + 1);
}

function duplicateDrt(row) {
  const data = window.DRT_QUEUE_DATA || [];
  const copy = { ...row };
  copy.id = getNextDrtId();
  copy.createdTimestamp = new Date().toISOString();
  copy.lastModified = "Just now";
  copy.status = "Draft";
  copy.assigned = false;
  copy.timeToTarget = "�";
  data.unshift(copy);
  renderTable();
}

function renderRow(row) {
  const tr = document.createElement("tr");
  tr.dataset.id = row.id;
  tr.style.cursor = "pointer";
  const createdDisplay = formatCreatedDate(row.createdTimestamp);
  const timeToTarget = row.timeToTarget || "—";
  const timeClass = timeToTargetClass(timeToTarget);
  tr.innerHTML = `
    <td>${row.id}</td>
    <td><span class="status-chip ${priorityClass(row.priority)}">${row.priority}</span></td>
    <td>${row.account}</td>
    <td>${row.aov || ""}</td>
    <td>${row.owner}</td>
    <td>${createdDisplay}</td>
    <td>${row.lastModified}</td>
    <td><span class="time-to-target ${timeClass}">${timeToTarget}</span></td>
    <td>${row.status}</td>
    <td class="col-actions">
      <div class="row-menu-wrap">
        <button type="button" class="row-menu-btn" aria-label="Row actions" aria-haspopup="true" aria-expanded="false">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="1"/><circle cx="12" cy="5" r="1"/><circle cx="12" cy="19" r="1"/></svg>
        </button>
        <div class="row-menu-dropdown" role="menu" hidden>
          <button type="button" class="row-menu-item" data-action="duplicate" role="menuitem">Duplicate</button>
        </div>
      </div>
    </td>
  `;
  tr.addEventListener("click", (e) => {
    if (e.target.closest(".row-menu-wrap")) return;
    window.location.href = "load-drt.html?id=" + encodeURIComponent(row.id);
  });
  const menuBtn = tr.querySelector(".row-menu-btn");
  const dropdown = tr.querySelector(".row-menu-dropdown");
  const duplicateBtn = tr.querySelector('[data-action="duplicate"]');
  menuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const isOpen = !dropdown.hidden;
    document.querySelectorAll(".row-menu-dropdown").forEach((d) => { d.hidden = true; });
    document.querySelectorAll(".row-menu-btn").forEach((b) => b.setAttribute("aria-expanded", "false"));
    if (!isOpen) {
      dropdown.hidden = false;
      menuBtn.setAttribute("aria-expanded", "true");
    }
  });
  duplicateBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    dropdown.hidden = true;
    menuBtn.setAttribute("aria-expanded", "false");
    duplicateDrt(row);
  });
  return tr;
}

function renderTable() {
  if (!drtRowsContainer) return;
  let data = getQueueFilteredData();
  data = getSearchFilteredData(data);
  data = getSortedData(data);
  drtRowsContainer.innerHTML = "";
  data.forEach((row) => drtRowsContainer.appendChild(renderRow(row)));
  updateSortArrows();
  updateTabCounts();
}

function updateSortArrows() {
  sortableHeaders.forEach((th) => {
    const arrow = th.querySelector(".sort-arrow");
    if (!arrow) return;
    const col = th.dataset.sort;
    th.classList.toggle("sort-active", sortColumn === col);
    arrow.classList.toggle("sort-desc", sortColumn === col && sortDirection === -1);
  });
}

function getUnreadUpdates() {
  const data = window.ALL_UPDATES_DATA || [];
  return data.filter((u) => !u.completed);
}

function renderUnreadUpdates() {
  if (!unreadUpdatesList) return;
  const data = getUnreadUpdates();
  const countEl = document.getElementById("unreadUpdatesCount");
  if (countEl) {
    countEl.textContent = data.length > 0 ? ` (${data.length})` : "";
    countEl.setAttribute("aria-label", data.length > 0 ? `${data.length} unread` : "No unread");
  }
  unreadUpdatesList.innerHTML = "";
  data.forEach((update) => {
    const li = document.createElement("li");
    li.id = "unread-update-" + update.id;
    li.className = "unread-update-item" + (update.completed ? " is-completed" : "");
    li.dataset.id = update.id;
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = !!update.completed;
    checkbox.setAttribute("aria-label", "Mark as complete");
    const content = document.createElement("span");
    content.className = "unread-update-content";
    const idPrefix = update.ticketId ? update.ticketId + " · " : "";
    const dateTime = update.createdAt || "";
    const prefix = dateTime ? dateTime + " – " : "";
    const displayText = idPrefix + prefix + (update.text || "");
    const text = document.createElement("span");
    text.className = "unread-update-text";
    text.textContent = displayText;
    content.appendChild(text);
    li.appendChild(checkbox);
    li.appendChild(content);
    checkbox.addEventListener("change", () => {
      update.completed = checkbox.checked;
      update.readOn = update.completed ? new Date().toLocaleString("en-US", { timeZone: "America/Los_Angeles", month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" }) + " PST" : null;
      li.classList.toggle("is-completed", update.completed);
      if (update.completed) {
        setTimeout(() => renderUnreadUpdates(), 600);
      }
    });
    unreadUpdatesList.appendChild(li);
  });
}

function getUpdateDisplayText(update) {
  const idPrefix = update.ticketId ? update.ticketId + " · " : "";
  const dateTime = update.createdAt || "";
  const prefix = dateTime ? dateTime + " – " : "";
  return idPrefix + prefix + (update.text || "");
}

function copyUpdateToClipboard(update, evt) {
  const text = getUpdateDisplayText(update);
  navigator.clipboard.writeText(text).then(() => {
    const btn = evt?.target?.closest(".copy-update-btn");
    if (btn) {
      const orig = btn.textContent;
      btn.textContent = "Copied!";
      btn.disabled = true;
      setTimeout(() => { btn.textContent = orig; btn.disabled = false; }, 1500);
    }
  });
}

function escapeHtml(s) {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}

function markUpdateAsRead(update, evt) {
  update.completed = true;
  update.readOn = new Date().toLocaleString("en-US", { timeZone: "America/Los_Angeles", month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" }) + " PST";

  const row = evt?.target?.closest("tr");
  if (row) {
    row.classList.remove("update-unread");
    row.classList.add("update-read");
    const statusCell = row.querySelector(".update-cell-status");
    const readOnCell = row.querySelector(".update-cell-readon");
    const actionsCell = row.querySelector(".update-cell-actions");
    if (statusCell) statusCell.innerHTML = '<span class="update-status-badge status-read">Read</span>';
    if (readOnCell) readOnCell.textContent = update.readOn;
    const markReadBtn = row.querySelector(".mark-read-btn");
    if (markReadBtn) markReadBtn.remove();
  }

  setTimeout(() => {
    renderUpdatesModal();
    renderUnreadUpdates();
  }, 600);
}

function parseUpdateDate(str) {
  if (!str) return 0;
  const cleaned = (str || "").replace(/\s*PST\s*$/i, "").trim();
  const d = new Date(cleaned);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function getSortedUpdatesForModal() {
  const data = window.ALL_UPDATES_DATA || [];
  const unread = data.filter((u) => !u.completed).sort((a, b) => parseUpdateDate(b.createdAt) - parseUpdateDate(a.createdAt));
  const read = data.filter((u) => u.completed).sort((a, b) => parseUpdateDate(b.readOn) - parseUpdateDate(a.readOn));
  return [...unread, ...read];
}

function getFilteredUpdatesForModal() {
  const query = (updatesSearchQuery || "").trim().toLowerCase();
  const data = getSortedUpdatesForModal();
  if (!query) return data;
  return data.filter((update) => {
    const haystack = [
      update.ticketId || "",
      update.text || "",
      update.createdAt || "",
      update.readOn || "",
    ].join(" ").toLowerCase();
    return haystack.includes(query);
  });
}

function renderUpdatesModal() {
  if (!updatesModalBody) return;
  const data = getFilteredUpdatesForModal();
  updatesModalBody.innerHTML = "";
  if (!data.length) {
    const tr = document.createElement("tr");
    tr.className = "update-empty-row";
    tr.innerHTML = `<td colspan="4" class="update-empty-cell">No DD Updates match this search.</td>`;
    updatesModalBody.appendChild(tr);
    return;
  }
  data.forEach((update) => {
    const tr = document.createElement("tr");
    tr.id = "update-" + update.id;
    tr.className = update.completed ? "update-read" : "update-unread";
    const status = update.completed ? "Read" : "Unread";
    const readOn = update.readOn || "—";
    const copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.className = "ghost-button copy-update-btn";
    copyBtn.textContent = "Copy to clipboard";
    copyBtn.addEventListener("click", (e) => copyUpdateToClipboard(update, e));
    const statusBadge = `<span class="update-status-badge ${update.completed ? "status-read" : "status-unread"}">${escapeHtml(status)}</span>`;
    tr.innerHTML = `
      <td class="update-cell-text">${escapeHtml(getUpdateDisplayText(update))}</td>
      <td class="update-cell-status">${statusBadge}</td>
      <td class="update-cell-readon">${escapeHtml(readOn)}</td>
      <td class="update-cell-actions"></td>
    `;
    const actionsCell = tr.querySelector(".update-cell-actions");
    if (!update.completed) {
      const markReadBtn = document.createElement("button");
      markReadBtn.type = "button";
      markReadBtn.className = "ghost-button mark-read-btn";
      markReadBtn.title = "mark as read";
      markReadBtn.setAttribute("aria-label", "mark as read");
      markReadBtn.innerHTML = `<span class="mark-read-icon" aria-hidden="true"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg></span>`;
      markReadBtn.addEventListener("click", (e) => markUpdateAsRead(update, e));
      actionsCell.appendChild(markReadBtn);
    }
    actionsCell.appendChild(copyBtn);
    updatesModalBody.appendChild(tr);
  });
}

if (updatesSearchInput) {
  updatesSearchInput.addEventListener("input", (event) => {
    updatesSearchQuery = event.target.value || "";
    renderUpdatesModal();
  });
}

function applyTheme(theme) {
  document.body.dataset.theme = theme;
  const isNight = theme === "night";
  if (themeLabel) {
    themeLabel.textContent = isNight ? "Night Mode" : "Day Mode";
  }
  window.localStorage.setItem("drt-theme", theme);
  themeToggle.setAttribute(
    "aria-label",
    isNight ? "Switch to day theme" : "Switch to night theme"
  );
}

function openModal(type) {
  const config = modalConfigs[type];
  if (!config) return;

  navActions.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.modal === type);
  });

  const isUpdatesModal = type === "updates";
  const isCreateModal = type === "create";
  const modal = modalBackdrop?.querySelector(".modal");
  if (modal) modal.classList.toggle("modal-updates", isUpdatesModal);
  if (modal) modal.classList.toggle("modal-create-drt", isCreateModal);
  if (modalLayout) modalLayout.hidden = isUpdatesModal || isCreateModal;
  if (createDrtModalContent) createDrtModalContent.hidden = !isCreateModal;
  if (updatesModalContent) updatesModalContent.hidden = !isUpdatesModal;

  modalEyebrow.textContent = config.eyebrow;
  modalTitle.textContent = config.title;
  modalCopy.textContent = config.copy;
  if (modalCopy) modalCopy.hidden = isUpdatesModal || isCreateModal;
  modalInput.placeholder = config.placeholder;
  modalSecondary.placeholder = config.notes;
  primaryAction.textContent = config.primary;
  secondaryAction.textContent = config.secondary;
  if (secondaryAction) secondaryAction.hidden = isUpdatesModal || isCreateModal;
  searchResults.hidden = !config.searchMode || isUpdatesModal || isCreateModal;

  if (isUpdatesModal) {
    if (updatesSearchInput) {
      updatesSearchInput.value = updatesSearchQuery;
    }
    renderUpdatesModal();
  } else if (isCreateModal) {
    const accountSearch = document.getElementById("createDrtAccountSearch");
    if (accountSearch) accountSearch.value = "";
  } else {
    modalInput.value = "";
    modalSecondary.value = "";
    searchQuery = "";
    renderTable();
    modalInput.focus();
  }

  modalBackdrop.hidden = false;
}

function closeModal() {
  modalBackdrop.hidden = true;
}

function filterRows(query) {
  searchQuery = query;
  renderTable();
}

themeToggle.addEventListener("click", () => {
  const nextTheme = document.body.dataset.theme === "day" ? "night" : "day";
  applyTheme(nextTheme);
});

navActions.forEach((button) => {
  button.addEventListener("click", () => {
    if (button.dataset.modal === "reports") {
      openReportsModal();
    } else {
      openModal(button.dataset.modal);
    }
  });
});

if (createNewBtn) {
  createNewBtn.addEventListener("click", () => openModal("create"));
}

const createDrtScratchBtn = document.querySelector(".create-drt-scratch-btn");
if (createDrtScratchBtn) {
  createDrtScratchBtn.addEventListener("click", () => {
    closeModal();
    window.location.href = "create-drt.html";
  });
}

if (reportsToggle && reportsSection) {
  reportsToggle.addEventListener("click", () => {
    const isCollapsed = reportsSection.classList.toggle("collapsed");
    reportsToggle.setAttribute("aria-expanded", !isCollapsed);
  });
}

const sidebarToggle = document.getElementById("sidebarToggle");
const appShell = document.querySelector(".app-shell");

function toggleSidebar() {
  const isCollapsed = appShell.classList.toggle("sidebar-collapsed");
  if (sidebarToggle) {
    sidebarToggle.setAttribute("aria-label", isCollapsed ? "Show menu" : "Hide menu");
    sidebarToggle.setAttribute("aria-expanded", !isCollapsed);
  }
}

if (sidebarToggle) {
  sidebarToggle.addEventListener("click", toggleSidebar);
}

tabPills.forEach((button) => {
  button.addEventListener("click", () => {
    tabPills.forEach((pill) => pill.classList.remove("is-selected"));
    button.classList.add("is-selected");
    currentQueue = button.dataset.queue || "highPriority";
    renderTable();
  });
});

if (queueSearchInput) {
  queueSearchInput.addEventListener("input", () => {
    searchQuery = queueSearchInput.value || "";
    renderTable();
  });
}

sortableHeaders.forEach((th) => {
  th.addEventListener("click", () => {
    const col = th.dataset.sort;
    if (sortColumn === col) {
      sortDirection *= -1;
    } else {
      sortColumn = col;
      sortDirection = 1;
    }
    renderTable();
  });
});

modalClose.addEventListener("click", closeModal);
secondaryAction.addEventListener("click", () => {
  if (modalEyebrow.textContent === "Search") {
    modalInput.value = "";
    modalSecondary.value = "";
    searchQuery = "";
    renderTable();
    return;
  }
  closeModal();
});

primaryAction.addEventListener("click", closeModal);

modalBackdrop.addEventListener("click", (event) => {
  if (event.target === modalBackdrop) {
    closeModal();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !modalBackdrop.hidden) {
    closeModal();
  }
});

modalInput.addEventListener("input", () => {
  if (modalEyebrow.textContent === "Search") {
    filterRows(modalInput.value);
  }
});

searchResults.querySelectorAll(".search-result").forEach((button) => {
  button.addEventListener("click", () => {
    modalInput.value = button.textContent.split("·")[0].trim();
    filterRows(modalInput.value);
  });
});

applyTheme(savedTheme === "night" ? "night" : "day");
currentQueue = "highPriority";
tabPills.forEach((pill) => {
  pill.classList.toggle("is-selected", pill.dataset.queue === "highPriority");
});

renderTable();
updateTabCounts();
renderUnreadUpdates();

document.addEventListener("click", () => {
  document.querySelectorAll(".row-menu-dropdown").forEach((d) => { d.hidden = true; });
  document.querySelectorAll(".row-menu-btn").forEach((b) => b.setAttribute("aria-expanded", "false"));
});

// ── Reports Modal ─────────────────────────────────────────────────────────

const rptModal      = document.getElementById("rptModal");
const rptModalPanel = document.getElementById("rptModalPanel");

// Salesforce blue color palette
const _RPT_CP = {
  series: [{ hi: "#0176d3", lo: "#014486" }, { hi: "#1b96ff", lo: "#0176d3" }],
  lineA: "#0176d3", lineB: "#1b96ff",
  gradA: "rgba(1,118,211,0.90)", gradB: "rgba(27,150,255,0.65)",
  areaA: "rgba(1,118,211,0.18)", areaB: "rgba(1,118,211,0.00)",
  grid: "#f0f7ff", gridStrong: "#7ec0f7",
  axisLbl: "#94a3b8", peakLbl: "#014486",
};

const RPT_CHARTS = [
  {
    label: "Closed",
    info: "Number of deals closed in current Qtr. The bar charts show the count across past 5 qtrs.",
    fmt: (v) => String(Math.round(v)),
    quarterly: {
      subtitle: "FY27 Q2 · deals closed per quarter",
      values: [89, 104, 113, 118, 121, 126],
      labels: ["Q3 FY25", "Q4 FY25", "Q1 FY26", "Q2 FY26", "Q3 FY26", "Q4 FY26"],
    },
    monthly: {
      subtitle: "Mar FY26 · deals closed per month",
      values: [18, 22, 31, 19, 24, 28, 21, 26, 30, 20, 25, 32],
      labels: ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
    },
  },
  {
    label: "Avg. Discount %",
    info: "Average discount percentage applied across deals. 11.8% is 20.2% below the quarterly ceiling, indicating healthy discount control.",
    fmt: (v) => v.toFixed(1) + "%",
    quarterly: {
      subtitle: "Average discount applied per quarter",
      values: [14.2, 13.8, 13.1, 12.6, 12.2, 11.8],
      labels: ["Q3 FY25", "Q4 FY25", "Q1 FY26", "Q2 FY26", "Q3 FY26", "Q4 FY26"],
    },
    monthly: {
      subtitle: "Average discount applied per month",
      values: [12.9, 12.4, 11.9, 12.6, 12.1, 11.6, 12.3, 11.9, 11.4, 12.1, 11.7, 11.8],
      labels: ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
    },
  },
  {
    label: "Avg. Won Rate",
    info: "Percentage of opportunities that converted to closed won. 63% of qualified opportunities were successfully closed in the period.",
    fmt: (v) => Math.round(v) + "%",
    quarterly: {
      subtitle: "Conversion view · % of deals closed won",
      values: [54, 57, 59, 61, 62, 63],
      labels: ["Q3 FY25", "Q4 FY25", "Q1 FY26", "Q2 FY26", "Q3 FY26", "Q4 FY26"],
    },
    monthly: {
      subtitle: "% of deals closed won per month",
      values: [58, 61, 64, 59, 62, 65, 60, 63, 66, 61, 62, 63],
      labels: ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
    },
  },
  {
    label: "NNAOV/ACV",
    info: "Growth vs. base — How much new value (NNA) you're adding vs. existing contract value (ACV).\n\nRatio — The 1.28x is the NNAOV/ACV ratio: for every 1 of ACV, you are adding about 1.28 of new annual value.\n\nPipeline health — A ratio above 1.0 typically means you're adding more new value than your current base.",
    fmt: (v) => v.toFixed(2) + "x",
    quarterly: {
      subtitle: "Value mix · NNAOV / ACV ratio per quarter",
      values: [0.92, 1.01, 1.10, 1.18, 1.23, 1.28],
      labels: ["Q3 FY25", "Q4 FY25", "Q1 FY26", "Q2 FY26", "Q3 FY26", "Q4 FY26"],
    },
    monthly: {
      subtitle: "Value mix · NNAOV / ACV ratio per month",
      values: [1.05, 1.12, 1.18, 1.09, 1.16, 1.22, 1.14, 1.20, 1.25, 1.18, 1.24, 1.28],
      labels: ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
    },
  },
];

let _rptActiveChart = 0;
let _rptChartType   = "line";
let _rptPeriod      = "quarterly";
let _rptMaximized   = false;

function _rptBuildLineChart(values, labels, W, H, fmt) {
  const pad = { top: 48, left: 68, right: 32, bottom: 44 };
  const iW = W - pad.left - pad.right;
  const iH = H - pad.top - pad.bottom;
  const n   = values.length;
  const mx  = Math.max(...values);
  const mn  = Math.min(...values);
  const rng = mx - mn || mx || 1;
  const dMax = mx + rng * 0.15;
  const dMin = Math.max(0, mn - rng * 0.05);
  const span = dMax - dMin || 1;
  const uid  = "rptl" + Math.random().toString(36).slice(2, 8);
  const dotR = 5.5;

  const pts = values.map((v, i) => ({
    x: pad.left + (i / (n - 1)) * iW,
    y: pad.top + iH * (1 - (v - dMin) / span),
    v, lbl: labels[i],
  }));

  const gridLines = Array.from({ length: 5 }, (_, i) => {
    const y = pad.top + (i / 4) * iH;
    const v = dMax - (i / 4) * span;
    return `<line x1="${pad.left}" y1="${y.toFixed(1)}" x2="${(pad.left + iW).toFixed(1)}" y2="${y.toFixed(1)}" stroke="${_RPT_CP.grid}" stroke-width="1.2"/>
      <text x="${(pad.left - 8).toFixed(1)}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="${_RPT_CP.axisLbl}" font-family="system-ui,sans-serif">${fmt(v)}</text>`;
  }).join("");

  const linePtsStr = pts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(" L ");
  const areaPath = `M ${pts[0].x.toFixed(1)},${pts[0].y.toFixed(1)} L ${linePtsStr.slice(linePtsStr.indexOf(" L ") + 3)} L ${pts[n-1].x.toFixed(1)},${(pad.top + iH).toFixed(1)} L ${pts[0].x.toFixed(1)},${(pad.top + iH).toFixed(1)} Z`;

  const maxV = Math.max(...values);
  const dots = pts.map((p) => {
    const isPeak = p.v === maxV;
    const c = isPeak ? _RPT_CP.lineA : _RPT_CP.lineB;
    return `${isPeak ? `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${dotR + 9}" fill="none" stroke="${_RPT_CP.lineA}" stroke-opacity="0.16" stroke-width="1.5"/>` : ""}
      <circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${dotR}" fill="#fff" stroke="${c}" stroke-width="2.6"/>
      <text x="${p.x.toFixed(1)}" y="${(p.y - dotR - 9).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="700" fill="${isPeak ? _RPT_CP.peakLbl : _RPT_CP.axisLbl}" font-family="system-ui,sans-serif">${fmt(p.v)}</text>`;
  }).join("");

  const xLabels = pts.map((p) =>
    `<text x="${p.x.toFixed(1)}" y="${(H - 14).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="600" fill="${_RPT_CP.axisLbl}" font-family="system-ui,sans-serif">${p.lbl}</text>`
  ).join("");

  return `<defs>
    <linearGradient id="${uid}-area" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="${_RPT_CP.areaA}"/>
      <stop offset="100%" stop-color="${_RPT_CP.areaB}"/>
    </linearGradient>
    <linearGradient id="${uid}-line" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0%" stop-color="${_RPT_CP.gradA}"/>
      <stop offset="100%" stop-color="${_RPT_CP.gradB}"/>
    </linearGradient>
  </defs>
  <rect width="${W}" height="${H}" fill="transparent"/>
  ${gridLines}
  <path d="${areaPath}" fill="url(#${uid}-area)"/>
  <path d="M ${linePtsStr}" fill="none" stroke="url(#${uid}-line)" stroke-width="2.8" stroke-linejoin="round" stroke-linecap="round"/>
  ${dots}
  ${xLabels}
  <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${(pad.top + iH).toFixed(1)}" stroke="${_RPT_CP.grid}" stroke-width="1"/>
  <line x1="${pad.left}" y1="${(pad.top + iH).toFixed(1)}" x2="${(pad.left + iW).toFixed(1)}" y2="${(pad.top + iH).toFixed(1)}" stroke="${_RPT_CP.gridStrong}" stroke-width="1.5"/>`;
}

function _rptBuildBarChart(values, labels, W, H, fmt) {
  const pad = { top: 48, left: 68, right: 32, bottom: 44 };
  const iW  = W - pad.left - pad.right;
  const iH  = H - pad.top - pad.bottom;
  const n   = values.length;
  const mx  = Math.max(...values, 1);
  const barW = iW / n * 0.54;
  const gap  = iW / n;
  const uid  = "rptb" + Math.random().toString(36).slice(2, 8);

  const gridLines = Array.from({ length: 5 }, (_, i) => {
    const y = pad.top + (i / 4) * iH;
    const v = mx * (1 - i / 4);
    return `<line x1="${pad.left}" y1="${y.toFixed(1)}" x2="${(pad.left + iW).toFixed(1)}" y2="${y.toFixed(1)}" stroke="${_RPT_CP.grid}" stroke-width="1.2"/>
      <text x="${(pad.left - 8).toFixed(1)}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="${_RPT_CP.axisLbl}" font-family="system-ui,sans-serif">${fmt(v)}</text>`;
  }).join("");

  const maxV = Math.max(...values);
  const bars = values.map((v, i) => {
    const barH = (v / mx) * iH;
    const x    = pad.left + gap * i + (gap - barW) / 2;
    const y    = pad.top + iH - barH;
    const isPeak = v === maxV;
    return `<defs>
      <linearGradient id="${uid}-b${i}" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="${_RPT_CP.series[0].hi}"/>
        <stop offset="100%" stop-color="${_RPT_CP.series[0].lo}"/>
      </linearGradient>
    </defs>
    <rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${barH.toFixed(1)}" rx="6" fill="url(#${uid}-b${i})"/>
    ${isPeak ? `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${barH.toFixed(1)}" rx="6" fill="none" stroke="${_RPT_CP.series[0].hi}" stroke-width="1.5" stroke-opacity="0.5"/>` : ""}
    <text x="${(x + barW / 2).toFixed(1)}" y="${(y - 9).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="700" fill="${isPeak ? _RPT_CP.peakLbl : _RPT_CP.axisLbl}" font-family="system-ui,sans-serif">${fmt(v)}</text>
    <text x="${(x + barW / 2).toFixed(1)}" y="${(H - 14).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="600" fill="${_RPT_CP.axisLbl}" font-family="system-ui,sans-serif">${labels[i]}</text>`;
  }).join("");

  return `<rect width="${W}" height="${H}" fill="transparent"/>
  ${gridLines}
  ${bars}
  <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${(pad.top + iH).toFixed(1)}" stroke="${_RPT_CP.grid}" stroke-width="1"/>
  <line x1="${pad.left}" y1="${(pad.top + iH).toFixed(1)}" x2="${(pad.left + iW).toFixed(1)}" y2="${(pad.top + iH).toFixed(1)}" stroke="${_RPT_CP.gridStrong}" stroke-width="1.5"/>`;
}

function _rptRenderChart() {
  const svgEl = document.getElementById("rptChartSvg");
  if (!svgEl) return;
  const chart  = RPT_CHARTS[_rptActiveChart];
  const period = chart[_rptPeriod];
  svgEl.setAttribute("viewBox", "0 0 800 400");
  svgEl.innerHTML = _rptChartType === "bar"
    ? _rptBuildBarChart(period.values, period.labels, 800, 400, chart.fmt)
    : _rptBuildLineChart(period.values, period.labels, 800, 400, chart.fmt);
}

function _rptSyncHeader() {
  const chart  = RPT_CHARTS[_rptActiveChart];
  const period = chart[_rptPeriod];
  const titleEl  = document.getElementById("rptModalTitle");
  const subText  = document.getElementById("rptSubtitleText");
  if (titleEl)  titleEl.textContent  = chart.label;
  if (subText)  subText.textContent  = period.subtitle;
  rptModal.querySelectorAll("[data-rpt-chart]").forEach((btn) =>
    btn.classList.toggle("active", Number(btn.dataset.rptChart) === _rptActiveChart)
  );
  rptModal.querySelectorAll("[data-rpt-type]").forEach((btn) =>
    btn.classList.toggle("active", btn.dataset.rptType === _rptChartType)
  );
  rptModal.querySelectorAll("[data-rpt-period]").forEach((btn) =>
    btn.classList.toggle("active", btn.dataset.rptPeriod === _rptPeriod)
  );
}

function openReportsModal() {
  if (!rptModal) return;
  navActions.forEach((b) => b.classList.toggle("is-active", b.dataset.modal === "reports"));
  _rptSyncHeader();
  _rptRenderChart();
  rptModal.hidden = false;
}

function closeReportsModal() {
  if (!rptModal) return;
  _rptSetMaximized(false);
  rptModal.hidden = true;
  navActions.forEach((b) => b.classList.remove("is-active"));
}

function _rptSetMaximized(on) {
  _rptMaximized = on;
  rptModalPanel?.classList.toggle("rpt-modal__panel--maximized", on);
  const btn = document.getElementById("rptModalMaximize");
  if (btn) btn.setAttribute("aria-label", on ? "Restore chart size" : "Maximize chart");
}

if (rptModal) {
  rptModal.addEventListener("click", (e) => {
    if (!(e.target instanceof Element)) return;
    if (e.target === rptModal || e.target.closest("[data-rpt-close]")) {
      if (_rptMaximized) _rptSetMaximized(false);
      else closeReportsModal();
      return;
    }
    const chartTab = e.target.closest("[data-rpt-chart]");
    if (chartTab) {
      _rptActiveChart = Number(chartTab.dataset.rptChart);
      _rptSyncHeader();
      _rptRenderChart();
      return;
    }
    const typeBtn = e.target.closest("[data-rpt-type]");
    if (typeBtn) {
      _rptChartType = typeBtn.dataset.rptType;
      _rptSyncHeader();
      _rptRenderChart();
      return;
    }
    const periodBtn = e.target.closest("[data-rpt-period]");
    if (periodBtn) {
      _rptPeriod = periodBtn.dataset.rptPeriod;
      _rptSyncHeader();
      _rptRenderChart();
      return;
    }
  });
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && rptModal && !rptModal.hidden) {
    if (_rptMaximized) _rptSetMaximized(false);
    else closeReportsModal();
  }
});


