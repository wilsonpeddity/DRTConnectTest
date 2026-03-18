const themeToggle = document.getElementById("themeToggle");
const themeLabel = themeToggle.querySelector(".theme-label");
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
const searchResults = document.getElementById("searchResults");
const secondaryAction = document.getElementById("secondaryAction");
const primaryAction = document.getElementById("primaryAction");
const navActions = document.querySelectorAll(".nav-action");
const createNewBtn = document.querySelector(".create-new-btn");
const reportsToggle = document.getElementById("reportsToggle");
const reportsSection = document.querySelector(".reports-section");
const tabPills = document.querySelectorAll(".tab-pill");
const drtRowsContainer = document.getElementById("drtRows");
const unreadUpdatesList = document.getElementById("unreadUpdatesList");
const sortableHeaders = document.querySelectorAll(".sortable");
const savedTheme = window.localStorage.getItem("drt-theme");

const CURRENT_USER = "Wilson Peddity";
let currentQueue = "highPriority";
let searchQuery = "";
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
    switch (queue) {
      case "highPriority": return row.priority === "High" && !isClosed;
      case "assigned": return row.assigned && !isClosed;
      case "myDrt": return row.owner === CURRENT_USER && !isClosed;
      case "allOpen": return !isClosed;
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
    switch (currentQueue) {
      case "highPriority": return row.priority === "High" && !isClosed;
      case "assigned": return row.assigned && !isClosed;
      case "myDrt": return row.owner === CURRENT_USER && !isClosed;
      case "allOpen": return !isClosed;
      case "closed": return isClosed;
      default: return true;
    }
  });
}

function getSearchFilteredData(data) {
  if (!searchQuery.trim()) return data;
  const q = searchQuery.trim().toLowerCase();
  return data.filter((row) => {
    const searchStr = [row.id, row.priority, row.account, row.aov, row.owner, row.status, row.timeToTarget].join(" ").toLowerCase();
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
  `;
  tr.addEventListener("click", () => {
    window.location.href = "load-drt.html?id=" + encodeURIComponent(row.id);
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
  unreadUpdatesList.innerHTML = "";
  data.forEach((update) => {
    const li = document.createElement("li");
    li.className = "unread-update-item" + (update.completed ? " is-completed" : "");
    li.dataset.id = update.id;
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = !!update.completed;
    checkbox.setAttribute("aria-label", "Mark as complete");
    const content = document.createElement("span");
    content.className = "unread-update-content";
    const dateTime = update.createdAt || "";
    const prefix = dateTime ? dateTime + " – " : "";
    const displayText = prefix + (update.text || "");
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
  const dateTime = update.createdAt || "";
  const prefix = dateTime ? dateTime + " – " : "";
  return prefix + (update.text || "");
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

function renderUpdatesModal() {
  if (!updatesModalBody) return;
  const data = getSortedUpdatesForModal();
  updatesModalBody.innerHTML = "";
  data.forEach((update) => {
    const tr = document.createElement("tr");
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

function applyTheme(theme) {
  document.body.dataset.theme = theme;
  const isNight = theme === "night";
  themeLabel.textContent = isNight ? "Night Mode" : "Day Mode";
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
  button.addEventListener("click", () => openModal(button.dataset.modal));
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
