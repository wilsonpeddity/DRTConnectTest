const themeToggle = document.getElementById("themeToggle");
const themeLabel = themeToggle.querySelector(".theme-label");
const modalBackdrop = document.getElementById("modalBackdrop");
const modalClose = document.getElementById("modalClose");
const modalEyebrow = document.getElementById("modalEyebrow");
const modalTitle = document.getElementById("modalTitle");
const modalCopy = document.getElementById("modalCopy");
const modalInput = document.getElementById("modalInput");
const modalSecondary = document.getElementById("modalSecondary");
const searchResults = document.getElementById("searchResults");
const secondaryAction = document.getElementById("secondaryAction");
const primaryAction = document.getElementById("primaryAction");
const navActions = document.querySelectorAll(".nav-action");
const tabPills = document.querySelectorAll(".tab-pill");
const drtRows = document.querySelectorAll("#drtRows tr");
const savedTheme = window.localStorage.getItem("drt-theme");

const modalConfigs = {
  create: {
    eyebrow: "Create",
    title: "Create New DRT",
    copy: "Capture the request, set the right priority, and route it into the DRT flow without leaving the dashboard.",
    primary: "Create DRT",
    secondary: "Cancel",
    placeholder: "Enter account, opportunity, or request name",
    notes: "Add business context, stakeholders, and required approvals",
    searchMode: false
  },
  priority: {
    eyebrow: "Priority",
    title: "High Priority Queue",
    copy: "Review urgent DRTs that need pricing, approvals, or leadership attention in the next few hours.",
    primary: "Open Queue",
    secondary: "Close",
    placeholder: "Filter by account or owner",
    notes: "List escalation notes or handoff details",
    searchMode: false
  },
  reports: {
    eyebrow: "Reports",
    title: "Generate Summary Report",
    copy: "Package pipeline movement, discount exceptions, and close-rate trends into a single manager-ready view.",
    primary: "Generate Report",
    secondary: "Cancel",
    placeholder: "Choose scope: region, team, or account",
    notes: "Add any metrics that need to be highlighted",
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
  }
};

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

  modalEyebrow.textContent = config.eyebrow;
  modalTitle.textContent = config.title;
  modalCopy.textContent = config.copy;
  modalInput.placeholder = config.placeholder;
  modalSecondary.placeholder = config.notes;
  primaryAction.textContent = config.primary;
  secondaryAction.textContent = config.secondary;
  searchResults.hidden = !config.searchMode;
  modalBackdrop.hidden = false;
  modalInput.value = "";
  modalSecondary.value = "";
  filterRows("");
  modalInput.focus();
}

function closeModal() {
  modalBackdrop.hidden = true;
}

function filterRows(query) {
  const normalized = query.trim().toLowerCase();

  drtRows.forEach((row) => {
    const haystack = row.dataset.search.toLowerCase();
    const matches = !normalized || haystack.includes(normalized);
    row.classList.toggle("is-hidden", !matches);
  });
}

themeToggle.addEventListener("click", () => {
  const nextTheme = document.body.dataset.theme === "day" ? "night" : "day";
  applyTheme(nextTheme);
});

navActions.forEach((button) => {
  button.addEventListener("click", () => openModal(button.dataset.modal));
});

tabPills.forEach((button) => {
  button.addEventListener("click", () => {
    tabPills.forEach((pill) => pill.classList.remove("is-selected"));
    button.classList.add("is-selected");
  });
});

modalClose.addEventListener("click", closeModal);
secondaryAction.addEventListener("click", () => {
  if (modalEyebrow.textContent === "Search") {
    modalInput.value = "";
    modalSecondary.value = "";
    filterRows("");
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
