(function () {
  const params = new URLSearchParams(window.location.search);
  const drtId = params.get("id");

  const createdDateEl = document.getElementById("createDrtCreatedDate");
  if (createdDateEl) {
    createdDateEl.value = "Mar 18, 2025 10:30 AM PST";
  }

  const modifiedDateEl = document.getElementById("createDrtModifiedDate");
  if (modifiedDateEl) {
    modifiedDateEl.value = "Mar 22, 2025 2:15 PM PST";
  }

  const targetDateEl = document.getElementById("createDrtTargetDate");
  if (targetDateEl) {
    const y = 2025;
    const m = String(Math.floor(Math.random() * 12) + 1).padStart(2, "0");
    const d = String(Math.floor(Math.random() * 28) + 1).padStart(2, "0");
    targetDateEl.value = `${m}/${d}/${y}`;
  }

  const loadDrtIdEl = document.getElementById("createDrtId");
  if (loadDrtIdEl) {
    loadDrtIdEl.textContent = drtId || "DRT-" + String(Math.floor(Math.random() * 90000) + 10000);
  }

  const assignedToInput = document.getElementById("createDrtAssignedTo");
  if (assignedToInput) {
    const queueData = window.DRT_QUEUE_DATA || [];
    const row = drtId ? queueData.find((r) => r.id === drtId) : null;
    const names = ["Maya Chen", "James Wilson", "Wilson Peddity", "Sarah Lee", "William Peddity", "Noor Patel", "Ishan Rao"];
    assignedToInput.value = (row && row.owner) ? row.owner : names[Math.floor(Math.random() * names.length)];
  }

  const assignToInput = document.getElementById("createDrtAssignTo");
  if (assignToInput && !assignToInput.placeholder) {
    assignToInput.placeholder = "Search teammate...";
  }

  function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  function formatDollarK(value) {
    return `$${Math.round(value)}K`;
  }

  function formatHealthPct(value) {
    return `${Math.round(value)}%`;
  }

  function buildSummarySeries(base, variancePct) {
    const current = Math.round(base * (1 + ((Math.random() * 2 - 1) * variancePct)));
    const years = Array.from({ length: 5 }, (_, index) => {
      const growth = 1 + (index * (variancePct * 0.55));
      return Math.round(base * growth * (1 + ((Math.random() * 2 - 1) * (variancePct * 0.45))));
    });
    const total = years.reduce((sum, value) => sum + value, 0);
    return { current, years, total };
  }

  function buildHealthSeries() {
    const current = randInt(92, 120);
    const years = Array.from({ length: 5 }, () => randInt(92, 120));
    const total = Math.round(years.reduce((sum, value) => sum + value, 0) / years.length);
    return { current, years, total };
  }

  function applyLoadDrtDealSummary() {
    const summaryData = {
      aov: buildSummarySeries(randInt(120, 420), 0.18),
      acv: buildSummarySeries(randInt(180, 560), 0.2),
      nnaov: buildSummarySeries(randInt(90, 300), 0.22),
      pipeline: buildHealthSeries(),
    };

    const formatters = {
      aov: formatDollarK,
      acv: formatDollarK,
      nnaov: formatDollarK,
      pipeline: formatHealthPct,
    };

    document.querySelectorAll(".create-drt-kpi-card[data-kpi]").forEach((card) => {
      const key = card.getAttribute("data-kpi");
      const valueEl = card.querySelector(".create-drt-kpi-value");
      const row = summaryData[key];
      const format = formatters[key];
      if (row && valueEl && format) {
        valueEl.textContent = format(key === "pipeline" ? row.current : row.years[0]);
      }
    });

    document.querySelectorAll(".create-drt-collapsed-metric[data-metric]").forEach((metric) => {
      const key = metric.getAttribute("data-metric");
      const valueEl = metric.querySelector(".create-drt-collapsed-metric-value");
      const row = summaryData[key];
      const format = formatters[key];
      if (row && valueEl && format) {
        valueEl.textContent = format(key === "pipeline" ? row.current : row.years[0]);
      }
    });

    const summaryRows = {
      AOV: "aov",
      ACV: "acv",
      NNAOV: "nnaov",
      Health: "pipeline",
    };
    const currentDashRows = new Set(["acv", "nnaov", "pipeline"]);

    document.querySelectorAll(".create-drt-summary-table tbody tr").forEach((tr) => {
      const label = tr.querySelector(".create-drt-summary-metric-label")?.textContent?.trim();
      const key = summaryRows[label];
      const row = summaryData[key];
      const format = formatters[key];
      if (!row || !format) return;
      const cells = Array.from(tr.querySelectorAll("td")).slice(1);
      const values = [row.current, ...row.years, row.total];
      cells.forEach((cell, index) => {
        if (values[index] === undefined) return;
        if (index === 0 && currentDashRows.has(key)) {
          cell.textContent = "-";
          return;
        }
        cell.textContent = format(values[index]);
      });
    });
  }

  applyLoadDrtDealSummary();

  const resultPicker = document.getElementById("loadDrtResultPicker");
  const resultHidden = document.getElementById("loadDrtResult");
  const resultDisplay = document.getElementById("loadDrtResultDisplay");
  const resultText = resultDisplay ? resultDisplay.querySelector(".create-drt-result-text") : null;
  const resultOptions = resultPicker ? resultPicker.querySelectorAll(".create-drt-result-option") : [];
  const saveBtn = document.getElementById("loadDrtSaveBtn");
  const saveBtnLabel = saveBtn ? saveBtn.querySelector(".create-drt-save-btn-label") : null;

  const RESULT_META = {
    "":   { label: "-",  cls: "create-drt-result-pill--empty" },
    won:  { label: "Won",  cls: "create-drt-result-pill--won" },
    lost: { label: "Lost", cls: "create-drt-result-pill--lost" },
    na:   { label: "NA",   cls: "create-drt-result-pill--na" },
  };

  function applySaveButtonState() {
    if (!saveBtn || !saveBtnLabel || !resultHidden) return;
    const hasResult = !!resultHidden.value;
    saveBtn.classList.toggle("create-drt-save-btn--close", hasResult);
    saveBtnLabel.textContent = hasResult ? "Close" : "Save";
    if (hasResult) {
      saveBtn.disabled = false;
      saveBtn.setAttribute("aria-disabled", "false");
    }
  }

  function setResult(value) {
    const normalized = RESULT_META[value] ? value : "";
    const meta = RESULT_META[normalized];
    if (resultHidden) resultHidden.value = normalized;
    if (resultDisplay) {
      resultDisplay.className = `create-drt-result-pill ${meta.cls}`;
    }
    if (resultText) {
      resultText.textContent = meta.label;
    }
    resultOptions.forEach((opt) => {
      opt.setAttribute("aria-selected", String(opt.dataset.value === normalized));
    });
    if (typeof window.refreshDealInfoSaveState === "function") {
      window.refreshDealInfoSaveState();
    }
    applySaveButtonState();
  }

  function openResultPicker() {
    if (resultPicker) resultPicker.setAttribute("aria-expanded", "true");
  }

  function closeResultPicker() {
    if (resultPicker) resultPicker.setAttribute("aria-expanded", "false");
  }

  function toggleResultPicker() {
    if (!resultPicker) return;
    if (resultPicker.getAttribute("aria-expanded") === "true") closeResultPicker();
    else openResultPicker();
  }

  if (resultPicker) {
    resultPicker.addEventListener("click", (e) => {
      const opt = e.target.closest(".create-drt-result-option");
      if (opt) {
        setResult(opt.dataset.value);
        closeResultPicker();
        return;
      }
      toggleResultPicker();
    });

    resultPicker.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        toggleResultPicker();
      } else if (e.key === "Escape") {
        closeResultPicker();
      } else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        const vals = ["", "won", "lost", "na"];
        const cur = vals.indexOf(resultHidden ? resultHidden.value : "");
        const next = e.key === "ArrowDown"
          ? Math.min(cur + 1, vals.length - 1)
          : Math.max(cur - 1, 0);
        setResult(vals[next]);
      }
    });

    document.addEventListener("click", (e) => {
      if (!resultPicker.contains(e.target)) closeResultPicker();
    });

    setResult("");
  }

  if (saveBtn) {
    saveBtn.addEventListener("click", () => {
      if (resultHidden && resultHidden.value) {
        window.location.href = "index.html";
      }
    });
  }

  if (typeof window.refreshDealInfoSaveState === "function") {
    const baseRefreshDealInfoSaveState = window.refreshDealInfoSaveState;
    window.refreshDealInfoSaveState = function () {
      const result = baseRefreshDealInfoSaveState();
      applySaveButtonState();
      return result;
    };
  }

  if (typeof window.setDealInfoBaseline === "function") {
    window.setDealInfoBaseline();
  }
})();
