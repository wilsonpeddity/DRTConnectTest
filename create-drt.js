(function () {
  const SKU_NAMES = ["Enterprise License", "Cloud Storage Pro", "API Access Tier", "Support Package", "Premium Add-on", "Base Subscription", "Professional Services", "Training Module"];
  const YES_NO = ["Y", "N"];
  const CAPPED = ["Capped", "Uncapped"];
  const CLOUD = ["AWS", "Azure", "GCP", "On-Prem"];

  function random(arr) {
    return arr[Math.floor(Math.random() * arr.length)];
  }
  function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }
  function randPrice() {
    return (Math.random() * 5000 + 50).toFixed(2);
  }
  function randPct() {
    return (Math.random() * 40 + 5).toFixed(1);
  }
  function randDate() {
    const m = String(randInt(1, 12)).padStart(2, "0");
    const d = String(randInt(1, 28)).padStart(2, "0");
    const y = String(randInt(24, 26)).padStart(2, "0");
    return `${m}/${d}/${y}`;
  }

  function generateRow(id) {
    const grp1 = [randInt(1, 50), ...Array(5).fill(0).map(() => randInt(1, 100))];
    const grp2 = [randPrice(), ...Array(5).fill(0).map(() => randPrice())];
    const grp3 = grp2.map((p, i) => (parseFloat(p) * grp1[i]).toFixed(2));
    const newTotal = grp3.reduce((a, b) => a + parseFloat(b), 0).toFixed(2);
    return {
      id,
      sku: "SKU-" + randInt(1000, 9999),
      additionalComment: "",
      applySelaDiscount: random(YES_NO),
      cappedUncapped: random(CAPPED),
      selaStdQuote: random(["SELA", "Std Quote"]),
      contractNumber: "CNT-" + randInt(100000, 999999),
      contractEndDate: randDate(),
      orgIdTenantTig: "ORG-" + randInt(100, 999),
      skuName: random(SKU_NAMES),
      grp1Curr: grp1[0],
      grp1Yr: grp1.slice(1),
      grp2Curr: grp2[0],
      grp2Yr: grp2.slice(1),
      grp3Curr: grp3[0],
      grp3Yr: grp3.slice(1),
      newTotal,
      addonUnitPrice: randPrice(),
      cloudOfSku: random(CLOUD),
      eligibleSignatureAddon: random(YES_NO),
      listPriceMonthly: randPrice(),
      exitDiscount: randPct() + "%",
      includePctBasedPricing: random(YES_NO),
      exitColorGuidance: "Green",
      exitApprovalLevel: "L" + randInt(1, 4),
      selaEligibility: random(YES_NO),
      awsEligibility: random(YES_NO),
    };
  }

  function generateBlankRow(id) {
    const emptyArr = ["", "", "", "", ""];
    return {
      id,
      sku: "",
      additionalComment: "",
      applySelaDiscount: "",
      cappedUncapped: "",
      selaStdQuote: "",
      contractNumber: "",
      contractEndDate: "",
      orgIdTenantTig: "",
      skuName: "",
      grp1Curr: "",
      grp1Yr: [...emptyArr],
      grp2Curr: "",
      grp2Yr: [...emptyArr],
      grp3Curr: "",
      grp3Yr: [...emptyArr],
      newTotal: "",
      addonUnitPrice: "",
      cloudOfSku: "",
      eligibleSignatureAddon: "",
      listPriceMonthly: "",
      exitDiscount: "",
      includePctBasedPricing: "",
      exitColorGuidance: "",
      exitApprovalLevel: "",
      selaEligibility: "",
      awsEligibility: "",
    };
  }

  function colToLetter(c) {
    if (c <= 0) return "";
    let s = "";
    while (c > 0) {
      const r = (c - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      c = Math.floor((c - 1) / 26);
    }
    return s;
  }

  function letterToCol(s) {
    let n = 0;
    for (let i = 0; i < s.length; i++) {
      n = n * 26 + (s.toUpperCase().charCodeAt(i) - 64);
    }
    return n;
  }

  function getCellValue(ref) {
    const table = document.getElementById("createDrtTable");
    if (!table) return 0;
    const match = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match) return 0;
    const refUpper = ref.toUpperCase();
    const inputs = table.querySelectorAll("input[data-cell]");
    for (const el of inputs) {
      if (el.getAttribute("data-cell") === refUpper) {
        const val = el.value;
        const n = parseFloat(String(val).replace(/[^0-9.-]/g, ""));
        return isNaN(n) ? 0 : n;
      }
    }
    const cells = table.querySelectorAll("td[data-cell]");
    for (const el of cells) {
      if (el.getAttribute("data-cell") === refUpper) {
        const span = el.querySelector(".create-drt-cell-value");
        if (span) {
          const val = span.textContent;
          const n = parseFloat(String(val).replace(/[^0-9.-]/g, ""));
          return isNaN(n) ? 0 : n;
        }
      }
    }
    return 0;
  }

  function evaluateFormula(formula) {
    const s = formula.trim();
    if (!s.startsWith("=")) return null;
    const expr = s.slice(1).trim();
    try {
      const sumMatch = expr.match(/^SUM\s*\(\s*([A-Z]+\d+)\s*:\s*([A-Z]+\d+)\s*\)\s*$/i);
      if (sumMatch) {
        const [, start, end] = sumMatch;
        const sc = start.match(/([A-Z]+)(\d+)/i), sr = end.match(/([A-Z]+)(\d+)/i);
        if (!sc || !sr) return null;
        const colStart = letterToCol(sc[1]);
        const colEnd = letterToCol(sr[1]);
        const rowStart = parseInt(sc[2], 10);
        const rowEnd = parseInt(sr[2], 10);
        let sum = 0;
        for (let r = rowStart; r <= rowEnd; r++) {
          for (let c = colStart; c <= colEnd; c++) {
            sum += getCellValue(colToLetter(c) + r);
          }
        }
        return sum;
      }
      const avgMatch = expr.match(/^AVERAGE\s*\(\s*([A-Z]+\d+)\s*:\s*([A-Z]+\d+)\s*\)\s*$/i);
      if (avgMatch) {
        const [, start, end] = avgMatch;
        const sc = start.match(/([A-Z]+)(\d+)/i), sr = end.match(/([A-Z]+)(\d+)/i);
        if (!sc || !sr) return null;
        const colStart = letterToCol(sc[1]);
        const colEnd = letterToCol(sr[1]);
        const rowStart = parseInt(sc[2], 10);
        const rowEnd = parseInt(sr[2], 10);
        let sum = 0;
        let count = 0;
        for (let r = rowStart; r <= rowEnd; r++) {
          for (let c = colStart; c <= colEnd; c++) {
            sum += getCellValue(colToLetter(c) + r);
            count++;
          }
        }
        return count ? sum / count : 0;
      }
      const replaced = expr.replace(/([A-Z]+\d+)/gi, (m) => getCellValue(m));
      if (!/^[\d.\s+\-*/().]+$/.test(replaced)) return null;
      return Function("return (" + replaced + ")")();
    } catch {
      return null;
    }
  }

  function renderRow(data, rowIndex) {
    const tr = document.createElement("tr");
    tr.dataset.id = data.id;
    tr.dataset.row = rowIndex;
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "create-drt-delete-btn";
    delBtn.title = "Delete row";
    delBtn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>';
    delBtn.addEventListener("click", () => tr.remove());

    const inp = (v, col) => `<td data-cell="${colToLetter(col)}${rowIndex}"><input type="text" class="create-drt-cell-input" value="${v}" data-cell="${colToLetter(col)}${rowIndex}"></td>`;
    const ro = (v, col) => `<td class="create-drt-cell-readonly" data-cell="${colToLetter(col)}${rowIndex}"><span class="create-drt-cell-value">${v}</span></td>`;
    const calcCell = (col) => `<td class="col-calc" data-cell="${colToLetter(col)}${rowIndex}"><input type="text" class="create-drt-cell-input create-drt-calc-input" value="" data-formula="" data-cell="${colToLetter(col)}${rowIndex}"></td>`;
    let col = 6;
    const cells = [
      `<td class="col-row-num">${rowIndex}</td>`,
      "<td class='col-delete'></td>",
      ...Array.from({ length: 5 }, (_, i) => calcCell(i + 1)),
      ro(data.sku, col++),
      inp(data.additionalComment, col++),
      inp(data.applySelaDiscount, col++),
      inp(data.cappedUncapped, col++),
      inp(data.selaStdQuote, col++),
      inp(data.contractNumber, col++),
      inp(data.contractEndDate, col++),
      inp(data.orgIdTenantTig, col++),
      inp(data.skuName, col++),
      inp(data.grp1Curr, col++),
      ...data.grp1Yr.map((v) => inp(v, col++)),
      inp(data.grp2Curr, col++),
      ...data.grp2Yr.map((v) => inp(v, col++)),
      ro(data.grp3Curr, col++),
      ...data.grp3Yr.map((v) => ro(v, col++)),
      ro(data.newTotal, col++),
      inp(data.addonUnitPrice, col++),
      ro(data.cloudOfSku, col++),
      ro(data.eligibleSignatureAddon, col++),
      ro(data.listPriceMonthly, col++),
      ro(data.exitDiscount, col++),
      inp(data.includePctBasedPricing, col++),
      inp(data.exitColorGuidance, col++),
      inp(data.exitApprovalLevel, col++),
      ro(data.selaEligibility, col++),
      ro(data.awsEligibility, col++),
    ];
    tr.innerHTML = cells.join("");
    tr.querySelector(".col-delete").appendChild(delBtn);
    return tr;
  }

  const tbody = document.getElementById("createDrtBody");
  const table = document.getElementById("createDrtTable");
  const toggle = document.getElementById("inputStyleToggle");
  const toggleLabel = toggle?.querySelector(".toggle-label");

  const page = document.querySelector(".create-drt-page");
  const formulaBar = document.getElementById("excelFormulaBar");
  const formulaInput = document.getElementById("excelFormulaInput");

  function getNextRowIndex() {
    if (!tbody) return 1;
    const rows = tbody.querySelectorAll("tr[data-id]");
    let max = 0;
    rows.forEach((r) => {
      const id = parseInt(r.dataset.id, 10);
      if (!isNaN(id) && id > max) max = id;
    });
    return max + 1;
  }

  function addRow() {
    if (!tbody) return;
    const input = document.getElementById("addRowsCount");
    const count = input ? Math.max(1, parseInt(input.value, 10) || 1) : 1;
    for (let i = 0; i < count; i++) {
      const nextId = getNextRowIndex();
      const tr = renderRow(generateBlankRow(nextId), nextId);
      tbody.insertBefore(tr, tbody.firstChild);
    }
    if (input) input.value = "";
    updateAddRowsButton();
    if (table?.classList.contains("input-style-excel")) setupExcelCellHandlers();
  }

  function updateAddRowsButton() {
    const input = document.getElementById("addRowsCount");
    const btn = document.getElementById("addRowBtn");
    if (!input || !btn) return;
    const val = parseInt(input.value, 10);
    btn.disabled = !(Number.isInteger(val) && val >= 1);
  }

  if (tbody) {
    const isLoadDrt = !!document.getElementById("loadDrtId");
    const rowGenerator = isLoadDrt ? generateRow : generateBlankRow;
    for (let i = 1; i <= 50; i++) {
      tbody.appendChild(renderRow(rowGenerator(i), i));
    }
  }

  const addRowBtn = document.getElementById("addRowBtn");
  const addRowsInput = document.getElementById("addRowsCount");
  if (addRowBtn && tbody) {
    addRowBtn.addEventListener("click", addRow);
  }
  if (addRowsInput) {
    addRowsInput.addEventListener("input", updateAddRowsButton);
    addRowsInput.addEventListener("change", updateAddRowsButton);
  }

  function getCellBelow(ref) {
    const m = String(ref).match(/^([A-Z]+)(\d+)$/i);
    if (!m) return null;
    return m[1] + (parseInt(m[2], 10) + 1);
  }

  function getCellRight(ref) {
    const m = String(ref).match(/^([A-Z]+)(\d+)$/i);
    if (!m) return null;
    const col = letterToCol(m[1]);
    return colToLetter(col + 1) + m[2];
  }

  function focusCell(ref) {
    if (!table) return;
    const inp = table.querySelector(`input[data-cell="${ref}"]`);
    if (inp) inp.focus();
  }

  function setupExcelCellHandlers() {
    if (!table?.classList.contains("input-style-excel")) return;
    table.querySelectorAll(".create-drt-cell-input").forEach((input) => {
      input.removeEventListener("focus", onExcelCellFocus);
      input.removeEventListener("blur", onExcelCellBlur);
      input.removeEventListener("keydown", onExcelCellKeydown);
      input.removeEventListener("input", onExcelCellInput);
      input.addEventListener("focus", onExcelCellFocus);
      input.addEventListener("blur", onExcelCellBlur);
      input.addEventListener("keydown", onExcelCellKeydown);
      input.addEventListener("input", onExcelCellInput);
    });
  }

  function onExcelCellKeydown(e) {
    if (e.key === "Enter") {
      const inp = e.target;
      const val = inp.value.trim();
      if (val.startsWith("=")) {
        e.preventDefault();
        const result = evaluateFormula(val);
        if (result !== null) {
          inp.dataset.formula = val;
          inp.value = String(typeof result === "number" ? (Math.round(result * 100) / 100) : result);
        }
      }
      updateFormulaSelectionMode(false);
      clearFormulaSelection();
      e.preventDefault();
      const ref = inp.dataset.cell;
      if (ref) {
        const next = getCellBelow(ref);
        if (next) focusCell(next);
      }
    } else if (e.key === "Tab") {
      const inp = e.target;
      updateFormulaSelectionMode(false);
      clearFormulaSelection();
      const ref = inp.dataset.cell;
      if (ref) {
        e.preventDefault();
        const next = getCellRight(ref);
        if (next) focusCell(next);
      }
    }
  }

  function onExcelCellFocus(e) {
    const inp = e.target;
    syncFormulaBarFromCell(inp);
  }

  function onExcelCellInput(e) {
    const inp = e.target;
    syncFormulaBarFromCell(inp);
    updateFormulaSelectionMode(hasUnclosedBracket());
  }

  function onExcelCellBlur(e) {
    const inp = e.target;
    const val = inp.value.trim();
    if (val.startsWith("=")) {
      const result = evaluateFormula(val);
      if (result !== null) {
        inp.dataset.formula = val;
        inp.value = typeof result === "number" ? (Math.round(result * 100) / 100) : result;
      }
    }
  }

  function applyFormulaFromBar() {
    if (!formulaInput) return;
    const cellRef = formulaInput.dataset.activeCell;
    if (!cellRef || !table?.classList.contains("input-style-excel")) return;
    const inp = table.querySelector(`input[data-cell="${cellRef}"]`);
    if (inp) {
      const val = formulaInput.value.trim();
      inp.value = val;
      inp.dataset.formula = val.startsWith("=") ? val : "";
      if (val.startsWith("=")) {
        const result = evaluateFormula(val);
        if (result !== null) {
          inp.value = String(typeof result === "number" ? (Math.round(result * 100) / 100) : result);
        }
      }
    }
    return cellRef;
  }

  if (formulaInput) {
    formulaInput.addEventListener("change", applyFormulaFromBar);
    formulaInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        const cellRef = applyFormulaFromBar();
        updateFormulaSelectionMode(false);
        clearFormulaSelection();
        if (cellRef) {
          const next = getCellBelow(cellRef);
          if (next) focusCell(next);
        }
        formulaInput.blur();
      }
    });
  }

  function insertCellRefIntoFormulaBar(ref) {
    if (!formulaInput) return;
    const start = formulaInput.selectionStart;
    const end = formulaInput.selectionEnd;
    const val = formulaInput.value;
    formulaInput.value = val.slice(0, start) + ref + val.slice(end);
    formulaInput.selectionStart = formulaInput.selectionEnd = start + ref.length;
    const cellRef = formulaInput.dataset.activeCell;
    if (cellRef && table) {
      const inp = table.querySelector(`input[data-cell="${cellRef}"]`);
      if (inp) {
        inp.value = formulaInput.value;
        if (formulaInput.value.trim().startsWith("=")) inp.dataset.formula = formulaInput.value;
        inp.focus();
      } else {
        formulaInput.focus();
      }
    } else {
      formulaInput.focus();
    }
  }

  function formatRange(startRef, endRef) {
    const m1 = String(startRef).match(/^([A-Z]+)(\d+)$/i);
    const m2 = String(endRef).match(/^([A-Z]+)(\d+)$/i);
    if (!m1 || !m2) return startRef;
    const c1 = letterToCol(m1[1]), r1 = parseInt(m1[2], 10);
    const c2 = letterToCol(m2[1]), r2 = parseInt(m2[2], 10);
    const sc = Math.min(c1, c2), ec = Math.max(c1, c2);
    const sr = Math.min(r1, r2), er = Math.max(r1, r2);
    return colToLetter(sc) + sr + ":" + colToLetter(ec) + er;
  }

  function hasUnclosedBracket() {
    if (!formulaInput) return false;
    const open = (formulaInput.value.match(/\(/g) || []).length;
    const close = (formulaInput.value.match(/\)/g) || []).length;
    return open > close;
  }

  function updateFormulaSelectionMode(active) {
    if (page) page.classList.toggle("formula-selection-mode", !!active);
  }

  function getCellsInRange(ref) {
    const cells = [];
    const rangeMatch = ref.match(/^([A-Z]+)(\d+)\s*:\s*([A-Z]+)(\d+)$/i);
    if (rangeMatch) {
      const [, c1, r1, c2, r2] = rangeMatch;
      const sc = Math.min(letterToCol(c1), letterToCol(c2));
      const ec = Math.max(letterToCol(c1), letterToCol(c2));
      const sr = Math.min(parseInt(r1, 10), parseInt(r2, 10));
      const er = Math.max(parseInt(r1, 10), parseInt(r2, 10));
      for (let r = sr; r <= er; r++) {
        for (let c = sc; c <= ec; c++) {
          cells.push(colToLetter(c) + r);
        }
      }
    } else {
      const single = ref.match(/^([A-Z]+\d+)$/i);
      if (single) cells.push(single[1].toUpperCase());
    }
    return cells;
  }

  function highlightFormulaSelection(ref) {
    if (!table) return;
    table.querySelectorAll(".formula-selection-highlight").forEach((el) => el.classList.remove("formula-selection-highlight"));
    getCellsInRange(ref).forEach((cellRef) => {
      const cell = table.querySelector(`td[data-cell="${cellRef}"]`);
      if (cell) cell.classList.add("formula-selection-highlight");
    });
  }

  function clearFormulaSelection() {
    if (!table) return;
    table.querySelectorAll(".formula-selection-highlight").forEach((el) => el.classList.remove("formula-selection-highlight"));
  }

  function syncFormulaBarFromCell(inp) {
    if (formulaInput && inp) {
      formulaInput.value = inp.dataset.formula || inp.value;
      formulaInput.dataset.activeCell = inp.dataset.cell || "";
    }
  }

  if (formulaInput) {
    formulaInput.addEventListener("focus", () => updateFormulaSelectionMode(hasUnclosedBracket()));
    formulaInput.addEventListener("blur", () => updateFormulaSelectionMode(false));
    formulaInput.addEventListener("input", () => {
      updateFormulaSelectionMode(hasUnclosedBracket());
    });
  }

  let selectionStartCell = null;

  if (table) {
    table.addEventListener("mousedown", (e) => {
      if (!table.classList.contains("input-style-excel")) return;
      if (!page?.classList.contains("formula-selection-mode")) return;
      const cell = e.target.closest("td[data-cell]");
      if (cell) {
        const ref = cell.dataset.cell;
        if (ref) {
          e.preventDefault();
          e.stopPropagation();
          selectionStartCell = ref;
        }
      }
    }, true);

    table.addEventListener("mouseup", (e) => {
      if (!table.classList.contains("input-style-excel")) return;
      if (!page?.classList.contains("formula-selection-mode")) return;
      const cell = e.target.closest("td[data-cell]");
      if (cell && selectionStartCell) {
        const endRef = cell.dataset.cell;
        if (endRef) {
          e.preventDefault();
          e.stopPropagation();
          const ref = selectionStartCell === endRef ? selectionStartCell : formatRange(selectionStartCell, endRef);
          insertCellRefIntoFormulaBar(ref);
          highlightFormulaSelection(ref);
        }
      }
      selectionStartCell = null;
    }, true);
  }

  document.addEventListener("mouseup", () => { selectionStartCell = null; });

  if (toggle && table) {
    table.classList.add("input-style-modern");
    toggle.addEventListener("click", () => {
      const isExcel = table.classList.toggle("input-style-excel");
      table.classList.toggle("input-style-modern", !isExcel);
      toggle.classList.toggle("excel-mode", isExcel);
      if (page) page.classList.toggle("excel-mode-active", isExcel);
      if (toggleLabel) toggleLabel.textContent = isExcel ? "Table" : "Excel";
      setupExcelCellHandlers();
    });
  }

  const themeToggle = document.getElementById("createDrtThemeToggle");
  const themeLabel = themeToggle?.querySelector(".theme-label");
  const savedTheme = window.localStorage.getItem("drt-theme");

  function applyTheme(theme) {
    document.body.dataset.theme = theme;
    const isNight = theme === "night";
    if (themeLabel) themeLabel.textContent = isNight ? "Night Mode" : "Day Mode";
    window.localStorage.setItem("drt-theme", theme);
    if (themeToggle) {
      themeToggle.setAttribute("aria-label", isNight ? "Switch to day theme" : "Switch to night theme");
    }
  }

  if (themeToggle) {
    themeToggle.addEventListener("click", () => {
      const nextTheme = document.body.dataset.theme === "day" ? "night" : "day";
      applyTheme(nextTheme);
    });
  }

  applyTheme(savedTheme === "night" ? "night" : "day");

  const drtIdEl = document.getElementById("createDrtId");
  if (drtIdEl) {
    drtIdEl.textContent = "DRT-" + String(Math.floor(Math.random() * 90000) + 10000);
  }

  const prioritySelect = document.getElementById("createDrtPriority");
  if (prioritySelect) {
    function updatePriorityClass() {
      prioritySelect.classList.remove("priority-high", "priority-medium", "priority-low");
      const v = prioritySelect.value;
      if (v === "high") prioritySelect.classList.add("priority-high");
      else if (v === "medium") prioritySelect.classList.add("priority-medium");
      else if (v === "low") prioritySelect.classList.add("priority-low");
    }
    updatePriorityClass();
    prioritySelect.addEventListener("change", updatePriorityClass);
  }

  const resultSelect = document.getElementById("createDrtResult");
  if (resultSelect) {
    function updateResultClass() {
      resultSelect.classList.remove("result-won", "result-lost");
      const v = resultSelect.value;
      if (v === "won") resultSelect.classList.add("result-won");
      else if (v === "lost") resultSelect.classList.add("result-lost");
    }
    updateResultClass();
    resultSelect.addEventListener("change", updateResultClass);
  }
})();
