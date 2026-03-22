(function () {
  const SKU_NAMES = ["Enterprise License", "Cloud Storage Pro", "API Access Tier", "Support Package", "Premium Add-on", "Base Subscription", "Professional Services", "Training Module"];
  const YES_NO = ["Y", "N"];
  const CAPPED = ["Capped", "Uncapped"];
  const CLOUD = ["AWS", "Azure", "GCP", "On-Prem"];
  const STATUS_VALUES = ["End of Sale", "End of Renewal", "End of Life"];

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
    const grp3 = grp2.map((p, i) => (parseFloat(p) * grp1[i] * 12).toFixed(2));
    const newTotal = grp3.reduce((a, b) => a + parseFloat(b), 0).toFixed(2);
    return {
      id,
      calcValues: Array.from({ length: 5 }, () => ""),
      sku: "SKU-" + randInt(1000, 9999),
      additionalComment: "",
      applySelaDiscount: random(YES_NO),
      cappedUncapped: random(CAPPED),
      selaStdQuote: random(["SELA", "Std Quote"]),
      contractNumber: "CNT-" + randInt(100000, 999999),
      contractEndDate: randDate(),
      orgIdTenantTig: "ORG-" + randInt(100, 999),
      pg: "",
      skuName: random(SKU_NAMES),
      grp1Curr: grp1[0],
      grp1Yr: grp1.slice(1),
      grp2Curr: grp2[0],
      grp2Yr: grp2.slice(1),
      grp3Curr: grp3[0],
      grp3Yr: grp3.slice(1),
      newTotal,
      addonUnitPrice: randPrice(),
      status: random(STATUS_VALUES),
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
      calcValues: [...emptyArr],
      sku: "",
      additionalComment: "",
      applySelaDiscount: "",
      cappedUncapped: "",
      selaStdQuote: "",
      contractNumber: "",
      contractEndDate: "",
      orgIdTenantTig: "",
      pg: "",
      skuName: "",
      grp1Curr: "",
      grp1Yr: [...emptyArr],
      grp2Curr: "",
      grp2Yr: [...emptyArr],
      grp3Curr: "",
      grp3Yr: [...emptyArr],
      newTotal: "",
      addonUnitPrice: "",
      status: "",
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

  function buildYearColumns(prefix, groupKey, kind, readonly, groupTitle) {
    const valueKey = prefix + "Yr";
    const headerClass = readonly ? "col-readonly" : "";
    const cellClass = readonly ? "create-drt-cell-readonly" : "";
    const columns = [{
      key: prefix + "Current",
      field: prefix + "Curr",
      label: "Current",
      kind,
      groupKey,
      groupTitle,
      headerClass,
      cellClass,
      minWidth: 108,
    }];
    for (let i = 0; i < 5; i++) {
      columns.push({
        key: prefix + "Year" + String(i + 1),
        label: "Year " + String(i + 1),
        kind,
        groupKey,
        groupTitle,
        headerClass,
        cellClass,
        minWidth: 108,
        value: (data) => data[valueKey]?.[i] || "",
      });
    }
    return columns;
  }

  const SUBSCRIPTION_COLUMNS = buildYearColumns("grp1", "subscriptions", "input", false, "Number of Subscriptions");
  const MONTHLY_PRICE_COLUMNS = buildYearColumns("grp2", "monthly-unit-price", "input", false, "Monthly Unit Price");
  const TOTAL_ANNUAL_COLUMNS = buildYearColumns("grp3", "total-annual-price", "readonly", true, "Total AOV");

  const TABLE_COLUMNS = [
    {
      key: "rowNum",
      label: "#",
      kind: "row-number",
      headerClass: "col-row-num col-row-num-table-view",
      cellClass: "col-row-num",
      minWidth: 36,
      width: 36,
      stickyModes: ["table", "excel"],
      lockResize: true,
      meta: true,
    },
    {
      key: "delete",
      label: "",
      kind: "delete",
      headerClass: "col-delete col-delete-table-view",
      cellClass: "col-delete",
      minWidth: 40,
      width: 40,
      stickyModes: ["table", "excel"],
      lockResize: true,
      meta: true,
    },
    ...Array.from({ length: 5 }, (_, index) => ({
      key: "calc" + String(index + 1),
      label: colToLetter(index + 1),
      kind: "calc",
      headerClass: "col-calc create-drt-excel-only",
      cellClass: "col-calc create-drt-excel-only",
      minWidth: 70,
      width: 70,
      stickyModes: ["excel"],
      excelOnly: true,
    })),
    {
      key: "chartAction",
      label: "",
      kind: "chart-action",
      headerClass: "col-chart-action create-drt-excel-only",
      cellClass: "col-chart-action create-drt-excel-only",
      minWidth: 56,
      width: 56,
      stickyModes: ["excel"],
      excelOnly: true,
    },
    {
      key: "sku",
      field: "sku",
      label: "SKU",
      kind: "readonly",
      headerClass: "col-readonly col-sku-header",
      cellClass: "col-sku create-drt-cell-readonly",
      minWidth: 128,
      width: 128,
      stickyModes: ["table", "excel"],
    },
    {
      key: "additionalComment",
      field: "additionalComment",
      label: "Additional Comments",
      kind: "input",
      minWidth: 240,
      width: 240,
    },
    {
      key: "applySelaDiscount",
      field: "applySelaDiscount",
      label: "Apply SELA Discount",
      kind: "select",
      source: ["Yes", "No"],
      minWidth: 190,
      width: 190,
    },
    {
      key: "cappedUncapped",
      field: "cappedUncapped",
      label: "Capped/Uncapped",
      kind: "select",
      source: ["Capped", "Uncapped"],
      minWidth: 190,
      width: 190,
    },
    {
      key: "selaStdQuote",
      field: "selaStdQuote",
      label: "SELA/Std Quote",
      kind: "input",
      minWidth: 170,
      width: 170,
    },
    {
      key: "contractNumber",
      field: "contractNumber",
      label: "Contract Number",
      kind: "input",
      minWidth: 170,
      width: 170,
    },
    {
      key: "contractEndDate",
      field: "contractEndDate",
      label: "Contract End Date",
      kind: "date",
      minWidth: 170,
      width: 170,
    },
    {
      key: "pg",
      field: "pg",
      label: "PG",
      kind: "readonly",
      headerClass: "col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 40,
      width: 40,
    },
    {
      key: "orgIdTenantTig",
      field: "orgIdTenantTig",
      label: "OrgID/Tenant/TIG",
      kind: "input",
      minWidth: 170,
      width: 170,
    },
    {
      key: "skuName",
      field: "skuName",
      label: "SKU Name",
      kind: "input",
      headerClass: "col-sku-name-header",
      cellClass: "col-sku-name",
      minWidth: 216,
      width: 216,
    },
    ...SUBSCRIPTION_COLUMNS,
    ...MONTHLY_PRICE_COLUMNS,
    ...TOTAL_ANNUAL_COLUMNS,
    {
      key: "newTotal",
      field: "newTotal",
      label: "Total",
      kind: "readonly",
      headerClass: "col-new-total col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 132,
      width: 132,
    },
    {
      key: "status",
      field: "status",
      label: "Status",
      kind: "readonly",
      headerClass: "col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 170,
      width: 170,
    },
    {
      key: "cloudOfSku",
      field: "cloudOfSku",
      label: "Cloud of SKU",
      kind: "readonly",
      headerClass: "col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 150,
      width: 150,
    },
    {
      key: "eligibleSignatureAddon",
      field: "eligibleSignatureAddon",
      label: "Eligible Signature Success Add-On Fees",
      kind: "readonly",
      headerClass: "col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 240,
      width: 240,
    },
    {
      key: "selaEligibility",
      field: "selaEligibility",
      label: "SELA Eligibility",
      kind: "readonly",
      headerClass: "col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 150,
      width: 150,
    },
    {
      key: "awsEligibility",
      field: "awsEligibility",
      label: "AWS Eligibility",
      kind: "readonly",
      headerClass: "col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 150,
      width: 150,
    },
    {
      key: "addonUnitPrice",
      field: "addonUnitPrice",
      label: "Addon Unit Price (Monthly)",
      kind: "input",
      numeric: true,
      headerClass: "col-scroll-start",
      minWidth: 190,
      width: 190,
    },
    {
      key: "listPriceMonthly",
      field: "listPriceMonthly",
      label: "List Price (Monthly)",
      kind: "readonly",
      headerClass: "col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 150,
      width: 150,
    },
    {
      key: "exitDiscount",
      field: "exitDiscount",
      label: "Exit Discount",
      kind: "readonly",
      headerClass: "col-readonly",
      cellClass: "create-drt-cell-readonly",
      minWidth: 130,
      width: 130,
    },
    {
      key: "includePctBasedPricing",
      field: "includePctBasedPricing",
      label: "Include in % Based Pricing?",
      kind: "input",
      minWidth: 250,
      width: 250,
    },
    {
      key: "exitColorGuidance",
      field: "exitColorGuidance",
      label: "Exit Color Guidance*",
      kind: "input",
      minWidth: 170,
      width: 170,
    },
    {
      key: "exitApprovalLevel",
      field: "exitApprovalLevel",
      label: "Exit Approval Level*",
      kind: "input",
      minWidth: 170,
      width: 170,
    },
  ];

  let excelColumnIndex = 0;
  TABLE_COLUMNS.forEach((column, index) => {
    column.index = index;
    if (!column.meta) {
      excelColumnIndex += 1;
      column.letter = colToLetter(excelColumnIndex);
    }
  });

  const COLUMN_BY_KEY = Object.fromEntries(TABLE_COLUMNS.map((column) => [column.key, column]));
  const EXCEL_CANVAS_KEYS = TABLE_COLUMNS.filter((column) => column.kind === "calc").map((column) => column.key);

  function getCellValue(ref) {
    const match = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match) return 0;
    const colIndex = letterToCol(match[1]) - 1;
    const rowIndex = parseInt(match[2], 10) - 1;
    const column = SPREADSHEET_COLUMNS[colIndex];
    const row = spreadsheetVisibleRows[rowIndex];
    if (!column || !row) return 0;
    const value = getSpreadsheetCellValue(column, row);
    const numeric = parseFloat(String(value ?? "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(numeric) ? numeric : 0;
  }

  function getInitials(name) {
    const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return "?";
    if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  function getOwnerChipTheme(initials) {
    const themes = [
      { bg: "#dbeafe", text: "#1d4ed8", border: "#bfdbfe" },
      { bg: "#dcfce7", text: "#15803d", border: "#bbf7d0" },
      { bg: "#fef3c7", text: "#b45309", border: "#fde68a" },
      { bg: "#fce7f3", text: "#be185d", border: "#fbcfe8" },
      { bg: "#ede9fe", text: "#6d28d9", border: "#ddd6fe" },
      { bg: "#cffafe", text: "#0f766e", border: "#a5f3fc" },
      { bg: "#fee2e2", text: "#b91c1c", border: "#fecaca" },
      { bg: "#e0f2fe", text: "#0369a1", border: "#bae6fd" },
    ];
    const hash = Array.from(String(initials || "")).reduce((sum, char) => sum + char.charCodeAt(0), 0);
    return themes[hash % themes.length];
  }

  function renderOwnerChips() {
    document.querySelectorAll("[data-owner-chip]").forEach((node) => {
      const name = String(node.dataset.ownerName || node.textContent || "").trim();
      const initials = getInitials(name);
      const theme = getOwnerChipTheme(initials);

      node.dataset.ownerName = name;
      node.innerHTML = "";

      const avatar = document.createElement("span");
      avatar.className = "create-drt-owner-avatar";
      avatar.textContent = initials;
      avatar.style.backgroundColor = theme.bg;
      avatar.style.color = theme.text;
      avatar.style.border = "1px solid " + theme.border;

      const label = document.createElement("span");
      label.className = "create-drt-owner-name";
      label.textContent = name;

      node.appendChild(avatar);
      node.appendChild(label);
    });
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

  function getColumnValue(column, data) {
    if (typeof column.value === "function") return column.value(data);
    if (column.field) return data[column.field];
    return "";
  }

  function renderColumnCell(column, data, rowIndex) {
    const className = [column.cellClass, "create-drt-cell", "create-drt-col-" + column.key]
      .filter(Boolean)
      .join(" ");
    if (column.kind === "row-number") {
      return `<td class="${className}" data-col-key="${column.key}">${rowIndex}</td>`;
    }
    if (column.kind === "delete") {
      return `<td class="${className}" data-col-key="${column.key}"></td>`;
    }
    const cellRef = (column.letter || "") + String(rowIndex);
    if (column.kind === "readonly") {
      const value = getColumnValue(column, data) || "";
      return `<td class="${className} create-drt-cell-readonly" data-col-key="${column.key}" data-cell="${cellRef}"><span class="create-drt-cell-value">${value}</span></td>`;
    }
    if (column.kind === "calc") {
      return `<td class="${className}" data-col-key="${column.key}" data-cell="${cellRef}"><input type="text" class="create-drt-cell-input create-drt-calc-input" value="" data-formula="" data-cell="${cellRef}"></td>`;
    }
    const value = getColumnValue(column, data) || "";
    return `<td class="${className}" data-col-key="${column.key}" data-cell="${cellRef}"><input type="text" class="create-drt-cell-input" value="${value}" data-cell="${cellRef}"></td>`;
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

    tr.innerHTML = TABLE_COLUMNS.map((column) => renderColumnCell(column, data, rowIndex)).join("");
    tr.querySelector('[data-col-key="delete"]').appendChild(delBtn);
    return tr;
  }

  const tbody = document.getElementById("createDrtBody");
  const table = document.getElementById("createDrtTable");
  const thead = document.getElementById("createDrtHead") || table?.querySelector("thead");
  const colgroup = document.getElementById("createDrtColgroup");
  const toggle = document.getElementById("inputStyleToggle");
  const toggleLabel = toggle?.querySelector(".toggle-label");

  const page = document.querySelector(".create-drt-page");
  const formulaBar = document.getElementById("excelFormulaBar");
  const formulaCellRef = document.getElementById("excelFormulaCellRef");
  const formulaInput = document.getElementById("excelFormulaInput");
  const excelToggle = document.getElementById("createDrtExcelToggle");
  const excelOutline = document.getElementById("createDrtExcelOutline");
  const addlInfoOutline = document.getElementById("createDrtAddlInfoOutline");
  const addlInfoToggleEl = document.getElementById("addlInfoToggleBtn");
  const pricingInfoOutline = document.getElementById("createDrtPricingInfoOutline");
  const pricingInfoToggleEl = document.getElementById("pricingInfoToggleBtn");
  const productInfoOutline = document.getElementById("createDrtProductInfoOutline");
  const productInfoGroupToggleEl = document.getElementById("productInfoGroupToggleBtn");
  const excelCanvasPanel = document.getElementById("createDrtExcelCanvasPanel");
  const spreadsheetHost = document.getElementById("createDrtSpreadsheet");
  const chartModal = document.getElementById("createDrtChartModal");
  const chartSubtitle = document.getElementById("createDrtChartSubtitle");
  const chartSvg = document.getElementById("createDrtChartSvg");
  const undoButton = document.getElementById("createDrtUndoBtn");
  const redoButton = document.getElementById("createDrtRedoBtn");
  const workpadCountNumber = document.getElementById("createDrtWorkpadCountNumber");
  const workpadCountLabel = document.getElementById("createDrtWorkpadCountLabel");
  const SPREADSHEET_COLUMNS = TABLE_COLUMNS.filter((column) => !column.meta);
  const SPREADSHEET_ROW_HEADER_WIDTH = 84;
  const EXCEL_OUTLINE_TOGGLE_SIZE = 28;
  const EXCEL_CANVAS_COLUMN_INDEXES = SPREADSHEET_COLUMNS
    .map((column, index) => (column.kind === "calc" ? index : -1))
    .filter((index) => index >= 0);
  const EXCEL_OUTLINE_COLUMN_KEYS = [...EXCEL_CANVAS_KEYS, "chartAction"];
  const EXCEL_CANVAS_TOTAL_WIDTH = SPREADSHEET_COLUMNS
    .filter((column) => EXCEL_OUTLINE_COLUMN_KEYS.includes(column.key))
    .reduce((sum, column) => sum + (column.width || column.minWidth || 120), 0);
  const EXCEL_CHART_COLUMN_WIDTH = COLUMN_BY_KEY.chartAction?.width || COLUMN_BY_KEY.chartAction?.minWidth || 56;
  const EXCEL_CANVAS_OUTLINE_WIDTH = EXCEL_CANVAS_TOTAL_WIDTH - EXCEL_CHART_COLUMN_WIDTH + (EXCEL_CHART_COLUMN_WIDTH / 2) + (EXCEL_OUTLINE_TOGGLE_SIZE / 2);
  const DETAIL_KEYS = ["sku", "additionalComment", "applySelaDiscount", "cappedUncapped", "selaStdQuote", "contractNumber", "contractEndDate", "pg", "orgIdTenantTig", "skuName"];
  const ADDITIONAL_INFO_KEYS = ["additionalComment", "applySelaDiscount", "cappedUncapped", "selaStdQuote", "contractNumber", "contractEndDate"];
  // Pricing Info covers all Number of Subscriptions + Monthly Unit Price columns (12 cols).
  // Anchor: grp3Current (Current under Total AOV) — toggle button centres over it.
  const PRICING_INFO_KEYS = [
    ...SUBSCRIPTION_COLUMNS.map((c) => c.key),
    ...MONTHLY_PRICE_COLUMNS.map((c) => c.key),
  ];
  const PRICING_INFO_ANCHOR_KEY = "grp3Current";
  const PRICING_INFO_ANCHOR_WIDTH = COLUMN_BY_KEY[PRICING_INFO_ANCHOR_KEY]?.minWidth || 108;
  function computePricingInfoOutlineMetrics() {
    const baseCols = getSpreadsheetBaseColumnsForCurrentView();
    // Also exclude addl-info columns when they are collapsed (they shift the position)
    const activeCols = additionalInfoCollapsed
      ? baseCols.filter((col) => !ADDITIONAL_INFO_KEYS.includes(col.key))
      : baseCols;
    let left = SPREADSHEET_ROW_HEADER_WIDTH;
    let width = 0;
    let found = false;
    for (const col of activeCols) {
      const w = col.width || col.minWidth || 120;
      if (PRICING_INFO_KEYS.includes(col.key)) { found = true; width += w; }
      else if (!found) { left += w; }
    }
    // Toggle is now on the LEFT — expanded width covers the group; collapsed shows toggle + label only.
    const expandedWidth = width;
    const collapsedWidth = EXCEL_OUTLINE_TOGGLE_SIZE + 90;
    return { left, expandedWidth, collapsedWidth };
  }
  let currentPricingInfoMetrics = { left: SPREADSHEET_ROW_HEADER_WIDTH, expandedWidth: 0, collapsedWidth: 0 };
  // OrgID/Tenant/TIG is the anchor column for the Additional Info outline toggle
  // (always visible, toggle button is centered over it — mirrors how chartAction anchors the Excel Canvas outline)
  const ADDL_INFO_ANCHOR_KEY = "orgIdTenantTig";
  const ADDL_INFO_ANCHOR_WIDTH = COLUMN_BY_KEY[ADDL_INFO_ANCHOR_KEY]?.width || COLUMN_BY_KEY[ADDL_INFO_ANCHOR_KEY]?.minWidth || 170;
  // Compute Additional Info outline position & width dynamically, respecting current excelCanvasCollapsed state
  function computeAddlInfoOutlineMetrics() {
    const baseCols = getSpreadsheetBaseColumnsForCurrentView();
    let left = SPREADSHEET_ROW_HEADER_WIDTH;
    let width = 0;
    let found = false;
    for (const col of baseCols) {
      const w = col.width || col.minWidth || 120;
      if (ADDITIONAL_INFO_KEYS.includes(col.key)) { found = true; width += w; }
      else if (!found) { left += w; }
    }
    // Toggle is now on the LEFT — expanded width covers the group; collapsed shows toggle + label only.
    const expandedWidth = width;
    const collapsedWidth = EXCEL_OUTLINE_TOGGLE_SIZE + 90;
    return { left, expandedWidth, collapsedWidth };
  }
  // Initialized to a safe default; recomputed in applyExcelCollapsed / applyAddlInfoCollapsed
  // (cannot call computeAddlInfoOutlineMetrics() here — excelCanvasCollapsed is declared further below)
  let currentAddlInfoMetrics = { left: SPREADSHEET_ROW_HEADER_WIDTH, expandedWidth: 0, collapsedWidth: 0 };

  function computeProductInfoGroupOutlineMetrics() {
    const baseCols = getSpreadsheetBaseColumnsForCurrentView();
    const afterAddl = additionalInfoCollapsed
      ? baseCols.filter((col) => !ADDITIONAL_INFO_KEYS.includes(col.key))
      : baseCols;
    const afterPricing = pricingInfoCollapsed
      ? afterAddl.filter((col) => !PRICING_INFO_KEYS.includes(col.key))
      : afterAddl;
    let left = SPREADSHEET_ROW_HEADER_WIDTH;
    let width = 0;
    let found = false;
    for (const col of afterPricing) {
      const w = col.width || col.minWidth || 120;
      if (PRODUCT_GROUP_KEYS.includes(col.key)) { found = true; width += w; }
      else if (!found) { left += w; }
    }
    const expandedWidth = width;
    const collapsedWidth = EXCEL_OUTLINE_TOGGLE_SIZE + 90;
    return { left, expandedWidth, collapsedWidth };
  }
  let currentProductInfoGroupMetrics = { left: SPREADSHEET_ROW_HEADER_WIDTH, expandedWidth: 0, collapsedWidth: 0 };
  const MERGED_DETAIL_HEADER_KEYS = DETAIL_KEYS.filter((key) => key !== "sku");
  const POST_KEYS = ["newTotal", "status", "cloudOfSku", "eligibleSignatureAddon", "selaEligibility", "awsEligibility", "addonUnitPrice", "listPriceMonthly", "exitDiscount", "includePctBasedPricing", "exitColorGuidance", "exitApprovalLevel"];
  const PRODUCT_GROUP_KEYS = ["status", "cloudOfSku", "eligibleSignatureAddon", "selaEligibility", "awsEligibility"];
  // Product Info columns get HOT group header — exclude from dark-blue merged header set
  const MERGED_POST_HEADER_KEYS = POST_KEYS.filter((key) => key !== "newTotal" && !PRODUCT_GROUP_KEYS.includes(key));

  function getSpreadsheetBaseColumnsForCurrentView() {
    if (!excelCanvasCollapsed) return SPREADSHEET_COLUMNS;
    return SPREADSHEET_COLUMNS.filter((column) => column.kind !== "calc");
  }

  let spreadsheetInstance = null;
  let drtRows = [];
  let spreadsheetVisibleRows = [];
  let hidePgGreen = false;
  let spreadsheetIsRendering = false;
  let activeSpreadsheetCell = null;
  let activeSpreadsheetSelection = null;
  let excelCanvasCollapsed = false;
  let additionalInfoCollapsed = false;
  let pricingInfoCollapsed = false;
  let productInfoGroupCollapsed = false;
  let spreadsheetScrollHolder = null;
  let currentWorkpadProductCount = 0;
  const spreadsheetFormulaStore = new Map();
  const pendingSpreadsheetFormulaChanges = new Map();
  // Explicit next-cell target set by handleSpreadsheetChange on Enter-key edits.
  // restoreSpreadsheetSelection() consumes this first, bypassing activeSpreadsheetCell.
  let _pendingEditSelection = null; // { x: colIndex, y: rowIndex }

  // ── Chart modal state ─────────────────────────────────────────────────────
  let _chartActiveRow = null;       // row data object currently displayed
  let _chartType = "line";          // "line" | "bar" | "radar"  (single-panel)
  let _chartLayout = 1;             // 1 | 2 | 3 panels
  let _chartPanelTypes = ["line", "bar", "radar"]; // type per slot

  // Vibrant chart color palette
  const _CP = {
    // Per-series colors (bars / radar axes)
    series: [
      { hi: "#7c3aed", lo: "#5b21b6" }, // violet
      { hi: "#db2777", lo: "#9d174d" }, // rose
      { hi: "#d97706", lo: "#92400e" }, // amber
      { hi: "#059669", lo: "#064e3b" }, // emerald
      { hi: "#0284c7", lo: "#0c4a6e" }, // sky
      { hi: "#9333ea", lo: "#6b21a8" }, // purple
    ],
    // Line chart
    lineA: "#7c3aed", lineB: "#2563eb",
    areaA: "rgba(124,58,237,0.22)", areaB: "rgba(37,99,235,0.00)",
    // Radar
    radarFill: "rgba(124,58,237,0.22)", radarStroke: "#7c3aed", radarGrid: "#ede9fe",
    // Shared
    grid: "#f1f5f9", gridStrong: "#c7d2fe",
    axisLbl: "#94a3b8", valueLbl: "#374151", peakLbl: "#5b21b6",
  };

  // ── Scenario state ────────────────────────────────────────────────────────
  let _scenarios = [];
  let _activeScenarioId = null;
  let _scenarioIdCounter = 0;
  let _setPriorityFn = null;        // ref to setPriority() exposed from priority-picker block
  const SCENARIO_COLORS = ["#6366f1","#8b5cf6","#0ea5e9","#10b981","#f59e0b","#ef4444","#ec4899","#14b8a6"];

  // ── Formula-range-insertion state ─────────────────────────────────────────
  // Tracks whether the user is clicking/dragging to select a range while
  // a formula is being typed (in the formula bar OR the HOT cell editor).
  let _fRangeActive = false;        // range-injection mode is on
  let _fRangeInsertPos = 0;         // char offset in formula where ref starts
  let _fRangeInsertLen = 0;         // length of the last injected ref (for replacement)
  let _fRangeSource = "bar";        // "bar" | "editor"
  let _fRangeSavedFormula = "";     // formula snapshot taken at mousedown (editor case)
  let _fRangeCooldown = false;      // true briefly after range mode ends — blocks overwrite
  let _fRangeTargetCell = null;     // {x, y} cell the formula belongs to (editor case)

  // ── Scenario helpers ──────────────────────────────────────────────────────
  function _makeScenarioId() { return "sc-" + (++_scenarioIdCounter); }

  function _escHtml(s) {
    return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function _captureFormState() {
    const get = (id) => document.getElementById(id)?.value || "";
    return {
      accountName: get("createDrtAccountName"),
      country:     get("createDrtCountry"),
      currency:    get("createDrtCurrency"),
      term:        get("createDrtTerm") || "1",
      priority:    get("createDrtPriority") || "low",
      targetDate:  get("createDrtTargetDate"),
      assignedTo:  get("createDrtAssignedTo"),
      assignTo:    get("createDrtAssignTo"),
    };
  }

  function _restoreFormState(formState) {
    if (!formState) return;
    const setVal = (id, val) => { const el = document.getElementById(id); if (el) el.value = val ?? ""; };
    setVal("createDrtAccountName", formState.accountName);
    setVal("createDrtCountry",     formState.country);
    setVal("createDrtCurrency",    formState.currency);
    setVal("createDrtTargetDate",  formState.targetDate);
    setVal("createDrtAssignedTo",  formState.assignedTo);
    setVal("createDrtAssignTo",    formState.assignTo);
    // Term — fire change event so collapsed-metrics label updates
    const termEl = document.getElementById("createDrtTerm");
    if (termEl) { termEl.value = formState.term || "1"; termEl.dispatchEvent(new Event("change")); }
    // Priority — use exposed fn if available, else set hidden value and update pill manually
    const pri = formState.priority || "low";
    if (_setPriorityFn) {
      _setPriorityFn(pri);
    } else {
      const hidden = document.getElementById("createDrtPriority");
      if (hidden) hidden.value = pri;
      const display = document.getElementById("createDrtPriorityDisplay");
      if (display) {
        const LABELS = { high: "High", medium: "Medium", low: "Low" };
        const CLS    = { high: "create-drt-priority-pill--high", medium: "create-drt-priority-pill--medium", low: "create-drt-priority-pill--low" };
        display.className = `create-drt-priority-pill ${CLS[pri] || CLS.low}`;
        const txt = display.querySelector(".create-drt-priority-text");
        if (txt) txt.textContent = LABELS[pri] || "Low";
      }
    }
  }

  function _deepCloneRows(rows) {
    return (rows || []).map((row) => ({
      ...row,
      calcValues: [...(row.calcValues || [])],
      grp1Yr:     [...(row.grp1Yr     || [])],
      grp2Yr:     [...(row.grp2Yr     || [])],
      grp3Yr:     [...(row.grp3Yr     || [])],
    }));
  }

  function _saveCurrentScenario() {
    const sc = _scenarios.find((s) => s.id === _activeScenarioId);
    if (!sc) return;
    sc.rows      = _deepCloneRows(drtRows);
    sc.formState = _captureFormState();
  }

  function _renderScenarioBar() {
    const tabsHost = document.getElementById("createDrtScenarioTabs");
    const countEl  = document.getElementById("createDrtScenarioCount");
    if (!tabsHost) return;

    tabsHost.innerHTML = _scenarios.map((sc) => {
      const isActive = sc.id === _activeScenarioId;
      const dotStyle = `background:${sc.color};box-shadow:0 0 6px ${sc.color}80`;
      return `
        <div class="create-drt-scenario-tab${isActive ? " create-drt-scenario-tab--active" : ""}"
             data-scenario-id="${sc.id}"
             tabindex="${isActive ? 0 : -1}"
             role="tab"
             aria-selected="${isActive}">
          <span class="create-drt-scenario-tab__dot" style="${dotStyle}"></span>
          <span class="create-drt-scenario-tab__name" data-scenario-name="${sc.id}">${_escHtml(sc.name)}</span>
          <button class="create-drt-scenario-tab__rename" data-scenario-rename="${sc.id}" type="button" title="Rename" aria-label="Rename ${_escHtml(sc.name)}">
            <svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
          </button>
          ${_scenarios.length > 1 ? `
          <button class="create-drt-scenario-tab__del" data-scenario-delete="${sc.id}" type="button" title="Delete scenario" aria-label="Delete ${_escHtml(sc.name)}">
            <svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="m18 6-12 12M6 6l12 12"/></svg>
          </button>` : ""}
        </div>`;
    }).join("");

    if (countEl) {
      countEl.textContent = _scenarios.length === 1 ? "1 scenario" : `${_scenarios.length} scenarios`;
    }
  }

  function _activateScenario(id) {
    if (id === _activeScenarioId) return;
    _saveCurrentScenario();
    _activeScenarioId = id;
    const sc = _scenarios.find((s) => s.id === id);
    if (!sc) return;
    drtRows = _deepCloneRows(sc.rows);
    _restoreFormState(sc.formState);
    renderSpreadsheet();
    _renderScenarioBar();
  }

  function _addScenario() {
    _saveCurrentScenario();
    const newId    = _makeScenarioId();
    const colorIdx = _scenarios.length % SCENARIO_COLORS.length;
    const num      = _scenarios.length + 1;
    // Name: "Scenario B", "Scenario C", ... (skip "A" = Base Case)
    const letter   = num <= 26 ? String.fromCharCode(64 + num) : String(num);
    _scenarios.push({
      id:        newId,
      name:      "Scenario " + letter,
      color:     SCENARIO_COLORS[colorIdx],
      rows:      _deepCloneRows(drtRows),
      formState: _captureFormState(),
    });
    // Activate the new scenario without re-saving current (already saved above)
    _activeScenarioId = newId;
    const sc = _scenarios[_scenarios.length - 1];
    drtRows = _deepCloneRows(sc.rows);
    _renderScenarioBar();
    // renderSpreadsheet is NOT needed here since we cloned — data is identical
    // But re-render to ensure HOT shows the new scenario's rows correctly
    renderSpreadsheet();
  }

  function _deleteScenario(id) {
    if (_scenarios.length <= 1) return; // always keep at least one scenario
    const idx = _scenarios.findIndex((s) => s.id === id);
    if (idx < 0) return;
    _scenarios.splice(idx, 1);
    if (id === _activeScenarioId) {
      // Activate the nearest scenario
      const nextIdx = Math.min(idx, _scenarios.length - 1);
      const sc      = _scenarios[nextIdx];
      _activeScenarioId = sc.id;
      drtRows = _deepCloneRows(sc.rows);
      _restoreFormState(sc.formState);
      renderSpreadsheet();
    }
    _renderScenarioBar();
  }

  function _renameScenario(id, name) {
    const sc = _scenarios.find((s) => s.id === id);
    if (!sc) return;
    sc.name = String(name || "").trim() || sc.name;
    _renderScenarioBar();
  }

  function _startScenarioRename(id) {
    const nameEl = document.querySelector(`[data-scenario-name="${id}"]`);
    if (!nameEl) return;
    const sc = _scenarios.find((s) => s.id === id);
    if (!sc) return;
    const originalName = sc.name;
    // Replace span content with an input
    const inp = document.createElement("input");
    inp.type  = "text";
    inp.value = originalName;
    inp.className = "create-drt-scenario-tab__rename-input";
    nameEl.textContent = "";
    nameEl.appendChild(inp);
    inp.select();
    inp.focus();
    let committed = false;
    const commit = () => {
      if (committed) return;
      committed = true;
      const newName = inp.value.trim() || originalName;
      _renameScenario(id, newName);
    };
    inp.addEventListener("blur",    commit);
    inp.addEventListener("keydown", (e) => {
      if (e.key === "Enter")  { inp.blur(); }
      if (e.key === "Escape") { committed = true; nameEl.textContent = originalName; _renderScenarioBar(); }
      e.stopPropagation();
    });
  }

  function buildCellRangeRef(col1, row1, col2, row2) {
    const c1 = Math.min(col1, col2);
    const c2 = Math.max(col1, col2);
    const r1 = Math.min(row1, row2);
    const r2 = Math.max(row1, row2);
    const a = getSpreadsheetCellRef(c1, r1);
    if (c1 === c2 && r1 === r2) return a;
    return a + ":" + getSpreadsheetCellRef(c2, r2);
  }

  function _injectRangeRef(ref) {
    if (_fRangeSource === "editor") {
      // Try to inject into HOT cell editor; if the editor was already closed
      // by HOT's own mousedown handling, fall through to the formula bar.
      const hot = getSpreadsheetWorksheet();
      let injectedIntoEditor = false;
      if (hot) {
        try {
          const editor = hot.getActiveEditor();
          if (editor && editor.isOpened()) {
            const cur = String(editor.getValue() || "");
            const before = cur.slice(0, _fRangeInsertPos);
            const after  = cur.slice(_fRangeInsertPos + _fRangeInsertLen);
            const next   = before + ref + after;
            _fRangeInsertLen = ref.length;
            editor.setValue(next);
            if (formulaInput) formulaInput.value = next;
            injectedIntoEditor = true;
          }
        } catch {}
      }
      if (!injectedIntoEditor && formulaInput) {
        // Editor was closed — inject into formula bar using the saved formula
        const cur    = _fRangeSavedFormula || formulaInput.value;
        const before = cur.slice(0, _fRangeInsertPos);
        const after  = cur.slice(_fRangeInsertPos + _fRangeInsertLen);
        const next   = before + ref + after;
        _fRangeInsertLen = ref.length;
        _fRangeSavedFormula = next;
        formulaInput.value = next;
      }
    } else {
      // Inject into formula bar
      if (!formulaInput) return;
      const cur    = formulaInput.value;
      const before = cur.slice(0, _fRangeInsertPos);
      const after  = cur.slice(_fRangeInsertPos + _fRangeInsertLen);
      const next   = before + ref + after;
      _fRangeInsertLen = ref.length;
      formulaInput.value = next;
      const pos = _fRangeInsertPos + ref.length;
      formulaInput.setSelectionRange(pos, pos);
    }
  }

  function parseNumericCell(value) {
    const parsed = parseFloat(String(value ?? "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }

  function formatDerivedAmount(quantity, price) {
    if (quantity === null || price === null) return "";
    return (quantity * price * 12).toFixed(2);
  }

  function ensureCalcValues(row) {
    if (!row) return [];
    if (!Array.isArray(row.calcValues)) {
      row.calcValues = Array.from({ length: EXCEL_CANVAS_KEYS.length }, () => "");
    }
    while (row.calcValues.length < EXCEL_CANVAS_KEYS.length) {
      row.calcValues.push("");
    }
    return row.calcValues;
  }

  function getCalcValueIndex(column) {
    return EXCEL_CANVAS_KEYS.indexOf(column?.key);
  }

  function getSpreadsheetFormulaKey(row, column) {
    return row && column ? String(row.id) + "::" + column.key : "";
  }

  function getStoredFormula(row, column) {
    return spreadsheetFormulaStore.get(getSpreadsheetFormulaKey(row, column)) || "";
  }

  function setStoredFormula(row, column, formula) {
    const key = getSpreadsheetFormulaKey(row, column);
    if (!key) return;
    if (formula) spreadsheetFormulaStore.set(key, formula);
    else spreadsheetFormulaStore.delete(key);
  }

  function setPendingFormula(row, column, formula) {
    const key = getSpreadsheetFormulaKey(row, column);
    if (!key) return;
    if (formula) pendingSpreadsheetFormulaChanges.set(key, formula);
    else pendingSpreadsheetFormulaChanges.delete(key);
  }

  function consumePendingFormula(row, column) {
    const key = getSpreadsheetFormulaKey(row, column);
    if (!key) return "";
    const formula = pendingSpreadsheetFormulaChanges.get(key) || "";
    pendingSpreadsheetFormulaChanges.delete(key);
    return formula;
  }

  function formatFormulaResult(result) {
    return typeof result === "number" ? String(Math.round(result * 100) / 100) : String(result ?? "");
  }

  function getSpreadsheetCellRef(x, y) {
    return colToLetter(x + 1) + String(y + 1);
  }

  function getVisibleSpreadsheetColumns() {
    const base = getSpreadsheetBaseColumnsForCurrentView();
    const afterAddl = additionalInfoCollapsed
      ? base.filter((column) => !ADDITIONAL_INFO_KEYS.includes(column.key))
      : base;
    const afterPricing = pricingInfoCollapsed
      ? afterAddl.filter((column) => !PRICING_INFO_KEYS.includes(column.key))
      : afterAddl;
    return productInfoGroupCollapsed
      ? afterPricing.filter((column) => !PRODUCT_GROUP_KEYS.includes(column.key))
      : afterPricing;
  }

  function canEditSpreadsheetColumn(column) {
    return !!column && !["readonly", "chart-action"].includes(column.kind);
  }

  function getLiveSpreadsheetValue(x, y) {
    const hot = getSpreadsheetWorksheet();
    if (!hot) return null;
    try {
      const val = hot.getDataAtCell(y, x);
      if (val !== undefined && val !== null) return String(val);
    } catch {}
    return null;
  }

  function syncWorksheetCellValue(x, y, value) {
    const hot = getSpreadsheetWorksheet();
    if (!hot) return;
    const normalizedValue = String(value ?? "");
    const previousRenderingState = spreadsheetIsRendering;
    spreadsheetIsRendering = true;
    try {
      hot.setDataAtCell(y, x, normalizedValue);
    } catch {}
    spreadsheetIsRendering = previousRenderingState;
  }

  function normalizeSpreadsheetEntry(column, value) {
    const rawValue = String(value ?? "");
    const trimmedValue = rawValue.trim();
    if (!canEditSpreadsheetColumn(column)) {
      return { displayValue: rawValue, formula: "" };
    }
    if (trimmedValue.startsWith("=")) {
      const result = evaluateFormula(trimmedValue);
      return {
        displayValue: result === null ? rawValue : formatFormulaResult(result),
        formula: trimmedValue,
      };
    }
    return { displayValue: rawValue, formula: "" };
  }

  function updateFormulaBarState(x, y) {
    // Handsontable uses (row, col) — our internal convention stays (x=col, y=row)
    const safeX = Number.parseInt(x, 10);
    const safeY = Number.parseInt(y, 10);
    const column = getVisibleSpreadsheetColumns()[safeX];
    const row = spreadsheetVisibleRows[safeY];
    activeSpreadsheetCell = column && row ? { key: column.key, y: safeY } : null;

    if (!formulaInput || !formulaCellRef) return;
    if (!column || !row) {
      formulaCellRef.textContent = "";
      formulaInput.value = "";
      formulaInput.readOnly = true;
      formulaInput.dataset.activeCell = "";
      formulaBar?.setAttribute("aria-disabled", "true");
      return;
    }

    const cellRef = getSpreadsheetCellRef(safeX, safeY);
    formulaCellRef.textContent = cellRef;
    formulaInput.dataset.activeCell = cellRef;
    formulaInput.value = getStoredFormula(row, column) || getLiveSpreadsheetValue(safeX, safeY) || getSpreadsheetCellValue(column, row);
    formulaInput.readOnly = !canEditSpreadsheetColumn(column);
    formulaBar?.setAttribute("aria-disabled", formulaInput.readOnly ? "true" : "false");
  }

  function restoreSpreadsheetSelection() {
    const hot = getSpreadsheetWorksheet();
    if (!hot || !spreadsheetVisibleRows.length) {
      updateFormulaBarState(-1, -1);
      return;
    }

    const visibleColumns = getVisibleSpreadsheetColumns();
    const maxX = visibleColumns.length - 1;
    const maxY = spreadsheetVisibleRows.length - 1;

    let selection;
    if (_pendingEditSelection) {
      // Consume the explicit target set by Enter-key edit (y = editedRow + 1).
      // This is more reliable than relying on HOT's enterMoves firing before re-render.
      const p = _pendingEditSelection;
      _pendingEditSelection = null;
      selection = {
        x: Math.min(Math.max(p.x, 0), maxX),
        y: Math.min(Math.max(p.y, 0), maxY),
      };
    } else {
      const activeX = activeSpreadsheetCell
        ? visibleColumns.findIndex((column) => column.key === activeSpreadsheetCell.key)
        : -1;
      selection = activeX >= 0
        ? {
            x: Math.min(Math.max(activeX, 0), maxX),
            y: Math.min(Math.max(activeSpreadsheetCell.y, 0), maxY),
          }
        : { x: 0, y: 0 };
    }

    try {
      hot.selectCell(selection.y, selection.x);
    } catch {}
    updateFormulaBarState(selection.x, selection.y);
  }

  function recalculateSpreadsheetFormulas() {
    spreadsheetVisibleRows.forEach((row) => {
      SPREADSHEET_COLUMNS.forEach((column) => {
        const formula = getStoredFormula(row, column);
        if (!formula || !canEditSpreadsheetColumn(column)) return;
        const result = evaluateFormula(formula);
        setRowValueFromSpreadsheet(column, row, result === null ? formula : formatFormulaResult(result));
      });
    });
  }

  function recalculateDerivedRow(row) {
    ensureCalcValues(row);
    const subscriptions = [row.grp1Curr, ...row.grp1Yr];
    const monthlyPrices = [row.grp2Curr, ...row.grp2Yr];
    const totals = subscriptions.map((subscription, index) => {
      return formatDerivedAmount(parseNumericCell(subscription), parseNumericCell(monthlyPrices[index]));
    });
    row.grp3Curr = totals[0];
    row.grp3Yr = totals.slice(1);

    const numericTotals = totals
      .map((value) => parseNumericCell(value))
      .filter((value) => value !== null);
    row.newTotal = numericTotals.length
      ? numericTotals.reduce((sum, value) => sum + value, 0).toFixed(2)
      : "";
  }

  function getSpreadsheetCellValue(column, row) {
    if (column.kind === "chart-action") {
      return "";
    }
    if (column.kind === "calc") {
      const calcValues = ensureCalcValues(row);
      const index = getCalcValueIndex(column);
      return index >= 0 ? calcValues[index] || "" : "";
    }
    return getColumnValue(column, row) || "";
  }

  function getChartSeries(row) {
    const values = [row.grp3Curr, ...(row.grp3Yr || [])]
      .map((value) => parseNumericCell(value))
      .map((value) => (value === null ? 0 : value));
    return values.some((value) => value !== 0) ? values : [0, 0, 0, 0, 0, 0];
  }

  // ── Chart helpers ─────────────────────────────────────────────────────────
  function _chartFmtVal(v) {
    if (v >= 1000000) return (v / 1000000).toFixed(1).replace(/\.0$/, "") + "M";
    if (v >= 1000)    return (v / 1000).toFixed(1).replace(/\.0$/, "") + "K";
    return Math.round(v).toString();
  }

  function _chartSmoothPath(pts) {
    if (!pts.length) return "";
    let d = `M ${pts[0].x.toFixed(2)},${pts[0].y.toFixed(2)}`;
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[Math.max(i - 1, 0)];
      const p1 = pts[i];
      const p2 = pts[i + 1];
      const p3 = pts[Math.min(i + 2, pts.length - 1)];
      const cp1x = (p1.x + (p2.x - p0.x) / 6).toFixed(2);
      const cp1y = (p1.y + (p2.y - p0.y) / 6).toFixed(2);
      const cp2x = (p2.x - (p3.x - p1.x) / 6).toFixed(2);
      const cp2y = (p2.y - (p3.y - p1.y) / 6).toFixed(2);
      d += ` C ${cp1x},${cp1y} ${cp2x},${cp2y} ${p2.x.toFixed(2)},${p2.y.toFixed(2)}`;
    }
    return d;
  }

  function _buildLineChart(values, labels, W, H, compact) {
    const s = compact ? 0.72 : 1;
    const pad = { top: Math.round(44*s), right: Math.round(36*s), bottom: Math.round(52*s), left: Math.round(60*s) };
    const iW = W - pad.left - pad.right;
    const iH = H - pad.top - pad.bottom;
    const mx = Math.max(...values, 1);
    const uid = "lc" + Date.now() + Math.random().toString(36).slice(2,6);
    const fs = compact ? 9.5 : 11;
    const fsLbl = compact ? 10 : 12;
    const dotR = compact ? 4.5 : 5.5;
    const lineW = compact ? 2.5 : 3.2;

    const pts = values.map((v, i) => ({
      x: pad.left + (iW * i) / Math.max(values.length - 1, 1),
      y: pad.top + iH - (v / mx) * iH,
      v, lbl: labels[i],
    }));

    const grid = Array.from({ length: 5 }, (_, i) => {
      const y = pad.top + (iH * i) / 4;
      const v = ((4 - i) / 4) * mx;
      return `<line x1="${pad.left}" y1="${y.toFixed(1)}" x2="${(pad.left+iW).toFixed(1)}" y2="${y.toFixed(1)}" stroke="${_CP.grid}" stroke-width="1"/>
        <text x="${(pad.left-7).toFixed(1)}" y="${(y+4).toFixed(1)}" text-anchor="end" font-size="${fs}" fill="${_CP.axisLbl}" font-family="system-ui,sans-serif">${_chartFmtVal(v)}</text>`;
    }).join("");

    const linePath = _chartSmoothPath(pts);
    const areaPath = `${linePath} L ${pts[pts.length-1].x.toFixed(1)},${(pad.top+iH).toFixed(1)} L ${pts[0].x.toFixed(1)},${(pad.top+iH).toFixed(1)} Z`;
    const maxV = Math.max(...values);

    const dots = pts.map((p) => {
      const isPeak = p.v === maxV;
      return `
        <circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${dotR+5}" fill="${_CP.lineA}" fill-opacity="0.09"/>
        ${isPeak ? `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${dotR+10}" fill="none" stroke="${_CP.lineA}" stroke-opacity="0.18" stroke-width="1.5"/>` : ""}
        <circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${dotR}" fill="#fff" stroke="${isPeak ? _CP.lineA : _CP.lineB}" stroke-width="2.6"/>
        <text x="${p.x.toFixed(1)}" y="${(p.y - dotR - 8).toFixed(1)}" text-anchor="middle" font-size="${fs}" font-weight="700" fill="${isPeak ? _CP.peakLbl : _CP.axisLbl}" font-family="system-ui,sans-serif">${_chartFmtVal(p.v)}</text>
      `;
    }).join("");

    const xLabels = pts.map((p) =>
      `<text x="${p.x.toFixed(1)}" y="${(H-14).toFixed(1)}" text-anchor="middle" font-size="${fsLbl}" font-weight="600" fill="${_CP.axisLbl}" font-family="system-ui,sans-serif">${p.lbl}</text>`
    ).join("");

    return `
      <defs>
        <linearGradient id="${uid}-area" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%"   stop-color="${_CP.lineA}" stop-opacity="0.28"/>
          <stop offset="80%"  stop-color="${_CP.lineB}" stop-opacity="0.04"/>
          <stop offset="100%" stop-color="${_CP.lineB}" stop-opacity="0"/>
        </linearGradient>
        <linearGradient id="${uid}-line" x1="0" y1="0" x2="1" y2="0">
          <stop offset="0%"   stop-color="${_CP.lineA}"/>
          <stop offset="100%" stop-color="${_CP.lineB}"/>
        </linearGradient>
      </defs>
      <rect width="${W}" height="${H}" fill="transparent"/>
      ${grid}
      <path d="${areaPath}" fill="url(#${uid}-area)"/>
      <path d="${linePath}" fill="none" stroke="url(#${uid}-line)" stroke-width="${lineW}" stroke-linecap="round" stroke-linejoin="round"/>
      ${dots}
      ${xLabels}
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${(pad.top+iH).toFixed(1)}" stroke="${_CP.grid}" stroke-width="1"/>
      <line x1="${pad.left}" y1="${(pad.top+iH).toFixed(1)}" x2="${(pad.left+iW).toFixed(1)}" y2="${(pad.top+iH).toFixed(1)}" stroke="${_CP.gridStrong}" stroke-width="1.5"/>
    `;
  }

  function _buildBarChart(values, labels, W, H, compact) {
    const s = compact ? 0.72 : 1;
    const pad = { top: Math.round(44*s), right: Math.round(32*s), bottom: Math.round(52*s), left: Math.round(60*s) };
    const iW = W - pad.left - pad.right;
    const iH = H - pad.top - pad.bottom;
    const n = values.length;
    const mx = Math.max(...values, 1);
    const slotW = iW / n;
    const barW = slotW * 0.52;
    const fs = compact ? 9.5 : 11;
    const fsLbl = compact ? 10 : 12;
    const uid = "bc" + Date.now() + Math.random().toString(36).slice(2,6);

    const grid = Array.from({ length: 5 }, (_, i) => {
      const y = pad.top + (iH * i) / 4;
      const v = ((4 - i) / 4) * mx;
      return `<line x1="${pad.left}" y1="${y.toFixed(1)}" x2="${(pad.left+iW).toFixed(1)}" y2="${y.toFixed(1)}" stroke="${_CP.grid}" stroke-width="1"/>
        <text x="${(pad.left-7).toFixed(1)}" y="${(y+4).toFixed(1)}" text-anchor="end" font-size="${fs}" fill="${_CP.axisLbl}" font-family="system-ui,sans-serif">${_chartFmtVal(v)}</text>`;
    }).join("");

    const defs = _CP.series.slice(0, n).map((c, i) =>
      `<linearGradient id="${uid}-b${i}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${c.hi}"/><stop offset="100%" stop-color="${c.lo}"/></linearGradient>`
    ).join("");

    const maxV = Math.max(...values);
    const bars = values.map((v, i) => {
      const barH = Math.max((v / mx) * iH, 2);
      const x = pad.left + slotW * i + (slotW - barW) / 2;
      const y = pad.top + iH - barH;
      const c = _CP.series[i % _CP.series.length];
      const isPeak = v === maxV;
      return `
        <rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${barH.toFixed(1)}" rx="6" fill="url(#${uid}-b${i % _CP.series.length})"/>
        ${isPeak ? `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${barH.toFixed(1)}" rx="6" fill="none" stroke="${c.hi}" stroke-width="1.5" stroke-opacity="0.5"/>` : ""}
        <text x="${(x+barW/2).toFixed(1)}" y="${(y-8).toFixed(1)}" text-anchor="middle" font-size="${fs}" font-weight="700" fill="${isPeak ? _CP.peakLbl : _CP.axisLbl}" font-family="system-ui,sans-serif">${_chartFmtVal(v)}</text>
        <text x="${(x+barW/2).toFixed(1)}" y="${(H-14).toFixed(1)}" text-anchor="middle" font-size="${fsLbl}" font-weight="600" fill="${_CP.axisLbl}" font-family="system-ui,sans-serif">${labels[i]}</text>
      `;
    }).join("");

    return `
      <defs>${defs}</defs>
      <rect width="${W}" height="${H}" fill="transparent"/>
      ${grid}
      ${bars}
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${(pad.top+iH).toFixed(1)}" stroke="${_CP.grid}" stroke-width="1"/>
      <line x1="${pad.left}" y1="${(pad.top+iH).toFixed(1)}" x2="${(pad.left+iW).toFixed(1)}" y2="${(pad.top+iH).toFixed(1)}" stroke="${_CP.gridStrong}" stroke-width="1.5"/>
    `;
  }

  function _buildRadarChart(values, labels, W, H, compact) {
    const s = compact ? 0.80 : 1;
    const cx = W / 2, cy = H / 2 + 4;
    const R = Math.min(W, H) * 0.33 * s;
    const n = values.length;
    const mx = Math.max(...values, 1);
    const uid = "rc" + Date.now() + Math.random().toString(36).slice(2,6);
    const fsLbl = compact ? 10.5 : 12.5;
    const fsRing = compact ? 8 : 9.5;
    const dotR = compact ? 4.5 : 5.5;

    const angles = Array.from({ length: n }, (_, i) => -Math.PI / 2 + (2 * Math.PI * i) / n);

    const polyPts = (frac) =>
      angles.map((a) => `${(cx+Math.cos(a)*R*frac).toFixed(2)},${(cy+Math.sin(a)*R*frac).toFixed(2)}`).join(" ");

    const rings = [0.25, 0.50, 0.75, 1.00].map((f) =>
      `<polygon points="${polyPts(f)}" fill="none" stroke="${f===1 ? _CP.gridStrong : _CP.radarGrid}" stroke-width="${f===1 ? 1.5 : 1}"/>`
    ).join("");

    const axes = angles.map((a) =>
      `<line x1="${cx.toFixed(2)}" y1="${cy.toFixed(2)}" x2="${(cx+Math.cos(a)*R).toFixed(2)}" y2="${(cy+Math.sin(a)*R).toFixed(2)}" stroke="${_CP.radarGrid}" stroke-width="1"/>`
    ).join("");

    const dataPts = values.map((v, i) => ({
      x: cx + Math.cos(angles[i]) * R * (v / mx),
      y: cy + Math.sin(angles[i]) * R * (v / mx),
    }));
    const dataPoly = dataPts.map((p) => `${p.x.toFixed(2)},${p.y.toFixed(2)}`).join(" ");

    const maxV = Math.max(...values);
    const dots = dataPts.map((p, i) => {
      const isPeak = values[i] === maxV;
      const c = _CP.series[i % _CP.series.length].hi;
      return `
        ${isPeak ? `<circle cx="${p.x.toFixed(2)}" cy="${p.y.toFixed(2)}" r="${dotR+5}" fill="${c}" fill-opacity="0.14"/>` : ""}
        <circle cx="${p.x.toFixed(2)}" cy="${p.y.toFixed(2)}" r="${dotR}" fill="#fff" stroke="${c}" stroke-width="2.6"/>
        <text x="${p.x.toFixed(2)}" y="${(p.y - dotR - 6).toFixed(2)}" text-anchor="middle" font-size="${compact ? 8.5 : 10}" font-weight="700" fill="${isPeak ? _CP.peakLbl : _CP.axisLbl}" font-family="system-ui,sans-serif">${_chartFmtVal(values[i])}</text>
      `;
    }).join("");

    const lblDist = R * 1.22;
    const lblElems = labels.map((lbl, i) => {
      const lx = cx + Math.cos(angles[i]) * lblDist;
      const ly = cy + Math.sin(angles[i]) * lblDist;
      const ca = Math.cos(angles[i]);
      const anchor = ca > 0.15 ? "start" : ca < -0.15 ? "end" : "middle";
      return `<text x="${lx.toFixed(2)}" y="${ly.toFixed(2)}" text-anchor="${anchor}" dominant-baseline="middle" font-size="${fsLbl}" font-weight="700" fill="#374151" font-family="system-ui,sans-serif">${lbl}</text>`;
    }).join("");

    const ringLabels = [0.25,0.50,0.75,1.00].map((f) => {
      const lx = cx + R * f + 5;
      return `<text x="${lx.toFixed(2)}" y="${(cy+4).toFixed(2)}" font-size="${fsRing}" fill="${_CP.gridStrong}" font-family="system-ui,sans-serif">${Math.round(f*100)}%</text>`;
    }).join("");

    return `
      <defs>
        <radialGradient id="${uid}-fill" cx="50%" cy="50%" r="50%">
          <stop offset="0%"   stop-color="${_CP.radarStroke}" stop-opacity="0.38"/>
          <stop offset="100%" stop-color="${_CP.radarStroke}" stop-opacity="0.07"/>
        </radialGradient>
      </defs>
      <rect width="${W}" height="${H}" fill="transparent"/>
      ${rings}
      ${ringLabels}
      ${axes}
      <polygon points="${dataPoly}" fill="url(#${uid}-fill)" stroke="${_CP.radarStroke}" stroke-width="2.4" stroke-linejoin="round"/>
      ${dots}
      ${lblElems}
    `;
  }

  // Renders SVG markup for a given chart type into an SVG element
  function _renderChartIntoSvg(svgEl, type, values, labels, compact) {
    if (!svgEl) return;
    const W = 800, H = 400;
    if (type === "bar")        svgEl.innerHTML = _buildBarChart(values, labels, W, H, compact);
    else if (type === "radar") svgEl.innerHTML = _buildRadarChart(values, labels, W, H, compact);
    else                       svgEl.innerHTML = _buildLineChart(values, labels, W, H, compact);
  }

  // ── Panel type icons (inline SVG strings) ────────────────────────────────
  const _PANEL_TYPE_ICONS = {
    line:  `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round"><path d="M3 3v18h18"/><path d="m7 15 4-4 3 3 5-7"/></svg>`,
    bar:   `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round"><rect x="3" y="12" width="4" height="9" rx="1"/><rect x="10" y="7" width="4" height="14" rx="1"/><rect x="17" y="3" width="4" height="18" rx="1"/></svg>`,
    radar: `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round"><polygon points="12 2 22 8.5 22 15.5 12 22 2 15.5 2 8.5"/><line x1="12" y1="2" x2="12" y2="22"/><line x1="2" y1="8.5" x2="22" y2="8.5"/><line x1="2" y1="15.5" x2="22" y2="15.5"/></svg>`,
  };
  const _PANEL_TYPE_LABELS = { line: "Trend", bar: "Bar", radar: "Radar" };

  // Builds and renders all chart panels based on current layout/types
  function _renderAllPanels() {
    const host = document.getElementById("createDrtChartPanels");
    if (!host || !_chartActiveRow) return;

    const values  = getChartSeries(_chartActiveRow);
    const labels  = ["Current", "Y1", "Y2", "Y3", "Y4", "Y5"];
    const compact = _chartLayout > 1;

    // Determine which types to show for each slot
    const slots = _chartLayout === 1
      ? [_chartType]
      : _chartLayout === 2
        ? [_chartPanelTypes[0], _chartPanelTypes[1]]
        : [_chartPanelTypes[0], _chartPanelTypes[1], _chartPanelTypes[2]];

    // Update modal class for CSS (hide global switcher when multi)
    const panel = chartModal?.querySelector(".create-drt-chart-modal__panel");
    if (panel) panel.classList.toggle("create-drt-chart-modal--multi", _chartLayout > 1);

    // Build panel HTML
    host.innerHTML = slots.map((type, idx) => {
      const types = ["line", "bar", "radar"];
      const tabsHtml = types.map((t) =>
        `<button class="create-drt-chart-panel-tab${t === type ? " active" : ""}" data-panel-idx="${idx}" data-panel-type="${t}" type="button">${_PANEL_TYPE_ICONS[t]} ${_PANEL_TYPE_LABELS[t]}</button>`
      ).join("");

      return `
        <div class="create-drt-chart-panel">
          ${_chartLayout > 1 ? `<div class="create-drt-chart-panel-tabs">${tabsHtml}</div>` : ""}
          <div class="create-drt-chart-panel__svg-wrap">
            <div class="create-drt-chart-zoom-controls" aria-label="Zoom controls">
              <button class="create-drt-chart-zoom-btn" data-zoom-in title="Zoom in">
                <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>
              </button>
              <button class="create-drt-chart-zoom-btn create-drt-chart-zoom-btn--reset" data-zoom-reset title="Reset zoom (double-click chart)">
                <svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/></svg>
              </button>
              <button class="create-drt-chart-zoom-btn" data-zoom-out title="Zoom out">
                <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M5 12h14"/></svg>
              </button>
            </div>
            <svg class="create-drt-chart-panel__svg" data-panel-svg="${idx}" viewBox="0 0 800 400" preserveAspectRatio="xMidYMid meet" aria-hidden="true"></svg>
            <div class="create-drt-chart-zoom-hint">Scroll to zoom &nbsp;·&nbsp; Drag to pan &nbsp;·&nbsp; Double-click to reset</div>
          </div>
        </div>
      `;
    }).join("");

    // Render SVG into each panel, then attach zoom behaviour
    slots.forEach((type, idx) => {
      const svgEl = host.querySelector(`[data-panel-svg="${idx}"]`);
      _renderChartIntoSvg(svgEl, type, values, labels, compact);
      _initChartZoom(svgEl.closest(".create-drt-chart-panel__svg-wrap"), svgEl);
    });
  }

  // ── Per-panel zoom / pan ──────────────────────────────────────────────────
  const _VB_DEFAULT = { x: 0, y: 0, w: 800, h: 400 };

  function _zoomSvg(svgEl, factor, cx, cy) {
    const vb = svgEl._zoomVB || { ..._VB_DEFAULT };
    const newW = Math.min(_VB_DEFAULT.w, Math.max(80,  vb.w / factor));
    const newH = Math.min(_VB_DEFAULT.h, Math.max(40,  vb.h / factor));
    const rx   = (cx - vb.x) / vb.w;   // relative position (0-1) within current vb
    const ry   = (cy - vb.y) / vb.h;
    let newX = cx - rx * newW;
    let newY = cy - ry * newH;
    newX = Math.max(0, Math.min(_VB_DEFAULT.w - newW, newX));
    newY = Math.max(0, Math.min(_VB_DEFAULT.h - newH, newY));
    svgEl._zoomVB = { x: newX, y: newY, w: newW, h: newH };
    svgEl.setAttribute("viewBox", `${newX.toFixed(2)} ${newY.toFixed(2)} ${newW.toFixed(2)} ${newH.toFixed(2)}`);
    _updateZoomResetBtn(svgEl);
  }

  function _resetZoomSvg(svgEl) {
    svgEl._zoomVB = { ..._VB_DEFAULT };
    svgEl.setAttribute("viewBox", `0 0 800 400`);
    _updateZoomResetBtn(svgEl);
  }

  function _updateZoomResetBtn(svgEl) {
    const wrap = svgEl.closest(".create-drt-chart-panel__svg-wrap");
    const btn  = wrap?.querySelector("[data-zoom-reset]");
    if (!btn) return;
    const vb = svgEl._zoomVB || _VB_DEFAULT;
    const isDefault = Math.abs(vb.w - 800) < 1 && Math.abs(vb.h - 400) < 1;
    btn.classList.toggle("create-drt-chart-zoom-btn--active", !isDefault);
  }

  function _clientToSvgCoords(svgEl, clientX, clientY) {
    const rect = svgEl.getBoundingClientRect();
    const vb   = svgEl._zoomVB || _VB_DEFAULT;
    return {
      x: ((clientX - rect.left) / rect.width)  * vb.w + vb.x,
      y: ((clientY - rect.top)  / rect.height) * vb.h + vb.y,
    };
  }

  function _initChartZoom(wrap, svgEl) {
    // Remove any previously attached zoom listeners by replacing the wrap node
    // (simpler than tracking AbortControllers across re-renders)
    const oldHandlers = wrap._zoomCleanup;
    if (oldHandlers) oldHandlers();

    svgEl._zoomVB = { ..._VB_DEFAULT };

    // ── Scroll-wheel zoom ────────────────────────────────────────────────────
    function onWheel(e) {
      e.preventDefault();
      e.stopPropagation();
      const coords = _clientToSvgCoords(svgEl, e.clientX, e.clientY);
      const factor = e.deltaY < 0 ? 1.18 : (1 / 1.18);
      _zoomSvg(svgEl, factor, coords.x, coords.y);
    }
    wrap.addEventListener("wheel", onWheel, { passive: false });

    // ── Double-click reset ───────────────────────────────────────────────────
    function onDblClick() { _resetZoomSvg(svgEl); }
    wrap.addEventListener("dblclick", onDblClick);

    // ── Drag-to-pan ──────────────────────────────────────────────────────────
    let _pan = null;
    function onMouseDown(e) {
      if (e.button !== 0 || e.target.closest(".create-drt-chart-zoom-controls")) return;
      const vb = svgEl._zoomVB || _VB_DEFAULT;
      const isZoomed = vb.w < _VB_DEFAULT.w - 1;
      if (!isZoomed) return; // only pan when zoomed in
      _pan = { sx: e.clientX, sy: e.clientY, vb: { ...vb } };
      wrap.classList.add("create-drt-chart-panning");
      e.preventDefault();
    }
    function onMouseMove(e) {
      if (!_pan) return;
      const rect = svgEl.getBoundingClientRect();
      const vb   = _pan.vb;
      const dx   = -((e.clientX - _pan.sx) / rect.width)  * vb.w;
      const dy   = -((e.clientY - _pan.sy) / rect.height) * vb.h;
      const newX = Math.max(0, Math.min(_VB_DEFAULT.w - vb.w, vb.x + dx));
      const newY = Math.max(0, Math.min(_VB_DEFAULT.h - vb.h, vb.y + dy));
      svgEl._zoomVB = { ...vb, x: newX, y: newY };
      svgEl.setAttribute("viewBox", `${newX.toFixed(2)} ${newY.toFixed(2)} ${vb.w.toFixed(2)} ${vb.h.toFixed(2)}`);
    }
    function onMouseUp() {
      _pan = null;
      wrap.classList.remove("create-drt-chart-panning");
    }
    wrap.addEventListener("mousedown", onMouseDown);
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);

    // ── Pinch-to-zoom (touch) ────────────────────────────────────────────────
    let _pinchDist = null;
    function onTouchStart(e) {
      if (e.touches.length === 2) {
        _pinchDist = Math.hypot(
          e.touches[0].clientX - e.touches[1].clientX,
          e.touches[0].clientY - e.touches[1].clientY
        );
      }
    }
    function onTouchMove(e) {
      if (e.touches.length !== 2 || _pinchDist === null) return;
      e.preventDefault();
      const dist = Math.hypot(
        e.touches[0].clientX - e.touches[1].clientX,
        e.touches[0].clientY - e.touches[1].clientY
      );
      const factor = dist / _pinchDist;
      const mx = (e.touches[0].clientX + e.touches[1].clientX) / 2;
      const my = (e.touches[0].clientY + e.touches[1].clientY) / 2;
      const coords = _clientToSvgCoords(svgEl, mx, my);
      _zoomSvg(svgEl, factor, coords.x, coords.y);
      _pinchDist = dist;
    }
    function onTouchEnd() { _pinchDist = null; }
    wrap.addEventListener("touchstart", onTouchStart, { passive: true });
    wrap.addEventListener("touchmove",  onTouchMove,  { passive: false });
    wrap.addEventListener("touchend",   onTouchEnd,   { passive: true });

    // ── Zoom button controls ─────────────────────────────────────────────────
    function onZoomIn() {
      const vb = svgEl._zoomVB || _VB_DEFAULT;
      _zoomSvg(svgEl, 1.35, vb.x + vb.w / 2, vb.y + vb.h / 2);
    }
    function onZoomOut() {
      const vb = svgEl._zoomVB || _VB_DEFAULT;
      _zoomSvg(svgEl, 1 / 1.35, vb.x + vb.w / 2, vb.y + vb.h / 2);
    }
    function onZoomReset() { _resetZoomSvg(svgEl); }

    wrap.querySelector("[data-zoom-in]")?.addEventListener("click",  onZoomIn);
    wrap.querySelector("[data-zoom-out]")?.addEventListener("click", onZoomOut);
    wrap.querySelector("[data-zoom-reset]")?.addEventListener("click", onZoomReset);

    // Cleanup for next render
    wrap._zoomCleanup = () => {
      wrap.removeEventListener("wheel",       onWheel);
      wrap.removeEventListener("dblclick",    onDblClick);
      wrap.removeEventListener("mousedown",   onMouseDown);
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup",   onMouseUp);
      wrap.removeEventListener("touchstart",  onTouchStart);
      wrap.removeEventListener("touchmove",   onTouchMove);
      wrap.removeEventListener("touchend",    onTouchEnd);
    };
  }

  function renderChartSvg(row) {
    _chartActiveRow = row || _chartActiveRow;
    _renderAllPanels();
  }

  function openChartModal(row) {
    if (!chartModal || !row || !String(row.skuName || "").trim()) return;
    _chartActiveRow = row;
    if (chartSubtitle) {
      chartSubtitle.textContent = `${row.skuName} · ${row.sku || "SKU"}`;
    }
    // Sync global type tabs
    chartModal.querySelectorAll(".create-drt-chart-type-btn").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.chartType === _chartType);
    });
    // Sync layout tabs
    chartModal.querySelectorAll(".create-drt-chart-layout-btn").forEach((btn) => {
      btn.classList.toggle("active", Number(btn.dataset.chartLayout) === _chartLayout);
    });
    _renderAllPanels();
    chartModal.hidden = false;
  }

  function closeChartModal() {
    if (!chartModal) return;
    _setChartMaximized(false); // always restore before closing
    chartModal.hidden = true;
  }

  // ── Maximize / restore ───────────────────────────────────────────────────
  const _chartPanel     = chartModal?.querySelector(".create-drt-chart-modal__panel");
  const _chartMaxBtn    = document.getElementById("createDrtChartMaximize");
  let   _chartMaximized = false;

  function _setChartMaximized(on) {
    _chartMaximized = on;
    _chartPanel?.classList.toggle("create-drt-chart-modal__panel--maximized", on);
    if (_chartMaxBtn) {
      _chartMaxBtn.setAttribute("aria-label", on ? "Restore chart size" : "Maximize chart");
      _chartMaxBtn.title = on ? "Restore (Esc)" : "Maximize";
    }
    // Re-render panels so SVGs resize to the new panel dimensions
    requestAnimationFrame(() => _renderAllPanels());
  }

  _chartMaxBtn?.addEventListener("click", () => _setChartMaximized(!_chartMaximized));

  function renderChartActionCell(instance, td, rowIndex) {
    Handsontable.dom.empty(td);
    td.classList.add("htCenter", "htMiddle", "create-drt-chart-icon-cell");
    const row = spreadsheetVisibleRows[rowIndex];
    const hasSkuName = !!String(row?.skuName || "").trim();
    const button = document.createElement("button");
    button.type = "button";
    button.className = "create-drt-chart-button";
    button.dataset.chartRow = String(rowIndex);
    button.disabled = !hasSkuName;
    button.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><polyline points="7 14 11 9 14 12 19 5"/><circle cx="19" cy="5" r="1.8" fill="currentColor" stroke="none"/><circle cx="14" cy="12" r="1.5" fill="currentColor" stroke="none"/><circle cx="11" cy="9" r="1.5" fill="currentColor" stroke="none"/></svg>';
    button.title = hasSkuName ? "Open SKU trend chart" : "";
    td.appendChild(button);
    return td;
  }

  function renderStandardSpreadsheetCell(instance, td, rowIndex, colIndex, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer(instance, td, rowIndex, colIndex, prop, value, cellProperties);
    return td;
  }

  function getVisibleProductCount(rows) {
    return (rows || []).reduce((count, row) => (
      String(row?.skuName || "").trim() ? count + 1 : count
    ), 0);
  }

  function updateWorkpadProductCountDisplay(count) {
    currentWorkpadProductCount = count;
    if (workpadCountNumber) workpadCountNumber.textContent = String(count);
    if (workpadCountLabel) workpadCountLabel.textContent = count === 1 ? "Product" : "Products";
  }

  // Glossy pill/rectangle for PG column price-tier indicator.
  // Each call gets a unique gradient ID so multiple cells don't share the same SVG defs.
  let _pgGlossSeq = 0;
  function renderGlossRect(color) {
    const id = "pg-gl-" + (++_pgGlossSeq);
    return [
      `<svg xmlns="http://www.w3.org/2000/svg" width="30" height="14" viewBox="0 0 30 14" aria-hidden="true">`,
      `<defs>`,
      `<linearGradient id="${id}" x1="0" y1="0" x2="0" y2="1">`,
      `<stop offset="0%"   stop-color="rgba(255,255,255,0.52)"/>`,
      `<stop offset="48%"  stop-color="rgba(255,255,255,0.08)"/>`,
      `<stop offset="49%"  stop-color="rgba(0,0,0,0.00)"/>`,
      `<stop offset="100%" stop-color="rgba(0,0,0,0.12)"/>`,
      `</linearGradient>`,
      `</defs>`,
      // Shadow
      `<rect x="1.5" y="2.5" width="27" height="11" rx="4" fill="rgba(0,0,0,0.18)"/>`,
      // Base colour
      `<rect x="1" y="1" width="27" height="11" rx="4" fill="${color}"/>`,
      // Gloss overlay
      `<rect x="1" y="1" width="27" height="11" rx="4" fill="url(#${id})"/>`,
      // Inner highlight border
      `<rect x="1" y="1" width="27" height="11" rx="4" fill="none" stroke="rgba(255,255,255,0.45)" stroke-width="0.8"/>`,
      `</svg>`,
    ].join("");
  }

  // PG column renderer — shows a glossy coloured pill when Term = 1, SKU Name is set,
  // and Monthly Unit Price Year 1 is known (red < $100, orange $100–$200, green > $200).
  function renderPgCell(instance, td, rowIndex, colIndex, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer(instance, td, rowIndex, colIndex, prop, value, cellProperties);
    // Remove any leftover indicator from a previous render pass.
    td.querySelectorAll(".create-drt-pg-indicator").forEach((el) => el.remove());

    const termEl = document.getElementById("createDrtTerm");
    const term = termEl ? String(termEl.value || "1").trim() : "1";
    if (term !== "1") return td;

    const row = spreadsheetVisibleRows[rowIndex];
    if (!row) return td;
    if (!String(row.skuName || "").trim()) return td;

    const price = parseFloat(Array.isArray(row.grp2Yr) ? row.grp2Yr[0] : "");
    if (isNaN(price)) return td;

    let color, tier;
    if (price < 100)        { color = "#d93636"; tier = "low"; }
    else if (price <= 200)  { color = "#e87318"; tier = "mid"; }
    else                    { color = "#18a84e"; tier = "high"; }

    const indicator = document.createElement("span");
    indicator.className = "create-drt-pg-indicator";
    indicator.dataset.pgTier = tier;
    indicator.dataset.pgPrice = price.toFixed(2);
    indicator.innerHTML = renderGlossRect(color);
    td.appendChild(indicator);
    return td;
  }

  function renderChartHeaderIcon() {
    return '<span class="create-drt-chart-header-icon" aria-hidden="true"><svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="m7 15 4-4 3 3 5-7"/></svg></span>';
  }

  function renderDeleteRowIcon() {
    return '<svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M10 11v6"/><path d="M14 11v6"/></svg>';
  }

  function decorateSpreadsheetHeader(col, TH) {
    if (col < 0 || !TH) return;
    const column = getVisibleSpreadsheetColumns()[col];
    if (!column) return;
    const headerRow = TH.parentElement;
    const rowIndex = headerRow?.parentElement ? Array.from(headerRow.parentElement.children).indexOf(headerRow) : -1;
    TH.style.display = "";
    TH.rowSpan = 1;
    // No rowSpan — visual "merge" is achieved via CSS (matching dark background,
    // no border between rows). This avoids HOT virtual-scroll column-count confusion.
    const applyMergedHeader = (label, ...classNames) => {
      if (rowIndex === 0) {
        TH.classList.add("create-drt-merged-header-cell", ...classNames);
        TH.innerHTML = `<span class="create-drt-merged-header-label">${label}</span>`;
        return true;
      }
      // Row 1: empty dark continuation — stays in DOM (no display:none) so HOT
      // column layout is never confused by missing cells.
      TH.classList.add("create-drt-merged-subheader-hidden");
      TH.innerHTML = "";
      return true;
    };
    if (column.key === "chartAction") {
      if (rowIndex === 0) {
        TH.classList.add("create-drt-chart-header-cell", "create-drt-merged-header-cell");
        TH.innerHTML = renderChartHeaderIcon();
        return;
      }
      TH.classList.add("create-drt-merged-subheader-hidden");
      TH.innerHTML = "";
      return;
    }
    if (column.key === "sku") {
      applyMergedHeader("SKU", "create-drt-sku-merged-header");
      return;
    }
    if (MERGED_DETAIL_HEADER_KEYS.includes(column.key)) {
      applyMergedHeader(column.label);
      return;
    }
    if (MERGED_POST_HEADER_KEYS.includes(column.key)) {
      applyMergedHeader(column.label);
      return;
    }
  }

  function decorateSpreadsheetRowHeader(row, TH) {
    if (row < 0 || !TH) return;
    const dataRow = spreadsheetVisibleRows[row];
    TH.classList.add("create-drt-row-header-cell");
    TH.innerHTML = "";

    const wrap = document.createElement("div");
    wrap.className = "create-drt-row-header-wrap";

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "create-drt-row-delete-btn";
    deleteBtn.dataset.deleteRowId = String(dataRow?.id || "");
    const originalRowNum = (drtRows.indexOf(dataRow) + 1) || (row + 1);
    deleteBtn.setAttribute("aria-label", `Delete row ${originalRowNum}`);
    deleteBtn.innerHTML = renderDeleteRowIcon();

    const num = document.createElement("span");
    num.className = "create-drt-row-header-number";
    num.textContent = String(originalRowNum);

    wrap.appendChild(deleteBtn);
    wrap.appendChild(num);
    TH.appendChild(wrap);
  }

  function buildSpreadsheetData(rows) {
    const visibleColumns = getVisibleSpreadsheetColumns();
    return rows.map((row) => visibleColumns.map((column) => getSpreadsheetCellValue(column, row)));
  }

  function buildSpreadsheetColumns() {
    const visibleColumns = getVisibleSpreadsheetColumns();
    return visibleColumns.map((column, index) => {
      const isReadonly = column.kind === "readonly"; // chart-action excluded — no grey bg
      const isSelect = column.kind === "select";
      const isDate = column.kind === "date";
      const base = {
        data: index,
        width: column.width || column.minWidth || 120,
        readOnly: isReadonly || column.kind === "chart-action", // still non-editable
        className: column.key === "pg"
          ? "create-drt-pg-cell htCenter htMiddle"
          : column.kind === "readonly"
            ? "htDimmed"
            : column.kind === "chart-action"
              ? "create-drt-chart-icon-cell htCenter htMiddle"
              : "",
      };
      if (isSelect) {
        return { ...base, type: "dropdown", source: column.source || [], strict: false };
      }
      if (isDate) {
        return {
          ...base,
          type: "date",
          dateFormat: "MM/DD/YYYY",
          correctFormat: true,
          allowInvalid: true,
          renderer: renderStandardSpreadsheetCell,
        };
      }
      return {
        ...base,
        type: "text",
        renderer:
          column.kind === "chart-action" ? renderChartActionCell :
          column.key === "pg" ? renderPgCell :
          renderStandardSpreadsheetCell,
      };
    });
  }

  // Keys whose row-0 header visually merges both rows (CSS continuity, no DOM rowSpan).
  // Row-1 for these columns returns "" so HOT renders an empty cell that we style dark.
  const MERGED_HEADER_KEYS = new Set([
    "chartAction", "sku",
    ...MERGED_DETAIL_HEADER_KEYS,
    ...MERGED_POST_HEADER_KEYS,
  ]);

  function buildSpreadsheetColHeaders() {
    return getVisibleSpreadsheetColumns().map((column) =>
      MERGED_HEADER_KEYS.has(column.key) ? "" : column.label
    );
  }

  function buildSpreadsheetNestedHeaders() {
    const visibleColumns = getVisibleSpreadsheetColumns();
    const hasChartAction = visibleColumns.some((column) => column.key === "chartAction");
    const visibleCalcCount = visibleColumns.filter((column) => column.kind === "calc").length;
    const detailCount = visibleColumns.filter((column) => DETAIL_KEYS.includes(column.key)).length;
    const subscriptionCount = visibleColumns.filter((column) => SUBSCRIPTION_COLUMNS.some((item) => item.key === column.key)).length;
    const monthlyCount = visibleColumns.filter((column) => MONTHLY_PRICE_COLUMNS.some((item) => item.key === column.key)).length;
    const annualCount = visibleColumns.filter((column) => TOTAL_ANNUAL_COLUMNS.some((item) => item.key === column.key) || column.key === "newTotal").length;
    const productInfoGroupCount = visibleColumns.filter((column) => PRODUCT_GROUP_KEYS.includes(column.key)).length;
    const postCount = visibleColumns.filter((column) => MERGED_POST_HEADER_KEYS.includes(column.key)).length;
    const topRow = [];
    if (visibleCalcCount) topRow.push({ label: "Workpad", colspan: visibleCalcCount });
    if (hasChartAction) topRow.push({ label: "", colspan: 1 });
    if (detailCount) {
      topRow.push(...Array.from({ length: detailCount }, () => ({ label: "", colspan: 1 })));
    }
    if (subscriptionCount) topRow.push({ label: "Number of Subscriptions", colspan: subscriptionCount });
    if (monthlyCount) topRow.push({ label: "Monthly Unit Price", colspan: monthlyCount });
    if (annualCount) topRow.push({ label: "Total AOV", colspan: annualCount });
    if (productInfoGroupCount) topRow.push({ label: "Product Info", colspan: productInfoGroupCount });
    if (postCount) {
      topRow.push(...Array.from({ length: postCount }, () => ({ label: "", colspan: 1 })));
    }
    return [topRow, buildSpreadsheetColHeaders()];
  }

  function getSpreadsheetWorksheet() {
    return spreadsheetInstance && !spreadsheetInstance.isDestroyed ? spreadsheetInstance : null;
  }

  function getSpreadsheetUndoRedoPlugin(hot = getSpreadsheetWorksheet()) {
    if (!hot || typeof hot.getPlugin !== "function") return null;
    try {
      return hot.getPlugin("undoRedo");
    } catch {
      return null;
    }
  }

  function refreshSpreadsheetHistoryButtons() {
    const plugin = getSpreadsheetUndoRedoPlugin();
    const canUndo = !!(plugin && typeof plugin.isUndoAvailable === "function" && plugin.isUndoAvailable());
    const canRedo = !!(plugin && typeof plugin.isRedoAvailable === "function" && plugin.isRedoAvailable());

    if (undoButton) {
      undoButton.disabled = !canUndo;
      undoButton.setAttribute("aria-disabled", String(!canUndo));
    }
    if (redoButton) {
      redoButton.disabled = !canRedo;
      redoButton.setAttribute("aria-disabled", String(!canRedo));
    }
  }

  function setRowValueFromSpreadsheet(column, row, value) {
    if (!column || !row || ["readonly", "chart-action"].includes(column.kind)) return false;
    if (column.kind === "calc") {
      const calcValues = ensureCalcValues(row);
      const index = getCalcValueIndex(column);
      if (index < 0) return false;
      calcValues[index] = value;
      return true;
    }
    const yearMatch = column.key.match(/^(grp[123])Year([1-5])$/);
    if (yearMatch) {
      row[yearMatch[1] + "Yr"][parseInt(yearMatch[2], 10) - 1] = value;
      return true;
    }
    if (column.field) {
      row[column.field] = value;
      return true;
    }
    return false;
  }

  function getFilteredRows() {
    const filterEl = document.getElementById("productFilter");
    const value = (filterEl?.value || "").trim();
    const termEl = document.getElementById("createDrtTerm");
    const term = termEl ? String(termEl.value || "1").trim() : "1";
    return drtRows.filter((row) => {
      const skuValue = String(row.skuName || "").trim();
      // SKU name filter
      if (value && (value === "__empty__" ? skuValue : skuValue !== value)) return false;
      // Hide PG green-dot rows: Term=1, SKU set, Yr1 price > $200
      if (hidePgGreen && term === "1" && skuValue) {
        const price = parseFloat(Array.isArray(row.grp2Yr) ? row.grp2Yr[0] : "");
        if (!isNaN(price) && price > 200) return false;
      }
      return true;
    });
  }

  // Handsontable afterChange: changes = [[row, col, oldVal, newVal], ...]
  function handleSpreadsheetChange(changes, source) {
    if (spreadsheetIsRendering) return;
    if (!changes) return;
    let needsRender = false;
    let lastEditedY = null;
    let lastEditedX = null;
    let editedSkuName = false;
    changes.forEach(([y, x, , value]) => {
      const column = getVisibleSpreadsheetColumns()[x];
      const row = spreadsheetVisibleRows[y];
      if (!column || !row) return;
      const pendingFormula = consumePendingFormula(row, column);
      const nextValue = pendingFormula
        ? {
            displayValue: (() => {
              const result = evaluateFormula(pendingFormula);
              return result === null ? String(value ?? "") : formatFormulaResult(result);
            })(),
            formula: pendingFormula,
          }
        : normalizeSpreadsheetEntry(column, value);
      if (!setRowValueFromSpreadsheet(column, row, nextValue.displayValue)) return;
      setStoredFormula(row, column, nextValue.formula);
      if (pendingFormula || nextValue.formula) {
        syncWorksheetCellValue(x, y, nextValue.displayValue);
      }
      if (column.key === "grp1Current" || column.key.startsWith("grp1Year") || column.key === "grp2Current" || column.key.startsWith("grp2Year")) {
        recalculateDerivedRow(row);
      }
      if (column.key === "skuName") {
        if (source === "edit") { editedSkuName = true; }
        else { populateProductFilter(); }
      }
      needsRender = true;
      if (source === "edit") { lastEditedY = y; lastEditedX = x; }
    });
    if (needsRender) {
      updateWorkpadProductCountDisplay(getVisibleProductCount(drtRows));
      recalculateSpreadsheetFormulas();

      if (source === "edit" && lastEditedY !== null && lastEditedX !== null) {
        // ── Fast path: return immediately so HOT can finish its internal Enter-key
        // processing (editor close → enterMoves → selection update) without any
        // interference.  We defer derived-cell updates (grp3, newTotal, formulas)
        // to a setTimeout(0) which fires after HOT is completely done.
        const hot = getSpreadsheetWorksheet();
        if (hot) {
          setTimeout(() => {
            try {
              if (hot.isDestroyed) return;
              // Deferred SKU filter update — skip applyProductFilter() to avoid
              // a full renderSpreadsheet() that would destroy the HOT instance.
              if (editedSkuName) populateProductFilter(/* skipApply */ true);
              const newData = buildSpreadsheetData(spreadsheetVisibleRows);
              spreadsheetIsRendering = true;
              const updates = [];
              for (let r = 0; r < newData.length; r++) {
                for (let c = 0; c < newData[r].length; c++) {
                  if (String(hot.getDataAtCell(r, c) ?? "") !== String(newData[r][c] ?? "")) {
                    updates.push([r, c, newData[r][c]]);
                  }
                }
              }
              if (updates.length) hot.setDataAtCell(updates);
              spreadsheetIsRendering = false;
              hot.render();
              refreshSpreadsheetHistoryButtons();
            } catch {}
          }, 0);
          return; // ← let HOT finish enterMoves naturally
        }
        // Fallback if HOT instance gone: full re-render with pending selection
        const maxY = spreadsheetVisibleRows.length - 1;
        _pendingEditSelection = {
          x: lastEditedX,
          y: Math.min(lastEditedY + 1, maxY),
        };
      }

      renderSpreadsheet();
      requestAnimationFrame(refreshSpreadsheetHistoryButtons);
    }
  }

  // Handsontable beforeChange: changes = [[row, col, oldVal, newVal], ...]
  function handleSpreadsheetBeforeChange(changes) {
    if (!changes) return;
    changes.forEach((change) => {
      const [y, x, , rawInputValue] = change;
      const column = getVisibleSpreadsheetColumns()[x];
      const row = spreadsheetVisibleRows[y];
      if (!column || !row || !canEditSpreadsheetColumn(column)) return;
      const rawValue = String(rawInputValue ?? "");
      const trimmedValue = rawValue.trim();

      // Numeric-only columns: strip anything that isn't a digit, decimal point,
      // or leading minus sign.  Formulas (starting with "=") are still allowed.
      if (column.numeric && !trimmedValue.startsWith("=")) {
        const sanitized = trimmedValue.replace(/[^0-9.\-]/g, "");
        // Keep at most one decimal point and one leading minus
        const normalized = sanitized.replace(/(?!^)-/g, "").replace(/(\..*)\./g, "$1");
        change[3] = normalized;
        setPendingFormula(row, column, "");
        return;
      }

      if (!trimmedValue.startsWith("=")) {
        setPendingFormula(row, column, "");
        return;
      }
      const result = evaluateFormula(trimmedValue);
      setPendingFormula(row, column, trimmedValue);
      // Mutate the change in-place so Handsontable stores the computed value
      change[3] = result === null ? rawValue : formatFormulaResult(result);
    });
  }

  function getWorksheetSelectionCoords(hot) {
    if (!hot) return null;
    try {
      const selected = hot.getSelected();
      if (selected && selected.length) {
        const [row, col] = selected[0];
        if (Number.isFinite(row) && Number.isFinite(col)) {
          return { x: col, y: row };
        }
      }
    } catch {}
    return null;
  }

  function syncFormulaBarFromWorksheetSelection(hot) {
    const coords = getWorksheetSelectionCoords(hot || getSpreadsheetWorksheet());
    if (coords) {
      updateFormulaBarState(coords.x, coords.y);
      return true;
    }
    return false;
  }

  // Handsontable afterSelection: (row, col, row2, col2)
  function handleSpreadsheetSelection(row, col, row2, col2) {
    // HOT fires afterSelection during its own constructor (default cell 0,0).
    // Ignore those calls — they would overwrite activeSpreadsheetCell with (0,0)
    // and break the Enter-moves-down logic that sets activeSpreadsheetCell before
    // re-rendering.  restoreSpreadsheetSelection() handles the initial placement.
    if (spreadsheetIsRendering) return;

    // After range selection ends, skip the next afterSelection call so it
    // doesn't overwrite the formula bar with the selected cell's value.
    if (_fRangeCooldown) {
      _fRangeCooldown = false;
      return;
    }

    // ── Formula-range injection ──────────────────────────────────────────────
    // When the user is typing a formula (in the formula bar or HOT cell editor)
    // and then clicks/drags to select a range, inject the range reference into
    // the formula at the cursor position — just like Excel.
    if (_fRangeActive) {
      const ref = buildCellRangeRef(col, row, col2 ?? col, row2 ?? row);
      _injectRangeRef(ref);
      // Show range in the cell-reference box & green visual cue on formula bar
      if (formulaCellRef) formulaCellRef.textContent = ref;
      updateFormulaSelectionMode(true);
      return; // Don't change activeSpreadsheetCell while selecting a range
    }

    activeSpreadsheetSelection = {
      startX: Math.min(col, col2 ?? col),
      endX: Math.max(col, col2 ?? col),
      startY: Math.min(row, row2 ?? row),
      endY: Math.max(row, row2 ?? row),
    };
    updateFormulaBarState(col, row);
    window.setTimeout(() => {
      syncFormulaBarFromWorksheetSelection();
    }, 0);
  }

  function refreshSpreadsheetSelection() {
    const hot = getSpreadsheetWorksheet();
    if (!hot || !activeSpreadsheetSelection) return;
    const { startY, startX, endY, endX } = activeSpreadsheetSelection;
    try {
      if (typeof hot.deselectCell === "function") {
        hot.deselectCell();
      }
    } catch {}
    try {
      hot.selectCell(startY, startX, endY, endX, false, false);
    } catch {}
  }

  function getSpreadsheetCellFromDom(target) {
    const element = target instanceof Element ? target : target?.parentElement;
    if (!element || !spreadsheetHost) return null;
    const hot = getSpreadsheetWorksheet();
    if (!hot) return null;
    try {
      const coords = hot.getCoords(element.closest("td, th"));
      if (coords) return { x: coords.col, y: coords.row };
    } catch {}
    return null;
  }

  function updateFormulaBarFromSpreadsheetDom(target) {
    window.setTimeout(() => {
      if (syncFormulaBarFromWorksheetSelection()) return;
      const coords = getSpreadsheetCellFromDom(target);
      if (!coords) return;
      updateFormulaBarState(coords.x, coords.y);
    }, 0);
  }

  function renderSpreadsheet() {
    if (!spreadsheetHost) return;
    spreadsheetVisibleRows = getFilteredRows();
    updateWorkpadProductCountDisplay(getVisibleProductCount(drtRows));
    spreadsheetVisibleRows.forEach(ensureCalcValues);
    recalculateSpreadsheetFormulas();

    if (typeof window.Handsontable === "undefined") {
      spreadsheetHost.innerHTML = "<p>Spreadsheet library failed to load. Check network access to jsDelivr and refresh the page.</p>";
      return;
    }

    const data = buildSpreadsheetData(spreadsheetVisibleRows);
    // Persist horizontal scroll so the viewport doesn't jump on re-render.
    const savedScrollLeft = spreadsheetScrollHolder?.scrollLeft || 0;
    const savedScrollTop  = spreadsheetScrollHolder?.scrollTop  || 0;
    spreadsheetIsRendering = true;
    if (spreadsheetInstance) {
      try {
        spreadsheetInstance.destroy();
      } catch {}
      spreadsheetInstance = null;
    }
    if (spreadsheetScrollHolder) {
      spreadsheetScrollHolder.removeEventListener("scroll", syncExcelOutlineWithScroll);
      spreadsheetScrollHolder = null;
    }
    refreshSpreadsheetHistoryButtons();
    spreadsheetHost.innerHTML = "";
    spreadsheetInstance = new Handsontable(spreadsheetHost, {
      data,
      colHeaders: buildSpreadsheetColHeaders(),
      nestedHeaders: buildSpreadsheetNestedHeaders(),
      columnHeaderHeight: [34, 52],
      columns: buildSpreadsheetColumns(),
      rowHeaders: true,
      rowHeaderWidth: SPREADSHEET_ROW_HEADER_WIDTH,
      rowHeights: 22,
      width: "100%",
      height: "100%",
      licenseKey: "non-commercial-and-evaluation",
      stretchH: "none",
      manualColumnResize: true,
      manualRowResize: true,
      undo: true,
      columnSorting: false,
      fillHandle: false,
      contextMenu: false,
      autoWrapRow: false,
      autoWrapCol: false,
      enterMoves: { row: 1, col: 0 },
      outsideClickDeselects: false,
      fixedColumnsStart: 0,
      afterGetColHeader: decorateSpreadsheetHeader,
      afterGetRowHeader: decorateSpreadsheetRowHeader,
      afterRender: function() {
        // afterGetRowHeader is unreliable for row 0 when nestedHeaders is active —
        // HOT may reset the TH content after the hook fires. Re-apply any missing decorations.
        const ths = this.rootElement?.querySelectorAll(".ht_clone_inline_start tbody tr > th");
        if (!ths?.length) return;
        ths.forEach((th) => {
          if (th.querySelector(".create-drt-row-header-wrap")) return;
          try {
            const coords = this.getCoords(th);
            if (coords && coords.row >= 0) decorateSpreadsheetRowHeader(coords.row, th);
          } catch {}
        });
      },
      beforeChange: handleSpreadsheetBeforeChange,
      afterChange: handleSpreadsheetChange,
      afterSelection: handleSpreadsheetSelection,
      afterUndo: refreshSpreadsheetHistoryButtons,
      afterRedo: refreshSpreadsheetHistoryButtons,
      afterBeginEditing: function() {
        // When HOT's cell editor opens, sync _formulaBarInFormula so the
        // capture-phase mousedown knows formula-range mode is available.
        // We poll via keydown on the editor textarea.
        try {
          const editor = this.getActiveEditor();
          const textarea = editor?.TEXTAREA || editor?.textareaStyle?.parentNode?.querySelector("textarea");
          if (textarea) {
            const onEditorKey = () => {
              const val = String(textarea.value || "");
              _formulaBarInFormula = val.startsWith("=");
              if (_formulaBarInFormula) {
                _fRangeInsertPos = val.length;
                _fRangeInsertLen = 0;
              }
              // mirror into formula bar display
              if (formulaInput && !formulaInput.readOnly) {
                formulaInput.value = val;
              }
              updateFormulaSelectionMode(_formulaBarInFormula && hasUnclosedBracket());
            };
            textarea.addEventListener("input", onEditorKey);
            textarea.addEventListener("keyup", onEditorKey);
            // remove on next editor-close so listeners don't accumulate
            textarea._drtCleanup = () => {
              textarea.removeEventListener("input", onEditorKey);
              textarea.removeEventListener("keyup", onEditorKey);
            };
          }
        } catch {}
      },
      afterEditorClosed: function() {
        // If _fRangeActive is true, the editor was closed because the user
        // clicked on the grid to select a range for a formula.  DON'T clear
        // the flag — afterSelection still needs it.  The mouseup handler
        // will clear it after injection is complete.
        if (!_fRangeActive) {
          _formulaBarInFormula = false;
          updateFormulaSelectionMode(false);
        }
        try {
          const editor = this.getActiveEditor();
          const textarea = editor?.TEXTAREA || editor?.textareaStyle?.parentNode?.querySelector("textarea");
          if (textarea?._drtCleanup) { textarea._drtCleanup(); delete textarea._drtCleanup; }
        } catch {}
      },
    });
    spreadsheetIsRendering = false;
    bindSpreadsheetScrollSync();
    refreshSpreadsheetHistoryButtons();
    // Restore horizontal/vertical scroll position after HOT re-creation.
    if (spreadsheetScrollHolder && (savedScrollLeft || savedScrollTop)) {
      spreadsheetScrollHolder.scrollLeft = savedScrollLeft;
      spreadsheetScrollHolder.scrollTop  = savedScrollTop;
      syncExcelOutlineWithScroll();
    }
    // Remeasure header heights then restore selection — both deferred one frame
    // so HOT's internal viewport is fully laid out before selectCell() is called.
    // Capture the flag NOW (before the rAF) so we know whether this re-render
    // was triggered by an Enter-key edit even after _pendingEditSelection is consumed.
    const _hadPendingEdit = !!_pendingEditSelection;
    requestAnimationFrame(() => {
      try { spreadsheetInstance?.refreshDimensions(); } catch {}
      restoreSpreadsheetSelection();
      // If the re-render was triggered by Enter, the browser will fire the keyup
      // event for that Enter key AFTER this rAF callback — which steals focus back
      // to document.body.  A nested setTimeout(0) fires after all pending events
      // (including keyup) have processed, so listen() reliably reclaims focus.
      if (_hadPendingEdit) {
        setTimeout(() => {
          try {
            if (spreadsheetInstance && !spreadsheetInstance.isDestroyed) {
              spreadsheetInstance.listen();
            }
          } catch {}
        }, 0);
      }
    });
  }

  function syncExcelOutlineWithScroll() {
    const scrollLeft = spreadsheetScrollHolder?.scrollLeft || 0;
    if (excelOutline) {
      const workpadHideAfter = excelCanvasCollapsed
        ? EXCEL_OUTLINE_TOGGLE_SIZE + 90
        : SPREADSHEET_ROW_HEADER_WIDTH + EXCEL_CANVAS_TOTAL_WIDTH;
      excelOutline.style.setProperty("--excel-outline-shift", `${scrollLeft}px`);
      excelOutline.classList.toggle("create-drt-excel-outline-hidden", scrollLeft > workpadHideAfter);
    }
    if (addlInfoOutline) {
      addlInfoOutline.style.setProperty("--excel-outline-shift", `${scrollLeft}px`);
      const hideAfter = currentAddlInfoMetrics.left - SPREADSHEET_ROW_HEADER_WIDTH + currentAddlInfoMetrics.expandedWidth;
      addlInfoOutline.classList.toggle("create-drt-excel-outline-hidden", scrollLeft > hideAfter);
    }
    if (pricingInfoOutline) {
      pricingInfoOutline.style.setProperty("--excel-outline-shift", `${scrollLeft}px`);
      const hideAfter = currentPricingInfoMetrics.left - SPREADSHEET_ROW_HEADER_WIDTH + currentPricingInfoMetrics.expandedWidth;
      pricingInfoOutline.classList.toggle("create-drt-excel-outline-hidden", scrollLeft > hideAfter);
    }
    if (productInfoOutline) {
      productInfoOutline.style.setProperty("--excel-outline-shift", `${scrollLeft}px`);
      const hideAfter = currentProductInfoGroupMetrics.left - SPREADSHEET_ROW_HEADER_WIDTH + currentProductInfoGroupMetrics.expandedWidth;
      productInfoOutline.classList.toggle("create-drt-excel-outline-hidden", scrollLeft > hideAfter);
    }
  }

  function bindSpreadsheetScrollSync() {
    spreadsheetScrollHolder = spreadsheetHost?.querySelector(".ht_master .wtHolder") || null;
    if (!spreadsheetScrollHolder) {
      excelOutline?.style.setProperty("--excel-outline-shift", "0px");
      addlInfoOutline?.style.setProperty("--excel-outline-shift", "0px");
      pricingInfoOutline?.style.setProperty("--excel-outline-shift", "0px");
      productInfoOutline?.style.setProperty("--excel-outline-shift", "0px");
      syncExcelOutlineWithScroll();
      return;
    }
    spreadsheetScrollHolder.addEventListener("scroll", syncExcelOutlineWithScroll, { passive: true });
    syncExcelOutlineWithScroll();
  }

  function setSpreadsheetOnlyMode() {
    if (page) page.classList.add("excel-mode-active");
    if (excelCanvasPanel) excelCanvasPanel.hidden = false;
    table?.classList.remove("input-style-excel");
    table?.classList.add("input-style-modern");
    buildTableHeader();
    setupColumnResize();
    syncTableLayout();
    renderSpreadsheet();
  }

  function getViewMode() {
    return table?.classList.contains("input-style-excel") ? "excel" : "table";
  }

  function appendHeaderContent(th, column, mode, options) {
    if (options?.blank) {
      th.innerHTML = "";
      return;
    }
    if (column.groupTitle) {
      const stack = document.createElement("span");
      stack.className = "create-drt-header-stack";

      const kicker = document.createElement("span");
      kicker.className = "create-drt-header-kicker";
      kicker.textContent = column.groupTitle;

      const text = document.createElement("span");
      text.className = "create-drt-header-text";
      text.textContent = column.label;

      stack.appendChild(kicker);
      stack.appendChild(text);
      th.appendChild(stack);
      return;
    }

    const text = document.createElement("span");
    text.className = "create-drt-header-text";
    text.textContent = column.label;
    th.appendChild(text);
  }

  function createLeafHeaderCell(column, options) {
    const th = document.createElement("th");
    th.className = [
      "create-drt-header-cell",
      options?.mode === "excel" ? "create-drt-header-cell--excel" : "create-drt-header-cell--table",
      options?.blank ? "create-drt-header-placeholder" : "",
      column.headerClass,
      column.excelOnly ? "create-drt-excel-only" : "",
      "create-drt-col-" + column.key,
    ].filter(Boolean).join(" ");
    th.dataset.colKey = column.key;
    th.dataset.colIndex = String(column.index);
    if (column.lockResize) th.dataset.lockResize = "true";
    appendHeaderContent(th, column, options?.mode || "table", options);
    return th;
  }

  function createGroupHeaderCell(label, keys, className, options) {
    const th = document.createElement("th");
    th.className = [
      "create-drt-group-header",
      className,
      options?.mode === "excel" ? "create-drt-header-cell--excel" : "create-drt-header-cell--table",
      options?.excelOnly ? "create-drt-excel-only" : "",
    ].filter(Boolean).join(" ");
    const text = document.createElement("span");
    text.className = "create-drt-header-text";
    text.textContent = label;
    th.appendChild(text);
    th.colSpan = keys.length;
    if (options?.stickyAnchor) th.dataset.stickyAnchor = options.stickyAnchor;
    return th;
  }

  function buildTableHeader() {
    if (!thead) return;
    thead.innerHTML = "";
    const mode = getViewMode();

    if (mode === "excel") {
      const alphabetRow = document.createElement("tr");
      alphabetRow.className = "create-drt-header-row-0 create-drt-alphabet-row";
      TABLE_COLUMNS.forEach((column) => {
        const th = document.createElement("th");
        th.className = [
          column.meta ? column.headerClass : "col-alphabet",
          column.excelOnly ? "create-drt-excel-only" : "",
          "create-drt-alphabet-cell",
          "create-drt-col-" + column.key,
        ].filter(Boolean).join(" ");
        th.textContent = column.meta ? "" : (column.letter || "");
        th.dataset.colKey = column.key;
        th.dataset.colIndex = String(column.index);
        if (column.lockResize) th.dataset.lockResize = "true";
        alphabetRow.appendChild(th);
      });
      thead.appendChild(alphabetRow);
    }

    const headerRow = document.createElement("tr");
    headerRow.className = "create-drt-header-row-1";

    if (mode === "excel") {
      headerRow.appendChild(createLeafHeaderCell(COLUMN_BY_KEY.rowNum, { mode, blank: true }));
      headerRow.appendChild(createLeafHeaderCell(COLUMN_BY_KEY.delete, { mode, blank: true }));
      headerRow.appendChild(createGroupHeaderCell("Workpad", EXCEL_CANVAS_KEYS, "col-excel-canvas", { mode, excelOnly: true, stickyAnchor: "calc1" }));
      TABLE_COLUMNS
        .filter((column) => !column.meta && !column.excelOnly)
        .forEach((column) => headerRow.appendChild(createLeafHeaderCell(column, { mode })));
    } else {
      TABLE_COLUMNS
        .filter((column) => !column.excelOnly)
        .forEach((column) => headerRow.appendChild(createLeafHeaderCell(column, { mode })));
    }

    thead.appendChild(headerRow);
  }

  function buildColgroup() {
    if (!colgroup) return;
    colgroup.innerHTML = "";
    TABLE_COLUMNS.forEach((column) => {
      const col = document.createElement("col");
      col.className = [
        "create-drt-col",
        "create-drt-col-" + column.key,
        column.excelOnly ? "create-drt-excel-only" : "",
      ].filter(Boolean).join(" ");
      col.dataset.colKey = column.key;
      col.dataset.colIndex = String(column.index);
      if (column.width) col.style.width = String(column.width) + "px";
      if (column.minWidth) col.style.minWidth = String(column.minWidth) + "px";
      colgroup.appendChild(col);
    });
  }

  function getColumnWidth(column) {
    const col = colgroup?.querySelector(`col[data-col-key="${column.key}"]`);
    const explicitWidth = parseFloat(col?.style.width || "");
    if (!Number.isNaN(explicitWidth) && explicitWidth > 0) return explicitWidth;
    const headerCell = table?.querySelector(`th[data-col-key="${column.key}"]`);
    const bodyCell = table?.querySelector(`td[data-col-key="${column.key}"]`);
    return headerCell?.getBoundingClientRect().width || bodyCell?.getBoundingClientRect().width || column.minWidth || column.width || 80;
  }

  function syncStickyColumns() {
    if (!table) return;
    const mode = table.classList.contains("input-style-excel") ? "excel" : "table";
    table.querySelectorAll(".create-drt-sticky-col").forEach((cell) => {
      cell.classList.remove("create-drt-sticky-col");
      cell.style.left = "";
    });

    let left = 0;
    TABLE_COLUMNS
      .filter((column) => Array.isArray(column.stickyModes) && column.stickyModes.includes(mode))
      .forEach((column) => {
        table.querySelectorAll(`[data-col-key="${column.key}"]`).forEach((cell) => {
          cell.classList.add("create-drt-sticky-col");
          cell.style.left = String(left) + "px";
        });
        left += getColumnWidth(column);
      });

    const excelCanvasGroup = table.querySelector('[data-sticky-anchor="calc1"]');
    if (mode === "excel" && excelCanvasGroup) {
      const anchorCell = table.querySelector('[data-col-key="calc1"]');
      excelCanvasGroup.classList.add("create-drt-sticky-col");
      excelCanvasGroup.style.left = anchorCell?.style.left || "76px";
    }
  }

  function setupColumnResize() {
    if (!table || !colgroup) return;
    table.querySelectorAll("th[data-col-index]").forEach((th) => {
      if (th.dataset.lockResize === "true" || th.querySelector(".create-drt-col-resizer")) return;
      const resizer = document.createElement("div");
      resizer.className = "create-drt-col-resizer";
      resizer.title = "Drag to resize column";
      resizer.addEventListener("mousedown", (e) => {
        e.preventDefault();
        e.stopPropagation();
        const startX = e.pageX;
        const colIndex = parseInt(th.dataset.colIndex || "-1", 10);
        const col = colgroup.querySelector(`col[data-col-index="${colIndex}"]`);
        if (!col) return;
        const column = TABLE_COLUMNS[colIndex];
        const startWidth = getColumnWidth(column);
        const onMove = (ev) => {
          const dx = ev.pageX - startX;
          const newWidth = Math.max(24, startWidth + dx);
          col.style.width = String(newWidth) + "px";
          col.style.minWidth = String(newWidth) + "px";
          syncStickyColumns();
        };
        const onUp = () => {
          document.removeEventListener("mousemove", onMove);
          document.removeEventListener("mouseup", onUp);
          document.body.style.cursor = "";
          document.body.style.userSelect = "";
        };
        document.body.style.cursor = "col-resize";
        document.body.style.userSelect = "none";
        document.addEventListener("mousemove", onMove);
        document.addEventListener("mouseup", onUp);
      });
      th.style.position = "relative";
      th.appendChild(resizer);
    });
  }

  function syncTableLayout() {
    syncStickyColumns();
  }

  function buildTableStructure() {
    buildColgroup();
    buildTableHeader();
    setupColumnResize();
    syncTableLayout();
  }

  function getNextRowIndex() {
    let max = 0;
    drtRows.forEach((row) => {
      const id = parseInt(row.id, 10);
      if (!isNaN(id) && id > max) max = id;
    });
    return max + 1;
  }

  function addRow() {
    const input = document.getElementById("addRowsCount");
    const count = input ? Math.max(1, parseInt(input.value, 10) || 1) : 1;
    for (let i = 0; i < count; i++) {
      const nextId = getNextRowIndex();
      drtRows.unshift(generateBlankRow(nextId));
    }
    if (input) input.value = "";
    updateAddRowsButton();
    populateProductFilter();
    renderSpreadsheet();
  }

  function updateAddRowsButton() {
    const input = document.getElementById("addRowsCount");
    const btn = document.getElementById("addRowBtn");
    if (!input || !btn) return;
    const val = parseInt(input.value, 10);
    btn.disabled = !(Number.isInteger(val) && val >= 1);
  }

  buildTableStructure();

  {
    const isLoadDrt = !!document.getElementById("loadDrtId");
    const rowGenerator = isLoadDrt ? generateRow : generateBlankRow;
    for (let i = 1; i <= 50; i++) {
      const row = rowGenerator(i);
      recalculateDerivedRow(row);
      drtRows.push(row);
      if (tbody) tbody.appendChild(renderRow(row, i));
    }
  }
  syncTableLayout();

  if (table) {
    table.addEventListener("keydown", (e) => {
      if (!e.target.matches("input.create-drt-cell-input")) return;
      if (e.key !== "Enter") return;
      e.preventDefault();
      const ref = e.target.dataset.cell;
      if (ref) {
        const next = getCellBelow(ref);
        if (next) focusCell(next);
      }
    });
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

  const productFilter      = document.getElementById("productFilter");
  const skuComboEl         = document.getElementById("createDrtSkuCombo");
  const skuSearchInput     = document.getElementById("productFilterSearch");
  const skuComboList       = document.getElementById("createDrtSkuList");
  const skuClearBtn        = document.getElementById("productFilterClear");

  if (productFilter) {
    productFilter.addEventListener("change", applyProductFilter);
  }

  // All SKU names available (kept in sync by populateProductFilter)
  let _skuAllNames = [];

  const SKU_COMBO_THRESHOLD = 5;

  function getUniqueSkuNamesFromTable() {
    const names = new Set();
    drtRows.forEach((row) => {
      const skuValue = String(row.skuName || "").trim();
      if (skuValue) names.add(skuValue);
    });
    return { names: [...names].sort() };
  }

  function populateProductFilter(skipApply) {
    if (!productFilter) return;
    const currentValue = productFilter.value;
    while (productFilter.options.length > 1) productFilter.options.remove(1);
    const { names } = getUniqueSkuNamesFromTable();
    _skuAllNames = names;
    names.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      productFilter.appendChild(opt);
    });

    // Restore selection
    const validValue = currentValue && names.includes(currentValue) ? currentValue : "";
    productFilter.value = validValue;

    // Switch between plain select and searchable combobox
    const useCombo = names.length > SKU_COMBO_THRESHOLD;
    productFilter.hidden  = useCombo;
    if (skuComboEl) skuComboEl.hidden = !useCombo;

    if (useCombo && skuSearchInput) {
      skuSearchInput.value = validValue;
      if (skuClearBtn) skuClearBtn.hidden = !validValue;
      _renderSkuComboList(skuSearchInput.value);
    }

    if (!skipApply) applyProductFilter();
  }

  // Build the filtered <li> list inside the dropdown
  function _renderSkuComboList(query) {
    if (!skuComboList) return;
    skuComboList.innerHTML = "";
    const q = (query || "").trim().toLowerCase();

    // "All SKUs" entry
    const allLi = document.createElement("li");
    allLi.className = "create-drt-sku-combo__item" + (productFilter.value === "" ? " create-drt-sku-combo__item--active" : "");
    allLi.setAttribute("role", "option");
    allLi.dataset.value = "";
    allLi.textContent = "All SKUs";
    skuComboList.appendChild(allLi);

    const filtered = q ? _skuAllNames.filter((n) => n.toLowerCase().includes(q)) : _skuAllNames;

    filtered.forEach((name) => {
      const li = document.createElement("li");
      li.className = "create-drt-sku-combo__item" + (productFilter.value === name ? " create-drt-sku-combo__item--active" : "");
      li.setAttribute("role", "option");
      li.dataset.value = name;
      if (q) {
        const idx = name.toLowerCase().indexOf(q);
        li.innerHTML = _escHtml(name.slice(0, idx))
          + `<mark class="create-drt-sku-combo__mark">${_escHtml(name.slice(idx, idx + q.length))}</mark>`
          + _escHtml(name.slice(idx + q.length));
      } else {
        li.textContent = name;
      }
      skuComboList.appendChild(li);
    });

    if (filtered.length === 0 && q) {
      const empty = document.createElement("li");
      empty.className = "create-drt-sku-combo__empty";
      empty.textContent = "No matching SKUs";
      skuComboList.appendChild(empty);
    }
  }

  function _openSkuCombo() {
    if (!skuComboList) return;
    _renderSkuComboList(skuSearchInput?.value || "");
    skuComboList.hidden = false;
    if (skuSearchInput) skuSearchInput.setAttribute("aria-expanded", "true");
  }

  function _closeSkuCombo() {
    if (!skuComboList) return;
    skuComboList.hidden = true;
    if (skuSearchInput) skuSearchInput.setAttribute("aria-expanded", "false");
  }

  function _selectSkuComboValue(value) {
    productFilter.value = value;
    if (skuSearchInput) skuSearchInput.value = value;
    if (skuClearBtn) skuClearBtn.hidden = !value;
    _closeSkuCombo();
    applyProductFilter();
  }

  // Combobox events
  skuSearchInput?.addEventListener("focus", _openSkuCombo);

  skuSearchInput?.addEventListener("input", () => {
    const q = skuSearchInput.value;
    if (skuClearBtn) skuClearBtn.hidden = !q;
    _renderSkuComboList(q);
    skuComboList.hidden = false;
    // Auto-apply when text exactly matches a SKU or is empty
    if (!q) {
      productFilter.value = "";
      applyProductFilter();
    } else if (_skuAllNames.includes(q)) {
      productFilter.value = q;
      applyProductFilter();
    }
  });

  skuSearchInput?.addEventListener("keydown", (e) => {
    if (e.key === "Escape") { _closeSkuCombo(); skuSearchInput.blur(); }
    if (e.key === "Enter") {
      const active = skuComboList?.querySelector(".create-drt-sku-combo__item--active");
      if (active) _selectSkuComboValue(active.dataset.value);
      else _closeSkuCombo();
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      const items = [...(skuComboList?.querySelectorAll(".create-drt-sku-combo__item") || [])];
      const idx = items.findIndex((i) => i.classList.contains("create-drt-sku-combo__item--active"));
      items.forEach((i) => i.classList.remove("create-drt-sku-combo__item--active"));
      const next = items[Math.min(idx + 1, items.length - 1)];
      next?.classList.add("create-drt-sku-combo__item--active");
      next?.scrollIntoView({ block: "nearest" });
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      const items = [...(skuComboList?.querySelectorAll(".create-drt-sku-combo__item") || [])];
      const idx = items.findIndex((i) => i.classList.contains("create-drt-sku-combo__item--active"));
      items.forEach((i) => i.classList.remove("create-drt-sku-combo__item--active"));
      const prev = items[Math.max(idx - 1, 0)];
      prev?.classList.add("create-drt-sku-combo__item--active");
      prev?.scrollIntoView({ block: "nearest" });
    }
  });

  skuComboList?.addEventListener("mousedown", (e) => {
    e.preventDefault(); // prevent input blur before click registers
    const item = e.target.closest("[data-value]");
    if (item && !item.classList.contains("create-drt-sku-combo__empty")) {
      _selectSkuComboValue(item.dataset.value);
    }
  });

  skuClearBtn?.addEventListener("mousedown", (e) => {
    e.preventDefault();
    _selectSkuComboValue("");
    skuSearchInput?.focus();
  });

  document.addEventListener("click", (e) => {
    if (!e.target.closest("#createDrtSkuCombo")) _closeSkuCombo();
  });

  if (productFilter) {
    populateProductFilter();
    productFilter.addEventListener("focus", populateProductFilter);
  }

  function applyProductFilter() {
    if (!productFilter) return;
    renderSpreadsheet();
  }

  // ── Hide PG Green toggle switch ───────────────────────────────────────────
  const hidePgGreenToggle = document.getElementById("createDrtHidePgGreenToggle");
  if (hidePgGreenToggle) {
    hidePgGreenToggle.addEventListener("change", () => {
      hidePgGreen = hidePgGreenToggle.checked;
      renderSpreadsheet();
    });
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
    const match = String(ref).match(/^([A-Z]+)(\d+)$/i);
    const hot = getSpreadsheetWorksheet();
    if (match && hot) {
      const col = letterToCol(match[1]) - 1;
      const row = parseInt(match[2], 10) - 1;
      try {
        hot.selectCell(row, col);
      } catch {}
      return;
    }
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
    if (!formulaInput || !activeSpreadsheetCell) return null;
    const { key, y } = activeSpreadsheetCell;
    const column = SPREADSHEET_COLUMNS.find((item) => item.key === key);
    const row = spreadsheetVisibleRows[y];
    if (!column || !row || !canEditSpreadsheetColumn(column)) return null;

    const nextValue = normalizeSpreadsheetEntry(column, formulaInput.value);
    if (!setRowValueFromSpreadsheet(column, row, nextValue.displayValue)) return null;
    setStoredFormula(row, column, nextValue.formula);

    if (column.key === "grp1Current" || column.key.startsWith("grp1Year") || column.key === "grp2Current" || column.key.startsWith("grp2Year")) {
      recalculateDerivedRow(row);
    }

    if (column.key === "skuName") {
      populateProductFilter();
    }

    recalculateSpreadsheetFormulas();
    renderSpreadsheet();
    const visibleX = getVisibleSpreadsheetColumns().findIndex((item) => item.key === column.key);
    return visibleX >= 0 ? getSpreadsheetCellRef(visibleX, y) : null;
  }

  if (formulaInput) {
    formulaInput.addEventListener("change", applyFormulaFromBar);
    formulaInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        const cellRef = applyFormulaFromBar();
        if (cellRef) {
          const next = getCellBelow(cellRef);
          if (next) focusCell(next);
        }
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
    formulaInput.focus();
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
    return ref;
  }

  function clearFormulaSelection() {
    return;
  }

  function syncFormulaBarFromCell(inp) {
    if (!inp?.dataset?.cell) return;
    const match = inp.dataset.cell.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return;
    updateFormulaBarState(letterToCol(match[1]) - 1, parseInt(match[2], 10) - 1);
  }

  // ── Formula-bar active-formula tracking ──────────────────────────────────
  // We track this separately from document.activeElement because HOT may use
  // a capture-phase listener that changes focus BEFORE our bubbling listener fires.
  let _formulaBarInFormula = false; // true while formula bar has a "=" formula

  if (formulaInput) {
    const _syncFormulaBarFlag = () => {
      if (formulaInput.readOnly) { _formulaBarInFormula = false; return; }
      const val = formulaInput.value;
      _formulaBarInFormula = val.startsWith("=");
      updateFormulaSelectionMode(hasUnclosedBracket());
      // Update insert position when the user moves the cursor or types
      if (_formulaBarInFormula && !_fRangeActive) {
        _fRangeInsertPos = formulaInput.selectionStart ?? val.length;
        _fRangeInsertLen = 0;
      }
    };
    formulaInput.addEventListener("input",  _syncFormulaBarFlag);
    formulaInput.addEventListener("keyup",  _syncFormulaBarFlag);
    formulaInput.addEventListener("click",  _syncFormulaBarFlag);
    formulaInput.addEventListener("blur",   () => {
      // Don't clear _formulaBarInFormula on blur — we check it in the
      // capture-phase mousedown below (before focus actually changes).
    });
  }

  // ── Formula-range mousedown (CAPTURE PHASE) ───────────────────────────────
  // Capture phase fires before HOT's own listeners, so we can set _fRangeActive
  // while the formula bar is still the focused element and its value is intact.
  if (spreadsheetHost) {
    spreadsheetHost.addEventListener("mousedown", (e) => {
      // ① HOT cell-editor formula mode — check FIRST, before HOT closes
      //    the editor on this same mousedown.
      const hot = getSpreadsheetWorksheet();
      if (hot) {
        try {
          const editor = hot.getActiveEditor();
          if (editor && typeof editor.isOpened === "function" && editor.isOpened()) {
            const val = String(editor.getValue() || "");
            if (val.startsWith("=")) {
              _fRangeActive    = true;
              _fRangeSource    = "editor";
              _fRangeInsertPos = val.length;
              _fRangeInsertLen = 0;
              // Save the formula so _injectRangeRef can work even after
              // HOT closes the editor.
              _fRangeSavedFormula = val;
              // Remember which cell this formula belongs to
              const sel = hot.getSelected();
              _fRangeTargetCell = sel ? { y: sel[0][0], x: sel[0][1] } : null;
              return;
            }
          }
        } catch {}
      }
      // ② Formula-bar formula mode (user was typing in the formula bar)
      if (_formulaBarInFormula && formulaInput) {
        _fRangeActive   = true;
        _fRangeSource   = "bar";
        // _fRangeInsertPos already set by _syncFormulaBarFlag; keep as-is
        _fRangeInsertLen = 0;
        return;
      }
    }, true /* capture */);
  }

  // On mouseup: exit range mode AFTER HOT fires its final afterSelection.
  // Using setTimeout(0) ensures HOT's bubbling mouseup (which triggers afterSelection)
  // completes before we clear _fRangeActive and refocus the formula bar.
  document.addEventListener("mouseup", () => {
    if (!_fRangeActive) return;
    const src = _fRangeSource;
    window.setTimeout(() => {
      _fRangeActive = false;
      // Prevent the very next afterSelection from overwriting the formula bar.
      _fRangeCooldown = true;
      if ((src === "bar" || src === "editor") && formulaInput) {
        // For editor source: the editor was closed by the click, so the
        // formula now lives in the formula bar.  Focus it so the user can
        // continue typing or press Enter to commit.
        if (src === "editor" && _fRangeTargetCell) {
          // Point the formula bar back at the original cell (not the range cell)
          const tc = _fRangeTargetCell;
          const formulaCellRef = document.querySelector(".create-drt-formula-ref");
          if (formulaCellRef) formulaCellRef.textContent = getSpreadsheetCellRef(tc.x, tc.y);
          // Update activeSpreadsheetCell so Enter commits to the right cell
          const col = getVisibleSpreadsheetColumns()[tc.x];
          if (col) activeSpreadsheetCell = { key: col.key, y: tc.y };
          formulaInput.dataset.activeCell = getSpreadsheetCellRef(tc.x, tc.y);
          formulaInput.readOnly = false;
        }
        formulaInput.focus();
        _formulaBarInFormula = formulaInput.value.startsWith("=");
        const endPos = _fRangeInsertPos + _fRangeInsertLen;
        try { formulaInput.setSelectionRange(endPos, endPos); } catch {}
        updateFormulaSelectionMode(hasUnclosedBracket());
      } else {
        updateFormulaSelectionMode(false);
      }
    }, 0);
  });

  // Selection sync is handled by Handsontable's afterSelection hook above.
  // We still listen to clicks so the formula bar stays in sync when the user
  // clicks outside the focused cell (e.g. scrolls and clicks).
  if (spreadsheetHost) {
    spreadsheetHost.addEventListener("click", (event) => {
      const deleteButton = event.target instanceof Element ? event.target.closest("[data-delete-row-id]") : null;
      if (deleteButton) {
        const rowId = deleteButton.getAttribute("data-delete-row-id");
        if (rowId) {
          drtRows = drtRows.filter((row) => String(row.id) !== rowId);
          populateProductFilter();
          renderSpreadsheet();
        }
        return;
      }
      const chartButton = event.target instanceof Element ? event.target.closest("[data-chart-row]") : null;
      if (chartButton) {
        const rowIndex = Number.parseInt(chartButton.getAttribute("data-chart-row") || "", 10);
        const row = spreadsheetVisibleRows[rowIndex];
        openChartModal(row);
        return;
      }
      if (!_fRangeActive && !_fRangeCooldown) syncFormulaBarFromWorksheetSelection();
    });
  }

  if (undoButton) {
    undoButton.addEventListener("click", () => {
      const plugin = getSpreadsheetUndoRedoPlugin();
      if (!plugin || typeof plugin.isUndoAvailable !== "function" || !plugin.isUndoAvailable()) return;
      if (typeof plugin.undo === "function") plugin.undo();
      else {
        const hot = getSpreadsheetWorksheet();
        if (hot && typeof hot.undo === "function") hot.undo();
      }
      requestAnimationFrame(refreshSpreadsheetHistoryButtons);
    });
  }

  if (redoButton) {
    redoButton.addEventListener("click", () => {
      const plugin = getSpreadsheetUndoRedoPlugin();
      if (!plugin || typeof plugin.isRedoAvailable !== "function" || !plugin.isRedoAvailable()) return;
      if (typeof plugin.redo === "function") plugin.redo();
      else {
        const hot = getSpreadsheetWorksheet();
        if (hot && typeof hot.redo === "function") hot.redo();
      }
      requestAnimationFrame(refreshSpreadsheetHistoryButtons);
    });
  }

  chartModal?.addEventListener("click", (event) => {
    if (!(event.target instanceof Element)) return;

    // Close
    const closeTrigger = event.target.closest("[data-chart-close]");
    if (closeTrigger) { closeChartModal(); return; }

    // Layout switcher
    const layoutBtn = event.target.closest("[data-chart-layout]");
    if (layoutBtn && _chartActiveRow) {
      _chartLayout = Number(layoutBtn.dataset.chartLayout) || 1;
      chartModal.querySelectorAll(".create-drt-chart-layout-btn").forEach((btn) => {
        btn.classList.toggle("active", Number(btn.dataset.chartLayout) === _chartLayout);
      });
      // Show/hide global type switcher
      const typeSwitcher = chartModal.querySelector("#createDrtChartTypeSwitcher");
      if (typeSwitcher) typeSwitcher.style.display = _chartLayout > 1 ? "none" : "";
      _renderAllPanels();
      return;
    }

    // Global type switcher (single-panel mode)
    const typeBtn = event.target.closest("[data-chart-type]");
    if (typeBtn && !typeBtn.closest("[data-panel-type]") && _chartActiveRow) {
      _chartType = typeBtn.dataset.chartType;
      chartModal.querySelectorAll(".create-drt-chart-type-btn").forEach((btn) => {
        btn.classList.toggle("active", btn.dataset.chartType === _chartType);
      });
      _renderAllPanels();
      return;
    }

    // Per-panel type switcher (multi-panel mode)
    const panelTypeBtn = event.target.closest("[data-panel-type]");
    if (panelTypeBtn && _chartActiveRow) {
      const idx  = Number(panelTypeBtn.dataset.panelIdx);
      const type = panelTypeBtn.dataset.panelType;
      _chartPanelTypes[idx] = type;
      // Update tab active state within this panel only
      const panel = panelTypeBtn.closest(".create-drt-chart-panel");
      if (panel) {
        panel.querySelectorAll(".create-drt-chart-panel-tab").forEach((btn) => {
          btn.classList.toggle("active", btn.dataset.panelType === type);
        });
        const svgEl = panel.querySelector("[data-panel-svg]");
        const values = getChartSeries(_chartActiveRow);
        const labels = ["Current", "Y1", "Y2", "Y3", "Y4", "Y5"];
        _renderChartIntoSvg(svgEl, type, values, labels, true);
      }
      return;
    }
  });

  // ── PG pill hover tooltip ────────────────────────────────────────────────
  const pgTooltip = document.createElement("div");
  pgTooltip.id = "createDrtPgTooltip";
  pgTooltip.className = "create-drt-pg-tooltip";
  pgTooltip.setAttribute("aria-hidden", "true");
  pgTooltip.innerHTML = [
    `<div class="create-drt-pg-tooltip-row create-drt-pg-tooltip-row--low">`,
    `<span class="create-drt-pg-tooltip-val">&lt; $100</span>`,
    `</div>`,
    `<div class="create-drt-pg-tooltip-row create-drt-pg-tooltip-row--mid">`,
    `<span class="create-drt-pg-tooltip-val">$100 to $200</span>`,
    `</div>`,
    `<div class="create-drt-pg-tooltip-row create-drt-pg-tooltip-row--high">`,
    `<span class="create-drt-pg-tooltip-val">&gt; $200</span>`,
    `</div>`,
  ].join("");
  document.body.appendChild(pgTooltip);

  let _pgTooltipHideTimer = null;
  function showPgTooltip(indicatorEl) {
    clearTimeout(_pgTooltipHideTimer);
    const tier = indicatorEl.dataset.pgTier;   // "low" | "mid" | "high"
    pgTooltip.querySelectorAll(".create-drt-pg-tooltip-row").forEach((row) => {
      const active = row.classList.contains("create-drt-pg-tooltip-row--" + tier);
      row.classList.toggle("create-drt-pg-tooltip-row--active", active);
    });
    const rect = indicatorEl.getBoundingClientRect();
    pgTooltip.style.left = (rect.left + rect.width / 2) + "px";
    pgTooltip.style.top  = (rect.top - 8) + "px";
    pgTooltip.classList.add("create-drt-pg-tooltip--visible");
  }
  function hidePgTooltip() {
    _pgTooltipHideTimer = setTimeout(() => {
      pgTooltip.classList.remove("create-drt-pg-tooltip--visible");
    }, 120);
  }

  document.addEventListener("mouseover", (e) => {
    let ind = e.target instanceof Element ? e.target.closest(".create-drt-pg-indicator") : null;
    if (!ind && e.target instanceof Element) {
      ind = e.target.closest("td.create-drt-pg-cell")?.querySelector(".create-drt-pg-indicator") || null;
    }
    if (ind) showPgTooltip(ind);
  });
  document.addEventListener("mouseout", (e) => {
    let ind = e.target instanceof Element ? e.target.closest(".create-drt-pg-indicator") : null;
    if (!ind && e.target instanceof Element) {
      ind = e.target.closest("td.create-drt-pg-cell")?.querySelector(".create-drt-pg-indicator") || null;
    }
    if (!ind) return;
    // Only hide when the mouse leaves the indicator entirely — not when moving into a child (e.g. the SVG)
    let toEl = e.relatedTarget instanceof Element ? e.relatedTarget.closest(".create-drt-pg-indicator") : null;
    if (!toEl && e.relatedTarget instanceof Element) {
      toEl = e.relatedTarget.closest("td.create-drt-pg-cell")?.querySelector(".create-drt-pg-indicator") || null;
    }
    if (toEl !== ind) hidePgTooltip();
  });

  // SOX Pre-validate modal
  const soxModal = document.getElementById("soxModal");
  const soxPrevalidateBtn = document.getElementById("soxPrevalidateBtn");
  function openSoxModal() { if (soxModal) soxModal.hidden = false; }
  function closeSoxModal() { if (soxModal) soxModal.hidden = true; }
  soxPrevalidateBtn?.addEventListener("click", openSoxModal);

  // Export to Excel
  const exportToExcelBtn = document.getElementById("exportToExcelBtn");
  exportToExcelBtn?.addEventListener("click", () => {
    _exportDrtToExcel();
  });

  function _exportDrtToExcel() {
    if (typeof XLSX === "undefined") {
      alert("Excel export library not loaded. Please check your internet connection and refresh the page.");
      return;
    }

    // ── Columns to export: all spreadsheet columns except chart-action ──────
    const exportCols = SPREADSHEET_COLUMNS.filter((c) => c.kind !== "chart-action");
    const nCols = exportCols.length;
    const CALC_KEYS = ["calc1", "calc2", "calc3", "calc4", "calc5"];

    // ── Form / deal state ───────────────────────────────────────────────────
    const _fget = (id) => document.getElementById(id)?.value || "";
    const accountName = _fget("createDrtAccountName");
    const country     = _fget("createDrtCountry");
    const currency    = _fget("createDrtCurrency");
    const term        = _fget("createDrtTerm") || "1";
    const priority    = _fget("createDrtPriority") || "";
    const targetDate  = _fget("createDrtTargetDate");
    const assignedTo  = _fget("createDrtAssignedTo") || _fget("createDrtAssignTo");

    // ── Active scenario ─────────────────────────────────────────────────────
    const activeScenario = _scenarios.find((s) => s.id === _activeScenarioId);
    const scenarioName   = activeScenario?.name || "Scenario A";

    // ── Column group extents (for merged header cells) ──────────────────────
    const _ci = (key) => exportCols.findIndex((c) => c.key === key);
    const calcStart  = exportCols.findIndex((c) => c.kind === "calc");
    const calcCount  = exportCols.filter((c) => c.kind === "calc").length;
    const subsStart  = _ci("grp1Current");
    const subsCount  = exportCols.filter((c) => SUBSCRIPTION_COLUMNS.some((s) => s.key === c.key)).length;
    const monthStart = _ci("grp2Current");
    const monthCount = exportCols.filter((c) => MONTHLY_PRICE_COLUMNS.some((s) => s.key === c.key)).length;
    const aovStart   = _ci("grp3Current");
    const aovCount   = exportCols.filter((c) => TOTAL_ANNUAL_COLUMNS.some((s) => s.key === c.key) || c.key === "newTotal").length;
    const prodStart  = _ci("status");
    const prodCount  = exportCols.filter((c) => PRODUCT_GROUP_KEYS.includes(c.key)).length;

    // ── Build array-of-arrays ───────────────────────────────────────────────
    const aoa = [];

    // Row 0 — deal metadata spread across columns
    const metaRow = new Array(nCols).fill("");
    const _setMeta = (col, val) => { if (col < nCols) metaRow[col] = val; };
    _setMeta(0,  `Scenario: ${scenarioName}`);
    _setMeta(2,  `Account: ${accountName}`);
    _setMeta(4,  `Country: ${country}`);
    _setMeta(6,  `Currency: ${currency}`);
    _setMeta(8,  `Term: Y${term}`);
    _setMeta(10, `Priority: ${priority}`);
    _setMeta(12, `Target Date: ${targetDate}`);
    _setMeta(14, `Assigned To: ${assignedTo}`);
    aoa.push(metaRow);

    // Row 1 — spacer
    aoa.push(new Array(nCols).fill(""));

    // Row 2 — nested group headers
    const nestedRow = new Array(nCols).fill("");
    if (calcStart  >= 0)               nestedRow[calcStart]  = "Workpad";
    if (subsStart  >= 0)               nestedRow[subsStart]  = "Number of Subscriptions";
    if (monthStart >= 0)               nestedRow[monthStart] = "Monthly Unit Price";
    if (aovStart   >= 0)               nestedRow[aovStart]   = "Total AOV";
    if (prodStart  >= 0 && prodCount > 0) nestedRow[prodStart] = "Product Info";
    aoa.push(nestedRow);

    // Row 3 — individual column labels
    aoa.push(exportCols.map((c) => c.label || ""));

    // Rows 4+ — data
    drtRows.forEach((row) => {
      if (!row) return;

      // Group-header sentinel rows: output as a single labelled separator
      if (row._isGroupHeader) {
        const groupRow = new Array(nCols).fill("");
        const labelCol = calcStart >= 0 ? calcStart : 0;
        groupRow[labelCol] = `── ${row.label || row.skuName || "Group"} ──`;
        aoa.push(groupRow);
        return;
      }

      aoa.push(exportCols.map((col) => {
        if (col.kind === "calc") {
          const idx = CALC_KEYS.indexOf(col.key);
          return idx >= 0 ? (row.calcValues?.[idx] ?? "") : "";
        }
        const raw = getColumnValue(col, row) ?? "";
        // Return numbers as numbers so Excel can format them
        if (raw !== "" && !isNaN(Number(raw))) return Number(raw);
        return raw;
      }));
    });

    // ── Build worksheet ─────────────────────────────────────────────────────
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Column widths  (px ÷ 7 ≈ Excel character units)
    ws["!cols"] = exportCols.map((c) => ({
      wch: Math.max(8, Math.round((c.width || c.minWidth || 100) / 7)),
    }));

    // Cell merges
    const NR = 2; // nested-header row index (0-based)
    const merges = [];
    const _merge = (r, cStart, span) => {
      if (cStart >= 0 && span > 1) merges.push({ s: { r, c: cStart }, e: { r, c: cStart + span - 1 } });
    };
    _merge(0,  0,         4);           // metadata label area
    _merge(NR, calcStart,  calcCount);
    _merge(NR, subsStart,  subsCount);
    _merge(NR, monthStart, monthCount);
    _merge(NR, aovStart,   aovCount);
    _merge(NR, prodStart,  prodCount);
    ws["!merges"] = merges;

    // Freeze panes: lock the 4 header rows so data scrolls underneath
    ws["!freeze"] = { xSplit: 0, ySplit: 4 };

    // ── Workbook ─────────────────────────────────────────────────────────────
    const wb = XLSX.utils.book_new();

    // Main data sheet named after the active scenario
    const safeSheet = (scenarioName || "DRT").replace(/[\\/*?[\]:]/g, "").slice(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, safeSheet);

    // "All Scenarios" summary sheet (only when >1 scenario exists)
    if (_scenarios.length > 1) {
      const infoAoa = [
        ["Scenario", "Account", "Country", "Currency", "Term", "Priority", "Target Date", "Assigned To"],
        ..._scenarios.map((sc) => {
          const fs = sc.formState || {};
          return [
            sc.name,
            fs.accountName  || "",
            fs.country      || "",
            fs.currency     || "",
            fs.term         || "",
            fs.priority     || "",
            fs.targetDate   || "",
            fs.assignedTo   || fs.assignTo || "",
          ];
        }),
      ];
      const infoWs = XLSX.utils.aoa_to_sheet(infoAoa);
      infoWs["!cols"] = [28, 28, 18, 14, 10, 14, 18, 24].map((wch) => ({ wch }));
      XLSX.utils.book_append_sheet(wb, infoWs, "All Scenarios");
    }

    // ── Trigger download ─────────────────────────────────────────────────────
    const safeDeal     = (accountName || "DRT-Export").replace(/[^a-zA-Z0-9_\- ]/g, "").trim() || "DRT-Export";
    const safeScenario = scenarioName.replace(/[^a-zA-Z0-9_\- ]/g, "").trim();
    const dateStr      = new Date().toISOString().slice(0, 10);
    const filename     = `${safeDeal}_${safeScenario}_${dateStr}.xlsx`;

    XLSX.writeFile(wb, filename);
  }

  soxModal?.addEventListener("click", (event) => {
    const closeTrigger = event.target instanceof Element ? event.target.closest("[data-sox-close]") : null;
    if (closeTrigger) closeSoxModal();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && chartModal && !chartModal.hidden) {
      if (_chartMaximized) _setChartMaximized(false); // restore first, then close on next Esc
      else closeChartModal();
    }
    if (event.key === "Escape" && soxModal && !soxModal.hidden) {
      closeSoxModal();
    }
  });

  table?.classList.add("input-style-modern");
  setSpreadsheetOnlyMode();

  window.addEventListener("resize", syncTableLayout);

  const themeToggle = document.getElementById("createDrtThemeToggle");

  const headerToggle = document.getElementById("createDrtHeaderToggle");
  const savedTheme = window.localStorage.getItem("drt-theme");
  const savedExcelCollapsed = window.localStorage.getItem("create-drt-excel-collapsed") === "true";

  function applyTheme(theme) {
    document.body.dataset.theme = theme;
    const isNight = theme === "night";

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
  renderOwnerChips();

  function applyExcelCollapsed(collapsed) {
    excelCanvasCollapsed = !!collapsed;
    if (excelOutline) {
      // Label+button sit above the row-number column (offset=0); line extends across calc columns.
      const collapsedWidth = EXCEL_OUTLINE_TOGGLE_SIZE + 90;
      excelOutline.style.setProperty("--excel-outline-offset", "0px");
      excelOutline.style.setProperty("--excel-outline-width", `${collapsed ? collapsedWidth : SPREADSHEET_ROW_HEADER_WIDTH + EXCEL_CANVAS_TOTAL_WIDTH}px`);
    }
    if (excelToggle) {
      excelToggle.textContent = collapsed ? "+" : "-";
      excelToggle.title = collapsed ? "Show Workpad columns" : "Hide Workpad columns";
      excelToggle.setAttribute("aria-label", collapsed ? "Show Workpad columns" : "Hide Workpad columns");
      excelToggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
    }
    // Calc columns toggling shifts the pixel position of Additional Info columns — recompute and reapply
    currentAddlInfoMetrics = computeAddlInfoOutlineMetrics();
    if (addlInfoOutline) {
      const addlWidth = additionalInfoCollapsed ? currentAddlInfoMetrics.collapsedWidth : currentAddlInfoMetrics.expandedWidth;
      addlInfoOutline.style.setProperty("--excel-outline-offset", `${currentAddlInfoMetrics.left}px`);
      addlInfoOutline.style.setProperty("--excel-outline-width", `${addlWidth}px`);
    }
    // Calc-column toggle shifts pricing info columns too — recompute
    currentPricingInfoMetrics = computePricingInfoOutlineMetrics();
    if (pricingInfoOutline) {
      const pricingWidth = pricingInfoCollapsed ? currentPricingInfoMetrics.collapsedWidth : currentPricingInfoMetrics.expandedWidth;
      pricingInfoOutline.style.setProperty("--excel-outline-offset", `${currentPricingInfoMetrics.left}px`);
      pricingInfoOutline.style.setProperty("--excel-outline-width", `${pricingWidth}px`);
    }
    // Also shifts product info columns
    currentProductInfoGroupMetrics = computeProductInfoGroupOutlineMetrics();
    if (productInfoOutline) {
      const piWidth = productInfoGroupCollapsed ? currentProductInfoGroupMetrics.collapsedWidth : currentProductInfoGroupMetrics.expandedWidth;
      productInfoOutline.style.setProperty("--excel-outline-offset", `${currentProductInfoGroupMetrics.left}px`);
      productInfoOutline.style.setProperty("--excel-outline-width", `${piWidth}px`);
    }
    window.localStorage.setItem("create-drt-excel-collapsed", String(collapsed));
    if (spreadsheetInstance) renderSpreadsheet();
  }

  if (excelToggle) {
    excelToggle.addEventListener("click", () => {
      applyExcelCollapsed(!excelCanvasCollapsed);
    });
  }

  function applyAddlInfoCollapsed(collapsed) {
    additionalInfoCollapsed = !!collapsed;
    currentAddlInfoMetrics = computeAddlInfoOutlineMetrics();
    if (addlInfoOutline) {
      addlInfoOutline.style.setProperty("--excel-outline-offset", `${currentAddlInfoMetrics.left}px`);
      addlInfoOutline.style.setProperty("--excel-outline-width", `${collapsed ? currentAddlInfoMetrics.collapsedWidth : currentAddlInfoMetrics.expandedWidth}px`);
    }
    if (addlInfoToggleEl) {
      addlInfoToggleEl.textContent = collapsed ? "+" : "-";
      addlInfoToggleEl.title = collapsed ? "Show Additional Inputs columns" : "Hide Additional Inputs columns";
      addlInfoToggleEl.setAttribute("aria-label", collapsed ? "Show Additional Inputs columns" : "Hide Additional Inputs columns");
      addlInfoToggleEl.setAttribute("aria-expanded", collapsed ? "false" : "true");
    }
    // Collapsing addl-info shifts pricing columns — recompute pricing position
    currentPricingInfoMetrics = computePricingInfoOutlineMetrics();
    if (pricingInfoOutline) {
      const pricingWidth = pricingInfoCollapsed ? currentPricingInfoMetrics.collapsedWidth : currentPricingInfoMetrics.expandedWidth;
      pricingInfoOutline.style.setProperty("--excel-outline-offset", `${currentPricingInfoMetrics.left}px`);
      pricingInfoOutline.style.setProperty("--excel-outline-width", `${pricingWidth}px`);
    }
    // Also shifts product info columns
    currentProductInfoGroupMetrics = computeProductInfoGroupOutlineMetrics();
    if (productInfoOutline) {
      const piWidth = productInfoGroupCollapsed ? currentProductInfoGroupMetrics.collapsedWidth : currentProductInfoGroupMetrics.expandedWidth;
      productInfoOutline.style.setProperty("--excel-outline-offset", `${currentProductInfoGroupMetrics.left}px`);
      productInfoOutline.style.setProperty("--excel-outline-width", `${piWidth}px`);
    }
    window.localStorage.setItem("create-drt-addl-info-collapsed", String(collapsed));
    if (spreadsheetInstance) renderSpreadsheet();
  }

  if (addlInfoToggleEl) {
    addlInfoToggleEl.addEventListener("click", () => {
      applyAddlInfoCollapsed(!additionalInfoCollapsed);
    });
  }

  function applyPricingInfoCollapsed(collapsed) {
    pricingInfoCollapsed = !!collapsed;
    currentPricingInfoMetrics = computePricingInfoOutlineMetrics();
    if (pricingInfoOutline) {
      pricingInfoOutline.style.setProperty("--excel-outline-offset", `${currentPricingInfoMetrics.left}px`);
      pricingInfoOutline.style.setProperty("--excel-outline-width", `${collapsed ? currentPricingInfoMetrics.collapsedWidth : currentPricingInfoMetrics.expandedWidth}px`);
    }
    if (pricingInfoToggleEl) {
      pricingInfoToggleEl.textContent = collapsed ? "+" : "-";
      pricingInfoToggleEl.title = collapsed ? "Show Pricing columns" : "Hide Pricing columns";
      pricingInfoToggleEl.setAttribute("aria-label", collapsed ? "Show Pricing columns" : "Hide Pricing columns");
      pricingInfoToggleEl.setAttribute("aria-expanded", collapsed ? "false" : "true");
    }
    // Pricing toggle shifts product info outline position
    currentProductInfoGroupMetrics = computeProductInfoGroupOutlineMetrics();
    if (productInfoOutline) {
      const piWidth = productInfoGroupCollapsed ? currentProductInfoGroupMetrics.collapsedWidth : currentProductInfoGroupMetrics.expandedWidth;
      productInfoOutline.style.setProperty("--excel-outline-offset", `${currentProductInfoGroupMetrics.left}px`);
      productInfoOutline.style.setProperty("--excel-outline-width", `${piWidth}px`);
    }
    window.localStorage.setItem("create-drt-pricing-info-collapsed", String(collapsed));
    if (spreadsheetInstance) renderSpreadsheet();
  }

  if (pricingInfoToggleEl) {
    pricingInfoToggleEl.addEventListener("click", () => {
      applyPricingInfoCollapsed(!pricingInfoCollapsed);
    });
  }

  function applyProductInfoGroupCollapsed(collapsed) {
    productInfoGroupCollapsed = !!collapsed;
    currentProductInfoGroupMetrics = computeProductInfoGroupOutlineMetrics();
    if (productInfoOutline) {
      productInfoOutline.style.setProperty("--excel-outline-offset", `${currentProductInfoGroupMetrics.left}px`);
      productInfoOutline.style.setProperty("--excel-outline-width",
        `${collapsed ? currentProductInfoGroupMetrics.collapsedWidth : currentProductInfoGroupMetrics.expandedWidth}px`);
    }
    if (productInfoGroupToggleEl) {
      productInfoGroupToggleEl.textContent = collapsed ? "+" : "-";
      productInfoGroupToggleEl.title = collapsed ? "Show Product Info columns" : "Hide Product Info columns";
      productInfoGroupToggleEl.setAttribute("aria-label", collapsed ? "Show Product Info columns" : "Hide Product Info columns");
      productInfoGroupToggleEl.setAttribute("aria-expanded", collapsed ? "false" : "true");
    }
    window.localStorage.setItem("create-drt-product-info-collapsed", String(collapsed));
    if (spreadsheetInstance) renderSpreadsheet();
  }

  if (productInfoGroupToggleEl) {
    productInfoGroupToggleEl.addEventListener("click", () => {
      applyProductInfoGroupCollapsed(!productInfoGroupCollapsed);
    });
  }

  const savedAddlInfoCollapsed = window.localStorage.getItem("create-drt-addl-info-collapsed") === "true";
  const savedPricingInfoCollapsed = window.localStorage.getItem("create-drt-pricing-info-collapsed") === "true";
  const savedProductInfoGroupCollapsed = window.localStorage.getItem("create-drt-product-info-collapsed") === "true";
  applyExcelCollapsed(savedExcelCollapsed);
  applyAddlInfoCollapsed(savedAddlInfoCollapsed);
  applyPricingInfoCollapsed(savedPricingInfoCollapsed);
  applyProductInfoGroupCollapsed(savedProductInfoGroupCollapsed);

  function applyHeaderCollapsed(collapsed) {
    document.body.classList.toggle("create-drt-header-collapsed", collapsed);
    if (headerToggle) {
      headerToggle.textContent = collapsed ? "Show Details" : "Hide Details";
      headerToggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
    }
    window.localStorage.setItem("create-drt-header-collapsed", String(collapsed));
  }

  if (headerToggle) {
    headerToggle.addEventListener("click", () => {
      applyHeaderCollapsed(!document.body.classList.contains("create-drt-header-collapsed"));
    });
  }

  applyHeaderCollapsed(false);

  const headerCollapseActionButtons = document.querySelectorAll(".create-drt-save-btn, .create-drt-draft-btn");
  headerCollapseActionButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      applyHeaderCollapsed(true);
    });
  });

  const drtIdEl = document.getElementById("createDrtId");
  if (drtIdEl) {
    const isLoadDrtPage = !!document.getElementById("loadDrtId");
    const urlId = new URLSearchParams(window.location.search).get("id");
    drtIdEl.textContent = (isLoadDrtPage && urlId) ? urlId : "—";
  }

  function updateCollapsedMetricsTerm() {
    const termSelect = document.getElementById("createDrtTerm");
    const termLabels = document.querySelectorAll(".create-drt-collapsed-metric-term");
    const term = termSelect ? termSelect.value || "1" : "1";
    const label = "Year " + term;
    termLabels.forEach((el) => { el.textContent = label; });
  }
  const termSelect = document.getElementById("createDrtTerm");
  if (termSelect) {
    termSelect.addEventListener("change", updateCollapsedMetricsTerm);
    updateCollapsedMetricsTerm();
  }

  const dealInfoTrackedFieldIds = [
    "createDrtAccountName",
    "createDrtCountry",
    "createDrtCurrency",
    "createDrtTerm",
    "createDrtPriority",
    "createDrtTargetDate",
    "createDrtAssignTo",
    "loadDrtResult",
  ];
  const dealInfoRequiredFieldIds = [
    "createDrtAccountName",
    "createDrtCountry",
    "createDrtCurrency",
    "createDrtTerm",
    "createDrtPriority",
  ];
  const dealInfoSaveButtons = Array.from(document.querySelectorAll(".create-drt-save-btn"));
  const dealInfoDraftButtons = Array.from(document.querySelectorAll(".create-drt-draft-btn"));
  let dealInfoInitialSnapshot = null;

  function normalizeDealInfoValue(value) {
    return String(value ?? "").trim();
  }

  function getDealInfoSnapshot() {
    const snapshot = {};
    dealInfoTrackedFieldIds.forEach((id) => {
      const el = document.getElementById(id);
      snapshot[id] = normalizeDealInfoValue(el?.value);
    });
    return snapshot;
  }

  function isDealInfoValid(snapshot) {
    return dealInfoRequiredFieldIds.every((id) => !!normalizeDealInfoValue(snapshot[id]));
  }

  function isDealInfoDirty(snapshot) {
    if (!dealInfoInitialSnapshot) return false;
    return dealInfoTrackedFieldIds.some((id) => normalizeDealInfoValue(snapshot[id]) !== normalizeDealInfoValue(dealInfoInitialSnapshot[id]));
  }

  function refreshDealInfoSaveState() {
    const snapshot = getDealInfoSnapshot();
    const isDirty = isDealInfoDirty(snapshot);
    const shouldEnable = isDealInfoValid(snapshot) && isDirty;
    dealInfoSaveButtons.forEach((btn) => {
      btn.disabled = !shouldEnable;
    });
    dealInfoDraftButtons.forEach((btn) => {
      btn.disabled = !isDirty;
    });
    return shouldEnable;
  }

  function setDealInfoBaseline() {
    dealInfoInitialSnapshot = getDealInfoSnapshot();
    refreshDealInfoSaveState();
  }

  window.refreshDealInfoSaveState = refreshDealInfoSaveState;
  window.setDealInfoBaseline = setDealInfoBaseline;

  dealInfoTrackedFieldIds.forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("input", refreshDealInfoSaveState);
    el.addEventListener("change", refreshDealInfoSaveState);
  });

  requestAnimationFrame(() => {
    setDealInfoBaseline();
  });

  // ── Priority Pill Picker ──
  const priorityPicker = document.getElementById("createDrtPriorityPicker");
  if (priorityPicker) {
    const priorityHidden   = document.getElementById("createDrtPriority");
    const priorityDisplay  = document.getElementById("createDrtPriorityDisplay");
    const priorityOptions  = priorityPicker.querySelectorAll(".create-drt-priority-option");

    const PRIORITY_META = {
      high:   { label: "High",   cls: "create-drt-priority-pill--high"   },
      medium: { label: "Medium", cls: "create-drt-priority-pill--medium" },
      low:    { label: "Low",    cls: "create-drt-priority-pill--low"    },
    };

    function setPriority(value) {
      const meta = PRIORITY_META[value] || PRIORITY_META.low;
      // Update hidden input
      priorityHidden.value = value;
      // Update displayed pill
      priorityDisplay.className = `create-drt-priority-pill ${meta.cls}`;
      priorityDisplay.querySelector(".create-drt-priority-text").textContent = meta.label;
      // Update aria-selected on options
      priorityOptions.forEach((opt) => {
        const selected = opt.dataset.value === value;
        opt.setAttribute("aria-selected", String(selected));
      });
      refreshDealInfoSaveState();
    }

    function openPicker() {
      priorityPicker.setAttribute("aria-expanded", "true");
    }
    function closePicker() {
      priorityPicker.setAttribute("aria-expanded", "false");
    }
    function togglePicker() {
      if (priorityPicker.getAttribute("aria-expanded") === "true") closePicker();
      else openPicker();
    }

    priorityPicker.addEventListener("click", (e) => {
      const opt = e.target.closest(".create-drt-priority-option");
      if (opt) {
        setPriority(opt.dataset.value);
        closePicker();
        return;
      }
      togglePicker();
    });

    priorityPicker.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        togglePicker();
      } else if (e.key === "Escape") {
        closePicker();
      } else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        const vals = ["low", "medium", "high"];
        const cur  = vals.indexOf(priorityHidden.value);
        const next = e.key === "ArrowDown"
          ? Math.min(cur + 1, vals.length - 1)
          : Math.max(cur - 1, 0);
        setPriority(vals[next]);
      }
    });

    // Close when clicking outside
    document.addEventListener("click", (e) => {
      if (!priorityPicker.contains(e.target)) closePicker();
    });

    // Expose setPriority so the scenario manager can restore priority state
    _setPriorityFn = setPriority;
    // Initialise with default (low)
    setPriority(priorityHidden.value || "low");
  }

  // ── Scenario Bar — event delegation ──────────────────────────────────────
  const scenarioBarEl     = document.getElementById("createDrtScenarioBar");
  const scenarioAddBtn    = document.getElementById("createDrtScenarioAddBtn");
  const scenarioToggleBtn = document.getElementById("createDrtScenarioToggleBtn");
  const scenarioRestoreBtn = document.getElementById("createDrtScenarioRestoreBtn");

  // Hide / show the entire scenario bar
  let _scenarioBarCollapsed = window.localStorage.getItem("drt-scenario-bar-collapsed") === "true";

  function _applyScenarioBarCollapsed(collapsed) {
    _scenarioBarCollapsed = collapsed;
    if (scenarioBarEl) {
      scenarioBarEl.classList.toggle("create-drt-scenario-bar--collapsed", collapsed);
    }
    if (scenarioToggleBtn) {
      scenarioToggleBtn.setAttribute("aria-expanded", collapsed ? "false" : "true");
      scenarioToggleBtn.title = collapsed ? "Show scenarios bar" : "Hide scenarios bar";
    }
    if (scenarioRestoreBtn) {
      scenarioRestoreBtn.hidden = !collapsed;
      // Show the active scenario name instead of the generic "Scenarios" label
      const nameEl = scenarioRestoreBtn.querySelector("span");
      if (nameEl) {
        const active = _scenarios.find((s) => s.id === _activeScenarioId);
        const full = active?.name || "Scenarios";
        nameEl.textContent = full.length > 8 ? full.slice(0, 8) + ".." : full;
        nameEl.title = full.length > 8 ? full : "";
      }
    }
    window.localStorage.setItem("drt-scenario-bar-collapsed", String(collapsed));
  }

  scenarioToggleBtn?.addEventListener("click", () => {
    _applyScenarioBarCollapsed(!_scenarioBarCollapsed);
  });

  scenarioRestoreBtn?.addEventListener("click", () => {
    _applyScenarioBarCollapsed(false);
  });

  if (scenarioBarEl) {
    scenarioBarEl.addEventListener("click", (e) => {
      // Ignore clicks on the toggle button (handled above)
      if (e.target instanceof Element && e.target.closest("#createDrtScenarioToggleBtn")) return;

      const deleteBtn = e.target instanceof Element ? e.target.closest("[data-scenario-delete]") : null;
      const renameBtn = e.target instanceof Element ? e.target.closest("[data-scenario-rename]") : null;
      const tab       = e.target instanceof Element ? e.target.closest(".create-drt-scenario-tab") : null;

      if (deleteBtn) { _deleteScenario(deleteBtn.dataset.scenarioDelete); return; }
      if (renameBtn) { _startScenarioRename(renameBtn.dataset.scenarioRename); return; }
      if (tab) {
        const id = tab.dataset.scenarioId;
        if (id && id !== _activeScenarioId) _activateScenario(id);
      }
    });

    // Double-click on tab name to rename
    scenarioBarEl.addEventListener("dblclick", (e) => {
      const nameEl = e.target instanceof Element ? e.target.closest("[data-scenario-name]") : null;
      if (nameEl) _startScenarioRename(nameEl.dataset.scenarioName);
    });
  }

  scenarioAddBtn?.addEventListener("click", () => { _addScenario(); });

  // ── Initialize scenarios — one "Base Case" from current state ─────────────
  {
    const firstId = _makeScenarioId();
    _scenarios = [{
      id:        firstId,
      name:      "Base Case",
      color:     SCENARIO_COLORS[0],
      rows:      _deepCloneRows(drtRows),
      formState: _captureFormState(),
    }];
    _activeScenarioId = firstId;
    _renderScenarioBar();
    _applyScenarioBarCollapsed(_scenarioBarCollapsed);
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
