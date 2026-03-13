const state = {
  workbooks: [],
  selectedMainWorkbook: null,
  selectedReferenceWorkbook: null,
  mainSheetTabs: [],
  referenceSheetTabs: [],
  mainAvailableMonths: new Map(),
  cache: new Map(),
  selectedRefCanonical: null,
  selectedMainSheetName: null,
  selectedReferenceSheetName: null,
  pairVersion: null,
  mainStyledRequestKey: null,
  pollHandle: null,
  isRefreshing: false,
};

const POLL_INTERVAL_MS = 5000;

const el = {
  uploadInput: document.getElementById("uploadInput"),
  uploadBtn: document.getElementById("uploadBtn"),
  mainWorkbookSelect: document.getElementById("mainWorkbookSelect"),
  referenceWorkbookSelect: document.getElementById("referenceWorkbookSelect"),
  mainWorkbookName: document.getElementById("mainWorkbookName"),
  referenceWorkbookName: document.getElementById("referenceWorkbookName"),
  mainPanelTitle: document.getElementById("mainPanelTitle"),
  referencePanelTitle: document.getElementById("referencePanelTitle"),
  mainTabSelect: document.getElementById("mainTabSelect"),
  mainSheetTabs: document.getElementById("mainSheetTabs"),
  referenceTabSelect: document.getElementById("referenceTabSelect"),
  referenceSheetTabs: document.getElementById("referenceSheetTabs"),
  mainModeSelect: document.getElementById("mainModeSelect"),
  mainNInput: document.getElementById("mainNInput"),
  mainMonthWrapper: document.getElementById("mainMonthWrapper"),
  mainMonthSelect: document.getElementById("mainMonthSelect"),
  refModeSelect: document.getElementById("refModeSelect"),
  refNInput: document.getElementById("refNInput"),
  refMonthWrapper: document.getElementById("refMonthWrapper"),
  refMonthSelect: document.getElementById("refMonthSelect"),
  metricSelect: document.getElementById("metricSelect"),
  searchInput: document.getElementById("searchInput"),
  statusText: document.getElementById("statusText"),
  mainTable: document.getElementById("mainTable"),
  refTable: document.getElementById("refTable"),
  mainMeta: document.getElementById("mainMeta"),
  refMeta: document.getElementById("refMeta"),
};

const monthLabels = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

async function fetchJson(url) {
  const res = await fetch(url);
  if (!res.ok) {
    let message = `Request failed (${res.status}) for ${url}`;
    try {
      const payload = await res.json();
      if (payload && payload.error) {
        message = payload.error;
      }
    } catch {
      // Keep fallback message.
    }
    throw new Error(message);
  }
  return res.json();
}

function workbookQuery() {
  const params = new URLSearchParams();
  if (state.selectedMainWorkbook) {
    params.set("main", state.selectedMainWorkbook);
  }
  if (state.selectedReferenceWorkbook) {
    params.set("reference", state.selectedReferenceWorkbook);
  }
  return params.toString();
}

function cacheKey(canonical) {
  const version = state.pairVersion || "noversion";
  return `${state.selectedMainWorkbook}::${state.selectedReferenceWorkbook}::${version}::${canonical}`;
}

function populateWorkbookSelect(selectElement, selectedName) {
  selectElement.innerHTML = "";
  for (const workbook of state.workbooks) {
    const option = document.createElement("option");
    option.value = workbook;
    option.textContent = workbook;
    selectElement.appendChild(option);
  }
  if (selectedName) {
    selectElement.value = selectedName;
  }
}

function updateWorkbookLabels() {
  el.mainWorkbookName.textContent = state.selectedMainWorkbook || "-";
  el.referenceWorkbookName.textContent = state.selectedReferenceWorkbook || "-";
  el.mainPanelTitle.textContent = `Main View (${state.selectedMainWorkbook || "-"})`;
  el.referencePanelTitle.textContent = `Reference Detail (${state.selectedReferenceWorkbook || "-"})`;
}

function setEmptyMain(message) {
  el.mainTable.innerHTML = `<div class="empty">${message}</div>`;
  el.mainMeta.textContent = "";
}

function setEmptyRef(message) {
  unwrapSplitViewport(el.refTable);
  el.refTable.innerHTML = `<tbody><tr><td class="empty">${message}</td></tr></tbody>`;
  el.refMeta.textContent = "";
}

function getMainTabMeta(sheetName = state.selectedMainSheetName) {
  return state.mainSheetTabs.find((tab) => tab.sheet_name === sheetName) || null;
}

function getReferenceTabMeta(sheetName = state.selectedReferenceSheetName) {
  return state.referenceSheetTabs.find((tab) => tab.sheet_name === sheetName) || null;
}

function getPayloadForCanonical(canonical) {
  if (!canonical) {
    return null;
  }
  return state.cache.get(cacheKey(canonical)) || null;
}

function tableWrapFor(table) {
  return table ? table.closest(".table-wrap") : null;
}

function clearStickyStyles(table) {
  if (!table) {
    return;
  }
  for (const cell of table.querySelectorAll("th, td")) {
    cell.classList.remove("sticky-col", "sticky-col-boundary", "sticky-col-head");
    cell.style.removeProperty("left");
    cell.style.removeProperty("z-index");
    cell.style.removeProperty("background-color");
  }
}

function unwrapSplitViewport(table) {
  if (!table) {
    return table;
  }

  const rightPane = table.parentElement;
  if (!rightPane || !rightPane.classList.contains("split-pane-right")) {
    return table;
  }

  const splitView = rightPane.parentElement;
  if (!splitView || !splitView.classList.contains("split-view")) {
    return table;
  }

  const host = splitView.parentElement;
  if (!host) {
    return table;
  }

  host.insertBefore(table, splitView);
  splitView.remove();

  const wrap = tableWrapFor(table);
  if (wrap) {
    wrap.classList.remove("table-wrap-split");
  }
  return table;
}

async function onMainTabChange(sheetName) {
  state.selectedMainSheetName = sheetName;
  state.mainStyledRequestKey = null;

  const tabMeta = getMainTabMeta(sheetName);
  if (tabMeta && tabMeta.canonical) {
    await ensureSheetLoaded(tabMeta.canonical);
  }

  rebuildMainTabs(state.selectedMainSheetName);
  await render();
}

function rebuildMainTabs(preferredSheetName) {
  if (!state.mainSheetTabs.length) {
    state.selectedMainSheetName = null;
    el.mainTabSelect.innerHTML = "";
    el.mainSheetTabs.innerHTML = "";
    setEmptyMain("No sheets found in the selected overview workbook.");
    return;
  }

  const validSheet = state.mainSheetTabs.some((tab) => tab.sheet_name === preferredSheetName)
    ? preferredSheetName
    : state.mainSheetTabs[0].sheet_name;

  state.selectedMainSheetName = validSheet;
  el.mainTabSelect.innerHTML = "";
  for (const tab of state.mainSheetTabs) {
    const option = document.createElement("option");
    option.value = tab.sheet_name;
    option.textContent = tab.sheet_name;
    el.mainTabSelect.appendChild(option);
  }
  el.mainTabSelect.value = validSheet;
  el.mainSheetTabs.innerHTML = "";

  for (const tab of state.mainSheetTabs) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `sheet-tab-btn${tab.sheet_name === validSheet ? " active" : ""}`;
    button.textContent = tab.sheet_name;
    if (!tab.filterable) {
      button.title = "Rendered as full sheet (non-month layout)";
    }
    button.addEventListener("click", () => {
      onMainTabChange(tab.sheet_name).catch((err) => {
        el.statusText.textContent = err.message;
      });
    });
    el.mainSheetTabs.appendChild(button);
  }

  const activeBtn = el.mainSheetTabs.querySelector(".sheet-tab-btn.active");
  if (activeBtn) {
    activeBtn.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
}

async function onReferenceTabChange(sheetName) {
  state.selectedReferenceSheetName = sheetName;
  const tabMeta = getReferenceTabMeta(sheetName);
  state.selectedRefCanonical = tabMeta ? tabMeta.canonical : null;
  if (state.selectedRefCanonical) {
    await ensureSheetLoaded(state.selectedRefCanonical);
  }

  rebuildReferenceTabs(state.selectedReferenceSheetName);
  await render();
}

function rebuildReferenceTabs(preferredSheetName) {
  if (!state.referenceSheetTabs.length) {
    state.selectedReferenceSheetName = null;
    state.selectedRefCanonical = null;
    el.referenceTabSelect.innerHTML = "";
    el.referenceSheetTabs.innerHTML = "";
    setEmptyRef("No sheets found in the selected detail workbook.");
    return;
  }

  const validSheet = state.referenceSheetTabs.some((tab) => tab.sheet_name === preferredSheetName)
    ? preferredSheetName
    : state.referenceSheetTabs[0].sheet_name;

  state.selectedReferenceSheetName = validSheet;
  const selectedTab = getReferenceTabMeta(validSheet);
  state.selectedRefCanonical = selectedTab ? selectedTab.canonical : null;

  el.referenceTabSelect.innerHTML = "";
  for (const tab of state.referenceSheetTabs) {
    const option = document.createElement("option");
    option.value = tab.sheet_name;
    option.textContent = tab.sheet_name;
    el.referenceTabSelect.appendChild(option);
  }
  el.referenceTabSelect.value = validSheet;

  el.referenceSheetTabs.innerHTML = "";
  for (const tab of state.referenceSheetTabs) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `sheet-tab-btn${tab.sheet_name === validSheet ? " active" : ""}`;
    button.textContent = tab.sheet_name;
    if (!tab.filterable) {
      button.title = "Detail parser not available for this sheet.";
    }
    button.addEventListener("click", () => {
      onReferenceTabChange(tab.sheet_name).catch((err) => {
        el.statusText.textContent = err.message;
      });
    });
    el.referenceSheetTabs.appendChild(button);
  }

  const activeBtn = el.referenceSheetTabs.querySelector(".sheet-tab-btn.active");
  if (activeBtn) {
    activeBtn.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
}

async function loadWorkbookOptions() {
  const payload = await fetchJson("/api/workbooks");
  state.workbooks = payload.workbooks || [];

  if (!state.workbooks.length) {
    throw new Error("No supported Excel files found in this folder.");
  }

  state.selectedMainWorkbook = payload.default_main || state.workbooks[0];
  state.selectedReferenceWorkbook = payload.default_reference || state.workbooks[0];

  populateWorkbookSelect(el.mainWorkbookSelect, state.selectedMainWorkbook);
  populateWorkbookSelect(el.referenceWorkbookSelect, state.selectedReferenceWorkbook);
  updateWorkbookLabels();
}

async function uploadSelectedWorkbooks() {
  const files = Array.from(el.uploadInput.files || []);
  if (!files.length) {
    el.statusText.textContent = "Choose one or more Excel files first.";
    return;
  }

  const formData = new FormData();
  for (const file of files) {
    formData.append("files", file);
  }

  el.statusText.textContent = `Uploading ${files.length} file(s)...`;
  const res = await fetch("/api/upload-workbooks", { method: "POST", body: formData });
  if (!res.ok) {
    let message = `Upload failed (${res.status})`;
    try {
      const payload = await res.json();
      if (payload && payload.error) {
        message = payload.error;
      }
    } catch {
      // Keep fallback message.
    }
    throw new Error(message);
  }

  const payload = await res.json();
  state.workbooks = payload.workbooks || state.workbooks;
  state.selectedMainWorkbook = payload.default_main || state.selectedMainWorkbook;
  state.selectedReferenceWorkbook = payload.default_reference || state.selectedReferenceWorkbook;
  state.cache.clear();
  state.mainStyledRequestKey = null;

  populateWorkbookSelect(el.mainWorkbookSelect, state.selectedMainWorkbook);
  populateWorkbookSelect(el.referenceWorkbookSelect, state.selectedReferenceWorkbook);
  updateWorkbookLabels();
  await loadSheets();

  el.uploadInput.value = "";
  const skippedCount = Array.isArray(payload.skipped_files) ? payload.skipped_files.length : 0;
  const skippedText = skippedCount ? ` · skipped ${skippedCount} unsupported file(s)` : "";
  el.statusText.textContent = `Uploaded ${files.length} file(s)${skippedText}`;
}

function applySheetsPayload(payload) {
  const previousVersion = state.pairVersion;
  state.mainSheetTabs = payload.main_sheet_tabs || [];
  state.referenceSheetTabs = payload.reference_sheet_tabs || [];
  state.pairVersion = payload.version || null;

  if (payload.main_workbook) {
    state.selectedMainWorkbook = payload.main_workbook;
  }
  if (payload.reference_workbook) {
    state.selectedReferenceWorkbook = payload.reference_workbook;
  }

  updateWorkbookLabels();
  const changed = previousVersion !== state.pairVersion;
  if (changed) {
    state.mainStyledRequestKey = null;
    state.mainAvailableMonths.clear();
  }
  return changed;
}

async function loadSheets() {
  const query = workbookQuery();
  const payload = await fetchJson(`/api/sheets?${query}`);
  applySheetsPayload(payload);

  const previousMainSheet = state.selectedMainSheetName;
  rebuildMainTabs(previousMainSheet);

  const previousReferenceSheet = state.selectedReferenceSheetName;
  rebuildReferenceTabs(previousReferenceSheet);

  if (state.selectedRefCanonical) {
    await ensureSheetLoaded(state.selectedRefCanonical);
  }

  const currentMainTab = getMainTabMeta();
  if (currentMainTab && currentMainTab.canonical) {
    await ensureSheetLoaded(currentMainTab.canonical);
  }

  await render();
}

async function ensureSheetLoaded(canonical) {
  const key = cacheKey(canonical);
  if (state.cache.has(key)) {
    return state.cache.get(key);
  }

  el.statusText.textContent = `Loading ${canonical}...`;
  const query = workbookQuery();
  const payload = await fetchJson(`/api/sheet/${encodeURIComponent(canonical)}?${query}`);
  state.cache.set(key, payload);
  el.statusText.textContent = "";
  return payload;
}

function modeIsSameMonthYears(modeValue) {
  return modeValue === "same_month_years";
}

function currentN(inputElement) {
  const value = Number.parseInt(inputElement.value, 10);
  if (!Number.isFinite(value) || value < 1) {
    return 1;
  }
  return Math.min(value, 60);
}

function filterElementsForView(view) {
  if (view === "main") {
    return {
      modeSelect: el.mainModeSelect,
      nInput: el.mainNInput,
      monthWrapper: el.mainMonthWrapper,
      monthSelect: el.mainMonthSelect,
    };
  }
  return {
    modeSelect: el.refModeSelect,
    nInput: el.refNInput,
    monthWrapper: el.refMonthWrapper,
    monthSelect: el.refMonthSelect,
  };
}

function updateMonthControl(sheetData, view) {
  const controls = filterElementsForView(view);
  const previousValue = controls.monthSelect.value;

  if (!modeIsSameMonthYears(controls.modeSelect.value)) {
    controls.monthWrapper.classList.add("hidden");
    controls.monthSelect.innerHTML = "";
    return false;
  }

  let availableMonths = [];
  if (
    view === "main" &&
    state.selectedMainSheetName &&
    state.mainAvailableMonths.has(state.selectedMainSheetName)
  ) {
    const detected = state.mainAvailableMonths.get(state.selectedMainSheetName) || [];
    availableMonths = [...new Set(detected)].sort((a, b) => a - b);
  } else if (sheetData && Array.isArray(sheetData.months) && sheetData.months.length) {
    availableMonths = [...new Set(sheetData.months.map((item) => item.month))].sort((a, b) => a - b);
  } else if (view === "main") {
    availableMonths = Array.from({ length: 12 }, (_, idx) => idx + 1);
  }

  if (!availableMonths.length) {
    controls.monthWrapper.classList.add("hidden");
    controls.monthSelect.innerHTML = "";
    return false;
  }

  controls.monthWrapper.classList.remove("hidden");
  controls.monthSelect.innerHTML = "";

  for (const monthIndex of availableMonths) {
    const option = document.createElement("option");
    option.value = String(monthIndex);
    option.textContent = monthLabels[monthIndex - 1];
    controls.monthSelect.appendChild(option);
  }

  const current = Number.parseInt(controls.monthSelect.dataset.current || "", 10);
  const preferred = availableMonths.includes(current)
    ? current
    : availableMonths[availableMonths.length - 1];
  controls.monthSelect.value = String(preferred);
  return controls.monthSelect.value !== previousValue;
}

function hasAnyMetricValue(monthCell) {
  if (!monthCell) {
    return false;
  }
  return monthCell.pk !== null || monthCell.bottle !== null || monthCell.liter !== null;
}

function monthCoverageMap(sheetData) {
  const coverage = new Map();
  for (const month of sheetData.months) {
    coverage.set(month.key, 0);
  }

  for (const row of sheetData.rows) {
    for (const month of sheetData.months) {
      const monthCell = row.values[month.key];
      if (hasAnyMetricValue(monthCell)) {
        coverage.set(month.key, (coverage.get(month.key) || 0) + 1);
      }
    }
  }

  return coverage;
}

function pickMonths(sheetData, view) {
  const sorted = [...sheetData.months].sort((a, b) => a.key.localeCompare(b.key));
  const controls = filterElementsForView(view);
  const n = currentN(controls.nInput);
  const coverage = monthCoverageMap(sheetData);
  const minRowsForPopulated = Math.max(2, Math.ceil(sheetData.rows.length * 0.05));

  if (!modeIsSameMonthYears(controls.modeSelect.value)) {
    const populated = sorted.filter((month) => (coverage.get(month.key) || 0) >= minRowsForPopulated);
    const fallback = sorted.filter((month) => (coverage.get(month.key) || 0) > 0);
    const source = populated.length ? populated : fallback.length ? fallback : sorted;
    return source.slice(-n);
  }

  const selectedMonth = Number.parseInt(controls.monthSelect.value, 10);
  const filtered = sorted.filter((item) => item.month === selectedMonth);
  return filtered.slice(-n);
}

function annotateCellGrid(table) {
  const rows = Array.from(table.rows);
  const spanMap = [];
  let maxCols = 0;

  for (const row of rows) {
    for (let idx = 0; idx < spanMap.length; idx += 1) {
      if (spanMap[idx] > 0) {
        spanMap[idx] -= 1;
      }
    }

    let col = 0;
    for (const cell of row.cells) {
      while (spanMap[col] > 0) {
        col += 1;
      }

      const colSpan = Math.max(1, Number.parseInt(cell.getAttribute("colspan") || "1", 10) || 1);
      const rowSpan = Math.max(1, Number.parseInt(cell.getAttribute("rowspan") || "1", 10) || 1);

      cell.dataset.colStart = String(col);
      cell.dataset.colSpan = String(colSpan);

      if (rowSpan > 1) {
        for (let offset = 0; offset < colSpan; offset += 1) {
          const target = col + offset;
          spanMap[target] = Math.max(spanMap[target] || 0, rowSpan - 1);
        }
      }

      col += colSpan;
      if (col > maxCols) {
        maxCols = col;
      }
    }
  }

  return maxCols;
}

function measureColumnWidths(table, totalCols) {
  const widths = new Array(totalCols).fill(0);
  const colElements = Array.from(table.querySelectorAll("colgroup col"));
  if (colElements.length >= totalCols) {
    for (let idx = 0; idx < totalCols; idx += 1) {
      const width = Math.ceil(colElements[idx].getBoundingClientRect().width);
      if (Number.isFinite(width) && width > 0) {
        widths[idx] = width;
      }
    }
  }

  for (const cell of table.querySelectorAll("th, td")) {
    const start = Number.parseInt(cell.dataset.colStart || "-1", 10);
    const span = Number.parseInt(cell.dataset.colSpan || "1", 10);
    if (!Number.isFinite(start) || start < 0 || !Number.isFinite(span) || span !== 1) {
      continue;
    }
    const width = Math.ceil(cell.getBoundingClientRect().width);
    if (Number.isFinite(width) && width > widths[start]) {
      widths[start] = width;
    }
  }

  for (let idx = 0; idx < widths.length; idx += 1) {
    if (!widths[idx] || widths[idx] < 1) {
      widths[idx] = idx > 0 ? widths[idx - 1] : 96;
    }
    widths[idx] = Math.max(48, widths[idx]);
  }

  return widths;
}

function shouldSplitFrozenViewport(table, frozenCount, widths, totalCols) {
  if (frozenCount < 2 || frozenCount >= totalCols) {
    return false;
  }

  const frozenWidth = widths.slice(0, frozenCount).reduce((sum, width) => sum + width, 0);
  const totalWidth = widths.reduce((sum, width) => sum + width, 0);
  const wrap = tableWrapFor(table);
  const viewportWidth = wrap ? wrap.clientWidth : table.parentElement ? table.parentElement.clientWidth : 0;

  if (viewportWidth > 0 && frozenWidth > viewportWidth * 0.48) {
    return true;
  }
  if (totalWidth > 0 && frozenWidth / totalWidth > 0.55) {
    return true;
  }
  return frozenCount >= 8;
}

function pruneTableColumns(table, startCol, endColExclusive) {
  const colgroup = table.querySelector("colgroup");
  if (colgroup) {
    const cols = Array.from(colgroup.children);
    for (let idx = 0; idx < cols.length; idx += 1) {
      if (idx < startCol || idx >= endColExclusive) {
        cols[idx].remove();
      }
    }
  }

  for (const row of Array.from(table.rows)) {
    for (const cell of Array.from(row.cells)) {
      const cellStart = Number.parseInt(cell.dataset.colStart || "-1", 10);
      const cellSpan = Math.max(1, Number.parseInt(cell.dataset.colSpan || "1", 10) || 1);
      if (!Number.isFinite(cellStart) || cellStart < 0) {
        continue;
      }

      const cellEnd = cellStart + cellSpan;
      const overlapStart = Math.max(cellStart, startCol);
      const overlapEnd = Math.min(cellEnd, endColExclusive);
      const visibleSpan = overlapEnd - overlapStart;

      if (visibleSpan <= 0) {
        cell.remove();
        continue;
      }

      if (visibleSpan !== cellSpan) {
        cell.colSpan = visibleSpan;
      }

      cell.dataset.colStart = String(overlapStart - startCol);
      cell.dataset.colSpan = String(visibleSpan);
    }
  }
}

function applySplitViewport(table, frozenCount, totalCols, widths) {
  const host = table.parentElement;
  if (!host) {
    return false;
  }

  const leftTable = table.cloneNode(true);
  clearStickyStyles(leftTable);
  clearStickyStyles(table);

  pruneTableColumns(leftTable, 0, frozenCount);
  pruneTableColumns(table, frozenCount, totalCols);

  const splitView = document.createElement("div");
  splitView.className = "split-view";

  const leftPane = document.createElement("div");
  leftPane.className = "split-pane split-pane-left";
  leftPane.appendChild(leftTable);

  const rightPane = document.createElement("div");
  rightPane.className = "split-pane split-pane-right";
  rightPane.appendChild(table);

  splitView.appendChild(leftPane);
  splitView.appendChild(rightPane);
  host.replaceChild(splitView, table);

  const wrap = tableWrapFor(table);
  if (wrap) {
    wrap.classList.add("table-wrap-split");
    const height = wrap.clientHeight;
    if (height > 0) {
      const maxHeight = `${height}px`;
      leftPane.style.maxHeight = maxHeight;
      rightPane.style.maxHeight = maxHeight;
    }
  }

  const leftWidth = widths.slice(0, frozenCount).reduce((sum, width) => sum + width, 0);
  const rightWidth = widths.slice(frozenCount).reduce((sum, width) => sum + width, 0);
  leftPane.style.minWidth = leftWidth > 0 ? `${Math.min(leftWidth, 560)}px` : "";
  rightPane.style.minWidth = rightWidth > 0 ? "0" : "";

  let syncLock = false;
  leftPane.addEventListener(
    "scroll",
    () => {
      if (syncLock) {
        return;
      }
      syncLock = true;
      rightPane.scrollTop = leftPane.scrollTop;
      syncLock = false;
    },
    { passive: true },
  );
  rightPane.addEventListener(
    "scroll",
    () => {
      if (syncLock) {
        return;
      }
      syncLock = true;
      leftPane.scrollTop = rightPane.scrollTop;
      syncLock = false;
    },
    { passive: true },
  );

  return true;
}

function enhanceFrozenViewport(table, frozenCount) {
  if (!table) {
    return;
  }

  const normalizedTable = unwrapSplitViewport(table);
  const wrap = tableWrapFor(normalizedTable);
  if (wrap) {
    wrap.classList.remove("table-wrap-split");
  }

  if (!Number.isFinite(frozenCount) || frozenCount < 1) {
    clearStickyStyles(normalizedTable);
    return;
  }

  const totalCols = annotateCellGrid(normalizedTable);
  if (!totalCols) {
    return;
  }

  const clampedFrozen = Math.min(Math.max(1, frozenCount), totalCols);
  const widths = measureColumnWidths(normalizedTable, totalCols);
  if (shouldSplitFrozenViewport(normalizedTable, clampedFrozen, widths, totalCols)) {
    const applied = applySplitViewport(normalizedTable, clampedFrozen, totalCols, widths);
    if (applied) {
      return;
    }
  }

  applyFrozenColumns(normalizedTable, clampedFrozen);
}

function parseCssColor(text) {
  if (!text) {
    return null;
  }
  const value = text.trim().toLowerCase();
  if (!value || value === "transparent") {
    return { r: 0, g: 0, b: 0, a: 0 };
  }

  const rgbMatch = value.match(/^rgba?\(([^)]+)\)$/);
  if (rgbMatch) {
    const parts = rgbMatch[1].split(",").map((part) => part.trim());
    if (parts.length < 3) {
      return null;
    }
    const r = Number.parseFloat(parts[0]);
    const g = Number.parseFloat(parts[1]);
    const b = Number.parseFloat(parts[2]);
    const a = parts.length >= 4 ? Number.parseFloat(parts[3]) : 1;
    if (![r, g, b, a].every((part) => Number.isFinite(part))) {
      return null;
    }
    return {
      r: Math.max(0, Math.min(255, r)),
      g: Math.max(0, Math.min(255, g)),
      b: Math.max(0, Math.min(255, b)),
      a: Math.max(0, Math.min(1, a)),
    };
  }

  const hexMatch = value.match(/^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/);
  if (hexMatch) {
    const hex = hexMatch[1];
    if (hex.length === 3) {
      const r = Number.parseInt(hex[0] + hex[0], 16);
      const g = Number.parseInt(hex[1] + hex[1], 16);
      const b = Number.parseInt(hex[2] + hex[2], 16);
      return { r, g, b, a: 1 };
    }
    if (hex.length === 6) {
      const r = Number.parseInt(hex.slice(0, 2), 16);
      const g = Number.parseInt(hex.slice(2, 4), 16);
      const b = Number.parseInt(hex.slice(4, 6), 16);
      return { r, g, b, a: 1 };
    }
    const a = Number.parseInt(hex.slice(0, 2), 16) / 255;
    const r = Number.parseInt(hex.slice(2, 4), 16);
    const g = Number.parseInt(hex.slice(4, 6), 16);
    const b = Number.parseInt(hex.slice(6, 8), 16);
    return { r, g, b, a: Math.max(0, Math.min(1, a)) };
  }

  return null;
}

function blendColors(foreground, background) {
  const alpha = foreground.a ?? 1;
  const r = foreground.r * alpha + background.r * (1 - alpha);
  const g = foreground.g * alpha + background.g * (1 - alpha);
  const b = foreground.b * alpha + background.b * (1 - alpha);
  return {
    r: Math.round(Math.max(0, Math.min(255, r))),
    g: Math.round(Math.max(0, Math.min(255, g))),
    b: Math.round(Math.max(0, Math.min(255, b))),
  };
}

function toSolidRgbString(color) {
  return `rgb(${Math.round(color.r)}, ${Math.round(color.g)}, ${Math.round(color.b)})`;
}

function resolveOpaqueStickyBackground(cell, table) {
  const fallbackBaseText =
    getComputedStyle(document.documentElement).getPropertyValue("--bg-soft").trim() || "#102733";
  const fallbackBase = parseCssColor(fallbackBaseText) || { r: 16, g: 39, b: 51, a: 1 };

  const candidates = [
    getComputedStyle(cell).backgroundColor,
    cell.parentElement ? getComputedStyle(cell.parentElement).backgroundColor : null,
    getComputedStyle(table).backgroundColor,
  ];

  for (const candidate of candidates) {
    const parsed = parseCssColor(candidate || "");
    if (!parsed || (parsed.a ?? 0) <= 0) {
      continue;
    }
    if ((parsed.a ?? 1) >= 0.995) {
      return toSolidRgbString(parsed);
    }
    return toSolidRgbString(blendColors(parsed, fallbackBase));
  }

  return toSolidRgbString(fallbackBase);
}

function applyFrozenColumns(table, frozenCount) {
  if (!table || !Number.isFinite(frozenCount) || frozenCount < 1) {
    return;
  }

  const totalCols = annotateCellGrid(table);
  if (!totalCols) {
    return;
  }

  const clampedFrozen = Math.min(frozenCount, totalCols);
  if (clampedFrozen < 1) {
    return;
  }

  const widths = measureColumnWidths(table, totalCols);
  const leftOffsets = new Array(clampedFrozen).fill(0);
  let runningLeft = 0;
  for (let idx = 0; idx < clampedFrozen; idx += 1) {
    leftOffsets[idx] = runningLeft;
    runningLeft += widths[idx];
  }

  for (const cell of table.querySelectorAll("th, td")) {
    cell.classList.remove("sticky-col", "sticky-col-boundary", "sticky-col-head");
    cell.style.removeProperty("left");
    cell.style.removeProperty("z-index");
    cell.style.removeProperty("background-color");
  }

  for (const cell of table.querySelectorAll("th, td")) {
    const start = Number.parseInt(cell.dataset.colStart || "-1", 10);
    const span = Number.parseInt(cell.dataset.colSpan || "1", 10);
    if (!Number.isFinite(start) || start < 0 || !Number.isFinite(span) || span < 1) {
      continue;
    }

    const end = start + span;
    if (start >= clampedFrozen || end > clampedFrozen) {
      continue;
    }

    cell.classList.add("sticky-col");
    cell.style.left = `${leftOffsets[start]}px`;
    if (cell.tagName === "TH") {
      cell.classList.add("sticky-col-head");
      cell.style.zIndex = "12";
    } else {
      cell.style.zIndex = "9";
    }
    if (end === clampedFrozen) {
      cell.classList.add("sticky-col-boundary");
    }

    cell.style.backgroundColor = resolveOpaqueStickyBackground(cell, table);
  }
}

function formatValue(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "number") {
    const rounded =
      Math.abs(value) >= 1000
        ? value.toLocaleString(undefined, { maximumFractionDigits: 2 })
        : value.toFixed(2);
    if (Number.isInteger(value)) {
      return value.toLocaleString();
    }
    return rounded.replace(/\.00$/, "");
  }
  return String(value);
}

function filteredRows(rows) {
  const q = el.searchInput.value.trim().toLowerCase();
  if (!q) {
    return rows;
  }
  return rows.filter((row) => row.product_name.toLowerCase().includes(q));
}

function renderTable(target, sheetData, selectedMonths) {
  const metric = el.metricSelect.value;
  const rows = filteredRows(sheetData.rows);

  if (!rows.length) {
    target.innerHTML =
      '<tbody><tr><td class="empty">No rows match current filter.</td></tr></tbody>';
    return;
  }

  const columns = [
    { key: "sr", label: "Sr", cls: "right" },
    { key: "product_name", label: "Product Name", cls: "left" },
    { key: "ml", label: "Ml", cls: "right" },
    { key: "packing", label: "Packing", cls: "left" },
  ];

  const dynamicColumns = [];
  for (const month of selectedMonths) {
    if (metric === "all") {
      dynamicColumns.push(
        { monthKey: month.key, metric: "pk", label: `${month.label} PK` },
        { monthKey: month.key, metric: "bottle", label: `${month.label} Bottle` },
        { monthKey: month.key, metric: "liter", label: `${month.label} Liter` },
      );
    } else {
      dynamicColumns.push({
        monthKey: month.key,
        metric,
        label: month.label,
      });
    }
  }

  if (!dynamicColumns.length) {
    target.innerHTML =
      '<tbody><tr><td class="empty">No month columns matched this filter. Adjust ref mode, month, or N value.</td></tr></tbody>';
    return;
  }

  const thead = `
    <thead>
      <tr>
        ${columns.map((col) => `<th class="${col.cls}">${col.label}</th>`).join("")}
        ${dynamicColumns.map((col) => `<th class="right">${col.label}</th>`).join("")}
      </tr>
    </thead>
  `;

  const bodyRows = rows.map((row) => {
    const staticCells = columns
      .map((col) => `<td class="${col.cls}">${formatValue(row[col.key])}</td>`)
      .join("");

    const dynamicCells = dynamicColumns
      .map((col) => {
        const monthCell = row.values[col.monthKey] || {};
        return `<td class="right">${formatValue(monthCell[col.metric])}</td>`;
      })
      .join("");

    return `<tr>${staticCells}${dynamicCells}</tr>`;
  });

  target.innerHTML = `${thead}<tbody>${bodyRows.join("")}</tbody>`;
}

function renderMeta(target, sheetData, selectedMonths) {
  target.textContent = `${sheetData.sheet_name} · rows ${sheetData.rows.length} · showing ${selectedMonths.length} month group(s)`;
}

function mainStyledKey(selectedMonths) {
  const monthsPart = selectedMonths.map((item) => item.key).join(",");
  const mode = el.mainModeSelect.value;
  const nValue = currentN(el.mainNInput);
  const monthValue = modeIsSameMonthYears(mode) ? el.mainMonthSelect.value || "auto" : "none";
  return `${state.selectedMainWorkbook}::${state.selectedReferenceWorkbook}::${state.pairVersion}::${state.selectedMainSheetName}::${mode}::${nValue}::${monthValue}::${monthsPart}`;
}

async function renderMainStyled(selectedMonths) {
  if (!state.selectedMainSheetName) {
    setEmptyMain("No main sheet selected.");
    return;
  }

  const requestKey = mainStyledKey(selectedMonths);
  if (requestKey === state.mainStyledRequestKey) {
    return;
  }

  state.mainStyledRequestKey = requestKey;
  const query = workbookQuery();
  const monthKeys = selectedMonths.map((item) => item.key).join(",");
  const mode = el.mainModeSelect.value;
  const nValue = currentN(el.mainNInput);
  const selectedMonth = Number.parseInt(el.mainMonthSelect.value, 10);
  const url =
    `/api/main-styled-sheet?${query}` +
    `&sheet=${encodeURIComponent(state.selectedMainSheetName)}` +
    `&month_keys=${encodeURIComponent(monthKeys)}` +
    `&mode=${encodeURIComponent(mode)}` +
    `&n=${encodeURIComponent(String(nValue))}` +
    (Number.isFinite(selectedMonth) ? `&month=${encodeURIComponent(String(selectedMonth))}` : "");

  const payload = await fetchJson(url);
  if (state.mainStyledRequestKey !== requestKey) {
    return;
  }

  if (Array.isArray(payload.available_months) && payload.available_months.length) {
    const normalizedMonths = payload.available_months
      .map((value) => Number.parseInt(String(value), 10))
      .filter((value) => Number.isFinite(value) && value >= 1 && value <= 12)
      .sort((a, b) => a - b);
    if (state.selectedMainSheetName) {
      state.mainAvailableMonths.set(state.selectedMainSheetName, [...new Set(normalizedMonths)]);
    }
  } else if (state.selectedMainSheetName) {
    state.mainAvailableMonths.delete(state.selectedMainSheetName);
  }

  const monthAdjusted = updateMonthControl(null, "main");
  if (monthAdjusted && modeIsSameMonthYears(el.mainModeSelect.value)) {
    el.mainMonthSelect.dataset.current = el.mainMonthSelect.value;
    state.mainStyledRequestKey = null;
    await render();
    return;
  }

  el.mainTable.innerHTML = payload.html;
  const mainSourceTable = el.mainTable.querySelector("table.main-source-table");
  const frozenCount = Number.parseInt(String(payload.frozen_count ?? 0), 10) || 0;
  applyFrozenColumns(mainSourceTable, frozenCount);

  const monthCount = (payload.selected_month_labels || []).length;
  if (monthCount > 0) {
    el.mainMeta.textContent = `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · ${monthCount} month group(s)`;
  } else if (payload.filterable) {
    el.mainMeta.textContent = `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · fixed-layout mode`;
  } else {
    el.mainMeta.textContent = `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · full-sheet mode`;
  }
}

function renderReferencePanel(referenceSheet, selectedMonthsRef, referenceTab) {
  if (!referenceSheet) {
    if (referenceTab && !referenceTab.filterable) {
      setEmptyRef("Selected reference sheet is not in a compatible monthly-detail format.");
      el.refMeta.textContent = `${referenceTab.sheet_name} · non-filterable`;
      return;
    }
    setEmptyRef("No compatible monthly-detail data found for this reference sheet.");
    if (referenceTab) {
      el.refMeta.textContent = `${referenceTab.sheet_name} · no parsed month groups`;
    }
    return;
  }

  renderTable(el.refTable, referenceSheet, selectedMonthsRef);
  applyFrozenColumns(el.refTable, 4);
  renderMeta(el.refMeta, referenceSheet, selectedMonthsRef);

  if (!selectedMonthsRef.length) {
    el.refMeta.textContent += " · no months matched this mode";
  }
}

async function render() {
  const mainTab = getMainTabMeta();
  if (!mainTab) {
    setEmptyMain("No main sheet selected.");
    return;
  }

  const referenceTab = getReferenceTabMeta();
  const refPayload = getPayloadForCanonical(state.selectedRefCanonical);
  const refSheetData = refPayload ? refPayload.reference : null;

  const mainPayload = mainTab.canonical ? getPayloadForCanonical(mainTab.canonical) : null;
  const mainSheetData = mainPayload ? mainPayload.main : null;

  updateMonthControl(mainSheetData, "main");
  updateMonthControl(refSheetData, "ref");

  const selectedMonthsMain = mainSheetData ? pickMonths(mainSheetData, "main") : [];
  await renderMainStyled(selectedMonthsMain);

  const selectedMonthsRef = refSheetData ? pickMonths(refSheetData, "ref") : [];
  renderReferencePanel(refSheetData, selectedMonthsRef, referenceTab);

  const mainModeText = modeIsSameMonthYears(el.mainModeSelect.value)
    ? `Main: same month over past ${currentN(el.mainNInput)} year(s)`
    : `Main: past ${currentN(el.mainNInput)} month group(s)`;

  const refModeText = modeIsSameMonthYears(el.refModeSelect.value)
    ? `Ref: same month over past ${currentN(el.refNInput)} year(s)`
    : `Ref: past ${currentN(el.refNInput)} populated month(s)`;

  el.statusText.textContent = `Live refresh every ${POLL_INTERVAL_MS / 1000}s · ${mainModeText} · ${refModeText}`;
}

async function refreshRealtime() {
  if (state.isRefreshing || !state.selectedMainWorkbook || !state.selectedReferenceWorkbook) {
    return;
  }

  state.isRefreshing = true;
  try {
    const query = workbookQuery();
    const payload = await fetchJson(`/api/sheets?${query}`);
    const versionChanged = payload.version !== state.pairVersion;

    if (!versionChanged) {
      return;
    }

    applySheetsPayload(payload);

    const previousReferenceSheet = state.selectedReferenceSheetName;
    rebuildReferenceTabs(previousReferenceSheet);
    if (state.selectedRefCanonical) {
      await ensureSheetLoaded(state.selectedRefCanonical);
    }

    const previousMainSheet = state.selectedMainSheetName;
    rebuildMainTabs(previousMainSheet);
    const mainTab = getMainTabMeta();
    if (mainTab && mainTab.canonical) {
      await ensureSheetLoaded(mainTab.canonical);
    }

    await render();
  } catch (err) {
    el.statusText.textContent = err.message;
  } finally {
    state.isRefreshing = false;
  }
}

function startRealtimePolling() {
  if (state.pollHandle) {
    clearInterval(state.pollHandle);
  }
  state.pollHandle = setInterval(() => {
    refreshRealtime().catch((err) => {
      el.statusText.textContent = err.message;
    });
  }, POLL_INTERVAL_MS);
}

async function onWorkbookChange() {
  state.selectedMainWorkbook = el.mainWorkbookSelect.value;
  state.selectedReferenceWorkbook = el.referenceWorkbookSelect.value;
  state.mainStyledRequestKey = null;
  updateWorkbookLabels();
  await loadSheets();
}

function bindEvents() {
  el.uploadBtn.addEventListener("click", () => {
    uploadSelectedWorkbooks().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.mainWorkbookSelect.addEventListener("change", () => {
    onWorkbookChange().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.referenceWorkbookSelect.addEventListener("change", () => {
    onWorkbookChange().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.referenceTabSelect.addEventListener("change", () => {
    onReferenceTabChange(el.referenceTabSelect.value).catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.mainTabSelect.addEventListener("change", () => {
    onMainTabChange(el.mainTabSelect.value).catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.mainModeSelect.addEventListener("change", () => {
    state.mainStyledRequestKey = null;
    render().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.mainNInput.addEventListener("input", () => {
    state.mainStyledRequestKey = null;
    render().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.mainMonthSelect.addEventListener("change", () => {
    el.mainMonthSelect.dataset.current = el.mainMonthSelect.value;
    state.mainStyledRequestKey = null;
    render().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.refModeSelect.addEventListener("change", () => {
    render().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.refNInput.addEventListener("input", () => {
    render().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.refMonthSelect.addEventListener("change", () => {
    el.refMonthSelect.dataset.current = el.refMonthSelect.value;
    render().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.metricSelect.addEventListener("change", () => {
    render().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });

  el.searchInput.addEventListener("input", () => {
    render().catch((err) => {
      el.statusText.textContent = err.message;
    });
  });
}

(async function init() {
  try {
    bindEvents();
    await loadWorkbookOptions();
    await loadSheets();
    startRealtimePolling();
  } catch (err) {
    el.statusText.textContent = err.message;
  }
})();
