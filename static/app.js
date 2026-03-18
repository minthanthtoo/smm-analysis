const App = window.SMMApp || {};
const state = App.state;
const el = App.el;
const monthLabels = App.monthLabels;
const ribbonApi = App.ribbon || {};
const appUtils = App.utils || {};

if (!state || !el || !Array.isArray(monthLabels)) {
  throw new Error("App modules failed to initialize. Load app_state.js and ribbon.js before app.js.");
}

const requiredUtils = [
  "formatClockTime",
  "formatMetricNumber",
  "normalizeSheetName",
  "sanitizeSheetTabs",
  "roleLabel",
  "normalizeRoleToken",
  "regionLabel",
  "escapeHtml",
  "modeIsSameMonthYears",
  "modeIsMultiMonthYears",
  "modeUsesMonthSelector",
  "currentN",
  "hasAnyMetricValue",
  "monthCoverageMap",
  "isNumericCellText",
  "selectionScopeLabel",
  "normalizeViewScope",
  "columnToExcelLabel",
  "pointToAddress",
  "parseNumericCellValue",
  "formatValue",
];
const missingUtils = requiredUtils.filter((name) => typeof appUtils[name] !== "function");
if (missingUtils.length > 0) {
  throw new Error(`App utility module missing helpers: ${missingUtils.join(", ")}`);
}

const {
  formatClockTime,
  formatMetricNumber,
  normalizeSheetName,
  sanitizeSheetTabs,
  roleLabel,
  normalizeRoleToken,
  regionLabel,
  escapeHtml,
  modeIsSameMonthYears,
  modeIsMultiMonthYears,
  modeUsesMonthSelector,
  currentN,
  hasAnyMetricValue,
  monthCoverageMap,
  isNumericCellText,
  selectionScopeLabel,
  normalizeViewScope,
  columnToExcelLabel,
  pointToAddress,
  parseNumericCellValue,
  formatValue,
} = appUtils;

const POLL_INTERVAL_MS = 30000;
const MIN_ZOOM_PERCENT = 60;
const MAX_ZOOM_PERCENT = 180;
const ZOOM_STEP = 10;
const LONG_PRESS_OPEN_MS = 480;
const MIN_VIEWS_MAIN_RATIO = 0.28;
const MAX_VIEWS_MAIN_RATIO = 0.78;
const VIEWS_SPLIT_KEY_STEP = 0.03;
const THEME_PREF_STORAGE_KEY = "smm.themePreference";
const RIBBON_COLLAPSED_STORAGE_KEY = "smm.ribbonCollapsed";
const MIN_COLUMN_WIDTH_PX = 56;
const MAX_COLUMN_WIDTH_PX = 1400;
const MAX_AUTO_FROZEN_ROWS = 6;
const AUTO_FROZEN_ROW_RATIO_THRESHOLD = 0.25;
const GRID_SELECTION_EDGE_CLASSNAMES = [
  "grid-cell-edge-top",
  "grid-cell-edge-right",
  "grid-cell-edge-bottom",
  "grid-cell-edge-left",
];
const GRID_SELECTION_CLASSNAMES = ["grid-cell-selected", "grid-cell-active", ...GRID_SELECTION_EDGE_CLASSNAMES];
const FORMULA_FUNCTION_NAMES = new Set([
  "SUM",
  "AVERAGE",
  "AVG",
  "MIN",
  "MAX",
  "COUNT",
  "COUNTA",
  "IF",
  "ABS",
  "ROUND",
  "ROUNDUP",
  "ROUNDDOWN",
  "LEN",
]);
const themeMediaQuery =
  typeof window.matchMedia === "function" ? window.matchMedia("(prefers-color-scheme: dark)") : null;
let floatingMetricsRafId = 0;
let inMemoryClipboardText = "";
let activeColumnResize = null;

function readStoredRibbonCollapsed() {
  try {
    if (!window.localStorage) {
      return null;
    }
    const rawValue = window.localStorage.getItem(RIBBON_COLLAPSED_STORAGE_KEY);
    if (rawValue === "true") {
      return true;
    }
    if (rawValue === "false") {
      return false;
    }
  } catch {
    // Ignore storage read failures.
  }
  return null;
}

function storeRibbonCollapsed(collapsed) {
  try {
    if (!window.localStorage) {
      return;
    }
    window.localStorage.setItem(RIBBON_COLLAPSED_STORAGE_KEY, collapsed ? "true" : "false");
  } catch {
    // Ignore storage write failures.
  }
}

function scheduleFloatingLayoutMetricsSync() {
  if (typeof window.requestAnimationFrame !== "function") {
    syncFloatingLayoutMetrics();
    return;
  }
  if (floatingMetricsRafId) {
    return;
  }
  floatingMetricsRafId = window.requestAnimationFrame(() => {
    floatingMetricsRafId = 0;
    syncFloatingLayoutMetrics();
  });
}

const createThemeApi = App.createThemeApi;
if (typeof createThemeApi !== "function") {
  throw new Error("Theme module failed to initialize. Load theme.js before app.js.");
}
const { readStoredThemePreference, applyThemePreference } = createThemeApi({
  state,
  el,
  themeMediaQuery,
  storageKey: THEME_PREF_STORAGE_KEY,
  scheduleLayoutSync: scheduleFloatingLayoutMetricsSync,
});

function copySelectOptions(source, target) {
  if (typeof ribbonApi.copySelectOptions === "function") {
    ribbonApi.copySelectOptions(source, target);
    return;
  }
  if (!source || !target) {
    return;
  }
  target.innerHTML = "";
}

function setRibbonTab(name) {
  if (typeof ribbonApi.setRibbonTab === "function") {
    ribbonApi.setRibbonTab(el, name);
  }
  scheduleFloatingLayoutMetricsSync();
}

function triggerChange(control) {
  if (typeof ribbonApi.triggerChange === "function") {
    ribbonApi.triggerChange(control);
    return;
  }
  if (!control) {
    return;
  }
  control.dispatchEvent(new Event("change", { bubbles: true }));
}

function triggerInput(control) {
  if (typeof ribbonApi.triggerInput === "function") {
    ribbonApi.triggerInput(control);
    return;
  }
  if (!control) {
    return;
  }
  control.dispatchEvent(new Event("input", { bubbles: true }));
}

function syncUploadRegionInputs(source) {
  if (typeof ribbonApi.syncUploadRegionInputs === "function") {
    ribbonApi.syncUploadRegionInputs(el, source);
    return;
  }
  const value = source && "value" in source ? source.value : "";
  if (el.uploadRegionInput && source !== el.uploadRegionInput) {
    el.uploadRegionInput.value = value;
  }
}

function fullscreenElementActive() {
  return (
    document.fullscreenElement ||
    document.webkitFullscreenElement ||
    document.mozFullScreenElement ||
    document.msFullscreenElement ||
    null
  );
}

function isFullscreenSupported() {
  return Boolean(
    document.fullscreenEnabled ||
    document.webkitFullscreenEnabled ||
    document.mozFullScreenEnabled ||
    document.msFullscreenEnabled,
  );
}

function requestFullscreenCompat(target) {
  if (!target) {
    return Promise.reject(new Error("Fullscreen target is unavailable."));
  }
  if (typeof target.requestFullscreen === "function") {
    return Promise.resolve(target.requestFullscreen());
  }
  if (typeof target.webkitRequestFullscreen === "function") {
    return Promise.resolve(target.webkitRequestFullscreen());
  }
  if (typeof target.mozRequestFullScreen === "function") {
    return Promise.resolve(target.mozRequestFullScreen());
  }
  if (typeof target.msRequestFullscreen === "function") {
    return Promise.resolve(target.msRequestFullscreen());
  }
  return Promise.reject(new Error("Fullscreen is not supported by this browser."));
}

function exitFullscreenCompat() {
  if (typeof document.exitFullscreen === "function") {
    return Promise.resolve(document.exitFullscreen());
  }
  if (typeof document.webkitExitFullscreen === "function") {
    return Promise.resolve(document.webkitExitFullscreen());
  }
  if (typeof document.mozCancelFullScreen === "function") {
    return Promise.resolve(document.mozCancelFullScreen());
  }
  if (typeof document.msExitFullscreen === "function") {
    return Promise.resolve(document.msExitFullscreen());
  }
  return Promise.resolve();
}

function syncFullscreenToggleButton() {
  const active = Boolean(fullscreenElementActive());
  document.body.classList.toggle("app-fullscreen-active", active);
  if (!el.fullscreenToggleBtn) {
    return;
  }
  const supported = isFullscreenSupported() || active;
  el.fullscreenToggleBtn.disabled = !supported;
  el.fullscreenToggleBtn.setAttribute("aria-pressed", active ? "true" : "false");
  setText(el.fullscreenToggleBtn, active ? "Exit full screen" : "Full screen");
  el.fullscreenToggleBtn.title = active ? "Exit full screen (Esc)" : "Enter full screen";
}

async function toggleFullscreenMode() {
  if (!isFullscreenSupported() && !fullscreenElementActive()) {
    setText(el.statusText, "Fullscreen is not supported in this browser.");
    syncFullscreenToggleButton();
    return;
  }
  if (fullscreenElementActive()) {
    await exitFullscreenCompat();
    setText(el.statusText, "Exited full screen.");
    return;
  }
  const target = document.documentElement || document.body;
  await requestFullscreenCompat(target);
  setText(el.statusText, "Entered full screen.");
}

function syncRibbonFromCore() {
  if (typeof ribbonApi.syncRibbonFromCore === "function") {
    ribbonApi.syncRibbonFromCore({
      el,
      state,
      roleLabel,
      setText,
    });
  }
}

function scrollToNode(node) {
  if (typeof ribbonApi.scrollToNode === "function") {
    ribbonApi.scrollToNode(node);
    return;
  }
  if (!node || typeof node.scrollIntoView !== "function") {
    return;
  }
  node.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function refreshAllCommandData() {
  beginBusy("Refreshing all sections...");
  try {
    await refreshAccessContext();
    await loadWorkbookOptions();
    await loadSheets();
    await loadFiles();
  } finally {
    endBusy();
  }
}

function syncFloatingLayoutMetrics() {
  const root = document.documentElement;
  if (!root) {
    return;
  }

  const height = (node) => {
    if (!node || !(node instanceof HTMLElement)) {
      return 0;
    }
    if (node.hidden) {
      return 0;
    }
    const rect = node.getBoundingClientRect();
    if (!Number.isFinite(rect.height) || rect.height <= 0) {
      return 0;
    }
    return Math.ceil(rect.height);
  };

  const chromeHeight = height(el.excelChrome);
  const workbookBarHeight = height(el.workbookBars);
  const ribbonTabsBarHeight = height(el.ribbonTabsBar);
  const ribbonPanelHeight = height(el.ribbonRoot);
  const formulaBarHeight = height(el.formulaBar);

  let topOffset = 0;
  root.style.setProperty("--floating-top-chrome-offset", `${topOffset}px`);
  topOffset += chromeHeight;
  root.style.setProperty("--floating-top-workbook-offset", `${topOffset}px`);
  topOffset += workbookBarHeight;
  root.style.setProperty("--floating-top-tabs-offset", `${topOffset}px`);
  topOffset += ribbonTabsBarHeight;
  root.style.setProperty("--floating-top-ribbon-offset", `${topOffset}px`);
  topOffset += ribbonPanelHeight;
  root.style.setProperty("--floating-top-formula-offset", `${topOffset}px`);
  topOffset += formulaBarHeight;
  root.style.setProperty("--floating-top-total", `${topOffset}px`);

  const statusHeight = height(el.excelStatusBar);
  const tabsDockHeight = height(el.sheetTabsDock);
  root.style.setProperty("--floating-bottom-status-h", `${statusHeight}px`);
  root.style.setProperty("--floating-bottom-tabs-h", `${tabsDockHeight}px`);
  root.style.setProperty("--floating-bottom-total", `${statusHeight + tabsDockHeight}px`);
}

function setText(node, value) {
  if (!node) {
    return;
  }
  node.textContent = value;
}

function appendText(node, value) {
  if (!node) {
    return;
  }
  node.textContent = `${node.textContent || ""}${value}`;
}

function isEditableEventTarget(target) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  if (target.isContentEditable) {
    return true;
  }
  const tag = target.tagName;
  return tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT";
}

function setChip(node, text, tone) {
  if (!node) {
    return;
  }
  node.textContent = text;
  if (tone) {
    node.dataset.state = tone;
  }
}

function clampZoomPercent(value) {
  const parsed = Number.parseInt(String(value || ""), 10);
  if (!Number.isFinite(parsed)) {
    return 100;
  }
  const snapped = Math.round(parsed / ZOOM_STEP) * ZOOM_STEP;
  return Math.max(MIN_ZOOM_PERCENT, Math.min(MAX_ZOOM_PERCENT, snapped));
}

function roleOnboardingText(role) {
  if (role === "rsm") {
    return "RSM: Owner သတ်မှတ်ထားသော Region များထဲတွင် workbook/file များကို စီမံနိုင်ပါသည်။";
  }
  if (role === "asm") {
    return "ASM: Assigned region ထဲမှ ခွင့်ပြုထားသော township sheet များကိုသာ ကြည့်နိုင်ပါသည်။";
  }
  if (role === "user") {
    return "User: Assigned RSM scope ထဲရှိ region data ကိုသာ ကြည့်နိုင်ပါသည်။";
  }
  return "Owner: Region အားလုံးကို ကြည့်နိုင်ပြီး RSM/ASM mapping နှင့် regional permission ကို စီမံနိုင်ပါသည်။";
}

function updateRoleCards(role) {
  const roleMap = {
    owner: el.roleCardOwner,
    rsm: el.roleCardRsm,
    asm: el.roleCardAsm,
    user: el.roleCardUser,
  };
  for (const [key, node] of Object.entries(roleMap)) {
    if (!node) {
      continue;
    }
    node.dataset.active = key === role ? "true" : "false";
  }
}

function updateLoadHealth(stageText) {
  const hasMain = Boolean(state.selectedMainWorkbook);
  const hasReference = Boolean(state.selectedReferenceWorkbook);
  const mainTabs = state.mainSheetTabs.length;
  const refTabs = state.referenceSheetTabs.length;

  setChip(
    el.mainLoadChip,
    `Main workbook: ${hasMain ? state.selectedMainWorkbook : "မရွေးရသေး"}`,
    hasMain ? "ok" : "warn",
  );
  setChip(
    el.referenceLoadChip,
    `Reference workbook: ${hasReference ? state.selectedReferenceWorkbook : "မရွေးရသေး"}`,
    hasReference ? "ok" : "warn",
  );
  setChip(el.mainTabsChip, `Main tabs: ${mainTabs}`, mainTabs > 0 ? "ok" : "warn");
  setChip(el.referenceTabsChip, `Ref tabs: ${refTabs}`, refTabs > 0 ? "ok" : "warn");

  const stage =
    stageText ||
    (state.lastLoadAt
      ? `Status: နောက်ဆုံး load အချိန် ${formatClockTime(state.lastLoadAt)}`
      : "Status: စတင်ရန်စောင့်နေသည်");
  const tone = state.busyDepth > 0 ? "busy" : stageText ? "warn" : state.lastLoadAt ? "ok" : "warn";
  setChip(el.loadStageChip, stage, tone);

  if (el.roleOnboardingText) {
    setText(el.roleOnboardingText, roleOnboardingText(state.viewerRole));
  }
  updateRoleCards(state.viewerRole);
  if (el.statusHint) {
    if (hasMain && hasReference && (mainTabs === 0 || refTabs === 0)) {
      setText(
        el.statusHint,
        "Tab များမတွေ့ပါက Role/Region scope ကိုပြန်စစ်ပါ။ Hidden sheet များသည် tab တွင်မပါဝင်ပါ။",
      );
    } else {
      const mainPart = mainTabs > 0 ? `${mainTabs} main tabs` : "main tabs မရှိသေး";
      const refPart = refTabs > 0 ? `${refTabs} ref tabs` : "ref tabs မရှိသေး";
      setText(el.statusHint, `Visible sheets only loaded: ${mainPart} / ${refPart}. Hidden sheets မပါဝင်ပါ။`);
    }
  }
}

function beginBusy(message) {
  state.busyDepth += 1;
  document.body.classList.add("is-loading");
  setInteractiveControlsDisabled(true);
  if (message) {
    setText(el.statusText, message);
  }
  updateLoadHealth(`Status: ${message || "လုပ်ဆောင်နေသည်..."}`);
}

function endBusy() {
  state.busyDepth = Math.max(0, state.busyDepth - 1);
  if (state.busyDepth === 0) {
    document.body.classList.remove("is-loading");
    setInteractiveControlsDisabled(false);
  }
  updateLoadHealth();
}

function setStatusError(err) {
  const message = err instanceof Error ? err.message : String(err || "Unknown error");
  setText(el.statusText, message);
  updateLoadHealth(`Status: ${message}`);
}

function setRoleInfoTriggerExpanded(expanded) {
  const value = expanded ? "true" : "false";
  if (el.roleInfoBtn) {
    el.roleInfoBtn.setAttribute("aria-expanded", value);
  }
  if (el.roleInfoInlineBtn) {
    el.roleInfoInlineBtn.setAttribute("aria-expanded", value);
  }
}

function getRoleInfoModalFocusables() {
  if (!el.roleInfoModal) {
    return [];
  }
  return Array.from(
    el.roleInfoModal.querySelectorAll(
      "button, [href], input, select, textarea, [tabindex]:not([tabindex='-1'])",
    ),
  ).filter((node) => !node.hasAttribute("disabled") && node.getAttribute("aria-hidden") !== "true");
}

function setModalOpen(open, triggerElement = null) {
  if (!el.roleInfoModal) {
    return;
  }
  if (open) {
    const activeElement = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    state.modalReturnFocus =
      triggerElement instanceof HTMLElement ? triggerElement : activeElement || state.modalReturnFocus;
    el.roleInfoModal.hidden = false;
    el.roleInfoModal.setAttribute("aria-hidden", "false");
    setRoleInfoTriggerExpanded(true);
    document.body.classList.add("modal-open");
    const focusTarget =
      el.roleInfoCloseBtn ||
      el.roleInfoDoneBtn ||
      el.roleInfoModal.querySelector(".modal-card");
    if (focusTarget && typeof focusTarget.focus === "function") {
      window.requestAnimationFrame(() => {
        focusTarget.focus();
      });
    }
    return;
  }
  el.roleInfoModal.hidden = true;
  el.roleInfoModal.setAttribute("aria-hidden", "true");
  setRoleInfoTriggerExpanded(false);
  document.body.classList.remove("modal-open");
  const returnTarget = state.modalReturnFocus;
  state.modalReturnFocus = null;
  if (returnTarget && document.contains(returnTarget)) {
    window.requestAnimationFrame(() => {
      returnTarget.focus();
    });
  }
}

function setOnboardingExpanded(expanded) {
  if (!el.onboardingSteps || !el.onboardingToggleBtn) {
    return;
  }
  el.onboardingSteps.classList.toggle("hidden", !expanded);
  setText(
    el.onboardingToggleBtn,
    expanded ? "အသေးစိတ် onboarding ပိတ်ရန်" : "အသေးစိတ် onboarding ပြရန်",
  );
}

function setRibbonCollapsed(collapsed, options = {}) {
  const { persist = true } = options;
  state.ribbonCollapsed = Boolean(collapsed);
  document.body.classList.toggle("ribbon-collapsed", state.ribbonCollapsed);
  if (el.ribbonToggleBtn) {
    el.ribbonToggleBtn.setAttribute("aria-expanded", state.ribbonCollapsed ? "false" : "true");
    setText(el.ribbonToggleBtn, state.ribbonCollapsed ? "Expand Ribbon" : "Collapse Ribbon");
  }
  if (persist) {
    storeRibbonCollapsed(state.ribbonCollapsed);
  }
  scheduleFloatingLayoutMetricsSync();
}

function normalizeZoomScope(scope) {
  if (scope === "main" || scope === "reference") {
    return scope;
  }
  return normalizeViewScope(state.activeViewScope || "main");
}

function ensureZoomState() {
  if (!state.sheetZoomByScope || typeof state.sheetZoomByScope !== "object") {
    const fallback = clampZoomPercent(state.sheetZoomPercent || 100);
    state.sheetZoomByScope = {
      main: fallback,
      reference: fallback,
    };
  }
  const mainZoom = clampZoomPercent(state.sheetZoomByScope.main ?? state.sheetZoomPercent ?? 100);
  const referenceZoom = clampZoomPercent(state.sheetZoomByScope.reference ?? state.sheetZoomPercent ?? 100);
  state.sheetZoomByScope.main = mainZoom;
  state.sheetZoomByScope.reference = referenceZoom;
}

function zoomPercentForScope(scope) {
  ensureZoomState();
  const normalized = normalizeZoomScope(scope);
  return clampZoomPercent(state.sheetZoomByScope[normalized] ?? state.sheetZoomPercent ?? 100);
}

function applyScopeZoom(scope) {
  const normalized = normalizeZoomScope(scope);
  const zoom = zoomPercentForScope(normalized);
  const wrap = tableWrapForScope(normalized);
  if (wrap) {
    wrap.style.setProperty("--sheet-zoom", String(zoom / 100));
  }
}

function syncZoomControlForScope(scope = null) {
  const normalized = normalizeZoomScope(scope);
  const zoom = zoomPercentForScope(normalized);
  state.sheetZoomPercent = zoom;
  if (el.zoomRangeInput && el.zoomRangeInput.value !== String(zoom)) {
    el.zoomRangeInput.value = String(zoom);
  }
  if (el.zoomValueLabel) {
    setText(el.zoomValueLabel, `${zoom}%`);
  }
}

function syncAllScopeZooms() {
  ensureZoomState();
  applyScopeZoom("main");
  applyScopeZoom("reference");
  syncZoomControlForScope();
}

function setSheetZoom(percent, scope = null) {
  ensureZoomState();
  const normalized = normalizeZoomScope(scope);
  const zoom = clampZoomPercent(percent);
  state.sheetZoomByScope[normalized] = zoom;
  state.sheetZoomPercent = zoom;
  applyScopeZoom(normalized);
  syncZoomControlForScope();
}

function adjustSheetZoom(delta, scope = null) {
  const normalized = normalizeZoomScope(scope);
  const next = clampZoomPercent(zoomPercentForScope(normalized) + delta);
  setSheetZoom(next, normalized);
}

function resetSelectionStatusBar() {
  refreshStatusSelectionScopeLabel();
  setText(el.statusSelectionAddress, "Ready");
  setText(el.statusSelectionRange, "Select cells to see details");
  setText(el.statusSelectionCellCount, "Cells: 0");
  setText(el.statusSelectionCount, "CountA: 0");
  setText(el.statusSelectionNumericCount, "Numbers: 0");
  setText(el.statusSelectionSum, "Sum: -");
  setText(el.statusSelectionAvg, "Avg: -");
  setText(el.statusSelectionMin, "Min: -");
  setText(el.statusSelectionMax, "Max: -");
  if (el.formulaNameBox) {
    el.formulaNameBox.value = "A1";
  }
  if (el.formulaInput) {
    el.formulaInput.value = "";
  }
  scheduleFloatingLayoutMetricsSync();
}

function clampViewsMainRatio(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 0.54;
  }
  return Math.min(MAX_VIEWS_MAIN_RATIO, Math.max(MIN_VIEWS_MAIN_RATIO, numeric));
}

function setViewsSplitRatio(ratio) {
  state.viewsSplitRatio = clampViewsMainRatio(ratio);
  if (!el.viewsSplit) {
    return;
  }
  const percent = `${(state.viewsSplitRatio * 100).toFixed(2)}%`;
  el.viewsSplit.style.setProperty("--views-main-width", percent);
}

function isDetailPanelVisible() {
  const drawerMode = document.body.classList.contains("detail-drawer-mode");
  if (drawerMode) {
    return Boolean(state.referenceDrawerExpanded);
  }
  return !Boolean(state.desktopDetailCollapsed);
}

function syncSheetTabsDockVisibility() {
  if (!el.sheetTabsDock || !el.referenceSheetTabs || !el.mainSheetTabs) {
    return;
  }

  const referenceLane = el.referenceSheetTabs.closest(".sheet-tabs-dock-lane");
  const mainLane = el.mainSheetTabs ? el.mainSheetTabs.closest(".sheet-tabs-dock-lane") : null;
  const drawerMode = document.body.classList.contains("detail-drawer-mode");
  const collapsedDesktop =
    document.body.classList.contains("desktop-detail-collapsed") || Boolean(state.desktopDetailCollapsed);
  const drawerExpanded = drawerMode
    ? Boolean(
        el.referenceViewPanel instanceof HTMLElement
          ? el.referenceViewPanel.classList.contains("drawer-expanded")
          : state.referenceDrawerExpanded,
      )
    : true;

  let panelVisibleByLayout = !collapsedDesktop;
  if (el.referenceViewPanel instanceof HTMLElement) {
    const style = window.getComputedStyle(el.referenceViewPanel);
    const rect = el.referenceViewPanel.getBoundingClientRect();
    panelVisibleByLayout =
      style.display !== "none" &&
      style.visibility !== "hidden" &&
      rect.width > 2 &&
      rect.height > 2;
  }

  const showDetailTabs = drawerMode ? drawerExpanded : !collapsedDesktop && panelVisibleByLayout;

  if (referenceLane instanceof HTMLElement) {
    referenceLane.hidden = !showDetailTabs;
    referenceLane.setAttribute("aria-hidden", showDetailTabs ? "false" : "true");
    referenceLane.style.display = showDetailTabs ? "" : "none";
  }
  el.sheetTabsDock.classList.toggle("single-lane", !showDetailTabs);
  document.body.classList.toggle("detail-tabs-hidden", !showDetailTabs);

  if (showDetailTabs) {
    document.documentElement.style.removeProperty("--sheet-tabs-dock-height");
  } else if (mainLane instanceof HTMLElement) {
    const laneHeight = Math.max(36, Math.ceil(mainLane.getBoundingClientRect().height));
    document.documentElement.style.setProperty("--sheet-tabs-dock-height", `${laneHeight}px`);
  }

  scheduleFloatingLayoutMetricsSync();
}

function setDesktopDetailCollapsed(collapsed) {
  const drawerMode = document.body.classList.contains("detail-drawer-mode");
  state.desktopDetailCollapsed = Boolean(collapsed);
  const collapsedDesktop = state.desktopDetailCollapsed && !drawerMode;
  document.body.classList.toggle("desktop-detail-collapsed", collapsedDesktop);

  if (collapsedDesktop && !hasGridSelection() && state.activeViewScope === "reference") {
    setActiveViewScope("main");
  }

  if (el.referencePanelToggleBtn) {
    el.referencePanelToggleBtn.setAttribute("aria-expanded", collapsedDesktop ? "false" : "true");
    setText(el.referencePanelToggleBtn, collapsedDesktop ? "Expand detail pane" : "Collapse detail pane");
  }

  if (el.viewsSplitHandle) {
    el.viewsSplitHandle.hidden = drawerMode || collapsedDesktop;
  }

  syncSheetTabsDockVisibility();
}

function shouldUseReferenceDrawerMode() {
  const width = Number(window.innerWidth || document.documentElement.clientWidth || 0);
  const height = Number(window.innerHeight || document.documentElement.clientHeight || 0);

  if (typeof window.matchMedia === "function" && window.matchMedia("(orientation: portrait)").matches) {
    return true;
  }

  if (width > 0 && height > 0 && height > width * 1.08) {
    return true;
  }

  return false;
}

function setReferenceDrawerExpanded(expanded) {
  if (!el.referenceViewPanel) {
    return;
  }

  const drawerMode = document.body.classList.contains("detail-drawer-mode");
  const effectiveExpanded = drawerMode ? Boolean(expanded) : true;
  state.referenceDrawerExpanded = effectiveExpanded;

  el.referenceViewPanel.classList.toggle("drawer-expanded", effectiveExpanded);
  document.body.classList.toggle("detail-drawer-open", drawerMode && effectiveExpanded);

  if (el.referenceDrawerToggleBtn) {
    el.referenceDrawerToggleBtn.setAttribute("aria-expanded", effectiveExpanded ? "true" : "false");
    setText(el.referenceDrawerToggleBtn, effectiveExpanded ? "Collapse details" : "Expand details");
  }

  if (el.referenceDrawerBody) {
    el.referenceDrawerBody.setAttribute("aria-hidden", drawerMode && !effectiveExpanded ? "true" : "false");
  }

  syncSheetTabsDockVisibility();
}

function syncReferenceDrawerMode() {
  const drawerMode = shouldUseReferenceDrawerMode();
  document.body.classList.toggle("detail-drawer-mode", drawerMode);

  if (el.referenceDrawerToggleBtn) {
    el.referenceDrawerToggleBtn.hidden = !drawerMode;
  }
  if (el.referencePanelToggleBtn) {
    el.referencePanelToggleBtn.hidden = drawerMode;
  }

  if (drawerMode) {
    setDesktopDetailCollapsed(false);
    setReferenceDrawerExpanded(state.referenceDrawerExpanded);
  } else {
    setReferenceDrawerExpanded(true);
    setDesktopDetailCollapsed(state.desktopDetailCollapsed);
    setViewsSplitRatio(state.viewsSplitRatio);
  }
  syncSheetTabsDockVisibility();
  syncFloatingLayoutMetrics();
}

function resizeViewsSplitFromClientX(clientX) {
  if (!el.viewsSplit || !el.viewsSplitHandle) {
    return;
  }
  const rect = el.viewsSplit.getBoundingClientRect();
  if (!Number.isFinite(rect.width) || rect.width <= 0) {
    return;
  }
  const handleRect = el.viewsSplitHandle.getBoundingClientRect();
  const handleWidth = Math.max(6, Number(handleRect.width || 0));
  const usableWidth = Math.max(1, rect.width - handleWidth);
  const ratio = (Number(clientX) - rect.left - handleWidth / 2) / usableWidth;
  setViewsSplitRatio(ratio);
}

function beginViewsSplitDrag(event) {
  if (!event || !el.viewsSplit || !el.viewsSplitHandle) {
    return;
  }
  if (Number(event.button) !== 0) {
    return;
  }
  if (document.body.classList.contains("detail-drawer-mode") || state.desktopDetailCollapsed) {
    return;
  }

  state.viewsSplitDragActive = true;
  el.viewsSplit.classList.add("is-resizing");

  const pointerId = Number(event.pointerId);
  if (Number.isFinite(pointerId) && typeof el.viewsSplitHandle.setPointerCapture === "function") {
    try {
      el.viewsSplitHandle.setPointerCapture(pointerId);
    } catch (_err) {
      // Ignore capture failures and still resize using document listeners.
    }
  }

  const stopDrag = () => {
    if (!state.viewsSplitDragActive) {
      return;
    }
    state.viewsSplitDragActive = false;
    el.viewsSplit.classList.remove("is-resizing");
    document.removeEventListener("pointermove", onPointerMove);
    document.removeEventListener("pointerup", stopDrag);
    document.removeEventListener("pointercancel", stopDrag);

    if (Number.isFinite(pointerId) && typeof el.viewsSplitHandle.releasePointerCapture === "function") {
      try {
        el.viewsSplitHandle.releasePointerCapture(pointerId);
      } catch (_err) {
        // Ignore capture release failures.
      }
    }
  };

  const onPointerMove = (moveEvent) => {
    if (!state.viewsSplitDragActive) {
      return;
    }
    resizeViewsSplitFromClientX(moveEvent.clientX);
  };

  document.addEventListener("pointermove", onPointerMove);
  document.addEventListener("pointerup", stopDrag);
  document.addEventListener("pointercancel", stopDrag);
  resizeViewsSplitFromClientX(event.clientX);
}

function setInteractiveControlsDisabled(disabled) {
  const controls = [
    el.userSelect,
    el.regionSelect,
    el.uploadInput,
    el.uploadRegionInput,
    el.uploadBtn,
    el.ribbonToggleBtn,
    el.mainWorkbookSelect,
    el.referenceWorkbookSelect,
    el.referenceWorkbookMirrorSelect,
    el.mainModeSelect,
    el.mainNInput,
    el.mainMonthSelect,
    el.referenceTabSelect,
    el.mainTabSelect,
    el.refModeSelect,
    el.refNInput,
    el.refMonthSelect,
    el.metricSelect,
    el.searchInput,
    el.onboardingToggleBtn,
    el.roleInfoBtn,
    el.roleInfoInlineBtn,
    el.fullscreenToggleBtn,
    el.titleUserSelect,
    el.logoutCurrentUserBtn,
    el.themeSelect,
    el.assignRsmUsernameInput,
    el.assignRsmDisplayInput,
    el.assignRsmRegionsInput,
    el.mapUserUsernameInput,
    el.mapUserRsmSelect,
    el.switchUserRoleUsernameInput,
    el.switchUserRoleDisplayInput,
    el.switchUserRoleSelect,
    el.switchUserRoleRsmSelect,
    el.switchUserRoleRegionsInput,
    el.assignAsmUsernameInput,
    el.assignAsmDisplayInput,
    el.assignAsmRsmSelect,
    el.assignAsmRegionSelect,
    el.assignAsmTownshipsSelect,
    el.asmTownshipUserSelect,
    el.asmTownshipRegionSelect,
    el.asmTownshipSelect,
    el.refreshFilesBtn,
    el.ribbonRefreshAllBtn,
    el.ribbonRefreshFilesBtn,
    el.ribbonRoleGuideBtn,
    el.ribbonOnboardingBtn,
    el.ribbonGoMainBtn,
    el.ribbonGoRefBtn,
    el.ribbonGoAccessBtn,
    el.ribbonGoFilesBtn,
    el.ribbonUserSelect,
    el.ribbonRegionSelect,
    el.ribbonMainWorkbookSelect,
    el.ribbonReferenceWorkbookSelect,
    el.ribbonMainModeSelect,
    el.ribbonMainNInput,
    el.ribbonMainMonthSelect,
    el.ribbonRefModeSelect,
    el.ribbonRefNInput,
    el.ribbonRefMonthSelect,
    el.ribbonMetricSelect,
    el.ribbonSearchInput,
    el.ribbonCollapseAllGroupsBtn,
    el.ribbonExpandAllGroupsBtn,
    el.ribbonOpenRoleGuideBtn,
    el.ribbonJumpAssignRsmBtn,
    el.ribbonJumpMapUserBtn,
    el.ribbonJumpAssignAsmBtn,
    el.ribbonJumpTownshipBtn,
    el.ribbonUploadInput,
    el.ribbonUploadRegionInput,
    el.ribbonUploadBtn,
    el.ribbonFreezePanesBtn,
    el.ribbonFreezeTopRowBtn,
    el.ribbonFreezeFirstColBtn,
    el.ribbonUnfreezePanesBtn,
    el.ribbonSelectRowBtn,
    el.ribbonSelectColumnBtn,
    el.ribbonSelectAllBtn,
    el.ribbonInsertRowAboveBtn,
    el.ribbonInsertRowBelowBtn,
    el.ribbonDeleteRowBtn,
    el.ribbonHideRowsBtn,
    el.ribbonUnhideRowsBtn,
    el.ribbonInsertColLeftBtn,
    el.ribbonInsertColRightBtn,
    el.ribbonDeleteColBtn,
    el.ribbonHideColsBtn,
    el.ribbonUnhideColsBtn,
    el.referenceDrawerToggleBtn,
    el.referencePanelToggleBtn,
    el.viewsSplitHandle,
    el.zoomOutBtn,
    el.zoomRangeInput,
    el.zoomInBtn,
    el.zoomResetBtn,
    el.formulaNameBox,
    el.formulaInput,
    el.formulaCancelBtn,
    el.formulaApplyBtn,
    el.formulaFxBtn,
    el.ctxCopyBtn,
    el.ctxPasteBtn,
    el.ctxClearSelectionBtn,
    el.ctxFreezePanesBtn,
    el.ctxUnfreezePanesBtn,
    el.ctxHideRowsBtn,
    el.ctxUnhideRowsBtn,
    el.ctxHideColsBtn,
    el.ctxUnhideColsBtn,
    el.ctxHideBothBtn,
    el.ctxUnhideBothBtn,
    el.ctxInsertRowAboveBtn,
    el.ctxInsertRowBelowBtn,
    el.ctxDeleteRowBtn,
    el.ctxInsertColLeftBtn,
    el.ctxInsertColRightBtn,
    el.ctxDeleteColBtn,
    el.ctxSelectRowBtn,
    el.ctxSelectColumnBtn,
    el.ctxSelectAllBtn,
    el.ctxOpenDetailsBtn,
    el.ctxZoomInBtn,
    el.ctxZoomOutBtn,
    el.sheetTabCtxRenameBtn,
    ...(Array.isArray(el.ribbonTabButtons) ? el.ribbonTabButtons : []),
  ];
  for (const control of controls) {
    if (!control) {
      continue;
    }
    control.disabled = Boolean(disabled);
  }

  for (const renameInput of Array.from(document.querySelectorAll(".sheet-tab-rename-input"))) {
    renameInput.disabled = Boolean(disabled);
  }
}

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

function scopeQuery() {
  const params = new URLSearchParams();
  params.set("user", state.currentUser || "owner");
  params.set("region", state.selectedRegion || "ALL");
  return params.toString();
}

function workbookQuery() {
  const params = new URLSearchParams(scopeQuery());
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
  return `${state.viewerRole}::${state.selectedRegion}::${state.selectedMainWorkbook}::${state.selectedReferenceWorkbook}::${version}::${canonical}`;
}

function populateWorkbookSelect(selectElement, selectedName, workbookOptions = null) {
  if (!selectElement) {
    return;
  }
  const options = Array.isArray(workbookOptions) ? workbookOptions : state.workbooks;
  selectElement.innerHTML = "";
  for (const workbook of options) {
    const option = document.createElement("option");
    option.value = workbook;
    option.textContent = workbook;
    selectElement.appendChild(option);
  }
  if (selectedName && options.includes(selectedName)) {
    selectElement.value = selectedName;
    return;
  }
  if (options.length) {
    selectElement.value = options[0];
  }
}

const FILE_VIEW_MODE_OPTIONS = [
  { value: "auto", label: "Auto" },
  { value: "main", label: "Main only" },
  { value: "detail", label: "Detail only" },
  { value: "both", label: "Main + Detail" },
];

function normalizeFileViewModeClient(rawMode) {
  const token = String(rawMode || "")
    .trim()
    .toLowerCase();
  if (token === "main") {
    return "main";
  }
  if (token === "detail" || token === "reference" || token === "ref") {
    return "detail";
  }
  if (token === "both") {
    return "both";
  }
  return "auto";
}

function fileViewModeLabel(mode) {
  const normalized = normalizeFileViewModeClient(mode);
  const option = FILE_VIEW_MODE_OPTIONS.find((item) => item.value === normalized);
  return option ? option.label : "Auto";
}

function displayUserOption(user) {
  const username = user && typeof user.username === "string" ? user.username : "";
  const displayName = user && typeof user.display_name === "string" ? user.display_name : "";
  const role = user && typeof user.role === "string" ? roleLabel(user.role) : "";
  if (displayName) {
    return role ? `${displayName} · ${role}` : displayName;
  }
  return role ? `${username || "-"} · ${role}` : username || "-";
}

function populateUserSelect() {
  const users = Array.isArray(state.loginUsers) && state.loginUsers.length ? state.loginUsers : Array.isArray(state.users) ? state.users : [];
  if (!el.userSelect && !el.titleUserSelect) {
    return;
  }
  if (el.userSelect) {
    el.userSelect.innerHTML = "";
  }
  if (el.titleUserSelect) {
    el.titleUserSelect.innerHTML = "";
  }

  for (const user of users) {
    if (!user || typeof user !== "object" || !user.username) {
      continue;
    }
    const label = displayUserOption(user);
    if (el.userSelect) {
      const option = document.createElement("option");
      option.value = user.username;
      option.textContent = label;
      option.title = user.username;
      el.userSelect.appendChild(option);
    }
    if (el.titleUserSelect) {
      const option = document.createElement("option");
      option.value = user.username;
      option.textContent = label;
      option.title = user.username;
      el.titleUserSelect.appendChild(option);
    }
  }

  if (!users.some((user) => user.username === state.currentUser)) {
    const fallback = users[0];
    state.currentUser = fallback && fallback.username ? fallback.username : state.currentUser;
  }
  if (state.currentUser && el.userSelect) {
    el.userSelect.value = state.currentUser;
  }
  if (state.currentUser && el.titleUserSelect) {
    el.titleUserSelect.value = state.currentUser;
  }
  if (el.titleUserSelect) {
    el.titleUserSelect.disabled = users.length <= 1;
  }
  if (el.logoutCurrentUserBtn) {
    el.logoutCurrentUserBtn.disabled = users.length < 1;
  }
  syncRibbonFromCore();
}

function populateRegionSelect() {
  const regions = Array.isArray(state.regions) ? state.regions : [];
  const allowAll = Boolean(state.canViewAllRegions);
  el.regionSelect.innerHTML = "";

  if (allowAll) {
    const allOption = document.createElement("option");
    allOption.value = "ALL";
    allOption.textContent = "All regions";
    el.regionSelect.appendChild(allOption);
  }

  for (const region of regions) {
    const option = document.createElement("option");
    option.value = region;
    option.textContent = regionLabel(region);
    el.regionSelect.appendChild(option);
  }

  if (!allowAll && state.selectedRegion === "ALL" && regions.length) {
    state.selectedRegion = regions[0];
  }
  if (state.selectedRegion !== "ALL" && !regions.includes(state.selectedRegion)) {
    state.selectedRegion = allowAll ? "ALL" : regions[0] || "ALL";
  }

  el.regionSelect.value = state.selectedRegion || (allowAll ? "ALL" : regions[0] || "ALL");
  el.regionSelect.disabled = !allowAll && regions.length <= 1;
  syncRibbonFromCore();
}

function applyScopePayload(payload) {
  if (!payload || typeof payload !== "object") {
    return;
  }
  const currentUserPayload = payload.current_user;
  if (currentUserPayload && typeof currentUserPayload === "object") {
    if (currentUserPayload.username) {
      state.currentUser = currentUserPayload.username;
    }
    if (currentUserPayload.display_name) {
      state.currentUserDisplayName = currentUserPayload.display_name;
    }
    if (currentUserPayload.role) {
      state.viewerRole = normalizeRoleToken(currentUserPayload.role);
    }
  }
  if (payload.current_user && typeof payload.current_user === "string") {
    state.currentUser = payload.current_user;
  }
  if (payload.viewer_role) {
    state.viewerRole = normalizeRoleToken(payload.viewer_role);
  }
  if (Array.isArray(payload.users)) {
    state.users = payload.users;
  }
  if (Array.isArray(payload.login_users)) {
    state.loginUsers = payload.login_users;
  } else {
    state.loginUsers = Array.isArray(payload.users) ? payload.users : state.loginUsers;
  }
  if (Array.isArray(payload.regions)) {
    state.regions = payload.regions;
  }
  if (Array.isArray(payload.workbooks)) {
    state.workbooks = payload.workbooks;
  }
  if (Array.isArray(payload.main_workbooks)) {
    state.mainWorkbooks = payload.main_workbooks;
  } else if (Array.isArray(payload.workbooks)) {
    state.mainWorkbooks = payload.workbooks;
  }
  if (Array.isArray(payload.reference_workbooks)) {
    state.referenceWorkbooks = payload.reference_workbooks;
  } else if (Array.isArray(payload.workbooks)) {
    state.referenceWorkbooks = payload.workbooks;
  }
  if (Array.isArray(payload.all_regions)) {
    state.allRegions = payload.all_regions;
  }
  if (payload.selected_region) {
    state.selectedRegion = payload.selected_region;
  }
  if (payload.permissions && typeof payload.permissions === "object") {
    state.permissions = {
      ...state.permissions,
      ...payload.permissions,
    };
  }
  if (payload.assignments && typeof payload.assignments === "object") {
    state.assignments = {
      rsm_regions: payload.assignments.rsm_regions || {},
      user_to_rsm: payload.assignments.user_to_rsm || {},
      asm_townships: payload.assignments.asm_townships || {},
    };
  }
  if (payload.region_townships && typeof payload.region_townships === "object") {
    state.regionTownships = payload.region_townships;
  }
  if (Object.prototype.hasOwnProperty.call(payload, "can_view_all_regions")) {
    state.canViewAllRegions = Boolean(payload.can_view_all_regions);
  } else {
    state.canViewAllRegions = state.viewerRole === "owner";
  }
  populateUserSelect();
  populateRegionSelect();
}

function updateWorkbookLabels() {
  setText(el.mainWorkbookName, state.selectedMainWorkbook || "-");
  setText(el.referenceWorkbookName, state.selectedReferenceWorkbook || "-");
  setText(el.mainWorkbookBarName, state.selectedMainWorkbook || "-");
  setText(el.referenceWorkbookBarName, state.selectedReferenceWorkbook || "-");
  setText(el.mainTabsDockLabel, `Main Sheets · ${state.selectedMainWorkbook || "-"}`);
  setText(el.referenceTabsDockLabel, `Detail Sheets · ${state.selectedReferenceWorkbook || "-"}`);
  setText(el.currentUserName, state.currentUserDisplayName || state.currentUser || "-");
  setText(el.viewerRoleName, roleLabel(state.viewerRole));
  setText(el.regionName, regionLabel(state.selectedRegion));
  setText(el.mainPanelTitle, `Main View (${state.selectedMainWorkbook || "-"})`);
  setText(el.referencePanelTitle, `Reference Detail (${state.selectedReferenceWorkbook || "-"})`);
  refreshStatusSelectionScopeLabel();
  updateLoadHealth();
  syncRibbonFromCore();
  syncFloatingLayoutMetrics();
}

function setSelectOptions(selectElement, values, selectedValue = null) {
  if (!selectElement) {
    return;
  }
  selectElement.innerHTML = "";
  for (const value of values) {
    const option = document.createElement("option");
    option.value = String(value);
    option.textContent = regionLabel(String(value));
    selectElement.appendChild(option);
  }
  if (selectedValue !== null && values.includes(selectedValue)) {
    selectElement.value = selectedValue;
    return;
  }
  if (values.length) {
    selectElement.value = String(values[0]);
  }
}

function setSelectOptionsFromUsers(selectElement, users, roleFilter = null, selectedValue = null) {
  if (!selectElement) {
    return;
  }
  const filtered = users.filter((user) => !roleFilter || user.role === roleFilter);
  selectElement.innerHTML = "";
  for (const user of filtered) {
    const option = document.createElement("option");
    option.value = user.username;
    option.textContent = displayUserOption(user);
    selectElement.appendChild(option);
  }
  if (selectedValue && filtered.some((user) => user.username === selectedValue)) {
    selectElement.value = selectedValue;
    return;
  }
  if (filtered.length) {
    selectElement.value = filtered[0].username;
  }
}

function selectedMultiValues(selectElement) {
  if (!selectElement) {
    return [];
  }
  return Array.from(selectElement.selectedOptions || [])
    .map((option) => option.value)
    .filter(Boolean);
}

function setTownshipSelect(selectElement, region, selectedTownships = []) {
  if (!selectElement) {
    return;
  }
  const townships = Array.isArray(state.regionTownships[region]) ? state.regionTownships[region] : [];
  selectElement.innerHTML = "";
  for (const township of townships) {
    const option = document.createElement("option");
    option.value = township;
    option.textContent = township;
    option.selected = selectedTownships.includes(township);
    selectElement.appendChild(option);
  }
}

function syncSwitchUserRoleControls() {
  if (!el.switchUserRoleSelect) {
    return;
  }
  const role = normalizeRoleToken(el.switchUserRoleSelect.value || "user");
  const needsRsm = role === "user" || role === "asm";
  const needsRegions = role === "rsm";

  if (el.switchUserRoleRsmWrap) {
    el.switchUserRoleRsmWrap.classList.toggle("hidden", !needsRsm);
  }
  if (el.switchUserRoleRsmSelect) {
    el.switchUserRoleRsmSelect.disabled = !needsRsm;
  }
  if (el.switchUserRoleRegionsWrap) {
    el.switchUserRoleRegionsWrap.classList.toggle("hidden", !needsRegions);
  }
  if (el.switchUserRoleRegionsInput) {
    el.switchUserRoleRegionsInput.disabled = !needsRegions;
  }
}

function renderUsersList() {
  if (!el.usersList) {
    return;
  }
  const users = Array.isArray(state.users) ? state.users : [];
  if (!users.length) {
    el.usersList.innerHTML = '<div class="empty">No visible users in current scope.</div>';
    return;
  }
  const rows = users
    .map((user) => {
      const manager = state.assignments.user_to_rsm[user.username] || "-";
      const regions = Array.isArray(user.rsm_regions) ? user.rsm_regions.join(", ") : "";
      const asmRegions = Array.isArray(user.asm_regions) ? user.asm_regions.join(", ") : "";
      return `
        <tr>
          <td class="left">${escapeHtml(user.display_name || user.username)}</td>
          <td class="left">${escapeHtml(user.username)}</td>
          <td class="left">${escapeHtml(roleLabel(user.role))}</td>
          <td class="left">${escapeHtml(manager)}</td>
          <td class="left">${escapeHtml(regions || asmRegions || "-")}</td>
        </tr>
      `;
    })
    .join("");

  el.usersList.innerHTML = `
    <div class="table-wrap compact-wrap">
      <table>
        <thead>
          <tr>
            <th class="left">Display Name</th>
            <th class="left">Username</th>
            <th class="left">Role</th>
            <th class="left">Mapped RSM</th>
            <th class="left">Region Scope</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function syncAccessControls() {
  const permissions = state.permissions || {};
  const users = Array.isArray(state.users) ? state.users : [];
  const rsmUsers = users.filter((user) => user.role === "rsm");
  const asmUsers = users.filter((user) => user.role === "asm");
  const regions = Array.isArray(state.regions) ? state.regions : [];

  if (el.assignRsmForm) {
    el.assignRsmForm.classList.toggle("hidden", !permissions.can_manage_rsm);
  }
  if (el.mapUserRsmForm) {
    el.mapUserRsmForm.classList.toggle("hidden", !permissions.can_manage_rsm);
  }
  if (el.switchUserRoleForm) {
    el.switchUserRoleForm.classList.toggle("hidden", !permissions.can_manage_rsm);
  }
  if (el.assignAsmForm) {
    el.assignAsmForm.classList.toggle("hidden", !permissions.can_manage_asm);
  }
  if (el.asmTownshipForm) {
    el.asmTownshipForm.classList.toggle("hidden", !permissions.can_manage_asm);
  }
  if (el.uploadBtn) {
    el.uploadBtn.disabled = !permissions.can_upload;
  }
  if (el.uploadInput) {
    el.uploadInput.disabled = !permissions.can_upload;
  }
  if (el.uploadRegionInput) {
    el.uploadRegionInput.disabled = !permissions.can_upload;
  }
  if (el.ribbonUploadBtn) {
    el.ribbonUploadBtn.disabled = !permissions.can_upload;
  }
  if (el.ribbonUploadInput) {
    el.ribbonUploadInput.disabled = !permissions.can_upload;
  }
  if (el.ribbonUploadRegionInput) {
    el.ribbonUploadRegionInput.disabled = !permissions.can_upload;
  }

  setSelectOptionsFromUsers(el.mapUserRsmSelect, rsmUsers, null, el.mapUserRsmSelect?.value || null);
  setSelectOptionsFromUsers(
    el.switchUserRoleRsmSelect,
    rsmUsers,
    null,
    el.switchUserRoleRsmSelect?.value || null,
  );
  setSelectOptionsFromUsers(el.assignAsmRsmSelect, rsmUsers, null, el.assignAsmRsmSelect?.value || null);
  setSelectOptionsFromUsers(
    el.asmTownshipUserSelect,
    asmUsers,
    null,
    el.asmTownshipUserSelect?.value || null,
  );
  setSelectOptions(el.assignAsmRegionSelect, regions, el.assignAsmRegionSelect?.value || null);
  setSelectOptions(el.asmTownshipRegionSelect, regions, el.asmTownshipRegionSelect?.value || null);

  if (el.assignAsmRsmWrap) {
    el.assignAsmRsmWrap.classList.toggle("hidden", state.viewerRole !== "owner");
  }
  syncSwitchUserRoleControls();

  const assignAsmRegion = el.assignAsmRegionSelect ? el.assignAsmRegionSelect.value : "";
  setTownshipSelect(el.assignAsmTownshipsSelect, assignAsmRegion, []);

  const asmUser = el.asmTownshipUserSelect ? el.asmTownshipUserSelect.value : "";
  const asmRegion = el.asmTownshipRegionSelect ? el.asmTownshipRegionSelect.value : "";
  const assignedTownships =
    state.assignments &&
    state.assignments.asm_townships &&
    state.assignments.asm_townships[asmUser] &&
    Array.isArray(state.assignments.asm_townships[asmUser][asmRegion])
      ? state.assignments.asm_townships[asmUser][asmRegion]
      : [];
  setTownshipSelect(el.asmTownshipSelect, asmRegion, assignedTownships);

  setText(
    el.accessHint,
    `${roleLabel(state.viewerRole)} (${state.currentUser || "-"}) · ` +
      `RSM manage: ${permissions.can_manage_rsm ? "yes" : "no"} · ` +
      `ASM manage: ${permissions.can_manage_asm ? "yes" : "no"} · ` +
      `Upload: ${permissions.can_upload ? "yes" : "no"}`,
  );

  renderUsersList();
  syncRibbonFromCore();
}

function scopedUrl(path) {
  const query = scopeQuery();
  return `${path}?${query}`;
}

async function postScopedJson(path, payload) {
  const mergedPayload = {
    ...payload,
    user: state.currentUser,
    region: state.selectedRegion,
  };
  const res = await fetch(scopedUrl(path), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(mergedPayload),
  });
  if (!res.ok) {
    let message = `Request failed (${res.status})`;
    try {
      const body = await res.json();
      if (body && body.error) {
        message = body.error;
      }
    } catch {
      // Keep fallback message.
    }
    throw new Error(message);
  }
  return res.json();
}

async function patchScopedJson(path, payload) {
  const res = await fetch(scopedUrl(path), {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    let message = `Request failed (${res.status})`;
    try {
      const body = await res.json();
      if (body && body.error) {
        message = body.error;
      }
    } catch {
      // Keep fallback message.
    }
    throw new Error(message);
  }
  return res.json();
}

async function deleteScoped(path) {
  const res = await fetch(scopedUrl(path), { method: "DELETE" });
  if (!res.ok) {
    let message = `Request failed (${res.status})`;
    try {
      const body = await res.json();
      if (body && body.error) {
        message = body.error;
      }
    } catch {
      // Keep fallback message.
    }
    throw new Error(message);
  }
  return res.json();
}

async function refreshAccessContext() {
  try {
    const payload = await fetchJson(scopedUrl("/api/access/context"));
    applyScopePayload(payload);
    syncAccessControls();
    updateWorkbookLabels();
    return payload;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error || "");
    if (!message.includes("(404)")) {
      throw error;
    }

    // Compatibility mode for older backend versions that don't expose /api/access/context.
    const legacy = await fetchJson(scopedUrl("/api/workbooks"));
    const fallbackUser = String(legacy.current_user || state.currentUser || "owner");
    const fallbackRole = normalizeRoleToken(legacy.viewer_role || state.viewerRole || "owner");
    const fallbackRegions = Array.isArray(legacy.regions) ? legacy.regions : [];
    applyScopePayload({
      ...legacy,
      current_user: {
        username: fallbackUser,
        display_name: fallbackUser,
        role: fallbackRole,
        assigned_rsm: null,
        rsm_regions: [],
        asm_regions: [],
      },
      users: [
        {
          username: fallbackUser,
          display_name: fallbackUser,
          role: fallbackRole,
          assigned_rsm: null,
          rsm_regions: [],
          asm_regions: [],
        },
      ],
      login_users: [
        {
          username: fallbackUser,
          display_name: fallbackUser,
          role: fallbackRole,
          assigned_rsm: null,
          rsm_regions: [],
          asm_regions: [],
        },
      ],
      permissions: {
        can_upload: fallbackRole === "owner" || fallbackRole === "rsm",
        can_manage_rsm: false,
        can_manage_asm: false,
        can_manage_files: false,
      },
      assignments: {
        rsm_regions: {},
        user_to_rsm: {},
        asm_townships: {},
      },
      all_regions: fallbackRegions,
      region_townships: {},
      can_view_all_regions: fallbackRole === "owner",
    });
    syncAccessControls();
    updateWorkbookLabels();
    setText(el.statusText, "Server is using compatibility mode (access APIs not available).");
    return legacy;
  }
}

function filesEditableRegionOptions() {
  if (state.viewerRole === "owner") {
    return state.allRegions.length ? state.allRegions : state.regions;
  }
  return state.regions;
}

function renderFilesTable() {
  if (!el.filesTableBody) {
    return;
  }
  const files = Array.isArray(state.fileRows) ? state.fileRows : [];
  if (!files.length) {
    el.filesTableBody.innerHTML =
      '<tr><td colspan="4" class="empty">No files found in this scope.</td></tr>';
    return;
  }

  const regionOptions = filesEditableRegionOptions();
  const rows = files
    .map((file) => {
      const nameEscaped = escapeHtml(file.name);
      const regionEscaped = escapeHtml(file.region);
      const viewMode = normalizeFileViewModeClient(file.view_mode);
      const mainEnabled = file.main_enabled !== false;
      const detailEnabled = file.detail_enabled !== false;
      const badges = [];
      if (mainEnabled) {
        badges.push('<span class="file-view-chip">Main</span>');
      }
      if (detailEnabled) {
        badges.push('<span class="file-view-chip">Detail</span>');
      }
      if (!badges.length) {
        badges.push('<span class="file-view-chip file-view-chip-muted">None</span>');
      }
      const modeBadgeHtml = `<div class="file-role-chips">${badges.join("")}</div>`;
      const uploadBadge = file.uploaded
        ? '<span class="file-uploaded-chip">Uploaded</span>'
        : "";
      const nameMetaHtml = uploadBadge
        ? `<div class="file-name-meta">${uploadBadge}</div>`
        : "";

      let regionControl = `<span class="file-region-readonly">${regionEscaped}</span>`;
      let viewControl = (
        `<div class="file-view-control-wrap">` +
        `<span class="file-view-readonly-label">${escapeHtml(fileViewModeLabel(viewMode))}</span>` +
        `${modeBadgeHtml}` +
        `</div>`
      );
      let actions = '<span class="muted-inline">View only</span>';
      if (file.can_update) {
        const selectOptions = regionOptions
          .map((region) => {
            const selected = region === file.region ? ' selected="selected"' : "";
            return `<option value="${escapeHtml(region)}"${selected}>${escapeHtml(region)}</option>`;
          })
          .join("");
        regionControl = `<select class="file-region-select file-inline-control" data-file="${nameEscaped}">${selectOptions}</select>`;
        const viewOptions = FILE_VIEW_MODE_OPTIONS
          .map((option) => {
            const selected = option.value === viewMode ? ' selected="selected"' : "";
            return `<option value="${escapeHtml(option.value)}"${selected}>${escapeHtml(option.label)}</option>`;
          })
          .join("");
        viewControl = (
          `<div class="file-view-control-wrap">` +
          `<select class="file-view-mode-select file-inline-control" data-file="${nameEscaped}">${viewOptions}</select>` +
          `${modeBadgeHtml}` +
          `</div>`
        );
      }
      if (file.can_update || file.can_delete) {
        actions = "";
        if (file.can_update) {
          actions += `<button type="button" class="action-btn action-btn-sm file-save-btn" data-file="${nameEscaped}">Save role</button>`;
        }
        if (file.can_delete) {
          actions += `<button type="button" class="action-btn action-btn-sm action-btn-danger file-delete-btn" data-file="${nameEscaped}">Delete file</button>`;
        }
      }
      return (
        `<tr>` +
        `<td class="left file-name-cell"><span class="file-name-text" title="${nameEscaped}">${nameEscaped}</span>${nameMetaHtml}</td>` +
        `<td class="left">${regionControl}</td>` +
        `<td class="left">${viewControl}</td>` +
        `<td class="left file-action-cell">${actions}</td>` +
        `</tr>`
      );
    })
    .join("");
  el.filesTableBody.innerHTML = rows;
}

async function loadFiles() {
  try {
    const payload = await fetchJson(scopedUrl("/api/files"));
    state.fileRows = Array.isArray(payload.files) ? payload.files : [];
    renderFilesTable();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error || "");
    if (!message.includes("(404)")) {
      throw error;
    }
    state.fileRows = [];
    renderFilesTable();
  }
}

async function refreshAccessAndFiles() {
  await refreshAccessContext();
  await loadFiles();
}

function setEmptyMain(message) {
  el.mainTable.innerHTML = `<div class="empty">${message}</div>`;
  setText(el.mainMeta, "");
  if (state.selectionScope === "main") {
    clearGridSelectionModel();
  }
}

function setEmptyRef(message) {
  unwrapSplitViewport(el.refTable);
  el.refTable.innerHTML = `<tbody><tr><td class="empty">${message}</td></tr></tbody>`;
  setText(el.refMeta, "");
  if (state.selectionScope === "reference") {
    clearGridSelectionModel();
  }
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

function excelRowHeightPx(row) {
  if (!row || !row.dataset) {
    return null;
  }
  const rawHeight = row.dataset.excelRowHeight;
  if (!rawHeight) {
    return null;
  }
  const parsed = Number.parseFloat(rawHeight);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }
  return Math.max(1, Math.round(parsed));
}

function restoreExcelRowHeight(row) {
  if (!row) {
    return null;
  }
  const rowHeightPx = excelRowHeightPx(row);
  if (!rowHeightPx) {
    row.style.removeProperty("height");
    row.style.removeProperty("min-height");
    return null;
  }
  const rowHeightCss = `${rowHeightPx}px`;
  row.style.height = rowHeightCss;
  row.style.minHeight = rowHeightCss;
  return rowHeightPx;
}

function effectiveRowHeightPx(row) {
  if (!row) {
    return 1;
  }
  const measured = Math.ceil(row.getBoundingClientRect().height);
  if (Number.isFinite(measured) && measured > 0) {
    return measured;
  }
  return excelRowHeightPx(row) || 1;
}

function clearStickyStyles(table) {
  if (!table) {
    return;
  }
  for (const row of table.rows || []) {
    restoreExcelRowHeight(row);
  }
  for (const cell of table.querySelectorAll("th, td")) {
    cell.classList.remove(
      "sticky-col",
      "sticky-col-boundary",
      "sticky-col-head",
      "sticky-row",
      "sticky-row-boundary",
      "split-head-cell",
    );
    cell.style.removeProperty("left");
    cell.style.removeProperty("position");
    cell.style.removeProperty("top");
    cell.style.removeProperty("z-index");
    cell.style.removeProperty("background-color");
    cell.style.removeProperty("height");
    cell.style.removeProperty("min-height");
    cell.style.removeProperty("box-sizing");
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

  const rowSyncObserver = splitView.__rowSyncObserver;
  if (rowSyncObserver && typeof rowSyncObserver.disconnect === "function") {
    rowSyncObserver.disconnect();
  }
  const splitCleanup = splitView.__splitCleanup;
  if (typeof splitCleanup === "function") {
    splitCleanup();
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
  setActiveViewScope("main");
  state.selectedMainSheetName = sheetName;
  state.mainStyledRequestKey = null;

  const tabMeta = getMainTabMeta(sheetName);
  if (tabMeta && tabMeta.canonical) {
    await ensureSheetLoaded(tabMeta.canonical);
  }

  rebuildMainTabs(state.selectedMainSheetName);
  await render();
}

function canEditWorkbookTabs() {
  const permissions = state.permissions || {};
  return Boolean(permissions.can_manage_files || permissions.can_upload);
}

async function renameWorkbookSheet(view, oldSheetName, nextSheetName) {
  const oldName = String(oldSheetName || "").trim();
  const newName = String(nextSheetName || "").trim();
  if (!oldName || !newName || oldName === newName) {
    return false;
  }

  beginBusy(`Renaming sheet "${oldName}"...`);
  try {
    const payload = await postScopedJson("/api/workbook/rename-sheet", {
      workbook: view,
      old_sheet_name: oldName,
      new_sheet_name: newName,
      main: state.selectedMainWorkbook,
      reference: state.selectedReferenceWorkbook,
    });

    state.cache.clear();
    state.mainStyledRequestKey = null;
    state.mainAvailableMonths.clear();
    state.collapsedRowGroups.clear();
    if (view === "main") {
      if (state.selectedMainSheetName === oldName) {
        state.selectedMainSheetName = payload.new_sheet_name || newName;
      }
    } else {
      if (state.selectedReferenceSheetName === oldName) {
        state.selectedReferenceSheetName = payload.new_sheet_name || newName;
      }
    }

    await loadSheets();
    setText(
      el.statusText,
      `Renamed ${view} sheet "${oldName}" to "${payload.new_sheet_name || newName}".`,
    );
    return true;
  } finally {
    endBusy();
  }
}

function beginSheetTabRename(view, tabName, tabItem, tabButton, editButton) {
  if (!canEditWorkbookTabs() || !tabItem || !tabButton) {
    return;
  }
  if (tabItem.dataset.renaming === "true") {
    return;
  }

  const currentName = String(tabName || "").trim();
  if (!currentName) {
    return;
  }

  tabItem.dataset.renaming = "true";
  tabButton.hidden = true;
  if (editButton) {
    editButton.hidden = true;
  }

  const input = document.createElement("input");
  input.type = "text";
  input.className = "sheet-tab-rename-input";
  input.maxLength = 31;
  input.value = currentName;
  input.setAttribute("aria-label", `Rename ${view} sheet ${currentName}`);
  tabItem.appendChild(input);

  let closed = false;
  const restore = () => {
    if (closed) {
      return;
    }
    closed = true;
    tabItem.dataset.renaming = "false";
    input.remove();
    tabButton.hidden = false;
    if (editButton) {
      editButton.hidden = false;
    }
  };

  const commit = async () => {
    if (closed) {
      return;
    }
    const candidate = String(input.value || "").trim();
    if (!candidate || candidate === currentName) {
      restore();
      tabButton.focus();
      return;
    }

    closed = true;
    tabItem.dataset.renaming = "false";
    input.remove();

    try {
      await renameWorkbookSheet(view, currentName, candidate);
    } catch (err) {
      if (tabItem.isConnected) {
        tabButton.hidden = false;
        if (editButton) {
          editButton.hidden = false;
        }
      }
      setStatusError(err);
    }
  };

  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      commit().catch((err) => {
        setStatusError(err);
      });
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      restore();
      tabButton.focus();
    }
  });
  input.addEventListener("blur", () => {
    commit().catch((err) => {
      setStatusError(err);
    });
  });

  window.requestAnimationFrame(() => {
    input.focus();
    input.select();
  });
}

function rebuildMainTabs(preferredSheetName) {
  hideSheetTabContextMenu();
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
    const tabItem = document.createElement("div");
    tabItem.className = "sheet-tab-item";

    const button = document.createElement("button");
    button.type = "button";
    button.className = `sheet-tab-btn${tab.sheet_name === validSheet ? " active" : ""}`;
    button.textContent = tab.sheet_name;
    button.dataset.sheetName = tab.sheet_name;
    const titleParts = [tab.sheet_name];
    if (!tab.filterable) {
      titleParts.push("Rendered as full sheet (non-month layout)");
    } else if (canEditWorkbookTabs()) {
      titleParts.push("Right-click or long-press to rename this sheet");
    }
    button.title = titleParts.join(" · ");
    button.addEventListener("click", (event) => {
      if (state.sheetTabLongPressTriggered) {
        state.sheetTabLongPressTriggered = false;
        event.preventDefault();
        return;
      }
      hideSheetTabContextMenu();
      onMainTabChange(tab.sheet_name).catch((err) => {
        setStatusError(err);
      });
    });
    tabItem.appendChild(button);
    bindSheetTabContextMenu(button, "main", tab.sheet_name, tabItem);

    el.mainSheetTabs.appendChild(tabItem);
  }

  const activeBtn = el.mainSheetTabs.querySelector(".sheet-tab-btn.active");
  if (activeBtn) {
    activeBtn.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
}

async function onReferenceTabChange(sheetName) {
  setActiveViewScope("reference");
  if (document.body.classList.contains("detail-drawer-mode")) {
    setReferenceDrawerExpanded(true);
  } else if (state.desktopDetailCollapsed) {
    setDesktopDetailCollapsed(false);
  }
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
  hideSheetTabContextMenu();
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
    const tabItem = document.createElement("div");
    tabItem.className = "sheet-tab-item";

    const button = document.createElement("button");
    button.type = "button";
    button.className = `sheet-tab-btn${tab.sheet_name === validSheet ? " active" : ""}`;
    button.textContent = tab.sheet_name;
    button.dataset.sheetName = tab.sheet_name;
    const titleParts = [tab.sheet_name];
    if (!tab.filterable) {
      titleParts.push("Detail parser not available for this sheet.");
    } else if (canEditWorkbookTabs()) {
      titleParts.push("Right-click or long-press to rename this sheet");
    }
    button.title = titleParts.join(" · ");
    button.addEventListener("click", (event) => {
      if (state.sheetTabLongPressTriggered) {
        state.sheetTabLongPressTriggered = false;
        event.preventDefault();
        return;
      }
      hideSheetTabContextMenu();
      onReferenceTabChange(tab.sheet_name).catch((err) => {
        setStatusError(err);
      });
    });
    tabItem.appendChild(button);
    bindSheetTabContextMenu(button, "reference", tab.sheet_name, tabItem);

    el.referenceSheetTabs.appendChild(tabItem);
  }

  const activeBtn = el.referenceSheetTabs.querySelector(".sheet-tab-btn.active");
  if (activeBtn) {
    activeBtn.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
}

async function loadWorkbookOptions() {
  beginBusy("Excel file များကို ရှာဖွေနေသည်...");
  const scope = scopeQuery();
  try {
    const payload = await fetchJson(`/api/workbooks?${scope}`);
    applyScopePayload(payload);
    state.workbooks = Array.isArray(payload.workbooks) ? payload.workbooks : [];
    state.mainWorkbooks = Array.isArray(payload.main_workbooks) ? payload.main_workbooks : state.workbooks;
    state.referenceWorkbooks = Array.isArray(payload.reference_workbooks) ? payload.reference_workbooks : state.workbooks;

    const hasMainChoices = Array.isArray(state.mainWorkbooks) && state.mainWorkbooks.length > 0;
    const hasReferenceChoices = Array.isArray(state.referenceWorkbooks) && state.referenceWorkbooks.length > 0;
    if (!hasMainChoices || !hasReferenceChoices) {
      state.selectedMainWorkbook = null;
      state.selectedReferenceWorkbook = null;
      populateWorkbookSelect(el.mainWorkbookSelect, null, []);
      populateWorkbookSelect(el.referenceWorkbookSelect, null, []);
      updateWorkbookLabels();
      if (!hasMainChoices && !hasReferenceChoices) {
        setEmptyMain("No workbook available in current scope.");
        setEmptyRef("No workbook available in current scope.");
      } else if (!hasMainChoices) {
        setEmptyMain("No workbook flagged for Main view in current scope.");
        setEmptyRef("Choose a main-view workbook in Files tab.");
      } else {
        setEmptyMain("Choose a detail-view workbook in Files tab.");
        setEmptyRef("No workbook flagged for Detail view in current scope.");
      }
      return;
    }

    state.selectedMainWorkbook = payload.default_main || state.mainWorkbooks[0];
    state.selectedReferenceWorkbook = payload.default_reference || state.referenceWorkbooks[0];

    populateWorkbookSelect(el.mainWorkbookSelect, state.selectedMainWorkbook, state.mainWorkbooks);
    populateWorkbookSelect(
      el.referenceWorkbookSelect,
      state.selectedReferenceWorkbook,
      state.referenceWorkbooks,
    );
    state.selectedMainWorkbook = el.mainWorkbookSelect ? el.mainWorkbookSelect.value || state.selectedMainWorkbook : state.selectedMainWorkbook;
    state.selectedReferenceWorkbook = el.referenceWorkbookSelect
      ? el.referenceWorkbookSelect.value || state.selectedReferenceWorkbook
      : state.selectedReferenceWorkbook;
    updateWorkbookLabels();
  } finally {
    endBusy();
  }
}

async function uploadSelectedWorkbooks() {
  if (!state.permissions.can_upload) {
    setText(el.statusText, "Current user does not have upload permission.");
    return;
  }

  const primaryFiles = Array.from((el.uploadInput && el.uploadInput.files) || []);
  const ribbonFiles = Array.from((el.ribbonUploadInput && el.ribbonUploadInput.files) || []);
  const files = [...primaryFiles, ...ribbonFiles];
  if (!files.length) {
    setText(el.statusText, "Choose one or more Excel files first.");
    return;
  }

  const formData = new FormData();
  for (const file of files) {
    formData.append("files", file);
  }
  formData.append("user", state.currentUser || "owner");
  formData.append("region", state.selectedRegion || "ALL");
  const uploadRegionValue = (
    (el.ribbonUploadRegionInput ? el.ribbonUploadRegionInput.value : "") ||
    (el.uploadRegionInput ? el.uploadRegionInput.value : "")
  ).trim();
  if (uploadRegionValue) {
    formData.append("upload_region", uploadRegionValue);
  } else if (state.selectedRegion && state.selectedRegion !== "ALL") {
    formData.append("upload_region", state.selectedRegion);
  }

  beginBusy(`Excel file ${files.length} ခု upload လုပ်နေသည်...`);
  try {
    setText(el.statusText, `Uploading ${files.length} file(s)...`);
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
    applyScopePayload(payload);
    state.workbooks = Array.isArray(payload.workbooks) ? payload.workbooks : state.workbooks;
    state.mainWorkbooks = Array.isArray(payload.main_workbooks) ? payload.main_workbooks : state.workbooks;
    state.referenceWorkbooks = Array.isArray(payload.reference_workbooks)
      ? payload.reference_workbooks
      : state.workbooks;
    state.selectedMainWorkbook = payload.default_main || state.selectedMainWorkbook || state.mainWorkbooks[0] || null;
    state.selectedReferenceWorkbook =
      payload.default_reference || state.selectedReferenceWorkbook || state.referenceWorkbooks[0] || null;
    state.cache.clear();
    state.mainStyledRequestKey = null;
    state.mainAvailableMonths.clear();

    populateWorkbookSelect(el.mainWorkbookSelect, state.selectedMainWorkbook, state.mainWorkbooks);
    populateWorkbookSelect(
      el.referenceWorkbookSelect,
      state.selectedReferenceWorkbook,
      state.referenceWorkbooks,
    );
    state.selectedMainWorkbook = el.mainWorkbookSelect ? el.mainWorkbookSelect.value || state.selectedMainWorkbook : state.selectedMainWorkbook;
    state.selectedReferenceWorkbook = el.referenceWorkbookSelect
      ? el.referenceWorkbookSelect.value || state.selectedReferenceWorkbook
      : state.selectedReferenceWorkbook;
    updateWorkbookLabels();
    await loadSheets();
    await refreshAccessAndFiles();

    if (el.uploadInput) {
      el.uploadInput.value = "";
    }
    if (el.ribbonUploadInput) {
      el.ribbonUploadInput.value = "";
    }
    if (el.uploadRegionInput) {
      el.uploadRegionInput.value = "";
    }
    if (el.ribbonUploadRegionInput) {
      el.ribbonUploadRegionInput.value = "";
    }
    const skippedCount = Array.isArray(payload.skipped_files) ? payload.skipped_files.length : 0;
    const skippedText = skippedCount ? ` · skipped ${skippedCount} unsupported file(s)` : "";
    const uploadedRegions = Array.isArray(payload.uploaded_regions)
      ? [...new Set(payload.uploaded_regions.map((item) => (item && item.region ? String(item.region) : "")))]
          .filter(Boolean)
          .join(", ")
      : "";
    const uploadRegionText = uploadedRegions || regionLabel(state.selectedRegion);
    setText(el.statusText, `Uploaded ${files.length} file(s) to ${uploadRegionText}${skippedText}`);
    syncRibbonFromCore();
  } finally {
    endBusy();
  }
}

function applySheetsPayload(payload) {
  const previousVersion = state.pairVersion;
  applyScopePayload(payload);
  state.mainSheetTabs = sanitizeSheetTabs(payload.main_sheet_tabs, payload.main_sheet_names);
  state.referenceSheetTabs = sanitizeSheetTabs(
    payload.reference_sheet_tabs,
    payload.reference_sheet_names,
  );
  state.pairVersion = payload.version || null;
  state.workbooks = payload.available_workbooks || state.workbooks;
  state.mainWorkbooks = Array.isArray(payload.main_workbooks)
    ? payload.main_workbooks
    : state.mainWorkbooks.length
      ? state.mainWorkbooks
      : state.workbooks;
  state.referenceWorkbooks = Array.isArray(payload.reference_workbooks)
    ? payload.reference_workbooks
    : state.referenceWorkbooks.length
      ? state.referenceWorkbooks
      : state.workbooks;

  if (payload.main_workbook) {
    state.selectedMainWorkbook = payload.main_workbook;
  }
  if (payload.reference_workbook) {
    state.selectedReferenceWorkbook = payload.reference_workbook;
  }

  populateWorkbookSelect(el.mainWorkbookSelect, state.selectedMainWorkbook, state.mainWorkbooks);
  populateWorkbookSelect(
    el.referenceWorkbookSelect,
    state.selectedReferenceWorkbook,
    state.referenceWorkbooks,
  );
  state.selectedMainWorkbook = el.mainWorkbookSelect ? el.mainWorkbookSelect.value || state.selectedMainWorkbook : state.selectedMainWorkbook;
  state.selectedReferenceWorkbook = el.referenceWorkbookSelect
    ? el.referenceWorkbookSelect.value || state.selectedReferenceWorkbook
    : state.selectedReferenceWorkbook;
  updateWorkbookLabels();
  const changed = previousVersion !== state.pairVersion;
  if (changed) {
    state.mainStyledRequestKey = null;
    state.mainAvailableMonths.clear();
    state.collapsedRowGroups.clear();
  }
  state.lastLoadAt = new Date();
  updateLoadHealth();
  return changed;
}

async function loadSheets() {
  if (!state.selectedMainWorkbook || !state.selectedReferenceWorkbook) {
    state.mainSheetTabs = [];
    state.referenceSheetTabs = [];
    rebuildMainTabs(null);
    rebuildReferenceTabs(null);
    updateLoadHealth("Status: No workbook selected for this scope.");
    return;
  }
  beginBusy("Workbook tabs နဲ့ sheets များကို load လုပ်နေသည်...");
  try {
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
  } finally {
    endBusy();
  }
}

async function ensureSheetLoaded(canonical) {
  const key = cacheKey(canonical);
  if (state.cache.has(key)) {
    return state.cache.get(key);
  }

  setText(el.statusText, `Loading ${canonical}...`);
  const query = workbookQuery();
  const payload = await fetchJson(`/api/sheet/${encodeURIComponent(canonical)}?${query}`);
  state.cache.set(key, payload);
  setText(el.statusText, "");
  return payload;
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

function readSelectedMonthValues(monthSelect) {
  if (!monthSelect) {
    return [];
  }
  const selectedValues = Array.from(monthSelect.selectedOptions || [])
    .map((option) => Number.parseInt(option.value, 10))
    .filter((value) => Number.isFinite(value) && value >= 1 && value <= 12);
  return [...new Set(selectedValues)].sort((a, b) => a - b);
}

function writeSelectedMonthValues(monthSelect, monthValues) {
  if (!monthSelect) {
    return [];
  }
  const normalized = [...new Set((Array.isArray(monthValues) ? monthValues : [])
    .map((value) => Number.parseInt(String(value), 10))
    .filter((value) => Number.isFinite(value) && value >= 1 && value <= 12))].sort((a, b) => a - b);
  const selectedSet = new Set(normalized.map((value) => String(value)));
  for (const option of Array.from(monthSelect.options || [])) {
    option.selected = selectedSet.has(option.value);
  }
  return normalized;
}

function monthSelectionSignature(modeValue, monthSelect) {
  if (modeIsSameMonthYears(modeValue)) {
    return monthSelect ? monthSelect.value || "auto" : "auto";
  }
  if (modeIsMultiMonthYears(modeValue)) {
    return readSelectedMonthValues(monthSelect).join(",") || "auto";
  }
  return "none";
}

function selectedMonthsForMode(modeValue, monthSelect) {
  if (!monthSelect) {
    return [];
  }
  if (modeIsSameMonthYears(modeValue)) {
    const monthValue = Number.parseInt(monthSelect.value, 10);
    return Number.isFinite(monthValue) && monthValue >= 1 && monthValue <= 12 ? [monthValue] : [];
  }
  if (modeIsMultiMonthYears(modeValue)) {
    return readSelectedMonthValues(monthSelect);
  }
  return [];
}

function updateMonthControl(sheetData, view) {
  const controls = filterElementsForView(view);
  const modeValue = controls.modeSelect.value;
  const previousSignature = monthSelectionSignature(modeValue, controls.monthSelect);

  if (!modeUsesMonthSelector(modeValue)) {
    controls.monthWrapper.classList.add("hidden");
    controls.monthSelect.multiple = false;
    controls.monthSelect.removeAttribute("size");
    syncRibbonFromCore();
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
    controls.monthSelect.multiple = false;
    controls.monthSelect.removeAttribute("size");
    syncRibbonFromCore();
    return false;
  }

  const isMultiMode = modeIsMultiMonthYears(modeValue);
  const previousMultiValues = readSelectedMonthValues(controls.monthSelect);
  const storedMultiValues = String(controls.monthSelect.dataset.multi || "")
    .split(",")
    .map((token) => Number.parseInt(token, 10))
    .filter((value) => Number.isFinite(value) && value >= 1 && value <= 12);
  controls.monthWrapper.classList.remove("hidden");
  controls.monthSelect.multiple = isMultiMode;
  if (isMultiMode) {
    // Keep ribbon month list compact in fixed-height panels.
    controls.monthSelect.setAttribute("size", String(Math.min(Math.max(availableMonths.length, 3), 4)));
  } else {
    controls.monthSelect.removeAttribute("size");
  }
  controls.monthSelect.innerHTML = "";

  for (const monthIndex of availableMonths) {
    const option = document.createElement("option");
    option.value = String(monthIndex);
    option.textContent = monthLabels[monthIndex - 1];
    controls.monthSelect.appendChild(option);
  }

  if (isMultiMode) {
    const preferredMulti = [...new Set([...previousMultiValues, ...storedMultiValues])].filter((monthValue) =>
      availableMonths.includes(monthValue),
    );
    const fallbackMonth = availableMonths[availableMonths.length - 1];
    const normalizedMulti = writeSelectedMonthValues(
      controls.monthSelect,
      preferredMulti.length ? preferredMulti : [fallbackMonth],
    );
    controls.monthSelect.dataset.multi = normalizedMulti.join(",");
    controls.monthSelect.dataset.current = normalizedMulti.length ? String(normalizedMulti[normalizedMulti.length - 1]) : "";
  } else {
    const current = Number.parseInt(controls.monthSelect.dataset.current || "", 10);
    const preferred = availableMonths.includes(current)
      ? current
      : availableMonths[availableMonths.length - 1];
    controls.monthSelect.value = String(preferred);
    controls.monthSelect.dataset.current = String(preferred);
    controls.monthSelect.dataset.multi = String(preferred);
  }

  controls.monthSelect.title = isMultiMode
    ? "Select one or more months. Hold Ctrl/Cmd for non-adjacent selections."
    : "";
  syncRibbonFromCore();
  return monthSelectionSignature(modeValue, controls.monthSelect) !== previousSignature;
}

function pickMonths(sheetData, view) {
  const sorted = [...sheetData.months].sort((a, b) => a.key.localeCompare(b.key));
  const controls = filterElementsForView(view);
  const n = currentN(controls.nInput);
  const coverage = monthCoverageMap(sheetData);
  const minRowsForPopulated = Math.max(2, Math.ceil(sheetData.rows.length * 0.05));
  const modeValue = controls.modeSelect.value;

  if (!modeUsesMonthSelector(modeValue)) {
    const populated = sorted.filter((month) => (coverage.get(month.key) || 0) >= minRowsForPopulated);
    const fallback = sorted.filter((month) => (coverage.get(month.key) || 0) > 0);
    const source = populated.length ? populated : fallback.length ? fallback : sorted;
    return source.slice(-n);
  }

  const monthSelections = selectedMonthsForMode(modeValue, controls.monthSelect);
  if (!monthSelections.length) {
    return sorted.slice(-n);
  }

  if (modeIsSameMonthYears(modeValue)) {
    const selectedMonth = monthSelections[monthSelections.length - 1];
    const filtered = sorted.filter((item) => item.month === selectedMonth);
    return filtered.slice(-n);
  }

  const picked = [];
  for (const monthSelection of monthSelections) {
    const filtered = sorted.filter((item) => item.month === monthSelection);
    picked.push(...filtered.slice(-n));
  }
  const byKey = new Map();
  for (const item of picked) {
    byKey.set(item.key, item);
  }
  return [...byKey.values()].sort((a, b) => a.key.localeCompare(b.key));
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

function tableSelectionColOffset(table) {
  return Math.max(0, Number.parseInt(String(table && table.dataset ? table.dataset.selectionColOffset || "0" : "0"), 10) || 0);
}

function ensureColgroupColumns(table, colCount) {
  if (!table || !Number.isFinite(colCount) || colCount < 1) {
    return null;
  }
  let colgroup = table.querySelector(":scope > colgroup");
  if (!colgroup) {
    colgroup = document.createElement("colgroup");
    table.insertBefore(colgroup, table.firstChild);
  }
  const cols = Array.from(colgroup.children || []);
  if (cols.length > colCount) {
    for (let idx = cols.length - 1; idx >= colCount; idx -= 1) {
      cols[idx].remove();
    }
  } else if (cols.length < colCount) {
    for (let idx = cols.length; idx < colCount; idx += 1) {
      colgroup.appendChild(document.createElement("col"));
    }
  }
  return colgroup;
}

function applyColumnWidthOverridesToTable(table, scope) {
  if (!table) {
    return 0;
  }
  const normalizedScope = normalizeViewScope(scope);
  const viewLayout = viewLayoutForScope(normalizedScope);
  viewLayout.columnWidths = normalizeColumnWidthMap(viewLayout.columnWidths);
  const totalCols = annotateCellGrid(table);
  if (!totalCols) {
    return 0;
  }

  const colgroup = ensureColgroupColumns(table, totalCols);
  if (!colgroup) {
    return totalCols;
  }
  const offset = tableSelectionColOffset(table);
  const colElements = Array.from(colgroup.children || []);
  for (let idx = 0; idx < colElements.length; idx += 1) {
    const globalIndex = offset + idx;
    const width = normalizeColumnWidthPx(viewLayout.columnWidths[String(globalIndex)]);
    if (Number.isFinite(width)) {
      colElements[idx].style.width = `${width}px`;
      colElements[idx].style.minWidth = `${width}px`;
      colElements[idx].style.maxWidth = `${width}px`;
    } else {
      colElements[idx].style.removeProperty("width");
      colElements[idx].style.removeProperty("min-width");
      colElements[idx].style.removeProperty("max-width");
    }
  }
  return totalCols;
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

function shouldSplitFrozenViewport(table, frozenCount, widths, totalCols, frozenRows) {
  if (Number.isFinite(frozenRows) && frozenRows > 0) {
    return false;
  }
  const hasMergedCells = Array.from(table.querySelectorAll("th, td")).some((cell) => {
    const rowSpan = Number.parseInt(cell.getAttribute("rowspan") || "1", 10);
    const colSpan = Number.parseInt(cell.getAttribute("colspan") || "1", 10);
    return rowSpan > 1 || colSpan > 1;
  });
  if (hasMergedCells) {
    return false;
  }
  if (frozenCount < 2 || frozenCount >= totalCols) {
    return false;
  }

  const frozenWidth = widths.slice(0, frozenCount).reduce((sum, width) => sum + width, 0);
  const totalWidth = widths.reduce((sum, width) => sum + width, 0);
  const wrap = tableWrapFor(table);
  const viewportWidth = wrap ? wrap.clientWidth : table.parentElement ? table.parentElement.clientWidth : 0;

  if (viewportWidth > 0 && viewportWidth <= 1024) {
    return true;
  }

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

function clearSplitRowHeights(table) {
  if (!table) {
    return;
  }
  for (const row of Array.from(table.rows || [])) {
    restoreExcelRowHeight(row);
    for (const cell of Array.from(row.cells || [])) {
      cell.style.removeProperty("height");
      cell.style.removeProperty("min-height");
      cell.style.removeProperty("box-sizing");
    }
  }
}

function syncSplitRowHeights(leftTable, rightTable) {
  if (!leftTable || !rightTable) {
    return;
  }

  const leftRows = Array.from(leftTable.rows);
  const rightRows = Array.from(rightTable.rows);
  if (!leftRows.length || !rightRows.length) {
    return;
  }

  clearSplitRowHeights(leftTable);
  clearSplitRowHeights(rightTable);

  const sharedRowCount = Math.min(leftRows.length, rightRows.length);
  for (let idx = 0; idx < sharedRowCount; idx += 1) {
    const leftHeight = Math.ceil(leftRows[idx].getBoundingClientRect().height);
    const rightHeight = Math.ceil(rightRows[idx].getBoundingClientRect().height);
    const syncedHeight = Math.max(leftHeight, rightHeight);
    if (syncedHeight > 0) {
      const heightValue = `${syncedHeight}px`;
      leftRows[idx].style.height = heightValue;
      rightRows[idx].style.height = heightValue;
      for (const cell of Array.from(leftRows[idx].cells || [])) {
        cell.style.boxSizing = "border-box";
        cell.style.minHeight = heightValue;
        cell.style.height = heightValue;
      }
      for (const cell of Array.from(rightRows[idx].cells || [])) {
        cell.style.boxSizing = "border-box";
        cell.style.minHeight = heightValue;
        cell.style.height = heightValue;
      }
    }
  }
}

function detectSplitHeaderRowCount(table) {
  const rows = Array.from(table.rows || []);
  if (!rows.length) {
    return 0;
  }

  const scanLimit = Math.min(rows.length, 6);
  let headerRows = 0;
  for (let idx = 0; idx < scanLimit; idx += 1) {
    const row = rows[idx];
    const cells = Array.from(row.cells || []);
    if (!cells.length) {
      headerRows += 1;
      continue;
    }

    const hasMergedCells = cells.some((cell) => {
      const rowSpan = Number.parseInt(cell.getAttribute("rowspan") || "1", 10);
      const colSpan = Number.parseInt(cell.getAttribute("colspan") || "1", 10);
      return rowSpan > 1 || colSpan > 1;
    });
    const firstText = cells[0] ? cells[0].textContent || "" : "";
    if (idx > 0 && isNumericCellText(firstText) && !hasMergedCells) {
      break;
    }

    headerRows += 1;
  }

  return Math.max(1, Math.min(headerRows, 4));
}

function applySplitHeaderLayering(leftTable, rightTable) {
  if (!leftTable || !rightTable) {
    return;
  }

  const headerRows = detectSplitHeaderRowCount(rightTable);
  if (!headerRows) {
    return;
  }

  for (const table of [leftTable, rightTable]) {
    for (const cell of table.querySelectorAll("th, td")) {
      if (!cell.classList.contains("split-head-cell")) {
        continue;
      }
      cell.classList.remove("split-head-cell");
      cell.style.removeProperty("position");
      cell.style.removeProperty("z-index");
      cell.style.removeProperty("background-color");
    }

    const rows = Array.from(table.rows || []);
    const cappedRows = Math.min(headerRows, rows.length);
    for (let idx = 0; idx < cappedRows; idx += 1) {
      const z = 60 - idx;
      for (const cell of Array.from(rows[idx].cells || [])) {
        cell.classList.add("split-head-cell");
        cell.style.position = "relative";
        cell.style.zIndex = String(z);
        cell.style.backgroundColor = resolveOpaqueStickyBackground(cell, table);
      }
    }
  }
}

function applySplitViewport(table, frozenCount, totalCols, widths) {
  const host = table.parentElement;
  if (!host) {
    return false;
  }
  const wrap = tableWrapFor(table);
  const viewportWidth = wrap ? wrap.clientWidth : host.clientWidth;
  const frozenWidth = widths.slice(0, frozenCount).reduce((sum, width) => sum + width, 0);
  if (viewportWidth > 0 && frozenWidth > viewportWidth * 0.72) {
    return false;
  }

  const leftTable = table.cloneNode(true);
  clearStickyStyles(leftTable);
  clearStickyStyles(table);
  leftTable.dataset.selectionColOffset = "0";
  table.dataset.selectionColOffset = String(frozenCount);

  pruneTableColumns(leftTable, 0, frozenCount);
  pruneTableColumns(table, frozenCount, totalCols);

  const splitView = document.createElement("div");
  splitView.className = "split-view";
  splitView.style.setProperty("--split-left-width", `${Math.max(180, Math.ceil(frozenWidth))}px`);

  const leftPane = document.createElement("div");
  leftPane.className = "split-pane split-pane-left";
  leftPane.appendChild(leftTable);

  const rightPane = document.createElement("div");
  rightPane.className = "split-pane split-pane-right";

  splitView.appendChild(leftPane);
  splitView.appendChild(rightPane);
  if (table.parentElement !== host) {
    return false;
  }
  try {
    host.replaceChild(splitView, table);
  } catch {
    return false;
  }
  rightPane.appendChild(table);
  leftPane.scrollLeft = 0;

  let splitDisposed = false;
  const syncRows = () => {
    if (splitDisposed || !splitView.isConnected) {
      return;
    }
    syncSplitRowHeights(leftTable, table);
    applySplitHeaderLayering(leftTable, table);
  };
  const syncTimeouts = [];
  const queueDelayedSync = (delayMs) => {
    const handle = window.setTimeout(() => {
      syncRows();
    }, delayMs);
    syncTimeouts.push(handle);
  };
  syncRows();
  if (typeof requestAnimationFrame === "function") {
    requestAnimationFrame(() => {
      syncRows();
      requestAnimationFrame(syncRows);
    });
  }
  queueDelayedSync(60);
  queueDelayedSync(220);
  queueDelayedSync(560);
  if (document.fonts && document.fonts.ready && typeof document.fonts.ready.then === "function") {
    document.fonts.ready.then(() => {
      syncRows();
    }).catch(() => {});
  }
  if (typeof ResizeObserver !== "undefined") {
    const observer = new ResizeObserver(() => {
      syncRows();
    });
    observer.observe(leftPane);
    observer.observe(rightPane);
    observer.observe(leftTable);
    observer.observe(table);
    splitView.__rowSyncObserver = observer;
  }
  let tableMutationObserver = null;
  if (typeof MutationObserver !== "undefined") {
    tableMutationObserver = new MutationObserver(() => {
      syncRows();
    });
    tableMutationObserver.observe(leftTable, { childList: true, subtree: true, characterData: true });
    tableMutationObserver.observe(table, { childList: true, subtree: true, characterData: true });
  }

  if (wrap) {
    wrap.classList.add("table-wrap-split");
    const height = wrap.clientHeight;
    if (height > 0) {
      const maxHeight = `${height}px`;
      leftPane.style.maxHeight = maxHeight;
      rightPane.style.maxHeight = maxHeight;
    }
  }
  leftPane.style.minWidth = "0";
  rightPane.style.minWidth = "0";
  const onWindowResize = () => {
    syncRows();
  };
  window.addEventListener("resize", onWindowResize);
  splitView.__splitCleanup = () => {
    splitDisposed = true;
    window.removeEventListener("resize", onWindowResize);
    for (const handle of syncTimeouts) {
      window.clearTimeout(handle);
    }
    if (tableMutationObserver && typeof tableMutationObserver.disconnect === "function") {
      tableMutationObserver.disconnect();
    }
  };

  let syncLock = false;
  leftPane.addEventListener(
    "scroll",
    () => {
      if (leftPane.scrollLeft !== 0) {
        leftPane.scrollLeft = 0;
      }
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

function enhanceFrozenViewport(scope, table, frozenCount, frozenRows) {
  if (!table) {
    return;
  }

  const normalizedTable = unwrapSplitViewport(table);
  normalizedTable.dataset.selectionColOffset = "0";
  const wrap = tableWrapFor(normalizedTable);
  if (wrap) {
    wrap.classList.remove("table-wrap-split");
  }
  applyColumnWidthOverridesToTable(normalizedTable, scope);

  const hasFrozenCols = Number.isFinite(frozenCount) && frozenCount > 0;
  const hasFrozenRows = Number.isFinite(frozenRows) && frozenRows > 0;
  if (!hasFrozenCols && !hasFrozenRows) {
    clearStickyStyles(normalizedTable);
    return;
  }
  clearStickyStyles(normalizedTable);

  const totalCols = annotateCellGrid(normalizedTable);
  const totalRows = Array.from(normalizedTable.rows || []).length;
  if (!totalCols || !totalRows) {
    return;
  }

  const clampedFrozenCols = hasFrozenCols ? Math.min(Math.max(1, frozenCount), totalCols) : 0;
  const clampedFrozenRows = hasFrozenRows ? Math.min(Math.max(1, frozenRows), totalRows) : 0;
  const widths = measureColumnWidths(normalizedTable, totalCols);
  if (shouldSplitFrozenViewport(normalizedTable, clampedFrozenCols, widths, totalCols, clampedFrozenRows)) {
    const applied = applySplitViewport(normalizedTable, clampedFrozenCols, totalCols, widths);
    if (applied) {
      return;
    }
  }

  applyFrozenColumns(normalizedTable, clampedFrozenCols);
  applyFrozenRows(normalizedTable, clampedFrozenRows);
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

function cssVarValue(style, name) {
  if (!style || !name) {
    return "";
  }
  return style.getPropertyValue(name).trim();
}

function resolveOpaqueStickyBackground(cell, table) {
  const rootStyle = getComputedStyle(document.documentElement);
  const bodyStyle = document.body ? getComputedStyle(document.body) : null;
  const tableStyle = table ? getComputedStyle(table) : null;
  const cellStyle = cell ? getComputedStyle(cell) : null;

  const tableBodyBg =
    cssVarValue(tableStyle, "--theme-table-bg") ||
    cssVarValue(bodyStyle, "--theme-table-bg") ||
    cssVarValue(rootStyle, "--theme-table-bg");
  const tableHeadBg =
    cssVarValue(tableStyle, "--theme-table-head-bg") ||
    cssVarValue(bodyStyle, "--theme-table-head-bg") ||
    cssVarValue(rootStyle, "--theme-table-head-bg");
  const stickyBg =
    cssVarValue(tableStyle, "--sticky-bg") ||
    cssVarValue(bodyStyle, "--sticky-bg") ||
    cssVarValue(rootStyle, "--sticky-bg");
  const softBg = cssVarValue(bodyStyle, "--bg-soft") || cssVarValue(rootStyle, "--bg-soft");
  const tableBgColor = tableStyle ? tableStyle.backgroundColor : "";

  const baseColor =
    parseCssColor(tableBodyBg) ||
    parseCssColor(tableBgColor) ||
    parseCssColor(stickyBg) ||
    parseCssColor(softBg) || { r: 245, g: 245, b: 245, a: 1 };
  const opaqueBase =
    baseColor.a >= 1 ? baseColor : blendColors(baseColor, { r: 245, g: 245, b: 245, a: 1 });

  if (cellStyle) {
    const cellColor = parseCssColor(cellStyle.backgroundColor);
    if (cellColor && cellColor.a > 0) {
      if (cellColor.a >= 1) {
        return toSolidRgbString(cellColor);
      }
      return toSolidRgbString(blendColors(cellColor, opaqueBase));
    }
  }

  const isHeaderCell = Boolean(cell && cell.tagName === "TH");
  const fallbackColor =
    parseCssColor(isHeaderCell ? tableHeadBg : tableBodyBg) ||
    parseCssColor(stickyBg) ||
    parseCssColor(tableBgColor) ||
    parseCssColor(softBg);
  if (fallbackColor) {
    if (fallbackColor.a >= 1) {
      return toSolidRgbString(fallbackColor);
    }
    return toSolidRgbString(blendColors(fallbackColor, opaqueBase));
  }

  return toSolidRgbString(opaqueBase);
}

function refreshFrozenSurfaceColors(scope = null) {
  const scopes = scope ? [normalizeViewScope(scope)] : ["main", "reference"];
  for (const scopeName of scopes) {
    for (const table of tablesForScope(scopeName)) {
      for (const cell of table.querySelectorAll("th, td")) {
        if (
          !cell.classList.contains("sticky-col") &&
          !cell.classList.contains("sticky-row") &&
          !cell.classList.contains("split-head-cell")
        ) {
          continue;
        }
        cell.style.backgroundColor = resolveOpaqueStickyBackground(cell, table);
      }
    }
  }
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

function applyFrozenRows(table, frozenRows) {
  if (!table) {
    return;
  }

  const allCells = Array.from(table.querySelectorAll("th, td"));
  for (const cell of allCells) {
    cell.classList.remove("sticky-row", "sticky-row-boundary");
    cell.style.removeProperty("top");
    if (!cell.classList.contains("sticky-col")) {
      cell.style.removeProperty("z-index");
    }
  }

  if (!Number.isFinite(frozenRows) || frozenRows < 1) {
    return;
  }

  const rows = Array.from(table.rows || []);
  const clampedRows = Math.min(frozenRows, rows.length);
  if (clampedRows < 1) {
    return;
  }

  let runningTop = 0;
  for (let idx = 0; idx < clampedRows; idx += 1) {
    const row = rows[idx];
    const rowHeight = effectiveRowHeightPx(row);
    for (const cell of Array.from(row.cells || [])) {
      cell.classList.add("sticky-row");
      cell.style.top = `${runningTop}px`;
      const baseZ = cell.classList.contains("sticky-col") ? 28 : 18;
      cell.style.zIndex = String(baseZ - idx);
      cell.style.backgroundColor = resolveOpaqueStickyBackground(cell, table);
      if (idx === clampedRows - 1) {
        cell.classList.add("sticky-row-boundary");
      }
    }
    runningTop += rowHeight;
  }
}

function workbookNameForScope(scope) {
  const normalized = normalizeViewScope(scope);
  if (normalized === "reference") {
    return state.selectedReferenceWorkbook || "Detail file";
  }
  return state.selectedMainWorkbook || "Main file";
}

function hasGridSelection() {
  return Boolean(state.selectionScope && Array.isArray(state.selectionRanges) && state.selectionRanges.length);
}

function statusScopeLabel(scope, hasSelection) {
  const normalized = normalizeViewScope(scope);
  const phase = hasSelection ? "Selection" : "View";
  const view = normalized === "reference" ? "Detail" : "Main";
  return `${view} ${phase} · ${workbookNameForScope(normalized)}`;
}

function refreshStatusSelectionScopeLabel() {
  if (!el.statusSelectionScope) {
    return;
  }
  const hasSelection = hasGridSelection();
  const scope = hasSelection ? state.selectionScope : state.activeViewScope;
  setText(el.statusSelectionScope, statusScopeLabel(scope, hasSelection));
}

function setActiveViewScope(scope, options = {}) {
  state.activeViewScope = normalizeViewScope(scope);
  syncZoomControlForScope(state.activeViewScope);
  if (options.refresh !== false && !hasGridSelection()) {
    refreshStatusSelectionScopeLabel();
  }
  syncRibbonGridControlState();
}

function tableWrapForScope(scope) {
  if (scope === "main") {
    return el.mainTableWrap;
  }
  if (scope === "reference") {
    return el.referenceTableWrap;
  }
  return null;
}

function tablesForScope(scope) {
  const wrap = tableWrapForScope(scope);
  if (!wrap) {
    return [];
  }
  return Array.from(wrap.querySelectorAll("table"));
}

function viewLayoutForScope(scope) {
  const normalized = normalizeViewScope(scope);
  if (!state.viewLayoutOverrides || typeof state.viewLayoutOverrides !== "object") {
    state.viewLayoutOverrides = {};
  }
  if (!state.viewLayoutOverrides[normalized]) {
    state.viewLayoutOverrides[normalized] = {
      frozenColsOverride: null,
      frozenRowsOverride: null,
      lastAppliedFrozenCols: 0,
      lastAppliedFrozenRows: 0,
      columnWidths: {},
      hiddenRows: new Set(),
      hiddenCols: new Set(),
    };
  }
  const viewLayout = state.viewLayoutOverrides[normalized];
  if (!Number.isFinite(Number(viewLayout.lastAppliedFrozenCols))) {
    viewLayout.lastAppliedFrozenCols = 0;
  }
  if (!Number.isFinite(Number(viewLayout.lastAppliedFrozenRows))) {
    viewLayout.lastAppliedFrozenRows = 0;
  }
  if (!(viewLayout.hiddenRows instanceof Set)) {
    viewLayout.hiddenRows = new Set();
  }
  if (!(viewLayout.hiddenCols instanceof Set)) {
    viewLayout.hiddenCols = new Set();
  }
  viewLayout.columnWidths = normalizeColumnWidthMap(viewLayout.columnWidths);
  return viewLayout;
}

function effectiveFreezeForScope(scope, defaultFrozenCols = 0, defaultFrozenRows = 0) {
  const viewLayout = viewLayoutForScope(scope);
  const rawCols = Number.isFinite(viewLayout.frozenColsOverride) ? viewLayout.frozenColsOverride : defaultFrozenCols;
  const rawRows = Number.isFinite(viewLayout.frozenRowsOverride) ? viewLayout.frozenRowsOverride : defaultFrozenRows;
  return {
    cols: Math.max(0, Number.parseInt(String(rawCols || 0), 10) || 0),
    rows: Math.max(0, Number.parseInt(String(rawRows || 0), 10) || 0),
  };
}

function normalizedFreezeForTable(table, requestedCols, requestedRows, options = {}) {
  const totalCols = annotateCellGrid(table);
  const totalRows = Array.from(table.rows || []).length;

  let cols = Math.max(0, Number.parseInt(String(requestedCols || 0), 10) || 0);
  let rows = Math.max(0, Number.parseInt(String(requestedRows || 0), 10) || 0);

  if (totalCols <= 1) {
    cols = 0;
  } else if (cols >= totalCols) {
    cols = totalCols - 1;
  }

  if (totalRows <= 1) {
    rows = 0;
  } else if (rows >= totalRows) {
    rows = totalRows - 1;
  }

  if (options.autoAdjustRows !== false && rows > 0 && totalRows >= 8) {
    const detectedHeaderRows = Math.max(1, detectSplitHeaderRowCount(table));
    const adaptiveRowCap = Math.max(1, Math.min(MAX_AUTO_FROZEN_ROWS, detectedHeaderRows));
    if (rows > adaptiveRowCap && rows / totalRows >= AUTO_FROZEN_ROW_RATIO_THRESHOLD) {
      rows = adaptiveRowCap;
    }
  }

  if (options.autoAdjustCols !== false && cols > 0 && totalCols >= 5) {
    // Keep a small non-frozen area so frozen panes remain usable after column pruning.
    const softMaxCols = Math.max(1, totalCols - 2);
    if (cols > softMaxCols) {
      cols = softMaxCols;
    }
  }

  return { cols, rows };
}

function primaryTableForScope(scope) {
  const wrap = tableWrapForScope(scope);
  if (!wrap) {
    return null;
  }
  const splitRightTable = wrap.querySelector(".split-pane-right table");
  if (splitRightTable) {
    return splitRightTable;
  }
  return wrap.querySelector("table");
}

function isColumnRangeFullyHidden(hiddenCols, startCol, endCol) {
  if (!(hiddenCols instanceof Set) || hiddenCols.size === 0) {
    return false;
  }
  for (let col = startCol; col <= endCol; col += 1) {
    if (!hiddenCols.has(col)) {
      return false;
    }
  }
  return true;
}

function applyScopeVisibilityOverrides(scope) {
  const normalized = normalizeViewScope(scope);
  const viewLayout = viewLayoutForScope(normalized);
  const hiddenRows = viewLayout.hiddenRows;
  const hiddenCols = viewLayout.hiddenCols;

  for (const table of tablesForScope(normalized)) {
    annotateSelectionGridForTable(table);

    const rows = Array.from(table.rows || []);
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex];
      const shouldHideRow = hiddenRows.has(rowIndex);
      if (shouldHideRow) {
        row.dataset.contextHiddenRow = "true";
        row.style.display = "none";
      } else if (row.dataset.contextHiddenRow === "true") {
        delete row.dataset.contextHiddenRow;
        row.style.removeProperty("display");
      }
    }

    for (const cell of table.querySelectorAll("th, td")) {
      const colStart = Number.parseInt(cell.dataset.gridColStart || "", 10);
      const colEnd = Number.parseInt(cell.dataset.gridColEnd || "", 10);
      const hasColRange = Number.isFinite(colStart) && Number.isFinite(colEnd) && colStart >= 0 && colEnd >= colStart;
      const shouldHideCol = hasColRange && isColumnRangeFullyHidden(hiddenCols, colStart, colEnd);
      if (shouldHideCol) {
        cell.dataset.contextHiddenCol = "true";
        cell.style.display = "none";
      } else if (cell.dataset.contextHiddenCol === "true") {
        delete cell.dataset.contextHiddenCol;
        cell.style.removeProperty("display");
      }
    }
  }
}

function measuredGlobalColumnWidthsForScope(scope) {
  const normalized = normalizeViewScope(scope);
  const measured = {};
  for (const table of tablesForScope(normalized)) {
    const totalCols = annotateCellGrid(table);
    if (!totalCols) {
      continue;
    }
    const widths = measureColumnWidths(table, totalCols);
    const offset = tableSelectionColOffset(table);
    for (let idx = 0; idx < widths.length; idx += 1) {
      const globalIndex = offset + idx;
      const width = normalizeColumnWidthPx(widths[idx]);
      if (!Number.isFinite(width)) {
        continue;
      }
      measured[String(globalIndex)] = Math.max(width, measured[String(globalIndex)] || 0);
    }
  }
  return measured;
}

function applyLiveColumnWidthsForScope(scope) {
  const normalized = normalizeViewScope(scope);
  const viewLayout = viewLayoutForScope(normalized);
  for (const table of tablesForScope(normalized)) {
    applyColumnWidthOverridesToTable(table, normalized);
  }

  const wrap = tableWrapForScope(normalized);
  if (!wrap) {
    return;
  }
  const splitView = wrap.querySelector(".split-view");
  if (splitView) {
    const frozenCols = Math.max(0, Number.parseInt(String(viewLayout.lastAppliedFrozenCols || 0), 10) || 0);
    if (frozenCols > 0) {
      const measured = measuredGlobalColumnWidthsForScope(normalized);
      let frozenWidth = 0;
      for (let col = 0; col < frozenCols; col += 1) {
        const override = normalizeColumnWidthPx(viewLayout.columnWidths[String(col)]);
        const measuredWidth = normalizeColumnWidthPx(measured[String(col)]);
        frozenWidth += override || measuredWidth || 96;
      }
      splitView.style.setProperty("--split-left-width", `${Math.max(180, Math.ceil(frozenWidth))}px`);
    }
    const leftTable = wrap.querySelector(".split-pane-left table");
    const rightTable = wrap.querySelector(".split-pane-right table");
    if (leftTable && rightTable) {
      syncSplitRowHeights(leftTable, rightTable);
      applySplitHeaderLayering(leftTable, rightTable);
    }
    return;
  }

  const primary = primaryTableForScope(normalized);
  if (!primary) {
    return;
  }
  if (viewLayout.lastAppliedFrozenCols > 0) {
    applyFrozenColumns(primary, viewLayout.lastAppliedFrozenCols);
  }
  if (viewLayout.lastAppliedFrozenRows > 0) {
    applyFrozenRows(primary, viewLayout.lastAppliedFrozenRows);
  }
}

function collectResizableHeaderCells(table) {
  const cellsByLocalColumn = new Map();
  if (!table) {
    return cellsByLocalColumn;
  }
  annotateCellGrid(table);
  const rows = Array.from(table.rows || []);
  if (!rows.length) {
    return cellsByLocalColumn;
  }
  const headerRowCount = Math.max(1, detectSplitHeaderRowCount(table));
  const rowLimit = Math.min(rows.length, headerRowCount);
  for (let rowIndex = 0; rowIndex < rowLimit; rowIndex += 1) {
    const row = rows[rowIndex];
    for (const cell of Array.from(row.cells || [])) {
      const colStart = Number.parseInt(cell.dataset.colStart || "-1", 10);
      const colSpan = Math.max(1, Number.parseInt(cell.dataset.colSpan || "1", 10) || 1);
      if (!Number.isFinite(colStart) || colStart < 0 || colSpan !== 1) {
        continue;
      }
      cellsByLocalColumn.set(colStart, cell);
    }
  }
  return cellsByLocalColumn;
}

function endColumnResize(event) {
  if (!activeColumnResize) {
    return;
  }
  const session = activeColumnResize;
  activeColumnResize = null;
  document.body.classList.remove("column-resizing");
  window.removeEventListener("pointermove", onColumnResizeMove, true);
  window.removeEventListener("pointerup", endColumnResize, true);
  window.removeEventListener("pointercancel", endColumnResize, true);

  if (
    session.captureTarget &&
    Number.isFinite(session.pointerId) &&
    typeof session.captureTarget.releasePointerCapture === "function"
  ) {
    try {
      if (session.captureTarget.hasPointerCapture(session.pointerId)) {
        session.captureTarget.releasePointerCapture(session.pointerId);
      }
    } catch {
      // Ignore pointer capture release failures.
    }
  }

  const normalized = normalizeViewScope(session.scope);
  const viewLayout = viewLayoutForScope(normalized);
  applyScopeLayoutOverrides(
    normalized,
    viewLayout.lastAppliedFrozenCols || 0,
    viewLayout.lastAppliedFrozenRows || 0,
  );
  if (event && typeof event.preventDefault === "function") {
    event.preventDefault();
  }
}

function onColumnResizeMove(event) {
  if (!activeColumnResize) {
    return;
  }
  const session = activeColumnResize;
  const delta = Number.parseFloat(String(event.clientX || 0)) - session.startClientX;
  const nextWidth = normalizeColumnWidthPx(session.startWidth + delta);
  if (!Number.isFinite(nextWidth) || nextWidth === session.currentWidth) {
    event.preventDefault();
    return;
  }
  session.currentWidth = nextWidth;
  const normalized = normalizeViewScope(session.scope);
  const viewLayout = viewLayoutForScope(normalized);
  viewLayout.columnWidths[String(session.columnIndex)] = nextWidth;
  applyLiveColumnWidthsForScope(normalized);
  event.preventDefault();
}

function beginColumnResize(event, scope, columnIndex) {
  if (typeof PointerEvent === "undefined") {
    return;
  }
  if (!(event instanceof PointerEvent)) {
    return;
  }
  if (event.button !== 0 && event.pointerType !== "touch") {
    return;
  }

  const normalized = normalizeViewScope(scope);
  const table = primaryTableForScope(normalized);
  if (!table) {
    return;
  }
  const viewLayout = viewLayoutForScope(normalized);
  const measured = measuredGlobalColumnWidthsForScope(normalized);
  const widthFromLayout = normalizeColumnWidthPx(viewLayout.columnWidths[String(columnIndex)]);
  const widthFromMeasure = normalizeColumnWidthPx(measured[String(columnIndex)]);
  const initialWidth = widthFromLayout || widthFromMeasure || 96;
  viewLayout.columnWidths[String(columnIndex)] = initialWidth;

  activeColumnResize = {
    scope: normalized,
    columnIndex,
    startClientX: Number.parseFloat(String(event.clientX || 0)),
    startWidth: initialWidth,
    currentWidth: initialWidth,
    pointerId: Number.isFinite(event.pointerId) ? event.pointerId : null,
    captureTarget: event.currentTarget instanceof HTMLElement ? event.currentTarget : null,
  };

  if (
    activeColumnResize.captureTarget &&
    Number.isFinite(activeColumnResize.pointerId) &&
    typeof activeColumnResize.captureTarget.setPointerCapture === "function"
  ) {
    try {
      activeColumnResize.captureTarget.setPointerCapture(activeColumnResize.pointerId);
    } catch {
      // Ignore pointer capture failures.
    }
  }

  document.body.classList.add("column-resizing");
  window.addEventListener("pointermove", onColumnResizeMove, true);
  window.addEventListener("pointerup", endColumnResize, true);
  window.addEventListener("pointercancel", endColumnResize, true);
  event.preventDefault();
  event.stopPropagation();
}

function bindColumnResizersForScope(scope) {
  const normalized = normalizeViewScope(scope);
  for (const table of tablesForScope(normalized)) {
    for (const handle of Array.from(table.querySelectorAll(".col-resize-handle"))) {
      handle.remove();
    }
    for (const cell of Array.from(table.querySelectorAll(".col-resize-cell"))) {
      cell.classList.remove("col-resize-cell");
    }

    const byLocalCol = collectResizableHeaderCells(table);
    const offset = tableSelectionColOffset(table);
    for (const [localCol, cell] of byLocalCol.entries()) {
      if (!(cell instanceof HTMLElement)) {
        continue;
      }
      const globalCol = offset + localCol;
      const handle = document.createElement("button");
      handle.type = "button";
      handle.className = "col-resize-handle";
      handle.setAttribute("aria-label", `Resize column ${columnToExcelLabel(globalCol)}`);
      handle.dataset.resizeScope = normalized;
      handle.dataset.resizeCol = String(globalCol);
      handle.addEventListener("pointerdown", (event) => {
        beginColumnResize(event, normalized, globalCol);
      });
      cell.classList.add("col-resize-cell");
      cell.appendChild(handle);
    }
  }
}

function applyScopeLayoutOverrides(scope, defaultFrozenCols = 0, defaultFrozenRows = 0) {
  const normalized = normalizeViewScope(scope);
  const table = primaryTableForScope(normalized);
  if (!table) {
    return false;
  }
  const normalizedTable = unwrapSplitViewport(table);
  const viewLayout = viewLayoutForScope(normalized);
  const freeze = effectiveFreezeForScope(normalized, defaultFrozenCols, defaultFrozenRows);
  const freezeNormalized = normalizedFreezeForTable(normalizedTable, freeze.cols, freeze.rows, {
    autoAdjustCols: !Number.isFinite(viewLayout.frozenColsOverride),
    autoAdjustRows: !Number.isFinite(viewLayout.frozenRowsOverride),
  });
  enhanceFrozenViewport(normalized, normalizedTable, freezeNormalized.cols, freezeNormalized.rows);
  viewLayout.lastAppliedFrozenCols = freezeNormalized.cols;
  viewLayout.lastAppliedFrozenRows = freezeNormalized.rows;
  applyScopeVisibilityOverrides(normalized);
  if (state.selectionScope === normalized && Array.isArray(state.selectionRanges) && state.selectionRanges.length) {
    refreshGridSelectionVisuals();
  } else {
    annotateSelectionGridForScope(normalized);
    syncRibbonGridControlState();
  }
  bindColumnResizersForScope(normalized);
  return true;
}

function annotateSelectionGridForTable(table) {
  if (!table) {
    return;
  }
  const colOffset = Number.parseInt(table.dataset.selectionColOffset || "0", 10) || 0;
  const rows = Array.from(table.rows || []);
  const spanMap = [];

  let maxCols = colOffset;
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    for (let idx = 0; idx < spanMap.length; idx += 1) {
      if (spanMap[idx] > 0) {
        spanMap[idx] -= 1;
      }
    }

    let col = 0;
    for (const cell of Array.from(rows[rowIndex].cells || [])) {
      while (spanMap[col] > 0) {
        col += 1;
      }

      const colSpan = Math.max(1, Number.parseInt(cell.getAttribute("colspan") || "1", 10) || 1);
      const rowSpan = Math.max(1, Number.parseInt(cell.getAttribute("rowspan") || "1", 10) || 1);
      const colStart = colOffset + col;
      const colEnd = colStart + colSpan - 1;
      const rowEnd = rowIndex + rowSpan - 1;

      cell.dataset.gridRow = String(rowIndex);
      cell.dataset.gridRowEnd = String(rowEnd);
      cell.dataset.gridRowSpan = String(rowSpan);
      cell.dataset.gridColStart = String(colStart);
      cell.dataset.gridColEnd = String(colEnd);
      cell.dataset.gridColSpan = String(colSpan);
      table.classList.add("grid-selection-enabled");

      if (rowSpan > 1) {
        for (let offset = 0; offset < colSpan; offset += 1) {
          const target = col + offset;
          spanMap[target] = Math.max(spanMap[target] || 0, rowSpan - 1);
        }
      }

      col += colSpan;
      maxCols = Math.max(maxCols, colOffset + col);
    }
  }

  table.dataset.selectionRowCount = String(rows.length);
  table.dataset.selectionColCount = String(maxCols);
}

function annotateSelectionGridForScope(scope) {
  for (const table of tablesForScope(scope)) {
    annotateSelectionGridForTable(table);
  }
}

function normalizeSelectionRange(range) {
  return {
    rowStart: Math.min(range.rowStart, range.rowEnd),
    rowEnd: Math.max(range.rowStart, range.rowEnd),
    colStart: Math.min(range.colStart, range.colEnd),
    colEnd: Math.max(range.colStart, range.colEnd),
  };
}

function rangeFromPoints(a, b) {
  return normalizeSelectionRange({
    rowStart: a.row,
    rowEnd: b.row,
    colStart: a.col,
    colEnd: b.col,
  });
}

function selectionExtentForPoint(scope, point) {
  if (!scope || !point) {
    return null;
  }
  const row = Number.parseInt(String(point.row), 10);
  const col = Number.parseInt(String(point.col), 10);
  if (!Number.isFinite(row) || !Number.isFinite(col)) {
    return null;
  }

  const normalizedScope = normalizeViewScope(scope);
  for (const table of tablesForScope(normalizedScope)) {
    annotateSelectionGridForTable(table);
    const targetRow = table.rows && row >= 0 && row < table.rows.length ? table.rows[row] : null;
    if (!targetRow) {
      continue;
    }
    for (const cell of Array.from(targetRow.cells || [])) {
      const colStart = Number.parseInt(cell.dataset.gridColStart || "", 10);
      const colEnd = Number.parseInt(cell.dataset.gridColEnd || "", 10);
      if (!Number.isFinite(colStart) || !Number.isFinite(colEnd) || col < colStart || col > colEnd) {
        continue;
      }
      const rowStart = Number.parseInt(cell.dataset.gridRow || "", 10);
      const rowEnd = Number.parseInt(cell.dataset.gridRowEnd || cell.dataset.gridRow || "", 10);
      if (!Number.isFinite(rowStart) || !Number.isFinite(rowEnd)) {
        break;
      }
      return {
        rowStart,
        rowEnd,
        colStart,
        colEnd,
      };
    }
  }

  return {
    rowStart: row,
    rowEnd: row,
    colStart: col,
    colEnd: col,
  };
}

function rangeFromScopePoints(scope, anchorPoint, focusPoint) {
  const anchorExtent = selectionExtentForPoint(scope, anchorPoint);
  const focusExtent = selectionExtentForPoint(scope, focusPoint);
  if (!anchorExtent || !focusExtent) {
    return rangeFromPoints(anchorPoint, focusPoint);
  }
  return normalizeSelectionRange({
    rowStart: Math.min(anchorExtent.rowStart, focusExtent.rowStart),
    rowEnd: Math.max(anchorExtent.rowEnd, focusExtent.rowEnd),
    colStart: Math.min(anchorExtent.colStart, focusExtent.colStart),
    colEnd: Math.max(anchorExtent.colEnd, focusExtent.colEnd),
  });
}

function selectionScopeForCell(cell) {
  if (!cell) {
    return null;
  }
  if (el.mainTableWrap && el.mainTableWrap.contains(cell)) {
    return "main";
  }
  if (el.referenceTableWrap && el.referenceTableWrap.contains(cell)) {
    return "reference";
  }
  return null;
}

function selectionPointForCell(cell) {
  if (!cell) {
    return null;
  }
  const table = cell.closest("table");
  if (!table) {
    return null;
  }
  if (!cell.dataset.gridRow || !cell.dataset.gridColStart) {
    annotateSelectionGridForTable(table);
  }

  const row = Number.parseInt(cell.dataset.gridRow || "", 10);
  const col = Number.parseInt(cell.dataset.gridColStart || "", 10);
  if (!Number.isFinite(row) || !Number.isFinite(col)) {
    return null;
  }

  return { row, col };
}

function clearContextMenuLongPressTimer() {
  if (!state.contextMenuLongPressTimer) {
    return;
  }
  window.clearTimeout(state.contextMenuLongPressTimer);
  state.contextMenuLongPressTimer = null;
}

function hideGridContextMenu() {
  clearContextMenuLongPressTimer();
  if (!el.gridContextMenu) {
    return;
  }
  el.gridContextMenu.hidden = true;
  el.gridContextMenu.setAttribute("aria-hidden", "true");
  state.contextMenuOpen = false;
  state.contextMenuScope = null;
  state.contextMenuPoint = null;
}

function clearSheetTabContextLongPressTimer() {
  if (!state.sheetTabContextLongPressTimer) {
    return;
  }
  window.clearTimeout(state.sheetTabContextLongPressTimer);
  state.sheetTabContextLongPressTimer = null;
}

function hideSheetTabContextMenu() {
  clearSheetTabContextLongPressTimer();
  if (!el.sheetTabContextMenu) {
    return;
  }
  el.sheetTabContextMenu.hidden = true;
  el.sheetTabContextMenu.setAttribute("aria-hidden", "true");
  state.sheetTabContextMenuOpen = false;
  state.sheetTabContext = null;
}

function showSheetTabContextMenu(tabContext, clientX, clientY) {
  if (!el.sheetTabContextMenu || !tabContext) {
    return;
  }

  hideGridContextMenu();
  state.sheetTabContextMenuOpen = true;
  state.sheetTabContext = tabContext;

  const viewLabel = tabContext.view === "reference" ? "Detail Sheet" : "Main Sheet";
  if (el.sheetTabContextTitle) {
    setText(el.sheetTabContextTitle, `${viewLabel} · ${tabContext.sheetName}`);
  }
  if (el.sheetTabCtxRenameBtn) {
    el.sheetTabCtxRenameBtn.disabled = !canEditWorkbookTabs();
  }

  el.sheetTabContextMenu.hidden = false;
  el.sheetTabContextMenu.setAttribute("aria-hidden", "false");
  el.sheetTabContextMenu.style.left = "0px";
  el.sheetTabContextMenu.style.top = "0px";

  const menuRect = el.sheetTabContextMenu.getBoundingClientRect();
  const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
  const margin = 8;
  const maxLeft = Math.max(margin, viewportWidth - menuRect.width - margin);
  const maxTop = Math.max(margin, viewportHeight - menuRect.height - margin);
  const left = Math.min(Math.max(margin, Number(clientX)), maxLeft);
  const top = Math.min(Math.max(margin, Number(clientY)), maxTop);
  el.sheetTabContextMenu.style.left = `${Math.round(left)}px`;
  el.sheetTabContextMenu.style.top = `${Math.round(top)}px`;
}

function openSheetTabContextMenu(tabContext, clientX, clientY) {
  if (!tabContext || !tabContext.tabButton) {
    return;
  }
  const button = tabContext.tabButton;
  if (!(button instanceof HTMLElement)) {
    return;
  }

  const rect = button.getBoundingClientRect();
  const fallbackX = rect.left + Math.min(rect.width * 0.5, 120);
  const fallbackY = rect.bottom + 6;
  const posX = Number.isFinite(Number(clientX)) ? Number(clientX) : fallbackX;
  const posY = Number.isFinite(Number(clientY)) ? Number(clientY) : fallbackY;
  setActiveViewScope(tabContext.view);
  showSheetTabContextMenu(tabContext, posX, posY);
}

function renameSheetFromTabContextMenu() {
  if (!canEditWorkbookTabs()) {
    hideSheetTabContextMenu();
    return;
  }
  const tabContext = state.sheetTabContext;
  if (!tabContext) {
    hideSheetTabContextMenu();
    return;
  }
  const { view, sheetName, tabItem, tabButton } = tabContext;
  if (!(tabItem instanceof HTMLElement) || !(tabButton instanceof HTMLElement) || !tabItem.isConnected || !tabButton.isConnected) {
    hideSheetTabContextMenu();
    return;
  }
  hideSheetTabContextMenu();
  beginSheetTabRename(view, sheetName, tabItem, tabButton, null);
}

function bindSheetTabContextMenu(button, view, sheetName, tabItem) {
  if (!button || !tabItem) {
    return;
  }
  const tabContext = {
    view: view === "reference" ? "reference" : "main",
    sheetName: String(sheetName || ""),
    tabItem,
    tabButton: button,
  };

  button.addEventListener("contextmenu", (event) => {
    event.preventDefault();
    openSheetTabContextMenu(tabContext, event.clientX, event.clientY);
  });

  button.addEventListener(
    "touchstart",
    (event) => {
      if (!event.touches || event.touches.length !== 1) {
        clearSheetTabContextLongPressTimer();
        return;
      }
      const touch = event.touches[0];
      const touchX = Number(touch.clientX || 0);
      const touchY = Number(touch.clientY || 0);
      clearSheetTabContextLongPressTimer();
      state.sheetTabContextLongPressTimer = window.setTimeout(() => {
        state.sheetTabLongPressTriggered = true;
        window.setTimeout(() => {
          state.sheetTabLongPressTriggered = false;
        }, 1200);
        openSheetTabContextMenu(tabContext, touchX, touchY);
        state.sheetTabContextLongPressTimer = null;
      }, LONG_PRESS_OPEN_MS);
    },
    { passive: true },
  );

  const clearLongPress = () => {
    clearSheetTabContextLongPressTimer();
  };
  button.addEventListener(
    "touchmove",
    () => {
      clearLongPress();
    },
    { passive: true },
  );
  button.addEventListener(
    "touchend",
    () => {
      clearLongPress();
    },
    { passive: true },
  );
  button.addEventListener(
    "touchcancel",
    () => {
      clearLongPress();
    },
    { passive: true },
  );
}

function detailsContextActionState(scope) {
  const normalized = normalizeViewScope(scope);
  const drawerMode = document.body.classList.contains("detail-drawer-mode");
  if (drawerMode) {
    if (!state.referenceDrawerExpanded) {
      return { label: "Open details drawer", disabled: false };
    }
    if (normalized === "reference") {
      return { label: "Details drawer open", disabled: true };
    }
    return { label: "Focus details drawer", disabled: false };
  }
  if (state.desktopDetailCollapsed) {
    return { label: "Open details panel", disabled: false };
  }
  if (normalized === "reference") {
    return { label: "Details panel open", disabled: true };
  }
  return { label: "Focus details panel", disabled: false };
}

function openDetailsPanelFromContext() {
  const sourceScope = state.contextMenuScope || state.selectionScope || state.activeViewScope;
  if (document.body.classList.contains("detail-drawer-mode")) {
    setReferenceDrawerExpanded(true);
  } else if (state.desktopDetailCollapsed) {
    setDesktopDetailCollapsed(false);
  }
  setActiveViewScope("reference");
  setText(el.statusText, `Detail view ready · ${workbookNameForScope("reference")}`);
  if (sourceScope === "main" && !hasGridSelection()) {
    resetSelectionStatusBar();
  }
}

function showGridContextMenu(clientX, clientY, scope, point) {
  if (!el.gridContextMenu) {
    return;
  }
  if (!scope || !point) {
    hideGridContextMenu();
    return;
  }
  hideSheetTabContextMenu();

  state.contextMenuOpen = true;
  state.contextMenuScope = scope;
  state.contextMenuPoint = { ...point };
  if (el.gridContextTitle) {
    setText(el.gridContextTitle, `${selectionScopeLabel(scope)} · ${pointToAddress(point)}`);
  }

  const hasSelection = Boolean(state.selectionRanges && state.selectionRanges.length);
  const normalizedScope = normalizeViewScope(scope);
  const viewLayout = viewLayoutForScope(normalizedScope);
  const rowIndexes = collectContextRowIndices(normalizedScope);
  const colIndexes = collectContextColumnIndices(normalizedScope);
  const freeze = effectiveFreezeForScope(
    normalizedScope,
    Number(viewLayout.lastAppliedFrozenCols || 0),
    Number(viewLayout.lastAppliedFrozenRows || 0),
  );
  if (el.ctxCopyBtn) {
    el.ctxCopyBtn.disabled = !hasSelection;
  }
  if (el.ctxPasteBtn) {
    el.ctxPasteBtn.disabled = !point;
  }
  if (el.ctxClearSelectionBtn) {
    el.ctxClearSelectionBtn.disabled = !hasSelection;
  }
  if (el.ctxFreezePanesBtn) {
    setText(el.ctxFreezePanesBtn, `Freeze panes at ${pointToAddress(point)}`);
    el.ctxFreezePanesBtn.disabled = !point;
  }
  if (el.ctxUnfreezePanesBtn) {
    el.ctxUnfreezePanesBtn.disabled = freeze.cols < 1 && freeze.rows < 1;
  }
  if (el.ctxHideRowsBtn) {
    setText(el.ctxHideRowsBtn, rowIndexes.size > 0 ? `Hide rows (${rowIndexes.size})` : "Hide rows");
    el.ctxHideRowsBtn.disabled = rowIndexes.size < 1;
  }
  if (el.ctxUnhideRowsBtn) {
    const hiddenRowsCount = viewLayout.hiddenRows.size;
    setText(el.ctxUnhideRowsBtn, hiddenRowsCount > 0 ? `Unhide rows (${hiddenRowsCount})` : "Unhide rows");
    el.ctxUnhideRowsBtn.disabled = hiddenRowsCount < 1;
  }
  if (el.ctxHideColsBtn) {
    setText(el.ctxHideColsBtn, colIndexes.size > 0 ? `Hide columns (${colIndexes.size})` : "Hide columns");
    el.ctxHideColsBtn.disabled = colIndexes.size < 1;
  }
  if (el.ctxUnhideColsBtn) {
    const hiddenColsCount = viewLayout.hiddenCols.size;
    setText(el.ctxUnhideColsBtn, hiddenColsCount > 0 ? `Unhide columns (${hiddenColsCount})` : "Unhide columns");
    el.ctxUnhideColsBtn.disabled = hiddenColsCount < 1;
  }
  if (el.ctxHideBothBtn) {
    el.ctxHideBothBtn.disabled = rowIndexes.size < 1 && colIndexes.size < 1;
  }
  if (el.ctxUnhideBothBtn) {
    el.ctxUnhideBothBtn.disabled = viewLayout.hiddenRows.size < 1 && viewLayout.hiddenCols.size < 1;
  }
  const structure = structuralContextState(normalizedScope);
  if (el.ctxInsertRowAboveBtn) {
    setText(el.ctxInsertRowAboveBtn, "Insert row above");
    el.ctxInsertRowAboveBtn.disabled = !structure.canInsertRowAbove;
  }
  if (el.ctxInsertRowBelowBtn) {
    setText(el.ctxInsertRowBelowBtn, "Insert row below");
    el.ctxInsertRowBelowBtn.disabled = !structure.canInsertRowBelow;
  }
  if (el.ctxDeleteRowBtn) {
    const deletableRows = structure.deletableRowIndices.length;
    setText(el.ctxDeleteRowBtn, deletableRows > 0 ? `Delete row (${deletableRows})` : "Delete row");
    el.ctxDeleteRowBtn.disabled = !structure.canDeleteRow;
  }
  if (el.ctxInsertColLeftBtn) {
    setText(el.ctxInsertColLeftBtn, "Insert column left");
    el.ctxInsertColLeftBtn.disabled = !structure.canInsertColLeft;
  }
  if (el.ctxInsertColRightBtn) {
    setText(el.ctxInsertColRightBtn, "Insert column right");
    el.ctxInsertColRightBtn.disabled = !structure.canInsertColRight;
  }
  if (el.ctxDeleteColBtn) {
    const deletableCols = structure.selectedColumnIndices.length;
    setText(el.ctxDeleteColBtn, deletableCols > 0 ? `Delete column (${deletableCols})` : "Delete column");
    el.ctxDeleteColBtn.disabled = !structure.canDeleteCol;
  }
  if (el.ctxOpenDetailsBtn) {
    const detailsAction = detailsContextActionState(scope);
    setText(el.ctxOpenDetailsBtn, detailsAction.label);
    el.ctxOpenDetailsBtn.disabled = detailsAction.disabled;
  }

  el.gridContextMenu.hidden = false;
  el.gridContextMenu.setAttribute("aria-hidden", "false");
  el.gridContextMenu.style.left = "0px";
  el.gridContextMenu.style.top = "0px";

  const menuRect = el.gridContextMenu.getBoundingClientRect();
  const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
  const margin = 8;
  const maxLeft = Math.max(margin, viewportWidth - menuRect.width - margin);
  const maxTop = Math.max(margin, viewportHeight - menuRect.height - margin);
  const left = Math.min(Math.max(margin, clientX), maxLeft);
  const top = Math.min(Math.max(margin, clientY), maxTop);
  el.gridContextMenu.style.left = `${Math.round(left)}px`;
  el.gridContextMenu.style.top = `${Math.round(top)}px`;
}

function rangeToAddress(range) {
  if (!range) {
    return "";
  }
  const normalized = normalizeSelectionRange(range);
  const start = pointToAddress({ row: normalized.rowStart, col: normalized.colStart });
  const end = pointToAddress({ row: normalized.rowEnd, col: normalized.colEnd });
  return start === end ? start : `${start}:${end}`;
}

function rangeIncludesCell(range, rowStart, rowEnd, colStart, colEnd) {
  const normalized = normalizeSelectionRange(range);
  if (rowEnd < normalized.rowStart || rowStart > normalized.rowEnd) {
    return false;
  }
  return colEnd >= normalized.colStart && colStart <= normalized.colEnd;
}

function rangeFullyContainsCell(range, rowStart, rowEnd, colStart, colEnd) {
  const normalized = normalizeSelectionRange(range);
  return (
    rowStart >= normalized.rowStart &&
    rowEnd <= normalized.rowEnd &&
    colStart >= normalized.colStart &&
    colEnd <= normalized.colEnd
  );
}

function pointInRange(range, point) {
  const normalized = normalizeSelectionRange(range);
  return (
    point.row >= normalized.rowStart &&
    point.row <= normalized.rowEnd &&
    point.col >= normalized.colStart &&
    point.col <= normalized.colEnd
  );
}

function excelColumnLabelToIndex(label) {
  const normalized = String(label || "")
    .replace(/\$/g, "")
    .trim()
    .toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) {
    return null;
  }
  let value = 0;
  for (const char of normalized) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  const index = value - 1;
  return index >= 0 ? index : null;
}

function parseExcelAddress(addressText) {
  const raw = String(addressText || "").trim();
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(raw);
  if (!match) {
    return null;
  }
  const col = excelColumnLabelToIndex(match[1]);
  const row = Number.parseInt(match[2], 10) - 1;
  if (!Number.isFinite(col) || !Number.isFinite(row) || row < 0 || col < 0) {
    return null;
  }
  return { row, col };
}

function formulaScalarFromCell(cell) {
  if (!cell) {
    return "";
  }
  const text = String(cell.textContent || "").replace(/\s+/g, " ").trim();
  if (!text.length) {
    return "";
  }
  const numeric = parseNumericCellValue(text);
  if (Number.isFinite(numeric)) {
    return numeric;
  }
  const upper = text.toUpperCase();
  if (upper === "TRUE") {
    return true;
  }
  if (upper === "FALSE") {
    return false;
  }
  return text;
}

function flattenFormulaValues(value) {
  if (Array.isArray(value)) {
    return value.flatMap((item) => flattenFormulaValues(item));
  }
  return [value];
}

function formulaNumberValue(value, options = {}) {
  const { blankAsZero = true } = options;
  const scalar = Array.isArray(value) ? (value.length ? value[0] : "") : value;
  if (scalar === null || scalar === undefined || scalar === "") {
    return blankAsZero ? 0 : NaN;
  }
  if (typeof scalar === "number") {
    return Number.isFinite(scalar) ? scalar : NaN;
  }
  if (typeof scalar === "boolean") {
    return scalar ? 1 : 0;
  }
  const parsed = parseNumericCellValue(String(scalar));
  return Number.isFinite(parsed) ? parsed : NaN;
}

function formulaBooleanValue(value) {
  const scalar = Array.isArray(value) ? (value.length ? value[0] : "") : value;
  if (typeof scalar === "boolean") {
    return scalar;
  }
  if (typeof scalar === "number") {
    return scalar !== 0;
  }
  if (scalar === null || scalar === undefined || scalar === "") {
    return false;
  }
  const text = String(scalar).trim();
  if (!text) {
    return false;
  }
  if (/^(true|yes)$/i.test(text)) {
    return true;
  }
  if (/^(false|no)$/i.test(text)) {
    return false;
  }
  const numeric = parseNumericCellValue(text);
  if (Number.isFinite(numeric)) {
    return numeric !== 0;
  }
  return true;
}

function formulaCompareValues(left, right, op) {
  const leftNum = formulaNumberValue(left, { blankAsZero: false });
  const rightNum = formulaNumberValue(right, { blankAsZero: false });
  const bothNumeric = Number.isFinite(leftNum) && Number.isFinite(rightNum);
  const a = bothNumeric ? leftNum : String(Array.isArray(left) ? left[0] ?? "" : left ?? "");
  const b = bothNumeric ? rightNum : String(Array.isArray(right) ? right[0] ?? "" : right ?? "");

  if (op === "=") {
    return a === b;
  }
  if (op === "<>") {
    return a !== b;
  }
  if (op === "<") {
    return a < b;
  }
  if (op === "<=") {
    return a <= b;
  }
  if (op === ">") {
    return a > b;
  }
  if (op === ">=") {
    return a >= b;
  }
  throw new Error(`Unsupported comparison operator "${op}".`);
}

function formulaValuesFromRange(scope, pointCellMap, startAddress, endAddress) {
  const start = parseExcelAddress(startAddress);
  const end = parseExcelAddress(endAddress);
  if (!start || !end) {
    throw new Error(`Invalid range "${startAddress}:${endAddress}".`);
  }
  const rowStart = Math.min(start.row, end.row);
  const rowEnd = Math.max(start.row, end.row);
  const colStart = Math.min(start.col, end.col);
  const colEnd = Math.max(start.col, end.col);
  const values = [];
  const pointMap = pointCellMap || buildPointCellMap(scope);
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      values.push(formulaScalarFromCell(pointMap.get(`${row}:${col}`) || null));
    }
  }
  return values;
}

function formulaValueFromReference(scope, pointCellMap, address) {
  const point = parseExcelAddress(address);
  if (!point) {
    throw new Error(`Invalid reference "${address}".`);
  }
  const pointMap = pointCellMap || buildPointCellMap(scope);
  return formulaScalarFromCell(pointMap.get(`${point.row}:${point.col}`) || null);
}

function numericFormulaArgs(args) {
  const numeric = [];
  for (const arg of args) {
    for (const scalar of flattenFormulaValues(arg)) {
      const parsed = formulaNumberValue(scalar, { blankAsZero: false });
      if (Number.isFinite(parsed)) {
        numeric.push(parsed);
      }
    }
  }
  return numeric;
}

function countAFormulaArgs(args) {
  let count = 0;
  for (const arg of args) {
    for (const scalar of flattenFormulaValues(arg)) {
      if (scalar === null || scalar === undefined) {
        continue;
      }
      if (typeof scalar === "string" && scalar.trim() === "") {
        continue;
      }
      count += 1;
    }
  }
  return count;
}

function runFormulaFunction(name, args) {
  const fn = String(name || "").trim().toUpperCase();
  if (!FORMULA_FUNCTION_NAMES.has(fn)) {
    throw new Error(`Unsupported function "${name}".`);
  }

  if (fn === "SUM") {
    return numericFormulaArgs(args).reduce((sum, value) => sum + value, 0);
  }
  if (fn === "AVERAGE" || fn === "AVG") {
    const values = numericFormulaArgs(args);
    if (!values.length) {
      throw new Error("AVERAGE requires at least one numeric value.");
    }
    return values.reduce((sum, value) => sum + value, 0) / values.length;
  }
  if (fn === "MIN") {
    const values = numericFormulaArgs(args);
    if (!values.length) {
      throw new Error("MIN requires at least one numeric value.");
    }
    return Math.min(...values);
  }
  if (fn === "MAX") {
    const values = numericFormulaArgs(args);
    if (!values.length) {
      throw new Error("MAX requires at least one numeric value.");
    }
    return Math.max(...values);
  }
  if (fn === "COUNT") {
    return numericFormulaArgs(args).length;
  }
  if (fn === "COUNTA") {
    return countAFormulaArgs(args);
  }
  if (fn === "IF") {
    if (args.length < 2) {
      throw new Error("IF requires at least 2 arguments.");
    }
    return formulaBooleanValue(args[0]) ? args[1] : args.length > 2 ? args[2] : false;
  }
  if (fn === "ABS") {
    const value = formulaNumberValue(args[0], { blankAsZero: false });
    if (!Number.isFinite(value)) {
      throw new Error("ABS requires a numeric value.");
    }
    return Math.abs(value);
  }
  if (fn === "ROUND") {
    const value = formulaNumberValue(args[0], { blankAsZero: false });
    if (!Number.isFinite(value)) {
      throw new Error("ROUND requires a numeric value.");
    }
    const digits = Math.trunc(formulaNumberValue(args[1] ?? 0));
    const factor = 10 ** digits;
    return Math.round(value * factor) / factor;
  }
  if (fn === "ROUNDUP") {
    const value = formulaNumberValue(args[0], { blankAsZero: false });
    if (!Number.isFinite(value)) {
      throw new Error("ROUNDUP requires a numeric value.");
    }
    const digits = Math.trunc(formulaNumberValue(args[1] ?? 0));
    const factor = 10 ** digits;
    if (value >= 0) {
      return Math.ceil(value * factor) / factor;
    }
    return Math.floor(value * factor) / factor;
  }
  if (fn === "ROUNDDOWN") {
    const value = formulaNumberValue(args[0], { blankAsZero: false });
    if (!Number.isFinite(value)) {
      throw new Error("ROUNDDOWN requires a numeric value.");
    }
    const digits = Math.trunc(formulaNumberValue(args[1] ?? 0));
    const factor = 10 ** digits;
    if (value >= 0) {
      return Math.floor(value * factor) / factor;
    }
    return Math.ceil(value * factor) / factor;
  }
  if (fn === "LEN") {
    const scalar = Array.isArray(args[0]) ? args[0][0] ?? "" : args[0];
    return String(scalar ?? "").length;
  }

  throw new Error(`Unsupported function "${name}".`);
}

function tokenizeFormulaExpression(expression) {
  const source = String(expression || "").trim();
  const tokens = [];
  let index = 0;

  const push = (type, value = null) => {
    tokens.push({ type, value });
  };

  while (index < source.length) {
    const char = source[index];

    if (/\s/.test(char)) {
      index += 1;
      continue;
    }

    if (char === "(") {
      push("lparen");
      index += 1;
      continue;
    }
    if (char === ")") {
      push("rparen");
      index += 1;
      continue;
    }
    if (char === ",") {
      push("comma");
      index += 1;
      continue;
    }
    if (char === ";") {
      push("comma");
      index += 1;
      continue;
    }
    if (char === ":") {
      push("colon");
      index += 1;
      continue;
    }
    if (char === "+" || char === "-" || char === "*" || char === "/" || char === "^" || char === "&") {
      push("operator", char);
      index += 1;
      continue;
    }
    if (char === "<" || char === ">" || char === "=") {
      const next = source[index + 1] || "";
      if ((char === "<" || char === ">") && next === "=") {
        push("operator", `${char}=`);
        index += 2;
        continue;
      }
      if (char === "<" && next === ">") {
        push("operator", "<>");
        index += 2;
        continue;
      }
      push("operator", char);
      index += 1;
      continue;
    }
    if (char === '"') {
      let value = "";
      index += 1;
      while (index < source.length) {
        const current = source[index];
        if (current === '"') {
          if (source[index + 1] === '"') {
            value += '"';
            index += 2;
            continue;
          }
          index += 1;
          break;
        }
        value += current;
        index += 1;
      }
      push("string", value);
      continue;
    }
    if (/\d|\./.test(char)) {
      const start = index;
      let dotCount = 0;
      while (index < source.length && /[\d.]/.test(source[index])) {
        if (source[index] === ".") {
          dotCount += 1;
        }
        index += 1;
      }
      const raw = source.slice(start, index);
      if (dotCount > 1 || raw === ".") {
        throw new Error(`Invalid number "${raw}".`);
      }
      const value = Number.parseFloat(raw);
      if (!Number.isFinite(value)) {
        throw new Error(`Invalid number "${raw}".`);
      }
      push("number", value);
      continue;
    }
    if (/[A-Za-z_$]/.test(char)) {
      const start = index;
      while (index < source.length && /[A-Za-z0-9_$]/.test(source[index])) {
        index += 1;
      }
      const raw = source.slice(start, index);
      const normalized = raw.replace(/\$/g, "");
      if (/^[A-Za-z]+\d+$/.test(normalized)) {
        push("ref", normalized.toUpperCase());
        continue;
      }
      if (/^(true|false)$/i.test(normalized)) {
        push("boolean", /^true$/i.test(normalized));
        continue;
      }
      push("ident", normalized.toUpperCase());
      continue;
    }

    throw new Error(`Unexpected character "${char}" in formula.`);
  }

  push("eof");
  return tokens;
}

function evaluateFormulaExpression(rawFormula, scope, pointCellMap = null) {
  const source = String(rawFormula || "").trim().replace(/^=/, "");
  if (!source) {
    return "";
  }
  const tokens = tokenizeFormulaExpression(source);
  let current = 0;
  const pointMap = pointCellMap || buildPointCellMap(scope);

  const peek = () => tokens[current] || { type: "eof", value: null };
  const previous = () => tokens[Math.max(0, current - 1)] || { type: "eof", value: null };
  const advance = () => {
    if (current < tokens.length) {
      current += 1;
    }
    return previous();
  };
  const check = (type, value = null) => {
    const token = peek();
    if (token.type !== type) {
      return false;
    }
    if (value === null) {
      return true;
    }
    return token.value === value;
  };
  const match = (type, value = null) => {
    if (!check(type, value)) {
      return false;
    }
    advance();
    return true;
  };
  const consume = (type, message, value = null) => {
    if (match(type, value)) {
      return previous();
    }
    throw new Error(message);
  };

  const parsePrimary = () => {
    if (match("number")) {
      return previous().value;
    }
    if (match("string")) {
      return previous().value;
    }
    if (match("boolean")) {
      return previous().value;
    }
    if (match("ref")) {
      const startRef = previous().value;
      if (match("colon")) {
        const endRef = consume("ref", "Expected cell reference after ':' in range.").value;
        return formulaValuesFromRange(scope, pointMap, startRef, endRef);
      }
      return formulaValueFromReference(scope, pointMap, startRef);
    }
    if (match("ident")) {
      const identifier = previous().value;
      if (!match("lparen")) {
        throw new Error(`Unknown identifier "${identifier}".`);
      }
      const args = [];
      if (!check("rparen")) {
        do {
          args.push(parseComparison());
        } while (match("comma"));
      }
      consume("rparen", `Expected ')' after ${identifier} arguments.`);
      return runFormulaFunction(identifier, args);
    }
    if (match("lparen")) {
      const nested = parseComparison();
      consume("rparen", "Expected ')' to close formula expression.");
      return nested;
    }
    throw new Error("Unexpected token in formula.");
  };

  const parseUnary = () => {
    if (match("operator", "+")) {
      return formulaNumberValue(parseUnary());
    }
    if (match("operator", "-")) {
      const value = formulaNumberValue(parseUnary());
      if (!Number.isFinite(value)) {
        throw new Error("Unary minus requires a numeric value.");
      }
      return -value;
    }
    return parsePrimary();
  };

  const parsePower = () => {
    let value = parseUnary();
    while (match("operator", "^")) {
      const right = parseUnary();
      const leftNum = formulaNumberValue(value, { blankAsZero: false });
      const rightNum = formulaNumberValue(right, { blankAsZero: false });
      if (!Number.isFinite(leftNum) || !Number.isFinite(rightNum)) {
        throw new Error("Power operator requires numeric values.");
      }
      value = leftNum ** rightNum;
    }
    return value;
  };

  const parseProduct = () => {
    let value = parsePower();
    while (check("operator", "*") || check("operator", "/")) {
      const op = advance().value;
      const right = parsePower();
      const leftNum = formulaNumberValue(value, { blankAsZero: false });
      const rightNum = formulaNumberValue(right, { blankAsZero: false });
      if (!Number.isFinite(leftNum) || !Number.isFinite(rightNum)) {
        throw new Error(`${op} requires numeric values.`);
      }
      if (op === "/" && rightNum === 0) {
        throw new Error("Division by zero.");
      }
      value = op === "*" ? leftNum * rightNum : leftNum / rightNum;
    }
    return value;
  };

  const parseSum = () => {
    let value = parseProduct();
    while (check("operator", "+") || check("operator", "-") || check("operator", "&")) {
      const op = advance().value;
      const right = parseProduct();
      if (op === "&") {
        const leftText = Array.isArray(value) ? value[0] ?? "" : value ?? "";
        const rightText = Array.isArray(right) ? right[0] ?? "" : right ?? "";
        value = `${leftText}${rightText}`;
        continue;
      }
      const leftNum = formulaNumberValue(value);
      const rightNum = formulaNumberValue(right);
      if (!Number.isFinite(leftNum) || !Number.isFinite(rightNum)) {
        throw new Error(`${op} requires numeric values.`);
      }
      value = op === "+" ? leftNum + rightNum : leftNum - rightNum;
    }
    return value;
  };

  const parseComparison = () => {
    let value = parseSum();
    while (
      check("operator", "=") ||
      check("operator", "<>") ||
      check("operator", "<") ||
      check("operator", "<=") ||
      check("operator", ">") ||
      check("operator", ">=")
    ) {
      const op = advance().value;
      const right = parseSum();
      value = formulaCompareValues(value, right, op);
    }
    return value;
  };

  const result = parseComparison();
  consume("eof", "Unexpected trailing formula content.");
  return result;
}

function selectionEdgeClassesForCell(rowStart, rowEnd, colStart, colEnd, ranges) {
  let top = false;
  let right = false;
  let bottom = false;
  let left = false;

  for (const range of ranges) {
    const normalized = normalizeSelectionRange(range);
    if (
      rowEnd < normalized.rowStart ||
      rowStart > normalized.rowEnd ||
      colEnd < normalized.colStart ||
      colStart > normalized.colEnd
    ) {
      continue;
    }

    if (normalized.rowStart === rowStart) {
      top = true;
    }
    if (normalized.colEnd === colEnd) {
      right = true;
    }
    if (normalized.rowEnd === rowEnd) {
      bottom = true;
    }
    if (normalized.colStart === colStart) {
      left = true;
    }

    if (top && right && bottom && left) {
      break;
    }
  }

  const classes = [];
  if (top) {
    classes.push("grid-cell-edge-top");
  }
  if (right) {
    classes.push("grid-cell-edge-right");
  }
  if (bottom) {
    classes.push("grid-cell-edge-bottom");
  }
  if (left) {
    classes.push("grid-cell-edge-left");
  }
  return classes;
}

function isPointInSelection(scope, point) {
  if (!scope || !point || state.selectionScope !== scope) {
    return false;
  }
  const ranges = Array.isArray(state.selectionRanges) ? state.selectionRanges : [];
  return ranges.some((range) => pointInRange(range, point));
}

function selectSingleGridPoint(scope, point) {
  if (!scope || !point) {
    return;
  }
  setActiveViewScope(scope, { refresh: false });
  annotateSelectionGridForScope(scope);
  state.selectionScope = scope;
  state.selectionAnchor = { ...point };
  state.selectionFocus = { ...point };
  state.selectionRanges = [rangeFromScopePoints(scope, point, point)];
  refreshGridSelectionVisuals();
}

function openContextMenuForCell(cell, clientX, clientY) {
  if (!cell) {
    hideGridContextMenu();
    return false;
  }
  const scope = selectionScopeForCell(cell);
  if (!scope) {
    hideGridContextMenu();
    return false;
  }
  setActiveViewScope(scope, { refresh: false });
  annotateSelectionGridForScope(scope);
  const point = selectionPointForCell(cell);
  if (!point) {
    hideGridContextMenu();
    return false;
  }
  if (!isPointInSelection(scope, point)) {
    selectSingleGridPoint(scope, point);
  }
  showGridContextMenu(clientX, clientY, scope, point);
  return true;
}

function boundsForScope(scope) {
  let maxRow = -1;
  let maxCol = -1;
  for (const table of tablesForScope(scope)) {
    annotateSelectionGridForTable(table);
    const rows = Number.parseInt(table.dataset.selectionRowCount || "0", 10) || 0;
    const cols = Number.parseInt(table.dataset.selectionColCount || "0", 10) || 0;
    if (rows > 0) {
      maxRow = Math.max(maxRow, rows - 1);
    }
    if (cols > 0) {
      maxCol = Math.max(maxCol, cols - 1);
    }
  }
  return { maxRow, maxCol };
}

function selectContextRow() {
  const scope = state.contextMenuScope;
  const point = state.contextMenuPoint;
  if (!scope || !point) {
    return;
  }
  const bounds = boundsForScope(scope);
  if (bounds.maxCol < 0) {
    return;
  }
  const range = normalizeSelectionRange({
    rowStart: point.row,
    rowEnd: point.row,
    colStart: 0,
    colEnd: bounds.maxCol,
  });
  setActiveViewScope(scope, { refresh: false });
  state.selectionScope = scope;
  state.selectionAnchor = { row: point.row, col: 0 };
  state.selectionFocus = { row: point.row, col: bounds.maxCol };
  state.selectionRanges = [range];
  refreshGridSelectionVisuals();
}

function selectContextColumn() {
  const scope = state.contextMenuScope;
  const point = state.contextMenuPoint;
  if (!scope || !point) {
    return;
  }
  const bounds = boundsForScope(scope);
  if (bounds.maxRow < 0) {
    return;
  }
  const range = normalizeSelectionRange({
    rowStart: 0,
    rowEnd: bounds.maxRow,
    colStart: point.col,
    colEnd: point.col,
  });
  setActiveViewScope(scope, { refresh: false });
  state.selectionScope = scope;
  state.selectionAnchor = { row: 0, col: point.col };
  state.selectionFocus = { row: bounds.maxRow, col: point.col };
  state.selectionRanges = [range];
  refreshGridSelectionVisuals();
}

function selectAllForContextScope() {
  const scope = state.contextMenuScope;
  if (!scope) {
    return;
  }
  const bounds = boundsForScope(scope);
  if (bounds.maxRow < 0 || bounds.maxCol < 0) {
    return;
  }
  const range = normalizeSelectionRange({
    rowStart: 0,
    rowEnd: bounds.maxRow,
    colStart: 0,
    colEnd: bounds.maxCol,
  });
  setActiveViewScope(scope, { refresh: false });
  state.selectionScope = scope;
  state.selectionAnchor = { row: 0, col: 0 };
  state.selectionFocus = { row: bounds.maxRow, col: bounds.maxCol };
  state.selectionRanges = [range];
  refreshGridSelectionVisuals();
}

function contextSelectionRangesForScope(scope) {
  const normalized = normalizeViewScope(scope);
  if (state.selectionScope === normalized && Array.isArray(state.selectionRanges) && state.selectionRanges.length) {
    return state.selectionRanges.map((range) => normalizeSelectionRange(range));
  }
  if (state.contextMenuScope === normalized && state.contextMenuPoint) {
    annotateSelectionGridForScope(normalized);
    return [normalizeSelectionRange(rangeFromScopePoints(normalized, state.contextMenuPoint, state.contextMenuPoint))];
  }
  return [];
}

function collectContextRowIndices(scope) {
  const indices = new Set();
  for (const range of contextSelectionRangesForScope(scope)) {
    for (let row = range.rowStart; row <= range.rowEnd; row += 1) {
      indices.add(row);
    }
  }
  return indices;
}

function collectContextColumnIndices(scope) {
  const indices = new Set();
  for (const range of contextSelectionRangesForScope(scope)) {
    for (let col = range.colStart; col <= range.colEnd; col += 1) {
      indices.add(col);
    }
  }
  return indices;
}

function freezeDefaultsForScope(scope) {
  const viewLayout = viewLayoutForScope(scope);
  return {
    cols: Math.max(0, Number.parseInt(String(viewLayout.lastAppliedFrozenCols || 0), 10) || 0),
    rows: Math.max(0, Number.parseInt(String(viewLayout.lastAppliedFrozenRows || 0), 10) || 0),
  };
}

function contextPointForScope(scope) {
  const normalized = normalizeViewScope(scope);
  if (state.contextMenuScope === normalized && state.contextMenuPoint) {
    return state.contextMenuPoint;
  }
  if (state.selectionScope === normalized && state.selectionFocus) {
    return state.selectionFocus;
  }
  return null;
}

function sortedContextRowIndices(scope) {
  const rows = Array.from(collectContextRowIndices(scope));
  if (rows.length) {
    return rows.sort((a, b) => a - b);
  }
  const point = contextPointForScope(scope);
  return point ? [point.row] : [];
}

function sortedContextColumnIndices(scope) {
  const cols = Array.from(collectContextColumnIndices(scope));
  if (cols.length) {
    return cols.sort((a, b) => a - b);
  }
  const point = contextPointForScope(scope);
  return point ? [point.col] : [];
}

function rowByGridIndex(table, gridRow) {
  if (!table || !Number.isFinite(Number(gridRow))) {
    return null;
  }
  for (const row of Array.from(table.rows || [])) {
    const firstCell = row.cells && row.cells.length ? row.cells[0] : null;
    if (!firstCell) {
      continue;
    }
    const rowIndex = Number.parseInt(firstCell.dataset.gridRow || "", 10);
    if (Number.isFinite(rowIndex) && rowIndex === gridRow) {
      return row;
    }
  }
  return null;
}

function rowIsDataRow(table, row) {
  if (!table || !row) {
    return false;
  }
  const parentTag = row.parentElement ? row.parentElement.tagName : "";
  if (table.tBodies && table.tBodies.length > 0) {
    return parentTag === "TBODY";
  }
  return parentTag !== "THEAD";
}

function tableIsSimpleRectGrid(table) {
  if (!table) {
    return false;
  }
  annotateSelectionGridForTable(table);
  const expectedCols = Number.parseInt(table.dataset.selectionColCount || "0", 10) || 0;
  if (expectedCols < 1) {
    return false;
  }
  for (const row of Array.from(table.rows || [])) {
    let rowCols = 0;
    for (const cell of Array.from(row.cells || [])) {
      const colSpan = Math.max(1, Number.parseInt(cell.getAttribute("colspan") || "1", 10) || 1);
      const rowSpan = Math.max(1, Number.parseInt(cell.getAttribute("rowspan") || "1", 10) || 1);
      if (colSpan !== 1 || rowSpan !== 1) {
        return false;
      }
      rowCols += 1;
    }
    if (rowCols !== expectedCols) {
      return false;
    }
  }
  return true;
}

function shiftIndexSetForInsert(indexSet, insertAt, count = 1) {
  if (!(indexSet instanceof Set) || !Number.isFinite(insertAt) || !Number.isFinite(count) || count < 1) {
    return;
  }
  const next = new Set();
  for (const rawIndex of indexSet) {
    const index = Number.parseInt(String(rawIndex), 10);
    if (!Number.isFinite(index) || index < 0) {
      continue;
    }
    next.add(index >= insertAt ? index + count : index);
  }
  indexSet.clear();
  for (const index of next) {
    indexSet.add(index);
  }
}

function shiftIndexSetForDelete(indexSet, removedIndices) {
  if (!(indexSet instanceof Set)) {
    return;
  }
  const sortedRemoved = Array.from(new Set(Array.isArray(removedIndices) ? removedIndices : []))
    .map((value) => Number.parseInt(String(value), 10))
    .filter((value) => Number.isFinite(value) && value >= 0)
    .sort((a, b) => a - b);
  if (!sortedRemoved.length) {
    return;
  }
  const removedSet = new Set(sortedRemoved);
  const next = new Set();
  for (const rawIndex of indexSet) {
    const index = Number.parseInt(String(rawIndex), 10);
    if (!Number.isFinite(index) || index < 0 || removedSet.has(index)) {
      continue;
    }
    let shift = 0;
    for (const removed of sortedRemoved) {
      if (removed < index) {
        shift += 1;
      } else {
        break;
      }
    }
    next.add(Math.max(0, index - shift));
  }
  indexSet.clear();
  for (const index of next) {
    indexSet.add(index);
  }
}

function normalizeColumnWidthPx(value) {
  const numeric = Number.parseFloat(String(value));
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return null;
  }
  return Math.max(MIN_COLUMN_WIDTH_PX, Math.min(MAX_COLUMN_WIDTH_PX, Math.round(numeric)));
}

function normalizeColumnWidthMap(columnWidths) {
  const next = {};
  if (!columnWidths || typeof columnWidths !== "object") {
    return next;
  }
  for (const [rawKey, rawWidth] of Object.entries(columnWidths)) {
    const index = Number.parseInt(String(rawKey), 10);
    const width = normalizeColumnWidthPx(rawWidth);
    if (!Number.isFinite(index) || index < 0 || !Number.isFinite(width)) {
      continue;
    }
    next[String(index)] = width;
  }
  return next;
}

function shiftColumnWidthMapForInsert(columnWidths, insertAt, count = 1) {
  if (!columnWidths || typeof columnWidths !== "object") {
    return {};
  }
  if (!Number.isFinite(insertAt) || !Number.isFinite(count) || count < 1) {
    return normalizeColumnWidthMap(columnWidths);
  }
  const next = {};
  for (const [rawKey, rawWidth] of Object.entries(columnWidths)) {
    const index = Number.parseInt(String(rawKey), 10);
    const width = normalizeColumnWidthPx(rawWidth);
    if (!Number.isFinite(index) || index < 0 || !Number.isFinite(width)) {
      continue;
    }
    const shifted = index >= insertAt ? index + count : index;
    next[String(shifted)] = width;
  }
  return next;
}

function shiftColumnWidthMapForDelete(columnWidths, removedIndices) {
  const sortedRemoved = Array.from(new Set(Array.isArray(removedIndices) ? removedIndices : []))
    .map((value) => Number.parseInt(String(value), 10))
    .filter((value) => Number.isFinite(value) && value >= 0)
    .sort((a, b) => a - b);
  if (!sortedRemoved.length) {
    return normalizeColumnWidthMap(columnWidths);
  }

  const removedSet = new Set(sortedRemoved);
  const next = {};
  for (const [rawKey, rawWidth] of Object.entries(columnWidths || {})) {
    const index = Number.parseInt(String(rawKey), 10);
    const width = normalizeColumnWidthPx(rawWidth);
    if (!Number.isFinite(index) || index < 0 || !Number.isFinite(width) || removedSet.has(index)) {
      continue;
    }
    let shift = 0;
    for (const removed of sortedRemoved) {
      if (removed < index) {
        shift += 1;
      } else {
        break;
      }
    }
    next[String(Math.max(0, index - shift))] = width;
  }
  return next;
}

function structuralContextState(scope) {
  const normalized = normalizeViewScope(scope);
  const point = contextPointForScope(normalized);
  const table = primaryTableForScope(normalized);
  const stateDefault = {
    table,
    point,
    canInsertRowAbove: false,
    canInsertRowBelow: false,
    canDeleteRow: false,
    canInsertColLeft: false,
    canInsertColRight: false,
    canDeleteCol: false,
    deletableRowIndices: [],
    selectedColumnIndices: [],
  };
  if (!table || !point) {
    return stateDefault;
  }

  annotateSelectionGridForTable(table);
  const selectedRowIndices = sortedContextRowIndices(normalized);
  const selectedColumnIndices = sortedContextColumnIndices(normalized);
  const selectedColumnSet = new Set(selectedColumnIndices);
  const totalCols = Number.parseInt(table.dataset.selectionColCount || "0", 10) || 0;

  const pointRow = rowByGridIndex(table, point.row);
  const pointIsDataRow = rowIsDataRow(table, pointRow);
  const deletableRowIndices = selectedRowIndices.filter((rowIndex) => {
    const row = rowByGridIndex(table, rowIndex);
    return rowIsDataRow(table, row);
  });
  const dataRowCount = Array.from(table.rows || []).filter((row) => rowIsDataRow(table, row)).length;

  const simpleRectGrid = tableIsSimpleRectGrid(table);
  const canDeleteCol = simpleRectGrid && selectedColumnSet.size > 0 && totalCols > selectedColumnSet.size;

  return {
    table,
    point,
    canInsertRowAbove: Boolean(pointIsDataRow),
    canInsertRowBelow: Boolean(pointIsDataRow),
    canDeleteRow: deletableRowIndices.length > 0 && dataRowCount > deletableRowIndices.length,
    canInsertColLeft: simpleRectGrid && totalCols > 0,
    canInsertColRight: simpleRectGrid && totalCols > 0,
    canDeleteCol,
    deletableRowIndices,
    selectedColumnIndices,
  };
}

function activeGridScopeForCommands() {
  return normalizeViewScope(state.selectionScope || state.activeViewScope || "main");
}

function withTemporaryGridContext(scope, point, callback) {
  const normalized = normalizeViewScope(scope);
  const previousScope = state.contextMenuScope;
  const previousPoint = state.contextMenuPoint;
  state.contextMenuScope = normalized;
  state.contextMenuPoint = point ? { ...point } : null;
  try {
    callback();
  } finally {
    state.contextMenuScope = previousScope;
    state.contextMenuPoint = previousPoint;
  }
}

function runRibbonGridAction(action, options = {}) {
  const scope = activeGridScopeForCommands();
  const point = contextPointForScope(scope);
  if (options.requirePoint && !point) {
    setText(el.statusText, options.emptyPointMessage || "Select a cell first.");
    syncRibbonGridControlState();
    return;
  }
  withTemporaryGridContext(scope, point, action);
  syncRibbonGridControlState();
}

function syncRibbonGridControlState() {
  const scope = activeGridScopeForCommands();
  const point = contextPointForScope(scope);
  const hasPoint = Boolean(point);
  const bounds = boundsForScope(scope);
  const hasGrid = bounds.maxRow >= 0 && bounds.maxCol >= 0;
  const viewLayout = viewLayoutForScope(scope);
  const freeze = effectiveFreezeForScope(
    scope,
    Number(viewLayout.lastAppliedFrozenCols || 0),
    Number(viewLayout.lastAppliedFrozenRows || 0),
  );
  const structure = structuralContextState(scope);
  const rowIndexes = collectContextRowIndices(scope);
  const colIndexes = collectContextColumnIndices(scope);
  const activeGroupCount = collectRowGroupContexts(scope).length;
  const groupScopeLabel = scope === "reference" ? "Detail view" : "Main view";

  if (el.ribbonGridScopeLabel) {
    const scopeName = scope === "reference" ? "Detail view" : "Main view";
    const pointLabel = hasPoint ? ` · Active: ${pointToAddress(point)}` : " · Select a cell to enable commands";
    setText(el.ribbonGridScopeLabel, `${scopeName}${pointLabel}`);
  }

  if (el.ribbonFreezePanesBtn) {
    setText(el.ribbonFreezePanesBtn, hasPoint ? `Freeze panes at ${pointToAddress(point)}` : "Freeze panes");
    el.ribbonFreezePanesBtn.disabled = !hasPoint;
  }
  if (el.ribbonFreezeTopRowBtn) {
    el.ribbonFreezeTopRowBtn.disabled = !hasGrid;
  }
  if (el.ribbonFreezeFirstColBtn) {
    el.ribbonFreezeFirstColBtn.disabled = !hasGrid;
  }
  if (el.ribbonUnfreezePanesBtn) {
    el.ribbonUnfreezePanesBtn.disabled = freeze.cols < 1 && freeze.rows < 1;
  }

  if (el.ribbonSelectRowBtn) {
    el.ribbonSelectRowBtn.disabled = !hasPoint;
  }
  if (el.ribbonSelectColumnBtn) {
    el.ribbonSelectColumnBtn.disabled = !hasPoint;
  }
  if (el.ribbonSelectAllBtn) {
    el.ribbonSelectAllBtn.disabled = !hasGrid;
  }

  if (el.ribbonInsertRowAboveBtn) {
    el.ribbonInsertRowAboveBtn.disabled = !structure.canInsertRowAbove;
  }
  if (el.ribbonInsertRowBelowBtn) {
    el.ribbonInsertRowBelowBtn.disabled = !structure.canInsertRowBelow;
  }
  if (el.ribbonDeleteRowBtn) {
    const deletableRows = structure.deletableRowIndices.length;
    setText(el.ribbonDeleteRowBtn, deletableRows > 0 ? `Delete row (${deletableRows})` : "Delete row");
    el.ribbonDeleteRowBtn.disabled = !structure.canDeleteRow;
  }
  if (el.ribbonHideRowsBtn) {
    setText(el.ribbonHideRowsBtn, rowIndexes.size > 0 ? `Hide rows (${rowIndexes.size})` : "Hide rows");
    el.ribbonHideRowsBtn.disabled = rowIndexes.size < 1;
  }
  if (el.ribbonUnhideRowsBtn) {
    const hiddenRowsCount = viewLayout.hiddenRows.size;
    setText(el.ribbonUnhideRowsBtn, hiddenRowsCount > 0 ? `Unhide rows (${hiddenRowsCount})` : "Unhide rows");
    el.ribbonUnhideRowsBtn.disabled = hiddenRowsCount < 1;
  }

  if (el.ribbonInsertColLeftBtn) {
    el.ribbonInsertColLeftBtn.disabled = !structure.canInsertColLeft;
  }
  if (el.ribbonInsertColRightBtn) {
    el.ribbonInsertColRightBtn.disabled = !structure.canInsertColRight;
  }
  if (el.ribbonDeleteColBtn) {
    const deletableCols = structure.selectedColumnIndices.length;
    setText(el.ribbonDeleteColBtn, deletableCols > 0 ? `Delete column (${deletableCols})` : "Delete column");
    el.ribbonDeleteColBtn.disabled = !structure.canDeleteCol;
  }
  if (el.ribbonHideColsBtn) {
    setText(el.ribbonHideColsBtn, colIndexes.size > 0 ? `Hide columns (${colIndexes.size})` : "Hide columns");
    el.ribbonHideColsBtn.disabled = colIndexes.size < 1;
  }
  if (el.ribbonUnhideColsBtn) {
    const hiddenColsCount = viewLayout.hiddenCols.size;
    setText(el.ribbonUnhideColsBtn, hiddenColsCount > 0 ? `Unhide columns (${hiddenColsCount})` : "Unhide columns");
    el.ribbonUnhideColsBtn.disabled = hiddenColsCount < 1;
  }
  if (el.ribbonCollapseAllGroupsBtn) {
    setText(
      el.ribbonCollapseAllGroupsBtn,
      activeGroupCount > 0 ? `Collapse all groups (${activeGroupCount})` : "Collapse all groups",
    );
    el.ribbonCollapseAllGroupsBtn.disabled = activeGroupCount < 1;
    el.ribbonCollapseAllGroupsBtn.title = `Spelling-based groups in ${groupScopeLabel}`;
  }
  if (el.ribbonExpandAllGroupsBtn) {
    setText(
      el.ribbonExpandAllGroupsBtn,
      activeGroupCount > 0 ? `Expand all groups (${activeGroupCount})` : "Expand all groups",
    );
    el.ribbonExpandAllGroupsBtn.disabled = activeGroupCount < 1;
    el.ribbonExpandAllGroupsBtn.title = `Spelling-based groups in ${groupScopeLabel}`;
  }
}

function createEmptyInsertedCell(row, sampleCell) {
  const parentTag = row && row.parentElement ? row.parentElement.tagName : "";
  const baseTag = sampleCell && sampleCell.tagName ? sampleCell.tagName : parentTag === "THEAD" ? "TH" : "TD";
  const newCell = document.createElement(baseTag.toLowerCase());
  if (sampleCell) {
    newCell.className = sampleCell.className;
  }
  if (baseTag === "TH" && sampleCell && sampleCell.hasAttribute("scope")) {
    newCell.setAttribute("scope", sampleCell.getAttribute("scope"));
  }
  newCell.textContent = "";
  return newCell;
}

function clearRowRuntimeStyles(row) {
  if (!row) {
    return;
  }
  delete row.dataset.contextHiddenRow;
  row.style.removeProperty("display");
  restoreExcelRowHeight(row);
  for (const cell of Array.from(row.cells || [])) {
    cell.classList.remove(
      ...GRID_SELECTION_CLASSNAMES,
      "sticky-col",
      "sticky-col-boundary",
      "sticky-col-head",
      "sticky-row",
      "sticky-row-boundary",
      "split-head-cell",
    );
    delete cell.dataset.contextHiddenCol;
    delete cell.dataset.gridRow;
    delete cell.dataset.gridRowEnd;
    delete cell.dataset.gridRowSpan;
    delete cell.dataset.gridColStart;
    delete cell.dataset.gridColEnd;
    delete cell.dataset.gridColSpan;
    delete cell.dataset.colStart;
    delete cell.dataset.colSpan;
    delete cell.dataset.formula;
    cell.style.removeProperty("display");
    cell.style.removeProperty("left");
    cell.style.removeProperty("top");
    cell.style.removeProperty("z-index");
    cell.style.removeProperty("background-color");
    cell.style.removeProperty("height");
    cell.style.removeProperty("min-height");
    cell.style.removeProperty("box-sizing");
    cell.textContent = "";
  }
}

function applyStructuralEditForScope(scope, statusMessage) {
  const normalized = normalizeViewScope(scope);
  const freezeDefaults = freezeDefaultsForScope(normalized);
  applyScopeLayoutOverrides(normalized, freezeDefaults.cols, freezeDefaults.rows);
  clearGridSelectionModel();
  setActiveViewScope(normalized);
  setText(el.statusText, statusMessage);
}

function insertRowFromContext(position) {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const structure = structuralContextState(scope);
  if (!structure.table || !structure.point) {
    setText(el.statusText, "Select a data row first.");
    return;
  }
  if (position === "above" && !structure.canInsertRowAbove) {
    setText(el.statusText, "Insert row is available on data rows only.");
    return;
  }
  if (position === "below" && !structure.canInsertRowBelow) {
    setText(el.statusText, "Insert row is available on data rows only.");
    return;
  }

  const targetRow = rowByGridIndex(structure.table, structure.point.row);
  if (!targetRow || !rowIsDataRow(structure.table, targetRow)) {
    setText(el.statusText, "Insert row is available on data rows only.");
    return;
  }

  const section = targetRow.parentElement;
  if (!section) {
    return;
  }

  const insertedRow = targetRow.cloneNode(true);
  clearRowRuntimeStyles(insertedRow);
  if (position === "below") {
    section.insertBefore(insertedRow, targetRow.nextSibling);
  } else {
    section.insertBefore(insertedRow, targetRow);
  }

  const viewLayout = viewLayoutForScope(scope);
  const insertAt = position === "below" ? structure.point.row + 1 : structure.point.row;
  shiftIndexSetForInsert(viewLayout.hiddenRows, insertAt, 1);
  applyStructuralEditForScope(scope, `Inserted 1 row (${position}) · ${selectionScopeLabel(scope)}`);
}

function deleteRowsFromContext() {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const structure = structuralContextState(scope);
  if (!structure.table) {
    setText(el.statusText, "No table available for row delete.");
    return;
  }
  if (!structure.canDeleteRow) {
    setText(el.statusText, "Delete row requires selecting data rows (and keeping at least one row).");
    return;
  }

  const sortedRowsDesc = [...structure.deletableRowIndices].sort((a, b) => b - a);
  for (const rowIndex of sortedRowsDesc) {
    const row = rowByGridIndex(structure.table, rowIndex);
    if (row && row.parentElement) {
      row.parentElement.removeChild(row);
    }
  }

  const viewLayout = viewLayoutForScope(scope);
  shiftIndexSetForDelete(viewLayout.hiddenRows, structure.deletableRowIndices);
  applyStructuralEditForScope(scope, `Deleted ${structure.deletableRowIndices.length} row(s) · ${selectionScopeLabel(scope)}`);
}

function insertColumnFromContext(side) {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const structure = structuralContextState(scope);
  if (!structure.table || !structure.point) {
    setText(el.statusText, "Select a column first.");
    return;
  }

  const canInsert = side === "right" ? structure.canInsertColRight : structure.canInsertColLeft;
  if (!canInsert) {
    setText(el.statusText, "Insert column is available only for non-merged grid sections.");
    return;
  }

  const selectedColumns = structure.selectedColumnIndices.length ? structure.selectedColumnIndices : [structure.point.col];
  const anchorCol =
    side === "right"
      ? Math.max(...selectedColumns.map((value) => Number.parseInt(String(value), 10) || 0))
      : Math.min(...selectedColumns.map((value) => Number.parseInt(String(value), 10) || 0));
  const insertAt = side === "right" ? anchorCol + 1 : anchorCol;

  for (const row of Array.from(structure.table.rows || [])) {
    const cells = Array.from(row.cells || []);
    const sampleIndex = side === "right" ? insertAt - 1 : insertAt;
    const safeSampleIndex = Math.max(0, Math.min(cells.length - 1, sampleIndex));
    const sampleCell = cells.length ? cells[safeSampleIndex] : null;
    const newCell = createEmptyInsertedCell(row, sampleCell);
    const refCell = cells[insertAt] || null;
    row.insertBefore(newCell, refCell);
  }

  const viewLayout = viewLayoutForScope(scope);
  shiftIndexSetForInsert(viewLayout.hiddenCols, insertAt, 1);
  viewLayout.columnWidths = shiftColumnWidthMapForInsert(viewLayout.columnWidths, insertAt, 1);
  applyStructuralEditForScope(scope, `Inserted 1 column (${side}) · ${selectionScopeLabel(scope)}`);
}

function deleteColumnsFromContext() {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const structure = structuralContextState(scope);
  if (!structure.table) {
    setText(el.statusText, "No table available for column delete.");
    return;
  }
  if (!structure.canDeleteCol) {
    setText(el.statusText, "Delete column is available only for non-merged grid sections.");
    return;
  }

  const sortedColsDesc = [...structure.selectedColumnIndices].sort((a, b) => b - a);
  for (const row of Array.from(structure.table.rows || [])) {
    const cells = Array.from(row.cells || []);
    for (const colIndex of sortedColsDesc) {
      if (colIndex >= 0 && colIndex < cells.length) {
        const targetCell = cells[colIndex];
        if (targetCell && targetCell.parentElement === row) {
          row.removeChild(targetCell);
        }
      }
    }
  }

  const viewLayout = viewLayoutForScope(scope);
  shiftIndexSetForDelete(viewLayout.hiddenCols, structure.selectedColumnIndices);
  viewLayout.columnWidths = shiftColumnWidthMapForDelete(viewLayout.columnWidths, structure.selectedColumnIndices);
  applyStructuralEditForScope(scope, `Deleted ${structure.selectedColumnIndices.length} column(s) · ${selectionScopeLabel(scope)}`);
}

function applyFreezeOverride(scope, freezeCols, freezeRows, statusLabel) {
  const normalized = normalizeViewScope(scope);
  const viewLayout = viewLayoutForScope(normalized);
  viewLayout.frozenRowsOverride = Math.max(0, Number.parseInt(String(freezeRows || 0), 10) || 0);
  viewLayout.frozenColsOverride = Math.max(0, Number.parseInt(String(freezeCols || 0), 10) || 0);

  const applied = applyScopeLayoutOverrides(normalized, viewLayout.frozenColsOverride, viewLayout.frozenRowsOverride);
  if (!applied) {
    setText(el.statusText, "No table available to apply freeze.");
    return false;
  }

  if (statusLabel === "unfreeze") {
    setText(el.statusText, `Panes unfrozen · ${selectionScopeLabel(normalized)}`);
  } else {
    setText(
      el.statusText,
      `${statusLabel} · ${selectionScopeLabel(normalized)} · rows ${viewLayout.frozenRowsOverride}, columns ${viewLayout.frozenColsOverride}`,
    );
  }
  return true;
}

function freezePanesFromContext() {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const point = state.contextMenuPoint || state.selectionFocus;
  if (!point) {
    setText(el.statusText, "Select a cell first.");
    return;
  }

  const freezeRows = Math.max(0, Number.parseInt(String(point.row || 0), 10) || 0);
  const freezeCols = Math.max(0, Number.parseInt(String(point.col || 0), 10) || 0);
  applyFreezeOverride(scope, freezeCols, freezeRows, "Freeze panes applied");
}

function freezeTopRowFromContext() {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const bounds = boundsForScope(scope);
  if (bounds.maxRow < 0) {
    setText(el.statusText, "No rows available to freeze.");
    return;
  }
  applyFreezeOverride(scope, 0, 1, "Freeze top row applied");
}

function freezeFirstColumnFromContext() {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const bounds = boundsForScope(scope);
  if (bounds.maxCol < 0) {
    setText(el.statusText, "No columns available to freeze.");
    return;
  }
  applyFreezeOverride(scope, 1, 0, "Freeze first column applied");
}

function unfreezePanesFromContext() {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  applyFreezeOverride(scope, 0, 0, "unfreeze");
}

function updateHiddenRowsFromContext(hide) {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const viewLayout = viewLayoutForScope(scope);
  const rowIndexes = collectContextRowIndices(scope);

  if (hide) {
    if (!rowIndexes.size) {
      setText(el.statusText, "Select row cells first.");
      return;
    }
    for (const rowIndex of rowIndexes) {
      viewLayout.hiddenRows.add(rowIndex);
    }
    applyScopeVisibilityOverrides(scope);
    if (state.selectionScope === scope) {
      refreshGridSelectionVisuals();
    }
    setText(el.statusText, `Hidden ${rowIndexes.size} row(s) · ${selectionScopeLabel(scope)}`);
    return;
  }

  if (rowIndexes.size) {
    let removed = 0;
    for (const rowIndex of rowIndexes) {
      if (viewLayout.hiddenRows.delete(rowIndex)) {
        removed += 1;
      }
    }
    if (removed > 0) {
      setText(el.statusText, `Unhidden selected row(s) · ${selectionScopeLabel(scope)}`);
    } else {
      viewLayout.hiddenRows.clear();
      setText(el.statusText, `Unhidden all rows · ${selectionScopeLabel(scope)}`);
    }
  } else {
    viewLayout.hiddenRows.clear();
    setText(el.statusText, `Unhidden all rows · ${selectionScopeLabel(scope)}`);
  }
  applyScopeVisibilityOverrides(scope);
  if (state.selectionScope === scope) {
    refreshGridSelectionVisuals();
  }
}

function updateHiddenColumnsFromContext(hide) {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const viewLayout = viewLayoutForScope(scope);
  const colIndexes = collectContextColumnIndices(scope);

  if (hide) {
    if (!colIndexes.size) {
      setText(el.statusText, "Select column cells first.");
      return;
    }
    for (const colIndex of colIndexes) {
      viewLayout.hiddenCols.add(colIndex);
    }
    applyScopeVisibilityOverrides(scope);
    if (state.selectionScope === scope) {
      refreshGridSelectionVisuals();
    }
    setText(el.statusText, `Hidden ${colIndexes.size} column(s) · ${selectionScopeLabel(scope)}`);
    return;
  }

  if (colIndexes.size) {
    let removed = 0;
    for (const colIndex of colIndexes) {
      if (viewLayout.hiddenCols.delete(colIndex)) {
        removed += 1;
      }
    }
    if (removed > 0) {
      setText(el.statusText, `Unhidden selected column(s) · ${selectionScopeLabel(scope)}`);
    } else {
      viewLayout.hiddenCols.clear();
      setText(el.statusText, `Unhidden all columns · ${selectionScopeLabel(scope)}`);
    }
  } else {
    viewLayout.hiddenCols.clear();
    setText(el.statusText, `Unhidden all columns · ${selectionScopeLabel(scope)}`);
  }
  applyScopeVisibilityOverrides(scope);
  if (state.selectionScope === scope) {
    refreshGridSelectionVisuals();
  }
}

function updateHiddenRowsAndColumnsFromContext(hide) {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope);
  const viewLayout = viewLayoutForScope(scope);
  const rowIndexes = collectContextRowIndices(scope);
  const colIndexes = collectContextColumnIndices(scope);

  if (hide) {
    if (!rowIndexes.size && !colIndexes.size) {
      setText(el.statusText, "Select cells first.");
      return;
    }
    for (const rowIndex of rowIndexes) {
      viewLayout.hiddenRows.add(rowIndex);
    }
    for (const colIndex of colIndexes) {
      viewLayout.hiddenCols.add(colIndex);
    }
    applyScopeVisibilityOverrides(scope);
    if (state.selectionScope === scope) {
      refreshGridSelectionVisuals();
    }
    setText(
      el.statusText,
      `Hidden rows ${rowIndexes.size}, columns ${colIndexes.size} · ${selectionScopeLabel(scope)}`,
    );
    return;
  }

  if (rowIndexes.size || colIndexes.size) {
    let removedRows = 0;
    let removedCols = 0;
    for (const rowIndex of rowIndexes) {
      if (viewLayout.hiddenRows.delete(rowIndex)) {
        removedRows += 1;
      }
    }
    for (const colIndex of colIndexes) {
      if (viewLayout.hiddenCols.delete(colIndex)) {
        removedCols += 1;
      }
    }
    if (removedRows > 0 || removedCols > 0) {
      setText(
        el.statusText,
        `Unhidden selected rows ${removedRows}, columns ${removedCols} · ${selectionScopeLabel(scope)}`,
      );
    } else {
      viewLayout.hiddenRows.clear();
      viewLayout.hiddenCols.clear();
      setText(el.statusText, `Unhidden all rows + columns · ${selectionScopeLabel(scope)}`);
    }
  } else {
    viewLayout.hiddenRows.clear();
    viewLayout.hiddenCols.clear();
    setText(el.statusText, `Unhidden all rows + columns · ${selectionScopeLabel(scope)}`);
  }
  applyScopeVisibilityOverrides(scope);
  if (state.selectionScope === scope) {
    refreshGridSelectionVisuals();
  }
}

function buildSelectionValueMap(scope) {
  const valueMap = new Map();
  for (const table of tablesForScope(scope)) {
    for (const cell of table.querySelectorAll("th, td")) {
      const row = Number.parseInt(cell.dataset.gridRow || "", 10);
      const colStart = Number.parseInt(cell.dataset.gridColStart || "", 10);
      const colEnd = Number.parseInt(cell.dataset.gridColEnd || "", 10);
      if (!Number.isFinite(row) || !Number.isFinite(colStart) || !Number.isFinite(colEnd)) {
        continue;
      }
      const textValue = String(cell.textContent || "").replace(/\s+/g, " ").trim();
      for (let col = colStart; col <= colEnd; col += 1) {
        const key = `${row}:${col}`;
        if (!valueMap.has(key)) {
          valueMap.set(key, textValue);
        }
      }
    }
  }
  return valueMap;
}

async function writeTextToClipboard(text) {
  const normalizedText = String(text || "");
  inMemoryClipboardText = normalizedText;
  if (navigator.clipboard && typeof navigator.clipboard.writeText === "function" && window.isSecureContext) {
    await navigator.clipboard.writeText(normalizedText);
    return;
  }
  const textarea = document.createElement("textarea");
  textarea.value = normalizedText;
  textarea.setAttribute("readonly", "true");
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  document.body.appendChild(textarea);
  textarea.select();
  const copied = document.execCommand("copy");
  textarea.remove();
  if (!copied) {
    throw new Error("Clipboard copy is not available in this browser context.");
  }
}

async function readTextFromClipboard() {
  if (navigator.clipboard && typeof navigator.clipboard.readText === "function" && window.isSecureContext) {
    try {
      const text = await navigator.clipboard.readText();
      if (typeof text === "string" && text.length > 0) {
        inMemoryClipboardText = text;
      }
      return String(text || "");
    } catch {
      // Fall back to in-memory clipboard when read permission is blocked.
    }
  }

  if (inMemoryClipboardText.length > 0) {
    return inMemoryClipboardText;
  }
  throw new Error("Clipboard paste is not available in this browser context.");
}

async function copyCurrentSelection() {
  const scope = state.selectionScope;
  const ranges = Array.isArray(state.selectionRanges) ? state.selectionRanges : [];
  if (!scope || !ranges.length) {
    setText(el.statusText, "Select cells first.");
    return;
  }

  annotateSelectionGridForScope(scope);
  const valueMap = buildSelectionValueMap(scope);
  const blocks = [];
  for (const rawRange of ranges) {
    const range = normalizeSelectionRange(rawRange);
    const rowLines = [];
    for (let row = range.rowStart; row <= range.rowEnd; row += 1) {
      const cells = [];
      for (let col = range.colStart; col <= range.colEnd; col += 1) {
        cells.push(valueMap.get(`${row}:${col}`) || "");
      }
      rowLines.push(cells.join("\t"));
    }
    blocks.push(rowLines.join("\n"));
  }

  await writeTextToClipboard(blocks.join("\n\n"));
  setText(el.statusText, `Copied ${ranges.length} selection area(s).`);
}

function parseClipboardMatrix(text) {
  const normalizedText = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  if (!normalizedText.length) {
    return [];
  }
  const primaryBlock = normalizedText.split(/\n{2,}/)[0] || "";
  if (!primaryBlock.length) {
    return [];
  }
  const lines = primaryBlock.split("\n");
  if (lines.length > 1 && lines[lines.length - 1] === "") {
    lines.pop();
  }
  if (!lines.length) {
    return [];
  }
  return lines.map((line) => line.split("\t"));
}

function buildPointCellMap(scope) {
  const pointCellMap = new Map();
  for (const table of tablesForScope(scope)) {
    annotateSelectionGridForTable(table);
    for (const cell of table.querySelectorAll("th, td")) {
      const rowStart = Number.parseInt(cell.dataset.gridRow || "", 10);
      const rowEnd = Number.parseInt(cell.dataset.gridRowEnd || cell.dataset.gridRow || "", 10);
      const colStart = Number.parseInt(cell.dataset.gridColStart || "", 10);
      const colEnd = Number.parseInt(cell.dataset.gridColEnd || "", 10);
      if (
        !Number.isFinite(rowStart) ||
        !Number.isFinite(rowEnd) ||
        !Number.isFinite(colStart) ||
        !Number.isFinite(colEnd)
      ) {
        continue;
      }
      for (let row = rowStart; row <= rowEnd; row += 1) {
        for (let col = colStart; col <= colEnd; col += 1) {
          const key = `${row}:${col}`;
          if (!pointCellMap.has(key)) {
            pointCellMap.set(key, cell);
          }
        }
      }
    }
  }
  return pointCellMap;
}

function pastePointForScope(scope) {
  const normalized = normalizeViewScope(scope);
  if (state.contextMenuScope === normalized && state.contextMenuPoint) {
    return { ...state.contextMenuPoint };
  }
  if (state.selectionScope === normalized && state.selectionFocus) {
    return { ...state.selectionFocus };
  }
  if (state.selectionScope === normalized && state.selectionAnchor) {
    return { ...state.selectionAnchor };
  }
  if (state.selectionScope === normalized && Array.isArray(state.selectionRanges) && state.selectionRanges.length) {
    const firstRange = normalizeSelectionRange(state.selectionRanges[0]);
    return { row: firstRange.rowStart, col: firstRange.colStart };
  }
  return null;
}

function formulaDisplayForCell(cell) {
  if (!cell) {
    return "";
  }
  const formula = String(cell.dataset.formula || "").trim();
  if (formula.length) {
    return formula;
  }
  return String(cell.textContent || "").replace(/\s+/g, " ").trim();
}

function formatFormulaResult(value) {
  if (Array.isArray(value)) {
    return formatFormulaResult(value.length ? value[0] : "");
  }
  if (value === null || value === undefined || value === "") {
    return "";
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      return "#NUM!";
    }
    if (Number.isInteger(value)) {
      return String(value);
    }
    return String(Number.parseFloat(value.toFixed(12)));
  }
  return String(value);
}

function focusCellForScope(scope, pointCellMap = null) {
  const point = pastePointForScope(scope);
  if (!point) {
    return null;
  }
  const map = pointCellMap || buildPointCellMap(scope);
  return map.get(`${point.row}:${point.col}`) || null;
}

function collectFormulaTargetCells(scope, pointCellMap) {
  const targets = [];
  const seen = new Set();
  const map = pointCellMap || buildPointCellMap(scope);
  const ranges =
    state.selectionScope === scope && Array.isArray(state.selectionRanges) ? state.selectionRanges : [];

  if (ranges.length) {
    for (const rawRange of ranges) {
      const range = normalizeSelectionRange(rawRange);
      for (let row = range.rowStart; row <= range.rowEnd; row += 1) {
        for (let col = range.colStart; col <= range.colEnd; col += 1) {
          const cell = map.get(`${row}:${col}`);
          if (!cell || seen.has(cell)) {
            continue;
          }
          seen.add(cell);
          targets.push(cell);
        }
      }
    }
  }

  if (!targets.length) {
    const focus = focusCellForScope(scope, map);
    if (focus) {
      targets.push(focus);
    }
  }

  return targets;
}

function jumpToAddressFromNameBox() {
  if (!el.formulaNameBox) {
    return;
  }
  const requested = String(el.formulaNameBox.value || "").trim();
  const point = parseExcelAddress(requested);
  if (!point) {
    setText(el.statusText, `Invalid address "${requested}". Use format like A1.`);
    return;
  }

  const scope = normalizeViewScope(state.selectionScope || state.activeViewScope || "main");
  const bounds = boundsForScope(scope);
  if (
    !Number.isFinite(bounds.maxRow) ||
    !Number.isFinite(bounds.maxCol) ||
    point.row < 0 ||
    point.col < 0 ||
    point.row > bounds.maxRow ||
    point.col > bounds.maxCol
  ) {
    setText(el.statusText, `Address ${requested.toUpperCase()} is outside the visible ${selectionScopeLabel(scope)}.`);
    return;
  }

  selectSingleGridPoint(scope, point);
  const pointMap = buildPointCellMap(scope);
  const cell = pointMap.get(`${point.row}:${point.col}`) || null;
  if (cell && typeof cell.scrollIntoView === "function") {
    cell.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  setText(el.statusText, `Selected ${pointToAddress(point)} · ${selectionScopeLabel(scope)}.`);
}

function cancelFormulaEdit() {
  const scope = normalizeViewScope(state.selectionScope || state.activeViewScope || "main");
  const pointMap = buildPointCellMap(scope);
  const focusCell = focusCellForScope(scope, pointMap);
  if (el.formulaInput) {
    el.formulaInput.value = formulaDisplayForCell(focusCell);
  }
  if (el.formulaNameBox && state.selectionFocus) {
    el.formulaNameBox.value = pointToAddress(state.selectionFocus);
  }
}

function applyFormulaInput() {
  if (!el.formulaInput) {
    return;
  }
  const rawInput = String(el.formulaInput.value || "");
  const trimmedInput = rawInput.trim();
  const scope = normalizeViewScope(state.selectionScope || state.activeViewScope || "main");
  const pointMap = buildPointCellMap(scope);
  const targetCells = collectFormulaTargetCells(scope, pointMap);

  if (!targetCells.length) {
    setText(el.statusText, "Select a destination cell first.");
    return;
  }

  let renderedValue = "";
  let storedFormula = "";
  if (!trimmedInput.length) {
    renderedValue = "";
    storedFormula = "";
  } else if (trimmedInput.startsWith("=")) {
    const computedValue = evaluateFormulaExpression(trimmedInput, scope, pointMap);
    renderedValue = formatFormulaResult(computedValue);
    storedFormula = trimmedInput;
  } else {
    renderedValue = rawInput;
    storedFormula = "";
  }

  for (const cell of targetCells) {
    cell.textContent = renderedValue;
    if (storedFormula) {
      cell.dataset.formula = storedFormula;
    } else {
      delete cell.dataset.formula;
    }
  }

  const focusPoint = pastePointForScope(scope);
  if (focusPoint) {
    selectSingleGridPoint(scope, focusPoint);
  } else {
    refreshGridSelectionVisuals();
  }

  if (el.formulaInput) {
    el.formulaInput.value = storedFormula || renderedValue;
  }

  const targetCount = targetCells.length;
  if (storedFormula) {
    setText(el.statusText, `Formula applied to ${targetCount} cell${targetCount === 1 ? "" : "s"} (${selectionScopeLabel(scope)}).`);
    return;
  }
  if (!trimmedInput.length) {
    setText(el.statusText, `Cleared ${targetCount} cell${targetCount === 1 ? "" : "s"} (${selectionScopeLabel(scope)}).`);
    return;
  }
  setText(el.statusText, `Updated ${targetCount} cell${targetCount === 1 ? "" : "s"} (${selectionScopeLabel(scope)}).`);
}

async function pasteClipboardIntoGrid() {
  const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope || "main");
  const startPoint = pastePointForScope(scope);
  if (!startPoint) {
    setText(el.statusText, "Select a destination cell first.");
    return;
  }

  const clipboardText = await readTextFromClipboard();
  const matrix = parseClipboardMatrix(clipboardText);
  if (!matrix.length) {
    setText(el.statusText, "Clipboard is empty.");
    return;
  }

  annotateSelectionGridForScope(scope);
  const pointCellMap = buildPointCellMap(scope);
  if (!pointCellMap.size) {
    setText(el.statusText, "No destination cells available for paste.");
    return;
  }

  const assignmentMap = new Map();
  const bounds = boundsForScope(scope);
  const selectionRanges =
    state.selectionScope === scope && Array.isArray(state.selectionRanges) ? state.selectionRanges : [];

  if (matrix.length === 1 && matrix[0].length === 1 && selectionRanges.length) {
    const value = String(matrix[0][0] ?? "");
    for (const rawRange of selectionRanges) {
      const range = normalizeSelectionRange(rawRange);
      for (let row = range.rowStart; row <= range.rowEnd; row += 1) {
        for (let col = range.colStart; col <= range.colEnd; col += 1) {
          const cell = pointCellMap.get(`${row}:${col}`);
          if (cell && !assignmentMap.has(cell)) {
            assignmentMap.set(cell, value);
          }
        }
      }
    }
  } else {
    for (let rowOffset = 0; rowOffset < matrix.length; rowOffset += 1) {
      const sourceRow = matrix[rowOffset] || [];
      for (let colOffset = 0; colOffset < sourceRow.length; colOffset += 1) {
        const targetRow = startPoint.row + rowOffset;
        const targetCol = startPoint.col + colOffset;
        if (targetRow < 0 || targetCol < 0 || targetRow > bounds.maxRow || targetCol > bounds.maxCol) {
          continue;
        }
        const cell = pointCellMap.get(`${targetRow}:${targetCol}`);
        if (cell && !assignmentMap.has(cell)) {
          assignmentMap.set(cell, String(sourceRow[colOffset] ?? ""));
        }
      }
    }
  }

  if (!assignmentMap.size) {
    setText(el.statusText, "Paste target is outside the visible grid.");
    return;
  }

  for (const [cell, value] of assignmentMap) {
    cell.textContent = value;
    delete cell.dataset.formula;
  }

  state.selectionScope = scope;
  if (!Array.isArray(state.selectionRanges) || !state.selectionRanges.length) {
    state.selectionAnchor = { ...startPoint };
    state.selectionFocus = { ...startPoint };
    state.selectionRanges = [rangeFromScopePoints(scope, startPoint, startPoint)];
  }
  refreshGridSelectionVisuals();

  const pastedRows = matrix.length;
  const pastedCols = matrix.reduce((max, row) => Math.max(max, row.length), 0);
  setText(
    el.statusText,
    `Pasted ${pastedRows}x${pastedCols} into ${selectionScopeLabel(scope)} (${assignmentMap.size} cell${
      assignmentMap.size === 1 ? "" : "s"
    }).`,
  );
}

function clearGridSelectionClasses(scope = null) {
  const scopes = scope ? [scope] : ["main", "reference"];
  for (const scopeName of scopes) {
    for (const table of tablesForScope(scopeName)) {
      for (const cell of table.querySelectorAll("th, td")) {
        cell.classList.remove(...GRID_SELECTION_CLASSNAMES);
      }
    }
  }
}

function updateSelectionStatusBar(scope, selectedCells, focusCellText = "", focusCell = null) {
  const ranges = Array.isArray(state.selectionRanges) ? state.selectionRanges : [];
  const focus = state.selectionFocus;

  if (!scope || !ranges.length || !selectedCells.length) {
    resetSelectionStatusBar();
    return;
  }

  const numericValues = [];
  let nonEmptyCount = 0;
  for (const cell of selectedCells) {
    const compactText = String(cell.textContent || "").replace(/\s+/g, " ").trim();
    if (compactText.length > 0) {
      nonEmptyCount += 1;
    }
    const numeric = parseNumericCellValue(cell.textContent || "");
    if (Number.isFinite(numeric)) {
      numericValues.push(numeric);
    }
  }

  const sum = numericValues.reduce((acc, value) => acc + value, 0);
  const avg = numericValues.length ? sum / numericValues.length : NaN;
  const min = numericValues.length
    ? numericValues.reduce((current, value) => (value < current ? value : current), numericValues[0])
    : NaN;
  const max = numericValues.length
    ? numericValues.reduce((current, value) => (value > current ? value : current), numericValues[0])
    : NaN;

  setActiveViewScope(scope, { refresh: false });
  setText(el.statusSelectionScope, statusScopeLabel(scope, true));
  setText(el.statusSelectionAddress, focus ? pointToAddress(focus) : "-");
  const lastRange = ranges[ranges.length - 1] || null;
  const rangeAddress = rangeToAddress(lastRange);
  const extraAreas = ranges.length > 1 ? ` (+${ranges.length - 1} area)` : "";
  setText(el.statusSelectionRange, `${rangeAddress}${extraAreas}`);
  setText(el.statusSelectionCellCount, `Cells: ${selectedCells.length}`);
  setText(el.statusSelectionCount, `CountA: ${nonEmptyCount}`);
  setText(el.statusSelectionNumericCount, `Numbers: ${numericValues.length}`);
  setText(el.statusSelectionSum, `Sum: ${numericValues.length ? formatMetricNumber(sum) : "-"}`);
  setText(el.statusSelectionAvg, `Avg: ${numericValues.length ? formatMetricNumber(avg) : "-"}`);
  setText(el.statusSelectionMin, `Min: ${numericValues.length ? formatMetricNumber(min) : "-"}`);
  setText(el.statusSelectionMax, `Max: ${numericValues.length ? formatMetricNumber(max) : "-"}`);
  if (el.formulaNameBox && document.activeElement !== el.formulaNameBox) {
    el.formulaNameBox.value = focus ? pointToAddress(focus) : "A1";
  }
  if (el.formulaInput && document.activeElement !== el.formulaInput) {
    const formulaValue = focusCell ? String(focusCell.dataset.formula || "").trim() : "";
    el.formulaInput.value = formulaValue || String(focusCellText || "");
  }
  scheduleFloatingLayoutMetricsSync();
}

function refreshGridSelectionVisuals() {
  const scope = state.selectionScope;
  const ranges = Array.isArray(state.selectionRanges) ? state.selectionRanges : [];
  const normalizedRanges = ranges.map((range) => normalizeSelectionRange(range));
  clearGridSelectionClasses();

  if (!scope || !normalizedRanges.length) {
    resetSelectionStatusBar();
    syncRibbonGridControlState();
    return;
  }

  annotateSelectionGridForScope(scope);
  const selectedCells = [];
  const focus = state.selectionFocus;
  let focusCell = null;
  let focusCellText = "";

  for (const table of tablesForScope(scope)) {
    for (const cell of table.querySelectorAll("th, td")) {
      const rowStart = Number.parseInt(cell.dataset.gridRow || "", 10);
      const rowEnd = Number.parseInt(cell.dataset.gridRowEnd || cell.dataset.gridRow || "", 10);
      const colStart = Number.parseInt(cell.dataset.gridColStart || "", 10);
      const colEnd = Number.parseInt(cell.dataset.gridColEnd || "", 10);
      if (
        !Number.isFinite(rowStart) ||
        !Number.isFinite(rowEnd) ||
        !Number.isFinite(colStart) ||
        !Number.isFinite(colEnd)
      ) {
        continue;
      }

      const isMergedCell = rowEnd > rowStart || colEnd > colStart;
      const selected = normalizedRanges.some((range) => {
        if (isMergedCell) {
          return rangeFullyContainsCell(range, rowStart, rowEnd, colStart, colEnd);
        }
        return rangeIncludesCell(range, rowStart, rowEnd, colStart, colEnd);
      });
      if (!selected) {
        continue;
      }
      cell.classList.add("grid-cell-selected");
      const edgeClasses = selectionEdgeClassesForCell(rowStart, rowEnd, colStart, colEnd, normalizedRanges);
      if (edgeClasses.length) {
        cell.classList.add(...edgeClasses);
      }
      selectedCells.push(cell);

      if (focus && focus.row >= rowStart && focus.row <= rowEnd && focus.col >= colStart && focus.col <= colEnd) {
        cell.classList.add("grid-cell-active");
        if (!focusCell) {
          focusCell = cell;
        }
        if (!focusCellText) {
          focusCellText = String(cell.textContent || "").replace(/\s+/g, " ").trim();
        }
      }
    }
  }

  updateSelectionStatusBar(scope, selectedCells, focusCellText, focusCell);
  syncRibbonGridControlState();
}

function clearGridSelectionModel() {
  state.selectionScope = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionFocus = null;
  state.selectionDrag = null;
  clearGridSelectionClasses();
  resetSelectionStatusBar();
  hideGridContextMenu();
  syncRibbonGridControlState();
}

function applySelectionDragPoint(point) {
  if (!state.selectionDrag || !point) {
    return;
  }
  const drag = state.selectionDrag;
  state.selectionScope = drag.scope;
  state.selectionFocus = point;
  state.selectionRanges = [...drag.baseRanges, rangeFromScopePoints(drag.scope, drag.anchor, point)];
  refreshGridSelectionVisuals();
}

function startGridSelection(event) {
  if (event.button !== 0) {
    return;
  }
  if (event.target instanceof HTMLElement && event.target.closest(".row-group-toggle")) {
    return;
  }
  const cell = event.target instanceof HTMLElement ? event.target.closest("th, td") : null;
  if (!cell) {
    return;
  }
  const scope = selectionScopeForCell(cell);
  if (!scope) {
    return;
  }
  setActiveViewScope(scope, { refresh: false });
  annotateSelectionGridForScope(scope);
  const point = selectionPointForCell(cell);
  if (!point) {
    return;
  }

  const sameScope = state.selectionScope === scope;
  const additive = Boolean(event.ctrlKey || event.metaKey);
  const extend = Boolean(event.shiftKey && sameScope && state.selectionAnchor);
  const anchor = extend ? state.selectionAnchor : point;
  const baseRanges = additive && sameScope ? [...(state.selectionRanges || [])] : [];

  state.selectionScope = scope;
  state.selectionAnchor = anchor;
  state.selectionFocus = point;
  state.selectionRanges = [...baseRanges, rangeFromScopePoints(scope, anchor, point)];
  state.selectionDrag = {
    scope,
    anchor,
    baseRanges,
  };

  event.preventDefault();
  refreshGridSelectionVisuals();
}

function moveGridSelection(event) {
  if (!state.selectionDrag || !(event.buttons & 1)) {
    return;
  }
  const cell = event.target instanceof HTMLElement ? event.target.closest("th, td") : null;
  if (!cell) {
    return;
  }
  const scope = selectionScopeForCell(cell);
  if (!scope || scope !== state.selectionDrag.scope) {
    return;
  }
  const point = selectionPointForCell(cell);
  applySelectionDragPoint(point);
}

function endGridSelection() {
  state.selectionDrag = null;
}

function filteredRows(rows) {
  const q = el.searchInput.value.trim().toLowerCase();
  if (!q) {
    return rows;
  }
  return rows.filter((row) => String(row.product_name || "").toLowerCase().includes(q));
}

const PRODUCT_GROUP_STOPWORDS = new Set([
  "vodka",
  "whisky",
  "whiskey",
  "beer",
  "rum",
  "gin",
  "wine",
  "brandy",
  "bottle",
  "liter",
  "litre",
  "liters",
  "litres",
  "ml",
  "pk",
]);
const PRODUCT_GROUP_PUNCT_RE = /["'`’“”()[\]{}<>.,:;!?/\\|_*+=~-]+/g;
const PRODUCT_GROUP_UNIT_RE = /\b\d+(?:\.\d+)?\s*(?:ml|l|liter|litre|liters|litres|cc)\b/gi;
const PRODUCT_GROUP_NUMBER_RE = /\b\d+(?:\.\d+)?\b/g;

function productTokensForGrouping(value) {
  const normalized = String(value || "")
    .toLowerCase()
    .replace(PRODUCT_GROUP_PUNCT_RE, " ")
    .replace(PRODUCT_GROUP_UNIT_RE, " ")
    .replace(PRODUCT_GROUP_NUMBER_RE, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!normalized) {
    return [];
  }
  const filteredTokens = normalized
    .split(" ")
    .map((token) => token.trim())
    .filter((token) => token.length > 1 && !PRODUCT_GROUP_STOPWORDS.has(token));
  if (filteredTokens.length) {
    return filteredTokens;
  }
  return normalized
    .split(" ")
    .map((token) => token.trim())
    .filter(Boolean);
}

function productGroupKey(value) {
  const tokens = productTokensForGrouping(value);
  return tokens.length ? tokens[0] : "";
}

function productGroupLabel(value, fallback) {
  const source = String(value || "").replace(PRODUCT_GROUP_PUNCT_RE, " ").trim();
  if (!source) {
    return fallback || "Group";
  }
  const firstToken = source.split(/\s+/).find(Boolean);
  return firstToken || fallback || "Group";
}

function groupedReferenceRows(rows) {
  const groups = [];
  const groupByKey = new Map();
  let looseCounter = 0;
  for (const row of rows) {
    const rootKey = productGroupKey(row.product_name);
    const groupKey = rootKey || `row-${looseCounter++}`;
    let group = groupByKey.get(groupKey);
    if (!group) {
      group = {
        key: groupKey,
        label: productGroupLabel(row.product_name, groupKey),
        rows: [],
      };
      groupByKey.set(groupKey, group);
      groups.push(group);
    }
    group.rows.push(row);
  }
  return groups;
}

function rowGroupStateKey(scope, sheetName, groupKey) {
  return `${scope || "reference"}::${sheetName || "-"}::${groupKey || "-"}`;
}

function isRowGroupCollapsed(scope, sheetName, groupKey) {
  if (!state.collapsedRowGroups || !(state.collapsedRowGroups instanceof Set)) {
    return false;
  }
  return state.collapsedRowGroups.has(rowGroupStateKey(scope, sheetName, groupKey));
}

function setRowGroupCollapsed(scope, sheetName, groupKey, collapsed) {
  if (!state.collapsedRowGroups || !(state.collapsedRowGroups instanceof Set)) {
    state.collapsedRowGroups = new Set();
  }
  const key = rowGroupStateKey(scope, sheetName, groupKey);
  if (collapsed) {
    state.collapsedRowGroups.add(key);
    return;
  }
  state.collapsedRowGroups.delete(key);
}

function applyRowGroupVisibility(table, groupKey, collapsed) {
  if (!table || !groupKey) {
    return;
  }
  const rows = table.tBodies.length ? Array.from(table.tBodies[0].rows) : Array.from(table.rows || []);
  for (const row of rows) {
    if (!(row instanceof HTMLElement)) {
      continue;
    }
    if ((row.dataset.groupKey || "") !== groupKey) {
      continue;
    }
    if ((row.dataset.groupPrimary || "") === "1") {
      continue;
    }
    row.classList.toggle("row-group-hidden", collapsed);
  }
}

function syncRowGroupToggleVisual(toggleBtn, collapsed) {
  if (!(toggleBtn instanceof HTMLElement)) {
    return;
  }
  toggleBtn.setAttribute("aria-expanded", collapsed ? "false" : "true");
  const caret = toggleBtn.querySelector(".row-group-caret");
  if (caret) {
    setText(caret, collapsed ? "▸" : "▾");
  }
}

function applyRowGroupStateForScope(scope, groupKey, collapsed) {
  const normalizedScope = normalizeViewScope(scope);
  for (const table of tablesForScope(normalizedScope)) {
    applyRowGroupVisibility(table, groupKey, collapsed);
    const toggleButtons = table.querySelectorAll(".row-group-toggle");
    for (const button of Array.from(toggleButtons)) {
      if (!(button instanceof HTMLElement)) {
        continue;
      }
      if ((button.dataset.groupKey || "") !== groupKey) {
        continue;
      }
      syncRowGroupToggleVisual(button, collapsed);
    }
  }
}

function collectRowGroupContexts(scope = "reference") {
  const normalizedScope = normalizeViewScope(scope);
  const contexts = [];
  const seen = new Set();
  for (const table of tablesForScope(normalizedScope)) {
    const toggleButtons = table.querySelectorAll(".row-group-toggle");
    for (const button of Array.from(toggleButtons)) {
      if (!(button instanceof HTMLElement)) {
        continue;
      }
      const groupKey = String(button.dataset.groupKey || "").trim();
      if (!groupKey) {
        continue;
      }
      const groupScope = normalizeViewScope(String(button.dataset.groupScope || normalizedScope).trim() || normalizedScope);
      const fallbackSheet = groupScope === "reference" ? state.selectedReferenceSheetName : state.selectedMainSheetName;
      const sheetName = String(button.dataset.groupSheet || fallbackSheet || "").trim();
      const contextKey = rowGroupStateKey(groupScope, sheetName, groupKey);
      if (seen.has(contextKey)) {
        continue;
      }
      seen.add(contextKey);
      contexts.push({ scope: groupScope, sheetName, groupKey });
    }
  }
  return contexts;
}

function setAllRowGroupsCollapsed(scope = "reference", collapsed = true) {
  const groupContexts = collectRowGroupContexts(scope);
  for (const context of groupContexts) {
    setRowGroupCollapsed(context.scope, context.sheetName, context.groupKey, collapsed);
    applyRowGroupStateForScope(context.scope, context.groupKey, collapsed);
  }
  return groupContexts.length;
}

function normalizeGroupingHeaderText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function findRowCellAtGridColumn(row, gridCol) {
  if (!row || !Number.isFinite(Number(gridCol))) {
    return null;
  }
  for (const cell of Array.from(row.cells || [])) {
    const colStart = Number.parseInt(cell.dataset.gridColStart || "", 10);
    const colEnd = Number.parseInt(cell.dataset.gridColEnd || "", 10);
    if (!Number.isFinite(colStart) || !Number.isFinite(colEnd)) {
      continue;
    }
    if (gridCol >= colStart && gridCol <= colEnd) {
      return cell;
    }
  }
  return null;
}

function detectMainProductColumn(mainTable) {
  if (!(mainTable instanceof HTMLElement)) {
    return null;
  }
  annotateSelectionGridForTable(mainTable);
  const rows = Array.from(mainTable.rows || []);
  if (!rows.length) {
    return null;
  }

  let bestMatch = null;
  const scanLimit = Math.min(rows.length, 18);
  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
    const row = rows[rowIndex];
    for (const cell of Array.from(row.cells || [])) {
      const token = normalizeGroupingHeaderText(cell.textContent || "");
      if (!token) {
        continue;
      }
      let score = 0;
      if (token.includes("productname")) {
        score = 5;
      } else if (token.includes("product")) {
        score = 2;
      }
      if (score < 1) {
        continue;
      }
      const colStart = Number.parseInt(cell.dataset.gridColStart || "", 10);
      const gridRow = Number.parseInt(cell.dataset.gridRow || "", 10);
      const gridRowEnd = Number.parseInt(cell.dataset.gridRowEnd || "", 10);
      if (!Number.isFinite(colStart) || !Number.isFinite(gridRow)) {
        continue;
      }
      const headerRowEnd = Number.isFinite(gridRowEnd) ? gridRowEnd : gridRow;
      const candidate = {
        score,
        col: colStart,
        dataStartRow: headerRowEnd + 1,
        headerRow: gridRow,
      };
      if (
        !bestMatch ||
        candidate.score > bestMatch.score ||
        (candidate.score === bestMatch.score && candidate.headerRow < bestMatch.headerRow)
      ) {
        bestMatch = candidate;
      }
    }
  }

  if (!bestMatch || !Number.isFinite(bestMatch.col)) {
    return null;
  }
  return {
    col: Math.max(0, bestMatch.col),
    dataStartRow: Math.max(0, Number.parseInt(String(bestMatch.dataStartRow || 0), 10) || 0),
  };
}

function applyMainRowGrouping(mainHost, sheetName) {
  const mainTable =
    mainHost instanceof HTMLElement ? mainHost.querySelector("table.main-source-table") : null;
  if (!(mainTable instanceof HTMLTableElement)) {
    return 0;
  }

  const productColumn = detectMainProductColumn(mainTable);
  if (!productColumn) {
    return 0;
  }

  const allRows = Array.from(mainTable.rows || []);
  if (!allRows.length) {
    return 0;
  }

  const groups = [];
  const groupsByKey = new Map();
  let looseCounter = 0;
  for (let rowIndex = productColumn.dataStartRow; rowIndex < allRows.length; rowIndex += 1) {
    const row = allRows[rowIndex];
    if (!(row instanceof HTMLTableRowElement)) {
      continue;
    }
    const productCell = findRowCellAtGridColumn(row, productColumn.col);
    if (!productCell) {
      continue;
    }
    const productText = String(productCell.textContent || "")
      .replace(/\s+/g, " ")
      .trim();
    if (!productText) {
      continue;
    }
    const rootKey = productGroupKey(productText);
    const groupKey = rootKey || `main-row-${looseCounter++}`;
    let group = groupsByKey.get(groupKey);
    if (!group) {
      group = {
        key: groupKey,
        label: productGroupLabel(productText, groupKey),
        members: [],
      };
      groupsByKey.set(groupKey, group);
      groups.push(group);
    }
    group.members.push({
      row,
      productCell,
      productText,
      productHtml: productCell.innerHTML,
    });
  }

  let groupedRows = 0;
  for (const group of groups) {
    if (!group || !Array.isArray(group.members) || group.members.length < 2) {
      continue;
    }
    const collapsed = isRowGroupCollapsed("main", sheetName, group.key);
    const groupKeyEscaped = escapeHtml(group.key);
    const groupLabelEscaped = escapeHtml(group.label);
    for (let idx = 0; idx < group.members.length; idx += 1) {
      const member = group.members[idx];
      if (!member || !(member.row instanceof HTMLElement) || !(member.productCell instanceof HTMLElement)) {
        continue;
      }
      const isPrimary = idx === 0;
      member.row.classList.add("row-group-member");
      member.row.dataset.groupKey = group.key;
      member.row.dataset.groupPrimary = isPrimary ? "1" : "0";
      member.productCell.classList.add("row-group-product-cell");
      const productTitle = escapeHtml(member.productText);
      if (isPrimary) {
        const arrow = collapsed ? "▸" : "▾";
        member.productCell.innerHTML =
          `<div class="row-group-cell">` +
          `<button type="button" class="row-group-toggle" data-group-key="${groupKeyEscaped}" data-group-scope="main" data-group-sheet="${escapeHtml(sheetName || "")}" aria-expanded="${collapsed ? "false" : "true"}">` +
          `<span class="row-group-caret">${arrow}</span>` +
          `<span class="row-group-title">${groupLabelEscaped}</span>` +
          `<span class="row-group-count">(${group.members.length})</span>` +
          `</button>` +
          `<span class="row-group-product" title="${productTitle}">${member.productHtml}</span>` +
          `</div>`;
      } else {
        member.row.classList.add("row-group-child");
        if (collapsed) {
          member.row.classList.add("row-group-hidden");
        }
        member.productCell.innerHTML = `<span class="row-group-product row-group-product-child" title="${productTitle}">${member.productHtml}</span>`;
      }
      groupedRows += 1;
    }
  }
  return groupedRows;
}

function monthYearFromEntry(monthEntry) {
  if (!monthEntry || typeof monthEntry !== "object") {
    return null;
  }
  const directYear = Number.parseInt(String(monthEntry.year ?? ""), 10);
  if (Number.isFinite(directYear) && directYear >= 1900) {
    return directYear;
  }
  const keyMatch = String(monthEntry.key || "").match(/^(20\d{2})-/);
  if (keyMatch) {
    const year = Number.parseInt(keyMatch[1], 10);
    if (Number.isFinite(year)) {
      return year;
    }
  }
  const labelMatch = String(monthEntry.label || "").match(/(20\d{2})/);
  if (labelMatch) {
    const year = Number.parseInt(labelMatch[1], 10);
    if (Number.isFinite(year)) {
      return year;
    }
  }
  return null;
}

function rowMetricValueByMonthKey(row, monthKey, metricName) {
  if (!row || !row.values || !monthKey || !metricName) {
    return null;
  }
  const monthCell = row.values[monthKey];
  if (!monthCell || typeof monthCell !== "object") {
    return null;
  }
  const rawValue = monthCell[metricName];
  if (typeof rawValue === "number" && Number.isFinite(rawValue)) {
    return rawValue;
  }
  const parsed = parseNumericCellValue(rawValue);
  return Number.isFinite(parsed) ? parsed : null;
}

function summarizeRowMetricByMonths(row, monthKeys, metricName) {
  const numericValues = [];
  for (const monthKey of Array.isArray(monthKeys) ? monthKeys : []) {
    const value = rowMetricValueByMonthKey(row, monthKey, metricName);
    if (Number.isFinite(value)) {
      numericValues.push(value);
    }
  }
  if (!numericValues.length) {
    return { total: null, avg: null };
  }
  const total = numericValues.reduce((sum, value) => sum + value, 0);
  const avg = total / numericValues.length;
  const round = (value) => Math.round(value * 10000) / 10000;
  return { total: round(total), avg: round(avg) };
}

function renderTable(target, sheetData, selectedMonths, scope = "reference") {
  const metric = el.metricSelect.value;
  const rows = filteredRows(sheetData.rows);

  if (!rows.length) {
    target.innerHTML =
      '<tbody><tr><td class="empty">No rows match current filter.</td></tr></tbody>';
    return { yearSummaryColumns: 0, yearSummaryYears: 0 };
  }

  const columns = [
    { key: "sr", label: "Sr", cls: "right" },
    { key: "product_name", label: "Product Name", cls: "left" },
    { key: "ml", label: "Ml", cls: "right" },
    { key: "packing", label: "Packing", cls: "left" },
  ];

  const normalizedScope = normalizeViewScope(scope);
  const selectedMonthNumbers = new Set(
    selectedMonths
      .map((month) => Number.parseInt(String(month && month.month !== undefined ? month.month : ""), 10))
      .filter((monthNum) => Number.isFinite(monthNum) && monthNum >= 1 && monthNum <= 12),
  );
  const isReferenceYearSummaryMode =
    normalizedScope === "reference" &&
    metric !== "all" &&
    modeIsMultiMonthYears(el.refModeSelect.value) &&
    selectedMonths.length >= 2 &&
    selectedMonthNumbers.size >= 2;

  const dynamicColumns = [];
  const summaryYears = [];
  let currentYear = null;
  let currentYearMonthKeys = [];
  const flushYearSummaryColumns = () => {
    if (!isReferenceYearSummaryMode) {
      return;
    }
    if (!Number.isFinite(currentYear) || !currentYearMonthKeys.length) {
      return;
    }
    dynamicColumns.push(
      {
        kind: "year_total",
        year: currentYear,
        metric,
        monthKeys: [...currentYearMonthKeys],
        label: `${currentYear} Total`,
      },
      {
        kind: "year_avg",
        year: currentYear,
        metric,
        monthKeys: [...currentYearMonthKeys],
        label: `${currentYear} Avg`,
      },
    );
    summaryYears.push(currentYear);
  };

  for (const month of selectedMonths) {
    const monthYear = monthYearFromEntry(month);
    if (isReferenceYearSummaryMode && Number.isFinite(currentYear) && monthYear !== currentYear) {
      flushYearSummaryColumns();
      currentYearMonthKeys = [];
      currentYear = monthYear;
    }
    if (isReferenceYearSummaryMode && !Number.isFinite(currentYear)) {
      currentYear = monthYear;
    }

    if (metric === "all") {
      dynamicColumns.push(
        { kind: "month", monthKey: month.key, metric: "pk", label: `${month.label} PK` },
        { kind: "month", monthKey: month.key, metric: "bottle", label: `${month.label} Bottle` },
        { kind: "month", monthKey: month.key, metric: "liter", label: `${month.label} Liter` },
      );
    } else {
      dynamicColumns.push({
        kind: "month",
        monthKey: month.key,
        metric,
        label: month.label,
      });
      if (isReferenceYearSummaryMode && Number.isFinite(monthYear)) {
        currentYearMonthKeys.push(month.key);
      }
    }
  }
  flushYearSummaryColumns();
  const uniqueSummaryYears = [...new Set(summaryYears)];

  if (!dynamicColumns.length) {
    target.innerHTML =
      '<tbody><tr><td class="empty">No month columns matched this filter. Adjust ref mode, month, or N value.</td></tr></tbody>';
    return { yearSummaryColumns: 0, yearSummaryYears: 0 };
  }

  const thead = `
    <thead>
      <tr>
        ${columns.map((col) => `<th class="${col.cls}">${col.label}</th>`).join("")}
        ${dynamicColumns
          .map((col) =>
            col.kind === "year_total" || col.kind === "year_avg"
              ? `<th class="right year-summary-col">${col.label}</th>`
              : `<th class="right">${col.label}</th>`,
          )
          .join("")}
      </tr>
    </thead>
  `;

  const groups = groupedReferenceRows(rows);
  const sheetName = String(sheetData.sheet_name || "");
  const bodyRows = [];
  for (const group of groups) {
    const groupSize = group.rows.length;
    const collapsed = groupSize > 1 && isRowGroupCollapsed(scope, sheetName, group.key);
    const groupKeyEscaped = escapeHtml(group.key);
    const groupLabelEscaped = escapeHtml(group.label);

    for (let idx = 0; idx < group.rows.length; idx += 1) {
      const row = group.rows[idx];
      const isPrimary = idx === 0;
      const rowClasses = [];
      if (groupSize > 1) {
        rowClasses.push("row-group-member");
        if (!isPrimary) {
          rowClasses.push("row-group-child");
          if (collapsed) {
            rowClasses.push("row-group-hidden");
          }
        }
      }

      const productValue = escapeHtml(formatValue(row.product_name));
      let productCellContent = productValue;
      if (groupSize > 1) {
        if (isPrimary) {
          const arrow = collapsed ? "▸" : "▾";
          productCellContent =
            `<div class="row-group-cell">` +
            `<button type="button" class="row-group-toggle" data-group-key="${groupKeyEscaped}" data-group-scope="${escapeHtml(scope)}" data-group-sheet="${escapeHtml(sheetName)}" aria-expanded="${collapsed ? "false" : "true"}">` +
            `<span class="row-group-caret">${arrow}</span>` +
            `<span class="row-group-title">${groupLabelEscaped}</span>` +
            `<span class="row-group-count">(${groupSize})</span>` +
            `</button>` +
            `<span class="row-group-product" title="${productValue}">${productValue}</span>` +
            `</div>`;
        } else {
          productCellContent = `<span class="row-group-product row-group-product-child" title="${productValue}">${productValue}</span>`;
        }
      }

      const staticCells = [
        `<td class="right">${formatValue(row.sr)}</td>`,
        `<td class="left row-group-product-cell">${productCellContent}</td>`,
        `<td class="right">${formatValue(row.ml)}</td>`,
        `<td class="left">${formatValue(row.packing)}</td>`,
      ].join("");

      const dynamicCells = dynamicColumns
        .map((col) => {
          if (col.kind === "year_total" || col.kind === "year_avg") {
            const summary = summarizeRowMetricByMonths(row, col.monthKeys, col.metric);
            const value = col.kind === "year_total" ? summary.total : summary.avg;
            return `<td class="right year-summary-cell">${formatValue(value)}</td>`;
          }
          const monthCell = row.values[col.monthKey] || {};
          return `<td class="right">${formatValue(monthCell[col.metric])}</td>`;
        })
        .join("");

      const rowClassAttr = rowClasses.length ? ` class="${rowClasses.join(" ")}"` : "";
      const rowDataAttr =
        groupSize > 1
          ? ` data-group-key="${groupKeyEscaped}" data-group-primary="${isPrimary ? "1" : "0"}"`
          : "";
      bodyRows.push(`<tr${rowClassAttr}${rowDataAttr}>${staticCells}${dynamicCells}</tr>`);
    }
  }

  target.innerHTML = `${thead}<tbody>${bodyRows.join("")}</tbody>`;
  return {
    yearSummaryColumns: uniqueSummaryYears.length * 2,
    yearSummaryYears: uniqueSummaryYears.length,
  };
}

function renderMeta(target, sheetData, selectedMonths) {
  setText(target, `${sheetData.sheet_name} · rows ${sheetData.rows.length} · showing ${selectedMonths.length} month group(s)`);
}

function mainStyledKey(selectedMonths) {
  const monthsPart = selectedMonths.map((item) => item.key).join(",");
  const mode = el.mainModeSelect.value;
  const nValue = currentN(el.mainNInput);
  const monthValue = monthSelectionSignature(mode, el.mainMonthSelect);
  return `${state.viewerRole}::${state.selectedRegion}::${state.selectedMainWorkbook}::${state.selectedReferenceWorkbook}::${state.pairVersion}::${state.selectedMainSheetName}::${mode}::${nValue}::${monthValue}::${monthsPart}`;
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
  const selectedMonthsByMode = selectedMonthsForMode(mode, el.mainMonthSelect);
  const selectedMonth = selectedMonthsByMode.length
    ? selectedMonthsByMode[selectedMonthsByMode.length - 1]
    : Number.NaN;
  const selectedMonthsCsv = selectedMonthsByMode.join(",");
  const url =
    `/api/main-styled-sheet?${query}` +
    `&sheet=${encodeURIComponent(state.selectedMainSheetName)}` +
    `&month_keys=${encodeURIComponent(monthKeys)}` +
    `&mode=${encodeURIComponent(mode)}` +
    `&n=${encodeURIComponent(String(nValue))}` +
    (Number.isFinite(selectedMonth) ? `&month=${encodeURIComponent(String(selectedMonth))}` : "") +
    (selectedMonthsCsv ? `&months=${encodeURIComponent(selectedMonthsCsv)}` : "");

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
  if (monthAdjusted && modeUsesMonthSelector(el.mainModeSelect.value)) {
    if (modeIsMultiMonthYears(el.mainModeSelect.value)) {
      el.mainMonthSelect.dataset.multi = readSelectedMonthValues(el.mainMonthSelect).join(",");
    } else {
      el.mainMonthSelect.dataset.current = el.mainMonthSelect.value;
    }
    state.mainStyledRequestKey = null;
    await render();
    return;
  }

  el.mainTable.innerHTML = payload.html;
  const groupedMainRows = applyMainRowGrouping(
    el.mainTable,
    String(payload.sheet_name || state.selectedMainSheetName || ""),
  );
  const payloadFrozenCols = Number.parseInt(String(payload.frozen_columns ?? payload.frozen_count ?? 0), 10) || 0;
  const payloadFrozenRows = Number.parseInt(String(payload.frozen_rows ?? 0), 10) || 0;
  applyScopeLayoutOverrides("main", payloadFrozenCols, payloadFrozenRows);

  const hiddenRowsSkipped = Number.parseInt(String(payload.hidden_rows_skipped ?? 0), 10) || 0;
  const hiddenColsSkipped = Number.parseInt(String(payload.hidden_columns_skipped ?? 0), 10) || 0;
  const mainLayout = viewLayoutForScope("main");
  const effectiveMainFreeze = {
    rows: Math.max(0, Number.parseInt(String(mainLayout.lastAppliedFrozenRows || 0), 10) || 0),
    cols: Math.max(0, Number.parseInt(String(mainLayout.lastAppliedFrozenCols || 0), 10) || 0),
  };
  const metaExtras = [];
  if (effectiveMainFreeze.rows !== payloadFrozenRows || effectiveMainFreeze.cols !== payloadFrozenCols) {
    metaExtras.push(
      `normalized freeze from ${payloadFrozenRows} row(s), ${payloadFrozenCols} column(s)`,
    );
  }
  if (effectiveMainFreeze.rows > 0 || effectiveMainFreeze.cols > 0) {
    metaExtras.push(`freeze ${effectiveMainFreeze.rows} row(s), ${effectiveMainFreeze.cols} column(s)`);
  }
  if (groupedMainRows > 1) {
    metaExtras.push(`grouped ${groupedMainRows} row(s)`);
  }
  if (hiddenRowsSkipped > 0 || hiddenColsSkipped > 0) {
    metaExtras.push(`hidden skipped ${hiddenRowsSkipped} row(s), ${hiddenColsSkipped} column(s)`);
  }
  if (mainLayout.hiddenRows.size > 0 || mainLayout.hiddenCols.size > 0) {
    metaExtras.push(`local hidden ${mainLayout.hiddenRows.size} row(s), ${mainLayout.hiddenCols.size} column(s)`);
  }
  const extrasText = metaExtras.length ? ` · ${metaExtras.join(" · ")}` : "";

  const monthCount = (payload.selected_month_labels || []).length;
  if (monthCount > 0) {
    setText(
      el.mainMeta,
      `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · ${monthCount} month group(s)${extrasText}`,
    );
  } else if (payload.filterable) {
    setText(
      el.mainMeta,
      `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · fixed-layout mode${extrasText}`,
    );
  } else {
    setText(
      el.mainMeta,
      `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · full-sheet mode${extrasText}`,
    );
  }
}

function renderReferencePanel(referenceSheet, selectedMonthsRef, referenceTab) {
  if (!referenceSheet) {
    if (referenceTab && !referenceTab.filterable) {
      setEmptyRef("Selected reference sheet is not in a compatible monthly-detail format.");
      setText(el.refMeta, `${referenceTab.sheet_name} · non-filterable`);
      return;
    }
    setEmptyRef("No compatible monthly-detail data found for this reference sheet.");
    if (referenceTab) {
      setText(el.refMeta, `${referenceTab.sheet_name} · no parsed month groups`);
    }
    return;
  }

  const renderInfo = renderTable(el.refTable, referenceSheet, selectedMonthsRef);
  const defaultRefFrozenCols = 4;
  const defaultRefFrozenRows = 0;
  applyScopeLayoutOverrides("reference", defaultRefFrozenCols, defaultRefFrozenRows);
  const refLayout = viewLayoutForScope("reference");
  const effectiveRefFreeze = {
    rows: Math.max(0, Number.parseInt(String(refLayout.lastAppliedFrozenRows || 0), 10) || 0),
    cols: Math.max(0, Number.parseInt(String(refLayout.lastAppliedFrozenCols || 0), 10) || 0),
  };
  renderMeta(el.refMeta, referenceSheet, selectedMonthsRef);
  if (effectiveRefFreeze.rows !== defaultRefFrozenRows || effectiveRefFreeze.cols !== defaultRefFrozenCols) {
    appendText(
      el.refMeta,
      ` · normalized freeze from ${defaultRefFrozenRows} row(s), ${defaultRefFrozenCols} column(s)`,
    );
  }
  if (effectiveRefFreeze.rows > 0 || effectiveRefFreeze.cols > 0) {
    appendText(el.refMeta, ` · freeze ${effectiveRefFreeze.rows} row(s), ${effectiveRefFreeze.cols} column(s)`);
  }
  if (refLayout.hiddenRows.size > 0 || refLayout.hiddenCols.size > 0) {
    appendText(el.refMeta, ` · local hidden ${refLayout.hiddenRows.size} row(s), ${refLayout.hiddenCols.size} column(s)`);
  }

  if (!selectedMonthsRef.length) {
    appendText(el.refMeta, " · no months matched this mode");
  }
  if (renderInfo && renderInfo.yearSummaryColumns > 0) {
    appendText(
      el.refMeta,
      ` · yearly summary columns: ${renderInfo.yearSummaryColumns} (${renderInfo.yearSummaryYears} year(s))`,
    );
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
  clearGridSelectionModel();

  const mainSelectedMonthValues = selectedMonthsForMode(el.mainModeSelect.value, el.mainMonthSelect);
  const mainSelectedMonthLabels = mainSelectedMonthValues.map((monthValue) => monthLabels[monthValue - 1]).filter(Boolean);
  const mainModeText = modeIsSameMonthYears(el.mainModeSelect.value)
    ? `Main: same month over past ${currentN(el.mainNInput)} year(s)`
    : modeIsMultiMonthYears(el.mainModeSelect.value)
      ? `Main: ${mainSelectedMonthLabels.join(", ") || "selected months"} over past ${currentN(el.mainNInput)} year(s)`
      : `Main: past ${currentN(el.mainNInput)} month group(s)`;

  const refSelectedMonthValues = selectedMonthsForMode(el.refModeSelect.value, el.refMonthSelect);
  const refSelectedMonthLabels = refSelectedMonthValues.map((monthValue) => monthLabels[monthValue - 1]).filter(Boolean);
  const refModeText = modeIsSameMonthYears(el.refModeSelect.value)
    ? `Ref: same month over past ${currentN(el.refNInput)} year(s)`
    : modeIsMultiMonthYears(el.refModeSelect.value)
      ? `Ref: ${refSelectedMonthLabels.join(", ") || "selected months"} over past ${currentN(el.refNInput)} year(s)`
      : `Ref: past ${currentN(el.refNInput)} populated month(s)`;

  const scopeText = `Scope: ${state.currentUser || "-"} (${roleLabel(state.viewerRole)}) / ${regionLabel(state.selectedRegion)}`;
  setText(el.statusText, `Live refresh every ${POLL_INTERVAL_MS / 1000}s · ${scopeText} · ${mainModeText} · ${refModeText}`);
  updateLoadHealth();
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
    setStatusError(err);
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
      setStatusError(err);
    });
  }, POLL_INTERVAL_MS);
}

function loginUsersForSwitch() {
  if (Array.isArray(state.loginUsers) && state.loginUsers.length) {
    return state.loginUsers;
  }
  return Array.isArray(state.users) ? state.users : [];
}

function userSummaryByUsername(username) {
  const normalizedUsername = String(username || "");
  const users = loginUsersForSwitch();
  return users.find((user) => user && user.username === normalizedUsername) || null;
}

function userLabelByUsername(username) {
  const user = userSummaryByUsername(username);
  if (user) {
    return user.display_name || user.username || String(username || "-");
  }
  return String(username || "-");
}

function logoutRankForRole(role) {
  if (role === "user") {
    return 0;
  }
  if (role === "asm") {
    return 1;
  }
  if (role === "rsm") {
    return 2;
  }
  return 3;
}

function bestLogoutTargetUsername(currentUsername) {
  const users = loginUsersForSwitch()
    .filter((user) => user && typeof user.username === "string" && user.username)
    .slice()
    .sort((a, b) => {
      const rankDiff = logoutRankForRole(a.role) - logoutRankForRole(b.role);
      if (rankDiff !== 0) {
        return rankDiff;
      }
      const nameA = String(a.display_name || a.username || "").toLowerCase();
      const nameB = String(b.display_name || b.username || "").toLowerCase();
      return nameA.localeCompare(nameB);
    });
  if (!users.length) {
    return null;
  }
  const preferred = users.find((user) => user.username !== currentUsername);
  return (preferred || users[0]).username || null;
}

async function logoutCurrentUser() {
  const previousUser = state.currentUser;
  const nextUser = bestLogoutTargetUsername(previousUser);
  if (!nextUser) {
    setText(el.statusText, "No available user to switch.");
    return;
  }
  if (nextUser === previousUser && loginUsersForSwitch().length <= 1) {
    setText(el.statusText, "No alternate user available to switch.");
    return;
  }

  if (el.userSelect) {
    el.userSelect.value = nextUser;
  }
  if (el.titleUserSelect) {
    el.titleUserSelect.value = nextUser;
  }
  if (el.regionSelect) {
    const hasAllRegion = Array.from(el.regionSelect.options || []).some((option) => option.value === "ALL");
    if (hasAllRegion) {
      el.regionSelect.value = "ALL";
    } else if (el.regionSelect.options.length) {
      el.regionSelect.value = el.regionSelect.options[0].value;
    }
  }

  await onScopeChange();
  setText(
    el.statusText,
    `Logged out ${userLabelByUsername(previousUser)}. Current user: ${userLabelByUsername(state.currentUser)}.`,
  );
}

async function onScopeChange() {
  state.currentUser = el.userSelect ? el.userSelect.value || state.currentUser || "owner" : state.currentUser;
  state.selectedRegion = el.regionSelect ? el.regionSelect.value || state.selectedRegion || "ALL" : state.selectedRegion;
  state.selectedMainWorkbook = null;
  state.selectedReferenceWorkbook = null;
  state.selectedMainSheetName = null;
  state.selectedReferenceSheetName = null;
  state.selectedRefCanonical = null;
  state.cache.clear();
  state.mainStyledRequestKey = null;
  state.mainAvailableMonths.clear();
  state.collapsedRowGroups.clear();
  state.fileRows = [];
  await refreshAccessContext();
  await loadWorkbookOptions();
  await loadSheets();
  await loadFiles();
}

async function onWorkbookChange() {
  state.selectedMainWorkbook = el.mainWorkbookSelect.value;
  state.selectedReferenceWorkbook = el.referenceWorkbookSelect.value;
  state.mainStyledRequestKey = null;
  state.collapsedRowGroups.clear();
  updateWorkbookLabels();
  await loadSheets();
}

function bindEvents() {
  if (el.themeSelect) {
    el.themeSelect.addEventListener("change", () => {
      applyThemePreference(el.themeSelect.value);
      if (typeof window.requestAnimationFrame === "function") {
        window.requestAnimationFrame(() => {
          refreshFrozenSurfaceColors();
        });
      } else {
        refreshFrozenSurfaceColors();
      }
    });
  }

  const handleSystemThemeChange = () => {
    if (state.themePreference !== "system") {
      return;
    }
    applyThemePreference("system", { persist: false });
    if (typeof window.requestAnimationFrame === "function") {
      window.requestAnimationFrame(() => {
        refreshFrozenSurfaceColors();
      });
    } else {
      refreshFrozenSurfaceColors();
    }
  };
  if (themeMediaQuery && typeof themeMediaQuery.addEventListener === "function") {
    themeMediaQuery.addEventListener("change", handleSystemThemeChange);
  } else if (themeMediaQuery && typeof themeMediaQuery.addListener === "function") {
    themeMediaQuery.addListener(handleSystemThemeChange);
  }
  window.addEventListener(
    "load",
    () => {
      scheduleFloatingLayoutMetricsSync();
    },
    { once: true },
  );
  if (document.fonts && document.fonts.ready && typeof document.fonts.ready.then === "function") {
    document.fonts.ready
      .then(() => {
        scheduleFloatingLayoutMetricsSync();
      })
      .catch(() => {
        // Ignore font loading errors and continue with existing measurements.
      });
  }

  if (typeof ribbonApi.bindRibbonTabEvents === "function") {
    ribbonApi.bindRibbonTabEvents({
      el,
      state,
      setRibbonTab,
      setRibbonCollapsed,
    });
  }
  if (typeof ribbonApi.bindMirrorControls === "function") {
    ribbonApi.bindMirrorControls({
      el,
      triggerChange,
      triggerInput,
      syncUploadRegionInputs,
    });
  }

  if (el.ribbonRefreshAllBtn) {
    el.ribbonRefreshAllBtn.addEventListener("click", () => {
      refreshAllCommandData().catch((err) => {
        setStatusError(err);
      });
    });
  }
  if (el.ribbonRefreshFilesBtn) {
    el.ribbonRefreshFilesBtn.addEventListener("click", () => {
      (async () => {
        beginBusy("Refreshing files...");
        try {
          await loadFiles();
        } finally {
          endBusy();
        }
      })().catch((err) => {
        setStatusError(err);
      });
    });
  }
  if (el.ribbonRoleGuideBtn) {
    el.ribbonRoleGuideBtn.addEventListener("click", (event) => {
      setModalOpen(true, event.currentTarget);
    });
  }
  if (el.ribbonOpenRoleGuideBtn) {
    el.ribbonOpenRoleGuideBtn.addEventListener("click", (event) => {
      setModalOpen(true, event.currentTarget);
    });
  }
  if (el.ribbonOnboardingBtn) {
    el.ribbonOnboardingBtn.addEventListener("click", () => {
      setRibbonTab("onboarding");
      const currentlyHidden = el.onboardingSteps ? el.onboardingSteps.classList.contains("hidden") : true;
      setOnboardingExpanded(currentlyHidden);
      const onboardingPanel =
        (el.onboardingSteps && el.onboardingSteps.closest('[data-ribbon-panel="onboarding"]')) ||
        (el.roleOnboardingText && el.roleOnboardingText.closest('[data-ribbon-panel="onboarding"]'));
      scrollToNode(onboardingPanel);
    });
  }

  const accessPanel =
    (el.accessHint && el.accessHint.closest('[data-ribbon-panel="roles"]')) ||
    (el.accessHint && el.accessHint.closest(".panel"));
  const mainPanel = el.mainTable ? el.mainTable.closest(".panel") : null;
  const refPanel = el.refTable ? el.refTable.closest(".panel") : null;
  const filesPanel =
    (el.filesTableBody && el.filesTableBody.closest('[data-ribbon-panel="files"]')) ||
    (el.filesTableBody && el.filesTableBody.closest(".panel"));

  if (el.ribbonGoMainBtn) {
    el.ribbonGoMainBtn.addEventListener("click", () => {
      setActiveViewScope("main");
      scrollToNode(mainPanel);
    });
  }
  if (el.ribbonGoRefBtn) {
    el.ribbonGoRefBtn.addEventListener("click", () => {
      if (document.body.classList.contains("detail-drawer-mode")) {
        setReferenceDrawerExpanded(true);
      } else if (state.desktopDetailCollapsed) {
        setDesktopDetailCollapsed(false);
      }
      setActiveViewScope("reference");
      scrollToNode(refPanel);
    });
  }
  if (el.ribbonGoAccessBtn) {
    el.ribbonGoAccessBtn.addEventListener("click", () => {
      setRibbonTab("roles");
      scrollToNode(accessPanel);
    });
  }
  if (el.ribbonGoFilesBtn) {
    el.ribbonGoFilesBtn.addEventListener("click", () => {
      setRibbonTab("files");
      scrollToNode(filesPanel);
    });
  }

  if (el.ribbonFreezePanesBtn) {
    el.ribbonFreezePanesBtn.addEventListener("click", () => {
      runRibbonGridAction(
        () => {
          freezePanesFromContext();
        },
        { requirePoint: true, emptyPointMessage: "Select a cell first." },
      );
    });
  }
  if (el.ribbonFreezeTopRowBtn) {
    el.ribbonFreezeTopRowBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        freezeTopRowFromContext();
      });
    });
  }
  if (el.ribbonFreezeFirstColBtn) {
    el.ribbonFreezeFirstColBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        freezeFirstColumnFromContext();
      });
    });
  }
  if (el.ribbonUnfreezePanesBtn) {
    el.ribbonUnfreezePanesBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        unfreezePanesFromContext();
      });
    });
  }

  if (el.ribbonSelectRowBtn) {
    el.ribbonSelectRowBtn.addEventListener("click", () => {
      runRibbonGridAction(
        () => {
          selectContextRow();
        },
        { requirePoint: true, emptyPointMessage: "Select a cell in the target row first." },
      );
    });
  }
  if (el.ribbonSelectColumnBtn) {
    el.ribbonSelectColumnBtn.addEventListener("click", () => {
      runRibbonGridAction(
        () => {
          selectContextColumn();
        },
        { requirePoint: true, emptyPointMessage: "Select a cell in the target column first." },
      );
    });
  }
  if (el.ribbonSelectAllBtn) {
    el.ribbonSelectAllBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        selectAllForContextScope();
      });
    });
  }
  if (el.ribbonInsertRowAboveBtn) {
    el.ribbonInsertRowAboveBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        insertRowFromContext("above");
      });
    });
  }
  if (el.ribbonInsertRowBelowBtn) {
    el.ribbonInsertRowBelowBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        insertRowFromContext("below");
      });
    });
  }
  if (el.ribbonDeleteRowBtn) {
    el.ribbonDeleteRowBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        deleteRowsFromContext();
      });
    });
  }
  if (el.ribbonHideRowsBtn) {
    el.ribbonHideRowsBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        updateHiddenRowsFromContext(true);
      });
    });
  }
  if (el.ribbonUnhideRowsBtn) {
    el.ribbonUnhideRowsBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        updateHiddenRowsFromContext(false);
      });
    });
  }
  if (el.ribbonInsertColLeftBtn) {
    el.ribbonInsertColLeftBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        insertColumnFromContext("left");
      });
    });
  }
  if (el.ribbonInsertColRightBtn) {
    el.ribbonInsertColRightBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        insertColumnFromContext("right");
      });
    });
  }
  if (el.ribbonDeleteColBtn) {
    el.ribbonDeleteColBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        deleteColumnsFromContext();
      });
    });
  }
  if (el.ribbonHideColsBtn) {
    el.ribbonHideColsBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        updateHiddenColumnsFromContext(true);
      });
    });
  }
  if (el.ribbonUnhideColsBtn) {
    el.ribbonUnhideColsBtn.addEventListener("click", () => {
      runRibbonGridAction(() => {
        updateHiddenColumnsFromContext(false);
      });
    });
  }
  if (el.ribbonCollapseAllGroupsBtn) {
    el.ribbonCollapseAllGroupsBtn.addEventListener("click", () => {
      const scope = activeGridScopeForCommands();
      const scopeLabel = scope === "reference" ? "Detail view" : "Main view";
      const changed = setAllRowGroupsCollapsed(scope, true);
      setText(
        el.statusText,
        changed > 0
          ? `Collapsed ${changed} spelling-based row group(s) · ${scopeLabel}`
          : `No row groups available in ${scopeLabel.toLowerCase()}.`,
      );
      syncRibbonGridControlState();
    });
  }
  if (el.ribbonExpandAllGroupsBtn) {
    el.ribbonExpandAllGroupsBtn.addEventListener("click", () => {
      const scope = activeGridScopeForCommands();
      const scopeLabel = scope === "reference" ? "Detail view" : "Main view";
      const changed = setAllRowGroupsCollapsed(scope, false);
      setText(
        el.statusText,
        changed > 0
          ? `Expanded ${changed} spelling-based row group(s) · ${scopeLabel}`
          : `No row groups available in ${scopeLabel.toLowerCase()}.`,
      );
      syncRibbonGridControlState();
    });
  }

  if (el.ribbonJumpAssignRsmBtn) {
    el.ribbonJumpAssignRsmBtn.addEventListener("click", () => {
      setRibbonTab("roles");
      scrollToNode(el.assignRsmForm && !el.assignRsmForm.classList.contains("hidden") ? el.assignRsmForm : accessPanel);
    });
  }
  if (el.ribbonJumpMapUserBtn) {
    el.ribbonJumpMapUserBtn.addEventListener("click", () => {
      setRibbonTab("roles");
      scrollToNode(el.mapUserRsmForm && !el.mapUserRsmForm.classList.contains("hidden") ? el.mapUserRsmForm : accessPanel);
    });
  }
  if (el.ribbonJumpAssignAsmBtn) {
    el.ribbonJumpAssignAsmBtn.addEventListener("click", () => {
      setRibbonTab("roles");
      scrollToNode(el.assignAsmForm && !el.assignAsmForm.classList.contains("hidden") ? el.assignAsmForm : accessPanel);
    });
  }
  if (el.ribbonJumpTownshipBtn) {
    el.ribbonJumpTownshipBtn.addEventListener("click", () => {
      setRibbonTab("roles");
      scrollToNode(el.asmTownshipForm && !el.asmTownshipForm.classList.contains("hidden") ? el.asmTownshipForm : accessPanel);
    });
  }

  if (el.referenceDrawerToggleBtn) {
    el.referenceDrawerToggleBtn.addEventListener("click", () => {
      setReferenceDrawerExpanded(!state.referenceDrawerExpanded);
    });
  }
  if (el.referencePanelToggleBtn) {
    el.referencePanelToggleBtn.addEventListener("click", () => {
      if (document.body.classList.contains("detail-drawer-mode")) {
        setReferenceDrawerExpanded(!state.referenceDrawerExpanded);
        return;
      }
      setDesktopDetailCollapsed(!state.desktopDetailCollapsed);
    });
  }
  if (el.referencePanelTitle) {
    el.referencePanelTitle.addEventListener("dblclick", () => {
      if (document.body.classList.contains("detail-drawer-mode")) {
        setReferenceDrawerExpanded(!state.referenceDrawerExpanded);
        return;
      }
      setDesktopDetailCollapsed(!state.desktopDetailCollapsed);
    });
  }
  if (el.viewsSplitHandle) {
    el.viewsSplitHandle.addEventListener("pointerdown", (event) => {
      event.preventDefault();
      beginViewsSplitDrag(event);
    });
    el.viewsSplitHandle.addEventListener("keydown", (event) => {
      if (document.body.classList.contains("detail-drawer-mode")) {
        return;
      }
      if (event.key === "ArrowLeft") {
        event.preventDefault();
        setDesktopDetailCollapsed(false);
        setViewsSplitRatio(state.viewsSplitRatio - VIEWS_SPLIT_KEY_STEP);
        return;
      }
      if (event.key === "ArrowRight") {
        event.preventDefault();
        setDesktopDetailCollapsed(false);
        setViewsSplitRatio(state.viewsSplitRatio + VIEWS_SPLIT_KEY_STEP);
        return;
      }
      if (event.key === "Home") {
        event.preventDefault();
        setDesktopDetailCollapsed(false);
        setViewsSplitRatio(MIN_VIEWS_MAIN_RATIO);
        return;
      }
      if (event.key === "End") {
        event.preventDefault();
        setDesktopDetailCollapsed(false);
        setViewsSplitRatio(MAX_VIEWS_MAIN_RATIO);
      }
    });
  }
  window.addEventListener(
    "resize",
    () => {
      syncReferenceDrawerMode();
    },
    { passive: true },
  );
  document.addEventListener("mouseup", () => {
    endGridSelection();
  });
  document.addEventListener("pointerdown", (event) => {
    const target = event.target;
    if (state.contextMenuOpen && el.gridContextMenu) {
      if (!(target instanceof Node && el.gridContextMenu.contains(target))) {
        hideGridContextMenu();
      }
    }
    if (state.sheetTabContextMenuOpen && el.sheetTabContextMenu) {
      if (!(target instanceof Node && el.sheetTabContextMenu.contains(target))) {
        hideSheetTabContextMenu();
      }
    }
  });
  document.addEventListener(
    "scroll",
    () => {
      if (state.contextMenuOpen) {
        hideGridContextMenu();
      }
      if (state.sheetTabContextMenuOpen) {
        hideSheetTabContextMenu();
      }
    },
    { passive: true, capture: true },
  );
  window.addEventListener(
    "resize",
    () => {
      if (state.contextMenuOpen) {
        hideGridContextMenu();
      }
      if (state.sheetTabContextMenuOpen) {
        hideSheetTabContextMenu();
      }
    },
    { passive: true },
  );

  const bindSelectionArea = (wrap) => {
    if (!wrap) {
      return;
    }
    const inferredScope = wrap === el.referenceTableWrap ? "reference" : "main";
    wrap.addEventListener("pointerenter", () => {
      setActiveViewScope(inferredScope);
    });
    wrap.addEventListener("mousedown", (event) => {
      startGridSelection(event);
    });
    wrap.addEventListener("mouseover", (event) => {
      moveGridSelection(event);
    });
    wrap.addEventListener("contextmenu", (event) => {
      const cell = event.target instanceof HTMLElement ? event.target.closest("th, td") : null;
      if (!cell) {
        return;
      }
      event.preventDefault();
      openContextMenuForCell(cell, event.clientX, event.clientY);
    });
    wrap.addEventListener(
      "touchstart",
      (event) => {
        if (!event.touches || event.touches.length !== 1) {
          clearContextMenuLongPressTimer();
          return;
        }
        const touch = event.touches[0];
        const touchX = Number(touch.clientX || 0);
        const touchY = Number(touch.clientY || 0);
        const cell = event.target instanceof HTMLElement ? event.target.closest("th, td") : null;
        if (!cell) {
          clearContextMenuLongPressTimer();
          return;
        }
        clearContextMenuLongPressTimer();
        state.contextMenuLongPressTimer = window.setTimeout(() => {
          const anchor = document.elementFromPoint(touchX, touchY);
          const targetCell = anchor instanceof HTMLElement ? anchor.closest("th, td") : cell;
          openContextMenuForCell(targetCell, touchX, touchY);
          state.contextMenuLongPressTimer = null;
        }, LONG_PRESS_OPEN_MS);
      },
      { passive: true },
    );
    wrap.addEventListener(
      "touchmove",
      () => {
        clearContextMenuLongPressTimer();
      },
      { passive: true },
    );
    wrap.addEventListener(
      "touchend",
      () => {
        clearContextMenuLongPressTimer();
      },
      { passive: true },
    );
    wrap.addEventListener(
      "touchcancel",
      () => {
        clearContextMenuLongPressTimer();
      },
      { passive: true },
    );
    wrap.addEventListener(
      "wheel",
      (event) => {
        if (!event.ctrlKey) {
          return;
        }
        event.preventDefault();
        const zoomScope = wrap === el.referenceTableWrap ? "reference" : "main";
        setActiveViewScope(zoomScope, { refresh: false });
        adjustSheetZoom(event.deltaY < 0 ? ZOOM_STEP : -ZOOM_STEP, zoomScope);
      },
      { passive: false },
    );
  };

  bindSelectionArea(el.mainTableWrap);
  bindSelectionArea(el.referenceTableWrap);

  const bindRowGroupToggleEvents = (wrap, fallbackScope) => {
    if (!wrap) {
      return;
    }
    wrap.addEventListener("mousedown", (event) => {
      const toggleBtn = event.target instanceof HTMLElement ? event.target.closest(".row-group-toggle") : null;
      if (!toggleBtn) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
    });
    wrap.addEventListener("click", (event) => {
      const toggleBtn = event.target instanceof HTMLElement ? event.target.closest(".row-group-toggle") : null;
      if (!toggleBtn || !(toggleBtn instanceof HTMLElement)) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      const groupKey = String(toggleBtn.dataset.groupKey || "").trim();
      const scope = normalizeViewScope(String(toggleBtn.dataset.groupScope || fallbackScope).trim() || fallbackScope);
      const defaultSheet = scope === "reference" ? state.selectedReferenceSheetName : state.selectedMainSheetName;
      const sheetName = String(toggleBtn.dataset.groupSheet || defaultSheet || "").trim();
      if (!groupKey) {
        return;
      }
      const expanded = toggleBtn.getAttribute("aria-expanded") === "true";
      const collapsed = expanded;
      setRowGroupCollapsed(scope, sheetName, groupKey, collapsed);
      applyRowGroupStateForScope(scope, groupKey, collapsed);
      syncRibbonGridControlState();
    });
  };

  bindRowGroupToggleEvents(el.mainTableWrap, "main");
  bindRowGroupToggleEvents(el.referenceTableWrap, "reference");

  if (el.ctxCopyBtn) {
    el.ctxCopyBtn.addEventListener("click", () => {
      copyCurrentSelection()
        .then(() => {
          hideGridContextMenu();
        })
        .catch((err) => {
          setStatusError(err);
          hideGridContextMenu();
        });
    });
  }
  if (el.ctxPasteBtn) {
    el.ctxPasteBtn.addEventListener("click", () => {
      pasteClipboardIntoGrid()
        .then(() => {
          hideGridContextMenu();
        })
        .catch((err) => {
          setStatusError(err);
          hideGridContextMenu();
        });
    });
  }
  if (el.ctxClearSelectionBtn) {
    el.ctxClearSelectionBtn.addEventListener("click", () => {
      clearGridSelectionModel();
    });
  }
  if (el.ctxFreezePanesBtn) {
    el.ctxFreezePanesBtn.addEventListener("click", () => {
      freezePanesFromContext();
      hideGridContextMenu();
    });
  }
  if (el.ctxUnfreezePanesBtn) {
    el.ctxUnfreezePanesBtn.addEventListener("click", () => {
      unfreezePanesFromContext();
      hideGridContextMenu();
    });
  }
  if (el.ctxHideRowsBtn) {
    el.ctxHideRowsBtn.addEventListener("click", () => {
      updateHiddenRowsFromContext(true);
      hideGridContextMenu();
    });
  }
  if (el.ctxUnhideRowsBtn) {
    el.ctxUnhideRowsBtn.addEventListener("click", () => {
      updateHiddenRowsFromContext(false);
      hideGridContextMenu();
    });
  }
  if (el.ctxHideColsBtn) {
    el.ctxHideColsBtn.addEventListener("click", () => {
      updateHiddenColumnsFromContext(true);
      hideGridContextMenu();
    });
  }
  if (el.ctxUnhideColsBtn) {
    el.ctxUnhideColsBtn.addEventListener("click", () => {
      updateHiddenColumnsFromContext(false);
      hideGridContextMenu();
    });
  }
  if (el.ctxHideBothBtn) {
    el.ctxHideBothBtn.addEventListener("click", () => {
      updateHiddenRowsAndColumnsFromContext(true);
      hideGridContextMenu();
    });
  }
  if (el.ctxUnhideBothBtn) {
    el.ctxUnhideBothBtn.addEventListener("click", () => {
      updateHiddenRowsAndColumnsFromContext(false);
      hideGridContextMenu();
    });
  }
  if (el.ctxInsertRowAboveBtn) {
    el.ctxInsertRowAboveBtn.addEventListener("click", () => {
      insertRowFromContext("above");
      hideGridContextMenu();
    });
  }
  if (el.ctxInsertRowBelowBtn) {
    el.ctxInsertRowBelowBtn.addEventListener("click", () => {
      insertRowFromContext("below");
      hideGridContextMenu();
    });
  }
  if (el.ctxDeleteRowBtn) {
    el.ctxDeleteRowBtn.addEventListener("click", () => {
      deleteRowsFromContext();
      hideGridContextMenu();
    });
  }
  if (el.ctxInsertColLeftBtn) {
    el.ctxInsertColLeftBtn.addEventListener("click", () => {
      insertColumnFromContext("left");
      hideGridContextMenu();
    });
  }
  if (el.ctxInsertColRightBtn) {
    el.ctxInsertColRightBtn.addEventListener("click", () => {
      insertColumnFromContext("right");
      hideGridContextMenu();
    });
  }
  if (el.ctxDeleteColBtn) {
    el.ctxDeleteColBtn.addEventListener("click", () => {
      deleteColumnsFromContext();
      hideGridContextMenu();
    });
  }
  if (el.ctxSelectRowBtn) {
    el.ctxSelectRowBtn.addEventListener("click", () => {
      selectContextRow();
      hideGridContextMenu();
    });
  }
  if (el.ctxSelectColumnBtn) {
    el.ctxSelectColumnBtn.addEventListener("click", () => {
      selectContextColumn();
      hideGridContextMenu();
    });
  }
  if (el.ctxSelectAllBtn) {
    el.ctxSelectAllBtn.addEventListener("click", () => {
      selectAllForContextScope();
      hideGridContextMenu();
    });
  }
  if (el.ctxOpenDetailsBtn) {
    el.ctxOpenDetailsBtn.addEventListener("click", () => {
      openDetailsPanelFromContext();
      hideGridContextMenu();
    });
  }
  if (el.sheetTabCtxRenameBtn) {
    el.sheetTabCtxRenameBtn.addEventListener("click", () => {
      renameSheetFromTabContextMenu();
    });
  }
  if (el.ctxZoomInBtn) {
    el.ctxZoomInBtn.addEventListener("click", () => {
      const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope || "main");
      setActiveViewScope(scope, { refresh: false });
      adjustSheetZoom(ZOOM_STEP, scope);
      hideGridContextMenu();
    });
  }
  if (el.ctxZoomOutBtn) {
    el.ctxZoomOutBtn.addEventListener("click", () => {
      const scope = normalizeViewScope(state.contextMenuScope || state.selectionScope || state.activeViewScope || "main");
      setActiveViewScope(scope, { refresh: false });
      adjustSheetZoom(-ZOOM_STEP, scope);
      hideGridContextMenu();
    });
  }

  if (el.zoomRangeInput) {
    el.zoomRangeInput.addEventListener("input", () => {
      setSheetZoom(el.zoomRangeInput.value);
    });
  }
  if (el.zoomOutBtn) {
    el.zoomOutBtn.addEventListener("click", () => {
      adjustSheetZoom(-ZOOM_STEP);
    });
  }
  if (el.zoomInBtn) {
    el.zoomInBtn.addEventListener("click", () => {
      adjustSheetZoom(ZOOM_STEP);
    });
  }
  if (el.zoomResetBtn) {
    el.zoomResetBtn.addEventListener("click", () => {
      setSheetZoom(100);
    });
  }

  if (el.formulaApplyBtn) {
    el.formulaApplyBtn.addEventListener("click", () => {
      try {
        applyFormulaInput();
      } catch (err) {
        setStatusError(err);
      }
    });
  }
  if (el.formulaCancelBtn) {
    el.formulaCancelBtn.addEventListener("click", () => {
      cancelFormulaEdit();
    });
  }
  if (el.formulaFxBtn && el.formulaInput) {
    el.formulaFxBtn.addEventListener("click", () => {
      if (!el.formulaInput.value.trim()) {
        el.formulaInput.value = "=";
      } else if (!el.formulaInput.value.trim().startsWith("=")) {
        el.formulaInput.value = `=${el.formulaInput.value.trim()}`;
      }
      el.formulaInput.focus();
      el.formulaInput.setSelectionRange(el.formulaInput.value.length, el.formulaInput.value.length);
    });
  }
  if (el.formulaInput) {
    el.formulaInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        try {
          applyFormulaInput();
        } catch (err) {
          setStatusError(err);
        }
        return;
      }
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        cancelFormulaEdit();
      }
    });
  }
  if (el.formulaNameBox) {
    el.formulaNameBox.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") {
        return;
      }
      event.preventDefault();
      jumpToAddressFromNameBox();
    });
  }

  if (el.ribbonUploadBtn) {
    el.ribbonUploadBtn.addEventListener("click", () => {
      uploadSelectedWorkbooks().catch((err) => {
        setStatusError(err);
      });
    });
  }

  if (el.onboardingToggleBtn) {
    el.onboardingToggleBtn.addEventListener("click", () => {
      const currentlyHidden = el.onboardingSteps ? el.onboardingSteps.classList.contains("hidden") : true;
      setOnboardingExpanded(currentlyHidden);
    });
  }
  if (el.roleInfoBtn) {
    el.roleInfoBtn.addEventListener("click", (event) => {
      setModalOpen(true, event.currentTarget);
    });
  }
  if (el.roleInfoInlineBtn) {
    el.roleInfoInlineBtn.addEventListener("click", (event) => {
      setModalOpen(true, event.currentTarget);
    });
  }
  if (el.fullscreenToggleBtn) {
    el.fullscreenToggleBtn.addEventListener("click", () => {
      toggleFullscreenMode()
        .catch((err) => {
          setStatusError(err);
        })
        .finally(() => {
          syncFullscreenToggleButton();
        });
    });
  }
  document.addEventListener("fullscreenchange", () => {
    syncFullscreenToggleButton();
    scheduleFloatingLayoutMetricsSync();
  });
  document.addEventListener("webkitfullscreenchange", () => {
    syncFullscreenToggleButton();
    scheduleFloatingLayoutMetricsSync();
  });
  if (el.roleInfoCloseBtn) {
    el.roleInfoCloseBtn.addEventListener("click", () => {
      setModalOpen(false);
    });
  }
  if (el.roleInfoDoneBtn) {
    el.roleInfoDoneBtn.addEventListener("click", () => {
      setModalOpen(false);
    });
  }
  if (el.roleInfoModal) {
    el.roleInfoModal.addEventListener("click", (event) => {
      if (event.target === el.roleInfoModal) {
        setModalOpen(false);
      }
    });
    el.roleInfoModal.addEventListener("keydown", (event) => {
      if (event.key !== "Tab" || el.roleInfoModal.hidden) {
        return;
      }
      const focusables = getRoleInfoModalFocusables();
      if (!focusables.length) {
        return;
      }
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const active = document.activeElement;

      if (event.shiftKey) {
        if (active === first || !focusables.includes(active)) {
          last.focus();
          event.preventDefault();
        }
        return;
      }

      if (active === last) {
        first.focus();
        event.preventDefault();
      }
    });
  }
  document.addEventListener("keydown", (event) => {
    const shortcutKey = String(event.key || "").toLowerCase();
    const hasShortcutModifier = event.ctrlKey || event.metaKey;
    if (hasShortcutModifier && !event.altKey && !isEditableEventTarget(event.target)) {
      if (shortcutKey === "c" && hasGridSelection()) {
        event.preventDefault();
        copyCurrentSelection().catch((err) => {
          setStatusError(err);
        });
        return;
      }
      if (shortcutKey === "v") {
        event.preventDefault();
        pasteClipboardIntoGrid().catch((err) => {
          setStatusError(err);
        });
        return;
      }
    }

    if (event.ctrlKey && event.key === "F1") {
      event.preventDefault();
      setRibbonCollapsed(!state.ribbonCollapsed);
      return;
    }

    if (event.key === "Escape" && el.roleInfoModal && !el.roleInfoModal.hidden) {
      setModalOpen(false);
      return;
    }

    if (event.key === "Escape" && state.contextMenuOpen) {
      hideGridContextMenu();
      return;
    }

    if (event.key === "Escape" && state.sheetTabContextMenuOpen) {
      hideSheetTabContextMenu();
      return;
    }

    if (event.key === "Escape" && state.selectionRanges && state.selectionRanges.length) {
      clearGridSelectionModel();
    }
  });

  if (el.userSelect) {
    el.userSelect.addEventListener("change", () => {
      onScopeChange().catch((err) => {
        setStatusError(err);
      });
    });
  }
  if (el.titleUserSelect) {
    el.titleUserSelect.addEventListener("change", () => {
      if (el.userSelect && el.userSelect.value !== el.titleUserSelect.value) {
        el.userSelect.value = el.titleUserSelect.value;
      }
      onScopeChange().catch((err) => {
        setStatusError(err);
      });
    });
  }
  if (el.logoutCurrentUserBtn) {
    el.logoutCurrentUserBtn.addEventListener("click", () => {
      logoutCurrentUser().catch((err) => {
        setStatusError(err);
      });
    });
  }

  if (el.regionSelect) {
    el.regionSelect.addEventListener("change", () => {
      onScopeChange().catch((err) => {
        setStatusError(err);
      });
    });
  }

  el.uploadBtn.addEventListener("click", () => {
    uploadSelectedWorkbooks().catch((err) => {
      setStatusError(err);
    });
  });

  if (el.assignAsmRegionSelect) {
    el.assignAsmRegionSelect.addEventListener("change", () => {
      const region = el.assignAsmRegionSelect.value;
      setTownshipSelect(el.assignAsmTownshipsSelect, region, []);
    });
  }
  if (el.asmTownshipRegionSelect) {
    el.asmTownshipRegionSelect.addEventListener("change", () => {
      const asmUser = el.asmTownshipUserSelect ? el.asmTownshipUserSelect.value : "";
      const region = el.asmTownshipRegionSelect.value;
      const selected =
        state.assignments &&
        state.assignments.asm_townships &&
        state.assignments.asm_townships[asmUser] &&
        Array.isArray(state.assignments.asm_townships[asmUser][region])
          ? state.assignments.asm_townships[asmUser][region]
          : [];
      setTownshipSelect(el.asmTownshipSelect, region, selected);
    });
  }
  if (el.asmTownshipUserSelect) {
    el.asmTownshipUserSelect.addEventListener("change", () => {
      const asmUser = el.asmTownshipUserSelect.value;
      const region = el.asmTownshipRegionSelect ? el.asmTownshipRegionSelect.value : "";
      const selected =
        state.assignments &&
        state.assignments.asm_townships &&
        state.assignments.asm_townships[asmUser] &&
        Array.isArray(state.assignments.asm_townships[asmUser][region])
          ? state.assignments.asm_townships[asmUser][region]
          : [];
      setTownshipSelect(el.asmTownshipSelect, region, selected);
    });
  }
  if (el.switchUserRoleSelect) {
    el.switchUserRoleSelect.addEventListener("change", () => {
      syncSwitchUserRoleControls();
    });
  }

  if (el.assignRsmForm) {
    el.assignRsmForm.addEventListener("submit", (event) => {
      event.preventDefault();
      (async () => {
        beginBusy("Saving RSM mapping...");
        try {
          const regions = (el.assignRsmRegionsInput.value || "")
            .split(",")
            .map((item) => item.trim())
            .filter(Boolean);
          await postScopedJson("/api/access/assign-rsm", {
            username: el.assignRsmUsernameInput.value,
            display_name: el.assignRsmDisplayInput.value,
            regions,
          });
          await refreshAccessAndFiles();
          await loadWorkbookOptions();
          await loadSheets();
          setText(el.statusText, "RSM assignment updated.");
        } finally {
          endBusy();
        }
      })().catch((err) => {
        setStatusError(err);
      });
    });
  }

  if (el.mapUserRsmForm) {
    el.mapUserRsmForm.addEventListener("submit", (event) => {
      event.preventDefault();
      (async () => {
        beginBusy("Assigning user to RSM...");
        try {
          await postScopedJson("/api/access/assign-user-to-rsm", {
            username: el.mapUserUsernameInput.value,
            rsm_username: el.mapUserRsmSelect.value,
          });
          await refreshAccessAndFiles();
          await loadWorkbookOptions();
          await loadSheets();
          setText(el.statusText, "User mapped to RSM.");
        } finally {
          endBusy();
        }
      })().catch((err) => {
        setStatusError(err);
      });
    });
  }

  if (el.switchUserRoleForm) {
    el.switchUserRoleForm.addEventListener("submit", (event) => {
      event.preventDefault();
      (async () => {
        beginBusy("Saving user role...");
        try {
          const role = normalizeRoleToken(el.switchUserRoleSelect ? el.switchUserRoleSelect.value : "user");
          const payload = {
            username: el.switchUserRoleUsernameInput.value,
            display_name: el.switchUserRoleDisplayInput.value,
            role,
          };
          if (role === "rsm") {
            payload.regions = (el.switchUserRoleRegionsInput.value || "")
              .split(",")
              .map((item) => item.trim())
              .filter(Boolean);
          } else if (el.switchUserRoleRsmSelect && el.switchUserRoleRsmSelect.value) {
            payload.rsm_username = el.switchUserRoleRsmSelect.value;
          }

          await postScopedJson("/api/access/set-user-role", payload);
          await refreshAccessAndFiles();
          await loadWorkbookOptions();
          await loadSheets();
          setText(el.statusText, "User role updated.");
        } finally {
          endBusy();
        }
      })().catch((err) => {
        setStatusError(err);
      });
    });
  }

  if (el.assignAsmForm) {
    el.assignAsmForm.addEventListener("submit", (event) => {
      event.preventDefault();
      (async () => {
        beginBusy("Saving ASM mapping...");
        try {
          const payload = {
            asm_username: el.assignAsmUsernameInput.value,
            display_name: el.assignAsmDisplayInput.value,
            region: el.assignAsmRegionSelect.value,
            townships: selectedMultiValues(el.assignAsmTownshipsSelect),
          };
          if (state.viewerRole === "owner" && el.assignAsmRsmSelect) {
            payload.rsm_username = el.assignAsmRsmSelect.value;
          }
          await postScopedJson("/api/access/assign-asm", payload);
          await refreshAccessAndFiles();
          await loadWorkbookOptions();
          await loadSheets();
          setText(el.statusText, "ASM assignment updated.");
        } finally {
          endBusy();
        }
      })().catch((err) => {
        setStatusError(err);
      });
    });
  }

  if (el.asmTownshipForm) {
    el.asmTownshipForm.addEventListener("submit", (event) => {
      event.preventDefault();
      (async () => {
        beginBusy("Updating ASM township permissions...");
        try {
          await postScopedJson("/api/access/set-asm-townships", {
            asm_username: el.asmTownshipUserSelect.value,
            region: el.asmTownshipRegionSelect.value,
            townships: selectedMultiValues(el.asmTownshipSelect),
          });
          await refreshAccessAndFiles();
          await loadWorkbookOptions();
          await loadSheets();
          setText(el.statusText, "ASM township permissions updated.");
        } finally {
          endBusy();
        }
      })().catch((err) => {
        setStatusError(err);
      });
    });
  }

  if (el.refreshFilesBtn) {
    el.refreshFilesBtn.addEventListener("click", () => {
      (async () => {
        beginBusy("Refreshing files...");
        try {
          await loadFiles();
        } finally {
          endBusy();
        }
      })().catch((err) => {
        setStatusError(err);
      });
    });
  }

  if (el.filesTableBody) {
    el.filesTableBody.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }

      const saveBtn = target.closest(".file-save-btn");
      if (saveBtn instanceof HTMLElement) {
        const filename = saveBtn.dataset.file || "";
        const row = saveBtn.closest("tr");
        const regionSelect = row ? row.querySelector(".file-region-select") : null;
        const viewModeSelect = row ? row.querySelector(".file-view-mode-select") : null;
        const region = regionSelect && "value" in regionSelect ? regionSelect.value : "";
        const viewMode = viewModeSelect && "value" in viewModeSelect ? viewModeSelect.value : "";
        if (!filename || !region || !viewMode) {
          return;
        }
        (async () => {
          beginBusy("Updating file settings...");
          try {
            await patchScopedJson(`/api/files/${encodeURIComponent(filename)}`, {
              region,
              view_mode: viewMode,
            });
            await loadFiles();
            await loadWorkbookOptions();
            await loadSheets();
            setText(
              el.statusText,
              `Updated ${filename} · ${regionLabel(region)} · ${fileViewModeLabel(viewMode)}.`,
            );
          } finally {
            endBusy();
          }
        })().catch((err) => {
          setStatusError(err);
        });
        return;
      }

      const deleteBtn = target.closest(".file-delete-btn");
      if (deleteBtn instanceof HTMLElement) {
        const filename = deleteBtn.dataset.file || "";
        if (!filename) {
          return;
        }
        if (!window.confirm(`Delete file "${filename}"?`)) {
          return;
        }
        (async () => {
          beginBusy("Deleting file...");
          try {
            await deleteScoped(`/api/files/${encodeURIComponent(filename)}`);
            await loadFiles();
            await loadWorkbookOptions();
            await loadSheets();
            setText(el.statusText, `Deleted ${filename}.`);
          } finally {
            endBusy();
          }
        })().catch((err) => {
          setStatusError(err);
        });
      }
    });
  }

  el.mainWorkbookSelect.addEventListener("change", () => {
    onWorkbookChange().catch((err) => {
      setStatusError(err);
    });
  });

  el.referenceWorkbookSelect.addEventListener("change", () => {
    onWorkbookChange().catch((err) => {
      setStatusError(err);
    });
  });

  el.referenceTabSelect.addEventListener("change", () => {
    onReferenceTabChange(el.referenceTabSelect.value).catch((err) => {
      setStatusError(err);
    });
  });

  el.mainTabSelect.addEventListener("change", () => {
    onMainTabChange(el.mainTabSelect.value).catch((err) => {
      setStatusError(err);
    });
  });

  el.mainModeSelect.addEventListener("change", () => {
    state.mainStyledRequestKey = null;
    render().catch((err) => {
      setStatusError(err);
    });
  });

  el.mainNInput.addEventListener("input", () => {
    state.mainStyledRequestKey = null;
    render().catch((err) => {
      setStatusError(err);
    });
  });

  el.mainMonthSelect.addEventListener("change", () => {
    if (modeIsMultiMonthYears(el.mainModeSelect.value)) {
      el.mainMonthSelect.dataset.multi = readSelectedMonthValues(el.mainMonthSelect).join(",");
    } else {
      el.mainMonthSelect.dataset.current = el.mainMonthSelect.value;
    }
    state.mainStyledRequestKey = null;
    render().catch((err) => {
      setStatusError(err);
    });
  });

  el.refModeSelect.addEventListener("change", () => {
    render().catch((err) => {
      setStatusError(err);
    });
  });

  el.refNInput.addEventListener("input", () => {
    render().catch((err) => {
      setStatusError(err);
    });
  });

  el.refMonthSelect.addEventListener("change", () => {
    if (modeIsMultiMonthYears(el.refModeSelect.value)) {
      el.refMonthSelect.dataset.multi = readSelectedMonthValues(el.refMonthSelect).join(",");
    } else {
      el.refMonthSelect.dataset.current = el.refMonthSelect.value;
    }
    render().catch((err) => {
      setStatusError(err);
    });
  });

  el.metricSelect.addEventListener("change", () => {
    render().catch((err) => {
      setStatusError(err);
    });
  });

  el.searchInput.addEventListener("input", () => {
    render().catch((err) => {
      setStatusError(err);
    });
  });
}

(async function init() {
  try {
    applyThemePreference(readStoredThemePreference(), { persist: false });
    setOnboardingExpanded(false);
    setModalOpen(false);
    const storedRibbonCollapsed = readStoredRibbonCollapsed();
    setRibbonCollapsed(storedRibbonCollapsed === null ? false : storedRibbonCollapsed, { persist: false });
    syncAllScopeZooms();
    resetSelectionStatusBar();
    setRibbonTab("home");
    setViewsSplitRatio(state.viewsSplitRatio);
    syncReferenceDrawerMode();
    updateLoadHealth();
    bindEvents();
    syncFullscreenToggleButton();
    syncFloatingLayoutMetrics();
    syncRibbonFromCore();
    syncRibbonGridControlState();
    await refreshAccessContext();
    await loadWorkbookOptions();
    await loadSheets();
    await loadFiles();
    syncRibbonGridControlState();
    startRealtimePolling();
  } catch (err) {
    setStatusError(err);
  }
})();
