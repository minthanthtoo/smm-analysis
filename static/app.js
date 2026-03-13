const state = {
  workbooks: [],
  regions: [],
  allRegions: [],
  users: [],
  assignments: {
    rsm_regions: {},
    user_to_rsm: {},
    asm_townships: {},
  },
  permissions: {
    can_upload: false,
    can_manage_rsm: false,
    can_manage_asm: false,
    can_manage_files: false,
  },
  viewerRole: "owner",
  currentUser: "owner",
  currentUserDisplayName: "Owner",
  canViewAllRegions: true,
  selectedRegion: "ALL",
  selectedMainWorkbook: null,
  selectedReferenceWorkbook: null,
  regionTownships: {},
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
  busyDepth: 0,
  lastLoadAt: null,
  fileRows: [],
};

const POLL_INTERVAL_MS = 30000;

const el = {
  userSelect: document.getElementById("userSelect"),
  regionSelect: document.getElementById("regionSelect"),
  uploadInput: document.getElementById("uploadInput"),
  uploadRegionInput: document.getElementById("uploadRegionInput"),
  uploadBtn: document.getElementById("uploadBtn"),
  onboardingToggleBtn: document.getElementById("onboardingToggleBtn"),
  onboardingSteps: document.getElementById("onboardingSteps"),
  roleInfoBtn: document.getElementById("roleInfoBtn"),
  roleInfoModal: document.getElementById("roleInfoModal"),
  roleInfoCloseBtn: document.getElementById("roleInfoCloseBtn"),
  mainWorkbookSelect: document.getElementById("mainWorkbookSelect"),
  referenceWorkbookSelect: document.getElementById("referenceWorkbookSelect"),
  mainWorkbookName: document.getElementById("mainWorkbookName"),
  referenceWorkbookName: document.getElementById("referenceWorkbookName"),
  currentUserName: document.getElementById("currentUserName"),
  viewerRoleName: document.getElementById("viewerRoleName"),
  regionName: document.getElementById("regionName"),
  roleOnboardingText: document.getElementById("roleOnboardingText"),
  roleCardOwner: document.getElementById("roleCardOwner"),
  roleCardRsm: document.getElementById("roleCardRsm"),
  roleCardAsm: document.getElementById("roleCardAsm"),
  roleCardUser: document.getElementById("roleCardUser"),
  mainLoadChip: document.getElementById("mainLoadChip"),
  referenceLoadChip: document.getElementById("referenceLoadChip"),
  mainTabsChip: document.getElementById("mainTabsChip"),
  referenceTabsChip: document.getElementById("referenceTabsChip"),
  loadStageChip: document.getElementById("loadStageChip"),
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
  statusHint: document.getElementById("statusHint"),
  mainTable: document.getElementById("mainTable"),
  refTable: document.getElementById("refTable"),
  mainMeta: document.getElementById("mainMeta"),
  refMeta: document.getElementById("refMeta"),
  accessHint: document.getElementById("accessHint"),
  usersList: document.getElementById("usersList"),
  assignRsmForm: document.getElementById("assignRsmForm"),
  assignRsmUsernameInput: document.getElementById("assignRsmUsernameInput"),
  assignRsmDisplayInput: document.getElementById("assignRsmDisplayInput"),
  assignRsmRegionsInput: document.getElementById("assignRsmRegionsInput"),
  mapUserRsmForm: document.getElementById("mapUserRsmForm"),
  mapUserUsernameInput: document.getElementById("mapUserUsernameInput"),
  mapUserRsmSelect: document.getElementById("mapUserRsmSelect"),
  assignAsmForm: document.getElementById("assignAsmForm"),
  assignAsmUsernameInput: document.getElementById("assignAsmUsernameInput"),
  assignAsmDisplayInput: document.getElementById("assignAsmDisplayInput"),
  assignAsmRsmWrap: document.getElementById("assignAsmRsmWrap"),
  assignAsmRsmSelect: document.getElementById("assignAsmRsmSelect"),
  assignAsmRegionSelect: document.getElementById("assignAsmRegionSelect"),
  assignAsmTownshipsSelect: document.getElementById("assignAsmTownshipsSelect"),
  asmTownshipForm: document.getElementById("asmTownshipForm"),
  asmTownshipUserSelect: document.getElementById("asmTownshipUserSelect"),
  asmTownshipRegionSelect: document.getElementById("asmTownshipRegionSelect"),
  asmTownshipSelect: document.getElementById("asmTownshipSelect"),
  refreshFilesBtn: document.getElementById("refreshFilesBtn"),
  filesTableBody: document.getElementById("filesTableBody"),
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

function setChip(node, text, tone) {
  if (!node) {
    return;
  }
  node.textContent = text;
  if (tone) {
    node.dataset.state = tone;
  }
}

function formatClockTime(value) {
  if (!(value instanceof Date)) {
    return "";
  }
  return value.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
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

function setModalOpen(open) {
  if (!el.roleInfoModal) {
    return;
  }
  if (open) {
    el.roleInfoModal.classList.remove("hidden");
    el.roleInfoModal.setAttribute("aria-hidden", "false");
    return;
  }
  el.roleInfoModal.classList.add("hidden");
  el.roleInfoModal.setAttribute("aria-hidden", "true");
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

function setInteractiveControlsDisabled(disabled) {
  const controls = [
    el.userSelect,
    el.regionSelect,
    el.uploadInput,
    el.uploadRegionInput,
    el.uploadBtn,
    el.mainWorkbookSelect,
    el.referenceWorkbookSelect,
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
    el.assignRsmUsernameInput,
    el.assignRsmDisplayInput,
    el.assignRsmRegionsInput,
    el.mapUserUsernameInput,
    el.mapUserRsmSelect,
    el.assignAsmUsernameInput,
    el.assignAsmDisplayInput,
    el.assignAsmRsmSelect,
    el.assignAsmRegionSelect,
    el.assignAsmTownshipsSelect,
    el.asmTownshipUserSelect,
    el.asmTownshipRegionSelect,
    el.asmTownshipSelect,
    el.refreshFilesBtn,
  ];
  for (const control of controls) {
    if (!control) {
      continue;
    }
    control.disabled = Boolean(disabled);
  }
}

function normalizeSheetName(value) {
  if (typeof value !== "string") {
    return "";
  }
  return value;
}

function sanitizeSheetTabs(tabs, allowedNames) {
  const nameSet = new Set(
    Array.isArray(allowedNames)
      ? allowedNames.map((item) => normalizeSheetName(item)).filter(Boolean)
      : [],
  );
  const useAllowedNames = nameSet.size > 0;
  const normalized = [];
  const seen = new Set();
  for (const rawTab of Array.isArray(tabs) ? tabs : []) {
    if (!rawTab || typeof rawTab !== "object") {
      continue;
    }
    const sheetName = normalizeSheetName(rawTab.sheet_name);
    if (!sheetName || seen.has(sheetName)) {
      continue;
    }
    if (useAllowedNames && !nameSet.has(sheetName)) {
      continue;
    }
    seen.add(sheetName);
    normalized.push({
      sheet_name: sheetName,
      canonical: typeof rawTab.canonical === "string" && rawTab.canonical ? rawTab.canonical : null,
      filterable: Boolean(rawTab.filterable),
      has_reference: Boolean(rawTab.has_reference),
      has_main: Boolean(rawTab.has_main),
    });
  }
  return normalized;
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

function roleLabel(role) {
  if (role === "regional_manager") {
    return "RSM";
  }
  if (role === "rsm") {
    return "RSM";
  }
  if (role === "asm") {
    return "ASM";
  }
  if (role === "user") {
    return "User";
  }
  return "Owner";
}

function normalizeRoleToken(role) {
  if (role === "regional_manager") {
    return "rsm";
  }
  if (role === "owner" || role === "rsm" || role === "asm" || role === "user") {
    return role;
  }
  return "owner";
}

function regionLabel(region) {
  return region === "ALL" ? "All regions" : region || "-";
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

function populateWorkbookSelect(selectElement, selectedName) {
  selectElement.innerHTML = "";
  for (const workbook of state.workbooks) {
    const option = document.createElement("option");
    option.value = workbook;
    option.textContent = workbook;
    selectElement.appendChild(option);
  }
  if (selectedName && state.workbooks.includes(selectedName)) {
    selectElement.value = selectedName;
    return;
  }
  if (state.workbooks.length) {
    selectElement.value = state.workbooks[0];
  }
}

function displayUserOption(user) {
  const username = user && typeof user.username === "string" ? user.username : "";
  const displayName = user && typeof user.display_name === "string" ? user.display_name : "";
  const role = user && typeof user.role === "string" ? roleLabel(user.role) : "";
  if (displayName && displayName !== username) {
    return `${displayName} (${username}) · ${role}`;
  }
  return `${username || "-"} · ${role}`;
}

function populateUserSelect() {
  if (!el.userSelect) {
    return;
  }
  const users = Array.isArray(state.users) ? state.users : [];
  el.userSelect.innerHTML = "";

  for (const user of users) {
    if (!user || typeof user !== "object" || !user.username) {
      continue;
    }
    const option = document.createElement("option");
    option.value = user.username;
    option.textContent = displayUserOption(user);
    el.userSelect.appendChild(option);
  }

  if (!users.some((user) => user.username === state.currentUser)) {
    const fallback = users[0];
    state.currentUser = fallback && fallback.username ? fallback.username : state.currentUser;
  }
  if (state.currentUser) {
    el.userSelect.value = state.currentUser;
  }
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
    option.textContent = region;
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
  if (Array.isArray(payload.regions)) {
    state.regions = payload.regions;
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
  setText(el.currentUserName, state.currentUserDisplayName || state.currentUser || "-");
  setText(el.viewerRoleName, roleLabel(state.viewerRole));
  setText(el.regionName, regionLabel(state.selectedRegion));
  setText(el.mainPanelTitle, `Main View (${state.selectedMainWorkbook || "-"})`);
  setText(el.referencePanelTitle, `Reference Detail (${state.selectedReferenceWorkbook || "-"})`);
  updateLoadHealth();
}

function setSelectOptions(selectElement, values, selectedValue = null) {
  if (!selectElement) {
    return;
  }
  selectElement.innerHTML = "";
  for (const value of values) {
    const option = document.createElement("option");
    option.value = String(value);
    option.textContent = String(value);
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

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
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

  setSelectOptionsFromUsers(el.mapUserRsmSelect, rsmUsers, null, el.mapUserRsmSelect?.value || null);
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
      '<tr><td colspan="3" class="empty">No files found in this scope.</td></tr>';
    return;
  }

  const regionOptions = filesEditableRegionOptions();
  const rows = files
    .map((file) => {
      const nameEscaped = escapeHtml(file.name);
      const regionEscaped = escapeHtml(file.region);
      let regionControl = regionEscaped;
      let actions = '<span class="muted-inline">View only</span>';
      if (file.can_update) {
        const selectOptions = regionOptions
          .map((region) => {
            const selected = region === file.region ? ' selected="selected"' : "";
            return `<option value="${escapeHtml(region)}"${selected}>${escapeHtml(region)}</option>`;
          })
          .join("");
        regionControl = `<select class="file-region-select" data-file="${nameEscaped}">${selectOptions}</select>`;
      }
      if (file.can_update || file.can_delete) {
        actions = "";
        if (file.can_update) {
          actions += `<button type="button" class="action-btn action-btn-sm file-save-btn" data-file="${nameEscaped}">Save region</button>`;
        }
        if (file.can_delete) {
          actions += `<button type="button" class="action-btn action-btn-sm action-btn-danger file-delete-btn" data-file="${nameEscaped}">Delete</button>`;
        }
      }
      return `<tr><td class="left">${nameEscaped}</td><td class="left">${regionControl}</td><td class="left file-action-cell">${actions}</td></tr>`;
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
}

function setEmptyRef(message) {
  unwrapSplitViewport(el.refTable);
  el.refTable.innerHTML = `<tbody><tr><td class="empty">${message}</td></tr></tbody>`;
  setText(el.refMeta, "");
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
  for (const row of table.rows || []) {
    row.style.removeProperty("height");
  }
  for (const cell of table.querySelectorAll("th, td")) {
    cell.classList.remove("sticky-col", "sticky-col-boundary", "sticky-col-head", "split-head-cell");
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
        setStatusError(err);
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
        setStatusError(err);
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
  beginBusy("Excel file များကို ရှာဖွေနေသည်...");
  const scope = scopeQuery();
  try {
    const payload = await fetchJson(`/api/workbooks?${scope}`);
    applyScopePayload(payload);
    state.workbooks = payload.workbooks || [];

    if (!state.workbooks.length) {
      state.selectedMainWorkbook = null;
      state.selectedReferenceWorkbook = null;
      populateWorkbookSelect(el.mainWorkbookSelect, null);
      populateWorkbookSelect(el.referenceWorkbookSelect, null);
      updateWorkbookLabels();
      setEmptyMain("No workbook available in current scope.");
      setEmptyRef("No workbook available in current scope.");
      return;
    }

    state.selectedMainWorkbook = payload.default_main || state.workbooks[0];
    state.selectedReferenceWorkbook = payload.default_reference || state.workbooks[0];

    populateWorkbookSelect(el.mainWorkbookSelect, state.selectedMainWorkbook);
    populateWorkbookSelect(el.referenceWorkbookSelect, state.selectedReferenceWorkbook);
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

  const files = Array.from(el.uploadInput.files || []);
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
  const uploadRegionValue = (el.uploadRegionInput.value || "").trim();
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
    state.workbooks = payload.workbooks || state.workbooks;
    state.selectedMainWorkbook = payload.default_main || state.selectedMainWorkbook;
    state.selectedReferenceWorkbook = payload.default_reference || state.selectedReferenceWorkbook;
    state.cache.clear();
    state.mainStyledRequestKey = null;
    state.mainAvailableMonths.clear();

    populateWorkbookSelect(el.mainWorkbookSelect, state.selectedMainWorkbook);
    populateWorkbookSelect(el.referenceWorkbookSelect, state.selectedReferenceWorkbook);
    updateWorkbookLabels();
    await loadSheets();
    await refreshAccessAndFiles();

    el.uploadInput.value = "";
    el.uploadRegionInput.value = "";
    const skippedCount = Array.isArray(payload.skipped_files) ? payload.skipped_files.length : 0;
    const skippedText = skippedCount ? ` · skipped ${skippedCount} unsupported file(s)` : "";
    const uploadedRegions = Array.isArray(payload.uploaded_regions)
      ? [...new Set(payload.uploaded_regions.map((item) => (item && item.region ? String(item.region) : "")))]
          .filter(Boolean)
          .join(", ")
      : "";
    const uploadRegionText = uploadedRegions || regionLabel(state.selectedRegion);
    setText(el.statusText, `Uploaded ${files.length} file(s) to ${uploadRegionText}${skippedText}`);
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

  if (payload.main_workbook) {
    state.selectedMainWorkbook = payload.main_workbook;
  }
  if (payload.reference_workbook) {
    state.selectedReferenceWorkbook = payload.reference_workbook;
  }

  populateWorkbookSelect(el.mainWorkbookSelect, state.selectedMainWorkbook);
  populateWorkbookSelect(el.referenceWorkbookSelect, state.selectedReferenceWorkbook);
  updateWorkbookLabels();
  const changed = previousVersion !== state.pairVersion;
  if (changed) {
    state.mainStyledRequestKey = null;
    state.mainAvailableMonths.clear();
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
    row.style.removeProperty("height");
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

function isNumericCellText(text) {
  const normalized = String(text || "")
    .replace(/,/g, "")
    .trim();
  return /^-?\d+(?:\.\d+)?$/.test(normalized);
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
  const root = getComputedStyle(document.documentElement);
  const stickyBg = root.getPropertyValue("--sticky-bg").trim();
  if (stickyBg) {
    return stickyBg;
  }
  const fallback = root.getPropertyValue("--bg-soft").trim();
  if (fallback) {
    return fallback;
  }
  const tableBg = getComputedStyle(table).backgroundColor;
  if (tableBg && tableBg !== "transparent" && tableBg !== "rgba(0, 0, 0, 0)") {
    return tableBg;
  }
  return "rgb(16, 39, 51)";
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
  setText(target, `${sheetData.sheet_name} · rows ${sheetData.rows.length} · showing ${selectedMonths.length} month group(s)`);
}

function mainStyledKey(selectedMonths) {
  const monthsPart = selectedMonths.map((item) => item.key).join(",");
  const mode = el.mainModeSelect.value;
  const nValue = currentN(el.mainNInput);
  const monthValue = modeIsSameMonthYears(mode) ? el.mainMonthSelect.value || "auto" : "none";
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
  enhanceFrozenViewport(mainSourceTable, frozenCount);

  const monthCount = (payload.selected_month_labels || []).length;
  if (monthCount > 0) {
    setText(el.mainMeta, `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · ${monthCount} month group(s)`);
  } else if (payload.filterable) {
    setText(el.mainMeta, `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · fixed-layout mode`);
  } else {
    setText(el.mainMeta, `${payload.sheet_name} · rows ${payload.row_count} · ${payload.col_count} columns · full-sheet mode`);
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

  renderTable(el.refTable, referenceSheet, selectedMonthsRef);
  enhanceFrozenViewport(el.refTable, 4);
  renderMeta(el.refMeta, referenceSheet, selectedMonthsRef);

  if (!selectedMonthsRef.length) {
    appendText(el.refMeta, " · no months matched this mode");
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
  updateWorkbookLabels();
  await loadSheets();
}

function bindEvents() {
  if (el.onboardingToggleBtn) {
    el.onboardingToggleBtn.addEventListener("click", () => {
      const currentlyHidden = el.onboardingSteps ? el.onboardingSteps.classList.contains("hidden") : true;
      setOnboardingExpanded(currentlyHidden);
    });
  }
  if (el.roleInfoBtn) {
    el.roleInfoBtn.addEventListener("click", () => {
      setModalOpen(true);
    });
  }
  if (el.roleInfoCloseBtn) {
    el.roleInfoCloseBtn.addEventListener("click", () => {
      setModalOpen(false);
    });
  }
  if (el.roleInfoModal) {
    el.roleInfoModal.addEventListener("click", (event) => {
      if (event.target === el.roleInfoModal) {
        setModalOpen(false);
      }
    });
  }
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      setModalOpen(false);
    }
  });

  if (el.userSelect) {
    el.userSelect.addEventListener("change", () => {
      onScopeChange().catch((err) => {
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
        const select = row ? row.querySelector(".file-region-select") : null;
        const region = select && "value" in select ? select.value : "";
        if (!filename || !region) {
          return;
        }
        (async () => {
          beginBusy("Updating file region...");
          try {
            await patchScopedJson(`/api/files/${encodeURIComponent(filename)}`, { region });
            await loadFiles();
            await loadWorkbookOptions();
            await loadSheets();
            setText(el.statusText, `Updated region for ${filename}.`);
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
    el.mainMonthSelect.dataset.current = el.mainMonthSelect.value;
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
    el.refMonthSelect.dataset.current = el.refMonthSelect.value;
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
    setOnboardingExpanded(false);
    updateLoadHealth();
    bindEvents();
    await refreshAccessContext();
    await loadWorkbookOptions();
    await loadSheets();
    await loadFiles();
    startRealtimePolling();
  } catch (err) {
    setStatusError(err);
  }
})();
