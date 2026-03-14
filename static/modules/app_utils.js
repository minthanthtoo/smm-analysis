(function initAppUtilsModule(global) {
  const App = (global.SMMApp = global.SMMApp || {});

  function formatClockTime(value) {
    if (!(value instanceof Date)) {
      return "";
    }
    return value.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  }

  function formatMetricNumber(value) {
    if (!Number.isFinite(value)) {
      return "-";
    }
    if (Math.abs(value) >= 1000) {
      return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
    }
    return value.toLocaleString(undefined, { maximumFractionDigits: 4 });
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

  function escapeHtml(value) {
    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
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

  function isNumericCellText(text) {
    const normalized = String(text || "")
      .replace(/,/g, "")
      .trim();
    return /^-?\d+(?:\.\d+)?$/.test(normalized);
  }

  function selectionScopeLabel(scope) {
    if (scope === "main") {
      return "Main Selection";
    }
    if (scope === "reference") {
      return "Detail Selection";
    }
    return "No Selection";
  }

  function normalizeViewScope(scope) {
    return scope === "reference" ? "reference" : "main";
  }

  function columnToExcelLabel(index) {
    let value = Math.max(0, Number.parseInt(String(index), 10)) + 1;
    let label = "";
    while (value > 0) {
      const rem = (value - 1) % 26;
      label = String.fromCharCode(65 + rem) + label;
      value = Math.floor((value - 1) / 26);
    }
    return label || "A";
  }

  function pointToAddress(point) {
    if (!point) {
      return "";
    }
    return `${columnToExcelLabel(point.col)}${point.row + 1}`;
  }

  function parseNumericCellValue(cellText) {
    const compact = String(cellText || "")
      .replace(/,/g, "")
      .trim();
    if (!compact) {
      return null;
    }
    const normalized =
      compact.startsWith("(") && compact.endsWith(")") ? `-${compact.slice(1, compact.length - 1)}` : compact;
    if (!/^-?\d+(?:\.\d+)?$/.test(normalized)) {
      return null;
    }
    const value = Number.parseFloat(normalized);
    return Number.isFinite(value) ? value : null;
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

  App.utils = {
    ...(App.utils || {}),
    formatClockTime,
    formatMetricNumber,
    normalizeSheetName,
    sanitizeSheetTabs,
    roleLabel,
    normalizeRoleToken,
    regionLabel,
    escapeHtml,
    modeIsSameMonthYears,
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
  };
})(window);
