(function initThemeModule(global) {
  const App = (global.SMMApp = global.SMMApp || {});

  function createThemeApi({ state, el, themeMediaQuery, storageKey, scheduleLayoutSync }) {
    function normalizeThemePreference(value) {
      if (value === "dark" || value === "light" || value === "system") {
        return value;
      }
      return "system";
    }

    function resolveEffectiveTheme(preference) {
      const normalized = normalizeThemePreference(preference);
      if (normalized === "dark" || normalized === "light") {
        return normalized;
      }
      if (themeMediaQuery) {
        return themeMediaQuery.matches ? "dark" : "light";
      }
      return "dark";
    }

    function readStoredThemePreference() {
      try {
        const raw = window.localStorage ? window.localStorage.getItem(storageKey) : null;
        return normalizeThemePreference(raw || "system");
      } catch (_err) {
        return "system";
      }
    }

    function writeStoredThemePreference(preference) {
      try {
        if (window.localStorage) {
          window.localStorage.setItem(storageKey, normalizeThemePreference(preference));
        }
      } catch (_err) {
        // Ignore storage write errors (private mode or blocked storage).
      }
    }

    function applyThemePreference(preference, options = {}) {
      const normalized = normalizeThemePreference(preference);
      const effective = resolveEffectiveTheme(normalized);
      state.themePreference = normalized;
      state.effectiveTheme = effective;

      if (document.body) {
        document.body.classList.remove("theme-system", "theme-dark", "theme-light");
        document.body.classList.add(`theme-${normalized}`);
        document.body.setAttribute("data-theme", effective);
      }

      if (el.themeSelect && el.themeSelect.value !== normalized) {
        el.themeSelect.value = normalized;
      }

      if (options.persist !== false) {
        writeStoredThemePreference(normalized);
      }

      scheduleLayoutSync();
    }

    return {
      normalizeThemePreference,
      resolveEffectiveTheme,
      readStoredThemePreference,
      writeStoredThemePreference,
      applyThemePreference,
    };
  }

  App.createThemeApi = createThemeApi;
})(window);
