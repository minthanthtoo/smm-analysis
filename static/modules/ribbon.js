(function initRibbonModule(global) {
  const App = (global.SMMApp = global.SMMApp || {});
  const DEFAULT_RIBBON_TAB = "home";
  const RIBBON_TAB_GROUPS = Object.freeze([
    Object.freeze({ id: "workspace", tabs: Object.freeze(["home", "filters"]) }),
    Object.freeze({ id: "access", tabs: Object.freeze(["onboarding", "account", "roles"]) }),
    Object.freeze({ id: "files", tabs: Object.freeze(["files"]) }),
  ]);

  function copySelectOptions(source, target) {
    if (!source || !target) {
      return;
    }
    const selectedValues = new Set(
      Array.from(source.selectedOptions || [])
        .map((option) => option.value)
        .filter((value) => value !== ""),
    );
    const isMultiple = Boolean(source.multiple);
    const sourceSize = Number.parseInt(source.getAttribute("size") || "", 10);
    target.innerHTML = "";
    target.multiple = isMultiple;
    if (Number.isFinite(sourceSize) && sourceSize > 1) {
      target.setAttribute("size", String(sourceSize));
    } else {
      target.removeAttribute("size");
    }
    for (const option of Array.from(source.options || [])) {
      const next = document.createElement("option");
      next.value = option.value;
      next.textContent = option.textContent;
      next.selected = selectedValues.has(option.value);
      target.appendChild(next);
    }
    if (!isMultiple) {
      const selected = source.value;
      if (selected && Array.from(target.options).some((option) => option.value === selected)) {
        target.value = selected;
      } else if (target.options.length) {
        target.value = target.options[0].value;
      }
    }
  }

  function slugifyRibbonToken(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "");
  }

  function tabTarget(button) {
    if (!button || !(button instanceof HTMLElement)) {
      return "";
    }
    return String(button.dataset.ribbonTarget || "").trim();
  }

  function panelTarget(panel) {
    if (!panel || !(panel instanceof HTMLElement)) {
      return "";
    }
    return String(panel.dataset.ribbonPanel || "").trim();
  }

  function resolveRibbonTarget(el, name) {
    const requested = String(name || "").trim();
    if (requested) {
      for (const button of el.ribbonTabButtons || []) {
        if (tabTarget(button) === requested) {
          return requested;
        }
      }
    }
    return DEFAULT_RIBBON_TAB;
  }

  function buildRibbonTabIndex(el) {
    const index = new Map();
    for (const button of el.ribbonTabButtons || []) {
      const target = tabTarget(button);
      if (!target) {
        continue;
      }
      index.set(target, button);
    }
    return index;
  }

  function buildRibbonPanelIndex(el) {
    const index = new Map();
    for (const panel of el.ribbonPanels || []) {
      const target = panelTarget(panel);
      if (!target) {
        continue;
      }
      index.set(target, panel);
    }
    return index;
  }

  function ensureRibbonAriaLink(button, panel, target) {
    if (!button || !panel) {
      return;
    }
    const safeTarget = slugifyRibbonToken(target) || DEFAULT_RIBBON_TAB;
    if (!button.id) {
      button.id = `ribbonTab-${safeTarget}`;
    }
    if (!panel.id) {
      panel.id = `ribbonPanel-${safeTarget}`;
    }
    button.setAttribute("aria-controls", panel.id);
    panel.setAttribute("aria-labelledby", button.id);
  }

  function annotateRibbonPanelSections(panel) {
    if (!panel || !(panel instanceof HTMLElement)) {
      return;
    }
    const directGroups = Array.from(panel.children || []).filter(
      (child) => child instanceof HTMLElement && child.classList.contains("ribbon-group"),
    );
    for (let idx = 0; idx < directGroups.length; idx += 1) {
      const group = directGroups[idx];
      const heading = group.querySelector(":scope > h3");
      const sectionToken = slugifyRibbonToken(heading ? heading.textContent : "") || `section-${idx + 1}`;
      group.dataset.ribbonSection = sectionToken;
    }
  }

  function applyRibbonGrouping(el) {
    const tabHost =
      el.ribbonTabButtons && el.ribbonTabButtons.length && el.ribbonTabButtons[0] instanceof HTMLElement
        ? el.ribbonTabButtons[0].parentElement
        : null;
    if (!tabHost) {
      return;
    }

    for (const separator of Array.from(tabHost.querySelectorAll(".ribbon-tab-group-separator"))) {
      separator.remove();
    }

    const tabIndex = buildRibbonTabIndex(el);
    const panelIndex = buildRibbonPanelIndex(el);
    let previousGroupHadTabs = false;

    for (const group of RIBBON_TAB_GROUPS) {
      const groupTabs = [];
      for (const target of group.tabs) {
        const button = tabIndex.get(target);
        const panel = panelIndex.get(target);
        if (!button) {
          continue;
        }
        button.dataset.ribbonGroup = group.id;
        if (panel) {
          panel.dataset.ribbonGroup = group.id;
          ensureRibbonAriaLink(button, panel, target);
          annotateRibbonPanelSections(panel);
        }
        groupTabs.push(button);
      }

      if (previousGroupHadTabs && groupTabs.length) {
        const separator = document.createElement("span");
        separator.className = "ribbon-tab-group-separator";
        separator.setAttribute("aria-hidden", "true");
        tabHost.insertBefore(separator, groupTabs[0]);
      }
      previousGroupHadTabs = previousGroupHadTabs || groupTabs.length > 0;
    }
  }

  function activeRibbonPanel(el) {
    for (const panel of el.ribbonPanels || []) {
      if (!panel || !(panel instanceof HTMLElement)) {
        continue;
      }
      if (panel.classList.contains("active") && !panel.hidden) {
        return panel;
      }
    }
    return null;
  }

  function queueRibbonPanelScrollerSync(el) {
    if (!el) {
      return;
    }
    if (el.__ribbonPanelScrollerRaf) {
      return;
    }
    const flush = () => {
      el.__ribbonPanelScrollerRaf = 0;
      syncRibbonPanelScroller(el);
    };
    if (typeof window.requestAnimationFrame === "function") {
      el.__ribbonPanelScrollerRaf = window.requestAnimationFrame(flush);
      return;
    }
    el.__ribbonPanelScrollerRaf = window.setTimeout(flush, 16);
  }

  function syncRibbonPanelScroller(el) {
    const leftBtn = el.ribbonPanelScrollLeftBtn;
    const rightBtn = el.ribbonPanelScrollRightBtn;
    const activePanel = activeRibbonPanel(el);
    const wrap = activePanel ? activePanel.closest(".ribbon-panel-wrap") : null;
    if (!leftBtn || !rightBtn || !activePanel || !wrap) {
      if (leftBtn) {
        leftBtn.hidden = true;
      }
      if (rightBtn) {
        rightBtn.hidden = true;
      }
      return;
    }

    const maxScroll = Math.max(0, activePanel.scrollWidth - activePanel.clientWidth);
    const canScroll = maxScroll > 6;

    wrap.classList.toggle("ribbon-panel-wrap-scrollable", canScroll);
    leftBtn.hidden = !canScroll;
    rightBtn.hidden = !canScroll;
    leftBtn.classList.toggle("is-visible", canScroll);
    rightBtn.classList.toggle("is-visible", canScroll);

    if (!canScroll) {
      leftBtn.disabled = true;
      rightBtn.disabled = true;
      return;
    }

    const scrollLeft = Math.max(0, activePanel.scrollLeft);
    leftBtn.disabled = scrollLeft <= 2;
    rightBtn.disabled = scrollLeft >= maxScroll - 2;
  }

  function scrollActiveRibbonPanel(el, direction) {
    const activePanel = activeRibbonPanel(el);
    if (!activePanel) {
      return;
    }

    const step = Math.max(220, Math.floor(activePanel.clientWidth * 0.68));
    activePanel.scrollBy({
      left: step * direction,
      behavior: "smooth",
    });

    window.setTimeout(() => {
      queueRibbonPanelScrollerSync(el);
    }, 200);
  }

  function bindRibbonPanelScroller(el) {
    if (!el || el.__ribbonPanelScrollerBound) {
      return;
    }
    el.__ribbonPanelScrollerBound = true;

    if (el.ribbonPanelScrollLeftBtn) {
      el.ribbonPanelScrollLeftBtn.addEventListener("click", () => {
        scrollActiveRibbonPanel(el, -1);
      });
    }
    if (el.ribbonPanelScrollRightBtn) {
      el.ribbonPanelScrollRightBtn.addEventListener("click", () => {
        scrollActiveRibbonPanel(el, 1);
      });
    }

    for (const panel of el.ribbonPanels || []) {
      if (!panel || !(panel instanceof HTMLElement)) {
        continue;
      }
      panel.addEventListener("scroll", () => {
        queueRibbonPanelScrollerSync(el);
      });
    }

    const onWindowResize = () => {
      queueRibbonPanelScrollerSync(el);
    };
    window.addEventListener("resize", onWindowResize, { passive: true });

    if (typeof ResizeObserver !== "undefined") {
      const observer = new ResizeObserver(() => {
        queueRibbonPanelScrollerSync(el);
      });
      for (const panel of el.ribbonPanels || []) {
        if (panel && panel instanceof HTMLElement) {
          observer.observe(panel);
        }
      }
      if (el.ribbonPanelScrollLeftBtn) {
        const wrap = el.ribbonPanelScrollLeftBtn.closest(".ribbon-panel-wrap");
        if (wrap && wrap instanceof HTMLElement) {
          observer.observe(wrap);
        }
      }
      el.__ribbonPanelScrollerObserver = observer;
    }

    queueRibbonPanelScrollerSync(el);
  }

  function setRibbonTab(el, name) {
    const activeName = resolveRibbonTarget(el, name);
    for (const button of el.ribbonTabButtons || []) {
      if (!button || !(button instanceof HTMLElement)) {
        continue;
      }
      const target = tabTarget(button);
      const active = target === activeName;
      button.classList.toggle("active", active);
      button.setAttribute("aria-selected", active ? "true" : "false");
      button.setAttribute("tabindex", active ? "0" : "-1");
    }
    for (const panel of el.ribbonPanels || []) {
      if (!panel || !(panel instanceof HTMLElement)) {
        continue;
      }
      const target = panelTarget(panel);
      const active = target === activeName;
      panel.classList.toggle("active", active);
      panel.hidden = !active;
      panel.setAttribute("aria-hidden", active ? "false" : "true");
    }
    queueRibbonPanelScrollerSync(el);
  }

  function triggerChange(control) {
    if (!control) {
      return;
    }
    control.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function triggerInput(control) {
    if (!control) {
      return;
    }
    control.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function syncUploadRegionInputs(el, source) {
    const value = source && "value" in source ? source.value : "";
    if (el.uploadRegionInput && source !== el.uploadRegionInput) {
      el.uploadRegionInput.value = value;
    }
    if (el.ribbonUploadRegionInput && source !== el.ribbonUploadRegionInput) {
      el.ribbonUploadRegionInput.value = value;
    }
  }

  function syncRibbonFromCore({ el, state, roleLabel, setText }) {
    copySelectOptions(el.userSelect, el.ribbonUserSelect);
    copySelectOptions(el.regionSelect, el.ribbonRegionSelect);
    copySelectOptions(el.mainWorkbookSelect, el.ribbonMainWorkbookSelect);
    copySelectOptions(el.referenceWorkbookSelect, el.ribbonReferenceWorkbookSelect);
    copySelectOptions(el.referenceWorkbookSelect, el.referenceWorkbookMirrorSelect);
    copySelectOptions(el.mainModeSelect, el.ribbonMainModeSelect);
    copySelectOptions(el.refModeSelect, el.ribbonRefModeSelect);
    copySelectOptions(el.metricSelect, el.ribbonMetricSelect);
    copySelectOptions(el.mainMonthSelect, el.ribbonMainMonthSelect);
    copySelectOptions(el.refMonthSelect, el.ribbonRefMonthSelect);

    if (el.ribbonMainNInput && el.mainNInput) {
      el.ribbonMainNInput.value = el.mainNInput.value;
    }
    if (el.ribbonRefNInput && el.refNInput) {
      el.ribbonRefNInput.value = el.refNInput.value;
    }
    if (el.ribbonSearchInput && el.searchInput) {
      const isTyping = document.activeElement === el.ribbonSearchInput;
      if (!isTyping) {
        el.ribbonSearchInput.value = el.searchInput.value;
      }
    }

    if (el.ribbonMainMonthWrapper && el.mainMonthWrapper) {
      const hidden = el.mainMonthWrapper.classList.contains("hidden");
      el.ribbonMainMonthWrapper.classList.toggle("hidden", hidden);
    }
    if (el.ribbonRefMonthWrapper && el.refMonthWrapper) {
      const hidden = el.refMonthWrapper.classList.contains("hidden");
      el.ribbonRefMonthWrapper.classList.toggle("hidden", hidden);
    }

    if (el.ribbonUserSelect && el.userSelect) {
      el.ribbonUserSelect.disabled = el.userSelect.disabled;
    }
    if (el.ribbonRegionSelect && el.regionSelect) {
      el.ribbonRegionSelect.disabled = el.regionSelect.disabled;
    }
    if (el.referenceWorkbookMirrorSelect && el.referenceWorkbookSelect) {
      el.referenceWorkbookMirrorSelect.disabled = el.referenceWorkbookSelect.disabled;
    }

    if (el.ribbonUploadRegionInput && el.uploadRegionInput) {
      const isTypingRegion = document.activeElement === el.ribbonUploadRegionInput;
      if (!isTypingRegion) {
        el.ribbonUploadRegionInput.value = el.uploadRegionInput.value;
      }
    }

    if (el.ribbonRoleSummary) {
      const permissions = state.permissions || {};
      setText(
        el.ribbonRoleSummary,
        `${roleLabel(state.viewerRole)} · Upload: ${permissions.can_upload ? "yes" : "no"} · RSM manage: ${
          permissions.can_manage_rsm ? "yes" : "no"
        } · ASM manage: ${permissions.can_manage_asm ? "yes" : "no"}`,
      );
    }
    queueRibbonPanelScrollerSync(el);
  }

  function floatingTopOffset() {
    const root = document.documentElement;
    if (!root) {
      return 0;
    }
    const cssValue = window.getComputedStyle(root).getPropertyValue("--floating-top-total");
    const numeric = Number.parseFloat(String(cssValue || "").trim());
    return Number.isFinite(numeric) && numeric > 0 ? numeric : 0;
  }

  function scrollRibbonPanelToNode(panel, node) {
    if (!panel || !node) {
      return;
    }
    const panelRect = panel.getBoundingClientRect();
    const nodeRect = node.getBoundingClientRect();
    const currentLeft = panel.scrollLeft;
    const targetLeft = currentLeft + (nodeRect.left - panelRect.left) - 16;
    const maxLeft = Math.max(0, panel.scrollWidth - panel.clientWidth);
    const clampedLeft = Math.min(maxLeft, Math.max(0, targetLeft));
    panel.scrollTo({
      left: clampedLeft,
      behavior: "smooth",
    });
  }

  function scrollWindowToNode(node) {
    if (!node) {
      return;
    }
    const fixedOffset = floatingTopOffset();
    const rect = node.getBoundingClientRect();
    const top = window.scrollY + rect.top - fixedOffset - 10;
    window.scrollTo({
      top: Math.max(0, top),
      behavior: "smooth",
    });
  }

  function scrollToNode(node) {
    if (!node || !(node instanceof HTMLElement)) {
      return;
    }
    const ribbonPanel = node.closest(".ribbon-panel");
    if (ribbonPanel instanceof HTMLElement) {
      scrollRibbonPanelToNode(ribbonPanel, node);
      return;
    }
    scrollWindowToNode(node);
  }

  function bindRibbonTabEvents({ el, state, setRibbonTab, setRibbonCollapsed }) {
    applyRibbonGrouping(el);
    bindRibbonPanelScroller(el);

    for (const ribbonTab of el.ribbonTabButtons || []) {
      if (!ribbonTab || !(ribbonTab instanceof HTMLElement)) {
        continue;
      }
      ribbonTab.addEventListener("click", () => {
        const targetTab = tabTarget(ribbonTab) || DEFAULT_RIBBON_TAB;
        const clickedActiveTab = ribbonTab.classList.contains("active");

        setRibbonTab(targetTab);
        if (state.ribbonCollapsed) {
          setRibbonCollapsed(false);
          return;
        }
        if (clickedActiveTab) {
          setRibbonCollapsed(true);
        }
        queueRibbonPanelScrollerSync(el);
      });
      ribbonTab.addEventListener("dblclick", (event) => {
        event.preventDefault();
        setRibbonCollapsed(!state.ribbonCollapsed);
        queueRibbonPanelScrollerSync(el);
      });
    }

    if (el.ribbonToggleBtn) {
      el.ribbonToggleBtn.addEventListener("click", () => {
        setRibbonCollapsed(!state.ribbonCollapsed);
        queueRibbonPanelScrollerSync(el);
      });
    }

    queueRibbonPanelScrollerSync(el);
  }

  function bindMirrorControls({ el, triggerChange, triggerInput, syncUploadRegionInputs }) {
    const mirrorSelectChange = (ribbonControl, coreControl) => {
      if (!ribbonControl || !coreControl) {
        return;
      }
      ribbonControl.addEventListener("change", () => {
        if (ribbonControl.multiple || coreControl.multiple) {
          const selectedValues = new Set(Array.from(ribbonControl.selectedOptions || []).map((option) => option.value));
          let changed = false;
          for (const option of Array.from(coreControl.options || [])) {
            const selected = selectedValues.has(option.value);
            if (option.selected !== selected) {
              option.selected = selected;
              changed = true;
            }
          }
          if (!changed) {
            return;
          }
          triggerChange(coreControl);
          return;
        }
        if (coreControl.value === ribbonControl.value) {
          return;
        }
        coreControl.value = ribbonControl.value;
        triggerChange(coreControl);
      });
    };

    const mirrorInputValue = (ribbonControl, coreControl) => {
      if (!ribbonControl || !coreControl) {
        return;
      }
      ribbonControl.addEventListener("input", () => {
        coreControl.value = ribbonControl.value;
        triggerInput(coreControl);
      });
    };

    mirrorSelectChange(el.ribbonUserSelect, el.userSelect);
    mirrorSelectChange(el.ribbonRegionSelect, el.regionSelect);
    mirrorSelectChange(el.ribbonMainWorkbookSelect, el.mainWorkbookSelect);
    mirrorSelectChange(el.ribbonReferenceWorkbookSelect, el.referenceWorkbookSelect);
    mirrorSelectChange(el.referenceWorkbookMirrorSelect, el.referenceWorkbookSelect);
    mirrorSelectChange(el.ribbonMainModeSelect, el.mainModeSelect);
    mirrorSelectChange(el.ribbonMainMonthSelect, el.mainMonthSelect);
    mirrorSelectChange(el.ribbonRefModeSelect, el.refModeSelect);
    mirrorSelectChange(el.ribbonRefMonthSelect, el.refMonthSelect);
    mirrorSelectChange(el.ribbonMetricSelect, el.metricSelect);
    mirrorInputValue(el.ribbonMainNInput, el.mainNInput);
    mirrorInputValue(el.ribbonRefNInput, el.refNInput);
    mirrorInputValue(el.ribbonSearchInput, el.searchInput);

    if (el.ribbonUploadRegionInput) {
      el.ribbonUploadRegionInput.addEventListener("input", () => {
        syncUploadRegionInputs(el.ribbonUploadRegionInput);
      });
    }
    if (el.uploadRegionInput) {
      el.uploadRegionInput.addEventListener("input", () => {
        syncUploadRegionInputs(el.uploadRegionInput);
      });
    }
  }

  App.ribbon = {
    copySelectOptions,
    setRibbonTab,
    triggerChange,
    triggerInput,
    syncUploadRegionInputs,
    syncRibbonFromCore,
    scrollToNode,
    bindRibbonTabEvents,
    bindMirrorControls,
  };
})(window);
