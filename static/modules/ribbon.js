(function initRibbonModule(global) {
  const App = (global.SMMApp = global.SMMApp || {});

  function copySelectOptions(source, target) {
    if (!source || !target) {
      return;
    }
    const selected = source.value;
    target.innerHTML = "";
    for (const option of Array.from(source.options || [])) {
      const next = document.createElement("option");
      next.value = option.value;
      next.textContent = option.textContent;
      target.appendChild(next);
    }
    if (selected && Array.from(target.options).some((option) => option.value === selected)) {
      target.value = selected;
    } else if (target.options.length) {
      target.value = target.options[0].value;
    }
  }

  function setRibbonTab(el, name) {
    const activeName = name || "home";
    for (const button of el.ribbonTabButtons || []) {
      if (!button || !(button instanceof HTMLElement)) {
        continue;
      }
      const target = button.dataset.ribbonTarget || "";
      const active = target === activeName;
      button.classList.toggle("active", active);
      button.setAttribute("aria-selected", active ? "true" : "false");
    }
    for (const panel of el.ribbonPanels || []) {
      if (!panel || !(panel instanceof HTMLElement)) {
        continue;
      }
      const target = panel.dataset.ribbonPanel || "";
      const active = target === activeName;
      panel.classList.toggle("active", active);
      panel.hidden = !active;
    }
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
  }

  function scrollToNode(node) {
    if (!node || typeof node.scrollIntoView !== "function") {
      return;
    }
    node.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function bindRibbonTabEvents({ el, state, setRibbonTab, setRibbonCollapsed }) {
    for (const ribbonTab of el.ribbonTabButtons || []) {
      if (!ribbonTab || !(ribbonTab instanceof HTMLElement)) {
        continue;
      }
      ribbonTab.addEventListener("click", () => {
        const targetTab = ribbonTab.dataset.ribbonTarget || "home";
        const clickedActiveTab = ribbonTab.classList.contains("active");

        setRibbonTab(targetTab);
        if (state.ribbonCollapsed) {
          setRibbonCollapsed(false);
          return;
        }
        if (clickedActiveTab) {
          setRibbonCollapsed(true);
        }
      });
      ribbonTab.addEventListener("dblclick", (event) => {
        event.preventDefault();
        setRibbonCollapsed(!state.ribbonCollapsed);
      });
    }

    if (el.ribbonToggleBtn) {
      el.ribbonToggleBtn.addEventListener("click", () => {
        setRibbonCollapsed(!state.ribbonCollapsed);
      });
    }
  }

  function bindMirrorControls({ el, triggerChange, triggerInput, syncUploadRegionInputs }) {
    const mirrorSelectChange = (ribbonControl, coreControl) => {
      if (!ribbonControl || !coreControl) {
        return;
      }
      ribbonControl.addEventListener("change", () => {
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
