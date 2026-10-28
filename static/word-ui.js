(() => {
  "use strict";

  const root = document.querySelector(".xp-desktop");
  if (!root) return;

  if (window.__wordUiCleanup) {
    try {
      window.__wordUiCleanup();
    } catch (_) {
      /* A previous render is already gone. */
    }
  }

  const cleanupFns = [];
  const on = (target, type, handler, options) => {
    target.addEventListener(type, handler, options);
    cleanupFns.push(() => target.removeEventListener(type, handler, options));
  };
  const $ = (selector, context = root) => context.querySelector(selector);
  const $$ = (selector, context = root) => Array.from(context.querySelectorAll(selector));

  const paper = $("#paper-stack");
  const editor = $("#editor-scroll");
  const titleText = $(".title-text");
  const taskLabel = $(".task-label");
  const taskButton = $("#task-button");
  const wordWindow = $(".word-window");
  const titlebar = $("#word-titlebar");
  const menuRoot = $(".menubar");
  const docmap = $("#docmap");
  const docmapFilter = $(".docmap-filter");
  const rulerShell = $("#ruler-shell");
  const standardToolbar = $(".toolbar-standard");
  const formatToolbar = $(".toolbar-format");
  const statusPage = $("#status-page");
  const statusPos = $("#status-pos");
  const statusLine = $("#status-line");
  const statusCol = $("#status-col");
  const statusZoom = $("#status-zoom");
  const statusLang = $("#status-lang");
  const startButton = $("#start-button");
  const startMenu = $("#start-menu");
  const clippy = $("#clippy");
  const desktopReopen = $("#desktop-reopen");
  const desktopShortcut = $(".desktop-shortcut", desktopReopen);
  const dialogLayer = $("#dialog-layer");
  const dialogBox = $("#word-dialog");
  const dialogTitlebar = $("#dialog-titlebar");
  const dialogTitle = $("#dialog-title");
  const dialogBody = $("#dialog-body");
  const dialogButtons = $("#dialog-buttons");
  const dialogClose = $(".dialog-close");
  const toast = $("#word-toast");
  const fileInput = $("#open-file");
  const comboPopup = $("#combo-popup");
  const shutdownScreen = $("#shutdown-screen");
  const shutdownMessage = $("#shutdown-message");
  const clockButton = $("#clock");
  const maxButton = $(".wc-max");
  const undoButton = $('[data-action="edit-undo"]');
  const redoButton = $('[data-action="edit-redo"]');
  const GITHUB_URL = "https://github.com/maidenator";
  const REPOSITORIES_URL = "https://github.com/maidenator?tab=repositories";

  paper.setAttribute("contenteditable", "true");
  paper.spellcheck = true;

  const originalHTML = paper.innerHTML;
  let currentFileName = "Portfolio.doc";
  let dirty = false;
  let savedRange = null;
  let currentPage = 1;
  let pageCount = 2;
  let zoomPercent = 100;
  let toastTimer = 0;
  let statusFrame = 0;
  let selectionFrame = 0;
  let dialogState = null;
  let comboState = null;
  let menuName = null;
  let docmapView = "headings";
  let dialogOffset = { x: 0, y: 0 };
  let dragState = null;
  let shortcutDrag = null;
  let suppressShortcutClick = false;
  let windowDrag = null;
  let windowOffset = { x: 0, y: 0 };
  let suppressTitlebarDblClick = false;
  let clockTimer = 0;

  const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)");

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function escapeAttr(value) {
    return escapeHtml(value).replace(/\n/g, " ");
  }

  function showToast(message, duration = 2800) {
    if (!toast) return;
    window.clearTimeout(toastTimer);
    toast.textContent = message;
    toast.hidden = false;
    toastTimer = window.setTimeout(() => {
      toast.hidden = true;
    }, duration);
  }

  function updateTitle() {
    const heading = `${currentFileName}${dirty ? "*" : ""} - Microsoft Word`;
    titleText.textContent = heading;
    taskLabel.textContent = heading;
    document.title = heading;
  }

  function markDirty() {
    if (!dirty) {
      dirty = true;
      updateTitle();
    }
  }

  function currentSelection() {
    return window.getSelection();
  }

  function selectionInsidePaper(sel = currentSelection()) {
    if (!sel || !sel.rangeCount || !sel.anchorNode) return false;
    return paper.contains(sel.anchorNode) && paper.contains(sel.focusNode);
  }

  function rememberSelection() {
    const sel = currentSelection();
    if (selectionInsidePaper(sel)) {
      savedRange = sel.getRangeAt(0).cloneRange();
    }
  }

  function restoreSelection() {
    if (!savedRange) return false;
    const sel = currentSelection();
    paper.focus({ preventScroll: true });
    sel.removeAllRanges();
    try {
      sel.addRange(savedRange);
    } catch (_) {
      return false;
    }
    return true;
  }

  function ensureDocumentSelection() {
    if (selectionInsidePaper()) return true;
    if (savedRange && restoreSelection()) return true;
    showToast("Click in the document first, then try that command.");
    return false;
  }

  function caretElement() {
    const sel = currentSelection();
    if (!sel || !sel.rangeCount || !sel.anchorNode) return null;
    const node = sel.anchorNode;
    return node.nodeType === 1 ? node : node.parentElement;
  }

  function currentTable() {
    const element = caretElement();
    return element ? element.closest("table") : null;
  }

  function currentCell() {
    const element = caretElement();
    return element ? element.closest("td, th") : null;
  }

  function afterEdit() {
    markDirty();
    refreshToolbar();
    scheduleStatus();
  }

  function executeCommand(command, value = null) {
    if (!ensureDocumentSelection()) return false;
    try {
      document.execCommand("styleWithCSS", false, true);
      document.execCommand(command, false, value);
    } catch (error) {
      console.error(error);
      showToast("Word could not run that command in this browser.");
      return false;
    }
    afterEdit();
    return true;
  }

  function insertHtml(html) {
    if (!ensureDocumentSelection()) return false;
    try {
      document.execCommand("insertHTML", false, html);
    } catch (error) {
      console.error(error);
      showToast("Word could not insert that content.");
      return false;
    }
    afterEdit();
    return true;
  }

  function wrapSelectionStyle(styleText) {
    if (!ensureDocumentSelection()) return false;
    const sel = currentSelection();
    const range = sel.getRangeAt(0);
    const span = document.createElement("span");
    span.style.cssText = styleText;
    if (range.collapsed) {
      span.appendChild(document.createTextNode("\u200b"));
      range.insertNode(span);
      const caret = document.createRange();
      caret.setStart(span.firstChild, 1);
      caret.collapse(true);
      sel.removeAllRanges();
      sel.addRange(caret);
    } else {
      try {
        range.surroundContents(span);
      } catch (_) {
        const fragment = range.extractContents();
        span.appendChild(fragment);
        range.insertNode(span);
      }
      sel.removeAllRanges();
      const select = document.createRange();
      select.selectNodeContents(span);
      sel.addRange(select);
    }
    afterEdit();
    return true;
  }

  function applyFontFamily(font) {
    return executeCommand("fontName", font);
  }

  function applyFontSize(points) {
    if (!ensureDocumentSelection()) return false;
    const sel = currentSelection();
    if (sel.isCollapsed) {
      return wrapSelectionStyle(`font-size:${points}pt`);
    }
    try {
      document.execCommand("styleWithCSS", false, true);
      document.execCommand("fontSize", false, "7");
    } catch (error) {
      console.error(error);
      return wrapSelectionStyle(`font-size:${points}pt`);
    }
    $$("span", paper).forEach((span) => {
      if (span.style.fontSize === "xxx-large") {
        span.style.fontSize = `${points}pt`;
      }
    });
    afterEdit();
    return true;
  }

  function applyBlockStyle(tag) {
    return executeCommand("formatBlock", `<${tag}>`);
  }

  function setCommandState(command, wanted) {
    if (!ensureDocumentSelection()) return;
    let active = false;
    try {
      active = document.queryCommandState(command);
    } catch (_) {
      active = false;
    }
    if (active !== wanted) {
      document.execCommand(command, false, null);
    }
  }

  function setDirtyValue() {
    markDirty();
    refreshToolbar();
    scheduleStatus();
  }

  function closeDialog(value) {
    if (dialogState) dialogState.finish(value);
  }

  function showDialog(options) {
    if (dialogState) closeDialog(null);
    dialogLayer.hidden = false;
    dialogBox.style.width = options.width || "380px";
    dialogBox.style.transform = "";
    dialogOffset = { x: 0, y: 0 };
    dialogTitle.textContent = options.title || "Microsoft Word";
    dialogBody.innerHTML = options.body || "";
    dialogButtons.innerHTML = "";

    return new Promise((resolve) => {
      const finish = (value) => {
        if (!dialogState || dialogState.finish !== finish) return;
        dialogLayer.hidden = true;
        dialogState = null;
        dialogBody.innerHTML = "";
        dialogButtons.innerHTML = "";
        dialogBox.style.transform = "";
        resolve(value);
      };

      dialogState = { finish };

      (options.buttons || []).forEach((button) => {
        const element = document.createElement("button");
        element.type = "button";
        element.className = `dialog-button${button.primary ? " primary" : ""}`;
        element.textContent = button.label;
        element.addEventListener("click", () => {
          if (button.collect) {
            const value = button.collect(dialogBody);
            if (value === false) return;
            finish(value);
            return;
          }
          finish(button.value === undefined ? null : button.value);
        });
        dialogButtons.appendChild(element);
      });

      if (options.onMount) options.onMount(dialogBody, finish);

      window.requestAnimationFrame(() => {
        const focus = options.focus
          ? dialogBody.querySelector(options.focus)
          : dialogBody.querySelector("input:not([type=hidden]), select, textarea, button");
        const target = focus || dialogButtons.querySelector("button");
        if (target) {
          target.focus();
          if (target.select && target.tagName === "INPUT") target.select();
        }
      });
    });
  }

  on(dialogClose, "click", () => closeDialog(null));

  on(dialogTitlebar, "pointerdown", (event) => {
    if (event.button !== 0 || event.target.closest("button")) return;
    dragState = {
      x: event.clientX,
      y: event.clientY,
      baseX: dialogOffset.x,
      baseY: dialogOffset.y,
    };
    dialogTitlebar.setPointerCapture(event.pointerId);
  });

  on(dialogTitlebar, "pointermove", (event) => {
    if (!dragState) return;
    dialogOffset = {
      x: dragState.baseX + event.clientX - dragState.x,
      y: dragState.baseY + event.clientY - dragState.y,
    };
    dialogBox.style.transform = `translate(${dialogOffset.x}px, ${dialogOffset.y}px)`;
  });

  const finishDrag = () => {
    dragState = null;
  };
  on(dialogTitlebar, "pointerup", finishDrag);
  on(dialogTitlebar, "pointercancel", finishDrag);

  const SHORTCUT_EDGE = 4;
  const SHORTCUT_THRESHOLD = 4;

  function placeShortcut(left, top) {
    if (!desktopShortcut || desktopReopen.hidden) return;
    const width = desktopShortcut.offsetWidth;
    const height = desktopShortcut.offsetHeight;
    const maxLeft = Math.max(SHORTCUT_EDGE, desktopReopen.clientWidth - width - SHORTCUT_EDGE);
    const maxTop = Math.max(SHORTCUT_EDGE, desktopReopen.clientHeight - height - SHORTCUT_EDGE);
    desktopShortcut.style.left = `${Math.min(Math.max(SHORTCUT_EDGE, left), maxLeft)}px`;
    desktopShortcut.style.top = `${Math.min(Math.max(SHORTCUT_EDGE, top), maxTop)}px`;
  }

  function clampShortcutToDesktop() {
    if (!desktopShortcut || desktopReopen.hidden) return;
    const left = parseFloat(desktopShortcut.style.left);
    const top = parseFloat(desktopShortcut.style.top);
    if (!Number.isFinite(left) || !Number.isFinite(top)) return;
    placeShortcut(left, top);
  }

  function finishShortcutDrag(event) {
    if (!shortcutDrag || event.pointerId !== shortcutDrag.pointerId) return;
    const moved = shortcutDrag.moved;
    shortcutDrag = null;
    desktopShortcut.classList.remove("is-dragging");
    if (desktopShortcut.hasPointerCapture(event.pointerId)) {
      desktopShortcut.releasePointerCapture(event.pointerId);
    }
    if (moved) {
      suppressShortcutClick = true;
      window.setTimeout(() => {
        suppressShortcutClick = false;
      }, 0);
    }
  }

  if (desktopShortcut) {
    on(desktopShortcut, "pointerdown", (event) => {
      if (event.button !== 0 || desktopReopen.hidden) return;
      const rect = desktopShortcut.getBoundingClientRect();
      const host = desktopReopen.getBoundingClientRect();
      shortcutDrag = {
        pointerId: event.pointerId,
        x: event.clientX,
        y: event.clientY,
        left: rect.left - host.left - desktopReopen.clientLeft,
        top: rect.top - host.top - desktopReopen.clientTop,
        moved: false,
      };
      desktopShortcut.setPointerCapture(event.pointerId);
    });

    on(desktopShortcut, "pointermove", (event) => {
      if (!shortcutDrag || event.pointerId !== shortcutDrag.pointerId) return;
      const dx = event.clientX - shortcutDrag.x;
      const dy = event.clientY - shortcutDrag.y;
      if (!shortcutDrag.moved && Math.hypot(dx, dy) < SHORTCUT_THRESHOLD) return;
      shortcutDrag.moved = true;
      desktopShortcut.classList.add("is-dragging");
      placeShortcut(shortcutDrag.left + dx, shortcutDrag.top + dy);
    });

    on(desktopShortcut, "pointerup", finishShortcutDrag);
    on(desktopShortcut, "pointercancel", finishShortcutDrag);

    on(window, "resize", clampShortcutToDesktop);
  }

  const WINDOW_EDGE = 4;
  const WINDOW_DRAG_THRESHOLD = 4;
  const WINDOW_MIN_VISIBLE = 64;

  function resetWindowOffset() {
    windowOffset = { x: 0, y: 0 };
    wordWindow.style.left = "";
    wordWindow.style.top = "";
  }

  function moveWindowTo(clientX, clientY) {
    if (!windowDrag) return;
    const host = root.getBoundingClientRect();
    const width = wordWindow.offsetWidth;
    const minLeft = host.left + WINDOW_EDGE;
    const maxLeft = Math.max(minLeft, host.right - width - WINDOW_EDGE);
    const minTop = host.top + WINDOW_EDGE;
    const maxTop = Math.max(minTop, host.bottom - WINDOW_MIN_VISIBLE);
    const left = Math.min(
      Math.max(clientX - windowDrag.grabX * width, minLeft),
      maxLeft
    );
    const top = Math.min(Math.max(clientY - windowDrag.grabY, minTop), maxTop);
    windowOffset = {
      x: left - windowDrag.baseLeft,
      y: top - windowDrag.baseTop,
    };
    wordWindow.style.left = `${windowOffset.x}px`;
    wordWindow.style.top = `${windowOffset.y}px`;
  }

  function finishWindowDrag(event) {
    if (!windowDrag || event.pointerId !== windowDrag.pointerId) return;
    const moved = windowDrag.moved;
    windowDrag = null;
    if (titlebar.hasPointerCapture(event.pointerId)) {
      titlebar.releasePointerCapture(event.pointerId);
    }
    if (moved) {
      suppressTitlebarDblClick = true;
      window.setTimeout(() => {
        suppressTitlebarDblClick = false;
      }, 250);
    }
  }

  function clampWindowToDesktop() {
    if (!root.classList.contains("is-restored")) return;
    if (root.classList.contains("is-minimized") || root.classList.contains("is-closed")) return;
    const host = root.getBoundingClientRect();
    const rect = wordWindow.getBoundingClientRect();
    if (!rect.width) return;
    const minLeft = host.left + WINDOW_EDGE;
    const maxLeft = Math.max(minLeft, host.right - rect.width - WINDOW_EDGE);
    const minTop = host.top + WINDOW_EDGE;
    const maxTop = Math.max(minTop, host.bottom - WINDOW_MIN_VISIBLE);
    const left = Math.min(Math.max(rect.left, minLeft), maxLeft);
    const top = Math.min(Math.max(rect.top, minTop), maxTop);
    if (left === rect.left && top === rect.top) return;
    windowOffset = {
      x: windowOffset.x + left - rect.left,
      y: windowOffset.y + top - rect.top,
    };
    wordWindow.style.left = `${windowOffset.x}px`;
    wordWindow.style.top = `${windowOffset.y}px`;
  }

  on(titlebar, "pointerdown", (event) => {
    if (event.button !== 0 || event.target.closest("button")) return;
    if (root.classList.contains("is-minimized") || root.classList.contains("is-closed")) return;
    const rect = wordWindow.getBoundingClientRect();
    windowDrag = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      grabX: (event.clientX - rect.left) / Math.max(rect.width, 1),
      grabY: Math.max(0, event.clientY - rect.top),
      baseLeft: rect.left - windowOffset.x,
      baseTop: rect.top - windowOffset.y,
      restored: root.classList.contains("is-restored"),
      moved: false,
    };
    titlebar.setPointerCapture(event.pointerId);
  });

  on(titlebar, "pointermove", (event) => {
    if (!windowDrag || event.pointerId !== windowDrag.pointerId) return;
    if (!windowDrag.moved) {
      const dx = event.clientX - windowDrag.startX;
      const dy = event.clientY - windowDrag.startY;
      if (Math.hypot(dx, dy) < WINDOW_DRAG_THRESHOLD) return;
      windowDrag.moved = true;
      if (!windowDrag.restored) {
        toggleRestore();
        const rect = wordWindow.getBoundingClientRect();
        windowDrag.baseLeft = rect.left;
        windowDrag.baseTop = rect.top;
        windowDrag.grabY = Math.min(windowDrag.grabY, rect.height);
      }
    }
    moveWindowTo(event.clientX, event.clientY);
  });

  on(titlebar, "pointerup", finishWindowDrag);
  on(titlebar, "pointercancel", finishWindowDrag);
  on(window, "resize", clampWindowToDesktop);

  on(dialogBox, "keydown", (event) => {
    if (event.key !== "Tab") return;
    const focusable = $$(
      'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), a[href]',
      dialogBox
    );
    if (!focusable.length) return;
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  });

  function scheduleStatus() {
    window.cancelAnimationFrame(statusFrame);
    statusFrame = window.requestAnimationFrame(refreshStatus);
  }

  function refreshStatus() {
    const sheets = $$(".sheet", paper);
    pageCount = Math.max(1, sheets.length);
    const anchor = editor.getBoundingClientRect().top + editor.clientHeight * 0.35;
    let page = 1;
    sheets.forEach((sheet, index) => {
      if (sheet.getBoundingClientRect().top <= anchor) page = index + 1;
    });
    if (editor.scrollTop <= 4) page = 1;
    currentPage = Math.min(pageCount, Math.max(1, page));
    statusPage.textContent = `Page ${currentPage} of ${pageCount}`;

    if (selectionInsidePaper()) {
      const sel = currentSelection();
      const range = sel.getRangeAt(0).cloneRange();
      let rect = range.getClientRects()[0];
      if (!rect) rect = range.getBoundingClientRect();
      const caret = caretElement();
      const sheet = caret ? caret.closest(".sheet") : null;
      if (rect && rect.width + rect.height > 0 && sheet) {
        const sheetRect = sheet.getBoundingClientRect();
        const style = window.getComputedStyle(sheet);
        const paddingTop = parseFloat(style.paddingTop) || 0;
        const lineHeight = parseFloat(style.lineHeight) || parseFloat(style.fontSize) * 1.45 || 24;
        const y = Math.max(0, rect.top - sheetRect.top - paddingTop);
        const line = Math.max(1, Math.round(y / lineHeight) + 1);
        const column = Math.max(1, sel.focusOffset + 1);
        const inches = Math.max(0, (rect.top - sheetRect.top) / 96);
        statusLine.textContent = `Ln ${line}`;
        statusCol.textContent = `Col ${column}`;
        statusPos.textContent = `At ${inches.toFixed(1)}"`;
      }
    }
  }

  function refreshToolbar() {
    const inside = selectionInsidePaper();
    $$("[data-state-command]").forEach((element) => {
      let active = false;
      if (inside) {
        try {
          active = document.queryCommandState(element.dataset.stateCommand);
        } catch (_) {
          active = false;
        }
      }
      element.classList.toggle("is-active", active);
      element.setAttribute("aria-pressed", String(active));
    });

    const styleCombo = $('[data-combo="style"]');
    const fontCombo = $('[data-combo="font"]');
    const sizeCombo = $('[data-combo="size"]');
    if (!styleCombo || !fontCombo || !sizeCombo) return;

    if (inside) {
      const element = caretElement();
      if (element) {
        const computed = window.getComputedStyle(element);
        const family = computed.fontFamily.split(",")[0].replace(/["']/g, "").trim();
        const points = Math.round((parseFloat(computed.fontSize) / 1.3333) * 10) / 10;
        setComboLabel(fontCombo, family || "Times New Roman");
        setComboLabel(sizeCombo, String(Number.isInteger(points) ? points : points.toFixed(1)));
        const block = element.closest("h1, h2, h3, p, li, td, th");
        let style = "Normal";
        if (block && block.tagName === "H1") style = "Title";
        else if (block && block.tagName === "H2") style = "Heading 1";
        else if (block && block.tagName === "H3") style = "Heading 2";
        setComboLabel(styleCombo, style);
      }
    }

    if (undoButton) undoButton.disabled = !inside || !document.queryCommandEnabled("undo");
    if (redoButton) redoButton.disabled = !inside || !document.queryCommandEnabled("redo");
  }

  function scheduleToolbar() {
    window.cancelAnimationFrame(selectionFrame);
    selectionFrame = window.requestAnimationFrame(refreshToolbar);
  }

  function refreshChecks() {
    const state = {
      docmap: !docmap.classList.contains("is-hidden"),
      ruler: !rulerShell.classList.contains("is-hidden"),
      "standard-toolbar": !standardToolbar.classList.contains("is-hidden"),
      "formatting-toolbar": !formatToolbar.classList.contains("is-hidden"),
      "window-document": !root.classList.contains("is-closed"),
    };
    $$("[data-check]").forEach((item) => {
      item.setAttribute("aria-checked", String(Boolean(state[item.dataset.check])));
    });
  }

  function setComboLabel(combo, label) {
    const element = $(".combo-label", combo);
    if (element) element.textContent = label;
  }

  function updateTableCommands() {
    const inTable = Boolean(currentTable());
    $$("[data-requires-table]").forEach((item) => {
      item.disabled = !inTable;
    });
  }

  const COMBO_OPTIONS = {
    style: [
      { value: "p", label: "Normal" },
      { value: "h2", label: "Heading 1" },
      { value: "h3", label: "Heading 2" },
      { value: "h1", label: "Title" },
    ],
    font: [
      { value: "Times New Roman", label: "Times New Roman" },
      { value: "Arial", label: "Arial" },
      { value: "Georgia", label: "Georgia" },
      { value: "Verdana", label: "Verdana" },
      { value: "Courier New", label: "Courier New" },
      { value: "Garamond", label: "Garamond" },
      { value: "Trebuchet MS", label: "Trebuchet MS" },
    ],
    size: [
      8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48,
    ].map((size) => ({ value: String(size), label: String(size) })),
    docmap: [
      { value: "headings", label: "Headings" },
      { value: "pages", label: "Pages" },
    ],
  };

  function closeCombo() {
    if (!comboState) return;
    comboState.combo.classList.remove("is-open");
    comboState.combo.setAttribute("aria-expanded", "false");
    comboPopup.hidden = true;
    comboPopup.innerHTML = "";
    comboState = null;
  }

  function openCombo(combo) {
    const key = combo.dataset.combo;
    const options = COMBO_OPTIONS[key];
    if (!options) return;
    closeCombo();
    rememberSelection();
    combo.classList.add("is-open");
    combo.setAttribute("aria-expanded", "true");
    comboPopup.innerHTML = "";
    const selected = $(".combo-label", combo)?.textContent || "";
    options.forEach((option) => {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "combo-option";
      item.dataset.value = option.value;
      item.textContent = option.label;
      item.setAttribute("role", "option");
      if (option.label === selected) {
        item.classList.add("is-selected");
        item.setAttribute("aria-selected", "true");
      }
      if (key === "font") item.style.fontFamily = option.value;
      if (key === "size") item.style.fontSize = "11px";
      item.addEventListener("click", () => {
        chooseComboValue(key, option.value, option.label);
      });
      comboPopup.appendChild(item);
    });

    comboPopup.hidden = false;
    const rect = combo.getBoundingClientRect();
    const popupRect = comboPopup.getBoundingClientRect();
    let left = rect.left;
    let top = rect.bottom + 1;
    if (left + popupRect.width > window.innerWidth - 4) {
      left = Math.max(4, window.innerWidth - popupRect.width - 4);
    }
    if (top + popupRect.height > window.innerHeight - 36) {
      top = Math.max(4, rect.top - popupRect.height - 1);
    }
    comboPopup.style.left = `${left}px`;
    comboPopup.style.top = `${top}px`;
    comboState = { combo };
  }

  function chooseComboValue(key, value, label) {
    closeCombo();
    if (key === "docmap") {
      setDocmapView(value);
      return;
    }
    if (key === "style") {
      applyBlockStyle(value);
      setComboLabel($('[data-combo="style"]'), label);
      return;
    }
    if (key === "font") {
      applyFontFamily(value);
      setComboLabel($('[data-combo="font"]'), label);
      return;
    }
    if (key === "size") {
      applyFontSize(value);
      setComboLabel($('[data-combo="size"]'), label);
    }
  }

  function setDocmapView(view) {
    docmapView = view;
    setComboLabel(docmapFilter, view === "pages" ? "Pages" : "Headings");
    $$("[data-docmap-view]", docmap).forEach((tree) => {
      tree.hidden = tree.dataset.docmapView !== view;
    });
  }

  function toggleDocmap() {
    docmap.classList.toggle("is-hidden");
    refreshChecks();
    showToast(docmap.classList.contains("is-hidden") ? "Document Map hidden." : "Document Map shown.");
  }

  function toggleRuler() {
    rulerShell.classList.toggle("is-hidden");
    refreshChecks();
  }

  function toggleToolbar(which) {
    const toolbar = which === "standard" ? standardToolbar : formatToolbar;
    toolbar.classList.toggle("is-hidden");
    refreshChecks();
  }

  function openMenu(menu, focusFirst = false) {
    closeMenu();
    menuName = menu.dataset.menuName;
    menu.classList.add("is-open");
    const title = $(".menu-title", menu);
    title.setAttribute("aria-expanded", "true");
    updateTableCommands();
    refreshChecks();
    if (focusFirst) {
      const first = $(".menu-item:not(:disabled)", menu);
      if (first) first.focus();
    }
  }

  function closeMenu() {
    if (!menuName) return;
    const menu = $(`.menu[data-menu-name="${menuName}"]`);
    if (menu) {
      menu.classList.remove("is-open");
      const title = $(".menu-title", menu);
      if (title) title.setAttribute("aria-expanded", "false");
    }
    menuName = null;
  }

  function openMenuByName(name, focusFirst = false) {
    const menu = $(`.menu[data-menu-name="${name}"]`);
    if (menu) openMenu(menu, focusFirst);
  }

  function toggleStart(force) {
    const shouldOpen = force === undefined ? startMenu.hidden : force;
    if (shouldOpen) {
      closeMenu();
      closeCombo();
      startMenu.hidden = false;
      startButton.classList.add("is-open");
      startButton.setAttribute("aria-expanded", "true");
      const first = $(".sm-item", startMenu);
      if (first) first.focus();
    } else {
      startMenu.hidden = true;
      startButton.classList.remove("is-open");
      startButton.setAttribute("aria-expanded", "false");
      startButton.focus();
    }
  }

  function closeStart() {
    if (startMenu.hidden) return;
    startMenu.hidden = true;
    startButton.classList.remove("is-open");
    startButton.setAttribute("aria-expanded", "false");
  }

  function minimizeWindow() {
    if (root.classList.contains("is-closed")) return;
    root.classList.add("is-minimized");
    taskButton.classList.remove("is-active");
    taskButton.setAttribute("aria-pressed", "false");
    closeStart();
    closeMenu();
    closeCombo();
  }

  function restoreWindow() {
    root.classList.remove("is-minimized");
    root.classList.remove("is-closed");
    taskButton.hidden = false;
    taskButton.classList.add("is-active");
    taskButton.setAttribute("aria-pressed", "true");
    desktopReopen.hidden = true;
    refreshChecks();
  }

  function reopenDocument() {
    restoreWindow();
    paper.focus({ preventScroll: true });
    showToast("Portfolio.doc is open.");
  }

  function toggleTask() {
    if (root.classList.contains("is-closed")) {
      reopenDocument();
      return;
    }
    if (root.classList.contains("is-minimized")) {
      restoreWindow();
      return;
    }
    minimizeWindow();
  }

  function showDesktop() {
    if (root.classList.contains("is-closed")) {
      showToast("The document is already closed. Click its desktop icon to reopen it.");
      return;
    }
    minimizeWindow();
  }

  function toggleRestore() {
    if (root.classList.contains("is-closed")) return;
    if (root.classList.contains("is-minimized")) {
      restoreWindow();
      return;
    }
    const restored = root.classList.toggle("is-restored");
    if (!restored) resetWindowOffset();
    maxButton.setAttribute("title", restored ? "Maximize" : "Restore Down");
    maxButton.setAttribute("aria-label", restored ? "Maximize" : "Restore Down");
  }

  function closeDocument() {
    root.classList.remove("is-minimized");
    root.classList.add("is-closed");
    desktopReopen.hidden = false;
    taskButton.hidden = true;
    closeStart();
    closeMenu();
    closeCombo();
    clampShortcutToDesktop();
    refreshChecks();
    showToast("Portfolio.doc closed. Click the desktop icon to reopen it.");
  }

  function scrollToHash(hash) {
    if (!hash || hash === "#") return;
    closeMenu();
    closeStart();
    closeCombo();
    const behavior = reduceMotion.matches ? "auto" : "smooth";
    if (hash === "#top") {
      editor.scrollTo({ top: 0, behavior });
      return;
    }
    const target = document.getElementById(hash.slice(1));
    if (!target) return;
    target.scrollIntoView({ behavior, block: "start" });
    target.classList.add("anchor-flash");
    window.setTimeout(() => target.classList.remove("anchor-flash"), 900);
  }

  const EXPORT_CSS = `
    body{font-family:"Times New Roman",Times,serif;font-size:12pt;line-height:1.45;color:#000;background:#fff;margin:0.55in}
    a{color:#0000cc;text-decoration:underline}
    h1{font-size:26pt;margin:0 0 4pt}
    h2{font-size:12pt;text-transform:uppercase;letter-spacing:.09em;border-bottom:1px solid #000;padding-bottom:2pt;margin:18pt 0 7pt}
    h3{font-size:15pt;margin:0 0 5pt}
    p{margin:0 0 7pt}
    table{width:100%;border-collapse:collapse;font-size:10pt;margin:4pt 0}
    th,td{border:1px solid #999;padding:4pt 6pt;vertical-align:top;text-align:left}
    th{background:#dcdcdc;font-size:9pt;text-transform:uppercase;letter-spacing:.06em}
    .sheet{margin:0;page-break-after:always;break-after:page}
    .sheet:last-child{page-break-after:auto;break-after:auto}
    .doc-head{text-align:center}
    .doc-eyebrow{font-size:9pt;letter-spacing:.18em;text-transform:uppercase;color:#555}
    .doc-role{font-size:11.5pt;font-style:italic}
    .doc-meta{font-size:9.5pt}
    .doc-rule{border:0;border-top:2px solid #000;border-bottom:1px solid #888;margin:14pt 0 0}
    .callout{border:1px solid #8c8a76;background:#fbfae6;padding:10pt 12pt;margin:4pt 0}
    .callout-label{font-size:8pt;font-weight:bold;letter-spacing:.16em;text-transform:uppercase;color:#6b5a00}
    .doc-note,.table-note{font-size:9.5pt;font-style:italic;color:#333}
    .sheet-footer{border-top:1px solid #bcbcbc;color:#666;font-size:8.5pt;padding-top:6pt;margin-top:18pt}
    .sr-only{display:none}
    .manual-page-break{page-break-after:always;break-after:page;border-top:1px dashed #999;color:#777;font-family:Verdana,sans-serif;font-size:8pt;padding:4pt 0;margin:12pt 0}
  `;

  function sanitizeFileName(name) {
    const clean = String(name || "Portfolio")
      .replace(/[\\/:*?"<>|]+/g, "-")
      .replace(/\s+/g, " ")
      .trim();
    return clean || "Portfolio";
  }

  function cleanExportMarkup(markup) {
    const holder = document.createElement("div");
    holder.innerHTML = markup.replace(/\u200b/g, "");
    holder.querySelectorAll("script, iframe, object, embed, link, meta, base, form").forEach((node) => node.remove());
    holder.querySelectorAll("*").forEach((element) => {
      Array.from(element.attributes).forEach((attribute) => {
        if (/^on/i.test(attribute.name)) element.removeAttribute(attribute.name);
      });
    });
    return holder.innerHTML;
  }

  function downloadDocument(name) {
    const requested = sanitizeFileName(name || currentFileName);
    const fileName = /\.(doc|docx|html?)$/i.test(requested) ? requested : `${requested}.doc`;
    const content = cleanExportMarkup(paper.innerHTML);
    const html = [
      "<!DOCTYPE html><html><head><meta charset=\"utf-8\">",
      `<title>${escapeHtml(fileName.replace(/\.(doc|docx|html?)$/i, ""))}</title>`,
      `<style>${EXPORT_CSS}</style></head><body>${content}</body></html>`,
    ].join("");
    const blob = new Blob(["\ufeff", html], { type: "application/msword;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    window.setTimeout(() => {
      URL.revokeObjectURL(url);
      link.remove();
    }, 1200);
    currentFileName = fileName;
    dirty = false;
    updateTitle();
    showToast(`Saved ${fileName}.`);
  }

  async function saveDocument() {
    downloadDocument(currentFileName);
  }

  async function saveAsDialog() {
    const suggested = currentFileName.replace(/\.docx?$/i, "") + ".doc";
    const value = await showDialog({
      title: "Save As",
      body: `
        <div class="dialog-row"><label for="save-name">File name:</label>
        <input class="dialog-input dialog-grow" id="save-name" type="text" value="${escapeAttr(suggested)}" spellcheck="false"></div>
        <div class="dialog-row"><label for="save-type">Save as type:</label>
        <select class="dialog-input dialog-grow" id="save-type"><option>Word Document (*.doc)</option></select></div>
        <p class="dialog-note">This demo saves a Word-compatible HTML document with a .doc extension.</p>`,
      buttons: [
        {
          label: "Save",
          primary: true,
          collect: (body) => {
            const input = body.querySelector("#save-name");
            const name = input.value.trim();
            if (!name) {
              input.focus();
              return false;
            }
            return name;
          },
        },
        { label: "Cancel", value: null },
      ],
      focus: "#save-name",
    });
    if (value) downloadDocument(value);
  }

  async function confirmDiscard(message) {
    if (!dirty) return true;
    const choice = await showDialog({
      title: "Microsoft Word",
      body: `<p>${escapeHtml(message || "Do you want to save the changes you made?")}</p>
        <p class="dialog-note">If you choose No, your changes since the last save will be lost.</p>`,
      buttons: [
        { label: "Yes", primary: true, value: "save" },
        { label: "No", value: "discard" },
        { label: "Cancel", value: "cancel" },
      ],
    });
    if (choice === "save") {
      downloadDocument(currentFileName);
      return true;
    }
    return choice === "discard";
  }

  async function newDocument() {
    if (!(await confirmDiscard("Do you want to save the changes to this document before creating a new one?"))) return;
    paper.innerHTML = `
      <article class="sheet" id="top">
        <header class="doc-head">
          <p class="doc-eyebrow">New Document</p>
          <h1 class="doc-name">Untitled</h1>
          <p class="doc-role">Start typing your document</p>
        </header>
        <hr class="doc-rule">
        <section class="doc-section">
          <p><br></p>
        </section>
        <footer class="sheet-footer"><span>Document1.doc</span><span>Page 1 of 1</span></footer>
      </article>`;
    currentFileName = "Document1.doc";
    dirty = false;
    updateTitle();
    editor.scrollTo({ top: 0, behavior: "auto" });
    const first = $(".doc-section p", paper);
    if (first) {
      paper.focus({ preventScroll: true });
      const range = document.createRange();
      range.setStart(first, 0);
      range.collapse(true);
      const sel = currentSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
    refreshStatus();
    refreshToolbar();
    showToast("Created Document1.doc. Use File > Revert to Portfolio to restore the portfolio.");
  }

  function resetToPortfolio(showMessage = true) {
    paper.innerHTML = originalHTML;
    currentFileName = "Portfolio.doc";
    dirty = false;
    updateTitle();
    editor.scrollTo({ top: 0, behavior: "auto" });
    refreshStatus();
    refreshToolbar();
    if (showMessage) showToast("Portfolio.doc restored.");
  }

  async function revertDocument() {
    if (!(await confirmDiscard("Revert to the original portfolio? Any unsaved changes will be lost."))) return;
    resetToPortfolio(true);
  }

  function loadPlainDocument(name, text) {
    const title = name.replace(/\.(txt|html?|doc)$/i, "") || "Untitled";
    const paragraphs = text
      .replace(/\r\n?/g, "\n")
      .split(/\n{2,}/)
      .map((part) => part.trim())
      .filter(Boolean)
      .slice(0, 160);
    paper.innerHTML = "";
    const article = document.createElement("article");
    article.className = "sheet";
    article.id = "top";
    const header = document.createElement("header");
    header.className = "doc-head";
    const eyebrow = document.createElement("p");
    eyebrow.className = "doc-eyebrow";
    eyebrow.textContent = "Opened Document";
    const heading = document.createElement("h1");
    heading.className = "doc-name";
    heading.textContent = title;
    const meta = document.createElement("p");
    meta.className = "doc-meta";
    meta.textContent = `Opened from ${name}`;
    header.append(eyebrow, heading, meta);
    const rule = document.createElement("hr");
    rule.className = "doc-rule";
    const section = document.createElement("section");
    section.className = "doc-section";
    (paragraphs.length ? paragraphs : ["(This document is empty.)"]).forEach((part) => {
      const paragraph = document.createElement("p");
      paragraph.textContent = part;
      section.appendChild(paragraph);
    });
    const footer = document.createElement("footer");
    footer.className = "sheet-footer";
    const footerA = document.createElement("span");
    footerA.textContent = name;
    const footerB = document.createElement("span");
    footerB.textContent = "Page 1 of 1";
    footer.append(footerA, footerB);
    article.append(header, rule, section, footer);
    paper.appendChild(article);
    currentFileName = name;
    dirty = false;
    updateTitle();
    editor.scrollTo({ top: 0, behavior: "auto" });
    refreshStatus();
    refreshToolbar();
    showToast(`Opened ${name}. File > Revert to Portfolio restores the portfolio.`);
  }

  async function loadLocalFile(file) {
    try {
      const buffer = await file.arrayBuffer();
      const sample = new Uint8Array(buffer.slice(0, 4000));
      if (sample.includes(0)) throw new Error("This looks like a binary Word document. Save it as text or HTML first.");
      const text = new TextDecoder("utf-8").decode(buffer);
      let plain = text;
      if (/\.html?$/i.test(file.name) || /<html|<body|<p[\s>]/i.test(text.slice(0, 800))) {
        const parsed = new DOMParser().parseFromString(text, "text/html");
        parsed.querySelectorAll("script, style, iframe, object, embed").forEach((node) => node.remove());
        const blocks = Array.from(parsed.body.querySelectorAll("p, h1, h2, h3, h4, li, blockquote, pre"))
          .map((node) => node.textContent.trim())
          .filter(Boolean);
        plain = blocks.length ? blocks.join("\n\n") : parsed.body.textContent || "";
      }
      if (!plain.trim()) plain = "(This document is empty.)";
      loadPlainDocument(file.name, plain);
    } catch (error) {
      console.error(error);
      await showDialog({
        title: "Microsoft Word",
        body: `<p>Word cannot open <strong>${escapeHtml(file.name)}</strong>.</p>
          <p class="dialog-note">${escapeHtml(error.message || "Use a plain text (.txt) or HTML file in this browser demo.")}</p>`,
        buttons: [{ label: "OK", primary: true, value: null }],
      });
    }
  }

  async function openDocumentDialog() {
    const choice = await showDialog({
      title: "Open",
      width: "440px",
      body: `
        <div class="dialog-row"><label>Look in:</label><span class="dialog-folder">My Documents</span></div>
        <div class="dialog-list" role="listbox" aria-label="Documents">
          <div class="dialog-list-item is-selected">
            <span class="sm-icon ql-doc" aria-hidden="true"></span>
            <span><strong>Portfolio.doc</strong><br><small>Microsoft Word document · 2 pages</small></span>
          </div>
        </div>
        <p class="dialog-note">You can also open a plain text or HTML file as a simplified document. Binary .doc files are not parsed in this browser demo.</p>`,
      buttons: [
        { label: "Open", primary: true, value: "portfolio" },
        { label: "Browse…", value: "browse" },
        { label: "Cancel", value: null },
      ],
    });
    if (choice === "browse") {
      fileInput.value = "";
      fileInput.click();
      return;
    }
    if (choice === "portfolio") {
      if (!(await confirmDiscard("Opening Portfolio.doc will discard unsaved changes. Continue?"))) return;
      resetToPortfolio(true);
    }
  }

  on(fileInput, "change", () => {
    const file = fileInput.files && fileInput.files[0];
    if (file) loadLocalFile(file);
  });

  async function insertHyperlinkDialog() {
    rememberSelection();
    const selected = currentSelection().toString().trim();
    const result = await showDialog({
      title: "Insert Hyperlink",
      width: "440px",
      body: `
        <p class="dialog-note">Link to an existing page, a section of this document, or an external site.</p>
        <div class="dialog-row"><label for="link-text">Text to display:</label>
        <input class="dialog-input dialog-grow" id="link-text" type="text" value="${escapeAttr(selected)}"></div>
        <div class="dialog-row"><label for="link-url">Address:</label>
        <input class="dialog-input dialog-grow" id="link-url" type="text" value="https://" spellcheck="false"></div>
        <p class="dialog-help">Use #profile, #work, or #contact to link to a section of this document.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => {
            let url = body.querySelector("#link-url").value.trim();
            let text = body.querySelector("#link-text").value.trim();
            if (!url) {
              body.querySelector("#link-url").focus();
              return false;
            }
            if (!/^(https?:|mailto:|#)/i.test(url)) url = `https://${url}`;
            if (!text) text = url;
            return { url, text };
          },
        },
        { label: "Cancel", value: null },
      ],
      focus: "#link-text",
    });
    if (!result) return;
    restoreSelection();
    insertHtml(
      `<a href="${escapeAttr(result.url)}" target="_blank" rel="noreferrer">${escapeHtml(result.text)}</a>`
    );
  }

  async function insertTableDialog() {
    rememberSelection();
    const result = await showDialog({
      title: "Insert Table",
      body: `
        <div class="dialog-row"><label for="table-rows">Number of rows:</label>
        <input class="dialog-input dialog-number" id="table-rows" type="number" min="1" max="20" value="3"></div>
        <div class="dialog-row"><label for="table-cols">Number of columns:</label>
        <input class="dialog-input dialog-number" id="table-cols" type="number" min="1" max="8" value="3"></div>
        <p class="dialog-note">Word will insert a simple bordered table at the cursor.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => {
            const rows = Math.min(20, Math.max(1, parseInt(body.querySelector("#table-rows").value, 10) || 3));
            const cols = Math.min(8, Math.max(1, parseInt(body.querySelector("#table-cols").value, 10) || 3));
            return { rows, cols };
          },
        },
        { label: "Cancel", value: null },
      ],
      focus: "#table-rows",
    });
    if (!result) return;
    restoreSelection();
    let markup = '<table class="doc-table"><tbody>';
    for (let row = 0; row < result.rows; row += 1) {
      markup += "<tr>";
      for (let col = 0; col < result.cols; col += 1) markup += "<td><br></td>";
      markup += "</tr>";
    }
    markup += "</tbody></table><p><br></p>";
    insertHtml(markup);
  }

  function insertColumn(side) {
    const cell = currentCell();
    const table = currentTable();
    if (!cell || !table) {
      showToast("Place the cursor inside a table first.");
      return;
    }
    const index = Array.from(cell.parentElement.children).indexOf(cell);
    $$("tr", table).forEach((row) => {
      const reference = row.children[index];
      if (!reference) return;
      const created = document.createElement(reference.tagName === "TH" ? "th" : "td");
      created.innerHTML = reference.tagName === "TH" ? "Column" : "<br>";
      if (side === "left") row.insertBefore(created, reference);
      else row.insertBefore(created, reference.nextSibling);
    });
    afterEdit();
    updateTableCommands();
  }

  function insertRow(position) {
    const cell = currentCell();
    const table = currentTable();
    if (!cell || !table) {
      showToast("Place the cursor inside a table first.");
      return;
    }
    const row = cell.parentElement;
    const clone = row.cloneNode(true);
    Array.from(clone.children).forEach((child) => {
      child.innerHTML = "<br>";
    });
    if (position === "above") row.parentElement.insertBefore(clone, row);
    else row.parentElement.insertBefore(clone, row.nextSibling);
    afterEdit();
    updateTableCommands();
  }

  function deleteRow() {
    const cell = currentCell();
    const table = currentTable();
    if (!cell || !table) {
      showToast("Place the cursor inside a table first.");
      return;
    }
    if (table.rows.length <= 1) {
      showToast("Word cannot delete the only row in the table.");
      return;
    }
    cell.parentElement.remove();
    afterEdit();
  }

  function deleteColumn() {
    const cell = currentCell();
    const table = currentTable();
    if (!cell || !table) {
      showToast("Place the cursor inside a table first.");
      return;
    }
    if (cell.parentElement.children.length <= 1) {
      showToast("Word cannot delete the only column in the table.");
      return;
    }
    const index = Array.from(cell.parentElement.children).indexOf(cell);
    $$("tr", table).forEach((row) => {
      if (row.children[index]) row.children[index].remove();
    });
    afterEdit();
    updateTableCommands();
  }

  function selectCurrentTable() {
    const table = currentTable();
    if (!table) {
      showToast("Place the cursor inside a table first.");
      return;
    }
    const range = document.createRange();
    range.selectNode(table);
    const sel = currentSelection();
    paper.focus({ preventScroll: true });
    sel.removeAllRanges();
    sel.addRange(range);
    table.scrollIntoView({ behavior: reduceMotion.matches ? "auto" : "smooth", block: "center" });
  }

  function insertPageBreak() {
    insertHtml('<div class="manual-page-break" contenteditable="false"><span>Page Break</span></div><p><br></p>');
  }

  function insertRule() {
    insertHtml('<hr class="doc-rule manual-rule"><p><br></p>');
  }

  async function insertDateDialog() {
    rememberSelection();
    const now = new Date();
    const longDate = now.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });
    const shortDate = now.toLocaleDateString();
    const time = now.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
    const value = await showDialog({
      title: "Date and Time",
      body: `
        <div class="dialog-row"><label for="date-format">Available formats:</label>
        <select class="dialog-input dialog-grow" id="date-format" size="5">
          <option>${escapeHtml(longDate)}</option>
          <option>${escapeHtml(time)}</option>
          <option>${escapeHtml(`${longDate} ${time}`)}</option>
          <option>${escapeHtml(shortDate)}</option>
        </select></div>
        <p class="dialog-note">The date and time will be inserted at the cursor as text.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => body.querySelector("#date-format").value,
        },
        { label: "Cancel", value: null },
      ],
      focus: "#date-format",
    });
    if (!value) return;
    restoreSelection();
    executeCommand("insertText", value);
  }

  function applyBorders(kind) {
    const table = currentTable();
    if (table) {
      table.classList.remove("doc-table-borderless", "doc-table-bordered");
      if (kind === "none") table.classList.add("doc-table-borderless");
      if (kind === "box") table.classList.add("doc-table-bordered");
      if (kind === "grid") table.classList.remove("doc-table-borderless", "doc-table-bordered");
      afterEdit();
      return;
    }
    const cell = currentCell();
    const element = caretElement();
    const block = element ? element.closest("p, li, h1, h2, h3") : null;
    if (!block) {
      showToast("Place the cursor in a paragraph or table first.");
      return;
    }
    block.classList.remove("doc-block-borderless", "doc-block-bordered");
    if (kind === "none") block.classList.add("doc-block-borderless");
    else block.classList.add("doc-block-bordered");
    afterEdit();
  }

  async function bordersDialog() {
    rememberSelection();
    const choice = await showDialog({
      title: "Borders and Shading",
      body: `
        <fieldset class="dialog-fieldset"><legend>Setting</legend>
          <label class="dialog-check"><input type="radio" name="border-kind" value="none"> None</label>
          <label class="dialog-check"><input type="radio" name="border-kind" value="box" checked> Box</label>
          <label class="dialog-check"><input type="radio" name="border-kind" value="grid"> Grid</label>
        </fieldset>
        <p class="dialog-note">Applies to the current table if the cursor is inside one, otherwise to the current paragraph.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => body.querySelector('input[name="border-kind"]:checked').value,
        },
        { label: "Cancel", value: null },
      ],
    });
    if (!choice) return;
    restoreSelection();
    applyBorders(choice);
  }

  function toggleHighlight() {
    if (!ensureDocumentSelection()) return;
    let current = "";
    try {
      current = document.queryCommandValue("hiliteColor");
    } catch (_) {
      current = "";
    }
    const active = current && !/transparent|rgba\(0, 0, 0, 0\)/i.test(current);
    executeCommand("hiliteColor", active ? "transparent" : "#ffff00");
  }

  async function fontColorDialog() {
    rememberSelection();
    const colors = [
      "#000000", "#404040", "#7f7f7f", "#bfbfbf", "#ffffff",
      "#7f0000", "#ff0000", "#ff7f7f", "#ff7f00", "#ffbf00",
      "#ffff00", "#bfff00", "#00a000", "#00ff00", "#00ffbf",
      "#00bfff", "#007fff", "#0000ff", "#7f00ff", "#ff00ff",
      "#7f007f", "#a0522d", "#c0c0c0", "#ffffe0", "#f5deb3",
      "#e6e6fa", "#dcdcdc", "#2f4f4f",
    ];
    const chosen = await showDialog({
      title: "Font Color",
      body: `
        <p class="dialog-note">Choose a color for the selected text.</p>
        <div class="dialog-color-grid">
          ${colors.map((color) => `<button type="button" class="dialog-color" data-color="${color}" style="background:${color}" title="${color}" aria-label="${color}"></button>`).join("")}
        </div>
        <p class="dialog-help">Automatic color is black.</p>`,
      buttons: [{ label: "Cancel", value: null }],
      onMount: (body, finish) => {
        $$(".dialog-color", body).forEach((swatch) => {
          swatch.addEventListener("click", () => finish(swatch.dataset.color));
        });
      },
    });
    if (!chosen) return;
    restoreSelection();
    executeCommand("foreColor", chosen);
  }

  async function fontDialog() {
    rememberSelection();
    if (!selectionInsidePaper() && !savedRange) {
      showToast("Select the text you want to format first.");
      return;
    }
    const element = caretElement() || paper;
    const computed = window.getComputedStyle(element);
    const currentFont = computed.fontFamily.split(",")[0].replace(/["']/g, "").trim();
    const currentSize = Math.round((parseFloat(computed.fontSize) / 1.3333) * 10) / 10;
    const result = await showDialog({
      title: "Font",
      width: "420px",
      body: `
        <div class="dialog-row"><label for="font-family">Font:</label>
        <select class="dialog-input dialog-grow" id="font-family">
          ${["Times New Roman", "Arial", "Georgia", "Verdana", "Courier New", "Garamond", "Trebuchet MS"].map((font) => `<option${font === currentFont ? " selected" : ""}>${font}</option>`).join("")}
        </select></div>
        <div class="dialog-row"><label for="font-size">Size:</label>
        <input class="dialog-input dialog-number" id="font-size" type="number" min="8" max="72" step="1" value="${Number.isFinite(currentSize) ? currentSize : 12}"></div>
        <div class="dialog-row dialog-row-top"><span class="dialog-label">Style:</span>
          <div class="dialog-checks">
            <label class="dialog-check"><input type="checkbox" id="font-bold"> Bold</label>
            <label class="dialog-check"><input type="checkbox" id="font-italic"> Italic</label>
            <label class="dialog-check"><input type="checkbox" id="font-underline"> Underline</label>
          </div>
        </div>
        <div class="font-preview" id="font-preview">The quick brown fox jumps over the lazy dog.</div>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => ({
            font: body.querySelector("#font-family").value,
            size: Math.min(72, Math.max(8, parseInt(body.querySelector("#font-size").value, 10) || 12)),
            bold: body.querySelector("#font-bold").checked,
            italic: body.querySelector("#font-italic").checked,
            underline: body.querySelector("#font-underline").checked,
          }),
        },
        { label: "Cancel", value: null },
      ],
      onMount: (body) => {
        const preview = body.querySelector("#font-preview");
        const update = () => {
          preview.style.fontFamily = body.querySelector("#font-family").value;
          preview.style.fontSize = `${body.querySelector("#font-size").value || 12}px`;
          preview.style.fontWeight = body.querySelector("#font-bold").checked ? "700" : "400";
          preview.style.fontStyle = body.querySelector("#font-italic").checked ? "italic" : "normal";
          preview.style.textDecoration = body.querySelector("#font-underline").checked ? "underline" : "none";
        };
        $$("select, input", body).forEach((input) => input.addEventListener("input", update));
        update();
      },
    });
    if (!result) return;
    restoreSelection();
    applyFontFamily(result.font);
    applyFontSize(result.size);
    setCommandState("bold", result.bold);
    setCommandState("italic", result.italic);
    setCommandState("underline", result.underline);
    afterEdit();
  }

  async function styleDialog() {
    rememberSelection();
    const styles = [
      { tag: "p", name: "Normal", sample: "Body text" },
      { tag: "h2", name: "Heading 1", sample: "SECTION HEADING" },
      { tag: "h3", name: "Heading 2", sample: "Subheading" },
      { tag: "h1", name: "Title", sample: "Document Title" },
    ];
    const chosen = await showDialog({
      title: "Style",
      body: `
        <p class="dialog-note">Pick a paragraph style for the current paragraph.</p>
        <div class="dialog-style-list">
          ${styles.map((style) => `<button type="button" class="dialog-style" data-tag="${style.tag}"><span class="style-sample style-${style.tag}">${style.sample}</span><span>${style.name}</span></button>`).join("")}
        </div>`,
      buttons: [{ label: "Cancel", value: null }],
      onMount: (body, finish) => {
        $$(".dialog-style", body).forEach((button) => {
          button.addEventListener("click", () => finish(button.dataset.tag));
        });
      },
    });
    if (!chosen) return;
    restoreSelection();
    applyBlockStyle(chosen);
  }

  function findInDocument(query) {
    if (!query) return false;
    const walker = document.createTreeWalker(paper, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        if (!node.data.trim()) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      },
    });
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    if (!nodes.length) return false;
    const selection = currentSelection();
    let startNode = null;
    let startOffset = 0;
    if (selection.rangeCount && selectionInsidePaper(selection)) {
      startNode = selection.focusNode;
      startOffset = selection.focusOffset;
    }
    const needle = query.toLowerCase();
    const startIndex = Math.max(0, nodes.indexOf(startNode));
    for (let offset = 0; offset < nodes.length; offset += 1) {
      const node = nodes[(startIndex + offset) % nodes.length];
      const haystack = node.data.toLowerCase();
      let found = -1;
      if (offset === 0 && node === startNode) found = haystack.indexOf(needle, startOffset);
      else found = haystack.indexOf(needle);
      if (found >= 0) {
        const range = document.createRange();
        range.setStart(node, found);
        range.setEnd(node, found + query.length);
        paper.focus({ preventScroll: true });
        selection.removeAllRanges();
        selection.addRange(range);
        const parent = node.parentElement;
        if (parent) {
          parent.scrollIntoView({
            behavior: reduceMotion.matches ? "auto" : "smooth",
            block: "center",
          });
          parent.classList.add("find-hit");
          window.setTimeout(() => parent.classList.remove("find-hit"), 900);
        }
        refreshStatus();
        return true;
      }
    }
    return false;
  }

  async function findDialog() {
    const selected = currentSelection() && selectionInsidePaper()
      ? currentSelection().toString().slice(0, 80)
      : "";
    const query = await showDialog({
      title: "Find and Replace",
      body: `
        <div class="dialog-row"><label for="find-text">Find what:</label>
        <input class="dialog-input dialog-grow" id="find-text" type="text" value="${escapeAttr(selected)}"></div>
        <p class="dialog-note">Searches the document text. Press Enter to find the next match.</p>`,
      buttons: [
        {
          label: "Find Next",
          primary: true,
          collect: (body) => body.querySelector("#find-text").value.trim(),
        },
        { label: "Cancel", value: null },
      ],
      focus: "#find-text",
      onMount: (body) => {
        body.querySelector("#find-text").addEventListener("keydown", (event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            const button = dialogButtons.querySelector(".dialog-button.primary");
            if (button) button.click();
          }
        });
      },
    });
    if (!query) return;
    if (!findInDocument(query)) {
      showToast(`Word has finished searching the document for "${query}".`);
    }
  }

  function wordCounts() {
    const text = paper.innerText || paper.textContent || "";
    const words = text.trim().split(/\s+/).filter(Boolean).length;
    const characters = text.replace(/\s/g, "").length;
    const paragraphs = $$("p, li, td, th", paper).length;
    return { words, characters, paragraphs, pages: $$(".sheet", paper).length };
  }

  async function wordCountDialog() {
    const counts = wordCounts();
    await showDialog({
      title: "Word Count",
      body: `
        <table class="dialog-table">
          <tbody>
            <tr><th>Pages</th><td>${counts.pages}</td></tr>
            <tr><th>Words</th><td>${counts.words}</td></tr>
            <tr><th>Characters (no spaces)</th><td>${counts.characters}</td></tr>
            <tr><th>Paragraphs</th><td>${counts.paragraphs}</td></tr>
          </tbody>
        </table>
        <label class="dialog-check"><input type="checkbox" checked> Include textboxes, footnotes and endnotes</label>`,
      buttons: [{ label: "Close", primary: true, value: null }],
    });
  }

  async function propertiesDialog() {
    const counts = wordCounts();
    const size = Math.max(1, Math.round(new Blob([paper.innerHTML]).size / 1024));
    const created = new Date().toLocaleDateString(undefined, {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    await showDialog({
      title: "Portfolio.doc Properties",
      width: "440px",
      body: `
        <div class="dialog-tabs"><span class="is-active">General</span><span>Summary</span><span>Statistics</span></div>
        <table class="dialog-table dialog-table-props">
          <tbody>
            <tr><th>File name</th><td>${escapeHtml(currentFileName)}</td></tr>
            <tr><th>Type</th><td>Microsoft Word Document</td></tr>
            <tr><th>Location</th><td>My Documents</td></tr>
            <tr><th>Size</th><td>${size} KB</td></tr>
            <tr><th>Created</th><td>${escapeHtml(created)}</td></tr>
            <tr><th>Modified</th><td>${escapeHtml(new Date().toLocaleString())}</td></tr>
            <tr><th>Author</th><td>Gabriel Elorde</td></tr>
            <tr><th>Pages</th><td>${counts.pages}</td></tr>
            <tr><th>Words</th><td>${counts.words}</td></tr>
          </tbody>
        </table>`,
      buttons: [{ label: "OK", primary: true, value: null }],
    });
  }

  async function optionsDialog() {
    const wasDocmap = !docmap.classList.contains("is-hidden");
    const wasRuler = !rulerShell.classList.contains("is-hidden");
    const wasClippy = clippy.style.display !== "none";
    const result = await showDialog({
      title: "Options",
      width: "440px",
      body: `
        <fieldset class="dialog-fieldset"><legend>View</legend>
          <label class="dialog-check"><input type="checkbox" id="opt-docmap"${wasDocmap ? " checked" : ""}> Show the Document Map</label>
          <label class="dialog-check"><input type="checkbox" id="opt-ruler"${wasRuler ? " checked" : ""}> Show the ruler</label>
          <label class="dialog-check"><input type="checkbox" id="opt-clippy"${wasClippy ? " checked" : ""}> Show the paperclip assistant</label>
        </fieldset>
        <fieldset class="dialog-fieldset"><legend>Spelling</legend>
          <label class="dialog-check"><input type="checkbox" id="opt-spell"${paper.spellcheck ? " checked" : ""}> Check spelling as you type</label>
        </fieldset>
        <p class="dialog-note">These options apply to this browser session only.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => ({
            docmap: body.querySelector("#opt-docmap").checked,
            ruler: body.querySelector("#opt-ruler").checked,
            clippy: body.querySelector("#opt-clippy").checked,
            spell: body.querySelector("#opt-spell").checked,
          }),
        },
        { label: "Cancel", value: null },
      ],
    });
    if (!result) return;
    docmap.classList.toggle("is-hidden", !result.docmap);
    rulerShell.classList.toggle("is-hidden", !result.ruler);
    clippy.style.display = result.clippy ? "" : "none";
    paper.spellcheck = result.spell;
    refreshChecks();
    showToast("Options saved for this session.");
  }

  function setZoom(percent) {
    zoomPercent = Math.min(200, Math.max(50, Math.round(percent)));
    paper.style.zoom = String(zoomPercent / 100);
    editor.classList.toggle("is-zoomed", zoomPercent > 100);
    statusZoom.textContent = `${zoomPercent}%`;
    refreshStatus();
  }

  async function zoomDialog() {
    const value = await showDialog({
      title: "Zoom",
      body: `
        <div class="dialog-row"><label for="zoom-percent">Zoom to:</label>
        <select class="dialog-input dialog-number" id="zoom-percent">
          ${[50, 75, 100, 125, 150, 200].map((value) => `<option value="${value}"${value === zoomPercent ? " selected" : ""}>${value}%</option>`).join("")}
        </select></div>
        <p class="dialog-note">Use 100% for the intended Word page size.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => parseInt(body.querySelector("#zoom-percent").value, 10),
        },
        { label: "Cancel", value: null },
      ],
      focus: "#zoom-percent",
    });
    if (value) setZoom(value);
  }

  async function fullScreen() {
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else await root.requestFullscreen();
    } catch (error) {
      console.error(error);
      showToast("Your browser did not allow full screen. Press F11 instead.");
    }
  }

  async function languageDialog() {
    const value = await showDialog({
      title: "Language",
      body: `
        <div class="dialog-row"><label for="language">Mark selected text as:</label>
        <select class="dialog-input dialog-grow" id="language">
          <option value="en-US">English (U.S.)</option>
          <option value="en-GB">English (U.K.)</option>
          <option value="fil-PH">Filipino (Philippines)</option>
          <option value="es-ES">Spanish (Spain)</option>
        </select></div>
        <p class="dialog-note">The language controls the browser spell checker.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => body.querySelector("#language").value,
        },
        { label: "Cancel", value: null },
      ],
      focus: "#language",
    });
    if (!value) return;
    const names = {
      "en-US": "English (U.S.)",
      "en-GB": "English (U.K.)",
      "fil-PH": "Filipino (Philippines)",
      "es-ES": "Spanish (Spain)",
    };
    statusLang.textContent = names[value] || value;
    paper.setAttribute("lang", value);
    showToast(`Language set to ${names[value] || value}.`);
  }

  async function goToPageDialog() {
    const sheets = $$(".sheet", paper);
    if (sheets.length <= 1) {
      showToast("This document has only one page.");
      return;
    }
    const value = await showDialog({
      title: "Go To Page",
      body: `
        <div class="dialog-row"><label for="page-number">Page number:</label>
        <input class="dialog-input dialog-number" id="page-number" type="number" min="1" max="${sheets.length}" value="${currentPage}"></div>
        <p class="dialog-note">Pages ${sheets.length} · Enter a number between 1 and ${sheets.length}.</p>`,
      buttons: [
        {
          label: "Go To",
          primary: true,
          collect: (body) => parseInt(body.querySelector("#page-number").value, 10),
        },
        { label: "Cancel", value: null },
      ],
      focus: "#page-number",
    });
    if (!value) return;
    const index = Math.min(sheets.length, Math.max(1, value)) - 1;
    sheets[index].scrollIntoView({
      behavior: reduceMotion.matches ? "auto" : "smooth",
      block: "start",
    });
    showToast(`Moved to page ${index + 1} of ${sheets.length}.`);
  }

  async function spellingDialog() {
    const result = await showDialog({
      title: "Spelling and Grammar",
      width: "440px",
      body: `
        <p><strong>Spelling check is complete.</strong></p>
        <p>Your browser checks spelling as you type. Words it does not recognize are underlined with a red wavy line.</p>
        <label class="dialog-check"><input type="checkbox" id="spell-live"${paper.spellcheck ? " checked" : ""}> Check spelling as you type</label>
        <p class="dialog-note">Try typing a misspelled word in the document to see the browser underline it.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => ({ live: body.querySelector("#spell-live").checked }),
        },
        { label: "Cancel", value: null },
      ],
    });
    if (result === null) return;
    paper.spellcheck = Boolean(result.live);
    showToast(result.live ? "Spelling check is on." : "Spelling check is off.");
  }

  async function officeClipboard() {
    try {
      const text = await navigator.clipboard.readText();
      await showDialog({
        title: "Clipboard",
        body: `<p class="dialog-note">The most recent clipboard item:</p><pre class="dialog-clipboard">${escapeHtml(text.slice(0, 3000)) || "(empty)"}</pre>`,
        buttons: [
          {
            label: "Paste",
            primary: true,
            value: "paste",
          },
          { label: "Close", value: null },
        ],
      }).then((choice) => {
        if (choice === "paste") {
          if (ensureDocumentSelection()) executeCommand("insertText", text);
        }
      });
    } catch (error) {
      console.error(error);
      await showDialog({
        title: "Clipboard",
        body: "<p>Your browser blocked clipboard access.</p><p class=\"dialog-note\">Use Ctrl+C and Ctrl+V in the document instead.</p>",
        buttons: [{ label: "OK", primary: true, value: null }],
      });
    }
  }

  async function helpTopicsDialog() {
    const topics = [
      ["Formatting text", "Select text, then use the B, I, and U buttons or choose Format > Font."],
      ["Saving and printing", "File > Save downloads a Word-compatible .doc file. File > Print prints only the document pages."],
      ["Document Map", "Use View > Document Map, or switch between Headings and Pages in the map pane."],
      ["Editing", "Click in the document to place the cursor. Typing changes the document; use Ctrl+Z to undo."],
      ["Navigation", "Use the document map, the phone section bar, or any underlined link."],
      ["Speech and key shortcuts", "Press Alt+F for the File menu, Ctrl+S to save, Ctrl+P to print, and F7 for spelling."],
      ["Restoring the portfolio", "If you create a new document or open a file, choose File > Revert to Portfolio."],
    ];
    await showDialog({
      title: "Portfolio Help",
      width: "500px",
      body: `
        <div class="dialog-row"><label for="help-search">Search:</label>
        <input class="dialog-input dialog-grow" id="help-search" type="search" placeholder="Type a topic"></div>
        <div class="help-topics" id="help-topics">
          ${topics.map(([title, text]) => `<section class="help-topic"><h3>${title}</h3><p>${text}</p></section>`).join("")}
        </div>
        <p class="dialog-note">This is a portfolio built as a Word 2003 homage. It is not affiliated with Microsoft.</p>`,
      buttons: [{ label: "Close", primary: true, value: null }],
      focus: "#help-search",
      onMount: (body) => {
        const search = body.querySelector("#help-search");
        search.addEventListener("input", () => {
          const query = search.value.toLowerCase();
          $$(".help-topic", body).forEach((topic) => {
            topic.hidden = Boolean(query) && !topic.textContent.toLowerCase().includes(query);
          });
        });
      },
    });
  }

  async function shortcutsDialog() {
    await showDialog({
      title: "Keyboard Shortcuts",
      width: "460px",
      body: `
        <table class="dialog-table dialog-shortcuts">
          <tbody>
            <tr><th>Ctrl+N</th><td>New document</td></tr>
            <tr><th>Ctrl+O</th><td>Open</td></tr>
            <tr><th>Ctrl+S</th><td>Save a .doc copy</td></tr>
            <tr><th>Ctrl+P</th><td>Print the document</td></tr>
            <tr><th>Ctrl+F</th><td>Find</td></tr>
            <tr><th>Ctrl+K</th><td>Insert hyperlink</td></tr>
            <tr><th>Ctrl+B / I / U</th><td>Bold, italic, underline</td></tr>
            <tr><th>Ctrl+Z / Y</th><td>Undo, redo</td></tr>
            <tr><th>Ctrl+Enter</th><td>Insert page break</td></tr>
            <tr><th>F7</th><td>Spelling and grammar</td></tr>
            <tr><th>Alt+F</th><td>Open the File menu</td></tr>
            <tr><th>Esc</th><td>Close menus and dialogs</td></tr>
          </tbody>
        </table>`,
      buttons: [{ label: "Close", primary: true, value: null }],
    });
  }

  async function aboutDialog(wordOnly = false) {
    const counts = wordCounts();
    await showDialog({
      title: wordOnly ? "About Microsoft Word" : "About Portfolio.doc",
      width: "470px",
      body: `
        <div class="about-banner">
          <span class="word-icon about-word-icon" aria-hidden="true"></span>
          <span class="about-title">Portfolio.doc <small>2003</small></span>
        </div>
        <p class="about-copy">A hand-built HTML and CSS homage to Microsoft Word 2003 running on a Windows XP desktop.</p>
        <p class="about-copy">This document contains the portfolio of <strong>Gabriel Elorde</strong>: ${counts.pages} pages, ${counts.words} words, and ${counts.paragraphs} paragraphs of profile, projects, education, toolkit, and contact details.</p>
        <table class="dialog-table dialog-shortcuts">
          <tbody>
            <tr><th>Edition</th><td>Portfolio Web Edition</td></tr>
            <tr><th>Interface</th><td>Word 2003 style · XP taskbar</td></tr>
            <tr><th>Renderer</th><td>${escapeHtml(navigator.userAgent.split(") ")[0])}</td></tr>
            <tr><th>Author</th><td>Gabriel Elorde</td></tr>
          </tbody>
        </table>
        <p class="dialog-note">Not affiliated with Microsoft. All chrome is drawn with CSS and inline SVG; no Microsoft logos are used.</p>`,
      buttons: [{ label: "OK", primary: true, value: null }],
    });
  }

  async function shutdownDialog() {
    const choice = await showDialog({
      title: "Turn off computer",
      body: `
        <p class="shutdown-question">What do you want the computer to do?</p>
        <label class="dialog-check"><input type="radio" name="shutdown-kind" value="off" checked> Turn off</label>
        <label class="dialog-check"><input type="radio" name="shutdown-kind" value="restart"> Restart</label>
        <label class="dialog-check"><input type="radio" name="shutdown-kind" value="standby"> Stand by</label>
        <p class="dialog-note">This is a simulated XP shutdown. Your browser and this portfolio stay running.</p>`,
      buttons: [
        {
          label: "OK",
          primary: true,
          collect: (body) => body.querySelector('input[name="shutdown-kind"]:checked').value,
        },
        { label: "Cancel", value: null },
      ],
    });
    if (!choice) return;
    closeStart();
    shutdownScreen.hidden = false;
    if (choice === "restart") {
      shutdownMessage.textContent = "Windows is restarting…";
      window.setTimeout(() => {
        shutdownScreen.hidden = true;
        restoreWindow();
        showToast("Windows has restarted. Portfolio.doc is ready.");
      }, 1500);
    } else if (choice === "standby") {
      shutdownMessage.textContent = "Standing by… click Restart to wake the computer.";
    } else {
      shutdownMessage.textContent = "It's now safe to turn off your computer.";
    }
  }

  function restartSession() {
    shutdownScreen.hidden = true;
    restoreWindow();
    showToast("Welcome back. Portfolio.doc is ready.");
  }

  function tickClock() {
    const now = new Date();
    let hours = now.getHours();
    const minutes = String(now.getMinutes()).padStart(2, "0");
    const suffix = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    clockButton.textContent = `${hours}:${minutes} ${suffix}`;
    clockButton.title = now.toLocaleDateString(undefined, {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
    });
  }

  async function showClock() {
    const now = new Date();
    showToast(now.toLocaleString(undefined, {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "numeric",
      minute: "2-digit",
    }));
  }

  function printDocument() {
    closeMenu();
    closeStart();
    window.print();
  }

  function copySelection(cut) {
    if (!ensureDocumentSelection()) return;
    try {
      document.execCommand(cut ? "cut" : "copy");
      if (cut) afterEdit();
      showToast(cut ? "Selection cut to the clipboard." : "Selection copied to the clipboard.");
    } catch (error) {
      console.error(error);
      showToast("Your browser blocked clipboard access. Use Ctrl+C or Ctrl+X.");
    }
  }

  async function pasteSelection() {
    if (!ensureDocumentSelection()) return;
    try {
      const text = await navigator.clipboard.readText();
      if (!text) {
        showToast("The clipboard is empty.");
        return;
      }
      executeCommand("insertText", text);
    } catch (error) {
      console.error(error);
      showToast("Your browser blocked clipboard access. Use Ctrl+V in the document.");
    }
  }

  function selectAllDocument() {
    const range = document.createRange();
    range.selectNodeContents(paper);
    const sel = currentSelection();
    paper.focus({ preventScroll: true });
    sel.removeAllRanges();
    sel.addRange(range);
    refreshToolbar();
    refreshStatus();
  }

  function undoEdit() {
    if (!ensureDocumentSelection()) return;
    document.execCommand("undo");
    afterEdit();
  }

  function redoEdit() {
    if (!ensureDocumentSelection()) return;
    document.execCommand("redo");
    afterEdit();
  }

  function showAssistant() {
    clippy.style.display = "";
    clippy.open = true;
    clippy.scrollIntoView({ behavior: reduceMotion.matches ? "auto" : "smooth", block: "center" });
    const summary = $(".clippy-summary", clippy);
    if (summary) summary.focus();
  }

  const actions = {
    "file-new": newDocument,
    "file-open": openDocumentDialog,
    "file-save": saveDocument,
    "file-save-as": saveAsDialog,
    "file-print": printDocument,
    "file-properties": propertiesDialog,
    "file-revert": revertDocument,
    "file-exit": requestClose,

    "edit-undo": undoEdit,
    "edit-redo": redoEdit,
    "edit-cut": () => copySelection(true),
    "edit-copy": () => copySelection(false),
    "edit-paste": pasteSelection,
    "edit-select-all": selectAllDocument,
    "edit-find": findDialog,
    "edit-clipboard": officeClipboard,

    "view-docmap": toggleDocmap,
    "view-ruler": toggleRuler,
    "view-standard": () => toggleToolbar("standard"),
    "view-format": () => toggleToolbar("formatting"),
    "view-zoom": zoomDialog,
    "view-fullscreen": fullScreen,

    "insert-page-break": insertPageBreak,
    "insert-date": insertDateDialog,
    "insert-rule": insertRule,
    "insert-hyperlink": insertHyperlinkDialog,
    "insert-table": insertTableDialog,

    "format-font": fontDialog,
    "format-style": styleDialog,
    "format-bold": () => executeCommand("bold"),
    "format-italic": () => executeCommand("italic"),
    "format-underline": () => executeCommand("underline"),
    "format-bullets": () => executeCommand("insertUnorderedList"),
    "format-numbers": () => executeCommand("insertOrderedList"),
    "format-indent": () => executeCommand("indent"),
    "format-border": bordersDialog,
    "format-highlight": toggleHighlight,
    "format-font-color": fontColorDialog,

    "align-left": () => executeCommand("justifyLeft"),
    "align-center": () => executeCommand("justifyCenter"),
    "align-right": () => executeCommand("justifyRight"),
    "align-justify": () => executeCommand("justifyFull"),

    "tools-spelling": spellingDialog,
    "tools-word-count": wordCountDialog,
    "tools-options": optionsDialog,
    "tools-language": languageDialog,

    "table-row-above": () => insertRow("above"),
    "table-row-below": () => insertRow("below"),
    "table-col-left": () => insertColumn("left"),
    "table-col-right": () => insertColumn("right"),
    "table-delete-row": deleteRow,
    "table-delete-col": deleteColumn,
    "table-select": selectCurrentTable,

    "window-minimize": minimizeWindow,
    "window-restore": toggleRestore,
    "window-close": requestClose,
    "focus-document": () => {
      if (root.classList.contains("is-closed")) reopenDocument();
      else if (root.classList.contains("is-minimized")) restoreWindow();
      else {
        paper.focus({ preventScroll: true });
        if (selectionInsidePaper()) rememberSelection();
      }
    },
    "task-toggle": toggleTask,
    "show-desktop": showDesktop,
    reopen: reopenDocument,
    "go-to-page": goToPageDialog,

    "help-topics": helpTopicsDialog,
    "help-shortcuts": shortcutsDialog,
    "help-about": () => aboutDialog(false),
    "help-about-word": () => aboutDialog(true),
    "help-assistant": showAssistant,
    "help-contact": () => window.open(GITHUB_URL, "_blank", "noopener"),

    "start-shutdown": shutdownDialog,
    "restart-session": restartSession,
    clock: showClock,
    "tray-volume": () => showToast("Volume: 100% · No audio device is available in this demo."),
    "tray-network": () => showToast("Local Area Connection · Connected at 100.0 Mbps"),
  };

  async function requestClose() {
    if (root.classList.contains("is-closed")) return;
    if (dirty) {
      const choice = await showDialog({
        title: "Microsoft Word",
        body: `<p>Do you want to save the changes you made to <strong>${escapeHtml(currentFileName)}</strong>?</p>`,
        buttons: [
          { label: "Yes", primary: true, value: "save" },
          { label: "No", value: "discard" },
          { label: "Cancel", value: "cancel" },
        ],
      });
      if (choice === "cancel" || choice === null) return;
      if (choice === "save") downloadDocument(currentFileName);
    }
    closeDocument();
  }

  function runAction(name) {
    const handler = actions[name];
    if (!handler) {
      showToast("That command is not available in this demo.");
      return;
    }
    closeMenu();
    closeStart();
    closeCombo();
    try {
      const result = handler();
      if (result && typeof result.catch === "function") {
        result.catch((error) => {
          console.error(error);
          showToast("Word could not complete that command.");
        });
      }
    } catch (error) {
      console.error(error);
      showToast("Word could not complete that command.");
    }
  }

  on(root, "mousedown", (event) => {
    const target = event.target.closest(
      ".tb, .fmt-btn, .combo, .menu-title, .menu-item, .window-controls button, " +
      ".task-btn, .start-btn, .tray button, .statusbar button, .docmap-filter, .help-box"
    );
    if (target) event.preventDefault();
  });

  on(comboPopup, "mousedown", (event) => event.preventDefault());

  on(startButton, "click", (event) => {
    event.preventDefault();
    toggleStart();
  });

  on(menuRoot, "click", (event) => {
    const title = event.target.closest(".menu-title");
    if (!title) return;
    event.preventDefault();
    event.stopPropagation();
    const menu = title.closest(".menu");
    if (menu.classList.contains("is-open")) closeMenu();
    else openMenu(menu);
  });

  on(menuRoot, "pointerover", (event) => {
    const title = event.target.closest(".menu-title");
    if (!title || !menuName) return;
    const menu = title.closest(".menu");
    if (menu.dataset.menuName !== menuName) openMenu(menu);
  });

  on(menuRoot, "keydown", (event) => {
    const title = event.target.closest(".menu-title");
    const item = event.target.closest(".menu-item");
    if (!title && !item) return;
    const menu = event.target.closest(".menu");
    const items = $$(".menu-item:not(:disabled)", menu);
    if (event.key === "ArrowDown" && title) {
      event.preventDefault();
      openMenu(menu, true);
      return;
    }
    if ((event.key === "Enter" || event.key === " ") && title) {
      event.preventDefault();
      openMenu(menu, true);
      return;
    }
    if (event.key === "ArrowDown" && item) {
      event.preventDefault();
      const index = items.indexOf(item);
      (items[index + 1] || items[0]).focus();
      return;
    }
    if (event.key === "ArrowUp" && item) {
      event.preventDefault();
      const index = items.indexOf(item);
      (items[index - 1] || items[items.length - 1]).focus();
      return;
    }
    if (event.key === "Home" && item) {
      event.preventDefault();
      if (items[0]) items[0].focus();
      return;
    }
    if (event.key === "End" && item) {
      event.preventDefault();
      if (items[items.length - 1]) items[items.length - 1].focus();
      return;
    }
    if (event.key === "ArrowRight" || event.key === "ArrowLeft") {
      event.preventDefault();
      const all = $$(".menu");
      const index = all.indexOf(menu);
      const next = all[(index + (event.key === "ArrowRight" ? 1 : all.length - 1)) % all.length];
      if (next) {
        openMenu(next);
        const nextTitle = $(".menu-title", next);
        if (nextTitle) nextTitle.focus();
      }
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      const label = $(".menu-title", menu);
      closeMenu();
      if (label) label.focus();
      return;
    }
    if (event.key === "Tab") closeMenu();
  });

  on(startMenu, "keydown", (event) => {
    const items = $$(".sm-item, .sm-footer button", startMenu);
    const index = items.indexOf(document.activeElement);
    if (event.key === "ArrowDown") {
      event.preventDefault();
      (items[index + 1] || items[0]).focus();
    } else if (event.key === "ArrowUp") {
      event.preventDefault();
      (items[index - 1] || items[items.length - 1]).focus();
    } else if (event.key === "Escape") {
      event.preventDefault();
      closeStart();
    }
  });

  on(document, "pointerdown", (event) => {
    if (!event.target.closest(".menu")) closeMenu();
    if (!event.target.closest("#start-menu") && !event.target.closest("#start-button")) closeStart();
    if (!event.target.closest(".combo") && !event.target.closest("#combo-popup")) closeCombo();
  });

  on(root, "click", (event) => {
    if (suppressShortcutClick && event.target.closest(".desktop-shortcut")) {
      suppressShortcutClick = false;
      event.preventDefault();
      return;
    }

    const actionElement = event.target.closest("[data-action]");
    if (actionElement && !actionElement.disabled) {
      event.preventDefault();
      if (actionElement.closest(".clippy")) clippy.open = false;
      const action = actionElement.dataset.action;
      runAction(action);
      return;
    }

    const combo = event.target.closest("[data-combo]");
    if (combo) {
      event.preventDefault();
      if (comboState && comboState.combo === combo) closeCombo();
      else openCombo(combo);
      return;
    }

    const anchor = event.target.closest("a");
    if (!anchor) return;
    const href = anchor.getAttribute("href") || "";
    if (href.startsWith("#")) {
      event.preventDefault();
      if (anchor.closest(".clippy")) clippy.open = false;
      scrollToHash(href);
      return;
    }
    if (paper.contains(anchor)) {
      event.preventDefault();
      window.open(anchor.href, "_blank", "noopener");
    }
  });

  on(document, "keydown", (event) => {
    if (event.key === "Escape") {
      if (dialogState) {
        event.preventDefault();
        closeDialog(null);
        return;
      }
      if (!comboPopup.hidden) {
        event.preventDefault();
        closeCombo();
        return;
      }
      if (!startMenu.hidden) {
        event.preventDefault();
        closeStart();
        return;
      }
      if (menuName) {
        event.preventDefault();
        closeMenu();
        return;
      }
    }

    if (dialogState) return;

    const ctrl = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();

    if (event.altKey && !ctrl && !event.shiftKey) {
      const menus = {
        f: "file",
        e: "edit",
        v: "view",
        i: "insert",
        o: "format",
        t: "tools",
        a: "table",
        w: "window",
        h: "help",
      };
      if (menus[key]) {
        event.preventDefault();
        openMenuByName(menus[key], true);
      }
      return;
    }

    if (ctrl && !event.altKey) {
      if (key === "s") {
        event.preventDefault();
        if (event.shiftKey) saveAsDialog();
        else saveDocument();
        return;
      }
      if (key === "p") {
        event.preventDefault();
        printDocument();
        return;
      }
      if (key === "o") {
        event.preventDefault();
        openDocumentDialog();
        return;
      }
      if (key === "n") {
        event.preventDefault();
        newDocument();
        return;
      }
      if (key === "f" || key === "h") {
        event.preventDefault();
        findDialog();
        return;
      }
      if (key === "k") {
        event.preventDefault();
        insertHyperlinkDialog();
        return;
      }
      if (key === "a" && selectionInsidePaper()) {
        event.preventDefault();
        selectAllDocument();
        return;
      }
      if (["b", "i", "u"].includes(key) && selectionInsidePaper()) {
        event.preventDefault();
        runAction(`format-${{ b: "bold", i: "italic", u: "underline" }[key]}`);
        return;
      }
      if ((key === "z" || key === "y") && selectionInsidePaper()) {
        window.setTimeout(afterEdit, 0);
        return;
      }
      if (event.key === "Enter" && selectionInsidePaper()) {
        event.preventDefault();
        insertPageBreak();
        return;
      }
    }

    if (event.key === "F1") {
      event.preventDefault();
      helpTopicsDialog();
    } else if (event.key === "F7") {
      event.preventDefault();
      spellingDialog();
    } else if (event.key === "F11") {
      event.preventDefault();
      fullScreen();
    } else if (event.key === "F12") {
      event.preventDefault();
      saveAsDialog();
    }
  });

  on(titlebar, "dblclick", (event) => {
    if (suppressTitlebarDblClick) {
      suppressTitlebarDblClick = false;
      return;
    }
    if (!event.target.closest("button")) toggleRestore();
  });

  on(paper, "input", () => {
    afterEdit();
    scheduleStatus();
  });
  on(paper, "mouseup", () => {
    rememberSelection();
    scheduleToolbar();
    scheduleStatus();
    updateTableCommands();
  });
  on(paper, "keyup", () => {
    rememberSelection();
    scheduleToolbar();
    scheduleStatus();
    updateTableCommands();
  });
  on(document, "selectionchange", () => {
    if (!selectionInsidePaper()) return;
    rememberSelection();
    scheduleToolbar();
    scheduleStatus();
    updateTableCommands();
  });
  on(editor, "scroll", scheduleStatus, { passive: true });
  on(window, "resize", () => {
    closeCombo();
    scheduleStatus();
  });
  on(document, "fullscreenchange", () => {
    showToast(document.fullscreenElement ? "Full screen on. Press Esc to exit." : "Full screen off.");
  });

  updateTitle();
  refreshChecks();
  refreshToolbar();
  refreshStatus();
  updateTableCommands();
  statusZoom.textContent = "100%";
  tickClock();
  clockTimer = window.setInterval(tickClock, 30000);

  const cleanup = () => {
    cleanupFns.splice(0).forEach((fn) => fn());
    window.clearTimeout(toastTimer);
    window.clearInterval(clockTimer);
    if (dialogState) closeDialog(null);
    if (window.__wordUiCleanup === cleanup) window.__wordUiCleanup = null;
  };
  window.__wordUiCleanup = cleanup;
})();
