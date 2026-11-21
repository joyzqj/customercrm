/*!
 * 客户通 CustomerCRM — js/core.js
 * 基础工具：DOM、转义、下载、Toast、Modal、Confirm、图标、localStorage 安全封装。
 * 浏览器挂到 window.CrmCore。
 */
(function (root) {
  'use strict';

  var core = {};

  /* ------------------------------------------------------------------ */
  /* 1. DOM 与转义                                                        */
  /* ------------------------------------------------------------------ */

  var ESC_MAP = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

  core.esc = function (value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/[&<>"']/g, function (ch) { return ESC_MAP[ch]; });
  };

  core.qs = function (selector, scope) {
    return (scope || document).querySelector(selector);
  };

  core.qsa = function (selector, scope) {
    return Array.prototype.slice.call((scope || document).querySelectorAll(selector));
  };

  core.on = function (target, type, handler, options) {
    if (!target) return function () {};
    target.addEventListener(type, handler, options);
    return function off() { target.removeEventListener(type, handler, options); };
  };

  core.delegate = function (scope, selector, type, handler) {
    return core.on(scope, type, function (event) {
      var node = event.target;
      while (node && node !== scope) {
        if (node.matches && node.matches(selector)) {
          handler.call(node, event, node);
          return;
        }
        node = node.parentNode;
      }
    });
  };

  core.debounce = function (fn, wait) {
    var timer = null;
    return function debounced() {
      var args = arguments;
      var self = this;
      clearTimeout(timer);
      timer = setTimeout(function () { fn.apply(self, args); }, wait);
    };
  };

  /** 计算属性并生成内联 style 字符串（用于条形图宽度） */
  core.styleWidth = function (percent) {
    var value = Math.max(0, Math.min(100, Number(percent) || 0));
    return 'width:' + value.toFixed(1) + '%';
  };

  core.isEmpty = function (value) {
    return value === null || value === undefined || String(value).trim() === '';
  };

  core.formatInt = function (value) {
    var n = Number(value) || 0;
    return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  };

  core.percent = function (part, total, digits) {
    if (!total) return '0%';
    var p = (Number(part) || 0) / total * 100;
    return p.toFixed(digits === undefined ? 0 : digits) + '%';
  };

  core.truncate = function (text, max) {
    var s = String(text === null || text === undefined ? '' : text);
    var limit = max || 60;
    return s.length > limit ? s.slice(0, limit - 1) + '…' : s;
  };

  core.dayLabel = function (days) {
    if (days === null || days === undefined || !isFinite(days)) return '尚未跟进';
    if (days <= 0) return '今天';
    return days + ' 天前';
  };

  /* ------------------------------------------------------------------ */
  /* 2. 本地存储（localStorage 不可用时降级为内存）                        */
  /* ------------------------------------------------------------------ */

  var memory = {};

  core.storage = {
    available: function () {
      try {
        var key = '__crm_probe__';
        window.localStorage.setItem(key, '1');
        window.localStorage.removeItem(key);
        return true;
      } catch (err) { return false; }
    },
    get: function (key, fallback) {
      try {
        var raw = window.localStorage.getItem(key);
        if (raw === null) return key in memory ? memory[key] : fallback;
        return JSON.parse(raw);
      } catch (err) {
        return key in memory ? memory[key] : fallback;
      }
    },
    set: function (key, value) {
      memory[key] = value;
      try {
        window.localStorage.setItem(key, JSON.stringify(value));
        return true;
      } catch (err) {
        return false;
      }
    },
    remove: function (key) {
      delete memory[key];
      try { window.localStorage.removeItem(key); return true; } catch (err) { return false; }
    }
  };

  /* ------------------------------------------------------------------ */
  /* 3. 下载 / 剪贴板                                                     */
  /* ------------------------------------------------------------------ */

  core.download = function (filename, data, mime) {
    try {
      var blob = data instanceof Blob ? data : new Blob([data], { type: mime || 'application/octet-stream' });
      var url = URL.createObjectURL(blob);
      var link = document.createElement('a');
      link.href = url;
      link.download = filename;
      link.rel = 'noopener';
      link.style.display = 'none';
      document.body.appendChild(link);
      link.click();
      setTimeout(function () {
        URL.revokeObjectURL(url);
        if (link.parentNode) link.parentNode.removeChild(link);
      }, 1200);
      return true;
    } catch (err) {
      return false;
    }
  };

  core.saveTextFile = function (filename, text, mime) {
    var type = mime || 'text/plain;charset=utf-8';
    return core.download(filename, new Blob([text], { type: type }), type);
  };

  core.copyText = function (text) {
    return new Promise(function (resolve) {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () { resolve(true); }, function () { resolve(false); });
        return;
      }
      try {
        var area = document.createElement('textarea');
        area.value = text;
        area.setAttribute('readonly', 'readonly');
        area.style.position = 'fixed';
        area.style.opacity = '0';
        document.body.appendChild(area);
        area.select();
        var ok = document.execCommand('copy');
        document.body.removeChild(area);
        resolve(!!ok);
      } catch (err) { resolve(false); }
    });
  };

  core.pickFile = function (accept) {
    return new Promise(function (resolve) {
      var input = document.createElement('input');
      input.type = 'file';
      input.accept = accept || '';
      input.style.display = 'none';
      input.addEventListener('change', function () {
        var file = input.files && input.files[0];
        if (!file) { resolve(null); return; }
        var reader = new FileReader();
        reader.onload = function () { resolve({ name: file.name, text: String(reader.result || '') }); };
        reader.onerror = function () { resolve(null); };
        reader.readAsText(file, 'utf-8');
      });
      document.body.appendChild(input);
      input.click();
      setTimeout(function () { if (input.parentNode) input.parentNode.removeChild(input); }, 60000);
    });
  };

  /* ------------------------------------------------------------------ */
  /* 4. Toast                                                            */
  /* ------------------------------------------------------------------ */

  var TOAST_ICONS = {
    ok: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m5 13 4 4L19 7"></path></svg>',
    error: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><path d="M12 8v5"></path><path d="M12 16.5h.01"></path><circle cx="12" cy="12" r="9"></circle></svg>',
    info: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="M12 11v5M12 8h.01"></path></svg>'
  };

  /**
   * 轻提示。type: 'ok' | 'error' | 'info'
   * opts: { timeout, action: { label, onClick } }
   */
  core.toast = function (message, type, opts) {
    var stack = document.getElementById('toast-stack');
    if (!stack) return null;
    var kind = type === 'ok' || type === 'error' ? type : 'info';
    var node = document.createElement('div');
    node.className = 'toast toast--' + kind;
    node.setAttribute('role', kind === 'error' ? 'alert' : 'status');
    node.innerHTML =
      '<span class="toast-icon" aria-hidden="true">' + (TOAST_ICONS[kind] || TOAST_ICONS.info) + '</span>' +
      '<span class="toast-text">' + core.esc(message) + '</span>' +
      (opts && opts.action ? '<button type="button" class="toast-action btn btn--sm btn--ghost">' + core.esc(opts.action.label) + '</button>' : '') +
      '<button type="button" class="toast-close" aria-label="关闭提示">' +
      '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18"></path></svg></button>';
    stack.appendChild(node);

    var timer = null;
    function dismiss() {
      if (!node.parentNode) return;
      clearTimeout(timer);
      node.classList.add('is-leaving');
      setTimeout(function () { if (node.parentNode) node.parentNode.removeChild(node); }, 220);
    }
    var closeBtn = node.querySelector('.toast-close');
    if (closeBtn) closeBtn.addEventListener('click', dismiss);
    if (opts && opts.action) {
      var actionBtn = node.querySelector('.toast-action');
      if (actionBtn) {
        actionBtn.addEventListener('click', function () {
          try { opts.action.onClick(); } finally { dismiss(); }
        });
      }
    }
    timer = setTimeout(dismiss, (opts && opts.timeout) || 3200);
    // 最多保留 4 条，避免堆叠过多
    var nodes = stack.querySelectorAll('.toast');
    if (nodes.length > 4) stack.removeChild(nodes[0]);
    return node;
  };

  /* ------------------------------------------------------------------ */
  /* 5. Modal                                                            */
  /* ------------------------------------------------------------------ */

  var modalStack = [];

  function syncModalState() {
    var open = modalStack.length > 0;
    document.body.classList.toggle('is-modal-open', open);
  }

  var modal = {
    /**
     * open({ id, title, subtitle, body, footer, size, closable, onMount, onClose })
     * body / footer 为 HTML 字符串；返回 { id, root, close }
     */
    open: function (options) {
      var opts = options || {};
      var rootEl = document.getElementById('modal-root');
      if (!rootEl) return null;

      var id = opts.id || ('modal_' + (++modal._seq));
      var existing = rootEl.querySelector('[data-modal-id="' + id + '"]');
      if (existing) rootEl.removeChild(existing);

      var backdrop = document.createElement('div');
      backdrop.className = 'modal-backdrop';
      backdrop.setAttribute('data-modal-id', id);

      var sizeClass = opts.size === 'lg' ? ' modal--lg' : opts.size === 'sm' ? ' modal--sm' : '';
      var titleId = id + '-title';
      backdrop.innerHTML =
        '<div class="modal' + sizeClass + '" role="dialog" aria-modal="true" aria-labelledby="' + titleId + '">' +
          '<div class="modal-head">' +
            '<div class="modal-heading">' +
              '<h2 class="modal-title" id="' + titleId + '">' + core.esc(opts.title || '') + '</h2>' +
              (opts.subtitle ? '<p class="modal-sub">' + core.esc(opts.subtitle) + '</p>' : '') +
            '</div>' +
            '<button type="button" class="icon-btn modal-close" data-modal-close aria-label="关闭">' +
              '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18"></path></svg>' +
            '</button>' +
          '</div>' +
          '<div class="modal-body">' + (opts.body || '') + '</div>' +
          (opts.footer ? '<div class="modal-foot">' + opts.footer + '</div>' : '') +
        '</div>';

      rootEl.appendChild(backdrop);
      modalStack.push({ id: id, node: backdrop, onClose: opts.onClose || null });
      syncModalState();

      var previousFocus = document.activeElement;

      backdrop.addEventListener('mousedown', function (event) {
        if (event.target === backdrop && opts.closable !== false) modal.close(id);
      });
      // 所有 [data-modal-close]（头部 ×、底部取消）都关闭当前弹窗
      backdrop.addEventListener('click', function (event) {
        var target = event.target;
        if (!target || !target.closest) return;
        if (target.closest('[data-modal-close]')) {
          event.preventDefault();
          modal.close(id);
        }
      });

      setTimeout(function () {
        var focusTarget = backdrop.querySelector('[autofocus]') ||
          backdrop.querySelector('input:not([type="hidden"]):not([disabled]), select:not([disabled]), textarea:not([disabled])') ||
          backdrop.querySelector('.modal-close');
        if (focusTarget) { try { focusTarget.focus({ preventScroll: false }); } catch (err) { focusTarget.focus(); } }
      }, 30);

      var handle = {
        id: id,
        root: backdrop,
        close: function () { modal.close(id); },
        restoreFocus: previousFocus
      };
      if (typeof opts.onMount === 'function') {
        try { opts.onMount(backdrop, handle); } catch (err) { console.warn('[CustomerCRM] modal onMount 失败：', err); }
      }
      return handle;
    },

    close: function (id) {
      var index = -1;
      if (id === undefined) index = modalStack.length - 1;
      else {
        for (var i = 0; i < modalStack.length; i += 1) { if (modalStack[i].id === id) { index = i; break; } }
      }
      if (index === -1) return;
      var entry = modalStack.splice(index, 1)[0];
      if (entry.node && entry.node.parentNode) entry.node.parentNode.removeChild(entry.node);
      syncModalState();
      if (typeof entry.onClose === 'function') {
        try { entry.onClose(); } catch (err) { console.warn('[CustomerCRM] modal onClose 失败：', err); }
      }
      var top = modalStack[modalStack.length - 1];
      if (top && top.node) {
        var focusTarget = top.node.querySelector('input, select, textarea, .modal-close');
        if (focusTarget) { try { focusTarget.focus({ preventScroll: true }); } catch (err) { /* ignore */ } }
      }
    },

    closeAll: function () {
      while (modalStack.length) modal.close();
    },

    isOpen: function () { return modalStack.length > 0; },

    _seq: 0
  };

  core.modal = modal;

  /**
   * Promise 版确认框：{ title, message, confirmText, cancelText, danger }
   * 返回 Promise<boolean>
   */
  core.confirm = function (options) {
    var opts = options || {};
    return new Promise(function (resolve) {
      var settled = false;
      function settle(value) {
        if (settled) return;
        settled = true;
        resolve(value);
      }
      var handle = core.modal.open({
        title: opts.title || '请确认',
        size: 'sm',
        body: '<p class="confirm-text">' + core.esc(opts.message || '') + '</p>',
        footer:
          '<button type="button" class="btn btn--ghost" data-confirm-cancel>' + core.esc(opts.cancelText || '取消') + '</button>' +
          '<button type="button" class="btn ' + (opts.danger ? 'btn--danger' : 'btn--primary') + '" data-confirm-ok>' + core.esc(opts.confirmText || '确认') + '</button>',
        onClose: function () { settle(false); }
      });
      if (!handle) { settle(false); return; }
      var okBtn = handle.root.querySelector('[data-confirm-ok]');
      var cancelBtn = handle.root.querySelector('[data-confirm-cancel]');
      if (okBtn) okBtn.addEventListener('click', function () { settle(true); core.modal.close(handle.id); });
      if (cancelBtn) cancelBtn.addEventListener('click', function () { core.modal.close(handle.id); });
    });
  };

  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && modalStack.length) {
      var top = modalStack[modalStack.length - 1];
      var dialog = top.node.querySelector('.modal');
      if (dialog) core.modal.close(top.id);
    }
  });

  /* ------------------------------------------------------------------ */
  /* 6. 图标（内联 SVG，stroke=currentColor）                              */
  /* ------------------------------------------------------------------ */

  var ICON_PATHS = {
    search: '<circle cx="11" cy="11" r="7"></circle><path d="m20 20-3.6-3.6"></path>',
    plus: '<path d="M12 5v14M5 12h14"></path>',
    edit: '<path d="M12 20h9"></path><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"></path>',
    trash: '<path d="M3 6h18"></path><path d="M8 6V4h8v2"></path><path d="M19 6l-1 14H6L5 6"></path><path d="M10 11v5M14 11v5"></path>',
    download: '<path d="M12 3v12"></path><path d="m7 10 5 5 5-5"></path><path d="M4 21h16"></path>',
    upload: '<path d="M12 17V5"></path><path d="m7 10 5-5 5 5"></path><path d="M4 21h16"></path>',
    calendar: '<rect x="3" y="5" width="18" height="16" rx="2"></rect><path d="M16 3v4M8 3v4M3 11h18"></path>',
    users: '<path d="M17 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path><circle cx="9.5" cy="7" r="4"></circle><path d="M22 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path>',
    close: '<path d="M6 6l12 12M18 6 6 18"></path>',
    check: '<path d="m5 13 4 4L19 7"></path>',
    alert: '<path d="M12 9v4M12 17h.01"></path><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"></path>',
    chevronDown: '<path d="m6 9 6 6 6-6"></path>',
    chevronRight: '<path d="m9 6 6 6-6 6"></path>',
    arrowRight: '<path d="M4 12h15"></path><path d="m13 6 6 6-6 6"></path>',
    sparkles: '<path d="M12 3v3M12 18v3M3 12h3M18 12h3M5.6 5.6l2.1 2.1M16.3 16.3l2.1 2.1M18.4 5.6l-2.1 2.1M7.7 16.3l-2.1 2.1"></path><circle cx="12" cy="12" r="3.4"></circle>',
    refresh: '<path d="M21 12a9 9 0 1 1-3.2-6.9"></path><path d="M21 4v5h-5"></path>',
    filter: '<path d="M3 5h18l-7 8v6l-4 2v-8z"></path>',
    clipboard: '<rect x="8" y="3" width="8" height="4" rx="1"></rect><path d="M16 5h2a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2h2"></path>',
    database: '<ellipse cx="12" cy="5" rx="8" ry="3"></ellipse><path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5"></path><path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"></path>',
    file: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"></path><path d="M14 3v5h5"></path><path d="M9 13h6M9 17h4"></path>',
    clock: '<circle cx="12" cy="12" r="9"></circle><path d="M12 7v5l3.2 2"></path>',
    building: '<rect x="4" y="3" width="16" height="18" rx="1.6"></rect><path d="M9 7h2M13 7h2M9 11h2M13 11h2M9 15h2M13 15h2"></path>',
    globe: '<circle cx="12" cy="12" r="9"></circle><path d="M3 12h18"></path><path d="M12 3a15 15 0 0 1 0 18a15 15 0 0 1 0-18z"></path>',
    chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"></path>',
    target: '<circle cx="12" cy="12" r="8"></circle><circle cx="12" cy="12" r="3.4"></circle>',
    settings: '<path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3"></path><path d="M1 14h6M9 8h6M17 16h6"></path>',
    help: '<circle cx="12" cy="12" r="9"></circle><path d="M9.1 9a3 3 0 0 1 5.8 1c0 2-3 2.5-3 4"></path><path d="M12 17h.01"></path>',
    history: '<path d="M3 12a9 9 0 1 0 3.2-6.9"></path><path d="M3 4v5h5"></path><path d="M12 8v4.5l3 1.8"></path>',
    copy: '<rect x="9" y="9" width="12" height="12" rx="2"></rect><path d="M5 15H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v1"></path>'
  };

  core.icon = function (name, size) {
    var path = ICON_PATHS[name];
    if (!path) return '';
    var s = size || 16;
    return '<svg viewBox="0 0 24 24" width="' + s + '" height="' + s + '" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + path + '</svg>';
  };

  /* ------------------------------------------------------------------ */
  /* 7. 其它                                                             */
  /* ------------------------------------------------------------------ */

  /** 让 .table-wrap 内的滚动条不遮挡圆角（用于重绘后夹紧） */
  core.scrollIntoViewIfNeeded = function (element) {
    if (!element) return;
    try { element.scrollIntoView({ block: 'nearest', inline: 'nearest' }); } catch (err) { /* ignore */ }
  };

  core.focusFirst = function (scope) {
    var target = scope && scope.querySelector('input:not([type="hidden"]), select, textarea, button');
    if (target) { try { target.focus(); } catch (err) { /* ignore */ } }
  };

  core.isTextInput = function (element) {
    if (!element) return false;
    var tag = element.tagName;
    return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || element.isContentEditable;
  };

  root.CrmCore = core;
})(typeof window !== 'undefined' ? window : globalThis);
