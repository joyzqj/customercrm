/*!
 * 客户通 CustomerCRM — js/app.js
 * 路由 / 页面渲染 / 表格 / 表单 / 智能识别 / 仪表盘 / 设置 / 帮助。
 * 依赖：js/schema.js、js/core.js、js/parser.js、js/xlsx.js、js/store.js
 */
(function (root) {
  'use strict';

  var schema = root.CRM_SCHEMA;
  var core = root.CrmCore;
  var Store = root.CrmStore;
  var esc = core.esc;
  var icon = core.icon;

  var PAYLOAD_ROUTES = ['dashboard', 'customers', 'recognize', 'settings', 'help'];

  var PAGE_META = {
    dashboard: { title: '仪表盘', crumb: '业务 / 仪表盘' },
    customers: { title: '客户台账', crumb: '业务 / 客户台账' },
    recognize: { title: '智能识别', crumb: '业务 / 智能识别' },
    settings: { title: '设置', crumb: '系统 / 设置' },
    help: { title: '使用说明', crumb: '系统 / 使用说明' }
  };

  var DEFAULT_FILTERS = { q: '', country: '', status: '', deal: '', sort: 'updated-desc' };

  var state = {
    route: 'dashboard',
    filters: Object.assign({}, DEFAULT_FILTERS),
    expanded: {},
    recog: { text: '', result: null },
    lastRecogCount: 0
  };

  var unsubscribers = [];

  /* ================================================================== */
  /* 1. 路由                                                             */
  /* ================================================================== */

  function normalizeHash(hash) {
    var value = String(hash || '').replace(/^#/, '');
    if (!value) return '#/dashboard';
    if (value.charAt(0) !== '/') value = '/' + value;
    var parts = value.split('/').filter(Boolean);
    var name = parts[0] || 'dashboard';
    if (PAYLOAD_ROUTES.indexOf(name) === -1) name = 'dashboard';
    return '#/' + name;
  }

  function parseRoute() {
    var value = normalizeHash(root.location.hash);
    var name = value.replace(/^#\//, '');
    return { name: name, hash: value };
  }

  function navigate(hash) {
    var target = normalizeHash(hash);
    if (root.location.hash === target) {
      render();
      return;
    }
    root.location.hash = target;
  }

  function getRoute() {
    return parseRoute().name;
  }

  /* ================================================================== */
  /* 2. 通用片段                                                          */
  /* ================================================================== */

  function badge(text, tone) {
    return '<span class="badge badge--' + esc(tone || 'neutral') + '">' + esc(text || '—') + '</span>';
  }

  function flagText(country) {
    var flag = country ? schema.countryFlag(country) : '';
    return flag ? '<span class="flag" aria-hidden="true">' + flag + '</span>' : '';
  }

  function ageCell(rec) {
    if (rec.daysSince === null || rec.daysSince === undefined) {
      return '<span class="age age--none">尚未跟进</span>';
    }
    var label = core.dayLabel(rec.daysSince);
    return '<span class="age age--' + esc(rec.ageTone) + '">' + esc(label) + '</span>' +
      '<span class="age-date">' + esc(schema.formatShort(rec.lastFollowupAt)) + '</span>';
  }

  function emptyState(options) {
    var opts = options || {};
    return '<div class="empty">' +
      '<div class="empty-art" aria-hidden="true">' +
        '<svg viewBox="0 0 120 90" width="120" height="90" fill="none">' +
          '<rect x="16" y="18" width="88" height="60" rx="8" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="1.5"></rect>' +
          '<path d="M28 34h64M28 46h48M28 58h56" stroke="#cbd5e1" stroke-width="2.6" stroke-linecap="round"></path>' +
          '<circle cx="94" cy="66" r="14" fill="#ccfbf1" stroke="#0d9488" stroke-width="1.6"></circle>' +
          '<path d="M88 66h12M94 60v12" stroke="#0f766e" stroke-width="2" stroke-linecap="round"></path>' +
        '</svg>' +
      '</div>' +
      '<h3 class="empty-title">' + esc(opts.title || '暂无数据') + '</h3>' +
      '<p class="empty-text">' + esc(opts.text || '') + '</p>' +
      (opts.actions ? '<div class="empty-actions">' + opts.actions + '</div>' : '') +
      '</div>';
  }

  function fieldHintFor(key) {
    var field = schema.fieldByKey(key);
    return field && field.hint ? field.hint : '';
  }

  /* ================================================================== */
  /* 3. 客户表单                                                          */
  /* ================================================================== */

  function fieldControl(field, value) {
    var id = 'f-' + field.key;
    var name = field.key;
    var common = 'id="' + id + '" name="' + name + '" class="' + (field.type === 'select' ? 'select' : 'input') + '"';

    if (field.type === 'date') {
      return '<input type="date" ' + common + ' value="' + esc(value || '') + '" required>';
    }
    if (field.type === 'select' && field.options === 'status') {
      return '<select ' + common + '>' + schema.FOLLOWUP_STATUS.map(function (status) {
        return '<option value="' + esc(status) + '"' + (status === (value || '新询盘') ? ' selected' : '') + '>' + esc(status) + '</option>';
      }).join('') + '</select>';
    }
    if (field.type === 'select' && field.options === 'deal') {
      return '<select ' + common + '>' + schema.DEAL_STATUS.map(function (status) {
        return '<option value="' + esc(status) + '"' + (status === (value || '未成交') ? ' selected' : '') + '>' + esc(status) + '</option>';
      }).join('') + '</select>';
    }
    if (field.type === 'select') {
      return '<input type="text" ' + common + ' list="country-options" value="' + esc(value || '') + '" placeholder="' + esc(field.placeholder || '') + '" autocomplete="off">';
    }
    if (field.type === 'textarea') {
      return '<textarea ' + common.replace('class="input"', 'class="textarea"') + ' rows="3" placeholder="' + esc(field.placeholder || '') + '">' + esc(value || '') + '</textarea>';
    }
    return '<input type="text" ' + common + ' value="' + esc(value || '') + '" placeholder="' + esc(field.placeholder || '') + '">';
  }

  function fieldHtml(field, value, full) {
    var label = esc(field.label);
    return '<div class="field' + (full || field.type === 'textarea' ? ' field--full' : '') + '" data-field="' + field.key + '">' +
      '<label class="field-label" for="f-' + field.key + '">' + label +
      (field.required ? '<span class="req" aria-hidden="true">*</span>' : '') +
      '</label>' +
      fieldControl(field, value) +
      (field.hint ? '<p class="field-hint">' + esc(field.hint) + '</p>' : '') +
      '<p class="field-error" data-error-for="' + field.key + '" hidden></p>' +
      '</div>';
  }

  function followupRowHtml(item, index) {
    return '<div class="followup-row" data-followup-row>' +
      '<input type="date" class="input followup-date-input" data-followup-date value="' + esc(item.date || '') + '" aria-label="第 ' + (index + 1) + ' 次跟进日期">' +
      '<input type="text" class="input followup-note-input" data-followup-note value="' + esc(item.note || '') + '" placeholder="跟进内容，如：发送报价单，客户要求 5% 折扣" aria-label="第 ' + (index + 1) + ' 次跟进内容" maxlength="300">' +
      '<button type="button" class="icon-btn icon-btn--sm icon-btn--danger" data-action="remove-followup-row" aria-label="删除这条跟进记录">' + icon('trash', 14) + '</button>' +
      '</div>';
  }

  function customerFormHtml(record, notice) {
    var rec = record || {};
    var groups = schema.FIELD_GROUPS.map(function (group) {
      return '<div class="form-section">' +
        '<div class="form-section-title">' + esc(group.title) + '</div>' +
        (group.hint ? '<p class="field-hint form-section-hint">' + esc(group.hint) + '</p>' : '') +
        '<div class="form-grid">' + group.keys.map(function (key) {
          var field = schema.fieldByKey(key);
          return fieldHtml(field, rec[key] || '');
        }).join('') + '</div>' +
        '</div>';
    }).join('');

    var followups = (rec.followups || []).slice(0, 3);
    var rows = followups.map(followupRowHtml).join('');

    return '<form id="customer-form" data-testid="form-customer" data-mode="' + esc(rec.__mode || 'create') + '" data-id="' + esc(rec.id || '') + '" class="customer-form" novalidate autocomplete="off">' +
      (notice ? '<p class="form-notice">' + icon('sparkles', 15) + '<span>' + esc(notice) + '</span></p>' : '') +
      '<div class="form-error" id="form-error" data-testid="form-error" hidden></div>' +
      groups +
      '<div class="form-section">' +
        '<div class="form-section-title">最近三次跟进</div>' +
        '<p class="field-hint form-section-hint">记录最近三次跟进（最新的排最前），「距上次跟进」自动按第一条日期计算。</p>' +
        '<div class="followup-editor" id="followup-editor">' +
          (rows || '<p class="followup-empty">还没有跟进记录，点下方按钮补充第一次跟进。</p>') +
        '</div>' +
        '<button type="button" class="btn btn--ghost btn--sm" data-action="add-followup-row">' + icon('plus', 14) + ' 添加一次跟进</button>' +
      '</div>' +
      '<datalist id="country-options">' + schema.countryOptions().map(function (name) {
        return '<option value="' + esc(name) + '"></option>';
      }).join('') + '</datalist>' +
      '</form>';
  }

  function openCustomerModal(record, options) {
    var opts = options || {};
    var isEdit = opts.mode === 'edit' && record && record.id;
    var data = isEdit ? Store.get(record.id) : schema.normalize(record || schema.emptyRecord());
    if (!data) data = schema.emptyRecord();
    data.__mode = isEdit ? 'edit' : 'create';

    var handle = core.modal.open({
      id: 'customer-modal',
      size: 'lg',
      title: isEdit ? '编辑客户' : '新建客户',
      subtitle: isEdit ? (data.customerName + (data.companyName ? ' · ' + data.companyName : '')) : '填写客户与公司信息，保存到本机浏览器',
      body: customerFormHtml(data, opts.notice),
      footer: '<button type="button" class="btn btn--ghost" data-testid="btn-cancel-customer" data-modal-close>取消</button>' +
        '<button type="button" class="btn btn--primary" data-testid="btn-save-customer" data-action="save-customer-form">' + icon('check', 16) + ' 保存客户</button>',
      onMount: function (container) {
        var firstEmpty = null;
        core.qsa('.form-grid input, .form-grid select, .form-grid textarea', container).forEach(function (el) {
          if (!firstEmpty && el.name !== 'followupDate' && !el.value) firstEmpty = el;
        });
        if (firstEmpty) setTimeout(function () { try { firstEmpty.focus(); } catch (err) { /* ignore */ } }, 60);
      }
    });
    return handle;
  }

  function collectFormValues() {
    var form = document.getElementById('customer-form');
    if (!form) return null;
    var data = {};
    schema.FIELD_KEYS.forEach(function (key) {
      var el = form.elements[key];
      data[key] = el ? String(el.value || '').trim() : '';
    });
    var followups = [];
    core.qsa('[data-followup-row]', form).forEach(function (row) {
      var dateEl = row.querySelector('[data-followup-date]');
      var noteEl = row.querySelector('[data-followup-note]');
      var date = dateEl ? String(dateEl.value || '').trim() : '';
      var note = noteEl ? String(noteEl.value || '').trim() : '';
      if (!date && !note) return;
      followups.push({ date: date, note: note });
    });
    data.followups = followups;
    data.__mode = form.getAttribute('data-mode');
    data.__id = form.getAttribute('data-id');
    return data;
  }

  function clearFormErrors(form) {
    if (!form) return;
    core.qsa('.field-error', form).forEach(function (el) { el.hidden = true; el.textContent = ''; });
    core.qsa('[aria-invalid="true"]', form).forEach(function (el) { el.removeAttribute('aria-invalid'); });
    var box = document.getElementById('form-error');
    if (box) { box.hidden = true; box.textContent = ''; }
  }

  var ERROR_LABELS = { followupDate: '最近三次跟进' };

  function showFormErrors(form, errors) {
    var keys = Object.keys(errors);
    var firstEl = null;
    keys.forEach(function (key) {
      var el = form.elements[key];
      var tip = form.querySelector('[data-error-for="' + key + '"]');
      if (tip) { tip.textContent = errors[key]; tip.hidden = false; }
      if (el) {
        el.setAttribute('aria-invalid', 'true');
        if (!firstEl) firstEl = el;
      }
    });
    if (firstEl) { try { firstEl.focus(); } catch (err) { /* ignore */ } }
    var box = document.getElementById('form-error');
    if (box && keys.length) {
      box.hidden = false;
      box.textContent = '还有 ' + keys.length + ' 项需要修正：' + keys.map(function (key) {
        return ERROR_LABELS[key] || schema.fieldLabel(key);
      }).join('、');
    }
  }

  function saveCustomerForm() {
    var form = document.getElementById('customer-form');
    if (!form) return null;
    var raw = collectFormValues();
    clearFormErrors(form);

    var errors = {};
    if (!schema.isValidISO(raw.firstInquiryAt)) errors.firstInquiryAt = '请选择首次询盘日期';
    if (!raw.customerName) errors.customerName = '请填写客户名称';
    else if (raw.customerName.length > 120) errors.customerName = '客户名称请不要超过 120 字';

    var partial = { firstInquiryAt: raw.firstInquiryAt, customerName: raw.customerName, followupStatus: raw.followupStatus, dealStatus: raw.dealStatus };
    var statusCheck = schema.validate(partial);
    Object.keys(statusCheck.errors).forEach(function (key) { if (!errors[key]) errors[key] = statusCheck.errors[key]; });

    raw.followups.forEach(function (item, index) {
      if (!item.date) errors.followupDate = '第 ' + (index + 1) + ' 条跟进缺少日期';
      else if (!schema.isValidISO(schema.parseLooseDate(item.date))) errors.followupDate = '第 ' + (index + 1) + ' 条跟进日期格式不正确';
    });
    if (raw.followups.length > 3) errors.followupDate = '最多记录最近三次跟进';

    if (Object.keys(errors).length) {
      showFormErrors(form, errors);
      return null;
    }

    var record = schema.normalize(raw);
    var mode = raw.__mode;
    var id = raw.__id;
    var saved;
    if (mode === 'edit' && id && Store.get(id)) {
      saved = Store.update(id, record);
      core.toast('客户资料已更新', 'ok');
    } else {
      saved = Store.add(record);
      core.toast('已新增客户「' + (saved.customerName || saved.companyName || '未命名') + '」', 'ok');
    }
    core.modal.close('customer-modal');
    render();
    return saved;
  }

  function addFollowupRow() {
    var editor = document.getElementById('followup-editor');
    if (!editor) return;
    var rows = core.qsa('[data-followup-row]', editor);
    var emptyTip = editor.querySelector('.followup-empty');
    if (emptyTip) emptyTip.parentNode.removeChild(emptyTip);
    if (rows.length >= 3) {
      core.toast('最多保留最近三次跟进记录', 'info');
      return;
    }
    var holder = document.createElement('div');
    holder.innerHTML = followupRowHtml({ date: schema.todayISO(), note: '' }, rows.length);
    editor.insertBefore(holder.firstChild, editor.firstChild);
  }

  function removeFollowupRow(button) {
    var editor = document.getElementById('followup-editor');
    var row = button ? button.closest('[data-followup-row]') : null;
    if (!editor || !row) return;
    if (row.parentNode) row.parentNode.removeChild(row);
    if (!core.qsa('[data-followup-row]', editor).length) {
      editor.innerHTML = '<p class="followup-empty">还没有跟进记录，点下方按钮补充第一次跟进。</p>';
    }
  }

  /* ================================================================== */
  /* 4. 跟进快捷弹窗                                                      */
  /* ================================================================== */

  function openFollowupModal(id) {
    var rec = Store.get(id);
    if (!rec) { core.toast('客户不存在或已被删除', 'error'); return null; }
    var today = schema.todayISO();
    var body = '<div class="followup-customer">' +
        '<div class="cell-strong">' + esc(rec.customerName || '未命名客户') + '</div>' +
        '<div class="cell-muted">' + esc(rec.companyName || '—') + ' · ' + esc(rec.country || '未填写国家') + '</div>' +
      '</div>' +
      '<div class="followup-history">当前共 ' + rec.followups.length + ' 条跟进记录，新增后只保留最近三次。</div>' +
      '<div class="field">' +
        '<label class="field-label" for="followup-date">跟进日期<span class="req" aria-hidden="true">*</span></label>' +
        '<input type="date" class="input" id="followup-date" value="' + today + '">' +
      '</div>' +
      '<div class="field">' +
        '<label class="field-label" for="followup-note">跟进内容<span class="req" aria-hidden="true">*</span></label>' +
        '<textarea class="textarea" id="followup-note" rows="3" placeholder="例如：电话沟通，客户确认 2000pcs 试单意向"></textarea>' +
        '<p class="field-error" id="followup-error" hidden></p>' +
      '</div>';

    var handle = core.modal.open({
      id: 'followup-modal',
      title: '记录一次跟进',
      subtitle: '保存后自动重算「距上次跟进」',
      body: body,
      footer: '<button type="button" class="btn btn--ghost" data-modal-close>取消</button>' +
        '<button type="button" class="btn btn--primary" data-action="save-followup" data-id="' + esc(id) + '">' + icon('check', 16) + ' 保存跟进</button>',
      onMount: function (container) {
        var note = container.querySelector('#followup-note');
        if (note) setTimeout(function () { try { note.focus(); } catch (err) { /* ignore */ } }, 60);
      }
    });
    return handle;
  }

  function saveFollowup(id) {
    var dateEl = document.getElementById('followup-date');
    var noteEl = document.getElementById('followup-note');
    var errorEl = document.getElementById('followup-error');
    var date = dateEl ? dateEl.value : '';
    var note = noteEl ? String(noteEl.value || '').trim() : '';
    var error = '';
    if (!schema.isValidISO(date)) error = '请选择跟进日期';
    else if (!note) error = '请填写跟进内容';
    if (error) {
      if (errorEl) { errorEl.textContent = error; errorEl.hidden = false; }
      core.toast(error, 'error');
      return false;
    }
    var rec = Store.get(id);
    if (!rec) { core.toast('客户不存在', 'error'); return false; }
    var merged = [{ date: date, note: note }].concat(rec.followups || []);
    Store.update(id, { followups: merged });
    core.modal.close('followup-modal');
    core.toast('已记录跟进（保留最近三次）', 'ok');
    render();
    return true;
  }

  /* ================================================================== */
  /* 5. 筛选与排序                                                        */
  /* ================================================================== */

  function activeFilterCount() {
    var f = state.filters;
    var count = 0;
    if (f.q.trim()) count += 1;
    if (f.country) count += 1;
    if (f.status) count += 1;
    if (f.deal) count += 1;
    return count;
  }

  function matchRecord(rec, filter) {
    if (filter.country && rec.country !== filter.country) return false;
    if (filter.status && rec.followupStatus !== filter.status) return false;
    if (filter.deal && rec.dealStatus !== filter.deal) return false;
    var q = filter.q.trim().toLowerCase();
    if (!q) return true;
    var haystack = [
      rec.customerName, rec.companyName, rec.product, rec.country, rec.coreReason,
      rec.followupStatus, rec.dealStatus, rec.firstInquiryAt, rec.foundedAt,
      rec.teamSize, rec.annualRevenue
    ].concat((rec.followups || []).map(function (item) { return item.note; })).join(' ').toLowerCase();
    return haystack.indexOf(q) !== -1;
  }

  function sortRecords(list) {
    var sort = state.filters.sort;
    var arr = list.slice();
    if (sort === 'firstInquiryAt-desc') arr.sort(function (a, b) { return (b.firstInquiryAt || '').localeCompare(a.firstInquiryAt || ''); });
    else if (sort === 'firstInquiryAt-asc') arr.sort(function (a, b) { return (a.firstInquiryAt || '').localeCompare(b.firstInquiryAt || ''); });
    else if (sort === 'customerName-asc') arr.sort(function (a, b) { return String(a.customerName || '').localeCompare(String(b.customerName || ''), 'zh-Hans-CN'); });
    else if (sort === 'daysSince-desc') arr.sort(function (a, b) {
      if (a.daysSince === null && b.daysSince === null) return 0;
      if (a.daysSince === null) return 1;
      if (b.daysSince === null) return -1;
      return b.daysSince - a.daysSince;
    });
    else if (sort === 'daysSince-asc') arr.sort(function (a, b) {
      if (a.daysSince === null && b.daysSince === null) return 0;
      if (a.daysSince === null) return 1;
      if (b.daysSince === null) return -1;
      return a.daysSince - b.daysSince;
    });
    else arr.sort(function (a, b) { return (b.updatedAt || '').localeCompare(a.updatedAt || ''); });
    return arr;
  }

  function filteredDerived() {
    return sortRecords(Store.deriveAll().filter(function (rec) { return matchRecord(rec, state.filters); }));
  }

  function allDerived() {
    return Store.deriveAll();
  }

  /* ================================================================== */
  /* 6. 仪表盘                                                            */
  /* ================================================================== */

  function statCard(options) {
    var tone = options.tone ? ' stat-card--' + options.tone : '';
    return '<article class="stat-card' + tone + '">' +
      '<span class="stat-icon" aria-hidden="true">' + icon(options.icon || 'chart', 20) + '</span>' +
      '<div class="stat-body">' +
        '<span class="stat-label">' + esc(options.label) + '</span>' +
        '<strong class="stat-value"' + (options.testid ? ' data-testid="' + options.testid + '"' : '') + '>' + esc(String(options.value)) + '</strong>' +
        '<span class="stat-hint">' + esc(options.hint || '') + '</span>' +
      '</div>' +
      '</article>';
  }

  function barList(items) {
    if (!items.length) return '<p class="hint">暂无数据</p>';
    var max = items.reduce(function (acc, item) { return Math.max(acc, item.count); }, 0) || 1;
    return '<div class="bar-list">' + items.map(function (item) {
      var tone = item.tone && item.tone !== 'neutral' && item.tone !== 'brand' ? ' bar-fill--' + item.tone : '';
      return '<div class="bar-row">' +
        '<div class="bar-meta"><span>' + (item.flag ? '<span class="flag" aria-hidden="true">' + item.flag + '</span> ' : '') + esc(item.name) + '</span><b>' + item.count + '</b></div>' +
        '<div class="bar-track"><span class="bar-fill' + tone + '" style="' + core.styleWidth(item.count / max * 100) + '"></span></div>' +
        '</div>';
    }).join('') + '</div>';
  }

  function renderDashboard() {
    var stats = Store.stats();
    var monthPrefix = schema.todayISO().slice(0, 7);
    var newThisMonth = allDerived().filter(function (rec) {
      return rec.firstInquiryAt.slice(0, 7) === monthPrefix;
    }).length;

    var countryItems = stats.byCountry.slice(0, 6).map(function (item) {
      return { name: item.name, count: item.count, flag: schema.countryFlag(item.name) };
    });
    var statusItems = stats.byStatus.map(function (item) {
      return { name: item.status, count: item.count, tone: item.tone };
    });

    var timeline = stats.recent.length ? '<ol class="timeline">' + stats.recent.map(function (item) {
      return '<li class="timeline-item">' +
        '<span class="timeline-dot" aria-hidden="true"></span>' +
        '<div class="timeline-body">' +
          '<div class="timeline-date">' + esc(item.date) + ' · ' + esc(core.dayLabel(item.daysAgo)) + ' ' + badge(item.followupStatus, schema.statusTone(item.followupStatus)) + '</div>' +
          '<button type="button" class="timeline-text" data-action="edit-customer" data-id="' + esc(item.id) + '">' +
            '<b>' + esc(item.customerName || '未命名客户') + '</b>' + (item.companyName ? ' · ' + esc(item.companyName) : '') + '：' + esc(item.note || '（无备注）') +
          '</button>' +
        '</div>' +
        '</li>';
    }).join('') + '</ol>' : '<p class="hint">最近还没有跟进记录，去客户台账记一次跟进吧。</p>';

    var overdue = stats.overdueList.length ? '<ul class="priority-list">' + stats.overdueList.map(function (rec) {
      return '<li class="priority-item">' +
        '<div class="priority-main">' +
          '<b>' + esc(rec.customerName || '未命名客户') + '</b>' +
          '<span class="priority-meta">' + esc(rec.companyName || '—') + ' · ' + esc(rec.country || '未填写') + ' · ' + badge(rec.followupStatus, rec.statusTone) + '</span>' +
        '</div>' +
        '<span class="age age--' + esc(rec.ageTone) + '">' + esc(core.dayLabel(rec.daysSince)) + '</span>' +
        '<button type="button" class="btn btn--sm btn--ghost" data-action="add-followup" data-id="' + esc(rec.id) + '" data-testid="btn-add-followup">' + icon('plus', 14) + ' 记跟进</button>' +
        '</li>';
    }).join('') + '</ul>' : '<p class="hint">没有超期未跟进的客户，节奏很棒。</p>';

    return '<div class="page">' +
      '<div class="page-head">' +
        '<div class="page-head-main">' +
          '<h2 class="page-title">客户经营概览</h2>' +
          '<p class="page-sub">共 ' + stats.total + ' 位客户 · 已成交 ' + stats.deal + ' 位 · 待跟进 ' + stats.open + ' 位 · 超期未跟进 ' + stats.overdue + ' 位（≥' + stats.staleDays + ' 天）</p>' +
        '</div>' +
        '<div class="page-head-actions">' +
          '<button type="button" class="btn btn--ghost" data-action="goto-recognize">' + icon('sparkles', 16) + ' 智能识别</button>' +
          '<button type="button" class="btn btn--ghost" data-action="export-excel">' + icon('download', 16) + ' 导出 Excel</button>' +
          '<button type="button" class="btn btn--primary" data-action="new-customer" data-testid="new-customer">' + icon('plus', 16) + ' 新建客户</button>' +
        '</div>' +
      '</div>' +

      '<div class="stats-grid">' +
        statCard({ tone: '', icon: 'users', label: '客户总数', value: stats.total, hint: '本月新询盘 ' + newThisMonth + ' 位', testid: 'stat-total' }) +
        statCard({ tone: 'ok', icon: 'check', label: '已成交', value: stats.deal, hint: '成交率 ' + core.percent(stats.deal, stats.total), testid: 'stat-deal' }) +
        statCard({ tone: 'warn', icon: 'clock', label: '待跟进', value: stats.open, hint: '距上次跟进 ≥' + stats.warnDays + ' 天', testid: 'stat-open' }) +
        statCard({ tone: 'danger', icon: 'alert', label: '超期未跟进', value: stats.overdue, hint: '≥' + stats.staleDays + ' 天未联系 · 优先处理', testid: 'stat-overdue' }) +
      '</div>' +

      '<div class="page-grid page-grid--2">' +
        '<section class="card">' +
          '<div class="card-head"><div><h3 class="card-title">客户国家分布</h3><p class="card-sub">按客户数量排序，最多显示 6 个</p></div><div class="card-actions">' + icon('globe', 18) + '</div></div>' +
          '<div class="card-body">' + barList(countryItems) + '</div>' +
        '</section>' +
        '<section class="card">' +
          '<div class="card-head"><div><h3 class="card-title">跟进状态分布</h3><p class="card-sub">平均距上次跟进 ' + (stats.avgDaysSince === null ? '—' : stats.avgDaysSince + ' 天') + '</p></div><div class="card-actions">' + icon('chart', 18) + '</div></div>' +
          '<div class="card-body">' + barList(statusItems) + '</div>' +
        '</section>' +
      '</div>' +

      '<div class="page-grid page-grid--2">' +
        '<section class="card">' +
          '<div class="card-head"><div><h3 class="card-title">最近跟进动态</h3><p class="card-sub">最新 8 条跟进记录</p></div><div class="card-actions"><button type="button" class="btn btn--sm btn--ghost" data-action="goto-customers">全部客户 ' + icon('arrowRight', 14) + '</button></div></div>' +
          '<div class="card-body">' + timeline + '</div>' +
        '</section>' +
        '<section class="card">' +
          '<div class="card-head"><div><h3 class="card-title">优先跟进</h3><p class="card-sub">超过 ' + stats.staleDays + ' 天未联系，建议本周触达</p></div><div class="card-actions">' + icon('target', 18) + '</div></div>' +
          '<div class="card-body">' + overdue + '</div>' +
        '</section>' +
      '</div>' +
      '</div>';
  }

  /* ================================================================== */
  /* 7. 客户台账                                                          */
  /* ================================================================== */

  function sortButton(key, label) {
    var sort = state.filters.sort;
    var active = sort === key + '-asc' || sort === key + '-desc';
    var dir = sort === key + '-asc' ? 'asc' : sort === key + '-desc' ? 'desc' : '';
    var arrow = dir === 'asc' ? '▲' : dir === 'desc' ? '▼' : '↕';
    return '<button type="button" class="sort-btn' + (active ? ' is-active' : '') + '" data-action="sort" data-sort="' + key + '">' + esc(label) +
      ' <span class="sort-arrow" aria-hidden="true">' + arrow + '</span></button>';
  }

  function customerRowHtml(rec) {
    var detailId = 'detail-' + rec.id;
    var reason = rec.coreReason || '';
    return '<tr class="row' + (state.expanded[rec.id] ? ' row--expanded' : '') + '" data-id="' + esc(rec.id) + '" data-testid="customer-row">' +
        '<td class="col-expand"><button type="button" class="row-toggle icon-btn icon-btn--sm" data-action="toggle-row" data-id="' + esc(rec.id) + '" data-testid="row-toggle" aria-expanded="' + (state.expanded[rec.id] ? 'true' : 'false') + '" aria-controls="' + detailId + '" aria-label="展开最近三次跟进">' + icon('chevronRight', 14) + '</button></td>' +
        '<td data-col="firstInquiryAt" class="cell-num">' + esc(rec.firstInquiryAt) + '</td>' +
        '<td data-col="age">' + ageCell(rec) + '</td>' +
        '<td data-col="customerName" class="cell-strong">' + esc(rec.customerName || '未命名') + '</td>' +
        '<td data-col="product" class="cell-muted" title="' + esc(rec.product) + '">' + esc(rec.product || '—') + '</td>' +
        '<td data-col="country" class="cell-country">' + flagText(rec.country) + esc(rec.country || '—') + '</td>' +
        '<td data-col="companyName" title="' + esc(rec.companyName) + '">' + esc(rec.companyName || '—') + '</td>' +
        '<td data-col="foundedAt" class="cell-num">' + esc(rec.foundedAt || '—') + '</td>' +
        '<td data-col="teamSize">' + esc(rec.teamSize || '—') + '</td>' +
        '<td data-col="annualRevenue">' + esc(rec.annualRevenue || '—') + '</td>' +
        '<td data-col="followupStatus">' + badge(rec.followupStatus, rec.statusTone) + '</td>' +
        '<td data-col="dealStatus">' + badge(rec.dealStatus, rec.dealTone) + '</td>' +
        '<td data-col="coreReason" class="cell-reason" title="' + esc(reason) + '">' + esc(reason || '—') + '</td>' +
        '<td class="col-actions"><div class="row-actions">' +
          '<button type="button" class="icon-btn icon-btn--sm" data-action="add-followup" data-id="' + esc(rec.id) + '" data-testid="btn-add-followup" title="记录跟进" aria-label="记录跟进">' + icon('plus', 15) + '</button>' +
          '<button type="button" class="icon-btn icon-btn--sm" data-action="edit-customer" data-id="' + esc(rec.id) + '" data-testid="btn-edit-customer" title="编辑客户" aria-label="编辑客户">' + icon('edit', 15) + '</button>' +
          '<button type="button" class="icon-btn icon-btn--sm icon-btn--danger" data-action="delete-customer" data-id="' + esc(rec.id) + '" data-testid="btn-delete-customer" title="删除客户" aria-label="删除客户">' + icon('trash', 15) + '</button>' +
        '</div></td>' +
      '</tr>' +
      '<tr class="row-detail" id="' + detailId + '" data-detail="' + esc(rec.id) + '"' + (state.expanded[rec.id] ? '' : ' hidden') + '>' +
        '<td colspan="14">' +
          '<div class="row-detail-inner">' +
            '<div class="row-detail-meta">' +
              '<span><i>成立时间</i>' + esc(rec.foundedAt || '—') + '</span>' +
              '<span><i>团队规模</i>' + esc(rec.teamSize || '—') + '</span>' +
              '<span><i>年营收</i>' + esc(rec.annualRevenue || '—') + '</span>' +
              '<span><i>首次询盘</i>' + esc(rec.firstInquiryAt) + '（' + (rec.inquiryDays === null ? '—' : '已跟进 ' + rec.inquiryDays + ' 天') + '）</span>' +
              '<span><i>最近更新</i>' + esc(schema.formatDateTime(rec.updatedAt)) + '</span>' +
            '</div>' +
            (rec.followups.length
              ? '<div class="followup-strip">' + rec.followups.map(function (item, index) {
                  return '<div class="followup-chip">' +
                    '<span class="followup-date">' + (index === 0 ? '最近 · ' : '第 ' + (index + 1) + ' 次 · ') + esc(item.date) + '</span>' +
                    '<span class="followup-note">' + esc(item.note || '（无备注）') + '</span>' +
                    '</div>';
                }).join('') + '</div>'
              : '<p class="followup-empty">还没有跟进记录。点击「记录一次跟进」，系统会自动计算距上次跟进的天数。</p>') +
            '<div class="row-detail-actions">' +
              '<button type="button" class="btn btn--sm btn--ghost" data-action="add-followup" data-id="' + esc(rec.id) + '" data-testid="btn-add-followup">' + icon('plus', 14) + ' 记录一次跟进</button>' +
              '<button type="button" class="btn btn--sm btn--ghost" data-action="edit-customer" data-id="' + esc(rec.id) + '">' + icon('edit', 14) + ' 编辑客户</button>' +
              '<span class="cell-muted">首次询盘 ' + esc(rec.firstInquiryAt) + ' · 已跟进 ' + (rec.inquiryDays === null ? '—' : rec.inquiryDays + ' 天') + ' · 创建于 ' + esc(schema.formatDateTime(rec.createdAt)) + '</span>' +
            '</div>' +
          '</div>' +
        '</td>' +
      '</tr>';
  }

  function countryFilterOptions() {
    var map = {};
    allDerived().forEach(function (rec) { if (rec.country) map[rec.country] = (map[rec.country] || 0) + 1; });
    return Object.keys(map).map(function (name) { return { name: name, count: map[name] }; })
      .sort(function (a, b) { return b.count - a.count || a.name.localeCompare(b.name, 'zh-Hans-CN'); });
  }

  function tableRegionHtml() {
    var total = Store.list().length;
    var records = filteredDerived();
    var f = state.filters;
    var isEmptyAll = total === 0;

    if (isEmptyAll) {
      return '<div class="card"><div class="card-body">' + emptyState({
        title: '还没有客户记录',
        text: '可以粘贴一段询盘文本让系统自动填表，或者直接手动新建第一位客户。',
        actions: '<button type="button" class="btn btn--primary" data-action="goto-recognize">' + icon('sparkles', 16) + ' 智能识别填表</button>' +
          '<button type="button" class="btn btn--ghost" data-action="new-customer">' + icon('plus', 16) + ' 新建客户</button>' +
          '<button type="button" class="btn btn--ghost" data-testid="btn-seed-demo" data-action="seed-demo">载入示例数据</button>'
      }) + '</div></div>';
    }

    if (!records.length) {
      return '<div class="card"><div class="card-body">' + emptyState({
        title: '没有匹配的客户',
        text: '试试更换关键词，或重置筛选条件（当前 ' + activeFilterCount() + ' 个筛选生效）。',
        actions: '<button type="button" class="btn btn--primary" data-action="reset-filters">重置筛选</button>'
      }) + '</div></div>' +
        '<div class="table-footer"><span class="table-count">共 ' + total + ' 条记录，筛选后 0 条</span></div>';
    }

    var dealCount = records.filter(function (rec) { return rec.isDeal; }).length;
    var openCount = records.filter(function (rec) { return rec.isOpen; }).length;

    return '<div class="table-wrap">' +
      '<table class="data-table data-table--customers">' +
        '<thead><tr>' +
          '<th class="col-expand" scope="col"><span class="sr-only">展开</span></th>' +
          '<th data-col="firstInquiryAt" scope="col">' + sortButton('firstInquiryAt', '首次询盘时间') + '</th>' +
          '<th data-col="age" scope="col">' + sortButton('daysSince', '距上次跟进') + '</th>' +
          '<th data-col="customerName" scope="col">' + sortButton('customerName', '客户名称') + '</th>' +
          '<th data-col="product" scope="col">咨询产品</th>' +
          '<th data-col="country" scope="col">国家</th>' +
          '<th data-col="companyName" scope="col">公司名</th>' +
          '<th data-col="foundedAt" scope="col">成立时间</th>' +
          '<th data-col="teamSize" scope="col">团队规模</th>' +
          '<th data-col="annualRevenue" scope="col">年营收</th>' +
          '<th data-col="followupStatus" scope="col">跟进状态</th>' +
          '<th data-col="dealStatus" scope="col">是否成交</th>' +
          '<th data-col="coreReason" scope="col">核心原因</th>' +
          '<th class="col-actions" scope="col">操作</th>' +
        '</tr></thead>' +
        '<tbody>' + records.map(customerRowHtml).join('') + '</tbody>' +
      '</table>' +
      '</div>' +
      '<div class="table-footer">' +
        '<span class="table-count">显示 <b>' + records.length + '</b> / 共 <b>' + total + '</b> 条 · 已成交 <b>' + dealCount + '</b> · 待跟进 <b>' + openCount + '</b>' +
          (activeFilterCount() ? ' · <span class="filter-on">已筛选 ' + activeFilterCount() + ' 项</span>' : '') + '</span>' +
        '<span class="cell-muted">点击行首箭头查看最近三次跟进 · 数据保存在本机浏览器</span>' +
      '</div>';
  }

  function renderCustomers() {
    var total = Store.list().length;
    var records = filteredDerived();
    var countries = countryFilterOptions();
    var f = state.filters;

    return '<div class="page">' +
      '<div class="page-head">' +
        '<div class="page-head-main">' +
          '<h2 class="page-title">客户台账</h2>' +
          '<p class="page-sub">共 ' + total + ' 位客户 · 当前显示 ' + records.length + ' 位 · 11 个核心字段 + 最近三次跟进</p>' +
        '</div>' +
        '<div class="page-head-actions">' +
          '<button type="button" class="btn btn--ghost" data-action="goto-recognize" data-testid="btn-recognize">' + icon('sparkles', 16) + ' 智能识别</button>' +
          '<button type="button" class="btn btn--ghost" data-action="export-json">' + icon('database', 16) + ' 备份 JSON</button>' +
          '<button type="button" class="btn btn--primary" data-action="export-excel" data-testid="btn-export-excel">' + icon('download', 16) + ' 一键导出 Excel</button>' +
        '</div>' +
      '</div>' +

      '<div class="toolbar">' +
        '<div class="toolbar-left">' +
          '<label class="search-box">' + icon('search', 16) +
            '<input type="search" class="input search-input" id="customer-search" data-testid="customer-search" placeholder="搜索客户 / 公司 / 产品 / 国家 / 原因 / 跟进内容…" value="' + esc(f.q) + '" aria-label="搜索客户">' +
          '</label>' +
          '<select class="select select--sm filter-select" id="filter-country" data-testid="filter-country" aria-label="按国家筛选">' +
            '<option value="">全部国家</option>' +
            countries.map(function (item) {
              return '<option value="' + esc(item.name) + '"' + (f.country === item.name ? ' selected' : '') + '>' + esc(item.name) + '（' + item.count + '）</option>';
            }).join('') +
          '</select>' +
          '<select class="select select--sm filter-select" id="filter-status" data-testid="filter-status" aria-label="按跟进状态筛选">' +
            '<option value="">全部状态</option>' +
            schema.FOLLOWUP_STATUS.map(function (status) {
              return '<option value="' + esc(status) + '"' + (f.status === status ? ' selected' : '') + '>' + esc(status) + '</option>';
            }).join('') +
          '</select>' +
          '<select class="select select--sm filter-select" id="filter-deal" data-testid="filter-deal" aria-label="按是否成交筛选">' +
            '<option value="">成交与否</option>' +
            schema.DEAL_STATUS.map(function (status) {
              return '<option value="' + esc(status) + '"' + (f.deal === status ? ' selected' : '') + '>' + esc(status) + '</option>';
            }).join('') +
          '</select>' +
          '<button type="button" class="btn btn--ghost btn--sm" data-action="reset-filters" data-testid="filter-reset">重置</button>' +
        '</div>' +
        '<div class="toolbar-right">' +
          '<span class="count-pill">显示 <b>' + records.length + '</b> / ' + total + ' 条</span>' +
          '<button type="button" class="btn btn--primary btn--sm" data-action="new-customer" data-testid="new-customer">' + icon('plus', 14) + ' 新建客户</button>' +
        '</div>' +
      '</div>' +

      '<div id="customers-table-region">' + tableRegionHtml() + '</div>' +
      '</div>';
  }

  function updateCustomersTable() {
    var region = document.getElementById('customers-table-region');
    if (!region) return;
    region.innerHTML = tableRegionHtml();
    var pill = document.querySelector('.count-pill');
    if (pill) pill.innerHTML = '显示 <b>' + filteredDerived().length + '</b> / ' + Store.list().length + ' 条';
  }

  function refreshTableSortIndicators() {
    core.qsa('.sort-btn').forEach(function (btn) {
      var key = btn.getAttribute('data-sort');
      var sort = state.filters.sort;
      btn.classList.toggle('is-active', sort.indexOf(key + '-') === 0);
      var arrow = btn.querySelector('.sort-arrow');
      if (arrow) {
        arrow.textContent = sort === key + '-asc' ? '▲' : sort === key + '-desc' ? '▼' : '↕';
      }
    });
  }

  /* ================================================================== */
  /* 8. 智能识别                                                          */
  /* ================================================================== */

  function parserReady() {
    return root.CrmParser && typeof root.CrmParser.recognize === 'function';
  }

  function sampleTexts() {
    if (parserReady() && Array.isArray(root.CrmParser.SAMPLE_TEXTS) && root.CrmParser.SAMPLE_TEXTS.length) {
      return root.CrmParser.SAMPLE_TEXTS;
    }
    return [{ label: '示例文本', text: '2026年9月12日收到德国客户 Michael Braun 的询盘，公司 Braun Trading GmbH 成立于1998年，团队 50-100 人，年营收约 1200 万欧元，咨询不锈钢保温杯 500ml，目前已报价。' }];
  }

  function recognizeResultHtml() {
    var result = state.recog.result;
    if (!result) {
      return '<div class="recog-empty">' +
        '<p class="panel-hint">粘贴询盘文本后点击「开始识别」，系统会把时间、客户、产品、国家、公司、规模、营收、状态等字段逐项填好。</p>' +
        '<ul class="help-list">' +
          '<li>支持中文、英文、中英混排</li>' +
          '<li>识别不到的字段会标灰，可手动补填后再入表</li>' +
          '<li>原文证据会显示在字段下方，方便核对</li>' +
        '</ul>' +
        '</div>';
    }

    var fields = schema.FIELDS.map(function (field) {
      var value = (result.fields && result.fields[field.key]) || '';
      var evidence = result.evidence && result.evidence[field.key] ? result.evidence[field.key].snippet : '';
      var filled = !!String(value).trim();
      return '<div class="recog-item ' + (filled ? 'recog-item--filled' : 'recog-item--missing') + '" data-recog-item="' + field.key + '">' +
        '<div class="recog-key">' + esc(field.label) + (field.required ? '<span class="req" aria-hidden="true">*</span>' : '') + '</div>' +
        '<input type="text" class="input recog-input" data-recog-key="' + field.key + '" value="' + esc(value) + '" placeholder="未识别，可手动补充" aria-label="' + esc(field.label) + '">' +
        (evidence ? '<div class="recog-evidence" title="' + esc(evidence) + '">原文：' + esc(evidence) + '</div>' : '') +
        '</div>';
    }).join('');

    var followups = result.followups && result.followups.length
      ? '<div class="recog-sub">识别到 ' + result.followups.length + ' 条跟进记录（保存时取最近三次）</div>' +
        '<div class="followup-strip">' + result.followups.map(function (item, index) {
          return '<div class="followup-chip"><span class="followup-date">#' + (index + 1) + ' · ' + esc(item.date) + '</span><span class="followup-note">' + esc(item.note || '（无描述）') + '</span></div>';
        }).join('') + '</div>'
      : '<div class="recog-sub">没有识别到独立的跟进记录，可保存后在「记跟进」中补充。</div>';

    var warnings = (result.warnings || []).map(function (warn) {
      return '<li>' + esc(warn.message || warn.code || '') + '</li>';
    }).join('');

    var filledCount = schema.FIELDS.filter(function (field) {
      return String((result.fields && result.fields[field.key]) || '').trim();
    }).length;

    return '<div class="recog-summary">' +
        '<span class="recog-count">已识别 <b>' + filledCount + '</b> / ' + schema.FIELDS.length + ' 项字段</span>' +
        '<span class="panel-hint">可直接修改下方内容，确认无误后填入表格</span>' +
      '</div>' +
      '<div class="recog-list">' + fields + '</div>' +
      followups +
      (warnings ? '<ul class="recog-warnings">' + warnings + '</ul>' : '') +
      '<div class="recog-actions">' +
        '<button type="button" class="btn btn--ghost" data-action="recognize-clear-result">清空结果</button>' +
        '<button type="button" class="btn btn--ghost" data-testid="btn-add-recognized" data-action="recognize-add">' + icon('plus', 16) + ' 直接新增一行</button>' +
        '<button type="button" class="btn btn--primary" data-testid="btn-fill-customer" data-action="recognize-fill">' + icon('arrowRight', 16) + ' 填入表格并核对</button>' +
      '</div>';
  }

  function renderRecognize() {
    var samples = sampleTexts();
    return '<div class="page page--recognize">' +
      '<div class="page-head">' +
        '<div class="page-head-main">' +
          '<h2 class="page-title">智能识别 · 一段话自动填表</h2>' +
          '<p class="page-sub">粘贴客户询盘邮件、微信聊天、展会名片速记，自动拆分出 11 个字段和最近三次跟进。</p>' +
        '</div>' +
        '<div class="page-head-actions">' +
          '<button type="button" class="btn btn--ghost" data-action="goto-customers">' + icon('users', 16) + ' 返回客户台账</button>' +
        '</div>' +
      '</div>' +
      '<div class="recog-layout">' +
        '<section class="recog-panel recog-panel--source">' +
          '<div class="panel-head">' +
            '<div class="panel-title">询盘文本</div>' +
            '<p class="panel-hint">中文 / 英文 / 中英混排都可以，越具体识别越准</p>' +
          '</div>' +
          '<div class="sample-row">' +
            '<span class="panel-hint">试试示例：</span>' +
            samples.map(function (sample, index) {
              return '<button type="button" class="sample-pill" data-action="recognize-sample" data-index="' + index + '">' + esc(sample.label || ('示例 ' + (index + 1))) + '</button>';
            }).join('') +
          '</div>' +
          '<textarea class="textarea textarea--lg" id="recognize-input" data-testid="recognize-input" placeholder="例如：2026年9月12日收到德国客户 Michael Braun 的询盘，公司 Braun Trading GmbH 成立于 1998 年，团队 50-100 人，年营收约 1200 万欧元，咨询 500ml 不锈钢保温杯，目前已报价……">' + esc(state.recog.text) + '</textarea>' +
          '<div class="recog-source-foot">' +
            '<span class="recog-count" id="recog-char-count">' + (state.recog.text ? state.recog.text.length : 0) + ' 字</span>' +
            '<div class="recog-actions">' +
              '<button type="button" class="btn btn--ghost btn--sm" data-action="recognize-clear">清空</button>' +
              '<button type="button" class="btn btn--primary" data-testid="btn-run-recognize" data-action="recognize-run">' + icon('sparkles', 16) + ' 开始识别</button>' +
            '</div>' +
          '</div>' +
        '</section>' +
        '<section class="recog-panel recog-panel--result" data-testid="recognize-result">' +
          '<div class="panel-head">' +
            '<div class="panel-title">识别结果</div>' +
            '<p class="panel-hint">逐项核对后再入表</p>' +
          '</div>' +
          '<div class="recog-result-body" id="recog-result-body">' + recognizeResultHtml() + '</div>' +
        '</section>' +
      '</div>' +
      '</div>';
  }

  function updateRecognizeResult() {
    var body = document.getElementById('recog-result-body');
    if (body) body.innerHTML = recognizeResultHtml();
  }

  function runRecognize() {
    if (!parserReady()) {
      core.toast('识别引擎未加载（js/parser.js 缺失），请刷新页面', 'error');
      return null;
    }
    var area = document.getElementById('recognize-input');
    var text = area ? area.value : state.recog.text;
    state.recog.text = text;
    if (!String(text || '').trim()) {
      core.toast('请先粘贴一段询盘文本', 'error');
      if (area) { try { area.focus(); } catch (err) { /* ignore */ } }
      return null;
    }
    var result;
    try {
      result = root.CrmParser.recognize(text, { today: schema.todayISO() });
    } catch (err) {
      core.toast('识别失败：' + (err && err.message ? err.message : '未知错误'), 'error');
      return null;
    }
    state.recog.result = result;
    updateRecognizeResult();
    var filled = schema.FIELDS.filter(function (field) {
      return String((result.fields && result.fields[field.key]) || '').trim();
    }).length;
    core.toast('识别完成：填好 ' + filled + ' / ' + schema.FIELDS.length + ' 项字段', filled >= 5 ? 'ok' : 'info');
    return result;
  }

  function readRecognizedValues() {
    var data = {};
    schema.FIELD_KEYS.forEach(function (key) { data[key] = ''; });
    core.qsa('[data-recog-key]').forEach(function (input) {
      data[input.getAttribute('data-recog-key')] = String(input.value || '').trim();
    });
    return data;
  }

  function applyRecognition(mode) {
    var result = state.recog.result;
    if (!result) {
      core.toast('请先运行一次识别', 'error');
      return null;
    }
    var values = readRecognizedValues();
    var record = schema.normalize(values);
    var followups = (result.followups || []).slice(0, 3);
    if (followups.length) record.followups = schema.normalize({ followups: followups }).followups;

    if (mode === 'add') {
      var errors = {};
      if (!record.customerName) errors.customerName = '识别结果里没有客户名称，请补充';
      if (!schema.isValidISO(record.firstInquiryAt)) errors.firstInquiryAt = '缺少询盘日期';
      if (Object.keys(errors).length) {
        core.toast(Object.keys(errors).map(function (key) { return errors[key]; }).join('；'), 'error');
        return null;
      }
      var saved = Store.add(record);
      core.toast('已新增客户「' + saved.customerName + '」，可在客户台账查看', 'ok', {
        action: { label: '去看看', onClick: function () { navigate('#/customers'); } }
      });
      return saved;
    }

    openCustomerModal(record, { mode: 'create', notice: '已根据识别结果自动填入 ' + schema.FIELDS.filter(function (field) {
      return String(values[field.key] || '').trim();
    }).length + ' 项字段，请核对后保存。' });
    return record;
  }

  /* ================================================================== */
  /* 9. 设置 / 帮助                                                       */
  /* ================================================================== */

  function renderSettings() {
    var settings = Store.settings();
    var stats = Store.stats();
    var storageOk = Store.storageAvailable();
    var lastSaved = Store.savedAt();
    return '<div class="page page--settings">' +
      '<div class="page-head">' +
        '<div class="page-head-main">' +
          '<h2 class="page-title">设置</h2>' +
          '<p class="page-sub">数据全部保存在本机浏览器，换电脑请先导出 JSON 备份。</p>' +
        '</div>' +
      '</div>' +
      '<div class="settings-list">' +
        '<section class="card">' +
          '<div class="card-head"><div><h3 class="card-title">数据管理</h3><p class="card-sub">备份、恢复与示例数据</p></div><div class="card-actions">' + icon('database', 18) + '</div></div>' +
          '<div class="card-body">' +
            settingRow('导出 JSON 备份', '把全部客户、设置导出为一个 JSON 文件，可直接保存到网盘或仓库。', '<button type="button" class="btn btn--ghost" data-action="export-json">' + icon('download', 16) + ' 导出备份</button>') +
            settingRow('从 JSON 恢复', '选择此前导出的备份文件，合并进当前数据（同 id 覆盖）。', '<button type="button" class="btn btn--ghost" data-action="import-json">' + icon('upload', 16) + ' 选择文件</button>') +
            settingRow('载入示例数据', '写入 12 位演示客户，用于体验表格、仪表盘与导出效果。', '<button type="button" class="btn btn--ghost" data-testid="btn-seed-demo" data-action="seed-demo">载入示例数据</button>') +
            '<div class="danger-zone">' +
              settingRow('清空全部数据', '删除本机保存的所有客户记录，此操作不可撤销，建议先导出备份。', '<button type="button" class="btn btn--danger" data-testid="btn-clear-all" data-action="clear-all">清空数据</button>') +
            '</div>' +
          '</div>' +
        '</section>' +
        '<section class="card">' +
          '<div class="card-head"><div><h3 class="card-title">跟进提醒阈值</h3><p class="card-sub">影响「距上次跟进」的颜色与仪表盘统计</p></div><div class="card-actions">' + icon('clock', 18) + '</div></div>' +
          '<div class="card-body">' +
            '<div class="form-grid">' +
              '<div class="field"><label class="field-label" for="setting-warnDays">待跟进阈值（天）</label>' +
                '<input type="number" class="input" id="setting-warnDays" min="1" max="180" step="1" value="' + settings.warnDays + '">' +
                '<p class="field-hint">超过该天数显示琥珀色提醒，默认 7 天</p></div>' +
              '<div class="field"><label class="field-label" for="setting-staleDays">超期阈值（天）</label>' +
                '<input type="number" class="input" id="setting-staleDays" min="1" max="365" step="1" value="' + settings.staleDays + '">' +
                '<p class="field-hint">超过该天数显示红色并计入「超期未跟进」，默认 14 天</p></div>' +
            '</div>' +
            '<div class="data-actions"><button type="button" class="btn btn--primary btn--sm" data-action="save-settings">' + icon('check', 16) + ' 保存设置</button></div>' +
          '</div>' +
        '</section>' +
        '<section class="card">' +
          '<div class="card-head"><div><h3 class="card-title">存储信息</h3><p class="card-sub">localStorage · 数据不出电脑</p></div><div class="card-actions">' + icon('file', 18) + '</div></div>' +
          '<div class="card-body">' +
            '<dl class="info-list">' +
              '<div><dt>客户记录</dt><dd>' + stats.total + ' 位（已成交 ' + stats.deal + ' 位）</dd></div>' +
              '<div><dt>存储位置</dt><dd>浏览器 localStorage · 键名 <code class="mono">' + esc(Store.KEY) + '</code></dd></div>' +
              '<div><dt>可用状态</dt><dd>' + (storageOk ? '<span class="badge badge--ok">可用</span>' : '<span class="badge badge--danger">不可用</span>') + '</dd></div>' +
              '<div><dt>最后保存</dt><dd>' + esc(schema.formatDateTime(lastSaved)) + '</dd></div>' +
              '<div><dt>系统版本</dt><dd>客户通 CustomerCRM v' + esc(schema.VERSION) + '</dd></div>' +
            '</dl>' +
          '</div>' +
        '</section>' +
      '</div>' +
      '</div>';
  }

  function settingRow(title, desc, action) {
    return '<div class="setting-row">' +
      '<div class="setting-info">' +
        '<div class="setting-title">' + esc(title) + '</div>' +
        '<p class="setting-desc">' + esc(desc) + '</p>' +
      '</div>' +
      '<div class="setting-actions">' + action + '</div>' +
      '</div>';
  }

  function renderHelp() {
    var fieldRows = schema.FIELDS.map(function (field) {
      return '<tr><td class="cell-strong">' + esc(field.label) + (field.required ? '<span class="req" aria-hidden="true">*</span>' : '') + '</td>' +
        '<td>' + esc(field.hint || '') + '</td>' +
        '<td class="cell-muted">' + esc(field.placeholder || '—') + '</td></tr>';
    }).join('');

    return '<div class="page page--help">' +
      '<div class="page-head"><div class="page-head-main">' +
        '<h2 class="page-title">使用说明</h2>' +
        '<p class="page-sub">三分钟上手：录入客户 → 记录跟进 → 导出 Excel；也可以直接粘贴一段询盘让系统自动填表。</p>' +
      '</div></div>' +
      '<div class="help-grid">' +
        '<section class="help-card">' +
          '<h3>① 三种录入方式</h3>' +
          '<ol class="help-list">' +
            '<li><b>智能识别</b>：把询盘邮件 / 聊天记录整段粘进「智能识别」，点「开始识别」，核对后「填入表格并核对」。</li>' +
            '<li><b>手动新建</b>：点右上角「新建客户」，必填「首次询盘时间」和「客户名称」。</li>' +
            '<li><b>导入备份</b>：设置页「从 JSON 恢复」，可把同事的台账合并进来。</li>' +
          '</ol>' +
        '</section>' +
        '<section class="help-card">' +
          '<h3>② 跟进与「距上次跟进」</h3>' +
          '<ul class="help-list">' +
            '<li>每位客户最多记录 <b>最近三次</b>跟进，最新一条排最前。</li>' +
            '<li>「距上次跟进」= 今天 − 最近一条跟进日期；没有跟进记录时显示「尚未跟进」。</li>' +
            '<li>颜色：<span class="badge badge--ok">7 天内</span> <span class="badge badge--warn">7-13 天</span> <span class="badge badge--danger">≥14 天</span>，阈值可在设置页调整。</li>' +
            '<li>表格行首箭头展开，可看到三条跟进记录并直接补记。</li>' +
          '</ul>' +
        '</section>' +
        '<section class="help-card">' +
          '<h3>③ 一键导出 Excel</h3>' +
          '<ul class="help-list">' +
            '<li>客户台账页「一键导出 Excel」：有筛选时导出<b>当前筛选结果</b>，否则导出全部。</li>' +
            '<li>文件含两个工作表：<b>客户总表</b>（19 列，含最近三次跟进与距上次跟进天数）与<b>跟进明细</b>（每条跟进一行）。</li>' +
            '<li>表头冻结、自动筛选、列宽与换行都已设置，可直接打印或发给同事。</li>' +
          '</ul>' +
        '</section>' +
        '<section class="help-card">' +
          '<h3>④ 数据安全与迁移</h3>' +
          '<ul class="help-list">' +
            '<li>数据只存在本机浏览器，不上传任何服务器；清除浏览器数据会丢失，请定期导出 JSON 备份。</li>' +
            '<li>换电脑：旧电脑导出 JSON → 新电脑「从 JSON 恢复」。</li>' +
            '<li>本系统是纯静态网页，可部署到 GitHub Pages / 内网静态服务器。</li>' +
          '</ul>' +
        '</section>' +
      '</div>' +
      '<section class="card">' +
        '<div class="card-head"><div><h3 class="card-title">字段说明</h3><p class="card-sub">表格、表单与 Excel 共用同一套字段</p></div></div>' +
        '<div class="card-body">' +
          '<div class="table-wrap"><table class="data-table"><thead><tr><th>字段</th><th>说明</th><th>示例</th></tr></thead><tbody>' + fieldRows + '</tbody></table></div>' +
        '</div>' +
      '</section>' +
      '<section class="card">' +
        '<div class="card-head"><div><h3 class="card-title">智能识别怎么写更容易识别</h3><p class="card-sub">规则引擎：标签越清晰、日期越规范，命中率越高</p></div></div>' +
        '<div class="card-body">' +
          '<ul class="help-list help-list--wide">' +
            '<li>带上 <b>日期 + 客户名 + 公司名 + 国家</b> 四要素，例如「2026年9月12日收到德国客户 Michael Braun 的询盘，公司 Braun Trading GmbH……」。</li>' +
            '<li>用「成立于」「团队」「年营收」这类标签词描述公司背景，识别成立时间 / 团队规模 / 年营收最稳。</li>' +
            '<li>跟进记录写成「日期 + 做了什么」，例如「11月18日邮件确认样品尺寸」，会被归入最近三次跟进。</li>' +
            '<li>识别不出的字段会在结果里标灰，可手动补填；所有字段都能在保存前修改。</li>' +
            '<li>快捷键：<span class="kbd">/</span> 聚焦搜索，<span class="kbd">N</span> 新建客户，<span class="kbd">Esc</span> 关闭弹窗。</li>' +
          '</ul>' +
        '</div>' +
      '</section>' +
      '</div>';
  }

  /* ================================================================== */
  /* 10. 页面外壳与渲染                                                   */
  /* ================================================================== */

  function updateChrome(routeName) {
    var meta = PAGE_META[routeName] || PAGE_META.dashboard;
    var crumbs = document.getElementById('crumbs');
    var title = document.getElementById('page-title');
    if (crumbs) crumbs.textContent = meta.crumb;
    if (title) title.textContent = meta.title;
    document.title = meta.title + ' · 客户通 CustomerCRM';
    core.qsa('.nav-item').forEach(function (item) {
      item.classList.toggle('is-active', item.getAttribute('data-nav') === routeName);
    });
    var topbarNew = document.querySelector('.btn-new-customer');
    if (topbarNew) topbarNew.style.display = routeName === 'customers' || routeName === 'dashboard' ? 'none' : '';
  }

  function render() {
    var route = parseRoute();
    state.route = route.name;
    updateChrome(route.name);
    var viewInner = document.getElementById('view-inner');
    if (!viewInner) return;
    var html;
    if (route.name === 'dashboard') html = renderDashboard();
    else if (route.name === 'customers') html = renderCustomers();
    else if (route.name === 'recognize') html = renderRecognize();
    else if (route.name === 'settings') html = renderSettings();
    else html = renderHelp();
    viewInner.innerHTML = html;
    closeSidebar();
    if (route.name === 'recognize') {
      var area = document.getElementById('recognize-input');
      if (area) {
        area.addEventListener('input', function () {
          state.recog.text = area.value;
          var counter = document.getElementById('recog-char-count');
          if (counter) counter.textContent = area.value.length + ' 字';
        });
      }
    }
  }

  function saveStateIndicator(text, stateValue) {
    var wrap = document.getElementById('save-state');
    var label = document.getElementById('save-state-text');
    if (!wrap || !label) return;
    wrap.setAttribute('data-state', stateValue || 'saved');
    label.textContent = text || '已保存';
  }

  function onStoreChange(event) {
    var now = new Date();
    var hh = String(now.getHours()).padStart(2, '0');
    var mm = String(now.getMinutes()).padStart(2, '0');
    saveStateIndicator('已保存 ' + hh + ':' + mm, 'saved');
    if (event && event.type === 'settings') {
      core.toast('设置已保存', 'ok');
    }
    if (state.route === 'recognize' && document.getElementById('recog-result-body')) {
      // 识别页不因数据变化重绘（避免丢失输入），仅更新保存指示
      return;
    }
    if (!core.modal.isOpen()) render();
  }

  /* ================================================================== */
  /* 11. 交互分发                                                         */
  /* ================================================================== */

  function toggleRow(id) {
    var detail = document.querySelector('.row-detail[data-detail="' + id + '"]');
    var row = document.querySelector('.row[data-id="' + id + '"]');
    if (!detail || !row) return;
    var expanded = !detail.hidden;
    if (expanded) {
      detail.hidden = true;
      row.classList.remove('row--expanded');
    } else {
      detail.hidden = false;
      row.classList.add('row--expanded');
    }
    state.expanded[id] = !expanded;
    var toggle = row.querySelector('.row-toggle');
    if (toggle) {
      toggle.setAttribute('aria-expanded', String(!expanded));
      toggle.classList.toggle('is-open', !expanded);
    }
  }

  function deleteCustomer(id) {
    var rec = Store.get(id);
    if (!rec) { core.toast('客户不存在或已被删除', 'error'); return Promise.resolve(false); }
    return core.confirm({
      title: '删除客户',
      message: '确定删除「' + (rec.customerName || rec.companyName || '未命名客户') + '」？该客户的跟进记录会一起删除，且不可恢复。',
      confirmText: '删除',
      danger: true
    }).then(function (ok) {
      if (!ok) return false;
      Store.remove(id);
      delete state.expanded[id];
      core.toast('已删除客户', 'ok');
      return true;
    });
  }

  function exportExcel() {
    if (!root.CrmXlsx || typeof root.CrmXlsx.exportCustomers !== 'function') {
      core.toast('导出引擎未加载（js/xlsx.js 缺失），请刷新页面', 'error');
      return { filename: '', bytes: 0 };
    }
    var records = state.filters && activeFilterCount() ? filteredDerived() : Store.list();
    if (!records.length) {
      core.toast('没有可导出的客户记录', 'error');
      return { filename: '', bytes: 0 };
    }
    try {
      var result = root.CrmXlsx.exportCustomers(records, { today: schema.todayISO() });
      if (!result || !result.bytes) {
        core.toast('导出失败：未生成文件', 'error');
        return { filename: '', bytes: 0 };
      }
      if (typeof root.CrmXlsx.download === 'function') {
        root.CrmXlsx.download(result.bytes, result.filename);
      } else if (result.bytes instanceof Blob) {
        core.download(result.filename, result.bytes);
      } else {
        core.download(result.filename, new Blob([result.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      }
      core.toast('已导出 ' + records.length + ' 位客户：' + result.filename, 'ok', { timeout: 5000 });
      return { filename: result.filename, bytes: result.bytes.length || (result.bytes.byteLength || 0) };
    } catch (err) {
      core.toast('导出失败：' + (err && err.message ? err.message : '未知错误'), 'error');
      return { filename: '', bytes: 0 };
    }
  }

  function exportJson() {
    var text = Store.exportJSON();
    var stamp = schema.todayISO().replace(/-/g, '') + '_' + String(new Date().getHours()).padStart(2, '0') + String(new Date().getMinutes()).padStart(2, '0');
    var filename = '客户通备份_' + stamp + '.json';
    var ok = core.saveTextFile(filename, text, 'application/json;charset=utf-8');
    core.toast(ok ? '已导出备份：' + filename : '备份导出失败', ok ? 'ok' : 'error');
    return { filename: filename, text: text };
  }

  function importJson() {
    core.pickFile('.json,application/json').then(function (file) {
      if (!file) return;
      var result = Store.importJSON(file.text, 'merge');
      if (!result.ok) {
        core.toast(result.error || '导入失败', 'error');
        return;
      }
      core.toast('已从备份导入 ' + result.imported + ' 位客户' + (result.skipped ? '（跳过 ' + result.skipped + ' 条无效记录）' : ''), 'ok');
    });
  }

  function saveSettings() {
    var warnEl = document.getElementById('setting-warnDays');
    var staleEl = document.getElementById('setting-staleDays');
    var warnDays = Number(warnEl && warnEl.value);
    var staleDays = Number(staleEl && staleEl.value);
    if (!isFinite(warnDays) || warnDays < 1 || warnDays > 180) { core.toast('待跟进阈值需在 1-180 天之间', 'error'); return null; }
    if (!isFinite(staleDays) || staleDays < warnDays || staleDays > 365) { core.toast('超期阈值需 ≥ 待跟进阈值，且不超过 365 天', 'error'); return null; }
    return Store.updateSettings({ warnDays: Math.round(warnDays), staleDays: Math.round(staleDays) });
  }

  function clearAll() {
    return core.confirm({
      title: '清空全部数据',
      message: '将删除本机保存的全部客户记录，操作不可恢复。建议先在设置页导出 JSON 备份。确定继续？',
      confirmText: '清空数据',
      danger: true
    }).then(function (ok) {
      if (!ok) return false;
      Store.clear();
      core.toast('已清空全部客户数据', 'ok');
      return true;
    });
  }

  function seedDemo() {
    var existing = Store.list().length;
    var proceed = existing
      ? core.confirm({
          title: '载入示例数据',
          message: '当前已有 ' + existing + ' 位客户，载入示例数据会覆盖现有数据（建议先在设置页导出 JSON 备份）。确定继续？',
          confirmText: '覆盖并载入',
          danger: true
        })
      : Promise.resolve(true);
    return proceed.then(function (ok) {
      if (!ok) return 0;
      var count = Store.seedDemo(true);
      core.toast('已载入 ' + count + ' 位示例客户', 'ok');
      return count;
    });
  }

  function handleAction(action, element) {
    var id = element.getAttribute('data-id');
    switch (action) {
      case 'toggle-sidebar': toggleSidebar(); break;
      case 'new-customer': openCustomerModal(null, { mode: 'create' }); break;
      case 'edit-customer': openCustomerModal({ id: id }, { mode: 'edit' }); break;
      case 'delete-customer': deleteCustomer(id); break;
      case 'toggle-row': toggleRow(id); break;
      case 'add-followup': openFollowupModal(id); break;
      case 'save-followup': saveFollowup(id); break;
      case 'add-followup-row': addFollowupRow(); break;
      case 'remove-followup-row': removeFollowupRow(element); break;
      case 'save-customer-form': saveCustomerForm(); break;
      case 'goto-customers': navigate('#/customers'); break;
      case 'goto-recognize': navigate('#/recognize'); break;
      case 'export-excel': exportExcel(); break;
      case 'export-json': exportJson(); break;
      case 'import-json': importJson(); break;
      case 'seed-demo': seedDemo(); break;
      case 'clear-all': clearAll(); break;
      case 'save-settings': saveSettings(); break;
      case 'reset-filters': resetFilters(); break;
      case 'sort': {
        var key = element.getAttribute('data-sort');
        state.filters.sort = nextSort(key);
        updateCustomersTable();
        refreshTableSortIndicators();
        break;
      }
      case 'recognize-sample': loadSample(Number(element.getAttribute('data-index'))); break;
      case 'recognize-run': runRecognize(); break;
      case 'recognize-clear': clearRecognizeInput(); break;
      case 'recognize-clear-result':
        state.recog.result = null;
        updateRecognizeResult();
        break;
      case 'recognize-fill': applyRecognition('fill'); break;
      case 'recognize-add': applyRecognition('add'); break;
      default: break;
    }
  }

  function nextSort(key) {
    var current = state.filters.sort;
    if (key === 'daysSince') return current === 'daysSince-desc' ? 'daysSince-asc' : 'daysSince-desc';
    if (current === key + '-desc') return key + '-asc';
    if (current === key + '-asc') return 'updated-desc';
    return key + '-desc';
  }

  function resetFilters() {
    state.filters = Object.assign({}, DEFAULT_FILTERS);
    if (state.route === 'customers') render();
  }

  function loadSample(index) {
    var samples = sampleTexts();
    var sample = samples[index];
    if (!sample) return;
    var area = document.getElementById('recognize-input');
    if (area) {
      area.value = sample.text;
      state.recog.text = sample.text;
      var counter = document.getElementById('recog-char-count');
      if (counter) counter.textContent = sample.text.length + ' 字';
    }
    runRecognize();
  }

  function clearRecognizeInput() {
    var area = document.getElementById('recognize-input');
    if (area) { area.value = ''; area.focus(); }
    state.recog.text = '';
    state.recog.result = null;
    var counter = document.getElementById('recog-char-count');
    if (counter) counter.textContent = '0 字';
    updateRecognizeResult();
  }

  /* ---------------- 侧栏 ---------------- */

  function toggleSidebar(force) {
    var sidebar = document.getElementById('sidebar');
    var backdrop = document.getElementById('sidebar-backdrop');
    var toggle = document.querySelector('.menu-toggle');
    if (!sidebar) return;
    var open = typeof force === 'boolean' ? force : !sidebar.classList.contains('is-open');
    sidebar.classList.toggle('is-open', open);
    if (backdrop) backdrop.classList.toggle('is-open', open);
    if (toggle) toggle.setAttribute('aria-expanded', String(open));
    document.body.classList.toggle('sidebar-open', open);
  }

  function closeSidebar() {
    if (document.getElementById('sidebar') && document.getElementById('sidebar').classList.contains('is-open')) toggleSidebar(false);
  }

  /* ---------------- 全局事件 ---------------- */

  function bindGlobalEvents() {
    core.on(document, 'click', function (event) {
      var target = event.target;
      var actionEl = target && target.closest ? target.closest('[data-action]') : null;
      if (actionEl && !actionEl.disabled) {
        event.preventDefault();
        handleAction(actionEl.getAttribute('data-action'), actionEl);
        return;
      }
      var navLink = target && target.closest ? target.closest('a[href^="#/"]') : null;
      if (navLink) { closeSidebar(); }
    });

    document.addEventListener('submit', function (event) {
      if (event.target && event.target.id === 'customer-form') {
        event.preventDefault();
        saveCustomerForm();
      }
    });

    document.addEventListener('keydown', function (event) {
      var active = document.activeElement;
      var typing = core.isTextInput(active);
      if (event.key === '/' && !typing && !core.modal.isOpen()) {
        event.preventDefault();
        if (state.route !== 'customers') {
          navigate('#/customers');
          setTimeout(function () {
            var input = document.getElementById('customer-search');
            if (input) input.focus();
          }, 60);
        } else {
          var input = document.getElementById('customer-search');
          if (input) input.focus();
        }
        return;
      }
      if ((event.key === 'n' || event.key === 'N') && !typing && !core.modal.isOpen() && !event.metaKey && !event.ctrlKey) {
        event.preventDefault();
        openCustomerModal(null, { mode: 'create' });
      }
    });

    var searchInput = null;
    core.on(document, 'input', function (event) {      var target = event.target;
      if (!target) return;
      if (target.id === 'customer-search') {
        state.filters.q = target.value;
        updateCustomersTable();
        return;
      }
      if (target.hasAttribute && target.hasAttribute('data-recog-key')) {
        var item = target.closest('.recog-item');
        if (item) {
          var filled = String(target.value || '').trim() !== '';
          item.classList.toggle('recog-item--filled', filled);
          item.classList.toggle('recog-item--missing', !filled);
        }
      }
    });

    core.on(document, 'change', function (event) {
      var target = event.target;
      if (!target) return;
      var map = { 'filter-country': 'country', 'filter-status': 'status', 'filter-deal': 'deal' };
      var key = map[target.id];
      if (!key) return;
      state.filters[key] = target.value;
      updateCustomersTable();
    });

    var backdrop = document.getElementById('sidebar-backdrop');
    if (backdrop) backdrop.addEventListener('click', function () { toggleSidebar(false); });

    root.addEventListener('hashchange', function () { render(); });
  }

  /* ================================================================== */
  /* 12. 初始化与对外接口                                                 */
  /* ================================================================== */

  function init() {
    var info = Store.init();
    unsubscribers.push(Store.subscribe(onStoreChange));
    bindGlobalEvents();

    if (!root.location.hash) {
      try { root.history.replaceState(null, '', '#/dashboard'); } catch (err) { /* file:// 下可能失败，忽略 */ }
    }
    render();
    saveStateIndicator(info.storageOk ? '已就绪' : '仅本次有效', info.storageOk ? 'saved' : 'warn');
    if (!info.storageOk) {
      core.toast('浏览器禁用了本地存储，本次修改只在当前页面有效，请及时导出备份', 'error', { timeout: 6000 });
    }
    var helpLink = document.querySelector('.topbar-help');
    if (helpLink) helpLink.setAttribute('aria-label', '使用说明');
  }

  root.CustomerCRM = {
    version: schema.VERSION,
    app: {
      navigate: navigate,
      getRoute: getRoute,
      refresh: render,
      newCustomer: function () { return openCustomerModal(null, { mode: 'create' }); },
      editCustomer: function (id) { return openCustomerModal({ id: id }, { mode: 'edit' }); },
      saveCustomerForm: saveCustomerForm,
      deleteCustomer: deleteCustomer,
      addFollowup: function (id, item) {
        var rec = Store.get(id);
        if (!rec) return null;
        var merged = [{ date: item.date, note: item.note }].concat(rec.followups || []);
        return Store.update(id, { followups: merged });
      },
      recognize: function (text) {
        if (!parserReady()) return null;
        return root.CrmParser.recognize(text, { today: schema.todayISO() });
      },
      applyRecognition: applyRecognition,
      exportExcel: exportExcel,
      exportJson: exportJson,
      stats: function () { return Store.stats(); },
      filtered: function () { return filteredDerived(); },
      toast: core.toast,
      openFollowupModal: openFollowupModal,
      state: state,
      render: render
    },
    store: Store,
    schema: schema,
    parser: root.CrmParser || null,
    xlsx: root.CrmXlsx || null
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})(typeof window !== 'undefined' ? window : globalThis);
