/*!
 * 客户通 CustomerCRM — js/store.js
 * localStorage 仓储：客户记录、设置、示例数据、JSON 备份/恢复、统计聚合。
 * 依赖 js/schema.js 与 js/core.js；浏览器挂到 window.CrmStore。
 */
(function (root) {
  'use strict';

  var schema = root.CRM_SCHEMA;
  var core = root.CrmCore;

  var KEY = schema.STORAGE_KEY;
  var PAYLOAD_VERSION = 1;

  var DEFAULT_SETTINGS = { warnDays: 7, staleDays: 14, seeded: false };

  var state = null;       // { version, savedAt, settings, customers }
  var listeners = [];
  var storageOk = true;

  /* ------------------------------------------------------------------ */
  /* 1. 示例数据                                                          */
  /* ------------------------------------------------------------------ */

  /** 相对今天的天数偏移 → 'YYYY-MM-DD' */
  function ago(days) {
    return schema.addDays(schema.todayISO(), -Math.abs(Number(days) || 0));
  }

  var DEMO_BLUEPRINT = [
    {
      customerName: 'Michael Braun', companyName: 'Braun Trading GmbH', country: '德国',
      product: '不锈钢保温杯 500ml（定制 LOGO）', foundedAt: '1998', teamSize: '50-100 人',
      annualRevenue: '1200 万欧元', followupStatus: '已报价', dealStatus: '未成交',
      coreReason: '报价高出目标价约 8%，客户等终端零售客户确认后回复',
      firstInquiryAt: -71,
      followups: [
        { date: -3, note: '邮件确认样品尺寸与 LOGO 印刷位置，等寄样确认' },
        { date: -11, note: '电话沟通，客户确认 2000pcs 试单意向' },
        { date: -24, note: '首次邮件回复，发送价目表与产品册' }
      ]
    },
    {
      customerName: 'Elena Rossi', companyName: 'Rossi Import S.r.l.', country: '意大利',
      product: '陶瓷餐具套装 16 件套', foundedAt: '2005', teamSize: '20-50 人',
      annualRevenue: '800 万欧元', followupStatus: '谈判中', dealStatus: '未成交',
      coreReason: '客户要求 30 天交期，我方产线最快 45 天，正在协调排产',
      firstInquiryAt: -48,
      followups: [
        { date: -12, note: '视频会议确认包装方式，交期仍是唯一卡点' },
        { date: -26, note: '客户下单前审核通过，要求更新 PI' }
      ]
    },
    {
      customerName: 'Ahmed Hassan', companyName: 'El Nile Import & Export Co., Ltd', country: '埃及',
      product: '户外折叠椅', foundedAt: '2011', teamSize: '15 人',
      annualRevenue: '300 万美元', followupStatus: '寄样中', dealStatus: '未成交',
      coreReason: '样品已寄出，等待客户测试反馈后再谈价格',
      firstInquiryAt: -19,
      followups: [
        { date: -5, note: 'DHL 单号已发客户，预计 4 天到开罗' },
        { date: -14, note: '客户确认样品颜色与钢管厚度' }
      ]
    },
    {
      customerName: 'Sarah Thompson', companyName: 'Thompson Homewares Ltd', country: '英国',
      product: '硅胶厨具套装 12 件套', foundedAt: '2016', teamSize: '10-20 人',
      annualRevenue: '250 万英镑', followupStatus: '已成交', dealStatus: '已成交',
      coreReason: '价格与交期都满足，首单 5000 套，已收 30% 定金',
      firstInquiryAt: -35,
      followups: [
        { date: -1, note: '收到 30% 定金，安排下周投产' },
        { date: -8, note: '客户确认 PI 与包装稿' },
        { date: -21, note: '寄出样品，客户测试通过' }
      ]
    },
    {
      customerName: 'Lucas Silva', companyName: 'Silva Comércio Ltda', country: '巴西',
      product: '不锈钢水壶 750ml', foundedAt: '2009', teamSize: '100-200 人',
      annualRevenue: '2000 万雷亚尔', followupStatus: '新询盘', dealStatus: '未成交',
      coreReason: '刚接触，需求与目标价待确认',
      firstInquiryAt: -2,
      followups: [
        { date: -2, note: '阿里国际站询盘，已回复问需求数量与目标价' }
      ]
    },
    {
      customerName: 'Yuki Tanaka', companyName: 'Tanaka Trading Co., Ltd.', country: '日本',
      product: '保温饭盒 1.2L', foundedAt: '1985', teamSize: '30 人',
      annualRevenue: '5 亿日元', followupStatus: '已回复', dealStatus: '未成交',
      coreReason: '客户要求 JIS 认证与日文包装，认证周期约 6 周',
      firstInquiryAt: -62,
      followups: [
        { date: -18, note: '发送认证方案与报价，客户内部评审中' },
        { date: -40, note: '客户索取 JIS 认证资料' }
      ]
    },
    {
      customerName: 'Olga Petrova', companyName: 'Petrova Distribution LLC', country: '俄罗斯',
      product: '玻璃保鲜盒 5 件套', foundedAt: '2013', teamSize: '40-60 人',
      annualRevenue: '6 亿卢布', followupStatus: '暂停跟进', dealStatus: '未成交',
      coreReason: '付款方式谈不拢：客户坚持 OA 60 天，我方只接受 T/T 30% 定金',
      firstInquiryAt: -95,
      followups: [
        { date: -35, note: '客户表示可接受 LC at sight，等其银行额度' },
        { date: -58, note: '发送付款条件说明，客户未回复' }
      ]
    },
    {
      customerName: 'David Okafor', companyName: 'Okafor Global Ventures', country: '尼日利亚',
      product: '密胺餐具 10 件套', foundedAt: '2018', teamSize: '8 人',
      annualRevenue: '120 万美元', followupStatus: '已流失', dealStatus: '未成交',
      coreReason: '目标价低于我方成本 12%，且首单仅 300 套，无法开模',
      firstInquiryAt: -120,
      followups: [
        { date: -60, note: '客户最终从印度供应商下单，标记流失' },
        { date: -88, note: '二次报价仍低于成本，未达成' }
      ]
    },
    {
      customerName: 'Peter van Dijk', companyName: 'Van Dijk Trading B.V.', country: '荷兰',
      product: '保温杯 OEM（客户品牌）', foundedAt: '1994', teamSize: '50 人',
      annualRevenue: '1500 万欧元', followupStatus: '已成交', dealStatus: '已成交',
      coreReason: '老客户返单，签年度框架协议，年采购 5 万只',
      firstInquiryAt: -160,
      followups: [
        { date: -4, note: '框架协议盖章回传，Q1 首单 1.2 万只' },
        { date: -30, note: '客户来厂验厂通过' },
        { date: -75, note: '寄出 3 款样品供终端客户选型' }
      ]
    },
    {
      customerName: 'Kim Min-su', companyName: 'Seoul Living Co., Ltd', country: '韩国',
      product: '硅胶折叠水杯', foundedAt: '2015', teamSize: '25 人',
      annualRevenue: '120 亿韩元', followupStatus: '已报价', dealStatus: '未成交',
      coreReason: '价格可接受，客户等其终端客户确认颜色方案',
      firstInquiryAt: -28,
      followups: [
        { date: -9, note: '发送色卡与报价，客户内部会议定色' },
        { date: -20, note: '客户确认 3000pcs 试单，要求改包装' }
      ]
    },
    {
      customerName: 'James Carter', companyName: 'Carter Outdoor Supply Inc.', country: '美国',
      product: '不锈钢运动水壶 1L', foundedAt: '2007', teamSize: '200-500 人',
      annualRevenue: '8000 万美元', followupStatus: '已报价', dealStatus: '未成交',
      coreReason: '客户要求 FDA + Prop 65 检测报告，检测费用正在核算',
      firstInquiryAt: -14,
      followups: [
        { date: -7, note: '报价已发，附检测方案与费用预估' },
        { date: -13, note: '客户索取 FDA 检测报告模板' }
      ]
    },
    {
      customerName: 'Sophie Martin', companyName: 'Martin Cuisine SAS', country: '法国',
      product: '不粘锅套装', foundedAt: '2019', teamSize: '12 人',
      annualRevenue: '400 万欧元', followupStatus: '新询盘', dealStatus: '未成交',
      coreReason: '',
      firstInquiryAt: -1,
      followups: [
        { date: -1, note: '网站表单询盘，尚未回复' }
      ]
    }
  ];

  /* 示例客户的「产品名称 / 型号 / 主营业务 / 联系方式」（与产品库预置数据对应） */
  var DEMO_EXTRA = {
    'Michael Braun': { product: '不锈钢保温杯', productModel: '500ml', mainBusiness: '户外用品进口分销', email: 'm.braun@braun-trading.de', phone: '+49 711 4567 890', address: 'Industriestr. 12, 70565 Stuttgart, Germany' },
    'Elena Rossi': { product: '陶瓷餐具套装', productModel: '16 件套', mainBusiness: '家居餐具进口分销', email: 'e.rossi@rossiimport.it', phone: '+39 02 8765 4321', address: 'Via Brera 18, 20121 Milano, Italy' },
    'Ahmed Hassan': { product: '户外折叠椅', productModel: '标准款', mainBusiness: '户外用品批发零售', email: 'a.hassan@elnile-import.com', phone: '+20 100 234 5678', address: '25 Nile Corniche, Cairo 11511, Egypt' },
    'Sarah Thompson': { product: '硅胶厨具套装', productModel: '12 件套', mainBusiness: '厨房用品连锁零售', email: 's.thompson@thompsonhomewares.co.uk', phone: '+44 20 7946 0123', address: '14 Camden High St, London NW1 0JH, UK' },
    'Lucas Silva': { product: '不锈钢运动水壶', productModel: '1L', mainBusiness: '运动户外电商', email: 'lucas@silvacomercio.com.br', phone: '+55 11 98765 4321', address: 'Av. Paulista 1500, São Paulo 01310-100, Brazil' },
    'Yuki Tanaka': { product: '保温饭盒', productModel: '1.2L', mainBusiness: '日用百货批发', email: 'y.tanaka@tanaka-trading.jp', phone: '+81 3 1234 5678', address: '2-1-1 Uchisaiwaicho, Chiyoda-ku, Tokyo 100-0011, Japan' },
    'Olga Petrova': { product: '玻璃保鲜盒', productModel: '5 件套', mainBusiness: '商超渠道分销', email: 'o.petrova@petrova-dist.ru', phone: '+7 495 123 4567', address: 'Leninsky Prospekt 42, Moscow 119334, Russia' },
    'David Okafor': { product: '密胺餐具', productModel: '10 件套', mainBusiness: '酒店用品进口', email: 'd.okafor@okaforglobal.ng', phone: '+234 803 123 4567', address: '12 Broad St, Lagos Island, Lagos, Nigeria' },
    'Peter van Dijk': { product: '不锈钢保温杯', productModel: '750ml', mainBusiness: '自有品牌 OEM 采购', email: 'p.vandijk@vandijktrading.nl', phone: '+31 20 123 4567', address: 'Keizersgracht 120, 1015 CW Amsterdam, Netherlands' },
    'Kim Min-su': { product: '硅胶折叠水杯', productModel: '500ml', mainBusiness: '生活用品电商', email: 'mkim@seoulliving.kr', phone: '+82 2 1234 5678', address: '45 Teheran-ro, Gangnam-gu, Seoul 06234, Korea' },
    'James Carter': { product: '不锈钢保温杯', productModel: '1L', mainBusiness: '户外装备零售连锁', email: 'j.carter@carteroutdoor.com', phone: '+1 503 555 0142', address: '820 NW Marshall St, Portland, OR 97209, USA' },
    'Sophie Martin': { product: '不粘锅套装', productModel: '3 件套', mainBusiness: '厨房用品进口商', email: 's.martin@martincuisine.fr', phone: '+33 4 78 12 34 56', address: '22 Rue de la République, 69002 Lyon, France' }
  };

  function buildDemoRecords() {
    return DEMO_BLUEPRINT.map(function (item) {
      var extra = DEMO_EXTRA[item.customerName] || {};
      var rec = schema.normalize({
        id: schema.uid('c_'),
        firstInquiryAt: ago(item.firstInquiryAt),
        customerName: item.customerName,
        email: extra.email || '',
        phone: extra.phone || '',
        address: extra.address || '',
        product: extra.product || item.product,
        productModel: extra.productModel || item.productModel || '',
        country: item.country,
        companyName: item.companyName,
        mainBusiness: extra.mainBusiness || item.mainBusiness || '',
        foundedAt: item.foundedAt,
        teamSize: item.teamSize,
        annualRevenue: item.annualRevenue,
        followupStatus: item.followupStatus,
        dealStatus: item.dealStatus,
        coreReason: item.coreReason,
        followups: (item.followups || []).map(function (f) {
          return { date: ago(f.date), note: f.note };
        }),
        createdAt: new Date(Date.now() - Math.abs(item.firstInquiryAt) * 86400000).toISOString(),
        updatedAt: new Date(Date.now() - Math.abs((item.followups && item.followups[0] ? item.followups[0].date : 1)) * 86400000).toISOString()
      }, schema.todayISO());
      return rec;
    });
  }

  /* ------------------------------------------------------------------ */
  /* 2. 持久化                                                            */
  /* ------------------------------------------------------------------ */

  function emptyPayload() {
    return {
      version: PAYLOAD_VERSION,
      savedAt: schema.nowISO(),
      settings: JSON.parse(JSON.stringify(DEFAULT_SETTINGS)),
      customers: [],
      products: []
    };
  }

  function normalizeProduct(raw) {
    if (!raw || typeof raw !== 'object') return null;
    var name = String(raw.name || raw.product || raw.productName || '').trim().slice(0, 120);
    var model = String(raw.model || raw.productModel || '').trim().slice(0, 80);
    if (!name) return null;
    return {
      id: String(raw.id || '').trim() || schema.uid('p_'),
      name: name,
      model: model,
      createdAt: String(raw.createdAt || '').trim() || schema.nowISO()
    };
  }

  function defaultProducts() {
    return schema.DEFAULT_PRODUCTS.map(function (item) {
      return normalizeProduct({ name: item.name, model: item.model });
    });
  }

  function normalizePayload(raw) {
    var payload = emptyPayload();
    if (!raw || typeof raw !== 'object') return payload;
    if (raw.settings && typeof raw.settings === 'object') {
      payload.settings = {
        warnDays: finiteOr(raw.settings.warnDays, DEFAULT_SETTINGS.warnDays),
        staleDays: finiteOr(raw.settings.staleDays, DEFAULT_SETTINGS.staleDays),
        seeded: !!raw.settings.seeded
      };
    }
    var list = Array.isArray(raw.customers) ? raw.customers : Array.isArray(raw.records) ? raw.records : [];
    payload.customers = list.map(function (item) { return schema.normalize(item); }).filter(Boolean);
    if (Array.isArray(raw.products)) {
      payload.products = raw.products.map(normalizeProduct).filter(Boolean);
    } else {
      /* 从 v1.0 升级：没有产品库字段时写入预置产品 */
      payload.products = defaultProducts();
    }
    payload.savedAt = typeof raw.savedAt === 'string' ? raw.savedAt : schema.nowISO();
    return payload;
  }

  function finiteOr(value, fallback) {
    var n = Number(value);
    return isFinite(n) && n >= 0 ? Math.round(n) : fallback;
  }

  function load() {
    if (state) return state;
    var raw = core.storage.get(KEY, null);
    var payload = normalizePayload(raw);
    var isFirstRun = !raw;
    if (isFirstRun && payload.customers.length === 0) {
      payload.customers = buildDemoRecords();
      payload.settings.seeded = true;
      payload.products = defaultProducts();
      payload.savedAt = schema.nowISO();
      var ok = core.storage.set(KEY, payload);
      storageOk = ok;
    }
    if (!payload.products.length && payload.customers.length === 0 && !payload.settings.seeded) {
      payload.products = defaultProducts();
    }
    state = payload;
    return state;
  }

  function persist(reason) {
    state.savedAt = schema.nowISO();
    var ok = core.storage.set(KEY, state);
    storageOk = ok;
    notify(reason || 'save');
    return ok;
  }

  function notify(type, detail) {
    listeners.slice().forEach(function (fn) {
      try { fn({ type: type, detail: detail || null, state: state }); } catch (err) {
        console.warn('[CustomerCRM] store 订阅回调异常：', err);
      }
    });
  }

  /* ------------------------------------------------------------------ */
  /* 3. 查询与变更                                                        */
  /* ------------------------------------------------------------------ */

  function list() {
    var payload = load();
    return payload.customers.slice().sort(function (a, b) {
      var ua = a.updatedAt || '';
      var ub = b.updatedAt || '';
      if (ua === ub) return (b.firstInquiryAt || '').localeCompare(a.firstInquiryAt || '');
      return ub.localeCompare(ua);
    });
  }

  function get(id) {
    if (!id) return null;
    var found = load().customers.filter(function (item) { return item.id === id; })[0];
    return found || null;
  }

  function add(record) {
    var payload = load();
    var rec = schema.normalize(record);
    if (!rec.id) rec.id = schema.uid('c_');
    if (get(rec.id)) rec.id = schema.uid('c_');
    rec.createdAt = schema.nowISO();
    rec.updatedAt = rec.createdAt;
    payload.customers.push(rec);
    persist('add');
    return rec;
  }

  function update(id, patch) {
    var payload = load();
    for (var i = 0; i < payload.customers.length; i += 1) {
      if (payload.customers[i].id !== id) continue;
      var merged = {};
      var base = payload.customers[i];
      Object.keys(base).forEach(function (key) { merged[key] = base[key]; });
      Object.keys(patch || {}).forEach(function (key) { merged[key] = patch[key]; });
      var next = schema.normalize(merged);
      next.id = id;
      next.createdAt = base.createdAt;
      next.updatedAt = schema.nowISO();
      payload.customers[i] = next;
      persist('update');
      return next;
    }
    return null;
  }

  function remove(id) {
    var payload = load();
    var before = payload.customers.length;
    payload.customers = payload.customers.filter(function (item) { return item.id !== id; });
    if (payload.customers.length === before) return false;
    persist('remove');
    return true;
  }

  function replaceAll(records) {
    var payload = load();
    payload.customers = (records || []).map(function (item) { return schema.normalize(item); });
    persist('replace');
    return payload.customers.length;
  }

  function clear() {
    var payload = load();
    payload.customers = [];
    payload.settings.seeded = true;
    persist('clear');
    return true;
  }

  function seedDemo(force) {
    var payload = load();
    if (!force && payload.customers.length > 0) return 0;
    payload.customers = buildDemoRecords();
    payload.settings.seeded = true;
    persist('seed');
    return payload.customers.length;
  }

  function isEmpty() {
    return load().customers.length === 0;
  }

  /* ------------------------------------------------------------------ */
  /* 3.2 产品库（咨询产品预录入：产品名称 + 型号）                          */
  /* ------------------------------------------------------------------ */

  function listProducts() {
    return load().products.slice().sort(function (a, b) {
      var byName = String(a.name).localeCompare(String(b.name), 'zh-Hans-CN');
      if (byName !== 0) return byName;
      return String(a.model || '').localeCompare(String(b.model || ''), 'zh-Hans-CN');
    });
  }

  function productNames() {
    var seen = {};
    var names = [];
    listProducts().forEach(function (item) {
      if (seen[item.name]) return;
      seen[item.name] = true;
      names.push(item.name);
    });
    return names;
  }

  function productModels(name) {
    var target = String(name || '').trim();
    if (!target) return [];
    return listProducts().filter(function (item) {
      return item.name === target && item.model;
    }).map(function (item) { return item.model; });
  }

  /** 新增产品；name+model 已存在时直接返回已有项（不重复写入） */
  function addProduct(raw) {
    var payload = load();
    var product = normalizeProduct(raw);
    if (!product) return null;
    var exists = payload.products.filter(function (item) {
      return item.name === product.name && String(item.model || '') === String(product.model || '');
    })[0];
    if (exists) return exists;
    payload.products.push(product);
    persist('product-add');
    return product;
  }

  function removeProduct(id) {
    var payload = load();
    var before = payload.products.length;
    payload.products = payload.products.filter(function (item) { return item.id !== id; });
    if (payload.products.length === before) return false;
    persist('product-remove');
    return true;
  }

  function seedProducts(force) {
    var payload = load();
    if (!force && payload.products.length) return payload.products.length;
    payload.products = defaultProducts();
    persist('product-seed');
    return payload.products.length;
  }

  /** 产品库使用热度：{ '名称|型号': 客户数 } */
  function productUsage() {
    var usage = {};
    load().customers.forEach(function (rec) {
      var key = String(rec.product || '').trim() + '|' + String(rec.productModel || '').trim();
      if (!String(rec.product || '').trim()) return;
      usage[key] = (usage[key] || 0) + 1;
    });
    return usage;
  }

  /* ------------------------------------------------------------------ */
  /* 4. 设置                                                              */
  /* ------------------------------------------------------------------ */

  function settings() {
    return JSON.parse(JSON.stringify(load().settings));
  }

  function updateSettings(patch) {
    var payload = load();
    var next = payload.settings;
    if (patch && typeof patch === 'object') {
      if (patch.warnDays !== undefined) next.warnDays = finiteOr(patch.warnDays, next.warnDays);
      if (patch.staleDays !== undefined) next.staleDays = finiteOr(patch.staleDays, next.staleDays);
      if (patch.seeded !== undefined) next.seeded = !!patch.seeded;
    }
    persist('settings');
    return settings();
  }

  /* ------------------------------------------------------------------ */
  /* 5. 派生与统计                                                        */
  /* ------------------------------------------------------------------ */

  function deriveAll() {
    var opts = settings();
    var today = schema.todayISO();
    return list().map(function (rec) {
      return schema.derive(rec, { today: today, warnDays: opts.warnDays, staleDays: opts.staleDays });
    });
  }

  function stats() {
    var opts = settings();
    var today = schema.todayISO();
    var derived = deriveAll();
    var total = derived.length;
    var deal = 0;
    var open = 0;
    var overdue = 0;
    var due = 0;
    var daysSum = 0;
    var daysCount = 0;
    var countryMap = {};
    var statusMap = {};
    var recent = [];
    var overdueList = [];

    derived.forEach(function (rec) {
      if (rec.isDeal) deal += 1;
      if (rec.isOpen) open += 1;
      if (rec.isOverdue) { overdue += 1; overdueList.push(rec); }
      if (rec.isDue) due += 1;
      if (rec.daysSince !== null && isFinite(rec.daysSince)) { daysSum += rec.daysSince; daysCount += 1; }
      var country = rec.country || '未填写';
      countryMap[country] = (countryMap[country] || 0) + 1;
      var status = rec.followupStatus || '新询盘';
      statusMap[status] = (statusMap[status] || 0) + 1;
      rec.followups.forEach(function (f) {
        recent.push({
          id: rec.id,
          customerName: rec.customerName,
          companyName: rec.companyName,
          country: rec.country,
          date: f.date,
          note: f.note,
          daysAgo: Math.max(0, schema.daysBetween(f.date, today)),
          followupStatus: rec.followupStatus,
          dealStatus: rec.dealStatus
        });
      });
    });

    recent.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
    overdueList.sort(function (a, b) { return (b.daysSince || 0) - (a.daysSince || 0); });

    var byCountry = Object.keys(countryMap).map(function (name) {
      return { name: name, count: countryMap[name] };
    }).sort(function (a, b) { return b.count - a.count; });

    var byStatus = schema.FOLLOWUP_STATUS.map(function (status) {
      return { status: status, count: statusMap[status] || 0, tone: schema.statusTone(status) };
    }).filter(function (item) { return item.count > 0; });
    // 未在枚举内的历史状态兜底展示
    Object.keys(statusMap).forEach(function (status) {
      if (schema.FOLLOWUP_STATUS.indexOf(status) !== -1) return;
      byStatus.push({ status: status, count: statusMap[status], tone: schema.statusTone(status) });
    });

    return {
      total: total,
      deal: deal,
      open: open,
      overdue: overdue,
      due: due,
      dealRate: total ? deal / total : 0,
      avgDaysSince: daysCount ? Math.round(daysSum / daysCount) : null,
      byCountry: byCountry,
      byStatus: byStatus,
      recent: recent.slice(0, 8),
      overdueList: overdueList.slice(0, 6),
      warnDays: opts.warnDays,
      staleDays: opts.staleDays,
      today: today
    };
  }

  /* ------------------------------------------------------------------ */
  /* 6. 备份 / 恢复                                                       */
  /* ------------------------------------------------------------------ */

  function exportJSON() {
    var payload = load();
    return JSON.stringify({
      app: 'CustomerCRM',
      version: schema.VERSION,
      exportedAt: schema.nowISO(),
      settings: payload.settings,
      products: payload.products,
      customers: payload.customers
    }, null, 2);
  }

  /**
   * 导入 JSON 备份。mode: 'merge'（默认，同 id 覆盖）| 'replace'（清空后导入）
   * 返回 { ok, imported, skipped, total, mode, products }
   */
  function importJSON(text, mode) {
    var parsed;
    try { parsed = JSON.parse(text); } catch (err) {
      return { ok: false, error: 'JSON 解析失败：' + (err && err.message ? err.message : '格式不正确') };
    }
    var incoming = Array.isArray(parsed) ? parsed : (parsed && (parsed.customers || parsed.records)) || null;
    if (!Array.isArray(incoming)) return { ok: false, error: '文件里没有找到客户数组（customers）' };
    var payload = load();
    var importMode = mode === 'replace' ? 'replace' : 'merge';
    if (importMode === 'replace') payload.customers = [];
    var imported = 0;
    var skipped = 0;
    incoming.forEach(function (item) {
      if (!item || typeof item !== 'object') { skipped += 1; return; }
      var rec = schema.normalize(item);
      if (!rec.customerName && !rec.companyName && !rec.firstInquiryAt) { skipped += 1; return; }
      var existingIndex = -1;
      for (var i = 0; i < payload.customers.length; i += 1) {
        if (payload.customers[i].id === rec.id) { existingIndex = i; break; }
      }
      if (existingIndex === -1) payload.customers.push(rec);
      else payload.customers[existingIndex] = rec;
      imported += 1;
    });
    var importedProducts = 0;
    if (Array.isArray(parsed.products)) {
      if (importMode === 'replace') payload.products = [];
      parsed.products.forEach(function (raw) {
        var product = normalizeProduct(raw);
        if (!product) return;
        var dup = payload.products.some(function (item) {
          return item.name === product.name && String(item.model || '') === String(product.model || '');
        });
        if (dup) return;
        payload.products.push(product);
        importedProducts += 1;
      });
    }
    if (parsed.settings && importMode === 'merge') {
      payload.settings.warnDays = finiteOr(parsed.settings.warnDays, payload.settings.warnDays);
      payload.settings.staleDays = finiteOr(parsed.settings.staleDays, payload.settings.staleDays);
    }
    payload.settings.seeded = true;
    persist('import');
    return {
      ok: true,
      imported: imported,
      skipped: skipped,
      total: payload.customers.length,
      mode: importMode,
      products: importedProducts
    };
  }

  /* ------------------------------------------------------------------ */
  /* 7. 导出                                                             */
  /* ------------------------------------------------------------------ */

  root.CrmStore = {
    KEY: KEY,
    DEFAULT_SETTINGS: DEFAULT_SETTINGS,

    init: function () {
      storageOk = core.storage.available();
      var payload = load();
      return { total: payload.customers.length, seeded: payload.settings.seeded, storageOk: storageOk };
    },
    subscribe: function (fn) {
      if (typeof fn !== 'function') return function () {};
      listeners.push(fn);
      return function () {
        var index = listeners.indexOf(fn);
        if (index !== -1) listeners.splice(index, 1);
      };
    },

    list: list,
    get: get,
    add: add,
    update: update,
    remove: remove,
    replaceAll: replaceAll,
    clear: clear,
    seedDemo: seedDemo,
    isEmpty: isEmpty,

    listProducts: listProducts,
    productNames: productNames,
    productModels: productModels,
    addProduct: addProduct,
    removeProduct: removeProduct,
    seedProducts: seedProducts,
    productUsage: productUsage,

    settings: settings,
    updateSettings: updateSettings,
    deriveAll: deriveAll,
    stats: stats,

    exportJSON: exportJSON,
    importJSON: importJSON,

    savedAt: function () { return load().savedAt; },
    storageOk: function () { return storageOk; },
    storageAvailable: function () { return core.storage.available(); },
    demoRecords: function () { return buildDemoRecords(); }
  };
})(typeof window !== 'undefined' ? window : globalThis);
