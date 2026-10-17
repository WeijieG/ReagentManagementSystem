/* 全局状态。
 *
 * 只有一个 reactive 对象，通过 provide/inject 下发，不引入状态库。
 * 跨窗口的耦合靠后端推送的 db.changed 事件解除 —— 旧实现里
 * 入库/出库对话框关闭后由父窗口回头调 refresh_widget()，新模型下改为
 * 后端主动通知、前端各自刷新当前页。
 */
window.RMS = window.RMS || {};

(function (RMS) {
  'use strict';

  var bridge = RMS.bridge;

  /* 界面风格清单。明暗色不再单独切换，而是并入风格：
   * 每一项同时决定皮肤（data-style）与明暗（data-theme）。
   * 新增风格只需要两步：
   *   1. 在 css 里补一套 [data-style="<id>"][data-theme="..."] 规则；
   *   2. 往下面加一行（value 用小写字母/数字/连字符）。
   * 设置页的下拉读这份清单，后端不需要跟着改。
   */
  var STYLES = [
    { value: 'light', label: '默认（白）', style: 'default', theme: 'light' },
    { value: 'dark',  label: '默认（黑）', style: 'default', theme: 'dark' },
    // 科技只有深色版
    { value: 'tech',  label: '科技（黑）', style: 'tech',    theme: 'dark' },
    { value: 'minimal', label: '极简（白）', style: 'minimal', theme: 'light' },
    { value: 'lively',  label: '灵动（白）', style: 'lively',  theme: 'light' },
    { value: 'ios',     label: 'iOS（白）',  style: 'ios',     theme: 'light' },
  ];

  function findStyle(value) {
    for (var i = 0; i < STYLES.length; i++) {
      if (STYLES[i].value === value) { return STYLES[i]; }
    }
    return STYLES[0];
  }

  var store = Vue.reactive({
    // 应用信息
    ready: false,
    fatal: '',
    version: '',
    dbPath: '',

    // 界面风格：light（默认白） / dark（默认黑） / tech（科技黑）
    style: 'light',

    // 免责声明
    disclaimerAccepted: true,

    // 开库一致性自检：修复过的孤立名称 / 修不了的违例
    dbRepaired: [],
    dbIssues: [],

    // 网络组件（socket/ssl）是否可用；不可用时仅"检查更新"受影响
    networkOk: true,
    networkError: '',

    // 库存页
    inventory: {
      rows: [],
      total: 0,
      page: 1,
      pageSize: 50,
      pages: 1,
      loading: false,
      filters: { name: '', batch: '', expiry: 'all' },
      sort: { col: 'id', dir: 'asc' },
    },

    // 出入库记录页
    records: {
      rows: [],
      total: 0,
      page: 1,
      pageSize: 50,
      pages: 1,
      loading: false,
      filters: { type: 'all', name: '', batch: '', start: '', end: '' },
    },

    // 试剂名称页
    names: {
      rows: [],
      loading: false,
      filter: '',
    },

    // 概览页
    dashboard: {
      data: null,
      loading: false,
    },

    // 后端推送的失效信号：数据库有变动时自增，各页 watch 它重新取数
    dirty: 0,

    toasts: [],
    nextToastId: 1,
  });

  /* ---------------- 界面风格 ---------------- */

  function applyStyle(style) {
    // 未知风格（例如旧配置里的值）回落到白色默认，避免界面停在半套样式上
    var resolved = findStyle(style);
    store.style = resolved.value;
    document.documentElement.setAttribute('data-style', resolved.style);
    document.documentElement.setAttribute('data-theme', resolved.theme);
  }

  function setStyle(style) {
    applyStyle(style);
    // 落盘到 userData.ini。file:// 源下 localStorage 不可用，只能走后端。
    bridge.call('style.set', { style: store.style }).catch(function (err) {
      console.warn('[style] 保存界面风格偏好失败:', err.message);
    });
  }

  /* ---------------- Toast ---------------- */

  function toast(kind, text, timeoutMs) {
    var id = store.nextToastId++;
    store.toasts.push({ id: id, kind: kind, text: text });
    var ttl = timeoutMs === undefined ? (kind === 'error' ? 6000 : 3000) : timeoutMs;
    if (ttl > 0) {
      setTimeout(function () { dismissToast(id); }, ttl);
    }
    return id;
  }

  function dismissToast(id) {
    for (var i = 0; i < store.toasts.length; i++) {
      if (store.toasts[i].id === id) {
        store.toasts.splice(i, 1);
        return;
      }
    }
  }

  /* ---------------- 初始化 ---------------- */

  function bootstrap() {
    return bridge.call('app.bootstrap').then(function (info) {
      store.version = info.version || '';
      store.dbPath = info.dbPath || '';
      store.disclaimerAccepted = !!info.disclaimerAccepted;
      store.dbRepaired = info.dbRepaired || [];
      store.dbIssues = info.dbIssues || [];
      store.networkOk = info.networkOk !== false;
      store.networkError = info.networkError || '';
      applyStyle(info.style || 'light');
      store.ready = true;

      // 网络组件坏掉只影响"检查更新"，不影响库存与出入库 ——
      // 说明清楚，别让用户以为整个程序有问题
      if (!store.networkOk) {
        toast('warning',
          '网络组件不可用，「检查更新」会失败。库存、出入库、导出等核心功能不受影响。' +
          '程序目录下的 network-error.txt 记录了详细信息。', 12000);
      }

      // 老库自检结果要让用户看见，而不是等他写数据时才撞上报错
      if (store.dbIssues.length) {
        toast('error',
          '数据库存在 ' + store.dbIssues.length + ' 条无法自动修复的记录不一致，' +
          '相关记录的出入库操作可能失败。建议先备份数据库。', 15000);
      } else if (store.dbRepaired.length) {
        toast('info',
          '已自动修复 ' + store.dbRepaired.length + ' 个缺失的试剂名称。', 6000);
      }

      return info;
    });
  }

  // 后端推送的事件统一在这里落到状态上
  bridge.on('db.changed', function () {
    store.dirty++;
  });

  bridge.on('toast', function (data) {
    if (data && data.text) { toast(data.kind || 'info', data.text); }
  });

  bridge.on('style.changed', function (data) {
    if (data && data.style) { applyStyle(data.style); }
  });

  RMS.store = store;
  RMS.styles = STYLES;
  RMS.bootstrap = bootstrap;
  RMS.setStyle = setStyle;
  RMS.toast = toast;
  RMS.dismissToast = dismissToast;
}(window.RMS));
