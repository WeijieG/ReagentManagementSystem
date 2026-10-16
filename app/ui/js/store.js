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

  var store = Vue.reactive({
    // 应用信息
    ready: false,
    fatal: '',
    version: '',
    dbPath: '',

    // 主题：light / dark / system
    theme: 'system',
    resolvedTheme: 'light',

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

  /* ---------------- 主题 ---------------- */

  function systemPrefersDark() {
    // matchMedia 在 Chromium 83 可用
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  }

  function applyTheme(mode) {
    var resolved = mode;
    if (mode === 'system') {
      resolved = systemPrefersDark() ? 'dark' : 'light';
    }
    store.resolvedTheme = resolved;
    document.documentElement.setAttribute('data-theme', resolved);
  }

  function setTheme(mode) {
    store.theme = mode;
    applyTheme(mode);
    // 落盘到 userData.ini。file:// 源下 localStorage 不可用，只能走后端。
    bridge.call('theme.set', { mode: mode }).catch(function (err) {
      console.warn('[theme] 保存主题偏好失败:', err.message);
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
      store.theme = info.theme || 'system';
      store.dbRepaired = info.dbRepaired || [];
      store.dbIssues = info.dbIssues || [];
      store.networkOk = info.networkOk !== false;
      store.networkError = info.networkError || '';
      applyTheme(store.theme);
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

  bridge.on('theme.changed', function (data) {
    if (data && data.mode) { store.theme = data.mode; applyTheme(data.mode); }
  });

  // 跟随系统主题：仅在用户选「跟随系统」时生效
  if (window.matchMedia) {
    var mq = window.matchMedia('(prefers-color-scheme: dark)');
    var onChange = function () {
      if (store.theme === 'system') { applyTheme('system'); }
    };
    if (mq.addEventListener) { mq.addEventListener('change', onChange); }
    else if (mq.addListener) { mq.addListener(onChange); }
  }

  RMS.store = store;
  RMS.bootstrap = bootstrap;
  RMS.setTheme = setTheme;
  RMS.toast = toast;
  RMS.dismissToast = dismissToast;
}(window.RMS));
