/* 极简 hash 路由。
 *
 * 不引入 vue-router：它的 createWebHashHistory 在 file:// 这个不透明源上
 * 对 history.state / base 的处理不稳定，而且多一个需要在 Chromium 83 上
 * 单独验证的产物。这里 40 行足够。
 */
window.RMS = window.RMS || {};

(function (RMS) {
  'use strict';

  var ROUTES = [
    { key: 'dashboard', title: '概览' },
    { key: 'inventory', title: '库存管理' },
    { key: 'records',   title: '出入库记录' },
    { key: 'names',     title: '试剂名称' },
    // bottom: true 的路由固定在侧栏底部，不随主导航列表滚动
    { key: 'settings',  title: '设置', bottom: true },
  ];

  var KEYS = ROUTES.map(function (r) { return r.key; });

  function parseHash() {
    var raw = (window.location.hash || '').replace(/^#\/?/, '');
    return KEYS.indexOf(raw) >= 0 ? raw : 'inventory';
  }

  var current = Vue.ref(parseHash());

  window.addEventListener('hashchange', function () {
    current.value = parseHash();
  });

  function go(key) {
    if (KEYS.indexOf(key) < 0) { return; }
    window.location.hash = '#/' + key;
    // hashchange 在部分场景下不会立即触发，这里兜一手
    current.value = key;
  }

  function titleOf(key) {
    for (var i = 0; i < ROUTES.length; i++) {
      if (ROUTES[i].key === key) { return ROUTES[i].title; }
    }
    return '';
  }

  RMS.router = {
    routes: ROUTES,
    current: current,
    go: go,
    titleOf: titleOf,
  };
}(window.RMS));
