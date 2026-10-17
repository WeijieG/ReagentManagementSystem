/* 极简 hash 路由。
 *
 * 不引入 vue-router：它的 createWebHashHistory 在 file:// 这个不透明源上
 * 对 history.state / base 的处理不稳定，而且多一个需要在 Chromium 83 上
 * 单独验证的产物。这里 40 行足够。
 */
window.RMS = window.RMS || {};

(function (RMS) {
  'use strict';

  // icon 是 24x24 视图框里的 SVG path，侧栏收起后只靠它区分入口。
  var ROUTES = [
    { key: 'dashboard', title: '概览',
      icon: 'M3 3h7v7H3z M14 3h7v7h-7z M14 14h7v7h-7z M3 14h7v7H3z' },
    { key: 'inventory', title: '库存管理',
      icon: 'M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z' +
            ' M3.27 6.96 12 12.01l8.73-5.05 M12 22.08V12' },
    { key: 'records',   title: '出入库记录',
      icon: 'M8 6h13 M8 12h13 M8 18h13 M3 6h.01 M3 12h.01 M3 18h.01' },
    { key: 'names',     title: '试剂名称',
      icon: 'M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z M7 7h.01' },
    // bottom: true 的路由固定在侧栏底部，不随主导航列表滚动
    { key: 'settings',  title: '设置', bottom: true,
      icon: 'M4 21v-7 M4 10V3 M12 21v-9 M12 8V3 M20 21v-5 M20 12V3 M1 14h6 M9 8h6 M17 16h6' },
  ];

  var KEYS = ROUTES.map(function (r) { return r.key; });

  function parseHash() {
    var raw = (window.location.hash || '').replace(/^#\/?/, '');
    // 无 hash 或未知路由时回到概览
    return KEYS.indexOf(raw) >= 0 ? raw : 'dashboard';
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
