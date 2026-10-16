/* 应用入口：组装外壳、注册组件、挂载。
 *
 * 必须放在所有其它 js 之后加载。
 */
window.RMS = window.RMS || {};

(function (RMS) {
  'use strict';

  var store = RMS.store;
  var router = RMS.router;

  /* ============================================================
     全局错误兜底
     —— 前端一旦抛异常，若不做处理就只表现为白屏，而目标机器上
        往往没有 DevTools。这里把任何未捕获的错误渲染成可见面板。
     ============================================================ */
  function showFatal(message, detail) {
    var el = document.getElementById('fatal-panel');
    if (!el) { return; }
    el.style.display = 'block';
    var msg = document.getElementById('fatal-message');
    var det = document.getElementById('fatal-detail');
    if (msg) { msg.textContent = message; }
    if (det) { det.textContent = detail || ''; }
  }

  window.addEventListener('error', function (e) {
    console.error('[未捕获错误]', e.message, e.filename + ':' + e.lineno);
    showFatal('界面发生错误',
      (e.message || '') + '\n' + (e.filename || '') + ':' + (e.lineno || ''));
  });

  window.addEventListener('unhandledrejection', function (e) {
    var reason = e.reason;
    var text = (reason && reason.message) ? reason.message : String(reason);
    console.error('[未处理的 Promise 拒绝]', reason);
    // 只有通道建立失败这类基础故障才值得弹面板，业务错误由各处自己 toast
    if (/通道|WebChannel|后端未注册/.test(text)) {
      showFatal('无法与后端通信', text);
    }
  });

  /* ============================================================
     通用组件
     ============================================================ */

  var ToastStack = {
    template: [
      '<div class="toast-stack">',
      '  <div v-for="t in toasts" :key="t.id" class="toast" :class="\'toast-\' + t.kind">',
      '    <div class="toast-text">{{ t.text }}</div>',
      '    <button class="toast-close" @click="dismiss(t.id)">&times;</button>',
      '  </div>',
      '</div>',
    ].join(''),
    computed: {
      toasts: function () { return store.toasts; },
    },
    methods: {
      dismiss: function (id) { RMS.dismissToast(id); },
    },
  };

  var SidebarNav = {
    props: { current: String },
    template: [
      '<nav class="sidebar-nav">',
      '  <div class="nav-group-label">库存</div>',
      '  <button v-for="r in routes" :key="r.key"',
      '          class="nav-item" :class="{ \'is-active\': r.key === current }"',
      '          @click="go(r.key)">',
      '    <span class="nav-item-dot"></span>{{ r.title }}',
      '  </button>',
      '</nav>',
    ].join(''),
    computed: {
      routes: function () { return router.routes; },
    },
    methods: {
      go: function (key) { router.go(key); },
    },
  };

  /* ============================================================
     免责声明闸门
     —— 未接受前挡住整个界面，对应旧实现里 init_ui 之前就跑的 check_disclaimer
     ============================================================ */
  var DisclaimerGate = {
    components: { 'ui-modal': RMS.components['ui-modal'] },
    template: [
      '<ui-modal title="用前须知" size="lg" :dismissible="false">',
      '  <div class="legal-text" v-if="text">{{ text }}</div>',
      '  <div v-else class="loading-overlay"><span class="spinner"></span> 正在加载…</div>',
      '  <div class="checkbox-row">',
      '    <input type="checkbox" id="disclaimer-ok" v-model="checked" />',
      '    <label for="disclaimer-ok">我已阅读、理解并同意接受上述条款</label>',
      '  </div>',
      '  <template #footer>',
      '    <button class="btn" @click="reject">拒绝</button>',
      '    <button class="btn btn-primary" :disabled="!checked" @click="accept">接受</button>',
      '  </template>',
      '</ui-modal>',
    ].join(''),
    data: function () {
      return { text: '', checked: false };
    },
    created: function () {
      var self = this;
      RMS.bridge.call('app.disclaimer').then(function (res) {
        self.text = res.text;
      }).catch(function () {
        self.text = '无法加载说明文本，请联系管理员。';
      });
    },
    methods: {
      accept: function () {
        var self = this;
        RMS.bridge.call('disclaimer.accept', { accepted: true }).then(function () {
          store.disclaimerAccepted = true;
          self.$emit('done');
        });
      },
      reject: function () {
        RMS.bridge.call('disclaimer.accept', { accepted: false }).then(function () {
          RMS.bridge.call('app.quit');
        });
      },
    },
  };

  /* ============================================================
     外壳
     ============================================================ */
  var AppShell = {
    components: {
      'sidebar-nav': SidebarNav,
      'toast-stack': ToastStack,
      'disclaimer-gate': DisclaimerGate,
      'inbound-drawer': RMS.components['inbound-drawer'],
      'outbound-drawer': RMS.components['outbound-drawer'],
    },
    template: [
      '<div class="app-shell">',
      '  <aside class="sidebar">',
      '    <div class="sidebar-brand">',
      '      <div class="sidebar-brand-mark"></div>',
      '      <div class="sidebar-brand-text">试剂管理</div>',
      '    </div>',
      '    <sidebar-nav :current="current"></sidebar-nav>',
      '    <div class="sidebar-footer">{{ versionLabel }}</div>',
      '  </aside>',

      '  <main class="main">',
      '    <header class="topbar">',
      '      <div class="topbar-title">{{ title }}</div>',
      '      <div class="topbar-actions">',
      '        <button class="btn btn-primary" @click="inboundOpen = true">入库</button>',
      '        <button class="btn" @click="outboundOpen = true">出库</button>',
      '        <button class="icon-btn" :title="themeTitle" @click="toggleTheme">',
      '          {{ themeIcon }}</button>',
      '      </div>',
      '    </header>',
      '    <section class="content">',
      '      <component :is="screenComponent"',
      '                 @open-inbound="inboundOpen = true"',
      '                 @open-outbound="outboundOpen = true"',
      '                 @goto-records="gotoRecords"></component>',
      '    </section>',
      '  </main>',

      '  <inbound-drawer v-if="inboundOpen" @close="closeDrawers"></inbound-drawer>',
      '  <outbound-drawer v-if="outboundOpen" @close="closeDrawers"></outbound-drawer>',
      '  <disclaimer-gate v-if="showDisclaimer" @done="onDisclaimerDone"></disclaimer-gate>',
      '  <toast-stack></toast-stack>',
      '</div>',
    ].join(''),
    data: function () {
      return {
        inboundOpen: false,
        outboundOpen: false,
        disclaimerDismissed: false,
      };
    },
    computed: {
      current: function () { return router.current.value; },
      title: function () { return router.titleOf(router.current.value); },
      screenComponent: function () { return 'screen-' + router.current.value; },
      versionLabel: function () { return store.version ? ('版本 ' + store.version) : ''; },
      themeIcon: function () { return store.resolvedTheme === 'dark' ? '☽' : '☀'; },
      themeTitle: function () {
        return store.resolvedTheme === 'dark' ? '切换到浅色主题' : '切换到深色主题';
      },
      showDisclaimer: function () {
        return store.ready && !store.disclaimerAccepted && !this.disclaimerDismissed;
      },
    },
    methods: {
      toggleTheme: function () {
        RMS.setTheme(store.resolvedTheme === 'dark' ? 'light' : 'dark');
      },
      closeDrawers: function () {
        this.inboundOpen = false;
        this.outboundOpen = false;
      },
      gotoRecords: function () { router.go('records'); },
      onDisclaimerDone: function () { this.disclaimerDismissed = true; },
    },
  };

  /* ============================================================
     占位页面 —— 尚未迁移的页面用这个兜底，避免整页空白
     ============================================================ */
  function placeholder(title) {
    return {
      template: [
        '<div class="page"><div class="card"><div class="card-body">',
        '  <div style="font-size:var(--text-lg);font-weight:600;margin-bottom:8px">',
        '    ' + title + '</div>',
        '  <div class="muted">该页面尚未迁移。</div>',
        '</div></div></div>',
      ].join(''),
    };
  }

  function registerAll(app) {
    app.component('app-shell', AppShell);

    // 共享组件
    var components = RMS.components || {};
    Object.keys(components).forEach(function (name) {
      app.component(name, components[name]);
    });

    // 页面
    var screens = RMS.screens || {};
    Object.keys(screens).forEach(function (name) {
      app.component(name, screens[name]);
    });

    // 未注册的路由落到占位组件
    router.routes.forEach(function (r) {
      var name = 'screen-' + r.key;
      if (!screens[name]) { app.component(name, placeholder(r.title)); }
    });
  }

  /* ============================================================
     启动
     ============================================================ */
  function start() {
    var app = Vue.createApp(AppShell);
    app.config.errorHandler = function (err, vm, info) {
      console.error('[Vue 错误]', err, info);
      showFatal('界面渲染出错',
        (err && err.message ? err.message : String(err)) + '\n' + (info || ''));
    };
    registerAll(app);
    app.mount('#app');

    RMS.bootstrap().catch(function (err) {
      showFatal('启动失败', err && err.message ? err.message : String(err));
    });
  }

  // 键盘导航才显示焦点环（:focus-visible 需要 Chrome 86）
  document.addEventListener('keydown', function () {
    document.body.classList.add('kb-focus');
  });
  document.addEventListener('mousedown', function () {
    document.body.classList.remove('kb-focus');
  });

  RMS.app = { start: start, showFatal: showFatal };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
}(window.RMS));
