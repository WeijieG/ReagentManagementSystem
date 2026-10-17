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
    // collapsed 时只留图标：文字隐藏，改为原生 title 提示
    props: { current: String, collapsed: Boolean },
    template: [
      '<nav class="sidebar-nav">',
      '  <button v-for="r in mainRoutes" :key="r.key"',
      '          class="nav-item" :class="{ \'is-active\': r.key === current }"',
      '          :data-label="r.title" :aria-label="collapsed ? r.title : null"',
      '          @click="go(r.key)">',
      '    <svg class="nav-item-icon" viewBox="0 0 24 24" aria-hidden="true">',
      '      <path :d="r.icon"></path>',
      '    </svg>',
      '    <span class="nav-item-label">{{ r.title }}</span>',
      '  </button>',
      '</nav>',
      /* 设置入口固定在底部，不跟随上面的列表滚动 */
      '<div class="sidebar-bottom" v-if="bottomRoutes.length">',
      '  <button v-for="r in bottomRoutes" :key="r.key"',
      '          class="nav-item" :class="{ \'is-active\': r.key === current }"',
      '          :data-label="r.title" :aria-label="collapsed ? r.title : null"',
      '          @click="go(r.key)">',
      '    <svg class="nav-item-icon" viewBox="0 0 24 24" aria-hidden="true">',
      '      <path :d="r.icon"></path>',
      '    </svg>',
      '    <span class="nav-item-label">{{ r.title }}</span>',
      '  </button>',
      '</div>',
    ].join(''),
    computed: {
      mainRoutes: function () {
        return router.routes.filter(function (r) { return !r.bottom; });
      },
      bottomRoutes: function () {
        return router.routes.filter(function (r) { return r.bottom; });
      },
    },
    methods: {
      go: function (key) { router.go(key); },
    },
  };

  /* ============================================================
     使用须知闸门
     —— 未接受前挡住整个界面，对应旧实现里 init_ui 之前就跑的 check_disclaimer
     ============================================================ */
  var DisclaimerGate = {
    components: { 'ui-modal': RMS.components['ui-modal'] },
    template: [
      '<ui-modal title="使用须知" size="lg" :dismissible="false">',
      '  <div class="legal-text" v-if="text">{{ text }}</div>',
      '  <div v-else class="loading-overlay"><span class="spinner"></span> 正在加载…</div>',
      '  <div class="checkbox-row">',
      '    <input type="checkbox" id="disclaimer-ok" v-model="checked" />',
      '    <label for="disclaimer-ok">我已阅读、理解并同意遵守上述使用须知</label>',
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
      'ui-clock': RMS.components['ui-clock'],
    },
    template: [
      '<div class="app-shell" :class="{ \'is-collapsed\': sidebarCollapsed }">',
      '  <aside class="sidebar">',
      '    <div class="sidebar-brand">',
      '      <div class="sidebar-brand-mark"></div>',
      '      <div class="sidebar-brand-text">试剂管理</div>',
      /* 收起后品牌让位，这个按钮停在原来的位置，图标方向随之翻转 */
      '      <button class="sidebar-toggle" :title="toggleTitle"',
      '              @click="sidebarCollapsed = !sidebarCollapsed">',
      '        <svg class="sidebar-toggle-icon" viewBox="0 0 24 24" aria-hidden="true">',
      '          <rect x="3" y="3" width="18" height="18" rx="2"></rect>',
      '          <path :d="toggleIconPath"></path>',
      '        </svg>',
      '      </button>',
      '    </div>',
      '    <sidebar-nav :current="current"',
      '                 :collapsed="sidebarCollapsed"></sidebar-nav>',
      '    <div class="sidebar-footer">{{ versionLabel }}</div>',
      '  </aside>',

      '  <main class="main">',
      '    <header class="topbar">',
      '      <div class="topbar-title">{{ title }}</div>',
      '      <div class="topbar-actions">',
      '        <div class="topbar-actions-main">',
      '          <button class="btn btn-in" @click="inboundOpen = true">',
      '            <svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true">',
      '              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>',
      '              <path d="M7 10l5 5 5-5"></path>',
      '              <path d="M12 15V3"></path>',
      '            </svg>',
      '            入库',
      '          </button>',
      '          <button class="btn btn-out" @click="outboundOpen = true">',
      '            <svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true">',
      '              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>',
      '              <path d="M17 8l-5-5-5 5"></path>',
      '              <path d="M12 3v12"></path>',
      '            </svg>',
      '            出库',
      '          </button>',
      '        </div>',
      '        <ui-clock></ui-clock>',
      '      </div>',
      '    </header>',
      '    <section class="content" ref="content">',
      '      <component :is="screenComponent"',
      '                 @goto-records="gotoRecords"></component>',
      '    </section>',
      /* 置顶/置底按钮：跟随当前可滚动的区域（页面内容或表格） */
      '    <transition name="scroll-fade">',
      '      <div class="scroll-controls" v-if="hasScroll">',
      '        <transition name="scroll-btn">',
      '          <button v-if="canScrollUp" class="scroll-btn" title="回到顶部"',
      '                  @click="scrollTo(\'top\')">',
      '            <svg viewBox="0 0 24 24" aria-hidden="true">',
      '              <path d="M12 19V5"></path>',
      '              <path d="M5 12l7-7 7 7"></path>',
      '            </svg>',
      '          </button>',
      '        </transition>',
      '        <transition name="scroll-btn">',
      '          <button v-if="canScrollDown" class="scroll-btn" title="跳到底部"',
      '                  @click="scrollTo(\'bottom\')">',
      '            <svg viewBox="0 0 24 24" aria-hidden="true">',
      '              <path d="M12 5v14"></path>',
      '              <path d="M5 12l7 7 7-7"></path>',
      '            </svg>',
      '          </button>',
      '        </transition>',
      '      </div>',
      '    </transition>',
      '  </main>',

      /* transition 让抽屉在离开时也有滑出 + 遮罩淡出 */
      '  <transition name="drawer">',
      '    <inbound-drawer v-if="inboundOpen" @close="closeDrawers"></inbound-drawer>',
      '  </transition>',
      '  <transition name="drawer">',
      '    <outbound-drawer v-if="outboundOpen" @close="closeDrawers"></outbound-drawer>',
      '  </transition>',
      '  <disclaimer-gate v-if="showDisclaimer" @done="onDisclaimerDone"></disclaimer-gate>',
      '  <toast-stack></toast-stack>',
      '</div>',
    ].join(''),
    data: function () {
      return {
        inboundOpen: false,
        outboundOpen: false,
        disclaimerDismissed: false,
        // 只在本进程内有效：file:// 下 localStorage 不可用，不为这点偏好动后端
        sidebarCollapsed: false,
        // 滚动控件状态：hasScroll 表示当前区域可滚动
        hasScroll: false,
        canScrollUp: false,
        canScrollDown: false,
      };
    },
    computed: {
      current: function () { return router.current.value; },
      title: function () { return router.titleOf(router.current.value); },
      screenComponent: function () { return 'screen-' + router.current.value; },
      versionLabel: function () { return store.version ? ('版本 ' + store.version) : ''; },
      toggleTitle: function () {
        return this.sidebarCollapsed ? '展开侧栏' : '收起侧栏';
      },
      // 方框里一条竖线：线在左表示收起，线在右表示展开
      toggleIconPath: function () {
        return this.sidebarCollapsed ? 'M15 3v18' : 'M9 3v18';
      },
      showDisclaimer: function () {
        return store.ready && !store.disclaimerAccepted && !this.disclaimerDismissed;
      },
      dirty: function () { return store.dirty; },
    },
    watch: {
      // 换页或数据刷新后，滚动区域可能变化，重新判断一次
      current: function () { this.scheduleScrollRefresh(); },
      dirty: function () { this.scheduleScrollRefresh(); },
    },
    mounted: function () {
      window.addEventListener('keydown', this.onGlobalKeydown);
      window.addEventListener('resize', this.refreshScrollTarget);
      // scroll 不冒泡，用捕获阶段监听内容区与表格内部的滚动
      document.addEventListener('scroll', this.onAnyScroll, true);
      this.scheduleScrollRefresh();
    },
    beforeUnmount: function () {
      window.removeEventListener('keydown', this.onGlobalKeydown);
      window.removeEventListener('resize', this.refreshScrollTarget);
      document.removeEventListener('scroll', this.onAnyScroll, true);
      clearTimeout(this._scrollTimer);
    },
    methods: {
      /* 全局快捷键：Esc 关闭出入库界面；出入库快捷键打开对应抽屉。
         抽屉已打开、弹窗打开或正在输入框中打字时都不响应。 */
      onGlobalKeydown: function (e) {
        if (e.key === 'Escape' && (this.inboundOpen || this.outboundOpen)) {
          e.preventDefault();
          this.closeDrawers();
          return;
        }
        if (this.inboundOpen || this.outboundOpen) { return; }
        if (this.showDisclaimer) { return; }
        if (document.querySelector('.modal-scrim')) { return; }

        var el = e.target;
        var tag = (el && el.tagName) || '';
        if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT'
            || (el && el.isContentEditable)) {
          return;
        }

        var key = RMS.shortcutFromEvent(e);
        if (!key) { return; }
        if (key === store.shortcuts.inbound) {
          e.preventDefault();
          this.inboundOpen = true;
        } else if (key === store.shortcuts.outbound) {
          e.preventDefault();
          this.outboundOpen = true;
        }
      },
      closeDrawers: function () {
        this.inboundOpen = false;
        this.outboundOpen = false;
      },
      gotoRecords: function () { router.go('records'); },
      onDisclaimerDone: function () { this.disclaimerDismissed = true; },
      /* ---------- 置顶 / 置底 ---------- */

      /* 找到当前真正滚动的区域：列表页是表格内部，其余页面是内容区 */
      refreshScrollTarget: function () {
        var content = this.$refs.content;
        if (!content) { return; }

        var table = content.querySelector('.table-wrap');
        var target = null;
        if (content.scrollHeight > content.clientHeight + 4) {
          target = content;
        } else if (table && table.scrollHeight > table.clientHeight + 4) {
          target = table;
        }
        this._scrollEl = target;
        this.updateScrollState();
      },

      scheduleScrollRefresh: function () {
        // 数据是异步加载的，换页后稍等再判断可滚动区域
        var self = this;
        clearTimeout(this._scrollTimer);
        this._scrollTimer = setTimeout(function () { self.refreshScrollTarget(); }, 200);
        setTimeout(function () { self.refreshScrollTarget(); }, 900);
      },

      updateScrollState: function () {
        var el = this._scrollEl;
        if (!el) {
          this.hasScroll = false;
          this.canScrollUp = false;
          this.canScrollDown = false;
          return;
        }
        var max = el.scrollHeight - el.clientHeight;
        this.hasScroll = max > 4;
        this.canScrollUp = el.scrollTop > 4;
        this.canScrollDown = el.scrollTop < max - 4;
      },

      onAnyScroll: function (e) {
        var el = e.target;
        if (!el || !el.classList) { return; }
        var isContent = el.classList.contains('content');
        var isTable = el.classList.contains('table-wrap');
        if (!isContent && !isTable) { return; }

        if (el === this._scrollEl) {
          this.updateScrollState();
          return;
        }
        // 滚动目标变了（例如换页后表格开始滚动），重新判定一次
        this.refreshScrollTarget();
      },

      scrollTo: function (where) {
        var el = this._scrollEl;
        if (!el) { return; }
        var top = where === 'bottom' ? el.scrollHeight : 0;
        if (el.scrollTo) { el.scrollTo({ top: top, behavior: 'smooth' }); }
        else { el.scrollTop = top; }
      },
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
