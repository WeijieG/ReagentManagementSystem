/* 可复用组件。注册到 RMS.components，由 app.js 统一接进 Vue 应用。 */
window.RMS = window.RMS || {};
window.RMS.components = window.RMS.components || {};

(function (RMS) {
  'use strict';

  var components = RMS.components;

  /* ============================================================
     效期徽章 —— 阈值与后端 _expiry_state 一致
     ============================================================ */
  components['ui-expiry-badge'] = {
    props: {
      state: String,      // expired | warn30 | warn90 | ok
      daysLeft: Number,
    },
    template: '<span class="badge" :class="cls">{{ label }}</span>',
    computed: {
      cls: function () {
        return {
          expired: 'badge-expired',
          warn30: 'badge-warn30',
          warn90: 'badge-warn90',
          ok: 'badge-ok',
        }[this.state] || 'badge-neutral';
      },
      label: function () {
        if (this.state === 'expired') {
          return '已过期 ' + Math.abs(this.daysLeft || 0) + ' 天';
        }
        if (this.daysLeft === null || this.daysLeft === undefined) { return '—'; }
        return '剩 ' + this.daysLeft + ' 天';
      },
    },
  };

  /* ============================================================
     系统时钟
     ------------------------------------------------------------
     业务时间优先使用网络时间，断网时回退本机时间。顶栏常驻一个
     时钟，方便第一时间发现电脑日期/时间不对。

     网络校时见 app/handlers/timecheck.py；断网时弹窗提醒，本机时间
     与网络时间相差过大时也会提醒用户校正系统时钟。
     ============================================================ */
  var WEEKDAYS = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
  var TIME_RECHECK_MS = 30 * 60 * 1000;   // 每 30 分钟重新校一次

  components['ui-clock'] = {
    template: [
      '<div class="topbar-clock" :class="{ \'is-suspect\': suspect }" :title="title">',
      '  <span class="topbar-clock-source" :class="sourceClass">{{ sourceLabel }}</span>',
      '  <span class="topbar-clock-date">{{ dateText }}</span>',
      '  <span class="topbar-clock-week">{{ weekText }}</span>',
      '  <span class="topbar-clock-time">{{ timeText }}</span>',
      '</div>',
      /* 校时失败/偏差过大时的提示。ui-modal 已全局注册。 */
      '<ui-modal v-if="warn" :title="warn.title" @close="warn = null">',
      '  <div class="time-warn-text">{{ warn.text }}</div>',
      '  <div v-if="warn.detail" class="time-warn-detail">',
      '    <button type="button" class="link-btn" @click="showDetail = !showDetail">',
      '      {{ showDetail ? "收起技术详情" : "查看技术详情" }}</button>',
      '    <div v-if="showDetail" class="mono">{{ warn.detail }}</div>',
      '  </div>',
      '  <template #footer>',
      '    <button class="btn btn-primary" @click="warn = null">我知道了</button>',
      '  </template>',
      '</ui-modal>',
    ].join(''),
    data: function () {
      return {
        now: new Date(),
        timer: null,
        recheckTimer: null,
        checkTimeout: null,
        // 网络时间基准：上次校时的网络毫秒数 + 当时的单调时钟读数
        networkBaseMs: 0,
        perfBaseMs: 0,
        lastWallMs: Date.now(),
        lastPerfMs: null,
        // checking：正在请求网络时间；network：已用网络时间；local：断网回退
        syncState: 'checking',
        offsetMs: 0,
        offsetSeconds: 0,
        warnThresholdSeconds: 120,
        netSource: '',
        netError: '',
        lastCheckAt: 0,
        networkWarned: false,
        offsetWarned: false,
        warn: null,
        showDetail: false,
      };
    },
    computed: {
      /* 联网时按网络基准 + 单调时钟推进，断网时就是本机时间。
         这样用户中途修改系统时间也不会让显示跳错。 */
      displayNow: function () {
        if (this.syncState === 'network' && this.networkBaseMs) {
          var perf = (window.performance && window.performance.now)
            ? window.performance.now() : null;
          if (perf !== null && this.perfBaseMs) {
            return new Date(this.networkBaseMs + (perf - this.perfBaseMs));
          }
          return new Date(this.now.getTime() + this.offsetMs);
        }
        return this.now;
      },
      valid: function () { return !isNaN(this.displayNow.getTime()); },
      dateText: function () {
        if (!this.valid) { return '------ -- --'; }
        return [
          this.displayNow.getFullYear(),
          pad2(this.displayNow.getMonth() + 1),
          pad2(this.displayNow.getDate()),
        ].join('-');
      },
      weekText: function () {
        return this.valid ? WEEKDAYS[this.displayNow.getDay()] : '--';
      },
      timeText: function () {
        if (!this.valid) { return '--:--:--'; }
        return [
          pad2(this.displayNow.getHours()),
          pad2(this.displayNow.getMinutes()),
          pad2(this.displayNow.getSeconds()),
        ].join(':');
      },
      sourceLabel: function () {
        if (this.syncState === 'checking') { return '校时'; }
        return this.syncState === 'network' ? '网络' : '本机';
      },
      sourceClass: function () {
        return {
          'is-checking': this.syncState === 'checking',
          'is-network': this.syncState === 'network',
          'is-local': this.syncState === 'local',
        };
      },
      suspect: function () {
        if (!this.valid) { return true; }
        // 断网回退到本机时间：提醒用户自己确认时间是否正确
        if (this.syncState === 'local') { return true; }
        // 本机年份明显不合理（刚装机没设时间、CMOS 电池没电等）
        var year = this.now.getFullYear();
        if (year < 2000 || year > 2100) { return true; }
        // 网络时间正常，但本机时间跑偏：出入库记录仍会写错日期
        return this.syncState === 'network'
          && Math.abs(this.offsetSeconds) >= this.warnThresholdSeconds;
      },
      title: function () {
        if (this.syncState === 'checking') { return '正在校验网络时间…'; }
        if (this.syncState === 'local') {
          return '未能获取网络时间，当前显示本机时间；请确认电脑日期与时间是否正确';
        }
        if (Math.abs(this.offsetSeconds) >= this.warnThresholdSeconds) {
          return '本机时间与网络时间相差约 '
            + humanizeSeconds(this.offsetSeconds) + '，请校正电脑时间';
        }
        var host = (this.netSource || '').replace(/^https?:\/\//, '').replace(/\/$/, '');
        return '已使用网络时间' + (host ? '（' + host + '）' : '');
      },
    },
    methods: {
      refresh: function () {
        var wall = Date.now();
        var perf = (window.performance && window.performance.now)
          ? window.performance.now() : null;

        // 两次 tick 之间本机时钟被手动改动：网络基准仍有效，
        // 显示继续按单调时钟推进，同时立即重新校时确认。
        if (perf !== null && this.lastPerfMs !== null) {
          var expected = this.lastWallMs + (perf - this.lastPerfMs);
          if (Math.abs(wall - expected) > 5000) {
            this.now = new Date(wall);
            this.lastWallMs = wall;
            this.lastPerfMs = perf;
            this.requestTimeCheck();
            return;
          }
        }

        this.lastWallMs = wall;
        this.lastPerfMs = perf;
        this.now = new Date(wall);
      },

      requestTimeCheck: function () {
        var self = this;
        if (this.checkTimeout) { clearTimeout(this.checkTimeout); }
        // 兜底：后台线程即使被 DNS 卡住，界面也不会一直停在“校时”
        this.checkTimeout = setTimeout(function () {
          if (self.syncState === 'checking') {
            self.applyTimeResult({ networkOk: false, error: '网络校时请求超时' });
          }
        }, 12000);

        RMS.bridge.call('time.check').catch(function (err) {
          // 桥本身出错时也按“拿不到网络时间”处理
          self.applyTimeResult({ networkOk: false, error: err.message });
        });
      },

      applyTimeResult: function (data) {
        data = data || {};
        if (this.checkTimeout) {
          clearTimeout(this.checkTimeout);
          this.checkTimeout = null;
        }
        this.lastCheckAt = Date.now();
        this.warnThresholdSeconds = data.warnThresholdSeconds || 120;

        if (data.networkOk && data.networkEpochMs) {
          this.syncState = 'network';
          this.offsetMs = data.networkEpochMs - (data.systemEpochMs || 0);
          this.networkBaseMs = data.networkEpochMs;
          this.perfBaseMs = (window.performance && window.performance.now)
            ? window.performance.now() : 0;
          this.offsetSeconds = data.offsetSeconds || 0;
          this.netSource = data.source || '';
          this.netError = '';

          if (Math.abs(this.offsetSeconds) >= this.warnThresholdSeconds
              && !this.offsetWarned) {
            this.offsetWarned = true;
            this.showDetail = false;
            this.warn = {
              title: '系统时间需要校正',
              text: '本机时间与网络时间相差约 '
                + humanizeSeconds(this.offsetSeconds) + '。\n\n'
                + '网络时间：' + (data.networkTime || '—') + '\n'
                + '本机时间：' + (data.systemTime || '—') + '\n\n'
                + '出入库记录会自动使用网络时间，但仍建议在 Windows 的'
                + '“日期和时间”设置中校正系统时间，否则其他程序与文件时间会受影响。',
            };
          }
          return;
        }

        this.syncState = 'local';
        this.offsetMs = 0;
        this.offsetSeconds = 0;
        this.netSource = '';
        this.netError = data.error || '';

        if (!this.networkWarned) {
          this.networkWarned = true;
          this.showDetail = false;
          this.warn = {
            title: '网络时间获取失败',
            text: '未能获取网络时间，当前显示本机系统时间。\n\n'
              + '请确认电脑的日期与时间是否正确，否则出入库记录可能记错日期。',
            detail: this.netError ? shortenText(this.netError, 300) : '',
          };
        }
      },

      onFocusOrVisible: function () {
        this.refresh();
        // 最小化期间定时器可能被节流，回到前台时顺手补一次校时
        if (Date.now() - this.lastCheckAt > TIME_RECHECK_MS) {
          this.requestTimeCheck();
        }
      },
    },
    mounted: function () {
      var self = this;
      // 每秒刷新：时间被调整或秒针停住都能马上看出来
      this.timer = setInterval(this.refresh, 1000);
      this.recheckTimer = setInterval(function () {
        self.requestTimeCheck();
      }, TIME_RECHECK_MS);

      document.addEventListener('visibilitychange', this.onFocusOrVisible);
      window.addEventListener('focus', this.onFocusOrVisible);
      RMS.bridge.on('time.synced', this.applyTimeResult);

      this.refresh();
      this.requestTimeCheck();
    },
    beforeUnmount: function () {
      if (this.timer) { clearInterval(this.timer); }
      if (this.recheckTimer) { clearInterval(this.recheckTimer); }
      if (this.checkTimeout) { clearTimeout(this.checkTimeout); }
      document.removeEventListener('visibilitychange', this.onFocusOrVisible);
      window.removeEventListener('focus', this.onFocusOrVisible);
      // bridge.on 没有对应的 off；时钟是常驻组件，不会反复挂载
    },
  };

  function pad2(n) {
    return n < 10 ? '0' + n : String(n);
  }

  function humanizeSeconds(seconds) {
    var total = Math.abs(Math.round(seconds));
    if (total < 60) { return total + ' 秒'; }
    var minutes = Math.round(total / 60);
    if (minutes < 60) { return minutes + ' 分钟'; }
    var hours = Math.floor(minutes / 60);
    var rest = minutes % 60;
    return hours + ' 小时' + (rest ? ' ' + rest + ' 分钟' : '');
  }

  function shortenText(text, max) {
    var value = String(text || '').replace(/\s+/g, ' ').trim();
    return value.length > max ? value.slice(0, max) + '…' : value;
  }

  /* ============================================================
     空状态
     ============================================================ */
  components['ui-empty'] = {
    props: {
      title: { type: String, default: '暂无数据' },
      hint: { type: String, default: '' },
    },
    template: [
      '<div class="empty">',
      '  <div class="empty-title">{{ title }}</div>',
      '  <div v-if="hint" class="empty-hint">{{ hint }}</div>',
      '</div>',
    ].join(''),
  };

  /* ============================================================
     批号选择框
     ------------------------------------------------------------
     输入框 + 下拉候选。候选按当前试剂名称从后端拉取，便于没有 GS1
     条码时直接选批号；同时保留自由输入，入库遇到新批号不必先进名称表。
     `withStock` 打开时在候选右侧显示当前库存（出库用）。
     ============================================================ */
  components['ui-batch-picker'] = {
    props: {
      modelValue: { type: String, default: '' },
      name: { type: String, default: '' },
      withStock: Boolean,
      placeholder: { type: String, default: '输入或下拉选择批号' },
    },
    emits: ['update:modelValue', 'change'],
    template: [
      '<div class="combo" ref="root">',
      '  <div class="combo-field">',
      '    <input class="input combo-input" ref="input" :value="modelValue"',
      '           :placeholder="placeholder" autocomplete="off"',
      '           @input="onInput" @focus="onFocus"',
      '           @keydown.down.prevent="move(1)" @keydown.up.prevent="move(-1)"',
      '           @keydown.enter.prevent="onEnter" @keydown.esc="open = false"',
      '           @keydown.tab="open = false" />',
      '    <button type="button" class="combo-toggle" tabindex="-1"',
      '            @mousedown.prevent @click="toggle">&#9662;</button>',
      '  </div>',
      '  <div v-if="open" class="combo-menu">',
      '    <div v-for="(item, i) in filtered" :key="item.batch" class="combo-item"',
      '         :class="{ \'is-active\': i === active }"',
      '         @mousedown.prevent="pick(item)">',
      '      <span class="mono combo-batch">{{ item.batch }}</span>',
      '      <span v-if="withStock" class="combo-stock">库存 {{ item.quantity }}</span>',
      '    </div>',
      '    <div v-if="filtered.length === 0" class="combo-empty">{{ emptyHint }}</div>',
      '  </div>',
      '</div>',
    ].join(''),
    data: function () {
      return {
        items: [],
        open: false,
        active: -1,
        loading: false,
        typed: false,     // 只有用户手动输入时才按输入内容过滤
        loadedName: '',   // 已拉取过批号的名称，避免重复请求
        timer: null,
        seq: 0,
      };
    },
    computed: {
      filtered: function () {
        var q = this.typed ? (this.modelValue || '').trim().toLowerCase() : '';
        if (!q) { return this.items; }
        return this.items.filter(function (item) {
          return item.batch.toLowerCase().indexOf(q) >= 0;
        });
      },
      emptyHint: function () {
        if (!this.name) { return '请先选择试剂名称'; }
        if (this.loading) { return '加载中…'; }
        return '该名称下暂无批号，可直接输入新批号';
      },
    },
    watch: {
      name: function () {
        this.items = [];
        this.loadedName = '';
        this.typed = false;
        this.scheduleLoad();
      },
    },
    mounted: function () {
      document.addEventListener('mousedown', this.onDocDown);
      this.scheduleLoad();
    },
    beforeUnmount: function () {
      document.removeEventListener('mousedown', this.onDocDown);
      if (this.timer) { clearTimeout(this.timer); }
    },
    methods: {
      onDocDown: function (e) {
        if (this.$refs.root && !this.$refs.root.contains(e.target)) {
          this.open = false;
        }
      },
      scheduleLoad: function () {
        var self = this;
        if (this.timer) { clearTimeout(this.timer); }
        // 名称逐字输入时防抖，避免每敲一个字符就过一次桥
        this.timer = setTimeout(function () { self.load(); }, 180);
      },
      load: function () {
        var self = this;
        var name = (this.name || '').trim();
        if (!name) {
          this.items = [];
          this.loading = false;
          return;
        }

        var seq = ++this.seq;
        this.loading = true;
        RMS.bridge.call('batches.list', { name: name, withStock: !!(this.withStock) })
          .then(function (data) {
            if (seq !== self.seq) { return; }
            self.items = (data.items || []).map(function (item) {
              return typeof item === 'string'
                ? { batch: item, quantity: null }
                : item;
            });
            self.loadedName = name;
            self.active = -1;
          })
          .catch(function () {
            if (seq === self.seq) { self.items = []; }
          })
          .then(function () {
            if (seq === self.seq) { self.loading = false; }
          });
      },
      ensureLoaded: function () {
        var name = (this.name || '').trim();
        if (name && this.loadedName !== name && !this.loading) { this.load(); }
      },
      emitValue: function (value) {
        this.$emit('update:modelValue', value);
        this.$emit('change', value);
      },
      onInput: function (e) {
        this.typed = true;
        this.open = true;
        this.active = -1;
        this.emitValue(e.target.value);
      },
      onFocus: function () {
        this.open = true;
        this.typed = false;
        this.ensureLoaded();
      },
      toggle: function () {
        this.open = !this.open;
        if (this.open) {
          this.typed = false;
          this.active = -1;
          this.ensureLoaded();
        }
      },
      move: function (delta) {
        this.open = true;
        this.typed = false;
        var count = this.filtered.length;
        if (!count) { this.active = -1; return; }
        var next = this.active + delta;
        if (next < 0) { next = count - 1; }
        if (next >= count) { next = 0; }
        this.active = next;
      },
      onEnter: function () {
        if (this.open && this.active >= 0 && this.filtered[this.active]) {
          this.pick(this.filtered[this.active]);
          return;
        }
        this.open = false;
      },
      pick: function (item) {
        this.open = false;
        this.active = -1;
        this.typed = false;
        this.emitValue(item.batch);
      },
    },
  };

  /* ============================================================
     下拉选择
     ------------------------------------------------------------
     替代原生 <select>：原生弹层由浏览器/系统绘制，不同主题下
     配色不可控（深色主题曾出现浅底浅字）。这里复用批号选择器的
     下拉面板样式，保证所有下拉的观感一致。
     ============================================================ */
  components['ui-select'] = {
    props: {
      modelValue: { type: [String, Number], default: '' },
      options: { type: Array, default: function () { return []; } },
      placeholder: { type: String, default: '请选择' },
      dropup: Boolean,     // 靠近页面底部时向上展开
    },
    emits: ['update:modelValue', 'change'],
    template: [
      '<div class="select-combo" ref="root">',
      '  <button type="button" class="input select select-trigger"',
      '          :class="{ \'is-open\': open }"',
      '          :aria-expanded="open ? \'true\' : \'false\'" aria-haspopup="listbox"',
      '          @click="toggle"',
      '          @keydown.down.prevent="move(1)" @keydown.up.prevent="move(-1)"',
      '          @keydown.enter.prevent="onEnter" @keydown.space.prevent="onEnter"',
      '          @keydown.esc="close" @keydown.tab="close">',
      '    <span class="select-value">{{ currentLabel }}</span>',
      '    <span class="select-caret">&#9662;</span>',
      '  </button>',
      '  <div v-if="open" class="combo-menu select-menu" :class="{ \'is-dropup\': dropup }">',
      '    <div v-for="(opt, i) in options" :key="opt.value"',
      '         class="combo-item select-item"',
      '         :class="{ \'is-active\': i === active, \'is-selected\': isSelected(opt) }"',
      '         @mousedown.prevent="pick(opt)">',
      '      <span class="select-item-label">{{ opt.label }}</span>',
      '      <span v-if="isSelected(opt)" class="select-check">&#10003;</span>',
      '    </div>',
      '    <div v-if="!options.length" class="combo-empty">无可选项</div>',
      '  </div>',
      '</div>',
    ].join(''),
    data: function () {
      return { open: false, active: -1 };
    },
    computed: {
      currentLabel: function () {
        var value = String(this.modelValue);
        for (var i = 0; i < this.options.length; i++) {
          if (String(this.options[i].value) === value) {
            return this.options[i].label;
          }
        }
        return this.placeholder;
      },
    },
    mounted: function () {
      document.addEventListener('mousedown', this.onDocDown);
    },
    beforeUnmount: function () {
      document.removeEventListener('mousedown', this.onDocDown);
      this.raiseCard(false);
    },
    methods: {
      onDocDown: function (e) {
        if (this.$refs.root && !this.$refs.root.contains(e.target)) {
          this.close();
        }
      },
      /* 卡片皮肤带 backdrop-filter 时会形成独立层叠上下文，
         必须把整张卡片抬高，否则下拉会被后面的卡片盖住 */
      raiseCard: function (raised) {
        if (this._card) {
          this._card.classList.toggle('has-open-select', raised);
        }
        if (!raised) { return; }
        var el = this.$el;
        while (el && el !== document.body) {
          if (el.classList && el.classList.contains('card')) {
            this._card = el;
            el.classList.add('has-open-select');
            return;
          }
          el = el.parentNode;
        }
      },
      close: function () {
        this.open = false;
        this.raiseCard(false);
      },
      isSelected: function (opt) {
        return String(opt.value) === String(this.modelValue);
      },
      selectedIndex: function () {
        for (var i = 0; i < this.options.length; i++) {
          if (this.isSelected(this.options[i])) { return i; }
        }
        return -1;
      },
      toggle: function () {
        this.open = !this.open;
        if (this.open) {
          this.active = this.selectedIndex();
          this.raiseCard(true);
        } else {
          this.raiseCard(false);
        }
      },
      move: function (delta) {
        if (!this.open) {
          this.open = true;
          this.active = this.selectedIndex();
          this.raiseCard(true);
          return;
        }
        var count = this.options.length;
        if (!count) { return; }
        var next = this.active + delta;
        if (next < 0) { next = count - 1; }
        if (next >= count) { next = 0; }
        this.active = next;
      },
      onEnter: function () {
        if (this.open && this.active >= 0 && this.options[this.active]) {
          this.pick(this.options[this.active]);
          return;
        }
        this.open = !this.open;
        if (this.open) {
          this.active = this.selectedIndex();
          this.raiseCard(true);
        } else {
          this.raiseCard(false);
        }
      },
      pick: function (opt) {
        this.close();
        this.active = -1;
        this.$emit('update:modelValue', opt.value);
        this.$emit('change', opt.value);
      },
    },
  };

  /* ============================================================
     分页条
     ============================================================ */
  components['ui-pager'] = {
    components: { 'ui-select': components['ui-select'] },
    props: {
      page: { type: Number, default: 1 },
      pages: { type: Number, default: 1 },
      total: { type: Number, default: 0 },
      pageSize: { type: Number, default: 50 },
      pageSizes: { type: Array, default: function () { return [20, 50, 100, 200]; } },
      loading: Boolean,
    },
    emits: ['change', 'size-change'],
    template: [
      '<div class="pager">',
      '  <span>每页</span>',
      '  <ui-select class="pager-size" :model-value="pageSize"',
      '             :options="sizeOptions" dropup @change="onSize"></ui-select>',
      '  <span class="pager-info">共 {{ total }} 条，第 {{ page }} / {{ pages }} 页</span>',
      '  <span class="spacer"></span>',
      '  <span v-if="loading" class="muted" style="margin-right:12px">',
      '    <span class="spinner"></span> 加载中',
      '  </span>',
      '  <button class="btn btn-sm" :disabled="page <= 1" @click="go(1)">首页</button>',
      '  <button class="btn btn-sm" style="margin-left:6px" :disabled="page <= 1" @click="go(page - 1)">上一页</button>',
      '  <button class="btn btn-sm" style="margin-left:6px" :disabled="page >= pages" @click="go(page + 1)">下一页</button>',
      '  <button class="btn btn-sm" style="margin-left:6px" :disabled="page >= pages" @click="go(pages)">末页</button>',
      '</div>',
    ].join(''),
    computed: {
      sizeOptions: function () {
        return this.pageSizes.map(function (size) {
          return { value: size, label: String(size) };
        });
      },
    },
    methods: {
      go: function (p) {
        var target = Math.min(Math.max(1, p), this.pages);
        if (target !== this.page) { this.$emit('change', target); }
      },
      onSize: function (value) {
        this.$emit('size-change', parseInt(value, 10));
      },
    },
  };

  /* ============================================================
     弹窗
     ============================================================ */
  components['ui-modal'] = {
    props: {
      title: String,
      size: { type: String, default: '' },       // '' | 'lg' | 'xl'
      dismissible: { type: Boolean, default: true },
    },
    template: [
      '<div class="modal-scrim" @mousedown.self="onScrim">',
      '  <div class="modal" :class="sizeClass">',
      '    <div class="modal-header">',
      '      <div class="modal-title">{{ title }}</div>',
      '      <button v-if="dismissible" class="icon-btn modal-close" @click="close">&times;</button>',
      '    </div>',
      '    <div class="modal-body"><slot></slot></div>',
      '    <div class="modal-footer" v-if="$slots.footer"><slot name="footer"></slot></div>',
      '  </div>',
      '</div>',
    ].join(''),
    computed: {
      sizeClass: function () {
        return this.size ? ('modal-' + this.size) : '';
      },
    },
    methods: {
      close: function () { this.$emit('close'); },
      onScrim: function () {
        if (this.dismissible) { this.$emit('close'); }
      },
    },
    mounted: function () {
      // 打开弹窗时把焦点收进来，避免扫码枪仍打在背景输入框上
      var el = this.$el.querySelector('input, button');
      if (el) { el.focus(); }
    },
  };
}(window.RMS));
