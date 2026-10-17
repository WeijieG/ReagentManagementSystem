/* 设置页。原先是一个 300x350 的「更多功能」弹窗，现在拆成带分组的整页。 */
window.RMS = window.RMS || {};
window.RMS.screens = window.RMS.screens || {};

(function (RMS) {
  'use strict';

  var bridge = RMS.bridge;
  var store = RMS.store;

  RMS.screens['screen-settings'] = {
    components: {
      'ui-modal': RMS.components['ui-modal'],
      'ui-select': RMS.components['ui-select'],
    },
    template: [
      '<div class="page page-narrow">',

      /* 导入后必须重启才生效 —— 当前连接里读的还是旧库 */
      '  <div v-if="needRestart" class="card" style="margin-bottom:20px;border-color:var(--warning)">',
      '    <div class="card-body" style="display:flex;align-items:center">',
      '      <div style="flex:1">',
      '        <div style="font-weight:600;margin-bottom:4px">数据库已导入</div>',
      '        <div class="muted">需要重启程序，新数据才会生效。</div>',
      '      </div>',
      '      <button class="btn btn-primary" @click="restart">立即重启</button>',
      '    </div>',
      '  </div>',

      /* ---- 外观 ---- */
      '  <div class="card" style="margin-bottom:20px">',
      '    <div class="card-header"><div class="card-title">外观</div></div>',
      '    <div class="card-body">',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">界面风格</div>',
      '          <div class="setting-desc">',
      '            明暗已并入风格，选择后立即生效',
      '          </div>',
      '        </div>',
      '        <ui-select class="setting-select" :model-value="uiStyle"',
      '                   :options="styles" @change="setStyle"></ui-select>',
      '      </div>',
      '    </div>',
      '  </div>',

      /* ---- 快捷键 ---- */
      '  <div class="card" style="margin-bottom:20px">',
      '    <div class="card-header"><div class="card-title">快捷键</div></div>',
      '    <div class="card-body">',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">入库</div>',
      '          <div class="setting-desc">',
      '            聚焦后按下要使用的按键；支持字母、数字与 F1–F12',
      '          </div>',
      '        </div>',
      '        <input class="input shortcut-input" readonly',
      '               :value="capturing === \'inbound\' ? \'按下按键…\' : (shortcuts.inbound || \'未设置\')"',
      '               @focus="startCapture(\'inbound\')"',
      '               @blur="endCapture"',
      '               @keydown="onShortcutKey($event, \'inbound\')" />',
      '        <button class="btn btn-sm" style="margin-left:8px"',
      '                @click="clearShortcut(\'inbound\')">清除</button>',
      '      </div>',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">出库</div>',
      '          <div class="setting-desc">',
      '            出入库界面打开时快捷键自动失效；按 Esc 可关闭界面',
      '          </div>',
      '        </div>',
      '        <input class="input shortcut-input" readonly',
      '               :value="capturing === \'outbound\' ? \'按下按键…\' : (shortcuts.outbound || \'未设置\')"',
      '               @focus="startCapture(\'outbound\')"',
      '               @blur="endCapture"',
      '               @keydown="onShortcutKey($event, \'outbound\')" />',
      '        <button class="btn btn-sm" style="margin-left:8px"',
      '                @click="clearShortcut(\'outbound\')">清除</button>',
      '      </div>',
      '    </div>',
      '  </div>',

      /* ---- 数据 ---- */
      '  <div class="card" style="margin-bottom:20px">',
      '    <div class="card-header"><div class="card-title">数据</div></div>',
      '    <div class="card-body">',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">备份数据库</div>',
      '          <div class="setting-desc">导出一份完整的数据库副本到指定位置</div>',
      '        </div>',
      '        <button class="btn" :disabled="busy" @click="backup">立即备份</button>',
      '      </div>',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">导入数据库</div>',
      '          <div class="setting-desc">',
      '            用备份文件替换当前数据。原库会自动备份到 Backups 目录，操作后需要重启。',
      '          </div>',
      '        </div>',
      '        <button class="btn" :disabled="busy" @click="importDb">选择文件</button>',
      '      </div>',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">导出出入库记录</div>',
      '          <div class="setting-desc">按日期范围导出为 Excel，可在「出入库记录」页做更细的筛选</div>',
      '        </div>',
      '        <button class="btn" @click="$emit(\'goto-records\')">前往导出</button>',
      '      </div>',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">数据库位置</div>',
      '          <div class="setting-desc mono" style="word-break:break-all">{{ dbPath }}</div>',
      '        </div>',
      '      </div>',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">数据一致性</div>',
      '          <div class="setting-desc">',
      '            <span v-if="dbIssues.length" style="color:var(--danger)">',
      '              发现 {{ dbIssues.length }} 条记录不一致，相关记录的改名操作会失败；建议先备份数据库。',
      '            </span>',
      '            <span v-else-if="dbRepaired.length">',
      '              启动时已自动修复 {{ dbRepaired.length }} 个缺失的试剂名称，当前无异常。',
      '            </span>',
      '            <span v-else>正常，未发现记录不一致。</span>',
      '          </div>',
      '        </div>',
      '      </div>',
      '    </div>',
      '  </div>',

      /* ---- 更新 ---- */
      '  <div class="card" style="margin-bottom:20px">',
      '    <div class="card-header"><div class="card-title">更新</div></div>',
      '    <div class="card-body">',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">当前版本</div>',
      '          <div class="setting-desc">{{ version || "未知" }}</div>',
      '        </div>',
      '        <button class="btn" :disabled="checking || !networkOk" @click="checkUpdate">',
      '          {{ checking ? "检查中…" : "检查更新" }}</button>',
      '      </div>',
      '      <div v-if="!networkOk" class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label" style="color:var(--warning)">网络组件不可用</div>',
      '          <div class="setting-desc">',
      '            本机缺少可用的网络运行库，因此无法检查更新。',
      '            库存管理、出入库登记、数据导出等核心功能不受影响。',
      '          </div>',
      '          <div class="setting-desc mono" style="margin-top:4px;word-break:break-all">{{ networkError }}</div>',
      '        </div>',
      '      </div>',
      '      <div v-if="updateInfo" class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">发现新版本 {{ updateInfo.version }}</div>',
      '          <div class="setting-desc" style="white-space:pre-wrap">{{ updateInfo.changelog || "无更新说明" }}</div>',
      '          <div v-if="downloading" class="progress" style="margin-top:10px;max-width:320px">',
      '            <div class="progress-bar" :style="{ width: percent + \'%\' }"></div>',
      '          </div>',
      '        </div>',
      '        <button class="btn btn-primary" :disabled="downloading" @click="startUpdate">',
      '          {{ downloading ? percent + "%" : "立即更新" }}</button>',
      '      </div>',
      '    </div>',
      '  </div>',

      /* ---- 关于 ---- */
      '  <div class="card">',
      '    <div class="card-header"><div class="card-title">关于</div></div>',
      '    <div class="card-body">',
      '      <div class="setting-row">',
      '        <div class="setting-main">',
      '          <div class="setting-label">试剂库存管理系统</div>',
      '          <div class="setting-desc">版本 {{ version || "未知" }}</div>',
      '        </div>',
      '        <button class="btn" @click="showAbout = true">开源声明</button>',
      '      </div>',
      '    </div>',
      '  </div>',

      '  <ui-modal v-if="showAbout" title="开源声明" size="lg" @close="showAbout = false">',
      '    <div class="legal-text">{{ aboutText }}</div>',
      '  </ui-modal>',

      '</div>',
    ].join(''),
    data: function () {
      return {
        styles: RMS.styles,
        busy: false,
        checking: false,
        downloading: false,
        percent: 0,
        updateInfo: null,
        needRestart: false,
        showAbout: false,
        aboutText: '',
        capturing: '',   // 正在录入快捷键的行：inbound / outbound
      };
    },
    computed: {
      uiStyle: function () { return store.style; },
      shortcuts: function () { return store.shortcuts; },
      version: function () { return store.version; },
      dbPath: function () { return store.dbPath; },
      dbIssues: function () { return store.dbIssues; },
      dbRepaired: function () { return store.dbRepaired; },
      networkOk: function () { return store.networkOk; },
      networkError: function () { return store.networkError; },
    },
    created: function () {
      var self = this;
      bridge.call('about.text').then(function (res) {
        self.aboutText = res.text;
      }).catch(function () { self.aboutText = '无法加载说明文本。'; });

      // 下载进度由后端推送
      bridge.on('update.progress', function (data) {
        if (!self.downloading) { return; }
        self.percent = data.percent || 0;
      });
      bridge.on('update.available', function (data) {
        self.updateInfo = data;
      });
      bridge.on('update.error', function (data) {
        self.downloading = false;
        self.percent = 0;
        RMS.toast('error', '更新失败：' + (data.message || '未知错误'));
      });
      bridge.on('update.installing', function () {
        RMS.toast('info', '正在启动安装程序，程序即将退出…');
      });
    },
    methods: {
      setStyle: function (style) { RMS.setStyle(style); },

      /* ---------- 快捷键录入 ---------- */
      startCapture: function (which) { this.capturing = which; },
      endCapture: function () { this.capturing = ''; },
      onShortcutKey: function (e, which) {
        e.preventDefault();
        e.stopPropagation();

        if (e.key === 'Escape') {
          this.capturing = '';
          e.target.blur();
          return;
        }
        if (e.key === 'Backspace' || e.key === 'Delete') {
          this.capturing = '';
          e.target.blur();
          RMS.setShortcut(which, '');
          return;
        }

        var key = RMS.shortcutFromEvent(e);
        if (!key) {
          RMS.toast('warning', '该按键不可用，请使用字母、数字或 F1–F12');
          return;
        }
        var other = which === 'inbound' ? 'outbound' : 'inbound';
        if (store.shortcuts[other] === key) {
          RMS.toast('warning', '该按键已用于' + (other === 'inbound' ? '入库' : '出库'));
          return;
        }

        this.capturing = '';
        e.target.blur();
        RMS.setShortcut(which, key);
      },
      clearShortcut: function (which) { RMS.setShortcut(which, ''); },

      backup: function () {
        var self = this;
        this.busy = true;
        bridge.call('data.backup').then(function (res) {
          self.busy = false;
          if (res.cancelled) { return; }
          RMS.toast('success', '数据库已备份到：' + res.path);
        }).catch(function (err) {
          self.busy = false;
          RMS.toast('error', '备份失败：' + err.message);
        });
      },

      importDb: function () {
        var self = this;
        this.busy = true;
        bridge.call('data.import').then(function (res) {
          self.busy = false;
          if (res.cancelled) { return; }
          // 导入是覆盖数据库文件，当前连接里的还是旧数据，必须重启才生效
          self.needRestart = true;
          RMS.toast('success', '已导入。需要重启程序使新数据生效。');
        }).catch(function (err) {
          self.busy = false;
          RMS.toast('error', '导入失败：' + err.message);
        });
      },

      restart: function () {
        RMS.toast('info', '正在重启…');
        bridge.call('data.restart');
      },

      checkUpdate: function () {
        var self = this;
        this.checking = true;
        bridge.call('update.check').then(function (res) {
          self.checking = false;
          if (!res.available) {
            RMS.toast('info', '当前已是最新版本。');
            return;
          }
          self.updateInfo = res.info;
        }).catch(function (err) {
          self.checking = false;
          RMS.toast('error', '检查更新失败：' + err.message);
        });
      },

      startUpdate: function () {
        var self = this;
        this.downloading = true;
        this.percent = 0;
        bridge.call('update.start').then(function () {
          // 实际引入由 update.progress / update.installing 推送驱动
        }).catch(function (err) {
          self.downloading = false;
          RMS.toast('error', '更新失败：' + err.message);
        });
      },
    },
  };
}(window.RMS));
