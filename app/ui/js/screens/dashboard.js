/* 概览页。
 *
 * 只重组既有数据，不新增采集：KPI 与清单全部由 reagents / inbound / outbound
 * 表现算，与旧版的口径一致。
 */
window.RMS = window.RMS || {};
window.RMS.screens = window.RMS.screens || {};

(function (RMS) {
  'use strict';

  var bridge = RMS.bridge;
  var store = RMS.store;

  RMS.screens['screen-dashboard'] = {
    components: {
      'ui-empty': RMS.components['ui-empty'],
      'ui-expiry-badge': RMS.components['ui-expiry-badge'],
    },
    template: [
      '<div class="page">',

      '  <div class="stat-grid">',
      '    <div class="stat-card">',
      '      <div class="stat-label">在库品规</div>',
      '      <div class="stat-value is-accent">{{ data.skuCount || 0 }}</div>',
      '      <div class="stat-foot">名称 + 批号组合数</div>',
      '    </div>',
      '    <div class="stat-card">',
      '      <div class="stat-label">库存总量</div>',
      '      <div class="stat-value">{{ data.totalQuantity || 0 }}</div>',
      '      <div class="stat-foot">全部在库记录之和</div>',
      '    </div>',
      '    <div class="stat-card">',
      '      <div class="stat-label">30 天内到期</div>',
      '      <div class="stat-value stat-value-lg"',
      '           :class="{ \'is-warning\': data.expiringSoonQuantity > 0 }">',
      '        {{ data.expiringSoonQuantity || 0 }}</div>',
      '      <div class="stat-foot">',
      '        <span class="stat-foot-num">{{ data.expiringSoonCount || 0 }}</span>',
      '        个批次 · 需尽快使用</div>',
      '    </div>',
      '    <div class="stat-card">',
      '      <div class="stat-label">已过期</div>',
      '      <div class="stat-value stat-value-lg"',
      '           :class="{ \'is-danger\': data.expiredQuantity > 0 }">',
      '        {{ data.expiredQuantity || 0 }}</div>',
      '      <div class="stat-foot">',
      '        <span class="stat-foot-num">{{ data.expiredCount || 0 }}</span>',
      '        个批次 · 建议尽快处理</div>',
      '    </div>',
      '  </div>',

      /* 左右两张卡：需关注批次 + 最近操作，各取前几条，展开走专门页面 */
      '  <div class="dashboard-grid">',

      '    <div class="card dash-card">',
      '      <div class="card-header dash-card-header">',
      '        <div class="card-title">需关注批次</div>',
      '        <div class="chip-group">',
      '          <button class="chip" :class="{ \'is-active\': attentionFilter === \'all\' }"',
      '                  @click="attentionFilter = \'all\'">全部 {{ attentionAll.length }}</button>',
      '          <button class="chip" :class="{ \'is-active\': attentionFilter === \'expired\' }"',
      '                  @click="attentionFilter = \'expired\'">已过期 {{ attentionExpired.length }}</button>',
      '          <button class="chip" :class="{ \'is-active\': attentionFilter === \'soon\' }"',
      '                  @click="attentionFilter = \'soon\'">30 天内 {{ attentionSoon.length }}</button>',
      '        </div>',
      '        <button class="link-btn" @click="gotoInventory">查看全部 →</button>',
      '      </div>',
      '      <div v-if="loading && !data.attention" class="loading-overlay">',
      '        <span class="spinner"></span> 正在加载…</div>',
      '      <ui-empty v-else-if="attentionRows.length === 0"',
      '                title="没有需要关注的批次"',
      '                hint="所选范围内暂无到期或临期试剂"></ui-empty>',
      '      <div v-else class="dash-list">',
      '        <div v-for="r in attentionRows" :key="r.name + r.batch" class="dash-item">',
      '          <div class="dash-item-main">',
      '            <div class="dash-item-title">{{ r.name }}</div>',
      '            <div class="dash-item-sub mono">{{ r.batch }}</div>',
      '          </div>',
      '          <ui-expiry-badge :state="r.state" :days-left="r.daysLeft"></ui-expiry-badge>',
      '          <div class="dash-item-qty">{{ r.quantity }}</div>',
      '        </div>',
      '      </div>',
      '    </div>',

      '    <div class="card dash-card">',
      '      <div class="card-header dash-card-header">',
      '        <div class="card-title">最近操作</div>',
      '        <button class="link-btn" @click="gotoRecords">查看全部 →</button>',
      '      </div>',
      '      <div v-if="loading && !data.recent" class="loading-overlay">',
      '        <span class="spinner"></span> 正在加载…</div>',
      '      <ui-empty v-else-if="recentRows.length === 0"',
      '                title="暂无出入库记录"',
      '                hint="完成一次入库或出库后这里会显示"></ui-empty>',
      '      <div v-else class="dash-list">',
      '        <div v-for="(r, i) in recentRows" :key="i" class="dash-item">',
      '          <span class="badge"',
      '                :class="r.kind === \'入库\' ? \'badge-in\' : \'badge-out\'">{{ r.kind }}</span>',
      '          <div class="dash-item-main">',
      '            <div class="dash-item-title">{{ r.name }}</div>',
      '            <div class="dash-item-sub">',
      '              <span class="mono">{{ r.batch }}</span> · {{ r.operator || "—" }}</div>',
      '          </div>',
      '          <div class="dash-item-time">{{ shortDate(r.date) }}</div>',
      '          <div class="dash-item-qty">{{ r.quantity }}</div>',
      '        </div>',
      '      </div>',
      '    </div>',

      '  </div>',

      '</div>',
    ].join(''),
    data: function () {
      return {
        loading: true,
        attentionFilter: 'all',   // all | expired | soon
      };
    },
    computed: {
      data: function () { return store.dashboard.data || {}; },
      dirty: function () { return store.dirty; },

      // 后端按「已过期 + 30 天内」返回，这里补上状态并按严重程度分档
      attentionAll: function () {
        return (this.data.attention || []).map(function (r) {
          var state = r.daysLeft < 0 ? 'expired' : 'warn30';
          return {
            name: r.name, batch: r.batch, quantity: r.quantity,
            daysLeft: r.daysLeft, state: state,
          };
        });
      },
      attentionExpired: function () {
        return this.attentionAll.filter(function (r) { return r.state === 'expired'; });
      },
      attentionSoon: function () {
        return this.attentionAll.filter(function (r) { return r.state !== 'expired'; });
      },
      attentionRows: function () {
        var rows = this.attentionFilter === 'expired' ? this.attentionExpired
                 : this.attentionFilter === 'soon' ? this.attentionSoon
                 : this.attentionAll;
        return rows.slice(0, 6);
      },
      recentRows: function () {
        return (this.data.recent || []).slice(0, 6);
      },
    },
    watch: {
      dirty: function () { this.load(); },
    },
    mounted: function () { this.load(); },
    methods: {
      /* ---------- 跳转到专门页面，并带上与首页一致的筛选条件 ---------- */
      gotoInventory: function () {
        var expiry = this.attentionFilter === 'expired' ? 'expired'
                   : this.attentionFilter === 'soon' ? 'soon'
                   : 'attention';
        store.pending.inventory = { name: '', batch: '', expiry: expiry };
        RMS.router.go('inventory');
      },
      gotoRecords: function () {
        // 首页的「最近操作」不限制日期与类型，对应记录页的「全部」
        store.pending.records = {
          type: 'all', name: '', batch: '', start: '', end: '', preset: 'all',
        };
        RMS.router.go('records');
      },
      // "yyyy-MM-dd HH:mm:ss:zzz" → "MM-dd HH:mm"
      shortDate: function (value) {
        if (!value || value.length < 16) { return value || ''; }
        return value.slice(5, 16);
      },

      load: function () {
        var self = this;
        self.loading = true;
        return bridge.call('dashboard.summary').then(function (data) {
          store.dashboard.data = data;
          self.loading = false;
        }).catch(function (err) {
          self.loading = false;
          RMS.toast('error', '加载概览失败：' + err.message);
        });
      },
    },
  };
}(window.RMS));
