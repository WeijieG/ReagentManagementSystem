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
      '      <div class="stat-value" :class="{ \'is-warning\': data.expiringSoonCount > 0 }">',
      '        {{ data.expiringSoonCount || 0 }}</div>',
      '      <div class="stat-foot">需尽快安排使用</div>',
      '    </div>',
      '    <div class="stat-card">',
      '      <div class="stat-label">已过期</div>',
      '      <div class="stat-value" :class="{ \'is-danger\': data.expiredCount > 0 }">',
      '        {{ data.expiredCount || 0 }}</div>',
      '      <div class="stat-foot">建议尽快处理</div>',
      '    </div>',
      '  </div>',

      '  <div class="card" style="margin-bottom:20px">',
      '    <div class="card-header">',
      '      <div class="card-title">需关注批次</div>',
      '      <div class="muted" style="margin-left:auto">已过期与 30 天内到期</div>',
      '    </div>',
      '    <div v-if="loading" class="loading-overlay"><span class="spinner"></span> 正在加载…</div>',
      '    <ui-empty v-else-if="!data.attention || data.attention.length === 0"',
      '              title="没有需要关注的批次"',
      '              hint="所有在库试剂的有效期都在 30 天以外"></ui-empty>',
      '    <div v-else class="table-wrap">',
      '      <table class="table">',
      '        <thead><tr>',
      '          <th>试剂名称</th><th>批号</th><th>有效期至</th>',
      '          <th class="col-center">剩余有效期</th><th class="col-num">库存</th>',
      '        </tr></thead>',
      '        <tbody>',
      '          <tr v-for="(r, i) in data.attention" :key="i"',
      '              :class="{ \'is-expired\': r.state === \'expired\' }">',
      '            <td>{{ r.name }}</td>',
      '            <td class="col-mono">{{ r.batch }}</td>',
      '            <td>{{ r.expiryDate }}</td>',
      '            <td class="col-center">',
      '              <ui-expiry-badge :state="r.state" :days-left="r.daysLeft"></ui-expiry-badge>',
      '            </td>',
      '            <td class="col-num">{{ r.quantity }}</td>',
      '          </tr>',
      '        </tbody>',
      '      </table>',
      '    </div>',
      '  </div>',

      '  <div class="card">',
      '    <div class="card-header"><div class="card-title">最近操作</div></div>',
      '    <ui-empty v-if="!data.recent || data.recent.length === 0"',
      '              title="暂无出入库记录" hint="完成一次入库或出库后这里会显示"></ui-empty>',
      '    <div v-else class="table-wrap">',
      '      <table class="table">',
      '        <thead><tr>',
      '          <th>类型</th><th>试剂名称</th><th>批号</th>',
      '          <th>时间</th><th class="col-num">数量</th><th>操作员</th>',
      '        </tr></thead>',
      '        <tbody>',
      '          <tr v-for="(r, i) in data.recent" :key="i">',
      '            <td><span class="badge"',
      '                      :class="r.kind === \'入库\' ? \'badge-in\' : \'badge-out\'">',
      '              {{ r.kind }}</span></td>',
      '            <td>{{ r.name }}</td>',
      '            <td class="col-mono">{{ r.batch }}</td>',
      '            <td>{{ r.date }}</td>',
      '            <td class="col-num">{{ r.quantity }}</td>',
      '            <td>{{ r.operator || "—" }}</td>',
      '          </tr>',
      '        </tbody>',
      '      </table>',
      '    </div>',
      '  </div>',

      '</div>',
    ].join(''),
    data: function () {
      return { loading: true };
    },
    computed: {
      data: function () { return store.dashboard.data || {}; },
      dirty: function () { return store.dirty; },
    },
    watch: {
      dirty: function () { this.load(); },
    },
    mounted: function () { this.load(); },
    methods: {
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
