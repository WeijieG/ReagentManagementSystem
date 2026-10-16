/* 出入库记录页。
 *
 * 旧实现（CombinedRecordPage.load_data）会一次性把全部入库+出库记录拉到内存里
 * 合并排序再塞进表格，数据一大就卡死。这里改为服务端 UNION ALL + 分页。
 */
window.RMS = window.RMS || {};
window.RMS.screens = window.RMS.screens || {};

(function (RMS) {
  'use strict';

  var bridge = RMS.bridge;
  var store = RMS.store;

  function isoDate(d) {
    var m = d.getMonth() + 1;
    var day = d.getDate();
    return d.getFullYear() + '-' + (m < 10 ? '0' + m : m) + '-' + (day < 10 ? '0' + day : day);
  }

  function daysAgo(n) {
    var d = new Date();
    d.setDate(d.getDate() - n);
    return isoDate(d);
  }

  function today() { return isoDate(new Date()); }

  function firstOfMonth() {
    var d = new Date();
    d.setDate(1);
    return isoDate(d);
  }

  var PRESETS = [
    { key: 'today', label: '今天',     range: function () { return [today(), today()]; } },
    { key: '7d',    label: '近 7 天',  range: function () { return [daysAgo(6), today()]; } },
    { key: '30d',   label: '近 30 天', range: function () { return [daysAgo(29), today()]; } },
    { key: 'month', label: '本月',     range: function () { return [firstOfMonth(), today()]; } },
    { key: 'all',   label: '全部',     range: function () { return ['', '']; } },
  ];

  RMS.screens['screen-records'] = {
    components: {
      'ui-pager': RMS.components['ui-pager'],
      'ui-empty': RMS.components['ui-empty'],
    },
    template: [
      '<div class="page">',

      '  <div class="page-toolbar">',
      '    <div class="scan-input" style="width:320px" :class="scanState">',
      '      <input class="input" v-model="scanText" placeholder="扫描条码筛选记录…"',
      '             @keyup.enter="onScan" />',
      '    </div>',
      '    <select class="input select" style="width:120px" v-model="filters.type"',
      '            @change="reload">',
      '      <option value="all">全部记录</option>',
      '      <option value="inbound">仅入库</option>',
      '      <option value="outbound">仅出库</option>',
      '    </select>',
      '    <input class="input" style="width:160px" v-model="filters.name"',
      '           placeholder="试剂名称" @input="onFilterChange" />',
      '    <input class="input" style="width:130px" v-model="filters.batch"',
      '           placeholder="批号" @input="onFilterChange" />',
      '    <button class="btn btn-primary" @click="exportXlsx">导出 Excel</button>',
      '  </div>',

      '  <div class="page-toolbar">',
      '    <span class="muted" style="margin-right:4px">日期范围</span>',
      '    <button v-for="p in presets" :key="p.key"',
      '            class="btn btn-sm" :class="{ \'btn-primary\': activePreset === p.key }"',
      '            style="margin-right:6px" @click="applyPreset(p)">{{ p.label }}</button>',
      '    <input type="date" class="input" style="width:150px" v-model="filters.start"',
      '           @change="onDateChange" />',
      '    <span class="muted">至</span>',
      '    <input type="date" class="input" style="width:150px" v-model="filters.end"',
      '           @change="onDateChange" />',
      '  </div>',

      '  <div class="card">',
      '    <div v-if="loading && rows.length === 0" class="loading-overlay">',
      '      <span class="spinner"></span> 正在加载…',
      '    </div>',
      '    <ui-empty v-else-if="rows.length === 0"',
      '              title="没有符合条件的记录"',
      '              hint="试试放宽日期范围或清除筛选条件"></ui-empty>',
      '    <div v-else class="table-wrap">',
      '      <table class="table">',
      '        <thead><tr>',
      '          <th>类型</th><th class="col-num">ID</th><th>试剂名称</th>',
      '          <th>批号</th><th>日期</th><th class="col-num">数量</th>',
      '          <th>操作员</th><th class="col-num">操作后库存</th>',
      '        </tr></thead>',
      '        <tbody>',
      '          <tr v-for="r in rows" :key="r.kind + r.id">',
      '            <td>',
      '              <span class="badge" :class="r.kind === \'入库\' ? \'badge-in\' : \'badge-out\'">',
      '                {{ r.kind }}</span>',
      '            </td>',
      '            <td class="col-num">{{ r.id }}</td>',
      '            <td>{{ r.name }}</td>',
      '            <td class="col-mono">{{ r.batch }}</td>',
      '            <td>{{ r.date }}</td>',
      '            <td class="col-num">{{ r.quantity }}</td>',
      '            <td>{{ r.operator || "—" }}</td>',
      '            <td class="col-num">{{ r.remaining }}</td>',
      '          </tr>',
      '        </tbody>',
      '      </table>',
      '    </div>',
      '    <ui-pager v-if="rows.length > 0"',
      '              :page="page" :pages="pages" :total="total" :page-size="pageSize"',
      '              :loading="loading"',
      '              @change="goPage" @size-change="changeSize"></ui-pager>',
      '  </div>',

      '</div>',
    ].join(''),
    data: function () {
      return {
        presets: PRESETS,
        activePreset: '30d',
        rows: [],
        total: 0,
        page: 1,
        pages: 1,
        pageSize: 50,
        loading: false,
        filters: { type: 'all', name: '', batch: '', start: daysAgo(29), end: today() },
        scanText: '',
        scanState: '',
        generation: 0,
      };
    },
    computed: {
      dirty: function () { return store.dirty; },
    },
    watch: {
      dirty: function () { this.reload(); },
    },
    mounted: function () {
      this.fetch();
    },
    methods: {
      reload: function () {
        this.page = 1;
        this.fetch();
      },
      goPage: function (p) { this.page = p; this.fetch(); },
      changeSize: function (s) { this.pageSize = s; this.page = 1; this.fetch(); },

      fetch: function () {
        var self = this;
        var gen = ++this.generation;
        this.loading = true;

        bridge.call('records.page', {
          page: this.page,
          pageSize: this.pageSize,
          filters: this.filters,
        }).then(function (data) {
          if (gen !== self.generation) { return; }
          self.rows = data.rows;
          self.total = data.total;
          self.pages = data.pages;
          self.page = data.page;
          self.loading = false;
        }).catch(function (err) {
          if (gen !== self.generation) { return; }
          self.loading = false;
          RMS.toast('error', '加载记录失败：' + err.message);
        });
      },

      onFilterChange: function () {
        var self = this;
        clearTimeout(this._t);
        this._t = setTimeout(function () { self.reload(); }, 250);
      },
      onDateChange: function () {
        this.activePreset = '';
        this.reload();
      },
      applyPreset: function (preset) {
        var range = preset.range();
        this.filters.start = range[0];
        this.filters.end = range[1];
        this.activePreset = preset.key;
        this.reload();
      },

      onScan: function () {
        var raw = (this.scanText || '').trim();
        if (!raw) { return; }
        var self = this;

        bridge.call('gs1.parse', { raw: raw }).then(function (res) {
          var hit = false;
          if (res.name) { self.filters.name = res.name; hit = true; }
          if (res.batch) { self.filters.batch = res.batch; hit = true; }
          if (!hit) {
            self.scanState = 'is-miss';
            RMS.toast('warning', '条码中没有可用的定位信息。');
            return;
          }
          self.scanState = 'is-hit';
          self.filters.start = '';
          self.filters.end = '';
          self.activePreset = 'all';
          self.reload();
          setTimeout(function () { self.scanState = ''; }, 1200);
        }).catch(function (err) {
          RMS.toast('error', '条码解析失败：' + err.message);
        });
      },

      exportXlsx: function () {
        var self = this;
        var start = this.filters.start;
        var end = this.filters.end;
        if (!start || !end) {
          RMS.toast('warning', '导出需要指定日期范围。');
          return;
        }
        RMS.toast('info', '正在生成 Excel…');
        bridge.call('records.export', { start: start, end: end }).then(function (res) {
          RMS.toast('success', '已导出到：' + res.path);
        }).catch(function (err) {
          if (err.code === 'CANCELLED') { return; }
          RMS.toast('error', '导出失败：' + err.message);
        });
      },
    },
  };
}(window.RMS));
