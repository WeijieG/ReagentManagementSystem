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

  // key 必须与后端 RECORD_SORT_COLUMNS 的白名单一致
  var COLUMNS = [
    { key: '',          label: '类型',       sortable: false },
    { key: '',          label: 'ID',         cls: 'col-num', sortable: false },
    { key: 'name',      label: '试剂名称',   sortable: true },
    { key: 'batch',     label: '批号',       cls: 'col-mono', sortable: true },
    { key: 'date',      label: '日期',       sortable: true },
    { key: 'quantity',  label: '数量',       cls: 'col-num', sortable: true },
    { key: 'operator',  label: '操作员',     sortable: true },
    { key: 'remaining', label: '操作后库存', cls: 'col-num', sortable: true },
  ];

  RMS.screens['screen-records'] = {
    components: {
      'ui-pager': RMS.components['ui-pager'],
      'ui-select': RMS.components['ui-select'],
      'ui-empty': RMS.components['ui-empty'],
    },
    template: [
      '<div class="page page-fill">',

      /* 筛选栏按窗口宽度分档排布，避免控件零散换行 */
      '  <div class="filter-toolbar filter-toolbar-records">',
      '    <div class="scan-input filter-scan" :class="scanState">',
      '      <input class="input" v-model="scanText" placeholder="扫描条码筛选记录…"',
      '             @keyup.enter="onScan" />',
      '    </div>',
      '    <ui-select class="filter-type" :model-value="filters.type"',
      '               :options="typeOptions" @change="onTypeChange"></ui-select>',
      '    <input class="input filter-name" v-model="filters.name"',
      '           placeholder="试剂名称" @input="onFilterChange" />',
      '    <input class="input filter-batch" v-model="filters.batch"',
      '           placeholder="批号" @input="onFilterChange" />',
      '    <button class="btn btn-primary filter-export" @click="exportXlsx">导出 Excel</button>',
      '  </div>',

      '  <div class="page-toolbar date-toolbar">',
      '    <div class="date-presets">',
      '      <span class="muted date-label">日期范围</span>',
      '      <button v-for="p in presets" :key="p.key"',
      '              class="btn date-preset" :class="{ \'btn-primary\': activePreset === p.key }"',
      '              @click="applyPreset(p)">{{ p.label }}</button>',
      '    </div>',
      '    <div class="date-range">',
      '      <input type="date" class="input date-input" v-model="filters.start"',
      '             @change="onDateChange" />',
      '      <span class="muted date-sep">至</span>',
      '      <input type="date" class="input date-input" v-model="filters.end"',
      '             @change="onDateChange" />',
      '    </div>',
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
      '          <th v-for="(c, i) in columns" :key="i"',
      '              :class="[c.cls, { \'is-sortable\': c.sortable,',
      '                               \'is-sorted\': c.sortable && sort.col === c.key }]"',
      '              @click="c.sortable && toggleSort(c.key)">',
      '            {{ c.label }}',
      '            <span v-if="c.sortable" class="table-sort-mark">{{ sortMark(c.key) }}</span>',
      '          </th>',
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
        columns: COLUMNS,
        typeOptions: [
          { value: 'all', label: '全部记录' },
          { value: 'inbound', label: '仅入库' },
          { value: 'outbound', label: '仅出库' },
        ],
        activePreset: '30d',
        rows: [],
        total: 0,
        page: 1,
        pages: 1,
        pageSize: 50,
        loading: false,
        filters: { type: 'all', name: '', batch: '', start: daysAgo(29), end: today() },
        sort: { col: 'date', dir: 'desc' },
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
      // 从概览「最近操作」跳过来时，展示与首页一致的全部记录条件
      var pending = store.pending && store.pending.records;
      if (pending) {
        this.filters = {
          type: pending.type || 'all',
          name: pending.name || '',
          batch: pending.batch || '',
          start: pending.start || '',
          end: pending.end || '',
        };
        this.activePreset = pending.preset || '';
        store.pending.records = null;
      }
      this.fetch();
    },
    methods: {
      reload: function () {
        this.page = 1;
        this.fetch();
      },
      goPage: function (p) { this.page = p; this.fetch(); },
      changeSize: function (s) { this.pageSize = s; this.page = 1; this.fetch(); },

      toggleSort: function (key) {
        if (this.sort.col === key) {
          this.sort.dir = this.sort.dir === 'asc' ? 'desc' : 'asc';
        } else {
          this.sort.col = key;
          this.sort.dir = 'asc';
        }
        this.page = 1;
        this.fetch();
      },
      sortMark: function (key) {
        if (this.sort.col !== key) { return '↕'; }
        return this.sort.dir === 'asc' ? '↑' : '↓';
      },

      fetch: function () {
        var self = this;
        var gen = ++this.generation;
        this.loading = true;

        bridge.call('records.page', {
          page: this.page,
          pageSize: this.pageSize,
          filters: this.filters,
          sort: this.sort,
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
      onTypeChange: function (value) {
        this.filters.type = value;
        this.reload();
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
