/* 库存管理页 —— 主工作流界面。
 *
 * 分页与排序都在服务端做（LIMIT/OFFSET + 白名单列），
 * 过桥只传当页几十行，绝不过桥传全表。
 */
window.RMS = window.RMS || {};
window.RMS.screens = window.RMS.screens || {};

(function (RMS) {
  'use strict';

  var bridge = RMS.bridge;
  var store = RMS.store;

  // 列定义：key 必须与后端 INVENTORY_SORT_COLUMNS 的白名单一致
  var COLUMNS = [
    { key: 'id',             label: 'ID',       cls: 'col-num',   sortable: true,  width: 64 },
    { key: 'name',           label: '试剂名称', sortable: true },
    { key: 'batch',          label: '批号',     cls: 'col-mono',  sortable: true },
    { key: 'productionDate', label: '生产日期', sortable: true },
    { key: 'expiryDate',     label: '有效期至', sortable: true },
    { key: 'quantity',       label: '库存数量', cls: 'col-num',   sortable: true },
    { key: 'daysLeft',       label: '剩余有效期', cls: 'col-center', sortable: true },
    { key: 'gtin',           label: 'GTIN',     cls: 'col-mono',  sortable: false },
  ];

  var EXPIRY_OPTIONS = [
    { value: 'all',       label: '全部' },
    { value: 'attention', label: '需关注' },
    { value: 'expired',   label: '已过期' },
    { value: 'soon',      label: '30 天内到期' },
    { value: 'near',      label: '90 天内到期' },
    { value: 'valid',     label: '正常' },
  ];

  RMS.screens['screen-inventory'] = {
    components: {
      'ui-pager': RMS.components['ui-pager'],
      'ui-select': RMS.components['ui-select'],
      'ui-empty': RMS.components['ui-empty'],
      'ui-expiry-badge': RMS.components['ui-expiry-badge'],
    },
    template: [
      '<div class="page page-fill">',

      /* 筛选栏按窗口宽度分档排布；出入库统一走顶栏，这里不放动作按钮 */
      '  <div class="filter-toolbar filter-toolbar-inventory">',
      '    <div class="scan-input filter-scan" :class="scanState">',
      '      <input class="input" ref="scan" v-model="scanText"',
      '             placeholder="扫描条码定位试剂…" @keyup.enter="onScan" />',
      '    </div>',
      '    <input class="input filter-name" v-model="filters.name"',
      '           placeholder="试剂名称" @input="onFilterChange" />',
      '    <input class="input filter-batch" v-model="filters.batch"',
      '           placeholder="批号" @input="onFilterChange" />',
      '    <ui-select class="filter-expiry" :model-value="filters.expiry"',
      '               :options="expiryOptions" @change="onExpiryChange"></ui-select>',
      '    <button class="btn filter-reset" @click="resetFilters">重置</button>',
      '  </div>',

      '  <div class="card">',
      '    <div v-if="loading && rows.length === 0" class="loading-overlay">',
      '      <span class="spinner"></span> 正在加载…',
      '    </div>',

      '    <ui-empty v-else-if="rows.length === 0"',
      '              title="没有符合条件的库存记录"',
      '              :hint="emptyHint"></ui-empty>',

      '    <div v-else class="table-wrap">',
      '      <table class="table">',
      '        <thead><tr>',
      '          <th v-for="c in columns" :key="c.key"',
      '              :class="[c.cls, { \'is-sortable\': c.sortable,',
      '                               \'is-sorted\': sort.col === c.key }]"',
      '              @click="c.sortable && toggleSort(c.key)">',
      '            {{ c.label }}',
      '            <span v-if="c.sortable" class="table-sort-mark">{{ sortMark(c.key) }}</span>',
      '          </th>',
      '        </tr></thead>',
      '        <tbody>',
      '          <tr v-for="r in rows" :key="r.id"',
      '              :class="{ \'is-expired\': r.state === \'expired\' }">',
      '            <td class="col-num">{{ r.id }}</td>',
      '            <td>{{ r.name }}</td>',
      '            <td class="col-mono">{{ r.batch }}</td>',
      '            <td>{{ r.productionDate }}</td>',
      '            <td>{{ r.expiryDate }}</td>',
      '            <td class="col-num">{{ r.quantity }}</td>',
      '            <td class="col-center">',
      '              <ui-expiry-badge :state="r.state" :days-left="r.daysLeft"></ui-expiry-badge>',
      '            </td>',
      '            <td class="col-mono">{{ r.gtin || "—" }}</td>',
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
        columns: COLUMNS,
        expiryOptions: EXPIRY_OPTIONS,
        rows: [],
        total: 0,
        page: 1,
        pages: 1,
        pageSize: 50,
        loading: false,
        filters: { name: '', batch: '', expiry: 'all' },
        sort: { col: 'id', dir: 'asc' },
        scanText: '',
        scanState: '',
        // 请求代际号：只接受最新一次请求的结果，避免快速输入时旧响应覆盖新数据
        generation: 0,
      };
    },
    computed: {
      emptyHint: function () {
        return (this.filters.name || this.filters.batch || this.filters.expiry !== 'all')
          ? '试试放宽筛选条件'
          : '还没有库存记录，先用「入库」添加试剂';
      },
      dirty: function () { return store.dirty; },
    },
    watch: {
      dirty: function () { this.reload(); },
    },
    mounted: function () {
      // 从概览跳过来时带上一次性筛选条件（需关注批次 → 库存）
      var pending = store.pending && store.pending.inventory;
      if (pending) {
        this.filters = {
          name: pending.name || '',
          batch: pending.batch || '',
          expiry: pending.expiry || 'all',
        };
        store.pending.inventory = null;
      }
      this.reload();
    },
    methods: {
      /* ---------- 取数 ---------- */
      reload: function () {
        this.page = 1;
        this.fetch();
      },
      goPage: function (p) {
        this.page = p;
        this.fetch();
      },
      changeSize: function (size) {
        this.pageSize = size;
        this.page = 1;
        this.fetch();
      },
      fetch: function () {
        var self = this;
        var gen = ++this.generation;
        this.loading = true;

        bridge.call('inventory.page', {
          page: this.page,
          pageSize: this.pageSize,
          filters: this.filters,
          sort: this.sort,
        }).then(function (data) {
          if (gen !== self.generation) { return; }   // 已被更新的请求取代
          self.rows = data.rows;
          self.total = data.total;
          self.pages = data.pages;
          self.page = data.page;
          self.loading = false;
        }).catch(function (err) {
          if (gen !== self.generation) { return; }
          self.loading = false;
          RMS.toast('error', '加载库存失败：' + err.message);
        });
      },

      /* ---------- 筛选 ---------- */
      onFilterChange: function () {
        // 简单防抖：连续输入时只在停顿后发一次请求
        var self = this;
        clearTimeout(this._filterTimer);
        this._filterTimer = setTimeout(function () { self.reload(); }, 250);
      },
      onExpiryChange: function (value) {
        this.filters.expiry = value;
        this.reload();
      },
      resetFilters: function () {
        this.filters = { name: '', batch: '', expiry: 'all' };
        this.sort = { col: 'id', dir: 'asc' };
        this.reload();
      },

      /* ---------- 排序 ---------- */
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

      /* ---------- 扫码 ---------- */
      onScan: function () {
        var raw = (this.scanText || '').trim();
        if (!raw) { return; }
        var self = this;

        bridge.call('gs1.parse', { raw: raw }).then(function (res) {
          if (!res.recognised) {
            self.scanState = 'is-miss';
            RMS.toast('warning', '无法识别该条码。若扫码枪输出的是原始 GS1 格式，请确认已升级到支持 FNC1 的版本。');
            return;
          }

          var hit = false;
          if (res.name) { self.filters.name = res.name; hit = true; }
          if (res.batch) { self.filters.batch = res.batch; hit = true; }

          if (!hit && res.gtin) {
            self.scanState = 'is-miss';
            RMS.toast('warning', '该 GTIN 尚未绑定试剂名称，可到「试剂名称」页绑定。');
            return;
          }

          if (!hit) {
            self.scanState = 'is-miss';
            RMS.toast('warning', '条码中没有可用的定位信息。');
            return;
          }

          self.scanState = 'is-hit';
          self.reload();
          setTimeout(function () { self.scanState = ''; }, 1200);
        }).catch(function (err) {
          self.scanState = 'is-miss';
          RMS.toast('error', '条码解析失败：' + err.message);
        });
      },
    },
  };
}(window.RMS));
