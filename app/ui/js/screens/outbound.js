/* 出库抽屉。与入库的差别：实时显示当前库存，且数量不得超过库存。 */
window.RMS = window.RMS || {};
window.RMS.screens = window.RMS.screens || {};

(function (RMS) {
  'use strict';

  var bridge = RMS.bridge;

  RMS.components['outbound-drawer'] = {
    template: [
      '<div class="drawer-scrim" @mousedown.self="close">',
      '  <div class="drawer">',
      '    <div class="modal-header">',
      '      <div class="modal-title">出库登记</div>',
      '      <button class="icon-btn modal-close" @click="close">&times;</button>',
      '    </div>',

      '    <div class="modal-body" style="display:flex">',

      '      <div style="width:220px;flex-shrink:0;margin-right:20px">',
      '        <div class="field-label">试剂名称列表</div>',
      '        <input class="input" style="margin-bottom:8px" v-model="nameQuery"',
      '               placeholder="筛选…" />',
      '        <div style="max-height:420px;overflow-y:auto;border:1px solid var(--border);border-radius:var(--radius-sm)">',
      '          <div v-for="n in filteredNames" :key="n"',
      '               style="padding:8px 12px;cursor:pointer;font-size:var(--text-sm);border-bottom:1px solid var(--border-subtle)"',
      '               @click="pickName(n)">{{ n }}</div>',
      '          <div v-if="filteredNames.length === 0" class="muted"',
      '               style="padding:12px;font-size:var(--text-sm)">无匹配名称</div>',
      '        </div>',
      '      </div>',

      '      <div style="flex:1;min-width:0">',
      '        <div class="field">',
      '          <label class="field-label">条码扫描</label>',
      '          <div class="scan-input" :class="scanState">',
      '            <input class="input" ref="scan" v-model="scanText"',
      '                   placeholder="扫描试剂条码（GS1 标准）" @keyup.enter="onScan" />',
      '          </div>',
      '          <div v-if="gtin" class="field-hint">GTIN：<span class="mono">{{ gtin }}</span></div>',
      '        </div>',

      '        <div class="field">',
      '          <label class="field-label">试剂名称 <span style="color:var(--danger)">*</span></label>',
      '          <input class="input" v-model="form.name" placeholder="输入或从左侧选择" />',
      '        </div>',

      '        <div class="field">',
      '          <label class="field-label">批号 <span style="color:var(--danger)">*</span></label>',
      '          <input class="input" list="bl-out" v-model="form.batch"',
      '                 placeholder="输入或扫码带出" @change="refreshStock" @input="refreshStock" />',
      '          <datalist id="bl-out"><option v-for="b in batches" :key="b" :value="b"></option></datalist>',
      '        </div>',

      '        <div class="field">',
      '          <label class="field-label">当前库存</label>',
      '          <div style="display:flex;align-items:center">',
      '            <span class="stat-value" style="font-size:var(--text-xl)"',
      '                  :class="{ \'is-danger\': stock === 0 && form.batch }">{{ stockLabel }}</span>',
      '            <span v-if="form.batch && stock === 0" class="muted" style="margin-left:10px">',
      '              该批号无可用库存</span>',
      '          </div>',
      '        </div>',

      '        <div style="display:flex">',
      '          <div class="field" style="flex:1;margin-right:12px">',
      '            <label class="field-label">出库数量 <span style="color:var(--danger)">*</span></label>',
      '            <input class="input" v-model="form.quantity" placeholder="正整数"',
      '                   @keyup.enter="submit" />',
      '          </div>',
      '          <div class="field" style="flex:1">',
      '            <label class="field-label">操作员</label>',
      '            <input class="input" v-model="form.operator" placeholder="姓名"',
      '                   @keyup.enter="submit" />',
      '          </div>',
      '        </div>',

      '        <div v-if="error" class="field-error" style="margin-bottom:12px">{{ error }}</div>',

      '        <div style="display:flex;align-items:center">',
      '          <button class="btn btn-primary btn-lg" :disabled="busy" @click="submit">',
      '            {{ busy ? "提交中…" : "确定出库" }}</button>',
      '          <span class="muted" style="margin-left:12px;font-size:var(--text-sm)">',
      '            提交后表单会保留，可直接扫下一件</span>',
      '          <span class="spacer" style="flex:1"></span>',
      '          <button class="btn" @click="close">关闭</button>',
      '        </div>',
      '      </div>',

      '    </div>',
      '  </div>',
      '</div>',
    ].join(''),
    data: function () {
      return {
        names: [],
        nameQuery: '',
        batches: [],
        stock: null,
        scanText: '',
        scanState: '',
        gtin: '',
        busy: false,
        error: '',
        form: { name: '', batch: '', quantity: '', operator: '' },
      };
    },
    computed: {
      filteredNames: function () {
        var q = this.nameQuery.trim().toLowerCase();
        if (!q) { return this.names; }
        return this.names.filter(function (n) { return n.toLowerCase().indexOf(q) >= 0; });
      },
      stockLabel: function () {
        return this.stock === null ? '—' : String(this.stock);
      },
    },
    mounted: function () {
      this.loadNames();
      this.focusScan();
    },
    methods: {
      close: function () { this.$emit('close'); },
      focusScan: function () {
        var self = this;
        this.$nextTick(function () {
          if (self.$refs.scan) { self.$refs.scan.focus(); }
        });
      },
      loadNames: function () {
        var self = this;
        bridge.call('names.list').then(function (data) {
          self.names = data.rows.map(function (r) { return r.name; });
        });
      },
      pickName: function (n) {
        this.form.name = n;
        this.loadBatches();
        this.refreshStock();
      },
      loadBatches: function () {
        var self = this;
        if (!this.form.name) { this.batches = []; return; }
        bridge.call('batches.list', { name: this.form.name }).then(function (data) {
          self.batches = data.items;
        });
      },

      refreshStock: function () {
        var self = this;
        if (!this.form.name || !this.form.batch) {
          this.stock = null;
          return;
        }
        bridge.call('reagent.details', {
          name: this.form.name, batch: this.form.batch,
        }).then(function (res) {
          self.stock = res.found ? res.quantity : null;
        }).catch(function () { self.stock = null; });
      },

      onScan: function () {
        var raw = (this.scanText || '').trim();
        if (!raw) { return; }
        var self = this;

        bridge.call('gs1.parse', { raw: raw }).then(function (res) {
          if (!res.recognised) {
            self.scanState = 'is-miss';
            self.error = '无法识别该条码。';
            return;
          }
          self.scanState = 'is-hit';
          self.error = '';
          if (res.gtin) { self.gtin = res.gtin; }
          if (res.name) { self.form.name = res.name; self.loadBatches(); }
          if (res.batch) { self.form.batch = res.batch; }
          self.scanText = '';
          setTimeout(function () { self.scanState = ''; }, 1200);
        }).catch(function (err) {
          self.scanState = 'is-miss';
          self.error = '解析失败：' + err.message;
        });
      },

      submit: function () {
        var self = this;
        this.error = '';

        if (!this.form.name.trim() || !this.form.batch.trim() || !this.form.quantity) {
          this.error = '请填写试剂名称、批号和数量。';
          return;
        }

        this.busy = true;
        bridge.call('outbound.submit', {
          name: this.form.name.trim(),
          batch: this.form.batch.trim(),
          quantity: this.form.quantity,
          operator: this.form.operator.trim(),
        }).then(function (res) {
          self.busy = false;
          RMS.toast('success',
            '已出库：' + res.name + '（批号 ' + res.batch + '）数量 ' + res.quantity);
          self.form.batch = '';
          self.form.quantity = '';
          self.scanText = '';
          self.gtin = '';
          self.stock = null;
          self.focusScan();
        }).catch(function (err) {
          self.busy = false;
          self.error = err.message;
        });
      },
    },
  };
}(window.RMS));
