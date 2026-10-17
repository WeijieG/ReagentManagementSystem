/* 入库抽屉。
 *
 * 保留旧实现的关键交互：提交成功后**不关闭**，而是清空表单并把焦点交回扫码框，
 * 便于连续扫码入库（原代码里 `# self.accept()` 是注释掉的，就是这个意图）。
 */
window.RMS = window.RMS || {};
window.RMS.screens = window.RMS.screens || {};

(function (RMS) {
  'use strict';

  var bridge = RMS.bridge;

  function isoDate(d) {
    var m = d.getMonth() + 1;
    var day = d.getDate();
    return d.getFullYear() + '-' + (m < 10 ? '0' + m : m) + '-' + (day < 10 ? '0' + day : day);
  }

  function monthsFromNow(n) {
    var d = new Date();
    d.setMonth(d.getMonth() + n);
    return isoDate(d);
  }

  RMS.components['inbound-drawer'] = {
    template: [
      '<div class="drawer-scrim" @mousedown.self="close">',
      '  <div class="drawer">',
      '    <div class="modal-header">',
      '      <div class="modal-title">入库登记</div>',
      '      <button class="icon-btn modal-close" @click="close">&times;</button>',
      '    </div>',

      '    <div class="modal-body" style="display:flex">',

      /* 左侧：名称列表 */
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

      /* 右侧：表单 */
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
      '          <ui-batch-picker v-model="form.batch" :name="form.name"',
      '                           placeholder="输入、扫码或下拉选择批号" />',
      '        </div>',

      '        <div style="display:flex">',
      '          <div class="field" style="flex:1;margin-right:12px">',
      '            <label class="field-label">生产日期</label>',
      '            <input type="date" class="input" v-model="form.productionDate" />',
      '          </div>',
      '          <div class="field" style="flex:1">',
      '            <label class="field-label">有效期至</label>',
      '            <input type="date" class="input" v-model="form.expiryDate" />',
      '          </div>',
      '        </div>',

      '        <div style="display:flex">',
      '          <div class="field" style="flex:1;margin-right:12px">',
      '            <label class="field-label">数量 <span style="color:var(--danger)">*</span></label>',
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
      '            {{ busy ? "提交中…" : "确定入库" }}</button>',
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
        scanText: '',
        scanState: '',
        gtin: '',
        busy: false,
        error: '',
        form: {
          name: '',
          batch: '',
          productionDate: monthsFromNow(-6),
          expiryDate: monthsFromNow(12),
          quantity: '',
          operator: '',
        },
      };
    },
    computed: {
      filteredNames: function () {
        var q = this.nameQuery.trim().toLowerCase();
        if (!q) { return this.names; }
        return this.names.filter(function (n) { return n.toLowerCase().indexOf(q) >= 0; });
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
          if (res.name) { self.form.name = res.name; }
          if (res.batch) { self.form.batch = res.batch; }
          if (res.productionDate) { self.form.productionDate = res.productionDate; }
          if (res.expiryDate) { self.form.expiryDate = res.expiryDate; }
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
        bridge.call('inbound.submit', {
          name: this.form.name.trim(),
          batch: this.form.batch.trim(),
          quantity: this.form.quantity,
          operator: this.form.operator.trim(),
          gtin: this.gtin,
          qrcode: '',
          productionDate: this.form.productionDate,
          expiryDate: this.form.expiryDate,
        }).then(function (res) {
          self.busy = false;
          RMS.toast('success',
            '已入库：' + res.name + '（批号 ' + res.batch + '）数量 ' + res.quantity);
          self.resetAfterSubmit();
        }).catch(function (err) {
          self.busy = false;
          self.error = err.message;
        });
      },

      /** 提交后保留名称/日期/操作员，只清掉本次扫出的内容，方便连续录入。 */
      resetAfterSubmit: function () {
        this.form.batch = '';
        this.form.quantity = '';
        this.scanText = '';
        this.gtin = '';
        this.error = '';
        this.focusScan();
      },
    },
  };
}(window.RMS));
