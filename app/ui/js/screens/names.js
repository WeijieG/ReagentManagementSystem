/* 试剂名称管理页。
 * 原先是弹窗，现在转为整页 —— 名称列表可能很长，整页空间更合适。
 */
window.RMS = window.RMS || {};
window.RMS.screens = window.RMS.screens || {};

(function (RMS) {
  'use strict';

  var bridge = RMS.bridge;

  var EditModal = {
    props: {
      mode: String,          // 'add' | 'edit'
      row: Object,           // edit 时有值
    },
    components: { 'ui-modal': RMS.components['ui-modal'] },
    template: [
      '<ui-modal :title="mode === \'add\' ? \'新增试剂名称\' : \'编辑试剂名称\'"',
      '           @close="$emit(\'close\')">',
      '  <div v-if="mode === \'edit\'" class="field">',
      '    <label class="field-label">当前名称</label>',
      '    <div>{{ row.name }}</div>',
      '  </div>',

      '  <div class="field">',
      '    <label class="field-label">扫码绑定 GTIN（选填）</label>',
      '    <div class="scan-input" :class="scanState">',
      '      <input class="input" ref="scan" v-model="scanText"',
      '             placeholder="扫描试剂条码自动提取 GTIN" @keyup.enter="onScan" />',
      '    </div>',
      '    <div v-if="gtin" class="field-hint">已识别 GTIN：<span class="mono">{{ gtin }}</span></div>',
      '    <div v-else class="field-hint">',
      '      绑定后，扫码即可自动带出试剂名称。',
      '      <span v-if="mode === \'edit\'">留空则保持原有绑定。</span>',
      '    </div>',
      '  </div>',

      '  <div class="field">',
      '    <label class="field-label">试剂名称 <span style="color:var(--danger)">*</span></label>',
      '    <input class="input" v-model="name" :placeholder="namePlaceholder" @keyup.enter="submit" />',
      '    <div v-if="error" class="field-error">{{ error }}</div>',
      '  </div>',

      '  <template #footer>',
      '    <button class="btn" @click="$emit(\'close\')">取消</button>',
      '    <button class="btn btn-primary" :disabled="busy" @click="submit">',
      '      {{ busy ? "保存中…" : "保存" }}</button>',
      '  </template>',
      '</ui-modal>',
    ].join(''),
    data: function () {
      var isEdit = this.mode === 'edit';
      return {
        name: isEdit ? this.row.name : '',
        gtin: isEdit ? (this.row.gtin || '') : '',
        originalGtin: isEdit ? (this.row.gtin || '') : '',
        scanText: '',
        scanState: '',
        error: '',
        busy: false,
      };
    },
    computed: {
      namePlaceholder: function () {
        return this.mode === 'edit' ? '留空则保持原名不变' : '例如：生理盐水';
      },
    },
    methods: {
      onScan: function () {
        var raw = (this.scanText || '').trim();
        if (!raw) { return; }
        var self = this;
        bridge.call('gs1.parse', { raw: raw }).then(function (res) {
          if (res.gtin) {
            self.gtin = res.gtin;
            self.scanState = 'is-hit';
            // GTIN 未绑定过名称时，顺手把数据库里的名称填进来，省一次手输
            if (res.name && self.mode === 'add' && !self.name) { self.name = res.name; }
          } else {
            self.scanState = 'is-miss';
            self.error = '条码里没有 GTIN 字段。';
          }
          self.scanText = '';
          setTimeout(function () { self.scanState = ''; }, 1200);
        }).catch(function (err) {
          self.error = '解析失败：' + err.message;
        });
      },

      submit: function () {
        var self = this;
        this.error = '';

        if (this.mode === 'add' && !this.name.trim()) {
          this.error = '试剂名称不能为空。';
          return;
        }

        this.busy = true;
        var promise;
        if (this.mode === 'add') {
          promise = bridge.call('names.add', { name: this.name.trim(), gtin: this.gtin });
        } else {
          promise = bridge.call('names.update', {
            oldName: this.row.name,
            oldGtin: this.originalGtin,
            newName: this.name.trim(),
            newGtin: this.gtin,
          });
        }

        promise.then(function () {
          self.busy = false;
          RMS.toast('success', self.mode === 'add' ? '已新增试剂名称' : '已保存修改');
          self.$emit('close', true);
        }).catch(function (err) {
          self.busy = false;
          self.error = err.message;
        });
      },
    },
  };

  var ConfirmModal = {
    props: { text: String },
    components: { 'ui-modal': RMS.components['ui-modal'] },
    template: [
      '<ui-modal title="确认删除">',
      '  <div>{{ text }}</div>',
      '  <template #footer>',
      '    <button class="btn" @click="$emit(\'close\')">取消</button>',
      '    <button class="btn btn-danger" @click="$emit(\'close\', true)">删除</button>',
      '  </template>',
      '</ui-modal>',
    ].join(''),
  };

  RMS.screens['screen-names'] = {
    components: {
      'ui-empty': RMS.components['ui-empty'],
      'name-edit': EditModal,
      'name-confirm': ConfirmModal,
    },
    template: [
      '<div class="page">',
      '  <div class="page-toolbar">',
      '    <input class="input" style="width:260px" v-model="filter"',
      '           placeholder="搜索试剂名称…" />',
      '    <span class="spacer"></span>',
      '    <button class="btn btn-primary" @click="openAdd">新增试剂名称</button>',
      '  </div>',

      '  <div class="card">',
      '    <div v-if="loading" class="loading-overlay">',
      '      <span class="spinner"></span> 正在加载…',
      '    </div>',
      '    <ui-empty v-else-if="filtered.length === 0"',
      '              :title="filter ? \'没有匹配的名称\' : \'还没有试剂名称\'"',
      '              :hint="filter ? \'试试换个关键词\' : \'点击右上角「新增试剂名称」开始\'"></ui-empty>',
      '    <div v-else class="table-wrap">',
      '      <table class="table">',
      '        <thead><tr>',
      '          <th>试剂名称</th><th>GTIN 绑定</th><th style="width:180px">操作</th>',
      '        </tr></thead>',
      '        <tbody>',
      '          <tr v-for="r in filtered" :key="r.name">',
      '            <td>{{ r.name }}</td>',
      '            <td>',
      '              <span v-if="r.gtin" class="mono">{{ r.gtin }}</span>',
      '              <span v-else class="badge badge-neutral">未绑定</span>',
      '            </td>',
      '            <td>',
      '              <button class="btn btn-sm" @click="openEdit(r)">编辑</button>',
      '              <button class="btn btn-sm btn-ghost" style="margin-left:6px;color:var(--danger)"',
      '                      @click="askDelete(r)">删除</button>',
      '            </td>',
      '          </tr>',
      '        </tbody>',
      '      </table>',
      '    </div>',
      '  </div>',

      '  <name-edit v-if="editing" :mode="editing.mode" :row="editing.row"',
      '             @close="onEditClosed"></name-edit>',
      '  <name-confirm v-if="pendingDelete" :text="deleteText"',
      '                @close="onDeleteClosed"></name-confirm>',
      '</div>',
    ].join(''),
    data: function () {
      return {
        rows: [],
        loading: true,
        filter: '',
        editing: null,
        pendingDelete: null,
      };
    },
    computed: {
      filtered: function () {
        var q = this.filter.trim().toLowerCase();
        if (!q) { return this.rows; }
        return this.rows.filter(function (r) {
          return r.name.toLowerCase().indexOf(q) >= 0
              || (r.gtin || '').indexOf(q) >= 0;
        });
      },
      deleteText: function () {
        return this.pendingDelete
          ? '确定要删除试剂名称「' + this.pendingDelete.name + '」吗？若该名称已有库存或出入库记录，将无法删除。'
          : '';
      },
    },
    mounted: function () {
      this.load();
    },
    methods: {
      load: function () {
        var self = this;
        this.loading = true;
        return bridge.call('names.list').then(function (data) {
          self.rows = data.rows;
          self.loading = false;
        }).catch(function (err) {
          self.loading = false;
          RMS.toast('error', '加载试剂名称失败：' + err.message);
        });
      },
      openAdd: function () { this.editing = { mode: 'add', row: null }; },
      openEdit: function (row) { this.editing = { mode: 'edit', row: row }; },
      onEditClosed: function (saved) {
        this.editing = null;
        if (saved) { this.load(); }
      },
      askDelete: function (row) { this.pendingDelete = row; },
      onDeleteClosed: function (confirmed) {
        var row = this.pendingDelete;
        this.pendingDelete = null;
        if (!confirmed || !row) { return; }

        var self = this;
        bridge.call('names.delete', { name: row.name }).then(function () {
          RMS.toast('success', '已删除「' + row.name + '」');
          self.load();
        }).catch(function (err) {
          RMS.toast('error', err.message);
        });
      },
    },
  };
}(window.RMS));
