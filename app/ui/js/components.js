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
     分页条
     ============================================================ */
  components['ui-pager'] = {
    props: {
      page: { type: Number, default: 1 },
      pages: { type: Number, default: 1 },
      total: { type: Number, default: 0 },
      pageSize: { type: Number, default: 50 },
      pageSizes: { type: Array, default: function () { return [20, 50, 100, 200]; } },
      loading: Boolean,
    },
    template: [
      '<div class="pager">',
      '  <span>每页</span>',
      '  <select class="input select pager-size" :value="pageSize" @change="onSize">',
      '    <option v-for="s in pageSizes" :key="s" :value="s">{{ s }}</option>',
      '  </select>',
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
    methods: {
      go: function (p) {
        var target = Math.min(Math.max(1, p), this.pages);
        if (target !== this.page) { this.$emit('change', target); }
      },
      onSize: function (e) {
        this.$emit('size-change', parseInt(e.target.value, 10));
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
