/* 前端 polyfill —— 必须在 vue.global.prod.js 之前加载。
 *
 * 内嵌内核是 Chromium 83（V8 8.3）。本文件补齐 Vue 3.4+ 会用到、
 * 而 Chromium 83 没有的几个方法，作为版本误升级时的兜底。
 *
 * 注意：本文件自身不能用逻辑赋值运算符（||= 是 Chrome 85 才有的语法），
 * 在 Chromium 83 上是解析错误，会导致整个文件不执行。
 */
(function () {
  'use strict';

  // String.prototype.replaceAll —— Chrome 85
  if (!String.prototype.replaceAll) {
    String.prototype.replaceAll = function (search, replacement) {
      var subject = String(this);
      if (search instanceof RegExp) {
        // 与原语义一致：正则必须是全局的
        if (!search.global) {
          throw new TypeError('replaceAll must be called with a global RegExp');
        }
        return subject.replace(search, replacement);
      }
      var needle = String(search);
      if (needle === '') {
        return subject.split('').join(String(replacement));
      }
      return subject.split(needle).join(String(replacement));
    };
  }

  // Array.prototype.at / String.prototype.at —— Chrome 92
  if (!Array.prototype.at) {
    Object.defineProperty(Array.prototype, 'at', {
      value: function (index) {
        var i = Math.trunc(index) || 0;
        if (i < 0) { i += this.length; }
        return (i < 0 || i >= this.length) ? undefined : this[i];
      },
      writable: true,
      configurable: true,
    });
  }
  if (!String.prototype.at) {
    Object.defineProperty(String.prototype, 'at', {
      value: function (index) {
        return Array.prototype.at.call(String(this), index);
      },
      writable: true,
      configurable: true,
    });
  }

  // Array.prototype.findLast / findLastIndex —— Chrome 97
  if (!Array.prototype.findLast) {
    Object.defineProperty(Array.prototype, 'findLast', {
      value: function (predicate, thisArg) {
        for (var i = this.length - 1; i >= 0; i--) {
          if (predicate.call(thisArg, this[i], i, this)) { return this[i]; }
        }
        return undefined;
      },
      writable: true,
      configurable: true,
    });
  }

  // Object.hasOwn —— Chrome 93
  if (!Object.hasOwn) {
    Object.hasOwn = function (obj, key) {
      return Object.prototype.hasOwnProperty.call(obj, key);
    };
  }

  // Array.prototype.flat / flatMap —— Chrome 69，理论上够，但补齐无副作用
  if (!Array.prototype.flat) {
    Object.defineProperty(Array.prototype, 'flat', {
      value: function (depth) {
        var d = depth === undefined ? 1 : depth;
        return this.reduce(function (acc, val) {
          if (Array.isArray(val) && d > 0) {
            return acc.concat(val.flat(d - 1));
          }
          return acc.concat(val);
        }, []);
      },
      writable: true,
      configurable: true,
    });
  }

  // Promise.allSettled —— Chrome 76
  if (!Promise.allSettled) {
    Promise.allSettled = function (iterable) {
      return Promise.all(Array.prototype.map.call(iterable, function (item) {
        return Promise.resolve(item).then(
          function (value) { return { status: 'fulfilled', value: value }; },
          function (reason) { return { status: 'rejected', reason: reason }; }
        );
      }));
    };
  }

  // AggregateError —— Chrome 85
  if (typeof window.AggregateError === 'undefined') {
    window.AggregateError = function AggregateError(errors, message) {
      var e = new Error(message);
      e.name = 'AggregateError';
      e.errors = errors;
      return e;
    };
  }
}());
