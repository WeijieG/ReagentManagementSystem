/* QWebChannel 客户端。
 *
 * 两个必须遵守的时序约束（都是实测踩出来的）：
 *   1. qt.webChannelTransport 由 Qt 注入，时机**晚于**页面内联脚本的执行。
 *      直接 new QWebChannel(qt.webChannelTransport, ...) 会读到 undefined，
 *      报 "Cannot read property 'send' of undefined"。必须轮询等待。
 *   2. qwebchannel.js 必须是随包内置的真实文件。file:// 源加载不了 qrc://，
 *      引用它会静默失败 —— 页面照常渲染但整个界面是死的。
 *
 * file:// 源的额外限制（设计上要绕开，不是 bug）：
 *   - 不能用 ES modules / fetch（CORS 拦截）→ 全部走 QWebChannel
 *   - 不能用 localStorage → 界面风格偏好存到 userData.ini，经 style.set 落盘
 *   - 不能用 Web Worker → 耗时活交给 Python 的 QThread
 */
window.RMS = window.RMS || {};

(function (RMS) {
  'use strict';

  var POLL_INTERVAL_MS = 25;
  var POLL_MAX_ATTEMPTS = 400;   // 约 10 秒

  var api = null;
  var listeners = {};
  var resolveReady, rejectReady;

  var readyPromise = new Promise(function (resolve, reject) {
    resolveReady = resolve;
    rejectReady = reject;
  });

  function onNotify(raw) {
    var evt;
    try {
      evt = JSON.parse(raw);
    } catch (e) {
      console.error('[bridge] 收到无法解析的事件:', raw);
      return;
    }
    var handlers = listeners[evt.event] || [];
    for (var i = 0; i < handlers.length; i++) {
      try {
        handlers[i](evt.data);
      } catch (e) {
        console.error('[bridge] 事件处理器出错 (' + evt.event + '):', e);
      }
    }
    var any = listeners['*'] || [];
    for (var j = 0; j < any.length; j++) {
      try { any[j](evt); } catch (e) { /* 忽略通配处理器的错误 */ }
    }
  }

  function initChannel() {
    var attempts = 0;

    function attempt() {
      if (typeof qt === 'undefined' || !qt.webChannelTransport) {
        if (++attempts > POLL_MAX_ATTEMPTS) {
          rejectReady(new Error(
            '无法建立与后端的通信通道（等待传输对象超时）。\n' +
            '可能原因：页面不是用 load() 从文件加载的，或未调用 setWebChannel()。'
          ));
          return;
        }
        return setTimeout(attempt, POLL_INTERVAL_MS);
      }

      new QWebChannel(qt.webChannelTransport, function (channel) {
        api = channel.objects.api;
        if (!api) {
          rejectReady(new Error('后端未注册名为 api 的对象。'));
          return;
        }
        if (api.notify && api.notify.connect) {
          api.notify.connect(onNotify);
        }
        resolveReady(api);
      });
    }

    attempt();
  }

  /**
   * 调用后端。返回 Promise，成功时给出 data，失败时以 error 拒绝，
   * error 上带 code / detail 两个字段。
   */
  function call(action, payload) {
    return readyPromise.then(function (apiObject) {
      return new Promise(function (resolve, reject) {
        apiObject.invoke(action, JSON.stringify(payload || {}), function (raw) {
          var env;
          try {
            env = JSON.parse(raw);
          } catch (e) {
            reject(new Error('后端返回了无法解析的内容: ' + String(raw).slice(0, 200)));
            return;
          }
          if (env && env.ok) {
            resolve(env.data);
            return;
          }
          var info = (env && env.error) || {};
          var err = new Error(info.message || '未知错误');
          err.code = info.code || 'INTERNAL';
          err.detail = info.detail || '';
          reject(err);
        });
      });
    });
  }

  /** 订阅后端推送。event 为 '*' 时接收全部事件。 */
  function on(event, handler) {
    if (!listeners[event]) { listeners[event] = []; }
    listeners[event].push(handler);
  }

  function ready() {
    return readyPromise;
  }

  RMS.bridge = {
    ready: ready,
    call: call,
    on: on,
    _init: initChannel,
  };

  initChannel();
}(window.RMS));
