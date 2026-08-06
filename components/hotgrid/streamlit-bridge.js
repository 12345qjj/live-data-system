/*
 * Minimal Streamlit component bridge (classic script, no ESM / ArrowTable deps).
 *
 * Streamlit >= 1.45 no longer auto-injects `window.Streamlit` into
 * path-based components, so the official ESM build (which imports
 * ArrowTable -> apache-arrow) can't be loaded without a bundler.
 * This file re-implements exactly the subset the hotgrid component needs,
 * using the stable postMessage protocol Streamlit has shipped since v0.63.
 *
 * APIs implemented:
 *   - Streamlit.RENDER_EVENT         ("streamlit:render")
 *   - Streamlit.events (EventTarget) -> dispatches RENDER_EVENT with {args}
 *   - Streamlit.setComponentReady()  -> tell parent we're listening
 *   - Streamlit.setComponentValue(v) -> send JSON value back to Python
 *   - Streamlit.setFrameHeight(h)    -> resize the iframe
 */
(function () {
  "use strict";

  var RENDER_EVENT = "streamlit:render";
  var registered = false;

  function sendBackMsg(type, data) {
    var msg = { isStreamlitMessage: true, type: type };
    for (var k in data) {
      if (Object.prototype.hasOwnProperty.call(data, k)) msg[k] = data[k];
    }
    window.parent.postMessage(msg, "*");
  }

  function onMessage(event) {
    var data = event && event.data;
    if (!data || data.type !== RENDER_EVENT) return;
    var args = data.args || {};
    var detail = {
      disabled: Boolean(data.disabled),
      args: args,
      theme: data.theme
    };
    var ev = new CustomEvent(RENDER_EVENT, { detail: detail });
    Streamlit.events.dispatchEvent(ev);
  }

  var Streamlit = {
    API_VERSION: 1,
    RENDER_EVENT: RENDER_EVENT,
    events: new EventTarget(),
    lastFrameHeight: undefined,

    setComponentReady: function () {
      if (!registered) {
        window.addEventListener("message", onMessage);
        registered = true;
      }
      sendBackMsg("streamlit:componentReady", { apiVersion: Streamlit.API_VERSION });
    },

    setFrameHeight: function (height) {
      if (height === undefined || height === null) {
        height = document.body.scrollHeight;
      }
      if (height === Streamlit.lastFrameHeight) return;
      Streamlit.lastFrameHeight = height;
      sendBackMsg("streamlit:setFrameHeight", { height: height });
    },

    setComponentValue: function (value) {
      sendBackMsg("streamlit:setComponentValue", { value: value, dataType: "json" });
    }
  };

  window.Streamlit = Streamlit;
  // Let the component know the bridge is ready (defensive; component also
  // checks window.Streamlit synchronously since this script runs first).
  window.dispatchEvent(new Event("streamlit-bridge-ready"));
})();
