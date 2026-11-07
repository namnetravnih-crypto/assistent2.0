/* Ассистент — масштаб интерфейса 80 % (задан в theme.css через zoom).
   Мышь и getBoundingClientRect дают реальные пиксели экрана, а вёрстка считается в «css-пикселях» —
   k() возвращает их отношение, чтобы координаты (печать, кисть, подсказки) попадали точно. */
(function () {
  'use strict';
  let probe = null;
  function k() {
    if (!probe) { probe = document.createElement('div'); probe.style.cssText = 'position:absolute;left:-9999px;top:0;width:100px;height:1px;visibility:hidden;pointer-events:none'; (document.body || document.documentElement).appendChild(probe); }
    const w = probe.getBoundingClientRect().width;
    return w > 0 ? w / 100 : 1;
  }
  window.UIZoom = { k };
})();
