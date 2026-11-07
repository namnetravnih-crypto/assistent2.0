(() => {
'use strict';

const $ = id => document.getElementById(id);
const MAX_SIDE = 1800;                                // максимальная сторона рабочего изображения, px
const FULL_SIDE = 6000;                               // максимальная сторона исходной страницы / фото до обрезки, px
const DPI = 600;                                      // разрешение выходного PNG
const TARGET_MM = 40;                                 // диаметр печати в готовом файле
const TARGET_PX = Math.round(TARGET_MM / 25.4 * DPI); // 945 px

const el = {
  file: $('fileInput'), label: $('uploadLabel'), drop: $('dropZone'),
  workspace: document.querySelector('.workspace'), panes: $('panes'),
  container: $('canvasContainer'), canvas: $('mainCanvas'), orig: $('origCanvas'),
  cursor: $('brushCursor'), cursor2: $('brushCursor2'),
  status: $('statusBar').firstElementChild, fname: $('fileName'), autoInfo: $('autoInfo'),
  btnAuto: $('btnAuto'), manual: $('manualBox'),
  pageBar: $('pageBar'), pageBtns: $('pageBtns'),
  autoCrop: $('autoCropCheck'), clip: $('clipCheck'), autoRound: $('autoRoundCheck'),
  removeText: $('removeTextCheck'), pText: $('paramText'),
  pBright: $('paramBrightness'), pChroma: $('paramChroma'),
  pShadow: $('paramShadows'), pSmooth: $('paramSmoothing'),
  presetChips: $('presetChips'), btnPick: $('btnPick'), pPick: $('paramPick'), pickChips: $('pickChips'),
  btnMask: $('btnMagicCursor'), btnEraser: $('btnEraser'), btnRestore: $('btnRestore'), btnTextMask: $('btnTextMask'),
  pEdge: $('paramEdgeSens'),
  brush: $('brushSize'), brushVal: $('valBrushSize'),
  undo: $('btnUndo'), round: $('btnMakeRound'),
  dl: $('btnDownload'), reset: $('btnReset'), bgSwitch: $('bgSwitch')
};
const ctx = el.canvas.getContext('2d', { willReadFrequently: true });
const octx = el.orig.getContext('2d');

/* ---------- состояние ---------- */
let fullCanvas = null;             // загруженное изображение / страница целиком (до автообрезки)
let origCanvas = null;             // рабочее изображение (после автообрезки, не больше MAX_SIDE)
let src = null, srcPix = null, W = 0, H = 0;
let grid = null;                   // карта бумаги (для подавления бликов и теней)
let flat = null, flatKey = -1;     // изображение после выравнивания освещения (RGB)
let ana = null;                    // результат анализа: контур печати, цвета, подпись
let base = null;                   // результат автоматической обработки
let cur = null;                    // альфа после ручных правок
let imgOut = null;
let edits = [], history = [];
let tool = null;
const sizes = { mask: 60, eraser: 15, restore: 40, textmask: 40 };
let dispScale = 1;
let stroke = null, lastPt = null;
let fileLabel = 'stamp';
let picks = [];                    // выбранные пипеткой цвета
const presetsOn = new Set();       // включённые быстрые оттенки
let bboxDirty = true, bbox = null;
let loadNotes = [], procNotes = [];
let pdfDoc = null, pdfPage = 1, pdfCount = 1, pdfAuto = 0;
let zoneRegion = null, lastRegion = null;   // рабочая зона: выбранная вручную и фактически использованная (в пикселях исходного листа)

const clamp = (v, a, b) => v < a ? a : (v > b ? b : v);
const smooth = t => t <= 0 ? 0 : (t >= 1 ? 1 : t * t * (3 - 2 * t));
const setStatus = t => { el.status.textContent = t; };
const thrOf = th => 5 + th * 1.1;  // ползунок 0..100 → порог 5..115
/* «Уточнение края»: чувствительность 0..100 → множитель порога плотности. 50 — прежнее поведение (множитель 1);
   ниже — остаётся только самый плотный оттиск, выше — сохраняются и более тонкие, бледные линии. */
const edgeFactor = s => s <= 50 ? 1 + (50 - s) / 50 * 0.4 : 1 - (s - 50) / 50 * 0.55;

/* ---------- параметры ---------- */
function P() {
  return {
    rm: el.removeText.checked, text: +el.pText.value,
    th: +el.pBright.value, sens: +el.pChroma.value,
    shadow: +el.pShadow.value, smooth: +el.pSmooth.value,
    tol: +el.pPick.value, edge: +el.pEdge.value
  };
}
/* =====================================================================
   Общие инструменты: карта бумаги, морфология, компоненты, подгонка эллипса
   ===================================================================== */

/* Карта бумаги: по блокам берём цвет яркой части (80-й перцентиль) — это бумага с бликами и тенями.
   Блоки, целиком залитые краской, заменяются значением соседей, чтобы сплошной оттиск не «растворялся». */
function makeGrid(pix, w, h) {
  const B = Math.max(24, Math.round(Math.min(w, h) / 12));
  const gw = Math.ceil(w / B), gh = Math.ceil(h / B), M = gw * gh;
  let g = new Float32Array(M * 3);
  const gl = new Float32Array(M);
  const hist = new Uint32Array(256);
  const st = w * h > 160000 ? 2 : 1;
  for (let gy = 0; gy < gh; gy++) {
    for (let gx = 0; gx < gw; gx++) {
      const x0 = gx * B, y0 = gy * B, x1 = Math.min(w, x0 + B), y1 = Math.min(h, y0 + B);
      hist.fill(0); let cnt = 0;
      for (let y = y0; y < y1; y += st) for (let x = x0; x < x1; x += st) {
        const i = (y * w + x) * 4;
        hist[(pix[i] * 77 + pix[i + 1] * 150 + pix[i + 2] * 29) >> 8]++; cnt++;
      }
      let acc = 0, p = 255; const target = cnt * 0.2;
      for (let l = 255; l >= 0; l--) { acc += hist[l]; if (acc >= target) { p = l; break; } }
      let sr = 0, sg = 0, sb = 0, n = 0;
      for (let y = y0; y < y1; y += st) for (let x = x0; x < x1; x += st) {
        const i = (y * w + x) * 4;
        if (((pix[i] * 77 + pix[i + 1] * 150 + pix[i + 2] * 29) >> 8) >= p) {
          sr += pix[i]; sg += pix[i + 1]; sb += pix[i + 2]; n++;
        }
      }
      const idx = gy * gw + gx, o = idx * 3;
      if (n) { g[o] = sr / n; g[o + 1] = sg / n; g[o + 2] = sb / n; } else { g[o] = g[o + 1] = g[o + 2] = 255; }
      gl[idx] = 0.299 * g[o] + 0.587 * g[o + 1] + 0.114 * g[o + 2];
    }
  }
  const sorted = Array.from(gl).sort((a, b) => a - b);
  const med = sorted[M >> 1];
  const good = new Uint8Array(M);
  for (let i = 0; i < M; i++) good[i] = gl[i] >= med * 0.6 ? 1 : 0;
  for (let pass = 0; pass < 30; pass++) {
    const upd = [];
    for (let gy = 0; gy < gh; gy++) for (let gx = 0; gx < gw; gx++) {
      const idx = gy * gw + gx;
      if (good[idx]) continue;
      let r = 0, gg = 0, b = 0, c = 0;
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const yy = gy + dy, xx = gx + dx;
        if (yy < 0 || yy >= gh || xx < 0 || xx >= gw) continue;
        const j = yy * gw + xx;
        if (!good[j]) continue;
        r += g[j * 3]; gg += g[j * 3 + 1]; b += g[j * 3 + 2]; c++;
      }
      if (c) upd.push([idx, r / c, gg / c, b / c]);
    }
    if (!upd.length) break;
    for (const u of upd) { g[u[0] * 3] = u[1]; g[u[0] * 3 + 1] = u[2]; g[u[0] * 3 + 2] = u[3]; good[u[0]] = 1; }
  }
  for (let pass = 0; pass < 2; pass++) {
    const t = new Float32Array(g.length);
    for (let gy = 0; gy < gh; gy++) for (let gx = 0; gx < gw; gx++) {
      for (let k = 0; k < 3; k++) {
        let s = 0, c = 0;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const yy = gy + dy, xx = gx + dx;
          if (yy < 0 || yy >= gh || xx < 0 || xx >= gw) continue;
          s += g[(yy * gw + xx) * 3 + k]; c++;
        }
        t[(gy * gw + gx) * 3 + k] = s / c;
      }
    }
    g = t;
  }
  return { B, gw, gh, g };
}

/* Выравнивание освещения: каждый пиксель делится на цвет бумаги вокруг него (m — сила 0..1) */
function makeFlat(pix, w, h, grd, m) {
  const out = new Uint8ClampedArray(w * h * 3);
  const { g, gw, gh, B } = grd;
  const xa = new Int32Array(w), xb = new Int32Array(w), xf = new Float32Array(w);
  for (let x = 0; x < w; x++) {
    const f = clamp((x + 0.5) / B - 0.5, 0, gw - 1);
    const i0 = Math.min(Math.floor(f), gw - 1);
    xa[x] = i0; xb[x] = Math.min(i0 + 1, gw - 1); xf[x] = f - i0;
  }
  for (let y = 0; y < h; y++) {
    const f = clamp((y + 0.5) / B - 0.5, 0, gh - 1);
    const y0 = Math.min(Math.floor(f), gh - 1), y1 = Math.min(y0 + 1, gh - 1), yf = f - y0;
    const rA = y0 * gw, rB = y1 * gw;
    for (let x = 0; x < w; x++) {
      const idx = y * w + x, i = idx * 4, o = idx * 3;
      let r = pix[i], gg = pix[i + 1], b = pix[i + 2];
      if (m > 0) {
        const p00 = (rA + xa[x]) * 3, p01 = (rA + xb[x]) * 3, p10 = (rB + xa[x]) * 3, p11 = (rB + xb[x]) * 3;
        const fx = xf[x];
        const bgR = Math.max(30, (g[p00] * (1 - fx) + g[p01] * fx) * (1 - yf) + (g[p10] * (1 - fx) + g[p11] * fx) * yf);
        const bgG = Math.max(30, (g[p00 + 1] * (1 - fx) + g[p01 + 1] * fx) * (1 - yf) + (g[p10 + 1] * (1 - fx) + g[p11 + 1] * fx) * yf);
        const bgB = Math.max(30, (g[p00 + 2] * (1 - fx) + g[p01 + 2] * fx) * (1 - yf) + (g[p10 + 2] * (1 - fx) + g[p11 + 2] * fx) * yf);
        r += (Math.min(255, r * 255 / bgR) - r) * m;
        gg += (Math.min(255, gg * 255 / bgG) - gg) * m;
        b += (Math.min(255, b * 255 / bgB) - b) * m;
      }
      out[o] = r; out[o + 1] = gg; out[o + 2] = b;
    }
  }
  return out;
}

/* Порог Отцу: граница между бумагой и краской */
function otsu(hist, total) {
  let sumAll = 0; for (let i = 0; i < 256; i++) sumAll += i * hist[i];
  let wB = 0, sumB = 0, best = -1, t0 = 40;
  for (let t = 0; t < 256; t++) {
    wB += hist[t]; if (!wB) continue;
    const wF = total - wB; if (!wF) break;
    sumB += t * hist[t];
    const mB = sumB / wB, mF = (sumAll - sumB) / wF;
    const v = wB * wF * (mB - mF) * (mB - mF);
    if (v > best) { best = v; t0 = t; }
  }
  return t0;
}

/* Бинарная морфология квадратным окном радиуса r (dil=true — расширение, false — сжатие) */
function morph(mask, w, h, r, dil) {
  if (r <= 0) return Uint8Array.from(mask);
  const win = 2 * r + 1, t = new Uint8Array(w * h), o = new Uint8Array(w * h);
  for (let y = 0; y < h; y++) {
    const row = y * w; let sum = 0;
    for (let x = -r; x <= r; x++) sum += mask[row + clamp(x, 0, w - 1)];
    for (let x = 0; x < w; x++) {
      t[row + x] = dil ? (sum > 0 ? 1 : 0) : (sum === win ? 1 : 0);
      sum += mask[row + Math.min(x + r + 1, w - 1)] - mask[row + Math.max(x - r, 0)];
    }
  }
  for (let x = 0; x < w; x++) {
    let sum = 0;
    for (let y = -r; y <= r; y++) sum += t[clamp(y, 0, h - 1) * w + x];
    for (let y = 0; y < h; y++) {
      o[y * w + x] = dil ? (sum > 0 ? 1 : 0) : (sum === win ? 1 : 0);
      sum += t[Math.min(y + r + 1, h - 1) * w + x] - t[Math.max(y - r, 0) * w + x];
    }
  }
  return o;
}

function boxBlurF(a, w, h, r) {
  const win = 2 * r + 1, t = new Float32Array(a.length), o = new Float32Array(a.length);
  for (let y = 0; y < h; y++) {
    const row = y * w; let sum = 0;
    for (let x = -r; x <= r; x++) sum += a[row + clamp(x, 0, w - 1)];
    for (let x = 0; x < w; x++) {
      t[row + x] = sum / win;
      sum += a[row + Math.min(x + r + 1, w - 1)] - a[row + Math.max(x - r, 0)];
    }
  }
  for (let x = 0; x < w; x++) {
    let sum = 0;
    for (let y = -r; y <= r; y++) sum += t[clamp(y, 0, h - 1) * w + x];
    for (let y = 0; y < h; y++) {
      o[y * w + x] = sum / win;
      sum += t[Math.min(y + r + 1, h - 1) * w + x] - t[Math.max(y - r, 0) * w + x];
    }
  }
  return o;
}

/* Связные компоненты (8 соседей). comps[0] — пустая заглушка. */
function components(mask, w, h) {
  const N = w * h, labels = new Int32Array(N), q = new Int32Array(N), comps = [null];
  let id = 0;
  for (let s = 0; s < N; s++) {
    if (!mask[s] || labels[s]) continue;
    id++; let head = 0, tail = 0; q[tail++] = s; labels[s] = id;
    let x0 = w, y0 = h, x1 = -1, y1 = -1, sx = 0, sy = 0;
    while (head < tail) {
      const i = q[head++], x = i % w, y = (i - x) / w;
      sx += x; sy += y;
      if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
      for (let dy = -1; dy <= 1; dy++) {
        const ny = y + dy; if (ny < 0 || ny >= h) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const nx = x + dx; if (nx < 0 || nx >= w) continue;
          const j = ny * w + nx;
          if (mask[j] && !labels[j]) { labels[j] = id; q[tail++] = j; }
        }
      }
    }
    comps.push({ id, n: tail, x0, y0, x1, y1, cx: sx / tail, cy: sy / tail });
  }
  return { labels, comps };
}

/* Подгонка эллипса (общее коническое сечение) методом наименьших квадратов с отбрасыванием выбросов */
function conicFit(pts) {
  const n = pts.length;
  let mx = 0, my = 0;
  for (const p of pts) { mx += p[0]; my += p[1]; }
  mx /= n; my /= n;
  let sr = 0; for (const p of pts) sr += Math.hypot(p[0] - mx, p[1] - my);
  const s = sr / n || 1;

  const solve = P_ => {
    const A = []; for (let i = 0; i < 5; i++) A.push(new Float64Array(6));
    for (const p of P_) {
      const x = (p[0] - mx) / s, y = (p[1] - my) / s, f = [x * x, x * y, y * y, x, y];
      for (let i = 0; i < 5; i++) { for (let j = 0; j < 5; j++) A[i][j] += f[i] * f[j]; A[i][5] += f[i]; }
    }
    for (let i = 0; i < 5; i++) {
      let piv = i;
      for (let r = i + 1; r < 5; r++) if (Math.abs(A[r][i]) > Math.abs(A[piv][i])) piv = r;
      if (Math.abs(A[piv][i]) < 1e-12) return null;
      const tmp = A[i]; A[i] = A[piv]; A[piv] = tmp;
      for (let r = 0; r < 5; r++) {
        if (r === i) continue;
        const k = A[r][i] / A[i][i];
        for (let c = i; c < 6; c++) A[r][c] -= k * A[i][c];
      }
    }
    const a = A[0][5] / A[0][0], b = A[1][5] / A[1][1], c = A[2][5] / A[2][2], d = A[3][5] / A[3][3], e = A[4][5] / A[4][4];
    const det = a * c - b * b / 4;
    if (det <= 1e-9 || a + c <= 0) return null;
    const cx = (-(d / 2) * c + (e / 2) * (b / 2)) / det;
    const cy = ((d / 2) * (b / 2) - (e / 2) * a) / det;
    const k = 1 + a * cx * cx + b * cx * cy + c * cy * cy;
    if (k <= 1e-9) return null;
    const m11 = a / k, m12 = b / (2 * k), m22 = c / k;
    const tr = m11 + m22, dt = m11 * m22 - m12 * m12, dc = Math.sqrt(Math.max(0, tr * tr / 4 - dt));
    const l1 = tr / 2 + dc, l2 = tr / 2 - dc;
    if (l2 <= 1e-9) return null;
    return {
      cx, cy, m11, m12, m22,
      aN: 1 / Math.sqrt(l2), bN: 1 / Math.sqrt(l1),
      th: 0.5 * Math.atan2(2 * m12, m11 - m22) + Math.PI / 2
    };
  };
  const resid = (f, p) => {
    const dx = (p[0] - mx) / s - f.cx, dy = (p[1] - my) / s - f.cy;
    return Math.sqrt(Math.max(0, f.m11 * dx * dx + 2 * f.m12 * dx * dy + f.m22 * dy * dy)) - 1;
  };

  let f = solve(pts);
  if (!f) return null;
  let r2 = 0; for (const p of pts) { const r = resid(f, p); r2 += r * r; }
  const rms = Math.sqrt(r2 / n);
  const lim = Math.max(0.04, 2.2 * rms);
  const kept = pts.filter(p => Math.abs(resid(f, p)) <= lim);
  if (kept.length >= n * 0.6 && kept.length < n) { const f2 = solve(kept); if (f2) f = f2; }
  // качество: среднеквадратичное отклонение лучших 80% точек (хвосты подписи и росчерки на контур не влияют)
  const rs = pts.map(p => { const r = Math.min(Math.abs(resid(f, p)), 0.5); return r * r; }).sort((x, y) => x - y);
  const m80 = Math.max(10, Math.floor(n * 0.8));
  let q2 = 0; for (let i = 0; i < m80; i++) q2 += rs[i];
  return {
    cx: mx + f.cx * s, cy: my + f.cy * s, a: f.aN * s, b: f.bN * s, th: f.th,
    q: Math.sqrt(q2 / m80), n
  };
}

/* Эллипс по внешней границе компоненты: 360 лучей из центра, берём самую дальнюю точку компоненты на каждом луче */
function fitEllipse(labels, L, c, w, h) {
  const pts = [], maxR = Math.hypot(c.x1 - c.x0, c.y1 - c.y0) + 2;
  for (let k = 0; k < 360; k++) {
    const an = k * Math.PI / 180, dx = Math.cos(an), dy = Math.sin(an);
    let last = -1;
    for (let r = 0; r <= maxR; r += 1) {
      const x = Math.round(c.cx + dx * r), y = Math.round(c.cy + dy * r);
      if (x < 0 || y < 0 || x >= w || y >= h) break;
      if (labels[y * w + x] === L) last = r;
    }
    if (last > 0) pts.push([c.cx + dx * last, c.cy + dy * last]);
  }
  if (pts.length < 40) return null;
  return conicFit(pts);
}

function ellExtents(f) {
  const c = Math.cos(f.th), s = Math.sin(f.th);
  return {
    hw: Math.sqrt(f.a * f.a * c * c + f.b * f.b * s * s),
    hh: Math.sqrt(f.a * f.a * s * s + f.b * f.b * c * c)
  };
}

/* =====================================================================
   Быстрый поиск печати на странице / фото
   ===================================================================== */
function quickStampBox(canvas) {
  const cw = canvas.width, chh = canvas.height;
  const k = Math.min(1, 900 / Math.max(cw, chh));
  const w = Math.max(16, Math.round(cw * k)), h = Math.max(16, Math.round(chh * k));
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const cx = c.getContext('2d', { willReadFrequently: true });
  cx.fillStyle = '#fff'; cx.fillRect(0, 0, w, h);
  cx.imageSmoothingQuality = 'high';
  cx.drawImage(canvas, 0, 0, w, h);
  const pix = cx.getImageData(0, 0, w, h).data;
  const fl = makeFlat(pix, w, h, makeGrid(pix, w, h), 0.85);

  const N = w * h, sA = new Uint8ClampedArray(N), chA = new Uint8ClampedArray(N), dkA = new Uint8ClampedArray(N);
  const hist = new Uint32Array(256);
  for (let i = 0; i < N; i++) {
    const r = fl[i * 3], g = fl[i * 3 + 1], b = fl[i * 3 + 2];
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b), ch = mx - mn;
    const dark = 255 - (0.299 * r + 0.587 * g + 0.114 * b);
    chA[i] = ch; dkA[i] = dark;
    const s = Math.min(255, Math.max(dark, ch * 1.7));
    sA[i] = s; hist[s | 0]++;
  }
  const T = clamp(otsu(hist, N) * 0.55, 14, 85);
  const chrom = new Uint8Array(N), neut = new Uint8Array(N);
  let nc = 0, nn = 0;
  for (let i = 0; i < N; i++) {
    if (sA[i] < T) continue;
    neut[i] = 1; nn++;
    if (chA[i] >= 20 && chA[i] >= 0.3 * (dkA[i] + 1)) { chrom[i] = 1; nc++; }
  }
  const rr = Math.max(1, Math.round(Math.max(w, h) / 450));
  const ringIn = mask => {
    const d = morph(mask, w, h, rr, true);
    const { labels, comps } = components(d, w, h);
    let best = null;
    const order = comps.slice(1).sort((a, b) => b.n - a.n).slice(0, 8);
    for (const cp of order) {
      const bw = cp.x1 - cp.x0 + 1, bh = cp.y1 - cp.y0 + 1;
      if (Math.min(bw, bh) < 0.04 * Math.max(w, h) || cp.n < 60) continue;
      const f = fitEllipse(labels, cp.id, cp, w, h);
      if (!f || f.q >= 0.07 || f.a / f.b > 2.5 || f.b < 0.015 * Math.max(w, h)) continue;
      if (cp.n / (Math.PI * f.a * f.b) > 0.9) continue;       // сплошное пятно — не печать
      const score = Math.sqrt(f.a * f.b) / Math.max(w, h) * (1 - f.q / 0.07);
      if (!best || score > best.score) best = { score, fit: f };
    }
    return best;
  };
  let res = nc >= 80 ? ringIn(chrom) : null;
  if (!res && nn >= 80) res = ringIn(neut);
  if (!res) return null;

  const f = res.fit, e = ellExtents(f), m = 1.6;
  const x0 = clamp((f.cx - e.hw * m) / k, 0, cw), x1 = clamp((f.cx + e.hw * m) / k, 0, cw);
  const y0 = clamp((f.cy - e.hh * m) / k, 0, chh), y1 = clamp((f.cy + e.hh * m) / k, 0, chh);
  return {
    box: { x0: Math.floor(x0), y0: Math.floor(y0), x1: Math.ceil(x1), y1: Math.ceil(y1) },
    score: res.score,
    areaRatio: ((x1 - x0) * (y1 - y0)) / (cw * chh)
  };
}

/* =====================================================================
   Загрузка файлов: изображения и PDF
   ===================================================================== */
let pdfjsPromise = null;
function loadPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (!pdfjsPromise) {
    const base = window.PDFJS_BASE || 'assets/vendor/pdfjs/';
    pdfjsPromise = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = base + 'pdf.min.js';
      s.onload = () => {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = base + 'pdf.worker.min.js';
        res(window.pdfjsLib);
      };
      s.onerror = () => { pdfjsPromise = null; rej(new Error('Не удалось загрузить модуль чтения PDF (файл assets/vendor/pdfjs/pdf.min.js не найден).')); };
      document.head.appendChild(s);
    });
  }
  return pdfjsPromise;
}

async function renderPdfPage(n, maxSide) {
  const page = await pdfDoc.getPage(n);
  const v1 = page.getViewport({ scale: 1 });
  const scale = maxSide / Math.max(v1.width, v1.height);
  const vp = page.getViewport({ scale });
  const c = document.createElement('canvas');
  c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  const cx = c.getContext('2d');
  cx.fillStyle = '#fff'; cx.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: cx, viewport: vp }).promise;
  return c;
}

async function showPdfPage(n) {
  if (!pdfDoc) return;
  n = clamp(n, 1, pdfCount);
  zoneRegion = null;
  setStatus('Чтение страницы ' + n + ' из ' + pdfCount + '…');
  const c = await renderPdfPage(n, FULL_SIDE);
  pdfPage = n;
  updatePdfBar();
  el.fname.textContent = (el.fname.dataset.base || '') + ' (стр. ' + n + ' из ' + pdfCount + ')';
  startFromFull(c, pdfCount > 1 ? ['PDF: показана страница ' + n + ' из ' + pdfCount + (n === pdfAuto ? ' — на ней программа нашла печать. Другие страницы — кнопками внизу.' : '.') ] : []);
}
function updatePdfBar() {
  const multi = !!pdfDoc && pdfCount > 1;
  el.pageBar.hidden = !multi;
  el.workspace.classList.toggle('has-pages', multi);
  el.pageBtns.innerHTML = '';
  if (!multi) { layout(); return; }
  for (let n = 1; n <= pdfCount; n++) {
    const b = document.createElement('button'); b.type = 'button'; b.className = 'pg' + (n === pdfPage ? ' on' : '');
    b.textContent = n + (n === pdfAuto ? ' ★' : '');
    b.setAttribute('data-tip', 'Страница ' + n + (n === pdfAuto ? ': здесь программа сама нашла печать.' : '. Нажмите — программа будет искать печать на этой странице.'));
    b.addEventListener('click', () => { if (n !== pdfPage) showPdfPage(n).catch(e => setStatus('Ошибка: ' + e.message)); });
    el.pageBtns.appendChild(b);
  }
  const on = el.pageBtns.querySelector('.on'); if (on && on.scrollIntoView) on.scrollIntoView({ block: 'nearest', inline: 'center' });
  layout();
}

function imageToFull(file) {
  return new Promise((res, rej) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const k = Math.min(1, FULL_SIDE / Math.max(img.naturalWidth, img.naturalHeight));
      const c = document.createElement('canvas');
      c.width = Math.max(1, Math.round(img.naturalWidth * k));
      c.height = Math.max(1, Math.round(img.naturalHeight * k));
      const cx = c.getContext('2d');
      cx.fillStyle = '#fff'; cx.fillRect(0, 0, c.width, c.height);
      cx.imageSmoothingQuality = 'high';
      cx.drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(url);
      res(c);
    };
    img.onerror = () => { URL.revokeObjectURL(url); rej(new Error('Не удалось открыть изображение (формат не поддерживается).')); };
    img.src = url;
  });
}

/* Автообрезка по найденной печати и приведение к рабочему размеру */
function startFromFull(full, notes) {
  fullCanvas = full;
  loadNotes = (notes || []).slice();
  let work = full, region = null;
  if (zoneRegion) {
    region = { x: Math.round(zoneRegion.x), y: Math.round(zoneRegion.y), w: Math.max(64, Math.round(zoneRegion.w)), h: Math.max(64, Math.round(zoneRegion.h)) };
    loadNotes.push('Рабочая зона выбрана вручную — печать ищется только внутри рамки.');
  } else if (el.autoCrop.checked) {
    const qb = quickStampBox(full);
    if (qb && qb.areaRatio < 0.7) {
      const b = qb.box;
      region = { x: b.x0, y: b.y0, w: Math.max(64, b.x1 - b.x0), h: Math.max(64, b.y1 - b.y0) };
      loadNotes.push('Печать найдена автоматически, и вокруг неё вырезана область для обработки.');
    } else if (!qb && Math.max(full.width, full.height) > MAX_SIDE) {
      loadNotes.push('Круглая печать автоматически не найдена — обрабатывается всё изображение целиком.');
    }
  }
  lastRegion = region;
  const sw = region ? region.w : full.width, sh = region ? region.h : full.height;
  const k = Math.min(1, MAX_SIDE / Math.max(sw, sh));
  const c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(sw * k)); c.height = Math.max(1, Math.round(sh * k));
  const cx = c.getContext('2d');
  cx.fillStyle = '#fff'; cx.fillRect(0, 0, c.width, c.height);
  cx.imageSmoothingQuality = 'high';
  if (region) cx.drawImage(full, region.x, region.y, region.w, region.h, 0, 0, c.width, c.height);
  else cx.drawImage(full, 0, 0, c.width, c.height);
  origCanvas = c;
  resetEverything();
}

async function handleFile(file) {
  if (!file) return;
  const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
  const isImg = file.type.startsWith('image/');
  if (!isPdf && !isImg) { setStatus('Нужен файл изображения или PDF.'); return; }
  setStatus('Загрузка…');
  try {
    fileLabel = file.name.replace(/\.[^.]+$/, '') || 'stamp';
    el.fname.dataset.base = file.name;
    el.fname.textContent = file.name;
    zoneRegion = null;
    if (isPdf) {
      pdfAuto = 0;
      await openPdfTracked(file);
    } else {
      pdfDoc = null; updatePdfBar();
      startFromFull(await imageToFull(file), []);
    }
  } catch (err) {
    setStatus('Ошибка: ' + err.message);
  }
}
async function openPdfTracked(file) {
  const lib = await loadPdfJs();
  const data = await file.arrayBuffer();
  pdfDoc = await lib.getDocument({ data }).promise;
  pdfCount = pdfDoc.numPages;
  let best = 1;
  if (pdfCount > 1) {
    let bestScore = 0;
    const n = Math.min(pdfCount, 30);
    for (let p = 1; p <= n; p++) {
      setStatus('Поиск печати в PDF: страница ' + p + ' из ' + n + '…');
      const c = await renderPdfPage(p, 900);
      const qb = quickStampBox(c);
      if (qb && qb.score > bestScore) { bestScore = qb.score; best = p; }
    }
    pdfAuto = bestScore > 0 ? best : 0;
  }
  await showPdfPage(best);
}

/* =====================================================================
   Подготовка источника
   ===================================================================== */
function setSource(canvas, opts) {
  opts = opts || {};
  src = canvas; W = canvas.width; H = canvas.height;
  srcPix = canvas.getContext('2d').getImageData(0, 0, W, H).data;
  el.canvas.width = W; el.canvas.height = H;
  el.orig.width = W; el.orig.height = H;
  imgOut = ctx.createImageData(W, H);
  flat = null; flatKey = -1; bbox = null; bboxDirty = true; ana = null;
  grid = makeGrid(srcPix, W, H);
  el.drop.style.display = 'none';
  el.panes.style.display = 'flex';
  layout();
  redrawOrig();
  if (opts.auto) autoAll(true);
  else if (opts.reanalyze) {
    analyze();
    if (opts.finish) finishAuto(); else { process(); }
  } else { process(); }
}

function redrawOrig() {
  if (!src) return;
  octx.clearRect(0, 0, el.orig.width, el.orig.height);
  octx.drawImage(src, 0, 0);
}
function layout() {
  if (!src) return;
  const stacked = getComputedStyle(el.panes).flexDirection === 'column';   // совпадает с адаптивной вёрсткой в CSS
  const aw = stacked ? el.workspace.clientWidth - 40 : Math.floor((el.workspace.clientWidth - 60 - 20) / 2) - 4;
  const extra = el.workspace.classList.contains('has-pages') ? 64 : 0;
  const ah = stacked ? Math.max(260, window.innerHeight / window.UIZoom.k() * 0.5) : el.workspace.clientHeight - 60 - 34 - 4 - extra;
  dispScale = Math.max(0.05, Math.min(aw / W, ah / H, 3));
  for (const c of [el.canvas, el.orig]) {
    c.style.width = Math.round(W * dispScale) + 'px';
    c.style.height = Math.round(H * dispScale) + 'px';
  }
}

function ensureFlat(shadow) {
  if (flat && flatKey === shadow) return;
  flat = makeFlat(srcPix, W, H, grid, clamp(shadow / 100, 0, 1));
  flatKey = shadow;
}

/* ---------- быстрые оттенки ---------- */
function presetHit(key, lum, ch, dark, hue) {
  const neutral = ch < 0.2 * dark + 6;               // почти бесцветный пиксель
  switch (key) {
    case 'black': return neutral && lum < 200;
    case 'gray':  return neutral && lum >= 90 && lum < 240;
    default: {
      if (neutral || ch < 10) return false;          // оттенок есть только у цветных пикселей
      if (key === 'red')    return hue >= 340 || hue < 18;
      if (key === 'orange') return hue >= 18 && hue < 70;
      if (key === 'green')  return hue >= 70 && hue < 165;
      if (key === 'cyan')   return hue >= 165 && hue < 200;
      if (key === 'blue')   return hue >= 200 && hue < 250;
      if (key === 'violet') return hue >= 250 && hue < 340;
      return false;
    }
  }
}

/* =====================================================================
   Анализ: контур печати, цвета красок, подпись
   ===================================================================== */
function discMask(fit, f) {
  const m = new Uint8Array(W * H);
  const cs = Math.cos(fit.th), sn = Math.sin(fit.th), ia = 1 / (fit.a * f), ib = 1 / (fit.b * f);
  for (let y = 0; y < H; y++) {
    const dy = y - fit.cy;
    for (let x = 0; x < W; x++) {
      const dx = x - fit.cx;
      const u = (dx * cs + dy * sn) * ia, v = (-dx * sn + dy * cs) * ib;
      if (u * u + v * v <= 1) m[y * W + x] = 1;
    }
  }
  return m;
}

/* k-means по направлениям поглощения света (не зависят от разбавления белым) */
function clusterDirs(V, n) {
  const dist = (i, c) => Math.hypot(V[i * 3] - c[0], V[i * 3 + 1] - c[1], V[i * 3 + 2] - c[2]);
  let mean = [0, 0, 0];
  for (let i = 0; i < n; i++) { mean[0] += V[i * 3]; mean[1] += V[i * 3 + 1]; mean[2] += V[i * 3 + 2]; }
  mean = mean.map(v => v / n);
  let first = 0, bd = 1e9;
  for (let i = 0; i < n; i++) { const d = dist(i, mean); if (d < bd) { bd = d; first = i; } }
  const cents = [[V[first * 3], V[first * 3 + 1], V[first * 3 + 2]]];
  for (let k = 1; k < 3; k++) {
    let far = -1, fd = 0;
    for (let i = 0; i < n; i++) {
      let d = 1e9; for (const c of cents) d = Math.min(d, dist(i, c));
      if (d > fd) { fd = d; far = i; }
    }
    if (far < 0 || fd < 0.12) break;
    cents.push([V[far * 3], V[far * 3 + 1], V[far * 3 + 2]]);
  }
  const K = cents.length, cnt = new Int32Array(K);
  for (let it = 0; it < 8; it++) {
    const sum = cents.map(() => [0, 0, 0]); cnt.fill(0);
    for (let i = 0; i < n; i++) {
      let bk = 0, bdd = 1e9;
      for (let k = 0; k < K; k++) { const d = dist(i, cents[k]); if (d < bdd) { bdd = d; bk = k; } }
      sum[bk][0] += V[i * 3]; sum[bk][1] += V[i * 3 + 1]; sum[bk][2] += V[i * 3 + 2]; cnt[bk]++;
    }
    for (let k = 0; k < K; k++) {
      if (!cnt[k]) continue;
      const m = Math.hypot(sum[k][0], sum[k][1], sum[k][2]) || 1;
      cents[k] = [sum[k][0] / m, sum[k][1] / m, sum[k][2] / m];
    }
  }
  let protos = cents.map((v, k) => ({ v, n: cnt[k] })).filter(p => p.n > 0);
  // близкие цвета объединяем: это одна и та же краска разной плотности
  let merged = true;
  while (merged && protos.length > 1) {
    merged = false;
    outer: for (let i = 0; i < protos.length; i++) for (let j = i + 1; j < protos.length; j++) {
      const a = protos[i], b = protos[j];
      if (Math.hypot(a.v[0] - b.v[0], a.v[1] - b.v[1], a.v[2] - b.v[2]) < 0.20) {
        const w = a.n + b.n, v = [0, 1, 2].map(t => (a.v[t] * a.n + b.v[t] * b.n) / w);
        const m = Math.hypot(v[0], v[1], v[2]) || 1;
        protos.splice(j, 1); protos[i] = { v: v.map(t => t / m), n: w };
        merged = true; break outer;
      }
    }
  }
  protos.sort((a, b) => b.n - a.n);
  const total = protos.reduce((s, p) => s + p.n, 0);
  const keep = protos.filter((p, i) => i === 0 || p.n / total >= 0.02);
  return keep;
}

function analyze() {
  ana = null;
  if (!src) return;
  const p = P(); ensureFlat(p.shadow);
  const N = W * H, T = thrOf(p.th), cm = 1 + 2 * p.sens / 100;
  const sArr = new Uint8ClampedArray(N);
  let Tc = Math.max(2 * T, 70), cntCore = 0;
  for (let idx = 0; idx < N; idx++) {
    const o = idx * 3, r = flat[o], g = flat[o + 1], b = flat[o + 2];
    const mx = r > g ? (r > b ? r : b) : (g > b ? g : b), mn = r < g ? (r < b ? r : b) : (g < b ? g : b);
    const dark = 255 - (0.299 * r + 0.587 * g + 0.114 * b), ch = (mx - mn) * cm;
    const s = dark > ch ? dark : ch;
    sArr[idx] = s; if (s >= Tc) cntCore++;
  }
  if (cntCore < 300) { Tc = T * 1.3; cntCore = 0; for (let i = 0; i < N; i++) if (sArr[i] >= Tc) cntCore++; }
  if (cntCore < 60) { ana = { fit: null, protos: [], stampPts: new Int32Array(0) }; return; }

  const step = Math.max(1, Math.floor(cntCore / 20000));
  const V = new Float32Array(Math.ceil(cntCore / step) * 3 + 3);
  let ns = 0, ctr = 0;
  for (let idx = 0; idx < N; idx++) {
    if (sArr[idx] < Tc) continue;
    if (ctr++ % step) continue;
    const o = idx * 3, w0 = 255 - flat[o], w1 = 255 - flat[o + 1], w2 = 255 - flat[o + 2];
    const len = Math.hypot(w0, w1, w2) || 1;
    V[ns * 3] = w0 / len; V[ns * 3 + 1] = w1 / len; V[ns * 3 + 2] = w2 / len; ns++;
  }
  const protos = clusterDirs(V, ns);

  // класс каждого пикселя краски — ближайший цвет
  const cls = new Uint8Array(N).fill(255), lenA = new Uint8ClampedArray(N);
  const K = protos.length;
  for (let idx = 0; idx < N; idx++) {
    if (sArr[idx] < T) continue;
    const o = idx * 3, w0 = 255 - flat[o], w1 = 255 - flat[o + 1], w2 = 255 - flat[o + 2];
    const len = Math.hypot(w0, w1, w2) || 1;
    lenA[idx] = len;
    if (K === 1) { cls[idx] = 0; continue; }
    const v0 = w0 / len, v1 = w1 / len, v2 = w2 / len;
    let bk = 0, bd = 1e9;
    for (let k = 0; k < K; k++) {
      const q = protos[k].v, d = (v0 - q[0]) * (v0 - q[0]) + (v1 - q[1]) * (v1 - q[1]) + (v2 - q[2]) * (v2 - q[2]);
      if (d < bd) { bd = d; bk = k; }
    }
    cls[idx] = bk;
  }

  // контур печати по каждому цвету (в уменьшенном масштабе)
  const ds = Math.max(1, Math.ceil(Math.max(W, H) / 700)), w2 = Math.ceil(W / ds), h2 = Math.ceil(H / ds);
  const ringFor = c => {
    const m = new Uint8Array(w2 * h2);
    for (let y = 0; y < H; y++) {
      const row = y * W, r2 = ((y / ds) | 0) * w2;
      for (let x = 0; x < W; x++) {
        const v = cls[row + x];
        if (c < 0 ? v !== 255 : v === c) m[r2 + ((x / ds) | 0)] = 1;
      }
    }
    const d = morph(m, w2, h2, 1, true);
    const { labels, comps } = components(d, w2, h2);
    let big = null;
    for (let i = 1; i < comps.length; i++) if (!big || comps[i].n > big.n) big = comps[i];
    if (!big || big.n < 30) return null;
    const f = fitEllipse(labels, big.id, big, w2, h2);
    if (!f) return null;
    return { cx: (f.cx + 0.5) * ds - 0.5, cy: (f.cy + 0.5) * ds - 0.5, a: f.a * ds, b: f.b * ds, th: f.th, q: f.q };
  };
  let stampIdx = 0, fit = null;
  if (K > 1) {
    let bestScore = 0;
    for (let c = 0; c < K; c++) {
      const f = ringFor(c);
      if (!f || f.q >= 0.10 || f.a / f.b > 3) continue;
      const sc = Math.sqrt(f.a * f.b) * (1 - f.q / 0.10);
      if (sc > bestScore) { bestScore = sc; stampIdx = c; fit = f; }
    }
    if (!fit) fit = ringFor(-1);
  } else fit = ringFor(0);

  const stampPts = [];
  if (fit) {
    const cs = Math.cos(fit.th), sn = Math.sin(fit.th);
    const stp = Math.max(1, Math.floor(Math.sqrt(Math.PI * fit.a * fit.b * 0.1 / 7000)));
    const x0 = Math.max(0, Math.floor(fit.cx - Math.max(fit.a, fit.b) * 1.3)), x1 = Math.min(W - 1, Math.ceil(fit.cx + Math.max(fit.a, fit.b) * 1.3));
    const y0 = Math.max(0, Math.floor(fit.cy - Math.max(fit.a, fit.b) * 1.3)), y1 = Math.min(H - 1, Math.ceil(fit.cy + Math.max(fit.a, fit.b) * 1.3));
    for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) {
      const i = y * W + x, c = cls[i]; if (c === 255 || sArr[i] < Tc * 0.8) continue;
      const dx = x - fit.cx, dy = y - fit.cy;
      const u = (dx * cs + dy * sn) / fit.a, v = (-dx * sn + dy * cs) / fit.b, rho = Math.sqrt(u * u + v * v);
      if (rho < 0.62 && c === stampIdx && ((x + y * 7) % stp === 0) && stampPts.length < 9000) stampPts.push(x, y);
    }
  }

  ana = { protos, stampIdx, fit, stampPts: new Int32Array(stampPts) };
}

/* =====================================================================
   Основная обработка
   ===================================================================== */
function process() {
  if (!src) return;
  const p = P(), N = W * H;
  ensureFlat(p.shadow);
  const T = thrOf(p.th), soft = Math.max(10, T * 0.6), cm = 1 + 2 * p.sens / 100;
  const q0 = 0.06 + p.text * 0.0022, q1 = q0 + 0.12;                   // граница «текст / цветной оттиск» по отношению цветность / тёмность
  const presets = Array.from(presetsOn), nPre = presets.length;
  const nPick = picks.length;
  const perpMax = p.tol * 1.2;                 // допуск по направлению цвета (в единицах 0..255)

  const clipM = (el.clip.checked && ana && ana.fit && ana.fit.q < 0.12) ? discMask(ana.fit, 1.12) : null;

  const a = new Uint8ClampedArray(N), perm = new Uint8ClampedArray(N), sc = new Uint8ClampedArray(N);
  const aN = p.rm ? new Uint8ClampedArray(N) : null;   // то же самое, но без удаления чёрного и серого текста — для инструмента «Маска»
  const od = imgOut.data;

  for (let idx = 0; idx < N; idx++) {
    const o = idx * 3, i = idx * 4;
    const r = flat[o], g = flat[o + 1], b = flat[o + 2];
    const mx = r > g ? (r > b ? r : b) : (g > b ? g : b);
    const mn = r < g ? (r < b ? r : b) : (g < b ? g : b);
    const ch = mx - mn;
    const lum = 0.299 * r + 0.587 * g + 0.114 * b;
    const dark = 255 - lum;
    const s = dark > ch * cm ? dark : ch * cm;
    sc[idx] = s;
    let al = smooth((s - T) / soft), alN = al;
    perm[idx] = smooth((s - T * 0.4) / (soft * 0.6)) * 255;
    if (alN > 0) {
      if (p.rm) al *= smooth((ch / (dark + 1) - q0) / (q1 - q0)) * smooth((ch - 4) / 8);
      if (clipM && !clipM[idx]) { al = 0; alN = 0; }
      if (nPre && alN > 0) {
        let hue = 0;
        if (ch > 0) {
          if (mx === r) hue = ((g - b) / ch) % 6; else if (mx === g) hue = (b - r) / ch + 2; else hue = (r - g) / ch + 4;
          hue *= 60; if (hue < 0) hue += 360;
        }
        for (let k = 0; k < nPre; k++) if (presetHit(presets[k], lum, ch, dark, hue)) { al = 0; alN = 0; break; }
      }
      if (nPick && alN > 0) {
        // сравниваем «направление» цвета относительно белой бумаги: оно не меняется на полупрозрачных краях линий
        const w0 = 255 - r, w1 = 255 - g, w2 = 255 - b;
        let mm = 0;
        for (let k = 0; k < nPick; k++) {
          const v = picks[k].v;                       // единичный вектор выбранного цвета
          const dp = w0 * v[0] + w1 * v[1] + w2 * v[2];
          const e0 = w0 - dp * v[0], e1 = w1 - dp * v[1], e2 = w2 - dp * v[2];
          const perp = Math.sqrt(e0 * e0 + e1 * e1 + e2 * e2);
          const m = dp < 0 ? 0 : 1 - smooth((perp - perpMax * 0.6) / (perpMax * 0.4 + 0.001));
          if (m > mm) mm = m;
        }
        al *= 1 - mm; alN *= 1 - mm;
      }
    }
    a[idx] = al * 255;
    if (aN) aN[idx] = alN * 255;
    // цвет: убираем примесь белой бумаги на полупрозрачных краях (по «полной» плотности — чтобы вернувшийся маской текст имел верный цвет)
    const ia = alN > 0.001 ? 1 / (alN > 0.35 ? alN : 0.35) : 1;
    od[i]     = 255 - (255 - r) * ia;
    od[i + 1] = 255 - (255 - g) * ia;
    od[i + 2] = 255 - (255 - b) * ia;
  }

  const minSpeck = Math.max(10, Math.round(N * 0.00001));
  removeSpecks(a, minSpeck);
  const alpha = boxBlur(a, p.smooth);
  let alphaNoText = alpha;
  if (aN) { removeSpecks(aN, minSpeck); alphaNoText = boxBlur(aN, p.smooth); }
  base = { alpha, alphaNoText, perm, sc, T, T2: T * 2 + 20 };
  render();
}

function removeSpecks(a, minSize) {
  const N = W * H, vis = new Uint8Array(N), q = new Int32Array(N);
  for (let s = 0; s < N; s++) {
    if (vis[s] || a[s] < 40) continue;
    let head = 0, tail = 0; q[tail++] = s; vis[s] = 1;
    while (head < tail) {
      const i = q[head++], x = i % W, y = (i - x) / W;
      for (let dy = -1; dy <= 1; dy++) {
        const ny = y + dy; if (ny < 0 || ny >= H) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const nx = x + dx; if (nx < 0 || nx >= W) continue;
          const j = ny * W + nx;
          if (!vis[j] && a[j] >= 40) { vis[j] = 1; q[tail++] = j; }
        }
      }
    }
    if (tail < minSize) for (let k = 0; k < tail; k++) a[q[k]] = 0;
  }
}

function boxBlur(a, r) {
  if (!r) return a;
  const win = 2 * r + 1, t = new Uint8ClampedArray(a.length), o = new Uint8ClampedArray(a.length);
  for (let y = 0; y < H; y++) {
    const row = y * W; let sum = 0;
    for (let x = -r; x <= r; x++) sum += a[row + clamp(x, 0, W - 1)];
    for (let x = 0; x < W; x++) {
      t[row + x] = sum / win;
      sum += a[row + Math.min(x + r + 1, W - 1)] - a[row + Math.max(x - r, 0)];
    }
  }
  for (let x = 0; x < W; x++) {
    let sum = 0;
    for (let y = -r; y <= r; y++) sum += t[clamp(y, 0, H - 1) * W + x];
    for (let y = 0; y < H; y++) {
      o[y * W + x] = sum / win;
      sum += t[Math.min(y + r + 1, H - 1) * W + x] - t[Math.max(y - r, 0) * W + x];
    }
  }
  return o;
}

/* =====================================================================
   Автоматика
   ===================================================================== */
function autoCalibrate() {
  const shadow = 85;
  ensureFlat(shadow);
  const hist = new Uint32Array(256);
  let total = 0;
  for (let y = 0; y < H; y += 2) for (let x = 0; x < W; x += 2) {
    const o = (y * W + x) * 3, r = flat[o], g = flat[o + 1], b = flat[o + 2];
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
    const dark = 255 - (0.299 * r + 0.587 * g + 0.114 * b);
    const s = Math.min(255, Math.max(dark, (mx - mn) * 1.7));
    hist[s | 0]++; total++;
  }
  const tOtsu = otsu(hist, total);
  const T = clamp(tOtsu * 0.55, 14, 85);
  // цветной оттиск или чёрно-серый? доля цветных пикселей среди «краски»
  let ink = 0, chrom = 0;
  for (let y = 0; y < H; y += 2) for (let x = 0; x < W; x += 2) {
    const o = (y * W + x) * 3, r = flat[o], g = flat[o + 1], b = flat[o + 2];
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
    const dark = 255 - (0.299 * r + 0.587 * g + 0.114 * b);
    if (Math.max(dark, (mx - mn) * 1.7) >= tOtsu) { ink++; if (mx - mn >= 40) chrom++; }
  }
  const colored = ink > 30 && chrom / ink >= 0.3;

  el.pShadow.value = shadow;
  el.pBright.value = Math.round(clamp((T - 5) / 1.1, 0, 100));
  el.pChroma.value = 35;
  el.pSmooth.value = 0;                 // сглаживание краёв по умолчанию — минимум
  el.removeText.checked = false;        // «Убрать чёрный и серый текст» по умолчанию выключено
  el.pText.value = 55;
  presetsOn.clear(); picks = [];
  for (const b of el.presetChips.children) b.classList.remove('on');
  renderPicks();
  updateLabels();
  procNotes.push('Убраны фон, блики и тени.');
  if (colored) procNotes.push('Оттиск цветной: если вокруг остался чёрный или серый текст, включите «Убрать чёрный и серый текст» на верхней панели.');
}

function autoAll(allowRound) {
  procNotes = [];
  autoCalibrate();
  analyze();
  if (allowRound && el.autoRound.checked && ana && ana.fit && ana.fit.q < 0.06) {
    const ratio = Math.max(ana.fit.a, ana.fit.b) / Math.min(ana.fit.a, ana.fit.b);
    if (ratio >= 1.04 && ratio < 2.5 && makeRound(true)) return;
  }
  finishAuto();
}

function finishAuto() {
  process();
  showAutoInfo();
  setStatus('Файл загружен и очищен автоматически. Слева — исходный вид, справа — результат.');
}

function showAutoInfo() {
  el.autoInfo.innerHTML = '';
  const notes = loadNotes.concat(procNotes);
  if (!notes.length) { el.autoInfo.textContent = 'После загрузки файла печать очищается автоматически.'; return; }
  for (const n of notes) {
    const d = document.createElement('div'); d.textContent = '• ' + n; el.autoInfo.appendChild(d);
  }
  const d = document.createElement('div'); d.className = 'auto-more';
  d.textContent = 'Если результат не устраивает, воспользуйтесь инструментами сверху и «Тонкой настройкой».';
  el.autoInfo.appendChild(d);
}

/* ---------- ручные правки ---------- */
function applyCircle(a, x, y, r, type, sens) {
  const x0 = Math.max(0, Math.floor(x - r)), x1 = Math.min(W - 1, Math.ceil(x + r));
  const y0 = Math.max(0, Math.floor(y - r)), y1 = Math.min(H - 1, Math.ceil(y + r));
  const r2 = r * r, sc = base.sc, perm = base.perm, T2 = base.T2 * edgeFactor(sens == null ? 50 : sens), noText = base.alphaNoText;
  for (let yy = y0; yy <= y1; yy++) {
    const dy = yy - y;
    for (let xx = x0; xx <= x1; xx++) {
      const dx = xx - x;
      if (dx * dx + dy * dy > r2) continue;
      const i = yy * W + xx;
      if (type === 'eraser') a[i] = 0;
      else if (type === 'mask') { if (sc[i] < T2) a[i] = 0; }            // «Уточнение края»: остаётся только плотный оттиск
      else if (type === 'textmask') { if (noText[i] > a[i]) a[i] = noText[i]; }   // «Маска»: возвращаем чёрный и серый текст, убранный галочкой
      else if (type === 'restore') { if (perm[i] > a[i]) a[i] = perm[i]; }  // возвращаем убранное очисткой
    }
  }
}

function render() {
  if (!base) return;
  cur = base.alpha.slice();
  for (const e of edits) for (const pt of e.pts) applyCircle(cur, pt[0], pt[1], e.r, e.type, e.sens);
  bboxDirty = true;
  paint();
}

function computeBBox() {
  let x0 = W, y0 = H, x1 = -1, y1 = -1;
  for (let y = 0; y < H; y++) {
    const row = y * W;
    for (let x = 0; x < W; x++) {
      if (cur[row + x] > 40) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    }
  }
  bbox = x1 < 0 ? null : { x0, y0, x1, y1, w: x1 - x0 + 1, h: y1 - y0 + 1 };
  bboxDirty = false;
}

function paint() {
  const out = imgOut.data, N = W * H;
  for (let i = 0; i < N; i++) out[i * 4 + 3] = cur[i];
  ctx.putImageData(imgOut, 0, 0);
}

let paintQueued = false;
function schedulePaint() {
  if (paintQueued) return; paintQueued = true;
  requestAnimationFrame(() => { paintQueued = false; if (base) paint(); });
}
let procQueued = false;
function scheduleProcess() {
  if (procQueued) return; procQueued = true;
  requestAnimationFrame(() => { procQueued = false; process(); });
}

/* =====================================================================
   Круглая печать: эллипс → круг, затем разворот по тексту в центре
   ===================================================================== */
function fitFromCov() {
  let n = 0, sx = 0, sy = 0, sxx = 0, syy = 0, sxy = 0;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if (cur[y * W + x] > 128) { n++; sx += x; sy += y; sxx += x * x; syy += y * y; sxy += x * y; }
  }
  if (n < 50) return null;
  const mx = sx / n, my = sy / n;
  const cxx = sxx / n - mx * mx, cyy = syy / n - my * my, cxy = sxy / n - mx * my;
  const th = 0.5 * Math.atan2(2 * cxy, cxx - cyy);
  const c = Math.cos(th), s = Math.sin(th);
  let umin = 1e9, umax = -1e9, vmin = 1e9, vmax = -1e9;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if (cur[y * W + x] > 128) {
      const dx = x - mx, dy = y - my;
      const u = dx * c + dy * s, v = -dx * s + dy * c;
      if (u < umin) umin = u; if (u > umax) umax = u;
      if (v < vmin) vmin = v; if (v > vmax) vmax = v;
    }
  }
  const cu = (umax + umin) / 2, cv = (vmax + vmin) / 2;
  return { cx: mx + cu * c - cv * s, cy: my + cu * s + cv * c, a: (umax - umin) / 2, b: (vmax - vmin) / 2, th, q: 1 };
}

/* Угол разворота: текст в центре печати должен лечь в горизонтальные строки */
function findRotation(fit, pts) {
  if (!pts || pts.length < 400) return 0;
  const R = Math.max(fit.a, fit.b), cs = Math.cos(fit.th), sn = Math.sin(fit.th);
  const n = pts.length / 2, ux = new Float32Array(n), uy = new Float32Array(n);
  for (let i = 0; i < n; i++) {
    const dx = pts[i * 2] - fit.cx, dy = pts[i * 2 + 1] - fit.cy;
    const tu = (dx * cs + dy * sn) * R / fit.a, tv = (-dx * sn + dy * cs) * R / fit.b;
    ux[i] = (tu * cs - tv * sn) / R; uy[i] = (tu * sn + tv * cs) / R;
  }
  const BINS = 64, hist = new Float32Array(BINS), scores = [];
  let best = -1, bestA = 0;
  for (let a = -45; a <= 45; a += 0.5) {
    const ph = a * Math.PI / 180, sp = Math.sin(ph), cp = Math.cos(ph);
    hist.fill(0);
    for (let i = 0; i < n; i++) {
      const y = -ux[i] * sp + uy[i] * cp;
      const b = Math.floor((y + 0.65) / 1.3 * BINS);
      if (b >= 0 && b < BINS) hist[b]++;
    }
    let sc = 0; for (let b = 0; b < BINS; b++) sc += hist[b] * hist[b];
    scores.push(sc);
    if (sc > best) { best = sc; bestA = a; }
  }
  const sorted = scores.slice().sort((x, y) => x - y), med = sorted[sorted.length >> 1];
  if (best / med < 1.12 || Math.abs(bestA) < 0.7) return 0;
  return bestA;
}

function makeRound(auto) {
  auto = auto === true;
  if (!src || (!base && !auto)) { setStatus('Сначала загрузите файл.'); return false; }
  let fit = ana && ana.fit && ana.fit.q < 0.12 ? ana.fit : null;
  if (!fit && cur && !auto) fit = fitFromCov();
  if (!fit) { if (!auto) setStatus('Печать на изображении не найдена — подстройте параметры.'); return false; }
  const R = Math.max(fit.a, fit.b);
  const ang = findRotation(fit, ana && ana.stampPts);
  const ph = ang * Math.PI / 180, cph = Math.cos(ph), sph = Math.sin(ph);
  const cs = Math.cos(fit.th), sn = Math.sin(fit.th), ka = fit.a / R, kb = fit.b / R;
  const q = Math.min(1, 2400 / (2 * R * 1.2));
  const S = Math.ceil(2 * R * 1.2 * q);
  const out = new ImageData(S, S), od = out.data;
  for (let y = 0; y < S; y++) for (let x = 0; x < S; x++) {
    const px = (x - S / 2 + 0.5) / q, py = (y - S / 2 + 0.5) / q;
    const rx = px * cph - py * sph, ry = px * sph + py * cph;                 // поворот по тексту
    const tu = (rx * cs + ry * sn) * ka, tv = (-rx * sn + ry * cs) * kb;       // вдоль осей эллипса
    const sx0 = clamp(fit.cx + tu * cs - tv * sn, 0, W - 1.001);
    const sy0 = clamp(fit.cy + tu * sn + tv * cs, 0, H - 1.001);
    const ix = Math.floor(sx0), iy = Math.floor(sy0), fx = sx0 - ix, fy = sy0 - iy;
    const i00 = (iy * W + ix) * 4, i01 = i00 + 4, i10 = i00 + W * 4, i11 = i10 + 4;
    const o = (y * S + x) * 4;
    for (let k = 0; k < 3; k++) {
      od[o + k] = (srcPix[i00 + k] * (1 - fx) + srcPix[i01 + k] * fx) * (1 - fy) +
                  (srcPix[i10 + k] * (1 - fx) + srcPix[i11 + k] * fx) * fy;
    }
    od[o + 3] = 255;
  }
  const c2 = document.createElement('canvas'); c2.width = S; c2.height = S;
  c2.getContext('2d').putImageData(out, 0, 0);
  const ratio = R / Math.min(fit.a, fit.b);
  const note = 'Печать была искажена (оси отличались в ' + ratio.toFixed(2).replace('.', ',') + ' раза) — приведена к кругу' +
    (ang ? ', строки текста внутри выровнены (поворот ' + ang.toFixed(1).replace('.', ',') + '°).' : '; строки текста внутри уже горизонтальны.');
  if (auto) {
    procNotes.push(note);
    setSource(c2, { reanalyze: true, finish: true });
  } else {
    history.push({ type: 'round', prevSrc: src, prevEdits: edits });
    edits = [];
    setSource(c2, { reanalyze: true });
    setStatus(note + ' Ручные правки кистью сброшены — лучше применять эту функцию до ручной чистки.');
  }
  return true;
}

/* ---------- отмена ---------- */
function undo() {
  const h = history.pop();
  if (!h) { setStatus('Нечего отменять.'); return; }
  if (h.type === 'edit') { edits.pop(); render(); }
  else if (h.type === 'round') { edits = h.prevEdits; setSource(h.prevSrc, { reanalyze: true }); }
  setStatus('Последнее действие отменено.');
}

function resetEverything() {
  if (!origCanvas) return;
  edits = []; history = [];
  el.pPick.value = 25;
  const c = document.createElement('canvas');
  c.width = origCanvas.width; c.height = origCanvas.height;
  c.getContext('2d').drawImage(origCanvas, 0, 0);
  setSource(c, { auto: true });
}

/* ---------- скачивание ---------- */
const crcT = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
function crc32(buf) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) c = crcT[(c ^ buf[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function addDpi(bytes, dpi) {            // вставка блока pHYs сразу после IHDR (смещение 33)
  const ppm = Math.round(dpi / 0.0254);
  const chunk = new Uint8Array(21), dv = new DataView(chunk.buffer);
  dv.setUint32(0, 9); chunk.set([0x70, 0x48, 0x59, 0x73], 4);
  dv.setUint32(8, ppm); dv.setUint32(12, ppm); chunk[16] = 1;
  dv.setUint32(17, crc32(chunk.subarray(4, 17)));
  const out = new Uint8Array(bytes.length + 21);
  out.set(bytes.subarray(0, 33)); out.set(chunk, 33); out.set(bytes.subarray(33), 54);
  return out;
}

async function buildExport() {
  if (!base) return null;
  // размер считаем по результату до состаривания, чтобы диаметр не «плавал» от выпавших краёв
  let x0 = W, y0 = H, x1 = -1, y1 = -1;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if (cur[y * W + x] > 40) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  }
  if (x1 < 0) return null;
  let out, ratio = 1;
  {
    const bw = x1 - x0 + 1, bh = y1 - y0 + 1;
    const k = TARGET_PX / Math.max(bw, bh);
    out = document.createElement('canvas'); out.width = TARGET_PX; out.height = TARGET_PX;
    const oc = out.getContext('2d');
    oc.imageSmoothingEnabled = true; oc.imageSmoothingQuality = 'high';
    const dw = bw * k, dh = bh * k;
    oc.drawImage(el.canvas, x0, y0, bw, bh, (TARGET_PX - dw) / 2, (TARGET_PX - dh) / 2, dw, dh);
    ratio = Math.max(bw, bh) / Math.min(bw, bh);
  }
  const blob = await new Promise(res => out.toBlob(res, 'image/png'));
  const bytes = addDpi(new Uint8Array(await blob.arrayBuffer()), DPI);
  return { bytes, ratio, w: out.width, h: out.height };
}

async function download() {
  if (!base) { setStatus('Сначала загрузите файл.'); return; }
  const r = await buildExport();
  if (!r) { setStatus('Результат пуст — нечего скачивать.'); return; }
  const url = URL.createObjectURL(new Blob([r.bytes], { type: 'image/png' }));
  const a = document.createElement('a');
  a.href = url; a.download = fileLabel + '_stamp_40mm.png';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  setStatus('PNG сохранён: ' + TARGET_PX + '×' + TARGET_PX + ' px, ' + DPI + ' dpi, диаметр 40 мм.' +
      (r.ratio > 1.05 ? ' Печать вытянута — нажмите «Сделать оттиск круглым» для точного круга.' : ''));
}

/* =====================================================================
   Интерфейс
   ===================================================================== */
function updateLabels() {
  $('val1').textContent = el.pBright.value;
  $('val2').textContent = el.pChroma.value;
  $('val3').textContent = el.pShadow.value;
  $('val4').textContent = el.pSmooth.value;
  $('valText').textContent = el.pText.value;
  $('valEdgeSens').textContent = el.pEdge.value;
  $('tbTextPower').classList.toggle('off', !el.removeText.checked);   // ползунок «Сила лишнего текста» работает только при включённой галочке
  $('valPick').textContent = el.pPick.value;
  el.brushVal.textContent = el.brush.value + ' px';
}

function setTool(t) {
  tool = tool === t ? null : t;
  el.btnMask.classList.toggle('tool-active', tool === 'mask');
  el.btnEraser.classList.toggle('tool-active', tool === 'eraser');
  el.btnRestore.classList.toggle('tool-active', tool === 'restore');
  el.btnTextMask.classList.toggle('tool-active', tool === 'textmask');
  el.btnPick.classList.toggle('tool-active', tool === 'pick');
  if (tool === 'eraser') { el.brush.min = 1; el.brush.max = 80; }
  else { el.brush.min = 5; el.brush.max = 180; }
  if (tool && sizes[tool]) el.brush.value = sizes[tool];
  updateLabels();
  const hints = { mask: '<b>Уточнение края:</b> проведите по результату — внутри круга останется только плотный оттиск. Насколько «строго» — регулирует ползунок «Чувствительность» под кнопкой.', textmask: '<b>Маска:</b> нажмите или проведите по результату — в этом месте вернётся чёрный и серый текст, который убрала галочка «Убрать чёрный и серый текст». Ширина круга — ползунок «Размер» или колесо мыши.', eraser: '<b>Ластик:</b> проведите по результату — всё под кругом сотрётся.', restore: '<b>Вернуть:</b> проведите по результату — вернётся то, что убрала очистка.', pick: '<b>Пипетка:</b> кликните по цвету, который нужно убрать.' };
  const th = document.getElementById('toolHint'); if (th) th.innerHTML = hints[tool] || 'Выберите инструмент, затем проведите по <b>правому</b> изображению (результату).';
  el.canvas.style.cursor = tool === 'pick' ? 'crosshair' : (isBrush() ? 'none' : 'default');
  el.orig.style.cursor = tool === 'pick' ? 'crosshair' : 'default';
  if (!isBrush()) el.cursor.style.display = 'none';
  el.cursor2.style.display = 'none';
  if (tool === 'pick') setStatus('Пипетка: кликните по цвету, который нужно удалить — на левом или правом изображении.');
  if (tool === 'textmask') setStatus(el.removeText.checked ? 'Маска: нажмите на результат там, где нужно вернуть чёрный и серый текст.' : 'Маска действует вместе с галочкой «Убрать чёрный и серый текст»: сначала включите её, затем возвращайте текст маской в нужных местах.');
}

const isBrush = () => tool === 'mask' || tool === 'eraser' || tool === 'restore' || tool === 'textmask';

function moveCursorOn(e, canvas, cursorEl, d, color) {
  const cr = canvas.getBoundingClientRect();
  cursorEl.style.display = 'block';
  cursorEl.style.width = d + 'px'; cursorEl.style.height = d + 'px';
  const zk = window.UIZoom.k();                    // масштаб сайта 80 %
  cursorEl.style.left = ((e.clientX - cr.left) / zk) + 'px';
  cursorEl.style.top = ((e.clientY - cr.top) / zk) + 'px';
  cursorEl.style.borderColor = color;
}
function moveCursor(e) {
  if (!isBrush()) return;
  moveCursorOn(e, el.canvas, el.cursor, Math.max(3, sizes[tool]),
    tool === 'eraser' ? 'rgba(255,120,120,.95)' : (tool === 'restore' ? 'rgba(120,255,160,.95)' : (tool === 'textmask' ? 'rgba(255,200,70,.98)' : 'rgba(255,255,255,.85)')));
}
function imgPos(e, canvas) {
  const r = canvas.getBoundingClientRect();
  return [(e.clientX - r.left) * W / r.width, (e.clientY - r.top) * H / r.height];
}

function addPoint(e, canvas) {
  const [x, y] = imgPos(e, canvas);
  const put = (px, py) => { stroke.pts.push([px, py]); applyCircle(cur, px, py, stroke.r, stroke.type, stroke.sens); };
  if (lastPt) {
    const dx = x - lastPt[0], dy = y - lastPt[1], dist = Math.hypot(dx, dy);
    const step = Math.max(1, stroke.r * 0.5), n = Math.floor(dist / step);
    for (let i = 1; i <= n; i++) put(lastPt[0] + dx * i / n, lastPt[1] + dy * i / n);
    if (n === 0) put(x, y);
  } else put(x, y);
  lastPt = [x, y];
  schedulePaint();
}

/* пипетка */
function pickColor(e, canvas) {
  if (!flat) return;
  const [x, y] = imgPos(e, canvas);
  const cx = Math.round(x), cy = Math.round(y);
  let r = 0, g = 0, b = 0, n = 0;
  for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) {
    const xx = cx + dx, yy = cy + dy;
    if (xx < 0 || yy < 0 || xx >= W || yy >= H) continue;
    const o = (yy * W + xx) * 3; r += flat[o]; g += flat[o + 1]; b += flat[o + 2]; n++;
  }
  if (!n) return;
  r /= n; g /= n; b /= n;
  const w0 = 255 - r, w1 = 255 - g, w2 = 255 - b, len = Math.hypot(w0, w1, w2);
  if (len < 12) { setStatus('Это почти белый фон — выберите цвет на линии оттиска или текста.'); return; }
  picks.push({ v: [w0 / len, w1 / len, w2 / len], hex: 'rgb(' + Math.round(r) + ',' + Math.round(g) + ',' + Math.round(b) + ')' });
  renderPicks();
  scheduleProcess();
  setStatus('Цвет добавлен в список удаляемых. Можно выбрать ещё или изменить «Допуск оттенка».');
}
function renderPicks() {
  el.pickChips.innerHTML = '';
  picks.forEach((pk, i) => {
    const s = document.createElement('span'); s.className = 'pick-chip';
    s.innerHTML = '<i style="background:' + pk.hex + '"></i>цвет ' + (i + 1);
    const x = document.createElement('button'); x.textContent = '×'; x.title = 'Убрать из списка';
    x.addEventListener('click', () => { picks.splice(i, 1); renderPicks(); scheduleProcess(); });
    s.appendChild(x); el.pickChips.appendChild(s);
  });
}

el.canvas.addEventListener('pointerdown', e => {
  if (!base) return;
  if (tool === 'pick') { pickColor(e, el.canvas); return; }
  if (!isBrush()) return;
  e.preventDefault();
  try { el.canvas.setPointerCapture(e.pointerId); } catch (_) {}
  stroke = { type: tool, r: Math.max(0.5, sizes[tool] / 2 / dispScale), sens: +el.pEdge.value, pts: [] };   // размер задан в пикселях экрана
  lastPt = null;
  addPoint(e, el.canvas);
});
el.orig.addEventListener('pointerdown', e => { if (base && tool === 'pick') pickColor(e, el.orig); });
el.canvas.addEventListener('pointermove', e => {
  moveCursor(e);
  if (stroke) addPoint(e, el.canvas);
});
const endStroke = () => {
  if (!stroke) return;
  if (stroke.pts.length) { edits.push(stroke); history.push({ type: 'edit' }); }
  stroke = null; lastPt = null;
  bboxDirty = true; schedulePaint();
};
el.canvas.addEventListener('pointerup', endStroke);
el.canvas.addEventListener('pointercancel', endStroke);
el.canvas.addEventListener('pointerleave', () => { if (!stroke) el.cursor.style.display = 'none'; });
el.canvas.addEventListener('pointerenter', moveCursor);

/* размер ластика, «Уточнения края», «Маски» и «Вернуть» — колесом мыши над результатом */
el.canvas.addEventListener('wheel', e => {
  if (!isBrush()) return;
  e.preventDefault();
  const cur0 = sizes[tool], f = e.deltaY < 0 ? 1.12 : 1 / 1.12;
  let v = Math.round(cur0 * f); if (v === cur0) v += e.deltaY < 0 ? 1 : -1;
  v = clamp(v, +el.brush.min, +el.brush.max);
  sizes[tool] = v; el.brush.value = v; updateLabels(); moveCursor(e);
}, { passive: false });

el.brush.addEventListener('input', () => {
  if (tool && sizes[tool] !== undefined) sizes[tool] = +el.brush.value;
  updateLabels();
});

el.pEdge.addEventListener('input', updateLabels);   // действует на новые мазки «Уточнения края»; уже сделанные не меняются
[el.pBright, el.pChroma, el.pShadow, el.pSmooth, el.pText, el.pPick].forEach(s =>
  s.addEventListener('input', () => { updateLabels(); scheduleProcess(); }));
el.removeText.addEventListener('change', () => { updateLabels(); scheduleProcess(); });
for (const b of el.presetChips.children) {
  b.addEventListener('click', () => {
    const k = b.dataset.preset;
    if (presetsOn.has(k)) { presetsOn.delete(k); b.classList.remove('on'); } else { presetsOn.add(k); b.classList.add('on'); }
    scheduleProcess();
  });
}
el.btnPick.addEventListener('click', () => setTool('pick'));
el.clip.addEventListener('change', () => { if (base) scheduleProcess(); });
el.autoCrop.addEventListener('change', () => { if (fullCanvas) startFromFull(fullCanvas, loadNotes.filter(n => !/^Печать найдена|^Круглая печать авто|^Рабочая зона/.test(n))); });
el.autoRound.addEventListener('change', () => { if (origCanvas) resetEverything(); });

el.btnMask.addEventListener('click', () => setTool('mask'));
el.btnEraser.addEventListener('click', () => setTool('eraser'));
el.btnRestore.addEventListener('click', () => setTool('restore'));
el.btnTextMask.addEventListener('click', () => setTool('textmask'));
el.undo.addEventListener('click', undo);
el.round.addEventListener('click', () => makeRound(false));
el.dl.addEventListener('click', download);
async function addToFile() {
  if (!base) { setStatus('Сначала загрузите файл.'); return; }
  const r = await buildExport();
  if (!r) { setStatus('Результат пуст — нечего сохранять.'); return; }
  const SC = window.StampCache, res = await SC.save(SC.bytesToDataUrl(r.bytes), r.w, r.h);
  if (res.ok) setStatus('Печать сохранена в памяти как «Печать ' + res.slot + '» (занято ' + SC.count() + ' из ' + SC.MAX + ') на 30 минут. Откройте «Автозаполнение» — в каждом блоке её можно наложить на документ (40 мм).');
  else if (res.reason === 'full') setStatus('В памяти уже ' + SC.MAX + ' печати. Освободите место: нажмите ✕ у ненужной печати в списке «Печати в памяти» — и повторите.');
  else setStatus('Не удалось сохранить печать в память браузера (мало места). Скачайте PNG вручную.');
}
/* панель «Печати в памяти»: миниатюры, время до стирания, удаление по одной */
(function memPanel() {
  const SC = window.StampCache, list = $('memList'), cnt = $('memCount'), clr = $('memClear'); let sig = '';
  function draw() {
    const l = SC.list(); cnt.textContent = l.length + ' из ' + SC.MAX; clr.hidden = !l.length;
    const s = l.map(r => r.id + ':' + Math.ceil((r.expires - Date.now()) / 60000)).join(','); if (s === sig) return; sig = s;
    list.innerHTML = '';
    l.forEach(r => {
      const it = document.createElement('div'); it.className = 'mem-item';
      const im = document.createElement('img'); im.alt = 'Печать ' + r.slot; im.src = r.dataUrl;
      const tx = document.createElement('span'); tx.innerHTML = '<b>Печать ' + r.slot + '</b><small>ещё ≈ ' + Math.max(1, Math.ceil((r.expires - Date.now()) / 60000)) + ' мин</small>';
      const x = document.createElement('button'); x.type = 'button'; x.className = 'mem-x'; x.textContent = '✕'; x.setAttribute('data-tip', 'Стереть «Печать ' + r.slot + '» из памяти.');
      x.addEventListener('click', () => { SC.remove(r.id); setStatus('«Печать ' + r.slot + '» стёрта из памяти.'); });
      it.append(im, tx, x); list.append(it);
    });
  }
  clr.addEventListener('click', () => { SC.clear(); setStatus('Все печати стёрты из памяти.'); });
  SC.onChange(() => { sig = ''; draw(); }); setInterval(draw, 20000); draw();
})();
$('btnAddToFile').addEventListener('click', addToFile);
el.btnAuto.addEventListener('click', () => {
  if (!origCanvas || !src) { setStatus('Файл ещё не загружен.'); return; }
  edits = []; history = []; redrawOrig();
  autoAll(false);
});
el.reset.addEventListener('click', () => {
  if (!origCanvas) { setStatus('Файл ещё не загружен.'); return; }
  resetEverything(); setStatus('Всё сброшено к исходному изображению.');
});
el.bgSwitch.addEventListener('click', e => {
  const b = e.target.closest('button[data-bg]'); if (!b) return;
  for (const x of el.bgSwitch.children) x.classList.toggle('on', x === b);
  el.container.classList.remove('bg-checker', 'bg-white', 'bg-dark');
  el.container.classList.add('bg-' + b.dataset.bg);
});

document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); }
});

/* загрузка: кнопка */
el.file.addEventListener('change', () => {
  const f = el.file.files[0];
  handleFile(f).finally(() => { el.file.value = ''; });
});

/* загрузка: перетаскивание (кнопка, окно-заглушка и вся рабочая область) */
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => e.preventDefault());
function bindDrop(node, hoverNode) {
  node.addEventListener('dragover', e => { e.preventDefault(); hoverNode.classList.add('dragover'); });
  node.addEventListener('dragleave', () => hoverNode.classList.remove('dragover'));
  node.addEventListener('drop', e => {
    e.preventDefault(); e.stopPropagation();
    hoverNode.classList.remove('dragover');
    handleFile(e.dataTransfer.files && e.dataTransfer.files[0]);
  });
}
bindDrop(el.label, el.label);
bindDrop(el.drop, el.drop);
bindDrop(el.workspace, el.drop);
el.drop.addEventListener('click', () => el.file.click());

/* ---------- рабочая зона: ручной выбор области поиска печати на листе ---------- */
(function zoneUi() {
  const zp = $('zonePanel'), zc = $('zoneCanvas'), zr = $('zoneRect'), zw = $('zoneWrap'), zs = $('zoneStage');
  let rect = null, drag = null;
  const FW = () => fullCanvas.width, FH = () => fullCanvas.height, MIN = 40;
  function fit() {
    const aw = Math.max(200, zs.clientWidth - 32), ah = Math.max(160, zs.clientHeight - 32);
    const k = Math.min(aw / FW(), ah / FH(), 3);
    zc.style.width = Math.round(FW() * k) + 'px'; zc.style.height = Math.round(FH() * k) + 'px';
  }
  function draw() {
    zr.style.left = (rect.x / FW() * 100) + '%'; zr.style.top = (rect.y / FH() * 100) + '%';
    zr.style.width = (rect.w / FW() * 100) + '%'; zr.style.height = (rect.h / FH() * 100) + '%';
  }
  function limit() {
    rect.w = clamp(rect.w, MIN, FW()); rect.h = clamp(rect.h, MIN, FH());
    rect.x = clamp(rect.x, 0, FW() - rect.w); rect.y = clamp(rect.y, 0, FH() - rect.h);
  }
  function open() {
    if (!fullCanvas) { setStatus('Сначала загрузите файл.'); return; }
    zp.hidden = false;
    zc.width = FW(); zc.height = FH(); zc.getContext('2d').drawImage(fullCanvas, 0, 0);
    fit();
    const r = zoneRegion || lastRegion;
    rect = r ? { x: r.x, y: r.y, w: r.w, h: r.h } : { x: FW() * 0.2, y: FH() * 0.2, w: FW() * 0.6, h: FH() * 0.6 };
    limit(); draw();
  }
  const close = () => { zp.hidden = true; };
  const pt = e => { const r = zc.getBoundingClientRect(); return [(e.clientX - r.left) / r.width * FW(), (e.clientY - r.top) / r.height * FH()]; };
  zr.addEventListener('pointerdown', e => {
    e.preventDefault(); try { zr.setPointerCapture(e.pointerId); } catch (_) {}
    const [x, y] = pt(e); drag = { mode: e.target.dataset && e.target.dataset.h ? e.target.dataset.h : 'move', x, y, r: Object.assign({}, rect) };
  });
  zr.addEventListener('pointermove', e => {
    if (!drag) return;
    const [x, y] = pt(e), dx = x - drag.x, dy = y - drag.y, o = drag.r;
    if (drag.mode === 'move') { rect.x = o.x + dx; rect.y = o.y + dy; }
    else {
      let x0 = o.x, y0 = o.y, x1 = o.x + o.w, y1 = o.y + o.h;
      if (/w/.test(drag.mode)) x0 = clamp(o.x + dx, 0, x1 - MIN); if (/e/.test(drag.mode)) x1 = clamp(o.x + o.w + dx, x0 + MIN, FW());
      if (/n/.test(drag.mode)) y0 = clamp(o.y + dy, 0, y1 - MIN); if (/s/.test(drag.mode)) y1 = clamp(o.y + o.h + dy, y0 + MIN, FH());
      rect = { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
    }
    limit(); draw();
  });
  const end = () => { drag = null; };
  zr.addEventListener('pointerup', end); zr.addEventListener('pointercancel', end);
  // клик по листу вне рамки — переносит рамку в это место
  zc.addEventListener('pointerdown', e => { const [x, y] = pt(e); rect.x = x - rect.w / 2; rect.y = y - rect.h / 2; limit(); draw(); });
  // колесо мыши — масштаб рамки вокруг её центра (вверх — меньше, вниз — больше)
  zs.addEventListener('wheel', e => {
    if (zp.hidden) return; e.preventDefault();
    const f = e.deltaY < 0 ? 1 / 1.1 : 1.1, cx = rect.x + rect.w / 2, cy = rect.y + rect.h / 2;
    rect.w *= f; rect.h *= f; rect.w = clamp(rect.w, MIN, FW()); rect.h = clamp(rect.h, MIN, FH());
    rect.x = cx - rect.w / 2; rect.y = cy - rect.h / 2; limit(); draw();
  }, { passive: false });
  $('btnZone').addEventListener('click', () => (zp.hidden ? open() : close()));
  $('zoneClose').addEventListener('click', close);
  $('zoneApply').addEventListener('click', () => {
    zoneRegion = { x: rect.x, y: rect.y, w: rect.w, h: rect.h }; close();
    startFromFull(fullCanvas, loadNotes.filter(n => !/^Печать найдена|^Круглая печать авто|^Рабочая зона/.test(n)));
    setStatus('Рабочая зона применена: печать ищется внутри выбранной рамки.');
  });
  $('zoneAuto').addEventListener('click', () => {
    zoneRegion = null; close();
    startFromFull(fullCanvas, loadNotes.filter(n => !/^Печать найдена|^Круглая печать авто|^Рабочая зона/.test(n)));
  });
  window.addEventListener('resize', () => { if (!zp.hidden && fullCanvas) { fit(); draw(); } });
})();

window.addEventListener('resize', layout);
updateLabels();
setStatus('Готов к работе. Загрузите изображение или PDF.');

/* точка входа для автоматической проверки */
window.StampStudio = {
  loadCanvas(c) { fileLabel = 'test'; startFromFull(c, []); },
  handleFile,
  process, makeRound, buildExport,
  info() {
    return {
      W, H, notes: loadNotes.concat(procNotes),
      pdf: { page: pdfPage, count: pdfCount }, removeText: el.removeText.checked, smooth: +el.pSmooth.value,
      ana: ana && { fit: ana.fit,
        protos: ana.protos.map(p => ({ v: p.v.map(x => +x.toFixed(3)), n: p.n })), stampIdx: ana.stampIdx }
    };
  },
  alphaStats() { let n = 0; for (let i = 0; i < cur.length; i++) if (cur[i] > 128) n++; return { n, W, H }; },
  alphaIn(x0, y0, x1, y1) { let n = 0; for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) if (cur[y * W + x] > 128) n++; return n; }
};
})();
