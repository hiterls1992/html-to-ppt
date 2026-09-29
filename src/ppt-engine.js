
/* ============================================================
   HTML 16:9 高保真可编辑 PPT 导出引擎 V1.0
   架构（模块化分区，单文件内嵌，离线可用）：
   [1] CoordinateTransform  统一坐标转换（唯一入口，禁止旁路计算）
   [2] StyleParser          颜色/阴影/字体解析（getComputedStyle 计算值）
   [3] FallbackRenderer     元素级高清 PNG 兜底（html2canvas，3840×2160）
   [4] Classifier/Converter DOM 递归分类 → 原生 PPT 对象
   [5] Exporter             三模式（智能混合/高保真/最大可编辑）+ 校验
   设计基准 1280×720（16:9），与 1920×1080 同比例，换算公式一致：
   pptX = htmlX / 1280 × 13.333
   ============================================================ */
(function () {
  'use strict';

  /* ============ [1] CoordinateTransform ============ */
  var DESIGN = { W: 1280, H: 720 };
  var SLIDE = { W: 13.333, H: 7.5 };
  function px2in(px) { return Math.round(px / 96 * 1000) / 1000; }
  /* 依当前屏幕缩放与 .slide 视口矩形，返回元素 → 英寸坐标的唯一转换器 */
  function makeCT(scale, R) {
    return function (r) {
      return {
        x: Math.round((r.left - R.left) / scale / DESIGN.W * SLIDE.W * 1000) / 1000,
        y: Math.round((r.top - R.top) / scale / DESIGN.H * SLIDE.H * 1000) / 1000,
        w: Math.round(r.width / scale / DESIGN.W * SLIDE.W * 1000) / 1000,
        h: Math.round(r.height / scale / DESIGN.H * SLIDE.H * 1000) / 1000
      };
    };
  }

  /* ============ [2] StyleParser ============ */
  function parseColor(c) {
    if (!c) return null;
    c = String(c).trim();
    var m = c.match(/^rgba?\(([^)]+)\)$/i);
    if (m) {
      var p = m[1].split(',').map(function (x) { return parseFloat(x); });
      var a = p.length > 3 ? p[3] : 1;
      if (a <= 0.02) return null;
      return { hex: rgbHex(p[0], p[1], p[2]), a: a };
    }
    if (c.charAt(0) === '#') {
      var h = c.substr(1);
      if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
      if (h.length >= 6) return { hex: h.substr(0, 6).toUpperCase(), a: 1 };
    }
    return null;
  }
  function rgbHex(r, g, b) {
    function h(v) { return ('0' + Math.max(0, Math.min(255, Math.round(v))).toString(16)).slice(-2); }
    return (h(r) + h(g) + h(b)).toUpperCase();
  }
  function cssVar(el, name) {
    var v = getComputedStyle(el).getPropertyValue(name);
    return v ? v.trim() : '';
  }
  /* 简单 box-shadow → PPT Shadow；多重复阴影取第一条（复杂阴影走 Fallback） */
  function parseBoxShadow(sh) {
    if (!sh || sh === 'none') return null;
    var m = sh.match(/rgba?\(([^)]+)\)\s+([-0-9.]+)px\s+([-0-9.]+)px\s+([-0-9.]+)px/);
    if (!m) return null;
    var p = m[1].split(',').map(function (x) { return parseFloat(x); });
    var a = p.length > 3 ? p[3] : 1;
    if (a <= 0.02) return null;
    var x = parseFloat(m[2]), y = parseFloat(m[3]), blur = parseFloat(m[4]);
    var ang = Math.atan2(y, x) * 180 / Math.PI;
    return { type: 'outer', angle: Math.round(ang), blur: Math.max(1, blur * 0.75),
             color: rgbHex(p[0], p[1], p[2]), offset: Math.max(0.5, Math.sqrt(x * x + y * y) * 0.75),
             opacity: a };
  }
  function parseTextShadow(sh) {
    if (!sh || sh === 'none') return null;
    var m = sh.match(/rgba?\(([^)]+)\)\s+([-0-9.]+)px\s+([-0-9.]+)px\s+([-0-9.]+)px/);
    if (!m) return null;
    var p = m[1].split(',').map(function (x) { return parseFloat(x); });
    var blur = parseFloat(m[4]);
    return { type: 'outer', angle: 90, blur: Math.max(1, blur * 0.75),
             color: rgbHex(p[0], p[1], p[2]), offset: 1,
             opacity: p.length > 3 ? p[3] : 1 };
  }
  /* 字体保持：取引号主字体；白名单外降级为微软雅黑并计数 */
  var SAFE_FONTS = { 'microsoft yahei': 'Microsoft YaHei', consolas: 'Consolas',
                     arial: 'Arial', helvetica: 'Helvetica', 'times new roman': 'Times New Roman' };
  function fontOf(cs, stats) {
    var m = cs.fontFamily.match(/"([^"]+)"/);
    var f = (m ? m[1] : cs.fontFamily.split(',')[0] || '').replace(/"/g, '').trim();
    var hit = SAFE_FONTS[f.toLowerCase()];
    if (!hit) { if (stats) stats.fontDegrade++; return 'Microsoft YaHei'; }
    return hit;
  }

  /* ============ [3] FallbackRenderer（元素级 PNG，scale=3 → 3840×2160 级） ============ */
  function elementPNG(el) {
    if (!window.html2canvas) return Promise.reject(new Error('html2canvas 不可用'));
    return window.html2canvas(el, { scale: 3, backgroundColor: null, logging: false })
      .then(function (cv) { return cv.toDataURL('image/png'); });
  }
  function slidePNG(sl) {
    if (!window.html2canvas) return Promise.reject(new Error('html2canvas 不可用'));
    return window.html2canvas(sl, { scale: 3, backgroundColor: null, logging: false,
      onclone: function (doc) { var s = doc.querySelector('.slide'); if (s) s.style.transform = 'none'; } })
      .then(function (cv) { return cv.toDataURL('image/png'); });
  }

  /* ============ [4] Classifier / Converter（DOM 递归分类） ============ */
  function hasCls(el, name) {
    return (' ' + (el.className || '') + ' ').indexOf(' ' + name + ' ') >= 0;
  }
  function isComplex(el) {
    var cs = getComputedStyle(el);
    return (cs.backdropFilter && cs.backdropFilter !== 'none') ||
           (cs.filter && cs.filter !== 'none') ||
           (cs.clipPath && cs.clipPath !== 'none');
  }
  function hasDirectText(el) {
    for (var i = 0; i < el.childNodes.length; i++) {
      var n = el.childNodes[i];
      if (n.nodeType === 3 && n.nodeValue && n.nodeValue.trim()) return true;
    }
    return false;
  }
  function hasVisibleBox(el) {
    var cs = getComputedStyle(el);
    var bg = parseColor(cs.backgroundColor);
    var bw = parseFloat(cs.borderTopWidth) || 0;
    return !!(bg || bw > 0);
  }
  function isPill(el) { return hasCls(el, 'tag') || hasCls(el, 'org') || hasCls(el, 'chip'); }

  function clone(o) { var r = {}; for (var k in o) r[k] = o[k]; return r; }
  function runs(el, base) {
    var arr = [];
    el.childNodes.forEach(function (n) {
      if (n.nodeType === 3) {
        if (n.nodeValue) arr.push({ text: n.nodeValue, options: clone(base) });
      } else if (n.nodeType === 1) {
        if (n.tagName === 'BR') { arr.push({ text: '\n', options: clone(base) }); return; }
        var cs = getComputedStyle(n);
        var o = clone(base);
        var c = parseColor(cs.color); if (c) o.color = c.hex;
        o.bold = parseInt(cs.fontWeight, 10) >= 600;
        o.fontSize = px2pt(parseFloat(cs.fontSize));
        if (cs.fontStyle === 'italic') o.italic = true;
        if ((cs.textDecorationLine || '').indexOf('underline') >= 0) o.underline = true;
        arr.push({ text: n.textContent, options: o });
      }
    });
    if (!arr.length) arr.push({ text: el.textContent, options: clone(base) });
    return arr;
  }
  function px2pt(px) { return Math.round(px * 0.75 * 10) / 10; }

  /* 递归分类转换：每个元素 try/catch，单元素失败不中断导出（Fallback 原则） */
  function walk(ctx, el) {
    try { walkInner(ctx, el); }
    catch (e) {
      ctx.stats.fail++;
      try { console.error('[PPTX 导出] 元素转换失败，已回退：', el.tagName, el.className || '', (e && e.message) || e); } catch (_e) {}
      if (ctx.mode !== 'maxedit' && window.html2canvas) {
        /* Level 2→3：元素级 PNG 兜底 */
        ctx.tasks.push(
          elementPNG(el).then(function (url) {
            var b = ctx.ct(el.getBoundingClientRect());
            ctx.slide.addImage({ data: url, x: b.x, y: b.y, w: b.w, h: b.h });
            ctx.stats.images++; ctx.stats.fallbacks++;
          }).catch(function () { ctx.stats.fail++; })
        );
      }
    }
  }
  function walkInner(ctx, el) {
    var cs = getComputedStyle(el);

    /* 装饰元素与复杂元素：使用 Phase1 预捕获的透明 PNG，按 DOM 顺序原位插入（z 序正确） */
    var isDeco = hasCls(el, 'glow') || hasCls(el, 'glow2');
    if (isDeco || isComplex(el)) {
      if (ctx.mode === 'maxedit') return;
      var purl = ctx.pngMap.get(el);
      if (purl) {
        var pb = ctx.ct(el.getBoundingClientRect());
        ctx.slide.addImage({ data: purl, x: pb.x, y: pb.y, w: pb.w, h: pb.h });
        ctx.stats.images++; ctx.stats.fallbacks++;
      } else if (!isDeco) {
        /* 无位图（关闭 Fallback 或捕获失败）→ 高亮色实底近似，不丢失元素 */
        emitBox(ctx, el);
      }
      return;
    }
    /* 表格 → 原生表格 */
    if (el.tagName === 'TABLE') { emitTable(ctx, el); return; }
    /* 分隔线 → accent 细条 */
    if (hasCls(el, 'divider') || hasCls(el, 'cline')) { emitLine(ctx, el); return; }
    /* 胶囊标签：文字 + 圆角矩形一体 */
    if (isPill(el)) { emitPill(ctx, el); return; }
    /* 图片占位框：虚线圆角矩形（子元素继续递归出图标与文字） */
    if (hasCls(el, 'ph')) { emitPh(ctx, el); }
    /* 占位小图标：圆(日) + 三角(山) */
    if (hasCls(el, 'ic')) { emitIcon(ctx, el); return; }

    /* 自身背景/边框 → 原生 Shape（先画底，子文本在其上层） */
    var boxed = false;
    if (el !== ctx.root && hasVisibleBox(el)) { emitBox(ctx, el); boxed = true; }
    /* 自身直接文本 → 原生文本框（富文本 runs 已并入高亮子元素 span/b/small，
       输出后不再递归子元素，防止高亮文字被重复绘制） */
    if (hasDirectText(el)) { emitText(ctx, el, boxed); return; }
    /* 递归子元素 */
    el.childNodes.forEach(function (n) { if (n.nodeType === 1) walk(ctx, n); });
  }

  function emitBox(ctx, el) {
    var b = ctx.ct(el.getBoundingClientRect());
    if (b.w < 0.03 || b.h < 0.03) return;
    var cs = getComputedStyle(el);
    var fill = parseColor(cs.backgroundColor);
    var bc = parseColor(cs.borderTopColor);
    var bw = parseFloat(cs.borderTopWidth) || 0;
    var opts = { x: b.x, y: b.y, w: b.w, h: b.h, rectRadius: Math.min(0.15, px2in(parseFloat(cs.borderTopLeftRadius) || 0)) };
    if (fill) opts.fill = { color: fill.hex, transparency: fill.a < 0.98 ? Math.round((1 - fill.a) * 100) : 0 };
    if (bc && bw > 0) opts.line = { color: bc.hex, width: Math.max(0.5, bw * 0.75) };
    var sh = parseBoxShadow(cs.boxShadow);
    if (sh) opts.shadow = sh;
    ctx.slide.addShape('roundRect', opts);
    ctx.stats.shapes++;
  }
  function emitText(ctx, el, inBox) {
    var b = ctx.ct(el.getBoundingClientRect());
    if (b.w < 0.05 || b.h < 0.02) return;
    var cs = getComputedStyle(el);
    var col = parseColor(cs.color);
    var fs0 = parseFloat(cs.fontSize);
    /* 单行判定：元素设计高度 < 1.8 倍字号 → 禁止 PPT 内换行（防字体度量差异导致的意外折行） */
    var singleLine = b.h * 96 < fs0 * 1.8;
    var o = {
      x: b.x, y: b.y, w: b.w, h: b.h,
      fontSize: px2pt(fs0),
      fontFace: fontOf(cs, ctx.stats),
      color: (col && col.hex) || '333333',
      bold: parseInt(cs.fontWeight, 10) >= 600,
      italic: cs.fontStyle === 'italic',
      underline: (cs.textDecorationLine || '').indexOf('underline') >= 0,
      align: cs.textAlign === 'center' ? 'center' : (cs.textAlign === 'right' ? 'right' : 'left'),
      valign: 'top', margin: 0, fit: 'shrink',
      wrap: !singleLine
    };
    /* 行距定稿（用户 2026-09-29）：多行文本统一 1.5 倍——PPT 对中文字体的
       multiple 语义偏松，跟随浏览器计算值（1.75）会明显过疏 */
    if (!singleLine) o.lineSpacingMultiple = 1.5;
    var ls = parseFloat(cs.letterSpacing);
    if (ls > 0) o.charSpacing = px2pt(ls);
    if (col && col.a < 0.98) o.transparency = Math.round((1 - col.a) * 100);
    var ts = parseTextShadow(cs.textShadow);
    if (ts) o.shadow = ts;
    if (inBox) { o.margin = 2; }   /* 在色块内的文字留少量内边距防贴边 */
    ctx.textOps.push(function () {
      ctx.slide.addText(runs(el, o), o);
      ctx.stats.texts++;
    });
  }
  function emitPill(ctx, el) {
    var b = ctx.ct(el.getBoundingClientRect());
    if (b.w < 0.05 || b.h < 0.05) return;
    var cs = getComputedStyle(el);
    var fill = parseColor(cs.backgroundColor) || parseColor(cssVar(el, '--bg-highlight'));
    var line = parseColor(cssVar(el, '--border'));
    var tc = parseColor(cs.color);
    ctx.slide.addText(el.textContent.trim(), {
      shape: 'roundRect', x: b.x, y: b.y, w: b.w, h: b.h, rectRadius: b.h / 2,
      fill: (fill) ? { color: fill.hex, transparency: fill.a < 0.98 ? Math.round((1 - fill.a) * 100) : 0 }
                   : { color: 'EEEEEE' },
      line: (line && line.hex) ? { color: line.hex, width: 0.75 } : undefined,
      color: (tc && tc.hex) || '333333',
      fontSize: px2pt(parseFloat(cs.fontSize)), fontFace: fontOf(cs, ctx.stats),
      bold: parseInt(cs.fontWeight, 10) >= 600,
      align: 'center', valign: 'middle', margin: 0,
      charSpacing: parseFloat(cs.letterSpacing) > 0 ? px2pt(parseFloat(cs.letterSpacing)) : undefined
    });
    ctx.stats.shapes++; ctx.stats.texts++;
  }
  function emitLine(ctx, el) {
    var b = ctx.ct(el.getBoundingClientRect());
    if (b.w < 0.05) return;
    var ac = parseColor(cssVar(el, '--accent'));
    ctx.slide.addShape('roundRect', { x: b.x, y: b.y, w: b.w, h: Math.max(b.h, 0.02),
      rectRadius: 0.02, fill: { color: (ac && ac.hex) || '888888' } });
    ctx.stats.shapes++;
  }
  function emitPh(ctx, el) {
    var b = ctx.ct(el.getBoundingClientRect());
    if (b.w < 0.05 || b.h < 0.05) return;
    var hl = parseColor(cssVar(el, '--bg-highlight'));
    var st = parseColor(cssVar(el, '--border-strong'));
    ctx.slide.addShape('roundRect', { x: b.x, y: b.y, w: b.w, h: b.h, rectRadius: 0.08,
      fill: { color: (hl && hl.hex) || 'F2F2F2' },
      line: (st && st.hex) ? { color: st.hex, width: 0.75, dashType: 'dash' } : undefined });
    ctx.stats.shapes++;
  }
  function emitIcon(ctx, el) {
    var b = ctx.ct(el.getBoundingClientRect());
    var st = parseColor(cssVar(el, '--border-strong'));
    var c = (st && st.hex) || '999999';
    ctx.slide.addShape('ellipse', { x: b.x + b.w * 0.55, y: b.y + b.h * 0.10,
      w: b.w * 0.18, h: b.w * 0.18, fill: { color: c } });
    ctx.slide.addShape('triangle', { x: b.x + b.w * 0.04, y: b.y + b.h * 0.42,
      w: b.w * 0.55, h: b.h * 0.55, fill: { color: c } });
    ctx.slide.addShape('triangle', { x: b.x + b.w * 0.36, y: b.y + b.h * 0.53,
      w: b.w * 0.42, h: b.h * 0.42, fill: { color: c, transparency: 45 } });
    ctx.stats.shapes += 3;
  }
  function emitTable(ctx, tbl) {
    var b = ctx.ct(tbl.getBoundingClientRect());
    var rows = [];
    tbl.querySelectorAll('tr').forEach(function (tr) {
      var cells = [];
      tr.querySelectorAll('th,td').forEach(function (c) {
        var cs = getComputedStyle(c);
        var col = parseColor(cs.color);
        var bg = parseColor(cs.backgroundColor);
        cells.push({ text: c.textContent.trim(), options: {
          bold: c.tagName === 'TH' || c.classList.contains('hl'),
          color: (col && col.hex) || '333333',
          fill: { color: (bg && bg.hex) || 'FFFFFF' },
          align: 'center', valign: 'middle',
          fontFace: fontOf(cs, ctx.stats), fontSize: px2pt(parseFloat(cs.fontSize))
        }});
      });
      rows.push(cells);
    });
    var firstRow = tbl.querySelector('tr');
    var colW = [];
    firstRow.querySelectorAll('th,td').forEach(function (c) {
      colW.push(ctx.ct(c.getBoundingClientRect()).w);
    });
    var bch = parseColor(getComputedStyle(tbl.querySelector('td')).borderTopColor);
    ctx.slide.addTable(rows, { x: b.x, y: b.y, w: b.w, colW: colW,
      border: { pt: 0.75, color: (bch && bch.hex) || '999999' },
      fontFace: 'Microsoft YaHei', valign: 'middle' });
    ctx.stats.tables++;
  }

  /* ============ [5] Exporter ============ */
  function buildSlide(pptx, fr, mode, stats, tasks, pngMap) {
    var sl = fr.querySelector('.slide');
    if (!sl) throw new Error('页面缺少 .slide');
    var scale = parseFloat(fr.style.width) / 1280 || 1;
    var R = sl.getBoundingClientRect();
    var slide = pptx.addSlide();
    var panel = parseColor(cssVar(sl, '--bg-panel'));
    slide.background = { color: (panel && panel.hex) || 'FFFFFF' };
    var ctx = { pptx: pptx, slide: slide, mode: mode, stats: stats, tasks: tasks,
                root: sl, ct: makeCT(scale, R), textOps: [], pngMap: pngMap || new Map() };
    Array.prototype.forEach.call(sl.childNodes, function (n) {
      if (n.nodeType === 1) walk(ctx, n);
    });
    /* 第二遍输出文本：保证溢出文字不被后续兄弟色块遮挡（修复行距导致的文字被盖） */
    ctx.textOps.forEach(function (op) { op(); });
    return slide;
  }

  /* 写文件 + 导出后校验摘要（页/文本/形状/图片/表格/回退/失败/字体降级） */
  function finishPptx(pptx, stats) {
    var hint = document.getElementById('hint');
    hint.textContent = '正在生成 PPT 文件…';
    var btEl = document.querySelector('.bar .bt');
    var bt = btEl ? btEl.textContent : '演示文稿';
    var deck = bt.split('（')[0].split('·')[0].trim() || '演示文稿';
    var themeName = (typeof THEME_NAMES !== 'undefined' && typeof currentTheme !== 'undefined')
      ? THEME_NAMES[currentTheme] : '';
    var fname = deck + (themeName ? '-' + themeName : '') + '.pptx';
    return pptx.writeFile({ fileName: fname }).then(function () {
      var msg = 'PPT 生成成功（' + fname + '）· 页 ' + stats.slides + ' · 文本 ' + stats.texts +
                ' · 形状 ' + stats.shapes + ' · 图片 ' + stats.images + ' · 表格 ' + stats.tables;
      if (stats.fallbacks) msg += ' · 位图回退 ' + stats.fallbacks;
      if (stats.fail) msg += ' · 失败 ' + stats.fail;
      if (stats.fontDegrade) msg += ' · 字体降级 ' + stats.fontDegrade;
      hint.textContent = msg;
      return stats;
    });
  }
  function run(mode, opts) {
    var hint = document.getElementById('hint');
    var pptx = new window.PptxGenJS();
    pptx.defineLayout({ name: 'W169E', width: 13.333, height: 7.5 });
    pptx.layout = 'W169E';
    pptx.author = 'AI Showcase';
    var stats = { slides: 0, texts: 0, shapes: 0, images: 0, tables: 0,
                  fallbacks: 0, fail: 0, fontDegrade: 0 };
    var tasks = [];
    hint.textContent = '正在分析页面…';
    /* Phase 1：预捕获装饰/复杂元素透明 PNG（先分析 → 再转换 → 最后生成） */
    var pngMap = new Map();
    var pre = [];
    if (mode !== 'maxedit' && window.html2canvas && (!opts || opts.autoFb !== false)) {
      Array.prototype.forEach.call(document.querySelectorAll('.slide .glow, .slide .glow2'), function (el) {
        pre.push(elementPNG(el).then(function (url) { pngMap.set(el, url); }).catch(function () {}));
      });
      Array.prototype.forEach.call(document.querySelectorAll('.slide *'), function (el) {
        if (pngMap.has(el)) return;
        try { if (isComplex(el)) pre.push(elementPNG(el).then(function (url) { pngMap.set(el, url); }).catch(function () {})); } catch (_e) {}
      });
    }
    /* 高保真模式：整页高清位图（视觉 100%） */
    if (mode === 'hifi') {
      return Promise.all(pre).then(function () {
        Array.prototype.forEach.call(document.querySelectorAll('.frame'), function (fr) {
          var sl = fr.querySelector('.slide');
          if (!sl) return;
          stats.slides++;
          tasks.push(slidePNG(sl).then(function (url) {
            var slide = pptx.addSlide();
            var panel = parseColor(cssVar(sl, '--bg-panel'));
            slide.background = { color: (panel && panel.hex) || 'FFFFFF' };
            slide.addImage({ data: url, x: 0, y: 0, w: SLIDE.W, h: SLIDE.H });
            stats.images++; stats.fallbacks++;
          }).catch(function () { stats.fail++; }));
        });
        return Promise.all(tasks);
      }).then(function () { return finishPptx(pptx, stats); });
    }
    /* 等待预捕获完成后，再进入原生转换（保证 pngMap 就绪） */
    return Promise.all(pre).then(function () {
      hint.textContent = '正在转换文字与图形…';
      var frames = document.querySelectorAll('.frame');
      frames.forEach(function (fr) {
        try {
          buildSlide(pptx, fr, mode, stats, tasks, pngMap);
          stats.slides++;
        } catch (e) {
          /* Level 4：整页截图兜底 */
          stats.fail++;
          if (mode !== 'maxedit' && window.html2canvas) {
            tasks.push(
              slidePNG(fr.querySelector('.slide')).then(function (url) {
                var slide = pptx.addSlide();
                slide.addImage({ data: url, x: 0, y: 0, w: SLIDE.W, h: SLIDE.H });
                stats.images++; stats.fallbacks++;
              }).catch(function () {})
            );
          }
        }
      });
      hint.textContent = '正在处理装饰与复杂元素…';
      return Promise.all(tasks).then(function () {
        return finishPptx(pptx, stats);
      });
    });
  }

  /* 写文件 + 导出后校验摘要（页/文本/形状/图片/表格/回退/失败/字体降级） */
  function finishPptx(pptx, stats) {
    var hint = document.getElementById('hint');
    hint.textContent = '正在生成 PPT 文件…';
    var btEl = document.querySelector('.bar .bt');
    var bt = btEl ? btEl.textContent : '演示文稿';
    var deck = bt.split('（')[0].split('·')[0].trim() || '演示文稿';
    var themeName = (typeof THEME_NAMES !== 'undefined' && typeof currentTheme !== 'undefined')
      ? THEME_NAMES[currentTheme] : '';
    var fname = deck + (themeName ? '-' + themeName : '') + '.pptx';
    return pptx.writeFile({ fileName: fname }).then(function () {
      var msg = 'PPT 生成成功（' + fname + '）· 页 ' + stats.slides + ' · 文本 ' + stats.texts +
                ' · 形状 ' + stats.shapes + ' · 图片 ' + stats.images + ' · 表格 ' + stats.tables;
      if (stats.fallbacks) msg += ' · 位图回退 ' + stats.fallbacks;
      if (stats.fail) msg += ' · 失败 ' + stats.fail;
      if (stats.fontDegrade) msg += ' · 字体降级 ' + stats.fontDegrade;
      hint.textContent = msg;
      return stats;
    });
  }

  /* ============ UI（与页面风格一致的下拉面板） ============ */
  function injectUI() {
    var bar = document.querySelector('.bar');
    if (!bar || document.getElementById('btnPpt')) return;
    var style = document.createElement('style');
    style.textContent =
      '.pptx-panel{position:fixed;top:60px;right:20px;z-index:120;width:250px;' +
      'background:var(--bg-panel);border:1px solid var(--border);border-radius:10px;' +
      'box-shadow:0 10px 30px var(--shadow-strong);padding:14px 16px;display:none;' +
      'font-size:13px;color:var(--text-primary);}' +
      '.pptx-panel.open{display:block;}' +
      '.pptx-panel h4{font-size:14px;color:var(--text-highlight);margin-bottom:8px;}' +
      '.pptx-panel label{display:flex;align-items:center;gap:8px;margin:6px 0;cursor:pointer;}' +
      '.pptx-panel .go{width:100%;margin-top:10px;padding:7px 0;border-radius:16px;' +
      'border:none;background:var(--accent);color:var(--accent-contrast);cursor:pointer;' +
      'font-size:13px;font-family:inherit;}' +
      '.pptx-panel .go:disabled{opacity:.55;cursor:wait;}' +
      '.pptx-panel .sec{font-size:12px;color:var(--text-muted);margin:8px 0 2px;}';
    document.head.appendChild(style);

    var btn = document.createElement('button');
    btn.id = 'btnPpt';
    btn.textContent = '导出 PPT ▼';
    var btnExport = document.getElementById('btnExport');
    if (btnExport) bar.insertBefore(btn, btnExport); else bar.appendChild(btn);

    var panel = document.createElement('div');
    panel.className = 'pptx-panel';
    panel.id = 'pptxPanel';
    panel.innerHTML =
      '<h4>导出 PowerPoint</h4>' +
      '<div class="sec">导出模式</div>' +
      '<label><input type="radio" name="pptx-mode" value="hybrid" checked>智能混合（推荐）</label>' +
      '<label><input type="radio" name="pptx-mode" value="hifi">高保真（整页高清位图）</label>' +
      '<label><input type="radio" name="pptx-mode" value="maxedit">最大可编辑（无装饰位图）</label>' +
      '<div class="sec">选项</div>' +
      '<label><input type="checkbox" checked disabled>保持 16:9（13.33×7.5 in）</label>' +
      '<label><input type="checkbox" id="pptx-font" checked>保持网页字体</label>' +
      '<label><input type="checkbox" id="pptx-fallback" checked>自动 Fallback（复杂元素转高清位图）</label>' +
      '<button class="go" id="pptx-go">开始导出</button>';
    bar.appendChild(panel);

    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      panel.classList.toggle('open');
    });
    panel.addEventListener('click', function (e) { e.stopPropagation(); });
    document.addEventListener('click', function () { panel.classList.remove('open'); });

    panel.querySelector('#pptx-go').addEventListener('click', function () {
      var go = this;
      var hint = document.getElementById('hint');
      var mode = panel.querySelector('input[name="pptx-mode"]:checked').value;
      var keepFont = panel.querySelector('#pptx-font').checked;
      var autoFb = panel.querySelector('#pptx-fallback').checked;
      go.disabled = true;
      Promise.resolve().then(function () { return run(mode, { keepFont: keepFont, autoFb: autoFb }); })
        .catch(function (e) {
          console.error(e);
          hint.textContent = 'PPT 导出失败：' + ((e && e.message) || e);
        })
        .then(function () { go.disabled = false; });
    });
  }

  function onReady(fn) {
    if (document.readyState !== 'loading') fn();
    else document.addEventListener('DOMContentLoaded', fn);
  }
  onReady(function () {
    if (!window.PptxGenJS) return;
    injectUI();
    window.__pptxRun = run;   /* 测试钩子：vm 集成测试与控制台调用入口 */
  });
})();


