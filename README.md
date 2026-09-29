# html-to-ppt — HTML 16:9 高保真可编辑 PPT 导出引擎 V1.0

将 16:9 的 HTML 成果展示页 / 幻灯片在浏览器内一键导出为**标准 `.pptx` 文件**。
不是截图拼图，而是混合渲染引擎：**简单元素转 PPT 原生可编辑对象，复杂元素逐级 Fallback**。

> 核心理念：**HTML 是设计源文件，PPT 是输出格式。**
> 不强迫 HTML 适应 PPT，而是建立一层智能转换引擎。

## 功能特性

- **三种导出模式**
  - 智能混合（默认）：文本/形状/表格 → PPT 原生对象；装饰与复杂元素 → 元素级 4K 透明 PNG
  - 高保真：整页 scale=3 高清位图，视觉 100% 一致
  - 最大可编辑：无任何位图，纯原生对象
- **统一坐标转换**：设计基准 1280×720（16:9）→ 13.333×7.5 in，全项目唯一换算入口
- **样式保真**：基于 `getComputedStyle` 计算值转换——字号/颜色/粗细/斜体/下划线/对齐/行高/字距/圆角/边框/填充透明度/简单阴影均跟随
- **原生表格**：表头底色、斑马纹、单元格高亮原样映射，PPT 内可直接编辑
- **四级 Fallback**：原生对象 → 位图 → 整页截图，单元素失败不中断导出
- **导出后校验**：页数/文本/形状/图片/表格/回退/失败/字体降级 八项摘要
- 依赖 PptxGenJS 3.12.0（MIT）与 html2canvas 1.4.1（MIT），均离线可用

## 快速开始

```html
<body>
  <!-- 你的 16:9 页面（设计基准 1280×720，每页 .frame > .slide） -->
  <div class="frame" data-name="01-页面名">
    <section class="slide">
      <h1>标题<span>强调词</span></h1>
      <p class="desc">正文……</p>
    </section>
  </div>

  <!-- 引入两个脚本即可 -->
  <script src="vendor/pptxgen.bundle.js"></script>
  <script src="src/ppt-engine.js"></script>
</body>
```

点击顶栏「导出 PPT ▼」→ 选择模式 → 开始导出，生成 `页面名-主题.pptx`。

也可控制台调用：`window.__pptxRun('hybrid' | 'hifi' | 'maxedit')`。

完整示例见 [demo/](demo/)（双击 `demo/ppt-template.html` 即可体验，含 12 页样板：封面/目录/章节/内容/图片/表格/致谢）。

## 目录结构

```
html-to-ppt/
├── src/ppt-engine.js          # 导出引擎（CoordinateTransform / StyleParser /
│                              #   FallbackRenderer / Classifier / Exporter 五模块）
├── vendor/pptxgen.bundle.js   # PptxGenJS 3.12.0（MIT，含 JSZip）
├── demo/ppt-template.html     # 演示：12 页 PPT 模版（封面/目录/章节/图片/表格/致谢）
├── demo/collection.html       # 演示：6 页成果合集（多版式内容页）
├── LICENSE
└── README.md
```

## 转换策略（智能混合模式）

| HTML 元素 | PPT 对象 |
|---|---|
| h1~h4 / p / span / b / small（含富文本混排） | 原生文本框（runs 保留各子元素字号颜色粗细） |
| 卡片 / 面板 / 指标块 / 标签胶囊 | 圆角矩形（填充含透明度、边框、圆角、简单阴影） |
| table | 原生表格（表头底色、斑马纹、单元格高亮） |
| 分隔线 / 细线 | 色条 Shape |
| 径向光晕等装饰、backdrop-filter / filter / clip-path 元素 | 元素级 3840×2160 透明 PNG（Phase 1 预捕获，z 序正确） |
| 单元素转换失败 | 逐级 Fallback，永不中断导出 |

单行判定：元素高度 < 1.8×字号 → `wrap:false`，避免 PPT 字体度量差异导致的意外折行。

## 已知限制

- 网页辉光 / 径向渐变装饰以位图形式呈现（高明度半透明，观感一致但不可编辑）
- 文本行距在 PPT 中统一为 1.5 倍（PPT 对中文字体的倍数语义与浏览器不一致）
- 版式为坐标近似还原（±像素级），非像素级复制
- 图片占位框导出后需在 PPT 中手动替换为真实图片

## 验证环境

- Chrome / Edge（Chromium 内核，需支持 `getComputedStyle`、html2canvas、ES5）
- 导出的 .pptx 在 Microsoft PowerPoint 与 WPS Office 中验证可正常打开、文字与表格可编辑

## License

MIT（引擎部分）。依赖库：PptxGenJS MIT © Brent Ely；html2canvas MIT © Niklas von Hertzen。
