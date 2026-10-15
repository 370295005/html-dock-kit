import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const outputDirectory = path.dirname(fileURLToPath(import.meta.url));
await mkdir(outputDirectory, { recursive: true });

// Reuse the committed diagnostic fixture so generation works with plain Node
// and produces the exact same image bytes on every machine.
const png = await readFile(path.join(outputDirectory, 'diagnostic.png'));
const base64 = png.toString('base64');
const dataUrl = `data:image/png;base64,${base64}`;
const whitespaceDataUrl = `data:image/png;base64,${base64.replace(/.{72}/g, '$&\n \t')}`;
const parameterDataUrl = `data:image/png;name=diagnostic.png;base64,${base64}`;
const imageSha256 = createHash('sha256').update(png).digest('hex');

const escapeHtml = value => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
const imageSize = 'width="132" height="44" style="width:132px;height:44px;vertical-align:middle;"';
const imageCases = [
  {
    label: 'IMG-01 · 双引号',
    source: '<img src="data:image/png;base64,…">',
    html: `<img src="${dataUrl}" ${imageSize} alt="IMG-01 double quoted source">`,
  },
  {
    label: 'IMG-02 · 双引号重复',
    source: '与 IMG-01 的 MIME 和字节完全一致',
    html: `<img src="${dataUrl}" ${imageSize} alt="IMG-02 exact duplicate">`,
  },
  {
    label: 'IMG-03 · 单引号',
    source: "<img src='data:image/png;base64,…'>",
    html: `<img src='${dataUrl}' ${imageSize} alt="IMG-03 single quoted source">`,
  },
  {
    label: 'IMG-04 · 大写标签属性',
    source: '<IMG SRC="data:image/png;base64,…">',
    html: `<IMG SRC="${dataUrl}" ${imageSize} ALT="IMG-04 uppercase IMG SRC">`,
  },
  {
    label: 'IMG-05 · 无引号 src',
    source: '<img src=data:image/png;base64,…>',
    html: `<img src=${dataUrl} ${imageSize} alt="IMG-05 unquoted source">`,
  },
  {
    label: 'IMG-06 · base64 空白',
    source: 'base64 中插入换行、空格、制表符',
    html: `<img src="${whitespaceDataUrl}" ${imageSize} alt="IMG-06 base64 with whitespace">`,
  },
  {
    label: 'IMG-07 · MIME 参数',
    source: 'data:image/png;name=diagnostic.png;base64,…',
    html: `<img src="${parameterDataUrl}" ${imageSize} alt="IMG-07 MIME name parameter">`,
  },
];

const css = `
  @page { size: Letter; margin: 1in; }
  html { background-color: #ffffff; }
  body { font-family: Calibri, Arial, "Microsoft YaHei", sans-serif; font-size: 11pt; color: #1e293b; line-height: 1.35; margin: 0; }
  p { margin: 0 0 8pt; }
  h1 { margin: 0 0 11pt; font-size: 23pt; line-height: 1.12; color: #123047; font-weight: 700; }
  h2 { margin: 12pt 0 6pt; font-size: 13pt; color: #123047; }
  h3 { margin: 9pt 0 5pt; font-size: 11pt; color: #334155; }
  .eyebrow { color: #64748b; font-size: 8pt; font-weight: 700; letter-spacing: 1px; margin-bottom: 7pt; }
  .lead { color: #475569; font-size: 10pt; }
  .caption { color: #64748b; font-size: 8pt; line-height: 1.3; }
  .page-start { page-break-before: always; }
  .class-blue { color: #2563eb; font-family: Georgia, "Times New Roman", serif; font-size: 17pt; font-weight: 700; }
  .class-shading { background-color: #ede9fe; color: #6d28d9; padding: 8pt; border: 1pt solid #c4b5fd; text-align: center; font-size: 12pt; }
  .class-highlight { background-color: #fde68a; color: #78350f; }
  .class-indent { margin-left: 24pt; text-indent: 12pt; line-height: 1.7; font-size: 10pt; }
  .class-table { border-collapse: collapse; width: 100%; table-layout: fixed; font-size: 9pt; }
  .class-table th, .class-table td { border: 1pt solid #94a3b8; padding: 6pt; vertical-align: top; }
  .class-table th { font-weight: 700; text-align: left; background-color: #e2e8f0; color: #0f172a; }
  .class-cell { background-color: #dcfce7; color: #166534; }
  .class-cell-purple { background-color: #ede9fe; color: #6d28d9; }
  .small-table { font-size: 8pt; }
  .small-table td, .small-table th { padding: 4pt; }
  ul, ol { margin-top: 4pt; margin-bottom: 7pt; padding-left: 22pt; }
  li { margin-bottom: 4pt; line-height: 1.35; }
  .quote { margin: 7pt 0; padding: 8pt 10pt; border-left: 3pt solid #7c3aed; background-color: #f5f3ff; color: #5b21b6; font-size: 10pt; }
  code { font-family: Consolas, "Courier New", monospace; font-size: 8pt; color: #475569; }
  .image-case td { padding: 5pt 6pt; vertical-align: middle; }
  .image-label { font-weight: 700; font-size: 9pt; color: #123047; margin-bottom: 3pt; }
  .image-source { font-size: 8pt; line-height: 1.2; color: #64748b; }
`;

const complexHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>复杂 HTML → DOCX 导出对照测试 · 4 页</title>
  <style>${css}</style>
</head>
<body>
  <p class="eyebrow">SYNTHETIC EXPORT FIXTURE · 01 / 04 · LETTER · 1-INCH MARGINS</p>
  <h1 style="color:#123047;background-color:#e0f2fe;padding:7pt;">01 · 颜色、背景与文字</h1>
  <p class="lead">本页使用虚构测试文本。C 表示 CSS class，I 表示内联 style；保留标签是为了对照 HTML、旧库和新库的导出结果。</p>

  <table class="class-table" style="margin:9pt 0;" cellspacing="0" cellpadding="0">
    <tr><th style="width:50%;">C01 · CSS class</th><th style="width:50%;">I01 · 内联 style</th></tr>
    <tr>
      <td><p class="class-shading">紫色文字 + 淡紫底色<br>class-shading</p></td>
      <td><p style="background-color:#ffedd5;color:#c2410c;padding:8pt;border:1pt solid #fdba74;text-align:center;font-size:12pt;">橙色文字 + 淡橙底色<br>inline style</p></td>
    </tr>
  </table>

  <p class="class-blue">C02 · 蓝色 / 17 pt / Georgia / 加粗</p>
  <p style="color:#be123c;background-color:#fff1f2;font-size:14pt;font-family:Arial,sans-serif;text-decoration:underline;line-height:1.6;">I02 · 深红文字、淡粉背景、14 pt、Arial、下划线</p>
  <p style="text-align:right;color:#0f766e;font-size:11pt;">I03 · 右对齐 / 青绿色 / 11 pt → RIGHT</p>
  <p style="text-align:center;color:#7c3aed;font-size:12pt;font-style:italic;">I04 · 居中 / 紫色 / 斜体 · CENTER</p>
  <p style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">I05 · <strong>粗体 Bold</strong> / <em>斜体 Italic</em> / <u>下划线 Underline</u> / <span style="text-decoration:line-through;">删除线 Strike</span> / <span style="background-color:#fde68a;color:#78350f;">黄色高亮 Highlight</span></p>
  <p class="class-indent">C03 · 首行缩进、整体缩进与 1.7 倍行距。第一行从这里开始；这一段重复使用中文排版测试文本，观察文字是否按可用宽度换行，段落缩进是否保留，第二行与第三行是否保持清晰的行距。字体与颜色均为静态设置。</p>

  <h2>BG · 三种背景设置</h2>
  <p class="class-shading">BG01 · class 指定的段落背景与边框</p>
  <p style="background-color:#dcfce7;color:#166534;padding:5pt;border:1pt solid #86efac;">BG02 · 内联段落背景：#dcfce7</p>
  <table class="class-table small-table" cellspacing="0" cellpadding="0"><tr>
    <td style="background-color:#dbeafe;color:#1d4ed8;">BG03 · td 内联蓝色底纹</td>
    <td bgcolor="#ffe4e6" style="background-color:#ffe4e6;color:#be123c;">BG04 · td 同时含 bgcolor 与 style</td>
  </tr></table>
  <p class="caption" style="margin-top:7pt;">字符检查：中文 &amp; English · &lt;标签文字&gt; · &quot;引号&quot; · 5&nbsp;km · café / e&#769; · 世界 🌍 · 状态 ✅ · 金额 ¥1,234.50。</p>

  <h1 class="page-start" style="page-break-before:always;color:#166534;background-color:#dcfce7;padding:7pt;">02 · 表格、合并与底纹</h1>
  <p class="lead">T01 同时包含 colspan、rowspan、嵌套表格和显式单元格底色。内容保持短小，便于比较单页布局。</p>
  <table class="class-table" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;">
    <colgroup><col style="width:18%;"><col style="width:25%;"><col style="width:37%;"><col style="width:20%;"></colgroup>
    <tr><th colspan="4" bgcolor="#123047" style="background-color:#123047;color:#ffffff;text-align:center;">T01 · 四列合并表头 / colspan="4"</th></tr>
    <tr>
      <th rowspan="2" style="background-color:#dbeafe;vertical-align:middle;">分组<br>rowspan="2"</th>
      <th colspan="2" style="background-color:#e0e7ff;text-align:center;">内容组 / colspan="2"</th>
      <th rowspan="2" style="background-color:#dbeafe;vertical-align:middle;text-align:right;">数值</th>
    </tr>
    <tr><th style="background-color:#eef2ff;">状态</th><th style="background-color:#eef2ff;">备注</th></tr>
    <tr>
      <td rowspan="2" style="background-color:#f0fdf4;vertical-align:middle;font-weight:700;">A 组<br>跨两行</td>
      <td class="class-cell">C04 · class 绿底</td><td style="background-color:#f0fdf4;">虚构记录 A-01：中文与 English。</td><td style="background-color:#f0fdf4;text-align:right;">128.50</td>
    </tr>
    <tr>
      <td style="background-color:#fff7ed;color:#c2410c;">I06 · inline 橙底</td><td style="background-color:#fff7ed;">虚构记录 A-02：&amp; / &lt; / &gt;。</td><td style="background-color:#fff7ed;text-align:right;">96.00</td>
    </tr>
    <tr>
      <td style="background-color:#f8fafc;font-weight:700;">B 组</td>
      <td colspan="3" style="background-color:#f8fafc;">
        <p style="margin-bottom:4pt;font-weight:700;">T02 · 合并单元格中的嵌套表格</p>
        <table class="class-table small-table" style="width:100%;border-collapse:collapse;" cellspacing="0" cellpadding="0">
          <tr><th style="background-color:#ede9fe;color:#6d28d9;">子项</th><th style="background-color:#ede9fe;color:#6d28d9;">值</th></tr>
          <tr><td style="background-color:#ffffff;">B-01 · 蓝色文字</td><td style="color:#2563eb;background-color:#ffffff;">Nested table ✅</td></tr>
          <tr><td style="background-color:#fef3c7;">B-02 · 黄色底纹</td><td style="background-color:#fef3c7;text-align:right;">42</td></tr>
        </table>
      </td>
    </tr>
    <tr><td colspan="3" style="background-color:#e2e8f0;font-weight:700;text-align:right;">T01 小计 · 合并三列</td><td style="background-color:#e2e8f0;text-align:right;font-weight:700;">224.50</td></tr>
  </table>

  <h2>T03 · 每个单元格直接设置的交替底纹</h2>
  <table class="class-table small-table" cellspacing="0" cellpadding="0">
    <tr><th style="width:22%;">编号</th><th style="width:48%;">虚构内容</th><th style="width:30%;text-align:center;">对齐</th></tr>
    <tr><td style="background-color:#dbeafe;">STRIPE-01</td><td style="background-color:#dbeafe;">第一行蓝色底纹 / #dbeafe</td><td style="background-color:#dbeafe;text-align:center;">居中 CENTER</td></tr>
    <tr><td style="background-color:#ffffff;">STRIPE-02</td><td style="background-color:#ffffff;">第二行白色底纹 / #ffffff</td><td style="background-color:#ffffff;text-align:right;">右对齐 RIGHT</td></tr>
    <tr><td style="background-color:#dbeafe;">STRIPE-03</td><td style="background-color:#dbeafe;">第三行蓝色底纹 / #dbeafe</td><td style="background-color:#dbeafe;text-align:left;">左对齐 LEFT</td></tr>
  </table>
  <h2>T04 · 单元格垂直对齐与边框</h2>
  <table class="class-table small-table" style="height:42pt;" cellspacing="0" cellpadding="0"><tr style="height:42pt;">
    <td style="vertical-align:top;background-color:#f0fdf4;border:1pt solid #16a34a;">TOP · 顶部</td>
    <td style="vertical-align:middle;background-color:#fffbeb;border:1pt dashed #d97706;">MIDDLE · 中部</td>
    <td style="vertical-align:bottom;background-color:#fff1f2;border:1pt dotted #e11d48;">BOTTOM · 底部</td>
  </tr></table>
  <p class="caption" style="margin-top:7pt;">预期检查：合并结构完整、嵌套表格可见、不同底纹可辨、数值列右对齐。表格内容均为合成样本。</p>

  <h1 class="page-start" style="page-break-before:always;color:#6d28d9;background-color:#ede9fe;padding:7pt;">03 · 嵌套列表与 Unicode</h1>
  <p class="lead">L01 与 L02 展示有序、无序和混合嵌套。U01 展示多语言、实体与组合字符；字体是否含有对应字形可单独观察。</p>
  <h2>L01 · 有序列表中的嵌套列表</h2>
  <ol style="list-style-type:decimal;">
    <li><strong style="color:#2563eb;">一级步骤：准备合成文本</strong>
      <ol type="A" style="list-style-type:upper-alpha;">
        <li>二级 A：检查中文、标点和 &amp; 实体。</li>
        <li>二级 B：检查文字样式。
          <ul style="list-style-type:square;">
            <li><span style="background-color:#fef3c7;color:#92400e;">三级方块：黄色高亮</span></li>
            <li><span style="color:#16a34a;">三级方块：绿色文字</span></li>
          </ul>
        </li>
      </ol>
    </li>
    <li><strong style="color:#7c3aed;">一级步骤：核对表格和图片</strong>，确认编号连续。</li>
    <li>一级步骤：记录对照观察，无需推断实际业务结论。</li>
  </ol>
  <h2>L02 · 无序列表中的有序列表</h2>
  <ul style="list-style-type:disc;">
    <li>一级圆点：左侧缩进与项目符号。
      <ul style="list-style-type:circle;">
        <li>二级空心圆：短文本一。</li>
        <li>二级空心圆：短文本二。
          <ol type="a" style="list-style-type:lower-alpha;">
            <li>三级 a：字母编号。</li>
            <li>三级 b：字母编号。</li>
          </ol>
        </li>
      </ul>
    </li>
    <li>一级圆点：<u>下划线</u>、<em>斜体</em>、<strong>粗体</strong>。</li>
  </ul>
  <blockquote class="quote">Q01 · 引用块使用左侧紫色边框、浅紫背景与 10 pt 文字。这段内容仅用于排版检查：在同一段落里观察背景、边框、内边距和换行。</blockquote>
  <h2>U01 · 多语言、符号与长段落</h2>
  <p style="font-family:Calibri,Arial,'Microsoft YaHei',sans-serif;line-height:1.5;">中文：春风拂过测试页面，标点包括「引号」、（括号）、顿号与省略号……。English: café, naïve, résumé. 组合字符：e&#769; / a&#778;。日本語：こんにちは。한국어: 안녕하세요。العربية: مرحبًا بالعالم。Emoji：🌍 ✅ 🚀 🧪。</p>
  <p style="text-align:justify;line-height:1.5;color:#334155;">U02 · 这是一段较长的合成中文文字，用来观察两端对齐、自然换行、行距与字符编码。文本中的颜色和背景都来自明确的静态样式；每一项测试都带有固定编号，方便逐项记录。在导出后，可以检查中英文是否完整、实体是否显示为字符、单词是否按段落宽度排列，以及段落末尾是否保持可读。以下是实体文字：A &amp; B、&lt;sample&gt;、&quot;quoted&quot;、10&nbsp;mm、© 2026。</p>
  <p class="caption">本测试不使用远程资源、动态脚本或实际业务意见。分页由下一页标题上的 page-break-before: always 显式指定。</p>

  <h1 class="page-start" style="page-break-before:always;color:#0f766e;background-color:#ccfbf1;padding:7pt;">04 · 嵌入图片与属性写法</h1>
  <p class="lead">下面 7 行应显示同一张 PNG，蓝、绿、黄、红四块及右侧棋盘均可见。行号和写法各不相同，图片字节完全相同。</p>
  <p class="caption">PNG 原始尺寸 264 × 88；显示尺寸 132 × 44。IMG-01～06 使用 image/png，IMG-07 额外含 name 参数。</p>
  <table class="class-table image-case" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;">
    <tr><th style="width:64%;">编号与输入写法</th><th style="width:36%;text-align:center;">预期图片</th></tr>
    ${imageCases.map((image, index) => `<tr>
      <td style="background-color:${index % 2 === 0 ? '#f0fdfa' : '#ffffff'};"><p class="image-label">${image.label}</p><p class="image-source"><code>${escapeHtml(image.source)}</code></p></td>
      <td style="background-color:${index % 2 === 0 ? '#f0fdfa' : '#ffffff'};text-align:center;">${image.html}</td>
    </tr>`).join('\n')}
  </table>
  <h2>IMG · 验证记录</h2>
  <p style="background-color:#fef3c7;color:#92400e;padding:6pt;">目视：7 张图都应出现，比例相同、颜色相同；没有图片的位置可以按 IMG 编号记录。数据：相同 MIME 与相同字节可复用同一个 MHTML 图片部件。</p>
  <p class="caption">PNG 字节数：${png.length}。SHA-256：<br><code>${imageSha256}</code></p>
  <p class="caption">结构预期：main HTML 中 7 个真实 img；按完整 Content-Type 与字节去重后为 2 个图片部件（image/png × 6；image/png; name=diagnostic.png × 1）。</p>
</body>
</html>`;

const boundaryHtml = `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>HTML → DOCX CSS 边界样本</title><style>
  @page { size: Letter; margin: 1in; }
  body { margin:0;font:11pt/1.45 Calibri,Arial,"Microsoft YaHei",sans-serif;color:#1e293b; }
  h1 { font-size:22pt;color:#123047; }
  h2 { font-size:13pt;margin:18pt 0 7pt; }
  .flex { display:flex;gap:12px;justify-content:space-between; }
  .flex div { flex:1;background:#dbeafe;color:#1d4ed8;padding:12px;text-align:center; }
  .grid { display:grid;grid-template-columns:1fr 1fr;gap:10px; }
  .grid div { background:#ede9fe;color:#6d28d9;padding:12px;text-align:center; }
  .decorative { padding:22px;border-radius:18px;background-color:#0f766e;background-image:linear-gradient(90deg,#0f766e,#2563eb);box-shadow:0 8px 20px #94a3b8;color:white; }
  .background-image { height:88px;background-color:#fff7ed;background-image:url("${dataUrl}");background-repeat:no-repeat;background-position:center;background-size:264px 88px;border:1px solid #fdba74; }
</style></head><body>
  <h1>CSS 边界样本 · 单页</h1>
  <p>以下为合成测试内容，用于观察现代浏览器布局在 Word HTML 导入中的表现。主复杂样本单独使用更直接的文字、表格和 img 写法。</p>
  <h2>B01 · Flex 三列</h2><div class="flex"><div>左侧 A</div><div>中间 B</div><div>右侧 C</div></div>
  <h2>B02 · Grid 2 × 2</h2><div class="grid"><div>单元格 1</div><div>单元格 2</div><div>单元格 3</div><div>单元格 4</div></div>
  <h2>B03 · 渐变 / 圆角 / 阴影</h2><div class="decorative">浏览器中呈现从青绿到蓝色的渐变、圆角与阴影。</div>
  <h2>B04 · CSS background-image</h2><div class="background-image"></div>
  <p style="font-size:9pt;color:#64748b;">B04 是 CSS url() 资源，没有真实 img 节点；用于区分 CSS 背景图与 img src 图片的处理路径。</p>
</body></html>`;

await writeFile(path.join(outputDirectory, 'complex.html'), complexHtml, 'utf8');
await writeFile(path.join(outputDirectory, 'boundary.html'), boundaryHtml, 'utf8');
console.log(JSON.stringify({
  html: path.join(outputDirectory, 'complex.html'),
  boundaryHtml: path.join(outputDirectory, 'boundary.html'),
  png: path.join(outputDirectory, 'diagnostic.png'),
  targetPages: 4,
  imageElements: imageCases.length,
  distinctImageBytes: 1,
  expectedNewMhtmlImageParts: 2,
  pngBytes: png.length,
  pngSha256: imageSha256,
  htmlBytes: Buffer.byteLength(complexHtml, 'utf8'),
}, null, 2));
