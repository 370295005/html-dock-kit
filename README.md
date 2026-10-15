# html-docx-kit

用 TypeScript 实现的 HTML → DOCX npm 库。参考 [html-docx-js](https://github.com/evidenceprime/html-docx-js) 的 **altChunk + MHTML** 路线，支持 Node.js 和现代浏览器。HTML 的实际转换由 Microsoft Word 在打开文档时完成。

当前**尚未发布到 npm，也未检查包名是否可用**。从源码仓库执行 `npm ci`、`npm pack` 后，可安装生成的本地 npm 包：

```sh
npm install ./html-docx-kit-0.1.1.tgz
```

## Node.js

```ts
import { writeFileSync } from 'node:fs';
import { asBuffer } from 'html-docx-kit/node';

const html = `<!doctype html>
<html><head><style>
  body { font-family: Arial, sans-serif; }
  h1 { color: #245da8; }
  table { border-collapse: collapse; }
  td, th { border: 1px solid #aaa; padding: 6px; }
</style></head><body>
  <h1>项目报告</h1>
  <p>支持中文、标题、列表、表格和行内样式。</p>
  <table><tr><th>项目</th><th>状态</th></tr>
    <tr><td>HTML 导出</td><td>完成</td></tr></table>
</body></html>`;

writeFileSync('report.docx', asBuffer(html, {
  pageSize: 'A4',
  orientation: 'portrait',
  margins: { top: 720, bottom: 720 },
  properties: { title: '项目报告', creator: '产品团队' },
}));
```

CommonJS 也可使用：

```js
const { asBuffer } = require('html-docx-kit/node');
require('node:fs').writeFileSync('report.docx', asBuffer('<h1>Hello</h1>'));
```

## 浏览器

通过构建工具导入：

```ts
import { asBlob } from 'html-docx-kit';

const blob = asBlob('<h1>你好，Word</h1><p>HTML 片段也可以导出。</p>');
const url = URL.createObjectURL(blob);
const link = document.createElement('a');
link.href = url;
link.download = 'report.docx';
document.body.appendChild(link);
link.click();
link.remove();
setTimeout(() => URL.revokeObjectURL(url), 1000);
```

无需构建工具时，可引入 `dist/html-docx-kit.global.js`，通过全局变量 `htmlDocx.asBlob(html, options)` 使用。打开 `examples/browser.html` 即可试用，生成过程在本地完成。

## API

| API | 返回值 | 用途 |
| --- | --- | --- |
| `htmlToDocx(html, options?)` | `Uint8Array` | 通用入口，同步生成 DOCX |
| `asBlob(html, options?)` | `Blob` | 浏览器下载；在现代 Node 中也始终返回 Blob |
| `asBuffer(html, options?)` | `Buffer` | 仅从 `html-docx-kit/node` 导入 |
| `DOCX_MIME_TYPE` | `string` | DOCX 标准 MIME 类型 |

主入口也提供默认导出 `{ htmlToDocx, asBlob }`。所有输出函数都是同步函数，不需要 `await`。接受完整 HTML 或 HTML 片段；解析器会补齐 `html/head/body` 并统一为 UTF-8。空字符串会报错。

| 选项 | 默认值 | 说明 |
| --- | --- | --- |
| `orientation` | `'portrait'` | `'portrait'` 或 `'landscape'` |
| `pageSize` | `'letter'` | `'letter'`、`'A4'` 或 `{ width, height }`，自定义尺寸单位为 twips，方向按长短边确定 |
| `margins` | 见下方 | 可只覆盖部分页边距 |
| `externalImages` | `'error'` | 外部或相对图片地址默认报错；`'preserve'` 原样保留地址，不下载、不嵌入 |
| `inlineStyles` | `true` | 导出前内联文档中 `<style>` 的静态规则；`false` 保留 0.1.0 的直接封装行为 |
| `compression` | `6` | ZIP 压缩等级，整数 `0–9` |
| `properties` | 空字符串 | `{ title?, creator?, description? }`，写入并转义文档元数据 |

长度单位与参考项目一致：**1 英寸 = 1440 twips，1 磅 = 20 twips**。默认页边距为 `top/right/bottom/left: 1440`、`header/footer: 720`、`gutter: 0`。尺寸和页边距必须为整数，页边距必须留下有效内容区域。`header/footer` 设置页眉、页脚的距离，并不创建页眉、页脚内容。

## 图片

将图片放在 `<img src="data:image/png;base64,...">` 中。支持单引号、双引号、无引号、大小写标签、带参数的 data URI，以及百分号编码的图片 data URI。相同 MIME 类型、相同字节的图片只打包一次。

```ts
const html = `<p>图示：</p><img src="data:image/png;base64,${base64}" width="200">`;
const bytes = htmlToDocx(html);
```

PNG、JPEG、GIF 等图片的实际显示由 Word 的导入器决定。库会封装 `image/*` 数据，但不把 SVG、WebP 等格式转为 PNG。远程图片需由调用方预先下载并转为 data URI；在浏览器中受该图片服务器的 CORS 设置约束。

缺少或为空的 `src` 不会生成图片部件。非空 `srcset` 会报错：请先选择一张图片写入 `src`。`<picture>` 的候选源、CSS 背景图片、外部样式表等资源不会自动打包，建议使用 `<img src>` 和内联 `<style>`。

## 实现与兼容性

```text
HTML → 静态 CSS 内联 → parse5 解析/补齐文档 → 提取图片 → UTF-8/base64 MHTML
     → OOXML 关系与页面设置 → fflate ZIP → .docx
                                              ↓ Word 打开
                                      转换为 Word 文档内容
```

生成的包包含 `[Content_Types].xml`、根关系、`word/document.xml`、`word/_rels/document.xml.rels`、`word/afchunk.mht` 和 `docProps/core.xml`。`w:altChunk` 使用 Microsoft Word 要求的大小写敏感 `aFChunk` 关系类型；MHTML 类型为 `message/rfc822`。详见 [Microsoft 说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c391c28f-1b03-4a21-a4f8-4d9cddd4a95c)。

HTML 和图片均使用完整 base64 编码，按 [RFC 2045](https://www.rfc-editor.org/rfc/rfc2045) 折行为 76 字符，并使用 CRLF；MHTML 资源引用遵循 [RFC 2557](https://www.rfc-editor.org/rfc/rfc2557)。固定 ZIP 时间戳使同样输入和选项生成相同字节。

- 目标是支持 altChunk 的 Microsoft Word；参考项目明确标注 LibreOffice、Google Docs 不支持这条路线，旧版 Mac Word 2008 也不支持。不承诺其他办公软件兼容性。
- 默认用 Juice 的同步客户端把静态 CSS 写入元素的 style，减少 Word 对多 class 和后代选择器导入不完整的问题。实际支持的 CSS 属性仍由 Word 决定；flex、grid、复杂网页布局与浏览器视觉效果不能保证一致。不会请求外部样式表或执行脚本，也不会计算网页运行后的样式。
- 这是基于 altChunk 的文档封装，不是直接将每个 HTML 节点转换为原生 WordprocessingML。需要广泛办公软件兼容时，应另做原生 OOXML 转换引擎。
- 转换器不会执行 HTML 中的 JavaScript，也不是 HTML 清洗器。保留用户输入的内容，调用方可按业务需要先清洗。
- Node.js 要求 18.17+，支持 Node 20；浏览器使用现代环境并提供 atob。`asBlob` 需要 Blob 支持。同步处理会占用调用线程，大文档可放入 Worker。

## 表格内边距与居中

0.1.1 修复了 Word 导入 `class="class-table small-table"` 等组合时，后代规则中的单元格 padding 可能丢失的问题。普通调用无需改变，默认开启样式内联。更多样式生效后，行高和分页可能随之变化；需要原有直接封装方式时可传 `{ inlineStyles: false }`。

内边距和垂直对齐是两个独立设置。需要居中时，同时明确单元格的 vertical-align 和内容段落的 margin：

```html
<td style="padding:4pt; vertical-align:middle; background-color:#ffe4e6;">
  <p style="margin:0;">带内边距、垂直居中的内容</p>
</td>
```

`examples/table-padding/padding.html` 包含多 class 层叠、三种背景写法和 top/middle/bottom 对齐；运行 `node examples/table-padding/export.mjs` 可生成开启/关闭内联的两份对照 DOCX。库按源样式保留对齐方式，不会统一强制居中。

## 开发与验证

```sh
npm ci
npm run typecheck
npm test
npm run example
npm pack
```

`dist` 同时包含 ESM、CommonJS、浏览器 IIFE 及 `.d.ts/.d.cts` 类型声明。运行时代码已打包，无需额外安装运行时依赖。源码、构建依赖及锁文件保留在项目中。

测试检查 ZIP 与 OOXML 关系、MIME 解码、中文和 emoji、图片字节与去重、页面选项、错误处理、ESM/CJS 和无 Node 全局变量的浏览器包。格式检查不能代替真实 Word 的排版验收；可用 `examples/sample.docx` 在目标 Word 版本中检查。

0.1.1：42 项自动化测试通过，TypeScript 构建通过；npm tarball 在独立项目中离线安装后，ESM/CJS、Blob/Buffer 及 NodeNext 类型声明检查通过。已在 macOS Microsoft Word 16.113.4 验证多 class 表格的四边 4pt 内边距、零段距及 top/middle/bottom 对齐，详见 `examples/table-padding/TEST-RESULTS.md`。中文、标题样式、列表、表格和嵌入 PNG 也已在上述 Word 版本验收。其他 Word 版本仍需按使用场景验收。

## 与 html-docx-js 的关系

沿用其技术思路，重新实现 TypeScript 代码；未复制其 CoffeeScript 源码。相比参考项目，补充类型声明、明确的输出类型、HTML 片段支持、标准 MIME 编码、更完整的图片属性解析、图片去重、A4/自定义页面、参数校验和可复现输出。

MIT 许可证。已打包依赖的许可证保留在 `THIRD_PARTY_NOTICES.txt`。
