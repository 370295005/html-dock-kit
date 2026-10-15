# 单元格内边距与垂直对齐验收

版本：html-docx-kit 0.1.1。应用：macOS Microsoft Word 16.113.4。

## 原因与修复

旧的复杂样本中，BG03/BG04 的 table 同时使用 class-table 与 small-table。Word 没有完整应用这组 class 的后代规则，导致原定四边 4pt 的 padding 没有进入单元格属性；单个 class 的表格可以保留 6pt。新旧包的历史对照都存在这个缺失，因此“效果相同”不代表完整还原了 CSS。

0.1.1 默认先用 Juice 客户端解析静态规则，按优先级写入元素的 style，再清理已经内联的复杂选择器规则，减少 Word 再次解释它们造成的段距偏差。保留全局 body/p 等基础规则，让 Word 继续建立默认字体和段落样式；不改写 @page、@media 等条件块和伪元素规则，不下载外部资源。

## 实测

| 检查 | 结果 |
| --- | --- |
| 原复杂样本 BG03/BG04 | 上、右、下、左全部为 80 twips，即 4pt；蓝色和粉色底纹保留 |
| padding.html 的 7 个测试单元格 | 四边全部为 80 twips |
| 4 个背景单元格 | verticalAlign=center，段后间距为 0 |
| TOP / MIDDLE / BOTTOM | 分别保持 top / center / bottom，段后间距为 0 |
| 字体默认规则 | body 中的 Arial 规则保留，未因清理 CSS 改成另一种西文字体 |
| 自动化测试 | 42 项通过，包括原 24 项与新增 18 项样式测试；TypeScript 构建通过 |
| npm tarball 离线安装 | 独立消费目录仅安装 1 个包，ESM/CJS、Blob/Buffer 与 NodeNext 类型声明通过，无额外运行时依赖 |
| 原复杂样本的图片 | 7 个 Word drawing 保留，全部可见 |

验收打开的是转换器生成的原始 DOCX，随后在 Word 另存为原生副本，并从 Word 本地打印成 PDF 检查真实颜色和排版。JSON 记录来自该原生副本的 WordprocessingML，PDF/原生副本仅用于 QA。

## 居中写法

padding 与 vertical-align 相互独立。原复杂样本 BG 行的源规则是 vertical-align:top，所以修复后仍按顶部对齐；库不会强制改成居中。

需要居中时，可以这样写：

```html
<td style="padding:4pt; vertical-align:middle; background-color:#ffe4e6;">
  <p style="margin:0;">BG04 · td 同时含 bgcolor 与 style</p>
</td>
```

零段距防止 Word 的默认段后间距使文字看起来偏上。padding.html 同时以 class 后代规则设置这些属性，已验证内联后生效。

更多 CSS 生效会改变行高和分页。本次原复杂样本从历史对照的 4 页变为 6 页：表格、图片和文字均完整，表格后的说明文字延续到下一页。0.1.1 不承诺与 0.1.0 的页数一致；可用 inlineStyles:false 关闭预处理。其他 Word 版本仍需按实际材料验收。

## 文件与复现

- padding.html：自包含 HTML，包含三种背景方式和三种垂直对齐。
- export.mjs：生成 padding-fixed.docx 和 padding-without-inlining.docx。
- padding-fixed.docx：开启默认内联的原始导出。
- padding-without-inlining.docx：同一 HTML 关闭预处理的对照导出。
- word-audit.json：最终样本的 Word 单元格属性记录。
- audit-word.py：检查 Word 另存副本中的 7 个诊断单元格。
- ../complex-export/results/complex-fixed-v0.1.1.docx：原复杂 HTML 的 0.1.1 导出。

在库工作区执行 `node examples/table-padding/export.mjs`。若需复核原生属性，在 Word 打开生成的文档并另存一份副本，再执行 `python3 examples/table-padding/audit-word.py /path/to/native-copy.docx`。
