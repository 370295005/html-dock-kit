# HTML 转 DOCX 复杂样本对照测试

> 此处记录的是 0.1.0 的历史对照结果。后续发现 BG03/BG04 等多 class 表格的单元格 padding 被新旧包共同忽略；前 3 页相同仅证明两者一致，不能证明完整还原了 CSS。0.1.1 已增加 CSS 内联处理，针对 padding 和段距的 Word 验证见 `../table-padding/TEST-RESULTS.md`。

本次用同一份 HTML，分别调用 html-docx-kit 0.1.0 和 html-docx-js 0.3.1 的浏览器 asBlob API。两边使用相同的 Letter 纵向纸张和 1440 twips 页边距。实测应用为 macOS Microsoft Word 16.113.4。

## 结果

常见静态 CSS 可以生效，但需要在传入的 HTML 中提供样式。新旧库都把 HTML 交给 Word 导入，因此本次复杂样本的前 3 页在 Word 打印 PDF 转成 144 DPI PNG 后逐像素一致。换包不会自动增加 Word 的 CSS 支持范围。

| 检查项 | 新包 | 参考旧包 |
| --- | --- | --- |
| 文字颜色、字号、加粗、斜体、下划线 | 生效 | 生效 |
| style 标签中的单个 class 规则 | 生效 | 生效 |
| 行内文字背景、段落背景、表格底色 | 生效 | 生效 |
| 合并单元格、嵌套表格、嵌套列表、显式分页 | 保留 | 保留 |
| 主复杂样本页数 | 4 页 | 4 页 |
| 前 3 页的 Word 打印渲染 | 与旧包逐像素相同 | 与新包逐像素相同 |
| 图片实际有效嵌入 | 7 / 7 | 6 / 7 |
| 单引号、无引号、含空白的 base64 图片 | 统一进入 MHTML，导入尺寸一致 | 未进入 MHTML，Word 补救后尺寸约小 25% |
| 带 name 参数的图片 data URI | 有效嵌入 | 留为外部 data URI，Word 保存时地址截断，显示损坏占位 |
| 主样本原始 DOCX 大小 | 15,605 字节 | 56,623 字节 |
| flex、grid 布局 | 未还原，变为上下排列 | 相同 |
| 渐变、圆角、阴影 | 未还原；保留显式纯色背景 | 相同 |
| CSS background-image 中的 data URI | 未显示、未嵌入图片部件 | 相同 |

style 标签、class 属性和行内 style 属性在两个原始导出中都完整保留。Word 保存后的原生文档中，C01 的紫色文字 6D28D9 / 淡紫背景 EDE9FE / 12pt，C02 的蓝色 2563EB / Georgia / 17pt，以及 I02 的深红 BE123C / 淡粉 FFF1F2 / Arial / 14pt / 下划线均有明确样式记录，且新旧一致。

## 图片尺寸的限制

本次新库实现了 7 个不同 src 写法的图片导入尺寸一致，但不代表浏览器的 132px 被绝对还原。Word 对 MHTML 图片和直接 data URI 的尺寸解释不同；PNG 的 DPI 与 Word 的导入方式仍会影响物理尺寸。需要严格尺寸时应在目标 Word 版本中验收，并优先测试 pt/cm 等明确物理单位。

## 文件

- complex.html：4 个分页章节，包含文字与背景、复杂表格、Unicode 和 7 种图片源写法。
- boundary.html：flex、grid、渐变、圆角、阴影和 CSS 背景图的边界样本。
- results/complex-new.docx 与 complex-old.docx：同一复杂 HTML 的原始新旧导出。
- results/boundary-new.docx 与 boundary-old.docx：同一 CSS 边界 HTML 的原始新旧导出。
- generate.mjs：使用内置 diagnostic.png 生成自包含 HTML 的脚本，无需图片处理依赖。
- compare.mjs：使用两个独立浏览器上下文进行相同参数导出。
- results/structural-audit.json：ZIP、MIME、图片字节和 Word 原生样式的检查结果。

结果文件是转换器直接生成的 DOCX。为检查颜色和实际导入结果，另建了 Word 原生保存副本及本地打印 PDF；这些仅作为内部 QA，未替换原始导出。

## 复现

在库工作区运行：

```sh
node examples/complex-export/compare.mjs --reference /path/to/html-docx-js/dist/html-docx.js
```

参考包可以先用 npm pack html-docx-js@0.3.1 下载，再解压取 package/dist/html-docx.js。compare.mjs 只读取参考代码，不安装或改动业务系统。两个 HTML 样本均不依赖远程图片或外部样式表。

这些结论仅对应本次输入和上述 Word 版本。复杂选择器、多个 class 组合、字体替代及其他 Word 版本仍需按实际材料验收。
