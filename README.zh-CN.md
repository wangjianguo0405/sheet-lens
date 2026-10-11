# SheetLens

[English](README.md) | **简体中文** | [日本語](README.ja.md)

快速、本地优先的 Excel 查看器。打开工作簿就看到 Excel 会显示的内容——不用等 Excel。

基于 Tauri 2、Rust（单元格取值用 [calamine](https://github.com/tafia/calamine)）和 React 19 + Tailwind。网格是自己写的渲染器，而不是套用某个电子表格库：它直接读工作簿的 XML 还原布局——列宽、行高、合并、边框、填充、字体、打印区域，以及浮在单元格上方的形状和图片——再用 CSS 再现 Excel 的渲染结果。

它是照着日文设计文档（画面定義書 / 機能定義書）打磨出来的，所以 CJK 字体、注音（ルビ）以及这类文档惯用的排版方式就是它调校的目标场景。`xlsx`、`xlsm`、`xls`、`xlsb`、`ods` 都能打开；旧格式没有 zip 结构可读，因此退回默认样式。

## 能做什么

- 工作表列表、搜索、跳转、列筛选和公式栏
- 单元格、整行/整列、矩形区域选择，Ctrl+C 按 Excel 自己的格式复制：列间用制表符，行间用换行
- Ctrl+滚轮缩放，走 Excel 自己的档位，右下角有它那个控件
- 单元格上方的形状和图片：标注、连接线、花括号、渐变——按 ECMA-376 的预设几何表绘制，而不是逆向猜出来的
- 打印区域按分页预览的样子呈现：区域外着色，边界上那条粗蓝线
- 合并区域、文字溢出到空的邻格、缩进，以及按 Excel 的方式裁决共享边上的边框

## 运行

```bash
pnpm install
pnpm tauri dev                   # 开发，带热重载
pnpm tauri build --no-bundle     # 绿色版 exe 输出到 src-tauri/target/release
```

## 测试

```bash
cd src-tauri && cargo test
```

测试套件跑的是 `samples/sample-design.xlsx`，一个为压解析器而造的合成工作簿：四个工作表、一个合并标题、显式列宽、自动筛选、边框、填充、对齐与缩进，外加一张 12,000 行的工作表用来测分页。

另有两个 ignored 测试改为读你自己机器上的工作簿——这是拿真实文档验证解析的唯一办法：

```bash
SHEETLENS_XLSX="/path/to/book.xlsx" SHEETLENS_SHEET="Sheet1" \
  cargo test --release profile_switch -- --ignored --nocapture
```

## Excel 还原度是怎么验的

对着 Excel 本身、在真实文档上验：

- `Shape.CopyPicture` 会把 Excel 自己的渲染结果放进剪贴板——位图，或者是选中形状时带真实路径命令的 SVG。这是看一个标注或花括号*应该*长什么样最快的办法。
- `ExportAsFixedFormat` 把打印区域输出成矢量 PDF。
- ignored 的 dump 测试会打印每张表的形状、锚点与颜色，或者某个形状解算出的宽度和位置，于是差异可以当数字读，而不是靠猜。
- 开发构建里 `Ctrl+Shift+D` 会在屏幕上给出同一批解算结果：工作表的字体与实测数字宽度、默认列宽行高、选中单元格解算出的样式与边框，以及每个形状的框。

这件事难，大半是因为 Excel 的布局单位是它自己的。列宽是工作簿字体最大数字宽度的倍数——游ゴシック 是 8px，Calibri 是 7px——而文件里没有任何地方记录是哪一个；行高是磅；单元格的内边距也由同一个数字宽度推出来。

## 未实现

正则搜索、冻结窗格、拆分视图、书签、最近打开、对比；条件格式，以及遗留的 VML（批注、表单控件）。形状方面：`rect`、`ellipse`、`pentagon`、`line`、花括号和 `borderCallout2` 之外的预设都退回矩形，径向渐变按线性画，`pattFill` 按无填充画，标注的尾巴是三点折线。WMF 和 EMF 图片完全画不出来。
