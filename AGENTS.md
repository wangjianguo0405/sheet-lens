# AGENTS.md

## 项目

**SheetLens** — 本地优先的 Excel 查看器（Tauri 2 + React 19 + Tailwind）。

- 打开工作簿后，展示 Excel 实际的渲染结果（单元格、合并、边框、填充、字体、打印区域、浮在单元格之上的形状和图片）。
- 渲染由自研网格渲染器负责，而不是现成的电子表格库；布局直接读工作簿 XML。
- 支持 `xlsx`、`xlsm`、`xls`、`xlsb`、`ods`；老格式没有 zip 结构，所以回退到普通默认样式。

## 命令

```bash
pnpm install
pnpm dev                 # Vite 开发，热重载
pnpm build               # tsc + vite build
pnpm tauri dev           # Tauri 开发（热重载）
pnpm tauri build --no-bundle  # 便携 exe，输出到 src-tauri/target/release
```

## 测试

```bash
cd src-tauri && cargo test
```

套件针对 `samples/sample-design.xlsx` 运行。两个被忽略的测试会读取你自己机器上的工作簿，这是校对真实文档解析的唯一方式。

## 技术栈

- Tauri 2（Rust，后端）
- React 19 + Tailwind（前端）
- pnpm 11.7.0，Node >= 22
- TypeScript 严格模式

## 约定

- 保持现有项目约定。
- 文件变更较少、集中处理。
- 非必要不引入新依赖。

## 标准对照

清单在 `dev-harness/ooxml-inventory.md`，分两轴：字段（标准定义、文件在用、代码没读）和画法（Excel 怎么画）。

- 新字段落地前，先写清它和哪些既有功能共用同一块 XML 或同一块画面区域，并为组合造一张夹具（例：`<row hidden>` × `autoFilter`）。
- 准绳是 Excel 自己：COM 读数、`CopyPicture`、`ExportAsFixedFormat` 导出的 PDF。`dump_web_fixture` 只覆盖解析层，前端绘制的改动要回到真实文档的页面上量。
- 修完用户报的 bug 必须回填清单：哪一条本该覆盖、为什么没覆盖、清单改了什么。只改代码不改清单不算完。
