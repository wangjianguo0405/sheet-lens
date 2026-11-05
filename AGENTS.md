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

## 格式与检查

```bash
pnpm fmt         # prettier + cargo fmt，两边都写回
pnpm lint        # eslint
pnpm lint:fix    # eslint --fix
pnpm fmt:check   # 三项全查：prettier、cargo fmt、eslint（CI 用的是这一条）
```

- 前端格式：prettier，规则在 `.prettierrc.json`（100 列、双引号、分号、尾逗号）。
- Rust 格式：rustfmt，规则在 `src-tauri/rustfmt.toml`（`max_width = 120`）。
- 检查：eslint，规则在 `eslint.config.js`。开的是 `rules-of-hooks`、`exhaustive-deps`、`no-misused-promises`、`no-floating-promises`、`no-unsafe-*`；关掉的是 React Compiler 那一组（`refs`、`set-state-in-effect`、`immutability` 等）——这个画布网格本来就靠 ref 和命令式绘制，那组规则描述的编码风格本 app 不走。
- `release.yml` 在编译前跑这几项，不过就不打包。

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
