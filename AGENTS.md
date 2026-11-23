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

## 技术栈

- Tauri 2（Rust，后端）
- React 19 + Tailwind（前端）
- pnpm 11.7.0，Node >= 22
- TypeScript 严格模式

## 约定

- 保持现有项目约定。
- 文件变更较少、集中处理。
- 非必要不引入新依赖。
- 界面文案三份同时加、键名相同（`src/locales/en.json`、`zh-CN.json`、`ja.json`）；界面上有语言切换，默认跟随系统，选择会记住。
- 右键菜单顺序固定，所有菜单共用一份（参照 Windows 资源管理器）；新加项按类别插入，菜单里没有的类别跳过：
  1. 打开 / 重新加载 / 打开方式——条目自己的动作
  2. 展开 / 折叠——树形列表收起或铺开，只有文件树的目录行有
  3. 复制——子菜单固定「文件名 → 绝对路径 → 相对路径」，目录没有文件名
  4. 打开文件位置 / 在资源管理器中打开——告诉你它在哪
  5. 固定 / 重命名——整理类
  6. 删除 / 关闭——破坏性动作，永远最后，前面加分隔线

## 标准对照

清单在 `dev-harness/ooxml-inventory.md`，分两轴：字段（标准定义、文件在用、代码没读）和画法（Excel 怎么画）。它记录的是本机语料与标准文档的对照，含真实文档信息，按约定不入库（`dev-harness/` 在 gitignore 里），新克隆的仓库里没有这份文件。

- 新字段落地前，先写清它和哪些既有功能共用同一块 XML 或同一块画面区域，并为组合造一张夹具（例：`<row hidden>` × `autoFilter`）。
- 准绳是 Excel 自己：COM 读数、`CopyPicture`、`ExportAsFixedFormat` 导出的 PDF。`dump_web_fixture` 只覆盖解析层，前端绘制的改动要回到真实文档的页面上量。
- 修完用户报的 bug 必须回填清单：哪一条本该覆盖、为什么没覆盖、清单改了什么。只改代码不改清单不算完。
