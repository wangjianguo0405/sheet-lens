import { readFileSync, writeFileSync, existsSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const here = dirname(fileURLToPath(import.meta.url))

const MAIN = [
  {
    id: 'e', file: 'e-window-lens.svg', ref: 'reference2.webp',
    name: 'E · 窗口 + 镜片（照第二张重画）',
    note: '无底圆，直接是窗口本体：蓝标题栏 + 三个圆点、表头三格、四行单元格，镜片压住右下角。环和把手保留一点渐变来留住原图的立体感，但没有高光和投影。',
  },
  {
    id: 'd', file: 'd-lens-card.svg', ref: 'reference.png',
    name: 'D · 圆底 + 卡片 + 镜片（照第一张重画）',
    note: '蓝绿渐变圆底，白卡片压在圆上。整体更「应用图标」，但少了窗口那层意思。',
  },
]

const SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256]

function inline(svg, uid, size) {
  return svg
    .replace(/<\?xml[^>]*\?>/, '')
    .replace(/width="1024" height="1024"/, `width="${size}" height="${size}"`)
    .replace(/id="([\w-]+)"/g, `id="$1-${uid}"`)
    .replace(/url\(#([\w-]+)\)/g, `url(#$1-${uid})`)
    .replace(/\s+/g, ' ')
    .trim()
}

const read = f => readFileSync(join(here, f), 'utf8')
const dataUri = f => existsSync(join(here, f))
  ? `data:image/${f.endsWith('.webp') ? 'webp' : 'png'};base64,${readFileSync(join(here, f)).toString('base64')}`
  : ''

const cards = MAIN.map(v => {
  const svg = read(v.file)
  const ref = dataUri(v.ref)
  const strip = SIZES.map(s => `
    <div class="sz"><span class="szlabel">${s}</span>${inline(svg, `${v.id}${s}`, s)}</div>`).join('')
  return `
  <section class="card">
    <header>
      <h2>${v.name}</h2>
      <p>${v.note}</p>
      <code>${v.file}</code>
    </header>
    <div class="compare">
      ${ref ? `<div class="pane"><span class="cap">参考图（AI 生成）</span><img src="${ref}" alt="reference"></div>` : ''}
      <div class="pane"><span class="cap">扁平矢量 256</span>${inline(svg, `${v.id}cmp`, 256)}</div>
      <div class="pane"><span class="cap">64</span>${inline(svg, `${v.id}cmp64`, 64)}</div>
      <div class="pane"><span class="cap">32</span>${inline(svg, `${v.id}cmp32`, 32)}</div>
    </div>
    <div class="strip">${strip}</div>
    <div class="mocks">
      <div class="mock taskbar-dark">
        <span class="mock-label">任务栏 · 深色</span>
        <div class="tb">${[16, 24, 32, 40].map(n => inline(svg, `${v.id}td${n}`, n)).join('')}</div>
      </div>
      <div class="mock taskbar-light">
        <span class="mock-label">任务栏 · 浅色</span>
        <div class="tb">${[16, 24, 32, 40].map(n => inline(svg, `${v.id}tl${n}`, n)).join('')}</div>
      </div>
      <div class="mock tabbar">
        <span class="mock-label">窗口标题栏</span>
        <div class="win"><span class="dot"></span>${inline(svg, `${v.id}win`, 16)}<span>SheetLens — 売上実績.xlsx</span></div>
      </div>
    </div>
  </section>`
}).join('')

const html = `<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>SheetLens 图标</title>
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 32px 28px 64px;
    font: 14px/1.6 -apple-system, "Segoe UI", "Noto Sans SC", sans-serif;
    background: #f4f6fa; color: #1a1f2b;
  }
  h1 { font-size: 20px; margin: 0 0 4px; letter-spacing: -0.01em; }
  .lede { margin: 0 0 26px; color: #5c6579; font-size: 13px; }
  .card { background: #fff; border: 1px solid #e2e7f0; border-radius: 14px; padding: 22px 24px; margin-bottom: 22px; }
  h2 { font-size: 15px; margin: 0 0 2px; }
  .card header p { margin: 0 0 4px; color: #5c6579; font-size: 13px; }
  code { font: 12px ui-monospace, Menlo, Consolas, monospace; color: #8b93a7; }
  .compare { display: flex; gap: 22px; align-items: flex-end; flex-wrap: wrap; margin: 18px 0 4px; }
  .pane { display: flex; flex-direction: column; gap: 8px; align-items: center; }
  .pane .cap { font-size: 11px; letter-spacing: 0.04em; text-transform: uppercase; color: #9aa2b4; }
  .pane img { width: 232px; height: auto; display: block; }
  .pane > svg { display: block; }
  .strip { display: flex; align-items: flex-end; gap: 22px; flex-wrap: wrap; padding: 18px 16px; background: #fbfcfe; border: 1px solid #e9edf5; border-radius: 10px; margin-top: 18px; }
  .sz { display: flex; flex-direction: column; align-items: center; gap: 6px; }
  .szlabel { font-size: 10px; color: #9aa2b4; font-variant-numeric: tabular-nums; }
  .sz > svg { display: block; }
  .mocks { display: flex; gap: 14px; margin-top: 18px; flex-wrap: wrap; }
  .mock { border-radius: 10px; padding: 12px 14px; display: flex; flex-direction: column; gap: 10px; }
  .mock-label { font-size: 11px; letter-spacing: 0.04em; text-transform: uppercase; }
  .taskbar-dark { background: #1c2027; } .taskbar-dark .mock-label { color: #6f7888; }
  .taskbar-light { background: #eef1f7; border: 1px solid #e2e7f0; } .taskbar-light .mock-label { color: #9aa2b4; }
  .tabbar { background: #fff; border: 1px solid #e2e7f0; } .tabbar .mock-label { color: #9aa2b4; }
  .tb { display: flex; align-items: center; gap: 12px; }
  .win { display: flex; align-items: center; gap: 8px; font-size: 12px; color: #5c6579; }
  .win .dot { width: 9px; height: 9px; border-radius: 3px; background: #d7dde8; }
</style>
</head>
<body>
  <h1>SheetLens 图标</h1>
  <p class="lede">两张参考图各重画一版，都是 1024 矢量。每张卡片左边是参考图，右边是 256 / 64 / 32 的并排，下面是真实尺寸。重点看 16 和 32。</p>
  ${cards}
</body>
</html>`

writeFileSync(join(here, 'preview.html'), html)
console.log('preview.html written')
