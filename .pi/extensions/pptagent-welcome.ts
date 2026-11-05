// @ts-nocheck
// PPTAgent 欢迎面板：pi 打开本项目时，在输入框上方展示项目简介与全部选项。
// 机制：session_start 事件 + ctx.ui.setWidget({ placement: "aboveEditor" })。
// 对齐：内容行按「显示宽度」（CJK 记 2 列、其余记 1 列）在运行时补齐，边框不会错位。

var W = 74 // 内容区显示宽度（不含左右边框各 1 列）

// 与终端一致的显示宽度估算：只用「确定宽」字符 —— CJK 全宽，ASCII/制表线半宽。
// 刻意避开 U+00B7、U+2014、U+2192 等歧义宽度字符（在不同终端下宽窄不定，会破坏对齐）。
function dwidth(s) {
  var w = 0
  for (var _i = 0, _a = Array.from(s); _i < _a.length; _i++) {
    var cp = _a[_i].codePointAt(0)
    if (
      (cp >= 0x1100 && cp <= 0x115f) ||
      (cp >= 0x2e80 && cp <= 0x303e) ||
      (cp >= 0x3041 && cp <= 0x33ff) ||
      (cp >= 0x3400 && cp <= 0x4dbf) ||
      (cp >= 0x4e00 && cp <= 0x9fff) ||
      (cp >= 0xa000 && cp <= 0xa4cf) ||
      (cp >= 0xac00 && cp <= 0xd7a3) ||
      (cp >= 0xf900 && cp <= 0xfaff) ||
      (cp >= 0xfe30 && cp <= 0xfe4f) ||
      (cp >= 0xff00 && cp <= 0xff60) ||
      (cp >= 0xffe0 && cp <= 0xffe6) ||
      (cp >= 0x20000 && cp <= 0x3fffd)
    ) {
      w += 2
    } else {
      w += 1
    }
  }
  return w
}

function padLine(s) {
  var gap = W - dwidth(s)
  return gap > 0 ? s + ' '.repeat(gap) : s
}

function renderBox(lines) {
  var out = ['\u256d' + '\u2500'.repeat(W + 2) + '\u256e']
  for (var _i = 0, lines_1 = lines; _i < lines_1.length; _i++) {
    var l = lines_1[_i]
    out.push(l === '' ? '\u2502' + ' '.repeat(W + 2) + '\u2502' : '\u2502 ' + padLine(l) + ' \u2502')
  }
  out.push('\u2570' + '\u2500'.repeat(W + 2) + '\u256f')
  return out
}

// 注意：pi 1.4.0 对字符串数组 widget 硬截断在 MAX_WIDGET_LINES = 10 行
// （dist/modes/interactive/interactive-mode.js）。本面板恰好 10 行，
// 再加行会被切掉并显示 "... (widget truncated)"；要放更多内容得换组件工厂形式。
var PANEL_LINES = [
  '  PPTAgent：文档 / 思路 -> 高信息密度、图表丰富的 .pptx（内网私有模型）',
  '  入口：/ppt <文档路径>，或直接说需求（例：把 input.md 做成汇报）',
  '  选项（报编号即可切换；不指定则按默认值走）',
  '  类型  1 研讨型（默认）   2 汇报型（结论先行）',
  '  风格  3 经典藏青（默认） 4 深蓝商务 5 明亮宣讲 6 墨绿战略 7 紫灰科技',
  '  速度  8 慢（默认，大纲过目+warning归零）  9 快（少停顿直接出图）',
  '  流程：需求澄清 -> 输入分析 -> 大纲 -> 成批生成 -> build 一条命令出图',
  '  交付：.pi/out/ 主题.pptx（自动通过质检与字体自检）',
]

export default function (pi) {
  pi.on('session_start', async function (_event, ctx) {
    if (!ctx || !ctx.hasUI) return
// 挂在编辑器（输入框）上方：下方空间小会截断长面板，上方空间大且随会话内容滚动。
    ctx.ui.setWidget('pptagent-welcome', renderBox(PANEL_LINES), {
      placement: 'aboveEditor',
    })
  })
}
