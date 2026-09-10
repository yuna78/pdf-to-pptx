---
name: pdf-to-pptx
description: >-
  把 PDF 转成**可编辑**的 PowerPoint——文字变真文本框、色块和线条变原生矢量形状、
  照片变图片，不是每页拍一张图。走 LibreOffice 的 PDF 导入过滤器，再自动修掉它那几个
  对中文不友好的默认行为（字体名无空格 / 文本框叠字 / 汉字变康熙部首 / 逐字空格）。
  当用户说「PDF 转 PPT」「pdf to pptx」「这份 PDF 我想改字」「只有 PDF 没有源文件」
  「方案 PDF 转成可编辑 PPT」「把这个 PDF 做成 PowerPoint」时使用。
  与 `html-to-pptx` 的区别：那个只吃 HTML，喂 PDF 会被直接拒。
  扫描件（无文字层）不适用——只会得到一堆图片，得先 OCR。
---

# pdf-to-pptx — PDF 转可编辑 PowerPoint

完整文档见 [README.md](./README.md)（中文）/ [README.en.md](./README.en.md)。
本文件是给 agent 看的操作要点。

## 什么时候用

- ✅ 对方只给了 PDF，你要改字、换 logo、改署名再发出去。
- ✅ 自己的旧材料只剩 PDF，源文件找不到了。
- ❌ **扫描件 / 纯图 PDF**（无文字层）→ 先 OCR。脚本会警告。
- ❌ 输入是 HTML → 用 `html-to-pptx`。

**先问一句：能不能直接管对方要源 PPTX。** 要得到就别转——任何转换都有损。

## 用法

```bash
# 默认产物与 PDF 同目录同名
./bin/pdf-to-pptx path/to/deck.pdf

# 转一份新 PDF 的第一次，一定加 --verify
./bin/pdf-to-pptx deck.pdf --verify --verify-pages 1,5,6

# 自检依赖
./bin/pdf-to-pptx --doctor
```

常用参数见 README 的参数表；最常要的三个是
`--verify`、`--cjk-font`、`--no-postfix`（对拍用）。

## agent 必须遵守的两条

1. **转完必须看对照图，不许只看统计数字就汇报成功。**
   脚本打印的「形状 / 文本框 / 字符」清单看不出叠字、错位、文字被裁。
   跑 `--verify`，然后用 Read 逐张看 `cmp-NN.png`——至少看**文字最密的一页**和**封面**。
   催生这个工具的那份 deck 就是统计全绿、某一页整块叠字。

2. **看到「兼容字符没有对照表」的警告要如实转达。**
   那些字看着和正常汉字一样，但复制出去搜不到。脚本不会瞎猜，会原样保留并列出来。
   补 `scripts/convert.py` 的 `RADICAL_SUPPLEMENT` 一行即可解决。

## 汇报时要说清的边界

- **表格不是原生 PPT 表格**：13 行的表 = 约 55 个独立文本框 + 线条。
  改字行，插入行 / 调列宽 / 套表格样式不行。
- **中文粗体取决于字体是否装了**：渲染偏细 ≠ 粗体丢了，先查 `run.font.bold` 再下结论。
- **渐变 / 透明会被简化**成近似纯色。

## 验收清单

- [ ] 页数与原稿一致（页数不一致脚本会警告）
- [ ] `--verify` 对照图看过，重点是文字最密的一页和封面
- [ ] 没有叠字 / 文字跑出色块 / 文字被裁
- [ ] 没有「兼容字符无法还原」的残留警告
- [ ] 要外发的话，确认署名、logo、页脚是否需要换

## 仓库

本 skill 同时是开源仓库 <https://github.com/yuna78/pdf-to-pptx>（MIT）。
它首先是一个**通用命令行工具**（`bin/pdf-to-pptx`，终端 / CI / 任何脚本都能调，不依赖 Claude），
其次才是 agent skill；`SKILL.md` 是标准 skill 描述文件，任何支持 skill 或外部命令的 agent 都能用。
改动这里的文件就是改仓库，改完记得 push。
