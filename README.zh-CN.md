# pdf-to-pptx

把 PDF 转成**可编辑**的 PowerPoint —— 文字是真文本框、色块是真矢量形状、照片是真图片，
不是每页拍一张截图。

[English](./README.md) · MIT 许可证

![转换前后对比](./examples/before-after.png)

中间那一栏是直接跑 `soffice --convert-to pptx` 的结果：文本框宽度是照着原字体量的，
换一台没有那个字体的机器，文字就在框里折行、撞进下一行。右边是同一份文件过这个工具的结果。

---

## 为什么需要它

LibreOffice 本来就有 PDF 导入过滤器，而且做得相当好——它把 PDF 的文字对象还原成文本框、
把矢量路径还原成形状。一条命令能到 90%：

```bash
soffice --headless --infilter="impress_pdf_import" --convert-to pptx deck.pdf
```

剩下的 10% 决定了产物是「一个文件」还是「一个能发给客户的文件」。这个工具就是那 10%，
而且**每一条都是对中文不友好的默认行为，且都要等你发出去之后才会发现**：

| # | LibreOffice 吐出来的 | 会坏在哪 | 修法 |
|---|---|---|---|
| 1 | 字体名没有空格：`MicrosoftYaHei` | PowerPoint 认不出这个名字，整篇乱替换 | 改名为 `Microsoft YaHei` |
| 1b | 只写 `a:latin`，从不写 `a:ea` | 每个 run 的东亚字体那一半没有声明 | 给含中文的 run 补 `a:ea` |
| 2 | 文本框按原字体量宽度，自动换行还开着 | 字体一被替换就重排到下一行——叠字 | `word_wrap = False`，去掉 autofit |
| 3 | 部分汉字变成康熙部首（`⽰` U+2F70 而不是 `示` U+793A） | 看着一模一样，但复制和搜索会静默失败 | NFKC + 一张人工核对的部首对照表 |
| 4 | 原稿有字距调整时，每个汉字之间被塞一个空格（`擅 长 领 域`） | 同上——文字「看着」是对的，但搜不到 | 在「每个汉字都被单空格隔开且≥3 字」时合并 |

第 3、4 条是大多数人不知道要去检查的。一份 deck 可以看上去像素级完美，同时里面的文字谁也搜不到。

## 安装

不打包任何二进制。两个系统依赖，一条命令自检：

```bash
./bin/pdf-to-pptx --doctor
```

| 依赖 | 装法 | 用在哪 |
|---|---|---|
| LibreOffice | `brew install --cask libreoffice` · `apt install libreoffice-impress` | 全部功能 |
| poppler | `brew install poppler` · `apt install poppler-utils` | `--verify` 和扫描件预检 |
| python-pptx, pillow | 首次运行自动装进 `.venv` | 全部功能 |

LibreOffice 装完 ~800 MB、许可证是 MPL-2.0。为四个平台各塞一份进仓库是荒谬的，
所以它保持为前置依赖——`html-to-pptx` 对 Chrome 也是这么处理的。

## 用法

```bash
# 产物默认与输入同目录同名
./bin/pdf-to-pptx deck.pdf

# 渲染回 PDF 并生成逐页左右对照图。转任何一份新 deck 的第一次都该加上。
./bin/pdf-to-pptx deck.pdf --verify

# 只比对关心的几页
./bin/pdf-to-pptx deck.pdf --verify --verify-pages 1,5,6
```

| 参数 | 用途 |
|---|---|
| `-o, --out` | 输出路径或目录（默认与输入同目录同名） |
| `--verify` | 渲染回 PDF，生成逐页左右对照图 |
| `--verify-pages 1,5,6` | 只比对这几页 |
| `--verify-dpi N` | 对照图分辨率（默认 80） |
| `--cjk-font "PingFang SC"` | 强制指定东亚字体（默认取承载中文最多的那个） |
| `--font-map "FooBar=Foo Bar"` | 追加改名规则，可重复 |
| `--keep-wrap` | 保留 LibreOffice 的换行设置（跳过修补 2） |
| `--no-text-cleanup` | 保留兼容字符与逐字空格（跳过修补 3、4） |
| `--no-postfix` | 交出 LibreOffice 原始产物，一点不改 |
| `--soffice PATH` | 指定 LibreOffice 可执行文件 |

## 要看对照图，别信统计数字

工具会打印逐页清单：

```
页数 10
  页     形状    文本框     字符    矢量    图片
  1      24      4     36    18     1
  5      85     55    815    29     1
```

**这张表说明不了画面是对的。** 叠字、色块跑偏、文字被裁，在它眼里全都看不见。
催生这个工具的那份 deck，统计清单完美无缺，其中一页却整块糊掉。`--verify` 就是干这个的——
它会输出 `cmp-NN.png`，原稿和产物并排。去看。尤其是文字最密的一页和封面。

## 已知边界

- **扫描件（无文字层）不在范围内。** 得到的是图片不是文字。工具会检测并警告，请先 OCR。
- **表格不是原生 PPT 表格。** 13 行的表会变成约 55 个独立文本框加一堆线条。
  改单元格文字没问题；插入行、调列宽、套表格样式做不到。
  如果对方需要**当表格来用**，这条路满足不了。
- **渐变、透明、混合模式**会塌成近似纯色。品牌感强的封面页务必对着对照图核。
- **中文粗体取决于字体是否真的装了。** 在没有微软雅黑的机器上渲染偏细，
  不等于粗体丢了——下结论前先查 `run.font.bold`。
- **部首对照表是人工核对的，不是穷举的。** 表外的兼容字符会原样保留并警告，绝不瞎猜。
  在 `scripts/convert.py` 的 `RADICAL_SUPPLEMENT` 里补一行就是一个 PR。
- **能不用就不用。** 先向给你 PDF 的人要源 `.pptx`。任何转换都有损；
  这个工具是源文件确实找不回来时的最优解，不是「有源文件也用它」的理由。

## 配合 Claude Code 使用

把这个仓库放进 `~/.claude/skills/pdf-to-pptx/`（或任意 skills 目录），
`SKILL.md` 会让它成为一个 skill——说「把这份 PDF 转成 PPT」就会触发。
它同时也是一个不需要 Claude 的普通命令行工具。

## 测试

```bash
python3 -m venv .venv && .venv/bin/pip install python-pptx pillow pytest
.venv/bin/python -m pytest tests/ -q
```

测试样例 `tests/fixtures/sample-deck.pdf` 是一份关于虚构公司的合成 deck，
由 `tests/make_fixture.py` 生成，**不含任何人的真实材料**。
测试断言的是「每条修补相对 `--no-postfix` 确实改变了什么」——
所以某条修补哪天悄悄失效，测试会挂，而不是空转通过。

## 许可证说明

本项目 MIT。它通过 `subprocess` 以**独立进程**方式调用 LibreOffice（MPL-2.0）
和 poppler（GPL-2.0/3.0），并不链接它们的库。按 GPL 的通行解释，这属于聚合而非衍生作品，
copyleft 条款不会传染过来。这是行业通行理解，**不是法律意见**——
若打算再分发打包版本，请自行咨询律师。
