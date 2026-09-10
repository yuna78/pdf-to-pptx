#!/usr/bin/env python3
"""生成测试用的合成 PDF —— 一份「幻灯片式、中英混排、含密集表格」的假方案。

Generate the synthetic test fixture: a slide-shaped, CJK+Latin, table-heavy fake deck.

刻意不用任何真实客户材料。内容全部虚构。
Deliberately contains no real client material. All content is invented.

依赖 Google Chrome（headless 打印 PDF）。产物已提交进仓库，所以跑测试的人
不需要 Chrome —— 只有要重新生成 fixture 时才需要。
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "fixtures" / "sample-deck.pdf"

CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "google-chrome",
    "chromium",
    "chromium-browser",
]

HTML = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><style>
  @page { size: 1280px 720px; margin: 0; }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans SC", sans-serif; }
  .slide { width: 1280px; height: 720px; padding: 48px 64px; position: relative;
           page-break-after: always; overflow: hidden; background: #fff; }
  .band { position: absolute; top: 0; left: 0; width: 100%; height: 14px; background: #4a2f6f; }
  .foot { position: absolute; bottom: 22px; right: 64px; color: #8a8a8a; font-size: 13px; }
  h1 { color: #4a2f6f; font-size: 40px; margin: 24px 0 6px; }
  h2 { color: #4a2f6f; font-size: 28px; margin: 18px 0 14px; }
  .sub { color: #666; font-size: 18px; margin: 0 0 26px; }
  .cover-title { font-size: 46px; font-weight: 700; color: #4a2f6f; line-height: 1.35; }
  ul { font-size: 17px; line-height: 1.9; padding-left: 22px; }
  li b { color: #4a2f6f; }
  table { border-collapse: collapse; width: 100%; font-size: 13px; margin-top: 8px; }
  th { background: #4a2f6f; color: #fff; padding: 7px 9px; text-align: left; }
  td { border: 1px solid #d6cde4; padding: 6px 9px; vertical-align: top; }
  td.t { white-space: nowrap; font-weight: 700; color: #4a2f6f; }
  .cards { display: flex; gap: 18px; margin-top: 20px; }
  .card { flex: 1; border: 2px solid #4a2f6f; border-radius: 10px; padding: 16px; }
  .card h3 { margin: 0 0 8px; color: #4a2f6f; font-size: 19px; }
  .card p { margin: 0; font-size: 14px; line-height: 1.7; color: #444; }
  .bar { background: #4a2f6f; color: #fff; border-radius: 6px; padding: 9px 14px;
         margin-bottom: 9px; font-size: 14px; }
  .bar b { display: block; font-size: 13px; font-weight: 700; }
</style></head><body>

<div class="slide">
  <div class="band"></div>
  <div style="margin-top:210px" class="cover-title">示例科技有限公司<br>年度业务回顾与规划</div>
  <p class="sub" style="margin-top:18px">Sample Corp · Annual Business Review</p>
  <div class="foot">1</div>
</div>

<div class="slide">
  <div class="band"></div>
  <h1>一、三种交付模式</h1>
  <p class="sub">同样叫「交付」，成本结构和适用场景并不一样。</p>
  <div class="cards">
    <div class="card"><h3>MODE 01 标准化</h3><p>模板复用、边际成本低，适合需求高度重复的场景。重点看复用率与维护成本。</p></div>
    <div class="card"><h3>MODE 02 半定制</h3><p>骨架固定、内容替换，适合同行业不同客户。重点看抽象层的稳定性。</p></div>
    <div class="card"><h3>MODE 03 全定制</h3><p>逐单设计，适合旗舰客户与灯塔项目。重点看人力投入与回款周期。</p></div>
  </div>
  <div class="foot">2</div>
</div>

<div class="slide">
  <div class="band"></div>
  <h1>二、全年排期</h1>
  <table>
    <tr><th style="width:130px">时间</th><th style="width:230px">模块</th><th>内容</th></tr>
    <tr><td class="t">09:30-09:50</td><td class="t">开场：目标对齐</td><td>回顾上一年度结论，明确本次评审的三个判断标准</td></tr>
    <tr><td class="t">09:50-11:00</td><td class="t">板块一：市场与竞争</td><td>区域份额、价格带迁移、主要竞争者的动作与我们的应对</td></tr>
    <tr><td class="t">11:00-11:15</td><td class="t">茶歇</td><td>短暂休息</td></tr>
    <tr><td class="t">11:15-11:50</td><td class="t">板块二：产品线复盘</td><td>按毛利与增速four象限分档，确定收缩、维持与加投的品类</td></tr>
    <tr><td class="t">11:50-12:20</td><td class="t">板块三：渠道效率</td><td>把订单、履约和售后打通成一张表，清理重复与冲突口径</td></tr>
    <tr><td class="t">12:20-13:20</td><td class="t">午餐</td><td>午餐与交流</td></tr>
    <tr><td class="t">13:20-14:35</td><td class="t">工作坊一：预算重排</td><td>基于上午结论重排明年预算，现场改写口径、重组科目并检查勾稽关系</td></tr>
    <tr><td class="t">14:35-15:45</td><td class="t">工作坊二：组织盘点</td><td>盘点关键岗位与继任梯队，标出单点依赖并给出补位方案</td></tr>
    <tr><td class="t">15:45-16:00</td><td class="t">茶歇</td><td>短暂休息</td></tr>
    <tr><td class="t">16:00-17:00</td><td class="t">工作坊三：风险清单</td><td>把已知风险按影响与概率排序，逐条指定责任人与观察指标</td></tr>
    <tr><td class="t">17:00-17:30</td><td class="t">总结与答疑</td><td>回顾三个板块与三场工作坊的结论，集中回答问题</td></tr>
  </table>
  <div class="foot">3</div>
</div>

<div class="slide">
  <div class="band"></div>
  <h1>三、负责人简介 | 张三 Alex</h1>
  <p class="sub">业务运营实战专家</p>
  <ul>
    <li>示例科技首席运营官，某大学管理学硕士。精通 Operations Research 与多渠道履约体系设计，</li>
    <li><b>深谙如何将复杂的运营方法（S&amp;OP、精益）转化为零售、制造等垂直行业的可执行动作，帮助企业实现从 0 到 1 的体系搭建。</b></li>
    <li>为多家企业提供运营转型实战培训与咨询服务，深入一线，积累了丰富的第一手经验。</li>
  </ul>
  <h2 style="font-size:20px">擅长领域：</h2>
  <ul style="font-size:15px">
    <li>利用标准化流程高效搭建公司级的运营、履约及服务体系</li>
    <li>Excel / SQL 办公与分析赋能</li>
    <li>指标口径治理、数据清洗、自动化实践</li>
  </ul>
  <div class="foot">4</div>
</div>

<div class="slide">
  <div class="band"></div>
  <h1>四、公司简介</h1>
  <p style="font-size:16px;line-height:1.75;color:#333">
    Sample Corp is a fictional company created solely as a test fixture for the
    pdf-to-pptx converter. Any resemblance to a real organisation is coincidental.<br>
    示例科技是为测试而虚构的公司，<b>95%</b>的内容由脚本生成，与任何真实机构无关。
  </p>
  <div style="margin-top:20px">
    <div class="bar"><b>More than 15 years of service experience in a fictional market</b>虚构市场中超过15年的经营与服务经验积累</div>
    <div class="bar"><b>6 professional consultants with extensive experience</b>由来自不同领域的6位资深顾问组成</div>
    <div class="bar"><b>Completed more than 30 projects within one year</b>平均每年完成30个左右的定制项目和相关咨询服务</div>
    <div class="bar"><b>High-standard complete working process &amp; quality control</b>高质量的交付源于与客户的深度定制和共创经验</div>
  </div>
  <div class="foot">5</div>
</div>

</body></html>
"""


def find_chrome() -> str:
    for c in CHROME_CANDIDATES:
        p = c if (c.startswith("/") and Path(c).exists()) else shutil.which(c)
        if p:
            return p
    print("找不到 Chrome / Chromium，无法重新生成 fixture。"
          "（仓库里已提交现成的 fixtures/sample-deck.pdf，跑测试不需要 Chrome。）",
          file=sys.stderr)
    sys.exit(1)


def main() -> int:
    chrome = find_chrome()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        html = Path(tmp) / "fixture.html"
        html.write_text(HTML, encoding="utf-8")
        subprocess.run(
            [chrome, "--headless", "--disable-gpu", "--no-sandbox",
             f"--print-to-pdf={OUT}", "--no-pdf-header-footer",
             f"--user-data-dir={tmp}/profile", f"file://{html}"],
            check=True, capture_output=True,
        )
    print(f"fixture: {OUT}  ({OUT.stat().st_size // 1024} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
