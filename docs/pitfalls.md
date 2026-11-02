# Claude 制作 PPT：坑点与省钱指南

## 一、核心原则

### 1. 不要让 Claude 直接读大文件
**反例**：
```python
# ❌ 1.3MB SVG 直接 Read → 上下文爆炸
Read(file_path="第08页_完整改进_字体层级加G2.svg")
```

**正解**：
```bash
# ✓ 预提取坐标到轻量文本文件
grep -oP 'x="\K[^"]+|y="\K[^"]+|font-size="\K[^"]+' input.svg > coords.txt
```

子代理如果读了超过 200KB 的单个文件，autocompact 会在 3 轮内触发 thrashing。

---

### 2. 不要让子代理"审查代码后再补齐"
**反例**：
```
先审一遍这些 G2 分支，再补齐剩下的部分
```
这会让子代理花 5–10 轮反复 Read 已改文件，token 浪费且容易超轮次限制。

**正解**：
```
只改 slides_hw.py 的 p6 pool() 函数：GIF w=8.0，步骤字号 22/18，WARN 两行 w=3.35。
不要读其他函数，不要审查已有 G2 分支。
```

明确**单个函数**的改动点，禁止"先看再说"。

---

### 3. 验证时必须构建完整 deck
**反例**：
```python
# ❌ 只构建单页 → 页码、隐藏页数错误
prs = new_deck()
slides_end.progress(prs)
prs.save("deck.pptx")  # 只有 1 页
```

**正解**：
```python
# ✓ 完整构建：main_slides + backups + keep_left_margin
for f in build.main_slides():
    f(prs)
build.slides_end.backups(prs)
build.keep_left_margin(prs)
# 应为 22 页，其中 6 页隐藏
```

lines.py / check_layout 等审计脚本依赖正确的页码索引。

---

### 4. 图片检查用视觉模型，不要 Read 图片
**反例**：
```python
Read(file_path="p-08.png")  # ❌ 二进制图片进上下文
```

**正解**：
```bash
# 调用本地 imgcheck 脚本（通过视觉模型做 OCR + 布局分析）
imgcheck p-08.png
```

---

## 二、PPT 生成的 token 陷阱

### 1. python-pptx 的冗长 XML
**现象**：一个 22 页 PPTX 用 python-pptx 生成后，`presentation.xml` 可能有 50–100KB。

**省 token 方案**：
- 不要在每次改动后 Read 整个 PPTX 文件验证
- 用专用审计脚本（lines.py / check_layout.py）输出文本结果，先落盘再 grep
- 只在最终验收时渲染 PDF → PNG

---

### 2. 子代理的"部分结果"陷阱
**现象**：子代理达到 30 轮限制后，返回 `NOTE: this agent stopped at its 30-turn limit before finishing. The text below is PARTIAL output`。

**应对**：
- 任务拆解时，单个子代理的范围不超过 3–5 个函数
- 禁止"审查 + 补齐"两阶段任务，直接给出改动点
- 用 `SendMessage` 续接时，附上当前 lines.py 的具体输出（不是"还有孤字"，而是"p14 ORPHAN '准' at line 234"）

---

### 3. 并行子代理的文件冲突
**反例**：
```
三个 builder 同时改 slides_hw.py / slides_algo.py / charts.py
→ charts.py 被 hw 和 algo 都依赖 → 潜在竞态
```

**正解**：
- 明确"不要改 theme.py / charts.py，如需改动停下并报告"
- 每个子代理输出到独立 `/tmp/g2/{hw,algo,end}/` 验证目录
- 主调度最后统一构建一次

---

## 三、排版审计的成本控制

### 1. lines.py 输出先落盘
**反例**：
```bash
python3 lines.py deck.pdf 1 2 3  # ❌ 输出直接进上下文
```

**正解**：
```bash
python3 lines.py deck.pdf 1 2 3 > /tmp/lines_out.txt
grep ORPHAN /tmp/lines_out.txt  # ✓ 只看关键行
```

16 页的 lines.py 全输出可能 200+ 行，只关注 ORPHAN / OVER 行即可。

---

### 2. PDF 渲染只在验收时做一次
**流程**：
```
改代码 → py_compile 语法检查 → (迭代)
         ↓
    最终一次：PPTX → PDF → PNG → lines.py
```

不要每改一个函数就渲染一次 PDF。

---

### 3. 批量图片检查
如果要对比多页：
```bash
for i in 08 10 11 12; do
  imgcheck p-$i.png > /tmp/check_$i.txt &
done
wait
grep -h "问题\|overlaps\|orphan" /tmp/check_*.txt
```

---

## 四、字体与光晕的坑

### 1. PowerPoint 的 glow 渲染不可控
**问题**：python-pptx 可以设置 `glow_effect`（`color`, `radius`, `transparency`），但：
- LibreOffice 导出 PDF 时，glow 被 **忽略**
- PowerPoint 2016/2019 对 glow 的渲染半径与 python-pptx 的 `Pt()` 单位不完全一致
- 预览 PDF 里看不到光晕 ≠ 实际投影时没效果

**应对**：
- 只在 G2 需求明确时才用 glow，否则优先调整背景 / 文字颜色对比度
- 验收时必须在 PowerPoint 中打开 PPTX，不能只看 LibreOffice 转的 PDF

---

### 2. 中文字体映射
**问题**：`python-pptx` 的 `font.name = "Noto Sans CJK SC"` 只设置了 Latin 字体槽，中文走 fallback。

**修正**：
```python
from pptx.oxml.ns import qn
run._element.rPr.rFonts.set(qn('a:ea'), 'Noto Sans CJK SC')
```

需要显式设置 East Asian (`a:ea`) 字体。

---

### 3. 字号放大后的孤字问题
**现象**：原 15pt 能容纳的文字，放大到 20pt 后自动换行，末行只剩一个字"入"。

**工具**：
```python
def lines(t, size, w):
    """G2 only: if t overflows w, break once after punctuation 
    so two lines are balanced."""
    # 在标点处断开，使两行宽度接近
```

但 `lines()` 在顿号"、"后断开可能仍造成孤字，需人工指定断点：
```python
text = ["算法指令、", "固件接入"] if G2 else "算法指令、固件接入"
```

---

## 五、Workflow / Agent 使用建议

### 1. 何时用 Workflow
**适用**：
- 需要 5+ 个独立审查维度（bugs / perf / security / ...）并行跑
- 每个维度的 verify 可以立即开始，不依赖其他维度

**不适用**：
- PPT 生成：slides_hw / slides_algo / slides_end 互相依赖 theme.py 和 build.py
- 依赖关系复杂 → Workflow 的 pipeline() 语义不清晰

---

### 2. 何时用 Agent
**适用**：
- 单个子任务明确（"只改 slides_end.py 的 p14–p16"）
- 有独立的验证目录（`/tmp/g2/end/`）

**不适用**：
- 子任务需要"审查 + 补齐"两阶段
- 子任务需要读超过 3 个文件才能决定怎么改

---

### 3. 子代理的 prompt 长度
**经验值**：
- 简短任务：≤500 字
- 中等任务：500–1000 字
- 超过 1200 字 → 拆成两个子任务

prompt 太长 → 子代理容易"理解偏差"，反而需要更多轮澄清。

---

## 六、验收清单模板

```bash
# 1. 语法检查
python3 -m py_compile build/slides_*.py build/theme.py

# 2. 构建完整 deck
python3 build_g2.py --preview  # 或 /tmp/g2/bfull.py

# 3. 页数与隐藏页
python3 -c "from pptx import Presentation; p=Presentation('deck.pptx'); 
  hidden=sum(1 for s in p.slides if s.element.get('show')=='0'); 
  print(f'{len(p.slides)} slides, {hidden} hidden')"
# 预期：22 slides, 6 hidden

# 4. 孤字 / 越界 / 重叠
python3 lines.py deck.pdf 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 > /tmp/lines.txt
grep -E "ORPHAN|OVER" /tmp/lines.txt
# 预期：clean（无输出）

# 5. 字体检查
pdffonts deck.pdf | grep -v "Noto Sans CJK SC"
# 预期：只有 Noto Sans CJK SC Medium / Bold

# 6. 与原版文字 diff
python3 -c "from pptx import Presentation
old=Presentation('output/SEA-Loop_中期答辩.pptx')
new=Presentation('output/SEA-Loop_教师答辩G2_可编辑版.pptx')
for i,(o,n) in enumerate(zip(old.slides, new.slides)):
  ot=''.join(sh.text for sh in o.shapes if sh.has_text_frame)
  nt=''.join(sh.text for sh in n.shapes if sh.has_text_frame)
  if ot!=nt: print(f'p{i+1}:', ot[:30], '→', nt[:30])"
# 预期：只有 p16 "感谢聆听" → "感谢垂听"

# 7. 关键页与参考图对比
imgcheck p-08.png > /tmp/c08.txt
imgcheck p-10.png > /tmp/c10.txt
# 人工检查 /tmp/c08.txt 是否报"字号偏小"或"重叠"
```

---

## 七、推荐工作流

```
1. 人工：提取大文件关键信息（SVG 坐标、图片 OCR）→ 轻量 .txt
2. 主调度：拆解任务（p1–3 / p4–8 / p9–13 / p14–16），写 3 个子代理 prompt
3. 子代理：并行改代码，各自验证到独立目录（/tmp/g2/{a,b,c}/）
4. 主调度：统一构建完整 deck → 跑 lines.py + imgcheck
5. 人工：在 PowerPoint 中打开 PPTX，确认光晕 / 字体实际效果
```

---

## 八、不要做的事

1. ❌ 让 Claude 从零"设计" PPT 配色 / 布局
   - 现有模板 + 参数化改动 >> 从头生成
2. ❌ 用 Claude 批量生成 100 页内容相似的 slides
   - 数据驱动 + Jinja2 模板 >> python-pptx 逐页调用
3. ❌ 让 Claude "美化" PPT
   - 美化 = 主观判断，Claude 会过度添加装饰元素
   - 改为：明确字号 / 间距 / 对齐的数值要求
4. ❌ 期望 Claude 完美还原设计稿的像素级位置
   - PowerPoint 的坐标系、字体渲染与设计工具（Figma / Sketch）不同
   - 改为：给出相对位置关系（"图表在标题下方 0.5 英寸"）

---

## 附录：常用工具脚本

### A. 提取 SVG 坐标
```python
# extract_coords.py
import re, sys
from xml.etree import ElementTree as ET

tree = ET.parse(sys.argv[1])
for text in tree.findall('.//{http://www.w3.org/2000/svg}text'):
    x = text.get('x', '0')
    y = text.get('y', '0')
    fs = text.get('font-size', '12')
    fw = text.get('font-weight', '400')
    fill = text.get('fill', '#000000')
    anchor = text.get('text-anchor', 'start')
    content = ''.join(text.itertext()).strip()
    print(f"x={x} y={y} size={fs} wt={fw} fill={fill} anchor={anchor} | {content}")
```

### B. 批量 imgcheck
```bash
# check_pages.sh
for i in $(seq -w 1 16); do
  echo "=== p$i ===" >> /tmp/all_checks.txt
  imgcheck p-$i.png >> /tmp/all_checks.txt 2>&1
done
grep -E "孤字|重叠|溢出|orphan|overlap" /tmp/all_checks.txt
```

### C. PPTX 文字提取
```python
# extract_text.py
from pptx import Presentation
import sys
prs = Presentation(sys.argv[1])
for i, slide in enumerate(prs.slides, 1):
    print(f"\n=== Slide {i} ===")
    for shape in slide.shapes:
        if shape.has_text_frame:
            print(shape.text)
```

---

**最后建议**：Claude 擅长"结构化代码生成 + 参数化调整"，不擅长"审美设计 + 像素级还原"。把 PPT 制作当成"数据可视化的代码工程"，而不是"平面设计"，会更高效。
