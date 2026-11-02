# Claude PPT Workflow

用 Claude 辅助生成 PPTX 的工作流与经验总结

## 核心理念

基于真实项目的试错（22 页学术答辩 PPT，总投入 ≥$50，最后一天约 $30），得出推荐流程：

**三阶段工作流：**

1. **调教前端审美好的模型生成 pptx**（给它设计规范/参考样张，生成结构化代码或文件）
2. **人工或模型微调**（不在字体渲染问题上死磕）
3. **交给生图模型精修**（逐页生成视觉最终版，结果为图片，不可编辑，按需取舍）

### 为什么不全程用 Claude 生成可编辑 PPTX？

实战踩过的坑：
- **上下文爆炸**：让子代理读 1.3MB SVG 设计稿 → 崩溃（发生 2 次）
- **轮次浪费**：子代理花 30 轮反复"审查"已有代码，到达限制仍未完成
- **字体渲染不一致**：python-pptx 生成的 PPTX 在 PowerPoint/LibreOffice 上显示不同
- **并行协调成本高**：多个子代理修改共享模块 → 冲突频繁

---

## 三阶段工作流详解

### 阶段 1：内容生成（前端审美模型）

**工具**：Claude Opus / GPT-4 / Gemini 1.5 Pro  
**产出**：结构化大纲（Markdown 或 YAML）+ 关键数据

```markdown
# Slide 1: 封面
- 标题：项目名称
- 副标题：答辩类型
- 背景：场景配图（待生成）

# Slide 2: 问题陈述
- 痛点 1：现状问题描述
- 痛点 2：缺失能力
- 数据：关键指标
```

**优势**：
- 模型专注于内容逻辑，不陷入像素级排版
- Markdown 便于版本控制和人工审校
- 成本低

---

### 阶段 2：布局与数据可视化（python-pptx + 人工/模型）

**工具**：python-pptx + Claude（仅用于微调具体问题）  
**产出**：可编辑的 PPTX，图表/数据准确

**推荐做法：模板化生成**

```python
# 用 Jinja2 模板 + 数据驱动，避免逐页调用 Claude
from pptx import Presentation
from jinja2 import Template

template = Template("""
{% for slide in slides %}
def slide_{{ loop.index }}(prs):
    s = prs.slides.add_slide(prs.slide_layouts[5])
    # 标题、内容、图表...
{% endfor %}
""")

code = template.render(slides=data["slides"])
exec(code)
prs.save("draft.pptx")
```

**人工检查点**：
- 字号溢出：用 `tools/lines.py` 检测孤字、越界
- 图表坐标：确认 X/Y 轴标签不重叠
- 数据准确性：对照原始数据源，防止模型幻觉

**模型辅助（仅在遇到具体问题时）**：

```bash
# ✅ 明确具体问题
claude: "slides_hw.py 的 pool() 函数，第 3 个要点文字溢出，
        请将 textbox 宽度从 3.2 改为 3.5 英寸，字号从 17 降到 16"

# ❌ 不要让 Claude 重新生成整页
claude: "优化 slide 8 的布局"
```

---

### 阶段 3：视觉精修（生图模型）

**工具**：生图模型（如 Stable Diffusion 等）  
**产出**：高质量 PNG/PDF，用于最终投影

**为什么需要这一步？**

python-pptx 的字体渲染问题：
- 同一 PPTX 在不同软件（PowerPoint / LibreOffice / Keynote）显示不同
- 光晕、阴影、渐变效果在某些环境不生效
- 中文字体需手动设置 `a:ea` 属性，否则 fallback 到宋体

生图模型可以绕过这些兼容性问题。

**工作流**：

```bash
# 每页提取布局描述 → 生成提示词 → 调用生图 API
for i in $(seq 1 16); do
  python3 tools/extract_layout.py draft.pptx $i > prompt_$i.txt
  # 调用生图 API
done
```

**人工审校**：
- 文字准确性（生图模型可能改变数字/标点）
- 品牌一致性（logo、配色）
- 数据图表（数值是否一致）

---

## 核心踩坑总结

### 1. 不要让 Claude 直接读大文件

**反例**：让子代理读 1.3MB SVG → 上下文爆炸崩溃（本项目发生 2 次）

**正解**：
```bash
# 预提取关键信息到轻量文本文件（5 分钟人工）
python3 tools/extract_coords.py design.svg > coords.txt  # 4KB
```

---

### 2. 不要让子代理"审查代码后再补齐"

会让子代理花 5–10 轮反复 Read 已改文件，token 浪费且容易超轮次限制。

**正解**：明确**单个函数**的改动点，禁止"先看再说"。

---

### 3. 验证时必须构建完整 deck

**反例**：只构建单页 → 页码、隐藏页数错误

**正解**：完整构建 22 页，然后用 `lines.py deck.pdf 8 10 11` 只检查目标页

---

### 4. 中文字体映射

python-pptx 的 `font.name = "字体"` 只设置 Latin 槽，中文需显式设置 East Asian：

```python
from pptx.oxml.ns import qn
run._element.rPr.rFonts.set(qn('a:ea'), 'Noto Sans CJK SC')
```

---

### 5. LibreOffice 预览 ≠ PowerPoint 渲染

LibreOffice 导出 PDF 时会忽略某些效果（glow、阴影）。验收时必须在 PowerPoint 中打开 PPTX。

---

### 6. 单字孤行问题（中文特有）

字号放大后容易出现末行只剩一个字"入""准"等。需要在标点处手动断行或调整 textbox 宽度。

---

## 目录结构

```
claude-ppt-workflow/
├── README.md                    # 本文件
├── docs/
│   ├── pitfalls.md             # 详细坑点与省 token 指南
│   ├── workflow-guide.md       # 三阶段工作流详解
│   └── cost-analysis.md        # 成本分析（定性）
├── tools/
│   ├── extract_coords.py       # 从 SVG 提取坐标
│   ├── lines.py                # 孤字/越界/重叠检查
│   └── bfull.py                # 完整 deck 构建模板
├── skills/
│   └── ppt-audit/              # Claude Code Skill：PPT 验收
└── examples/
    └── sea-loop-case-study.md  # 真实案例分析
```

---

## 快速开始

### 1. 安装依赖

```bash
pip install python-pptx jinja2 pillow
```

### 2. 生成你的第一个 PPT

```bash
# 从 Markdown 大纲生成（需自行实现 md2pptx）
python3 generate_from_outline.py outline.md -o draft.pptx

# 审计排版问题（需先转 PDF）
soffice --headless --convert-to pdf draft.pptx
python3 tools/lines.py draft.pdf 1 2 3
```

### 3. 使用工具

**检查孤字/重叠/越界**：
```bash
python3 tools/lines.py deck.pdf 8 10 11
# 输出示例：
# p08 ORPHAN '入'
# p10 OVER '传感器' ~ '接口' (12x8)
# p11 clean
```

**提取 SVG 坐标**（避免让 Claude 读大文件）：
```bash
python3 tools/extract_coords.py design.svg > coords.txt
```

---

## Token 省钱清单

1. **预处理大文件**：>200KB 的 SVG/PDF，先提取关键信息
2. **模板化生成**：用 Jinja2 + 数据驱动，不要逐页调用 Claude
3. **明确任务边界**：子代理任务 = "改这 5 行代码"，不是"审查后决定"
4. **验证逻辑分离**：语法检查（py_compile）每次改动都做，视觉渲染（PDF）只做一次
5. **工具路径查询**：用 `which` 查询工具路径，不要让代理猜测
6. **并行任务隔离**：并行子代理不能共享依赖文件（theme.py / charts.py）

---

## 真实案例

本项目源于一个 22 页学术答辩 PPT 的制作：
- **成本**：≥$50（最后一天约 $30）
- **有效产出**：16 页主讲 + 6 页备答的可编辑 PPTX
- **失败尝试**：投影版因代理崩溃、轮次超限而终止

详见 `examples/sea-loop-case-study.md`。

---

## 工具使用

### lines.py - 排版审计

```bash
python3 tools/lines.py deck.pdf 1 2 3
# 检查第 1、2、3 页的：
# - ORPHAN: 单字孤行（中文）
# - EDGE: 越界文本
# - OVER: 重叠文本框
```

### bfull.py - 构建脚本模板

```bash
python3 tools/bfull.py --build-dir /path/to/build --output /path/to/output
# 完整构建 deck（主讲页 + 备答页 + 隐藏页）
# 自动转 PDF + 渲染 PNG
```

### ppt-audit - 验收 Skill

```bash
# 在 Claude Code 中调用
python3 skills/ppt-audit/ppt-audit deck.pptx
# 检查：页数、隐藏页、字体、单字孤行
```

---

## 适用场景

**推荐使用本工作流**：
- 技术答辩、学术汇报（数据图表为主）
- 需要频繁更新数据的报告
- 批量生成相似结构的 PPT（课程系列、月报）

**不推荐**：
- 简单文字 PPT（<10 页，无图表）→ 直接用 Canva 模板
- 极致视觉要求（品牌发布会）→ 专业设计师
- 一次性需求 + 不需要修改 → 在线工具（Gamma.ai 等）

---

## 许可证

MIT License - 详见 LICENSE 文件

---

## 致谢

本项目的经验来自真实项目的试错。感谢 python-pptx 社区的文档与示例。
