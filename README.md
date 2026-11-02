# Claude PPT Workflow

AI 辅助制作演示文稿的实战经验与工具集

## 项目背景

在 SEA-Loop 中期答辩 PPT 制作过程中，我们尝试了多种 AI 辅助方案，投入约 $15–20 美元后总结出一套相对高效的工作流。本项目记录这些经验，避免后来者重复踩坑。

## 推荐工作流

```
前端审美模型（生成）→ 人工/模型微调 → 生图模型（精修）
     ↓                    ↓                  ↓
  Claude Opus       python-pptx 调整    DALL-E / Midjourney
  结构化内容         字号/布局/数据      最终视觉呈现
```

### 为什么不是"全程 Claude 生成可编辑 PPTX"？

1. **审美天花板**：Claude 擅长结构化内容，但配色/排版的审美判断不如前端优化的多模态模型
2. **Token 成本**：22 页 PPT 完整生成 + 3 轮迭代 ≈ 175K–300K tokens（$2–4），且易因大文件/并行冲突返工
3. **字体渲染坑**：python-pptx 生成的 PPTX 在 PowerPoint/LibreOffice 上渲染不一致，光晕/阴影效果预览困难

### 推荐的三阶段方案

#### 阶段 1：内容生成（前端审美模型）
**工具**：Claude Opus / GPT-4 / Gemini 1.5 Pro  
**产出**：带占位符的 Markdown 大纲 + 关键数据

```markdown
# Slide 1: 封面
- 标题：SEA-Loop 自主浮标系统
- 副标题：中期答辩
- 背景：海洋场景（待配图）

# Slide 2: 问题陈述
- 痛点 1：传统浮标需要船只维护，成本高
- 痛点 2：缺乏长期自主观测能力
- 数据：年均维护成本 $50K，可用率 <60%
```

**优势**：
- 模型专注于内容逻辑，不陷入像素级排版
- Markdown 便于版本控制和人工审校
- 成本：~5K tokens

#### 阶段 2：布局与数据可视化（python-pptx + 人工/模型）
**工具**：python-pptx + Claude Sonnet（微调助手）  
**产出**：可编辑的 PPTX，图表/数据准确

```python
# 模板化生成，避免逐页调用 Claude
from pptx import Presentation
from jinja2 import Template

template = Template("""
{% for slide in slides %}
def slide_{{ loop.index }}(prs):
    s = prs.slides.add_slide(prs.slide_layouts[5])
    title = s.shapes.title
    title.text = "{{ slide.title }}"
    # ... 数据驱动的图表生成
{% endfor %}
""")
```

**人工介入点**：
- 调整字号、间距（用 `imgcheck` 或肉眼检查溢出）
- 修正图表坐标偏移
- 验证数据准确性（防止模型幻觉）

**模型辅助**：
```bash
# 只让 Claude 改具体问题，不要"全盘生成"
claude: "slides_hw.py 的 p6 pool 函数，GIF 宽度改为 8.0 英寸，步骤字号 22/18"
```

**成本**：~20K–40K tokens（如果避开大文件陷阱）

#### 阶段 3：视觉精修（生图模型）
**工具**：DALL-E 3 / Midjourney / Stable Diffusion  
**产出**：高质量 PNG/PDF，用于最终投影

```bash
# 每页截图 → 提示词工程 → 重新生图
for i in $(seq 1 16); do
  # 提取当前页的布局描述
  python3 extract_layout.py deck.pptx $i > prompt_$i.txt
  
  # 调用生图 API
  curl -X POST https://api.openai.com/v1/images/generations \
    -d "prompt=$(cat prompt_$i.txt), professional presentation style, 16:9"
done
```

**优势**：
- 生图模型的字体渲染、光影效果远超 python-pptx
- 可以"参考原版 + 微调提示词"，而非从零设计
- 不受 PowerPoint 兼容性困扰

**成本**：~$0.04/图 × 16 页 = $0.64

---

## 目录结构

```
claude-ppt-workflow/
├── README.md                    # 本文件
├── docs/
│   ├── pitfalls.md             # 坑点详解（token 陷阱、字体问题等）
│   ├── workflow-guide.md       # 三阶段工作流详解
│   └── cost-analysis.md        # 成本对比（全程 Claude vs 混合方案）
├── tools/
│   ├── extract_coords.py       # 从 SVG 提取坐标（避免读大文件）
│   ├── lines.py                # 孤字/越界/重叠检查
│   ├── check_layout.py         # 布局审计
│   └── imgcheck_batch.sh       # 批量图片检查
├── skills/
│   └── ppt-audit/              # Claude Code Skill：PPT 验收清单
└── examples/
    └── sea-loop-case-study.md  # 真实案例：中期答辩 PPT 制作记录
```

---

## 快速开始

### 1. 安装依赖
```bash
pip install python-pptx jinja2 pillow
```

### 2. 生成你的第一个 PPT
```bash
# 从 Markdown 大纲生成
python3 tools/md2pptx.py outline.md -o draft.pptx

# 审计排版问题
python3 tools/lines.py draft.pdf 1 2 3

# （可选）批量生图精修
bash tools/render_with_dalle.sh draft.pptx output/
```

### 3. 在 Claude Code 中使用 Skill
```bash
# 复制 skill 到 Claude 项目
cp -r skills/ppt-audit ~/.claude/skills/

# 在 Claude Code 中调用
/ppt-audit deck.pptx
```

---

## 成本对比

| 方案 | Token 消耗 | 美元成本 | 质量 | 耗时 |
|-----|-----------|---------|-----|-----|
| 全程 Claude 生成可编辑 PPTX | 175K–300K | $2–4 | 6/10 | 4–6 小时 |
| **推荐：三阶段混合** | 25K–45K | $0.5–1 + $0.64 图 | 8/10 | 2–3 小时 |
| 纯人工 + 设计师 | 0 | $100–300 | 9/10 | 1–2 天 |

**结论**：三阶段方案在成本、质量、效率上找到最佳平衡点。

---

## 真实案例

### SEA-Loop 中期答辩 G2 版尝试

**需求**：22 页可编辑 PPTX，字号放大（34–36pt 标题），微光晕保护文字，按设计稿精确对齐。

**实际消耗**：
- 主会话：56K tokens
- 3 个子代理：187K tokens（2 个因读 1.3MB SVG 崩溃）
- 总计：~243K tokens ≈ $3.5

**有效输出**：只有 3/16 页通过验证（slides_end.py 的 p14–p16）

**失败原因**：
1. 子代理读了设计稿 SVG（1.3MB）→ 上下文爆炸
2. 并行 3 个 builder 改同一套代码 → 协调成本高
3. PowerPoint glow 效果在 LibreOffice 预览时不可见 → 返工

**教训**：见 `docs/pitfalls.md`

---

## 贡献指南

欢迎补充：
- 更多语言的模板（目前只有中文案例）
- 其他工具集成（Figma → PPTX、Canva API 等）
- 成本优化技巧

提交 PR 前请确保：
1. 代码通过 `ruff check`
2. 文档用 `markdownlint` 校验
3. 示例 PPTX 文件 <5MB（用 Git LFS 或外链）

---

## 许可证

MIT License

---

## 致谢

本项目的经验来自 SEA-Loop 团队在实战中的试错，特别感谢：
- 用户 ansty 提供的真实需求与成本约束
- Claude Code 的子代理机制（虽然踩了坑，但验证了并行协作的可能性）
- python-pptx 社区的文档与示例

---

**如果这个项目帮你省了钱或时间，请给个 ⭐️**
