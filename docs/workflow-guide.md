# PPT 生成三阶段工作流详解

## 核心理念

**分工明确，各取所长**：
- **前端模型**：擅长内容结构和审美判断
- **python-pptx**：擅长数据可视化和精确布局
- **生图模型**：擅长视觉呈现和最终渲染

不要让任何一个环节做它不擅长的事。

## 阶段 1：内容生成（前端审美模型）

### 目标
生成结构化的内容大纲，明确每页的叙事逻辑、关键数据、配图需求。

### 工具选择
| 模型 | 优势 | 适用场景 |
|-----|-----|---------|
| **Claude Opus** | 长文本理解强，逻辑连贯 | 技术答辩、研究报告 |
| GPT-4 | 创意发散，风格多样 | 产品发布会、营销材料 |
| Gemini 1.5 Pro | 多模态理解（可参考图片） | 已有设计稿需改编 |

### Prompt 模板

```markdown
你是专业的演示文稿策划师。请为以下主题生成 15–20 页 PPT 的结构化大纲：

**主题**：[项目名称]
**受众**：[目标听众背景]
**时长**：[演讲时长]
**核心目标**：[展示什么]

**要求**：
1. 每页包含：标题、2–3 个关键信息点、配图建议、讲稿提示
2. 数据可视化优先用图表，避免大段文字
3. 开场 3 页建立问题场景，中间 10 页展开技术方案，结尾 3 页总结与展望
4. 标注哪些页需要动画效果（如流程图逐步展开）

**输出格式**：Markdown 或 YAML
```

### 输出示例

```yaml
slides:
  - id: 1
    title: "封面"
    content:
      - 项目名称
      - 副标题
      - 团队信息
    visual: "背景图 + logo"
    notes: "简短问候，直接切入主题"
  
  - id: 2
    title: "我们要解决什么问题？"
    content:
      - 痛点 1：现状描述
      - 痛点 2：缺失能力
      - 痛点 3：成本/效率问题
    visual: "对比图：现状 vs 理想"
    notes: "用数据说话，建立紧迫感"
```

---

## 阶段 2：布局与数据可视化（python-pptx）

### 目标
将大纲转化为可编辑的 PPTX，图表准确，布局合理，字体/颜色符合品牌规范。

### 工作流

#### 2.1 模板化生成（推荐）
**避免逐页调用 Claude**，用 Jinja2 模板 + 数据驱动：

```python
# generate.py
from pptx import Presentation
from jinja2 import Template
import yaml

# 加载大纲
with open("outline.yaml") as f:
    data = yaml.safe_load(f)

# 生成 python-pptx 代码
template = Template(open("templates/slide_template.j2").read())
code = template.render(slides=data["slides"])

# 执行生成
exec(code)
prs.save("draft.pptx")
```

**模板示例**（`templates/slide_template.j2`）：
```python
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

{% for slide in slides %}
def slide_{{ slide.id }}():
    layout = prs.slide_layouts[5]  # blank
    s = prs.slides.add_slide(layout)
    
    # 标题
    title = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
    title.text = "{{ slide.title }}"
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.bold = True
    
    # 内容点
    y = 1.5
    {% for point in slide.content %}
    txt = s.shapes.add_textbox(Inches(0.8), Inches(y), Inches(11), Inches(0.6))
    txt.text = "• {{ point }}"
    txt.text_frame.paragraphs[0].font.size = Pt(20)
    y += 0.7
    {% endfor %}
    
    # TODO: 根据 slide.visual 添加图片/图表

slide_{{ slide.id }}()
{% endfor %}
```

#### 2.2 图表生成（数据驱动）
```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# 示例：成本对比柱状图
chart_data = CategoryChartData()
chart_data.categories = ['方案A', '方案B', '方案C']
chart_data.add_series('年均成本', (50, 35, 12))

x, y, cx, cy = Inches(1), Inches(3), Inches(6), Inches(3.5)
chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
).chart
```

#### 2.3 人工微调点
生成初版后，人工检查：
1. **字号溢出**：用 `tools/lines.py` 检测孤字、越界
2. **图表坐标**：确认 X/Y 轴标签不重叠
3. **配色一致性**：定义 `COLORS = {"primary": "214C40", ...}` 并全局使用
4. **数据准确性**：对照原始数据源，防止模型幻觉

#### 2.4 模型辅助微调（可选）
**只在遇到具体问题时调用 Claude**：
```bash
# ❌ 错误：让 Claude 重新生成整页
claude: "优化 slide 8 的布局"

# ✅ 正确：明确具体问题
claude: "slides_hw.py 的 pool() 函数，第 3 个要点文字溢出，
        请将 textbox 宽度从 3.2 改为 3.5 英寸，字号从 17 降到 16"
```

**Prompt 模板**：
```markdown
文件：slides_hw.py，函数：pool()
问题：第 3 个要点在当前字号下溢出 textbox
要求：
1. 将该 textbox 宽度改为 3.5 英寸
2. 字号改为 16pt
3. 不要改动其他要点，不要改函数签名

只输出修改后的 pool() 函数代码。
```

---

## 阶段 3：视觉精修（生图模型）

### 目标
生成最终投影用的高质量 PNG/PDF，字体渲染完美，光影效果专业。

### 为什么需要这一步？
python-pptx 生成的 PPTX 存在以下问题：
1. **字体渲染不一致**：同一 PPTX 在 PowerPoint / LibreOffice / Mac Keynote 上显示不同
2. **特效受限**：glow、阴影、渐变在某些环境下不生效
3. **中文字体映射**：需要手动设置 `a:ea` 属性，否则 fallback 到宋体

生图模型可以绕过这些兼容性问题。

### 工具选择
| 模型 | 优势 | 适用场景 |
|-----|-----|---------|
| **生图模型** | 视觉效果好，字体渲染准确 | 技术类 PPT，图表为主 |
| Stable Diffusion | 开源免费，可本地部署 | 大批量生成，预算紧张 |

### 工作流

#### 3.1 提取布局描述
```python
# tools/extract_layout.py
from pptx import Presentation
import sys

prs = Presentation(sys.argv[1])
slide_idx = int(sys.argv[2]) - 1
slide = prs.slides[slide_idx]

# 提取文本内容
texts = []
for shape in slide.shapes:
    if shape.has_text_frame:
        texts.append(shape.text)

# 提取图片位置
images = []
for shape in slide.shapes:
    if shape.shape_type == 13:  # PICTURE
        images.append(f"{shape.left//914.4:.1f}in × {shape.top//914.4:.1f}in")

print(f"""
Slide {slide_idx + 1} Layout:
Title: {texts[0] if texts else 'None'}
Content: {' | '.join(texts[1:3])}
Images: {', '.join(images)}
Style: Professional tech presentation, 16:9 aspect ratio
""")
```

#### 3.2 生成提示词
```bash
python3 tools/extract_layout.py draft.pptx 8 > prompt_08.txt
cat prompt_08.txt
```

输出示例：
```
Slide 8 Layout:
Title: 控制策略：三层闭环
Content: 深度控制 PID | 预报驱动的采样策略 | 故障保护与安全上浮
Images: 2.5in × 1.8in (流程图), 8.0in × 3.2in (深度曲线图表)
Style: Professional tech presentation, 16:9 aspect ratio
```

#### 3.3 调用生图 API
```python
# 示例：调用生图模型
import sys

prompt = open(sys.argv[1]).read()
# 调用你选择的生图 API
# response = ...
```

#### 3.4 批量处理
```bash
# 批量生成所有页
for i in $(seq 1 16); do
  python3 tools/extract_layout.py draft.pptx $i > prompt_$i.txt
  python3 tools/render_with_imagegen.py prompt_$i.txt output/slide_$i.png
  sleep 2  # 避免 rate limit
done

# 合并为 PDF
convert output/slide_*.png final.pdf
```

### 人工审校点
1. **文字准确性**：生图模型可能改变标点、数字
2. **品牌一致性**：logo、配色、字体是否符合规范
3. **数据图表**：数值是否与 python-pptx 版本一致

---

## 完整流程对比

| 方案 | 特点 | 适用场景 |
|-----|-----|---------|
| **全程 Claude 生成可编辑 PPTX** | 完全可编辑，但成本较高，易踩坑 | 需要频繁修改数据 |
| **三阶段混合（推荐）** | 平衡成本与质量，视觉效果好 | 通用场景 |
| **纯人工 + 设计师** | 质量最高，但耗时长、成本高 | 品牌发布会、高端商务 |

---

## 进阶优化

### 1. 缓存复用
如果多个 PPT 共享相同模板（如系列课程、月度汇报）：
```python
# 第一次生成模板
template_code = claude("生成企业汇报 PPT 的 python-pptx 模板...")

# 后续只改数据
data = {"Q1_revenue": "1.2M", "Q2_revenue": "1.5M"}
exec(template_code.format(**data))
```

### 2. 本地 SD 替代商业 API
如果每月生成大量页面：
```bash
# 部署 Stable Diffusion XL
docker run -d -p 7860:7860 --gpus all \
  ghcr.io/automatic1111/stable-diffusion-webui

# 调用本地 API
curl -X POST http://localhost:7860/sdapi/v1/txt2img \
  -d '{"prompt": "slide layout...", "width": 1920, "height": 1080}'
```

成本优势：本地算力成本远低于商业 API。

### 3. 自动化 CI/CD（可选想法）
```yaml
# .github/workflows/generate-ppt.yml
name: Generate PPT
on:
  push:
    paths:
      - 'data/outline.yaml'
jobs:
  generate:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - run: python3 generate.py
      - run: bash tools/render_all.sh
      - uses: actions/upload-artifact@v3
        with:
          name: final-ppt
          path: output/final.pdf
```

每次更新大纲，自动生成新版 PPT。

---

## 常见问题

**Q: 为什么不让 Claude 直接输出 python-pptx 代码？**  
A: 可以，但逐页调用成本高且容易出错（坐标偏移、字体映射）。模板化 + 数据驱动更稳定。

**Q: 生图模型生成的 PNG 能转回可编辑 PPTX 吗？**  
A: 不能。如果需要后续修改，保留阶段 2 的 PPTX 作为"源文件"，阶段 3 的 PNG 仅用于最终投影。

**Q: 如何处理动画效果？**  
A: python-pptx 对动画支持有限，建议：
1. 简单动画（淡入/飞入）：用 python-pptx 的相关 API
2. 复杂动画：在 PowerPoint 中手动添加
3. 视频片段：用其他工具预渲染后插入 PPTX

**Q: 三阶段方案适合所有场景吗？**  
A: 不适合：
- 简单的文字 PPT（<10 页，无图表）：直接用 Canva 模板
- 需要频繁修改数据：保留在阶段 2，不走阶段 3
- 极致视觉要求（品牌发布会）：跳过阶段 2，直接专业设计

---

## 下一步

- 阅读 `docs/pitfalls.md` 了解详细坑点
- 查看 `examples/sea-loop-case-study.md` 真实案例
- 使用 `tools/lines.py` 检查生成的 PPT 排版质量
