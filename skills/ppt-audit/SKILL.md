---
name: ppt-audit
description: Audit PPT layout and content quality
---

# PPT Audit Skill

Validates generated PowerPoint presentations for common issues.

## Usage

```bash
/ppt-audit <deck.pptx>
```

## Checks Performed

1. **Page count**: Verify expected number of slides and hidden pages
2. **Orphan text**: Single-character lines that should wrap
3. **Overlapping shapes**: Text boxes colliding with each other
4. **Font consistency**: All fonts match specification
5. **Image resolution**: Check if images are high-quality enough
6. **Data accuracy**: Compare chart values with source data

## Output

```
✓ 22 slides, 6 hidden (expected)
✗ p14 ORPHAN '准' at y=280
✗ p15 OVER '入' ~ '接口记录' (17x14)
✓ Fonts: Noto Sans CJK SC Medium/Bold only
⚠ p08 chart: 1 label overlaps axis
```

## Requirements

- python-pptx
- pdf2image (for rendering checks)
- Pillow

## Configuration

Edit `~/.claude/skills/ppt-audit/config.yaml`:

```yaml
expected_slides: 22
hidden_slides: 6
fonts_allowed:
  - "Noto Sans CJK SC"
orphan_threshold: 2  # characters
```
