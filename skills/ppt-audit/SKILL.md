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

1. **Slide count**: Total slides and hidden slide count
2. **Font consistency**: Check if all text runs use consistent font names
3. **Orphan characters**: Any single-character line (任何只有单个字的行)

## Output

```
✓ 22 slides, 6 hidden
✓ Fonts: Noto Sans CJK SC
⚠ p14 possible orphan: '准'
```

## Requirements

- python-pptx
