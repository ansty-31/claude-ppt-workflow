#!/usr/bin/env python3
"""Extract text coordinates from SVG files to avoid reading large files in Claude."""
import re
import sys
from xml.etree import ElementTree as ET

def extract_svg_coords(svg_path):
    """Extract x, y, font-size, font-weight, fill, text-anchor from SVG text elements."""
    try:
        tree = ET.parse(svg_path)
    except Exception as e:
        print(f"Error parsing {svg_path}: {e}", file=sys.stderr)
        sys.exit(1)

    # Find all text elements
    namespaces = {'svg': 'http://www.w3.org/2000/svg'}
    texts = tree.findall('.//svg:text', namespaces)

    if not texts:
        # Try without namespace
        texts = tree.findall('.//text')

    print(f"# Extracted from {svg_path}")
    print(f"# Format: x=<pt> y=<baseline_pt> size=<pt> wt=<weight> fill=<color> anchor=<start|middle|end> | <text_content>")
    print()

    for text in texts:
        x = text.get('x', '0')
        y = text.get('y', '0')
        font_size = text.get('font-size', '12')
        font_weight = text.get('font-weight', '400')
        fill = text.get('fill', '#000000').lstrip('#')
        anchor = text.get('text-anchor', 'start')

        # Get text content (including nested tspan)
        content = ''.join(text.itertext()).strip()

        if content:
            print(f"x={x} y={y} size={font_size} wt={font_weight} fill={fill} anchor={anchor} | {content}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: extract_coords.py <input.svg>", file=sys.stderr)
        print("", file=sys.stderr)
        print("Extract text coordinates from SVG to lightweight text format.", file=sys.stderr)
        print("This avoids loading 1MB+ SVG files into Claude's context.", file=sys.stderr)
        sys.exit(1)

    extract_svg_coords(sys.argv[1])
