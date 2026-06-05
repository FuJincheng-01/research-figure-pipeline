#!/usr/bin/env python3
"""Create reference crops and SVG wrappers from an approved master figure.

These crops are for coordinates, style reference, user-approved sample preservation,
or emergency fallback. In the normal workflow, final modules are regenerated with
ChatGPT image2 after the full master is approved.

Manifest format:
{
  "master": "figure_master.png",
  "out_dir": "assets/modules",
  "modules": {
    "header": [0, 0, 1600, 120],
    "module_1": [20, 140, 360, 850]
  }
}
"""

from __future__ import annotations

import argparse
import base64
import json
from pathlib import Path

from PIL import Image


def svg_wrap(png_path: Path, width: int, height: int, svg_path: Path) -> None:
    data = base64.b64encode(png_path.read_bytes()).decode("ascii")
    svg = f'''<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <image href="data:image/png;base64,{data}" x="0" y="0" width="{width}" height="{height}" preserveAspectRatio="none"/>
</svg>
'''
    svg_path.write_text(svg, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("manifest", help="JSON manifest with master, out_dir, reference modules")
    args = parser.parse_args()

    manifest_path = Path(args.manifest)
    cfg = json.loads(manifest_path.read_text(encoding="utf-8"))
    base = manifest_path.parent
    master_path = (base / cfg["master"]).resolve()
    out_dir = (base / cfg.get("out_dir", "assets/modules")).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    master = Image.open(master_path).convert("RGB")
    rows = []
    for name, box in cfg["modules"].items():
        crop = master.crop(tuple(box))
        png_path = out_dir / f"{name}.png"
        svg_path = out_dir / f"{name}.svg"
        crop.save(png_path, quality=96)
        svg_wrap(png_path, crop.width, crop.height, svg_path)
        rows.append(f"{name}\tbox={tuple(box)}\tsize={crop.size}\tpng={png_path}\tsvg={svg_path}")

    (out_dir / "manifest.txt").write_text("\n".join(rows), encoding="utf-8")
    print(out_dir / "manifest.txt")


if __name__ == "__main__":
    main()
