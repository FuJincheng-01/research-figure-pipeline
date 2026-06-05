#!/usr/bin/env python3
"""Compare an exported preview with a reference image and save an amplified diff."""

from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageChops, ImageStat


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("reference")
    parser.add_argument("preview")
    parser.add_argument("--out", default="preview_diff_x4.png")
    args = parser.parse_args()

    ref = Image.open(args.reference).convert("RGB")
    preview = Image.open(args.preview).convert("RGB")
    preview_resized = preview.resize(ref.size, Image.Resampling.LANCZOS)
    diff = ImageChops.difference(ref, preview_resized)
    stat = ImageStat.Stat(diff)
    mean_abs = sum(stat.mean) / 3
    rms = (sum(v * v for v in stat.rms) / 3) ** 0.5
    amp = diff.point(lambda p: min(255, p * 4))
    out = Path(args.out)
    amp.save(out)
    print(f"reference={ref.size} preview={preview.size}")
    print(f"mean_abs={mean_abs:.3f} rms={rms:.3f} diff={out}")


if __name__ == "__main__":
    main()
