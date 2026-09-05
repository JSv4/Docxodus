#!/usr/bin/env python3
"""Assemble captured PNG frames into the README GIF.

Not part of the test suite — the companion to capture-redline-gif.mjs. Frames are
downscaled and quantised to a shared adaptive palette so the animation stays
legible (the wire console is small monospace text) without the file getting silly.
"""
import os
import sys
from PIL import Image

frame_dir = os.environ.get("FRAME_DIR", "/tmp/redline-frames")
out_path = sys.argv[1] if len(sys.argv) > 1 else "/tmp/redline-theater.gif"
width = int(os.environ.get("WIDTH", "960"))
fps = float(os.environ.get("FPS", "10"))
colors = int(os.environ.get("COLORS", "128"))

names = sorted(n for n in os.listdir(frame_dir) if n.endswith(".png"))
if not names:
    sys.exit(f"no frames in {frame_dir}")

frames = []
for name in names:
    img = Image.open(os.path.join(frame_dir, name)).convert("RGB")
    if img.width != width:
        height = round(img.height * width / img.width)
        img = img.resize((width, height), Image.LANCZOS)
    frames.append(img)

# One palette for the whole animation: a per-frame palette makes the background
# shimmer between frames, which on a dark UI is very visible.
palette_source = frames[len(frames) // 2].quantize(colors=colors, method=Image.MEDIANCUT)
quantised = [f.quantize(palette=palette_source, dither=Image.FLOYDSTEINBERG) for f in frames]

quantised[0].save(
    out_path,
    save_all=True,
    append_images=quantised[1:],
    duration=int(1000 / fps),
    loop=0,
    optimize=True,
    disposal=2,
)
size_mb = os.path.getsize(out_path) / (1024 * 1024)
print(f"{out_path}: {len(quantised)} frames, {frames[0].width}x{frames[0].height}, {size_mb:.2f} MB")
