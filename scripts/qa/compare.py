# Compares two screenshot folders from shoot.cjs and lists the pages that differ.
# usage: python scripts/qa/compare.py <old dir> <new dir> [min changed area in px, default 0]
# A small area threshold (e.g. 2000) ignores font anti-aliasing noise between builds.
import glob
import os
import sys

from PIL import Image, ImageChops

old, new = sys.argv[1], sys.argv[2]
min_area = int(sys.argv[3]) if len(sys.argv) > 3 else 0
changed = []
for f in sorted(glob.glob(os.path.join(new, '*.png'))):
    name = os.path.basename(f)
    ref = os.path.join(old, name)
    if not os.path.exists(ref):
        changed.append(f'{name} (new)')
        continue
    a = Image.open(f).convert('RGB')
    b = Image.open(ref).convert('RGB')
    if a.size != b.size:
        changed.append(f'{name} (size {b.size} -> {a.size})')
        continue
    box = ImageChops.difference(a, b).getbbox()
    if box and (box[2] - box[0]) * (box[3] - box[1]) > min_area:
        changed.append(f'{name} {box}')
print(f'{len(glob.glob(os.path.join(new, "*.png")))} pages compared; {len(changed)} changed')
for c in changed:
    print('  ', c)
