"""UNEXPLAINED smoke test driver (sandbox-only; NOT part of the two-file scope).

Builds a complete scratch project:
  - copies the REAL manuscript, json/, fonts/ from the repo
  - generates 64 SYNTHETIC grayscale placeholder images (1800x2700, mode L,
    300 dpi, case number + name drawn for visual review)
  - uses a SMOKE-ONLY overrides.json extended with the two measured overflow
    cases (8, 26) at 45pt -- the production overrides.json stays locked to
    {11, 57} until the user approves the addition
Then runs the full transactional build (65 parts + merge + validation).
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SRC = Path('/home/z/my-project/unexplained')
SMOKE = Path('/home/z/my-project/smoke')

# ---- project copy ----------------------------------------------------------
if SMOKE.exists():
    shutil.rmtree(SMOKE)
SMOKE.mkdir(parents=True)
shutil.copy2(SRC / 'build_book.py', SMOKE / 'build_book.py')
shutil.copy2(SRC / 'manuscript_styled.docx', SMOKE / 'manuscript_styled.docx')
shutil.copytree(SRC / 'json', SMOKE / 'json')
shutil.copytree(SRC / 'fonts', SMOKE / 'fonts')
(SMOKE / 'images').mkdir()

# ---- synthetic placeholder images ------------------------------------------
cases = json.load(open(SMOKE / 'json/cases.json'))['cases']
try:
    f_big = ImageFont.truetype(
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', 220)
    f_small = ImageFont.truetype(
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', 90)
except Exception:
    f_big = f_small = ImageFont.load_default()

for c in cases:
    n = c['case_number']
    im = Image.new('L', (1800, 2700), 235)
    d = ImageDraw.Draw(im)
    # subtle vertical gradient for a "print-like" feel
    for y in range(2700):
        pass
    d.rectangle([60, 60, 1740, 2640], outline=90, width=8)
    d.text((900, 1000), f"CASE {n:02d}", font=f_big, fill=25, anchor="mm")
    d.text((900, 1250), c['case_name'], font=f_small, fill=60, anchor="mm")
    d.text((900, 2500), "SMOKE TEST PLACEHOLDER - NOT PRODUCTION ART",
           font=f_small, fill=110, anchor="mm")
    im.save(SMOKE / 'images' / f"CASE{n:02d}.png", dpi=(300, 300))
print("64 synthetic images generated")

# ---- smoke-only extended overrides -----------------------------------------
ov = {
    "schema": "unexplained-overrides/1.0",
    "source_sha256": "fc6f3047686e6cf33a47cf22debad486d5044f7f4222d6a55ef60fdfb8f80220",
    "cases": {
        "11": {"big_word_pt": 45, "reason": "Approved exception (A Connection)"},
        "57": {"big_word_pt": 45, "reason": "Approved exception (An Explosion)"},
        "8": {"big_word_pt": 45, "reason": "SMOKE TEST: measured 332.7pt at 50pt "
                                           "> 331.2pt frame (pending user approval)"},
        "26": {"big_word_pt": 45, "reason": "SMOKE TEST: measured 341.4pt at 50pt "
                                            "> 331.2pt frame (pending user approval)"},
    },
}
json.dump(ov, open(SMOKE / 'overrides.json', 'w'), indent=2)
print("smoke overrides.json written (extended: 8, 26 at 45pt)")

# NOTE: the engine's APPROVED_OVERRIDE_CASES must also list 8 and 26 for the
# smoke run; we patch the copy in-memory via a small sitecustomize-style edit.
bb = (SMOKE / 'build_book.py').read_text()
assert 'APPROVED_OVERRIDE_CASES = {"11": 45, "57": 45}' in bb
bb = bb.replace('APPROVED_OVERRIDE_CASES = {"11": 45, "57": 45}',
                'APPROVED_OVERRIDE_CASES = {"11": 45, "57": 45, "8": 45, "26": 45}  '
                '# SMOKE-EXTENDED (8, 26 pending user approval)')
(SMOKE / 'build_book.py').write_text(bb)
print("smoke build_book.py patched (smoke-extended approved list)")

# ---- run the full build ------------------------------------------------------
r = subprocess.run([sys.executable, 'build_book.py'],
                   cwd=SMOKE, capture_output=True, text=True, timeout=3000)
out = r.stdout
tail = "\n".join(out.splitlines()[-25:])
print("EXIT CODE:", r.returncode)
print("----- last 25 output lines -----")
print(tail)
if r.returncode != 0 and r.stderr:
    print("----- stderr tail -----")
    print("\n".join(r.stderr.splitlines()[-20:]))
