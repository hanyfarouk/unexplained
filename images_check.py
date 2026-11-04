import os
import sys
from PIL import Image

ROOT = os.path.dirname(os.path.abspath(__file__))
IMAGES_DIR = os.path.join(ROOT, "images")

EXPECTED_WIDTH_PX = 1800  # 6 in * 300 DPI
EXPECTED_HEIGHT_PX = 2700  # 9 in * 300 DPI


def main():
  print("=== UNEXPLAINED: 64 IMAGES KDP VALIDATOR ===")
  errors = 0
  warnings = 0

  for i in range(1, 65):
    fname = f"CASE{i:02d}.png"
    path = os.path.join(IMAGES_DIR, fname)
    if not os.path.exists(path):
      print(f"[CRITICAL-IMAGE] Missing file: {fname}")
      errors += 1
      continue

    try:
      with Image.open(path) as img:
        w, h = img.size
        mode = img.mode
        dpi = img.info.get("dpi", (0, 0))

        # Check grayscale mode ('L' or 'LA' or 'I;16')
        if mode not in ("L", "LA", "I;16"):
          print(
              f"[WARN-MODE] {fname}: mode is '{mode}' (expected Grayscale 'L')."
              " Engine can convert on the fly."
          )
          warnings += 1

        # Check effective pixel dimensions for 6x9 @ 300 DPI
        if w < EXPECTED_WIDTH_PX or h < EXPECTED_HEIGHT_PX:
          print(
              f"[CRITICAL-RES] {fname}: {w}x{h}px is below 300 DPI target"
              f" ({EXPECTED_WIDTH_PX}x{EXPECTED_HEIGHT_PX}px)."
          )
          errors += 1
        else:
          print(f"[OK] {fname}: {w}x{h}px, mode={mode}, dpi={dpi}")
    except Exception as e:
      print(f"[CRITICAL-IMAGE] Corrupt file {fname}: {e}")
      errors += 1

  print(
      f"\nSUMMARY: {64 - errors}/64 passed, {warnings} warnings, {errors}"
      " critical errors."
  )
  if errors > 0:
    sys.exit(1)


if __name__ == "__main__":
  main()