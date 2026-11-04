import os
import sys
from fontTools.ttLib import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont as RLTTFont

ROOT = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(ROOT, "fonts")

REQUIRED_FONTS = {
    "EBGaramond-Regular": "EBGaramond-Regular.ttf",
    "EBGaramond-Medium": "EBGaramond-Medium.ttf",
    "EBGaramond-Italic": "EBGaramond-Italic.ttf",
    "EBGaramond-SemiBold": "EBGaramond-SemiBold.ttf",
    "EBGaramond-Bold": "EBGaramond-Bold.ttf",
    "LibreBaskerville-Regular": "LibreBaskerville-Regular.ttf",
    "LibreBaskerville-Bold": "LibreBaskerville-Bold.ttf",
}

# Critical non-ASCII glyphs used in manuscript_styled.docx
CRITICAL_CHARS = "ʻ→≠₂°~—–…“”‘’©öääáéàî"


def main():
  print("=== UNEXPLAINED: FONTS & GLYPH VALIDATOR ===")
  errors = 0

  for name, filename in REQUIRED_FONTS.items():
    path = os.path.join(FONTS_DIR, filename)
    if not os.path.exists(path):
      print(f"[CRITICAL-FONT] Missing font file: {path}")
      errors += 1
      continue

    # 1. Test ReportLab TTFont loading (ensures TrueType outlines, not CFF .otf)
    try:
      pdfmetrics.registerFont(RLTTFont(name, path))
    except Exception as e:
      print(f"[CRITICAL-FONT] ReportLab rejected {filename}: {e}")
      errors += 1
      continue

    # 2. Check character map coverage via fontTools
    tt = TTFont(path)
    cmap = tt.getBestCmap() or {}
    missing = [
         f"{ch} (U+{ord(ch):04X})" for ch in CRITICAL_CHARS if ord(ch) not in cmap
    ]
    if missing:
      print(
          f"[WARN-GLYPH] {filename} is missing glyphs: {', '.join(missing)}"
      )
    else:
      print(f"[OK] {filename} loaded & covers all critical glyphs.")

  if errors > 0:
    print(f"\nFAILED with {errors} critical font error(s).")
    sys.exit(1)
  print("\nSUCCESS: All fonts validated for ReportLab & KDP.")


if __name__ == "__main__":
  main()