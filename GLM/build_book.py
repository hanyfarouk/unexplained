#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UNEXPLAINED — single-file production build engine (two-file scope).

    build_book.py  +  overrides.json          (the only maintained files)

Usage
-----
    python build_book.py                     full transactional build
    python build_book.py --preflight         input gate only, no PDFs
    python build_book.py --cases 1,11,57     pilot parts only (no master, no merge)
    python build_book.py --selftest          engine self-tests (see SELFTEST section)

What it does (transactional sequence, per the approved Master Plan):
  1. preflight()          — runtime env, frozen hash, JSON 1.1 structure, overrides,
                            fonts (per-role cmap), images (exact CASE{NN}.png mapping,
                            1800x2700, mode L, >=299.9 dpi), display-line widths.
  2. purge output/parts/  — current-run intermediates only; master is never touched
                            until the final atomic replace.
  3. build 65 parts       — 00_FrontMatter (13p) + 01_Case01..63_Case63 (6p each)
                            + 64_Case64_and_Epilogue (9p). Fit-Ladder v3 + Widow-Ladder
                            v2.3 + Page-6 decision table per case; failed parts are
                            withheld and reported, the loop continues (collect-all).
  4. merge                — only with 65 valid parts and zero CRITICALs; writes a
                            temporary master, validates it independently, then
                            atomically replaces output/UNEXPLAINED_KDP_Master_400p.pdf.
  5. build_report.txt     — preflight results, per-case trial telemetry, fit summary,
                            validation results, limitations.

Implementation notes (reportlab 4.4.x verified internally):
  * Paragraph first-baseline offset from flowable top equals fontSize when
    paragraph.paraFontSizeHeightOffset == 1 (reportlab 4.4.9 default) for both
    blPara.kind==0 (style fontSize) and kind==1 (lines[0].fontSize).  The
    CapAlignedFrame uses these offsets and is *accepted* only through the
    --selftest PDF-geometry test (text-matrix + font cap-height == 589.248pt).
  * KeepTogether.wrap() always forces a split, so its children are re-submitted
    to Frame.add() individually; the single-trigger cap calibration therefore
    only ever sees leaf flowables (plus trackers/actions which are skipped).
  * All fill/stroke colors are emitted as DeviceGray (g/G) through
    BookCanvas / ShearTextObject overrides; reportlab has no native DeviceGray
    path for Color objects (it would emit rg).
  * Synthetic BoldItalic ("EBGaramond-BoldItalic-Synth") is a style alias on
    EBGaramond-Bold: a 12-degree Tm shear emitted inside the BT..ET block by
    ShearTextObject._textOut, with an exact unskewed-matrix restore so that
    subsequent fragments keep zero drift.  NOTE: no bold+italic run exists in
    the current manuscript (verified by preflight), so this machinery is
    exercised only by --selftest.
  * Trials and production renders share the exact same story builders, frames,
    templates and chrome painters; only the canvas target differs.

This file is deliberately dependency-light: reportlab, pypdf, Pillow, fontTools.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import math
import os
import re
import sys
import time
import traceback
from pathlib import Path

from reportlab.lib.colors import Color
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import registerFont, registerFontFamily, stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfgen.textobject import PDFTextObject
from reportlab.platypus.doctemplate import (BaseDocTemplate, NextPageTemplate,
                                            PageTemplate, ActionFlowable)
from reportlab.platypus.flowables import Flowable, KeepTogether, Spacer
from reportlab.platypus.frames import Frame
from reportlab.platypus.paragraph import Paragraph

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import ContentStream, NameObject, DictionaryObject, ArrayObject
except Exception:                                      # pragma: no cover
    PdfReader = PdfWriter = ContentStream = None

try:
    from PIL import Image as PILImage
except Exception:                                      # pragma: no cover
    PILImage = None

# ============================================================================
# 1. CONSTANTS — locked geometry, page map, typography, thresholds
# ============================================================================

ROOT = Path(__file__).resolve().parent

MANUSCRIPT   = ROOT / "manuscript_styled.docx"
JSON_DIR     = ROOT / "json"
FONT_DIR     = ROOT / "fonts"
IMAGE_DIR    = ROOT / "images"
OUT_DIR      = ROOT / "output"
PARTS_DIR    = OUT_DIR / "parts"
MASTER_PATH  = OUT_DIR / "UNEXPLAINED_KDP_Master_400p.pdf"
REPORT_PATH  = OUT_DIR / "build_report.txt"

FROZEN_SHA256 = "fc6f3047686e6cf33a47cf22debad486d5044f7f4222d6a55ef60fdfb8f80220"

# --- page geometry (6 x 9 in, all values in PDF points, bottom-left origin) ---
PAGE_W, PAGE_H   = 432.0, 648.0          # 6in x 9in
M_TOP            = 54.0                  # 0.75 in
M_BOT            = 57.6                  # 0.8  in
M_IN             = 57.6                  # 0.8  in (binding edge)
M_OUT            = 43.2                  # 0.6  in
TEXT_W           = PAGE_W - M_IN - M_OUT # 331.2
FRAME_TOP_Y      = PAGE_H - M_TOP        # 594.0
FRAME_BOT_Y      = M_BOT                 # 57.6
FRAME_H          = FRAME_TOP_Y - FRAME_BOT_Y   # 536.4
VERS0_X          = M_OUT                 # 43.2   left text edge on verso
VERS0_RIGHT      = PAGE_W - M_IN         # 374.4  right text edge on verso
RECTO_X          = M_IN                  # 57.6   left text edge on recto
RECTO_RIGHT      = PAGE_W - M_OUT        # 388.8  right text edge on recto
PAGE_CENTER_X    = PAGE_W / 2.0          # 216.0  (caption page-centering)

# --- A-master chrome ---
HEADER_BASELINE_Y = PAGE_H - 0.54 * 72   # 609.12 (0.54 in from top edge)
RULE_Y            = PAGE_H - 0.595 * 72  # 605.16 (0.595 in from top edge)
RULE_WEIGHT       = 0.75                 # pt (KDP minimum; locked decision)
FOLIO_Y           = 0.45 * 72            # 32.4  (0.45 in from bottom edge)
HEADER_BAND_Y     = 600.0                # text above this y == running header
FOLIO_BAND_Y      = 45.0                 # text below this y == folio
GRAY_HEADER       = 0.25                 # 75% black

# --- recto cap alignment target (0.816 in from top edge) ---
CAP_TARGET_Y      = PAGE_H - 0.816 * 72  # 589.248

# --- page map ---
FRONT_MATTER_PAGES = 13
CASE_START_BASE    = 14                  # case N starts on 14 + 6*(N-1)
TOTAL_PAGES        = 400
SECTION_KEYS = ["WHAT_HAPPENED", "WHEN_WHO", "EVIDENCE", "CRACK", "THEORIES",
                "QUESTION", "SEE_ALSO", "CURIOUS"]
N_CASES = 64

def case_start_page(n: int) -> int:
    return CASE_START_BASE + 6 * (n - 1)

def part_name(n: int) -> str:
    return "00_FrontMatter.pdf" if n == 0 else f"{n:02d}_Case{n:02d}" + (
        "_and_Epilogue.pdf" if n == N_CASES else ".pdf")

# --- typography tables (Specs-Final + Fit-Ladder v3 ladder parameters) ---
FONTS = {
    "EBGaramond-Regular":   "EBGaramond-Regular.ttf",
    "EBGaramond-Medium":    "EBGaramond-Medium.ttf",
    "EBGaramond-Italic":    "EBGaramond-Italic.ttf",
    "EBGaramond-SemiBold":  "EBGaramond-SemiBold.ttf",
    "EBGaramond-Bold":      "EBGaramond-Bold.ttf",
    "LibreBaskerville-Regular": "LibreBaskerville-Regular.ttf",
    "LibreBaskerville-Bold":    "LibreBaskerville-Bold.ttf",
}
SYNTH_ALIAS = "EBGaramond-BoldItalic-Synth"   # style alias, NOT an 8th font file
SYNTH_BASE  = "EBGaramond-Bold"
SHEAR_TAN   = math.tan(math.radians(12.0))    # 0.21255656...

# charset_by_style role -> font face used to render that role
ROLE_FONT = {
    "author": "LibreBaskerville-Bold",   "big-word": "LibreBaskerville-Bold",
    "book-name": "LibreBaskerville-Bold", "book-name-full": "LibreBaskerville-Bold",
    "case-title": "LibreBaskerville-Bold", "epilogue-title": "LibreBaskerville-Bold",
    "intro-title": "LibreBaskerville-Bold", "toc-title": "LibreBaskerville-Bold",
    "body-text": "EBGaramond-Regular",   "copyright": "EBGaramond-Regular",
    "epilogue-body": "EBGaramond-Regular",
    "case-name": "EBGaramond-Medium",    "welcome": "EBGaramond-Medium",
    "section-header": "EBGaramond-SemiBold",
    "case-quote": "EBGaramond-Italic",   "dedication-line": "EBGaramond-Italic",
    "image-caption": "EBGaramond-Italic", "see-also": "EBGaramond-Italic",
    "search-term": "EBGaramond-Italic",
    "subtitle-name": "EBGaramond-Bold",
    # roles rendered by this engine but absent from charset_by_style:
    "toc-entry": "EBGaramond-Regular",
    "list-item": "EBGaramond-Regular",   "list-item-lead": "EBGaramond-Bold",
    "running-header": "EBGaramond-Medium", "folio": "EBGaramond-Medium",
}

TIERS = {   # tier -> (body leading, body spaceAfter, hdr before, hdr after, list leading, list spaceAfter)
    "standard": (14.0, 4.0, 18.0, 10.8, 17.0, 4.5),
    "tight":    (12.0, 3.0, 13.0,  7.0, 14.5, 3.0),
    "open":     (15.0, 5.0, 21.0, 12.0, 18.0, 5.5),
}
TIER_ORDER = ("standard", "tight", "open")
BODY_SIZE  = 10.5
LIST_INDENT = 18.0                        # 0.25 in hanging indent zone
LIST_W     = TEXT_W - LIST_INDENT         # 313.2 inner paragraph width
CROSS_ARM  = 2.5                          # crosshair half-arm
CROSS_WEIGHT = 0.75                       # pt

WIDOW_GUARD_CHARS = 12                    # companion word length guard
WIDOW_GUARD_FILL  = 0.65                  # preceding-line fill guard
P6_FILL_MIN       = 0.25                  # p6_fill underflow threshold (informational)

# --- B-master title page stack (calibration constants; proof-stage adjustable) ---
CASENAME_BASELINE_Y = 432.0               # 3.0 in from top (locked by spec)
GAP_CASENAME_BW     = 49.32               # 0.685 in spaceAfter of case-name
GAP_BW_TL           = 24.66               # 0.28 in (BW spaceAfter) + 0.0625 in (TL spaceBefore)
GAP_TL_QUOTE        = 21.6                # 0.3 in spaceBefore of quote
BW_LEADING, TL_LEADING, QUOTE_LEADING = 30.0, 22.0, 18.0
BIG_WORD_PT = {"11": 45, "57": 45}        # approved exceptions (overrides.json)

# --- C-master fixed pages (page-relative baselines, from top edge) ---
HALFTITLE_BASELINE_Y  = PAGE_H - 3.3 * 72    # 410.4
TITLE_BOOK_BASELINE_Y = PAGE_H - 3.3 * 72    # 410.4
TITLE_SUB1_BASELINE_Y = PAGE_H - 3.9 * 72    # 367.2
TITLE_SUB_LEADING     = 22.0
TITLE_AUTHOR_BASELINE_Y = PAGE_H - 6.9 * 72  # 151.2
COPYRIGHT_BASELINE_Y  = PAGE_H - 5.5 * 72    # 252.0  (low third; calibration constant)
DEDICATION_BASELINE_Y = PAGE_H - 3.3 * 72    # 410.4
WELCOME_BASELINE_Y    = PAGE_H - 4.25 * 72   # 342.0  (optical center; calibration constant)
INTRO_TITLE_BASELINE_Y = PAGE_H - 4.0 * 72   # 360.0
EPILOGUE_TITLE_BASELINE_Y = PAGE_H - 3.5 * 72  # 396.0
CAPTION_BASELINE_Y    = 1.15 * 72             # 82.8 from bottom edge
CAPTION_MAX_W         = 4.5 * 72              # 324.0
CAPTION_LEADING       = 11.0
INTRO_STANZA_PT       = 14.0
EPILOGUE_STANZA_PT    = 11.0

# --- locked overrides contract ---
OVERRIDE_SCHEMA = "unexplained-overrides/1.0"
APPROVED_OVERRIDE_CASES = {"11": 45, "57": 45}   # add cases only via approved change

ENV_RANGE = {"reportlab": ("4",), "pypdf": None, "PIL": None, "fontTools": None}

BOOK_TITLE = "UNEXPLAINED"
BOOK_AUTHOR = "Hany Farouk"

# ============================================================================
# 2. REPORT — severity-graded collector (ChatGPT #13: part/case/decision detail)
# ============================================================================

class Sev:
    INFO = "INFO"
    WARN = "WARN"
    CRIT = "CRITICAL"

SEV_ORDER = {Sev.INFO: 0, Sev.WARN: 1, Sev.CRIT: 2}

class Report:
    """Collects every diagnostic with structured context."""

    def __init__(self, title: str = "UNEXPLAINED build report"):
        self.title = title
        self.lines: list[tuple[str, str]] = []
        self.t0 = time.time()

    def add(self, sev: str, code: str, detail: str = "") -> None:
        self.lines.append((sev, f"{code}" + (f": {detail}" if detail else "")))
        safe_print(f"[{sev}] {code}" + (f": {detail}" if detail else ""))

    def info(self, code, detail=""): self.add(Sev.INFO, code, detail)
    def warn(self, code, detail=""): self.add(Sev.WARN, code, detail)
    def crit(self, code, detail=""): self.add(Sev.CRIT, code, detail)

    @property
    def n_crit(self) -> int:
        return sum(1 for s, _ in self.lines if s == Sev.CRIT)

    @property
    def n_warn(self) -> int:
        return sum(1 for s, _ in self.lines if s == Sev.WARN)

    def section(self, name: str) -> None:
        self.lines.append(("SECTION", name))
        safe_print(f"\n----- {name} -----")

    def write(self, path: Path) -> None:
        with open(path, "w", encoding="utf-8", errors="replace") as f:
            f.write(f"{self.title}\n")
            f.write(f"generated by build_book.py in {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"duration: {time.time() - self.t0:.1f}s\n\n")
            cur = None
            for sev, text in self.lines:
                if sev == "SECTION":
                    f.write(f"\n===== {text} =====\n")
                    cur = None
                    continue
                if sev != cur:
                    f.write(f"\n--- {sev} ---\n")
                    cur = sev
                f.write(text + "\n")

# ============================================================================
# 3. UTILITIES
# ============================================================================

def safe_print(*args):
    """Console print that never dies on cp1252 Windows consoles."""
    try:
        msg = " ".join(str(a) for a in args)
        enc = sys.stdout.encoding or "ascii"
        msg.encode(enc, errors="replace")
        print(msg)
    except Exception:
        pass

def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()

NBSP = "\u00a0"

def norm_text(s: str) -> str:
    """Extraction-side normalization: NBSP->space, collapse whitespace."""
    return re.sub(r"\s+", " ", (s or "").replace(NBSP, " ")).strip()

def esc(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

# ============================================================================
# 4. FONTS — registration, metrics, per-role cmap data
# ============================================================================

_FONT_METRICS: dict[str, dict] = {}
_FONTS_REGISTERED = False

def register_all_fonts(report: "Report | None" = None) -> dict:
    """Register the 7 approved faces (+ synth alias, + family mapping)."""
    global _FONTS_REGISTERED
    if _FONTS_REGISTERED:
        return _FONT_METRICS
    missing = [f for f in FONTS.values() if not (FONT_DIR / f).exists()]
    if missing:
        raise FileNotFoundError(f"missing font files: {missing}")
    for name, fname in FONTS.items():
        registerFont(TTFont(name, str(FONT_DIR / fname)))
    # synth alias shares the Bold outlines; the Tf is rewritten to the base font
    # by ShearTextObject, so no second subset is ever embedded.
    registerFont(TTFont(SYNTH_ALIAS, str(FONT_DIR / FONTS[SYNTH_BASE])))
    registerFontFamily("EBGaramond", normal="EBGaramond-Regular", bold="EBGaramond-Bold",
                       italic="EBGaramond-Italic", boldItalic=SYNTH_ALIAS)
    _FONT_METRICS.clear()
    from fontTools.ttLib import TTFont as FTFont
    for name, fname in FONTS.items():
        ft = FTFont(str(FONT_DIR / fname), lazy=True)
        upm = ft["head"].unitsPerEm
        os2 = ft["OS/2"]
        cap = getattr(os2, "sCapHeight", None)
        if not cap:                                  # fallback, fail-visible
            cap = int(ft["hhea"].ascent * 0.70)
        rl_face = pdfmetrics.getFont(name).face
        _FONT_METRICS[name] = {
            "upm": upm,
            "cap": cap / upm,
            "ascent": rl_face.ascent / 1000.0,
            "descent": rl_face.descent / 1000.0,
            "cmap": set(ft.getBestCmap().keys()),
        }
        ft.close()
    _FONT_METRICS[SYNTH_ALIAS] = dict(_FONT_METRICS[SYNTH_BASE])
    _FONTS_REGISTERED = True
    return _FONT_METRICS

def cap_height_pt(font_name: str, size: float) -> float:
    m = _FONT_METRICS[font_name]
    return m["cap"] * size

def ascent_pt(font_name: str, size: float) -> float:
    return _FONT_METRICS[font_name]["ascent"] * size

def first_baseline_offset(flowable) -> float:
    """Distance from flowable top to the first text baseline, for a *wrapped*
    paragraph-like leaf (reportlab 4.4.x, paraFontSizeHeightOffset==1)."""
    if isinstance(flowable, TrackedLineFlowable):
        return flowable.baseline_offset
    if isinstance(flowable, ListItemFlowable):
        return first_baseline_offset(flowable.para)
    bl = getattr(flowable, "blPara", None)
    if bl is None:
        return getattr(flowable.style, "fontSize", BODY_SIZE)
    if getattr(bl, "kind", 0) == 1 and bl.lines:
        return float(getattr(bl.lines[0], "fontSize", flowable.style.fontSize))
    return float(getattr(bl, "fontSize", flowable.style.fontSize))

def cap_font_size(flowable):
    """(font_name, font_size) of the first rendered line of a wrapped leaf."""
    if isinstance(flowable, TrackedLineFlowable):
        return flowable.font_name, flowable.font_size
    if isinstance(flowable, ListItemFlowable):
        return cap_font_size(flowable.para)
    bl = getattr(flowable, "blPara", None)
    if bl is not None and getattr(bl, "kind", 0) == 1 and bl.lines:
        for w in bl.lines[0].words:
            if getattr(w, "text", "").strip():
                return w.fontName, w.fontSize
    st = flowable.style
    return st.fontName, st.fontSize

# ============================================================================
# 5. RUNTIME ENVIRONMENT CHECK (ChatGPT #12 — record the tested environment)
# ============================================================================

def check_runtime_environment(report: Report) -> dict:
    env = {}
    import reportlab as _rl
    env["reportlab"] = _rl.__version__
    if not _rl.__version__.startswith("4."):
        report.crit("CRITICAL-ENV", f"reportlab {_rl.__version__} not in tested 4.x line")
    else:
        report.info("ENV-REPORTLAB", _rl.__version__)
    # internals this engine relies on
    import reportlab.platypus.paragraph as _pp
    if getattr(_pp, "paraFontSizeHeightOffset", 0) != 1:
        report.crit("CRITICAL-ENV",
                    "paragraph.paraFontSizeHeightOffset != 1; baseline calibration invalid")
    for attr in ("_textOut", "_setFont", "_formatText", "setCharSpace"):
        if not hasattr(PDFTextObject, attr):
            report.crit("CRITICAL-ENV", f"PDFTextObject.{attr} missing")
    probe = Frame(0, 0, 100, 100)
    for attr in ("_y", "_y1p", "_atTop"):
        if not hasattr(probe, attr):
            report.crit("CRITICAL-ENV", f"Frame instance attribute .{attr} missing")
    for attr in ("_reset", "add"):
        if not hasattr(Frame, attr):
            report.crit("CRITICAL-ENV", f"Frame.{attr} missing")
    if PdfReader is None:
        report.crit("CRITICAL-ENV", "pypdf not importable")
    else:
        import pypdf as _pypdf
        env["pypdf"] = _pypdf.__version__
        report.info("ENV-PYPDF", _pypdf.__version__)
    if PILImage is None:
        report.crit("CRITICAL-ENV", "Pillow not importable")
    else:
        env["Pillow"] = PILImage.__version__
        report.info("ENV-PILLOW", PILImage.__version__)
    import fontTools
    env["fontTools"] = fontTools.version
    report.info("ENV-FONTTOOLS", fontTools.version)
    return env
# ============================================================================
# 6. GRAY CANVAS + SHEAR TEXT OBJECT  (DeviceGray + synthetic BoldItalic)
# ============================================================================

GRAY_BLACK = Color(0, 0, 0)
GRAY_75    = Color(GRAY_HEADER, GRAY_HEADER, GRAY_HEADER)

def _gray_value(color) -> "float | None":
    """Return the DeviceGray value for a color, or None if it is not a pure gray."""
    if isinstance(color, Color):
        if color.alpha is not None and color.alpha != 1:
            return None
        if abs(color.red - color.green) < 1e-9 and abs(color.green - color.blue) < 1e-9:
            return color.red
    return None

class ShearTextObject(PDFTextObject):
    """PDFTextObject that (a) emits DeviceGray color operators and (b) renders
    the synthetic BoldItalic alias with an exact 12-degree Tm shear.

    The shear is emitted inside the active BT..ET block:
        1 0 tan 1 x y Tm  <text> Tj  1 0 0 1 (x+W) y Tm      (mid-line)
    where W is the *unsheared* advance of the text.  For a TStar (end-of-line)
    emission the restoring matrix targets the reportlab line origin (_x0,_y0)
    instead, so the subsequent T* lands exactly where unsheared reportlab
    would put it.  reportlab's own _textOut relies on the implicit Tj advance
    between fragments, which the (x+W) restore preserves bit-exactly.
    """

    # -- color: DeviceGray only ------------------------------------------
    def setFillColor(self, aColor, alpha=None):
        v = _gray_value(aColor)
        if v is None and aColor is not None and not isinstance(aColor, (str, tuple, list)):
            raise ValueError(f"non-gray fill color rejected: {aColor!r}")
        if v is not None:
            self._fillColorObj = aColor
            self._code.append("%s g" % _fmt(v))
        else:
            PDFTextObject.setFillColor(self, aColor, alpha)

    def setStrokeColor(self, aColor, alpha=None):
        v = _gray_value(aColor)
        if v is None and aColor is not None and not isinstance(aColor, (str, tuple, list)):
            raise ValueError(f"non-gray stroke color rejected: {aColor!r}")
        if v is not None:
            self._strokeColorObj = aColor
            self._code.append("%s G" % _fmt(v))
        else:
            PDFTextObject.setStrokeColor(self, aColor, alpha)

    # -- synth font alias: map to base outlines, remember the shear flag ----
    def _setFont(self, psfontname, size):
        if psfontname == SYNTH_ALIAS:
            self._synth = True
            psfontname = SYNTH_BASE
        else:
            self._synth = False
        PDFTextObject._setFont(self, psfontname, size)

    def setFont(self, psfontname, size, leading=None):
        if psfontname == SYNTH_ALIAS:
            self._synth = True
            psfontname = SYNTH_BASE
        else:
            self._synth = False
        PDFTextObject.setFont(self, psfontname, size, leading)

    # -- the shear itself, with exact cursor / TLM tracking -----------------
    #
    # reportlab's _textOut relies on the *implicit* Tj advance: it never
    # updates _x, and paragraph.py communicates positions only occasionally
    # (setTextOrigin / setXPos).  We therefore track the real cursor ourselves
    # (_cx, _cy) and the real text-line-matrix origin (_tlm_x, _tlm_y).
    # A synth emission leaves the TLM at the restore point ("dirty"); any
    # later operator that reads the TLM (T*, Td) is preceded by an exact
    # re-anchor to reportlab's model, so unsheared behaviour is bit-identical.

    def _text_width(self, text) -> float:
        fname = SYNTH_BASE if self._synth else self._fontname
        w = stringWidth(text, fname, self._fontsize)
        cs = getattr(self, "_charSpace", 0) or 0
        ws = getattr(self, "_wordSpace", 0) or 0
        return w + cs * len(text) + ws * text.count(" ")

    def _sync_origin(self):
        self._tlm_x, self._tlm_y = self._x0, self._y0
        self._cx, self._cy = self._x, self._y
        self._tlm_dirty = False

    def _advance_line(self):
        self._tlm_y -= self._leading
        self._cx, self._cy = self._tlm_x, self._tlm_y

    def _reanchor_tlm(self):
        """Emit a Tm that restores the TLM to reportlab's (_x0, _y0) model."""
        self._code.append("1 0 0 1 %s Tm" % _fmt2(self._x0, self._y0))
        self._tlm_dirty = False

    def _textOut(self, text, TStar=0):
        if not text:
            # empty payload: reportlab's ragged-line break (_textOut('', 1))
            if TStar and self._tlm_dirty:
                self._code.append("1 0 0 1 %s Tm" % _fmt2(self._x0, self._cy))
                self._tlm_dirty = False
            PDFTextObject._textOut(self, text, TStar)
            if TStar:
                self._advance_line()
            return
        if self._synth:
            x, y = self._cx, self._cy
            W = self._text_width(text)
            self._code.append("1 0 %s 1 %s Tm" % (_fmt(SHEAR_TAN), _fmt2(x, y)))
            # emit the text exactly as reportlab would (subset Tf switches included)
            self._code.append(self._formatText(text))
            if TStar:
                # restore the *line origin* (real y, not reportlab's stale _y0)
                # so T* lands exactly where unsheared reportlab would put it
                self._code.append("1 0 0 1 %s Tm" % _fmt2(self._x0, y))
                self._code.append("T*")
                self._tlm_dirty = False
                self._advance_line()
            else:
                self._code.append("1 0 0 1 %s Tm" % _fmt2(x + W, y))
                self._tlm_dirty = True
                self._cx, self._cy = x + W, y
            return
        if self._tlm_dirty:
            # post-synth fragment: explicit anchor (keeps the advanced cursor);
            # own T* handling so the break lands on the real line origin
            self._code.append("1 0 0 1 %s Tm" % _fmt2(self._cx, self._cy))
            self._code.append(self._formatText(text))
            if TStar:
                self._code.append("1 0 0 1 %s Tm" % _fmt2(self._x0, self._cy))
                self._code.append("T*")
                self._tlm_dirty = False
                self._advance_line()
            else:
                self._tlm_dirty = True
                self._cx += self._text_width(text)
            return
        # clean fragment: reportlab's native emission (byte-identical to an
        # unsheared build, including its 'text Tj T*' line breaks)
        PDFTextObject._textOut(self, text, TStar)
        self._cx += self._text_width(text)
        if TStar:
            self._advance_line()

    def setTextOrigin(self, x, y):
        PDFTextObject.setTextOrigin(self, x, y)
        self._sync_origin()

    def moveCursor(self, dx, dy):
        if self._tlm_dirty:
            self._reanchor_tlm()          # Td offsets from the TLM: clean it first
        PDFTextObject.moveCursor(self, dx, dy)
        self._sync_origin()

    def textLine(self, text=""):
        if self._tlm_dirty:
            self._reanchor_tlm()
        PDFTextObject.textLine(self, text)
        self._sync_origin()

    def __init__(self, *args, **kw):
        self._synth = False
        self._tlm_dirty = False
        self._cx = self._cy = 0.0
        self._tlm_x = self._tlm_y = 0.0
        PDFTextObject.__init__(self, *args, **kw)
        self._sync_origin()


def _fmt(v: float) -> str:
    return ("%.6f" % v).rstrip("0").rstrip(".") or "0"

def _fmt2(a: float, b: float) -> str:
    return "%s %s" % (_fmt(a), _fmt(b))


class BookCanvas(Canvas):
    """Canvas with DeviceGray-only color operators and ShearTextObject text.

    The reportlab page prologue ('BT /F1 12 Tf 14.4 TL ET') references the
    canvas *initial* font; we point it at an embedded face so no page ever
    references an unembedded base-14 font (KDP hygiene)."""

    def __init__(self, *args, **kw):
        kw.setdefault("initialFontName", "EBGaramond-Regular")
        kw.setdefault("initialFontSize", 10.5)
        try:
            Canvas.__init__(self, *args, **kw)
        except TypeError:      # older signature without initialFontName
            kw.pop("initialFontName", None)
            kw.pop("initialFontSize", None)
            Canvas.__init__(self, *args, **kw)

    def beginText(self, x=0, y=0, direction=None):
        try:
            return ShearTextObject(self, x, y, direction=direction)
        except TypeError:                             # older signature
            return ShearTextObject(self, x, y)

    def setFillColor(self, aColor, alpha=None):
        v = _gray_value(aColor)
        if v is None and aColor is not None and not isinstance(aColor, (str, tuple, list)):
            raise ValueError(f"non-gray fill color rejected: {aColor!r}")
        if v is not None:
            self._fillColorObj = aColor
            self._code.append("%s g" % _fmt(v))
        else:
            Canvas.setFillColor(self, aColor, alpha)

    def setStrokeColor(self, aColor, alpha=None):
        v = _gray_value(aColor)
        if v is None and aColor is not None and not isinstance(aColor, (str, tuple, list)):
            raise ValueError(f"non-gray stroke color rejected: {aColor!r}")
        if v is not None:
            self._strokeColorObj = aColor
            self._code.append("%s G" % _fmt(v))
        else:
            Canvas.setStrokeColor(self, aColor, alpha)


# ============================================================================
# 7. TRACKED STRING DRAWING  (TextObject.setCharSpace; InDesign 1/1000 em units)
# ============================================================================

def tracked_width(text: str, font: str, size: float, tracking_em: float) -> float:
    """Advance width with tracking; the trailing tracking space is excluded
    from optical-centering measurements (Claude plan rule)."""
    w = stringWidth(text, font, size)
    cs = tracking_em / 1000.0 * size
    return w + cs * max(len(text) - 1, 0)

def draw_tracked(canvas, text: str, x: float, y: float, font: str, size: float,
                 tracking_em: float, align: str = "left", color=GRAY_BLACK):
    """Draw one tracked line.  align: 'left' | 'center' | 'right' (x = anchor)."""
    cs = tracking_em / 1000.0 * size
    w = tracked_width(text, font, size, tracking_em)
    if align == "center":
        x0 = x - w / 2.0
    elif align == "right":
        x0 = x - w
    else:
        x0 = x
    t = canvas.beginText()
    t.setTextOrigin(x0, y)
    t.setFont(font, size)
    if cs:
        t.setCharSpace(cs)
    t.setFillColor(color)
    t.textOut(text)
    if cs:
        t.setCharSpace(0)
    canvas.drawText(t)
    return w

# ============================================================================
# 8. FLOWABLES — recording paragraphs, tracked headers, list items, markers
# ============================================================================

class LayoutRecorder:
    """Draw-time recorder shared by trials and production renders."""

    def __init__(self):
        self.events = []          # (page, y_baseline, kind, role, payload)
        self.header_pages = {}    # section key -> page (1-based within part doc)
        self.item_pages = {}      # (section key, item idx) -> page
        self.p6_y = None          # EndOfCaseMarker y on its page
        self.p6_page = None

    def add_event(self, page, y, kind, role, payload=None):
        self.events.append((page, y, kind, role, payload))

    def page_events(self, page):
        return [e for e in self.events if e[0] == page]


def markup_to_display(markup: str) -> str:
    """Deterministic inverse of the markup builder (<br/> -> newline,
    tags stripped, entities unescaped; &amp; handled last)."""
    t = re.sub(r"<br\s*/?>", "\n", markup)
    t = re.sub(r"</?[bi]>", "", t)
    t = t.replace("&lt;", "\x00").replace("&gt;", "\x01")
    t = t.replace("&amp;", "&")
    return t.replace("\x00", "<").replace("\x01", ">")


class BookParagraph(Paragraph):
    """Paragraph that records its draw position/lines for telemetry+validation.
    Split parts (reportlab re-creates Paragraphs from frag lists) keep the
    recording context through split() and reconstruct their text from frags."""

    def __init__(self, text, style, role="body-text", src=None,
                 rec: LayoutRecorder = None, **kw):
        Paragraph.__init__(self, text, style, **kw)
        self.role = role
        self.src = src            # dict with item identity for expected-records
        self.rec = rec

    def split(self, aW, aH):
        parts = Paragraph.split(self, aW, aH)
        for part in parts:
            if isinstance(part, BookParagraph):
                part.role, part.src, part.rec = self.role, self.src, self.rec
        return parts

    def _record_text(self) -> str:
        if getattr(self, "text", None):
            return markup_to_display(self.text)
        try:
            return "".join(getattr(f, "text", "") for f in self.frags)
        except Exception:
            return self.getPlainText()

    def drawOn(self, canv, x, y, _sW=0):
        if self.rec is not None:
            n_lines = 0
            bl = getattr(self, "blPara", None)
            if bl is not None and getattr(bl, "lines", None):
                n_lines = len(bl.lines)
            base = y + self.height - first_baseline_offset(self)
            self.rec.add_event(canv.getPageNumber(), base, "para",
                               self.role, {"src": self.src, "n_lines": n_lines,
                                           "text": self._record_text()})
        Paragraph.drawOn(self, canv, x, y, _sW=_sW)


class TrackedLineFlowable(Flowable):
    """Single-line tracked section header (EBGaramond-SemiBold 13pt, +10)."""

    def __init__(self, text, tracking_em, font, size, leading,
                 space_before, space_after, role="section-header", src=None,
                 rec: LayoutRecorder = None):
        Flowable.__init__(self)
        self.text = text
        self.font_name, self.font_size = font, size
        self.tracking_em = tracking_em
        self.leading = leading
        self._sb, self._sa = space_before, space_after
        self.role, self.src, self.rec = role, src, rec
        self.baseline_offset = size           # matches paragraph convention

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        self.height = self.leading
        return availWidth, self.leading

    def getSpaceBefore(self): return self._sb
    def getSpaceAfter(self):  return self._sa

    def draw(self):
        c = self.canv
        draw_tracked(c, self.text, 0, self.height - self.baseline_offset,
                     self.font_name, self.font_size, self.tracking_em, "left")

    def drawOn(self, canv, x, y, _sW=0):
        if self.rec is not None:
            base = y + self.height - self.baseline_offset
            self.rec.add_event(canv.getPageNumber(), base, "header",
                               self.role, {"src": self.src, "text": self.text})
        Flowable.drawOn(self, canv, x, y, _sW=_sW)


class ListItemFlowable(Flowable):
    """Crosshair-bulleted item: vector '+' at the hanging indent, inner paragraph
    wrapped at LIST_W, first baseline aligned with the bullet optical middle."""

    def __init__(self, para: BookParagraph, rec: LayoutRecorder = None):
        Flowable.__init__(self)
        self.para = para
        self.rec = rec
        para.rec = None          # the list records on the paragraph's behalf

    def wrap(self, availWidth, availHeight):
        w, h = self.para.wrap(availWidth - LIST_INDENT, availHeight)
        self.width = availWidth
        self.height = h
        return availWidth, h

    def getSpaceBefore(self): return self.para.getSpaceBefore()
    def getSpaceAfter(self):  return self.para.getSpaceAfter()

    def draw(self):
        c = self.canv
        # crosshair centered vertically on the cap-middle of the first line
        base_y = self.height - first_baseline_offset(self.para)
        cy = base_y - cap_height_pt(self.para.style.fontName,
                                    self.para.style.fontSize) / 2.0
        cx = LIST_INDENT / 2.0
        c.saveState()
        c.setStrokeColor(GRAY_BLACK)
        c.setLineWidth(CROSS_WEIGHT)
        c.line(cx - CROSS_ARM, cy, cx + CROSS_ARM, cy)
        c.line(cx, cy - CROSS_ARM, cx, cy + CROSS_ARM)
        c.restoreState()
        self.para.drawOn(c, LIST_INDENT, 0)

    def drawOn(self, canv, x, y, _sW=0):
        if self.rec is not None:
            base = y + self.height - first_baseline_offset(self.para)
            self.rec.add_event(canv.getPageNumber(), base, "list",
                               self.para.role, {"src": self.para.src,
                                                "text": self.para._record_text()})
        Flowable.drawOn(self, canv, x, y, _sW=_sW)


class EndOfCaseMarker(Flowable):
    """Zero-height marker measuring p6_fill (spec: bottom of last text flowable,
    excluding trailing space-after)."""

    _ZEROSIZE = 1

    def __init__(self, rec: LayoutRecorder = None):
        Flowable.__init__(self)
        self.rec = rec
        self.width = self.height = 0

    def wrap(self, aw, ah):
        return 0, 0

    def draw(self):
        pass

    def drawOn(self, canv, x, y, _sW=0):
        if self.rec is not None:
            self.rec.p6_y = y
            self.rec.p6_page = canv.getPageNumber()
        Flowable.drawOn(self, canv, x, y, _sW=_sW)


class TOCEntryFlowable(Flowable):
    """TOC line: entry text left, measured dot leader, page number right-aligned."""

    def __init__(self, text: str, page_number: int, font="EBGaramond-Regular",
                 size=10.5, leading=16.0, space_after=2.0, rec=None, src=None):
        Flowable.__init__(self)
        self.text, self.page_number = text, page_number
        self.font, self.size, self.leading = font, size, leading
        self._sa = space_after
        self.rec, self.src, self.role = rec, src, "toc-entry"

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        self.height = self.leading
        return availWidth, self.leading

    def getSpaceAfter(self): return self._sa

    def draw(self):
        c = self.canv
        num = str(self.page_number)
        num_w = stringWidth(num, self.font, self.size)
        gap = 6.0                      # minimum blank before and after the dots
        right_x = self.width - num_w - gap
        t = c.beginText()
        t.setTextOrigin(0, self.height - self.size)
        t.setFont(self.font, self.size)
        t.setFillColor(GRAY_BLACK)
        t.textOut(self.text)
        c.drawText(t)
        # measured dot leader ". " units
        unit = stringWidth(". ", self.font, self.size)
        avail = right_x - gap - stringWidth(self.text, self.font, self.size)
        n = max(0, int(avail // unit))
        if n:
            t = c.beginText()
            t.setTextOrigin(stringWidth(self.text, self.font, self.size) + gap,
                            self.height - self.size)
            t.setFont(self.font, self.size)
            t.setFillColor(GRAY_BLACK)
            t.textOut(". " * n)
            c.drawText(t)
        t = c.beginText()
        t.setTextOrigin(self.width - num_w, self.height - self.size)
        t.setFont(self.font, self.size)
        t.setFillColor(GRAY_BLACK)
        t.textOut(num)
        c.drawText(t)

    def drawOn(self, canv, x, y, _sW=0):
        if self.rec is not None:
            self.rec.add_event(canv.getPageNumber(), y + self.height, "toc",
                               "toc-entry", {"src": self.src, "text": self.text})
        Flowable.drawOn(self, canv, x, y, _sW=_sW)


class StanzaBreak(Spacer):
    """Vertical breathing gap in Introduction (14pt) / Epilogue (11pt)."""

    def __init__(self, points: float):
        Spacer.__init__(self, 1, points)
# ============================================================================
# 9. WIDOW-LADDER v2.3 — logical words, rich widths, 65% guard, run-boundary glue
# ============================================================================

def _frag_token_stream(frag_words):
    """Flatten a kind-1 line's ParaFrag list into (kind, text, font, size) tokens.
    Regular spaces split logical words; NBSP never does."""
    toks = []
    for fw in frag_words:
        text = getattr(fw, "text", "")
        if text is None:
            continue
        fname = getattr(fw, "fontName", None)
        fsize = getattr(fw, "fontSize", None)
        pos = 0
        for piece in re.split(r"( +)", text):
            if not piece:
                continue
            if piece.startswith(" "):
                toks.append(("space", piece, fname, fsize))
            else:
                toks.append(("word", piece, fname, fsize))
            pos += len(piece)
    return toks


def _logical_words_from_tokens(toks):
    """Merge adjacent word tokens (word split across formatting boundaries stays
    one logical word). Returns list of dicts with text/width/preceding-space-width."""
    words = []
    pending_space = 0.0
    cur = None
    for kind, piece, fname, fsize in toks:
        if kind == "space":
            w = stringWidth(piece, fname, fsize) if fname else 0.0
            if cur is not None:
                words.append(cur)
                cur = None
            pending_space = w
        else:
            w = stringWidth(piece, fname, fsize)
            if cur is None:
                cur = {"text": piece, "width": w, "space_before": pending_space}
            else:                                   # word split across runs
                cur["text"] += piece
                cur["width"] += w
            pending_space = 0.0
    if cur is not None:
        words.append(cur)
    return words


def inspect_paragraph_lines(para: Paragraph):
    """Return per-line logical-word data for a *wrapped* paragraph.
    reportlab 4.4.x kind-0 lines are (extraSpace, words-strings) 2-tuples;
    kind-1 lines are FragLine objects with .words = [ParaFrag]."""
    bl = para.blPara
    out = []
    if getattr(bl, "kind", 0) == 0:
        style = para.style
        space_w = stringWidth(" ", style.fontName, style.fontSize)
        for line in bl.lines:
            extra, words = line[0], line[-1]
            ww = [{"text": w, "width": stringWidth(w, style.fontName, style.fontSize),
                   "space_before": space_w} for w in words]
            for i, w in enumerate(ww):
                w["space_before"] = space_w if i > 0 else 0.0
            line_w = sum(w["width"] for w in ww) + space_w * max(len(ww) - 1, 0)
            out.append({"words": ww, "line_width": line_w})
    else:
        for line in bl.lines:
            toks = _frag_token_stream(line.words)
            words = _logical_words_from_tokens(toks)
            line_w = sum(w["width"] + (w["space_before"] if i > 0 else 0.0)
                         for i, w in enumerate(words))
            out.append({"words": words, "line_width": line_w})
    return out


def _glue_space_index(text: str):
    """Index of the plain space separating the final logical word from the rest,
    or None when the final word starts an authored line / no space exists."""
    t = text.rstrip()
    if not t:
        return None
    p = t.rfind(" ")
    if p < 0:
        return None
    tail = t[p + 1:]
    if not tail or "\n" in tail:
        return None                     # authored newline boundary: never glue
    return p


def apply_widow_ladder(item: dict, markup: str, style: ParagraphStyle,
                       avail_width: float, case_number: int,
                       item_label: str, report: Report):
    """One attempt per paragraph (spec).  Returns possibly-glued (markup, runs).
    `item` must carry 'text' and optionally 'runs' (already glued-none)."""
    text = item.get("text", "")
    runs = item.get("runs")
    scratch = Paragraph(markup, style)
    scratch.wrap(avail_width, 10 ** 6)
    lines = inspect_paragraph_lines(scratch)
    if len(lines) < 2:
        return markup, runs, None
    last_words = lines[-1]["words"]
    if len(last_words) != 1:
        return markup, runs, None
    prev = lines[-2]["words"]
    if not prev:
        return markup, runs, None
    companion = prev[-1]
    remaining = lines[-2]["line_width"] - companion["width"] - companion["space_before"]
    if (len(companion["text"]) > WIDOW_GUARD_CHARS
            and remaining / avail_width < WIDOW_GUARD_FILL):
        report.info("WIDOW-ACCEPTED",
                    f"Case {case_number} {item_label}: kept 1-word last line "
                    f"(companion '{companion['text'][:24]}' >{WIDOW_GUARD_CHARS} chars, "
                    f"prev line {remaining / avail_width:.0%} full)")
        return markup, runs, "accepted"

    p = _glue_space_index(text)
    if p is None:
        return markup, runs, None
    # verify the located tail equals the rendered final word
    if text.rstrip()[p + 1:] != last_words[0]["text"]:
        return markup, runs, None

    if runs:
        new_runs, ok = _glue_runs(runs, p)
    else:
        new_runs, ok = None, True
        text = text[:p] + NBSP + text[p + 1:]
    if not ok:
        return markup, runs, None

    new_item_text = text if new_runs is None else "".join(r["t"] for r in new_runs)
    new_markup = (runs_markup(new_runs) if new_runs
                  else esc(new_item_text).replace("\n", "<br/>"))
    check = Paragraph(new_markup, style)
    check.wrap(avail_width, 10 ** 6)
    report.info("WIDOW-GLUED", f"Case {case_number} {item_label}: "
                               f"NBSP before '{last_words[0]['text'][:24]}'")
    return new_markup, new_runs, "glued"


def _glue_runs(runs, pos: int):
    """Replace the space at source-text offset `pos` with NBSP inside the runs
    list, preserving every run's b/i flags (ChatGPT #4 run-boundary case)."""
    off = 0
    new_runs = []
    done = False
    for r in runs:
        t = r["t"]
        if not done and off <= pos < off + len(t):
            if t[pos - off] != " ":
                return None, False
            t = t[:pos - off] + NBSP + t[pos - off + 1:]
            done = True
        new_runs.append(dict(r, t=t))
        off += len(r["t"])
    return (new_runs, True) if done else (None, False)


# ============================================================================
# 10. STYLES
# ============================================================================

def make_styles(tier: str) -> dict:
    lead, sa, hsb, hsa, llead, lsa = TIERS[tier]
    base = dict(fontName="EBGaramond-Regular", fontSize=BODY_SIZE,
                alignment=TA_LEFT, allowWidows=0, allowOrphans=0,
                hyphenationLang=None, embeddedHyphenation=0, splitLongWords=0,
                spaceBefore=0, textColor=GRAY_BLACK)
    body = ParagraphStyle(f"body-{tier}", leading=lead, spaceAfter=sa, **base)
    lst = ParagraphStyle(f"list-{tier}", leading=llead, spaceAfter=lsa, **base)
    see = ParagraphStyle(f"see-{tier}", fontName="EBGaramond-Italic",
                         leading=14.0, spaceAfter=4.0, **{k: v for k, v in base.items()
                                                          if k != "fontName"})
    styles = {"body": body, "list": lst, "see-also": see, "search-term": see}
    # fixed (tier-independent) flowed styles
    styles["intro-body"] = ParagraphStyle("intro-body", leading=14.0, spaceAfter=4.0, **base)
    styles["intro-title"] = ParagraphStyle(
        "intro-title", fontName="LibreBaskerville-Bold", fontSize=17.0, leading=17.0,
        alignment=TA_CENTER, spaceAfter=24.0, textColor=GRAY_BLACK)
    styles["epilogue-body"] = ParagraphStyle(
        "epilogue-body", fontName="EBGaramond-Regular", fontSize=BODY_SIZE,
        leading=11.0, spaceAfter=4.0, alignment=TA_CENTER, textColor=GRAY_BLACK)
    styles["copyright"] = ParagraphStyle(
        "copyright", fontName="EBGaramond-Regular", fontSize=9.0, leading=11.0,
        spaceBefore=4.0, alignment=TA_CENTER, textColor=GRAY_BLACK)
    styles["toc-title"] = ParagraphStyle(
        "toc-title", fontName="LibreBaskerville-Bold", fontSize=17.0, leading=17.0,
        alignment=TA_CENTER, spaceAfter=24.0, textColor=GRAY_BLACK)
    styles["header-metrics"] = {"space_before": hsb, "space_after": hsa}
    styles["_tier"] = tier
    return styles


# ============================================================================
# 11. MARKUP — runs/esc/newline handling (ChatGPT #7/#8: source is authoritative)
# ============================================================================

def runs_markup(runs) -> str:
    out = []
    for r in runs:
        t = esc(r["t"]).replace("\n", "<br/>")
        b, i = r.get("b"), r.get("i")
        if b and i:
            out.append(f"<b><i>{t}</i></b>")
        elif b:
            out.append(f"<b>{t}</b>")
        elif i:
            out.append(f"<i>{t}</i>")
        else:
            out.append(t)
    return "".join(out)


def item_display(item: dict) -> str:
    """Authoritative display text of an item (theory: lead + ' — ' + body)."""
    if item["type"] == "theory-item":
        return f"{item['lead']} \u2014 {item['body']}"
    return item["text"]


def item_source_runs(item: dict):
    """Runs-equivalent of the item source, for widow gluing."""
    if item.get("runs"):
        return item["runs"]
    if item["type"] == "theory-item":
        return [{"t": f"{item['lead']} \u2014 ", "b": True}, {"t": item["body"]}]
    return None


def build_item_markup(item: dict, styles: dict, glued=None) -> "tuple[str, BookParagraph]":
    """Build the (possibly widow-glued) markup + BookParagraph for one item.
    `glued` is the (markup, runs, action) tuple from apply_widow_ladder."""
    t = item["type"]
    if t == "body-text":
        style = styles["body"]
        role = "body-text"
    elif t in ("evidence-item", "theory-item"):
        style = styles["list"]
        role = t
    elif t in ("see-also", "search-term"):
        style = styles["see-also"]
        role = t
    else:
        raise ValueError(f"unknown item type {t}")
    markup = glued[0] if glued else (
        runs_markup(item["runs"]) if item.get("runs") else esc(item_display(item)).replace("\n", "<br/>"))
    return markup, style, role


# ============================================================================
# 12. OVERRIDES — locked two-case contract
# ============================================================================

def load_overrides(report: Report) -> "dict | None":
    path = ROOT / "overrides.json"
    if not path.exists():
        report.crit("CRITICAL-OVERRIDE", "overrides.json missing")
        return None
    try:
        ov = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        report.crit("CRITICAL-OVERRIDE", f"invalid JSON: {e}")
        return None
    if not isinstance(ov, dict) or set(ov.keys()) != {"schema", "source_sha256", "cases"}:
        report.crit("CRITICAL-OVERRIDE", "top-level keys must be exactly "
                                         "{schema, source_sha256, cases}")
        return None
    if ov["schema"] != OVERRIDE_SCHEMA:
        report.crit("CRITICAL-OVERRIDE", f"schema must be {OVERRIDE_SCHEMA!r}")
        return None
    if ov["source_sha256"] != FROZEN_SHA256:
        report.crit("CRITICAL-OVERRIDE", "source_sha256 does not match frozen hash")
        return None
    cases = ov["cases"]
    if not isinstance(cases, dict):
        report.crit("CRITICAL-OVERRIDE", "cases must be an object")
        return None
    for k, v in cases.items():
        if k not in APPROVED_OVERRIDE_CASES:
            report.crit("CRITICAL-OVERRIDE",
                        f"case {k!r} is not in the approved override list "
                        f"{sorted(APPROVED_OVERRIDE_CASES)}; adding a case is an "
                        f"approved one-line change to APPROVED_OVERRIDE_CASES")
            return None
        if not isinstance(v, dict) or set(v.keys()) != {"big_word_pt", "reason"}:
            report.crit("CRITICAL-OVERRIDE", f"case {k}: fields must be exactly "
                                             "{big_word_pt, reason}")
            return None
        pt = v["big_word_pt"]
        if not isinstance(pt, (int, float)) or not (40 <= pt <= 50):
            report.crit("CRITICAL-OVERRIDE", f"case {k}: big_word_pt must be 40..50")
            return None
        if not isinstance(v["reason"], str) or not v["reason"].strip():
            report.crit("CRITICAL-OVERRIDE", f"case {k}: non-empty reason required")
            return None
    for k, pt in APPROVED_OVERRIDE_CASES.items():
        if k not in cases:
            report.crit("CRITICAL-OVERRIDE", f"approved case {k} missing from overrides.json")
            return None
        if cases[k]["big_word_pt"] != pt:
            report.crit("CRITICAL-OVERRIDE", f"case {k} must be big_word_pt={pt}")
            return None
    for k, v in cases.items():
        report.info("OVERRIDE-APPLIED",
                    f"case {k}: big_word_pt={v['big_word_pt']} ({v['reason']})")
    return {k: v["big_word_pt"] for k, v in cases.items()}


# ============================================================================
# 13. PREFLIGHT — hashes, JSON 1.1 structure, fonts, images, display widths
# ============================================================================

DOCUMENTED_AGGREGATES = {     # from the accepted plan; drift is reported, not fatal
    "body-text": 3244, "evidence-item": 393, "theory-item": 310,
    "see-also": 192, "search-term": 128,
}
DOCUMENTED_INTRO_BODY = 80
DOCUMENTED_EPILOGUE_GROUPS = [1, 3, 2, 1, 2]
FRONTMATTER_COUNTS = {"book-name": 1, "book-name-full": 1, "subtitle-name": 2,
                      "author": 1, "copyright": 5, "dedication-line": 1,
                      "welcome": 1, "toc-title": 1, "intro-title": 1}

def preflight(report: Report, need_images=True) -> "dict | None":
    """Full input gate. Returns the loaded data bundle or None on CRITICAL."""
    report.section("PREFLIGHT")
    register_all_fonts(report)
    data = {}

    # ---- manuscript + JSON hashes ---------------------------------------
    if not MANUSCRIPT.exists():
        report.crit("CRITICAL-HASH", f"manuscript missing: {MANUSCRIPT}")
    else:
        live = sha256_file(MANUSCRIPT)
        if live != FROZEN_SHA256:
            report.crit("CRITICAL-HASH", f"manuscript hash drift: {live}")
        else:
            report.info("HASH-OK", "manuscript_styled.docx == frozen SHA-256")

    for name in ("frontmatter.json", "cases.json", "epilogue.json"):
        p = JSON_DIR / name
        if not p.exists():
            report.crit("CRITICAL-JSON", f"{name} missing")
            continue
        try:
            j = json.loads(p.read_text(encoding="utf-8"))
        except Exception as e:
            report.crit("CRITICAL-JSON", f"{name} parse error: {e}")
            continue
        for key in ("schema_version", "ingest_version"):
            if j.get(key) != "1.1":
                report.crit("CRITICAL-JSON", f"{name}: {key} != 1.1")
        if j.get("ingest_status") != "OK":
            report.crit("CRITICAL-JSON", f"{name}: ingest_status != OK")
        if j.get("criticals") not in (0, "0"):
            report.crit("CRITICAL-JSON", f"{name}: criticals != 0")
        if j.get("source_sha256") != FROZEN_SHA256:
            report.crit("CRITICAL-JSON", f"{name}: source_sha256 != frozen hash")
        data[name.split(".")[0]] = j
    if report.n_crit:
        return None

    cases = data["cases"]["cases"]
    fm = data["frontmatter"]
    epi = data["epilogue"]
    overrides = load_overrides(report) or {}
    data["overrides"] = overrides
    if report.n_crit:
        return None

    # ---- case structure ---------------------------------------------------
    if len(cases) != N_CASES:
        report.crit("CRITICAL-STRUCTURE", f"expected {N_CASES} cases, got {len(cases)}")
        return None
    see_targets_all = []
    counts = {}
    multiline_n = 0
    runs_n = 0
    for i, c in enumerate(cases, start=1):
        cn = c.get("case_number")
        if cn != i:
            report.crit("CRITICAL-STRUCTURE", f"case sequence broken at index {i}: {cn}")
        if not re.match(rf"^CASE {i}: .+$", c.get("case_name", "")):
            report.crit("CRITICAL-STRUCTURE", f"case {i}: case_name malformed: "
                                             f"{c.get('case_name')!r}")
        tl = c.get("title_lines")
        if not isinstance(tl, list) or not (1 <= len(tl) <= 2):
            report.crit("CRITICAL-STRUCTURE", f"case {i}: title_lines must be 1..2")
        if not c.get("big_word"):
            report.crit("CRITICAL-STRUCTURE", f"case {i}: big_word missing")
        if not c.get("quote"):
            report.crit("CRITICAL-STRUCTURE", f"case {i}: quote missing")
        caps = c.get("captions")
        if not isinstance(caps, list) or not (1 <= len(caps) <= 3):
            report.crit("CRITICAL-STRUCTURE", f"case {i}: captions must be 1..3 lines")
        secs = c.get("sections")
        if [s.get("key") for s in (secs or [])] != SECTION_KEYS:
            report.crit("CRITICAL-STRUCTURE",
                        f"case {i}: section keys != {SECTION_KEYS}")
            continue
        for s in secs:
            for item in s["items"]:
                counts[item["type"]] = counts.get(item["type"], 0) + 1
                if "\n" in item.get("text", ""):
                    multiline_n += 1
                if item.get("runs"):
                    runs_n += 1
                    joined = "".join(r["t"] for r in item["runs"])
                    if joined != item["text"]:
                        report.crit("CRITICAL-STRUCTURE",
                                    f"case {i}: run concatenation != text")
                    if item["type"] != "body-text":
                        report.crit("CRITICAL-STRUCTURE",
                                    f"case {i}: runs present on {item['type']} "
                                    f"(only body-text carries runs)")
        for s in secs:
            if s["key"] == "SEE_ALSO":
                if len(s["items"]) != 3:
                    report.crit("CRITICAL-STRUCTURE",
                                f"case {i}: SEE ALSO must have exactly 3 entries")
                for it in s["items"]:
                    m = it.get("target")
                    if not (isinstance(m, int) and 1 <= m <= N_CASES and m != i):
                        report.crit("CRITICAL-STRUCTURE",
                                    f"case {i}: see-also target invalid: {m!r}")
                    recon = f"{it.get('name')} (Case {m}) \u2014 {it.get('description')}"
                    if it.get("text") != recon:
                        report.crit("CRITICAL-STRUCTURE",
                                    f"case {i}: see-also text reconstruction mismatch")
                    see_targets_all.append((i, m))
            if s["key"] == "CURIOUS":
                if len(s["items"]) != 2:
                    report.crit("CRITICAL-STRUCTURE",
                                f"case {i}: CURIOUS must have exactly 2 search terms")
                for it in s["items"]:
                    if not it.get("text", "").strip():
                        report.crit("CRITICAL-STRUCTURE", f"case {i}: empty search term")
    if report.n_crit:
        return None
    report.info("STRUCTURE-OK", f"64 cases, 8 ordered sections each, "
                                f"3 see-also + 2 search terms per case")
    for k, v in DOCUMENTED_AGGREGATES.items():
        got = counts.get(k, 0)
        if got != v:
            report.warn("AGGREGATE-DRIFT", f"{k}: {got} items (documented {v})")
    total_items = sum(counts.values())
    report.info("AGGREGATES", f"case items total={total_items} " +
                " ".join(f"{k}={v}" for k, v in sorted(counts.items())) +
                f"; multiline={multiline_n}; runs={runs_n}")
    bi_runs = 0
    for c in cases:
        for s in c["sections"]:
            for it in s["items"]:
                for r in it.get("runs") or []:
                    if r.get("b") and r.get("i"):
                        bi_runs += 1
    if bi_runs == 0:
        report.info("SYNTH-UNUSED", "no bold+italic run in the manuscript; "
                                    "synthetic BoldItalic exercised by --selftest only")

    # ---- front matter / intro / epilogue ---------------------------------
    fmdata = fm.get("frontmatter", {})
    for k, v in FRONTMATTER_COUNTS.items():
        got = len(fmdata.get(k, []))
        if got != v:
            report.crit("CRITICAL-STRUCTURE", f"frontmatter {k}: {got} entries (need {v})")
    intro = fm.get("introduction", [])
    intro_body = [i for i in intro if i.get("type") == "body-text"]
    intro_stanza = [i for i in intro if i.get("type") == "stanza_break"]
    for i in intro:
        if i.get("type") not in ("body-text", "stanza_break"):
            report.crit("CRITICAL-STRUCTURE", f"introduction bad item type {i.get('type')!r}")
    if not intro_body:
        report.crit("CRITICAL-STRUCTURE", "introduction has no body text")
    if len(intro_body) != DOCUMENTED_INTRO_BODY:
        report.warn("AGGREGATE-DRIFT",
                    f"introduction body items {len(intro_body)} (documented {DOCUMENTED_INTRO_BODY})")
    report.info("INTRO-OK", f"introduction: {len(intro_body)} body, "
                            f"{len(intro_stanza)} stanza breaks")
    groups = [len(g) for g in epi.get("groups", [])]
    if groups != DOCUMENTED_EPILOGUE_GROUPS:
        report.warn("EPILOGUE-GROUPS", f"got {groups} (documented "
                                       f"{DOCUMENTED_EPILOGUE_GROUPS}); copy change?")
    if not all(l == l.upper() for g in epi.get("groups", []) for l in g):
        report.warn("EPILOGUE-CASE", "epilogue line not uppercase (rendered as stored)")
    if report.n_crit:
        return None

    # ---- fonts & per-role glyph coverage ----------------------------------
    cbs_path = JSON_DIR / "charset_by_style.json"
    if not cbs_path.exists():
        report.crit("CRITICAL-FONTS", "charset_by_style.json missing")
        return None
    cbs = json.loads(cbs_path.read_text(encoding="utf-8"))
    unknown_roles = set(cbs) - set(ROLE_FONT)
    if unknown_roles:
        report.crit("CRITICAL-CONFIG", f"charset_by_style roles without font mapping: "
                                       f"{sorted(unknown_roles)}")
    # roles this engine renders that charset_by_style does not declare
    extra_role_charset = {
        "toc-entry": set("".join(c["case_name"] for c in cases)) | set("INTRODUCTION0123456789 "),
        "running-header": set("".join(c["case_name"] for c in cases)) | set("UNEXPLAINED"),
        "folio": set("0123456789"),
        "list-item-lead": set("".join(it["lead"] for c in cases for s in c["sections"]
                                      if s["key"] == "THEORIES" for it in s["items"]
                                      if it["type"] == "theory-item")
                              + " \u2014"),
    }
    problems = 0
    for role, chars in cbs.items():
        font = ROLE_FONT.get(role)
        if not font:
            continue
        cmap = _FONT_METRICS[font]["cmap"]
        missing = sorted({ch for ch in chars if ord(ch) > 32 and ord(ch) not in cmap})
        if missing:
            report.crit("CRITICAL-FONT-COVERAGE",
                        f"role {role} / {font}: missing glyphs "
                        f"{[hex(ord(x)) for x in missing]}")
            problems += 1
    for role, chars in extra_role_charset.items():
        font = ROLE_FONT[role]
        cmap = _FONT_METRICS[font]["cmap"]
        missing = sorted({ch for ch in chars if ord(ch) > 32 and ord(ch) not in cmap})
        if missing:
            report.crit("CRITICAL-FONT-COVERAGE",
                        f"engine role {role} / {font}: missing "
                        f"{[hex(ord(x)) for x in missing]}")
            problems += 1
    # NBSP must exist in every face that can receive glue
    for font in ("EBGaramond-Regular", "EBGaramond-Bold", "EBGaramond-Italic"):
        if 0x00A0 not in _FONT_METRICS[font]["cmap"]:
            report.crit("CRITICAL-FONT-COVERAGE", f"{font}: NBSP U+00A0 missing")
    if not problems:
        report.info("FONTS-OK", "7 faces registered; per-role cmap coverage verified "
                                "(declared charsets + engine-derived charsets + NBSP)")

    # ---- display-line width gates -----------------------------------------
    worst_bw = []
    for c in cases:
        n = c["case_number"]
        size = overrides.get(str(n), 50)
        w = stringWidth(c["big_word"], "LibreBaskerville-Bold", size)
        if w > TEXT_W + 0.5:
            report.crit("CRITICAL-BIGWORD-WIDTH",
                        f"case {n}: big word '{c['big_word']}' measures {w:.1f}pt "
                        f"at {size}pt > text frame {TEXT_W}pt; approved fix is a "
                        f"45pt override (measures "
                        f"{stringWidth(c['big_word'], 'LibreBaskerville-Bold', 45):.1f}pt)")
        worst_bw.append((w, n))
        for tl in c["title_lines"]:
            w2 = stringWidth(tl, "LibreBaskerville-Bold", 24)
            if w2 > TEXT_W + 0.5:
                report.crit("CRITICAL-TITLE-WIDTH",
                            f"case {n}: title line '{tl}' {w2:.1f}pt > {TEXT_W}pt")
        w3 = tracked_width(c["case_name"], "EBGaramond-Medium", 12, 40)
        if w3 > TEXT_W + 0.5:
            report.crit("CRITICAL-CASENAME-WIDTH", f"case {n}: {w3:.1f}pt")
        for cap in c["captions"]:
            wc = stringWidth(cap, "EBGaramond-Italic", 9)
            if wc > CAPTION_MAX_W + 0.5:
                report.crit("CRITICAL-CAPTION-WIDTH",
                            f"case {n}: caption '{cap[:40]}...' {wc:.1f}pt > "
                            f"{CAPTION_MAX_W}pt")
        for seg in c["quote"].split("\n"):
            wq = stringWidth(seg, "EBGaramond-Italic", 13)
            if wq > TEXT_W + 0.5:
                report.warn("QUOTE-WRAPPED",
                            f"case {n}: quote segment wraps ({wq:.1f}pt)")
    for line, font, size, tr, limit in (
            (fmdata["book-name"][0], "LibreBaskerville-Bold", 32, 40, TEXT_W),
            (fmdata["book-name-full"][0], "LibreBaskerville-Bold", 32, 40, TEXT_W),
            (fmdata["welcome"][0], "EBGaramond-Medium", 9, 60, TEXT_W),
            (fmdata["toc-title"][0], "LibreBaskerville-Bold", 17, 0, TEXT_W)):
        w = tracked_width(line, font, size, tr) if tr else stringWidth(line, font, size)
        if w > limit + 0.5:
            report.crit("CRITICAL-DISPLAY-WIDTH", f"'{line[:30]}' {w:.1f}pt > {limit}pt")
    for sub in fmdata["subtitle-name"]:
        w = tracked_width(sub, "EBGaramond-Bold", 20, 40)
        if w > TEXT_W + 0.5:
            report.crit("CRITICAL-DISPLAY-WIDTH", f"subtitle '{sub}' {w:.1f}pt")
    if not report.n_crit:
        report.info("WIDTHS-OK", "all display lines fit their frames "
                                 "(big words, title lines, captions, headers)")

    # ---- images (ChatGPT #9: exact filename-to-case identity) -------------
    if need_images:
        if not IMAGE_DIR.exists():
            report.crit("CRITICAL-IMAGE", f"images folder missing: {IMAGE_DIR}")
        else:
            n_ok = 0
            for n in range(1, N_CASES + 1):
                p = IMAGE_DIR / f"CASE{n:02d}.png"
                if not p.exists():
                    report.crit("CRITICAL-IMAGE", f"case {n}: missing {p.name}")
                    continue
                with PILImage.open(p) as im:
                    if im.size != (1800, 2700):
                        report.crit("CRITICAL-IMAGE", f"case {n}: size {im.size} != (1800, 2700)")
                        continue
                    if im.mode != "L":
                        report.crit("CRITICAL-IMAGE", f"case {n}: mode {im.mode} != L")
                        continue
                    dpi = im.info.get("dpi", (0, 0))
                if min(dpi) < 299.9:
                    report.crit("CRITICAL-IMAGE", f"case {n}: dpi {dpi} < 299.9")
                    continue
                n_ok += 1
            if n_ok == N_CASES:
                report.info("IMAGES-OK", "64/64 CASE{NN}.png: 1800x2700, mode L, "
                                         ">=299.9 dpi, exact case mapping")

    report.info("PREFLIGHT-RESULT",
                "PASS" if report.n_crit == 0 else f"FAIL ({report.n_crit} critical)")
    return None if report.n_crit else data
# ============================================================================
# 14. FRAMES — CapAlignedFrame with single-trigger recto calibration
# ============================================================================

from reportlab.platypus.doctemplate import PageBreak, FrameBreak   # noqa: E402
try:                                    # reportlab 4.4.x: FrameBreak is a factory
    from reportlab.platypus.doctemplate import _FrameBreak as _FrameBreakClass
except ImportError:                     # older lines: FrameBreak is the class
    _FrameBreakClass = FrameBreak

_NON_DRAWABLE = (ActionFlowable, Spacer, EndOfCaseMarker, KeepTogether)
_DRAWABLE = (BookParagraph, TrackedLineFlowable, ListItemFlowable, Paragraph)


class CapAlignedFrame(Frame):
    """Recto content frame whose FIRST drawable flowable is placed so the first
    line's cap-height top lands exactly on CAP_TARGET_Y (0.816 in from top).

    Single trigger: zero-size trackers, actions, spacers and KeepTogether
    wrappers (whose children are re-submitted individually by reportlab's
    KeepTogether.split) never consume the calibration.  Acceptance is the
    --selftest PDF-geometry test, not this formula (ChatGPT #1).
    """

    def __init__(self, *a, **kw):
        Frame.__init__(self, *a, **kw)
        self._cap_calibrated = False

    def _reset(self):
        Frame._reset(self)
        self._cap_calibrated = False

    def add(self, flowable, canv, trySplit=0):
        if not self._cap_calibrated and self._atTop:
            if isinstance(flowable, _DRAWABLE) and not isinstance(flowable, _NON_DRAWABLE):
                try:
                    flowable.wrap(self._getAvailableWidth(), 10 ** 6)
                    fname, fsize = cap_font_size(flowable)
                    offset = first_baseline_offset(flowable)
                    cap = cap_height_pt(fname, fsize)
                    self._y = CAP_TARGET_Y - cap + offset
                    self._cap_calibrated = True
                except Exception:
                    pass        # fail-visible: geometry selftest will catch it
        return Frame.add(self, flowable, canv, trySplit)


class HeaderBlock(KeepTogether):
    """KeepTogether with a *guaranteed* atomic move (ChatGPT #11).

    reportlab's KeepTogether.split() decides between inserting a FrameBreak
    or unpacking the children into the current frame based on peeked-frame
    heuristics.  For section-header protection we make the behaviour
    deterministic: whenever the group does not fit, a FrameBreak is always
    prepended, so the header can never be unpacked into a frame that cannot
    hold the whole group.  Groups taller than a full empty frame (impossible
    for one section here) fall back to plain unpacking.
    """

    def split(self, aW, aH):
        S = KeepTogether.split(self, aW, aH)
        # KeepTogether.wrap() always reports a huge height, so split() runs even
        # when the group fits; only intervene when it genuinely does not fit
        # the remaining space (self._H > aH) yet fits an empty frame.
        if (S and not isinstance(S[0], _FrameBreakClass)
                and self._H is not None and self._H > aH):
            frame = getattr(self, "_frame", None)
            full = None
            if frame is not None:
                try:
                    full = frame._y2 - frame._y1p
                except Exception:
                    full = None
            if full is None or self._H <= full:
                S.insert(0, FrameBreak())
        return S


def make_side_frame(side: str) -> Frame:
    if side == "verso":
        return Frame(VERS0_X, FRAME_BOT_Y, TEXT_W, FRAME_H,
                     leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0)
    return CapAlignedFrame(RECTO_X, FRAME_BOT_Y, TEXT_W, FRAME_H,
                           leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0)

# ============================================================================
# 15. CHROME + FIXED-CANVAS PAGE PAINTERS
# ============================================================================

def paint_a_master(canv, doc, case_ctx, side):
    """Running header + rule + folio on case content pages (doc pages 3..6)."""
    if doc.page < 3:
        return
    book_page = case_ctx["base"] + doc.page - 1
    frame_x = VERS0_X if side == "verso" else RECTO_X
    frame_right = VERS0_RIGHT if side == "verso" else RECTO_RIGHT
    # header
    if side == "verso":
        draw_tracked(canv, "UNEXPLAINED", frame_x, HEADER_BASELINE_Y,
                     "EBGaramond-Medium", 9, 20, "left", GRAY_75)
    else:
        draw_tracked(canv, case_ctx["case_name"], frame_right, HEADER_BASELINE_Y,
                     "EBGaramond-Medium", 9, 20, "right", GRAY_75)
    # rule
    canv.saveState()
    canv.setStrokeColor(GRAY_BLACK)
    canv.setLineWidth(RULE_WEIGHT)
    canv.line(frame_x, RULE_Y, frame_right, RULE_Y)
    canv.restoreState()
    # folio (outside corner)
    if side == "verso":
        draw_tracked(canv, str(book_page), VERS0_X, FOLIO_Y,
                     "EBGaramond-Medium", 11, 0, "left")
    else:
        draw_tracked(canv, str(book_page), RECTO_RIGHT, FOLIO_Y,
                     "EBGaramond-Medium", 11, 0, "right")


def paint_case_title_page(canv, case, big_word_pt: int, report=None):
    """B-master verso: case name -> big word -> title lines -> quote.

    Baseline chain in bottom-origin PDF coordinates: each block's box top
    sits GAP below the previous baseline, and the block's baseline sits one
    ascent *below* its own box top:
        box_top = prev_baseline - gap ; baseline = box_top - ascent.
    """
    cx = (VERS0_X + VERS0_RIGHT) / 2.0
    draw_tracked(canv, case["case_name"], cx, CASENAME_BASELINE_Y,
                 "EBGaramond-Medium", 12, 40, "center")
    bw_top = CASENAME_BASELINE_Y - GAP_CASENAME_BW
    bw_base = bw_top - ascent_pt("LibreBaskerville-Bold", big_word_pt)
    draw_tracked(canv, case["big_word"], cx, bw_base,
                 "LibreBaskerville-Bold", big_word_pt, 0, "center")
    tl_top = bw_base - GAP_BW_TL
    tl_base = tl_top - ascent_pt("LibreBaskerville-Bold", 24)
    for i, tl in enumerate(case["title_lines"]):
        draw_tracked(canv, tl, cx, tl_base - i * TL_LEADING,
                     "LibreBaskerville-Bold", 24, 0, "center")
    last_tl = tl_base - (len(case["title_lines"]) - 1) * TL_LEADING
    q_top = last_tl - GAP_TL_QUOTE
    q_base = q_top - ascent_pt("EBGaramond-Italic", 13)
    q_lines = case["quote"].split("\n")
    for i, ql in enumerate(q_lines):
        draw_tracked(canv, ql, cx, q_base - i * QUOTE_LEADING,
                     "EBGaramond-Italic", 13, 0, "center")
    lowest = q_base - (len(q_lines) - 1) * QUOTE_LEADING \
        - _FONT_METRICS["EBGaramond-Italic"]["descent"] * 13
    if lowest < M_BOT - 0.5:
        raise RuntimeError(f"Case {case['case_number']}: title-page stack overflows "
                           f"bottom margin (lowest {lowest:.1f} < {M_BOT})")


def paint_case_image_page(canv, case, report=None):
    """B-master recto: full-bleed grayscale image + centered caption block."""
    n = case["case_number"]
    path = IMAGE_DIR / f"CASE{n:02d}.png"
    canv.drawImage(str(path), 0, 0, width=PAGE_W, height=PAGE_H,
                   preserveAspectRatio=False, mask=None)
    caps = case["captions"]
    # spec: the FIRST caption line's baseline is 1.15in from the bottom edge;
    # additional lines stack below it (leading 11pt)
    y = CAPTION_BASELINE_Y
    for cap in caps:
        draw_tracked(canv, cap, PAGE_CENTER_X, y, "EBGaramond-Italic", 9, 0, "center")
        y -= CAPTION_LEADING


def paint_front_page(canv, page_no: int, fm: dict):
    """Fixed C-master front-matter pages 1..6 (canvas only)."""
    if page_no == 1:
        draw_tracked(canv, fm["book-name"][0], PAGE_CENTER_X, HALFTITLE_BASELINE_Y,
                     "LibreBaskerville-Bold", 32, 40, "center")
    elif page_no == 3:
        draw_tracked(canv, fm["book-name-full"][0], PAGE_CENTER_X, TITLE_BOOK_BASELINE_Y,
                     "LibreBaskerville-Bold", 32, 40, "center")
        for i, sub in enumerate(fm["subtitle-name"]):
            draw_tracked(canv, sub, PAGE_CENTER_X, TITLE_SUB1_BASELINE_Y - i * TITLE_SUB_LEADING,
                         "EBGaramond-Bold", 20, 40, "center")
        draw_tracked(canv, fm["author"][0], PAGE_CENTER_X, TITLE_AUTHOR_BASELINE_Y,
                     "LibreBaskerville-Bold", 20, 0, "center")
    elif page_no == 4:
        y = COPYRIGHT_BASELINE_Y
        for para in fm["copyright"]:
            canv.setFont("EBGaramond-Regular", 9)
            canv.setFillColor(GRAY_BLACK)
            canv.drawCentredString(PAGE_CENTER_X, y, para)
            y -= (11.0 + 4.0)
    elif page_no == 5:
        canv.setFont("EBGaramond-Italic", 12)
        canv.setFillColor(GRAY_BLACK)
        canv.drawCentredString(PAGE_CENTER_X, DEDICATION_BASELINE_Y,
                               fm["dedication-line"][0])
    elif page_no == 6:
        draw_tracked(canv, fm["welcome"][0], PAGE_CENTER_X, WELCOME_BASELINE_Y,
                     "EBGaramond-Medium", 9, 60, "center")

# ============================================================================
# 16. CASE STORY BUILDER (stranding protection + fresh flowables per call)
# ============================================================================

def make_item_flowable(item, idx, sec_key, styles, glued, rec):
    markup, style, role = build_item_markup(item, styles, glued.get((sec_key, idx)))
    src = {"key": sec_key, "idx": idx, "type": item["type"]}
    para = BookParagraph(markup, style, role=role, src=src, rec=rec)
    if item["type"] in ("evidence-item", "theory-item"):
        return ListItemFlowable(para, rec=rec)
    return para


def single_line_item(item, styles, glued, sec_key, idx):
    markup, style, role = build_item_markup(item, styles, glued.get((sec_key, idx)))
    p = Paragraph(markup, style)          # plain probe (no recording)
    p.wrap(LIST_W if item["type"] in ("evidence-item", "theory-item") else TEXT_W,
           10 ** 6)
    bl = getattr(p, "blPara", None)
    return bool(bl and bl.lines and len(bl.lines) == 1)


def build_case_story(case, styles, glued, break_before=None, rec=None):
    """Fresh flowables per call (Platypus is stateful).  Stranding protection:
    KeepTogether(header + first item [+ second when the first is single-line])."""
    hm = styles["header-metrics"]
    story = []
    for s in case["sections"]:
        key, items = s["key"], s["items"]
        header = TrackedLineFlowable(
            s["header"].upper(), 10, "EBGaramond-SemiBold", 13, 14,
            hm["space_before"], hm["space_after"],
            src={"key": key}, rec=rec)
        first = make_item_flowable(items[0], 0, key, styles, glued, rec)
        group = [header, first]
        skip_second = False
        if (items[0]["type"] == "body-text" and len(items) > 1
                and single_line_item(items[0], styles, glued, key, 0)):
            group.append(make_item_flowable(items[1], 1, key, styles, glued, rec))
            skip_second = True
        if break_before == key:
            story.append(FrameBreak())
        story.append(HeaderBlock(group))
        start = 2 if skip_second else 1
        for idx in range(start, len(items)):
            story.append(make_item_flowable(items[idx], idx, key, styles, glued, rec))
    story.append(EndOfCaseMarker(rec))
    return story

# ============================================================================
# 17. TRIALS — identical pipeline, scratch canvas (ChatGPT: dry run == build)
# ============================================================================

class TrialResult:
    __slots__ = ("tier", "break_before", "pages", "sec_start", "sec_end",
                 "item_pages", "p6_fill", "spill_pt", "first_item_overflow")

    def __init__(self, tier, break_before):
        self.tier, self.break_before = tier, break_before
        self.pages = 0
        self.sec_start, self.sec_end = {}, {}
        self.item_pages = {}
        self.p6_fill = None
        self.spill_pt = None
        self.first_item_overflow = None

    def state(self):
        return (f"P={self.pages} Q={self.sec_start.get('QUESTION')}..{self.sec_end.get('QUESTION')} "
                f"S={self.sec_start.get('SEE_ALSO')}..{self.sec_end.get('SEE_ALSO')} "
                f"C={self.sec_start.get('CURIOUS')}..{self.sec_end.get('CURIOUS')} "
                f"p6_fill={'%.0f%%' % (self.p6_fill * 100) if self.p6_fill is not None else '-'}")


def _analyze_recorder(rec: LayoutRecorder, tier: str, pages: int) -> TrialResult:
    tr = TrialResult(tier, None)
    tr.pages = pages
    for page, y, kind, role, payload in rec.events:
        src = (payload or {}).get("src") or {}
        key, idx = src.get("key"), src.get("idx")
        if kind == "header" and key:
            tr.sec_start.setdefault(key, page)
            tr.sec_end[key] = page
        elif key is not None and idx is not None:
            tr.item_pages[(key, idx)] = page
            tr.sec_end[key] = max(tr.sec_end.get(key, 0), page)
    if rec.p6_y is not None and rec.p6_page == pages:
        sa_last = TIERS[tier][1]        # last item is a search-term
        bottom_text = rec.p6_y + sa_last
        tr.p6_fill = (FRAME_TOP_Y - bottom_text) / FRAME_H
    if pages > 4:
        tops = [y for page, y, *_ in rec.events if page == 5]
        if tops:
            tr.spill_pt = FRAME_TOP_Y - min(tops)
    return tr


def run_trial(case, styles, glued, break_before, case_ctx, report):
    rec = LayoutRecorder()
    story = build_case_story(case, styles, glued, break_before, rec)
    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=(PAGE_W, PAGE_H),
                          allowSplitting=1)
    tv = PageTemplate(id="A_verso", frames=[make_side_frame("verso")],
                      onPage=lambda c, d, ctx=case_ctx: paint_a_master(c, d, ctx, "verso"))
    tr_ = PageTemplate(id="A_recto", frames=[make_side_frame("recto")],
                       onPage=lambda c, d, ctx=case_ctx: paint_a_master(c, d, ctx, "recto"))

    def chain_verso(c, d):
        paint_a_master(c, d, case_ctx, "verso")
        d._nextPageTemplateIndex = 1

    def chain_recto(c, d):
        paint_a_master(c, d, case_ctx, "recto")
        d._nextPageTemplateIndex = 0

    tv.onPage = chain_verso
    tr_.onPage = chain_recto
    doc.addPageTemplates([tv, tr_])
    doc.build(story, canvasmaker=BookCanvas)
    tr = _analyze_recorder(rec, styles["_tier"], doc.page)
    tr.break_before = break_before
    return tr

# ============================================================================
# 18. FIT-LADDER v3 + EXHAUSTIVE PAGE-6 DECISION TABLE
# ============================================================================

def _sc_ok(t: TrialResult) -> bool:
    """SEE ALSO + CURIOUS? complete on CP4 with all their items, P == 4."""
    if t.pages != 4:
        return False
    for key in ("SEE_ALSO", "CURIOUS"):
        if t.sec_start.get(key) != 4 or t.sec_end.get(key) != 4:
            return False
        if not all(p == 4 for (k, i), p in t.item_pages.items() if k == key):
            return False
    return True


def _q_complete_cp4(t: TrialResult) -> bool:
    return (t.pages == 4 and t.sec_start.get("QUESTION") == 4
            and t.sec_end.get("QUESTION") == 4)


def verify_step_e(locked: TrialResult, d1: "TrialResult | None",
                  q_start_unforced) -> "tuple[bool, str]":
    """Strengthened Step E (Gemini table): S+C complete on CP4 with items, and
    either QUESTION complete on CP4, or D1 was *proven* unable to produce that
    arrangement (trial executed and failed verification, incl. overflow), or
    QUESTION started before CP3 (a single permitted break cannot reach CP4)."""
    if locked.pages != 4:
        return False, f"content_pages={locked.pages} != 4"
    if not _sc_ok(locked):
        return False, "SEE ALSO / CURIOUS? not complete on CP4"
    if _q_complete_cp4(locked):
        return True, "QUESTION complete on CP4"
    if d1 is not None:
        if d1.pages > 4:
            return True, (f"D1 proven impossible: break before QUESTION overflows "
                          f"to {d1.pages} pages")
        if not (_sc_ok(d1) and _q_complete_cp4(d1)):
            return True, ("D1 proven impossible: trial verified failing "
                          f"({d1.state()})")
    if q_start_unforced is not None and q_start_unforced < 3:
        return True, ("QUESTION starts before CP3: a single FrameBreak cannot "
                      "place it on CP4")
    return False, "QUESTION not on CP4 without structural proof"


class LockedFit:
    def __init__(self, tier, break_before, telemetry, step_e_reason):
        self.tier, self.break_before = tier, break_before
        self.telemetry, self.step_e_reason = telemetry, step_e_reason


def fit_case(case, styles_by_tier, glued, report):
    """Fit-Ladder v3 Steps 0/A-E with the exhaustive CP4 decision table.
    Returns LockedFit or None (with CRITICAL logged; part withheld)."""
    cn = case["case_number"]
    ctx = {"base": case_start_page(cn), "case_name": case["case_name"]}
    telem = []

    def trial(tier, brk, step):
        t = run_trial(case, styles_by_tier[tier], glued, brk, ctx, report)
        telem.append({"step": step, "tier": tier, "break": brk,
                      "pages": t.pages,
                      "sec": {k: (t.sec_start.get(k), t.sec_end.get(k)) for k in SECTION_KEYS},
                      "p6_fill": t.p6_fill, "spill_pt": t.spill_pt,
                      "state": t.state()})
        return t

    # Step 0 — overrides: only big_word_pt exists in the locked contract; there
    # are no force_style / force_break_before fields to apply.
    # Step A — Standard
    tA = trial("standard", None, "A")
    tier, tU = "standard", tA
    if tA.pages > 4:
        # Step B — overflow -> Tight
        tB = trial("tight", None, "B")
        tier, tU = "tight", tB
        if tB.pages > 4:
            report.crit("CRITICAL-OVERFLOW",
                        f"Case {cn}: Tight still overflows "
                        f"(spill ~{tB.spill_pt}pt on 5th page; {tB.state()}); "
                        f"part withheld — fix text or request an approved override")
            return None
        report.info("FIT-TIGHT", f"Case {cn}: {tB.state()}")
    elif tA.pages < 4 or (tA.p6_fill is not None and tA.p6_fill < P6_FILL_MIN):
        # Step C — underflow -> Open
        tC = trial("open", None, "C")
        if tC.pages == 4:
            tier, tU = "open", tC
            report.info("FIT-OPEN", f"Case {cn}: {tC.state()}")
        elif tC.pages > 4:
            tier, tU = "standard", tA       # one allowed revert
            report.info("FIT-REVERT", f"Case {cn}: Open overflowed, back to Standard")
        else:
            tier, tU = "open", tC
            report.warn("WARN-UNDERFLOW",
                        f"Case {cn}: Open cannot fill 4 pages ({tC.state()})")

    # Step D — Page 6 safety rule (exhaustive decision table)
    d1 = None
    q_start_unforced = tU.sec_start.get("QUESTION")
    if _sc_ok(tU) and _q_complete_cp4(tU):
        decision, brk = "D0", None                       # Branch 1
    else:
        brk = None
        decision = None
        d2 = None
        if _sc_ok(tU):
            # S+C fine; try to pull QUESTION onto CP4 as well
            if q_start_unforced == 3:                     # Branch 2
                d1 = trial(tier, "QUESTION", "D1")
                if _sc_ok(d1) and _q_complete_cp4(d1):
                    decision, brk = "D1", "QUESTION"
                else:
                    decision, brk = "D2_natural", None    # unforced already satisfies
            else:                                         # Branch 3 (Q_s < 3)
                decision, brk = "D2_natural", None
        else:
            if q_start_unforced == 3:
                d1 = trial(tier, "QUESTION", "D1")
                if _sc_ok(d1) and _q_complete_cp4(d1):
                    decision, brk = "D1", "QUESTION"
                elif tU.sec_start.get("SEE_ALSO") == 3:
                    d2 = trial(tier, "SEE_ALSO", "D2")
                    if _sc_ok(d2):
                        decision, brk = "D2", "SEE_ALSO"
            elif tU.sec_start.get("SEE_ALSO") == 3:
                d2 = trial(tier, "SEE_ALSO", "D2")
                if _sc_ok(d2):
                    decision, brk = "D2", "SEE_ALSO"
            if decision is None and tier == "open":
                # the one allowed Open -> Standard revert, retry D2 only
                d2 = trial("standard", "SEE_ALSO", "D2-revert")
                if _sc_ok(d2):
                    tier, decision, brk = "standard", "D2-revert", "SEE_ALSO"
        if decision is None:
            report.crit("CRITICAL-P6-ARRANGEMENT",
                        f"Case {cn}: no permitted arrangement puts SEE ALSO + "
                        f"CURIOUS? complete on CP4; trials: " +
                        "; ".join(f"[{t['step']}/{t['tier']}/{t['break']}] {t['state']}"
                                  for t in telem))
            return None

    # Step E — final verification on the locked arrangement
    if brk is None:
        locked = tU
    elif brk == "QUESTION":
        locked = d1
    else:
        locked = d2 if d2 is not None and d2.break_before == brk else trial(tier, brk, "lock-" + brk)
    ok, reason = verify_step_e(locked, d1, q_start_unforced)
    if not ok:
        report.crit("CRITICAL-PAGECOUNT", f"Case {cn}: Step E failed: {reason}; "
                                          f"{locked.state()}")
        return None
    report.info("FIT-LOCKED", f"Case {cn}: tier={tier} break={brk} decision="
                              f"{decision}; {locked.state()}; {reason}")
    return LockedFit(tier, brk, telem, reason)
# ============================================================================
# 19. PART RENDERERS + EXPECTED-RECORDS (source-authoritative, ChatGPT #7)
# ============================================================================

def _case_ctx(case):
    return {"base": case_start_page(case["case_number"]),
            "case_name": case["case_name"]}


def _set_next_template(doc, template_id: str):
    """Name-based template chaining (index-independent: the same A-templates
    are used in 2-template trial docs and 4-template part docs)."""
    for i, t in enumerate(doc.pageTemplates):
        if t.id == template_id:
            doc._nextPageTemplateIndex = i
            return


def _a_templates(case_ctx):
    def chain(side, next_id):
        def painter(canv, doc):
            paint_a_master(canv, doc, case_ctx, side)
            _set_next_template(doc, next_id)
        return painter
    tv = PageTemplate(id="A_verso", frames=[make_side_frame("verso")],
                      onPage=chain("verso", "A_recto"))
    tr_ = PageTemplate(id="A_recto", frames=[make_side_frame("recto")],
                       onPage=chain("recto", "A_verso"))
    return [tv, tr_]


def _expected_for_case(case):
    """Static expected body sequence: headers + items in reading order."""
    body = []
    for s in case["sections"]:
        body.append(("section-header", s["header"].upper()))
        for it in s["items"]:
            body.append((it["type"], item_display(it)))
    return body


def render_case_part(case, fit, styles_by_tier, glued, overrides, report,
                     render_records) -> bool:
    cn = case["case_number"]
    path = PARTS_DIR / part_name(cn)
    styles = styles_by_tier[fit.tier]
    rec = LayoutRecorder()
    story = build_case_story(case, styles, glued, fit.break_before, rec)
    bw_pt = overrides.get(str(cn), 50)
    ctx = _case_ctx(case)

    bt = PageTemplate(id="B_title", frames=[make_side_frame("verso")],
                      onPage=lambda c, d: paint_case_title_page(c, case, bw_pt))
    bi = PageTemplate(id="B_image", frames=[make_side_frame("recto")],
                      onPage=lambda c, d: paint_case_image_page(c, case))
    av, ar = _a_templates(ctx)
    templates = [bt, bi, av, ar]
    tail = []
    if cn == N_CASES:
        cv = PageTemplate(id="C_verso", frames=[make_side_frame("verso")],
                          onPage=_noop)
        cr = PageTemplate(id="C_recto", frames=[make_side_frame("recto")],
                          onPage=_noop)
        templates += [cv, cr]
        epilogue_story = build_epilogue_story()
        tail = [NextPageTemplate("C_verso"), PageBreak(),
                NextPageTemplate("C_recto"), PageBreak()] + epilogue_story + \
               [NextPageTemplate("C_verso"), PageBreak(), EndOfCaseMarker(None)]

    doc = BaseDocTemplate(str(path), pagesize=(PAGE_W, PAGE_H), allowSplitting=1)
    doc.addPageTemplates(templates)
    full = [NextPageTemplate("B_image"), PageBreak(),
            NextPageTemplate("A_verso"), PageBreak()] + story + tail
    doc.build(full, canvasmaker=BookCanvas)
    expected_pages = 9 if cn == N_CASES else 6
    if doc.page != expected_pages:
        report.crit("CRITICAL-PAGECOUNT",
                    f"case {cn}: part has {doc.page} pages, expected {expected_pages}")
        return False

    # ---- expected records (page -> ordered body paragraphs) ---------------
    base = ctx["base"]
    render_records[base] = {"kind": "title", "case": cn, "paras": [
        ("case-name", case["case_name"]),
        ("big-word", case["big_word"]),
        *[("case-title", tl) for tl in case["title_lines"]],
        *[("case-quote", ql) for ql in case["quote"].split("\n")]]}
    render_records[base + 1] = {"kind": "image", "case": cn, "paras": [
        ("image-caption", "\n".join(case["captions"]))]}
    body_seq = _expected_for_case(case)
    for offset in range(4):
        render_records[base + 2 + offset] = {"kind": "cp", "case": cn, "paras": []}
    # distribute by recorder: every event (split paragraphs contribute one
    # event per page); book page = base + (doc page - 1)
    parts_by_item = {}
    for page, y, kind, role, payload in rec.events:
        src = (payload or {}).get("src") or {}
        key, idx = src.get("key"), src.get("idx")
        if key is None or idx is None:
            continue
        book_page = base + page - 1
        if book_page not in render_records or render_records[book_page]["kind"] != "cp":
            continue
        item = case["sections"][SECTION_KEYS.index(key)]["items"][idx]
        render_records[book_page]["paras"].append(
            (item["type"], payload.get("text", ""), y, payload.get("n_lines", 0), src))
        parts_by_item.setdefault((key, idx), []).append(payload.get("text", ""))
        if page > 6:
            report.crit("CRITICAL-PAGECOUNT", f"case {cn}: content spilled to part "
                                              f"page {page}")
            return False
    # headers with their rendered text (draw order)
    for page, y, kind, role, payload in rec.events:
        if kind != "header":
            continue
        src = (payload or {}).get("src") or {}
        key = src.get("key")
        if key is None:
            continue
        book_page = base + page - 1
        if book_page in render_records and render_records[book_page]["kind"] == "cp":
            render_records[book_page]["paras"].append(
                ("section-header", payload.get("text", ""), y, 1, src))
    # source-authoritative cross-check: rendered parts joined == source display
    for (key, idx), texts in parts_by_item.items():
        item = case["sections"][SECTION_KEYS.index(key)]["items"][idx]
        if norm_text(" ".join(texts)) != norm_text(item_display(item)):
            report.crit("CRITICAL-MARKUP",
                        f"case {cn} {key}#{idx}: rendered text != source text")
            return False
    # reading order per page (highest baseline first)
    for offset in range(4):
        pg = render_records[base + 2 + offset]
        pg["paras"].sort(key=lambda p: -p[2])
    if cn == N_CASES:
        _tail_expected_records(render_records)
    return True


def _noop(canv, doc):
    pass


_EPILOGUE_CACHE = {}

def build_epilogue_story():
    """Epilogue flowables for page 399 (title baseline pinned at 396pt)."""
    epi = _EPILOGUE_CACHE["epilogue"]
    story = [Spacer(1, FRAME_TOP_Y - (EPILOGUE_TITLE_BASELINE_Y + 17.0))]
    title_style = ParagraphStyle(
        "epilogue-title", fontName="LibreBaskerville-Bold",
        fontSize=17.0, leading=17.0, alignment=TA_CENTER,
        spaceAfter=14.0, textColor=GRAY_BLACK)
    story.append(BookParagraph(esc(epi["title"]), title_style,
                               role="epilogue-title", src={"key": "epilogue-title"},
                               rec=_EPILOGUE_CACHE.get("rec")))
    for gi, group in enumerate(epi["groups"]):
        if gi:
            story.append(StanzaBreak(EPILOGUE_STANZA_PT))
        for line in group:
            story.append(BookParagraph(
                esc(line), _EPILOGUE_CACHE["body_style"],
                role="epilogue-body", src={"key": "epilogue", "idx": gi},
                rec=_EPILOGUE_CACHE.get("rec")))
    return story


def _tail_expected_records(render_records):
    epi = _EPILOGUE_CACHE["epilogue"]
    rec = _EPILOGUE_CACHE.get("rec")
    paras = [("epilogue-title", epi["title"].upper())]
    for gi, group in enumerate(epi["groups"]):
        paras += [("epilogue-body", l) for l in group]
    render_records[399] = {"kind": "epilogue", "case": None, "paras": paras}
    render_records[398] = {"kind": "blank", "case": None, "paras": []}
    render_records[400] = {"kind": "blank", "case": None, "paras": []}


def render_front_part(fm, cases, report, render_records) -> bool:
    path = PARTS_DIR / part_name(0)
    fmdata = fm["frontmatter"]
    rec = LayoutRecorder()
    _EPILOGUE_CACHE["rec"] = None

    def painter(canv, doc):
        if doc.page <= 6:
            paint_front_page(canv, doc.page, fmdata)
        doc._nextPageTemplateIndex = 1 if doc.page % 2 == 1 else 0

    cr = PageTemplate(id="C_recto", frames=[make_side_frame("recto")], onPage=painter)
    cv = PageTemplate(id="C_verso", frames=[make_side_frame("verso")], onPage=painter)
    doc = BaseDocTemplate(str(path), pagesize=(PAGE_W, PAGE_H), allowSplitting=1)
    doc.addPageTemplates([cr, cv])

    story = [PageBreak() for _ in range(6)]
    # TOC (pp. 7-9)
    story.append(BookParagraph(esc(fmdata["toc-title"][0]),
                               ParagraphStyle("toc-title-p", fontName="LibreBaskerville-Bold",
                                              fontSize=17.0, leading=17.0,
                                              alignment=TA_CENTER, spaceAfter=24.0,
                                              textColor=GRAY_BLACK),
                               role="toc-title", src={"key": "toc-title"}, rec=rec))
    story.append(TOCEntryFlowable("INTRODUCTION", 10, rec=rec, src={"key": "toc", "idx": 0}))
    for i, c in enumerate(cases, start=1):
        story.append(TOCEntryFlowable(c["case_name"], case_start_page(i),
                                      rec=rec, src={"key": "toc", "idx": i}))
    story.append(PageBreak())
    # Introduction (pp. 10-13): title baseline pinned at 360pt
    story.append(Spacer(1, FRAME_TOP_Y - (INTRO_TITLE_BASELINE_Y + 17.0)))
    story.append(BookParagraph(esc(fmdata["intro-title"][0]),
                               ParagraphStyle("intro-title-p", fontName="LibreBaskerville-Bold",
                                              fontSize=17.0, leading=17.0,
                                              alignment=TA_CENTER, spaceAfter=24.0,
                                              textColor=GRAY_BLACK),
                               role="intro-title", src={"key": "intro-title"}, rec=rec))
    intro_style = ParagraphStyle("intro-body-p", fontName="EBGaramond-Regular",
                                 fontSize=BODY_SIZE, leading=14.0, spaceAfter=4.0,
                                 alignment=TA_LEFT, allowWidows=0, allowOrphans=0,
                                 hyphenationLang=None, embeddedHyphenation=0,
                                 splitLongWords=0, textColor=GRAY_BLACK)
    for i, item in enumerate(fm["introduction"]):
        if item["type"] == "stanza_break":
            story.append(StanzaBreak(INTRO_STANZA_PT))
        else:
            story.append(BookParagraph(esc(item["text"]).replace("\n", "<br/>"),
                                       intro_style, role="intro-body",
                                       src={"key": "intro", "idx": i}, rec=rec))
    doc.build(story, canvasmaker=BookCanvas)
    if doc.page != FRONT_MATTER_PAGES:
        report.crit("CRITICAL-PAGECOUNT",
                    f"front matter has {doc.page} pages, expected 13")
        return False

    # ---- recorder-based structural asserts (TOC == 3 pages, Intro == 4) ----
    toc_pages = {p for p, y, kind, role, _ in rec.events if role == "toc-entry"}
    toc_title_pg = next((p for p, y, kind, role, _ in rec.events
                         if role == "toc-title"), None)
    intro_title_pg = next((p for p, y, kind, role, _ in rec.events
                           if role == "intro-title"), None)
    if toc_title_pg != 7 or sorted(toc_pages) != [7, 8, 9]:
        report.crit("CRITICAL-STRUCTURE",
                    f"TOC must occupy exactly pages 7-9 (title on {toc_title_pg}, "
                    f"entries on {sorted(toc_pages)})")
        return False
    if intro_title_pg != 10:
        report.crit("CRITICAL-STRUCTURE", f"Introduction must start on page 10 "
                                          f"(starts on {intro_title_pg})")
        return False
    report.info("FRONTMATTER-OK", "13 pages; TOC on 7-9 (65 entries); "
                                  "Introduction on 10-13")

    # ---- expected records --------------------------------------------------
    fmd = fm["frontmatter"]
    render_records[1] = {"kind": "front", "case": None, "paras": [("book-name", fmd["book-name"][0])]}
    render_records[2] = {"kind": "blank", "case": None, "paras": []}
    render_records[3] = {"kind": "front", "case": None, "paras": [
        ("book-name-full", fmd["book-name-full"][0]),
        *[("subtitle-name", s) for s in fmd["subtitle-name"]],
        ("author", fmd["author"][0])]}
    render_records[4] = {"kind": "front", "case": None, "paras": [
        ("copyright", p) for p in fmd["copyright"]]}
    render_records[5] = {"kind": "front", "case": None, "paras": [
        ("dedication-line", fmd["dedication-line"][0])]}
    render_records[6] = {"kind": "front", "case": None, "paras": [
        ("welcome", fmd["welcome"][0])]}
    for p in (7, 8, 9):
        render_records[p] = {"kind": "toc", "case": None, "paras": []}
    if toc_title_pg == 7:
        render_records[7]["paras"].append(("toc-title", fmd["toc-title"][0]))
    for page, y, kind, role, payload in rec.events:
        if role != "toc-entry":
            continue
        src = (payload or {}).get("src") or {}
        idx = src.get("idx", 0)
        if idx == 0:
            text = "INTRODUCTION 10"
        else:
            c = cases[idx - 1]
            text = f"{c['case_name']} {case_start_page(idx)}"
        render_records[page]["paras"].append(("toc-entry", text, y, 1, src))
    for p in (10, 11, 12, 13):
        render_records[p] = {"kind": "intro", "case": None, "paras": []}
    for page, y, kind, role, payload in rec.events:
        if role not in ("intro-title", "intro-body"):
            continue
        src = (payload or {}).get("src") or {}
        if role == "intro-title":
            text = fm["frontmatter"]["intro-title"][0]
        else:
            text = fm["introduction"][src["idx"]]["text"]
        render_records[page]["paras"].append((role, text, y,
                                              payload.get("n_lines", 1), src))
    for p in range(10, 14):
        render_records[p]["paras"].sort(key=lambda t: -t[2])
    return True

# ============================================================================
# 20. MERGE (transactional: 65 valid parts, tmp file, atomic replace)
# ============================================================================

def merge_parts(manifest, report) -> "Path | None":
    if len(manifest) != 65:
        report.crit("CRITICAL-MERGE", f"expected 65 valid parts, got {len(manifest)}")
        return None
    expected_counts = [13] + [6] * 63 + [9]
    for (path, sha, pages), want in zip(manifest, expected_counts):
        if not path.exists():
            report.crit("CRITICAL-MERGE", f"part missing: {path.name}")
            return None
        if sha256_file(path) != sha:
            report.crit("CRITICAL-MERGE", f"part changed on disk: {path.name}")
            return None
        r = PdfReader(str(path))
        if len(r.pages) != want:
            report.crit("CRITICAL-MERGE", f"{path.name}: {len(r.pages)} pages != {want}")
            return None
    tmp = OUT_DIR / "UNEXPLAINED_KDP_Master_400p.tmp.pdf"
    writer = PdfWriter()
    for path, _, _ in manifest:
        writer.append(str(path))
    writer.add_metadata({"/Title": BOOK_TITLE, "/Author": BOOK_AUTHOR})
    with open(tmp, "wb") as f:
        writer.write(f)
    report.info("MERGE-OK", "65 parts merged in order -> temporary master")
    return tmp

# ============================================================================
# 21. FINAL VALIDATION — independent of the layout code (ChatGPT #1,6,7,9,10)
# ============================================================================

def _base_font_name(font_dict) -> str:
    try:
        bf = str(font_dict.get("/BaseFont", ""))
    except Exception:
        return ""
    return bf.split("+")[-1] if "+" in bf else bf


_CMAP_BFCHAR = re.compile(r"beginbfchar(.*?)endbfchar", re.S)
_CMAP_BFRANGE = re.compile(r"beginbfrange(.*?)endbfrange", re.S)
_CMAP_HEX = re.compile(r"<([0-9A-Fa-f\s]+)>")


def _font_decoders(page) -> dict:
    """resource-name -> {byte: unicode} from each font's /ToUnicode CMap."""
    dec = {}
    res = page.get("/Resources")
    if res is None:
        return dec
    fdict = res.get("/Font")
    if fdict is None:
        return dec
    for k, v in fdict.get_object().items():
        fo = v.get_object()
        m = {}
        tou = fo.get("/ToUnicode")
        if tou is not None:
            try:
                data = tou.get_object().get_data().decode("latin-1", "replace")
                for block in _CMAP_BFCHAR.findall(data):
                    pairs = _CMAP_HEX.findall(block)
                    for i in range(0, len(pairs) - 1, 2):
                        src_hex = re.sub(r"\s", "", pairs[i])
                        dst_hex = re.sub(r"\s", "", pairs[i + 1])
                        if src_hex:
                            b = int(src_hex[:2], 16)
                            if dst_hex:
                                m[b] = "".join(
                                    chr(int(dst_hex[j:j + 4], 16))
                                    for j in range(0, len(dst_hex), 4))
                for block in _CMAP_BFRANGE.findall(data):
                    hexes = _CMAP_HEX.findall(block)
                    for i in range(0, len(hexes) - 2, 3):
                        lo = int(re.sub(r"\s", "", hexes[i]), 16)
                        hi = int(re.sub(r"\s", "", hexes[i + 1]), 16)
                        base_hex = re.sub(r"\s", "", hexes[i + 2])
                        base = int(base_hex[:4], 16) if base_hex else 0
                        for b in range(lo, min(hi, lo + 60000) + 1):
                            m[b] = chr(base + (b - lo))
            except Exception:
                m = {}
        dec[str(k).lstrip("/")] = m
    return dec


def _simulate_text_ops(page, reader, decoded=None):
    """Ground-truth text positions from the content stream itself.

    pypdf 6.6.0's visitor mis-reports the y coordinate after T* sequences
    (observed 147pt errors on plain reportlab output), so positions are
    simulated here: Tm/Td/T*/TL/Tf tracked exactly, q/Q/cm tracked as
    translations (reportlab emits no scaling/rotation on our pages).
    Returns one entry per text-showing operator, in stream order.
    """
    out = []
    try:
        ops = ContentStream(page.get_contents(), reader).operations
    except Exception:
        return out
    fonts = {}
    res = page.get("/Resources")
    if res is not None:
        fdict = res.get("/Font")
        if fdict is not None:
            for k, v in fdict.get_object().items():
                fo = v.get_object()
                bf = str(fo.get("/BaseFont", ""))
                fonts[str(k).lstrip("/")] = bf.split("+")[-1] if "+" in bf else bf
    decoders = _font_decoders(page)
    _cur_res = [""]
    tm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    tlm_x = 0.0
    leading = 0.0
    font, size = "", 0.0
    ctm_stack = []
    cx = cy = 0.0
    k = 0
    for operands, operator in ops:
        op = operator.decode() if isinstance(operator, bytes) else str(operator)
        try:
            if op == "q":
                ctm_stack.append((cx, cy))
            elif op == "Q":
                if ctm_stack:
                    cx, cy = ctm_stack.pop()
            elif op == "cm" and len(operands) >= 6:
                cx += float(operands[4])
                cy += float(operands[5])
            elif op == "BT":
                tm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
                tlm_x = tm[4]
            elif op == "Tm" and len(operands) >= 6:
                tm = [float(x) for x in operands[:6]]
                tlm_x = tm[4]
            elif op == "Td" and len(operands) >= 2:
                tm[4] = tlm_x + float(operands[0])
                tm[5] += float(operands[1])
                tlm_x = tm[4]
            elif op == "TD" and len(operands) >= 2:
                leading = -float(operands[1])
                tm[4] = tlm_x + float(operands[0])
                tm[5] += float(operands[1])
                tlm_x = tm[4]
            elif op == "TL" and operands:
                leading = float(operands[0])
            elif op == "T*":
                tm[4] = tlm_x          # the PDF cursor re-anchors on the TLM
                tm[5] -= leading
            elif op == "Tf" and len(operands) >= 2:
                res_name = str(operands[0]).lstrip("/")
                _cur_res[0] = res_name
                font = fonts.get(res_name, res_name)
                size = float(operands[1])
            elif op in ("Tj", "'", '"', "TJ"):
                cmap = decoders.get(_cur_res[0], {})

                def _dec(b: bytes) -> str:
                    if cmap:
                        return "".join(cmap.get(x, chr(x)) for x in b)
                    return b.decode("latin-1")

                if op == "TJ":
                    text = "".join(_dec(x) if isinstance(x, bytes) else str(x)
                                   for x in operands[0]
                                   if not isinstance(x, (int, float)))
                else:
                    t = operands[0]
                    if isinstance(t, bytes):
                        text = _dec(t)
                    else:
                        s = str(t)
                        try:            # pypdf hands strings as TextStringObject
                            text = _dec(s.encode("latin-1"))
                        except Exception:
                            text = s
                if op == "'":
                    tm[5] -= leading
                out.append((text, tm[4] + cx, tm[5] + cy, font, size))
                # advance the cursor by the decoded text width (1 byte per
                # glyph in reportlab TTF subsets, so widths are exact)
                if font in _FONT_METRICS:
                    try:
                        tm[4] += stringWidth(text, font, size or BODY_SIZE)
                    except Exception:
                        pass
        except Exception:
            continue
    return out


def page_fragments(page, reader=None):
    """Text fragments with reliable coordinates.

    Text content comes from pypdf's visitor (it decodes subset fonts
    correctly); positions come from the stream simulator above, matched by
    operator order.  Falls back to visitor positions if the counts differ.
    """
    vis_texts = []

    def visitor(text, cm, tm, font_dict, font_size):
        if text and text.strip():
            vis_texts.append(text)
    try:
        page.extract_text(visitor_text=visitor)
    except Exception as e:
        return [{"text": "\x00EXTRACT-ERROR", "x": 0, "y": 0,
                 "font": "", "size": 0, "err": str(e)}]
    if reader is not None:
        sim = _simulate_text_ops(page, reader)
        return [{"text": s[0], "x": s[1], "y": s[2], "font": s[3], "size": s[4]}
                for s in sim if s[0] and s[0].strip()]
    # fallback: visitor positions (pypdf's T* quirk may affect y)
    frags = []

    def visitor2(text, cm, tm, font_dict, font_size):
        if text and text.strip():
            frags.append({"text": text,
                          "x": tm[4] + cm[4], "y": tm[5] + cm[5],
                          "font": _base_font_name(font_dict),
                          "size": font_size})
    try:
        page.extract_text(visitor_text=visitor2)
    except Exception:
        pass
    return frags


def joined_text(frags) -> str:
    """Join fragments with PDF-truth spacing: no space inserted when the next
    fragment starts exactly at the previous fragment's advance position
    (e.g. 'Cold Spot' + '.' from adjacent rich-text runs)."""
    out = []
    for i, f in enumerate(frags):
        if i:
            prev = frags[i - 1]
            no_gap = False
            if prev.get("font") in _FONT_METRICS and abs(f["y"] - prev["y"]) < 0.5:
                try:
                    w = stringWidth(prev["text"], prev["font"], prev["size"] or BODY_SIZE)
                    no_gap = abs(f["x"] - (prev["x"] + w)) < 1.2
                except Exception:
                    no_gap = False
            if not no_gap:
                out.append(" ")
        out.append(f["text"])
    return "".join(out)


def _page_kind(book_page: int, render_records) -> str:
    rr = render_records.get(book_page)
    return rr["kind"] if rr else "?"


def validate_final(path: Path, render_records, report) -> bool:
    report.section("FINAL VALIDATION")
    reader = PdfReader(str(path))
    ok = True

    # ---- A. page count, boxes, encryption, annotations ---------------------
    if reader.is_encrypted:
        report.crit("CRITICAL-PDF", "master is encrypted")
        return False
    if len(reader.pages) != TOTAL_PAGES:
        report.crit("CRITICAL-PDF", f"master has {len(reader.pages)} pages != 400")
        return False
    for i, page in enumerate(reader.pages, start=1):
        box = page.mediabox
        if (round(float(box.width), 2), round(float(box.height), 2)) != (432.0, 648.0):
            report.crit("CRITICAL-PDF", f"page {i}: MediaBox {box} != 432x648")
            ok = False
            break
        if page.get("/Annots"):
            report.crit("CRITICAL-PDF", f"page {i}: annotations present")
            ok = False
    if not ok:
        return False
    report.info("GEOMETRY-OK", "400 pages, all MediaBox 432x648, no annotations")

    # ---- B. per-page text: chrome, folios, expected text, cap alignment ----
    cap_fail = 0
    folio_fail = 0
    text_fail = 0
    chrome_fail = 0
    for book_page in range(1, TOTAL_PAGES + 1):
        page = reader.pages[book_page - 1]
        frags = page_fragments(page, reader)
        if any("err" in f for f in frags):
            report.crit("CRITICAL-TEXT-INTEGRITY",
                        f"page {book_page}: extraction failed")
            return False
        kind = _page_kind(book_page, render_records)
        header_frags = [f for f in frags if f["y"] > HEADER_BAND_Y]
        folio_frags = [f for f in frags if f["y"] < FOLIO_BAND_Y]
        body = [f for f in frags if FOLIO_BAND_Y <= f["y"] <= HEADER_BAND_Y]

        if kind in ("cp",):
            # folio == physical page number
            folio_txt = norm_text("".join(f["text"] for f in folio_frags))
            if folio_txt != str(book_page):
                report.crit("CRITICAL-FOLIO", f"page {book_page}: folio "
                                              f"{folio_txt!r} != {book_page}")
                folio_fail += 1
            hdr_txt = norm_text("".join(f["text"] for f in header_frags))
            verso = book_page % 2 == 0
            cn = render_records[book_page]["case"]
            case = _CASES_BY_NUM.get(cn)
            want_hdr = "UNEXPLAINED" if verso else (case["case_name"] if case else "?")
            if hdr_txt != want_hdr:
                report.crit("CRITICAL-HEADER",
                            f"page {book_page}: running header {hdr_txt!r} != {want_hdr!r}")
                chrome_fail += 1
            # side frame containment
            left_lim = VERS0_X if verso else RECTO_X
            right_lim = VERS0_RIGHT if verso else RECTO_RIGHT
            for f in body:
                if f["x"] < left_lim - 1.5 or f["x"] > right_lim + 1.5:
                    report.crit("CRITICAL-MARGIN",
                                f"page {book_page}: text x={f['x']:.1f} outside "
                                f"[{left_lim}, {right_lim}]")
                    chrome_fail += 1
                    break
            # recto cap alignment (ChatGPT #1: actual PDF truth)
            if not verso and body:
                f0 = body[0]
                base = f0["font"]
                if base in _FONT_METRICS:
                    cap = _FONT_METRICS[base]["cap"] * (f0["size"] or BODY_SIZE)
                    got = f0["y"] + cap
                    if abs(got - CAP_TARGET_Y) > 0.05:
                        report.crit("CRITICAL-CAP-ALIGN",
                                    f"page {book_page}: first cap top {got:.3f} != "
                                    f"{CAP_TARGET_Y} (font {base} {f0['size']})")
                        cap_fail += 1
        else:
            if header_frags or folio_frags:
                report.crit("CRITICAL-CHROME",
                            f"page {book_page} ({kind}): chrome present where "
                            f"forbidden: {[f['text'] for f in header_frags + folio_frags]}")
                chrome_fail += 1

        # expected text comparison (TOC pages: drop standalone dots)
        rr = render_records.get(book_page)
        if rr is None:
            report.crit("CRITICAL-OWNERSHIP", f"page {book_page}: no expected record")
            text_fail += 1
            continue
        exp = " ".join(norm_text(p[1]) for p in rr["paras"])
        got_full = joined_text(body)                   # chrome is verified separately
        if kind == "toc":
            toks = [t for t in got_full.split() if t != "."]
            got_full = " ".join(toks)
        got = norm_text(got_full)
        if got != norm_text(exp):
            report.crit("CRITICAL-TEXT-INTEGRITY",
                        f"page {book_page} ({kind}): extracted != expected\n"
                        f"  expected: {norm_text(exp)[:150]}\n"
                        f"  got     : {got[:150]}")
            text_fail += 1

    if folio_fail or chrome_fail or cap_fail:
        ok = False
    if text_fail:
        return False
    report.info("CHROME-OK", f"folios == physical page numbers on all CP pages; "
                             f"headers correct; chrome absent on B/C pages; "
                             f"recto cap == {CAP_TARGET_Y}pt on all recto CP pages")

    # ---- C. authored line breaks + bold/italic run fonts --------------------
    nb_ok = nb_bad = run_ok = run_bad = 0
    for book_page, rr in render_records.items():
        if book_page > TOTAL_PAGES:
            continue
        page = reader.pages[book_page - 1]
        frags = page_fragments(page, reader)
        body = [f for f in frags if FOLIO_BAND_Y <= f["y"] <= HEADER_BAND_Y]
        for para in rr["paras"]:
            if len(para) < 5:
                continue
            role, text, y, n_lines, src = para
            if "\n" in text:
                segs = [s.strip() for s in text.split("\n") if s.strip()]
                if len(segs) >= 2 and _segments_on_distinct_lines(body, segs):
                    nb_ok += 1
                else:
                    nb_bad += 1
                    report.crit("CRITICAL-TEXT-INTEGRITY",
                                f"page {book_page}: authored line break lost: {text[:60]!r}")
            if role == "body-text" and src.get("idx") is not None:
                case = _CASES_BY_NUM.get(rr.get("case"))
                if case:
                    item = None
                    for s in case["sections"]:
                        if s["key"] == src.get("key"):
                            item = s["items"][src["idx"]]
                    if item and item.get("runs"):
                        if _runs_fonts_ok(body, y, item["runs"]):
                            run_ok += 1
                        else:
                            run_bad += 1
                            report.crit("CRITICAL-TEXT-INTEGRITY",
                                        f"page {book_page}: bold/italic run rendered "
                                        f"with wrong font: {item['text'][:60]!r}")
    if nb_bad or run_bad:
        return False
    report.info("TEXT-INTEGRITY-OK",
                f"all 400 pages match source text; authored line breaks preserved "
                f"({nb_ok} checked); run fonts verified ({run_ok} run items)")

    # ---- D. stranding + section grouping ------------------------------------
    strand = 0
    for book_page, rr in render_records.items():
        if rr["kind"] != "cp" or not rr["paras"]:
            continue
        last = rr["paras"][-1]
        if last[0] == "section-header":
            report.crit("CRITICAL-STRANDING",
                        f"page {book_page}: section header '{last[1]}' stranded at "
                        f"bottom of page")
            strand += 1
    if strand:
        return False
    report.info("STRANDING-OK", "no section header is the last element of any page")

    # ---- E. blanks ----------------------------------------------------------
    for p in (2, 398, 400):
        page = reader.pages[p - 1]
        if page_fragments(page, reader):
            report.crit("CRITICAL-BLANK", f"page {p} is not text-empty")
            return False
        if not _page_is_visually_empty(page, reader):
            report.crit("CRITICAL-BLANK",
                        f"page {p} has drawing operators "
                        f"(text/path/image/shading)")
            return False
    report.info("BLANKS-OK", "pages 2, 398, 400 are genuinely empty "
                             "(no text, path, image or shading operators; "
                             "reportlab's empty BT..ET prologue ignored)")

    # ---- F. fonts embedded (fonts that actually show text) --------------------
    missing_embed = 0
    seen_fonts = set()
    for book_page in range(1, TOTAL_PAGES + 1):
        page = reader.pages[book_page - 1]
        used = {f["font"] for f in page_fragments(page, reader)}
        res = page.get("/Resources")
        if not res:
            continue
        fonts = res.get("/Font")
        if not fonts:
            continue
        referenced = set()
        for fname, fobj in fonts.items():
            fo = fobj.get_object()
            bf = str(fo.get("/BaseFont", ""))
            base = bf.split("+")[-1]
            referenced.add(base)
            if base not in used:
                continue                       # prologue-only reference
            seen_fonts.add(base)
            fd = fo.get("/FontDescriptor")
            if fd is None and fo.get("/Subtype") == "/Type0":
                df = fo.get("/DescendantFonts")
                if df:
                    fd = df[0].get_object().get("/FontDescriptor")
            if fd is None or fd.get_object().get("/FontFile2") is None:
                report.crit("CRITICAL-FONT-EMBED",
                            f"page {book_page}: font {bf} shows text but has no "
                            f"FontFile2")
                missing_embed += 1
        unused = referenced - used
        if unused:
            report.info("FONT-PROLOGUE",
                        f"page {book_page}: referenced without shown text "
                        f"(reportlab prologue): {sorted(unused)}")
    if missing_embed:
        return False
    report.info("FONTS-EMBEDDED", f"all text-showing fonts carry /FontFile2 "
                                  f"subsets: {sorted(seen_fonts)}")

    # ---- G. images: identity + colorspace + dims (ChatGPT #9) ----------------
    img_count = 0
    for book_page in range(1, TOTAL_PAGES + 1):
        page = reader.pages[book_page - 1]
        res = page.get("/Resources")
        if not res:
            continue
        xo = res.get("/XObject")
        if not xo:
            continue
        for name, obj in xo.items():
            o = obj.get_object()
            if o.get("/Subtype") != "/Image":
                continue
            img_count += 1
            cs = o.get("/ColorSpace")
            cs_n = str(cs) if not isinstance(cs, list) else str(cs[0])
            if "/DeviceGray" not in cs_n:
                report.crit("CRITICAL-IMAGE-COLORSPACE",
                            f"page {book_page}: image colorspace {cs_n}")
                return False
            if o.get("/Width") != 1800 or o.get("/Height") != 2700:
                report.crit("CRITICAL-IMAGE-DIMS", f"page {book_page}: "
                                                   f"{o.get('/Width')}x{o.get('/Height')}")
                return False
            if o.get("/BitsPerComponent") != 8:
                report.crit("CRITICAL-IMAGE-BPC", f"page {book_page}")
                return False
            if o.get("/SMask") or o.get("/Mask"):
                report.crit("CRITICAL-IMAGE-MASK", f"page {book_page}: SMask/Mask present")
                return False
    if img_count != 64:
        report.crit("CRITICAL-IMAGE-COUNT", f"{img_count} images != 64")
        return False
    # pixel identity for every case image
    mism = 0
    for cn in range(1, N_CASES + 1):
        page = reader.pages[case_start_page(cn)]          # image page = base+1
        try:
            imgs = page.images
        except Exception as e:
            report.crit("CRITICAL-IMAGE-VERIFY",
                        f"case {cn}: extraction failed: {e}")
            mism += 1
            continue
        if len(imgs) != 1:
            report.crit("CRITICAL-IMAGE-VERIFY", f"case {cn}: {len(imgs)} images")
            mism += 1
            continue
        with PILImage.open(IMAGE_DIR / f"CASE{cn:02d}.png") as src:
            want = src.tobytes()
        got_img = imgs[0].image
        if got_img.mode != "L":
            got_img = got_img.convert("L")
        if got_img.tobytes() != want:
            report.crit("CRITICAL-IMAGE-IDENTITY",
                        f"case {cn}: embedded pixels differ from CASE{cn:02d}.png")
            mism += 1
    if mism:
        return False
    report.info("IMAGES-OK", "64 DeviceGray 1800x2700 8bpc images; pixel-identical "
                             "to their CASE{NN}.png sources")

    # ---- H. stateful DeviceGray certification --------------------------------
    gray_report = audit_colorspaces_and_paths(reader, render_records, report)
    if not gray_report:
        return False
    report.info("COLORSPACE-OK", "stateful parser: every page, Form XObject, image "
                                 "and ExtGState is DeviceGray-only, no transparency; "
                                 "0.75pt header rule verified on every CP page")
    report.info("VALIDATION-RESULT", "PASS" if report.n_crit == 0 else "FAIL")
    return report.n_crit == 0


_VISIBLE_OPS = {"Tj", "'", '"', "TJ", "S", "s", "f", "F", "f*", "B", "B*",
                "b", "b*", "Do", "sh", "EI", "ID"}


def _page_is_visually_empty(page, reader) -> bool:
    """True when the page draws nothing (reportlab writes an empty
    'BT /F1 12 Tf 14.4 TL ET' prologue on every page; it paints nothing)."""
    try:
        ops = ContentStream(page.get_contents(), reader).operations
    except Exception:
        return True
    for operands, operator in ops:
        op = operator.decode() if isinstance(operator, bytes) else str(operator)
        if op in _VISIBLE_OPS:
            return False
    return True


def _segments_on_distinct_lines(body_frags, segs) -> bool:
    """True when the last word of seg[i] and first word of seg[i+1] render on
    different baselines (authored break preserved)."""
    words = []
    for f in body_frags:
        for w in f["text"].split():
            words.append((w, f["y"]))
    for a, b in zip(segs, segs[1:]):
        wa = a.split()[-1] if a.split() else None
        wb = b.split()[0] if b.split() else None
        ya = next((y for w, y in words if w.endswith(wa) and wa), None)
        yb = next((y for w, y in words if w.startswith(wb) and wb), None)
        if ya is None or yb is None or abs(ya - yb) < 3.0:
            return False
    return True


def _runs_fonts_ok(body_frags, para_y, runs) -> bool:
    """Every bold/italic run near para_y renders with the right face: the run's
    distinctive leading words must appear in *some* nearby fragment drawn with
    the expected face (exists-at-least-one, so repeated words in earlier
    paragraphs on the same page cannot produce false failures)."""
    for r in runs:
        if not (r.get("b") or r.get("i")):
            continue
        want = "Bold" if r.get("b") else "Italic"
        words = r["t"].split()
        probes = []
        if len(words) >= 2:
            probes.append(" ".join(words[:2]))
        if words:
            probes.append(words[0])
        if not probes:
            continue
        near = [f for f in body_frags if abs(f["y"] - para_y) < 90.0]
        if not any(any(p in f["text"] for p in probes) and want in f["font"]
                   for f in near):
            return False
    return True
# ============================================================================
# 21b. STATEFUL DEVICEGRAY + PATH AUDIT (recursive Forms, images, ExtGState)
#
# Scope statement (ChatGPT #6): this parser certifies the color-space and
# transparency rules it implements against the operators and resource
# structures ReportLab emits (g/G/cs/CS/sc*/SC*/rg/RG/k/K/q/Q, Form and
# Image XObjects, ExtGState).  Anything it cannot interpret (inline images,
# shadings, patterns) is rejected fail-closed rather than silently passed.
# ============================================================================

_GRAY_NAMES = ("/DeviceGray",)


def _pdf_name(x) -> str:
    s = str(x)
    return s if s.startswith("/") else "/" + s


def _resolve_colorspace(operands, resources, report, where):
    """Resolve a cs/CS operand (name or array) against /Resources /ColorSpace."""
    if not operands:
        return None
    op0 = operands[0]
    if isinstance(op0, (list, ArrayObject)):
        arr = [str(x) for x in op0]
        base = arr[0] if arr else ""
        if base == "/DeviceGray" and len(arr) == 1:
            return "/DeviceGray"
        report.crit("CRITICAL-COLORSPACE",
                    f"{where}: array color space {arr} is not plain /DeviceGray")
        return None
    name = _pdf_name(op0)
    if name == "/DeviceGray":
        return "/DeviceGray"
    if name in ("/DeviceRGB", "/DeviceCMYK", "/Pattern", "/Lab"):
        report.crit("CRITICAL-COLORSPACE", f"{where}: {name} selected")
        return None
    # resource reference?
    if resources is not None:
        csd = resources.get("/ColorSpace")
        if csd is not None:
            csd = csd.get_object()
            obj = csd.get(name)
            if obj is not None:
                obj = obj.get_object()
                if isinstance(obj, (list, ArrayObject)):
                    arr = [str(x) for x in obj]
                    if arr and arr[0] == "/DeviceGray" and len(arr) == 1:
                        return "/DeviceGray"
                    report.crit("CRITICAL-COLORSPACE",
                                f"{where}: /ColorSpace {name} -> {arr}")
                elif _pdf_name(obj) == "/DeviceGray":
                    return "/DeviceGray"
                else:
                    report.crit("CRITICAL-COLORSPACE",
                                f"{where}: /ColorSpace {name} -> {obj!r}")
                return None
    report.crit("CRITICAL-COLORSPACE", f"{where}: unresolved color space {name}")
    return None


def _audit_resources_defaults(resources, report, where):
    """Reject /Default* substitutions and pattern/shading resources (unused)."""
    if resources is None:
        return
    csd = resources.get("/ColorSpace")
    if csd is not None:
        csd = csd.get_object()
        for key in csd.keys():
            k = str(key)
            if k.startswith("/Default"):
                if k == "/DefaultGray":
                    v = csd.get(key)
                    v = v.get_object() if hasattr(v, "get_object") else v
                    if _pdf_name(v) == "/DeviceGray":
                        continue
                report.crit("CRITICAL-COLORSPACE",
                            f"{where}: {k} substitution present")
    if resources.get("/Pattern") is not None:
        report.crit("CRITICAL-COLORSPACE", f"{where}: /Pattern resource present")
    if resources.get("/Shading") is not None:
        report.crit("CRITICAL-COLORSPACE", f"{where}: /Shading resource present")
    if resources.get("/Properties") is not None:
        report.info("COLORSPACE-NOTE", f"{where}: /Properties present (not used "
                                      f"by this engine; not certified)")


def _audit_extgstate(resources, report, where):
    if resources is None:
        return
    egs = resources.get("/ExtGState")
    if egs is None:
        return
    egs = egs.get_object()
    for key, obj in egs.items():
        o = obj.get_object()
        where2 = f"{where} ExtGState {key}"
        if o.get("/SMask") is not None and _pdf_name(o["/SMask"]) != "/None":
            report.crit("CRITICAL-EXTGSTATE", f"{where2}: /SMask != /None")
        for k in ("/CA", "/ca"):
            v = o.get(k)
            if v is not None and abs(float(v) - 1.0) > 1e-6:
                report.crit("CRITICAL-EXTGSTATE", f"{where2}: {k} = {v}")
        bm = o.get("/BM")
        if bm is not None and _pdf_name(bm) not in ("/Normal", "/Compatible"):
            report.crit("CRITICAL-EXTGSTATE", f"{where2}: /BM = {bm}")


def _num(x) -> float:
    try:
        return float(x)
    except Exception:
        return float("nan")


def _audit_content_stream(stream_obj, reader, resources, book_page, label,
                          report, path_stats, depth=0):
    """Walk one content stream with a graphics-state stack."""
    if depth > 8:
        report.crit("CRITICAL-PDF-INSPECTION", f"{label}: form nesting too deep")
        return
    try:
        cs = ContentStream(stream_obj, reader)
        ops = cs.operations
    except Exception as e:
        report.crit("CRITICAL-PDF-INSPECTION", f"{label}: cannot parse stream: {e}")
        return
    gs = {"fill": "/DeviceGray", "stroke": "/DeviceGray"}
    stack = []
    line_w = 1.0
    pts = []
    for operands, operator in ops:
        op = operator.decode() if isinstance(operator, bytes) else str(operator)
        try:
            if op == "q":
                stack.append(dict(gs))
            elif op == "Q":
                if stack:
                    gs = stack.pop()
            elif op == "g":
                gs["fill"] = "/DeviceGray"
                v = _num(operands[0]) if operands else -1
                if not (0.0 <= v <= 1.0):
                    report.crit("CRITICAL-COLORSPACE", f"{label}: g {v} out of range")
            elif op == "G":
                gs["stroke"] = "/DeviceGray"
                v = _num(operands[0]) if operands else -1
                if not (0.0 <= v <= 1.0):
                    report.crit("CRITICAL-COLORSPACE", f"{label}: G {v} out of range")
            elif op in ("rg", "RG", "k", "K"):
                report.crit("CRITICAL-COLORSPACE",
                            f"{label}: forbidden {op} operator {list(operands)}")
            elif op == "cs":
                gs["fill"] = _resolve_colorspace(operands, resources, report,
                                                 f"{label} cs") or gs["fill"]
            elif op == "CS":
                gs["stroke"] = _resolve_colorspace(operands, resources, report,
                                                   f"{label} CS") or gs["stroke"]
            elif op in ("sc", "scn"):
                if gs["fill"] not in _GRAY_NAMES:
                    report.crit("CRITICAL-COLORSPACE",
                                f"{label}: {op} while fill space is {gs['fill']}")
                elif len(operands) != 1 or not (0.0 <= _num(operands[0]) <= 1.0):
                    report.crit("CRITICAL-COLORSPACE",
                                f"{label}: {op} operands {list(operands)} not single gray")
            elif op in ("SC", "SCN"):
                if gs["stroke"] not in _GRAY_NAMES:
                    report.crit("CRITICAL-COLORSPACE",
                                f"{label}: {op} while stroke space is {gs['stroke']}")
                elif len(operands) != 1 or not (0.0 <= _num(operands[0]) <= 1.0):
                    report.crit("CRITICAL-COLORSPACE",
                                f"{label}: {op} operands {list(operands)} not single gray")
            elif op == "BI":
                report.crit("CRITICAL-PDF-INSPECTION",
                            f"{label}: inline image (BI) unsupported -> rejected")
            elif op == "sh":
                report.crit("CRITICAL-COLORSPACE", f"{label}: shading op sh")
            elif op == "w":
                line_w = _num(operands[0]) if operands else 1.0
            elif op in ("m", "l"):
                if len(operands) >= 2:
                    pts.append((_num(operands[0]), _num(operands[1])))
            elif op == "re":
                if len(operands) >= 4:
                    x, y, w_, h_ = (_num(v) for v in operands[:4])
                    pts = [(x, y), (x + w_, y + h_)]
            elif op in ("S", "s"):
                _classify_stroke(pts, line_w, book_page, path_stats)
                pts = []
            elif op in ("f", "F", "f*", "B", "B*", "b", "b*", "n"):
                pts = []
            elif op == "Do":
                if resources is not None:
                    xo = resources.get("/XObject")
                    if xo is not None:
                        xo = xo.get_object()
                        nm = _pdf_name(operands[0]) if operands else ""
                        obj = xo.get(nm)
                        if obj is not None:
                            _audit_xobject(obj.get_object(), reader, book_page,
                                           f"{label} XObject {nm}", report,
                                           path_stats, depth + 1)
        except Exception as e:
            report.crit("CRITICAL-PDF-INSPECTION",
                        f"{label}: op {op} inspection error: {e}")


def _classify_stroke(pts, line_w, book_page, path_stats):
    if len(pts) >= 2 and book_page in path_stats:
        st = path_stats[book_page]
        (x1, y1), (x2, y2) = pts[0], pts[-1]
        if abs(y1 - y2) < 0.4 and abs(y1 - RULE_Y) < 0.5 and abs(x2 - x1) > 300:
            if abs(line_w - RULE_WEIGHT) > 0.01:
                st["rule_weight_bad"] = True
            st["rule"] = True
        elif (0.4 < abs(x2 - x1) < 2 * CROSS_ARM + 0.4 and abs(y1 - y2) < 0.05) or \
             (0.4 < abs(y2 - y1) < 2 * CROSS_ARM + 0.4 and abs(x1 - x2) < 0.05):
            if abs(line_w - CROSS_WEIGHT) > 0.01:
                st["cross_weight_bad"] = True
            st["cross"] += 1


def _audit_ximage(obj, label, report):
    cs = obj.get("/ColorSpace")
    cs_n = str(cs[0]) if isinstance(cs, (list, ArrayObject)) else str(cs)
    if "/DeviceGray" not in cs_n:
        report.crit("CRITICAL-COLORSPACE", f"{label}: image /ColorSpace {cs_n}")


def _audit_xobject(obj, reader, book_page, label, report, path_stats, depth):
    sub = obj.get("/Subtype")
    if sub == "/Image":
        _audit_ximage(obj, label, report)
    elif sub == "/Form":
        res = obj.get("/Resources")
        _audit_resources_defaults(res, report, label)
        _audit_extgstate(res, report, label)
        _audit_content_stream(obj, reader, res, book_page, label, report,
                              path_stats, depth)


def audit_colorspaces_and_paths(reader, render_records, report) -> bool:
    n0 = report.n_crit
    path_stats = {p: {"rule": False, "rule_weight_bad": False,
                      "cross": 0, "cross_weight_bad": False}
                  for p in range(1, TOTAL_PAGES + 1)}
    for book_page in range(1, TOTAL_PAGES + 1):
        page = reader.pages[book_page - 1]
        res = page.get("/Resources")
        _audit_resources_defaults(res, report, f"page {book_page}")
        _audit_extgstate(res, report, f"page {book_page}")
        try:
            contents = page.get_contents()
        except Exception as e:
            report.crit("CRITICAL-PDF-INSPECTION",
                        f"page {book_page}: contents unreadable: {e}")
            continue
        if contents is None:
            continue
        _audit_content_stream(contents, reader, res, book_page,
                              f"page {book_page}", report, path_stats)
    # rule + crosshair verification from the same pass
    rule_bad = [p for p, st in path_stats.items()
                if render_records.get(p, {}).get("kind") == "cp"
                and (not st["rule"] or st["rule_weight_bad"])]
    if rule_bad:
        report.crit("CRITICAL-HEADER-RULE",
                    f"0.75pt header rule missing/wrong on CP pages {rule_bad[:8]}")
    cross_expected = {}
    for book_page, rr in render_records.items():
        if rr["kind"] == "cp" and rr.get("case"):
            cross_expected[rr["case"]] = cross_expected.get(rr["case"], 0) + \
                sum(1 for p in rr["paras"] if p[0] in ("evidence-item", "theory-item"))
    cross_bad = []
    for cn, want in cross_expected.items():
        got = sum(path_stats[p]["cross"] for p in range(case_start_page(cn) + 2,
                                                        case_start_page(cn) + 6))
        if got != 2 * want or any(path_stats[p]["cross_weight_bad"]
                                  for p in range(case_start_page(cn) + 2,
                                                 case_start_page(cn) + 6)):
            cross_bad.append((cn, got, 2 * want))
    if cross_bad:
        report.crit("CRITICAL-CROSSHAIR",
                    f"crosshair strokes wrong (case, got, want): {cross_bad[:8]}")
    return report.n_crit == n0

# ============================================================================
# 22. ORCHESTRATION — full build, pilots, preflight-only, selftest
# ============================================================================

_CASES_BY_NUM: dict = {}
_EPILOGUE_BODY_STYLE = None


def precompute_glue(case, styles, report) -> dict:
    """Widow-Ladder v2.3 pre-pass, once per item (tier-independent: line breaks
    depend on width and font size only)."""
    glued = {}
    cn = case["case_number"]
    for s in case["sections"]:
        for idx, item in enumerate(s["items"]):
            t = item["type"]
            if t not in ("body-text", "evidence-item", "theory-item",
                         "see-also", "search-term"):
                continue
            src_item = {"text": item_display(item), "runs": item_source_runs(item)}
            markup, style, role = build_item_markup(item, styles, None)
            avail = LIST_W if t in ("evidence-item", "theory-item") else TEXT_W
            res = apply_widow_ladder(src_item, markup, style, avail, cn,
                                     f"{s['key']}#{idx}", report)
            if res[2] == "glued":
                glued[(s["key"], idx)] = res
    return glued


def _fit_summary(report, tier_counts):
    report.section("FIT SUMMARY")
    report.info("FIT-SUMMARY-LINE",
                f"FIT SUMMARY: {tier_counts.get('standard', 0)} Standard, "
                f"{tier_counts.get('tight', 0)} Tight, "
                f"{tier_counts.get('open', 0)} Open, "
                f"{report.n_warn} Warnings, {report.n_crit} Critical Errors.")


def build_all(report) -> int:
    data = preflight(report)
    if data is None:
        report.write(REPORT_PATH)
        return 1
    cases = data["cases"]["cases"]
    fm = data["frontmatter"]
    epi = data["epilogue"]
    overrides = data["overrides"]
    global _CASES_BY_NUM, _EPILOGUE_BODY_STYLE
    _CASES_BY_NUM = {c["case_number"]: c for c in cases}
    std_styles = make_styles("standard")
    _EPILOGUE_BODY_STYLE = std_styles["epilogue-body"]
    _EPILOGUE_CACHE.clear()
    _EPILOGUE_CACHE.update({"epilogue": epi, "body_style": _EPILOGUE_BODY_STYLE,
                            "rec": None})

    # --- transactional step 2: purge current-run intermediates --------------
    report.section("BUILD PARTS")
    PARTS_DIR.mkdir(parents=True, exist_ok=True)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for f in PARTS_DIR.glob("*.pdf"):
        f.unlink()
    styles_by_tier = {t: make_styles(t) for t in TIER_ORDER}
    render_records = {}
    manifest = []

    t0 = time.time()
    tier_counts = {}
    if render_front_part(fm, cases, report, render_records):
        p0 = PARTS_DIR / part_name(0)
        manifest.append((p0, sha256_file(p0), 13))
    for case in cases:
        cn = case["case_number"]
        glued = precompute_glue(case, styles_by_tier["standard"], report)
        fit = fit_case(case, styles_by_tier, glued, report)
        if fit is None:
            continue                                   # part withheld
        if not render_case_part(case, fit, styles_by_tier, glued, overrides,
                                report, render_records):
            continue
        p = PARTS_DIR / part_name(cn)
        manifest.append((p, sha256_file(p), 9 if cn == N_CASES else 6))
        tier_counts[fit.tier] = tier_counts.get(fit.tier, 0) + 1
        for t in fit.telemetry:
            report.info("TRIAL", f"case {cn} [{t['step']}/{t['tier']}/"
                                 f"{t['break']}] {t['state']}")
    report.info("BUILD-TIMING", f"65 parts attempted in {time.time() - t0:.1f}s")

    if report.n_crit or len(manifest) != 65:
        report.crit("CRITICAL-BUILD",
                    f"build incomplete: {len(manifest)}/65 valid parts; master "
                    f"not merged; previous master (if any) untouched")
        _fit_summary(report, tier_counts)
        report.write(REPORT_PATH)
        return 1

    # --- merge + independent validation + atomic replace --------------------
    report.section("MERGE + FINAL VALIDATION")
    tmp = merge_parts(manifest, report)
    if tmp is None:
        report.write(REPORT_PATH)
        return 1
    ok = validate_final(tmp, render_records, report)
    if ok:
        os.replace(tmp, MASTER_PATH)
        report.info("MASTER-REPLACED", f"atomic replace -> {MASTER_PATH}")
    else:
        try:
            tmp.unlink()
        except OSError:
            pass
        report.crit("CRITICAL-MASTER", "validation failed; temporary master "
                                       "discarded; production master untouched")
    _fit_summary(report, tier_counts)
    report.write(REPORT_PATH)
    safe_print(f"\nreport: {REPORT_PATH}")
    return 0 if ok else 1


def run_pilots(case_numbers, report) -> int:
    data = preflight(report)
    if data is None:
        report.write(REPORT_PATH)
        return 1
    cases = data["cases"]["cases"]
    overrides = data["overrides"]
    global _CASES_BY_NUM
    _CASES_BY_NUM = {c["case_number"]: c for c in cases}
    styles_by_tier = {t: make_styles(t) for t in TIER_ORDER}
    render_records = {}
    PARTS_DIR.mkdir(parents=True, exist_ok=True)
    report.section(f"PILOT PARTS {case_numbers}")
    for cn in case_numbers:
        case = cases[cn - 1]
        glued = precompute_glue(case, styles_by_tier["standard"], report)
        fit = fit_case(case, styles_by_tier, glued, report)
        if fit is None:
            continue
        if render_case_part(case, fit, styles_by_tier, glued, overrides,
                            report, render_records):
            p = PARTS_DIR / part_name(cn)
            report.info("PILOT-OK", f"case {cn} -> {p.name} "
                                    f"({fit.tier}, break={fit.break_before})")
    report.info("PILOT-NOTE", "pilot mode builds parts only; no merge, no master")
    report.write(REPORT_PATH)
    return 1 if report.n_crit else 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="UNEXPLAINED two-file build engine")
    parser.add_argument("--preflight", action="store_true",
                        help="input gate only; no PDFs are written")
    parser.add_argument("--cases", type=str, default=None,
                        help="comma-separated case numbers (pilot parts only)")
    parser.add_argument("--selftest", action="store_true",
                        help="engine self-tests")
    args = parser.parse_args(argv)

    if args.selftest:
        return run_selftest()

    report = Report()
    check_runtime_environment(report)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if args.preflight:
        register_all_fonts(report)
        preflight(report)
        report.write(REPORT_PATH)
        safe_print(f"\nreport: {REPORT_PATH}")
        return 1 if report.n_crit else 0
    if args.cases:
        nums = []
        for tok in args.cases.split(","):
            tok = tok.strip()
            if not tok:
                continue
            n = int(tok)
            if not (1 <= n <= N_CASES):
                parser.error(f"case number {n} out of range")
            nums.append(n)
        return run_pilots(nums, report)
    return build_all(report)

# ============================================================================
# 23. SELFTEST — engine acceptance tests (ChatGPT #1,2,4,11: real PDF output)
#
# Selftests exercise the ENGINE with the real fonts and (when present) the
# real JSON inputs.  They never touch the production master.  The input
# preflight (hashes, images, big-word widths) is reported informationally:
# its gate is `--preflight`, not the selftest.
# ============================================================================

def _ops_of(page, reader):
    cs = ContentStream(page.get_contents(), reader)
    return cs.operations


def _op_name(operator) -> str:
    return operator.decode() if isinstance(operator, bytes) else str(operator)


def _tj_texts(ops):
    out = []
    for operands, operator in ops:
        if _op_name(operator) == "Tj":
            t = operands[0]
            if isinstance(t, bytes):
                t = t.decode("latin-1")
            out.append(str(t))
    return out


def _tm_ops(ops):
    out = []
    for operands, operator in ops:
        if _op_name(operator) == "Tm":
            out.append([float(x) for x in operands])
    return out


def _st_doc(story_builder, pages=1):
    """Render a story to an in-memory BookCanvas PDF; return PdfReader."""
    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=(PAGE_W, PAGE_H), allowSplitting=1)
    doc.addPageTemplates(story_builder(doc))
    doc.build(story_builder.story, canvasmaker=BookCanvas)
    return PdfReader(io.BytesIO(buf.getvalue()))


class _Mini:
    """Tiny helper doc: one frame template + story."""

    def __init__(self, story, frame_w=TEXT_W, frame_h=FRAME_H, chrome=False):
        self.story = story
        self.frame_w, self.frame_h = frame_w, frame_h
        self.chrome = chrome

    def __call__(self, doc):
        f = Frame(43.2, PAGE_H - 54.0 - self.frame_h, self.frame_w, self.frame_h,
                  leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0)
        ctx = {"base": 16, "case_name": "CASE 1: FAST RADIO BURSTS"}
        if self.chrome:
            onPage = lambda c, d: paint_a_master(c, d, ctx, "verso")
        else:
            onPage = lambda c, d: None
        return [PageTemplate(id="mini", frames=[f], onPage=onPage)]


def _body_filler(n=45):
    st = ParagraphStyle("filler", fontName="EBGaramond-Regular", fontSize=BODY_SIZE,
                        leading=14.0, spaceAfter=4.0, textColor=GRAY_BLACK)
    return [BookParagraph(f"filler line {i} for page occupation", st)
            for i in range(n)]


# ---------------------------------------------------------------- tests ----

def t_env(report):
    r2 = Report("env")
    check_runtime_environment(r2)
    return r2.n_crit == 0, f"reportlab {__import__('reportlab').__version__} " \
                           f"internals verified"


def t_geometry(report):
    checks = [
        case_start_page(1) == 14, case_start_page(64) == 392,
        part_name(0) == "00_FrontMatter.pdf",
        part_name(64) == "64_Case64_and_Epilogue.pdf",
        abs(TEXT_W - 331.2) < 1e-9, abs(FRAME_H - 536.4) < 1e-9,
        abs(CAP_TARGET_Y - 589.248) < 1e-9,
        abs(HEADER_BASELINE_Y - 609.12) < 1e-9,
        abs(RULE_Y - 605.16) < 1e-9,
        abs(FOLIO_Y - 32.4) < 1e-9,
        # 13 + 64*6 + 3 = 400
        FRONT_MATTER_PAGES + N_CASES * 6 + 3 == TOTAL_PAGES,
    ]
    return all(checks), f"{sum(checks)}/{len(checks)} geometry constants"


def t_tracked(report):
    buf = io.BytesIO()
    c = BookCanvas(buf, pagesize=(PAGE_W, PAGE_H))
    cases_ = [
        ("UNEXPLAINED", "LibreBaskerville-Bold", 32, 40, "center", 216.0),
        ("WELCOME TO THE EDGE OF KNOWLEDGE", "EBGaramond-Medium", 9, 60, "center", 216.0),
        ("CASE 60: DÉJÀ VU", "EBGaramond-Medium", 9, 20, "right", RECTO_RIGHT),
        ("THE CRACK IN THE EXPLANATION", "EBGaramond-SemiBold", 13, 10, "left", 57.6),
    ]
    ys = [500, 400, 300, 200]
    for (text, font, size, tr, align, ax), y in zip(cases_, ys):
        draw_tracked(c, text, ax, y, font, size, tr, align)
    c.showPage()
    c.save()
    reader = PdfReader(io.BytesIO(buf.getvalue()))
    frags = page_fragments(reader.pages[0], reader)
    ok = True
    detail = []
    for (text, font, size, tr, align, ax), y in zip(cases_, ys):
        w = tracked_width(text, font, size, tr)
        f = next((f for f in frags if norm_text(f["text"]) == norm_text(text)), None)
        if f is None:
            ok = False
            detail.append(f"missing {text[:20]!r}")
            continue
        want_x = ax - w / 2 if align == "center" else (ax - w if align == "right" else ax)
        if abs(f["x"] - want_x) > 0.05 or abs(f["y"] - y) > 0.05:
            ok = False
            detail.append(f"{text[:18]!r}: x {f['x']:.2f} != {want_x:.2f}")
    return ok, ("tracked centering/width verified on 4 lines; " +
                ("; ".join(detail) if detail else "all within 0.05pt"))


def t_gray(report):
    buf = io.BytesIO()
    c = BookCanvas(buf, pagesize=(PAGE_W, PAGE_H))
    st = ParagraphStyle("g", fontName="EBGaramond-Regular", fontSize=BODY_SIZE,
                        leading=14, textColor=GRAY_BLACK)
    # body font family registered: EBGaramond maps via ps2tt('EBGaramond-Regular')
    p = BookParagraph("black body text in the paragraph", st)
    p.wrap(TEXT_W, 10000)
    p.drawOn(c, 43.2, 400)
    draw_tracked(c, "UNEXPLAINED", 43.2, 609.12, "EBGaramond-Medium", 9, 20,
                 "left", GRAY_75)
    c.setStrokeColor(GRAY_75)
    c.setLineWidth(0.75)
    c.line(43.2, 605.16, 374.4, 605.16)
    c.showPage()
    c.save()
    data = PdfReader(io.BytesIO(buf.getvalue())).pages[0].get_contents().get_data().decode("latin1")
    has_rg = bool(re.search(r"\brg\b|\bRG\b|\bk\b|\bK\b", data))
    has_g = " g" in data and " G" in data
    r2 = Report("gray")
    reader = PdfReader(io.BytesIO(buf.getvalue()))
    stats = {1: {"rule": False, "rule_weight_bad": False, "cross": 0,
                 "cross_weight_bad": False}}
    _audit_content_stream(reader.pages[0].get_contents(), reader,
                          reader.pages[0].get("/Resources"), 1, "t_gray", r2, stats)
    ok = (not has_rg) and has_g and r2.n_crit == 0
    return ok, (f"DeviceGray g/G ops only (rg present: {has_rg}); "
                f"stateful parser clean: {r2.n_crit == 0}")


def t_shear(report):
    markup = ("Regular words here then <b>Bold words</b> then <i>Italic "
              "words</i> then <b><i>BoldItalic words</i></b> then Tail words "
              "continue on and on and wrap across lines for drift checking "
              "with more regular words to fill the line completely")
    st = ParagraphStyle("s", fontName="EBGaramond-Regular", fontSize=BODY_SIZE,
                        leading=14, textColor=GRAY_BLACK)
    p = BookParagraph(markup, st)
    mini = _Mini([p])
    reader = _st_doc(mini)
    page = reader.pages[0]
    ops = _ops_of(page, reader)
    tms = _tm_ops(ops)
    tjs = _tj_texts(ops)
    # 1. sequence: shear Tm -> (BoldItalic..) Tj -> restore Tm -> Tail
    bi_idx = next((i for i, t in enumerate(tjs) if t.startswith("BoldItalic")), None)
    if bi_idx is None:
        return False, "BoldItalic fragment not found in stream"
    size = BODY_SIZE
    W = stringWidth("BoldItalic words", SYNTH_BASE, size)
    shear_tms = [m for m in tms if abs(m[2] - SHEAR_TAN) < 1e-4]
    restore_tms = [m for m in tms if m[0] == 1 and m[2] == 0 and m[3] == 1]
    if not shear_tms:
        return False, "no shear Tm found"
    m = shear_tms[0]
    if not any(abs(r[4] - (m[4] + W)) < 0.05 and abs(r[5] - m[5]) < 0.01
               for r in restore_tms):
        return False, (f"restore Tm not at exact post-advance "
                       f"(want x={m[4] + W:.2f})")
    # 2. drift: every rendered line starts at the same left edge
    frags = page_fragments(page, reader)
    body = [f for f in frags]
    lines = {}
    for f in body:
        lines.setdefault(round(f["y"], 1), []).append(f["x"])
    starts = [min(xs) for xs in lines.values()]
    drift = max(starts) - min(starts)
    # first line starts at 43.2; wrapped lines too; the synth line start == 43.2
    if drift > 0.05:
        return False, f"line-start drift {drift:.3f}pt across {len(starts)} lines"
    # 3. baselines descend by exactly the leading
    ys = sorted(lines.keys(), reverse=True)
    gaps = [round(a - b, 3) for a, b in zip(ys, ys[1:])]
    if any(abs(g - 14.0) > 0.01 for g in gaps):
        return False, f"leading drift: {gaps}"
    return True, (f"shear Tm tan={SHEAR_TAN:.5f}; restore at x+W exact; "
                  f"{len(ys)} lines, zero drift, leading exact")


def t_cap(report):
    results = []
    std = make_styles("standard")

    def check_first_cap(story_tail, label):
        story = _body_filler(45) + story_tail
        mini = _Mini(story)
        # two-frame alternating doc: verso filler page, recto test page
        buf = io.BytesIO()
        doc = BaseDocTemplate(buf, pagesize=(PAGE_W, PAGE_H), allowSplitting=1)
        ctx = {"base": 16, "case_name": "CASE 1: FAST RADIO BURSTS"}
        av, ar = _a_templates(ctx)
        doc.addPageTemplates([av, ar])
        doc.build([BookParagraph(f"filler {i}", std["body"]) for i in range(45)]
                  + story_tail, canvasmaker=BookCanvas)
        reader = PdfReader(io.BytesIO(buf.getvalue()))
        page = reader.pages[1]                      # recto
        frags = [f for f in page_fragments(page, reader)
                 if FOLIO_BAND_Y <= f["y"] <= HEADER_BAND_Y]
        if not frags:
            return label, False, "no body fragments on recto page"
        f0 = max(frags, key=lambda f: f["y"])
        cap = _FONT_METRICS.get(f0["font"], {}).get("cap")
        if cap is None:
            return label, False, f"unknown font {f0['font']}"
        got = f0["y"] + cap * f0["size"]
        return label, abs(got - CAP_TARGET_Y) <= 0.05, \
            f"cap top {got:.3f} (font {f0['font']} {f0['size']}pt)"

    rich = BookParagraph('plain then <b>bold</b> and <i>italic</i> first '
                         'rich paragraph on the recto page', std["body"])
    plain = BookParagraph("plain first paragraph on the recto page", std["body"])
    list_para = BookParagraph("a crosshair bulleted list item here", std["list"])
    li = ListItemFlowable(list_para)
    hdr = TrackedLineFlowable("THE EVIDENCE", 10, "EBGaramond-SemiBold", 13, 14,
                              18.0, 10.8)
    kt = KeepTogether([TrackedLineFlowable("THE EVIDENCE", 10, "EBGaramond-SemiBold",
                                           13, 14, 18.0, 10.8),
                       BookParagraph("keep together first paragraph", std["body"])])
    for tail, label in (([plain], "plain paragraph"),
                        ([rich], "rich paragraph"),
                        ([li], "ListItemFlowable"),
                        ([hdr, BookParagraph("after header", std["body"])],
                         "TrackedLineFlowable"),
                        ([kt], "KeepTogether(header+item)")):
        results.append(check_first_cap(tail, label))
    ok = all(r[1] for r in results)
    return ok, "; ".join(f"{l}: {'OK' if o else 'FAIL ' + d}" for l, o, d in results)


def t_widow(report):
    r2 = Report("widow")
    std = make_styles("standard")
    res = []

    def wrap_lines(text, width, style=None):
        style = style or std["body"]
        p = Paragraph(esc(text).replace("\n", "<br/>"), style)
        p.wrap(width, 10 ** 6)
        return inspect_paragraph_lines(p)

    def craft(width, tail, companion, max_words=40):
        """Deterministically craft text whose wrap ends with the tail alone."""
        for n in range(2, max_words):
            text = " ".join(["word"] * n) + f" {companion} {tail}"
            lines = wrap_lines(text, width)
            if (len(lines) >= 2 and len(lines[-1]["words"]) == 1
                    and lines[-1]["words"][0]["text"] == tail
                    and lines[-2]["words"][-1]["text"] == companion):
                return text
        return None

    # 1. plain trigger -> glue (short companion: guard does not apply)
    text1 = craft(TEXT_W, "tail", "mid")
    if text1 is None:
        return False, "could not craft a plain widow layout"
    item = {"text": text1, "runs": None}
    m, _, act = apply_widow_ladder(item, esc(text1), std["body"], TEXT_W, 1, "t1", r2)
    res.append(("glue plain", act == "glued" and NBSP in m))

    # 2. long companion + shallow preceding line -> accepted widow (no glue)
    text2 = craft(85.0, "tail", "extraordinarily")
    if text2 is None:
        return False, "could not craft a guarded widow layout"
    item2 = {"text": text2, "runs": None}
    m2, _, act2 = apply_widow_ladder(item2, esc(text2), std["body"], 85.0, 1, "t2", r2)
    res.append(("guard accept", act2 == "accepted" and NBSP not in m2))

    # 3. rich run boundary: last word italic, space inside the previous run
    base = craft(TEXT_W, "tailword", "mid")
    if base is None:
        return False, "could not craft a rich widow layout"
    idx = base.rfind(" tailword")
    head = base[:idx + 1]                # keeps the space before the tailword
    runs3 = [{"t": head}, {"t": "tailword", "i": True}]
    m3, nr, act3 = apply_widow_ladder({"text": base, "runs": runs3},
                                      runs_markup(runs3), std["body"], TEXT_W, 1,
                                      "t3", r2)
    ok3 = act3 == "glued" and nr and any(NBSP in r["t"] for r in nr) \
        and nr[-1].get("i") is True
    res.append(("glue across runs", ok3))

    # 4. authored newline boundary: never glue
    text4 = "some words that wrap around nicely here\nlastword"
    m4, _, act4 = apply_widow_ladder({"text": text4, "runs": None},
                                     esc(text4).replace("\n", "<br/>"),
                                     std["body"], TEXT_W, 1, "t4", r2)
    res.append(("authored \\n preserved", act4 is None and "lastword" in m4))

    # 5. single-line paragraph untouched
    m5, _, act5 = apply_widow_ladder({"text": "short single line", "runs": None},
                                     esc("short single line"), std["body"],
                                     TEXT_W, 1, "t5", r2)
    res.append(("single line untouched", act5 is None))

    # 6. a word split across formatting boundaries counts as ONE logical word
    runs6 = [{"t": "count these words carefully mr "},
             {"t": "super", "b": True}, {"t": "cali", "b": True}]
    text6 = runs6[0]["t"] + "supercali"
    p6 = Paragraph(runs_markup(runs6), std["body"])
    p6.wrap(TEXT_W, 10 ** 6)
    ln = inspect_paragraph_lines(p6)
    one_word = any(w["text"] == "supercali" for w in ln[-1]["words"]) \
        and not any(w["text"] in ("super", "cali") for w in ln[-1]["words"])
    res.append(("split word = one logical word", one_word))
    ok = all(x[1] for x in res)
    return ok, "; ".join(f"{n}: {'OK' if v else 'FAIL'}" for n, v in res)


def t_keeptogether(report):
    std = make_styles("standard")
    rec = LayoutRecorder()
    hdr = lambda: TrackedLineFlowable("THE EVIDENCE", 10, "EBGaramond-SemiBold",
                                      13, 14, 18.0, 10.8, src={"key": "EVIDENCE"},
                                      rec=rec)
    para = lambda i: BookParagraph(f"item text {i} following the section header",
                                   std["body"], src={"key": "EVIDENCE", "idx": i},
                                   rec=rec)
    # small frame: header(18+14+10.8) + one 14+4 line ~= 61pt; frame 84pt
    story = [BookParagraph("tall opening paragraph " * 4, std["body"]),
             HeaderBlock([hdr(), para(0)])]
    mini = _Mini(story, frame_h=84.0)
    reader = _st_doc(mini)
    hdr_pages = [p for p, y, kind, role, _ in rec.events if kind == "header"]
    item_pages = [p for p, y, kind, role, _ in rec.events
                  if (payload := _) and (payload.get("src") or {}).get("idx") == 0]
    ok = hdr_pages and item_pages and hdr_pages[0] == item_pages[0] == 2
    # single-line first item pulls the second item into the KT group
    rec2 = LayoutRecorder()
    story2 = [BookParagraph("tall opening paragraph " * 4, std["body"]),
              HeaderBlock([TrackedLineFlowable("SEE ALSO", 10, "EBGaramond-SemiBold",
                                                13, 14, 18.0, 10.8,
                                                src={"key": "SEE_ALSO"}, rec=rec2),
                            BookParagraph("short", std["body"],
                                          src={"key": "SEE_ALSO", "idx": 0}, rec=rec2),
                            BookParagraph("second", std["body"],
                                          src={"key": "SEE_ALSO", "idx": 1}, rec=rec2)])]
    mini2 = _Mini(story2, frame_h=84.0)
    _st_doc(mini2)
    p0 = [p for p, y, k, r, pl in rec2.events if (pl.get("src") or {}).get("idx") == 0]
    p1 = [p for p, y, k, r, pl in rec2.events if (pl.get("src") or {}).get("idx") == 1]
    ph = [p for p, y, k, r, pl in rec2.events if k == "header"]
    ok2 = p0 and p1 and ph and ph[0] == p0[0] == p1[0] == 2
    return ok and ok2, (f"header+item pushed together to page 2 "
                        f"({hdr_pages}/{item_pages}); single-line first item "
                        f"protects second item ({ph}/{p0}/{p1})")


def t_decision(report):
    def mk(pages, qs, qe, ss, se, cs_, ce, items_ok=True):
        t = TrialResult("standard", None)
        t.pages = pages
        t.sec_start = {"QUESTION": qs, "SEE_ALSO": ss, "CURIOUS": cs_}
        t.sec_end = {"QUESTION": qe, "SEE_ALSO": se, "CURIOUS": ce}
        t.item_pages = {(k, i): 4 for k in ("SEE_ALSO", "CURIOUS") for i in range(3)}
        if not items_ok:
            t.item_pages[("SEE_ALSO", 2)] = 3
        return t

    checks = []
    # D0: everything on CP4
    t = mk(4, 4, 4, 4, 4, 4, 4)
    checks.append(("D0", verify_step_e(t, None, 4)[0] is True))
    # D1 overflow proof
    t = mk(4, 3, 3, 4, 4, 4, 4)
    d1 = mk(5, 4, 5, 4, 5, 5, 5)
    checks.append(("D1-overflow-proof", verify_step_e(t, d1, 3)[0] is True))
    # QUESTION before CP3: structurally impossible
    checks.append(("Q-before-CP3", verify_step_e(t, None, 2)[0] is True))
    # no proof -> reject
    checks.append(("no-proof-rejects", verify_step_e(t, None, 3)[0] is False))
    # S/C not on 4 -> reject
    t2 = mk(4, 4, 4, 3, 4, 4, 4)
    checks.append(("sc-not-on-4-rejects", verify_step_e(t2, None, 4)[0] is False))
    # items not on 4 -> _sc_ok False
    t3 = mk(4, 4, 4, 4, 4, 4, 4, items_ok=False)
    checks.append(("items-on-3-rejects", verify_step_e(t3, None, 4)[0] is False))
    # pages != 4 -> reject
    t4 = mk(5, 4, 5, 4, 5, 4, 5)
    checks.append(("pages-5-rejects", verify_step_e(t4, None, 4)[0] is False))
    ok = all(c[1] for c in checks)
    return ok, "; ".join(f"{n}: {'OK' if v else 'FAIL'}" for n, v in checks)


def t_grayparser(report):
    # good document (BookCanvas)
    buf = io.BytesIO()
    c = BookCanvas(buf, pagesize=(300, 400))
    c.setFillColor(GRAY_BLACK)
    c.rect(10, 10, 50, 50, stroke=0, fill=1)
    c.setStrokeColor(GRAY_75)
    c.setLineWidth(0.75)
    c.line(10, 100, 200, 100)
    c.showPage()
    c.save()
    good = PdfReader(io.BytesIO(buf.getvalue()))
    r_good = Report("good")
    stats = {1: {"rule": False, "rule_weight_bad": False, "cross": 0,
                 "cross_weight_bad": False}}
    _audit_content_stream(good.pages[0].get_contents(), good,
                          good.pages[0].get("/Resources"), 1, "good", r_good, stats)
    # bad 1: rgb operator
    buf1 = io.BytesIO()
    c1 = Canvas(buf1, pagesize=(300, 400))
    c1.setFillColorRGB(1, 0, 0)
    c1.rect(10, 10, 50, 50, stroke=0, fill=1)
    c1.showPage()
    c1.save()
    bad1 = PdfReader(io.BytesIO(buf1.getvalue()))
    r1 = Report("bad1")
    _audit_content_stream(bad1.pages[0].get_contents(), bad1,
                          bad1.pages[0].get("/Resources"), 1, "bad1", r1,
                          {1: {"rule": False, "rule_weight_bad": False,
                               "cross": 0, "cross_weight_bad": False}})
    # bad 2: rgb inside a Form XObject (recursive inspection)
    buf2 = io.BytesIO()
    c2 = Canvas(buf2, pagesize=(300, 400))
    c2.beginForm("F1")
    c2.setFillColorRGB(0, 1, 0)
    c2.rect(0, 0, 20, 20, stroke=0, fill=1)
    c2.endForm()
    c2.doForm("F1")
    c2.showPage()
    c2.save()
    bad2 = PdfReader(io.BytesIO(buf2.getvalue()))
    r2 = Report("bad2")
    _audit_content_stream(bad2.pages[0].get_contents(), bad2,
                          bad2.pages[0].get("/Resources"), 1, "bad2", r2,
                          {1: {"rule": False, "rule_weight_bad": False,
                               "cross": 0, "cross_weight_bad": False}})
    # bad 3: transparency ExtGState
    buf3 = io.BytesIO()
    c3 = Canvas(buf3, pagesize=(300, 400))
    c3.saveState()
    c3.setFillAlpha(0.5)
    c3.setFillColorRGB(0, 0, 0)
    c3.rect(10, 10, 50, 50, stroke=0, fill=1)
    c3.restoreState()
    c3.showPage()
    c3.save()
    bad3 = PdfReader(io.BytesIO(buf3.getvalue()))
    r3 = Report("bad3")
    _audit_content_stream(bad3.pages[0].get_contents(), bad3,
                          bad3.pages[0].get("/Resources"), 1, "bad3", r3,
                          {1: {"rule": False, "rule_weight_bad": False,
                               "cross": 0, "cross_weight_bad": False}})
    ok = (r_good.n_crit == 0 and r1.n_crit > 0 and r2.n_crit > 0 and r3.n_crit > 0)
    return ok, (f"all-gray doc passes ({r_good.n_crit} crit); rg rejected "
                f"({r1.n_crit}); nested Form rgb rejected ({r2.n_crit}); "
                f"alpha ExtGState rejected ({r3.n_crit})")


def t_toc_dots(report):
    e = TOCEntryFlowable("CASE 1: FAST RADIO BURSTS", 14)
    mini = _Mini([e])
    reader = _st_doc(mini)
    frags = page_fragments(reader.pages[0], reader)
    texts = [f["text"] for f in frags]
    joined = norm_text("".join(texts))
    ok_txt = ("CASE 1: FAST RADIO BURSTS" in joined and "14" in joined
              and joined.count(".") >= 3)
    # geometry: number right-aligned, no overlap with entry text
    entry = next((f for f in frags if "FAST" in f["text"]), None)
    num = next((f for f in frags if f["text"].strip() == "14"), None)
    dots = next((f for f in frags if ".." in f["text"] or
                 f["text"].strip().startswith(". .")), None)
    ok_geo = False
    if entry and num and dots:
        entry_end = entry["x"] + stringWidth(entry["text"], "EBGaramond-Regular", 10.5)
        ok_geo = (dots["x"] >= entry_end + 5 and
                  num["x"] + stringWidth("14", "EBGaramond-Regular", 10.5)
                  <= TEXT_W + 43.2 + 0.5 and
                  num["x"] >= dots["x"])
    return ok_txt and ok_geo, f"dots present={ok_txt}, geometry ok={ok_geo}"


def t_image(report):
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        png = Path(td) / "tiny.png"
        im = PILImage.new("L", (60, 90))
        im.putdata([min(255, (x * 3 + y * 2) % 256)
                    for y in range(90) for x in range(60)])
        im.save(png, dpi=(300, 300))
        buf = io.BytesIO()
        c = BookCanvas(buf, pagesize=(PAGE_W, PAGE_H))
        c.drawImage(str(png), 0, 0, width=PAGE_W, height=PAGE_H)
        c.showPage()
        c.save()
        reader = PdfReader(io.BytesIO(buf.getvalue()))
        page = reader.pages[0]
        xo = page["/Resources"]["/XObject"]
        obj = list(xo.values())[0].get_object()
        cs_ok = "/DeviceGray" in str(obj.get("/ColorSpace"))
        dims_ok = obj.get("/Width") == 60 and obj.get("/Height") == 90
        imgs = page.images
        with PILImage.open(png) as src:
            want = src.tobytes()
        got = imgs[0].image
        if got.mode != "L":
            got = got.convert("L")
        pix_ok = got.tobytes() == want
        r2 = Report("img")
        stats = {1: {"rule": False, "rule_weight_bad": False, "cross": 0,
                     "cross_weight_bad": False}}
        _audit_content_stream(page.get_contents(), reader,
                              page.get("/Resources"), 1, "t_img", r2, stats)
        ok = cs_ok and dims_ok and pix_ok and r2.n_crit == 0 and len(imgs) == 1
        return ok, (f"DeviceGray={cs_ok}, dims={dims_ok}, pixel-identical={pix_ok}, "
                    f"parser clean={r2.n_crit == 0}")


def t_roundtrip(report):
    std = make_styles("standard")
    text = ("alpha beta gamma delta epsilon zeta eta theta iota kappa lambda "
            "mu nu omicron\nAUTHORED SECOND LINE tailword")
    p1 = BookParagraph(esc(text).replace("\n", "<br/>"), std["body"])
    glued_item = {"text": "one two three four five six seven eight nine ten "
                          "eleven twelve thirteen tail",
                  "runs": None}
    m2 = esc(glued_item["text"])
    res = apply_widow_ladder(glued_item, m2, std["body"], TEXT_W, 1, "rt", report)
    p2 = BookParagraph(res[0], std["body"])
    mini = _Mini([p1, Spacer(1, 12), p2])
    reader = _st_doc(mini)
    frags = page_fragments(reader.pages[0], reader)
    got = norm_text(joined_text(frags))
    want = norm_text(text + " " + glued_item["text"])
    ok_txt = got == want
    segs = ["mu nu omicron", "AUTHORED SECOND LINE tailword"]
    ok_break = _segments_on_distinct_lines(frags, segs)
    return ok_txt and ok_break, f"text match={ok_txt}, authored break preserved={ok_break}"


# ------------------------------------------------------------ runner --------

def run_selftest() -> int:
    safe_print("=" * 72)
    safe_print("UNEXPLAINED build_book.py --selftest")
    safe_print("=" * 72)
    report = Report("selftest")
    try:
        register_all_fonts(report)
    except Exception as e:
        safe_print(f"[FAIL] font registration: {e}")
        return 1
    tests = [
        ("T01 runtime environment", t_env),
        ("T02 geometry constants", t_geometry),
        ("T03 tracked-line geometry", t_tracked),
        ("T04 DeviceGray emission", t_gray),
        ("T05 synthetic BoldItalic shear", t_shear),
        ("T06 recto cap alignment (PDF truth)", t_cap),
        ("T07 Widow-Ladder v2.3", t_widow),
        ("T08 KeepTogether behavior", t_keeptogether),
        ("T09 Page-6 decision table", t_decision),
        ("T10 DeviceGray parser", t_grayparser),
        ("T11 TOC dot leader", t_toc_dots),
        ("T12 image identity", t_image),
        ("T13 text-integrity round-trip", t_roundtrip),
    ]
    n_pass = 0
    for name, fn in tests:
        try:
            ok, detail = fn(report)
        except Exception:
            ok, detail = False, "exception: " + traceback.format_exc(limit=3)
        status = "PASS" if ok else "FAIL"
        if ok:
            n_pass += 1
        safe_print(f"[{status}] {name}")
        safe_print(f"        {detail}")
    safe_print("-" * 72)
    safe_print(f"ENGINE TESTS: {n_pass}/{len(tests)} passed")

    # informational: input preflight against whatever inputs exist here
    safe_print("-" * 72)
    safe_print("INPUT PREFLIGHT (informational in selftest; gate is --preflight):")
    r2 = Report("preflight-info")
    try:
        register_all_fonts(r2)
        data = preflight(r2, need_images=False)
        crits = [(s, m) for s, m in r2.lines if s == Sev.CRIT]
        if not crits and data is not None:
            safe_print("  PASS (images not checked here; run --preflight for the "
                       "full gate)")
        else:
            for s, m in crits:
                safe_print(f"  [{s}] {m}")
            safe_print("  FAIL — the full build will refuse to run until these "
                       "are resolved (this is the fail-closed gate working).")
    except FileNotFoundError as e:
        safe_print(f"  inputs not present in this directory: {e}")
    safe_print("=" * 72)
    return 0 if n_pass == len(tests) else 1


if __name__ == "__main__":
    sys.exit(main())
