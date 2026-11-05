# UNEXPLAINED — Implementation Notes, Test Evidence & Limitations

**Decision taken:** Option 1 — the Gemini plan executed with ChatGPT's 14
clarifications applied, plus a small number of engine-level improvements
listed below. The architecture was not redesigned (per the locked review).

**Delivered files** (copy into `E:\python-projects\unexplained\`):

| File | Role |
|---|---|
| `build_book.py` | The only production script (preflight → 65 parts → merge → independent validation → atomic master replace) |
| `overrides.json` | Locked to the approved contract: Cases 11 & 57 at 45 pt |
| `requirements.txt` | The versions the engine was actually tested against |
| `SELFTEST_RESULTS.txt` | Real output of `python build_book.py --selftest` (exit 0) |
| `SMOKE_TEST/` | Full-pipeline verification artifacts (see below) |

---

## 1. ChatGPT review → implementation map

| # | Review point | Where it lives |
|---|---|---|
| 1 | Recto cap alignment accepted from actual PDF output | `CapAlignedFrame` calibrates `_y` once per recto page; acceptance = T06 selftest + `CRITICAL-CAP-ALIGN` in `validate_final` (text-matrix y + real OS/2 cap height == 589.248 pt ± 0.05 on **every** recto CP page of the merged book) |
| 2 | Synthetic BoldItalic visual + stream verification | `ShearTextObject` (12° `Tm` shear inside BT..ET, exact unskewed restore); T05 verifies stream sequence, exact post-advance restore, zero line-start drift, exact leading across wrapped lines. **Note: the manuscript contains zero bold+italic runs (preflight `SYNTH-UNUSED`), so this machinery is exercised by selftest only** |
| 3 | CP4 QUESTION placement needs structural proof | `fit_case()` exhaustive decision table (D0 / D1 / D2_natural / D2 / Open→Standard revert / reject); `verify_step_e()` accepts QUESTION-off-CP4 only with a *measured* D1 failure or QUESTION starting before CP3; full per-trial telemetry in the report |
| 4 | Widow ladder on reconstructed logical words | `_frag_token_stream` + `_logical_words_from_tokens` (words split across `<b>`/`<i>` boundaries count as one); NBSP inserted inside the runs structure preserving flags (T07) |
| 5 | 65 % guard on actual measurements | `apply_widow_ladder`: guard uses real rendered widths of the penultimate line minus the companion word and its space, over the style's actual available width |
| 6 | DeviceGray validator resolves resources, states limits | `_audit_content_stream`: recursive Form XObjects, resource-level `/ColorSpace` resolution (named + array), `/Default*` substitution rejection, `sc/SC/scn/SCN` under non-gray state rejection, ExtGState transparency, images; inline images/shading/patterns are **rejected fail-closed** (not silently passed). The parser certifies the structures ReportLab emits — it is not a universal PDF color certifier |
| 7 | Frozen JSON is authoritative | Expected page text comes from `cases.json`/`frontmatter.json`/`epilogue.json`; render records only locate *where* things landed; build-time `CRITICAL-MARKUP` cross-checks rendered-vs-source per item |
| 8 | Case preservation vs uppercase | `case_name`/TOC entries render with exact stored casing (already uppercase in the data); `.upper()` only on section-header roles; integrity compares role-transformed text against source |
| 9 | Image identity | Exact `CASE{NN}.png` mapping enforced in preflight; post-merge every embedded image is **pixel-compared** to its source PNG (`CRITICAL-IMAGE-IDENTITY`) |
| 10 | Exact page ownership | Part manifest (page counts 13/6×63/9) + folio == physical page number on every CP page + per-page expected-text records; pages 14–15 treated as Case 1's B-spread, not front matter |
| 11 | KeepTogether behavior verified | T08 exercises the production structures; because reportlab's `KeepTogether.split()` FrameBreak heuristic is not guaranteed, `HeaderBlock` prepends a FrameBreak deterministically whenever the group genuinely does not fit |
| 12 | ReportLab runtime recorded | `check_runtime_environment()` records versions + verifies the internals relied on (`paraFontSizeHeightOffset == 1`, `PDFTextObject._textOut/_setFont/_formatText`, Frame instance attrs, pypdf/Pillow/fontTools). Tested on **reportlab 4.4.9, pypdf 6.6.0, Pillow 11.3.0, fontTools 4.62.1, Python 3.12** |
| 13 | Transactional sequence | preflight → purge parts → build 65 (collect-all, failed parts withheld) → 65 + zero-criticals gate → tmp merge → independent validation → `os.replace` atomic master swap; the master is never touched by a failed build |
| 14 | Honest delivery | Everything claimed below was actually executed in this sandbox; the limitations section lists what was **not** executed |

## 2. Findings that need your decision

**Cases 8 and 26 big words overflow the text frame at 50 pt** (measured with
Libre Baskerville Bold against the 331.2 pt frame):

- Case 8 `An Equation`: **332.7 pt** at 50 pt (299.4 pt at 45 pt)
- Case 26 `A Settlement`: **341.4 pt** at 50 pt (307.2 pt at 45 pt)

The delivered `overrides.json` stays locked to `{11, 57}`. When you approve,
the fix is two lines in `overrides.json` **plus** the same two case numbers in
`APPROVED_OVERRIDE_CASES` (one constant at the top of `build_book.py`) —
preflight will then pass the width gate. Until then the full build correctly
refuses to run (`CRITICAL-BIGWORD-WIDTH`).

## 3. What was executed (all outputs are real)

1. `--selftest`: **13/13 engine tests pass** (see `SELFTEST_RESULTS.txt`).
2. `--preflight` against the real repo inputs: hash, JSON 1.1 structure,
   aggregates (4,267 case items — exactly the documented counts), fonts with
   per-role cmap coverage, display widths — all pass except the two big-word
   overflows above and the (here absent) images folder.
3. **Full-pipeline smoke build**: the real manuscript, real JSON, real fonts,
   64 *synthetic* grayscale placeholder images (1800×2700, mode L, 300 dpi,
   clearly watermarked "SMOKE TEST PLACEHOLDER — NOT PRODUCTION ART"), and a
   smoke-only overrides extension for Cases 8/26. Result:
   - 65/65 parts; **FIT SUMMARY: 39 Standard, 1 Tight, 24 Open, 0 Warnings, 0 Critical Errors**
   - 203 widow-glues, 0 WIDOW-ACCEPTED
   - 47 D0 / 17 D1 page-6 decisions
   - independent validation of the merged 400-page PDF: **PASS** — geometry,
     folios, headers, chrome suppression, recto cap == 589.248 pt on all 63
     recto CP pages, text integrity on all 400 pages (authored line breaks
     preserved, bold/italic run fonts verified), stranding, blanks, embedded
     fonts, 64 pixel-identical DeviceGray images, stateful DeviceGray +
     0.75 pt rule audit
   - atomic master replacement executed
   - `SMOKE_TEST/UNEXPLAINED_ENGINE_SMOKE_400p_NOT_PRODUCTION.pdf` is that
     output: open it to review everything **except** artwork. Do not upload
     it to KDP.

## 4. Limitations / not executed (per ChatGPT #14)

- **No production build**: your real `images/` folder was not in the shared
  repo, so the smoke build used synthetic placeholders. Run
  `python build_book.py` on your machine (after the Cases 8/26 approval) for
  the production master.
- **Human visual proof** of representative spreads remains mandatory; the
  programmatic visual check performed here (11/11 pages: correct ink bands)
  is not a substitute.
- **KDP Print Previewer** check not performed (needs the KDP portal).
- The stateful color parser certifies the structures ReportLab emits
  (`g/G`, `cs/CS`, XObjects, ExtGState); unknown structures fail closed
  rather than pass silently.
- `--cases` pilot mode and the `--preflight` gate were exercised; pilot
  part-rendering itself is the same code path as the full build (proven by
  the smoke run).
- Calibration constants flagged in the code (welcome-page optical center,
  copyright band, title-page stack gaps) match the accepted plan values;
  they are grouped at the top of the file for your proof-stage pass.

## 5. How to run on your machine

```
cd E:\python-projects\unexplained
.\env\Scripts\python.exe build_book.py --selftest      # engine tests
.\env\Scripts\python.exe build_book.py --preflight     # input gate
.\env\Scripts\python.exe build_book.py --cases 1,11,57 # pilot parts
.\env\Scripts\python.exe build_book.py                 # full build
```

The runtime gate verifies the ReportLab internals; if your installed 4.x
differs and fails the gate, `pip install -r requirements.txt`.
