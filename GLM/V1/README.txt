SMOKE TEST ARTIFACTS (sandbox verification - NOT production output)
===================================================================

smoke_driver.py
    Rebuilds this verification: copies the REAL manuscript + json/ + fonts/,
    generates 64 synthetic grayscale placeholder images (watermarked
    "SMOKE TEST PLACEHOLDER - NOT PRODUCTION ART"), extends the approved
    override list with Cases 8 and 26 at 45pt (the two measured overflows,
    pending your approval), and runs the full transactional build.

build_report.txt
    The actual report of that run: preflight, per-case fit telemetry,
    validation results. FIT SUMMARY: 39 Standard, 1 Tight, 24 Open,
    0 Warnings, 0 Critical Errors.

UNEXPLAINED_ENGINE_SMOKE_400p_NOT_PRODUCTION.pdf
    The 400-page result. All text, fonts, layout, folios, TOC, epilogue are
    the real book; the 64 case images are synthetic placeholders. Useful for
    reviewing everything except artwork. DO NOT upload to KDP.

The delivered ../build_book.py differs from the copy used for this smoke run
in exactly ONE line (verified by diff): APPROVED_OVERRIDE_CASES is extended
with cases 8 and 26 in the smoke copy, because those two 45pt overrides are
still pending your approval. The production file stays locked to {11, 57}.
