#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UNEXPLAINED - ingest.py  v1.1
Manuscript (DOCX) -> structured JSON + QA report.

Revised from Gemini Argon's v1 against the locked spec (Ingest Normalization
Rules v1). Architecture kept: SHA-256 fingerprint, style names as identifiers,
state-machine extraction, intermediate JSON, machine-readable QA.

PRINCIPLES
  * The DOCX supplies STYLE NAMES and TEXT only. Fonts, sizes, spacing and
    colours set in Word are ignored. Only inline bold/italic in `body-text`
    is read.
  * Text is stored AS WRITTEN. No upper-casing, no period stripping, no
    NBSPs. (All Caps etc. is the layout engine's job, applied by style.)
  * Nothing is silently dropped. Unknown, misplaced or hidden content is a
    CRITICAL error.
  * JSON is always written (diagnostics), but every JSON file carries
    `ingest_status` and `criticals`. build_book.py MUST refuse to run when
    ingest_status != "OK", when the three source_sha256 values differ, or
    when the hash no longer matches the manuscript.

USAGE
    python ingest.py                      # manuscript_styled.docx next to script
    python ingest.py --docx other.docx
    python ingest.py --allow-unfrozen     # skip the frozen-hash check (drafts)

JSON SCHEMA (schema_version 1.1)
  frontmatter.json
    frontmatter : {style: [line, ...]}   (book-name, book-name-full,
                   subtitle-name, author, copyright, dedication-line, welcome,
                   toc-title, intro-title)
    introduction: [{"type":"body-text","text":..,"runs"?:[{"t":..,"b"?:true,"i"?:true}]},
                   {"type":"stanza_break"}, ...]
  cases.json -> cases: [{
      case_number, case_name, case_topic, big_word, title_lines[],
      quote : "line1\\nline2"   (ONE block per case; "\\n" = forced line break)
      captions[], sections:[{key, header, items:[...]}] }]
      item types: body-text | evidence-item{text} | theory-item{lead,body}
                  | see-also{text,name,target,description} | search-term{text}
  epilogue.json
    title, groups: [[line, ...], ...]   (groups separated by stanza breaks)
  charset.txt            every unique character in the stored text
  charset_by_style.json  {style: "chars"}  -> lets fonts_check test each
                         style's characters against that style's own font
  qa_report.txt
"""

import argparse
import hashlib
import json
import os
import re
import sys
from collections import defaultdict

from docx import Document

VERSION = "1.1"
SCHEMA_VERSION = "1.1"

ROOT = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DOCX = os.path.join(ROOT, "manuscript_styled.docx")
JSON_DIR = os.path.join(ROOT, "json")
IMAGES_DIR = os.path.join(ROOT, "images")
QA_REPORT_PATH = os.path.join(ROOT, "qa_report.txt")

# SHA-256 of the frozen manuscript (manuscript_FINAL_v3 == manuscript_styled).
# Update ONLY when the manuscript is deliberately re-frozen.
FROZEN_SHA256 = "fc6f3047686e6cf33a47cf22debad486d5044f7f4222d6a55ef60fdfb8f80220"

# --------------------------------------------------------------------------
# Style whitelist (exact, case-sensitive) and regions
# --------------------------------------------------------------------------
FRONT_STYLES = (
    "book-name", "book-name-full", "subtitle-name", "author", "copyright",
    "dedication-line", "welcome", "toc-title",
)
ALLOWED_STYLES = set(FRONT_STYLES) | {
    "intro-title", "body-text", "case-name", "big-word", "case-title",
    "case-quote", "image-caption", "section-header", "list-item",
    "see-also", "search-term", "epilogue-title", "epilogue-body",
}
EXPECTED_FRONT_COUNTS = {
    "book-name": 1, "book-name-full": 1, "subtitle-name": 2, "author": 1,
    "copyright": 5, "dedication-line": 1, "welcome": 1, "toc-title": 1,
    "intro-title": 1,
}

# Section headers matched by prefix (printed text is kept as written).
SECTIONS = [
    ("WHAT_HAPPENED", "WHAT HAPPENED"),
    ("WHEN_WHO", "WHEN &"),
    ("EVIDENCE", "THE EVIDENCE"),
    ("CRACK", "THE CRACK"),
    ("THEORIES", "THEORIES"),
    ("QUESTION", "THE QUESTION"),
    ("SEE_ALSO", "SEE ALSO"),
    ("CURIOUS", "CURIOUS?"),
]
SECTION_ORDER = [k for k, _ in SECTIONS]
NO_BODY_SECTIONS = {"SEE_ALSO", "CURIOUS"}      # body-text not allowed here
LIST_SECTIONS = {"EVIDENCE", "THEORIES"}

EXPECTED_EPILOGUE_GROUPS = [1, 3, 2, 1, 2]
N_CASES = 64

INVISIBLE = {ord(c): None for c in "\u00ad\u200b\u200c\u200d\u2060\ufeff"}
OPENERS = set(" \t\n([{\u201c\u2018\u2014\u2013/")

MARK_RE = re.compile(r"^(?:[\u2022\u00b7\u25aa\u25e6\u25cf\u25a0+*\-\u2013\u2014]|\d{1,2}[.)])\s+")
CASE_NAME_RE = re.compile(r"^CASE (\d+): (.+)$")
SEE_ALSO_RE = re.compile(r"^(.+?) \(Case (\d+)\) \u2014 (.+)$")


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def preview(text, n=60):
    t = " ".join(str(text).split())
    return t if len(t) <= n else t[:n] + "..."


def norm_name(s):
    """Loose comparison key for case names (See Also vs target)."""
    s = s.casefold().replace("\u02bb", "").replace("\u2019", "")
    s = re.sub(r"^the ", "", s.strip())
    return re.sub(r"[^\w]+", "", s)


def norm_name_exact(s):
    """Like norm_name but keeps hyphens/spaces, so '52 Hertz' != '52-Hertz'."""
    s = s.casefold().replace("\u02bb", "").replace("\u2019", "")
    s = re.sub(r"^the ", "", s.strip())
    return re.sub(r"\s+", " ", s)


class Ingest:
    def __init__(self, docx_path, allow_unfrozen=False):
        self.docx_path = docx_path
        self.allow_unfrozen = allow_unfrozen
        self.qa = []                       # (level, line)
        self.chars = defaultdict(set)      # style -> set(chars)
        self.sq_converted = 0
        self.cases = []
        self.cur = None                    # current case dict
        self.cs = None                     # current case parse state
        self.front = {s: [] for s in list(FRONT_STYLES) + ["intro-title"]}
        self.intro = []
        self.epi_title = ""
        self.epi_title_count = 0
        self.epi_groups = [[]]

    # ------------------------------------------------------------------ logging
    def log(self, level, code, where, msg):
        self.qa.append((level, f"{level}-{code} ({where}): {msg}"))

    # ------------------------------------------------------------ text helpers
    def smart(self, text, where):
        """Straight quotes -> curly, decided over the WHOLE string.
        ' -> right single quote/apostrophe (locked rule).
        "  -> opening after start/space/bracket/dash, otherwise closing."""
        out, n = [], 0
        for i, ch in enumerate(text):
            if ch == '"':
                prev = text[i - 1] if i else " "
                out.append("\u201c" if prev in OPENERS else "\u201d")
                n += 1
            elif ch == "'":
                out.append("\u2019")
                n += 1
            else:
                out.append(ch)
        if n:
            self.sq_converted += n
            self.log("INFO", "SMARTQUOTE", where,
                     f"converted {n} straight quote(s) in '{preview(text)}'")
        return "".join(out)

    def check_quotes(self, text, where):
        depth, bad = 0, False
        for ch in text:
            if ch == "\u201c":
                depth += 1
                if depth > 1:
                    bad = True
            elif ch == "\u201d":
                depth -= 1
                if depth < 0:
                    bad = True
        if bad:
            self.log("WARN", "QUOTE-ORDER", where,
                     f"curly double quotes nested or out of order in '{preview(text)}'")
        elif depth != 0:
            self.log("WARN", "QUOTE-BALANCE", where,
                     f"unbalanced curly double quotes in '{preview(text)}'")

    def clean_lines(self, raw, style, where, smart=True):
        """Soft-break aware line splitter. Returns normalised non-empty lines."""
        stripped = raw.translate(INVISIBLE)
        if stripped != raw:
            self.log("WARN", "INVISIBLE-CHAR", where,
                     f"removed soft-hyphen/zero-width characters in '{preview(raw)}'")
        s = (stripped.replace("\r\n", "\n").replace("\r", "\n")
             .replace("\x0b", "\n").replace("\x0c", "\n")
             .replace("\u00a0", " ").replace("\t", " "))
        lines = [re.sub(r" {2,}", " ", ln).strip() for ln in s.split("\n")]
        lines = [ln for ln in lines if ln]
        if smart and lines:
            lines = self.smart("\n".join(lines), where).split("\n")
            for ln in lines:
                self.chars[style].update(ln)
        return lines

    def one_line(self, raw, style, where):
        """Soft breaks removed (joined by a space)."""
        return " ".join(self.clean_lines(raw, style, where))

    def body_runs(self, para, style, where):
        """body-text: soft breaks kept as '\\n'; inline bold/italic kept.
        text == ''.join(run texts) is guaranteed."""
        runs = para.runs
        if "".join(r.text for r in runs) != para.text:
            self.log("CRITICAL", "TEXT-MISMATCH", where,
                     f"run text differs from paragraph text (hyperlink/field?) in '{preview(para.text)}'")
        chars = []
        for r in runs:
            b, i = bool(r.bold), bool(r.italic)
            for ch in r.text.translate(INVISIBLE):
                chars.append((ch, b, i))
        if "".join(c for c, _, _ in chars) != para.text:
            self.log("WARN", "INVISIBLE-CHAR", where,
                     f"removed soft-hyphen/zero-width characters in '{preview(para.text)}'")
        # normalise whitespace per line, on (char, bold, italic) tuples
        lines, line = [], []
        for ch, b, i in chars:
            if ch in "\r\x0b\x0c\n":
                lines.append(line)
                line = []
            elif ch in " \t\u00a0":
                line.append((" ", b, i))
            else:
                line.append((ch, b, i))
        lines.append(line)
        out = []
        for ln in lines:
            norm = []
            for ch, b, i in ln:
                if ch == " ":
                    if not norm or norm[-1][0] == " ":
                        continue
                    norm.append((" ", norm[-1][1], norm[-1][2]))
                else:
                    norm.append((ch, b, i))
            while norm and norm[-1][0] == " ":
                norm.pop()
            if not norm:
                continue
            if out:
                out.append(("\n", out[-1][1], out[-1][2]))
            out.extend(norm)
        text = "".join(c for c, _, _ in out)
        text = self.smart(text, where) if text else text   # 1:1 char replacement
        out = [(t, b, i) for t, (_, b, i) in zip(text, out)]
        self.check_quotes(text, where)
        self.chars[style].update(text.replace("\n", ""))
        has_emph = any((b or i) and c not in " \n" for c, b, i in out)
        if not has_emph:
            return text, None
        runs_out = []
        for ch, b, i in out:
            if runs_out and runs_out[-1][1] == b and runs_out[-1][2] == i:
                runs_out[-1][0].append(ch)
            else:
                runs_out.append(([ch], b, i))
        result = []
        for chs, b, i in runs_out:
            e = {"t": "".join(chs)}
            if b:
                e["b"] = True
            if i:
                e["i"] = True
            result.append(e)
        return text, result

    # ----------------------------------------------------------- document-level
    def check_hidden_content(self, doc):
        """Content python-docx would silently skip -> CRITICAL."""
        body = doc.element.body
        checks = (
            ("w:ins", "tracked insertion"), ("w:del", "tracked deletion"),
            ("w:moveFrom", "tracked move"), ("w:moveTo", "tracked move"),
            ("w:sdt", "content control"), ("w:txbxContent", "text box"),
        )
        for tag, label in checks:
            n = len(body.xpath(f".//{tag}"))
            if n:
                self.log("CRITICAL", "HIDDEN-CONTENT", "document",
                         f"{n} {label} element(s) ({tag}); text inside would be skipped. "
                         "Accept/reject changes or remove it in Word.")
        if doc.tables:
            self.log("CRITICAL", "HIDDEN-CONTENT", "document",
                     f"{len(doc.tables)} table(s) found; table text is not read.")

    # --------------------------------------------------------------------- main
    def run(self):
        print(f"=== UNEXPLAINED: INGEST & QA VALIDATOR (v{VERSION}) ===")
        if not os.path.exists(self.docx_path):
            print(f"[CRITICAL-FILE] Manuscript not found at: {self.docx_path}")
            return 1
        os.makedirs(JSON_DIR, exist_ok=True)

        self.sha = sha256_of(self.docx_path)
        if self.sha != FROZEN_SHA256:
            if self.allow_unfrozen:
                self.log("WARN", "HASH-UNFROZEN", "manuscript",
                         f"SHA-256 {self.sha[:16]}... differs from the frozen hash")
            else:
                self.log("CRITICAL", "HASH-FROZEN", "manuscript",
                         f"SHA-256 is {self.sha}, expected {FROZEN_SHA256}. "
                         "Restore the frozen file, or re-freeze deliberately "
                         "(update FROZEN_SHA256) / use --allow-unfrozen for drafts.")

        try:
            doc = Document(self.docx_path)
        except Exception as exc:                      # noqa: BLE001
            print(f"[CRITICAL-FILE] Cannot open manuscript: {exc}")
            return 1
        self.check_hidden_content(doc)

        region = "front"
        for idx, para in enumerate(doc.paragraphs, start=1):
            raw = para.text
            sname = para.style.name if para.style is not None else ""
            where = f"para #{idx}"

            if not raw.strip():                       # Global Strip Rule
                if para._p.xpath(".//w:drawing | .//w:pict | .//w:object"):
                    self.log("WARN", "DRAWING", where,
                             "empty paragraph holds an embedded graphic (ignored)")
                if region == "intro" and self.intro and self.intro[-1]["type"] != "stanza_break":
                    self.intro.append({"type": "stanza_break"})
                elif region == "epilogue" and self.epi_groups[-1]:
                    self.epi_groups.append([])
                continue

            if sname not in ALLOWED_STYLES:
                self.log("CRITICAL", "STYLE", where,
                         f"style '{sname}' is not allowed on text: '{preview(raw)}'")
                continue

            # ---- region transitions
            if sname == "intro-title":
                if region != "front":
                    self.log("CRITICAL", "STRUCTURE", where, "intro-title out of place or repeated")
                    continue
                region = "intro"
                self.front["intro-title"].extend(self.clean_lines(raw, sname, where))
                continue
            if sname == "case-name":
                if region == "front":
                    self.log("CRITICAL", "STRUCTURE", where, "case-name before intro-title")
                    region = "cases"
                elif region == "epilogue":
                    self.log("CRITICAL", "STRUCTURE", where, "case-name after the Epilogue")
                    continue
                else:
                    region = "cases"
                self.start_case(raw, where)
                continue
            if sname == "epilogue-title":
                self.epi_title_count += 1
                if region == "epilogue":
                    self.log("CRITICAL", "STRUCTURE", where, "epilogue-title repeated")
                    continue
                self.finish_case()
                region = "epilogue"
                self.epi_title = " ".join(self.clean_lines(raw, sname, where))
                continue

            # ---- per-region content
            if region == "front":
                if sname in FRONT_STYLES:
                    self.front[sname].extend(self.clean_lines(raw, sname, where))
                else:
                    self.log("CRITICAL", "STRUCTURE", where,
                             f"style '{sname}' is not allowed in the front matter: '{preview(raw)}'")
            elif region == "intro":
                if sname == "body-text":
                    text, runs = self.body_runs(para, sname, "Introduction")
                    block = {"type": "body-text", "text": text}
                    if runs:
                        block["runs"] = runs
                    self.intro.append(block)
                else:
                    self.log("CRITICAL", "STRUCTURE", where,
                             f"style '{sname}' is not allowed in the Introduction: '{preview(raw)}'")
            elif region == "epilogue":
                if sname == "epilogue-body":
                    self.epi_groups[-1].extend(self.clean_lines(raw, sname, "Epilogue"))
                else:
                    self.log("CRITICAL", "STRUCTURE", where,
                             f"style '{sname}' is not allowed in the Epilogue: '{preview(raw)}'")
            else:
                self.case_para(para, sname, where)

        self.finish_case()
        self.validate_all()
        return self.write_outputs()

    # ------------------------------------------------------------------- cases
    def start_case(self, raw, where):
        self.finish_case()
        text = " ".join(self.clean_lines(raw, "case-name", where))
        m = CASE_NAME_RE.match(text)
        num = int(m.group(1)) if m else len(self.cases) + 1
        label = f"Case {num}"
        if not m:
            self.log("CRITICAL", "STRUCTURE", label,
                     f"case-name '{text}' does not match '^CASE (\\d+): (.+)$'")
        self.cur = {
            "case_number": num,
            "case_name": text,
            "case_topic": m.group(2).strip() if m else text,
            "big_word": "",
            "title_lines": [],
            "quote": "",
            "captions": [],
            "sections": [],
        }
        self.cs = {"label": label, "phase": "title", "big_seen": False,
                   "quote_lines": [], "last_style": "case-name",
                   "section": None}

    def finish_case(self):
        if self.cur is None:
            return
        cs, c = self.cs, self.cur
        if cs["quote_lines"]:
            joined = self.smart("\n".join(cs["quote_lines"]), cs["label"])
            self.check_quotes(joined, cs["label"])
            self.chars["case-quote"].update(joined.replace("\n", ""))
            c["quote"] = joined
        self.cases.append(c)
        self.cur = None
        self.cs = None

    def add_item(self, item):
        self.cur["sections"][-1]["items"].append(item)

    def case_para(self, para, sname, where):
        c, cs = self.cur, self.cs
        label = cs["label"]
        raw = para.text
        in_sections = bool(c["sections"])
        cur_key = cs["section"]

        if sname in ("big-word", "case-title"):
            if cs["phase"] != "title" or in_sections:
                self.log("CRITICAL", "STRUCTURE", label,
                         f"{sname} outside the title block: '{preview(raw)}'")
            else:
                lines = self.clean_lines(raw, sname, label)
                if sname == "big-word":
                    if cs["big_seen"]:
                        self.log("CRITICAL", "STRUCTURE", label, "more than one big-word paragraph")
                    else:
                        cs["big_seen"] = True
                        if len(lines) > 1:
                            self.log("WARN", "BIGWORD-MULTILINE", label,
                                     f"big-word has {len(lines)} lines; only the first is the big line")
                        c["big_word"] = lines[0] if lines else ""
                        c["title_lines"].extend(lines[1:])
                else:
                    if not cs["big_seen"]:
                        self.log("CRITICAL", "STRUCTURE", label,
                                 f"case-title before big-word: '{preview(raw)}'")
                    c["title_lines"].extend(lines)
            cs["last_style"] = sname
            return

        cs["phase"] = "body"          # any other style ends the title block

        if sname == "case-quote":
            if in_sections:
                self.log("CRITICAL", "STRUCTURE", label, "case-quote after a section header")
            elif cs["quote_lines"] and cs["last_style"] != "case-quote":
                self.log("CRITICAL", "STRUCTURE", label,
                         "case-quote is split into several blocks (must be one block per case)")
                cs["quote_lines"].extend(self.clean_lines(raw, sname, label, smart=False))
            else:
                cs["quote_lines"].extend(self.clean_lines(raw, sname, label, smart=False))

        elif sname == "image-caption":
            if in_sections:
                self.log("CRITICAL", "STRUCTURE", label, "image-caption after a section header")
            c["captions"].extend(self.clean_lines(raw, sname, label))

        elif sname == "section-header":
            text = self.one_line(raw, sname, label)
            key = "UNKNOWN"
            for k, prefix in SECTIONS:
                if text.upper().startswith(prefix):
                    key = k
                    break
            if key == "UNKNOWN":
                self.log("CRITICAL", "STRUCTURE", label, f"unknown section header '{text}'")
            c["sections"].append({"key": key, "header": text, "items": []})
            cs["section"] = key

        elif not in_sections:
            self.log("CRITICAL", "STRUCTURE", label,
                     f"{sname} before the first section header: '{preview(raw)}'")

        elif sname == "body-text":
            if cur_key in NO_BODY_SECTIONS:
                self.log("CRITICAL", "STRUCTURE", label,
                         f"body-text inside section {cur_key}: '{preview(raw)}'")
            text, runs = self.body_runs(para, sname, label)
            item = {"type": "body-text", "text": text}
            if runs:
                item["runs"] = runs
            self.add_item(item)

        elif sname == "list-item":
            if cur_key not in LIST_SECTIONS:
                self.log("CRITICAL", "STRUCTURE", label,
                         f"list-item inside section {cur_key}: '{preview(raw)}'")
            text = self.one_line(raw, sname, label)
            if MARK_RE.match(text):
                self.log("WARN", "LISTMARKER-STRIPPED", label,
                         f"typed bullet/number removed from '{preview(text)}' "
                         "(bullets are drawn by the engine)")
                text = MARK_RE.sub("", text)
            self.check_quotes(text, label)
            if cur_key == "THEORIES":
                if " \u2014 " in text:
                    lead, rest = text.split(" \u2014 ", 1)
                    lead, rest = lead.strip(), rest.strip()
                    if not lead or not rest:
                        self.log("WARN", "LISTITEM-NOLEAD", label,
                                 f"empty lead or body in '{preview(text)}'")
                    self.add_item({"type": "theory-item", "lead": lead, "body": rest})
                else:
                    self.log("WARN", "LISTITEM-NOLEAD", label,
                             f"Theories list-item without ' \u2014 ': '{preview(text)}'")
                    self.add_item({"type": "theory-item", "lead": "", "body": text})
            else:
                self.add_item({"type": "evidence-item", "text": text})

        elif sname == "see-also":
            if cur_key != "SEE_ALSO":
                self.log("CRITICAL", "STRUCTURE", label,
                         f"see-also outside the SEE ALSO section: '{preview(raw)}'")
            for ln in self.clean_lines(raw, sname, label):
                m = SEE_ALSO_RE.match(ln)
                item = {"type": "see-also", "text": ln}
                if m:
                    item.update({"name": m.group(1).strip(),
                                 "target": int(m.group(2)),
                                 "description": m.group(3).strip()})
                else:
                    self.log("CRITICAL", "STRUCTURE", label,
                             f"see-also format mismatch: '{ln}'")
                self.add_item(item)

        elif sname == "search-term":
            if cur_key != "CURIOUS":
                self.log("CRITICAL", "STRUCTURE", label,
                         f"search-term outside the CURIOUS? section: '{preview(raw)}'")
            for ln in self.clean_lines(raw, sname, label):
                self.check_quotes(ln, label)
                self.add_item({"type": "search-term", "text": ln})

        else:
            self.log("CRITICAL", "STRUCTURE", label,
                     f"style '{sname}' is not allowed inside a case: '{preview(raw)}'")

        cs["last_style"] = sname

    # -------------------------------------------------------------- validation
    def validate_all(self):
        for k, exp in EXPECTED_FRONT_COUNTS.items():
            got = len(self.front[k])
            if got != exp:
                self.log("CRITICAL", "STRUCTURE", f"Front Matter '{k}'",
                         f"expected {exp}, got {got}")
        if not any(b["type"] == "body-text" for b in self.intro):
            self.log("CRITICAL", "STRUCTURE", "Introduction", "no body-text found")
        while self.intro and self.intro[-1]["type"] == "stanza_break":
            self.intro.pop()

        self.epi_groups = [g for g in self.epi_groups if g]
        if self.epi_title_count != 1:
            self.log("CRITICAL", "STRUCTURE", "Epilogue",
                     f"expected exactly 1 epilogue-title, found {self.epi_title_count}")
        if not self.epi_groups:
            self.log("CRITICAL", "STRUCTURE", "Epilogue", "no epilogue-body text found")
        counts = [len(g) for g in self.epi_groups]
        if self.epi_groups and counts != EXPECTED_EPILOGUE_GROUPS:
            self.log("WARN", "EPILOGUE-GROUPS", "Epilogue",
                     f"expected {EXPECTED_EPILOGUE_GROUPS}, got {counts} "
                     "(a copy change, not an error)")

        if len(self.cases) != N_CASES:
            self.log("CRITICAL", "STRUCTURE", "Book",
                     f"expected {N_CASES} cases, found {len(self.cases)}")
        topics = {c["case_number"]: c["case_topic"] for c in self.cases}
        images_ok = os.path.isdir(IMAGES_DIR)
        if not images_ok:
            self.log("CRITICAL", "IMAGES", "Book", f"images folder not found: {IMAGES_DIR}")

        seen_names = set()
        for pos, c in enumerate(self.cases, start=1):
            n = c["case_number"]
            label = f"Case {n}"
            if n != pos:
                self.log("CRITICAL", "STRUCTURE", label, f"out of sequence at position {pos}")
            if c["case_name"] in seen_names:
                self.log("CRITICAL", "STRUCTURE", label, "duplicate case-name")
            seen_names.add(c["case_name"])

            total = (1 if c["big_word"] else 0) + len(c["title_lines"])
            if not c["big_word"] or total not in (2, 3):
                self.log("CRITICAL", "STRUCTURE", label,
                         f"title lines={total} (expected 2 or 3): "
                         f"{[c['big_word']] + c['title_lines']}")
            elif not re.match(r"^(A|An|The)\s", c["big_word"]):
                self.log("INFO", "BIGWORD-NOARTICLE", label,
                         f"big line '{c['big_word']}' has no leading A/An/The")

            if not c["quote"]:
                self.log("CRITICAL", "STRUCTURE", label, "missing case-quote")
            if not 1 <= len(c["captions"]) <= 3:
                self.log("CRITICAL", "STRUCTURE", label,
                         f"image-caption lines={len(c['captions'])} (expected 1..3)")

            keys = [s["key"] for s in c["sections"]]
            if keys != SECTION_ORDER:
                self.log("CRITICAL", "STRUCTURE", label,
                         f"section headers {keys} != expected {SECTION_ORDER}")

            sa = [i for s in c["sections"] for i in s["items"] if i["type"] == "see-also"]
            st = [i for s in c["sections"] for i in s["items"] if i["type"] == "search-term"]
            if len(sa) != 3:
                self.log("CRITICAL", "STRUCTURE", label, f"expected 3 see-also entries, found {len(sa)}")
            targets, descs = [], []
            for e in sa:
                if "target" not in e:
                    continue
                t = e["target"]
                if t == n or t not in topics:
                    self.log("CRITICAL", "STRUCTURE", label,
                             f"invalid see-also target Case {t} in '{e['text']}'")
                    continue
                targets.append(t)
                descs.append(re.sub(r"[\s.]+$", "", e["description"]).casefold())
                if norm_name(e["name"]) != norm_name(topics[t]):
                    self.log("WARN", "SEEALSO-NAME", label,
                             f"'{e['name']}' != Case {t} name '{topics[t]}'")
                elif norm_name_exact(e["name"]) != norm_name_exact(topics[t]):
                    self.log("INFO", "SEEALSO-NAME-STYLE", label,
                             f"'{e['name']}' vs Case {t} name '{topics[t]}' differ only in hyphen/space")
            if len(set(targets)) != len(targets):
                self.log("CRITICAL", "STRUCTURE", label, f"duplicate see-also target in {targets}")
            if len(set(descs)) != len(descs):
                self.log("WARN", "SEEALSO-DUPLICATE-DESC", label,
                         "two see-also entries share the same description")

            if len(st) != 2:
                self.log("CRITICAL", "STRUCTURE", label, f"expected 2 search-term lines, found {len(st)}")
            for e in st:
                if not (e["text"].startswith("\u201c") and e["text"].endswith("\u201d")):
                    self.log("WARN", "SEARCHTERM-QUOTES", label,
                             f"search term not wrapped in curly quotes: '{e['text']}'")

            for s in c["sections"]:
                if not s["items"]:
                    self.log("WARN", "EMPTY-SECTION", label, f"section '{s['header']}' has no content")

            if images_ok and not os.path.exists(os.path.join(IMAGES_DIR, f"CASE{n:02d}.png")):
                self.log("CRITICAL", "IMAGES", label, f"missing image CASE{n:02d}.png")

    # ----------------------------------------------------------------- outputs
    def write_outputs(self):
        criticals = [m for lv, m in self.qa if lv == "CRITICAL"]
        warns = [m for lv, m in self.qa if lv == "WARN"]
        infos = [m for lv, m in self.qa if lv == "INFO"]
        status = "OK" if not criticals else "FAILED"
        head = {
            "schema_version": SCHEMA_VERSION,
            "ingest_version": VERSION,
            "source_file": os.path.basename(self.docx_path),
            "source_sha256": self.sha,
            "ingest_status": status,
            "criticals": len(criticals),
        }

        def dump(name, payload):
            with open(os.path.join(JSON_DIR, name), "w", encoding="utf-8") as f:
                json.dump(dict(head, **payload), f, indent=2, ensure_ascii=False)

        dump("frontmatter.json", {"frontmatter": self.front, "introduction": self.intro})
        dump("cases.json", {"cases": self.cases})
        dump("epilogue.json", {"title": self.epi_title, "groups": self.epi_groups})

        union = set()
        for s in self.chars.values():
            union |= s
        union.discard("\n")
        with open(os.path.join(JSON_DIR, "charset.txt"), "w", encoding="utf-8") as f:
            f.write("".join(sorted(union)))
        with open(os.path.join(JSON_DIR, "charset_by_style.json"), "w", encoding="utf-8") as f:
            json.dump({k: "".join(sorted(v - {"\n"})) for k, v in sorted(self.chars.items())},
                      f, indent=2, ensure_ascii=False)

        with open(QA_REPORT_PATH, "w", encoding="utf-8") as f:
            f.write(f"=== UNEXPLAINED: INGEST QA REPORT (v{VERSION}) ===\n")
            f.write(f"Manuscript SHA-256: {self.sha}\n")
            f.write(f"Frozen hash match: {'YES' if self.sha == FROZEN_SHA256 else 'NO'}\n")
            f.write(f"Cases parsed: {len(self.cases)}\n")
            f.write(f"Unique characters: {len(union)}\n")
            f.write(f"Straight quotes converted: {self.sq_converted}\n\n")
            for _, m in self.qa:
                f.write(m + "\n")
            if not self.qa:
                f.write("ALL INGEST CHECKS PASSED: 0 CRITICAL, 0 WARNINGS.\n")
            f.write(f"\nINGEST SUMMARY: {len(self.cases)} cases, "
                    f"{len(criticals)} Critical, {len(warns)} Warnings, {len(infos)} Info.\n")

        print(f"Parsed {len(self.cases)} cases.")
        print(f"Ingest QA Summary: {len(criticals)} Criticals, {len(warns)} Warnings, {len(infos)} Info.")
        for lv, m in self.qa:
            if lv != "INFO":
                print("  " + m)
        if criticals:
            print("FAILED: JSON written for diagnosis only; build_book.py must not run.")
            return 1
        print("SUCCESS: frontmatter.json, cases.json, epilogue.json, charset.txt, "
              "charset_by_style.json and qa_report.txt generated.")
        return 0


def main():
    ap = argparse.ArgumentParser(description="UNEXPLAINED manuscript ingest")
    ap.add_argument("--docx", default=DEFAULT_DOCX, help="path to the styled manuscript")
    ap.add_argument("--allow-unfrozen", action="store_true",
                    help="do not fail when the SHA-256 differs from the frozen hash")
    args = ap.parse_args()
    sys.exit(Ingest(args.docx, args.allow_unfrozen).run())


if __name__ == "__main__":
    main()
