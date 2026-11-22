#!/usr/bin/env python3
"""docsync - helper CLI for the Doc Sync Agent.

The AI tool (Copilot, Antigravity, Claude Code, Cursor, ...) does the thinking.
This script does the deterministic work:

  init       save the documentation root folder (asked once)
  scan       index all docs, learn the folder/naming/heading conventions,
             and collect what changed in the code since the last run
  note       record a change manually (when work is not in git)
  set-mode   choose "new-version" (file_v2.docx) or "in-place" (with backup)
  apply      apply an edits.json: update section / add section /
             append to section / create new document
  status     show current configuration

Run `python docsync.py <command> -h` for options.
"""
import argparse
import copy
import difflib
import json
import math
import os
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:  # pragma: no cover
    sys.exit("Missing dependency. Run: pip install -r requirements.txt")

HERE = Path(__file__).resolve().parent
STATE = Path(os.environ.get("DOCSYNC_HOME") or (HERE / ".docsync"))
CONFIG = STATE / "config.json"
NOTES = STATE / "manual_changes.md"
WORK_DIR_NAME = "_docsync"  # lives inside the docs root: backups + changelog
SUPPORTED = {".docx", ".md", ".txt"}
EDITABLE = {".docx", ".md"}
SKIP_DIRS = {WORK_DIR_NAME, ".git", "node_modules", "__pycache__", ".docsync"}
LOCKFILE_EXCLUDES = [
    ":(exclude)*.lock", ":(exclude)package-lock.json", ":(exclude)*.min.js",
    ":(exclude)*.map", ":(exclude)*.svg",
]
VERSION_RE = re.compile(r"^(?P<base>.+?)(?P<sep>[_\- ])v(?P<n>\d+)$", re.I)


class EditError(Exception):
    pass


def die(msg, code=1):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(code)


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# --------------------------------------------------------------------------
# config
# --------------------------------------------------------------------------
def load_config(required=True):
    if CONFIG.exists():
        return json.loads(CONFIG.read_text("utf-8"))
    if required:
        die("Not configured yet. Run: python docsync.py init")
    return {}


def save_config(cfg):
    STATE.mkdir(parents=True, exist_ok=True)
    CONFIG.write_text(json.dumps(cfg, indent=2), "utf-8")


def docs_root(cfg):
    root = Path(cfg.get("docs_root", ""))
    if not root.is_dir():
        die(f"Docs root not found: {root}. Re-run: python docsync.py init --reset")
    return root


# --------------------------------------------------------------------------
# text helpers
# --------------------------------------------------------------------------
def norm(s):
    """Normalise a heading for matching: drop numbering, case, spacing."""
    s = re.sub(r"^\s*\d+(?:\.\d+)*[.)]?\s+", "", s.strip())
    return re.sub(r"\s+", " ", s).casefold()


def clean_inline(text):
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1 (\2)", text)
    return text.replace("`", "")


def split_bold(text):
    text = clean_inline(text)
    out = []
    for seg in re.split(r"(\*\*[^*]+\*\*)", text):
        if not seg:
            continue
        if seg.startswith("**") and seg.endswith("**") and len(seg) > 4:
            out.append((seg[2:-2], True))
        else:
            out.append((seg, False))
    return out or [("", False)]


def parse_body(text):
    """Tiny markdown subset -> blocks: p, b (bullet), n (numbered), t (table), code."""
    lines = text.replace("\r\n", "\n").strip("\n").split("\n")
    out, para, i = [], [], 0

    def flush():
        if para:
            out.append(("p", " ".join(x.strip() for x in para)))
            para.clear()

    while i < len(lines):
        s = lines[i].strip()
        if not s:
            flush()
            i += 1
            continue
        if s.startswith("```"):
            flush()
            i += 1
            code = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i])
                i += 1
            i += 1
            out.append(("code", code))
            continue
        if s.startswith("|"):
            flush()
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                is_sep = cells and all(re.fullmatch(r":?-{2,}:?", c) for c in cells)
                if not is_sep:
                    rows.append(cells)
                i += 1
            if rows:
                out.append(("t", rows))
            continue
        m = re.match(r"^[-*\u2022]\s+(.*)", s)
        if m:
            flush()
            out.append(("b", m.group(1)))
            i += 1
            continue
        m = re.match(r"^(\d+)[.)]\s+(.*)", s)
        if m:
            flush()
            out.append(("n", m.group(1), m.group(2)))
            i += 1
            continue
        para.append(s)
        i += 1
    flush()
    return out


# --------------------------------------------------------------------------
# docx helpers
# --------------------------------------------------------------------------
HEADING_RE = re.compile(r"^heading\s*(\d)$", re.I)


def style_name(p):
    try:
        return p.style.name or ""
    except Exception:
        return ""


def para_level(p):
    m = HEADING_RE.match(style_name(p))
    return int(m.group(1)) if m else None


def has_media(el):
    return any(True for _ in el.iter(qn("w:drawing"), qn("w:pict"), qn("w:object")))


def is_bullet_numbering(doc, num_id, ilvl):
    """True if numbering definition num_id/ilvl is a bullet list."""
    try:
        root = doc.part.numbering_part.element
        abstract = None
        for num in root.findall(qn("w:num")):
            if num.get(qn("w:numId")) == str(num_id):
                abstract = num.find(qn("w:abstractNumId")).get(qn("w:val"))
                break
        if abstract is None:
            return False
        for an in root.findall(qn("w:abstractNum")):
            if an.get(qn("w:abstractNumId")) == abstract:
                for lvl in an.findall(qn("w:lvl")):
                    if lvl.get(qn("w:ilvl")) == str(ilvl):
                        fmt = lvl.find(qn("w:numFmt"))
                        return fmt is not None and fmt.get(qn("w:val")) == "bullet"
    except Exception:
        pass
    return False


class Styler:
    """Builds new paragraphs/headings/tables by cloning the document's own ones,
    so new content inherits the existing fonts, spacing and numbering."""

    NORMAL_NAMES = {"normal", "body text", "body text 2", "default", "text body"}

    def __init__(self, dd):
        self.dd = dd
        self.doc = dd.doc
        self.h = {}
        self.title = None
        self.bullet = None
        self.table_style = None
        candidates = []
        for el in dd.blocks():
            if el.tag == qn("w:tbl"):
                if self.table_style is None:
                    try:
                        self.table_style = Table(el, self.doc._body).style.name
                    except Exception:
                        pass
                continue
            p = dd.par(el)
            name = style_name(p)
            lvl = para_level(p)
            ppr = el.find(qn("w:pPr"))
            numpr = ppr.find(qn("w:numPr")) if ppr is not None else None
            if lvl:
                if lvl not in self.h:
                    self.h[lvl] = copy.deepcopy(el)
            elif name == "Title":
                if self.title is None:
                    self.title = copy.deepcopy(el)
            elif name.startswith("List") or numpr is not None:
                is_bullet = "bullet" in name.lower()
                if numpr is not None and not is_bullet:
                    nid = numpr.find(qn("w:numId"))
                    il = numpr.find(qn("w:ilvl"))
                    is_bullet = is_bullet_numbering(
                        self.doc,
                        nid.get(qn("w:val")) if nid is not None else None,
                        il.get(qn("w:val")) if il is not None else 0)
                if is_bullet and self.bullet is None:
                    self.bullet = copy.deepcopy(el)
            elif p.text.strip():
                candidates.append((name, el))
        self.body = self._pick_body(candidates)

    def _pick_body(self, candidates):
        def first_bold(el):
            r = next(el.iter(qn("w:r")), None)
            rpr = r.find(qn("w:rPr")) if r is not None else None
            return rpr is not None and rpr.find(qn("w:b")) is not None

        def txt(el):
            return "".join(t.text or "" for t in el.iter(qn("w:t")))

        for want_normal in (True, False):
            for name, el in candidates:
                if want_normal and name.lower() not in self.NORMAL_NAMES:
                    continue
                if not first_bold(el) and len(txt(el)) >= 20:
                    return copy.deepcopy(el)
        return copy.deepcopy(candidates[0][1]) if candidates else None

    # -- low level
    def _fresh(self, style=None):
        try:
            p = self.doc.add_paragraph(style=style) if style else self.doc.add_paragraph()
        except KeyError:
            p = self.doc.add_paragraph()
        el = p._p
        el.getparent().remove(el)
        return el

    def _fill(self, el, text, rpr=None, md=True, force_bold=False):
        p = Paragraph(el, self.doc._body)
        segments = split_bold(text) if md else [(clean_inline(text), False)]
        for seg, bold in segments:
            r = p.add_run(seg)
            if rpr is not None:
                r._r.insert(0, copy.deepcopy(rpr))
            if bold or force_bold:
                r.bold = True

    def _clone(self, tpl, text, md=True, strip_bold=True):
        el = copy.deepcopy(tpl)
        rpr = None
        r0 = next(el.iter(qn("w:r")), None)
        if r0 is not None and r0.find(qn("w:rPr")) is not None:
            rpr = copy.deepcopy(r0.find(qn("w:rPr")))
            if strip_bold:
                for tag in ("w:b", "w:bCs"):
                    for b in rpr.findall(qn(tag)):
                        rpr.remove(b)
        for ch in list(el):
            if ch.tag != qn("w:pPr"):
                el.remove(ch)
        ppr = el.find(qn("w:pPr"))
        if ppr is not None:
            for sect in ppr.findall(qn("w:sectPr")):  # never copy a section break
                ppr.remove(sect)
        self._fill(el, text, rpr, md=md)
        return el

    # -- public builders
    def para(self, text, tpl=None):
        t = tpl if tpl is not None else self.body
        if t is not None:
            return self._clone(t, text)
        el = self._fresh()
        self._fill(el, text)
        return el

    def bullet_item(self, text, tpl=None):
        if self.bullet is not None:
            return self._clone(self.bullet, text)
        try:
            self.doc.styles["List Bullet"]
            el = self._fresh("List Bullet")
            self._fill(el, text)
            return el
        except KeyError:
            return self.para("\u2022 " + text, tpl)

    def heading(self, level, text):
        tpl = self.h.get(level)
        if tpl is not None:
            return self._clone(tpl, text, md=False, strip_bold=False)
        try:
            self.doc.styles[f"Heading {level}"]
            el = self._fresh(f"Heading {level}")
            self._fill(el, text, md=False)
            return el
        except KeyError:
            el = self._fresh()
            self._fill(el, text, md=False, force_bold=True)
            return el

    def title_par(self, text):
        if self.title is not None:
            return self._clone(self.title, text, md=False, strip_bold=False)
        try:
            self.doc.styles["Title"]
            el = self._fresh("Title")
            self._fill(el, text, md=False)
            return el
        except KeyError:
            return self.heading(1, text)

    def code_lines(self, lines, tpl=None):
        out = []
        for ln in lines:
            el = self.para(ln or " ", tpl)
            for r in Paragraph(el, self.doc._body).runs:
                r.font.name = "Consolas"
            out.append(el)
        return out

    def table(self, rows):
        n, m = len(rows), max(len(r) for r in rows)
        t = self.doc.add_table(rows=n, cols=m)
        for style in (self.table_style, "Table Grid"):
            if not style:
                continue
            try:
                t.style = style
                break
            except Exception:
                continue
        for i, row in enumerate(rows):
            for j in range(m):
                cell = t.cell(i, j)
                cell.text = clean_inline(row[j]) if j < len(row) else ""
                if i == 0:
                    for par in cell.paragraphs:
                        for r in par.runs:
                            r.bold = True
        el = t._tbl
        el.getparent().remove(el)
        return el


class DocxDoc:
    kind = "docx"

    def __init__(self, path=None):
        self.path = path
        self.doc = Document(str(path)) if path else Document()
        self.body = self.doc.element.body
        self.styler = Styler(self)

    def par(self, el):
        return Paragraph(el, self.doc._body)

    def blocks(self):
        return [c for c in self.body.iterchildren() if c.tag in (qn("w:p"), qn("w:tbl"))]

    def level_of(self, el):
        return para_level(self.par(el)) if el.tag == qn("w:p") else None

    def block_text(self, el):
        if el.tag == qn("w:tbl"):
            t = Table(el, self.doc._body)
            rows = []
            for r in t.rows:
                cells = []
                prev = None
                for c in r.cells:
                    if c._tc is prev:
                        continue
                    prev = c._tc
                    cells.append(c.text.replace("\n", " ").strip())
                rows.append("| " + " | ".join(cells) + " |")
            return "\n".join(rows)
        p = self.par(el)
        text = p.text
        name = style_name(p)
        if name == "Title":
            return f"[Title] {text}"
        ppr = el.find(qn("w:pPr"))
        if name.startswith("List") or (ppr is not None and ppr.find(qn("w:numPr")) is not None):
            text = "- " + text
        if has_media(el):
            text += " [image/object]"
        return text

    # -- reading
    def sections(self):
        out = [{"heading": None, "level": 0, "lines": []}]
        for el in self.blocks():
            lvl = self.level_of(el)
            if lvl:
                out.append({"heading": self.par(el).text, "level": lvl, "lines": []})
            else:
                t = self.block_text(el)
                if t.strip():
                    out[-1]["lines"].append(t)
        return out

    def stats(self):
        bl = self.blocks()
        return {
            "tables": sum(1 for e in bl if e.tag == qn("w:tbl")),
            "bullets": sum(1 for e in bl if self.block_text(e).startswith("- ")),
            "title_style": any(style_name(self.par(e)) == "Title" for e in bl if e.tag == qn("w:p")),
            "front_matter": False,
        }

    # -- locating
    def find(self, heading, occurrence=None):
        bl = self.blocks()
        lv = [self.level_of(e) for e in bl]
        key = norm(heading)
        hits = [i for i, l in enumerate(lv) if l and norm(self.par(bl[i]).text) == key]
        if not hits:
            avail = [self.par(bl[i]).text for i, l in enumerate(lv) if l]
            raise EditError(f'heading not found: "{heading}". Headings in file: {avail}')
        if occurrence:
            if occurrence > len(hits):
                raise EditError(f'heading "{heading}" has only {len(hits)} occurrence(s)')
            return bl, lv, hits[occurrence - 1]
        if len(hits) > 1:
            raise EditError(f'heading "{heading}" appears {len(hits)} times; add "occurrence": 1..{len(hits)}')
        return bl, lv, hits[0]

    @staticmethod
    def _range(lv, h, subtree):
        end = len(lv)
        for j in range(h + 1, len(lv)):
            if lv[j] and (not subtree or lv[j] <= lv[h]):
                end = j
                break
        return h + 1, end

    def _insert_after(self, anchor, els):
        if anchor is None:
            for k, e in enumerate(els):
                self.body.insert(k, e)
            return
        for e in els:
            anchor.addnext(e)
            anchor = e

    def build(self, text, tpl=None):
        s, out = self.styler, []
        for b in parse_body(text):
            if b[0] == "p":
                out.append(s.para(b[1], tpl))
            elif b[0] == "b":
                out.append(s.bullet_item(b[1], tpl))
            elif b[0] == "n":
                out.append(s.para(f"{b[1]}. {b[2]}", tpl))
            elif b[0] == "code":
                out.extend(s.code_lines(b[1], tpl))
            elif b[0] == "t":
                out.append(s.table(b[1]))
        return out

    def _first_body_par(self, els):
        for e in els:
            if e.tag != qn("w:p") or not self.par(e).text.strip():
                continue
            ppr = e.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:numPr")) is not None:
                continue
            if self.level_of(e):
                continue
            return e
        return None

    # -- operations
    def section_text(self, heading, occurrence=None):
        bl, lv, h = self.find(heading, occurrence)
        s, e = self._range(lv, h, False)
        return "\n".join(self.block_text(x) for x in bl[s:e])

    def replace_section(self, heading, text, occurrence=None, allow_media_loss=False):
        bl, lv, h = self.find(heading, occurrence)
        s, e = self._range(lv, h, False)
        old = bl[s:e]
        if not allow_media_loss and any(has_media(x) for x in old):
            raise EditError(
                f'section "{heading}" contains an image/object that would be deleted. '
                'Use append_to_section instead, or set "allow_media_loss": true if intended.')
        old_text = "\n".join(self.block_text(x) for x in old)
        tpl = self._first_body_par(old)
        for x in old:
            self.body.remove(x)
        self._insert_after(bl[h], self.build(text, tpl))
        return old_text

    def append_to_section(self, heading, text, occurrence=None):
        bl, lv, h = self.find(heading, occurrence)
        s, e = self._range(lv, h, False)
        anchor = bl[e - 1] if e > s else bl[h]
        tpl = self._first_body_par(bl[s:e])
        self._insert_after(anchor, self.build(text, tpl))

    def add_section(self, heading, level, text, after_heading=None, occurrence=None):
        bl = self.blocks()
        if after_heading:
            bl, lv, h = self.find(after_heading, occurrence)
            _, e = self._range(lv, h, True)
            anchor = bl[e - 1]
        else:
            anchor = bl[-1] if bl else None
        els = [self.styler.heading(level, heading)] + self.build(text)
        self._insert_after(anchor, els)

    def save(self, dest):
        bl = self.blocks()
        if bl and bl[-1].tag == qn("w:tbl"):  # Word wants a paragraph after a final table
            bl[-1].addnext(self.styler._fresh())
        dest = Path(dest)
        tmp = dest.with_name(dest.name + ".docsync.tmp")
        self.doc.save(str(tmp))
        os.replace(tmp, dest)


class MdDoc:
    kind = "md"

    def __init__(self, path=None):
        raw = Path(path).read_bytes().decode("utf-8", "replace") if path else ""
        self.nl = "\r\n" if "\r\n" in raw else "\n"
        self.final_nl = raw.endswith("\n") or not raw
        self.lines = raw.splitlines()

    def levels(self):
        lv, fence = [], False
        for ln in self.lines:
            if re.match(r"^\s*(```|~~~)", ln):
                fence = not fence
                lv.append(None)
                continue
            m = None if fence else re.match(r"^(#{1,6})\s+\S", ln)
            lv.append(len(m.group(1)) if m else None)
        return lv

    @staticmethod
    def htext(ln):
        return re.match(r"^#{1,6}\s+(.*?)\s*$", ln).group(1)

    def sections(self):
        lv = self.levels()
        out = [{"heading": None, "level": 0, "lines": []}]
        for ln, l in zip(self.lines, lv):
            if l:
                out.append({"heading": self.htext(ln), "level": l, "lines": []})
            elif ln.strip():
                out[-1]["lines"].append(ln)
        return out

    def stats(self):
        txt = "\n".join(self.lines)
        return {
            "tables": len(re.findall(r"^\|.*\|\s*$", txt, re.M)) // 3 if "|" in txt else 0,
            "bullets": len(re.findall(r"^\s*[-*+]\s+", txt, re.M)),
            "title_style": False,
            "front_matter": bool(self.lines and self.lines[0].strip() == "---"),
        }

    def find(self, heading, occurrence=None):
        lv = self.levels()
        key = norm(heading)
        hits = [i for i, l in enumerate(lv) if l and norm(self.htext(self.lines[i])) == key]
        if not hits:
            avail = [self.htext(self.lines[i]) for i, l in enumerate(lv) if l]
            raise EditError(f'heading not found: "{heading}". Headings in file: {avail}')
        if occurrence:
            if occurrence > len(hits):
                raise EditError(f'heading "{heading}" has only {len(hits)} occurrence(s)')
            return lv, hits[occurrence - 1]
        if len(hits) > 1:
            raise EditError(f'heading "{heading}" appears {len(hits)} times; add "occurrence": 1..{len(hits)}')
        return lv, hits[0]

    _range = staticmethod(DocxDoc._range)

    def section_text(self, heading, occurrence=None):
        lv, h = self.find(heading, occurrence)
        s, e = self._range(lv, h, False)
        return "\n".join(self.lines[s:e]).strip("\n")

    def replace_section(self, heading, text, occurrence=None, allow_media_loss=False):
        lv, h = self.find(heading, occurrence)
        s, e = self._range(lv, h, False)
        old = "\n".join(self.lines[s:e]).strip("\n")
        new = text.replace("\r\n", "\n").strip("\n").split("\n")
        self.lines[s:e] = [""] + new + ([""] if e < len(self.lines) else [])
        return old

    def append_to_section(self, heading, text, occurrence=None):
        lv, h = self.find(heading, occurrence)
        s, e = self._range(lv, h, False)
        body = text.replace("\r\n", "\n").strip("\n").split("\n")
        last = None
        for i in range(e - 1, s - 1, -1):
            if self.lines[i].strip():
                last = i
                break
        if last is None:
            self.lines[h + 1:h + 1] = [""] + body
            return
        list_re = r"^\s*([-*+]|\d+[.)])\s"
        glue = re.match(list_re, self.lines[last]) and re.match(list_re, body[0])
        self.lines[last + 1:last + 1] = (body if glue else [""] + body)

    def add_section(self, heading, level, text, after_heading=None, occurrence=None):
        if after_heading:
            lv, h = self.find(after_heading, occurrence)
            _, e = self._range(lv, h, True)
        else:
            e = len(self.lines)
        ins = e
        while ins > 0 and not self.lines[ins - 1].strip():
            ins -= 1
        block = ["", "#" * level + " " + heading, ""] + text.replace("\r\n", "\n").strip("\n").split("\n")
        if ins == 0:
            block = block[1:]
        if ins == e and e < len(self.lines):
            block.append("")
        self.lines[ins:ins] = block

    def save(self, dest):
        dest = Path(dest)
        data = self.nl.join(self.lines) + (self.nl if self.final_nl else "")
        tmp = dest.with_name(dest.name + ".docsync.tmp")
        tmp.write_bytes(data.encode("utf-8"))
        os.replace(tmp, dest)


def open_doc(path):
    ext = Path(path).suffix.lower()
    if ext == ".docx":
        return DocxDoc(path)
    if ext == ".md":
        return MdDoc(path)
    raise EditError(f"editing {ext or 'this file type'} is not supported (only .docx and .md)")


def make_new_docx(template, title, intro, sections):
    dd = DocxDoc(template if template else None)
    for el in dd.blocks():  # keep styles/headers/footers/page setup, drop the body
        dd.body.remove(el)
    els = []
    if title:
        els.append(dd.styler.title_par(title))
        dd.doc.core_properties.title = title
    if intro:
        els += dd.build(intro)
    for s in sections:
        els.append(dd.styler.heading(int(s.get("level", 1)), s["heading"]))
        els += dd.build(s.get("body", ""))
    dd._insert_after(None, els)
    return dd


def make_new_md(title, intro, sections):
    md = MdDoc()
    lines = []
    if title:
        lines += [f"# {title}", ""]
    if intro:
        lines += [intro.strip("\n"), ""]
    for s in sections:
        lvl = int(s.get("level", 2 if title else 1))
        lines += ["#" * lvl + " " + s["heading"], "", s.get("body", "").strip("\n"), ""]
    md.lines = lines[:-1] if lines else []
    return md


# --------------------------------------------------------------------------
# scanning
# --------------------------------------------------------------------------
def collect_files(root):
    files = []
    for dp, dns, fns in os.walk(root):
        dns[:] = sorted(d for d in dns if d not in SKIP_DIRS and not d.startswith("."))
        for f in sorted(fns):
            if f.startswith(("~$", ".")):
                continue
            if Path(f).suffix.lower() in SUPPORTED:
                files.append(Path(dp) / f)
    return files


def read_doc(path, root):
    ext = path.suffix.lower()
    rel = path.relative_to(root).as_posix()
    if ext == ".docx":
        d = DocxDoc(path)
        return {"rel": rel, "ext": ext, "sections": d.sections(), "stats": d.stats()}
    if ext == ".md":
        d = MdDoc(path)
        return {"rel": rel, "ext": ext, "sections": d.sections(), "stats": d.stats()}
    raw = path.read_bytes().decode("utf-8", "replace")
    return {"rel": rel, "ext": ext,
            "sections": [{"heading": None, "level": 0, "lines": raw.splitlines()}],
            "stats": {"tables": 0, "bullets": 0, "title_style": False, "front_matter": False}}


def find_superseded(docs):
    """Map older versions (X.docx, X_v2.docx) to the newest one."""
    groups = defaultdict(list)
    for d in docs:
        p = Path(d["rel"])
        m = VERSION_RE.match(p.stem)
        base, n = (m.group("base"), int(m.group("n"))) if m else (p.stem, 1)
        groups[(p.parent.as_posix(), base.casefold(), p.suffix.lower())].append((n, d["rel"]))
    out = {}
    for items in groups.values():
        if len(items) > 1:
            items.sort()
            latest = items[-1][1]
            for _, rel in items[:-1]:
                out[rel] = latest
    return out


def write_index(docs, superseded, root, max_chars):
    out = ["# Documentation index", "",
           f"Generated {now()} from `{root}`.",
           "Heading text below is what you use in edits.json (`heading` field).",
           "Edit only the newest version of a document.", ""]
    for d in docs:
        tag = f"  [SUPERSEDED by {superseded[d['rel']]} - do not edit]" if d["rel"] in superseded else ""
        editable = "yes" if d["ext"] in EDITABLE else "no (read-only)"
        out.append(f"## DOC: {d['rel']}{tag}")
        out.append(f"format: {d['ext']} | sections: {len(d['sections'])} | editable: {editable}")
        out.append("")
        for s in d["sections"]:
            if s["heading"] is None:
                if not s["lines"]:
                    continue
                out.append("### (preface - before first heading)")
            else:
                out.append(f"### [H{s['level']}] {s['heading']}")
            text = "\n".join(s["lines"])
            if len(text) > max_chars:
                text = text[:max_chars] + f"\n[... truncated, {len(text) - max_chars} more chars]"
            out.append(text if text.strip() else "(empty)")
            out.append("")
    (STATE / "index.md").write_text("\n".join(out), "utf-8")


def name_style(stem):
    s = VERSION_RE.sub(lambda m: m.group("base"), stem)
    if "_" in s:
        return "lower_snake_case" if s == s.lower() else "Words_With_Underscores"
    if "-" in s:
        return "kebab-case" if s == s.lower() else "Words-With-Dashes"
    if " " in s:
        return "Words with spaces"
    if re.search(r"[a-z][A-Z]", s):
        return "PascalCase/camelCase"
    return "single word"


def write_structure(docs, superseded, root):
    live = [d for d in docs if d["rel"] not in superseded]
    n = max(len(live), 1)
    out = ["# Documentation structure profile", "",
           f"Learned from {len(live)} current documents in `{root}` ({now()}).",
           "Use this when a NEW document is needed: match folder, name, skeleton and style.", ""]

    # folders
    out += ["## Folder layout", ""]
    by_folder = defaultdict(list)
    for d in live:
        by_folder[Path(d["rel"]).parent.as_posix()].append(d)
    for folder in sorted(by_folder):
        files = [Path(d["rel"]).name for d in by_folder[folder]]
        shown = ", ".join(files[:6]) + (f", ... (+{len(files) - 6})" if len(files) > 6 else "")
        out.append(f"- `{'/' if folder == '.' else folder + '/'}` - {len(files)} doc(s): {shown}")
    out.append("")

    # naming
    out += ["## Naming conventions", ""]
    exts = Counter(d["ext"] for d in live)
    out.append("- Formats: " + ", ".join(f"{e} x{c}" for e, c in exts.most_common()))
    styles = Counter(name_style(Path(d["rel"]).stem) for d in live)
    for style, c in styles.most_common():
        ex = [Path(d["rel"]).name for d in live if name_style(Path(d["rel"]).stem) == style][:3]
        out.append(f"- {style}: {c}/{len(live)} files, e.g. {', '.join(ex)}")
    toks = Counter()
    firsts = Counter()
    for d in live:
        parts = [t for t in re.split(r"[_\- ]+", VERSION_RE.sub(lambda m: m.group("base"), Path(d["rel"]).stem)) if t]
        if len(parts) > 1:
            toks[parts[-1].casefold()] += 1
            firsts[parts[0].casefold()] += 1
    common = [f"`{t}` x{c}" for t, c in toks.most_common(4) if c >= 2]
    if common:
        out.append("- Common ending words: " + ", ".join(common))
    common = [f"`{t}` x{c}" for t, c in firsts.most_common(4) if c >= 2]
    if common:
        out.append("- Common starting words: " + ", ".join(common))
    versioned = sum(1 for d in docs if VERSION_RE.match(Path(d["rel"]).stem))
    out.append(f"- Version suffix files (`name_v2.ext`): {versioned}")
    out.append("")

    # headings
    out += ["## Typical heading skeleton", ""]
    seen = defaultdict(lambda: {"docs": 0, "pos": [], "text": "", "levels": Counter()})
    for d in live:
        heads = [s for s in d["sections"] if s["heading"]]
        used = set()
        for i, s in enumerate(heads):
            k = norm(s["heading"])
            if not k or k in used:
                continue
            used.add(k)
            e = seen[k]
            e["docs"] += 1
            e["pos"].append(i / max(len(heads) - 1, 1))
            e["text"] = e["text"] or s["heading"]
            e["levels"][s["level"]] += 1
    threshold = max(2, math.ceil(0.3 * len(live))) if len(live) > 1 else 1
    skeleton = [(sum(e["pos"]) / len(e["pos"]), e) for e in seen.values() if e["docs"] >= threshold]
    skeleton.sort(key=lambda t: t[0])
    if skeleton:
        for i, (_, e) in enumerate(skeleton, 1):
            out.append(f"{i}. {e['text']} (H{e['levels'].most_common(1)[0][0]}, in {e['docs']}/{len(live)} docs)")
        out.append("")
        out.append("Reuse these headings in this order for new documents; skip any that do not apply.")
    else:
        out.append("No shared headings found; follow the nearest sibling document.")
    out.append("")

    # style cues
    out += ["## Style cues", ""]
    lvls = Counter(s["level"] for d in live for s in d["sections"] if s["heading"])
    out.append("- Heading levels used: " + ", ".join(f"H{l} x{c}" for l, c in sorted(lvls.items())))
    out.append(f"- Documents with tables: {sum(1 for d in live if d['stats']['tables'])}/{len(live)}")
    out.append(f"- Documents with bullet lists: {sum(1 for d in live if d['stats']['bullets'])}/{len(live)}")
    out.append(f"- Documents using a Title style: {sum(1 for d in live if d['stats']['title_style'])}/{len(live)}")
    fm = sum(1 for d in live if d["stats"]["front_matter"])
    if fm:
        out.append(f"- Markdown docs with front matter: {fm}")
    words = [sum(len(l.split()) for s in d["sections"] for l in s["lines"]) for d in live]
    out.append(f"- Average length: {sum(words) // n} words")
    out.append("")

    # index-like docs
    out += ["## Index / hub documents (update these when adding a new document)", ""]
    hubs = []
    if len(live) <= 400:
        stems = {d["rel"]: Path(d["rel"]).stem for d in live}
        for d in live:
            text = " ".join(l for s in d["sections"] for l in [s["heading"] or ""] + s["lines"]).casefold()
            hits = sum(1 for r, st in stems.items() if r != d["rel"] and len(st) >= 4
                       and st.casefold().replace("_", " ") in text.replace("_", " "))
            if hits >= 3 or re.search(r"(?i)^(index|toc|contents|readme|home)$", Path(d["rel"]).stem):
                hubs.append((d["rel"], hits))
    out += [f"- `{r}` (references {h} other docs)" for r, h in hubs] or ["- none detected"]
    out.append("")

    # exemplars
    out += ["## Exemplar outlines (nearest siblings to copy from)", ""]
    for folder in sorted(by_folder):
        for d in by_folder[folder][:2]:
            heads = [f"H{s['level']} {s['heading']}" for s in d["sections"] if s["heading"]][:12]
            out.append(f"- `{d['rel']}`: " + (" / ".join(heads) if heads else "(no headings)"))
    out.append("")
    (STATE / "structure.md").write_text("\n".join(out), "utf-8")


def git(repo, *args):
    try:
        r = subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        return r.returncode, r.stdout
    except FileNotFoundError:
        return 127, ""


def build_changes(cfg, since, limit, max_diff, total_cap=80000):
    out = ["# Changes since last run", "", f"Generated {now()}", ""]
    pending = None
    repo = cfg.get("code_repo")
    if repo and Path(repo).is_dir():
        rc, head = git(repo, "rev-parse", "HEAD")
        if rc != 0:
            out += [f"(Could not read git history in `{repo}`.)", ""]
        else:
            head = head.strip()
            pending = head
            base = since or cfg.get("last_commit")
            if base and git(repo, "cat-file", "-e", f"{base}^{{commit}}")[0] != 0:
                out += [f"(Reference `{base}` not found; showing the latest {limit} commits.)", ""]
                base = None
            if base == head:
                out += [f"No new commits since `{base[:8]}`.", ""]
            else:
                rng = [f"{base}..HEAD"] if base else ["-n", str(limit)]
                rc, log = git(repo, "log", *rng, "--reverse", "--no-merges",
                              "--format=%H%x1f%an%x1f%ad%x1f%B%x1e", "--date=short")
                commits = [c.strip("\n") for c in log.split("\x1e") if c.strip()]
                header = f"since `{base[:8]}`" if base else f"latest {limit} commits (first run)"
                out += [f"## Code commits ({len(commits)}) {header}", f"Repo: `{repo}`", ""]
                used = 0
                for c in commits:
                    h, an, ad, msg = (c.split("\x1f") + ["", "", "", ""])[:4]
                    h = h.strip()
                    out += [f"### {h[:8]} - {msg.strip().splitlines()[0] if msg.strip() else '(no message)'}",
                            f"{an}, {ad}", "", msg.strip(), ""]
                    _, stat = git(repo, "show", "--stat", "--format=", "--no-color", h)
                    out += ["Files:", "```", stat.strip(), "```", ""]
                    _, patch = git(repo, "show", "--format=", "--no-color", "-U2", h, "--", ".", *LOCKFILE_EXCLUDES)
                    if len(patch) > max_diff:
                        patch = patch[:max_diff] + f"\n[... diff truncated, {len(patch) - max_diff} more chars]"
                    used += len(patch)
                    if used > total_cap:
                        out += ["[diff omitted: total size cap reached; use the file list above]", ""]
                    else:
                        out += ["```diff", patch.strip(), "```", ""]
    else:
        out += ["No code repo configured (set one with `python docsync.py init --code-repo <path>`).", ""]
    if NOTES.exists() and NOTES.read_text("utf-8").strip():
        out += ["## Manual change notes (from the user)", "", NOTES.read_text("utf-8").strip(), ""]
    body = "\n".join(out)
    if "## Code commits" not in body and "## Manual" not in body:
        if pending:
            out += ["**Nothing new was found in git.** If the work is not in git (or was already "
                    "documented), ask the user to describe the change, then run "
                    '`python docsync.py note "<description>"` and scan again.', ""]
        else:
            out += ["**No change source available.** Ask the user to describe what changed, then run "
                    '`python docsync.py note "<description>"` and scan again.', ""]
    (STATE / "changes.md").write_text("\n".join(out), "utf-8")
    return pending


# --------------------------------------------------------------------------
# commands
# --------------------------------------------------------------------------
def cmd_init(a):
    cfg = load_config(False)
    if cfg.get("docs_root") and not a.root and not a.reset:
        print(f"Already configured.\n  docs root : {cfg['docs_root']}\n  code repo : {cfg.get('code_repo') or '(none)'}"
              f"\n  mode      : {cfg.get('mode') or '(not chosen yet)'}\n"
              "Use `init --reset` or `init --root <path>` to change.")
        return
    root = a.root
    if not root:
        if not sys.stdin.isatty():
            die("NEEDS_INPUT: the documents root folder is not known. Ask the user for the folder path, "
                'then run: python docsync.py init --root "<path>" [--code-repo "<path>"]')
        root = input("Root folder where your documents are stored: ").strip().strip('"')
    p = Path(root).expanduser().resolve()
    if not p.is_dir():
        die(f"Not a folder: {p}")
    cfg["docs_root"] = str(p)
    repo = a.code_repo
    if repo is None and sys.stdin.isatty() and not cfg.get("code_repo"):
        repo = input("Code repository path for change detection (Enter to skip): ").strip().strip('"')
    if repo:
        rp = Path(repo).expanduser().resolve()
        if not rp.is_dir():
            die(f"Code repo not found: {rp}")
        cfg["code_repo"] = str(rp)
    cfg.pop("last_commit", None) if a.reset else None
    save_config(cfg)
    print(f"Saved. Docs root: {p}\nCode repo: {cfg.get('code_repo') or '(none)'}\nNext: python docsync.py scan")


def cmd_scan(a):
    cfg = load_config()
    root = docs_root(cfg)
    STATE.mkdir(parents=True, exist_ok=True)
    docs, skipped = [], []
    for f in collect_files(root):
        try:
            docs.append(read_doc(f, root))
        except Exception as e:  # corrupt / locked / password-protected
            skipped.append((f.relative_to(root).as_posix(), str(e)))
    superseded = find_superseded(docs)
    write_index(docs, superseded, root, a.max_section_chars)
    write_structure(docs, superseded, root)
    pending = build_changes(cfg, a.since, a.limit, a.max_diff_chars)
    cfg["pending_head"] = pending
    save_config(cfg)
    print(f"Scanned {len(docs)} documents ({len(superseded)} superseded versions)")
    for rel, err in skipped:
        print(f"  skipped {rel}: {err}")
    print(f"Wrote:\n  {STATE / 'index.md'}\n  {STATE / 'structure.md'}\n  {STATE / 'changes.md'}")
    if not cfg.get("mode"):
        print("Mode not chosen yet: ask the user, then run `python docsync.py set-mode new-version|in-place`")


def cmd_note(a):
    STATE.mkdir(parents=True, exist_ok=True)
    with NOTES.open("a", encoding="utf-8") as f:
        f.write(f"- [{now()}] {a.text.strip()}\n")
    print(f"Note saved to {NOTES}")


def cmd_set_mode(a):
    cfg = load_config()
    cfg["mode"] = a.mode
    save_config(cfg)
    print(f"Mode set to: {a.mode}")


def cmd_status(a):
    cfg = load_config(False)
    if not cfg:
        print("Not configured. Run: python docsync.py init")
        return
    print(json.dumps(cfg, indent=2))
    print("Manual notes pending:", "yes" if NOTES.exists() and NOTES.read_text("utf-8").strip() else "no")


def safe_path(root, rel):
    root = root.resolve()
    p = (root / rel).resolve()
    try:
        parts = p.relative_to(root).parts
    except ValueError:
        raise EditError(f"path escapes the docs root: {rel}")
    if WORK_DIR_NAME in parts:
        raise EditError(f"path is inside the reserved {WORK_DIR_NAME} folder: {rel}")
    return p


def next_version_path(path):
    m = VERSION_RE.match(path.stem)
    base, sep, n = (m.group("base"), m.group("sep"), int(m.group("n")) + 1) if m else (path.stem, "_", 2)
    while True:
        cand = path.with_name(f"{base}{sep}v{n}{path.suffix}")
        if not cand.exists():
            return cand
        n += 1


def cmd_apply(a):
    cfg = load_config()
    root = docs_root(cfg)
    mode = a.mode or cfg.get("mode")
    if not mode:
        die("NEEDS_INPUT: mode not chosen. Ask the user 'new version or update existing?', then run "
            "`python docsync.py set-mode new-version|in-place`")
    try:
        spec = json.loads(Path(a.edits).read_text("utf-8-sig"))
    except Exception as e:
        die(f"Cannot read edits file: {e}")
    edits = spec.get("edits") if isinstance(spec, dict) else spec
    if not isinstance(edits, list) or not edits:
        die('edits file needs a non-empty "edits" list')
    task = (spec.get("task") if isinstance(spec, dict) else None) or "(no task label)"

    opened, created, results, errors = {}, {}, [], []

    def get_doc(p):
        if p not in opened:
            if not p.is_file():
                raise EditError(f"file not found: {p.relative_to(root).as_posix()}")
            opened[p] = open_doc(p)
        return opened[p]

    for i, e in enumerate(edits, 1):
        label = f"edit #{i} ({e.get('type', '?')} {e.get('path', '?')})"
        try:
            t, rel = e.get("type"), e.get("path")
            if not t or not rel:
                raise EditError('"type" and "path" are required')
            p = safe_path(root, rel)
            reason = e.get("reason", "")
            occ = e.get("occurrence")

            if t == "new_document":
                if p.suffix.lower() not in EDITABLE:
                    raise EditError("new documents must be .docx or .md")
                if p.exists() or p in created:
                    raise EditError("file already exists (choose another name, or use update/add edits)")
                if p in opened:
                    raise EditError("cannot edit and create the same file in one batch")
                secs = e.get("sections", [])
                if any("heading" not in s for s in secs):
                    raise EditError("every section needs a heading")
                note = ""
                if p.suffix.lower() == ".docx":
                    tp = safe_path(root, e["template_from"]) if e.get("template_from") else None
                    if tp and (not tp.is_file() or tp.suffix.lower() != ".docx"):
                        raise EditError(f"template_from is not an existing .docx: {e['template_from']}")
                    if not tp:
                        note = " (no template_from: default Word styling used)"
                    doc = make_new_docx(tp, e.get("title"), e.get("intro"), secs)
                else:
                    doc = make_new_md(e.get("title"), e.get("intro"), secs)
                created[p] = doc
                results.append({"type": "NEW", "path": rel, "summary": f"{len(secs)} section(s){note}",
                                "reason": reason, "diff": ""})
                continue

            doc = get_doc(p)
            if t == "update_section":
                old = doc.replace_section(e["heading"], e["new_text"], occ, bool(e.get("allow_media_loss")))
                diff = "\n".join(difflib.unified_diff(old.splitlines(), e["new_text"].strip().splitlines(),
                                                      "before", "after", lineterm="", n=1))
                results.append({"type": "UPDATE", "path": rel, "summary": f'section "{e["heading"]}"',
                                "reason": reason, "diff": diff})
            elif t == "add_section":
                doc.add_section(e["heading"], int(e.get("level", 2)), e.get("body", ""),
                                e.get("after_heading"), occ)
                where = f' after "{e["after_heading"]}"' if e.get("after_heading") else " at end"
                results.append({"type": "ADD SECTION", "path": rel,
                                "summary": f'"{e["heading"]}"{where}', "reason": reason, "diff": ""})
            elif t == "append_to_section":
                doc.append_to_section(e["heading"], e["body"], occ)
                results.append({"type": "APPEND", "path": rel, "summary": f'to section "{e["heading"]}"',
                                "reason": reason, "diff": ""})
            else:
                raise EditError(f"unknown edit type: {t}")
        except KeyError as ke:
            errors.append(f"{label}: missing field {ke}")
        except EditError as ex:
            errors.append(f"{label}: {ex}")
        except Exception as ex:  # unexpected - still abort cleanly
            errors.append(f"{label}: {type(ex).__name__}: {ex}")

    if errors:
        print("Nothing was written. Fix these and run again:")
        for m in errors:
            print("  -", m)
        sys.exit(2)

    print(f"Task: {task}\nMode: {mode}{'  (dry run - nothing written)' if a.dry_run else ''}\n")
    for r in results:
        print(f"[{r['type']}] {r['path']}: {r['summary']}")
        if r["reason"]:
            print(f"    reason: {r['reason']}")
        if r["diff"]:
            print("    " + r["diff"].replace("\n", "\n    "))
    if a.dry_run:
        return

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    written = []
    out_names = {}
    try:
        for p, doc in opened.items():
            rel = p.relative_to(root).as_posix()
            if mode == "new-version":
                dest = next_version_path(p)
            else:
                backup = root / WORK_DIR_NAME / "backups" / stamp / rel
                backup.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(p, backup)
                dest = p
            doc.save(dest)
            out_names[rel] = dest.relative_to(root).as_posix()
            written.append(f"{rel} -> {out_names[rel]}" if dest != p else f"{rel} (updated, backup kept)")
        for p, doc in created.items():
            p.parent.mkdir(parents=True, exist_ok=True)
            doc.save(p)
            written.append(f"{p.relative_to(root).as_posix()} (new)")
    except PermissionError as ex:
        die(f"Could not write {ex.filename}. Close the file in Word/OneDrive and re-run.")

    log = root / WORK_DIR_NAME / "changelog.md"
    log.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"## {now()} - {task}", f"Mode: {mode}", ""]
    for r in results:
        path = out_names.get(r["path"], r["path"])
        why = f" - {r['reason']}" if r["reason"] else ""
        lines.append(f"- {r['type']}: `{path}` {r['summary']}{why}")
    lines.append("")
    with log.open("a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    if cfg.get("pending_head"):
        cfg["last_commit"] = cfg["pending_head"]
        save_config(cfg)
    if NOTES.exists():
        NOTES.write_text("", "utf-8")
    print("\nWritten:")
    for w in written:
        print("  " + w)
    print(f"Changelog: {log}")


def main():
    ap = argparse.ArgumentParser(prog="docsync", description="Doc Sync Agent helper")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("init", help="save docs root (asks once)")
    p.add_argument("--root", help="folder where documents are stored")
    p.add_argument("--code-repo", help="git repo whose commits drive doc changes (optional)")
    p.add_argument("--reset", action="store_true", help="re-ask / start over")
    p.set_defaults(fn=cmd_init)

    p = sub.add_parser("scan", help="index docs, learn structure, collect changes")
    p.add_argument("--since", help="git ref to diff from (default: last applied commit)")
    p.add_argument("--limit", type=int, default=10, help="commits to read on first run (default 10)")
    p.add_argument("--max-diff-chars", type=int, default=4000, help="diff size cap per commit")
    p.add_argument("--max-section-chars", type=int, default=3000, help="section text cap in index.md")
    p.set_defaults(fn=cmd_scan)

    p = sub.add_parser("note", help="record a change manually")
    p.add_argument("text")
    p.set_defaults(fn=cmd_note)

    p = sub.add_parser("set-mode", help="new-version or in-place")
    p.add_argument("mode", choices=["new-version", "in-place"])
    p.set_defaults(fn=cmd_set_mode)

    p = sub.add_parser("apply", help="apply an edits.json")
    p.add_argument("edits", help="path to edits.json")
    p.add_argument("--dry-run", action="store_true", help="validate and show diffs, write nothing")
    p.add_argument("--mode", choices=["new-version", "in-place"], help="override saved mode")
    p.set_defaults(fn=cmd_apply)

    p = sub.add_parser("status", help="show configuration")
    p.set_defaults(fn=cmd_status)

    args = ap.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()
