#!/usr/bin/env python3
"""Minimal Markdown -> .docx, enough for agent-written drafts.

Supports: # / ## / ### headings, paragraphs, **bold**, *italic*, `code`, bullet and numbered
lists, images on their own line (![caption](path)), simple pipe tables, and a
{BIBLIOGRAPHY} paragraph. Citation placeholders ({@KEY}) are kept as plain text for
zotero_docx.py. For anything richer, build the .docx your own way (pandoc, Word, a journal
template): zotero_docx.py works on any .docx that contains the placeholders.

Run:  python3 scripts/md_to_docx.py draft.md draft.docx
"""
import os
import re
import sys

from docx import Document
from docx.shared import Cm, Pt

INLINE = re.compile(r"(\*\*[^*]+\*\*|\*[^*\s][^*]*\*|`[^`]+`)")
IMAGE = re.compile(r"^!\[([^\]]*)\]\(([^)\s]+)[^)]*\)\s*$")


def add_inline(par, text):
    for part in INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            par.add_run(part[2:-2]).bold = True
        elif part.startswith("`") and part.endswith("`"):
            run = par.add_run(part[1:-1])
            run.font.name = "Courier New"
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            par.add_run(part[1:-1]).italic = True
        else:
            par.add_run(part)


def split_row(line):
    cells, depth, cur = [], 0, ""
    for ch in line.strip().strip("|"):
        depth += ch == "{"
        depth -= ch == "}"
        if ch == "|" and depth == 0:
            cells.append(cur.strip())
            cur = ""
        else:
            cur += ch
    cells.append(cur.strip())
    return cells


def markdown_to_document(md, base_dir="."):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    lines = md.splitlines()
    i, para = 0, []

    def flush():
        if para:
            add_inline(doc.add_paragraph(), " ".join(s.strip() for s in para))
            del para[:]

    while i < len(lines):
        line = lines[i]
        s = line.strip()
        h = re.match(r"^(#{1,4})\s+(.*)$", s)
        if not s:
            flush()
        elif h:
            flush()
            doc.add_heading(h.group(2).strip(), level=len(h.group(1)))
        elif s == "{BIBLIOGRAPHY}":
            flush()
            doc.add_paragraph("{BIBLIOGRAPHY}")
        elif IMAGE.match(s):
            flush()
            m = IMAGE.match(s)
            path = m.group(2) if os.path.isabs(m.group(2)) else os.path.join(base_dir, m.group(2))
            if os.path.exists(path):
                doc.add_picture(path, width=Cm(15))
            else:
                doc.add_paragraph("[missing image: %s]" % m.group(2))
            if m.group(1):
                cap = doc.add_paragraph()
                add_inline(cap, m.group(1))
        elif re.match(r"^[-*+]\s+", s):
            flush()
            add_inline(doc.add_paragraph(style="List Bullet"), re.sub(r"^[-*+]\s+", "", s))
        elif re.match(r"^\d+[.)]\s+", s):
            flush()
            add_inline(doc.add_paragraph(style="List Number"), re.sub(r"^\d+[.)]\s+", "", s))
        elif s.startswith("|") and i + 1 < len(lines) and re.match(r"^\|?\s*:?-{3,}", lines[i + 1].strip()):
            flush()
            rows = [split_row(s)]
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(split_row(lines[i]))
                i += 1
            ncol = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=ncol)
            table.style = "Table Grid"
            for r, row in enumerate(rows):
                for c, val in enumerate(row):
                    cell = table.cell(r, c)
                    cell.paragraphs[0].text = ""
                    add_inline(cell.paragraphs[0], val)
                    if r == 0:
                        for run in cell.paragraphs[0].runs:
                            run.bold = True
            continue
        else:
            para.append(line)
        i += 1
    flush()
    return doc


def main():
    if len(sys.argv) != 3:
        sys.exit("usage: md_to_docx.py input.md output.docx")
    with open(sys.argv[1], encoding="utf-8") as f:
        doc = markdown_to_document(f.read(), os.path.dirname(os.path.abspath(sys.argv[1])))
    doc.save(sys.argv[2])
    print("wrote " + sys.argv[2])


if __name__ == "__main__":
    main()
