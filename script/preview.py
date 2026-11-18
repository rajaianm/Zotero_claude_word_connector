#!/usr/bin/env python3
"""Readable copy of a Markdown draft: placeholders become author-year citations that link to
the item in Zotero (zotero://select/...) or, for works not in Zotero, to the PDF given in
--sources. The file ends with the list of sources it cites. The draft is not modified.

    drafts/intro.md  ->  drafts/preview/intro.md

Open the preview in VS Code's Markdown preview (or any Markdown viewer) and click a citation.

Run:  python3 scripts/preview.py drafts/intro.md [more.md ...] [--map citations/zotero_map.json]
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from zotero_docx import PLACEHOLDER, apa_cite_one, narrative_author, parse_placeholder  # noqa: E402


def target(entry, out_dir):
    if entry.get("key"):
        return "zotero://select/library/items/%s" % entry["key"]
    if entry.get("pdf"):
        return os.path.relpath(os.path.abspath(entry["pdf"]), out_dir)
    return None


def link(text, entry, out_dir):
    url = target(entry, out_dir)
    if not url:
        return text
    title = (entry["csl"].get("title") or "").replace('"', "'")[:150]
    return '[%s](<%s> "%s")' % (text, url, title)


def render(md, zmap, out_dir, cited):
    def sub(m):
        items = parse_placeholder(m.group(1))
        if items is None:
            return "**[MALFORMED %s]**" % m.group(0)
        missing = [i[0] for i in items if i[0] not in zmap]
        if missing:
            return "**[UNKNOWN %s]**" % ", ".join(missing)
        for i in items:
            if i[0] not in cited:
                cited.append(i[0])
        if m.group(2) and len(items) == 1:
            rid, label, loc = items[0]
            e = zmap[rid]
            inner = apa_cite_one(e["csl"], label, loc).split(", ", 1)[1]
            return "%s (%s)" % (link(narrative_author(e["csl"]), e, out_dir), inner)
        return "(%s)" % "; ".join(link(apa_cite_one(zmap[r]["csl"], lab, loc), zmap[r], out_dir)
                                  for r, lab, loc in items)
    return PLACEHOLDER.sub(sub, md)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--map", default="citations/zotero_map.json")
    ap.add_argument("--out", help="output folder (default: <draft folder>/preview)")
    args = ap.parse_args()
    with open(args.map, encoding="utf-8") as f:
        zmap = json.load(f)
    for path in args.files:
        out_dir = args.out or os.path.join(os.path.dirname(os.path.abspath(path)), "preview")
        os.makedirs(out_dir, exist_ok=True)
        with open(path, encoding="utf-8") as f:
            md = f.read()
        cited = []
        body = render(md, zmap, out_dir, cited).replace(
            "{BIBLIOGRAPHY}", "*(the Zotero bibliography goes here in the Word file)*")
        lines = ["", "---", "", "### Sources cited (%d)" % len(cited), ""]
        for rid in cited:
            e = zmap[rid]
            how = "in Zotero" if e.get("key") else "not in Zotero (embedded)"
            lines.append("- **%s** %s — %s — %s" % (rid, link(apa_cite_one(e["csl"]), e, out_dir),
                                                     e["csl"].get("title") or "", how))
        out = os.path.join(out_dir, os.path.basename(path))
        with open(out, "w", encoding="utf-8") as f:
            f.write(body.rstrip() + "\n" + "\n".join(lines) + "\n")
        print("%s  (%d sources)" % (out, len(cited)))


if __name__ == "__main__":
    main()
