#!/usr/bin/env python3
"""Turn citation placeholders into live Zotero citations in a Word file.

Input is a .docx (placeholders typed anywhere in the text) or a Markdown file (converted
with md_to_docx.py first). Output is a .docx that the Zotero Word plugin treats exactly as
if you had inserted every citation by hand: click Zotero -> Refresh once, then change the
citation style whenever you like in Zotero -> Document Preferences.

Placeholders
  {@KEY}                 parenthetical       (Smith et al., 2020)
  {@KEY1; @KEY2}         several works       (Lee, 2019; Smith et al., 2020)
  {@KEY, p. 12}          with a locator      (Smith et al., 2020, p. 12)
  {@KEY|n}               narrative           Smith et al. (2020)
  {BIBLIOGRAPHY}         alone in a paragraph: the Zotero bibliography

KEY is a Zotero item key or one of your own ids from --sources (see zotero_link.py).

Run:
  python3 scripts/zotero_docx.py draft.md  draft.docx
  python3 scripts/zotero_docx.py draft.docx draft_zotero.docx --style http://www.zotero.org/styles/ieee
"""
import argparse
import copy
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PLACEHOLDER = re.compile(r"\{(@[^{}|]+?)(\|n)?\}")
ITEM = re.compile(r"^@([A-Za-z0-9_:-]+)(?:\s*,\s*(.+))?$")
SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
LABELS = {"p.": "page", "pp.": "page", "page": "page", "pages": "page", "ch.": "chapter",
          "chap.": "chapter", "sec.": "section", "fig.": "figure", "figs.": "figure",
          "tab.": "table", "para.": "paragraph", "vol.": "volume", "eq.": "equation"}


def parse_placeholder(inner):
    """'@A, p. 12; @B' -> [('A', 'page', '12'), ('B', None, None)], or None if malformed."""
    items = []
    for part in inner.split(";"):
        m = ITEM.match(part.strip())
        if not m:
            return None
        label = locator = None
        if m.group(2):
            loc = m.group(2).strip()
            first = loc.split(None, 1)
            if len(first) == 2 and first[0].lower() in LABELS:
                label, locator = LABELS[first[0].lower()], first[1]
            else:
                label, locator = "page", loc
        items.append((m.group(1), label, locator))
    return items


def short_id(*parts):
    return hashlib.sha1("|".join(parts).encode()).hexdigest()[:8]


# ------------------------------------------------------------------ APA pre-rendering
# Only what Word shows before the first Zotero refresh. Zotero re-renders everything.
def family(n):
    return n.get("family") or n.get("literal") or "?"


def year_of(csl):
    try:
        return str(csl["issued"]["date-parts"][0][0])
    except Exception:
        return "n.d."


def apa_author(csl):
    names = csl.get("author") or csl.get("editor") or []
    if not names:
        return (csl.get("title") or "?")[:40]
    if len(names) == 1:
        return family(names[0])
    if len(names) == 2:
        return "%s & %s" % (family(names[0]), family(names[1]))
    return "%s et al." % family(names[0])


def narrative_author(csl):
    """Author text written before a narrative citation: 'Lee and Kim', not 'Lee & Kim'."""
    return apa_author(csl).replace(" & ", " and ")


def apa_cite_one(csl, label=None, locator=None):
    s = "%s, %s" % (apa_author(csl), year_of(csl))
    if locator:
        short = {"page": "pp." if re.search(r"[-–,]", locator) else "p."}.get(label, label or "")
        s += ", %s %s" % (short, locator)
    return s


def apa_reference(csl):
    names = csl.get("author") or []

    def fmt(n):
        if "literal" in n:
            return n["literal"]
        initials = " ".join(p[0] + "." for p in re.split(r"[\s-]+", n.get("given", "")) if p)
        return "%s, %s" % (n.get("family", "?"), initials) if initials else n.get("family", "?")
    if len(names) > 20:
        names = names[:19] + [names[-1]]
    auth = ", ".join(fmt(n) for n in names[:-1])
    auth = (auth + ", & " + fmt(names[-1])) if len(names) > 1 else (fmt(names[0]) if names else "")
    parts = ["%s (%s)." % (auth, year_of(csl)), "%s." % (csl.get("title") or "").rstrip(".")]
    src = csl.get("container-title") or ""
    if csl.get("volume"):
        src += ", %s" % csl["volume"] + ("(%s)" % csl["issue"] if csl.get("issue") else "")
    if csl.get("page"):
        src += ", %s" % csl["page"]
    if src:
        parts.append(src + ".")
    if csl.get("DOI"):
        parts.append("https://doi.org/" + csl["DOI"])
    return " ".join(p for p in parts if p.strip(" ."))


# ------------------------------------------------------------------ field XML
def make_run(src_run, text=None, fld=None, instr=None):
    r = OxmlElement("w:r")
    rpr = src_run.find(qn("w:rPr")) if src_run is not None else None
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    if fld:
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), fld)
    elif instr is not None:
        el = OxmlElement("w:instrText")
        el.set(qn("xml:space"), "preserve")
        el.text = instr
    else:
        el = OxmlElement("w:t")
        el.set(qn("xml:space"), "preserve")
        el.text = text
    r.append(el)
    return r


def field_runs(src_run, instr, shown):
    return [make_run(src_run, fld="begin"), make_run(src_run, instr=instr),
            make_run(src_run, fld="separate"), make_run(src_run, text=shown),
            make_run(src_run, fld="end")]


ZOTERO_KEY_CHARS = "23456789ABCDEFGHIJKLMNPQRSTUVWXYZ"


def library_prefix(zmap):
    """'http://zotero.org/users/local/XXXXXXXX', taken from any linked item."""
    for e in zmap.values():
        if e.get("uri"):
            return e["uri"].rsplit("/items/", 1)[0]
    return "http://zotero.org/users/local/embedded"


def placeholder_uri(prefix, rid):
    """A well-formed item URI that points to no item.

    The Zotero Word plugin needs a URI on every citation item. If it gets none, it reads the
    item id as a numeric database id, fails, and shows "Zotero experienced an error updating
    your document". With a URI that matches nothing, it falls back to the embedded itemData,
    which is what we want for works that are not in your library."""
    h = int(hashlib.sha1(("embedded|" + rid).encode()).hexdigest(), 16)
    key = ""
    for _ in range(8):
        h, r = divmod(h, len(ZOTERO_KEY_CHARS))
        key += ZOTERO_KEY_CHARS[r]
    return "%s/items/%s" % (prefix, key)


def citation_instr(items, zmap, counter, narrative=False):
    cites, shown = [], []
    prefix = library_prefix(zmap)
    for rid, label, locator in items:
        entry = zmap[rid]
        csl = dict(entry["csl"])
        uri = entry.get("uri") or placeholder_uri(prefix, rid)
        csl["id"] = uri
        item = {"id": uri, "uris": [uri], "itemData": csl}
        if locator:
            item["locator"], item["label"] = locator, label
        if narrative:
            item["suppress-author"] = True
        cites.append(item)
        if narrative:
            shown.append(apa_cite_one(csl, label, locator).split(", ", 1)[1])
        else:
            shown.append(apa_cite_one(csl, label, locator))
    text = "(%s)" % "; ".join(shown)
    payload = {"citationID": short_id("cite", str(counter), *[i[0] for i in items]),
               "properties": {"formattedCitation": text, "plainCitation": text, "noteIndex": 0},
               "citationItems": cites, "schema": SCHEMA}
    return " ADDIN ZOTERO_ITEM CSL_CITATION %s " % json.dumps(payload, ensure_ascii=False), text


# ------------------------------------------------------------------ document walk
def all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def merge_split_placeholders(p):
    """Word often splits typed text into several runs (spell check, autocorrect, edits).
    Join the runs a placeholder spans, so each placeholder sits in one run."""
    while True:
        runs = p.runs
        spans, pos = [], 0
        for r in runs:
            spans.append((pos, pos + len(r.text)))
            pos += len(r.text)
        text = "".join(r.text for r in runs)
        for m in PLACEHOLDER.finditer(text):
            first = next(i for i, (a, b) in enumerate(spans) if a <= m.start() < b)
            last = next(i for i, (a, b) in enumerate(spans) if a < m.end() <= b)
            if first != last:
                runs[first].text = "".join(r.text for r in runs[first:last + 1])
                for r in runs[first + 1:last + 1]:
                    r.text = ""
                break
        else:
            return


def process(doc, zmap):
    cited, counter, bib_paras, problems = [], 0, [], []
    for p in all_paragraphs(doc):
        if p.text.strip() == "{BIBLIOGRAPHY}":
            bib_paras.append(p)
            continue
        if "{@" not in p.text:
            continue
        merge_split_placeholders(p)
        for run in list(p.runs):
            text = run.text
            if not PLACEHOLDER.search(text):
                continue
            r_el, new, pos = run._r, [], 0
            for m in PLACEHOLDER.finditer(text):
                items = parse_placeholder(m.group(1))
                unknown = [i[0] for i in items or [] if i[0] not in zmap]
                if items is None or unknown:
                    problems.append(m.group(0) + ("  (unknown id: %s)" % ", ".join(unknown) if unknown
                                                  else "  (malformed)"))
                    continue  # leave the placeholder in the text so it is easy to find
                narrative = bool(m.group(2)) and len(items) == 1
                counter += 1
                instr, shown = citation_instr(items, zmap, counter, narrative)
                before = text[pos:m.start()]
                if narrative:
                    before += narrative_author(zmap[items[0][0]]["csl"]) + " "
                if before:
                    new.append(make_run(r_el, text=before))
                new.extend(field_runs(r_el, instr, shown))
                for i in items:
                    if i[0] not in cited:
                        cited.append(i[0])
                pos = m.end()
            if not new:
                continue
            if pos < len(text):
                new.append(make_run(r_el, text=text[pos:]))
            for el in new:
                r_el.addprevious(el)
            r_el.getparent().remove(r_el)

    for p in bib_paras:
        src = p.runs[0]._r if p.runs else None
        refs = sorted((apa_reference(zmap[i]["csl"]) for i in cited), key=lambda s: s.lower())
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        instr = ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '
        runs = [make_run(src, fld="begin"), make_run(src, instr=instr), make_run(src, fld="separate")]
        for k, ref in enumerate(refs):
            runs.append(make_run(src, text=ref))
            if k < len(refs) - 1:
                br = OxmlElement("w:r")
                br.append(OxmlElement("w:br"))
                runs.append(br)
        runs.append(make_run(src, fld="end"))
        for el in runs:
            p._p.append(el)
    return cited, counter, problems, len(bib_paras)


# ------------------------------------------------------------------ Zotero document prefs
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"


def add_zotero_prefs(path, style, locale="en-US"):
    """Store Zotero's document preferences (style, field type) as ZOTERO_PREF_n custom
    properties, the way the plugin itself does, so Word opens with the chosen style."""
    pref = ('<data data-version="3" zotero-version="7.0"><session id="%s"/>'
            '<style id="%s" locale="%s" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
            '<prefs><pref name="fieldType" value="Field"/>'
            '<pref name="automaticJournalAbbreviations" value="true"/></prefs></data>'
            ) % (short_id(path), style, locale)
    chunks = [pref[i:i + 240] for i in range(0, len(pref), 240)]
    esc = lambda s: s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    props = "".join(
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="ZOTERO_PREF_%d">'
        '<vt:lpwstr>%s</vt:lpwstr></property>' % (k + 2, k + 1, esc(c)) for k, c in enumerate(chunks))
    custom = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Properties xmlns="%s" xmlns:vt="%s">%s</Properties>' % (CUSTOM_NS, VT_NS, props))

    fd, tmp = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name == "docProps/custom.xml":
                continue  # replaced below (a document's own custom properties are dropped)
            if name == "[Content_Types].xml" and b"docProps/custom.xml" not in data:
                data = data.replace(b"</Types>", b'<Override PartName="/docProps/custom.xml" ContentType='
                                    b'"application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')
            if name == "_rels/.rels" and CUSTOM_REL.encode() not in data:
                data = data.replace(b"</Relationships>", ('<Relationship Id="rIdZoteroCustom" Type="%s" '
                                    'Target="docProps/custom.xml"/></Relationships>' % CUSTOM_REL).encode())
            zout.writestr(name, data)
        zout.writestr("docProps/custom.xml", custom)
    shutil.move(tmp, path)


def main():
    ap = argparse.ArgumentParser(description="Placeholders -> live Zotero citations in .docx")
    ap.add_argument("input", help=".docx or .md")
    ap.add_argument("output", help=".docx")
    ap.add_argument("--map", default="citations/zotero_map.json", help="from zotero_link.py")
    ap.add_argument("--style", default="http://www.zotero.org/styles/apa",
                    help="CSL style id, e.g. http://www.zotero.org/styles/ieee")
    ap.add_argument("--locale", default="en-US")
    args = ap.parse_args()
    if not os.path.exists(args.map):
        sys.exit("No %s. Run scripts/zotero_link.py first." % args.map)
    with open(args.map, encoding="utf-8") as f:
        zmap = json.load(f)

    if args.input.lower().endswith((".md", ".markdown", ".txt")):
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from md_to_docx import markdown_to_document
        with open(args.input, encoding="utf-8") as f:
            doc = markdown_to_document(f.read(), base_dir=os.path.dirname(os.path.abspath(args.input)))
    else:
        doc = Document(args.input)

    cited, n, problems, n_bib = process(doc, zmap)
    doc.save(args.output)
    add_zotero_prefs(args.output, args.style, args.locale)

    linked = sum(1 for i in cited if zmap[i].get("uri"))
    print("%s: %d citations of %d distinct works (%d linked to Zotero items, %d embedded)." % (
        args.output, n, len(cited), linked, len(cited) - linked))
    if not n_bib:
        print("Note: no {BIBLIOGRAPHY} paragraph, so no reference list (add one, or insert it in Word).")
    for p in problems:
        print("NOT CONVERTED: " + p)
    print("Next: open in Word -> Zotero tab -> Refresh. Change style: Zotero -> Document Preferences.")
    if problems:
        sys.exit(1)


if __name__ == "__main__":
    main()
