#!/usr/bin/env python3
"""Build the citation map that the writing agent and zotero_docx.py use.

It reads your Zotero library (or one collection) through Zotero's local API and writes:

  citations/zotero_map.json   id -> {key, uri, match, csl}. Every Zotero item is listed under
                              its item key (e.g. "8XK2M4QZ"). If you pass --sources, your own
                              ids (e.g. "P017") are added and matched to Zotero items by DOI,
                              then by title (+ year).
  citations/library.md        a compact list of what can be cited. Give this to the agent.
  citations/unmatched.txt     sources that are not in Zotero (they still cite correctly; their
                              data is embedded in the Word file).
  citations/unmatched.ris     the same sources as RIS, to import into Zotero if you like
                              (then run this script again to link them).

Zotero 7: Settings -> Advanced -> tick "Allow other applications on this computer to
communicate with Zotero". Nothing leaves your computer.

Examples:
  python3 scripts/zotero_link.py                                  # whole library
  python3 scripts/zotero_link.py --collection "My review"          # one collection
  python3 scripts/zotero_link.py --sources examples/sources.json   # + your own ids
  python3 scripts/zotero_link.py --sources examples/sources.json --offline   # no Zotero needed
"""
import argparse
import json
import os
import re
import sys
import urllib.parse
import urllib.request

API = "http://127.0.0.1:23119/api"


# ------------------------------------------------------------------ Zotero local API
def get(url):
    with urllib.request.urlopen(url, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def paged(url):
    out, start = [], 0
    sep = "&" if "?" in url else "?"
    while True:
        page = get("%s%slimit=100&start=%d" % (url, sep, start))
        if isinstance(page, dict):
            page = page.get("items", [])
        if not page:
            break
        out.extend(page)
        start += len(page)
        if len(page) < 100:
            break
    return out


def find_collection(api, name):
    cols = paged("%s/users/0/collections" % api)
    hits = [c for c in cols if c.get("data", {}).get("name", "").lower() == name.lower()]
    if not hits:
        names = sorted(c.get("data", {}).get("name", "") for c in cols)
        sys.exit("No collection named %r. Collections found: %s" % (name, ", ".join(names) or "none"))
    return hits[0]["key"]


def fetch_items(api, collection=None):
    if collection:
        key = find_collection(api, collection)
        url = "%s/users/0/collections/%s/items/top?format=csljson" % (api, urllib.parse.quote(key))
    else:
        url = "%s/users/0/items/top?format=csljson" % api
    return paged(url)


# ------------------------------------------------------------------ matching helpers
def norm_doi(doi):
    if not doi or not isinstance(doi, str):
        return None
    doi = re.sub(r"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", doi.strip().lower())
    return doi.rstrip(".") or None


def norm_title(t):
    return re.sub(r"[^a-z0-9]+", "", (t or "").lower())


def csl_year(item):
    try:
        return int(item["issued"]["date-parts"][0][0])
    except Exception:
        return None


def parse_names(authors):
    """CSL names from a list of names or from a string like 'Smith, J.; Lee, K.'."""
    if not authors:
        return []
    if isinstance(authors, list):
        parts = [str(a).strip() for a in authors if str(a).strip()]
    else:
        parts = [p.strip() for p in re.split(r";|\band\b|&", str(authors)) if p.strip()]
        if len(parts) == 1 and parts[0].count(",") >= 2:  # "Smith, J., Lee, K."
            bits = [b.strip() for b in parts[0].split(",")]
            parts = [", ".join(bits[i:i + 2]) for i in range(0, len(bits), 2)]
    names = []
    for p in parts:
        p = re.sub(r"\bet al\.?", "", p).strip(" ,.")
        if not p:
            continue
        if "," in p:
            fam, giv = [x.strip() for x in p.split(",", 1)]
        else:
            toks = p.split()
            fam, giv = toks[-1], " ".join(toks[:-1])
        names.append({"family": fam, "given": giv} if giv else {"literal": fam})
    return names


def csl_from_source(sid, src):
    """CSL-JSON for a source that is not in Zotero. Zotero embeds it in the Word file."""
    csl = {"id": "embedded-" + sid, "type": src.get("type") or "article-journal",
           "title": src.get("title")}
    names = parse_names(src.get("authors"))
    if names:
        csl["author"] = names
    if str(src.get("year") or "").isdigit():
        csl["issued"] = {"date-parts": [[int(src["year"])]]}
    for field, csl_field in (("venue", "container-title"), ("volume", "volume"),
                             ("issue", "issue"), ("pages", "page"), ("url", "URL")):
        if src.get(field):
            csl[csl_field] = str(src[field])
    if norm_doi(src.get("doi")):
        csl["DOI"] = src["doi"]
    return csl


def to_ris(src):
    lines = ["TY  - JOUR"]
    for n in parse_names(src.get("authors")):
        lines.append("AU  - " + (n.get("literal") or "%s, %s" % (n["family"], n["given"])))
    for tag, key in (("TI", "title"), ("PY", "year"), ("JO", "venue"), ("VL", "volume"),
                     ("IS", "issue"), ("SP", "pages"), ("DO", "doi"), ("UR", "url")):
        if src.get(key):
            lines.append("%s  - %s" % (tag, src[key]))
    lines.append("ER  - ")
    return "\n".join(lines)


def short_cite(csl):
    names = csl.get("author") or csl.get("editor") or []
    fam = lambda n: n.get("family") or n.get("literal") or "?"
    if not names:
        who = (csl.get("title") or "?")[:30]
    elif len(names) == 1:
        who = fam(names[0])
    elif len(names) == 2:
        who = "%s & %s" % (fam(names[0]), fam(names[1]))
    else:
        who = "%s et al." % fam(names[0])
    return "%s (%s)" % (who, csl_year(csl) or "n.d.")


def md_cell(s):
    return str(s or "").replace("|", "/").replace("\n", " ").strip()


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--collection", help="only items in this Zotero collection (by name)")
    ap.add_argument("--sources", help="JSON list of your own sources: id, title, authors, year, doi, ...")
    ap.add_argument("--out", default="citations", help="output folder (default: citations)")
    ap.add_argument("--api", default=API)
    ap.add_argument("--offline", action="store_true",
                    help="do not contact Zotero; embed every --sources entry (for testing)")
    args = ap.parse_args()

    lib = []
    if not args.offline:
        try:
            lib = fetch_items(args.api, args.collection)
        except SystemExit:
            raise
        except Exception as exc:
            sys.exit("Cannot reach Zotero's local API at %s (%s).\nIs Zotero running, with Settings -> "
                     "Advanced -> 'Allow other applications on this computer to communicate with "
                     "Zotero' ticked? Use --offline to test without Zotero." % (args.api, exc))
    elif not args.sources:
        sys.exit("--offline needs --sources.")

    out = {}
    by_doi, by_title = {}, {}
    for it in lib:
        uri = it["id"]
        key = uri.rsplit("/", 1)[-1]
        out[key] = {"key": key, "uri": uri, "match": "zotero", "csl": it}
        if norm_doi(it.get("DOI")):
            by_doi[norm_doi(it["DOI"])] = key
        by_title.setdefault(norm_title(it.get("title")), []).append(key)

    unmatched, ris = [], []
    sources = []
    if args.sources:
        with open(args.sources, encoding="utf-8") as f:
            sources = json.load(f)
        if isinstance(sources, dict):  # {"P017": {...}} is accepted too
            sources = [dict(v, id=k) for k, v in sources.items()]
    for src in sources:
        sid = src["id"]
        if sid in out:
            sys.exit("Source id %r clashes with a Zotero item key; rename it." % sid)
        key, how = by_doi.get(norm_doi(src.get("doi"))), "doi"
        if key is None:
            cands = by_title.get(norm_title(src.get("title")), [])
            if len(cands) > 1 and str(src.get("year") or "").isdigit():
                cands = [c for c in cands if csl_year(out[c]["csl"]) == int(src["year"])] or cands
            key, how = (cands[0], "title") if cands else (None, None)
        if key is not None:
            out[sid] = dict(out[key], match=how)
        else:
            out[sid] = {"key": None, "uri": None, "match": "embedded",
                        "csl": csl_from_source(sid, src)}
            if src.get("pdf"):
                out[sid]["pdf"] = src["pdf"]
            unmatched.append("%s | %s | %s" % (sid, src.get("year"), (src.get("title") or "")[:120]))
            ris.append(to_ris(src))

    os.makedirs(args.out, exist_ok=True)
    with open(os.path.join(args.out, "zotero_map.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, indent=1, ensure_ascii=False)
    with open(os.path.join(args.out, "unmatched.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(unmatched) + ("\n" if unmatched else ""))
    ris_path = os.path.join(args.out, "unmatched.ris")
    if ris:
        with open(ris_path, "w", encoding="utf-8") as f:
            f.write("\n\n".join(ris) + "\n")
    elif os.path.exists(ris_path):
        os.remove(ris_path)

    # The list the writing agent reads to choose what to cite.
    own = {s["id"] for s in sources}
    rows = ["# Citable works", "",
            "Cite with `{@ID}`, several with `{@ID1; @ID2}`, narrative with `{@ID|n}`,",
            "a page with `{@ID, p. 12}`. Use only the IDs below.", "",
            "| ID | Zotero key | Short cite | Title | DOI |", "|---|---|---|---|---|"]
    listed = [i for i in out if i in own] + [k for k in out if k not in own and
                                             not any(out[s].get("key") == k for s in own)]
    for i in listed:
        e = out[i]
        rows.append("| %s | %s | %s | %s | %s |" % (
            i, e.get("key") or "(embedded)", md_cell(short_cite(e["csl"])),
            md_cell(e["csl"].get("title"))[:140], md_cell(e["csl"].get("DOI"))))
    with open(os.path.join(args.out, "library.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(rows) + "\n")

    print("Zotero items read: %d%s" % (len(lib), " (collection %r)" % args.collection if args.collection else ""))
    if sources:
        n = len(sources)
        cnt = lambda m: sum(1 for s in sources if out[s["id"]]["match"] == m)
        print("Your sources: %d -> %d by DOI, %d by title, %d not in Zotero (embedded)"
              % (n, cnt("doi"), cnt("title"), cnt("embedded")))
        for u in unmatched:
            print("  not in Zotero: " + u)
        if ris:
            print("  import %s into Zotero and re-run to link them (optional)" % ris_path)
    print("Wrote %s/zotero_map.json and %s/library.md (%d citable ids)" % (args.out, args.out, len(listed)))


if __name__ == "__main__":
    main()
