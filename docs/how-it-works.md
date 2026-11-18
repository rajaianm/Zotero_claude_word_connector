# How it works

## 1. Reading the library
Zotero 7 has a local, read-only API at `http://127.0.0.1:23119/api/` that uses the same
URLs as the Zotero web API, with `users/0` meaning "the local user". `zotero_link.py` asks
for top-level items as CSL-JSON:

```
GET /api/users/0/items/top?format=csljson&limit=100&start=0
GET /api/users/0/collections                                  (to find a collection by name)
GET /api/users/0/collections/<KEY>/items/top?format=csljson
```

In this format, each item's `id` is its URI, for example
`http://zotero.org/users/local/AbC123xy/items/8XK2M4QZ`. The last part is the item key.

## 2. A Zotero citation inside Word
In the Word plugin's default "Fields" mode, each citation is a complex field:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> ADDIN ZOTERO_ITEM CSL_CITATION {JSON} </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>(Example et al., 2021)</w:t></w:r>          <- text shown before refresh
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

The JSON looks like this:

```json
{
  "citationID": "28eb7ca8",
  "properties": {"formattedCitation": "(Example et al., 2021)",
                 "plainCitation": "(Example et al., 2021)", "noteIndex": 0},
  "citationItems": [
    {"id": "http://zotero.org/users/local/AbC123xy/items/8XK2M4QZ",
     "uris": ["http://zotero.org/users/local/AbC123xy/items/8XK2M4QZ"],
     "itemData": { "...full CSL-JSON of the item..." },
     "locator": "7", "label": "page",
     "suppress-author": true}
  ],
  "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
}
```

On *Refresh*, the plugin uses the URIs to find each item in your library. If an item is
found, its current library data is used. If not, the plugin falls back to the embedded
`itemData`. Either way it re-renders the citation in the document's style.

`locator` and `label` come from `{@KEY, p. 7}`. `suppress-author` comes from `{@KEY|n}`:
the citation shows only "(2023)", and the script writes the author names as normal text
before it ("Placeholder and Mock (2023)"), which is how narrative citations work in Zotero.

## 3. Works that are not in your library
The plugin requires a URI on every item. An item with no `uris` makes it try to read the
`id` as a numeric database id. That fails, and you get *"Zotero experienced an error
updating your document"*.

`zotero_docx.py` therefore gives every unmatched work a well-formed URI in your library's
namespace, using a key built from a hash of the ID that matches no real item:

```
http://zotero.org/users/local/<your library id>/items/9F9EITBK
```

The lookup finds nothing, so the plugin uses the embedded `itemData`, which is what we
want. The same ID always gets the same key, so repeated citations of a work are recognised
as one work in the bibliography.

## 4. The bibliography
A paragraph containing only `{BIBLIOGRAPHY}` becomes:

```
ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY
```

with a pre-rendered APA list as the shown text. Zotero rebuilds it on refresh.

## 5. Document preferences
The plugin stores the style and other settings as custom document properties named
`ZOTERO_PREF_1`, `ZOTERO_PREF_2`, … in `docProps/custom.xml`. Together they hold XML split
into 240-character chunks:

```xml
<data data-version="3" zotero-version="7.0">
  <session id="…"/>
  <style id="http://www.zotero.org/styles/apa" locale="en-US" hasBibliography="1"
         bibliographyStyleHasBeenSet="1"/>
  <prefs><pref name="fieldType" value="Field"/>
         <pref name="automaticJournalAbbreviations" value="true"/></prefs>
</data>
```

`--style` and `--locale` set these values. Without them, the plugin would ask for a style
at the first refresh.

## 6. Why placeholders?
- **The agent never handles formatted references,** so it cannot get a year, an initial
  or an "et al." wrong, and cannot invent a work.
- **Every citation is traceable** to an ID in `zotero_map.json`, so a reviewer agent or a
  script can check it.
- **A Markdown draft stays readable and easy to diff,** and `preview.py` turns it into
  clickable author–year text.
