# Troubleshooting

**`Cannot reach Zotero's local API`**
Start Zotero. In *Settings* → *Advanced*, tick "Allow other applications on this computer
to communicate with Zotero". Test it with:
`curl "http://127.0.0.1:23119/api/users/0/items/top?limit=1"`.
To try the scripts without Zotero, use `--offline --sources …`.

**`No collection named '…'`**
The name must match exactly (case does not matter). The error lists the collections it
found. Sub-collections are separate collections: give the sub-collection's own name.

**`NOT CONVERTED: {@XYZ} (unknown id)`**
The ID is not in `citations/zotero_map.json`. Either it is a typo, or the item was added
to Zotero after you last ran `zotero_link.py`, or it is outside the `--collection` you
used. Fix it and run again. The script leaves unknown placeholders in the text so they are
easy to find, and it never guesses.

**`NOT CONVERTED: … (malformed)`**
Each item must start with `@`, and items are separated by `;`. For example:
`{@A; @B, p. 3}`.

**"Zotero experienced an error updating your document" in Word**
This happens with files made by other tools that leave out item URIs. Files from this
`zotero_docx.py` always include them (see [how-it-works.md](how-it-works.md#3-works-that-are-not-in-your-library)).
If you still see it, re-create the file with the current scripts and send an issue with
the field code. To see it, right-click a citation → *Toggle Field Codes*.

**Word asks which citation style to use**
The document preferences are missing, for example because the file was re-saved by
another program. Choose a style. Everything else still works.

**Some sources are "not in Zotero" although they are**
They are matched by DOI, then by exact title after removing punctuation and case. Check
the DOI and title in `sources.json` against Zotero, or add the DOI to the Zotero item. Then
run `zotero_link.py` again. Unmatched works still cite correctly. They are just not linked
to your library item.

**The same work appears twice in the bibliography**
It was cited once as a Zotero key and once as an unmatched own ID. Fix the match (see
above) so both point to the same Zotero item.

**Citations show raw `{@…}` after refresh**
That text was not converted. Run `zotero_docx.py` on the file and read its output.
Placeholders in footnotes, headers and text boxes are not supported.
