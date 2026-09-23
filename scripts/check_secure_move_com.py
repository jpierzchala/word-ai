"""Synthetic Windows Word COM check of the section-move package from section-move.js.

Word's Range.WordOpenXML / InsertXML use the same Flat OPC round trip as Office.js
getOoxml / insertOoxml. This validates the package transform (sentinel paragraph,
temporary bookmark names, stripped section properties) and the expected-sequence
check against real Word. It is NOT an end-to-end test of the Office.js add-in.
Requires Word, pywin32, Node and `npm ci --ignore-scripts` in secure-addin.
No user document is opened, changed or overwritten.
"""
import json
import os
from pathlib import Path
import struct
import subprocess
import tempfile
import zlib
import win32com.client

REPO = Path(__file__).resolve().parents[1] / "secure-addin"
NODE = """
import fs from 'node:fs';
import {DOMParser, XMLSerializer} from '@xmldom/xmldom';
import {inspectSection, sectionDigest, preparePackage, findSection, resolveDestination, expectedSequence, compareSequence} from './section-move.js';
const input = JSON.parse(fs.readFileSync(process.env.INPUT_JSON, 'utf8'));
let out;
if (input.mode === 'prepare') {
  const xml = fs.readFileSync(input.xml, 'utf8');
  const {blockers, bookmarks, counts, ends_with_table} = inspectSection(xml, DOMParser);
  const prepared = preparePackage(xml, DOMParser, XMLSerializer);
  fs.writeFileSync(input.xml + '.prepared', prepared.xml);
  // Two separate Word exports of the unchanged range must fingerprint the same.
  out = {blockers, bookmarks, counts, ends_with_table, renames: prepared.bookmarks,
    digests: [await sectionDigest(xml, DOMParser), await sectionDigest(fs.readFileSync(input.xml + '.again', 'utf8'), DOMParser)]};
} else if (input.mode === 'plan') {
  const section = findSection(input.rows, input.heading, input.finalEmpty);
  out = {section, ...resolveDestination(input.rows, section, input.target, input.position)};
} else {
  const expected = expectedSequence(input.rows, input.section, input.destination);
  out = {...compareSequence(input.after, expected), seam: expected.sentinelIndex};
}
process.stdout.write(JSON.stringify(out));
"""


def node(payload, temp):
    path = Path(temp) / "input.json"
    path.write_text(json.dumps(payload), encoding="utf8")
    result = subprocess.run(["node", "--input-type=module", "-e", NODE], cwd=REPO, capture_output=True,
                            text=True, check=True, env={**os.environ, "INPUT_JSON": str(path)})
    return json.loads(result.stdout)


def png():
    def chunk(name, data):
        return struct.pack(">I", len(data)) + name + data + struct.pack(">I", zlib.crc32(name + data))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 2, 1, 8, 2, 0, 0, 0)) +
            chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00\x00\xff\x00")) + chunk(b"IEND", b""))


LINES = [("Intro", -1), ("A", -2), ("a1 ", -1), ("B", -2), ("B.1", -3), ("Lead-in: text", -1), ("see ", -1),
         ("PICTURE", -1), ("Figure ", -1), ("bookmarked", -1), ("has footnote", -1), ("TABLE", -1),
         ("C", -2), ("c1", -1), ("", -1)]


# A section ending with a genuine empty paragraph next to Word's seam artifacts.
PLAIN = [("A", -2), ("a1", -1), ("B", -2), ("b1", -1), ("", -3), ("C", -2), ("c1", -1), ("", -1)]


def build(word, temp, lines=None):
    doc = word.Documents.Add()
    lines = lines or LINES
    doc.Content.Text = "\r".join(text for text, _ in lines)
    for i, (_, style) in enumerate(lines, 1):
        doc.Paragraphs(i).Style = doc.Styles(style)
    if lines is not LINES:
        return doc
    find = lambda text: next(doc.Paragraphs(i).Range for i in range(1, doc.Paragraphs.Count + 1)
                             if doc.Paragraphs(i).Range.Text.rstrip("\r") == text)
    end = lambda r: doc.Range(r.End - 1, r.End - 1)
    lead = find("Lead-in: text"); doc.Range(lead.Start, lead.Start + 8).Font.Bold = True
    marked = find("bookmarked"); doc.Bookmarks.Add("bm_test", doc.Range(marked.Start, marked.End - 1))
    heading = find("B.1"); doc.Bookmarks.Add("_Ref900001", doc.Range(heading.Start, heading.End - 1))
    target = find("A"); doc.Bookmarks.Add("_Toc900002", doc.Range(target.Start, target.End - 1))  # a TOC anchor on a move target
    doc.Fields.Add(end(find("see ")), -1, r"REF _Ref900001 \h", False)
    doc.Fields.Add(end(find("a1 ")), -1, r"REF bm_test \h", False)
    doc.Fields.Add(end(find("Figure ")), -1, r"SEQ Figure \* ARABIC", False)
    doc.Footnotes.Add(Range=end(find("has footnote")), Text="note text")
    picture = find("PICTURE"); image = Path(temp) / "check.png"; image.write_bytes(png())
    target = doc.Range(picture.Start, picture.End - 1); target.Text = ""
    doc.InlineShapes.AddPicture(str(image), False, True, target)
    table = doc.Tables.Add(find("TABLE"), 2, 2)  # the section B ends with this table
    for (row, column), value in {(1, 1): "x1", (1, 2): "x2", (2, 1): "y1", (2, 2): "y2"}.items():
        table.Cell(row, column).Range.Text = value
    doc.Fields.Update()
    return doc


def rows(doc):
    out = []
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        level = p.Range.Tables(1).NestingLevel if p.Range.Information(12) else 0
        out.append({"uniqueLocalId": f"p{i}", "text": p.Range.Text.replace("\r", "").replace("\x07", ""),
                    "style": p.Style.NameLocal, "styleBuiltIn": "Other", "outlineLevel": p.OutlineLevel,
                    "tableNestingLevel": level})
    return out


def state(doc):
    doc.Fields.Update()
    doc.Bookmarks.ShowHidden = True
    return {"tables": doc.Tables.Count, "pictures": doc.InlineShapes.Count, "footnotes": doc.Footnotes.Count,
            "sections": doc.Sections.Count, "bookmarks": sorted(b.Name for b in doc.Bookmarks if b.Name != "_GoBack"),
            "fields": sorted(f.Result.Text.strip() for f in doc.Fields)}


def move(word, temp, heading, target, position, lines=None):
    doc = build(word, temp, lines)
    try:
        before, facts = rows(doc), state(doc)
        pid = lambda text: next(r["uniqueLocalId"] for r in before if r["text"] == text and r["tableNestingLevel"] == 0)
        plan = node({"mode": "plan", "rows": before, "heading": pid(heading), "target": pid(target),
                     "position": position, "finalEmpty": before[-1]["text"] == ""}, temp)
        section = plan["section"]
        paragraph = lambda i: doc.Paragraphs(i + 1).Range
        source = doc.Range(paragraph(section["start"]).Start, paragraph(section["end"]).Start)
        xml_path = Path(temp) / "section.xml"
        xml_path.write_text(source.WordOpenXML, encoding="utf8")
        Path(str(xml_path) + ".again").write_text(source.WordOpenXML, encoding="utf8")
        prepared = node({"mode": "prepare", "xml": str(xml_path)}, temp)
        assert not prepared["blockers"], prepared["blockers"]
        assert prepared["digests"][0] == prepared["digests"][1], "section digest is not deterministic"
        destination = paragraph(plan["destination"])
        doc.Bookmarks.ShowHidden = True
        # Bookmarks starting at the insertion point grow over the copy (live-ops.js restores them).
        landing = {b.Name: b.Range.Text for b in doc.Bookmarks if b.Range.Start == destination.Start and b.Name != "_GoBack"}
        doc.Range(destination.Start, destination.Start).InsertXML(Path(str(xml_path) + ".prepared").read_text(encoding="utf8"))
        source.Delete()
        compare = lambda: node({"mode": "compare", "rows": before, "section": section,
                                "destination": plan["destination"], "after": rows(doc)}, temp)
        result, seam_removed = compare(), 0
        if result["outcome"] == "seam":  # same cleanup as live-ops.js
            for i in reversed(range(result["extras"])):
                doc.Paragraphs(result["seam"] + i + 1).Range.Delete()
            seam_removed, result = result["extras"], compare()
        outcome, after = result["outcome"], rows(doc)
        doc.Bookmarks.ShowHidden = True
        for rename in prepared["renames"]:
            mark = doc.Bookmarks(rename["temp"]); where = mark.Range; mark.Delete(); doc.Bookmarks.Add(rename["name"], where)
        grown = [name for name, text in landing.items() if doc.Bookmarks(name).Range.Text != text]
        start = doc.Paragraphs(result.get("seam", 0) + 1).Range.Start if landing else None
        for name in grown:
            doc.Bookmarks.Add(name, doc.Range(start, doc.Bookmarks(name).Range.End))
        assert all(doc.Bookmarks(name).Range.Text == text for name, text in landing.items()), landing
        facts_after = state(doc)
        assert outcome == "exact", (heading, target, position, outcome, [r["text"] for r in after])
        assert facts_after == facts, (facts, facts_after)
        assert not any("Error!" in field for field in facts_after["fields"]), facts_after["fields"]
        return {**prepared, "seam_removed": seam_removed, "grown": grown}
    finally:
        doc.Close(0)


word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
try:
    with tempfile.TemporaryDirectory() as temp:
        cases = [("B", "A", "before"), ("B", "c1", "after"), ("A", "C", "before"), ("C", "A", "before"),
                 ("B", "A", "before", PLAIN), ("B", "c1", "after", PLAIN)]
        for case in cases:
            prepared = move(word, temp, *case)
            case = case[:3] + (("empty-ending",) if len(case) > 3 else ())
            print("PASS", case, "seam paragraphs removed:", prepared["seam_removed"], "restored destination bookmarks:", prepared["grown"],
                  "ends_with_table" if prepared["ends_with_table"] else "", prepared["counts"])
    print("PASS: Word COM kept order, styles, tables, pictures, footnotes, sections, bookmarks and field results "
          "for sections moved up, down, to the end and from the end (COM, not Office.js).")
finally:
    word.Quit()
