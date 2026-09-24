"""Synthetic Windows Word COM check of list_restart (secure-addin/document-edit.js).

1. Reproduces the 2026-09-23 failure: Paragraph.SeparateList is refused ("This
   command is not available.") exactly on the first paragraph of a numbering
   instance (w:num); in the reported layout that is the first item of 9.1, a
   second instance of the 7.1 list definition. The add-in no longer uses it.
2. Runs sequences of restarts the way the add-in does (restartPackage, insert at
   the end of the item's content, remove the empty paragraph that keeps the
   item's former mark, restartOutcome) and requires the numbering the preview
   promises (the item 1, the run after it shifted, all else unchanged), no change
   to other paragraphs, styles, bookmarks, comments, fields or list definitions.
   In single-level lists that is exactly Word's own "Restart at 1"
   (NumberingRestart). In multi-level lists with several instances Word's own
   command may copy the list definition, so later blocks keep counting the old
   one; that comparison is reported, not required.
Range.WordOpenXML / InsertXML use the same Flat OPC import as Office.js
getOoxml / insertOoxml, but the live add-in showed differences (see
docs/SECURE-LIVE.md); this is NOT an end-to-end test of the add-in.
Requires Word, pywin32, Node and `npm ci --ignore-scripts` in secure-addin.
Uses a separate invisible Word instance; no user document is opened or changed.
"""
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import win32com.client

ADDIN = Path(__file__).resolve().parents[1] / "secure-addin"
NODE = """
import fs from 'node:fs';
import {DOMParser, XMLSerializer} from '@xmldom/xmldom';
import {restartPackage, restartOutcome} from './document-edit.js';
const input = JSON.parse(fs.readFileSync(process.env.INPUT_JSON, 'utf8'));
const read = name => fs.readFileSync(input[name], 'utf8');
let out;
try {
  if (input.mode === 'package') {
    const r = restartPackage(read('xml'), 0, DOMParser, XMLSerializer);
    fs.writeFileSync(input.xml + '.pkg', r.xml);
    out = {ok: true, hidden: r.hidden};
  } else out = {ok: true, problems: restartOutcome(read('before'), read('after'), 0, DOMParser, null, input.hidden)};
} catch (error) { out = {ok: false, error: error.message}; }
process.stdout.write(JSON.stringify(out));
"""
NORMAL, H1, H2, LIST_NUMBER = -1, -2, -3, -50


def layout():
    """The reported document: 7.1 with 7 items, a long gap with headings and a table, 9.1 with 4 and 9.2 with 3 items."""
    rows = [(H1, "7 Requirements"), (H2, "7.1 Functional")] + [(LIST_NUMBER, f"a{i}") for i in range(1, 8)]
    rows += [(H2, "7.2 Other"), (NORMAL, "gap A"), (H1, "8 Design"), (H2, "8.1 Architecture"), (NORMAL, "TABLE")]
    rows += [(NORMAL, f"filler {i}") for i in range(20)]
    rows += [(H1, "9 Delivery"), (H2, "9.1 Plan")] + [(LIST_NUMBER, f"b{i}") for i in range(1, 5)]
    rows += [(H2, "9.2 Risks")] + [(LIST_NUMBER, f"c{i}") for i in range(1, 4)]
    rows += [(H1, "10 End"), (NORMAL, "end"), (NORMAL, "")]
    return rows


def find(doc, text):
    for i in range(1, doc.Paragraphs.Count + 1):
        if doc.Paragraphs(i).Range.Text.rstrip("\r\x07").startswith(text):
            return doc.Paragraphs(i)
    raise KeyError(text)


def build(word, path, mode="style"):
    """style: List Number (single-level); direct: numbering button list; multilevel: 9-level outline list."""
    doc = word.Documents.Add()
    rows = layout()
    doc.Content.Text = "\r".join(text for _, text in rows)
    for i, (style, _) in enumerate(rows, 1):
        doc.Paragraphs(i).Style = doc.Styles(NORMAL if mode != "style" and style == LIST_NUMBER else style)
    if mode != "style":
        template = doc.ListTemplates.Add(True) if mode == "multilevel" else None
        for first, last in [("a1", "a7"), ("b1", "b4"), ("c1", "c3")]:
            block = doc.Range(find(doc, first).Range.Start, find(doc, last).Range.End)
            if template is None:
                block.ListFormat.ApplyNumberDefault()
                template = find(doc, first).Range.ListFormat.ListTemplate
            else:
                block.ListFormat.ApplyListTemplate(template, first != "a1")
    table = doc.Tables.Add(find(doc, "TABLE").Range, 2, 2)
    table.Cell(1, 1).Range.Text = "Component"
    table.Borders.Enable = True
    doc.SaveAs2(str(path), 16)
    doc.Close(0)


def renumber(src, dst, plan):
    """A second w:num of a1's list definition for paragraphs in plan (exact text -> "new"; "old" = a1's instance),
    as assembled documents have."""
    with zipfile.ZipFile(src) as z:
        files = {n: z.read(n) for n in z.namelist()}
    document, numbering = files["word/document.xml"].decode("utf8"), files["word/numbering.xml"].decode("utf8")
    first = next(p for p in re.findall(r"<w:p[ >].*?</w:p>", document, re.S) if re.search(r"<w:t(?: [^>]*)?>a1</w:t>", p))
    style = re.search(r'<w:pStyle w:val="([^"]+)"', first).group(1)
    styles = files["word/styles.xml"].decode("utf8")
    used = (re.findall(r'<w:numId w:val="(\d+)"', first) or
            re.findall(r'<w:numId w:val="(\d+)"', re.search(rf'<w:style [^>]*w:styleId="{style}".*?</w:style>', styles, re.S).group(0)))[0]
    abstract = re.search(rf'<w:num w:numId="{used}"[^>]*><w:abstractNumId w:val="(\d+)"', numbering).group(1)
    fresh = str(max(map(int, re.findall(r'<w:num w:numId="(\d+)"', numbering))) + 1)
    ids = {"new": fresh, "old": used}
    def fix(match):
        p = match.group(0)
        text = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p))
        if text in plan:
            p = re.sub(r"<w:numPr>.*?</w:numPr>", "", p)
            p = re.sub(r"(<w:pStyle [^>]*/>)", r'\1<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%s"/></w:numPr>' % ids[plan[text]], p, count=1)
        return p
    files["word/document.xml"] = re.sub(r"<w:p[ >].*?</w:p>", fix, document, flags=re.S).encode("utf8")
    files["word/numbering.xml"] = numbering.replace("</w:numbering>", f'<w:num w:numId="{fresh}"><w:abstractNumId w:val="{abstract}"/></w:num></w:numbering>').encode("utf8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)


def numbers(doc):
    return {p.Range.Text.rstrip("\r")[:2]: p.Range.ListFormat.ListValue
            for p in (doc.Paragraphs(i) for i in range(1, doc.Paragraphs.Count + 1)) if p.Range.ListFormat.ListType}


def structure(doc):
    """Per paragraph: text, style, level and its numbering instance by content; plus the list definitions."""
    xml = doc.Content.WordOpenXML
    numbering = re.search(r"<w:numbering .*?</w:numbering>", xml, re.S).group(0)
    nums = {n: (a, re.sub(r"\s+w16cid:durableId=\"\d+\"", "", body))
            for n, a, body in re.findall(r'<w:num w:numId="(\d+)"[^>]*><w:abstractNumId w:val="(\d+)"/>(.*?)</w:num>', numbering, re.S)}
    styles = re.search(r"<w:styles .*?</w:styles>", xml, re.S).group(0)
    style_num = dict(re.findall(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>(?:(?!</w:style>).)*?<w:numId w:val="(\d+)"/>', styles, re.S))
    body = re.search(r"<w:body>(.*)</w:body>", xml, re.S).group(1)
    rows = []
    for p in re.findall(r"<w:p[ >].*?</w:p>", body, re.S):
        text = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p))
        style = (re.findall(r'<w:pStyle w:val="([^"]+)"', p) or ["Normal"])[0]
        num = (re.findall(r'<w:numId w:val="(\d+)"', p) or [style_num.get(style)])[0]
        rows.append((text, style, re.findall(r'<w:ilvl w:val="(\d)"', p), nums.get(num)))
    abstracts = sorted(re.findall(r"<w:nsid w:val=\"([0-9A-F]+)\"", numbering))
    return rows, abstracts


def node(temp, payload):
    path = Path(temp) / "input.json"
    path.write_text(json.dumps(payload), encoding="utf8")
    result = subprocess.run(["node", "--input-type=module", "-e", NODE], cwd=ADDIN, capture_output=True, text=True,
                            check=True, env={**os.environ, "INPUT_JSON": str(path)})
    return json.loads(result.stdout)


def add_in_restart(doc, temp, target):
    """The add-in's write: the restart package at the end of the item's content, then the empty paragraph
    that keeps the item's former mark is removed."""
    item = find(doc, target)
    whole = Path(temp) / "whole.xml"
    whole.write_text(item.Range.WordOpenXML, encoding="utf8")
    planned = node(temp, {"mode": "package", "xml": str(whole)})
    if not planned["ok"]:
        return "refused: " + planned["error"], []
    count = doc.Paragraphs.Count
    end = item.Range.End - 1
    doc.Range(end, end).InsertXML(Path(str(whole) + ".pkg").read_text(encoding="utf8"))
    item = find(doc, target)
    after_item = item.Next()
    if doc.Paragraphs.Count != count + 1 or after_item is None or after_item.Range.Text != "\r":
        return "start_override", ["unexpected split"]
    after_item.Range.Delete()
    after = Path(temp) / "after.xml"
    after.write_text(find(doc, target).Range.WordOpenXML, encoding="utf8")
    return "start_override", node(temp, {"mode": "outcome", "before": str(whole), "after": str(after), "hidden": planned["hidden"]})["problems"]


def promised(before, target):
    """The preview's model: the item becomes 1, the run of consecutive numbers after it shifts, the rest stays."""
    keys = list(before)
    index = keys.index(target)
    shift, expected = before[target] - 1, dict(before)
    for j in range(index, len(keys)):
        if j > index and before[keys[j]] != before[keys[j - 1]] + 1:
            break
        expected[keys[j]] = before[keys[j]] - shift
    return expected


def word_restart(word, doc, target):
    doc.Activate()
    point = find(doc, target).Range
    point.Collapse(1)
    point.Select()
    word.CommandBars.ExecuteMso("NumberingRestart")


def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    report, failures = {"separate_list": {}, "sequences": {}}, []
    try:
        with tempfile.TemporaryDirectory() as temp:
            base = {mode: Path(temp) / f"{mode}.docx" for mode in ("style", "direct", "multilevel")}
            for mode, path in base.items():
                build(word, path, mode)
            second = {f"b{i}": "new" for i in range(1, 5)} | {f"c{i}": "new" for i in range(1, 4)}
            variants = dict(base)
            for mode in ("style", "multilevel"):
                variants[f"{mode}-continued"] = Path(temp) / f"{mode}-continued.docx"  # the reported layout
                renumber(base[mode], variants[f"{mode}-continued"], second)
            variants["interleaved"] = Path(temp) / "interleaved.docx"
            renumber(base["style"], variants["interleaved"], second | {"b3": "old", "b4": "old"})
            variants["rich"] = Path(temp) / "rich.docx"
            renumber(base["style"], variants["rich"], second)
            doc = word.Documents.Open(str(variants["rich"]))
            item = find(doc, "b1")
            doc.Range(item.Range.End - 1, item.Range.End - 1).InsertAfter(" lead: text ")
            doc.Range(item.Range.Start, item.Range.Start + 2).Font.Bold = True
            doc.Bookmarks.Add("bm_item", doc.Range(item.Range.Start, item.Range.Start + 2))
            doc.Comments.Add(doc.Range(item.Range.Start + 3, item.Range.Start + 7), "review note")
            doc.Fields.Add(doc.Range(item.Range.End - 1, item.Range.End - 1), 33)  # PAGE
            doc.Save()
            doc.Close(0)
            refused = {("style-continued", "b1"), ("multilevel-continued", "b1"), ("interleaved", "b1"), ("rich", "b1")}
            for name, path in variants.items():
                for target in ("b1", "b2", "c1"):
                    doc = word.Documents.Open(str(path), False, True)
                    try:
                        try:
                            find(doc, target).SeparateList()
                            outcome = "ok"
                        except Exception as error:  # COM error with Word's message
                            outcome = str(error.excepinfo[2] if getattr(error, "excepinfo", None) else error)
                    finally:
                        doc.Close(0)
                    report["separate_list"][f"{name}:{target}"] = outcome
                    if (outcome != "ok") != ((name, target) in refused):
                        failures.append(f"SeparateList {name}:{target} -> {outcome}")
            sequences = [(name, [t]) for name in variants for t in ("b1", "b3")]
            sequences += [(name, seq) for name in ("style-continued", "multilevel-continued")
                          for seq in (["b1", "c1"], ["c1", "b1"], ["b1", "c1", "b3"], ["a3", "a5"], ["a4", "a2"])]
            for name, seq in sequences:
                label = f"{name}:{','.join(seq)}"
                # Word opens one file once: compare two copies.
                copies = [Path(temp) / f"{side}.docx" for side in ("ours", "words")]
                for copy in copies:
                    shutil.copyfile(variants[name], copy)
                ours, words = (word.Documents.Open(str(copy), False, True) for copy in copies)
                try:
                    steps = []
                    for target in seq:
                        rows_before, abstracts_before = structure(ours)
                        numbers_before = numbers(ours)
                        counts = (ours.Paragraphs.Count, ours.Bookmarks.Count, ours.Comments.Count, ours.Fields.Count,
                                  sorted(s.NameLocal for s in ours.Styles if s.InUse))
                        method, problems = add_in_restart(ours, temp, target)
                        steps.append(method)
                        if method.startswith("refused"):
                            break  # the preview would refuse; nothing written
                        word_restart(word, words, target)
                        rows_after, abstracts_after = structure(ours)
                        index = [r[0][:2] for r in rows_before].index(target)
                        checks = {"numbers": numbers(ours) == promised(numbers_before, target), "outcome": problems == [],
                                  "objects": counts == (ours.Paragraphs.Count, ours.Bookmarks.Count, ours.Comments.Count, ours.Fields.Count,
                                                        sorted(s.NameLocal for s in ours.Styles if s.InUse)),
                                  "texts": [r[0] for r in rows_after] == [r[0] for r in rows_before]}
                        if method == "start_override":
                            checks["definitions"] = abstracts_after == abstracts_before
                            checks["other_paragraphs"] = [r for i, r in enumerate(rows_after) if i != index] == [r for i, r in enumerate(rows_before) if i != index]
                        failures += [f"{label} {target} ({method}): {k}" for k, ok in checks.items() if not ok]
                        if name.startswith("style") and numbers(ours) != numbers(words):
                            failures.append(f"{label} {target}: differs from Word's Restart at 1 in a single-level list")
                    if any(s.startswith("refused") for s in steps):
                        failures.append(f"{label}: unexpected methods {steps}")
                    report["sequences"][label] = {"methods": steps, "numbers": list(numbers(ours).values()), "word": list(numbers(words).values()),
                                                  "same_as_word": numbers(ours) == numbers(words)}
                finally:
                    ours.Close(0)
                    words.Close(0)
    finally:
        word.Quit()
    print(json.dumps({"report": report, "failures": failures}, indent=1, ensure_ascii=False))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
