"""Create the synthetic document for the live Office.js acceptance of list_restart.

Reproduces the 2026-09-23 layout: one List Number list over 7.1 (1-7), a long gap
with headings and a table, 9.1 (8-11) and 9.2 (12-14). As in assembled documents,
9.1 starts its own numbering instance (w:num) of the same list definition, so
Paragraph.separateList is refused on its first item. 9.3 holds an independent
list (other lists must not change); 9.4 continues the list (15-16) and its first
item ends a Word section, which the preview must refuse.
Uses a separate invisible Word instance; never opens or changes user documents.
Usage: python make_list_restart_live_test_doc.py <output.docx>
"""
from pathlib import Path
import re
import sys
import tempfile
import zipfile
import win32com.client

NORMAL, TITLE, H1, H2, LIST_NUMBER = -1, -63, -2, -3, -50
ROWS = ([(TITLE, "List restart live test"), (H1, "7 Requirements"), (H2, "7.1 Functional")] +
        [(LIST_NUMBER, f"Requirement {i}") for i in range(1, 8)] +
        [(H2, "7.2 Notes"), (NORMAL, "Notes paragraph."), (H1, "8 Design"), (H2, "8.1 Architecture"),
         (NORMAL, "The components table follows."), (NORMAL, "TABLE")] +
        [(NORMAL, f"Design paragraph {i}.") for i in range(1, 11)] +
        [(H1, "9 Delivery"), (H2, "9.1 Plan")] + [(LIST_NUMBER, f"Plan step {c}") for c in "ABCD"] +
        [(H2, "9.2 Risks")] + [(LIST_NUMBER, f"Risk {c}") for c in "ABC"] +
        [(H2, "9.3 Independent list")] + [(NORMAL, f"Other {w}") for w in ("one", "two", "three")] +
        [(H2, "9.4 Section break case"), (LIST_NUMBER, "Closing item X"), (LIST_NUMBER, "Closing item Y"),
         (H1, "10 End"), (NORMAL, "End of document."), (NORMAL, "")])
CONTINUED = {f"Plan step {c}" for c in "ABCD"} | {f"Risk {c}" for c in "ABC"} | {"Closing item X", "Closing item Y"}
SECTION = ('<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
           'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def find(doc, text):
    return next(doc.Paragraphs(i) for i in range(1, doc.Paragraphs.Count + 1) if doc.Paragraphs(i).Range.Text.rstrip("\r") == text)


def assemble(src, dst):
    """9.1, 9.2 and 9.4 items get a second w:num of the List Number definition; X ends a section."""
    with zipfile.ZipFile(src) as z:
        files = {n: z.read(n) for n in z.namelist()}
    numbering = files["word/numbering.xml"].decode("utf8")
    style_num = re.search(r'<w:style [^>]*w:styleId="ListNumber".*?<w:numId w:val="(\d+)"/>', files["word/styles.xml"].decode("utf8"), re.S).group(1)
    abstract = re.search(rf'<w:num w:numId="{style_num}"[^>]*><w:abstractNumId w:val="(\d+)"/>', numbering).group(1)
    fresh = max(map(int, re.findall(r'<w:num w:numId="(\d+)"', numbering))) + 1
    files["word/numbering.xml"] = numbering.replace("</w:numbering>", f'<w:num w:numId="{fresh}"><w:abstractNumId w:val="{abstract}"/></w:num></w:numbering>').encode("utf8")
    def fix(match):
        p = match.group(0)
        text = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p))
        if text in CONTINUED:
            p = re.sub(r"(<w:pStyle [^>]*/>)", rf'\1<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{fresh}"/></w:numPr>', p, count=1)
        if text == "Closing item X":
            p = p.replace("</w:pPr>", SECTION + "</w:pPr>", 1)
        return p
    files["word/document.xml"] = re.sub(r"<w:p[ >].*?</w:p>", fix, files["word/document.xml"].decode("utf8"), flags=re.S).encode("utf8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)


def main(output):
    output = Path(output).resolve()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        with tempfile.TemporaryDirectory() as temp:
            doc = word.Documents.Add()
            doc.Content.Text = "\r".join(text for _, text in ROWS)
            for i, (style, _) in enumerate(ROWS, 1):
                doc.Paragraphs(i).Style = doc.Styles(style)
            doc.Range(find(doc, "Other one").Range.Start, find(doc, "Other three").Range.End).ListFormat.ApplyNumberDefault()
            table = doc.Tables.Add(find(doc, "TABLE").Range, 3, 2)
            for (row, column), value in {(1, 1): "Component", (1, 2): "Role", (2, 1): "API", (2, 2): "Gateway",
                                         (3, 1): "Queue", (3, 2): "Buffer"}.items():
                table.Cell(row, column).Range.Text = value
            table.Borders.Enable = True
            raw = Path(temp) / "raw.docx"
            doc.SaveAs2(str(raw), 16)
            doc.Close(0)
            assembled = Path(temp) / "assembled.docx"
            assemble(raw, assembled)
            doc = word.Documents.Open(str(assembled))
            numbers = [(doc.Paragraphs(i).Range.Text.rstrip("\r"), doc.Paragraphs(i).Range.ListFormat.ListString)
                       for i in range(1, doc.Paragraphs.Count + 1) if doc.Paragraphs(i).Range.ListFormat.ListType]
            expected = ([f"{i}." for i in range(1, 8)] + [f"{i}." for i in range(8, 15)] + ["1.", "2.", "3.", "15.", "16."])
            if [n for _, n in numbers] != expected:
                raise RuntimeError(f"unexpected numbering: {numbers}")
            try:
                find(doc, "Plan step A").SeparateList()
                raise RuntimeError("SeparateList unexpectedly available on Plan step A; the layout does not reproduce the report")
            except Exception as error:  # COM error expected: "This command is not available."
                if "not available" not in str(getattr(error, "excepinfo", None) or error):
                    raise
            output.parent.mkdir(parents=True, exist_ok=True)
            doc.SaveAs2(str(output), 16)
            doc.Close(0)
            print("created", output, numbers)
    finally:
        word.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
