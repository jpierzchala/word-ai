"""Create the synthetic document for the live Office.js acceptance of structural operations.

Uses a separate invisible Word instance; never opens or changes user documents.
Usage: python make_structural_live_test_doc.py <output.docx>
"""
from pathlib import Path
import struct
import sys
import tempfile
import zlib
import win32com.client


def png():
    def chunk(name, data):
        return struct.pack(">I", len(data)) + name + data + struct.pack(">I", zlib.crc32(name + data))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 2, 1, 8, 2, 0, 0, 0)) +
            chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00\x00\xff\x00")) + chunk(b"IEND", b""))


NORMAL, H1, H2 = -1, -2, -3
LINES = [
    ("Structural operations live test", -63),  # Title
    ("TOC", NORMAL),
    ("Alpha", H1), ("Alpha body.", NORMAL),
    ("We suggest a Graph based approach (DELETE ME).", NORMAL),
    ("Lead-in: the rest of this paragraph stays regular.", NORMAL),
    ("one", NORMAL), ("two", NORMAL), ("three", NORMAL), ("Between the lists.", NORMAL), ("four", NORMAL), ("five", NORMAL),
    ("Alpha after lists.", NORMAL),
    ("Beta", H1), ("Beta sub", H2), ("Beta sub body.", NORMAL), ("PICTURE", NORMAL), ("Figure 1: synthetic picture", NORMAL),
    ("bullet candidate one", NORMAL), ("bullet candidate two", NORMAL), ("TABLE", NORMAL),
    ("Gamma", H1), ("Gamma body.", NORMAL), ("", NORMAL),
]


def main(output):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Add()
        doc.Content.Text = "\r".join(text for text, _ in LINES)
        for i, (_, style) in enumerate(LINES, 1):
            doc.Paragraphs(i).Style = doc.Styles(style)
        find = lambda text: next(doc.Paragraphs(i).Range for i in range(1, doc.Paragraphs.Count + 1)
                                 if doc.Paragraphs(i).Range.Text.rstrip("\r") == text)
        numbered = lambda first, last: doc.Range(find(first).Start, find(last).End)
        numbered("one", "three").ListFormat.ApplyNumberDefault()
        # "four" and "five" continue the same list (4., 5.): the restart target.
        numbered("four", "five").ListFormat.ApplyListTemplate(find("one").ListFormat.ListTemplate, True)
        numbered("bullet candidate one", "bullet candidate two").ListFormat.ApplyNumberDefault()
        with tempfile.TemporaryDirectory() as temp:
            image = Path(temp) / "live.png"
            image.write_bytes(png())
            picture = find("PICTURE")
            target = doc.Range(picture.Start, picture.End - 1)
            target.Text = ""
            doc.InlineShapes.AddPicture(str(image), False, True, target)
        table = doc.Tables.Add(find("TABLE"), 2, 2)  # Beta ends with the components table
        for (row, column), value in {(1, 1): "Component", (1, 2): "Role", (2, 1): "API", (2, 2): "Gateway"}.items():
            table.Cell(row, column).Range.Text = value
        table.Rows(1).HeadingFormat = True
        table.Borders.Enable = True
        toc = find("TOC")
        doc.TablesOfContents.Add(doc.Range(toc.Start, toc.End - 1), True, 1, 2)  # headings get _Toc bookmarks
        Path(output).parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs2(str(Path(output).resolve()), 16)  # wdFormatXMLDocument
        doc.Close(0)
        print("created", output)
    finally:
        word.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
