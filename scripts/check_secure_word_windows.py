"""Synthetic Windows Word COM check of the same replacement XML as Office.js.
This is NOT an end-to-end test of the Office.js add-in or its consent UI.
Requires Word, pywin32, Node and `npm ci --ignore-scripts` in secure-addin.
"""
import os
from pathlib import Path
import subprocess
import win32com.client

root = Path(os.environ["LOCALAPPDATA"]) / "WordAiSecure" / "validation"
root.mkdir(parents=True, exist_ok=True)
repo = Path(__file__).resolve().parents[1] / "secure-addin"
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
doc = None
try:
    doc = word.Documents.Add()
    old = "Word AI synthetic pilot text."
    prefix = "BEFORE\r"
    doc.Content.Text = prefix + old + "\rAFTER"
    target = doc.Range(len(prefix), len(prefix) + len(old))
    target.Font.Bold = -1
    control = doc.ContentControls.Add(0, target)
    control.Tag = "WORD-AI:pilot"
    original_id = control.ID
    before_text = doc.Content.Text
    xml_path = root / "word-range.xml"
    xml_path.write_text(control.Range.WordOpenXML, encoding="utf8")
    program = """
import fs from 'node:fs';
import {DOMParser, XMLSerializer} from '@xmldom/xmldom';
import {replacementXml} from './safety.js';
const xml = fs.readFileSync(process.env.INPUT_XML, 'utf8');
process.stdout.write(replacementXml(xml, 'Word AI verified replacement', DOMParser, XMLSerializer));
"""
    replacement = subprocess.run(["node", "--input-type=module", "-e", program], cwd=repo,
        env={**os.environ, "INPUT_XML": str(xml_path)}, capture_output=True, text=True, check=True).stdout
    control.Range.InsertXML(replacement)
    assert doc.ContentControls.Count == 1, "Control lost"
    current = doc.ContentControls.Item(1)
    assert current.ID == original_id and current.Tag == "WORD-AI:pilot", "Control identity lost"
    assert current.Range.Text == "Word AI verified replacement", repr(current.Range.Text)
    assert current.Range.Font.Bold == -1, "Bold formatting lost"
    assert doc.Content.Text == before_text.replace(old, "Word AI verified replacement", 1), repr(doc.Content.Text)
    # No user document is opened, changed or overwritten.
    print("PASS: actual Word text, bold formatting, control identity and outside text preserved (COM, not Office.js).")
finally:
    if doc is not None:
        doc.Close(0)
    word.Quit()
