import test from "node:test";
import assert from "node:assert/strict";
import {DOMParser, XMLSerializer} from "@xmldom/xmldom";
import {inspectXml, replacementXml, structuralDigest, W} from "../safety.js";

const xml = inner => `<w:document xmlns:w="${W}"><w:body>${inner}</w:body></w:document>`;
const p = inner => `<w:p>${inner}</w:p>`;
const run = text => `<w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r>`;

test("single styled run keeps properties and safely escapes user text", () => {
  const source = xml(p(`<w:pPr><w:pStyle w:val="Normal"/></w:pPr>${run("before")}`));
  const output = replacementXml(source, "<script>& new", DOMParser, XMLSerializer);
  const result = inspectXml(output, DOMParser);
  assert.equal(result.oldText, "<script>& new");
  assert.equal(result.paragraph.getElementsByTagNameNS(W, "b").length, 1);
  assert.equal(result.paragraph.getElementsByTagNameNS(W, "pStyle")[0].getAttributeNS(W, "val"), "Normal");
  assert.equal(result.paragraph.getElementsByTagName("script").length, 0);
});

for (const [name, body] of Object.entries({
  table: `<w:tbl><w:tr><w:tc>${p(run("a"))}</w:tc></w:tr></w:tbl>`,
  field: p(`<w:r><w:fldChar w:fldCharType="begin"/></w:r>${run("a")}`),
  hyperlink: p(`<w:hyperlink>${run("a")}</w:hyperlink>`),
  comments: p(`<w:commentRangeStart w:id="1"/>${run("a")}`),
  revision: p(`<w:ins>${run("a")}</w:ins>`),
  drawing: p(`${run("a")}<w:r><w:drawing/></w:r>`),
  nested: p(`<w:sdt>${run("a")}</w:sdt>`),
  numbering: p(`<w:pPr><w:numPr/></w:pPr>${run("a")}`),
  multiRun: p(run("a") + run("b")),
  multiParagraph: p(run("a")) + p(run("b")),
  sectionBreakInParagraph: p(`<w:pPr><w:sectPr/></w:pPr>${run("a")}`),
  foreignNamespace: p(`<evil:t xmlns:evil="https://evil.test">a</evil:t>${run("b")}`),
})) {
  test(`reject ${name} before mutation`, () => assert.throws(() => replacementXml(xml(body), "new", DOMParser, XMLSerializer)));
}

test("reject DTD and non-inline replacement", () => {
  assert.throws(() => inspectXml(`<!DOCTYPE a [<!ENTITY x SYSTEM "file:///secrets">]>${xml(p(run("a")))}`, DOMParser));
  assert.throws(() => replacementXml(xml(p(run("a"))), "new\nparagraph", DOMParser, XMLSerializer));
});

test("Word synthetic trailing section is never reinserted", () => {
  const source = xml(p(run("before")) + '<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>');
  const output = replacementXml(source, "after", DOMParser, XMLSerializer);
  assert.equal(inspectXml(output, DOMParser).oldText, "after");
  assert.ok(!output.includes("sectPr"));
  assert.ok(!output.includes("pgSz"));
});

test("structural fingerprint ignores export IDs, prefixes and namespace declarations", async () => {
  const first = xml(p(run("before"))).replace('<w:p>', '<w:p w:rsidR="AABBCCDD" xmlns:unused="urn:first">');
  const second = first.replace('AABBCCDD', '11223344').replace('urn:first', 'urn:second').replaceAll('w:', 'q:').replace('xmlns:w=', 'xmlns:q=');
  assert.equal(await structuralDigest(first, DOMParser), await structuralDigest(second, DOMParser));
});

test("structural fingerprint rejects stale text and direct formatting", async () => {
  const first = xml(p(run("before")));
  const hash = await structuralDigest(first, DOMParser);
  for (const changed of [first.replace('before', 'after'), first.replace('<w:b/>', '<w:i/>'), first.replace('w:val="24"', 'w:val="26"')]) {
    assert.notEqual(hash, await structuralDigest(changed, DOMParser));
  }
});

test("formatting dependencies remain part of the fingerprint", async () => {
  const pkg = 'http://schemas.microsoft.com/office/2006/xmlPackage';
  const first = `<pkg:package xmlns:pkg="${pkg}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${xml(p(run('before')))}</pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles><w:style w:styleId="Normal"><w:rPr><w:sz w:val="24"/></w:rPr></w:style></w:styles></pkg:xmlData></pkg:part></pkg:package>`;
  assert.notEqual(await structuralDigest(first, DOMParser), await structuralDigest(first.replace('w:val="24"/></w:rPr></w:style>', 'w:val="26"/></w:rPr></w:style>'), DOMParser));
});
