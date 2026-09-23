import test from "node:test";
import assert from "node:assert/strict";
import {DOMParser, XMLSerializer} from "@xmldom/xmldom";
import {W} from "../safety.js";
import {compareSequence, emptyParagraphXml, expectedSequence, fieldNesting, findSection, headingLevel, inspectSection, preparePackage,
  resolveDestination, sectionDigest, startsInsideField} from "../section-move.js";
import {paragraphDigest, parseFormat, planDeleteParagraph, planFormat, planList, planTableRow, searchText} from "../document-edit.js";
globalThis.DOMParser = DOMParser;

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const p = (text, style = "Normal", extra = {}) => ({uniqueLocalId: `id-${text || "empty"}-${style}`, text, style,
  styleBuiltIn: style.replace(" ", ""), outlineLevel: 10, tableNestingLevel: 0, ...extra});
const doc = () => [p("Intro"), p("A", "Heading 1"), p("a1"), p("B", "Heading 1"), p("B.1", "Heading 2"), p("b1"),
  p("x1", "Normal", {uniqueLocalId: "cell-1", tableNestingLevel: 1}), p("C", "Heading 1"), p("c1"), p("")];
const flat = (body, parts = "") => `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906"/></w:sectPr></w:body></w:document></pkg:xmlData></pkg:part>${parts}</pkg:package>`;
const para = (text, inner = "") => `<w:p w:rsidR="00AB12CD">${inner}<w:r><w:t>${text}</w:t></w:r></w:p>`;

test("heading levels come from built-in headings or custom outline levels outside tables", () => {
  assert.equal(headingLevel(p("A", "Heading 2")), 2);
  assert.equal(headingLevel({...p("A", "Custom"), styleBuiltIn: "Other", outlineLevel: 3}), 3);
  assert.equal(headingLevel({...p("A"), outlineLevel: 1}), null, "Normal paragraphs are never headings");
  assert.equal(headingLevel({...p("A", "Heading 1"), tableNestingLevel: 1}), null);
});

test("section spans subsections and tables until the next heading of the same level", () => {
  const rows = doc();
  assert.deepEqual(findSection(rows, "id-B-Heading 1", true), {start: 3, end: 7, level: 1});
  assert.deepEqual(findSection(rows, "id-B.1-Heading 2", true), {start: 4, end: 7, level: 2});
  // The last section stops before the document's empty final paragraph.
  assert.deepEqual(findSection(rows, "id-C-Heading 1", true), {start: 7, end: 9, level: 1});
  assert.throws(() => findSection(rows, "id-C-Heading 1", false), /ostatni akapit/);
  assert.throws(() => findSection(rows, "id-a1-Normal", true), /nie jest nagłówkiem/);
  assert.throws(() => findSection(rows, "missing", true), /Odczytaj snapshot/);
});

test("destination excludes the section itself, no-op positions, tables and the document end", () => {
  const rows = doc(), section = findSection(rows, "id-B-Heading 1", true);
  assert.deepEqual(resolveDestination(rows, section, "id-A-Heading 1", "before"), {target: 1, destination: 1});
  assert.deepEqual(resolveDestination(rows, section, "id-c1-Normal", "after"), {target: 8, destination: 9});
  assert.throws(() => resolveDestination(rows, section, "id-b1-Normal", "before"), /wewnątrz/);
  assert.throws(() => resolveDestination(rows, section, "id-C-Heading 1", "before"), /już znajduje/);
  assert.throws(() => resolveDestination(rows, section, "id-a1-Normal", "after"), /już znajduje/);
  assert.throws(() => resolveDestination(rows, section, "cell-1", "before"), /tabeli/);
  assert.throws(() => resolveDestination(rows, section, "id-empty-Normal", "after"), /ostatnim akapitem/);
  const tableNext = [...rows.slice(0, 2), p("t", "Normal", {uniqueLocalId: "cell-2", tableNestingLevel: 1}), ...rows.slice(2)];
  assert.throws(() => resolveDestination(tableNext, findSection(tableNext, "id-B-Heading 1", true), "id-A-Heading 1", "after"), /tabela/);
});

test("expected order detects exact moves, seam artifacts and anything else", () => {
  const rows = doc(), section = findSection(rows, "id-B-Heading 1", true);
  const expected = expectedSequence(rows, section, 1);
  const moved = [rows[0], ...rows.slice(3, 7), rows[1], rows[2], ...rows.slice(7)];
  assert.deepEqual(compareSequence(moved, expected), {outcome: "exact", extras: 0});
  const seam = [...moved.slice(0, 5), p(""), p(""), ...moved.slice(5)];
  assert.deepEqual(compareSequence(seam, expected), {outcome: "seam", extras: 2});
  assert.equal(expected.sentinelIndex, 5);
  assert.equal(compareSequence([...moved.slice(0, 5), p(""), p(""), p(""), ...moved.slice(5)], expected).outcome, "mismatch");
  assert.equal(compareSequence([...moved.slice(0, 5), p("x"), ...moved.slice(5)], expected).outcome, "mismatch");
  assert.equal(compareSequence([...moved.slice(0, 5), p("", "Normal", {tableNestingLevel: 1}), ...moved.slice(5)], expected).outcome, "mismatch");
  assert.equal(compareSequence([...moved.slice(0, 6), ...moved.slice(7)], expected).outcome, "mismatch");
  // Moving down: the destination index shifts by the moved length.
  const down = expectedSequence(rows, section, 9);
  assert.deepEqual(down.sequence.map(r => r[0]), ["Intro", "A", "a1", "C", "c1", "B", "B.1", "b1", "x1", ""]);
  assert.equal(down.sentinelIndex, 9);
});

test("section inspection blocks unsafe objects and reports what moves", () => {
  const clean = flat(para("H") + `<w:tbl><w:tr><w:tc>${para("cell")}</w:tc></w:tr></w:tbl>`);
  const ok = inspectSection(clean);
  assert.deepEqual(ok.blockers, []);
  assert.equal(ok.counts.tables, 1); assert.equal(ok.ends_with_table, true);
  const cases = {
    "komentarze": para("H", '<w:commentRangeStart w:id="1"/>'),
    "śledzone zmiany (rewizje)": `<w:p><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p>`,
    "kontrolki zawartości": `<w:sdt><w:sdtContent>${para("x")}</w:sdtContent></w:sdt>`,
    "podział sekcji Worda": `<w:p><w:pPr><w:sectPr/></w:pPr></w:p>`,
    "pole przekraczające granicę sekcji": `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>`,
    "zakładka przekraczająca granicę sekcji": `<w:p><w:bookmarkStart w:id="4" w:name="x"/></w:p>`,
  };
  for (const [blocker, body] of Object.entries(cases))
    assert.ok(inspectSection(flat(para("H") + body)).blockers.includes(blocker), blocker);
  const counted = inspectSection(flat(para("H", '<w:bookmarkStart w:id="1" w:name="_Toc1"/><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/>') +
    '<w:p><w:fldSimple w:instr="SEQ Figure"/><w:r><w:footnoteReference w:id="2"/></w:r></w:p>'));
  assert.deepEqual(counted.bookmarks, ["_Toc1"]);
  assert.equal(counted.counts.fields, 1); assert.equal(counted.counts.footnotes, 1);
  assert.throws(() => inspectSection('<!DOCTYPE x [<!ENTITY a "b">]>' + clean), /Nieobsługiwany/);
});

test("section digest ignores export noise but tracks text, formatting parts and media", async () => {
  const media = data => `<pkg:part pkg:name="/word/media/image1.png"><pkg:binaryData>${data}</pkg:binaryData></pkg:part>`;
  const base = await sectionDigest(flat(para("H"), media("AAAA")));
  assert.equal(base, await sectionDigest(flat(para("H").replace("00AB12CD", "00FFFFFF"), media("AA\nAA"))));
  assert.equal(base, await sectionDigest(flat('<w:p w:rsidR="1"><w:r><w:lastRenderedPageBreak/><w:t>H</w:t></w:r></w:p>', media("AAAA"))));
  assert.notEqual(base, await sectionDigest(flat(para("h"), media("AAAA"))));
  assert.notEqual(base, await sectionDigest(flat(para("H"), media("BBBB"))));
  const styles = s => `<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles xmlns:w="${W}"><w:style w:styleId="${s}"/></w:styles></pkg:xmlData></pkg:part>`;
  assert.notEqual(await sectionDigest(flat(para("H"), styles("A"))), await sectionDigest(flat(para("H"), styles("B"))));
});

test("package strips section properties, renames bookmarks and appends a sentinel paragraph", () => {
  const xml = flat(para("H", '<w:bookmarkStart w:id="1" w:name="_Toc42"/><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/>') + para("body"));
  const prepared = preparePackage(xml, DOMParser, XMLSerializer, "cafe0001");
  assert.deepEqual(prepared.bookmarks, [{name: "_Toc42", temp: "WAIcafe0001_0"}]);
  const out = new DOMParser().parseFromString(prepared.xml, "application/xml");
  const body = out.getElementsByTagNameNS(W, "body")[0];
  const children = Array.from(body.childNodes).filter(n => n.nodeType === 1);
  assert.deepEqual(children.map(n => n.localName), ["p", "p", "p"]);
  assert.equal(children[2].childNodes.length, 0, "sentinel is an empty paragraph");
  assert.equal(out.getElementsByTagNameNS(W, "sectPr").length, 0);
  const names = Array.from(out.getElementsByTagNameNS(W, "bookmarkStart")).map(n => n.getAttributeNS(W, "name"));
  assert.deepEqual(names, ["WAIcafe0001_0"]);
  assert.equal(out.getElementsByTagNameNS(W, "bookmarkEnd").length, 1, "_GoBack removed on both ends");
  assert.ok(prepared.bookmarks.every(b => b.temp.length <= 40 && /^[A-Za-z][A-Za-z0-9_]+$/.test(b.temp)));
});

const single = body => `<w:document xmlns:w="${W}"><w:body><w:p>${body}</w:p></w:body></w:document>`;
const current = (body = "<w:r><w:t>one two</w:t></w:r>", text = "one two") => ({xml: single(body), text, hash: "h", style: "Normal"});
const context = {isLastInBody: false, endsSection: false, inContentControl: false, tableNestingLevel: 0,
  previous: {text: "prev", tableNestingLevel: 0}, next: {text: "next", tableNestingLevel: 0}, pictures: 0};

test("paragraph deletion discloses dependent objects and blocks unsafe structure", () => {
  const plan = planDeleteParagraph(current(), {expected_sha256: "h", text: "", find: ""}, context);
  assert.equal(plan.before, "one two"); assert.equal(plan.next_text, "next"); assert.equal(plan.note, null);
  const withField = planDeleteParagraph(current('<w:bookmarkStart w:id="1" w:name="_Ref1"/><w:fldSimple w:instr="SEQ"/>'), {expected_sha256: "h", text: "", find: ""}, context);
  assert.deepEqual(withField.bookmarks, ["_Ref1"]); assert.equal(withField.field_count, 1); assert.match(withField.note, /odsyłacze/);
  const payload = {expected_sha256: "h", text: "", find: ""};
  assert.throws(() => planDeleteParagraph(current(), {...payload, expected_sha256: "x"}, context), /zmienił/);
  assert.throws(() => planDeleteParagraph(current(), {...payload, text: "x"}, context), /nie przyjmuje/);
  assert.throws(() => planDeleteParagraph(current('<w:commentRangeStart w:id="1"/>'), payload, context), /komentarze/);
  assert.throws(() => planDeleteParagraph(current('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'), payload, context), /pole/);
  assert.throws(() => planDeleteParagraph(current(), payload, {...context, isLastInBody: true}), /ostatni/);
  assert.throws(() => planDeleteParagraph(current(), payload, {...context, endsSection: true}), /podział sekcji/);
  assert.throws(() => planDeleteParagraph(current(), payload, {...context, inContentControl: true}), /kontrolce/);
  const tables = {...context, previous: {text: "", tableNestingLevel: 1}, next: {text: "", tableNestingLevel: 1}};
  assert.throws(() => planDeleteParagraph(current(), payload, tables), /scaliłoby/);
  assert.ok(planDeleteParagraph(current(), payload, {...tables, tableNestingLevel: 1}), "cell paragraph between cells of the same table");
  const drawing = '<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/></w:drawing></w:r>';
  assert.throws(() => planDeleteParagraph(current(drawing, ""), payload, context), /pływające/);
  const chart = '<w:r><w:drawing><a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></w:drawing></w:r>';
  assert.throws(() => planDeleteParagraph(current(chart, ""), payload, context), /wykresy/);
  const picture = '<w:r><w:drawing><a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></w:drawing></w:r>';
  assert.equal(planDeleteParagraph(current(picture, ""), payload, {...context, pictures: 1}).inline_picture_count, 1);
});

test("fragment formatting parses explicit tokens and escapes Word search codes", () => {
  assert.deepEqual(parseFormat("bold,no_italic"), {bold: true, italic: false});
  for (const bad of ["", "bold,no_bold", "bold,bold", "heavy"]) assert.throws(() => parseFormat(bad));
  assert.equal(searchText("x^2"), "x^^2");
  assert.throws(() => searchText("^".repeat(128)), /255/);
  assert.equal(planFormat(current(), {expected_sha256: "h", text: "bold", find: "one"}).target, "one");
  assert.equal(planFormat(current(), {expected_sha256: "h", text: "italic", find: ""}).whole_paragraph, true);
  assert.throws(() => planFormat(current("<w:r><w:t>one one</w:t></w:r>", "one one"), {expected_sha256: "h", text: "bold", find: "one"}), /dokładnie raz/);
  assert.throws(() => planFormat(current('<w:ins w:id="1"/>'), {expected_sha256: "h", text: "bold", find: "one"}), /śledzone/);
});

test("table row plan requires a simple table, data-row template and matching cell count", () => {
  const table = {nestingLevel: 1, merged: false, rowIndex: 1,
    rows: [{cellCount: 2, isHeader: true, values: ["Name", "Role"]}, {cellCount: 2, isHeader: false, values: ["a", "b"]}]};
  const payload = {expected_sha256: "h", position: "after", cells: ["c", "d"]};
  const cell = {...current(), tableNestingLevel: 1};
  assert.deepEqual(planTableRow(cell, payload, table).reference_row, ["a", "b"]);
  assert.throws(() => planTableRow(cell, {...payload, cells: ["c"]}, table), /2 komórek/);
  assert.throws(() => planTableRow(cell, {...payload, cells: ["c", "two\nlines"]}, table), /jednowierszowych/);
  assert.throws(() => planTableRow(cell, payload, {...table, rowIndex: 0}), /nagłówka/);
  assert.throws(() => planTableRow(cell, payload, {...table, merged: true}), /scalone/);
  assert.throws(() => planTableRow({...cell, tableNestingLevel: 2}, payload, {...table, nestingLevel: 2}), /niezagnieżdżone/);
});

test("list plans state their scope and refuse headings, no-ops and unsupported restarts", () => {
  const items = [{id: "a", text: "one", level: 0, listString: "1."}, {id: "b", text: "two", level: 0, listString: "2."},
    {id: "c", text: "sub", level: 1, listString: "a."}];
  const list = {isListItem: true, level: 0, listString: "2.", levelTypes: ["Number", "Number"], items, index: 1, value: 2, canSeparate: true};
  const payload = operation => ({expected_sha256: "h", operation, text: "", find: ""});
  assert.equal(planList(current(), {...payload("list_level"), text: "1"}, list).new_level, 1);
  assert.throws(() => planList(current(), {...payload("list_level"), text: "0"}, list), /ma już/);
  assert.match(planList(current(), {...payload("list_type"), text: "bullet"}, list).scope, /poziomie 0 \(2\)/);
  assert.throws(() => planList(current(), {...payload("list_type"), text: "number"}, list), /ma już/);
  assert.equal(planList(current(), payload("list_restart"), list).method, "separate_list");
  assert.equal(planList(current(), payload("list_restart"), {...list, index: 0, value: 5}).method, "set_starting_number");
  assert.throws(() => planList(current(), payload("list_restart"), {...list, index: 0, value: 1}), /już od 1/);
  assert.throws(() => planList(current(), payload("list_restart"), {...list, canSeparate: false}), /WordApiDesktop 1.4/);
  assert.throws(() => planList(current(), payload("list_restart"), {...list, levelTypes: ["Bullet"]}), /numerowanym/);
  assert.throws(() => planList({...current(), style: "Heading2"}, payload("list_restart"), list), /set_style/);
  assert.throws(() => planList(current(), payload("list_restart"), {isListItem: false}), /nie jest elementem/);
  assert.throws(() => planList(current(), payload("list_restart"), {...list, index: -1}), /treści głównej/);
  assert.throws(() => planFormat(current('<w:sdt/><w:r><w:t>one two</w:t></w:r>'), {expected_sha256: "h", text: "bold", find: "one"}), /kontrolkę/);
});

test("page-break cache does not change a paragraph fingerprint", async () => {
  assert.equal(await paragraphDigest(single("<w:r><w:t>x</w:t></w:r>")),
    await paragraphDigest(single("<w:r><w:lastRenderedPageBreak/><w:t>x</w:t></w:r>")));
});

const fld = type => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
test("field nesting is walked in order, so fields crossing both section boundaries are caught", () => {
  const crossing = flat(`<w:p>${fld("end")}<w:r><w:t>H</w:t></w:r></w:p><w:p>${fld("begin")}</w:p>`);
  assert.ok(inspectSection(crossing).blockers.includes("pole przekraczające granicę sekcji"));
  const nested = flat(`<w:p>${fld("begin")}${fld("begin")}${fld("separate")}${fld("end")}${fld("separate")}${fld("end")}<w:r><w:t>H</w:t></w:r></w:p>`);
  assert.deepEqual(inspectSection(nested).blockers, []);
  const fallback = `<w:p xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${fld("begin")}<mc:AlternateContent><mc:Choice/><mc:Fallback>${fld("end")}</mc:Fallback></mc:AlternateContent>${fld("end")}</w:p>`;
  assert.equal(fieldNesting(new DOMParser().parseFromString(`<x xmlns:w="${W}">${fallback}</x>`, "application/xml").documentElement).balanced, true);
  assert.equal(startsInsideField(flat(`<w:p>${fld("end")}<w:r><w:t>Intro</w:t></w:r></w:p>`)), true);
  assert.equal(startsInsideField(flat(`<w:p>${fld("begin")}${fld("end")}<w:r><w:t>Intro</w:t></w:r></w:p>`)), false);
});

test("bookmarks Word cannot restore and objects inside notes block a move", () => {
  assert.ok(inspectSection(flat(para("H", '<w:bookmarkStart w:id="1" w:name="getting-started"/><w:bookmarkEnd w:id="1"/>')))
    .blockers.some(b => b.includes("getting-started")));
  const note = inner => `<pkg:part pkg:name="/word/footnotes.xml"><pkg:xmlData><w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:bookmarkStart w:id="7" w:name="sep"/></w:p></w:footnote><w:footnote w:id="2"><w:p>${inner}</w:p></w:footnote></w:footnotes></pkg:xmlData></pkg:part>`;
  assert.deepEqual(inspectSection(flat(para("H"), note("<w:r><w:t>plain note</w:t></w:r>"))).blockers, []);
  assert.ok(inspectSection(flat(para("H"), note('<w:bookmarkStart w:id="3" w:name="fnTarget"/>'))).blockers.some(b => b.startsWith("przypisy")));
});

test("only a paragraph without runs, fields, drawings or bookmarks counts as an empty final paragraph", () => {
  assert.equal(emptyParagraphXml(flat('<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p>')), true);
  assert.equal(emptyParagraphXml(flat('<w:p/><w:p><w:r><w:rPr/><w:t></w:t></w:r></w:p>')), true);
  assert.equal(emptyParagraphXml(flat('<w:p><w:r><w:drawing/></w:r></w:p>')), false);
  assert.equal(emptyParagraphXml(flat('<w:p/><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>')), false);
  assert.equal(emptyParagraphXml(flat('<w:p><w:r><w:t> </w:t></w:r></w:p>')), false);
  assert.equal(emptyParagraphXml(flat('<w:p><w:bookmarkStart w:id="1" w:name="x"/></w:p>')), false);
});

test("drawing identifiers regenerated by each Word export do not change fingerprints", async () => {
  const pic = id => flat(`<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" wp14:anchorId="${id}" wp14:editId="${id}"/></w:drawing></w:r></w:p>`);
  assert.equal(await sectionDigest(pic("39015504")), await sectionDigest(pic("5C678BFC")));
});
