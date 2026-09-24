import test from "node:test";
import assert from "node:assert/strict";
import {DOMParser, XMLSerializer} from "@xmldom/xmldom";
import {W} from "../safety.js";
import {compareSequence, emptyParagraphXml, expectedSequence, fieldNesting, findSection, headingLevel, inspectSection, preparePackage,
  resolveDestination, sectionDigest, startsInsideField} from "../section-move.js";
import {levelStart, paragraphDigest, parseFormat, planDeleteParagraph, planFormat, planList, planTableRow, restartOutcome, restartPackage,
  restartRun, searchText} from "../document-edit.js";
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
  const restart = {package: {xml: "<pkg/>", shape: "s", own_start: null}, run: ["later-list-item"], sharesInstance: true};
  const midList = planList(current(), payload("list_restart"), list, restart);
  assert.equal(midList.method, "start_override"); assert.equal(midList.renumbered_item_count, 2, "the run may reach later Office.js lists");
  assert.throws(() => planList(current(), payload("list_restart"), list, {...restart, package: {...restart.package, own_start: "5"}}), /własną wartość początkową \(5\)/);
  assert.equal(planList(current(), payload("list_restart"), list, {...restart, package: {...restart.package, own_start: "5"}, sharesInstance: false}).method, "start_override");
  // Level 0: even the first item of a list restarts through a new instance (it may continue an earlier one).
  assert.equal(planList(current(), payload("list_restart"), {...list, index: 0, value: 5}, restart).method, "start_override");
  assert.throws(() => planList(current(), payload("list_restart"), {...list, index: 0, value: 5}), /brak eksportu akapitu\. Użyj/);
  const deeper = {...list, level: 1, index: 2, value: 3, listString: "c."};
  assert.equal(planList(current(), payload("list_restart"), deeper, {levelStart: 3}).method, "set_starting_number", "the number comes from the level's start value");
  assert.throws(() => planList(current(), payload("list_restart"), deeper, {levelStart: 1}), /kontynuuje wcześniejszy element/);
  assert.throws(() => planList(current(), payload("list_restart"), {...list, index: 0, value: 1}), /już od 1/);
  assert.throws(() => planList(current(), payload("list_restart"), {...list, value: 1}, restart), /już od 1/, "an existing restart is a no-op");
  assert.throws(() => planList(current(), payload("list_restart"), list, null), /brak eksportu akapitu.*Restart at 1/);
  assert.throws(() => planList(current(), payload("list_restart"), list, {reason: "akapit kończy sekcję Worda"}), /kończy sekcję Worda/);
  assert.throws(() => planList(current(), payload("list_restart"), {...list, value: null}, restart), /WordApiDesktop 1.3/);
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

// Whole-range exports as Word produces them (verified through COM): w:pPr with
// w:numPr or a numbered style, and a numbering part with the list definition.
const CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const numbered = ({pPr = '<w:pStyle w:val="ListNumber"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:ind w:left="720"/>',
  runs = '<w:r><w:rPr><w:b/></w:rPr><w:t>Lead:</w:t></w:r><w:bookmarkStart w:id="1" w:name="_Ref1"/><w:r><w:t xml:space="preserve"> text</w:t></w:r><w:bookmarkEnd w:id="1"/>',
  wrap = p => p, nums = '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2" w16cid:durableId="42"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"/></w:lvlOverride></w:num>',
  nsid = '<w:nsid w:val="FFFFFF88"/>', numbering = true} = {}) =>
  flat(wrap(`<w:p w14:paraId="1A2B3C4D" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:pPr>${pPr}</w:pPr>${runs}</w:p>`),
    `<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/><w:style w:type="paragraph" w:styleId="ListBase"><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ListNumber"><w:basedOn w:val="ListBase"/></w:style></w:styles></pkg:xmlData></pkg:part>` +
    (numbering ? `<pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData><w:numbering xmlns:w="${W}" xmlns:w16cid="${CID}"><w:abstractNum w:abstractNumId="0">${nsid}<w:lvl w:ilvl="0"><w:start w:val="1"/></w:lvl></w:abstractNum>${nums}<w:numIdMacAtCleanup w:val="2"/></w:numbering></pkg:xmlData></pkg:part>` : ""));
const parsed = xml => new DOMParser().parseFromString(xml, "application/xml");
const kids = node => Array.from(node.childNodes).filter(n => n.nodeType === 1);
const numById = (doc, id) => Array.from(doc.getElementsByTagNameNS(W, "num")).find(n => n.getAttributeNS(W, "numId") === id);

test("restart package: one empty paragraph with the item's properties and a new instance of the same list starting at 1", () => {
  const result = restartPackage(numbered(), 0, DOMParser, XMLSerializer);
  assert.equal(result.num_id, "2"); assert.equal(result.nsid, "FFFFFF88");
  assert.equal(parsed(result.xml).getElementsByTagNameNS(W, "nsid")[0].getAttributeNS(W, "val"), "FFFFFF88", "the package keeps the definition Word merges by");
  const doc = parsed(result.xml);
  const body = doc.getElementsByTagNameNS(W, "body")[0];
  assert.deepEqual(kids(body).map(n => n.localName), ["p", "p", "sectPr"]);
  assert.equal(kids(kids(body)[1]).length, 0, "an empty sentinel takes over the item's former mark");
  const paragraph = kids(body)[0];
  assert.deepEqual(kids(paragraph).map(n => n.localName), ["pPr"], "no runs: the item's text is never rewritten");
  assert.equal(paragraph.attributes.length, 0, "no paragraph identity copied");
  const pPr = kids(paragraph)[0];
  assert.deepEqual(kids(pPr).map(n => n.localName), ["pStyle", "numPr", "ind"]);
  assert.deepEqual(kids(kids(pPr)[1]).map(n => [n.localName, n.getAttributeNS(W, "val")]), [["ilvl", "0"], ["numId", "3"]]);
  const fresh = numById(doc, "3");
  assert.equal(kids(fresh)[0].getAttributeNS(W, "val"), "0", "same abstractNum");
  assert.equal(fresh.hasAttributeNS(CID, "durableId"), false, "Word assigns the durable ID");
  const overrides = kids(fresh).filter(n => n.localName === "lvlOverride");
  assert.deepEqual(overrides.map(o => [o.getAttributeNS(W, "ilvl"), kids(o).map(n => n.localName)]), [["1", ["lvl"]], ["0", ["startOverride"]]]);
  assert.equal(kids(overrides[1])[0].getAttributeNS(W, "val"), "1");
  assert.equal(kids(fresh.parentNode).at(-1).localName, "numIdMacAtCleanup", "the new instance follows the other w:num elements");
  assert.equal(kids(fresh.parentNode).at(-2), fresh);
  assert.equal(kids(numById(doc, "2")).length, 2, "the item's former instance is unchanged");
});

test("restart package resolves numbering through the style chain and replaces an existing start override", () => {
  const styled = parsed(restartPackage(numbered({pPr: '<w:pStyle w:val="ListNumber"/><w:jc w:val="left"/>'}), 0, DOMParser, XMLSerializer).xml);
  const pPr = styled.getElementsByTagNameNS(W, "pPr")[0];
  assert.deepEqual(kids(pPr).map(n => n.localName), ["pStyle", "numPr", "jc"]);
  assert.equal(numById(styled, "3").getElementsByTagNameNS(W, "abstractNumId")[0].getAttributeNS(W, "val"), "0");
  const overridden = numbered({nums: '<w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/><w:lvl w:ilvl="0"/></w:lvlOverride></w:num>'});
  const override = parsed(restartPackage(overridden, 0, DOMParser, XMLSerializer).xml);
  const fresh = numById(override, "3").getElementsByTagNameNS(W, "lvlOverride")[0];
  assert.deepEqual(kids(fresh).map(n => [n.localName, n.getAttributeNS(W, "val")]), [["startOverride", "1"], ["lvl", null]]);
  assert.throws(() => restartPackage(numbered({wrap: p => `<w:tbl><w:tr><w:tc>${p}</w:tc></w:tr></w:tbl>`}), 0, DOMParser, XMLSerializer),
    /leży w tabeli/, "table-cell items are not validated in Word");
});

test("restart package refuses what Word would not restart as previewed", () => {
  const cases = [
    [{pPr: '<w:pStyle w:val="ListNumber"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:sectPr/>'}, /kończy sekcję Worda/],
    [{pPr: '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:pPrChange w:id="1"/>'}, /śledzoną zmianę/],
    [{pPr: '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:rPr><w:ins w:id="2" w:author="a"/></w:rPr>'}, /śledzoną zmianę/],
    [{wrap: p => `<w:sdt><w:sdtContent>${p}</w:sdtContent></w:sdt>`}, /kontrolce zawartości/],
    [{numbering: false}, /definicji numeracji/],
    [{pPr: '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'}, /definicji numeracji/],
    [{pPr: '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>'}, /poziom numeracji/],
    [{nsid: ""}, /nsid/],
    [{runs: '<w:r><w:t>x</w:t></w:r></w:p><w:p><w:r><w:t>y</w:t></w:r>'}, /nieoczekiwaną strukturę \(p\(r=1,t=1,num\) p\(r=1,t=1\) sectPr\)/],
  ];
  for (const [options, message] of cases)
    assert.throws(() => restartPackage(numbered(options), 0, DOMParser, XMLSerializer), message, JSON.stringify(options));
  assert.throws(() => restartPackage('<!DOCTYPE x [<!ENTITY a "b">]>' + numbered(), 0, DOMParser, XMLSerializer), /Nieobsługiwany/);
});

test("restart outcome accepts only a restarting numbering instance of the same list definition", () => {
  const before = numbered();
  const restarted = '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>';
  const pPr = id => `<w:pStyle w:val="ListNumber"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr><w:ind w:left="720"/>`;
  assert.deepEqual(restartOutcome(before, numbered({pPr: pPr(3), nums: restarted}), 0, DOMParser), []);
  assert.ok(restartOutcome(before, numbered({pPr: pPr(3), nums: restarted, runs: "<w:r><w:t>Lead: text</w:t></w:r>"}), 0, DOMParser).includes("treść akapitu"));
  assert.ok(restartOutcome(before, numbered({pPr: pPr(3).replace("720", "360"), nums: restarted}), 0, DOMParser).includes("właściwości akapitu"));
  assert.ok(restartOutcome(before, numbered({pPr: pPr(2), nums: restarted}), 0, DOMParser).some(p => p.includes("startOverride")));
  assert.ok(restartOutcome(before, numbered({pPr: pPr(3), nums: restarted, nsid: '<w:nsid w:val="12345678"/>'}), 0, DOMParser).includes("definicja listy akapitu"));
});

test("the run a restart renumbers follows the list definition across numbering instances", () => {
  const styles = `<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles xmlns:w="${W}"/></pkg:xmlData></pkg:part>`;
  const numbering = `<pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="AAAA0000"/></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:nsid w:val="BBBB1111"/></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num></w:numbering></pkg:xmlData></pkg:part>`;
  const p = (num, ilvl = 0) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${num}"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
  const span = parts => flat(parts.join("") + "<w:p/>", styles + numbering);  // Office.js appends an empty paragraph
  const all = (...rows) => rows.map(([id, level, value]) => ({id, level, value}));
  // item t (8) on instance 1; u on instance 1, then v, w on instance 2 of the same definition; o on another definition in between.
  const layout = all(["t", 0, 8], ["u", 0, 9], ["o", 0, 1], ["v", 0, 10], ["d", 1, 1], ["w", 0, 11], ["r", 0, 1], ["z", 0, 2]);
  const xml = span([p(1), p(1), p(3), p(2), p(2, 1), p(2), p(2), p(2)]);
  assert.deepEqual(restartRun(xml, layout, 0, 0, DOMParser), {run: ["u", "v", "w"], sharesInstance: true});
  assert.deepEqual(restartRun(span([p(1), p(2), p(3)]), all(["t", 0, 8], ["v", 0, 9], ["o", 0, 1]), 0, 0, DOMParser), {run: ["v"], sharesInstance: false});
  assert.throws(() => restartRun(span([p(1), p(2)]), layout, 0, 0, DOMParser), /nie odpowiada elementom list \(2\/8\)/);
  assert.throws(() => restartRun(span([p(1), p(1, 1)]), all(["t", 0, 8], ["u", 0, 9]), 0, 0, DOMParser), /poziomy w eksporcie/);
  // A shallower item of the definition ends the run of a deeper restart.
  assert.deepEqual(restartRun(span([p(1, 1), p(1, 1), p(1), p(1, 1)]), all(["s", 1, 3], ["u", 1, 4], ["q", 0, 2], ["x", 1, 1]), 0, 1, DOMParser).run, ["u"]);
});

test("a level's start value comes from the instance's override, else from the definition", () => {
  assert.equal(levelStart(numbered(), 0, DOMParser), 1);
  const five = numbered({nums: '<w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>'});
  assert.equal(levelStart(five, 0, DOMParser), 5);
});

const restarted = '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="5"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>';
const levels = count => Array.from({length: count}, (_, l) => `<w:lvl w:ilvl="${l}"><w:start w:val="${l === 8 ? 3 : 1}"/></w:lvl>`).join("");
const withLevels = (xml, count) => xml.replace('<w:lvl w:ilvl="0"><w:start w:val="1"/></w:lvl>', levels(count));

test("a restart that would equal an existing instance gets a neutral override of another level", () => {
  // Single-level definition (List Number): level 8 is undefined, so the override gets a unique start value exports omit.
  const single = restartPackage(numbered({nums: restarted}), 0, DOMParser, XMLSerializer);
  assert.equal(single.hidden, true);
  const hidden = kids(numById(parsed(single.xml), "6")).filter(n => n.localName === "lvlOverride");
  assert.deepEqual(hidden.map(o => o.getAttributeNS(W, "ilvl")), ["0", "8"]);
  const value = Number(kids(hidden[1])[0].getAttributeNS(W, "val"));
  assert.ok(value >= 1000 && value < 10000, String(value));
  // Multi-level: level 8 restarts at its own start (3), visible in exports.
  const multi = restartPackage(withLevels(numbered({nums: restarted}), 9), 0, DOMParser, XMLSerializer);
  assert.equal(multi.hidden, false);
  assert.deepEqual(kids(numById(parsed(multi.xml), "6")).filter(n => n.localName === "lvlOverride").map(o => [o.getAttributeNS(W, "ilvl"), kids(o)[0].getAttributeNS(W, "val")]),
    [["0", "1"], ["8", "3"]]);
  const both = restarted +
    '<w:num w:numId="6"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride><w:lvlOverride w:ilvl="8"><w:startOverride w:val="3"/></w:lvlOverride></w:num>';
  const third = numById(parsed(restartPackage(withLevels(numbered({nums: both}), 9), 0, DOMParser, XMLSerializer).xml), "7");
  assert.deepEqual(kids(third).filter(n => n.localName === "lvlOverride").map(o => o.getAttributeNS(W, "ilvl")), ["0", "8", "7"]);
  assert.equal(restartPackage(numbered(), 0, DOMParser, XMLSerializer).hidden, false, "no override needed when nothing collides");
});

test("the restart shape changes when the item's properties or another instance of its definition change", () => {
  const nums = '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>';
  const plain = restartPackage(numbered({nums}), 0, DOMParser, XMLSerializer).shape;
  assert.equal(restartPackage(numbered({nums}), 0, DOMParser, XMLSerializer).shape, plain, "stable for the same export");
  const sibling = nums + '<w:num w:numId="7"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride></w:num>';
  assert.notEqual(restartPackage(numbered({nums: sibling}), 0, DOMParser, XMLSerializer).shape, plain);
  const indented = '<w:pStyle w:val="ListNumber"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:ind w:left="360"/>';
  assert.notEqual(restartPackage(numbered({nums, pPr: indented}), 0, DOMParser, XMLSerializer).shape, plain);
});

test("restart outcome rejects an item that ended up on an existing instance, unless only a hidden override tells them apart", () => {
  const pPr = id => `<w:pStyle w:val="ListNumber"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr><w:ind w:left="720"/>`;
  const before = numbered({nums: restarted});
  assert.ok(restartOutcome(before, numbered({pPr: pPr(5), nums: restarted}), 0, DOMParser).includes("Word użył istniejącego wystąpienia listy"));
  assert.deepEqual(restartOutcome(before, numbered({pPr: pPr(5), nums: restarted}), 0, DOMParser, null, true), []);
});

test("list restart plan refuses with the reason Word could not restart", () => {
  const items = [{id: "a", text: "one", level: 0, listString: "1.", value: 1}, {id: "b", text: "two", level: 0, listString: "2.", value: 2}];
  const list = {isListItem: true, level: 0, listString: "2.", levelTypes: ["Number"], items, index: 1, value: 2};
  const payload = {expected_sha256: "h", operation: "list_restart", text: "", find: ""};
  assert.throws(() => planList(current(), payload, list, {reason: "akapit kończy sekcję Worda."}), /kończy sekcję Worda\. Użyj w Wordzie/);
});

test("an Office.js Whole export with an extra empty paragraph still yields the item, checked against its text", () => {
  const trailing = numbered().replace("<w:sectPr>", '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p><w:sectPr>');
  const plain = restartPackage(numbered(), 0, DOMParser, XMLSerializer, "Lead: text");
  const extra = restartPackage(trailing, 0, DOMParser, XMLSerializer, "Lead: text");
  assert.equal(extra.shape, plain.shape, "the artifact paragraph does not change the plan");
  assert.throws(() => restartPackage(trailing, 0, DOMParser, XMLSerializer, "another item"), /nie odpowiada tekstowi akapitu \(p\(r=2,t=10,num\) p\(r=0,t=0\) sectPr\)/);
  assert.deepEqual(restartOutcome(trailing, numbered(), 0, DOMParser, "Lead: text").filter(p => p !== "brak startOverride=1 w numeracji akapitu" &&
    p !== "Word użył istniejącego wystąpienia listy"), [], "before and after exports may differ in artifacts only");
});
