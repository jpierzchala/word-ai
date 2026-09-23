// Pure planning for moving a heading's section through Word's own OOXML export.
// Word merges the LAST paragraph of an inserted package into the destination
// paragraph (its mark and properties are dropped). A trailing empty sentinel
// paragraph absorbs that merge, so every real paragraph and a closing table keep
// their structure. Bookmarks are inserted under temporary names because Word
// drops inserted bookmarks whose names still exist at insertion time.
import {W, canonicalNode, digest} from "./safety.js";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
// A section with several 2 MiB pictures exports as Base64 inside Flat OPC.
export const MAX_SECTION_XML_CHARS = 64 * 1024 * 1024;
const REVISIONS = new Set(["ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveToRangeStart", "pPrChange",
  "rPrChange", "sectPrChange", "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange", "numberingChange",
  "cellIns", "cellDel", "cellMerge", "customXmlInsRangeStart", "customXmlDelRangeStart"]);
const COMMENTS = new Set(["commentRangeStart", "commentRangeEnd", "commentReference"]);
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
// Names Range.insertBookmark accepts; any other name could not be restored.
export const BOOKMARK_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,39}$/;
const elements = node => Array.from(node.childNodes).filter(n => n.nodeType === 1);
const inFallback = node => { for (let n = node.parentNode; n; n = n.parentNode) if (n.namespaceURI === MC && n.localName === "Fallback") return true; return false; };

// Field characters in document order (mc:Fallback duplicates ignored).
// balanced: every field opened here also ends here; startsInside: a field
// end or separator appears before its begin, i.e. the content starts inside
// a field that began earlier.
export function fieldNesting(root) {
  let depth = 0, startsInside = false;
  for (const node of Array.from(root.getElementsByTagNameNS(W, "fldChar"))) {
    if (inFallback(node)) continue;
    const type = node.getAttributeNS(W, "fldCharType");
    if (type === "begin") depth++;
    else if (type === "end") depth--;
    if (depth < 0 || (type === "separate" && depth === 0)) startsInside = true;
    if (depth < 0) depth = 0;
  }
  return {balanced: !startsInside && depth === 0, startsInside};
}

function parse(xml, Parser, limit = MAX_SECTION_XML_CHARS) {
  if (typeof xml !== "string" || /<!DOCTYPE|<!ENTITY/i.test(xml)) throw Error("Nieobsługiwany eksport XML.");
  if (xml.length > limit) throw Error(`Eksport ma ${xml.length} znaków; limit wynosi ${limit}.`);
  const doc = new Parser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw Error("Błędny XML.");
  return doc;
}

// Destination paragraph export: does it begin inside a field such as a TOC?
export function startsInsideField(xml, Parser = DOMParser) {
  return fieldNesting(documentBody(parse(xml, Parser))).startsInside;
}

// Word can export a blank trailing range as more than one empty paragraph or
// materialize an empty text run. Neither contains content that a move can lose.
export function emptyParagraphXml(xml, Parser = DOMParser) {
  const children = elements(documentBody(parse(xml, Parser)));
  const paragraphs = children.filter(n => n.namespaceURI === W && n.localName === "p");
  if (!paragraphs.length || children.some(n => n.namespaceURI !== W || !["p", "sectPr"].includes(n.localName))) return false;
  const emptyRun = run => elements(run).every(n => n.namespaceURI === W &&
    (n.localName === "rPr" || (n.localName === "t" && !n.textContent)));
  return paragraphs.every(p => elements(p).every(n => n.namespaceURI === W &&
    (["pPr", "proofErr", "bookmarkEnd"].includes(n.localName) ||
      (n.localName === "bookmarkStart" && n.getAttributeNS(W, "name") === "_GoBack") ||
      (n.localName === "r" && emptyRun(n)))));
}

export function emptyParagraphShape(xml, Parser = DOMParser) {
  const children = elements(documentBody(parse(xml, Parser)));
  const paragraphs = children.filter(n => n.namespaceURI === W && n.localName === "p");
  return {body: children.slice(0, 8).map(n => n.localName), count: paragraphs.length,
    paragraphs: paragraphs.slice(0, 3).map(p => elements(p).slice(0, 8).map(n => n.localName))};
}

export function headingLevel(paragraph) {
  if (paragraph.tableNestingLevel) return null;
  const match = /^Heading([1-9])$/.exec(paragraph.styleBuiltIn || "");
  if (match) return Number(match[1]);
  // Custom heading styles (styleBuiltIn "Other") carry an outline level 1–9.
  const level = paragraph.outlineLevel;
  return paragraph.styleBuiltIn === "Other" && Number.isInteger(level) && level >= 1 && level <= 9 ? level : null;
}

// paragraphs: body paragraphs in document order (including table cells) with
// uniqueLocalId, text, style, styleBuiltIn, outlineLevel and tableNestingLevel.
export function findSection(paragraphs, headingId, finalParagraphEmpty) {
  const start = paragraphs.findIndex(p => p.uniqueLocalId === headingId);
  if (start < 0) throw Error("Nagłówek nie należy do treści głównej dokumentu. Odczytaj snapshot ponownie.");
  const level = headingLevel(paragraphs[start]);
  if (!level) throw Error("Wskazany akapit nie jest nagłówkiem (Heading1–9 lub poziom konspektu 1–9) poza tabelą.");
  let end = paragraphs.length;
  for (let i = start + 1; i < paragraphs.length; i++) {
    const other = headingLevel(paragraphs[i]);
    if (other && other <= level) { end = i; break; }
  }
  if (end === paragraphs.length) {
    // The final paragraph mark of a document can be neither moved nor deleted.
    const last = paragraphs.length - 1;
    if (last === start || !finalParagraphEmpty)
      throw Error("Sekcja obejmuje ostatni akapit dokumentu. Dodaj pusty akapit na końcu (insert_after) i ponów podgląd.");
    end = last;
  }
  return {start, end, level};
}

export function resolveDestination(paragraphs, section, targetId, position) {
  const target = paragraphs.findIndex(p => p.uniqueLocalId === targetId);
  if (target < 0) throw Error("Akapit docelowy nie należy do treści głównej dokumentu. Odczytaj snapshot ponownie.");
  if (paragraphs[target].tableNestingLevel) throw Error("Cel przeniesienia nie może leżeć w tabeli.");
  if (target >= section.start && target < section.end) throw Error("Cel leży wewnątrz przenoszonej sekcji.");
  if (!["before", "after"].includes(position)) throw Error("Nieprawidłowe położenie celu.");
  const destination = position === "before" ? target : target + 1;
  if (destination >= paragraphs.length)
    throw Error("Nie można wstawić za ostatnim akapitem dokumentu. Dodaj pusty akapit na końcu i wskaż go z position=before.");
  if (paragraphs[destination].tableNestingLevel)
    throw Error("Za celem zaczyna się tabela. Wskaż akapit za tabelą z position=before.");
  if (destination === section.start || destination === section.end) throw Error("Sekcja już znajduje się w tym miejscu.");
  return {target, destination};
}

const key = p => [p.text, p.style, p.tableNestingLevel || 0];
export function expectedSequence(paragraphs, section, destination) {
  const moved = paragraphs.slice(section.start, section.end);
  const rest = [...paragraphs.slice(0, section.start), ...paragraphs.slice(section.end)];
  const at = destination > section.start ? destination - moved.length : destination;
  return {sequence: [...rest.slice(0, at), ...moved, ...rest.slice(at)].map(key), sentinelIndex: at + moved.length};
}

// Word keeps the sentinel, and after a closing table also its own implicit
// paragraph, as empty paragraphs at the seam before the destination. Returns
// {outcome: "exact"} or {outcome: "seam", extras} when only 1–2 such empty
// body paragraphs at sentinelIndex differ (they must then be deleted), else
// {outcome: "mismatch"}.
export const MAX_SEAM_PARAGRAPHS = 2;
export function compareSequence(actual, expected) {
  const same = (a, b) => a.length === b.length && a.every((x, i) => JSON.stringify(x) === JSON.stringify(b[i]));
  const rows = actual.map(key);
  if (same(rows, expected.sequence)) return {outcome: "exact", extras: 0};
  const at = expected.sentinelIndex;
  for (let extras = 1; extras <= MAX_SEAM_PARAGRAPHS; extras++) {
    const seam = rows.slice(at, at + extras);
    if (rows.length === expected.sequence.length + extras && seam.length === extras &&
        seam.every(r => r[0] === "" && r[2] === 0) && same([...rows.slice(0, at), ...rows.slice(at + extras)], expected.sequence))
      return {outcome: "seam", extras};
  }
  return {outcome: "mismatch", extras: 0};
}

function documentBody(doc) {
  const parts = Array.from(doc.getElementsByTagNameNS(PKG, "part"));
  const main = parts.filter(p => p.getAttributeNS(PKG, "name") === "/word/document.xml");
  const scope = main.length === 1 ? main[0] : parts.length ? null : doc;
  const bodies = scope ? scope.getElementsByTagNameNS(W, "body") : [];
  if (bodies.length !== 1) throw Error("Nieobsługiwana struktura eksportu sekcji.");
  return bodies[0];
}

export function parseSection(xml, Parser = DOMParser) {
  const doc = parse(xml, Parser);
  const body = documentBody(doc);
  const children = elements(body);
  // Word wraps every range export in the section properties of its location.
  const content = children.length && children[children.length - 1].namespaceURI === W &&
    children[children.length - 1].localName === "sectPr" ? children.slice(0, -1) : children;
  return {doc, body, content};
}

export function inspectSection(xml, Parser = DOMParser) {
  const {doc, body, content} = parseSection(xml, Parser);
  const all = Array.from(body.getElementsByTagName("*"));
  const count = (ns, name) => all.filter(n => n.namespaceURI === ns && n.localName === name).length;
  const blockers = [];
  if (all.some(n => n.namespaceURI === W && REVISIONS.has(n.localName))) blockers.push("śledzone zmiany (rewizje)");
  if (all.some(n => n.namespaceURI === W && COMMENTS.has(n.localName))) blockers.push("komentarze");
  if (count(W, "sdt")) blockers.push("kontrolki zawartości");
  if (all.some(n => n.namespaceURI === W && n.localName === "sectPr" && n.parentNode?.localName === "pPr"))
    blockers.push("podział sekcji Worda");
  if (count(W, "altChunk") || count(W, "subDoc")) blockers.push("osadzone dokumenty");
  const fieldChars = all.filter(n => n.namespaceURI === W && n.localName === "fldChar" && !inFallback(n)).map(n => n.getAttributeNS(W, "fldCharType"));
  if (!fieldNesting(body).balanced) blockers.push("pole przekraczające granicę sekcji");
  const starts = all.filter(n => n.namespaceURI === W && n.localName === "bookmarkStart");
  const ends = new Set(all.filter(n => n.namespaceURI === W && n.localName === "bookmarkEnd").map(n => n.getAttributeNS(W, "id")));
  const bookmarks = starts.map(n => n.getAttributeNS(W, "name")).filter(name => name !== "_GoBack");
  if (starts.some(n => !ends.has(n.getAttributeNS(W, "id")))) blockers.push("zakładka przekraczająca granicę sekcji");
  const unnamed = bookmarks.filter(name => !BOOKMARK_NAME.test(name));
  if (unnamed.length) blockers.push(`zakładki o nazwach, których Word nie pozwala odtworzyć (${unnamed.slice(0, 5).join(", ")})`);
  // Notes travel inside the package; their bookmarks would be dropped as
  // duplicates at insertion and revisions, comments or controls in them escape the body checks.
  for (const part of Array.from(doc.getElementsByTagNameNS(PKG, "part"))) {
    if (!/^\/word\/(footnotes|endnotes)\.xml$/.test(part.getAttributeNS(PKG, "name"))) continue;
    const notes = Array.from(part.getElementsByTagNameNS(W, "*"))
      .filter(n => ["footnote", "endnote"].includes(n.localName) && !["separator", "continuationSeparator", "continuationNotice"].includes(n.getAttributeNS(W, "type")));
    const inner = notes.flatMap(n => Array.from(n.getElementsByTagNameNS(W, "*")));
    if (inner.some(n => REVISIONS.has(n.localName) || COMMENTS.has(n.localName) || n.localName === "sdt" ||
        (n.localName === "bookmarkStart" && n.getAttributeNS(W, "name") !== "_GoBack")))
      blockers.push("przypisy z zakładkami, komentarzami, rewizjami lub kontrolkami");
  }
  if (!content.length || content[0].namespaceURI !== W || content[0].localName !== "p") blockers.push("sekcja nie zaczyna się akapitem");
  const topTables = content.filter(n => n.namespaceURI === W && n.localName === "tbl").length;
  return {doc, blockers, bookmarks,
    counts: {paragraphs: count(W, "p"), tables: count(W, "tbl"), top_level_tables: topTables,
      inline_pictures: count(WP, "inline"), floating_objects: count(WP, "anchor"),
      fields: fieldChars.filter(t => t === "begin").length + count(W, "fldSimple"),
      footnotes: count(W, "footnoteReference"), endnotes: count(W, "endnoteReference"),
      direct_list_paragraphs: count(W, "numPr"), bookmarks: bookmarks.length},
    ends_with_table: content.length > 0 && content[content.length - 1].localName === "tbl"};
}

// Structure, formatting dependencies and media of the export; technical IDs,
// proofing marks and layout-only page-break caches are ignored by canonicalNode.
export async function sectionDigest(xml, Parser = DOMParser) {
  const {doc, content} = parseSection(xml, Parser);
  const parts = Array.from(doc.getElementsByTagNameNS(PKG, "part")).map(p => [p.getAttributeNS(PKG, "name"), p]);
  const formatting = parts
    .filter(([name]) => /^\/word\/(styles(?:WithEffects)?\.xml|fontTable\.xml|numbering\.xml|theme\/[^/]+\.xml)$/.test(name))
    .map(([name, p]) => [name, Array.from(p.childNodes).map(canonicalNode).filter(x => x !== null)])
    .sort((a, b) => a[0].localeCompare(b[0]));
  const media = [];
  for (const [name, p] of parts) {
    const binary = p.getElementsByTagNameNS(PKG, "binaryData")[0];
    if (binary) media.push([name, await digest(binary.textContent.replace(/\s+/g, ""))]);
  }
  media.sort((a, b) => a[0].localeCompare(b[0]));
  return digest(JSON.stringify([content.map(canonicalNode), formatting, media]));
}

export function randomToken() {
  return Array.from(crypto.getRandomValues(new Uint8Array(4)), b => b.toString(16).padStart(2, "0")).join("");
}

// Package for insertion: no synthetic section properties, temporary bookmark
// names (restored after the original is deleted) and a trailing sentinel.
export function preparePackage(xml, Parser = DOMParser, Serializer = XMLSerializer, token = randomToken()) {
  const {doc, body, content} = parseSection(xml, Parser);
  const children = elements(body);
  if (children.length !== content.length) body.removeChild(children[children.length - 1]);
  const renames = [];
  const starts = Array.from(body.getElementsByTagNameNS(W, "bookmarkStart"));
  for (const start of starts) {
    const name = start.getAttributeNS(W, "name");
    if (name === "_GoBack") {
      const id = start.getAttributeNS(W, "id");
      for (const end of Array.from(body.getElementsByTagNameNS(W, "bookmarkEnd")))
        if (end.getAttributeNS(W, "id") === id) end.parentNode.removeChild(end);
      start.parentNode.removeChild(start);
      continue;
    }
    // Visible, letter-first temporary name: resolvable like any user bookmark.
    const temp = `WAI${token}_${renames.length}`;
    start.setAttributeNS(W, "w:name", temp);
    renames.push({name, temp});
  }
  body.appendChild(doc.createElementNS(W, "w:p"));
  return {xml: new Serializer().serializeToString(doc), bookmarks: renames};
}
