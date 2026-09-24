import {W, canonicalNode, digest} from "./safety.js";
import {fieldNesting} from "./section-move.js";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
// Flat OPC includes Base64 media and formatting dependencies. An allowed 2 MiB
// image alone needs ~2.8 million characters; the former 2M text guard was too small.
export const MAX_PARAGRAPH_XML_CHARS = 8 * 1024 * 1024;
export const STYLES = ["Normal", "Title", "Subtitle", ...Array.from({length: 9}, (_, i) => `Heading${i + 1}`)];

export function paragraphXml(xml, Parser = DOMParser) {
  if (typeof xml !== "string" || /<!DOCTYPE|<!ENTITY/i.test(xml)) throw Error("Nieobsługiwany eksport XML.");
  if (xml.length > MAX_PARAGRAPH_XML_CHARS)
    throw Error(`Eksport akapitu ma ${xml.length} znaków; limit wynosi ${MAX_PARAGRAPH_XML_CHARS}.`);
  const doc = new Parser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw Error("Błędny XML.");
  const body = doc.getElementsByTagNameNS(W, "body");
  if (body.length !== 1) throw Error("Nieobsługiwana struktura zakresu.");
  const paragraphs = body[0].getElementsByTagNameNS(W, "p");
  if (paragraphs.length !== 1) throw Error("Zakres obejmuje złożony obiekt lub wiele akapitów.");
  const paragraph = paragraphs[0];
  const children = node => Array.from(node.childNodes).filter(n => n.nodeType === 1);
  let plain = true;
  const runProperties = [];
  for (const child of children(paragraph)) {
    if (child.namespaceURI !== W) {plain = false; continue;}
    if (child.localName === "proofErr") continue; // ephemeral spelling markers
    if (child.localName === "pPr") {
      if (child.getElementsByTagNameNS(W, "sectPr").length || child.getElementsByTagNameNS(W, "pPrChange").length) plain = false;
      continue;
    }
    if (child.localName !== "r") {plain = false; continue;}
    const properties = [];
    for (const node of children(child)) {
      if (node.namespaceURI !== W || !["rPr", "t", "tab", "br"].includes(node.localName)) plain = false;
      if (node.localName === "rPr") {
        if (node.getElementsByTagNameNS(W, "rPrChange").length) plain = false;
        properties.push(...children(node).map(canonicalNode).filter(x => x !== null));
      }
    }
    runProperties.push(JSON.stringify(properties));
  }
  const uniform = new Set(runProperties).size <= 1;
  return {doc, paragraph, plain, uniform};
}

export async function paragraphDigest(xml, Parser = DOMParser) {
  const {doc, paragraph} = paragraphXml(xml, Parser);
  const formatting = Array.from(doc.getElementsByTagNameNS(PKG, "part"))
    .map(p => [p.getAttributeNS(PKG, "name"), p])
    .filter(([name]) => /^\/word\/(styles(?:WithEffects)?\.xml|fontTable\.xml|numbering\.xml|theme\/[^/]+\.xml)$/.test(name))
    .map(([name, p]) => [name, Array.from(p.childNodes).map(canonicalNode).filter(x => x !== null)])
    .sort((a, b) => a[0].localeCompare(b[0]));
  return digest(JSON.stringify([canonicalNode(paragraph), formatting]));
}

export function planEdit(current, payload) {
  if (current.hash !== payload.expected_sha256) throw Error("Akapit zmienił się od odczytu. Odczytaj go ponownie.");
  if (typeof payload.text !== "string" || payload.text.length > 20000 || /[\x00-\x1f\x7f]/.test(payload.text)) throw Error("Tekst musi być pojedynczą linią do 20000 znaków.");
  const {plain, uniform} = paragraphXml(current.xml);
  const operation = payload.operation;
  if (["replace_text", "replace_paragraph"].includes(operation) && !plain)
    throw Error("Akapit zawiera obiekty, pola, zakładki, komentarze lub rewizje. Zapis tekstowy został zablokowany.");
  if (operation === "replace_paragraph" && !uniform)
    throw Error("Akapit ma różne formatowania. Użyj celowanego replace_text zamiast zastępować cały akapit.");
  let after = payload.text;
  if (operation === "replace_text") {
    const find = payload.find;
    if (!find || find.length > 255 || /[\x00-\x1f\x7f]/.test(find)) throw Error("Podaj unikalny tekst do zamiany (1–255 znaków).");
    searchText(find);  // the apply searches the ^-escaped form, which must also fit
    const first = current.text.indexOf(find);
    if (first < 0 || current.text.indexOf(find, first + 1) >= 0) throw Error("Tekst nie występuje dokładnie raz w akapicie.");
    after = current.text.slice(0, first) + payload.text + current.text.slice(first + find.length);
  } else if (operation === "set_style") {
    if (!STYLES.includes(payload.text)) throw Error("Nieobsługiwany styl wbudowany.");
    after = current.text;
  } else if (!["replace_paragraph", "insert_before", "insert_after"].includes(operation)) throw Error("Nieobsługiwana operacja.");
  return {operation, before: current.text, after, before_style: current.style, new_style: operation === "set_style" ? payload.text : null,
    formatting_note: operation === "replace_text" ? "Formatowanie poza dopasowanym tekstem zostaje zachowane; nowy tekst przejmuje format początku dopasowania." : null};
}

async function validateImageAsset(payload, decode) {
  if (!payload.image || !["image/png", "image/jpeg"].includes(payload.image.mime) ||
      typeof payload.image_base64 !== "string" || payload.image_base64.length > 2796204)
    throw Error("Obsługiwane są PNG/JPEG do 2 MiB.");
  const bytes = Uint8Array.from(atob(payload.image_base64), ch => ch.charCodeAt(0));
  if (!bytes.length || bytes.length > 2097152) throw Error("Obraz jest zbyt duży.");
  const hash = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)), b => b.toString(16).padStart(2, "0")).join("");
  if (hash !== payload.image.sha256) throw Error("Dane obrazu nie zgadzają się z podglądem.");
  const bitmap = await decode(new Blob([bytes], {type: payload.image.mime}));
  const width = bitmap.width, height = bitmap.height;
  bitmap.close();
  if (width !== payload.image.pixel_width || height !== payload.image.pixel_height ||
      width < 1 || height < 1 || width > 8192 || height > 8192 || width * height > 20000000)
    throw Error("Nieprawidłowe wymiary obrazu; zapisz go ponownie jako PNG/JPEG.");
  return {hash, width, height};
}

export async function planImage(current, payload, decode = createImageBitmap) {
  if (!current.hash || current.hash !== payload.expected_sha256) throw Error("Akapit zmienił się od odczytu. Odczytaj go ponownie.");
  paragraphXml(current.xml);
  if (payload.operation !== "insert_image" || !["before", "after"].includes(payload.position) ||
      !Number.isInteger(payload.width_pt) || payload.width_pt < 24 || payload.width_pt > 500 ||
      typeof payload.caption !== "string" || payload.caption.length > 500 ||
      typeof payload.alt_text !== "string" || payload.alt_text.length > 1000 ||
      /[\x00-\x1f\x7f]/.test(payload.caption + payload.alt_text)) throw Error("Nieprawidłowy plan wstawienia obrazu.");
  const {hash, width, height} = await validateImageAsset(payload, decode);
  const heightPt = payload.width_pt * height / width;
  if (heightPt > 700 || Math.abs(heightPt - payload.height_pt) > 0.01) throw Error("Obraz przekracza obsługiwany rozmiar strony.");
  return {operation: "insert_image", before: current.text, position: payload.position,
    image_name: payload.image_name, image_sha256: hash, mime: payload.image.mime,
    width_pt: payload.width_pt, height_pt: heightPt, caption: payload.caption, alt_text: payload.alt_text};
}

export async function planImageChange(current, payload, pictures, decode = createImageBitmap) {
  if (!current.hash || current.hash !== payload.expected_sha256) throw Error("Akapit zmienił się od odczytu. Odczytaj go ponownie.");
  paragraphXml(current.xml);
  if (!Array.isArray(pictures) || !Number.isInteger(payload.image_index) || payload.image_index < 0 ||
      payload.image_index >= pictures.length) throw Error("Wskazany obraz nie istnieje w tym akapicie. Odczytaj snapshot ponownie.");
  const target = pictures[payload.image_index];
  const previous = {width_pt: target.width, height_pt: target.height,
    alt_text: target.altTextDescription || "", alt_title: target.altTextTitle || ""};
  if (payload.operation === "delete_image") return {operation: payload.operation, before: current.text,
    image_index: payload.image_index, image_count_before: pictures.length, previous_image: previous};
  if (payload.operation !== "replace_image" || !Number.isInteger(payload.width_pt) ||
      (payload.width_pt !== 0 && (payload.width_pt < 24 || payload.width_pt > 500)) ||
      typeof payload.alt_text !== "string" || payload.alt_text.length > 1000 || /[\x00-\x1f\x7f]/.test(payload.alt_text))
    throw Error("Nieprawidłowy plan podmiany obrazu.");
  const {hash, width, height} = await validateImageAsset(payload, decode);
  const widthPt = payload.width_pt || target.width;
  const heightPt = widthPt * height / width;
  if (widthPt < 1 || widthPt > 500 || heightPt > 700) throw Error("Obraz przekracza obsługiwany rozmiar strony.");
  return {operation: payload.operation, before: current.text, image_index: payload.image_index,
    image_count_before: pictures.length, previous_image: previous, image_name: payload.image_name,
    image_sha256: hash, mime: payload.image.mime, width_pt: widthPt, height_pt: heightPt, alt_text: payload.alt_text};
}

const REVISION_NODES = new Set(["ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveToRangeStart", "pPrChange", "rPrChange"]);
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DML = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture";
// Objects inside one exported paragraph that block a structural edit or that
// the preview must disclose.
export function paragraphObjects(paragraph) {
  const all = Array.from(paragraph.getElementsByTagName("*"));
  const named = name => all.filter(n => n.namespaceURI === W && n.localName === name);
  const blockers = [];
  if (all.some(n => n.namespaceURI === W && REVISION_NODES.has(n.localName))) blockers.push("śledzone zmiany (rewizje)");
  if (named("commentRangeStart").length || named("commentRangeEnd").length || named("commentReference").length) blockers.push("komentarze");
  if (named("sdt").length) blockers.push("kontrolki zawartości");
  const chars = named("fldChar").map(n => n.getAttributeNS(W, "fldCharType"));
  if (!fieldNesting(paragraph).balanced) blockers.push("pole obejmujące inne akapity");
  // Floating drawings, VML shapes, OLE objects, charts and diagrams are not
  // inline pictures, so neither snapshot nor preview would disclose them.
  const embedded = all.filter(n => (n.namespaceURI === WP && n.localName === "anchor") ||
    (n.namespaceURI === W && ["object", "pict"].includes(n.localName)) ||
    (n.namespaceURI === DML && n.localName === "graphicData" && n.getAttribute("uri") !== PICTURE_URI)).length;
  if (embedded) blockers.push("obiekty pływające, wykresy lub obiekty osadzone");
  // The number mark that starts a footnote/endnote text.
  if (named("footnoteRef").length || named("endnoteRef").length) blockers.push("numer przypisu");
  return {blockers, fields: chars.filter(t => t === "begin").length + named("fldSimple").length,
    bookmarks: named("bookmarkStart").map(n => n.getAttributeNS(W, "name")).filter(n => n !== "_GoBack"),
    footnotes: named("footnoteReference").length + named("endnoteReference").length};
}

function requireFresh(current, payload) {
  if (!current.hash || current.hash !== payload.expected_sha256) throw Error("Akapit zmienił się od odczytu. Odczytaj go ponownie.");
}

// context: facts read from Word about the paragraph's surroundings.
export function planDeleteParagraph(current, payload, context) {
  requireFresh(current, payload);
  if (payload.text || payload.find) throw Error("delete_paragraph nie przyjmuje text ani find.");
  const objects = paragraphObjects(paragraphXml(current.xml).paragraph);
  if (objects.blockers.length) throw Error(`Akapit zawiera: ${objects.blockers.join(", ")}. Usunięcie zablokowane.`);
  if (context.isLastInBody) throw Error("To ostatni akapit swojego obszaru (dokumentu, komórki, nagłówka lub przypisu); Word nie pozwala go usunąć.");
  if (context.endsSection) throw Error("Akapit zawiera podział sekcji Worda; usunięcie zmieniłoby układ stron.");
  if (context.inContentControl) throw Error("Akapit leży w kontrolce zawartości; usunięcie zablokowane.");
  const level = context.tableNestingLevel || 0;
  if (context.previous && context.next && context.previous.tableNestingLevel > level && context.next.tableNestingLevel > level)
    throw Error("Akapit rozdziela dwie tabele; jego usunięcie scaliłoby je.");
  const disclosed = objects.fields || objects.bookmarks.length || context.pictures || objects.footnotes;
  return {operation: "delete_paragraph", before: current.text, before_style: current.style,
    inline_picture_count: context.pictures, field_count: objects.fields, bookmarks: objects.bookmarks,
    footnote_count: objects.footnotes, previous_text: context.previous ? context.previous.text.slice(0, 200) : null,
    next_text: context.next ? context.next.text.slice(0, 200) : null,
    note: disclosed ? "Usunięcie obejmuje pola, zakładki, przypisy i obrazy tego akapitu; odsyłacze do jego zakładek przestaną działać." : null};
}

export const FORMAT_TOKENS = ["bold", "italic", "underline", "no_bold", "no_italic", "no_underline"];
export function parseFormat(text) {
  const tokens = typeof text === "string" && text ? text.split(",") : [];
  if (!tokens.length) throw Error("Podaj formatowanie.");
  const format = {};
  for (const token of tokens) {
    if (!FORMAT_TOKENS.includes(token)) throw Error("Formatowanie: bold, italic, underline, no_bold, no_italic, no_underline (po przecinku).");
    const name = token.replace(/^no_/, "");
    if (name in format) throw Error("Sprzeczne lub powtórzone formatowanie.");
    format[name] = !token.startsWith("no_");
  }
  return format;
}

// Word interprets ^ codes even without wildcards; a literal caret is ^^.
export function searchText(find) {
  const escaped = typeof find === "string" ? find.replace(/\^/g, "^^") : "";
  if (!escaped || escaped.length > 255 || /[\x00-\x1f\x7f]/.test(find))
    throw Error("Podaj unikalny tekst do wyszukania (1–255 znaków, ^ liczy się podwójnie).");
  return escaped;
}

export function planFormat(current, payload) {
  requireFresh(current, payload);
  const format = parseFormat(payload.text);
  const objects = paragraphObjects(paragraphXml(current.xml).paragraph);
  if (objects.blockers.includes("śledzone zmiany (rewizje)")) throw Error("Akapit zawiera śledzone zmiany; formatowanie zablokowane.");
  if (objects.blockers.includes("kontrolki zawartości")) throw Error("Akapit zawiera kontrolkę zawartości; formatowanie zablokowane.");
  if (payload.find) {
    searchText(payload.find);
    const first = current.text.indexOf(payload.find);
    if (first < 0 || current.text.indexOf(payload.find, first + 1) >= 0) throw Error("Tekst nie występuje dokładnie raz w akapicie.");
  } else if (!current.text) throw Error("Akapit nie zawiera tekstu do sformatowania.");
  return {operation: "format_text", before: current.text, target: payload.find || current.text,
    whole_paragraph: !payload.find, format};
}

// table: facts about the anchor's table read through table.rows (never through
// the cell's own indexes, which fail for vertically merged cells).
export function planTableRow(current, payload, table) {
  requireFresh(current, payload);
  const cells = payload.cells;
  if (!Array.isArray(cells) || !cells.length || cells.length > 63 ||
      cells.some(c => typeof c !== "string" || c.length > 4000 || /[\x00-\x1f\x7f]/.test(c)))
    throw Error("Podaj 1–63 jednowierszowych tekstów komórek.");
  if (!["before", "after"].includes(payload.position)) throw Error("Nieprawidłowe położenie wiersza.");
  if (table.nestingLevel !== 1 || current.tableNestingLevel !== 1) throw Error("Obsługiwane są tylko tabele niezagnieżdżone.");
  if (table.merged) throw Error("Tabela ma scalone lub nieregularne komórki; dodawanie wiersza zablokowane.");
  const row = table.rows[table.rowIndex];
  if (!row) throw Error("Nie znaleziono wiersza odniesienia. Odczytaj snapshot ponownie.");
  if (row.isHeader) throw Error("Wiersz nagłówka nie może być wzorcem nowego wiersza; wskaż akapit w wierszu danych.");
  if (cells.length !== row.cellCount) throw Error(`Wiersz ma ${row.cellCount} komórek; podano ${cells.length}.`);
  return {operation: "insert_table_row", position: payload.position, reference_row_index: table.rowIndex,
    row_count_before: table.rows.length, reference_row: row.values, cells,
    formatting_note: "Nowy wiersz dziedziczy formatowanie wiersza odniesienia."};
}

const elements = node => Array.from(node.childNodes).filter(n => n.nodeType === 1);
const child = (node, name) => node ? elements(node).find(n => n.namespaceURI === W && n.localName === name) || null : null;
const valueOf = node => node ? node.getAttributeNS(W, "val") : null;
const packagePart = (doc, name) => Array.from(doc.getElementsByTagNameNS(PKG, "part")).find(p => p.getAttributeNS(PKG, "name") === name) || null;
// Paragraph properties that precede w:numPr in the schema sequence (CT_PPrBase).
const BEFORE_NUMPR = new Set(["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl"]);

// Numbering of an exported paragraph: direct w:numPr, completed from its style chain.
function paragraphNumbering(doc, pPr) {
  const direct = child(pPr, "numPr");
  let numId = valueOf(child(direct, "numId")), ilvl = valueOf(child(direct, "ilvl"));
  const stylesPart = packagePart(doc, "/word/styles.xml");
  const styles = stylesPart ? Array.from(stylesPart.getElementsByTagNameNS(W, "style")).filter(s => s.getAttributeNS(W, "type") === "paragraph") : [];
  const styleId = valueOf(child(pPr, "pStyle"));
  let style = styleId !== null ? styles.find(s => s.getAttributeNS(W, "styleId") === styleId) : styles.find(s => s.getAttributeNS(W, "default") === "1");
  for (let hops = 0; style && hops < 20 && (numId === null || ilvl === null); hops++) {
    const numPr = child(child(style, "pPr"), "numPr");
    if (numId === null) numId = valueOf(child(numPr, "numId"));
    if (ilvl === null) ilvl = valueOf(child(numPr, "ilvl"));
    const basedOn = valueOf(child(style, "basedOn"));
    style = basedOn !== null ? styles.find(s => s.getAttributeNS(W, "styleId") === basedOn) : null;
  }
  const numbering = packagePart(doc, "/word/numbering.xml")?.getElementsByTagNameNS(W, "numbering")[0] || null;
  const nums = numbering ? elements(numbering).filter(n => n.namespaceURI === W && n.localName === "num") : [];
  const num = nums.find(n => n.getAttributeNS(W, "numId") === numId) || null;
  const abstractId = valueOf(child(num, "abstractNumId"));
  const abstract = abstractId === null ? null : elements(numbering).find(n => n.namespaceURI === W && n.localName === "abstractNum" &&
    n.getAttributeNS(W, "abstractNumId") === abstractId) || null;
  return {numId, ilvl: ilvl ?? "0", numbering, nums, num, abstract, nsid: valueOf(child(abstract, "nsid")),
    siblings: nums.filter(n => abstractId !== null && valueOf(child(n, "abstractNumId")) === abstractId)};
}

// The start value a numbering instance sets for a level itself (lvlOverride), or null.
function instanceStart(num, level) {
  const override = elements(num).find(n => n.namespaceURI === W && n.localName === "lvlOverride" && n.getAttributeNS(W, "ilvl") === String(level));
  return valueOf(child(override, "startOverride")) ?? valueOf(child(child(override, "lvl"), "start"));
}

// The start value of a level: the instance's override, else the definition's w:start (0 when absent).
export function levelStart(xml, level, Parser = DOMParser, text = null) {
  const {numbering} = restartFacts(xml, level, Parser, text);
  const lvl = elements(numbering.abstract).find(n => n.namespaceURI === W && n.localName === "lvl" && n.getAttributeNS(W, "ilvl") === String(level));
  return Number(instanceStart(numbering.num, level) ?? valueOf(child(lvl, "start")) ?? "0");
}

// A numbering instance by content, its level overrides: numIds differ between exports.
const overridesOf = num => JSON.stringify(elements(num).filter(n => n.namespaceURI === W && n.localName === "lvlOverride")
  .map(n => JSON.stringify(canonicalNode(n))).sort());

function withoutNumPr(pPr) {
  if (!pPr) return null;
  const copy = pPr.cloneNode(true);
  for (const n of elements(copy).filter(n => n.namespaceURI === W && n.localName === "numPr")) copy.removeChild(n);
  return copy;
}

const blankParagraph = p => elements(p).every(n => n.namespaceURI === W && (["pPr", "proofErr"].includes(n.localName) ||
  (["bookmarkStart", "bookmarkEnd"].includes(n.localName) && !n.getAttributeNS(W, "name")?.replace("_GoBack", ""))));
const paragraphText = p => Array.from(p.getElementsByTagNameNS(W, "t")).map(t => t.textContent).join("");
const squash = text => text.replace(/[\s\u0000-\u001f\u00a0]/g, "");
// Counts only, never document text: body children, and per paragraph runs, text length and numPr.
const describeBody = body => elements(body).map(n => n.localName !== "p" ? n.localName :
  `p(r=${n.getElementsByTagNameNS(W, "r").length},t=${paragraphText(n).length}${child(child(n, "pPr"), "numPr") ? ",num" : ""})`).join(" ");

// The item's paragraph in its Whole-range export (Content exports omit w:pPr).
// Office.js exports extra empty paragraphs with it (live Word 2026-09-24), COM
// exactly one: the item is the only paragraph with content. text: the item's
// text from Office.js, checked when given.
function wholeParagraph(xml, Parser, text) {
  if (typeof xml !== "string" || /<!DOCTYPE|<!ENTITY/i.test(xml)) throw Error("Nieobsługiwany eksport XML.");
  if (xml.length > MAX_PARAGRAPH_XML_CHARS) throw Error(`Eksport akapitu ma ${xml.length} znaków; limit wynosi ${MAX_PARAGRAPH_XML_CHARS}.`);
  const doc = new Parser().parseFromString(xml, "application/xml");
  const bodies = doc.getElementsByTagNameNS(W, "body");
  if (doc.getElementsByTagName("parsererror").length || bodies.length !== 1) throw Error("Błędny XML.");
  const body = bodies[0];
  const direct = elements(body).filter(n => n.namespaceURI === W && n.localName === "p");
  if (elements(body).some(n => !(n.namespaceURI === W && ["p", "sectPr"].includes(n.localName))) ||
      body.getElementsByTagNameNS(W, "p").length !== direct.length)
    throw Error(`akapit leży w tabeli lub kontrolce zawartości albo zawiera obiekty z własnymi akapitami (eksport: ${describeBody(body)})`);
  const filled = direct.filter(p => !blankParagraph(p));
  const numbered = direct.filter(p => child(child(p, "pPr"), "numPr"));
  const paragraph = filled.length === 1 ? filled[0] : !filled.length && numbered.length === 1 ? numbered[0] : null;
  if (!paragraph) throw Error(`eksport akapitu ma nieoczekiwaną strukturę (${describeBody(body)})`);
  if (typeof text === "string" && squash(paragraphText(paragraph)) !== squash(text))
    throw Error(`eksport nie odpowiada tekstowi akapitu (${describeBody(body)})`);
  return {doc, paragraph};
}

// The item's paragraph and numbering from its Whole-range export.
function restartFacts(xml, level, Parser, text) {
  const {doc, paragraph} = wholeParagraph(xml, Parser, text);
  const pPr = child(paragraph, "pPr");
  if (pPr?.getElementsByTagNameNS(W, "sectPr").length) throw Error("akapit kończy sekcję Worda");
  if (pPr && Array.from(pPr.getElementsByTagName("*")).some(n => n.namespaceURI === W && (REVISION_NODES.has(n.localName) || n.localName === "numberingChange")))
    throw Error("właściwości akapitu zawierają śledzoną zmianę");
  const numbering = paragraphNumbering(doc, pPr);
  if (!numbering.numId || numbering.numId === "0" || !numbering.abstract) throw Error("eksport nie zawiera definicji numeracji akapitu");
  if (Number(numbering.ilvl) !== level) throw Error("poziom numeracji w OOXML różni się od poziomu listy");
  // Word merges the inserted numbering into the existing list definition by its nsid.
  if (!numbering.nsid) throw Error("definicja listy nie ma identyfikatora w:nsid");
  const content = elements(paragraph).filter(n => n !== pPr).map(canonicalNode).filter(x => x !== null);
  return {doc, pPr, numbering, content: JSON.stringify(content), properties: JSON.stringify(canonicalNode(withoutNumPr(pPr) ?? doc.createElementNS(W, "w:pPr")))};
}

// Word's own "Restart at 1" (verified through COM on Word M365): a new numbering
// instance (w:num) of the same list definition with startOverride=1 on the
// item's level, applied to this paragraph only. Later items keep their numbering
// instance and continue 2, 3, ... because Word counts per list definition.
// The package holds an empty paragraph carrying the item's own properties with
// the new w:numPr, then an empty sentinel: Office.js keeps the destination's
// paragraph mark for the last inserted paragraph (live Word 2026-09-24). Inserted
// at the end of the item's content, the item's runs end with the package's mark
// in a new paragraph; the item's former mark and ID stay on an empty paragraph
// after it, which the executor deletes.
// Paragraph.separateList cannot do this for the first paragraph of a numbering
// instance that continues an earlier one ("This command is not available.").
export function restartPackage(xml, level, Parser = DOMParser, Serializer = XMLSerializer, text = null) {
  const facts = restartFacts(xml, level, Parser, text);
  const {doc, pPr} = facts;
  const {numbering, nums, num, abstract} = facts.numbering;
  const fresh = String(Math.max(0, ...nums.map(n => Number(n.getAttributeNS(W, "numId"))).filter(Number.isInteger)) + 1);
  const instance = num.cloneNode(true);
  // Durable IDs are per instance; Word assigns one to the new instance.
  for (const a of Array.from(instance.attributes)) if (a.namespaceURI !== W && a.namespaceURI !== "http://www.w3.org/2000/xmlns/")
    instance.removeAttributeNS(a.namespaceURI, a.localName);
  instance.setAttributeNS(W, "w:numId", fresh);
  let override = elements(instance).find(n => n.namespaceURI === W && n.localName === "lvlOverride" && n.getAttributeNS(W, "ilvl") === String(level));
  if (!override) {
    override = doc.createElementNS(W, "w:lvlOverride");
    override.setAttributeNS(W, "w:ilvl", String(level));
    instance.appendChild(override);
  }
  for (const n of elements(override).filter(n => n.namespaceURI === W && n.localName === "startOverride")) override.removeChild(n);
  const start = doc.createElementNS(W, "w:startOverride");
  start.setAttributeNS(W, "w:val", "1");
  override.insertBefore(start, override.firstChild);
  // Word's import reuses an existing instance with the same overrides instead of
  // adding one (COM: a second restart in a list changed nothing; a restart before
  // an existing restart removed it). Keep the new instance distinct with a neutral
  // override of another level. A defined level restarts at its own start value
  // (0 when w:start is absent), which exports show. A level the definition does
  // not define (single-level lists such as List Number) has no items to number;
  // Word keeps such an override and matches on it, but exports omit it, so it gets
  // a unique start value (verified through COM and live Office.js 2026-09-24).
  const taken = new Set(facts.numbering.siblings.map(overridesOf));
  let hidden = false;
  for (let spare = 8; taken.has(overridesOf(instance)); spare--) {
    if (spare < 0) throw Error("definicja listy nie pozwala utworzyć odrębnego wystąpienia");
    const overridden = elements(instance).some(n => n.namespaceURI === W && n.localName === "lvlOverride" && n.getAttributeNS(W, "ilvl") === String(spare));
    const lvl = elements(abstract).find(n => n.namespaceURI === W && n.localName === "lvl" && n.getAttributeNS(W, "ilvl") === String(spare));
    if (spare === level || overridden) continue;
    hidden ||= !lvl;
    const neutral = doc.createElementNS(W, "w:lvlOverride");
    neutral.setAttributeNS(W, "w:ilvl", String(spare));
    const neutralStart = doc.createElementNS(W, "w:startOverride");
    neutralStart.setAttributeNS(W, "w:val", lvl ? valueOf(child(lvl, "start")) ?? "0" : String(1000 + crypto.getRandomValues(new Uint16Array(1))[0] % 9000));
    neutral.appendChild(neutralStart);
    instance.appendChild(neutral);
  }
  numbering.insertBefore(instance, nums[nums.length - 1].nextSibling);
  const properties = withoutNumPr(pPr) ?? doc.createElementNS(W, "w:pPr");
  const numPr = doc.createElementNS(W, "w:numPr");
  for (const [name, value] of [["ilvl", String(level)], ["numId", fresh]]) {
    const node = doc.createElementNS(W, `w:${name}`);
    node.setAttributeNS(W, "w:val", value);
    numPr.appendChild(node);
  }
  properties.insertBefore(numPr, elements(properties).find(n => !(n.namespaceURI === W && BEFORE_NUMPR.has(n.localName))) || null);
  const paragraph = doc.createElementNS(W, "w:p");
  paragraph.appendChild(properties);
  const body = doc.getElementsByTagNameNS(W, "body")[0];
  for (const n of elements(body)) if (!(n.namespaceURI === W && n.localName === "sectPr")) body.removeChild(n);
  body.insertBefore(doc.createElementNS(W, "w:p"), body.firstChild);  // sentinel
  body.insertBefore(paragraph, body.firstChild);
  return {xml: new Serializer().serializeToString(doc), num_id: facts.numbering.numId, nsid: facts.numbering.nsid, hidden,
    own_start: instanceStart(num, level),
    shape: JSON.stringify([facts.properties, facts.content, facts.numbering.numId, level, canonicalNode(num), canonicalNode(abstract), [...taken].sort()])};
}

// The item's own export after the write: same content and properties apart from
// w:numPr, which must now select a new instance of the same list definition that
// restarts at 1 on the item's level. hidden: the instance differs from the others
// only by an override exports omit, so they cannot tell it apart; the numbers
// checked by the executor still would.
export function restartOutcome(beforeXml, afterXml, level, Parser = DOMParser, text = null, hidden = false) {
  const before = restartFacts(beforeXml, level, Parser, text), after = restartFacts(afterXml, level, Parser, text);
  const problems = [];
  if (after.content !== before.content) problems.push("treść akapitu");
  if (after.properties !== before.properties) problems.push("właściwości akapitu");
  if (after.numbering.nsid !== before.numbering.nsid) problems.push("definicja listy akapitu");
  const override = elements(after.numbering.num).find(n => n.namespaceURI === W && n.localName === "lvlOverride" && n.getAttributeNS(W, "ilvl") === String(level));
  if (valueOf(child(override, "startOverride")) !== "1") problems.push("brak startOverride=1 w numeracji akapitu");
  if (!hidden && before.numbering.siblings.some(n => overridesOf(n) === overridesOf(after.numbering.num))) problems.push("Word użył istniejącego wystąpienia listy");
  return problems;
}

// First item of its level in the list (by document order).
export const firstOfLevel = list => list.items.findIndex(i => i.level === list.level) === list.index;

// The run a level restart renumbers, across Office.js lists: Office.js groups items
// by numbering instance while Word counts per list definition (live Word
// 2026-09-24), so later lists of the same definition shift too. xml: one Whole
// export from the item to the body's last list paragraph; all: every list item
// of the body in order ({id, level, value}); target: the item's index in all.
// Returns the ids of later items of the item's definition and level whose numbers
// continue its run (it ends at a shallower item or an existing restart), and
// whether a later paragraph uses the item's own numbering instance.
export const MAX_SPAN_XML_CHARS = 32 * 1024 * 1024;
export function restartRun(xml, all, target, level, Parser = DOMParser) {
  if (typeof xml !== "string" || /<!DOCTYPE|<!ENTITY/i.test(xml)) throw Error("Nieobsługiwany eksport XML.");
  if (xml.length > MAX_SPAN_XML_CHARS) throw Error(`eksport od elementu do końca list ma ${xml.length} znaków; limit wynosi ${MAX_SPAN_XML_CHARS}`);
  const doc = new Parser().parseFromString(xml, "application/xml");
  const bodies = doc.getElementsByTagNameNS(W, "body");
  if (doc.getElementsByTagName("parsererror").length || bodies.length !== 1) throw Error("Błędny XML.");
  const numbered = Array.from(bodies[0].getElementsByTagNameNS(W, "p")).map(p => paragraphNumbering(doc, child(p, "pPr")))
    .filter(n => n.numId && n.numId !== "0");
  if (numbered.length !== all.length - target)
    throw Error(`eksport od elementu do końca list nie odpowiada elementom list (${numbered.length}/${all.length - target})`);
  const {nsid, numId} = numbered[0];
  if (!nsid) throw Error("definicja listy nie ma identyfikatora w:nsid");
  const run = [];
  let previous = all[target].value, open = true, sharesInstance = false;
  for (let j = 1; j < numbered.length; j++) {
    const item = all[target + j], n = numbered[j];
    if (n.numId === numId) sharesInstance = true;
    if (!open || n.nsid !== nsid) continue;
    if (Number(n.ilvl) !== item.level) throw Error("poziomy w eksporcie nie odpowiadają poziomom list");
    if (item.level < level || typeof previous !== "number" || (item.level === level && item.value !== previous + 1)) { open = false; continue; }
    if (item.level !== level) continue;
    run.push(item.id);
    previous = item.value;
  }
  return {run, sharesInstance};
}

export const LIST_TYPES = ["bullet", "number"];
// list: facts about the paragraph's list gathered from Word by walking body
// paragraphs (List.paragraphs is unreliable on Win32). restart: for a level-0
// restart {package, run, sharesInstance} (restartPackage, restartRun), for the
// first item of a deeper level {levelStart}, or {reason} why it is unusable.
export function planList(current, payload, list, restart = null) {
  requireFresh(current, payload);
  if (payload.find) throw Error("Operacje na listach nie przyjmują find.");
  if (!list.isListItem) throw Error("Akapit nie jest elementem listy.");
  // Membership, guard and verification are read from the main body only.
  if (list.index < 0) throw Error("Operacje na listach obsługują listy w treści głównej dokumentu.");
  if (/^Heading[1-9]$/.test(current.style || "")) throw Error("Numeracja nagłówków wynika ze stylu; użyj set_style.");
  const levelItems = list.items.filter(i => i.level === list.level);
  const base = {operation: payload.operation, before: current.text, list_string: list.listString, level: list.level,
    level_type: list.levelTypes[list.level], list_item_count: list.items.length, level_item_count: levelItems.length,
    first_item: list.items.length ? list.items[0].text.slice(0, 200) : null,
    last_item: list.items.length ? list.items[list.items.length - 1].text.slice(0, 200) : null};
  if (payload.operation === "list_level") {
    if (!/^[0-8]$/.test(payload.text)) throw Error("Poziom listy: 0–8.");
    if (Number(payload.text) === list.level) throw Error("Element ma już ten poziom.");
    return {...base, new_level: Number(payload.text), scope: "Tylko ten element listy."};
  }
  if (payload.operation === "list_type") {
    if (!LIST_TYPES.includes(payload.text)) throw Error("Typ listy: bullet lub number.");
    if (base.level_type === (payload.text === "bullet" ? "Bullet" : "Number")) throw Error("Ten poziom listy ma już wskazany typ.");
    return {...base, new_type: payload.text, scope: `Wszystkie elementy tej listy na poziomie ${list.level} (${levelItems.length}).`};
  }
  if (payload.operation === "list_restart") {
    if (payload.text) throw Error("list_restart nie przyjmuje text.");
    if (base.level_type !== "Number") throw Error("Numerację od 1 można rozpocząć tylko na poziomie numerowanym.");
    // Without WordApiDesktop 1.3 the rendered value cannot be read to verify the restart.
    if (list.value === null) throw Error("Rozpoczęcie numeracji wymaga Worda z WordApiDesktop 1.3 (Microsoft 365 2507 lub nowszy).");
    if (list.value === 1) throw Error("Numeracja zaczyna się już od 1 na tym elemencie.");
    // Level 0 always restarts like Word's Restart at 1: the first item of an
    // Office.js list can still continue an earlier instance of its definition
    // (live Word 2026-09-24: "Plan step A" first of its list with number 8), and
    // setLevelStartingNumber cannot change that.
    const refuse = reason => Error(`Nie można bezpiecznie rozpocząć numeracji od 1 na tym elemencie: ${(reason || "brak eksportu akapitu").replace(/\.$/, "")}. ` +
      "Użyj w Wordzie polecenia „Uruchom ponownie od 1” (Restart at 1) albo uzgodnionego fallbacku COM.");
    if (list.level !== 0) {
      if (!firstOfLevel(list)) throw Error("Restart w środku listy jest obsługiwany dla elementów poziomu 0.");
      // setLevelStartingNumber changes the level's start value; a number that
      // continues an earlier item stays (live Word 2026-09-24).
      if (restart?.levelStart !== list.value) throw refuse(restart?.reason || "numer kontynuuje wcześniejszy element, a nie wynika z wartości początkowej poziomu");
      return {...base, method: "set_starting_number", current_value: list.value, scope: "Wartość początkowa tego poziomu listy."};
    }
    if (!restart?.package) throw refuse(restart?.reason);
    // An instance's own start value applies at its first paragraph: moving the
    // item off it would hand that value to the next item.
    if (restart.package.own_start !== null && restart.sharesInstance)
      throw refuse(`wystąpienie listy tego elementu ma własną wartość początkową (${restart.package.own_start}), którą Word przeniósłby na następny element`);
    const following = restart.run;
    return {...base, method: "start_override", current_value: list.value, renumbered_item_count: following.length + 1,
      scope: `Tylko ten akapit dostaje nowe wystąpienie tej samej definicji listy od 1, jak „Uruchom ponownie od 1” w Wordzie. ` +
        `Następne elementy liczą dalej: ${following.length} kolejnych elementów tego poziomu (także w dalszych listach tej samej definicji) zmieni numer; etykiety ich elementów podrzędnych ` +
        "z numerem nadrzędnym (np. 9.1) zmienią się razem z nimi. Treść akapitu się nie zmienia, ale dostaje on nowe paragraph_id."};
  }
  throw Error("Nieobsługiwana operacja listy.");
}
