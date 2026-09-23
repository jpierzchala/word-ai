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

export const LIST_TYPES = ["bullet", "number"];
// list: facts about the paragraph's list gathered from Word by walking body
// paragraphs (List.paragraphs is unreliable on Win32).
export function planList(current, payload, list) {
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
    const first = list.items.findIndex(i => i.level === list.level) === list.index;
    if (first && list.value === 1) throw Error("Numeracja tego poziomu zaczyna się już od 1 na tym elemencie.");
    if (!first && !list.canSeparate)
      throw Error("Rozpoczęcie numeracji w środku listy wymaga Worda z WordApiDesktop 1.4 (Microsoft 365 2508 lub nowszy).");
    if (!first && list.level !== 0) throw Error("Restart w środku listy jest obsługiwany dla elementów poziomu 0.");
    return {...base, method: first ? "set_starting_number" : "separate_list", current_value: list.value,
      scope: first ? "Wartość początkowa tego poziomu listy." : "Ten element i następne elementy listy tworzą nową listę od 1."};
  }
  throw Error("Nieobsługiwana operacja listy.");
}
