// Fail closed: this profile supports one plain paragraph, one uniformly styled run.
export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const allowed = new Set(["p", "pPr", "pStyle", "spacing", "jc", "ind", "keepNext", "keepLines", "widowControl",
  "r", "rPr", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "u", "color", "sz", "szCs", "lang", "t"]);

export function inspectXml(xml, Parser = DOMParser) {
  if (typeof xml !== "string" || xml.length > 200000 || /<!DOCTYPE|<!ENTITY/i.test(xml)) throw Error("Nieobsługiwany XML.");
  const doc = new Parser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw Error("Błędny XML.");
  const bodies = doc.getElementsByTagNameNS(W, "body");
  if (bodies.length !== 1) throw Error("Nieobsługiwana struktura zakresu.");
  const body = bodies[0];
  const children = Array.from(body.childNodes).filter(n => n.nodeType === 1);
  // Word wraps even an inline selection in a synthetic section. Accept only a
  // trailing body-level sectPr; it is excluded from the replacement package.
  if (children.length === 2 && children[1].namespaceURI === W && children[1].localName === "sectPr") children.pop();
  if (children.length !== 1 || children[0].namespaceURI !== W || children[0].localName !== "p")
    throw Error("Dozwolony jest tylko jeden prosty akapit, bez sekcji, tabel i obiektów.");
  const paragraph = children[0];
  const all = [paragraph, ...Array.from(paragraph.getElementsByTagName("*"))];
  for (const node of all) {
    if (node.namespaceURI !== W || !allowed.has(node.localName))
      throw Error(`Zakres złożony (${node.localName}): zapis zablokowany.`);
  }
  if (paragraph.getElementsByTagNameNS(W, "r").length !== 1 || paragraph.getElementsByTagNameNS(W, "t").length !== 1)
    throw Error("Zakres musi zawierać jeden fragment tekstu o jednolitym formatowaniu.");
  const text = paragraph.getElementsByTagNameNS(W, "t")[0];
  if (text.parentNode.localName !== "r" || text.parentNode.parentNode !== paragraph)
    throw Error("Zagnieżdżona struktura tekstu jest niedozwolona.");
  // Only the exact selected paragraph is reinserted. Drop other Flat OPC parts,
  // avoiding document-wide style/relationship updates from a range package.
  return {doc, paragraph, text, oldText: text.textContent || ""};
}

export function replacementXml(xml, text, Parser = DOMParser, Serializer = XMLSerializer) {
  if (typeof text !== "string" || !text || text.length > 20000 || /[\x00-\x1f\x7f]/.test(text))
    throw Error("Nowy tekst musi być pojedynczą linią do 20000 znaków.");
  const selected = inspectXml(xml, Parser);
  selected.text.textContent = text;
  selected.text.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
  const paragraph = new Serializer().serializeToString(selected.paragraph);
  return `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="${W}"><w:body>${paragraph}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
}

export async function digest(text) {
  const bytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(bytes)).map(b => b.toString(16).padStart(2, "0")).join("");
}

// Word's Flat OPC export includes package metadata and regenerated revision IDs.
// Hash the selected structure and formatting dependencies, not transport bytes.
export function canonicalNode(node) {
  if (node.nodeType === 3 || node.nodeType === 4) {
    return node.parentNode?.namespaceURI === W && node.parentNode?.localName === "t"
      ? node.nodeValue : (node.nodeValue.trim() ? node.nodeValue : null);
  }
  if (node.nodeType !== 1) return null;
  // lastRenderedPageBreak is a pagination cache that Word refreshes on relayout.
  if (node.namespaceURI === W && ["rsid", "proofErr", "lastRenderedPageBreak"].includes(node.localName)) return null;
  const attributes = Array.from(node.attributes).filter(a =>
    a.namespaceURI !== "http://www.w3.org/2000/xmlns/" &&
    !(a.namespaceURI === W && /^rsid/.test(a.localName)) &&
    !(a.namespaceURI === "http://schemas.microsoft.com/office/word/2010/wordml" && ["paraId", "textId"].includes(a.localName)) &&
    // Drawing identifiers Word regenerates on every export of the same picture.
    !(a.namespaceURI === "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" && ["anchorId", "editId"].includes(a.localName))
  ).map(a => [a.namespaceURI || "", a.localName, a.value])
    .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
  return [node.namespaceURI || "", node.localName, attributes,
    Array.from(node.childNodes).map(canonicalNode).filter(n => n !== null)];
}

export async function structuralDigest(xml, Parser = DOMParser) {
  const {doc, paragraph} = inspectXml(xml, Parser);
  const dependencies = Array.from(doc.getElementsByTagNameNS(PKG, "part"))
    .map(part => [part.getAttributeNS(PKG, "name"), part])
    .filter(([name]) => /^\/word\/(styles(?:WithEffects)?\.xml|fontTable\.xml|theme\/[^/]+\.xml)$/.test(name))
    .map(([name, part]) => [name, Array.from(part.childNodes).map(canonicalNode).filter(n => n !== null)])
    .sort((a, b) => a[0].localeCompare(b[0]));
  return digest(JSON.stringify([canonicalNode(paragraph), dependencies]));
}
