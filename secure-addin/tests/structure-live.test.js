// Structural operations through the real task pane executor against a small
// fake Word object model: preview -> guard -> apply -> verification.
import test from "node:test";
import assert from "node:assert/strict";
import {DOMParser, XMLSerializer} from "@xmldom/xmldom";
import {W} from "../safety.js";
import {paragraphDigest} from "../document-edit.js";

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const esc = s => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
const pXml = r => `<w:p><w:pPr><w:pStyle w:val="${r.style}"/></w:pPr>${r.fieldEnd ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>' : ""}${r.bookmark ? `<w:bookmarkStart w:id="9" w:name="${r.bookmark}"/><w:bookmarkEnd w:id="9"/>` : ""}${r.text ? `<w:r><w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>` : ""}</w:p>`;
const single = r => `<w:document xmlns:w="${W}"><w:body>${pXml(r)}</w:body></w:document>`;
// Word's list model as observed through COM: one counter per list definition
// (w:abstractNum, "abstract" here) on level 0; each paragraph selects a numbering
// instance (w:num); an instance's startOverride applies at its first paragraph,
// which also starts a new Word/Office.js list.
const numOf = r => String(r.list.num ?? r.list.id);
const abstractOf = r => String(r.list.abstract ?? r.list.id);
// Exports use integer ids and one nsid per definition, like Word.
const idOf = (state, map, key) => (state[map][key] ??= Object.keys(state[map]).length + 1);
function numberingXml(state, abstracts) {
  const levels = state.multiLevel ? 9 : 1;
  let xml = "";
  for (const a of abstracts) xml += `<w:abstractNum w:abstractNumId="${idOf(state, "abstractIds", a)}"><w:nsid w:val="0000AB${String(idOf(state, "abstractIds", a)).padStart(2, "0")}"/>` +
    Array.from({length: levels}, (_, l) => `<w:lvl w:ilvl="${l}"><w:start w:val="1"/></w:lvl>`).join("") + "</w:abstractNum>";
  const nums = [...new Map(state.rows.filter(r => r.list && abstracts.includes(abstractOf(r))).map(r => [numOf(r), abstractOf(r)])).entries()];
  for (const [num, a] of nums) xml += `<w:num w:numId="${idOf(state, "numIds", num)}" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" w16cid:durableId="7${idOf(state, "numIds", num)}"><w:abstractNumId w:val="${idOf(state, "abstractIds", a)}"/>` +
    Object.entries(state.numOverrides[num] || {}).filter(([l]) => Number(l) < levels).map(([l, v]) => `<w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="${v}"/></w:lvlOverride>`).join("") + "</w:num>";
  return `<pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData><w:numbering xmlns:w="${W}">${xml}</w:numbering></pkg:xmlData></pkg:part>`;
}
const listParagraphXml = (state, r) => !r.list ? pXml(r) :
  `<w:p><w:pPr><w:pStyle w:val="${r.style}"/><w:numPr><w:ilvl w:val="${r.list.level}"/><w:numId w:val="${idOf(state, "numIds", numOf(r))}"/></w:numPr>${r.sectionBreak ? "<w:sectPr/>" : ""}</w:pPr>${r.text ? `<w:r><w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>` : ""}</w:p>`;
// Whole export of paragraphs [first..last]: w:pPr with w:numPr and every instance of their definitions.
function wholeXml(state, rows) {
  const abstracts = [...new Set(rows.filter(r => r.list).map(abstractOf))];
  // Office.js adds an empty paragraph to Whole exports (live Word 2026-09-24).
  const trailing = state.trailingExportParagraph ? '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>' : "";
  return `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${W}"><w:body>${rows.map(r => listParagraphXml(state, r)).join("")}${trailing}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part>${numberingXml(state, abstracts)}</pkg:package>`;
}
function flat(records) {
  let body = "", i = 0;
  while (i < records.length) {
    if (records[i].tableNestingLevel) {
      let rows = "";
      while (i < records.length && records[i].tableNestingLevel) rows += `<w:tr><w:tc>${pXml(records[i++])}</w:tc></w:tr>`;
      body += `<w:tbl>${rows}</w:tbl><w:p/>`;  // Word exports an implicit paragraph after a closing table
    } else body += pXml(records[i++]);
  }
  return `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
}
const nul = () => ({isNullObject: true, load() {}});
const collection = items => ({items, load() {}, getLast() { return items[items.length - 1]; }});
let bootNumber = 0;

async function boot(rows, options = {}) {
  const state = {rows: rows.map(r => ({tableNestingLevel: 0, style: "Normal", ...r})), results: [], writes: 0, command: null,
    bookmarks: new Map(), next: 0, spanExports: 0, desktop: options.desktop ?? 1.4};
  // Bookmarks span paragraph records [start, end]; Word-like extension on insertion.
  for (const r of state.rows) if (r.bookmark) state.bookmarks.set(r.bookmark, {start: r, end: r});
  for (const [name, [from, to]] of Object.entries(options.bookmarks || {}))
    state.bookmarks.set(name, {start: state.rows.find(r => r.id === from), end: state.rows.find(r => r.id === to)});
  state.bookmarkText = name => { const m = state.bookmarks.get(name); return m && state.rows.slice(state.rows.indexOf(m.start), state.rows.indexOf(m.end) + 1).map(r => r.text).join("\r"); };
  state.separateCalls = 0;
  Object.assign(state, {multiLevel: Boolean(options.multiLevel), trailingExportParagraph: options.trailingExportParagraph ?? true, numOverrides: {}, numList: {}, abstractStart: {}, abstractIds: {}, numIds: {}, wholeExports: 0});
  for (const r of state.rows) if (r.list) {
    // Definition and instance stay when Word moves an item to another list.
    r.list.abstract ??= String(r.list.id);
    r.list.num = String(r.list.num ?? r.list.id);
    state.numList[numOf(r)] ??= r.list.id;
    if (r.list.override != null) state.numOverrides[numOf(r)] ??= {0: r.list.override};
  }
  let nextList = 100;
  const renumber = abstract => {
    const members = state.rows.filter(r => r.list && abstractOf(r) === abstract);
    if (!members.length) return;
    const firsts = new Set(), seen = new Set();
    for (const r of members) if (!seen.has(numOf(r))) { seen.add(numOf(r)); firsts.add(r); }
    let counter = (state.abstractStart[abstract] ?? members[0].list.value) - 1;
    for (const r of members) {
      const override = state.numOverrides[numOf(r)]?.[0];
      if (firsts.has(r) && override != null) counter = override - 1;
      // Office.js groups list items by numbering instance (live Word 2026-09-24).
      const list = state.numList[numOf(r)] ??= nextList++;
      state.lists[list] ??= [...state.lists[r.list.id]];
      r.list.id = list;
      if (r.list.level !== 0) continue;
      counter++;
      r.list.value = counter; r.list.string = `${counter}.`;
    }
  };
  // insertOoxml of the restart package at the end of an item's content (live
  // Word, Office.js): the last inserted paragraph keeps the item's mark, so a
  // package without a sentinel changes nothing; with one, the item's runs end
  // with the package's mark in a new paragraph and the former mark and ID stay on
  // an empty paragraph after it. The package's definition merges by nsid; an
  // instance with the same overrides of that definition is reused.
  const restartThroughOoxml = (rec, xml, location) => {
    if (location !== "End") throw Error(`unexpected insert location ${location}`);
    if (options.ooxmlRefused) {
      if (options.rereadFails) state.rereadFails = true;
      throw Object.assign(Error("Microsoft Word: This command is not available."), {debugInfo: {errorLocation: "Range.insertOoxml"}});
    }
    state.writes++;
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W, "body")[0].getElementsByTagNameNS(W, "p"));
    if (!paragraphs.length || paragraphs.some(p => p.getElementsByTagNameNS(W, "r").length)) throw Error("package must hold empty paragraphs");
    if (paragraphs.length === 1) return;  // the only paragraph merges into the item's mark: its properties are dropped
    const numId = paragraphs[0].getElementsByTagNameNS(W, "numId")[0].getAttributeNS(W, "val");
    const num = Array.from(doc.getElementsByTagNameNS(W, "num")).find(n => n.getAttributeNS(W, "numId") === numId);
    const nsid = Array.from(doc.getElementsByTagNameNS(W, "abstractNum")).find(a => a.getAttributeNS(W, "abstractNumId") ===
      num.getElementsByTagNameNS(W, "abstractNumId")[0].getAttributeNS(W, "val")).getElementsByTagNameNS(W, "nsid")[0].getAttributeNS(W, "val");
    if (nsid !== `0000AB${String(state.abstractIds[abstractOf(rec)]).padStart(2, "0")}`) throw Error("package lost the list definition");
    const overrides = Object.fromEntries(Array.from(num.getElementsByTagNameNS(W, "lvlOverride")).map(o =>
      [o.getAttributeNS(W, "ilvl"), Number(o.getElementsByTagNameNS(W, "startOverride")[0].getAttributeNS(W, "val"))]));
    // Failures in the middle of Word's batch.
    const partial = () => { throw Error("GeneralException"); };
    if (options.ooxmlPartial) { rec.list.value = overrides[0]; rec.list.string = `${overrides[0]}.`; partial(); }
    if (options.extraThrow) { state.rows.splice(index(rec) + 1, 0, {id: `new-${state.next++}`, text: "", style: "Normal", tableNestingLevel: 0}); partial(); }
    if (options.propertiesThrow) { rec.list.num = `p${state.next++}`; partial(); }  // another instance, same numbers
    if (options.numbersThrow) { state.rows.find(r => r.list && abstractOf(r) !== abstractOf(rec)).list.string = "9."; partial(); }
    if (options.ooxmlExtraParagraph) state.rows.splice(index(rec) + 1, 0, {id: `new-${state.next++}`, text: "", style: rec.style, tableNestingLevel: 0});
    const abstract = abstractOf(rec);
    const key = JSON.stringify(Object.entries(overrides).sort());
    const same = Object.keys(state.numOverrides).find(n => JSON.stringify(Object.entries(state.numOverrides[n]).sort()) === key &&
      state.rows.some(r => r.list && numOf(r) === n && abstractOf(r) === abstract));
    const target = same ?? `r${state.next++}`;
    if (!same) state.numOverrides[target] = overrides;
    const text = {id: `${rec.id}~${state.next++}`, text: rec.text, style: rec.style, tableNestingLevel: rec.tableNestingLevel, list: {...rec.list, num: target}};
    if (options.splitKeepsId) {  // not what Word does: the ID stays with the text
      state.rows.splice(index(rec) + 1, 0, {...text, text: "", list: {...rec.list}});
      rec.list = text.list;
      renumber(abstract);
      return;
    }
    state.rows.splice(index(rec), 0, text);
    rec.text = "";
    // Someone types into the empty paragraph before the executor removes it.
    if (options.typeIntoEmpty) state.onIdentity = () => { rec.text = "typed meanwhile"; state.onIdentity = null; };
    if (options.ooxmlNoContinue) {  // a separate definition: later items keep counting the old one
      text.list.abstract = `x${state.next++}`;
      renumber(text.list.abstract);
    }
    renumber(abstract);
  };
  const proxies = new WeakMap();
  const tableProxies = new WeakMap();
  const index = rec => state.rows.indexOf(rec);
  const styleBuiltIn = style => /^Heading [1-9]$/.test(style) ? style.replace(" ", "") : style === "Normal" ? "Normal" : "Other";
  const range = (kind, a, b) => ({
    kind, a, b, isNullObject: false, load() {},
    get text() {
      if (kind === "span") return state.inserted && options.untrackedRange ? "shifted" : state.rows.slice(index(a), index(b)).map(r => r.text).join("\r");
      if (kind === "bookmark") return state.rows.slice(index(a), index(b) + 1).map(r => r.text).join("\r");
      return a.text;
    },
    track() { return this; },
    getRange(location) { return range(location, location === "End" ? (b || a) : a); },
    get font() { return a.font ??= {bold: false, italic: false, underline: "None", load() {}}; },
    inlinePictures: collection([]), tables: collection([]), contentControls: collection([]),
    parentContentControlOrNullObject: nul(),
    get listFormat() { return {load() {}, get listValue() { return a.list?.value ?? null; }}; },
    getOoxml() {
      if (options.exportFails || state.rereadFails) throw Error("export");
      if (kind === "spanWhole") return {value: wholeXml(state, state.rows.slice(index(a), index(b) + 1))};
      if (kind === "Whole" && a.list && ++state.wholeExports === options.mutateOnWholeExport) a.list.num = `m${state.next++}`;
      if (kind === "span" && ++state.spanExports === 2 && options.editAfterExport) {
        const value = flat(state.rows.slice(index(a), index(b)));
        a.text += " (co-author)";  // typed while the apply export is being hashed
        return {value};
      }
      return {value: kind === "span" ? flat(state.rows.slice(index(a), index(b))) : kind === "table"
        ? `<w:document xmlns:w="${W}"><w:body><w:tbl>${a.table.rows.map(r => `<w:tr>${r.map(c => `<w:tc><w:p><w:r><w:t>${c}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`).join("")}</w:tbl></w:body></w:document>`
        : kind === "Whole" && a.list ? wholeXml(state, [a]) : single(a)};
    },
    expandTo(other) {
      if (kind === "Whole" && other.kind === "Whole") return range("spanWhole", a, other.a);
      return other.kind === "End" ? range("anchor", a, other.a) : range("span", a, other.a);
    },
    compareLocationWith(other) {
      const at = index(a), start = index(other.a), end = index(other.b);
      if (kind === "bookmark") return {value: at >= start && index(b) < end ? "Inside" : "OverlapsBefore"};
      if (other.kind === "bookmark") return {value: at === start ? "InsideStart" : start < at && at <= end ? "Inside" : at < start ? "Before" : "After"};
      return {value: at < start ? "Before" : at > end ? "After" : at === end ? "AdjacentAfter" : "InsideStart"};
    },
    getBookmarks() {
      const marks = [...state.bookmarks].filter(([, m]) => kind === "Start"
        ? index(m.start) <= index(a) && index(m.end) >= index(a)
        : index(m.end) >= index(a) && index(m.start) < index(b));
      return {value: marks.map(([n]) => n)};
    },
    search(escaped) {
      // Word reads ^^ as a literal caret; each match has its own font.
      const text = escaped.replace(/\^\^/g, "^");
      const match = {text, load() {}, font: {bold: false, italic: false, underline: "None", load() {}}};
      if (a.text.includes(text)) (a.matches ??= []).push(match);
      return {items: a.text.includes(text) ? [match] : [], load() {}};
    },
    insertOoxml(xml, location) {
      if (kind === "Content") return restartThroughOoxml(a, xml, location);
      state.writes++;
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const body = doc.getElementsByTagNameNS(W, "body")[0];
      const inserted = [];
      for (const node of Array.from(body.childNodes).filter(n => n.nodeType === 1)) {
        const records = node.localName === "tbl" ? Array.from(node.getElementsByTagNameNS(W, "p")).map(p => [p, 1]) : [[node, 0]];
        for (const [p, level] of records) {
          const style = p.getElementsByTagNameNS(W, "pStyle")[0]?.getAttributeNS(W, "val") ?? "Normal";
          const rec = {id: `new-${state.next++}`, text: Array.from(p.getElementsByTagNameNS(W, "t")).map(t => t.textContent).join(""),
            style, tableNestingLevel: level};
          const mark = p.getElementsByTagNameNS(W, "bookmarkStart")[0];
          // Word drops inserted bookmarks whose names still exist.
          if (mark && !state.bookmarks.has(mark.getAttributeNS(W, "name"))) state.bookmarks.set(mark.getAttributeNS(W, "name"), {start: rec, end: rec});
          inserted.push(rec);
        }
      }
      // Like Word, a bookmark starting at the insertion point grows over the copy.
      for (const m of state.bookmarks.values()) if (m.start === a) m.start = inserted[0];
      state.rows.splice(index(a), 0, ...inserted);
      state.inserted = true;
    },
    delete() {
      state.writes++;
      if (options.revokeOnWrite) element("disconnect").listeners.click();
      const removed = state.rows.splice(index(a), index(b) - index(a));
      for (const [name, m] of state.bookmarks) if (removed.includes(m.start) || removed.includes(m.end)) state.bookmarks.delete(name);
    },
    insertBookmark(name) { state.bookmarks.set(name, {start: a, end: kind === "bookmark" || kind === "anchor" ? b : a}); },
  });
  const paragraph = rec => {
    if (proxies.has(rec)) return proxies.get(rec);
    const proxy = {
      isNullObject: false, load() {},
      get uniqueLocalId() { return rec.id; }, get text() { return rec.text; }, get style() { return rec.style; },
      get styleBuiltIn() { return styleBuiltIn(rec.style); }, get outlineLevel() { return 10; },
      get tableNestingLevel() { return rec.tableNestingLevel; }, get isListItem() { return Boolean(rec.list); },
      getRange: location => range(location, rec),
      get parentBody() { return {type: "MainDoc", load() {}, get paragraphs() { return collection(state.rows.map(paragraph)); }}; },
      parentContentControlOrNullObject: nul(),
      get listItemOrNullObject() {
        if (!rec.list) return nul();
        return {isNullObject: false, load() {}, get level() { return rec.list.level; }, set level(v) { state.writes++; rec.list.level = v; },
          get listString() { return rec.list.string; }};
      },
      get listOrNullObject() {
        if (!rec.list) return nul();
        const members = () => state.rows.filter(r => r.list?.id === rec.list.id);
        return {isNullObject: false, load() {}, get id() { return rec.list.id; }, get levelTypes() { return state.lists[rec.list.id]; },
          setLevelBullet(level) { state.writes++; state.lists[rec.list.id][level] = "Bullet"; },
          setLevelNumbering(level) { state.writes++; state.lists[rec.list.id][level] = "Number"; },
          setLevelStartingNumber(level, start) { state.writes++; state.abstractStart[abstractOf(rec)] = start; renumber(abstractOf(rec)); }};
      },
      // The add-in no longer calls separateList: Word refuses it on the first
      // paragraph of a numbering instance ("This command is not available.").
      separateList() {
        state.separateCalls++;
        throw Object.assign(Error("Microsoft Word: This command is not available."), {debugInfo: {errorLocation: "Paragraph.separateList"}});
      },
       get parentTableOrNullObject() {
         if (!rec.table) return nul();
         const t = rec.table;
         const cached = tableProxies.get(t);
         if (options.cachedTableProxyPerRun && cached?.runNumber === state.runNumber) return cached.proxy;
         const snapshot = options.staleTableProxy ? {header: t.header, rows: t.rows.map(r => [...r])} : null;
         const tableProxy = {isNullObject: false, load() {}, nestingLevel: 1, isUniform: true,
           get headerRowCount() { return snapshot ? snapshot.header : t.header; },
           get rowCount() { return snapshot ? snapshot.rows.length : t.rows.length; },
           get values() { return snapshot ? snapshot.rows : options.sharedTableValues ? t.rows : t.rows.map(r => [...r]); },
           getRange: () => range("table", rec),
           get rows() { return collection(t.rows.map((r, i) => ({cellCount: r.length, isHeader: i < t.header,
             insertRows(where, count, values) { state.writes++; t.rows.splice(where === "Before" ? i : i + 1, 0, [...values[0]]); }}))); }};
         if (options.cachedTableProxyPerRun) tableProxies.set(t, {runNumber: state.runNumber, proxy: tableProxy});
         return tableProxy;
      },
      get parentTableCell() { return {load() {}, rowIndex: rec.row}; },
      delete() {
        state.writes++;
        if (options.brokenDelete) return;
        state.rows.splice(index(rec), 1);
        if (rec.list) renumber(abstractOf(rec));  // Word renumbers the list at once
      },
    };
    proxies.set(rec, proxy);
    return proxy;
  };
  state.lists = options.lists || {};
  const elements = new Map();
  const element = id => elements.get(id) || (elements.set(id, {listeners: {}, addEventListener(n, f) { this.listeners[n] = f; }}), elements.get(id));
  globalThis.DOMParser = DOMParser; globalThis.XMLSerializer = XMLSerializer;
  globalThis.document = {getElementById: element};
  globalThis.window = {setInterval: fn => { state.tick = fn; }};
  globalThis.Office = {AsyncResultStatus: {Succeeded: "ok"}, HostType: {Word: "Word"}, onReady: fn => fn({host: "Word"}),
    context: {requirements: {isSetSupported: (name, version) => name === "WordApi" || (name === "WordApiDesktop" && Number(version) <= state.desktop)},
      document: {getFilePropertiesAsync: cb => { state.onIdentity?.(); cb({status: "ok", value: {url: "https://test/doc.docx"}}); }}}};
  const ctx = {sync: async () => {}, document: {
    load() {}, changeTrackingMode: "Off",
    getParagraphByUniqueLocalId: id => paragraph(state.rows.find(r => r.id === id)),
    get body() { return {paragraphs: collection(state.rows.map(paragraph)), inlinePictures: collection([]), tables: collection([]),
      fields: collection([]), footnotes: collection([]), endnotes: collection([]),
      get lists() { return collection(Object.entries(state.lists).map(([id, levelTypes]) => ({id: Number(id), levelTypes}))); }}; },
    get sections() { return collection([{body: {paragraphs: collection(state.rows.map(paragraph))}}]); },
    getBookmarkRangeOrNullObject: name => state.bookmarks.has(name) ? range("bookmark", state.bookmarks.get(name).start, state.bookmarks.get(name).end) : nul(),
    deleteBookmark: name => { state.bookmarks.delete(name); },
  }};
  globalThis.Word = {run: fn => { state.runNumber = (state.runNumber || 0) + 1; return fn(ctx); }};
  globalThis.fetch = async (url, request) => {
    const body = JSON.parse(request.body); let data = {};
    if (url === "/office/pair") data = {pairing_id: "p", secret: "s"};
    if (url === "/office/pair-status") data = {status: "connected", session_id: "s1", capability: "c1"};
    if (url === "/office/poll") { data = {command: state.command}; state.command = null; }
    if (url === "/office/result") state.results.push(body.result);
    return {ok: true, status: 200, json: async () => data};
  };
  await import(`../taskpane.js?structure=${++bootNumber}`);
  await element("connect").listeners.click(); await state.tick();
  state.hash = async id => paragraphDigest(single(state.rows.find(r => r.id === id)));
  // Preview through the executor, then apply with the returned guard as the bridge would.
  state.run = async payload => {
    state.command = {command_id: "preview", type: "preview", payload}; await state.tick();
    const preview = state.results.at(-1);
    if (!preview.ok) return {preview};
    state.command = {command_id: "apply", type: "apply", payload: {...payload, guard_sha256: preview.guard_sha256}}; await state.tick();
    return {preview, applied: state.results.at(-1)};
  };
  return state;
}

const outline = () => [
  {id: "intro", text: "Intro"}, {id: "h-a", text: "A", style: "Heading 1"}, {id: "a1", text: "a1"},
  {id: "h-b", text: "B", style: "Heading 1", bookmark: "_Toc1"}, {id: "h-b1", text: "B.1", style: "Heading 2"},
  {id: "b1", text: "Lead-in: text"}, {id: "cell1", text: "x1", tableNestingLevel: 1}, {id: "cell2", text: "x2", tableNestingLevel: 1},
  {id: "h-c", text: "C", style: "Heading 1"}, {id: "c1", text: "c1"}, {id: "final", text: ""}];

test("move_section previews the range and moves it, removing seam paragraphs and keeping bookmarks", async () => {
  const s = await boot(outline());
  const payload = {operation: "move_section", paragraph_id: "h-b", expected_sha256: await s.hash("h-b"),
    target_paragraph_id: "h-a", target_expected_sha256: await s.hash("h-a"), position: "before"};
  s.command = {command_id: "preview", type: "preview", payload}; await s.tick();
  const preview = s.results.at(-1);
  assert.equal(preview.ok, true, JSON.stringify(preview));
  assert.equal(preview.paragraph_count, 5); assert.equal(preview.subsection_count, 1);
  assert.equal(preview.last_paragraph.in_table, true); assert.equal(preview.section_ends_before, "C");
  assert.equal(preview.target_ooxml_sha256, payload.target_expected_sha256);
  assert.equal(s.writes, 0, "preview never writes");
  s.command = {command_id: "apply", type: "apply", payload: {...payload, guard_sha256: preview.guard_sha256}}; await s.tick();
  const applied = s.results.at(-1);
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.deepEqual(s.rows.map(r => r.text), ["Intro", "B", "B.1", "Lead-in: text", "x1", "x2", "A", "a1", "C", "c1", ""]);
  assert.equal(applied.seam_paragraphs_removed, 2, "implicit post-table paragraph and sentinel");
  assert.equal(s.bookmarkText("_Toc1"), "B", "bookmark restored on the moved heading");
  assert.ok(![...s.bookmarks.keys()].some(n => n.startsWith("WAI")));
});

test("move_section apply refuses a stale structure without writing", async () => {
  const s = await boot(outline());
  const payload = {operation: "move_section", paragraph_id: "h-b", expected_sha256: await s.hash("h-b"),
    target_paragraph_id: "c1", target_expected_sha256: await s.hash("c1"), position: "after"};
  s.command = {command_id: "preview", type: "preview", payload}; await s.tick();
  const preview = s.results.at(-1); assert.equal(preview.ok, true, JSON.stringify(preview));
  s.rows.splice(5, 0, {id: "late", text: "added by a co-author", style: "Normal", tableNestingLevel: 0});
  s.command = {command_id: "apply", type: "apply", payload: {...payload, guard_sha256: preview.guard_sha256}}; await s.tick();
  assert.equal(s.results.at(-1).ok, false); assert.equal(s.results.at(-1).submitted, false);
  assert.match(s.results.at(-1).error, /zmieniła się od podglądu/); assert.equal(s.writes, 0);
});

test("move_section to the end of the document and a target inside the section", async () => {
  const s = await boot(outline());
  const down = {operation: "move_section", paragraph_id: "h-a", expected_sha256: await s.hash("h-a"),
    target_paragraph_id: "c1", target_expected_sha256: await s.hash("c1"), position: "after"};
  const {applied} = await s.run(down);
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.deepEqual(s.rows.map(r => r.text), ["Intro", "B", "B.1", "Lead-in: text", "x1", "x2", "C", "c1", "A", "a1", ""]);
  const inner = {...down, paragraph_id: "h-b", expected_sha256: await s.hash("h-b"), target_paragraph_id: "b1", target_expected_sha256: await s.hash("b1")};
  const {preview} = await s.run(inner);
  assert.equal(preview.ok, false); assert.match(preview.error, /wewnątrz/);
});

test("delete_paragraph removes exactly one paragraph and verifies its neighbours", async () => {
  const s = await boot(outline());
  const {preview, applied} = await s.run({operation: "delete_paragraph", paragraph_id: "a1", expected_sha256: await s.hash("a1"), text: "", find: ""});
  assert.equal(preview.previous_text, "A"); assert.equal(preview.next_text, "B");
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.deepEqual(s.rows.map(r => r.id).slice(0, 3), ["intro", "h-a", "h-b"]); assert.equal(s.writes, 1);
  const last = await s.run({operation: "delete_paragraph", paragraph_id: "final", expected_sha256: await s.hash("final"), text: "", find: ""});
  assert.equal(last.preview.ok, false); assert.match(last.preview.error, /ostatni akapit/);
});

test("format_text bolds a unique lead-in and verifies the font", async () => {
  const s = await boot(outline());
  const {preview, applied} = await s.run({operation: "format_text", paragraph_id: "b1", expected_sha256: await s.hash("b1"), text: "bold", find: "Lead-in:"});
  assert.deepEqual(preview.before_format, {bold: false, italic: false, underline: "None"});
  assert.equal(applied.ok, true, JSON.stringify(applied)); assert.equal(applied.format.bold, true);
  const record = s.rows.find(r => r.id === "b1");
  assert.equal(record.matches.at(-1).font.bold, true, "the lead-in is bold");
  assert.equal(record.font?.bold ?? false, false, "the rest of the paragraph is not");
  const caret = await boot([{id: "m", text: "E = mc^2 holds"}, {id: "end", text: ""}]);
  const math = await caret.run({operation: "format_text", paragraph_id: "m", expected_sha256: await caret.hash("m"), text: "italic", find: "mc^2"});
  assert.equal(math.applied.ok, true, JSON.stringify(math.applied));
});

test("list_level and list_restart change only the addressed item", async () => {
  const rows = [{id: "i1", text: "one", list: {id: 1, level: 0, string: "4.", value: 4}},
    {id: "i2", text: "two", list: {id: 1, level: 0, string: "5.", value: 5}}, {id: "i3", text: "three", list: {id: 1, level: 0, string: "6.", value: 6}},
    {id: "end", text: ""}];
  const s = await boot(rows, {lists: {1: ["Number", "Number"]}});
  const level = await s.run({operation: "list_level", paragraph_id: "i3", expected_sha256: await s.hash("i3"), text: "1", find: ""});
  assert.equal(level.applied.ok, true, JSON.stringify(level.applied)); assert.equal(s.rows[2].list.level, 1);
  const restart = await s.run({operation: "list_restart", paragraph_id: "i2", expected_sha256: await s.hash("i2"), text: "", find: ""});
  assert.equal(restart.preview.method, "start_override");
  assert.equal(restart.applied.ok, true, JSON.stringify(restart.applied)); assert.equal(restart.applied.list_value, 1);
  assert.equal(s.rows[0].list.string, "4.", "earlier item keeps its number");
  assert.equal(s.rows[2].list.value, 6, "a deeper item keeps its number");
  assert.equal(s.separateCalls, 0);
});

// The same layout as live Word showed it on the synthetic document: Office.js
// groups by numbering instance, so 9.1's first item is the first item of its list
// although it reads 8.
const perInstance = () => continued().map(r => r.list?.num === 2 ? {...r, list: {...r.list, id: 7, abstract: "1"}} : r);

// The layout observed 2026-09-23: 1-7, a long gap with headings and a table, then
// 9.1 (8-11) and 9.2 (12-14). 9.1 starts its own numbering instance (num 2) of
// the same list definition, so Word refuses Paragraph.separateList there.
const continued = () => [
  {id: "h7", text: "7.1 Functional", style: "Heading 2"},
  ...[1, 2, 3, 4, 5, 6, 7].map(n => ({id: `a${n}`, text: `first ${n}`, style: "List Number", list: {id: 1, level: 0, string: `${n}.`, value: n}})),
  {id: "h8", text: "8 Design", style: "Heading 1"}, {id: "gap", text: "Gap paragraph."},
  {id: "cell", text: "cell", tableNestingLevel: 1}, {id: "h9", text: "9.1 Plan", style: "Heading 2"},
  ...[8, 9, 10, 11].map(n => ({id: `b${n}`, text: `second ${n}`, style: "List Number", list: {id: 1, level: 0, string: `${n}.`, value: n, num: 2}})),
  {id: "h92", text: "9.2 Risks", style: "Heading 2"},
  ...[12, 13, 14].map(n => ({id: `c${n}`, text: `third ${n}`, style: "List Number", list: {id: 1, level: 0, string: `${n}.`, value: n, num: 2}})),
  {id: "o1", text: "other one", list: {id: 5, level: 0, string: "1.", value: 1}}, {id: "o2", text: "other two", list: {id: 5, level: 0, string: "2.", value: 2}},
  {id: "end", text: ""}];
const LISTS = {lists: {1: ["Number"], 5: ["Number"]}};
const restartAt = async (s, id) => s.run({operation: "list_restart", paragraph_id: id, expected_sha256: await s.hash(id), text: "", find: ""});
const numbers = (s, pattern) => s.rows.filter(r => pattern.test(r.id) && r.list).map(r => r.list.value);

test("list_restart on the first item of a continued block restarts at 1 like Word's Restart at 1, without separateList", async () => {
  const s = await boot(continued(), {...LISTS, desktop: 1.3});
  const {preview, applied} = await restartAt(s, "b8");
  assert.equal(preview.ok, true, JSON.stringify(preview));
  assert.equal(preview.method, "start_override"); assert.equal(preview.current_value, 8);
  assert.equal(preview.renumbered_item_count, 7); assert.equal(preview.list_item_count, 14);
  assert.equal(applied.ok, true, JSON.stringify(applied)); assert.equal(applied.method, "start_override");
  assert.deepEqual(numbers(s, /^a\d/), [1, 2, 3, 4, 5, 6, 7]);
  assert.deepEqual(numbers(s, /^[bc]\d/), [1, 2, 3, 4, 5, 6, 7]);
  assert.deepEqual(numbers(s, /^o\d/), [1, 2], "the other list is untouched");
  // Office.js moves the item's text into a new paragraph; its former ID and mark go with the removed empty paragraph.
  assert.equal(applied.previous_paragraph_id, "b8"); assert.equal(s.rows.some(r => r.id === "b8"), false);
  const item = s.rows.find(r => r.id === applied.paragraph_id);
  assert.equal(item.text, "second 8"); assert.equal(s.rows.length, continued().length);
  assert.notEqual(item.list.id, 1, "the restarted item has its own instance");
  assert.equal(s.separateCalls, 0); assert.equal(s.writes, 2, "insertion and removal of the empty paragraph");
});

test("the first item of an Office.js list that continues an earlier instance restarts through a new instance", async () => {
  const s = await boot(perInstance(), {lists: {1: ["Number"], 5: ["Number"], 7: ["Number"]}});
  const {preview, applied} = await restartAt(s, "b8");
  assert.equal(preview.ok, true, JSON.stringify(preview));
  assert.equal(preview.list_item_count, 7); assert.equal(preview.first_item, "second 8");
  assert.equal(preview.method, "start_override", "setLevelStartingNumber could not change a continued number");
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.deepEqual(numbers(s, /^a\d/), [1, 2, 3, 4, 5, 6, 7]);
  assert.deepEqual(numbers(s, /^[bc]\d/), [1, 2, 3, 4, 5, 6, 7]);
});

test("a restart renumbers later Office.js lists of the same definition and says so in the preview", async () => {
  const s = await boot(perInstance(), {lists: {1: ["Number"], 5: ["Number"], 7: ["Number"]}});
  const {preview, applied} = await restartAt(s, "a3");
  assert.equal(preview.list_item_count, 7, "its own Office.js list"); assert.equal(preview.renumbered_item_count, 12);
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.deepEqual(numbers(s, /^a\d/), [1, 2, 1, 2, 3, 4, 5]);
  assert.deepEqual(numbers(s, /^[bc]\d/), [6, 7, 8, 9, 10, 11, 12]);
  assert.deepEqual(numbers(s, /^o\d/), [1, 2]);
  const order = await boot(perInstance(), {lists: {1: ["Number"], 5: ["Number"], 7: ["Number"]}});
  for (const id of ["b8", "c12", "a3"]) assert.equal((await restartAt(order, id)).applied.ok, true, id);
  assert.deepEqual(numbers(order, /^[abc]\d/), [1, 2, 1, 2, 3, 4, 5, 1, 2, 3, 4, 1, 2, 3]);
});

test("the first item of an instance with its own start value is refused: Word would hand that value on", async () => {
  const rows = perInstance().map(r => r.list?.num === 2 ? {...r, list: {...r.list, value: r.list.value - 3, string: `${r.list.value - 3}.`, override: r.id === "b8" ? 5 : undefined}} : r);
  const s = await boot(rows, {lists: {1: ["Number"], 5: ["Number"], 7: ["Number"]}});
  const {preview} = await restartAt(s, "b8");
  assert.equal(preview.ok, false); assert.match(preview.error, /własną wartość początkową \(5\).*Restart at 1/);
  assert.equal(s.writes, 0);
});

test("a deeper-level number that continues an earlier item is refused instead of an ineffective setLevelStartingNumber", async () => {
  const rows = [{id: "p1", text: "parent", list: {id: 1, level: 0, string: "1.", value: 1}},
    {id: "s1", text: "sub one", list: {id: 8, abstract: "1", num: 8, level: 1, string: "c.", value: 3}},
    {id: "s2", text: "sub two", list: {id: 8, abstract: "1", num: 8, level: 1, string: "d.", value: 4}}, {id: "end", text: ""}];
  const s = await boot(rows, {lists: {1: ["Number", "Number"], 8: ["Number", "Number"]}, multiLevel: true});
  const {preview} = await restartAt(s, "s1");
  assert.equal(preview.ok, false); assert.match(preview.error, /kontynuuje wcześniejszy element/); assert.equal(s.writes, 0);
});

for (const [mode, problem, writes] of [["typeIntoEmpty", /pusty akapit zmienił się przed usunięciem/, 2], ["splitKeepsId", /układ akapitów po wstawieniu/, 1],
  ["brokenDelete", /układ akapitów/, 2]]) {
  test(`the removal of the empty paragraph never hides a surprise: ${mode}`, async () => {
    const s = await boot(continued(), {...LISTS, [mode]: true});
    const {applied} = await restartAt(s, "b8");
    assert.equal(applied.ok, false); assert.equal(applied.submitted, true); assert.equal(applied.unchanged, false);
    assert.match(applied.error, problem); assert.equal(s.writes, writes);
    assert.ok(s.rows.some(r => r.text === "second 8"), "the item's text is never removed");
  });
}

// Single-level definitions (List Number) cannot tell a second restart instance
// apart by a visible override; a unique override of an undefined level does.
for (const [order, expected] of [[["b8", "c12"], [1, 2, 3, 4, 1, 2, 3]], [["c12", "b8"], [1, 2, 3, 4, 1, 2, 3]],
  [["a3", "a5"], null], [["a4", "a2"], null], [["b8", "c12", "b10"], [1, 2, 1, 2, 1, 2, 3]]]) {
  test(`restarts ${order.join(" then ")} in a single-level list each get their own instance`, async () => {
    const s = await boot(continued(), LISTS);
    for (const id of order) {
      const {preview, applied} = await restartAt(s, id);
      assert.equal(preview.method, "start_override", id);
      assert.equal(applied.ok, true, `${id}: ${JSON.stringify(applied)}`);
    }
    const values = numbers(s, /^[abc]\d/);
    if (expected) assert.deepEqual(numbers(s, /^[bc]\d/), expected);
    else if (order[0] === "a3") assert.deepEqual(values.slice(0, 7), [1, 2, 1, 2, 1, 2, 3], "a5 restarts, a3's restart stays");
    else assert.deepEqual(values.slice(0, 7), [1, 1, 2, 1, 2, 3, 4], "a2 restarts before a4's restart, which stays");
    assert.deepEqual(numbers(s, /^o\d/), [1, 2]);
    assert.equal(s.separateCalls, 0);
  });
}

test("multi-level lists keep every restart a new instance through a neutral deeper-level override", async () => {
  const s = await boot(continued(), {...LISTS, multiLevel: true});
  for (const id of ["b8", "c12", "b10"]) {
    const {preview, applied} = await restartAt(s, id);
    assert.equal(preview.method, "start_override", id);
    assert.equal(applied.ok, true, `${id}: ${JSON.stringify(applied)}`);
  }
  assert.deepEqual(numbers(s, /^[bc]\d/), [1, 2, 1, 2, 1, 2, 3]);
  assert.deepEqual(numbers(s, /^a\d/), [1, 2, 3, 4, 5, 6, 7]);
  assert.equal(s.separateCalls, 0);
});

test("a Word refusal whose fresh read matches the preview is reported as unchanged (bridge: failed, not unknown)", async () => {
  const s = await boot(continued(), {...LISTS, ooxmlRefused: true});
  const {preview, applied} = await restartAt(s, "b8");
  assert.equal(preview.ok, true, JSON.stringify(preview));
  assert.equal(applied.ok, false); assert.equal(applied.submitted, true); assert.equal(applied.unchanged, true);
  assert.match(applied.error, /This command is not available\..*nie zmienił/); assert.equal(applied.api, "Range.insertOoxml");
  assert.equal(s.rows.find(r => r.id === "b8").list.value, 8);
});

// Each case changes the document in a way only one of the post-error checks can see.
for (const [mode, seen] of [["ooxmlPartial", "item number"], ["extraThrow", "paragraph sequence"], ["propertiesThrow", "restart shape in the guard"],
  ["numbersThrow", "numbers of other lists"], ["rereadFails", "failed re-read"]]) {
  test(`a Word error stays an unknown outcome when the ${seen} differs`, async () => {
    const s = await boot(continued(), {...LISTS, [mode]: true, ooxmlRefused: mode === "rereadFails"});
    const {applied} = await restartAt(s, "b8");
    assert.equal(applied.ok, false); assert.equal(applied.submitted, true); assert.equal(applied.unchanged, false, JSON.stringify(applied));
  });
}

for (const [mode, problem] of [["ooxmlNoContinue", /dalsze elementy listy/], ["ooxmlExtraParagraph", /układ akapitów po wstawieniu/]]) {
  test(`restart verification catches ${mode}`, async () => {
    const s = await boot(continued(), {...LISTS, [mode]: true});
    const {applied} = await restartAt(s, "b8");
    assert.equal(applied.ok, false); assert.equal(applied.submitted, true); assert.equal(applied.unchanged, false);
    assert.match(applied.error, problem); assert.match(applied.error, /Ctrl\+Z/);
  });
}

test("paragraph properties changed after the preview or during the write never pass as verified", async () => {
  const s = await boot(continued(), LISTS);
  const payload = {operation: "list_restart", paragraph_id: "b8", expected_sha256: await s.hash("b8"), text: "", find: ""};
  s.command = {command_id: "preview", type: "preview", payload}; await s.tick();
  const preview = s.results.at(-1);
  s.rows.find(r => r.id === "b8").list.num = "moved";  // another instance, same number: invisible to the Content hash
  s.command = {command_id: "apply", type: "apply", payload: {...payload, guard_sha256: preview.guard_sha256}}; await s.tick();
  assert.match(s.results.at(-1).error, /zmieniła się od podglądu/); assert.equal(s.writes, 0);
  // Whole exports: preview, apply plan, pre-write check (3), same batch as the write (4).
  const late = await boot(continued(), {...LISTS, mutateOnWholeExport: 3});
  assert.match((await restartAt(late, "b8")).applied.error, /zmieniła się od podglądu/); assert.equal(late.writes, 0);
  const racing = await boot(continued(), {...LISTS, mutateOnWholeExport: 4});
  const raced = (await restartAt(racing, "b8")).applied;
  assert.equal(raced.submitted, true); assert.match(raced.error, /zmieniły się w chwili zapisu/);
});

test("a restart stops shifting at an item that already restarts", async () => {
  // c12 is a restarted instance (num 3, startOverride 1) of the same definition; Word gives it its own list.
  const rows = continued();
  for (const [i, r] of rows.filter(r => /^c\d/.test(r.id)).entries())
    r.list = {...r.list, id: 6, abstract: "1", override: i ? undefined : 1, value: i + 1, string: `${i + 1}.`, num: 3};
  for (const [multiLevel, method] of [[true, "start_override"], [false, "start_override"]]) {
    const s = await boot(rows.map(r => ({...r, list: r.list && {...r.list}})), {lists: {1: ["Number"], 5: ["Number"], 6: ["Number"]}, multiLevel});
    const {preview, applied} = await restartAt(s, "b9");
    assert.equal(preview.method, method, `multiLevel ${multiLevel}`); assert.equal(preview.renumbered_item_count, 3);
    assert.equal(applied.ok, true, JSON.stringify(applied));
    assert.deepEqual(numbers(s, /^[bc]\d/), [8, 1, 2, 3, 1, 2, 3]);
    assert.equal(s.rows.find(r => r.id === "c12").list.num, "3", "the existing restart is kept");
  }
});

test("list_restart preview refuses what it cannot restart safely, before any write", async () => {
  const rows = continued(); rows.find(r => r.id === "b8").sectionBreak = true;
  rows.splice(rows.findIndex(r => r.id === "b10"), 0, {id: "t1", text: "in a cell", tableNestingLevel: 1, style: "List Number", list: {id: 1, level: 0, string: "10.", value: 10, num: 2}});
  const s = await boot(rows, LISTS);
  const blocked = await restartAt(s, "b8");
  assert.equal(blocked.preview.ok, false); assert.match(blocked.preview.error, /kończy sekcję Worda.*Restart at 1/);
  const cell = await restartAt(s, "t1");
  assert.equal(cell.preview.ok, false); assert.match(cell.preview.error, /leży w tabeli/);
  const first = await restartAt(s, "o1");
  assert.equal(first.preview.ok, false); assert.match(first.preview.error, /już od 1/);
  assert.equal(s.writes + s.separateCalls, 0);
});

test("insert_table_row adds one verified row after a data row", async () => {
  const table = {header: 1, rows: [["Component", "Role"], ["API", "Gateway"]]};
  const s = await boot([{id: "cell", text: "API", tableNestingLevel: 1, table, row: 1}, {id: "end", text: ""}]);
  const {preview, applied} = await s.run({operation: "insert_table_row", paragraph_id: "cell", expected_sha256: await s.hash("cell"),
    position: "after", cells: ["Queue", "Buffer"]});
  assert.deepEqual(preview.reference_row, ["API", "Gateway"]);
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.deepEqual(table.rows, [["Component", "Role"], ["API", "Gateway"], ["Queue", "Buffer"]]);
  const wrong = await s.run({operation: "insert_table_row", paragraph_id: "cell", expected_sha256: await s.hash("cell"), position: "after", cells: ["x"]});
  assert.equal(wrong.preview.ok, false); assert.match(wrong.preview.error, /2 komórek/);
});

for (const proxyMode of ["staleTableProxy", "sharedTableValues"]) {
  test(`insert_table_row verifies a fresh table with ${proxyMode}`, async () => {
    const table = {header: 1, rows: [["Component", "Role"], ["API", "Gateway"]]};
    const s = await boot([{id: "cell", text: "API", tableNestingLevel: 1, table, row: 1}, {id: "end", text: ""}],
      {[proxyMode]: true});
    const {applied} = await s.run({operation: "insert_table_row", paragraph_id: "cell", expected_sha256: await s.hash("cell"),
      position: "after", cells: ["Queue", "Buffer"]});
    assert.equal(applied.ok, true, JSON.stringify(applied));
    assert.deepEqual(table.rows, [["Component", "Role"], ["API", "Gateway"], ["Queue", "Buffer"]]);
  });
}

test("insert_table_row verifies after a host keeps the table proxy stale within one Word.run", async () => {
  const table = {header: 1, rows: [["Component", "Role"], ["API", "Gateway"]]};
  const s = await boot([{id: "cell", text: "API", tableNestingLevel: 1, table, row: 1}, {id: "end", text: ""}],
    {cachedTableProxyPerRun: true, staleTableProxy: true});
  const {applied} = await s.run({operation: "insert_table_row", paragraph_id: "cell", expected_sha256: await s.hash("cell"),
    position: "after", cells: ["Queue", "Buffer"]});
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.equal(applied.row_count, 3);
  assert.deepEqual(table.rows, [["Component", "Role"], ["API", "Gateway"], ["Queue", "Buffer"]]);
});

test("snapshot exposes list numbering for list paragraphs only", async () => {
  const s = await boot([{id: "i1", text: "one", list: {id: 1, level: 0, string: "4.", value: 4}}, {id: "end", text: ""}], {lists: {1: ["Number"]}});
  s.command = {command_id: "snap", type: "snapshot", payload: {scope: "body", start: 0, limit: 10, query: ""}}; await s.tick();
  const result = s.results.at(-1);
  assert.equal(result.complete, true, JSON.stringify(result.coverage_errors));
  assert.deepEqual(result.paragraphs.map(p => p.list_item), [{level: 0, list_string: "4."}, null]);
});

test("a failed check after a submitted write is reported as submitted (bridge marks it unknown)", async () => {
  const s = await boot(outline(), {brokenDelete: true});
  const {applied} = await s.run({operation: "delete_paragraph", paragraph_id: "a1", expected_sha256: await s.hash("a1"), text: "", find: ""});
  assert.equal(applied.ok, false); assert.equal(applied.submitted, true); assert.match(applied.error, /Ctrl\+Z/);
  const m = await boot(outline(), {untrackedRange: true});
  const move = await m.run({operation: "move_section", paragraph_id: "h-b", expected_sha256: await m.hash("h-b"),
    target_paragraph_id: "h-a", target_expected_sha256: await m.hash("h-a"), position: "before"});
  assert.equal(move.applied.ok, false); assert.equal(move.applied.submitted, true);
  assert.match(move.applied.error, /nie został usunięty/);
  assert.equal(m.rows.filter(r => r.text === "B.1").length, 2, "a duplicate, never lost content");
});

test("a stale paragraph hash stops every structural preview before any write", async () => {
  const s = await boot(outline());
  for (const payload of [{operation: "delete_paragraph", paragraph_id: "a1", text: "", find: ""},
    {operation: "format_text", paragraph_id: "b1", text: "bold", find: "Lead-in:"},
    {operation: "move_section", paragraph_id: "h-b", target_paragraph_id: "h-a", target_expected_sha256: await s.hash("h-a"), position: "before"}]) {
    s.command = {command_id: "p", type: "preview", payload: {...payload, expected_sha256: "0".repeat(64)}}; await s.tick();
    assert.equal(s.results.at(-1).ok, false, payload.operation); assert.equal(s.results.at(-1).submitted, false);
  }
  assert.equal(s.writes, 0);
});

test("move_section refuses edits made in the section while the export was hashed", async () => {
  const s = await boot(outline(), {editAfterExport: true});
  const {applied} = await s.run({operation: "move_section", paragraph_id: "h-b", expected_sha256: await s.hash("h-b"),
    target_paragraph_id: "h-a", target_expected_sha256: await s.hash("h-a"), position: "before"});
  assert.equal(applied.ok, false); assert.equal(applied.submitted, false); assert.equal(s.writes, 0);
  assert.equal(s.rows.find(r => r.id === "h-b").text, "B (co-author)", "the co-author edit survives");
});

test("revocation during the move stops cleanup and bookmark restore batches", async () => {
  const s = await boot(outline(), {revokeOnWrite: true});
  await s.run({operation: "move_section", paragraph_id: "h-b", expected_sha256: await s.hash("h-b"),
    target_paragraph_id: "h-a", target_expected_sha256: await s.hash("h-a"), position: "before"});
  assert.equal(s.writes, 2, "only the atomic insert+delete batch ran");
  assert.ok([...s.bookmarks.keys()].some(n => n.startsWith("WAI")), "no bookmark batch after revocation");
});

test("a bookmark starting at the destination keeps its original extent after the move", async () => {
  const rows = outline(); rows[1] = {...rows[1], bookmark: "_Toc0"};
  const s = await boot(rows);
  const {preview, applied} = await s.run({operation: "move_section", paragraph_id: "h-b", expected_sha256: await s.hash("h-b"),
    target_paragraph_id: "h-a", target_expected_sha256: await s.hash("h-a"), position: "before"});
  assert.deepEqual(preview.destination_bookmarks, ["_Toc0"]);
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.deepEqual(applied.destination_bookmarks_restored, ["_Toc0"]);
  assert.equal(s.bookmarkText("_Toc0"), "A", "the TOC bookmark of the target heading does not swallow the moved section");
});

test("destinations inside a bookmark or a field result are refused before any write", async () => {
  const spanning = await boot(outline(), {bookmarks: {span: ["intro", "a1"]}});
  const blocked = await spanning.run({operation: "move_section", paragraph_id: "h-b", expected_sha256: await spanning.hash("h-b"),
    target_paragraph_id: "h-a", target_expected_sha256: await spanning.hash("h-a"), position: "before"});
  assert.equal(blocked.preview.ok, false); assert.match(blocked.preview.error, /obejmująca miejsce docelowe/);
  const rows = outline(); rows[1] = {...rows[1], fieldEnd: true};  // a TOC ends at the start of heading A
  const toc = await boot(rows);
  const inside = await toc.run({operation: "move_section", paragraph_id: "h-b", expected_sha256: await toc.hash("h-b"),
    target_paragraph_id: "h-a", target_expected_sha256: await toc.hash("h-a"), position: "before"});
  assert.equal(inside.preview.ok, false); assert.match(inside.preview.error, /wewnątrz pola/);
  assert.equal(spanning.writes + toc.writes, 0);
});

test("list_restart needs WordApiDesktop 1.3 to verify the rendered number", async () => {
  const s = await boot([{id: "i1", text: "one", list: {id: 1, level: 0, string: "c)", value: 3}}, {id: "end", text: ""}], {lists: {1: ["Number"]}, desktop: 0});
  const {preview} = await s.run({operation: "list_restart", paragraph_id: "i1", expected_sha256: await s.hash("i1"), text: "", find: ""});
  assert.equal(preview.ok, false); assert.match(preview.error, /WordApiDesktop 1.3/); assert.equal(s.writes, 0);
});
