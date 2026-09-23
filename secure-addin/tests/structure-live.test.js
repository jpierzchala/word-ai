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
      if (options.exportFails) throw Error("export");
      if (kind === "span" && ++state.spanExports === 2 && options.editAfterExport) {
        const value = flat(state.rows.slice(index(a), index(b)));
        a.text += " (co-author)";  // typed while the apply export is being hashed
        return {value};
      }
      return {value: kind === "span" ? flat(state.rows.slice(index(a), index(b))) : kind === "table"
        ? `<w:document xmlns:w="${W}"><w:body><w:tbl>${a.table.rows.map(r => `<w:tr>${r.map(c => `<w:tc><w:p><w:r><w:t>${c}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`).join("")}</w:tbl></w:body></w:document>`
        : single(a)};
    },
    expandTo(other) { return other.kind === "End" ? range("anchor", a, other.a) : range("span", a, other.a); },
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
    insertOoxml(xml) {
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
          setLevelStartingNumber() { state.writes++; members().forEach((r, i) => { r.list.value = i + 1; r.list.string = `${i + 1}.`; }); }};
      },
      separateList() {
        state.writes++;
        const members = state.rows.filter(r => r.list?.id === rec.list.id);
        const tail = members.slice(members.indexOf(rec));
        state.lists[99] = [...state.lists[rec.list.id]];
        tail.forEach((r, i) => { r.list = {...r.list, id: 99, value: i + 1, string: `${i + 1}.`}; });
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
      delete() { state.writes++; if (!options.brokenDelete) state.rows.splice(index(rec), 1); },
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
      document: {getFilePropertiesAsync: cb => cb({status: "ok", value: {url: "https://test/doc.docx"}})}}};
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
  assert.equal(restart.preview.method, "separate_list");
  assert.equal(restart.applied.ok, true, JSON.stringify(restart.applied)); assert.equal(restart.applied.list_value, 1);
  assert.equal(s.rows[0].list.string, "4.", "earlier item keeps its number");
});

test("list_restart mid-list requires WordApiDesktop 1.4", async () => {
  const s = await boot([{id: "i1", text: "one", list: {id: 1, level: 0, string: "1.", value: 1}},
    {id: "i2", text: "two", list: {id: 1, level: 0, string: "2.", value: 2}}, {id: "end", text: ""}], {lists: {1: ["Number"]}, desktop: 1.3});
  const {preview} = await s.run({operation: "list_restart", paragraph_id: "i2", expected_sha256: await s.hash("i2"), text: "", find: ""});
  assert.equal(preview.ok, false); assert.match(preview.error, /WordApiDesktop 1.4/); assert.equal(s.writes, 0);
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
