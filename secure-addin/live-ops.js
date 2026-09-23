// Office.js execution for structural operations. Each operation re-derives its
// plan on apply and must match the preview's structural guard before writing.
import {W, canonicalNode, digest} from "./safety.js";
import {paragraphDigest, planDeleteParagraph, planFormat, planList, planTableRow, searchText} from "./document-edit.js";
import {BOOKMARK_NAME, compareSequence, emptyParagraphShape, emptyParagraphXml, expectedSequence, findSection, headingLevel, inspectSection,
  preparePackage, resolveDestination, sectionDigest, startsInsideField} from "./section-move.js";

export const STRUCTURAL = new Set(["delete_paragraph", "format_text", "list_restart", "list_type", "list_level",
  "insert_table_row", "move_section"]);
const stale = () => Error("Struktura dokumentu zmieniła się od podglądu. Przygotuj nowy podgląd.");
const guardOf = parts => digest(JSON.stringify(parts));
const plain = p => p.isNullObject ? null : {uniqueLocalId: p.uniqueLocalId, text: p.text, tableNestingLevel: p.tableNestingLevel};

// env: {beforeWrite(limitMs), beforeFollowUp(), submitted()} supplied by the
// command executor. beforeFollowUp re-checks revocation and document identity
// before every write batch after the first one.
export async function runStructural(current, payload, mode, env) {
  const operation = {delete_paragraph: deleteParagraph, format_text: formatText, list_restart: listChange,
    list_type: listChange, list_level: listChange, insert_table_row: tableRow, move_section: moveSection}[payload.operation];
  if (!operation || !["preview", "apply"].includes(mode)) throw Error("Nieobsługiwane polecenie.");
  return operation(current, payload, mode, env);
}

const DELETABLE_BODIES = new Set(["MainDoc", "Section", "TableCell", "Header", "Footer", "Footnote", "Endnote", "NoteItem"]);
const sequenceOf = collection => collection.items.map(p => [p.text, p.tableNestingLevel]);

async function deleteParagraph(current, payload, mode, env) {
  const {ctx, paragraph} = current;
  // Neighbours come from the parent body's own collection: getNext/getPrevious
  // semantics across table cells are undocumented.
  const body = paragraph.parentBody;
  body.load("type");
  const siblings = body.paragraphs;
  siblings.load("items/uniqueLocalId,items/text,items/tableNestingLevel");
  const whole = paragraph.getRange("Whole");
  const pictures = whole.inlinePictures;
  pictures.load("items");
  const wholeXml = whole.getOoxml();
  const control = paragraph.parentContentControlOrNullObject;
  control.load("isNullObject");
  const sections = ctx.document.sections;
  sections.load("items");
  await ctx.sync();
  // The last paragraph of every section but the final one carries its break.
  const breaks = sections.items.slice(0, -1).map(s => s.body.paragraphs.getLast());
  breaks.forEach(p => p.load("uniqueLocalId"));
  await ctx.sync();
  const index = siblings.items.findIndex(p => p.uniqueLocalId === payload.paragraph_id);
  if (index < 0 || !DELETABLE_BODIES.has(body.type)) throw Error("Nie ustalono położenia akapitu. Odczytaj snapshot ponownie.");
  const neighbour = i => siblings.items[i] ? plain(siblings.items[i]) : null;
  const context = {isLastInBody: index === siblings.items.length - 1,
    endsSection: breaks.some(p => p.uniqueLocalId === payload.paragraph_id) || sectionBreakIn(wholeXml.value),
    inContentControl: !control.isNullObject, tableNestingLevel: paragraph.tableNestingLevel,
    previous: neighbour(index - 1), next: neighbour(index + 1), pictures: pictures.items.length};
  const plan = planDeleteParagraph(current, payload, context);
  const guard = await guardOf(["delete_paragraph", current.hash, body.type, index, context.previous, context.next,
    context.pictures, siblings.items.length]);
  if (mode === "preview") return {ok: true, ...plan, guard_sha256: guard};
  if (guard !== payload.guard_sha256) throw stale();
  const before = sequenceOf(siblings);
  await env.beforeWrite();
  env.submitted();
  paragraph.delete();
  await ctx.sync();
  const after = body.paragraphs;
  after.load("items/text,items/tableNestingLevel");
  await ctx.sync();
  const expected = [...before.slice(0, index), ...before.slice(index + 1)];
  if (JSON.stringify(sequenceOf(after)) !== JSON.stringify(expected))
    throw Error("Weryfikacja usunięcia akapitu nie powiodła się. Sprawdź dokument (Ctrl+Z cofa zmianę).");
  return {ok: true, operation: plan.operation, deleted_text: plan.before, previous_text: plan.previous_text,
    next_text: plan.next_text, paragraph_count: after.items.length};
}

function sectionBreakIn(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (/<!DOCTYPE|<!ENTITY/i.test(xml) || doc.getElementsByTagName("parsererror").length) throw Error("Błędny eksport akapitu.");
  return Array.from(doc.getElementsByTagNameNS(W, "sectPr")).some(n => n.parentNode?.localName === "pPr");
}

const FONT = "bold,italic,underline";
const fontState = font => ({bold: font.bold, italic: font.italic, underline: font.underline});
async function formatText(current, payload, mode, env) {
  const {ctx} = current;
  const plan = planFormat(current, payload);
  let target = current.range;
  if (payload.find) {
    const matches = current.range.search(searchText(payload.find), {matchCase: true, matchWholeWord: false,
      matchWildcards: false, ignorePunct: false, ignoreSpace: false});
    matches.load("items/text");
    await ctx.sync();
    if (matches.items.length !== 1 || matches.items[0].text !== payload.find)
      throw Error("Wyszukanie Worda nie wskazało dokładnie jednego zgodnego fragmentu.");
    target = matches.items[0];
  }
  target.font.load(FONT);
  await ctx.sync();
  const before = fontState(target.font);
  if (mode === "preview") return {ok: true, ...plan, before_format: before};
  await env.beforeWrite();
  env.submitted();
  for (const [name, on] of Object.entries(plan.format)) target.font[name] = name === "underline" ? (on ? "Single" : "None") : on;
  await ctx.sync();
  target.font.load(FONT);
  current.range.load("text");
  await ctx.sync();
  const after = fontState(target.font);
  const matches = Object.entries(plan.format).every(([name, on]) =>
    name === "underline" ? after.underline === (on ? "Single" : "None") : after[name] === on);
  // Same Content range the pre-write text came from.
  if (!matches || current.range.text.replace(/\r$/, "") !== current.text)
    throw Error("Weryfikacja formatowania nie powiodła się.");
  const updated = current.paragraph.getRange("Content").getOoxml();
  await ctx.sync();
  return {ok: true, operation: plan.operation, paragraph_id: payload.paragraph_id, target: plan.target,
    before_format: before, format: after, ooxml_sha256: await paragraphDigest(updated.value)};
}

// Every list paragraph of the body with its list and rendered number.
// List.paragraphs can return other lists' paragraphs on Win32 (Office.js #6602),
// and List.id is not guaranteed stable across writes, so membership is read
// per paragraph and checks across a write compare paragraph IDs.
async function listState(current, payload) {
  const {ctx} = current;
  // Fresh proxies on every read: after separateList a cached List object may
  // still be the list the paragraph left.
  const paragraph = ctx.document.getParagraphByUniqueLocalId(payload.paragraph_id);
  const item = paragraph.listItemOrNullObject, list = paragraph.listOrNullObject;
  item.load("isNullObject,level,listString");
  list.load("isNullObject,id,levelTypes");
  const paragraphs = ctx.document.body.paragraphs;
  paragraphs.load("items/uniqueLocalId,items/text,items/isListItem");
  const lists = ctx.document.body.lists;
  lists.load("items/id,items/levelTypes");
  const format = Office.context.requirements.isSetSupported("WordApiDesktop", "1.3") ? paragraph.getRange("Content").listFormat : null;
  if (format) format.load("listValue");
  await ctx.sync();
  if (item.isNullObject || list.isNullObject) return {isListItem: false};
  const candidates = paragraphs.items.filter(p => p.isListItem).map(p => {
    const l = p.listOrNullObject, i = p.listItemOrNullObject;
    l.load("isNullObject,id");
    i.load("isNullObject,level,listString");
    return {p, l, i};
  });
  await ctx.sync();
  const all = candidates.filter(x => !x.l.isNullObject && !x.i.isNullObject).map(x => ({id: x.p.uniqueLocalId,
    list: x.l.id, text: x.p.text, level: x.i.level, listString: x.i.listString}));
  const items = all.filter(x => x.list === list.id).map(({list: _, ...x}) => x);
  return {isListItem: true, paragraph, item, list, id: list.id, level: item.level, listString: item.listString,
    levelTypes: list.levelTypes, items, all, index: items.findIndex(i => i.id === payload.paragraph_id),
    value: format ? format.listValue : null, canSeparate: Office.context.requirements.isSetSupported("WordApiDesktop", "1.4"),
    others: lists.items.filter(l => l.id !== list.id).map(l => JSON.stringify(l.levelTypes)).sort()};
}

async function listChange(current, payload, mode, env) {
  const {ctx} = current;
  const state = await listState(current, payload);
  const plan = planList(current, payload, state);
  const guard = await guardOf([payload.operation, current.hash, state.levelTypes, state.items, state.others]);
  if (mode === "preview") return {ok: true, ...plan, guard_sha256: guard};
  if (guard !== payload.guard_sha256) throw stale();
  await env.beforeWrite();
  env.submitted();
  if (payload.operation === "list_level") state.item.level = plan.new_level;
  else if (payload.operation === "list_type") {
    if (plan.new_type === "bullet") state.list.setLevelBullet(state.level, "Solid");
    else state.list.setLevelNumbering(state.level, "Arabic", [state.level, "."]);
  } else if (plan.method === "set_starting_number") state.list.setLevelStartingNumber(state.level, 1);
  else state.paragraph.separateList();
  await ctx.sync();
  let after = await listState(current, payload);
  if (plan.method === "separate_list" && after.isListItem && after.value !== null && after.value !== 1) {
    // The separated list restarts at its level's starting number; make it 1,
    // but only on a list that no longer contains the earlier items.
    const earlier = new Set(state.items.slice(0, state.index).map(i => i.id));
    if (after.items.some(i => earlier.has(i.id))) throw Error("Word nie rozdzielił listy. Sprawdź dokument (Ctrl+Z cofa zmianę).");
    await env.beforeFollowUp();
    after.list.setLevelStartingNumber(after.level, 1);
    await ctx.sync();
    after = await listState(current, payload);
  }
  const problems = [];
  const before = new Map(state.all.map(x => [x.id, x]));
  const now = new Map(after.isListItem ? after.all.map(x => [x.id, x]) : []);
  const self = now.get(payload.paragraph_id);
  if (!after.isListItem || !self) problems.push("akapit przestał być elementem listy");
  else {
    if (self.text !== before.get(payload.paragraph_id)?.text) problems.push("tekst akapitu");
    const others = [...before.values()].filter(x => x.id !== payload.paragraph_id);
    if (payload.operation === "list_level") {
      if (after.level !== plan.new_level) problems.push("poziom elementu");
      if (others.some(x => now.get(x.id)?.level !== x.level)) problems.push("poziomy innych elementów");
    } else if (payload.operation === "list_type") {
      if (after.levelTypes[state.level] !== (plan.new_type === "bullet" ? "Bullet" : "Number")) problems.push("typ poziomu");
      if (JSON.stringify(after.others) !== JSON.stringify(state.others)) problems.push("zmieniły się inne listy");
    } else {
      if (after.value !== null ? after.value !== 1 : !/^\D*1\D*$/.test(after.listString)) problems.push("numer elementu");
      const earlier = state.items.slice(0, Math.max(state.index, 0));
      if (earlier.some(x => now.get(x.id)?.listString !== x.listString)) problems.push("wcześniejsze elementy listy");
    }
    if (others.some(x => now.get(x.id)?.text !== x.text)) problems.push("tekst innych elementów");
    // Items of other lists keep their numbers whatever the operation.
    if (state.all.some(x => x.list !== state.id && now.get(x.id)?.listString !== x.listString)) problems.push("numeracja innych list");
  }
  if (problems.length) throw Error(`Weryfikacja listy nie powiodła się (${problems.join(", ")}). Sprawdź dokument (Ctrl+Z cofa zmianę).`);
  return {ok: true, operation: plan.operation, paragraph_id: payload.paragraph_id, level: after.level,
    level_type: after.levelTypes[after.level], list_string: after.listString, list_value: after.value,
    list_item_count: after.items.length};
}

function tableFacts(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (/<!DOCTYPE|<!ENTITY/i.test(xml) || doc.getElementsByTagName("parsererror").length) throw Error("Błędny eksport tabeli.");
  const table = doc.getElementsByTagNameNS(W, "tbl")[0];
  if (!table) throw Error("Nieobsługiwany eksport tabeli.");
  const all = Array.from(table.getElementsByTagNameNS(W, "*"));
  const merged = all.some(n => n.localName === "vMerge" || n.localName === "hMerge" ||
    (n.localName === "gridSpan" && n.getAttributeNS(W, "val") !== "1") || (n.localName === "tbl" && n !== table));
  if (all.some(n => n.localName === "sdt")) throw Error("Tabela zawiera kontrolki zawartości; dodawanie wiersza zablokowane.");
  if (all.some(n => ["ins", "del", "trPrChange", "tcPrChange", "tblPrChange", "cellIns", "cellDel", "cellMerge"].includes(n.localName)))
    throw Error("Tabela zawiera śledzone zmiany; dodawanie wiersza zablokowane.");
  return {canonical: canonicalNode(table), merged};
}

async function tableRow(current, payload, mode, env) {
  const {ctx, paragraph} = current;
  const table = paragraph.parentTableOrNullObject;
  table.load("isNullObject,nestingLevel,rowCount,headerRowCount,isUniform,values");
  await ctx.sync();
  if (table.isNullObject) throw Error("Akapit nie leży w tabeli.");
  const rows = table.rows;
  rows.load("items/cellCount,items/isHeader");
  const xml = table.getRange("Whole").getOoxml();
  await ctx.sync();
  const facts = tableFacts(xml.value);
  const merged = facts.merged || !table.isUniform || new Set(rows.items.map(r => r.cellCount)).size > 1;
  let rowIndex = -1;
  if (!merged) {
    // Reading a cell's indexes through its paragraph fails for vertically
    // merged cells (Office.js #3200); merged tables are rejected above.
    const cell = paragraph.parentTableCell;
    cell.load("rowIndex");
    await ctx.sync();
    rowIndex = cell.rowIndex;
  }
  // Keep the pre-write values independent of the Office.js proxy. Some hosts
  // update an already-loaded array when the table changes.
  const values = table.values.map(row => [...row]);
  const shape = {nestingLevel: table.nestingLevel, merged, rowIndex,
    rows: rows.items.map((r, i) => ({cellCount: r.cellCount, isHeader: r.isHeader, values: values[i]}))};
  const plan = planTableRow({...current, tableNestingLevel: paragraph.tableNestingLevel}, payload, shape);
  const guard = await guardOf(["insert_table_row", current.hash, facts.canonical, rowIndex, table.rowCount,
    table.headerRowCount, payload.position, payload.cells]);
  if (mode === "preview") return {ok: true, ...plan, header_row_count: table.headerRowCount, guard_sha256: guard};
  if (guard !== payload.guard_sha256) throw stale();
  const headerRows = table.headerRowCount, rowCount = table.rowCount;
  await env.beforeWrite();
  env.submitted();
  rows.items[rowIndex].insertRows(payload.position === "before" ? "Before" : "After", 1, [payload.cells]);
  await ctx.sync();
  // Word can keep every table proxy stale for the rest of this Word.run batch.
  // The caller verifies in a separate Word.run before reporting success.
  return {operation: plan.operation, verify_after_run: {paragraph_id: payload.paragraph_id,
    inserted: payload.position === "before" ? rowIndex : rowIndex + 1,
    row_count_before: rowCount, header_rows: headerRows, values_before: values, cells: payload.cells}};
}

export async function verifyTableRowAfterWrite(check) {
  return Word.run(async ctx => {
    const paragraph = ctx.document.getParagraphByUniqueLocalId(check.paragraph_id);
    const verified = paragraph.parentTableOrNullObject;
    verified.load("isNullObject,rowCount,headerRowCount,values");
    await ctx.sync();
    const actual = verified.isNullObject ? [] : verified.values;
    const inserted = check.inserted;
    const checks = {
      table_found: !verified.isNullObject,
      row_count: !verified.isNullObject && verified.rowCount === check.row_count_before + 1,
      header_row_count: !verified.isNullObject && verified.headerRowCount === check.header_rows,
      inserted_cells: JSON.stringify(actual[inserted]) === JSON.stringify(check.cells),
      existing_cells: JSON.stringify(actual.filter((_, i) => i !== inserted)) === JSON.stringify(check.values_before),
    };
    if (Object.values(checks).some(ok => !ok))
      throw Error(`Weryfikacja nowego wiersza nie powiodła się. Sprawdź tabelę (Ctrl+Z cofa zmianę). Kontrole: ${JSON.stringify(checks)}; wiersze: ${verified.isNullObject ? "brak tabeli" : verified.rowCount}/${check.row_count_before + 1}; nagłówki: ${verified.isNullObject ? "brak tabeli" : verified.headerRowCount}/${check.header_rows}.`);
    return {ok: true, operation: "insert_table_row", row_index: inserted, row_count: verified.rowCount, cells: actual[inserted]};
  });
}

const PARAGRAPH_FIELDS = "items/uniqueLocalId,items/text,items/style,items/styleBuiltIn,items/outlineLevel,items/tableNestingLevel";
const rowsOf = collection => collection.items.map(p => ({uniqueLocalId: p.uniqueLocalId, text: p.text, style: p.style,
  styleBuiltIn: p.styleBuiltIn, outlineLevel: p.outlineLevel, tableNestingLevel: p.tableNestingLevel}));
const INSIDE = new Set(["Inside", "InsideStart", "InsideEnd", "Equal"]);
const WITHIN_FIELD = new Set(["Inside", "InsideStart", "InsideEnd"]);

// Counts compared before and after a move (fields and notes: WordApi 1.4/1.5).
function objectCounts(ctx) {
  const body = ctx.document.body;
  const collections = {pictures: body.inlinePictures, tables: body.tables, fields: body.fields,
    footnotes: body.footnotes, endnotes: body.endnotes, sections: ctx.document.sections};
  Object.values(collections).forEach(c => c.load("items"));
  return {collections, read: () => Object.fromEntries(Object.entries(collections).map(([k, c]) => [k, c.items.length]))};
}

async function sectionPlan(current, payload) {
  const {ctx} = current;
  const paragraphs = ctx.document.body.paragraphs;
  paragraphs.load(PARAGRAPH_FIELDS);
  const counts = objectCounts(ctx);
  await ctx.sync();
  const list = rowsOf(paragraphs);
  const finalXml = paragraphs.items[paragraphs.items.length - 1].getRange("Whole").getOoxml();
  await ctx.sync();
  const finalEmpty = emptyParagraphXml(finalXml.value);
  let section;
  try { section = findSection(list, payload.paragraph_id, finalEmpty); }
  catch (error) {
    if (!finalEmpty && error.message.startsWith("Sekcja obejmuje ostatni akapit"))
      throw Error(`${error.message} Eksport końca: ${JSON.stringify(emptyParagraphShape(finalXml.value))}; długość tekstu: ${list[list.length - 1].text.length}.`);
    throw error;
  }
  const {target: targetIndex, destination} = resolveDestination(list, section, payload.target_paragraph_id, payload.position);
  const heading = paragraphs.items[section.start], boundary = paragraphs.items[section.end];
  const target = paragraphs.items[targetIndex], landing = paragraphs.items[destination];
  const targetXml = target.getRange("Content").getOoxml();
  const landingXml = landing.getRange("Content").getOoxml();
  // Tracked: both ranges are used again after the document has changed.
  const range = heading.getRange("Whole").expandTo(boundary.getRange("Start")).track();
  const insertion = landing.getRange("Start").track();
  range.load("text");
  const placement = insertion.compareLocationWith(range);
  const xml = range.getOoxml();
  const bookmarks = range.getBookmarks(true, false);
  const nearby = insertion.getBookmarks(true, true);
  const fieldPlacement = counts.collections.fields.items.map(f => insertion.compareLocationWith(f.result));
  const pictures = range.inlinePictures, tables = range.tables, controls = range.contentControls;
  pictures.load("items"); tables.load("items"); controls.load("items");
  const outerSource = range.parentContentControlOrNullObject, outerTarget = insertion.parentContentControlOrNullObject;
  outerSource.load("isNullObject"); outerTarget.load("isNullObject");
  await ctx.sync();
  // Index arithmetic already excludes no-op and inner targets; Word must agree.
  if (!["Before", "After"].includes(placement.value)) throw Error("Cel leży wewnątrz lub na granicy przenoszonej sekcji.");
  const inspection = inspectSection(xml.value);
  const blockers = inspection.blockers;
  if ((controls.items.length || !outerSource.isNullObject || !outerTarget.isNullObject) && !blockers.includes("kontrolki zawartości"))
    blockers.push("kontrolki zawartości");
  // Content inserted inside a field result (e.g. a TOC) is discarded by the next field update.
  if (fieldPlacement.some(r => WITHIN_FIELD.has(r.value)) || startsInsideField(landingXml.value))
    blockers.push("miejsce docelowe wewnątrz pola (np. spisu treści)");
  // Resolve every bookmark first; comparing a null object would abort the batch.
  const resolve = names => names.filter(n => n !== "_GoBack").map(name => {
    const r = ctx.document.getBookmarkRangeOrNullObject(name);
    r.load("isNullObject,text");
    return {name, r};
  });
  const inner = resolve(bookmarks.value), near = resolve(nearby.value);
  await ctx.sync();
  if ([...inner, ...near].some(m => m.r.isNullObject)) throw Error("Nie udało się odczytać zakładki w sekcji lub w miejscu docelowym. Odczytaj snapshot ponownie.");
  // A bookmark reaching outside the section would be cut by the move.
  const relations = inner.map(m => m.r.compareLocationWith(range));
  // Word extends bookmarks that start at or span the insertion point over the copy.
  const landingMarks = near.map(m => ({...m, relation: insertion.compareLocationWith(m.r)}));
  await ctx.sync();
  if (relations.some(r => !INSIDE.has(r.value))) blockers.push("zakładka przekraczająca granicę sekcji");
  if (landingMarks.some(m => m.relation.value === "Inside")) blockers.push("zakładka obejmująca miejsce docelowe");
  if (landingMarks.some(m => !BOOKMARK_NAME.test(m.name))) blockers.push("zakładka w miejscu docelowym o nazwie, której Word nie pozwala odtworzyć");
  if (blockers.length) throw Error(`Sekcja zawiera: ${[...new Set(blockers)].join(", ")}. Przeniesienie zablokowane.`);
  const targetHash = await paragraphDigest(targetXml.value);
  const digestValue = await sectionDigest(xml.value);
  const moved = list.slice(section.start, section.end);
  const guard = await guardOf(["move_section", current.hash, targetHash, digestValue, moved.map(p => p.uniqueLocalId),
    list[section.end].uniqueLocalId, list[destination].uniqueLocalId, [list[destination].text, list[destination].style],
    payload.position, list.length, landingMarks.map(m => [m.name, m.r.text])]);
  return {list, section, destination, range, insertion, target, xml: xml.value, digest: digestValue, sourceText: range.text,
    inspection, targetHash, guard, counts: counts.read(), landingMarks: landingMarks.map(m => ({name: m.name, text: m.r.text})),
    sectionCounts: {pictures: pictures.items.length, tables: tables.items.length}};
}

async function moveSection(current, payload, mode, env) {
  const {ctx} = current;
  if (!current.hash || current.hash !== payload.expected_sha256) throw Error("Nagłówek zmienił się od odczytu. Odczytaj go ponownie.");
  let plan;
  try { plan = await sectionPlan(current, payload); }
  catch (error) { error.message = `Plan sekcji: ${error.message}`; throw error; }
  if (plan.targetHash !== payload.target_expected_sha256) throw Error("Akapit docelowy zmienił się od odczytu. Odczytaj go ponownie.");
  const {list, section, destination, inspection} = plan;
  const moved = list.slice(section.start, section.end);
  const last = moved[moved.length - 1];
  const summary = {operation: "move_section", heading: moved[0].text.slice(0, 200), heading_level: section.level,
    first_paragraph: {paragraph_id: moved[0].uniqueLocalId, text: moved[0].text.slice(0, 200)},
    last_paragraph: {paragraph_id: last.uniqueLocalId, text: last.text.slice(0, 200), in_table: Boolean(last.tableNestingLevel)},
    paragraph_count: moved.length, subsection_count: moved.slice(1).filter(p => headingLevel(p)).length,
    table_count: plan.sectionCounts.tables, inline_picture_count: plan.sectionCounts.pictures,
    floating_object_count: inspection.counts.floating_objects, field_count: inspection.counts.fields,
    footnote_count: inspection.counts.footnotes + inspection.counts.endnotes, bookmark_count: inspection.counts.bookmarks,
    ends_with_table: inspection.ends_with_table,
    section_ends_before: section.end === list.length - 1 && headingLevel(list[section.end]) === null ? "koniec dokumentu" : list[section.end].text.slice(0, 200),
    target: {paragraph_id: payload.target_paragraph_id, position: payload.position},
    destination_before: {text: list[destination].text.slice(0, 200), style: list[destination].style},
    destination_bookmarks: plan.landingMarks.map(m => m.name),
    note: "Po przeniesieniu zaktualizuj pola (numery podpisów, spis treści, odsyłacze): Ctrl+A, F9."};
  if (mode === "preview") return {ok: true, ...summary, ooxml_sha256: current.hash, target_ooxml_sha256: plan.targetHash,
    guard_sha256: plan.guard};
  if (plan.guard !== payload.guard_sha256) throw stale();
  const expected = expectedSequence(list, section, destination);
  const prepared = preparePackage(plan.xml);
  // Last light read. Hashing the export takes seconds; text, structure,
  // pictures or target edits made meanwhile would otherwise be overwritten by
  // the older copy.
  const fresh = ctx.document.body.paragraphs;
  fresh.load(PARAGRAPH_FIELDS);
  const freshPictures = plan.range.inlinePictures;
  freshPictures.load("items");
  const freshTarget = ctx.document.getParagraphByUniqueLocalId(payload.target_paragraph_id).getRange("Content").getOoxml();
  try { await ctx.sync(); }
  catch (error) { error.message = `Kontrola przed zapisem sekcji: ${error.message}`; throw error; }
  if (JSON.stringify(rowsOf(fresh)) !== JSON.stringify(list) || freshPictures.items.length !== plan.sectionCounts.pictures ||
      await paragraphDigest(freshTarget.value) !== plan.targetHash) throw stale();
  // Identity, revocation and write window are the last check before the write.
  try { await env.beforeWrite(); }
  catch (error) { error.message = `Ostatnia kontrola przed zapisem sekcji: ${error.message}`; throw error; }
  env.submitted();
  // Batch 1: the copy. The source is exported once more in the same batch,
  // immediately before the insertion.
  const committed = plan.range.getOoxml();
  plan.insertion.insertOoxml(prepared.xml, "Before");
  plan.range.load("text");
  await ctx.sync();
  // Delete only what is still exactly the previewed original: a last-moment
  // edit, or a range that did not follow the insertion, leaves a duplicate
  // instead of removing the wrong content.
  if (plan.range.text !== plan.sourceText || await sectionDigest(committed.value) !== plan.digest)
    throw Error("Kopia sekcji została wstawiona, ale oryginał zmienił się w chwili zapisu i nie został usunięty. Usuń duplikat albo cofnij (Ctrl+Z).");
  // Batch 2: remove the original.
  await env.beforeFollowUp();
  plan.range.delete();
  await ctx.sync();
  // Batch 3: bookmark names and extents, before any check that could stop here.
  const temps = prepared.bookmarks.map(b => ({...b, range: ctx.document.getBookmarkRangeOrNullObject(b.temp)}));
  temps.forEach(b => b.range.load("isNullObject"));
  const landing = plan.landingMarks.map(m => ({...m, range: ctx.document.getBookmarkRangeOrNullObject(m.name)}));
  landing.forEach(m => m.range.load("isNullObject,text"));
  await ctx.sync();
  const extended = landing.filter(m => !m.range.isNullObject && m.range.text !== m.text);
  if (temps.length || extended.length) {
    await env.beforeFollowUp();
    for (const b of temps) if (!b.range.isNullObject) { b.range.insertBookmark(b.name); ctx.document.deleteBookmark(b.temp); }
    // A bookmark that started at the insertion point now also covers the copy; restore its extent.
    for (const m of extended) (m.text ? plan.insertion.expandTo(m.range.getRange("End")) : plan.insertion).insertBookmark(m.name);
    await ctx.sync();
  }
  const reload = async () => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load(PARAGRAPH_FIELDS);
    await ctx.sync();
    return paragraphs;
  };
  let after = await reload();
  let comparison = compareSequence(rowsOf(after), expected);
  const seamRemoved = comparison.extras;
  if (comparison.outcome === "seam") {
    // Only empty body paragraphs Word left between the copy and the destination.
    await env.beforeFollowUp();
    for (let i = comparison.extras - 1; i >= 0; i--) after.items[expected.sentinelIndex + i].delete();
    await ctx.sync();
    after = await reload();
    comparison = compareSequence(rowsOf(after), expected);
  }
  if (comparison.outcome !== "exact")
    throw Error("Weryfikacja kolejności akapitów po przeniesieniu nie powiodła się. Sprawdź dokument (Ctrl+Z cofa zmianę).");
  const counts = objectCounts(ctx);
  const restored = prepared.bookmarks.map(b => ({name: b.name, now: ctx.document.getBookmarkRangeOrNullObject(b.name),
    temp: ctx.document.getBookmarkRangeOrNullObject(b.temp)}));
  restored.forEach(x => { x.now.load("isNullObject"); x.temp.load("isNullObject"); });
  const kept = plan.landingMarks.map(m => ({...m, now: ctx.document.getBookmarkRangeOrNullObject(m.name)}));
  kept.forEach(m => m.now.load("isNullObject,text"));
  await ctx.sync();
  const lost = [...restored.filter(x => x.now.isNullObject || !x.temp.isNullObject).map(x => x.name),
    ...kept.filter(m => m.now.isNullObject || m.now.text !== m.text).map(m => m.name)];
  const changed = Object.entries(counts.read()).filter(([k, v]) => v !== plan.counts[k]).map(([k]) => k);
  if (lost.length || changed.length)
    throw Error(`Przeniesienie wykonane, ale weryfikacja nie powiodła się (${[...lost.map(n => `zakładka ${n}`), ...changed].join(", ")}). Sprawdź dokument (Ctrl+Z cofa zmianę).`);
  return {ok: true, ...summary, seam_paragraphs_removed: seamRemoved, bookmarks_restored: prepared.bookmarks.length,
    destination_bookmarks_restored: extended.map(m => m.name), paragraph_count_after: after.items.length};
}
