import {paragraphXml, paragraphDigest, planEdit, planImage, planImageChange, searchText} from "./document-edit.js";
import {STRUCTURAL, runStructural, verifyTableRowAfterWrite} from "./live-ops.js";

const GRANT = "document-read-write-v2";
const $ = id => document.getElementById(id);
const status = text => { $("status").textContent = text; };
let session = null;
let pairing = null;
let generation = 0;
let polling = false;
let granting = false;

function documentUrl() {
  return new Promise((resolve, reject) => Office.context.document.getFilePropertiesAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value.url)
      reject(Error("Zapisz dokument przed udostępnieniem."));
    else resolve(result.value.url);
  }));
}

async function request(route, body, token) {
  const response = await fetch(`/office/${route}`, {method: "POST", cache: "no-store", credentials: "omit",
    headers: {"Content-Type": "application/json", ...(token ? {"Authorization": `Bearer ${token}`} : {})},
    body: JSON.stringify(body), signal: AbortSignal.timeout(10000)});
  if (!response.ok) {
    const error = Error(`Połączenie wygasło lub żądanie zostało odrzucone (${response.status}).`);
    error.status = response.status;
    throw error;
  }
  return response.json();
}

function renderAccess() {
  $("connect").disabled = Boolean(session || pairing || granting);
  $("disconnect").disabled = !session && !pairing && !granting;
  $("document").textContent = (session || pairing)?.document || "";
  $("access").textContent = session ? "Dostęp do dokumentu aktywny" : pairing ? "Oczekiwanie na agenta" : "Brak dostępu";
}

function resetAccess(message) {
  generation++;
  session = null;
  pairing = null;
  granting = false;
  renderAccess();
  status(message);
}

async function revokeAccess() {
  const previous = session || pairing;
  const wasSession = Boolean(session);
  resetAccess("Dostęp cofnięty. Zapis wysłany już do Worda mógł się zakończyć.");
  if (!previous) return;
  const route = wasSession ? "disconnect" : "pair-cancel";
  const body = wasSession ? {session_id: previous.session_id} : {pairing_id: previous.pairing_id};
  try { await request(route, {document: previous.document, ...body}, previous.capability || previous.secret); }
  catch { status("Dostęp w panelu cofnięty. Mostek nie odpowiada; sesja po jego stronie wygaśnie automatycznie."); }
}

async function grantAccess() {
  if (session || pairing || granting) return;
  granting = true;
  renderAccess();
  const gen = ++generation;
  try {
    const url = await documentUrl();
    const ticket = await request("pair", {document: url, access: GRANT});
    if (gen !== generation) {
      await request("pair-cancel", {pairing_id: ticket.pairing_id, document: url}, ticket.secret).catch(() => {});
      return;
    }
    granting = false;
    pairing = {...ticket, document: url};
    renderAccess();
    status("Powiedz Codexowi lub Claude Code, co zrobić w tym dokumencie. Agent połączy się z udzielonym dostępem.");
  } catch (error) { if (gen === generation) resetAccess(error.message); }
}

async function verifyIdentity(expected, gen) {
  if (generation !== gen || (session || pairing) !== expected) throw Error("Dostęp został cofnięty.");
  if (await documentUrl() !== expected.document) {
    await revokeAccess();
    throw Error("Dokument zmienił tożsamość. Udostępnij go ponownie.");
  }
  if (generation !== gen || (session || pairing) !== expected) throw Error("Dostęp został cofnięty.");
}

async function api(route, extra, expected, gen) {
  await verifyIdentity(expected, gen);
  return request(route, {session_id: expected.session_id, document: expected.document, ...extra}, expected.capability);
}

async function snapshot(payload, expected, gen) {
  await verifyIdentity(expected, gen);
  return Word.run(async ctx => {
    const bodies = [];
    if (payload.scope === "body") bodies.push({body: ctx.document.body, label: "body"});
    else if (payload.scope === "footnotes" || payload.scope === "endnotes")
      bodies.push({body: payload.scope === "footnotes" ? ctx.document.getFootnoteBody() : ctx.document.getEndnoteBody(), label: payload.scope});
    else {
      const sections = ctx.document.sections;
      sections.load("items"); await ctx.sync();
      for (let i = 0; i < sections.items.length; i++) for (const kind of ["Primary", "FirstPage", "EvenPages"])
        bodies.push({body: payload.scope === "headers" ? sections.items[i].getHeader(kind) : sections.items[i].getFooter(kind), label: `${payload.scope}:${i}:${kind}`});
    }
    const coverageErrors = [];
    for (const entry of bodies) {
      entry.paragraphs = entry.body.paragraphs;
      try {
        entry.paragraphs.load("items/uniqueLocalId,items/text,items/style,items/styleBuiltIn,items/tableNestingLevel");
        await ctx.sync();
      } catch (error) {
        if (payload.scope === "body") throw error;
        entry.paragraphs = null;
        coverageErrors.push({location: entry.label, error: error.code || error.message,
          api: error.debugInfo?.errorLocation || null});
      }
    }
    const seen = new Set();
    const all = [];
    for (const entry of bodies) for (const paragraph of entry.paragraphs?.items || []) {
      if (!seen.has(paragraph.uniqueLocalId)) {
        seen.add(paragraph.uniqueLocalId);
        all.push({paragraph, location: entry.label});
      }
    }
    const matches = all.filter(p => !payload.query || p.paragraph.text.toLocaleLowerCase().includes(payload.query.toLocaleLowerCase()));
    const page = matches.slice(payload.start, payload.start + payload.limit);
    for (const p of page) {
      try {
        p.range = p.paragraph.getRange("Content");
        p.xml = p.range.getOoxml();
        p.pictures = p.range.inlinePictures;
        p.pictures.load("items/width,items/height,items/altTextTitle,items/altTextDescription");
        await ctx.sync();
      } catch (error) {
        p.xml = null;
        p.pictures = null;
        coverageErrors.push({location: p.location, paragraph_id: p.paragraph.uniqueLocalId,
          error: error.code || error.message, api: error.debugInfo?.errorLocation || null});
      }
    }
    // List numbering is not part of paragraph text; read it separately so a
    // list failure never costs the structural hash of the page.
    try {
      for (const p of page) {
        p.list = p.paragraph.listItemOrNullObject;
        p.list.load("isNullObject,level,listString");
      }
      await ctx.sync();
    } catch (error) {
      for (const p of page) p.list = null;
      coverageErrors.push({location: payload.scope, error: error.code || error.message, api: "listItemOrNullObject"});
    }
    const paragraphs = [];
    for (const p of page) {
      let hash = null, editSupport = "read_only";
      try {
        const shape = paragraphXml(p.xml?.value);
        hash = await paragraphDigest(p.xml.value);
        editSupport = shape.plain ? (shape.uniform ? "text_and_insert" : "targeted_text_and_insert") : "insert_or_style_only";
      } catch (error) {
        // Readable text is not evidence of complete structural/image coverage.
        if (p.xml) coverageErrors.push({location: p.location, paragraph_id: p.paragraph.uniqueLocalId,
          error: error.message, api: "paragraphDigest"});
      }
      const inlinePictures = (p.pictures?.items || []).map((picture, imageIndex) => ({image_index: imageIndex,
        width_pt: picture.width, height_pt: picture.height, alt_title: picture.altTextTitle || "",
        alt_text: picture.altTextDescription || ""}));
      paragraphs.push({paragraph_id: p.paragraph.uniqueLocalId, text: p.paragraph.text.slice(0, 500),
        truncated: p.paragraph.text.length > 500, text_length: p.paragraph.text.length, style: p.paragraph.style,
        location: p.location, table_nesting_level: p.paragraph.tableNestingLevel, ooxml_sha256: hash, edit_support: editSupport,
        inline_picture_count: inlinePictures.length, inline_pictures: inlinePictures,
        list_item: p.list && !p.list.isNullObject ? {level: p.list.level, list_string: p.list.listString} : null});
    }
    await verifyIdentity(expected, gen);
    return {ok: true, scope: payload.scope, total: all.length, matched: matches.length, start: payload.start,
      next_start: payload.start + page.length < matches.length ? payload.start + page.length : null,
      paragraphs, complete: coverageErrors.length === 0, coverage_errors: coverageErrors,
      coverage: "Paragraphs in the requested scope, including table cells and inline-picture metadata. Floating shapes/textboxes are not enumerated. Any coverage_errors mean unread portions, not empty content."};
  });
}

async function withParagraph(payload, expected, gen, action, writing=false) {
  await verifyIdentity(expected, gen);
  return Word.run(async ctx => {
    const paragraph = ctx.document.getParagraphByUniqueLocalId(payload.paragraph_id);
    paragraph.load("uniqueLocalId,text,style,styleBuiltIn,tableNestingLevel");
    ctx.document.load("changeTrackingMode");
    const range = paragraph.getRange("Content");
    range.load("text");
    const parent = range.parentContentControlOrNullObject;
    parent.load("isNullObject,cannotEdit");
    const xml = range.getOoxml();
    await ctx.sync();
    if (paragraph.uniqueLocalId !== payload.paragraph_id) throw Error("Tożsamość akapitu uległa zmianie.");
    if (writing && ctx.document.changeTrackingMode !== "Off") throw Error("Śledzenie zmian jest włączone. Ten profil nie zmienia go automatycznie.");
    if (writing && !parent.isNullObject && parent.cannotEdit) throw Error("Akapit jest zablokowany przez kontrolkę Worda.");
    let hash = null;
    try { hash = await paragraphDigest(xml.value); } catch (error) { if (writing) throw error; }
    const current = {ctx, paragraph, range, xml: xml.value, text: range.text.replace(/\r$/, ""), style: paragraph.styleBuiltIn, hash};
    await verifyIdentity(expected, gen);
    return action(current);
  });
}

async function execute(command, expected, gen) {
  let started = false, submitted = false;
  try {
    await api("begin", {command_id: command.command_id}, expected, gen);
    started = true;
    const begun = performance.now();
    let result;
    if (command.type === "snapshot") result = await snapshot(command.payload, expected, gen);
    else result = await withParagraph(command.payload, expected, gen, async current => {
      if (command.type === "read") return {ok: true, paragraph_id: command.payload.paragraph_id,
        text: current.text.slice(command.payload.start, command.payload.start + command.payload.limit), total: current.text.length,
        next_start: command.payload.start + command.payload.limit < current.text.length ? command.payload.start + command.payload.limit : null,
        style: current.style, ooxml_sha256: current.hash};
      if (STRUCTURAL.has(command.payload.operation)) {
        const result = await runStructural(current, command.payload, command.type, {
          submitted: () => { submitted = true; },
          beforeFollowUp: () => verifyIdentity(expected, gen),
          // Same guard as the text path: fresh hash, identity and write window
          // (20 s: structural edits read the whole body before writing).
          beforeWrite: async (limit = 20000) => {
            const latest = current.ctx.document.getParagraphByUniqueLocalId(command.payload.paragraph_id);
            const fresh = latest.getRange("Content").getOoxml(); await current.ctx.sync();
            if (await paragraphDigest(fresh.value) !== current.hash) throw Error("Wykryto zmianę podczas przygotowania zapisu.");
            await verifyIdentity(expected, gen);
            if (performance.now() - begun > limit) throw Error("Okno wykonania wygasło przed zapisem.");
          }});
        return command.type === "preview" ? {ooxml_sha256: current.hash, ...result} : result;
      }
      let plan, targetPicture = null;
      if (command.payload.operation === "insert_image") plan = await planImage(current, command.payload);
      else if (["replace_image", "delete_image"].includes(command.payload.operation)) {
        const pictures = current.range.inlinePictures;
        pictures.load("items/width,items/height,items/altTextTitle,items/altTextDescription");
        await current.ctx.sync();
        plan = await planImageChange(current, command.payload, pictures.items);
        targetPicture = pictures.items[plan.image_index];
        const pictureParent = targetPicture.parentContentControlOrNullObject;
        pictureParent.load("isNullObject,cannotEdit");
        await current.ctx.sync();
        if (!pictureParent.isNullObject && pictureParent.cannotEdit) throw Error("Obraz jest zablokowany przez kontrolkę Worda.");
      } else plan = planEdit(current, command.payload);
      if (command.type === "preview") return {ok: true, ...plan, ooxml_sha256: current.hash};
      if (command.type !== "apply") throw Error("Nieobsługiwane polecenie.");
      let matches;
      if (plan.operation === "replace_text") {
        matches = current.range.search(searchText(command.payload.find), {matchCase: true, matchWholeWord: false, matchWildcards: false, ignorePunct: false, ignoreSpace: false});
        matches.load("items/text"); await current.ctx.sync();
        if (matches.items.length !== 1 || matches.items[0].text !== command.payload.find) throw Error("Wyszukanie Worda nie wskazało dokładnie jednego zgodnego fragmentu.");
      }
      const fresh = current.range.getOoxml(); await current.ctx.sync();
      if (await paragraphDigest(fresh.value) !== current.hash) throw Error("Wykryto zmianę podczas przygotowania zapisu.");
      await verifyIdentity(expected, gen);
      if (performance.now() - begun > 5000) throw Error("Okno wykonania wygasło przed zapisem.");
      if (plan.operation === "insert_image") {
        // Failed batches can leave partial content; never automatically retry.
        submitted = true;
        const imageParagraph = current.paragraph.insertParagraph("", plan.position === "before" ? "Before" : "After");
        imageParagraph.styleBuiltIn = "Normal";
        const picture = imageParagraph.insertInlinePictureFromBase64(command.payload.image_base64, "End");
        picture.lockAspectRatio = false;
        picture.width = plan.width_pt; picture.height = plan.height_pt;
        picture.lockAspectRatio = true; picture.altTextDescription = plan.alt_text;
        const captionParagraph = plan.caption ? imageParagraph.insertParagraph(plan.caption, "After") : null;
        if (captionParagraph) captionParagraph.styleBuiltIn = "Normal";
        await current.ctx.sync();
        picture.load("width,height,altTextDescription");
        const actualParagraph = picture.paragraph;
        actualParagraph.load("uniqueLocalId");
        if (captionParagraph) captionParagraph.load("text");
        await current.ctx.sync();
        if (Math.abs(picture.width - plan.width_pt) > 0.2 || Math.abs(picture.height - plan.height_pt) > 0.2 ||
            picture.altTextDescription !== plan.alt_text || (captionParagraph && captionParagraph.text.replace(/\r$/, "") !== plan.caption))
          throw Error("Weryfikacja obrazu lub podpisu nie powiodła się.");
        return {ok: true, operation: plan.operation, paragraph_id: actualParagraph.uniqueLocalId,
          image_sha256: plan.image_sha256, image_name: plan.image_name, width_pt: picture.width,
          height_pt: picture.height, caption: plan.caption, alt_text: picture.altTextDescription};
      }
      if (plan.operation === "replace_image") {
        // Replace is a single Office.js mutation. Unknown outcomes must never be retried automatically.
        submitted = true;
        const picture = targetPicture.insertInlinePictureFromBase64(command.payload.image_base64, "Replace");
        picture.lockAspectRatio = false;
        picture.width = plan.width_pt; picture.height = plan.height_pt;
        picture.lockAspectRatio = true; picture.altTextDescription = plan.alt_text;
        await current.ctx.sync();
        picture.load("width,height,altTextDescription");
        const actualParagraph = picture.paragraph;
        actualParagraph.load("uniqueLocalId");
        const afterPictures = actualParagraph.getRange("Content").inlinePictures;
        afterPictures.load("items");
        await current.ctx.sync();
        if (actualParagraph.uniqueLocalId !== command.payload.paragraph_id || afterPictures.items.length !== plan.image_count_before ||
            Math.abs(picture.width - plan.width_pt) > 0.2 || Math.abs(picture.height - plan.height_pt) > 0.2 ||
            picture.altTextDescription !== plan.alt_text) throw Error("Weryfikacja podmienionego obrazu nie powiodła się.");
        const updated = actualParagraph.getRange("Content").getOoxml(); await current.ctx.sync();
        return {ok: true, operation: plan.operation, paragraph_id: actualParagraph.uniqueLocalId,
          image_index: plan.image_index, image_count: afterPictures.items.length, image_sha256: plan.image_sha256,
          image_name: plan.image_name, width_pt: picture.width, height_pt: picture.height,
          alt_text: picture.altTextDescription, previous_image: plan.previous_image,
          ooxml_sha256: await paragraphDigest(updated.value)};
      }
      if (plan.operation === "delete_image") {
        submitted = true;
        targetPicture.delete();
        await current.ctx.sync();
        current.paragraph.load("uniqueLocalId");
        const afterPictures = current.paragraph.getRange("Content").inlinePictures;
        afterPictures.load("items");
        await current.ctx.sync();
        if (current.paragraph.uniqueLocalId !== command.payload.paragraph_id || afterPictures.items.length !== plan.image_count_before - 1)
          throw Error("Weryfikacja usunięcia obrazu nie powiodła się.");
        return {ok: true, operation: plan.operation, paragraph_id: current.paragraph.uniqueLocalId,
          image_index: plan.image_index, image_count: afterPictures.items.length, previous_image: plan.previous_image};
      }
      let changed = current.paragraph;
      if (plan.operation === "replace_text") matches.items[0].insertText(command.payload.text, "Replace");
      else if (plan.operation === "replace_paragraph") current.range.insertText(command.payload.text, "Replace");
      else if (plan.operation === "set_style") current.paragraph.styleBuiltIn = command.payload.text;
      else changed = current.paragraph.insertParagraph(command.payload.text, plan.operation === "insert_before" ? "Before" : "After");
      submitted = true;
      await current.ctx.sync();
      changed.load("uniqueLocalId,text,styleBuiltIn"); await current.ctx.sync();
      const actual = changed.text.replace(/\r$/, "");
      const expectedText = plan.operation.startsWith("insert_") ? command.payload.text : plan.after;
      if (actual !== expectedText) throw Error("Tekst po zapisie różni się od podglądu.");
      if (plan.operation === "set_style" && changed.styleBuiltIn !== command.payload.text) throw Error("Weryfikacja stylu nie powiodła się.");
      // Word's Content range has no OOXML for a newly inserted empty paragraph.
      // Its new ID and empty text above are the available post-write evidence.
      if (plan.operation.startsWith("insert_") && expectedText === "")
        return {ok: true, operation: plan.operation, paragraph_id: changed.uniqueLocalId, text: actual,
          style: changed.styleBuiltIn, ooxml_sha256: null};
      const updated = changed.getRange("Content").getOoxml(); await current.ctx.sync();
      return {ok: true, operation: plan.operation, paragraph_id: changed.uniqueLocalId, text: actual,
        style: changed.styleBuiltIn, ooxml_sha256: await paragraphDigest(updated.value)};
    }, command.type !== "read");
    if (result?.verify_after_run) {
      await verifyIdentity(expected, gen);
      result = await verifyTableRowAfterWrite(result.verify_after_run);
      await verifyIdentity(expected, gen);
    }
    await api("result", {command_id: command.command_id, result}, expected, gen);
    if (command.type === "apply") {
      const labels = {insert_image: "Wstawiono obraz", replace_image: "Podmieniono obraz", delete_image: "Usunięto obraz", replace_text: "Zmieniono fragment tekstu", replace_paragraph: "Zmieniono akapit", insert_before: "Dodano akapit", insert_after: "Dodano akapit", set_style: "Zmieniono styl akapitu",
        delete_paragraph: "Usunięto akapit", format_text: "Zmieniono formatowanie fragmentu", list_restart: "Rozpoczęto numerację od 1", list_type: "Zmieniono typ listy", list_level: "Zmieniono poziom listy", insert_table_row: "Dodano wiersz tabeli", move_section: "Przeniesiono sekcję"};
      $("last-change").textContent = `${new Date().toLocaleTimeString()} — ${labels[result.operation]}. Wynik sprawdzony.`;
      status("Zmiana wykonana w ramach udzielonego dostępu do dokumentu.");
    }
  } catch (error) {
    if (started && session === expected && generation === gen) {
      await api("result", {command_id: command.command_id, result: {ok: false, submitted, error: error.message,
        api: error.debugInfo?.errorLocation || null}}, expected, gen).catch(() => {});
    }
    if (generation === gen) status(`${error.message}${submitted ? " Wynik może być niepewny — nie ponawiaj zapisu automatycznie." : ""}`);
  }
}

async function poll() {
  if (polling || (!session && !pairing)) return;
  polling = true;
  const expected = session || pairing, gen = generation;
  try {
    await verifyIdentity(expected, gen);
    if (!session) {
      const result = await request("pair-status", {pairing_id: expected.pairing_id, document: expected.document}, expected.secret);
      if (generation !== gen) return;
      if (result.status === "connected") {
        session = {session_id: result.session_id, capability: result.capability, document: expected.document};
        pairing = null; renderAccess();
        status("Połączono. Codex i Claude Code mogą wykonywać zlecone zmiany w tym dokumencie. Dostęp możesz cofnąć w każdej chwili.");
      }
      return;
    }
    const response = await api("poll", {}, expected, gen);
    if (response.command && generation === gen) await execute(response.command, expected, gen);
  } catch (error) {
    if (generation === gen) {
      if (error.status === 400 || error.status === 401) resetAccess("Dostęp wygasł lub mostek został uruchomiony ponownie. Udostępnij dokument ponownie.");
      else status(error.message);
    }
  } finally { polling = false; }
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Word || !Office.context.requirements.isSetSupported("WordApi", "1.6")) {
    status("Wymagany jest Word Microsoft 365 z WordApi 1.6."); $("connect").disabled = true; return;
  }
  $("connect").addEventListener("click", grantAccess);
  $("disconnect").addEventListener("click", revokeAccess);
  renderAccess();
  status("Zapisz dokument i kliknij „Udostępnij dokument AI”. Nie musisz zaznaczać tekstu.");
  window.setInterval(poll, 1200);
});
