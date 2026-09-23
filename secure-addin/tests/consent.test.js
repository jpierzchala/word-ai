import test from "node:test";
import assert from "node:assert/strict";
import {DOMParser, XMLSerializer} from "@xmldom/xmldom";
import {W} from "../safety.js";
import {paragraphDigest} from "../document-edit.js";
const xml = text => `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`;
let number = 0;
async function boot() {
  const state = {text: "before", writes: 0, command: null, results: [], cancelled: false, url: "https://test/document.docx", pictures: []};
  const elements = new Map();
  const element = id => {
    if (!elements.has(id)) elements.set(id, {listeners: {}, addEventListener(n, f) {this.listeners[n] = f;}});
    return elements.get(id);
  };
  globalThis.DOMParser = DOMParser; globalThis.XMLSerializer = XMLSerializer;
  globalThis.document = {getElementById: element};
  globalThis.window = {setInterval: fn => {state.tick = fn;}};
  globalThis.Office = {AsyncResultStatus: {Succeeded: "ok"}, HostType: {Word: "Word"},
    context: {requirements: {isSetSupported: () => true}, document: {getFilePropertiesAsync: cb => cb({status: "ok", value: {url: state.url}})}}, onReady: fn => fn({host: "Word"})};
  const pictureCollection = {load() {}, get items() {return state.pictures;}};
  const range = {load() {}, parentContentControlOrNullObject: {isNullObject: true, load() {}}, inlinePictures: pictureCollection,
    get text() {return state.text;}, getOoxml: () => { if (state.exportFails) throw Object.assign(Error("Export failed"), {code: "GeneralException"}); return {value: state.exportXml || xml(state.text)}; },
    insertText(text) {state.text = text; state.writes++;}};
  const paragraph = {uniqueLocalId: "11111111-1111-1111-1111-111111111111", styleBuiltIn: "Normal", load() {}, get text() {return state.text;}, getRange: () => range,
    listItemOrNullObject: {isNullObject: true, load() {}}};
  const ctx = {sync: async () => {if (state.onSync) await state.onSync();}, document: {load() {}, changeTrackingMode: "Off", getParagraphByUniqueLocalId: () => paragraph}};
  ctx.document.body = {paragraphs: {load() {}, items: [paragraph]}};
  state.ctx = ctx;
  const imageParagraph = {uniqueLocalId: "22222222-2222-2222-2222-222222222222", load() {},
    insertParagraph(text) {state.writes++; return {text, load() {}};},
    insertInlinePictureFromBase64() { if (state.imageFailure) throw Error("Partial image failure"); state.writes++; return picture; }};
  const picture = {paragraph: imageParagraph, load() {}};
  const existingPicture = {paragraph, width: 144, height: 72, altTextTitle: "Old title", altTextDescription: "Old alt", load() {},
    parentContentControlOrNullObject: {isNullObject: true, load() {}},
    insertInlinePictureFromBase64() {if (state.imageFailure) throw Error("Partial image failure"); state.writes++;
      const replacement = {...existingPicture, paragraph, load() {}}; state.pictures[0] = replacement;
      if (state.afterReplacementXml) state.exportXml = state.afterReplacementXml;
      return replacement;},
    delete() {state.writes++; state.pictures.splice(state.pictures.indexOf(existingPicture), 1);}};
  state.existingPicture = existingPicture;
  paragraph.insertParagraph = () => {state.writes++; return imageParagraph;};
  globalThis.createImageBitmap = async () => {if (state.onDecode) await state.onDecode(); return {width: 2, height: 1, close() {}};};
  globalThis.Word = {run: fn => fn(ctx)};
  globalThis.fetch = async (url, options) => {
    const body = JSON.parse(options.body); let data = {}, ok = true;
    if (url === "/office/pair") data = {pairing_id: "p1", secret: "ticket"};
    if (url === "/office/pair-status") data = {status: "connected", session_id: "s1", capability: "c1"};
    if (url === "/office/poll") {data = {command: state.command}; state.command = null;}
    if (url === "/office/begin") {ok = !state.cancelled; if (state.mutateAtBegin) state.text = "coauthor";}
    if (url === "/office/result") state.results.push(body.result);
    return {ok, status: ok ? 200 : 400, json: async () => data};
  };
  await import(`../taskpane.js?test=${++number}`);
  state.click = id => element(id).listeners.click();
  await state.click("connect"); await state.tick();
  state.prepare = async () => {state.command = {command_id: "c1", type: "apply", payload: {paragraph_id: paragraph.uniqueLocalId,
    operation: "replace_paragraph", text: "after", find: "", expected_sha256: await paragraphDigest(xml("before"))}};};
  return state;
}
test("document session applies requested edit without a second UI confirmation", async () => {
  const s = await boot(); await s.prepare(); await s.tick();
  assert.equal(s.writes, 1); assert.equal(s.results[0].ok, true); await s.tick(); assert.equal(s.writes, 1);
});
test("stale preview rejects write", async () => {
  const s = await boot(); await s.prepare(); s.mutateAtBegin = true; await s.tick();
  assert.equal(s.writes, 0); assert.equal(s.results[0].submitted, false);
});
test("cancelled command never writes", async () => {
  const s = await boot(); await s.prepare(); s.cancelled = true; await s.tick(); assert.equal(s.writes, 0);
});
test("revoke before poll prevents writes", async () => {
  const s = await boot(); await s.prepare(); await s.click("disconnect"); await s.tick(); assert.equal(s.writes, 0);
});
test("revoke during Word sync prevents writes", async () => {
  const s = await boot(); await s.prepare(); s.onSync = async () => {s.onSync = null; await s.click("disconnect");};
  await s.tick(); assert.equal(s.writes, 0);
});
test("document identity change revokes before writing", async () => {
  const s = await boot(); await s.prepare(); s.url = "https://test/other.docx"; await s.tick(); assert.equal(s.writes, 0);
});
test("tracking enabled blocks edits without altering tracking", async () => {
  const s = await boot(); await s.prepare(); s.ctx.document.changeTrackingMode = "TrackAll"; await s.tick();
  assert.equal(s.writes, 0); assert.equal(s.ctx.document.changeTrackingMode, "TrackAll");
});

test("snapshot reports failed XML export without claiming full coverage or losing readable text", async () => {
  const s = await boot(); s.exportFails = true;
  s.command = {command_id: "snapshot1", type: "snapshot", payload: {scope: "body", start: 0, limit: 10, query: ""}};
  await s.tick(); const result = s.results[0];
  assert.equal(result.ok, true); assert.equal(result.complete, false);
  assert.equal(result.paragraphs[0].text, "before"); assert.equal(result.paragraphs[0].ooxml_sha256, null);
  assert.equal(result.coverage_errors[0].error, "GeneralException"); assert.equal(s.writes, 0);
});

async function prepareImage(s) {
  const encoded = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAD0In+KAAAAD0lEQVR4nGP4z8DA8J8BAAf/Af9jgLz7AAAAAElFTkSuQmCC";
  const sha256 = Buffer.from(await crypto.subtle.digest("SHA-256", Buffer.from(encoded, "base64"))).toString("hex");
  s.command = {command_id: "image1", type: "apply", payload: {paragraph_id: "11111111-1111-1111-1111-111111111111",
    operation: "insert_image", position: "after", expected_sha256: await paragraphDigest(xml("before")), image_base64: encoded,
    image: {mime: "image/png", sha256, pixel_width: 2, pixel_height: 1}, image_name: "test.png",
    width_pt: 144, height_pt: 72, caption: "Caption", alt_text: "Test image"}};
}
test("session inserts image and caption with verified dimensions without another approval", async () => {
  const s = await boot(); await prepareImage(s); await s.tick();
  assert.equal(s.writes, 3); assert.equal(s.results[0].ok, true);
  assert.equal(s.results[0].width_pt, 144); assert.equal(s.results[0].caption, "Caption");
});
test("revocation during image decoding prevents any mutation", async () => {
  const s = await boot(); await prepareImage(s); s.onDecode = () => s.click("disconnect");
  await s.tick(); assert.equal(s.writes, 0);
});
test("partial image failure is reported as submitted and must not be retried", async () => {
  const s = await boot(); await prepareImage(s); s.imageFailure = true; await s.tick();
  assert.equal(s.writes, 1); assert.equal(s.results[0].ok, false); assert.equal(s.results[0].submitted, true);
});

test("snapshot reports structural parsing failures as incomplete coverage", async () => {
  const s = await boot(); s.exportXml = " ".repeat(8 * 1024 * 1024 + 1);
  s.command = {command_id: "snapshot-large", type: "snapshot", payload: {scope: "body", start: 0, limit: 10, query: ""}};
  await s.tick();
  assert.equal(s.results[0].complete, false);
  assert.equal(s.results[0].paragraphs[0].ooxml_sha256, null);
  assert.equal(s.results[0].coverage_errors[0].api, "paragraphDigest");
});

async function prepareImageChange(s, operation) {
  const encoded = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAD0In+KAAAAD0lEQVR4nGP4z8DA8J8BAAf/Af9jgLz7AAAAAElFTkSuQmCC";
  const sha256 = Buffer.from(await crypto.subtle.digest("SHA-256", Buffer.from(encoded, "base64"))).toString("hex");
  s.pictures.push(s.existingPicture);
  s.command = {command_id: "image-change", type: "apply", payload: {paragraph_id: "11111111-1111-1111-1111-111111111111",
    operation, image_index: 0, expected_sha256: await paragraphDigest(xml("before")), ...(operation === "replace_image" ? {
      image_base64: encoded, image: {mime: "image/png", sha256, pixel_width: 2, pixel_height: 1}, image_name: "new.png",
      width_pt: 0, alt_text: "New alt"} : {})}};
}
test("session replaces one existing image in place", async () => {
  const s = await boot(); await prepareImageChange(s, "replace_image"); await s.tick();
  assert.equal(s.writes, 1); assert.equal(s.results[0].ok, true);
  assert.equal(s.results[0].operation, "replace_image"); assert.equal(s.results[0].image_count, 1);
});
test("session deletes only the selected image", async () => {
  const s = await boot(); await prepareImageChange(s, "delete_image"); await s.tick();
  assert.equal(s.writes, 1); assert.equal(s.results[0].ok, true);
  assert.equal(s.results[0].operation, "delete_image"); assert.equal(s.results[0].image_count, 0);
});

test("replacement verification accepts a Flat OPC export containing an allowed 2 MiB image", async () => {
  const s = await boot(); await prepareImageChange(s, "replace_image");
  // Word's range export includes media as Base64 in addition to the paragraph XML.
  s.afterReplacementXml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${xml("before")}</pkg:xmlData></pkg:part><pkg:part pkg:name="/word/media/image.png"><pkg:binaryData>${Buffer.alloc(2 * 1024 * 1024).toString("base64")}</pkg:binaryData></pkg:part></pkg:package>`;
  await s.tick();
  assert.equal(s.writes, 1);
  assert.equal(s.results[0].ok, true, JSON.stringify(s.results[0]));
  assert.match(s.results[0].ooxml_sha256, /^[0-9a-f]{64}$/);
});
