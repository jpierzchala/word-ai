import test from "node:test";
import assert from "node:assert/strict";
import {DOMParser} from "@xmldom/xmldom";
import {W} from "../safety.js";
import {paragraphDigest, paragraphXml, planEdit, planImage, planImageChange, MAX_PARAGRAPH_XML_CHARS} from "../document-edit.js";
globalThis.DOMParser = DOMParser;
const xml = body => `<w:document xmlns:w="${W}"><w:body><w:p>${body}</w:p></w:body></w:document>`;
const run = (text, properties="") => `<w:r>${properties}<w:t>${text}</w:t></w:r>`;
const mixed = xml(run("one ") + run("two", "<w:rPr><w:b/></w:rPr>"));
const current = {xml: mixed, hash: "hash", text: "one two", style: "Normal"};
const payload = {expected_sha256: "hash", operation: "replace_text", find: "two", text: "three"};
test("mixed formatting supports targeted replacement but not whole paragraph flattening", () => {
  assert.equal(planEdit(current, payload).after, "one three");
  assert.throws(() => planEdit(current, {...payload, operation: "replace_paragraph"}), /formatowania/);
});
test("fields, comments, revisions and bookmarks prevent text replacement", () => {
  for (const node of ["<w:fldSimple/>", "<w:commentRangeStart/>", "<w:ins/>", "<w:bookmarkStart/>"])
    assert.throws(() => planEdit({...current, xml: xml(run("one two") + node)}, payload), /zablokowany/);
});
test("stale hash and ambiguous literal are rejected", () => {
  assert.throws(() => planEdit(current, {...payload, expected_sha256: "stale"}), /zmienił/);
  assert.throws(() => planEdit({...current, text: "two two"}, payload), /dokładnie raz/);
});
test("empty paragraphs, insertion and built-in styles are supported", () => {
  assert.equal(paragraphXml(xml("")).plain, true);
  assert.equal(planEdit(current, {...payload, operation: "insert_after"}).after, "three");
  assert.equal(planEdit(current, {...payload, operation: "set_style", text: "Heading1"}).new_style, "Heading1");
});
test("structural hash detects text and formatting changes, ignores proofing markers", async () => {
  const a = xml(run("text"));
  assert.notEqual(await paragraphDigest(a), await paragraphDigest(xml(run("other"))));
  assert.notEqual(await paragraphDigest(a), await paragraphDigest(xml(run("text", "<w:rPr><w:b/></w:rPr>"))));
  assert.equal(await paragraphDigest(a), await paragraphDigest(xml('<w:proofErr w:type="spellStart"/>' + run("text"))));
});

const encodedImage = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAD0In+KAAAAD0lEQVR4nGP4z8DA8J8BAAf/Af9jgLz7AAAAAElFTkSuQmCC";
async function imagePayload() {
  const bytes = Buffer.from(encodedImage, "base64");
  const hash = Buffer.from(await crypto.subtle.digest("SHA-256", bytes)).toString("hex");
  return {operation: "insert_image", position: "after", expected_sha256: "hash", image_base64: encodedImage,
    image_name: "test.png", image: {mime: "image/png", sha256: hash, pixel_width: 2, pixel_height: 1},
    width_pt: 144, height_pt: 72, caption: "Caption", alt_text: "Two colors"};
}
const decodeImage = async () => ({width: 2, height: 1, close() {}});
test("image plan preserves ratio and returns metadata without binary content", async () => {
  const plan = await planImage(current, await imagePayload(), decodeImage);
  assert.equal(plan.width_pt, 144); assert.equal(plan.height_pt, 72);
  assert.equal(plan.caption, "Caption"); assert.ok(!JSON.stringify(plan).includes(encodedImage));
});
test("image plan rejects stale anchors, changed bytes and undecodable images", async () => {
  const payload = await imagePayload();
  await assert.rejects(planImage(current, {...payload, expected_sha256: "stale"}, decodeImage), /zmienił/);
  await assert.rejects(planImage(current, {...payload, image_base64: "YWJj"}, decodeImage), /Dane obrazu/);
  await assert.rejects(planImage(current, payload, async () => {throw Error("decode failure");}), /decode failure/);
});

test("larger media export allowance still rejects oversized XML and DTDs", () => {
  assert.throws(() => paragraphXml(" ".repeat(MAX_PARAGRAPH_XML_CHARS + 1)), /limit wynosi/);
  assert.throws(() => paragraphXml('<!DOCTYPE w:document [<!ENTITY x "unsafe">]>' + xml("")), /Nieobsługiwany/);
});

const pictures = [{width: 200, height: 100, altTextTitle: "Old title", altTextDescription: "Old alt"}];
test("image replacement targets an existing picture and can preserve its width", async () => {
  const payload = {...await imagePayload(), operation: "replace_image", image_index: 0, width_pt: 0, alt_text: "New alt"};
  const plan = await planImageChange(current, payload, pictures, decodeImage);
  assert.equal(plan.width_pt, 200); assert.equal(plan.height_pt, 100);
  assert.equal(plan.previous_image.alt_text, "Old alt"); assert.equal(plan.alt_text, "New alt");
});
test("image deletion previews only the selected picture and rejects stale indices", async () => {
  const plan = await planImageChange(current, {operation: "delete_image", image_index: 0, expected_sha256: "hash"}, pictures, decodeImage);
  assert.equal(plan.image_count_before, 1); assert.equal(plan.previous_image.alt_title, "Old title");
  await assert.rejects(planImageChange(current, {operation: "delete_image", image_index: 1, expected_sha256: "hash"}, pictures, decodeImage), /nie istnieje/);
});
