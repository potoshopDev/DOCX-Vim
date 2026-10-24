import assert from "node:assert/strict";
import test from "node:test";
import { draftNeedsWrite, localNeedsWrite } from "../src/save-decision.ts";
import type { Draft } from "../src/draft.ts";

const original: Draft = {
  version: 1, id: "00000000-0000-0000-0000-000000000001", name: "Текст.docx",
  document: { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Привет" }] }] },
  updatedAt: "2026-10-24T00:00:00Z", sourcePath: "C:/Текст.docx", savePath: "C:/Текст.docx",
  sourceFingerprint: "file-revision", localSavePath: "C:/Текст.docx", localSaveContentHash: "doc-hash",
};

test("focus, cursor and timestamp changes do not write draft or DOCX", () => {
  const focused = { ...original, updatedAt: "2026-10-24T00:01:00Z", selection: { anchor: 3, head: 3 }, scrollTop: 850 };
  assert.equal(draftNeedsWrite(original, focused), false);
  assert.equal(localNeedsWrite(focused, "doc-hash"), false);
});

test("text and formatting changes are durable changes", () => {
  const changedText = { ...original, document: { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Привет!" }] }] } };
  const changedFormat = { ...original, document: { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Привет", marks: [{ type: "bold" }] }] }] } };
  assert.equal(draftNeedsWrite(original, changedText), true);
  assert.equal(draftNeedsWrite(original, changedFormat), true);
  assert.equal(localNeedsWrite(changedText, "new-doc-hash"), true);
});

test("failed or redirected save is retried despite unchanged text", () => {
  assert.equal(localNeedsWrite({ ...original, localSaveContentHash: undefined }, "doc-hash"), true);
  assert.equal(localNeedsWrite({ ...original, savePath: "D:/copy.docx" }, "doc-hash"), true);
  assert.equal(draftNeedsWrite(original, { ...original, savePath: "D:/copy.docx" }), true);
});

test("undo back to saved content does not rewrite DOCX", () => {
  const changed = { ...original, document: { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Другое" }] }] } };
  assert.equal(draftNeedsWrite(changed, original), true);
  assert.equal(localNeedsWrite(original, "doc-hash"), false);
});
