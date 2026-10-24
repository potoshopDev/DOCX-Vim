import { readDraft, writeDraft, type Draft } from "./draft";
import type { JSONContent } from "@tiptap/core";

export async function fingerprint(bytes: Uint8Array): Promise<string> {
  const hash = await crypto.subtle.digest("SHA-256", bytes.slice().buffer as ArrayBuffer);
  return Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export function documentContentHash(document: JSONContent): Promise<string> {
  return fingerprint(new TextEncoder().encode(JSON.stringify(document)));
}

export async function preserveDraftVersion(draft: Draft): Promise<Draft> {
  const version: Draft = {
    ...draft,
    id: crypto.randomUUID(),
    name: `${draft.name.replace(/\.docx$/i, "")} — версия до конфликта`,
    restoreSourcePath: draft.sourcePath ?? draft.restoreSourcePath,
    sourcePath: undefined,
    sourceFingerprint: undefined,
    savePath: undefined,
    localSavePath: undefined,
    localSaveContentHash: undefined,
    cloudTargets: {},
    updatedAt: new Date().toISOString(),
  };
  await writeDraft(version);
  return version;
}

export async function replaceSourceDraft(draft: Draft): Promise<Draft | null> {
  const previous = await readDraft(draft.id);
  const preserved = previous && JSON.stringify(previous.document) !== JSON.stringify(draft.document)
    ? await preserveDraftVersion(previous) : null;
  await writeDraft(draft);
  return preserved;
}
