import type { Draft } from "./draft";

// Cursor movement and window focus never change the durable document revision.
export function draftNeedsWrite(previous: Draft | null, next: Draft): boolean {
  if (!previous) return true;
  return JSON.stringify(previous.document) !== JSON.stringify(next.document) ||
    previous.name !== next.name || previous.sourcePath !== next.sourcePath ||
    previous.restoreSourcePath !== next.restoreSourcePath ||
    previous.sourceFingerprint !== next.sourceFingerprint ||
    previous.savePath !== next.savePath || previous.localSavePath !== next.localSavePath ||
    previous.localSaveContentHash !== next.localSaveContentHash ||
    previous.viewMode !== next.viewMode ||
    JSON.stringify(previous.cloudTargets ?? {}) !== JSON.stringify(next.cloudTargets ?? {});
}

export function localNeedsWrite(draft: Draft, documentHash: string): boolean {
  return !!draft.savePath &&
    (draft.localSavePath !== draft.savePath || draft.localSaveContentHash !== documentHash);
}
