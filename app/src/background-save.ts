import type { JSONContent } from "@tiptap/core";
import { isTauri } from "@tauri-apps/api/core";
import { BaseDirectory, mkdir, readFile, rename, writeFile } from "@tauri-apps/plugin-fs";
import { writeDraft, type Draft } from "./draft";
import { fingerprint } from "./source-version";

type Snapshot = { draft: Draft; token: number };
type Waiter = { token: number; resolve: () => void; reject: (error: Error) => void };
type Entry = { latest?: Snapshot; timer?: number; running: boolean; suspended: boolean; idle: (() => void)[]; nextToken: number; waiters: Waiter[] };

export class LocalSourceConflict extends Error {}

export class BackgroundSaveFailure extends Error {
  constructor(readonly draftSaved: boolean) { super("Background save failed"); }
}

export function exportInWorker(document: JSONContent): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./docx-export.worker.ts", import.meta.url), { type: "module" });
    worker.onmessage = (event: MessageEvent<{ ok: boolean; bytes?: ArrayBuffer }>) => {
      worker.terminate();
      if (event.data.ok && event.data.bytes) resolve(new Uint8Array(event.data.bytes));
      else reject(new Error("DOCX export failed"));
    };
    worker.onerror = () => { worker.terminate(); reject(new Error("DOCX worker failed")); };
    worker.postMessage(document);
  });
}

export class BackgroundSaver {
  private entries = new Map<string, Entry>();
  private revisions = new Map<string, string>();
  constructor(private readonly onSaved: (draft: Draft) => void,
    private readonly delayMs = 800,
    private readonly onSnapshot?: (draft: Draft, bytes: Uint8Array) => void,
    private readonly onConflict?: (draft: Draft) => void,
    private readonly onSourceSaved?: (revision: string) => void) {}

  private entry(id: string): Entry {
    let entry = this.entries.get(id);
    if (!entry) {
      entry = { running: false, suspended: false, idle: [], nextToken: 0, waiters: [] };
      this.entries.set(id, entry);
    }
    return entry;
  }

  schedule(draft: Draft): void {
    const entry = this.entry(draft.id);
    if (entry.suspended) return;
    entry.latest = { draft, token: ++entry.nextToken };
    if (entry.timer) window.clearTimeout(entry.timer);
    entry.timer = window.setTimeout(() => { entry.timer = undefined; void this.run(draft.id); }, this.delayMs);
  }

  flush(draft: Draft): Promise<void> {
    const entry = this.entry(draft.id);
    if (entry.suspended) return Promise.reject(new LocalSourceConflict("Source changed"));
    const token = ++entry.nextToken;
    entry.latest = { draft, token };
    if (entry.timer) window.clearTimeout(entry.timer);
    entry.timer = undefined;
    const result = new Promise<void>((resolve, reject) => entry.waiters.push({ token, resolve, reject }));
    void this.run(draft.id);
    return result;
  }

  async suspend(id: string): Promise<void> {
    const entry = this.entry(id);
    entry.suspended = true;
    entry.latest = undefined;
    if (entry.timer) window.clearTimeout(entry.timer);
    entry.timer = undefined;
    for (const waiter of entry.waiters) waiter.reject(new LocalSourceConflict("Source changed"));
    entry.waiters = [];
    if (entry.running) await new Promise<void>((resolve) => entry.idle.push(resolve));
  }

  private async run(id: string): Promise<void> {
    const entry = this.entry(id);
    if (entry.running) return;
    entry.running = true;
    try {
      while (entry.latest && !entry.suspended) {
        const snapshot = entry.latest;
        entry.latest = undefined;
        let draftSaved = false;
        try {
          await writeDraft(snapshot.draft);
          draftSaved = true;
          if ((snapshot.draft.savePath || Object.values(snapshot.draft.cloudTargets ?? {}).some(Boolean)) && isTauri()) {
            const bytes = await exportInWorker(snapshot.draft.document);
            await mkdir("docx-backups", { baseDir: BaseDirectory.AppData, recursive: true });
            const backup = `docx-backups/${id}.docx`;
            const temporary = `${backup}.tmp`;
            await writeFile(temporary, bytes, { baseDir: BaseDirectory.AppData });
            await rename(temporary, backup, { oldPathBaseDir: BaseDirectory.AppData, newPathBaseDir: BaseDirectory.AppData });
            if (snapshot.draft.savePath) {
              const mainLocalSource = snapshot.draft.sourcePath === snapshot.draft.savePath;
              if (mainLocalSource) {
                const expected = this.revisions.get(id) ?? snapshot.draft.sourceFingerprint;
                if (!expected) throw new LocalSourceConflict("Local source revision unknown");
                if (expected !== "new" && await fingerprint(await readFile(snapshot.draft.savePath)) !== expected)
                  throw new LocalSourceConflict("Local source changed");
              }
              await writeFile(snapshot.draft.savePath, bytes);
              if (mainLocalSource) {
                const revision = await fingerprint(bytes);
                this.revisions.set(id, revision);
                snapshot.draft.sourceFingerprint = revision;
                await writeDraft(snapshot.draft);
                this.onSourceSaved?.(revision);
              }
            }
            this.onSnapshot?.(snapshot.draft, bytes);
          }
          if (!entry.latest) this.onSaved(snapshot.draft);
          entry.waiters = entry.waiters.filter((waiter) => {
            if (waiter.token > snapshot.token) return true;
            waiter.resolve();
            return false;
          });
        } catch (error) {
          if (error instanceof LocalSourceConflict) {
            entry.suspended = true;
            entry.latest = undefined;
            if (entry.timer) window.clearTimeout(entry.timer);
            entry.timer = undefined;
            for (const waiter of entry.waiters) waiter.reject(error);
            entry.waiters = [];
            this.onConflict?.(snapshot.draft);
            break;
          }
          const failure = new BackgroundSaveFailure(draftSaved);
          entry.waiters = entry.waiters.filter((waiter) => {
            if (waiter.token > snapshot.token) return true;
            waiter.reject(failure);
            return false;
          });
          if (!entry.latest) {
            entry.latest = snapshot;
            entry.timer = window.setTimeout(() => {
              entry.timer = undefined;
              void this.run(id);
            }, 15_000);
            break;
          }
        }
      }
    } finally { entry.running = false; entry.idle.splice(0).forEach((resolve) => resolve()); }
  }
}
