import type { JSONContent } from "@tiptap/core";
import { isTauri } from "@tauri-apps/api/core";
import { BaseDirectory, mkdir, rename, writeFile } from "@tauri-apps/plugin-fs";
import { writeDraft, type Draft } from "./draft";

type Snapshot = { draft: Draft; token: number };
type Waiter = { token: number; resolve: () => void; reject: (error: Error) => void };
type Entry = { latest?: Snapshot; timer?: number; running: boolean; nextToken: number; waiters: Waiter[] };

export class BackgroundSaveFailure extends Error {
  constructor(readonly draftSaved: boolean) { super("Background save failed"); }
}

function exportInWorker(document: JSONContent): Promise<Uint8Array> {
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
  constructor(private readonly onSaved: (id: string, path: string | undefined) => void,
    private readonly delayMs = 800) {}

  private entry(id: string): Entry {
    let entry = this.entries.get(id);
    if (!entry) {
      entry = { running: false, nextToken: 0, waiters: [] };
      this.entries.set(id, entry);
    }
    return entry;
  }

  schedule(draft: Draft): void {
    const entry = this.entry(draft.id);
    entry.latest = { draft, token: ++entry.nextToken };
    if (entry.timer) window.clearTimeout(entry.timer);
    entry.timer = window.setTimeout(() => { entry.timer = undefined; void this.run(draft.id); }, this.delayMs);
  }

  flush(draft: Draft): Promise<void> {
    const entry = this.entry(draft.id);
    const token = ++entry.nextToken;
    entry.latest = { draft, token };
    if (entry.timer) window.clearTimeout(entry.timer);
    entry.timer = undefined;
    const result = new Promise<void>((resolve, reject) => entry.waiters.push({ token, resolve, reject }));
    void this.run(draft.id);
    return result;
  }

  private async run(id: string): Promise<void> {
    const entry = this.entry(id);
    if (entry.running) return;
    entry.running = true;
    try {
      while (entry.latest) {
        const snapshot = entry.latest;
        entry.latest = undefined;
        let draftSaved = false;
        try {
          await writeDraft(snapshot.draft);
          draftSaved = true;
          if (snapshot.draft.savePath && isTauri()) {
            const bytes = await exportInWorker(snapshot.draft.document);
            await mkdir("docx-backups", { baseDir: BaseDirectory.AppData, recursive: true });
            const backup = `docx-backups/${id}.docx`;
            const temporary = `${backup}.tmp`;
            await writeFile(temporary, bytes, { baseDir: BaseDirectory.AppData });
            await rename(temporary, backup, { oldPathBaseDir: BaseDirectory.AppData, newPathBaseDir: BaseDirectory.AppData });
            await writeFile(snapshot.draft.savePath, bytes);
          }
          if (!entry.latest) this.onSaved(id, snapshot.draft.savePath);
          entry.waiters = entry.waiters.filter((waiter) => {
            if (waiter.token > snapshot.token) return true;
            waiter.resolve();
            return false;
          });
        } catch (error) {
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
    } finally { entry.running = false; }
  }
}
