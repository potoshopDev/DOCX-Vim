import { BaseDirectory, mkdir, readFile, readTextFile, rename, writeFile, writeTextFile } from "@tauri-apps/plugin-fs";
import { readDraft, type Draft } from "./draft";
import { preserveDraftVersion } from "./source-version";
import { isRemoteTarget, RemoteConflict, uploadRemote, verifyRemoteRevision } from "./cloud-remote";

export type CloudProvider = "oneDrive" | "yandexDisk";
export const cloudProviders: CloudProvider[] = ["oneDrive", "yandexDisk"];
export const cloudName: Record<CloudProvider, string> = { oneDrive: "OneDrive", yandexDisk: "Яндекс Диск" };
export type CloudStatus = "pending" | "writing" | "saved" | "retry" | "conflict";
export type CloudState = { path: string; hash?: string; revision?: string; status: CloudStatus; conflictCopy?: string; conflictDraftId?: string; updatedAt?: string; savedDraftUpdatedAt?: string };
export type CloudStates = Partial<Record<CloudProvider, CloudState>>;
type Pending = { path: string; bytes: Uint8Array; draftUpdatedAt: string };
type Queue = { pending?: Pending; running: boolean; timer?: number; attempts: number };

const directory = "cloud-sync";
const retryDelay = (attempts: number) => Math.min(5 * 60_000, 5_000 * 2 ** Math.min(attempts, 6));
const filename = (id: string) => {
  if (!/^[a-f\d-]{36}$/i.test(id)) throw new Error("Неверный ID документа");
  return `${directory}/${id}.json`;
};

async function sha256(bytes: Uint8Array): Promise<string> {
  const hash = await crypto.subtle.digest("SHA-256", bytes.slice().buffer);
  return Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export async function readCloudStates(id: string): Promise<CloudStates> {
  try { return JSON.parse(await readTextFile(filename(id), { baseDir: BaseDirectory.AppData })) as CloudStates; }
  catch { return {}; }
}

async function writeCloudStates(id: string, states: CloudStates): Promise<void> {
  await mkdir(directory, { baseDir: BaseDirectory.AppData, recursive: true });
  const temporary = `${filename(id)}.tmp`;
  await writeTextFile(temporary, JSON.stringify(states), { baseDir: BaseDirectory.AppData });
  await rename(temporary, filename(id), { oldPathBaseDir: BaseDirectory.AppData, newPathBaseDir: BaseDirectory.AppData });
}

export async function rememberOpenedRemote(id: string, provider: CloudProvider, path: string, revision: string): Promise<void> {
  const states = await readCloudStates(id);
  states[provider] = { path, revision, status: "saved", updatedAt: new Date().toISOString() };
  await writeCloudStates(id, states);
}

export class CloudSync {
  private queues = new Map<CloudProvider, Queue>();
  private retired = false;
  private states: CloudStates = {};
  private loaded = false;
  private loading?: Promise<void>;
  private writing: Promise<void> = Promise.resolve();
  constructor(private readonly id: string, private readonly onChange: (states: CloudStates) => void) {}

  async load(): Promise<void> {
    if (this.loaded) return;
    if (!this.loading) this.loading = readCloudStates(this.id).then((states) => {
      this.states = states;
      this.loaded = true;
      this.onChange({ ...this.states });
    });
    await this.loading;
  }

  schedule(draft: Draft, bytes: Uint8Array): void {
    if (this.retired) return;
    if (!this.loaded) { void this.load().then(() => this.schedule(draft, bytes)); return; }
    for (const provider of cloudProviders) {
      const path = draft.cloudTargets?.[provider];
      if (!path) continue;
      let queue = this.queues.get(provider);
      if (!queue) { queue = { running: false, attempts: 0 }; this.queues.set(provider, queue); }
      queue.pending = { path, bytes, draftUpdatedAt: draft.updatedAt };
      const state = this.states[provider];
      if (!state || state.path !== path) this.states[provider] = { path, status: "pending" };
      else if (state.status !== "conflict") this.states[provider] = { ...state, status: "pending" };
      this.onChange({ ...this.states });
      if (queue.timer) clearTimeout(queue.timer);
      queue.timer = window.setTimeout(() => { queue!.timer = undefined; void this.run(provider); }, 250);
    }
  }

  disconnect(provider: CloudProvider): void {
    const queue = this.queues.get(provider);
    if (queue?.timer) clearTimeout(queue.timer);
    if (queue) queue.pending = undefined;
    delete this.states[provider];
    this.onChange({ ...this.states });
    void this.persist();
  }

  rebind(provider: CloudProvider, path: string): void {
    const queue = this.queues.get(provider);
    if (queue?.timer) clearTimeout(queue.timer);
    if (queue) { queue.pending = undefined; queue.attempts = 0; }
    this.states[provider] = { path, status: "pending" };
    this.onChange({ ...this.states });
    void this.persist();
  }

  async acceptOpenedRemote(provider: CloudProvider, path: string, revision: string): Promise<void> {
    await this.load();
    const queue = this.queues.get(provider);
    if (queue?.timer) clearTimeout(queue.timer);
    if (queue) { queue.pending = undefined; queue.attempts = 0; }
    this.states[provider] = { path, revision, status: "saved", updatedAt: new Date().toISOString() };
    this.onChange({ ...this.states });
    await this.persist();
  }

  retry(): void {
    if (this.retired) return;
    for (const provider of cloudProviders) {
      const queue = this.queues.get(provider);
      if (queue?.pending && !queue.running) void this.run(provider);
    }
  }

  private async run(provider: CloudProvider): Promise<void> {
    if (this.retired) return;
    const queue = this.queues.get(provider);
    if (!queue || queue.running) return;
    queue.running = true;
    try {
      while (queue.pending && !this.retired) {
        const pending = queue.pending;
        queue.pending = undefined;
        const state = this.states[provider];
        if (!state || state.path !== pending.path) continue;
        try {
          if (state.status === "conflict") {
            await this.keepConflictCopy(provider, state, pending.bytes);
            continue;
          }
          await this.update(provider, { ...state, status: "writing" });
          const newHash = await sha256(pending.bytes);
          if (isRemoteTarget(pending.path)) {
            if (state.hash === newHash && state.revision) {
              await verifyRemoteRevision(provider, pending.path, state.revision);
              await this.update(provider, { ...state, status: "saved", savedDraftUpdatedAt: pending.draftUpdatedAt });
              continue;
            }
            const revision = await uploadRemote(provider, pending.path, pending.bytes, state.revision);
            await this.update(provider, { path: pending.path, hash: newHash, revision, status: "saved", updatedAt: new Date().toISOString(), savedDraftUpdatedAt: pending.draftUpdatedAt });
            queue.attempts = 0;
            continue;
          }
          if (state.hash) {
            const remoteHash = await sha256(await readFile(pending.path));
            if (remoteHash !== state.hash) {
              await this.keepConflictCopy(provider, state, pending.bytes);
              continue;
            }
          }
          if (state.hash !== newHash) await writeFile(pending.path, pending.bytes);
          await this.update(provider, { path: pending.path, hash: newHash, status: "saved", updatedAt: new Date().toISOString(), savedDraftUpdatedAt: pending.draftUpdatedAt });
          queue.attempts = 0;
        } catch (error) {
          if (error instanceof RemoteConflict) {
            try { await this.keepConflictCopy(provider, state, pending.bytes); } catch { queue.pending = queue.pending ?? pending; this.setRetry(provider, queue); }
            continue;
          }
          queue.pending = queue.pending ?? pending;
          try { await this.update(provider, { ...state, status: "retry" }); } catch { /* Retry can recover the ledger too. */ }
          this.setRetry(provider, queue);
          break;
        }
      }
    } finally { queue.running = false; }
  }

  private setRetry(provider: CloudProvider, queue: Queue): void {
    if (this.retired) return;
    if (queue.timer) clearTimeout(queue.timer);
    queue.timer = window.setTimeout(() => { queue.timer = undefined; void this.run(provider); }, retryDelay(++queue.attempts));
  }

  private async keepConflictCopy(provider: CloudProvider, state: CloudState, bytes: Uint8Array): Promise<void> {
    await mkdir("cloud-conflicts", { baseDir: BaseDirectory.AppData, recursive: true });
    const copy = state.conflictCopy ?? `cloud-conflicts/${this.id}-${provider}.docx`;
    await writeFile(copy, bytes, { baseDir: BaseDirectory.AppData });
    let conflictDraftId = state.conflictDraftId;
    if (!conflictDraftId) {
      const draft = await readDraft(this.id);
      if (draft && draft.sourcePath !== state.path) conflictDraftId = (await preserveDraftVersion(draft)).id;
    }
    await this.update(provider, { ...state, status: "conflict", conflictCopy: copy, conflictDraftId });
  }

  private async update(provider: CloudProvider, state: CloudState): Promise<void> {
    if (this.states[provider]?.path !== state.path) return;
    this.states[provider] = state;
    this.onChange({ ...this.states });
    await this.persist();
  }

  private persist(): Promise<void> {
    this.writing = this.writing.catch(() => {}).then(() => writeCloudStates(this.id, { ...this.states }));
    return this.writing;
  }

  async retire(): Promise<void> {
    this.retired = true;
    for (const queue of this.queues.values()) {
      if (queue.timer) clearTimeout(queue.timer);
      queue.timer = undefined;
      queue.pending = undefined;
    }
    while ([...this.queues.values()].some((queue) => queue.running))
      await new Promise((resolve) => window.setTimeout(resolve, 50));
    await this.writing.catch(() => {});
  }
}
