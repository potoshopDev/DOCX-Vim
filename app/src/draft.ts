import type { JSONContent } from "@tiptap/core";
import { isTauri } from "@tauri-apps/api/core";
import { BaseDirectory, mkdir, readDir, readTextFile, rename, writeTextFile } from "@tauri-apps/plugin-fs";

export type Draft = {
  version: 1;
  id: string;
  name: string;
  document: JSONContent;
  updatedAt: string;
  sourcePath?: string;
  savePath?: string;
  viewMode?: "page" | "mobile";
  selection?: { anchor: number; head: number };
  scrollTop?: number;
};
const directory = "drafts";
const key = (id: string) => `docx-vim-draft-${id}`;
const path = (id: string) => {
  if (!/^[a-f\d-]{36}$/i.test(id)) throw new Error("Неверный ID документа");
  return `${directory}/${id}.json`;
};

export async function readDraft(id: string): Promise<Draft | null> {
  try {
    const raw = isTauri()
      ? await readTextFile(path(id), { baseDir: BaseDirectory.AppData })
      : localStorage.getItem(key(id));
    if (!raw) return null;
    const draft = JSON.parse(raw) as Draft;
    return draft.version === 1 && draft.id === id && draft.document?.type === "doc" ? draft : null;
  } catch { return null; }
}

export async function listDrafts(): Promise<Draft[]> {
  let ids: string[];
  if (isTauri()) {
    try {
      ids = (await readDir(directory, { baseDir: BaseDirectory.AppData }))
        .map((entry) => entry.name.replace(/\.json$/, ""))
        .filter((id) => /^[a-f\d-]{36}$/i.test(id));
    } catch { return []; }
  } else {
    ids = Object.keys(localStorage)
      .filter((entry) => entry.startsWith("docx-vim-draft-"))
      .map((entry) => entry.slice("docx-vim-draft-".length));
  }
  return (await Promise.all(ids.map(readDraft))).filter((draft): draft is Draft => draft !== null)
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function writeDraft(draft: Draft): Promise<void> {
  const raw = JSON.stringify(draft);
  if (isTauri()) {
    await mkdir(directory, { baseDir: BaseDirectory.AppData, recursive: true });
    const destination = path(draft.id);
    const temporary = `${destination}.tmp`;
    await writeTextFile(temporary, raw, { baseDir: BaseDirectory.AppData });
    await rename(temporary, destination, { oldPathBaseDir: BaseDirectory.AppData, newPathBaseDir: BaseDirectory.AppData });
  } else localStorage.setItem(key(draft.id), raw);
}
