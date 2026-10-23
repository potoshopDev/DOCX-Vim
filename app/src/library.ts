import { isTauri } from "@tauri-apps/api/core";
import { BaseDirectory, mkdir, readTextFile, rename, writeTextFile } from "@tauri-apps/plugin-fs";

export type RecentDocument = { id: string; name: string; sourcePath?: string; openedAt: string };
const file = "library/recent.json";
const storageKey = "docx-vim-recent";

export async function listRecent(): Promise<RecentDocument[]> {
  try {
    const raw = isTauri() ? await readTextFile(file, { baseDir: BaseDirectory.AppData }) : localStorage.getItem(storageKey);
    const entries = JSON.parse(raw || "[]");
    return Array.isArray(entries) ? entries.filter((item): item is RecentDocument =>
      typeof item?.id === "string" && typeof item?.name === "string" && typeof item?.openedAt === "string") : [];
  } catch { return []; }
}

export async function rememberDocument(item: Omit<RecentDocument, "openedAt">): Promise<void> {
  const entries = (await listRecent()).filter((entry) => entry.id !== item.id &&
    (!item.sourcePath || entry.sourcePath !== item.sourcePath));
  entries.unshift({ ...item, openedAt: new Date().toISOString() });
  const raw = JSON.stringify(entries.slice(0, 100));
  if (isTauri()) {
    await mkdir("library", { baseDir: BaseDirectory.AppData, recursive: true });
    await writeTextFile(`${file}.tmp`, raw, { baseDir: BaseDirectory.AppData });
    await rename(`${file}.tmp`, file, { oldPathBaseDir: BaseDirectory.AppData, newPathBaseDir: BaseDirectory.AppData });
  } else localStorage.setItem(storageKey, raw);
}
