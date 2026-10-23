import { isTauri } from "@tauri-apps/api/core";
import { BaseDirectory, mkdir, readTextFile, rename, writeTextFile } from "@tauri-apps/plugin-fs";

export type Language = "ru" | "en";
export type PersonalDictionary = Record<Language, string[]>;
const path = "library/personal-dictionary.json";
const key = "docx-vim-personal-dictionary";
const empty = (): PersonalDictionary => ({ ru: [], en: [] });
let writeQueue = Promise.resolve();

export async function readPersonalDictionary(): Promise<PersonalDictionary> {
  try {
    const raw = isTauri() ? await readTextFile(path, { baseDir: BaseDirectory.AppData }) : localStorage.getItem(key);
    const value = JSON.parse(raw || "{}");
    const words = empty();
    for (const language of ["ru", "en"] as const) {
      if (Array.isArray(value[language])) words[language] = [...new Set(value[language].filter((word: unknown) => typeof word === "string" && /^[\p{L}ёЁ'-]{2,64}$/u.test(word)).map((word: string) => word.toLocaleLowerCase(language)))].sort();
    }
    return words;
  } catch { return empty(); }
}

export function writePersonalDictionary(words: PersonalDictionary): Promise<void> {
  const raw = JSON.stringify(words);
  writeQueue = writeQueue.catch(() => {}).then(async () => {
    if (isTauri()) {
      await mkdir("library", { baseDir: BaseDirectory.AppData, recursive: true });
      await writeTextFile(`${path}.tmp`, raw, { baseDir: BaseDirectory.AppData });
      await rename(`${path}.tmp`, path, { oldPathBaseDir: BaseDirectory.AppData, newPathBaseDir: BaseDirectory.AppData });
    } else localStorage.setItem(key, raw);
    dispatchEvent(new CustomEvent("personal-dictionary-changed", { detail: words }));
  });
  return writeQueue;
}
