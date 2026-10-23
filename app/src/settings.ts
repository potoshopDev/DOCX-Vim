import { isTauri } from "@tauri-apps/api/core";
import { BaseDirectory, mkdir, readTextFile, rename, writeTextFile } from "@tauri-apps/plugin-fs";

export type Settings = {
  autosaveMs: number;
  language: "ru" | "en";
  vimLayout: "physical" | "translated";
  defaultView: "mobile" | "page";
  vimEnabled: boolean;
  hideMenu: boolean;
  autocorrect: boolean;
  autocomplete: boolean;
  grammarUrl: string;
};
export const defaultSettings: Settings = {
  autosaveMs: 800, language: "ru", vimLayout: "physical", defaultView: "mobile", vimEnabled: true, hideMenu: false,
  autocorrect: true, autocomplete: true, grammarUrl: "",
};
const file = "library/settings.json";
const storageKey = "docx-vim-settings";
let pendingWrite: Promise<void> = Promise.resolve();

export async function readSettings(): Promise<Settings> {
  try {
    const raw = isTauri() ? await readTextFile(file, { baseDir: BaseDirectory.AppData }) : localStorage.getItem(storageKey);
    const value = JSON.parse(raw || "{}");
    return {
      autosaveMs: [800, 2000, 5000].includes(value.autosaveMs) ? value.autosaveMs : defaultSettings.autosaveMs,
      language: value.language === "en" ? "en" : "ru",
      vimLayout: value.vimLayout === "translated" ? "translated" : "physical",
      defaultView: value.defaultView === "page" ? "page" : "mobile",
      vimEnabled: value.vimEnabled !== false,
      hideMenu: value.hideMenu === true,
      autocorrect: value.autocorrect !== false,
      autocomplete: value.autocomplete !== false,
      grammarUrl: typeof value.grammarUrl === "string" ? value.grammarUrl.trim() : "",
    };
  } catch { return defaultSettings; }
}

export function writeSettings(settings: Settings): Promise<void> {
  const raw = JSON.stringify(settings);
  pendingWrite = pendingWrite.catch(() => {}).then(async () => {
    if (isTauri()) {
      await mkdir("library", { baseDir: BaseDirectory.AppData, recursive: true });
      await writeTextFile(`${file}.tmp`, raw, { baseDir: BaseDirectory.AppData });
      await rename(`${file}.tmp`, file, { oldPathBaseDir: BaseDirectory.AppData, newPathBaseDir: BaseDirectory.AppData });
    } else localStorage.setItem(storageKey, raw);
  });
  return pendingWrite;
}
