import { useCallback, useEffect, useRef, useState } from "react";
import { generateJSON } from "@tiptap/core";
import { isTauri } from "@tauri-apps/api/core";
import { basename } from "@tauri-apps/api/path";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { open } from "@tauri-apps/plugin-dialog";
import { readFile } from "@tauri-apps/plugin-fs";
import { importDocx } from "./docx";
import { editorExtensions } from "./editor-extensions";
import { listDrafts, readDraft, writeDraft, type Draft } from "./draft";
import { listRecent, rememberDocument, type RecentDocument } from "./library";
import { defaultSettings, readSettings, writeSettings, type Settings } from "./settings";
import { openDocumentWindow, openWindowIds } from "./windows";

export default function StartScreen() {
  const [drafts, setDrafts] = useState<Draft[]>([]);
  const [recent, setRecent] = useState<RecentDocument[]>([]);
  const [openIds, setOpenIds] = useState<string[]>([]);
  const [settings, setSettings] = useState<Settings>(defaultSettings);
  const [showSettings, setShowSettings] = useState(false);
  const [busy, setBusy] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const refresh = useCallback(async () => {
    const [savedDrafts, savedRecent, windows] = await Promise.all([listDrafts(), listRecent(), openWindowIds()]);
    setDrafts(savedDrafts); setRecent(savedRecent); setOpenIds(windows);
  }, []);
  useEffect(() => {
    void refresh();
    void readSettings().then(setSettings);
    const onFocus = () => { void refresh(); };
    addEventListener("focus", onFocus);
    return () => removeEventListener("focus", onFocus);
  }, [refresh]);
  useEffect(() => {
    if (!isTauri() || /Android/i.test(navigator.userAgent)) return;
    const main = getCurrentWindow();
    const listener = main.onCloseRequested(async (event) => {
      if ((await openWindowIds()).length > 0) { event.preventDefault(); await main.hide(); }
    });
    return () => { void listener.then((unlisten) => unlisten()); };
  }, []);

  const openDraft = async (draft: Draft) => {
    await rememberDocument({ id: draft.id, name: draft.name, sourcePath: draft.sourcePath });
    await openDocumentWindow(draft.id, draft.name);
    await refresh();
  };

  const importBytes = async (bytes: Uint8Array, name: string, sourcePath?: string, id: string = crypto.randomUUID()) => {
    const imported = await importDocx(bytes);
    const draft: Draft = {
      version: 1, id, name, document: generateJSON(imported.html || "<p></p>", editorExtensions),
      updatedAt: new Date().toISOString(), sourcePath,
    };
    await writeDraft(draft);
    await openDraft(draft);
  };

  const openFile = async () => {
    if (busy) return;
    if (!isTauri()) { fileInput.current?.click(); return; }
    setBusy(true);
    try {
      const selected = await open({ multiple: false, filters: [{ name: "Документ Word", extensions: ["docx"] }] });
      if (typeof selected !== "string") return;
      const existingDraft = drafts.find((draft) => draft.sourcePath === selected);
      if (existingDraft) { await openDraft(existingDraft); return; }
      const existing = recent.find((item) => item.sourcePath === selected);
      if (existing) {
        const draft = await readDraft(existing.id);
        if (draft) { await openDraft(draft); return; }
      }
      const bytes = await readFile(selected);
      const name = await basename(selected).catch(() => selected.split(/[\\/]/).pop() || "Документ.docx");
      await importBytes(bytes, name, selected, existing?.id);
    } catch { /* The start screen keeps the user's documents available. */ }
    finally { setBusy(false); }
  };

  const newDocument = async () => {
    if (busy) return;
    setBusy(true);
    try {
      const draft: Draft = {
        version: 1, id: crypto.randomUUID(), name: "Новый документ",
        document: { type: "doc", content: [{ type: "paragraph" }] },
        updatedAt: new Date().toISOString(),
      };
      await writeDraft(draft);
      await openDraft(draft);
    } catch { /* The draft is left on disk for recovery. */ }
    finally { setBusy(false); }
  };

  const openRecent = async (item: RecentDocument) => {
    if (busy) return;
    setBusy(true);
    try {
      const draft = await readDraft(item.id);
      if (draft) await openDraft(draft);
      else if (item.sourcePath) await importBytes(await readFile(item.sourcePath), item.name, item.sourcePath, item.id);
    } catch { /* A missing source does not remove the recent entry. */ }
    finally { setBusy(false); }
  };

  const changeSetting = (patch: Partial<Settings>) => {
    const next = { ...settings, ...patch };
    setSettings(next);
    void writeSettings(next).catch(() => {});
  };

  return <main className="start-screen">
    <header className="start-header">
      <div><h1>DOCX Vim</h1><p>Документы и черновики</p></div>
      <button type="button" onClick={() => setShowSettings((value) => !value)} aria-expanded={showSettings}>Настройки</button>
    </header>
    <div className="start-content">
      <div className="start-actions">
        <button type="button" onClick={newDocument} disabled={busy}>Новый документ</button>
        <button type="button" onClick={openFile} disabled={busy}>Открыть DOCX</button>
        <input ref={fileInput} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" hidden
          onChange={async (event) => {
            const file = event.target.files?.[0];
            if (file) { setBusy(true); try { await importBytes(new Uint8Array(await file.arrayBuffer()), file.name); } finally { setBusy(false); } }
            event.target.value = "";
          }} />
      </div>
      {showSettings && <section className="start-panel settings-panel" aria-label="Настройки приложения">
        <h2>Настройки</h2>
        <label>Автосохранение
          <select value={settings.autosaveMs} onChange={(event) => changeSetting({ autosaveMs: Number(event.target.value) })}>
            <option value={800}>Через 0,8 секунды</option><option value={2000}>Через 2 секунды</option><option value={5000}>Через 5 секунд</option>
          </select>
        </label>
        <label>Язык проверки
          <select value={settings.language} onChange={(event) => changeSetting({ language: event.target.value as Settings["language"] })}>
            <option value="ru">Русский</option><option value="en">Английский</option>
          </select>
        </label>
        <label>Раскладка команд Vim
          <select value={settings.vimLayout} onChange={(event) => changeSetting({ vimLayout: event.target.value as Settings["vimLayout"] })}>
            <option value="physical">По физическим клавишам</option><option value="translated">Перевод русской раскладки</option>
          </select>
        </label>
        <label>Вид документа по умолчанию
          <select value={settings.defaultView} onChange={(event) => changeSetting({ defaultView: event.target.value as Settings["defaultView"] })}>
            <option value="mobile">Мобильный вид</option><option value="page">Страница</option>
          </select>
        </label>
      </section>}
      <section className="start-panel"><h2>Черновики</h2>
        {drafts.length ? <ul className="document-list">{drafts.map((draft) => <li key={draft.id}>
          <button type="button" onClick={() => void openDraft(draft).catch(() => {})}>
            <strong>{draft.name}</strong><span>{openIds.includes(draft.id) ? "Открыт · " : ""}{new Date(draft.updatedAt).toLocaleString("ru-RU")}</span>
          </button>
        </li>)}</ul> : <p className="empty-list">Пока нет черновиков</p>}
      </section>
      <section className="start-panel"><h2>Недавние документы</h2>
        {recent.length ? <ul className="document-list">{recent.map((item) => <li key={item.id}>
          <button type="button" onClick={() => void openRecent(item)} disabled={busy}>
            <strong>{item.name}</strong><span>{openIds.includes(item.id) ? "Открыт · " : ""}{new Date(item.openedAt).toLocaleString("ru-RU")}</span>
          </button>
        </li>)}</ul> : <p className="empty-list">Открытые файлы появятся здесь</p>}
      </section>
    </div>
  </main>;
}
