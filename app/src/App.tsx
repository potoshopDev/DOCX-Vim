import { useEffect, useRef, useState } from "react";
import { EditorContent, useEditor, useEditorState } from "@tiptap/react";
import { isTauri } from "@tauri-apps/api/core";
import { basename } from "@tauri-apps/api/path";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { open, save } from "@tauri-apps/plugin-dialog";
import { readFile } from "@tauri-apps/plugin-fs";
import StarterKit from "@tiptap/starter-kit";
import TextAlign from "@tiptap/extension-text-align";
import { exportDocx, importDocx } from "./docx";
import { BackgroundSaveFailure, BackgroundSaver } from "./background-save";
import { listDrafts, type Draft } from "./draft";
import "./App.css";

type ViewMode = "page" | "mobile";

function App() {
  const [viewMode, setViewMode] = useState<ViewMode>("mobile");
  const [documentId, setDocumentId] = useState<string>(() => crypto.randomUUID());
  const [documentName, setDocumentName] = useState("Новый документ");
  const [recovery, setRecovery] = useState<Draft[]>([]);
  const [ready, setReady] = useState(false);
  const [revision, setRevision] = useState(0);
  const [saveStatus, setSaveStatus] = useState("Подготовка…");
  const [busy, setBusy] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const suppressUpdate = useRef(false);
  const currentId = useRef(documentId);
  const currentName = useRef(documentName);
  const currentSourcePath = useRef<string | undefined>(undefined);
  const currentSavePath = useRef<string | undefined>(undefined);
  const currentReady = useRef(false);
  const saver = useRef<BackgroundSaver | null>(null);
  if (!saver.current) saver.current = new BackgroundSaver((id, path) => {
    if (id === currentId.current) setSaveStatus(`${path ? "DOCX и черновик сохранены" : "Черновик сохранён"} ${new Date().toLocaleTimeString("ru-RU")}`);
  });
  const currentDraft = (): Draft => ({
    version: 1, id: currentId.current, name: currentName.current,
    document: editor!.getJSON(), updatedAt: new Date().toISOString(),
    sourcePath: currentSourcePath.current, savePath: currentSavePath.current,
  });
  const editor = useEditor({
    extensions: [
      StarterKit.configure({ link: { openOnClick: false } }),
      TextAlign.configure({ types: ["heading", "paragraph"] }),
    ],
    content:
      "<h1>Новый документ</h1><p>Выделите текст, чтобы изменить его оформление. Проверка кириллицы: Ёж, ёлка, «кавычки», — тире и emoji 👋.</p><p>Переключите вид документа: текст и выделение останутся в одном редакторе.</p>",
    editorProps: {
      attributes: {
        lang: "ru",
        "aria-label": "Текст документа",
        spellcheck: "true",
      },
    },
    onUpdate: () => {
      if (!suppressUpdate.current) setRevision((value) => value + 1);
    },
  });
  useEffect(() => {
    listDrafts().then((drafts) => {
      setRecovery(drafts);
      setReady(drafts.length === 0);
      currentReady.current = drafts.length === 0;
      setSaveStatus(drafts.length ? "Есть сохранённые черновики" : "Готово");
    }).catch(() => { setReady(true); currentReady.current = true; setSaveStatus(""); });
  }, []);
  useEffect(() => { editor?.setEditable(ready); }, [editor, ready]);
  useEffect(() => {
    if (!editor || !ready) return;
    setSaveStatus("Сохранение в фоне…");
    saver.current!.schedule(currentDraft());
  }, [editor, ready, revision, documentId, documentName]);

  const persistCurrent = async (draftIsEnough = false) => {
    if (!editor || !currentReady.current) return;
    try { await saver.current!.flush(currentDraft()); }
    catch (error) {
      if (!draftIsEnough || !(error instanceof BackgroundSaveFailure && error.draftSaved)) throw error;
    }
  };
  useEffect(() => {
    if (!editor) return;
    const saveWhenHidden = () => {
      if (document.visibilityState === "hidden" && currentReady.current) {
        void saver.current!.flush(currentDraft()).catch(() => {});
      }
    };
    document.addEventListener("visibilitychange", saveWhenHidden);
    if (!isTauri() || /Android/i.test(navigator.userAgent)) {
      return () => document.removeEventListener("visibilitychange", saveWhenHidden);
    }
    const windowHandle = getCurrentWindow();
    let closing = false;
    const listener = windowHandle.onCloseRequested(async (event) => {
      if (!currentReady.current || closing) return;
      event.preventDefault();
      closing = true;
      try {
        await persistCurrent(true);
        await windowHandle.destroy();
      } catch {
        closing = false;
        setSaveStatus("");
      }
    }).catch(() => () => {});
    return () => {
      document.removeEventListener("visibilitychange", saveWhenHidden);
      void listener.then((unlisten) => unlisten()).catch(() => {});
    };
  }, [editor]);

  const showDocument = (id: string, name: string, content: string | Draft["document"], paths: Pick<Draft, "sourcePath" | "savePath"> = {}) => {
    if (!editor) return;
    suppressUpdate.current = true;
    editor.commands.setContent(content);
    suppressUpdate.current = false;
    currentId.current = id;
    currentName.current = name;
    currentSourcePath.current = paths.sourcePath;
    currentSavePath.current = paths.savePath;
    setDocumentId(id);
    setDocumentName(name);
    setReady(true);
    currentReady.current = true;
    setRevision((value) => value + 1);
  };

  const restoreDraft = (draft: Draft) => {
    showDocument(draft.id, draft.name, draft.document, draft);
    setRecovery([]);
  };

  const startNew = async () => {
    if (!editor) return;
    try { await persistCurrent(true); } catch { setSaveStatus(""); return; }
    showDocument(crypto.randomUUID(), "Новый документ", "<p></p>");
    setRecovery([]);
  };

  const loadBytes = async (bytes: Uint8Array, name: string, sourcePath?: string) => {
    if (!editor) return;
    setBusy(true);
    try {
      const imported = await importDocx(bytes);
      await persistCurrent(true);
      showDocument(crypto.randomUUID(), name, imported.html || "<p></p>", { sourcePath });
      setRecovery([]);
      setSaveStatus("Открыт документ; исходный файл не изменён");
    } catch { setSaveStatus(""); }
    finally { setBusy(false); }
  };

  const openDocument = async () => {
    if (!editor || busy) return;
    if (!isTauri()) { fileInput.current?.click(); return; }
    try {
      const selected = await open({ multiple: false, filters: [{ name: "Документ Word", extensions: ["docx"] }] });
      if (typeof selected !== "string") return;
      const bytes = await readFile(selected);
      const name = await basename(selected).catch(() => selected.split(/[\\/]/).pop() ?? "Документ.docx");
      await loadBytes(bytes, name, selected);
    } catch { setSaveStatus(""); }
  };

  const saveDocument = async () => {
    if (!editor || busy || !ready) return;
    setBusy(true);
    try {
      const baseName = documentName.replace(/\.docx$/i, "");
      const suggestedName = `${baseName}${currentSourcePath.current && !currentSavePath.current ? " (DOCX Vim)" : ""}.docx`;
      if (isTauri()) {
        if (!currentSavePath.current) {
          const selected = await save({ defaultPath: suggestedName, filters: [{ name: "Документ Word", extensions: ["docx"] }] });
          if (!selected) return;
          currentSavePath.current = selected;
        }
        await persistCurrent();
      } else {
        const blob = await exportDocx(editor.getJSON());
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = suggestedName;
        anchor.click();
        window.setTimeout(() => URL.revokeObjectURL(url), 1000);
        await persistCurrent();
      }
      if (!isTauri()) setSaveStatus(`DOCX сохранён ${new Date().toLocaleTimeString("ru-RU")}`);
    } catch { setSaveStatus(""); }
    finally { setBusy(false); }
  };
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) =>
      current && {
        bold: current.isActive("bold"),
        italic: current.isActive("italic"),
        underline: current.isActive("underline"),
        strike: current.isActive("strike"),
        bulletList: current.isActive("bulletList"),
        orderedList: current.isActive("orderedList"),
        link: current.isActive("link"),
        heading1: current.isActive("heading", { level: 1 }),
        heading2: current.isActive("heading", { level: 2 }),
        heading3: current.isActive("heading", { level: 3 }),
        alignLeft: current.isActive({ textAlign: "left" }),
        alignCenter: current.isActive({ textAlign: "center" }),
        alignRight: current.isActive({ textAlign: "right" }),
      },
  });

  const setLink = () => {
    if (!editor) return;
    const previous = editor.getAttributes("link").href ?? "";
    const href = window.prompt("Адрес ссылки (пустое поле удалит ссылку)", previous);
    if (href === null) return;
    if (href.trim()) {
      editor.chain().focus().extendMarkRange("link").setLink({ href: href.trim() }).run();
    } else {
      editor.chain().focus().unsetLink().run();
    }
  };

  const formatButton = (label: string, title: string, selected: boolean, action: () => void) => (
    <button
      type="button"
      title={title}
      aria-label={title}
      aria-pressed={selected}
      onClick={action}
      disabled={!editor}
    >
      {label}
    </button>
  );

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <h1>DOCX Vim</h1>
          <p>{documentName}</p>
        </div>
        <div className="file-actions" role="group" aria-label="Файл">
          <button type="button" onClick={startNew} disabled={!editor || busy}>Новый</button>
          <button type="button" onClick={openDocument} disabled={!editor || busy}>Открыть DOCX</button>
          <button type="button" onClick={saveDocument} disabled={!editor || busy || !ready}>Сохранить DOCX</button>
          <input ref={fileInput} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" hidden onChange={async (event) => {
            const file = event.target.files?.[0];
            if (file) await loadBytes(new Uint8Array(await file.arrayBuffer()), file.name);
            event.target.value = "";
          }} />
        </div>
        <div className="view-switch" role="group" aria-label="Вид документа">
          <button type="button" aria-pressed={viewMode === "page"} onClick={() => setViewMode("page")}>
            Страница
          </button>
          <button type="button" aria-pressed={viewMode === "mobile"} onClick={() => setViewMode("mobile")}>
            Мобильный вид
          </button>
        </div>
      </header>

      {recovery.length > 0 && (
        <aside className="recovery-banner" role="status">
          <strong>Найдены черновики:</strong>
          {recovery.map((draft) => <button key={draft.id} type="button" onClick={() => restoreDraft(draft)}>{draft.name}</button>)}
          <button type="button" onClick={() => { setRecovery([]); setReady(true); currentReady.current = true; }}>Начать новый</button>
        </aside>
      )}
      <div className="toolbar" role="toolbar" aria-label="Форматирование документа">
        <select
          aria-label="Стиль абзаца"
          value={active?.heading1 ? "h1" : active?.heading2 ? "h2" : active?.heading3 ? "h3" : "p"}
          disabled={!editor}
          onChange={(event) => {
            const value = event.target.value;
            if (value === "p") editor?.chain().focus().setParagraph().run();
            else editor?.chain().focus().setHeading({ level: Number(value.slice(1)) as 1 | 2 | 3 }).run();
          }}
        >
          <option value="p">Обычный текст</option>
          <option value="h1">Заголовок 1</option>
          <option value="h2">Заголовок 2</option>
          <option value="h3">Заголовок 3</option>
        </select>
        <span className="toolbar-divider" />
        {formatButton("Ж", "Полужирный", !!active?.bold, () => editor?.chain().focus().toggleBold().run())}
        {formatButton("К", "Курсив", !!active?.italic, () => editor?.chain().focus().toggleItalic().run())}
        {formatButton("Ч", "Подчёркнутый", !!active?.underline, () => editor?.chain().focus().toggleUnderline().run())}
        {formatButton("З", "Зачёркнутый", !!active?.strike, () => editor?.chain().focus().toggleStrike().run())}
        <span className="toolbar-divider" />
        {formatButton("• ≡", "Маркированный список", !!active?.bulletList, () => editor?.chain().focus().toggleBulletList().run())}
        {formatButton("1. ≡", "Нумерованный список", !!active?.orderedList, () => editor?.chain().focus().toggleOrderedList().run())}
        <span className="toolbar-divider" />
        {formatButton("Слева", "По левому краю", !!active?.alignLeft, () => editor?.chain().focus().setTextAlign("left").run())}
        {formatButton("Центр", "По центру", !!active?.alignCenter, () => editor?.chain().focus().setTextAlign("center").run())}
        {formatButton("Справа", "По правому краю", !!active?.alignRight, () => editor?.chain().focus().setTextAlign("right").run())}
        <span className="toolbar-divider" />
        {formatButton("Ссылка", "Добавить или изменить ссылку", !!active?.link, setLink)}
        {formatButton("↶", "Отменить", false, () => editor?.chain().focus().undo().run())}
        {formatButton("↷", "Повторить", false, () => editor?.chain().focus().redo().run())}
      </div>

      <section className={`editor-area editor-area--${viewMode}`} aria-label="Редактор документа">
        <EditorContent editor={editor} className="document" />
      </section>

      <footer className="status-bar">
        <span>Обычный ввод</span>
        <span>{saveStatus}</span>
      </footer>
    </main>
  );
}

export default App;
