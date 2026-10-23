import { useEffect, useRef, useState } from "react";
import { EditorContent, useEditor, useEditorState } from "@tiptap/react";
import { isTauri } from "@tauri-apps/api/core";
import { onBackButtonPress } from "@tauri-apps/api/app";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { save } from "@tauri-apps/plugin-dialog";
import { exportDocx } from "./docx";
import { BackgroundSaveFailure, BackgroundSaver } from "./background-save";
import { readDraft, type Draft } from "./draft";
import { editorExtensions } from "./editor-extensions";
import { readSettings } from "./settings";
import { openWindowIds, registerDocumentTask, showLibrary } from "./windows";
import "./App.css";

type ViewMode = "page" | "mobile";

function DocumentEditor({ documentId }: { documentId: string }) {
  const [viewMode, setViewMode] = useState<ViewMode>("mobile");
  const [documentName, setDocumentName] = useState("Новый документ");
  const [ready, setReady] = useState(false);
  const [revision, setRevision] = useState(0);
  const [saveStatus, setSaveStatus] = useState("Подготовка…");
  const [busy, setBusy] = useState(false);
  const suppressUpdate = useRef(false);
  const editorArea = useRef<HTMLElement>(null);
  const currentId = useRef(documentId);
  const currentName = useRef(documentName);
  const currentSourcePath = useRef<string | undefined>(undefined);
  const currentSavePath = useRef<string | undefined>(undefined);
  const currentViewMode = useRef<ViewMode>("mobile");
  const currentReady = useRef(false);
  const saver = useRef<BackgroundSaver | null>(null);
  if (!saver.current) saver.current = new BackgroundSaver((_id, path) => {
    setSaveStatus(`${path ? "DOCX и черновик сохранены" : "Черновик сохранён"} ${new Date().toLocaleTimeString("ru-RU")}`);
  });
  const currentDraft = (): Draft => ({
    version: 1, id: currentId.current, name: currentName.current,
    document: editor!.getJSON(), updatedAt: new Date().toISOString(),
    sourcePath: currentSourcePath.current, savePath: currentSavePath.current,
    viewMode: currentViewMode.current,
    selection: editor ? { anchor: editor.state.selection.anchor, head: editor.state.selection.head } : undefined,
    scrollTop: editorArea.current?.scrollTop,
  });
  const editor = useEditor({
    extensions: editorExtensions,
    content: "<p></p>",
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
    if (!editor) return;
    let cancelled = false;
    Promise.all([readDraft(documentId), readSettings()]).then(([draft, settings]) => {
      if (cancelled || !draft) return;
      saver.current = new BackgroundSaver((_id, path) => {
        setSaveStatus(`${path ? "DOCX и черновик сохранены" : "Черновик сохранён"} ${new Date().toLocaleTimeString("ru-RU")}`);
      }, settings.autosaveMs);
      suppressUpdate.current = true;
      editor.commands.setContent(draft.document);
      suppressUpdate.current = false;
      currentName.current = draft.name;
      currentSourcePath.current = draft.sourcePath;
      currentSavePath.current = draft.savePath;
      setDocumentName(draft.name);
      const mode = draft.viewMode || settings.defaultView;
      currentViewMode.current = mode;
      setViewMode(mode);
      editor.view.dom.setAttribute("lang", settings.language);
      if (draft.selection) {
        const max = editor.state.doc.content.size;
        editor.commands.setTextSelection({ from: Math.min(draft.selection.anchor, max), to: Math.min(draft.selection.head, max) });
      }
      currentReady.current = true;
      setReady(true);
      setSaveStatus("Готово");
      void registerDocumentTask(documentId, draft.name);
      if (draft.scrollTop != null) requestAnimationFrame(() => {
        if (editorArea.current) editorArea.current.scrollTop = draft.scrollTop!;
      });
      if (isTauri()) void getCurrentWindow().setTitle(draft.name);
    }).catch(() => setSaveStatus(""));
    return () => { cancelled = true; };
  }, [editor, documentId]);
  useEffect(() => { editor?.setEditable(ready); }, [editor, ready]);
  useEffect(() => {
    if (!editor || !ready) return;
    setSaveStatus("Сохранение в фоне…");
    saver.current!.schedule(currentDraft());
  }, [editor, ready, revision, documentId, documentName, viewMode]);

  const persistCurrent = async (draftIsEnough = false) => {
    if (!editor || !currentReady.current) return;
    try { await saver.current!.flush(currentDraft()); }
    catch (error) {
      if (!draftIsEnough || !(error instanceof BackgroundSaveFailure && error.draftSaved)) throw error;
    }
  };
  useEffect(() => {
    if (!isTauri() || !/Android/i.test(navigator.userAgent)) return;
    const listener = onBackButtonPress(() => {
      void (async () => {
        try { await persistCurrent(true); } catch { return; }
        await showLibrary();
        if (getCurrentWindow().label !== "main") await getCurrentWindow().destroy();
      })();
    });
    return () => { void listener.then((handle) => handle.unregister()); };
  }, [editor]);
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
        if (windowHandle.label !== "main" && (await openWindowIds()).length <= 1) await showLibrary();
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
          <h1>{documentName}</h1>
        </div>
        <div className="file-actions" role="group" aria-label="Файл">
          <button type="button" onClick={() => void showLibrary()}>Документы</button>
          <button type="button" onClick={saveDocument} disabled={!editor || busy || !ready}>Сохранить DOCX</button>
        </div>
        <div className="view-switch" role="group" aria-label="Вид документа">
          <button type="button" aria-pressed={viewMode === "page"} onClick={() => { currentViewMode.current = "page"; setViewMode("page"); }}>
            Страница
          </button>
          <button type="button" aria-pressed={viewMode === "mobile"} onClick={() => { currentViewMode.current = "mobile"; setViewMode("mobile"); }}>
            Мобильный вид
          </button>
        </div>
      </header>
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

      <section ref={editorArea} className={`editor-area editor-area--${viewMode}`} aria-label="Редактор документа">
        <EditorContent editor={editor} className="document" />
      </section>

      <footer className="status-bar">
        <span>Обычный ввод</span>
        <span>{saveStatus}</span>
      </footer>
    </main>
  );
}

export default DocumentEditor;
