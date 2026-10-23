import { useEffect, useRef, useState } from "react";
import { EditorContent, useEditor, useEditorState } from "@tiptap/react";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { onBackButtonPress } from "@tauri-apps/api/app";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import { emitTo, listen } from "@tauri-apps/api/event";
import { save } from "@tauri-apps/plugin-dialog";
import { exportDocx } from "./docx";
import { BackgroundSaveFailure, BackgroundSaver } from "./background-save";
import { readDraft, type Draft } from "./draft";
import { editorExtensions } from "./editor-extensions";
import { readSettings } from "./settings";
import { VimMode as VimExtension, getVimMode } from "vim-prosemirror/tiptap";
import "vim-prosemirror/style.css";
import { VimAdapter, type VimMode } from "./vim-adapter";
import { focusAdjacentWindow, isDocumentSaveOwner, openDocumentWindow, openWindowIds, registerDocumentTask, showLibrary } from "./windows";
import "./App.css";

type ViewMode = "page" | "mobile";

function DocumentEditor({ documentId }: { documentId: string }) {
  const [viewMode, setViewMode] = useState<ViewMode>("mobile");
  const [documentName, setDocumentName] = useState("Новый документ");
  const [ready, setReady] = useState(false);
  const [revision, setRevision] = useState(0);
  const [saveStatus, setSaveStatus] = useState("Подготовка…");
  const [busy, setBusy] = useState(false);
  const [vimEnabled, setVimEnabled] = useState(true);
  const [hideMenu, setHideMenu] = useState(false);
  const [vimMode, setVimMode] = useState<VimMode>("normal");
  const [showHeadings, setShowHeadings] = useState(false);
  const [commandLine, setCommandLine] = useState<string | null>(null);
  const [searchLine, setSearchLine] = useState<{ backward: boolean; query: string } | null>(null);
  const commandInput = useRef<HTMLInputElement>(null);
  const searchInput = useRef<HTMLInputElement>(null);
  useEffect(() => { if (commandLine !== null) commandInput.current?.focus(); }, [commandLine]);
  useEffect(() => { if (searchLine !== null) searchInput.current?.focus(); }, [searchLine]);
  const vim = useRef<VimAdapter | null>(null);
  const documentChannel = useRef<BroadcastChannel | null>(null);
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
    extensions: [...editorExtensions, VimExtension],
    content: "<p></p>",
    editorProps: {
      handleKeyDown: (_view, event) => vim.current?.handle(event) ?? false,
      attributes: {
        lang: "ru",
        "aria-label": "Текст документа",
        spellcheck: "true",
      },
    },
    onUpdate: () => {
      if (!suppressUpdate.current) {
        setRevision((value) => value + 1);
        documentChannel.current?.postMessage({ type: "document", document: editor?.getJSON() });
      }
    },
    onTransaction: ({ editor: current }) => setVimMode(getVimMode(current)),
  });
  useEffect(() => {
    if (!editor) return;
    let cancelled = false;
    Promise.all([readDraft(documentId), readSettings()]).then(([draft, settings]) => {
      if (cancelled || !draft) return;
      setVimEnabled(settings.vimEnabled);
      setHideMenu(settings.hideMenu);
      saver.current = new BackgroundSaver((_id, path) => {
        setSaveStatus(`${path ? "DOCX и черновик сохранены" : "Черновик сохранён"} ${new Date().toLocaleTimeString("ru-RU")}`);
      }, settings.autosaveMs);
      suppressUpdate.current = true;
      editor.chain().setMeta("addToHistory", false).setContent(draft.document, { emitUpdate: false }).run();
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
      requestAnimationFrame(() => editor.commands.focus());
      setSaveStatus("Готово");
      void registerDocumentTask(isTauri() ? getCurrentWindow().label.slice(4) : documentId, draft.name);
      if (draft.scrollTop != null) requestAnimationFrame(() => {
        if (editorArea.current) editorArea.current.scrollTop = draft.scrollTop!;
      });
      if (isTauri()) void getCurrentWindow().setTitle(draft.name);
    }).catch(() => setSaveStatus(""));
    return () => { cancelled = true; };
  }, [editor, documentId]);
  useEffect(() => {
    if (!editor || !ready) return;
    const channel = new BroadcastChannel(`docx-vim-${documentId}`);
    documentChannel.current = channel;
    channel.onmessage = (event: MessageEvent<{ type: string; document?: Draft["document"]; savePath?: string }>) => {
      if (event.data.type === "active") { editor.setEditable(false); return; }
      if (event.data.type === "save-request") {
        void isDocumentSaveOwner(documentId).then((owner) => {
          if (!owner) return;
          if (event.data.savePath) currentSavePath.current = event.data.savePath;
          void saver.current?.flush(currentDraft()).catch(() => {});
        });
        return;
      }
      if (event.data.type !== "document" || !event.data.document) return;
      const anchor = editor.state.selection.anchor;
      const head = editor.state.selection.head;
      suppressUpdate.current = true;
      editor.chain().setMeta("addToHistory", false).setContent(event.data.document, { emitUpdate: false }).run();
      const max = editor.state.doc.content.size;
      editor.commands.setTextSelection({ from: Math.min(anchor, max), to: Math.min(head, max) });
      suppressUpdate.current = false;
      setRevision((value) => value + 1);
    };
    const onFocus = () => {
      editor.setEditable(true);
      channel.postMessage({ type: "active" });
      setRevision((value) => value + 1);
    };
    addEventListener("focus", onFocus);
    if (document.hasFocus()) onFocus();
    return () => { removeEventListener("focus", onFocus); channel.close(); documentChannel.current = null; };
  }, [editor, ready, documentId]);
  useEffect(() => {
    if (!editor) return;
    vim.current = new VimAdapter(editor, (key) => {
      if (["s", "v"].includes(key)) void persistCurrent(true).then(() => openDocumentWindow(documentId, documentName, true)).catch(() => {});
      else if (["w", "W", "p", "h", "j", "k", "l"].includes(key)) void focusAdjacentWindow(key === "W" || key === "p" || key === "h" || key === "k" ? -1 : 1);
    }, () => { setCommandLine(""); requestAnimationFrame(() => commandInput.current?.focus()); },
    (backward) => { setSearchLine({ backward, query: "" }); requestAnimationFrame(() => searchInput.current?.focus()); });
    vim.current.setEnabled(vimEnabled);
    setVimMode(vim.current.mode());
    return () => { vim.current = null; };
  }, [editor, vimEnabled, documentId, documentName]);
  useEffect(() => { editor?.setEditable(ready); }, [editor, ready]);
  useEffect(() => {
    const refreshSettings = () => { void readSettings().then((settings) => {
      setVimEnabled(settings.vimEnabled); setHideMenu(settings.hideMenu);
    }); if (currentReady.current) requestAnimationFrame(() => editor?.commands.focus()); };
    addEventListener("focus", refreshSettings);
    return () => removeEventListener("focus", refreshSettings);
  }, [editor]);
  useEffect(() => {
    if (!isTauri() || !editor) return;
    const listener = getCurrentWindow().listen("focus-editor", () => {
      if (currentReady.current) requestAnimationFrame(() => editor.commands.focus());
    });
    return () => { void listener.then((unlisten) => unlisten()); };
  }, [editor]);
  useEffect(() => {
    if (!editor || !ready) return;
    setSaveStatus("Сохранение в фоне…");
    void isDocumentSaveOwner(documentId).then((owner) => {
      if (owner) saver.current!.schedule(currentDraft());
    });
  }, [editor, ready, revision, documentId, documentName, viewMode]);

  const persistCurrent = async (draftIsEnough = false) => {
    if (!editor || !currentReady.current) return;
    if (!(await isDocumentSaveOwner(documentId)) && documentChannel.current) {
      documentChannel.current.postMessage({ type: "save-request", savePath: currentSavePath.current });
      return;
    }
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
        void isDocumentSaveOwner(documentId).then((owner) => {
          if (owner) void saver.current!.flush(currentDraft()).catch(() => {});
        });
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

  const saveDocument = async (): Promise<boolean> => {
    if (!editor || busy || !ready) return false;
    setBusy(true);
    try {
      const baseName = documentName.replace(/\.docx$/i, "");
      const suggestedName = `${baseName}${currentSourcePath.current && !currentSavePath.current ? " (DOCX Vim)" : ""}.docx`;
      if (isTauri()) {
        if (!currentSavePath.current) {
          const selected = await save({ defaultPath: suggestedName, filters: [{ name: "Документ Word", extensions: ["docx"] }] });
          if (!selected) return false;
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
      return true;
    } catch { setSaveStatus(""); return false; }
    finally { setBusy(false); }
  };

  const quitDocument = async (): Promise<void> => {
    try { await persistCurrent(true); } catch { return; }
    await showLibrary();
    if (isTauri() && getCurrentWindow().label !== "main") await getCurrentWindow().destroy();
  };

  const writeQuitAll = async (): Promise<void> => {
    if (!isTauri()) { if (await saveDocument()) await quitDocument(); return; }
    const windows = (await WebviewWindow.getAll()).filter((item) => item.label.startsWith("doc-"));
    if (!windows.length) { if (await saveDocument()) await quitDocument(); return; }
    const owners = [...new Map(windows.slice().sort((a, b) => a.label.length - b.label.length || a.label.localeCompare(b.label))
      .map((item) => [item.label.slice(4).split("--")[0], item])).values()];
    for (const owner of owners) {
      if (owner.label === getCurrentWindow().label) {
        if (!(await saveDocument())) return;
        continue;
      }
      if (/Android/i.test(navigator.userAgent)) await invoke("focus_task", { id: owner.label.slice(4) });
      else await owner.setFocus();
      const token = crypto.randomUUID();
      const result = await new Promise<boolean>((resolve) => {
        let done = false;
        let unlisten: (() => void) | null = null;
        const timer = window.setTimeout(() => { if (!done) { done = true; unlisten?.(); resolve(false); } }, 120_000);
        void listen<{ token: string; ok: boolean }>("vim-write-result", (event) => {
          if (event.payload.token !== token || done) return;
          done = true; window.clearTimeout(timer); unlisten?.(); resolve(event.payload.ok);
        }).then((stop) => {
          unlisten = stop;
          void emitTo(owner.label, "vim-write-request", { token, replyTo: getCurrentWindow().label }).catch(() => {
            if (!done) { done = true; window.clearTimeout(timer); stop(); resolve(false); }
          });
        }).catch(() => { if (!done) { done = true; window.clearTimeout(timer); resolve(false); } });
      });
      if (!result) return;
    }
    for (const other of windows) if (other.label !== getCurrentWindow().label) await emitTo(other.label, "vim-quit-request");
    await quitDocument();
  };

  const saveHandler = useRef(saveDocument);
  const quitHandler = useRef(quitDocument);
  saveHandler.current = saveDocument;
  quitHandler.current = quitDocument;
  useEffect(() => {
    if (!isTauri()) return;
    const writeListener = getCurrentWindow().listen<{ token: string; replyTo: string }>("vim-write-request", async (event) => {
      const ok = await saveHandler.current();
      await emitTo(event.payload.replyTo, "vim-write-result", { token: event.payload.token, ok });
    });
    const quitListener = getCurrentWindow().listen("vim-quit-request", () => { void quitHandler.current(); });
    return () => { void writeListener.then((stop) => stop()); void quitListener.then((stop) => stop()); };
  }, []);
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
  const headings = useEditorState({
    editor,
    selector: ({ editor: current }) => {
      const items: { pos: number; level: number; text: string }[] = [];
      current?.state.doc.descendants((node, pos) => {
        if (node.type.name === "heading") items.push({ pos, level: node.attrs.level as number, text: node.textContent || "Без названия" });
      });
      return items;
    },
  }) ?? [];

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
      {!hideMenu && <header className="app-header">
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
      </header>}
      {!hideMenu && <div className="toolbar" role="toolbar" aria-label="Форматирование документа">
        <button type="button" onClick={() => setShowHeadings((value) => !value)} aria-expanded={showHeadings}>Заголовки</button>
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
      </div>}

      {showHeadings && <nav className="headings-panel" aria-label="Навигация по заголовкам">
        {headings.length ? headings.map((heading) => <button key={heading.pos} type="button" style={{ paddingLeft: `${heading.level}rem` }}
          onClick={() => { editor?.chain().focus().setTextSelection(heading.pos + 1).scrollIntoView().run(); setShowHeadings(false); }}>
          {heading.text}
        </button>) : <span>Заголовков пока нет</span>}
      </nav>}

      <section ref={editorArea} className={`editor-area editor-area--${viewMode}`} aria-label="Редактор документа">
        <EditorContent editor={editor} className="document" />
      </section>

      <footer className="status-bar">
        <span>{vimEnabled ? `Vim ${vimMode}` : "Обычный ввод"}</span>
        {commandLine !== null && <label className="vim-command">:
          <input ref={commandInput} aria-label="Команда Vim" value={commandLine}
            onChange={(event) => setCommandLine(event.target.value)}
            onKeyDown={(event) => {
              event.stopPropagation();
              if (event.key === "Escape") { setCommandLine(null); editor?.commands.focus(); }
              if (event.key === "Enter") {
                const command = commandLine.trim().replace(/^:/, "");
                setCommandLine(null);
                if (/^(?:sp(?:lit)?|vs(?:plit)?)$/.test(command)) void persistCurrent(true).then(() => openDocumentWindow(documentId, documentName, true)).then(() => {
                  if (!isTauri()) editor?.commands.focus();
                }).catch(() => {});
                else if (/^(?:w|write)$/.test(command)) void saveDocument().then(() => editor?.commands.focus());
                else if (/^(?:q|quit)$/.test(command)) void quitDocument();
                else if (/^(?:wq|x)$/.test(command)) void saveDocument().then((saved) => { if (saved) void quitDocument(); else editor?.commands.focus(); });
                else if (/^(?:wqa|wqall|xa|xall)$/.test(command)) void writeQuitAll();
                else editor?.commands.focus();
              }
            }} />
        </label>}
        {searchLine !== null && <label className="vim-command">{searchLine.backward ? "?" : "/"}
          <input ref={searchInput} aria-label="Поиск Vim" value={searchLine.query}
            onChange={(event) => setSearchLine({ ...searchLine, query: event.target.value })}
            onKeyDown={(event) => {
              event.stopPropagation();
              if (event.key === "Escape") { setSearchLine(null); editor?.commands.focus(); }
              if (event.key === "Enter") {
                vim.current?.submitSearch(searchLine.query, searchLine.backward);
                setSearchLine(null);
                editor?.commands.focus();
              }
            }} />
        </label>}
        {hideMenu && <button type="button" onClick={() => setShowHeadings((value) => !value)}>Заголовки</button>}
        <span>{saveStatus}</span>
      </footer>
    </main>
  );
}

export default DocumentEditor;
