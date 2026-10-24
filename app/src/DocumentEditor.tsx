import { useEffect, useMemo, useRef, useState } from "react";
import { flushSync } from "react-dom";
import { EditorContent, useEditor, useEditorState } from "@tiptap/react";
import type { EditorState } from "@tiptap/pm/state";
import type { EditorProps } from "@tiptap/pm/view";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { onBackButtonPress } from "@tauri-apps/api/app";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import { emitTo, listen } from "@tauri-apps/api/event";
import { save } from "@tauri-apps/plugin-dialog";
import { exportDocx } from "./docx";
import { BackgroundSaveFailure, BackgroundSaver } from "./background-save";
import { readDraft, writeDraft, type Draft } from "./draft";
import { editorExtensions } from "./editor-extensions";
import { readSettings } from "./settings";
import { defaultSettings, type Settings } from "./settings";
import { readPersonalDictionary, writePersonalDictionary, type PersonalDictionary, type Language } from "./personal-dictionary";
import { TextCheck, textCheckKey, type TextIssue } from "./text-check-extension";
import { checkGrammar } from "./grammar";
import { VimMode as VimExtension, getVimMode } from "vim-prosemirror/tiptap";
import "vim-prosemirror/style.css";
import { VimAdapter, type VimMode } from "./vim-adapter";
import { focusAdjacentWindow, isDocumentSaveOwner, openDocumentWindow, openWindowIds, registerDocumentTask, showLibrary } from "./windows";
import "./App.css";

type ViewMode = "page" | "mobile";

function wordAtCaret(state: EditorState): { from: number; to: number; word: string } | null {
  if (!state.selection.empty) return null;
  const $from = state.selection.$from;
  if (!$from.parent.isTextblock) return null;
  const before = $from.nodeBefore?.text ?? "";
  const after = $from.nodeAfter?.text ?? "";
  const text = before + after;
  for (const match of text.matchAll(/[A-Za-zА-Яа-яЁё]+(?:[’'-][A-Za-zА-Яа-яЁё]+)*/gu)) {
    const start = match.index, end = start + match[0].length;
    if (start <= before.length && before.length <= end) {
      const from = $from.pos - (before.length - start);
      return { from, to: from + match[0].length, word: match[0] };
    }
  }
  return null;
}

function DocumentEditor({ documentId }: { documentId: string }) {
  const [viewMode, setViewMode] = useState<ViewMode>("mobile");
  const [documentName, setDocumentName] = useState("Новый документ");
  const [ready, setReady] = useState(false);
  const [revision, setRevision] = useState(0);
  const [saveStatus, setSaveStatus] = useState("Подготовка…");
  const [busy, setBusy] = useState(false);
  const [vimEnabled, setVimEnabled] = useState(true);
  const [hideMenu, setHideMenu] = useState(false);
  const [assistSettings, setAssistSettings] = useState<Settings>(defaultSettings);
  const [personal, setPersonal] = useState<PersonalDictionary>({ ru: [], en: [] });
  const [selectedIssue, setSelectedIssue] = useState<TextIssue | null>(null);
  const [correctionPending, setCorrectionPending] = useState(false);
  const [correctionIndex, setCorrectionIndex] = useState(0);
  const [completion, setCompletion] = useState<{ from: number; to: number; suggestion: string } | null>(null);
  const [selectionRevision, setSelectionRevision] = useState(0);
  const issuesRef = useRef<TextIssue[]>([]);
  const selectedIssueRef = useRef<TextIssue | null>(null);
  const correctionIndexRef = useRef(0);
  const correctionRequest = useRef(0);
  const correctionPendingRef = useRef(false);
  const pendingCorrectionOffset = useRef(0);
  const ignoredIssues = useRef(new Set<string>());
  const completionRef = useRef<typeof completion>(null);
  const correctionPanel = useRef<HTMLDivElement>(null);
  const assistSettingsRef = useRef(assistSettings);
  const personalRef = useRef(personal);
  const checkWorker = useRef<Worker | null>(null);
  const nextWorkerId = useRef(0);
  const workerRequests = useRef(new Map<number, (response: any) => void>());
  assistSettingsRef.current = assistSettings;
  personalRef.current = personal;
  completionRef.current = completion;
  selectedIssueRef.current = selectedIssue;
  correctionIndexRef.current = correctionIndex;
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
  const correctionChoiceCount = (issue: TextIssue) => issue.suggestions.length + (issue.kind === "spelling" ? 3 : 2);
  const applyCorrectionChoice = (index: number): boolean => {
    const issue = selectedIssueRef.current;
    if (!issue || !editor) return false;
    const suggestionCount = issue.suggestions.length;
    if (index < 0 || index >= correctionChoiceCount(issue)) return false;
    if (index < suggestionCount) {
      if (editor.state.doc.textBetween(issue.from, issue.to) !== issue.word) return false;
      editor.chain().focus().insertContentAt({ from: issue.from, to: issue.to }, issue.suggestions[index]).run();
    } else if (issue.kind === "spelling" && index === suggestionCount) {
      const language: Language = /[А-Яа-яЁё]/.test(issue.word) ? "ru" : "en";
      const word = issue.word.toLocaleLowerCase(language);
      const next = { ...personalRef.current, [language]: [...new Set([...personalRef.current[language], word])].sort() };
      personalRef.current = next;
      setPersonal(next);
      void writePersonalDictionary(next).catch(() => {});
    } else if (index === correctionChoiceCount(issue) - 2) {
      ignoredIssues.current.add(`${issue.kind}:${issue.word}`);
      const next = issuesRef.current.filter((item) => !ignoredIssues.current.has(`${item.kind}:${item.word}`));
      issuesRef.current = next;
      editor.view.dispatch(editor.state.tr.setMeta(textCheckKey, next).setMeta("addToHistory", false));
    }
    correctionRequest.current++;
    correctionPendingRef.current = false;
    selectedIssueRef.current = null;
    setSelectedIssue(null);
    setCorrectionPending(false);
    editor.commands.focus();
    return true;
  };
  const cycleCorrection = (direction: 1 | -1): boolean => {
    const issue = selectedIssueRef.current;
    if (issue && correctionPendingRef.current) {
      pendingCorrectionOffset.current += direction;
      return true;
    }
    if (!issue) return false;
    const count = correctionChoiceCount(issue);
    const next = (correctionIndexRef.current + direction + count) % count;
    correctionIndexRef.current = next;
    flushSync(() => setCorrectionIndex(next));
    return true;
  };
  const acceptCorrection = (): boolean => {
    const issue = selectedIssueRef.current;
    return issue ? applyCorrectionChoice(correctionIndexRef.current) : false;
  };
  const showCorrectionsAtCaret = (direction: 1 | -1) => {
    if (!editor) return;
    const target = wordAtCaret(editor.state);
    if (!target) { selectedIssueRef.current = null; correctionPendingRef.current = false; setSelectedIssue(null); setCorrectionPending(false); return; }
    const issue = issuesRef.current.find((item) => item.from === target.from && item.to === target.to);
    if (issue?.kind === "grammar") {
      selectedIssueRef.current = issue;
      setSelectedIssue(issue);
      correctionPendingRef.current = false;
      setCorrectionPending(false);
      const initialIndex = direction === -1 ? Math.max(0, issue.suggestions.length - 1) : 0;
      correctionIndexRef.current = initialIndex;
      setCorrectionIndex(initialIndex);
      return;
    }
    const request = ++correctionRequest.current;
    const caret = editor.state.selection.from;
    const language: Language = /[А-Яа-яЁё]/.test(target.word) ? "ru" : "en";
    const pending: TextIssue = issue ?? { ...target, kind: "spelling", suggestions: [] };
    selectedIssueRef.current = pending;
    correctionPendingRef.current = true;
    pendingCorrectionOffset.current = direction === -1 ? -1 : 0;
    flushSync(() => { setSelectedIssue(pending); setCorrectionPending(true); });
    void askWorker({ type: "suggest", word: target.word, language, personal: personalRef.current })
      .then((result: { suggestions?: string[]; correct?: boolean }) => {
        if (request !== correctionRequest.current || !editor || editor.state.selection.from !== caret ||
          editor.state.doc.textBetween(target.from, target.to) !== target.word) return;
        if (result.correct) {
          selectedIssueRef.current = null;
          correctionPendingRef.current = false;
          flushSync(() => { setSelectedIssue(null); setCorrectionPending(false); });
          return;
        }
        const next = { ...pending, suggestions: result.suggestions ?? [] };
        selectedIssueRef.current = next;
        correctionPendingRef.current = false;
        const index = next.suggestions.length
          ? (pendingCorrectionOffset.current % next.suggestions.length + next.suggestions.length) % next.suggestions.length : 0;
        correctionIndexRef.current = index;
        flushSync(() => { setSelectedIssue(next); setCorrectionIndex(index); setCorrectionPending(false); });
      });
  };
  const stableExtensions = useMemo(() => [...editorExtensions, VimExtension, TextCheck], []);
  const editorHandlers = useRef<EditorProps>({});
  editorHandlers.current = {
      handleKeyDown: (_view, event) => {
        const editable = !vim.current?.enabled || vim.current?.mode() === "insert";
        const correctionKey = event.ctrlKey && !event.altKey && !event.metaKey && !event.shiftKey
          ? event.code === "KeyN" || ["n", "N", "т", "Т"].includes(event.key) ? 1
            : event.code === "KeyP" || ["p", "P", "з", "З"].includes(event.key) ? -1 : 0
          : 0;
        if (correctionKey) {
          event.preventDefault();
          event.stopPropagation();
          const direction = correctionKey as 1 | -1;
          if (!cycleCorrection(direction)) showCorrectionsAtCaret(direction);
          return true;
        }
        const correction = selectedIssueRef.current;
        if (event.key === "Tab" && correction) {
          event.preventDefault();
          event.stopPropagation();
          if (!correctionPendingRef.current) acceptCorrection();
          return true;
        }
        if (event.key === "Tab" && (!editable || !completionRef.current)) {
          event.preventDefault();
          event.stopPropagation();
          showCorrectionsAtCaret(1);
          if (assistSettingsRef.current.autocomplete) setSelectionRevision((value) => value + 1);
          return true;
        }
        if (event.key === "Escape" && correction) {
          event.preventDefault();
          correctionRequest.current++;
          selectedIssueRef.current = null;
          correctionPendingRef.current = false;
          setSelectedIssue(null);
          setCorrectionPending(false);
          return true;
        }
        if (editable && completionRef.current && (event.key === "Tab" || (event.ctrlKey && event.code === "Space"))) {
          event.preventDefault();
          const item = completionRef.current;
          editor?.chain().focus().insertContentAt({ from: item.from, to: item.to }, item.suggestion).run();
          setCompletion(null);
          return true;
        }
        if (editable && event.key === "Escape" && completionRef.current) { setCompletion(null); return true; }
        if (editable && event.ctrlKey && event.code === "Space") { setSelectionRevision((value) => value + 1); event.preventDefault(); return true; }
        return vim.current?.handle(event) ?? false;
      },
      handleTextInput: (view, from, to, text) => {
        if (!assistSettingsRef.current.autocorrect || (vim.current?.enabled && vim.current.mode() !== "insert")) return false;
        const $from = view.state.doc.resolve(from);
        const before = $from.parent.textBetween(Math.max(0, $from.parentOffset - 2), $from.parentOffset);
        if (text === "-" && before.endsWith("-")) {
          view.dispatch(view.state.tr.insertText("—", from - 1, to));
          return true;
        }
        if (text === '"') {
          const prior = $from.parent.textBetween(0, $from.parentOffset);
          const opening = !prior || /[\s([{«—]/u.test(prior.slice(-1));
          view.dispatch(view.state.tr.insertText(opening ? "«" : "»", from, to));
          return true;
        }
        return false;
      },
      handleClick: (_view, pos) => {
        const issue = issuesRef.current.find((item) => item.from <= pos && pos <= item.to);
        selectedIssueRef.current = issue ?? null;
        correctionPendingRef.current = false;
        setSelectedIssue(issue ?? null);
        setCorrectionPending(false);
        correctionIndexRef.current = 0;
        setCorrectionIndex(0);
        if (issue?.kind === "spelling") {
          const language: Language = /[А-Яа-яЁё]/.test(issue.word) ? "ru" : "en";
          void askWorker({ type: "suggest", word: issue.word, language, personal: personalRef.current }).then((result: { suggestions: string[] }) => {
            setSelectedIssue((current) => {
              if (current?.from !== issue.from || current?.to !== issue.to) return current;
              const next = { ...current, suggestions: result.suggestions };
              selectedIssueRef.current = next;
              return next;
            });
          });
        }
        return false;
      },
    };
  const stableEditorProps = useMemo<EditorProps>(() => ({
      handleKeyDown: (view, event) => editorHandlers.current.handleKeyDown?.(view, event) ?? false,
      handleTextInput: (view, from, to, text, deflt) => editorHandlers.current.handleTextInput?.(view, from, to, text, deflt) ?? false,
      handleClick: (view, pos, event) => editorHandlers.current.handleClick?.(view, pos, event) ?? false,
      attributes: {
        lang: "ru",
        "aria-label": "Текст документа",
        spellcheck: "false",
      },
    }), []);
  const editor = useEditor({
    extensions: stableExtensions,
    content: "<p></p>",
    editorProps: stableEditorProps,
    onUpdate: () => {
      if (!suppressUpdate.current) {
        correctionRequest.current++;
        selectedIssueRef.current = null;
        correctionPendingRef.current = false;
        setSelectedIssue(null);
        setCorrectionPending(false);
        setRevision((value) => value + 1);
        documentChannel.current?.postMessage({ type: "document", document: editor?.getJSON() });
      }
    },
    onTransaction: ({ editor: current, transaction }) => {
      setVimMode(getVimMode(current));
      if (transaction.selectionSet || transaction.docChanged) setSelectionRevision((value) => value + 1);
    },
  });
  useEffect(() => {
    if (!selectedIssue) return;
    correctionPanel.current?.focus({ preventScroll: true });
    correctionPanel.current?.querySelector<HTMLButtonElement>(".correction-choice--active")
      ?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [selectedIssue, correctionIndex]);
  useEffect(() => {
    const worker = new Worker(new URL("./text-check.worker.ts", import.meta.url), { type: "module" });
    checkWorker.current = worker;
    worker.onmessage = (event: MessageEvent<{ id: number }>) => {
      const resolve = workerRequests.current.get(event.data.id);
      if (resolve) { workerRequests.current.delete(event.data.id); resolve(event.data); }
    };
    void readPersonalDictionary().then(setPersonal);
    const onDictionary = (event: Event) => setPersonal((event as CustomEvent<PersonalDictionary>).detail);
    addEventListener("personal-dictionary-changed", onDictionary);
    return () => { worker.terminate(); checkWorker.current = null; workerRequests.current.clear(); removeEventListener("personal-dictionary-changed", onDictionary); };
  }, []);
  const askWorker = (message: object): Promise<any> => new Promise((resolve) => {
    const id = ++nextWorkerId.current;
    workerRequests.current.set(id, resolve);
    checkWorker.current?.postMessage({ ...message, id });
  });
  useEffect(() => {
    if (!editor) return;
    let cancelled = false;
    Promise.all([readDraft(documentId), readSettings()]).then(([draft, settings]) => {
      if (cancelled || !draft) return;
      setVimEnabled(settings.vimEnabled);
      setHideMenu(settings.hideMenu);
      setAssistSettings(settings);
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
    const android = isTauri() && /Android/i.test(navigator.userAgent);
    documentChannel.current = channel;
    channel.onmessage = (event: MessageEvent<{ type: string; document?: Draft["document"]; savePath?: string }>) => {
      if (event.data.type === "active") { if (!android) editor.setEditable(false); return; }
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
      if (!android && document.visibilityState === "hidden") return;
      editor.setEditable(true);
      channel.postMessage({ type: "active" });
      setRevision((value) => value + 1);
      requestAnimationFrame(() => editor.commands.focus());
    };
    const onVisible = () => { if (document.visibilityState === "visible") onFocus(); };
    addEventListener("focus", onFocus);
    document.addEventListener("visibilitychange", onVisible);
    const resumed = android ? getCurrentWindow().listen("tauri://resumed", onFocus) : null;
    const focusEditor = isTauri() ? getCurrentWindow().listen("focus-editor", onFocus) : null;
    if (android || document.hasFocus()) onFocus();
    return () => {
      removeEventListener("focus", onFocus);
      document.removeEventListener("visibilitychange", onVisible);
      if (resumed) void resumed.then((unlisten) => unlisten());
      if (focusEditor) void focusEditor.then((unlisten) => unlisten());
      channel.close(); documentChannel.current = null;
    };
  }, [editor, ready, documentId]);
  useEffect(() => {
    if (!editor) return;
    vim.current = new VimAdapter(editor, (key) => {
      if (["s", "v"].includes(key)) void persistCurrent(true).then(() => openDocumentWindow(documentId, documentName, true)).catch(() => {});
      else if (["w", "W", "p", "h", "j", "k", "l"].includes(key)) void focusAdjacentWindow(key === "W" || key === "p" || key === "h" || key === "k" ? -1 : 1);
    }, () => { setCommandLine(""); requestAnimationFrame(() => commandInput.current?.focus()); },
    (backward) => { setSearchLine({ backward, query: "" }); requestAnimationFrame(() => searchInput.current?.focus()); },
    (direction) => {
      const issues = issuesRef.current
        .filter((issue) => issue.from >= 1 && issue.to <= editor.state.doc.content.size &&
          editor.state.doc.textBetween(issue.from, issue.to) === issue.word)
        .sort((left, right) => left.from - right.from);
      if (!issues.length) return;
      const position = editor.state.selection.from;
      const target = direction === 1
        ? issues.find((issue) => issue.from > position) ?? issues[0]
        : [...issues].reverse().find((issue) => issue.from < position) ?? issues[issues.length - 1];
      editor.chain().focus().setTextSelection(target.from).scrollIntoView().run();
    });
    vim.current.setEnabled(vimEnabled);
    setVimMode(vim.current.mode());
    return () => { vim.current = null; };
  }, [editor, vimEnabled, documentId, documentName]);
  useEffect(() => { editor?.setEditable(ready); }, [editor, ready]);
  useEffect(() => {
    const refreshSettings = () => { void readSettings().then((settings) => {
      setVimEnabled(settings.vimEnabled); setHideMenu(settings.hideMenu);
      setAssistSettings(settings);
    }); void readPersonalDictionary().then(setPersonal); if (currentReady.current) requestAnimationFrame(() => editor?.commands.focus()); };
    addEventListener("focus", refreshSettings);
    return () => removeEventListener("focus", refreshSettings);
  }, [editor]);
  useEffect(() => {
    if (!editor || !ready) return;
    setSaveStatus("Сохранение в фоне…");
    void isDocumentSaveOwner(documentId).then((owner) => {
      if (owner) saver.current!.schedule(currentDraft());
    });
  }, [editor, ready, revision, documentId, documentName, viewMode]);

  useEffect(() => {
    if (!editor || !ready || !checkWorker.current) return;
    let cancelled = false;
    const controller = new AbortController();
    const timer = window.setTimeout(() => {
      const segments: { from: number; text: string }[] = [];
      editor.state.doc.descendants((node, pos) => {
        if (node.isText && node.text) segments.push({ from: pos, text: node.text });
      });
      const snapshot = editor.state.doc.textContent;
      void askWorker({ type: "check", segments, personal }).then(async (result: { issues: TextIssue[] }) => {
        if (cancelled || snapshot !== editor.state.doc.textContent) return;
        let all = (result.issues ?? []).filter((issue) => !ignoredIssues.current.has(`${issue.kind}:${issue.word}`));
        issuesRef.current = all;
        editor.view.dispatch(editor.state.tr.setMeta(textCheckKey, all).setMeta("addToHistory", false));
        if (assistSettings.grammarUrl.trim()) {
          try { all = [...all, ...(await checkGrammar(segments, assistSettings.language, assistSettings.grammarUrl, controller.signal))
            .filter((issue) => !ignoredIssues.current.has(`${issue.kind}:${issue.word}`))]; }
          catch { /* Grammar server may be offline; local spelling remains available. */ }
        }
        if (cancelled || snapshot !== editor.state.doc.textContent) return;
        issuesRef.current = all;
        editor.view.dispatch(editor.state.tr.setMeta(textCheckKey, all).setMeta("addToHistory", false));
      });
    }, 500);
    return () => { cancelled = true; controller.abort(); clearTimeout(timer); };
  }, [editor, ready, revision, personal, assistSettings.grammarUrl, assistSettings.language]);

  useEffect(() => {
    if (!editor || !ready || !assistSettings.autocomplete || (vimEnabled && vimMode !== "insert")) {
      setCompletion(null); return;
    }
    const { $from, empty } = editor.state.selection;
    if (!empty || !$from.parent.isTextblock) { setCompletion(null); return; }
    const before = $from.parent.textBetween(0, $from.parentOffset);
    const prefix = before.match(/[A-Za-zА-Яа-яЁё]{3,32}$/u)?.[0];
    if (!prefix) { setCompletion(null); return; }
    const from = $from.pos - prefix.length, to = $from.pos;
    const language: Language = /[А-Яа-яЁё]/.test(prefix) ? "ru" : "en";
    let cancelled = false;
    const timer = window.setTimeout(() => {
      void askWorker({ type: "complete", prefix, language, personal }).then((result: { suggestion: string }) => {
        if (cancelled || !result.suggestion || editor.state.selection.from !== to) return;
        setCompletion({ from, to, suggestion: /^[A-ZА-ЯЁ]/.test(prefix) ? result.suggestion[0].toLocaleUpperCase(language) + result.suggestion.slice(1) : result.suggestion });
      });
    }, 300);
    return () => { cancelled = true; clearTimeout(timer); };
  }, [editor, ready, selectionRevision, vimEnabled, vimMode, assistSettings.autocomplete, personal]);

  const persistCurrent = async (draftIsEnough = false) => {
    if (!editor || !currentReady.current) return;
    if (!(await isDocumentSaveOwner(documentId)) && documentChannel.current) {
      await writeDraft(currentDraft());
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
        if (getCurrentWindow().label.startsWith("doc-")) await invoke("close_task");
        else await showLibrary();
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
    if (isTauri() && /Android/i.test(navigator.userAgent) && getCurrentWindow().label.startsWith("doc-")) {
      await invoke("close_task");
      return;
    }
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
  const correctionChoices = selectedIssue
    ? [...selectedIssue.suggestions, ...(selectedIssue.kind === "spelling" ? ["Добавить в словарь"] : []), "Игнорировать", "Закрыть"]
    : [];

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

      {selectedIssue && <div ref={correctionPanel} className="correction-panel" role="group" tabIndex={-1}
        aria-label="Исправление текста"
        onKeyDown={(event) => {
          const direction = event.ctrlKey && !event.altKey && !event.metaKey && !event.shiftKey
            ? event.code === "KeyN" || ["n", "N", "т", "Т"].includes(event.key) ? 1
              : event.code === "KeyP" || ["p", "P", "з", "З"].includes(event.key) ? -1 : 0
            : 0;
          if (direction) {
            event.preventDefault(); event.stopPropagation();
            cycleCorrection(direction as 1 | -1);
          } else if (event.key === "Tab") {
            event.preventDefault(); event.stopPropagation();
            if (!correctionPendingRef.current) acceptCorrection();
          } else if (event.key === "Escape") {
            event.preventDefault(); event.stopPropagation();
            applyCorrectionChoice(correctionChoices.length - 1);
          }
        }}>
        <span>{selectedIssue.message || `Слово «${selectedIssue.word}» не найдено в словаре`}</span>
        {correctionPending && <span>Подбираем варианты…</span>}
        {!correctionPending && !selectedIssue.suggestions.length && <span>Вариантов нет</span>}
        <span>Выбрано: {correctionChoices[correctionIndex]} · Ctrl+N/P — выбор · Tab — выполнить</span>
        {correctionChoices.map((choice, index) => <button type="button" key={`${choice}-${index}`}
          className={index === correctionIndex ? "correction-choice--active" : ""} aria-pressed={index === correctionIndex}
          onClick={() => applyCorrectionChoice(index)}>{choice}</button>)}
      </div>}

      <footer className="status-bar">
        <span>{vimEnabled ? `Vim ${vimMode}` : "Обычный ввод"}</span>
        {completion && <span className="completion-hint">{completion.suggestion} · Tab / Ctrl+Space</span>}
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
