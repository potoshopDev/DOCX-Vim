import { useEffect, useMemo, useRef, useState } from "react";
import { flushSync } from "react-dom";
import { EditorContent, useEditor, useEditorState } from "@tiptap/react";
import { generateJSON } from "@tiptap/core";
import type { EditorState } from "@tiptap/pm/state";
import type { EditorProps } from "@tiptap/pm/view";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { onBackButtonPress } from "@tauri-apps/api/app";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import { emit, emitTo, listen } from "@tauri-apps/api/event";
import { save } from "@tauri-apps/plugin-dialog";
import { BaseDirectory, readFile, writeFile } from "@tauri-apps/plugin-fs";
import { exportDocx, importDocx } from "./docx";
import { BackgroundSaveFailure, BackgroundSaver } from "./background-save";
import { CloudSync, cloudProviders, rememberOpenedRemote, type CloudProvider, type CloudStates } from "./cloud-sync";
import { downloadYandexDocx } from "./cloud-remote";
import { readDraft, writeDraft, type Draft } from "./draft";
import { rememberDocument } from "./library";
import { editorExtensions } from "./editor-extensions";
import { fingerprint, preserveDraftVersion, replaceSourceDraft } from "./source-version";
import { readSettings } from "./settings";
import { defaultSettings, type Settings } from "./settings";
import { readPersonalDictionary, writePersonalDictionary, type PersonalDictionary, type Language } from "./personal-dictionary";
import { TextCheck, textCheckKey, type TextIssue } from "./text-check-extension";
import { checkGrammar } from "./grammar";
import { VimMode as VimExtension, getVimMode } from "vim-prosemirror/tiptap";
import "vim-prosemirror/style.css";
import { VimAdapter, type VimMode } from "./vim-adapter";
import { focusAdjacentWindow, isDocumentSaveOwner, openDocumentWindow, openWindowIds, registerDocumentTask, showLibrary, showYandexLibrary } from "./windows";
import "./App.css";

type ViewMode = "page" | "mobile";
type SessionSnapshot = { revision: number; editorLabel: string; draft: Draft; unconfirmedFrom?: string | null };
type CorrectionItem = TextIssue | { kind: "none" | "pending"; from: number; to: number; word: string; suggestions: string[]; message: string };

function wordAtCaret(state: EditorState): { from: number; to: number; word: string } | null {
  if (!state.selection.empty) return null;
  const $from = state.selection.$from;
  if (!$from.parent.isTextblock) return null;
  const before = $from.nodeBefore?.text ?? "";
  const after = $from.nodeAfter?.text ?? "";
  if (after && !/^[A-Za-zА-Яа-яЁё]/u.test(after)) return null;
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
  const [sourceLabel, setSourceLabel] = useState("Черновик");
  const [ready, setReady] = useState(false);
  const [revision, setRevision] = useState(0);
  const [contentRevision, setContentRevision] = useState(0);
  const [localConflict, setLocalConflict] = useState(false);
  const [saveStatus, setSaveStatus] = useState("Подготовка…");
  const [lastLocalSave, setLastLocalSave] = useState<{ path: string; updatedAt: string } | null>(null);
  const [busy, setBusy] = useState(false);
  const [cloudStates, setCloudStates] = useState<CloudStates>({});
  const [vimEnabled, setVimEnabled] = useState(true);
  const [hideMenu, setHideMenu] = useState(false);
  const [assistSettings, setAssistSettings] = useState<Settings>(defaultSettings);
  const [personal, setPersonal] = useState<PersonalDictionary>({ ru: [], en: [] });
  const [selectedIssue, setSelectedIssue] = useState<CorrectionItem | null>(null);
  const [correctionPending, setCorrectionPending] = useState(false);
  const [correctionIndex, setCorrectionIndex] = useState(0);
  const [completion, setCompletion] = useState<{ from: number; to: number; prefix: string; suggestions: string[]; index: number } | null>(null);
  const [selectionRevision, setSelectionRevision] = useState(0);
  const issuesRef = useRef<TextIssue[]>([]);
  const selectedIssueRef = useRef<CorrectionItem | null>(null);
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
  const sessionRevision = useRef(-1);
  const latestDraftUpdatedAt = useRef("");
  const hasEditLease = useRef(false);
  const cloudHandoffBlocked = useRef(false);
  const handoffRecoveryId = useRef<string | null>(null);
  const handoffBackupTimer = useRef<number | null>(null);
  const publicationQueue = useRef<Promise<void>>(Promise.resolve());
  const publicationFailed = useRef(false);
  const suppressUpdate = useRef(false);
  const editorArea = useRef<HTMLElement>(null);
  const currentId = useRef(documentId);
  const currentName = useRef(documentName);
  const currentSourcePath = useRef<string | undefined>(undefined);
  const currentRestoreSourcePath = useRef<string | undefined>(undefined);
  const currentSourceFingerprint = useRef<string | undefined>(undefined);
  const currentSavePath = useRef<string | undefined>(undefined);
  const currentCloudTargets = useRef<Draft["cloudTargets"]>({});
  const currentViewMode = useRef<ViewMode>("mobile");
  const currentReady = useRef(false);
  const conflictInProgress = useRef(false);
  const saver = useRef<BackgroundSaver | null>(null);
  const cloudSync = useRef<CloudSync | null>(null);
  const recordSaved = (draft: Draft) => {
    setSaveStatus(`${draft.savePath ? "DOCX и черновик сохранены" : "Черновик сохранён"} ${new Date().toLocaleTimeString("ru-RU")}`);
    if (draft.savePath && draft.savePath === draft.sourcePath && !draft.sourcePath.startsWith("remote:")) {
      setLastLocalSave({ path: draft.savePath, updatedAt: draft.updatedAt });
      if (isTauri()) void emit("docx-vim-local-saved", { id: draft.id, path: draft.savePath, updatedAt: draft.updatedAt }).catch(() => {});
    }
  };
  const reportCloudStates = (states: CloudStates) => {
    setCloudStates(states);
    if (isTauri() && hasEditLease.current)
      void emit("docx-vim-cloud-states", { id: documentId, states }).catch(() => {});
  };
  if (!cloudSync.current) cloudSync.current = new CloudSync(documentId, reportCloudStates);
  if (!saver.current) saver.current = new BackgroundSaver(recordSaved,
  800, (draft, bytes) => { if (!cloudHandoffBlocked.current) cloudSync.current?.schedule(draft, bytes); }, () => setLocalConflict(true),
  (revision) => { currentSourceFingerprint.current = revision; });
  const currentDraft = (): Draft => ({
    version: 1, id: currentId.current, name: currentName.current,
    document: editor!.getJSON(), updatedAt: new Date(Math.max(Date.now(), (Date.parse(latestDraftUpdatedAt.current) || 0) + 1)).toISOString(),
    sourcePath: currentSourcePath.current, savePath: currentSavePath.current,
    restoreSourcePath: currentRestoreSourcePath.current,
    sourceFingerprint: currentSourceFingerprint.current,
    cloudTargets: currentCloudTargets.current,
    viewMode: currentViewMode.current,
    selection: editor ? { anchor: editor.state.selection.anchor, head: editor.state.selection.head } : undefined,
    scrollTop: editorArea.current?.scrollTop,
  });
  const writeHandoffBackup = (draft: Draft): Promise<void> => {
    handoffRecoveryId.current ??= crypto.randomUUID();
    return writeDraft({ ...draft, id: handoffRecoveryId.current,
      name: `${draft.name.replace(/\.docx$/i, "")} — версия при переключении окон`,
      sourcePath: undefined, restoreSourcePath: draft.sourcePath ?? draft.restoreSourcePath,
      sourceFingerprint: undefined, savePath: undefined, cloudTargets: {}, updatedAt: new Date().toISOString() });
  };
  const scheduleHandoffBackup = (draft: Draft) => {
    if (handoffBackupTimer.current !== null) clearTimeout(handoffBackupTimer.current);
    handoffBackupTimer.current = window.setTimeout(() => {
      handoffBackupTimer.current = null;
      void writeHandoffBackup(draft).catch(() => {});
    }, 500);
  };
  const publishCurrent = (): Promise<void> => {
    if (!isTauri() || !hasEditLease.current || !currentReady.current) return publicationQueue.current;
    const draft = currentDraft();
    latestDraftUpdatedAt.current = draft.updatedAt;
    publicationQueue.current = publicationQueue.current.then(async () => {
      const snapshot = await invoke<SessionSnapshot>("publish_document_session", { id: documentId, draft });
      sessionRevision.current = Math.max(sessionRevision.current, snapshot.revision);
    }).catch(async () => {
      publicationFailed.current = true;
      hasEditLease.current = false;
      editor?.setEditable(false);
      setSaveStatus("Не удалось передать изменение между окнами");
      await preserveDraftVersion(draft).catch(() => {});
      throw new Error("Document publication failed");
    });
    return publicationQueue.current;
  };
  const correctionChoiceCount = (issue: CorrectionItem) => issue.kind === "none" || issue.kind === "pending" ? 1 : issue.suggestions.length + (issue.kind === "spelling" ? 3 : 2);
  const applyCorrectionChoice = (index: number): boolean => {
    const issue = selectedIssueRef.current;
    if (!issue || !editor || (isTauri() && !hasEditLease.current)) return false;
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
    if (!target) {
      correctionRequest.current++;
      const caret = editor.state.selection.from;
      const notice: CorrectionItem = { kind: "none", from: caret, to: caret, word: "", suggestions: [], message: "Ошибок нет" };
      selectedIssueRef.current = notice;
      correctionPendingRef.current = false;
      correctionIndexRef.current = 0;
      flushSync(() => { setSelectedIssue(notice); setCorrectionIndex(0); setCorrectionPending(false); });
      return;
    }
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
    const pending: CorrectionItem = { ...target, kind: "pending", suggestions: [], message: "Проверяем слово…" };
    selectedIssueRef.current = pending;
    correctionPendingRef.current = true;
    pendingCorrectionOffset.current = direction === -1 ? -1 : 0;
    correctionIndexRef.current = 0;
    flushSync(() => { setSelectedIssue(pending); setCorrectionIndex(0); setCorrectionPending(true); });
    void askWorker({ type: "suggest", word: target.word, language, personal: personalRef.current })
      .then((result: { suggestions?: string[]; correct?: boolean }) => {
        if (request !== correctionRequest.current || !editor || editor.state.selection.from !== caret ||
          editor.state.doc.textBetween(target.from, target.to) !== target.word) return;
        if (result.correct) {
          const notice: CorrectionItem = { ...target, kind: "none", suggestions: [], message: "Ошибок нет" };
          selectedIssueRef.current = notice;
          correctionPendingRef.current = false;
          correctionIndexRef.current = 0;
          flushSync(() => { setSelectedIssue(notice); setCorrectionIndex(0); setCorrectionPending(false); });
          return;
        }
        const next: TextIssue = { ...(issue ?? { ...target, kind: "spelling" }), suggestions: result.suggestions ?? [] };
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
        if (isTauri() && (!hasEditLease.current || !_view.editable)) { event.preventDefault(); return true; }
        const editable = !vim.current?.enabled || vim.current?.mode() === "insert";
        const candidate = completionRef.current;
        const activeCompletion = candidate && _view.state.selection.from === candidate.to &&
          _view.state.doc.textBetween(candidate.from, candidate.to) === candidate.prefix ? candidate : null;
        const correctionKey = event.ctrlKey && !event.altKey && !event.metaKey && !event.shiftKey
          ? event.code === "KeyN" || ["n", "N", "т", "Т"].includes(event.key) ? 1
            : event.code === "KeyP" || ["p", "P", "з", "З"].includes(event.key) ? -1 : 0
          : 0;
        if (correctionKey) {
          event.preventDefault();
          event.stopPropagation();
          const direction = correctionKey as 1 | -1;
          const currentCompletion = activeCompletion;
          if (editable && !selectedIssueRef.current && currentCompletion) {
            const next = { ...currentCompletion,
              index: (currentCompletion.index + direction + currentCompletion.suggestions.length) % currentCompletion.suggestions.length };
            completionRef.current = next;
            setCompletion(next);
            return true;
          }
          if (!cycleCorrection(direction)) showCorrectionsAtCaret(direction);
          return true;
        }
        const correction = selectedIssueRef.current;
        if ((event.key === "Tab" || event.key === "Enter" || event.key === " ") && correction && !event.ctrlKey && !event.altKey && !event.metaKey) {
          event.preventDefault();
          event.stopPropagation();
          if (!correctionPendingRef.current || correction.kind === "pending") acceptCorrection();
          return true;
        }
        if (event.key === "Tab" && (!editable || !activeCompletion)) {
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
        if (editable && activeCompletion && (event.key === "Tab" || (event.ctrlKey && event.code === "Space"))) {
          event.preventDefault();
          const item = activeCompletion;
          editor?.chain().focus().insertContentAt({ from: item.from, to: item.to }, item.suggestions[item.index]).run();
          setCompletion(null);
          return true;
        }
        if (editable && event.key === "Escape" && completionRef.current) { setCompletion(null); return true; }
        if (editable && event.ctrlKey && event.code === "Space") { setSelectionRevision((value) => value + 1); event.preventDefault(); return true; }
        return vim.current?.handle(event) ?? false;
      },
      handleTextInput: (view, from, to, text) => {
        if (isTauri() && !hasEditLease.current) return true;
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
        if (isTauri() && !hasEditLease.current) return false;
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
    editable: false,
    editorProps: stableEditorProps,
    onUpdate: () => {
      if (!suppressUpdate.current) {
        correctionRequest.current++;
        selectedIssueRef.current = null;
        correctionPendingRef.current = false;
        setSelectedIssue(null);
        setCorrectionPending(false);
        setRevision((value) => value + 1);
        setContentRevision((value) => value + 1);
        if (isTauri()) void publishCurrent().catch(() => {});
        else documentChannel.current?.postMessage({ type: "document", document: editor?.getJSON() });
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
    worker.postMessage({ id: 0, type: "preload" });
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
    if (!isTauri()) return;
    const listener = listen("source-reload", () => location.reload());
    return () => { void listener.then((unlisten) => unlisten()); };
  }, []);
  useEffect(() => {
    if (!editor) return;
    let cancelled = false;
    Promise.all([readDraft(documentId), readSettings()]).then(async ([storedDraft, settings]) => {
      if (cancelled || !storedDraft) return;
      let draft = storedDraft;
      let session = isTauri() ? await invoke<SessionSnapshot | null>("get_document_session", { id: documentId }) : null;
      if (isTauri() && !session && !(await isDocumentSaveOwner(documentId))) {
        for (let attempt = 0; attempt < 100 && !cancelled && !session; attempt++) {
          await new Promise((resolve) => window.setTimeout(resolve, 100));
          session = await invoke<SessionSnapshot | null>("get_document_session", { id: documentId });
          if (await isDocumentSaveOwner(documentId)) break;
        }
      }
      if (isTauri() && !session && !(await isDocumentSaveOwner(documentId))) throw new Error("Document owner has not opened the session");
      if (!session && draft.sourcePath && await isDocumentSaveOwner(documentId)) {
        try {
          const source = draft.sourcePath;
          const remote = source.startsWith("remote:yandexDisk:");
          const { bytes, revision } = remote
            ? await downloadYandexDocx(source)
            : { bytes: await readFile(source), revision: undefined };
          const imported = await importDocx(bytes);
          if (cancelled) return;
          draft = {
            ...draft,
            document: generateJSON(imported.html || "<p></p>", editorExtensions),
            sourceFingerprint: remote ? undefined : await fingerprint(bytes),
            savePath: remote ? undefined : source,
            cloudTargets: remote ? { yandexDisk: source } : {},
            updatedAt: new Date().toISOString(),
          };
          await replaceSourceDraft(draft);
          if (remote && revision) await rememberOpenedRemote(documentId, "yandexDisk", source, revision);
        } catch {
          await showLibrary().catch(() => {});
          return;
        }
      }
      if (isTauri()) {
        session ??= await invoke<SessionSnapshot>("join_document_session", { id: documentId, draft });
        draft = session.draft;
        sessionRevision.current = session.revision;
        hasEditLease.current = session.editorLabel === getCurrentWindow().label;
        cloudHandoffBlocked.current = !!session.unconfirmedFrom;
      }
      latestDraftUpdatedAt.current = draft.updatedAt;
      await cloudSync.current?.load();
      if (cancelled) return;
      setVimEnabled(settings.vimEnabled);
      setHideMenu(settings.hideMenu);
      setAssistSettings(settings);
      saver.current = new BackgroundSaver(recordSaved, settings.autosaveMs, (savedDraft, bytes) => {
        if (!cloudHandoffBlocked.current) cloudSync.current?.schedule(savedDraft, bytes);
      },
      () => setLocalConflict(true), (revision) => {
        currentSourceFingerprint.current = revision;
        if (isTauri()) void invoke("update_document_source_fingerprint", { id: documentId, fingerprint: revision }).catch(() => {});
      });
      suppressUpdate.current = true;
      editor.chain().setMeta("addToHistory", false).setContent(draft.document, { emitUpdate: false }).run();
      suppressUpdate.current = false;
      currentName.current = draft.name;
      currentSourcePath.current = draft.sourcePath;
      currentRestoreSourcePath.current = draft.restoreSourcePath;
      currentSourceFingerprint.current = draft.sourceFingerprint;
      currentSavePath.current = draft.savePath;
      currentCloudTargets.current = draft.cloudTargets ?? {};
      setDocumentName(draft.name);
      setSourceLabel(draft.sourcePath?.startsWith("remote:yandexDisk:") ? "Яндекс Диск" : draft.sourcePath ? "Локальный" : "Черновик");
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
      if (!isTauri() || hasEditLease.current) requestAnimationFrame(() => editor.commands.focus());
      setSaveStatus("Готово");
      void registerDocumentTask(isTauri() ? getCurrentWindow().label.slice(4) : documentId, draft.name);
      if (draft.scrollTop != null) requestAnimationFrame(() => {
        if (editorArea.current) editorArea.current.scrollTop = draft.scrollTop!;
      });
      if (isTauri()) void getCurrentWindow().setTitle(draft.name);
    }).catch(() => setSaveStatus(""));
    return () => { cancelled = true; };
  }, [editor, documentId]);
  const applySessionSnapshot = (snapshot: SessionSnapshot, force = false) => {
    if (!editor || snapshot.draft.id !== documentId || (!force && snapshot.revision <= sessionRevision.current)) return;
    sessionRevision.current = snapshot.revision;
    latestDraftUpdatedAt.current = snapshot.draft.updatedAt;
    cloudHandoffBlocked.current = !!snapshot.unconfirmedFrom;
    const draft = snapshot.draft;
    currentName.current = draft.name;
    currentSourcePath.current = draft.sourcePath;
    currentRestoreSourcePath.current = draft.restoreSourcePath;
    currentSourceFingerprint.current = draft.sourceFingerprint;
    currentSavePath.current = draft.savePath;
    currentCloudTargets.current = draft.cloudTargets ?? {};
    setDocumentName(draft.name);
    setSourceLabel(draft.sourcePath?.startsWith("remote:yandexDisk:") ? "Яндекс Диск" : draft.sourcePath ? "Локальный" : "Черновик");
    if (force || snapshot.editorLabel !== getCurrentWindow().label) {
      const anchor = editor.state.selection.anchor;
      const head = editor.state.selection.head;
      suppressUpdate.current = true;
      try {
        editor.chain().setMeta("addToHistory", false).setContent(draft.document, { emitUpdate: false }).run();
        const max = editor.state.doc.content.size;
        editor.commands.setTextSelection({ from: Math.min(anchor, max), to: Math.min(head, max) });
      } finally { suppressUpdate.current = false; }
      setRevision((value) => value + 1);
      setContentRevision((value) => value + 1);
    }
  };
  useEffect(() => {
    if (!editor || !ready) return;
    if (isTauri()) {
      const label = getCurrentWindow().label;
      let stopped = false;
      let acquiring = false;
      let yielding = false;
      let queuedSaveRevision = sessionRevision.current;
      let wasSaveOwner = hasEditLease.current;
      const checkSaveOwner = async () => {
        if (yielding) return;
        const owner = await isDocumentSaveOwner(documentId);
        if (stopped || yielding) return;
        if (owner && wasSaveOwner && cloudHandoffBlocked.current) {
          const snapshot = await invoke<SessionSnapshot>("get_document_session", { id: documentId });
          if (!snapshot.unconfirmedFrom) {
            cloudHandoffBlocked.current = false;
            applySessionSnapshot(snapshot);
            saver.current?.schedule(snapshot.draft);
          }
        }
        if (owner && wasSaveOwner !== true) {
          if (wasSaveOwner === false) {
            cloudSync.current = new CloudSync(documentId, reportCloudStates);
            await cloudSync.current.load();
          }
          const snapshot = await invoke<SessionSnapshot>("get_document_session", { id: documentId });
          cloudHandoffBlocked.current = !!snapshot.unconfirmedFrom;
          if (snapshot.revision > 0 || wasSaveOwner === false) {
            queuedSaveRevision = Math.max(queuedSaveRevision, snapshot.revision);
            if (snapshot.unconfirmedFrom) scheduleHandoffBackup(snapshot.draft);
            else saver.current?.schedule(snapshot.draft);
          }
          if (wasSaveOwner === false && document.visibilityState === "visible") void onFocus();
        }
        wasSaveOwner = owner;
      };
      const onSnapshot = (snapshot: SessionSnapshot) => {
        if (snapshot.draft.id !== documentId || snapshot.revision < sessionRevision.current) return;
        cloudHandoffBlocked.current = !!snapshot.unconfirmedFrom;
        if (snapshot.editorLabel !== label && hasEditLease.current) {
          hasEditLease.current = false;
          editor.setEditable(false);
          wasSaveOwner = false;
          void cloudSync.current?.retire().catch(() => {});
          setRevision((value) => value + 1);
        }
        if (snapshot.unconfirmedFrom === label && snapshot.editorLabel !== label)
          void cloudSync.current?.retire().then(() => invoke("confirm_document_handoff", { id: documentId })).catch(() => {});
        applySessionSnapshot(snapshot);
        if (snapshot.revision > queuedSaveRevision) {
          queuedSaveRevision = snapshot.revision;
          void isDocumentSaveOwner(documentId).then((owner) => {
            if (owner && !stopped && !yielding && snapshot.revision === queuedSaveRevision) {
              if (snapshot.unconfirmedFrom) scheduleHandoffBackup(snapshot.draft);
              else saver.current?.schedule(snapshot.draft);
            }
          });
        }
      };
      const onFocus = async () => {
        if (stopped || yielding || publicationFailed.current || conflictInProgress.current || (document.visibilityState === "hidden" && !/Android/i.test(navigator.userAgent)) || acquiring) return;
        acquiring = true;
        editor.setEditable(false);
        try {
          const current = await invoke<SessionSnapshot>("get_document_session", { id: documentId });
          cloudHandoffBlocked.current = !!current.unconfirmedFrom;
          if (current.unconfirmedFrom === label && current.editorLabel !== label) {
            await cloudSync.current?.retire();
            await invoke("confirm_document_handoff", { id: documentId });
            wasSaveOwner = false;
          }
          if (current.editorLabel === label) {
            applySessionSnapshot(current, current.revision > sessionRevision.current || !hasEditLease.current);
            hasEditLease.current = true;
            editor.setEditable(true);
            setRevision((value) => value + 1);
            requestAnimationFrame(() => editor.commands.focus());
            return;
          }
          hasEditLease.current = false;
          const snapshot = await invoke<SessionSnapshot>("acquire_document_session", { id: documentId });
          if (stopped) return;
          applySessionSnapshot(snapshot, true);
          hasEditLease.current = true;
          await checkSaveOwner();
          editor.setEditable(true);
          setRevision((value) => value + 1);
          requestAnimationFrame(() => editor.commands.focus());
        } catch { if (!stopped) setSaveStatus("Не удалось передать право редактирования"); }
        finally { acquiring = false; }
      };
      const onYield = async (payload: { id: string; editorLabel: string }) => {
        if (payload.id !== documentId || payload.editorLabel !== label || !hasEditLease.current || yielding) return;
        yielding = true;
        editor.setEditable(false);
        setRevision((value) => value + 1);
        selectedIssueRef.current = null;
        setSelectedIssue(null);
        setCompletion(null);
        setCommandLine(null);
        setSearchLine(null);
        try {
          await publicationQueue.current;
          const snapshot = await invoke<SessionSnapshot>("get_document_session", { id: documentId });
          await saver.current?.flush(snapshot.draft);
          if (currentSourceFingerprint.current) await invoke("update_document_source_fingerprint", {
            id: documentId, fingerprint: currentSourceFingerprint.current,
          });
          await cloudSync.current?.retire();
          wasSaveOwner = false;
          await invoke("release_document_session", { id: documentId });
          hasEditLease.current = false;
          yielding = false;
        } catch { setSaveStatus("Не удалось передать изменение между окнами"); }
      };
      const onSaveRequest = async (payload: { id: string; token: string; replyTo: string }) => {
        if (payload.id !== documentId) return;
        let ok = false;
        try {
          const snapshot = await invoke<SessionSnapshot>("get_document_session", { id: documentId });
          if (!yielding && snapshot.editorLabel === label) {
            if (snapshot.unconfirmedFrom) await writeHandoffBackup(snapshot.draft);
            else { await saver.current?.flush(snapshot.draft); ok = true; }
          }
        } catch { /* The requester keeps its window open when saving fails. */ }
        await emitTo(payload.replyTo, "docx-vim-save-result", { token: payload.token, ok }).catch(() => {});
      };
      const listeners = Promise.all([
        listen<SessionSnapshot>("docx-vim-session", (event) => onSnapshot(event.payload)),
        listen<SessionSnapshot>("docx-vim-lease", (event) => {
          if (event.payload.draft.id !== documentId) return;
          if (event.payload.editorLabel !== label) { hasEditLease.current = false; editor.setEditable(false); }
          if (event.payload.unconfirmedFrom === label && event.payload.editorLabel !== label)
            void cloudSync.current?.retire().then(() => invoke("confirm_document_handoff", { id: documentId })).catch(() => {});
          setRevision((value) => value + 1);
          applySessionSnapshot(event.payload);
        }),
        listen<SessionSnapshot>("docx-vim-fingerprint", (event) => {
          if (event.payload.draft.id === documentId) currentSourceFingerprint.current = event.payload.draft.sourceFingerprint;
        }),
        listen<{ id: string; path: string; updatedAt: string }>("docx-vim-local-saved", (event) => {
          if (event.payload.id === documentId) setLastLocalSave({ path: event.payload.path, updatedAt: event.payload.updatedAt });
        }),
        listen<{ id: string; states: CloudStates }>("docx-vim-cloud-states", (event) => {
          if (event.payload.id === documentId) setCloudStates(event.payload.states);
        }),
        listen<SessionSnapshot>("docx-vim-handoff-confirmed", (event) => {
          if (event.payload.draft.id !== documentId || event.payload.revision < sessionRevision.current) return;
          cloudHandoffBlocked.current = false;
          applySessionSnapshot(event.payload);
          setRevision((value) => value + 1);
          void isDocumentSaveOwner(documentId).then((owner) => {
            if (owner && !stopped) saver.current?.schedule(event.payload.draft);
          });
        }),
        listen<{ id: string; editorLabel: string }>("docx-vim-yield", (event) => { void onYield(event.payload); }),
        getCurrentWindow().listen<{ id: string; token: string; replyTo: string }>("docx-vim-save-request", (event) => { void onSaveRequest(event.payload); }),
        getCurrentWindow().listen("focus-editor", () => { void onFocus(); }),
        getCurrentWindow().listen("tauri://resumed", () => { void onFocus(); }),
      ]).then(async (unlisteners) => {
        if (stopped) { unlisteners.forEach((stop) => stop()); return []; }
        const latest = await invoke<SessionSnapshot>("get_document_session", { id: documentId });
        onSnapshot(latest);
        void checkSaveOwner().catch(() => {});
        void onFocus();
        return unlisteners;
      });
      const ownerTimer = window.setInterval(() => { void checkSaveOwner().catch(() => {}); }, 2_000);
      const focus = () => { void onFocus(); };
      const visible = () => { if (document.visibilityState === "visible") void onFocus(); };
      addEventListener("focus", focus);
      document.addEventListener("visibilitychange", visible);
      return () => {
        stopped = true;
        clearInterval(ownerTimer);
        if (handoffBackupTimer.current !== null) clearTimeout(handoffBackupTimer.current);
        removeEventListener("focus", focus);
        document.removeEventListener("visibilitychange", visible);
        void listeners.then((unlisteners) => unlisteners.forEach((stop) => stop()));
      };
    }
    const channel = new BroadcastChannel(`docx-vim-${documentId}`);
    const android = isTauri() && /Android/i.test(navigator.userAgent);
    documentChannel.current = channel;
    channel.onmessage = (event: MessageEvent<{ type: string; document?: Draft["document"]; savePath?: string; cloudTargets?: Draft["cloudTargets"]; provider?: CloudProvider; path?: string }>) => {
      if (event.data.type === "source-reload") { location.reload(); return; }
      if (event.data.type === "active") { if (!android) editor.setEditable(false); return; }
      if (event.data.type === "cloud-rebind" && event.data.provider && event.data.path) {
        void isDocumentSaveOwner(documentId).then((owner) => { if (owner) cloudSync.current?.rebind(event.data.provider!, event.data.path!); });
        return;
      }
      if (event.data.type === "cloud-disconnect" && event.data.provider) {
        void isDocumentSaveOwner(documentId).then((owner) => { if (owner) cloudSync.current?.disconnect(event.data.provider!); });
        return;
      }
      if (event.data.type === "save-request") {
        void isDocumentSaveOwner(documentId).then((owner) => {
          if (!owner) return;
          if (event.data.savePath) currentSavePath.current = event.data.savePath;
          if (event.data.cloudTargets) currentCloudTargets.current = event.data.cloudTargets;
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
      setContentRevision((value) => value + 1);
    };
    const onFocus = () => {
      if (!android && document.visibilityState === "hidden") return;
      if (conflictInProgress.current) return;
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
  useEffect(() => { if (!isTauri()) editor?.setEditable(ready); }, [editor, ready]);
  useEffect(() => {
    const refreshSettings = () => { void readSettings().then((settings) => {
      setVimEnabled(settings.vimEnabled); setHideMenu(settings.hideMenu);
      setAssistSettings(settings);
    }); void readPersonalDictionary().then(setPersonal); if (currentReady.current && (!isTauri() || hasEditLease.current)) requestAnimationFrame(() => editor?.commands.focus()); };
    addEventListener("focus", refreshSettings);
    return () => removeEventListener("focus", refreshSettings);
  }, [editor]);
  useEffect(() => {
    if (!editor || !ready || contentRevision === 0 || conflictInProgress.current) return;
    if (isTauri()) return;
    setSaveStatus("Сохранение в фоне…");
    void isDocumentSaveOwner(documentId).then((owner) => {
      if (owner) saver.current!.schedule(currentDraft());
    });
  }, [editor, ready, contentRevision, documentId]);
  useEffect(() => {
    const source = currentSourcePath.current;
    const cloudConflict = source?.startsWith("remote:yandexDisk:") && cloudStates.yandexDisk?.status === "conflict";
    if (!editor || !ready || (!localConflict && !cloudConflict) || conflictInProgress.current || !source) return;
    conflictInProgress.current = true;
    editor.setEditable(false);
    setSaveStatus("Сохраняем отдельную версию черновика…");
    void (async () => {
      let preserved = false;
      try {
        if (!(await isDocumentSaveOwner(documentId))) return;
        await saver.current?.suspend(documentId);
        const authoritative = isTauri()
          ? (await invoke<SessionSnapshot>("get_document_session", { id: documentId })).draft
          : currentDraft();
        await preserveDraftVersion(authoritative);
        preserved = true;
        const remote = source.startsWith("remote:yandexDisk:");
        const { bytes, revision } = remote
          ? await downloadYandexDocx(source)
          : { bytes: await readFile(source), revision: undefined };
        const imported = await importDocx(bytes);
        const mainDraft: Draft = {
          ...authoritative,
          document: generateJSON(imported.html || "<p></p>", editorExtensions),
          sourcePath: source,
          sourceFingerprint: remote ? undefined : await fingerprint(bytes),
          savePath: remote ? undefined : source,
          cloudTargets: remote ? { yandexDisk: source } : {},
          selection: undefined,
          scrollTop: 0,
          updatedAt: new Date().toISOString(),
        };
        await writeDraft(mainDraft);
        if (remote && revision) await cloudSync.current?.acceptOpenedRemote("yandexDisk", source, revision);
        if (isTauri()) {
          await invoke("reset_document_session", { id: documentId, draft: mainDraft });
          for (const other of await WebviewWindow.getAll()) {
            if (other.label !== getCurrentWindow().label && (other.label === `doc-${documentId}` || other.label.startsWith(`doc-${documentId}--`)))
              void emitTo(other.label, "source-reload").catch(() => {});
          }
        } else documentChannel.current?.postMessage({ type: "source-reload" });
        location.reload();
      } catch {
        setSaveStatus(preserved ? "Версия сохранена в черновиках" : "");
        if (preserved) {
          await showLibrary().catch(() => {});
          if (isTauri() && getCurrentWindow().label.startsWith("doc-")) {
            if (/Android/i.test(navigator.userAgent)) await invoke("close_task").catch(() => {});
            else await getCurrentWindow().destroy().catch(() => {});
          }
        }
      }
    })();
  }, [editor, ready, localConflict, cloudStates.yandexDisk?.status, documentId]);
  useEffect(() => {
    if (ready && editor) void isDocumentSaveOwner(documentId).then((owner) => {
      if (!owner) return;
      if (isTauri()) void invoke<SessionSnapshot>("get_document_session", { id: documentId })
        .then((snapshot) => snapshot.unconfirmedFrom
          ? writeHandoffBackup(snapshot.draft)
          : writeDraft({ ...snapshot.draft, viewMode })).catch(() => {});
      else void writeDraft(currentDraft()).catch(() => {});
    });
  }, [ready, viewMode]);
  useEffect(() => {
    const retry = () => cloudSync.current?.retry();
    addEventListener("online", retry);
    addEventListener("focus", retry);
    return () => { removeEventListener("online", retry); removeEventListener("focus", retry); };
  }, []);

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
    const prefix = before.match(/[A-Za-zА-Яа-яЁё]{2,32}$/u)?.[0];
    if (!prefix) { setCompletion(null); return; }
    const from = $from.pos - prefix.length, to = $from.pos;
    const language: Language = /[А-Яа-яЁё]/.test(prefix) ? "ru" : "en";
    let cancelled = false;
    const timer = window.setTimeout(() => {
      void askWorker({ type: "complete", prefix, language, personal }).then((result: { suggestions?: string[] }) => {
        if (cancelled || !result.suggestions?.length || editor.state.selection.from !== to ||
          editor.state.doc.textBetween(from, to) !== prefix) return;
        const suggestions = result.suggestions.map((word) => /^[A-ZА-ЯЁ]/.test(prefix)
          ? word[0].toLocaleUpperCase(language) + word.slice(1) : word);
        setCompletion({ from, to, prefix, suggestions, index: 0 });
      });
    }, 90);
    return () => { cancelled = true; clearTimeout(timer); };
  }, [editor, ready, selectionRevision, vimEnabled, vimMode, assistSettings.autocomplete, personal]);

  const persistCurrent = async (draftIsEnough = false) => {
    if (!editor || !currentReady.current) return;
    if (isTauri()) {
      if (hasEditLease.current) await publishCurrent();
      await publicationQueue.current;
      const snapshot = await invoke<SessionSnapshot>("get_document_session", { id: documentId });
      if (snapshot.unconfirmedFrom) {
        await writeHandoffBackup(snapshot.draft);
        if (draftIsEnough) return;
        throw new Error("Waiting for previous window to finish saving");
      }
      if (await isDocumentSaveOwner(documentId)) {
        try { await saver.current!.flush(snapshot.draft); }
        catch (error) {
          if (!draftIsEnough || !(error instanceof BackgroundSaveFailure && error.draftSaved)) throw error;
        }
        return;
      }
      if (!snapshot.editorLabel) throw new Error("Save owner missing");
      const token = crypto.randomUUID();
      await new Promise<void>((resolve, reject) => {
        let stop: (() => void) | undefined;
        const timer = window.setTimeout(() => { stop?.(); reject(new Error("Save request timed out")); }, 20_000);
        void getCurrentWindow().listen<{ token: string; ok: boolean }>("docx-vim-save-result", (event) => {
          if (event.payload.token !== token) return;
          clearTimeout(timer); stop?.();
          if (event.payload.ok) resolve(); else reject(new Error("Save request failed"));
        }).then((unlisten) => {
          stop = unlisten;
          void emitTo(snapshot.editorLabel, "docx-vim-save-request", { id: documentId, token, replyTo: getCurrentWindow().label })
            .catch((error) => { clearTimeout(timer); stop?.(); reject(error); });
        }).catch((error) => { clearTimeout(timer); reject(error); });
      });
      return;
    }
    if (!(await isDocumentSaveOwner(documentId)) && documentChannel.current) {
      await writeDraft(currentDraft());
      documentChannel.current.postMessage({ type: "save-request", savePath: currentSavePath.current, cloudTargets: currentCloudTargets.current });
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
        if (getCurrentWindow().label.startsWith("doc-")) {
          await invoke("leave_document_session", { id: documentId });
          await invoke("close_task");
        }
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
          if (owner) void (async () => {
            const snapshot = isTauri()
              ? await invoke<SessionSnapshot>("get_document_session", { id: documentId })
              : null;
            if (snapshot?.unconfirmedFrom) await writeHandoffBackup(snapshot.draft);
            else await saver.current!.flush(snapshot?.draft ?? currentDraft());
          })().catch(() => {});
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
        await invoke("leave_document_session", { id: documentId });
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
    if (!editor || busy || !ready || cloudHandoffBlocked.current) return false;
    setBusy(true);
    try {
      if (currentSourcePath.current?.startsWith("remote:")) {
        await persistCurrent();
        cloudSync.current?.retry();
        setSaveStatus("Сохранение в облако запущено");
        return true;
      }
      const baseName = documentName.replace(/\.docx$/i, "");
      const suggestedName = `${baseName}${currentSourcePath.current && !currentSavePath.current ? " (DOCX Vim)" : ""}.docx`;
      if (isTauri()) {
        if (!currentSavePath.current) {
          const selected = await save({ defaultPath: suggestedName, filters: [{ name: "Документ Word", extensions: ["docx"] }] });
          if (!selected) return false;
          currentSavePath.current = selected;
          currentSourcePath.current = selected;
          currentRestoreSourcePath.current = undefined;
          currentSourceFingerprint.current = "new";
        }
        await persistCurrent();
        if (currentSourcePath.current === currentSavePath.current && currentSourcePath.current) {
          await rememberDocument({ id: documentId, name: documentName, sourcePath: currentSourcePath.current });
          setSourceLabel("Локальный");
        }
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

  const exportConflict = async (provider: CloudProvider): Promise<void> => {
    const copy = cloudStates[provider]?.conflictCopy;
    if (!copy) return;
    const destination = await save({ defaultPath: `${documentName.replace(/\.docx$/i, "")}-${provider}-конфликт.docx`, filters: [{ name: "Документ Word", extensions: ["docx"] }] });
    if (!destination) return;
    try { await writeFile(destination, await readFile(copy, { baseDir: BaseDirectory.AppData })); }
    catch { /* Recovery copy remains in AppData/cloud-conflicts. */ }
  };

  const quitDocument = async (): Promise<void> => {
    try { await persistCurrent(true); } catch { return; }
    if (isTauri() && /Android/i.test(navigator.userAgent) && getCurrentWindow().label.startsWith("doc-")) {
      await invoke("leave_document_session", { id: documentId });
      await invoke("close_task");
      return;
    }
    await showLibrary();
    if (isTauri() && getCurrentWindow().label !== "main") {
      await invoke("leave_document_session", { id: documentId });
      await getCurrentWindow().destroy();
    }
  };

  const writeQuitAll = async (): Promise<void> => {
    if (!isTauri()) { if (await saveDocument()) await quitDocument(); return; }
    const windows = (await WebviewWindow.getAll()).filter((item) => item.label.startsWith("doc-"));
    if (!windows.length) { if (await saveDocument()) await quitDocument(); return; }
    const ids = [...new Set(windows.map((item) => item.label.slice(4).split("--")[0]))];
    const owners = (await Promise.all(ids.map(async (id) => {
      const session = await invoke<SessionSnapshot | null>("get_document_session", { id });
      return windows.find((item) => item.label === session?.editorLabel);
    }))).filter((item): item is WebviewWindow => !!item);
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
    ? selectedIssue.kind === "none" || selectedIssue.kind === "pending" ? ["Закрыть"]
      : [...selectedIssue.suggestions, ...(selectedIssue.kind === "spelling" ? ["Добавить в словарь"] : []), "Игнорировать", "Закрыть"]
    : [];
  const sourceSaveComplete = !cloudHandoffBlocked.current && !!latestDraftUpdatedAt.current && (
    sourceLabel === "Яндекс Диск"
      ? cloudStates.yandexDisk?.status === "saved" && cloudStates.yandexDisk.path === currentSourcePath.current
        && cloudStates.yandexDisk.savedDraftUpdatedAt === latestDraftUpdatedAt.current
      : sourceLabel === "Локальный" && lastLocalSave?.path === currentSourcePath.current
        && lastLocalSave?.updatedAt === latestDraftUpdatedAt.current
  );

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
      disabled={!editor || !editor.isEditable}
    >
      {label}
    </button>
  );

  return (
    <main className="app-shell">
      {!hideMenu && <header className="app-header">
        <div>
          <h1>{documentName}</h1><small>{sourceLabel}</small>
        </div>
        <div className="file-actions" role="group" aria-label="Файл">
          <button type="button" onClick={() => void showLibrary()}>Документы</button>
          <button type="button" onClick={() => void showYandexLibrary()}>Яндекс Диск</button>
          <button type="button" onClick={saveDocument} disabled={!editor || busy || !ready || cloudHandoffBlocked.current || (isTauri() && !hasEditLease.current)}>Сохранить DOCX</button>
          {cloudProviders.filter((provider) => cloudStates[provider]?.status === "conflict").map((provider) =>
            <button key={provider} type="button" onClick={() => void exportConflict(provider)}>Сохранить конфликтную копию {provider === "yandexDisk" ? "Яндекса" : "OneDrive"}</button>)}
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
          disabled={!editor || !editor.isEditable}
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
          } else if (["Tab", "Enter", " "].includes(event.key) && !event.ctrlKey && !event.altKey && !event.metaKey) {
            event.preventDefault(); event.stopPropagation();
            if (!correctionPendingRef.current || selectedIssue.kind === "pending") acceptCorrection();
          } else if (event.key === "Escape") {
            event.preventDefault(); event.stopPropagation();
            applyCorrectionChoice(correctionChoices.length - 1);
          }
        }}>
        <span>{selectedIssue.message || `Слово «${selectedIssue.word}» не найдено в словаре`}</span>
        {correctionPending && <span>Подбираем варианты…</span>}
        {!correctionPending && selectedIssue.kind !== "none" && !selectedIssue.suggestions.length && <span>Вариантов нет</span>}
        {selectedIssue.kind !== "none" && selectedIssue.kind !== "pending" && <span>Выбрано: {correctionChoices[correctionIndex]} · Ctrl+N/P — выбор · Tab/Enter/Пробел — выполнить</span>}
        {correctionChoices.map((choice, index) => <button type="button" key={`${choice}-${index}`}
          className={index === correctionIndex ? "correction-choice--active" : ""} aria-pressed={index === correctionIndex}
          onClick={() => applyCorrectionChoice(index)}>{choice}</button>)}
      </div>}

      <footer className="status-bar">
        <span>{isTauri() && !hasEditLease.current ? "Просмотр — редактирование в другом окне"
          : cloudHandoffBlocked.current ? "Ожидаем сохранения из предыдущего окна; изменения в резервном черновике"
          : vimEnabled ? `Vim ${vimMode}` : "Обычный ввод"}</span>
        {completion && <span className="completion-hint">{completion.suggestions[completion.index]} · {completion.index + 1}/{completion.suggestions.length} · Ctrl+N/P — выбор · Tab — вставить</span>}
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
        {hideMenu && <><button type="button" onClick={() => setShowHeadings((value) => !value)}>Заголовки</button><button type="button" onClick={() => void showYandexLibrary()}>Яндекс Диск</button></>}
        <span className="source-indicator" aria-label={sourceLabel === "Яндекс Диск" ? "Файл в облаке" : "Файл не в облаке"}>
          {sourceLabel === "Яндекс Диск" ? "✓ Облако" : `✕ ${sourceLabel}`}
          {sourceSaveComplete && <span className="source-save-check" role="img" aria-label="Текущие изменения сохранены в выбранный файл" title="Текущие изменения сохранены в выбранный файл">✓</span>}
          {sourceLabel === "Яндекс Диск" && cloudStates.yandexDisk?.status === "retry" ? " · ожидает синхронизации" : ""}
          {sourceLabel === "Яндекс Диск" && cloudStates.yandexDisk?.status === "conflict" ? " · конфликт версии" : ""}
        </span>
        <span>{saveStatus}</span>
      </footer>
    </main>
  );
}

export default DocumentEditor;
