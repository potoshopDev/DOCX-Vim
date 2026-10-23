import type { Editor, JSONContent } from "@tiptap/core";
import { vimPluginKey } from "vim-prosemirror";
import { Selection, TextSelection } from "@tiptap/pm/state";

const russian = "йцукенгшщзхъфывапролджэячсмитьбю";
const english = "qwertyuiop[]asdfghjkl;'zxcvbnm,.";

export type VimMode = "normal" | "insert" | "visual" | "visual-line" | "replace";

function commandKey(event: KeyboardEvent): string {
  if (event.key === "?" || event.key === ":") return event.key;
  if (event.code.startsWith("Key")) {
    const letter = event.code.slice(3).toLowerCase();
    return event.shiftKey ? letter.toUpperCase() : letter;
  }
  if (event.code === "Semicolon" && event.shiftKey) return ":";
  if (event.code === "Digit4" && event.shiftKey) return "$";
  if (event.code === "Slash" && !event.shiftKey) return "/";
  const index = russian.indexOf(event.key.toLowerCase());
  const base = index < 0 ? event.key : english[index];
  if (event.shiftKey && base === ";") return ":";
  return event.shiftKey && /^[a-z]$/.test(base) ? base.toUpperCase() : base;
}

function withKey(event: KeyboardEvent, key: string): void {
  if (event.key !== key) Object.defineProperty(event, "key", { configurable: true, value: key });
}

export class VimAdapter {
  private windowPending = false;
  private backwardSearch = false;
  private preferredX: number | null = null;
  private ownClipboard: string | null = null;
  private ownContent: JSONContent[] | null = null;
  private lastWordAction: { operator: "d" | "c"; object: "w" | "W"; around: boolean } | null = null;
  enabled = true;
  constructor(private editor: Editor, private windowCommand: (key: string) => void,
    private commandInput: () => void, private searchInput: (backward: boolean) => void) {}

  mode(): VimMode {
    return (vimPluginKey.getState(this.editor.state)?.mode ?? "normal") as VimMode;
  }

  setEnabled(enabled: boolean): void {
    this.enabled = enabled;
    this.windowPending = false;
    const state = vimPluginKey.getState(this.editor.state);
    if (!state) return;
    state.mode = enabled ? "normal" : "insert";
    state.count = null;
    state.operator = null;
    state.searchActive = false;
    this.editor.view.dispatch(this.editor.state.tr);
  }

  private openLines(count: number, above: boolean): void {
    const view = this.editor.view;
    const state = view.state;
    const $head = state.selection.$head;
    const block = $head.depth ? $head.depth : 1;
    const at = above ? $head.before(block) : $head.after(block);
    const paragraph = state.schema.nodes.paragraph;
    if (!paragraph) return;
    const lines = Array.from({ length: Math.min(999, count) }, () => paragraph.create());
    const tr = state.tr.insert(at, lines);
    const first = above ? at + 1 : at + (count - 1) * paragraph.create().nodeSize + 1;
    tr.setSelection(TextSelection.create(tr.doc, Math.min(first, tr.doc.content.size - 1)));
    const vim = vimPluginKey.getState(state);
    if (vim) { vim.mode = "insert"; vim.count = null; }
    view.dispatch(tr.scrollIntoView());
  }

  submitSearch(query: string, backward: boolean): void {
    if (!query) return;
    const view = this.editor.view;
    const state = view.state;
    const hits: number[] = [];
    state.doc.descendants((node, pos) => {
      if (!node.isText) return;
      const text = node.text!.toLocaleLowerCase();
      const needle = query.toLocaleLowerCase();
      let from = 0; let found = -1;
      while ((found = text.indexOf(needle, from)) >= 0) { hits.push(pos + found); from = found + 1; }
    });
    const vim = vimPluginKey.getState(state);
    if (vim) {
      vim.searchActive = false;
      vim.searchQuery = "";
      vim.searchTerm = query;
      vim.searchWholeWord = false;
      vim.searchHighlightsVisible = true;
      this.backwardSearch = backward;
    }
    const head = state.selection.head;
    const target = backward ? hits.filter((pos) => pos < head).pop() ?? hits[hits.length - 1]
      : hits.find((pos) => pos > head) ?? hits[0];
    if (target !== undefined) view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, target)).scrollIntoView());
    else view.dispatch(state.tr);
  }

  private lineBounds(pos: number): { from: number; to: number } {
    const view = this.editor.view;
    const $pos = view.state.doc.resolve(pos);
    const start = $pos.start($pos.depth);
    const end = $pos.end($pos.depth);
    const top = view.coordsAtPos(pos).top;
    const sameRow = (at: number) => Math.abs(view.coordsAtPos(at).top - top) < 2;
    let low = start; let high = pos;
    while (low < high) {
      const mid = Math.floor((low + high) / 2);
      if (sameRow(mid)) high = mid; else low = mid + 1;
    }
    const from = low;
    low = pos; high = end;
    while (low < high) {
      const mid = Math.ceil((low + high) / 2);
      if (sameRow(mid)) low = mid; else high = mid - 1;
    }
    return { from, to: low };
  }

  private moveScreenRow(pos: number, direction: number): number {
    const view = this.editor.view;
    const rect = view.coordsAtPos(pos);
    const lineHeight = Number.parseFloat(getComputedStyle(view.dom).lineHeight) || rect.bottom - rect.top || 24;
    const x = this.preferredX ?? rect.left;
    this.preferredX = x;
    for (let step = 0; step < 24; step++) {
      const target = view.posAtCoords({ left: x, top: (direction > 0 ? rect.bottom + 1 : rect.top - 1) + direction * lineHeight * step / 4 });
      if (!target) continue;
      const $target = view.state.doc.resolve(target.pos);
      const cursor = $target.parent.isTextblock ? target.pos
        : (Selection.findFrom($target, direction, true) ?? Selection.findFrom($target, -direction, true))?.$from.pos;
      if (cursor === undefined) continue;
      const targetTop = view.coordsAtPos(cursor).top;
      if (direction * (targetTop - rect.top) > 2) return cursor;
    }
    return pos;
  }

  private screenSelection(anchor: number, head: number): void {
    const state = this.editor.state;
    const a = this.lineBounds(anchor);
    const h = this.lineBounds(head);
    const from = Math.min(a.from, h.from);
    const to = Math.max(a.to, h.to);
    this.editor.view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, from, to)).scrollIntoView());
  }

  private wordObject(object: "w" | "W", around: boolean, operator: "d" | "c" | "y", count: number): boolean {
    const view = this.editor.view;
    const state = view.state;
    const $head = state.selection.$head;
    if (!$head.parent.isTextblock) return false;
    const text = $head.parent.textContent;
    const offset = $head.pos - $head.start();
    const pattern = object === "w" ? /[\p{L}\p{N}_]+|[^\p{L}\p{N}_\s]+/gu : /\S+/gu;
    const tokens = [...text.matchAll(pattern)].map((match) => ({ from: match.index, to: match.index + match[0].length }));
    const index = tokens.findIndex((token) => token.from <= offset && offset < token.to || token.from > offset);
    if (index < 0) return false;
    const first = tokens[index];
    const last = tokens[Math.min(tokens.length - 1, index + count - 1)];
    let from = first.from; let to = last.to;
    if (around) {
      const trailing = /^\s+/u.exec(text.slice(to))?.[0].length ?? 0;
      if (trailing) to += trailing;
      else from -= /\s+$/u.exec(text.slice(0, from))?.[0].length ?? 0;
    }
    const absoluteFrom = $head.start() + from;
    const absoluteTo = $head.start() + to;
    this.ownClipboard = text.slice(from, to);
    this.ownContent = state.doc.slice(absoluteFrom, absoluteTo).content.toJSON() as JSONContent[];
    void navigator.clipboard?.writeText(this.ownClipboard).catch(() => {});
    if (operator !== "y") {
      const tr = state.tr.delete(absoluteFrom, absoluteTo);
      tr.setSelection(TextSelection.create(tr.doc, Math.min(absoluteFrom, tr.doc.content.size - 1)));
      view.dispatch(tr.scrollIntoView());
      this.lastWordAction = { operator, object, around };
    }
    const vim = vimPluginKey.getState(view.state);
    if (vim) {
      vim.operator = null; vim.count = null;
      (vim as typeof vim & { _textObjectType?: string })._textObjectType = undefined;
      if (operator === "c") vim.mode = "insert";
      view.dispatch(view.state.tr);
    }
    return true;
  }

  handle(event: KeyboardEvent): boolean {
    const state = vimPluginKey.getState(this.editor.state);
    if (!state) return false;
    if (!this.enabled) {
      if (event.key === "Escape") { event.preventDefault(); return true; }
      return false;
    }
    if (event.isComposing || state.searchActive || state.mode === "insert" || state.mode === "replace") return false;
    if (event.ctrlKey && !event.altKey && (event.code === "KeyW" || /[wц]/i.test(event.key))) {
      this.windowPending = true; event.preventDefault(); return true;
    }
    if (event.ctrlKey || event.altKey || event.metaKey) return false;
    const key = commandKey(event);
    if (this.windowPending) {
      this.windowPending = false;
      event.preventDefault();
      this.windowCommand(key);
      return true;
    }
    if (key === ":") {
      event.preventDefault();
      this.commandInput();
      return true;
    }
    if (key === "?" || key === "/") {
      event.preventDefault();
      this.searchInput(key === "?");
      return true;
    }
    if (key === "V" && state.operator === null) {
      event.preventDefault();
      if (state.mode === "visual-line") {
        state.mode = "normal"; state.visualAnchor = null; state.visualHead = null;
        this.editor.view.dispatch(this.editor.state.tr.setSelection(TextSelection.create(this.editor.state.doc, this.editor.state.selection.from)));
      } else {
        const pos = state.visualHead ?? this.editor.state.selection.head;
        state.mode = "visual-line"; state.visualAnchor = pos; state.visualHead = pos;
        this.screenSelection(pos, pos);
      }
      return true;
    }
    if ((key === "j" || key === "k") && state.operator === null) {
      event.preventDefault();
      const count = Math.min(state.count ?? 1, 999);
      let pos = state.visualHead ?? this.editor.state.selection.head;
      for (let index = 0; index < count; index++) pos = this.moveScreenRow(pos, key === "j" ? 1 : -1);
      state.count = null;
      if (state.mode === "visual-line" && state.visualAnchor !== null) {
        state.visualHead = pos;
        this.screenSelection(state.visualAnchor, pos);
      } else if (state.mode === "visual" && state.visualAnchor !== null) {
        state.visualHead = pos;
        const from = Math.min(state.visualAnchor, pos);
        const to = Math.min(this.editor.state.doc.content.size, Math.max(state.visualAnchor, pos) + 1);
        this.editor.view.dispatch(this.editor.state.tr.setSelection(TextSelection.create(this.editor.state.doc, from, to)).scrollIntoView());
      } else this.editor.view.dispatch(this.editor.state.tr.setSelection(TextSelection.create(this.editor.state.doc, pos)).scrollIntoView());
      return true;
    }
    if (key !== "j" && key !== "k") this.preferredX = null;
    if (this.backwardSearch && (key === "n" || key === "N")) {
      withKey(event, key === "n" ? "N" : "n");
      return false;
    }
    if ((key === "o" || key === "O") && (state.count ?? 1) > 1 && state.operator === null) {
      event.preventDefault(); this.openLines(state.count!, key === "O"); return true;
    }
    const pending = (state as typeof state & { _textObjectType?: string })._textObjectType;
    if (state.operator && (pending === "a" || pending === "i") && (key === "w" || key === "W")) {
      event.preventDefault();
      this.wordObject(key, pending === "a", state.operator, state.count ?? 1);
      return true;
    }
    if (key === "." && this.lastWordAction && this.ownClipboard !== null) {
      event.preventDefault();
      const action = this.lastWordAction;
      this.wordObject(action.object, action.around, action.operator, 1);
      return true;
    }
    if ((key === "p" || key === "P") && this.ownClipboard !== null) {
      event.preventDefault();
      const pos = this.editor.state.selection.head + (key === "p" ? 1 : 0);
      const content = this.ownContent?.length ? Array.from({ length: state.count ?? 1 }, () => this.ownContent!).flat()
        : this.ownClipboard.repeat(state.count ?? 1);
      this.editor.chain().focus().insertContentAt(Math.min(pos, this.editor.state.doc.content.size - 1), content).run();
      state.count = null;
      return true;
    }
    if (key === "d" || key === "c" || key === "y") { this.ownClipboard = null; this.ownContent = null; }
    if (state.mode === "visual-line" && state.visualAnchor !== null) {
      queueMicrotask(() => {
        const current = vimPluginKey.getState(this.editor.state);
        if (current?.mode === "visual-line" && current.visualAnchor !== null && current.visualHead !== null)
          this.screenSelection(current.visualAnchor, current.visualHead);
      });
    }
    withKey(event, key);
    return false;
  }
}
