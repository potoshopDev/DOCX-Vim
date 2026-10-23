import type { Editor, JSONContent } from "@tiptap/core";
import { vimPluginKey } from "vim-prosemirror";
import { TextSelection } from "@tiptap/pm/state";

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
  private backwardSearchInput = false;
  private ownClipboard: string | null = null;
  private ownContent: JSONContent[] | null = null;
  private lastWordAction: { operator: "d" | "c"; object: "w" | "W"; around: boolean } | null = null;
  enabled = true;
  constructor(private editor: Editor, private windowCommand: (key: string) => void,
    private commandInput: () => void) {}

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

  private searchBackwards(query: string): void {
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
      this.backwardSearch = true;
    }
    const head = state.selection.head;
    const previous = hits.filter((pos) => pos < head).pop() ?? hits[hits.length - 1];
    if (previous !== undefined) view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, previous)).scrollIntoView());
    else view.dispatch(state.tr);
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
    if (state.searchActive && this.backwardSearchInput && (event.key === "Enter" || event.key === "Escape")) {
      event.preventDefault();
      const query = state.searchQuery;
      this.backwardSearchInput = false;
      const prefix = this.editor.view.dom.parentElement?.querySelector(".vim-search-prefix");
      if (prefix) prefix.textContent = "/";
      if (event.key === "Enter" && query) this.searchBackwards(query);
      else { state.searchActive = false; state.searchQuery = ""; this.editor.view.dispatch(this.editor.state.tr); }
      return true;
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
    if (key === "?") {
      event.preventDefault();
      this.backwardSearchInput = true;
      state.searchActive = true;
      state.searchQuery = "";
      const prefix = this.editor.view.dom.parentElement?.querySelector(".vim-search-prefix");
      if (prefix) prefix.textContent = "?";
      this.editor.view.dispatch(this.editor.state.tr);
      return true;
    }
    if (key === "/") this.backwardSearch = false;
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
    withKey(event, key);
    return false;
  }
}
