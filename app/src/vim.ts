import type { Editor } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";

export type VimMode = "NORMAL" | "INSERT" | "VISUAL";
const ru = "йцукенгшщзхъфывапролджэячсмитьбю";
const en = "qwertyuiop[]asdfghjkl;'zxcvbnm,.";
const words = /[\p{L}\p{N}_]+/gu;

export class VimController {
  mode: VimMode = "NORMAL";
  private pending = "";
  private count = "";
  private anchor = 0;
  private register = "";
  private search = "";
  constructor(private editor: Editor, private changed: (mode: VimMode) => void,
    private windowCommand: (key: string) => void, private searchInput: () => void) {}

  reset() { this.pending = ""; this.count = ""; this.mode = "NORMAL"; this.changed(this.mode); }
  private modeTo(mode: VimMode) { this.mode = mode; this.pending = ""; this.count = ""; this.changed(mode); }
  private set(pos: number, visual = this.mode === "VISUAL") {
    const { state, view } = this.editor;
    const next = Math.max(1, Math.min(pos, state.doc.content.size - 1));
    const selection = TextSelection.create(state.doc, visual ? this.anchor : next, next);
    view.dispatch(state.tr.setSelection(selection).scrollIntoView());
  }
  private move(key: string, count: number): boolean {
    const { state, view } = this.editor;
    const pos = state.selection.head;
    const $pos = state.doc.resolve(pos);
    const start = $pos.start(); const end = $pos.end();
    if (key === "h") this.set(Math.max(start, pos - count));
    else if (key === "l") this.set(Math.min(end, pos + count));
    else if (key === "0") this.set(start);
    else if (key === "$" ) this.set(end);
    else if (key === "j" || key === "k") {
      for (let i = 0; i < count; i++) {
        const at = view.coordsAtPos(view.state.selection.head);
        const hit = view.posAtCoords({ left: at.left, top: at.top + (key === "j" ? 1 : -1) * Math.max(18, at.bottom - at.top + 4) });
        if (hit) this.set(hit.pos);
      }
    } else if (key === "G") this.set(state.doc.content.size - 1);
    else if (key === "gg") this.set(1);
    else if (["w", "b", "e"].includes(key)) {
      const matches: { start: number; end: number }[] = [];
      state.doc.descendants((node, offset) => {
        if (node.isText) for (const match of node.text!.matchAll(words)) {
          matches.push({ start: offset + match.index, end: offset + match.index + match[0].length - 1 });
        }
      });
      let target = pos;
      for (let i = 0; i < count; i++) {
        const found = key === "b" ? matches.filter((m) => m.start < target).slice(-1)[0]
          : matches.find((m) => (key === "e" ? m.end : m.start) > target);
        if (found) target = key === "e" ? found.end : found.start;
      }
      this.set(target);
    } else return false;
    return true;
  }
  handle(event: KeyboardEvent): boolean {
    if (event.isComposing || event.key === "Process") return false;
    if (event.key === "Escape") { event.preventDefault(); this.modeTo("NORMAL"); this.editor.commands.focus(); return true; }
    if (this.mode === "INSERT") return false;
    if (event.ctrlKey && !event.altKey && (event.code === "KeyW" || event.key.toLowerCase() === "w" || event.key.toLowerCase() === "ц")) { event.preventDefault(); this.pending = "window"; return true; }
    if (event.ctrlKey || event.altKey || event.metaKey) return false;
    const physical = event.code.startsWith("Key") ? event.code.slice(3).toLowerCase() : "";
    const translated = ru.indexOf(event.key.toLowerCase());
    const fallback = translated >= 0 ? en[translated] : event.key;
    const key = event.code === "Semicolon" && event.shiftKey ? ":"
      : event.code === "Slash" && !event.shiftKey ? "/"
      : event.code === "Digit4" && event.shiftKey ? "$"
      : physical ? (event.shiftKey ? physical.toUpperCase() : physical)
      : event.shiftKey && fallback === ";" ? ":"
      : event.shiftKey && /[a-z]/.test(fallback) ? fallback.toUpperCase()
      : fallback;
    if (this.pending === "window") { this.pending = ""; this.windowCommand(key); event.preventDefault(); return true; }
    if (key.length !== 1 && key !== "Escape") return false;
    event.preventDefault();
    if (this.pending === "g") { this.pending = ""; if (key === "g") this.move("gg", 1); return true; }
    if (this.pending === "d" || this.pending === "y") {
      const operation = this.pending; this.pending = "";
      if (key === operation) {
        const { $from } = this.editor.state.selection;
        const from = $from.before($from.depth); const to = $from.after($from.depth);
        this.register = this.editor.state.doc.textBetween(from, to, "\n") + "\n";
        if (operation === "d") this.editor.chain().focus().deleteRange({ from, to }).run();
      }
      return true;
    }
    if (/^[1-9]$/.test(key) || (this.count && key === "0")) { this.count += key; return true; }
    const count = Math.min(999, Number(this.count) || 1); this.count = "";
    if (this.move(key, count)) return true;
    if (key === "g") this.pending = "g";
    else if (key === "i" || key === "a") { if (key === "a") this.move("l", 1); this.modeTo("INSERT"); }
    else if (key === "o") { this.editor.chain().focus().setTextSelection(this.editor.state.selection.$from.end()).enter().run(); this.modeTo("INSERT"); }
    else if (key === "v") { this.anchor = this.editor.state.selection.head; this.modeTo("VISUAL"); }
    else if (key === "x") { const p = this.editor.state.selection.head; this.editor.chain().focus().deleteRange({ from: p, to: Math.min(p + count, this.editor.state.selection.$from.end()) }).run(); }
    else if (key === "d" || key === "y") {
      if (this.mode === "VISUAL") {
        const { from, to } = this.editor.state.selection;
        this.register = this.editor.state.doc.textBetween(from, to, "\n");
        if (key === "d") this.editor.chain().focus().deleteSelection().run();
        this.modeTo("NORMAL");
      } else this.pending = key;
    }
    else if (key === "p") this.editor.chain().focus().insertContent(this.register).run();
    else if (key === "u") this.editor.commands.undo();
    else if (key === ":") {
      const command = window.prompt("Команда Vim (:split, :vsplit)");
      if (command && /^(?:\s*:?\s*)(?:split|vsplit|sp|vs)\s*$/.test(command)) this.windowCommand("s");
    }
    else if (key === "/") this.searchInput();
    else if (key === "n" && this.search) this.findNext();
    return true;
  }
  find(query: string) { this.search = query; this.findNext(); }
  private findNext() {
    const hits: number[] = [];
    this.editor.state.doc.descendants((node, pos) => {
      if (!node.isText) return;
      const text = node.text!.toLocaleLowerCase();
      let from = 0; let found: number;
      while ((found = text.indexOf(this.search.toLocaleLowerCase(), from)) >= 0) {
        hits.push(pos + found); from = found + 1;
      }
    });
    const head = this.editor.state.selection.head;
    const next = hits.find((hit) => hit > head) ?? hits[0];
    if (next !== undefined) this.set(next);
  }
}
