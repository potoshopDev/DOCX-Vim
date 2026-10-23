import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

export type TextIssue = { from: number; to: number; word: string; suggestions: string[]; kind: "spelling" | "grammar"; message?: string };
export const textCheckKey = new PluginKey<DecorationSet>("text-check");

export const TextCheck = Extension.create({
  name: "textCheck",
  addProseMirrorPlugins() {
    return [new Plugin({
      key: textCheckKey,
      state: {
        init: () => DecorationSet.empty,
        apply(tr, previous) {
          const issues = tr.getMeta(textCheckKey) as TextIssue[] | undefined;
          if (issues) return DecorationSet.create(tr.doc, issues.filter((issue) => issue.from < issue.to && issue.to <= tr.doc.content.size)
            .map((issue) => Decoration.inline(issue.from, issue.to, { class: `text-issue text-issue--${issue.kind}` })));
          return tr.docChanged ? previous.map(tr.mapping, tr.doc) : previous;
        },
      },
      props: { decorations: (state) => textCheckKey.getState(state) },
    })];
  },
});
