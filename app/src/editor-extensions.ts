import StarterKit from "@tiptap/starter-kit";
import TextAlign from "@tiptap/extension-text-align";

export const editorExtensions = [
  StarterKit.configure({ link: { openOnClick: false } }),
  TextAlign.configure({ types: ["heading", "paragraph"] }),
];
