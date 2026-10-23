import type { JSONContent } from "@tiptap/core";
import mammoth from "mammoth";
import DOMPurify from "dompurify";
import JSZip from "jszip";
import {
  AlignmentType, Document, ExternalHyperlink, HeadingLevel,
  LevelFormat, Packer, Paragraph, TextRun,
} from "docx";

export type ImportedDocx = { html: string; warnings: string[] };

export async function importDocx(bytes: Uint8Array): Promise<ImportedDocx> {
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (!xml) throw new Error("Файл не содержит документа Word.");
  const warnings: string[] = [];
  const checks: [RegExp, string][] = [
    [/<w:tbl(?:\s|>)/, "Таблицы пока не поддерживаются и могут быть потеряны."],
    [/<(?:w:drawing|w:pict)(?:\s|>)/, "Изображения пока не сохраняются."],
    [/<w:(?:ins|del)(?:\s|>)/, "Исправления с отслеживанием изменений не сохраняются."],
    [/<w:br[^>]*w:type="page"/, "Разрывы страниц не сохраняются."],
    [/<w:commentRangeStart(?:\s|>)/, "Комментарии не сохраняются."],
  ];
  for (const [pattern, warning] of checks) if (pattern.test(xml)) warnings.push(warning);
  if (Object.keys(zip.files).some((name) => /^word\/(?:header|footer|footnotes|endnotes)/.test(name))) {
    warnings.push("Колонтитулы и сноски не сохраняются.");
  }
  const result = await mammoth.convertToHtml({ arrayBuffer: buffer });
  for (const message of result.messages) {
    if (message.type === "warning") warnings.push(message.message);
  }
  // Imported document content is untrusted HTML.
  const html = DOMPurify.sanitize(result.value, {
    ALLOWED_TAGS: ["p", "br", "strong", "b", "em", "i", "u", "s", "strike", "a", "h1", "h2", "h3", "ul", "ol", "li", "blockquote", "hr"],
    ALLOWED_ATTR: ["href"],
  });
  return { html, warnings: [...new Set(warnings)] };
}

function inlineRuns(nodes: JSONContent[] = []): (TextRun | ExternalHyperlink)[] {
  const runs: (TextRun | ExternalHyperlink)[] = [];
  for (const node of nodes) {
    if (node.type === "hardBreak") { runs.push(new TextRun({ break: 1 })); continue; }
    if (node.type !== "text") continue;
    const marks = node.marks ?? [];
    const run = new TextRun({
      text: node.text ?? "",
      bold: marks.some((mark) => mark.type === "bold"),
      italics: marks.some((mark) => mark.type === "italic"),
      underline: marks.some((mark) => mark.type === "underline") ? {} : undefined,
      strike: marks.some((mark) => mark.type === "strike"),
    });
    const link = marks.find((mark) => mark.type === "link")?.attrs?.href;
    runs.push(typeof link === "string" && /^https?:\/\//i.test(link)
      ? new ExternalHyperlink({ link, children: [run] })
      : run);
  }
  return runs;
}

export async function exportDocx(json: JSONContent): Promise<Blob> {
  const paragraphs: Paragraph[] = [];
  const add = (node: JSONContent, list?: "bullet" | "ordered", depth = 0) => {
    if (node.type === "bulletList" || node.type === "orderedList") {
      node.content?.forEach((child) => add(child, node.type === "bulletList" ? "bullet" : "ordered", depth));
      return;
    }
    if (node.type === "listItem") {
      node.content?.forEach((child) => add(child, list, depth + 1));
      return;
    }
    if (node.type === "blockquote") {
      node.content?.forEach((child) => add(child, list, depth));
      return;
    }
    if (node.type !== "paragraph" && node.type !== "heading") return;
    const level = Number(node.attrs?.level ?? 1);
    const alignment = node.attrs?.textAlign;
    paragraphs.push(new Paragraph({
      children: inlineRuns(node.content),
      heading: node.type === "heading" ? [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3][level - 1] : undefined,
      alignment: alignment === "center" ? AlignmentType.CENTER : alignment === "right" ? AlignmentType.RIGHT : undefined,
      bullet: list === "bullet" ? { level: Math.min(depth - 1, 8) } : undefined,
      numbering: list === "ordered" ? { reference: "ordered", level: Math.min(depth - 1, 8) } : undefined,
    }));
  };
  json.content?.forEach((node) => add(node));
  const doc = new Document({
    numbering: { config: [{ reference: "ordered", levels: Array.from({ length: 9 }, (_, level) => ({ level, format: LevelFormat.DECIMAL, text: `%${level + 1}.` })) }] },
    sections: [{ children: paragraphs.length ? paragraphs : [new Paragraph("")] }],
  });
  return Packer.toBlob(doc);
}
