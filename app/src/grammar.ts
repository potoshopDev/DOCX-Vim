import { isTauri } from "@tauri-apps/api/core";
import { fetch as tauriFetch } from "@tauri-apps/plugin-http";
import type { Language } from "./personal-dictionary";
import type { TextIssue } from "./text-check-extension";

type Match = { offset: number; length: number; message?: string; replacements?: { value: string }[]; rule?: { issueType?: string } };
type RelativeIssue = Omit<TextIssue, "from" | "to"> & { offset: number; length: number };
const cache = new Map<string, RelativeIssue[]>();

export function grammarEndpoint(base: string): string | null {
  try {
    const url = new URL(base);
    if (url.protocol !== "https:" && !(url.protocol === "http:" && ["localhost", "127.0.0.1", "10.0.2.2"].includes(url.hostname))) return null;
    if (url.hostname === "api.languagetool.org") return null;
    url.pathname = `${url.pathname.replace(/\/(?:v2\/check)?\/?$/, "")}/v2/check`;
    return url.toString();
  } catch { return null; }
}

export async function checkGrammar(segments: { from: number; text: string }[], language: Language, base: string, signal: AbortSignal): Promise<TextIssue[]> {
  const endpoint = grammarEndpoint(base);
  if (!endpoint) return [];
  const issues: TextIssue[] = [];
  for (const segment of segments.filter((item) => item.text.trim().length >= 8 && item.text.length < 10000)) {
    if (signal.aborted) return [];
    const key = `${endpoint}|${language}|${segment.text}`;
    let found = cache.get(key);
    if (!found) {
      const body = new URLSearchParams({ text: segment.text, language: language === "ru" ? "ru-RU" : "en-US" });
      const response = await (isTauri() ? tauriFetch : fetch)(endpoint, { method: "POST", body, signal });
      if (!response.ok) continue;
      const data = await response.json() as { matches?: Match[] };
      found = (data.matches ?? []).filter((match) => match.rule?.issueType !== "misspelling" && match.length > 0)
        .map((match) => ({ offset: match.offset, length: match.length, kind: "grammar" as const,
          word: segment.text.slice(match.offset, match.offset + match.length), message: match.message,
          suggestions: (match.replacements ?? []).slice(0, 4).map((item) => item.value) }));
      cache.set(key, found);
      if (cache.size > 200) cache.delete(cache.keys().next().value!);
    }
    for (const issue of found) issues.push({ ...issue, from: segment.from + issue.offset, to: segment.from + issue.offset + issue.length });
  }
  return issues;
}
