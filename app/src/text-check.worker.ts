import nspell from "nspell";
import type { Language, PersonalDictionary } from "./personal-dictionary";

type Segment = { from: number; text: string };
type Issue = { from: number; to: number; word: string; suggestions: string[]; kind: "spelling" };
type Request = { id: number; type: "preload" } | { id: number; type: "check"; segments: Segment[]; personal: PersonalDictionary } | { id: number; type: "complete"; prefix: string; language: Language; personal: PersonalDictionary } | { id: number; type: "suggest"; word: string; language: Language; personal: PersonalDictionary };
const spelling = new Map<Language, ReturnType<typeof nspell>>();
const roots = new Map<Language, string[]>();
const loading = new Map<Language, Promise<ReturnType<typeof nspell>>>();

async function load(language: Language) {
  if (spelling.has(language)) return spelling.get(language)!;
  let promise = loading.get(language);
  if (!promise) {
    promise = (async () => {
      const [aff, dic] = await Promise.all(["aff", "dic"].map(async (ext) => {
        const response = await fetch(`/dictionaries/${language}/index.${ext}`);
        if (!response.ok) throw new Error(`Dictionary ${language} unavailable`);
        return response.text();
      }));
      const checker = nspell({ aff, dic });
      spelling.set(language, checker);
      roots.set(language, dic.split(/\r?\n/).slice(1).map((line) => line.split(/[\/\t ]/, 1)[0].toLocaleLowerCase(language)).filter((word) => word.length >= 3).sort());
      return checker;
    })();
    loading.set(language, promise);
  }
  try { return await promise; }
  finally { loading.delete(language); }
}

self.onmessage = async (event: MessageEvent<Request>) => {
  const request = event.data;
  try {
    if (request.type === "preload") {
      await Promise.all([load("ru"), load("en")]);
      self.postMessage({ id: request.id, type: "preload" });
    } else if (request.type === "check") {
      const [ru, en] = await Promise.all([load("ru"), load("en")]);
      const issues: Issue[] = [];
      for (const segment of request.segments) {
        for (const match of segment.text.matchAll(/[\p{L}]+(?:[’'-][\p{L}]+)*/gu)) {
          const word = match[0];
          if (word.length < 2 || word.length > 64) continue;
          const language: Language = /[А-Яа-яЁё]/.test(word) ? "ru" : "en";
          if (!new RegExp(language === "ru" ? "^[А-Яа-яЁё’'-]+$" : "^[A-Za-z’'-]+$").test(word)) continue;
          const lower = word.toLocaleLowerCase(language);
          if (request.personal[language].includes(lower) || (language === "ru" ? ru : en).correct(word)) continue;
          issues.push({ from: segment.from + match.index!, to: segment.from + match.index! + word.length,
            word, suggestions: [], kind: "spelling" });
        }
      }
      self.postMessage({ id: request.id, type: "check", issues });
    } else if (request.type === "suggest") {
      const checker = await load(request.language);
      const correct = checker.correct(request.word) || request.personal[request.language].includes(request.word.toLocaleLowerCase(request.language));
      self.postMessage({ id: request.id, type: "suggest", suggestions: correct ? [] : checker.suggest(request.word).slice(0, 6), correct });
    } else {
      const checker = await load(request.language);
      const prefix = request.prefix.toLocaleLowerCase(request.language);
      const words = roots.get(request.language)!;
      let lo = 0, hi = words.length;
      while (lo < hi) { const mid = (lo + hi) >>> 1; if (words[mid] < prefix) lo = mid + 1; else hi = mid; }
      const candidates = [...new Set([...request.personal[request.language], ...words.slice(lo, lo + 300)])]
        .filter((word) => word.length > prefix.length && word.startsWith(prefix) && checker.correct(word));
      const suggestions = candidates.sort((a, b) => a.length - b.length || a.localeCompare(b)).slice(0, 25);
      if (!suggestions.length) suggestions.push(...checker.suggest(prefix).filter((word) => word.length > prefix.length).slice(0, 6));
      self.postMessage({ id: request.id, type: "complete", suggestions });
    }
  } catch { self.postMessage({ id: request.id, type: request.type, issues: [], suggestions: [] }); }
};
