import { supportsMultipleWindows } from "@tauri-apps/api/app";
import { isTauri } from "@tauri-apps/api/core";
import { invoke } from "@tauri-apps/api/core";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import { getCurrentWindow } from "@tauri-apps/api/window";

const labelFor = (id: string) => `doc-${id}`;
export const routeDocumentId = () => new URLSearchParams(location.search).get("doc");

export async function openDocumentWindow(id: string, name: string, duplicate = false): Promise<void> {
  const android = isTauri() && /Android/i.test(navigator.userAgent);
  if (!isTauri() || (!android && !(await supportsMultipleWindows()))) {
    history.pushState(null, "", `?doc=${encodeURIComponent(id)}`);
    dispatchEvent(new PopStateEvent("popstate"));
    return;
  }
  const label = duplicate ? `${labelFor(id)}--${crypto.randomUUID().slice(0, 8)}` : labelFor(id);
  const existing = duplicate ? null : (await WebviewWindow.getByLabel(label))
    ?? (await WebviewWindow.getAll()).find((item) => item.label.startsWith(`${label}--`)) ?? null;
  if (existing) {
    if (android) await invoke("focus_task", { id: existing.label.slice(4) });
    else { await existing.show(); await existing.setFocus(); }
    return;
  }
  try {
    const window = new WebviewWindow(label, {
      url: `index.html?doc=${encodeURIComponent(id)}`,
      title: name,
      width: 940,
      height: 720,
      ...(android ? { activityName: "DocumentActivity" } : {}),
    });
    await new Promise<void>((resolve, reject) => {
      void window.once("tauri://created", () => resolve());
      void window.once("tauri://error", (event) => reject(new Error(String(event.payload))));
    });
  } catch (error) {
    if (!android) throw error;
    history.pushState(null, "", `?doc=${encodeURIComponent(id)}`);
    dispatchEvent(new PopStateEvent("popstate"));
  }
}

export async function focusAdjacentWindow(direction: number): Promise<void> {
  if (!isTauri()) return;
  const windows = (await WebviewWindow.getAll()).filter((item) => item.label.startsWith("doc-"));
  if (windows.length < 2) return;
  const current = getCurrentWindow().label;
  const index = windows.findIndex((item) => item.label === current);
  const target = windows[(index + direction + windows.length) % windows.length];
  if (/Android/i.test(navigator.userAgent)) await invoke("focus_task", { id: target.label.slice(4) });
  else { await target.show(); await target.setFocus(); }
}

export async function isDocumentSaveOwner(id: string): Promise<boolean> {
  if (!isTauri()) return true;
  const labels = (await WebviewWindow.getAll()).map((item) => item.label).filter((label) => label === labelFor(id) || label.startsWith(`${labelFor(id)}--`));
  labels.sort((a, b) => a.length - b.length || a.localeCompare(b));
  return labels[0] === getCurrentWindow().label;
}

export async function showLibrary(): Promise<void> {
  const android = isTauri() && /Android/i.test(navigator.userAgent);
  if (!isTauri() || (!android && !(await supportsMultipleWindows())) || (android && getCurrentWindow().label === "main")) {
    history.pushState(null, "", location.pathname);
    dispatchEvent(new PopStateEvent("popstate"));
    return;
  }
  if (android) { await invoke("focus_task", { id: "main" }); return; }
  const main = await WebviewWindow.getByLabel("main");
  if (main) { await main.show(); await main.setFocus(); }
}

export async function registerDocumentTask(id: string, name: string): Promise<void> {
  if (isTauri() && /Android/i.test(navigator.userAgent)) await invoke("register_task", { id, name });
}

export async function openWindowIds(): Promise<string[]> {
  if (!isTauri()) return routeDocumentId() ? [routeDocumentId()!] : [];
  return (await WebviewWindow.getAll()).map((window) => window.label)
    .filter((label) => label.startsWith("doc-")).map((label) => label.slice(4).split("--")[0]);
}
