import { supportsMultipleWindows } from "@tauri-apps/api/app";
import { isTauri } from "@tauri-apps/api/core";
import { invoke } from "@tauri-apps/api/core";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";

const labelFor = (id: string) => `doc-${id}`;
export const routeDocumentId = () => new URLSearchParams(location.search).get("doc");

export async function openDocumentWindow(id: string, name: string): Promise<void> {
  if (!isTauri() || !(await supportsMultipleWindows())) {
    history.pushState(null, "", `?doc=${encodeURIComponent(id)}`);
    dispatchEvent(new PopStateEvent("popstate"));
    return;
  }
  const label = labelFor(id);
  const existing = await WebviewWindow.getByLabel(label);
  const android = /Android/i.test(navigator.userAgent);
  if (existing) {
    if (android) await invoke("focus_task", { id });
    else { await existing.show(); await existing.setFocus(); }
    return;
  }
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
}

export async function showLibrary(): Promise<void> {
  if (!isTauri() || !(await supportsMultipleWindows())) {
    history.pushState(null, "", location.pathname);
    dispatchEvent(new PopStateEvent("popstate"));
    return;
  }
  if (/Android/i.test(navigator.userAgent)) { await invoke("focus_task", { id: "main" }); return; }
  const main = await WebviewWindow.getByLabel("main");
  if (main) { await main.show(); await main.setFocus(); }
}

export async function registerDocumentTask(id: string, name: string): Promise<void> {
  if (isTauri() && /Android/i.test(navigator.userAgent)) await invoke("register_task", { id, name });
}

export async function openWindowIds(): Promise<string[]> {
  if (!isTauri()) return routeDocumentId() ? [routeDocumentId()!] : [];
  return (await WebviewWindow.getAll()).map((window) => window.label)
    .filter((label) => label.startsWith("doc-")).map((label) => label.slice(4));
}
