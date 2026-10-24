import { fetch as tauriFetch } from "@tauri-apps/plugin-http";
import { accessToken } from "./cloud-auth";
import type { CloudProvider } from "./cloud-sync";
import { readSettings } from "./settings";

export class RemoteConflict extends Error {}
type Meta = { revision: string; exists: boolean };
export type RemoteFile = { name: string; path: string; modified?: string };
const diskBase = "https://cloud-api.yandex.net/v1/disk/resources";

export function remoteTarget(provider: CloudProvider, id: string, title: string): string {
  const name = title.replace(/\.docx$/i, "").replace(/[\\/:*?"<>|\x00-\x1f]/g, " ").trim().slice(0, 90) || "Документ";
  return `remote:${provider}:${encodeURIComponent(`DOCX Vim - ${name} - ${id.slice(0, 8)}.docx`)}`;
}

export function isRemoteTarget(path: string): boolean { return /^remote:(oneDrive|yandexDisk):/.test(path); }

export function yandexTargetFromPath(path: string): string {
  if (!path.startsWith("disk:/") || !path.toLowerCase().endsWith(".docx") || path.split("/").some((part) => part === ".."))
    throw new Error("Invalid Yandex Disk path");
  const relative = path.slice("disk:/".length);
  return `remote:yandexDisk:${encodeURIComponent(relative.includes("/") ? path : relative)}`;
}

function targetName(provider: CloudProvider, target: string): string {
  const prefix = `remote:${provider}:`;
  if (!target.startsWith(prefix)) throw new Error("Invalid remote target");
  const name = decodeURIComponent(target.slice(prefix.length));
  if (!name || name.includes("\\") || !name.toLowerCase().endsWith(".docx") ||
    (provider === "oneDrive" && name.includes("/")) ||
    (provider === "yandexDisk" && name.includes("/") && (!name.startsWith("disk:/") || name.split("/").some((part) => part === ".."))))
    throw new Error("Invalid remote name");
  return name;
}

const diskPath = (name: string) => name.startsWith("disk:/") ? name : `disk:/${name}`;

async function credentials(provider: CloudProvider): Promise<{ token: string }> {
  const settings = await readSettings();
  const clientId = provider === "oneDrive" ? settings.oneDriveClientId : settings.yandexClientId;
  return { token: await accessToken(provider, clientId) };
}

export async function verifyRemoteRevision(provider: CloudProvider, target: string, revision: string): Promise<void> {
  const name = targetName(provider, target);
  const { token } = await credentials(provider);
  const meta = provider === "oneDrive" ? await oneDriveMeta(name, token) : await yandexMeta(name, token);
  if (!meta.exists || meta.revision !== revision) throw new RemoteConflict("Remote document changed");
}

async function oneDriveMeta(name: string, token: string): Promise<Meta> {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(name)}`;
  const response = await tauriFetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (response.status === 404) return { revision: "", exists: false };
  if (!response.ok) throw new Error(`OneDrive metadata ${response.status}`);
  const data = await response.json() as { eTag?: string; file?: object };
  if (!data.file || !data.eTag) throw new Error("OneDrive metadata missing");
  return { revision: data.eTag, exists: true };
}

async function yandexMeta(name: string, token: string): Promise<Meta> {
  const url = `${diskBase}?${new URLSearchParams({ path: diskPath(name), fields: "md5,modified,type" })}`;
  const response = await tauriFetch(url, { headers: { Authorization: `OAuth ${token}` } });
  if (response.status === 404) return { revision: "", exists: false };
  if (!response.ok) throw new Error(`Yandex metadata ${response.status}`);
  const data = await response.json() as { md5?: string; modified?: string; type?: string };
  if (data.type !== "file" || (!data.md5 && !data.modified)) throw new Error("Yandex metadata missing");
  return { revision: data.md5 ? `md5:${data.md5}` : `modified:${data.modified}`, exists: true };
}

export async function listYandexDocx(): Promise<RemoteFile[]> {
  const { token } = await credentials("yandexDisk");
  const files: RemoteFile[] = [];
  for (let offset = 0; offset < 5000; offset += 200) {
    const url = `${diskBase}/files?${new URLSearchParams({ limit: "200", offset: String(offset), fields: "items.name,items.path,items.modified" })}`;
    const response = await tauriFetch(url, { headers: { Authorization: `OAuth ${token}` } });
    if (!response.ok) throw new Error(`Yandex file list ${response.status}`);
    const data = await response.json() as { items?: { name?: string; path?: string; modified?: string }[] };
    const batch = data.items ?? [];
    for (const item of batch) if (item.name?.toLowerCase().endsWith(".docx") && item.path?.startsWith("disk:/")) {
      files.push({ name: item.name, path: yandexTargetFromPath(item.path), modified: item.modified });
    }
    if (batch.length < 200) break;
  }
  return files.sort((a, b) => (b.modified ?? "").localeCompare(a.modified ?? ""));
}

export async function downloadYandexDocx(target: string): Promise<{ bytes: Uint8Array; revision: string }> {
  const name = targetName("yandexDisk", target);
  const { token } = await credentials("yandexDisk");
  const before = await yandexMeta(name, token);
  if (!before.exists) throw new Error("Yandex document not found");
  const url = `${diskBase}/download?${new URLSearchParams({ path: diskPath(name) })}`;
  const linkResponse = await tauriFetch(url, { headers: { Authorization: `OAuth ${token}` } });
  if (!linkResponse.ok) throw new Error(`Yandex download link ${linkResponse.status}`);
  const link = await linkResponse.json() as { href?: string };
  if (!link.href?.startsWith("https://")) throw new Error("Invalid Yandex download link");
  const response = await tauriFetch(link.href);
  if (!response.ok) throw new Error(`Yandex download ${response.status}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  const after = await yandexMeta(name, token);
  if (!after.exists || before.revision !== after.revision) throw new RemoteConflict("Yandex document changed during download");
  return { bytes, revision: after.revision };
}

export async function uploadRemote(provider: CloudProvider, target: string, bytes: Uint8Array, knownRevision?: string): Promise<string> {
  const name = targetName(provider, target);
  const { token } = await credentials(provider);
  const meta = provider === "oneDrive" ? await oneDriveMeta(name, token) : await yandexMeta(name, token);
  if (knownRevision ? !meta.exists || meta.revision !== knownRevision : meta.exists) throw new RemoteConflict("Remote document changed");
  if (provider === "oneDrive") {
    if (bytes.byteLength > 250_000_000) throw new Error("OneDrive small upload limit");
    const response = await tauriFetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(name)}:/content`, {
      method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ...(knownRevision ? { "If-Match": knownRevision } : {}) }, body: bytes.slice().buffer as ArrayBuffer,
    });
    if (response.status === 409 || response.status === 412) throw new RemoteConflict("OneDrive version changed");
    if (!response.ok) throw new Error(`OneDrive upload ${response.status}`);
    const result = await response.json() as { eTag?: string };
    if (!result.eTag) throw new Error("OneDrive revision missing");
    return result.eTag;
  }
  const uploadUrl = `${diskBase}/upload?${new URLSearchParams({ path: diskPath(name), overwrite: String(!!knownRevision) })}`;
  const linkResponse = await tauriFetch(uploadUrl, { headers: { Authorization: `OAuth ${token}` } });
  if (linkResponse.status === 409 || linkResponse.status === 412) throw new RemoteConflict("Yandex version changed");
  if (!linkResponse.ok) throw new Error(`Yandex upload link ${linkResponse.status}`);
  const link = await linkResponse.json() as { href?: string; method?: string };
  if (!link.href || !link.href.startsWith("https://") || link.method !== "PUT") throw new Error("Invalid Yandex upload link");
  const uploaded = await tauriFetch(link.href, { method: "PUT", body: bytes.slice().buffer as ArrayBuffer,
    headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
  if (!uploaded.ok) throw new Error(`Yandex upload ${uploaded.status}`);
  for (let attempt = 0; attempt < 8; attempt++) {
    const updated = await yandexMeta(name, token);
    if (updated.exists && (updated.revision !== knownRevision || uploaded.status === 201)) return updated.revision;
    await new Promise<void>((resolve) => window.setTimeout(resolve, 1000));
  }
  throw new Error("Yandex upload awaiting processing");
}
