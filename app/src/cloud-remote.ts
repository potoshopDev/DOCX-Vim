import { fetch as tauriFetch } from "@tauri-apps/plugin-http";
import { accessToken } from "./cloud-auth";
import type { CloudProvider } from "./cloud-sync";
import { readSettings } from "./settings";

export class RemoteConflict extends Error {}
type Meta = { revision: string; exists: boolean };
const diskBase = "https://cloud-api.yandex.net/v1/disk/resources";

export function remoteTarget(provider: CloudProvider, id: string, title: string): string {
  const name = title.replace(/\.docx$/i, "").replace(/[\\/:*?"<>|\x00-\x1f]/g, " ").trim().slice(0, 90) || "Документ";
  return `remote:${provider}:${encodeURIComponent(`DOCX Vim - ${name} - ${id.slice(0, 8)}.docx`)}`;
}

export function isRemoteTarget(path: string): boolean { return /^remote:(oneDrive|yandexDisk):/.test(path); }

function targetName(provider: CloudProvider, target: string): string {
  const prefix = `remote:${provider}:`;
  if (!target.startsWith(prefix)) throw new Error("Invalid remote target");
  const name = decodeURIComponent(target.slice(prefix.length));
  if (!name || name.includes("/") || name.includes("\\") || !name.toLowerCase().endsWith(".docx")) throw new Error("Invalid remote name");
  return name;
}

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
  const url = `${diskBase}?${new URLSearchParams({ path: `disk:/${name}`, fields: "md5,modified,type" })}`;
  const response = await tauriFetch(url, { headers: { Authorization: `OAuth ${token}` } });
  if (response.status === 404) return { revision: "", exists: false };
  if (!response.ok) throw new Error(`Yandex metadata ${response.status}`);
  const data = await response.json() as { md5?: string; modified?: string; type?: string };
  if (data.type !== "file" || (!data.md5 && !data.modified)) throw new Error("Yandex metadata missing");
  return { revision: data.md5 ? `md5:${data.md5}` : `modified:${data.modified}`, exists: true };
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
  const uploadUrl = `${diskBase}/upload?${new URLSearchParams({ path: `disk:/${name}`, overwrite: String(!!knownRevision) })}`;
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
