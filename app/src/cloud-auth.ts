import { fetch as tauriFetch } from "@tauri-apps/plugin-http";
import { deletePasswords, getPasswords, setPasswords } from "tauri-plugin-keyring-store-api";
import type { CloudProvider } from "./cloud-sync";

type Token = { accessToken: string; refreshToken?: string; expiresAt: number; clientId: string };
type TokenResponse = { access_token?: string; refresh_token?: string; expires_in?: number; error?: string };
export type DeviceCode = { device_code: string; user_code: string; verification_uri: string; interval: number; expires_in: number };
const account = (provider: CloudProvider) => `docxvim.oauth.${provider}`;
const microsoftBase = "https://login.microsoftonline.com/common/oauth2/v2.0";
const yandexBase = "https://oauth.yandex.com";

async function postForm(url: string, fields: Record<string, string>): Promise<TokenResponse & Record<string, unknown>> {
  const response = await tauriFetch(url, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams(fields) });
  const data = await response.json() as TokenResponse & Record<string, unknown>;
  if (!response.ok && data.error !== "authorization_pending" && data.error !== "slow_down") throw new Error(data.error ?? `HTTP ${response.status}`);
  return data;
}

async function readToken(provider: CloudProvider): Promise<Token | null> {
  const [raw] = await getPasswords([account(provider)]);
  if (!raw) return null;
  try { return JSON.parse(raw) as Token; } catch { return null; }
}

async function storeToken(provider: CloudProvider, clientId: string, data: TokenResponse, previous?: Token): Promise<void> {
  if (!data.access_token) throw new Error("OAuth token missing");
  const token: Token = {
    accessToken: data.access_token,
    refreshToken: data.refresh_token ?? previous?.refreshToken,
    expiresAt: Date.now() + Math.max(60, Number(data.expires_in) || 3600) * 1000,
    clientId,
  };
  await setPasswords([{ account: account(provider), secret: JSON.stringify(token) }]);
}

export async function isConnected(provider: CloudProvider, clientId: string): Promise<boolean> {
  if (!clientId) return false;
  const token = await readToken(provider);
  return !!token && token.clientId === clientId;
}

export async function disconnect(provider: CloudProvider): Promise<void> { await deletePasswords([account(provider)]); }

export async function accessToken(provider: CloudProvider, clientId: string): Promise<string> {
  const token = await readToken(provider);
  if (!token || token.clientId !== clientId) throw new Error("Account not connected");
  if (token.expiresAt > Date.now() + 90_000) return token.accessToken;
  if (!token.refreshToken) throw new Error("Account needs authorization");
  const base = provider === "oneDrive" ? microsoftBase : yandexBase;
  const fields = { grant_type: "refresh_token", refresh_token: token.refreshToken, client_id: clientId };
  const refreshed = await postForm(`${base}/token`, provider === "oneDrive" ? { ...fields, scope: "offline_access Files.ReadWrite" } : fields);
  await storeToken(provider, clientId, refreshed, token);
  return refreshed.access_token!;
}

export async function requestOneDriveCode(clientId: string): Promise<DeviceCode> {
  const data = await postForm(`${microsoftBase}/devicecode`, { client_id: clientId, scope: "offline_access Files.ReadWrite" });
  if (typeof data.device_code !== "string" || typeof data.user_code !== "string" || typeof data.verification_uri !== "string") throw new Error("Device code missing");
  return data as unknown as DeviceCode;
}

export async function completeOneDrive(clientId: string, code: DeviceCode, signal: AbortSignal): Promise<void> {
  const deadline = Date.now() + code.expires_in * 1000;
  let interval = Math.max(5, code.interval || 5) * 1000;
  while (Date.now() < deadline && !signal.aborted) {
    await new Promise<void>((resolve) => window.setTimeout(resolve, interval));
    if (signal.aborted) return;
    const data = await postForm(`${microsoftBase}/token`, {
      grant_type: "urn:ietf:params:oauth:grant-type:device_code", client_id: clientId, device_code: code.device_code,
    });
    if (data.error === "authorization_pending") continue;
    if (data.error === "slow_down") { interval += 5000; continue; }
    await storeToken("oneDrive", clientId, data);
    return;
  }
  throw new Error("Device code expired");
}

function base64Url(bytes: Uint8Array): string {
  return btoa(String.fromCharCode(...bytes)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

export async function beginYandex(clientId: string): Promise<{ url: string; verifier: string }> {
  const verifier = base64Url(crypto.getRandomValues(new Uint8Array(48)));
  const challenge = base64Url(new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))));
  const url = new URL(`${yandexBase}/authorize`);
  url.search = new URLSearchParams({ response_type: "code", client_id: clientId,
    redirect_uri: `${yandexBase}/verification_code`, scope: "cloud_api:disk.read,cloud_api:disk.write",
    code_challenge: challenge, code_challenge_method: "S256" }).toString();
  return { url: url.toString(), verifier };
}

export async function completeYandex(clientId: string, code: string, verifier: string): Promise<void> {
  const data = await postForm(`${yandexBase}/token`, { grant_type: "authorization_code", code: code.trim(),
    client_id: clientId, code_verifier: verifier });
  await storeToken("yandexDisk", clientId, data);
}
