import type { BrowserWindow } from 'electron';
import type { ResolvedConfig } from '../../types';

export type BuildAnilistSetupUrlDeps = {
  authorizeUrl: string;
  clientId: string;
  responseType: string;
  redirectUri?: string;
};

export type ConsumeAnilistSetupCallbackUrlDeps = {
  rawUrl: string;
  saveToken: (token: string) => boolean;
  setCachedToken: (token: string) => void;
  setResolvedState: (resolvedAt: number) => void;
  setSetupPageOpened: (opened: boolean) => void;
  onSuccess: () => void;
  closeWindow: () => void;
};

export function isAnilistTrackingEnabled(resolved: ResolvedConfig): boolean {
  return resolved.anilist.enabled;
}

export function buildAnilistSetupUrl(params: BuildAnilistSetupUrlDeps): string {
  const authorizeUrl = new URL(params.authorizeUrl);
  authorizeUrl.searchParams.set('client_id', params.clientId);
  authorizeUrl.searchParams.set('response_type', params.responseType);
  if (params.redirectUri && params.redirectUri.trim().length > 0) {
    authorizeUrl.searchParams.set('redirect_uri', params.redirectUri);
  }
  return authorizeUrl.toString();
}

export function extractAnilistAccessTokenFromUrl(rawUrl: string): string | null {
  try {
    const parsed = new URL(rawUrl);

    const fromQuery = parsed.searchParams.get('access_token')?.trim();
    if (fromQuery && fromQuery.length > 0) {
      return fromQuery;
    }

    const hash = parsed.hash.startsWith('#') ? parsed.hash.slice(1) : parsed.hash;
    if (hash.length === 0) {
      return null;
    }
    const hashParams = new URLSearchParams(hash);
    const fromHash = hashParams.get('access_token')?.trim();
    if (fromHash && fromHash.length > 0) {
      return fromHash;
    }
    return null;
  } catch {
    return null;
  }
}

export function findAnilistSetupDeepLinkArgvUrl(argv: readonly string[]): string | null {
  for (const value of argv) {
    if (value.startsWith('subminer://anilist-setup')) {
      return value;
    }
  }
  return null;
}

export function consumeAnilistSetupCallbackUrl(deps: ConsumeAnilistSetupCallbackUrlDeps): boolean {
  const token = extractAnilistAccessTokenFromUrl(deps.rawUrl);
  if (!token) {
    return false;
  }

  if (!deps.saveToken(token)) {
    deps.setSetupPageOpened(true);
    return true;
  }

  const resolvedAt = Date.now();
  deps.setCachedToken(token);
  deps.setResolvedState(resolvedAt);
  deps.setSetupPageOpened(false);
  deps.onSuccess();
  deps.closeWindow();
  return true;
}

export function openAnilistSetupInBrowser(params: {
  authorizeUrl: string;
  openExternal: (url: string) => Promise<void>;
  logError: (message: string, error: unknown) => void;
}): void {
  void params.openExternal(params.authorizeUrl).catch((error) => {
    params.logError('Failed to open AniList authorize URL in browser', error);
  });
}

export function buildAnilistSetupFallbackHtml(params: {
  reason: string;
  authorizeUrl: string;
  developerSettingsUrl: string;
}): string {
  const safeReason = params.reason.replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const safeAuth = params.authorizeUrl.replace(/"/g, '&quot;');
  const safeDev = params.developerSettingsUrl.replace(/"/g, '&quot;');
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>AniList Setup</title></head>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; padding: 24px; line-height: 1.5;">
  <h1>AniList setup</h1>
  <p>Automatic page load failed (${safeReason}).</p>
  <p><a href="${safeAuth}">Open AniList authorize page</a></p>
  <p><a href="${safeDev}">Open AniList developer settings</a></p>
</body></html>`;
}

export function buildAnilistManualTokenEntryHtml(params: {
  authorizeUrl: string;
  developerSettingsUrl: string;
}): string {
  const safeAuth = params.authorizeUrl.replace(/"/g, '&quot;');
  const safeDev = params.developerSettingsUrl.replace(/"/g, '&quot;');
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>AniList Setup</title></head>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; padding: 24px; line-height: 1.5;">
  <h1>AniList setup</h1>
  <p>Authorize in browser, then paste the access token below.</p>
  <p><a href="${safeAuth}" target="_blank" rel="noreferrer">Open AniList authorize page</a></p>
  <p><a href="${safeDev}" target="_blank" rel="noreferrer">Open AniList developer settings</a></p>
  <form id="token-form">
    <label for="token">Access token</label><br />
    <input id="token" style="width: 100%; max-width: 760px; margin: 8px 0; padding: 8px;" autocomplete="off" />
    <br />
    <button type="submit" style="padding: 8px 12px;">Continue</button>
  </form>
  <script>
    const form = document.getElementById('token-form');
    const token = document.getElementById('token');
    form?.addEventListener('submit', (event) => {
      event.preventDefault();
      const rawToken = String(token?.value || '').trim();
      if (rawToken) {
        window.location.href = 'subminer://anilist-setup?access_token=' + encodeURIComponent(rawToken);
      }
    });
  </script>
</body></html>`;
}

export function loadAnilistSetupFallback(params: {
  setupWindow: BrowserWindow;
  reason: string;
  authorizeUrl: string;
  developerSettingsUrl: string;
  logWarn: (message: string, data: unknown) => void;
}): void {
  const html = buildAnilistSetupFallbackHtml({
    reason: params.reason,
    authorizeUrl: params.authorizeUrl,
    developerSettingsUrl: params.developerSettingsUrl,
  });
  void params.setupWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
  params.logWarn('Loaded AniList setup fallback page', { reason: params.reason });
}

export function loadAnilistManualTokenEntry(params: {
  setupWindow: BrowserWindow;
  authorizeUrl: string;
  developerSettingsUrl: string;
  logWarn: (message: string, data: unknown) => void;
}): void {
  const html = buildAnilistManualTokenEntryHtml({
    authorizeUrl: params.authorizeUrl,
    developerSettingsUrl: params.developerSettingsUrl,
  });
  void params.setupWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
  params.logWarn('Loaded AniList manual token entry page', {
    authorizeUrl: params.authorizeUrl,
  });
}
