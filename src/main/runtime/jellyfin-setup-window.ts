import { normalizeJellyfinRecentServers } from './jellyfin-cli-auth';
import { i18n } from '../../i18n/index.js';

type JellyfinSession = {
  serverUrl: string;
  username: string;
  accessToken: string;
  userId: string;
};

type JellyfinClientInfo = {
  clientName: string;
  clientVersion: string;
  deviceId: string;
};

type FocusableWindowLike = {
  focus: () => void;
};

type JellyfinSetupWebContentsLike = {
  on: (event: 'will-navigate', handler: (event: unknown, url: string) => void) => void;
  executeJavaScript?: (code: string, userGesture?: boolean) => Promise<unknown>;
};

type JellyfinSetupWindowLike = FocusableWindowLike & {
  webContents: JellyfinSetupWebContentsLike;
  loadURL: (url: string) => unknown;
  on: (event: 'closed', handler: () => void) => void;
  isDestroyed: () => boolean;
  close: () => void;
};

export type JellyfinSetupAction = 'login' | 'logout' | 'done';

export type JellyfinSetupSubmission = {
  action: JellyfinSetupAction;
  server: string;
  username: string;
  password: string;
};

export type JellyfinSetupViewState = {
  selectedServerUrl: string;
  username: string;
  hasStoredSession: boolean;
  statusMessage: string;
  statusKind: 'idle' | 'success' | 'error' | 'loading';
};

type JellyfinSetupViewOverrides = {
  selectedServerUrl?: string;
  username?: string;
  statusMessage?: string;
  statusKind?: JellyfinSetupViewState['statusKind'];
};

export type JellyfinSetupIpcResult = {
  handled: boolean;
  statusMessage?: string;
  statusKind?: JellyfinSetupViewState['statusKind'];
};

type RegisterJellyfinSetupIpcHandler = (
  handler: (submission: unknown) => Promise<JellyfinSetupIpcResult>,
) => () => void;

function escapeHtmlAttr(value: string): string {
  return value.replace(/"/g, '&quot;');
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function normalizeSetupAction(value: unknown): JellyfinSetupAction {
  return value === 'logout' || value === 'done' ? value : 'login';
}

function normalizeString(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

export function createMaybeFocusExistingJellyfinSetupWindowHandler(deps: {
  getSetupWindow: () => FocusableWindowLike | null;
}) {
  return (): boolean => {
    const window = deps.getSetupWindow();
    if (!window) {
      return false;
    }
    window.focus();
    return true;
  };
}

export function buildJellyfinSetupViewState(input: {
  config: {
    serverUrl?: string | null;
    username?: string | null;
    recentServers?: unknown[];
  };
  defaultServerUrl: string;
  hasStoredSession: boolean;
  statusMessage?: string;
  statusKind?: JellyfinSetupViewState['statusKind'];
  selectedServerUrl?: string;
  username?: string;
}): JellyfinSetupViewState {
  const configServer = normalizeJellyfinRecentServers([input.config.serverUrl || ''])[0] || '';
  const recentServers = normalizeJellyfinRecentServers(input.config.recentServers || []);
  const defaultServer = normalizeJellyfinRecentServers([input.defaultServerUrl])[0] || '';

  const selectedServerUrl =
    normalizeJellyfinRecentServers([input.selectedServerUrl || ''])[0] ||
    configServer ||
    recentServers[0] ||
    defaultServer;

  return {
    selectedServerUrl,
    username: input.username ?? input.config.username ?? '',
    hasStoredSession: input.hasStoredSession,
    statusMessage: input.statusMessage || '',
    statusKind: input.statusKind || 'idle',
  };
}

export function buildJellyfinSetupFormHtml(state: JellyfinSetupViewState): string {
  const statusClass = `status ${state.statusKind}`;
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Jellyfin Setup</title>
  <style>
    :root {
      color-scheme: dark;
      --ctp-red: #ed8796;
      --ctp-peach: #f5a97f;
      --ctp-yellow: #eed49f;
      --ctp-green: #a6da95;
      --ctp-blue: #8aadf4;
      --ctp-lavender: #b7bdf8;
      --ctp-text: #cad3f5;
      --ctp-subtext1: #b8c0e0;
      --ctp-subtext0: #a5adcb;
      --ctp-overlay2: #939ab7;
      --ctp-overlay1: #8087a2;
      --ctp-overlay0: #6e738d;
      --ctp-surface1: #494d64;
      --ctp-surface0: #363a4f;
      --ctp-base: #24273a;
      --ctp-mantle: #1e2030;
      --ctp-crust: #181926;
      --line: rgba(110, 115, 141, 0.28);
      --line-soft: rgba(110, 115, 141, 0.14);
      --text: var(--ctp-text);
      --muted: var(--ctp-subtext0);
    }
    * { box-sizing: border-box; }
    html, body { width: 100%; height: 100%; margin: 0; }
    html { background: var(--ctp-base); }
    body {
      min-height: 100vh;
      background: var(--ctp-base);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif;
      font-size: 13px;
      line-height: 1.45;
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }
    main { padding: 32px 22px; max-width: 520px; margin: 0 auto; }
    h1 { margin: 0 0 6px; font-size: 20px; font-weight: 800; color: var(--ctp-text); letter-spacing: -0.01em; }
    p { margin: 0 0 18px; color: var(--muted); font-size: 13px; line-height: 1.5; }
    label { display: block; margin: 14px 0 6px; font-size: 11px; font-weight: 800; color: var(--ctp-overlay2); text-transform: uppercase; letter-spacing: 0.1em; }
    input {
      width: 100%;
      padding: 9px 11px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: rgba(24, 25, 38, 0.85);
      color: var(--text);
      font: inherit;
      outline: none;
      transition: border-color 140ms ease, box-shadow 140ms ease, background 140ms ease;
    }
    input::placeholder { color: var(--ctp-overlay0); }
    input:hover { border-color: rgba(138, 173, 244, 0.32); }
    input:focus {
      border-color: rgba(138, 173, 244, 0.65);
      background: rgba(24, 25, 38, 0.95);
      box-shadow: 0 0 0 3px rgba(138, 173, 244, 0.15);
    }
    button {
      height: 36px;
      padding: 0 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      font: inherit;
      font-weight: 700;
      font-size: 13px;
      cursor: pointer;
      transition: background 140ms ease, border-color 140ms ease, color 140ms ease, transform 60ms ease;
    }
    button:active { transform: translateY(1px); }
    button:disabled { cursor: wait; opacity: 0.7; }
    button.primary {
      border-color: transparent;
      background: var(--ctp-blue);
      color: var(--ctp-crust);
    }
    button.primary:hover:not(:disabled) { filter: brightness(1.06); }
    button.primary:disabled {
      background: rgba(54, 58, 79, 0.55);
      color: var(--ctp-overlay0);
      border-color: var(--line);
    }
    button.secondary {
      background: rgba(54, 58, 79, 0.5);
      color: var(--text);
    }
    button.secondary:hover:not(:disabled) {
      border-color: rgba(138, 173, 244, 0.45);
      background: rgba(73, 77, 100, 0.6);
      color: var(--ctp-lavender);
    }
    button.danger {
      background: rgba(237, 135, 150, 0.12);
      color: var(--ctp-red);
      border-color: rgba(237, 135, 150, 0.45);
    }
    button.danger:hover:not(:disabled) {
      background: rgba(237, 135, 150, 0.22);
    }
    .actions { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 18px; }
    .actions .primary { grid-column: 1 / -1; }
    .status {
      min-height: 18px;
      margin-top: 14px;
      font-size: 12.5px;
      color: var(--muted);
    }
    .status:empty { display: none; }
    .status.loading,
    .status.success,
    .status.error {
      padding: 10px 12px;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: var(--ctp-surface0);
      font-weight: 600;
    }
    .status.success {
      border-color: rgba(166, 218, 149, 0.45);
      background: rgba(166, 218, 149, 0.1);
      color: var(--ctp-green);
    }
    .status.error {
      border-color: rgba(237, 135, 150, 0.55);
      background: rgba(237, 135, 150, 0.1);
      color: var(--ctp-red);
    }
    .hint {
      margin-top: 16px;
      font-size: 11.5px;
      color: var(--ctp-overlay2);
      line-height: 1.55;
    }
  </style>
</head>
<body>
  <main>
    <h1>Jellyfin Setup</h1>
    <p>Enter your Jellyfin server URL, sign in, and SubMiner will save a session token for Jellyfin commands and cast discovery.</p>
    <form id="form">
      <label for="server">Server URL</label>
      <input id="server" name="server" value="${escapeHtmlAttr(state.selectedServerUrl)}" required />
      <label for="username">Username</label>
      <input id="username" name="username" value="${escapeHtmlAttr(state.username)}" required />
      <label for="password">Password</label>
      <input id="password" name="password" type="password" required />
      <div id="status" class="${statusClass}" aria-live="polite">${escapeHtml(state.statusMessage)}</div>
      <div class="actions">
        <button class="primary" type="submit">Login</button>
        ${
          state.hasStoredSession
            ? '<button id="logout" class="danger" type="button">Logout</button>'
            : '<span></span>'
        }
        <button id="done" class="secondary" type="button">Done</button>
      </div>
      <div class="hint">Equivalent CLI: --jellyfin-login --jellyfin-server ... --jellyfin-username ... --jellyfin-password ...</div>
    </form>
  </main>
  <script>
    const form = document.getElementById("form");
    const server = document.getElementById("server");
    const username = document.getElementById("username");
    const password = document.getElementById("password");
    const status = document.getElementById("status");
    const buttons = Array.from(document.querySelectorAll("button"));
    function setBusy(message) {
      if (status) {
        status.textContent = message;
        status.className = "status loading";
      }
      for (const button of buttons) button.disabled = true;
    }
    function setStatus(message, kind) {
      if (status) {
        status.textContent = message;
        status.className = "status " + kind;
      }
      for (const button of buttons) button.disabled = false;
    }
    async function submitAction(action) {
      const serverValue = String(server?.value || "");
      const usernameValue = String(username?.value || "");
      const passwordValue = String(password?.value || "");
      setBusy(action === "login" ? "Logging in to Jellyfin..." : action === "logout" ? "Logging out..." : "Closing...");
      const bridge = window.subminerJellyfinSetup;
      if (bridge?.submit) {
        try {
          const result = await bridge.submit({
            action,
            server: serverValue,
            username: usernameValue,
            password: passwordValue,
          });
          if (result?.handled === false) {
            setStatus(result.statusMessage || "Jellyfin setup action was not accepted.", result.statusKind || "error");
          }
        } catch (error) {
          const message = error && typeof error === "object" && "message" in error ? String(error.message) : String(error || "Unknown error");
          setStatus("Jellyfin setup action failed: " + message, "error");
        }
        return;
      }
      const params = new URLSearchParams();
      params.set("action", action);
      if (action === "login") {
        params.set("server", serverValue);
        params.set("username", usernameValue);
        window.__subminerJellyfinPassword = passwordValue;
      }
      window.location.href = "subminer://jellyfin-setup?" + params.toString();
    }
    form?.addEventListener("submit", (event) => {
      event.preventDefault();
      submitAction("login");
    });
    document.getElementById("logout")?.addEventListener("click", () => submitAction("logout"));
    document.getElementById("done")?.addEventListener("click", () => submitAction("done"));
  </script>
</body>
</html>`;
}

export function parseJellyfinSetupSubmissionUrl(rawUrl: string): {
  action: JellyfinSetupAction;
  server: string;
  username: string;
  password: string;
} | null {
  if (!rawUrl.startsWith('subminer://jellyfin-setup')) {
    return null;
  }
  const parsed = new URL(rawUrl);
  const rawAction = parsed.searchParams.get('action') || 'login';
  const action: JellyfinSetupAction =
    rawAction === 'logout' || rawAction === 'done' ? rawAction : 'login';
  return {
    action,
    server: parsed.searchParams.get('server') || '',
    username: parsed.searchParams.get('username') || '',
    password: parsed.searchParams.get('password') || '',
  };
}

export function normalizeJellyfinSetupIpcSubmission(
  value: unknown,
): JellyfinSetupSubmission | null {
  if (!isRecord(value)) {
    return null;
  }
  return {
    action: normalizeSetupAction(value.action),
    server: normalizeString(value.server),
    username: normalizeString(value.username),
    password: normalizeString(value.password),
  };
}

export function buildJellyfinSetupSubmissionUrl(submission: JellyfinSetupSubmission): string {
  const params = new URLSearchParams();
  params.set('action', submission.action);
  if (submission.action === 'login') {
    params.set('server', submission.server);
    params.set('username', submission.username);
  }
  return `subminer://jellyfin-setup?${params.toString()}`;
}

export function createHandleJellyfinSetupSubmissionHandler(deps: {
  parseSubmissionUrl: (
    rawUrl: string,
  ) => { action: JellyfinSetupAction; server: string; username: string; password: string } | null;
  authenticateWithPassword: (
    server: string,
    username: string,
    password: string,
    clientInfo: JellyfinClientInfo,
  ) => Promise<JellyfinSession>;
  getJellyfinClientInfo: () => JellyfinClientInfo;
  saveStoredSession: (session: { accessToken: string; userId: string }) => void;
  clearStoredSession: () => void;
  patchJellyfinConfig: (session: JellyfinSession) => void;
  persistAuthenticatedSession?: (session: JellyfinSession, clientInfo: JellyfinClientInfo) => void;
  restartRemoteSession?: () => Promise<void> | void;
  stopRemoteSession?: () => void;
  logInfo: (message: string) => void;
  logError: (message: string, error: unknown) => void;
  showMpvOsd: (message: string) => void;
  closeSetupWindow: () => void;
  reloadSetupWindow: (state?: JellyfinSetupViewOverrides) => void;
}) {
  let loginInFlight = false;

  return async (rawUrl: string, passwordOverride?: string): Promise<boolean> => {
    const submission = deps.parseSubmissionUrl(rawUrl);
    if (!submission) {
      return false;
    }

    if (submission.action === 'done') {
      deps.closeSetupWindow();
      return true;
    }

    if (submission.action === 'logout') {
      try {
        deps.clearStoredSession();
        deps.stopRemoteSession?.();
        deps.logInfo('Cleared stored Jellyfin auth session.');
        deps.showMpvOsd(i18n.t('osd.jellyfinLoggedOut'));
        deps.reloadSetupWindow({
          statusMessage: i18n.t('jellyfin.sessionCleared'),
          statusKind: 'success',
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        deps.logError('Jellyfin logout failed', error);
        deps.showMpvOsd(i18n.t('osd.jellyfinLogoutFailed', { message }));
        deps.reloadSetupWindow({
          statusMessage: message,
          statusKind: 'error',
        });
      }
      return true;
    }

    if (loginInFlight) {
      deps.showMpvOsd(i18n.t('osd.jellyfinLoginProgress'));
      deps.reloadSetupWindow({
        selectedServerUrl: submission.server,
        username: submission.username,
        statusMessage: i18n.t('jellyfin.loginInProgress'),
        statusKind: 'loading',
      });
      return true;
    }

    loginInFlight = true;
    try {
      const clientInfo = deps.getJellyfinClientInfo();
      const session = await deps.authenticateWithPassword(
        submission.server,
        submission.username,
        passwordOverride ?? submission.password,
        clientInfo,
      );
      if (deps.persistAuthenticatedSession) {
        deps.persistAuthenticatedSession(session, clientInfo);
      } else {
        deps.saveStoredSession({ accessToken: session.accessToken, userId: session.userId });
        deps.patchJellyfinConfig(session);
      }
      await deps.restartRemoteSession?.();
      deps.logInfo(`Jellyfin setup saved for ${session.username}.`);
      deps.showMpvOsd(i18n.t('osd.jellyfinLoginSuccess'));
      deps.reloadSetupWindow({
        selectedServerUrl: session.serverUrl,
        username: session.username,
        statusMessage: i18n.t('jellyfin.loginSuccess', { username: session.username }),
        statusKind: 'success',
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      deps.logError('Jellyfin setup failed', error);
      deps.showMpvOsd(i18n.t('osd.jellyfinLoginFailed', { message }));
      deps.reloadSetupWindow({
        selectedServerUrl: submission.server,
        username: submission.username,
        statusMessage: message,
        statusKind: 'error',
      });
    } finally {
      loginInFlight = false;
    }
    return true;
  };
}

export function createHandleJellyfinSetupNavigationHandler(deps: {
  setupSchemePrefix: string;
  handleSubmission: (rawUrl: string) => Promise<unknown>;
  logError: (message: string, error: unknown) => void;
}) {
  return (params: { url: string; preventDefault: () => void }): boolean => {
    if (!params.url.startsWith(deps.setupSchemePrefix)) {
      return false;
    }
    params.preventDefault();
    void deps.handleSubmission(params.url).catch((error) => {
      deps.logError('Failed handling Jellyfin setup submission', error);
    });
    return true;
  };
}

async function readJellyfinSetupPasswordFromWindow(
  setupWindow: JellyfinSetupWindowLike,
): Promise<string | undefined> {
  const executeJavaScript = setupWindow.webContents.executeJavaScript;
  if (!executeJavaScript) {
    return undefined;
  }

  const value = await executeJavaScript(
    `(() => {
      const input = document.getElementById("password");
      const password = String(window.__subminerJellyfinPassword || input?.value || "");
      window.__subminerJellyfinPassword = "";
      if (input) input.value = "";
      return password;
    })()`,
    true,
  );
  return typeof value === 'string' ? value : '';
}

export function createHandleJellyfinSetupWindowClosedHandler(deps: {
  clearSetupWindow: () => void;
}) {
  return (): void => {
    deps.clearSetupWindow();
  };
}

export function createHandleJellyfinSetupWindowOpenedHandler(deps: { setSetupWindow: () => void }) {
  return (): void => {
    deps.setSetupWindow();
  };
}

export function createOpenJellyfinSetupWindowHandler<
  TWindow extends JellyfinSetupWindowLike,
>(deps: {
  maybeFocusExistingSetupWindow: () => boolean;
  createSetupWindow: () => TWindow;
  getResolvedJellyfinConfig: () => {
    serverUrl?: string | null;
    username?: string | null;
    recentServers?: unknown[];
  };
  buildSetupFormHtml: (state: JellyfinSetupViewState) => string;
  parseSubmissionUrl: (
    rawUrl: string,
  ) => { action: JellyfinSetupAction; server: string; username: string; password: string } | null;
  authenticateWithPassword: (
    server: string,
    username: string,
    password: string,
    clientInfo: JellyfinClientInfo,
  ) => Promise<JellyfinSession>;
  getJellyfinClientInfo: () => JellyfinClientInfo;
  saveStoredSession: (session: { accessToken: string; userId: string }) => void;
  clearStoredSession: () => void;
  patchJellyfinConfig: (session: JellyfinSession) => void;
  persistAuthenticatedSession?: (session: JellyfinSession, clientInfo: JellyfinClientInfo) => void;
  restartRemoteSession?: () => Promise<void> | void;
  stopRemoteSession?: () => void;
  logInfo: (message: string) => void;
  logError: (message: string, error: unknown) => void;
  showMpvOsd: (message: string) => void;
  clearSetupWindow: () => void;
  setSetupWindow: (window: TWindow) => void;
  registerSetupIpcHandler?: RegisterJellyfinSetupIpcHandler;
  encodeURIComponent: (value: string) => string;
  defaultServerUrl: string;
  hasStoredSession: () => boolean;
}) {
  return (): void => {
    if (deps.maybeFocusExistingSetupWindow()) {
      return;
    }

    const setupWindow = deps.createSetupWindow();
    const loadSetupForm = (overrides: JellyfinSetupViewOverrides = {}) => {
      const state = buildJellyfinSetupViewState({
        config: deps.getResolvedJellyfinConfig(),
        defaultServerUrl: deps.defaultServerUrl,
        hasStoredSession: deps.hasStoredSession(),
        selectedServerUrl: overrides.selectedServerUrl,
        username: overrides.username,
        statusMessage: overrides.statusMessage,
        statusKind: overrides.statusKind,
      });
      const formHtml = deps.buildSetupFormHtml(state);
      void setupWindow.loadURL(`data:text/html;charset=utf-8,${deps.encodeURIComponent(formHtml)}`);
    };
    const handleSubmission = createHandleJellyfinSetupSubmissionHandler({
      parseSubmissionUrl: (rawUrl) => deps.parseSubmissionUrl(rawUrl),
      authenticateWithPassword: (server, username, password, clientInfo) =>
        deps.authenticateWithPassword(server, username, password, clientInfo),
      getJellyfinClientInfo: () => deps.getJellyfinClientInfo(),
      saveStoredSession: (session) => deps.saveStoredSession(session),
      clearStoredSession: () => deps.clearStoredSession(),
      patchJellyfinConfig: (session) => deps.patchJellyfinConfig(session),
      persistAuthenticatedSession: deps.persistAuthenticatedSession
        ? (session, clientInfo) => deps.persistAuthenticatedSession?.(session, clientInfo)
        : undefined,
      restartRemoteSession: deps.restartRemoteSession
        ? () => deps.restartRemoteSession?.()
        : undefined,
      stopRemoteSession: deps.stopRemoteSession ? () => deps.stopRemoteSession?.() : undefined,
      logInfo: (message) => deps.logInfo(message),
      logError: (message, error) => deps.logError(message, error),
      showMpvOsd: (message) => deps.showMpvOsd(message),
      closeSetupWindow: () => {
        if (!setupWindow.isDestroyed()) {
          setupWindow.close();
        }
      },
      reloadSetupWindow: (state) => {
        if (!setupWindow.isDestroyed()) {
          loadSetupForm(state);
        }
      },
    });
    const unregisterSetupIpcHandler = deps.registerSetupIpcHandler?.(async (payload) => {
      const submission = normalizeJellyfinSetupIpcSubmission(payload);
      if (!submission) {
        return {
          handled: false,
          statusMessage: 'Invalid Jellyfin setup request.',
          statusKind: 'error',
        };
      }
      const handled = await handleSubmission(
        buildJellyfinSetupSubmissionUrl(submission),
        submission.password,
      );
      return { handled };
    });
    const handleNavigation = createHandleJellyfinSetupNavigationHandler({
      setupSchemePrefix: 'subminer://jellyfin-setup',
      handleSubmission: async (rawUrl) => {
        const submission = deps.parseSubmissionUrl(rawUrl);
        let password: string | undefined;
        if (submission?.action === 'login' && !submission.password) {
          try {
            password = await readJellyfinSetupPasswordFromWindow(setupWindow);
          } catch (error) {
            deps.logError('Failed reading Jellyfin setup password', error);
            password = '';
          }
        }
        return handleSubmission(rawUrl, password);
      },
      logError: (message, error) => deps.logError(message, error),
    });
    const handleWindowClosed = createHandleJellyfinSetupWindowClosedHandler({
      clearSetupWindow: () => deps.clearSetupWindow(),
    });
    const handleWindowOpened = createHandleJellyfinSetupWindowOpenedHandler({
      setSetupWindow: () => deps.setSetupWindow(setupWindow),
    });

    setupWindow.webContents.on('will-navigate', (event, url) => {
      handleNavigation({
        url,
        preventDefault: () => {
          if (event && typeof event === 'object' && 'preventDefault' in event) {
            const typedEvent = event as { preventDefault?: () => void };
            typedEvent.preventDefault?.();
          }
        },
      });
    });
    loadSetupForm();
    setupWindow.on('closed', () => {
      unregisterSetupIpcHandler?.();
      handleWindowClosed();
    });
    handleWindowOpened();
  };
}
