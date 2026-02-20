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
};

type JellyfinSetupWindowLike = FocusableWindowLike & {
  webContents: JellyfinSetupWebContentsLike;
  loadURL: (url: string) => unknown;
  on: (event: 'closed', handler: () => void) => void;
  isDestroyed: () => boolean;
  close: () => void;
};

function escapeHtmlAttr(value: string): string {
  return value.replace(/"/g, '&quot;');
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

export function buildJellyfinSetupFormHtml(defaultServer: string, defaultUser: string): string {
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Jellyfin Setup</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 0; background: #0b1020; color: #e5e7eb; }
    main { padding: 20px; }
    h1 { margin: 0 0 8px; font-size: 22px; }
    p { margin: 0 0 14px; color: #cbd5e1; font-size: 13px; line-height: 1.4; }
    label { display: block; margin: 10px 0 4px; font-size: 13px; }
    input { width: 100%; box-sizing: border-box; padding: 9px 10px; border: 1px solid #334155; border-radius: 8px; background: #111827; color: #e5e7eb; }
    button { margin-top: 16px; width: 100%; padding: 10px 12px; border: 0; border-radius: 8px; font-weight: 600; cursor: pointer; background: #2563eb; color: #f8fafc; }
    .hint { margin-top: 12px; font-size: 12px; color: #94a3b8; }
  </style>
</head>
<body>
  <main>
    <h1>Jellyfin Setup</h1>
    <p>Login info is used to fetch a token and save Jellyfin config values.</p>
    <form id="form">
      <label for="server">Server URL</label>
      <input id="server" name="server" value="${escapeHtmlAttr(defaultServer)}" required />
      <label for="username">Username</label>
      <input id="username" name="username" value="${escapeHtmlAttr(defaultUser)}" required />
      <label for="password">Password</label>
      <input id="password" name="password" type="password" required />
      <button type="submit">Save and Login</button>
      <div class="hint">Equivalent CLI: --jellyfin-login --jellyfin-server ... --jellyfin-username ... --jellyfin-password ...</div>
    </form>
  </main>
  <script>
    const form = document.getElementById("form");
    form?.addEventListener("submit", (event) => {
      event.preventDefault();
      const data = new FormData(form);
      const params = new URLSearchParams();
      params.set("server", String(data.get("server") || ""));
      params.set("username", String(data.get("username") || ""));
      params.set("password", String(data.get("password") || ""));
      window.location.href = "subminer://jellyfin-setup?" + params.toString();
    });
  </script>
</body>
</html>`;
}

export function parseJellyfinSetupSubmissionUrl(rawUrl: string): {
  server: string;
  username: string;
  password: string;
} | null {
  if (!rawUrl.startsWith('subminer://jellyfin-setup')) {
    return null;
  }
  const parsed = new URL(rawUrl);
  return {
    server: parsed.searchParams.get('server') || '',
    username: parsed.searchParams.get('username') || '',
    password: parsed.searchParams.get('password') || '',
  };
}

export function createHandleJellyfinSetupSubmissionHandler(deps: {
  parseSubmissionUrl: (rawUrl: string) => { server: string; username: string; password: string } | null;
  authenticateWithPassword: (
    server: string,
    username: string,
    password: string,
    clientInfo: JellyfinClientInfo,
  ) => Promise<JellyfinSession>;
  getJellyfinClientInfo: () => JellyfinClientInfo;
  patchJellyfinConfig: (session: JellyfinSession) => void;
  logInfo: (message: string) => void;
  logError: (message: string, error: unknown) => void;
  showMpvOsd: (message: string) => void;
  closeSetupWindow: () => void;
}) {
  return async (rawUrl: string): Promise<boolean> => {
    const submission = deps.parseSubmissionUrl(rawUrl);
    if (!submission) {
      return false;
    }

    try {
      const session = await deps.authenticateWithPassword(
        submission.server,
        submission.username,
        submission.password,
        deps.getJellyfinClientInfo(),
      );
      deps.patchJellyfinConfig(session);
      deps.logInfo(`Jellyfin setup saved for ${session.username}.`);
      deps.showMpvOsd('Jellyfin login success');
      deps.closeSetupWindow();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      deps.logError('Jellyfin setup failed', error);
      deps.showMpvOsd(`Jellyfin login failed: ${message}`);
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

export function createHandleJellyfinSetupWindowClosedHandler(deps: {
  clearSetupWindow: () => void;
}) {
  return (): void => {
    deps.clearSetupWindow();
  };
}

export function createHandleJellyfinSetupWindowOpenedHandler(deps: {
  setSetupWindow: () => void;
}) {
  return (): void => {
    deps.setSetupWindow();
  };
}

export function createOpenJellyfinSetupWindowHandler<TWindow extends JellyfinSetupWindowLike>(deps: {
  maybeFocusExistingSetupWindow: () => boolean;
  createSetupWindow: () => TWindow;
  getResolvedJellyfinConfig: () => { serverUrl?: string | null; username?: string | null };
  buildSetupFormHtml: (defaultServer: string, defaultUser: string) => string;
  parseSubmissionUrl: (rawUrl: string) => { server: string; username: string; password: string } | null;
  authenticateWithPassword: (
    server: string,
    username: string,
    password: string,
    clientInfo: JellyfinClientInfo,
  ) => Promise<JellyfinSession>;
  getJellyfinClientInfo: () => JellyfinClientInfo;
  patchJellyfinConfig: (session: JellyfinSession) => void;
  logInfo: (message: string) => void;
  logError: (message: string, error: unknown) => void;
  showMpvOsd: (message: string) => void;
  clearSetupWindow: () => void;
  setSetupWindow: (window: TWindow) => void;
  encodeURIComponent: (value: string) => string;
}) {
  return (): void => {
    if (deps.maybeFocusExistingSetupWindow()) {
      return;
    }

    const setupWindow = deps.createSetupWindow();
    const defaults = deps.getResolvedJellyfinConfig();
    const defaultServer = defaults.serverUrl || 'http://127.0.0.1:8096';
    const defaultUser = defaults.username || '';
    const formHtml = deps.buildSetupFormHtml(defaultServer, defaultUser);
    const handleSubmission = createHandleJellyfinSetupSubmissionHandler({
      parseSubmissionUrl: (rawUrl) => deps.parseSubmissionUrl(rawUrl),
      authenticateWithPassword: (server, username, password, clientInfo) =>
        deps.authenticateWithPassword(server, username, password, clientInfo),
      getJellyfinClientInfo: () => deps.getJellyfinClientInfo(),
      patchJellyfinConfig: (session) => deps.patchJellyfinConfig(session),
      logInfo: (message) => deps.logInfo(message),
      logError: (message, error) => deps.logError(message, error),
      showMpvOsd: (message) => deps.showMpvOsd(message),
      closeSetupWindow: () => {
        if (!setupWindow.isDestroyed()) {
          setupWindow.close();
        }
      },
    });
    const handleNavigation = createHandleJellyfinSetupNavigationHandler({
      setupSchemePrefix: 'subminer://jellyfin-setup',
      handleSubmission: (rawUrl) => handleSubmission(rawUrl),
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
    void setupWindow.loadURL(
      `data:text/html;charset=utf-8,${deps.encodeURIComponent(formHtml)}`,
    );
    setupWindow.on('closed', () => {
      handleWindowClosed();
    });
    handleWindowOpened();
  };
}
