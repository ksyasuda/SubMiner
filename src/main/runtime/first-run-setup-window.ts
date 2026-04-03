import { getFirstRunSetupCompletionMessage } from './first-run-setup-service';

type FocusableWindowLike = {
  focus: () => void;
};

type FirstRunSetupWebContentsLike = {
  on: (event: 'will-navigate', handler: (event: unknown, url: string) => void) => void;
};

type FirstRunSetupWindowLike = FocusableWindowLike & {
  webContents: FirstRunSetupWebContentsLike;
  loadURL: (url: string) => unknown;
  on: (event: 'closed', handler: () => void) => void;
  isDestroyed: () => boolean;
  close: () => void;
};

export type FirstRunSetupAction =
  | 'configure-mpv-executable-path'
  | 'install-plugin'
  | 'configure-windows-mpv-shortcuts'
  | 'open-yomitan-settings'
  | 'refresh'
  | 'finish';

export interface FirstRunSetupSubmission {
  action: FirstRunSetupAction;
  mpvExecutablePath?: string;
  startMenuEnabled?: boolean;
  desktopEnabled?: boolean;
}

export interface FirstRunSetupHtmlModel {
  configReady: boolean;
  dictionaryCount: number;
  canFinish: boolean;
  externalYomitanConfigured: boolean;
  pluginStatus: 'installed' | 'required' | 'failed';
  pluginInstallPathSummary: string | null;
  mpvExecutablePath: string;
  mpvExecutablePathStatus: 'blank' | 'configured' | 'invalid';
  windowsMpvShortcuts: {
    supported: boolean;
    startMenuEnabled: boolean;
    desktopEnabled: boolean;
    startMenuInstalled: boolean;
    desktopInstalled: boolean;
    status: 'installed' | 'optional' | 'skipped' | 'failed';
  };
  message: string | null;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function renderStatusBadge(value: string, tone: 'ready' | 'warn' | 'muted' | 'danger'): string {
  return `<span class="badge ${tone}">${escapeHtml(value)}</span>`;
}

export function buildFirstRunSetupHtml(model: FirstRunSetupHtmlModel): string {
  const pluginActionLabel =
    model.pluginStatus === 'installed' ? 'Reinstall mpv plugin' : 'Install mpv plugin';
  const pluginLabel =
    model.pluginStatus === 'installed'
      ? 'Installed'
      : model.pluginStatus === 'failed'
        ? 'Failed'
        : 'Required';
  const pluginTone =
    model.pluginStatus === 'installed'
      ? 'ready'
      : model.pluginStatus === 'failed'
        ? 'danger'
        : 'warn';
  const windowsShortcutLabel =
    model.windowsMpvShortcuts.status === 'installed'
      ? 'Installed'
      : model.windowsMpvShortcuts.status === 'skipped'
        ? 'Skipped'
        : model.windowsMpvShortcuts.status === 'failed'
          ? 'Failed'
          : 'Optional';
  const windowsShortcutTone =
    model.windowsMpvShortcuts.status === 'installed'
      ? 'ready'
      : model.windowsMpvShortcuts.status === 'failed'
        ? 'danger'
        : model.windowsMpvShortcuts.status === 'skipped'
          ? 'muted'
          : 'warn';
  const mpvExecutablePathLabel =
    model.mpvExecutablePathStatus === 'configured'
      ? 'Configured'
      : model.mpvExecutablePathStatus === 'invalid'
        ? 'Invalid'
        : 'Blank';
  const mpvExecutablePathTone =
    model.mpvExecutablePathStatus === 'configured'
      ? 'ready'
      : model.mpvExecutablePathStatus === 'invalid'
        ? 'danger'
        : 'muted';
  const mpvExecutablePathCurrent =
    model.mpvExecutablePathStatus === 'blank'
      ? 'blank (PATH discovery)'
      : model.mpvExecutablePathStatus === 'invalid'
        ? `${model.mpvExecutablePath} (invalid; file not found)`
        : model.mpvExecutablePath;
  const mpvExecutablePathCard = model.windowsMpvShortcuts.supported
    ? `
    <div class="card block">
      <div class="card-head">
        <div>
          <strong>mpv executable path</strong>
          <div class="meta">Leave blank to auto-discover mpv.exe from PATH.</div>
          <div class="meta">Current: ${escapeHtml(mpvExecutablePathCurrent)}</div>
        </div>
        ${renderStatusBadge(mpvExecutablePathLabel, mpvExecutablePathTone)}
      </div>
      <form
        class="path-form"
        onsubmit="event.preventDefault(); const params = new URLSearchParams({ action: 'configure-mpv-executable-path', mpvExecutablePath: document.getElementById('mpv-executable-path').value }); window.location.href = 'subminer://first-run-setup?' + params.toString();"
      >
        <input
          id="mpv-executable-path"
          type="text"
          value="${escapeHtml(model.mpvExecutablePath)}"
          placeholder="C:\\Program Files\\mpv\\mpv.exe"
        />
        <button type="submit">Save mpv executable path</button>
      </form>
    </div>`
    : '';
  const windowsShortcutCard = model.windowsMpvShortcuts.supported
    ? `
    <div class="card block">
      <div class="card-head">
        <div>
          <strong>Windows mpv launcher</strong>
          <div class="meta">Create standalone \`SubMiner mpv\` shortcuts that run \`SubMiner.exe --launch-mpv\`.</div>
          <div class="meta">Installed: Start Menu ${model.windowsMpvShortcuts.startMenuInstalled ? 'yes' : 'no'}, Desktop ${model.windowsMpvShortcuts.desktopInstalled ? 'yes' : 'no'}</div>
        </div>
        ${renderStatusBadge(windowsShortcutLabel, windowsShortcutTone)}
      </div>
      <form
        class="shortcut-form"
        onsubmit="event.preventDefault(); const params = new URLSearchParams({ action: 'configure-windows-mpv-shortcuts', startMenu: document.getElementById('shortcut-start-menu').checked ? '1' : '0', desktop: document.getElementById('shortcut-desktop').checked ? '1' : '0' }); window.location.href = 'subminer://first-run-setup?' + params.toString();"
      >
        <label><input id="shortcut-start-menu" type="checkbox" ${model.windowsMpvShortcuts.startMenuEnabled ? 'checked' : ''} /> Create Start Menu shortcut</label>
        <label><input id="shortcut-desktop" type="checkbox" ${model.windowsMpvShortcuts.desktopEnabled ? 'checked' : ''} /> Create Desktop shortcut</label>
        <button type="submit">Apply mpv launcher shortcuts</button>
      </form>
    </div>`
    : '';

  const yomitanMeta = model.externalYomitanConfigured
    ? 'External profile configured. SubMiner is reusing that Yomitan profile for this setup run.'
    : `${model.dictionaryCount} installed`;
  const yomitanBadgeLabel = model.externalYomitanConfigured
    ? 'External'
    : model.dictionaryCount >= 1
      ? 'Ready'
      : 'Missing';
  const yomitanBadgeTone = model.externalYomitanConfigured
    ? 'ready'
    : model.dictionaryCount >= 1
      ? 'ready'
      : 'warn';
  const blockerMessage = getFirstRunSetupCompletionMessage(model);
  const footerMessage = blockerMessage
    ? blockerMessage
    : model.canFinish
      ? model.externalYomitanConfigured
        ? 'Finish stays unlocked while SubMiner is reusing an external Yomitan profile. If you later launch without yomitan.externalProfilePath, setup will require at least one internal dictionary.'
        : 'Finish stays unlocked once the mpv plugin is installed and Yomitan reports at least one installed dictionary.'
      : 'Finish stays locked until the mpv plugin is installed and Yomitan reports at least one installed dictionary.';

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>SubMiner First-Run Setup</title>
  <style>
    :root {
      color-scheme: dark;
      --base: #24273a;
      --mantle: #1e2030;
      --surface: #363a4f;
      --surface-strong: #494d64;
      --text: #cad3f5;
      --muted: #b8c0e0;
      --blue: #8aadf4;
      --green: #a6da95;
      --yellow: #eed49f;
      --red: #ed8796;
    }
    body {
      margin: 0;
      background: linear-gradient(180deg, var(--mantle), var(--base));
      color: var(--text);
      font: 13px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    main {
      padding: 18px;
    }
    h1 {
      margin: 0 0 6px;
      font-size: 18px;
    }
    p {
      margin: 0 0 14px;
      color: var(--muted);
    }
    .card {
      background: rgba(54, 58, 79, 0.92);
      border: 1px solid rgba(202, 211, 245, 0.08);
      border-radius: 12px;
      padding: 12px;
      margin-bottom: 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
    }
    .card.block {
      display: block;
    }
    .card-head {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
    }
    .meta {
      color: var(--muted);
      font-size: 12px;
    }
    .shortcut-form {
      display: grid;
      gap: 8px;
      margin-top: 12px;
    }
    label {
      color: var(--muted);
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .badge {
      display: inline-flex;
      align-items: center;
      border-radius: 999px;
      padding: 4px 9px;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.03em;
    }
    .badge.ready { background: rgba(166, 218, 149, 0.16); color: var(--green); }
    .badge.warn { background: rgba(238, 212, 159, 0.18); color: var(--yellow); }
    .badge.muted { background: rgba(184, 192, 224, 0.12); color: var(--muted); }
    .badge.danger { background: rgba(237, 135, 150, 0.16); color: var(--red); }
    .path-form {
      display: grid;
      gap: 8px;
      margin-top: 12px;
    }
    .path-form input[type='text'] {
      width: 100%;
      box-sizing: border-box;
      border: 1px solid rgba(202, 211, 245, 0.12);
      border-radius: 10px;
      padding: 9px 10px;
      color: var(--text);
      background: rgba(30, 32, 48, 0.72);
      font: inherit;
    }
    .path-form input[type='text']::placeholder {
      color: rgba(184, 192, 224, 0.65);
    }
    .actions {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      margin-top: 14px;
    }
    button {
      border: 0;
      border-radius: 10px;
      padding: 10px 12px;
      cursor: pointer;
      font-weight: 700;
      color: var(--text);
      background: var(--surface);
    }
    button.primary {
      background: var(--blue);
      color: #1e2030;
    }
    button.ghost {
      background: transparent;
      border: 1px solid rgba(202, 211, 245, 0.12);
    }
    button:disabled {
      cursor: not-allowed;
      opacity: 0.55;
    }
    .message {
      min-height: 18px;
      margin-top: 12px;
      color: var(--muted);
    }
    .footer {
      margin-top: 10px;
      color: var(--muted);
      font-size: 12px;
    }
  </style>
</head>
<body>
  <main>
    <h1>SubMiner setup</h1>
    <div class="card">
      <div>
        <strong>Config file</strong>
        <div class="meta">Default config directory seeded automatically.</div>
      </div>
      ${renderStatusBadge(model.configReady ? 'Ready' : 'Missing', model.configReady ? 'ready' : 'danger')}
    </div>
    <div class="card">
      <div>
        <strong>mpv plugin</strong>
        <div class="meta">${escapeHtml(model.pluginInstallPathSummary ?? 'Default mpv scripts location')}</div>
        <div class="meta">Required before SubMiner setup can finish.</div>
      </div>
      ${renderStatusBadge(pluginLabel, pluginTone)}
    </div>
    <div class="card">
      <div>
        <strong>Yomitan dictionaries</strong>
        <div class="meta">${escapeHtml(yomitanMeta)}</div>
      </div>
      ${renderStatusBadge(yomitanBadgeLabel, yomitanBadgeTone)}
    </div>
    ${mpvExecutablePathCard}
    ${windowsShortcutCard}
    <div class="actions">
      <button onclick="window.location.href='subminer://first-run-setup?action=install-plugin'">${pluginActionLabel}</button>
      <button onclick="window.location.href='subminer://first-run-setup?action=open-yomitan-settings'">Open Yomitan Settings</button>
      <button class="ghost" onclick="window.location.href='subminer://first-run-setup?action=refresh'">Refresh status</button>
      <button class="primary" ${model.canFinish ? '' : 'disabled'} onclick="window.location.href='subminer://first-run-setup?action=finish'">Finish setup</button>
    </div>
    <div class="message">${model.message ? escapeHtml(model.message) : ''}</div>
    <div class="footer">${escapeHtml(footerMessage)}</div>
  </main>
</body>
</html>`;
}

export function parseFirstRunSetupSubmissionUrl(rawUrl: string): FirstRunSetupSubmission | null {
  if (!rawUrl.startsWith('subminer://first-run-setup')) {
    return null;
  }
  const parsed = new URL(rawUrl);
  const action = parsed.searchParams.get('action');
  if (
    action !== 'configure-mpv-executable-path' &&
    action !== 'install-plugin' &&
    action !== 'configure-windows-mpv-shortcuts' &&
    action !== 'open-yomitan-settings' &&
    action !== 'refresh' &&
    action !== 'finish'
  ) {
    return null;
  }
  if (action === 'configure-mpv-executable-path') {
    return {
      action,
      mpvExecutablePath: parsed.searchParams.get('mpvExecutablePath') ?? '',
    };
  }
  if (action === 'configure-windows-mpv-shortcuts') {
    return {
      action,
      startMenuEnabled: parsed.searchParams.get('startMenu') === '1',
      desktopEnabled: parsed.searchParams.get('desktop') === '1',
    };
  }
  return { action };
}

export function createMaybeFocusExistingFirstRunSetupWindowHandler(deps: {
  getSetupWindow: () => FocusableWindowLike | null;
}) {
  return (): boolean => {
    const window = deps.getSetupWindow();
    if (!window) return false;
    window.focus();
    return true;
  };
}

export function createHandleFirstRunSetupNavigationHandler(deps: {
  parseSubmissionUrl: (rawUrl: string) => FirstRunSetupSubmission | null;
  handleAction: (submission: FirstRunSetupSubmission) => Promise<unknown>;
  logError: (message: string, error: unknown) => void;
}) {
  return (params: { url: string; preventDefault: () => void }): boolean => {
    if (!params.url.startsWith('subminer://first-run-setup')) {
      params.preventDefault();
      return true;
    }
    params.preventDefault();
    let submission: FirstRunSetupSubmission | null;
    try {
      submission = deps.parseSubmissionUrl(params.url);
    } catch {
      return true;
    }
    if (!submission) return true;
    void deps.handleAction(submission).catch((error) => {
      deps.logError('Failed handling first-run setup action', error);
    });
    return true;
  };
}

export function createOpenFirstRunSetupWindowHandler<
  TWindow extends FirstRunSetupWindowLike,
>(deps: {
  maybeFocusExistingSetupWindow: () => boolean;
  createSetupWindow: () => TWindow;
  getSetupSnapshot: () => Promise<FirstRunSetupHtmlModel>;
  buildSetupHtml: (model: FirstRunSetupHtmlModel) => string;
  parseSubmissionUrl: (rawUrl: string) => FirstRunSetupSubmission | null;
  handleAction: (submission: FirstRunSetupSubmission) => Promise<{ closeWindow?: boolean } | void>;
  markSetupInProgress: () => Promise<unknown>;
  markSetupCancelled: () => Promise<unknown>;
  isSetupCompleted: () => boolean;
  shouldQuitWhenClosedIncomplete: () => boolean;
  quitApp: () => void;
  clearSetupWindow: () => void;
  setSetupWindow: (window: TWindow) => void;
  encodeURIComponent: (value: string) => string;
  logError: (message: string, error: unknown) => void;
}) {
  return (): void => {
    if (deps.maybeFocusExistingSetupWindow()) {
      return;
    }

    const setupWindow = deps.createSetupWindow();
    deps.setSetupWindow(setupWindow);

    const render = async (): Promise<void> => {
      const model = await deps.getSetupSnapshot();
      const html = deps.buildSetupHtml(model);
      await setupWindow.loadURL(`data:text/html;charset=utf-8,${deps.encodeURIComponent(html)}`);
    };

    const handleNavigation = createHandleFirstRunSetupNavigationHandler({
      parseSubmissionUrl: deps.parseSubmissionUrl,
      handleAction: async (submission) => {
        const result = await deps.handleAction(submission);
        if (result?.closeWindow) {
          if (!setupWindow.isDestroyed()) {
            setupWindow.close();
          }
          return;
        }
        if (!setupWindow.isDestroyed()) {
          await render();
        }
      },
      logError: deps.logError,
    });

    setupWindow.webContents.on('will-navigate', (event, url) => {
      handleNavigation({
        url,
        preventDefault: () => {
          if (event && typeof event === 'object' && 'preventDefault' in event) {
            (event as { preventDefault?: () => void }).preventDefault?.();
          }
        },
      });
    });

    setupWindow.on('closed', () => {
      const setupCompleted = deps.isSetupCompleted();
      if (!setupCompleted) {
        void deps.markSetupCancelled().catch((error) => {
          deps.logError('Failed marking first-run setup cancelled', error);
        });
      }
      deps.clearSetupWindow();
      if (!setupCompleted && deps.shouldQuitWhenClosedIncomplete()) {
        deps.quitApp();
      }
    });

    void deps
      .markSetupInProgress()
      .then(() => render())
      .catch((error) => deps.logError('Failed opening first-run setup window', error));
  };
}
