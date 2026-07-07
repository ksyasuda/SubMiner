import { getFirstRunSetupCompletionMessage } from './first-run-setup-service';
import type {
  BunSnapshot,
  CommandLineLauncherSnapshot,
  LauncherSnapshot,
} from './command-line-launcher';
import { i18n } from '../../i18n/index.js';

type FocusableWindowLike = {
  focus: () => void;
  show?: () => void;
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
  | 'remove-legacy-plugin'
  | 'configure-windows-mpv-shortcuts'
  | 'install-bun'
  | 'install-command-line-launcher'
  | 'open-yomitan-settings'
  | 'open-config-settings'
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
  legacyMpvPluginPaths?: string[];
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
  commandLineLauncher: CommandLineLauncherSnapshot;
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

function formatCommand(command: string[] | null): string {
  return command?.join(' ') ?? i18n.t('setup.bun.meta.noCommand');
}

function getBunStatusLabel(status: BunSnapshot['status']): string {
  switch (status) {
    case 'ready':
      return i18n.t('setup.status.ready');
    case 'installing':
      return i18n.t('setup.status.installing');
    case 'failed':
      return i18n.t('setup.status.failed');
    case 'missing':
      return i18n.t('setup.status.missing');
  }
}

function getLauncherStatusLabel(status: LauncherSnapshot['status']): string {
  switch (status) {
    case 'ready':
      return i18n.t('setup.status.ready');
    case 'installed_bun_missing':
      return i18n.t('setup.status.installedBunMissing');
    case 'not_installed':
      return i18n.t('setup.status.notInstalled');
    case 'not_on_path':
      return i18n.t('setup.status.notOnPath');
    case 'shadowed':
      return i18n.t('setup.status.shadowed');
    case 'not_installable':
      return i18n.t('setup.status.notInstallable');
    case 'failed':
      return i18n.t('setup.status.failed');
  }
}

function getToolTone(status: BunSnapshot['status']): 'ready' | 'warn' | 'muted' | 'danger' {
  if (status === 'ready') return 'ready';
  if (status === 'failed') return 'danger';
  if (status === 'installing') return 'muted';
  return 'warn';
}

function getLauncherTone(
  status: LauncherSnapshot['status'],
): 'ready' | 'warn' | 'muted' | 'danger' {
  if (status === 'ready') return 'ready';
  if (status === 'failed') return 'danger';
  if (status === 'installed_bun_missing' || status === 'not_installed') return 'warn';
  return 'muted';
}

function renderCommandLineLauncherSection(
  commandLineLauncher: CommandLineLauncherSnapshot,
): string {
  if (!commandLineLauncher.supported) {
    return '';
  }

  const bun = commandLineLauncher.bun;
  const launcher = commandLineLauncher.launcher;
  const bunMeta =
    bun.status === 'ready'
      ? [
          bun.commandPath ? i18n.t('setup.bun.meta.path', { value: bun.commandPath }) : null,
          bun.version ? i18n.t('setup.bun.meta.version', { value: bun.version }) : null,
        ].filter(Boolean)
      : [
          bun.installMethod ? i18n.t('setup.bun.meta.method', { value: bun.installMethod }) : null,
          i18n.t('setup.bun.meta.command', { value: formatCommand(bun.installCommand) }),
          bun.message,
        ].filter(Boolean);
  const launcherMeta = [
    launcher.commandPath ? i18n.t('setup.bun.meta.command', { value: launcher.commandPath }) : null,
    launcher.installPath ? i18n.t('setup.bun.meta.installTarget', { value: launcher.installPath }) : null,
    launcher.pathDir ? i18n.t('setup.bun.meta.pathDir', { value: launcher.pathDir }) : null,
    launcher.shadowedBy ? i18n.t('setup.bun.meta.shadowedBy', { value: launcher.shadowedBy }) : null,
    launcher.message,
    bun.status !== 'ready' ? i18n.t('setup.bun.meta.warning') : null,
  ].filter(Boolean);
  const bunInstallButton =
    bun.status === 'missing' || bun.status === 'failed'
      ? `<button onclick="window.location.href='subminer://first-run-setup?action=install-bun'">${escapeHtml(i18n.t('setup.bun.installBun'))}</button>`
      : '';
  const launcherButtonDisabled = launcher.status === 'not_installable' ? 'disabled' : '';

  return `
    <section class="setup-section">
      <div class="section-head">
        <h2>${escapeHtml(i18n.t('setup.bunSection.title'))}</h2>
        <div class="meta">${escapeHtml(i18n.t('setup.bunSection.subtitle'))}</div>
      </div>
      <div class="card block">
        <div class="card-head">
          <div>
            <strong>${escapeHtml(i18n.t('setup.bun.bunRuntime'))}</strong>
            ${bunMeta.map((line) => `<div class="meta">${escapeHtml(String(line))}</div>`).join('')}
          </div>
          ${renderStatusBadge(getBunStatusLabel(bun.status), getToolTone(bun.status))}
        </div>
        <div class="inline-actions">
          ${bunInstallButton}
          <button class="ghost" onclick="window.location.href='subminer://first-run-setup?action=refresh'">${escapeHtml(i18n.t('setup.bun.refresh'))}</button>
        </div>
      </div>
      <div class="card block">
        <div class="card-head">
          <div>
            <strong>${escapeHtml(i18n.t('setup.bun.subminerLauncher'))}</strong>
            ${launcherMeta.map((line) => `<div class="meta">${escapeHtml(String(line))}</div>`).join('')}
          </div>
          ${renderStatusBadge(getLauncherStatusLabel(launcher.status), getLauncherTone(launcher.status))}
        </div>
        <div class="inline-actions">
          <button ${launcherButtonDisabled} onclick="window.location.href='subminer://first-run-setup?action=install-command-line-launcher'">${escapeHtml(i18n.t('setup.bun.installLauncher'))}</button>
          <button class="ghost" onclick="window.location.href='subminer://first-run-setup?action=refresh'">${escapeHtml(i18n.t('setup.bun.refresh'))}</button>
        </div>
      </div>
    </section>`;
}

export function buildFirstRunSetupHtml(model: FirstRunSetupHtmlModel): string {
  const legacyMpvPluginPaths = model.legacyMpvPluginPaths ?? [];
  const finishButtonLabel =
    legacyMpvPluginPaths.length > 0 && model.canFinish
      ? i18n.t('setup.action.continueWithoutRemoving')
      : i18n.t('setup.action.finish');
  const windowsShortcutLabel =
    model.windowsMpvShortcuts.status === 'installed'
      ? i18n.t('setup.status.installed')
      : model.windowsMpvShortcuts.status === 'skipped'
        ? i18n.t('setup.status.skipped')
        : model.windowsMpvShortcuts.status === 'failed'
          ? i18n.t('setup.status.failed')
          : i18n.t('setup.status.optional');
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
      ? i18n.t('setup.status.configured')
      : model.mpvExecutablePathStatus === 'invalid'
        ? i18n.t('setup.status.invalid')
        : i18n.t('setup.status.blank');
  const mpvExecutablePathTone =
    model.mpvExecutablePathStatus === 'configured'
      ? 'ready'
      : model.mpvExecutablePathStatus === 'invalid'
        ? 'danger'
        : 'muted';
  const mpvExecutablePathCurrent =
    model.mpvExecutablePathStatus === 'blank'
      ? i18n.t('setup.mpvExe.blankLabel')
      : model.mpvExecutablePathStatus === 'invalid'
        ? i18n.t('setup.mpvExe.invalidLabel', { path: model.mpvExecutablePath })
        : model.mpvExecutablePath;
  const mpvExecutablePathCard = model.windowsMpvShortcuts.supported
    ? `
    <div class="card block">
      <div class="card-head">
        <div>
          <strong>${escapeHtml(i18n.t('setup.mpvExe.title'))}</strong>
          <div class="meta">${escapeHtml(i18n.t('setup.mpvExe.leaveBlank'))}</div>
          <div class="meta">${escapeHtml(i18n.t('setup.mpvExe.current', { value: mpvExecutablePathCurrent }))}</div>
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
          aria-label="${escapeHtml(i18n.t('setup.mpvExe.ariaLabel'))}"
          value="${escapeHtml(model.mpvExecutablePath)}"
          placeholder="${escapeHtml(i18n.t('setup.mpvExe.placeholder'))}"
        />
        <button type="submit">${escapeHtml(i18n.t('setup.mpvExe.save'))}</button>
      </form>
    </div>`
    : '';
  const windowsShortcutCard = model.windowsMpvShortcuts.supported
    ? `
    <div class="card block">
      <div class="card-head">
        <div>
          <strong>${escapeHtml(i18n.t('setup.shortcuts.title'))}</strong>
          <div class="meta">${escapeHtml(i18n.t('setup.shortcuts.subtitle'))}</div>
          <div class="meta">${escapeHtml(i18n.t('setup.shortcuts.installedStatus', { startMenu: i18n.t(model.windowsMpvShortcuts.startMenuInstalled ? 'setup.shortcuts.yes' : 'setup.shortcuts.no'), desktop: i18n.t(model.windowsMpvShortcuts.desktopInstalled ? 'setup.shortcuts.yes' : 'setup.shortcuts.no') }))}</div>
        </div>
        ${renderStatusBadge(windowsShortcutLabel, windowsShortcutTone)}
      </div>
      <form
        class="shortcut-form"
        onsubmit="event.preventDefault(); const params = new URLSearchParams({ action: 'configure-windows-mpv-shortcuts', startMenu: document.getElementById('shortcut-start-menu').checked ? '1' : '0', desktop: document.getElementById('shortcut-desktop').checked ? '1' : '0' }); window.location.href = 'subminer://first-run-setup?' + params.toString();"
      >
        <label><input id="shortcut-start-menu" type="checkbox" ${model.windowsMpvShortcuts.startMenuEnabled ? 'checked' : ''} /> ${escapeHtml(i18n.t('setup.shortcuts.startMenuCheckbox'))}</label>
        <label><input id="shortcut-desktop" type="checkbox" ${model.windowsMpvShortcuts.desktopEnabled ? 'checked' : ''} /> ${escapeHtml(i18n.t('setup.shortcuts.desktopCheckbox'))}</label>
        <button type="submit">${escapeHtml(i18n.t('setup.shortcuts.apply'))}</button>
      </form>
    </div>`
    : '';
  const legacyPluginCard =
    legacyMpvPluginPaths.length > 0
      ? `
    <div class="card block">
      <div class="card-head">
        <div>
          <strong>${escapeHtml(i18n.t('setup.legacy.title'))}</strong>
          <div class="meta">${escapeHtml(i18n.t('setup.legacy.subtitle'))}</div>
        </div>
        ${renderStatusBadge(i18n.t('setup.status.found'), 'warn')}
      </div>
      <ul class="legacy-paths">
        ${legacyMpvPluginPaths.map((pluginPath) => `<li>${escapeHtml(pluginPath)}</li>`).join('')}
      </ul>
      <button class="legacy-remove" onclick="if (confirm(&quot;${escapeHtml(i18n.t('setup.legacy.confirmRemove'))}&quot;)) window.location.href='subminer://first-run-setup?action=remove-legacy-plugin'">${escapeHtml(i18n.t('setup.legacy.remove'))}</button>
    </div>`
      : '';

  const yomitanMeta = model.externalYomitanConfigured
    ? i18n.t('setup.yomitan.externalMeta')
    : i18n.t('setup.yomitan.countMeta', { count: model.dictionaryCount });
  const yomitanBadgeLabel = model.externalYomitanConfigured
    ? i18n.t('setup.status.external')
    : model.dictionaryCount >= 1
      ? i18n.t('setup.status.ready')
      : i18n.t('setup.status.missing');
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
        ? i18n.t('setup.footer.external')
        : i18n.t('setup.footer.unlocked')
      : i18n.t('setup.footer.locked');

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(i18n.t('setup.title'))}</title>
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
    html,
    body {
      min-height: 100%;
    }
    body {
      margin: 0;
      min-height: 100vh;
      background: linear-gradient(180deg, var(--mantle), var(--base));
      color: var(--text);
      font: 13px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    main {
      box-sizing: border-box;
      min-height: 100vh;
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
    .setup-section {
      margin-top: 10px;
    }
    .section-head {
      margin: 14px 0 8px;
    }
    .section-head h2 {
      margin: 0;
      font-size: 14px;
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
    .inline-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 12px;
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
    button.legacy-remove {
      display: inline-flex;
      justify-content: center;
      align-items: center;
      min-width: 220px;
      border: 1px solid rgba(237, 135, 150, 0.38);
      background: rgba(237, 135, 150, 0.14);
      color: #f5b1ba;
    }
    button.legacy-remove:hover {
      background: rgba(237, 135, 150, 0.22);
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
    .legacy-paths {
      margin: 10px 0 12px;
      padding-left: 18px;
      color: var(--muted);
      font-size: 12px;
      overflow-wrap: anywhere;
    }
  </style>
</head>
<body>
  <main>
    <h1>${escapeHtml(i18n.t('setup.heading'))}</h1>
    <div class="card">
      <div>
        <strong>${escapeHtml(i18n.t('setup.configFile.title'))}</strong>
        <div class="meta">${escapeHtml(i18n.t('setup.configFile.subtitle'))}</div>
      </div>
      ${renderStatusBadge(model.configReady ? i18n.t('setup.status.ready') : i18n.t('setup.status.missing'), model.configReady ? 'ready' : 'danger')}
    </div>
    <div class="card">
      <div>
        <strong>${escapeHtml(i18n.t('setup.yomitan.title'))}</strong>
        <div class="meta">${escapeHtml(yomitanMeta)}</div>
      </div>
      ${renderStatusBadge(yomitanBadgeLabel, yomitanBadgeTone)}
    </div>
    ${mpvExecutablePathCard}
    ${windowsShortcutCard}
    ${renderCommandLineLauncherSection(model.commandLineLauncher)}
    ${legacyPluginCard}
    <div class="actions">
      <button onclick="window.location.href='subminer://first-run-setup?action=open-yomitan-settings'">${escapeHtml(i18n.t('setup.action.openYomitan'))}</button>
      <button class="ghost" onclick="window.location.href='subminer://first-run-setup?action=refresh'">${escapeHtml(i18n.t('setup.action.refresh'))}</button>
      <button onclick="window.location.href='subminer://first-run-setup?action=open-config-settings'">${escapeHtml(i18n.t('setup.action.openSubminer'))}</button>
      <button class="primary" ${model.canFinish ? '' : 'disabled'} onclick="window.location.href='subminer://first-run-setup?action=finish'">${escapeHtml(finishButtonLabel)}</button>
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
    action !== 'remove-legacy-plugin' &&
    action !== 'configure-windows-mpv-shortcuts' &&
    action !== 'install-bun' &&
    action !== 'install-command-line-launcher' &&
    action !== 'open-yomitan-settings' &&
    action !== 'open-config-settings' &&
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
    window.show?.();
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
  handleAction: (
    submission: FirstRunSetupSubmission,
  ) => Promise<{ closeWindow?: boolean; skipRender?: boolean } | void>;
  markSetupInProgress: () => Promise<unknown>;
  markSetupCancelled: () => Promise<unknown>;
  isSetupCompleted: () => boolean;
  shouldQuitWhenClosedIncomplete: () => boolean;
  shouldQuitWhenClosedCompleted?: () => boolean;
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
    setupWindow.show?.();
    setupWindow.focus();

    const render = async (): Promise<void> => {
      const model = await deps.getSetupSnapshot();
      if (setupWindow.isDestroyed()) {
        return;
      }
      const html = deps.buildSetupHtml(model);
      if (setupWindow.isDestroyed()) {
        return;
      }
      await setupWindow.loadURL(`data:text/html;charset=utf-8,${deps.encodeURIComponent(html)}`);
      if (!setupWindow.isDestroyed()) {
        setupWindow.show?.();
        setupWindow.focus();
      }
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
        if (result?.skipRender) {
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
      if (
        (setupCompleted && deps.shouldQuitWhenClosedCompleted?.()) ||
        (!setupCompleted && deps.shouldQuitWhenClosedIncomplete())
      ) {
        deps.quitApp();
      }
    });

    void deps
      .markSetupInProgress()
      .then(() => render())
      .catch((error) => deps.logError('Failed opening first-run setup window', error));
  };
}
