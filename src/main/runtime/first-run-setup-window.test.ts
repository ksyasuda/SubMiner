import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildFirstRunSetupHtml,
  createHandleFirstRunSetupNavigationHandler,
  createMaybeFocusExistingFirstRunSetupWindowHandler,
  createOpenFirstRunSetupWindowHandler,
  parseFirstRunSetupSubmissionUrl,
} from './first-run-setup-window';
import type { CommandLineLauncherSnapshot } from './command-line-launcher';

function createCommandLineLauncherSnapshot(
  overrides: Partial<CommandLineLauncherSnapshot> = {},
): CommandLineLauncherSnapshot {
  return {
    supported: true,
    bun: {
      status: 'missing',
      commandPath: null,
      version: null,
      installMethod: 'official-script',
      installCommand: ['bash', '-lc', 'curl -fsSL https://bun.com/install | bash'],
      message: null,
    },
    launcher: {
      status: 'not_installed',
      commandPath: null,
      installPath: '/home/tester/.local/bin/subminer',
      pathDir: '/home/tester/.local/bin',
      shadowedBy: null,
      message: null,
    },
    ...overrides,
  };
}

test('buildFirstRunSetupHtml renders macchiato setup actions and disabled finish state', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 0,
    canFinish: false,
    externalYomitanConfigured: false,
    pluginStatus: 'required',
    pluginInstallPathSummary: null,
    mpvExecutablePath: '',
    mpvExecutablePathStatus: 'blank',
    windowsMpvShortcuts: {
      supported: false,
      startMenuEnabled: true,
      desktopEnabled: true,
      startMenuInstalled: false,
      desktopInstalled: false,
      status: 'optional',
    },
    commandLineLauncher: createCommandLineLauncherSnapshot(),
    message: 'Waiting for dictionaries',
  });

  assert.match(html, /SubMiner setup/);
  assert.doesNotMatch(html, /Install legacy mpv plugin/);
  assert.doesNotMatch(html, /action=install-plugin/);
  assert.match(html, /Ready/);
  assert.doesNotMatch(html, /Bundled ready/);
  assert.match(html, /Managed mpv launches use the bundled runtime plugin\./);
  assert.match(html, /Open Yomitan Settings/);
  assert.match(html, /Finish setup/);
  assert.match(html, /disabled/);
});

test('buildFirstRunSetupHtml switches plugin action to reinstall when already installed', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 1,
    canFinish: true,
    externalYomitanConfigured: false,
    pluginStatus: 'installed',
    pluginInstallPathSummary: '/tmp/mpv',
    mpvExecutablePath: 'C:\\Program Files\\mpv\\mpv.exe',
    mpvExecutablePathStatus: 'configured',
    windowsMpvShortcuts: {
      supported: true,
      startMenuEnabled: true,
      desktopEnabled: true,
      startMenuInstalled: true,
      desktopInstalled: false,
      status: 'installed',
    },
    commandLineLauncher: createCommandLineLauncherSnapshot(),
    message: null,
  });

  assert.doesNotMatch(html, /Reinstall mpv plugin/);
  assert.doesNotMatch(html, /action=install-plugin/);
  assert.match(html, /mpv executable path/);
  assert.match(html, /Leave blank to auto-discover mpv\.exe from PATH\./);
  assert.match(html, /aria-label="Path to mpv\.exe"/);
  assert.match(html, /SubMiner-managed mpv launches use the bundled runtime plugin\./);
});

test('buildFirstRunSetupHtml shows legacy mpv plugin removal action with confirmation', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 1,
    canFinish: true,
    externalYomitanConfigured: false,
    pluginStatus: 'installed',
    pluginInstallPathSummary: '/tmp/mpv',
    legacyMpvPluginPaths: ['/tmp/mpv/scripts/subminer', '/tmp/mpv/scripts/subminer.lua'],
    mpvExecutablePath: '',
    mpvExecutablePathStatus: 'blank',
    windowsMpvShortcuts: {
      supported: false,
      startMenuEnabled: true,
      desktopEnabled: true,
      startMenuInstalled: false,
      desktopInstalled: false,
      status: 'optional',
    },
    commandLineLauncher: createCommandLineLauncherSnapshot(),
    message: null,
  });

  assert.match(html, /Legacy mpv plugin/);
  assert.match(html, /Legacy detected/);
  assert.match(html, /\/tmp\/mpv\/scripts\/subminer/);
  assert.match(html, /\/tmp\/mpv\/scripts\/subminer\.lua/);
  assert.match(html, /Remove legacy mpv plugin/);
  assert.match(html, /class="legacy-remove"/);
  assert.match(html, /\.legacy-remove/);
  assert.match(html, /Continue without removing/);
  assert.match(
    html,
    /Remove these SubMiner mpv plugin files from mpv.s scripts directory\? This stops regular mpv from loading SubMiner\./,
  );
  assert.match(html, /action=remove-legacy-plugin/);
});

test('buildFirstRunSetupHtml marks an invalid configured mpv path as invalid', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 1,
    canFinish: true,
    externalYomitanConfigured: false,
    pluginStatus: 'installed',
    pluginInstallPathSummary: '/tmp/mpv',
    mpvExecutablePath: 'C:\\Broken\\mpv.exe',
    mpvExecutablePathStatus: 'invalid',
    windowsMpvShortcuts: {
      supported: true,
      startMenuEnabled: true,
      desktopEnabled: true,
      startMenuInstalled: false,
      desktopInstalled: false,
      status: 'optional',
    },
    commandLineLauncher: createCommandLineLauncherSnapshot(),
    message: null,
  });

  assert.match(html, />Invalid</);
  assert.match(html, /Current: C:\\Broken\\mpv\.exe \(invalid; file not found\)/);
});

test('buildFirstRunSetupHtml explains the config blocker when setup is missing config', () => {
  const html = buildFirstRunSetupHtml({
    configReady: false,
    dictionaryCount: 0,
    canFinish: false,
    externalYomitanConfigured: false,
    pluginStatus: 'required',
    pluginInstallPathSummary: null,
    mpvExecutablePath: '',
    mpvExecutablePathStatus: 'blank',
    windowsMpvShortcuts: {
      supported: false,
      startMenuEnabled: true,
      desktopEnabled: true,
      startMenuInstalled: false,
      desktopInstalled: false,
      status: 'optional',
    },
    commandLineLauncher: createCommandLineLauncherSnapshot(),
    message: null,
  });

  assert.match(html, /Create or provide the config file before finishing setup\./);
});

test('buildFirstRunSetupHtml explains external yomitan mode and keeps finish enabled', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 0,
    canFinish: true,
    externalYomitanConfigured: true,
    pluginStatus: 'installed',
    pluginInstallPathSummary: null,
    mpvExecutablePath: '',
    mpvExecutablePathStatus: 'blank',
    windowsMpvShortcuts: {
      supported: false,
      startMenuEnabled: true,
      desktopEnabled: true,
      startMenuInstalled: false,
      desktopInstalled: false,
      status: 'optional',
    },
    commandLineLauncher: createCommandLineLauncherSnapshot(),
    message: null,
  });

  assert.match(html, /External profile configured/);
  assert.match(
    html,
    /Finish stays unlocked while SubMiner is reusing an external Yomitan profile\./,
  );
});

test('parseFirstRunSetupSubmissionUrl parses supported custom actions', () => {
  assert.deepEqual(
    parseFirstRunSetupSubmissionUrl(
      'subminer://first-run-setup?action=configure-mpv-executable-path&mpvExecutablePath=C%3A%5CApps%5Cmpv%5Cmpv.exe',
    ),
    {
      action: 'configure-mpv-executable-path',
      mpvExecutablePath: 'C:\\Apps\\mpv\\mpv.exe',
    },
  );
  assert.deepEqual(parseFirstRunSetupSubmissionUrl('subminer://first-run-setup?action=refresh'), {
    action: 'refresh',
  });
  assert.deepEqual(
    parseFirstRunSetupSubmissionUrl('subminer://first-run-setup?action=install-bun'),
    {
      action: 'install-bun',
    },
  );
  assert.deepEqual(
    parseFirstRunSetupSubmissionUrl(
      'subminer://first-run-setup?action=install-command-line-launcher',
    ),
    {
      action: 'install-command-line-launcher',
    },
  );
  assert.deepEqual(
    parseFirstRunSetupSubmissionUrl('subminer://first-run-setup?action=remove-legacy-plugin'),
    {
      action: 'remove-legacy-plugin',
    },
  );
  assert.equal(
    parseFirstRunSetupSubmissionUrl('subminer://first-run-setup?action=skip-plugin'),
    null,
  );
  assert.equal(parseFirstRunSetupSubmissionUrl('https://example.com'), null);
});

test('buildFirstRunSetupHtml renders command-line launcher section and actions', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 1,
    canFinish: true,
    externalYomitanConfigured: false,
    pluginStatus: 'installed',
    pluginInstallPathSummary: null,
    mpvExecutablePath: '',
    mpvExecutablePathStatus: 'blank',
    windowsMpvShortcuts: {
      supported: false,
      startMenuEnabled: true,
      desktopEnabled: true,
      startMenuInstalled: false,
      desktopInstalled: false,
      status: 'optional',
    },
    commandLineLauncher: createCommandLineLauncherSnapshot({
      bun: {
        status: 'failed',
        commandPath: null,
        version: null,
        installMethod: 'official-script',
        installCommand: ['bash', '-lc', 'curl -fsSL https://bun.com/install | bash'],
        message: 'network failed',
      },
      launcher: {
        status: 'installed_bun_missing',
        commandPath: '/home/tester/.local/bin/subminer',
        installPath: '/home/tester/.local/bin/subminer',
        pathDir: '/home/tester/.local/bin',
        shadowedBy: null,
        message: 'Bun is missing.',
      },
    }),
    message: null,
  });

  assert.match(html, /Command line launcher/);
  assert.match(html, /Optional\. Setup can finish without Bun or the launcher\./);
  assert.match(html, /Bun runtime/);
  assert.match(html, /Failed/);
  assert.match(html, /bash -lc curl -fsSL https:\/\/bun\.com\/install \| bash/);
  assert.match(html, /Install Bun/);
  assert.match(html, /action=install-bun/);
  assert.match(html, /SubMiner launcher/);
  assert.match(html, /Installed, Bun missing/);
  assert.match(html, /\/home\/tester\/\.local\/bin\/subminer/);
  assert.match(html, /action=install-command-line-launcher/);
  assert.match(
    html,
    /<button class="primary"  onclick="window\.location\.href='subminer:\/\/first-run-setup\?action=finish'">Finish setup<\/button>/,
  );
});

test('buildFirstRunSetupHtml disables launcher install when no target is installable', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 1,
    canFinish: true,
    externalYomitanConfigured: false,
    pluginStatus: 'installed',
    pluginInstallPathSummary: null,
    mpvExecutablePath: '',
    mpvExecutablePathStatus: 'blank',
    windowsMpvShortcuts: {
      supported: false,
      startMenuEnabled: true,
      desktopEnabled: true,
      startMenuInstalled: false,
      desktopInstalled: false,
      status: 'optional',
    },
    commandLineLauncher: createCommandLineLauncherSnapshot({
      launcher: {
        status: 'not_installable',
        commandPath: null,
        installPath: null,
        pathDir: null,
        shadowedBy: null,
        message: 'No writable PATH directory found.',
      },
    }),
    message: null,
  });

  assert.match(
    html,
    /<button disabled onclick="window\.location\.href='subminer:\/\/first-run-setup\?action=install-command-line-launcher'">Install launcher<\/button>/,
  );
});

test('first-run setup window handler focuses existing window', () => {
  const calls: string[] = [];
  const maybeFocus = createMaybeFocusExistingFirstRunSetupWindowHandler({
    getSetupWindow: () => ({
      focus: () => calls.push('focus'),
    }),
  });

  assert.equal(maybeFocus(), true);
  assert.deepEqual(calls, ['focus']);
});

test('first-run setup navigation handler prevents default and dispatches supported action', async () => {
  const calls: string[] = [];
  const handleNavigation = createHandleFirstRunSetupNavigationHandler({
    parseSubmissionUrl: (url) => parseFirstRunSetupSubmissionUrl(url),
    handleAction: async (submission) => {
      calls.push(submission.action);
    },
    logError: (message) => calls.push(message),
  });

  const prevented = handleNavigation({
    url: 'subminer://first-run-setup?action=refresh',
    preventDefault: () => calls.push('preventDefault'),
  });

  assert.equal(prevented, true);
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(calls, ['preventDefault', 'refresh']);
});

test('first-run setup parser rejects legacy global plugin install action', () => {
  assert.equal(
    parseFirstRunSetupSubmissionUrl('subminer://first-run-setup?action=install-plugin'),
    null,
  );
});

test('first-run setup navigation handler swallows stale custom-scheme actions', () => {
  const calls: string[] = [];
  const handleNavigation = createHandleFirstRunSetupNavigationHandler({
    parseSubmissionUrl: (url) => parseFirstRunSetupSubmissionUrl(url),
    handleAction: async (submission) => {
      calls.push(submission.action);
    },
    logError: (message) => calls.push(message),
  });

  const prevented = handleNavigation({
    url: 'subminer://first-run-setup?action=skip-plugin',
    preventDefault: () => calls.push('preventDefault'),
  });

  assert.equal(prevented, true);
  assert.deepEqual(calls, ['preventDefault']);
});

test('closing incomplete first-run setup quits app outside background mode', async () => {
  const calls: string[] = [];
  let closedHandler: (() => void) | undefined;
  const handler = createOpenFirstRunSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => false,
    createSetupWindow: () =>
      ({
        webContents: {
          on: () => {},
        },
        loadURL: async () => undefined,
        on: (event: 'closed', callback: () => void) => {
          if (event === 'closed') {
            closedHandler = callback;
          }
        },
        isDestroyed: () => false,
        close: () => calls.push('close-window'),
        focus: () => {},
      }) as never,
    getSetupSnapshot: async () => ({
      configReady: false,
      dictionaryCount: 0,
      canFinish: false,
      externalYomitanConfigured: false,
      pluginStatus: 'required',
      pluginInstallPathSummary: null,
      mpvExecutablePath: '',
      mpvExecutablePathStatus: 'blank',
      windowsMpvShortcuts: {
        supported: false,
        startMenuEnabled: true,
        desktopEnabled: true,
        startMenuInstalled: false,
        desktopInstalled: false,
        status: 'optional',
      },
      commandLineLauncher: createCommandLineLauncherSnapshot(),
      message: null,
    }),
    buildSetupHtml: () => '<html></html>',
    parseSubmissionUrl: () => null,
    handleAction: async () => undefined,
    markSetupInProgress: async () => undefined,
    markSetupCancelled: async () => {
      calls.push('cancelled');
    },
    isSetupCompleted: () => false,
    shouldQuitWhenClosedIncomplete: () => true,
    quitApp: () => {
      calls.push('quit');
    },
    clearSetupWindow: () => {
      calls.push('clear');
    },
    setSetupWindow: () => {
      calls.push('set');
    },
    encodeURIComponent: (value) => value,
    logError: () => {},
  });

  handler();
  if (typeof closedHandler !== 'function') {
    throw new Error('expected closed handler');
  }
  closedHandler();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(calls, ['set', 'cancelled', 'clear', 'quit']);
});

test('closing completed first-run setup quits app when completion policy allows it', async () => {
  const calls: string[] = [];
  let closedHandler: (() => void) | undefined;
  const handler = createOpenFirstRunSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => false,
    createSetupWindow: () =>
      ({
        webContents: {
          on: () => {},
        },
        loadURL: async () => undefined,
        on: (event: 'closed', callback: () => void) => {
          if (event === 'closed') {
            closedHandler = callback;
          }
        },
        isDestroyed: () => false,
        close: () => calls.push('close-window'),
        focus: () => {},
      }) as never,
    getSetupSnapshot: async () => ({
      configReady: true,
      dictionaryCount: 1,
      canFinish: true,
      externalYomitanConfigured: false,
      pluginStatus: 'installed',
      pluginInstallPathSummary: null,
      mpvExecutablePath: '',
      mpvExecutablePathStatus: 'blank',
      windowsMpvShortcuts: {
        supported: false,
        startMenuEnabled: true,
        desktopEnabled: true,
        startMenuInstalled: false,
        desktopInstalled: false,
        status: 'optional',
      },
      commandLineLauncher: createCommandLineLauncherSnapshot(),
      message: null,
    }),
    buildSetupHtml: () => '<html></html>',
    parseSubmissionUrl: () => null,
    handleAction: async () => undefined,
    markSetupInProgress: async () => undefined,
    markSetupCancelled: async () => {
      calls.push('cancelled');
    },
    isSetupCompleted: () => true,
    shouldQuitWhenClosedIncomplete: () => true,
    shouldQuitWhenClosedCompleted: () => true,
    quitApp: () => {
      calls.push('quit');
    },
    clearSetupWindow: () => {
      calls.push('clear');
    },
    setSetupWindow: () => {
      calls.push('set');
    },
    encodeURIComponent: (value) => value,
    logError: () => {},
  });

  handler();
  if (typeof closedHandler !== 'function') {
    throw new Error('expected closed handler');
  }
  closedHandler();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(calls, ['set', 'clear', 'quit']);
});
