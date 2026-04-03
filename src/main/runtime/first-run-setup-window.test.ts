import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildFirstRunSetupHtml,
  createHandleFirstRunSetupNavigationHandler,
  createMaybeFocusExistingFirstRunSetupWindowHandler,
  createOpenFirstRunSetupWindowHandler,
  parseFirstRunSetupSubmissionUrl,
} from './first-run-setup-window';

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
    message: 'Waiting for dictionaries',
  });

  assert.match(html, /SubMiner setup/);
  assert.match(html, /Install mpv plugin/);
  assert.match(html, /Required before SubMiner setup can finish/);
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
    message: null,
  });

  assert.match(html, /Reinstall mpv plugin/);
  assert.match(html, /mpv executable path/);
  assert.match(html, /Leave blank to auto-discover mpv\.exe from PATH\./);
  assert.match(html, /aria-label="Path to mpv\.exe"/);
  assert.match(
    html,
    /Finish stays unlocked once the mpv plugin is installed and Yomitan reports at least one installed dictionary\./,
  );
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
    message: null,
  });

  assert.match(html, /External profile configured/);
  assert.match(html, /Finish stays unlocked while SubMiner is reusing an external Yomitan profile\./);
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
  assert.equal(parseFirstRunSetupSubmissionUrl('subminer://first-run-setup?action=skip-plugin'), null);
  assert.equal(parseFirstRunSetupSubmissionUrl('https://example.com'), null);
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

test('first-run setup navigation handler prevents default and dispatches action', async () => {
  const calls: string[] = [];
  const handleNavigation = createHandleFirstRunSetupNavigationHandler({
    parseSubmissionUrl: (url) => parseFirstRunSetupSubmissionUrl(url),
    handleAction: async (submission) => {
      calls.push(submission.action);
    },
    logError: (message) => calls.push(message),
  });

  const prevented = handleNavigation({
    url: 'subminer://first-run-setup?action=install-plugin',
    preventDefault: () => calls.push('preventDefault'),
  });

  assert.equal(prevented, true);
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(calls, ['preventDefault', 'install-plugin']);
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
