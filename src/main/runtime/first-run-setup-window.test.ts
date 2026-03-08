import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildFirstRunSetupHtml,
  createHandleFirstRunSetupNavigationHandler,
  createMaybeFocusExistingFirstRunSetupWindowHandler,
  parseFirstRunSetupSubmissionUrl,
} from './first-run-setup-window';

test('buildFirstRunSetupHtml renders macchiato setup actions and disabled finish state', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 0,
    canFinish: false,
    pluginStatus: 'optional',
    pluginInstallPathSummary: null,
    message: 'Waiting for dictionaries',
  });

  assert.match(html, /SubMiner setup/);
  assert.match(html, /Install mpv plugin/);
  assert.match(html, /Open Yomitan Settings/);
  assert.match(html, /Finish setup/);
  assert.match(html, /disabled/);
});

test('buildFirstRunSetupHtml switches plugin action to reinstall when already installed', () => {
  const html = buildFirstRunSetupHtml({
    configReady: true,
    dictionaryCount: 1,
    canFinish: true,
    pluginStatus: 'installed',
    pluginInstallPathSummary: '/tmp/mpv',
    message: null,
  });

  assert.match(html, /Reinstall mpv plugin/);
});

test('parseFirstRunSetupSubmissionUrl parses supported custom actions', () => {
  assert.deepEqual(parseFirstRunSetupSubmissionUrl('subminer://first-run-setup?action=refresh'), {
    action: 'refresh',
  });
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
    handleAction: async (action) => {
      calls.push(action);
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
