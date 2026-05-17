import test from 'node:test';
import assert from 'node:assert/strict';
import { createOpenConfigSettingsWindowHandler } from './config-settings-window';

test('createOpenConfigSettingsWindowHandler focuses existing settings window', () => {
  const calls: string[] = [];
  const existing = {
    isDestroyed: () => false,
    focus: () => calls.push('focus'),
    loadFile: () => calls.push('load'),
    on: () => {},
  };

  const open = createOpenConfigSettingsWindowHandler({
    getSettingsWindow: () => existing,
    setSettingsWindow: () => calls.push('set'),
    createSettingsWindow: () => {
      throw new Error('Should not create a second window.');
    },
    settingsHtmlPath: '/tmp/settings.html',
  });

  assert.equal(open(), true);
  assert.deepEqual(calls, ['focus']);
});

test('createOpenConfigSettingsWindowHandler creates window and clears closed state', () => {
  const calls: string[] = [];
  const handlers: { closed?: () => void } = {};
  const created = {
    isDestroyed: () => false,
    focus: () => calls.push('focus'),
    loadFile: (path: string) => calls.push(`load:${path}`),
    on: (event: string, handler: () => void) => {
      if (event === 'closed') handlers.closed = handler;
    },
  };

  const open = createOpenConfigSettingsWindowHandler({
    getSettingsWindow: () => null,
    setSettingsWindow: (window) => calls.push(window ? 'set:window' : 'set:null'),
    createSettingsWindow: () => created,
    settingsHtmlPath: '/tmp/settings.html',
  });

  assert.equal(open(), true);
  assert.deepEqual(calls, ['load:/tmp/settings.html', 'set:window', 'focus']);
  assert.ok(handlers.closed);
  handlers.closed();
  assert.equal(calls.at(-1), 'set:null');
});
