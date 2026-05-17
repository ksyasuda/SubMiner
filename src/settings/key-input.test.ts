import test from 'node:test';
import assert from 'node:assert/strict';
import type { Keybinding } from '../types/runtime';
import {
  buildMpvKeybindingConfigValue,
  createMpvKeybindingRows,
  keyboardEventToConfigKey,
} from './key-input';

test('keyboardEventToConfigKey formats Electron accelerators from learned input', () => {
  assert.equal(
    keyboardEventToConfigKey(
      { code: 'KeyS', key: 's', ctrlKey: true, altKey: false, shiftKey: true, metaKey: false },
      'accelerator',
    ),
    'CommandOrControl+Shift+S',
  );
  assert.equal(
    keyboardEventToConfigKey(
      { code: 'Slash', key: '/', ctrlKey: false, altKey: true, shiftKey: false, metaKey: false },
      'accelerator',
    ),
    'Alt+Slash',
  );
});

test('keyboardEventToConfigKey formats DOM code bindings from learned input', () => {
  assert.equal(
    keyboardEventToConfigKey(
      { code: 'KeyJ', key: 'j', ctrlKey: true, altKey: false, shiftKey: true, metaKey: false },
      'dom-code',
    ),
    'Ctrl+Shift+KeyJ',
  );
  assert.equal(
    keyboardEventToConfigKey(
      {
        code: 'Backquote',
        key: '`',
        ctrlKey: false,
        altKey: false,
        shiftKey: false,
        metaKey: false,
      },
      'dom-code',
    ),
    'Backquote',
  );
});

test('keyboardEventToConfigKey formats bare key-code fields without modifiers', () => {
  assert.equal(
    keyboardEventToConfigKey(
      { code: 'KeyW', key: 'w', ctrlKey: true, altKey: true, shiftKey: false, metaKey: false },
      'code',
    ),
    'KeyW',
  );
});

test('MPV keybinding rows save default key moves as a disable plus replacement', () => {
  const defaults: Keybinding[] = [{ key: 'Space', command: ['cycle', 'pause'] }];
  const rows = createMpvKeybindingRows(defaults, []);
  rows[0]!.key = 'KeyP';

  assert.deepEqual(buildMpvKeybindingConfigValue(defaults, rows), [
    { key: 'Space', command: null },
    { key: 'KeyP', command: ['cycle', 'pause'] },
  ]);
});

test('MPV keybinding rows reopen moved default bindings as their default row', () => {
  const defaults: Keybinding[] = [{ key: 'Space', command: ['cycle', 'pause'] }];

  assert.deepEqual(
    createMpvKeybindingRows(defaults, [
      { key: 'Space', command: null },
      { key: 'KeyP', command: ['cycle', 'pause'] },
    ]),
    [
      {
        defaultKey: 'Space',
        key: 'KeyP',
        command: ['cycle', 'pause'],
        commandText: '["cycle","pause"]',
        isDefault: true,
      },
    ],
  );
});

test('MPV keybinding rows omit unchanged default bindings from config value', () => {
  const defaults: Keybinding[] = [
    { key: 'Space', command: ['cycle', 'pause'] },
    { key: 'KeyF', command: ['cycle', 'fullscreen'] },
  ];

  assert.deepEqual(
    buildMpvKeybindingConfigValue(defaults, createMpvKeybindingRows(defaults, [])),
    [],
  );
});
