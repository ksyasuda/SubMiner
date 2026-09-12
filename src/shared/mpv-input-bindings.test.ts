import assert from 'node:assert/strict';
import test from 'node:test';
import {
  keyboardEventToMpvKey,
  normalizeMpvInputKey,
  parseMpvInputBindingKeys,
} from './mpv-input-bindings';

test('mpv discovery validates entries and excludes inactive, mouse, sequence, and SubMiner keys', () => {
  assert.deepEqual(
    parseMpvInputBindingKeys([
      { key: 'r', cmd: 'script-binding replay/run', priority: 1, owner: 'replay' },
      { key: 'r', cmd: 'show-text duplicate', priority: 0 },
      { key: 'Ctrl+A', cmd: 'show-text shifted', priority: 1 },
      { key: 'g-g', cmd: 'seek 0', priority: 1 },
      { key: 'MBTN_LEFT', cmd: 'cycle pause', priority: 1 },
      { key: 'q', cmd: 'quit', priority: -1 },
      { key: 's', cmd: 'screenshot', priority: 1 },
      { key: 's', cmd: 'script-binding subminer/session', priority: 5, owner: 'subminer' },
      { key: 't', cmd: 'script-message subminer-toggle', priority: 1 },
      { key: 'x', cmd: 'ignore', priority: NaN },
      { key: 'z', cmd: 5, priority: 1 },
      null,
    ]),
    ['r', 'ctrl+A'],
  );
  assert.deepEqual(parseMpvInputBindingKeys({ key: 'r' }), []);
});

test('mpv keys retain printable characters and normalize modifiers', () => {
  assert.equal(normalizeMpvInputKey('Alt+Ctrl+Shift+a'), 'ctrl+alt+A');
  assert.equal(normalizeMpvInputKey('Ctrl++'), 'ctrl++');
  assert.equal(normalizeMpvInputKey('Shift+LEFT'), 'shift+LEFT');
  assert.equal(normalizeMpvInputKey('F12'), 'F12');
  assert.equal(normalizeMpvInputKey('UNMAPPED'), null);
});

test('keyboard conversion respects layout characters and skips composition and AltGr', () => {
  const event = {
    key: 'A',
    ctrlKey: true,
    shiftKey: true,
    altKey: false,
    metaKey: false,
    isComposing: false,
    getModifierState: () => false,
  };
  assert.equal(keyboardEventToMpvKey(event), 'ctrl+A');
  assert.equal(keyboardEventToMpvKey({ ...event, key: '#', ctrlKey: false }), 'SHARP');
  assert.equal(keyboardEventToMpvKey({ ...event, key: 'ArrowLeft', ctrlKey: false }), 'shift+LEFT');
  assert.equal(keyboardEventToMpvKey({ ...event, key: 'Dead' }), null);
  assert.equal(keyboardEventToMpvKey({ ...event, isComposing: true }), null);
  assert.equal(
    keyboardEventToMpvKey({ ...event, getModifierState: (key) => key === 'AltGraph' }),
    null,
  );
});

test('only the highest-priority active binding determines SubMiner ownership', () => {
  const user = { key: 'r', cmd: 'script-binding replay/run', is_weak: false, priority: 12 };
  const plugin = {
    key: 'r',
    cmd: 'script-binding subminer/run',
    owner: 'subminer',
    is_weak: true,
    priority: 2,
  };
  assert.deepEqual(parseMpvInputBindingKeys([plugin, user]), ['r']);
  assert.deepEqual(parseMpvInputBindingKeys([user, plugin]), ['r']);
  assert.deepEqual(parseMpvInputBindingKeys([user, { ...plugin, priority: -1 }]), ['r']);
  assert.deepEqual(
    parseMpvInputBindingKeys([user, { ...plugin, is_weak: false, priority: 15 }]),
    [],
  );
});

test('SubMiner ownership excludes only its script commands and respects explicit owners', () => {
  assert.deepEqual(
    parseMpvInputBindingKeys([
      { key: 'a', cmd: 'show-text "subminer/readme"', priority: 1 },
      { key: 'b', cmd: 'run subminer-helper', priority: 1 },
      { key: 'c', cmd: 'script-message subminer-toggle', owner: 'other-script', priority: 1 },
      { key: 'd', cmd: 'script-binding subminer/action', owner: 'other-script', priority: 1 },
      { key: 'e', cmd: '  script-binding "subminer/action"', priority: 1 },
      { key: 'f', cmd: 'script-message subminer-toggle', priority: 1 },
      { key: 'g', cmd: 'ignore', owner: 'subminer', priority: 1 },
    ]),
    ['a', 'b', 'c', 'd'],
  );
});

test('SubMiner ownership recognizes leading mpv prefixes without matching command arguments', () => {
  assert.deepEqual(
    parseMpvInputBindingKeys([
      { key: 'a', cmd: 'no-osd script-binding subminer/action', priority: 1 },
      { key: 'b', cmd: '  repeatable\tasync raw script-binding "subminer/action"', priority: 1 },
      { key: 'c', cmd: 'osd-msg-bar sync script-message subminer-toggle', priority: 1 },
      { key: 'd', cmd: 'no-osd show-text "script-binding subminer/action"', priority: 1 },
      { key: 'e', cmd: 'show-text "no-osd script-binding subminer/action"', priority: 1 },
      { key: 'f', cmd: 'no-osd script-binding other/action', priority: 1 },
      { key: 'g', cmd: 'no-osd script-binding subminer/action', owner: 'other', priority: 1 },
    ]),
    ['d', 'e', 'f', 'g'],
  );
});
