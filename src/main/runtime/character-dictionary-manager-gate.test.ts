import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import {
  CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE,
  openCharacterDictionaryManagerWithConfigGate,
  type CharacterDictionaryManagerNotificationType,
} from './character-dictionary-manager-gate';

function makeDeps(options: {
  enabled?: boolean;
  notificationType?: CharacterDictionaryManagerNotificationType;
}) {
  const calls: string[] = [];
  return {
    calls,
    deps: {
      isCharacterDictionaryEnabled: () => options.enabled ?? false,
      getNotificationType: () => options.notificationType ?? 'osd',
      openManager: () => calls.push('open'),
      showOsd: (message: string) => calls.push(`osd:${message}`),
      showOverlayNotification: (payload: { title: string; body?: string }) =>
        calls.push(`overlay:${payload.title}:${payload.body ?? ''}`),
      showDesktopNotification: (title: string, opts: { body: string }) =>
        calls.push(`system:${title}:${opts.body}`),
      logWarn: (message: string) => calls.push(`warn:${message}`),
    },
  };
}

test('opens character dictionary manager when character dictionary is enabled', () => {
  const { calls, deps } = makeDeps({ enabled: true, notificationType: 'both' });

  openCharacterDictionaryManagerWithConfigGate(deps);

  assert.deepEqual(calls, ['open']);
});

test('routes disabled manager notification to configured surfaces', () => {
  for (const [type, expected] of [
    ['osd', [`osd:${CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE}`]],
    ['system', [`system:SubMiner:${CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE}`]],
    [
      'both',
      [
        `overlay:SubMiner:${CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE}`,
        `system:SubMiner:${CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE}`,
      ],
    ],
    [
      'osd-system',
      [
        `osd:${CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE}`,
        `system:SubMiner:${CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE}`,
      ],
    ],
    ['none', []],
  ] as const) {
    const { calls, deps } = makeDeps({ enabled: false, notificationType: type });

    openCharacterDictionaryManagerWithConfigGate(deps);

    assert.deepEqual(calls, expected);
  }
});
