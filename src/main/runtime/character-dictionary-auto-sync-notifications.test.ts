import assert from 'node:assert/strict';
import test from 'node:test';
import {
  notifyCharacterDictionaryAutoSyncStatus,
  type CharacterDictionaryAutoSyncNotificationEvent,
} from './character-dictionary-auto-sync-notifications';

function makeEvent(
  phase: CharacterDictionaryAutoSyncNotificationEvent['phase'],
  message: string,
): CharacterDictionaryAutoSyncNotificationEvent {
  return {
    phase,
    mediaId: 101291,
    mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
    message,
  };
}

test('auto sync notifications send osd updates for progress phases', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('checking', 'checking'), {
    getNotificationType: () => 'osd',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('generating', 'generating'), {
    getNotificationType: () => 'osd',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('syncing', 'syncing'), {
    getNotificationType: () => 'osd',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('importing', 'importing'), {
    getNotificationType: () => 'osd',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('ready', 'ready'), {
    getNotificationType: () => 'osd',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, [
    'osd:checking',
    'osd:generating',
    'osd:syncing',
    'osd:importing',
    'osd:ready',
  ]);
});

test('auto sync notifications route both to overlay and system only', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('syncing', 'syncing'), {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
    showOverlayNotification: (payload) =>
      calls.push(
        `overlay:${payload.id}:${payload.title}:${payload.body}:${payload.persistent ? 'pin' : 'auto'}`,
      ),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('ready', 'ready'), {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
    showOverlayNotification: (payload) =>
      calls.push(
        `overlay:${payload.id}:${payload.title}:${payload.body}:${payload.persistent ? 'pin' : 'auto'}`,
      ),
  });

  assert.deepEqual(calls, [
    'overlay:character-dictionary-auto-sync:Character dictionary:syncing:pin',
    'desktop:SubMiner:syncing',
    'overlay:character-dictionary-auto-sync:Character dictionary:ready:auto',
    'desktop:SubMiner:ready',
  ]);
});

test('auto sync notifications fall back to desktop when overlay routing is unavailable', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('building', 'building'), {
    getNotificationType: () => undefined,
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, ['desktop:SubMiner:building']);
});

test('auto sync notifications keep osd-system on legacy surfaces', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('syncing', 'syncing'), {
    getNotificationType: () => 'osd-system',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
    showOverlayNotification: (payload) => calls.push(`overlay:${payload.body}`),
  });

  assert.deepEqual(calls, ['osd:syncing', 'desktop:SubMiner:syncing']);
});

test('auto sync notifications keep osd-system desktop delivery even when osd is unavailable', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('generating', 'generating'), {
    getNotificationType: () => 'osd-system',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
      return false;
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('ready', 'ready'), {
    getNotificationType: () => 'osd-system',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
      return false;
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, [
    'osd:generating',
    'desktop:SubMiner:generating',
    'osd:ready',
    'desktop:SubMiner:ready',
  ]);
});

test('auto sync notifications send osd-system desktop updates with startup sequencer', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('importing', 'importing'), {
    getNotificationType: () => 'osd-system',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
    startupOsdSequencer: {
      notifyCharacterDictionaryStatus: (event) => {
        calls.push(`sequencer:${event.phase}:${event.message}`);
        return false;
      },
    },
  });

  assert.deepEqual(calls, ['sequencer:importing:importing', 'desktop:SubMiner:importing']);
});
