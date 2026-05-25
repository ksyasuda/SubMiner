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

test('auto sync notifications never send desktop notifications', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('syncing', 'syncing'), {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('importing', 'importing'), {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('ready', 'ready'), {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('failed', 'failed'), {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, ['osd:syncing', 'osd:importing', 'osd:ready', 'osd:failed']);
});

test('auto sync notifications fall back to desktop for long progress when osd is unavailable', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('generating', 'generating'), {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
      return false;
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });
  notifyCharacterDictionaryAutoSyncStatus(makeEvent('ready', 'ready'), {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
      return false;
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, ['osd:generating', 'desktop:SubMiner:generating', 'osd:ready']);
});

test('auto sync notifications fall back to desktop when startup sequencer cannot show osd', () => {
  const calls: string[] = [];

  notifyCharacterDictionaryAutoSyncStatus(makeEvent('importing', 'importing'), {
    getNotificationType: () => 'both',
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

  assert.deepEqual(calls, [
    'sequencer:importing:importing',
    'desktop:SubMiner:importing',
  ]);
});
