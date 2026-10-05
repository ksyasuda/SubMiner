import assert from 'node:assert/strict';
import test from 'node:test';
import {
  confirmBucketDelete,
  confirmDayGroupDelete,
  confirmEpisodeDelete,
  confirmSessionDelete,
  setDeleteConfirmPresenter,
} from './delete-confirm';

interface NativeStatsApi {
  confirmNativeDialog?: (message: string) => boolean;
  beginNativeDialog?: () => void;
  endNativeDialog?: () => void;
}

type ElectronGlobal = typeof globalThis & { electronAPI?: { stats?: NativeStatsApi } };

/**
 * Runs `run` with a fake browser `confirm` (answering `confirmAnswer`) and, when
 * `electronStats` is given, a fake Electron stats bridge built from the shared `calls`
 * log. Calls into either end up in `calls`, in order.
 */
async function withConfirmGlobals(
  options: {
    confirmAnswer?: boolean;
    electronStats?: (calls: string[]) => NativeStatsApi;
  },
  run: (calls: string[]) => Promise<void>,
): Promise<void> {
  const calls: string[] = [];
  const globals = globalThis as ElectronGlobal;
  const originalConfirm = globals.confirm;
  const originalElectronAPI = globals.electronAPI;

  globals.confirm = ((message?: string) => {
    calls.push(`browser-confirm:${message ?? ''}`);
    return options.confirmAnswer ?? true;
  }) as typeof globalThis.confirm;
  globals.electronAPI = options.electronStats ? { stats: options.electronStats(calls) } : undefined;

  try {
    await run(calls);
  } finally {
    globals.confirm = originalConfirm;
    globals.electronAPI = originalElectronAPI;
  }
}

const SESSION_WARNING = 'Delete this session and all associated data?';

test('confirmSessionDelete falls back to the browser confirm with the shared warning copy', async () => {
  await withConfirmGlobals({}, async (calls) => {
    assert.equal(await confirmSessionDelete(), true);
    assert.deepEqual(calls, [`browser-confirm:${SESSION_WARNING}`]);
  });
});

test('confirmSessionDelete suspends stats overlay layering around native confirm', async () => {
  await withConfirmGlobals(
    {
      electronStats: (calls) => ({
        beginNativeDialog: () => calls.push('begin-native-dialog'),
        endNativeDialog: () => calls.push('end-native-dialog'),
      }),
    },
    async (calls) => {
      assert.equal(await confirmSessionDelete(), true);
      assert.deepEqual(calls, [
        'begin-native-dialog',
        `browser-confirm:${SESSION_WARNING}`,
        'end-native-dialog',
      ]);
    },
  );
});

test('confirmSessionDelete uses parented Electron confirm when available', async () => {
  await withConfirmGlobals(
    {
      electronStats: (calls) => ({
        confirmNativeDialog: (message) => {
          calls.push(`native-confirm:${message}`);
          return false;
        },
        beginNativeDialog: () => calls.push('begin-native-dialog'),
        endNativeDialog: () => calls.push('end-native-dialog'),
      }),
    },
    async (calls) => {
      assert.equal(await confirmSessionDelete(), false);
      assert.deepEqual(calls, [`native-confirm:${SESSION_WARNING}`]);
    },
  );
});

test('confirmSessionDelete uses the registered stats presenter before native or browser confirm', async () => {
  await withConfirmGlobals(
    {
      electronStats: (calls) => ({
        confirmNativeDialog: (message) => {
          calls.push(`native-confirm:${message}`);
          return true;
        },
      }),
    },
    async (calls) => {
      const unregister = setDeleteConfirmPresenter(async (message) => {
        calls.push(`presenter:${message}`);
        return false;
      });

      try {
        assert.equal(await confirmSessionDelete(), false);
        assert.deepEqual(calls, [`presenter:${SESSION_WARNING}`]);
      } finally {
        unregister();
      }
    },
  );
});

const copyCases = [
  {
    name: 'confirmDayGroupDelete includes the day label and count',
    ask: () => confirmDayGroupDelete('Today', 3),
    message: 'Delete all 3 sessions from Today and all associated data?',
  },
  {
    name: 'confirmDayGroupDelete uses singular for one session',
    ask: () => confirmDayGroupDelete('Yesterday', 1),
    message: 'Delete this session from Yesterday and all associated data?',
  },
  {
    name: 'confirmBucketDelete names the episode and count for multiple sessions',
    ask: () => confirmBucketDelete('My Episode', 3),
    message: 'Delete all 3 sessions of "My Episode" from this day and all associated data?',
  },
  {
    name: 'confirmBucketDelete uses a clean singular form for one session',
    ask: () => confirmBucketDelete('Solo Episode', 1),
    message: 'Delete this session of "Solo Episode" from this day and all associated data?',
  },
  {
    name: 'confirmEpisodeDelete includes the episode title',
    ask: () => confirmEpisodeDelete('Episode 4'),
    message: 'Delete "Episode 4" and all its sessions?',
  },
];

for (const { name, ask, message } of copyCases) {
  test(`${name} and returns the user's answer`, async () => {
    for (const answer of [true, false]) {
      await withConfirmGlobals({ confirmAnswer: answer }, async (calls) => {
        assert.equal(await ask(), answer);
        assert.deepEqual(calls, [`browser-confirm:${message}`]);
      });
    }
  });
}
