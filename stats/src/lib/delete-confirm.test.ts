import assert from 'node:assert/strict';
import test from 'node:test';
import {
  confirmBucketDelete,
  confirmDayGroupDelete,
  confirmEpisodeDelete,
  confirmSessionDelete,
  setDeleteConfirmPresenter,
} from './delete-confirm';

test('confirmSessionDelete uses the shared session delete warning copy', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(await confirmSessionDelete(), true);
    assert.deepEqual(calls, ['Delete this session and all associated data?']);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmSessionDelete suspends stats overlay layering around native confirm', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  const originalElectronAPI = (
    globalThis as typeof globalThis & {
      electronAPI?: {
        stats?: {
          beginNativeDialog?: () => void;
          endNativeDialog?: () => void;
        };
      };
    }
  ).electronAPI;
  (
    globalThis as typeof globalThis & {
      electronAPI?: {
        stats?: {
          beginNativeDialog?: () => void;
          endNativeDialog?: () => void;
        };
      };
    }
  ).electronAPI = {
    stats: {
      beginNativeDialog: () => calls.push('begin-native-dialog'),
      endNativeDialog: () => calls.push('end-native-dialog'),
    },
  };
  globalThis.confirm = ((message?: string) => {
    calls.push(`confirm:${message ?? ''}`);
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(await confirmSessionDelete(), true);
    assert.deepEqual(calls, [
      'begin-native-dialog',
      'confirm:Delete this session and all associated data?',
      'end-native-dialog',
    ]);
  } finally {
    globalThis.confirm = originalConfirm;
    (
      globalThis as typeof globalThis & {
        electronAPI?: {
          stats?: {
            beginNativeDialog?: () => void;
            endNativeDialog?: () => void;
          };
        };
      }
    ).electronAPI = originalElectronAPI;
  }
});

test('confirmSessionDelete uses parented Electron confirm when available', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  const originalElectronAPI = (
    globalThis as typeof globalThis & {
      electronAPI?: {
        stats?: {
          confirmNativeDialog?: (message: string) => boolean;
          beginNativeDialog?: () => void;
          endNativeDialog?: () => void;
        };
      };
    }
  ).electronAPI;
  (
    globalThis as typeof globalThis & {
      electronAPI?: {
        stats?: {
          confirmNativeDialog?: (message: string) => boolean;
          beginNativeDialog?: () => void;
          endNativeDialog?: () => void;
        };
      };
    }
  ).electronAPI = {
    stats: {
      confirmNativeDialog: (message) => {
        calls.push(`native-confirm:${message}`);
        return false;
      },
      beginNativeDialog: () => calls.push('begin-native-dialog'),
      endNativeDialog: () => calls.push('end-native-dialog'),
    },
  };
  globalThis.confirm = ((message?: string) => {
    calls.push(`browser-confirm:${message ?? ''}`);
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(await confirmSessionDelete(), false);
    assert.deepEqual(calls, ['native-confirm:Delete this session and all associated data?']);
  } finally {
    globalThis.confirm = originalConfirm;
    (
      globalThis as typeof globalThis & {
        electronAPI?: {
          stats?: {
            confirmNativeDialog?: (message: string) => boolean;
            beginNativeDialog?: () => void;
            endNativeDialog?: () => void;
          };
        };
      }
    ).electronAPI = originalElectronAPI;
  }
});

test('confirmSessionDelete uses the registered stats presenter before native or browser confirm', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  const originalElectronAPI = (
    globalThis as typeof globalThis & {
      electronAPI?: {
        stats?: {
          confirmNativeDialog?: (message: string) => boolean;
        };
      };
    }
  ).electronAPI;
  (
    globalThis as typeof globalThis & {
      electronAPI?: {
        stats?: {
          confirmNativeDialog?: (message: string) => boolean;
        };
      };
    }
  ).electronAPI = {
    stats: {
      confirmNativeDialog: (message) => {
        calls.push(`native-confirm:${message}`);
        return true;
      },
    },
  };
  globalThis.confirm = ((message?: string) => {
    calls.push(`browser-confirm:${message ?? ''}`);
    return true;
  }) as typeof globalThis.confirm;

  const unregister = setDeleteConfirmPresenter(async (message) => {
    calls.push(`presenter:${message}`);
    return false;
  });

  try {
    assert.equal(await confirmSessionDelete(), false);
    assert.deepEqual(calls, ['presenter:Delete this session and all associated data?']);
  } finally {
    unregister();
    globalThis.confirm = originalConfirm;
    (
      globalThis as typeof globalThis & {
        electronAPI?: {
          stats?: {
            confirmNativeDialog?: (message: string) => boolean;
          };
        };
      }
    ).electronAPI = originalElectronAPI;
  }
});

test('confirmDayGroupDelete includes the day label and count in the warning copy', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(await confirmDayGroupDelete('Today', 3), true);
    assert.deepEqual(calls, ['Delete all 3 sessions from Today and all associated data?']);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmDayGroupDelete uses singular for one session', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(await confirmDayGroupDelete('Yesterday', 1), true);
    assert.deepEqual(calls, ['Delete this session from Yesterday and all associated data?']);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmBucketDelete asks about merging multiple sessions of the same episode', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(await confirmBucketDelete('My Episode', 3), true);
    assert.deepEqual(calls, [
      'Delete all 3 sessions of "My Episode" from this day and all associated data?',
    ]);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmBucketDelete uses a clean singular form for one session', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return false;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(await confirmBucketDelete('Solo Episode', 1), false);
    assert.deepEqual(calls, [
      'Delete this session of "Solo Episode" from this day and all associated data?',
    ]);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmEpisodeDelete includes the episode title in the shared warning copy', async () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return false;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(await confirmEpisodeDelete('Episode 4'), false);
    assert.deepEqual(calls, ['Delete "Episode 4" and all its sessions?']);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});
