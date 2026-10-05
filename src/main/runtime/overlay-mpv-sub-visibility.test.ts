import assert from 'node:assert/strict';
import test from 'node:test';

import {
  createEnsureOverlayMpvSubtitlesHiddenHandler,
  createRestoreOverlayMpvSubtitlesHandler,
} from './overlay-mpv-sub-visibility';

type VisibilityState = {
  savedSubVisibility: boolean | null;
  revision: number;
};

function makeEnsure(requestProperty: (name: string) => Promise<unknown>) {
  const state: VisibilityState = { savedSubVisibility: null, revision: 0 };
  const hidden: boolean[] = [];
  const warnings: string[] = [];
  const ensureHidden = createEnsureOverlayMpvSubtitlesHiddenHandler({
    getMpvClient: () => ({ connected: true, requestProperty }),
    getSavedSubVisibility: () => state.savedSubVisibility,
    setSavedSubVisibility: (visible) => {
      state.savedSubVisibility = visible;
    },
    getRevision: () => state.revision,
    setRevision: (revision) => {
      state.revision = revision;
    },
    setMpvSubVisibility: (visible) => {
      hidden.push(visible);
    },
    logWarn: (message) => warnings.push(message),
  });
  return { ensureHidden, state, hidden, warnings };
}

function makeRestore(initial: {
  saved: boolean | null;
  revision: number;
  connected: boolean;
  keepSuppressed: boolean;
}) {
  const state: VisibilityState = {
    savedSubVisibility: initial.saved,
    revision: initial.revision,
  };
  const calls: boolean[] = [];
  const restore = createRestoreOverlayMpvSubtitlesHandler({
    getSavedSubVisibility: () => state.savedSubVisibility,
    setSavedSubVisibility: (visible) => {
      state.savedSubVisibility = visible;
    },
    getRevision: () => state.revision,
    setRevision: (revision) => {
      state.revision = revision;
    },
    isMpvConnected: () => initial.connected,
    shouldKeepSuppressedFromVisibleOverlayBinding: () => initial.keepSuppressed,
    setMpvSubVisibility: (visible) => {
      calls.push(visible);
    },
  });
  return { restore, state, calls };
}

test('ensure overlay mpv subtitle suppression captures previous visibility then hides subtitles', async () => {
  const { ensureHidden, state, hidden } = makeEnsure(async () => 'no');

  await ensureHidden();

  assert.equal(state.savedSubVisibility, false);
  assert.equal(state.revision, 1);
  assert.deepEqual(hidden, [false]);
});

// mpv reports sub-visibility as strings, booleans or numbers depending on the path.
const CAPTURED_VISIBILITY_CASES: Array<[unknown, boolean]> = [
  ['no', false],
  ['false', false],
  ['0', false],
  [false, false],
  [0, false],
  ['yes', true],
  [true, true],
  [1, true],
  ['unrecognized', true],
];

for (const [reported, expected] of CAPTURED_VISIBILITY_CASES) {
  test(`ensure captures sub-visibility ${JSON.stringify(reported)} as ${expected}`, async () => {
    const { ensureHidden, state } = makeEnsure(async () => reported);

    await ensureHidden();

    assert.equal(state.savedSubVisibility, expected);
  });
}

test('ensure falls back to visible restore and logs when the visibility read fails', async () => {
  const { ensureHidden, state, hidden, warnings } = makeEnsure(async () => {
    throw new Error('mpv gone');
  });

  await ensureHidden();

  assert.equal(state.savedSubVisibility, true);
  assert.deepEqual(hidden, [false]);
  assert.equal(warnings.length, 1);
});

test('ensure drops a captured visibility when a newer toggle superseded it', async () => {
  let resolveVisibility!: (value: unknown) => void;
  const { ensureHidden, state } = makeEnsure(
    () =>
      new Promise((resolve) => {
        resolveVisibility = resolve;
      }),
  );

  const pending = ensureHidden();
  state.revision += 1;
  resolveVisibility('no');
  await pending;

  assert.equal(state.savedSubVisibility, null);
});

const RESTORE_CASES: Array<{
  name: string;
  initial: Parameters<typeof makeRestore>[0];
  force?: boolean;
  expected: { saved: boolean | null; revision: number; calls: boolean[] };
}> = [
  {
    name: 'restore overlay mpv subtitle suppression restores saved visibility',
    initial: { saved: false, revision: 4, connected: true, keepSuppressed: false },
    expected: { saved: null, revision: 5, calls: [false] },
  },
  {
    name: 'restore keeps mpv subtitles hidden when visible-overlay binding still requires suppression',
    initial: { saved: true, revision: 9, connected: true, keepSuppressed: true },
    expected: { saved: true, revision: 10, calls: [false] },
  },
  {
    name: 'forced restore ignores visible-overlay suppression during app shutdown',
    initial: { saved: true, revision: 9, connected: true, keepSuppressed: true },
    force: true,
    expected: { saved: null, revision: 10, calls: [true] },
  },
  {
    name: 'restore defers mpv subtitle restore while mpv is disconnected',
    initial: { saved: true, revision: 2, connected: false, keepSuppressed: false },
    expected: { saved: true, revision: 3, calls: [] },
  },
];

for (const c of RESTORE_CASES) {
  test(c.name, () => {
    const { restore, state, calls } = makeRestore(c.initial);

    restore(c.force ? { force: true } : undefined);

    assert.equal(state.savedSubVisibility, c.expected.saved);
    assert.equal(state.revision, c.expected.revision);
    assert.deepEqual(calls, c.expected.calls);
  });
}
