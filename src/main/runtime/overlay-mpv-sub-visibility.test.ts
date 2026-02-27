import assert from 'node:assert/strict';
import test from 'node:test';

import {
  createEnsureOverlayMpvSubtitlesHiddenHandler,
  createRestoreOverlayMpvSubtitlesHandler,
} from './overlay-mpv-sub-visibility';

type VisibilityState = {
  savedSubVisibility: boolean | null;
  savedSecondarySubVisibility: boolean | null;
  revision: number;
};

test('ensure overlay mpv subtitle suppression captures previous visibility then hides subtitles', async () => {
  const state: VisibilityState = {
    savedSubVisibility: null,
    savedSecondarySubVisibility: null,
    revision: 0,
  };
  const calls: boolean[] = [];

  const ensureHidden = createEnsureOverlayMpvSubtitlesHiddenHandler({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (_name: string) => 'no',
    }),
    getSavedSubVisibility: () => state.savedSubVisibility,
    setSavedSubVisibility: (visible) => {
      state.savedSubVisibility = visible;
    },
    getSavedSecondarySubVisibility: () => state.savedSecondarySubVisibility,
    setSavedSecondarySubVisibility: (visible) => {
      state.savedSecondarySubVisibility = visible;
    },
    getRevision: () => state.revision,
    setRevision: (revision) => {
      state.revision = revision;
    },
    setMpvSubVisibility: (visible) => {
      calls.push(visible);
    },
    setMpvSecondarySubVisibility: (visible) => {
      calls.push(visible);
    },
    logWarn: () => {},
  });

  await ensureHidden();

  assert.equal(state.savedSubVisibility, false);
  assert.equal(state.savedSecondarySubVisibility, false);
  assert.equal(state.revision, 1);
  assert.deepEqual(calls, [false, false]);
});

test('restore overlay mpv subtitle suppression restores saved visibility', () => {
  const state: VisibilityState = {
    savedSubVisibility: false,
    savedSecondarySubVisibility: true,
    revision: 4,
  };
  const calls: boolean[] = [];

  const restore = createRestoreOverlayMpvSubtitlesHandler({
    getSavedSubVisibility: () => state.savedSubVisibility,
    setSavedSubVisibility: (visible) => {
      state.savedSubVisibility = visible;
    },
    getSavedSecondarySubVisibility: () => state.savedSecondarySubVisibility,
    setSavedSecondarySubVisibility: (visible) => {
      state.savedSecondarySubVisibility = visible;
    },
    getRevision: () => state.revision,
    setRevision: (revision) => {
      state.revision = revision;
    },
    isMpvConnected: () => true,
    shouldKeepSuppressedFromVisibleOverlayBinding: () => false,
    setMpvSubVisibility: (visible) => {
      calls.push(visible);
    },
    setMpvSecondarySubVisibility: (visible) => {
      calls.push(visible);
    },
  });

  restore();

  assert.equal(state.savedSubVisibility, null);
  assert.equal(state.savedSecondarySubVisibility, null);
  assert.equal(state.revision, 5);
  assert.deepEqual(calls, [false, true]);
});

test('restore keeps mpv subtitles hidden when visible-overlay binding still requires suppression', () => {
  const state: VisibilityState = {
    savedSubVisibility: true,
    savedSecondarySubVisibility: true,
    revision: 9,
  };
  const calls: boolean[] = [];

  const restore = createRestoreOverlayMpvSubtitlesHandler({
    getSavedSubVisibility: () => state.savedSubVisibility,
    setSavedSubVisibility: (visible) => {
      state.savedSubVisibility = visible;
    },
    getSavedSecondarySubVisibility: () => state.savedSecondarySubVisibility,
    setSavedSecondarySubVisibility: (visible) => {
      state.savedSecondarySubVisibility = visible;
    },
    getRevision: () => state.revision,
    setRevision: (revision) => {
      state.revision = revision;
    },
    isMpvConnected: () => true,
    shouldKeepSuppressedFromVisibleOverlayBinding: () => true,
    setMpvSubVisibility: (visible) => {
      calls.push(visible);
    },
    setMpvSecondarySubVisibility: (visible) => {
      calls.push(visible);
    },
  });

  restore();

  assert.equal(state.savedSubVisibility, true);
  assert.equal(state.savedSecondarySubVisibility, true);
  assert.equal(state.revision, 10);
  assert.deepEqual(calls, [false, false]);
});

test('restore defers mpv subtitle restore while mpv is disconnected', () => {
  const state: VisibilityState = {
    savedSubVisibility: true,
    savedSecondarySubVisibility: false,
    revision: 2,
  };
  const calls: boolean[] = [];

  const restore = createRestoreOverlayMpvSubtitlesHandler({
    getSavedSubVisibility: () => state.savedSubVisibility,
    setSavedSubVisibility: (visible) => {
      state.savedSubVisibility = visible;
    },
    getSavedSecondarySubVisibility: () => state.savedSecondarySubVisibility,
    setSavedSecondarySubVisibility: (visible) => {
      state.savedSecondarySubVisibility = visible;
    },
    getRevision: () => state.revision,
    setRevision: (revision) => {
      state.revision = revision;
    },
    isMpvConnected: () => false,
    shouldKeepSuppressedFromVisibleOverlayBinding: () => false,
    setMpvSubVisibility: (visible) => {
      calls.push(visible);
    },
    setMpvSecondarySubVisibility: (visible) => {
      calls.push(visible);
    },
  });

  restore();

  assert.equal(state.savedSubVisibility, true);
  assert.equal(state.revision, 3);
  assert.deepEqual(calls, []);
});
