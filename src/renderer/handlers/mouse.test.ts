import assert from 'node:assert/strict';
import test from 'node:test';

import { createMouseHandlers } from './mouse.js';

function createClassList() {
  const classes = new Set<string>();
  return {
    add: (...tokens: string[]) => {
      for (const token of tokens) {
        classes.add(token);
      }
    },
    remove: (...tokens: string[]) => {
      for (const token of tokens) {
        classes.delete(token);
      }
    },
    contains: (token: string) => classes.has(token),
  };
}

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

function createMouseTestContext() {
  const overlayClassList = createClassList();
  const subtitleRootClassList = createClassList();
  const subtitleContainerClassList = createClassList();

  const ctx = {
    dom: {
      overlay: {
        classList: overlayClassList,
      },
      subtitleRoot: {
        classList: subtitleRootClassList,
      },
      subtitleContainer: {
        classList: subtitleContainerClassList,
        style: { cursor: '' },
        addEventListener: () => {},
      },
      secondarySubContainer: {
        addEventListener: () => {},
      },
    },
    platform: {
      shouldToggleMouseIgnore: false,
      isMacOSPlatform: false,
    },
    state: {
      isOverSubtitle: false,
      isDragging: false,
      dragStartY: 0,
      startYPercent: 0,
    },
  };

  return ctx;
}

test('auto-pause on subtitle hover pauses on enter and resumes on leave when enabled', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];

  const handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => true,
    getPlaybackPaused: async () => false,
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  await handlers.handleMouseEnter();
  await handlers.handleMouseLeave();

  assert.deepEqual(mpvCommands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'pause', 'no'],
  ]);
});

test('auto-pause on subtitle hover does not unpause when playback becomes paused on leave', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];
  const playbackPausedStates = [false, true];

  const handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => true,
    getPlaybackPaused: async () => playbackPausedStates.shift() ?? true,
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  await handlers.handleMouseEnter();
  await handlers.handleMouseLeave();

  assert.deepEqual(mpvCommands, [['set_property', 'pause', 'yes']]);
});

test('auto-pause on subtitle hover is skipped when disabled in config', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];

  const handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => false,
    getPlaybackPaused: async () => false,
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  await handlers.handleMouseEnter();
  await handlers.handleMouseLeave();

  assert.deepEqual(mpvCommands, []);
});

test('pending hover pause check is ignored when mouse leaves before pause state resolves', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];
  const deferred = createDeferred<boolean | null>();

  const handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => true,
    getPlaybackPaused: async () => deferred.promise,
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  const enterPromise = handlers.handleMouseEnter();
  await handlers.handleMouseLeave();
  deferred.resolve(false);
  await enterPromise;

  assert.deepEqual(mpvCommands, []);
});
