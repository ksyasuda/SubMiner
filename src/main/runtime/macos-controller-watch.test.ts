import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';
import {
  handleBackgroundControllerPress,
  parseControllerWatchLine,
  resolveMacOSControllerWatchCommand,
} from './macos-controller-watch';

test('parseControllerWatchLine reads press lines only', () => {
  assert.deepEqual(parseControllerWatchLine('press 7 2\n'), { buttonIndex: 7, controllerCount: 2 });
  assert.equal(parseControllerWatchLine('press 7'), null);
  assert.equal(parseControllerWatchLine('release 7 1'), null);
});

function fakeFs(files: Record<string, number>) {
  return {
    existsSync: (filePath: string) => filePath in files,
    statSync: (filePath: string) => ({ mtimeMs: files[filePath] }),
  } as unknown as Parameters<typeof resolveMacOSControllerWatchCommand>[1];
}

test('resolveMacOSControllerWatchCommand prefers the packaged binary', () => {
  const packaged = path.join('/res', 'scripts', 'macos-controller-watch');
  assert.deepEqual(
    resolveMacOSControllerWatchCommand(
      { resourcesPath: '/res', mainDirname: '/repo/dist' },
      fakeFs({ [packaged]: 1 }),
    ),
    { command: packaged, args: [] },
  );
});

test('resolveMacOSControllerWatchCommand runs the source when the build is stale', () => {
  const built = path.join('/repo/dist', 'scripts', 'macos-controller-watch');
  const source = path.join('/repo', 'scripts', 'macos-controller-watch.swift');
  const options = { resourcesPath: undefined, mainDirname: '/repo/dist' };

  assert.deepEqual(
    resolveMacOSControllerWatchCommand(options, fakeFs({ [built]: 2, [source]: 1 })),
    {
      command: built,
      args: [],
    },
  );
  assert.deepEqual(
    resolveMacOSControllerWatchCommand(options, fakeFs({ [built]: 1, [source]: 2 })),
    {
      command: 'swift',
      args: [source],
    },
  );
  assert.equal(resolveMacOSControllerWatchCommand(options, fakeFs({})), null);
});

function window(name: string, visible = true) {
  return { name, isDestroyed: () => false, isVisible: () => visible };
}

function run(overrides: {
  controllerCount?: number;
  enabled?: boolean;
  mpvFocused?: boolean;
  modal?: ReturnType<typeof window> | null;
  main?: ReturnType<typeof window> | null;
}): string[] {
  const calls: string[] = [];
  const press = { buttonIndex: 4, controllerCount: overrides.controllerCount ?? 1 };
  handleBackgroundControllerPress(press, {
    isControllerEnabled: () => overrides.enabled ?? true,
    isMpvFocused: () => overrides.mpvFocused ?? true,
    getModalWindow: () => overrides.modal ?? null,
    getMainWindow: () => (overrides.main === undefined ? window('main') : overrides.main),
    focusWindow: (target) => calls.push(`focus:${target.name}`),
    replayPress: (target, index) => calls.push(`replay:${target.name}:${index}`),
  });
  return calls;
}

test('a press while mpv is in front focuses the overlay and replays it', () => {
  assert.deepEqual(run({}), ['focus:main', 'replay:main:4']);
});

test('with several pads connected a press focuses the overlay without replaying', () => {
  assert.deepEqual(run({ controllerCount: 2 }), ['focus:main']);
});

test('an open modal window takes the press instead of the main overlay', () => {
  assert.deepEqual(run({ modal: window('modal') }), ['focus:modal', 'replay:modal:4']);
});

test('presses are ignored unless mpv is in front, the controller is on and an overlay shows', () => {
  assert.deepEqual(run({ mpvFocused: false }), []);
  assert.deepEqual(run({ enabled: false }), []);
  assert.deepEqual(run({ main: window('main', false) }), []);
});
