import assert from 'node:assert/strict';
import test, { beforeEach } from 'node:test';
import {
  isDockIconRetained,
  releaseDockIcon,
  resetDockIconRetentionForTests,
  retainDockIcon,
} from './dock-icon-visibility';

function createDockRecorder() {
  const calls: string[] = [];
  return {
    calls,
    dock: {
      show: () => {
        calls.push('show');
      },
      hide: () => {
        calls.push('hide');
      },
    },
  };
}

beforeEach(() => {
  resetDockIconRetentionForTests();
});

test('retainDockIcon shows the dock once and marks retention on darwin', () => {
  const recorder = createDockRecorder();

  retainDockIcon({ dock: recorder.dock, platform: 'darwin' });
  retainDockIcon({ dock: recorder.dock, platform: 'darwin' });

  assert.deepEqual(recorder.calls, ['show']);
  assert.equal(isDockIconRetained(), true);
});

test('releaseDockIcon rehides only when the last retainer releases and overlay still needs it', () => {
  const recorder = createDockRecorder();

  retainDockIcon({ dock: recorder.dock, platform: 'darwin' });
  retainDockIcon({ dock: recorder.dock, platform: 'darwin' });
  releaseDockIcon({ dock: recorder.dock, platform: 'darwin', shouldRehide: () => true });
  assert.deepEqual(recorder.calls, ['show']);
  assert.equal(isDockIconRetained(), true);

  releaseDockIcon({ dock: recorder.dock, platform: 'darwin', shouldRehide: () => true });
  assert.deepEqual(recorder.calls, ['show', 'hide']);
  assert.equal(isDockIconRetained(), false);
});

test('releaseDockIcon leaves the dock visible when no overlay needs the accessory transform', () => {
  const recorder = createDockRecorder();

  retainDockIcon({ dock: recorder.dock, platform: 'darwin' });
  releaseDockIcon({ dock: recorder.dock, platform: 'darwin', shouldRehide: () => false });

  assert.deepEqual(recorder.calls, ['show']);
  assert.equal(isDockIconRetained(), false);
});

test('release without retain and non-darwin platforms are no-ops', () => {
  const recorder = createDockRecorder();

  releaseDockIcon({ dock: recorder.dock, platform: 'darwin', shouldRehide: () => true });
  retainDockIcon({ dock: recorder.dock, platform: 'linux' });
  releaseDockIcon({ dock: recorder.dock, platform: 'win32', shouldRehide: () => true });

  assert.deepEqual(recorder.calls, []);
  assert.equal(isDockIconRetained(), false);
});
