import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildMediaRuntimeMainDepsHandler } from './media-runtime-main-deps';

test('media runtime main deps builder maps state and subtitle broadcast channel', () => {
  const calls: string[] = [];
  let currentPath: string | null = '/tmp/a.mkv';
  let subtitlePosition: unknown = null;
  let currentTitle: string | null = 'Title';

  const deps = createBuildMediaRuntimeMainDepsHandler({
    isRemoteMediaPath: (mediaPath) => mediaPath.startsWith('http'),
    loadSubtitlePosition: () => ({ x: 1 }) as never,
    getCurrentMediaPath: () => currentPath,
    getPendingSubtitlePosition: () => null,
    getSubtitlePositionsDir: () => '/tmp/subs',
    setCurrentMediaPath: (mediaPath) => {
      currentPath = mediaPath;
      calls.push(`path:${String(mediaPath)}`);
    },
    clearPendingSubtitlePosition: () => calls.push('clear-pending'),
    setSubtitlePosition: (position) => {
      subtitlePosition = position;
      calls.push('set-position');
    },
    broadcastToOverlayWindows: (channel, payload) =>
      calls.push(`broadcast:${channel}:${JSON.stringify(payload)}`),
    getCurrentMediaTitle: () => currentTitle,
    setCurrentMediaTitle: (title) => {
      currentTitle = title;
      calls.push(`title:${String(title)}`);
    },
  })();

  assert.equal(deps.isRemoteMediaPath('http://x'), true);
  assert.equal(deps.getCurrentMediaPath(), '/tmp/a.mkv');
  assert.equal(deps.getSubtitlePositionsDir(), '/tmp/subs');
  assert.deepEqual(deps.loadSubtitlePosition(), { x: 1 });
  deps.setCurrentMediaPath('/tmp/b.mkv');
  deps.clearPendingSubtitlePosition();
  deps.setSubtitlePosition({ line: 1 } as never);
  deps.broadcastSubtitlePosition({ line: 1 } as never);
  deps.setCurrentMediaTitle('Next');
  assert.equal(currentPath, '/tmp/b.mkv');
  assert.deepEqual(subtitlePosition, { line: 1 });
  assert.equal(currentTitle, 'Next');
  assert.deepEqual(calls, [
    'path:/tmp/b.mkv',
    'clear-pending',
    'set-position',
    'broadcast:subtitle-position:set:{"line":1}',
    'title:Next',
  ]);
});
