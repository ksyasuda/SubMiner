import assert from 'node:assert/strict';
import test from 'node:test';
import { createWatchConfigPathHandler } from './config-hot-reload-main-deps';

test('watch config path handler watches file directly when config exists', () => {
  const calls: string[] = [];
  const watchConfigPath = createWatchConfigPathHandler({
    fileExists: () => true,
    dirname: (path) => path.split('/').slice(0, -1).join('/'),
    watchPath: (targetPath, nextListener) => {
      calls.push(`watch:${targetPath}`);
      nextListener('change', 'ignored');
      return { close: () => calls.push('close') };
    },
  });

  const watcher = watchConfigPath('/tmp/config.jsonc', () => calls.push('change'));
  watcher.close();
  assert.deepEqual(calls, ['watch:/tmp/config.jsonc', 'change', 'close']);
});

test('watch config path handler filters directory events to config files only', () => {
  const calls: string[] = [];
  const watchConfigPath = createWatchConfigPathHandler({
    fileExists: () => false,
    dirname: (path) => path.split('/').slice(0, -1).join('/'),
    watchPath: (_targetPath, nextListener) => {
      nextListener('change', 'foo.txt');
      nextListener('change', 'config.json');
      nextListener('change', 'config.jsonc');
      nextListener('change', null);
      return { close: () => {} };
    },
  });

  watchConfigPath('/tmp/config.jsonc', () => calls.push('change'));
  assert.deepEqual(calls, ['change', 'change', 'change']);
});
