import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createInvokeStatsWordHelperHandler,
  createReadStatsYomitanDeckNameHandler,
} from './stats-word-helper-client';

test('word helper client returns note id when helper responds before exit', async () => {
  const calls: string[] = [];
  const handler = createInvokeStatsWordHelperHandler({
    createTempDir: () => '/tmp/stats-word-helper',
    joinPath: (...parts) => parts.join('/'),
    spawnHelper: async (options) => {
      calls.push(
        `spawnHelper:${options.scriptPath}:${options.responsePath}:${options.userDataPath}:${options.word}`,
      );
      return new Promise<number>((resolve) => setTimeout(() => resolve(0), 20));
    },
    waitForResponse: async (responsePath) => {
      calls.push(`waitForResponse:${responsePath}`);
      return { ok: true, noteId: 123 };
    },
    removeDir: (targetPath) => {
      calls.push(`removeDir:${targetPath}`);
    },
  });

  const noteId = await handler({
    helperScriptPath: '/tmp/stats-word-helper.js',
    userDataPath: '/tmp/SubMiner',
    word: '猫',
  });

  assert.equal(noteId, 123);
  assert.deepEqual(calls, [
    'spawnHelper:/tmp/stats-word-helper.js:/tmp/stats-word-helper/response.json:/tmp/SubMiner:猫',
    'waitForResponse:/tmp/stats-word-helper/response.json',
    'removeDir:/tmp/stats-word-helper',
  ]);
});

test('word helper client returns Yomitan deck name from helper read-deck mode', async () => {
  const calls: string[] = [];
  const handler = createReadStatsYomitanDeckNameHandler({
    createTempDir: () => '/tmp/stats-word-helper',
    joinPath: (...parts) => parts.join('/'),
    spawnHelper: async (options) => {
      calls.push(
        `spawnHelper:${options.scriptPath}:${options.responsePath}:${options.userDataPath}:${options.mode}`,
      );
      return new Promise<number>((resolve) => setTimeout(() => resolve(0), 20));
    },
    waitForResponse: async (responsePath) => {
      calls.push(`waitForResponse:${responsePath}`);
      return { ok: true, deckName: ' Minecraft ' };
    },
    removeDir: (targetPath) => {
      calls.push(`removeDir:${targetPath}`);
    },
  });

  const deckName = await handler({
    helperScriptPath: '/tmp/stats-word-helper.js',
    userDataPath: '/tmp/SubMiner',
  });

  assert.equal(deckName, 'Minecraft');
  assert.deepEqual(calls, [
    'spawnHelper:/tmp/stats-word-helper.js:/tmp/stats-word-helper/response.json:/tmp/SubMiner:deck-name',
    'waitForResponse:/tmp/stats-word-helper/response.json',
    'removeDir:/tmp/stats-word-helper',
  ]);
});

test('word helper client throws helper response errors', async () => {
  const handler = createInvokeStatsWordHelperHandler({
    createTempDir: () => '/tmp/stats-word-helper',
    joinPath: (...parts) => parts.join('/'),
    spawnHelper: async () => 0,
    waitForResponse: async () => ({ ok: false, error: 'helper failed' }),
    removeDir: () => {},
  });

  await assert.rejects(
    async () =>
      handler({
        helperScriptPath: '/tmp/stats-word-helper.js',
        userDataPath: '/tmp/SubMiner',
        word: '猫',
      }),
    /helper failed/,
  );
});
