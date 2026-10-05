import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createInvokeStatsWordHelperHandler,
  createReadStatsYomitanDeckNameHandler,
  type StatsWordHelperResponse,
  type StatsWordHelperSpawnOptions,
} from './stats-word-helper-client';

type Script = {
  response: StatsWordHelperResponse;
  exitStatus?: number;
  // When true the helper exits before any response is readable, so the client must read it after exit.
  helperExitsFirst?: boolean;
};

function createHarness(script: Script) {
  const spawned: StatsWordHelperSpawnOptions[] = [];
  const removedDirs: string[] = [];
  let waitCalls = 0;
  const deps = {
    createTempDir: () => '/tmp/stats-word-helper',
    joinPath: (...parts: string[]) => parts.join('/'),
    spawnHelper: async (options: StatsWordHelperSpawnOptions) => {
      spawned.push(options);
      const status = script.exitStatus ?? 0;
      if (script.helperExitsFirst) return status;
      return new Promise<number>((resolve) => setTimeout(() => resolve(status), 5));
    },
    waitForResponse: async () => {
      waitCalls += 1;
      if (script.helperExitsFirst && waitCalls === 1) return new Promise<never>(() => {});
      return script.response;
    },
    removeDir: (targetPath: string) => {
      removedDirs.push(targetPath);
    },
  };
  return { deps, spawned, removedDirs };
}

const addWordOptions = {
  helperScriptPath: '/tmp/stats-word-helper.js',
  userDataPath: '/tmp/SubMiner',
  word: '猫',
};
const deckNameOptions = {
  helperScriptPath: '/tmp/stats-word-helper.js',
  userDataPath: '/tmp/SubMiner',
};

const ADD_WORD_FAILURES: Array<{ name: string; script: Script; error: RegExp }> = [
  {
    name: 'helper reports an error response',
    script: { response: { ok: false, error: 'helper failed' } },
    error: /helper failed/,
  },
  {
    name: 'helper reports failure without a message',
    script: { response: { ok: false } },
    error: /Stats word helper failed/,
  },
  {
    name: 'response has no note id',
    script: { response: { ok: true } },
    error: /Stats word helper failed/,
  },
  {
    name: 'helper exits non-zero before responding',
    script: { response: { ok: true, noteId: 1 }, exitStatus: 3, helperExitsFirst: true },
    error: /exited before response \(status 3\)/,
  },
  {
    name: 'helper exits non-zero after a successful response',
    script: { response: { ok: true, noteId: 1 }, exitStatus: 2 },
    error: /exited with status 2/,
  },
];

test('word helper client returns note id and spawns the helper in add-word mode', async () => {
  const { deps, spawned, removedDirs } = createHarness({ response: { ok: true, noteId: 123 } });

  const noteId = await createInvokeStatsWordHelperHandler(deps)(addWordOptions);

  assert.equal(noteId, 123);
  assert.equal(spawned.length, 1);
  assert.deepEqual(spawned[0], {
    scriptPath: addWordOptions.helperScriptPath,
    responsePath: '/tmp/stats-word-helper/response.json',
    userDataPath: addWordOptions.userDataPath,
    mode: 'add-word',
    word: '猫',
  });
  assert.deepEqual(removedDirs, ['/tmp/stats-word-helper']);
});

test('word helper client reads the response after a clean helper exit', async () => {
  const { deps } = createHarness({
    response: { ok: true, noteId: 7 },
    helperExitsFirst: true,
  });

  assert.equal(await createInvokeStatsWordHelperHandler(deps)(addWordOptions), 7);
});

for (const c of ADD_WORD_FAILURES) {
  test(`word helper client rejects and cleans up when ${c.name}`, async () => {
    const { deps, removedDirs } = createHarness(c.script);

    await assert.rejects(createInvokeStatsWordHelperHandler(deps)(addWordOptions), c.error);
    assert.deepEqual(removedDirs, ['/tmp/stats-word-helper']);
  });
}

test('word helper client returns the trimmed Yomitan deck name from deck-name mode', async () => {
  const { deps, spawned, removedDirs } = createHarness({
    response: { ok: true, deckName: ' Minecraft ' },
  });

  const deckName = await createReadStatsYomitanDeckNameHandler(deps)(deckNameOptions);

  assert.equal(deckName, 'Minecraft');
  assert.equal(spawned[0]?.mode, 'deck-name');
  assert.equal(spawned[0]?.word, undefined);
  assert.deepEqual(removedDirs, ['/tmp/stats-word-helper']);
});

test('word helper client rejects a deck-name response without a deck name', async () => {
  const { deps, removedDirs } = createHarness({ response: { ok: true } });

  await assert.rejects(
    createReadStatsYomitanDeckNameHandler(deps)(deckNameOptions),
    /Stats word helper failed/,
  );
  assert.deepEqual(removedDirs, ['/tmp/stats-word-helper']);
});
