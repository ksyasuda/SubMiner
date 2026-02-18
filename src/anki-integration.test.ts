import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AnkiIntegration } from './anki-integration';

interface IntegrationTestContext {
  integration: AnkiIntegration;
  calls: {
    findNotes: number;
    notesInfo: number;
  };
  stateDir: string;
}

function createIntegrationTestContext(
  options: {
    highlightEnabled?: boolean;
    onFindNotes?: () => Promise<number[]>;
    onNotesInfo?: () => Promise<unknown[]>;
    stateDirPrefix?: string;
  } = {},
): IntegrationTestContext {
  const calls = {
    findNotes: 0,
    notesInfo: 0,
  };

  const stateDir = fs.mkdtempSync(
    path.join(os.tmpdir(), options.stateDirPrefix ?? 'subminer-anki-integration-'),
  );
  const knownWordCacheStatePath = path.join(stateDir, 'known-words-cache.json');

  const client = {
    findNotes: async () => {
      calls.findNotes += 1;
      if (options.onFindNotes) {
        return options.onFindNotes();
      }
      return [] as number[];
    },
    notesInfo: async () => {
      calls.notesInfo += 1;
      if (options.onNotesInfo) {
        return options.onNotesInfo();
      }
      return [] as unknown[];
    },
  } as {
    findNotes: () => Promise<number[]>;
    notesInfo: () => Promise<unknown[]>;
  };

  const integration = new AnkiIntegration(
    {
      nPlusOne: {
        highlightEnabled: options.highlightEnabled ?? true,
      },
    },
    {} as never,
    {} as never,
    undefined,
    undefined,
    undefined,
    knownWordCacheStatePath,
  );

  const integrationWithClient = integration as unknown as {
    client: {
      findNotes: () => Promise<number[]>;
      notesInfo: () => Promise<unknown[]>;
    };
  };
  integrationWithClient.client = client;

  const privateState = integration as unknown as {
    knownWordsScope: string;
    knownWordsLastRefreshedAtMs: number;
  };
  privateState.knownWordsScope = 'is:note';
  privateState.knownWordsLastRefreshedAtMs = Date.now();

  return {
    integration,
    calls,
    stateDir,
  };
}

function cleanupIntegrationTestContext(ctx: IntegrationTestContext): void {
  fs.rmSync(ctx.stateDir, { recursive: true, force: true });
}

test('AnkiIntegration.refreshKnownWordCache bypasses stale checks', async () => {
  const ctx = createIntegrationTestContext();

  try {
    await ctx.integration.refreshKnownWordCache();

    assert.equal(ctx.calls.findNotes, 1);
    assert.equal(ctx.calls.notesInfo, 0);
  } finally {
    cleanupIntegrationTestContext(ctx);
  }
});

test('AnkiIntegration.refreshKnownWordCache skips work when highlight mode is disabled', async () => {
  const ctx = createIntegrationTestContext({
    highlightEnabled: false,
    stateDirPrefix: 'subminer-anki-integration-disabled-',
  });

  try {
    await ctx.integration.refreshKnownWordCache();

    assert.equal(ctx.calls.findNotes, 0);
    assert.equal(ctx.calls.notesInfo, 0);
  } finally {
    cleanupIntegrationTestContext(ctx);
  }
});

test('AnkiIntegration.refreshKnownWordCache deduplicates concurrent refreshes', async () => {
  let releaseFindNotes: (() => void) | undefined;
  const findNotesPromise = new Promise<void>((resolve) => {
    releaseFindNotes = resolve;
  });

  const ctx = createIntegrationTestContext({
    onFindNotes: async () => {
      await findNotesPromise;
      return [] as number[];
    },
    stateDirPrefix: 'subminer-anki-integration-concurrent-',
  });

  const first = ctx.integration.refreshKnownWordCache();
  await Promise.resolve();
  const second = ctx.integration.refreshKnownWordCache();

  if (releaseFindNotes !== undefined) {
    releaseFindNotes();
  }

  await Promise.all([first, second]);

  try {
    assert.equal(ctx.calls.findNotes, 1);
    assert.equal(ctx.calls.notesInfo, 0);
  } finally {
    cleanupIntegrationTestContext(ctx);
  }
});
