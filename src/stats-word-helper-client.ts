export type StatsWordHelperResponse = {
  ok: boolean;
  noteId?: number;
  deckName?: string;
  error?: string;
};

type StatsWordHelperMode = 'add-word' | 'deck-name';

export type StatsWordHelperSpawnOptions = {
  scriptPath: string;
  responsePath: string;
  userDataPath: string;
  mode: StatsWordHelperMode;
  word?: string;
};

export function createInvokeStatsWordHelperHandler(deps: {
  createTempDir: (prefix: string) => string;
  joinPath: (...parts: string[]) => string;
  spawnHelper: (options: StatsWordHelperSpawnOptions) => Promise<number>;
  waitForResponse: (responsePath: string) => Promise<StatsWordHelperResponse>;
  removeDir: (targetPath: string) => void;
}) {
  return async (options: {
    helperScriptPath: string;
    userDataPath: string;
    word: string;
  }): Promise<number> => {
    const tempDir = deps.createTempDir('subminer-stats-word-helper-');
    const responsePath = deps.joinPath(tempDir, 'response.json');

    try {
      const helperExitPromise = deps.spawnHelper({
        scriptPath: options.helperScriptPath,
        responsePath,
        userDataPath: options.userDataPath,
        mode: 'add-word',
        word: options.word,
      });

      const startupResult = await Promise.race([
        deps
          .waitForResponse(responsePath)
          .then((response) => ({ kind: 'response' as const, response })),
        helperExitPromise.then((status) => ({ kind: 'exit' as const, status })),
      ]);

      let response: StatsWordHelperResponse;
      if (startupResult.kind === 'response') {
        response = startupResult.response;
      } else {
        if (startupResult.status !== 0) {
          throw new Error(
            `Stats word helper exited before response (status ${startupResult.status}).`,
          );
        }
        response = await deps.waitForResponse(responsePath);
      }

      const exitStatus = await helperExitPromise;
      if (exitStatus !== 0) {
        throw new Error(`Stats word helper exited with status ${exitStatus}.`);
      }
      if (!response.ok || typeof response.noteId !== 'number') {
        throw new Error(response.error || 'Stats word helper failed.');
      }
      return response.noteId;
    } finally {
      deps.removeDir(tempDir);
    }
  };
}

export function createReadStatsYomitanDeckNameHandler(deps: {
  createTempDir: (prefix: string) => string;
  joinPath: (...parts: string[]) => string;
  spawnHelper: (options: StatsWordHelperSpawnOptions) => Promise<number>;
  waitForResponse: (responsePath: string) => Promise<StatsWordHelperResponse>;
  removeDir: (targetPath: string) => void;
}) {
  return async (options: { helperScriptPath: string; userDataPath: string }): Promise<string> => {
    const tempDir = deps.createTempDir('subminer-stats-word-helper-');
    const responsePath = deps.joinPath(tempDir, 'response.json');

    try {
      const helperExitPromise = deps.spawnHelper({
        scriptPath: options.helperScriptPath,
        responsePath,
        userDataPath: options.userDataPath,
        mode: 'deck-name',
      });

      const startupResult = await Promise.race([
        deps
          .waitForResponse(responsePath)
          .then((response) => ({ kind: 'response' as const, response })),
        helperExitPromise.then((status) => ({ kind: 'exit' as const, status })),
      ]);

      let response: StatsWordHelperResponse;
      if (startupResult.kind === 'response') {
        response = startupResult.response;
      } else {
        if (startupResult.status !== 0) {
          throw new Error(
            `Stats word helper exited before response (status ${startupResult.status}).`,
          );
        }
        response = await deps.waitForResponse(responsePath);
      }

      const exitStatus = await helperExitPromise;
      if (exitStatus !== 0) {
        throw new Error(`Stats word helper exited with status ${exitStatus}.`);
      }
      if (!response.ok || typeof response.deckName !== 'string') {
        throw new Error(response.error || 'Stats word helper failed.');
      }
      return response.deckName.trim();
    } finally {
      deps.removeDir(tempDir);
    }
  };
}
