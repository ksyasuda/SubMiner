// Shared harness for the Yomitan parser-runtime and scan-runtime tests: fake
// parser-window deps whose injected scripts run in a vm context, plus the
// backend stubs the scanner tests drive them with. Kept out of the test files
// so the runtime tests and the in-page scanner tests can share one setup.
import * as vm from 'node:vm';

export function createDeps(
  executeJavaScript: (script: string) => Promise<unknown>,
  options?: {
    createYomitanExtensionWindow?: (pageName: string) => Promise<unknown>;
  },
) {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async (script: string) => await executeJavaScript(script),
    },
  };

  return {
    getYomitanExt: () => ({ id: 'ext-id' }) as never,
    getYomitanParserWindow: () => parserWindow as never,
    setYomitanParserWindow: () => undefined,
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => undefined,
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => undefined,
    createYomitanExtensionWindow: options?.createYomitanExtensionWindow as never,
  };
}

function createYomitanScriptSandbox(handler: (action: string, params: unknown) => unknown) {
  const storage: Record<string, unknown> = {};
  return {
    chrome: {
      storage: {
        local: {
          get: async () => ({ ...storage }),
          set: async (value: Record<string, unknown>) => {
            Object.assign(storage, value);
          },
        },
      },
      runtime: {
        lastError: null,
        sendMessage: (
          payload: { action?: string; params?: unknown },
          callback: (response: { result?: unknown; error?: { message?: string } }) => void,
        ) => {
          try {
            // Messages cross a process boundary in Electron; clone params into
            // the host realm so handlers can compare them structurally.
            const params =
              payload.params === undefined ? undefined : structuredClone(payload.params);
            callback({ result: handler(payload.action ?? '', params) });
          } catch (error) {
            callback({ error: { message: (error as Error).message } });
          }
        },
      },
    },
    Array,
    Error,
    JSON,
    Map,
    Math,
    Number,
    Object,
    Promise,
    RegExp,
    Set,
    String,
  };
}

export async function runInjectedYomitanScript(
  script: string,
  handler: (action: string, params: unknown) => unknown,
): Promise<unknown> {
  // Clone results into the host realm, matching Electron's process boundary.
  return structuredClone(await vm.runInNewContext(script, createYomitanScriptSandbox(handler)));
}

// Persistent page context shared across executeJavaScript calls, matching the
// real parser window: the scan runtime is installed once via
// globalThis.__subminerYomitanScan and per-line calls reuse it (and its
// cross-line termsFind cache).
function createPersistentYomitanScriptRunner(
  handler: (action: string, params: unknown) => unknown,
): (script: string) => Promise<unknown> {
  const context = vm.createContext(createYomitanScriptSandbox(handler));
  // Clone results into the host realm, matching Electron's process boundary.
  return async (script: string) => structuredClone(await vm.runInContext(script, context));
}

// Deps whose parser window executes every injected script (profile metadata,
// scan runtime install, per-line scan calls, parseText fallback) inside one
// persistent vm context, dispatching backend actions to `handler`.
export function createScanDeps(
  handler: (action: string, params: unknown) => unknown,
  options?: { onScript?: (script: string) => void },
) {
  const runScript = createPersistentYomitanScriptRunner(handler);
  return createDeps(async (script) => {
    options?.onScript?.(script);
    return await runScript(script);
  });
}

export function countTermsFindLookups(lookups: string[], prefix: string): number {
  return lookups.filter((lookupText) => lookupText.startsWith(prefix)).length;
}

export interface TermsFindResult {
  originalTextLength: number;
  dictionaryEntries: unknown[];
}

// One termsFind dictionary entry whose headword exactly matches `originalText`
// (the scanned surface, defaulting to the term). `dictionary` attaches a
// definition from that dictionary, e.g. a SubMiner character dictionary.
export function termEntry(
  term: string,
  reading: string,
  options: { originalText?: string; isPrimary?: boolean; dictionary?: string } = {},
) {
  return {
    headwords: [
      {
        term,
        reading,
        sources: [
          {
            originalText: options.originalText ?? term,
            isPrimary: options.isPrimary ?? true,
            matchType: 'exact',
          },
        ],
      },
    ],
    ...(options.dictionary ? { definitions: [{ dictionary: options.dictionary }] } : {}),
  };
}

export function termsFound(
  originalTextLength: number,
  ...dictionaryEntries: unknown[]
): TermsFindResult {
  return { originalTextLength, dictionaryEntries };
}

// Scan deps backed by a fake Yomitan backend. The active profile enables
// `dictionaries` (in priority order; omitted means no dictionary list),
// getDictionaryInfo returns `dictionaryInfo`, and termsFind answers through
// `termsFind` (null/undefined means no match) while recording each text in
// `lookups`. `actions` handles any other backend action and overrides the
// defaults above; anything else throws. Every action name lands in `actionLog`.
export function createBackendDeps(
  options: {
    dictionaries?: string[];
    dictionaryInfo?: unknown[];
    termsFind?: (text: string) => TermsFindResult | null | undefined;
    lookups?: string[];
    actions?: Record<string, (params: unknown) => unknown>;
    actionLog?: string[];
    onScript?: (script: string) => void;
  } = {},
) {
  return createScanDeps(
    (action, params) => {
      options.actionLog?.push(action);
      const override = options.actions?.[action];
      if (override) {
        return override(params);
      }
      if (action === 'optionsGetFull') {
        return {
          profileCurrent: 0,
          profiles: [
            {
              options: {
                scanning: { length: 40 },
                ...(options.dictionaries
                  ? {
                      dictionaries: options.dictionaries.map((name, id) => ({
                        name,
                        enabled: true,
                        id,
                      })),
                    }
                  : {}),
              },
            },
          ],
        };
      }
      if (action === 'getDictionaryInfo') {
        return options.dictionaryInfo ?? [];
      }
      if (action === 'termsFind' && options.termsFind) {
        const text = (params as { text?: string } | undefined)?.text ?? '';
        options.lookups?.push(text);
        return options.termsFind(text) ?? termsFound(0);
      }
      throw new Error(`unexpected action: ${action}`);
    },
    { onScript: options.onScript },
  );
}

// Deps whose hidden settings.html window runs injected scripts against a fake
// `__subminerYomitanSettingsAutomation` bridge, like the real settings page.
export function createSettingsAutomationDeps(automation: Record<string, unknown>) {
  const context = vm.createContext({
    __subminerYomitanSettingsAutomation: { ready: true, ...automation },
    setTimeout,
  });
  const settingsWindow = {
    isDestroyed: () => false,
    destroy: () => undefined,
    webContents: {
      executeJavaScript: async (script: string) => await vm.runInContext(script, context),
    },
  };
  return createDeps(async () => true, {
    createYomitanExtensionWindow: async (pageName) =>
      pageName === 'settings.html' ? settingsWindow : null,
  });
}

export const SUBMINER_TEST_CHARACTER_DICTIONARY = 'SubMiner Character Dictionary (AniList 1)';

// Backend stub for the greedy name pre-pass: one character name (ミナト) in a
// line of ordinary words, with the SubMiner character dictionary enabled
// unless `dictionaries` says otherwise.
export const NAME_SCAN_WORDS: Array<[string, string, string, boolean]> = [
  ['ミナト', 'ミナト', 'みなと', true],
  ['は', 'は', 'は', false],
  ['まだ', 'まだ', 'まだ', false],
  ['学校', '学校', 'がっこう', false],
  ['に', 'に', 'に', false],
  ['いない', 'いる', 'いる', false],
];

export function createNameScanDeps(
  lookups: string[],
  words: Array<[string, string, string, boolean]> = NAME_SCAN_WORDS,
  dictionaries: string[] = ['JMdict', SUBMINER_TEST_CHARACTER_DICTIONARY],
) {
  return createBackendDeps({
    dictionaries,
    lookups,
    termsFind: (text) => {
      const match = words.find(([surface]) => text.startsWith(surface));
      if (!match) {
        return null;
      }
      const [surface, term, reading, isName] = match;
      return termsFound(
        surface.length,
        termEntry(term, reading, {
          originalText: surface,
          dictionary: isName ? SUBMINER_TEST_CHARACTER_DICTIONARY : 'JMdict',
        }),
      );
    },
  });
}
