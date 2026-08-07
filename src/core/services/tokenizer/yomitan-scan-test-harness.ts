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
  return {
    chrome: {
      runtime: {
        lastError: null,
        sendMessage: (
          payload: { action?: string; params?: unknown },
          callback: (response: { result?: unknown; error?: { message?: string } }) => void,
        ) => {
          try {
            callback({ result: handler(payload.action ?? '', payload.params) });
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
  return await vm.runInNewContext(script, createYomitanScriptSandbox(handler));
}

// Persistent page context shared across executeJavaScript calls, matching the
// real parser window: the scan runtime is installed once via
// globalThis.__subminerYomitanScan and per-line calls reuse it (and its
// cross-line termsFind cache).
function createPersistentYomitanScriptRunner(
  handler: (action: string, params: unknown) => unknown,
): (script: string) => Promise<unknown> {
  const context = vm.createContext(createYomitanScriptSandbox(handler));
  return async (script: string) => await vm.runInContext(script, context);
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

// Backend stub for the greedy name pre-pass: one character name (ミナト) in a
// line of ordinary words, with the SubMiner character dictionary enabled.
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
) {
  return createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JMdict', enabled: true, id: 0 },
                {
                  name: 'SubMiner Character Dictionary (AniList 1)',
                  enabled: true,
                  id: 1,
                },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    lookups.push(text);
    for (const [surface, term, reading, isName] of words) {
      if (text.startsWith(surface)) {
        return {
          originalTextLength: surface.length,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term,
                  reading,
                  sources: [{ originalText: surface, isPrimary: true, matchType: 'exact' }],
                },
              ],
              definitions: [
                { dictionary: isName ? 'SubMiner Character Dictionary (AniList 1)' : 'JMdict' },
              ],
            },
          ],
        };
      }
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });
}
