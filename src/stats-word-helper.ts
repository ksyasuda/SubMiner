import fs from 'node:fs';
import path from 'node:path';
import { app, protocol } from 'electron';
import type { BrowserWindow, Extension, Session } from 'electron';
import { ConfigService } from './config/service';
import { createLogger, setLogLevel } from './logger';
import { loadYomitanExtension } from './core/services/yomitan-extension-loader';
import {
  createHachidoriExtensionRuntime,
  getHachidoriSession,
} from './core/services/hachidori-extension';
import {
  getPreferredYomitanAnkiServerUrl,
  shouldForceOverrideYomitanAnkiServer,
} from './main/runtime/yomitan-anki-server';
import {
  addYomitanNoteViaSearch,
  getYomitanCurrentAnkiDeckName,
  syncYomitanDefaultAnkiServer,
} from './core/services/tokenizer/yomitan-parser-runtime';
import type { StatsWordHelperResponse } from './stats-word-helper-client';
import { clearYomitanExtensionRuntimeState } from './core/services/yomitan-extension-runtime-state';

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'chrome-extension',
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: true,
      bypassCSP: true,
    },
  },
]);

const logger = createLogger('stats-word-helper');

function readFlagValue(argv: string[], flag: string): string | undefined {
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg) continue;
    if (arg === flag) {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        return value;
      }
      return undefined;
    }
    if (arg.startsWith(`${flag}=`)) {
      return arg.split('=', 2)[1];
    }
  }
  return undefined;
}

function writeResponse(responsePath: string | undefined, payload: StatsWordHelperResponse): void {
  if (!responsePath) return;
  fs.mkdirSync(path.dirname(responsePath), { recursive: true });
  fs.writeFileSync(responsePath, JSON.stringify(payload, null, 2), 'utf8');
}

const responsePath = readFlagValue(process.argv, '--stats-word-helper-response-path')?.trim();
const userDataPath = readFlagValue(process.argv, '--stats-word-helper-user-data-path')?.trim();
const word = readFlagValue(process.argv, '--stats-word-helper-word');
const readDeck = process.argv.includes('--stats-word-helper-read-deck');
const logLevel = readFlagValue(process.argv, '--log-level');

if (logLevel) {
  setLogLevel(logLevel, 'cli');
}

if (!userDataPath || (!word && !readDeck)) {
  writeResponse(responsePath, {
    ok: false,
    error: 'Missing stats word helper arguments.',
  });
  app.exit(1);
}

app.setName('SubMiner');
app.setPath('userData', userDataPath!);

let yomitanExt: Extension | null = null;
let yomitanSession: Session | null = null;
let yomitanParserWindow: BrowserWindow | null = null;
let yomitanParserReadyPromise: Promise<void> | null = null;
let yomitanParserInitPromise: Promise<boolean> | null = null;

function cleanup(): void {
  clearYomitanExtensionRuntimeState({
    getYomitanParserWindow: () => yomitanParserWindow,
    setYomitanParserWindow: () => {
      yomitanParserWindow = null;
    },
    setYomitanParserReadyPromise: () => {
      yomitanParserReadyPromise = null;
    },
    setYomitanParserInitPromise: () => {
      yomitanParserInitPromise = null;
    },
    setYomitanExtension: () => {
      yomitanExt = null;
    },
    setYomitanSession: () => {
      yomitanSession = null;
    },
  });
}

async function main(): Promise<void> {
  try {
    const configService = new ConfigService(userDataPath!);
    const config = configService.getConfig();
    // Mine with the same backend the app uses so notes come from the user's
    // dictionaries and Anki templates, not an empty profile.
    const extension =
      config.dictionaryBackend === 'hachidori'
        ? await createHachidoriExtensionRuntime(userDataPath!)
            .ensureLoaded()
            .then((loaded) => {
              yomitanExt = loaded;
              yomitanSession = getHachidoriSession();
              return loaded;
            })
        : await loadYomitanExtension({
            userDataPath: userDataPath!,
            getYomitanParserWindow: () => yomitanParserWindow,
            setYomitanParserWindow: (window) => {
              yomitanParserWindow = window;
            },
            setYomitanParserReadyPromise: (promise) => {
              yomitanParserReadyPromise = promise;
            },
            setYomitanParserInitPromise: (promise) => {
              yomitanParserInitPromise = promise;
            },
            setYomitanExtension: (extensionValue) => {
              yomitanExt = extensionValue;
            },
            setYomitanSession: (sessionValue) => {
              yomitanSession = sessionValue;
            },
          });
    if (!extension) {
      throw new Error(`${config.dictionaryBackend} extension failed to load.`);
    }

    const yomitanDeps = {
      getYomitanExt: () => yomitanExt,
      getYomitanSession: () => yomitanSession,
      getYomitanParserWindow: () => yomitanParserWindow,
      setYomitanParserWindow: (window: BrowserWindow | null) => {
        yomitanParserWindow = window;
      },
      getYomitanParserReadyPromise: () => yomitanParserReadyPromise,
      setYomitanParserReadyPromise: (promise: Promise<void> | null) => {
        yomitanParserReadyPromise = promise;
      },
      getYomitanParserInitPromise: () => yomitanParserInitPromise,
      setYomitanParserInitPromise: (promise: Promise<boolean> | null) => {
        yomitanParserInitPromise = promise;
      },
    };

    if (readDeck) {
      const deckName = await getYomitanCurrentAnkiDeckName(yomitanDeps, logger);
      writeResponse(responsePath, {
        ok: true,
        deckName,
      });
      cleanup();
      app.exit(0);
      return;
    }

    await syncYomitanDefaultAnkiServer(
      getPreferredYomitanAnkiServerUrl(config.ankiConnect),
      yomitanDeps,
      logger,
      {
        forceOverride: shouldForceOverrideYomitanAnkiServer(config.ankiConnect),
        deck: config.ankiConnect?.deck,
        ankiConfig: config.ankiConnect,
      },
    );

    const addResult = await addYomitanNoteViaSearch(word!, yomitanDeps, logger);

    const noteId = addResult.noteId;
    if (typeof noteId !== 'number') {
      throw new Error(`${config.dictionaryBackend} failed to create note.`);
    }

    writeResponse(responsePath, {
      ok: true,
      noteId,
    });
    cleanup();
    app.exit(0);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('Stats word helper failed', message);
    writeResponse(responsePath, {
      ok: false,
      error: message,
    });
    cleanup();
    app.exit(1);
  }
}

void app.whenReady().then(() => main());
