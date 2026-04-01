import fs from 'node:fs';
import path from 'node:path';
import { app, protocol } from 'electron';
import type { BrowserWindow, Extension, Session } from 'electron';
import { ConfigService } from './config/service';
import { createLogger, setLogLevel } from './logger';
import { loadYomitanExtension } from './core/services/yomitan-extension-loader';
import {
  addYomitanNoteViaSearch,
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
const logLevel = readFlagValue(process.argv, '--log-level');

if (logLevel) {
  setLogLevel(logLevel, 'cli');
}

if (!userDataPath || !word) {
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
    const extension = await loadYomitanExtension({
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
      throw new Error('Yomitan extension failed to load.');
    }

    await syncYomitanDefaultAnkiServer(
      config.ankiConnect?.url || 'http://127.0.0.1:8765',
      {
        getYomitanExt: () => yomitanExt,
        getYomitanSession: () => yomitanSession,
        getYomitanParserWindow: () => yomitanParserWindow,
        setYomitanParserWindow: (window) => {
          yomitanParserWindow = window;
        },
        getYomitanParserReadyPromise: () => yomitanParserReadyPromise,
        setYomitanParserReadyPromise: (promise) => {
          yomitanParserReadyPromise = promise;
        },
        getYomitanParserInitPromise: () => yomitanParserInitPromise,
        setYomitanParserInitPromise: (promise) => {
          yomitanParserInitPromise = promise;
        },
      },
      logger,
      { forceOverride: true },
    );

    const addResult = await addYomitanNoteViaSearch(
      word!,
      {
        getYomitanExt: () => yomitanExt,
        getYomitanSession: () => yomitanSession,
        getYomitanParserWindow: () => yomitanParserWindow,
        setYomitanParserWindow: (window) => {
          yomitanParserWindow = window;
        },
        getYomitanParserReadyPromise: () => yomitanParserReadyPromise,
        setYomitanParserReadyPromise: (promise) => {
          yomitanParserReadyPromise = promise;
        },
        getYomitanParserInitPromise: () => yomitanParserInitPromise,
        setYomitanParserInitPromise: (promise) => {
          yomitanParserInitPromise = promise;
        },
      },
      logger,
    );

    const noteId = addResult.noteId;
    if (typeof noteId !== 'number') {
      throw new Error('Yomitan failed to create note.');
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
