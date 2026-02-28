import type { BrowserWindow, Extension } from 'electron';

interface LoggerLike {
  error: (message: string, ...args: unknown[]) => void;
  info?: (message: string, ...args: unknown[]) => void;
}

interface YomitanParserRuntimeDeps {
  getYomitanExt: () => Extension | null;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
}

export interface YomitanTermFrequency {
  term: string;
  reading: string | null;
  dictionary: string;
  frequency: number;
  displayValue: string | null;
  displayValueParsed: boolean;
}

export interface YomitanTermReadingPair {
  term: string;
  reading: string | null;
}

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object');
}

function asPositiveInteger(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    return null;
  }
  return Math.max(1, Math.floor(value));
}

function toYomitanTermFrequency(value: unknown): YomitanTermFrequency | null {
  if (!isObject(value)) {
    return null;
  }

  const term = typeof value.term === 'string' ? value.term.trim() : '';
  const dictionary = typeof value.dictionary === 'string' ? value.dictionary.trim() : '';
  const frequency = asPositiveInteger(value.frequency);
  if (!term || !dictionary || frequency === null) {
    return null;
  }

  const reading =
    value.reading === null
      ? null
      : typeof value.reading === 'string'
        ? value.reading
        : null;
  const displayValue =
    value.displayValue === null
      ? null
      : typeof value.displayValue === 'string'
        ? value.displayValue
        : null;
  const displayValueParsed = value.displayValueParsed === true;

  return {
    term,
    reading,
    dictionary,
    frequency,
    displayValue,
    displayValueParsed,
  };
}

function normalizeTermReadingList(termReadingList: YomitanTermReadingPair[]): YomitanTermReadingPair[] {
  const normalized: YomitanTermReadingPair[] = [];
  const seen = new Set<string>();

  for (const pair of termReadingList) {
    const term = typeof pair.term === 'string' ? pair.term.trim() : '';
    if (!term) {
      continue;
    }
    const reading =
      typeof pair.reading === 'string' && pair.reading.trim().length > 0 ? pair.reading.trim() : null;
    const key = `${term}\u0000${reading ?? ''}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    normalized.push({ term, reading });
  }

  return normalized;
}

async function ensureYomitanParserWindow(
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<boolean> {
  const electron = await import('electron');
  const yomitanExt = deps.getYomitanExt();
  if (!yomitanExt) {
    return false;
  }

  const currentWindow = deps.getYomitanParserWindow();
  if (currentWindow && !currentWindow.isDestroyed()) {
    return true;
  }

  const existingInitPromise = deps.getYomitanParserInitPromise();
  if (existingInitPromise) {
    return existingInitPromise;
  }

  const initPromise = (async () => {
    const { BrowserWindow, session } = electron;
    const parserWindow = new BrowserWindow({
      show: false,
      width: 800,
      height: 600,
      webPreferences: {
        contextIsolation: true,
        nodeIntegration: false,
        session: session.defaultSession,
      },
    });
    deps.setYomitanParserWindow(parserWindow);

    deps.setYomitanParserReadyPromise(
      new Promise((resolve, reject) => {
        parserWindow.webContents.once('did-finish-load', () => resolve());
        parserWindow.webContents.once('did-fail-load', (_event, _errorCode, errorDescription) => {
          reject(new Error(errorDescription));
        });
      }),
    );

    parserWindow.on('closed', () => {
      if (deps.getYomitanParserWindow() === parserWindow) {
        deps.setYomitanParserWindow(null);
        deps.setYomitanParserReadyPromise(null);
      }
    });

    try {
      await parserWindow.loadURL(`chrome-extension://${yomitanExt.id}/search.html`);
      const readyPromise = deps.getYomitanParserReadyPromise();
      if (readyPromise) {
        await readyPromise;
      }

      return true;
    } catch (err) {
      logger.error('Failed to initialize Yomitan parser window:', (err as Error).message);
      if (!parserWindow.isDestroyed()) {
        parserWindow.destroy();
      }
      if (deps.getYomitanParserWindow() === parserWindow) {
        deps.setYomitanParserWindow(null);
        deps.setYomitanParserReadyPromise(null);
      }

      return false;
    } finally {
      deps.setYomitanParserInitPromise(null);
    }
  })();

  deps.setYomitanParserInitPromise(initPromise);
  return initPromise;
}

export async function requestYomitanParseResults(
  text: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<unknown[] | null> {
  const yomitanExt = deps.getYomitanExt();
  if (!text || !yomitanExt) {
    return null;
  }

  const isReady = await ensureYomitanParserWindow(deps, logger);
  const parserWindow = deps.getYomitanParserWindow();
  if (!isReady || !parserWindow || parserWindow.isDestroyed()) {
    return null;
  }

  const script = `
    (async () => {
      const invoke = (action, params) =>
        new Promise((resolve, reject) => {
          chrome.runtime.sendMessage({ action, params }, (response) => {
            if (chrome.runtime.lastError) {
              reject(new Error(chrome.runtime.lastError.message));
              return;
            }
            if (!response || typeof response !== "object") {
              reject(new Error("Invalid response from Yomitan backend"));
              return;
            }
            if (response.error) {
              reject(new Error(response.error.message || "Yomitan backend error"));
              return;
            }
            resolve(response.result);
          });
        });

      const optionsFull = await invoke("optionsGetFull", undefined);
      const profileIndex = optionsFull.profileCurrent;
      const scanLength =
        optionsFull.profiles?.[profileIndex]?.options?.scanning?.length ?? 40;

      return await invoke("parseText", {
        text: ${JSON.stringify(text)},
        optionsContext: { index: profileIndex },
        scanLength,
        useInternalParser: true,
        useMecabParser: true
      });
    })();
  `;

  try {
    const parseResults = await parserWindow.webContents.executeJavaScript(script, true);
    return Array.isArray(parseResults) ? parseResults : null;
  } catch (err) {
    logger.error('Yomitan parser request failed:', (err as Error).message);
    return null;
  }
}

export async function requestYomitanTermFrequencies(
  termReadingList: YomitanTermReadingPair[],
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<YomitanTermFrequency[]> {
  const normalizedTermReadingList = normalizeTermReadingList(termReadingList);
  const yomitanExt = deps.getYomitanExt();
  if (normalizedTermReadingList.length === 0 || !yomitanExt) {
    return [];
  }

  const isReady = await ensureYomitanParserWindow(deps, logger);
  const parserWindow = deps.getYomitanParserWindow();
  if (!isReady || !parserWindow || parserWindow.isDestroyed()) {
    return [];
  }

  const script = `
    (async () => {
      const invoke = (action, params) =>
        new Promise((resolve, reject) => {
          chrome.runtime.sendMessage({ action, params }, (response) => {
            if (chrome.runtime.lastError) {
              reject(new Error(chrome.runtime.lastError.message));
              return;
            }
            if (!response || typeof response !== "object") {
              reject(new Error("Invalid response from Yomitan backend"));
              return;
            }
            if (response.error) {
              reject(new Error(response.error.message || "Yomitan backend error"));
              return;
            }
            resolve(response.result);
          });
        });

      const optionsFull = await invoke("optionsGetFull", undefined);
      const profileIndex = optionsFull.profileCurrent;
      const dictionariesRaw = optionsFull.profiles?.[profileIndex]?.options?.dictionaries ?? [];
      const dictionaries = Array.isArray(dictionariesRaw)
        ? dictionariesRaw
            .filter((entry) => entry && typeof entry === "object" && entry.enabled === true && typeof entry.name === "string")
            .map((entry) => entry.name)
        : [];

      if (dictionaries.length === 0) {
        return [];
      }

      return await invoke("getTermFrequencies", {
        termReadingList: ${JSON.stringify(normalizedTermReadingList)},
        dictionaries
      });
    })();
  `;

  try {
    const rawResult = await parserWindow.webContents.executeJavaScript(script, true);
    if (!Array.isArray(rawResult)) {
      return [];
    }

    return rawResult
      .map((entry) => toYomitanTermFrequency(entry))
      .filter((entry): entry is YomitanTermFrequency => entry !== null);
  } catch (err) {
    logger.error('Yomitan term frequency request failed:', (err as Error).message);
    return [];
  }
}

export async function syncYomitanDefaultAnkiServer(
  serverUrl: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<boolean> {
  const normalizedTargetServer = serverUrl.trim();
  if (!normalizedTargetServer) {
    return false;
  }

  const isReady = await ensureYomitanParserWindow(deps, logger);
  const parserWindow = deps.getYomitanParserWindow();
  if (!isReady || !parserWindow || parserWindow.isDestroyed()) {
    return false;
  }

  const script = `
    (async () => {
      const invoke = (action, params) =>
        new Promise((resolve, reject) => {
          chrome.runtime.sendMessage({ action, params }, (response) => {
            if (chrome.runtime.lastError) {
              reject(new Error(chrome.runtime.lastError.message));
              return;
            }
            if (!response || typeof response !== "object") {
              reject(new Error("Invalid response from Yomitan backend"));
              return;
            }
            if (response.error) {
              reject(new Error(response.error.message || "Yomitan backend error"));
              return;
            }
            resolve(response.result);
          });
        });

      const targetServer = ${JSON.stringify(normalizedTargetServer)};
      const optionsFull = await invoke("optionsGetFull", undefined);
      const profiles = Array.isArray(optionsFull.profiles) ? optionsFull.profiles : [];
      if (profiles.length === 0) {
        return { updated: false, reason: "no-profiles" };
      }

      const defaultProfile = profiles[0];
      if (!defaultProfile || typeof defaultProfile !== "object") {
        return { updated: false, reason: "invalid-default-profile" };
      }

      defaultProfile.options = defaultProfile.options && typeof defaultProfile.options === "object"
        ? defaultProfile.options
        : {};
      defaultProfile.options.anki = defaultProfile.options.anki && typeof defaultProfile.options.anki === "object"
        ? defaultProfile.options.anki
        : {};

      const currentServerRaw = defaultProfile.options.anki.server;
      const currentServer = typeof currentServerRaw === "string" ? currentServerRaw.trim() : "";
      const canReplaceDefault =
        currentServer.length === 0 || currentServer === "http://127.0.0.1:8765";
      if (!canReplaceDefault || currentServer === targetServer) {
        return { updated: false, reason: "no-change", currentServer, targetServer };
      }

      defaultProfile.options.anki.server = targetServer;
      await invoke("setAllSettings", { value: optionsFull, source: "subminer" });
      return { updated: true, currentServer, targetServer };
    })();
  `;

  try {
    const result = await parserWindow.webContents.executeJavaScript(script, true);
    const updated =
      typeof result === 'object' &&
      result !== null &&
      (result as { updated?: unknown }).updated === true;
    if (updated) {
      logger.info?.(`Updated Yomitan default profile Anki server to ${normalizedTargetServer}`);
      return true;
    }
    return false;
  } catch (err) {
    logger.error('Failed to sync Yomitan default profile Anki server:', (err as Error).message);
    return false;
  }
}
