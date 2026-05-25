import type { BrowserWindow, Extension, Session } from 'electron';
import * as fs from 'fs';
import * as http from 'http';
import * as path from 'path';
import { selectYomitanParseTokens } from './parser-selection-stage';

interface LoggerLike {
  error: (message: string, ...args: unknown[]) => void;
  info?: (message: string, ...args: unknown[]) => void;
}

interface YomitanParserRuntimeDeps {
  getYomitanExt: () => Extension | null;
  getYomitanSession?: () => Session | null;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  createYomitanExtensionWindow?: (pageName: string) => Promise<BrowserWindow | null>;
}

type YomitanFrequencyMode = 'occurrence-based' | 'rank-based';

export interface YomitanDictionaryInfo {
  title: string;
  revision?: string | number;
  frequencyMode?: YomitanFrequencyMode;
}

export interface YomitanTermFrequency {
  term: string;
  reading: string | null;
  hasReading: boolean;
  dictionary: string;
  dictionaryPriority: number;
  frequency: number;
  displayValue: string | null;
  displayValueParsed: boolean;
  frequencyDerivedFromDisplayValue: boolean;
}

export interface YomitanTermReadingPair {
  term: string;
  reading: string | null;
}

export interface YomitanScanToken {
  surface: string;
  reading: string;
  headword: string;
  startPos: number;
  endPos: number;
  isNameMatch?: boolean;
  frequencyRank?: number;
  wordClasses?: string[];
}

interface YomitanProfileMetadata {
  profileIndex: number;
  scanLength: number;
  dictionaries: string[];
  dictionaryPriorityByName: Record<string, number>;
  dictionaryFrequencyModeByName: Partial<Record<string, YomitanFrequencyMode>>;
}

export interface YomitanAddNoteResult {
  noteId: number | null;
  duplicateNoteIds: number[];
}

const DEFAULT_YOMITAN_SCAN_LENGTH = 40;
const yomitanProfileMetadataByWindow = new WeakMap<BrowserWindow, YomitanProfileMetadata>();
const yomitanFrequencyCacheByWindow = new WeakMap<
  BrowserWindow,
  Map<string, YomitanTermFrequency[]>
>();

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object');
}

function isScanTokenArray(value: unknown): value is YomitanScanToken[] {
  return (
    Array.isArray(value) &&
    value.every(
      (entry) =>
        isObject(entry) &&
        typeof entry.surface === 'string' &&
        typeof entry.reading === 'string' &&
        typeof entry.headword === 'string' &&
        typeof entry.startPos === 'number' &&
        typeof entry.endPos === 'number' &&
        (entry.isNameMatch === undefined || typeof entry.isNameMatch === 'boolean') &&
        (entry.frequencyRank === undefined || typeof entry.frequencyRank === 'number') &&
        (entry.wordClasses === undefined ||
          (Array.isArray(entry.wordClasses) &&
            entry.wordClasses.every((wordClass) => typeof wordClass === 'string'))),
    )
  );
}

function hasSameTokenSpans(left: YomitanScanToken[], right: YomitanScanToken[]): boolean {
  if (left.length !== right.length) {
    return false;
  }

  return left.every((token, index) => {
    const other = right[index];
    return (
      other !== undefined &&
      token.surface === other.surface &&
      token.startPos === other.startPos &&
      token.endPos === other.endPos
    );
  });
}

function makeTermReadingCacheKey(term: string, reading: string | null): string {
  return `${term}\u0000${reading ?? ''}`;
}

function getWindowFrequencyCache(window: BrowserWindow): Map<string, YomitanTermFrequency[]> {
  let cache = yomitanFrequencyCacheByWindow.get(window);
  if (!cache) {
    cache = new Map<string, YomitanTermFrequency[]>();
    yomitanFrequencyCacheByWindow.set(window, cache);
  }
  return cache;
}

function clearWindowCaches(window: BrowserWindow): void {
  yomitanProfileMetadataByWindow.delete(window);
  yomitanFrequencyCacheByWindow.delete(window);
}
export function clearYomitanParserCachesForWindow(window: BrowserWindow): void {
  clearWindowCaches(window);
}

function asPositiveInteger(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    return null;
  }
  return Math.max(1, Math.floor(value));
}

function parsePositiveFrequencyString(value: string): number | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  const numericMatch = trimmed.match(/[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/)?.[0];
  if (!numericMatch) {
    return null;
  }

  const parsed = Number.parseFloat(numericMatch);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }

  const normalized = Math.floor(parsed);
  if (!Number.isFinite(normalized) || normalized <= 0) {
    return null;
  }

  return normalized;
}

function parsePositiveFrequencyValue(value: unknown): number | null {
  const numeric = asPositiveInteger(value);
  if (numeric !== null) {
    return numeric;
  }

  if (typeof value === 'string') {
    return parsePositiveFrequencyString(value);
  }

  if (Array.isArray(value)) {
    for (const item of value) {
      const parsed = parsePositiveFrequencyValue(item);
      if (parsed !== null) {
        return parsed;
      }
    }
  }

  return null;
}

function parseDisplayFrequencyValue(value: unknown): number | null {
  if (typeof value === 'string') {
    const leadingDigits = value.trim().match(/^\d+/)?.[0];
    if (!leadingDigits) {
      return null;
    }
    const parsed = Number.parseInt(leadingDigits, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
  }

  return parsePositiveFrequencyValue(value);
}

function toYomitanTermFrequency(value: unknown): YomitanTermFrequency | null {
  if (!isObject(value)) {
    return null;
  }

  const term = typeof value.term === 'string' ? value.term.trim() : '';
  const dictionary = typeof value.dictionary === 'string' ? value.dictionary.trim() : '';
  const rawFrequency = parsePositiveFrequencyValue(value.frequency);
  const displayValueRaw = value.displayValue;
  const parsedDisplayFrequency =
    displayValueRaw !== null && displayValueRaw !== undefined
      ? parseDisplayFrequencyValue(displayValueRaw)
      : null;
  const frequency = parsedDisplayFrequency ?? rawFrequency;
  if (!term || !dictionary || frequency === null) {
    return null;
  }
  const dictionaryPriorityRaw = (value as { dictionaryPriority?: unknown }).dictionaryPriority;
  const dictionaryPriority =
    typeof dictionaryPriorityRaw === 'number' && Number.isFinite(dictionaryPriorityRaw)
      ? Math.max(0, Math.floor(dictionaryPriorityRaw))
      : Number.MAX_SAFE_INTEGER;

  const reading =
    value.reading === null ? null : typeof value.reading === 'string' ? value.reading : null;
  const hasReading = value.hasReading === false ? false : reading !== null;
  const displayValue = typeof displayValueRaw === 'string' ? displayValueRaw : null;
  const displayValueParsed = value.displayValueParsed === true;

  return {
    term,
    reading,
    hasReading,
    dictionary,
    dictionaryPriority,
    frequency,
    displayValue,
    displayValueParsed,
    frequencyDerivedFromDisplayValue: parsedDisplayFrequency !== null,
  };
}

function normalizeTermReadingList(
  termReadingList: YomitanTermReadingPair[],
): YomitanTermReadingPair[] {
  const normalized: YomitanTermReadingPair[] = [];
  const seen = new Set<string>();

  for (const pair of termReadingList) {
    const term = typeof pair.term === 'string' ? pair.term.trim() : '';
    if (!term) {
      continue;
    }
    const reading =
      typeof pair.reading === 'string' && pair.reading.trim().length > 0
        ? pair.reading.trim()
        : null;
    const key = `${term}\u0000${reading ?? ''}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    normalized.push({ term, reading });
  }

  return normalized;
}

function toYomitanProfileMetadata(value: unknown): YomitanProfileMetadata | null {
  if (!isObject(value)) {
    return null;
  }

  const profileIndexRaw = value.profileIndex ?? value.profileCurrent;
  const profileIndex =
    typeof profileIndexRaw === 'number' && Number.isFinite(profileIndexRaw)
      ? Math.max(0, Math.floor(profileIndexRaw))
      : 0;
  const scanLengthRaw =
    value.scanLength ??
    (Array.isArray(value.profiles) && isObject(value.profiles[profileIndex])
      ? (value.profiles[profileIndex] as { options?: { scanning?: { length?: unknown } } }).options
          ?.scanning?.length
      : undefined);
  const scanLength =
    typeof scanLengthRaw === 'number' && Number.isFinite(scanLengthRaw)
      ? Math.max(1, Math.floor(scanLengthRaw))
      : DEFAULT_YOMITAN_SCAN_LENGTH;
  const dictionariesRaw =
    value.dictionaries ??
    (Array.isArray(value.profiles) && isObject(value.profiles[profileIndex])
      ? (value.profiles[profileIndex] as { options?: { dictionaries?: unknown[] } }).options
          ?.dictionaries
      : undefined);
  const dictionaries = Array.isArray(dictionariesRaw)
    ? dictionariesRaw
        .map((entry, index) => {
          if (typeof entry === 'string') {
            return { name: entry.trim(), priority: index };
          }
          if (!isObject(entry) || entry.enabled === false || typeof entry.name !== 'string') {
            return null;
          }
          const normalizedName = entry.name.trim();
          if (!normalizedName) {
            return null;
          }
          const priorityRaw = (entry as { id?: unknown }).id;
          const priority =
            typeof priorityRaw === 'number' && Number.isFinite(priorityRaw)
              ? Math.max(0, Math.floor(priorityRaw))
              : index;
          return { name: normalizedName, priority };
        })
        .filter((entry): entry is { name: string; priority: number } => entry !== null)
        .sort((a, b) => a.priority - b.priority)
        .map((entry) => entry.name)
        .filter((entry) => entry.length > 0)
    : [];
  const dictionaryPriorityByNameRaw = value.dictionaryPriorityByName;
  const dictionaryPriorityByName: Record<string, number> = {};
  if (isObject(dictionaryPriorityByNameRaw)) {
    for (const [name, priorityRaw] of Object.entries(dictionaryPriorityByNameRaw)) {
      if (typeof priorityRaw !== 'number' || !Number.isFinite(priorityRaw)) {
        continue;
      }
      const normalizedName = name.trim();
      if (!normalizedName) {
        continue;
      }
      dictionaryPriorityByName[normalizedName] = Math.max(0, Math.floor(priorityRaw));
    }
  }

  for (let index = 0; index < dictionaries.length; index += 1) {
    const dictionary = dictionaries[index];
    if (!dictionary) {
      continue;
    }
    if (dictionaryPriorityByName[dictionary] === undefined) {
      dictionaryPriorityByName[dictionary] = index;
    }
  }

  const dictionaryFrequencyModeByNameRaw = value.dictionaryFrequencyModeByName;
  const dictionaryFrequencyModeByName: Partial<Record<string, YomitanFrequencyMode>> = {};
  if (isObject(dictionaryFrequencyModeByNameRaw)) {
    for (const [name, frequencyModeRaw] of Object.entries(dictionaryFrequencyModeByNameRaw)) {
      const normalizedName = name.trim();
      if (!normalizedName) {
        continue;
      }
      if (frequencyModeRaw !== 'occurrence-based' && frequencyModeRaw !== 'rank-based') {
        continue;
      }
      dictionaryFrequencyModeByName[normalizedName] = frequencyModeRaw;
    }
  }

  return {
    profileIndex,
    scanLength,
    dictionaries,
    dictionaryPriorityByName,
    dictionaryFrequencyModeByName,
  };
}

function normalizeFrequencyEntriesWithPriority(
  rawResult: unknown[],
  dictionaryPriorityByName: Record<string, number>,
  dictionaryFrequencyModeByName: Partial<Record<string, YomitanFrequencyMode>>,
): YomitanTermFrequency[] {
  const normalized: YomitanTermFrequency[] = [];
  for (const entry of rawResult) {
    const frequency = toYomitanTermFrequency(entry);
    if (!frequency) {
      continue;
    }

    if (dictionaryFrequencyModeByName[frequency.dictionary] === 'occurrence-based') {
      continue;
    }

    const dictionaryPriority = dictionaryPriorityByName[frequency.dictionary];
    normalized.push({
      ...frequency,
      dictionaryPriority:
        dictionaryPriority !== undefined ? dictionaryPriority : frequency.dictionaryPriority,
    });
  }

  return normalized;
}

function groupFrequencyEntriesByPair(
  entries: YomitanTermFrequency[],
): Map<string, YomitanTermFrequency[]> {
  const grouped = new Map<string, YomitanTermFrequency[]>();
  for (const entry of entries) {
    const reading =
      typeof entry.reading === 'string' && entry.reading.trim().length > 0
        ? entry.reading.trim()
        : null;
    const key = makeTermReadingCacheKey(entry.term.trim(), reading);
    const existing = grouped.get(key);
    if (existing) {
      existing.push(entry);
      continue;
    }
    grouped.set(key, [entry]);
  }
  return grouped;
}

function groupFrequencyEntriesByTerm(
  entries: YomitanTermFrequency[],
): Map<string, YomitanTermFrequency[]> {
  const grouped = new Map<string, YomitanTermFrequency[]>();
  for (const entry of entries) {
    const term = entry.term.trim();
    if (!term) {
      continue;
    }

    const existing = grouped.get(term);
    if (existing) {
      existing.push(entry);
      continue;
    }
    grouped.set(term, [entry]);
  }
  return grouped;
}

async function requestYomitanProfileMetadata(
  parserWindow: BrowserWindow,
  logger: LoggerLike,
): Promise<YomitanProfileMetadata | null> {
  const cached = yomitanProfileMetadataByWindow.get(parserWindow);
  if (cached) {
    return cached;
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
      const profileIndex =
        typeof optionsFull.profileCurrent === "number" && Number.isFinite(optionsFull.profileCurrent)
          ? Math.max(0, Math.floor(optionsFull.profileCurrent))
          : 0;
      const scanLengthRaw = optionsFull.profiles?.[profileIndex]?.options?.scanning?.length;
      const scanLength =
        typeof scanLengthRaw === "number" && Number.isFinite(scanLengthRaw)
          ? Math.max(1, Math.floor(scanLengthRaw))
          : ${DEFAULT_YOMITAN_SCAN_LENGTH};
      const dictionariesRaw = optionsFull.profiles?.[profileIndex]?.options?.dictionaries ?? [];
      const dictionaryEntries = Array.isArray(dictionariesRaw)
        ? dictionariesRaw
            .filter((entry) => entry && typeof entry === "object" && entry.enabled === true && typeof entry.name === "string")
            .map((entry, index) => ({
              name: entry.name,
              id: typeof entry.id === "number" && Number.isFinite(entry.id) ? Math.max(0, Math.floor(entry.id)) : index
            }))
            .sort((a, b) => a.id - b.id)
        : [];
      const dictionaries = dictionaryEntries.map((entry) => entry.name);
      const dictionaryPriorityByName = dictionaryEntries.reduce((acc, entry, index) => {
        acc[entry.name] = index;
        return acc;
      }, {});
      let dictionaryFrequencyModeByName = {};
      try {
        const dictionaryInfo = await invoke("getDictionaryInfo", undefined);
        dictionaryFrequencyModeByName = Array.isArray(dictionaryInfo)
          ? dictionaryInfo.reduce((acc, entry) => {
              if (!entry || typeof entry !== "object" || typeof entry.title !== "string") {
                return acc;
              }
              if (
                entry.frequencyMode === "occurrence-based" ||
                entry.frequencyMode === "rank-based"
              ) {
                acc[entry.title] = entry.frequencyMode;
              }
              return acc;
            }, {})
          : {};
      } catch {
        dictionaryFrequencyModeByName = {};
      }

      return {
        profileIndex,
        scanLength,
        dictionaries,
        dictionaryPriorityByName,
        dictionaryFrequencyModeByName
      };
    })();
  `;

  try {
    const rawMetadata = await parserWindow.webContents.executeJavaScript(script, true);
    const metadata = toYomitanProfileMetadata(rawMetadata);
    if (!metadata) {
      return null;
    }
    yomitanProfileMetadataByWindow.set(parserWindow, metadata);
    return metadata;
  } catch (err) {
    logger.error('Yomitan parser metadata request failed:', (err as Error).message);
    return null;
  }
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
    const yomitanSession = deps.getYomitanSession?.() ?? session.defaultSession;
    const parserWindow = new BrowserWindow({
      show: false,
      width: 800,
      height: 600,
      webPreferences: {
        contextIsolation: true,
        nodeIntegration: false,
        session: yomitanSession,
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
      clearWindowCaches(parserWindow);
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
      clearWindowCaches(parserWindow);
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

async function createYomitanExtensionWindow(
  pageName: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<BrowserWindow | null> {
  if (typeof deps.createYomitanExtensionWindow === 'function') {
    return await deps.createYomitanExtensionWindow(pageName);
  }

  const electron = await import('electron');
  const yomitanExt = deps.getYomitanExt();
  if (!yomitanExt) {
    return null;
  }

  const { BrowserWindow, session } = electron;
  const yomitanSession = deps.getYomitanSession?.() ?? session.defaultSession;
  const window = new BrowserWindow({
    show: false,
    width: 1200,
    height: 800,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      session: yomitanSession,
    },
  });

  try {
    await new Promise<void>((resolve, reject) => {
      window.webContents.once('did-finish-load', () => resolve());
      window.webContents.once('did-fail-load', (_event, _errorCode, errorDescription) => {
        reject(new Error(errorDescription));
      });
      void window
        .loadURL(`chrome-extension://${yomitanExt.id}/${pageName}`)
        .catch((error: Error) => reject(error));
    });
    return window;
  } catch (err) {
    logger.error(`Failed to create hidden Yomitan ${pageName} window: ${(err as Error).message}`);
    if (!window.isDestroyed()) {
      window.destroy();
    }
    return null;
  }
}

async function invokeYomitanSettingsAutomation<T>(
  script: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<T | null> {
  const settingsWindow = await createYomitanExtensionWindow('settings.html', deps, logger);
  if (!settingsWindow || settingsWindow.isDestroyed()) {
    return null;
  }

  try {
    await settingsWindow.webContents.executeJavaScript(
      `
        (async () => {
          const deadline = Date.now() + 10000;
          while (Date.now() < deadline) {
            if (globalThis.__subminerYomitanSettingsAutomation?.ready === true) {
              return true;
            }
            await new Promise((resolve) => setTimeout(resolve, 50));
          }
          throw new Error("Yomitan settings automation bridge did not become ready");
        })();
      `,
      true,
    );

    return (await settingsWindow.webContents.executeJavaScript(script, true)) as T;
  } catch (err) {
    logger.error('Failed to drive Yomitan settings automation:', (err as Error).message);
    return null;
  } finally {
    if (!settingsWindow.isDestroyed()) {
      settingsWindow.destroy();
    }
  }
}

async function serveDictionaryZipOnce<T>(
  zipPath: string,
  callback: (url: string) => Promise<T>,
): Promise<T> {
  const fileName = path.basename(zipPath);
  const token = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  const requestPath = `/${token}/${encodeURIComponent(fileName)}`;
  let served = false;
  const server = http.createServer((request, response) => {
    if (request.method === 'OPTIONS') {
      response.writeHead(204, {
        'access-control-allow-origin': '*',
        'access-control-allow-methods': 'GET, OPTIONS',
      });
      response.end();
      return;
    }
    if (request.method !== 'GET' || request.url !== requestPath || served) {
      response.writeHead(404, { 'access-control-allow-origin': '*' });
      response.end();
      return;
    }
    served = true;
    let size = 0;
    try {
      size = fs.statSync(zipPath).size;
    } catch {
      response.writeHead(500, { 'access-control-allow-origin': '*' });
      response.end();
      return;
    }
    response.writeHead(200, {
      'access-control-allow-origin': '*',
      'content-length': String(size),
      'content-type': 'application/zip',
    });
    const stream = fs.createReadStream(zipPath);
    stream.on('error', () => {
      if (!response.headersSent) {
        response.writeHead(500, { 'access-control-allow-origin': '*' });
        response.end();
        return;
      }
      response.destroy();
    });
    stream.pipe(response);
  });

  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => resolve());
  });

  try {
    const address = server.address();
    if (!address || typeof address === 'string') {
      throw new Error('Dictionary import server did not bind to a TCP port.');
    }
    return await callback(`http://127.0.0.1:${address.port}${requestPath}`);
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
}

const YOMITAN_SCANNING_HELPERS = String.raw`
      const HIRAGANA_CONVERSION_RANGE = [0x3041, 0x3096];
      const KATAKANA_CONVERSION_RANGE = [0x30a1, 0x30f6];
      const KANA_PROLONGED_SOUND_MARK_CODE_POINT = 0x30fc;
      const KATAKANA_SMALL_KA_CODE_POINT = 0x30f5;
      const KATAKANA_SMALL_KE_CODE_POINT = 0x30f6;
      const KANA_RANGES = [[0x3040, 0x309f], [0x30a0, 0x30ff]];
      const JAPANESE_RANGES = [[0x3040, 0x30ff], [0x3400, 0x9fff]];
      function isCodePointInRange(codePoint, range) { return codePoint >= range[0] && codePoint <= range[1]; }
      function isCodePointInRanges(codePoint, ranges) { return ranges.some((range) => isCodePointInRange(codePoint, range)); }
      function isCodePointKana(codePoint) { return isCodePointInRanges(codePoint, KANA_RANGES); }
      function isCodePointJapanese(codePoint) { return isCodePointInRanges(codePoint, JAPANESE_RANGES); }
      function createFuriganaSegment(text, reading) { return {text, reading}; }
      function getProlongedHiragana(previousCharacter) {
        switch (previousCharacter) {
          case "あ": case "か": case "が": case "さ": case "ざ": case "た": case "だ": case "な": case "は": case "ば": case "ぱ": case "ま": case "や": case "ら": case "わ": case "ぁ": case "ゃ": case "ゎ": return "あ";
          case "い": case "き": case "ぎ": case "し": case "じ": case "ち": case "ぢ": case "に": case "ひ": case "び": case "ぴ": case "み": case "り": case "ぃ": return "い";
          case "う": case "く": case "ぐ": case "す": case "ず": case "つ": case "づ": case "ぬ": case "ふ": case "ぶ": case "ぷ": case "む": case "ゆ": case "る": case "ぅ": case "ゅ": return "う";
          case "え": case "け": case "げ": case "せ": case "ぜ": case "て": case "で": case "ね": case "へ": case "べ": case "ぺ": case "め": case "れ": case "ぇ": return "え";
          case "お": case "こ": case "ご": case "そ": case "ぞ": case "と": case "ど": case "の": case "ほ": case "ぼ": case "ぽ": case "も": case "よ": case "ろ": case "を": case "ぉ": case "ょ": return "う";
          default: return null;
        }
      }
      function getFuriganaKanaSegments(text, reading) {
        const newSegments = [];
        let start = 0;
        let state = (reading[0] === text[0]);
        for (let i = 1; i < text.length; ++i) {
          const newState = (reading[i] === text[i]);
          if (state === newState) { continue; }
          newSegments.push(createFuriganaSegment(text.substring(start, i), state ? '' : reading.substring(start, i)));
          state = newState;
          start = i;
        }
        newSegments.push(createFuriganaSegment(text.substring(start), state ? '' : reading.substring(start)));
        return newSegments;
      }
      function convertKatakanaToHiragana(text, keepProlongedSoundMarks = false) {
        let result = '';
        const offset = (HIRAGANA_CONVERSION_RANGE[0] - KATAKANA_CONVERSION_RANGE[0]);
        for (let char of text) {
          const codePoint = char.codePointAt(0);
          switch (codePoint) {
            case KATAKANA_SMALL_KA_CODE_POINT:
            case KATAKANA_SMALL_KE_CODE_POINT:
              break;
            case KANA_PROLONGED_SOUND_MARK_CODE_POINT:
              if (!keepProlongedSoundMarks && result.length > 0) {
                const char2 = getProlongedHiragana(result[result.length - 1]);
                if (char2 !== null) { char = char2; }
              }
              break;
            default:
              if (isCodePointInRange(codePoint, KATAKANA_CONVERSION_RANGE)) {
                char = String.fromCodePoint(codePoint + offset);
              }
              break;
          }
          result += char;
        }
        return result;
      }
      function segmentizeFurigana(reading, readingNormalized, groups, groupsStart) {
        const groupCount = groups.length - groupsStart;
        if (groupCount <= 0) { return reading.length === 0 ? [] : null; }
        const group = groups[groupsStart];
        const {isKana, text} = group;
        if (isKana) {
          if (group.textNormalized !== null && readingNormalized.startsWith(group.textNormalized)) {
            const segments = segmentizeFurigana(reading.substring(text.length), readingNormalized.substring(text.length), groups, groupsStart + 1);
            if (segments !== null) {
              if (reading.startsWith(text)) { segments.unshift(createFuriganaSegment(text, '')); }
              else { segments.unshift(...getFuriganaKanaSegments(text, reading)); }
              return segments;
            }
          }
          return null;
        }
        let result = null;
        for (let i = reading.length; i >= text.length; --i) {
          const segments = segmentizeFurigana(reading.substring(i), readingNormalized.substring(i), groups, groupsStart + 1);
          if (segments !== null) {
            if (result !== null) { return null; }
            segments.unshift(createFuriganaSegment(text, reading.substring(0, i)));
            result = segments;
          }
          if (groupCount === 1) { break; }
        }
        return result;
      }
      function distributeFurigana(term, reading) {
        if (reading === term) { return [createFuriganaSegment(term, '')]; }
        const groups = [];
        let groupPre = null;
        let isKanaPre = null;
        for (const c of term) {
          const isKana = isCodePointKana(c.codePointAt(0));
          if (isKana === isKanaPre) { groupPre.text += c; }
          else {
            groupPre = {isKana, text: c, textNormalized: null};
            groups.push(groupPre);
            isKanaPre = isKana;
          }
        }
        for (const group of groups) {
          if (group.isKana) { group.textNormalized = convertKatakanaToHiragana(group.text); }
        }
        const segments = segmentizeFurigana(reading, convertKatakanaToHiragana(reading), groups, 0);
        return segments !== null ? segments : [createFuriganaSegment(term, reading)];
      }
      function getStemLength(text1, text2) {
        const minLength = Math.min(text1.length, text2.length);
        if (minLength === 0) { return 0; }
        let i = 0;
        while (true) {
          const char1 = text1.codePointAt(i);
          const char2 = text2.codePointAt(i);
          if (char1 !== char2) { break; }
          const charLength = String.fromCodePoint(char1).length;
          i += charLength;
          if (i >= minLength) {
            if (i > minLength) { i -= charLength; }
            break;
          }
        }
        return i;
      }
      function distributeFuriganaInflected(term, reading, source) {
        const termNormalized = convertKatakanaToHiragana(term);
        const readingNormalized = convertKatakanaToHiragana(reading);
        const sourceNormalized = convertKatakanaToHiragana(source);
        let mainText = term;
        let stemLength = getStemLength(termNormalized, sourceNormalized);
        const readingStemLength = getStemLength(readingNormalized, sourceNormalized);
        if (readingStemLength > 0 && readingStemLength >= stemLength) {
          mainText = reading;
          stemLength = readingStemLength;
          reading = source.substring(0, stemLength) + reading.substring(stemLength);
        }
        const segments = [];
        if (stemLength > 0) {
          mainText = source.substring(0, stemLength) + mainText.substring(stemLength);
          const segments2 = distributeFurigana(mainText, reading);
          let consumed = 0;
          for (const segment of segments2) {
            const start = consumed;
            consumed += segment.text.length;
            if (consumed < stemLength) { segments.push(segment); }
            else if (consumed === stemLength) { segments.push(segment); break; }
            else {
              if (start < stemLength) { segments.push(createFuriganaSegment(mainText.substring(start, stemLength), '')); }
              break;
            }
          }
        }
        if (stemLength < source.length) {
          const remainder = source.substring(stemLength);
          const last = segments[segments.length - 1];
          if (last && last.reading.length === 0) { last.text += remainder; }
          else { segments.push(createFuriganaSegment(remainder, '')); }
        }
        return segments;
      }
      function parsePositiveFrequencyNumber(value) {
        if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
          return Math.max(1, Math.floor(value));
        }
        if (typeof value === 'string') {
          const numericMatch = value.trim().match(/[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/)?.[0];
          if (!numericMatch) { return null; }
          const parsed = Number.parseFloat(numericMatch);
          if (!Number.isFinite(parsed) || parsed <= 0) { return null; }
          return Math.max(1, Math.floor(parsed));
        }
        if (Array.isArray(value)) {
          for (const item of value) {
            const parsed = parsePositiveFrequencyNumber(item);
            if (parsed !== null) { return parsed; }
          }
        }
        return null;
      }
      function parseDisplayFrequencyNumber(value) {
        if (typeof value === 'string') {
          const leadingDigits = value.trim().match(/^\d+/)?.[0];
          if (!leadingDigits) { return null; }
          const parsed = Number.parseInt(leadingDigits, 10);
          return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
        }
        return parsePositiveFrequencyNumber(value);
      }
      function getFrequencyDictionaryName(frequency) {
        const candidates = [
          frequency?.dictionary,
          frequency?.dictionaryName,
          frequency?.name,
          frequency?.title,
          frequency?.dictionaryTitle,
          frequency?.dictionaryAlias
        ];
        for (const candidate of candidates) {
          if (typeof candidate === 'string' && candidate.trim().length > 0) {
            return candidate.trim();
          }
        }
        return null;
      }
      function getBestFrequencyRank(dictionaryEntry, headwordIndex, dictionaryPriorityByName, dictionaryFrequencyModeByName) {
        let best = null;
        const headwordCount = Array.isArray(dictionaryEntry?.headwords) ? dictionaryEntry.headwords.length : 0;
        for (const frequency of dictionaryEntry?.frequencies || []) {
          if (!frequency || typeof frequency !== 'object') { continue; }
          const frequencyHeadwordIndex = frequency.headwordIndex;
          if (typeof frequencyHeadwordIndex === 'number') {
            if (frequencyHeadwordIndex !== headwordIndex) { continue; }
          } else if (headwordCount > 1) {
            continue;
          }
          const dictionary = getFrequencyDictionaryName(frequency);
          if (!dictionary) { continue; }
          if (dictionaryFrequencyModeByName[dictionary] === 'occurrence-based') { continue; }
          const rank =
            parseDisplayFrequencyNumber(frequency.displayValue) ??
            parsePositiveFrequencyNumber(frequency.frequency);
          if (rank === null) { continue; }
          const priorityRaw = dictionaryPriorityByName[dictionary];
          const fallbackPriority =
            typeof frequency.dictionaryIndex === 'number' && Number.isFinite(frequency.dictionaryIndex)
              ? Math.max(0, Math.floor(frequency.dictionaryIndex))
              : Number.MAX_SAFE_INTEGER;
          const priority =
            typeof priorityRaw === 'number' && Number.isFinite(priorityRaw)
              ? Math.max(0, Math.floor(priorityRaw))
              : fallbackPriority;
          if (best === null || priority < best.priority || (priority === best.priority && rank < best.rank)) {
            best = { priority, rank };
          }
        }
        return best?.rank ?? null;
      }
      function hasExactSource(headword, token, requirePrimary) {
        for (const src of headword.sources || []) {
          if (src.originalText !== token) { continue; }
          if (requirePrimary && !src.isPrimary) { continue; }
          if (src.matchType !== 'exact') { continue; }
          return true;
        }
        return false;
      }
      function collectExactHeadwordMatches(dictionaryEntries, token, requirePrimary) {
        const matches = [];
        for (const dictionaryEntry of dictionaryEntries || []) {
          const headwords = Array.isArray(dictionaryEntry?.headwords) ? dictionaryEntry.headwords : [];
          for (let headwordIndex = 0; headwordIndex < headwords.length; headwordIndex += 1) {
            const headword = headwords[headwordIndex];
            if (!hasExactSource(headword, token, requirePrimary)) { continue; }
            matches.push({ dictionaryEntry, headword, headwordIndex });
          }
        }
        return matches;
      }
      function sameHeadword(match, preferredMatch) {
        if (!match || !preferredMatch) {
          return false;
        }
        if (match.headword?.term !== preferredMatch.headword?.term) {
          return false;
        }
        const matchReading = typeof match.headword?.reading === 'string' ? match.headword.reading : '';
        const preferredReading =
          typeof preferredMatch.headword?.reading === 'string' ? preferredMatch.headword.reading : '';
        if (!matchReading || !preferredReading) {
          return true;
        }
        return matchReading === preferredReading;
      }
      function getBestFrequencyRankForMatches(matches, dictionaryPriorityByName, dictionaryFrequencyModeByName) {
        let best = null;
        for (const match of matches) {
          const rank = getBestFrequencyRank(
            match.dictionaryEntry,
            match.headwordIndex,
            dictionaryPriorityByName,
            dictionaryFrequencyModeByName
          );
          if (rank === null) { continue; }
          if (best === null || rank < best) {
            best = rank;
          }
        }
        return best;
      }
      function getPreferredHeadword(dictionaryEntries, token, dictionaryPriorityByName, dictionaryFrequencyModeByName) {
        function normalizeWordClasses(headword) {
          if (!Array.isArray(headword?.wordClasses)) { return undefined; }
          const classes = headword.wordClasses.filter((wordClass) => typeof wordClass === "string" && wordClass.trim().length > 0);
          return classes.length > 0 ? classes : undefined;
        }
        function appendDictionaryNames(target, value) {
          if (!value || typeof value !== 'object') {
            return;
          }
          const candidates = [
            value.dictionary,
            value.dictionaryName,
            value.name,
            value.title,
            value.dictionaryTitle,
            value.dictionaryAlias
          ];
          for (const candidate of candidates) {
            if (typeof candidate === 'string' && candidate.trim().length > 0) {
              target.push(candidate.trim());
            }
          }
        }
        function getDictionaryEntryNames(entry) {
          const names = [];
          appendDictionaryNames(names, entry);
          for (const definition of entry?.definitions || []) {
            appendDictionaryNames(names, definition);
          }
          for (const frequency of entry?.frequencies || []) {
            appendDictionaryNames(names, frequency);
          }
          for (const pronunciation of entry?.pronunciations || []) {
            appendDictionaryNames(names, pronunciation);
          }
          return names;
        }
        function isNameDictionaryEntry(entry) {
          if (!includeNameMatchMetadata || !entry || typeof entry !== 'object') {
            return false;
          }
          return getDictionaryEntryNames(entry).some((name) => name.startsWith("SubMiner Character Dictionary"));
        }
        const exactPrimaryMatches = collectExactHeadwordMatches(dictionaryEntries, token, true);
        let matchedNameDictionary = false;
        if (includeNameMatchMetadata) {
          for (const dictionaryEntry of dictionaryEntries || []) {
            if (!isNameDictionaryEntry(dictionaryEntry)) { continue; }
            for (const match of exactPrimaryMatches) {
              if (match.dictionaryEntry !== dictionaryEntry) { continue; }
              matchedNameDictionary = true;
              break;
            }
            if (matchedNameDictionary) { break; }
          }
        }
        const preferredMatch = exactPrimaryMatches[0];
        if (preferredMatch) {
          const exactFrequencyMatches = collectExactHeadwordMatches(dictionaryEntries, token, false)
            .filter((match) => sameHeadword(match, preferredMatch));
          return {
            term: preferredMatch.headword.term,
            reading: preferredMatch.headword.reading,
            wordClasses: normalizeWordClasses(preferredMatch.headword),
            isNameMatch: matchedNameDictionary || isNameDictionaryEntry(preferredMatch.dictionaryEntry),
            frequencyRank: getBestFrequencyRankForMatches(
              exactFrequencyMatches.length > 0 ? exactFrequencyMatches : exactPrimaryMatches,
              dictionaryPriorityByName,
              dictionaryFrequencyModeByName
            )
          };
        }
        return null;
      }
`;

function buildYomitanScanningScript(
  text: string,
  profileIndex: number,
  scanLength: number,
  includeNameMatchMetadata: boolean,
  dictionaryPriorityByName: Record<string, number>,
  dictionaryFrequencyModeByName: Partial<Record<string, YomitanFrequencyMode>>,
): string {
  return `
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
${YOMITAN_SCANNING_HELPERS}
      const includeNameMatchMetadata = ${includeNameMatchMetadata ? 'true' : 'false'};
      const dictionaryPriorityByName = ${JSON.stringify(dictionaryPriorityByName)};
      const dictionaryFrequencyModeByName = ${JSON.stringify(dictionaryFrequencyModeByName)};
      const text = ${JSON.stringify(text)};
      const details = {matchType: "exact", deinflect: true};
      const tokens = [];
      let i = 0;
      while (i < text.length) {
        const codePoint = text.codePointAt(i);
        const character = String.fromCodePoint(codePoint);
        const substring = text.substring(i, i + ${scanLength});
        const result = await invoke("termsFind", { text: substring, details, optionsContext: { index: ${profileIndex} } });
        const dictionaryEntries = Array.isArray(result?.dictionaryEntries) ? result.dictionaryEntries : [];
        const originalTextLength = typeof result?.originalTextLength === "number" ? result.originalTextLength : 0;
        if (dictionaryEntries.length > 0 && originalTextLength > 0 && (originalTextLength !== character.length || isCodePointJapanese(codePoint))) {
          const source = substring.substring(0, originalTextLength);
          const preferredHeadword = getPreferredHeadword(
            dictionaryEntries,
            source,
            dictionaryPriorityByName,
            dictionaryFrequencyModeByName
          );
          if (preferredHeadword && typeof preferredHeadword.term === "string") {
            const reading = typeof preferredHeadword.reading === "string" ? preferredHeadword.reading : "";
            const segments = distributeFuriganaInflected(preferredHeadword.term, reading, source);
            const tokenPayload = {
              surface: segments.map((segment) => segment.text).join("") || source,
              reading: segments.map((segment) => typeof segment.reading === "string" ? segment.reading : "").join(""),
              headword: preferredHeadword.term,
              startPos: i,
              endPos: i + originalTextLength,
              isNameMatch: includeNameMatchMetadata && preferredHeadword.isNameMatch === true,
              frequencyRank:
                typeof preferredHeadword.frequencyRank === "number" && Number.isFinite(preferredHeadword.frequencyRank)
                  ? Math.max(1, Math.floor(preferredHeadword.frequencyRank))
                  : undefined,
            };
            if (Array.isArray(preferredHeadword.wordClasses) && preferredHeadword.wordClasses.length > 0) {
              tokenPayload.wordClasses = preferredHeadword.wordClasses;
            }
            tokens.push(tokenPayload);
            i += originalTextLength;
            continue;
          }
        }
        i += character.length;
      }
      return tokens;
    })();
  `;
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

  const metadata = await requestYomitanProfileMetadata(parserWindow, logger);
  const script =
    metadata !== null
      ? `
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

      return await invoke("parseText", {
        text: ${JSON.stringify(text)},
        optionsContext: { index: ${metadata.profileIndex} },
        scanLength: ${metadata.scanLength},
        useInternalParser: true,
        useMecabParser: false
      });
    })();
  `
      : `
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
        optionsFull.profiles?.[profileIndex]?.options?.scanning?.length ?? ${DEFAULT_YOMITAN_SCAN_LENGTH};

      return await invoke("parseText", {
        text: ${JSON.stringify(text)},
        optionsContext: { index: profileIndex },
        scanLength,
        useInternalParser: true,
        useMecabParser: false
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

export async function requestYomitanScanTokens(
  text: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
  options?: {
    includeNameMatchMetadata?: boolean;
  },
): Promise<YomitanScanToken[] | null> {
  const yomitanExt = deps.getYomitanExt();
  if (!text || !yomitanExt) {
    return null;
  }

  const isReady = await ensureYomitanParserWindow(deps, logger);
  const parserWindow = deps.getYomitanParserWindow();
  if (!isReady || !parserWindow || parserWindow.isDestroyed()) {
    return null;
  }

  const parseResults = await requestYomitanParseResults(text, deps, logger);
  const selectedParseTokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  const parseScanTokens =
    selectedParseTokens?.map((token) => ({
      surface: token.surface,
      reading: token.reading,
      headword: token.headword,
      startPos: token.startPos,
      endPos: token.endPos,
    })) ?? null;

  const metadata = await requestYomitanProfileMetadata(parserWindow, logger);
  const profileIndex = metadata?.profileIndex ?? 0;
  const scanLength = metadata?.scanLength ?? DEFAULT_YOMITAN_SCAN_LENGTH;

  try {
    const rawResult = await parserWindow.webContents.executeJavaScript(
      buildYomitanScanningScript(
        text,
        profileIndex,
        scanLength,
        options?.includeNameMatchMetadata === true,
        metadata?.dictionaryPriorityByName ?? {},
        metadata?.dictionaryFrequencyModeByName ?? {},
      ),
      true,
    );
    if (isScanTokenArray(rawResult)) {
      if (parseScanTokens && parseScanTokens.length > 0) {
        return hasSameTokenSpans(parseScanTokens, rawResult) ? rawResult : parseScanTokens;
      }
      return rawResult;
    }
    if (Array.isArray(rawResult)) {
      const selectedTokens = selectYomitanParseTokens(rawResult, () => false, 'headword');
      return (
        selectedTokens?.map((token) => ({
          surface: token.surface,
          reading: token.reading,
          headword: token.headword,
          startPos: token.startPos,
          endPos: token.endPos,
        })) ?? null
      );
    }
    if (parseScanTokens && parseScanTokens.length > 0) {
      return parseScanTokens;
    }
    return null;
  } catch (err) {
    if (parseScanTokens && parseScanTokens.length > 0) {
      return parseScanTokens;
    }
    logger.error('Yomitan scanner request failed:', (err as Error).message);
    return null;
  }
}

async function fetchYomitanTermFrequencies(
  parserWindow: BrowserWindow,
  termReadingList: YomitanTermReadingPair[],
  metadata: YomitanProfileMetadata | null,
  logger: LoggerLike,
): Promise<YomitanTermFrequency[] | null> {
  if (metadata && metadata.dictionaries.length > 0) {
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

        return await invoke("getTermFrequencies", {
          termReadingList: ${JSON.stringify(termReadingList)},
          dictionaries: ${JSON.stringify(metadata.dictionaries)}
        });
      })();
    `;

    try {
      const rawResult = await parserWindow.webContents.executeJavaScript(script, true);
      return Array.isArray(rawResult)
        ? normalizeFrequencyEntriesWithPriority(
            rawResult,
            metadata.dictionaryPriorityByName,
            metadata.dictionaryFrequencyModeByName,
          )
        : [];
    } catch (err) {
      logger.error('Yomitan term frequency request failed:', (err as Error).message);
      return null;
    }
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
      const dictionaryEntries = Array.isArray(dictionariesRaw)
        ? dictionariesRaw
            .filter((entry) => entry && typeof entry === "object" && entry.enabled === true && typeof entry.name === "string")
            .map((entry, index) => ({
              name: entry.name,
              id: typeof entry.id === "number" && Number.isFinite(entry.id) ? Math.floor(entry.id) : index
            }))
            .sort((a, b) => a.id - b.id)
        : [];
      const dictionaries = dictionaryEntries.map((entry) => entry.name);
      const dictionaryPriorityByName = dictionaryEntries.reduce((acc, entry, index) => {
        acc[entry.name] = index;
        return acc;
      }, {});

      if (dictionaries.length === 0) {
        return [];
      }

      const rawFrequencies = await invoke("getTermFrequencies", {
        termReadingList: ${JSON.stringify(termReadingList)},
        dictionaries
      });

      if (!Array.isArray(rawFrequencies)) {
        return [];
      }

      return rawFrequencies
        .filter((entry) => entry && typeof entry === "object")
        .map((entry) => ({
          ...entry,
          dictionaryPriority:
            typeof entry.dictionary === "string" && dictionaryPriorityByName[entry.dictionary] !== undefined
              ? dictionaryPriorityByName[entry.dictionary]
              : Number.MAX_SAFE_INTEGER
        }));
    })();
  `;

  try {
    const rawResult = await parserWindow.webContents.executeJavaScript(script, true);
    return Array.isArray(rawResult)
      ? rawResult
          .map((entry) => toYomitanTermFrequency(entry))
          .filter((entry): entry is YomitanTermFrequency => entry !== null)
      : [];
  } catch (err) {
    logger.error('Yomitan term frequency request failed:', (err as Error).message);
    return null;
  }
}

function cacheFrequencyEntriesForPairs(
  frequencyCache: Map<string, YomitanTermFrequency[]>,
  termReadingList: YomitanTermReadingPair[],
  fetchedEntries: YomitanTermFrequency[],
): void {
  const groupedByPair = groupFrequencyEntriesByPair(fetchedEntries);
  const groupedByTerm = groupFrequencyEntriesByTerm(fetchedEntries);
  for (const pair of termReadingList) {
    const key = makeTermReadingCacheKey(pair.term, pair.reading);
    const exactEntries = groupedByPair.get(key);
    const termEntries = groupedByTerm.get(pair.term) ?? [];
    frequencyCache.set(key, exactEntries ?? termEntries);
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

  const metadata = await requestYomitanProfileMetadata(parserWindow, logger);
  const frequencyCache = getWindowFrequencyCache(parserWindow);
  const missingTermReadingList: YomitanTermReadingPair[] = [];

  const buildCachedResult = (): YomitanTermFrequency[] => {
    const result: YomitanTermFrequency[] = [];
    for (const pair of normalizedTermReadingList) {
      const key = makeTermReadingCacheKey(pair.term, pair.reading);
      const cached = frequencyCache.get(key);
      if (cached && cached.length > 0) {
        result.push(...cached);
      }
    }
    return result;
  };

  for (const pair of normalizedTermReadingList) {
    const key = makeTermReadingCacheKey(pair.term, pair.reading);
    if (!frequencyCache.has(key)) {
      missingTermReadingList.push(pair);
    }
  }

  if (missingTermReadingList.length === 0) {
    return buildCachedResult();
  }

  const fetchedEntries = await fetchYomitanTermFrequencies(
    parserWindow,
    missingTermReadingList,
    metadata,
    logger,
  );
  if (fetchedEntries === null) {
    return buildCachedResult();
  }

  cacheFrequencyEntriesForPairs(frequencyCache, missingTermReadingList, fetchedEntries);

  const fallbackTermReadingList = normalizeTermReadingList(
    missingTermReadingList
      .filter((pair) => pair.reading !== null)
      .map((pair) => {
        const key = makeTermReadingCacheKey(pair.term, pair.reading);
        const cachedEntries = frequencyCache.get(key);
        if (cachedEntries && cachedEntries.length > 0) {
          return null;
        }

        const fallbackKey = makeTermReadingCacheKey(pair.term, null);
        const cachedFallback = frequencyCache.get(fallbackKey);
        if (cachedFallback && cachedFallback.length > 0) {
          frequencyCache.set(key, cachedFallback);
          return null;
        }

        return { term: pair.term, reading: null };
      })
      .filter((pair): pair is { term: string; reading: null } => pair !== null),
  ).filter((pair) => !frequencyCache.has(makeTermReadingCacheKey(pair.term, pair.reading)));

  let fallbackFetchedEntries: YomitanTermFrequency[] = [];

  if (fallbackTermReadingList.length > 0) {
    const fallbackFetchResult = await fetchYomitanTermFrequencies(
      parserWindow,
      fallbackTermReadingList,
      metadata,
      logger,
    );
    if (fallbackFetchResult !== null) {
      fallbackFetchedEntries = fallbackFetchResult;
      cacheFrequencyEntriesForPairs(
        frequencyCache,
        fallbackTermReadingList,
        fallbackFetchedEntries,
      );
    }

    for (const pair of missingTermReadingList) {
      if (pair.reading === null) {
        continue;
      }
      const key = makeTermReadingCacheKey(pair.term, pair.reading);
      const cachedEntries = frequencyCache.get(key);
      if (cachedEntries && cachedEntries.length > 0) {
        continue;
      }
      const fallbackEntries = frequencyCache.get(makeTermReadingCacheKey(pair.term, null));
      if (fallbackEntries && fallbackEntries.length > 0) {
        frequencyCache.set(key, fallbackEntries);
      }
    }
  }

  const allFetchedEntries = [...fetchedEntries, ...fallbackFetchedEntries];
  const queriedTerms = new Set(
    [...missingTermReadingList, ...fallbackTermReadingList].map((pair) => pair.term),
  );
  const cachedResult = buildCachedResult();
  const unmatchedEntries = allFetchedEntries.filter(
    (entry) => !queriedTerms.has(entry.term.trim()),
  );
  return [...cachedResult, ...unmatchedEntries];
}

export async function syncYomitanDefaultAnkiServer(
  serverUrl: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
  options?: {
    forceOverride?: boolean;
  },
): Promise<boolean> {
  const normalizedTargetServer = serverUrl.trim();
  if (!normalizedTargetServer) {
    return false;
  }
  const forceOverride = options?.forceOverride === true;

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
      const forceOverride = ${forceOverride ? 'true' : 'false'};
      const optionsFull = await invoke("optionsGetFull", undefined);
      const profiles = Array.isArray(optionsFull.profiles) ? optionsFull.profiles : [];
      if (profiles.length === 0) {
        return { updated: false, reason: "no-profiles" };
      }

      const profileCurrent = Number.isInteger(optionsFull.profileCurrent)
        ? optionsFull.profileCurrent
        : 0;
      const targetProfile = profiles[profileCurrent];
      if (!targetProfile || typeof targetProfile !== "object") {
        return { updated: false, reason: "invalid-default-profile" };
      }

      targetProfile.options = targetProfile.options && typeof targetProfile.options === "object"
        ? targetProfile.options
        : {};
      targetProfile.options.anki = targetProfile.options.anki && typeof targetProfile.options.anki === "object"
        ? targetProfile.options.anki
        : {};

      const currentServerRaw = targetProfile.options.anki.server;
      const currentServer = typeof currentServerRaw === "string" ? currentServerRaw.trim() : "";
      if (currentServer === targetServer) {
        return { updated: false, matched: true, reason: "already-target", currentServer, targetServer };
      }
      const canReplaceCurrent =
        forceOverride || currentServer.length === 0 || currentServer === "http://127.0.0.1:8765";
      if (!canReplaceCurrent) {
        return { updated: false, matched: false, reason: "blocked-existing-server", currentServer, targetServer };
      }

      targetProfile.options.anki.server = targetServer;
      await invoke("setAllSettings", { value: optionsFull, source: "subminer" });
      return { updated: true, matched: true, currentServer, targetServer };
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
    const matchedWithoutUpdate =
      isObject(result) &&
      result.updated === false &&
      (result as { matched?: unknown }).matched === true;
    if (matchedWithoutUpdate) {
      return true;
    }
    const blockedByExistingServer =
      isObject(result) &&
      result.updated === false &&
      (result as { matched?: unknown }).matched === false &&
      typeof (result as { reason?: unknown }).reason === 'string';
    if (blockedByExistingServer) {
      logger.info?.(
        `Skipped syncing Yomitan Anki server (reason=${String((result as { reason: string }).reason)})`,
      );
      return false;
    }
    const checkedWithoutUpdate =
      typeof result === 'object' &&
      result !== null &&
      (result as { updated?: unknown }).updated === false;
    return checkedWithoutUpdate;
  } catch (err) {
    logger.error('Failed to sync Yomitan default profile Anki server:', (err as Error).message);
    return false;
  }
}

function buildYomitanInvokeScript(actionLiteral: string, paramsLiteral: string): string {
  return `
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

      return await invoke(${actionLiteral}, ${paramsLiteral});
    })();
  `;
}

async function invokeYomitanBackendAction<T>(
  action: string,
  params: unknown,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<T | null> {
  const isReady = await ensureYomitanParserWindow(deps, logger);
  const parserWindow = deps.getYomitanParserWindow();
  if (!isReady || !parserWindow || parserWindow.isDestroyed()) {
    return null;
  }

  const script = buildYomitanInvokeScript(
    JSON.stringify(action),
    params === undefined ? 'undefined' : JSON.stringify(params),
  );

  try {
    return (await parserWindow.webContents.executeJavaScript(script, true)) as T;
  } catch (err) {
    logger.error(`Yomitan backend action failed (${action}):`, (err as Error).message);
    return null;
  }
}

function createDefaultDictionarySettings(name: string, enabled: boolean): Record<string, unknown> {
  return {
    name,
    alias: name,
    enabled,
    allowSecondarySearches: false,
    definitionsCollapsible: 'not-collapsible',
    partsOfSpeechFilter: true,
    useDeinflections: true,
    styles: '',
  };
}

function getTargetProfileIndices(
  optionsFull: Record<string, unknown>,
  profileScope: 'all' | 'active',
): number[] {
  const profiles = Array.isArray(optionsFull.profiles) ? optionsFull.profiles : [];
  if (profileScope === 'active') {
    const profileCurrent =
      typeof optionsFull.profileCurrent === 'number' && Number.isFinite(optionsFull.profileCurrent)
        ? Math.max(0, Math.floor(optionsFull.profileCurrent))
        : 0;
    return profileCurrent < profiles.length ? [profileCurrent] : [];
  }
  return profiles.map((_profile, index) => index);
}

export async function getYomitanDictionaryInfo(
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<YomitanDictionaryInfo[]> {
  const result = await invokeYomitanBackendAction<unknown>(
    'getDictionaryInfo',
    undefined,
    deps,
    logger,
  );
  if (!Array.isArray(result)) {
    return [];
  }

  return result
    .filter((entry): entry is Record<string, unknown> => isObject(entry))
    .map((entry) => {
      const title = typeof entry.title === 'string' ? entry.title.trim() : '';
      const revision = entry.revision;
      const frequencyMode: YomitanFrequencyMode | undefined =
        entry.frequencyMode === 'occurrence-based' || entry.frequencyMode === 'rank-based'
          ? entry.frequencyMode
          : undefined;
      return {
        title,
        revision:
          typeof revision === 'string' || typeof revision === 'number' ? revision : undefined,
        frequencyMode,
      };
    })
    .filter((entry) => entry.title.length > 0);
}

export async function getYomitanSettingsFull(
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<Record<string, unknown> | null> {
  const result = await invokeYomitanBackendAction<unknown>(
    'optionsGetFull',
    undefined,
    deps,
    logger,
  );
  return isObject(result) ? result : null;
}

export async function setYomitanSettingsFull(
  value: Record<string, unknown>,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
  source = 'subminer',
): Promise<boolean> {
  const result = await invokeYomitanBackendAction<unknown>(
    'setAllSettings',
    { value, source },
    deps,
    logger,
  );
  return result !== null;
}

export async function importYomitanDictionaryFromZip(
  zipPath: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<boolean> {
  const normalizedZipPath = zipPath.trim();
  if (!normalizedZipPath || !fs.existsSync(normalizedZipPath)) {
    logger.error(`Dictionary ZIP not found: ${zipPath}`);
    return false;
  }

  const supportsUrlImport = await invokeYomitanSettingsAutomation<boolean>(
    `
      (() => typeof globalThis.__subminerYomitanSettingsAutomation.importDictionaryArchiveUrl === "function")();
    `,
    deps,
    logger,
  );

  const result =
    supportsUrlImport === true
      ? await serveDictionaryZipOnce(normalizedZipPath, async (archiveUrl) =>
          invokeYomitanSettingsAutomation<boolean>(
            `
              (async () => {
                await globalThis.__subminerYomitanSettingsAutomation.importDictionaryArchiveUrl(
                  ${JSON.stringify(archiveUrl)}
                );
                return true;
              })();
            `,
            deps,
            logger,
          ),
        )
      : await invokeYomitanSettingsAutomation<boolean>(
          `
            (async () => {
              await globalThis.__subminerYomitanSettingsAutomation.importDictionaryArchiveBase64(
                ${JSON.stringify(fs.readFileSync(normalizedZipPath).toString('base64'))},
                ${JSON.stringify(path.basename(normalizedZipPath))}
              );
              return true;
            })();
          `,
          deps,
          logger,
        );
  return result === true;
}

export async function deleteYomitanDictionaryByTitle(
  dictionaryTitle: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<boolean> {
  const normalizedTitle = dictionaryTitle.trim();
  if (!normalizedTitle) {
    return false;
  }
  const result = await invokeYomitanSettingsAutomation<boolean>(
    `
      (async () => {
        await globalThis.__subminerYomitanSettingsAutomation.deleteDictionary(
          ${JSON.stringify(normalizedTitle)}
        );
        return true;
      })();
    `,
    deps,
    logger,
  );
  return result === true;
}

export async function upsertYomitanDictionarySettings(
  dictionaryTitle: string,
  profileScope: 'all' | 'active',
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<boolean> {
  const normalizedTitle = dictionaryTitle.trim();
  if (!normalizedTitle) {
    return false;
  }

  const optionsFull = await getYomitanSettingsFull(deps, logger);
  if (!optionsFull) {
    return false;
  }

  const profiles = Array.isArray(optionsFull.profiles) ? optionsFull.profiles : [];
  const indices = getTargetProfileIndices(optionsFull, profileScope);
  let changed = false;

  for (const index of indices) {
    const profile = profiles[index];
    if (!isObject(profile)) {
      continue;
    }

    if (!isObject(profile.options)) {
      profile.options = {};
    }
    const profileOptions = profile.options as Record<string, unknown>;
    if (!Array.isArray(profileOptions.dictionaries)) {
      profileOptions.dictionaries = [];
    }

    const dictionaries = profileOptions.dictionaries as unknown[];
    const existingIndex = dictionaries.findIndex(
      (entry) =>
        isObject(entry) &&
        typeof (entry as { name?: unknown }).name === 'string' &&
        (entry as { name: string }).name.trim() === normalizedTitle,
    );

    if (existingIndex >= 0) {
      const existing = dictionaries[existingIndex] as Record<string, unknown>;
      if (existing.enabled !== true) {
        existing.enabled = true;
        changed = true;
      }
      if (typeof existing.alias !== 'string' || existing.alias.trim().length === 0) {
        existing.alias = normalizedTitle;
        changed = true;
      }
      continue;
    }

    dictionaries.push(createDefaultDictionarySettings(normalizedTitle, true));
    changed = true;
  }

  if (!changed) {
    return false;
  }

  return await setYomitanSettingsFull(optionsFull, deps, logger);
}

export async function removeYomitanDictionarySettings(
  dictionaryTitle: string,
  profileScope: 'all' | 'active',
  mode: 'delete' | 'disable',
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<boolean> {
  const normalizedTitle = dictionaryTitle.trim();
  if (!normalizedTitle) {
    return false;
  }

  const optionsFull = await getYomitanSettingsFull(deps, logger);
  if (!optionsFull) {
    return false;
  }

  const profiles = Array.isArray(optionsFull.profiles) ? optionsFull.profiles : [];
  const indices = getTargetProfileIndices(optionsFull, profileScope);
  let changed = false;

  for (const index of indices) {
    const profile = profiles[index];
    if (!isObject(profile) || !isObject(profile.options)) {
      continue;
    }
    const profileOptions = profile.options as Record<string, unknown>;
    if (!Array.isArray(profileOptions.dictionaries)) {
      continue;
    }

    const dictionaries = profileOptions.dictionaries as unknown[];
    if (mode === 'delete') {
      const before = dictionaries.length;
      profileOptions.dictionaries = dictionaries.filter(
        (entry) =>
          !(
            isObject(entry) &&
            typeof (entry as { name?: unknown }).name === 'string' &&
            (entry as { name: string }).name.trim() === normalizedTitle
          ),
      );
      if ((profileOptions.dictionaries as unknown[]).length !== before) {
        changed = true;
      }
      continue;
    }

    for (const entry of dictionaries) {
      if (
        !isObject(entry) ||
        typeof (entry as { name?: unknown }).name !== 'string' ||
        (entry as { name: string }).name.trim() !== normalizedTitle
      ) {
        continue;
      }
      const dictionaryEntry = entry as Record<string, unknown>;
      if (dictionaryEntry.enabled !== false) {
        dictionaryEntry.enabled = false;
        changed = true;
      }
    }
  }

  if (!changed) {
    return false;
  }

  return await setYomitanSettingsFull(optionsFull, deps, logger);
}

export async function addYomitanNoteViaSearch(
  word: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<YomitanAddNoteResult> {
  const isReady = await ensureYomitanParserWindow(deps, logger);
  const parserWindow = deps.getYomitanParserWindow();
  if (!isReady || !parserWindow || parserWindow.isDestroyed()) {
    return { noteId: null, duplicateNoteIds: [] };
  }

  const escapedWord = JSON.stringify(word);

  const script = `
    (async () => {
      if (typeof window.__subminerAddNote !== 'function') {
        throw new Error('Yomitan search page bridge not initialized');
      }
      return await window.__subminerAddNote(${escapedWord});
    })();
  `;

  try {
    const result = await parserWindow.webContents.executeJavaScript(script, true);
    if (typeof result === 'number') {
      return {
        noteId: Number.isInteger(result) && result > 0 ? result : null,
        duplicateNoteIds: [],
      };
    }
    if (result && typeof result === 'object' && !Array.isArray(result)) {
      const envelope = result as {
        noteId?: unknown;
        duplicateNoteIds?: unknown;
      };
      return {
        noteId:
          typeof envelope.noteId === 'number' &&
          Number.isInteger(envelope.noteId) &&
          envelope.noteId > 0
            ? envelope.noteId
            : null,
        duplicateNoteIds: Array.isArray(envelope.duplicateNoteIds)
          ? envelope.duplicateNoteIds.filter(
              (entry): entry is number =>
                typeof entry === 'number' && Number.isInteger(entry) && entry > 0,
            )
          : [],
      };
    }
    return { noteId: null, duplicateNoteIds: [] };
  } catch (err) {
    logger.error('Yomitan addNoteFromWord failed:', (err as Error).message);
    return { noteId: null, duplicateNoteIds: [] };
  }
}
