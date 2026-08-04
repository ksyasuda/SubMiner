import type { BrowserWindow, Extension, Session } from 'electron';
import * as fs from 'fs';
import * as http from 'http';
import * as path from 'path';
import { selectYomitanParseTokens } from './parser-selection-stage';
import {
  buildYomitanScanCallScript,
  buildYomitanScanNameCandidatesScript,
  CHARACTER_DICTIONARY_TITLE_PREFIX,
  YOMITAN_SCAN_RUNTIME_INSTALL_SCRIPT,
  YOMITAN_SCAN_RUNTIME_MISSING_SENTINEL,
  type YomitanFrequencyMode,
} from './yomitan-scan-runtime-script';

interface LoggerLike {
  error: (message: string, ...args: unknown[]) => void;
  info?: (message: string, ...args: unknown[]) => void;
  warn?: (message: string, ...args: unknown[]) => void;
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
  headwordReading?: string;
  startPos: number;
  endPos: number;
  isNameMatch?: boolean;
  frequencyRank?: number;
  wordClasses?: string[];
  isUnparsedRun?: boolean;
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
const yomitanProfileDiagnosticsLoggedByWindow = new WeakSet<BrowserWindow>();
const yomitanFrequencyCacheByWindow = new WeakMap<
  BrowserWindow,
  Map<string, YomitanTermFrequency[]>
>();
// Epoch passed with every scan request; the in-window termsFind cache clears
// itself when the epoch changes (dictionary imports, settings changes).
const yomitanScanCacheEpochByWindow = new WeakMap<BrowserWindow, number>();

function getYomitanScanCacheEpoch(window: BrowserWindow): number {
  return yomitanScanCacheEpochByWindow.get(window) ?? 0;
}

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
        (entry.headwordReading === undefined || typeof entry.headwordReading === 'string') &&
        typeof entry.startPos === 'number' &&
        typeof entry.endPos === 'number' &&
        (entry.isNameMatch === undefined || typeof entry.isNameMatch === 'boolean') &&
        (entry.isUnparsedRun === undefined || typeof entry.isUnparsedRun === 'boolean') &&
        (entry.frequencyRank === undefined || typeof entry.frequencyRank === 'number') &&
        (entry.wordClasses === undefined ||
          (Array.isArray(entry.wordClasses) &&
            entry.wordClasses.every((wordClass) => typeof wordClass === 'string'))),
    )
  );
}

// Maps a parse-selected token to the scanner-token shape carried out of the
// parser runtime, used by the parseText fallback path when the in-window
// scanner is unavailable.
function toYomitanScanToken(token: {
  surface: string;
  reading: string;
  headword: string;
  startPos: number;
  endPos: number;
  isUnparsedRun?: boolean;
}): YomitanScanToken {
  return {
    surface: token.surface,
    reading: token.reading,
    headword: token.headword,
    startPos: token.startPos,
    endPos: token.endPos,
    ...(token.isUnparsedRun === true ? { isUnparsedRun: true } : {}),
  };
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
  yomitanScanCacheEpochByWindow.set(window, getYomitanScanCacheEpoch(window) + 1);
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
    logYomitanProfileDiagnostics(parserWindow, metadata, logger);
    return metadata;
  } catch (err) {
    logger.error('Yomitan parser metadata request failed:', (err as Error).message);
    return null;
  }
}

function logYomitanProfileDiagnostics(
  parserWindow: BrowserWindow,
  metadata: YomitanProfileMetadata,
  logger: LoggerLike,
): void {
  if (yomitanProfileDiagnosticsLoggedByWindow.has(parserWindow)) {
    return;
  }
  yomitanProfileDiagnosticsLoggedByWindow.add(parserWindow);

  const visibleDictionaries = metadata.dictionaries.slice(0, 8);
  const details = {
    profileIndex: metadata.profileIndex,
    scanLength: metadata.scanLength,
    dictionaryCount: metadata.dictionaries.length,
    dictionaries: visibleDictionaries,
    omittedDictionaryCount: Math.max(0, metadata.dictionaries.length - visibleDictionaries.length),
  };

  if (metadata.dictionaries.length === 0) {
    const logWarning = logger.warn ?? logger.info;
    logWarning?.(
      'Yomitan active profile has no enabled dictionaries; lookup popups may not show definitions.',
      details,
    );
    return;
  }

  logger.info?.('Yomitan active profile dictionaries loaded.', details);
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
      // Eagerly install the scan runtime so the first subtitle line does not
      // pay the install round trip; failures fall back to the per-request
      // install-and-retry path.
      await installYomitanScanRuntime(parserWindow).catch(() => {});

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

async function installYomitanScanRuntime(parserWindow: BrowserWindow): Promise<void> {
  await parserWindow.webContents.executeJavaScript(YOMITAN_SCAN_RUNTIME_INSTALL_SCRIPT, true);
  // A fresh runtime has no candidate list; force the next scan to reinstall it.
  yomitanScanNameCandidateKeyByWindow.delete(parserWindow);
}

// Key of the character-name candidate list currently installed in each parser
// window, so an unchanged list costs nothing per line.
const yomitanScanNameCandidateKeyByWindow = new WeakMap<BrowserWindow, string>();

async function ensureYomitanScanNameCandidates(
  parserWindow: BrowserWindow,
  nameCandidates: { key: string; forms: string[] } | null,
  logger: LoggerLike,
): Promise<void> {
  const installedKey = yomitanScanNameCandidateKeyByWindow.get(parserWindow);
  const nextKey = nameCandidates?.key ?? '';
  if (installedKey === nextKey) {
    return;
  }

  try {
    await parserWindow.webContents.executeJavaScript(
      buildYomitanScanNameCandidatesScript(nameCandidates),
      true,
    );
    yomitanScanNameCandidateKeyByWindow.set(parserWindow, nextKey);
  } catch (err) {
    // The scan falls back to checking every position when the list is absent,
    // so a failed install costs speed, never a missed name.
    logger.warn?.(
      'Failed to install Yomitan character-name scan candidates:',
      (err as Error).message,
    );
    yomitanScanNameCandidateKeyByWindow.delete(parserWindow);
  }
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

// parseText fallback for when the in-window scanner cannot run (script eval
// failure, unexpected payload). The scanner walk is the primary tokenizer and
// emits its own filler runs, so this extra full parse only happens on errors.
async function requestYomitanParseFallbackTokens(
  text: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<YomitanScanToken[] | null> {
  const parseResults = await requestYomitanParseResults(text, deps, logger);
  const selectedTokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  const parseScanTokens = selectedTokens?.map(toYomitanScanToken) ?? null;
  return parseScanTokens && parseScanTokens.length > 0 ? parseScanTokens : null;
}

export async function requestYomitanScanTokens(
  text: string,
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
  options?: {
    includeNameMatchMetadata?: boolean;
    currentCharacterDictionaryMediaId?: number | null;
    nameCandidates?: { key: string; forms: string[] } | null;
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

  const metadata = await requestYomitanProfileMetadata(parserWindow, logger);
  const profileIndex = metadata?.profileIndex ?? 0;
  const scanLength = metadata?.scanLength ?? DEFAULT_YOMITAN_SCAN_LENGTH;
  const includeNameMatchMetadata = options?.includeNameMatchMetadata === true;
  const greedyNameScanEnabled =
    includeNameMatchMetadata &&
    (metadata?.dictionaries ?? []).some((name) =>
      name.startsWith(CHARACTER_DICTIONARY_TITLE_PREFIX),
    );

  // Candidate name forms let the in-page pre-pass skip positions where no
  // character name can start. Installed only when it changes (per media), so
  // the per-line call stays a single tiny script.
  const nameCandidates = greedyNameScanEnabled ? (options?.nameCandidates ?? null) : null;
  await ensureYomitanScanNameCandidates(parserWindow, nameCandidates, logger);

  const callScript = buildYomitanScanCallScript({
    text,
    profileIndex,
    scanLength,
    includeNameMatchMetadata,
    greedyNameScanEnabled,
    currentCharacterDictionaryMediaId:
      typeof options?.currentCharacterDictionaryMediaId === 'number' &&
      Number.isFinite(options.currentCharacterDictionaryMediaId) &&
      options.currentCharacterDictionaryMediaId > 0
        ? Math.floor(options.currentCharacterDictionaryMediaId)
        : null,
    dictionaryPriorityByName: metadata?.dictionaryPriorityByName ?? {},
    dictionaryFrequencyModeByName: metadata?.dictionaryFrequencyModeByName ?? {},
    cacheEpoch: getYomitanScanCacheEpoch(parserWindow),
    nameCandidateKey: nameCandidates?.key ?? null,
  });

  try {
    let rawResult = await parserWindow.webContents.executeJavaScript(callScript, true);
    if (rawResult === YOMITAN_SCAN_RUNTIME_MISSING_SENTINEL) {
      // First request for this window, or the page reloaded and dropped the
      // installed runtime: install and retry once. The candidate list lives in
      // the same page state, so it has to be reinstalled alongside it.
      await installYomitanScanRuntime(parserWindow);
      await ensureYomitanScanNameCandidates(parserWindow, nameCandidates, logger);
      rawResult = await parserWindow.webContents.executeJavaScript(callScript, true);
    }
    if (isScanTokenArray(rawResult)) {
      // Filler-only results carry no dictionary match; keep the historical
      // contract of returning null so callers fall back to raw text.
      return rawResult.some((token) => token.isUnparsedRun !== true) ? rawResult : null;
    }
    logger.error('Yomitan scanner returned an unexpected payload; using parseText fallback.');
    return await requestYomitanParseFallbackTokens(text, deps, logger);
  } catch (err) {
    logger.error('Yomitan scanner request failed:', (err as Error).message);
    return await requestYomitanParseFallbackTokens(text, deps, logger);
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
    deck?: string;
  },
): Promise<boolean> {
  const normalizedTargetServer = serverUrl.trim();
  if (!normalizedTargetServer) {
    return false;
  }
  const forceOverride = options?.forceOverride === true;
  const normalizedTargetDeck = options?.deck?.trim() ?? '';

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
      const targetDeck = ${JSON.stringify(normalizedTargetDeck)};
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
      let changed = false;
      if (currentServer !== targetServer) {
        const canReplaceCurrent =
          forceOverride || currentServer.length === 0 || currentServer === "http://127.0.0.1:8765";
        if (!canReplaceCurrent) {
          return { updated: false, matched: false, reason: "blocked-existing-server", currentServer, targetServer };
        }

        targetProfile.options.anki.server = targetServer;
        changed = true;
      }

      if (targetDeck) {
        const cardFormats = Array.isArray(targetProfile.options.anki.cardFormats)
          ? targetProfile.options.anki.cardFormats
          : [];
        for (const cardFormat of cardFormats) {
          if (
            !cardFormat ||
            typeof cardFormat !== "object" ||
            cardFormat.type !== "term" ||
            cardFormat.enabled === false
          ) {
            continue;
          }
          const currentDeck = typeof cardFormat.deck === "string" ? cardFormat.deck.trim() : "";
          if (currentDeck !== targetDeck) {
            cardFormat.deck = targetDeck;
            changed = true;
          }
        }

        const terms = targetProfile.options.anki.terms;
        if (terms && typeof terms === "object") {
          const currentTermDeck = typeof terms.deck === "string" ? terms.deck.trim() : "";
          if (currentTermDeck !== targetDeck) {
            terms.deck = targetDeck;
            changed = true;
          }
        }
      }

      if (!changed) {
        return { updated: false, matched: true, reason: "already-target", currentServer, targetServer, targetDeck };
      }

      await invoke("setAllSettings", { value: optionsFull, source: "subminer" });
      return { updated: true, matched: true, currentServer, targetServer, targetDeck };
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

function readDeckName(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function getYomitanDeckFromProfileOptions(profileOptions: Record<string, unknown>): string {
  const anki = profileOptions.anki;
  if (!isObject(anki)) {
    return '';
  }

  const cardFormats = Array.isArray(anki.cardFormats) ? anki.cardFormats : [];
  const enabledCardFormats = cardFormats
    .filter((cardFormat): cardFormat is Record<string, unknown> => isObject(cardFormat))
    .filter((cardFormat) => cardFormat.enabled !== false);

  const termDeck = enabledCardFormats.find(
    (cardFormat) => cardFormat.type === 'term' && readDeckName(cardFormat.deck).length > 0,
  );
  if (termDeck) {
    return readDeckName(termDeck.deck);
  }

  const firstDeck = enabledCardFormats
    .map((cardFormat) => readDeckName(cardFormat.deck))
    .find((deckName) => deckName.length > 0);
  if (firstDeck) {
    return firstDeck;
  }

  const terms = anki.terms;
  if (isObject(terms)) {
    const legacyTermDeck = readDeckName(terms.deck);
    if (legacyTermDeck) {
      return legacyTermDeck;
    }
  }

  const kanji = anki.kanji;
  return isObject(kanji) ? readDeckName(kanji.deck) : '';
}

export function extractYomitanCurrentAnkiDeckName(optionsFull: Record<string, unknown>): string {
  const profiles = Array.isArray(optionsFull.profiles) ? optionsFull.profiles : [];
  if (profiles.length === 0) {
    return '';
  }

  const profileCurrent =
    typeof optionsFull.profileCurrent === 'number' && Number.isFinite(optionsFull.profileCurrent)
      ? Math.max(0, Math.floor(optionsFull.profileCurrent))
      : 0;
  const targetProfile = profiles[profileCurrent];
  if (!isObject(targetProfile) || !isObject(targetProfile.options)) {
    return '';
  }

  return getYomitanDeckFromProfileOptions(targetProfile.options as Record<string, unknown>);
}

export async function getYomitanCurrentAnkiDeckName(
  deps: YomitanParserRuntimeDeps,
  logger: LoggerLike,
): Promise<string> {
  const optionsFull = await getYomitanSettingsFull(deps, logger);
  return optionsFull ? extractYomitanCurrentAnkiDeckName(optionsFull) : '';
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
