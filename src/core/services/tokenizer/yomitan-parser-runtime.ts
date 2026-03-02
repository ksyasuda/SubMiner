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
  dictionaryPriority: number;
  frequency: number;
  displayValue: string | null;
  displayValueParsed: boolean;
}

export interface YomitanTermReadingPair {
  term: string;
  reading: string | null;
}

interface YomitanProfileMetadata {
  profileIndex: number;
  scanLength: number;
  dictionaries: string[];
  dictionaryPriorityByName: Record<string, number>;
}

const DEFAULT_YOMITAN_SCAN_LENGTH = 40;
const yomitanProfileMetadataByWindow = new WeakMap<BrowserWindow, YomitanProfileMetadata>();
const yomitanFrequencyCacheByWindow = new WeakMap<BrowserWindow, Map<string, YomitanTermFrequency[]>>();

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object');
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

  const numericPrefix = trimmed.match(/^\d[\d,]*/)?.[0];
  if (!numericPrefix) {
    return null;
  }

  const chunks = numericPrefix.split(',');
  const normalizedNumber =
    chunks.length <= 1
      ? chunks[0] ?? ''
      : chunks.slice(1).every((chunk) => /^\d{3}$/.test(chunk))
        ? chunks.join('')
        : (chunks[0] ?? '');
  const parsed = Number.parseInt(normalizedNumber, 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }

  return parsed;
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
      ? parsePositiveFrequencyValue(displayValueRaw)
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
    value.reading === null
      ? null
      : typeof value.reading === 'string'
        ? value.reading
        : null;
  const displayValue = typeof displayValueRaw === 'string' ? displayValueRaw : null;
  const displayValueParsed = value.displayValueParsed === true;

  return {
    term,
    reading,
    dictionary,
    dictionaryPriority,
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

  return {
    profileIndex,
    scanLength,
    dictionaries,
    dictionaryPriorityByName,
  };
}

function normalizeFrequencyEntriesWithPriority(
  rawResult: unknown[],
  dictionaryPriorityByName: Record<string, number>,
): YomitanTermFrequency[] {
  const normalized: YomitanTermFrequency[] = [];
  for (const entry of rawResult) {
    const frequency = toYomitanTermFrequency(entry);
    if (!frequency) {
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
      typeof entry.reading === 'string' && entry.reading.trim().length > 0 ? entry.reading.trim() : null;
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

      return { profileIndex, scanLength, dictionaries, dictionaryPriorityByName };
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
        ? normalizeFrequencyEntriesWithPriority(rawResult, metadata.dictionaryPriorityByName)
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
      .filter((pair): pair is YomitanTermReadingPair => pair !== null),
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
      cacheFrequencyEntriesForPairs(frequencyCache, fallbackTermReadingList, fallbackFetchedEntries);
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
  const unmatchedEntries = allFetchedEntries.filter((entry) => !queriedTerms.has(entry.term.trim()));
  return [...cachedResult, ...unmatchedEntries];
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
