import fs from 'fs';
import path from 'path';

import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import { getConfiguredWordFieldName } from '../anki-field-config';
import { AnkiConnectConfig } from '../types/anki';
import type { KnownWordMaturityTier } from '../types/subtitle';
import { createLogger } from '../logger';
import {
  KNOWN_WORD_MATURITY_RULES_VERSION,
  classifyKnownWordNoteTier,
  fetchKnownWordMaturityTierSets,
  getKnownWordMaturityEnabled,
  getMatureIntervalThresholdDays,
  maxKnownWordMaturityTier,
  sanitizeKnownWordMaturityTier,
} from './known-word-maturity';
import {
  DEFAULT_KNOWN_WORD_READING_FIELDS,
  KnownWordEntry,
  convertKatakanaToHiragana,
  isReadingFieldName,
  knownWordEntryListsEqual,
  normalizeKnownReadingForLookup,
  normalizeKnownWordEntryList,
  parseFuriganaAnnotatedText,
} from './known-word-entries';

const log = createLogger('anki').child('integration.known-word-cache');

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export function getKnownWordCacheRefreshIntervalMinutes(config: AnkiConnectConfig): number {
  const refreshMinutes = config.knownWords?.refreshMinutes;
  return typeof refreshMinutes === 'number' && Number.isFinite(refreshMinutes) && refreshMinutes > 0
    ? refreshMinutes
    : DEFAULT_ANKI_CONNECT_CONFIG.knownWords.refreshMinutes;
}

export function getKnownWordCacheScopeForConfig(config: AnkiConnectConfig): string {
  const configuredDecks = config.knownWords?.decks;
  if (configuredDecks && typeof configuredDecks === 'object' && !Array.isArray(configuredDecks)) {
    const normalizedDecks = Object.entries(configuredDecks)
      .map(([deckName, fields]) => {
        const name = trimToNonEmptyString(deckName);
        if (!name) return null;
        const normalizedFields = Array.isArray(fields)
          ? [
              ...new Set(
                fields
                  .map(String)
                  .map(trimToNonEmptyString)
                  .filter((field): field is string => Boolean(field)),
              ),
            ].sort()
          : [];
        return [name, normalizedFields];
      })
      .filter((entry): entry is [string, string[]] => entry !== null)
      .sort(([a], [b]) => a.localeCompare(b));
    if (normalizedDecks.length > 0) {
      return `decks:${JSON.stringify(normalizedDecks)}`;
    }
  }

  const configuredDeck = trimToNonEmptyString(config.deck);
  return configuredDeck ? `deck:${configuredDeck}` : 'all';
}

export function getKnownWordCacheLifecycleConfig(config: AnkiConnectConfig): string {
  const payload: Record<string, unknown> = {
    refreshMinutes: getKnownWordCacheRefreshIntervalMinutes(config),
    scope: getKnownWordCacheScopeForConfig(config),
    fieldsWord: trimToNonEmptyString(config.fields?.word) ?? '',
  };
  // The maturity fields are only added while enabled so persisted caches from
  // before the feature existed (or with it off) keep their identity.
  // maturityRules is the classification-rule version: bump it whenever the tier
  // queries change meaning so existing caches refetch instead of serving tiers
  // computed under the old rules.
  if (getKnownWordMaturityEnabled(config)) {
    payload.maturity = getMatureIntervalThresholdDays(config);
    payload.maturityRules = KNOWN_WORD_MATURITY_RULES_VERSION;
  }
  return JSON.stringify(payload);
}

export interface KnownWordCacheNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

interface KnownWordCacheStateV1 {
  readonly version: 1;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly words: string[];
}

interface KnownWordCacheStateV2 {
  readonly version: 2;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly words: string[];
  readonly notes: Record<string, string[]>;
}

interface KnownWordCacheStateV3 {
  readonly version: 3;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly notes: Record<string, KnownWordEntry[]>;
}

interface KnownWordCacheStateV4 {
  readonly version: 4;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly notes: Record<string, KnownWordEntry[]>;
  readonly tiers: Record<string, KnownWordMaturityTier>;
}

type KnownWordCacheState =
  | KnownWordCacheStateV1
  | KnownWordCacheStateV2
  | KnownWordCacheStateV3
  | KnownWordCacheStateV4;

const NO_READING_KEY = '';

interface KnownWordCacheClient {
  findNotes: (
    query: string,
    options?: {
      maxRetries?: number;
    },
  ) => Promise<unknown>;
  notesInfo: (noteIds: number[]) => Promise<unknown>;
}

interface KnownWordCacheDeps {
  client: KnownWordCacheClient;
  getConfig: () => AnkiConnectConfig;
  knownWordCacheStatePath?: string;
  showStatusNotification: (message: string) => void;
}

type KnownWordQueryScope = {
  query: string;
  fields: string[];
};

export class KnownWordCacheManager {
  private knownWordsLastRefreshedAtMs = 0;
  private knownWordsStateKey = '';
  // word → (hiragana reading | NO_READING_KEY → note ids). NO_READING_KEY
  // entries fail open: the word matches regardless of the token's reading.
  private wordReadingNoteIds = new Map<string, Map<string, Set<number>>>();
  // hiragana reading → note ids, so kana tokens still match by reading alone.
  private readingNoteIds = new Map<string, Set<number>>();
  private noteEntriesById = new Map<number, KnownWordEntry[]>();
  private noteTierById = new Map<number, KnownWordMaturityTier>();
  private knownWordsRefreshTimer: ReturnType<typeof setInterval> | null = null;
  private knownWordsRefreshTimeout: ReturnType<typeof setTimeout> | null = null;
  private isRefreshingKnownWords = false;
  private readonly statePath: string;

  constructor(private readonly deps: KnownWordCacheDeps) {
    this.statePath = path.normalize(
      deps.knownWordCacheStatePath || path.join(process.cwd(), 'known-words-cache.json'),
    );
  }

  isKnownWord(
    text: string,
    reading?: string,
    options?: { allowReadingOnlyMatch?: boolean },
  ): boolean {
    if (!this.isKnownWordCacheEnabled()) {
      return false;
    }

    const normalized = this.normalizeKnownWordForLookup(text);
    if (normalized.length === 0) {
      return false;
    }

    const knownReadings = this.wordReadingNoteIds.get(normalized);
    if (knownReadings && knownReadings.size > 0) {
      const normalizedReading =
        typeof reading === 'string' ? normalizeKnownReadingForLookup(reading) : '';
      return (
        normalizedReading.length === 0 ||
        knownReadings.has(NO_READING_KEY) ||
        knownReadings.has(normalizedReading)
      );
    }

    // Callers that look up a kanji token's reading (not subtitle text) must
    // opt out of the reading-only fallback: readingNoteIds holds readings of
    // every note including kanji words, so 渓谷's けいこく would match a
    // mined 警告/けいこく.
    if (options?.allowReadingOnlyMatch === false) {
      return false;
    }

    // Reading-only fallback, except for single-kana text: particles and
    // interjections (よ, ね, え…) would otherwise borrow the reading of an
    // unrelated note (夜「よ」, 絵「え」) and count as known.
    const hiragana = convertKatakanaToHiragana(normalized);
    if ([...hiragana].length === 1) {
      return false;
    }
    return this.readingNoteIds.has(hiragana);
  }

  // Maturity tier for a matching known word, following the exact matching
  // rules of isKnownWord. A match with no tier data (tier fetch failed or
  // pre-v4 cache) returns null so rendering falls back to the single
  // known-word color.
  getKnownWordTier(
    text: string,
    reading?: string,
    options?: { allowReadingOnlyMatch?: boolean },
  ): KnownWordMaturityTier | null {
    if (!getKnownWordMaturityEnabled(this.deps.getConfig())) {
      return null;
    }

    return this.maxTierForNotes(null, this.getKnownWordMatchNoteIds(text, reading, options));
  }

  // Note ids a known-word lookup matches, using the same matching rules as
  // getKnownWordTier. Exposed for diagnostics (see
  // scripts/verify-known-word-highlights.ts), which audits a rendered tier
  // against the live card data of the notes that produced it.
  getKnownWordMatchNoteIds(
    text: string,
    reading?: string,
    options?: { allowReadingOnlyMatch?: boolean },
  ): Set<number> {
    const matches = new Set<number>();
    const normalized = this.normalizeKnownWordForLookup(text);
    if (normalized.length === 0) {
      return matches;
    }

    const knownReadings = this.wordReadingNoteIds.get(normalized);
    if (knownReadings && knownReadings.size > 0) {
      const normalizedReading =
        typeof reading === 'string' ? normalizeKnownReadingForLookup(reading) : '';
      if (normalizedReading.length === 0) {
        for (const noteIds of knownReadings.values()) {
          for (const noteId of noteIds) {
            matches.add(noteId);
          }
        }
        return matches;
      }
      for (const key of [NO_READING_KEY, normalizedReading]) {
        for (const noteId of knownReadings.get(key) ?? []) {
          matches.add(noteId);
        }
      }
      return matches;
    }

    if (options?.allowReadingOnlyMatch === false) {
      return matches;
    }

    const hiragana = convertKatakanaToHiragana(normalized);
    if ([...hiragana].length === 1) {
      return matches;
    }
    for (const noteId of this.readingNoteIds.get(hiragana) ?? []) {
      matches.add(noteId);
    }
    return matches;
  }

  private maxTierForNotes(
    current: KnownWordMaturityTier | null,
    noteIds: ReadonlySet<number>,
  ): KnownWordMaturityTier | null {
    let tier = current;
    for (const noteId of noteIds) {
      tier = maxKnownWordMaturityTier(tier, this.noteTierById.get(noteId) ?? null);
      if (tier === 'mature') {
        break;
      }
    }
    return tier;
  }

  refresh(force = false): Promise<void> {
    return this.refreshKnownWords(force);
  }

  startLifecycle(): void {
    this.stopLifecycle();
    if (!this.isKnownWordCacheEnabled()) {
      log.info('Known-word cache disabled; clearing local cache state');
      this.clearKnownWordCacheState();
      return;
    }

    const refreshMinutes = this.getKnownWordRefreshIntervalMs() / 60_000;
    const scope = getKnownWordCacheScopeForConfig(this.deps.getConfig());
    log.info(
      'Known-word cache lifecycle enabled',
      `scope=${scope}`,
      `refreshMinutes=${refreshMinutes}`,
      `cachePath=${this.statePath}`,
    );

    this.loadKnownWordCacheState();
    this.scheduleKnownWordRefreshLifecycle();
  }

  stopLifecycle(): void {
    if (this.knownWordsRefreshTimeout) {
      clearTimeout(this.knownWordsRefreshTimeout);
      this.knownWordsRefreshTimeout = null;
    }
    if (this.knownWordsRefreshTimer) {
      clearInterval(this.knownWordsRefreshTimer);
      this.knownWordsRefreshTimer = null;
    }
  }

  appendFromNoteInfo(noteInfo: KnownWordCacheNoteInfo): boolean {
    if (!this.isKnownWordCacheEnabled() || !this.shouldAddMinedWordsImmediately()) {
      return false;
    }

    let didMutateCache = false;
    const currentStateKey = this.getKnownWordCacheStateKey();
    if (this.knownWordsStateKey && this.knownWordsStateKey !== currentStateKey) {
      didMutateCache = this.wordReadingNoteIds.size > 0 || this.noteEntriesById.size > 0;
      this.clearKnownWordCacheState();
    }
    if (!this.knownWordsStateKey) {
      this.knownWordsStateKey = currentStateKey;
    }

    const preferredFields = this.getImmediateAppendFields();
    if (!preferredFields) {
      return didMutateCache;
    }

    const nextEntries = this.extractKnownWordEntriesFromNoteInfo(noteInfo, preferredFields);
    const changed = this.replaceNoteSnapshot(noteInfo.noteId, nextEntries);
    if (!changed) {
      return didMutateCache;
    }

    // A just-mined card has never been reviewed.
    if (
      this.isMaturityTrackingEnabled() &&
      this.noteEntriesById.has(noteInfo.noteId) &&
      !this.noteTierById.has(noteInfo.noteId)
    ) {
      this.noteTierById.set(noteInfo.noteId, 'new');
    }

    if (this.knownWordsLastRefreshedAtMs <= 0) {
      this.knownWordsLastRefreshedAtMs = Date.now();
    }
    this.persistKnownWordCacheState();
    log.info(
      'Known-word cache updated in-session',
      `noteId=${noteInfo.noteId}`,
      `wordCount=${nextEntries.length}`,
      `scope=${getKnownWordCacheScopeForConfig(this.deps.getConfig())}`,
    );
    return true;
  }

  clearKnownWordCacheState(): void {
    this.clearInMemoryState();
    this.knownWordsStateKey = this.getKnownWordCacheStateKey();
    try {
      if (fs.existsSync(this.statePath)) {
        fs.unlinkSync(this.statePath);
      }
    } catch (error) {
      log.warn('Failed to clear known-word cache state:', (error as Error).message);
    }
  }

  private async refreshKnownWords(force = false): Promise<void> {
    if (!this.isKnownWordCacheEnabled()) {
      log.debug('Known-word cache refresh skipped; feature disabled');
      return;
    }
    if (this.isRefreshingKnownWords) {
      log.debug('Known-word cache refresh skipped; already refreshing');
      return;
    }
    if (!force && !this.isKnownWordCacheStale()) {
      log.debug('Known-word cache refresh skipped; cache is fresh');
      return;
    }

    const frozenStateKey = this.getKnownWordCacheStateKey();
    this.isRefreshingKnownWords = true;
    try {
      const noteFieldsById = await this.fetchKnownWordNoteFieldsById();
      const maturityTrackingEnabled = this.isMaturityTrackingEnabled();
      let maturityFetchFailed = false;
      let tierSets = null;
      if (maturityTrackingEnabled) {
        try {
          tierSets = await fetchKnownWordMaturityTierSets(
            (query, options) => this.deps.client.findNotes(query, options),
            this.getKnownWordQueryScopes().map((scope) => scope.query),
            getMatureIntervalThresholdDays(this.deps.getConfig()),
          );
        } catch (error) {
          maturityFetchFailed = true;
          log.warn('Failed to fetch known-word maturity tiers:', (error as Error).message);
        }
      }
      const currentNoteIds = Array.from(noteFieldsById.keys()).sort((a, b) => a - b);

      if (this.noteEntriesById.size === 0) {
        await this.rebuildFromCurrentNotes(currentNoteIds, noteFieldsById);
      } else {
        const currentNoteIdSet = new Set(currentNoteIds);
        for (const noteId of Array.from(this.noteEntriesById.keys())) {
          if (!currentNoteIdSet.has(noteId)) {
            this.removeNoteSnapshot(noteId);
          }
        }

        if (currentNoteIds.length > 0) {
          const noteInfos = await this.fetchKnownWordNotesInfo(currentNoteIds);
          for (const noteInfo of noteInfos) {
            this.replaceNoteSnapshot(
              noteInfo.noteId,
              this.extractKnownWordEntriesFromNoteInfo(
                noteInfo,
                noteFieldsById.get(noteInfo.noteId),
              ),
            );
          }
        }
      }

      this.noteTierById = new Map();
      if (tierSets) {
        for (const noteId of currentNoteIds) {
          this.noteTierById.set(noteId, classifyKnownWordNoteTier(noteId, tierSets));
        }
      }

      this.knownWordsLastRefreshedAtMs = Date.now();
      this.knownWordsStateKey = frozenStateKey;
      this.persistKnownWordCacheState();
      log.info(
        'Known-word cache refreshed',
        `noteCount=${currentNoteIds.length}`,
        `wordCount=${this.wordReadingNoteIds.size}`,
        tierSets
          ? `maturityTiers=${this.noteTierById.size}`
          : maturityFetchFailed
            ? 'maturityTiers=fetch-failed'
            : 'maturityTiers=off',
      );
    } catch (error) {
      log.warn('Failed to refresh known-word cache:', (error as Error).message);
      this.deps.showStatusNotification('AnkiConnect: unable to refresh known words');
    } finally {
      this.isRefreshingKnownWords = false;
    }
  }

  private isKnownWordCacheEnabled(): boolean {
    const config = this.deps.getConfig();
    return config.knownWords?.highlightEnabled === true || config.nPlusOne?.enabled === true;
  }

  private isMaturityTrackingEnabled(): boolean {
    return getKnownWordMaturityEnabled(this.deps.getConfig());
  }

  private shouldAddMinedWordsImmediately(): boolean {
    return this.deps.getConfig().knownWords?.addMinedWordsImmediately !== false;
  }

  private getKnownWordRefreshIntervalMs(): number {
    return getKnownWordCacheRefreshIntervalMinutes(this.deps.getConfig()) * 60_000;
  }

  private getDefaultKnownWordFields(): string[] {
    const configuredWordField = getConfiguredWordFieldName(this.deps.getConfig());
    return this.withDefaultReadingFields([configuredWordField, 'Word']);
  }

  // Reading fields are always probed (even when a deck configures explicit
  // word fields) so entries can carry the reading their note teaches.
  private withDefaultReadingFields(fields: string[]): string[] {
    return [...new Set([...fields, ...DEFAULT_KNOWN_WORD_READING_FIELDS])];
  }

  private getKnownWordDecks(): string[] {
    const configuredDecks = this.deps.getConfig().knownWords?.decks;
    if (configuredDecks && typeof configuredDecks === 'object' && !Array.isArray(configuredDecks)) {
      return Object.keys(configuredDecks)
        .map((d) => d.trim())
        .filter((d) => d.length > 0);
    }

    const deck = this.deps.getConfig().deck?.trim();
    return deck ? [deck] : [];
  }

  private getConfiguredFields(): string[] {
    return this.getDefaultKnownWordFields();
  }

  private getImmediateAppendFields(): string[] | null {
    const configuredDecks = this.deps.getConfig().knownWords?.decks;
    if (configuredDecks && typeof configuredDecks === 'object' && !Array.isArray(configuredDecks)) {
      const trimmedDeckEntries = Object.entries(configuredDecks)
        .map(([deckName, fields]) => [deckName.trim(), fields] as const)
        .filter(([deckName]) => deckName.length > 0);

      const currentDeck = this.deps.getConfig().deck?.trim();
      const selectedDeckEntry =
        currentDeck !== undefined && currentDeck.length > 0
          ? (trimmedDeckEntries.find(([deckName]) => deckName === currentDeck) ?? null)
          : trimmedDeckEntries.length === 1
            ? (trimmedDeckEntries[0] ?? null)
            : null;

      if (!selectedDeckEntry) {
        const configuredFields = trimmedDeckEntries.flatMap(([, fields]) =>
          Array.isArray(fields) ? fields : [],
        );
        const normalizedFields = [
          ...new Set(
            configuredFields
              .map(String)
              .map((field) => field.trim())
              .filter((field) => field.length > 0),
          ),
        ];
        return normalizedFields.length > 0
          ? this.withDefaultReadingFields(normalizedFields)
          : this.getDefaultKnownWordFields();
      }

      const deckFields = selectedDeckEntry[1];
      if (Array.isArray(deckFields)) {
        const normalizedFields = [
          ...new Set(
            deckFields
              .map(String)
              .map((field) => field.trim())
              .filter((field) => field.length > 0),
          ),
        ];
        if (normalizedFields.length > 0) {
          return this.withDefaultReadingFields(normalizedFields);
        }
      }

      return this.getDefaultKnownWordFields();
    }

    return this.getConfiguredFields();
  }

  private getKnownWordQueryScopes(): KnownWordQueryScope[] {
    const configuredDecks = this.deps.getConfig().knownWords?.decks;
    if (configuredDecks && typeof configuredDecks === 'object' && !Array.isArray(configuredDecks)) {
      const scopes: KnownWordQueryScope[] = [];
      for (const [deckName, fields] of Object.entries(configuredDecks)) {
        const trimmedDeckName = deckName.trim();
        if (!trimmedDeckName) {
          continue;
        }
        const normalizedFields = Array.isArray(fields)
          ? [
              ...new Set(
                fields
                  .map(String)
                  .map((field) => field.trim())
                  .filter(Boolean),
              ),
            ]
          : [];
        scopes.push({
          query: `deck:"${escapeAnkiSearchValue(trimmedDeckName)}"`,
          fields:
            normalizedFields.length > 0
              ? this.withDefaultReadingFields(normalizedFields)
              : this.getDefaultKnownWordFields(),
        });
      }
      if (scopes.length > 0) {
        return scopes;
      }
    }

    return [{ query: this.buildKnownWordsQuery(), fields: this.getDefaultKnownWordFields() }];
  }

  private buildKnownWordsQuery(): string {
    const decks = this.getKnownWordDecks();
    if (decks.length === 0) {
      return '';
    }

    if (decks.length === 1) {
      return `deck:"${escapeAnkiSearchValue(decks[0]!)}"`;
    }

    const deckQueries = decks.map((deck) => `deck:"${escapeAnkiSearchValue(deck)}"`);
    return `(${deckQueries.join(' OR ')})`;
  }

  private getKnownWordCacheStateKey(): string {
    return getKnownWordCacheLifecycleConfig(this.deps.getConfig());
  }

  private isKnownWordCacheStale(): boolean {
    if (!this.isKnownWordCacheEnabled()) {
      return true;
    }
    if (this.knownWordsStateKey !== this.getKnownWordCacheStateKey()) {
      return true;
    }
    if (this.knownWordsLastRefreshedAtMs <= 0) {
      return true;
    }
    return Date.now() - this.knownWordsLastRefreshedAtMs >= this.getKnownWordRefreshIntervalMs();
  }

  private async fetchKnownWordNoteFieldsById(): Promise<Map<number, string[]>> {
    const scopes = this.getKnownWordQueryScopes();
    const noteFieldsById = new Map<number, string[]>();
    log.debug(
      'Refreshing known-word cache',
      `queries=${scopes.map((scope) => scope.query).join(' | ')}`,
    );

    for (const scope of scopes) {
      const noteIds = (await this.deps.client.findNotes(scope.query, {
        maxRetries: 0,
      })) as number[];

      for (const noteId of noteIds) {
        if (!Number.isInteger(noteId) || noteId <= 0) {
          continue;
        }
        const existingFields = noteFieldsById.get(noteId) ?? [];
        noteFieldsById.set(noteId, [...new Set([...existingFields, ...scope.fields])]);
      }
    }

    return noteFieldsById;
  }

  private scheduleKnownWordRefreshLifecycle(): void {
    const refreshIntervalMs = this.getKnownWordRefreshIntervalMs();
    const scheduleInterval = () => {
      this.knownWordsRefreshTimer = setInterval(() => {
        void this.refreshKnownWords();
      }, refreshIntervalMs);
    };

    const initialDelayMs = this.getMsUntilNextRefresh();
    this.knownWordsRefreshTimeout = setTimeout(() => {
      this.knownWordsRefreshTimeout = null;
      void this.refreshKnownWords();
      scheduleInterval();
    }, initialDelayMs);
  }

  private getMsUntilNextRefresh(): number {
    if (this.knownWordsStateKey !== this.getKnownWordCacheStateKey()) {
      return 0;
    }
    if (this.knownWordsLastRefreshedAtMs <= 0) {
      return 0;
    }
    const remainingMs =
      this.getKnownWordRefreshIntervalMs() - (Date.now() - this.knownWordsLastRefreshedAtMs);
    return Math.max(0, remainingMs);
  }

  private async rebuildFromCurrentNotes(
    noteIds: number[],
    noteFieldsById: Map<number, string[]>,
  ): Promise<void> {
    this.clearInMemoryState();
    if (noteIds.length === 0) {
      return;
    }

    const noteInfos = await this.fetchKnownWordNotesInfo(noteIds);
    for (const noteInfo of noteInfos) {
      this.replaceNoteSnapshot(
        noteInfo.noteId,
        this.extractKnownWordEntriesFromNoteInfo(noteInfo, noteFieldsById.get(noteInfo.noteId)),
      );
    }
  }

  private async fetchKnownWordNotesInfo(noteIds: number[]): Promise<KnownWordCacheNoteInfo[]> {
    const noteInfos: KnownWordCacheNoteInfo[] = [];
    const chunkSize = 50;
    for (let i = 0; i < noteIds.length; i += chunkSize) {
      const chunk = noteIds.slice(i, i + chunkSize);
      const notesInfoResult = (await this.deps.client.notesInfo(chunk)) as unknown[];
      const chunkInfos = notesInfoResult as KnownWordCacheNoteInfo[];
      for (const noteInfo of chunkInfos) {
        if (
          !noteInfo ||
          !Number.isInteger(noteInfo.noteId) ||
          noteInfo.noteId <= 0 ||
          typeof noteInfo.fields !== 'object' ||
          noteInfo.fields === null ||
          Array.isArray(noteInfo.fields)
        ) {
          continue;
        }
        noteInfos.push(noteInfo);
      }
    }
    return noteInfos;
  }

  private replaceNoteSnapshot(noteId: number, nextEntries: KnownWordEntry[]): boolean {
    const normalizedEntries = normalizeKnownWordEntryList(nextEntries);
    const previousEntries = this.noteEntriesById.get(noteId) ?? [];
    if (knownWordEntryListsEqual(previousEntries, normalizedEntries)) {
      return false;
    }

    this.removeEntriesFromIndexes(noteId, previousEntries);
    if (normalizedEntries.length > 0) {
      this.noteEntriesById.set(noteId, normalizedEntries);
      this.addEntriesToIndexes(noteId, normalizedEntries);
    } else {
      this.noteEntriesById.delete(noteId);
      this.noteTierById.delete(noteId);
    }
    return true;
  }

  private removeNoteSnapshot(noteId: number): void {
    const previousEntries = this.noteEntriesById.get(noteId);
    if (!previousEntries) {
      return;
    }
    this.noteEntriesById.delete(noteId);
    this.noteTierById.delete(noteId);
    this.removeEntriesFromIndexes(noteId, previousEntries);
  }

  private addEntriesToIndexes(noteId: number, entries: KnownWordEntry[]): void {
    for (const entry of entries) {
      const readingKey = entry.reading ?? NO_READING_KEY;
      let readings = this.wordReadingNoteIds.get(entry.word);
      if (!readings) {
        readings = new Map();
        this.wordReadingNoteIds.set(entry.word, readings);
      }
      let noteIds = readings.get(readingKey);
      if (!noteIds) {
        noteIds = new Set();
        readings.set(readingKey, noteIds);
      }
      noteIds.add(noteId);
      if (entry.reading) {
        let readingNotes = this.readingNoteIds.get(entry.reading);
        if (!readingNotes) {
          readingNotes = new Set();
          this.readingNoteIds.set(entry.reading, readingNotes);
        }
        readingNotes.add(noteId);
      }
    }
  }

  private removeEntriesFromIndexes(noteId: number, entries: KnownWordEntry[]): void {
    for (const entry of entries) {
      const readingKey = entry.reading ?? NO_READING_KEY;
      const readings = this.wordReadingNoteIds.get(entry.word);
      if (readings) {
        const noteIds = readings.get(readingKey);
        if (noteIds) {
          noteIds.delete(noteId);
          if (noteIds.size === 0) {
            readings.delete(readingKey);
            if (readings.size === 0) {
              this.wordReadingNoteIds.delete(entry.word);
            }
          }
        }
      }
      if (entry.reading) {
        const readingNotes = this.readingNoteIds.get(entry.reading);
        if (readingNotes) {
          readingNotes.delete(noteId);
          if (readingNotes.size === 0) {
            this.readingNoteIds.delete(entry.reading);
          }
        }
      }
    }
  }

  private clearInMemoryState(): void {
    this.wordReadingNoteIds = new Map();
    this.readingNoteIds = new Map();
    this.noteEntriesById = new Map();
    this.noteTierById = new Map();
    this.knownWordsLastRefreshedAtMs = 0;
  }

  private loadKnownWordCacheState(): void {
    try {
      if (!fs.existsSync(this.statePath)) {
        this.clearInMemoryState();
        this.knownWordsStateKey = this.getKnownWordCacheStateKey();
        return;
      }

      const raw = fs.readFileSync(this.statePath, 'utf-8');
      if (!raw.trim()) {
        this.clearInMemoryState();
        this.knownWordsStateKey = this.getKnownWordCacheStateKey();
        return;
      }

      const parsed = JSON.parse(raw) as unknown;
      if (!this.isKnownWordCacheStateValid(parsed)) {
        this.clearInMemoryState();
        this.knownWordsStateKey = this.getKnownWordCacheStateKey();
        return;
      }

      if (parsed.scope !== this.getKnownWordCacheStateKey()) {
        this.clearInMemoryState();
        this.knownWordsStateKey = this.getKnownWordCacheStateKey();
        return;
      }

      this.clearInMemoryState();
      if (parsed.version === 3 || parsed.version === 4) {
        for (const [noteIdKey, entries] of Object.entries(parsed.notes)) {
          const noteId = Number.parseInt(noteIdKey, 10);
          if (!Number.isInteger(noteId) || noteId <= 0) {
            continue;
          }
          const normalizedEntries = normalizeKnownWordEntryList(entries);
          if (normalizedEntries.length === 0) {
            continue;
          }
          this.noteEntriesById.set(noteId, normalizedEntries);
          this.addEntriesToIndexes(noteId, normalizedEntries);
        }
        if (parsed.version === 4) {
          for (const [noteIdKey, tier] of Object.entries(parsed.tiers)) {
            const noteId = Number.parseInt(noteIdKey, 10);
            const sanitizedTier = sanitizeKnownWordMaturityTier(tier);
            if (sanitizedTier && this.noteEntriesById.has(noteId)) {
              this.noteTierById.set(noteId, sanitizedTier);
            }
          }
        }
        this.knownWordsLastRefreshedAtMs = parsed.refreshedAtMs;
        this.knownWordsStateKey = parsed.scope;
        return;
      }

      if (parsed.version === 2) {
        // Older states have no readings; load them reading-less (fail-open,
        // matching the old behavior) but leave the cache marked stale so the
        // next refresh upgrades entries with readings from Anki.
        for (const [noteIdKey, words] of Object.entries(parsed.notes)) {
          const noteId = Number.parseInt(noteIdKey, 10);
          if (!Number.isInteger(noteId) || noteId <= 0) {
            continue;
          }
          const normalizedEntries = normalizeKnownWordEntryList(
            words.map((word) => ({ word: this.normalizeKnownWordForLookup(word), reading: null })),
          );
          if (normalizedEntries.length === 0) {
            continue;
          }
          this.noteEntriesById.set(noteId, normalizedEntries);
          this.addEntriesToIndexes(noteId, normalizedEntries);
        }
        this.knownWordsStateKey = parsed.scope;
        return;
      }

      // v1 has no per-note snapshots to convert; refetch from Anki.
      this.knownWordsStateKey = this.getKnownWordCacheStateKey();
    } catch (error) {
      log.warn('Failed to load known-word cache state:', (error as Error).message);
      this.clearInMemoryState();
      this.knownWordsStateKey = this.getKnownWordCacheStateKey();
    }
  }

  private persistKnownWordCacheState(): void {
    try {
      const notes: Record<string, KnownWordEntry[]> = {};
      const tiers: Record<string, KnownWordMaturityTier> = {};
      for (const [noteId, entries] of this.noteEntriesById.entries()) {
        if (entries.length > 0) {
          notes[String(noteId)] = entries;
          const tier = this.noteTierById.get(noteId);
          if (tier) {
            tiers[String(noteId)] = tier;
          }
        }
      }

      const state: KnownWordCacheStateV4 = {
        version: 4,
        refreshedAtMs: this.knownWordsLastRefreshedAtMs,
        scope: this.knownWordsStateKey,
        notes,
        tiers,
      };
      fs.writeFileSync(this.statePath, JSON.stringify(state), 'utf-8');
    } catch (error) {
      log.warn('Failed to persist known-word cache state:', (error as Error).message);
    }
  }

  private isKnownWordCacheStateValid(value: unknown): value is KnownWordCacheState {
    if (typeof value !== 'object' || value === null) return false;
    const candidate = value as Record<string, unknown>;
    if (
      candidate.version !== 1 &&
      candidate.version !== 2 &&
      candidate.version !== 3 &&
      candidate.version !== 4
    ) {
      return false;
    }
    if (typeof candidate.refreshedAtMs !== 'number') return false;
    if (typeof candidate.scope !== 'string') return false;
    if (candidate.version === 1 || candidate.version === 2) {
      if (!Array.isArray(candidate.words)) return false;
      if (!candidate.words.every((entry: unknown) => typeof entry === 'string')) {
        return false;
      }
    }
    if (candidate.version === 4) {
      // Per-tier values are sanitized entry-by-entry at load time.
      if (
        typeof candidate.tiers !== 'object' ||
        candidate.tiers === null ||
        Array.isArray(candidate.tiers)
      ) {
        return false;
      }
    }
    if (candidate.version === 2 || candidate.version === 3 || candidate.version === 4) {
      if (
        typeof candidate.notes !== 'object' ||
        candidate.notes === null ||
        Array.isArray(candidate.notes)
      ) {
        return false;
      }
      const isValidNoteEntry =
        candidate.version === 2
          ? (entry: unknown): boolean => typeof entry === 'string'
          : (entry: unknown): boolean =>
              typeof entry === 'object' &&
              entry !== null &&
              typeof (entry as KnownWordEntry).word === 'string' &&
              ((entry as KnownWordEntry).reading === null ||
                typeof (entry as KnownWordEntry).reading === 'string');
      if (
        !Object.values(candidate.notes as Record<string, unknown>).every(
          (noteEntries) => Array.isArray(noteEntries) && noteEntries.every(isValidNoteEntry),
        )
      ) {
        return false;
      }
    }
    return true;
  }

  private extractKnownWordEntriesFromNoteInfo(
    noteInfo: KnownWordCacheNoteInfo,
    preferredFields = this.getConfiguredFields(),
  ): KnownWordEntry[] {
    const wordValues: string[] = [];
    let noteReading: string | null = null;
    for (const preferredField of preferredFields) {
      const fieldName = resolveFieldName(Object.keys(noteInfo.fields), preferredField);
      if (!fieldName) continue;

      const raw = noteInfo.fields[fieldName]?.value;
      if (!raw) continue;

      const cleaned = this.normalizeRawKnownWordValue(raw);
      if (!cleaned) continue;

      if (isReadingFieldName(preferredField)) {
        const normalizedReading = normalizeKnownReadingForLookup(cleaned);
        if (normalizedReading) {
          noteReading ??= normalizedReading;
          continue;
        }
        // Non-kana content in a reading field: treat it as a word so decks
        // with repurposed reading fields keep matching (fail-open).
      }
      wordValues.push(cleaned);
    }

    const entries: KnownWordEntry[] = [];
    for (const value of wordValues) {
      const parsed = parseFuriganaAnnotatedText(value);
      const word = parsed.text.trim().toLowerCase();
      if (!word) continue;
      const inlineReading = parsed.reading ? normalizeKnownReadingForLookup(parsed.reading) : '';
      entries.push({ word, reading: inlineReading || noteReading });
    }

    // Kana-only notes (reading field but no word field) stay matchable.
    if (entries.length === 0 && noteReading) {
      entries.push({ word: noteReading, reading: noteReading });
    }

    return normalizeKnownWordEntryList(entries);
  }

  private normalizeRawKnownWordValue(value: string): string {
    return value
      .replace(/<[^>]*>/g, '')
      .replace(/\u3000/g, ' ')
      .trim();
  }

  private normalizeKnownWordForLookup(value: string): string {
    return this.normalizeRawKnownWordValue(value).toLowerCase();
  }
}

function resolveFieldName(availableFieldNames: string[], preferredName: string): string | null {
  const exact = availableFieldNames.find((name) => name === preferredName);
  if (exact) return exact;

  const lower = preferredName.toLowerCase();
  return availableFieldNames.find((name) => name.toLowerCase() === lower) || null;
}

function escapeAnkiSearchValue(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/\"/g, '\\"')
    .replace(/([:*?()\[\]{}])/g, '\\$1');
}
