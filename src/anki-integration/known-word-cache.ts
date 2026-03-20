import fs from 'fs';
import path from 'path';

import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import { getConfiguredWordFieldName } from '../anki-field-config';
import { AnkiConnectConfig } from '../types';
import { createLogger } from '../logger';

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
  return configuredDeck ? `deck:${configuredDeck}` : 'is:note';
}

export function getKnownWordCacheLifecycleConfig(config: AnkiConnectConfig): string {
  return JSON.stringify({
    refreshMinutes: getKnownWordCacheRefreshIntervalMinutes(config),
    scope: getKnownWordCacheScopeForConfig(config),
    fieldsWord: trimToNonEmptyString(config.fields?.word) ?? '',
  });
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

type KnownWordCacheState = KnownWordCacheStateV1 | KnownWordCacheStateV2;

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
  private knownWords: Set<string> = new Set();
  private wordReferenceCounts = new Map<string, number>();
  private noteWordsById = new Map<number, string[]>();
  private knownWordsRefreshTimer: ReturnType<typeof setInterval> | null = null;
  private knownWordsRefreshTimeout: ReturnType<typeof setTimeout> | null = null;
  private isRefreshingKnownWords = false;
  private readonly statePath: string;

  constructor(private readonly deps: KnownWordCacheDeps) {
    this.statePath = path.normalize(
      deps.knownWordCacheStatePath || path.join(process.cwd(), 'known-words-cache.json'),
    );
  }

  isKnownWord(text: string): boolean {
    if (!this.isKnownWordCacheEnabled()) {
      return false;
    }

    const normalized = this.normalizeKnownWordForLookup(text);
    return normalized.length > 0 ? this.knownWords.has(normalized) : false;
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

  appendFromNoteInfo(noteInfo: KnownWordCacheNoteInfo): void {
    if (!this.isKnownWordCacheEnabled() || !this.shouldAddMinedWordsImmediately()) {
      return;
    }

    const currentStateKey = this.getKnownWordCacheStateKey();
    if (this.knownWordsStateKey && this.knownWordsStateKey !== currentStateKey) {
      this.clearKnownWordCacheState();
    }
    if (!this.knownWordsStateKey) {
      this.knownWordsStateKey = currentStateKey;
    }

    const preferredFields = this.getImmediateAppendFields();
    if (!preferredFields) {
      return;
    }

    const nextWords = this.extractNormalizedKnownWordsFromNoteInfo(noteInfo, preferredFields);
    const changed = this.replaceNoteSnapshot(noteInfo.noteId, nextWords);
    if (!changed) {
      return;
    }

    if (this.knownWordsLastRefreshedAtMs <= 0) {
      this.knownWordsLastRefreshedAtMs = Date.now();
    }
    this.persistKnownWordCacheState();
    log.info(
      'Known-word cache updated in-session',
      `noteId=${noteInfo.noteId}`,
      `wordCount=${nextWords.length}`,
      `scope=${getKnownWordCacheScopeForConfig(this.deps.getConfig())}`,
    );
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
      const currentNoteIds = Array.from(noteFieldsById.keys()).sort((a, b) => a - b);

      if (this.noteWordsById.size === 0) {
        await this.rebuildFromCurrentNotes(currentNoteIds, noteFieldsById);
      } else {
        const currentNoteIdSet = new Set(currentNoteIds);
        for (const noteId of Array.from(this.noteWordsById.keys())) {
          if (!currentNoteIdSet.has(noteId)) {
            this.removeNoteSnapshot(noteId);
          }
        }

        if (currentNoteIds.length > 0) {
          const noteInfos = await this.fetchKnownWordNotesInfo(currentNoteIds);
          for (const noteInfo of noteInfos) {
            this.replaceNoteSnapshot(
              noteInfo.noteId,
              this.extractNormalizedKnownWordsFromNoteInfo(
                noteInfo,
                noteFieldsById.get(noteInfo.noteId),
              ),
            );
          }
        }
      }

      this.knownWordsLastRefreshedAtMs = Date.now();
      this.knownWordsStateKey = frozenStateKey;
      this.persistKnownWordCacheState();
      log.info(
        'Known-word cache refreshed',
        `noteCount=${currentNoteIds.length}`,
        `wordCount=${this.knownWords.size}`,
      );
    } catch (error) {
      log.warn('Failed to refresh known-word cache:', (error as Error).message);
      this.deps.showStatusNotification('AnkiConnect: unable to refresh known words');
    } finally {
      this.isRefreshingKnownWords = false;
    }
  }

  private isKnownWordCacheEnabled(): boolean {
    return this.deps.getConfig().knownWords?.highlightEnabled === true;
  }

  private shouldAddMinedWordsImmediately(): boolean {
    return this.deps.getConfig().knownWords?.addMinedWordsImmediately !== false;
  }

  private getKnownWordRefreshIntervalMs(): number {
    return getKnownWordCacheRefreshIntervalMinutes(this.deps.getConfig()) * 60_000;
  }

  private getDefaultKnownWordFields(): string[] {
    const configuredWordField = getConfiguredWordFieldName(this.deps.getConfig());
    return [...new Set([configuredWordField, 'Word', 'Reading', 'Word Reading'])];
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
          ? trimmedDeckEntries.find(([deckName]) => deckName === currentDeck) ?? null
          : trimmedDeckEntries.length === 1
            ? trimmedDeckEntries[0] ?? null
            : null;

      if (!selectedDeckEntry) {
        return null;
      }

      const deckFields = selectedDeckEntry[1];
      if (Array.isArray(deckFields)) {
        const normalizedFields = [
          ...new Set(
            deckFields.map(String).map((field) => field.trim()).filter((field) => field.length > 0),
          ),
        ];
        if (normalizedFields.length > 0) {
          return normalizedFields;
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
          ? [...new Set(fields.map(String).map((field) => field.trim()).filter(Boolean))]
          : [];
        scopes.push({
          query: `deck:"${escapeAnkiSearchValue(trimmedDeckName)}"`,
          fields: normalizedFields.length > 0 ? normalizedFields : this.getDefaultKnownWordFields(),
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
      return 'is:note';
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
    log.debug('Refreshing known-word cache', `queries=${scopes.map((scope) => scope.query).join(' | ')}`);

    for (const scope of scopes) {
      const noteIds = (await this.deps.client.findNotes(scope.query, {
        maxRetries: 0,
      })) as number[];

      for (const noteId of noteIds) {
        if (!Number.isInteger(noteId) || noteId <= 0) {
          continue;
        }
        const existingFields = noteFieldsById.get(noteId) ?? [];
        noteFieldsById.set(
          noteId,
          [...new Set([...existingFields, ...scope.fields])],
        );
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
        this.extractNormalizedKnownWordsFromNoteInfo(noteInfo, noteFieldsById.get(noteInfo.noteId)),
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
        if (!noteInfo || !Number.isInteger(noteInfo.noteId) || noteInfo.noteId <= 0) {
          continue;
        }
        noteInfos.push(noteInfo);
      }
    }
    return noteInfos;
  }

  private replaceNoteSnapshot(noteId: number, nextWords: string[]): boolean {
    const normalizedWords = normalizeKnownWordList(nextWords);
    const previousWords = this.noteWordsById.get(noteId) ?? [];
    if (knownWordListsEqual(previousWords, normalizedWords)) {
      return false;
    }

    this.removeWordsFromCounts(previousWords);
    if (normalizedWords.length > 0) {
      this.noteWordsById.set(noteId, normalizedWords);
      this.addWordsToCounts(normalizedWords);
    } else {
      this.noteWordsById.delete(noteId);
    }
    return true;
  }

  private removeNoteSnapshot(noteId: number): void {
    const previousWords = this.noteWordsById.get(noteId);
    if (!previousWords) {
      return;
    }
    this.noteWordsById.delete(noteId);
    this.removeWordsFromCounts(previousWords);
  }

  private addWordsToCounts(words: string[]): void {
    for (const word of words) {
      const nextCount = (this.wordReferenceCounts.get(word) ?? 0) + 1;
      this.wordReferenceCounts.set(word, nextCount);
      this.knownWords.add(word);
    }
  }

  private removeWordsFromCounts(words: string[]): void {
    for (const word of words) {
      const nextCount = (this.wordReferenceCounts.get(word) ?? 0) - 1;
      if (nextCount > 0) {
        this.wordReferenceCounts.set(word, nextCount);
      } else {
        this.wordReferenceCounts.delete(word);
        this.knownWords.delete(word);
      }
    }
  }

  private clearInMemoryState(): void {
    this.knownWords = new Set();
    this.wordReferenceCounts = new Map();
    this.noteWordsById = new Map();
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
      if (parsed.version === 2) {
        for (const [noteIdKey, words] of Object.entries(parsed.notes)) {
          const noteId = Number.parseInt(noteIdKey, 10);
          if (!Number.isInteger(noteId) || noteId <= 0) {
            continue;
          }
          const normalizedWords = normalizeKnownWordList(words);
          if (normalizedWords.length === 0) {
            continue;
          }
          this.noteWordsById.set(noteId, normalizedWords);
          this.addWordsToCounts(normalizedWords);
        }
      } else {
        for (const value of parsed.words) {
          const normalized = this.normalizeKnownWordForLookup(value);
          if (!normalized) {
            continue;
          }
          this.knownWords.add(normalized);
          this.wordReferenceCounts.set(normalized, 1);
        }
      }

      this.knownWordsLastRefreshedAtMs = parsed.refreshedAtMs;
      this.knownWordsStateKey = parsed.scope;
    } catch (error) {
      log.warn('Failed to load known-word cache state:', (error as Error).message);
      this.clearInMemoryState();
      this.knownWordsStateKey = this.getKnownWordCacheStateKey();
    }
  }

  private persistKnownWordCacheState(): void {
    try {
      const notes: Record<string, string[]> = {};
      for (const [noteId, words] of this.noteWordsById.entries()) {
        if (words.length > 0) {
          notes[String(noteId)] = words;
        }
      }

      const state: KnownWordCacheStateV2 = {
        version: 2,
        refreshedAtMs: this.knownWordsLastRefreshedAtMs,
        scope: this.knownWordsStateKey,
        words: Array.from(this.knownWords),
        notes,
      };
      fs.writeFileSync(this.statePath, JSON.stringify(state), 'utf-8');
    } catch (error) {
      log.warn('Failed to persist known-word cache state:', (error as Error).message);
    }
  }

  private isKnownWordCacheStateValid(value: unknown): value is KnownWordCacheState {
    if (typeof value !== 'object' || value === null) return false;
    const candidate = value as Record<string, unknown>;
    if (candidate.version !== 1 && candidate.version !== 2) return false;
    if (typeof candidate.refreshedAtMs !== 'number') return false;
    if (typeof candidate.scope !== 'string') return false;
    if (!Array.isArray(candidate.words)) return false;
    if (!candidate.words.every((entry: unknown) => typeof entry === 'string')) {
      return false;
    }
    if (candidate.version === 2) {
      if (
        typeof candidate.notes !== 'object' ||
        candidate.notes === null ||
        Array.isArray(candidate.notes)
      ) {
        return false;
      }
      if (
        !Object.values(candidate.notes as Record<string, unknown>).every(
          (entry) =>
            Array.isArray(entry) && entry.every((word: unknown) => typeof word === 'string'),
        )
      ) {
        return false;
      }
    }
    return true;
  }

  private extractNormalizedKnownWordsFromNoteInfo(
    noteInfo: KnownWordCacheNoteInfo,
    preferredFields = this.getConfiguredFields(),
  ): string[] {
    const words: string[] = [];
    for (const preferredField of preferredFields) {
      const fieldName = resolveFieldName(Object.keys(noteInfo.fields), preferredField);
      if (!fieldName) continue;

      const raw = noteInfo.fields[fieldName]?.value;
      if (!raw) continue;

      const normalized = this.normalizeKnownWordForLookup(raw);
      if (normalized) {
        words.push(normalized);
      }
    }
    return normalizeKnownWordList(words);
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

function normalizeKnownWordList(words: string[]): string[] {
  return [...new Set(words.map((word) => word.trim()).filter((word) => word.length > 0))].sort();
}

function knownWordListsEqual(left: string[], right: string[]): boolean {
  if (left.length !== right.length) {
    return false;
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false;
    }
  }
  return true;
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
