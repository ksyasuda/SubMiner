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

interface KnownWordCacheState {
  readonly version: 1;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly words: string[];
}

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

export class KnownWordCacheManager {
  private knownWordsLastRefreshedAtMs = 0;
  private knownWordsStateKey = '';
  private knownWords: Set<string> = new Set();
  private knownWordsRefreshTimer: ReturnType<typeof setInterval> | null = null;
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
    void this.refreshKnownWords();
    const refreshIntervalMs = this.getKnownWordRefreshIntervalMs();
    this.knownWordsRefreshTimer = setInterval(() => {
      void this.refreshKnownWords();
    }, refreshIntervalMs);
  }

  stopLifecycle(): void {
    if (this.knownWordsRefreshTimer) {
      clearInterval(this.knownWordsRefreshTimer);
      this.knownWordsRefreshTimer = null;
    }
  }

  appendFromNoteInfo(noteInfo: KnownWordCacheNoteInfo): void {
    if (!this.isKnownWordCacheEnabled()) {
      return;
    }

    const currentStateKey = this.getKnownWordCacheStateKey();
    if (this.knownWordsStateKey && this.knownWordsStateKey !== currentStateKey) {
      this.clearKnownWordCacheState();
    }
    if (!this.knownWordsStateKey) {
      this.knownWordsStateKey = currentStateKey;
    }

    let addedCount = 0;
    for (const rawWord of this.extractKnownWordsFromNoteInfo(noteInfo)) {
      const normalized = this.normalizeKnownWordForLookup(rawWord);
      if (!normalized || this.knownWords.has(normalized)) {
        continue;
      }
      this.knownWords.add(normalized);
      addedCount += 1;
    }

    if (addedCount > 0) {
      if (this.knownWordsLastRefreshedAtMs <= 0) {
        this.knownWordsLastRefreshedAtMs = Date.now();
      }
      this.persistKnownWordCacheState();
      log.info(
        'Known-word cache updated in-session',
        `added=${addedCount}`,
        `scope=${getKnownWordCacheScopeForConfig(this.deps.getConfig())}`,
      );
    }
  }

  clearKnownWordCacheState(): void {
    this.knownWords = new Set();
    this.knownWordsLastRefreshedAtMs = 0;
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

    this.isRefreshingKnownWords = true;
    try {
      const query = this.buildKnownWordsQuery();
      log.debug('Refreshing known-word cache', `query=${query}`);
      const noteIds = (await this.deps.client.findNotes(query, {
        maxRetries: 0,
      })) as number[];

      const nextKnownWords = new Set<string>();
      if (noteIds.length > 0) {
        const chunkSize = 50;
        for (let i = 0; i < noteIds.length; i += chunkSize) {
          const chunk = noteIds.slice(i, i + chunkSize);
          const notesInfoResult = (await this.deps.client.notesInfo(chunk)) as unknown[];
          const notesInfo = notesInfoResult as KnownWordCacheNoteInfo[];

          for (const noteInfo of notesInfo) {
            for (const word of this.extractKnownWordsFromNoteInfo(noteInfo)) {
              const normalized = this.normalizeKnownWordForLookup(word);
              if (normalized) {
                nextKnownWords.add(normalized);
              }
            }
          }
        }
      }

      this.knownWords = nextKnownWords;
      this.knownWordsLastRefreshedAtMs = Date.now();
      this.knownWordsStateKey = this.getKnownWordCacheStateKey();
      this.persistKnownWordCacheState();
      log.info(
        'Known-word cache refreshed',
        `noteCount=${noteIds.length}`,
        `wordCount=${nextKnownWords.size}`,
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

  private getKnownWordRefreshIntervalMs(): number {
    return getKnownWordCacheRefreshIntervalMinutes(this.deps.getConfig()) * 60_000;
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
    const configuredDecks = this.deps.getConfig().knownWords?.decks;
    if (configuredDecks && typeof configuredDecks === 'object' && !Array.isArray(configuredDecks)) {
      const allFields = new Set<string>();
      for (const fields of Object.values(configuredDecks)) {
        if (Array.isArray(fields)) {
          for (const f of fields) {
            if (typeof f === 'string' && f.trim()) allFields.add(f.trim());
          }
        }
      }
      if (allFields.size > 0) return [...allFields];
    }
    const configuredWordField = getConfiguredWordFieldName(this.deps.getConfig());
    return [...new Set([configuredWordField, 'Word', 'Reading', 'Word Reading'])];
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

  private loadKnownWordCacheState(): void {
    try {
      if (!fs.existsSync(this.statePath)) {
        this.knownWords = new Set();
        this.knownWordsLastRefreshedAtMs = 0;
        this.knownWordsStateKey = this.getKnownWordCacheStateKey();
        return;
      }

      const raw = fs.readFileSync(this.statePath, 'utf-8');
      if (!raw.trim()) {
        this.knownWords = new Set();
        this.knownWordsLastRefreshedAtMs = 0;
        this.knownWordsStateKey = this.getKnownWordCacheStateKey();
        return;
      }

      const parsed = JSON.parse(raw) as unknown;
      if (!this.isKnownWordCacheStateValid(parsed)) {
        this.knownWords = new Set();
        this.knownWordsLastRefreshedAtMs = 0;
        this.knownWordsStateKey = this.getKnownWordCacheStateKey();
        return;
      }

      if (parsed.scope !== this.getKnownWordCacheStateKey()) {
        this.knownWords = new Set();
        this.knownWordsLastRefreshedAtMs = 0;
        this.knownWordsStateKey = this.getKnownWordCacheStateKey();
        return;
      }

      const nextKnownWords = new Set<string>();
      for (const value of parsed.words) {
        const normalized = this.normalizeKnownWordForLookup(value);
        if (normalized) {
          nextKnownWords.add(normalized);
        }
      }

      this.knownWords = nextKnownWords;
      this.knownWordsLastRefreshedAtMs = parsed.refreshedAtMs;
      this.knownWordsStateKey = parsed.scope;
    } catch (error) {
      log.warn('Failed to load known-word cache state:', (error as Error).message);
      this.knownWords = new Set();
      this.knownWordsLastRefreshedAtMs = 0;
      this.knownWordsStateKey = this.getKnownWordCacheStateKey();
    }
  }

  private persistKnownWordCacheState(): void {
    try {
      const state: KnownWordCacheState = {
        version: 1,
        refreshedAtMs: this.knownWordsLastRefreshedAtMs,
        scope: this.knownWordsStateKey,
        words: Array.from(this.knownWords),
      };
      fs.writeFileSync(this.statePath, JSON.stringify(state), 'utf-8');
    } catch (error) {
      log.warn('Failed to persist known-word cache state:', (error as Error).message);
    }
  }

  private isKnownWordCacheStateValid(value: unknown): value is KnownWordCacheState {
    if (typeof value !== 'object' || value === null) return false;
    const candidate = value as Partial<KnownWordCacheState>;
    if (candidate.version !== 1) return false;
    if (typeof candidate.refreshedAtMs !== 'number') return false;
    if (typeof candidate.scope !== 'string') return false;
    if (!Array.isArray(candidate.words)) return false;
    if (!candidate.words.every((entry) => typeof entry === 'string')) {
      return false;
    }
    return true;
  }

  private extractKnownWordsFromNoteInfo(noteInfo: KnownWordCacheNoteInfo): string[] {
    const words: string[] = [];
    const configuredFields = this.getConfiguredFields();
    for (const preferredField of configuredFields) {
      const fieldName = resolveFieldName(Object.keys(noteInfo.fields), preferredField);
      if (!fieldName) continue;

      const raw = noteInfo.fields[fieldName]?.value;
      if (!raw) continue;

      const extracted = this.normalizeRawKnownWordValue(raw);
      if (extracted) {
        words.push(extracted);
      }
    }
    return words;
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
