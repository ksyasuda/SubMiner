import fs from 'fs';
import path from 'path';

import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import { AnkiConnectConfig } from '../types';
import { createLogger } from '../logger';

const log = createLogger('anki').child('integration.known-word-cache');

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
  private knownWordsScope = '';
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
    const scope = this.getKnownWordCacheScope();
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

    const currentScope = this.getKnownWordCacheScope();
    if (this.knownWordsScope && this.knownWordsScope !== currentScope) {
      this.clearKnownWordCacheState();
    }
    if (!this.knownWordsScope) {
      this.knownWordsScope = currentScope;
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
        `scope=${currentScope}`,
      );
    }
  }

  clearKnownWordCacheState(): void {
    this.knownWords = new Set();
    this.knownWordsLastRefreshedAtMs = 0;
    this.knownWordsScope = this.getKnownWordCacheScope();
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
      this.knownWordsScope = this.getKnownWordCacheScope();
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
    const minutes = this.deps.getConfig().knownWords?.refreshMinutes;
    const safeMinutes =
      typeof minutes === 'number' && Number.isFinite(minutes) && minutes > 0
        ? minutes
        : DEFAULT_ANKI_CONNECT_CONFIG.knownWords.refreshMinutes;
    return safeMinutes * 60_000;
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
    return ['Expression', 'Word', 'Reading', 'Word Reading'];
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

  private getKnownWordCacheScope(): string {
    const decks = this.getKnownWordDecks();
    if (decks.length === 0) {
      return 'is:note';
    }
    return `decks:${JSON.stringify(decks)}`;
  }

  private isKnownWordCacheStale(): boolean {
    if (!this.isKnownWordCacheEnabled()) {
      return true;
    }
    if (this.knownWordsScope !== this.getKnownWordCacheScope()) {
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
        this.knownWordsScope = this.getKnownWordCacheScope();
        return;
      }

      const raw = fs.readFileSync(this.statePath, 'utf-8');
      if (!raw.trim()) {
        this.knownWords = new Set();
        this.knownWordsLastRefreshedAtMs = 0;
        this.knownWordsScope = this.getKnownWordCacheScope();
        return;
      }

      const parsed = JSON.parse(raw) as unknown;
      if (!this.isKnownWordCacheStateValid(parsed)) {
        this.knownWords = new Set();
        this.knownWordsLastRefreshedAtMs = 0;
        this.knownWordsScope = this.getKnownWordCacheScope();
        return;
      }

      if (parsed.scope !== this.getKnownWordCacheScope()) {
        this.knownWords = new Set();
        this.knownWordsLastRefreshedAtMs = 0;
        this.knownWordsScope = this.getKnownWordCacheScope();
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
      this.knownWordsScope = parsed.scope;
    } catch (error) {
      log.warn('Failed to load known-word cache state:', (error as Error).message);
      this.knownWords = new Set();
      this.knownWordsLastRefreshedAtMs = 0;
      this.knownWordsScope = this.getKnownWordCacheScope();
    }
  }

  private persistKnownWordCacheState(): void {
    try {
      const state: KnownWordCacheState = {
        version: 1,
        refreshedAtMs: this.knownWordsLastRefreshedAtMs,
        scope: this.knownWordsScope,
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
