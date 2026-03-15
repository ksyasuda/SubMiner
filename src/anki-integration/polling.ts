export interface PollingRunnerDeps {
  getDeck: () => string | undefined;
  getPollingRate: () => number;
  findNotes: (
    query: string,
    options?: {
      maxRetries?: number;
    },
  ) => Promise<number[]>;
  shouldAutoUpdateNewCards: () => boolean;
  processNewCard: (noteId: number) => Promise<void>;
  recordCardsAdded?: (count: number, noteIds: number[]) => void;
  isUpdateInProgress: () => boolean;
  setUpdateInProgress: (value: boolean) => void;
  getTrackedNoteIds: () => Set<number>;
  setTrackedNoteIds: (noteIds: Set<number>) => void;
  showStatusNotification: (message: string) => void;
  logDebug: (...args: unknown[]) => void;
  logInfo: (...args: unknown[]) => void;
  logWarn: (...args: unknown[]) => void;
}

export class PollingRunner {
  private pollingInterval: ReturnType<typeof setInterval> | null = null;
  private initialized = false;
  private backoffMs = 200;
  private maxBackoffMs = 5000;
  private nextPollTime = 0;

  constructor(private readonly deps: PollingRunnerDeps) {}

  get isRunning(): boolean {
    return this.pollingInterval !== null;
  }

  start(): void {
    if (this.pollingInterval) {
      this.stop();
    }

    void this.pollOnce();
    this.pollingInterval = setInterval(() => {
      void this.pollOnce();
    }, this.deps.getPollingRate());
  }

  stop(): void {
    if (this.pollingInterval) {
      clearInterval(this.pollingInterval);
      this.pollingInterval = null;
    }
  }

  async pollOnce(): Promise<void> {
    if (this.deps.isUpdateInProgress()) return;
    if (Date.now() < this.nextPollTime) return;

    this.deps.setUpdateInProgress(true);
    try {
      const query = this.deps.getDeck() ? `"deck:${this.deps.getDeck()}" added:1` : 'added:1';
      const noteIds = await this.deps.findNotes(query, {
        maxRetries: 0,
      });
      const currentNoteIds = new Set(noteIds);

      const previousNoteIds = this.deps.getTrackedNoteIds();
      if (!this.initialized) {
        this.deps.setTrackedNoteIds(currentNoteIds);
        this.initialized = true;
        this.deps.logInfo(`AnkiConnect initialized with ${currentNoteIds.size} existing cards`);
        this.backoffMs = 200;
        return;
      }

      const newNoteIds = Array.from(currentNoteIds).filter((id) => !previousNoteIds.has(id));

      if (newNoteIds.length > 0) {
        this.deps.logInfo('Found new cards:', newNoteIds);

        for (const noteId of newNoteIds) {
          previousNoteIds.add(noteId);
        }
        this.deps.setTrackedNoteIds(previousNoteIds);
        this.deps.recordCardsAdded?.(newNoteIds.length, newNoteIds);

        if (this.deps.shouldAutoUpdateNewCards()) {
          for (const noteId of newNoteIds) {
            await this.deps.processNewCard(noteId);
          }
        } else {
          this.deps.logInfo(
            'New card detected (auto-update disabled). Press Ctrl+V to update from clipboard.',
          );
        }
      }

      if (this.backoffMs > 200) {
        this.deps.logInfo('AnkiConnect connection restored');
      }
      this.backoffMs = 200;
    } catch (error) {
      const wasBackingOff = this.backoffMs > 200;
      this.backoffMs = Math.min(this.backoffMs * 2, this.maxBackoffMs);
      this.nextPollTime = Date.now() + this.backoffMs;
      if (!wasBackingOff) {
        this.deps.logWarn('AnkiConnect polling failed, backing off...');
        this.deps.showStatusNotification('AnkiConnect: unable to connect');
      }
      this.deps.logWarn((error as Error).message);
    } finally {
      this.deps.setUpdateInProgress(false);
    }
  }

  async poll(): Promise<void> {
    if (this.pollingInterval) {
      return;
    }
    return this.pollOnce();
  }
}
