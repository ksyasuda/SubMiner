import { readFile, writeFile, rename, rm, mkdir } from 'node:fs/promises';
import path from 'node:path';
import type { BridgePreference } from './types';

/**
 * Persists each source's preference array verbatim, keyed by bridge source id.
 *
 * Extensions keep credentials in here (the Jellyfin source stores a password),
 * so the file is written with owner-only permissions.
 */
export class PreferenceStore {
  private readonly file: string;
  private cache: Record<string, BridgePreference[]> | null = null;
  /**
   * Mutations run one at a time. Two concurrent load-modify-persist cycles
   * starting on a cold cache would each read their own object, and the later
   * write would drop the earlier one's edit.
   */
  private queue: Promise<unknown> = Promise.resolve();

  constructor(file: string) {
    this.file = file;
  }

  private enqueue<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.queue.then(operation, operation);
    // Keep the chain alive after a rejection so one failure cannot wedge it.
    this.queue = result.catch(() => undefined);
    return result;
  }

  private async load(): Promise<Record<string, BridgePreference[]>> {
    if (this.cache !== null) return this.cache;
    try {
      const parsed = JSON.parse(await readFile(this.file, 'utf8')) as unknown;
      this.cache =
        parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed)
          ? (parsed as Record<string, BridgePreference[]>)
          : {};
    } catch {
      // Missing or corrupt file starts empty rather than blocking the browser.
      this.cache = {};
    }
    return this.cache;
  }

  async get(sourceId: string): Promise<BridgePreference[]> {
    return this.enqueue(async () => {
      const all = await this.load();
      return all[sourceId] ?? [];
    });
  }

  async set(sourceId: string, preferences: BridgePreference[]): Promise<void> {
    await this.enqueue(async () => {
      const all = await this.load();
      all[sourceId] = preferences;
      await this.persist(all);
    });
  }

  /**
   * Drop every saved value whose key starts with `prefix`.
   *
   * Removing an extension should not leave its credentials on disk, and a
   * source id is not knowable once the APK is gone — so callers pass the
   * package name and this clears anything recorded under it.
   */
  async clear(prefix: string): Promise<void> {
    await this.enqueue(async () => {
      const all = await this.load();
      let changed = false;
      for (const key of Object.keys(all)) {
        if (key === prefix || key.startsWith(`${prefix}:`)) {
          delete all[key];
          changed = true;
        }
      }
      if (changed) await this.persist(all);
    });
  }

  /**
   * Write through a temporary file and rename into place.
   *
   * A write interrupted partway would otherwise leave truncated JSON, and
   * `load()` treats unparseable content as empty — which would quietly discard
   * every saved credential.
   */
  private async persist(all: Record<string, BridgePreference[]>): Promise<void> {
    await mkdir(path.dirname(this.file), { recursive: true });
    const temporary = `${this.file}.tmp`;
    try {
      await writeFile(temporary, JSON.stringify(all, null, 2), { mode: 0o600 });
      await rename(temporary, this.file);
    } catch (error) {
      await rm(temporary, { force: true }).catch(() => undefined);
      throw error;
    }
  }
}
