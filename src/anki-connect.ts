/*
 * SubMiner - Subtitle mining overlay for mpv
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import axios, { AxiosInstance } from 'axios';
import http from 'http';
import https from 'https';
import { createLogger } from './logger';

const log = createLogger('anki');

interface AnkiConnectRequest {
  action: string;
  version: number;
  params: Record<string, unknown>;
}

interface AnkiConnectResponse {
  result: unknown;
  error: string | null;
}

export class AnkiConnectClient {
  private client: AxiosInstance;
  private backoffMs = 200;
  private maxBackoffMs = 5000;
  private consecutiveFailures = 0;
  private maxConsecutiveFailures = 5;

  constructor(url: string) {
    const httpAgent = new http.Agent({
      keepAlive: false,
      keepAliveMsecs: 1000,
      maxSockets: 5,
      maxFreeSockets: 2,
      timeout: 10000,
    });

    const httpsAgent = new https.Agent({
      keepAlive: false,
      keepAliveMsecs: 1000,
      maxSockets: 5,
      maxFreeSockets: 2,
      timeout: 10000,
    });

    this.client = axios.create({
      baseURL: url,
      timeout: 10000,
      httpAgent,
      httpsAgent,
    });
  }

  private async sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private isRetryableError(error: unknown): boolean {
    if (!error || typeof error !== 'object') return false;

    const code = (error as Record<string, unknown>).code;
    const message =
      typeof (error as Record<string, unknown>).message === 'string'
        ? ((error as Record<string, unknown>).message as string).toLowerCase()
        : '';

    return (
      code === 'ECONNRESET' ||
      code === 'ETIMEDOUT' ||
      code === 'ENOTFOUND' ||
      code === 'ECONNREFUSED' ||
      code === 'EPIPE' ||
      message.includes('socket hang up') ||
      message.includes('network error') ||
      message.includes('timeout')
    );
  }

  async invoke(
    action: string,
    params: Record<string, unknown> = {},
    options: { timeout?: number; maxRetries?: number } = {},
  ): Promise<unknown> {
    const maxRetries = options.maxRetries ?? 3;
    let lastError: Error | null = null;

    const isMediaUpload = action === 'storeMediaFile';
    const requestTimeout = options.timeout || (isMediaUpload ? 30000 : 10000);

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        if (attempt > 0) {
          const delay = Math.min(this.backoffMs * Math.pow(2, attempt - 1), this.maxBackoffMs);
          log.info(`AnkiConnect ${action} retry ${attempt}/${maxRetries} after ${delay}ms delay`);
          await this.sleep(delay);
        }

        const response = await this.client.post<AnkiConnectResponse>(
          '',
          {
            action,
            version: 6,
            params,
          } as AnkiConnectRequest,
          {
            timeout: requestTimeout,
          },
        );

        this.consecutiveFailures = 0;
        this.backoffMs = 200;

        if (response.data.error) {
          throw new Error(response.data.error);
        }

        return response.data.result;
      } catch (error) {
        lastError = error as Error;
        this.consecutiveFailures++;

        if (!this.isRetryableError(error) || attempt === maxRetries) {
          if (this.consecutiveFailures < this.maxConsecutiveFailures) {
            log.error(
              `AnkiConnect error (attempt ${this.consecutiveFailures}/${this.maxConsecutiveFailures}):`,
              lastError.message,
            );
          } else if (this.consecutiveFailures === this.maxConsecutiveFailures) {
            log.error('AnkiConnect: Too many consecutive failures, suppressing further error logs');
          }
          throw lastError;
        }
      }
    }

    throw lastError || new Error('Unknown error');
  }

  async findNotes(query: string, options?: { maxRetries?: number }): Promise<number[]> {
    const result = await this.invoke('findNotes', { query }, options);
    return (result as number[]) || [];
  }

  async deckNames(): Promise<string[]> {
    const result = await this.invoke('deckNames');
    return Array.isArray(result)
      ? result.filter((value): value is string => typeof value === 'string').sort()
      : [];
  }

  async modelNames(): Promise<string[]> {
    const result = await this.invoke('modelNames');
    return Array.isArray(result)
      ? result.filter((value): value is string => typeof value === 'string').sort()
      : [];
  }

  async modelFieldNames(modelName: string): Promise<string[]> {
    const result = await this.invoke('modelFieldNames', { modelName });
    return Array.isArray(result)
      ? result.filter((value): value is string => typeof value === 'string').sort()
      : [];
  }

  private async noteInfosForDeck(
    deckName: string,
    sampleSize = 100,
  ): Promise<Record<string, unknown>[]> {
    const escapedDeckName = deckName.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    const noteIds = await this.findNotes(`deck:"${escapedDeckName}"`, { maxRetries: 0 });
    if (noteIds.length === 0) {
      return [];
    }

    const finiteSampleSize = Number.isFinite(sampleSize) ? sampleSize : 0;
    const normalizedSampleSize = Math.min(noteIds.length, Math.max(0, Math.floor(finiteSampleSize)));
    if (normalizedSampleSize === 0) {
      return [];
    }

    return this.notesInfo(noteIds.slice(0, normalizedSampleSize));
  }

  async fieldNamesForDeck(deckName: string, sampleSize = 100): Promise<string[]> {
    const noteInfos = await this.noteInfosForDeck(deckName, sampleSize);
    const fields = new Set<string>();
    for (const noteInfo of noteInfos) {
      const noteFields = noteInfo.fields;
      if (!noteFields || typeof noteFields !== 'object' || Array.isArray(noteFields)) {
        continue;
      }
      for (const fieldName of Object.keys(noteFields)) {
        fields.add(fieldName);
      }
    }
    return [...fields].sort();
  }

  async modelNamesForDeck(deckName: string, sampleSize = 100): Promise<string[]> {
    const noteInfos = await this.noteInfosForDeck(deckName, sampleSize);
    const modelNames = new Set<string>();
    for (const noteInfo of noteInfos) {
      const modelName = noteInfo.modelName;
      if (typeof modelName === 'string' && modelName.length > 0) {
        modelNames.add(modelName);
      }
    }
    return [...modelNames].sort();
  }

  async notesInfo(noteIds: number[]): Promise<Record<string, unknown>[]> {
    const result = await this.invoke('notesInfo', { notes: noteIds });
    return (result as Record<string, unknown>[]) || [];
  }

  async updateNoteFields(noteId: number, fields: Record<string, string>): Promise<void> {
    await this.invoke('updateNoteFields', {
      note: {
        id: noteId,
        fields,
      },
    });
  }

  async storeMediaFile(filename: string, data: Buffer): Promise<void> {
    const base64Data = data.toString('base64');
    const sizeKB = Math.round(base64Data.length / 1024);
    log.info(`Uploading media file: ${filename} (${sizeKB}KB)`);

    await this.invoke(
      'storeMediaFile',
      {
        filename,
        data: base64Data,
      },
      { timeout: 30000 },
    );
  }

  async addNote(
    deckName: string,
    modelName: string,
    fields: Record<string, string>,
    tags: string[] = [],
  ): Promise<number> {
    const note: {
      deckName: string;
      modelName: string;
      fields: Record<string, string>;
      tags?: string[];
    } = { deckName, modelName, fields };
    if (tags.length > 0) {
      note.tags = tags;
    }

    const result = await this.invoke('addNote', {
      note,
    });
    return result as number;
  }

  async addTags(noteIds: number[], tags: string[]): Promise<void> {
    if (noteIds.length === 0 || tags.length === 0) {
      return;
    }

    await this.invoke('addTags', {
      notes: noteIds,
      tags: tags.join(' '),
    });
  }

  async deleteNotes(noteIds: number[]): Promise<void> {
    await this.invoke('deleteNotes', { notes: noteIds });
  }

  async retrieveMediaFile(filename: string): Promise<string> {
    const result = await this.invoke('retrieveMediaFile', { filename });
    return (result as string) || '';
  }

  resetBackoff(): void {
    this.backoffMs = 200;
    this.consecutiveFailures = 0;
  }
}
