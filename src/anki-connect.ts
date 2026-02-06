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

import axios, { AxiosInstance } from "axios";
import http from "http";
import https from "https";

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
  private url: string;
  private backoffMs = 200;
  private maxBackoffMs = 5000;
  private consecutiveFailures = 0;
  private maxConsecutiveFailures = 5;

  constructor(url: string) {
    this.url = url;

    const httpAgent = new http.Agent({
      keepAlive: true,
      keepAliveMsecs: 1000,
      maxSockets: 5,
      maxFreeSockets: 2,
      timeout: 10000,
    });

    const httpsAgent = new https.Agent({
      keepAlive: true,
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
    if (!error || typeof error !== "object") return false;

    const code = (error as Record<string, unknown>).code;
    const message =
      typeof (error as Record<string, unknown>).message === "string"
        ? ((error as Record<string, unknown>).message as string).toLowerCase()
        : "";

    return (
      code === "ECONNRESET" ||
      code === "ETIMEDOUT" ||
      code === "ENOTFOUND" ||
      code === "ECONNREFUSED" ||
      code === "EPIPE" ||
      message.includes("socket hang up") ||
      message.includes("network error") ||
      message.includes("timeout")
    );
  }

  async invoke(
    action: string,
    params: Record<string, unknown> = {},
    options: { timeout?: number; maxRetries?: number } = {},
  ): Promise<unknown> {
    const maxRetries = options.maxRetries ?? 3;
    let lastError: Error | null = null;

    const isMediaUpload = action === "storeMediaFile";
    const requestTimeout = options.timeout || (isMediaUpload ? 30000 : 10000);

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        if (attempt > 0) {
          const delay = Math.min(
            this.backoffMs * Math.pow(2, attempt - 1),
            this.maxBackoffMs,
          );
          console.log(
            `AnkiConnect retry ${attempt}/${maxRetries} after ${delay}ms delay`,
          );
          await this.sleep(delay);
        }

        const response = await this.client.post<AnkiConnectResponse>(
          "",
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
            console.error(
              `AnkiConnect error (attempt ${this.consecutiveFailures}/${this.maxConsecutiveFailures}):`,
              lastError.message,
            );
          } else if (this.consecutiveFailures === this.maxConsecutiveFailures) {
            console.error(
              "AnkiConnect: Too many consecutive failures, suppressing further error logs",
            );
          }
          throw lastError;
        }
      }
    }

    throw lastError || new Error("Unknown error");
  }

  async findNotes(
    query: string,
    options?: { maxRetries?: number },
  ): Promise<number[]> {
    const result = await this.invoke("findNotes", { query }, options);
    return (result as number[]) || [];
  }

  async notesInfo(noteIds: number[]): Promise<Record<string, unknown>[]> {
    const result = await this.invoke("notesInfo", { notes: noteIds });
    return (result as Record<string, unknown>[]) || [];
  }

  async updateNoteFields(
    noteId: number,
    fields: Record<string, string>,
  ): Promise<void> {
    await this.invoke("updateNoteFields", {
      note: {
        id: noteId,
        fields,
      },
    });
  }

  async storeMediaFile(filename: string, data: Buffer): Promise<void> {
    const base64Data = data.toString("base64");
    const sizeKB = Math.round(base64Data.length / 1024);
    console.log(`Uploading media file: ${filename} (${sizeKB}KB)`);

    await this.invoke(
      "storeMediaFile",
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
  ): Promise<number> {
    const result = await this.invoke("addNote", {
      note: { deckName, modelName, fields },
    });
    return result as number;
  }

  async deleteNotes(noteIds: number[]): Promise<void> {
    await this.invoke("deleteNotes", { notes: noteIds });
  }

  async retrieveMediaFile(filename: string): Promise<string> {
    const result = await this.invoke("retrieveMediaFile", { filename });
    return (result as string) || "";
  }

  resetBackoff(): void {
    this.backoffMs = 200;
    this.consecutiveFailures = 0;
  }
}
