/*
 * SubMiner - All-in-one sentence mining overlay
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

import * as childProcess from 'child_process';
import { PartOfSpeech, Token, MecabStatus } from './types';
import { createLogger } from './logger';

export { PartOfSpeech };

const log = createLogger('mecab');

function mapPartOfSpeech(pos1: string): PartOfSpeech {
  switch (pos1) {
    case '名詞':
      return PartOfSpeech.noun;
    case '動詞':
      return PartOfSpeech.verb;
    case '形容詞':
      return PartOfSpeech.i_adjective;
    case '形状詞':
    case '形容動詞':
      return PartOfSpeech.na_adjective;
    case '助詞':
      return PartOfSpeech.particle;
    case '助動詞':
      return PartOfSpeech.bound_auxiliary;
    case '記号':
    case '補助記号':
      return PartOfSpeech.symbol;
    default:
      return PartOfSpeech.other;
  }
}

export function parseMecabLine(line: string): Token | null {
  if (!line || line === 'EOS' || line.trim() === '') {
    return null;
  }

  const tabIndex = line.indexOf('\t');
  if (tabIndex === -1) {
    return null;
  }

  const surface = line.substring(0, tabIndex);
  const featureString = line.substring(tabIndex + 1);
  const features = featureString.split(',');

  const pos1 = features[0] || '';
  const pos2 = features[1] || '';
  const pos3 = features[2] || '';
  const pos4 = features[3] || '';
  const inflectionType = features[4] || '';
  const inflectionForm = features[5] || '';
  const lemma = features[6] || surface;
  const reading = features[7] || '';
  const pronunciation = features[8] || '';

  return {
    word: surface,
    partOfSpeech: mapPartOfSpeech(pos1),
    pos1,
    pos2,
    pos3,
    pos4,
    inflectionType,
    inflectionForm,
    headword: lemma !== '*' ? lemma : surface,
    katakanaReading: reading !== '*' ? reading : '',
    pronunciation: pronunciation !== '*' ? pronunciation : '',
  };
}

export interface MecabTokenizerOptions {
  mecabCommand?: string;
  dictionaryPath?: string;
  idleShutdownMs?: number;
  spawnFn?: typeof childProcess.spawn;
  execSyncFn?: typeof childProcess.execSync;
  setTimeoutFn?: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  clearTimeoutFn?: (timer: ReturnType<typeof setTimeout>) => void;
}

interface MecabQueuedRequest {
  text: string;
  retryCount: number;
  resolve: (tokens: Token[] | null) => void;
}

interface MecabActiveRequest extends MecabQueuedRequest {
  lines: string[];
  stderr: string;
}

export class MecabTokenizer {
  private static readonly DEFAULT_IDLE_SHUTDOWN_MS = 30_000;
  private static readonly MAX_RETRY_COUNT = 1;

  private mecabPath: string | null = null;
  private mecabCommand: string;
  private dictionaryPath: string | null;
  private available: boolean = false;
  private enabled: boolean = true;
  private idleShutdownMs: number;
  private readonly spawnFn: typeof childProcess.spawn;
  private readonly execSyncFn: typeof childProcess.execSync;
  private readonly setTimeoutFn: (
    callback: () => void,
    delayMs: number,
  ) => ReturnType<typeof setTimeout>;
  private readonly clearTimeoutFn: (timer: ReturnType<typeof setTimeout>) => void;
  private mecabProcess: ReturnType<typeof childProcess.spawn> | null = null;
  private idleShutdownTimer: ReturnType<typeof setTimeout> | null = null;
  private stdoutBuffer = '';
  private requestQueue: MecabQueuedRequest[] = [];
  private activeRequest: MecabActiveRequest | null = null;

  constructor(options: MecabTokenizerOptions = {}) {
    this.mecabCommand = options.mecabCommand?.trim() || 'mecab';
    this.dictionaryPath = options.dictionaryPath?.trim() || null;
    this.idleShutdownMs = Math.max(
      0,
      Math.floor(options.idleShutdownMs ?? MecabTokenizer.DEFAULT_IDLE_SHUTDOWN_MS),
    );
    this.spawnFn = options.spawnFn ?? childProcess.spawn;
    this.execSyncFn = options.execSyncFn ?? childProcess.execSync;
    this.setTimeoutFn =
      options.setTimeoutFn ?? ((callback, delayMs) => setTimeout(callback, delayMs));
    this.clearTimeoutFn = options.clearTimeoutFn ?? ((timer) => clearTimeout(timer));
  }

  async checkAvailability(): Promise<boolean> {
    try {
      const command = this.mecabCommand;
      const result = command.includes('/')
        ? command
        : this.execSyncFn(`which ${command}`, { encoding: 'utf-8' });
      const resolvedPath = String(result).trim();
      if (resolvedPath) {
        this.mecabPath = resolvedPath;
        this.available = true;
        log.info('MeCab found at:', this.mecabPath);
        return true;
      }
    } catch (err) {
      log.info('MeCab not found on system');
    }

    this.stopPersistentProcess();
    this.available = false;
    return false;
  }

  async tokenize(text: string): Promise<Token[] | null> {
    const normalizedText = text.replace(/\r?\n/g, ' ').trim();
    if (!this.available || !this.enabled || !normalizedText) {
      return null;
    }

    return new Promise((resolve) => {
      this.clearIdleShutdownTimer();
      this.requestQueue.push({
        text: normalizedText,
        retryCount: 0,
        resolve,
      });
      this.processQueue();
    });
  }

  private processQueue(): void {
    if (this.activeRequest) {
      return;
    }

    const request = this.requestQueue.shift();
    if (!request) {
      this.scheduleIdleShutdown();
      return;
    }

    if (!this.ensurePersistentProcess()) {
      this.retryOrResolveRequest(request);
      this.processQueue();
      return;
    }

    this.activeRequest = {
      ...request,
      lines: [],
      stderr: '',
    };

    try {
      this.mecabProcess?.stdin?.write(`${request.text}\n`);
    } catch (error) {
      log.error('Failed to write to MeCab process:', (error as Error).message);
      this.retryOrResolveRequest(request);
      this.activeRequest = null;
      this.stopPersistentProcess();
      this.processQueue();
    }
  }

  private retryOrResolveRequest(request: MecabQueuedRequest): void {
    if (request.retryCount < MecabTokenizer.MAX_RETRY_COUNT && this.enabled && this.available) {
      this.requestQueue.push({
        ...request,
        retryCount: request.retryCount + 1,
      });
      return;
    }
    request.resolve(null);
  }

  private ensurePersistentProcess(): boolean {
    if (this.mecabProcess) {
      return true;
    }

    const mecabArgs: string[] = [];
    if (this.dictionaryPath) {
      mecabArgs.push('-d', this.dictionaryPath);
    }

    let mecab: ReturnType<typeof childProcess.spawn>;
    try {
      mecab = this.spawnFn(this.mecabPath ?? this.mecabCommand, mecabArgs, {
        stdio: ['pipe', 'pipe', 'pipe'],
      });
    } catch (error) {
      log.error('Failed to spawn MeCab:', (error as Error).message);
      return false;
    }

    if (!mecab.stdin || !mecab.stdout || !mecab.stderr) {
      log.error('Failed to spawn MeCab: missing stdio pipes');
      try {
        mecab.kill();
      } catch {}
      return false;
    }

    this.stdoutBuffer = '';
    mecab.stdout.on('data', (data: Buffer | string) => {
      this.handleStdoutChunk(data.toString());
    });
    mecab.stderr.on('data', (data: Buffer | string) => {
      if (!this.activeRequest) {
        return;
      }
      this.activeRequest.stderr += data.toString();
    });
    mecab.on('error', (error: Error) => {
      this.handlePersistentProcessEnded(mecab, `spawn error: ${error.message}`);
    });
    mecab.on('close', (code: number | null) => {
      this.handlePersistentProcessEnded(mecab, `exit code ${String(code)}`);
    });

    this.mecabProcess = mecab;
    return true;
  }

  private handleStdoutChunk(chunk: string): void {
    this.stdoutBuffer += chunk;
    while (true) {
      const newlineIndex = this.stdoutBuffer.indexOf('\n');
      if (newlineIndex === -1) {
        break;
      }
      const line = this.stdoutBuffer.slice(0, newlineIndex).replace(/\r$/, '');
      this.stdoutBuffer = this.stdoutBuffer.slice(newlineIndex + 1);
      this.handleStdoutLine(line);
    }
  }

  private handleStdoutLine(line: string): void {
    if (!this.activeRequest) {
      return;
    }
    if (line === 'EOS') {
      this.resolveActiveRequest();
      return;
    }
    if (!line.trim()) {
      return;
    }
    this.activeRequest.lines.push(line);
  }

  private resolveActiveRequest(): void {
    const current = this.activeRequest;
    if (!current) {
      return;
    }
    this.activeRequest = null;

    const tokens: Token[] = [];
    for (const line of current.lines) {
      const token = parseMecabLine(line);
      if (token) {
        tokens.push(token);
      }
    }

    if (tokens.length === 0 && current.text.trim().length > 0) {
      const trimmedStdout = current.lines.join('\n').trim();
      const trimmedStderr = current.stderr.trim();
      if (trimmedStdout) {
        log.warn(
          'MeCab returned no parseable tokens.',
          `command=${this.mecabPath ?? this.mecabCommand}`,
          `stdout=${trimmedStdout.slice(0, 1024)}`,
        );
      }
      if (trimmedStderr) {
        log.warn('MeCab stderr while tokenizing:', trimmedStderr);
      }
    }

    current.resolve(tokens);
    this.processQueue();
  }

  private handlePersistentProcessEnded(
    process: ReturnType<typeof childProcess.spawn>,
    reason: string,
  ): void {
    if (this.mecabProcess !== process) {
      return;
    }

    this.mecabProcess = null;
    this.stdoutBuffer = '';
    this.clearIdleShutdownTimer();

    const pending: MecabQueuedRequest[] = [];
    if (this.activeRequest) {
      pending.push({
        text: this.activeRequest.text,
        retryCount: this.activeRequest.retryCount,
        resolve: this.activeRequest.resolve,
      });
    }
    this.activeRequest = null;
    if (this.requestQueue.length > 0) {
      pending.push(...this.requestQueue);
    }
    this.requestQueue = [];

    if (pending.length > 0) {
      log.warn(
        `MeCab parser process ended during active work (${reason}); retrying pending request(s).`,
      );
      for (const request of pending) {
        this.retryOrResolveRequest(request);
      }
      this.processQueue();
    }
  }

  private scheduleIdleShutdown(): void {
    this.clearIdleShutdownTimer();
    if (this.idleShutdownMs <= 0 || !this.mecabProcess) {
      return;
    }
    this.idleShutdownTimer = this.setTimeoutFn(() => {
      this.idleShutdownTimer = null;
      if (this.activeRequest || this.requestQueue.length > 0) {
        return;
      }
      this.stopPersistentProcess();
    }, this.idleShutdownMs);
    const timerWithUnref = this.idleShutdownTimer as { unref?: () => void };
    if (typeof timerWithUnref.unref === 'function') {
      timerWithUnref.unref();
    }
  }

  private clearIdleShutdownTimer(): void {
    if (!this.idleShutdownTimer) {
      return;
    }
    this.clearTimeoutFn(this.idleShutdownTimer);
    this.idleShutdownTimer = null;
  }

  private stopPersistentProcess(): void {
    const process = this.mecabProcess;
    if (!process) {
      return;
    }
    this.mecabProcess = null;
    this.stdoutBuffer = '';
    this.clearIdleShutdownTimer();
    try {
      process.kill();
    } catch {}
  }

  getStatus(): MecabStatus {
    return {
      available: this.available,
      enabled: this.enabled,
      path: this.mecabPath,
    };
  }

  setEnabled(enabled: boolean): void {
    this.enabled = enabled;
    if (!enabled) {
      const pending: MecabQueuedRequest[] = [];
      if (this.activeRequest) {
        pending.push({
          text: this.activeRequest.text,
          retryCount: MecabTokenizer.MAX_RETRY_COUNT,
          resolve: this.activeRequest.resolve,
        });
      }
      if (this.requestQueue.length > 0) {
        pending.push(...this.requestQueue);
      }
      this.activeRequest = null;
      this.requestQueue = [];
      for (const request of pending) {
        request.resolve(null);
      }
      this.stopPersistentProcess();
    }
  }
}

export { mapPartOfSpeech };
