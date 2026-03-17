import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import WebSocket from 'ws';
import { createLogger } from '../../logger';
import type { MergedToken, SubtitleData } from '../../types';

const logger = createLogger('main:subtitle-ws');

export function hasMpvWebsocketPlugin(): boolean {
  const mpvWebsocketPath = path.join(os.homedir(), '.config', 'mpv', 'mpv_websocket');
  return fs.existsSync(mpvWebsocketPath);
}

export type SubtitleWebsocketFrequencyOptions = {
  enabled: boolean;
  topX: number;
  mode: 'single' | 'banded';
};

type SerializedSubtitleToken = Pick<
  MergedToken,
  | 'surface'
  | 'reading'
  | 'headword'
  | 'startPos'
  | 'endPos'
  | 'partOfSpeech'
  | 'isMerged'
  | 'isKnown'
  | 'isNPlusOneTarget'
  | 'frequencyRank'
  | 'jlptLevel'
> & {
  isNameMatch: boolean;
  className: string;
  frequencyRankLabel: string | null;
  jlptLevelLabel: string | null;
};

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function hasPrioritizedNameMatch(token: MergedToken): boolean {
  return token.isNameMatch === true;
}

function computeFrequencyClass(
  token: MergedToken,
  options: SubtitleWebsocketFrequencyOptions,
): string | null {
  if (hasPrioritizedNameMatch(token)) return null;
  if (!options.enabled) return null;
  if (typeof token.frequencyRank !== 'number' || !Number.isFinite(token.frequencyRank)) return null;

  const rank = Math.max(1, Math.floor(token.frequencyRank));
  const topX = Math.max(1, Math.floor(options.topX));
  if (rank > topX) return null;

  if (options.mode === 'banded') {
    const band = Math.min(5, Math.max(1, Math.ceil((rank / topX) * 5)));
    return `word-frequency-band-${band}`;
  }

  return 'word-frequency-single';
}

function getFrequencyRankLabel(
  token: MergedToken,
  options: SubtitleWebsocketFrequencyOptions,
): string | null {
  if (hasPrioritizedNameMatch(token)) return null;
  if (!options.enabled) return null;
  if (typeof token.frequencyRank !== 'number' || !Number.isFinite(token.frequencyRank)) return null;

  const rank = Math.max(1, Math.floor(token.frequencyRank));
  const topX = Math.max(1, Math.floor(options.topX));
  return rank <= topX ? String(rank) : null;
}

function getJlptLevelLabel(token: MergedToken): string | null {
  if (hasPrioritizedNameMatch(token)) {
    return null;
  }

  return token.jlptLevel ?? null;
}

function computeWordClass(token: MergedToken, options: SubtitleWebsocketFrequencyOptions): string {
  const classes = ['word'];

  if (hasPrioritizedNameMatch(token)) {
    classes.push('word-name-match');
  } else if (token.isNPlusOneTarget) {
    classes.push('word-n-plus-one');
  } else if (token.isKnown) {
    classes.push('word-known');
  }

  if (!hasPrioritizedNameMatch(token) && token.jlptLevel) {
    classes.push(`word-jlpt-${token.jlptLevel.toLowerCase()}`);
  }

  if (!token.isKnown && !token.isNPlusOneTarget && !token.isNameMatch) {
    const frequencyClass = computeFrequencyClass(token, options);
    if (frequencyClass) {
      classes.push(frequencyClass);
    }
  }

  return classes.join(' ');
}

function serializeWordDataAttributes(
  token: MergedToken,
  options: SubtitleWebsocketFrequencyOptions,
): string {
  const attributes: string[] = [];

  if (token.reading) {
    attributes.push(`data-reading="${escapeHtml(token.reading)}"`);
  }
  if (token.headword) {
    attributes.push(`data-headword="${escapeHtml(token.headword)}"`);
  }

  const frequencyRankLabel = getFrequencyRankLabel(token, options);
  if (frequencyRankLabel) {
    attributes.push(`data-frequency-rank="${escapeHtml(frequencyRankLabel)}"`);
  }

  const jlptLevelLabel = getJlptLevelLabel(token);
  if (jlptLevelLabel) {
    attributes.push(`data-jlpt-level="${escapeHtml(jlptLevelLabel)}"`);
  }

  return attributes.length > 0 ? ` ${attributes.join(' ')}` : '';
}

function serializeSubtitleToken(
  token: MergedToken,
  options: SubtitleWebsocketFrequencyOptions,
): SerializedSubtitleToken {
  const prioritizedNameMatch = hasPrioritizedNameMatch(token);

  return {
    surface: token.surface,
    reading: token.reading,
    headword: token.headword,
    startPos: token.startPos,
    endPos: token.endPos,
    partOfSpeech: token.partOfSpeech,
    isMerged: token.isMerged,
    isKnown: token.isKnown,
    isNPlusOneTarget: prioritizedNameMatch ? false : token.isNPlusOneTarget,
    isNameMatch: token.isNameMatch ?? false,
    jlptLevel: prioritizedNameMatch ? undefined : token.jlptLevel,
    frequencyRank: prioritizedNameMatch ? undefined : token.frequencyRank,
    className: computeWordClass(token, options),
    frequencyRankLabel: getFrequencyRankLabel(token, options),
    jlptLevelLabel: getJlptLevelLabel(token),
  };
}

export function serializeSubtitleMarkup(
  payload: SubtitleData,
  options: SubtitleWebsocketFrequencyOptions,
): string {
  if (!payload.tokens || payload.tokens.length === 0) {
    return escapeHtml(payload.text).replaceAll('\n', '<br>');
  }

  const chunks: string[] = [];
  for (const token of payload.tokens) {
    const klass = computeWordClass(token, options);
    const attrs = serializeWordDataAttributes(token, options);
    const parts = token.surface.split('\n');
    for (let index = 0; index < parts.length; index += 1) {
      const part = parts[index];
      if (part) {
        chunks.push(`<span class="${klass}"${attrs}>${escapeHtml(part)}</span>`);
      }
      if (index < parts.length - 1) {
        chunks.push('<br>');
      }
    }
  }

  return chunks.join('');
}

export function serializeSubtitleWebsocketMessage(
  payload: SubtitleData,
  options: SubtitleWebsocketFrequencyOptions,
): string {
  return JSON.stringify({
    version: 1,
    text: payload.text,
    sentence: serializeSubtitleMarkup(payload, options),
    tokens: payload.tokens?.map((token) => serializeSubtitleToken(token, options)) ?? [],
  });
}

export function serializeInitialSubtitleWebsocketMessage(
  payload: SubtitleData | null,
  options: SubtitleWebsocketFrequencyOptions,
): string | null {
  if (!payload || !payload.text.trim()) {
    return null;
  }

  return serializeSubtitleWebsocketMessage(payload, options);
}

export class SubtitleWebSocket {
  private server: WebSocket.Server | null = null;
  private latestMessage = '';

  public isRunning(): boolean {
    return this.server !== null;
  }

  public hasClients(): boolean {
    return (this.server?.clients.size ?? 0) > 0;
  }

  public start(
    port: number,
    getCurrentSubtitleData: () => SubtitleData | null,
    getFrequencyOptions: () => SubtitleWebsocketFrequencyOptions,
  ): void {
    this.server = new WebSocket.Server({ port, host: '127.0.0.1' });

    this.server.on('connection', (ws: WebSocket) => {
      logger.info('WebSocket client connected');
      if (this.latestMessage) {
        ws.send(this.latestMessage);
        return;
      }

      const currentMessage = serializeInitialSubtitleWebsocketMessage(
        getCurrentSubtitleData(),
        getFrequencyOptions(),
      );
      if (currentMessage) {
        ws.send(currentMessage);
      }
    });

    this.server.on('error', (err: Error) => {
      logger.error('WebSocket server error:', err.message);
    });

    logger.info(`Subtitle WebSocket server running on ws://127.0.0.1:${port}`);
  }

  public broadcast(payload: SubtitleData, options: SubtitleWebsocketFrequencyOptions): void {
    if (!this.server) return;
    const message = serializeSubtitleWebsocketMessage(payload, options);
    this.latestMessage = message;
    for (const client of this.server.clients) {
      if (client.readyState === WebSocket.OPEN) {
        client.send(message);
      }
    }
  }

  public stop(): void {
    if (this.server) {
      this.server.close();
      this.server = null;
    }
    this.latestMessage = '';
  }
}
