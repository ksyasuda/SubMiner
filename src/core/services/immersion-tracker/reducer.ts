import path from 'node:path';
import type { ProbeMetadata, SessionState } from './types';
import { SOURCE_TYPE_REMOTE } from './types';

export function createInitialSessionState(
  sessionId: number,
  videoId: number,
  startedAtMs: number,
): SessionState {
  return {
    sessionId,
    videoId,
    startedAtMs,
    currentLineIndex: 0,
    totalWatchedMs: 0,
    activeWatchedMs: 0,
    linesSeen: 0,
    wordsSeen: 0,
    tokensSeen: 0,
    cardsMined: 0,
    lookupCount: 0,
    lookupHits: 0,
    pauseCount: 0,
    pauseMs: 0,
    seekForwardCount: 0,
    seekBackwardCount: 0,
    mediaBufferEvents: 0,
    lastWallClockMs: 0,
    lastMediaMs: null,
    lastPauseStartMs: null,
    isPaused: false,
    pendingTelemetry: true,
  };
}

export function resolveBoundedInt(
  value: number | undefined,
  fallback: number,
  min: number,
  max: number,
): number {
  if (!Number.isFinite(value)) return fallback;
  const candidate = Math.floor(value as number);
  if (candidate < min || candidate > max) return fallback;
  return candidate;
}

export function sanitizePayload(payload: Record<string, unknown>, maxPayloadBytes: number): string {
  const json = JSON.stringify(payload);
  return json.length <= maxPayloadBytes ? json : JSON.stringify({ truncated: true });
}

export function calculateTextMetrics(value: string): {
  words: number;
  tokens: number;
} {
  const words = value.split(/\s+/).filter(Boolean).length;
  const cjkCount = value.match(/[\u3040-\u30ff\u4e00-\u9fff]/g)?.length ?? 0;
  const tokens = Math.max(words, cjkCount);
  return { words, tokens };
}

export function secToMs(seconds: number): number {
  const coerced = Number(seconds);
  if (!Number.isFinite(coerced)) return 0;
  return Math.round(coerced * 1000);
}

export function normalizeMediaPath(mediaPath: string | null): string {
  if (!mediaPath || !mediaPath.trim()) return '';
  return mediaPath.trim();
}

export function normalizeText(value: string | null | undefined): string {
  if (!value) return '';
  return value.trim().replace(/\s+/g, ' ');
}

export function buildVideoKey(mediaPath: string, sourceType: number): string {
  if (sourceType === SOURCE_TYPE_REMOTE) {
    return `remote:${mediaPath}`;
  }
  return `local:${mediaPath}`;
}

export function isRemoteSource(mediaPath: string): boolean {
  return /^[a-z][a-z0-9+.-]*:\/\//i.test(mediaPath);
}

export function deriveCanonicalTitle(mediaPath: string): string {
  if (isRemoteSource(mediaPath)) {
    try {
      const parsed = new URL(mediaPath);
      const parts = parsed.pathname.split('/').filter(Boolean);
      if (parts.length > 0) {
        const leaf = decodeURIComponent(parts[parts.length - 1]!);
        return normalizeText(leaf.replace(/\.[^/.]+$/, ''));
      }
      return normalizeText(parsed.hostname) || 'unknown';
    } catch {
      return normalizeText(mediaPath);
    }
  }

  const filename = path.basename(mediaPath);
  return normalizeText(filename.replace(/\.[^/.]+$/, ''));
}

export function parseFps(value?: string): number | null {
  if (!value || typeof value !== 'string') return null;
  const [num, den] = value.split('/');
  const n = Number(num);
  const d = Number(den);
  if (!Number.isFinite(n) || !Number.isFinite(d) || d === 0) return null;
  const fps = n / d;
  return Number.isFinite(fps) ? Math.round(fps * 100) : null;
}

export function hashToCode(input?: string): number | null {
  if (!input) return null;
  let hash = 0;
  for (let i = 0; i < input.length; i += 1) {
    hash = (hash * 31 + input.charCodeAt(i)) & 0x7fffffff;
  }
  return hash || null;
}

export function emptyMetadata(): ProbeMetadata {
  return {
    durationMs: null,
    codecId: null,
    containerId: null,
    widthPx: null,
    heightPx: null,
    fpsX100: null,
    bitrateKbps: null,
    audioCodecId: null,
  };
}

export function toNullableInt(value: number | null | undefined): number | null {
  if (value === null || value === undefined || !Number.isFinite(value)) return null;
  return value;
}
