import type { SessionSummary } from '../types/stats';

export interface SessionBucket {
  key: string;
  videoId: number | null;
  sessions: SessionSummary[];
  totalActiveMs: number;
  totalCardsMined: number;
  representativeSession: SessionSummary;
}

export function groupSessionsByVideo(sessions: SessionSummary[]): SessionBucket[] {
  const byKey = new Map<string, SessionSummary[]>();
  for (const session of sessions) {
    const hasVideoId =
      typeof session.videoId === 'number' &&
      Number.isFinite(session.videoId) &&
      session.videoId > 0;
    const key = hasVideoId ? `v-${session.videoId}` : `s-${session.sessionId}`;
    const existing = byKey.get(key);
    if (existing) existing.push(session);
    else byKey.set(key, [session]);
  }

  const buckets: SessionBucket[] = [];
  for (const [key, group] of byKey) {
    const sorted = [...group].sort((a, b) => b.startedAtMs - a.startedAtMs);
    const representative = sorted[0]!;
    buckets.push({
      key,
      videoId:
        typeof representative.videoId === 'number' && representative.videoId > 0
          ? representative.videoId
          : null,
      sessions: sorted,
      totalActiveMs: sorted.reduce((s, x) => s + x.activeWatchedMs, 0),
      totalCardsMined: sorted.reduce((s, x) => s + x.cardsMined, 0),
      representativeSession: representative,
    });
  }

  return buckets;
}
