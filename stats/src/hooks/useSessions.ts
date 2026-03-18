import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { SessionSummary, SessionTimelinePoint, SessionEvent } from '../types/stats';

export function useSessions(limit = 50) {
  const [sessions, setSessions] = useState<SessionSummary[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    const client = getStatsClient();
    client
      .getSessions(limit)
      .then((nextSessions) => {
        if (cancelled) return;
        setSessions(nextSessions);
      })
      .catch((err) => {
        if (cancelled) return;
        setError(err.message);
      })
      .finally(() => {
        if (cancelled) return;
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [limit]);

  return { sessions, loading, error };
}

export interface KnownWordsTimelinePoint {
  linesSeen: number;
  knownWordsSeen: number;
}

export function useSessionDetail(sessionId: number | null) {
  const [timeline, setTimeline] = useState<SessionTimelinePoint[]>([]);
  const [events, setEvents] = useState<SessionEvent[]>([]);
  const [knownWordsTimeline, setKnownWordsTimeline] = useState<KnownWordsTimelinePoint[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setError(null);
    if (sessionId == null) {
      setTimeline([]);
      setEvents([]);
      setKnownWordsTimeline([]);
      setLoading(false);
      return () => {
        cancelled = true;
      };
    }
    setLoading(true);
    setTimeline([]);
    setEvents([]);
    setKnownWordsTimeline([]);
    const client = getStatsClient();
    Promise.all([
      client.getSessionTimeline(sessionId),
      client.getSessionEvents(sessionId),
      client.getSessionKnownWordsTimeline(sessionId),
    ])
      .then(([nextTimeline, nextEvents, nextKnownWords]) => {
        if (cancelled) return;
        setTimeline(nextTimeline);
        setEvents(nextEvents);
        setKnownWordsTimeline(nextKnownWords);
      })
      .catch((err) => {
        if (cancelled) return;
        setError(err.message);
      })
      .finally(() => {
        if (cancelled) return;
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [sessionId]);

  return { timeline, events, knownWordsTimeline, loading, error };
}
