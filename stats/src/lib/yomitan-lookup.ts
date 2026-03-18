import type { SessionEvent } from '../types/stats';
import { EventType } from '../types/stats';

export interface LookupRateDisplay {
  shortValue: string;
  longValue: string;
}

export function buildLookupRateDisplay(
  yomitanLookupCount: number,
  tokensSeen: number,
): LookupRateDisplay | null {
  if (!Number.isFinite(yomitanLookupCount) || !Number.isFinite(tokensSeen) || tokensSeen <= 0) {
    return null;
  }
  const per100 = ((Math.max(0, yomitanLookupCount) / tokensSeen) * 100).toFixed(1);
  return {
    shortValue: `${per100} / 100 tokens`,
    longValue: `${per100} lookups per 100 tokens`,
  };
}

export function getYomitanLookupEvents(events: SessionEvent[]): SessionEvent[] {
  return events.filter((event) => event.eventType === EventType.YOMITAN_LOOKUP);
}
