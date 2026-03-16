import { useState } from 'react';
import {
  formatDuration,
  formatRelativeDate,
  formatNumber,
  todayLocalDay,
  localDayFromMs,
} from '../../lib/formatters';
import { BASE_URL } from '../../lib/api-client';
import type { SessionSummary } from '../../types/stats';

interface RecentSessionsProps {
  sessions: SessionSummary[];
}

interface AnimeGroup {
  key: string;
  animeId: number | null;
  animeTitle: string | null;
  videoId: number | null;
  sessions: SessionSummary[];
  totalCards: number;
  totalWords: number;
  totalActiveMs: number;
}

function groupSessionsByDay(sessions: SessionSummary[]): Map<string, SessionSummary[]> {
  const groups = new Map<string, SessionSummary[]>();
  const today = todayLocalDay();

  for (const session of sessions) {
    const sessionDay = localDayFromMs(session.startedAtMs);
    let label: string;
    if (sessionDay === today) {
      label = 'Today';
    } else if (sessionDay === today - 1) {
      label = 'Yesterday';
    } else {
      label = new Date(session.startedAtMs).toLocaleDateString(undefined, {
        month: 'long',
        day: 'numeric',
      });
    }
    const group = groups.get(label);
    if (group) {
      group.push(session);
    } else {
      groups.set(label, [session]);
    }
  }

  return groups;
}

function groupSessionsByAnime(sessions: SessionSummary[]): AnimeGroup[] {
  const map = new Map<string, AnimeGroup>();

  for (const session of sessions) {
    const key =
      session.animeId != null
        ? `anime-${session.animeId}`
        : session.videoId != null
          ? `video-${session.videoId}`
          : `session-${session.sessionId}`;

    const existing = map.get(key);
    if (existing) {
      existing.sessions.push(session);
      existing.totalCards += session.cardsMined;
      existing.totalWords += session.wordsSeen;
      existing.totalActiveMs += session.activeWatchedMs;
    } else {
      map.set(key, {
        key,
        animeId: session.animeId,
        animeTitle: session.animeTitle,
        videoId: session.videoId,
        sessions: [session],
        totalCards: session.cardsMined,
        totalWords: session.wordsSeen,
        totalActiveMs: session.activeWatchedMs,
      });
    }
  }

  return Array.from(map.values());
}

function CoverThumbnail({ videoId, title }: { videoId: number | null; title: string }) {
  const fallbackChar = title.charAt(0) || '?';

  if (!videoId) {
    return (
      <div className="w-12 h-16 rounded bg-ctp-surface2 flex items-center justify-center text-ctp-overlay2 text-lg font-bold shrink-0">
        {fallbackChar}
      </div>
    );
  }

  return (
    <img
      src={`${BASE_URL}/api/stats/media/${videoId}/cover`}
      alt=""
      className="w-12 h-16 rounded object-cover shrink-0 bg-ctp-surface2"
      onError={(e) => {
        const target = e.currentTarget;
        target.style.display = 'none';
        const placeholder = document.createElement('div');
        placeholder.className =
          'w-12 h-16 rounded bg-ctp-surface2 flex items-center justify-center text-ctp-overlay2 text-lg font-bold shrink-0';
        placeholder.textContent = fallbackChar;
        target.parentElement?.insertBefore(placeholder, target);
      }}
    />
  );
}

function SessionItem({ session }: { session: SessionSummary }) {
  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-3 flex items-center gap-3">
      <CoverThumbnail videoId={session.videoId} title={session.canonicalTitle ?? 'Unknown'} />
      <div className="min-w-0 flex-1">
        <div className="text-sm font-medium text-ctp-text truncate">
          {session.canonicalTitle ?? 'Unknown Media'}
        </div>
        <div className="text-xs text-ctp-overlay2">
          {formatRelativeDate(session.startedAtMs)} · {formatDuration(session.activeWatchedMs)}{' '}
          active
        </div>
      </div>
      <div className="flex gap-4 text-xs text-center shrink-0">
        <div>
          <div className="text-ctp-green font-medium font-mono tabular-nums">
            {formatNumber(session.cardsMined)}
          </div>
          <div className="text-ctp-overlay2">cards</div>
        </div>
        <div>
          <div className="text-ctp-mauve font-medium font-mono tabular-nums">
            {formatNumber(session.wordsSeen)}
          </div>
          <div className="text-ctp-overlay2">words</div>
        </div>
      </div>
    </div>
  );
}

function AnimeGroupRow({ group }: { group: AnimeGroup }) {
  const [expanded, setExpanded] = useState(false);

  if (group.sessions.length === 1) {
    return <SessionItem session={group.sessions[0]!} />;
  }

  const displayTitle = group.animeTitle ?? group.sessions[0]?.canonicalTitle ?? 'Unknown Media';
  const mostRecentSession = group.sessions[0]!;

  return (
    <div>
      <button
        type="button"
        onClick={() => setExpanded(!expanded)}
        className="w-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-3 flex items-center gap-3 hover:border-ctp-surface2 transition-colors text-left"
      >
        <CoverThumbnail videoId={mostRecentSession.videoId} title={displayTitle} />
        <div className="min-w-0 flex-1">
          <div className="text-sm font-medium text-ctp-text truncate">{displayTitle}</div>
          <div className="text-xs text-ctp-overlay2">
            {group.sessions.length} sessions · {formatDuration(group.totalActiveMs)} active
          </div>
        </div>
        <div className="flex gap-4 text-xs text-center shrink-0">
          <div>
            <div className="text-ctp-green font-medium font-mono tabular-nums">
              {formatNumber(group.totalCards)}
            </div>
            <div className="text-ctp-overlay2">cards</div>
          </div>
          <div>
            <div className="text-ctp-mauve font-medium font-mono tabular-nums">
              {formatNumber(group.totalWords)}
            </div>
            <div className="text-ctp-overlay2">words</div>
          </div>
        </div>
        <div
          className={`text-ctp-overlay2 text-xs transition-transform ${expanded ? 'rotate-90' : ''}`}
        >
          {'\u25B8'}
        </div>
      </button>
      {expanded && (
        <div className="ml-6 mt-1 space-y-1">
          {group.sessions.map((s) => (
            <div
              key={s.sessionId}
              className="bg-ctp-mantle border border-ctp-surface0 rounded-lg p-2.5 flex items-center gap-3"
            >
              <CoverThumbnail videoId={s.videoId} title={s.canonicalTitle ?? 'Unknown'} />
              <div className="min-w-0 flex-1">
                <div className="text-sm font-medium text-ctp-subtext1 truncate">
                  {s.canonicalTitle ?? 'Unknown Media'}
                </div>
                <div className="text-xs text-ctp-overlay2">
                  {formatRelativeDate(s.startedAtMs)} · {formatDuration(s.activeWatchedMs)} active
                </div>
              </div>
              <div className="flex gap-4 text-xs text-center shrink-0">
                <div>
                  <div className="text-ctp-green font-medium font-mono tabular-nums">
                    {formatNumber(s.cardsMined)}
                  </div>
                  <div className="text-ctp-overlay2">cards</div>
                </div>
                <div>
                  <div className="text-ctp-mauve font-medium font-mono tabular-nums">
                    {formatNumber(s.wordsSeen)}
                  </div>
                  <div className="text-ctp-overlay2">words</div>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export function RecentSessions({ sessions }: RecentSessionsProps) {
  if (sessions.length === 0) {
    return (
      <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
        <div className="text-sm text-ctp-overlay2">No sessions yet</div>
      </div>
    );
  }

  const groups = groupSessionsByDay(sessions);

  return (
    <div className="space-y-4">
      {Array.from(groups.entries()).map(([dayLabel, daySessions]) => {
        const animeGroups = groupSessionsByAnime(daySessions);
        return (
          <div key={dayLabel}>
            <div className="flex items-center gap-3 mb-2">
              <h3 className="text-xs font-semibold text-ctp-overlay2 uppercase tracking-widest shrink-0">
                {dayLabel}
              </h3>
              <div className="flex-1 h-px bg-gradient-to-r from-ctp-surface1 to-transparent" />
            </div>
            <div className="space-y-2">
              {animeGroups.map((group) => (
                <AnimeGroupRow key={group.key} group={group} />
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}
