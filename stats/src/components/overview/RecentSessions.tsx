import { useState } from 'react';
import {
  formatDuration,
  formatRelativeDate,
  formatNumber,
  formatSessionDayLabel,
} from '../../lib/formatters';
import { CoverThumbnail } from './CoverThumbnail';
import { useCoverImages } from '../../hooks/useCoverImages';
import type { CoverImageMap } from '../../lib/cover-images';
import { getSessionDisplayWordCount } from '../../lib/session-word-count';
import { getSessionNavigationTarget } from '../../lib/stats-navigation';
import type { SessionSummary } from '../../types/stats';

interface RecentSessionsProps {
  sessions: SessionSummary[];
  onNavigateToMediaDetail: (videoId: number, sessionId?: number | null) => void;
  onNavigateToSession: (sessionId: number) => void;
  onDeleteSession: (session: SessionSummary) => void;
  onDeleteDayGroup: (dayLabel: string, daySessions: SessionSummary[]) => void;
  onDeleteAnimeGroup: (sessions: SessionSummary[]) => void;
  deletingIds: Set<number>;
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
  totalKnownWords: number;
}

function groupSessionsByDay(sessions: SessionSummary[]): Map<string, SessionSummary[]> {
  const groups = new Map<string, SessionSummary[]>();

  for (const session of sessions) {
    const dayLabel = formatSessionDayLabel(session.startedAtMs);
    const group = groups.get(dayLabel);
    if (group) {
      group.push(session);
    } else {
      groups.set(dayLabel, [session]);
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
    const displayWordCount = getSessionDisplayWordCount(session);
    if (existing) {
      existing.sessions.push(session);
      existing.totalCards += session.cardsMined;
      existing.totalWords += displayWordCount;
      existing.totalActiveMs += session.activeWatchedMs;
      existing.totalKnownWords += session.knownWordsSeen;
    } else {
      map.set(key, {
        key,
        animeId: session.animeId,
        animeTitle: session.animeTitle,
        videoId: session.videoId,
        sessions: [session],
        totalCards: session.cardsMined,
        totalWords: displayWordCount,
        totalActiveMs: session.activeWatchedMs,
        totalKnownWords: session.knownWordsSeen,
      });
    }
  }

  return Array.from(map.values());
}

function SessionItem({
  session,
  onNavigateToMediaDetail,
  onNavigateToSession,
  onDelete,
  deleteDisabled,
  coverImages,
}: {
  session: SessionSummary;
  onNavigateToMediaDetail: (videoId: number, sessionId?: number | null) => void;
  onNavigateToSession: (sessionId: number) => void;
  onDelete: () => void;
  deleteDisabled: boolean;
  coverImages: CoverImageMap;
}) {
  const displayWordCount = getSessionDisplayWordCount(session);
  const navigationTarget = getSessionNavigationTarget(session);

  return (
    <div className="relative group">
      <button
        type="button"
        onClick={() => {
          if (navigationTarget.type === 'media-detail') {
            onNavigateToMediaDetail(navigationTarget.videoId, navigationTarget.sessionId);
            return;
          }
          onNavigateToSession(navigationTarget.sessionId);
        }}
        className="w-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-3 pr-10 flex items-center gap-3 hover:border-ctp-surface2 transition-colors text-left cursor-pointer"
      >
        <CoverThumbnail
          animeId={session.animeId}
          videoId={session.videoId}
          title={session.canonicalTitle ?? 'Unknown'}
          coverImages={coverImages}
        />
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
            <div className="text-ctp-cards-mined font-medium font-mono tabular-nums">
              {formatNumber(session.cardsMined)}
            </div>
            <div className="text-ctp-overlay2">cards</div>
          </div>
          <div>
            <div className="text-ctp-mauve font-medium font-mono tabular-nums">
              {formatNumber(displayWordCount)}
            </div>
            <div className="text-ctp-overlay2">words</div>
          </div>
          <div>
            <div className="text-ctp-green font-medium font-mono tabular-nums">
              {formatNumber(session.knownWordsSeen)}
            </div>
            <div className="text-ctp-overlay2">known words</div>
          </div>
        </div>
      </button>
      <button
        type="button"
        onClick={onDelete}
        disabled={deleteDisabled}
        aria-label={`Delete session ${session.canonicalTitle ?? 'Unknown Media'}`}
        className="absolute right-2 top-1/2 -translate-y-1/2 w-5 h-5 rounded border border-ctp-surface2 text-transparent hover:border-ctp-red/50 hover:text-ctp-red hover:bg-ctp-red/10 transition-colors opacity-0 group-hover:opacity-100 focus:opacity-100 flex items-center justify-center disabled:opacity-40 disabled:cursor-not-allowed"
        title="Delete session"
      >
        {'\u2715'}
      </button>
    </div>
  );
}

function AnimeGroupRow({
  group,
  onNavigateToMediaDetail,
  onNavigateToSession,
  onDeleteSession,
  onDeleteAnimeGroup,
  deletingIds,
  coverImages,
}: {
  group: AnimeGroup;
  onNavigateToMediaDetail: (videoId: number, sessionId?: number | null) => void;
  onNavigateToSession: (sessionId: number) => void;
  onDeleteSession: (session: SessionSummary) => void;
  onDeleteAnimeGroup: (group: AnimeGroup) => void;
  deletingIds: Set<number>;
  coverImages: CoverImageMap;
}) {
  const [expanded, setExpanded] = useState(false);
  const groupDeleting = group.sessions.some((s) => deletingIds.has(s.sessionId));

  if (group.sessions.length === 1) {
    const s = group.sessions[0]!;
    return (
      <SessionItem
        session={s}
        onNavigateToMediaDetail={onNavigateToMediaDetail}
        onNavigateToSession={onNavigateToSession}
        onDelete={() => onDeleteSession(s)}
        deleteDisabled={deletingIds.has(s.sessionId)}
        coverImages={coverImages}
      />
    );
  }

  const displayTitle = group.animeTitle ?? group.sessions[0]?.canonicalTitle ?? 'Unknown Media';
  const mostRecentSession = group.sessions[0]!;
  const disclosureId = `recent-sessions-${mostRecentSession.sessionId}`;

  return (
    <div className="group/anime">
      <div className="relative">
        <button
          type="button"
          onClick={() => setExpanded((value) => !value)}
          aria-expanded={expanded}
          aria-controls={disclosureId}
          className="w-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-3 pr-10 flex items-center gap-3 hover:border-ctp-surface2 transition-colors text-left"
        >
          <CoverThumbnail
            animeId={group.animeId}
            videoId={mostRecentSession.videoId}
            title={displayTitle}
            coverImages={coverImages}
          />
          <div className="min-w-0 flex-1">
            <div className="text-sm font-medium text-ctp-text truncate">{displayTitle}</div>
            <div className="text-xs text-ctp-overlay2">
              {group.sessions.length} sessions · {formatDuration(group.totalActiveMs)} active
            </div>
          </div>
          <div className="flex gap-4 text-xs text-center shrink-0">
            <div>
              <div className="text-ctp-cards-mined font-medium font-mono tabular-nums">
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
            <div>
              <div className="text-ctp-green font-medium font-mono tabular-nums">
                {formatNumber(group.totalKnownWords)}
              </div>
              <div className="text-ctp-overlay2">known words</div>
            </div>
          </div>
          <div
            className={`text-ctp-overlay2 text-xs transition-transform ${expanded ? 'rotate-90' : ''}`}
            aria-hidden="true"
          >
            {'\u25B8'}
          </div>
        </button>
        <button
          type="button"
          onClick={() => onDeleteAnimeGroup(group)}
          disabled={groupDeleting}
          aria-label={`Delete all sessions for ${displayTitle}`}
          className="absolute right-2 top-1/2 -translate-y-1/2 w-5 h-5 rounded border border-ctp-surface2 text-transparent hover:border-ctp-red/50 hover:text-ctp-red hover:bg-ctp-red/10 transition-colors opacity-0 group-hover/anime:opacity-100 focus:opacity-100 flex items-center justify-center disabled:opacity-40 disabled:cursor-not-allowed"
          title={`Delete all sessions for ${displayTitle}`}
        >
          {groupDeleting ? '\u2026' : '\u2715'}
        </button>
      </div>
      {expanded && (
        <div
          id={disclosureId}
          role="region"
          aria-label={`${displayTitle} sessions`}
          className="ml-6 mt-1 space-y-1"
        >
          {group.sessions.map((s) => {
            const navigationTarget = getSessionNavigationTarget(s);

            return (
              <div key={s.sessionId} className="relative group/nested">
                <button
                  type="button"
                  onClick={() => {
                    if (navigationTarget.type === 'media-detail') {
                      onNavigateToMediaDetail(navigationTarget.videoId, navigationTarget.sessionId);
                      return;
                    }
                    onNavigateToSession(navigationTarget.sessionId);
                  }}
                  className="w-full bg-ctp-mantle border border-ctp-surface0 rounded-lg p-2.5 pr-10 flex items-center gap-3 hover:border-ctp-surface1 transition-colors text-left cursor-pointer"
                >
                  <CoverThumbnail
                    animeId={s.animeId}
                    videoId={s.videoId}
                    title={s.canonicalTitle ?? 'Unknown'}
                    coverImages={coverImages}
                  />
                  <div className="min-w-0 flex-1">
                    <div className="text-sm font-medium text-ctp-subtext1 truncate">
                      {s.canonicalTitle ?? 'Unknown Media'}
                    </div>
                    <div className="text-xs text-ctp-overlay2">
                      {formatRelativeDate(s.startedAtMs)} · {formatDuration(s.activeWatchedMs)}{' '}
                      active
                    </div>
                  </div>
                  <div className="flex gap-4 text-xs text-center shrink-0">
                    <div>
                      <div className="text-ctp-cards-mined font-medium font-mono tabular-nums">
                        {formatNumber(s.cardsMined)}
                      </div>
                      <div className="text-ctp-overlay2">cards</div>
                    </div>
                    <div>
                      <div className="text-ctp-mauve font-medium font-mono tabular-nums">
                        {formatNumber(getSessionDisplayWordCount(s))}
                      </div>
                      <div className="text-ctp-overlay2">words</div>
                    </div>
                    <div>
                      <div className="text-ctp-green font-medium font-mono tabular-nums">
                        {formatNumber(s.knownWordsSeen)}
                      </div>
                      <div className="text-ctp-overlay2">known words</div>
                    </div>
                  </div>
                </button>
                <button
                  type="button"
                  onClick={() => onDeleteSession(s)}
                  disabled={deletingIds.has(s.sessionId)}
                  aria-label={`Delete session ${s.canonicalTitle ?? 'Unknown Media'}`}
                  className="absolute right-2 top-1/2 -translate-y-1/2 w-5 h-5 rounded border border-ctp-surface2 text-transparent hover:border-ctp-red/50 hover:text-ctp-red hover:bg-ctp-red/10 transition-colors opacity-0 group-hover/nested:opacity-100 focus:opacity-100 flex items-center justify-center disabled:opacity-40 disabled:cursor-not-allowed"
                  title="Delete session"
                >
                  {'\u2715'}
                </button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

export function RecentSessions({
  sessions,
  onNavigateToMediaDetail,
  onNavigateToSession,
  onDeleteSession,
  onDeleteDayGroup,
  onDeleteAnimeGroup,
  deletingIds,
}: RecentSessionsProps) {
  const coverImages = useCoverImages(sessions);

  if (sessions.length === 0) {
    return (
      <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
        <div className="text-sm text-ctp-overlay2">No sessions yet</div>
      </div>
    );
  }

  const groups = groupSessionsByDay(sessions);
  const anyDeleting = deletingIds.size > 0;

  return (
    <div className="space-y-4">
      {Array.from(groups.entries()).map(([dayLabel, daySessions]) => {
        const animeGroups = groupSessionsByAnime(daySessions);
        const groupDeleting = daySessions.some((s) => deletingIds.has(s.sessionId));
        return (
          <div key={dayLabel} className="group/day">
            <div className="flex items-center gap-3 mb-2">
              <h3 className="text-xs font-semibold text-ctp-overlay2 uppercase tracking-widest shrink-0">
                {dayLabel}
              </h3>
              <div className="flex-1 h-px bg-gradient-to-r from-ctp-surface1 to-transparent" />
              <button
                type="button"
                onClick={() => onDeleteDayGroup(dayLabel, daySessions)}
                disabled={anyDeleting}
                aria-label={`Delete all sessions from ${dayLabel}`}
                className="shrink-0 text-xs text-transparent hover:text-ctp-red transition-colors opacity-0 group-hover/day:opacity-100 focus:opacity-100 disabled:opacity-40 disabled:cursor-not-allowed"
                title={`Delete all sessions from ${dayLabel}`}
              >
                {groupDeleting ? '\u2026' : '\u2715'}
              </button>
            </div>
            <div className="space-y-2">
              {animeGroups.map((group) => (
                <AnimeGroupRow
                  key={group.key}
                  group={group}
                  onNavigateToMediaDetail={onNavigateToMediaDetail}
                  onNavigateToSession={onNavigateToSession}
                  onDeleteSession={onDeleteSession}
                  onDeleteAnimeGroup={(g) => onDeleteAnimeGroup(g.sessions)}
                  deletingIds={deletingIds}
                  coverImages={coverImages}
                />
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}
