import type { OverviewSummary } from '../../lib/dashboard-data';
import { formatNumber } from '../../lib/formatters';
import { Tooltip } from '../layout/Tooltip';

interface KnownWordsSummary {
  totalUniqueWords: number;
  knownWordCount: number;
}

interface TrackingSnapshotProps {
  summary: OverviewSummary;
  showTrackedCardNote?: boolean;
  knownWordsSummary: KnownWordsSummary | null;
}

export function TrackingSnapshot({
  summary,
  showTrackedCardNote = false,
  knownWordsSummary,
}: TrackingSnapshotProps) {
  const knownWordPercent =
    knownWordsSummary && knownWordsSummary.totalUniqueWords > 0
      ? Math.round((knownWordsSummary.knownWordCount / knownWordsSummary.totalUniqueWords) * 100)
      : null;

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-sm font-semibold text-ctp-text">Tracking Snapshot</h3>
      <p className="mt-1 mb-3 text-xs text-ctp-overlay2">
        Lifetime totals sourced from summary tables.
      </p>
      {showTrackedCardNote && (
        <div className="mb-3 rounded-lg border border-ctp-surface2 bg-ctp-surface1/50 px-3 py-2 text-xs text-ctp-subtext0">
          No lifetime card totals in the summary table yet. New cards mined after this fix will
          appear here.
        </div>
      )}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-sm">
        <Tooltip text="Total immersion sessions recorded across all time">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Sessions</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-lavender">
              {formatNumber(summary.totalSessions)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Total active watch time across all sessions">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Watch Time</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-mauve">
              {summary.allTimeMinutes < 60
                ? `${summary.allTimeMinutes}m`
                : `${(summary.allTimeMinutes / 60).toFixed(1)}h`}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Number of distinct days with at least one session">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Active Days</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-peach">
              {formatNumber(summary.activeDays)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Average active watch time per session in minutes">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Avg Session</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-yellow">
              {formatNumber(summary.averageSessionMinutes)}
              <span className="text-sm text-ctp-overlay2 ml-0.5">min</span>
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Total unique episodes (videos) watched across all anime">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Episodes</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-blue">
              {formatNumber(summary.totalEpisodesWatched)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Number of anime series fully completed">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Anime</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-sapphire">
              {formatNumber(summary.totalAnimeCompleted)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Total Anki cards mined from subtitle lines across all sessions">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Cards Mined</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-cards-mined">
              {formatNumber(summary.totalTrackedCards)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Lifetime Yomitan lookups normalized by total tokens seen">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Lookup Rate</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-flamingo">
              {summary.lookupRate?.shortValue ?? '—'}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Total token occurrences encountered in today's sessions">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Tokens Today</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-sky">
              {formatNumber(summary.todayTokens)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Unique headwords seen for the first time today">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">
              New Words Today
            </div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-rosewater">
              {formatNumber(summary.newWordsToday)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text="Unique headwords seen for the first time this week">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">New Words</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-pink">
              {formatNumber(summary.newWordsThisWeek)}
            </div>
          </div>
        </Tooltip>
        {knownWordsSummary && knownWordsSummary.totalUniqueWords > 0 && (
          <Tooltip text="Words matched against your known-words list out of all unique words seen">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Known Words</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-green">
                {formatNumber(knownWordsSummary.knownWordCount)}
                <span className="text-sm text-ctp-overlay2 ml-1">
                  / {formatNumber(knownWordsSummary.totalUniqueWords)}
                </span>
                {knownWordPercent != null ? (
                  <span className="text-sm text-ctp-overlay2 ml-1">({knownWordPercent}%)</span>
                ) : null}
              </div>
            </div>
          </Tooltip>
        )}
      </div>
    </div>
  );
}
