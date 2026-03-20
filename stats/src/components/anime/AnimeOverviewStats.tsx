import { formatDuration, formatNumber } from '../../lib/formatters';
import { buildLookupRateDisplay } from '../../lib/yomitan-lookup';
import { Tooltip } from '../layout/Tooltip';
import type { AnimeDetailData } from '../../types/stats';

interface AnimeOverviewStatsProps {
  detail: AnimeDetailData['detail'];
  knownWordsSummary: {
    totalUniqueWords: number;
    knownWordCount: number;
  } | null;
}

interface MetricProps {
  label: string;
  value: string;
  unit?: string;
  color: string;
  tooltip: string;
  sub?: string;
}

function Metric({ label, value, unit, color, tooltip, sub }: MetricProps) {
  return (
    <Tooltip text={tooltip}>
      <div className="flex flex-col items-center gap-1 px-3 py-3 rounded-lg bg-ctp-surface1/40 hover:bg-ctp-surface1/70 transition-colors">
        <div className={`text-2xl font-bold font-mono tabular-nums ${color}`}>
          {value}
          {unit && <span className="text-sm font-normal text-ctp-overlay2 ml-0.5">{unit}</span>}
        </div>
        <div className="text-[11px] uppercase tracking-wider text-ctp-overlay2 font-medium">
          {label}
        </div>
        {sub && <div className="text-[11px] text-ctp-overlay1">{sub}</div>}
      </div>
    </Tooltip>
  );
}

export function AnimeOverviewStats({ detail, knownWordsSummary }: AnimeOverviewStatsProps) {
  const lookupRate = buildLookupRateDisplay(detail.totalYomitanLookupCount, detail.totalTokensSeen);

  const knownPct =
    knownWordsSummary && knownWordsSummary.totalUniqueWords > 0
      ? Math.round((knownWordsSummary.knownWordCount / knownWordsSummary.totalUniqueWords) * 100)
      : null;

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4 space-y-3">
      {/* Primary metrics - always 4 columns on sm+ */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <Metric
          label="Watch Time"
          value={formatDuration(detail.totalActiveMs)}
          color="text-ctp-blue"
          tooltip="Total active watch time for this anime"
        />
        <Metric
          label="Sessions"
          value={String(detail.totalSessions)}
          color="text-ctp-peach"
          tooltip="Number of immersion sessions on this anime"
        />
        <Metric
          label="Episodes"
          value={String(detail.episodeCount)}
          color="text-ctp-yellow"
          tooltip="Number of completed episodes for this anime"
        />
        <Metric
          label="Words Seen"
          value={formatNumber(detail.totalTokensSeen)}
          color="text-ctp-mauve"
          tooltip="Total word occurrences across all sessions"
        />
      </div>

      {/* Secondary metrics - fills row evenly */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <Metric
          label="Cards Mined"
          value={formatNumber(detail.totalCards)}
          color="text-ctp-cards-mined"
          tooltip="Anki cards created from subtitle lines in this anime"
        />
        <Metric
          label="Lookups"
          value={formatNumber(detail.totalYomitanLookupCount)}
          color="text-ctp-lavender"
          tooltip="Total Yomitan dictionary lookups during sessions"
        />
        {lookupRate ? (
          <Metric
            label="Lookup Rate"
            value={lookupRate.shortValue}
            color="text-ctp-sapphire"
            tooltip="Yomitan lookups per 100 words seen"
          />
        ) : (
          <Metric
            label="Lookup Rate"
            value="—"
            color="text-ctp-overlay2"
            tooltip="No lookups recorded yet"
          />
        )}
        {knownPct !== null ? (
          <Metric
            label="Known Words"
            value={`${knownPct}%`}
            color="text-ctp-green"
            tooltip={`${formatNumber(knownWordsSummary!.knownWordCount)} known out of ${formatNumber(knownWordsSummary!.totalUniqueWords)} unique words in this anime`}
          />
        ) : (
          <Metric
            label="Known Words"
            value="—"
            color="text-ctp-overlay2"
            tooltip="No word data available yet"
          />
        )}
      </div>
    </div>
  );
}
