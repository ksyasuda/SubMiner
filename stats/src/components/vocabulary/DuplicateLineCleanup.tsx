import { useCallback, useState } from 'react';
import { getStatsClient } from '../../hooks/useStatsApi';
import { formatNumber } from '../../lib/formatters';
import type { StatsDuplicateLineCleanupResult } from '../../types/stats';

interface DuplicateLineCleanupProps {
  onClose: () => void;
  /** Called after rows are actually removed, so the charts can reload. */
  onCleaned: () => void;
}

const LOOKBACK_CHOICES: Array<{ label: string; days: number | null }> = [
  { label: '7 days', days: 7 },
  { label: '30 days', days: 30 },
  { label: '90 days', days: 90 },
  { label: '1 year', days: 365 },
  { label: 'All time', days: null },
];

function formatTimecode(ms: number): string {
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${String(seconds).padStart(2, '0')}`;
}

export function DuplicateLineCleanup({ onClose, onCleaned }: DuplicateLineCleanupProps) {
  const [lookbackDays, setLookbackDays] = useState<number | null>(30);
  const [preview, setPreview] = useState<StatsDuplicateLineCleanupResult | null>(null);
  const [applied, setApplied] = useState<StatsDuplicateLineCleanupResult | null>(null);
  const [busy, setBusy] = useState<'scan' | 'apply' | null>(null);
  const [error, setError] = useState<string | null>(null);

  const run = useCallback(
    async (dryRun: boolean) => {
      setBusy(dryRun ? 'scan' : 'apply');
      setError(null);
      try {
        const result = await getStatsClient().cleanupDuplicateLines({ dryRun, lookbackDays });
        if (dryRun) {
          setPreview(result);
          setApplied(null);
        } else {
          setApplied(result);
          setPreview(null);
          onCleaned();
        }
      } catch (cause) {
        setError(cause instanceof Error ? cause.message : String(cause));
      } finally {
        setBusy(null);
      }
    },
    [lookbackDays, onCleaned],
  );

  const result = applied ?? preview;
  const nothingToDo = preview !== null && preview.removedLines === 0;

  return (
    <div className="fixed inset-0 z-50">
      <button
        type="button"
        aria-label="Close duplicate line cleanup"
        className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]"
        onClick={onClose}
      />
      <div className="absolute inset-x-0 top-1/2 mx-auto max-w-xl -translate-y-1/2 rounded-xl border border-ctp-surface1 bg-ctp-mantle shadow-2xl">
        <div className="flex items-center justify-between border-b border-ctp-surface1 px-5 py-4">
          <h2 className="text-sm font-semibold text-ctp-text">Duplicate Lines</h2>
          <button
            type="button"
            className="rounded-md border border-ctp-surface2 px-3 py-1.5 text-xs font-medium text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue"
            onClick={onClose}
          >
            Close
          </button>
        </div>

        <div className="space-y-4 px-5 py-4">
          <p className="text-xs leading-relaxed text-ctp-subtext0">
            Typeset subtitles — karaoke openings, animated signs — are authored as one event per
            animation frame, and older versions counted every frame as its own line. This finds
            those runs and collapses each one back to a single line, giving back the word and kanji
            counts they inflated. Ordinary repeated dialogue is left alone.
          </p>

          <div>
            <div className="mb-2 text-xs font-medium text-ctp-subtext1">Look back over</div>
            <div className="flex flex-wrap gap-2">
              {LOOKBACK_CHOICES.map((choice) => (
                <button
                  key={choice.label}
                  type="button"
                  disabled={busy !== null}
                  onClick={() => {
                    setLookbackDays(choice.days);
                    setPreview(null);
                    setApplied(null);
                  }}
                  className={`rounded-lg border px-3 py-1.5 text-xs transition disabled:opacity-50 ${
                    lookbackDays === choice.days
                      ? 'border-ctp-blue/50 bg-ctp-surface2 text-ctp-text'
                      : 'border-ctp-surface1 bg-ctp-surface0 text-ctp-overlay2 hover:text-ctp-subtext0'
                  }`}
                >
                  {choice.label}
                </button>
              ))}
            </div>
          </div>

          {error && (
            <div className="rounded-lg border border-ctp-red/30 bg-ctp-red/10 px-3 py-2 text-xs text-ctp-red">
              {error}
            </div>
          )}

          {result && (
            <div className="rounded-lg bg-ctp-surface0 px-4 py-3">
              <div className="text-sm text-ctp-text">
                {applied
                  ? `Removed ${formatNumber(applied.removedLines)} repeated lines`
                  : nothingToDo
                    ? 'No animation bursts found in this window'
                    : `Found ${formatNumber(preview!.burstGroups)} bursts covering ${formatNumber(preview!.removedLines)} extra lines`}
              </div>
              <div className="mt-1 text-xs text-ctp-overlay2">
                {formatNumber(result.scannedLines)} lines scanned ·{' '}
                {formatNumber(result.removedWordOccurrences)} word counts ·{' '}
                {formatNumber(result.removedKanjiOccurrences)} kanji counts
                {applied ? ' removed' : ' would be removed'}
              </div>

              {result.samples.length > 0 && (
                <div className="mt-3 max-h-52 space-y-1.5 overflow-y-auto">
                  {result.samples.map((sample) => (
                    <div
                      key={`${sample.videoId}:${sample.startMs}:${sample.text}`}
                      className="flex items-center justify-between gap-3 rounded-md bg-ctp-mantle px-3 py-1.5"
                    >
                      <div className="min-w-0">
                        <div className="truncate text-xs text-ctp-text">{sample.text}</div>
                        <div className="truncate text-[11px] text-ctp-overlay1">
                          {sample.videoTitle ?? `Video ${sample.videoId}`} ·{' '}
                          {formatTimecode(sample.startMs)}
                        </div>
                      </div>
                      <span className="shrink-0 text-xs text-ctp-peach">×{sample.frames}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          <div className="flex items-center justify-end gap-2">
            <button
              type="button"
              disabled={busy !== null}
              onClick={() => void run(true)}
              className="rounded-md border border-ctp-surface2 px-3 py-1.5 text-xs font-medium text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue disabled:opacity-50"
            >
              {busy === 'scan' ? 'Scanning…' : 'Scan'}
            </button>
            <button
              type="button"
              disabled={busy !== null || preview === null || nothingToDo}
              onClick={() => void run(false)}
              className="rounded-md border border-ctp-red/30 px-3 py-1.5 text-xs font-medium text-ctp-red transition hover:bg-ctp-red/10 disabled:opacity-40"
            >
              {busy === 'apply' ? 'Cleaning…' : 'Clean Up'}
            </button>
          </div>
          <p className="text-[11px] text-ctp-overlay1">
            Scan first: cleanup removes rows and cannot be undone. Session watch time and lines-seen
            totals are left untouched.
          </p>
        </div>
      </div>
    </div>
  );
}
