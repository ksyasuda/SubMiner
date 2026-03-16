import type { ExcludedWord } from '../../hooks/useExcludedWords';

interface ExclusionManagerProps {
  excluded: ExcludedWord[];
  onRemove: (w: ExcludedWord) => void;
  onClearAll: () => void;
  onClose: () => void;
}

export function ExclusionManager({ excluded, onRemove, onClearAll, onClose }: ExclusionManagerProps) {
  return (
    <div className="fixed inset-0 z-50">
      <button
        type="button"
        aria-label="Close exclusion manager"
        className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]"
        onClick={onClose}
      />
      <div className="absolute inset-x-0 top-1/2 mx-auto max-w-lg -translate-y-1/2 rounded-xl border border-ctp-surface1 bg-ctp-mantle shadow-2xl">
        <div className="flex items-center justify-between border-b border-ctp-surface1 px-5 py-4">
          <h2 className="text-sm font-semibold text-ctp-text">
            Excluded Words
            <span className="ml-2 text-ctp-overlay1 font-normal">({excluded.length})</span>
          </h2>
          <div className="flex items-center gap-2">
            {excluded.length > 0 && (
              <button
                type="button"
                className="rounded-md border border-ctp-red/30 px-3 py-1.5 text-xs font-medium text-ctp-red transition hover:bg-ctp-red/10"
                onClick={onClearAll}
              >
                Clear All
              </button>
            )}
            <button
              type="button"
              className="rounded-md border border-ctp-surface2 px-3 py-1.5 text-xs font-medium text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue"
              onClick={onClose}
            >
              Close
            </button>
          </div>
        </div>
        <div className="max-h-80 overflow-y-auto px-5 py-3">
          {excluded.length === 0 ? (
            <div className="py-6 text-center text-sm text-ctp-overlay2">
              No excluded words yet. Use the Exclude button on a word's detail panel to hide it from stats.
            </div>
          ) : (
            <div className="space-y-1.5">
              {excluded.map(w => (
                <div
                  key={`${w.headword}\0${w.word}\0${w.reading}`}
                  className="flex items-center justify-between rounded-lg bg-ctp-surface0 px-3 py-2"
                >
                  <div className="min-w-0">
                    <span className="text-sm font-medium text-ctp-text">{w.headword}</span>
                    {w.reading && w.reading !== w.headword && (
                      <span className="ml-2 text-xs text-ctp-subtext0">{w.reading}</span>
                    )}
                  </div>
                  <button
                    type="button"
                    className="shrink-0 rounded-md border border-ctp-surface2 px-2 py-1 text-xs text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue"
                    onClick={() => onRemove(w)}
                  >
                    Restore
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
