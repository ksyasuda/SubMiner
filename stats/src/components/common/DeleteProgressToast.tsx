import { useSyncExternalStore } from 'react';
import { getDeleteProgressSnapshot, subscribeDeleteProgress } from '../../lib/delete-progress';

/**
 * Global "deletion in progress" indicator.
 *
 * Mounted once at the app root, outside every tab panel, so it shows no matter
 * which view started the delete — per-tab copies were hidden along with their
 * `hidden` tab panel, and detail views had no indicator at all. State comes
 * from the delete-progress store rather than props so it also survives the
 * initiating component unmounting mid-request.
 *
 * Renders two signals: a sweeping bar pinned to the top edge (visible even when
 * the eye is on the content being deleted) and a bottom-right toast naming the
 * work.
 */
export function DeleteProgressToast() {
  const { count, label } = useSyncExternalStore(
    subscribeDeleteProgress,
    getDeleteProgressSnapshot,
    getDeleteProgressSnapshot,
  );

  if (count <= 0) return null;

  const message = count > 1 ? `Deleting ${count} items` : (label ?? 'Deleting');

  return (
    <>
      <div
        aria-hidden="true"
        className="fixed inset-x-0 top-0 z-[2147483646] h-0.5 overflow-hidden bg-ctp-surface1"
      >
        <div className="h-full w-1/3 animate-indeterminate rounded-full bg-gradient-to-r from-ctp-red via-ctp-peach to-ctp-red" />
      </div>
      <div
        role="status"
        aria-live="polite"
        className="fixed bottom-4 right-4 z-[2147483646] flex items-center gap-3 rounded-lg border border-ctp-surface1 bg-ctp-surface0 px-4 py-3 shadow-lg shadow-black/30"
      >
        <span
          aria-hidden="true"
          className="h-4 w-4 shrink-0 animate-spin rounded-full border-2 border-ctp-surface2 border-t-ctp-red"
        />
        <span className="text-sm text-ctp-text">{message}&hellip;</span>
      </div>
    </>
  );
}
