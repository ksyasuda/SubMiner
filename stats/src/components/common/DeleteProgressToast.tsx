interface DeleteProgressToastProps {
  /** Number of sessions currently being deleted. The toast is hidden when 0. */
  count: number;
}

/**
 * Fixed-position toast shown while session deletions are in flight.
 *
 * The per-row delete buttons are only visible on hover, so once the confirm
 * dialog closes the user has no signal that a (potentially slow) batch delete
 * is still running. This stays on screen, independent of hover, until the work
 * finishes.
 */
export function DeleteProgressToast({ count }: DeleteProgressToastProps) {
  if (count <= 0) return null;

  return (
    <div
      role="status"
      aria-live="polite"
      className="fixed bottom-4 right-4 z-50 flex items-center gap-3 rounded-lg border border-ctp-surface1 bg-ctp-surface0 px-4 py-3 shadow-lg shadow-black/30"
    >
      <span
        aria-hidden="true"
        className="h-4 w-4 shrink-0 animate-spin rounded-full border-2 border-ctp-surface2 border-t-ctp-red"
      />
      <span className="text-sm text-ctp-text">
        Deleting {count} session{count === 1 ? '' : 's'}&hellip;
      </span>
    </div>
  );
}
