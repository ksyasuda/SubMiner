type: fixed
area: sync

- Fixed the app staying alive in the background after closing a standalone Sync window (`subminer sync --ui`) on macOS. The quit re-issued after async will-quit cleanup was dropped by Electron when the cleanup settled within the same tick as the will-quit dispatch; the re-quit now runs on a fresh macrotask, and the standalone Sync window close path uses the same forced-exit fallback as SIGTERM. Windows opened from the tray menu of a running app are unaffected: closing them still leaves the app running.
- Fixed cancelling or timing out a sync run hanging when the sync child had already exited without emitting a terminal result: the run waited on `close`, which a descendant holding the inherited stdio pipes can delay indefinitely, stalling the quit-time sync shutdown. Cancellation now settles the run immediately once the child has exited.
