type: added
area: subsync

- Integrated built-in alass-based subtitle syncing with sign/song stripping, framerate and scene-cut/ad-break handling, audio alignment verification across all loaded/embedded/sidecar tracks, and built-in speech-activity audio fallback.
- Added one-click automatic syncing via the subtitle sidebar Sync button and `Ctrl+Alt+S` (`subsync.mode: "auto"`), automatically opening the guided Subtitle Sync modal with diagnostics, file browsing, external engine fallbacks, and manual offset only when guidance is needed.
- In-place sync (`subsync.replace`, on by default) keeps the untouched original next to the subtitle as `<file>.orig` on the first overwrite.
- Switching to other media while a sync is running cancels it right away instead of finishing in the background and blocking the next sync.
