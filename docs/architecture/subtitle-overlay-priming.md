<!-- read_when: changing visible overlay startup, Linux/X11 overlay window shape, mpv subtitle callbacks, or subtitle tokenization emission -->

# Subtitle Overlay Priming

Status: active
Last verified: 2026-06-01
Owner: Kyle Yasuda
Read when: debugging subtitle state or blank Linux/X11 overlay windows when the visible overlay is shown or recreated

Visible-overlay subtitle priming fills the overlay from mpv's current subtitle properties before
waiting for the next live mpv subtitle event. This avoids a stale or blank overlay when the user
manually shows the visible overlay while playback is already sitting on a subtitle.

On Linux/X11, visible-overlay show and later mpv bounds refreshes restore the Electron window shape
to the full current overlay bounds. Electron's `BrowserWindow.setShape()` applies a bounding shape,
not an input-only region; stale shapes can leave a mapped 1920x1080 overlay with smaller X11 shape
extents such as `800x600+0+0`, so renderer and websocket subtitle state are correct while bottom
subtitles do not draw.

## Entry Points

- `src/main.ts` calls `primeCurrentSubtitleForVisibleOverlay()` when manual visible-overlay show
  paths run.
- `src/main.ts` calls `restoreVisibleOverlayWindowShapeForShow()` before visible-overlay show
  actions on Linux, and `resetVisibleOverlayInputState()` restores a full shape instead of applying
  an empty shape.
- `src/main.ts` also restores the Linux/X11 shape after applying mpv overlay bounds, so a newly
  created 800x600 hidden Electron window cannot keep clipping after it is resized to mpv geometry.
- `primeCurrentSubtitleForVisibleOverlay()` delegates to
  `primeVisibleOverlaySubtitleFromMpv()` in `src/main/runtime/current-subtitle-snapshot.ts`.
- `restoreVisibleOverlayWindowShapeForShow()` delegates to `restoreLinuxOverlayWindowShape()` in
  `src/main/runtime/linux-overlay-window-shape.ts`.
- Inputs are callback deps, not globals: `getMpvClient`, `setCurrentSubText`,
  `getCurrentSubtitleData`, `consumeCachedSubtitle`, `onSubtitleChange`,
  `refreshCurrentSubtitle`, `emitSubtitle`, optional secondary-subtitle callbacks, and `logDebug`.

## Primary Subtitle Flow

1. Read the connected mpv client through `getMpvClient()`. Exit if no connected client.
2. Request mpv `sub-text`. On failure, log a
   `[visible-overlay-subtitle-prime] failed to read sub-text` debug line and exit.
3. Normalize non-string `sub-text` to `''`, then call `setCurrentSubText(text)` so app state
   matches mpv before any overlay emission.
4. Empty text: call `onSubtitleChange(text)`, emit `{ text, tokens: null }`, then prime secondary
   subtitles.
5. Current cached payload: if `getCurrentSubtitleData()?.text === text`, call
   `emitSubtitle(payload)` and `refreshCurrentSubtitle(text)`, then prime secondary subtitles.
6. Tokenization cache hit: call `consumeCachedSubtitle(text)`, `onSubtitleChange(text)`, and
   `emitSubtitle(cachedPayload)`, then prime secondary subtitles.
7. Cache miss: call `refreshCurrentSubtitle(text)` and let normal tokenization emit the final
   payload.

In `src/main.ts`, both `onSubtitleChange` and `refreshCurrentSubtitle` pause
`subtitlePrefetchService`, notify it with `onSeek(lastObservedTimePos)`, and then call the matching
`subtitleProcessingController` method. This gives the visible overlay priority over background
prefetch work and re-centers prefetch around the live playback time.

## Emitted State

- `emitSubtitle(payload)` maps to `emitSubtitlePayload(payload)`, which sends the normal
  annotated subtitle payload to overlay windows and subtitle websocket listeners.
- Secondary priming reads mpv `secondary-sub-text`, stores it in
  `mpvClient.currentSecondarySubText`, and broadcasts `secondary-subtitle:set` to overlay windows.
- If secondary `requestProperty` fails, the primary flow stays complete and only a debug line is
  written.

## Startup Ready Release

- `mpv.pauseUntilOverlayReady` waits for tokenization warmup plus visible-overlay readiness before
  releasing the mpv startup gate.
- Visible-overlay startup creates the tray and visible overlay shell before tokenization and
  annotation warmups continue. Cold `--start --background --managed-playback` launches still handle
  initial args before the deferred Yomitan wait.
- Overlay-routed startup notifications are queued in the main process until an overlay window has
  finished loading. Progress notifications with the same id are upserted so spinner ticks do not
  flood a cold-start overlay, while events with distinct history ids are retained for phase-level
  history such as character dictionary checking/building/importing.
- The mpv plugin has a 30-second fallback for cold starts; app-side retry/release budgets match that
  window so readiness can still arrive before fallback resumes playback.
- If mpv is already on a subtitle, SubMiner still prefers the resolved current subtitle payload and
  waits for a fresh measured subtitle rectangle before signaling readiness.
- If mpv is before the first subtitle, SubMiner sends a synthetic warm readiness payload after
  tokenization warmup and visible overlay content-ready. This releases playback without waiting for
  a later subtitle event that cannot happen while mpv is paused.

## Linux/X11 Window Shape

- `restoreLinuxOverlayWindowShape()` reads `BrowserWindow.getBounds()` and calls `setShape()` with
  one full-window rectangle: `{ x: 0, y: 0, width, height }`.
- Restore the shape after `setBounds()`/mpv geometry updates, not only before showing the overlay.
  Manual startup can create the hidden overlay at Electron's default 800x600 size before the window
  tracker applies the real mpv bounds.
- Do not use `setShape([])` as a passive reset for the visible overlay. On the tested X11/XWayland
  path, empty or stale bounding shapes produced invisible or clipped subtitles even though the
  overlay window remained mapped above mpv.
- Pointer pass-through should continue to use `setIgnoreMouseEvents(true, { forward: true })` and
  the Linux cursor-poll fallback, not bounding-shape clipping.

## Config And Migration

No config or schema migration. This workflow reuses existing mpv properties, overlay IPC events,
subtitle tokenization cache, and prefetch controls.
