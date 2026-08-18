<!-- read_when: changing visible overlay startup, Linux/X11 overlay window shape, mpv subtitle callbacks, or subtitle tokenization emission -->

# Subtitle Overlay Priming

Status: active
Last verified: 2026-08-17
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
7. Cache miss: call `refreshCurrentSubtitle(text)`. Normal processing emits a plain payload
   synchronously, then replaces it with the tokenized payload when ready.

Both `onSubtitleChange` and `refreshCurrentSubtitle` pause `subtitlePrefetchService` and then call
the matching `subtitleProcessingController` method, giving the visible overlay priority over
background prefetch work. Prefetch is not re-centered here: restarting the run per line
(`onSeek`) discarded the in-flight tokenization every time the subtitle changed, so only real
seeks restart it (see `onTimePosUpdate` in `src/main.ts`).

On an uncached autoplay prime the raw payload is emitted here and reported to the controller with
`notePlainSubtitleEmitted`, so the controller skips its own plain emit for that line and the
overlay receives one plain payload followed by the annotated one.

The pause is released by the controller's `onProcessingSettled` callback, which fires once it has
no work left. Emits do not release it: the first emit for an uncached line is the plain payload
that precedes tokenization, and a run can finish without emitting at all (a suppressed duplicate,
a failed tokenization). Both controller methods return whether processing is now pending, and the
caller resumes immediately when it is not — a repeated subtitle schedules no work, so no settle is
coming and prefetching would otherwise idle for the rest of the cue.

## Live Cue Delivery

- A tokenization cache miss emits the plain cue synchronously. Tokenization remains serialized so
  live work does not contend for Yomitan state.
- If a newer cue arrives while an older line is still tokenizing, the newer plain cue or empty
  clear payload is emitted immediately. The older tokenization result is dropped before it can
  replace the current cue.
- The current cue upgrades in place when its tokens and annotations are ready. This can reflow text
  or character images, but cue visibility does not wait for that work.

## Secondary Subtitle Flow

- `secondary-sub-text` remains the immediate fallback, so unreadable and remote subtitle sources
  still appear without waiting for file resolution.
- `secondary-subtitle-track.ts` resolves `secondary-sid` against mpv's track list. External tracks
  are read directly; supported embedded text tracks are extracted through the same ffmpeg-backed
  source resolver used by primary subtitle prefetching.
- The selected source is parsed with `parseSubtitleCues()`, including metadata-aware ASS duplicate
  and animation collapse. Playback `time-pos` selects the active parsed cue after applying
  `secondary-sub-delay`.
- The resolved text is stored in `mpvClient.currentSecondarySubText` before it is broadcast. The
  overlay, mining, timing tracker, and immersion statistics therefore consume the same secondary
  text when a readable source is available.
- Media and `secondary-sid` changes clear the previous parsed state before refreshing the source;
  track-list changes refresh without discarding an unchanged source. Observed
  `secondary-sub-delay` changes retime the active parsed cue without rereading the file. If loading,
  extraction, or parsing fails, the controller returns to live mpv text and the renderer's
  conservative short stack heuristic remains the final display fallback.

## Emitted State

- `emitSubtitle(payload)` maps to `emitSubtitlePayload(payload)`. Overlay windows and annotation
  websocket listeners receive both the immediate plain cue and its later annotation upgrade.
- The basic subtitle websocket receives the immediate plain cue only. Because its serialized
  payload discards annotations, the later upgrade would be an identical duplicate and is skipped
  when text and cue timing match.
- Secondary priming reads mpv `secondary-sub-text` and routes it through the secondary track
  controller. A parsed active cue replaces the live text when the selected source is readable.
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
- If the startup subtitle has no cached annotations yet, autoplay priming emits a raw first-paint
  subtitle payload before background tokenization. The tokenized payload replaces it when ready, but
  the visible overlay can paint and measure the line before the mpv startup gate resumes playback.
- If startup `sub-text` is temporarily empty, autoplay priming refreshes the active subtitle source
  and then awaits cue-based priming before synthetic warm readiness can proceed. A parsed current or
  imminent cue is treated as the startup subtitle so the visible overlay can paint and measure it
  before playback resumes.
- If mpv is before the first subtitle, SubMiner sends a synthetic warm readiness payload after
  tokenization warmup and visible overlay content-ready. This releases playback without waiting for
  a later subtitle event that cannot happen while mpv is paused.
- After a synthetic warm readiness release, SubMiner briefly polls/refreshes the current subtitle
  again. This covers Linux/mpv startup cases where `sub-text` is still empty while paused but becomes
  available right after playback resumes, without waiting for the next subtitle property change.

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
  the Linux cursor-poll fallback, not bounding-shape clipping. Note that on Windows click-through
  must go through `applyOverlayClickThrough()` (`src/core/services/overlay-click-through.ts`),
  which omits `forward: true` there: Electron implements forwarding with a global low-level mouse
  hook that lags mouse input system-wide whenever the main thread stalls; the Windows cursor poll
  handles overlay wake-up instead.
- Visible-overlay show/reset marks Linux pointer passthrough state dirty even when the logical
  interaction state is already inactive. The next cursor-poll tick must still reapply
  `setIgnoreMouseEvents(true, { forward: true })`; otherwise a newly shown Electron overlay can keep
  full-window input capture and block both mpv and overlay controls before the first subtitle
  measurement.
- Visible-overlay show also starts a short Linux input grace before the first content measurement.
  Native Wayland surfaces can become inert while `setIgnoreMouseEvents(true)` is active; keeping the
  overlay interactive during this startup gap lets notifications and overlay mouse bindings work
  until subtitle/sidebar/notification rectangles are reported.

## Config And Migration

No config or schema migration. This workflow reuses existing mpv properties, overlay IPC events,
subtitle tokenization cache, and prefetch controls.
