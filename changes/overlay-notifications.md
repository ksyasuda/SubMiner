type: changed
area: notifications
breaking: true

- Added overlay notifications with a Catppuccin Macchiato stack, a 3-second transient timeout, and persistent long-running job notifications for character dictionary sync.
- Added `notifications.overlayPosition` to place overlay notifications at the top left, top center, or top right; top right remains the default.
- Added a notification history panel (default `Ctrl/Cmd+N`, configurable via `shortcuts.toggleNotificationHistory`) that logs every notification shown during the session; the toggle works whether the overlay or mpv has focus, the panel slides in from the same edge as notifications (right when centered), and entries can be removed individually or cleared.
- Made the overlay error/recovery toast follow the configured `notifications.overlayPosition` instead of always pinning to the top-right corner, and kept the notification stack and history panel side synced from that position before first open so left-side history panels slide in from the left.
- Routed startup tokenization, subtitle annotation, and character dictionary status through queued overlay notifications for `overlay`/`both` instead of falling back to mpv OSD while the overlay loads; queued loading cards are shown before their ready update when both happen before the overlay is ready, and the bundled mpv plugin now only emits startup OSD messages for `osd` and `osd-system`.
- Preserved character dictionary checking/building/importing/ready phases in overlay notification history and sent those phases to system notifications when `notificationType` is `both`.
- Initialized the tray and visible overlay shell before deferred tokenization warmups finish on visible-overlay startup, while keeping playback paused until SubMiner reports autoplay readiness.
- Kept playback feedback such as subtitle visibility, subtitle track, and subtitle delay text on overlay/OSD surfaces only; desktop/system notifications are reserved for real notifications like mined cards, errors, and updates.
- Reused the active primary/secondary subtitle mode overlay notification while cycling modes so rapid toggles update one card instead of stacking duplicate feedback.
- Updated repeated progress notifications such as subsync syncing in place so their spinner stays live instead of flickering on every tick.
- Fixed mined-card overlay notifications so `overlay` and `both` modes show generated card thumbnails in both live cards and the notification history panel.
- Added Open in Anki buttons to mined-card overlay notifications and their history entries, with a direct AnkiConnect fallback when the live integration is unavailable.
- Added an Update button to overlay update-available notifications so users can start the app update flow from the notification.
- Fixed sentence-card mining so the Ctrl+S flow shows only the Anki update progress notification instead of also stacking a generic SubMiner toast.
- Fixed overlay notification layering so notification close/actions stay clickable above subtitle bars on Linux overlays.
- Fixed character dictionary sync so duplicate MPV media-path events do not repeat check/ready notifications for the same opened video.
- Changed `both` notification routing to mean overlay + system; users who used `both` for mpv OSD + system notifications should set `notificationType` to `osd-system` in `config.jsonc`.
- Kept `osd` and `osd-system` as config-file-only legacy notification values; Settings normally offers only overlay, system, both, and none, while still showing an already configured legacy value as selected.
