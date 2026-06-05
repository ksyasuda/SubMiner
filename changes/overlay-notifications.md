type: changed
area: notifications
breaking: true

- Added overlay notifications with a Catppuccin Macchiato stack, a 3-second transient timeout, and persistent long-running job notifications for character dictionary sync.
- Added `notifications.overlayPosition` to place overlay notifications at the top left, top center, or top right; top right remains the default.
- Routed startup tokenization and subtitle annotation status through the configured notification surfaces; the bundled mpv plugin now only emits startup OSD messages for `osd` and `osd-system`.
- Kept playback feedback such as subtitle visibility, subtitle track, and subtitle delay text on overlay/OSD surfaces only; desktop/system notifications are reserved for real notifications like mined cards, errors, and updates.
- Reused the active primary/secondary subtitle mode overlay notification while cycling modes so rapid toggles update one card instead of stacking duplicate feedback.
- Changed `both` notification routing to mean overlay + system; users who used `both` for mpv OSD + system notifications should set `notificationType` to `osd-system` in `config.jsonc`.
- Kept `osd` and `osd-system` as config-file-only legacy notification values; Settings normally offers only overlay, system, both, and none, while still showing an already configured legacy value as selected.
