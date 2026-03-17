type: fixed
area: overlay

- Fixed macOS fullscreen overlay stability by keeping the passive visible overlay from stealing focus, re-raising the overlay window when reasserting its macOS topmost level, and tolerating one transient macOS tracker/helper miss before hiding the overlay.
- Kept subtitle tokenization warmup one-shot for the lifetime of the app so later fullscreen/media churn on macOS does not replay the startup warmup gate after the first file is ready.
- Added a bounded macOS tracker loss-grace window so fullscreen enter/leave transitions do not immediately hide and reload the overlay when the helper briefly loses the mpv window.
- Skipped subtitle/tokenization refresh invalidation on character-dictionary auto-sync completion when the dictionary was already current, preventing startup flash/reload loops on unchanged media.
