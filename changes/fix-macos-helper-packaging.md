type: fixed
area: release

- Fixed macOS packaging so the compiled mpv window helper is built into `dist/scripts` and required in the app bundle, preventing the overlay from falling back to slow Swift source startup.
- Removed a stale Windows helper resource entry that produced harmless missing-file warnings during packaging.
