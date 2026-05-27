type: fixed
area: release

- Fixed macOS packaging so the compiled mpv window helper is built into `dist/scripts` and bundled, preventing the overlay from falling back to slow Swift source startup, and removed a stale Windows helper resource entry that produced harmless missing-file warnings.
- Fixed one-shot `make clean build install` flows so install picks up the AppImage built earlier in the same make invocation.
