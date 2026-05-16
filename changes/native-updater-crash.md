---
area: updates
---

- Avoided native `electron-updater` checks on unsupported builds while preserving direct writable AppImage updates, so tray and background update checks continue through GitHub release metadata without crashing the app.
