type: internal
area: build

- Moved packaging inputs (electron-builder resources, Bun runtime manifests and licenses, committed release notes) under `packaging/`.
- Build output now lands in `dist/`: packaged apps in `dist/release/`, staged Yomitan and Hachidori in `dist/yomitan/` and `dist/hachidori/`.
