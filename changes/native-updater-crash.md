type: fixed
area: updates

- Avoided native `electron-updater` checks where they are unsafe, so tray and background update checks continue through GitHub release metadata without crashing the app.
