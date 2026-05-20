type: fixed
area: updates

- Windows automatic updates now keep the native `electron-updater`/NSIS install path enabled while routing updater HTTP through main-process fetch, avoiding the delayed app exit seen shortly after launch without requiring `curl.exe`.
