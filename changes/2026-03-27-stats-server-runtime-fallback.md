type: fixed
area: stats

- Fixed stats startup so the immersion tracker can run when `Bun.serve` is unavailable.
- Stats server now falls back to a Node `http` listener in Electron/runtime paths that do not expose Bun.
