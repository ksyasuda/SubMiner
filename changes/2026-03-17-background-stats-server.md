type: added
area: stats

- Added `subminer stats -b` to start or reuse a dedicated background stats server without blocking normal SubMiner instances.
- Added `subminer stats -s` to stop the dedicated background stats server without closing browser tabs.
- Stats server startup now reuses a running background stats daemon instead of trying to bind a second local server in another SubMiner instance.
