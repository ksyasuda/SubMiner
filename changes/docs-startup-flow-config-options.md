type: changed
area: docs

- Documented all config options that were present in `config.example.jsonc` but missing from the configuration reference: `subtitleStyle.primaryDefaultMode`, `stats.markWatchedKey`, `immersionTracking.lifetimeSummaries.*`, and all seven `mpv.*` launcher options (`socketPath`, `backend`, `autoStartSubMiner`, `pauseUntilOverlayReady`, `subminerBinaryPath`, `aniskipEnabled`, `aniskipButtonKey`).
- Added a **Playback Startup Flow** diagram to the Architecture page showing how the managed launch (`subminer` CLI, app, Windows shortcut) injects the plugin, establishes the IPC socket, and brings up the overlay via the two convergent triggers.
- Added a **Runtime Sockets** section and diagram to the IPC + Runtime Contracts page showing the mpv IPC socket and app control socket topology.
- Added cross-reference pointers in the MPV Plugin and Troubleshooting pages.
