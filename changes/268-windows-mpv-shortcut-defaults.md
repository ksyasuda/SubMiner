type: fixed
area: launcher

- Fixed the Windows `SubMiner mpv` shortcut and `SubMiner.exe --launch-mpv` flow to launch mpv with SubMiner's required default args directly instead of requiring an `mpv.conf` profile named `subminer`.
- Clarified the Windows install and usage docs so the shortcut path is documented as self-contained, while the optional `subminer` mpv profile remains available for manual mpv launches.
- Hardened the first-run setup blocker copy and stale custom-scheme handling so setup messages stay aligned with config, plugin, and dictionary readiness.
