type: changed
area: setup

- SubMiner-managed mpv launches now inject the bundled mpv plugin when no global SubMiner plugin is installed, setup can remove detected legacy global plugin files via the OS trash, and legacy global plugin install entrypoints have been removed so regular mpv playback stays unaffected.
- Managed playback now surfaces the legacy plugin removal prompt before mpv starts, allowing users to trash the old global plugin and immediately relaunch with the bundled runtime plugin.
