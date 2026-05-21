type: fixed
area: updater

- Routed macOS supplemental GitHub release lookups through `/usr/bin/curl` instead of Electron `net.fetch`, eliminating the last Electron-networking path from background update checks and avoiding the network-service crashes seen in earlier prereleases.
