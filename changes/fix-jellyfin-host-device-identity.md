type: fixed
area: jellyfin

- Derived Jellyfin cast device identity from the OS hostname, always reports the client as SubMiner, and ignores legacy configurable Jellyfin client/device identity fields so multiple SubMiner installs no longer share the same remote-session identity.
