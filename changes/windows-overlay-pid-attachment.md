type: fixed
area: overlay

- Windows overlay attachment now identifies mpv through its IPC process ID, removing the dependency on PowerShell command-line lookups and allowing the IPC socket to be configured in `mpv.conf`.
