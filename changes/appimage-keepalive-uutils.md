type: fixed
area: linux

- Fixed AppImage builds on Ubuntu 26.04+ (uutils coreutils) releasing their mount before Chromium child processes finished shutting down, which brought back "Service Crash" notifications on quit.
