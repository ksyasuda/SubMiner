type: fixed
area: dictionary

- Character dictionary generation, merged rebuilds, and imports no longer freeze the app (and trigger the compositor's "application not responding" dialog) on large dictionaries; snapshot reads/writes, archive building, and the character image/name lookup caches now do their heavy work off the UI's critical path.
- Desktop progress notifications now update in place on Linux AppImage installs too: the AppImage's bundled libraries broke the system notify-send helper, which silently forced the flickering close-and-reopen notification fallback.
