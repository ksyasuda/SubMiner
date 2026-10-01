type: added
area: overlay

- Added scroll wheel keys (`WHEEL_UP`, `WHEEL_DOWN`, `WHEEL_LEFT`, `WHEEL_RIGHT`, with modifiers) to `keybindings`, including capture from the settings key editor.
- Mouse button and scroll wheel bindings from mpv (`input.conf` and mpv defaults, e.g. double-click fullscreen, wheel volume, back/forward for playlist) now work while the cursor is over the overlay. On Hyprland, where the overlay always receives input, they previously did nothing. SubMiner bindings and the right-click pause still take priority.
