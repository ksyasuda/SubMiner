type: added
area: overlay

- Added a `kwin` backend for KDE Plasma Wayland and auto-detected it in the launcher, app runtime, and mpv plugin.
- Added a KWin-backed mpv window tracker so the overlay can follow native Wayland mpv windows on Plasma.
- Expanded the KWin bridge so the overlay and mpv stay coupled on Plasma Wayland: overlay geometry follows mpv, minimizes/restores with mpv, and raises as a pair when either window is focused.
- Native Plasma Wayland still keeps pointer events enabled for the overlay; clickthrough remains a separate limitation.
