type: added
area: overlay

- Added Chrome Gamepad API controller support for keyboard-only overlay mode, including configurable logical bindings for lookup, mining, popup navigation, Yomitan audio, mpv pause, d-pad fallback navigation, and slower smooth popup scrolling.
- Added `Alt+C` controller selection and `Alt+Shift+C` controller debug modals, with preferred controller persistence and live raw input inspection.
- Added a transient in-overlay controller-detected indicator when a controller is first found.
- Fixed stale keyboard-only token highlight cleanup when keyboard-only mode turns off or the Yomitan popup closes.
