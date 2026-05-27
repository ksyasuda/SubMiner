type: changed
area: config

- Primary and secondary subtitle appearance now use color controls plus CSS declaration editors, saved as `subtitleStyle.css` and `subtitleStyle.secondary.css`; sidebar appearance uses `subtitleSidebar.css`.
- Moved known-word and N+1 annotation colors to `subtitleStyle.knownWordColor` and `subtitleStyle.nPlusOneColor`; legacy Anki color keys are still accepted with deprecation warnings.
- Updated subtitle font defaults to `Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP`.
- Existing configs are migrated automatically: legacy primary/secondary appearance options and hover token colors fold into `subtitleStyle.css`, and user config files are preserved during legacy compatibility handling.
- Live Settings saves apply subtitle CSS declarations immediately to open video overlays, and the generated example config uses the same CSS declaration paths.
