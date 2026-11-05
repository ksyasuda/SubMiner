type: added
area: subtitles

- Added subtitle filters (`subtitleFilter`) that remove unwanted text from primary subtitles before it reaches the overlay, sidebar, mined sentences and stats. Built-in presets cover text in parentheses (speaker names, readings, notes; whole-line inner thoughts are kept), square-bracket sound descriptions, speaker names before a colon, and music notes. Custom regex rules with a replacement can be added after the presets. Lines left with no words are hidden.
- Configure them in the settings window under Behavior → Subtitle Filters; changes apply live to the current subtitle file.
