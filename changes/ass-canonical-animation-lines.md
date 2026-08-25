type: fixed
area: subtitles

- Typeset ASS karaoke and animated signs no longer flood the primary overlay, subtitle sidebar, immersion history, or sentence mining with glyph fragments and per-frame color phases. The authored line is reconstructed once and shown from its first sung syllable until its hold ends, so a lyric's invisible lead-in and fade tail no longer surface it beside the line actually being sung.
- Reconstructed lines keep their authored word spacing, including gaps encoded only by fragment positions or hidden behind wide syllable chunks, so translations stop running words together.
- Ordinary repeated dialogue, separately positioned signs, wrapped lyric rows, and multi-row CC-style dialogue blocks are still published, and dialogue spoken while a song's animation is on screen stays intact instead of being replaced by the lyric.
- Decorative layers stay out of the published text: highlight sweeps, drop-shadow and glow copies, symbol-font glyph decoration, particle-build glyph swarms, sign walls, near-invisible texture strings, and hidden zero-scaled or zero-clipped text.
- On the live overlay, a finished lyric whose exit ghosts still linger in mpv's text keeps explaining those fragments instead of reappearing beside the line that has already taken over.
- Event-heavy karaoke files that stalled subtitle loading for several seconds now parse in well under a second.
- The secondary subtitle overlay collapses layered duplicate lines from animated tracks even when the full karaoke heuristic does not apply.
