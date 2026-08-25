type: fixed
area: overlay

- Secondary subtitles now parse the selected ASS/SRT/VTT source through the primary subtitle deduplication pipeline, so layered animation text stops appearing several times in the overlay, mined cards, and statistics.
- Long ASS lines repeated as dialogue or as positioned signs collapse when they differ only in whitespace or terminal punctuation, and dense multi-row sign layouts no longer become concatenated primary or secondary lines.
- Live mpv text remains the fallback for unreadable tracks and applies full-line duplicate filtering before display; a failed source refresh clears ASS-only cleanup so fallback text from other formats stays intact.
