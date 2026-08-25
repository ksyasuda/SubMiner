type: fixed
area: stats

- Typeset subtitles no longer flood the stats. Karaoke openings and animated signs are authored as one subtitle event per animation frame and immersion tracking counted every frame, which was enough to pin an OP lyric to the top of "Top Repeated Words" for good. Lines are now collapsed on the way in using the same rules the subtitle sidebar applies, with a strict fallback for shifted or unparsed sources. Ordinary repeated dialogue and rewatches are unaffected.
- Added a cleanup for stats already affected. The Vocabulary tab's **Duplicates** button scans a chosen window (7 days through all time), shows the bursts it found and the word and kanji counts they added, and collapses each run to one line once confirmed. `subminer stats cleanup --duplicate-lines` does the same from the terminal, with `--dry-run` and `--lookback-days <n>`. Watch time and lines-seen totals are left as recorded.
