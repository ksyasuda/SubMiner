type: changed
area: stats

- Added the Stats Search tab for realtime subtitle sentence search with media context, headword matching, and mining actions for source-backed sentence cards or exact-match word/audio cards.
- Improved Stats mining from Search and vocabulary examples: empty `ankiConnect.deck` can use Yomitan's mining deck, sentence cards are created before slow media generation finishes, stored/requested secondary subtitles are preserved before falling back to sidecar files or temporary alass-retimed English sidecars for sentence Selection Text, invalid stored timings are blocked before FFmpeg runs, future out-of-order subtitle timing pairs are skipped until valid timings arrive, and partial media failures are shown.
- Fixed Stats mining field/audio behavior so sentence clips update `SentenceAudio`, word audio uses the configured Yomitan sources, English subtitle text is not written onto word cards, and secondary subtitle auto-selection prefers regular English tracks over Signs/Songs tracks.
- Improved vocabulary review with remembered Hide Known/Hide Kana filters, duplicate-collapsed exclusions across token variants, and Related Seen Words matching based on shared readings or kanji.
- Reorganized the Stats Trends tab into clearer Activity, Cumulative Totals, Efficiency, Patterns, and Library sections, disambiguated per-period vs cumulative charts, and added Reading Speed (words/min) and Cards/Hour efficiency charts.
- Improved Stats browsing reliability by remembering library card size, retrying stored cover art without extra AniList lookups, preserving PNG/WebP cover MIME types, honoring custom AnkiConnect URLs for Browse, showing progress during session deletes, and making session deletes refresh faster.
