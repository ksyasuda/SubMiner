type: changed
area: stats

- Added the Stats Search tab for realtime subtitle sentence search with media context, headword matching, and mining actions for source-backed sentence cards or exact-match word/audio cards.
- Improved Stats mining from Search and vocabulary examples: empty `ankiConnect.deck` can use Yomitan's mining deck, sentence cards are created before slow media generation finishes, secondary subtitles from stored lines, sidecar files, or temporary alass-retimed English sidecars populate sentence Selection Text, invalid stored timings are blocked before FFmpeg runs, future out-of-order subtitle timing pairs are skipped until valid timings arrive, and partial media failures are shown.
- Fixed Stats mining field/audio behavior so sentence clips update `SentenceAudio`, word audio uses the configured Yomitan sources, English subtitle text is not written onto word cards, and secondary subtitle auto-selection prefers regular English tracks over Signs/Songs tracks.
- Improved vocabulary review with a Hide Kana filter, duplicate-collapsed exclusions across token variants, and Related Seen Words matching based on shared readings or kanji.
- Improved Stats browsing reliability by remembering library card size, retrying stored cover art without extra AniList lookups, showing progress during session deletes, and making session deletes refresh faster.
