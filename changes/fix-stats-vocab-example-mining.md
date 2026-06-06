type: fixed
area: stats

- Fixed vocab-page example sentence mining buttons failing when the Anki deck setting is blank or Yomitan card formats are ordered with a non-term card first.
- Fixed vocab-page example word and audio mining so English subtitles are only written to Selection Text for sentence cards, leaving word cards to show the normal Yomitan dictionary glossary.
- Fixed stats-page sentence mining audio updates so generated sentence clips populate `SentenceAudio` without also writing the same clip to the configured expression-audio field.
- Fixed stats-page word mining so the hidden Yomitan helper uses the same configured word-audio sources as the normal Yomitan plus button.
- Fixed stats-page sentence mining to use the current Anki deck/settings at request time, create direct sentence cards before slow media generation completes, and include stored English subtitle text as Selection Text.
- Fixed secondary subtitle auto-selection to prefer regular English tracks over Signs/Songs tracks when both are available.
