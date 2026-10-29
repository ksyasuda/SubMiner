type: fixed
area: anki

- `ankiConnect.fields.audio` now defaults to `SentenceAudio` instead of `ExpressionAudio`, so generated sentence audio no longer overwrites Yomitan's word audio on default configs. `fields.wordAudio` still defaults to `ExpressionAudio`. If you relied on the old default, set `fields.audio` to `ExpressionAudio` explicitly.
