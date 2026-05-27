type: changed
area: anki

- `ankiConnect.nPlusOne.enabled` is no longer implicitly set to `true` when `ankiConnect.knownWords.highlightEnabled` is `true`. Users who rely on known-word highlighting and want N+1 target highlighting must now set `ankiConnect.nPlusOne.enabled: true` explicitly.
- Updated known-word cache docs and examples to recommend expression/word fields and removed legacy-option references from user-facing config docs.
