type: changed
area: config

- `ankiConnect.nPlusOne.enabled` is no longer implicitly set to `true` when known-word highlighting is enabled; existing configs that already had N+1 highlighting keep it, but new configs leave it disabled unless `ankiConnect.nPlusOne.enabled` is set explicitly.
- Updated known-word cache docs and examples to recommend expression/word fields and removed legacy-option references from user-facing config docs.
