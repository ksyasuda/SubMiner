# JLPT Vocabulary Bundle (Offline)

## Bundle location

SubMiner expects the JLPT term-meta bank files to be available locally at:

- `vendor/yomitan-jlpt-vocab`

At runtime, SubMiner also searches these derived locations:

- `vendor/yomitan-jlpt-vocab`
- `vendor/yomitan-jlpt-vocab/vendor/yomitan-jlpt-vocab`
- `vendor/yomitan-jlpt-vocab/yomitan-jlpt-vocab`

and user-data/config fallback paths (see `getJlptDictionarySearchPaths` in `src/main.ts`).

## Required files

The expected files are:

- `term_meta_bank_1.json`
- `term_meta_bank_2.json`
- `term_meta_bank_3.json`
- `term_meta_bank_4.json`
- `term_meta_bank_5.json`

Each bank maps terms to frequency metadata; only entries with a `frequency.displayValue` are considered for JLPT tagging.

SubMiner also reuses the same `term_meta_bank_*.json` format for frequency-based subtitle highlighting. The default frequency source is now bundled as `vendor/jiten_freq_global`, so users can enable `subtitleStyle.frequencyDictionary` without extra setup.

## Source and update process

For reproducible updates:

1. Obtain the JLPT term-meta bank archive from the same upstream source that supplies the bundled Yomitan dictionary data.
2. Extract the five `term_meta_bank_*.json` files.
3. Place them into `vendor/yomitan-jlpt-vocab/`.
4. Commit the update with the source URL/version in the task notes.

This repository currently ships the folder path in `electron-builder` `extraResources` as:
`vendor/yomitan-jlpt-vocab -> yomitan-jlpt-vocab`.

## Deterministic fallback behavior on malformed inputs

`createJlptVocabularyLookupService()` follows these rules:

- If a bank file is missing, parsing fails, or the JSON shape is unsupported, that file is skipped and processing continues.
- If entries do not expose expected frequency metadata, they are skipped.
- If no usable bank entries are found, SubMiner initializes a no-op JLPT lookup (`null` for every token).
- In all fallback cases, subtitle rendering remains unchanged (no underlines are added).

## Bundle size and startup cost

Lookup work is currently a synchronous file read + parse at enable-time and then O(1) in-memory `Map` lookups during subtitle updates.

Practical guidance:

- Keep the JLPT bundle inside `vendor/yomitan-jlpt-vocab` to avoid network lookups.
- Measure bundle size with:
  - `du -sh vendor/yomitan-jlpt-vocab`
- If the JLPT source is updated, re-run `bun run build:appimage` / packaging and confirm startup logs do not report missing banks.
