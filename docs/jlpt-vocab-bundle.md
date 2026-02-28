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

SubMiner also reuses the same `term_meta_bank_*.json` format for frequency-based subtitle highlighting, using installed/default `frequency-dictionary` locations or an explicit `subtitleStyle.frequencyDictionary.sourcePath`.

## Source and update process

For reproducible updates:

1. Obtain the JLPT term-meta bank archive from the same upstream source that supplies the bundled Yomitan dictionary data.
2. Extract the five `term_meta_bank_*.json` files.
3. Place them into `vendor/yomitan-jlpt-vocab/`.
4. Commit the update with the source URL/version in the task notes.

This repository currently ships the folder path in `electron-builder` `extraResources` as:
`vendor/yomitan-jlpt-vocab -> yomitan-jlpt-vocab`.

## Fallback Behavior

If bank files are missing, malformed, or lack expected metadata, SubMiner skips them gracefully. When no usable entries are found, JLPT underlining is silently disabled and subtitle rendering remains unchanged.
