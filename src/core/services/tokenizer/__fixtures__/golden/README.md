# Tokenizer golden corpus

End-to-end regression fixtures for the tokenizer/annotation pipeline. Each
`.json` file captures one real subtitle line together with everything the
pipeline consumed while tokenizing it live:

- `recording.messages` — the raw Yomitan backend responses
  (`chrome.runtime.sendMessage` level: `optionsGetFull`, `getDictionaryInfo`,
  `parseText`, `termsFind`, `getTermFrequencies`), pruned to the fields the
  injected scanning helpers actually read.
- `recording.scripts` — sha256 → result pairs for each injected script, used
  only as a replay fallback if a script cannot run in the vm.
- `recording.mecab` — raw MeCab tokens keyed by the tokenized text.
- `config` — annotation toggles plus fixture-local known words, JLPT levels,
  and local frequency ranks (see `golden-corpus-harness.ts` for the simplified
  known-word semantics shared by recorder and replay).
- `expected.tokens` — the annotated tokens the full pipeline produced at
  record time. **This is the assertion.** Review it before committing; it is a
  characterization of current behavior, not a statement that the behavior is
  ideal.

`golden-corpus.test.ts` replays every fixture through the real
`tokenizeSubtitle` (scan-token merge, MeCab enrichment, frequency ranks,
annotation stage, noise suppression) by executing the real injected scripts in
a `node:vm` sandbox against the recorded responses — no Electron, no
dictionaries, no network.

## Recording a fixture

Requires a built Yomitan extension (`bun run build:yomitan`), your SubMiner
Yomitan profile (`~/.config/SubMiner`), and MeCab:

```sh
bun run record-tokenizer-fixture:electron -- \
  --name my-regression-case \
  --issue "#123" \
  --description "what behavior this pins down" \
  --known-word 私 --known-word 要る:いる \
  --jlpt 美しい=N4 \
  そのまま字幕の一行
```

Run with `--help` for all options (annotation toggles, match mode, overwrite
with `--force`, ...). The recorder prints the expected tokens for review and
the fixture replays immediately via:

```sh
bun test src/core/services/tokenizer/golden-corpus.test.ts
```

## When a fixture fails

A failure means the pipeline now produces different annotated tokens for that
line. If the change is intentional, re-record the fixture with `--force`
(same flags — they are stored in the fixture's `config`) and review the diff
of `expected.tokens`; the diff _is_ the behavior change. If the change is not
intentional, you found the regression before shipping it.

Fixture dictionaries reflect whatever was installed in the recording profile,
so re-recorded fixtures may differ in frequency ranks if dictionaries changed.

## Comparing against stock Yomitan

`bun run compare-yomitan-api:electron` diffs SubMiner's tokenization against a
stock Yomitan instance reached through the
[yomitan-api](https://github.com/yomidevs/yomitan-api) bridge
(`http://127.0.0.1:19633`, enable "Yomitan API" in the browser extension's
settings). Without arguments it compares every fixture text; pass sentences or
`--file <path>` for ad-hoc checks. It reports segmentation, reading, and
headword-form divergence and exits non-zero on any difference — useful for
spot-checking that the pipeline still matches what Yomitan itself would
produce, with your real browser profile and dictionaries as the reference.
