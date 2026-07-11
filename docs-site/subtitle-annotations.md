# Subtitle Annotations

SubMiner annotates subtitle tokens in real time as they appear in the overlay. Four annotation layers work together to surface useful context while you watch: **N+1 highlighting**, **character-name highlighting**, **frequency highlighting**, and **JLPT tagging**.

All four are opt-in and configured under `subtitleStyle`, `ankiConnect.knownWords`, and `ankiConnect.nPlusOne` in your config. They apply independently - you can enable any combination.

::: tip Tokenization
SubMiner's primary tokenizer is Yomitan itself - subtitle text is tokenized based entirely on the dictionaries you have installed in Yomitan. Installing many large dictionaries can increase noise and slow down lookups, so be selective about which dictionaries you install and their priority order.
:::

Before any of those layers render, SubMiner strips annotation metadata from tokens that are usually just subtitle glue or annotation noise. Standalone particles, auxiliaries, adnominals, common explanatory endings like `んです` / `のだ`, merged trailing quote-particle forms like `...って`, auxiliary-stem grammar tails like `そうだ` (MeCab POS3 `助動詞語幹`), repeated kana interjections, and similar non-lexical helper tokens remain hoverable in the subtitle text, but they render as plain tokens without known-word, N+1, frequency, JLPT, or name-match annotation styling.

## N+1 Word Highlighting

N+1 highlighting identifies sentences where you know every word except one, making them ideal mining targets. When enabled, SubMiner builds a local cache of your known vocabulary from Anki and highlights tokens accordingly.

**How it works:**

1. SubMiner queries your configured Anki decks for expression/word fields such as `Expression` or `Word`.
2. The results are cached locally (`known-words-cache.json`) and refreshed on a configurable interval.
3. When a subtitle line appears, each token is checked against the cache.
4. If exactly one unknown word remains in the sentence, it is highlighted with `subtitleStyle.nPlusOneColor` (default: `#c6a0f6`).
5. Already-known tokens can optionally display in `subtitleStyle.knownWordColor` (default: `#a6da95`).

**Key settings:**

| Option                                    | Default      | Description                                                                                                                                                   |
| ----------------------------------------- | ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ankiConnect.knownWords.highlightEnabled` | `false`      | Enable known-word cache lookups used by N+1 highlighting                                                                                                      |
| `ankiConnect.knownWords.refreshMinutes`   | `1440`       | Minutes between Anki cache refreshes                                                                                                                          |
| `ankiConnect.knownWords.decks`            | `{}`         | Deck→fields map for known-word cache queries                                                                                                                  |
| `ankiConnect.knownWords.matchMode`        | `"headword"` | `"headword"` (dictionary form) or `"surface"` (raw text)                                                                                                      |
| `ankiConnect.nPlusOne.enabled`            | `false`      | Enable N+1 target highlighting                                                                                                                               |
| `ankiConnect.nPlusOne.minSentenceWords`   | `3`          | Minimum tokens in a sentence for N+1 to trigger                                                                                                               |
| `subtitleStyle.nPlusOneColor`             | `#c6a0f6`    | Color for the single unknown target word                                                                                                                      |
| `subtitleStyle.knownWordColor`            | `#a6da95`    | Color for already-known tokens                                                                                                                                |

Prefer expression/word fields for `ankiConnect.knownWords.decks`. Reading-only fields can mark unrelated homophones as known, so only include them when that tradeoff is intentional.

::: tip
Set `refreshMinutes` to `1440` (24 hours) for daily sync if your Anki collection is large.
:::

## Character-Name Highlighting

Character-name matches are built from the active merged SubMiner character dictionary, which auto-syncs character data from AniList for your recently-watched titles. When the current AniList media ID is known, SubMiner ignores loaded entries from other titles for subtitle name matching and inline portraits. Matching names are highlighted in subtitles and become available for hover-driven Yomitan character profiles - portraits, roles, voice actors, and biographical detail.

**How it works:**

1. Subtitles are tokenized, then candidate name tokens are matched against the character dictionary via Yomitan's scanning pipeline.
2. Matching tokens receive a dedicated style distinct from N+1 and frequency layers.
3. This layer can be independently toggled with `subtitleStyle.nameMatchEnabled`.
4. When `subtitleStyle.nameMatchImagesEnabled` is also enabled, SubMiner shows the cached AniList portrait beside matched names.

**Key settings:**

| Option                                 | Default   | Description                                      |
| -------------------------------------- | --------- | ------------------------------------------------ |
| `subtitleStyle.nameMatchEnabled`       | `false`   | Enable character-name token highlighting         |
| `subtitleStyle.nameMatchImagesEnabled` | `false`   | Show small AniList portraits next to name tokens |
| `subtitleStyle.nameMatchColor`         | `#f5bde6` | Color used for character-name matches            |

For full details on dictionary generation, name variant expansion, auto-sync lifecycle, and configuration, see the dedicated [Character Dictionary](/character-dictionary) page.

## Frequency Highlighting

Frequency highlighting colors tokens based on how common they are, using dictionary frequency rank data. This helps you spot high-value vocabulary at a glance. For each token, ranks from the installed Yomitan frequency dictionaries are consulted in priority order: the highest-priority dictionary that has the term wins, lower-priority dictionaries fill in terms it lacks, and occurrence-based dictionaries are skipped.

**Modes:**

- **Single** - all highlighted tokens share one color (`singleColor`).
- **Banded** - tokens are assigned to five color bands from most common to least common within the `topX` window.

SubMiner looks up each token's `frequencyRank` from `term_meta_bank_*.json` files. Only tokens with a positive rank at or below `topX` are highlighted.

**Key settings:**

| Option                                           | Default      | Description                                                      |
| ------------------------------------------------ | ------------ | ---------------------------------------------------------------- |
| `subtitleStyle.frequencyDictionary.enabled`      | `false`      | Enable frequency highlighting                                    |
| `subtitleStyle.frequencyDictionary.topX`         | `10000`      | Max frequency rank to highlight                                  |
| `subtitleStyle.frequencyDictionary.mode`         | `"single"`   | `"single"` or `"banded"`                                         |
| `subtitleStyle.frequencyDictionary.matchMode`    | `"headword"` | `"headword"` or `"surface"`                                      |
| `subtitleStyle.frequencyDictionary.singleColor`  | `#f5a97f`    | Color for single mode                                            |
| `subtitleStyle.frequencyDictionary.bandedColors` | 5 colors[^1] | Array of five hex colors for banded mode                         |
| `subtitleStyle.frequencyDictionary.sourcePath`   | `""`         | Custom path to frequency dictionary root (empty = auto-discover) |

[^1]: Default banded palette (most common → least common): `#ed8796`, `#f5a97f`, `#f9e2af`, `#8bd5ca`, `#8aadf4`.

When `sourcePath` is omitted, SubMiner searches default install/runtime locations for `frequency-dictionary` directories automatically.

::: info
Frequency highlighting skips tokens that look like non-lexical noise (kana reduplication, short kana endings like `っ`), even when dictionary ranks exist. For merged kana tokens, SubMiner keeps a rank when the dictionary headword reading covers the full token (for example, `かと言って` / `かといって`), while grammar wrapped around a shorter lemma remains unannotated.
:::

::: info
Frequency, JLPT, and N+1 metadata are only shown for tokens that survive the subtitle-annotation noise filter. Standalone grammar tokens like `は`, `です`, and `この` are intentionally left unannotated even if a dictionary can assign them metadata.
:::

## JLPT Tagging

JLPT tagging adds colored underlines to tokens based on their JLPT level (N1–N5), giving you an at-a-glance sense of difficulty distribution in each subtitle line.

**How it works:**

SubMiner loads offline `term_meta_bank_*.json` files from `vendor/yomitan-jlpt-vocab` and matches each token's headword against the bank entries. Tokens with a recognized JLPT level receive a colored underline.

**Default colors:**

| Level | Color     | Preview |
| ----- | --------- | ------- |
| N1    | `#ed8796` | Red     |
| N2    | `#f5a97f` | Peach   |
| N3    | `#f9e2af` | Yellow  |
| N4    | `#8bd5ca` | Teal    |
| N5    | `#8aadf4` | Blue    |

All colors are customizable via the `subtitleStyle.jlptColors` object.

**Key settings:**

| Option                             | Default   | Description                   |
| ---------------------------------- | --------- | ----------------------------- |
| `subtitleStyle.enableJlpt`         | `false`   | Enable JLPT underline styling |
| `subtitleStyle.jlptColors.N1`–`N5` | see above | Per-level underline colors    |

## Runtime Toggles

These annotation layers can be toggled at runtime via the runtime options palette (`Ctrl/Cmd+Shift+O`) without restarting:

- `ankiConnect.knownWords.highlightEnabled` (`On` / `Off`)
- `ankiConnect.knownWords.matchMode`
- `ankiConnect.nPlusOne.enabled` (`On` / `Off`)
- `subtitleStyle.enableJlpt` (`On` / `Off`)
- `subtitleStyle.frequencyDictionary.enabled` (`On` / `Off`)

(Character-name matching, `subtitleStyle.nameMatchEnabled`, is toggled through config or the Settings window, not the runtime palette.)

Toggles only apply to new subtitle lines after the change - the currently displayed line is not re-tokenized in place.

## Rendering Priority

When multiple annotations apply to the same token, the visual priority is:

1. **Character-name match** (highest) - dictionary-driven character-name token styling; it clears the token's N+1, frequency, and JLPT annotations
2. **N+1 target** - the single unknown word in an N+1 sentence
3. **Known-word color** - already-learned token tint
4. **Frequency highlight** - common-word coloring (not applied when a higher layer already matched)
5. **JLPT underline** - level-based underline (stacks with N+1/known/frequency since it uses underline rather than text color, but not with a character-name match)
