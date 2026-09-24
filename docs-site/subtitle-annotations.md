# Subtitle annotations

SubMiner can color and underline words in the subtitle overlay: words you already know, the one new word in an N+1 line, common words, JLPT levels, and character names. Each layer is off by default and works on its own, so turn on only the ones you want.

Yomitan splits the subtitle into words, so your installed Yomitan dictionaries and their order decide where word boundaries fall. Grammar words such as particles (`は`), auxiliaries (`です`), and endings like `んです` stay hoverable but never get annotation colors.

Defaults for every key below are in the [configuration reference](/configuration).

## Known words {#known-words}

Colors every word that already appears in your Anki decks, so you can see how much of a line you know.

Needs: Anki running with AnkiConnect, and at least one deck in `ankiConnect.knownWords.decks`.

```jsonc
{
  "ankiConnect": {
    "knownWords": {
      "highlightEnabled": true,
      "decks": { "Kaishi 1.5k": ["Word"] },
    },
  },
}
```

Map each deck to its expression or word field. SubMiner also reads the note's reading field when it has one, so a known word only matches in the reading its card teaches.

| Key                                               | What it does                                                        |
| ------------------------------------------------- | ------------------------------------------------------------------- |
| `ankiConnect.knownWords.highlightEnabled`         | Turn known-word coloring on                                         |
| `ankiConnect.knownWords.decks`                    | Deck name to list of fields to read                                 |
| `ankiConnect.knownWords.matchMode`                | `headword` matches the dictionary form, `surface` the text as shown |
| `ankiConnect.knownWords.refreshMinutes`           | How often the known-word list is re-read from Anki                  |
| `ankiConnect.knownWords.addMinedWordsImmediately` | Count a word as known as soon as you mine it                        |
| `subtitleStyle.knownWordColor`                    | Color for known words                                               |

### Known-word maturity highlighting

Colors known words by how well you know them instead of using one color. Each word gets the tier of its most mature card:

- `new`: never studied
- `learning`: in the learning or relearning queue
- `young`: in review, interval below the threshold
- `mature`: in review, interval at or above `ankiConnect.knownWords.matureThresholdDays`

Turn it on with `ankiConnect.knownWords.maturityEnabled` (known-word highlighting must also be on). Set the colors under `subtitleStyle.knownWordMaturityColors` (`new`, `learning`, `young`, `mature`).

Tiers update when the known-word list refreshes, so a card that turns mature today keeps its old color until the next refresh. If your deck has no relearning steps, lapsed cards go straight back to review and show as `young`.

## N+1 word highlighting

An N+1 line has exactly one word you don't know. It is the easiest kind of sentence to mine, because the rest of the line gives you context. SubMiner colors that one unknown word.

Needs: the same Anki setup as [known words](#known-words) (`ankiConnect.knownWords.decks`). Known-word coloring itself can stay off.

| Key                                     | What it does                            |
| --------------------------------------- | --------------------------------------- |
| `ankiConnect.nPlusOne.enabled`          | Turn N+1 highlighting on                |
| `ankiConnect.nPlusOne.minSentenceWords` | Skip lines shorter than this many words |
| `subtitleStyle.nPlusOneColor`           | Color for the unknown word              |

## Frequency highlighting

Colors words by how common they are, so a rare word in an easy line stands out.

Needs: at least one frequency dictionary installed in Yomitan. When several are installed, SubMiner uses them in your Yomitan priority order. Occurrence-count dictionaries are skipped. You can also point `sourcePath` at a folder of Yomitan-format frequency files as a fallback.

| Key                                              | What it does                                                           |
| ------------------------------------------------ | ---------------------------------------------------------------------- |
| `subtitleStyle.frequencyDictionary.enabled`      | Turn frequency highlighting on                                         |
| `subtitleStyle.frequencyDictionary.topX`         | Only color words whose rank is this number or lower (1 is most common) |
| `subtitleStyle.frequencyDictionary.mode`         | `single` uses one color, `banded` splits the range into five colors    |
| `subtitleStyle.frequencyDictionary.singleColor`  | Color for `single` mode                                                |
| `subtitleStyle.frequencyDictionary.bandedColors` | Five colors for `banded` mode, most common first                       |
| `subtitleStyle.frequencyDictionary.matchMode`    | `headword` or `surface`, as for known words                            |
| `subtitleStyle.frequencyDictionary.sourcePath`   | Optional folder of frequency files                                     |

## JLPT tagging

Underlines each word in a color for its JLPT level, N1 to N5. The JLPT word lists ship with SubMiner, so there is nothing to install.

| Key                                   | What it does                   |
| ------------------------------------- | ------------------------------ |
| `subtitleStyle.enableJlpt`            | Turn JLPT underlines on        |
| `subtitleStyle.jlptColors.N1` to `N5` | Underline color for each level |

## Character names

Colors character names from the current show and lets you hover them for a portrait, role, and voice actor.

Needs: the [character dictionary](/character-dictionary), which SubMiner builds from AniList when you turn this on.

| Key                                    | What it does                                    |
| -------------------------------------- | ----------------------------------------------- |
| `subtitleStyle.nameMatchEnabled`       | Build the character dictionary and color names  |
| `subtitleStyle.nameMatchImagesEnabled` | Show a small portrait next to each matched name |
| `subtitleStyle.nameMatchColor`         | Color for character names                       |

## Toggling during playback

Open the runtime options palette (`Ctrl/Cmd+Shift+O`) to switch these without restarting:

- known-word highlighting, maturity colors, and known-word match mode
- N+1 highlighting
- JLPT tagging
- frequency highlighting

Character names are toggled in the config file or the Settings window. Changes apply from the next subtitle line.

## When layers overlap

If one word matches several layers, the first match in this list sets its color:

1. Character name (also removes N+1, frequency, and JLPT marks)
2. N+1 target
3. Known word
4. Frequency

JLPT is an underline, so it shows alongside any of these except a character name.
