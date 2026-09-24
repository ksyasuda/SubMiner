# Character dictionary

SubMiner builds a Yomitan dictionary of the characters in the show you are watching, using data from [AniList](https://anilist.co). Character names in subtitles get their own color, and hovering one shows the character's portrait, role, voice actor, and description.

Ordinary dictionaries rarely contain character names, so without this every name counts as an unknown word and throws off [N+1 highlighting](/subtitle-annotations#n-1-word-highlighting).

## Turning it on

1. Set `subtitleStyle.nameMatchEnabled` to `true`, or turn it on in the Settings window under Annotation Display, Character Names.
2. Optionally set `subtitleStyle.nameMatchImagesEnabled` to `true` to show a small portrait next to each name in the subtitle line.
3. Play an episode.

```jsonc
{
  "subtitleStyle": {
    "nameMatchEnabled": true,
    "nameMatchImagesEnabled": true,
  },
}
```

No AniList account is needed. Logging in to AniList is only for [watch progress sync](/anilist-integration).

The character dictionary does not work when `yomitan.externalProfilePath` is set, because SubMiner then uses another app's Yomitan profile read-only.

## What happens when you play something

When a new show starts, SubMiner:

1. Guesses the title from the filename and finds it on AniList.
2. Downloads the cast list and portraits.
3. Builds the dictionary and imports it into SubMiner's Yomitan.

A notification shows each step. Once it says the dictionary is ready, names match from the next subtitle line.

Each character gets entries for the full name, family name, given name, and common honorifics (`さん`, `君`, `ちゃん`, `先生`, and others), so `太郎さん` matches as well as `太郎`.

SubMiner keeps your most recent shows loaded in one merged dictionary. `anilist.characterDictionary.maxLoaded` sets how many. Starting another show drops the oldest one. Only the current show's characters are highlighted.

### How long it takes

The first time you watch a show, most of the time goes into downloading portraits, one at a time. A typical cast takes seconds to a minute. A very large cast takes much longer: One Piece has over a thousand characters and takes around 10 minutes. After that the show is cached, and later episodes load it from disk.

## Correcting AniList matches

SubMiner can match the wrong show when a filename is ambiguous, for example `Re - ZERO, Starting Life in Another World (2016)` matching a different `Re...` series. To fix it:

1. Press `Ctrl/Cmd+D` to open the character dictionary manager.
2. Click **Override**, edit the title if needed, search, and pick the right result.

From the command line:

```bash
# List AniList matches for a file
subminer dictionary --candidates "/path/to/episode.mkv"

# Save the correct AniList ID for that series
subminer dictionary --select 21355 "/path/to/episode.mkv"
```

The override applies to every episode of that season in the same folder. Other seasons are not affected, even when they share the folder. The override also sets which entry [AniList watch progress](/anilist-integration) updates, so one fix covers both.

## Managing loaded shows

The manager (`Ctrl/Cmd+D`) lists the shows in the merged dictionary and marks the current one.

- **Remove** drops a show from the dictionary. You cannot remove the show you are watching.
- **Up/Down** changes which show gets dropped first when a new one is added.
- **Override** replaces a show's AniList match.

## Generating from the command line

```bash
subminer dictionary /path/to/media
```

This builds a standalone dictionary file for that file or folder without playing it. With the AppImage directly, use `SubMiner.AppImage --dictionary`.

## Configuration

Defaults are in the [configuration reference](/configuration).

| Key                                                                    | What it does                                     |
| ---------------------------------------------------------------------- | ------------------------------------------------ |
| `subtitleStyle.nameMatchEnabled`                                       | Build the dictionary and color character names   |
| `subtitleStyle.nameMatchImagesEnabled`                                 | Show a portrait next to matched names            |
| `subtitleStyle.nameMatchColor`                                         | Color for character names                        |
| `anilist.characterDictionary.maxLoaded`                                | Number of recent shows kept in the dictionary    |
| `anilist.characterDictionary.collapsibleSections.description`          | Show the description expanded in the popup       |
| `anilist.characterDictionary.collapsibleSections.characterInformation` | Show age, birthday, and similar details expanded |
| `anilist.characterDictionary.collapsibleSections.voicedBy`             | Show the voice actor section expanded            |
| `shortcuts.openCharacterDictionaryManager`                             | Shortcut for the manager                         |

## Troubleshooting

**It seems stuck.** Check the notification. While generating it shows counts (`image 120/400`), an estimate of time left, and an elapsed clock. If the clock moves, it is still working, and large casts are slow (see [how long it takes](#how-long-it-takes)). If you missed the notification, open the notification history with `Ctrl/Cmd+N`. For errors, check the app log ([log locations](/troubleshooting)).

**Import failed or timed out.** Yomitan may have been busy importing another dictionary. Play the next episode or restart SubMiner. The import may have finished anyway, so check for the character popup before retrying.

**Names are not highlighted.** Check that `subtitleStyle.nameMatchEnabled` is `true`, that `yomitan.externalProfilePath` is empty, and that the show was found on AniList. The wrong show's cast means a wrong match. See [correcting AniList matches](#correcting-anilist-matches).

**Portraits are missing.** Portraits need AniList to have an image and the download to succeed. If you were offline during the first sync, delete that show's file from `character-dictionaries/snapshots/` in the SubMiner config directory and replay it.

SubMiner's generator is based on the [Japanese Character Name Dictionary](https://github.com/bee-san/Japanese_Character_Name_Dictionary) project, which also supports VNDB and works without SubMiner.
