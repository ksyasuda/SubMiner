# Jimaku integration

[Jimaku](https://jimaku.cc) is a community archive of Japanese subtitles for anime and live action. SubMiner searches it from the overlay, downloads the file you pick, and loads it into mpv.

## Setup

1. Create a free account at [jimaku.cc](https://jimaku.cc) and copy your API key.
2. Add the key to `config.jsonc`, either directly or through a command that prints it:

```jsonc
{
  "jimaku": {
    "apiKey": "YOUR_API_KEY",
    // or, to keep it out of the config file:
    // "apiKeyCommand": "pass jimaku/api-key",
  },
}
```

If both are set, `apiKey` wins. `apiKeyCommand` must print the key within 10 seconds. Without a key, the modal shows "Jimaku API key not set."

## Usage

1. Press `Ctrl+Shift+J` during playback.
2. SubMiner fills in the title, season, and episode from the file name. If it finds both a title and an episode, it searches right away. Otherwise, fix the fields and press `Enter`.
3. Pick the **Anime** or **Live action** tab. Switching tabs repeats the search.
4. Select an entry, then select a file. Files are filtered to the current episode. Click **Broaden search (all files)** to see every file in the entry.

The file is saved next to the video (or to a temp directory for streams) and loaded into mpv as a new subtitle track.

| Key              | Action                                 |
| ---------------- | -------------------------------------- |
| `Enter`          | Search, or select the highlighted item |
| `Up` / `Down`    | Move through entries or files          |
| `Left` / `Right` | Switch tabs                            |
| `Escape`         | Close                                  |

You can also open the modal with `subminer app --open-jimaku`, or change the shortcut with `shortcuts.openJimaku`.

The file name parser understands `S01E03`, `1x03`, `E03`, `EP03`, and `Title - 03 -` patterns, and reads the season from a parent folder such as `Season 2`. It ignores bracket tags like `[SubGroup]` and year tags like `(2024)`.

## Options

| Key                         | What it does                                                        |
| --------------------------- | ------------------------------------------------------------------- |
| `jimaku.apiKey`             | API key in plain text.                                              |
| `jimaku.apiKeyCommand`      | Shell command that prints the API key.                              |
| `jimaku.languagePreference` | Sorts files tagged with this language first: `ja`, `en`, or `none`. |
| `jimaku.maxEntryResults`    | Maximum entries per search.                                         |
| `jimaku.apiBaseUrl`         | API address. Change only for a mirror.                              |

See [Configuration](/configuration#jimaku) for defaults.

## Troubleshooting

**"Jimaku API key not set."** Set `jimaku.apiKey` or `jimaku.apiKeyCommand`. Run the command in your shell to confirm it prints only the key.

**HTTP 429.** You hit Jimaku's rate limit. Wait for the time shown in the message and retry.

**No entries found.** Search with just the show's name, without season or episode words. Jimaku matches against its own titles.

**The subtitle downloads but does not load.** SubMiner loads it over the mpv socket. Make sure mpv is still running and connected.
