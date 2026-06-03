# Jimaku Integration

[Jimaku](https://jimaku.cc) is a community-driven subtitle repository for anime - a shared online library of subtitle files contributed by other learners. SubMiner integrates with the Jimaku API so you can search, browse, and download Japanese subtitle files directly from the overlay - no alt-tabbing or manual file management required. Downloaded subtitles are loaded into mpv immediately.

::: tip Prerequisite: a free API key
You need a Jimaku account and an API key (a personal access string) before this feature works. Create an account at [jimaku.cc](https://jimaku.cc), copy your key, and add it to your config as shown under [Configuration](#configuration) below. Without a key, the search modal will report "Jimaku API key not set."
:::

## How It Works

The Jimaku integration runs through an in-overlay modal accessible via a keyboard shortcut (`Ctrl+Shift+J` by default).

When you open the modal, SubMiner parses the current video filename to extract a title, season, and episode number. Common naming conventions are supported - `S01E03`, `1x03`, `E03`, and dash-separated episode numbers all work. If the filename yields a high-confidence match (title + episode), SubMiner auto-searches immediately.

From there:

1. **Search** - SubMiner queries the Jimaku API with the parsed title. Results appear as a list of anime entries (Japanese and English names).
2. **Browse entries** - Select an entry to load its available subtitle files, filtered by episode if one was detected.
3. **Browse files** - Files show name, size, and last-modified date. If a language preference is configured, files are sorted accordingly (e.g., Japanese-tagged files first).
4. **Download** - Selecting a file downloads it to the same directory as the video (or a temp directory for remote/streamed media) and loads it into mpv as a new subtitle track.

If no files match the current episode filter, a "Show all files" button lets you broaden the search to all episodes for that entry.

### Modal Keyboard Shortcuts

| Key | Action |
| --- | --- |
| `Enter` (in text field) | Search |
| `Enter` (in list) | Select entry / download file |
| `Arrow Up` / `Arrow Down` | Navigate entries or files |
| `Escape` | Close modal |

## Configuration

Add a `jimaku` section to your `config.jsonc`:

```jsonc
{
  "jimaku": {
    "apiKey": "YOUR_API_KEY",
    "apiKeyCommand": "cat ~/.jimaku_key",
    "apiBaseUrl": "https://jimaku.cc",
    "languagePreference": "ja",
    "maxEntryResults": 10
  }
}
```

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `jimaku.apiKey` | `string` | - | Jimaku API key (plaintext). Mutually exclusive with `apiKeyCommand`. |
| `jimaku.apiKeyCommand` | `string` | - | Shell command that prints the API key to stdout. Useful for secret managers (e.g., `pass jimaku/api-key`). |
| `jimaku.apiBaseUrl` | `string` | `"https://jimaku.cc"` | Base URL for the Jimaku API. Only change this if using a mirror or local instance. |
| `jimaku.languagePreference` | `"ja"` \| `"en"` \| `"none"` | `"ja"` | Sort subtitle files by language tag. `"ja"` pushes Japanese-tagged files to the top; `"en"` does the same for English. `"none"` preserves the API order. |
| `jimaku.maxEntryResults` | `number` | `10` | Maximum number of anime entries returned per search. |

The keyboard shortcut is configured separately under `shortcuts`:

```jsonc
{
  "shortcuts": {
    "openJimaku": "Ctrl+Shift+J"
  }
}
```

### API Key

An API key is required to use the Jimaku integration. You can get one from [jimaku.cc](https://jimaku.cc). There are two ways to provide it:

- **`apiKey`** - set the key directly in config. Simple, but the key is stored in plaintext.
- **`apiKeyCommand`** - a shell command that outputs the key. Runs with a 10-second timeout. Preferred if you use a secret manager like `pass`, `gpg`, or a keychain tool.

If both are set, `apiKey` takes priority.

## Filename Parsing

SubMiner extracts media info from the current video path to pre-fill the search fields. The parser handles:

- **Season + episode patterns:** `S01E03`, `1x03`
- **Episode-only patterns:** `E03`, `EP03`, or dash-separated numbers like `Title - 03 -`
- **Bracket tags:** `[SubGroup]`, `[1080p]`, `[HEVC]` - stripped before title extraction
- **Year tags:** `(2024)` - stripped
- **Dots and underscores:** treated as spaces
- **Remote/streamed URLs:** SubMiner checks URL query parameters (`title`, `name`, `q`) and path segments to extract a meaningful title

If the parser produces a high-confidence result (title + episode both detected), the search runs automatically when the modal opens. Otherwise, you can adjust the fields manually before searching.

## Troubleshooting

**"Jimaku API key not set"**

Configure `jimaku.apiKey` or `jimaku.apiKeyCommand` in your config. If using `apiKeyCommand`, verify the command works in your shell: it should print the key and exit cleanly.

**"Jimaku request failed" or HTTP 429**

The Jimaku API has rate limits. If you see 429 errors, wait for the retry duration shown in the OSD message and try again. An API key provides higher rate limits.

**No entries found**

Try simplifying the title - remove season/episode qualifiers and search with just the anime name. Jimaku's search matches against its own database of anime titles, so the exact spelling matters.

**No files found for this episode**

The entry may not have per-episode files, or files may be named differently. Click "Show all files" to see everything available for the entry.

**Downloaded subtitle not loading**

Verify mpv is running and connected via IPC. SubMiner loads the subtitle by issuing a `sub-add` command over the mpv socket. If mpv is not connected, the download succeeds but the subtitle cannot be loaded.

## Related

- [Configuration Reference](/configuration#jimaku) - full config options
- [Mining Workflow](/mining-workflow#jimaku-subtitle-search) - how Jimaku fits into the sentence mining loop
- [Troubleshooting](/troubleshooting#jimaku) - additional error guidance
