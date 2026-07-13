# TsukiHime Integration

[TsukiHime](https://tsukihime.org) tracks anime torrent releases and extracts every attachment - including embedded subtitle tracks - from the release files, hosting them for direct download. SubMiner integrates with the TsukiHime API so you can pull English subtitles for the currently playing episode straight from the overlay, no torrent client involved. Downloaded subtitles are decompressed, saved next to the video, and loaded into mpv immediately.

This is the English-side companion to the [Jimaku integration](/jimaku-integration): Jimaku covers Japanese subtitles, TsukiHime covers your secondary language (English by default). Releases that ship multiple languages (e.g. Netflix `[MultiSub]` rips) expose them all; the modal's tabs pick which ones you see, and each download is saved with its own language suffix.

::: tip Successor to Animetosho
TsukiHime replaces [Animetosho](https://animetosho.org), which stops processing new releases in May 2026. TsukiHime imported the Animetosho index and mirrors its attachment storage, so older releases stay reachable alongside new ones.
:::

::: tip No API key required
Unlike Jimaku, TsukiHime needs no account or API key. The only requirement is the `xz` binary on your `PATH` - TsukiHime serves extracted subtitles xz-compressed, and SubMiner shells out to `xz` to decompress them. Most Linux distributions ship it by default (package `xz` or `xz-utils`).
:::

## How It Works

The integration runs through an in-overlay modal opened with `Ctrl+Shift+T` by default. The modal has two tabs that filter the subtitle tracks of the selected release by language: the first tab follows your `secondarySub.secondarySubLanguages` config (defaults to English when unset, and the tab is labeled accordingly - e.g. "German" if you configure `["de"]`), and the second tab is always **Japanese**. Tracks with no language tag stay visible on the first tab.

When you open the modal, SubMiner parses the current video filename to extract a title and episode number (same parser as Jimaku - `S01E03`, `1x03`, `E03`, and dash-separated numbers all work). If the filename yields a high-confidence match, SubMiner auto-searches immediately.

From there:

1. **Search** - SubMiner queries TsukiHime with `<title> <episode>`. Results appear as a list of releases (e.g. `[SubsPlease] ... - 28 (1080p)`), each showing size, file count, and the subtitle languages the release carries.
2. **Browse releases** - Select a release to list the text subtitle tracks extracted from its files. English tracks sort first; image-based tracks (PGS/VobSub) are filtered out.
3. **Download** - Selecting a track downloads the xz-compressed subtitle from TsukiHime's storage, decompresses it, saves it next to the video (or a temp directory for remote/streamed media), and loads it into mpv. Japanese tracks are selected as the **primary** subtitle; any other language loads as the **secondary** subtitle, so your Japanese primary track stays in place. The filename carries the track's language - `<video basename>.en.<ext>` for English, `.ja` for Japanese, and so on - so mpv and media servers detect the language correctly.

Because releases on TsukiHime are the same files circulating as torrents, picking the release that matches your local file (same group, same version) gives you subtitles with exact timing - no resync needed. If your file is a raw or from a different group, pick any release of the same episode and adjust timing with the [subtitle sync tools](/troubleshooting#subtitle-sync-subsync) (`Ctrl+Alt+S`) if necessary.

### Modal Keyboard Shortcuts

| Key                          | Action                          |
| ---------------------------- | ------------------------------- |
| `Enter` (in text field)      | Search                          |
| `Enter` (in list)            | Select release / download track |
| `Arrow Up` / `Arrow Down`    | Navigate releases or tracks     |
| `Arrow Left` / `Arrow Right` | Switch English / Japanese tab   |
| `Escape`                     | Close modal                     |

## Configuration

The integration works out of the box. An optional `tsukihime` section in `config.jsonc` tunes it:

```jsonc
{
  "tsukihime": {
    "apiBaseUrl": "https://api.tsukihime.org/v1",
    "maxSearchResults": 10,
  },
}
```

| Option                       | Type     | Default                          | Description                                                                |
| ---------------------------- | -------- | -------------------------------- | -------------------------------------------------------------------------- |
| `tsukihime.apiBaseUrl`       | `string` | `"https://api.tsukihime.org/v1"` | Base URL of the TsukiHime API. Only change this if using a mirror.         |
| `tsukihime.maxSearchResults` | `number` | `10`                             | Maximum number of releases returned per search (the API caps this at 100). |

The keyboard shortcut is configured separately under `shortcuts`:

```jsonc
{
  "shortcuts": {
    "openTsukihime": "Ctrl+Shift+T", // default; set to null to disable
  },
}
```

Existing Animetosho configuration remains compatible. SubMiner treats the old `animetosho` section and `shortcuts.openAnimetosho` setting as deprecated aliases. When old and current names are both present, `tsukihime` and `shortcuts.openTsukihime` take precedence.

## Other Ways to Open It

- CLI: `subminer --open-tsukihime`
- Keybinding command: bind any key to `["__tsukihime-open"]` in the `keybindings` array

The previous `--open-animetosho` flag and `__animetosho-open` keybinding command remain accepted as deprecated aliases.

## Troubleshooting

- **"xz binary not found"** - install `xz`/`xz-utils` with your package manager.
- **"Batch releases are not supported"** - TsukiHime only exposes extracted attachments for single-file torrents. Pick the single-episode release for your episode instead of a season batch.
- **"No text subtitle tracks in this release"** - the release only carries image-based subtitles (PGS/VobSub) or none at all; try a different release (fansub and SubsPlease-style releases almost always carry ASS tracks).
- **Timing is off** - the subtitle came from a different release than your video file. Use the subtitle sync modal (`Ctrl+Alt+S`) or pick the release matching your file exactly.
