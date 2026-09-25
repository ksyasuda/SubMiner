# TsukiHime integration

[TsukiHime](https://tsukihime.org) extracts the subtitle tracks from anime torrent releases and hosts them for download. SubMiner searches it from the overlay, so you can grab Japanese or secondary-language subtitles for the current episode without a torrent client. Most releases carry only English subtitles, so [Jimaku](/jimaku-integration) is usually the better source for Japanese.

## Setup

TsukiHime needs no account or API key. SubMiner needs the `xz` binary on your `PATH` to unpack downloads. Most Linux distributions ship it (package `xz` or `xz-utils`).

## Usage

1. Press `Ctrl+Shift+T` during playback.
2. SubMiner fills in the title and episode from the file name and searches right away when it finds both. Otherwise, fix the fields and press `Enter`.
3. Pick a tab. The first tab shows your secondary language (`secondarySub.secondarySubLanguages`, or English if unset). The second tab shows Japanese. Each tab lists only releases that carry that language.
4. Select a release, then a subtitle track.

The track is saved next to the video with a language suffix, such as `<video>.ja.ass` or `<video>.en.ass` (a temp directory is used for streams). A Japanese track becomes mpv's primary subtitle. A track from the secondary tab loads as the secondary subtitle and leaves the primary alone.

Pick the release that matches your video file, same group and same version, and the timing will line up. With any other release, fix the offset with subtitle sync (`Ctrl+Alt+S`).

| Key              | Action                                 |
| ---------------- | -------------------------------------- |
| `Enter`          | Search, or select the highlighted item |
| `Up` / `Down`    | Move through releases or tracks        |
| `Left` / `Right` | Switch tabs                            |
| `Escape`         | Close                                  |

You can also open the modal with `subminer app --open-tsukihime`, bind a key to `["__tsukihime-open"]` in `keybindings`, or change the shortcut with `shortcuts.openTsukihime`.

## Options

| Key                          | What it does                                           |
| ---------------------------- | ------------------------------------------------------ |
| `tsukihime.maxSearchResults` | Maximum releases per search. The API caps this at 100. |
| `tsukihime.apiBaseUrl`       | API address. Change only for a mirror.                 |

See [Configuration](/configuration#tsukihime) for defaults.

## Troubleshooting

**"xz binary not found."** Install `xz` or `xz-utils` with your package manager.

**"No releases with Japanese subtitles."** None of the results carry a Japanese track. Try another search, or use Jimaku.

**"Batch releases are not supported."** TsukiHime only has extracted tracks for single-file torrents. Pick the single-episode release.

**"No text subtitle tracks in this release."** The release has only image-based subtitles (PGS or VobSub) or none. Try another release.

**Timing is off.** The subtitle came from a different release than your video. Use subtitle sync (`Ctrl+Alt+S`) or pick the matching release.
