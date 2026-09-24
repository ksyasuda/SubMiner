# YouTube integration

Play a YouTube URL and SubMiner downloads its Japanese subtitles and loads them into mpv, so you can mine from it like a local file.

## Setup

Install [yt-dlp](https://github.com/yt-dlp/yt-dlp) and make sure it is on your `PATH`. If it is somewhere else, set `SUBMINER_YTDLP_BIN` to the full path of the binary.

## Usage

```bash
subminer https://www.youtube.com/watch?v=VIDEO_ID
subminer ytsearch:"keyword"     # plays the first search result
```

mpv starts paused while SubMiner fetches the subtitle list. It picks a primary and a secondary track, loads them, and resumes playback once the primary track is ready. A playlist link plays only the linked video.

SubMiner picks tracks in this order. Manual (uploaded) tracks win over auto-generated captions.

| Track     | Choice                                                                           |
| --------- | -------------------------------------------------------------------------------- |
| Primary   | Japanese manual, then Japanese auto, then any manual track, then the first track |
| Secondary | English manual, then English auto. Skipped if none exists.                       |

Press `Ctrl+Alt+C` during playback to open the subtitle picker. It lists every track with its language and kind, and lets you choose different primary and secondary tracks or retry a failed load.

## Secondary subtitle languages

YouTube secondary selection is fixed to English. `secondarySub.secondarySubLanguages` and `secondarySub.autoLoadSecondarySub` apply only to local files and Jellyfin. `secondarySub.defaultMode` still controls how the secondary bar is shown. Use the picker to load a different secondary language.

Likewise, `youtube.primarySubLanguages` does not change which YouTube track is picked. It sets which languages count as a primary subtitle for local and playlist subtitle selection and for the "primary subtitle missing" notification.

## Card media

By default, card audio and screenshots are cut from mpv's live YouTube stream. If card media fails with `403` errors, switch to the background cache:

```jsonc
{
  "youtube": {
    "mediaCache": { "mode": "background" },
  },
}
```

In background mode, SubMiner downloads the video with yt-dlp after playback starts. Cards you mine get their text fields right away, and audio and images are added once the download finishes. `youtube.mediaCache.maxHeight` caps the download resolution (`0` for no limit). If the download fails, SubMiner tells you and drops the pending media updates.

See [Configuration](/configuration#youtube-playback-settings) for all `youtube` options and defaults.

## Troubleshooting

**No Japanese subtitles.** The video may not have any. Open the picker with `Ctrl+Alt+C` to see what is available.

**yt-dlp not found.** Install it and put it on `PATH`, or set `SUBMINER_YTDLP_BIN`.

**Timeouts.** Each yt-dlp call times out after 15 seconds. Slow or rate-limited connections can hit this. Retry, or update yt-dlp.

**Poor subtitle quality.** Auto-generated captions are often inaccurate. SubMiner uses a manual track when one exists.

A missing or failed secondary track never blocks playback.

## Stats

The stats Library groups YouTube videos by channel. Choose **YouTube** in the Library filter to see them. See [Immersion tracking](/immersion-tracking).
