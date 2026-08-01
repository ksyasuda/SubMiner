# AniSkip Integration

SubMiner integrates with [AniSkip](https://aniskip.com) to automatically detect anime intro intervals and let you skip them with a single key press.

Intro detection runs in the SubMiner app over the mpv IPC socket. It is available whenever the overlay is connected to mpv - not just at launch - and covers every local file loaded during an mpv session, including playlist advances.

## Setup

AniSkip is enabled by default. Disable it or change the skip key in your config:

```jsonc
{
  "mpv": {
    "aniskipEnabled": true, // default: true
    "aniskipButtonKey": "TAB",
  },
}
```

Both settings hot-reload: changing them in your config takes effect immediately without restarting playback or mpv.

For best title and episode detection, install [`guessit`](https://github.com/guessit-io/guessit):

```bash
python3 -m pip install --user guessit
```

Without `guessit`, SubMiner falls back to an internal filename parser which handles most common naming conventions but may miss unusual formats.

## How It Works

On each local file load:

1. SubMiner infers the anime title, season, and episode number from the filename and path (using `guessit` if available, otherwise the built-in parser). Remote URLs are skipped entirely.
2. The title is matched against MyAnimeList to resolve a MAL id.
3. SubMiner queries the AniSkip API for an OP skip interval for that MAL id and episode.
4. If an interval is found, SubMiner adds `AniSkip Intro Start` and `AniSkip Intro End` chapter markers to the current file and binds the skip key (`mpv.aniskipButtonKey`, default `TAB`).
5. At the start of the intro, an OSD prompt appears for 3 seconds: `You can skip by pressing TAB` (reflects your configured key). Pressing the key at any point during the intro seeks to the intro end.

When a custom key (other than `TAB` or `y-k`) is configured, the legacy `y-k` chord is also bound as a fallback skip trigger.

Results are cached per file for the app session; only definitive "no intro found" results are cached, so transient lookup failures are retried on the next file load. Reload detection is also handled: if mpv reloads the same file, SubMiner re-applies the chapter markers without a new API lookup.

## Triggering from mpv

You can trigger AniSkip actions from mpv script-messages:

| Command                                   | Effect                                                                  |
| ----------------------------------------- | ----------------------------------------------------------------------- |
| `script-message subminer-skip-intro`      | Skip to the intro end immediately (same as pressing the key)            |
| `script-message subminer-aniskip-refresh` | Force a fresh lookup for the current file, discarding any cached result |

These are handled by the SubMiner app over the IPC socket.
