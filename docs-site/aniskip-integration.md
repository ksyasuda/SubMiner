# AniSkip integration

SubMiner looks up opening timestamps on [AniSkip](https://aniskip.com) so you can skip an anime's intro with one key.

## Setup

AniSkip is on by default. To turn it off or change the key:

```jsonc
{
  "mpv": {
    "aniskipEnabled": true,
    "aniskipButtonKey": "TAB",
  },
}
```

Both settings apply immediately, without restarting mpv. For better title and episode detection, install [guessit](https://github.com/guessit-io/guessit):

```bash
python3 -m pip install --user guessit
```

## Usage

When a local file loads, SubMiner reads the title and episode from the file name, finds the show on MyAnimeList, and asks AniSkip for the intro's timestamps. Streams and URLs are skipped.

If AniSkip has an intro, SubMiner adds `AniSkip Intro Start` and `AniSkip Intro End` chapters. When the intro starts, mpv shows "You can skip by pressing TAB" (with your key) for 3 seconds. Press the key any time during the intro to jump to its end.

With a custom key other than `TAB` or `y-k`, `y-k` also skips.

## Triggering from mpv

| Command                                   | What it does                                            |
| ----------------------------------------- | ------------------------------------------------------- |
| `script-message subminer-skip-intro`      | Skip to the end of the intro                            |
| `script-message subminer-aniskip-refresh` | Look up the current file again, ignoring cached results |

Use `subminer-aniskip-refresh` after a lookup failed or matched the wrong show.
