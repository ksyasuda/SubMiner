# AniSkip integration

SubMiner finds an anime's opening and ending so you can skip either with one key. It reads the file's own chapters first and asks [AniSkip](https://aniskip.com) for anything they lack.

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

Once a file's duration is known, SubMiner looks for the opening and ending in the file's chapters:

- By name: chapters called `Opening`, `OP`, `OP ...` or `オープニング`; and `Ending`, `ED`, `ED ...`, `Outro`, `End Credits`, `Closing Credits`, `エンディング`, or `Credits` in the second half of the episode.
- By length, in files with at least three chapters: a 55 to 150 second chapter in the first 30% of the episode (the one closest to 90 seconds) is the opening, and the first one after 70% is the ending.

If the chapters don't give both, SubMiner reads the title and episode from the file name, finds the show on MyAnimeList, and asks AniSkip for the rest, sending the file's length so the times match your cut (TV or Blu-ray). AniSkip lookups run for local files only; streams such as Jellyfin use their chapters alone.

For times that came from AniSkip, SubMiner adds `AniSkip Intro Start`/`End` and `AniSkip Outro Start`/`End` chapters. When the intro or outro starts, SubMiner shows "You can skip the intro by pressing TAB" (or "the outro", with your key) for 3 seconds. Press the key any time during either one to jump to its end. Nothing is skipped automatically.

With a custom key other than `TAB` or `y-k`, `y-k` also skips.

A failed lookup (AniSkip or MyAnimeList down, rate limited, or offline) is retried the next time the file loads. A show or episode AniSkip has no times for is remembered until SubMiner restarts.

## Triggering from mpv

| Command                                   | What it does                                            |
| ----------------------------------------- | ------------------------------------------------------- |
| `script-message subminer-skip-intro`      | Skip to the end of the intro or outro that is playing   |
| `script-message subminer-aniskip-refresh` | Look up the current file again, ignoring cached results |

Use `subminer-aniskip-refresh` after a lookup failed or matched the wrong show.
