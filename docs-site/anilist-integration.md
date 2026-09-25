# AniList integration

SubMiner updates your [AniList](https://anilist.co) watch progress when you finish an episode. The same connection supplies cover art for the stats dashboard and names for the [character dictionary](/character-dictionary).

## Setup

1. Set `anilist.enabled` to `true`:

   ```jsonc
   {
     "anilist": {
       "enabled": true,
     },
   }
   ```

2. Restart SubMiner. With no token stored, it opens the AniList setup window. You can also open it from the tray (**Configure AniList**) or with `subminer app --anilist-setup`.
3. Approve access on the AniList page. SubMiner receives the token through a `subminer://` link and stores it encrypted.

If the setup window does not render, SubMiner opens the authorization page in your browser instead. To skip the flow entirely, paste a token into `anilist.accessToken`.

On Linux, the token is stored with `gnome-libsecret` by default. If your keyring is unavailable, start it (gnome-keyring or KWallet) or launch SubMiner with `--password-store=basic_text`.

## How updates work

An episode counts as watched after 85% of its length and at least 10 minutes of playback. SubMiner then:

1. Reads the title, season, and episode from the file name and folder. Install [guessit](https://github.com/guessit-io/guessit) for better parsing. A folder named `Season 2` is a strong season hint.
2. Finds the matching AniList entry. For season 2 and later, it follows the show's sequels.
3. Sets your progress to that episode and marks the entry Watching, or Completed on the final episode.

The show must already be on your Planning or Watching list. SubMiner does not add new entries, and it never lowers your progress.

Failed updates are saved and retried in the background, up to 8 times with growing delays. The queue survives restarts.

## Fixing a wrong match

If a cover or title in the stats Library is wrong, open the title and use **Change AniList Entry**.

If SubMiner cannot find a later season, it skips the update rather than writing progress to season 1. Pin the right entry with the character dictionary's AniList override. See [Character dictionary](/character-dictionary).

## Commands

| Command                              | What it does                            |
| ------------------------------------ | --------------------------------------- |
| `subminer app --anilist-setup`       | Open the AniList setup window           |
| `subminer app --anilist-status`      | Show token state and retry queue counts |
| `subminer app --anilist-logout`      | Remove the stored token                 |
| `subminer app --anilist-retry-queue` | Retry one queued update now             |

## Options

| Key                   | What it does                                                  |
| --------------------- | ------------------------------------------------------------- |
| `anilist.enabled`     | Turns on progress updates.                                    |
| `anilist.accessToken` | Token override. Leave empty to use the token stored by setup. |

Character dictionary settings live under `anilist.characterDictionary` and are covered on the [Character dictionary](/character-dictionary) page. See [Configuration](/configuration#anilist) for defaults.

## Troubleshooting

**No update after an episode.** Check that `anilist.enabled` is `true` and that you watched at least 85% of the episode.

**"AniList update not possible."** Add the show to your Planning or Watching list, then mark the episode watched again.

**Wrong show or episode.** Install guessit and make sure it is on your `PATH`. Unusual file names parse poorly without it.

**Token errors.** Run `subminer app --anilist-status`. If the token is invalid, run `--anilist-logout`, then `--anilist-setup`.
