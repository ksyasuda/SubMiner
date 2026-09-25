# Jellyfin integration

If your anime lives on a [Jellyfin](https://jellyfin.org) server, SubMiner can appear as a cast target in any Jellyfin client. Cast an episode and it plays in SubMiner's mpv with the overlay and Yomitan lookup attached.

## Setup

You need a Jellyfin server (Jellyfin 12 is supported) and your username and password.

1. Start SubMiner and leave it in the system tray.
2. Open the tray menu and click **Configure Jellyfin**. You can also run `subminer jellyfin`.
3. Enter the **Server URL** (for example `http://127.0.0.1:8096`), **Username**, and **Password**, then click **Login**.

SubMiner stores an encrypted session token, not your password, and turns the integration on. Reopen the same window to switch servers or log out.

## Casting from Jellyfin

After you sign in, SubMiner connects to Jellyfin at startup and shows up in the cast ("Play on") menu under your computer's hostname. To connect for the current session only, tick **Jellyfin Discovery** in the tray menu.

1. In the Jellyfin web or mobile app, start playing an episode.
2. Open the cast menu and pick your computer.

SubMiner starts mpv if it is not already running. Pause, seek, stop, and track changes in the Jellyfin app are mirrored in mpv, and watch progress syncs back to Jellyfin. Playback resumes from Jellyfin's saved position.

SubMiner selects a Japanese subtitle track automatically and resets mpv's subtitle delay to zero. It direct-plays files when it can and asks Jellyfin to transcode the rest.

On Windows, casting finds mpv through `mpv.executablePath`, then `SUBMINER_MPV_PATH`, then `PATH`. An invalid `mpv.executablePath` stops mpv from starting.

## Playing from the terminal

The launcher can browse your libraries and play an item without a Jellyfin client:

```bash
subminer jellyfin -p       # fzf picker; `jf` is an alias for `jellyfin`
subminer -R jellyfin -p    # rofi picker
```

Sign in first. See [Launcher script](/launcher-script) for the other `jellyfin` subcommands.

## Options

All options are under **Settings > Integrations > Jellyfin**, or `jellyfin` in `config.jsonc`. See [Configuration](/configuration#jellyfin) for the full list and defaults.

| Key                        | What it does                                                         |
| -------------------------- | -------------------------------------------------------------------- |
| `enabled`                  | Turns the integration on. Set for you when you sign in.              |
| `serverUrl`                | Your Jellyfin server. Filled in when you sign in.                    |
| `remoteControlEnabled`     | Lets SubMiner act as a cast target.                                  |
| `remoteControlAutoConnect` | Connects at startup. Turn off to start discovery from the tray.      |
| `autoAnnounce`             | Re-announces the device on connect. Try it if SubMiner appears late. |
| `transcodeVideoCodec`      | Video codec requested when Jellyfin transcodes.                      |

For headless setups, `SUBMINER_JELLYFIN_ACCESS_TOKEN` and `SUBMINER_JELLYFIN_USER_ID` supply a session without the sign-in window. Treat the token store and `config.jsonc` as secrets.

## Troubleshooting

**SubMiner is missing from the cast menu.** Check that SubMiner is running, that you are signed in (log in again if the token expired), and that discovery is on. The Jellyfin client and SubMiner must use the same server.

**Casting starts but nothing plays.** Confirm the item plays in another Jellyfin client. If mpv was closed, give SubMiner a few seconds to start it.

**Linux token storage fails.** SubMiner stores the token with `gnome-libsecret` by default. Start your keyring, or pass `--password-store=basic_text`.
