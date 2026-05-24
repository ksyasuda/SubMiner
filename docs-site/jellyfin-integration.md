# Jellyfin Integration

[Jellyfin](https://jellyfin.org) is a free, self-hosted media server — think of it as your own private streaming service for video you own. If you keep your anime on a Jellyfin server, SubMiner can play episodes through mpv with the full mining overlay.

::: tip Who needs this?
This page is only relevant if you already run (or have access to) a Jellyfin server. If you watch local files or YouTube, you can skip it. The in-app setup window (`subminer jellyfin`) is the easiest starting point.
:::

SubMiner can act as a **cast-to-device target** for Jellyfin (similar to jellyfin-mpv-shim). Sign in once, turn on discovery, and SubMiner shows up in the "Play on…" / cast menu of any Jellyfin app — web, phone, or TV. Pick an episode, cast it to SubMiner, and it plays in SubMiner's mpv window with the full overlay and Yomitan click-to-lookup.

This is the recommended way to use Jellyfin with SubMiner. A terminal-only option is covered in [Launcher playback](#launcher-playback) at the end.

## Requirements

- A Jellyfin server plus your username and password
- SubMiner installed and running (see [Installation](/installation))
- On Linux, the session token is stored with `gnome-libsecret` by default

## Quick start

### 1. Start SubMiner

Launch SubMiner so it's running in the system tray.

### 2. Sign in to your server

Open the tray menu and click **Configure Jellyfin**. In the window that opens, enter your **Server URL** (for example `http://127.0.0.1:8096`), **Username**, and **Password**, then click **Login**.

On success, SubMiner:

- saves an encrypted session token — your password is never stored,
- turns the Jellyfin integration on, and
- remembers the server and username for next time.

Reopen this window any time to switch servers or **Logout**.

### 3. Turn on discovery

Discovery is what makes SubMiner appear as a cast target. Two ways to enable it:

- **For the current session** — open the tray menu and tick **Jellyfin Discovery**. (This item appears once you've signed in.)
- **Automatically on every launch** — already on by default. After your first sign-in, SubMiner auto-connects to Jellyfin at startup, so the cast target is ready without touching the tray. You can change this under [Settings](#settings).

### 4. Cast from any Jellyfin app

In the Jellyfin web UI or mobile app, start playing something, open the **cast / "Play on"** menu, and pick your device — SubMiner appears there named after your computer's hostname. Playback opens in SubMiner.

From then on, pause / resume / seek / stop and audio or subtitle track changes you make in the Jellyfin app are mirrored in SubMiner, and your watch progress syncs back to Jellyfin (now-playing and resume position).

## What happens during playback

- **mpv launches automatically.** If mpv isn't already running when you cast, SubMiner starts it with SubMiner defaults and the bundled mpv plugin, so keybindings work right away.
- **The overlay is managed by SubMiner,** so your configured `subtitleStyle` controls how subtitles look. Use the [overlay-toggle shortcut](/shortcuts) to hide it for a session.
- **Resume works.** If Jellyfin has a saved position for the item, SubMiner seeks there on load.
- **Direct play first.** When the source allows it and the container is in your direct-play allowlist, SubMiner streams the original file; otherwise it requests a transcoded stream from Jellyfin.
- **Japanese subtitles are auto-selected,** preferring Jellyfin's default and embedded tracks over external sidecar files when several match.
- **Subtitle timing is corrected when possible.** SubMiner removes Jellyfin's server-selected subtitle stream from the mpv load URL, suppresses the mpv plugin's one-shot subtitle auto-selection and overlay auto-start for managed Jellyfin loads, stages downloaded subtitle tracks without letting mpv auto-switch between tracks, then selects the Japanese track once after applying any saved or inferred timing delay. When Jellyfin provides both Japanese and English subtitle files, SubMiner compares their cue timelines and applies a global delay if one track is clearly offset. Manual delay shifts you make with SubMiner's adjacent-cue controls are saved per item and subtitle track, then restored the next time you select that track.

## Settings

All Jellyfin options live under **Settings → Integrations → Jellyfin** (open settings from the tray's **Open SubMiner Settings**). The ones that matter for casting:

| Setting                         | Default | What it does                                                                                                          |
| ------------------------------- | ------- | --------------------------------------------------------------------------------------------------------------------- |
| **Enabled**                     | Off     | Turns the Jellyfin integration on. Switched on for you when you sign in.                                              |
| **Server Url**                  | —       | Your Jellyfin server. Filled in when you sign in.                                                                     |
| **Remote Control Enabled**      | On      | Lets SubMiner act as a cast target.                                                                                   |
| **Remote Control Auto Connect** | On      | Connects to Jellyfin at startup so discovery is automatic. Turn off if you'd rather start it from the tray each time. |
| **Auto Announce**               | Off     | Re-broadcasts visibility on connect. Enable if your device is slow to appear in the cast menu.                        |

Prefer editing the config file? The same keys live under `jellyfin` in `config.jsonc`:

```jsonc
{
  "jellyfin": {
    "enabled": true,
    "serverUrl": "http://127.0.0.1:8096",
    "remoteControlEnabled": true,
    "remoteControlAutoConnect": true,
  },
}
```

See [Configuration](/configuration) for the full list (transcode codec, direct-play containers, default library, and more).

## Troubleshooting

**SubMiner doesn't appear in the cast menu**

- Make sure SubMiner is running.
- Make sure you're signed in — reopen **Configure Jellyfin** and log in again if your token expired.
- Make sure discovery is on (tray **Jellyfin Discovery**, or **Remote Control Auto Connect** in settings).
- Make sure SubMiner and the Jellyfin client point at the same server.

**Casting starts but nothing plays**

- Confirm the item plays normally in another Jellyfin client.
- If mpv was closed, give it a moment — SubMiner launches it on demand and retries.

**SubMiner keeps disconnecting**

- Check server/network stability and whether the session token has expired.

## Security notes

- The Jellyfin session (access token + user ID) is kept in SubMiner's local encrypted token storage. Your password is used only to log in and is never saved.
- Treat the token storage and your `config.jsonc` as secrets — don't commit them.
- Advanced/headless: the `SUBMINER_JELLYFIN_ACCESS_TOKEN` and `SUBMINER_JELLYFIN_USER_ID` environment variables can supply a session without the sign-in window.

## Launcher playback

If you'd rather stay in the terminal, the `subminer` launcher can browse and play Jellyfin media directly, without casting from a Jellyfin app:

```bash
subminer jellyfin -p      # alias: subminer jf -p
```

This opens an fzf picker (add `-R` for rofi) to browse your libraries and episodes, then plays the selected item in SubMiner's mpv with the same overlay, resume, and subtitle behavior described above. Sign in first (step 2) so the launcher can reach your server. See [Launcher Script](/launcher-script) for the rest of the launcher's features.
