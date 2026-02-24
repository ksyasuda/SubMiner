# Jellyfin Integration

SubMiner includes an optional Jellyfin CLI integration for:

- authenticating against a server
- listing libraries and media items
- launching item playback in the connected mpv instance
- receiving Jellyfin remote cast-to-device playback events in-app
- opening an in-app setup window for server/user/password input

## Requirements

- Jellyfin server URL and user credentials
- For `--jellyfin-play`: connected mpv IPC socket (`--start` or existing mpv plugin workflow)
- On Linux, token encryption defaults to `gnome-libsecret`; pass `--password-store=<backend>` to override.

## Setup

1. Set base config values (`config.jsonc`):

```jsonc
{
  "jellyfin": {
    "enabled": true,
    "serverUrl": "http://127.0.0.1:8096",
    "username": "your-user",
    "remoteControlEnabled": true,
    "remoteControlAutoConnect": true,
    "autoAnnounce": false,
    "remoteControlDeviceName": "SubMiner",
    "defaultLibraryId": "",
    "pullPictures": false,
    "iconCacheDir": "/tmp/subminer-jellyfin-icons",
    "directPlayPreferred": true,
    "directPlayContainers": ["mkv", "mp4", "webm", "mov", "flac", "mp3", "aac"],
    "transcodeVideoCodec": "h264",
  },
}
```

2. Authenticate:

```bash
subminer jellyfin
subminer jellyfin -l \
  --server http://127.0.0.1:8096 \
  --username your-user \
  --password 'your-password'
```

3. List libraries:

```bash
SubMiner.AppImage --jellyfin-libraries
```

Launcher wrapper equivalent for interactive playback flow:

```bash
subminer jellyfin -p
```

Launcher wrapper for Jellyfin cast discovery mode (foreground app process):

```bash
subminer jellyfin -d
```

`subminer jf ...` is an alias for `subminer jellyfin ...`.

To clear saved session credentials:

```bash
subminer jellyfin --logout
```

4. List items in a library:

```bash
SubMiner.AppImage --jellyfin-items --jellyfin-library-id LIBRARY_ID --jellyfin-search term
```

5. Start playback:

```bash
SubMiner.AppImage --start
SubMiner.AppImage --jellyfin-play --jellyfin-item-id ITEM_ID
```

Optional stream overrides:

- `--jellyfin-audio-stream-index N`
- `--jellyfin-subtitle-stream-index N`

## Playback Behavior

- Direct play is attempted first when:
  - `jellyfin.directPlayPreferred=true`
  - media source supports direct stream
  - source container matches `jellyfin.directPlayContainers`
- If direct play is not selected/available, SubMiner requests a Jellyfin transcoded stream (`master.m3u8`) using `jellyfin.transcodeVideoCodec`.
- Resume position (`PlaybackPositionTicks`) is applied via mpv seek.
- Media title is set in mpv as `[Jellyfin/<mode>] <title>`.

## Cast To Device Mode (jellyfin-mpv-shim style)

When SubMiner is running with a valid Jellyfin session, it can appear as a
remote playback target in Jellyfin's cast-to-device menu.

### Requirements

- `jellyfin.enabled=true`
- valid `jellyfin.serverUrl` and Jellyfin auth session (env override or stored login session)
- `jellyfin.remoteControlEnabled=true` (default)
- `jellyfin.remoteControlAutoConnect=true` (default)
- `jellyfin.autoAnnounce=false` by default (`true` enables auto announce/visibility check logs on connect)

### Behavior

- SubMiner connects to Jellyfin remote websocket and posts playback capabilities.
- `Play` events open media in mpv with the same defaults used by `--jellyfin-play`.
- If mpv IPC is not connected at cast time, SubMiner auto-launches mpv in idle mode with SubMiner defaults and retries playback.
- `Playstate` events map to mpv pause/resume/seek/stop controls.
- Stream selection commands (`SetAudioStreamIndex`, `SetSubtitleStreamIndex`) are mapped to mpv track selection.
- SubMiner reports start/progress/stop timeline updates back to Jellyfin so now-playing and resume state stay synchronized.
- `--jellyfin-remote-announce` forces an immediate capability re-broadcast and logs whether server sessions can see the device.

### Troubleshooting

- Device not visible in Jellyfin cast menu:
  - ensure SubMiner is running
  - ensure session token is valid (`--jellyfin-login` again if needed)
  - ensure `remoteControlEnabled` and `remoteControlAutoConnect` are true
- Cast command received but playback does not start:
  - verify mpv IPC can connect (`--start` flow)
  - verify item is playable from normal `--jellyfin-play --jellyfin-item-id ...`
- Frequent reconnects:
  - check Jellyfin server/network stability and token expiration

## Failure Handling

User-visible errors are shown through CLI logs and mpv OSD for:

- invalid credentials
- expired/invalid token
- server/network errors
- missing library/item identifiers
- no playable source
- mpv not connected for playback

## Security Notes and Limitations

- Jellyfin auth session (`accessToken` + `userId`) is stored in local encrypted token storage after login/setup.
- Launcher wrappers support `--password-store=<backend>` and forward it through to the app process.
- Optional environment overrides are supported: `SUBMINER_JELLYFIN_ACCESS_TOKEN` and `SUBMINER_JELLYFIN_USER_ID`.
- Treat both token storage and config files as secrets and avoid committing them.
- Password is used only for login and is not stored.
- Optional setup UI is available via `--jellyfin`; all actions are also available via CLI flags.
- `subminer` wrapper uses Jellyfin subcommands (`subminer jellyfin ...`, alias `subminer jf ...`). Use `SubMiner.AppImage` for direct `--jellyfin-libraries` and `--jellyfin-items`.
- For direct app CLI usage (`SubMiner.AppImage ...`), `--jellyfin-server` can override server URL for login/play flows without editing config.
