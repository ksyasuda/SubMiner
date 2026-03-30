type: fixed
area: main

- Keep integrated `--start --texthooker` launches on the full app-ready startup path so the texthooker page and websocket servers start together during normal playback startup.
- Stop the mpv/plugin auto-start flow from spawning a separate standalone texthooker helper during normal `subminer <video>` launches.
