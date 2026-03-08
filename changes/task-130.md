type: fixed
area: launcher

- Keep the background SubMiner process running after a launcher-managed mpv session exits so the next mpv instance can reconnect without restarting the app.
- Reuse prior tokenization readiness after the background app is already warm so reopening a video does not pause again waiting for duplicate warmup completion.
