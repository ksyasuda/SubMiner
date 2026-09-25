# WebSocket and texthooker API

SubMiner streams the current subtitle over local WebSockets and serves a texthooker page, so browser tools and your own scripts can follow along. This page is the reference for building a client. If you only want subtitles in a browser tab for Yomitan, see [Texthooker page](#texthooker-integration-guide).

| Surface               | Default                     | Purpose                                               |
| --------------------- | --------------------------- | ----------------------------------------------------- |
| `websocket`           | `ws://127.0.0.1:6677`       | Plain subtitle text                                   |
| `annotationWebsocket` | `ws://127.0.0.1:6678`       | Subtitle text plus token metadata and rendered HTML   |
| `texthooker`          | `http://127.0.0.1:5174`     | Bundled texthooker page, preconfigured for your setup |
| mpv plugin            | `script-message subminer-*` | Start, stop, toggle, and status automation inside mpv |

All servers bind to `127.0.0.1` only. There is no authentication.

## Enable the services

All three services are off by default. Turn on the ones you need in `config.jsonc`:

```jsonc
{
  "websocket": {
    "enabled": "auto",
    "port": 6677,
  },
  "annotationWebsocket": {
    "enabled": true,
    "port": 6678,
  },
  "texthooker": {
    "launchAtStartup": true,
    "openBrowser": false,
  },
}
```

- `websocket.enabled`: `true` always starts the plain stream. `"auto"` starts it unless the external `mpv_websocket` plugin is installed at `~/.config/mpv/mpv_websocket`.
- `annotationWebsocket.enabled`: starts the annotated stream. It is independent of `websocket`.
- `texthooker.launchAtStartup`: starts the texthooker page with the app.
- `texthooker.openBrowser`: opens the page in your browser when it starts.

See [Configuration](/configuration) for all related options.

## Subtitle WebSocket

`ws://127.0.0.1:6677`. Use it when you only need the current line as text.

The server pushes only; it ignores client messages. On connect it sends the latest subtitle if there is one, then a new message each time the subtitle changes. Reconnecting is up to the client.

```json
{
  "version": 1,
  "text": "無事",
  "sentence": "無事",
  "tokens": []
}
```

| Field      | Type   | Notes                                                              |
| ---------- | ------ | ------------------------------------------------------------------ |
| `version`  | number | Payload version, currently `1`                                     |
| `text`     | string | Raw subtitle text                                                  |
| `sentence` | string | HTML-escaped text with line breaks as `<br>`, no annotation markup |
| `tokens`   | array  | Always empty on this stream                                        |

## Annotation WebSocket

`ws://127.0.0.1:6678`. The same token data the bundled texthooker uses. Prefer this stream for new clients. It keeps running when the plain stream is auto-disabled by `mpv_websocket`.

When a line is not yet tokenized, the stream first sends it with an empty `tokens` array, then sends the annotated version when tokenization finishes. Treat each message as the complete current state and replace the previous one. The plain stream does not repeat the line for this upgrade.

```json
{
  "version": 1,
  "text": "無事",
  "sentence": "<span class=\"word word-known word-jlpt-n2\" data-reading=\"ぶじ\" data-headword=\"無事\" data-frequency-rank=\"745\" data-jlpt-level=\"N2\">無事</span>",
  "tokens": [
    {
      "surface": "無事",
      "reading": "ぶじ",
      "headword": "無事",
      "startPos": 0,
      "endPos": 2,
      "partOfSpeech": "other",
      "isMerged": false,
      "isKnown": true,
      "isNPlusOneTarget": false,
      "isNameMatch": false,
      "jlptLevel": "N2",
      "frequencyRank": 745,
      "className": "word word-known word-jlpt-n2",
      "frequencyRankLabel": "745",
      "jlptLevelLabel": "N2"
    }
  ]
}
```

| Token field           | Type             | Notes                                                                          |
| --------------------- | ---------------- | ------------------------------------------------------------------------------ |
| `surface`             | string           | Display text                                                                   |
| `reading`             | string           | Kana reading when available                                                    |
| `headword`            | string           | Dictionary headword when available                                             |
| `startPos` / `endPos` | number           | Character offsets in `text`                                                    |
| `partOfSpeech`        | string           | SubMiner part-of-speech label                                                  |
| `isMerged`            | boolean          | Token was merged from several parser tokens                                    |
| `isKnown`             | boolean          | Word is known                                                                  |
| `isNPlusOneTarget`    | boolean          | Token is the line's N+1 target                                                 |
| `isNameMatch`         | boolean          | Token matched a character name                                                 |
| `frequencyRank`       | number           | Frequency rank; omitted when unavailable or a name match                       |
| `jlptLevel`           | string           | JLPT level; omitted when unavailable or a name match                           |
| `className`           | string           | CSS class list for the token                                                   |
| `frequencyRankLabel`  | string or `null` | Rank label, set only when the rank is within your frequency highlight settings |
| `jlptLevelLabel`      | string or `null` | JLPT label for display                                                         |

### HTML markup

`sentence` is HTML rendered by SubMiner. Each token is a `<span>` with these classes as they apply:

- `word` on every token
- one of `word-name-match`, `word-n-plus-one`, or `word-known`
- `word-jlpt-n1` through `word-jlpt-n5`
- `word-frequency-single`, or `word-frequency-band-1` through `word-frequency-band-5`, on words that are not known, N+1, or names

Spans also carry `data-reading`, `data-headword`, `data-frequency-rank`, and `data-jlpt-level` when available. For a fully custom UI, ignore `sentence` and render from `tokens`.

## Texthooker page {#texthooker-integration-guide}

The bundled texthooker is a browser tab that updates live with the current subtitle, works with browser Yomitan, and uses SubMiner's colors. Start it with the app (`texthooker.launchAtStartup`) or from the launcher:

```bash
subminer texthooker      # start the texthooker
subminer texthooker -o   # start it and open the browser
```

SubMiner injects the page's settings into `window.localStorage` when it serves it: the WebSocket URL (`bannou-texthooker-websocketUrl`), the known, N+1, name, frequency, and JLPT coloring toggles, and CSS custom properties for the token colors. The page connects to the annotation stream if it is enabled, otherwise to the plain stream. With neither running, it has nothing to connect to.

## Build a client

A minimal browser client for the annotation stream:

```html
<!doctype html>
<meta charset="utf-8" />
<title>SubMiner client</title>
<div id="subtitle">Waiting for subtitles...</div>
<script>
  const subtitle = document.getElementById('subtitle');
  const ws = new WebSocket('ws://127.0.0.1:6678');

  ws.addEventListener('message', (event) => {
    const payload = JSON.parse(event.data);
    subtitle.innerHTML = payload.sentence || payload.text;
  });

  ws.addEventListener('close', () => {
    subtitle.textContent = 'Connection closed; reload or reconnect.';
  });
</script>
```

A Node client:

```js
import WebSocket from 'ws';

const ws = new WebSocket('ws://127.0.0.1:6678');

ws.on('message', (raw) => {
  const payload = JSON.parse(String(raw));
  console.log({
    text: payload.text,
    tokens: payload.tokens.length,
    firstToken: payload.tokens[0]?.surface ?? null,
  });
});
```

Tips:

- Handle empty `tokens` arrays. Text can arrive before tokenization finishes.
- Reconnect on disconnect yourself.
- Use `text` for logging and automation, and `sentence` or `tokens` for display.

### Forward lines to a webhook

SubMiner does not send webhooks itself. Relay the stream to your own endpoint instead. This example forwards only lines that contain an N+1 target:

```js
import WebSocket from 'ws';

const ws = new WebSocket('ws://127.0.0.1:6678');

ws.on('message', async (raw) => {
  const payload = JSON.parse(String(raw));
  const hasNPlusOne = payload.tokens.some((token) => token.isNPlusOneTarget);
  if (!hasNPlusOne) return;

  await fetch('http://127.0.0.1:5678/subminer/n-plus-one', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ text: payload.text, tokens: payload.tokens }),
  });
});
```

The same pattern works for n8n or Zapier workflows, Discord notifiers, or appending lines to a notes file.

## mpv script messages

SubMiner has no in-app plugin SDK. Besides the streams above, you can drive it from other mpv scripts and from the [launcher CLI](/launcher-script).

The mpv plugin accepts these script messages:

| Message                 | Action                                     |
| ----------------------- | ------------------------------------------ |
| `subminer-start`        | Start the overlay                          |
| `subminer-stop`         | Stop the overlay                           |
| `subminer-toggle`       | Toggle the visible overlay                 |
| `subminer-menu`         | Open the plugin menu                       |
| `subminer-options`      | Open the SubMiner settings window          |
| `subminer-restart`      | Restart the overlay                        |
| `subminer-status`       | Show overlay status on the mpv OSD         |
| `subminer-stats-toggle` | Show an OSD hint for the overlay stats key |

`subminer-start` accepts overrides for `backend` (`auto`, `hyprland`, `sway`, `x11`, `macos`), `socket`, `texthooker`, and `log-level`:

```text
script-message subminer-start backend=hyprland socket=/custom/path texthooker=no log-level=debug
```

The plugin also registers `subminer-autoplay-ready`, `subminer-visible-overlay-shown`, `subminer-visible-overlay-hidden`, `subminer-managed-subtitles-loading`, `subminer-overlay-loading-ready`, and `subminer-reload-session-bindings`. The SubMiner app sends these to keep the plugin in sync, so do not send them from your own scripts.

While the app is connected to mpv, it also handles two AniSkip messages over the mpv IPC socket: `subminer-skip-intro` skips the intro, and `subminer-aniskip-refresh` reloads intro data, for example after your script changes title or episode metadata.

## Related pages

- [Configuration](/configuration)
- [Mining workflow](/mining-workflow)
- [mpv plugin](/mpv-plugin)
- [Launcher script](/launcher-script)
- [Anki integration](/anki-integration)
