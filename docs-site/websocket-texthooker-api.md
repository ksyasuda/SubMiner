# WebSocket / Texthooker API & Integration

**Who this page is for:** developers and tinkerers who want to consume SubMiner's live subtitle stream from their own tools - a browser tab, an automation script, or another mpv plugin. If you just want subtitles in a browser tab for Yomitan, skip to [Texthooker Integration Guide](#texthooker-integration-guide); the rest is reference for building custom clients.

A *texthooker* is a page/tool that receives the text currently on screen so a dictionary extension (like Yomitan) can look words up. SubMiner ships its own texthooker UI and also broadcasts subtitle text over local WebSockets that any client can connect to.

SubMiner exposes a small set of local integration surfaces for browser tools, automation helpers, and mpv-driven workflows:

- **Subtitle WebSocket** at `ws://127.0.0.1:6677` by default for plain subtitle pushes.
- **Annotation WebSocket** at `ws://127.0.0.1:6678` by default for token-aware clients.
- **Texthooker HTTP UI** at `http://127.0.0.1:5174` by default for browser-based subtitle consumption.
- **mpv plugin script messages** for in-player automation and extension.

This page documents those integration points and shows how to build custom consumers around them.

## Quick Reference

| Surface | Default | Purpose |
| --- | --- | --- |
| `websocket` | `ws://127.0.0.1:6677` | Basic subtitle broadcast stream |
| `annotationWebsocket` | `ws://127.0.0.1:6678` | Structured stream with token metadata |
| `texthooker` | `http://127.0.0.1:5174` | Local texthooker UI with injected websocket config |
| mpv plugin | `script-message subminer-*` | Start/stop/toggle/status automation inside mpv |

## Enable and Configure the Services

SubMiner's integration ports are configured in `config.jsonc`. All three services are **off by default** - the block below shows the values to set to turn them on.

```jsonc
{
  "websocket": {
    "enabled": "auto",
    "port": 6677
  },
  "annotationWebsocket": {
    "enabled": true,
    "port": 6678
  },
  "texthooker": {
    "launchAtStartup": true,
    "openBrowser": false
  }
}
```

### How startup behaves

- `websocket.enabled` defaults to `false`. Set it to `"auto"` to start the basic subtitle websocket unless SubMiner detects the external `mpv_websocket` plugin, or `true` to always start it.
- `annotationWebsocket.enabled` defaults to `false` and is independent from `websocket`. Set it to `true` to start the annotated stream.
- `texthooker.launchAtStartup` defaults to `false`. Set it to `true` to start the local HTTP UI automatically.
- `texthooker.openBrowser` controls whether SubMiner opens the texthooker page in your browser when it starts.

If you use the [mpv plugin](/mpv-plugin), it can also start a texthooker-only helper process. The launcher derives the plugin's texthooker setting from your SubMiner config (`texthooker.launchAtStartup`) and injects it at runtime - there is no plugin config file to edit.

## Developer API Documentation

### 1. Subtitle WebSocket

Use the basic subtitle websocket when you only need the current subtitle line as plain text.

- **Default URL:** `ws://127.0.0.1:6677`
- **Transport:** local WebSocket server bound to `127.0.0.1`
- **Direction:** server push only
- **Client auth:** none
- **Reconnects:** client-managed

When a client connects, SubMiner immediately sends the latest subtitle payload if one is available. After that, it pushes a new message each time the current subtitle changes. Annotation-only upgrades do not repeat the same line on this basic stream.

#### Message shape

```json
{
  "version": 1,
  "text": "無事",
  "sentence": "無事",
  "tokens": []
}
```

#### Field reference

| Field | Type | Notes |
| --- | --- | --- |
| `version` | number | Current websocket payload version. Today this is `1`. |
| `text` | string | Raw subtitle text. |
| `sentence` | string | Plain subtitle text with line breaks represented as `<br>`. No annotation spans or attributes. |
| `tokens` | array | Always empty on the basic subtitle websocket. |

### 2. Annotation WebSocket

Use the annotation websocket for custom clients that want the same structured token payload the bundled texthooker UI consumes.

- **Default URL:** `ws://127.0.0.1:6678`
- **Payload shape:** JSON payload with `text`, rendered `sentence` HTML, and token metadata
- **Primary difference:** this stream is intended to stay on even when the basic websocket auto-disables because `mpv_websocket` is installed

In practice, if you are building a new client, prefer `annotationWebsocket` unless you specifically need compatibility with an existing `websocket` consumer.

On a tokenization cache miss, this stream first sends the cue as plain text with an empty `tokens` array, then sends the annotated replacement when tokenization finishes. Treat each message as the complete current state, replacing the previous payload.

#### Message shape

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

Each annotation token may include:

| Token field | Type | Notes |
| --- | --- | --- |
| `surface` | string | Display text for the token |
| `reading` | string | Kana reading when available |
| `headword` | string | Dictionary headword when available |
| `startPos` / `endPos` | number | Character offsets in the subtitle text |
| `partOfSpeech` | string | SubMiner token POS label |
| `isMerged` | boolean | Whether this token represents merged content |
| `isKnown` | boolean | Marked known by SubMiner's known-word logic |
| `isNPlusOneTarget` | boolean | True when the token is the sentence's N+1 target |
| `isNameMatch` | boolean | True for prioritized character-name matches |
| `frequencyRank` | number | Frequency rank when available |
| `jlptLevel` | string | JLPT level when available |
| `className` | string | CSS-ready class list derived from token state |
| `frequencyRankLabel` | string or `null` | Preformatted rank label for UIs |
| `jlptLevelLabel` | string or `null` | Preformatted JLPT label for UIs |

### 3. HTML markup conventions

The `sentence` field is pre-rendered HTML generated by SubMiner. Depending on token state, it can include classes such as:

- `word`
- `word-known`
- `word-n-plus-one`
- `word-name-match`
- `word-jlpt-n1` through `word-jlpt-n5`
- `word-frequency-single`
- `word-frequency-band-1` through `word-frequency-band-5`

SubMiner also adds tooltip-friendly data attributes when available:

- `data-reading`
- `data-headword`
- `data-frequency-rank`
- `data-jlpt-level`

If you need a fully custom UI, ignore `sentence` and render from `tokens` instead.

## Texthooker Integration Guide

### When to use the bundled texthooker page

Use texthooker when you want a browser tab that:

- updates live from current subtitles
- works well with browser-based Yomitan setups
- inherits SubMiner's coloring preferences and websocket URL automatically

Start it with either:

```bash
subminer texthooker
# or open the page immediately
subminer texthooker -o
```

or by leaving `texthooker.launchAtStartup` enabled.

### What SubMiner injects into the page

When SubMiner serves the local texthooker UI, it injects bootstrap values into `window.localStorage`, including:

- `bannou-texthooker-websocketUrl`
- coloring toggles for known/N+1/name/frequency/JLPT styling
- CSS custom properties for SubMiner's token colors

That means the bundled page already knows which websocket to connect to and which color palette to use.

### Build a custom websocket client

Here is a minimal browser client for the annotation stream:

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

### Build a custom Node client

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

### Integration tips

- Bind only to `127.0.0.1`; these services are local-only by design.
- Handle empty `tokens` arrays gracefully because subtitle text can arrive before tokenization completes.
- Reconnect on disconnect; SubMiner does not manage client reconnects for you.
- Prefer `payload.text` for logging/automation and `payload.sentence` or `payload.tokens` for UI rendering.

## Plugin Development

SubMiner does **not** currently expose a general-purpose third-party plugin SDK inside the app itself. Today, the supported extension surfaces are:

1. the local websocket streams
2. the local texthooker UI
3. the mpv Lua plugin's script-message API
4. the launcher CLI

### mpv script messages

The mpv plugin accepts these script messages:

```text
script-message subminer-start
script-message subminer-stop
script-message subminer-toggle
script-message subminer-menu
script-message subminer-options
script-message subminer-restart
script-message subminer-status
script-message subminer-autoplay-ready
script-message subminer-stats-toggle
script-message subminer-visible-overlay-shown
script-message subminer-visible-overlay-hidden
script-message subminer-managed-subtitles-loading
script-message subminer-overlay-loading-ready
script-message subminer-reload-session-bindings
```

The overlay/loading/session-binding messages are primarily sent by the SubMiner app to keep the plugin's state in sync. The AniSkip messages (`subminer-skip-intro`, `subminer-aniskip-refresh`) are handled by the SubMiner app over the mpv IPC socket while it is connected.

The start command also accepts inline overrides:

```text
script-message subminer-start backend=hyprland socket=/custom/path texthooker=no log-level=debug
```

### Practical extension patterns

#### Add another mpv script that coordinates with SubMiner

Examples:

- send `subminer-start` after your own media-selection script chooses a file
- send `subminer-status` before running follow-up automation
- send `subminer-aniskip-refresh` after you update title/episode metadata (handled by the SubMiner app)

#### Build a launcher wrapper

Examples:

- open a media picker, then call `subminer /path/to/file.mkv`
- launch browser-only subtitle tooling with `subminer texthooker -o`
- disable the helper UI for a session with `subminer --no-texthooker video.mkv`

#### Build an overlay-adjacent client

Examples:

- browser widget showing current subtitle + token breakdown
- local vocabulary capture helper that writes interesting lines to a file
- bridge service that forwards websocket events into your own workflow engine

## Webhook Examples

SubMiner does **not** currently send outbound webhooks by itself. The supported pattern is to consume the websocket locally and relay events into another system.

That still makes webhook-style automation straightforward.

### Example: forward subtitle lines to a local webhook receiver

```js
import WebSocket from 'ws';

const ws = new WebSocket('ws://127.0.0.1:6678');

ws.on('message', async (raw) => {
  const payload = JSON.parse(String(raw));

  await fetch('http://127.0.0.1:5678/subminer/subtitle', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      text: payload.text,
      tokens: payload.tokens,
      receivedAt: new Date().toISOString(),
    }),
  });
});
```

### Automation ideas

- **n8n / Make / Zapier relay:** send each subtitle line into an automation workflow for logging, translation, or summarization.
- **Discord / Slack notifier:** post only lines that contain unknown words or N+1 targets.
- **Obsidian / Markdown capture:** append subtitle lines plus token metadata to a daily immersion note.
- **Local LLM pipeline:** trigger a glossary, translation, or sentence-mining workflow whenever a new line arrives.

### Filtering example: only forward N+1 lines

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

## Recommended Integration Combinations

- **Browser Yomitan client:** `texthooker` + `annotationWebsocket`
- **Custom dashboard:** `annotationWebsocket` only
- **Lightweight subtitle mirror:** `websocket` only
- **mpv-side automation:** mpv plugin script messages + optional websocket relay
- **Webhook-style workflows:** `annotationWebsocket` + your own local relay service

## Related Pages

- [Configuration](/configuration#websocket-server)
- [Mining Workflow - Texthooker](/mining-workflow#texthooker)
- [MPV Plugin](/mpv-plugin)
- [Launcher Script](/launcher-script)
- [Anki Integration](/anki-integration#proxy-mode-setup-yomitan-texthooker)
