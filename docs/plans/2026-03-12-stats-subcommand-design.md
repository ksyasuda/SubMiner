# Stats Subcommand Design

**Problem:** Add a launcher command and matching app command that run only the stats dashboard stack: start the local stats server, initialize the data source it needs, and open the browser to the stats page.

**Constraints:**
- Public entrypoint is a launcher subcommand: `subminer stats`
- Reuse the existing app instance when one is already running
- Explicit `stats` launch overrides `stats.autoStartServer`
- If `immersionTracking.enabled` is `false`, fail with an error instead of opening an empty dashboard
- Scope limited to stats server + browser page; no overlay/mpv startup requirements

## Recommended Approach

Add a dedicated app CLI flag, `--stats`, and let the launcher subcommand forward into that path. Use the existing Electron single-instance flow so a second `subminer stats` invocation can be handled by the primary app instance. Add a small response-file handshake so the launcher can still return success or failure when work is delegated to an already-running primary instance.

## Runtime Flow

1. `subminer stats` runs in the launcher.
2. Launcher resolves the app binary and forwards:
   - `--stats`
   - `--log-level <level>` when provided
   - internal `--stats-response-path <tmpfile>`
3. Electron startup parses `--stats` as an app-starting command.
4. If this process becomes the primary instance, it runs a stats-only startup path:
   - load config
   - fail if `immersionTracking.enabled === false`
   - initialize immersion tracker
   - start stats server, forcing startup regardless of `stats.autoStartServer`
   - open `http://127.0.0.1:<stats.serverPort>`
   - write success/error to the response path
5. If the process is a secondary instance, Electron forwards argv to the primary instance through the existing single-instance event. The primary instance runs the same stats command handler and writes the response result to the temp file. The secondary process waits for that file and exits with the same status.

## Code Shape

### Launcher

- Add `stats` top-level subcommand in `launcher/config/cli-parser-builder.ts`
- Normalize that invocation in launcher arg parsing
- Add `launcher/commands/stats-command.ts`
- Dispatch it from `launcher/main.ts`
- Reuse existing app passthrough spawn helpers
- Add launcher tests for routing and forwarded argv

### Electron app

- Extend `src/cli/args.ts` with:
  - `stats: boolean`
  - `statsResponsePath?: string`
- Update app start gating so `--stats` starts the app
- Add a focused stats CLI runtime service instead of burying stats launch logic inside `main.ts`
- Reuse existing immersion tracker startup and stats server helpers where possible
- Add a single function that:
  - validates immersion tracking enabled
  - ensures tracker exists
  - ensures stats server exists
  - opens browser
  - reports completion/failure to optional response file

## Error Handling

- `immersionTracking.enabled === false`: hard failure, clear message
- tracker init failure: hard failure, clear message
- server start failure: hard failure, clear message
- browser open failure: hard failure, clear message
- response-path write failure: log warning; primary runtime behavior still follows command result

## Testing

- Launcher parser/routing tests for `subminer stats`
- Launcher forwarding test verifies `--stats` and `--stats-response-path`
- App CLI arg tests for `--stats`
- App runtime tests for:
  - stats command starts tracker/server/browser
  - stats command forces server start even when auto-start is off
  - stats command fails when immersion tracking is disabled
  - second-instance command path can surface failure via response file plumbing

## Docs

- Update CLI help text
- Update user docs where launcher/browser stats access is described
