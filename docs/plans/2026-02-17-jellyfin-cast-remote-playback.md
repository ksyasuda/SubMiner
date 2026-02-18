# Jellyfin Cast-to-Device Remote Playback Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Let Jellyfin users cast media to SubMiner so playback opens in mpv with SubMiner's existing Jellyfin subtitle defaults and control behavior.

**Architecture:** Add a dedicated Jellyfin remote-session service that connects as a long-lived cast target, handles inbound websocket events (`Play`, `Playstate`, selected `GeneralCommand`), and reports playback timeline state back to Jellyfin. Reuse existing Jellyfin playback planning and mpv command wiring so CLI `--jellyfin-play` and cast playback share one playback orchestration path.

**Tech Stack:** TypeScript, Node fetch/websocket-capable Jellyfin SDK, Electron main process lifecycle, mpv IPC socket client, node:test.

---

### Task 1: Add Remote-Control Config Surface

**Files:**
- Modify: `src/types.ts`
- Modify: `src/config/definitions.ts`
- Test: `src/config/definitions.test.ts` (if present) and/or config-related tests already in repo

**Step 1: Write failing config tests**

```ts
test("default jellyfin remote control config is enabled", () => {
  assert.equal(DEFAULT_CONFIG.jellyfin.remoteControlEnabled, true);
  assert.equal(DEFAULT_CONFIG.jellyfin.remoteControlAutoConnect, true);
});
```

**Step 2: Run test to verify it fails**

Run: `pnpm test -- src/config/definitions.test.ts`
Expected: FAIL on missing `remoteControlEnabled` keys.

**Step 3: Add minimal config fields and defaults**

```ts
remoteControlEnabled?: boolean;
remoteControlAutoConnect?: boolean;
remoteControlDeviceName?: string;
```

```ts
remoteControlEnabled: true,
remoteControlAutoConnect: true,
remoteControlDeviceName: "SubMiner",
```

**Step 4: Run targeted test to verify pass**

Run: `pnpm test -- src/config/definitions.test.ts`
Expected: PASS.

**Step 5: Commit**

```bash
git add src/types.ts src/config/definitions.ts src/config/definitions.test.ts
git commit -m "feat: add jellyfin remote-control configuration defaults"
```

### Task 2: Add Jellyfin Remote Session Service (Connection + Capabilities)

**Files:**
- Create: `src/core/services/jellyfin-remote.ts`
- Modify: `src/core/services/index.ts` (if export barrel used)
- Test: `src/core/services/jellyfin-remote.test.ts`

**Step 1: Write failing lifecycle/capabilities tests**

```ts
test("JellyfinRemoteSession posts capabilities on connect", async () => {
  // fake client, connect, assert post_capabilities called once
});

test("JellyfinRemoteSession reconnects after websocket disconnect", async () => {
  // emit disconnect, assert reconnect backoff scheduling
});
```

**Step 2: Run tests to verify failure**

Run: `pnpm test -- src/core/services/jellyfin-remote.test.ts`
Expected: FAIL because service does not exist.

**Step 3: Implement service with minimal API**

```ts
export interface JellyfinRemoteSessionService {
  start(): Promise<void>;
  stop(): Promise<void>;
  isConnected(): boolean;
}
```

- Build static capability payload with `PlayableMediaTypes`, `SupportsMediaControl`, and `SupportedCommands` including `Play`, `Playstate`, `GeneralCommand`.
- Connect using configured server/token/device/client identity.
- On disconnect, schedule bounded exponential reconnect.

**Step 4: Run tests to verify pass**

Run: `pnpm test -- src/core/services/jellyfin-remote.test.ts`
Expected: PASS.

**Step 5: Commit**

```bash
git add src/core/services/jellyfin-remote.ts src/core/services/jellyfin-remote.test.ts src/core/services/index.ts
git commit -m "feat: add jellyfin remote session service with capability registration"
```

### Task 3: Extract Shared Jellyfin->mpv Playback Orchestrator

**Files:**
- Modify: `src/main.ts`
- Optionally create: `src/core/services/jellyfin-playback-orchestrator.ts`
- Test: `src/core/services/jellyfin-playback-orchestrator.test.ts` or `src/main` service-level tests

**Step 1: Write failing test for shared playback handler**

```ts
test("playJellyfinItemInMpv applies defaults and loads file", async () => {
  // assert sub-auto no, loadfile replace, force-media-title, sid logic
});
```

**Step 2: Run test to verify failure**

Run: `pnpm test -- src/core/services/jellyfin-playback-orchestrator.test.ts`
Expected: FAIL on missing exported helper.

**Step 3: Refactor existing `--jellyfin-play` path to call shared helper**

- Move code currently in `runJellyfinCommand` `args.jellyfinPlay` block into one reusable function:
  - resolve playback plan
  - apply mpv defaults
  - load media URL
  - preload external subtitle tracks
  - pick JP primary / EN secondary tracks

**Step 4: Run targeted tests**

Run: `pnpm test -- src/core/services/jellyfin-playback-orchestrator.test.ts src/core/services/jellyfin.test.ts`
Expected: PASS.

**Step 5: Commit**

```bash
git add src/main.ts src/core/services/jellyfin-playback-orchestrator.ts src/core/services/jellyfin-playback-orchestrator.test.ts
git commit -m "refactor: share jellyfin mpv playback orchestration"
```

### Task 4: Bridge Inbound Remote Events to mpv Controls

**Files:**
- Modify: `src/core/services/jellyfin-remote.ts`
- Modify: `src/main.ts`
- Test: `src/core/services/jellyfin-remote.test.ts`

**Step 1: Add failing event-mapping tests**

```ts
test("Play event triggers shared jellyfin playback", async () => {});
test("Playstate Pause/Unpause/Stop/Seek map to mpv commands", async () => {});
test("GeneralCommand stream index updates are handled safely", async () => {});
```

**Step 2: Run tests to verify failure**

Run: `pnpm test -- src/core/services/jellyfin-remote.test.ts`
Expected: FAIL on unhandled events.

**Step 3: Implement inbound event handlers**

- `Play`: parse `ItemIds[0]`, `AudioStreamIndex`, `SubtitleStreamIndex`, `StartPositionTicks`; call shared playback helper.
- `Playstate`: `Pause`, `Unpause`, `PlayPause`, `Stop`, `Seek` -> mpv IPC commands.
- `GeneralCommand`: minimally handle `SetAudioStreamIndex`, `SetSubtitleStreamIndex`; log unsupported commands at debug.

**Step 4: Run tests to verify pass**

Run: `pnpm test -- src/core/services/jellyfin-remote.test.ts`
Expected: PASS.

**Step 5: Commit**

```bash
git add src/core/services/jellyfin-remote.ts src/core/services/jellyfin-remote.test.ts src/main.ts
git commit -m "feat: map jellyfin remote control events to mpv playback"
```

### Task 5: Add Timeline Reporting (Playing/Progress/Stopped)

**Files:**
- Modify: `src/core/services/jellyfin-remote.ts`
- Modify: `src/main.ts`
- Test: `src/core/services/jellyfin-remote.test.ts`

**Step 1: Add failing timeline payload tests**

```ts
test("timeline start/progress/stop payloads include item and position ticks", async () => {});
test("timeline reporting errors are swallowed and logged", async () => {});
```

**Step 2: Run tests to verify failure**

Run: `pnpm test -- src/core/services/jellyfin-remote.test.ts`
Expected: FAIL on missing reporting code.

**Step 3: Implement reporting loop**

- On active cast playback start: post `Sessions/Playing`.
- While active: periodic `Sessions/Playing/Progress` based on mpv `time-pos`, `pause`, and selected tracks.
- On stop/end/disconnect: post `Sessions/Playing/Stopped`.
- Guard all reporter calls with try/catch to avoid crashing playback thread.

**Step 4: Run tests to verify pass**

Run: `pnpm test -- src/core/services/jellyfin-remote.test.ts`
Expected: PASS.

**Step 5: Commit**

```bash
git add src/core/services/jellyfin-remote.ts src/core/services/jellyfin-remote.test.ts src/main.ts
git commit -m "feat: report jellyfin cast timeline from mpv state"
```

### Task 6: Wire Lifecycle Startup/Shutdown and CLI Controls

**Files:**
- Modify: `src/main.ts`
- Modify: `src/main/state.ts`
- Modify: `src/cli/args.ts` (optional status/debug flag)
- Modify: `src/core/services/cli-command.ts` and tests if new CLI flag added

**Step 1: Add failing startup lifecycle tests**

```ts
test("remote session auto-starts when jellyfin remoteControlAutoConnect true", () => {});
test("remote session stops during app shutdown", () => {});
```

**Step 2: Run tests to verify failure**

Run: `pnpm test -- src/core/services/app-ready.test.ts src/core/services/cli-command.test.ts`
Expected: FAIL on missing lifecycle behavior.

**Step 3: Implement lifecycle wiring**

- Initialize and start remote session service on app ready when Jellyfin config/session present.
- Keep service handle in app state.
- Stop service during cleanup path.
- Optional: add `--jellyfin-remote-status` command for debugging.

**Step 4: Run targeted tests**

Run: `pnpm test -- src/core/services/app-ready.test.ts src/core/services/cli-command.test.ts`
Expected: PASS.

**Step 5: Commit**

```bash
git add src/main.ts src/main/state.ts src/cli/args.ts src/core/services/cli-command.ts src/core/services/app-ready.test.ts src/core/services/cli-command.test.ts
git commit -m "feat: wire jellyfin remote session into app lifecycle"
```

### Task 7: Documentation + Regression Validation

**Files:**
- Modify: `docs/jellyfin-integration.md`
- Modify: `docs/mpv-plugin.md` (if workflow mentions cast target)

**Step 1: Add docs updates**

- New section: cast-to-device setup prerequisites.
- New section: expected behavior (launch, subtitles, controls, resume).
- New section: troubleshooting (token invalid, device not visible, websocket disconnected).

**Step 2: Run full relevant test suite**

Run: `pnpm test -- src/core/services/jellyfin.test.ts src/core/services/jellyfin-remote.test.ts src/core/services/cli-command.test.ts src/core/services/app-ready.test.ts`
Expected: PASS.

**Step 3: Run quality gate**

Run: `pnpm test`
Expected: PASS.

**Step 4: Commit docs + final integration**

```bash
git add docs/jellyfin-integration.md docs/mpv-plugin.md
git commit -m "docs: add jellyfin cast-to-device remote playback guide"
```
