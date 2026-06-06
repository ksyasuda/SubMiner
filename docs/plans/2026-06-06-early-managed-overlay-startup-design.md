<!-- read_when: changing managed mpv startup, pause-until-ready, or visible overlay boot ordering -->

# Early Managed Overlay Startup Design

Status: approved
Date: 2026-06-06

## Problem

Managed mpv startup can pause playback immediately, then leave SubMiner's tray and visible overlay
unavailable until Yomitan/tokenization warmups finish. Startup notifications therefore miss the
overlay surface and fall back to non-overlay status paths.

## Chosen Approach

For cold `--start --background --managed-playback` launches, handle initial args before waiting for
the deferred overlay warmup. That lets the tray and visible overlay shell initialize immediately
while the existing tokenization warmups continue in the background.

The mpv plugin pause gate stays armed. Playback release still waits for SubMiner's autoplay-ready
signal, which is emitted only after tokenization warmup and visible-overlay readiness. Existing
second-instance attach behavior remains unchanged: when the launcher finds an already-running
background app, it sends the same control command to that process and reuses its warmups/tokenizer.

## Checks

- Add a startup ordering regression test for managed background playback.
- Keep the existing deferred startup ordering for non-managed launches.
- Run the startup/runtime test slice plus SubMiner verification lane.
